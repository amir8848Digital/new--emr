VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{AB2CC32A-E652-42A2-B341-26ECE02F6386}#2.1#0"; "MwfCtl.ocx"
Begin VB.Form EmrFrpInv 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Invoice"
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
      Left            =   11700
      TabIndex        =   42
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
      TabIndex        =   41
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
      Left            =   5865
      TabIndex        =   51
      Top             =   9570
      Width           =   3555
      Begin MwfCtl.MWCTL_MED aktxt 
         Height          =   345
         Index           =   0
         Left            =   210
         TabIndex        =   52
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
         TabIndex        =   53
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
      Left            =   7080
      TabIndex        =   50
      Top             =   9585
      Visible         =   0   'False
      Width           =   855
      Begin CRVIEWER9LibCtl.CRViewer9 CRV_REP 
         Height          =   495
         Left            =   480
         TabIndex        =   47
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
      TabIndex        =   43
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
      Height          =   9735
      Left            =   0
      TabIndex        =   46
      Top             =   60
      Width           =   15255
      Begin TabDlg.SSTab TAB_REP 
         Height          =   9465
         Left            =   15
         TabIndex        =   45
         Top             =   45
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   16695
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "&Data Range"
         TabPicture(0)   =   "EmrFrpInv.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "fra_tabrep(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "&Group Sort"
         TabPicture(1)   =   "EmrFrpInv.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "fra_tabrep(1)"
         Tab(1).ControlCount=   1
         Begin VB.Frame fra_tabrep 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   9070
            Index           =   1
            Left            =   -74940
            TabIndex        =   49
            Top             =   360
            Width           =   15045
            Begin MwfCtl.MWCTL_GRP GRP_REP 
               Height          =   360
               Left            =   240
               TabIndex        =   44
               Top             =   360
               Width           =   8745
               _ExtentX        =   15425
               _ExtentY        =   635
            End
         End
         Begin VB.Frame fra_tabrep 
            BorderStyle     =   0  'None
            Height          =   9015
            Index           =   0
            Left            =   90
            TabIndex        =   48
            Top             =   390
            Width           =   15045
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Show Logo(s)"
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
               Index           =   14
               Left            =   120
               TabIndex        =   40
               Tag             =   "ShowDCWt"
               ToolTipText     =   "Check To Show Dia/CS Wt"
               Top             =   8175
               Width           =   3240
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Show Dia/CS Wt"
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
               Index           =   13
               Left            =   120
               TabIndex        =   39
               Tag             =   "ShowDCWt"
               ToolTipText     =   "Check To Show Dia/CS Wt"
               Top             =   7880
               Visible         =   0   'False
               Width           =   3240
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Sort by Order Cust"
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
               Index           =   12
               Left            =   6000
               TabIndex        =   8
               Tag             =   "ShowBin"
               ToolTipText     =   "Check to Sort by Order Customer"
               Top             =   1680
               Width           =   2400
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Show Accessories in Metal Loss Summary       "
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
               Index           =   11
               Left            =   5880
               TabIndex        =   17
               Tag             =   "ShowAccInMetLoss"
               ToolTipText     =   "Check To Show Accessories in Metal Loss Summary"
               Top             =   3000
               Width           =   4680
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Show CE Regn. No.     "
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
               Index           =   10
               Left            =   120
               TabIndex        =   38
               Tag             =   "ShowBin"
               ToolTipText     =   "Check To Show CE Regn No. instead of Bin No."
               Top             =   7580
               Width           =   3240
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Show Bag Details                     "
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
               Height          =   240
               Index           =   9
               Left            =   120
               TabIndex        =   37
               Tag             =   "ShowPic"
               ToolTipText     =   "Check To Show Picture"
               Top             =   7260
               Visible         =   0   'False
               Width           =   3240
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Show Making Charges"
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
               Index           =   8
               Left            =   120
               TabIndex        =   32
               Tag             =   "ShowMkChg"
               ToolTipText     =   "Check To Show Making Charges"
               Top             =   5565
               Width           =   3240
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Show Payment Terms               "
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
               Index           =   7
               Left            =   120
               TabIndex        =   30
               Tag             =   "ShowPayTerms"
               ToolTipText     =   "Check To Show Payment Terms "
               Top             =   5280
               Width           =   3240
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Show Grs Weight                     "
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
               Index           =   6
               Left            =   120
               TabIndex        =   28
               Tag             =   "ShowDCWt"
               ToolTipText     =   "Check To Show Gross Weight"
               Top             =   4995
               Width           =   3240
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Show Net Weight                      "
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
               Index           =   4
               Left            =   120
               TabIndex        =   26
               Tag             =   "ShowNetWt"
               ToolTipText     =   "Check To Show Net Weight"
               Top             =   4710
               Width           =   3240
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Show Decl For Cust                   "
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
               Index           =   5
               Left            =   120
               TabIndex        =   24
               Tag             =   "ShowCustDecl"
               ToolTipText     =   "Check To Show Declaration For Customer"
               Top             =   4425
               Width           =   3240
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Show Detail Line Demarcation "
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
               Index           =   3
               Left            =   120
               TabIndex        =   20
               Tag             =   "ShowDetLine"
               ToolTipText     =   "Check To Show Line In Detail Section "
               Top             =   3570
               Width           =   3240
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Show Metal Loss Summary       "
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
               Index           =   2
               Left            =   120
               TabIndex        =   16
               Tag             =   "ShowMetLs"
               ToolTipText     =   "Check To Show Metal Loss Summary"
               Top             =   3000
               Width           =   3240
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Show Cust PO, DsgCd               "
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
               Left            =   120
               TabIndex        =   14
               Tag             =   "ShowCustPoDsg"
               ToolTipText     =   "Check To Show Customer PO Number, Customer Design Code "
               Top             =   2715
               Width           =   3240
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Picture                                      "
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
               Left            =   120
               TabIndex        =   12
               Tag             =   "ShowPic"
               ToolTipText     =   "Check To Show Picture"
               Top             =   2430
               Width           =   3240
            End
            Begin VB.TextBox TxtInPrdDesc 
               Height          =   825
               Left            =   3165
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   34
               ToolTipText     =   "Enter Product Description (To Be Printed In Report)"
               Top             =   6120
               Width           =   6825
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   26
               Left            =   5145
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
               CmpStr          =   "Ih.InNo ="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   20
               Left            =   3945
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
               CmpStr          =   "Ih.InYy ="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   43
               Left            =   3165
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
               CmpStr          =   "Ih.InTc ="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   45
               Left            =   4365
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
               CmpStr          =   "Ih.InChr ="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   2
               Left            =   3585
               TabIndex        =   11
               ToolTipText     =   "Show Picture (Yes / No)"
               Top             =   2430
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
               Index           =   0
               Left            =   3165
               TabIndex        =   7
               ToolTipText     =   "Enter Value On Which Invoice Report Grouping Has To Be Done"
               Top             =   1575
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               DataField       =   "UOINGRP"
               IdName          =   "UOINGRP"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   1
               Left            =   3165
               TabIndex        =   9
               ToolTipText     =   "Enter Detail Sorting To Be Shown In Report"
               Top             =   1860
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               DataField       =   "UOINDET"
               IdName          =   "UOINDET"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   4
               Left            =   3165
               TabIndex        =   10
               ToolTipText     =   "Specify Whether To Show Invoice Or Customs Weight & Amount"
               Top             =   2145
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               DataType        =   4
               DataField       =   "UOINWT"
               IdName          =   "UOINWT"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   17
               Left            =   3165
               TabIndex        =   0
               ToolTipText     =   "Enter Company Code"
               Top             =   720
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               Enabled         =   0   'False
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOCOCDFR"
               IdName          =   "UOCOCDFR"
               CmpStr          =   "Ih.InCoCd = "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   9
               Left            =   3165
               TabIndex        =   5
               ToolTipText     =   "Enter From Export Number"
               Top             =   1290
               Width           =   2625
               _ExtentX        =   4630
               _ExtentY        =   503
               MaxLength       =   20
               DataField       =   "UOINEXPNOFR"
               IdName          =   "UOINEXPNOFR"
               CmpStr          =   "Ih.InExpNo ="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   3
               Left            =   3585
               TabIndex        =   13
               ToolTipText     =   "Show Customer PO Number, Customer Design Code (Yes / No)"
               Top             =   2715
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
               Index           =   12
               Left            =   3165
               TabIndex        =   18
               ToolTipText     =   "Enter Document  Type (Original/Duplicate/Triplicate)"
               Top             =   3285
               Width           =   2625
               _ExtentX        =   4630
               _ExtentY        =   503
               MaxLength       =   15
               IdName          =   "WDOCTYP"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   6
               Left            =   3585
               TabIndex        =   19
               ToolTipText     =   "Show Line In Detail Section ('Y'/'N')"
               Top             =   3570
               Visible         =   0   'False
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               Enabled         =   0   'False
               DataType        =   4
               DataField       =   "UOYN3"
               IdName          =   "UOYN3"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   7
               Left            =   3165
               TabIndex        =   21
               ToolTipText     =   "Enter Level Of Raw Material Detail To Be Shown"
               Top             =   3855
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               DataField       =   "UORMGRP"
               IdName          =   "UORMGRP"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   8
               Left            =   3165
               TabIndex        =   22
               ToolTipText     =   "Enter Level Of Metal Loss Detail To Be Shown"
               Top             =   4140
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               DataField       =   "UOPMCDFR"
               IdName          =   "UOPMCDFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   10
               Left            =   3585
               TabIndex        =   25
               ToolTipText     =   "Show Net Weight (Yes / No)"
               Top             =   4710
               Visible         =   0   'False
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               Enabled         =   0   'False
               DataType        =   4
               DataField       =   "UOYN4"
               IdName          =   "UOYN4"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   11
               Left            =   3585
               TabIndex        =   23
               ToolTipText     =   "Show Declaration For Customer (Yes / No)"
               Top             =   4425
               Visible         =   0   'False
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               Enabled         =   0   'False
               DataType        =   4
               DataField       =   "UOYN5"
               IdName          =   "UOYN5"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   13
               Left            =   3585
               TabIndex        =   27
               ToolTipText     =   "Show Diamond/ Color Stone Weight (Yes / No)"
               Top             =   4995
               Visible         =   0   'False
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               Enabled         =   0   'False
               DataType        =   4
               DataField       =   "UOYN6"
               IdName          =   "UOYN6"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   14
               Left            =   3165
               TabIndex        =   6
               ToolTipText     =   "Enter Export Numbers"
               Top             =   1290
               Width           =   6825
               _ExtentX        =   12039
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOINEXPNOSEL"
               ReCalcParent    =   "UOENTINEXPNO"
               IdName          =   "UOINEXPNOSEL"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   15
               Left            =   3165
               TabIndex        =   33
               ToolTipText     =   "Enter Product Description Code"
               Top             =   5835
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               DataField       =   "UOPMCDTO"
               IdName          =   "UOPMCDTO"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   16
               Left            =   3585
               TabIndex        =   29
               ToolTipText     =   "Show Payment Terms (Yes / No)"
               Top             =   5280
               Visible         =   0   'False
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               Enabled         =   0   'False
               DataType        =   4
               DataField       =   "UOYN7"
               IdName          =   "UOYN7"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   18
               Left            =   3165
               TabIndex        =   35
               ToolTipText     =   "Enter Invoice Exp No"
               Top             =   6930
               Width           =   6825
               _ExtentX        =   12039
               _ExtentY        =   503
               MaxLength       =   60
               DataType        =   4
               DataField       =   "UOENTINEXPNO"
               ReCalcOn        =   "UOINEXPNOSEL"
               IdName          =   "UOENTINEXPNO"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   5
               Left            =   3585
               TabIndex        =   15
               ToolTipText     =   "Show Metal Loss Summary (Y/N)"
               Top             =   3000
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
               Index           =   19
               Left            =   3585
               TabIndex        =   31
               ToolTipText     =   "Show Payment Terms (Yes / No)"
               Top             =   5565
               Visible         =   0   'False
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               Enabled         =   0   'False
               DataType        =   4
               DataField       =   "UOYN8"
               IdName          =   "UOYN8"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   405
               Index           =   21
               Left            =   3600
               TabIndex        =   36
               ToolTipText     =   "Show Bag(Yes / No)"
               Top             =   7200
               Visible         =   0   'False
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   714
               Enabled         =   0   'False
               DataType        =   4
               DataField       =   "UOYN9"
               IdName          =   "UOYN9"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   22
               Left            =   3600
               TabIndex        =   77
               ToolTipText     =   "Show CE Regn No.(Yes / No)"
               Top             =   7560
               Visible         =   0   'False
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               Enabled         =   0   'False
               DataType        =   4
               DataField       =   "UOYN10"
               IdName          =   "UOYN10"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   23
               Left            =   10785
               TabIndex        =   78
               ToolTipText     =   "Show Accessories in Metal Loss Summary (Y/N)"
               Top             =   3000
               Visible         =   0   'False
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               Enabled         =   0   'False
               DataType        =   4
               DataField       =   "UOYN11"
               IdName          =   "UOYN11"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   24
               Left            =   8640
               TabIndex        =   79
               ToolTipText     =   "Sort by Order Customer"
               Top             =   1680
               Visible         =   0   'False
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               Enabled         =   0   'False
               DataType        =   4
               DataField       =   "UOYN12"
               IdName          =   "UOYN12"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   27
               Left            =   3600
               TabIndex        =   80
               ToolTipText     =   "Show Diamond/ Color Stone Weight (Yes / No)"
               Top             =   7860
               Visible         =   0   'False
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               Enabled         =   0   'False
               DataType        =   4
               DataField       =   "UOYN13"
               IdName          =   "UOYN13"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   25
               Left            =   3600
               TabIndex        =   81
               ToolTipText     =   "Show Diamond/ Color Stone Weight (Yes / No)"
               Top             =   8160
               Visible         =   0   'False
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               Enabled         =   0   'False
               DataType        =   4
               DataField       =   "UOYN14"
               IdName          =   "UOYN14"
            End
            Begin VB.Label lblModInExpNos 
               BackStyle       =   0  'Transparent
               Caption         =   "Entered Invoice Exp Nos"
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
               Left            =   150
               TabIndex        =   76
               Top             =   6930
               Width           =   2565
            End
            Begin VB.Label LblShowPayTerms 
               BackStyle       =   0  'Transparent
               Caption         =   "Show Payment Terms"
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
               Left            =   165
               TabIndex        =   75
               Top             =   5280
               Visible         =   0   'False
               Width           =   2415
            End
            Begin VB.Label LblPrdDescCd 
               BackStyle       =   0  'Transparent
               Caption         =   "Product Description Code"
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
               TabIndex        =   74
               Top             =   5835
               Width           =   2865
            End
            Begin VB.Label LblInPrdDesc 
               BackStyle       =   0  'Transparent
               Caption         =   "Product Description"
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
               TabIndex        =   73
               Top             =   6120
               Width           =   2835
            End
            Begin VB.Label LblTcSlash 
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
               Left            =   3885
               TabIndex        =   72
               Top             =   1005
               Width           =   105
            End
            Begin VB.Label LblYySlash 
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
               Left            =   4305
               TabIndex        =   71
               Top             =   1005
               Width           =   105
            End
            Begin VB.Label LblShowDCWt 
               BackStyle       =   0  'Transparent
               Caption         =   "Show DC Weight"
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
               Left            =   165
               TabIndex        =   70
               Top             =   4995
               Visible         =   0   'False
               Width           =   1845
            End
            Begin VB.Label LblShowCustRem 
               BackStyle       =   0  'Transparent
               Caption         =   "Show Decl For Cust"
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
               Left            =   165
               TabIndex        =   69
               Top             =   4425
               Visible         =   0   'False
               Width           =   1905
            End
            Begin VB.Label LblShowNetWt 
               BackStyle       =   0  'Transparent
               Caption         =   "Show Net Weight"
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
               Left            =   165
               TabIndex        =   68
               Top             =   4710
               Visible         =   0   'False
               Width           =   1845
            End
            Begin VB.Label LblMetLsGrp 
               BackStyle       =   0  'Transparent
               Caption         =   "Metal Loss Level"
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
               TabIndex        =   67
               Top             =   4140
               Width           =   1845
            End
            Begin VB.Label LblRmGrp 
               BackStyle       =   0  'Transparent
               Caption         =   "RM Det Level"
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
               TabIndex        =   66
               ToolTipText     =   "Location"
               Top             =   3855
               Width           =   1845
            End
            Begin VB.Label LblDetLine 
               BackStyle       =   0  'Transparent
               Caption         =   "Show Detail Line Demarcation"
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
               Left            =   165
               TabIndex        =   65
               Top             =   3570
               Visible         =   0   'False
               Width           =   2985
            End
            Begin VB.Label lblDocTyp 
               BackStyle       =   0  'Transparent
               Caption         =   "Document Copy"
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
               Left            =   165
               TabIndex        =   64
               Top             =   3285
               Width           =   1815
            End
            Begin VB.Label LblMetLsSumm 
               BackStyle       =   0  'Transparent
               Caption         =   "Show Metal Loss Summary"
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
               Left            =   165
               TabIndex        =   63
               Top             =   3000
               Visible         =   0   'False
               Width           =   2775
            End
            Begin VB.Label LblInvNo 
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
               Left            =   165
               TabIndex        =   62
               Top             =   1005
               Width           =   1845
            End
            Begin VB.Label LblPic 
               BackStyle       =   0  'Transparent
               Caption         =   "Picture"
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
               Left            =   165
               TabIndex        =   61
               Top             =   2430
               Visible         =   0   'False
               Width           =   1845
            End
            Begin VB.Label LblGrpCombo 
               BackStyle       =   0  'Transparent
               Caption         =   "Grp Combination"
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
               TabIndex        =   60
               Top             =   1575
               Width           =   1845
            End
            Begin VB.Label LblDet 
               BackStyle       =   0  'Transparent
               Caption         =   "Detail Level"
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
               TabIndex        =   59
               Top             =   1860
               Width           =   1845
            End
            Begin VB.Label LblWtAs 
               BackStyle       =   0  'Transparent
               Caption         =   "Weight/ Value As"
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
               TabIndex        =   58
               Top             =   2145
               Width           =   1845
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
               TabIndex        =   57
               Top             =   720
               Width           =   1845
            End
            Begin VB.Label LblExpNo 
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
               Left            =   165
               TabIndex        =   56
               Top             =   1290
               Width           =   1845
            End
            Begin VB.Label lblRem 
               BackStyle       =   0  'Transparent
               Caption         =   "Show Cust PO, DsgCd"
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
               Left            =   165
               TabIndex        =   55
               Top             =   2715
               Visible         =   0   'False
               Width           =   2355
            End
            Begin VB.Label LblChrSlash 
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
               Left            =   5070
               TabIndex        =   54
               Top             =   1005
               Width           =   105
            End
         End
      End
   End
End
Attribute VB_Name = "EmrFrpInv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mRep As CRAXDRT.Report
Dim Rep As New EmrRepInv
Dim Rep1 As New EmrRepShpBill
Dim Rep2 As New EmrRepGRFormFrnt
'Dim Rep2 As New EmrRepEDFFormFrnt
Dim Rep3 As New EmrRepGSPFORM
Dim Rep4 As New EmrRepGRFormFrntNew   '*** Jay 2.14 P2
'Dim Rep4 As New EmrRepInvLocal         '*** (Bef 2.13) Jen
'Dim Rep5 As New EmrRepShpBillLocal     '*** (Bef 2.13) Jen

'*** (Bef 24/02/06) For Kama (Jen 2.12)
'Dim Rep6 As New EmrRepGRFormFrntLocal  '*** (Bef 2.13) Jen
'*** (Bef 24/02/06) For Kama (Jen 2.12)
'Dim Rep7 As New EmrRepGSPFORMLocal     '*** (Bef 2.13) Jen

Dim Rep8 As New EmrRepPckLstInvBag    '3.8.0

Dim wFlagLocalExp As String
Dim moCn As MwfLib.MDOConnection   '*** (Bef speed) bef 24/11/06

Dim mb_MultiInv As Boolean    '*** (Jen 2.13)

Private Sub SetGroupSort()
  ''GRP_REP.Add "Sub Ctg", "RmSCtg", "", "RmSCtg", "hRmSCtg", "", "", "", "(Select PDesc From Param Where PTyp= 'RMSCTG' and PMCd = RmCtg and PSCd= RmSCtg) "
End Sub
Private Sub ProcTooLarge()
  Select Case UCase(ADC.MenuCd)
  '*** Jay 3.3(FG)
  ' ***** Manali Trading Module
  Case UCase("RepInv"), UCase("RepInvLoose"), UCase("RepFgInv"), UCase("RepJtInv")
  '*** Jay 3.3(FG)
    '*** (Bef 2.13) Jen
    'If wFlagLocalExp = "Y" Then
    '  Set adc.RepSource = Rep4
    '  Set mRep = Rep4
    'Else
    '*** (Bef 2.13) Jen
      Set ADC.RepSource = Rep
      Set mRep = Rep
      
      '4.1.3.0
      If UCase(ADC.MenuCd) = UCase("RepInv") Then Call SetLogo(mRep)
      
    '*** (Bef 2.13) Jen
    'End If
    '*** (Bef 2.13) Jen
    
  'Case UCase("RepShpBill"), UCase("RepShpBillLoose")     '*** (Bef 2.13)
  '*** Jay 3.3(FG)
  ' ***** Manali Trading Module
  Case UCase("RepShpBill"), UCase("RepShpBillLoose"), UCase("RepFgShpBill"), UCase("RepJtShpBill") ', UCase("RepShpBillMulti") ******* Manali 3.2.2
  '*** Jay 3.3(FG)
    '*** (Bef 2.13) Jen
    'If wFlagLocalExp = "Y" Then
    '  Set adc.RepSource = Rep5
    '  Set mRep = Rep5
    'Else
    '*** (Bef 2.13) Jen
      Set ADC.RepSource = Rep1
      Set mRep = Rep1
      
    '*** (Bef 2.13) Jen
    'End If
    '*** (Bef 2.13) Jen
    
  '******* Manali 3.2.2 inv exp no field to accommodate upto 9 inv exp nos
  Case UCase("RepShpBillMulti")
      Set ADC.RepSource = Rep1
      Set mRep = Rep1
      If ADC("UoEntInExpNo") <> "" Then GRP_REP.SetFormula mRep, "wInExpNo", Trim("'" + ADC("UoEntInExpNo") + "'")
  '******* Manali 3.2.2 inv exp no field to accommodate upto 9 inv exp nos
    
  '*** Jay 3.3(FG)
  ' ***** Manali Trading Module
  Case UCase("RepGRFormFrnt"), UCase("RepGrFormLoose"), UCase("RepFgGRFrnt"), UCase("RepJtGRFrnt") ' , UCase("RepGrFormFrntMulti") **** Manali 3.2.2
  '*** Jay 3.3(FG)
    If wFlagLocalExp = "Y" Then
      '*** (Bef 2.13)
      'Set adc.RepSource = Rep6
      'Set mRep = Rep6
      '*** (Bef 2.13)
      
      '*** (Jen 2.13)
      Set ADC.RepSource = Rep1
      Set mRep = Rep1
      '*** (Jen 2.13)
    Else
      Set ADC.RepSource = Rep2
      Set mRep = Rep2
    End If
    
  '******* Manali 3.2.2 inv exp no field to accommodate upto 9 inv exp nos
  Case UCase("RepGrFormFrntMulti")
    If wFlagLocalExp = "Y" Then
      Set ADC.RepSource = Rep1
      Set mRep = Rep1
    Else
      Set ADC.RepSource = Rep2
      Set mRep = Rep2
    End If
    If ADC("UoEntInExpNo") <> "" Then GRP_REP.SetFormula mRep, "wInExpNo", Trim("'" + ADC("UoEntInExpNo") + "'")
  '******* Manali 3.2.2 inv exp no field to accommodate upto 9 inv exp nos
 
    
  '*** Jay 2.14 P2
  Case UCase("RepGrFormFrntNew"), UCase("RepGrFormFrntNewLoose")  ' , UCase("RepGrFormFrntNewMul")  **** Manali 3.2.2
    'If wFlagLocalExp = "Y" Then
      Set ADC.RepSource = Rep4
      Set mRep = Rep4
      'adc("UOINGRP") = "CMB5"
      ADC("UoInDet") = "PCS/PRS"
    'End If
  '*** Jay 2.14 P2
    
  '******* Manali 3.2.2 inv exp no field to accommodate upto 9 inv exp nos
  Case UCase("RepGrFormFrntNewMul")
      Set ADC.RepSource = Rep4
      Set mRep = Rep4
      ADC("UoInDet") = "PCS/PRS"
      If ADC("UoEntInExpNo") <> "" Then GRP_REP.SetFormula mRep, "wInExpNo", Trim("'" + ADC("UoEntInExpNo") + "'")
  '******* Manali 3.2.2 inv exp no field to accommodate upto 9 inv exp nos
    
  '****** (Jen 3.01) added RepGSPFORMMulti ******
  '*** Jay 3.3(FG)
  '****** (Jen 3.01) added RepGSPFORMMulti ******
  ' **** Manali Trading Module
  Case UCase("RepGSPFORM"), UCase("RepFgGSP"), UCase("RepJtGSP") ' , UCase("RepGSPFORMMulti") ****** Manali 3.2.2
  '*** Jay 3.3(FG)
    '*** (Bef 2.13) Jen
    'If wFlagLocalExp = "Y" Then
    '  Set adc.RepSource = Rep7
    '  Set mRep = Rep7
    'Else
    '*** (Bef 2.13) Jen
    
      Set ADC.RepSource = Rep3
      Set mRep = Rep3
    
    '*** (Bef 2.13) Jen
    'End If
    '*** (Bef 2.13) Jen
    
  '******* Manali 3.2.2 inv exp no field to accommodate upto 9 inv exp nos
  Case UCase("RepGSPFORMMulti")
      Set ADC.RepSource = Rep3
      Set mRep = Rep3
      If ADC("UoEntInExpNo") <> "" Then GRP_REP.SetFormula mRep, "wInExpNo", Trim("'" + ADC("UoEntInExpNo") + "'")
  '******* Manali 3.2.2 inv exp no field to accommodate upto 9 inv exp nos
  
  Case UCase("RepPckLstInvBag")
      Set ADC.RepSource = Rep8
      Set mRep = Rep8
  
  End Select
  ''Set adc.RepSource = Rep
End Sub

Private Sub adc_setreprecsource()
'*** (Jen 2.12)
On Error GoTo RepErr
'*** (Jen 2.12)

  '***Geeta ***Emr207**New Fields*InHETCTxt,InHStatOfOrg,InHImpValTxt,InHShpBillSgnBy,InHMetalBOETxt,
  '****Geeta****Emr208**New Fields **InHGldFiness,InHPlFiness,
  '***InHGjepcNo,InHGjepcValidUpTo,InHRbiPolicy,RmCtg,IrRmQty, IrRmIWt as qWtInCts,IrRmIVal as qVal ****
  '****Geeta***2.10***Replaced RmCtg with Space(11)
    
  ' Zubin 212 (InHCustRem, Sales Ctg and qDmSalCtgDesc added)
  '****** (Jen 2.12) (InBOELooseExp field added) ******
  '*** Report Sql ***
  '*** Jay 3.2.0 [OdPoNo] '*** Jay 3.2.0 [OdPicNm]
  ' **** Manali 3.03 - 25/06/08 - Pd Mod - 'L' fields added
  ' **** Manali 3.03 - 26/06/08 - Ag Name and address fields included
  ' **** Manali 3.5.0 - 11/12/08 - InHHideCommYN included
  ' **** Manali 3.6.0 - 17/11/09 - Circular No added
  ' **** Manali 3.8.1 - Self Payable - InDeduct, InDeductRs, InDeductDesc, InHSelfDscYN, InHSelfDeductYN added
  '    Select space(2) as qTag, IdTc, IdYy, IdChr, IdNo, IdSr, InExpNo, InCmCd, InDt, InLmgCst, InLmgDt, InLmgSal,
  '    InVaDC, InVaX, InCnvRt, InPrdDesc, InPkgWt, InGrNo, InGrDt, InSbNo, InSbDt, InAwbNo,
  '    InAwbDt, InHAwbNo, InHAwbDt, InHBinNo, InGspNo, InGspDt, InChqNo, InChqAmt, InPoNo,
  '    InPoDt, InCmPayTerms, InFOB,
  '    InFOBRs, InComm, InCommRs, InFrt, InFrtRs, InIns, InInsRs, InDeduct, InDeductRs, InDeductDsec, InTot, InTotRs, InTotAs, InTot, InTotRs, InTotAs,
  '    InDiaCost, InCSCost, InAdv, InDiaCostDesc, InCSCostDesc, InAdvDesc, InCls, InHBkName,
  '    InHBkAdd1, InHBkAdd2, InHBkAdd3, InHBkAc, InHFrgnBkName, InHFrgnBkAc, InHExpOff,
  '    InHRbiNo, InHIecNo, InHPreCarriage, InHPreCarrier, InHVessel, InHPortOfLd, InHConOfOrg,
  '    InHMarks, InHPkgDesc, InHItcNo, InHEtcZone, InHEtcSection, InHConsignee, InCmInsBy as InHInsBy,
  '    InHInsNo, InHAgName, InHAgAdd1, InHAgAdd2, InHAgAdd3, InCmName, InCmAdd1, InCmAdd2,
  '    InCmAdd3, InCmTel, InCmFax, InCmEMail, InCmCpName, InCmPortOfDisch, InCmFinDest, InCmBuyer,
  '    InCmAgName, InCmAgAdd1, InCmAgAdd2, InCmAgAdd3,
  '    InSelfName, InSelfAdd1, InSelfAdd2, InSelfAdd3, InSelfTel, InSelfFax,
  '    InSelfEMail, InHLicNo, InHETCSvg, InHAR4No, InHQCertNo, InHGSPOrgCrit, InHStatCd,
  '    InHFtRem, space(700) as InHBnkInstr, InConAdd1, InConAdd2, InConAdd3, InBuyAdd1, InBuyAdd2,
  '    InBuyAdd3, space(750) as InHUnResolutions,InHCustRem,space(180) as InHEximPolicy, InHGspEligibility, InMetalUtil, InhAuthSgn,
  '    InHETCTxt,InHStatOfOrg,InHImpValTxt,InHShpBillSgnBy,InHMetalBOETxt, InHGspPkgDesc, InBOELooseExp, InHHideCommYN, InHSelfDscYN, InHSelfDeductYN, InCircularNo,
  '    IdKt, DmCtg, DmSalCtg, IdDmCd, IdSfx, IdDmSz, DmUom, OmPONo, OdPoNo, OdPicNm, CdCustDmCd, OdDmCol,
  '    IdVaCtg, IdGldLs, IdQty, IrRmIWt as qInvGrWt, IrRmIWt as qInvGldWt,
  '    IrRmIWt as qInvPlWt, IrRmIWt as qInvSlWt, IrRmIWt as qInvPdWt, IrRmIWt as qInvAccWt, IrRmIWt as qInvDiaWt,
  '    IrRmIWt as qInvPureGldWt, IrRmIWt as qInvPurePlWt, IrRmIWt as qInvPureSlWt,  IrRmIWt as qInvPurePdWt,IrRmIWt as qInvPureGldLs,
  '    IrRmIWt as qInvPurePlLs, IrRmIWt as qInvPureSlLs,  IrRmIWt as qInvPurePdLs, IrRmIVal as qIGldVal, IrRmIVal as qIPlVal,
  '    IrRmIVal as qISlVal, IrRmIVal as qIPdVal, IrRmIVal as qIDCVal, IrSetIVal as qISetVal, iLIVal as qILabVal, IdiFOB,
  '    IrRmZWt as qZGrWt, IrRmZWt as qZGldWt, IrRmZWt as qZPlWt, IrRmZWt as qZSlWt, IrRmZWt as qZPdWt, IrRmZWt as qZAccWt,
  '    IrRmZWt As qZDiaWt, IrRmZWt as qZPureGldWt, IrRmZWt as qZPurePlWt, IrRmZWt as qZPureSlWt, IrRmZWt as qZPurePdWt,
  '    IrRmZWt as qZPureGldLs, IrRmZWt as qZPurePlLs, IrRmZWt as qZPureSlLs, IrRmZWt as qZPurepdLs, IrRmZVal as qZGldVal,
  '    IrRmZVal as qZPlVal, IrRmZVal as qZSlVal, IrRmZVal as qZPdVal, IrRmZVal as qZDCVal, IrSetZVal as qZSetVal, iLZVal as qZLabVal,
  '    IdZFOB, OrRmCtg, space(254) as qMainRmCtgDesc, space(254) as qDmCtgDesc, space(254) as qDmSalCtgDesc,
  '    space(254) as qDmColDesc, space(254) as qVaCtgDesc, InMMTC,InHGldFiness,InHPlFiness,InHSlFiness, InHPdFiness,
  '    InHGjepcNo,InHGjepcValidUpTo,InHRbiPolicy,Space(11) as RmCtg, RmKt, IrRmIRt as qRmRt,
  '    IrRmQty, IrRmIWt as qWtInCts, IrRmIWt as qWtInGms, IrRmIVal as qVal
  '    From InvRm, InvLab, RmMst, InvDsg, DsgMst, OrdMst, OrdDsg, OrdRm, InvHd, CustDsg
  '    Where 1 = 2
  '*** Report Sql ***
  
  '*** Loss Summary Sub Rep Query (Invoice Printing) '*** (Jen 2.13)
  'Select IdTc, IdYy, IdChr, IdNo, IdGldLs, IdKt, IrRmCd, RmCtg, RmKt, PDesc as RmCtgDesc,
  'IrRmIWt as qMetWt, IrRmIWt as qMetPureWt, IrRmIWt as qMetFineWt,
  'IrRmIWt As qInvLsWt, IrRmIWt As qInvPureLsWt, IrRmIWt as qInvFineLsWt,
  'IrRmIWt As qInvTotWt, IrRmIWt As qInvPureTotWt, IrRmIWt As qInvFineTotWt
  'from InvRm, RmMst, InvDsg, InvHd, Param where 1= 2
  '*** Loss Summary Sub Rep Query (Invoice Printing) '*** (Jen 2.13)
  
  '*** Loss Summary Sub Rep Query (Invoice Printing) '(Bef 2.13) Jen
  ''    Select InTc, InYy, InChr, InNo, IdGldLs, RmCtg,
  ''    IrRmCd, IrRmIWt as qMetWt, IrRmIWt as qMetPureWt, IrRmIWt as qInvLsWt,
  ''    IrRmIWt As qInvPureLsWt, IrRmIWt As qInvTotWt, IrRmIWt As qInvPureTotWt
  ''    from InvRm, RmMst, InvDsg, InvHd where 1= 2
  '*** Loss Summary Sub Rep Query (Invoice Printing) '(Bef 2.13) Jen
    
  '*********** Loss Summary For Local (Shipping Bill) *** (Jen 2.12)
  '   Select InTc, InYy, InChr, InNo,IdSr,IdKt,IdGldLs,IdValAddn, RmCtg, RmKt,
  '   IrRmCd, IrRmZRt as qRt, IrRmIWt as qMetWt, IrRmIWt as qMetPureWt,
  '   IrRmIWt as qInvLsWt, IrRmIWt As qInvPureLsWt, IrRmIWt As qInvTotWt,
  '   IrRmIWt As qInvPureTotWt, iLZVal As qLabVal,IrSetZVal as qSetVal,
  '   IrRmIVal As qMetVal, InHGldFiness As InFiness
  '   from InvRm, RmMst,InvDsg, InvHd,InvLab where 1= 2
  '*********** Loss Summary For Local (Shipping Bill) *** (Jen 2.12)
    
  '*** Second Sub Report Sql ***
  '      Select InTc, InYy, InChr, InNo, RmCtg, RmSCtg, IrRmQty, IrRmIWt as qWtInCts,
  '      IrRmIWt as qWtInGms, IrRmIRt as qRtPerCt, IrRmIVal as qVal from
  '      InvRm, RmMst, InvHd where 1= 2
  '*** Second Sub Report Sql ***
  
  '**** Invoice DsgDesc SubReport **For Local Invoice   ***Geeta****Emr208******
  ' Zubin 212 (Sales Ctg and qDmSalCtgDesc added)
  '   Select Space(2) as qTag, DmCtg,DmSalCtg,PDesc as DmCtgDesc,PDesc as DmSalCtgDesc,DmUom,IdKt,
  '   IdVaCtg,IdGldLs,OdDmCol,OrRmCtg,IdQty as qIdQty From InvDsg, DsgMst,OrdDsg,OrdRm,Param
  '   Where 1 = 2
  '********************************
    
  '**** Raw Material Summary SubReport (For Local Invoice & Loose Rm Invoice)  (Jen 2.12)
  ' Select RmCtg+ '-'+ PSCd as qRmCtg, RmSCtg, RmCd, IrRmIRt as qRmRt, IrRmQty,
  ' IrRmIWt As qWtInCts, IrRmIWt As qWtInGms, IrRmIVal As qVal, IrRmIVal as qValRs
  ' from RmMst, InvRm, Param where 1 = 2
  '**** Raw Material Summary SubReport (For Local Invoice & Loose Rm Invoice)  (Jen 2.12)
    
    
  Dim wRmWtOpt As String, wrepcnd As String, wCnd As String, wAndCnd As String, wHdStr1 As String, wStr1 As String, wStr2 As String, wStr3 As String, wSqlStrg As String, wSubRepStr As String
  Dim wRsMetUtlsd As MDORowSet, wMetUtlsdStr As String
  Dim wShowWtFld As String, wShowMetLsSumm As String, wShowCnvRt As String, wShowRtFld As String
  Dim wSubRepStr1 As String, wRmValOpt As String
  Dim wILabValStr As String, wZLabValStr As String, wISetValStr As String, wZSetValStr As String
  Dim wRsFiness As MDORowSet, ws_FinessStr As String, RsCnt As Integer
  Dim wd_GldOunce As Double, wd_PlOunce As Double, wd_LmgGms As Double, ws_GldStr As String
  Dim ws_PlStr As String, wd_LmpGms As Double
  Dim wb_MLsRmCd As Boolean     '****** Sachin 2.14.0 Next ******
  
  ' ****** Sachin 2.12 - 06/12/05 - SJM
  Dim wd_SlOunce As Double, wd_LmsGms As Double, ws_SlvStr As String
  ' ****** Sachin 2.12 - 06/12/05 - SJM
  
  ' **** Manali 3.03 - 25/06/08 - Pd Mod
  Dim wd_PdOunce As Double, wd_LmlGms As Double, ws_PdStr As String, ws_PdOunceDsc As String
  ' **** Manali 3.03 - 25/06/08 - Pd Mod
     
  '*** (Jen 2.12)
  Dim wRmRtOpt As String
  Dim ws_RmSubRep As String, ws_RmGrpOpt As String
  Dim ws_RmCtg As String, ws_RmSCtg As String, ws_RmCd As String, ws_RmRt As String
  Dim ws_ShowRmSubRepCnd As String
  '*** (Jen 2.12)
  
  '*** (Jen 2.13)
  Dim ws_CurCd As String
  Dim ws_MinCnd As String, ws_MinWhereCnd As String, ws_MinAndCnd As String ',ws_MultiWhereCnd As String, ws_MultiAndCnd As String
  Dim ws_InExpNo As String, ws_InTotAs As String, ws_PrdDesc As String, ws_MinExpNo As String
  Dim ws_GldOunceDsc As String, ws_PlOunceDsc As String, ws_SlOunceDsc As String
  
  If mb_MultiInv = False Then
  'If adc("UoInExpNoSel") = "" Then
    'mb_MultiInv = False
    wrepcnd = ADC.RepCond
    wCnd = IIF(wrepcnd <> "", " where " + wrepcnd, " ")
    wAndCnd = IIF(wrepcnd <> "", " and " + wrepcnd, " ")
    If ADC("UoInNoFr") <> 0 Then
      ws_MinCnd = " Ih1.InCoCd= '" + ADC("UoCoCdFr") + "' and Ih1.InTc= '" + ADC("UoInTcFr") + "' and " + _
                  "Ih1.InYy= '" + ADC("UoInYyFr") + "' and Ih1.InChr= '" + ADC("UoInChrFr") + "' and " + _
                  "Ih1.InNo= " + CStr(ADC("UoInNoFr")) + " "
    Else
      ws_MinCnd = " Ih1.InCoCd= '" + ADC("UoCoCdFr") + "' and Ih1.InExpNo= '" + ADC("UoInExpNoFr") + "' "
    End If
    
    ws_MinWhereCnd = " where " + ws_MinCnd
    ws_MinAndCnd = " and " + ws_MinCnd
    ws_InExpNo = " max(Ih1.InExpNo) as InExpNo "
    ws_InTotAs = " max(Ih1.InTotAsTxt) as InTotAs "
    ws_PrdDesc = " max(Ih1.InPrdDesc) as InPrdDesc "
  Else
    mb_MultiInv = True
    
    '*** (Bef 3.1.3)
    'wCnd = " where Ih.InExpNo in (" + adc("UoInExpNoSel") + ") "
    'wAndCnd = " and Ih.InExpNo in (" + adc("UoInExpNoSel") + ") "
    '*** (Bef 3.1.3)
    
    '*** (Jen 3.1.3)
    wCnd = " where " + IIF(gs_Partition = ctCurrPrtn, " Ih.InPrtKey='" + ctCurrPrtn + "' and ", "") + " Ih.InCoCd= '" + ADC("UoCoCdFr") + "' and Ih.InExpNo in (" + ADC("UoInExpNoSel") + ") "
    wAndCnd = " and Ih.InCoCd= '" + ADC("UoCoCdFr") + "' and Ih.InExpNo in (" + ADC("UoInExpNoSel") + ") "
    '*** (Jen 3.1.3)
    
    '*** 12/08/06
    ws_MinExpNo = moCn.GetFldVal("Select min(InExpNo) from InvHd Ih " + wCnd)
    ws_MinCnd = " Ih1.InCoCd= Ih.InCoCd and Ih1.InExpNo = '" + ws_MinExpNo + "' "     '*** (Jen 3.1.3)
    'ws_MinCnd = " Ih1.InExpNo = '" + ws_MinExpNo + "' "      '*** (Bef 3.1.3)
    
    '*** 12/08/06
    '*** (Bef 12/08/06) ws_MinCnd = " Ih1.InExpNo = '" + moCn.GetFldVal("Select min(InExpNo) from InvHd Ih " + wCnd) + "' "
    ws_MinWhereCnd = " where " + ws_MinCnd
    ws_MinAndCnd = " and " + ws_MinCnd
    ws_InExpNo = " '" + moCn.GetStrConcatVal("InExpNo", "InvHd Ih ", wCnd, ", ") + "' as InExpNo "
    
    'Debug.Print moCn.GetStrConcatVal("UOINEXPNOSEL", " UsrOpt ", " where UoUsrCd='" + gs_UsrCd + "' and UoMnuCd='" + adc.MenuCd + "' ", ", ")
    'Dim www As String
    'www = moCn.GetStrConcatVal("PMCd", "Param", " where PTyp= 'TC' ", ",")
    'MsgBox www
    
    Dim ws_MinCurCd As String, ws_MaxCurCd As String
    Dim ws_MinCnvFct As String, ws_MaxCnvFct As String
    Dim ws_MinCnvRt As String, ws_MaxCnvRt As String
    
    '*** (Jen 3.1.3) added Ih alias and wCnd
    ws_MinCurCd = moCn.GetFldVal("Select min(CmCurCd) from InvHd Ih join CustMst on CmCtg= 'C' and CmCd= InCmCd /* where InExpNo in (" + ADC("UoInExpNoSel") + ") */ " + wCnd)
    ws_MaxCurCd = moCn.GetFldVal("Select max(CmCurCd) from InvHd Ih join CustMst on CmCtg= 'C' and CmCd= InCmCd /* where InExpNo in (" + ADC("UoInExpNoSel") + ") */ " + wCnd)
    If ws_MinCurCd <> ws_MaxCurCd Then DispMsg "Invoice Export Number(s) Have Different Currency Codes", etError: wCnd = " and 1= 2": wAndCnd = " and 1= 2"
    
    ws_MinCnvFct = Format(moCn.GetFldVal("Select min(RTrim (LTrim (InMulDiv))+ Str(convert (decimal (16, 4), InCnvFct))) from InvHd Ih /* where InExpNo in (" + ADC("UoInExpNoSel") + ") */ " + wCnd), "####0.0000")
    ws_MaxCnvFct = Format(moCn.GetFldVal("Select max(RTrim (LTrim (InMulDiv))+ Str(convert (decimal (16, 4), InCnvFct))) from InvHd Ih /* where InExpNo in (" + ADC("UoInExpNoSel") + ") */ " + wCnd), "####0.0000")
    If ws_MinCnvFct <> ws_MaxCnvFct Then DispMsg "Invoice Export Number(s) Have Different Currency Conversion Factors", etError: wCnd = " and 1= 2": wAndCnd = " and 1= 2"
    
    ws_MinCnvRt = Format(moCn.GetFldVal("Select min(RTrim (LTrim (InRsMulDiv))+ Str(convert (decimal (16, 4), InCnvRt))) from InvHd Ih /* where InExpNo in (" + ADC("UoInExpNoSel") + ") */ " + wCnd), "####0.0000")
    ws_MaxCnvRt = Format(moCn.GetFldVal("Select max(RTrim (LTrim (InRsMulDiv))+ Str(convert (decimal (16, 4), InCnvRt))) from InvHd Ih /* where InExpNo in (" + ADC("UoInExpNoSel") + ") */ " + wCnd), "####0.0000")
    If ws_MinCnvRt <> ws_MaxCnvRt Then DispMsg "Invoice Export Number(s) Have Different Rupee Conversion Factors", etError: wCnd = " and 1= 2": wAndCnd = " and 1= 2"
    '*** (Jen 3.1.3) added Ih alias and wCnd

    ws_InTotAs = " '" + moCn.GetFldVal("Select (case when (sum(Ih.InFrt)<> 0 or sum(Ih.InFrtRs)<> 0) AND (sum(Ih.InIns)= 0 and sum(Ih.InInsRs)= 0) then 'C&F' " + _
                 " when (sum(Ih.InFrt)= 0 and sum(Ih.InFrtRs)= 0) AND (sum(Ih.InIns)<> 0 or sum(Ih.InInsRs)<> 0) then 'C&I' " + _
                 " when (sum(Ih.InFrt)<> 0 or sum(Ih.InFrtRs)<> 0) AND (sum(Ih.InIns)<> 0 or sum(Ih.InInsRs)<> 0) then 'CIF' " + _
                 " when sum(Ih.InFrt)= 0 and sum(Ih.InFrtRs)= 0 and sum(Ih.InIns)= 0 and sum(Ih.InInsRs)= 0 then 'FOB' else '' end) " + _
                 "from InvHd Ih " + wCnd) + "' as InTotAs "
                 
    If Trim(TxtInPrdDesc.text) <> "" Then
      ws_PrdDesc = " '" + Replace(Trim(TxtInPrdDesc.text), "'", "''") + "' as InPrdDesc "
    Else
      ws_PrdDesc = " max(Ih1.InPrdDesc) as InPrdDesc "
    End If
  End If
  '*** (Jen 2.13)
  
  wShowWtFld = moCn.GetFldVal("Select hInRepWtFld from Head where HCoCd= '" + ADC("UoCoCdFr") + "' and " + _
                                                                  "HCd= '" + ctSelfCmCd + "'")

  '*** Jay 3.3(FG)
  ' ***** Manali Trading Module
  wShowMetLsSumm = IIF(UCase(ADC.MenuCd) = UCase("RepInvLoose") Or UCase(ADC.MenuCd) = UCase("RepFgInv") Or UCase(ADC.MenuCd) = UCase("RepJtInv"), "N", ADC("UOYN2"))  '*** (Jen 2.13)
  '*** (Bef 2.13) wShowMetLsSumm = adc("UOYN2")
  'wShowMetLsSumm = IIF(UCase(adc.MenuCd) = UCase("RepInvLoose"), "N", adc("UOYN2"))   '*** (Jen 2.13)
  ''wShowMetLsSumm = moCn.GetFldVal("Select hInRepMetLs from Head where HCoCd = '" + ADC("UoCoCdFr") + "' and HCd= '" + ctSelfCmCd + "'")
  '*** Jay 3.3(FG)
  wShowCnvRt = moCn.GetFldVal("Select hInRepCnvRt from Head where HCoCd= '" + ADC("UoCoCdFr") + "' and " + _
                                                                  "HCd= '" + ctSelfCmCd + "'")
  wShowRtFld = moCn.GetFldVal("Select hInRepRtFld from Head where HCoCd= '" + ADC("UoCoCdFr") + "' and " + _
                                                                  "HCd= '" + ctSelfCmCd + "'")
                                                                 
                                                                    
  '*** (Bef 24/02/06) For Kama (Jen 2.12)
  'If wFlagLocalExp = "Y" Then
  '  wILabValStr = " (Select sum(iLIVal) from InvLab where iLTc= IdTc and iLYy= IDYy and " + _
  '                "iLChr= IdChr and iLNo= IdNo and iLSr= IdSr) "
  '  wZLabValStr = " (Select sum(iLZVal) from InvLab where iLTc= IdTc and iLYy= IDYy and " + _
  '                "iLChr= IdChr and iLNo= IdNo and iLSr= IdSr) "
  '  wISetValStr = " (Select sum(IrSetIVal) from InvRm where IrTc= IdTc and IrYy= IDYy and " + _
  '                "IrChr= IdChr and IrNo= IdNo and IrSr= IdSr) "
  '  wZSetValStr = " (Select sum(IrSetZVal) from InvRm where IrTc= IdTc and IrYy= IDYy and " + _
  '                "IrChr= IdChr and IrNo= IdNo and IrSr= IdSr) "
  'Else
  '*** (Bef 24/02/06) For Kama (Jen 2.12)
    wILabValStr = " 0 "
    wZLabValStr = " 0 "
  '*** (Bef 24/02/06) For Kama (Jen 2.12)
  'End If
  '*** (Bef 24/02/06) For Kama (Jen 2.12)
  
  Set Rep = Nothing
  Set Rep1 = Nothing
  Set Rep2 = Nothing
  Set Rep3 = Nothing
  Set Rep4 = Nothing    '*** Jay 2.14 P2
  'Set Rep4 = Nothing   '*** (Bef 2.13) Jen
  'Set Rep5 = Nothing   '*** (Bef 2.13) Jen
  'Set Rep6 = Nothing   '*** (24/02/06) For Kama (Jen 2.12)      '*** (Bef 2.13) Jen
  'Set Rep7 = Nothing   '*** (Bef 2.13) Jen
  Set Rep8 = Nothing
  
  Call ProcTooLarge
  Call DispCoNm
  ''Call GRP_REP.Gen3LRep
  ''grpflds = GRP_REP.GrpFldLst
   
    
  '*** For the Detail & Group Formula of the report
  Dim wDetArr() As String, wGrpArr() As String, wDetFormula As String, wDetDescFormula As String, wDetDesc2Str As String, wDetDesc2Formula As String, wGrpFormula As String, wGrpDescFormula As String
  Dim wDmUomFlag As Boolean, qDmCtgDesc As String, qMainRmCtgDesc As String, qDmColDesc As String, qVaCtgDesc As String, qMainMet As String, wOrJoin As String, i As Integer, qDmKtDesc As String
  'Dim wGldLsFlag As Boolean  '*** (Bef 2.13) Jen
  Dim wSubRepGrpFormula As String
  Dim wStr4 As String, wStr5 As String, wStr6 As String
  ' Zubin 212
  Dim qDmSalCtgDesc As String
  '*** Jay 3.3(FG)
  Dim wGrpSubRep() As String, ws_IdDmCd As String, ws_IdSfx As String, ws_IdDmSz As String
  Dim ws_IdVaCtg As String, Ws_IdKt As String, ws_IdQty As String
  '*** Jay 3.3(FG)
  qDmCtgDesc = "'' as qDmCtgDesc": qMainRmCtgDesc = "'' as qMainRmCtgDesc"
  qDmColDesc = "'' as qDmColDesc": qVaCtgDesc = "'' as qVaCtgDesc"
  qDmKtDesc = "'' as qDmKtDesc"
  qMainMet = "'' as OrRmCtg": wOrJoin = ""
  ' Zubin 212
  qDmSalCtgDesc = "'' as qDmSalCtgDesc"
  
  wDmUomFlag = False
  'wGldLsFlag = False '*** (Bef 2.13) Jen
  '*** Jay 3.3(FG)
  ' ***** Manali Trading Module
  If UCase(ADC.MenuCd) <> UCase("RepFgInv") And UCase(ADC.MenuCd) <> UCase("RepFgShpBill") And _
     UCase(ADC.MenuCd) <> UCase("RepFgGrFrnt") And UCase(ADC.MenuCd) <> UCase("RepFgGsp") And _
     UCase(ADC.MenuCd) <> UCase("RepJtInv") And UCase(ADC.MenuCd) <> UCase("RepJtShpBill") And _
     UCase(ADC.MenuCd) <> UCase("RepJtGrFrnt") And UCase(ADC.MenuCd) <> UCase("RepJtGsp") Then
    ws_IdDmCd = "IdDmCd":  ws_IdSfx = "IdSfx":  ws_IdDmSz = "IdDmSz"
    ws_IdVaCtg = "IdVaCtg":  Ws_IdKt = "IdKt":  ws_IdQty = "IdQty"
  Else
    ws_IdDmCd = "OdDmCd":  ws_IdSfx = "OdSfx":  ws_IdDmSz = "OdDmSz"
    ws_IdVaCtg = "OdVaCtg":  Ws_IdKt = "OdKt":  ws_IdQty = "OdOrdQty"
  End If
  '*** Jay 3.3(FG)
  
  wDetArr = Split(moCn.GetFldVal("Select PDesc225 from Param where PTyp= 'INDET' and " + _
            "PMCd= '" + ADC("UoInDet") + "' "), ",")
  For i = 0 To UBound(wDetArr)
    '###OldCode  wDetFormula = wDetFormula + IIf(wDetFormula = "", "", "+") + "Trim(UpperCase({rdo." + wDetArr(i) + "}))"
    '*** (Bef 2.13) wDetFormula = wDetFormula + IIF(wDetFormula = "", "", "+") + IIF(UCase(wDetArr(i)) = "IDSR", "ToText({rdo.IdSr}, '###0')", "Trim(UpperCase({rdo." + wDetArr(i) + "}))")
    ' **** Manali 3.8.0 - TrayNo added in Det Sort
    wDetFormula = wDetFormula + IIF(wDetFormula = "", "", "+") + IIF(UCase(wDetArr(i)) = "IDSR", "ToText({rdo.IdNo}, '######0')+ '/ ' + ToText({rdo.IdSr}, '###0')", IIF(UCase(wDetArr(i)) = "IDTRAYNO", "ToText({rdo.IDTrayNo}, '###0')", "Trim(UpperCase({rdo." + wDetArr(i) + "}))"))
    'Bef  3.8.0
    'wDetFormula = wDetFormula + IIF(wDetFormula = "", "", "+") + IIF(UCase(wDetArr(i)) = "IDSR", "ToText({rdo.IdNo}, '######0')+ '/ ' + ToText({rdo.IdSr}, '###0')", "Trim(UpperCase({rdo." + wDetArr(i) + "}))")
  Next i
  'wDetFormula = "{rdo.qTag}+' '+" + wDetFormula '*** New Code (Bef 2.12)
  wDetFormula = IIF(wDetFormula = "", "", "{rdo.qTag}+' '+" + wDetFormula) '*** (Jen 2.12)
  
  '*** Default value for wDetDesc2Str
  wDetDesc2Str = " '' as OmPONo, '' as OdPoNo, '' as OdPicNm, '' as CdCustDmCd, "    '*** Jay 3.2.0 [OdPoNo] '*** Jay 3.2.0 [OdPicNm]
  
  If UCase(ADC("UoInDet")) = "DMCTG" Then
    qDmCtgDesc = " (Select PDesc from Param where PTyp= 'DMCTG' and PMCd= max(DmCtg)) as qDmCtgDesc "
    wDetDescFormula = "Trim({rdo.qDmCtgDesc})"
    wDetDesc2Formula = ""
  ElseIf UCase(ADC("UoInDet")) = "PCS/PRS" Then
    wDetDescFormula = ""
    wDetDesc2Str = " '' as OmPONo, '' as OdPoNo, '' as OdPicNm, '' as CdCustDmCd, "    '*** Jay 3.2.0 [OdPoNo] '*** Jay 3.2.0 [OdPicNm]
    wDetDesc2Formula = "{rdo.DMUom}"
  ElseIf UCase(ADC("UoInDet")) = "DMCD" Then
    wDetDescFormula = "{rdo.IdDMCd}"
    wDetDesc2Formula = ""
  '****** Sachin 2.12  22/12/05 - Dsg Suffix in Invoice Reports Details Level
'  ElseIf UCase(adc("UoInDet")) = "DMSFX" Then
'    wDetDescFormula = "{rdo.IdDMCd}+" + "(If {rdo.IdSfx}= '' then '' else ' /   ') +Trim({rdo.IdSfx})"
'    wDetDesc2Formula = ""
  '****** Sachin 2.12  22/12/05 - Dsg Suffix in Invoice Reports Details Level
  ElseIf UCase(ADC("UoInDet")) = "INVSR" Then
    '**************Geeta*************Emr206********Running  Introduced *******************
    wDetDescFormula = "Trim(ToText({rdo.IdSr}, '###0'))+'    '+ Trim({rdo.IdDmCd} )+ " + _
                      "(If {rdo.IdDmSz}= '' then '' else ' /   ') +Trim({rdo.IdDmSz})"
    '*** Jay 3.2.0 [OdPoNo] '*** Jay 3.2.0 [OdPicNm]
    '*** Jay 3.3(FG)
    wDetDesc2Str = " max(OmPONo) as OmPONo, max(OdPoNo) as OdPoNo, Max(OdPicNm) as OdPicNm, IsNull((Select CdCustDmCd from CustDsg where CdCtg= 'C' and CdCd= max(OmCmCd) and CdOurDmCd= max(" + ws_IdDmCd + ") and CdOurSfx= max(" + ws_IdSfx + ") and CdOurDmSz= max(" + ws_IdDmSz + ")), " + _
                   " IsNull((Select CdCustDmCd from CustDsg where " + IIF(gs_Partition = ctCurrPrtn, " CdPrtKey='" + ctCurrPrtn + "' and ", "") + " CdCtg= 'C' and CdCd= max(OmCmCd) and CdOurDmCd= max(" + ws_IdDmCd + ") and CdOurSfx= max(" + ws_IdSfx + ") and CdOurDmSz= ''), " + _
                   " IsNull((Select CdCustDmCd from CustDsg where " + IIF(gs_Partition = ctCurrPrtn, " CdPrtKey='" + ctCurrPrtn + "' and ", "") + " CdCtg= 'C' and CdCd= max(OmCmCd) and CdOurDmCd= max(" + ws_IdDmCd + ") and CdOurSfx= '' and CdOurDmSz= max(" + ws_IdDmSz + ")), " + _
                   " IsNull((Select CdCustDmCd from CustDsg where " + IIF(gs_Partition = ctCurrPrtn, " CdPrtKey='" + ctCurrPrtn + "' and ", "") + " CdCtg= 'C' and CdCd= max(OmCmCd) and CdOurDmCd= max(" + ws_IdDmCd + ") and CdOurSfx= '' and CdOurDmSz= ''), '')))) as CdCustDmCd, "
    '*** Jay 3.3(FG)
    wDetDesc2Formula = "{rdo.OmPoNo}+" + "(If {rdo.OdPoNo}= '' then '' else ' - ') +Trim({rdo.OdPoNo})+" + "(If {rdo.CdCustDmCd}= '' then '' else ' /   ') +Trim({rdo.CdCustDmCd})"    '*** Jay 3.2.0 [OdPoNo]
  ElseIf UCase(ADC("UoInDet")) = "DMCD/SR" Then
    wDetDescFormula = "Trim(ToText({@wRunSrNo}, '###0'))+'    '+Trim({rdo.IdDmCd})+ " + _
                      "(If {rdo.IdDmSz}= '' then '' else ' /   ') +Trim({rdo.IdDmSz})"
    '*** Jay 3.2.0 [OdPoNo] '*** Jay 3.2.0 [OdPicNm]
    '*** Jay 3.3(FG)
    wDetDesc2Str = " max(OmPONo) as OmPONo, max(OdPoNo) as OdPoNo, Max(OdPicNm) as OdPicNm, IsNull((Select CdCustDmCd from CustDsg where CdCtg= 'C' and CdCd= max(OmCmCd) and CdOurDmCd= max(" + ws_IdDmCd + ") and CdOurSfx= max(" + ws_IdSfx + ") and CdOurDmSz= max(" + ws_IdDmSz + ")), " + _
                   " IsNull((Select CdCustDmCd from CustDsg where CdCtg= 'C' and CdCd= max(OmCmCd) and CdOurDmCd= max(" + ws_IdDmCd + ") and CdOurSfx= max(" + ws_IdSfx + ") and CdOurDmSz= ''), " + _
                   " IsNull((Select CdCustDmCd from CustDsg where CdCtg= 'C' and CdCd= max(OmCmCd) and CdOurDmCd= max(" + ws_IdDmCd + ") and CdOurSfx= '' and CdOurDmSz= max(" + ws_IdDmSz + ")), " + _
                   " IsNull((Select CdCustDmCd from CustDsg where CdCtg= 'C' and CdCd= max(OmCmCd) and CdOurDmCd= max(" + ws_IdDmCd + ") and CdOurSfx= '' and CdOurDmSz= ''), '')))) as CdCustDmCd, "
    '*** Jay 3.3(FG)
    wDetDesc2Formula = "{rdo.OmPoNo}+" + "(If {rdo.OdPoNo}= '' then '' else ' - ') +Trim({rdo.OdPoNo})+" + "(If {rdo.CdCustDmCd}= '' then '' else ' /   ') +Trim({rdo.CdCustDmCd})"    '*** Jay 3.2.0 [OdPoNo]
  ElseIf UCase(ADC("UoInDet")) = "INVSR/SR" Then
    ' **** Zubin 211 **** '
    wDetDescFormula = "Trim(ToText({@wRunSrNo}, '###0'))+'    '+ " + "Trim({rdo.IdDmCd} )+ " + _
                      "(If {rdo.IdDmSz}= '' then '' else ' /   ') +Trim({rdo.IdDmSz})"
    '*** Jay 3.2.0 [OdPoNo] '*** Jay 3.2.0 [OdPicNm]
    '*** Jay 3.3(FG)
    wDetDesc2Str = " max(OmPONo) as OmPONo, max(OdPoNo) as OdPoNo,Max(OdPicNm) as OdPicNm,  IsNull((Select CdCustDmCd from CustDsg where CdCtg= 'C' and CdCd= max(OmCmCd) and CdOurDmCd= max(" + ws_IdDmCd + ") and CdOurSfx= max(" + ws_IdSfx + ") and CdOurDmSz= max(" + ws_IdDmSz + ")), " + _
                   " IsNull((Select CdCustDmCd from CustDsg where CdCtg= 'C' and CdCd= max(OmCmCd) and CdOurDmCd= max(" + ws_IdDmCd + ") and CdOurSfx= max(" + ws_IdSfx + ") and CdOurDmSz= ''), " + _
                   " IsNull((Select CdCustDmCd from CustDsg where CdCtg= 'C' and CdCd= max(OmCmCd) and CdOurDmCd= max(" + ws_IdDmCd + ") and CdOurSfx= '' and CdOurDmSz= max(" + ws_IdDmSz + ")), " + _
                   " IsNull((Select CdCustDmCd from CustDsg where CdCtg= 'C' and CdCd= max(OmCmCd) and CdOurDmCd= max(" + ws_IdDmCd + ") and CdOurSfx= '' and CdOurDmSz= ''), '')))) as CdCustDmCd, "
    '*** Jay 3.3(FG)
    wDetDesc2Formula = "{rdo.OmPoNo}+" + "(If {rdo.OdPoNo}= '' then '' else ' - ') +Trim({rdo.OdPoNo})+" + "(If {rdo.CdCustDmCd}= '' then '' else ' /   ') +Trim({rdo.CdCustDmCd})"    '*** Jay 3.2.0 [OdPoNo]
    ' **** Zubin 211 **** '
    
  ' **** Manali 3.8.0 - TrayNo
  ElseIf UCase(ADC("UoInDet")) = "TRNO/SR" Then
    wDetDescFormula = "Trim(ToText({@wRunSrNo}, '###0'))+'    '+ Trim({rdo.IdDmCd} )+ " + "(If {rdo.IdDmSz}= '' then '' else ' /   ') +Trim({rdo.IdDmSz})"
    wDetDesc2Str = " max(OmPONo) as OmPONo, max(OdPoNo) as OdPoNo,Max(OdPicNm) as OdPicNm,  IsNull((Select CdCustDmCd from CustDsg where CdCtg= 'C' and CdCd= max(OmCmCd) and CdOurDmCd= max(" + ws_IdDmCd + ") and CdOurSfx= max(" + ws_IdSfx + ") and CdOurDmSz= max(" + ws_IdDmSz + ")), " + _
                   " IsNull((Select CdCustDmCd from CustDsg where CdCtg= 'C' and CdCd= max(OmCmCd) and CdOurDmCd= max(" + ws_IdDmCd + ") and CdOurSfx= max(" + ws_IdSfx + ") and CdOurDmSz= ''), " + _
                   " IsNull((Select CdCustDmCd from CustDsg where CdCtg= 'C' and CdCd= max(OmCmCd) and CdOurDmCd= max(" + ws_IdDmCd + ") and CdOurSfx= '' and CdOurDmSz= max(" + ws_IdDmSz + ")), " + _
                   " IsNull((Select CdCustDmCd from CustDsg where CdCtg= 'C' and CdCd= max(OmCmCd) and CdOurDmCd= max(" + ws_IdDmCd + ") and CdOurSfx= '' and CdOurDmSz= ''), '')))) as CdCustDmCd, "
    
    wDetDesc2Formula = "{rdo.OmPoNo}+" + "(If {rdo.OdPoNo}= '' then '' else ' - ') +Trim({rdo.OdPoNo})+" + "(If {rdo.CdCustDmCd}= '' then '' else ' /   ') +Trim({rdo.CdCustDmCd})"    '*** Jay 3.2.0 [OdPoNo]
    
  '****** Sachin 2.12 22/12/05 - Suffix in Detail Level for Inv Reports
  'ElseIf UCase(adc("UoInDet")) = "DMSFX/SR" Then
  '  wDetDescFormula = "Trim(ToText({@wRunSrNo}, '###0'))+'    '+Trim({rdo.IdDmCd})+ " + _
  '                    "(If {rdo.IdSfx}= '' then '' else ' /   ') +Trim({rdo.IdSfx})"
  '  wDetDesc2Str = " max(OmPONo) as OmPONo, IsNull((Select CdCustDmCd from CustDsg where CdCtg= 'C' and CdCd= max(OmCmCd) and CdOurDmCd= max(IdDmCd) and CdOurSfx= max(IdSfx) and CdOurDmSz= max(IdDmSz)), " + _
  '                 " IsNull((Select CdCustDmCd from CustDsg where CdCtg= 'C' and CdCd= max(OmCmCd) and CdOurDmCd= max(IdDmCd) and CdOurSfx= max(IdSfx) and CdOurDmSz= ''), " + _
  '                 " IsNull((Select CdCustDmCd from CustDsg where CdCtg= 'C' and CdCd= max(OmCmCd) and CdOurDmCd= max(IdDmCd) and CdOurSfx= '' and CdOurDmSz= max(IdDmSz)), " + _
  '                 " IsNull((Select CdCustDmCd from CustDsg where CdCtg= 'C' and CdCd= max(OmCmCd) and CdOurDmCd= max(IdDmCd) and CdOurSfx= '' and CdOurDmSz= ''), '')))) as CdCustDmCd, "
  '  wDetDesc2Formula = "{rdo.OmPoNo}+" + "(If {rdo.CdCustDmCd}= '' then '' else ' /   ') +Trim({rdo.CdCustDmCd})"

  '****** Sachin 2.12 22/12/05 - Suffix in Detail Level for Inv Reports
  End If
  
  wGrpArr = Split(moCn.GetFldVal("Select PDesc225 from Param where PTyp= 'INGR' and " + _
         "PMCd= '" + ADC("UoInGrp") + "' "), ",")
  For i = 0 To UBound(wGrpArr)
    If UCase(wGrpArr(i)) = "RMCTG" Then
      qMainRmCtgDesc = "(Select PDesc from Param where PTyp= 'RMCTG' and PMCd= max(OrRmCtg)) as qMainRmCtgDesc"
      wGrpDescFormula = wGrpDescFormula + IIF(wGrpDescFormula = "", "", "+', '+") + "Trim({rdo.qMainRmCtgDesc})"
      qMainMet = "max(OrRmCtg) as OrRmCtg "
      '****** Sachin 3.02 - Id fields in Joins
      '*** Jay 3.3(FG)
      ' ****** Manali Trading Module
      If UCase(ADC.MenuCd) <> UCase("RepFgInv") And UCase(ADC.MenuCd) <> UCase("RepFgShpBill") And _
          UCase(ADC.MenuCd) <> UCase("RepFgGrFrnt") And UCase(ADC.MenuCd) <> UCase("RepFgGsp") And _
          UCase(ADC.MenuCd) <> UCase("RepJtInv") And UCase(ADC.MenuCd) <> UCase("RepJtShpBill") And _
          UCase(ADC.MenuCd) <> UCase("RepJtGrFrnt") And UCase(ADC.MenuCd) <> UCase("RepJtGsp") Then
          wOrJoin = " join OrdRm on " + IIF(gs_Partition = ctCurrPrtn, " OrPrtKey=OdPrtKey and ", "") + " OrOdIdNo=OdIdNo and OrCoCd= OdCoCd and OrTc= OdTc and OrYy= OdYy " + _
                        " and OrChr= OdChr and OrNo= OdNo and OrSr= OdSr and OrMainMet= 'Y' "
      End If
      '*** Jay 3.3(FG)
    ElseIf UCase(wGrpArr(i)) = "KT" Then
      qDmKtDesc = "(Select PDesc from Param where PTyp= 'KT' and PMCd= max(IdKt)) as qDmKtDesc"
      wGrpDescFormula = wGrpDescFormula + IIF(wGrpDescFormula = "", "", "+', '+") + "Trim({rdo.qDmKtDesc})"
    ElseIf UCase(wGrpArr(i)) = "DMCOL" Then
      qDmColDesc = "(Select PDesc from Param where PTyp= 'DMCOL' and PMCd= max(OdDmCol)) as qDmColDesc"
      wGrpDescFormula = wGrpDescFormula + IIF(wGrpDescFormula = "", "", "+', '+") + "Trim({rdo.qDmColDesc})"
    ElseIf UCase(wGrpArr(i)) = "VACTG" Then
      '*** Jay 3.3(FG)
      ' ***** Manali Trading Module
      If UCase(ADC.MenuCd) <> UCase("RepFgInv") And UCase(ADC.MenuCd) <> UCase("RepFgShpBill") And _
          UCase(ADC.MenuCd) <> UCase("RepFgGrFrnt") And UCase(ADC.MenuCd) <> UCase("RepFgGsp") And _
          UCase(ADC.MenuCd) <> UCase("RepJtInv") And UCase(ADC.MenuCd) <> UCase("RepJtShpBill") And _
          UCase(ADC.MenuCd) <> UCase("RepJtGrFrnt") And UCase(ADC.MenuCd) <> UCase("RepJtGsp") Then
        qVaCtgDesc = "(Select PDesc from Param where PTyp= 'VACTG' and PMCd= max(" + ws_IdVaCtg + ")) as qVaCtgDesc"
        wGrpDescFormula = wGrpDescFormula + IIF(wGrpDescFormula = "", "", "+', '+") + "Trim({rdo.qVaCtgDesc})"
      Else
        qVaCtgDesc = "'' as qVaCtgDesc"
      End If
      '*** Jay 3.3(FG)
    ElseIf UCase(wGrpArr(i)) = "GLDLS" Then
      
      '*** (Bef 2.13) Jen
      'wGldLsFlag = True
      '*****Emr208****For wFlagLocalExp='Y' Don't Show GldLs%
      'wGrpDescFormula = wGrpDescFormula + IIF(wFlagLocalExp = "Y", "", IIF(wGrpDescFormula = "", "", "+', '+") + "'@'+ToText({rdo.IdGldLs}, '##0.000')+ '%'")
      '*** (Bef 2.13) Jen
      
      '*** Jay 3.3(FG)
      ' ***** Manali Trading Module
      If UCase(ADC.MenuCd) <> UCase("RepFgInv") And UCase(ADC.MenuCd) <> UCase("RepFgShpBill") And _
          UCase(ADC.MenuCd) <> UCase("RepFgGrFrnt") And UCase(ADC.MenuCd) <> UCase("RepFgGsp") And _
          UCase(ADC.MenuCd) <> UCase("RepJtInv") And UCase(ADC.MenuCd) <> UCase("RepJtShpBill") And _
          UCase(ADC.MenuCd) <> UCase("RepJtGrFrnt") And UCase(ADC.MenuCd) <> UCase("RepJtGsp") Then
        '*** (Jen 2.13)
        wGrpDescFormula = wGrpDescFormula + IIF(wGrpDescFormula = "", "", "+', '+") + "'@'+ToText({rdo.IdGldLs}, '##0.000')+ '%'"
        '*** (Jen 2.13)
      End If
      '*** Jay 3.3(FG)
    ElseIf UCase(wGrpArr(i)) = "PCS/PRS" Then
      wDmUomFlag = True
      wGrpDescFormula = wGrpDescFormula + IIF(wGrpDescFormula = "", "", "+', '+") + "Trim({rdo.DmUom})"
    ElseIf UCase(wGrpArr(i)) = "DMCTG" Then
      qDmCtgDesc = " (Select PDesc from Param where PTyp= 'DMCTG' and PMCd= max(DmCtg)) as qDmCtgDesc"
      wGrpDescFormula = wGrpDescFormula + IIF(wGrpDescFormula = "", "", "+', '+") + "Trim({rdo.qDmCtgDesc})"
    ' Zubin 212
    ElseIf UCase(wGrpArr(i)) = "SALCTG" Then
      qDmSalCtgDesc = " (Select PDesc from Param where PTyp= 'SALCTG' and PMCd= max(DmCtg) and PSCd= max(DmSalCtg)) as qDmSalCtgDesc "
      wGrpDescFormula = wGrpDescFormula + IIF(wGrpDescFormula = "", "", "+', '+") + "Trim({rdo.qDmSalCtgDesc})"
    ' Zubin 212
    End If
    wGrpFormula = wGrpFormula + IIF(wGrpFormula = "", "", "+") + "Trim(UpperCase(ToText({rdo." + moCn.GetFldVal("Select PDesc225 from Param where PTyp= 'INGRPRCP' and PMCd= '" + wGrpArr(i) + "'") + "})))"
    
    '****Geeta****11/09/04***Emr208***
    wSubRepGrpFormula = wSubRepGrpFormula + IIF(wSubRepGrpFormula = "", "", ",") + moCn.GetFldVal("Select PDesc225 from Param where PTyp= 'INGRPRCP' and PMCd= '" + wGrpArr(i) + "'")
    '*****************************
  Next i
  
  
  '*** Jay 3.3(FG)
  ' ***** Manali Trading Module
  If UCase(ADC.MenuCd) = UCase("RepFgInv") Or UCase(ADC.MenuCd) = UCase("RepFgShpBill") Or _
     UCase(ADC.MenuCd) = UCase("RepFgGrFrnt") Or UCase(ADC.MenuCd) = UCase("RepFgGsp") Or _
     UCase(ADC.MenuCd) = UCase("RepJtInv") Or UCase(ADC.MenuCd) = UCase("RepJtShpBill") Or _
          UCase(ADC.MenuCd) = UCase("RepJtGrFrnt") Or UCase(ADC.MenuCd) = UCase("RepJtGsp") Then
     wGrpSubRep = Split(wSubRepGrpFormula, ",")
     wSubRepGrpFormula = ""
     For i = 0 To UBound(wGrpSubRep)
      If UCase(wGrpSubRep(i)) = "IDGLDLS" Or UCase(wGrpSubRep(i)) = "IDVACTG" Then
          'Skip This Group For FG Invoice.
      ElseIf UCase(wGrpSubRep(i)) = "IDKT" Then
        wGrpSubRep(i) = UCase("OdKt")
        wSubRepGrpFormula = wSubRepGrpFormula + IIF(wSubRepGrpFormula = "", "", ",") + wGrpSubRep(i)
      Else
        wSubRepGrpFormula = wSubRepGrpFormula + IIF(wSubRepGrpFormula = "", "", ",") + wGrpSubRep(i)
      End If
     Next i
  End If
  '*** Jay 3.3(FG)
  
  '*** (Bef 2.12)
  'wSubRepGrpFormula = wSubRepGrpFormula + IIF(wSubRepGrpFormula = "", "", ",")
  'wGrpFormula = "{rdo.qTag}+' '+" + wGrpFormula '*** New Code
  '*** (Bef 2.12)
  
  '*** (Jen 2.12)
  
  
  If UCase(ADC.MenuCd) = UCase("RepInv") Or UCase(ADC.MenuCd) = UCase("RepInvLoose") Or UCase(ADC.MenuCd) = UCase("RepFgInv") _
            Or UCase(ADC.MenuCd) = UCase("RepJtInv") Or UCase(ADC.MenuCd) = UCase("RepPckLstInvBag") Then
      wSubRepGrpFormula = IIF(ADC("UOYN12") = "Y", "OdOmCmCd" + IIF(wSubRepGrpFormula = "", "", ","), "") + wSubRepGrpFormula
      wGrpFormula = IIF(ADC("UOYN12") = "Y", "{rdo.OmCmCd}+' '" + IIF(wGrpFormula = "", "", "+"), "") + wGrpFormula     '4.1.2
  End If
  
  wSubRepGrpFormula = IIF(wSubRepGrpFormula = "", "", wSubRepGrpFormula + IIF(wSubRepGrpFormula = "", "", ","))
  wGrpFormula = IIF(wGrpFormula = "", "", "{rdo.qTag}+' '+" + wGrpFormula) '*** New Code
      
  
  
  '*** (Jen 2.12)
  '*** For the Detail & Group Formula of the report
    
  '*** (Bef 2.13) ###
  'wrepcnd = adc.RepCond
  'wCnd = IIF(wrepcnd <> "", " where " + wrepcnd, " ")
  '*** (Bef 2.13)
               
  '*** (Bef 2.13) Jen
  'If (UCase(adc.MenuCd) = UCase("RepInv")) And wFlagLocalExp = "Y" And wGldLsFlag = False Then
  '  DispMsg "Group Option Which has Gold Loss % As Group Participant Has To Be Selected", etError: wCnd = wCnd + IIF(Trim(wCnd) <> "", " and 1= 2 ", " where 1= 2")
  'End If
  '*** (Bef 2.13) Jen
             
  ''wCommStr = "Select " + grpflds + ", "
  '*** Option for showing Invoice Wt or Customs Weight
  If UCase(ADC("UoInWt")) = "I" Then
      GRP_REP.SetFormula Rep, "wDetDWt", "Sum ({rdo.qInvDWt}, {@wDet})"
      GRP_REP.SetFormula Rep, "wTotDWt", "Sum ({rdo.qInvDWt})"
      
      GRP_REP.SetFormula Rep, "wDetCWt", "Sum ({rdo.qInvCWt}, {@wDet})"
      GRP_REP.SetFormula Rep, "wTotCWt", "Sum ({rdo.qInvCWt})"
  
      'GRP_REP.SetFormula Rep, "wNetWt", "{rdo.qInvGldWt} + {rdo.qInvPlWt} + {rdo.qInvSlWt} + {rdo.qInvPdWt}"
      GRP_REP.SetFormula Rep, "wNetWt", "{rdo.qInvGrWt}-{rdo.qInvDWt}/5-{rdo.qInvCWt}/5"
      
      GRP_REP.SetFormula mRep, "wDetGrWt", "Sum ({rdo.qInvGrWt}, {@wDet})"
      GRP_REP.SetFormula mRep, "wTotGrWt", "Sum ({rdo.qInvGrWt})"
      GRP_REP.SetFormula mRep, "wDetAmt", "Sum ({rdo.IdiFOB}, {@wDet})"
      
  '*** Jay 3.3(FG)
  ' ***** Manali Trading Module
    If UCase(ADC.MenuCd) <> UCase("RepFgInv") And UCase(ADC.MenuCd) <> UCase("RepFgShpBill") And _
        UCase(ADC.MenuCd) <> UCase("RepFgGrFrnt") And UCase(ADC.MenuCd) <> UCase("RepFgGsp") And _
        UCase(ADC.MenuCd) <> UCase("RepJtInv") And UCase(ADC.MenuCd) <> UCase("RepJtShpBill") And _
          UCase(ADC.MenuCd) <> UCase("RepJtGrFrnt") And UCase(ADC.MenuCd) <> UCase("RepJtGsp") Then
    '*** Jay 3.3(FG)
      wRmWtOpt = "IrRmIWt"
      wRmRtOpt = "IrRmIRt"  '*** (Jen 2.12)
      wRmValOpt = "IrRmIVal"
    '*** Jay 3.3(FG)
    Else
      wRmWtOpt = "OrWt"
      wRmRtOpt = "OrSalRt"
      wRmValOpt = "OrSalVal"
    End If
    '*** Jay 3.3(FG)
    
    '*** Jay 3.3(FG)
    ' ***** Manali Trading Module
    If UCase(ADC.MenuCd) = UCase("RepInv") Or UCase(ADC.MenuCd) = UCase("RepFgInv") Or UCase(ADC.MenuCd) = UCase("RepJtInv") Then mRep.TxtHead.SetText "INVOICE (Customer Copy)"
    '*** Jay 3.3(FG)

    '*** Option for showing/ hiding the Gross/Net weight field & Wt Column
    If UCase(ADC("UOYN6")) = "Y" Then       '4.1.2
      Rep.TxtGrWt.SetText "Grs Wt (Gms)"
    Else
      If ADC("UoYN4") = "Y" Then
        Rep.TxtGrWt.SetText "Net Wt (Gms)"
        GRP_REP.SetFormula Rep, "wDetGrWt", "Sum ({@wNetWt}, {@wDet})"
        GRP_REP.SetFormula Rep, "wTotGrWt", "Sum ({@wNetWt})"
      Else
        mRep.TxtGrWt.Suppress = True
        mRep.FldDetGrWt.Suppress = True
        mRep.FldTotGrWt.Suppress = True
      End If
    End If
    
  ElseIf UCase(ADC("UoInWt")) = "Z" Then
    GRP_REP.SetFormula mRep, "wDetGrWt", "Sum ({rdo.qZGrWt}, {@wDet})"
    GRP_REP.SetFormula mRep, "wTotGrWt", "Sum ({rdo.qZGrWt})"
    GRP_REP.SetFormula mRep, "wDetAmt", "Sum ({rdo.IdZFOB}, {@wDet})"
    wRmWtOpt = "IrRmZWt"
    wRmRtOpt = "IrRmZRt"    '*** (Jen 2.12)
    wRmValOpt = "IrRmZVal"
    ' ****** Manali 3.4.1 - 24/12/08 - MenuCd condition added
    ' ***** Manali Trading Module
    If UCase(ADC.MenuCd) = UCase("RepInv") Or UCase(ADC.MenuCd) = UCase("RepFgInv") Or UCase(ADC.MenuCd) = UCase("RepJtInv") Then mRep.TxtHead.SetText "INVOICE"
    ' ****** Manali 3.4.1 - 24/12/08 - MenuCd condition added
    
        GRP_REP.SetFormula Rep, "wDetDWt", "Sum ({rdo.qZDWt}, {@wDet})"
        GRP_REP.SetFormula Rep, "wTotDWt", "Sum ({rdo.qZDWt})"
        GRP_REP.SetFormula Rep, "wDetCWt", "Sum ({rdo.qZCWt}, {@wDet})"
        GRP_REP.SetFormula Rep, "wTotCWt", "Sum ({rdo.qZCWt})"
       ' GRP_REP.SetFormula Rep, "wNetWt", "{rdo.qZGldWt} + {rdo.qZPlWt} + {rdo.qZSlWt} + {rdo.qZPdWt}"
         GRP_REP.SetFormula Rep, "wNetWt", "{rdo.qZGrWt}-{rdo.qZDWt}/5-{rdo.qZCWt}/5"
    
    ' ***** Manali 3.6.0 - 30/09/09 - Show Net Wt instead of DC Wt
'    If adc("UoYN6") = "Y" Then
'      '*** (Jen 2.13)
'      GRP_REP.SetFormula Rep, "wDetDCWt", "Sum ({rdo.qZDiaWt}, {@wDet})"
'      GRP_REP.SetFormula Rep, "wTotDCWt", "Sum ({rdo.qZDiaWt})"
'      '*** (Jen 2.13)
'    Else
    
    '*** Option for showing/ hiding the Gross/Net weight field & Wt Column
    If UCase(ADC("UOYN6")) = "Y" Then       '4.1.2
      Rep.TxtGrWt.SetText "Grs Wt (Gms)"
    Else
      If ADC("UoYN4") = "Y" Then
        Rep.TxtGrWt.SetText "Net Wt (Gms)"
        GRP_REP.SetFormula Rep, "wDetGrWt", "Sum ({@wNetWt}, {@wDet})"
        GRP_REP.SetFormula Rep, "wTotGrWt", "Sum ({@wNetWt})"
      Else
        mRep.TxtGrWt.Suppress = True
        mRep.FldDetGrWt.Suppress = True
        mRep.FldTotGrWt.Suppress = True
      End If
    End If
    
  End If
  
    wHdStr1 = "Select '2H' as qTag, max(Ih1.InTc) as IdTc, max(Ih1.InYy) as IdYy, max(Ih1.InChr) as IdChr, " + _
            "max(Ih1.InNo) as IdNo, 0 as IdSr, " + ws_InExpNo + ", max(Ih1.InCmCd) as InCmCd, max(Ih1.InPrnDt) as InDt, " + _
            "max(Ih1.InLmgCst) as InLmgCst, max(Ih1.InLmgDt) as InLmgDt, max(Ih1.InLmgSal) as InLmgSal, max(Ih1.InVaDC) as InVaDC, " + _
            "max(Ih1.InVaX) as InVaX, " + _
            "max(case when Ih1.InRsMulDiv= 'D' and convert (decimal (16, 4), Ih1.InCnvRt)> 0 then 1/ Ih1.InCnvRt else Ih1.InCnvRt end) as InCnvRt, " + _
            ws_PrdDesc + ", sum(Ih.InPkgWt) as InPkgWt, max(Ih1.InGrNo) as InGrNo, " + _
            "max(Ih1.InGrDt) as InGrDt, max(Ih1.InSbNo) as InSbNo, max(Ih1.InSbDt) as InSbDt, max(Ih1.InAwbNo) as InAwbNo, max(Ih1.InAwbDt) as InAwbDt, " + _
            "max(Ih1.InHAwbNo) as InHAwbNo, max(Ih1.InHAwbDt) as InHAwbDt, max(Ih1.InHBinNo) as InHBinNo, max(Ih1.InCERegnNo) as InCERegnNo, max(Ih1.InGspNo) as InGspNo, max(Ih1.InGspDt) as InGspDt, " + _
            "max(Ih1.InChqNo) as InChqNo, sum(Ih.InChqAmt) as InChqAmt, max(Ih1.InPoNo) as InPoNo, max(Ih1.InPoDt) as InPoDt, max(Ih1.InCmPayTerms) as InCmPayTerms, " + _
            "sum(Ih.InFOB) as InFOB, sum(Ih.InFOBRs) as InFOBRs, sum(Ih.InComm) as InComm, sum(Ih.InCommRs) as InCommRs, sum(Ih.InFrt) as InFrt, " + _
            "sum(Ih.InFrtRs) as InFrtRs, sum(Ih.InIns) as InIns, sum(Ih.InInsRs) as InInsRs, sum(Ih.InDsc) as InDsc, " + _
            "sum(Ih.InDscRs) as InDscRs, max(Ih.InDscDesc) as InDscDesc, sum(Ih1.InDeduct) as InDeduct, sum(Ih1.InDeductRs) as InDeductRs, max(Ih.InDeductDesc) as InDeductDesc, sum(Ih.InTot) as InTot, sum(Ih.InTotRs) as InTotRs, " + ws_InTotAs + ", " + _
            "sum(Ih.InDiaCost) as InDiaCost, sum(Ih.InCSCost) as InCSCost, sum(Ih.InAdv) as InAdv, max(Ih.InDiaCostDesc) as InDiaCostDesc, " + _
            "max(Ih.InCSCostDesc) as InCSCostDesc, max(Ih.InAdvDesc) as InAdvDesc, max(Ih1.InCls) as InCls, max(Ih1.InHBkName) as InHBkName, " + _
            "max(Ih1.InHBkAdd1) as InHBkAdd1, max(Ih1.InHBkAdd2) as InHBkAdd2, max(Ih1.InHBkAdd3) as InHBkAdd3, max(Ih1.InHBkAc) as InHBkAc, " + _
            "max(Ih1.InHFrgnBkName) as InHFrgnBkName, max(Ih1.InHFrgnBkAc) as InHFrgnBkAc, max(Ih1.InHExpOff) as InHExpOff, max(Ih1.InHRbiNo) as InHRbiNo, " + _
            "max(Ih1.InHIecNo) as InHIecNo, max(Ih1.InHPreCarriage) as InHPreCarriage, max(Ih1.InHPreCarrier) as InHPreCarrier, max(Ih1.InHVessel) as InHVessel, " + _
            "max(Ih1.InHPortOfLd) as InHPortOfLd, max(Ih1.InHConOfOrg) as InHConOfOrg, max(Ih1.InHMarks) as InHMarks, max(Ih1.InHPkgDesc) as InHPkgDesc, "
  'sv.44 InTCS, InTCSRs added, InIrn added
  wHdStr1 = wHdStr1 + "max(Ih1.InHItcNo) as InHItcNo, max(Ih1.InHEtcZone) as InHEtcZone, max(Ih1.InHEtcSection) as InHEtcSection, max(Ih1.InHConsignee) as InHConsignee, " + _
            "max(Ih1.InCmInsBy) as InHInsBy, max(Ih1.InHInsNo) as InHInsNo, max(Ih1.InHAgName) as InHAgName, max(Ih1.InHAgAdd1) as InHAgAdd1, " + _
            "max(Ih1.InHAgAdd2) as InHAgAdd2, max(Ih1.InHAgAdd3) as InHAgAdd3, max(Ih1.InCmName) as InCmName, max(Ih1.InCmAdd1) as InCmAdd1, " + _
            "max(Ih1.InCmAdd2) as InCmAdd2, max(Ih1.InCmAdd3) as InCmAdd3, max(Ih1.InCmTel) as InCmTel, max(Ih1.InCmFax) as InCmFax, max(Ih1.InCmEMail) as InCmEMail, " + _
            "max(Ih1.InCmCpName) as InCmCpName, max(Ih1.InCmPortOfDisch) as InCmPortOfDisch, max(Ih1.InCmFinDest) as InCmFinDest, max(Ih1.InCmBuyer) as InCmBuyer, " + _
            "max(Ih1.InCmAgName) as InCmAgName, max(Ih1.InCmAgAdd1) as InCmAgAdd1, max(Ih1.InCmAgAdd2) as InCmAgAdd2, max(Ih1.InCmAgAdd3) as InCmAgAdd3, " + _
            "max(Ih1.InSelfName) as InSelfName, max(Ih1.InSelfAdd1) as InSelfAdd1, max(Ih1.InSelfAdd2) as InSelfAdd2, max(Ih1.InSelfAdd3) as InSelfAdd3, " + _
            "max(Ih1.InSelfTel) as InSelfTel, max(Ih1.InSelfFax) as InSelfFax, max(Ih1.InSelfEMail) as InSelfEMail, max(Ih1.InHLicNo) as InHLicNo, " + _
            "max(Ih1.InHETCSvg) as InHETCSvg, max(Ih1.InHAR4No) as InHAR4No, max(Ih1.InHQCertNo) as InHQCertNo, max(Ih1.InHGSPOrgCrit) as InHGSPOrgCrit, " + _
            "max(Ih1.InHStatCd) as InHStatCd, max(Ih1.InHFtRem) as InHFtRem, max(SubString(Ih1.InHBnkInstr,1, 700)) as InHBnkInstr, " + _
            "max(Ih1.InConAdd1) as InConAdd1, max(Ih1.InConAdd2) as InConAdd2, max(Ih1.InConAdd3) as InConAdd3, max(Ih1.InBuyAdd1) as InBuyAdd1, " + _
            "max(Ih1.InBuyAdd2) as InBuyAdd2, max(Ih1.InBuyAdd3) as InBuyAdd3, max(SubString(Ih1.InHUnResolutions, 1, 750)) as InHUnResolutions, " + _
            "max(Ih1.InHCustRem) as InHCustRem, max(Ih1.InHEximPolicy) as InHEximPolicy, max(Ih1.InHGspEligibility) as InHGspEligibility, " + _
            "max(Ih1.InMetalUtil) as InMetalUtil, max(Ih1.InhAuthSgn) as InhAuthSgn, max(Ih1.InHETCTxt) as InHETCTxt, max(Ih1.InHStatOfOrg) as InHStatOfOrg, " + _
            "max(Ih1.InHImpValTxt) as InHImpValTxt, max(Ih1.InHShpBillSgnBy) as InHShpBillSgnBy, max(Ih1.InHMetalBOETxt) as InHMetalBOETxt, " + _
            "max(Ih1.InHGspPkgDesc) as InHGspPkgDesc, max(Ih1.InBOELooseExp) as InBOELooseExp, max(Ih1.InHHideCommYN) as InHHideCommYN, max(Ih1.InHSelfDscYN) as InHSelfDscYN, max(Ih1.InHSelfDeductYN) as InHSelfDeductYN," + _
            "max(Ih1.InCircularNo) as InCircularNo, sum(Ih1.InTCS) as InTCS,sum(Ih1.InTCSRs) as InTCSRs, max(Ih1.InIRN) as InIRN,"
  '*** Jay 3.2.0 [OdPoNo] '*** Jay 3.2.0 [OdPicNm]
  ' **** Manali 3.03 - 25/06/08 - Pd Mod - 'L' added
  wHdStr1 = wHdStr1 + "'' as OmCmCd, '' as IdKt, '' as DmCtg, '' as DmSalCtg, '' as IdDmCd, '' as IdSfx, '' as IdDmSz, '' as DmUOM, " + _
            "'' as OmPONo, '' as OdPoNo, '' as OdPicNm, '' as CdCustDmCd, '' as OdDmCol, '' IdVaCtg, 0.0 as IdGldLs, " + _
            "0.0 as IdQty, 0.0 qInvGrWt, 0.0 as qInvGldWt, 0.0 as qInvPlWt, 0.0 as qInvSlWt, 0.0 as qInvPdWt, 0.0 as qInvAccWt, " + _
            "0.0 as qInvDWt,0.0 as qInvCWt, 0.0 as qInvPureGldWt,0.0 as qInvPurePlWt, 0.0 as qInvPureSlWt, 0.0 as qInvPurePdWt,0.0 as qInvPureGldLs, " + _
            "0.0 as qInvPurePlLs, 0.0 as qInvPureSlLs, 0.0 as qInvPurePdLs, 0.0 as qIGldVal, 0.0 as qIPlVal, 0.0 as qISlVal, 0.0 as qIPdVal, 0.0 as qIDVal,0.0 as qICVal, 0.0 as qISetVal, 0.00 as qILabVal, " + _
            "0.0 as IdiFOB, 0.0 as qZGrWt, 0.0 as qZGldWt, 0.0 as qZPlWt, 0.0 as qZSlWt, 0.0 as qZPdWt, 0.0 as qZAccWt, 0.0 as qZDWt, 0.0 as qZCWt, " + _
            "0.0 as qZPureGldWt, 0.0 as qZPurePlWt, 0.0 as qZPureSlWt, 0.0 as qZPurePdWt, 0.0 as qZPureGldLs, 0.0 as qZPurePlLs, 0.0 as qZPureSlLs, 0.0 as qZPurePdLs, " + _
            "0.0 as qZGldVal, 0.0 as qZPlVal, 0.0 as qZSlVal, 0.0 as qZPdVal, 0.0 as qZDVal, 0.0 as qZCVal, 0.0 as qZSetVal, 0.00 as qZLabVal, 0.0 as IdZFOB, " + _
            "'' as OrRmCtg, '' as qMainRmCtgDesc, '' as qDmCtgDesc, '' as qDmSalCtgDesc, '' as qDmColDesc,'' as qDmKtDesc, '' as qVaCtgDesc, " + _
            "max(Ih1.InMMTC) as InMMTC, max(Ih1.InHGldFiness) as InHGldFiness, max(Ih1.InHPlFiness) as InHPlFiness, max(Ih1.InHSlFiness) as InHSlFiness,  max(Ih1.InHPdFiness) as InHPdFiness, " + _
            "max(Ih1.InHGjepcNo) as InHGjepcNo, max(Ih1.InHGjepcValidUpTo) as InHGjepcValidUpTo, max(Ih1.InHRbiPolicy) as InHRbiPolicy, " + _
            "'' as RmCtg, '' as RmKt, 0 as qRmRt, 0 as IrRmQty, 0 as qWtInCts, 0 as qWtInGms, 0 as qVal, 0 as IdTrayNo " + _
            "from InvHd Ih " + _
            "join InvHd Ih1 on " + ws_MinCnd + wCnd
  
  '*** Jay 3.3(FG)
  ' ***** Manali Trading Module
  If UCase(ADC.MenuCd) <> UCase("RepFgInv") And UCase(ADC.MenuCd) <> UCase("RepFgShpBill") And _
     UCase(ADC.MenuCd) <> UCase("RepFgGrFrnt") And UCase(ADC.MenuCd) <> UCase("RepFgGsp") And _
     UCase(ADC.MenuCd) <> UCase("RepJtInv") And UCase(ADC.MenuCd) <> UCase("RepJtShpBill") And _
     UCase(ADC.MenuCd) <> UCase("RepJtGrFrnt") And UCase(ADC.MenuCd) <> UCase("RepJtGsp") Then
  '*** Jay 3.3(FG)
  
    '*** Jay 3.1.1 [Discount]
    wStr1 = "Select '1D' as qTag, IdTc, IdYy, IdChr, IdNo, IdSr, " + ws_InExpNo + ", max(Ih1.InCmCd) as InCmCd, max(Ih1.InPrnDt) as InDt, " + _
            "max(Ih1.InLmgCst) as InLmgCst, max(Ih1.InLmgDt) as InLmgDt, max(Ih1.InLmgSal) as InLmgSal, max(Ih1.InVaDC) as InVaDC, max(Ih1.InVaX) as InVaX, " + _
            "max(case when Ih1.InRsMulDiv= 'D' and convert (decimal (16, 4), Ih1.InCnvRt)> 0 then 1/ Ih1.InCnvRt else Ih1.InCnvRt end) as InCnvRt, " + _
            ws_PrdDesc + ", 0 as InPkgWt, max(Ih1.InGrNo) as InGrNo, max(Ih1.InGrDt) as InGrDt, max(Ih1.InSbNo) as InSbNo, " + _
            "max(Ih1.InSbDt) as InSbDt, max(Ih1.InAwbNo) as InAwbNo, max(Ih1.InAwbDt) as InAwbDt, max(Ih1.InHAwbNo) as InHAwbNo, max(Ih1.InHAwbDt) as InHAwbDt, max(Ih1.InHBinNo) as InHBinNo,  max(Ih1.InCERegnNo) as InCERegnNo, " + _
            "max(Ih1.InGspNo) as InGspNo, max(Ih1.InGspDt) as InGspDt, max(Ih1.InChqNo) as InChqNo, " + _
            "max(Ih1.InChqAmt) as InChqAmt, max(Ih1.InPoNo) as InPoNo, max(Ih1.InPoDt) as InPoDt, max(Ih1.InCmPayTerms) as InCmPayTerms, 0 as InFOB, " + _
            "0 as InFOBRs, 0 as InComm, 0 as InCommRs, 0 as InFrt, 0 as InFrtRs, 0 as InIns, 0 as InInsRs, " + _
            "0 as InDsc, 0 as InDscRs, max(Ih.InDscDesc) as InDscDesc, 0 as InDeduct, 0 as InDeductRs, max(Ih.InDeductDesc) as InDeductDesc, 0 as InTot, 0 as InTotRs, " + ws_InTotAs + ", 0 as InDiaCost, " + _
            "0 as InCSCost, 0 as InAdv, max(Ih.InDiaCostDesc) as InDiaCostDesc, max(Ih.InCSCostDesc) as InCSCostDesc, max(Ih.InAdvDesc) as InAdvDesc,  "
            
    ' ***** Manali 3.03 - 26/06/08 - Ag Name and Address included
    ' **** Manali 3.5.0 - 11/12/08 - InHHideCommYN included
    ' **** Manali 3.6.0 - 17/11/09 - Circular No added
    ' **** Manali 3.8.1 - Self Payable in Invoice
    'sv.44 InTCS, InTCSRs added, InIRN added
    wStr1 = wStr1 + "max(Ih1.InCls) as InCls, max(Ih1.InHBkName) as InHBkName, max(Ih1.InHBkAdd1) as InHBkAdd1, max(Ih1.InHBkAdd2) as InHBkAdd2, max(Ih1.InHBkAdd3) as InHBkAdd3, " + _
            "max(Ih1.InHBkAc) as InHBkAc, max(Ih1.InHFrgnBkName) as InHFrgnBkName, max(Ih1.InHFrgnBkAc) as InHFrgnBkAc, max(Ih1.InHExpOff) as InHExpOff, max(Ih1.InHRbiNo) as InHRbiNo, " + _
            "max(Ih1.InHIecNo) as InHIecNo, max(Ih1.InHPreCarriage) as InHPreCarriage, max(Ih1.InHPreCarrier) as InHPreCarrier, max(Ih1.InHVessel) as InHVessel, " + _
            "max(Ih1.InHPortOfLd) as InHPortOfLd, max(Ih1.InHConOfOrg) as InHConOfOrg, max(Ih1.InHMarks) as InHMarks, max(Ih1.InHPkgDesc) as InHPkgDesc, " + _
            "max(Ih1.InHItcNo) as InHItcNo, max(Ih1.InHEtcZone) as InHEtcZone, max(Ih1.InHEtcSection) as InHEtcSection, max(Ih1.InHConsignee) as InHConsignee, " + _
            "max(Ih1.InCmInsBy) as InHInsBy, max(Ih1.InHInsNo) as InHInsNo, max(Ih1.InHAgName) as InHAgName, max(Ih1.InHAgAdd1) as InHAgAdd1, max(Ih1.InHAgAdd2) as InHAgAdd2, " + _
            "max(Ih1.InHAgAdd3) as InHAgAdd3, max(Ih1.InCmName) as InCmName, max(Ih1.InCmAdd1) as InCmAdd1, max(Ih1.InCmAdd2) as InCmAdd2, max(Ih1.InCmAdd3) as InCmAdd3, " + _
            "max(Ih1.InCmTel) as InCmTel, max(Ih1.InCmFax) as InCmFax, max(Ih1.InCmEMail) as InCmEMail, max(Ih1.InCmCpName) as InCmCpName, max(Ih1.InCmPortOfDisch) as InCmPortOfDisch, " + _
            "max(Ih1.InCmFinDest) as InCmFinDest, max(Ih1.InCmBuyer) as InCmBuyer, " + _
            "max(Ih1.InCmAgName) as InCmAgName, max(Ih1.InCmAgAdd1) as InCmAgAdd1, max(Ih1.InCmAgAdd2) as InCmAgAdd2, max(Ih1.InCmAgAdd3) as InCmAgAdd3, " + _
            "max(Ih1.InSelfName) as InSelfName, max(Ih1.InSelfAdd1) as InSelfAdd1, " + _
            "max(Ih1.InSelfAdd2) as InSelfAdd2, max(Ih1.InSelfAdd3) as InSelfAdd3, max(Ih1.InSelfTel) as InSelfTel, max(Ih1.InSelfFax) as InSelfFax, max(Ih1.InSelfEMail) as InSelfEMail, " + _
            "max(Ih1.InHLicNo) as InHLicNo, max(Ih1.InHETCSvg) as InHETCSvg, max(Ih1.InHAR4No) as InHAR4No, max(Ih1.InHQCertNo) as InHQCertNo, max(Ih1.InHGSPOrgCrit) as InHGSPOrgCrit, " + _
            "max(Ih1.InHStatCd) as InHStatCd, max(Ih1.InHFtRem) as InHFtRem, max(SubString(Ih1.InHBnkInstr,1, 700)) as InHBnkInstr, max(Ih1.InConAdd1) as InConAdd1, max(Ih1.InConAdd2) as InConAdd2, " + _
            "max(Ih1.InConAdd3) as InConAdd3, max(Ih1.InBuyAdd1) as InBuyAdd1, max(Ih1.InBuyAdd2) as InBuyAdd2, max(Ih1.InBuyAdd3) as InBuyAdd3, max(SubString(Ih1.InHUnResolutions, 1, 750)) as InHUnResolutions, max(Ih1.InHCustRem) as InHCustRem, " + _
            "max(Ih1.InHEximPolicy) as InHEximPolicy, max(Ih1.InHGspEligibility) as InHGspEligibility, max(Ih1.InMetalUtil) as InMetalUtil, max(Ih1.InhAuthSgn) as InhAuthSgn, " + _
            "max(Ih1.InHETCTxt) as InHETCTxt, max(Ih1.InHStatOfOrg) as InHStatOfOrg, max(Ih1.InHImpValTxt) as InHImpValTxt, max(Ih1.InHShpBillSgnBy) as InHShpBillSgnBy, max(Ih1.InHMetalBOETxt) as InHMetalBOETxt, max(Ih1.InHGspPkgDesc) as InHGspPkgDesc, max(Ih1.InBOELooseExp) as InBOELooseExp, " + _
            "max(Ih1.InHHideCommYN) as InHHideCommYN, max(Ih1.InHSelfDscYN) as InHSelfDscYN, max(Ih1.InHSelfDeductYN) as InHSelfDeductYN, max(Ih1.InCircularNo) as InCircularNo, 0 as InTCS,0 as InTCSRs,max(Ih1.InIRN) as InIRN,"
    '*** Jay 3.1.1 [Discount]
    '****** (Jen 2.14 Next Patch 1 05/06/07) added R1 alias and "/ (case when R2.RmPurityZ> 0 then R2.RmPurityZ else 1 end)"
    ' Zubin 212 (DmSalCtg added)
    ' **** Manali 3.03 - 25/06/08 - Pd Mod - 'L' added
    '*** Jay 3.3(FG) - Variable Added
    wStr2 = "max(OmCmCd) as OmCmCd,max(" + Ws_IdKt + ") as IdKt, max(DmCtg) as DmCtg, max(DmSalCtg) as DmSalCtg, max(DmCd) as IdDmCd, max(" + ws_IdSfx + ") as IdSfx, max(" + ws_IdDmSz + ") as IdDmSz, max(DmUom) as DmUom, " + wDetDesc2Str + _
            " max(OdDmCol) as OdDmCol, max(" + ws_IdVaCtg + ") as IdVaCtg, max(IdGldLs) as IdGldLs, max(" + ws_IdQty + ") as IdQty, " + _
            "sum(IrRmIWt / (case when R1.RmCtg in ('C', 'D') then 5 else 1 end)) as qInvGrWt, " + _
            "sum(IrRmIWt * (case R1.RmCtg when 'G' then 1 else 0 end)) as qInvGldWt, " + _
            "sum(IrRmIWt * (case R1.RmCtg when 'P' then 1 else 0 end)) as qInvPlWt, " + _
            "sum(IrRmIWt * (case R1.RmCtg when 'S' then 1 else 0 end)) as qInvSlWt, " + _
            "sum(IrRmIWt * (case when R1.RmCtg in ('L','M') then 1 else 0 end)) as qInvPdWt, " + _
            "sum(IrRmIWt * (case R1.RmCtg when 'X' then 1 else 0 end)) as qInvAccWt, " + _
            "sum(IrRmIWt * (case when R1.RmCtg in ('D') then 1 else 0 end)) as qInvDWt, " + _
            "sum(IrRmIWt * (case when R1.RmCtg in ('C') then 1 else 0 end)) as qInvCWt, " + _
            "sum(IrRmIWt * R1.RmPurityZ * (case R1.RmCtg when 'G' then 1 else 0 end)/ (case when R2.RmPurityZ> 0 then R2.RmPurityZ else 1 end)) as qInvPureGldWt, " + _
            "sum(IrRmIWt * R1.RmPurityZ * (case R1.RmCtg when 'P' then 1 else 0 end)/ (case when R2.RmPurityZ> 0 then R2.RmPurityZ else 1 end)) as qInvPurePlWt, " + _
            "sum(IrRmIWt * R1.RmPurityZ * (case R1.RmCtg when 'S' then 1 else 0 end)/ (case when R2.RmPurityZ> 0 then R2.RmPurityZ else 1 end)) as qInvPureSlWt, " + _
            "sum(IrRmIWt * R1.RmPurityZ * (case when R1.RmCtg in ('L','M')  then 1 else 0 end)/ (case when R2.RmPurityZ> 0 then R2.RmPurityZ else 1 end)) as qInvPurePdWt, " + _
            "sum(IrRmIWt * R1.RmPurityZ * (case R1.RmCtg when 'G' then 1 else 0 end)* (IdGldLs/100)/ (case when R2.RmPurityZ> 0 then R2.RmPurityZ else 1 end)) as qInvPureGldLs, " + _
            "sum(IrRmIWt * R1.RmPurityZ * (case R1.RmCtg when 'P' then 1 else 0 end)* (IdGldLs/100)/ (case when R2.RmPurityZ> 0 then R2.RmPurityZ else 1 end)) as qInvPurePlLs, " + _
            "sum(IrRmIWt * R1.RmPurityZ * (case R1.RmCtg when 'S' then 1 else 0 end)* (IdGldLs/100)/ (case when R2.RmPurityZ> 0 then R2.RmPurityZ else 1 end)) as qInvPureSlLs, " + _
            "sum(IrRmIWt * R1.RmPurityZ * (case when R1.RmCtg in ('L','M') then 1 else 0 end)* (IdGldLs/100)/ (case when R2.RmPurityZ> 0 then R2.RmPurityZ else 1 end)) as qInvPurePdLs, " + _
            "sum(IrRmIVal * (case R1.RmCtg when 'G' then 1 else 0 end)) as qIGldVal, " + _
            "sum(IrRmIVal * (case R1.RmCtg when 'P' then 1 else 0 end)) as qIPlVal, " + _
            "sum(IrRmIVal * (case R1.RmCtg when 'S' then 1 else 0 end)) as qISlVal, " + _
            "sum(IrRmIVal * (case when R1.RmCtg in ('L','M') then 1 else 0 end)) as qIPdVal, " + _
            "sum(IrRmIVal * (case when R1.RmCtg in ('D') then 1 else 0 end)) as qIDVal, " + _
            "sum(IrRmIVal * (case when R1.RmCtg in ('C') then 1 else 0 end)) as qICVal, " + _
            "sum(IrSetIVal) as qISetVal, " + wILabValStr + " as qILabVal, max(IdiFOB) as IdiFOB, "
    
    '****** (Jen 2.14 Next Patch 1 05/06/07) added R1 alias and "/ (case when R2.RmPurityZ> 0 then R2.RmPurityZ else 1 end)"
    '*** *** (Jen 2.12) Note: added the fields RmKt, qRmRt, qWtInGms in the query *** ***'
    ' **** Manali 3.03 - 25/06/08 - Pd Mod - 'L' added
    wStr3 = "sum(IrRmZWt / (case when R1.RmCtg in ('C', 'D') then 5 else 1 end)) as qZGrWt, " + _
            "sum(IrRmZWt * (case R1.RmCtg when 'G' then 1 else 0 end)) as qZGldWt, " + _
            "sum(IrRmZWt * (case R1.RmCtg when 'P' then 1 else 0 end)) as qZPlWt, " + _
            "sum(IrRmZWt * (case R1.RmCtg when 'S' then 1 else 0 end)) as qZSlWt, " + _
            "sum(IrRmZWt * (case when R1.RmCtg in ('L','M') then 1 else 0 end)) as qZPdWt, " + _
            "sum(IrRmZWt * (case R1.RmCtg when 'X' then 1 else 0 end)) as qZAccWt, sum(IrRmZWt * (case when R1.RmCtg in ('D') then 1 else 0 end)) as qZDWt, " + _
            "sum(IrRmZWt * (case when R1.RmCtg in ('C') then 1 else 0 end)) as qZCWt, " + _
            "sum(IrRmZWt * R1.RmPurityZ * (case R1.RmCtg when 'G' then 1 else 0 end)/ (case when R2.RmPurityZ> 0 then R2.RmPurityZ else 1 end)) as qZPureGldWt, " + _
            "sum(IrRmZWt * R1.RmPurityZ * (case R1.RmCtg when 'P' then 1 else 0 end)/ (case when R2.RmPurityZ> 0 then R2.RmPurityZ else 1 end)) as qZPurePlWt, " + _
            "sum(IrRmZWt * R1.RmPurityZ * (case R1.RmCtg when 'S' then 1 else 0 end)/ (case when R2.RmPurityZ> 0 then R2.RmPurityZ else 1 end)) as qZPureSlWt, " + _
            "sum(IrRmZWt * R1.RmPurityZ * (case when R1.RmCtg in ('L','M') then 1 else 0 end)/ (case when R2.RmPurityZ> 0 then R2.RmPurityZ else 1 end)) as qZPurePdWt, " + _
            "sum(" + " (case when IrMainMet= 'Y' then IdCustmMtchWt- IrRmZWt else IrRmZWt* IdGldLs/100 end) " + "* R1.RmPurityZ * (case R1.RmCtg when 'G' then 1 else 0 end)/ (case when R2.RmPurityZ> 0 then R2.RmPurityZ else 1 end)) as qZPureGldLs, " + _
            "sum(" + " (case when IrMainMet= 'Y' then IdCustmMtchWt- IrRmZWt else IrRmZWt* IdGldLs/100 end) " + "* R1.RmPurityZ * (case R1.RmCtg when 'P' then 1 else 0 end)/ (case when R2.RmPurityZ> 0 then R2.RmPurityZ else 1 end)) as qZPurePlLs, " + _
            "sum(" + " (case when IrMainMet= 'Y' then IdCustmMtchWt- IrRmZWt else IrRmZWt* IdGldLs/100 end) " + "* R1.RmPurityZ * (case R1.RmCtg when 'S' then 1 else 0 end)/ (case when R2.RmPurityZ> 0 then R2.RmPurityZ else 1 end)) as qZPureSlLs, " + _
            "sum(" + " (case when IrMainMet= 'Y' then IdCustmMtchWt- IrRmZWt else IrRmZWt* IdGldLs/100 end) " + "* R1.RmPurityZ * (case when R1.RmCtg in ('L','M') then 1 else 0 end)/ (case when R2.RmPurityZ> 0 then R2.RmPurityZ else 1 end)) as qZPurePdLs, " + _
            "sum(IrRmZVal * (case R1.RmCtg when 'G' then 1 else 0 end)) as qZGldVal, " + _
            "sum(IrRmZVal * (case R1.RmCtg when 'P' then 1 else 0 end)) as qZPlVal, " + _
            "sum(IrRmZVal * (case R1.RmCtg when 'S' then 1 else 0 end)) as qZSlVal, " + _
            "sum(IrRmZVal * (case when R1.RmCtg in ('L','M') then 1 else 0 end)) as qZPdVal, " + _
            "sum(IrRmZVal * (case when R1.RmCtg in ('D') then 1 else 0 end)) as qZDVal,sum(IrRmZVal * (case when R1.RmCtg in ('C') then 1 else 0 end)) as qZCVal, "
    
            '*** (Bef 2.14 Next Patch 1 05/06/07)
            '"sum(IrRmZWt * RmPurityZ * (case RmCtg when 'G' then 1 else 0 end)) as qZPureGldWt, " + _
            '"sum(IrRmZWt * RmPurityZ * (case RmCtg when 'P' then 1 else 0 end)) as qZPurePlWt, " + _
            '"sum(IrRmZWt * RmPurityZ * (case RmCtg when 'S' then 1 else 0 end)) as qZPureSlWt, " + _
            '"sum(" + " (case when IrMainMet= 'Y' then IdCustmMtchWt- IrRmZWt else IrRmZWt* IdGldLs/100 end) " + "* RmPurityZ * (case RmCtg when 'G' then 1 else 0 end)) as qZPureGldLs, " + _
            '"sum(" + " (case when IrMainMet= 'Y' then IdCustmMtchWt- IrRmZWt else IrRmZWt* IdGldLs/100 end) " + "* RmPurityZ * (case RmCtg when 'P' then 1 else 0 end)) as qZPurePlLs, " + _
            '"sum(" + " (case when IrMainMet= 'Y' then IdCustmMtchWt- IrRmZWt else IrRmZWt* IdGldLs/100 end) " + "* RmPurityZ * (case RmCtg when 'S' then 1 else 0 end)) as qZPureSlLs, " + _
            '*** (Bef 2.14 Next Patch 1 05/06/07)
    
            '*** (04/08/06)
            '"sum(IrRmZWt * RmPurityZ * (case RmCtg when 'G' then 1 else 0 end)* IdGldLs/100) as qZPureGldLs, " + _
            '"sum(IrRmZWt * RmPurityZ * (case RmCtg when 'P' then 1 else 0 end)* IdGldLs/100) as qZPurePlLs, " + _
            '"sum(IrRmZWt * RmPurityZ * (case RmCtg when 'S' then 1 else 0 end)* IdGldLs/100) as qZPureSlLs, "
            '*** (04/08/06)
    
    '****** (Jen 2.14 Next Patch 1 05/06/07) added R2 join
    '****** Sachin 3.02 - Id fields in Joins
    ' **** Manali 3.03 - 25/06/08 - Pd Mod - 'L' added
    wStr3 = wStr3 + "sum(IrSetZVal) as qZSetVal, " + wZLabValStr + " as qZLabVal, max(IdZFOB) as IdZFOB, " + qMainMet + ", " + qMainRmCtgDesc + ", " + qDmCtgDesc + ", " + qDmSalCtgDesc + ", " + qDmColDesc + ", " + qDmKtDesc + ", " + qVaCtgDesc + ", max(Ih1.InMMTC) as InMMTC," + _
            " max(Ih1.InHGldFiness) as InHGldFiness, max(Ih1.InHPlFiness) as InHPlFiness, max(Ih1.InHSlFiness) as InHSlFiness,  max(Ih1.InHPdFiness) as InHPdFiness, max(Ih1.InHGjepcNo) as InHGjepcNo, max(Ih1.InHGjepcValidUpTo) as InHGjepcValidUpTo, Max(Ih1.InHRbiPolicy) as InHRbiPolicy, " + _
            " '' as RmCtg, '' as RmKt, 0 as qRmRt, 0 as IrRmQty, 0 as qWtInCts, 0 as qWtInGms, 0 as qVal, Max(IdTrayNo) as IdTrayNo " + _
            " from InvRm join RmMst R1 on " + IIF(gs_Partition = ctCurrPrtn, " R1.RmPrtKey=IrPrtKey and ", "") + " R1.RmCd= IrRmCd join RmMst R2 on R2.RmCd= R1.RmZCd " + IIF(gs_Partition = ctCurrPrtn, " And R2.RmPrtKey=R1.RmPrtKey ", "") + _
            " join InvDsg on IdIdNo=IrIdIdNo and IdCoCd= IrCoCd and IdTc= IrTc and IdYy= IrYy and IdChr= IrChr " + _
            " and IdNo= IrNo and IdSr= IrSr " + IIF(gs_Partition = ctCurrPrtn, " And IdPrtKey=IrPrtKey ", "") + _
            " join DsgMst on DmIdNo=IdDmIdNo and DmTcTyp= 'DM' and DmCd = IdDmCd and DmSz= '' " + IIF(gs_Partition = ctCurrPrtn, " And DmPrtKey=IdPrtKey ", "") + _
            " join OrdMst on OmCoCd= IdCoCd and OmTc= IdExpOdTc and OmYy= IdExpOdYy " + _
            " and OmChr= IdExpOdChr and OmNo= IdExpOdNo " + IIF(gs_Partition = ctCurrPrtn, " And OmPrtKey=IdPrtKey ", "") + _
            " join OrdDsg on OdCoCd= IdCoCd and OdTc= IdExpOdTc and OdYy= IdExpOdYy " + _
            " and OdChr= IdExpOdChr and OdNo= IdExpOdNo and OdSr= IdExpOdSr " + IIF(gs_Partition = ctCurrPrtn, " And OdPrtKey=IdPrtKey ", "") + wOrJoin + _
            " join InvHd Ih on Ih.InIdNo=IdInIdNo and Ih.InCoCd= IdCoCd and Ih.InTc = IdTc and Ih.InYy= IdYy and Ih.InChr= IdChr and Ih.InNo= IdNo " + IIF(gs_Partition = ctCurrPrtn, " And Ih.InPrtKey=IdPrtKey ", "") + _
            " join InvHd Ih1 on " + ws_MinCnd + wCnd + IIF(gs_Partition = ctCurrPrtn, " And IrPrtKey='" + ctCurrPrtn + "' ", "") + " Group by IdTc, IdYy, IdChr, IdNo, IdSr "
    '*** (Jen 2.13)
            
            
    '*** (Bef 2.13)
    ''*** *** (Jen 2.12) Note: added the fields RmKt, qRmRt, qWtInGms in the query *** ***'
    '' Zubin 212 (InHCustRem, DmSalCtg and qDmSalCtgDesc added)
    'wHdStr1 = "Select '2H' as qTag, InTc as IdTc, InYy as IdYy, InChr as IdChr, InNo as IdNo, 0 as IdSr, InExpNo, InCmCd, InPrnDt as InDt, InLmgCst, InLmgDt, " + _
    '          "InLmgSal, InVaDC, InVaX, InCnvRt, InPrdDesc, InPkgWt, InGrNo, InGrDt, InSbNo, " + _
    '          "InSbDt, InAwbNo, InAwbDt, InHAwbNo, InHAwbDt, InHBinNo, InGspNo, InGspDt, InChqNo, " + _
    '          "InChqAmt, InPoNo, InPoDt, InCmPayTerms, InFOB, InFOBRs, InComm, InCommRs, InFrt, " + _
    '          "InFrtRs, InIns, InInsRs, InTot, InTotRs, InTotAs, InDiaCost, InCSCost, InAdv, " + _
    '          "InDiaCostDesc, InCSCostDesc, InAdvDesc, InCls, InHBkName, InHBkAdd1, InHBkAdd2, " + _
    '          "InHBkAdd3, InHBkAc, InHFrgnBkName, InHFrgnBkAc, InHExpOff, InHRbiNo, InHIecNo, " + _
    '          "InHPreCarriage, InHPreCarrier, InHVessel, InHPortOfLd, InHConOfOrg, InHMarks, " + _
    '          "InHPkgDesc, InHItcNo, InHEtcZone, InHEtcSection, InHConsignee, InCmInsBy as InHInsBy, InHInsNo, " + _
    '          "InHAgName, InHAgAdd1, InHAgAdd2, InHAgAdd3, InCmName, InCmAdd1, InCmAdd2, InCmAdd3, " + _
    '          "InCmTel, InCmFax, InCmEMail, InCmCpName, InCmPortOfDisch, InCmFinDest, InCmBuyer, " + _
    '          "InSelfName, InSelfAdd1, InSelfAdd2, InSelfAdd3, InSelfTel, InSelfFax, InSelfEMail,InHLicNo, InHETCSvg, InHAR4No, " + _
    '          "InHQCertNo, InHGSPOrgCrit, InHStatCd, InHFtRem,SubString(InHBnkInstr,1, 700) as InHBnkInstr, InConAdd1, InConAdd2, InConAdd3, InBuyAdd1, " + _
    '          "InBuyAdd2,InBuyAdd3, SubString(InHUnResolutions, 1, 750) as InHUnResolutions, InHCustRem, InHEximPolicy, InHGspEligibility, InMetalUtil, InhAuthSgn, " + _
    '          "InHETCTxt,InHStatOfOrg,InHImpValTxt,InHShpBillSgnBy,InHMetalBOETxt,InHGspPkgDesc, InBOELooseExp, " + _
    '          "'' as IdKt, '' as DmCtg, '' as DmSalCtg, '' as IdDmCd, '' as IdSfx, '' as IdDmSz, '' as DmUOM, " + _
    '          "'' as OmPONo, '' as CdCustDmCd, '' as OdDmCol, '' IdVaCtg, 0.0 as IdGldLs, " + _
    '          "0.0 as IdQty, 0.0 qInvGrWt, 0.0 as qInvGldWt, 0.0 as qInvPlWt, 0.0 as qInvSlWt, 0.0 as qInvAccWt, " + _
    '          "0.0 as qInvDiaWt, 0.0 as qInvPureGldWt,0.0 as qInvPurePlWt, 0.0 as qInvPureSlWt, 0.0 as qInvPureGldLs, " + _
    '          "0.0 as qInvPurePlLs, 0.0 as qInvPureSlLs, 0.0 as qIGldVal, 0.0 as qIPlVal, 0.0 as qISlVal, 0.0 as qIDCVal, 0.0 as qISetVal, 0.00 as qILabVal, " + _
    '          "0.0 as IdiFOB, 0.0 as qZGrWt, 0.0 as qZGldWt, 0.0 as qZPlWt, 0.0 as qZSlWt, 0.0 as qZAccWt, 0.0 as qZDiaWt, " + _
    '          "0.0 as qZPureGldWt, 0.0 as qZPurePlWt, 0.0 as qZPureSlWt, 0.0 as qZPureGldLs, 0.0 as qZPurePlLs, 0.0 as qZPureSlLs, " + _
    '          "0.0 as qZGldVal, 0.0 as qZPlVal, 0.0 as qZSlVal, 0.0 as qZDCVal, 0.0 as qZSetVal, 0.00 as qZLabVal, 0.0 as IdZFOB, " + _
    '          "'' as OrRmCtg, '' as qMainRmCtgDesc, '' as qDmCtgDesc, '' as qDmSalCtgDesc, '' as qDmColDesc, '' as qVaCtgDesc, InMMTC, InHGldFiness, InHPlFiness, InHSlFiness, InHGjepcNo, InHGjepcValidUpTo, InHRbiPolicy, " + _
    '          "'' as RmCtg, '' as RmKt, 0 as qRmRt, 0 as IrRmQty, 0 as qWtInCts, 0 as qWtInGms, 0 as qVal from InvHd " + wCnd
    
    'wStr1 = "Select '1D' as qTag, IdTc, IdYy, IdChr, IdNo, IdSr, max(InExpNo) as InExpNo, max(InCmCd) as InCmCd, max(InPrnDt) as InDt, " + _
    '        "max(InLmgCst) as InLmgCst, max(InLmgDt) as InLmgDt, max(InLmgSal) as InLmgSal, max(InVaDC) as InVaDC, max(InVaX) as InVaX, " + _
    '        "max(InCnvRt) as InCnvRt, max(InPrdDesc) as InPrdDesc, max(InPkgWt) as InPkgWt, max(InGrNo) as InGrNo, max(InGrDt) as InGrDt, max(InSbNo) as InSbNo, " + _
    '        "max(InSbDt) as InSbDt, max(InAwbNo) as InAwbNo, max(InAwbDt) as InAwbDt, max(InHAwbNo) as InHAwbNo, max(InHAwbDt) as InHAwbDt, max(InHBinNo) as InHBinNo, " + _
    '        "max(InGspNo) as InGspNo, max(InGspDt) as InGspDt, max(InChqNo) as InChqNo, " + _
    '        "max(InChqAmt) as InChqAmt, max(InPoNo) as InPoNo, max(InPoDt) as InPoDt, max(InCmPayTerms) as InCmPayTerms, max(InFOB) as InFOB, " + _
    '        "max(InFOBRs) as InFOBRs, max(InComm) as InComm, max(InCommRs) as InCommRs, max(InFrt) as InFrt, max(InFrtRs) as InFrtRs, max(InIns) as InIns, " + _
    '        "max(InInsRs) as InInsRs, max(InTot) as InTot, max(InTotRs) as InTotRs, max(InTotAs) as InTotAs, max(InDiaCost) as InDiaCost, " + _
    '        "max(InCSCost) as InCSCost, max(InAdv) as InAdv, max(InDiaCostDesc) as InDiaCostDesc, max(InCSCostDesc) as InCSCostDesc, max(InAdvDesc) as InAdvDesc, " + _
    '        "max(InCls) as InCls, max(InHBkName) as InHBkName, max(InHBkAdd1) as InHBkAdd1, max(InHBkAdd2) as InHBkAdd2, max(InHBkAdd3) as InHBkAdd3, " + _
    '        "max(InHBkAc) as InHBkAc, max(InHFrgnBkName) as InHFrgnBkName, max(InHFrgnBkAc) as InHFrgnBkAc, max(InHExpOff) as InHExpOff, max(InHRbiNo) as InHRbiNo, " + _
    '        "max(InHIecNo) as InHIecNo, max(InHPreCarriage) as InHPreCarriage, max(InHPreCarrier) as InHPreCarrier, max(InHVessel) as InHVessel, " + _
    '        "max(InHPortOfLd) as InHPortOfLd, max(InHConOfOrg) as InHConOfOrg, max(InHMarks) as InHMarks, max(InHPkgDesc) as InHPkgDesc, " + _
    '        "max(InHItcNo) as InHItcNo, max(InHEtcZone) as InHEtcZone, max(InHEtcSection) as InHEtcSection, max(InHConsignee) as InHConsignee, " + _
    '        "max(InCmInsBy) as InHInsBy, max(InHInsNo) as InHInsNo, max(InHAgName) as InHAgName, max(InHAgAdd1) as InHAgAdd1, max(InHAgAdd2) as InHAgAdd2, " + _
    '        "max(InHAgAdd3) as InHAgAdd3, max(InCmName) as InCmName, max(InCmAdd1) as InCmAdd1, max(InCmAdd2) as InCmAdd2, max(InCmAdd3) as InCmAdd3, " + _
    '        "max(InCmTel) as InCmTel, max(InCmFax) as InCmFax, max(InCmEMail) as InCmEMail, max(InCmCpName) as InCmCpName, max(InCmPortOfDisch) as InCmPortOfDisch, " + _
    '        "max(InCmFinDest) as InCmFinDest, max(InCmBuyer) as InCmBuyer, max(InSelfName) as InSelfName, max(InSelfAdd1) as InSelfAdd1, " + _
    '        "max(InSelfAdd2) as InSelfAdd2, max(InSelfAdd3) as InSelfAdd3, max(InSelfTel) as InSelfTel, max(InSelfFax) as InSelfFax, max(InSelfEMail) as InSelfEMail, " + _
    '        "max(InHLicNo) as InHLicNo, max(InHETCSvg) as InHETCSvg, max(InHAR4No) as InHAR4No, max(InHQCertNo) as InHQCertNo, max(InHGSPOrgCrit) as InHGSPOrgCrit, " + _
    '        "max(InHStatCd) as InHStatCd, max(InHFtRem) as InHFtRem, max(SubString(InHBnkInstr,1, 700)) as InHBnkInstr, max(InConAdd1) as InConAdd1, max(InConAdd2) as InConAdd2, " + _
    '        "max(InConAdd3) as InConAdd3, max(InBuyAdd1) as InBuyAdd1, max(InBuyAdd2) as InBuyAdd2, max(InBuyAdd3) as InBuyAdd3, max(SubString(InHUnResolutions, 1, 750)) as InHUnResolutions, max(InHCustRem) as InHCustRem, " + _
    '        "max(InHEximPolicy) as InHEximPolicy, max(InHGspEligibility) as InHGspEligibility, max(InMetalUtil) as InMetalUtil, max(InhAuthSgn) as InhAuthSgn, " + _
    '        "max(InHETCTxt) as InHETCTxt,max(InHStatOfOrg) as InHStatOfOrg,max(InHImpValTxt) as InHImpValTxt,max(InHShpBillSgnBy) as InHShpBillSgnBy,max(InHMetalBOETxt) as InHMetalBOETxt, max(InHGspPkgDesc) as InHGspPkgDesc, max(InBOELooseExp) as InBOELooseExp, "
    '
    '' Zubin 212 (DmSalCtg added)
    'wStr2 = "max(IdKt) as IdKt, max(DmCtg) as DmCtg, max(DmSalCtg) as DmSalCtg, max(DmCd) as IdDmCd, max(IdSfx) as IdSfx, max(IdDmSz) as IdDmSz, max(DmUom) as DmUom, " + wDetDesc2Str + _
    '        " max(OdDmCol) as OdDmCol, max(IdVaCtg) as IdVaCtg, max(IdGldLs) as IdGldLs, max(IdQty) as IdQty, " + _
    '        "sum(IrRmIWt / (case when RmCtg in ('C', 'D') then 5 else 1 end)) as qInvGrWt, " + _
    '        "sum(IrRmIWt * (case RmCtg when 'G' then 1 else 0 end)) as qInvGldWt, " + _
    '        "sum(IrRmIWt * (case RmCtg when 'P' then 1 else 0 end)) as qInvPlWt, " + _
    '        "sum(IrRmIWt * (case RmCtg when 'S' then 1 else 0 end)) as qInvSlWt, " + _
    '        "sum(IrRmIWt * (case RmCtg when 'X' then 1 else 0 end)) as qInvAccWt, " + _
    '        "sum(IrRmIWt * (case when RmCtg in ('C', 'D') then 1 else 0 end)) as qInvDiaWt, " + _
    '        "sum(IrRmIWt * RmPurityZ * (case RmCtg when 'G' then 1 else 0 end)) as qInvPureGldWt, " + _
    '        "sum(IrRmIWt * RmPurityZ * (case RmCtg when 'P' then 1 else 0 end)) as qInvPurePlWt, " + _
    '        "sum(IrRmIWt * RmPurityZ * (case RmCtg when 'S' then 1 else 0 end)) as qInvPureSlWt, " + _
    '        "sum(IrRmIWt * RmPurityZ * (case RmCtg when 'G' then 1 else 0 end)* IdGldLs/100) as qInvPureGldLs, " + _
    '        "sum(IrRmIWt * RmPurityZ * (case RmCtg when 'P' then 1 else 0 end)* IdGldLs/100) as qInvPurePlLs, " + _
    '        "sum(IrRmIWt * RmPurityZ * (case RmCtg when 'S' then 1 else 0 end)* IdGldLs/100) as qInvPureSlLs, " + _
    '        "sum(IrRmIVal * (case RmCtg when 'G' then 1 else 0 end)) as qIGldVal, " + _
    '        "sum(IrRmIVal * (case RmCtg when 'P' then 1 else 0 end)) as qIPlVal, " + _
    '        "sum(IrRmIVal * (case RmCtg when 'S' then 1 else 0 end)) as qISlVal, " + _
    '        "sum(IrRmIVal * (case when RmCtg in ('C', 'D') then 1 else 0 end)) as qIDCVal, " + _
    '        "sum(IrSetIVal) as qISetVal, " + wILabValStr + " as qILabVal, max(IdiFOB) as IdiFOB, "
    '
    ''*** *** (Jen 2.12) Note: added the fields RmKt, qRmRt, qWtInGms in the query *** ***'
    'wStr3 = "sum(IrRmZWt / (case when RmCtg in ('C', 'D') then 5 else 1 end)) as qZGrWt, " + _
    '        "sum(IrRmZWt * (case RmCtg when 'G' then 1 else 0 end)) as qZGldWt, " + _
    '        "sum(IrRmZWt * (case RmCtg when 'P' then 1 else 0 end)) as qZPlWt, " + _
    '        "sum(IrRmZWt * (case RmCtg when 'S' then 1 else 0 end)) as qZSlWt, " + _
    '        "sum(IrRmZWt * (case RmCtg when 'X' then 1 else 0 end)) as qZAccWt, sum(IrRmZWt * (case when RmCtg in ('C', 'D') then 1 else 0 end)) as qZDiaWt, " + _
    '        "sum(IrRmZWt * RmPurityZ * (case RmCtg when 'G' then 1 else 0 end)) as qZPureGldWt, " + _
    '        "sum(IrRmZWt * RmPurityZ * (case RmCtg when 'P' then 1 else 0 end)) as qZPurePlWt, " + _
    '        "sum(IrRmZWt * RmPurityZ * (case RmCtg when 'S' then 1 else 0 end)) as qZPureSlWt, " + _
    '        "sum(IrRmZWt * RmPurityZ * (case RmCtg when 'G' then 1 else 0 end)* IdGldLs/100) as qZPureGldLs, " + _
    '        "sum(IrRmZWt * RmPurityZ * (case RmCtg when 'P' then 1 else 0 end)* IdGldLs/100) as qZPurePlLs, " + _
    '        "sum(IrRmZWt * RmPurityZ * (case RmCtg when 'S' then 1 else 0 end)* IdGldLs/100) as qZPureSlLs, " + _
    '        "sum(IrRmZVal * (case RmCtg when 'G' then 1 else 0 end)) as qZGldVal, " + _
    '        "sum(IrRmZVal * (case RmCtg when 'P' then 1 else 0 end)) as qZPlVal, " + _
    '        "sum(IrRmZVal * (case RmCtg when 'S' then 1 else 0 end)) as qZSlVal, " + _
    '        "sum(IrRmZVal * (case when RmCtg in ('C', 'D') then 1 else 0 end)) as qZDCVal, "
    'wStr3 = wStr3 + "sum(IrSetZVal) as qZSetVal, " + wZLabValStr + " as qZLabVal, max(IdZFOB) as IdZFOB, " + qMainMet + ", " + qMainRmCtgDesc + ", " + qDmCtgDesc + ", " + qDmSalCtgDesc + ", " + qDmColDesc + ", " + qVaCtgDesc + ", max(InMMTC) as InMMTC," + _
    '        "max(InHGldFiness) as InHGldFiness,max(InHPlFiness) as InHPlFiness,max(InHSlFiness) as InHSlFiness,max(InHGjepcNo) as InHGjepcNo,max(InHGjepcValidUpTo) as InHGjepcValidUpTo,Max(InHRbiPolicy) as InHRbiPolicy, " + _
    '        "'' as RmCtg, '' as RmKt, 0 as qRmRt, 0 as IrRmQty, 0 as qWtInCts, 0 as qWtInGms, 0 as qVal " + _
    '        " from InvRm join RmMst on RmCd = IrRmCd " + _
    '        "join InvDsg on IdCoCd= IrCoCd and IdTc= IrTc and IdYy= IrYy and IdChr= IrChr " + _
    '        "and IdNo= IrNo and IdSr= IrSr " + _
    '        "join DsgMst on DmTcTyp= 'DM' and DmCd = IdDmCd and DmSz= '' " + _
    '        "join OrdMst on OmCoCd= IdCoCd and OmTc= IdExpOdTc and OmYy= IdExpOdYy " + _
    '        "and OmChr= IdExpOdChr and OmNo= IdExpOdNo " + _
    '        "join OrdDsg on OdCoCd= IdCoCd and OdTc= IdExpOdTc and OdYy= IdExpOdYy " + _
    '        "and OdChr= IdExpOdChr and OdNo= IdExpOdNo and OdSr= IdExpOdSr " + wOrJoin + _
    '        " join InvHd on InCoCd= IdCoCd and InTc = IdTc and InYy= IdYy and InChr= IdChr and " + _
    '        "InNo= IdNo " + wCnd + " group by IdTc, IdYy, IdChr, IdNo, IdSr "
    '*** (Bef 2.13)
            
          

    '*** (Bef 2.12)
    '**********Geeta**********208****
    'wStr4 = "max(IdKt) as IdKt, max(DmCtg) as DmCtg, max(DmCd) as IdDmCd, max(IdSfx) as IdSfx, max(IdDmSz) as IdDmSz, max(DmUom) as DmUom, " + wDetDesc2Str + _
    '        " max(OdDmCol) as OdDmCol, max(IdVaCtg) as IdVaCtg, max(IdGldLs) as IdGldLs, 0  as qIdQty, " + _
    '        "0 as qInvGrWt, 0 as qInvGldWt, 0 as qInvPlWt, " + _
    '        "0 as qInvAccWt, 0 as qInvDiaWt, 0 as qInvPureGldWt, " + _
    '        "0 as qInvPurePlWt, 0 as qInvPureGldLs, 0 as qInvPurePlLs, " + _
    '        "0 as qIGldVal, 0 as qIPlVal, 0 as qIDCVal, 0 as qISetVal, 0 as qILabVal, 0 as qIdiFOB, "
    '**********Geeta**********208****
    '*** (Bef 2.12)
    
    
    '*** (Bef 2.12 Patch 2)
    '*** (Jen 2.12)
    '' Zubin 212 (DmSalCtg and qDmSalCtgDesc added)
    'wStr4 = "max(IdKt) as IdKt, max(DmCtg) as DmCtg, max(DmSalCtg) as DmSalCtg, max(DmCd) as IdDmCd, max(IdSfx) as IdSfx, max(IdDmSz) as IdDmSz, max(DmUom) as DmUom, " + wDetDesc2Str + _
    '        " max(OdDmCol) as OdDmCol, max(IdVaCtg) as IdVaCtg, max(IdGldLs) as IdGldLs, 0  as IdQty, " + _
    '        "0 as qInvGrWt, 0 as qInvGldWt, 0 as qInvPlWt, 0 as qInvSlWt, " + _
    '        "0 as qInvAccWt, 0 as qInvDiaWt, 0 as qInvPureGldWt, " + _
    '        "0 as qInvPurePlWt, 0 as qInvPureSlWt, 0 as qInvPureGldLs, 0 as qInvPurePlLs, 0 as qInvPureSlLs, " + _
    '        "0 as qIGldVal, 0 as qIPlVal, 0 as qISlVal, 0 as qIDCVal, 0 as qISetVal, 0 as qILabVal, 0 as qIdiFOB, " + _
    '        "0 as qZGrWt,0 as qZGldWt,0 as qZPlWt,0 as qZSlWt,0 as qZAccWt,0 as qZDiaWt, " + _
    '        "0 as qZPureGldWt,0 as qZPurePlWt,0 as qZPureSlWt,0 as qZPureGldLs, 0 as qZPurePlLs,0 as qZPureSlLs, " + _
    '        "0 as qZGldVal,0 as qZPlVal,0 as qZSlVal,0 as qZDCVal, 0 as qZSetVal, 0 as qZLabVal, 0 as IdZFOB, " + qMainMet + ", " + qMainRmCtgDesc + ", " + qDmCtgDesc + ", " + qDmSalCtgDesc + _
    '        ", " + qDmColDesc + ", " + qVaCtgDesc + ", max(InMMTC) as InMMTC, max(InHGldFiness) as InHGldFiness, max(InHPlFiness) as InHPlFiness, max(InHSlFiness) as InHSlFiness," + _
    '        "max(InHGjepcNo) as InHGjepcNo, max(InHGjepcValidUpTo) as InHGjepcValidUpTo, Max(InHRbiPolicy) as InHRbiPolicy, "
    '
    ''RmCtg, RmKt, IrRmIRt as qRmRt, sum(IrRmQty) as IrRmQty, IrRmIWt as qWtInCts, IrRmIWt as qWtInGms, IrRmIVal as qVal
    'wStr5 = "RmCtg+(Case When PSCd is Null Then '' Else '-'+IsNull(PSCd,'') End ) as RmCtg, " + _
    '        "RmKt, " + wRmRtOpt + " as qRmRt, sum(IrRmQty) as IrRmQty, " + _
    '        "sum(case when RmCtg in ('C', 'D') then " + wRmWtOpt + " else 0 end) as qWtInCts, " + _
    '        "sum(case when RmCtg not in ('C', 'D') then " + wRmWtOpt + " else 0 end) as qWtInGms, " + _
    '        "sum(" + wRmValOpt + ") as qVal " + _
    '        " from InvRm join RmMst on RmCd = IrRmCd " + _
    '        "Left Outer Join Param On PTyp='STNCLS' and PMCd= RMCTG and  (','+PDesc225+',') Like ('%,'+RmSCtg+',%') " + _
    '        " join InvDsg on IdCoCd= IrCoCd and IdTc= IrTc and IdYy= IrYy and IdChr= IrChr " + _
    '        "and IdNo= IrNo and IdSr= IrSr " + _
    '        "join DsgMst on DmTcTyp= 'DM' and DmCd = IdDmCd and DmSz= '' " + _
    '        "join OrdMst on OmCoCd= IdCoCd and OmTc= IdExpOdTc and OmYy= IdExpOdYy " + _
    '        "and OmChr= IdExpOdChr and OmNo= IdExpOdNo " + _
    '        "join OrdDsg on OdCoCd= IdCoCd and OdTc= IdExpOdTc and OdYy= IdExpOdYy " + _
    '        "and OdChr= IdExpOdChr and OdNo= IdExpOdNo and OdSr= IdExpOdSr " + wOrJoin + _
    '        " join InvHd on InCoCd= IdCoCd and InTc = IdTc and InYy= IdYy and InChr= IdChr and " + _
    '        "InNo= IdNo " + wCnd + " and RmCtg in ('D','C', 'G', 'P', 'S') " + _
    '        " group by IdTc, IdYy, IdChr, IdNo, IdSr, RmCtg, PSCd, RmKt, " + wRmRtOpt + " "
    '*** (Jen 2.12)
    '*** (Bef 2.12 Patch 2)
    
    '2.04-3 urmi DsgMst Join Clause added in above query: 'DmTcTyp= 'DM' and '
    
    '"left outer join CustDsg on CdCtg= 'C' and CdCd= OmCmCd and CdOurDmCd= IdDmCd and CdOurDmSz= '' and CdOurSfx= '' "
    
    '*** (Bef 2.13)
    'If UCase(adc.MenuCd) = UCase("RepInv") Or UCase(adc.MenuCd) = UCase("RepInvLoose") Then
    '  '*** (Bef 24/02/06) For Kama (Jen 2.12)
    '  'wSqlStrg = wHdStr1 + "Union All " + wStr1 + wStr2 + wStr3 + IIF(wFlagLocalExp = "Y", "Union All " + wStr1 + wStr4 + wStr5, "") + " Order By IdTc, IdYy, IdChr, IdNo, IdSr "
    '  '*** (Bef 24/02/06) For Kama (Jen 2.12)
    '
    '  '*** (24/02/06) For Kama (Jen 2.12)
    '  wSqlStrg = wHdStr1 + "Union All " + wStr1 + wStr2 + wStr3 + " Order By IdTc, IdYy, IdChr, IdNo, IdSr "
    '  '*** (24/02/06) For Kama (Jen 2.12)
    'Else
    '  wSqlStrg = wStr1 + wStr2 + wStr3 + " Order By IdTc, IdYy, IdChr, IdNo, IdSr "
    'End If
    '*** (Bef 2.13)

  '*** Jay 3.3(FG)
  Else
  '*** Jay 3.3(FG)

     '*** Jay 3.3(FG) [OdTc as added]
     ' **** Manali 3.5.0 - 11/12/08 - InHHideCommYN included
     ' **** Manali 3.6.0 - 17/11/09 - Circular No added
     ' **** Manali 3.8.1 - Self Payable In Invoice
    wStr1 = "Select '1D' as qTag, OdTc as IdTc, OdYy as IdYy, OdChr as IdChr, OdNo as IdNo, OdSr as IdSr, " + ws_InExpNo + ", max(Ih1.InCmCd) as InCmCd, max(Ih1.InPrnDt) as InDt, " + _
            "max(Ih1.InLmgCst) as InLmgCst, max(Ih1.InLmgDt) as InLmgDt, max(Ih1.InLmgSal) as InLmgSal, max(Ih1.InVaDC) as InVaDC, max(Ih1.InVaX) as InVaX, " + _
            "max(case when Ih1.InRsMulDiv= 'D' and convert (decimal (16, 4), Ih1.InCnvRt)> 0 then 1/ Ih1.InCnvRt else Ih1.InCnvRt end) as InCnvRt, " + _
            ws_PrdDesc + ", 0 as InPkgWt, max(Ih1.InGrNo) as InGrNo, max(Ih1.InGrDt) as InGrDt, max(Ih1.InSbNo) as InSbNo, " + _
            "max(Ih1.InSbDt) as InSbDt, max(Ih1.InAwbNo) as InAwbNo, max(Ih1.InAwbDt) as InAwbDt, max(Ih1.InHAwbNo) as InHAwbNo, max(Ih1.InHAwbDt) as InHAwbDt, max(Ih1.InHBinNo) as InHBinNo,  max(Ih1.InCERegnNo) as InCERegnNo, " + _
            "max(Ih1.InGspNo) as InGspNo, max(Ih1.InGspDt) as InGspDt, max(Ih1.InChqNo) as InChqNo, " + _
            "max(Ih1.InChqAmt) as InChqAmt, max(Ih1.InPoNo) as InPoNo, max(Ih1.InPoDt) as InPoDt, max(Ih1.InCmPayTerms) as InCmPayTerms, 0 as InFOB, " + _
            "0 as InFOBRs, 0 as InComm, 0 as InCommRs, 0 as InFrt, 0 as InFrtRs, 0 as InIns, 0 as InInsRs, " + _
            "0 as InDsc, 0 as InDscRs, max(Ih.InDscDesc) as InDscDesc, 0 as InDeduct, 0 as InDeductRs, max(Ih.InDeductDesc) as InDeductDesc, 0 as InTot, 0 as InTotRs, " + ws_InTotAs + ", 0 as InDiaCost, " + _
            "0 as InCSCost, 0 as InAdv, max(Ih.InDiaCostDesc) as InDiaCostDesc, max(Ih.InCSCostDesc) as InCSCostDesc, max(Ih.InAdvDesc) as InAdvDesc, "
    wStr1 = wStr1 + "max(Ih1.InCls) as InCls, max(Ih1.InHBkName) as InHBkName, max(Ih1.InHBkAdd1) as InHBkAdd1, max(Ih1.InHBkAdd2) as InHBkAdd2, max(Ih1.InHBkAdd3) as InHBkAdd3, " + _
            "max(Ih1.InHBkAc) as InHBkAc, max(Ih1.InHFrgnBkName) as InHFrgnBkName, max(Ih1.InHFrgnBkAc) as InHFrgnBkAc, max(Ih1.InHExpOff) as InHExpOff, max(Ih1.InHRbiNo) as InHRbiNo, " + _
            "max(Ih1.InHIecNo) as InHIecNo, max(Ih1.InHPreCarriage) as InHPreCarriage, max(Ih1.InHPreCarrier) as InHPreCarrier, max(Ih1.InHVessel) as InHVessel, " + _
            "max(Ih1.InHPortOfLd) as InHPortOfLd, max(Ih1.InHConOfOrg) as InHConOfOrg, max(Ih1.InHMarks) as InHMarks, max(Ih1.InHPkgDesc) as InHPkgDesc, " + _
            "max(Ih1.InHItcNo) as InHItcNo, max(Ih1.InHEtcZone) as InHEtcZone, max(Ih1.InHEtcSection) as InHEtcSection, max(Ih1.InHConsignee) as InHConsignee, " + _
            "max(Ih1.InCmInsBy) as InHInsBy, max(Ih1.InHInsNo) as InHInsNo, max(Ih1.InHAgName) as InHAgName, max(Ih1.InHAgAdd1) as InHAgAdd1, max(Ih1.InHAgAdd2) as InHAgAdd2, " + _
            "max(Ih1.InHAgAdd3) as InHAgAdd3, max(Ih1.InCmName) as InCmName, max(Ih1.InCmAdd1) as InCmAdd1, max(Ih1.InCmAdd2) as InCmAdd2, max(Ih1.InCmAdd3) as InCmAdd3, " + _
            "max(Ih1.InCmTel) as InCmTel, max(Ih1.InCmFax) as InCmFax, max(Ih1.InCmEMail) as InCmEMail, max(Ih1.InCmCpName) as InCmCpName, max(Ih1.InCmPortOfDisch) as InCmPortOfDisch, " + _
            "max(Ih1.InCmFinDest) as InCmFinDest, max(Ih1.InCmBuyer) as InCmBuyer, " + _
            "max(Ih1.InCmAgName) as InCmAgName, max(Ih1.InCmAgAdd1) as InCmAgAdd1, max(Ih1.InCmAgAdd2) as InCmAgAdd2, max(Ih1.InCmAgAdd3) as InCmAgAdd3, " + _
            "max(Ih1.InSelfName) as InSelfName, max(Ih1.InSelfAdd1) as InSelfAdd1, " + _
            "max(Ih1.InSelfAdd2) as InSelfAdd2, max(Ih1.InSelfAdd3) as InSelfAdd3, max(Ih1.InSelfTel) as InSelfTel, max(Ih1.InSelfFax) as InSelfFax, max(Ih1.InSelfEMail) as InSelfEMail, " + _
            "max(Ih1.InHLicNo) as InHLicNo, max(Ih1.InHETCSvg) as InHETCSvg, max(Ih1.InHAR4No) as InHAR4No, max(Ih1.InHQCertNo) as InHQCertNo, max(Ih1.InHGSPOrgCrit) as InHGSPOrgCrit, " + _
            "max(Ih1.InHStatCd) as InHStatCd, max(Ih1.InHFtRem) as InHFtRem, max(SubString(Ih1.InHBnkInstr,1, 700)) as InHBnkInstr, max(Ih1.InConAdd1) as InConAdd1, max(Ih1.InConAdd2) as InConAdd2, " + _
            "max(Ih1.InConAdd3) as InConAdd3, max(Ih1.InBuyAdd1) as InBuyAdd1, max(Ih1.InBuyAdd2) as InBuyAdd2, max(Ih1.InBuyAdd3) as InBuyAdd3, max(SubString(Ih1.InHUnResolutions, 1, 750)) as InHUnResolutions, max(Ih1.InHCustRem) as InHCustRem, " + _
            "max(Ih1.InHEximPolicy) as InHEximPolicy, max(Ih1.InHGspEligibility) as InHGspEligibility, max(Ih1.InMetalUtil) as InMetalUtil, max(Ih1.InhAuthSgn) as InhAuthSgn, " + _
            "max(Ih1.InHETCTxt) as InHETCTxt, max(Ih1.InHStatOfOrg) as InHStatOfOrg, max(Ih1.InHImpValTxt) as InHImpValTxt, max(Ih1.InHShpBillSgnBy) as InHShpBillSgnBy, max(Ih1.InHMetalBOETxt) as InHMetalBOETxt, max(Ih1.InHGspPkgDesc) as InHGspPkgDesc, max(Ih1.InBOELooseExp) as InBOELooseExp, " + _
            "max(Ih1.InHHideCommYN) as InHHideCommYN, max(Ih1.InHSelfDscYN) as InHSelfDscYN, max(Ih1.InHSelfDeductYN) as InHSelfDeductYN, max(Ih1.InCircularNo) as InCircularNo, "
    
      '*** Jay 3.3(FG)
      wStr2 = "max(OmCmCd) as OmCmCd,max(" + Ws_IdKt + ") as IdKt, max(DmCtg) as DmCtg, max(DmSalCtg) as DmSalCtg, max(DmCd) as IdDmCd, max(" + ws_IdSfx + ") as IdSfx, max(" + ws_IdDmSz + ") as IdDmSz, max(DmUom) as DmUom, " + wDetDesc2Str + _
            " max(OdDmCol) as OdDmCol, max(" + ws_IdVaCtg + ") as IdVaCtg, 0 as IdGldLs, max(" + ws_IdQty + ") as IdQty, " + _
            "sum(OrWt / (case when R1.RmCtg in ('C', 'D') then 5 else 1 end)) as qInvGrWt, " + _
            "sum(OrWt * (case R1.RmCtg when 'G' then 1 else 0 end)) as qInvGldWt, " + _
            "sum(OrWt * (case R1.RmCtg when 'P' then 1 else 0 end)) as qInvPlWt, " + _
            "sum(OrWt * (case R1.RmCtg when 'S' then 1 else 0 end)) as qInvSlWt, " + _
            "sum(OrWt * (case when R1.RmCtg in ('L','M') then 1 else 0 end)) as qInvPdWt, " + _
            "sum(OrWt * (case R1.RmCtg when 'X' then 1 else 0 end)) as qInvAccWt, " + _
            "sum(OrWt * (case when R1.RmCtg in ('D') then 1 else 0 end)) as qInvDWt, " + _
            "sum(OrWt * (case when R1.RmCtg in ('C') then 1 else 0 end)) as qInvCWt, " + _
            "sum(OrWt * R1.RmPurityZ * (case R1.RmCtg when 'G' then 1 else 0 end)/ (case when R2.RmPurityZ> 0 then R2.RmPurityZ else 1 end)) as qInvPureGldWt, " + _
            "sum(OrWt * R1.RmPurityZ * (case R1.RmCtg when 'P' then 1 else 0 end)/ (case when R2.RmPurityZ> 0 then R2.RmPurityZ else 1 end)) as qInvPurePlWt, " + _
            "sum(OrWt * R1.RmPurityZ * (case R1.RmCtg when 'S' then 1 else 0 end)/ (case when R2.RmPurityZ> 0 then R2.RmPurityZ else 1 end)) as qInvPureSlWt, " + _
            "sum(OrWt * R1.RmPurityZ * (case when R1.RmCtg in ('L','M') then 1 else 0 end)/ (case when R2.RmPurityZ> 0 then R2.RmPurityZ else 1 end)) as qInvPurePdWt, " + _
            "0 as qInvPureGldLs, " + _
            "0 as qInvPurePlLs, " + _
            "0 as qInvPureSlLs, " + _
            "0 as qInvPurePdLs, " + _
            "sum(OrSalVal * (case R1.RmCtg when 'G' then 1 else 0 end)) as qIGldVal, " + _
            "sum(OrSalVal * (case R1.RmCtg when 'P' then 1 else 0 end)) as qIPlVal, " + _
            "sum(OrSalVal * (case R1.RmCtg when 'S' then 1 else 0 end)) as qISlVal, " + _
            "sum(OrSalVal * (case when R1.RmCtg in ('L','M') then 1 else 0 end)) as qIPdVal, " + _
            "sum(OrSalVal * (case when R1.RmCtg in ('D') then 1 else 0 end)) as qIDVal, " + _
            "sum(OrSalVal * (case when R1.RmCtg in ('C') then 1 else 0 end)) as qICVal, " + _
            "sum(OrSetSalVal) as qISetVal, " + wILabValStr + " as qILabVal, max(OdSalPrc) as IdiFOB, "
    
    wStr3 = "sum(OrWt / (case when R1.RmCtg in ('C', 'D') then 5 else 1 end)) as qZGrWt, " + _
            "sum(OrWt * (case R1.RmCtg when 'G' then 1 else 0 end)) as qZGldWt, " + _
            "sum(OrWt * (case R1.RmCtg when 'P' then 1 else 0 end)) as qZPlWt, " + _
            "sum(OrWt * (case R1.RmCtg when 'S' then 1 else 0 end)) as qZSlWt, " + _
            "sum(OrWt * (case when R1.RmCtg in ('L','M') then 1 else 0 end)) as qZPdWt, " + _
            "sum(OrWt * (case R1.RmCtg when 'X' then 1 else 0 end)) as qZAccWt, sum(OrWt * (case when R1.RmCtg in ('D') then 1 else 0 end)) as qZDWt, " + _
            " sum(OrWt * (case when R1.RmCtg in ('C') then 1 else 0 end)) as qZCWt, " + _
            "sum(OrWt * R1.RmPurityZ * (case R1.RmCtg when 'G' then 1 else 0 end)/ (case when R2.RmPurityZ> 0 then R2.RmPurityZ else 1 end)) as qZPureGldWt, " + _
            "sum(OrWt * R1.RmPurityZ * (case R1.RmCtg when 'P' then 1 else 0 end)/ (case when R2.RmPurityZ> 0 then R2.RmPurityZ else 1 end)) as qZPurePlWt, " + _
            "sum(OrWt * R1.RmPurityZ * (case R1.RmCtg when 'S' then 1 else 0 end)/ (case when R2.RmPurityZ> 0 then R2.RmPurityZ else 1 end)) as qZPureSlWt, " + _
            "sum(OrWt * R1.RmPurityZ * (case when R1.RmCtg in ('L','M') then 1 else 0 end)/ (case when R2.RmPurityZ> 0 then R2.RmPurityZ else 1 end)) as qZPurePdWt, " + _
            "0 as qZPureGldLs,0 as qZPurePlLs,0 as qZPureSlLs,0 as qZPurePdLs, " + _
            "sum(OrSalVal * (case R1.RmCtg when 'G' then 1 else 0 end)) as qZGldVal, " + _
            "sum(OrSalVal * (case R1.RmCtg when 'P' then 1 else 0 end)) as qZPlVal, " + _
            "sum(OrSalVal * (case R1.RmCtg when 'S' then 1 else 0 end)) as qZSlVal, " + _
            "sum(OrSalVal * (case R1.RmCtg when 'S' then 1 else 0 end)) as qZPdVal, " + _
            "sum(OrSalVal * (case when R1.RmCtg in ('D') then 1 else 0 end)) as qZDVal, " + _
            "sum(OrSalVal * (case when R1.RmCtg in ('C') then 1 else 0 end)) as qZCVal, "
    
    wStr3 = wStr3 + "sum(OrSetSalVal) as qZSetVal, " + wZLabValStr + " as qZLabVal, max(OdSalPrc) as IdZFOB, " + qMainMet + ", " + qMainRmCtgDesc + ", " + qDmCtgDesc + ", " + qDmSalCtgDesc + ", " + qDmColDesc + ", " + qDmKtDesc + ", " + qVaCtgDesc + ", max(Ih1.InMMTC) as InMMTC," + _
            " max(Ih1.InHGldFiness) as InHGldFiness, max(Ih1.InHPlFiness) as InHPlFiness, max(Ih1.InHSlFiness) as InHSlFiness,  max(Ih1.InHPdFiness) as InHPdFiness, max(Ih1.InHGjepcNo) as InHGjepcNo, max(Ih1.InHGjepcValidUpTo) as InHGjepcValidUpTo, Max(Ih1.InHRbiPolicy) as InHRbiPolicy, " + _
            " '' as RmCtg, '' as RmKt, 0 as qRmRt, 0 as IrRmQty, 0 as qWtInCts, 0 as qWtInGms, 0 as qVal " + _
            " from OrdRm join RmMst R1 on " + IIF(gs_Partition = ctCurrPrtn, " R1.RmPrtKey=OrPrtKey and ", "") + " R1.RmCd= OrRmCd join RmMst R2 on R2.RmCd= R1.RmZCd " + IIF(gs_Partition = ctCurrPrtn, " And R2.RmPrtKey=R1.RmPrtKey ", "") + _
            " join OrdDsg on OdIdNo=OrOdIdNo and OdCoCd= OrCoCd and OdTc= OrTc and OdYy= OrYy and OdChr= OrChr " + _
            " and OdNo= OrNo and OdSr= OrSr " + IIF(gs_Partition = ctCurrPrtn, " And OdPrtKey=OrPrtKey ", "") + _
            " join DsgMst on DmIdNo=OdDmIdNo and DmTcTyp= 'DM' and DmCd = OdDmCd and DmSz= '' " + IIF(gs_Partition = ctCurrPrtn, " And DmPrtKey=OdPrtKey ", "") + _
            " join OrdMst on OmCoCd= OdCoCd and OmTc= OdTc and OmYy= OdYy " + _
            " and OmChr= OdChr and OmNo= OdNo " + IIF(gs_Partition = ctCurrPrtn, " And OmPrtKey=OdPrtKey ", "") + _
            " join InvHd Ih on Ih.InCoCd= OdCoCd and Ih.InTc = OdTc and Ih.InYy= OdYy and Ih.InChr= OdChr and Ih.InNo= OdNo " + IIF(gs_Partition = ctCurrPrtn, " And Ih.InPrtKey=OdPrtKey ", "") + _
            " join InvHd Ih1 on " + ws_MinCnd + wCnd + IIF(gs_Partition = ctCurrPrtn, " And OrPrtKey='" + ctCurrPrtn + "' ", "") + " Group by OdTc, OdYy, OdChr, OdNo, OdSr "
  '*** Jay 3.3(FG)
  End If
  '*** Jay 3.3(FG)
    
  '*** (Jen 2.13)
  wSqlStrg = wHdStr1 + " Union All " + wStr1 + wStr2 + wStr3 + " Order By qTag, IdTc, IdYy, IdChr, IdNo, IdSr "
  '*** (Jen 2.13)
  
  '4.1.0.0
'  If moCn.RecSeek(" Select 'x' From InvFgd Join InvHd On InCoCd=IfCoCd And InTc=IfTc And InYy=IfYy And InChr=IfChr And InNo=IfNo Where " + _
                    IIF(adc("UoInNoFr") <> 0, "IfTc= '" + adc("UoInTcFr") + "' and IfYy= '" + adc("UoInYyFr") + "' and IfChr= '" + adc("UoInChrFr") + "' and " + _
                    " IfNo= " + CStr(adc("UoInNoFr")), " InExpNo In (" + IIF(adc("UOINEXPNOFR") <> "", "'" + adc("UOINEXPNOFR") + "'", adc("UOINEXPNOSEL"))) + ")" + " and IfSr=0 ") Then
 '   adc.RepRecSource = "Select * From InvHd Where 1=2" + ctMaxDopOpt '****** Sachin 3.01 added ctMaxDopOpt
  '  DispMsg " IfSr=0 Record Exist Cannot Display Report", etError: Exit Sub
  'sv.40 company code added
   If moCn.RecSeek(" Select 'x' From InvFgd Join InvHd On InCoCd=IfCoCd And InTc=IfTc And InYy=IfYy And InChr=IfChr And InNo=IfNo Where " + _
                    IIF(ADC("UoInNoFr") <> 0, "IfCoCd='" + ADC("UoCoCdFr") + "' and IfTc= '" + ADC("UoInTcFr") + "' and IfYy= '" + ADC("UoInYyFr") + "' and IfChr= '" + ADC("UoInChrFr") + "' and " + _
                    " IfNo= " + CStr(ADC("UoInNoFr")), " InExpNo In (" + IIF(ADC("UOINEXPNOFR") <> "", "'" + ADC("UOINEXPNOFR") + "'", ADC("UOINEXPNOSEL")) + ")") + "" + " and IfSr=0 ") Then
   
    ADC.RepRecSource = "Select * From InvHd Where 1=2" + ctMaxDopOpt '****** Sachin 3.01 added ctMaxDopOpt
    DispMsg " IfSr=0 Record Exist Cannot Display Report", etError: Exit Sub
  Else
    If ADC("UOINEXPNOSEL") <> "" Then
        If moCn.GetFldVal(" Select Count(distinct InCmCd) From InvHd  Where InCoCd= '" + gs_CoCd + "' and InExpNo In (" + ADC("UOINEXPNOSEL") + ")") > 1 Then
            ADC.RepRecSource = "Select * From InvHd Where 1=2" + ctMaxDopOpt
            DispMsg " All invoices should be of Same Customer.", etError: Exit Sub
        End If
     End If
        ADC.RepRecSource = wSqlStrg + ctMaxDopOpt  '****** Sachin 3.01 ctMaxDopOpt added
  End If
        
        
 Dim wRsInHd As MDORowSet
  '*** (Bef 2.13)
  'Set wRsInHd = moCn.OpenRes("Select InFOB, InFOBRs, InComm, InCommRs, InFrt, InFrtRs, " + _
  '              "InIns, InInsRs, InTot, InTotRs, InTotAs, InDiaCost, InCSCost, InAdv, " + _
  '              "InExpNo, InPrnDt as InDt, InHFrgnBkName, InHBkName, InHBkAdd1, InHBkAdd2, " + _
  '              "InHBkAdd3, InSelfName, InHBkAc, InCnvRt " + _
  '              "from InvHd where InCoCd= '" + adc("UoCoCdFr") + "' and " + _
  '              "InTc= '" + adc("UoInTcFr") + "' and InYy= '" + adc("UoInYyFr") + "' and " + _
  '              "InChr= '" + adc("UoInChrFr") + "' and InNo= " + CStr(adc("UoInNoFr")) + " ")
  '*** (Bef 2.13)
    
  '*** (Jen 2.13)
  'max(Ih1.InCnvRt) as InCnvRt
  '*** Jay 3.3(FG) [when 'FG' then 'C' ]
  ' ***** Manali Trading Module - When 'JT' Then OmCmCtg
  Set wRsInHd = moCn.OpenRes("Select isnull(max(Ih.InNo),0) as InNo, sum(Ih.InFOB) as InFOB, sum(Ih.InFOBRs) as InFOBRs, sum(Ih.InComm) as InComm, " + _
                "sum(Ih.InCommRs) as InCommRs, sum(Ih.InFrt) as InFrt, sum(Ih.InFrtRs) as InFrtRs, " + _
                "sum(Ih.InIns) as InIns, sum(Ih.InInsRs) as InInsRs, sum(Ih.InTot) as InTot, sum(Ih.InTotRs) as InTotRs, " + _
                ws_InTotAs + ", sum(Ih.InDiaCost) as InDiaCost, sum(Ih.InCSCost) as InCSCost, " + _
                "sum(Ih.InAdv) as InAdv, " + ws_InExpNo + ", max(Ih1.InPrnDt) as InDt, max(Ih1.InHFrgnBkName) as InHFrgnBkName, " + _
                "max(Ih1.InHBkName) as InHBkName, max(Ih1.InHBkAdd1) as InHBkAdd1, max(Ih1.InHBkAdd2) as InHBkAdd2, " + _
                "max(Ih1.InHBkAdd3) as InHBkAdd3, max(Ih1.InSelfName) as InSelfName, " + _
                "max(Ih1.InHBkAc) as InHBkAc, max(case when Ih1.InRsMulDiv= 'D' and convert (decimal (16, 4), Ih1.InCnvRt)> 0 then 1/ Ih1.InCnvRt else Ih1.InCnvRt end) as InCnvRt, " + _
                "max(CmCurCd) as CmCurCd " + _
                "from InvHd Ih " + _
                "join Param on PTyp= 'TC' and PMCd= Ih.InTc and PSCd= '' " + _
                "join CustMst on CmCtg= (case PValue when 'IN' then 'C' when 'CTB' then 'T' when 'FG' then 'C' When 'JT' Then (Select OmCmCtg from OrdMst Where OmCoCd=InCoCd and OmTc=InTc and OmChr=InChr and OmNo=InNo and OmPrtKey=InPrtKey) else '' end) and CmCd= Ih.InCmCd " + _
                "join InvHd Ih1 on " + ws_MinCnd + wCnd + IIF(gs_Partition = ctCurrPrtn, " And Ih.InPrtKey='" + ctCurrPrtn + "' ", ""))
  
  If wRsInHd.RecCount > 0 And wRsInHd!InNo <> 0 Then
    ws_CurCd = UCase(Trim(wRsInHd!CmCurCd))

    wd_GldOunce = moCn.GetFldVal("Select PNum From Param Where PTyp= 'GMFACTOR' and PMCd= 'G' and PSCd= '" + ws_CurCd + "'")
    wd_PlOunce = moCn.GetFldVal("Select PNum From Param Where PTyp= 'GMFACTOR' and PMCd= 'P' and PSCd= '" + ws_CurCd + "'")
    wd_SlOunce = moCn.GetFldVal("Select PNum From Param Where PTyp= 'GMFACTOR' and PMCd= 'S' and PSCd= '" + ws_CurCd + "'")
    ' **** Manali 3.03 - 25/06/08 - Pd Mod
    wd_PdOunce = moCn.GetFldVal("Select PNum From Param Where PTyp= 'GMFACTOR' and PMCd= 'L' and PSCd= '" + ws_CurCd + "'")
    ' **** Manali 3.03 - 25/06/08 - Pd Mod
    
    ws_GldOunceDsc = IIF(ctDefCurncy = ws_CurCd, "Troy Ounce", CStr(wd_GldOunce) + " Grams")
    ws_PlOunceDsc = IIF(ctDefCurncy = ws_CurCd, "Troy Ounce", CStr(wd_PlOunce) + " Grams")
    ws_SlOunceDsc = IIF(ctDefCurncy = ws_CurCd, "Troy Ounce", CStr(wd_SlOunce) + " Grams")
    ' **** Manali 3.03 - 25/06/08 - Pd Mod
    ws_PdOunceDsc = IIF(ctDefCurncy = ws_CurCd, "Troy Ounce", CStr(wd_PdOunce) + " Grams")
    ' **** Manali 3.03 - 25/06/08 - Pd Mod
  End If
  '*** (Jen 2.13)
  '*****Bhavna addad for ce regn no
  
  If UCase(ADC.MenuCd) = UCase("RepInv") Or UCase(ADC.MenuCd) = UCase("RepShpBill") Or _
  UCase(ADC.MenuCd) = UCase("RepShpBillMulti") Or UCase(ADC.MenuCd) = UCase("RepInvLoose") Or _
  UCase(ADC.MenuCd) = UCase("RepShpBillLoose") Or _
  (wFlagLocalExp = "Y" And (UCase(ADC.MenuCd) = UCase("RepGRFormFrnt") Or UCase(ADC.MenuCd) = UCase("RepGrFormLoose") Or _
   UCase(ADC.MenuCd) = UCase("RepFgGRFrnt") Or UCase(ADC.MenuCd) = UCase("RepJtGRFrnt") Or _
   UCase(ADC.MenuCd) = UCase("RepGrFormFrntMulti"))) Then
    If UCase(ADC("UoYN10")) = "Y" Then
        mRep.TxtBinNo.Suppress = True
        mRep.FldBinNo.Suppress = True
        mRep.FldCERegnNo.Suppress = False

    Else
        mRep.FldBinNo.Suppress = False
        mRep.FldCERegnNo.Suppress = True
        mRep.TxtCERegnNo.Suppress = True

    End If
  Else
'    If UCase(adc("UoYN10")) = "Y" And ChkBoxArr(10).Visible = True Then mRep.TxtBinNo.SetText "CE Regn. No."
  End If
  '*****Bhavna addad for ce regn no
  
  Select Case UCase(ADC.MenuCd)
  '*** Jay 3.3(FG)
  ' ***** Manali Trading Module
  Case Is = UCase("RepInv"), UCase("RepInvLoose"), UCase("RepFgInv"), UCase("RepJtInv")
  '*** Jay 3.3(FG)
    GRP_REP.SetFormula mRep, "wGrp1", wGrpFormula
    GRP_REP.SetFormula mRep, "wGrp1Desc", wGrpDescFormula
    GRP_REP.SetFormula mRep, "wDet", wDetFormula
    GRP_REP.SetFormula mRep, "wDetDesc", wDetDescFormula
    GRP_REP.SetFormula mRep, "wDetDesc2", wDetDesc2Formula
    
    '***Geeta***Emr208
    If UCase(ADC("UoInDet")) = "" Then GRP_REP.SetFormula mRep, "wDet", wGrpFormula
    '***Geeta***Emr208
    
    '*** (Bef 2.13)
    'If wRsInHd.RecCount > 0 Then _
    'SetRepText mRep, "TxtAmtInWrds", "Total " + wRsInHd!InTotAs + " " + FnNumToWords(Val(Format(wRsInHd!InTot - wRsInHd!InDiaCost - wRsInHd!InCSCost - wRsInHd!InAdv, "######0.00")), "USD")
    '*** (Bef 2.13)
    
    '*** (Jen 2.13)
    If wRsInHd.RecCount > 0 And wRsInHd!InNo <> 0 Then _
    SetRepText mRep, "TxtAmtInWrds", "Total " + "Net Realisation" + " " + FnNumToWords(Val(Format(wRsInHd!InTot - wRsInHd!InDiaCost - wRsInHd!InCSCost - wRsInHd!InAdv, "######0.00")), ws_CurCd)
    '*** (Jen 2.13)
    
    '***(Bef 1.11) SetRepText mRep, "TxtAmtInWrds", "Total " + wRsInHd!InTotAs + " " + FnNumToWords(Val(Format(wRsInHd!InTot - wRsInHd!InDiaCost - wRsInHd!InCSCost - wRsInHd!InAdv, "######0")), "USD")
    'InDiaCost, InCSCost, InAdv,
    'SetRepText mRep, "TxtAmtInWrds", "Total " + wRsInHd!InTotAs + " " + FnNumToWords(Val(Format(wRsInHd!InTot, "######0")), "USD")
    '*** If wRsInHd.RecCount > 0 Then _
    '*** SetRepText mRep, "TxtInstr", "PLEASE INSTRUCT YOUR BANKERS TO TRANSFER THE INVOICE VALUE " + _
                     "US$   " + Format(wRsInHd!InTot, "######0.00") + "   QUOTING OUR INVOICE NO.   " + _
                     wRsInHd!InExpNo + "   DATED   " + Format(wRsInHd!InPrnDt, "dd/MM/yyyy") + _
                     "   BY AN AUTHENTICATED CABLE TESTED TELEX TO   " + wRsInHd!InHFrgnBkName + _
                     "   FOR THE CREDIT OF OUR A/C WITH   " + wRsInHd!InHBkName + ", " + _
                     wRsInHd!InHBkAdd1 + wRsInHd!InHBkAdd2 + wRsInHd!InHBkAdd3 + _
                     "   FAVOURING M/S.   " + wRsInHd!InSelfName + ",   CURRENT A/C NO. -   " + _
                     wRsInHd!InHBkAc
    
    If UCase(ADC("UoYN")) = "Y" Then
      mRep.PICSEC.Suppress = True
      ' **** Zubin 211 **** '
      If UCase(ADC("UoInDet")) = "DMCD" Or UCase(ADC("UoInDet")) = "INVSR" Or UCase(ADC("UoInDet")) = "INVSR/SR" Or _
         UCase(ADC("UoInDet")) = "DMCD/SR" Then mRep.PICSEC.Suppress = False
      ' **** Zubin 211 **** '
    Else
        mRep.PICSEC.Suppress = True
    End If
  
    '******************************** urmi *********************************************/
    If UCase(ADC("UoYN1")) = "Y" Then
      mRep.RemSec.Suppress = True
      ' **** Zubin 211 **** '
      If UCase(ADC("UoInDet")) = "DMCD" Or UCase(ADC("UoInDet")) = "INVSR" Or UCase(ADC("UoInDet")) = "INVSR/SR" Or _
         UCase(ADC("UoInDet")) = "DMCD/SR" Then mRep.RemSec.Suppress = False
      ' **** Zubin 211 **** '
    Else
        mRep.RemSec.Suppress = True
    End If
    '******************************** urmi *********************************************/
    
  
    '*** Option for showing/ hiding the Metal Loss Sub Report
    
    'Changed below in 4.1.3 to single line as giving error 'Procedure to large'
    mRep.Subreport1.Suppress = IIF(UCase(wShowMetLsSumm) = "Y", False, True)
        
'    If UCase(wShowMetLsSumm) = "Y" Then
'      mRep.Subreport1.Suppress = False
'      '*********Geeta****Emr208******
'    Else
'      mRep.Subreport1.Suppress = True
'      '*********Geeta****Emr208******
'    End If
  
  
  If UCase(wShowCnvRt) = "N" Then
      mRep.TxtCnvRt.Suppress = True: mRep.FldCnvRt.Suppress = True
      mRep.FldInvCommRs.Suppress = True: mRep.FldInvFOBRs.Suppress = True
      mRep.FldInvFrtRs.Suppress = True: mRep.FldInvInsRs.Suppress = True
      mRep.FldInvTotRs.Suppress = True: mRep.FldDiaCostRs.Suppress = True
      mRep.FldCSCostRs.Suppress = True: mRep.FldAdvRs.Suppress = True
      mRep.LblValInRs.Suppress = True
  End If
  
      'mRep.TxtCnvRt.Suppress = IIF(UCase(wShowCnvRt) = "Y", False, True)
      'mRep.FldCnvRt.Suppress = IIF(UCase(wShowCnvRt) = "Y", False, True)
      'mRep.FldInvFOBRs.Suppress = IIF(UCase(wShowCnvRt) = "Y", False, True)
      'mRep.FldInvFrtRs.Suppress = IIF(UCase(wShowCnvRt) = "Y", False, True)
'      mRep.FldInvInsRs.Suppress = IIF(UCase(wShowCnvRt) = "Y", False, True)
'      mRep.FldInvTotRs.Suppress = IIF(UCase(wShowCnvRt) = "Y", False, True)
'      mRep.FldDiaCostRs.Suppress = IIF(UCase(wShowCnvRt) = "Y", False, True)
'      mRep.FldCSCostRs.Suppress = IIF(UCase(wShowCnvRt) = "Y", False, True)
'      mRep.FldAdvRs.Suppress = IIF(UCase(wShowCnvRt) = "Y", False, True)
'      mRep.LblValInRs.Suppress = IIF(UCase(wShowCnvRt) = "Y", False, True)
'      mRep.FldInvCommRs.Suppress = True
 
  
    '*** Option for showing/ hiding the Conversion Rate/ Rupee equivalent
'    If UCase(wShowCnvRt) = "Y" Then
'    'If UCase(ADC("UoYN3")) = "Y" Then
'      mRep.TxtCnvRt.Suppress = False: mRep.FldCnvRt.Suppress = False
'      'mRep.FldInvCommRs.Suppress = False  ' **** Manali 3.5.0 - 11/12/08 - InHHideCommYN included
'      mRep.FldInvFOBRs.Suppress = False
'      mRep.FldInvFrtRs.Suppress = False: mRep.FldInvInsRs.Suppress = False
'      mRep.FldInvTotRs.Suppress = False: mRep.FldDiaCostRs.Suppress = False
'      mRep.FldCSCostRs.Suppress = False: mRep.FldAdvRs.Suppress = False
'      '****Geeta******Emr208******
'    Else
'      mRep.TxtCnvRt.Suppress = True: mRep.FldCnvRt.Suppress = True
'      mRep.FldInvCommRs.Suppress = True: mRep.FldInvFOBRs.Suppress = True
'      mRep.FldInvFrtRs.Suppress = True: mRep.FldInvInsRs.Suppress = True
'      mRep.FldInvTotRs.Suppress = True: mRep.FldDiaCostRs.Suppress = True
'      mRep.FldCSCostRs.Suppress = True: mRep.FldAdvRs.Suppress = True
'      '****Geeta******Emr208******
'    End If
    
    '*** Option for showing/ hiding the Rate Column
    If UCase(wShowRtFld) = "Y" Then
      mRep.TxtRt.Suppress = False
      
      '*** (24/02/06) For Kama (Jen 2.12) ###$$$%%%
      mRep.FldRt.Suppress = False
      '*** (24/02/06) For Kama (Jen 2.12) ###$$$%%%
      
    Else
      mRep.TxtRt.Suppress = True

      '*** (24/02/06) For Kama (Jen 2.12) ###$$$%%%
      mRep.FldRt.Suppress = True
      '*** (24/02/06) For Kama (Jen 2.12) ###$$$%%%
    End If
    
    
  
    '*** (Bef 2.13)
    'wrepcnd = adc.RepCond
    'wCnd = IIF(wrepcnd <> "", " and " + wrepcnd, "")
    '*** (Bef 2.13)
    
    'wSubRepWtStr = "sum(IrRmIWt) as qMetWt, " + _
                   "sum(IrRmIWt * RmPurityZ) as qMetPureWt, " + _
                   "sum(IrRmIWt * IdGldLs/100) as qLsWt, " + _
                   "sum(IrRmIWt * RmPurityZ * IdGldLs/100) as qPureLsWt, " + _
                   "sum(IrRmIWt * (1 + IdGldLs/100)) as qTotWt, " + _
                   "sum(IrRmIWt * RmPurityZ * (1 + IdGldLs/100)) as qPureTotWt "
    
        
    '*** (Bef 2.13)
    'If moCn.GetFldVal("Select count(*) from InvRm join RmMst on RmCd= IrRmCd " + _
    '   "where IrCoCd= '" + adc("UoCoCdFr") + "' and IrTc= '" + adc("UoInTcFr") + "' and " + _
    '   "IrYy= '" + adc("UoInYyFr") + "' and IrChr= '" + adc("UoInChrFr") + "' and " + _
    '   "IrNo= " + CStr(adc("UoInNoFr")) + " and RmCtg in ('G', 'P', 'S') ") <= 0 Then
    '  mRep.Subreport1.Suppress = True
    'End If
    '*** (Bef 2.13)
    
    '*** (Jen 2.13)
    ' **** Manali 3.03 - 25/06/08 - Pd Mod - 'L' added
    '*** Jay 3.3(FG)
    ' ***** Manali Trading Module
    If UCase(ADC.MenuCd) <> UCase("RepFgInv") And UCase(ADC.MenuCd) <> UCase("RepJtInv") Then
    '*** Jay 3.3(FG)
      If moCn.GetFldVal("Select count(*) from InvRm join RmMst on RmCd= IrRmCd " + IIF(gs_Partition = ctCurrPrtn, " And RmPrtKey=IrPrtKey", "") + _
         " join InvHd Ih on Ih.InCoCd= IrCoCd and Ih.InTc= IrTc and Ih.InYy= IrYy and Ih.InChr= IrChr and Ih.InNo= IrNo " + IIF(gs_Partition = ctCurrPrtn, " And Ih.InPrtKey=IrPrtKey", "") + _
         " join InvHd Ih1 on " + ws_MinCnd + wCnd + " and RmCtg in ('G', 'P', 'S', 'L','M') ") <= 0 Then
        mRep.Subreport1.Suppress = True
      End If
    '*** Jay 3.3(FG)
    Else
      If moCn.GetFldVal("Select count(*) from OrdRm join RmMst on RmCd= OrRmCd " + IIF(gs_Partition = ctCurrPrtn, " And RmPrtKey=OrPrtKey", "") + _
         " join InvHd Ih on Ih.InCoCd= OrCoCd and Ih.InTc= OrTc and Ih.InYy= OrYy and Ih.InChr= OrChr and Ih.InNo= OrNo " + IIF(gs_Partition = ctCurrPrtn, " And Ih.InPrtKey=OrPrtKey", "") + _
         " join InvHd Ih1 on " + ws_MinCnd + wCnd + " and RmCtg in ('G', 'P', 'S', 'L','M') ") <= 0 Then
        mRep.Subreport1.Suppress = True
      End If
    End If
    '*** Jay 3.3(FG)
    
    '*** (Bef 04/08/06)
    'wSubRepStr = "Select max(InTc) as InTc, max(InYy) as InYy, max(InChr) as InChr, " + _
                 "max(InNo) as InNo, IdGldLs, max(IdKt) as IdKt, max(IrRmCd) as IrRmCd, " + _
                 "RmCtg as RmCtg, (Select PDesc from Param where PTyp= 'RMCTG' and PMCd= RmCtg) as RmCtgDesc, " + _
                 "sum(" + wRmWtOpt + ") as qMetWt, " + _
                 "sum(" + wRmWtOpt + " * RmPurityZ) as qMetPureWt, " + _
                 "sum(" + wRmWtOpt + " * RmPurityZ/(case when (InHGldFiness> 0 and InHPlFiness> 0 and InHSlFiness> 0) then (case RmCtg when 'G' then InHGldFiness when 'P' then InHPlFiness when 'S' then InHSlFiness else 1 end) else 1 end)) as qMetFineWt, " + _
                 "sum(" + wRmWtOpt + " * IdGldLs/100) as qInvLsWt, " + _
                 "sum(" + wRmWtOpt + " * RmPurityZ * IdGldLs/100) as qInvPureLsWt, " + _
                 "sum(" + wRmWtOpt + " * RmPurityZ * IdGldLs/100/(case when (InHGldFiness> 0 and InHPlFiness> 0 and InHSlFiness> 0) then (case RmCtg when 'G' then InHGldFiness when 'P' then InHPlFiness when 'S' then InHSlFiness else 1 end) else 1 end)) as qInvFineLsWt, " + _
                 "sum(" + wRmWtOpt + " * (1 + IdGldLs/100)) as qInvTotWt, " + _
                 "sum(" + wRmWtOpt + " * RmPurityZ * (1 + IdGldLs/100)) as qInvPureTotWt, " + _
                 "sum(" + wRmWtOpt + " * RmPurityZ * (1 + IdGldLs/100)/(case when (InHGldFiness> 0 and InHPlFiness> 0 and InHSlFiness> 0) then (case RmCtg when 'G' then InHGldFiness when 'P' then InHPlFiness when 'S' then InHSlFiness else 1 end) else 1 end)) as qInvFineTotWt " + _
                 " from InvRm join RmMst on RmCd = IrRmCd " + _
                 "join InvDsg on IdCoCd= IrCoCd and IdTc= IrTc and IdYy= IrYy and " + _
                 "IdChr= IrChr and IdNo= IrNo and IdSr= IrSr " + _
                 "join InvHd Ih on InCoCd= IdCoCd and InTc = IdTc and InYy= IdYy and " + _
                 "InChr= IdChr and InNo= IdNo where RmCtg in ('G', 'P', 'S') " + _
                 wAndCnd + " Group By RmCtg, " + IIF(wFlagLocalExp = "Y", "IdKt ", "IrRmCd ") + ", IdGldLs "
    '*** (Bef 04/08/06)
    
    '****** Sachin 2.14.0 Next ******
    wb_MLsRmCd = moCn.RecSeek("Select 'x' From Head Where HCd='" + ctSelfCmCd + "' And HCoCd='" + ctSelfCoCd + "' And HInRepMetLsRmCd='Y'")
    '****** Sachin 2.14.0 Next ******
        
    
    '****** (07/11/06) added field RmKt in the Gold Loss Summary Sub Report ******'
    '*** (Jen 2.14 Next Patch 1 05/06/07)
    '****** Sachin 3.02 - Id fields in Joins
    ' **** Manali 3.03 - 25/06/08 - Pd Mod - 'L' added
    '*** Jay 3.3(FG)
    ' ***** Manali Trading Module
    If UCase(ADC.MenuCd) <> UCase("RepFgInv") And UCase(ADC.MenuCd) <> UCase("RepJtInv") Then
    '*** Jay 3.3(FG)
      wSubRepStr = "Select max(InTc) as InTc, max(InYy) as InYy, max(InChr) as InChr, " + _
                 "max(InNo) as InNo,(Case when R1.RmCtg in('X','M') Then 0 Else IdGldLs End) as IdGldLs,/* max(" + Ws_IdKt + ") */ Max(IdKt) as IdKt, max(IrRmCd) as IrRmCd, " + _
                 "R1.RmCtg as RmCtg, max(R1.RmKt) as RmKt, (Select PDesc from Param where PTyp= 'RMCTG' and PMCd= R1.RmCtg) as RmCtgDesc, " + _
                 "sum(" + wRmWtOpt + ") as qMetWt, " + _
                 "sum(" + wRmWtOpt + " * (Case when R1.RmCtg='X' Then B.RmPurityZ Else R1.RmPurityZ End)/ (case when R2.RmPurityZ> 0 then R2.RmPurityZ else 1 end)) as qMetPureWt, " + _
                 "sum(" + wRmWtOpt + " * (Case when R1.RmCtg='X' Then B.RmPurityZ Else R1.RmPurityZ End)/ (case when R2.RmPurityZ> 0 then R2.RmPurityZ else 1 end)/ (case when (InHGldFiness> 0 and InHPlFiness> 0 and InHSlFiness> 0 and InHPdFiness> 0) then (case R1.RmCtg when 'G' then InHGldFiness when 'P' then InHPlFiness when 'S' then InHSlFiness when 'L' then InHPdFiness else 1 end) else 1 end)) as qMetFineWt, " + _
                 "sum(" + IIF(UCase(ADC("UoInWt")) = "Z", " (case when IrMainMet= 'Y' then IdCustmMtchWt- " + wRmWtOpt + " else " + wRmWtOpt + "* (Case when R1.RmCtg in ('X','M') Then 0 Else IdGldLs End)/100 end) ", wRmWtOpt + "* (Case when R1.RmCtg in ('X','M') Then 0 Else IdGldLs End)/100") + ") as qInvLsWt, " + _
                 "sum(" + IIF(UCase(ADC("UoInWt")) = "Z", " (case when IrMainMet= 'Y' then IdCustmMtchWt- " + wRmWtOpt + " else " + wRmWtOpt + "* (Case when R1.RmCtg in ('X','M') Then 0 Else IdGldLs End)/100 end) ", wRmWtOpt + "* (Case when R1.RmCtg  in ('X','M') Then 0 Else IdGldLs End)/100") + "* R1.RmPurityZ/ (case when R2.RmPurityZ> 0 then R2.RmPurityZ else 1 end) ) as qInvPureLsWt, " + _
                 "sum(" + IIF(UCase(ADC("UoInWt")) = "Z", " (case when IrMainMet= 'Y' then IdCustmMtchWt- " + wRmWtOpt + " else " + wRmWtOpt + "* (Case when R1.RmCtg in ('X','M') Then 0 Else IdGldLs End)/100 end) ", wRmWtOpt + "* (Case when R1.RmCtg in ('X','M') Then 0 Else IdGldLs End)/100") + "* R1.RmPurityZ/ (case when R2.RmPurityZ> 0 then R2.RmPurityZ else 1 end)/ (case when (InHGldFiness> 0 and InHPlFiness> 0 and InHSlFiness> 0 and InHPdFiness> 0) then (case R1.RmCtg when 'G' then InHGldFiness when 'P' then InHPlFiness when 'S' then InHSlFiness when 'L' then InHPdFiness else 1 end) else 1 end)) as qInvFineLsWt, " + _
                 "sum(" + IIF(UCase(ADC("UoInWt")) = "Z", " (case when IrMainMet= 'Y' then IdCustmMtchWt else " + wRmWtOpt + "* (1 + (Case when R1.RmCtg in ('X','M') Then 0 Else IdGldLs End)/100) end) ", wRmWtOpt + "* (1 + (Case when R1.RmCtg in ('X','M') Then 0 Else IdGldLs End)/100)") + ") as qInvTotWt, " + _
                 "sum(" + IIF(UCase(ADC("UoInWt")) = "Z", " (case when IrMainMet= 'Y' then IdCustmMtchWt else " + wRmWtOpt + "* (1 + (Case when R1.RmCtg in ('X','M') Then 0 Else IdGldLs End)/100) end) ", wRmWtOpt + "* (1 + (Case when R1.RmCtg in ('X','M') Then 0 Else IdGldLs End)/100)") + "* (Case when R1.RmCtg='X' Then B.RmPurityZ Else R1.RmPurityZ End)/ (case when R2.RmPurityZ> 0 then R2.RmPurityZ else 1 end)) as qInvPureTotWt, " + _
                 "sum(" + IIF(UCase(ADC("UoInWt")) = "Z", " (case when IrMainMet= 'Y' then IdCustmMtchWt else " + wRmWtOpt + "* (1 + (Case when R1.RmCtg in ('X','M') Then 0 Else IdGldLs End)/100) end) ", wRmWtOpt + "* (1 + (Case when R1.RmCtg in ('X','M') Then 0 Else IdGldLs End)/100)") + "* R1.RmPurityZ/ (case when R2.RmPurityZ> 0 then R2.RmPurityZ else 1 end)/ (case when (InHGldFiness> 0 and InHPlFiness> 0 and InHSlFiness> 0 and InHPdFiness> 0) then (case R1.RmCtg when 'G' then InHGldFiness when 'P' then InHPlFiness when 'S' then InHSlFiness when 'L' then InHPdFiness else 1 end) else 1 end)) as qInvFineTotWt " + _
                 " from InvRm " + _
                 "      join RmMst R1 on R1.RmCd = IrRmCd " + IIF(gs_Partition = ctCurrPrtn, " And IrPrtKey=R1.RmPrtKey ", "") + _
                 "      Left Outer Join RmMst B On B.RmCd=R1.RmBaseCd  " + IIF(gs_Partition = ctCurrPrtn, " And IrPrtKey=B.RmPrtKey ", "") + _
                 "      Join RmMst R2 on R2.RmCd=IsNull(B.RmZCd,R1.RmZCd)  " + IIF(gs_Partition = ctCurrPrtn, " And R2.RmPrtKey=R1.RmPrtKey ", "") + _
                 "      join InvDsg on IdIdNo=IrIdIdNo and IdCoCd= IrCoCd and IdTc= IrTc and IdYy= IrYy and " + _
                 "              IdChr= IrChr and IdNo= IrNo and IdSr= IrSr " + IIF(gs_Partition = ctCurrPrtn, " And IdPrtKey=IrPrtKey ", "") + _
                 "      join InvHd Ih on InIdNo=IdInIdNo and InCoCd= IdCoCd and InTc = IdTc and InYy= IdYy and " + _
                 "              InChr= IdChr and InNo= IdNo " + IIF(gs_Partition = ctCurrPrtn, " And InPrtKey=IdPrtKey ", "") + " where R1.RmCtg in ('G', 'P', 'S', 'L','M'" + IIF(ADC("UOYN11") = "Y", ",'X')", ")") + IIF(gs_Partition = ctCurrPrtn, " And IrPrtKey='" + ctCurrPrtn + "' ", "") + _
                  wAndCnd + " Group By R1.RmCtg, " + IIF(wFlagLocalExp = "Y", IIF(wb_MLsRmCd = True, "IrRmCd ", "R1.RmKt "), IIF(wb_MLsRmCd = True, "IrRmCd", "R1.RmKt")) + ", IdGldLs "
                 
'                 wAndCnd " Group By R1.RmCtg, " + IIF(wFlagLocalExp = "Y", Ws_IdKt + " ", IIF(wb_MLsRmCd = True, "IrRmCd", "R1.RmKt")) + ", IdGldLs "
    '*** (Jen 2.14 Next Patch 1 05/06/07)
    '*** Jay 3.3(FG)
    Else
      wSubRepStr = "Select max(InTc) as InTc, max(InYy) as InYy, max(InChr) as InChr, " + _
                 "max(InNo) as InNo, 0 as IdGldLs, max(" + Ws_IdKt + ") as IdKt, max(OrRmCd) as IrRmCd, " + _
                 "R1.RmCtg as RmCtg, max(R1.RmKt) as RmKt, (Select PDesc from Param where PTyp= 'RMCTG' and PMCd= R1.RmCtg) as RmCtgDesc, " + _
                 "sum(" + wRmWtOpt + ") as qMetWt, " + _
                 "sum(" + wRmWtOpt + " * R1.RmPurityZ/ (case when R2.RmPurityZ> 0 then R2.RmPurityZ else 1 end)) as qMetPureWt, " + _
                 "sum(" + wRmWtOpt + " * R1.RmPurityZ/ (case when R2.RmPurityZ> 0 then R2.RmPurityZ else 1 end)/ (case when (InHGldFiness> 0 and InHPlFiness> 0 and InHSlFiness> 0 and InHPdFiness> 0) then (case R1.RmCtg when 'G' then InHGldFiness when 'P' then InHPlFiness when 'S' then InHSlFiness when 'L' then InHPdFiness else 1 end) else 1 end)) as qMetFineWt, " + _
                 "0 as qInvLsWt, " + _
                 "0 as qInvPureLsWt, " + _
                 "0 as qInvFineLsWt, " + _
                 "sum(" + wRmWtOpt + ") as qInvTotWt, " + _
                 "sum(" + wRmWtOpt + "* R1.RmPurityZ/ (case when R2.RmPurityZ> 0 then R2.RmPurityZ else 1 end)) as qInvPureTotWt, " + _
                 "sum(" + wRmWtOpt + "* R1.RmPurityZ/ (case when R2.RmPurityZ> 0 then R2.RmPurityZ else 1 end)/ (case when (InHGldFiness> 0 and InHPlFiness> 0 and InHSlFiness> 0 and InHPdFiness> 0) then (case R1.RmCtg when 'G' then InHGldFiness when 'P' then InHPlFiness when 'S' then InHSlFiness when 'L' then InHPdFiness else 1 end) else 1 end)) as qInvFineTotWt " + _
                 " from OrdRm join RmMst R1 on R1.RmCd = OrRmCd " + IIF(gs_Partition = ctCurrPrtn, " And OrPrtKey=R1.RmPrtKey ", "") + " join RmMst R2 on R2.RmCd= R1.RmZCd " + IIF(gs_Partition = ctCurrPrtn, " And R2.RmPrtKey=R1.RmPrtKey ", "") + _
                 "join OrdDsg on OdIdNo=OrOdIdNo and OdCoCd= OrCoCd and OdTc= OrTc and OdYy= OrYy and " + _
                 "OdChr= OrChr and OdNo= OrNo and OdSr= OrSr " + IIF(gs_Partition = ctCurrPrtn, " And OdPrtKey=OrPrtKey ", "") + _
                 "join InvHd Ih on InCoCd= OdCoCd and InTc = OdTc and InYy= OdYy and " + _
                 "InChr= OdChr and InNo= OdNo " + IIF(gs_Partition = ctCurrPrtn, " And InPrtKey=OdPrtKey ", "") + " where R1.RmCtg in ('G', 'P', 'S','L','M'" + IIF(ADC("UOYN11") = "Y", ",'X')", ")") + IIF(gs_Partition = ctCurrPrtn, " And OrPrtKey='" + ctCurrPrtn + "' ", "") + _
                 wAndCnd + " Group By R1.RmCtg, " + IIF(wFlagLocalExp = "Y", Ws_IdKt + " ", IIF(wb_MLsRmCd = True, "OrRmCd", "R1.RmKt"))
      'Debug.Print wSubRepStr
    End If
    '*** Jay 3.3(FG)
    
    
    
    '*** (Bef 2.14 Next Patch 1)
    ''*** (04/08/06)
    'wSubRepStr = "Select max(InTc) as InTc, max(InYy) as InYy, max(InChr) as InChr, " + _
    '             "max(InNo) as InNo, IdGldLs, max(IdKt) as IdKt, max(IrRmCd) as IrRmCd, " + _
    '             "RmCtg as RmCtg, max(RmKt) as RmKt, (Select PDesc from Param where PTyp= 'RMCTG' and PMCd= RmCtg) as RmCtgDesc, " + _
    '             "sum(" + wRmWtOpt + ") as qMetWt, " + _
    '             "sum(" + wRmWtOpt + " * RmPurityZ) as qMetPureWt, " + _
    '             "sum(" + wRmWtOpt + " * RmPurityZ/(case when (InHGldFiness> 0 and InHPlFiness> 0 and InHSlFiness> 0) then (case RmCtg when 'G' then InHGldFiness when 'P' then InHPlFiness when 'S' then InHSlFiness else 1 end) else 1 end)) as qMetFineWt, " + _
    '             "sum(" + IIF(UCase(adc("UoInWt")) = "Z", " (case when IrMainMet= 'Y' then IdCustmMtchWt- " + wRmWtOpt + " else " + wRmWtOpt + "* IdGldLs/100 end) ", wRmWtOpt + "* IdGldLs/100") + ") as qInvLsWt, " + _
    '             "sum(" + IIF(UCase(adc("UoInWt")) = "Z", " (case when IrMainMet= 'Y' then IdCustmMtchWt- " + wRmWtOpt + " else " + wRmWtOpt + "* IdGldLs/100 end) ", wRmWtOpt + "* IdGldLs/100") + "* RmPurityZ ) as qInvPureLsWt, " + _
    '             "sum(" + IIF(UCase(adc("UoInWt")) = "Z", " (case when IrMainMet= 'Y' then IdCustmMtchWt- " + wRmWtOpt + " else " + wRmWtOpt + "* IdGldLs/100 end) ", wRmWtOpt + "* IdGldLs/100") + "* RmPurityZ/ (case when (InHGldFiness> 0 and InHPlFiness> 0 and InHSlFiness> 0) then (case RmCtg when 'G' then InHGldFiness when 'P' then InHPlFiness when 'S' then InHSlFiness else 1 end) else 1 end)) as qInvFineLsWt, " + _
    '             "sum(" + IIF(UCase(adc("UoInWt")) = "Z", " (case when IrMainMet= 'Y' then IdCustmMtchWt else " + wRmWtOpt + "* (1 + IdGldLs/100) end) ", wRmWtOpt + "* (1 + IdGldLs/100)") + ") as qInvTotWt, " + _
    '             "sum(" + IIF(UCase(adc("UoInWt")) = "Z", " (case when IrMainMet= 'Y' then IdCustmMtchWt else " + wRmWtOpt + "* (1 + IdGldLs/100) end) ", wRmWtOpt + "* (1 + IdGldLs/100)") + "* RmPurityZ) as qInvPureTotWt, " + _
    '             "sum(" + IIF(UCase(adc("UoInWt")) = "Z", " (case when IrMainMet= 'Y' then IdCustmMtchWt else " + wRmWtOpt + "* (1 + IdGldLs/100) end) ", wRmWtOpt + "* (1 + IdGldLs/100)") + "* RmPurityZ/(case when (InHGldFiness> 0 and InHPlFiness> 0 and InHSlFiness> 0) then (case RmCtg when 'G' then InHGldFiness when 'P' then InHPlFiness when 'S' then InHSlFiness else 1 end) else 1 end)) as qInvFineTotWt " + _
    '             " from InvRm join RmMst on RmCd = IrRmCd " + _
    '             "join InvDsg on IdCoCd= IrCoCd and IdTc= IrTc and IdYy= IrYy and " + _
    '             "IdChr= IrChr and IdNo= IrNo and IdSr= IrSr " + _
    '             "join InvHd Ih on InCoCd= IdCoCd and InTc = IdTc and InYy= IdYy and " + _
    '             "InChr= IdChr and InNo= IdNo where RmCtg in ('G', 'P', 'S') " + _
    '             wAndCnd + " Group By RmCtg, " + IIF(wFlagLocalExp = "Y", "IdKt ", IIF(wb_MLsRmCd = True, "IrRmCd", "RmKt")) + ", IdGldLs "
    ''*** (04/08/06)
    '*** (Bef 2.14 Next Patch 1)
                 
    '*** (Bef 07/11/06) ***
    '" Group By RmCtg, " + IIF(wFlagLocalExp = "Y", "IdKt ", "IrRmCd ") + ", IdGldLs "
    '*** (Bef 07/11/06) ***
                 
    Call moCn.RepRes(mRep.Subreport1.OpenSubreport, wSubRepStr + ctMaxDopOpt)  '****** Sachin 3.01 ctMaxDopOpt added
    '*** (Jen 2.13)
    
    'sv.24 calling procedure for Stncls summary
    'it giving procesure too large error, so query for Stncls summary done in separately
    Call GetStnClsSumm(wAndCnd)
          
    If wFlagLocalExp = "Y" Then
      mRep.FldInvTotRs.Suppress = True
      
      '*** (Bef 2.13) Jen
      'Set wRsFiness = moCn.OpenRes("Select RmCtg,(case when RmCtg='G' then  max(InHGldFiness) When RmCtg='P' then max(InHPlFiness) " + _
      '                              "else max(InHSlFiness) end ) as qMetFiness , " + _
      '                              "max(" + IIF(UCase(adc("UoInWt")) = "I", "InLmgSal", "InLmgCst") + ") as qLmgVal," + _
      '                              "max(" + IIF(UCase(adc("UoInWt")) = "I", "InLmpSal", "InLmpCst") + ") as qLmpVal, " + _
      '                              "max(" + IIF(UCase(adc("UoInWt")) = "I", "InLmsSal", "InLmsCst") + ") as qLmsVal, " + _
      '                              "Case When RmCtg ='G' Then 'Gold In '+Str((case when RmCtg='G' then  max(InHGldFiness)  when RmCtg='P' then max(InHPlFiness) else max(InHSlFiness) end ),6,3)+ ' Fineness ' " + _
      '                              "     When RmCtg ='P' Then 'Platinum In '+ Str((case when RmCtg='G' then  max(InHGldFiness) when RmCtg='P' then max(InHPlFiness) else max(InHSlFiness) end ),6,3)+ ' Fineness ' " + _
      '                              "     When RmCtg ='S' Then 'Silver In '+ Str((case when RmCtg='G' then  max(InHGldFiness) when RmCtg='P' then max(InHPlFiness) else max(InHSlFiness) end ),6,3)+ ' Fineness ' " + _
      '                              "   End  as qFineStr " + _
      '                              "From InvHd Join InvRm on InCoCd= IrCoCd and InTc = IrTc and InYy= IrYy " + _
      '                              "and InChr= IrChr and  InNo= IrNo " + _
      '                              "join RmMst on RmCd = IrRmCd  where RmCtg in ('G', 'P', 'S') " + _
      '                              "and InCoCd= '" + adc("UoCoCdFr") + "'  and InTc= '" + adc("UoInTcFr") + "'" + _
      '                              "and InYy= '" + adc("UoInYyFr") + "' and InChr= '" + adc("UoInChrFr") + "'" + _
      '                              "and InNo=" + CStr(adc("UoInNoFr")) + " Group By RmCtg ")
      '*** (Bef 2.13) Jen
      
      '*** Jay 3.3(FG)
      ' ***** Manali Trading Module
      If UCase(ADC.MenuCd) <> UCase("RepFgInv") And UCase(ADC.MenuCd) <> UCase("RepJtInv") Then
      '*** Jay 3.3(FG)
        '*** (Jen 2.13)
        '****** Sachin 3.02 - Id fields In Joins
        ' **** Manali 3.03 - 25/06/08 - Pd Mod
        Set wRsFiness = moCn.OpenRes("Select RmCtg,(case when RmCtg= 'G' then max(InHGldFiness) " + _
                                   "When RmCtg= 'P' then max(InHPlFiness) When RmCtg= 'S' then max(InHSlFiness) else max(InHPdFiness) end) as qMetFiness, " + _
                                   "max(InLmgSal) as qLmgVal, max(InLmpSal) as qLmpVal, max(InLmsSal) as qLmsVal,  max(InLmlSal) as qLmlVal, " + _
                                   "Case When RmCtg ='G' Then 'Gold In '+Str((case when RmCtg='G' then  max(InHGldFiness)  when RmCtg='P' then max(InHPlFiness) when RmCtg='S' then max(InHSlFiness) else max(InHPdFiness) end ), 6, 3)+ ' Fineness ' " + _
                                   "     When RmCtg ='P' Then 'Platinum In '+ Str((case when RmCtg='G' then  max(InHGldFiness) when RmCtg='P' then max(InHPlFiness) when RmCtg='S' then max(InHSlFiness) else max(InHPdFiness) end ), 6, 3)+ ' Fineness ' " + _
                                   "     When RmCtg ='S' Then 'Silver In '+ Str((case when RmCtg='G' then  max(InHGldFiness) when RmCtg='P' then max(InHPlFiness) when RmCtg='S' then max(InHSlFiness) else max(InHPdFiness) end ), 6, 3)+ ' Fineness ' " + _
                                   "     When RmCtg ='L' Then 'Palladium In '+ Str((case when RmCtg='G' then  max(InHGldFiness) when RmCtg='P' then max(InHPlFiness) when RmCtg='S' then max(InHSlFiness) else max(InHPdFiness) end ), 6, 3)+ ' Fineness ' " + _
                                   "   End  as qFineStr " + _
                                   " From InvHd Ih Join InvRm on Ih.InCoCd= IrCoCd and Ih.InTc= IrTc and Ih.InYy= IrYy " + IIF(gs_Partition = ctCurrPrtn, " And Ih.InPrtKey=IrPrtKey ", "") + _
                                   " and Ih.InChr= IrChr and Ih.InNo= IrNo " + _
                                   "join RmMst on RmCd = IrRmCd  " + IIF(gs_Partition = ctCurrPrtn, " And RmPrtKey=IrPrtKey", "") + _
                                   " where " + IIF(gs_Partition = ctCurrPrtn, " InPrtKey='" + ctCurrPrtn + "' And ", "") + " RmCtg in ('G', 'P', 'S', 'L') " + wAndCnd + " Group By RmCtg ")
      '*** (Jen 2.13)
      '*** Jay 3.3(FG)
      Else
        Set wRsFiness = moCn.OpenRes("Select RmCtg,(case when RmCtg= 'G' then max(InHGldFiness) " + _
                                   "When RmCtg= 'P' then max(InHPlFiness) When RmCtg= 'S' then max(InHSlFiness) else max(InHPdFiness) end) as qMetFiness, " + _
                                   "max(InLmgSal) as qLmgVal, max(InLmpSal) as qLmpVal, max(InLmsSal) as qLmsVal,  max(InLmlSal) as qLmlVal, " + _
                                   "Case When RmCtg ='G' Then 'Gold In '+Str((case when RmCtg='G' then  max(InHGldFiness)  when RmCtg='P' then max(InHPlFiness) when RmCtg='S' then max(InHSlFiness) else max(InHPdFiness) end ), 6, 3)+ ' Fineness ' " + _
                                   "     When RmCtg ='P' Then 'Platinum In '+ Str((case when RmCtg='G' then  max(InHGldFiness) when RmCtg='P' then max(InHPlFiness) when RmCtg='S' then max(InHSlFiness) else max(InHPdFiness) end ), 6, 3)+ ' Fineness ' " + _
                                   "     When RmCtg ='S' Then 'Silver In '+ Str((case when RmCtg='G' then  max(InHGldFiness) when RmCtg='P' then max(InHPlFiness) when RmCtg='S' then max(InHSlFiness) else max(InHPdFiness) end ), 6, 3)+ ' Fineness ' " + _
                                   "     When RmCtg ='L' Then 'Palladium In '+ Str((case when RmCtg='G' then  max(InHGldFiness) when RmCtg='P' then max(InHPlFiness) when RmCtg='S' then max(InHSlFiness) else max(InHPdFiness) end ), 6, 3)+ ' Fineness ' " + _
                                   "   End  as qFineStr " + _
                                   " From InvHd Ih Join OrdRm on Ih.InCoCd= OrCoCd and Ih.InTc= OrTc and Ih.InYy= OrYy " + IIF(gs_Partition = ctCurrPrtn, " And Ih.InPrtKey=OrPrtKey ", "") + _
                                   " and Ih.InChr= OrChr and Ih.InNo= OrNo " + _
                                   "join RmMst on RmCd = OrRmCd  " + IIF(gs_Partition = ctCurrPrtn, " And RmPrtKey=OrPrtKey", "") + _
                                   " where " + IIF(gs_Partition = ctCurrPrtn, " InPrtKey='" + ctCurrPrtn + "' And ", "") + " RmCtg in ('G', 'P', 'S','L','M') " + wAndCnd + " Group By RmCtg ")
      End If
      '*** Jay 3.3(FG)
      
      If wRsFiness.RecCount > 0 Then _
        ws_FinessStr = "Statement Showing "
      RsCnt = 1
      
      '*** (Bef 2.13)
      'wd_GldOunce = moCn.GetFldVal("Select PNum From Param Where PTyp= 'OUNCE' and PMCd= 'G' and PSCd= ''")
      'wd_PlOunce = moCn.GetFldVal("Select PNum From Param Where PTyp= 'OUNCE' and PMCd= 'P' and PSCd= ''")
      '' ****** Sachin 2.12 - 06/12/05 - SJM
      'wd_SlOunce = moCn.GetFldVal("Select PNum From Param Where PTyp= 'OUNCE' and PMCd= 'S' and PSCd= ''")
      '' ****** Sachin 2.12 - 06/12/05 - SJM
      '*** (Bef 2.13)
      
      With wRsFiness
      Do While Not (.BOF Or .EOF)
        ws_FinessStr = ws_FinessStr + !qFineStr
        ws_FinessStr = ws_FinessStr + IIF(.RecCount > 1 And RsCnt < .RecCount, " And ", "")
        RsCnt = RsCnt + 1
        wd_LmgGms = Round(!qLmgVal / IIF(wd_GldOunce > 0, wd_GldOunce, 1), 4)
        wd_LmpGms = Round(!qLmpVal / IIF(wd_PlOunce > 0, wd_PlOunce, 1), 4)
        ' ****** Sachin 2.12 - 06/12/05 - SJM
        wd_LmsGms = Round(!qLmsVal / IIF(wd_SlOunce > 0, wd_SlOunce, 1), 4)
        ' ****** Sachin 2.12 - 06/12/05 - SJM
        ' **** Manali 3.03 - 25/06/08 - Pd Mod
        wd_LmlGms = Round(!qLmlVal / IIF(wd_PdOunce > 0, wd_PdOunce, 1), 4)
        ' **** Manali 3.03 - 25/06/08 - Pd Mod
        ws_GldStr = IIF(!RmCtg = "G", ws_CurCd + " " + Str(!qLmgVal) + _
                          " For " + ws_GldOunceDsc + " " + ws_CurCd + " " + _
                          Str(wd_LmgGms) + " For " + Str(Round(!qMetFiness, 3)) + " Fineness  " + ws_CurCd + " " + _
                          Str(Round(wd_LmgGms / !qMetFiness, 4)) + " For Pure 1 Gram ", ws_GldStr)
         ws_PlStr = IIF(!RmCtg = "P", ws_CurCd + " " + Str(!qLmpVal) + _
                          " For " + ws_PlOunceDsc + " " + ws_CurCd + " " + _
                          Str(wd_LmpGms) + " For " + Str(Round(!qMetFiness, 3)) + " Fineness  " + ws_CurCd + " " + _
                          Str(Round(wd_LmpGms / !qMetFiness, 4)) + " For Pure 1 Gram ", ws_PlStr)
        ' ****** Sachin 2.12 - 06/12/05 - SJM
        ws_SlvStr = IIF(!RmCtg = "S", ws_CurCd + " " + Str(!qLmsVal) + _
                          " For " + ws_SlOunceDsc + " " + ws_CurCd + " " + _
                          Str(wd_LmsGms) + " For " + Str(Round(!qMetFiness, 3)) + " Fineness  " + ws_CurCd + " " + _
                          Str(Round(wd_LmsGms / !qMetFiness, 4)) + " For Pure 1 Gram ", ws_SlvStr)
        ' ****** Sachin 2.12 - 06/12/05 - SJM
        ' **** Manali 3.03 - 25/06/08 - Pd Mod
        ws_PdStr = IIF(!RmCtg = "L", ws_CurCd + " " + Str(!qLmlVal) + _
                          " For " + ws_PdOunceDsc + " " + ws_CurCd + " " + _
                          Str(wd_LmlGms) + " For " + Str(Round(!qMetFiness, 3)) + " Fineness  " + ws_CurCd + " " + _
                          Str(Round(wd_LmlGms / !qMetFiness, 4)) + " For Pure 1 Gram ", ws_PdStr)
        ' **** Manali 3.03 - 25/06/08 - Pd Mod
        wRsFiness.MoveNext
      Loop
      End With
      ' **** Manali 3.03 - 25/06/08 - Pd Mod
      mRep.TxtOunceRt.SetText ws_GldStr + vbCr + ws_PlStr + vbCr + ws_SlvStr + vbCr + ws_PdStr
            
      '*** (Bef 2.13) Jen
      '*** (24/02/06) For Kama (Jen 2.12)
      'wSubRepStr = "Select max(InTc) as InTc, max(InYy) as InYy, max(InChr) as InChr, " + _
      '             "max(InNo) as InNo, IdGldLs, IdKt, " + _
      '             "RmCtg as RmCtg, (Select PDesc from Param where PTyp= 'RMCTG' and PMCd= RmCtg) as RmCtgDesc, " + _
      '             "sum(" + wRmWtOpt + ") as qMetWt, " + _
      '             "sum(" + wRmWtOpt + " * RmPurityZ) as qMetPureWt, " + _
      '             "sum(" + wRmWtOpt + " * RmPurityZ/(case when (InHGldFiness> 0 and InHPlFiness> 0 and InHSlFiness> 0) then (case RmCtg when 'G' then InHGldFiness when 'P' then InHPlFiness when 'S' then InHSlFiness else 1 end) else 1 end)) as qMetFineWt, " + _
      '             "sum(" + wRmWtOpt + " * IdGldLs/100) as qInvLsWt, " + _
      '             "sum(" + wRmWtOpt + " * RmPurityZ * IdGldLs/100) as qInvPureLsWt, " + _
      '             "sum(" + wRmWtOpt + " * RmPurityZ * IdGldLs/100/(case when (InHGldFiness> 0 and InHPlFiness> 0 and InHSlFiness> 0) then (case RmCtg when 'G' then InHGldFiness when 'P' then InHPlFiness when 'S' then InHSlFiness else 1 end) else 1 end)) as qInvFineLsWt, " + _
      '             "sum(" + wRmWtOpt + " * (1 + IdGldLs/100)) as qInvTotWt, " + _
      '             "sum(" + wRmWtOpt + " * RmPurityZ * (1 + IdGldLs/100)) as qInvPureTotWt, " + _
      '             "sum(" + wRmWtOpt + " * RmPurityZ * (1 + IdGldLs/100)/(case when (InHGldFiness> 0 and InHPlFiness> 0 and InHSlFiness> 0) then (case RmCtg when 'G' then InHGldFiness when 'P' then InHPlFiness when 'S' then InHSlFiness else 1 end) else 1 end)) as qInvFineTotWt " + _
      '             " from InvRm join RmMst on RmCd = IrRmCd " + _
      '             "join InvDsg on IdCoCd= IrCoCd and IdTc= IrTc and IdYy= IrYy and " + _
      '             "IdChr= IrChr and IdNo= IrNo and IdSr= IrSr " + _
      '             "join InvHd on InCoCd= IdCoCd and InTc = IdTc and InYy= IdYy and " + _
      '             "InChr= IdChr and InNo= IdNo where RmCtg in ('G', 'P', 'S') " + _
      '             wCnd + " Group By RmCtg, IdKt, IdGldLs "
      'Call moCn.RepRes(mRep.Subreport1.OpenSubreport, wSubRepStr)
      '*** (24/02/06) For Kama (Jen 2.12)
      '*** (Bef 2.13) Jen
      
      '****Geeta***Emr208
      ' Zubin 212 (Sales Ctg and qDmSalCtgDesc added)
      '*** (Bef 2.13)
      'wSubRepStr = "Select '1D' as qTag, DmCtg, max(DmSalCtg) as DmSalCtg, (Select PDesc From Param where PTyp= 'DMCTG'" + _
      '            " and PMCd= DmCtg and PSCd= '') as DmCtgDesc, " + _
      '            " (Select PDesc from Param where PTyp= 'SALCTG' and PMCd= DmCtg and " + _
      '            " PSCd= max(DmSalCtg)) as DmSalCtgDesc, max(DmUom) as DmUom,Max(IdKt) as IdKt," + _
      '            "Max(IdVaCtg) as IdVaCtg, Max(IdGldLs) as IdGldLs, Max(OdDmCol) as OdDmCol," + qMainMet + ", sum(IdQty) as qIdQty " + _
      '            "From InvDsg Join DsgMst  " + _
      '            "On DmTcTyp= 'DM' and IdDmCd= DmCd and DmSz= '' " + _
      '            "join OrdDsg on OdCoCd= IdCoCd and OdTc= IdExpOdTc and OdYy= IdExpOdYy " + _
      '            "and OdChr= IdExpOdChr and OdNo= IdExpOdNo and OdSr= IdExpOdSr " + wOrJoin + _
      '            "where IdCoCd= '" + adc("UoCoCdFr") + "' and IdTc= '" + adc("UoInTcFr") + "' and " + _
      '            "IdYy= '" + adc("UoInYyFr") + "' and IdChr= '" + adc("UoInChrFr") + "' and " + _
      '            "IdNo= " + CStr(adc("UoInNoFr")) + "  group by " + wSubRepGrpFormula + "DmCtg, DmUom"
      '*** (Bef 2.13)
      
      '*** Jay 3.3(FG)
      ' ***** Manali Trading Module
      If UCase(ADC.MenuCd) <> UCase("RepFgInv") And UCase(ADC.MenuCd) <> UCase("RepJtInv") Then
      '*** Jay 3.3(FG)
        '*** (Jen 2.13)
        '****** Sachin 3.02 - Id fields in Joins
        wSubRepStr = "Select '1D' as qTag,Max(OdOmCmCd) as OmCmCd, DmCtg, max(DmSalCtg) as DmSalCtg, (Select PDesc From Param where PTyp= 'DMCTG'" + _
                  " and PMCd= DmCtg and PSCd= '') as DmCtgDesc, " + _
                  " (Select PDesc from Param where PTyp= 'SALCTG' and PMCd= DmCtg and " + _
                  " PSCd= max(DmSalCtg)) as DmSalCtgDesc, max(DmUom) as DmUom,/*  Max(" + Ws_IdKt + ") */ Max(IdKt) as IdKt," + _
                  "Max(" + ws_IdVaCtg + ") as IdVaCtg, Max(IdGldLs) as IdGldLs, Max(OdDmCol) as OdDmCol," + qMainMet + ", sum(" + ws_IdQty + ") as qIdQty " + _
                  " From InvDsg Join DsgMst  " + _
                  " On DmIdNo=IdDmIdNo and DmTcTyp= 'DM' and IdDmCd= DmCd and DmSz= '' " + IIF(gs_Partition = ctCurrPrtn, " And IdPrtKey=DmPrtKey", "") + _
                  " join InvHd Ih on Ih.InIdNo=IdInIdNo and Ih.InCoCd= IdCoCd and Ih.InTc= IdTc and Ih.InYy= IdYy and Ih.InChr= IdChr and Ih.InNo= IdNo " + IIF(gs_Partition = ctCurrPrtn, " And Ih.InPrtKey=IdPrtKey ", "") + _
                  " join OrdDsg on OdCoCd= IdCoCd and OdTc= IdExpOdTc and OdYy= IdExpOdYy " + _
                  " and OdChr= IdExpOdChr and OdNo= IdExpOdNo and OdSr= IdExpOdSr " + IIF(gs_Partition = ctCurrPrtn, " And OdPrtKey=IdPrtKey", "") + wOrJoin + _
                  " " + wCnd + IIF(gs_Partition = ctCurrPrtn, " And IdPrtKey='" + ctCurrPrtn + "' ", "") + " group by " + wSubRepGrpFormula + "DmCtg, DmUom "
        '*** (Jen 2.13)
      '*** Jay 3.3(FG)
      Else
          wSubRepStr = "Select '1D' as qTag,Max(OdOmCmCd) as OmCmCd,DmCtg, max(DmSalCtg) as DmSalCtg, (Select PDesc From Param where PTyp= 'DMCTG'" + _
                  " and PMCd= DmCtg and PSCd= '') as DmCtgDesc, " + _
                  " (Select PDesc from Param where PTyp= 'SALCTG' and PMCd= DmCtg and " + _
                  " PSCd= max(DmSalCtg)) as DmSalCtgDesc, max(DmUom) as DmUom, /* Max(" + Ws_IdKt + ") */ Max(IdKt) as IdKt," + _
                  "Max(" + ws_IdVaCtg + ") as IdVaCtg, 0 as IdGldLs, Max(OdDmCol) as OdDmCol," + qMainMet + ", sum(" + ws_IdQty + ") as qIdQty " + _
                  " From OrdDsg Join DsgMst  " + _
                  " On DmIdNo=OdDmIdNo and DmTcTyp= 'DM' and OdDmCd= DmCd and DmSz= '' " + IIF(gs_Partition = ctCurrPrtn, " And OdPrtKey=DmPrtKey", "") + _
                  " join InvHd Ih on Ih.InCoCd= OdCoCd and Ih.InTc= OdTc and Ih.InYy= OdYy and Ih.InChr= OdChr and Ih.InNo= OdNo " + IIF(gs_Partition = ctCurrPrtn, " And Ih.InPrtKey=OdPrtKey ", "") + _
                  " " + wCnd + IIF(gs_Partition = ctCurrPrtn, " And OdPrtKey='" + ctCurrPrtn + "' ", "") + " group by " + wSubRepGrpFormula + "DmCtg, DmUom "
      End If
      '*** Jay 3.3(FG)
                                
      'Call moCn.RepRes(mRep.Subreport2.OpenSubreport, wSubRepStr) '(Bef 2.13) Jen
      '*** (Jen 2.13)
      Call moCn.RepRes(mRep.Subreport3.OpenSubreport, wSubRepStr + ctMaxDopOpt) '****** Sachin 3.01 ctMaxDopOpt added
      SetSubFormula mRep.Subreport3, "wGrp", wGrpFormula
      '*** (Jen 2.13)
      '**********Geeta*******Emr208***
      
      '*** (Jen 2.13)
      'Rep.Subreport1_TxtRmCdKt.SetText "KT"         '*** (Bef 07/11/06)
      Rep.Subreport1_TxtRmCdKt.SetText "Dsg KT"         '*** (07/11/06)
      Rep.Subreport1_DetInv.Suppress = True
      Rep.Subreport1_SecFGrp1Inv.Suppress = True
      '*** (Jen 2.13)
      
      ' ***** Manali 3.6.0 - 09/09/09 - Cost Details in DTA Invoice
      ' ***** Manali 3.6.1 - InFob changed to InFob+InDsc
      If UCase(ADC.MenuCd) = UCase("RepInv") And ADC("UoYN8") = "Y" Then
          Dim wRsCstDet As MDORowSet, ws_CstDet As String
          ' ***** Manali 3.8.0
          ' ***** Manali 3.8.1 - Self Payable Dsc and Deduct included in Other Cost
          ws_CstDet = "Select Sum(Case When TzRmCtg In ('G', 'P', 'S', 'L', 'M') Then TzPurAmt Else 0.0 End) as qMetCst, " + _
                        "Sum(Case When TzRmCtg In ('D', 'C') Then TzPurAmt Else 0.0 End) as qStudCst, " + _
                        "Sum(Case When TzRmCtg = 'X' Then TzPurAmt Else 0.0 End) as qAccCst, " + _
                        "Convert(Decimal(14,2), Max((Case When InHHideCommYN='Y' Then InComm Else 0.0 End)+(Case When InHSelfDscYN='Y' Then InDsc Else 0.0 End)+(Case When InHSelfDeductYN='Y' Then InDeduct  Else 0.0 End))) as qOthCst, " + _
                        "Max(InHHideCommYN) as InHHideCommYN, " + _
                        "Max(InFob) As InFob, Convert(Decimal(14,2), (Max(InFob) - (Sum(TzPurAmt)+Max((Case When InHHideCommYN='Y' Then InComm Else 0.0 End)+(Case When InHSelfDscYN='Y' Then InDsc Else 0.0 End)+(Case When InHSelfDeductYN='Y' Then InDeduct  Else 0.0 End))))) as qMkChg From TxndZ " + _
                        "Join InvHd Ih On InIdNo=TzInIdNo and InPrtKey=TzPrtKey and InCoCd=TzCoCd and InTc=TzTc and InYy=TzYy and InChr=TzChr and InNo=TzNo " + _
                        wCnd + IIF(gs_Partition = ctCurrPrtn, " And TzPrtKey='" + ctCurrPrtn + "' ", "") + _
                        " and TzBM='M' Group By TzCoCd, TzTc, TzYy, TzChr, TzNo "
            
            ' ***** bef 3.8.0
''          ws_CstDet = "Select Sum(Case When TzRmCtg In ('G', 'P', 'S', 'L', 'M') Then TzPurAmt Else 0.0 End) as qMetCst, " + _
''                        "Sum(Case When TzRmCtg In ('D', 'C') Then TzPurAmt Else 0.0 End) as qStudCst, " + _
''                        "Sum(Case When TzRmCtg = 'X' Then TzPurAmt Else 0.0 End) as qAccCst, " + _
''                        "Max(InFob+InDsc) As InFob, (Max(InFob+InDsc) - Sum(TzPurAmt)) as qMkChg From TxndZ " + _
''                        "Join InvHd Ih On InIdNo=TzInIdNo and InPrtKey=TzPrtKey and InCoCd=TzCoCd and InTc=TzTc and InYy=TzYy and InChr=TzChr and InNo=TzNo " + _
''                        wCnd + IIF(gs_Partition = ctCurrPrtn, " And TzPrtKey='" + ctCurrPrtn + "' ", "") + _
''                        " and TzBM='M' Group By TzCoCd, TzTc, TzYy, TzChr, TzNo "
          
          Set wRsCstDet = moCn.OpenRes(ws_CstDet)
          Dim wMkPcent As Single
                        
'          Rep.TxtMetCst.SetText "Total Cost Of Metal " + ws_CurCd
'          Rep.TxtStudCst.SetText "Total Cost Of Studding " + ws_CurCd
'          Rep.TxtAccCst.SetText "Total Cost Of Finding " + ws_CurCd
'          Rep.TxtOthCst.SetText "Other Cost " + ws_CurCd      ' ***** Manali 3.8.0
'          Rep.TxtMkChg.SetText "Making Charges " + ws_CurCd
          Do While Not (wRsCstDet.EOF Or wRsCstDet.BOF)
            Rep.FldMetCst.SetText IIF(wRsCstDet!qMetCst <> 0, Format(wRsCstDet!qMetCst, ".00"), "")
            Rep.FldStudCst.SetText IIF(wRsCstDet!qStudCst <> 0, Format(wRsCstDet!qStudCst, ".00"), "")
            Rep.FldAccCst.SetText IIF(wRsCstDet!qAccCst <> 0, Format(wRsCstDet!qAccCst, ".00"), "")
            Rep.FldOthCst.SetText Format(IIF(wRsCstDet!qOthCst <> 0, wRsCstDet!qOthCst, ""), ".00")    ' ***** Manali 3.8.0
            Rep.FldMkChg.SetText IIF(wRsCstDet!qMkChg <> 0, Format(wRsCstDet!qMkChg, ".00"), "")
            wMkPcent = Format(MWLib.Div(wRsCstDet!qMkChg, wRsCstDet!qMetCst + wRsCstDet!qStudCst + wRsCstDet!qAccCst) * 100, "000.00")  '+ wRsCstDet!qOthCst
            ' bef 3.8.0
            'wMkPcent = Format(MWLib.Div(wRsCstDet!qMkChg, wRsCstDet!InFob) * 100, "000.00")
            Rep.FldMkChgPc.SetText IIF(wMkPcent <> 0, wMkPcent, "")
            wRsCstDet.MoveNext
          Loop
          Rep.TxtMetCst.SetText IIF(Rep.FldMetCst.text = "", "", "Total Cost Of Metal " + ws_CurCd)
          Rep.TxtStudCst.SetText IIF(Rep.FldStudCst.text = "", "", "Total Cost Of Studding " + ws_CurCd)
          Rep.TxtAccCst.SetText IIF(Rep.FldAccCst.text = "", "", "Total Cost Of Finding " + ws_CurCd)
          Rep.TxtOthCst.SetText IIF(Rep.FldOthCst.text = "", "", "Other Cost " + ws_CurCd)      ' ***** Manali 3.8.0
          Rep.TxtMkChg.SetText IIF(Rep.FldMkChg.text = "", "", "Making Charges " + ws_CurCd)
          Rep.TxtMkChgPC.SetText IIF(Rep.FldMkChgPc.text = "", "", "Making Charges % ")     '+ ws_CurCd
          Rep.LineQty.Suppress = True
          Set wRsCstDet = Nothing
      ElseIf UCase(ADC.MenuCd) = UCase("RepInvLoose") Or UCase(ADC.MenuCd) = UCase("RepJtInv") Or UCase(ADC.MenuCd) = UCase("RepFgInv") Then
          Rep.SubRepSec.Suppress = True
      Else
          If ADC("UoYn2") = "N" Then
            Rep.SubRepSec.Suppress = True
          Else
            Rep.TxtMetCst.Suppress = True: Rep.TxtStudCst.Suppress = True
            Rep.TxtAccCst.Suppress = True: Rep.TxtMkChg.Suppress = True
            Rep.TxtMkChgPC.Suppress = True
            Rep.TxtOthCst.Suppress = True   ' ***** Manali 3.8.0
            Rep.LineQtyUp.Suppress = True: Rep.LineQtyDn.Suppress = True: Rep.LineMkChg.Suppress = True
          End If
      End If
      
      ' ***** Manali 3.6.0 - 09/09/09 - Cost Details in DTA Invoice
    Else
      '*** (Jen 2.13)
      'mRep.TxtInHGjepcNo.Suppress = True: mRep.FldInHGjepcNo.Suppress = True
      'mRep.TxtInHItcNo.Suppress = True: mRep.FldInHItcNo.Suppress = True
      'mRep.FldInMetalUtil.Suppress = True
      mRep.SecDescSubRep.Suppress = True: mRep.SecOunceRt.Suppress = True
      
      '****** Sachin 2.14.0 Next ******
      If wb_MLsRmCd = True Then
         Rep.Subreport1_fldRmCd.Suppress = False
         Rep.Subreport1_fldRmKt.Suppress = True
      Else
         Rep.Subreport1_fldRmCd.Suppress = True
         Rep.Subreport1_fldRmKt.Suppress = False
      End If
      '****** Sachin 2.14.0 Next ******
         
      'Rep.Subreport1_TxtRmCdKt.SetText "RmCode"    '*** (Bef 07/11/06)
      Rep.Subreport1_TxtRmCdKt.SetText IIF(wb_MLsRmCd = True, "RM CD", "RM KT")        '*** (07/11/06)
      'Rep.Subreport1_SecHGrp1.Suppress = True      'Commented in 3.11.0
      Rep.Subreport1_DetLocalInv.Suppress = True
      Rep.Subreport1_SecFGrp1LocalInv.Suppress = True
      '*** (Jen 2.13)
      
      '*** (Bef 2.13) Jen
      'wSubRepStr = "Select max(InTc) as InTc, max(InYy) as InYy, max(InChr) as InChr, " + _
      '         "max(InNo) as InNo, IdGldLs, " + _
      '         "max(RmCtg) as RmCtg, IrRmCd, sum(" + wRmWtOpt + ") as qMetWt, " + _
      '         "sum(" + wRmWtOpt + " * RmPurityZ) as qMetPureWt, " + _
      '         "sum(" + wRmWtOpt + " * IdGldLs/100) as qInvLsWt, " + _
      '         "sum(" + wRmWtOpt + " * RmPurityZ * IdGldLs/100) as qInvPureLsWt, " + _
      '         "sum(" + wRmWtOpt + " * (1 + IdGldLs/100)) as qInvTotWt, " + _
      '         "sum(" + wRmWtOpt + " * RmPurityZ * (1 + IdGldLs/100)) as qInvPureTotWt " + _
      '         " from InvRm join RmMst on RmCd = IrRmCd " + _
      '         "join InvDsg on IdCoCd= IrCoCd and IdTc= IrTc and IdYy= IrYy and " + _
      '         "IdChr= IrChr and IdNo= IrNo and IdSr= IrSr " + _
      '         "join InvHd on InCoCd= IdCoCd and InTc = IdTc and InYy= IdYy and " + _
      '         "InChr= IdChr and InNo= IdNo where RmCtg in ('G', 'P', 'S') " + _
      '         wCnd + " Group By IrRmCd, IdGldLs "
      'Call moCn.RepRes(mRep.Subreport1.OpenSubreport, wSubRepStr)
      '*** (Bef 2.13) Jen
      
      ' ***** Manali 3.6.0 - 09/09/09 - Cost Details in DTA Invoice
      If UCase(ADC.MenuCd) = UCase("RepInv") Then
        If ADC("UoYn2") = "N" Then
          Rep.SubRepSec.Suppress = True
        Else
          Rep.TxtMetCst.Suppress = True: Rep.TxtStudCst.Suppress = True
          Rep.TxtAccCst.Suppress = True: Rep.TxtMkChg.Suppress = True
          Rep.TxtMkChgPC.Suppress = True
          Rep.TxtOthCst.Suppress = True   ' ***** Manali 3.8.0
          Rep.LineQtyUp.Suppress = True: Rep.LineQtyDn.Suppress = True: Rep.LineMkChg.Suppress = True
        End If
      ElseIf UCase(ADC.MenuCd) = UCase("RepInvLoose") Or UCase(ADC.MenuCd) = UCase("RepJtInv") Or UCase(ADC.MenuCd) = UCase("RepFgInv") Then
        Rep.SubRepSec.Suppress = True
      End If
      ' ***** Manali 3.6.0 - 09/09/09 - Cost Details in DTA Invoice
    End If
    
    ' Zubin 212
    If UCase(ADC("UoYN5")) = "N" Then mRep.SecCustRem.Suppress = True
    ' Zubin 212
    
    '3.11.0
    '************Bhavna
    
    
'  If UCase(adc.MenuCd) = UCase("RepInv") Or UCase(adc.MenuCd) = UCase("RepShpBill") Then
'    If UCase(adc("UoYN10")) = "Y" Then
'        mRep.TxtBinNo.Suppress = True
'        mRep.FldBinNo.Suppress = True
'        mRep.FldCERegnNo.Suppress = False
'
'    Else
'        mRep.FldBinNo.Suppress = False
'        mRep.FldCERegnNo.Suppress = True
'        mRep.TxtCERegnNo.Suppress = True
'
'    End If
'  Else
'    If UCase(adc("UoYN10")) = "Y" Then mRep.TxtBinNo.SetText "CE Regn. No."
'  End If
        
    '*** (Jen 2.13)
    If mb_MultiInv = False Then
      If Trim(UCase(moCn.GetFldVal("Select min(InShowGldLs) from InvHd Ih " + wCnd))) = "N" Then
        Rep.Subreport1_TxtLsPrc.Suppress = True
        Rep.Subreport1_wLsPrcStr1.Suppress = True
        Rep.Subreport1_wLsPrcStr2.Suppress = True
      End If
    'Else
    '  If Trim(UCase(moCn.GetFldVal("Select min(InShowGldLs) from InvHd where InCoCd= '" + adc("UoCoCdFr") + "' and InExpNo= '" + ws_MinExpNo + "' "))) = "N" Then
    '    Rep.Subreport1_TxtLsPrc.Suppress = True
    '    Rep.Subreport1_wLsPrcStr1.Suppress = True
    '    Rep.Subreport1_wLsPrcStr2.Suppress = True
    '  End If
    End If
    
    Rep.TxtAmtUSD.SetText "Amount" + Chr(13) + ws_CurCd
    Rep.TxtFOB.SetText "FOB " + ws_CurCd
    Rep.LblComm.SetText "COM " + ws_CurCd
    Rep.LblFri.SetText "FRI " + ws_CurCd
    Rep.TxtIns.SetText "INS " + ws_CurCd
    Rep.TxtCur1.SetText ws_CurCd
    Rep.TxtCur2.SetText ws_CurCd
    Rep.TxtCur3.SetText ws_CurCd
    '*** (Jen 2.13)

  '*** Jay 3.3(FG)
  ' ***** Manali Trading Module
  'Case Is = UCase("RepShpBill"), UCase("RepShpBillLoose") '*** (Bef 2.13)
  Case Is = UCase("RepShpBill"), UCase("RepShpBillLoose"), UCase("RepShpBillMulti"), UCase("RepFgShpBill"), UCase("RepJtShpBill")
  '*** Jay 3.3(FG)
    GRP_REP.SetFormula mRep, "wGrp1", wGrpFormula
    GRP_REP.SetFormula mRep, "wGrp1Desc", wGrpDescFormula
    GRP_REP.SetFormula mRep, "wDet", wDetFormula
    GRP_REP.SetFormula mRep, "wDetDesc", wDetDescFormula

    '***Geeta***Emr208
    If UCase(ADC("UoInDet")) = "" Then _
    GRP_REP.SetFormula mRep, "wDet", wGrpFormula
    '***Geeta***Emr208

    If wRsInHd.RecCount > 0 And wRsInHd!InNo <> 0 Then
      '*** (Bef 2.13) SetRepText mRep, "TxtAmtInWrds", "Total " + FnNumToWords(Format(wRsInHd!InFOBRs, "######0.00"), "INR")
      '*** (Jen 2.13)
      '(Round Off Problem)  SetRepText mRep, "TxtAmtInWrds", "Total " + FnNumToWords(Val(Format(wRsInHd!InTotRs - ((Round(wRsInHd!InDiaCost * wRsInHd!InCnvRt, 0) + Round(wRsInHd!InCSCost * wRsInHd!InCnvRt, 0) + Round(wRsInHd!InAdv * wRsInHd!InCnvRt, 0))), "#########0.00")), "INR")
      '(Round Off Problem)  SetRepText mRep, "TxtAmtInWrds", "Total " + FnNumToWords(Val(Format(wRsInHd!InTotRs - ((wRsInHd!InDiaCost + wRsInHd!InCSCost + wRsInHd!InAdv) * wRsInHd!InCnvRt), "#########0.00")), "INR")
      'SetRepText mRep, "TxtAmtInWrds", "Total " + FnNumToWords(Val(Format(wRsInHd!InTotRs - (Val(Format(wRsInHd!InDiaCost * wRsInHd!InCnvRt, "###########0")) + Val(Format(wRsInHd!InCSCost * wRsInHd!InCnvRt, "###########0")) + Val(Format(wRsInHd!InAdv * wRsInHd!InCnvRt, "###########0"))), "#########0.00")), ctIndCurncy)
      SetRepText mRep, "TxtAmtInWrds", "Total " + FnNumToWords(Format(wRsInHd!InFOBRs, "######0.00"), ctIndCurncy)
      '*** (Jen 2.13)
    End If
    
    mRep.TxtDocTyp.SetText ADC("wDocTyp")   '************Geeta**********Emr207 *** 210 For Local also ******
    
    mRep.SecForGR.Suppress = True   '*** (Jen 2.13)
    
    '*** (Bef 2.13)
    'Set wRsMetUtlsd = moCn.OpenRes("Select distinct CmName from Txndz M " + _
    '              "join Txndz B on B.TzCoCd= M.TzCoCd and B.TzYy= M.TzRefYy " + _
    '              "and B.TzKey= M.TzRefKey join Txn on TCoCd= B.TzCoCd and TTc= B.TzTc and " + _
    '              "TYy= B.TzYy and TChr= B.TzChr and TNo= B.TzNo " + _
    '              "join RmMst on RmCd= M.TzRmZCd " + _
    '              "join CustMst on CmCtg= 'S' and CmCd= TSuppCd " + _
    '              "where M.TzCoCd= '" + adc("UoCoCdFr") + "' and " + _
    '              "M.TzTc= '" + adc("UoInTcFr") + "' and M.TzYy= '" + adc("UoInYyFr") + "' " + _
    '              "and M.TzChr= '" + adc("UoInChrFr") + "' and M.TzNo= " + CStr(adc("UoInNoFr")) + " " + _
    '              "and RmCtg in ('G', 'P', 'S')")
    '*** (Bef 2.13)
    
    '*** (Jen 2.13)
    '****** Sachin 3.02 - Id fields in Joins
    ' **** Manali 3.03 - 25/06/08 - Pd Mod - 'L' added
    Set wRsMetUtlsd = moCn.OpenRes("Select distinct '('+CmName +'-> '+TBoeNo+')' as CmName from Txndz M " + _
                      "join Txndz B on B.TzCoCd= M.TzCoCd and B.TzYy= M.TzRefYy " + _
                      "and B.TzKey= M.TzRefKey join Txn on TIdNo=B.TzTIdNo and TCoCd= B.TzCoCd and TTc= B.TzTc and " + _
                      "TYy= B.TzYy and TChr= B.TzChr and TNo= B.TzNo " + IIF(gs_Partition = ctCurrPrtn, " And TPrtKey=B.TzPrtKey ", "") + _
                      "join RmMst on RmCd= M.TzRmZCd " + IIF(gs_Partition = ctCurrPrtn, " And RmPrtKey=M.TzPrtKey ", "") + _
                      "join CustMst on CmCtg= 'S' and CmCd= TSuppCd " + _
                      "join InvHd Ih on Ih.InIdNo=M.TzInIdNo and Ih.InCoCd= M.TzCoCd and Ih.InTc= M.TzTc and Ih.InYy= M.TzYy " + _
                      "and Ih.InChr= M.TzChr and Ih.InNo= M.TzNo " + IIF(gs_Partition = ctCurrPrtn, " And Ih.InPrtKey=M.TzPrtKey ", "") + wCnd + IIF(gs_Partition = ctCurrPrtn, " And M.TzPrtKey='" + ctCurrPrtn + "' ", "") + " and RmCtg in ('G', 'P', 'S', 'L','M')")
    '*** (Jen 2.13)
    Dim strLMDesc As String
    Dim wrsrmctgd  As MDORowSet
    Do While Not (wRsMetUtlsd.EOF Or wRsMetUtlsd.BOF)
      wMetUtlsdStr = wMetUtlsdStr + IIF(wMetUtlsdStr = "", "", ", ") + wRsMetUtlsd!CmName
      wRsMetUtlsd.MoveNext
    Loop
    SetRepText mRep, "TxtMetUtlsd", "METAL UTILISED IS METAL IMPORTED/ PURCHASED FROM: " + wMetUtlsdStr
    
    If UCase(ADC.MenuCd) = UCase("RepShpBill") Then
      Set wrsrmctgd = moCn.OpenRes("Select PMCd as RmCtg, PDesc from Param where Ptyp = 'RMCTG'")
      Do While Not (wrsrmctgd.EOF Or wrsrmctgd.BOF)
        If (wrsrmctgd!RmCtg = "G" Or wrsrmctgd!RmCtg = "P" Or wrsrmctgd!RmCtg = "L" Or wrsrmctgd!RmCtg = "M" Or wrsrmctgd!RmCtg = "S") Then
          If (wrsrmctgd!RmCtg = "L" Or wrsrmctgd!RmCtg = "M") Then
            strLMDesc = strLMDesc + IIF(strLMDesc = "", "", " + ") + wrsrmctgd!PDesc
          Else
            SetRepText mRep, "txtCtg" + wrsrmctgd!RmCtg, wrsrmctgd!PDesc
          End If

        End If
        wrsrmctgd.MoveNext
      Loop
      SetRepText mRep, "txtCtgL", strLMDesc
    End If
    '*** Jay 3.3(FG)
    ' ***** Manali Trading Module
    If UCase(ADC.MenuCd) <> UCase("RepFgShpBill") And UCase(ADC.MenuCd) <> UCase("RepJtShpBill") Then
    '*** Jay 3.3(FG)
      '*** (Jen 2.13)
      '****** Sachin 3.02 - Id fields in Joins
      wSubRepStr = "Select '1D' as qTag,Max(OdOmCmCd) as OmCmCd, DmCtg, (Select PDesc From Param where PTyp= 'DMCTG'" + _
                  " and PMCd= DmCtg and PSCd= '') as  DmCtgDesc, " + _
                  " (Select PDesc from Param where PTyp= 'SALCTG' and PMCd= DmCtg and " + _
                  " PSCd= max(DmSalCtg)) as DmSalCtgDesc, " + _
                  "max(DmUom) as DmUom, Max(" + Ws_IdKt + ") as IdKt, " + _
                  "Max(" + ws_IdVaCtg + ") as IdVaCtg, Max(IdGldLs) as IdGldLs, Max(OdDmCol) as OdDmCol, " + qMainMet + ", sum(" + ws_IdQty + ") as qIdQty " + _
                  " From InvDsg Join DsgMst On DmIdNo=IdDmIdNo and DmTcTyp='DM' and IdDmCd=DmCd and DmSz='' " + IIF(gs_Partition = ctCurrPrtn, " And IdPrtKey=DmPrtKey", "") + _
                  " join InvHd Ih on Ih.InIdNo=IdInIdNo and Ih.InCoCd= IdCoCd and Ih.InTc= IdTc and Ih.InYy= IdYy and Ih.InChr= IdChr and Ih.InNo= IdNo " + IIF(gs_Partition = ctCurrPrtn, " And Ih.InPrtKey=IdPrtKey ", "") + _
                  " join OrdDsg on OdCoCd= IdCoCd and OdTc= IdExpOdTc and OdYy= IdExpOdYy " + _
                  " and OdChr= IdExpOdChr and OdNo= IdExpOdNo and OdSr= IdExpOdSr " + IIF(gs_Partition = ctCurrPrtn, " And OdPrtKey=IdPrtKey", "") + wOrJoin + _
                  " " + wCnd + IIF(gs_Partition = ctCurrPrtn, " And IdPrtKey='" + ctCurrPrtn + "' ", "") + " " + IIF(wFlagLocalExp = "Y", " ", " and 1= 2 ") + _
                  " group by " + wSubRepGrpFormula + "DmCtg, DmUom "
    '*** Jay 3.3(FG)
    Else
      wSubRepStr = "Select '1D' as qTag,Max(OdOmCmCd) as OmCmCd, DmCtg, (Select PDesc From Param where PTyp= 'DMCTG'" + _
                  " and PMCd= DmCtg and PSCd= '') as  DmCtgDesc, " + _
                  " (Select PDesc from Param where PTyp= 'SALCTG' and PMCd= DmCtg and " + _
                  " PSCd= max(DmSalCtg)) as DmSalCtgDesc, " + _
                  "max(DmUom) as DmUom, Max(" + Ws_IdKt + ") as IdKt, " + _
                  "Max(" + ws_IdVaCtg + ") as IdVaCtg, 0 as IdGldLs, Max(OdDmCol) as OdDmCol, " + qMainMet + ", sum(" + ws_IdQty + ") as qIdQty " + _
                  " From OrdDsg Join DsgMst On DmIdNo=OdDmIdNo and DmTcTyp='DM' and OdDmCd=DmCd and DmSz='' " + IIF(gs_Partition = ctCurrPrtn, " And OdPrtKey=DmPrtKey", "") + _
                  " join InvHd Ih on Ih.InCoCd= OdCoCd and Ih.InTc= OdTc and Ih.InYy= OdYy and Ih.InChr= OdChr and Ih.InNo= OdNo " + IIF(gs_Partition = ctCurrPrtn, " And Ih.InPrtKey=OdPrtKey ", "") + _
                  " join OrdRm on " + IIF(gs_Partition = ctCurrPrtn, " OrPrtKey=OdPrtKey and ", "") + " OrOdIdNo=OdIdNo and OrCoCd= OdCoCd and OrTc= OdTc and OrYy= OdYy " + _
                  " and OrChr= OdChr and OrNo= OdNo and OrSr= OdSr and OrMainMet= 'Y' " + _
                  " " + wCnd + IIF(gs_Partition = ctCurrPrtn, " And OdPrtKey='" + ctCurrPrtn + "' ", "") + " " + IIF(wFlagLocalExp = "Y", " ", " and 1= 2 ") + _
                  " group by " + wSubRepGrpFormula + "DmCtg, DmUom "
    End If
    '*** Jay 3.3(FG)
        
    Call moCn.RepRes(mRep.Subreport2.OpenSubreport, wSubRepStr + ctMaxDopOpt) '****** Sachin 3.01 ctMaxDopOpt added
    SetSubFormula mRep.Subreport2, "wGrp", wGrpFormula
    '*** (Jen 2.13)
    
    
    '*** (24/02/06) For Kama (Jen 2.12)
    If wFlagLocalExp = "Y" Then
      '*** (Bef 2.13)
      'wSubRepStr = "Select '1D' as qTag, DmCtg, (Select PDesc From Param where PTyp='DMCTG'" + _
      '              " and PMCd=DmCtg  and PSCd='') as  DmCtgDesc, " + _
      '              " (Select PDesc from Param where PTyp= 'SALCTG' and PMCd= DmCtg and " + _
      '              " PSCd= max(DmSalCtg)) as DmSalCtgDesc, " + _
      '              "max(DmUom) as DmUom, Max(IdKt) as IdKt, " + _
      '              "Max(IdVaCtg) as IdVaCtg, Max(IdGldLs) as IdGldLs, Max(OdDmCol) as OdDmCol, " + qMainMet + ", sum(IdQty) as qIdQty " + _
      '              "From InvDsg Join DsgMst  " + _
      '              "On DmTcTyp='DM' and IdDmCd=DmCd and DmSz='' " + _
      '              "join OrdDsg on OdCoCd= IdCoCd and OdTc= IdExpOdTc and OdYy= IdExpOdYy " + _
      '              "and OdChr= IdExpOdChr and OdNo= IdExpOdNo and OdSr= IdExpOdSr " + wOrJoin + _
      '              "where IdCoCd= '" + adc("UoCoCdFr") + "' and IdTc= '" + adc("UoInTcFr") + "' and " + _
      '              "IdYy= '" + adc("UoInYyFr") + "' and IdChr= '" + adc("UoInChrFr") + "' and " + _
      '              "IdNo= " + CStr(adc("UoInNoFr")) + "  group by " + wSubRepGrpFormula + "DmCtg, DmUom "
                    
      'Call moCn.RepRes(mRep.Subreport1.OpenSubreport, wSubRepStr)
      'SetSubFormula mRep.Subreport1, "wGrp", wGrpFormulaTmp
      '*** (Bef 2.13)
      
      '*** (Bef 2.13) Jen
      'Set wRsFiness = moCn.OpenRes("Select RmCtg,(case when RmCtg='G' then  max(InHGldFiness) when RmCtg='P' then  max(InHPlFiness) " + _
      '                        "else max(InHSlFiness) end ) as qMetFiness , " + _
      '                        "max(InLmgCst) as qLmgVal, max(InLmpCst) as qLmpVal, max(InLmsCst) as qLmsVal, " + _
      '                        "Case When RmCtg ='G' Then 'Gold In '+Str((case when RmCtg='G' then  max(InHGldFiness) when RmCtg='P' then max(InHPlFiness) else max(InHSlFiness) end ),6,3)+ ' Fineness ' " + _
      '                        "     When RmCtg ='P' Then 'Platinum In '+ +Str((case when RmCtg='G' then  max(InHGldFiness) when RmCtg='P' then  max(InHPlFiness) else max(InHSlFiness)  end ),6,3)+ ' Fineness ' " + _
      '                        "     When RmCtg ='S' Then 'Silver In '+ +Str((case when RmCtg='G' then  max(InHGldFiness) when RmCtg='P' then max(InHPlFiness) else max(InHSlFiness) end ),6,3)+ ' Fineness ' " + _
      '                        "   End  as qFineStr " + _
      '                        "From InvHd Join InvRm on InCoCd= IrCoCd and InTc = IrTc and InYy= IrYy " + _
      '                        "and InChr= IrChr and  InNo= IrNo " + _
      '                        "join RmMst on RmCd = IrRmCd  where RmCtg in ('G', 'P', 'S') " + _
      '                        "and InCoCd= '" + adc("UoCoCdFr") + "'  and InTc= '" + adc("UoInTcFr") + "'" + _
      '                        "and InYy= '" + adc("UoInYyFr") + "' and InChr= '" + adc("UoInChrFr") + "'" + _
      '                        "and InNo=" + CStr(adc("UoInNoFr")) + " Group By RmCtg ")
      '*** (Bef 2.13) Jen
      
      '*** Jay 3.3(FG)
      ' ***** Manali Trading Module
      If UCase(ADC.MenuCd) <> UCase("RepFgShpBill") And UCase(ADC.MenuCd) <> UCase("RepJtShpBill") Then
      '*** Jay 3.3(FG)
        '*** (Jen 2.13)
        '****** Sachin 3.02 - Id fields in Joins
        ' **** Manali 3.03 - 25/06/08 - Pd Mod - 'L' added
        Set wRsFiness = moCn.OpenRes("Select RmCtg,(case when RmCtg= 'G' then  max(InHGldFiness) when RmCtg= 'P' then  max(InHPlFiness) " + _
                           "when RmCtg= 'S' then  max(InHSlFiness) else max(InHPdFiness) end ) as qMetFiness , " + _
                           "max(InLmgSal) as qLmgVal, max(InLmpSal) as qLmpVal, max(InLmsSal) as qLmsVal,  max(InLmlSal) as qLmlVal, " + _
                           "Case When RmCtg ='G' Then 'Gold In '+Str((case when RmCtg='G' then  max(InHGldFiness) when RmCtg='P' then max(InHPlFiness) when RmCtg='S' then max(InHSlFiness) else max(InHPdFiness) end ), 6, 3)+ ' Fineness ' " + _
                           "     When RmCtg ='P' Then 'Platinum In '+ +Str((case when RmCtg='G' then  max(InHGldFiness) when RmCtg='P' then  max(InHPlFiness) when RmCtg='S' then max(InHSlFiness) else max(InHPdFiness) end ), 6, 3)+ ' Fineness ' " + _
                           "     When RmCtg ='S' Then 'Silver In '+ +Str((case when RmCtg='G' then  max(InHGldFiness) when RmCtg='P' then max(InHPlFiness) when RmCtg='S' then max(InHSlFiness) else max(InHPdFiness) end ), 6, 3)+ ' Fineness ' " + _
                           "     When RmCtg ='L' Then 'Palladium In '+ +Str((case when RmCtg='G' then  max(InHGldFiness) when RmCtg='P' then max(InHPlFiness) when RmCtg='S' then max(InHSlFiness) else max(InHPdFiness) end ), 6, 3)+ ' Fineness ' " + _
                           "   End  as qFineStr " + _
                           "From InvHd Ih Join InvRm on InCoCd= IrCoCd and InTc = IrTc and InYy= IrYy " + _
                           "and InChr= IrChr and  InNo= IrNo " + IIF(gs_Partition = ctCurrPrtn, " And InPrtKey=IrPrtKey ", "") + _
                           "join RmMst on RmCd = IrRmCd " + IIF(gs_Partition = ctCurrPrtn, " And RmPrtKey=IrPrtKey ", "") + "  where " + IIF(gs_Partition = ctCurrPrtn, " InPrtKey='" + ctCurrPrtn + "' And ", "") + " RmCtg in ('G', 'P', 'S', 'L') " + _
                           wAndCnd + " Group By RmCtg ")
      '*** (Jen 2.13)
      '*** Jay 3.3(FG)
      Else
        Set wRsFiness = moCn.OpenRes("Select RmCtg,(case when RmCtg= 'G' then  max(InHGldFiness) when RmCtg= 'P' then  max(InHPlFiness) " + _
                           "when RmCtg= 'S' then  max(InHSlFiness) else max(InHPdFiness) end ) as qMetFiness , " + _
                           "max(InLmgSal) as qLmgVal, max(InLmpSal) as qLmpVal, max(InLmsSal) as qLmsVal,  max(InLmlSal) as qLmlVal, " + _
                           "Case When RmCtg ='G' Then 'Gold In '+Str((case when RmCtg='G' then  max(InHGldFiness) when RmCtg='P' then max(InHPlFiness) when RmCtg='S' then max(InHSlFiness) else max(InHPdFiness) end ), 6, 3)+ ' Fineness ' " + _
                           "     When RmCtg ='P' Then 'Platinum In '+ +Str((case when RmCtg='G' then  max(InHGldFiness) when RmCtg='P' then  max(InHPlFiness) when RmCtg='S' then max(InHSlFiness) else max(InHPdFiness) end ), 6, 3)+ ' Fineness ' " + _
                           "     When RmCtg ='S' Then 'Silver In '+ +Str((case when RmCtg='G' then  max(InHGldFiness) when RmCtg='P' then max(InHPlFiness) when RmCtg='S' then max(InHSlFiness) else max(InHPdFiness) end ), 6, 3)+ ' Fineness ' " + _
                           "     When RmCtg ='L' Then 'Palladium In '+ +Str((case when RmCtg='G' then  max(InHGldFiness) when RmCtg='P' then max(InHPlFiness) when RmCtg='S' then max(InHSlFiness) else max(InHPdFiness) end ), 6, 3)+ ' Fineness ' " + _
                           "   End  as qFineStr " + _
                           "From InvHd Ih Join OrdRm on InCoCd= OrCoCd and InTc = OrTc and InYy= OrYy " + _
                           "and InChr= OrChr and  InNo= OrNo " + IIF(gs_Partition = ctCurrPrtn, " And InPrtKey=OrPrtKey ", "") + _
                           "join RmMst on RmCd = OrRmCd " + IIF(gs_Partition = ctCurrPrtn, " And RmPrtKey=OrPrtKey ", "") + "  where " + IIF(gs_Partition = ctCurrPrtn, " InPrtKey='" + ctCurrPrtn + "' And ", "") + " RmCtg in ('G', 'P', 'S', 'L') " + _
                           wAndCnd + " Group By RmCtg ")
      End If
      '*** Jay 3.3(FG)
      
      If wRsFiness.RecCount > 0 Then _
         ws_FinessStr = "Statement Showing "
      RsCnt = 1
      
      '*** (Bef 2.13)
      'wd_GldOunce = moCn.GetFldVal("Select PNum From Param Where PTyp='OUNCE' and PMCd='G' and PSCd=''")
      'wd_PlOunce = moCn.GetFldVal("Select PNum From Param Where PTyp='OUNCE' and PMCd='P' and PSCd=''")
      '' ****** Sachin 2.12 06/12/05 - SJM
      'wd_SlOunce = moCn.GetFldVal("Select PNum From Param Where PTyp='OUNCE' and PMCd='S' and PSCd=''")
      '' ****** Sachin 2.12 06/12/05 - SJM
      '*** (Bef 2.13)
      
      With wRsFiness
        Do While Not (.BOF Or .EOF)
          ws_FinessStr = ws_FinessStr + !qFineStr
          ws_FinessStr = ws_FinessStr + IIF(.RecCount > 1 And RsCnt < .RecCount, " And ", "")
          RsCnt = RsCnt + 1
          wd_LmgGms = Round(!qLmgVal / IIF(wd_GldOunce > 0, wd_GldOunce, 1), 4)
          wd_LmpGms = Round(!qLmpVal / IIF(wd_PlOunce > 0, wd_PlOunce, 1), 4)
          ' ****** Sachin 2.12 06/12/05 - SJM
          wd_LmsGms = Round(!qLmsVal / IIF(wd_SlOunce > 0, wd_SlOunce, 1), 4)
          ' ****** Sachin 2.12 06/12/05 - SJM
          ' **** Manali 3.03 - 25/06/08 - Pd Mod
          wd_LmlGms = Round(!qLmlVal / IIF(wd_PdOunce > 0, wd_PdOunce, 1), 4)
          ' **** Manali 3.03 - 25/06/08 - Pd Mod
          ws_GldStr = IIF(!RmCtg = "G", ws_CurCd + " " + Str(!qLmgVal) + " For " + ws_GldOunceDsc + " " + ws_CurCd + " " + _
                            Str(wd_LmgGms) + " For " + Str(Round(!qMetFiness, 3)) + " Fineness  " + ws_CurCd + " " + _
                            Str(Round(wd_LmgGms / !qMetFiness, 4)) + " For Pure 1 Gram ", ws_GldStr)
          ws_PlStr = IIF(!RmCtg = "P", ws_CurCd + " " + Str(!qLmpVal) + " For " + ws_PlOunceDsc + " " + ws_CurCd + " " + _
                            Str(wd_LmpGms) + " For " + Str(Round(!qMetFiness, 3)) + " Fineness  " + ws_CurCd + " " + _
                            Str(Round(wd_LmpGms / !qMetFiness, 4)) + " For Pure 1 Gram ", ws_PlStr)
          ' ****** Sachin 2.12 - 06/12/05 - SJM
          ws_SlvStr = IIF(!RmCtg = "S", ws_CurCd + " " + Str(!qLmsVal) + " For " + ws_SlOunceDsc + " " + ws_CurCd + " " + _
                            Str(wd_LmsGms) + " For " + Str(Round(!qMetFiness, 3)) + " Fineness  " + ws_CurCd + " " + _
                            Str(Round(wd_LmsGms / !qMetFiness, 4)) + " For Pure 1 Gram ", ws_SlvStr)
          ' ****** Sachin 2.12 - 06/12/05 - SJM
          ' **** Manali 3.03 - 25/06/08 - Pd Mod
          ws_PdStr = IIF(!RmCtg = "L", ws_CurCd + " " + Str(!qLmlVal) + " For " + ws_PdOunceDsc + " " + ws_CurCd + " " + _
                            Str(wd_LmlGms) + " For " + Str(Round(!qMetFiness, 3)) + " Fineness  " + ws_CurCd + " " + _
                            Str(Round(wd_LmlGms / !qMetFiness, 4)) + " For Pure 1 Gram ", ws_PdStr)
          ' **** Manali 3.03 - 25/06/08 - Pd Mod
          wRsFiness.MoveNext
        Loop
      End With
      ' **** Manali 3.03 - 25/06/08 - Pd Mod
      mRep.TxtOunceRt.SetText ws_GldStr + vbCr + ws_PlStr + vbCr + ws_SlvStr + vbCr + ws_PdStr
      
      'Change line below to single line in 4.1.3.0 as getting error 'Procedure too large'
      mRep.TxtOunceRt.Suppress = IIF(ws_GldStr = "" And ws_PlStr = "" And ws_SlvStr = "" And ws_PdStr = "", True, False)
      
'      If ws_GldStr = "" And ws_PlStr = "" And ws_SlvStr = "" And ws_PdStr = "" Then
'      ' **** Manali 3.03 - 25/06/08 - Pd Mod
'         mRep.TxtOunceRt.Suppress = True
'      Else
'          mRep.TxtOunceRt.Suppress = False
'      End If
    
      GRP_REP.SetFormula mRep, "wAuthSgn", "{rdo.InSelfName}"     '*** (Jen 2.13)
    
    Else
      '*** (Jen 2.13)
      GRP_REP.SetFormula mRep, "wAuthSgn", "{rdo.InHShpBillSgnBy}"
      mRep.TxtExpTrdCtrl.Suppress = True: mRep.TxtInHItcNo.Suppress = True
      mRep.FldInHItcNo.Suppress = True: mRep.FldInHEximPolicy.Suppress = True
      mRep.FldInCmPayTerms.Suppress = True: mRep.FldInCmPayTerms1.Suppress = True
      mRep.FldInHRbiPolicy.Suppress = True: mRep.FldInMetalUtil.Suppress = True
      mRep.TxtInHGjepcNo.Suppress = True: mRep.FldInHGjepcNo.Suppress = True
      mRep.TxtInHGjepcValidUpTo.Suppress = True: mRep.FldInHGjepcValidUpTo.Suppress = True
      mRep.TxtOunceRt.Suppress = True: mRep.SecDescSubRep.Suppress = True
      '*** (Jen 2.13)
    End If
    '*** (24/02/06) For Kama (Jen 2.12)
    
    '*** (Jen 2.13)
    mRep.TxtCurCd.SetText ws_CurCd
    mRep.TxtCurCd1.SetText ws_CurCd
    mRep.TxtCurCd2.SetText ws_CurCd
    mRep.TxtCurCd3.SetText ws_CurCd
    mRep.TxtCurCd4.SetText ws_CurCd
    mRep.TxtCurCd5.SetText ws_CurCd
    
    mRep.TxtCurSymbol1.SetText ws_CurCd
    mRep.TxtCurSymbol2.SetText ws_CurCd
    mRep.TxtCurSymbol3.SetText ws_CurCd
    mRep.TxtCurSymbol4.SetText ws_CurCd
    
    mRep.TxtOneCurCd.SetText "ONE " + ws_CurCd + " ="
    mRep.TxtNetRealisation.SetText "NET REALISATION " + ws_CurCd
    mRep.TxtFOB.SetText "FOB " + ws_CurCd
    '*** (Jen 2.13)
    
  '*** Jay 3.3(FG)
  ' ***** Manali Trading Module
  Case Is = UCase("RepGRFormFrnt"), UCase("RepGrFormLoose"), UCase("RepGrFormFrntMulti"), UCase("RepFgGRFrnt"), UCase("RepJtGRFrnt")
  '*** Jay 3.3(FG)
    GRP_REP.SetFormula mRep, "wGrp1", wGrpFormula
    GRP_REP.SetFormula mRep, "wGrp1Desc", wGrpDescFormula
    GRP_REP.SetFormula mRep, "wDet", wDetFormula
    GRP_REP.SetFormula mRep, "wDetDesc", wDetDescFormula
    
    '***Geeta***Emr208
    If UCase(ADC("UoInDet")) = "" Then _
    GRP_REP.SetFormula mRep, "wDet", wGrpFormula
    '***Geeta***Emr208

    If wRsInHd.RecCount > 0 And wRsInHd!InNo <> 0 Then
      '*** (Bef 2.13) SetRepText mRep, "TxtAmtInWrds", "Total " + FnNumToWords(Format(wRsInHd!InFOBRs, "######0.00"), "INR")
       
      '*** (Jen 2.13)
      'SetRepText mRep, "TxtAmtInWrds", "Total " + FnNumToWords(Val(Format(wRsInHd!InTotRs - (Val(Format(wRsInHd!InDiaCost * wRsInHd!InCnvRt, "###########0")) + Val(Format(wRsInHd!InCSCost * wRsInHd!InCnvRt, "###########0")) + Val(Format(wRsInHd!InAdv * wRsInHd!InCnvRt, "###########0"))), "#########0.00")), ctIndCurncy)
      SetRepText mRep, "TxtAmtInWrds", "Total " + FnNumToWords(Format(wRsInHd!InFOBRs, "######0.00"), ctIndCurncy)
      '*** (Jen 2.13)
    End If
    
    mRep.TxtDocTyp.SetText ADC("wDocTyp")   '************Geeta**********Emr207*210 For Local also*******
    
    '*** (Jen 2.14 Next)
    If ADC("UoYN7") = "Y" Then
      Rep2.FldInCmPayTerms.Suppress = False
    Else
      Rep2.FldInCmPayTerms.Suppress = True
    End If
    '*** (Jen 2.14 Next)
    
    '*** (Bef 2.13)
    'Set wRsMetUtlsd = moCn.OpenRes("Select distinct CmName from Txndz M " + _
    '              "join Txndz B on B.TzCoCd= M.TzCoCd and B.TzYy= M.TzRefYy " + _
    '              "and B.TzKey= M.TzRefKey join Txn on TCoCd= B.TzCoCd and TTc= B.TzTc and " + _
    '              "TYy= B.TzYy and TChr= B.TzChr and TNo= B.TzNo " + _
    '              "join RmMst on RmCd= M.TzRmZCd " + _
    '              "join CustMst on CmCtg= 'S' and CmCd= TSuppCd " + _
    '              "where M.TzCoCd= '" + adc("UoCoCdFr") + "' and " + _
    '              "M.TzTc= '" + adc("UoInTcFr") + "' and M.TzYy= '" + adc("UoInYyFr") + "' " + _
    '              "and M.TzChr= '" + adc("UoInChrFr") + "' and M.TzNo= " + CStr(adc("UoInNoFr")) + " " + _
    '              "and RmCtg in ('G', 'P', 'S')")
    '*** (Bef 2.13)
                  
    '*** (Jen 2.13)
    '****** Sachin 3.02 - Id fields in Joins
    ' **** Manali 3.03 - 25/06/08 - Pd Mod - 'L' added
    Set wRsMetUtlsd = moCn.OpenRes("Select distinct CmName from Txndz M " + _
                  "join Txndz B on " + IIF(gs_Partition = ctCurrPrtn, " B.TzPrtKey=M.TzPrtKey and ", "") + " B.TzCoCd= M.TzCoCd and B.TzYy= M.TzRefYy " + _
                  "and B.TzKey= M.TzRefKey join Txn on TCoCd= B.TzCoCd and TTc= B.TzTc and " + _
                  "TYy= B.TzYy and TChr= B.TzChr and TNo= B.TzNo " + _
                  "join RmMst on " + IIF(gs_Partition = ctCurrPrtn, " RmPrtKey=M.TzPrtKey and ", "") + " RmCd= M.TzRmZCd " + _
                  "join CustMst on CmCtg= 'S' and CmCd= TSuppCd " + _
                  "join InvHd Ih on Ih.InIdNo=M.TzInIdNo and Ih.InCoCd= M.TzCoCd and Ih.InTc= M.TzTc and Ih.InYy= M.TzYy " + _
                  "and Ih.InChr= M.TzChr and Ih.InNo= M.TzNo " + IIF(gs_Partition = ctCurrPrtn, " And Ih.InPrtKey=M.TzPrtKey ", "") + _
                  wCnd + IIF(gs_Partition = ctCurrPrtn, " And M.TzPrtKey='" + ctCurrPrtn + "' ", "") + " and RmCtg in ('G', 'P', 'S', 'L','M')")
    '*** (Jen 2.13)
                  
    Do While Not (wRsMetUtlsd.EOF Or wRsMetUtlsd.BOF)
      wMetUtlsdStr = wMetUtlsdStr + IIF(wMetUtlsdStr = "", "", ", ") + wRsMetUtlsd!CmName
      wRsMetUtlsd.MoveNext
    Loop
    SetRepText mRep, "TxtMetUtlsd", "METAL UTILISED IS METAL IMPORTED/ PURCHASED FROM: " + wMetUtlsdStr
    
    '*** (24/02/06) For Kama (Jen 2.12)
    If wFlagLocalExp = "Y" Then
                    
      '*** (Bef 2.13)
      'wSubRepStr = "Select '1D' as qTag, DmCtg, (Select PDesc From Param where PTyp='DMCTG'" + _
      '              " and PMCd=DmCtg  and PSCd='') as  DmCtgDesc, " + _
      '              " (Select PDesc from Param where PTyp= 'SALCTG' and PMCd= DmCtg and " + _
      '              " PSCd= max(DmSalCtg)) as DmSalCtgDesc, " + _
      '              "max(DmUom) as DmUom, Max(IdKt) as IdKt, " + _
      '              "Max(IdVaCtg) as IdVaCtg, Max(IdGldLs) as IdGldLs, Max(OdDmCol) as OdDmCol, " + qMainMet + ", sum(IdQty) as qIdQty " + _
      '              "From InvDsg Join DsgMst  " + _
      '              "On DmTcTyp='DM' and IdDmCd=DmCd and DmSz='' " + _
      '              "join OrdDsg on OdCoCd= IdCoCd and OdTc= IdExpOdTc and OdYy= IdExpOdYy " + _
      '              "and OdChr= IdExpOdChr and OdNo= IdExpOdNo and OdSr= IdExpOdSr " + wOrJoin + _
      '              "where IdCoCd= '" + adc("UoCoCdFr") + "' and IdTc= '" + adc("UoInTcFr") + "' and " + _
      '              "IdYy= '" + adc("UoInYyFr") + "' and IdChr= '" + adc("UoInChrFr") + "' and " + _
      '              "IdNo= " + CStr(adc("UoInNoFr")) + "  group by " + wSubRepGrpFormula + "DmCtg,DmUom"
      
      'Call moCn.RepRes(mRep.Subreport1.OpenSubreport, wSubRepStr)
      'SetSubFormula mRep.Subreport1, "wGrp", wGrpFormula
      '*** (Bef 2.13)
      
      '*** Jay 3.3(FG)
      ' ***** Manali Trading Module
      If UCase(ADC.MenuCd) <> UCase("RepFgGRFrnt") And UCase(ADC.MenuCd) <> UCase("RepJtGRFrnt") Then
      '*** Jay 3.3(FG)
        '*** (Jen 2.13)
        '****** Sachin 3.02 - Id fields in Joins
        wSubRepStr = "Select '1D' as qTag,Max(OdOmCmCd) as OmCmCd, DmCtg, (Select PDesc From Param where PTyp= 'DMCTG'" + _
              " and PMCd= DmCtg  and PSCd= '') as  DmCtgDesc, " + _
              " (Select PDesc from Param where PTyp= 'SALCTG' and PMCd= DmCtg and " + _
              " PSCd= max(DmSalCtg)) as DmSalCtgDesc, " + _
              "max(DmUom) as DmUom, Max(" + Ws_IdKt + ") as IdKt, Max(" + ws_IdVaCtg + ") as IdVaCtg, " + _
              "Max(IdGldLs) as IdGldLs, Max(OdDmCol) as OdDmCol, " + qMainMet + ", sum(" + ws_IdQty + ") as qIdQty " + _
              "From InvDsg Join DsgMst  " + _
              "On DmIdNo=IddmIdNo and DmTcTyp= 'DM' and IdDmCd= DmCd and DmSz= '' " + IIF(gs_Partition = ctCurrPrtn, " And IdPrtKey=DmPrtKey ", "") + _
              "join InvHd Ih on Ih.InIdNo=IdInIdNo and Ih.InCoCd= IdCoCd and Ih.InTc= IdTc and Ih.InYy= IdYy and Ih.InChr= IdChr and Ih.InNo= IdNo " + IIF(gs_Partition = ctCurrPrtn, " And Ih.InPrtKey=IdPrtKey ", "") + _
              "join OrdDsg on OdCoCd= IdCoCd and OdTc= IdExpOdTc and OdYy= IdExpOdYy " + _
              "and OdChr= IdExpOdChr and OdNo= IdExpOdNo and OdSr= IdExpOdSr " + IIF(gs_Partition = ctCurrPrtn, " And OdPrtKey=IdPrtKey", "") + wOrJoin + _
              " " + IIF(gs_Partition = ctCurrPrtn, " And IdPrtKey='" + ctCurrPrtn + "' ", "") + wCnd + " group by " + wSubRepGrpFormula + "DmCtg, DmUom"
      '*** Jay 3.3(FG)
      Else
        wSubRepStr = "Select '1D' as qTag,Max(OdOmCmCd) as OmCmCd, DmCtg, (Select PDesc From Param where PTyp= 'DMCTG'" + _
              " and PMCd= DmCtg  and PSCd= '') as  DmCtgDesc, " + _
              " (Select PDesc from Param where PTyp= 'SALCTG' and PMCd= DmCtg and " + _
              " PSCd= max(DmSalCtg)) as DmSalCtgDesc, " + _
              "max(DmUom) as DmUom, Max(" + Ws_IdKt + ") as IdKt, Max(" + ws_IdVaCtg + ") as IdVaCtg, " + _
              "0 as IdGldLs, Max(OdDmCol) as OdDmCol, " + qMainMet + ", sum(" + ws_IdQty + ") as qIdQty " + _
              "From OrdDsg Join DsgMst  " + _
              "On DmIdNo=OddmIdNo and DmTcTyp= 'DM' and OdDmCd= DmCd and DmSz= '' " + IIF(gs_Partition = ctCurrPrtn, " And OdPrtKey=DmPrtKey ", "") + _
              "join InvHd Ih on Ih.InCoCd= OdCoCd and Ih.InTc= OdTc and Ih.InYy= OdYy and Ih.InChr= OdChr and Ih.InNo= OdNo " + IIF(gs_Partition = ctCurrPrtn, " And Ih.InPrtKey=OdPrtKey ", "") + _
              " join OrdRm on " + IIF(gs_Partition = ctCurrPrtn, " OrPrtKey=OdPrtKey and ", "") + " OrOdIdNo=OdIdNo and OrCoCd= OdCoCd and OrTc= OdTc and OrYy= OdYy " + _
              " and OrChr= OdChr and OrNo= OdNo and OrSr= OdSr and OrMainMet= 'Y' " + _
              " " + IIF(gs_Partition = ctCurrPrtn, " And OdPrtKey='" + ctCurrPrtn + "' ", "") + wCnd + " group by " + wSubRepGrpFormula + "DmCtg, DmUom"
      End If
      '*** Jay 3.3(FG)
      
      Call moCn.RepRes(mRep.Subreport2.OpenSubreport, wSubRepStr + ctMaxDopOpt) '****** Sachin 3.01 ctMaxDopOpt added
      SetSubFormula mRep.Subreport2, "wGrp", wGrpFormula
      '*** (Jen 2.13)
            
      '*** (Bef 2.13)
      'Set wRsFiness = moCn.OpenRes("Select RmCtg,(case when RmCtg='G' then  max(InHGldFiness) when RmCtg='P' then  max(InHPlFiness) " + _
      '                        "else max(InHSlFiness) end ) as qMetFiness , " + _
      '                        "max(InLmgCst) as qLmgVal,max(InLmpCst) as qLmpVal, max(InLmsCst) as qLmsVal, " + _
      '                        "Case When RmCtg ='G' Then 'Gold In '+Str((case when RmCtg='G' then  max(InHGldFiness) when RmCtg='P' then max(InHPlFiness) else max(InHSlFiness) end ),6,3)+ ' Fineness ' " + _
      '                        "     When RmCtg ='P' Then 'Platinum In ' + Str((case when RmCtg='G' then  max(InHGldFiness) when RmCtg='P' then max(InHPlFiness) else max(InHSlFiness) end ),6,3)+ ' Fineness '  " + _
      '                        "     When RmCtg ='S' Then 'Silver In ' + Str((case when RmCtg='G' then  max(InHGldFiness) when RmCtg='P' then max(InHPlFiness) else max(InHSlFiness) end ),6,3)+ ' Fineness '  " + _
      '                        "  End  as qFineStr " + _
      '                        "From InvHd Join InvRm on InCoCd= IrCoCd and InTc = IrTc and InYy= IrYy " + _
      '                        "and InChr= IrChr and  InNo= IrNo " + _
      '                        "join RmMst on RmCd = IrRmCd  where RmCtg in ('G', 'P', 'S') " + _
      '                        "and InCoCd= '" + adc("UoCoCdFr") + "'  and InTc= '" + adc("UoInTcFr") + "'" + _
      '                        "and InYy= '" + adc("UoInYyFr") + "' and InChr= '" + adc("UoInChrFr") + "'" + _
      '                        "and InNo=" + CStr(adc("UoInNoFr")) + " Group By RmCtg ")
      '*** (Bef 2.13)
      
      '*** Jay 3.3(FG)
      ' ***** Manali Trading Module
      If UCase(ADC.MenuCd) <> UCase("RepFgGRFrnt") And UCase(ADC.MenuCd) <> UCase("RepJtGRFrnt") Then
      '*** Jay 3.3(FG)
        '*** (Jen 2.13)
        '****** Sachin 3.02 - Id fields in Joins
        ' **** Manali 3.03 - 25/06/08 - Pd Mod - 'L' added
        Set wRsFiness = moCn.OpenRes("Select RmCtg,(case when RmCtg='G' then  max(InHGldFiness) when RmCtg='P' then  max(InHPlFiness) " + _
                              "when RmCtg='S' then  max(InHSlFiness) else max(InHPdFiness) end ) as qMetFiness , " + _
                              "max(InLmgSal) as qLmgVal, max(InLmpSal) as qLmpVal, max(InLmsSal) as qLmsVal, max(InLmlSal) as qLmlVal, " + _
                              "Case When RmCtg ='G' Then 'Gold In '+Str((case when RmCtg='G' then  max(InHGldFiness) when RmCtg='P' then max(InHPlFiness) when RmCtg='S' then max(InHSlFiness) else max(InHPdFiness) end ),6,3)+ ' Fineness ' " + _
                              "     When RmCtg ='P' Then 'Platinum In ' + Str((case when RmCtg='G' then  max(InHGldFiness) when RmCtg='P' then max(InHPlFiness) when RmCtg='S' then max(InHSlFiness) else max(InHPdFiness) end ),6,3)+ ' Fineness '  " + _
                              "     When RmCtg ='S' Then 'Silver In ' + Str((case when RmCtg='G' then  max(InHGldFiness) when RmCtg='P' then max(InHPlFiness) when RmCtg='S' then max(InHSlFiness) else max(InHPdFiness) end ),6,3)+ ' Fineness '  " + _
                              "     When RmCtg ='L' Then 'Palladium In ' + Str((case when RmCtg='G' then  max(InHGldFiness) when RmCtg='P' then max(InHPlFiness) when RmCtg='S' then max(InHSlFiness) else max(InHPdFiness) end ),6,3)+ ' Fineness '  " + _
                              "  End  as qFineStr " + _
                              "From InvHd Ih Join InvRm on InCoCd= IrCoCd and InTc = IrTc and InYy= IrYy " + _
                              "and InChr= IrChr and InNo= IrNo " + IIF(gs_Partition = ctCurrPrtn, " And InPrtKey=IrPrtKey ", "") + _
                              "join RmMst on RmCd = IrRmCd  " + IIF(gs_Partition = ctCurrPrtn, " And RmPrtKey=IrPrtKey ", "") + _
                              "where " + IIF(gs_Partition = ctCurrPrtn, " InPrtKey='" + ctCurrPrtn + "' And ", "") + " RmCtg in ('G', 'P', 'S', 'L') " + wAndCnd + " Group By RmCtg ")
      '*** (Jen 2.13)
      '*** Jay 3.3(FG)
      Else
        Set wRsFiness = moCn.OpenRes("Select RmCtg,(case when RmCtg='G' then  max(InHGldFiness) when RmCtg='P' then  max(InHPlFiness) " + _
                              "when RmCtg='S' then  max(InHSlFiness) else max(InHPdFiness) end ) as qMetFiness , " + _
                              "max(InLmgSal) as qLmgVal, max(InLmpSal) as qLmpVal, max(InLmsSal) as qLmsVal, max(InLmlSal) as qLmlVal, " + _
                              "Case When RmCtg ='G' Then 'Gold In '+Str((case when RmCtg='G' then  max(InHGldFiness) when RmCtg='P' then max(InHPlFiness) when RmCtg='S' then max(InHSlFiness) else max(InHPdFiness) end ),6,3)+ ' Fineness ' " + _
                              "     When RmCtg ='P' Then 'Platinum In ' + Str((case when RmCtg='G' then  max(InHGldFiness) when RmCtg='P' then max(InHPlFiness) when RmCtg='S' then max(InHSlFiness) else max(InHPdFiness) end ),6,3)+ ' Fineness '  " + _
                              "     When RmCtg ='S' Then 'Silver In ' + Str((case when RmCtg='G' then  max(InHGldFiness) when RmCtg='P' then max(InHPlFiness) when RmCtg='S' then max(InHSlFiness) else max(InHPdFiness) end ),6,3)+ ' Fineness '  " + _
                              "     When RmCtg ='L' Then 'Palladium In ' + Str((case when RmCtg='G' then  max(InHGldFiness) when RmCtg='P' then max(InHPlFiness) when RmCtg='S' then max(InHSlFiness) else max(InHPdFiness) end ),6,3)+ ' Fineness '  " + _
                              "  End  as qFineStr " + _
                              "From InvHd Ih Join OrdRm on InCoCd= OrCoCd and InTc = OrTc and InYy= OrYy " + _
                              "and InChr= OrChr and InNo= OrNo " + IIF(gs_Partition = ctCurrPrtn, " And InPrtKey=OrPrtKey ", "") + _
                              "join RmMst on RmCd = OrRmCd  " + IIF(gs_Partition = ctCurrPrtn, " And RmPrtKey=OrPrtKey ", "") + _
                              "where " + IIF(gs_Partition = ctCurrPrtn, " InPrtKey='" + ctCurrPrtn + "' And ", "") + " RmCtg in ('G', 'P', 'S', 'L') " + wAndCnd + " Group By RmCtg ")
      End If
      '*** Jay 3.3(FG)
      
      If wRsFiness.RecCount > 0 Then _
         ws_FinessStr = "Statement Showing "
      RsCnt = 1
      
      '*** (Bef 2.13)
      'wd_GldOunce = moCn.GetFldVal("Select PNum From Param Where PTyp= 'OUNCE' and PMCd= 'G' and PSCd= ''")
      'wd_PlOunce = moCn.GetFldVal("Select PNum From Param Where PTyp= 'OUNCE' and PMCd= 'P' and PSCd= ''")
      '' ****** Sachin 2.12 06/12/05 - SJM
      'wd_SlOunce = moCn.GetFldVal("Select PNum From Param Where PTyp= 'OUNCE' and PMCd= 'S' and PSCd= ''")
      '' ****** Sachin 2.12 06/12/05 - SJM
      '*** (Bef 2.13)
      
      With wRsFiness
        Do While Not (.BOF Or .EOF)
          ws_FinessStr = ws_FinessStr + !qFineStr
          ws_FinessStr = ws_FinessStr + IIF(.RecCount > 1 And RsCnt < .RecCount, " And ", "")
          RsCnt = RsCnt + 1
          wd_LmgGms = Round(!qLmgVal / IIF(wd_GldOunce > 0, wd_GldOunce, 1), 4)
          wd_LmpGms = Round(!qLmpVal / IIF(wd_PlOunce > 0, wd_PlOunce, 1), 4)
          ' ****** Sachin 2.12 06/12/05 - SJM
          wd_LmsGms = Round(!qLmsVal / IIF(wd_SlOunce > 0, wd_SlOunce, 1), 4)
          ' ****** Sachin 2.12 06/12/05 - SJM
          ' **** Manali 3.03 - 25/06/08 - Pd Mod
          wd_LmlGms = Round(!qLmlVal / IIF(wd_PdOunce > 0, wd_PdOunce, 1), 4)
          ' **** Manali 3.03 - 25/06/08 - Pd Mod
          ws_GldStr = IIF(!RmCtg = "G", ws_CurCd + " " + Str(!qLmgVal) + " For " + ws_GldOunceDsc + " " + ws_CurCd + " " + _
                            Str(wd_LmgGms) + " For " + Str(Round(!qMetFiness, 3)) + " Fineness  " + ws_CurCd + " " + _
                            Str(Round(wd_LmgGms / !qMetFiness, 4)) + " For Pure 1 Gram ", ws_GldStr)
          ws_PlStr = IIF(!RmCtg = "P", ws_CurCd + " " + Str(!qLmpVal) + " For " + ws_PlOunceDsc + " " + ws_CurCd + " " + _
                            Str(wd_LmpGms) + " For " + Str(Round(!qMetFiness, 3)) + " Fineness  " + ws_CurCd + " " + _
                            Str(Round(wd_LmpGms / !qMetFiness, 4)) + " For Pure 1 Gram ", ws_PlStr)
          ' ****** Sachin 2.12 06/12/05 - SJM
          ws_SlvStr = IIF(!RmCtg = "S", ws_CurCd + " " + Str(!qLmsVal) + " For " + ws_SlOunceDsc + " " + ws_CurCd + " " + _
                            Str(wd_LmsGms) + " For " + Str(Round(!qMetFiness, 3)) + " Fineness  " + ws_CurCd + " " + _
                            Str(Round(wd_LmsGms / !qMetFiness, 4)) + " For Pure 1 Gram ", ws_SlvStr)
          ' ****** Sachin 2.12 06/12/05 - SJM
          ' **** Manali 3.03 - 25/06/08 - Pd Mod
          ws_PdStr = IIF(!RmCtg = "L", ws_CurCd + " " + Str(!qLmlVal) + " For " + ws_PdOunceDsc + " " + ws_CurCd + " " + _
                            Str(wd_LmlGms) + " For " + Str(Round(!qMetFiness, 3)) + " Fineness  " + ws_CurCd + " " + _
                            Str(Round(wd_LmlGms / !qMetFiness, 4)) + " For Pure 1 Gram ", ws_PdStr)
          ' **** Manali 3.03 - 25/06/08 - Pd Mod
          wRsFiness.MoveNext
        Loop
      End With
      ' **** Manali 3.03 - 25/06/08 - Pd Mod
      mRep.TxtOunceRt.SetText ws_GldStr + vbCr + ws_PlStr + vbCr + ws_SlvStr + vbCr + ws_PdStr
      If ws_GldStr = "" And ws_PlStr = "" And ws_SlvStr = "" And ws_PdStr = "" Then
      ' **** Manali 3.03 - 25/06/08 - Pd Mod
         mRep.TxtOunceRt.Suppress = True
      Else
          mRep.TxtOunceRt.Suppress = False
      End If

      '*** (Jen 2.13)
      GRP_REP.SetFormula mRep, "wAuthSgn", "{rdo.InSelfName}"     '*** (Jen 2.13)
      'mRep.SecForGR.Suppress = False
      mRep.TxtHead.Suppress = True
 'mRep.TxtInHBinNo.Suppress = True: mRep.FldInHBinNo.Suppress = True   '**GRForm Front Multi-inv giving error here**
  '    If UCase(adc.MenuCd) = UCase("RepInv") Then
   '     mRep.FldCERegnNo.Suppress = True
    '  End If

      mRep.TxtInMMTc.Suppress = True: mRep.FldInMMTc.Suppress = True
      
      mRep.TxtCurCd.SetText ws_CurCd
      mRep.TxtCurCd1.SetText ws_CurCd
      mRep.TxtCurCd2.SetText ws_CurCd
      mRep.TxtCurCd3.SetText ws_CurCd
      mRep.TxtCurCd4.SetText ws_CurCd
      mRep.TxtCurCd5.SetText ws_CurCd
      
      mRep.TxtCurSymbol1.SetText ws_CurCd
      mRep.TxtCurSymbol2.SetText ws_CurCd
      mRep.TxtCurSymbol3.SetText ws_CurCd
      mRep.TxtCurSymbol4.SetText ws_CurCd
      
      mRep.TxtOneCurCd.SetText "ONE " + ws_CurCd + " ="
      mRep.TxtNetRealisation.SetText "NET REALISATION " + ws_CurCd
      mRep.TxtFOB.SetText "FOB " + ws_CurCd
      '*** (Jen 2.13)
    
    Else
      '*** (Jen 2.13)
      mRep.TxtCurCd.SetText ws_CurCd
      mRep.TxtCurCd1.SetText ws_CurCd
      mRep.TxtCurCd2.SetText ws_CurCd
      mRep.TxtCurCd3.SetText ws_CurCd
      mRep.TxtCurCd4.SetText ws_CurCd
      mRep.TxtCurCd5.SetText ws_CurCd
      
      mRep.TxtCurSymbol1.SetText ws_CurCd
      mRep.TxtCurSymbol2.SetText ws_CurCd
      mRep.TxtCurSymbol3.SetText ws_CurCd
      mRep.TxtCurSymbol4.SetText ws_CurCd
      
      mRep.TxtOneCurCd.SetText "ONE " + ws_CurCd + " ="
      mRep.TxtNetRealisation.SetText "NET REALISATION " + ws_CurCd
      mRep.TxtFOB.SetText "FOB " + ws_CurCd
      '*** (Jen 2.13)
    End If
    '*** (24/02/06) For Kama (Jen 2.12)
    
      '*** Jay 2.14 P2
  Case Is = UCase("RepGrFormFrntNew"), UCase("RepGrFormNewLoose"), UCase("RepGrFormFrntNewMul")
    'GRP_REP.SetFormula mRep, "wGrp1", wGrpFormula
    'GRP_REP.SetFormula mRep, "wGrp1Desc", wGrpDescFormula
    GRP_REP.SetFormula mRep, "wDet", wDetFormula
    GRP_REP.SetFormula mRep, "wDetDesc", wDetDescFormula
    
    If UCase(ADC("UoInDet")) = "" Then _
    GRP_REP.SetFormula mRep, "wDet", wGrpFormula

    If wRsInHd.RecCount > 0 And wRsInHd!InNo <> 0 Then
      SetRepText mRep, "TxtAmtInWrds", "Total " + FnNumToWords(Format(wRsInHd!InFOBRs, "######0.00"), ctIndCurncy)
    End If
    
    mRep.TxtDocTyp.SetText ADC("wDocTyp")
    
    Rep4.FldInCmPayTerms.Suppress = IIF(ADC("UoYN7") = "Y", False, True)
    
'    If adc("UoYN7") = "Y" Then
'      Rep4.FldInCmPayTerms.Suppress = False
'    Else
'      Rep4.FldInCmPayTerms.Suppress = True
'    End If
                      
    ' **** Manali 3.03 - 25/06/08 - Pd Mod - 'L' added
    Set wRsMetUtlsd = moCn.OpenRes("Select distinct CmName from Txndz M " + _
                  "join Txndz B on " + IIF(gs_Partition = ctCurrPrtn, " B.TzPrtKey=M.TzPrtKey and ", "") + " B.TzCoCd= M.TzCoCd and B.TzYy= M.TzRefYy " + _
                  "and B.TzKey= M.TzRefKey join Txn on " + IIF(gs_Partition = ctCurrPrtn, " TPrtKey=B.TzPrtKey and ", "") + " TCoCd= B.TzCoCd and TTc= B.TzTc and " + _
                  "TYy= B.TzYy and TChr= B.TzChr and TNo= B.TzNo " + _
                  "join RmMst on RmCd= M.TzRmZCd " + IIF(gs_Partition = ctCurrPrtn, " And RmPrtKey=M.TzPrtKey ", "") + _
                  "join CustMst on CmCtg= 'S' and CmCd= TSuppCd " + _
                  "join InvHd Ih on Ih.InCoCd= M.TzCoCd and Ih.InTc= M.TzTc and Ih.InYy= M.TzYy " + _
                  "and Ih.InChr= M.TzChr and Ih.InNo= M.TzNo " + IIF(gs_Partition = ctCurrPrtn, " And Ih.InPrtKey=M.TzPrtKey ", "") + wCnd + IIF(gs_Partition = ctCurrPrtn, " And M.TzPrtKey='" + ctCurrPrtn + "' ", "") + " and RmCtg in ('G', 'P', 'S', 'L')")
                  
    Do While Not (wRsMetUtlsd.EOF Or wRsMetUtlsd.BOF)
      wMetUtlsdStr = wMetUtlsdStr + IIF(wMetUtlsdStr = "", "", ", ") + wRsMetUtlsd!CmName
      wRsMetUtlsd.MoveNext
    Loop
    
    SetRepText mRep, "TxtMetUtlsd", "METAL UTILISED IS METAL IMPORTED/ PURCHASED FROM: " + wMetUtlsdStr
    
    mRep.TxtCurCd1.SetText ws_CurCd
    mRep.TxtCurSymbol1.SetText ws_CurCd
    mRep.TxtOneCurCd.SetText "ONE " + ws_CurCd + " ="
  '*** Jay 2.14 P2
  
  '*** Jay 3.3(FG)
  ' ***** Manali Trading Module
  '****** (Jen 3.01) added RepGSPFORMMulti ******
  Case Is = UCase("RepGSPFORM"), UCase("RepGSPFORMMulti"), UCase("RepFgGsp"), UCase("RepJtGsp")
  '*** Jay 3.3(FG)
    GRP_REP.SetFormula mRep, "wGrp1", wGrpFormula
    GRP_REP.SetFormula mRep, "wGrp1Desc", wGrpDescFormula
    If wDmUomFlag = False Then
      mRep.FldDmUom.Suppress = True
      mRep.SecSummDmUom.Suppress = True
      mRep.SecTotQty.Suppress = False
    Else
      mRep.FldDmUom.Suppress = False
      mRep.SecSummDmUom.Suppress = False
      mRep.SecTotQty.Suppress = True
    End If

    '*** Jay 3.3(FG)
    ' ***** Manali Trading Module
    If UCase(ADC.MenuCd) <> UCase("RepFgGsp") And UCase(ADC.MenuCd) <> UCase("RepJtGsp") Then
    '*** Jay 3.3(FG)
      '*** (Jen 2.13)
      '****** Sachin 3.02 - Id fields in Joins
      wSubRepStr = "Select '1D' as qTag,Max(OdOmCmCd) as OmCmCd, DmCtg, max(DmSalCtg) as DmSalCtg, (Select PDesc From Param where PTyp= 'DMCTG'" + _
                " and PMCd=DmCtg  and PSCd='') as  DmCtgDesc, " + _
                " (Select PDesc from Param where PTyp= 'SALCTG' and PMCd= DmCtg and " + _
                " PSCd= max(DmSalCtg)) as DmSalCtgDesc, " + _
                "max(DmUom) as DmUom, Max(" + Ws_IdKt + ") as IdKt," + _
                "Max(" + ws_IdVaCtg + ") as IdVaCtg, Max(IdGldLs) as IdGldLs, Max(OdDmCol) as OdDmCol," + qMainMet + ", sum(" + ws_IdQty + ") as qIdQty " + _
                "From InvDsg Join DsgMst  " + _
                "On DmIdNo=IdDmIdNo and DmTcTyp= 'DM' and IdDmCd= DmCd and DmSz= '' " + IIF(gs_Partition = ctCurrPrtn, " And DmPrtKey=IdPrtKey ", "") + _
                "join InvHd Ih on Ih.InIdNo=IdInIdNo and Ih.InCoCd= IdCoCd and Ih.InTc= IdTc and Ih.InYy= IdYy and Ih.InChr= IdChr and Ih.InNo= IdNo " + IIF(gs_Partition = ctCurrPrtn, " And Ih.InPrtKey=IdPrtKey ", "") + _
                "join OrdDsg on OdCoCd= IdCoCd and OdTc= IdExpOdTc and OdYy= IdExpOdYy and OdChr= IdExpOdChr " + _
                "and OdNo= IdExpOdNo and OdSr= IdExpOdSr " + IIF(gs_Partition = ctCurrPrtn, " And OdPrtKey=IdPrtKey ", "") + wOrJoin + _
                " " + wCnd + IIF(gs_Partition = ctCurrPrtn, " And IdPrtKey='" + ctCurrPrtn + "' ", "") + " " + IIF(wFlagLocalExp = "Y", " ", " and 1= 2 ") + _
                " group by " + wSubRepGrpFormula + "DmCtg, DmUom "
    '*** Jay 3.3(FG)
    Else
      wSubRepStr = "Select '1D' as qTag,Max(OdOmCmCd) as OmCmCd, DmCtg, max(DmSalCtg) as DmSalCtg, (Select PDesc From Param where PTyp= 'DMCTG'" + _
          " and PMCd=DmCtg  and PSCd='') as  DmCtgDesc, " + _
          " (Select PDesc from Param where PTyp= 'SALCTG' and PMCd= DmCtg and " + _
          " PSCd= max(DmSalCtg)) as DmSalCtgDesc, " + _
          "max(DmUom) as DmUom, Max(" + Ws_IdKt + ") as IdKt," + _
          "Max(" + ws_IdVaCtg + ") as IdVaCtg, 0 as IdGldLs, Max(OdDmCol) as OdDmCol," + qMainMet + ", sum(" + ws_IdQty + ") as qIdQty " + _
          "From OrdDsg Join DsgMst  " + _
          "On DmIdNo=OdDmIdNo and DmTcTyp= 'DM' and OdDmCd= DmCd and DmSz= '' " + IIF(gs_Partition = ctCurrPrtn, " And DmPrtKey=OdPrtKey ", "") + _
          "join InvHd Ih on Ih.InCoCd= OdCoCd and Ih.InTc= OdTc and Ih.InYy= OdYy and Ih.InChr= OdChr and Ih.InNo= OdNo " + IIF(gs_Partition = ctCurrPrtn, " And Ih.InPrtKey=OdPrtKey ", "") + _
          " join OrdRm on " + IIF(gs_Partition = ctCurrPrtn, " OrPrtKey=OdPrtKey and ", "") + " OrOdIdNo=OdIdNo and OrCoCd= OdCoCd and OrTc= OdTc and OrYy= OdYy " + _
          " and OrChr= OdChr and OrNo= OdNo and OrSr= OdSr and OrMainMet= 'Y' " + _
          " " + wCnd + IIF(gs_Partition = ctCurrPrtn, " And OdPrtKey='" + ctCurrPrtn + "' ", "") + " " + IIF(wFlagLocalExp = "Y", " ", " and 1= 2 ") + _
          " group by " + wSubRepGrpFormula + "DmCtg, DmUom "
    End If
    '*** Jay 3.3(FG)
                        
    Call moCn.RepRes(mRep.Subreport1.OpenSubreport, wSubRepStr)
    SetSubFormula mRep.Subreport1, "wGrp", wGrpFormula
    '*** (Jen 2.13)
    
    If wFlagLocalExp = "N" Then
      mRep.TxtGldWt.Suppress = True
      mRep.FldwTotGldWt.Suppress = True
      mRep.TxtGldGms.Suppress = True
      mRep.TxtPlWt.Suppress = True
      mRep.FldwTotPlWt.Suppress = True
      mRep.TxtPlGms.Suppress = True
      ' ****** Sachin 2.12 - 05/12/05 - SJM
      mRep.TxtSlWt.Suppress = True
      mRep.FldwTotSlWt.Suppress = True
      mRep.TxtSlGms.Suppress = True
      ' ****** Sachin 2.12 - 05/12/05 - SJM
      
      ' 3.11.0
      mRep.TxtPdWt.Suppress = True
      mRep.FldwTotPdWt.Suppress = True
      mRep.TxtPdGms.Suppress = True
      
      ' Zubin 212
      If UCase(ADC("UoYn4")) = "N" Then
        mRep.TxtNetWt.Suppress = True
        mRep.SecNetWt.Suppress = True
      End If
      ' Zubin 212
      
      '*** (Jen 2.13)
      mRep.FldInCmCpName.Suppress = True
      mRep.FldInHMarks.Suppress = True
      mRep.TxtOneTinBox.Suppress = True
      mRep.TxtJewelDesc.Suppress = True
      mRep.SecInvDsgDescSubRep.Suppress = True
      'mRep.TxtGrp1GrsWt.Suppress = True: mRep.FldGrp1GrsWt.Suppress = True
      'mRep.TxtGrp1NetWt.Suppress = True: mRep.FldGrp1NetWt.Suppress = True
      '*** (Jen 2.13)
    ElseIf wFlagLocalExp = "Y" Then
      '****Geeta***Emr208
      ' Zubin 212 (Sales Ctg and qDmSalCtgDesc added)
      '*** (Bef 2.13)
      'wSubRepStr = "Select '1D' as qTag, DmCtg, max(DmSalCtg) as DmSalCtg, (Select PDesc From Param where PTyp='DMCTG'" + _
      '            " and PMCd=DmCtg  and PSCd='') as  DmCtgDesc, " + _
      '            " (Select PDesc from Param where PTyp= 'SALCTG' and PMCd= DmCtg and " + _
      '            " PSCd= max(DmSalCtg)) as DmSalCtgDesc, " + _
      '            "max(DmUom) as DmUom, Max(IdKt) as IdKt," + _
      '            "Max(IdVaCtg) as IdVaCtg, Max(IdGldLs) as IdGldLs, Max(OdDmCol) as OdDmCol," + qMainMet + ", sum(IdQty) as qIdQty " + _
      '            "From InvDsg Join DsgMst  " + _
      '            "On DmTcTyp='DM' and IdDmCd=DmCd and DmSz='' " + _
      '            "join OrdDsg on OdCoCd= IdCoCd and OdTc= IdExpOdTc and OdYy= IdExpOdYy " + _
      '            "and OdChr= IdExpOdChr and OdNo= IdExpOdNo and OdSr= IdExpOdSr " + wOrJoin + _
      '            "where IdCoCd= '" + adc("UoCoCdFr") + "' and IdTc= '" + adc("UoInTcFr") + "' and " + _
      '            "IdYy= '" + adc("UoInYyFr") + "' and IdChr= '" + adc("UoInChrFr") + "' and " + _
      '            "IdNo= " + CStr(adc("UoInNoFr")) + "  group by " + wSubRepGrpFormula + "DmCtg, DmUom "
      '
      'Call moCn.RepRes(mRep.Subreport1.OpenSubreport, wSubRepStr)
      'SetSubFormula mRep.Subreport1, "wGrp", wGrpFormula
      '*** (Bef 2.13)
      '**********Geeta*******Emr208***
    
      '*** (Jen 2.13)
      mRep.TxtAsPerSr.Suppress = True
      'mRep.FldGrp1Qty.Suppress = True
      'mRep.FldDmUom.Suppress = True
      'mRep.TxtSrNo.Suppress = True
      '*** (Jen 2.13)
    End If
      
  Case Is = UCase("RepPckLstInvBag")
'    GRP_REP.SetFormula mRep, "wGrp1", "if {rdo.qTag}='1D' Then " + wGrpFormula
'    GRP_REP.SetFormula mRep, "wGrp1Desc", wGrpDescFormula
'    GRP_REP.SetFormula mRep, "wDet", wDetFormula
'    GRP_REP.SetFormula mRep, "wDetDesc", wDetDescFormula
'    GRP_REP.SetFormula mRep, "wDetDesc2", wDetDesc2Formula
'
'    If UCase(adc("UoInDet")) = "" Then
'        GRP_REP.SetFormula mRep, "wDet", wGrpFormula
'    End If
'
'    If UCase(adc("UoYN")) = "Y" Then
'      mRep.PicSec.Suppress = True: mRep.PicSec1.Suppress = True
'      If UCase(adc("UoInDet")) = "DMCD" Or UCase(adc("UoInDet")) = "INVSR" Or UCase(adc("UoInDet")) = "INVSR/SR" Or _
'         UCase(adc("UoInDet")) = "DMCD/SR" Then If UCase(adc("UoYN9")) = "Y" Then mRep.PicSec1.Suppress = False Else mRep.PicSec.Suppress = False
'    Else
'        mRep.PicSec.Suppress = True: mRep.PicSec1.Suppress = True
'    End If
'    wSubRepStr = "Select '1D' as qTag, DmCtg, max(DmSalCtg) as DmSalCtg, (Select PDesc From Param where PTyp= 'DMCTG'" + _
'              " and PMCd= DmCtg and PSCd= '') as DmCtgDesc, " + _
'              " (Select PDesc from Param where PTyp= 'SALCTG' and PMCd= DmCtg and " + _
'              " PSCd= max(DmSalCtg)) as DmSalCtgDesc, max(DmUom) as DmUom,Max(" + Ws_IdKt + ") as IdKt," + _
'              "Max(" + ws_IdVaCtg + ") as IdVaCtg, Max(IdGldLs) as IdGldLs, Max(OdDmCol) as OdDmCol," + qMainMet + ", sum(" + ws_IdQty + ") as qIdQty " + _
'              " From InvDsg Join DsgMst  " + _
'              " On DmIdNo=IdDmIdNo and DmTcTyp= 'DM' and IdDmCd= DmCd and DmSz= '' " + IIF(gs_Partition = ctCurrPrtn, " And IdPrtKey=DmPrtKey", "") + _
'              " join InvHd Ih on Ih.InIdNo=IdInIdNo and Ih.InCoCd= IdCoCd and Ih.InTc= IdTc and Ih.InYy= IdYy and Ih.InChr= IdChr and Ih.InNo= IdNo " + IIF(gs_Partition = ctCurrPrtn, " And Ih.InPrtKey=IdPrtKey ", "") + _
'              " join OrdDsg on OdCoCd= IdCoCd and OdTc= IdExpOdTc and OdYy= IdExpOdYy " + _
'              " and OdChr= IdExpOdChr and OdNo= IdExpOdNo and OdSr= IdExpOdSr " + IIF(gs_Partition = ctCurrPrtn, " And OdPrtKey=IdPrtKey", "") + wOrJoin + _
'              " " + wCnd + IIF(gs_Partition = ctCurrPrtn, " And IdPrtKey='" + ctCurrPrtn + "' ", "") + " group by " + wSubRepGrpFormula + "DmCtg, DmUom "
'    Call moCn.RepRes(mRep.Subreport3.OpenSubreport, wSubRepStr + ctMaxDopOpt)
'    SetSubFormula mRep.Subreport3, "wGrp", wGrpFormula
'    Dim ws_PckLstSubRep As String
'    If adc("UoYN9") = "Y" Then
'      ws_PckLstSubRep = "Select '1D' as qTag, IdCoCd, IdTc, IdYy, IdChr, IdNo, IdSr, IfFdBYy, IfFdBChr, IfFdBNo, max(IfSrNo) as IfSrno, Sum(IfQty) as IfQty, Sum(IfGrsWt) as IfGrsWt, " + _
'                        "max(IdKt) as IdKt, max(DmCtg) as DmCtg, max(DmSalCtg) as DmSalCtg, max(DmCd) as IdDmCd, max(IdSfx) as IdSfx, max(IdDmSz) as IdDmSz, max(DmUom) as DmUom,  max(OmPONo) as OmPONo, max(OdPoNo) as OdPoNo, " + _
'                        "max(OdDmCol) as OdDmCol, max(IdVaCtg) as IdVaCtg, max(IdGldLs) as IdGldLs, " + qMainMet + _
'                        " from InvFgd  join InvDsg on IdInIdNo=IfInIdNo and IdCoCd= IfCoCd and IdTc= IfTc and IdYy= IfYy and IdChr= IfChr " + _
'                        " and IdNo= IfNo and IdSr= IfSr " + IIF(gs_Partition = ctCurrPrtn, " And IdPrtKey=IfPrtKey ", "") + _
'                        " join DsgMst on DmIdNo=IdDmIdNo and DmTcTyp= 'DM' and DmCd = IdDmCd and DmSz= '' " + IIF(gs_Partition = ctCurrPrtn, " And DmPrtKey=IdPrtKey ", "") + _
'                        " join OrdMst on OmCoCd= IdCoCd and OmTc= IdExpOdTc and OmYy= IdExpOdYy " + _
'                        " and OmChr= IdExpOdChr and OmNo= IdExpOdNo " + IIF(gs_Partition = ctCurrPrtn, " And OmPrtKey=IdPrtKey ", "") + _
'                        " join OrdDsg on OdCoCd= IdCoCd and OdTc= IdExpOdTc and OdYy= IdExpOdYy " + _
'                        " and OdChr= IdExpOdChr and OdNo= IdExpOdNo and OdSr= IdExpOdSr " + IIF(gs_Partition = ctCurrPrtn, " And OdPrtKey=IdPrtKey ", "") + wOrJoin + _
'                        " join InvHd Ih on Ih.InIdNo=IdInIdNo and Ih.InCoCd= IdCoCd and Ih.InTc = IdTc and Ih.InYy= IdYy and Ih.InChr= IdChr and Ih.InNo= IdNo " + IIF(gs_Partition = ctCurrPrtn, " And Ih.InPrtKey=IdPrtKey ", "") + _
'                        wCnd + IIF(gs_Partition = ctCurrPrtn, " And IfPrtKey='" + ctCurrPrtn + "' ", "") + " Group by IdCoCd, IdTc, IdYy, IdChr, IdNo, IdSr, IfFdBYy, IfFdBChr, IfFdBNo Order By IdSr, IfFdBYy, IfFdBChr, IfFdBNo "
'      Call moCn.RepRes(mRep.SubRepInvBag.OpenSubreport, ws_PckLstSubRep + ctMaxDopOpt)
'      SetSubFormula mRep.SubRepInvBag, "wSubDet", wDetFormula
'      SetSubFormula mRep.SubRepInvBag, "wSubGrp", wGrpFormula
'      If UCase(adc("UoInDet")) = "" Then
'          SetSubFormula mRep.SubRepInvBag, "wSubDet", wGrpFormula
'      End If
'    Else
'      mRep.SecInvBag.Suppress = True
'    End If
'    If wFlagLocalExp = "N" Then mRep.SecDescSubRep.Suppress = True
'    If UCase(adc("UoYN1")) = "Y" Then
'      mRep.RemSec.Suppress = True: mRep.RemSec1.Suppress = True
'      If UCase(adc("UoInDet")) = "DMCD" Or UCase(adc("UoInDet")) = "INVSR" Or UCase(adc("UoInDet")) = "INVSR/SR" Or _
'         UCase(adc("UoInDet")) = "DMCD/SR" Then If UCase(adc("UoYN9")) = "Y" Then mRep.RemSec1.Suppress = False Else mRep.RemSec.Suppress = False
'    Else
'        mRep.RemSec.Suppress = True: mRep.RemSec1.Suppress = True
'    End If
'
'    If adc("UoYn9") = "Y" Then
'      GRP_REP.SetFormula mRep, "wShowBag", "'Y'"
'    Else
'      GRP_REP.SetFormula mRep, "wShowBag", "'N'"
'    End If
  End Select
  '*** (Jen 2.12)
  
  '*** (Bef 2.14 P2 12/09/07)
  '*** (24/02/06) For Kama (Jen 2.12)
  'If UCase(adc.MenuCd) = UCase("RepInvLoose") Or UCase(adc.MenuCd) = UCase("RepShpBillLoose") Or _
  '   UCase(adc.MenuCd) = UCase("RepGrFormLoose") Then
  '*** (24/02/06) For Kama (Jen 2.12)
  '*** (Bef 2.14 P2 12/09/07)
  '*** Jay 2.14 P2 12/09/07
  If UCase(ADC.MenuCd) = UCase("RepInvLoose") Or UCase(ADC.MenuCd) = UCase("RepShpBillLoose") Or _
   UCase(ADC.MenuCd) = UCase("RepGrFormLoose") Or UCase(ADC.MenuCd) = UCase("RepGrFormNewLoose") Then
  '*** Jay 2.14 P2 12/09/07
    'RMGRP'-'INV'
    Select Case UCase(ADC("UoRmGrp"))
    Case Is = "N"
      ws_RmGrpOpt = " "
      ws_RmCtg = "'' "
      ws_RmSCtg = "'' "
      ws_RmCd = "'' "
      ws_RmRt = "'' "
      Rep.Subreport2_hSubRep2RmCtg.Suppress = True: Rep.Subreport2_FldSubRep2RmCtg.Suppress = True
      Rep.Subreport2_hSubRep2RmsCtg.Suppress = True: Rep.Subreport2_FldSubRep2RmSCtg.Suppress = True
      Rep.Subreport2_hSubRep2RmCd.Suppress = True: Rep.Subreport2_FldSubRep2RmCd.Suppress = True
      Rep.Subreport2_hSubRep2RmRt.Suppress = True: Rep.Subreport2_FldSubRep2RmRt.Suppress = True
      '*** (Bef 2.13) Jen
      'Rep4.Subreport3_hSubRep3RmCtg.Suppress = True: Rep4.Subreport3_FldSubRep3RmCtg.Suppress = True
      'Rep4.Subreport3_hSubRep3RmSCtg.Suppress = True: Rep4.Subreport3_FldSubRep3RmSCtg.Suppress = True
      'Rep4.Subreport3_hSubRep3RmCd.Suppress = True: Rep4.Subreport3_FldSubRep3RmCd.Suppress = True
      'Rep4.Subreport3_hSubRep3RmRt.Suppress = True: Rep4.Subreport3_FldSubRep3RmRt.Suppress = True
      '*** (Bef 2.13) Jen
    Case Is = "RMCDRT"
      ws_RmGrpOpt = ", RmCtg, RmSCtg, RmCd, " + wRmRtOpt + " "
      ws_RmCtg = "max(RmCtg) + max((Case When PSCd is Null Then '' Else '-' + IsNull(PSCd, '') End )) "
      ws_RmSCtg = "max(RmSCtg) "
      ws_RmCd = "max(RmCd) "
      ws_RmRt = wRmRtOpt
    Case Is = "RMCD"
      ws_RmGrpOpt = ", RmCtg, RmSCtg, RmCd "
      ws_RmCtg = "max(RmCtg) + max((Case When PSCd is Null Then '' Else '-' + IsNull(PSCd, '') End )) "
      ws_RmSCtg = "max(RmSCtg) "
      ws_RmCd = "max(RmCd) "
      ws_RmRt = "0 "
      Rep.Subreport2_hSubRep2RmRt.Suppress = True: Rep.Subreport2_FldSubRep2RmRt.Suppress = True
      '*** (Bef 2.13) Jen
      'Rep4.Subreport3_hSubRep3RmRt.Suppress = True: Rep4.Subreport3_FldSubRep3RmRt.Suppress = True
      '*** (Bef 2.13) Jen
    Case Is = "RMSCTG"
      ws_RmGrpOpt = ", RmCtg, RmSCtg "
      ws_RmCtg = "max(RmCtg) + max((Case When PSCd is Null Then '' Else '-' + IsNull(PSCd, '') End )) "
      ws_RmSCtg = "max(RmSCtg) "
      ws_RmCd = "'' "
      ws_RmRt = "0 "
      Rep.Subreport2_hSubRep2RmCd.Suppress = True: Rep.Subreport2_FldSubRep2RmCd.Suppress = True
      Rep.Subreport2_hSubRep2RmRt.Suppress = True: Rep.Subreport2_FldSubRep2RmRt.Suppress = True
      '*** (Bef 2.13) Jen
      'Rep4.Subreport3_hSubRep3RmCd.Suppress = True: Rep4.Subreport3_FldSubRep3RmCd.Suppress = True
      'Rep4.Subreport3_hSubRep3RmRt.Suppress = True: Rep4.Subreport3_FldSubRep3RmRt.Suppress = True
      '*** (Bef 2.13) Jen
    Case Is = "RMCTG"
      ws_RmGrpOpt = ", RmCtg"
      ws_RmCtg = "max(RmCtg) + max((Case When PSCd is Null Then '' Else '-' + IsNull(PSCd, '') End )) "
      ws_RmSCtg = "'' "
      ws_RmCd = "'' "
      ws_RmRt = "0 "
      Rep.Subreport2_hSubRep2RmsCtg.Suppress = True: Rep.Subreport2_FldSubRep2RmSCtg.Suppress = True
      Rep.Subreport2_hSubRep2RmCd.Suppress = True: Rep.Subreport2_FldSubRep2RmCd.Suppress = True
      Rep.Subreport2_hSubRep2RmRt.Suppress = True: Rep.Subreport2_FldSubRep2RmRt.Suppress = True
      '*** (Bef 2.13) Jen
      'Rep4.Subreport3_hSubRep3RmSCtg.Suppress = True: Rep4.Subreport3_FldSubRep3RmSCtg.Suppress = True
      'Rep4.Subreport3_hSubRep3RmCd.Suppress = True: Rep4.Subreport3_FldSubRep3RmCd.Suppress = True
      'Rep4.Subreport3_hSubRep3RmRt.Suppress = True: Rep4.Subreport3_FldSubRep3RmRt.Suppress = True
      '*** (Bef 2.13) Jen
    End Select
    If wRmWtOpt = "" Then wRmWtOpt = "IrRmZWt"
    If wRmValOpt = "" Then wRmValOpt = "IrRmZVal"
    
    '*** (Bef 2.13)
    'wrepcnd = adc.RepCond
    'wCnd = IIF(wrepcnd <> "", " where " + wrepcnd, "")
    '*** (Bef 2.13)
    
    '(Bef 2.13)  sum(" + wRmValOpt + ")* max(InCnvRt) as qValRs
    '****** (Jen 2.13) "sum(" + wRmValOpt + ")* max(case when InRsMulDiv= 'D' and convert (decimal (16, 4), InCnvRt)> 0 then 1/ InCnvRt else InCnvRt end) as qValRs " ******
    '****** Sachin 3.02 - Id fields in Joins
    'JewelSmith.2 (','+PDesc225+',') Like ('%,'+RmSCtg+',%') changed as PSCd =RmStnCls
    ws_RmSubRep = "Select " + ws_RmCtg + " as qRmCtg, " + ws_RmSCtg + " as RmSCtg, " + _
          ws_RmCd + " as RmCd, " + ws_RmRt + " as qRmRt, sum(IrRmQty) as IrRmQty, " + _
          "sum(case when RmCtg in ('C', 'D') then " + wRmWtOpt + " else 0 end) as qWtInCts, " + _
          "sum(case when RmCtg in ('C', 'D') then " + wRmWtOpt + "/5 else " + wRmWtOpt + " end) as qWtInGms, " + _
          "sum(" + wRmValOpt + ") as qVal, " + _
          "sum(" + wRmValOpt + ")* max(case when InRsMulDiv= 'D' and convert (decimal (16, 4), InCnvRt)> 0 then 1/ InCnvRt else InCnvRt end) as qValRs " + _
          " from InvRm join RmMst on RmCd = IrRmCd  " + IIF(gs_Partition = ctCurrPrtn, " And RmPrtKey=IrPrtKey ", "") + _
          "join InvHd Ih on InCoCd= IrCoCd and InTc = IrTc and InYy= IrYy and InChr= IrChr and " + _
          "     InNo= IrNo " + IIF(gs_Partition = ctCurrPrtn, " And InPrtKey=IrPrtKey ", "") + _
          "Left Outer Join Param On PTyp= 'STNCLS' and PMCd= RMCTG and PSCd =RmStnCls " + _
          " " + wCnd + IIF(gs_Partition = ctCurrPrtn, " And IrPrtKey='" + ctCurrPrtn + "' ", "") + ws_ShowRmSubRepCnd + _
          IIF(wFlagLocalExp = "Y" And UCase(ADC.MenuCd) = UCase("RepInv"), " and RmCtg in ('D', 'C', 'X') ", " ") + _
          "Group by IrTc, IrYy, IrChr, IrNo " + ws_RmGrpOpt + _
          " Order By IrTc, IrYy, IrChr, IrNo " + ws_RmGrpOpt
    
          'Group By IrTc, IrYy, IrChr, IrNo, IrSr, RmCtg, PSCd"
    
    '*** (24/02/06) For Kama (Jen 2.12)
    If UCase(ADC.MenuCd) = UCase("RepInvLoose") Then
    '*** (24/02/06) For Kama (Jen 2.12)
        
      '*** (Bef 2.13) Jen
      ''*** (24/02/06) For Kama (Jen 2.12)
      'If wFlagLocalExp = "Y" Then
      '  Call moCn.RepRes(mRep.Subreport3.OpenSubreport, ws_RmSubRep)
      'ElseIf wFlagLocalExp = "N" Then
      '*** (Bef 2.13) Jen
        Call moCn.RepRes(mRep.Subreport2.OpenSubreport, ws_RmSubRep)
      
      '*** (Bef 2.13) Jen
      'End If
      ''*** (24/02/06) For Kama (Jen 2.12)
      '*** (Bef 2.13) Jen
      
      'Diagold.3 - IGST must be shown as 0% in DTA Invoice
      If wFlagLocalExp = "N" Then
      mRep.TxtIGSTRate.Suppress = True
      End If
      
      mRep.FldTotPcs.Suppress = True: mRep.FldTotPrs.Suppress = True: mRep.FldTotSet.Suppress = True
      mRep.TxtPcs.Suppress = True: mRep.TxtPrs.Suppress = True: mRep.TxtSet.Suppress = True
      mRep.TxtGrWt.Suppress = True: mRep.TxtQty.Suppress = True: mRep.TxtRt.Suppress = True
      mRep.TxtAmtUSD.Suppress = True
      mRep.LineGrWt.Suppress = True: mRep.LineQty.Suppress = True: mRep.LineRt.Suppress = True
      Rep.LineQtyUp.Suppress = True: Rep.LineQtyDn.Suppress = True   ' ***** Manali 3.6.0 - 09/09/09 - Cost Details in DTA Invoice


      mRep.PHDET.Suppress = True: mRep.DETSEC.Suppress = True: mRep.PICSEC.Suppress = True
      mRep.RemSec.Suppress = True: mRep.G1F.Suppress = True: mRep.LineSec.Suppress = True
      
      '*** (Jen 2.13)
      mRep.TxtDCWt.Suppress = True: mRep.FldTotDCWt.Suppress = True
      mRep.TxtCSWt.Suppress = True: mRep.FldTotCSWt.Suppress = True     '4.1.2
      mRep.FldTotGrWt.Suppress = True
      '*** (Jen 2.13)
        
    ElseIf UCase(ADC.MenuCd) = UCase("RepShpBillLoose") Then
      
      '*** (Jen 2.13)
      Call moCn.RepRes(mRep.Subreport1.OpenSubreport, ws_RmSubRep)
      '*** (Jen 2.13)
      
      '*** (24/02/06) For Kama (Jen 2.12)
      If wFlagLocalExp = "Y" Then
        mRep.SecDescSubRep.Suppress = True
      '*** (Bef 2.13) Jen
      '  Call moCn.RepRes(mRep.Subreport2.OpenSubreport, ws_RmSubRep)
      '  mRep.G1HDesc.Suppress = True
      'ElseIf wFlagLocalExp = "N" Then
      '  Call moCn.RepRes(mRep.Subreport1.OpenSubreport, ws_RmSubRep)
      '*** (Bef 2.13) Jen
      End If
      '*** (24/02/06) For Kama (Jen 2.12)
      
      mRep.PHDET.Suppress = True: mRep.DETSEC.Suppress = True
      mRep.G1F.Suppress = True: mRep.TxtQty.Suppress = True
      mRep.LineQty.Suppress = True: mRep.TxtAmtRs.Suppress = True
      mRep.LineAmtRs.Suppress = True: mRep.FldTotQty.Suppress = True
    
    '****** Jay 2.14 P2 12/09/07 added RepGrFormNewLoose
    ElseIf UCase(ADC.MenuCd) = UCase("RepGrFormLoose") Or UCase(ADC.MenuCd) = UCase("RepGrFormNewLoose") Then
      '*** (Jen 2.13)
      Call moCn.RepRes(mRep.Subreport1.OpenSubreport, ws_RmSubRep)
      '*** (Jen 2.13)
      
      '*** (24/02/06) For Kama (Jen 2.12)
      mRep.PHDET.Suppress = True: mRep.DETSEC.Suppress = True
      mRep.G1F.Suppress = True: mRep.TxtQty.Suppress = True
      mRep.LineQty.Suppress = True: mRep.LineAmtRs.Suppress = True
      mRep.TxtAmtRs.Suppress = True: mRep.FldTotQty.Suppress = True
      If wFlagLocalExp = "Y" Then
        '*** (Jen 2.13)
        mRep.SecDescSubRep.Suppress = True
        'mRep.TxtAmtRs.Suppress = True mRep.LineQty.Suppress = True:
        'mRep.LineAmtRs.Suppress = True: mRep.FldTotQty.Suppress = True
        '*** (Jen 2.13)
        
        '*** (Bef 2.13) Jen
        'mRep.PhDet.Suppress = True: mRep.DETSEC.Suppress = True
        'mRep.G1F.Suppress = True: mRep.TxtQty.Suppress = True
        'mRep.TxtAmtRs.Suppress = True: mRep.FldDmUom.Suppress = True
        'mRep.FldQty.Suppress = True: mRep.FldAmtRs.Suppress = True
        'mRep.LineQty.Suppress = True: mRep.LineAmtRs.Suppress = True
        'mRep.FldTotQty.Suppress = True
        
        'Call moCn.RepRes(mRep.Subreport2.OpenSubreport, ws_RmSubRep)
        'mRep.G1HDesc.Suppress = True
        '*** (Bef 2.13) Jen
      ElseIf wFlagLocalExp = "N" Then
      
        'mRep.PHDET.Suppress = True: mRep.DETSEC.Suppress = True
        'mRep.G1F.Suppress = True: mRep.TxtQty.Suppress = True
        mRep.FldDmUom.Suppress = True  ' mRep.TxtAmtRs.Suppress = True:
        mRep.FldQty.Suppress = True: mRep.FldAmtRs.Suppress = True
        'mRep.LineQty.Suppress = True: mRep.LineAmtRs.Suppress = True
        'mRep.FldTotQty.Suppress = True
      
        '*** (Bef 2.13) Jen
        'Call moCn.RepRes(mRep.Subreport1.OpenSubreport, ws_RmSubRep)
        '*** (Bef 2.13) Jen
      End If
      '*** (24/02/06) For Kama (Jen 2.12)
    End If
  
  '*** (Bef 2.13)
  '*** (24/02/06) For Kama (Jen 2.12)
  'ElseIf UCase(adc.MenuCd) = UCase("RepInv") Or UCase(adc.MenuCd) = UCase("RepGRFormFrnt") Or _
  '   UCase(adc.MenuCd) = UCase("RepShpBill") Then
  '*** (24/02/06) For Kama (Jen 2.12)
  '*** (Bef 2.13)
  
  '*** Jay 3.3(FG)
  ' ***** Manali Trading Module
  '*** (Jen 2.13)
  ElseIf UCase(ADC.MenuCd) = UCase("RepInv") Or UCase(ADC.MenuCd) = UCase("RepGRFormFrnt") Or _
         UCase(ADC.MenuCd) = UCase("RepShpBill") Or UCase(ADC.MenuCd) = UCase("RepShpBillMulti") Or _
         UCase(ADC.MenuCd) = UCase("RepGrFormFrntMulti") Or UCase(ADC.MenuCd) = UCase("RepFgInv") Or _
         UCase(ADC.MenuCd) = UCase("RepFgShpBill") Or UCase(ADC.MenuCd) = UCase("RepFgGrFrnt") Or _
         UCase(ADC.MenuCd) = UCase("RepJtInv") Or UCase(ADC.MenuCd) = UCase("RepJtShpBill") Or _
         UCase(ADC.MenuCd) = UCase("RepJtGrFrnt") Then
  '*** (Jen 2.13)
  '*** Jay 3.3(FG)
  
    '*** Jay 3.3(FG)
    ' ***** Manali Trading Module
    If UCase(ADC.MenuCd) <> UCase("RepFgInv") And UCase(ADC.MenuCd) <> UCase("RepFgShpBill") And _
       UCase(ADC.MenuCd) <> UCase("RepFgGrFrnt") And _
       UCase(ADC.MenuCd) <> UCase("RepJtInv") And UCase(ADC.MenuCd) <> UCase("RepJtShpBill") And _
       UCase(ADC.MenuCd) <> UCase("RepJtGrFrnt") Then
      ws_RmSubRep = " Select RmCtg+ '-'+ PSCd as qRmCtg, RmSCtg, RmCd, IrRmIRt as qRmRt, " + _
                  "IrRmQty, IrRmIWt As qWtInCts, IrRmIWt As qWtInGms, IrRmIVal As qVal, " + _
                  "IrRmIVal As qValRs from RmMst, InvRm, Param where 1 = 2 "
    '*** Jay 3.3(FG)
    Else
      ws_RmSubRep = " Select RmCtg+ '-'+ PSCd as qRmCtg, RmSCtg, RmCd, OrSalRt as qRmRt, " + _
                  "OrQty as IrRmQty, OrWt As qWtInCts, OrWt As qWtInGms, OrSalVal As qVal, " + _
                  "OrSalVal As qValRs from RmMst, OrdRm, Param where 1 = 2 "
    End If
    '*** Jay 3.3(FG)
  
  
    '*** Jay 3.3(FG)
    ' ***** Manali Trading Module
    If UCase(ADC.MenuCd) = UCase("RepInv") Or UCase(ADC.MenuCd) = UCase("RepFgInv") Or UCase(ADC.MenuCd) = UCase("RepJtInv") Then
    '*** Jay 3.3(FG)

      '*** (Bef 2.13) Jen
      ''*** (24/02/06) For Kama (Jen 2.12)
      'If wFlagLocalExp = "Y" Then
      '  Call moCn.RepRes(mRep.Subreport3.OpenSubreport, ws_RmSubRep)
      'ElseIf wFlagLocalExp = "N" Then
      '*** (Bef 2.13) Jen
        Call moCn.RepRes(mRep.Subreport2.OpenSubreport, ws_RmSubRep)
      '*** (Bef 2.13) Jen
      'End If
      ''*** (24/02/06) For Kama (Jen 2.12)
      '*** (Bef 2.13) Jen
      
      'Diagold.3- IGST must be shown as 0% in DTA Invoice
      If wFlagLocalExp = "N" Then
      Rep.TxtIGSTRate.Suppress = True
      End If
     
      '4.1.2
      Rep.TxtDCWt.Suppress = IIF(ADC("UOYN13") = "Y", False, True)
      Rep.TxtCSWt.Suppress = IIF(ADC("UOYN13") = "Y", False, True)
      Rep.FldDCWt.Suppress = IIF(ADC("UOYN13") = "Y", False, True)
      Rep.FldCSWt.Suppress = IIF(ADC("UOYN13") = "Y", False, True)
      Rep.FldTotDCWt.Suppress = IIF(ADC("UOYN13") = "Y", False, True)
      Rep.FldTotCSWt.Suppress = IIF(ADC("UOYN13") = "Y", False, True)
      Call ChkBoxArr_Click(13)
    
    'ElseIf UCase(adc.MenuCd) = UCase("RepShpBill") Then  '*** (Bef 2.13)
    '*** Jay 3.3(FG)
    ' ***** Manali Trading Module
    ElseIf UCase(ADC.MenuCd) = UCase("RepShpBill") Or UCase(ADC.MenuCd) = UCase("RepShpBillMulti") Or _
           UCase(ADC.MenuCd) = UCase("RepFgShpBill") Or UCase(ADC.MenuCd) = UCase("RepJtShpBill") Then
    '*** Jay 3.3(FG)
      '*** (Bef 2.13) Jen
      'If wFlagLocalExp = "Y" Then
      '  Call moCn.RepRes(mRep.Subreport2.OpenSubreport, ws_RmSubRep) '*** (24/02/06) For Kama (Jen 2.12)
      'Else
      '  Call moCn.RepRes(mRep.Subreport1.OpenSubreport, ws_RmSubRep)
      'End If
      '*** (Bef 2.13) Jen
      Call moCn.RepRes(mRep.Subreport1.OpenSubreport, ws_RmSubRep) '*** (Jen 2.13)
      
      ' **** Manali 3.6.0 - 17/11/09 - Circular No Printed instead of GR No (InFob < 25000 and DTA )
      If wFlagLocalExp = "Y" Then
        GRP_REP.SetFormula mRep, "wGrNoDt", "if Maximum({rdo.InFOB})>= 25000 Then " + _
                                            " iif({rdo.InGrNo}<>'', {rdo.InGrNo}, '') " + _
                                            "+ iif({rdo.InGrNo}<>'' and {rdo.InGrDt}<> DateValue (1980,01,01), ' & ', '')  + " + _
                                            "iif({rdo.InGrDt}<> DateValue (1980,01,01), ToText({rdo.InGrDt}), '') " + _
                                            "Else {rdo.InCircularNo} "
        
      Else
        GRP_REP.SetFormula mRep, "wGrNoDt", " iif({rdo.InGrNo}<>'', {rdo.InGrNo}, '') " + _
                                            "+ iif({rdo.InGrNo}<>'' and {rdo.InGrDt}<> DateValue (1980,01,01), ' & ', '')  + " + _
                                            "iif({rdo.InGrDt}<> DateValue (1980,01,01), ToText({rdo.InGrDt}), '') "
      End If
      ' **** Manali 3.6.0 - 17/11/09 - Circular No Printed instead of GR No (InFob < 25000 and DTA )
      
    '*** Jay 3.3(FG)
    ' ***** Manali Trading Module
    ElseIf UCase(ADC.MenuCd) = UCase("RepGRFormFrnt") Or UCase(ADC.MenuCd) = UCase("RepGrFormFrntMulti") Or _
           UCase(ADC.MenuCd) = UCase("RepFgGrFrnt") Or UCase(ADC.MenuCd) = UCase("RepJtGrFrnt") Then
    '*** Jay 3.3(FG)

      '*** (Bef 2.13) Jen
      'If wFlagLocalExp = "Y" Then
      '  Call moCn.RepRes(mRep.Subreport2.OpenSubreport, ws_RmSubRep) '*** (24/02/06) For Kama (Jen 2.12)
      'Else
      '*** (Bef 2.13) Jen
        Call moCn.RepRes(mRep.Subreport1.OpenSubreport, ws_RmSubRep)
      '*** (Bef 2.13) Jen
      'End If
      '*** (Bef 2.13) Jen
    End If
    mRep.SecRmSummary.Suppress = True
    ''REP.SecRmSummary.KeepTogether = True
    ' **** Manali 3.6.0 - 17/11/09 - Circular No Printed instead of GR No (InFob < 25000 and DTA )
    GRP_REP.SetFormula mRep, "wGrNoDt", " iif({rdo.InGrNo}<>'', {rdo.InGrNo}, '') " + _
                                        "+ iif({rdo.InGrNo}<>'' and {rdo.InGrDt}<> DateValue (1980,01,01), ' & ', '')  + " + _
                                        "iif({rdo.InGrDt}<> DateValue (1980,01,01), ToText({rdo.InGrDt}), '') "
    ' **** Manali 3.6.0 - 17/11/09 - Circular No Printed instead of GR No (InFob < 25000 and DTA )
  End If
  '*** (Jen 2.12)
  
  ' **** Manali 3.03 - 26/06/08 *****
  'sv.24 stncls summary visible only for Show Dia/CS Wt=Yes in invoice printing report
  If UCase(ADC.MenuCd) = UCase("RepInv") And ADC("UoYN13") = "Y" Then Rep.Subreport4.Suppress = False
  
  If wFlagLocalExp = "N" Then
      Rep.FldInMetalUtil.Suppress = True
  Else
      Rep.FldInCmAgAdd1.Suppress = True
      Rep.FldInCmAgAdd2.Suppress = True
      Rep.FldInCmAgAdd3.Suppress = True
  End If
  ' **** Manali 3.03 - 26/06/08 *****
  'sv.44 update TCS description
  Dim wTCSPer As Double
  wTCSPer = moCn.GetFldVal("Select HTCSPer from Head where HCoCd= '" + ADC("UoCoCdFr") + "' and HCd= '" + ctSelfCmCd + "'")
  Rep.TCSDesc.SetText "@" + CStr(wTCSPer) + "%TCS"
  'sv.44 TCS details will show only for invoice printing report
  If UCase(ADC.MenuCd) <> UCase("RepInv") Then
    Rep.InTcsRs1.Suppress = True: Rep.InTCS1.Suppress = True
    Rep.TCSDesc.Suppress = True:  Rep.TxtTcsCur.Suppress = True
    Call QRCode
  End If
        
  CRV_REP.DisplayGroupTree = False
  
  '****** DemarkSec is necessary for the lines to extend till the page footer ******

  '**** Zubin 211 **** '
  '****** wDetLineYN exists only in EmrRepInv. Still ADC.MenuCd is not checked as no error is thrown if mRep<> EmrRepInv ******
    GRP_REP.SetFormula mRep, "wDetLineYN", IIF(ADC("UOYN3") = "Y", "'N'", "'Y'")
  '**** Zubin 211 **** '
  '*********************SRI****************************************
  Call ChngRepFormulaBasedOnChosenOptions
  
  Call DispReport(ADC) '****** Sachin 3.02 - Display Report ******
  
'*** (Jen 2.12)
  Exit Sub
RepErr:
  Call ChkSelScope(Err.Number, Err.Description)
'*** (Jen 2.12)
End Sub
Private Sub ChngRepFormulaBasedOnChosenOptions()
  If UCase(ADC.MenuCd) = UCase("RepInv") Then
    Dim wbShowDCWt, wbShowGrsWt, wbShowNetWt As Boolean
    wbShowDCWt = IIF(UCase(ADC("UoYN13") = "Y"), True, False)
    wbShowNetWt = IIF(UCase(ADC("UoYN4") = "Y"), True, False)
    wbShowGrsWt = IIF(UCase(ADC("UoYN6") = "Y"), True, False)
    Dim wDiaWtFld, wCSWtFld, wDiaCSWtFld, wGrsWtFld, wNetWtFld As String
    Dim wDiaWtHd, wCSWtHd, wDiaCSWtHd, wGrsWtHd, wNetWtHd As String
    Dim wCol1Hd, wCol2Hd, wCol3Hd As String
    Dim wCol1, wCol2, wCol3 As String
    
    If (UCase(ADC("UoInWt")) = "I") Then
      wDiaWtFld = "{rdo.qInvDWt}"
      wCSWtFld = "{rdo.qInvCWt}"
      wDiaCSWtFld = "{@qInvDCSWt}"
      wNetWtFld = "{@qInvNetWt}"
      wGrsWtFld = "{rdo.qInvGrWt}"
    ElseIf (UCase(ADC("UoInWt")) = "Z") Then
      wDiaWtFld = "{rdo.qZDWt}"
      wCSWtFld = "{rdo.qZCWt}"
      wGrsWtFld = "{rdo.qZGrWt}"
      wDiaCSWtFld = "{@qZDCSWt}"
      wNetWtFld = "{@qZNetWt}"
    End If
    
    Rep.TxtDCWt.Suppress = False
    Rep.TxtCSWt.Suppress = False
    Rep.TxtGrWt.Suppress = False
    Rep.FldDCWt.Suppress = False
    Rep.FldCSWt.Suppress = False
    Rep.FldDetGrWt.Suppress = False
    Rep.FldTotDCWt.Suppress = False
    Rep.FldTotCSWt.Suppress = False
    Rep.FldTotGrWt.Suppress = False

    wDiaWtHd = "Dia  Wt (Cts)"
    wCSWtHd = "  CS Wt (Cts)"
    wDiaCSWtHd = "DiaCSWt (Cts)"
    wGrsWtHd = "GrossWt (Gms)"
    wNetWtHd = "NetWt (Gms)"
    
    If (wbShowDCWt And wbShowNetWt And wbShowGrsWt) Then
      '1 Yes Yes Yes Dia + Cs    Net           Gross
        
      wCol1 = wDiaCSWtFld
      wCol2 = wNetWtFld
      wCol3 = wGrsWtFld
      
      wCol1Hd = wDiaCSWtHd
      wCol2Hd = wNetWtHd
      wCol3Hd = wGrsWtHd
      
      
    ElseIf (wbShowDCWt And wbShowNetWt And (Not wbShowGrsWt)) Then
      '2 Yes Yes No  Diamond     ColourStone   Net
      wCol1 = wDiaWtFld
      wCol2 = wCSWtFld
      wCol3 = wNetWtFld
      
      wCol1Hd = wDiaWtHd
      wCol2Hd = wCSWtHd
      wCol3Hd = wNetWtHd
      
    ElseIf (wbShowDCWt And (Not wbShowNetWt) And (Not wbShowGrsWt)) Then
      '3 Yes No  No  Diamond     ColourStone   Blank
      wCol1 = wDiaWtFld
      wCol2 = wCSWtFld
      wCol3 = "{@ValueZero}"
  
      
      wCol1Hd = wDiaWtHd
      wCol2Hd = wCSWtHd
      wCol3Hd = ""
      
    ElseIf (wbShowDCWt And (Not wbShowNetWt) And wbShowGrsWt) Then
      '4 Yes No  Yes Diamond     ColourStone   Gross
      wCol1 = wDiaWtFld
      wCol2 = wCSWtFld
      wCol3 = wGrsWtFld
  
      wCol1Hd = wDiaWtHd
      wCol2Hd = wCSWtHd
      wCol3Hd = wGrsWtHd
      
    ElseIf ((Not wbShowDCWt) And wbShowNetWt And wbShowGrsWt) Then
      '5 No  Yes Yes Blank       Net           Gross
      wCol1 = "{@ValueZero}"
      wCol2 = wNetWtFld
      wCol3 = wGrsWtFld
  
      wCol1Hd = ""
      wCol2Hd = wNetWtHd
      wCol3Hd = wGrsWtHd
      
    ElseIf ((Not wbShowDCWt) And (Not wbShowNetWt) And wbShowGrsWt) Then
      '6 No  No  Yes Blank       Blank         Gross
      wCol1 = "{@ValueZero}"
      wCol2 = "{@ValueZero}"
      wCol3 = wGrsWtFld
  
      wCol1Hd = ""
      wCol2Hd = ""
      wCol3Hd = wGrsWtHd
     
     ElseIf ((Not wbShowDCWt) And wbShowNetWt And (Not wbShowGrsWt)) Then
      ' 7 No  Yes No  Blank       Blank         Net
            wCol1 = "{@ValueZero}"
            wCol2 = "{@ValueZero}"
      wCol3 = wNetWtFld
  
      wCol1Hd = ""
      wCol2Hd = ""
      wCol3Hd = wNetWtHd
      
    ElseIf ((Not wbShowDCWt) And (Not wbShowNetWt) And (Not wbShowGrsWt)) Then
      ' 8 No  No  No  Blank       Blank         Blank
      wCol1 = "{@ValueZero}"
      wCol2 = "{@ValueZero}"
      wCol3 = "{@ValueZero}"
      
      wCol1Hd = ""
      wCol2Hd = ""
      wCol3Hd = ""
    End If
    
    GRP_REP.SetFormula Rep, "wDetDWt", "Sum(" + wCol1 + ", {@wDet})"
    GRP_REP.SetFormula Rep, "wDetCWt", "Sum(" + wCol2 + ", {@wDet})"
    GRP_REP.SetFormula mRep, "wDetGrWt", "Sum(" + wCol3 + ", {@wDet})"
    
    GRP_REP.SetFormula Rep, "wTotDWt", "Sum(" + wCol1 + ")"
    GRP_REP.SetFormula Rep, "wTotCWt", "Sum(" + wCol2 + ")"
    GRP_REP.SetFormula Rep, "wTotGrWt", "Sum(" + wCol3 + ")"
    
    Rep.TxtDCWt.SetText wCol1Hd
    Rep.TxtCSWt.SetText wCol2Hd
    Rep.TxtGrWt.SetText wCol3Hd
  End If
End Sub
'4.1.3
Private Sub SetLogo(pRep As CRAXDRT.Report)
If ADC("UOYN14") = "Y" Then
    pRep.PicLogo.Suppress = False
    pRep.PicDTCLogo.Suppress = False
End If
End Sub

' **** Manali 3.5.0 - 18/11/08 - Check Boxes Added
Private Sub ChkBoxArr_Click(Index As Integer)
With ChkBoxArr(Index)
  Select Case Index
  Case Is = 0       ' Tag = "ShowPic"
    If .Value = Checked Then ADC("UoYN") = "Y" Else ADC("UoYN") = "N"
  Case Is = 1      ' tag = "ShowCustPoDsg"
    If .Value = Checked Then ADC("UoYN1") = "Y" Else ADC("UoYN1") = "N"
  Case Is = 2      ' Tag = "ShowMetLs"
    If .Value = Checked Then ADC("UoYN2") = "Y" Else ADC("UoYN2") = "N"
    If .Value = Unchecked And ADC("UoYN11") = "Y" Then ADC("UoYN11") = "N": ChkBoxArr(11).Value = Unchecked '3.11.0
  Case Is = 3       ' tag = "ShowDetLine"
    If .Value = Checked Then ADC("UoYN3") = "Y" Else ADC("UoYN3") = "N"
  Case Is = 4      ' tag = "ShowNetWt"
    If .Value = Checked Then
      ADC("UoYN4") = "Y"
      If (ADC.MenuCd) <> UCase("RepInv") Then _
        ChkBoxArr(6).Value = Unchecked: ADC("UoYN6") = "N"    ' ***** Manali 3.6.0 - 30/09/09 - Show Net Wt instead of DC Wt
    Else
      ADC("UoYN4") = "N"
    End If
  Case Is = 5       ' Tag = "ShowCustDecl"
    If .Value = Checked Then ADC("UoYN5") = "Y" Else ADC("UoYN5") = "N"
  Case Is = 6        ' Tag = "ShowDCWt"
    If .Value = Checked Then
      ADC("UoYN6") = "Y"
      If (ADC.MenuCd) <> UCase("RepInv") Then _
        ChkBoxArr(4).Value = Unchecked: ADC("UoYN4") = "N"    ' ***** Manali 3.6.0 - 30/09/09 - Show Net Wt instead of DC Wt
    Else
      ADC("UoYN6") = "N"
    End If
  Case Is = 7        ' "ShowPayTerms"
    If .Value = Checked Then ADC("UoYN7") = "Y" Else ADC("UoYN7") = "N"
  ' **** Manali 3.6.0 - 09/09/09 - Cost Details in DTA Invoice
  Case Is = 8        ' "ShowMkChg"
    If .Value = Checked Then ADC("UoYN8") = "Y" Else ADC("UoYN8") = "N"
  ' **** Manali 3.6.0 - 09/09/09 - Cost Details in DTA Invoice
  Case Is = 9        ' "ShowMkChg"
    If .Value = Checked Then ADC("UoYN9") = "Y" Else ADC("UoYN9") = "N"
  
  Case Is = 10       ' "ShowBin"    3.11.0
    If .Value = Checked Then ADC("UoYN10") = "Y" Else ADC("UoYN10") = "N"
    
  Case Is = 11       ' "ShowAccInMetLoss 3.11.0
    If ADC("UoYN2") = "N" And .Value = Checked Then .Value = Unchecked  '3.11.0
    If .Value = Checked Then ADC("UoYN11") = "Y" Else ADC("UoYN11") = "N"
  
 Case Is = 12       ' "Sort by OrdCust"    4.1.2
    If .Value = Checked Then ADC("UoYN12") = "Y" Else ADC("UoYN12") = "N"
  
 Case Is = 13       ' "Show Dia/CS Wt"    4.1.2
    If .Value = Checked Then ADC("UoYN13") = "Y" Else ADC("UoYN13") = "N"
    
 Case Is = 14       ' "Show Logo"    4.1.3
    If .Value = Checked Then ADC("UoYN14") = "Y" Else ADC("UoYN14") = "N"
  
  End Select
End With
End Sub
' **** Manali 3.5.0 - 18/11/08 - Check Boxes Added

' **** Manali 3.5.0 - 18/11/08 - Check Boxes Added
Private Sub ChkBoxArr_GotFocus(Index As Integer)
  DispMsg ChkBoxArr(Index).ToolTipText, etInfo
End Sub
' **** Manali 3.5.0 - 18/11/08 - Check Boxes Added

'''' std code not to be changed
Private Sub Form_GotFocus()
  '*** Jenny Color (For Setting System Colors)
  SetMDISysColorsAfterLogin
  '*** Jenny Color (For Setting System Colors)
End Sub
Private Sub Form_Load()
  '*** Jenny Colour
  Dim ws_BtnStr As String
  ws_BtnStr = ""
  Call SetObjColors(Me, ws_BtnStr)
  '*** Jenny Colour
End Sub
Private Sub ADC_Load()
    Set moCn = ADC.Connection    '*** (Bef speed) bef 24/11/06
    
    '*** (Jen 2.13)
    '****** (Jen 3.01) added RepGSPFORMMulti
    '*** (Bef 2.14 P2 12/09/07)
    'If UCase(adc.MenuCd) = UCase("RepShpBillMulti") Or UCase(adc.MenuCd) = UCase("RepGrFormFrntMulti") Or UCase(adc.MenuCd) = UCase("RepGSPFORMMulti") Then
    '*** (Bef 2.14 P2 12/09/07)
    '*** Jay 2.14 P2 12/09/07
    If UCase(ADC.MenuCd) = UCase("RepShpBillMulti") Or UCase(ADC.MenuCd) = UCase("RepGrFormFrntMulti") Or _
    UCase(ADC.MenuCd) = UCase("RepGSPFORMMulti") Or UCase(ADC.MenuCd) = UCase("RepGrFormFrntNewMul") Then
    '*** Jay 2.14 P2 12/09/07
      Set ADC.FirNKeyCtl = ADC("UoInExpNoSel")
    Else
      Set ADC.FirNKeyCtl = ADC("UoInTcFr") 'adc("UoCoCdFr")
    End If
    '*** (Jen 2.13)
    
    'Set adc.FirNKeyCtl = adc("UoInTcFr") 'adc("UoCoCdFr")    '*** (Bef 2.13)
    
    
    Call InitProp(Me)
    
    ' **** Manali Trading Module
    Select Case UCase(ADC.MenuCd)
    Case Is = UCase("RepFgInv"), UCase("RepFgShpBill"), UCase("RepFgGrFrnt"), UCase("RepFgGsp")
      gs_InTcTyp = "FG"
    Case Is = UCase("RepJtInv"), UCase("RepJtShpBill"), UCase("RepJtGrFrnt"), UCase("RepJtGsp")
      gs_InTcTyp = ctJtTcTyp
    Case Else
      gs_InTcTyp = "IN"
    End Select
    ' **** Manali Trading Module
    
    ' Bef Trading Module
    '*** Jay 3.3(FG)
''    If UCase(adc.MenuCd) <> UCase("RepFgInv") And UCase(adc.MenuCd) <> UCase("RepFgShpBill") And _
''        UCase(adc.MenuCd) <> UCase("RepFgGrFrnt") And UCase(adc.MenuCd) <> UCase("RepFgGsp") Then
''        gs_InTcTyp = "IN"
''    Else
''        gs_InTcTyp = "FG"
''    End If
    '*** Jay 3.3(FG)
    
    '*** (Jen 2.12)
    gs_RmGrp = "INV"
    gs_Tbl = "Param"
    gs_PTyp = "GRLSSUMM"
    '*** (Jen 2.12)
          
    '********Geeta*****Emr207
    wFlagLocalExp = moCn.GetFldVal("Select hExpLocal from Head where HCoCd= '" + gs_CoCd + "' and " + _
                                                                  "HCd= '" + ctSelfCmCd + "' ")
 '   wFlagLocalExp = "N"  '**ramram
                                                                          
    lblDocTyp.Visible = False: ADC("wDocTyp").Visible = False
    '********Geeta*****Emr207
    
        '*** (Jen 2.13)
    '****** (Jen 3.01) added RepGSPFORMMulti ******
    '*** (Bef 2.14 P2 12/09/07)
    'If UCase(adc.MenuCd) = UCase("RepShpBillMulti") Or UCase(adc.MenuCd) = UCase("RepGrFormFrntMulti") _
    '  Or UCase(adc.MenuCd) = UCase("RepGSPFORMMulti") Then
    '*** (Bef 2.14 P2 12/09/07)
    '*** Jay 2.14 P2 12/09/07
    '****** Manali 3.2.2 14/05/08 - Modified exp no
    If UCase(ADC.MenuCd) = UCase("RepShpBillMulti") Or UCase(ADC.MenuCd) = UCase("RepGrFormFrntMulti") _
      Or UCase(ADC.MenuCd) = UCase("RepGSPFORMMulti") Or UCase(ADC.MenuCd) = UCase("RepGrFormFrntNewMul") Then
      '*** Jay 2.14 P2 12/09/07
      mb_MultiInv = True
      LblExpNo.Caption = "Exp No(s)": ADC("UoInExpNoSel").Visible = True: ADC("UoInExpNoSel").Enabled = True
      ADC("UoInExpNoFr").Visible = False: ADC("UoInExpNoFr").Enabled = False
      LblInvNo.Visible = False: LblTcSlash.Visible = False: LblYySlash.Visible = False: LblChrSlash.Visible = False
      ADC("UoInTcFr").Visible = False: ADC("UoInTcFr").Enabled = False
      ADC("UoInYyFr").Visible = False: ADC("UoInYyFr").Enabled = False
      ADC("UoInChrFr").Visible = False: ADC("UoInChrFr").Enabled = False
      ADC("UoInNoFr").Visible = False: ADC("UoInNoFr").Enabled = False
      LblInPrdDesc.Visible = True: TxtInPrdDesc.Visible = True: TxtInPrdDesc.Enabled = True
      LblPrdDescCd.Visible = True: ADC("UoPMCdTo").Visible = True: ADC("UoPMCdTo").Enabled = True
      lblModInExpNos.Visible = True: ADC("UoEntInExpNo").Visible = True: ADC("UoEntInExpNo").Enabled = True
    Else
      mb_MultiInv = False
      LblExpNo.Caption = "Exp No": ADC("UoInExpNoSel").Visible = False: ADC("UoInExpNoSel").Enabled = False
      ADC("UoInExpNoFr").Visible = True: ADC("UoInExpNoFr").Enabled = True
      LblInvNo.Visible = True: LblTcSlash.Visible = True: LblYySlash.Visible = True: LblChrSlash.Visible = True
      ADC("UoInTcFr").Visible = True: ADC("UoInTcFr").Enabled = True
      ADC("UoInYyFr").Visible = True: ADC("UoInYyFr").Enabled = True
      ADC("UoInChrFr").Visible = True: ADC("UoInChrFr").Enabled = True
      ADC("UoInNoFr").Visible = True: ADC("UoInNoFr").Enabled = True
      LblInPrdDesc.Visible = False: TxtInPrdDesc.Visible = False: TxtInPrdDesc.Enabled = False
      LblPrdDescCd.Visible = False: ADC("UoPMCdTo").Visible = False: ADC("UoPMCdTo").Enabled = False
      lblModInExpNos.Visible = False: ADC("UoEntInExpNo").Visible = False: ADC("UoEntInExpNo").Enabled = False
    End If
    '*** (Jen 2.13)
    
    Select Case UCase(ADC.MenuCd)
    '*** Jay 3.3(FG)
    ' ***** Manali Trading Module
    Case Is = UCase("RepInv"), UCase("RepFgInv"), UCase("RepJtInv")
    '*** Jay 3.3(FG)
      'Me.Caption = "Invoice Printing"
        '*** Jay 3.3(FG)
        ' ***** Manali Trading Module
      ADC("UoInWt") = IIF(UCase(ADC.MenuCd) = UCase("RepFgInv") Or UCase(ADC.MenuCd) = UCase("RepJtInv"), "I", moCn.GetFldVal("Select hInRepWtValAs from Head where HCoCd = '" + ADC("UoCoCdFr") + "' and " + _
                                     "HCd='" + ctSelfCmCd + "'"))
      If UCase(ADC.MenuCd) = UCase("RepFgInv") Or UCase(ADC.MenuCd) = UCase("RepJtInv") Then
        ADC("UoInWt").Visible = False: ADC("UoInWt").Enabled = False: LblWtAs.Visible = False
        ADC("UOYN2").Visible = False: ADC("UOYN2").Enabled = False: LblMetLsSumm.Visible = False
        ChkBoxArr(2).Enabled = False: ChkBoxArr(2).Visible = False    ' **** Manali 3.5.0 - 18/11/08 - Check Boxes Added
        ADC("UOYN11").Visible = False: ADC("UOYN11").Enabled = False    '3.11.0
        ChkBoxArr(11).Enabled = False: ChkBoxArr(11).Visible = False    '3.11.0
      End If
      
      ChkBoxArr(13).Enabled = True: ChkBoxArr(13).Visible = True    '4.1.2
      '*** Jay 3.3(FG)
      ' **** Zubin 211 **** '
      If UCase(wFlagLocalExp) = "Y" Then ChkBoxArr(3).Enabled = False: ChkBoxArr(3).Visible = False       ' **** Manali 3.5.0 - 18/11/08 - Check Boxes Added
                                         'adc("UOYN3").Visible = False: LblDetLine.Visible = False
      ' **** Zubin 211 **** '
      ' Zubin 212
      ' ***** Manali 3.6.0 - 30/09/09 - Show Net Wt instead of DC Wt
''      adc("UOYN4").Visible = False: LblShowNetWt.Visible = False
''      ChkBoxArr(4).Enabled = False: ChkBoxArr(4).Visible = False  ' **** Manali 3.5.0 - 18/11/08 - Check Boxes Added
      ' ***** Manali 3.6.0 - 30/09/09 - Show Net Wt instead of DC Wt
      ' Zubin 212
      
      LblShowPayTerms.Visible = False: ADC("UoYN7").Visible = False '*** (Jen 2.14 Next)
      ChkBoxArr(7).Enabled = False: ChkBoxArr(7).Visible = False    ' **** Manali 3.5.0 - 18/11/08 - Check Boxes Added
      
      
      ' **** Manali Trading Module
      If UCase(wFlagLocalExp) = "N" Or UCase(ADC.MenuCd) = UCase("RepJtInv") Or UCase(ADC.MenuCd) = UCase("RepFgInv") Then ChkBoxArr(8).Enabled = False: ChkBoxArr(8).Visible = False    ' ***** Manali 3.6.0 - 09/09/09 - Cost Details in DTA Invoice
   
   Case Is = UCase("RepInvLoose")
      'Me.Caption = "Invoice Printing"
      ADC("UoInWt") = moCn.GetFldVal("Select hInRepWtValAs from Head where HCoCd = '" + ADC("UoCoCdFr") + "' and " + _
                                     "HCd='" + ctSelfCmCd + "'")
      
      ' **** Zubin 211 **** '
      If UCase(wFlagLocalExp) = "Y" Then ChkBoxArr(3).Enabled = False: ChkBoxArr(3).Visible = False    ' **** Manali 3.5.0 - 18/11/08 - Check Boxes Added
                                         ' adc("UOYN3").Visible = False: LblDetLine.Visible = False
                                         
      ' **** Zubin 211 **** '
      ' Zubin 212
      ADC("UOYN4").Visible = False: LblShowNetWt.Visible = False
      ChkBoxArr(4).Enabled = False: ChkBoxArr(4).Visible = False  ' **** Manali 3.5.0 - 18/11/08 - Check Boxes Added
      ' Zubin 212
      
      '*** (Jen 2.13)
      ADC("UOYN2").Visible = False: ADC("UOYN2").Enabled = False: LblMetLsSumm.Visible = False
      ChkBoxArr(2).Enabled = False: ChkBoxArr(2).Visible = False    ' **** Manali 3.5.0 - 18/11/08 - Check Boxes Added
      
      '3.11.0 - Met Loss Accessories Option
      ADC("UOYN11").Visible = False: ADC("UOYN11").Enabled = False
      ChkBoxArr(11).Enabled = False: ChkBoxArr(11).Visible = False
      
      ADC("UOYN6").Visible = False: ADC("UOYN6").Enabled = False: LblShowDCWt.Visible = False
      ChkBoxArr(6).Enabled = False: ChkBoxArr(6).Visible = False    ' **** Manali 3.5.0 - 18/11/08 - Check Boxes Added
      '*** (Jen 2.13)
      
      LblShowPayTerms.Visible = False: ADC("UoYN7").Visible = False '*** (Jen 2.14 Next)
      ChkBoxArr(7).Enabled = False: ChkBoxArr(7).Visible = False    ' **** Manali 3.5.0 - 18/11/08 - Check Boxes Added
      ChkBoxArr(8).Enabled = False: ChkBoxArr(8).Visible = False    ' ***** Manali 3.6.0 - 09/09/09 - Cost Details in DTA Invoice
    '*** Jay 3.3(FG)
    'Case Is = UCase("RepShpBill"), UCase("RepShpBillLoose")  '*** (Bef 2.13)
    ' ***** Manali Trading Module
    Case Is = UCase("RepShpBill"), UCase("RepShpBillLoose"), UCase("RepShpBillMulti"), UCase("RepFgShpBill"), UCase("RepJtShpBill")
    '*** Jay 3.3(FG)
      'Me.Caption = "Shipping Bill"
      ADC("UoYN").Visible = False: ADC("UoYN").Enabled = False: LblPic.Visible = False
      ChkBoxArr(0).Enabled = False: ChkBoxArr(0).Visible = False    ' **** Manali 3.5.0 - 18/11/08 - Check Boxes Added
      
      ADC("UoYN12").Visible = False: ADC("UoYN12").Enabled = False
      ChkBoxArr(12).Enabled = False: ChkBoxArr(12).Visible = False    ' 4.1.2
      
      '********Geeta*****Emr207  210 For Local also
      lblDocTyp.Visible = True: ADC("wDocTyp").Visible = True
      '********Geeta*****Emr207
      '******************************** urmi *********************************************/
      ADC("UoYN1").Visible = False: ADC("UoYN1").Enabled = False: lblRem.Visible = False
      ChkBoxArr(1).Enabled = False: ChkBoxArr(1).Visible = False    ' **** Manali 3.5.0 - 18/11/08 - Check Boxes Added
      '******************************** urmi *********************************************/
      
      ADC("UoYN2").Visible = False: ADC("UoYN2").Enabled = False: LblMetLsSumm.Visible = False
      ChkBoxArr(2).Enabled = False: ChkBoxArr(2).Visible = False    ' **** Manali 3.5.0 - 18/11/08 - Check Boxes Added
      
      '3.11.0 - Accessories in Met Loss
      ADC("UoYN11").Visible = False: ADC("UoYN11").Enabled = False
      ChkBoxArr(11).Enabled = False: ChkBoxArr(11).Visible = False
            
      ADC("UoInWt").Visible = False: ADC("UoInWt").Enabled = False: LblWtAs.Visible = False
      ' **** Zubin 211 **** '
      ADC("UOYN3").Visible = False: LblDetLine.Visible = False
      ChkBoxArr(3).Enabled = False: ChkBoxArr(3).Visible = False    ' **** Manali 3.5.0 - 18/11/08 - Check Boxes Added
      ' **** Zubin 211 **** '
      ' Zubin 212
      ADC("UOYN4").Visible = False: LblShowNetWt.Visible = False
      ChkBoxArr(4).Enabled = False: ChkBoxArr(4).Visible = False    ' **** Manali 3.5.0 - 18/11/08 - Check Boxes Added
      ADC("UOYN5").Visible = False: LblShowCustRem.Visible = False
      ChkBoxArr(5).Enabled = False: ChkBoxArr(5).Visible = False    ' **** Manali 3.5.0 - 18/11/08 - Check Boxes Added
      ' Zubin 212
      '*** (Jen 2.13)
      ADC("UOYN6").Visible = False: ADC("UOYN6").Enabled = False: LblShowDCWt.Visible = False
      ChkBoxArr(6).Enabled = False: ChkBoxArr(6).Visible = False    ' **** Manali 3.5.0 - 18/11/08 - Check Boxes Added
      '*** (Jen 2.13)
      
      LblShowPayTerms.Visible = False: ADC("UoYN7").Visible = False '*** (Jen 2.14 Next)
      ChkBoxArr(7).Enabled = False: ChkBoxArr(7).Visible = False    ' **** Manali 3.5.0 - 18/11/08 - Check Boxes Added
      ChkBoxArr(8).Enabled = False: ChkBoxArr(8).Visible = False    ' ***** Manali 3.6.0 - 09/09/09 - Cost Details in DTA Invoice
    '*** Jay 3.3(FG)
    ' ***** Manali Trading Module
    Case Is = UCase("RepGRFormFrnt"), UCase("RepGrFormLoose"), UCase("RepGrFormFrntMulti"), UCase("RepFgGRFrnt"), UCase("RepJtGRFrnt")
    '*** Jay 3.3(FG)
      'Me.Caption = "GR Form (Front Side)"
      ADC("UoYN").Visible = False: ADC("UoYN").Enabled = False: LblPic.Visible = False
      ChkBoxArr(0).Enabled = False: ChkBoxArr(0).Visible = False    ' **** Manali 3.5.0 - 18/11/08 - Check Boxes Added
      '********Geeta*****Emr207 & 210 For Local also
      lblDocTyp.Visible = True: ADC("wDocTyp").Visible = True
      '********Geeta*****Emr207
      '******************************** urmi *********************************************/
      ADC("UoYN1").Visible = False: ADC("UoYN1").Enabled = False: lblRem.Visible = False
      ChkBoxArr(1).Enabled = False: ChkBoxArr(1).Visible = False    ' **** Manali 3.5.0 - 18/11/08 - Check Boxes Added
      '******************************** urmi *********************************************/
      
      ADC("UoYN12").Visible = False: ADC("UoYN12").Enabled = False
      ChkBoxArr(12).Enabled = False: ChkBoxArr(12).Visible = False    ' 4.1.2

      
      ADC("UoYN2").Visible = False: ADC("UoYN2").Enabled = False: LblMetLsSumm.Visible = False
      ChkBoxArr(2).Enabled = False: ChkBoxArr(2).Visible = False    ' **** Manali 3.5.0 - 18/11/08 - Check Boxes Added
      
      '3.11.0 - Accessories in Metal Loss
      ADC("UoYN11").Visible = False: ADC("UoYN11").Enabled = False
      ChkBoxArr(11).Enabled = False: ChkBoxArr(11).Visible = False
      
      ADC("UoInWt").Visible = False: ADC("UoInWt").Enabled = False: LblWtAs.Visible = False
      ' **** Zubin 211 **** '
      ADC("UOYN3").Visible = False: LblDetLine.Visible = False
      ChkBoxArr(3).Enabled = False: ChkBoxArr(3).Visible = False    ' **** Manali 3.5.0 - 18/11/08 - Check Boxes Added
      ' **** Zubin 211 **** '
      ' Zubin 212
      ADC("UOYN4").Visible = False: LblShowNetWt.Visible = False
      ChkBoxArr(4).Enabled = False: ChkBoxArr(4).Visible = False    ' **** Manali 3.5.0 - 18/11/08 - Check Boxes Added
      ADC("UOYN5").Visible = False: LblShowCustRem.Visible = False
      ChkBoxArr(5).Enabled = False: ChkBoxArr(5).Visible = False    ' **** Manali 3.5.0 - 18/11/08 - Check Boxes Added
      ' Zubin 212
      '*** (Jen 2.13)
      ADC("UOYN6").Visible = False: ADC("UOYN6").Enabled = False: LblShowDCWt.Visible = False
      ChkBoxArr(6).Enabled = False: ChkBoxArr(6).Visible = False    ' **** Manali 3.5.0 - 18/11/08 - Check Boxes Added
      '*** (Jen 2.13)
      
      'LblShowPayTerms.Visible = True: adc("UoYN7").Visible = True '*** (Jen 2.14 Next)
      ChkBoxArr(7).Enabled = True: ChkBoxArr(7).Visible = True    ' **** Manali 3.5.0 - 18/11/08 - Check Boxes Added
      ChkBoxArr(8).Enabled = False: ChkBoxArr(8).Visible = False    ' ***** Manali 3.6.0 - 09/09/09 - Cost Details in DTA Invoice
    '*** Jay 2.14 P2 12/09/07
    Case Is = UCase("RepGRFormFrntNew"), UCase("RepGrFormNewLoose"), UCase("RepGrFormFrntNewMul")
      ADC("UoYN").Visible = False: ADC("UoYN").Enabled = False: LblPic.Visible = False
      ChkBoxArr(0).Enabled = False: ChkBoxArr(0).Visible = False    ' **** Manali 3.5.0 - 18/11/08 - Check Boxes Added
      lblDocTyp.Visible = True: ADC("wDocTyp").Visible = True
      ADC("UoYN1").Visible = False: ADC("UoYN1").Enabled = False: lblRem.Visible = False
      ChkBoxArr(1).Enabled = False: ChkBoxArr(1).Visible = False    ' **** Manali 3.5.0 - 18/11/08 - Check Boxes Added
      ADC("UoYN2").Visible = False: ADC("UoYN2").Enabled = False: LblMetLsSumm.Visible = False
      ChkBoxArr(2).Enabled = False: ChkBoxArr(2).Visible = False    ' **** Manali 3.5.0 - 18/11/08 - Check Boxes Added
      
      ADC("UoYN12").Visible = False: ADC("UoYN12").Enabled = False
      ChkBoxArr(12).Enabled = False: ChkBoxArr(12).Visible = False    ' 4.1.2
      
      '3.11.0 - Accessories in Metal Loss
      ADC("UoYN11").Visible = False: ADC("UoYN11").Enabled = False
      ChkBoxArr(11).Enabled = False: ChkBoxArr(11).Visible = False
      
      ADC("UoInWt").Visible = False: ADC("UoInWt").Enabled = False: LblWtAs.Visible = False
      ADC("UOYN3").Visible = False: LblDetLine.Visible = False
      ChkBoxArr(3).Enabled = False: ChkBoxArr(3).Visible = False    ' **** Manali 3.5.0 - 18/11/08 - Check Boxes Added
      ADC("UOYN4").Visible = False: LblShowNetWt.Visible = False
      ChkBoxArr(4).Enabled = False: ChkBoxArr(4).Visible = False    ' **** Manali 3.5.0 - 18/11/08 - Check Boxes Added
      ADC("UOYN5").Visible = False: LblShowCustRem.Visible = False
      ChkBoxArr(5).Enabled = False: ChkBoxArr(5).Visible = False    ' **** Manali 3.5.0 - 18/11/08 - Check Boxes Added
      ADC("UOYN6").Visible = False: ADC("UOYN6").Enabled = False: LblShowDCWt.Visible = False
      ChkBoxArr(6).Enabled = False: ChkBoxArr(6).Visible = False    ' **** Manali 3.5.0 - 18/11/08 - Check Boxes Added
      'LblShowPayTerms.Visible = True: adc("UoYN7").Visible = True '*** (Jen 2.14 Next)
      ChkBoxArr(7).Enabled = True: ChkBoxArr(7).Visible = True    ' **** Manali 3.5.0 - 18/11/08 - Check Boxes Added
      ChkBoxArr(8).Enabled = False: ChkBoxArr(8).Visible = False    ' ***** Manali 3.6.0 - 09/09/09 - Cost Details in DTA Invoice
      LblGrpCombo.Visible = False: LblDet.Visible = False
      ADC("WDOCTYP").Visible = False: lblDocTyp.Visible = False
      ADC("UOINGRP").Visible = False: ADC("UOINDET").Visible = False
    '*** Jay 2.14 P2 12/09/07
    '****** (Jen 3.01) added RepGSPFORMMulti ******
    '*** Jay 3.3(FG)
    ' ***** Manali Trading Module
    Case Is = UCase("RepGSPFORM"), UCase("RepGSPFORMMulti"), UCase("RepFgGSP"), UCase("RepJtGSP")
    '*** Jay 3.3(FG)
      'Me.Caption = "GSP Form"
      ADC("UoInDet").Visible = False: ADC("UoInDet").Enabled = False: LblDet.Visible = False
      ADC("UoYN").Visible = False: ADC("UoYN").Enabled = False: LblPic.Visible = False
      ChkBoxArr(0).Enabled = False: ChkBoxArr(0).Visible = False    ' **** Manali 3.5.0 - 18/11/08 - Check Boxes Added
      
      '******************************** urmi *********************************************/
      ADC("UoYN1").Visible = False: ADC("UoYN1").Enabled = False: lblRem.Visible = False
      ChkBoxArr(1).Enabled = False: ChkBoxArr(1).Visible = False    ' **** Manali 3.5.0 - 18/11/08 - Check Boxes Added
      '******************************** urmi *********************************************/
      
      ADC("UoYN12").Visible = False: ADC("UoYN12").Enabled = False
      ChkBoxArr(12).Enabled = False: ChkBoxArr(12).Visible = False    ' 4.1.2
      
      ADC("UoYN2").Visible = False: ADC("UoYN2").Enabled = False: LblMetLsSumm.Visible = False
      ChkBoxArr(2).Enabled = False: ChkBoxArr(2).Visible = False    ' **** Manali 3.5.0 - 18/11/08 - Check Boxes Added
            
      '3.11.0 - Accessories In Metal Loss
      ADC("UoYN11").Visible = False: ADC("UoYN11").Enabled = False
      ChkBoxArr(11).Enabled = False: ChkBoxArr(11).Visible = False
      
      ADC("UoInWt").Visible = False: ADC("UoInWt").Enabled = False: LblWtAs.Visible = False
      ' **** Zubin 211 **** '
      ADC("UOYN3").Visible = False: LblDetLine.Visible = False
      ChkBoxArr(3).Enabled = False: ChkBoxArr(3).Visible = False    ' **** Manali 3.5.0 - 18/11/08 - Check Boxes Added
      ' **** Zubin 211 **** '
      ' Zubin 212
      If UCase(wFlagLocalExp) = "Y" Then ChkBoxArr(4).Enabled = False: ChkBoxArr(4).Visible = False    ' **** Manali 3.5.0 - 18/11/08 - Check Boxes Added
                                        'adc("UOYN4").Visible = False: LblShowNetWt.Visible = False

      ADC("UOYN5").Visible = False: LblShowCustRem.Visible = False
      ChkBoxArr(5).Enabled = False: ChkBoxArr(5).Visible = False    ' **** Manali 3.5.0 - 18/11/08 - Check Boxes Added
      ' Zubin 212
      '*** (Jen 2.13)
      ADC("UOYN6").Visible = False: ADC("UOYN6").Enabled = False: LblShowDCWt.Visible = False
      ChkBoxArr(6).Enabled = False: ChkBoxArr(6).Visible = False    ' **** Manali 3.5.0 - 18/11/08 - Check Boxes Added
      '*** (Jen 2.13)
      LblShowPayTerms.Visible = False: ADC("UoYN7").Visible = False '*** (Jen 2.14 Next)
      ChkBoxArr(7).Enabled = False: ChkBoxArr(7).Visible = False    ' **** Manali 3.5.0 - 18/11/08 - Check Boxes Added
      ChkBoxArr(8).Enabled = False: ChkBoxArr(8).Visible = False    ' ***** Manali 3.6.0 - 09/09/09 - Cost Details in DTA Invoice
      
    Case Is = UCase("RepPckLstInvBag")
      ChkBoxArr(2).Enabled = False: ChkBoxArr(2).Visible = False
      ChkBoxArr(11).Enabled = False: ChkBoxArr(11).Visible = False  '3.11.0
      ADC("UoInWt").Visible = False: ADC("UoInWt").Enabled = False: LblWtAs.Visible = False
      ChkBoxArr(3).Enabled = False: ChkBoxArr(3).Visible = False
      ChkBoxArr(4).Enabled = False: ChkBoxArr(4).Visible = False
      ChkBoxArr(5).Enabled = False: ChkBoxArr(5).Visible = False
      ChkBoxArr(6).Enabled = False: ChkBoxArr(6).Visible = False
      ChkBoxArr(7).Enabled = False: ChkBoxArr(7).Visible = False
      ChkBoxArr(8).Enabled = False: ChkBoxArr(8).Visible = False
      ChkBoxArr(9).Enabled = True: ChkBoxArr(9).Visible = True
    End Select
    
    
    
    
    
    Call SetGroupSort
End Sub
Private Sub ATXT_FldChange(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String, ByVal pv_NewValue As Variant)
  Dim ws_CmDets As MwfLib.MDORowSet
  ' **** Manali Trading Module
  Select Case UCase(ADC.MenuCd)
  Case Is = UCase("RepFgInv"), UCase("RepFgShpBill"), UCase("RepFgGrFrnt"), UCase("RepFgGsp")
    gs_InTcTyp = "FG"
  Case Is = UCase("RepJtInv"), UCase("RepJtShpBill"), UCase("RepJtGrFrnt"), UCase("RepJtGsp")
    gs_InTcTyp = ctJtTcTyp
  Case Else
    gs_InTcTyp = "IN"
  End Select
  ' **** Manali Trading Module
    
  gb_InHlpFrOm = False ' **** Manali 3.7.1

    ' Bef Trading Module
    '*** Jay 3.3(FG)
''  If UCase(adc.MenuCd) <> UCase("RepFgInv") And UCase(adc.MenuCd) <> UCase("RepFgShpBill") And _
''        UCase(adc.MenuCd) <> UCase("RepFgGrFrnt") And UCase(adc.MenuCd) <> UCase("RepFgGsp") Then
''        gs_InTcTyp = "IN"
''  Else
''      gs_InTcTyp = "FG"
''  End If
  '*** Jay 3.3(FG)
  '*** (Jen 2.12)
  gs_RmGrp = "INV"
  gs_Tbl = "Param"
  gs_PTyp = "GRLSSUMM"
  '*** (Jen 2.12)
  
  Call SetProp(Me, IdName, Valid, Cancel, ErrMsg)
  
  '*** (Jen 2.13)
  Select Case UCase(IdName)
  Case Is = UCase("UoInExpNoSel")
    If mb_MultiInv = True Then
      Dim ws_MinCurCd As String, ws_MaxCurCd As String
      Dim ws_MinCnvFct As String, ws_MaxCnvFct As String
      Dim ws_MinCnvRt As String, ws_MaxCnvRt As String
      'If adc("UoInExpNoSel") = "" Then Cancel = True: ErrMsg = "Invoice Export Number Scope Cannot Be Blank": Exit Sub
      If pv_NewValue <> "" Then
        '*** (Jen 3.1.3) added " InCoCd= '" + adc("UoCoCdFr") + "' and "
        'sv.40 company code added
        Cancel = Not (ChkVal("Select 'x' from InvHd where " + IIF(gs_Partition = ctCurrPrtn, " InPrtKey='" + ctCurrPrtn + "' and ", "") + " InCoCd= '" + ADC("UoCoCdFr") + "' and InExpNo in (" + pv_NewValue + ")", ErrMsg))
        If Cancel = True Then ErrMsg = "Invalid Inv Exp Number(s) Scope": Exit Sub
      
        ws_MinCurCd = moCn.GetFldVal("Select min(CmCurCd) from InvHd join CustMst on CmCtg= 'C' and CmCd= InCmCd where " + IIF(gs_Partition = ctCurrPrtn, " InPrtKey='" + ctCurrPrtn + "' and ", "") + " InCoCd= '" + ADC("UoCoCdFr") + "' and InExpNo in (" + pv_NewValue + ")")
        ws_MaxCurCd = moCn.GetFldVal("Select max(CmCurCd) from InvHd join CustMst on CmCtg= 'C' and CmCd= InCmCd where " + IIF(gs_Partition = ctCurrPrtn, " InPrtKey='" + ctCurrPrtn + "' and ", "") + " InCoCd= '" + ADC("UoCoCdFr") + "' and InExpNo in (" + pv_NewValue + ")")
        If ws_MinCurCd <> ws_MaxCurCd Then Cancel = True: ErrMsg = "Invoice Export Number(s) Have Different Currency Codes": Exit Sub
        
        'min (InCnvFct)
        ws_MinCnvFct = Format(moCn.GetFldVal("Select min (RTrim (LTrim (InMulDiv))+ Str(convert (decimal (16, 4), InCnvFct))) from InvHd where " + IIF(gs_Partition = ctCurrPrtn, " InPrtKey='" + ctCurrPrtn + "' and ", "") + " InCoCd= '" + ADC("UoCoCdFr") + "' and InExpNo in (" + pv_NewValue + ")"), "####0.0000")
        ws_MaxCnvFct = Format(moCn.GetFldVal("Select max (RTrim (LTrim (InMulDiv))+ Str(convert (decimal (16, 4), InCnvFct))) from InvHd where " + IIF(gs_Partition = ctCurrPrtn, " InPrtKey='" + ctCurrPrtn + "' and ", "") + " InCoCd= '" + ADC("UoCoCdFr") + "' and InExpNo in (" + pv_NewValue + ")"), "####0.0000")
        If ws_MinCnvFct <> ws_MaxCnvFct Then Cancel = True: ErrMsg = "Invoice Export Number(s) Have Different Currency Conversion Factors": Exit Sub
      
        'min(InCnvRt)
        ws_MinCnvRt = Format(moCn.GetFldVal("Select min (RTrim (LTrim (InRsMulDiv))+ Str(convert (decimal (16, 4), InCnvRt))) from InvHd where " + IIF(gs_Partition = ctCurrPrtn, " InPrtKey='" + ctCurrPrtn + "' and ", "") + " InCoCd= '" + ADC("UoCoCdFr") + "' and InExpNo in (" + pv_NewValue + ")"), "####0.0000")
        ws_MaxCnvRt = Format(moCn.GetFldVal("Select max (RTrim (LTrim (InRsMulDiv))+ Str(convert (decimal (16, 4), InCnvRt))) from InvHd where " + IIF(gs_Partition = ctCurrPrtn, " InPrtKey='" + ctCurrPrtn + "' and ", "") + " InCoCd= '" + ADC("UoCoCdFr") + "' and InExpNo in (" + pv_NewValue + ")"), "####0.0000")
        If ws_MinCnvRt <> ws_MaxCnvRt Then Cancel = True: ErrMsg = "Invoice Export Number(s) Have Different Rupee Conversion Factors": Exit Sub
      
        Cancel = IIF(moCn.GetFldVal("Select count('x') from InvHd where " + IIF(gs_Partition = ctCurrPrtn, " InPrtKey='" + ctCurrPrtn + "' and ", "") + " InCoCd= '" + ADC("UoCoCdFr") + "' and InExpNo in (" + pv_NewValue + ")") > 9, True, False)
        If Cancel = True Then ErrMsg = "Not More Than 9 Inv Exp Number(s) Can Be Selected": Exit Sub
        
        '******** Manali 3.2.2 - Inv Exp Limit extended to 9
        'Cancel = IIF(moCn.GetFldVal("Select count('x') from InvHd where " + IIF(gs_Partition = ctCurrPrtn, " InPrtKey='" + ctCurrPrtn + "' and ", "") + " InCoCd= '" + adc("UoCoCdFr") + "' and InExpNo in (" + pv_NewValue + ")") > 4, True, False)
        'If Cancel = True Then ErrMsg = "Not More Than 4 Inv Exp Number(s) Can Be Selected": Exit Sub
        '*** (Jen 3.1.3) added " InCoCd= '" + adc("UoCoCdFr") + "' and "
        
        '4.1.2
        Set ws_CmDets = moCn.OpenRes(" Select Top 1 CmInvShowGrsWtYN, CmInvShowDiaWtYN from InvHd " + _
                                     "    Join CustMst on CmCtg= 'C' and CmCd= InCmCd " + _
                                     " Where " + IIF(gs_Partition = ctCurrPrtn, " InPrtKey='" + ctCurrPrtn + "' And ", "") + " InCoCd= '" + ADC("UoCoCdFr") + "' and InExpNo In (" + pv_NewValue + ") Order by InCoCd, InTc, InYy, InChr, InNo")
        If Not (ws_CmDets.EOF Or ws_CmDets.BOF) Then
            ADC("UOYN4") = ws_CmDets!CmInvShowGrsWtYN
            ChkBoxArr(4).Value = IIF(ws_CmDets!CmInvShowGrsWtYN = "Y", Checked, Unchecked)
            ADC("UOYN6") = IIF(ws_CmDets!CmInvShowGrsWtYN = "Y", "N", "Y")
            ChkBoxArr(6).Value = IIF(ws_CmDets!CmInvShowGrsWtYN = "Y", Unchecked, Checked)
            ADC("UOYN13") = ws_CmDets!CmInvShowDiaWtYN
            ChkBoxArr(13).Value = IIF(ws_CmDets!CmInvShowDiaWtYN = "Y", Checked, Unchecked)
        End If
      Else
        'Cancel = True: ErrMsg = "Inv Exp Number(s) Scope Cannot Be Blank": Exit Sub
      End If
    End If
    
  Case Is = UCase("UOINEXPNOFR"), UCase("UOINTCFR"), UCase("UOINYYFR"), UCase("UOINCHRFR"), UCase("UOINNOFR")  '4.1.2
    Dim wAddlCnd As String
    wAddlCnd = IIF(Trim(ADC("UoInTcFr")) <> "", " And InTc='" + IIF(UCase(IdName) = UCase("UOINTCFR"), pv_NewValue, Trim(ADC("UoInTcFr"))) + "' ", "")
    wAddlCnd = wAddlCnd + IIF(Trim(ADC("UoInYyFr")) <> "", " And InYy='" + IIF(UCase(IdName) = UCase("UOINYYFR"), pv_NewValue, Trim(ADC("UoInYyFr"))) + "' ", "")
    wAddlCnd = wAddlCnd + IIF(Trim(ADC("UoInChrFr")) <> "", " And InChr='" + IIF(UCase(IdName) = UCase("UOINCHRFR"), pv_NewValue, Trim(ADC("UoInChrFr"))) + "' ", "")
    wAddlCnd = wAddlCnd + IIF(Trim(ADC("UoInNoFr")) <> "", " And InNo=" + CStr(IIF(UCase(IdName) = UCase("UOINNOFR"), pv_NewValue, Trim(ADC("UoInNoFr")))), "")
    wAddlCnd = wAddlCnd + IIF(Trim(ADC("UoInExpNoFr")) <> "", " And InExpNo='" + CStr(IIF(UCase(IdName) = UCase("UoInExpNoFr"), pv_NewValue, Trim(ADC("UoInExpNoFr")))) + "' ", "")
    Set ws_CmDets = moCn.OpenRes(" Select CmInvShowGrsWtYN, CmInvShowDiaWtYN from InvHd " + _
                                 "    Join CustMst on CmCtg= 'C' and CmCd= InCmCd " + _
                                 " Where " + IIF(gs_Partition = ctCurrPrtn, " InPrtKey='" + ctCurrPrtn + "' And ", "") + " InCoCd= '" + ADC("UoCoCdFr") + "' " + wAddlCnd)
    If Not (ws_CmDets.EOF Or ws_CmDets.BOF) Then
      ADC("UOYN4") = ws_CmDets!CmInvShowGrsWtYN
      ChkBoxArr(4).Value = IIF(ws_CmDets!CmInvShowGrsWtYN = "Y", Checked, Unchecked)
      ADC("UOYN6") = IIF(ws_CmDets!CmInvShowGrsWtYN = "Y", "N", "Y")
      ChkBoxArr(6).Value = IIF(ws_CmDets!CmInvShowGrsWtYN = "Y", Unchecked, Checked)
      ADC("UOYN13") = ws_CmDets!CmInvShowDiaWtYN
      ChkBoxArr(13).Value = IIF(ws_CmDets!CmInvShowDiaWtYN = "Y", Checked, Unchecked)
    End If
    
  Case Is = UCase("UoPMCdTo")
    TxtInPrdDesc.text = moCn.GetFldVal("Select PDesc225 from Param where PTyp= 'INVDESC' and PMCd = '" + pv_NewValue + "'")
    
  End Select
  '*** (Jen 2.13)
  
End Sub
Private Sub atxt_FldWhen(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String)
  ' **** Manali Trading Module
  Select Case UCase(ADC.MenuCd)
  Case Is = UCase("RepFgInv"), UCase("RepFgShpBill"), UCase("RepFgGrFrnt"), UCase("RepFgGsp")
    gs_InTcTyp = "FG"
  Case Is = UCase("RepJtInv"), UCase("RepJtShpBill"), UCase("RepJtGrFrnt"), UCase("RepJtGsp")
    gs_InTcTyp = ctJtTcTyp
  Case Else
    gs_InTcTyp = "IN"
  End Select
  ' **** Manali Trading Module

  gb_InHlpFrOm = False ' **** Manali 3.7.1

    ' **** Bef Trading Module
''  '*** Jay 3.3(FG)
''  If UCase(adc.MenuCd) <> UCase("RepFgInv") And UCase(adc.MenuCd) <> UCase("RepFgShpBill") And _
''        UCase(adc.MenuCd) <> UCase("RepFgGrFrnt") And _
''        UCase(adc.MenuCd) <> UCase("RepFgGsp") Then
''        gs_InTcTyp = "IN"
''  Else
''      gs_InTcTyp = "FG"
''  End If
''  '*** Jay 3.3(FG)
  
  '*** (Jen 2.12)
  gs_RmGrp = "INV"
  gs_Tbl = "Param"
  gs_PTyp = "GRLSSUMM"
  '*** (Jen 2.12)

  ' **** Zubin 211 (UoYN3 no more required) **** '
'''''  Select Case UCase(IdName)
'''''  Case Is = UCase("UoYN3")
'''''    If UCase(adc("UoInDet")) = "INVSR" Then
'''''       Cancel = False
'''''    Else
'''''      Cancel = True: ErrMsg = "Running Serial Can Be edited Only When Detail Level='INVSR' ": Exit Sub
'''''    End If
'''''  End Select
  ' **** Zubin 211 (UoYN3 no more required) **** '
  SetProp Me, IdName, When

  '*** (Jen 2.13)
  Select Case UCase(IdName)
  Case Is = UCase("UoPMCdTo")
    Call HlpList.PMCd("INVDESC")
  End Select
  '*** (Jen 2.13)

End Sub

Private Sub atxt_ReCalculate(Index As Integer, ByVal IdName As String, ByVal OldVal As Variant, ByVal RowNum As Integer, ByVal SourceName As String, ByVal SourceColName As String)
' **** Zubin 211 (UoYN3 no more required) **** '
''''''*********Geeta********Emr206*
'''''Select Case UCase(IdName)
'''''  Case Is = UCase("UoYN3")
'''''    If adc("UoInDet") <> "INVSR" Then
'''''      adc("UoYN3") = ""
'''''    End If
'''''  End Select
''''''*********Geeta********Emr206*
' **** Zubin 211 **** '
'****** Manali 3.2.2 Entered Exp no
Select Case UCase(IdName)
  Case Is = UCase("UoEntInExpNo")
     If ADC("UoInExpNoSel") <> "" Then
        Dim wCnd As String
        wCnd = " where " + IIF(gs_Partition = ctCurrPrtn, " Ih.InPrtKey='" + ctCurrPrtn + "' and ", "") + " Ih.InCoCd= '" + ADC("UoCoCdFr") + "' and Ih.InExpNo in (" + ADC("UoInExpNoSel") + ") "
        ADC("UoEntInExpNo") = Left$(moCn.GetStrConcatVal("InExpNo", "InvHd Ih ", wCnd, ", "), 60)
     Else
        ADC("UoEntInExpNo") = ""
     End If
End Select
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
  ADC("UoCoCdFr") = gs_CoCd
  ' ***** Manali Trading Module
  If UCase(ADC.MenuCd) = UCase("RepInv") Or UCase(ADC.MenuCd) = UCase("RepInvLoose") Then
        ADC("UOYN2") = moCn.GetFldVal("Select hInRepMetLs from Head where " + _
                    "HCoCd = '" + ADC("UoCoCdFr") + "' and HCd= '" + ctSelfCmCd + "'")
  ElseIf UCase(ADC.MenuCd) = UCase("RepFgInv") Or UCase(ADC.MenuCd) = UCase("RepJtInv") Then
        ADC("UOYN2") = "N"
  End If
                    
  ' ***** Bef Trading Module
''  If UCase(ADC.MenuCd) = UCase("RepInv") Or UCase(ADC.MenuCd) = UCase("RepInvLoose") Or _
''      UCase(ADC.MenuCd) = UCase("RepFgInv") Or UCase(ADC.MenuCd) = UCase("RepJtInv") Then _
''        ADC("UOYN2") = moCn.GetFldVal("Select hInRepMetLs from Head where " + _
''                    "HCoCd = '" + ADC("UoCoCdFr") + "' and HCd= '" + ctSelfCmCd + "'")
  
  ' **** Zubin 211 **** '
  If ADC("UOINDET") = "" And ADC("UOINDET").Visible = True Then ADC("UOINDET") = "INVSR/SR"
  Select Case UCase(ADC.MenuCd)
    '*** Jay 3.3(FG) [UCase("RepInv") added]
    ' ***** Manali Trading Module
    Case UCase("RepInv"), UCase("RepInvLoose"), UCase("RepInv"), UCase("RepJtInv")
    '*** Jay 3.3(FG) [UCase("RepInv") added]
      'If LblDetLine.Visible = True And adc("UOYN3") = "" Then adc("UOYN3") = "N"
      If wFlagLocalExp = "Y" Then
        ADC("UOYN3").Visible = False: LblDetLine.Visible = False: ADC("UOYN3") = "N"
        ChkBoxArr(3).Enabled = False: ChkBoxArr(3).Visible = False      ' **** Manali 3.5.0 - Check Boxes Added
      Else
        If ADC("UOYN3") = "" Then ADC("UOYN3") = "N"
      End If
  End Select
  ' **** Zubin 211 **** '
  '***** Bhavna added for ce regn no y/n
  
  If UCase(ADC.MenuCd) = UCase("RepInv") Or UCase(ADC.MenuCd) = UCase("RepShpBill") Or _
  UCase(ADC.MenuCd) = UCase("RepShpBillMulti") Or UCase(ADC.MenuCd) = UCase("RepInvLoose") Or _
  UCase(ADC.MenuCd) = UCase("RepShpBillLoose") Or _
  (wFlagLocalExp = "Y" And (UCase(ADC.MenuCd) = UCase("RepGRFormFrnt") Or UCase(ADC.MenuCd) = UCase("RepGrFormLoose") Or _
   UCase(ADC.MenuCd) = UCase("RepFgGRFrnt") Or UCase(ADC.MenuCd) = UCase("RepJtGRFrnt") Or _
   UCase(ADC.MenuCd) = UCase("RepGrFormFrntMulti"))) Then
        If ADC("UOYN10") = "" Then ADC("UOYN10") = "N"
        ChkBoxArr(10).Enabled = True: ChkBoxArr(10).Visible = True
    Else
        ADC("UOYN10") = "N"
        ChkBoxArr(10).Enabled = False: ChkBoxArr(10).Visible = False
        
   End If

  '***** Bhavna added for ce regn no y/n
  '*** (Jen 2.12)
  '*** (Jen 2.12) 18/03/06
  If UCase(ADC.MenuCd) = UCase("RepInvLoose") Then
  '*** (Jen 2.12) 18/03/06
    'And wFlagLocalExp = "N"
    ADC("UoRmGrp").Visible = True
    LblRmGrp.Visible = True
  Else
    ADC("UoRmGrp").Visible = False
    LblRmGrp.Visible = False
    'If wFlagLocalExp = "Y" Then
    ADC("UoRmGrp") = "RMCTG"  '*** Group By RmCtg, Stone Class
    'End If
  End If
  
  '****** Jay 2.14 P2 12/09/07 (added RepGrFormNewLoose)
  If UCase(ADC.MenuCd) = UCase("RepInvLoose") Or UCase(ADC.MenuCd) = UCase("RepShpBillLoose") Or _
     UCase(ADC.MenuCd) = UCase("RepGrFormLoose") Or UCase(ADC.MenuCd) = UCase("RepGrFormNewLoose") Then
    ADC("UOINGRP").Visible = False: ADC("UOINGRP").Enabled = False: ADC("UOINGRP") = "": LblGrpCombo.Visible = False
    ADC("UOINDET").Visible = False: ADC("UOINDET").Enabled = False: ADC("UOINDET") = "": LblDet.Visible = False
    ADC("UOYN").Visible = False: ADC("UOYN").Enabled = False: ADC("UOYN") = "N": LblPic.Visible = False
    ADC("UOYN1").Visible = False: ADC("UOYN1").Enabled = False: ADC("UOYN1") = "N": lblRem.Visible = False
    ADC("UOYN3").Visible = False: ADC("UOYN3").Enabled = False: ADC("UOYN3") = "N": LblDetLine.Visible = False
    ChkBoxArr(0).Enabled = False: ChkBoxArr(0).Visible = False      ' **** Manali 3.5.0 - Check Boxes Added
    ChkBoxArr(1).Enabled = False: ChkBoxArr(1).Visible = False      ' **** Manali 3.5.0 - Check Boxes Added
    ChkBoxArr(3).Enabled = False: ChkBoxArr(3).Visible = False      ' **** Manali 3.5.0 - Check Boxes Added
  End If
  
  ADC("UoPMCdFr").Visible = False: ADC("UoPMCdFr").Enabled = False
  LblMetLsGrp.Visible = False
  '*** (Jen 2.12)
  
  ' **** Manali 3.5.0 - Check Boxes Added
  ' Zubin 212
'''  If adc("UOYN4").Visible = True And adc("UOYN4") = "" Then adc("UOYN4") = "N"
'''  If adc("UOYN5").Visible = True And adc("UOYN5") = "" Then adc("UOYN5") = "N"
'''  If adc("UOYN6").Visible = True And adc("UOYN6") = "" Then adc("UOYN6") = "N"  '*** (Jen 2.13)
  ' Zubin 212
  '*** Jay 3.3(FG)
  ' ***** Manali Trading Module
  ' ***** Manali 3.6.1 - 'Z' wt for Shipping Bill
  If UCase(ADC.MenuCd) <> UCase("RepFgInv") And UCase(ADC.MenuCd) <> UCase("RepInv") And _
    UCase(ADC.MenuCd) <> UCase("RepInvLoose") And UCase(ADC.MenuCd) <> UCase("RepJtInv") Then ADC("UoInWt") = ""

' Bef 3.6.1
'  adc("UoInWt") = IIF(UCase(adc.MenuCd) = UCase("RepFgInv") Or UCase(adc.MenuCd) = UCase("RepJtInv"), "I", moCn.GetFldVal("Select hInRepWtValAs from Head where HCoCd = '" + adc("UoCoCdFr") + "' and " + _
'                                     "HCd='" + ctSelfCmCd + "'"))
  '*** Jay 3.3(FG)
  ADC("UoPMCdTo") = ""    '*** (Jen 2.13)
  
  ' **** Manali 3.5.0 - 18/11/08 - Check Boxes Added
  If ChkBoxArr(0).Visible = True And ADC("UOYN") = "" Then ADC("UOYN") = "N"
  If ChkBoxArr(1).Visible = True And ADC("UOYN1") = "" Then ADC("UOYN1") = "N"
  'If ChkBoxArr(2).Visible = True And adc("UOYN2") = "" Then adc("UOYN2") = "N"
  'If ChkBoxArr(3).Visible = True And adc("UOYN3") = "" Then adc("UOYN3") = "N"
  If ChkBoxArr(4).Visible = True And ADC("UOYN4") = "" Then ADC("UOYN4") = "N"
  If ChkBoxArr(5).Visible = True And ADC("UOYN5") = "" Then ADC("UOYN5") = "N"
  If ChkBoxArr(6).Visible = True And ADC("UOYN6") = "" Then ADC("UOYN6") = "N"  '*** (Jen 2.13)
  If ChkBoxArr(7).Visible = True And ADC("UOYN7") = "" Then ADC("UOYN7") = "N"
  If ChkBoxArr(8).Visible = True And ADC("UOYN8") = "" Then ADC("UOYN8") = "N"    ' ***** Manali 3.6.0 - 09/09/09 - Cost Details in DTA Invoice
  If ChkBoxArr(9).Visible = True And ADC("UOYN9") = "" Then ADC("UOYN9") = "N"
  
  If ChkBoxArr(10).Visible = True And ADC("UOYN10") = "" Then ADC("UOYN10") = "N"   '3.11.0
  If ChkBoxArr(11).Visible = True And ADC("UOYN11") = "" Then ADC("UOYN11") = "N"   '3.11.0
  
  If ChkBoxArr(12).Visible = True And ADC("UOYN12") = "" Then ADC("UOYN12") = "N"   '3.11.0
  If ChkBoxArr(13).Visible = True And ADC("UOYN13") = "" Then ADC("UOYN13") = "N"   '4.1.2
  If ChkBoxArr(14).Visible = True And ADC("UOYN14") = "" Then ADC("UOYN14") = "N"   '4.1.3
  
  If ADC("UoYN") = "Y" Then ChkBoxArr(0).Value = Checked Else ChkBoxArr(0).Value = Unchecked
  If ADC("UoYN1") = "Y" Then ChkBoxArr(1).Value = Checked Else ChkBoxArr(1).Value = Unchecked
  If ADC("UoYN2") = "Y" Then ChkBoxArr(2).Value = Checked Else ChkBoxArr(2).Value = Unchecked
  If ADC("UoYN3") = "Y" Then ChkBoxArr(3).Value = Checked Else ChkBoxArr(3).Value = Unchecked
  If ADC("UoYN4") = "Y" Then ChkBoxArr(4).Value = Checked Else ChkBoxArr(4).Value = Unchecked
  If ADC("UoYN5") = "Y" Then ChkBoxArr(5).Value = Checked Else ChkBoxArr(5).Value = Unchecked
  If ADC("UoYN6") = "Y" Then ChkBoxArr(6).Value = Checked Else ChkBoxArr(6).Value = Unchecked
  If ADC("UoYN7") = "Y" Then ChkBoxArr(7).Value = Checked Else ChkBoxArr(7).Value = Unchecked
  If ADC("UoYN8") = "Y" Then ChkBoxArr(8).Value = Checked Else ChkBoxArr(8).Value = Unchecked     ' ***** Manali 3.6.0 - 09/09/09 - Cost Details in DTA Invoice
  If ADC("UoYN9") = "Y" Then ChkBoxArr(9).Value = Checked Else ChkBoxArr(9).Value = Unchecked
  
  If ADC("UoYN10") = "Y" Then ChkBoxArr(10).Value = Checked Else ChkBoxArr(10).Value = Unchecked    '3.11.0
  If ADC("UoYN11") = "Y" Then ChkBoxArr(11).Value = Checked Else ChkBoxArr(11).Value = Unchecked    '3.11.0
  If ADC("UoYN12") = "Y" Then ChkBoxArr(12).Value = Checked Else ChkBoxArr(12).Value = Unchecked    '4.1.2
  If ADC("UoYN13") = "Y" Then ChkBoxArr(13).Value = Checked Else ChkBoxArr(13).Value = Unchecked    '4.1.2
  If ADC("UoYN14") = "Y" Then ChkBoxArr(14).Value = Checked Else ChkBoxArr(14).Value = Unchecked    '4.1.2
  ' **** Manali 3.5.0 - 18/11/08 - Check Boxes Added
End Sub
Private Sub DispCoNm()
  Dim ws_HName As String
  ws_HName = moCn.GetFldVal("Select HName from Head where HCoCd = '" + ADC("UoCoCdFr") + "' and " + _
                            "HCd = '" + ctSelfCmCd + "' ")
  mRep.wCoCd.SetText ws_HName
  
' ********* Manali 3.2.2 - Before 3.2.2
'  Dim ws_DtTm As String
'  ws_DtTm = Format$(moCn.SrvrDate, "DD-MMM-YYYY")
'  Rep.wSrvrDate.SetText ws_DtTm
'  ws_DtTm = Replace(Format(IIF(CDbl(moCn.SrvrTime) >= 13#, moCn.SrvrTime - 12, IIF(CDbl(moCn.SrvrTime) < 1#, moCn.SrvrTime + 12, moCn.SrvrTime)), "00.00"), ".", ":") + IIF(CDbl(moCn.SrvrTime) >= 12#, " PM", " AM")
'  Rep.wSrvrTime.SetText ws_DtTm

' ******** Manali 3.2.2 - User date and time in reports
  Dim ws_Dt As String, ws_Tm As String
    ws_Dt = Format$(moCn.SrvrDate, "DD-MMM-YYYY")
    ws_Tm = Replace(Format(IIF(CDbl(moCn.SrvrTime) >= 13#, moCn.SrvrTime - 12, IIF(CDbl(moCn.SrvrTime) < 1#, moCn.SrvrTime + 12, moCn.SrvrTime)), "00.00"), ".", ":") + IIF(CDbl(moCn.SrvrTime) >= 12#, " PM", " AM")
  'Rep.TxtRepUsrDtTime.SetText Trim(adc.UsrCd) + " / " + ws_Dt + " / " + ws_Tm
  
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
    
    '*** (09/08/05)
    Set moCn = Nothing   '*** (Bef speed) bef 24/11/06
    Set mRep = Nothing
    Set Rep = Nothing
    Set Rep1 = Nothing
    Set Rep2 = Nothing
    Set Rep3 = Nothing
    'Set Rep4 = Nothing '*** (Bef 2.13) Jen
    'Set Rep5 = Nothing '*** (Bef 2.13) Jen
    'Set Rep6 = Nothing '*** (24/02/06) For Kama (Jen 2.12)   '*** (Bef 2.13) Jen
    'Set Rep7 = Nothing '*** (Bef 2.13) Jen
    '*** (09/08/05)
End Sub

Private Sub GetStnClsSumm(ps_AndCnd As String)
  'sv.24 In Scope if Value as option selected as I/Z, wt and value will change in stncls summary table
  Dim wIrRmWt As String, wIrRmVal As String, wStnClsSummStr As String
  If (UCase(ADC("UoInWt")) = "I") Then
    wIrRmWt = "IrRmIWt"
    wIrRmVal = "IrRmIVal"
  ElseIf (UCase(ADC("UoInWt")) = "Z") Then
    wIrRmWt = "IrRmZWt"
    wIrRmVal = "IrRmZVal"
  End If
  
  'sv.24 query for stncls summary and calling subreport
  wStnClsSummStr = "Select IrTc, IrYy, IrChr, IrNo, RmStnCls as qRmStnCls, " + _
               "sum(IrRmQty) as IrRmQty, sum(" + wIrRmWt + ") as qRmWt, sum(" + wIrRmVal + ") as qRmVal " + _
               "from InvRm " + _
               "join RmMst on RmCd = IrRmCd " + IIF(gs_Partition = ctCurrPrtn, " And IrPrtKey=RmPrtKey ", "") + _
               "join InvDsg on IdIdNo=IrIdIdNo and IdCoCd= IrCoCd and IdTc= IrTc and IdYy= IrYy and " + _
               "IdChr= IrChr and IdNo= IrNo and IdSr= IrSr " + IIF(gs_Partition = ctCurrPrtn, " And IdPrtKey=IrPrtKey ", "") + _
               "join InvHd Ih on InIdNo=IdInIdNo and InCoCd= IdCoCd and InTc = IdTc and InYy= IdYy and " + _
               "InChr= IdChr and InNo= IdNo " + IIF(gs_Partition = ctCurrPrtn, " And InPrtKey=IdPrtKey ", "") + _
               "Left Outer Join Param on PTyp= 'STNCLS' and PMCD=RmCtg and PSCd=RmStnCls " + _
               "where IrRmCtg In ('D','C') " + IIF(gs_Partition = ctCurrPrtn, " and IrPrtKey='" + ctCurrPrtn + "' ", "") + _
                ps_AndCnd + " Group By IrTc, IrYy, IrChr, IrNo, RmStnCls "
  
  Call moCn.RepRes(mRep.Subreport4.OpenSubreport, wStnClsSummStr + ctMaxDopOpt)

End Sub
Private Sub QRCode()
  If UCase(ADC.MenuCd) <> UCase("RepInv") Then
    Rep.PicQRCode.Suppress = True:     Rep.InIRN1.Suppress = True
  End If
  
End Sub
