VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{AB2CC32A-E652-42A2-B341-26ECE02F6386}#2.1#0"; "mwfctlforspeed9.ocx"
Begin VB.Form EmrFrpInvCost 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Invoice Costing"
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
   Begin MwfCtl.MWCTL_BTN1 CmdChgImg 
      Height          =   390
      Left            =   10800
      TabIndex        =   136
      ToolTipText     =   "Analysis Scope Options"
      Top             =   9680
      Visible         =   0   'False
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   688
      BackColor       =   -2147483633
      MaskColor       =   8454143
      Caption         =   "Chg &Image"
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
   Begin MwfCtl.MWCTL_BTN1 CmdDispSc 
      Height          =   390
      Left            =   2310
      TabIndex        =   124
      ToolTipText     =   "Display Scope Options"
      Top             =   9690
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   688
      BackColor       =   -2147483633
      MaskColor       =   8454143
      Caption         =   "Disp&lay "
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
      Height          =   495
      Left            =   12300
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
      Left            =   4440
      TabIndex        =   60
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
      Left            =   6465
      TabIndex        =   68
      Top             =   9630
      Width           =   3555
      Begin MwfCtl.MWCTL_MED aktxt 
         Height          =   345
         Index           =   0
         Left            =   210
         TabIndex        =   69
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
         TabIndex        =   70
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
      Left            =   7680
      TabIndex        =   67
      Top             =   9645
      Visible         =   0   'False
      Width           =   855
      Begin CRVIEWER9LibCtl.CRViewer9 CRV_REP 
         Height          =   495
         Left            =   480
         TabIndex        =   65
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
      Height          =   9765
      Left            =   0
      TabIndex        =   64
      Top             =   60
      Width           =   15255
      Begin VB.Frame FraDispSc 
         Height          =   3225
         Left            =   6000
         TabIndex        =   104
         Top             =   6240
         Visible         =   0   'False
         Width           =   2535
         Begin VB.CheckBox ChkBoxArr 
            Alignment       =   1  'Right Justify
            Caption         =   "Design Qty              "
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
            Index           =   27
            Left            =   90
            TabIndex        =   121
            Tag             =   "ShowDsgQty"
            ToolTipText     =   "Check To Show Design Quantity"
            Top             =   2565
            Width           =   2115
         End
         Begin VB.CheckBox ChkBoxArr 
            Alignment       =   1  'Right Justify
            Caption         =   "Sub PO                   "
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
            Index           =   22
            Left            =   90
            TabIndex        =   123
            Tag             =   "ShowOdPO"
            ToolTipText     =   "Check To Show Sub PO"
            Top             =   2850
            Width           =   2115
         End
         Begin VB.CheckBox ChkBoxArr 
            Alignment       =   1  'Right Justify
            Caption         =   "Unit Price               "
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
            Index           =   21
            Left            =   90
            TabIndex        =   119
            Tag             =   "ShowUnitPrc"
            ToolTipText     =   "Check To Show Unit Price"
            Top             =   2280
            Width           =   2115
         End
         Begin VB.CheckBox ChkBoxArr 
            Alignment       =   1  'Right Justify
            Caption         =   "Rm Quantity           "
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
            Index           =   20
            Left            =   90
            TabIndex        =   117
            Tag             =   "ShowRmQty"
            ToolTipText     =   "Check To Show Rm Quantity"
            Top             =   1995
            Width           =   2115
         End
         Begin VB.CheckBox ChkBoxArr 
            Alignment       =   1  'Right Justify
            Caption         =   "Rm Wt Details         "
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
            Index           =   19
            Left            =   90
            TabIndex        =   115
            Tag             =   "ShowRmWt"
            ToolTipText     =   "Check To Show Rm Weight Details"
            Top             =   1710
            Width           =   2115
         End
         Begin VB.CheckBox ChkBoxArr 
            Alignment       =   1  'Right Justify
            Caption         =   "Rm SubCtg             "
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
            Index           =   18
            Left            =   90
            TabIndex        =   113
            Tag             =   "ShowRmSCtg"
            ToolTipText     =   "Check To Show Rm SubCtg Details"
            Top             =   1425
            Width           =   2115
         End
         Begin VB.CheckBox ChkBoxArr 
            Alignment       =   1  'Right Justify
            Caption         =   "Rm Code Details     "
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
            Index           =   17
            Left            =   90
            TabIndex        =   111
            Tag             =   "ShowRmCd"
            ToolTipText     =   "Check To Show Code Details"
            Top             =   1140
            Width           =   2115
         End
         Begin VB.CheckBox ChkBoxArr 
            Alignment       =   1  'Right Justify
            Caption         =   "Design Grs Wt         "
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
            Index           =   16
            Left            =   90
            TabIndex        =   109
            Tag             =   "ShowGrsWt"
            ToolTipText     =   "Check To Show Design Gross Wt"
            Top             =   855
            Width           =   2115
         End
         Begin VB.CheckBox ChkBoxArr 
            Alignment       =   1  'Right Justify
            Caption         =   "Design Col              "
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
            Index           =   15
            Left            =   90
            TabIndex        =   107
            Tag             =   "ShowDsgCol"
            ToolTipText     =   "Check To Show Color"
            Top             =   570
            Width           =   2115
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   53
            Left            =   2400
            TabIndex        =   106
            ToolTipText     =   "Show Design Color? (Yes/No)"
            Top             =   570
            Visible         =   0   'False
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   503
            Enabled         =   0   'False
            DataType        =   4
            DataField       =   "UOYN15"
            IdName          =   "UOYN15"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   54
            Left            =   2400
            TabIndex        =   108
            ToolTipText     =   "Show Design Grs Wt? (Yes / No)"
            Top             =   855
            Visible         =   0   'False
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   503
            Enabled         =   0   'False
            DataType        =   4
            DataField       =   "UOYN16"
            IdName          =   "UOYN16"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   56
            Left            =   2400
            TabIndex        =   112
            ToolTipText     =   "Show Rm Sub Category (Yes/No) ?"
            Top             =   1425
            Visible         =   0   'False
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   503
            Enabled         =   0   'False
            DataType        =   4
            DataField       =   "UOYN18"
            IdName          =   "UOYN18"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   61
            Left            =   2400
            TabIndex        =   110
            ToolTipText     =   "Show Rm Code Details (Yes / No)"
            Top             =   1140
            Visible         =   0   'False
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   503
            Enabled         =   0   'False
            DataType        =   4
            DataField       =   "UOYN17"
            IdName          =   "UOYN17"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   27
            Left            =   2400
            TabIndex        =   114
            ToolTipText     =   "Show Rm Wt Details (Yes/No) ?"
            Top             =   1710
            Visible         =   0   'False
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   503
            Enabled         =   0   'False
            DataType        =   4
            DataField       =   "UOYN19"
            IdName          =   "UOYN19"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   57
            Left            =   2400
            TabIndex        =   116
            ToolTipText     =   "Show Rm Quantity (Yes/No) ?"
            Top             =   1995
            Visible         =   0   'False
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   503
            Enabled         =   0   'False
            DataType        =   4
            DataField       =   "UOYN20"
            IdName          =   "UOYN20"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   58
            Left            =   2400
            TabIndex        =   118
            ToolTipText     =   "Show Unit Price (Yes/No)?"
            Top             =   2280
            Visible         =   0   'False
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   503
            Enabled         =   0   'False
            DataType        =   4
            DataField       =   "UOYN21"
            IdName          =   "UOYN21"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   59
            Left            =   2400
            TabIndex        =   122
            ToolTipText     =   "Show Sub PO (Yes/No) ?"
            Top             =   2850
            Visible         =   0   'False
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   503
            Enabled         =   0   'False
            DataType        =   4
            DataField       =   "UOYN22"
            IdName          =   "UOYN22"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   31
            Left            =   2400
            TabIndex        =   120
            ToolTipText     =   "Show Unit Price (Yes/No)?"
            Top             =   2565
            Visible         =   0   'False
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   503
            Enabled         =   0   'False
            DataType        =   4
            DataField       =   "UOYN27"
            IdName          =   "UOYN27"
         End
         Begin VB.Label LblMain 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "  Show Design Details"
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
            Left            =   0
            TabIndex        =   105
            Top             =   150
            Width           =   2910
         End
      End
      Begin TabDlg.SSTab TAB_REP 
         Height          =   9465
         Left            =   15
         TabIndex        =   63
         Top             =   45
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   16695
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "&Data Range"
         TabPicture(0)   =   "EmrFrpInvCost.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "fra_tabrep(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "&Group Sort"
         TabPicture(1)   =   "EmrFrpInvCost.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "fra_tabrep(1)"
         Tab(1).ControlCount=   1
         Begin VB.Frame fra_tabrep 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   9070
            Index           =   1
            Left            =   -74940
            TabIndex        =   89
            Top             =   360
            Width           =   15105
            Begin MwfCtl.MWCTL_GRP GRP_REP 
               Height          =   360
               Left            =   240
               TabIndex        =   90
               Top             =   360
               Width           =   8775
               _ExtentX        =   15478
               _ExtentY        =   635
            End
         End
         Begin VB.Frame fra_tabrep 
            BorderStyle     =   0  'None
            Height          =   9070
            Index           =   0
            Left            =   0
            TabIndex        =   66
            Top             =   360
            Width           =   15105
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Show Per Pc"
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
               Index           =   32
               Left            =   4320
               TabIndex        =   149
               Tag             =   "SortOrdCust"
               ToolTipText     =   "Check To Show Per Pc Values"
               Top             =   3840
               Visible         =   0   'False
               Width           =   2450
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Copy Picture"
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
               Index           =   31
               Left            =   4320
               TabIndex        =   147
               Tag             =   "SortOrdCust"
               ToolTipText     =   "Check To Copy Pictures to the Given folder"
               Top             =   3120
               Width           =   2450
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Show Bag / Stk Id"
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
               Left            =   45
               TabIndex        =   58
               Tag             =   "ShowRmVal"
               ToolTipText     =   "Check To Show Bag / Stock Id "
               Top             =   8400
               Width           =   2115
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Show Order Remarks            "
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   345
               Index           =   26
               Left            =   45
               TabIndex        =   59
               Tag             =   "ShowDsgRem"
               ToolTipText     =   "Check To Show Order Design Level Remark"
               Top             =   8640
               Width           =   3135
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Show Intl Qly"
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
               Index           =   30
               Left            =   3000
               TabIndex        =   50
               Tag             =   "ShowVchSr"
               ToolTipText     =   "Check To Show Intl Qly"
               Top             =   7470
               Visible         =   0   'False
               Width           =   2130
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Sort on Ord Cust"
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
               Index           =   29
               Left            =   4365
               TabIndex        =   24
               Tag             =   "SortOrdCust"
               ToolTipText     =   "Check To Sort On Order Customer Code"
               Top             =   2640
               Width           =   2415
            End
            Begin VB.Frame FraOptNetWt 
               BorderStyle     =   0  'None
               Height          =   285
               Left            =   1965
               TabIndex        =   134
               ToolTipText     =   "Select if to show GldAs Wt or Net Wt"
               Top             =   7770
               Width           =   3375
               Begin VB.OptionButton OptNetWt 
                  Caption         =   "Net Wt"
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
                  Left            =   1560
                  TabIndex        =   53
                  Top             =   0
                  Width           =   1365
               End
               Begin VB.OptionButton OptNetWt 
                  Caption         =   "GldAs Wt"
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
                  TabIndex        =   52
                  Top             =   0
                  Width           =   1365
               End
            End
            Begin VB.Frame FraLme 
               BorderStyle     =   0  'None
               Height          =   375
               Left            =   30
               TabIndex        =   130
               ToolTipText     =   "Select Any 2 LMEs"
               Top             =   8100
               Width           =   8895
               Begin VB.CheckBox ChkBoxLme 
                  Caption         =   "LMG"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   270
                  Index           =   0
                  Left            =   1950
                  TabIndex        =   54
                  Tag             =   "ShowVchSr"
                  ToolTipText     =   "Check To Display LMG in Report"
                  Top             =   0
                  Width           =   810
               End
               Begin VB.CheckBox ChkBoxLme 
                  Caption         =   "LMP"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   270
                  Index           =   1
                  Left            =   2950
                  TabIndex        =   55
                  Tag             =   "ShowVchSr"
                  ToolTipText     =   "Check To Display LMP in Report"
                  Top             =   0
                  Width           =   810
               End
               Begin VB.CheckBox ChkBoxLme 
                  Caption         =   "LMS"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   270
                  Index           =   2
                  Left            =   3950
                  TabIndex        =   56
                  Tag             =   "ShowVchSr"
                  ToolTipText     =   "Check To Display LMS in Report"
                  Top             =   0
                  Width           =   810
               End
               Begin VB.CheckBox ChkBoxLme 
                  Caption         =   "LML"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   270
                  Index           =   3
                  Left            =   4950
                  TabIndex        =   57
                  Tag             =   "ShowVchSr"
                  ToolTipText     =   "Check To Display LML in Report"
                  Top             =   0
                  Width           =   810
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   69
                  Left            =   5820
                  TabIndex        =   131
                  ToolTipText     =   "Enter From Design Category"
                  Top             =   30
                  Visible         =   0   'False
                  Width           =   885
                  _ExtentX        =   1561
                  _ExtentY        =   503
                  Enabled         =   0   'False
                  MaxLength       =   5
                  DataType        =   4
                  DataField       =   "UOLME1"
                  IdName          =   "UOLME1"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   70
                  Left            =   6660
                  TabIndex        =   132
                  ToolTipText     =   "Enter From Design Category"
                  Top             =   30
                  Visible         =   0   'False
                  Width           =   885
                  _ExtentX        =   1561
                  _ExtentY        =   503
                  Enabled         =   0   'False
                  MaxLength       =   5
                  DataType        =   4
                  DataField       =   "UOLME2"
                  IdName          =   "UOLME2"
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Select Any 2 LMEs"
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
                  Height          =   345
                  Index           =   4
                  Left            =   30
                  TabIndex        =   133
                  Top             =   0
                  Width           =   1845
               End
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Show Sr"
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
               Index           =   24
               Left            =   45
               TabIndex        =   49
               Tag             =   "ShowVchSr"
               ToolTipText     =   "Check To Show Serials"
               Top             =   7470
               Width           =   2130
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Show Excel Report"
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
               Index           =   23
               Left            =   45
               TabIndex        =   48
               Tag             =   "ShowExcel"
               ToolTipText     =   "Check To Show Report in Excel Format"
               Top             =   7155
               Width           =   2130
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Final Sort On Dsg Cd   "
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
               Left            =   4365
               TabIndex        =   18
               Tag             =   "FinalSrt"
               ToolTipText     =   "Check To Show Picture"
               Top             =   1785
               Width           =   2415
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Show Grp Comb Desc "
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
               Left            =   4365
               TabIndex        =   15
               Tag             =   "ShowGrpDesc"
               ToolTipText     =   "Check To Show Group Description"
               Top             =   1500
               Width           =   2415
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Sort on Inv Cust     "
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
               Left            =   45
               TabIndex        =   23
               Tag             =   "SortCust"
               ToolTipText     =   "Check To Sort On Invoice Customer Code"
               Top             =   2640
               Width           =   2115
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Show Running Sr   "
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
               Left            =   45
               TabIndex        =   47
               Tag             =   "ShowRunSr"
               ToolTipText     =   "Check To Show Picture"
               Top             =   6870
               Width           =   2115
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Show Distinct Design For Scope   "
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   480
               Index           =   3
               Left            =   45
               TabIndex        =   45
               Tag             =   "ShowDistDsg"
               ToolTipText     =   "Check To Show Distinct Design Scope"
               Top             =   6075
               Width           =   2115
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Show Summary      "
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
               Left            =   45
               TabIndex        =   43
               Tag             =   "ShowSumm"
               ToolTipText     =   "Check To Show Summary"
               Top             =   5775
               Width           =   2115
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Show Cust RmCd    "
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
               Left            =   45
               TabIndex        =   41
               Tag             =   "ShowCustRm"
               ToolTipText     =   "Check To Show Customer Rm Codes"
               Top             =   5490
               Width           =   2115
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Show Cust DsgCd   "
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
               Left            =   45
               TabIndex        =   39
               Tag             =   "ShowCustDsg"
               ToolTipText     =   "Check To Show Customer Design Code"
               Top             =   5205
               Width           =   2115
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Labour Value         "
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
               Left            =   60
               TabIndex        =   37
               Tag             =   "ShowLabVal"
               ToolTipText     =   "Check To Show Labour Value"
               Top             =   4920
               Width           =   2115
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Show Rm Value     "
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
               Left            =   60
               TabIndex        =   35
               Tag             =   "ShowRmVal"
               ToolTipText     =   "Check To Show Rm Value"
               Top             =   4635
               Width           =   2115
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Show Dsg Value     "
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
               Left            =   60
               TabIndex        =   33
               Tag             =   "ShowDsgVal"
               ToolTipText     =   "Check To Show Design Value"
               Top             =   4350
               Width           =   2115
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Show Average Wt  "
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
               TabIndex        =   29
               Tag             =   "ShowAvgWt"
               ToolTipText     =   "Check To Show Average Wts"
               Top             =   3780
               Width           =   2115
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Picture                    "
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
               TabIndex        =   31
               Tag             =   "ShowPic"
               ToolTipText     =   "Check To Show Picture"
               Top             =   4065
               Width           =   2115
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   18
               Left            =   5145
               TabIndex        =   5
               ToolTipText     =   "Enter To Invoice Year"
               Top             =   645
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               MaxLength       =   2
               DataType        =   4
               DataField       =   "UOINYYTO"
               IdName          =   "UOINYYTO"
               CmpStr          =   "IdYy <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   20
               Left            =   2745
               TabIndex        =   2
               ToolTipText     =   "Enter From Invoice Year"
               Top             =   645
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               MaxLength       =   2
               DataType        =   4
               DataField       =   "UOINYYFR"
               IdName          =   "UOINYYFR"
               CmpStr          =   "IdYy >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   43
               Left            =   1965
               TabIndex        =   1
               ToolTipText     =   "Enter From Invoice Tc"
               Top             =   645
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOINTCFR"
               IdName          =   "UOINTCFR"
               CmpStr          =   "IdTc >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   45
               Left            =   3165
               TabIndex        =   3
               ToolTipText     =   "Enter From Invoice Character"
               Top             =   645
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOINCHRFR"
               IdName          =   "UOINCHRFR"
               CmpStr          =   "IdChr >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   46
               Left            =   4365
               TabIndex        =   4
               ToolTipText     =   "Enter To Invoice Tc"
               Top             =   645
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOINTCTO"
               IdName          =   "UOINTCTO"
               CmpStr          =   "IdTc <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   47
               Left            =   5565
               TabIndex        =   6
               ToolTipText     =   "Enter To Invoice Character"
               Top             =   645
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOINCHRTO"
               IdName          =   "UOINCHRTO"
               CmpStr          =   "IdChr <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   26
               Left            =   1965
               TabIndex        =   7
               ToolTipText     =   "Enter From Invoice Number"
               Top             =   930
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOINNOFR"
               IdName          =   "UOINNOFR"
               CmpStr          =   "IdNo >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   29
               Left            =   4365
               TabIndex        =   9
               ToolTipText     =   "Enter To Invoice Number"
               Top             =   930
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOINNOTO"
               IdName          =   "UOINNOTO"
               CmpStr          =   "IdNo <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   0
               Left            =   1965
               TabIndex        =   25
               ToolTipText     =   "Enter Level Of Raw Material Detail To Be Shown"
               Top             =   2925
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
               Index           =   1
               Left            =   1965
               TabIndex        =   26
               ToolTipText     =   "Enter Level Of Labour Detail To Be Shown"
               Top             =   3210
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               DataField       =   "UOLABGRP"
               IdName          =   "UOLABGRP"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   2
               Left            =   2325
               TabIndex        =   30
               ToolTipText     =   "Show Picture (Yes / No)"
               Top             =   4065
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
               Index           =   3
               Left            =   1965
               TabIndex        =   27
               ToolTipText     =   "Specify Whether To Show Actual, Invoice Or Customs Values"
               Top             =   3495
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               DataType        =   4
               DataField       =   "UOINVAL"
               ReCalcParent    =   "UOCURCDFR"
               IdName          =   "UOINVAL"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   4
               Left            =   2325
               TabIndex        =   28
               ToolTipText     =   "Show Average Weights (Yes / No)"
               Top             =   3780
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
               Index           =   17
               Left            =   1965
               TabIndex        =   0
               ToolTipText     =   "Enter Company Code"
               Top             =   360
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               Enabled         =   0   'False
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOCOCDFR"
               IdName          =   "UOCOCDFR"
               CmpStr          =   "IdCoCd ="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   7
               Left            =   4365
               TabIndex        =   12
               ToolTipText     =   "Enter To Export Number"
               Top             =   1215
               Width           =   2415
               _ExtentX        =   4260
               _ExtentY        =   503
               MaxLength       =   20
               DataField       =   "UOINEXPNOTO"
               IdName          =   "UOINEXPNOTO"
               CmpStr          =   "InExpNo <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   9
               Left            =   1965
               TabIndex        =   11
               ToolTipText     =   "Enter From Export Number"
               Top             =   1215
               Width           =   2265
               _ExtentX        =   3995
               _ExtentY        =   503
               MaxLength       =   20
               DataField       =   "UOINEXPNOFR"
               IdName          =   "UOINEXPNOFR"
               CmpStr          =   "InExpNo >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   5
               Left            =   2325
               TabIndex        =   32
               ToolTipText     =   "Show Main Value (Yes / No)"
               Top             =   4350
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
               Index           =   14
               Left            =   2325
               TabIndex        =   36
               ToolTipText     =   "Show Labor Value (Yes / No)"
               Top             =   4920
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
               Index           =   6
               Left            =   1965
               TabIndex        =   20
               ToolTipText     =   "Enter From Invoice Number"
               Top             =   2355
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOCMCDFR"
               IdName          =   "UOCMCDFR"
               CmpStr          =   "InCmCd >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   8
               Left            =   4365
               TabIndex        =   21
               ToolTipText     =   "Enter To Invoice Number"
               Top             =   2355
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOCMCDTO"
               IdName          =   "UOCMCDTO"
               CmpStr          =   "InCmCd <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   10
               Left            =   2325
               TabIndex        =   22
               ToolTipText     =   "Show Value (Yes / No)"
               Top             =   2640
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
               Index           =   19
               Left            =   2325
               TabIndex        =   38
               ToolTipText     =   "Show Cust DsgCd (Yes / No)"
               Top             =   5205
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
               Index           =   22
               Left            =   2325
               TabIndex        =   40
               ToolTipText     =   "Show Cust RmCd (Yes / No)"
               Top             =   5490
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
               Index           =   11
               Left            =   1965
               TabIndex        =   13
               ToolTipText     =   "Enter First Group"
               Top             =   1500
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
               Index           =   12
               Left            =   1965
               TabIndex        =   16
               ToolTipText     =   "Enter Second Group"
               Top             =   1785
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               DataField       =   "UOINGRP1"
               IdName          =   "UOINGRP1"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   40
               Left            =   2325
               TabIndex        =   42
               ToolTipText     =   "Show Summary (Yes / No)"
               Top             =   5775
               Visible         =   0   'False
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               Enabled         =   0   'False
               DataType        =   4
               DataField       =   "UOYN9"
               IdName          =   "UOYN9"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   13
               Left            =   2325
               TabIndex        =   44
               ToolTipText     =   "Show Distinct Design(Y/N)"
               Top             =   6060
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
               Index           =   16
               Left            =   6975
               TabIndex        =   17
               ToolTipText     =   "Final Sorting on Design Code (Y/ N)"
               Top             =   1785
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
               Index           =   21
               Left            =   1965
               TabIndex        =   19
               ToolTipText     =   "Enter Detail Sorting To Be Shown In Report"
               Top             =   2070
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
               Index           =   23
               Left            =   1965
               TabIndex        =   46
               ToolTipText     =   "Enter Currency Code Whose Invoice You Wish To View Or Leave It Blank If You Wish To View All Invoices In Base Currency"
               Top             =   6585
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataType        =   4
               DataField       =   "UOCURCDFR"
               ReCalcOn        =   "UOINVAL"
               IdName          =   "UOCURCDFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   24
               Left            =   6975
               TabIndex        =   14
               ToolTipText     =   "Final Sorting on Design Code (Y/ N)"
               Top             =   1500
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
               Left            =   2325
               TabIndex        =   34
               ToolTipText     =   "Show Rm Value (Yes / No)"
               Top             =   4635
               Visible         =   0   'False
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               Enabled         =   0   'False
               DataType        =   4
               DataField       =   "UOYN14"
               IdName          =   "UOYN14"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   28
               Left            =   2880
               TabIndex        =   8
               ToolTipText     =   "Enter From Invoice Serial"
               Top             =   930
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOIDSRFR"
               IdName          =   "UOIDSRFR"
               CmpStr          =   "IdSr >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   30
               Left            =   5280
               TabIndex        =   10
               ToolTipText     =   "Enter From Invoice Serial"
               Top             =   930
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOIDSRTO"
               IdName          =   "UOIDSRTO"
               CmpStr          =   "IdSr <="
            End
            Begin MSComDlg.CommonDialog Cd1 
               Left            =   12240
               Top             =   8280
               _ExtentX        =   847
               _ExtentY        =   847
               _Version        =   393216
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   15
               Left            =   2325
               TabIndex        =   127
               ToolTipText     =   "Show Running Serial (Y/N)"
               Top             =   6870
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
               Height          =   285
               Index           =   60
               Left            =   2325
               TabIndex        =   128
               ToolTipText     =   "Show Components Details (Yes / No)"
               Top             =   7155
               Visible         =   0   'False
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               Enabled         =   0   'False
               DataType        =   4
               DataField       =   "UOYN23"
               IdName          =   "UOYN23"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   62
               Left            =   2325
               TabIndex        =   129
               ToolTipText     =   "Show Components Details (Yes / No)"
               Top             =   7440
               Visible         =   0   'False
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               Enabled         =   0   'False
               DataType        =   4
               DataField       =   "UOYN24"
               IdName          =   "UOYN24"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   74
               Left            =   7260
               TabIndex        =   51
               ToolTipText     =   "Show Components Details (Yes / No)"
               Top             =   7740
               Visible         =   0   'False
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               Enabled         =   0   'False
               DataType        =   4
               DataField       =   "UOYN28"
               IdName          =   "UOYN28"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   33
               Left            =   6975
               TabIndex        =   143
               ToolTipText     =   "Show Value (Yes / No)"
               Top             =   2640
               Visible         =   0   'False
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               Enabled         =   0   'False
               DataType        =   4
               DataField       =   "UOYN29"
               IdName          =   "UOYN29"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   34
               Left            =   5280
               TabIndex        =   144
               ToolTipText     =   "Show Components Details (Yes / No)"
               Top             =   7440
               Visible         =   0   'False
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               Enabled         =   0   'False
               DataType        =   4
               DataField       =   "UOYN30"
               IdName          =   "UOYN30"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   68
               Left            =   3480
               TabIndex        =   145
               ToolTipText     =   "Show Remarks in  Report? (Y/N)"
               Top             =   8640
               Visible         =   0   'False
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               Enabled         =   0   'False
               DataType        =   4
               DataField       =   "UOYN26"
               IdName          =   "UOYN26"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   35
               Left            =   2640
               TabIndex        =   146
               ToolTipText     =   "Show Bag / Stock Id (Yes / No)"
               Top             =   8400
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
               Index           =   36
               Left            =   6960
               TabIndex        =   150
               Top             =   3120
               Visible         =   0   'False
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               Enabled         =   0   'False
               DataType        =   4
               DataField       =   "UOYN31"
               IdName          =   "UOYN31"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   71
               Left            =   6555
               TabIndex        =   148
               ToolTipText     =   "Enter Picture Path Where To Be Copied"
               Top             =   3480
               Width           =   4740
               _ExtentX        =   8361
               _ExtentY        =   503
               MaxLength       =   100
               DataType        =   4
               IdName          =   "WTARPICPATH"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   37
               Left            =   6960
               TabIndex        =   152
               Top             =   3840
               Visible         =   0   'False
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               Enabled         =   0   'False
               DataType        =   4
               DataField       =   "UOYN32"
               IdName          =   "UOYN32"
            End
            Begin VB.Label LblPicPath 
               BackStyle       =   0  'Transparent
               Caption         =   "Picture Path"
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
               Left            =   4365
               TabIndex        =   151
               Top             =   3480
               Width           =   1425
            End
            Begin VB.Label LblNetWt 
               BackStyle       =   0  'Transparent
               Caption         =   "Show "
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
               TabIndex        =   135
               Top             =   7770
               Width           =   1830
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
               Left            =   2790
               TabIndex        =   126
               Top             =   930
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
               Left            =   5190
               TabIndex        =   125
               Top             =   930
               Width           =   105
            End
            Begin VB.Label LblShwVal 
               BackStyle       =   0  'Transparent
               Caption         =   "Show Design Value"
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
               Left            =   75
               TabIndex        =   103
               ToolTipText     =   "Location"
               Top             =   4350
               Visible         =   0   'False
               Width           =   1995
            End
            Begin VB.Label LblShowGrpDesc 
               BackStyle       =   0  'Transparent
               Caption         =   "Show Grp Comb Desc"
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
               Left            =   4380
               TabIndex        =   102
               ToolTipText     =   "Location"
               Top             =   1500
               Visible         =   0   'False
               Width           =   2085
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Currency"
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
               Index           =   175
               Left            =   75
               TabIndex        =   101
               Top             =   6585
               Width           =   1005
            End
            Begin VB.Label LblDet 
               BackStyle       =   0  'Transparent
               Caption         =   "Detail Sort"
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
               TabIndex        =   100
               Top             =   2070
               Width           =   1875
            End
            Begin VB.Label LblFinalSrt 
               BackStyle       =   0  'Transparent
               Caption         =   "Final Sort on Dsg Cd"
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
               Left            =   4380
               TabIndex        =   99
               ToolTipText     =   "Location"
               Top             =   1785
               Visible         =   0   'False
               Width           =   2085
            End
            Begin VB.Label LblShowRunSr 
               BackStyle       =   0  'Transparent
               Caption         =   "Show Running Sr"
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
               Left            =   60
               TabIndex        =   98
               ToolTipText     =   "Location"
               Top             =   6855
               Visible         =   0   'False
               Width           =   1845
            End
            Begin VB.Label LblDistDsg 
               BackStyle       =   0  'Transparent
               Caption         =   "Show Distinct Design for Scope"
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
               Height          =   525
               Left            =   75
               TabIndex        =   97
               ToolTipText     =   "Location"
               Top             =   6075
               Visible         =   0   'False
               Width           =   1845
            End
            Begin VB.Label LblShowSumm 
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
               Left            =   75
               TabIndex        =   96
               ToolTipText     =   "Location"
               Top             =   5775
               Visible         =   0   'False
               Width           =   1845
            End
            Begin VB.Label LblGrpCombo 
               BackStyle       =   0  'Transparent
               Caption         =   "Grp Combination1"
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
               TabIndex        =   95
               ToolTipText     =   "Location"
               Top             =   1500
               Width           =   1875
            End
            Begin VB.Label LblGrpCombo1 
               BackStyle       =   0  'Transparent
               Caption         =   "Grp Combination2"
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
               TabIndex        =   94
               ToolTipText     =   "Location"
               Top             =   1785
               Width           =   1875
            End
            Begin VB.Label lblCustRm 
               BackStyle       =   0  'Transparent
               Caption         =   "Show Cust RmCd"
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
               Left            =   75
               TabIndex        =   93
               ToolTipText     =   "Location"
               Top             =   5490
               Visible         =   0   'False
               Width           =   1845
            End
            Begin VB.Label Label3 
               BackStyle       =   0  'Transparent
               Caption         =   "Show Cust DsgCd"
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
               Left            =   75
               TabIndex        =   92
               ToolTipText     =   "Location"
               Top             =   5205
               Visible         =   0   'False
               Width           =   1845
            End
            Begin VB.Label LblSortCust 
               BackStyle       =   0  'Transparent
               Caption         =   "Sort on Cust Cd"
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
               Left            =   75
               TabIndex        =   91
               ToolTipText     =   "Sort On Customer Code ? (Y/N)"
               Top             =   2640
               Visible         =   0   'False
               Width           =   1845
            End
            Begin VB.Label LblCustCd 
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
               Left            =   75
               TabIndex        =   88
               ToolTipText     =   "Location"
               Top             =   2355
               Width           =   1845
            End
            Begin VB.Label LblRmVal 
               BackStyle       =   0  'Transparent
               Caption         =   "Show Rm Value"
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
               Left            =   75
               TabIndex        =   87
               Top             =   4635
               Visible         =   0   'False
               Width           =   1845
            End
            Begin VB.Label LblLabVal 
               BackStyle       =   0  'Transparent
               Caption         =   "Labour Value"
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
               Left            =   75
               TabIndex        =   86
               ToolTipText     =   "Location"
               Top             =   4920
               Visible         =   0   'False
               Width           =   1845
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
               Left            =   75
               TabIndex        =   85
               Top             =   1215
               Width           =   1425
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
               Left            =   75
               TabIndex        =   84
               Top             =   360
               Width           =   1335
            End
            Begin VB.Label LblAvgWt 
               BackStyle       =   0  'Transparent
               Caption         =   "Show Avg Wt"
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
               TabIndex        =   83
               Top             =   3780
               Width           =   1845
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
               Left            =   75
               TabIndex        =   82
               Top             =   3495
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
               Left            =   75
               TabIndex        =   81
               Top             =   4065
               Visible         =   0   'False
               Width           =   1845
            End
            Begin VB.Label LblLabGrp 
               BackStyle       =   0  'Transparent
               Caption         =   "Labour Det Level"
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
               TabIndex        =   80
               Top             =   3210
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
               Left            =   75
               TabIndex        =   79
               Top             =   2925
               Width           =   1845
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Inv Tc/Yy/Chr"
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
               Left            =   75
               TabIndex        =   78
               Top             =   645
               Width           =   1725
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Inv No"
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
               Left            =   75
               TabIndex        =   77
               Top             =   930
               Width           =   1605
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
               Left            =   5505
               TabIndex        =   76
               Top             =   645
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
               Left            =   3105
               TabIndex        =   75
               Top             =   645
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
               Left            =   5085
               TabIndex        =   74
               Top             =   645
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
               Left            =   2685
               TabIndex        =   73
               Top             =   645
               Width           =   105
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
               Left            =   4305
               TabIndex        =   72
               Top             =   30
               Width           =   1275
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "From "
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
               TabIndex        =   71
               Top             =   30
               Width           =   1335
            End
         End
      End
   End
   Begin VB.Frame FraChgImg 
      Height          =   1275
      Left            =   6000
      TabIndex        =   140
      Top             =   8280
      Visible         =   0   'False
      Width           =   6135
      Begin MwfCtl.MWCTL_BTN1 CmdChgImgGo 
         Height          =   390
         Left            =   5400
         TabIndex        =   139
         ToolTipText     =   "Analysis Scope Options"
         Top             =   675
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   688
         BackColor       =   -2147483633
         MaskColor       =   8454143
         Caption         =   "Go"
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
      Begin MwfCtl.MWCTL_BTN1 CmdSelFile 
         Height          =   285
         Left            =   4920
         TabIndex        =   138
         ToolTipText     =   "Analysis Scope Options"
         Top             =   720
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   503
         BackColor       =   -2147483633
         MaskColor       =   8454143
         Caption         =   "..."
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
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   32
         Left            =   1200
         TabIndex        =   137
         ToolTipText     =   "Enter Sort Option"
         Top             =   720
         Width           =   3780
         _ExtentX        =   6668
         _ExtentY        =   503
         MaxLength       =   0
         IdName          =   "WXLFILEPATH"
      End
      Begin VB.Label LblMain 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "  Select Report Excel File"
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
         Index           =   1
         Left            =   0
         TabIndex        =   142
         Top             =   240
         Width           =   6180
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Report File"
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
         Left            =   120
         TabIndex        =   141
         Top             =   720
         Width           =   1245
      End
   End
End
Attribute VB_Name = "EmrFrpInvCost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mRep As CRAXDRT.Report
Dim Rep As New EmrRepOrd
Dim Rep1 As New EmrRepOrdCat
Dim moCn As MwfLib.MDOConnection   '*** (Bef speed) bef 24/11/06
Dim oExcel, oBook, oSheet   ' ***** Manali 3.6.0 - Excel Format Report
Dim wErrMsg As String       ' ***** Manali 3.6.0 - Excel Format Report
Dim ws_BaseCurCd As String, wb_InBaseCur As Boolean, ws_CurJoin As String
Dim wChkSel As Integer
Dim wo_rsPic As MDORowSet
Private Sub SetGroupSort()
  'If ADC.MenuCd = UCase("RepInvCat") Then
  ' GRP_REP.Add "Cust Code", "InCmCd", "", "", "", "", "", "", "(Select CmName From CustMst Where CmCtg= 'C' and CmCd= InCmCd) "
  'End If
End Sub
Private Sub adc_setreprecsource()
    '*** (Jen 2.12)
On Error GoTo RepErr
  '*** (Jen 2.12)
'''*********************Emr205******Previous Sql**************************************
'''*** Report Sql for Det & Avg Ord ***
'''    Select OmCoCd, OmYy, OmTc, OmChr, OmNo, OmCmCd, OmDt, OmPoNo, OmLmgCst, OmLmgSal,OmLmpSal, OmLmsSal,OmLmlSal, OmDelDt, OdDelDt
'''    OmHCtCd, OdSr,
'''    ' ' as IdExpOdTc, ' ' as IdExpOdYy, ' ' as IdExpOdChr, 0 as IdExpOdNo, 0 as IdExpOdSr,
'''    OdDmCd, Space(17) as CdOurDmCd, OdDmSz, OdSfx, OdDmId, DmCtg, DmSalCtg, DmTcTyp, OdDmCol, 0 as OdTrayNo, OdPicNm, OdPoNo, OdKt as IdKt, DmUom, OdVaCtg as IdVaCtg, 0.00 as IdGldLs, Space(254) as qRmCtgDesc,
'''    space(254) as qDmCtgDesc, space(254) as qDmSalCtgDesc, space(254) as qDmColDesc,
'''    space(254) as qVaCtgDesc,
'''    OrRmCtg /*as qMainMetRmCtg*/, OdDmCol as OdDmCol3, OdGldAsWt, OdLabAsWt, OdSalPrc,
'''    OdOrdEnt as qOrdQty, OdSalPrc as qSalVal,
'''    OdDmPrdInst, OdCmPrdInst, OdCmStmpInst, OdSalRem, OdSzInst, OmSalRem, OmCmDelTerms, OmCmPayTerms, ' '  as qOdomPoNo,
'''     ' ' as CmLkUpMetLs, DmLsCtg, OdMulBy, ' ' as CmName,
'''    space(4) as qTag, 0 as qMainMet,Space(2) as qPtrYN,
'''    OrSrNo, OrRmCtg RmCtg, OrRmSCtg,space(14) as OrRmCd, RmQw, 0.0000 as OrLn1, ' ' as qRmSzDesc, 0.000 as OrLn2,
'''    0.0000 as OrRmPtr, OrQty, OrWt, OrLmeSal, OrSalRt, OrSetSCd,
'''    OrSetSalRt, OrSetSalVal, OrSalVal, Space(10) as RmRSzRt,
'''     OrAlySalRt, 0 as RmPuritRt,
'''    OlSrNo, OlMcd, OlSCd, PDesc as qLabDesc,
'''    OlQw , OlQty, OlSalRt, OlSalVal,
'''    OctSrNo , OctCd, OctSz, OctCtChr, OctParts, OctRmCd, OctDmCol, OctRmWt, OctDc, OctPrdSeq, OctMainPrtYn, OctDesc
'''    From OrdMst, OrdDsg, OrdRm, OrdLab, OrdCT, DsgMst, Param, RmMst
'''    Where 1 = 2
''  '*** Report Sql for Det & Avg Ord ***

'*** Jay 2.13(CT) ***
'*** Jay 3.3(FG)
'    Select OmCoCd, OmYy, OmTc, OmChr, OmNo, OmCmCd, OmDt, OmPoNo, OmPoDt, OmLmgCst, OmLmgSal,OmLmpSal, OmLmsSal,OmLmlSal, OmDelDt, OdDelDt
'    OmHCtCd, OdSr,
'    ' ' as IdExpOdTc, ' ' as IdExpOdYy, ' ' as IdExpOdChr, 0 as IdExpOdNo, 0 as IdExpOdSr,
'    OdDmCd, Space(17) as CdOurDmCd, OdDmSz, OdSfx, OdDmId, DmCtg, DmSalCtg, DmTcTyp, OdDmCol, 0 as OdTrayNo, OdPicNm, OdPoNo, OdKt as IdKt, DmUom, OdVaCtg as IdVaCtg, 0.00 as IdGldLs, Space(254) as qRmCtgDesc,
'    space(254) as qDmCtgDesc, space(254) as qDmSalCtgDesc, space(254) as qDmColDesc,
'    space(254) as qVaCtgDesc,
'    OrRmCtg /*as qMainMetRmCtg*/, OdDmCol as OdDmCol3, OdGldAsWt, OdLabAsWt, OdSalPrc,
'    OdOrdEnt as qOrdQty, OdSalPrc as qSalVal,
'    OdDmPrdInst, OdCmPrdInst, OdCmStmpInst, OdSalRem, OdSzInst, OmSalRem, OmCmDelTerms, OmCmPayTerms, ' '  as qOdomPoNo,
'    OdDelDt, ' ' as CmLkUpMetLs, DmLsCtg, OdMulBy, ' ' as CmName,
'    space(4) as qTag, 0 as qMainMet,Space(2) as qPtrYN,
'    OrSrNo, OrRmCtg RmCtg, OrRmSCtg,space(14) as OrRmCd, RmQw, 0.0000 as OrLn1, ' ' as qRmSzDesc, 0.000 as OrLn2,
'    0.0000 as OrRmPtr, OrQty, OrWt, OrLmeSal, OrSalRt, OrSetSCd,
'    OrSetSalRt, OrSetSalVal, OrSalVal, Space(10) as RmRSzRt,
'     OrAlySalRt, 0 as RmPuritRt,
'    OlSrNo, OlMcd, OlSCd, PDesc as qLabDesc,
'    OlQw , OlQty, OlSalRt, OlSalVal
'    From OrdMst, OrdDsg, OrdRm, OrdLab, DsgMst, Param, RmMst
'    Where 1 = 2
'*** Jay 2.13(CT) ***

  '*** Sub Report Sql for Det & Avg Ord ***
  '  Select OdYy, OdTc, OdChr, OdNo, OdSr, DmCtg, PDesc as qDesc, OdOrdEnt as qOrdQty,
  '  OdSalPrc as qSalVal From OrdDsg, DsgMst, Param Where 1= 2
  '*** Sub Report Sql for Det & Avg Ord ***
  
  '*******************Geeta****************************************************
  '***************Emr206**********************************************************
  '*** Report Sql for Catalogue ***
  ' Zubin 212 (Sales Ctg and qDmSalCtgDesc added)
  '*** Jay 3.3(FG)
'    Select OmCoCd, OmYy, OmTc, OmChr, OmNo, OmCmCd, OmDt, OmPoNo, OdSr,Space(17) as CdOurDmCd, OdDmCd, OdDmSz, OdSfx,OdDmId,
'    DmCtg, DmSalCtg, DmTcTyp, OdDmCol, OdKt as IdKt, OdVaCtg as IdVaCtg, DmUom, 0.00 as IdGldLs,
'    Space(254) as qRmCtgDesc,
'    space(254) as qDmCtgDesc, space(254) as qDmSalCtgDesc, space(254) as qDmColDesc,space(254) as qVaCtgDesc,
'    OrRmctg /*as qMainMetRmCtg*/, OdGldAsWt, OdLabAsWt, OdSalPrc, OdOrdEnt as qOrdQty, PDesc as qColDesc,
'    OdGldAsWt as qGrsWt From OrdMst,OrdRm, OrdDsg, DsgMst, Param Where 1 = 2

  '  Select OdYy, OdTc, OdChr, OdNo, OdSr, OrMainMet,
  '  OrRmCtg, 0 as qTag, OrRmSCtg, Space(14) as OrRmCd, OrQty, OrWt
  '  From OrdDsg, OrdRm Where 1 = 2
  '*** Report Sql for Catalogue ***
  
  If UCase(adc.MenuCd) = UCase("RepAvgInv") Or UCase(adc.MenuCd) = UCase("RepAvgInvCTB") Then
    adc("UoRmGrp") = "SR"
    adc("UoLabGrp") = "SR"
  End If
  
  Dim wExcelOrdBy As String
  Dim grpflds As String
  Dim wrepcnd As String, wCnd As String, wGldAsWt As String, wLabAsWt As String, wIdFOB As String, wIrRmWt As String
  Dim wIrRmRt As String, wIrSetRt As String, wIrSetVal As String, wIrRmVal As String
  Dim wiLRt As String, wiLVal As String, wCommStr As String, wRmStr As String, wRmGrp As String
  Dim wRmCommStr As String, wLabStr As String, wLabGrp As String, wLabCommStr As String
  Dim wSqlStrg As String, wDsgCtgStr As String, wRmDet As String
  Dim wRmDetSql As String
  
  Dim wiLQw As String     '*** (Jen 2.14 Next)
  
  
  '*************************** ZUBIN **************************
  ' 08th Dec 2003, EMR206
  Dim wOurDsgCd As String, wOurRmCd As String
  '*************************** ZUBIN **************************
  
  '*** Jay 2.13(CT) ***
  Dim wOrdCTCommStr As String
  '*** Jay 2.13(CT) ***

  ' ***** Manali 3.03 - 29/06/08
  Dim wIrLmeRt As String
  ' ***** Manali 3.03 - 29/06/08
  
  '*** Manali 3.03 - 07/07/08 - RmSzDesc
  Dim wRmSzDesc As String
  '*** Manali 3.03 - 07/07/08 - RmSzDesc
  
  Dim ws_OrdBy As String, wAlySalRt As String, wRmPurity As String
  
  Set Rep = Nothing
  Set Rep1 = Nothing

  Select Case UCase(adc.MenuCd)
  Case UCase("RepAvgInv"), UCase("RepDetInv"), UCase("RepAvgInvCTB"), UCase("RepDetInvCTB")
    Set adc.RepSource = Rep
    Set mRep = Rep
  Case UCase("RepInvCat"), UCase("RepInvCTBCat")
    Set adc.RepSource = Rep1
    Set mRep = Rep1
  End Select
  'Set adc.RepSource = Rep
  
  Dim wGrpStr As String
    
  Call DispCoNm
  Call SetMwName(mRep)   ' ***** Manali 3.6.1 - Microway footer
  If adc.MenuCd = UCase("RepInvCat") Then
    If adc("UoYN5") = "Y" Then
      GRP_REP.SetFormula mRep, "wGrpOpt", "{rdo.OmCmCd}"
      wExcelOrdBy = "InCmCd"
    Else
      GRP_REP.SetFormula mRep, "wGrpOpt", ""
    End If
  End If
  
  '6.1_27 if copy picture option is selected and path is blank then raise error message
  If UCase(adc.MenuCd) = UCase("RepDetInv") Or UCase(adc.MenuCd) = UCase("RepInvCat") Then
    If adc("UOYN31") = "Y" And adc("wTarPicPath") = "" Then _
      DispMsg "Picture Path Should Not Be Blank ", etError: Exit Sub
    'if target picture path is not valid then it will raise error
    If adc("UOYN31") = "Y" And adc("wTarPicPath") <> "" Then
      If GetAttr(adc("wTarPicPath")) <> vbDirectory Then GoTo RepErr
    End If
  End If
 
  ' Zubin 213
'  Dim ws_BaseCurCd As String, wb_InBaseCur As Boolean, ws_CurJoin As String
  ws_BaseCurCd = moCn.GetFldVal("Select HBaseCurCd from Head where HCd= '" + ctSelfCmCd + "' and " + _
                  "HCoCd= '" + adc("UoCoCdFr") + "'")
  If adc("UoCurCdFr") = "" Then
    wb_InBaseCur = True
    mRep.TxtValIn.SetText "Value In " + ws_BaseCurCd
    '6.1B-15 label2 & 3are added
    If UCase(adc.MenuCd) = UCase("RepInvCat") Or UCase(adc.MenuCd) = UCase("RepInvCTBCat") Then
      mRep.TxtUnitPrcCur.SetText "Price" + " (" + ws_BaseCurCd + ")"    ' **** Manali 3.5.0 - 14/11/08
      mRep.TxtUnitPrcCur2.SetText "Price" + " (" + ws_BaseCurCd + ")"    ' **** Manali 3.5.0 - 14/11/08
      mRep.TxtUnitPrcCur3.SetText "Price" + " (" + ws_BaseCurCd + ")"    ' **** Manali 3.5.0 - 14/11/08
    End If
      'mRep.TxtUnitPrcCur.SetText "Unit Price" + vbCrLf + "(" + ws_BaseCurCd + ")"
  Else
    wb_InBaseCur = False
    mRep.TxtValIn.SetText "Value In " + adc("UoCurCdFr")
    '6.1B-15 label2 & 3are added
    If UCase(adc.MenuCd) = UCase("RepInvCat") Or UCase(adc.MenuCd) = UCase("RepInvCTBCat") Then
      mRep.TxtUnitPrcCur.SetText "Price" + " (" + adc("UoCurCdFr") + ")"    ' **** Manali 3.5.0 - 14/11/08
      mRep.TxtUnitPrcCur2.SetText "Price" + " (" + adc("UoCurCdFr") + ")"    ' **** Manali 3.5.0 - 14/11/08
      mRep.TxtUnitPrcCur3.SetText "Price" + " (" + adc("UoCurCdFr") + ")"    ' **** Manali 3.5.0 - 14/11/08
    End If
      'mRep.TxtUnitPrcCur.SetText "Unit Price" + vbCrLf + "(" + adc("UoCurCdFr") + ")"
  End If
    
  ' ***** Manali 3.6.0 - Join Cond removed from  [adc("UoCurCdFr") <> ""]
  If UCase(adc.MenuCd) = UCase("RepDetInvCTB") Or UCase(adc.MenuCd) = UCase("RepAvgInvCTB") Or _
        UCase(adc.MenuCd) = UCase("RepInvCTBCat") Then
    ws_CurJoin = " Join CustMst On CmCtg= 'T' and CmCd= InCmCd "
  Else
    ws_CurJoin = " Join CustMst On CmCtg= 'C' and CmCd= InCmCd "
  End If
  ' ***** Manali 3.6.0 - Join Cond removed from  [adc("UoCurCdFr") = ""]
  
    wrepcnd = adc.RepCond
    ' Zubin 213
    wCnd = " where (Select PValue from Param where PTyp= 'TC' and PMCd= IdTc) = '" + _
           IIF(UCase(adc.MenuCd) = UCase("RepAvgInv") Or UCase(adc.MenuCd) = UCase("RepDetInv") Or _
           UCase(adc.MenuCd) = UCase("RepInvCat"), "IN", "CTB") + "' " + _
           IIF(wb_InBaseCur = False, " and CmCurCd= '" + adc("UoCurCdFr") + "'", "") + IIF(gs_Partition = ctCurrPrtn, " And InPrtKey='" + ctCurrPrtn + "' ", "")
           
    wCnd = wCnd + IIF(wrepcnd <> "", " and " + wrepcnd, "")
               
    Rep.FldOmDelDt.Suppress = True
    Rep.TxtSlash.Suppress = True
  '  Rep.SecRemDet.Suppress = True
  ' urmi 2.06 rem n OrdDet
   ' If adc("UOYN26") = "N" Then
        Rep.SecRemDet1.Suppress = True
        Rep.SecRemDet2.Suppress = True
        Rep.SecRemDet3.Suppress = True
        Rep.SecRemDet4.Suppress = True
        Rep.SecRemDet5.Suppress = True
        Rep.SecOmRemDet1.Suppress = True
        Rep.SecOmRemDet2.Suppress = True
        Rep.SecOmRemDet3.Suppress = True
    'End If
    ''wCommStr = "Select " + grpflds + ", "
    'UoYn2 Show val y/ n
    'Yn4 Show Lab Val y/n
    
    '*** Manali 3.03 - 07/07/08 - RmSzDesc
    wRmSzDesc = "case when (Select PDesc225 from Param where PTyp= 'RMSCTG' and PMCd= IsNull(Max(RisRmCtg), '') and PSCd= max(RmSCtg))= 'Y' then IsNull(Max(RisRmSzDesc), '') Else '' End "
    '*** Manali 3.03 - 07/07/08 - RmSzDesc
    
    If (adc("UoYN2")) = "Y" Or UCase(adc.MenuCd) = UCase("RepInvCat") Or UCase(adc.MenuCd) = UCase("RepInvCTBCat") Then
      If UCase(adc("UoInVal")) = "A" Then
        wGldAsWt = "IdAGldAsWt"
        wLabAsWt = "IdALabAsWt"   '****** Sachin 2.13.0 - 27-05-2006 ******
        wIdFOB = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then IdAVal/InCnvFct else IdAVal*InCnvFct end) ", "IdAVal")
        wIrRmWt = "IrRmAWt"
        ' ***** Manali 3.03 - 29/06/08 - LmeRt
        ' ***** Manali 3.6.0
        wIrLmeRt = " (case IrRmCtg when 'G' Then OmLmgCst when 'P' then OmLmpCst when 'S' then OmLmsCst when 'L' then OmLmlCst Else 0 End)  "
        ' ***Bef 3.6.0
'''        wIrLmeRt = " (case  OdIGldRtEqOrd  when 'Y' then (case IrRmCtg when 'G' Then OmLmgCst when 'P' then OmLmpCst when 'S' then OmLmsCst when 'L' then OmLmlCst Else 0 End)  " + _
'''                   " When 'N' then (case IrRmCtg when 'G' Then InLmgCst when 'P' then InLmpCst when 'S' then InLmsCst when 'L' then InLmlCst Else 0 End) " + _
'''                   " Else 0 end) "
        wIrRmRt = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then IrRmARt/InCnvFct else IrRmARt*InCnvFct end) ", "IrRmARt")
        wIrSetRt = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then IrSetARt/InCnvFct else IrSetARt*InCnvFct end) ", "IrSetARt")
        ' Zubin 212 (Before 01/12/05)
        'wIrSetVal = "IrSetAVal"
        '*** Jay 2.14Next [Show Rm Value]
        wIrRmVal = IIF(adc("UoYN14") = "Y", IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then IrRmAVal/InCnvFct else IrRmAVal*InCnvFct end) ", "IrRmAVal"), "0")
        '*** Jay 2.14Next [Show Rm Value]
        ' Zubin 212 (Before 01/12/05)
        'wiLRt = "iLARt"
        'wiLVal = "iLAVal"
        ' Zubin 212
        wiLQw = "iLCstQw"   '*** (Jen 2.14 Next)
        If UCase(adc("UoYN4")) = "Y" Or UCase(adc.MenuCd) = UCase("RepInvCat") Or UCase(adc.MenuCd) = UCase("RepInvCTBCat") Then
          wiLRt = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then iLARt/InCnvFct else iLARt*InCnvFct end) ", "iLARt")
          wiLVal = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then iLAVal/InCnvFct else iLAVal*InCnvFct end) ", "iLAVal")
          wIrSetVal = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then IrSetAVal/InCnvFct else IrSetAVal*InCnvFct end) ", "IrSetAVal")
        Else
          wiLRt = "0": wiLVal = "0": wIrSetVal = "0"
        End If
        ' Zubin 212
        ' ***** Manali 3.6.0 - 04/11/09 - AlySalRt, RmPurity added for Excel
        wAlySalRt = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then OrAlyCstRt/InCnvFct else OrAlyCstRt*InCnvFct end) ", "OrAlyCstRt")
        wRmPurity = "RmPurityWt"
        ' ***** Manali 3.6.0 - 04/11/09 - AlySalRt, RmPurity added for Excel
        '6.1B-15 PerPc/InvPc is added to invcat report
        If UCase(adc.MenuCd) = UCase("RepInvCat") Then
          mRep.TxtInvValues.SetText "(Actual Values: " + IIF(adc("UoYN32") = "Y", "Per Pc", "Inv Pc") + ")"
        Else
          mRep.TxtInvValues.SetText "(Actual Values)"
        End If
      '******* urmi Invoice RmStkRt
      ElseIf UCase(adc("UoInVal")) = "S" Then
        wGldAsWt = "IdAGldAsWt"
        wLabAsWt = "IdALabAsWt"     '****** Sachin 2.13.0 - 27-05-2006 ******
        wIdFOB = "0"
        wIrRmWt = "IrRmAWt"
        ' ***** Manali 3.03 - 29/06/08 - LmeRt
        wIrLmeRt = "0.00 "
        wIrRmRt = "IrStkRt"
        wIrSetRt = " (Case when InMulDiv= 'M' then IrSetARt/InCnvFct else IrSetARt*InCnvFct end) "
        wIrSetVal = " (Case when InMulDiv= 'M' then IrSetAVal/InCnvFct else IrSetAVal*InCnvFct end) "
        '*** Jay 2.14Next *** [Show Rm Value]
        wIrRmVal = IIF(adc("UoYN14") = "Y", "IrRmAWt * IrStkRt", "0")
        '*** Jay 2.14Next *** [Show Rm Value]
        wiLQw = "iLCstQw" '*** (Jen 2.14 Next)
        wiLRt = " (Case when InMulDiv= 'M' then iLARt/InCnvFct else iLARt*InCnvFct end) "
        wiLVal = " (Case when InMulDiv= 'M' then iLAVal/InCnvFct else iLAVal*InCnvFct end) "
        ' Zubin 212
        If UCase(adc("UoYN4")) = "N" Then
          Rep.Fldg2_LabVal.Suppress = True
          Rep.FldwRmSetRt.Suppress = True
          Rep.FldwRmSetVal.Suppress = True
          Rep.FldOlSalRt.Suppress = True
          Rep.FldOlSalVal.Suppress = True
        End If
        ' Zubin 212
        ' ***** Manali 3.6.0 - 04/11/09 - AlySalRt, RmPurity added for Excel
        wAlySalRt = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then OrAlyCstRt/InCnvFct else OrAlyCstRt*InCnvFct end) ", "OrAlyCstRt")
        wRmPurity = "RmPurityWt"
        ' ***** Manali 3.6.0 - 04/11/09 - AlySalRt, RmPurity added for Excel
        '6.1B-15 PerPc/InvPc is added to invcat report
        If UCase(adc.MenuCd) = UCase("RepInvCat") Then
          mRep.TxtInvValues.SetText "(Stock Values: " + IIF(adc("UoYN32") = "Y", "Per Pc", "Inv Pc") + ")"
        Else
          mRep.TxtInvValues.SetText "(Stock Values)"
        End If
        
        If UCase(adc.MenuCd) <> UCase("RepInvCat") And UCase(adc.MenuCd) <> UCase("RepInvCTBCat") Then
          ' Zubin 213 (Error rectified)
          '*** (Bef 2.14 Next Patch 2)
          'GRP_REP.SetFormula mRep, "wOdSalPrc", "if {rdo.qOrdQty} > 0 then " + _
          '            "(Sum({rdo.OrSalVal}, {rdo.OdSr})+{@g2_labVal})/{rdo.qOrdQty} else 0"
          'GRP_REP.SetFormula mRep, "wSalVal", "Sum({rdo.OrSalVal}, {rdo.OdSr})+{@g2_labVal}"
          '*** (Bef 2.14 Next Patch 2)
          
          '*** (Jen 2.14 Next Patch 2) ramram
'          GRP_REP.SetFormula mRep, "wOdSalPrc", "if {rdo.qOrdQty} > 0 then " + _
                      "(Sum({rdo.OrSalVal}, {@OdSr})+{@g2_labVal})/{rdo.qOrdQty} else 0"
 '         GRP_REP.SetFormula mRep, "wSalVal", "Sum({rdo.OrSalVal}, {@OdSr})+{@g2_labVal}"
          '*** (Jen 2.14 Next Patch 2)ramram
        End If
      '**************
      ElseIf UCase(adc("UoInVal")) = "I" Then
        wGldAsWt = "IdiGldAsWt"
        wLabAsWt = "IdILabAsWt"   '****** Sachin 2.13.0 - 27-05-2006 ******
        wIdFOB = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then IdIFob/InCnvFct else IdIFob*InCnvFct end) ", "IdIFob")
        wIrRmWt = "IrRmIWt"
        ' ***** Manali 3.03 - 29/06/08 - LmeRt
        wIrLmeRt = " (case  OdIGldRtEqOrd  when 'Y' then OrLmeSal  " + _
                   " When 'N' then (case IrRmCtg when 'G' Then InLmgSal when 'P' then InLmpSal when 'S' then InLmsSal when 'L' then InLmlSal Else 0 End) " + _
                   " Else 0 end) "
                   
        wIrRmRt = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then IrRmIRt/InCnvFct else IrRmIRt*InCnvFct end) ", "IrRmIRt")
        wIrSetRt = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then IrSetIRt/InCnvFct else IrSetIRt*InCnvFct end) ", "IrSetIRt")
        ' Zubin 212 (Before 01/12/05)
        'wIrSetVal = "IrSetIVal"
        '*** Jay 2.14Next *** [Show Rm Value]
        wIrRmVal = IIF(adc("UoYN14") = "Y", IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then IrRmIVal/InCnvFct else IrRmIVal*InCnvFct end) ", "IrRmIVal"), "0")
        '*** Jay 2.14Next *** [Show Rm Value]
        ' Zubin 212 (Before 01/12/05)
        'wiLRt = "iLiRt"
        'wiLVal = "iLiVal"
        ' Zubin 212
        wiLQw = "iLQw"  '*** (Jen 2.14 Next)
        If UCase(adc("UoYN4")) = "Y" Or UCase(adc.MenuCd) = UCase("RepInvCat") Or UCase(adc.MenuCd) = UCase("RepInvCTBCat") Then
          wiLRt = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then iLiRt/InCnvFct else iLiRt*InCnvFct end) ", "iLiRt")
          wiLVal = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then iLiVal/InCnvFct else iLiVal*InCnvFct end) ", "iLiVal")
          wIrSetVal = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then IrSetIVal/InCnvFct else IrSetIVal*InCnvFct end) ", "IrSetIVal")
        Else
          wiLRt = "0": wiLVal = "0": wIrSetVal = "0"
        End If
        ' Zubin 212
        ' ***** Manali 3.6.0 - 04/11/09 - AlySalRt, RmPurity added for Excel
        wAlySalRt = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then IrAlySalRt/InCnvFct else IrAlySalRt*InCnvFct end) ", "IrAlySalRt")
        wRmPurity = "RmPurityRt"
        ' ***** Manali 3.6.0 - 04/11/09 - AlySalRt, RmPurity added for Excel
        '6.1B-15 PerPc/InvPc is added to invcat report
        If UCase(adc.MenuCd) = UCase("RepInvCat") Then
          mRep.TxtInvValues.SetText "(Invoice Values: " + IIF(adc("UoYN32") = "Y", "Per Pc", "Inv Pc") + ")"
        Else
          mRep.TxtInvValues.SetText "(Invoice Values)"
        End If
      ElseIf UCase(adc("UoInVal")) = "Z" Then
        wGldAsWt = "IdZGldAsWt"
        wLabAsWt = "IdZLabAsWt"   '****** Sachin 2.13.0 - 27-05-2006 ******
        wIdFOB = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then IdZFob/InCnvFct else IdZFob*InCnvFct end) ", "IdZFob")
        wIrRmWt = "IrRmZWt"
        ' ***** Manali 3.03 - 29/06/08 - LmeRt
        wIrLmeRt = " (case  OdIGldRtEqOrd  when 'Y' then OrLmeSal  " + _
                   " When 'N' then (case IrRmCtg when 'G' Then InLmgSal when 'P' then InLmpSal when 'S' then InLmsSal when 'L' then InLmlSal Else 0 End) " + _
                   " Else 0 end) "
        wIrRmRt = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then IrRmZRt/InCnvFct else IrRmZRt*InCnvFct end) ", "IrRmZRt")
        wIrSetRt = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then IrSetZRt/InCnvFct else IrSetZRt*InCnvFct end) ", "IrSetZRt")
        ' Zubin 212 (Before 01/12/05)
        'wIrSetVal = "IrSetZVal"
        '*** Jay 2.14Next *** [Show Rm Value]
        wIrRmVal = IIF(adc("UoYN14") = "Y", IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then IrRmZVal/InCnvFct else IrRmZVal*InCnvFct end) ", "IrRmZVal"), "0")
        '*** Jay 2.14Next *** [Show Rm Value]
        ' Zubin 212 (Before 01/12/05)
        'wiLRt = "iLZRt"
        'wiLVal = "iLZVal"
        ' Zubin 212
        
        wiLQw = "iLQw"  '*** (Jen 2.14 Next)
        If UCase(adc("UoYN4")) = "Y" Or UCase(adc.MenuCd) = UCase("RepInvCat") Or UCase(adc.MenuCd) = UCase("RepInvCTBCat") Then
          wiLRt = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then iLZRt/InCnvFct else iLZRt*InCnvFct end) ", "iLZRt")
          wiLVal = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then iLZVal/InCnvFct else iLZVal*InCnvFct end) ", "iLZVal")
          wIrSetVal = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then IrSetZVal/InCnvFct else IrSetZVal*InCnvFct end) ", "IrSetZVal")
        Else
          wiLRt = "0": wiLVal = "0": wIrSetVal = "0"
        End If
        ' Zubin 212
        ' ***** Manali 3.6.0 - 04/11/09 - AlySalRt, RmPurity added for Excel
        wAlySalRt = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then IrAlySalRt/InCnvFct else IrAlySalRt*InCnvFct end) ", "IrAlySalRt")
        wRmPurity = "RmPurityZ"
        ' ***** Manali 3.6.0 - 04/11/09 - AlySalRt, RmPurity added for Excel
        '6.1B-15 PerPc/InvPc is added to invcat report
        If UCase(adc.MenuCd) = UCase("RepInvCat") Then
          mRep.TxtInvValues.SetText "(Custom Values: " + IIF(adc("UoYN32") = "Y", "Per Pc", "Inv Pc") + ")"
        Else
          mRep.TxtInvValues.SetText "(Custom Values)"
        End If
      End If
    ElseIf (adc("UoYN2")) = "N" And UCase(adc.MenuCd) <> UCase("RepInvCat") Then
      wIdFOB = "0"
      '*** Jay 2.14Next *** [Show Rm Value]
      'wIrRmVal = "0"
      If adc("UoYN14") = "Y" Then
        Select Case UCase(adc("UoInVal"))
          Case Is = UCase("A")
            wIrRmVal = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then IrRmAVal/InCnvFct else IrRmAVal*InCnvFct end) ", "IrRmAVal")
          Case Is = UCase("S")
            wIrRmVal = "IrRmAWt * IrStkRt"
          Case Is = UCase("I")
            wIrRmVal = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then IrRmIVal/InCnvFct else IrRmIVal*InCnvFct end) ", "IrRmIVal")
          Case Is = UCase("Z")
            wIrRmVal = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then IrRmZVal/InCnvFct else IrRmZVal*InCnvFct end) ", "IrRmZVal")
        End Select
      ElseIf adc("UoYN14") = "N" Then
        wIrRmVal = "0"
      End If
      '*** Jay 2.14Next *** [Show Rm Value]
      mRep.TxtInvValues.SetText ""
        
      If UCase(adc("UoInVal")) = "A" Then
        wGldAsWt = "IdAGldAsWt"
        wLabAsWt = "IdALabAsWt"   '****** Sachin 2.13.0 - 27-05-2006 ******
        wIrRmWt = "IrRmAWt"
        ' ***** Manali 3.03 - 29/06/08 - LmeRt
        ' ***** Manali 3.6.0
        wIrLmeRt = " (case IrRmCtg when 'G' Then OmLmgCst when 'P' then OmLmpCst when 'S' then OmLmsCst when 'L' then OmLmlCst Else 0 End)  "
        ' ***Bef 3.6.0
'''        wIrLmeRt = " (case  OdIGldRtEqOrd  when 'Y' then (case IrRmCtg when 'G' Then OmLmgCst when 'P' then OmLmpCst when 'S' then OmLmsCst when 'L' then OmLmlCst Else 0 End)  " + _
'''                   " When 'N' then (case IrRmCtg when 'G' Then InLmgCst when 'P' then InLmpCst when 'S' then InLmsCst when 'L' then InLmlCst Else 0 End) " + _
'''                   " Else 0 end) "
        wIrRmRt = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then IrRmARt/InCnvFct else IrRmARt*InCnvFct end) ", "IrRmARt")
        wIrSetRt = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then IrSetARt/InCnvFct else IrSetARt*InCnvFct end) ", "IrSetARt")
        wiLRt = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then iLARt/InCnvFct else iLARt*InCnvFct end) ", "iLARt")
        wiLQw = "iLCstQw" '*** (Jen 2.14 Next)
        wAlySalRt = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then OrAlyCstRt/InCnvFct else OrAlyCstRt*InCnvFct end) ", "OrAlyCstRt")
        ' ***** Manali 3.6.0 - 04/11/09 - AlySalRt, RmPurity added for Excel
        wRmPurity = "RmPurityWt"
        mRep.TxtInvValues.SetText "(Actual Values)"
        ' ***** Manali 3.6.0 - 04/11/09 - AlySalRt, RmPurity added for Excel
        
      '******* urmi Invoice RmStkRt
      ElseIf UCase(adc("UoInVal")) = "S" Then
        wGldAsWt = "IdAGldAsWt"
        wLabAsWt = "IdALabAsWt"
        wIrRmWt = "IrRmAWt"
        ' ***** Manali 3.03 - 29/06/08 - LmeRt
        'sv.24 In scope if Value As selected as S then giving error.
        'max removed from string value of wIrLmeRt
        wIrLmeRt = " (case  (OdIGldRtEqOrd)  when 'Y' then (OrLmeSal)  " + _
                   " When 'N' then (case (IrRmCtg) when 'G' Then (InLmgSal) when 'P' then (InLmpSal) when 'S' then (InLmsSal) when 'L' then (InLmlSal) Else 0 End) " + _
                   " Else 0 end) "
        wIrRmRt = "IrStkRt"
        wIrSetRt = " (Case when InMulDiv= 'M' then IrSetARt/InCnvFct else IrSetARt*InCnvFct end) "
        wiLRt = " (Case when InMulDiv= 'M' then iLARt/InCnvFct else iLARt*InCnvFct end) "
        wiLQw = "iLCstQw" '*** (Jen 2.14 Next)
        ' ***** Manali 3.6.0 - 04/11/09 - AlySalRt, RmPurity added for Excel
        wAlySalRt = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then OrAlyCstRt/InCnvFct else OrAlyCstRt*InCnvFct end) ", "OrAlyCstRt")
        wRmPurity = "RmPurityWt"
        ' ***** Manali 3.6.0 - 04/11/09 - AlySalRt, RmPurity added for Excel
        
        mRep.TxtInvValues.SetText "(Stock Values)"
      '*****************************
      ElseIf UCase(adc("UoInVal")) = "I" Then
        wGldAsWt = "IdiGldAsWt"
        wLabAsWt = "IdILabAsWt"   '****** Sachin 2.13.0 - 27-05-2006 ******
        wIrRmWt = "IrRmIWt"
        ' ***** Manali 3.03 - 29/06/08 - LmeRt
        wIrLmeRt = " (case  OdIGldRtEqOrd  when 'Y' then OrLmeSal  " + _
                   " When 'N' then (case IrRmCtg when 'G' Then InLmgSal when 'P' then InLmpSal when 'S' then InLmsSal when 'L' then InLmlSal Else 0 End) " + _
                   " Else 0 end) "
        wIrRmRt = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then IrRmIRt/InCnvFct else IrRmIRt*InCnvFct end) ", "IrRmIRt")
        wIrSetRt = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then IrSetIRt/InCnvFct else IrSetIRt*InCnvFct end) ", "IrSetIRt")
        wiLRt = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then iLiRt/InCnvFct else iLiRt*InCnvFct end) ", "iLiRt")
        wiLQw = "iLQw"  '*** (Jen 2.14 Next)
        ' ***** Manali 3.6.0 - 04/11/09 - AlySalRt, RmPurity added for Excel
        wAlySalRt = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then IrAlySalRt/InCnvFct else IrAlySalRt*InCnvFct end) ", "IrAlySalRt")
        wRmPurity = "RmPurityRt"
        ' ***** Manali 3.6.0 - 04/11/09 - AlySalRt, RmPurity added for Excel
        
        mRep.TxtInvValues.SetText "(Invoice Values)"
      
      ElseIf UCase(adc("UoInVal")) = "Z" Then
        wGldAsWt = "IdZGldAsWt"
        wLabAsWt = "IdZLabAsWt"   '****** Sachin 2.13.0 - 27-05-2006 ******
        wIrRmWt = "IrRmZWt"
        ' ***** Manali 3.03 - 29/06/08 - LmeRt
        wIrLmeRt = " (case  OdIGldRtEqOrd  when 'Y' then OrLmeSal  " + _
                   " When 'N' then (case IrRmCtg when 'G' Then InLmgSal when 'P' then InLmpSal when 'S' then InLmsSal when 'L' then InLmlSal Else 0 End) " + _
                   " Else 0 end) "
        wIrRmRt = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then IrRmZRt/InCnvFct else IrRmZRt*InCnvFct end) ", "IrRmZRt")
        wIrSetRt = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then IrSetZRt/InCnvFct else IrSetZRt*InCnvFct end) ", "IrSetZRt")
        wiLRt = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then iLZRt/InCnvFct else iLZRt*InCnvFct end) ", "iLZRt")
        wiLQw = "iLQw"  '*** (Jen 2.14 Next)
        ' ***** Manali 3.6.0 - 04/11/09 - AlySalRt, RmPurity added for Excel
        wAlySalRt = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then IrAlySalRt/InCnvFct else IrAlySalRt*InCnvFct end) ", "IrAlySalRt")
        wRmPurity = "RmPurityZ"
        ' ***** Manali 3.6.0 - 04/11/09 - AlySalRt, RmPurity added for Excel
        
        mRep.TxtInvValues.SetText "(Custom Values)"
      End If
      
      If (adc("UoYN4")) = "Y" Then
        If UCase(adc("UoInVal")) = "A" Or UCase(adc("UoInVal")) = "S" Then
          wIrSetVal = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then IrSetAVal/InCnvFct else IrSetAVal*InCnvFct end) ", "IrSetAVal")
          wiLVal = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then iLAVal/InCnvFct else iLAVal*InCnvFct end) ", "iLAVal")
          wiLRt = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then iLARt/InCnvFct else iLARt*InCnvFct end) ", "iLARt")
          'wiLQw = "iLCstQw"
        ElseIf UCase(adc("UoInVal")) = "I" Then
          wIrSetVal = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then IrSetIVal/InCnvFct else IrSetIVal*InCnvFct end) ", "IrSetIVal")
          wiLVal = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then iLiVal/InCnvFct else iLiVal*InCnvFct end) ", "iLiVal")
          wiLRt = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then iLiRt/InCnvFct else iLiRt*InCnvFct end) ", "iLiRt")
          'wiLQw = "iLQw"
        ElseIf UCase(adc("UoInVal")) = "Z" Then
          wIrSetVal = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then IrSetZVal/InCnvFct else IrSetZVal*InCnvFct end) ", "IrSetZVal")
          wiLVal = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then iLZVal/InCnvFct else iLZVal*InCnvFct end) ", "iLZVal")
          wiLRt = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then iLZRt/InCnvFct else iLZRt*InCnvFct end) ", "iLZRt")
          'wiLQw = "iLQw"
        End If
      Else
          wIrSetVal = "0"
          wiLVal = "0"
          wiLRt = "0"
          'wiLQw = "''"
      End If
    End If
     
    If adc("UoYN2") = "Y" Then
        ' ***** Manali 3.9.0 - Lme Selection
        Dim wLme1 As String, wLme2 As String
        wLme1 = adc("UoLme1"): wLme2 = adc("UoLme2")
        If wLme1 <> "" Then
          Rep.TxtLmg.SetText wLme1 + " Sales"
          GRP_REP.SetFormula Rep, "wLme1", "{rdo.Om" + wLme1 + "Sal}"
        Else
          Rep.TxtLmg.Suppress = True
        End If
        If wLme2 <> "" Then
          Rep.TxtLmp.SetText wLme2 + " Sales"
          GRP_REP.SetFormula Rep, "wLme2", "{rdo.Om" + wLme2 + "Sal}"
        Else
          Rep.TxtLmp.Suppress = True
        End If
        ' ***** Manali 3.9.0 - Lme Selection
        ' ***** Manali 3.10.0 - 02/04/12 - When No Lme Selected G,P Displayed (Original Behaviour)
        If wLme1 = "" And wLme2 = "" Then
          Rep.TxtLmg.Suppress = False: Rep.TxtLmp.Suppress = False
          GRP_REP.SetFormula Rep, "wLme1", "{rdo.OmLmgSal}"
          GRP_REP.SetFormula Rep, "wLme2", "{rdo.OmLmPSal}"
        End If
        ' ***** Manali 3.10.0 - 02/04/12 - When No Lme Selected G,P Displayed (Original Behaviour)
        
        ' Bef 3.9.0
''      Rep.TxtLmg.Suppress = False
''      Rep.TxtLmp.Suppress = False
''      Rep.FldLmg.Suppress = False
''      Rep.FldLmp.Suppress = False
    Else
      Rep.TxtLmg.Suppress = True
      Rep.TxtLmp.Suppress = True
      Rep.FldLmg.Suppress = True
      Rep.FldLmp.Suppress = True
    End If
      
    If moCn.RecSeek("Select PmCd from Param where PTyp = 'SYS' and PmCd = 'ECH'") <> False Then
      If (adc("UoYN2")) = "Y" Or UCase(adc.MenuCd) = UCase("RepOrdCat") Or UCase(adc.MenuCd) = UCase("RepInvCTBCat") Then
        GRP_REP.SetFormula Rep, "wLabRtYN", "'N'"
      ElseIf (adc("UoYN2")) = "N" And UCase(adc.MenuCd) <> UCase("RepOrdCat") And UCase(adc.MenuCd) <> UCase("RepInvCTBCat") Then
        If (adc("UoYN4")) = "Y" Then
          GRP_REP.SetFormula Rep, "wLabRtYN", "'N'"
        Else
          GRP_REP.SetFormula Rep, "wLabRtYN", "'Y'"
        End If
      End If
    End If
    'BnM.1-if Range size selected in Rm level then print Range size otherwise existing behaviour
    If adc("UoRmGrp") = UCase("RmRSzRt") Then
      Rep.TxtFrLn.Suppress = True
      Rep.TxtToLn.Suppress = True
      Rep.OrLn1.Suppress = True
      Rep.OrLn2.Suppress = True
      Rep.wRmSz.Suppress = True
      Rep.hRmSz.Suppress = True
      Rep.TxtRngSz.Suppress = False
      Rep.RmRSzRt.Suppress = False
    Else
      Rep.TxtFrLn.Suppress = True
      Rep.TxtToLn.Suppress = True
      Rep.OrLn1.Suppress = True
      Rep.OrLn2.Suppress = True
      Rep.hRmSz.Suppress = False
      Rep.wRmSz.Suppress = False
      Rep.TxtRngSz.Suppress = True
      Rep.RmRSzRt.Suppress = True
    End If
    
    '*************************** ZUBIN **************************
    ' 09th Dec 2003, EMR206
     wOurDsgCd = "Max(IdDmCd) as CdOurDmCd"
     wOurRmCd = "Max(IrRmCd) as OrRmCd"
    If UCase(adc.MenuCd) = UCase("RepDetInv") Or UCase(adc.MenuCd) = UCase("RepAvgInv") Or _
              UCase(adc.MenuCd) = UCase("RepAvgInvCTB") Or UCase(adc.MenuCd) = UCase("RepDetInvCTB") Then
      '*** Jay 2.14 *** (InCmCd replaced with IdExpCmCd)
      '****** Sachin 3.02 (Partition Key added)
      If UCase(adc("UoYN7")) = "Y" Then
        wOurDsgCd = "IsNull((Select  cast(CdCustDmCd as Varchar(40)) from CustDsg " + _
                          "Where " + IIF(gs_Partition = ctCurrPrtn, " CdPrtKey='" + ctCurrPrtn + "' And  ", "") + " CdCtg = 'C' and CdCd = Max(IdExpCmCd) and " + _
                          "CdOurDmCd = Max(IdDmCd) and CdOurSfx = Max(IdSfx) and " + _
                          "CdOurDmSz = Max(IdDmSz)), " + _
                        "IsNull((Select cast(CdCustDmCd as Varchar(40)) from CustDsg " + _
                              "where " + IIF(gs_Partition = ctCurrPrtn, " CdPrtKey='" + ctCurrPrtn + "' and ", "") + " CdCtg = 'C' and CdCd = Max(IdExpCmCd) and " + _
                              "CdOurDmCd = Max(IdDmCd) and CdOurSfx = '' and " + _
                              "CdOurDmSz = Max(IdDmSz)), " + _
                            "IsNull((Select cast(CdCustDmCd as Varchar(40)) from CustDsg " + _
                                  "where " + IIF(gs_Partition = ctCurrPrtn, " CdPrtKey='" + ctCurrPrtn + "' and  ", "") + " CdCtg = 'C' and CdCd = Max(IdExpCmCd) and " + _
                                  "CdOurDmCd = Max(IdDmCd) and CdOurSfx = Max(IdSfx) and " + _
                                  "CdOurDmSz = '' ), " + _
                                "IsNull((Select cast(CdCustDmCd as Varchar(40)) from CustDsg " + _
                                      "where " + IIF(gs_Partition = ctCurrPrtn, " CdPrtKey='" + ctCurrPrtn + "' and  ", "") + " CdCtg = 'C' and CdCd = Max(IdExpCmCd) and " + _
                                      "CdOurDmCd = Max(IdDmCd) and CdOurSfx = '' and " + _
                                      "CdOurDmSz = ''), '{' + max(IdDmCd) + '}')))) as CdOurDmCd"
      Else
        wOurDsgCd = "Max(IdDmCd) as CdOurDmCd"
      End If
    ElseIf UCase(adc.MenuCd) = UCase("RepInvCat") Or UCase(adc.MenuCd) = UCase("RepInvCTBCat") Then
      If UCase(adc("UoYN7")) = "Y" Then
        wOurDsgCd = "IsNull((Select cast(CdCustDmCd as Varchar(40)) from CustDsg " + _
                          "Where " + IIF(gs_Partition = ctCurrPrtn, " CdPrtKey='" + ctCurrPrtn + "' and ", "") + " CdCtg = 'C' and CdCd = IdExpCmCd and " + _
                          "CdOurDmCd = IdDmCd and CdOurSfx = IdSfx and " + _
                          "CdOurDmSz = IdDmSz ), " + _
                        "IsNull((Select cast(CdCustDmCd as Varchar(40)) from CustDsg " + _
                              "where " + IIF(gs_Partition = ctCurrPrtn, " CdPrtKey='" + ctCurrPrtn + "' and ", "") + " CdCtg = 'C' and CdCd = IdExpCmCd and " + _
                              "CdOurDmCd = IdDmCd and CdOurSfx = '' and " + _
                              "CdOurDmSz = IdDmSz ), " + _
                            "IsNull((Select cast(CdCustDmCd as Varchar(40)) from CustDsg " + _
                                  "where " + IIF(gs_Partition = ctCurrPrtn, " CdPrtKey='" + ctCurrPrtn + "' and ", "") + " CdCtg = 'C' and CdCd = IdExpCmCd and " + _
                                  "CdOurDmCd = IdDmCd and CdOurSfx = IdSfx and " + _
                                  "CdOurDmSz = '' ), " + _
                                "IsNull((Select cast(CdCustDmCd as Varchar(40)) from CustDsg " + _
                                      "where " + IIF(gs_Partition = ctCurrPrtn, " CdPrtKey='" + ctCurrPrtn + "' and ", "") + " CdCtg = 'C' and CdCd = IdExpCmCd and " + _
                                      "CdOurDmCd = IdDmCd and CdOurSfx = '' and " + _
                                      "CdOurDmSz = ''), '{' + IdDmCd + '}')))) as CdOurDmCd"
      Else
        wOurDsgCd = "IdDmCd as CdOurDmCd"
      End If
    End If
    '*** Jay 2.14 *** (InCmCd replaced with IdExpCmCd)
    
    '*** Jay 2.14 *** (InCmCd replaced with IdExpCmCd)
    If (UCase(adc.MenuCd) = UCase("RepDetInv") Or UCase(adc.MenuCd) = UCase("RepDetInvCTB")) And UCase(adc("UoYN6")) = "Y" Then
    '          wOurRmCd = "IsNull((Select Cast(CrCustRmCd as Varchar(14)) from CustRm where " + IIF(gs_Partition = ctCurrPrtn, " CrPrtKey='" + ctCurrPrtn + "' and ", "") + " CrCtg = 'C' and " + _
    '                  "CrCd = Max(IdExpCmCd) and CrRmCtg = Max(RmCtg) and " + _
    '                  "CrOurRmCd = Max(IrRmCd) and 0 Between CrFrLn and CrToLn)," + _
                      "'{' + Max(IrRmCd) + '}') as OrRmCd"
      '578.19 varChar(14) changed to varChar(16)
      wOurRmCd = " (Case when Max(IrCustRmCd) ='' Then '{'+Max(IrRmCd)+'}' Else Max(Cast(IrCustRmCd as varChar(16))) End) as OrRmCd "        '4.1.0.0
    Else
      wOurRmCd = "Max(IrRmCd) as OrRmCd"
    End If
    '*************************** ZUBIN **************************
    '*****************************Geeta************Emr206*************
    If (UCase(adc.MenuCd) = UCase("RepInvCat") Or UCase(adc.MenuCd) = UCase("RepInvCTBCat")) And UCase(adc("UoYN6")) = "Y" Then
    '          wOurRmCd = "IsNull((Select cast(CrCustRmCd as Varchar(14))  from CustRm where " + IIF(gs_Partition = ctCurrPrtn, " CrPrtKey='" + ctCurrPrtn + "' and ", "") + " CrCtg = 'C' and " + _
    '                  "CrCd = IdExpCmCd and CrRmCtg = RmCtg and " + _
    '                  "CrOurRmCd = IrRmCd and 0 Between CrFrLn and CrToLn)," + _
    '                  "'{' + IrRmCd + '}') as OrRmCd"
        '578.19 varChar(14) changed to varChar(16)
        wOurRmCd = " (Case when IrCustRmCd ='' Then '{'+IrRmCd+'}' Else Cast(IrCustRmCd as varChar(16)) End) as OrRmCd "
    ElseIf (UCase(adc.MenuCd) = UCase("RepInvCat") Or UCase(adc.MenuCd) = UCase("RepInvCTBCat")) And UCase(adc("UoYN6")) <> "Y" Then
        wOurRmCd = "IrRmCd as OrRmCd"
    End If
    '*************************************Geeta***************************************
    '*** Jay 2.14 *** (InCmCd replaced with IdExpCmCd)
     
    '*** For the Group Formula of the report
    '************Setting Group2 & Group3 Formula***********************
    
    Dim wGrpArr1() As String, wGrpFormula1 As String, wGrpDescFormula1 As String
    Dim wGrpArr2() As String, wGrpFormula2 As String, wGrpDescFormula2 As String
    Dim i As Integer, qDmCtgDesc As String, qRmCtgDesc As String, qDmColDesc As String, qDmKtDesc As String
    Dim qVaCtgDesc As String, qMainMet As String, wOrJoin As String
  
    ' **** Zubin 211 **** '
    Dim wDetArr() As String, wDetSrt As String
    ' **** Zubin 211 **** '
    ' Zubin 212
    Dim qDmSalCtgDesc As String
    
    Dim wInGrp2 As String, wDetSrtExc As String ' **** Manali 3.6.0
    
    qDmCtgDesc = "'' as qDmCtgDesc": qRmCtgDesc = "'' as qRmCtgDesc"
    qDmColDesc = "'' as qDmColDesc": qVaCtgDesc = "'' as qVaCtgDesc": qDmKtDesc = "'' as qDmKtDesc"
    qMainMet = "space(2) as OrRmCtg": wOrJoin = ""
    ' Zubin 212
    
    qDmSalCtgDesc = "'' as qDmSalCtgDesc"
   If adc("UoYN29") = "Y" Then
            wGrpFormula1 = "Trim({rdo.OdOmCmCd})"
            wExcelOrdBy = IIF(UCase(adc.MenuCd) = UCase("RepInvCat"), "OdOmCmCd", "Max(OdOmCmCd)")
            ws_OrdBy = "OdOmCmCd"
     End If
    
    
   If (UCase(adc.MenuCd) = UCase("RepDetInv") Or UCase(adc.MenuCd) = UCase("RepAvgInv") Or UCase(adc.MenuCd) = UCase("RepInvCat")) Then
   
    wGrpArr1 = Split(moCn.GetFldVal("Select PDesc225 from Param where PTyp= 'INGR' and " + _
           "PMCd= '" + adc("UoInGrp") + "' "), ",")
    For i = 0 To UBound(wGrpArr1)
     Select Case UCase(wGrpArr1(i))
      Case Is = "RMCTG"
          If UCase(adc.MenuCd) = UCase("RepDetInv") Or UCase(adc.MenuCd) = UCase("RepAvgInv") Then
          qRmCtgDesc = "IsNull((Select PDesc from Param where PTyp= 'RMCTG' and PMCd= " + _
                       "(select RmCtg from RmMst join InvRm on RmCd=IrRmCd " + IIF(gs_Partition = ctCurrPrtn, " And RmPrtKey=IrPrtKey ", "") + "  and  IrMainMet='Y' " + _
                       " Where " + IIF(gs_Partition = ctCurrPrtn, " IrPrtKey='" + ctCurrPrtn + "' And ", "") + " IrCocd = Max(InCocd) And IrTc = InTc  and IrYy = InYy and IrChr=InChr " + _
                       " and IrNo=InNo and IrSr=IdSr)),'') as qRmCtgDesc"
        Else
          qRmCtgDesc = "IsNull((Select PDesc from Param where PTyp= 'RMCTG' and PMCd= " + _
                     "(select RmCtg from RmMst,InvRm where RmCd=IrRmCd " + IIF(gs_Partition = ctCurrPrtn, " And RmPrtKey=IrPrtKey ", "") + "  and  IrMainMet='Y' " + _
                     " and IrCocd = InCocd And IrTc = InTc  and IrYy = InYy and IrChr=InChr " + _
                     " and IrNo=InNo and IrSr=IdSr " + IIF(gs_Partition = ctCurrPrtn, " And IrPrtKey=IdPrtKey ", "") + " )),'') as qRmCtgDesc"
        End If
        
        wGrpDescFormula1 = wGrpDescFormula1 + IIF(wGrpDescFormula1 = "", "", "+', '+") + "Trim({rdo.qRmCtgDesc})"
        qMainMet = "max(OrRmCtg) as OrRmCtg"
        '****** Sachin 3.02 - Id fields in Joins
        wOrJoin = " join OrdRm on OrOdIdNo=OdIdNo and OrCoCd= OdCoCd and OrTc= OdTc and OrYy= OdYy and OrChr= OdChr and OrNo= OdNo and OrSr= OdSr and OrMainMet= 'Y' "
      Case Is = "KT"
        If UCase(adc.MenuCd) = UCase("RepDetInv") Or UCase(adc.MenuCd) = UCase("RepAvgInv") Then
          qDmKtDesc = "(Select PDesc from Param where PTyp= 'KT' and PMCd= max(IdKt)) as qDmKtDesc"
        Else
          qDmKtDesc = "(Select PDesc from Param where PTyp= 'KT' and PMCd= IdKt) as qDmKtDesc"
        End If
        wGrpDescFormula1 = wGrpDescFormula1 + IIF(wGrpDescFormula1 = "", "", "+', '+") + "Trim({rdo.qDmKtDesc})"
      Case Is = "DMCOL"
        If UCase(adc.MenuCd) = UCase("RepDetInv") Or UCase(adc.MenuCd) = UCase("RepAvgInv") Then
          qDmColDesc = "(Select PDesc from Param where PTyp= 'DMCOL' and PMCd= max(OdDmCol)) as qDmColDesc"
        Else
          qDmColDesc = "(Select PDesc from Param where PTyp= 'DMCOL' and PMCd= OdDmCol) as qDmColDesc"
        End If
        wGrpDescFormula1 = wGrpDescFormula1 + IIF(wGrpDescFormula1 = "", "", "+', '+") + "Trim({rdo.qDmColDesc})"
      Case Is = "VACTG"
        If UCase(adc.MenuCd) = UCase("RepDetInv") Or UCase(adc.MenuCd) = UCase("RepAvgInv") Then
          qVaCtgDesc = "(Select PDesc from Param where PTyp= 'VACTG' and PMCd= max(IdVaCtg)) as qVaCtgDesc"
        Else
          qVaCtgDesc = "(Select PDesc from Param where PTyp= 'VACTG' and PMCd= IdVaCtg) as qVaCtgDesc"
        End If
        wGrpDescFormula1 = wGrpDescFormula1 + IIF(wGrpDescFormula1 = "", "", "+', '+") + "Trim({rdo.qVaCtgDesc})"
      Case Is = "PCS/PRS"
        wGrpDescFormula1 = wGrpDescFormula1 + IIF(wGrpDescFormula1 = "", "", "+', '+") + "{rdo.DmUom}"
      Case Is = "DMCTG"
        If UCase(adc.MenuCd) = UCase("RepDetInv") Or UCase(adc.MenuCd) = UCase("RepAvgInv") Then
          qDmCtgDesc = " (Select PDesc from Param where PTyp= 'DMCTG' and PMCd= max(DmCtg)) as qDmCtgDesc"
        Else
          qDmCtgDesc = " (Select PDesc from Param where PTyp= 'DMCTG' and PMCd= DmCtg) as qDmCtgDesc"
        End If
        wGrpDescFormula1 = wGrpDescFormula1 + IIF(wGrpDescFormula1 = "", "", "+', '+") + "Trim({rdo.qDmCtgDesc})"
      Case Is = "GLDLS"
         wGrpDescFormula1 = wGrpDescFormula1 + IIF(wGrpDescFormula1 = "", "", "+', '+") + "'@'+ToText({rdo.IdGldLs}, '##0.000')+ '%'"
      ' Zubin 212
      Case Is = "SALCTG"
        If UCase(adc.MenuCd) = UCase("RepDetInv") Or UCase(adc.MenuCd) = UCase("RepAvgInv") Then
          qDmSalCtgDesc = " (Select PDesc from Param where PTyp= 'SALCTG' and PMCd= max(DmCtg) and PSCd= max(DmSalCtg)) as qDmSalCtgDesc "
        Else
          qDmSalCtgDesc = " (Select PDesc from Param where PTyp= 'SALCTG' and PMCd= DmCtg and PSCd= DmSalCtg) as qDmSalCtgDesc "
        End If
        wGrpDescFormula1 = wGrpDescFormula1 + IIF(wGrpDescFormula1 = "", "", "+', '+") + "Trim({rdo.qDmSalCtgDesc})"
      ' Zubin 212
      End Select
          wGrpFormula1 = wGrpFormula1 + IIF(wGrpFormula1 = "", "", "+") + "Trim(UpperCase(ToText({rdo." + moCn.GetFldVal("Select PDesc225 from Param where PTyp= 'INGRPRCP' and PMCd= '" + wGrpArr1(i) + "'") + "})))"
      ' **** Manali 3.6.0 - 04/11/09 - Order By clause for Excel Report
      If adc("UoYN23") = "Y" Then ws_OrdBy = ws_OrdBy + IIF(ws_OrdBy = "", "", ",") + moCn.GetFldVal("Select PDesc225 from Param where PTyp= 'INGRPRCP' and PMCd= '" + wGrpArr1(i) + "'")
      ' **** Manali 3.6.0 - 04/11/09 - Order By clause for Excel Report
        If UCase(adc.MenuCd) = UCase("RepInvCat") Or UCase(adc.MenuCd) = UCase("RepInvCTBCat") Then
           wExcelOrdBy = wExcelOrdBy + IIF(wExcelOrdBy = "", "", ",") + moCn.GetFldVal("Select (Case PDesc225 when 'IdGldLS' " + _
                       " Then '(Select PNum from Param where PTyp=''VACTG'' and PMCd=IdVaCtg and PSCd='''')' " + _
                       " When 'OrRmCtg' Then 'IsNull((select max(RmCtg) from RmMst ,InvRm where " + IIF(gs_Partition = ctCurrPrtn, " RmPrtKey=IrPrtKey and ", "") + " RmCd=IrRmCd  and  " + IIF(gs_Partition = ctCurrPrtn, " IrPrtKey=InPrtKey and ", "") + _
                       " IrMainMet=''Y'' and IrCocd = InCocd And IrTc = InTc and IrYy = InYy and IrChr=InChr and IrNo=InNo and IrSr=IdSr),'''')' Else PDesc225 End)   " + _
                       " from Param where PTyp= 'INGRPRCP' and PMCd= '" + wGrpArr1(i) + "'")
        Else
            wExcelOrdBy = wExcelOrdBy + IIF(wExcelOrdBy = "", "", ",") + moCn.GetFldVal("Select 'Max('+PDesc225+')' from Param where PTyp= 'INGRPRCP' and PMCd= '" + wGrpArr1(i) + "'")
        End If
    Next i
    
    wGrpArr2 = Split(moCn.GetFldVal("Select PDesc225 from Param where PTyp= 'INGR' and " + _
           "PMCd= '" + adc("UoInGrp1") + "' "), ",")
    For i = 0 To UBound(wGrpArr2)
     Select Case UCase(wGrpArr2(i))
      Case Is = "RMCTG"
          If UCase(adc.MenuCd) = UCase("RepDetInv") Or UCase(adc.MenuCd) = UCase("RepAvgInv") Then
          qRmCtgDesc = "IsNull((Select PDesc from Param where PTyp= 'RMCTG' and PMCd= " + _
                       "(select RmCtg from RmMst join InvRm on RmCd=IrRmCd " + IIF(gs_Partition = ctCurrPrtn, " And RmPrtKey=IrPrtKey ", "") + " and  IrMainMet='Y'" + _
                       " Where " + IIF(gs_Partition = ctCurrPrtn, " IrPrtKey='" + ctCurrPrtn + "' and ", "") + " IrCocd = Max(InCocd) And IrTc = InTc  and IrYy = InYy and IrChr=InChr " + _
                       " and IrNo=InNo and IrSr=IdSr)),'') as qRmCtgDesc"
        Else
          qRmCtgDesc = "IsNull((Select PDesc from Param where PTyp= 'RMCTG' and PMCd= " + _
                     "(select RmCtg from RmMst,InvRm where RmCd=IrRmCd " + IIF(gs_Partition = ctCurrPrtn, " And RmPrtKey=IrPrtKey ", "") + "  and  IrMainMet='Y' " + _
                     " and IrCocd = InCocd And IrTc = InTc  and IrYy = InYy and IrChr=InChr " + _
                     " and IrNo=InNo and IrSr=IdSr " + IIF(gs_Partition = ctCurrPrtn, " And IrPrtKey=IdPrtKey ", "") + " )),'') as qRmCtgDesc"
        End If
        wGrpDescFormula2 = wGrpDescFormula2 + IIF(wGrpDescFormula2 = "", "", "+', '+") + "Trim({rdo.qRmCtgDesc})"
        qMainMet = "max(OrRmCtg) as OrRmCtg"
        '****** Sachin 3.02 - Id fields in Joins
        wOrJoin = " join OrdRm on OrOdIdNo=OdIdNo and OrCoCd= OdCoCd and OrTc= OdTc and OrYy= OdYy and OrChr= OdChr and OrNo= OdNo and OrSr= OdSr and OrMainMet= 'Y' "
      Case Is = "KT"
        If UCase(adc.MenuCd) = UCase("RepDetInv") Or UCase(adc.MenuCd) = UCase("RepAvgInv") Then
            qDmKtDesc = "(Select PDesc from Param where PTyp= 'KT' and PMCd= max(IdKt)) as qDmKtDesc"
        Else
            qDmKtDesc = "(Select PDesc from Param where PTyp= 'KT' and PMCd= IdKt) as qDmKtDesc"
        End If
        wGrpDescFormula2 = wGrpDescFormula2 + IIF(wGrpDescFormula2 = "", "", "+', '+") + "Trim({rdo.qDmKtDesc})"
      Case Is = "DMCOL"
        If UCase(adc.MenuCd) = UCase("RepDetInv") Or UCase(adc.MenuCd) = UCase("RepAvgInv") Then
          qDmColDesc = "(Select PDesc from Param where PTyp= 'DMCOL' and PMCd= max(OdDmCol)) as qDmColDesc"
        Else
          qDmColDesc = "(Select PDesc from Param where PTyp= 'DMCOL' and PMCd= OdDmCol) as qDmColDesc"
        End If
        wGrpDescFormula2 = wGrpDescFormula2 + IIF(wGrpDescFormula2 = "", "", "+', '+") + "Trim({rdo.qDmColDesc})"
      Case Is = "VACTG"
        If UCase(adc.MenuCd) = UCase("RepDetInv") Or UCase(adc.MenuCd) = UCase("RepAvgInv") Then
          qVaCtgDesc = "(Select PDesc from Param where PTyp= 'VACTG' and PMCd= max(IdVaCtg)) as qVaCtgDesc"
        Else
          qVaCtgDesc = "(Select PDesc from Param where PTyp= 'VACTG' and PMCd=IdVaCtg) as qVaCtgDesc"
        End If
        wGrpDescFormula2 = wGrpDescFormula2 + IIF(wGrpDescFormula2 = "", "", "+', '+") + "Trim({rdo.qVaCtgDesc})"
      Case Is = "PCS/PRS"
        wGrpDescFormula2 = wGrpDescFormula2 + IIF(wGrpDescFormula2 = "", "", "+', '+") + "{rdo.DmUom}"
      Case Is = "DMCTG"
        If UCase(adc.MenuCd) = UCase("RepDetInv") Or UCase(adc.MenuCd) = UCase("RepAvgInv") Then
          qDmCtgDesc = " (Select PDesc from Param where PTyp= 'DMCTG' and PMCd= max(DmCtg)) as qDmCtgDesc"
        Else
          qDmCtgDesc = " (Select PDesc from Param where PTyp= 'DMCTG' and PMCd= DmCtg) as qDmCtgDesc"
        End If
        wGrpDescFormula2 = wGrpDescFormula2 + IIF(wGrpDescFormula2 = "", "", "+', '+") + "Trim({rdo.qDmCtgDesc})"
      Case Is = "GLDLS"
         wGrpDescFormula2 = wGrpDescFormula2 + IIF(wGrpDescFormula2 = "", "", "+', '+") + "'@'+ToText({rdo.IdGldLs}, '##0.000')+ '%'"
      ' Zubin 212
      Case Is = "SALCTG"
        If UCase(adc.MenuCd) = UCase("RepDetInv") Or UCase(adc.MenuCd) = UCase("RepAvgInv") Then
          qDmSalCtgDesc = " (Select PDesc from Param where PTyp= 'SALCTG' and PMCd= max(DmCtg) and PSCd= max(DmSalCtg)) as qDmSalCtgDesc "
        Else
          qDmSalCtgDesc = " (Select PDesc from Param where PTyp= 'SALCTG' and PMCd= DmCtg and PSCd= DmSalCtg) as qDmSalCtgDesc "
        End If
        wGrpDescFormula2 = wGrpDescFormula2 + IIF(wGrpDescFormula2 = "", "", "+', '+") + "Trim({rdo.qDmSalCtgDesc})"
      ' Zubin 212
      End Select
      wGrpFormula2 = wGrpFormula2 + IIF(wGrpFormula2 = "", "", "+") + "Trim(UpperCase(ToText({rdo." + moCn.GetFldVal("Select PDesc225 from Param where PTyp= 'INGRPRCP' and PMCd= '" + wGrpArr2(i) + "'") + "})))"
      ' **** Manali 3.6.0 - 04/11/09 - Order By clause for Excel Report
         wInGrp2 = moCn.GetFldVal("Select PDesc225 from Param where PTyp= 'INGRPRCP' and PMCd= '" + wGrpArr2(i) + "'")
        
      If adc("UoYN23") = "Y" Then
        ws_OrdBy = ws_OrdBy + IIF(ws_OrdBy <> "" And InStr(1, ws_OrdBy, wInGrp2) = 0, ",", "") + IIF(InStr(1, ws_OrdBy, wInGrp2) = 0, wInGrp2, "")
      End If
      
     If UCase(adc.MenuCd) = UCase("RepInvCat") Or UCase(adc.MenuCd) = UCase("RepInvCTBCat") Then
        wInGrp2 = moCn.GetFldVal("Select (Case PDesc225 when 'IdGldLS' " + _
                       " Then '(Select PNum from Param where PTyp=''VACTG'' and PMCd=IdVaCtg and PSCd='''')' " + _
                       " When 'OrRmCtg' Then 'IsNull((select max(RmCtg) from RmMst ,InvRm where " + IIF(gs_Partition = ctCurrPrtn, " RmPrtKey=IrPrtKey and ", "") + " RmCd=IrRmCd  and  " + IIF(gs_Partition = ctCurrPrtn, " IrPrtKey=InPrtKey and ", "") + _
                       " IrMainMet=''Y'' and IrCocd = InCocd And IrTc = InTc and IrYy = InYy and IrChr=InChr and IrNo=InNo and IrSr=IdSr),'''')' Else PDesc225 End)  from Param where PTyp= 'INGRPRCP' and PMCd= '" + wGrpArr2(i) + "'")
     Else
       wInGrp2 = moCn.GetFldVal("Select 'Max('+PDesc225+')' from Param where PTyp= 'INGRPRCP' and PMCd= '" + wGrpArr2(i) + "'")
     End If
      
      ' **** Manali 3.6.0 - 04/11/09 - Order By clause for Excel Report
      wExcelOrdBy = wExcelOrdBy + IIF(wExcelOrdBy <> "" And InStr(1, wExcelOrdBy, wInGrp2) = 0, ",", "") + IIF(InStr(1, wExcelOrdBy, wInGrp2) = 0, wInGrp2, "")
    Next i
   End If
    '****************************************************************************************
    
    '*** Jay 2.14 ***
    wDetArr = Split(moCn.GetFldVal("Select PDesc225 from Param where PTyp= 'INDET' and " + _
              "PMCd= '" + adc("UoInDet") + "' "), ",")
    wDetSrt = ""
  '  For I = 0 To UBound(wDetArr)
  '    wDetSrt = wDetSrt + IIF(wDetSrt = "", "", "+") + IIF(UCase(wDetArr(I)) = "IDSR", "ToText({rdo.IdSr}, '###0')", "Trim(UpperCase({rdo." + wDetArr(I) + "}))")
  '  Next I
  '  wDetSrt = Replace(wDetSrt, "rdo.Id", "rdo.Od")
  '  wGrpFormula2 = wGrpFormula2 + IIF(wGrpFormula2 = "", "", "+") + wDetSrt
  '  ' **** Zubin 211 **** '
    For i = 0 To UBound(wDetArr)
      '*** JAY 3.1.2 DT:21/01/08 *** [change copied from 2.14 prj]
      ' **** Manali 3.6.0 - 10/10/09 - IdTrayNo/Sr added
      'wDetSrt = IIF(UCase(wDetArr(i)) = "IDSR", "ToText({rdo.IdSr}, '###0')", "Trim(UpperCase({rdo." + wDetArr(i) + "}))")
      'wDetSrt = wDetSrt + IIF(wDetSrt = "", "", "+") + IIF(UCase(wDetArr(i)) = "IDSR", "ToText({rdo.IdSr}, '###0')", "Trim(UpperCase({rdo." + wDetArr(i) + "}))")
      wDetSrt = wDetSrt + IIF(wDetSrt = "", "", "+") + IIF(UCase(wDetArr(i)) = "IDSR", "ToText({rdo.IdSr}, '###0')", IIF(UCase(wDetArr(i)) = "IDTRAYNO", "ToText({rdo.IdTrayNo}, '###0')", "Trim(UpperCase({rdo." + wDetArr(i) + "}))"))
      '*** JAY 3.1.2 DT:21/01/08 *** [change copied from 2.14 prj]
      wDetSrt = Replace(wDetSrt, "rdo.Id", "rdo.Od")
      If UCase(adc("UoInDet")) <> "DMCTG" And UCase(adc("UoInDet")) <> "PCS/PRS" Then
          GRP_REP.SetFormula mRep, "OdSr", wDetSrt
      Else
          wGrpFormula2 = wGrpFormula2 + IIF(wGrpFormula2 = "", "", "+") + wDetSrt
      End If
      ' **** Manali 3.6.0 - 04/11/09 - Order By clause for Excel Report
        wDetSrtExc = wDetArr(i)
        If UCase(adc.MenuCd) <> UCase("RepInvCat") And UCase(adc.MenuCd) <> UCase("RepInvCTBCat") Then _
           wDetSrtExc = IIF(InStr(1, wDetSrtExc, "Id") = 1, Replace(wDetSrtExc, "Id", "Od"), wDetSrtExc)
             
      If adc("UoYN23") = "Y" And UCase(wDetArr(i)) <> UCase("IDSR") Then
        ws_OrdBy = ws_OrdBy + IIF(ws_OrdBy <> "" And InStr(1, ws_OrdBy, wDetSrtExc) = 0, ",", "") + IIF(InStr(1, ws_OrdBy, wDetSrtExc) = 0, wDetSrtExc, "")
      End If
      ' **** Manali 3.6.0 - 04/11/09 - Order By clause for Excel Report
        wDetSrtExc = IIF(UCase(adc.MenuCd) = UCase("RepInvCat") Or UCase(adc.MenuCd) = UCase("RepInvCTBCat"), wDetSrtExc, " Max(" + wDetSrtExc + ")")
       wExcelOrdBy = wExcelOrdBy + IIF(wExcelOrdBy <> "" And InStr(1, wExcelOrdBy, wDetSrtExc) = 0, ",", "") + IIF(InStr(1, wExcelOrdBy, wDetSrtExc) = 0, wDetSrtExc, "")
    Next i
    
    '*** (Bef 16/05/07)
    'If UCase(adc("UoInDet")) = "DMCTG" Or UCase(adc("UoInDet")) = "PCS/PRS" Then _
    '                                    GRP_REP.SetFormula mRep, "OdSr", "{rdo.OdSr}"
    '*** (Bef 16/05/07)
    
    '*** (Jen 2.14 Patch 1) (16/05/07)
    If UCase(adc("UoInDet")) = "DMCTG" Or UCase(adc("UoInDet")) = "PCS/PRS" Then
      GRP_REP.SetFormula mRep, "OdSr", "{rdo.OdSr}"
    Else
      GRP_REP.SetFormula mRep, "OdSr", wDetSrt + IIF(Trim(wDetSrt) <> "", "+ '/' +", "") + "ToText({rdo.OdSr})"
    End If
    '*** (Jen 2.14 Patch 1) (16/05/07)
    
    If UCase(adc.MenuCd) <> UCase("RepInvCat") And UCase(adc.MenuCd) <> UCase("RepInvCTBCat") Then
      If adc("UOYN13") = "Y" Then 'And (adc("UoInGrp") <> "" And adc("UoInGrp1") <> "")
        Rep.Grp5H.Suppress = False
      Else
        Rep.Grp5H.Suppress = True
      End If
    End If
    '*** Jay 2.14 ***
    
    
    If UCase(adc.MenuCd) = UCase("RepDetInv") Or UCase(adc.MenuCd) = UCase("RepAvgInv") Or _
       UCase(adc.MenuCd) = UCase("RepDetInvCTB") Or UCase(adc.MenuCd) = UCase("RepAvgInvCTB") Then
       
    ' **** Manali 3.03 - 07/07/08
    Dim wRmIdSzJoin As String, wTblStr As String
    wRmIdSzJoin = "Left Outer Join RmIdSz on RisRmCtg=RmCtg and RisRmId=RmId and RisSz1=IrRmSz "
    ' **** Manali 3.03 - 07/07/08
    'BnM.3 qBagNo added
    '62 qHuidNo added
    '578.19 Space(14) changed to Space(16) to OrRmcd
    '6.1B Subcust added
    wTblStr = "Select OmCoCd, OmYy, OmTc, OmChr, OmNo, OmCmCd, OmDt, OmPONo,OmPoDt,OmLmgCst,OmLmgSal,OmLmpSal,OmLmsSal,OmLmlSal,OmDelDt,OdDelDt," + _
              " OmHCtCd, OdSr, OdTc as IdExpOdTc, OdYy as IdExpOdYy, OdChr as IdExpOdChr, OdNo as IdExpOdNo, OdSr as IdExpOdSr, " + _
              " OdDmCd,Space(40) as CdOurDmCd,OdDmSz,OdSfx,OdDmStkCoCd,OdDmStkYy, OdDmStkChr, OdDmStkNo, DmCtg,DmSalCtg,DmTcTyp,OdDmCol, 0 as OdTrayNo,OdPicNm,OdSubCust,OdPoNo,OdKt as IdKt,DmUom, OdVaCtg as IdVaCtg," + _
              " PNum as IdGldLs,space(254) as qRmCtgDesc,space(254) as qDmCtgDesc,space(254) as qDmSalCtgDesc,space(254) as qDmColDesc,space(254) as qDmKtDesc,space(254) as qVaCtgDesc, " + _
              " OrRmctg, OdDmCol as OdDmCol3,OdGldAs,OdLabAs,OdGldAsWt,OdLabAsWt,OdSalPrc,OdOrdEnt as qOrdQty, OdSalPrc as qSalVal,OdDmPrdInst, " + _
              " OdCmPrdInst,OdCmStmpInst, OdSalRem, OdSzInst,OmSalRem,OmCmDelTerms, OmCmPayTerms, OmPoNo AS qOdOmPoNo, " + _
              " CmLkUpMetLs,DmLsCtg,OdMulby,CmName, OdOmCmCd,DmWaxWt,Space(20) as qBagNo,Space(30) as qHuidNo,Space(4) as qTag,OrMainMet,0 as qMainMet,Space(1) as qPtrYN,  OrSrNo, OrRmCtg as RmCtg,OrRmSCtg, '' as OrLotNo, Space(16) as OrRmCd ,RmQw,OrLn1,OrLn2,OrLn3,PDesc225 as qRmSzDesc,OrRmPtr  , " + _
              " OrPrdQty as OrQty,OrPrdWt as OrWt,OrLmeSal,OrSalRt,OrSetSCd,OrSetSalRt,OrSetSalVal,OrSalVal,Space(10) as RmRSzRt,OrAlySalRt,RmPurityRt,0 as qRmCtgSr, " + _
              " 0 as OlSrNo,OlMCd,OlSCd, PDesc as qLabDesc,RmQW as OlQw,OlQty,OlSalRt,OlSalVal, 0 as OctSrNo, OctCd ,OctSz,OctCtChr, 0 as OctParts,OctRmCd,OctDmCol,OctRmWt,OctDc, " + _
              " OctPrdSeq , OctMainPrtYn, OctDesc Into TmpA From OrdMst,OrdDsg,OrdRm,OrdLab,DsgMst,CustMst,RmMst,RmRt,OrdCt,Param where 1=2"
         
      moCn.CreateTmpTable "TmpA", wTblStr
        
 Dim ws_InstrStr As String
ws_InstrStr = " max(Replace(OdDmPrdInst,char(13)+ char(10),' ')) as OdDmPrdInst, " + _
                            " max(Replace(OdCmPrdInst,char(13)+ char(10),' ')) as OdCmPrdInst, " + _
                            " max(Replace(OdCmStmpInst,char(13)+ char(10),' ')) as OdCmStmpInst, " + _
                            " max(Replace(OdSalRem,char(13)+ char(10),' ')) as OdSalRem, " + _
                            " max(Replace(OdSzInst,char(13)+ char(10),' ')) as OdSzInst, " + _
                            " max(Replace(OmSalRem,char(13)+ char(10),' ')) as OmSalRem, " + _
                            " max(Replace(OmCmDelTerms,char(13)+ char(10),' ')) as OmCmDelTerms, " + _
                            " max(Replace(OmCmPayTerms,char(13)+ char(10),' ')) as OmCmPayTerms, "

      'BnM.3 qBagNo added
      '62 qHuidNo added
      '6.1-13 Case When OmLkUpMetLs <> '' Then OmLkUpMetLs Else CmLkUpMetLs End replced as OmLkUpMetLs
      ' to avoid seek from custmst
      '6.1b OdSubCust added
      wCommStr = "Select InCoCd as OmCoCd, InYy as OmYy, InTc as OmTc, InChr as OmChr, InNo as OmNo, max(InCmCd) as OmCmCd, " + _
                 "max(InPrnDt) as OmDt, max(InExpNo) as OmPoNo, max(InPrnDt) as OmPoDt, " + _
                 "max(" + IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then InLmgCst/InCnvFct else InLmgCst*InCnvFct end) ", "InLmgCst") + ") as OmLmgCst, " + _
                 "max(" + IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then InLmgSal/InCnvFct else InLmgSal*InCnvFct end) ", "InLmgSal") + ") as OmLmgSal, " + _
                 "max(" + IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then InLmpSal/InCnvFct else InLmpSal*InCnvFct end) ", "InLmpSal") + ") as OmLmpSal, " + _
                 "max(" + IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then InLmsSal/InCnvFct else InLmsSal*InCnvFct end) ", "InLmsSal") + ") as OmLmsSal, " + _
                 "max(" + IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then InLmlSal/InCnvFct else InLmlSal*InCnvFct end) ", "InLmlSal") + ") as OmLmlSal, " + _
                 "max(InPrnDt) as OmDelDt, max(InPrnDt) as OdDelDt, '' as OmHCtCd, IdSr as OdSr, " + _
                 "max(IdExpOdTc) as IdExpOdTc, max(IdExpOdYy) as IdExpOdYy, max(IdExpOdChr) as IdExpOdChr, max(IdExpOdNo) as IdExpOdNo, max(IdExpOdSr) as IdExpOdSr, " + _
                 "max(IdDmCd) as OdDmCd," + wOurDsgCd + ", max(IdDmSz) as OdDmSz, max(IdSfx) as OdSfx,'' as OdDmStkCoCd, '' as OdDmStkYy, '' as OdDmStkChr, 0 as OdDmStkNo, " + _
                 "max(DmCtg) as DmCtg, max(DmSalCtg) as DmSalCtg, 'DM' as DmTcTyp, Max(OdDmCol) as OdDmCol, max(IdTrayNo) as OdTrayNo, " + _
                 "max(OdPicNm) as OdPicNm, Max(OdSubCust) as OdSubCust,max(OdPoNo) as OdPoNo, max(IdKt) as IdKt, " + _
                 "max(DmUom) as DmUom, max(IdVaCtg) as IdVaCtg,max(IdGldLs) as IdGldLs," + qRmCtgDesc + "," + qDmCtgDesc + "," + qDmSalCtgDesc + "," + qDmColDesc + "," + qDmKtDesc + "," + qVaCtgDesc + "," + _
                 "IsNull((select RmCtg from RmMst join InvRm on RmCd=IrRmCd " + IIF(gs_Partition = ctCurrPrtn, " And RmPrtKey=IrPrtKey ", "") + " and  " + _
                 "IrMainMet='Y' Where IrCocd = Max(InCocd) And IrTc = InTc  " + _
                 "and IrYy = InYy and IrChr=InChr and IrNo=InNo and IrSr=IdSr " + IIF(gs_Partition = ctCurrPrtn, " And IrPrtKey=Max(IdPrtKey) ", "") + "),'') as OrRmctg /*qMainMetRmCtg*/," + _
                 "'' as OdDmCol3, " + _
                 "Max(IdGldAs) as OdGldAs, Max(IdLabAs) as OdLabAs, Max (" + wGldAsWt + ") as OdGldAsWt, Max (" + wLabAsWt + ") as OdLabAsWt, Max(" + wIdFOB + "/(case when IdQty> 0 then IdQty else 1 end)) as OdSalPrc, " + _
                 "max(IdQty) as qOrdQty, max(" + wIdFOB + ") as qSalVal," + ws_InstrStr + " max(OmPoNo) as qOdOmPoNo, " + _
                 "max(OmLkUpMetLs) as CmLkUpMetLs, max(DmLsCtg) as DmLsCtg, 1 as OdMulby, max(CmName) as CmName, Max(OmCmCd) as OdOmCmCd,Max(DmWaxWt) as DmWaxWt, " + _
                 "(select max(IfFdBYy+'/'+IfFdBChr+'/'+LTrim(str(IfFdBNo))) from InvFgd where IfCoCd= InCoCd and IfTc= InTc and IfYy= InYy and IfChr= InChr and IfNo= InNo And IfSr =IdSr ) as qBagNo, " + _
                 "(select isnull(max(CdCertNo),'') from CertDet join InvFgd on CdBCoCd =IfCoCd AND CdBYy =IfFdBYy and CdBChr =IfFdBChr and CdBNo =IfFdBNo " + _
                 "where CdCtg='H' and IfCoCd= InCoCd and IfTc= InTc and IfYy= InYy and IfChr= InChr and IfNo= InNo And IfSr =IdSr) as qHuidNo, "

              
                 
                 '' as OdDmPrdInst, '' as OdCmPrdInst, " + _
                 "'' as OdCmStmpInst, '' as OdSalRem, '' as OdSzInst, '' as OmSalRem, '' as OmCmDelTerms, '' as OmCmPayTerms, max(OmPoNo) as qOdOmPoNo, " + _
                 "max(Case When OmLkUpMetLs <> '' Then OmLkUpMetLs Else CmLkUpMetLs End) as CmLkUpMetLs, max(DmLsCtg) as DmLsCtg, 1 as OdMulby, max(CmName) as CmName, Max(OmCmCd) as OdOmCmCd,Max(DmWaxWt) as DmWaxWt, "
  
             If adc("UoLabGrp") = "N" Then wIrSetVal = "0"

  
 
      Select Case UCase(adc("UoRmGrp"))
      ' ***** Manali 3.6.0 - 04/11/09 - OrAlySalRt added for Excel Report
      Case Is = UCase("N")
      'PREVIOUS - EMR205 <<<<<<====================================
        '  wRmStr = " 0 as IrSrNo, '*' as RmCtg, '*' as RmSCtg, '*' as IrRmCd, 0 as OrLn1, " + _
                 "0 as OrLn2, 0 as IrRmPtr, sum(IrRmQty), sum(" + wIrRmWt + "), 0 as IrRmRt, '*' as IrSetSCd, " + _
                 "0 as IrSetRt, sum(" + wIrSetVal + "), sum(" + wIrRmVal + "), Space(10) as RmRSzRt,"
      '  wRmGrp = " group by InYy, InTc, InChr, InNo, IdSr "
      '**********************************Geeta*******************************
      '******Emr206 **********RmWt for ('D','C') is not Converted into grams]
        ' ***** Manali 3.03 - 29/06/08 - LmeRt added
        '*** Manali 3.03 - 07/07/08 - RmSzDesc added
        wRmStr = " IrSrNo as OrSrNo, max(RmCtg) as RmCtg, max(RmSCtg) as OrRmSCtg,'' as OrLotNo,  " + wOurRmCd + ", max(RmQw) As RmQw, max(IrRmSz) as OrLn1, max(IrRmSz2) as OrLn2, max(IrRmSz3) as OrLn3, " + wRmSzDesc + " as qRmSzDesc,  0 as OrRmPtr, " + _
                 "sum(IrRmQty) as OrQty, sum(" + wIrRmWt + ") as OrWt, " + _
                 "max (" + IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then " + wIrLmeRt + " /InCnvFct else " + wIrLmeRt + " *InCnvFct end)", wIrLmeRt) + ")  as OrLmeSal, " + _
                 "max(" + wIrRmRt + ") as OrSalRt, max(IrSetSCd) as OrSetSCd, " + _
                 "max(" + wIrSetRt + ") as OrSetSalRt, sum(" + wIrSetVal + ") as OrSetSalVal, sum(" + wIrRmVal + ") as OrSalVal, Space(10) as RmRSzRt," + _
                 "max(" + wAlySalRt + ") as OrAlySalRt, "
        wRmGrp = " group by InPrtKey, InIdNo, InCoCd, InYy, InTc, InChr, InNo, IdSr, IrSrNo "
      Case Is = UCase("Sr")
          '*************************** ZUBIN **************************
        ' 08th Dec 2003, EMR206  Repalced IrRmcd with  " + wOurRmCd + "
        ' ***** Manali 3.03 - 29/06/08 - LmeRt added
        '*** Manali 3.03 - 07/07/08 - RmSzDesc added
        wRmStr = " IrSrNo as OrSrNo, max(RmCtg) as RmCtg, max(RmSCtg) as OrRmSCtg,'' as OrLotNo,  " + wOurRmCd + ", max(RmQw) As RmQw, max(IrRmSz) as OrLn1,  max(IrRmSz2) as OrLn2, max(IrRmSz3) as OrLn3, " + wRmSzDesc + " as qRmSzDesc,  0 as OrRmPtr, " + _
                 "sum(IrRmQty) as OrQty, sum(" + wIrRmWt + ") as OrWt, " + _
                 "max (" + IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then " + wIrLmeRt + " /InCnvFct else " + wIrLmeRt + " *InCnvFct end)", wIrLmeRt) + ")  as OrLmeSal, " + _
                 "max(" + wIrRmRt + ") as OrSalRt, max(IrSetSCd) as OrSetSCd, " + _
                 "max(" + wIrSetRt + ") as OrSetSalRt, sum(" + wIrSetVal + ") as OrSetSalVal, sum(" + wIrRmVal + ") as OrSalVal, Space(10) as RmRSzRt, " + _
                 "max(" + wAlySalRt + ") as OrAlySalRt, "
        wRmGrp = " group by InPrtKey, InIdNo, InCoCd, InYy, InTc, InChr, InNo, IdSr, IrSrNo "
      Case Is = UCase("RmCdRt")
  ' **** Zubin 211 **** '
  '''''      wRmStr = " 0 as OrSrNo, RmCtg, RmSCtg as OrRmSCtg, " + wOurRmCd + ", max(IrRmSz) as OrLn1, 0 as OrLn2, 0 as OrRmPtr, " + _
  '''''               "sum(IrRmQty) as OrQty, sum(" + wIrRmWt + ") as OrWt, " + wIrRmRt + " as OrSalRt, '*' as OrSetSCd, " + _
  '''''               "0 as OrSetSalRt, sum(" + wIrSetVal + ") as OrSetSalVal, sum(" + wIrRmVal + ") as OrSalVal, Space(10) as RmRSzRt,"
        ' ***** Manali 3.03 - 29/06/08 - LmeRt added
        '*** Manali 3.03 - 07/07/08 - RmSzDesc added
        wRmStr = " 0 as OrSrNo, RmCtg, RmSCtg as OrRmSCtg,'' as OrLotNo, " + wOurRmCd + ", RmQw, 0 as OrLn1, 0 as OrLn2, 0 as OrLn3, '' as qRmSzDesc,  0 as OrRmPtr, " + _
                 "sum(IrRmQty) as OrQty, sum(" + wIrRmWt + ") as OrWt, " + _
                 "max(" + IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then " + wIrLmeRt + " /InCnvFct else " + wIrLmeRt + " *InCnvFct end)", wIrLmeRt) + ")  as OrLmeSal, " + _
                 wIrRmRt + " as OrSalRt, '*' as OrSetSCd, " + _
                 "0 as OrSetSalRt, sum(" + wIrSetVal + ") as OrSetSalVal, sum(" + wIrRmVal + ") as OrSalVal, Space(10) as RmRSzRt, " + _
                 "max(" + wAlySalRt + ") as OrAlySalRt, "
        wRmGrp = " group by InPrtKey, InIdNo, InCoCd, InYy, InTc, InChr, InNo, IdSr, RmCtg, RmSCtg, IrRmCd, RmQw,  " + wIrRmRt
  ' **** Zubin 211 **** '
      Case Is = UCase("RmCd")
        ' ***** Manali 3.03 - 29/06/08 - LmeRt added
        '*** Manali 3.03 - 07/07/08 - RmSzDesc added
        wRmStr = " 0 as OrSrNo, RmCtg, RmSCtg as OrRmSCtg,'' as OrLotNo, " + wOurRmCd + ", max(RmQw) as RmQw, 0 as OrLn1, 0 as OrLn2, 0 as OrLn3, '' as qRmSzDesc,  0 as OrRmPtr, " + _
                 "sum(IrRmQty) as OrQty, sum(" + wIrRmWt + ") as OrWt, 0.00 as OrLmeSal, 0 as OrSalRt, '*' as OrSetSCd, 0 as OrSetSalRt, " + _
                 "sum(" + wIrSetVal + ") as OrSetSalVal, sum(" + wIrRmVal + ") as OrSalVal, Space(10) as RmRSzRt, 0.00 as OrAlySalRt, "
        wRmGrp = " group by InPrtKey, InIdNo, InCoCd, InYy, InTc, InChr, InNo, IdSr, RmCtg, RmSCtg, IrRmCd "
      Case Is = UCase("RmSCtg")
        ' ***** Manali 3.03 - 29/06/08 - LmeRt added
        '*** Manali 3.03 - 07/07/08 - RmSzDesc added
        wRmStr = " 0 as OrSrNo, RmCtg, RmSCtg as OrRmSCtg,'' as OrLotNo, '*' as OrRmCd, '*' as RmQw, 0 as OrLn1, 0 as OrLn2, 0 as OrLn3, '' as qRmSzDesc,  0 as OrRmPtr, " + _
                 "sum(IrRmQty) as OrQty, sum(" + wIrRmWt + ") as OrWt, 0.00 as OrLmeSal, 0 as OrSalRt, '' as OrSetSCd, 0 as OrSetSalRt, " + _
                 "sum(" + wIrSetVal + ") as OrSetSalVal, sum(" + wIrRmVal + ") as OrSalVal, Space(10) as RmRSzRt, 0.00 as OrAlySalRt, "
        wRmGrp = " group by InPrtKey, InIdNo, InCoCd, InYy, InTc, InChr, InNo, IdSr, RmCtg, RmSCtg "
      Case Is = UCase("RmCtg")
        ' ***** Manali 3.03 - 29/06/08 - LmeRt added
        '*** Manali 3.03 - 07/07/08 - RmSzDesc added
        wRmStr = " 0 as OrSrNo, RmCtg, '*'  as OrRmSCtg, '' as OrLotNo, '*' as OrRmCd, '*' as RmQw, 0 as OrLn1, 0 as OrLn2, 0 as OrLn3, '' as qRmSzDesc,  0 as OrRmPtr, " + _
                 "sum(IrRmQty) as OrQty, sum(" + wIrRmWt + ") as OrWt, 0.00 as OrLmeSal, 0 as OrSalRt, '*' as OrSetSCd, 0 as OrSetSalRt, " + _
                 "Sum (" + wIrSetVal + ") as OrSetSalVal, Sum(" + wIrRmVal + ") as OrSalVal, Space(10) as RmRSzRt, 0.00 As OrAlySalRt, "
        wRmGrp = " group by InPrtKey, InIdNo, InCoCd, InYy, InTc, InChr, InNo, IdSr, RmCtg "
      Case Is = UCase("RmRSzRt")
        ' ***** Manali 3.03 - 29/06/08 - LmeRt added
        '*** Manali 3.03 - 07/07/08 - RmSzDesc added
        wRmStr = " 0 as OrSrNo, max(RmCtg) as RmCtg, max(RmSCtg) as OrRmSCtg,'' as OrLotNo, " + wOurRmCd + ", max(RmQw) As RmQw, max(IrRmSz) as OrLn1, max(IrRmSz2) as OrLn2, max(IrRmSz3) as OrLn3, " + wRmSzDesc + " as qRmSzDesc,  0 as OrRmPtr, " + _
                 "sum(IrRmQty) as OrQty, sum(" + wIrRmWt + ") as OrWt, " + _
                 "max (" + IIF(wb_InBaseCur, " (Case when InMulDiv = 'M' then " + wIrLmeRt + " /InCnvFct else " + wIrLmeRt + " *InCnvFct end)", wIrLmeRt) + ")  as OrLmeSal, " + _
                 "max(" + wIrRmRt + ") as OrSalRt, max(IrSetSCd) as OrSetSCd, max(" + wIrSetRt + ") as OrSetSalRt, " + _
                 "sum (" + wIrSetVal + ") as OrSetSalVal, Sum(" + wIrRmVal + ") as OrSalVal, isnull(max(RrRngSz)," + GetRmSzDispFldNm("max(IrRmSz)", "max(IrRmSz2)", "max(IrRmSz3)") + ") as RmRSzRt, " + _
                 "max(" + wAlySalRt + ") as OrAlySalRt, "
        wRmGrp = " group by InPrtKey, InIdNo, InCoCd, InYy, InTc, InChr, InNo, IdSr, RmCtg, RmSCtg, RrRngSz "
      ' ***** Manali 3.6.0 - 25/09/09 - New RmGrp added - RmSCtgRt
      Case Is = UCase("RmSCtgRt")
        wRmStr = " 0 as OrSrNo, max(RmCtg) as RmCtg, max(RmSCtg) as OrRmSCtg,'' as OrLotNo, '*' as OrRmCd, max(RmQw) As RmQw, 0 as OrLn1, 0 as OrLn2, 0 as OrLn3, '' as qRmSzDesc,  0 as OrRmPtr, " + _
                 "sum(IrRmQty) as OrQty, sum(" + wIrRmWt + ") as OrWt, " + _
                 "max (" + IIF(wb_InBaseCur, " (Case when InMulDiv = 'M' then " + wIrLmeRt + " /InCnvFct else " + wIrLmeRt + " *InCnvFct end)", wIrLmeRt) + ")  as OrLmeSal, " + _
                 "max(" + wIrRmRt + ") as OrSalRt, '' as OrSetSCd, 0 as OrSetSalRt, " + _
                 "sum (" + wIrSetVal + ") as OrSetSalVal, Sum(" + wIrRmVal + ") as OrSalVal, Space(10) as RmRSzRt, " + _
                 "max(" + wAlySalRt + ") as OrAlySalRt, "
        wRmGrp = " group by InPrtKey, InIdNo, InCoCd, InYy, InTc, InChr, InNo, IdSr, RmCtg, RmSCtg, RmQW, " + wIrRmRt
      ' ***** Manali 3.6.0 - 25/09/09 - New RmGrp added - RmSCtgRt
      End Select
      '***********************Geeta*****************************
      '**************Emr206***Join On DsgMst,OrdDsg
      ' Zubin 213 (ws_CurJoin added)
      '*** Jay 2.13(CT) ***(Ordct flds added)
      '*** Jay 2.14Next ***[OctSrNo added]
      ' ***** Manali 3.03 -29/08/06 - joins with OrdRm and OrdMst added to get Lme Rt from Order
      wRmCommStr = "'1RM' as qTag, Max(IrMainMet) as OrMainMet, (case when max(IrMainMet)= 'Y' or max(IrRmCd)= " + _
               "IsNull((Select IrRmCd From InvRm where IrPrtKey=InPrtKey and IrCoCd= max(InCoCd) and IrTc= InTc and IrYy= InYy and " + _
               "IrChr= InChr and IrNo= InNo and IrSr= IdSr and IrMainMet= 'Y'), '') then 1 " + _
               "else 2 end) as qMainMet, (Select PValue3 from Param where PTyp= 'RMSCTG' " + _
               "and PMCd= Max(RmCtg) and PSCd= max(RmSCtg)) as qPtrYN, " + wRmStr + _
               " max(" + wRmPurity + ") as RmPurityRt, " + _
               "Max(Case When RmCtg in ('G', 'P', 'S', 'L','M') Then 1 When RmCtg In ('D', 'C') Then 2 When RmCtg = 'X' Then 3 Else 4 End) as qRmCtgSr, " + _
               "0 as OlSrNo,'' as OlMCd, '' as OlSCd, " + _
               "IsNull((select RTrim(PDesc) from Param where PTyp= 'LABSCD' and PMCd= 'SET' and " + _
               "PSCd= max(IrSetSCd)),'') as qLabDesc, " + _
               "'' as OlQw, 0 as OlQty, 0 as OlSalRt, 0 as OlSalVal, " + _
               "0 as OctSrNo, '' as OctCd , '' as OctSz, '' as OctCtChr, 0 as OctParts,  '' as OctRmCd, '' as OctDmCol, 0 as OctRmWt, " + _
               "'' as OctDc, '' as OctPrdSeq, '' as OctMainPrtYn, '' as OctDesc "

     wRmCommStr = wRmCommStr + "From InvHd " + _
               " join InvDsg on IdInIdNo=InIdNo and IdCoCd= InCoCd and IdTc= InTc and IdYy= InYy and IdChr= InChr and IdNo= InNo " + IIF(gs_Partition = ctCurrPrtn, " And IdPrtKey=InPrtKey ", "") + _
               " join InvRm on IrIdIdNo=IdIdNo and IrCoCd= IdCoCd and IrTc= IdTc and IrYy= IdYy " + _
               " and IrChr= IdChr and IrNo= IdNo and IrSr= IdSr " + IIF(gs_Partition = ctCurrPrtn, " And IrPrtKey=IdPrtKey ", "") + _
               " join RmMst on RmCd= IrRmCd " + IIF(gs_Partition = ctCurrPrtn, " And RmPrtKey=IrPrtKey ", "") + _
               " join DsgMst on DmIdNo=IdDmIdNo and DmTcTyp= 'DM' and DmCd= IdDmCd and DmSz= '' " + IIF(gs_Partition = ctCurrPrtn, " And DmPrtKey=IdPrtKey ", "") + _
               " join OrdDsg on OdCoCd=IdCoCd  and OdTc= IdExpOdTc and OdYy= IdExpOdYy " + _
               " and OdChr=IdExpOdChr and OdNo= IdExpOdNo and OdSr=IdExpOdSr  " + IIF(gs_Partition = ctCurrPrtn, " And OdPrtKey=IdPrtKey ", "") + _
               " join (Select OrPrtKey, OrOdIdNo, OrCoCd, OrTc, OrYy, OrChr, OrNo, OrSr, OrRmCd, max(OrLmeSal) as OrLmeSal, max(OrAlyCstRt) as OrAlyCstRt, max(OrAlySalRt) as  OrAlySalRt from OrdRm " + _
               " Group By OrPrtKey, OrOdIdNo, OrCoCd, OrTc, OrYy, OrChr, OrNo, OrSr, OrRmCd) OrdRm " + _
               " on OrCoCd=OdCoCd  and OrTc=  OdTc and OrYy=  OdYy  and OrChr= OdChr and OrNo= OdNo and OrSr=  OdSr   And OrPrtKey=OdPrtKey and OrRmCd=IrRmCd" + _
               " join OrdMst on OmCoCd=OdCoCd  and OmTc=  OdTc and OmYy=  OdYy  and OmChr= OdChr and OmNo= OdNo And  OmIdNo=OdOmIdNo " + IIF(gs_Partition = ctCurrPrtn, " And OmPrtKey=OdPrtKey ", "") + _
               " Left Outer join RmRt on RrTCTyp= 'RSZ' and RrCmCtg = 'C' and RrCmCd = 'ZSELF' and RrCmCurCd= '' and " + _
               " RrCtg = RmCtg and RrSCtg = RmSCtg and RrSTWGrp = '' and RrCd = '' and " + _
               " RrDmCtg = '' and RrLsCtg = '' and RrLabMCd = '' and RrFrLn <= IrRmSz and RrToLn >= IrRmSz " + IIF(gs_Partition = ctCurrPrtn, " And RrPrtKey=RmPrtKey ", "") + wRmIdSzJoin + _
               ws_CurJoin + _
               wCnd + IIF(gs_Partition = ctCurrPrtn, " And InPrtKey='" + ctCurrPrtn + "' ", "") + wRmGrp
               '(index= InvDsg) '(index= InvRm)
      '*** Jay 2.14Next ***[OctSrNo added]
      '*** Jay 2.13(CT) ***(Ordct flds added)
      
      Select Case UCase(adc("UoLabGrp"))
      Case Is = UCase("N")
        wLabStr = "0 as OlSrNo, '*' as OlMCd, '*' as OlSCd, " + _
                  "'*' as qLabDesc, '*' as OlQw, sum(iLQty) as OlQty, 0 as OlSalRt, sum(" + wiLVal + ") as OlSalVal "
        wLabGrp = " group by InCoCd, InYy, InTc, InChr, InNo, IdSr, iLSrNo "
      Case Is = UCase("Sr")
        '****** (Jen 2.14 Next) changed iLQw to wiLQw ******
        wLabStr = "iLSrNo as OlSrNo, max(iLMCd) as OlMcd, max(iLSCd) as OlSCd, " + _
                  "(select rtrim(PDesc) from Param where PTyp= 'LABSCD' and PMCd= max(iLMCd) and " + _
                  "PSCd= max(iLSCd)) as qLabDesc, max(" + wiLQw + ") as OlQw, sum(iLQty) as OlQty, max(" + wiLRt + ") as OlSalRt, " + _
                  "sum(" + wiLVal + ") as OlSalVal "
        wLabGrp = " group by InCoCd, InYy, InTc, InChr, InNo, IdSr, iLSrNo "
      Case Is = UCase("LabMCd")
        wLabStr = "0 as OlSrNo, iLMCd as OlMcd, '*' as OlSCd, " + _
                  "'*' as qLabDesc, '*' as OlQw, sum(iLQty) as OlQty, 0 as OlSalRt, sum(" + wiLVal + ") as OlSalVal "
        wLabGrp = " group by InCoCd, InYy, InTc, InChr, InNo, IdSr, iLMCd "
      Case Is = UCase("LabSCd")
        '****** (Jen 2.14 Next) changed iLQw to wiLQw ******
        wLabStr = "0 as OlSrNo, max(iLMCd) as OlMCd, iLSCd as OlSCd, " + _
                  "(select rtrim(PDesc) from Param where PTyp= 'LABSCD' and PMCd= max(iLMCd) and " + _
                  "PSCd= iLSCd) as qLabDesc, max(" + wiLQw + ") as OlQw, sum(iLQty) as OlQty, 0 as OlSalRt, sum(" + wiLVal + ") as OlSalVal "
        wLabGrp = " group by InCoCd, InYy, InTc, InChr, InNo, IdSr, iLSCd "
      End Select
      
      '***********************Geeta*****************************
      '**************Emr206***Join On DsgMst,OrdDsg
      ' Zubin 213 (ws_CurJoin added)
      '*** Jay 2.13(CT) ***(Ordct flds added)
      '*** Jay 2.14Next ***[OctSrNo added]
      ' ***** Manali 3.03 - 29/06/08 - LmeRt included
      '*** Manali 3.03 - 07/07/08 - RmSzDesc added
      ' ***** Manali 3.6.0- 04/11/09 - OrAlySalRt, RmPurityRt added for Excel Report
      wLabCommStr = "'2LAB' as qTag, '' as OrMainMet, 0 as qMainMet,'' as qPtrYN, 0 as OrSrNo, '' as RmCtg, '' as OrRmSCtg,'' as OrLotNo, " + _
                    "'' as OrRmCd, '' as RmQw, 0 as OrLn1, 0 as OrLn2, 0 as OrLn3, '' as qRmSzDesc,  0 as OrRmPtr, 0 as OrQty, 0 as OrWt, 0 as OrLmeSal, 0 as OrSalRt, " + _
                    "'' as OrSetSCd, 0 as OrSetSalRt, 0 as OrSetSalVal, 0 as OrSalVal, Space(10) as RmRSzRt," + _
                    "0 as OrAlySalRt, 0 as RmPurityRt, 10 as qRmCtgSr, " + _
                    wLabStr + ", 0 as OctSrNo, '' as OctCd , '' as OctSz, '' as OctCtChr, 0 as OctParts,  '' as OctRmCd, '' as OctDmCol, 0 as OctRmWt, " + _
                    "'' as OctDc, '' as OctPrdSeq, '' as OctMainPrtYn, '' as OctDesc " + _
                    "From InvHd " + _
                    " join InvDsg on IdInIdNo=InIdNo and IdCoCd= InCoCd and IdTc= InTc and IdYy= InYy and IdChr= InChr and IdNo= InNo " + IIF(gs_Partition = ctCurrPrtn, " And IdPrtKey=InPrtKey ", "") + _
                    " join InvLab on IlIdIdNo=IdIdNo and iLCoCd= IdCoCd and iLTc= IdTc and iLYy= IdYy and " + _
                    " iLChr= IdChr and iLNo= IdNo and iLSr= IdSr " + IIF(gs_Partition = ctCurrPrtn, " And IlPrtKey=IdPrtKey ", "") + _
                    " join OrdDsg on OdCoCd=IdCoCd  and OdTc= IdExpOdTc and OdYy= IdExpOdYy " + _
                    " and OdChr=IdExpOdChr and OdNo= IdExpOdNo and OdSr=IdExpOdSr  " + IIF(gs_Partition = ctCurrPrtn, " And OdPrtKey=IdPrtKey ", "") + _
                    " join OrdMst on OmCoCd=OdCoCd  and OmTc=  OdTc and OmYy=  OdYy  and OmChr= OdChr and OmNo= OdNo And  OmIdNo=OdOmIdNo " + IIF(gs_Partition = ctCurrPrtn, " And OmPrtKey=OdPrtKey ", "") + _
                    " join DsgMst on DmTcTyp= 'DM' and DmCd= IdDmCd and DmSz= '' " + IIF(gs_Partition = ctCurrPrtn, " And DmPrtKey=IdPrtKey ", "") + _
                    ws_CurJoin + _
                    wCnd + IIF(gs_Partition = ctCurrPrtn, " And InPrtKey='" + ctCurrPrtn + "' ", "") + wLabGrp
                    '(index= InvDsg) '(index= InvLab)
      '*** Jay 2.14Next ***[OctSrNo added]
      '*** Jay 2.13(CT) ***(Ordct flds added)
      
'      ws_OrdBy = ws_OrdBy + IIF(ws_OrdBy <> "", ",", "")
'      wSqlStrg = wCommStr + wRmCommStr + " Union All " + _
'                 wCommStr + wLabCommStr + _
'                 " Order by InCoCd, InYy, InTc, InChr, InNo, " + ws_OrdBy + " IdSr, qTag, qMainMet, OrSrNo, OlSrNo "
         ' **** Manali 3.8.0
        ws_OrdBy = ws_OrdBy + IIF(ws_OrdBy <> "", ",", "")
        wSqlStrg = wCommStr + wRmCommStr + " Union All " + _
                   wCommStr + wLabCommStr + _
                   " Order by InCoCd, InYy, InTc, InChr, InNo, " + ws_OrdBy + " IdSr, qTag, " + IIF(adc("UoYn23") = "Y", "qRmCtgSr", "qMainMet") + ", OrSrNo, OlSrNo "
        
                 
      moCn.Execute ("Insert Into TmpA " + wSqlStrg)
      wSqlStrg = "Select * From TmpA"
        
                 
      '*** Jay 2.13(CT) ***
      'Debug.Print wRmCommStr
      ' ***** Manali 3.6.0 - Excel Format Report
      If adc("UoYn23") = "Y" Then
        Dim wIdSrCnt As Integer
        Dim wRsInv As MDORowSet
        wIdSrCnt = moCn.GetFldVal("Select Count(IdSr) From InvHd " + _
                    " join InvDsg on IdInIdNo=InIdNo and IdCoCd= InCoCd and IdTc= InTc and IdYy= InYy and IdChr= InChr and IdNo= InNo " + IIF(gs_Partition = ctCurrPrtn, " And IdPrtKey=InPrtKey ", "") + _
                    " Join CustMst On CmCtg='C' and CmCd=InCmCd " + _
                    wCnd + " ")
        If wIdSrCnt = 0 Then
          DispMsg "No Data Found for given scope", etError: Exit Sub
        ElseIf wIdSrCnt > 285 Then
          DispMsg "No of Designs is greater than 285, Please Reduce the scope", etError: Exit Sub
        Else
          Set wRsInv = moCn.OpenRes(wSqlStrg + " order by OmCoCd, OmYy, OmTc, OmChr, OmNo, " + ws_OrdBy + " OdSr, qTag, /*qMainMet,*/ qRmCtgSr, OrSrNo, OlSrNo " + ctMaxDopOpt)
          wErrMsg = ""
          Call SetExcel(wRsInv)
          adc.RepRecSource = "Select * from Param Where 1=2"
          If wErrMsg <> "" Then
            DispMsg wErrMsg, etError
          End If
          Set wRsInv = Nothing
        End If
      ' ***** Manali 3.6.0 - Excel Format Report
      Else
        adc.RepRecSource = wSqlStrg + ctMaxDopOpt  '****** Sachin 3.01 added ctmaxdopopt
      End If
     
      
      wExcelOrdBy = wExcelOrdBy + IIF(wExcelOrdBy <> "", ",", "")
      Set wo_rsPic = moCn.OpenRes("Select Max(DmTcTyp) as DmTcTyp,Max(OdDmCd) as DmCd,Max(OdDmSz) as DmSz,OmCoCd,OmTc,OmYy,OmChr,OmNo,OdSr,Max(OdPicNm) as PicNm from TmpA" + _
                                  " Group By OmCoCd,OmTc,OmYy,OmChr,OmNo,OdSr " + _
                                  " order by OmCoCd, OmYy, OmTc, OmChr, OmNo, " + wExcelOrdBy + "OdSr ")
      If wo_rsPic.RecCount > 0 Then CmdChgImg.Visible = True
               
                  
      
      
      ' Zubin 213 (Base Curr)
      '****** Sachin 3.02 - Id fields In Joins
      If (adc("UoYN2")) = "Y" And UCase(adc("UoInVal")) = "S" Then _
          wIdFOB = "((Select sum(IrRmAWt * IrStkRt)+ " + _
                   "sum(" + IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then IrSetAVal/InCnvFct else IrSetAVal*InCnvFct end) ", "IrSetAVal") + ") " + _
                   "from InvRm " + IIF(wb_InBaseCur, " Join InvHd on InCoCd= IrCoCd and InTc= IrTc and InYy= IrYy and InChr= IrChr and InNo= IrNo " + IIF(gs_Partition = ctCurrPrtn, " And InPrtKey=IrPrtKey ", ""), "") + _
                   " where IrIdIdNo=IdIdNo and IrCoCd= IdCoCd and IrTc= IdTc and IrYy= IdYy and IrChr= IdChr and IrNo= IdNo and IrSr= IdSr " + IIF(gs_Partition = ctCurrPrtn, " And IrPrtKey=IdPrtKey ", "") + " ) " + _
                   " + " + _
                   "(Select sum(" + IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then iLAVAl/InCnvFct else iLAVAl*InCnvFct end) ", "iLAVAl") + ") " + _
                   "from InvLab " + IIF(wb_InBaseCur, " Join InvHd on InPrtKey=IlPrtKey and InCoCd= IlCoCd and InTc= IlTc and InYy= IlYy and InChr= IlChr and InNo= IlNo ", "") + _
                   " where IlIdIdNo=IdIdNo and IlCoCd= IdCoCd and IlTc= IdTc and IlYy= IdYy and IlChr= IdChr and IlNo= IdNo and IlSr= IdSr " + IIF(gs_Partition = ctCurrPrtn, " And IlPrtKey=IdPrtKey ", "") + ")) "
      ' Zubin 213 (ws_CurJoin added)
      '****** Sachin 3.02 - Id fields in Joins
      wDsgCtgStr = "Select IdYy as OdYy, IdTc as OdTc, IdChr as OdChr, IdNo as OdNo, IdSr as OdSr, DmCtg, (Select PDesc from Param " + _
                   "where PTyp= 'DMCTG' and PMCd= DmCtg) as qDesc, IdQty as qOrdQty " + IIF(wIdFOB <> "", ",", "") + _
                   wIdFOB + " as qSalVal From InvDsg join InvHd on InIdNo=IdInIdNo and InCoCd= IdCoCd and InTc= IdTc and " + _
                   "InYy= IdYy and InChr= IdChr and InNo= IdNo " + IIF(gs_Partition = ctCurrPrtn, " And InPrtKey=IdPrtKey ", "") + _
                   "join DsgMst on DmTcTyp= 'DM' and DmCd= IdDmCd and DmSz= '' " + IIF(gs_Partition = ctCurrPrtn, " And DmPrtKey=IdPrtKey ", "") + ws_CurJoin + wCnd + IIF(gs_Partition = ctCurrPrtn, " And IdPrtKey='" + ctCurrPrtn + "' ", "")
                 '(index= DsgMst)
  
      Call moCn.RepRes(mRep.Subreport1.OpenSubreport, wDsgCtgStr + ctMaxDopOpt)  '****** Sachin 3.01 added ctmaxdopopt
    '********************** RM Weight Box 31/07/03 **********************'
     'Dim wOdQty As String
     'wOdQty = IIf(adc("UoYN1") = "Y", "max(OdOrdQty)", "max(OdOrdEnt)") ???
     'sum(" + wIrRmVal + "),
      
      
  ' **** Zubin 211 **** '
  '''''    wRmDetSql = "Select IdTc as OdTc, IdYy as OdYy, IdChr as OdChr, IdNo as OdNo, RmCtg as OrRmCtg, sum(IrRmQty) as qRmQty, " + _
  '''''                "sum( " + wIrRmWt + ") as qRmWt, " + _
  '''''                "sum( " + wIrRmVal + " ) as qRmVal from InvHd " + _
  '''''                "Join InvDsg on IdCoCd= InCoCd and IdTc= InTc and IdYy= InYy and IdChr= InChr and IdNo= InNo " + _
  '''''                "Join InvRm on IrCoCd= IdCoCd and IrTc= IdTc and IrYy= IdYy and IrChr= IdChr and " + _
  '''''                "IrNo= IdNo And IrSr= IdSr " + _
  '''''                "Join RmMst on IrRmCd = RmCd " + wCnd + _
  '''''                " group by IdCoCd, IdTc, IdYy, IdChr, IdNo, RmCtg"
      ' Zubin 213 (ws_CurJoin added)
      '****** Sachin 3.02 - Id fields in Joins
      wRmDetSql = "Select IdTc as OdTc, IdYy as OdYy, IdChr as OdChr, IdNo as OdNo, RmCtg as OrRmCtg, " + _
                      "sum(IrRmQty) as qRmQty, 0 as qWSetQty, 0 as qHSetQty, " + _
                      "sum( " + wIrRmWt + ") as qRmWt, sum( " + wIrRmVal + " ) as qRmVal from InvHd " + _
                      " Join InvDsg on  IdInIdNo=InIdNo and IdCoCd= InCoCd and IdTc= InTc and IdYy= InYy and IdChr= InChr and IdNo= InNo " + IIF(gs_Partition = ctCurrPrtn, " And IdPrtKey=InPrtKey ", "") + _
                      " Join InvRm on IrIdIdNo=IdIdNo and IrCoCd= IdCoCd and IrTc= IdTc and IrYy= IdYy and IrChr= IdChr and " + _
                      " IrNo= IdNo And IrSr= IdSr " + IIF(gs_Partition = ctCurrPrtn, " And IrPrtKey=IdPrtKey ", "") + _
                      " Join RmMst on IrRmCd = RmCd " + IIF(gs_Partition = ctCurrPrtn, " And RmPrtKey=IrPrtKey ", "") + ws_CurJoin + wCnd + IIF(gs_Partition = ctCurrPrtn, " And InPrtKey='" + ctCurrPrtn + "' ", "") + _
                      " group by IdPrtKey, IdCoCd, IdTc, IdYy, IdChr, IdNo, RmCtg"
      
      Rep.Subreport2_hWSetQty.Suppress = True
      Rep.Subreport2_hHSetQty.Suppress = True
      
'      mRep.Subreport2_hWSetQty.Suppress = True
'      mRep.Subreport2_hHSetQty.Suppress = True
  ' **** Zubin 211 **** '
  
  
      Call moCn.RepRes(mRep.Subreport2.OpenSubreport, wRmDetSql)
      '**********************************************************************'
      'sv.24 query for stncls summary and calling subreport
      Dim wStnClsSummStr As String
      wStnClsSummStr = "Select IdTc as OdTc, IdYy as OdYy, IdChr as OdChr, IdNo as OdNo, PDesc as qRmStnCls, " + _
                      "sum(IrRmQty) as qRmQty, sum( " + wIrRmWt + ") as qRmWt, " + _
                      "sum( " + wIrRmVal + " ) as qRmVal from InvHd " + _
                      " Join InvDsg on  IdInIdNo=InIdNo and IdCoCd= InCoCd and IdTc= InTc and IdYy= InYy and IdChr= InChr and IdNo= InNo " + IIF(gs_Partition = ctCurrPrtn, " And IdPrtKey=InPrtKey ", "") + _
                      " Join InvRm on IrIdIdNo=IdIdNo and IrCoCd= IdCoCd and IrTc= IdTc and IrYy= IdYy and IrChr= IdChr and " + _
                      " IrNo= IdNo And IrSr= IdSr " + IIF(gs_Partition = ctCurrPrtn, " And IrPrtKey=IdPrtKey ", "") + _
                      " Join RmMst on IrRmCd = RmCd " + IIF(gs_Partition = ctCurrPrtn, " And RmPrtKey=IrPrtKey ", "") + ws_CurJoin + _
                      " Left Outer Join Param on PTyp= 'STNCLS' and PMCd=RmCtg and PSCd=RmStnCls " + _
                       wCnd + IIF(gs_Partition = ctCurrPrtn, " And InPrtKey='" + ctCurrPrtn + "' ", "") + _
                      " and IrRmCtg In ('D','C') " + _
                      " group by IdPrtKey, IdCoCd, IdTc, IdYy, IdChr, IdNo, IrRmCtg, RmStnCls, PDesc order by IdPrtKey, IdCoCd, IdTc, IdYy, IdChr, IdNo, IrRmCtg desc, qRmStnCls "
                      
      Call moCn.RepRes(mRep.Subreport3.OpenSubreport, wStnClsSummStr)
    
    
      If adc("UoYN") = "N" Then
        mRep.PICSEC.Suppress = True
      Else
        mRep.PICSEC.Suppress = False
      End If
    ElseIf UCase(adc.MenuCd) = UCase("RepInvCat") Or UCase(adc.MenuCd) = UCase("RepInvCTBCat") Then
      If UCase(adc("UoInVal")) = "S" Then
          wIdFOB = "((Select sum(IrRmAWt * IrStkRt)+ sum((Case when InMulDiv= 'M' then IrSetAVal/InCnvFct else IrSetAVal*InCnvFct end)) from InvRm Join InvHd on InCoCd= IrCoCd and InTc= IrTc and InYy= IrYy and InChr= IrChr and InNo= IrNo where IrPrtKey=IdPrtKey and IrIdIdNo=IdIdNo and IrCoCd= IdCoCd and IrTc= IdTc and " + _
                   "IrYy= IdYy and IrChr= IdChr and IrNo= IdNo and IrSr= IdSr) " + _
                   "  + " + _
                   " (Select sum((Case when InMulDiv= 'M' then iLAVAl/InCnvFct else iLAVAl*InCnvFct end)) from InvLab Join InvHd on InPrtKey=IlPrtKey and InCoCd= IlCoCd and InTc= IlTc and InYy= IlYy and InChr= IlChr and InNo= IlNo where IlCoCd= IdCoCd and IlTc= IdTc and " + _
                   " IlYy= IdYy and IlChr= IdChr and IlNo= IdNo and IlSr= IdSr))"
      End If
    
       wSqlStrg = "Select 0 as qRow, 0 as qCol,OmCoCd, OmYy, OmTc, OmChr, OmNo, OmCmCd,OdOmCmCd, OmDt, OmPoNo, OdSr ,Space(40) as CdOurDmCd, OdDmCd, OdDmSz, OdSfx, OdDmStkYy, OdDmStkChr, OdDmStkNo, DmCtg, DmSalCtg," + _
               "DmTcTyp, OdDmCol, OdPicNm, OdPoNo, OdKt as IdKt, OdVaCtg as IdVaCtg, DmUom ,PNum as IdGldLs, " + _
               "Space(254) as qRmCtgDesc,Space(254) as  qDmCtgDesc,Space(254) as  qDmSalCtgDesc,Space(254) as  qDmColDesc,Space(254) as qDmKtDesc,Space(254) as  qVaCtgDesc, " + _
               "OrRmCtg  as OrRmctg ,OdGldAsWt, OdLabAsWt,  OdSalPrc  as OdSalPrc, OdOrdEnt as qOrdQty, Space(254) as qColDesc, " + _
               "OrPrdWt as qGrsWt, 0 as OdTrayNo   Into TmpA " + _
               " From OrdMst,OrdDsg,Param,CustMst,DsgMst,OrdRm  where 1=2 "
      moCn.CreateTmpTable "TmpA", wSqlStrg
        
     wExcelOrdBy = IIF(wExcelOrdBy <> "", ",", "") + wExcelOrdBy
     '6.1B-15 if perpc option selected then divide the values by IdQty else as it is
     Dim wPerPcQty As String
     wPerPcQty = IIF(adc("UoYN32") = "Y", "IdQty", "1")
    ' ***** Manali 3.8.0 - IdTrayNo added
    '6.1B-15 wPerPcQty added for qGrsWt
    wSqlStrg = "Insert Into TmpA " + _
                 "Select (row_Number() over(" + IIF(adc("UOYN3") = "Y", " ", "Partition By InIdNo") + " Order By InCoCd,InTc,InYy,InChr,InNo" + wExcelOrdBy + ",IdSr)-1)/3 as qRow," + _
                 " row_Number() over(" + IIF(adc("UOYN3") = "Y", " ", "Partition By InIdNo") + " Order By InCoCd,InTc,InYy,InChr,InNo" + wExcelOrdBy + ",IdSr)%3 as qCol," + _
                 "InCoCd as OmCoCd, InYy as OmYy, InTc as OmTc, InChr as OmChr, InNo as OmNo, InCmCd as OmCmCd,OdOmCmCd, InPrnDt as OmDt, " + _
                 "InExpNo as OmPoNo," + IIF(adc("UOYN3") = "Y", "IdIdNo ", "IdSr") + " as OdSr," + wOurDsgCd + " ,IdDmCd as OdDmCd, IdDmSz as OdDmSz, IdSfx as OdSfx, '' as OdDmStkYy, '' as OdDmStkChr, 0 as OdDmStkNo, " + _
                 "DmCtg, DmSalCtg, 'DM' as DmTcTyp, OdDmCol, OdPicNm, OdPoNo, IdKt,IdVaCtg,DmUom," + _
                 "IdGldLs, " + qRmCtgDesc + "," + qDmCtgDesc + "," + qDmSalCtgDesc + "," + qDmColDesc + "," + qDmKtDesc + "," + qVaCtgDesc + _
                 ",IsNull((Select RmCtg from RmMst, InvRm where RmCd=IrRmCd  and  " + _
                 " IrMainMet='Y' and IrCocd = InCocd And IrTc = InTc  " + _
                 "and IrYy = InYy and IrChr=InChr and IrNo=InNo and IrSr=IdSr),'') as OrRmctg /*qMainMetRmCtg*/," + wGldAsWt + " as OdGldAsWt, " + wLabAsWt + " as OdLabAsWt, " + _
                 wIdFOB + "/(case when IdQty> 0 then IdQty else 1 end) as OdSalPrc, IdQty as qOrdQty, " + _
                 "(Select PDesc from Param where PTyp= 'DMCOL' and PMCd= OdDmCol) as qColDesc, " + _
                 "convert (decimal (10, 3),(Select sum(" + wIrRmWt + "/ (case when RmCtg in ('C', 'D') then 5 else 1 end)) " + _
                 "from InvRm, RmMst where " + IIF(gs_Partition = ctCurrPrtn, " RmPrtKey=IrPrtKey and ", "") + " RmCd= IrRmCd and " + IIF(gs_Partition = ctCurrPrtn, " IrPrtKey=IdPrtKey and ", "") + " IrCoCd= IdCoCd and IrTc= IdTc " + _
                 "and IrYy= IdYy and IrChr= IdChr and IrNo= IdNo and IrSr= IdSr )/" + wPerPcQty + ") as qGrsWt, IdTrayNo as OdTrayNo "
                 
      wSqlStrg = wSqlStrg + "From InvHd join InvDsg on IdInIdNo=InIdNo and IdCoCd= InCoCd and IdTc= InTc " + _
                 " and IdYy= InYy and IdChr= InChr and IdNo= InNo " + IIF(gs_Partition = ctCurrPrtn, " And IdPrtKey=InPrtKey ", "") + _
                 " join OrdDsg on OdCoCd= IdCoCd and OdTc= IdExpOdTc and OdYy= IdExpOdYy " + _
                 " and OdChr= IdExpOdChr and OdNo= IdExpOdNo and OdSr= IdExpOdSr " + IIF(gs_Partition = ctCurrPrtn, " And OdPrtKey=IdPrtKey ", "") + ws_CurJoin + _
                 " join DsgMst on DmIdNo=IdDmIdNo and DmTcTyp= 'DM' and DmCd= IdDmCd and DmSz= '' " + IIF(gs_Partition = ctCurrPrtn, " And DmPrtKey=IdPrtKey ", "") + _
                 IIF(adc("UOYN3") = "Y", "Join (Select Max(IdCoCd+IdTc+IdYy+IdChr+Str(IdNo)+Str(IdSr)) as qIdSr , " + _
                 " IdDmCd as qDmCd From InvDsg Ind Join InvHd " + _
                 " on IdInIdNo=InIdNo and IdCoCd= InCoCd and IdTc= InTc and IdYy= InYy and IdChr= InChr and IdNo= InNo  " + IIF(gs_Partition = ctCurrPrtn, " And IdPrtKey=InPrtKey ", "") + ws_CurJoin + _
                  wCnd + IIF(gs_Partition = ctCurrPrtn, " And InPrtKey='" + ctCurrPrtn + "' ", "") + " and InDt=(select max(InDt)  From InvDsg InvD Join InvHd " + _
                 " on IdInIdNo=InIdNo and IdCoCd= InCoCd and IdTc= InTc and IdYy= InYy and IdChr= InChr and IdNo= InNo " + IIF(gs_Partition = ctCurrPrtn, " And IdPrtKey=InPrtKey ", "") + " and InvD.IdDmCd=Ind.IdDmCd " + _
                 ws_CurJoin + wCnd + IIF(gs_Partition = ctCurrPrtn, " And InPrtKey='" + ctCurrPrtn + "' ", "") + " and InvD.IdDmCd=Ind.IdDmCd Group By IdDmCd )Group By IdDmCd ) Inv  " + _
                 " On IdCoCd+IdTc+IdYy+IdChr+Str(IdNo)+Str(IdSr)=qIdSr  ", " ") + _
                  wCnd + " Order By InCoCd,InTc,InYy,InChr,InNo" + wExcelOrdBy '+ ",IdSr"
                 '(index= InvDsg) '(index= OrdDsg)
                 '"and IrYy= IdYy and IrChr= IdChr and IrNo= IdNo and IrSr= IdSr)/ " + _
                 '"(case when IdQty> 0 then IdQty else 1 end) as qGrsWt " + _

        moCn.Execute (wSqlStrg)

     wSqlStrg = " Select  " + IIF(adc("UOYN3") = "Y", " OmCoCd,'' as OmYy,OmTc,'' as OmChr,0 as OmNo ", " OmCoCd,OmYy,OmTc,OmChr,OmNo") + ",qRow,Max(OmCmCd) as OmCmCd ,Max(OdOmCmCd) as OdOmCmCd,Max(OmDt) as OmDt,Max(OmPoNo) as OmPoNo, " + _
               " Max((Case when qCol=1 Then   OdSr else 0 End)) as OdSr1,Max((Case when qCol=1 Then   CdOurDmCd else '' End)) as CdOurDmCd1, " + _
               " Max((Case when qCol=1 Then   OdDmCd else '' End)) as OdDmCd1,Max((Case when qCol=1 Then   OdDmSz else '' End)) as OdDmSz1, " + _
               " Max((Case when qCol=1 Then   OdSfx else '' End)) as OdSfx1,Max((Case when qCol=1 Then   OdDmStkYy+'/'+OdDmStkChr+'/'+ltrim(str(OdDmStkNo)) else '' End)) as OdDmId1, " + _
               " Max((Case when qCol=1 Then   DmCtg else '' End)) as DmCtg1,Max((Case when qCol=1 Then   DmSalCtg else '' End)) as DmSalCtg1, " + _
               " Max((Case when qCol=1 Then   DmTcTyp else '' End)) as DmTcTyp1,Max((Case when qCol=1 Then   OdDmCol else '' End)) as OdDmCol1, " + _
               " Max((Case when qCol=1 Then   OdPicNm else '' End)) as OdPicNm1,Max((Case when qCol=1 Then   OdPoNo else '' End)) as OdPoNo1, " + _
               " Max((Case when qCol=1 Then   IdKt else '' End)) as IdKt1,Max((Case when qCol=1 Then   IdVaCtg else '' End)) as IdVaCtg1, " + _
               " Max((Case when qCol=1 Then   DmUom else '' End)) as DmUom1,Max((Case when qCol=1 Then   IdGldLs else '' End)) as IdVaCtg1, " + _
               " Max((Case when qCol=1 Then   qRmCtgDesc else '' End)) as qRmCtgDesc1,Max((Case when qCol=1 Then   qDmCtgDesc else '' End)) as qDmCtgDesc1, " + _
               " Max((Case when qCol=1 Then   qDmSalCtgDesc else '' End)) as qDmSalCtgDesc1,Max((Case when qCol=1 Then   qDmColDesc else '' End)) as qDmColDesc1,Max((Case when qCol=1 Then qDmKtDesc else '' End)) as qDmKtDesc1," + _
               " Max((Case when qCol=1 Then   qVaCtgDesc else '' End)) as qVaCtgDesc1,Max((Case when qCol=1 Then   OrRmctg else '' End)) as OrRmctg1, " + _
               " Max((Case when qCol=1 Then   OdGldAsWt else 0 End)) as OdGldAsWt1,Max((Case when qCol=1 Then   OdLabAsWt else 0 End)) as OdLabAsWt1, " + _
               " Max((Case when qCol=1 Then   OdSalPrc else 0 End)) as OdSalPrc1,Max((Case when qCol=1 Then   OdLabAsWt else 0 End)) as OdLabAsWt1, " + _
               " Max((Case when qCol=1 Then   qOrdQty else 0 End)) as qOrdQty1,Max((Case when qCol=1 Then   qColDesc else '' End)) as qColDesc1, " + _
               " Max((Case when qCol=1 Then   qGrsWt else 0 End)) as qGrsWt1,Max((Case when qCol=1 Then   OdTrayNo else '' End)) as OdTrayNo1, "
      wSqlStrg = wSqlStrg + "Max((Case when qCol=2 Then   OdSr else 0 End)) as OdSr2,Max((Case when qCol=2 Then   CdOurDmCd else '' End)) as CdOurDmCd2, " + _
                        " Max((Case when qCol=2 Then   OdDmCd else '' End)) as OdDmCd2,Max((Case when qCol=2 Then   OdDmSz else '' End)) as OdDmSz2," + _
                        " Max((Case when qCol=2 Then   OdSfx else '' End)) as OdSfx2,Max((Case when qCol=2 Then  OdDmStkYy+'/'+OdDmStkChr+'/'+ltrim(str(OdDmStkNo)) else '' End)) as OdDmId2, " + _
                        " Max((Case when qCol=2 Then   DmCtg else '' End)) as DmCtg2,Max((Case when qCol=2 Then   DmSalCtg else '' End)) as DmSalCtg2, " + _
                        " Max((Case when qCol=2 Then   DmTcTyp else '' End)) as DmTcTyp2,Max((Case when qCol=2 Then   OdDmCol else '' End)) as OdDmCol2, " + _
                        " Max((Case when qCol=2 Then   OdPicNm else '' End)) as OdPicNm2,Max((Case when qCol=2 Then   OdPoNo else '' End)) as OdPoNo2, " + _
                        " Max((Case when qCol=2 Then   IdKt else '' End)) as IdKt2,Max((Case when qCol=2 Then   IdVaCtg else '' End)) as IdVaCtg2, " + _
                        " Max((Case when qCol=2 Then   DmUom else '' End)) as DmUom2,Max((Case when qCol=2 Then   IdGldLs else '' End)) as IdVaCtg2, " + _
                        " Max((Case when qCol=2 Then   qRmCtgDesc else '' End)) as qRmCtgDesc2,Max((Case when qCol=2 Then   qDmCtgDesc else '' End)) as qDmCtgDesc2, " + _
                        " Max((Case when qCol=2 Then   qDmSalCtgDesc else '' End)) as qDmSalCtgDesc2,Max((Case when qCol=2 Then   qDmColDesc else '' End)) as qDmColDesc2,Max((Case when qCol=2 Then   qDmKtDesc else '' End)) as qDmKtDesc2, " + _
                        " Max((Case when qCol=2 Then   qVaCtgDesc else '' End)) as qVaCtgDesc2,Max((Case when qCol=2 Then   OrRmctg else '' End)) as OrRmctg2, " + _
                        " Max((Case when qCol=2 Then   OdGldAsWt else 0 End)) as OdGldAsWt2,Max((Case when qCol=2 Then   OdLabAsWt else 0 End)) as OdLabAsWt2, " + _
                        " Max((Case when qCol=2 Then   OdSalPrc else 0 End)) as OdSalPrc2,Max((Case when qCol=2 Then   OdLabAsWt else 0 End)) as OdLabAsWt2," + _
                        " Max((Case when qCol=2 Then   qOrdQty else 0 End)) as qOrdQty2,Max((Case when qCol=2 Then   qColDesc else '' End)) as qColDesc2, " + _
                        " Max((Case when qCol=2 Then   qGrsWt else 0 End)) as qGrsWt2,Max((Case when qCol=2 Then   OdTrayNo else '' End)) as OdTrayNo2, "
      wSqlStrg = wSqlStrg + " Max((Case when qCol=0 Then   OdSr else 0 End)) as OdSr3,Max((Case when qCol=0 Then   CdOurDmCd else '' End)) as CdOurDmCd3, " + _
                            " Max((Case when qCol=0 Then   OdDmCd else '' End)) as OdDmCd3,Max((Case when qCol=0 Then   OdDmSz else '' End)) as OdDmSz3, " + _
                            " Max((Case when qCol=0 Then   OdSfx else '' End)) as OdSfx3,Max((Case when qCol=0 Then   OdDmStkYy+'/'+OdDmStkChr+'/'+ltrim(str(OdDmStkNo)) else '' End)) as OdDmId3, " + _
                            " Max((Case when qCol=0 Then   DmCtg else '' End)) as DmCtg3,Max((Case when qCol=0 Then   DmSalCtg else '' End)) as DmSalCtg3, " + _
                            " Max((Case when qCol=0 Then   DmTcTyp else '' End)) as DmTcTyp3,Max((Case when qCol=0 Then   OdDmCol else '' End)) as OdDmCol3, " + _
                            " Max((Case when qCol=0 Then   OdPicNm else '' End)) as OdPicNm3,Max((Case when qCol=0 Then   OdPoNo else '' End)) as OdPoNo3, " + _
                            " Max((Case when qCol=0 Then   IdKt else '' End)) as IdKt3,Max((Case when qCol=0 Then   IdVaCtg else '' End)) as IdVaCtg3, " + _
                            " Max((Case when qCol=0 Then   DmUom else '' End)) as DmUom3,Max((Case when qCol=0 Then   IdGldLs else '' End)) as IdVaCtg3, " + _
                            " Max((Case when qCol=0 Then   qRmCtgDesc else '' End)) as qRmCtgDesc3,Max((Case when qCol=0 Then   qDmCtgDesc else '' End)) as qDmCtgDesc3, " + _
                            " Max((Case when qCol=0 Then   qDmSalCtgDesc else '' End)) as qDmSalCtgDesc3,Max((Case when qCol=0 Then   qDmColDesc else '' End)) as qDmColDesc3,Max((Case when qCol=0 Then   qDmKtDesc else '' End)) as qDmKtDesc3, " + _
                            " Max((Case when qCol=0 Then   qVaCtgDesc else '' End)) as qVaCtgDesc3,Max((Case when qCol=0 Then   OrRmctg else '' End)) as OrRmctg3, " + _
                            " Max((Case when qCol=0 Then   OdGldAsWt else 0 End)) as OdGldAsWt3,Max((Case when qCol=0 Then   OdLabAsWt else 0 End)) as OdLabAsWt3, " + _
                            " Max((Case when qCol=0 Then   OdSalPrc else 0 End)) as OdSalPrc3,Max((Case when qCol=0 Then   OdLabAsWt else 0 End)) as OdLabAsWt3, " + _
                            " Max((Case when qCol=0 Then   qOrdQty else 0 End)) as qOrdQty3,Max((Case when qCol=0 Then   qColDesc else '' End)) as qColDesc3, " + _
                            " Max((Case when qCol=0 Then   qGrsWt else 0 End)) as qGrsWt3,Max((Case when qCol=0 Then   OdTrayNo else '' End)) as OdTrayNo3 " + _
                            "  From TmpA Group By   " + IIF(adc("UOYN3") = "Y", "OmcoCd,OmTc ", " OmCoCd,OmYy,OmTc,OmChr,OmNo") + ",qRow Order By OmCoCd,OmYy,OmTc,OmChr,OmNo,qRow "
    
      
      adc.RepRecSource = wSqlStrg + ctMaxDopOpt  '****** Sachin 3.01 added ctmaxdopopt
      
      '*******************Geeta****************emr206 ***********Replace RmCd with " + wOurRmCd + "
      '****** Sachin 3.02 - Id fields in Joins
      '****** Manali 3.03 - 27/06/08 - Pd Mod - 'L' added
      '6.1B-15 wPerPcQty added for Qty & Wt
      wRmDet = "Select " + IIF(adc("UOYN3") = "Y", " '' as OdYy, IdTc as OdTc, '' as OdChr, 0 as OdNo, IdIdNo as OdSr ", " IdYy as OdYy, IdTc as OdTc, IdChr as OdChr, IdNo as OdNo, IdSr as OdSr") + ", IrMainMet as OrMainMet, " + _
               "RmCtg as OrRmCtg, (case RmCtg when 'G' then 1 when 'P' then 2 when 'S' then 3 when 'L' then 9 when 'D' then 4 " + _
               "when 'C' then 5 when 'X' then 6 when 'A' then 7 when 'M' then 8 else 0 end) " + _
               "as qTag, RmSCtg as OrRmSCtg,'' as OrLotNo," + IIF(adc("UoYN30") = "Y", "(Case When RmIntQly='' Then IrRmCd Else RmIntQly End)", " IrRmCd") + " as OrRmCd, " + _
               "IrRmQty/" + wPerPcQty + " as OrQty, convert (decimal (10, 3)," + wIrRmWt + "/" + wPerPcQty + ") as OrWt " + _
               "From InvDsg join InvHd on InIdNo=IdInIdNo and  InCoCd= IdCoCd and InTc= IdTc and " + _
               " InYy= IdYy and InChr= IdChr and InNo= IdNo " + IIF(gs_Partition = ctCurrPrtn, " And InPrtKey=IdPrtKey ", "") + _
               " join InvRm on IrIdIdNo=IdIdNo and IrCoCd= IdCoCd and IrTc= IdTc and IrYy= IdYy and " + _
               " IrChr= IdChr and IrNo= IdNo and IrSr= IdSr " + IIF(gs_Partition = ctCurrPrtn, " And IrPrtKey=IdPrtKey ", "") + _
               " join RmMst on RmCd= IrRmCd " + IIF(gs_Partition = ctCurrPrtn, " And RmPrtKey=IrPrtKey ", "") + ws_CurJoin + _
               wCnd + IIF(gs_Partition = ctCurrPrtn, " And InPrtKey='" + ctCurrPrtn + "' ", "")
    
               '(index= InvRm)
      'Call moCn.RepRes(mRep.Subreport1.OpenSubreport, wRmDet + ctMaxDopOpt)  '****** Sachin 3.01 added ctmaxdopopt
        
     Call moCn.RepRes(mRep.Subreport1.OpenSubreport, wRmDet + ctMaxDopOpt)  '****** Sachin 3.01 added ctMaxDopOpt
     Call moCn.RepRes(mRep.Subreport2.OpenSubreport, wRmDet + ctMaxDopOpt)  '****** Sachin 3.01 added ctMaxDopOpt
     Call moCn.RepRes(mRep.Subreport3.OpenSubreport, wRmDet + ctMaxDopOpt)  '****** Sachin 3.01 added ctMaxDopOpt
      
     Set wo_rsPic = moCn.OpenRes("Select DmTcTyp as DmTcTyp,OdDmCd as DmCd,OdDmSz as DmSz,OmCoCd,OmTc,OmYy,OmChr,OmNo,OdSr,OdPicNm as PicNm from TmpA" + _
                                  " order by OmCoCd, OmYy, OmTc, OmChr, OmNo,qRow,(Case when qCol=0 Then 3 Else qCol End) ")
     If wo_rsPic.RecCount > 0 Then CmdChgImg.Visible = True
      
    End If
    
    '6.1_27 if records exists then starts picture copy
    If (UCase(adc.MenuCd) = UCase("RepDetInv") Or UCase(adc.MenuCd) = UCase("RepInvCat")) And adc("UOYN31") = "Y" Then
      Dim wRsPicCpy As MDORowSet
      Set wRsPicCpy = moCn.OpenRes("Select * From TmpA")
      
      If wRsPicCpy.RecCount > 0 Then
        MWLib.BeginProcess Me, "Starting Picture Copy"
        Call PicCopy(wRsPicCpy, adc("wTarPicPath"))
        MWLib.EndProcess Me
        MsgBox "Picture Copy Completed"
      End If
    End If
  
    'Rep.Subreport1.OpenSubreport.Database.SetDataSource
  
    'Rm Lot No newly added in Detail Order Printing and Detail Voucher Printing,
    'So make visible false in Detail invoice printing report
    Rep.RmDet1.Suppress = True
    Rep.RmDetLine.Suppress = True
    Rep.TxtRmLotNo.Suppress = True
    
    If UCase(adc.MenuCd) = UCase("RepDetInv") Or UCase(adc.MenuCd) = UCase("RepAvgInv") Or UCase(adc.MenuCd) = UCase("RepInvCat") Then
      mRep.TxtCmCd.SetText "Customer"
    ElseIf UCase(adc.MenuCd) = UCase("RepDetInvCTB") Or UCase(adc.MenuCd) = UCase("RepAvgInvCTB") Or UCase(adc.MenuCd) = UCase("RepInvCTBCat") Then
      mRep.TxtCmCd.SetText "Contractor"
    End If
    
    If UCase(adc.MenuCd) = UCase("RepDetInv") Or UCase(adc.MenuCd) = UCase("RepAvgInv") Then
      Rep.TxtVchNo.SetText "Inv Vch"
      Rep.TxtPgVchNo.SetText "Inv Vch"
      Rep.TxtVchDt.SetText "Inv Dt"
      Rep.TxtPgVchDt.SetText "Inv Dt"
      Rep.TxtDocNo.SetText "Exp No"
      Rep.TxtDsgQty.SetText "Inv Qty"
      Rep.TxtDsgQty1.SetText "Inv Qty"
    ElseIf UCase(adc.MenuCd) = UCase("RepDetInvCTB") Or UCase(adc.MenuCd) = UCase("RepAvgInvCTB") Then
      Rep.TxtVchNo.SetText "CTB Vch"
      Rep.TxtPgVchNo.SetText "CTB Vch"
      Rep.TxtVchDt.SetText "CTB Dt"
      Rep.TxtPgVchDt.SetText "CTB Dt"
      Rep.TxtDocNo.SetText "Exp No"
      Rep.TxtDsgQty.SetText "CTB Qty"
      Rep.TxtDsgQty1.SetText "CTB Qty"
    ElseIf UCase(adc.MenuCd) = UCase("RepInvCat") Then
      mRep.TxtVchNo.SetText "Inv Vch"
      mRep.TxtVchDt.SetText "Inv Dt"
      mRep.TxtDocNo.SetText "Exp No"
      mRep.TxtPgVchNo.SetText "Inv Vch"
      mRep.TxtPgVchDt.SetText "Inv Dt"
      
      
    ElseIf UCase(adc.MenuCd) = UCase("RepInvCTBCat") Then
      mRep.TxtVchNo.SetText "CTB Vch"
      mRep.TxtPgVchNo.SetText "CTB Vch"
      mRep.TxtVchDt.SetText "CTB Dt"
      mRep.TxtPgVchDt.SetText "CTB Dt"
      mRep.TxtDocNo.SetText "Exp No"
    End If
    
    If UCase(adc.MenuCd) = UCase("RepDetInv") Or UCase(adc.MenuCd) = UCase("RepDetInvCTB") Then
      GRP_REP.SetFormula mRep, "wGrp3", "{rdo.qTag}+ToText({rdo.qMainMet}, '#')+ToText({rdo.OrSrNo}, '##0')+{rdo.RmCtg}+{rdo.OrRmSCtg}"
      If UCase(adc.MenuCd) = UCase("RepDetInv") Then mRep.TxtHead.SetText "Detail Invoice Printing"
      If UCase(adc.MenuCd) = UCase("RepDetInvCTB") Then mRep.TxtHead.SetText "Detail CTB Printing"
      mRep.PHDET.Suppress = False
      mRep.PhAvg.Suppress = True
      mRep.AvgDet.Suppress = True
      If adc("UoRmGrp") = "N" Then
        mRep.RmHead.Suppress = True
        mRep.RmDet.Suppress = True
      End If
      If adc("UoLabGrp") = "N" Then
        mRep.LABHEAD.Suppress = True
        mRep.LABDET.Suppress = True
      End If
      Rep.hOrRmPtr.Suppress = True: Rep.FldOrRmPtr.Suppress = True
      ' Zubin 212 (26/02/06)
      GRP_REP.SetFormula mRep, "wGldRtDAFlg", IIF(moCn.GetFldVal("Select HGldRtDAYN from Head where HCd= '" + ctSelfCmCd + "' and HCoCd= '" + adc("UOCOCDFR") + "'") = "Y", "'D'", "'A'")
      ' Zubin 212 (26/02/06)
      ' **** Manali 3.6.0
      If UCase(adc("UoInDet")) <> "TRNO/SR" Then
        Rep.FldTrayNo.Suppress = True: Rep.HDTrayNo.Suppress = True: Rep.HATrayNo.Suppress = True
      Else
        Rep.FldTrayNo.Suppress = False: Rep.HDTrayNo.Suppress = False: Rep.HATrayNo.Suppress = False
      End If
      ' **** Manali 3.6.0
      
      'If adc("UoYn4") <> "Y" Then Rep.SecLabSumm.Suppress = True
      
    ElseIf UCase(adc.MenuCd) = UCase("RepAvgInv") Or UCase(adc.MenuCd) = UCase("RepAvgInvCTB") Then
      GRP_REP.SetFormula mRep, "wGrp3", "{rdo.qTag}+ToText({rdo.qMainMet}, '#')+{rdo.RmCtg}+{rdo.OrRmSCtg}"
      If UCase(adc.MenuCd) = UCase("RepAvgInv") Then mRep.TxtHead.SetText "Average Invoice"
      If UCase(adc.MenuCd) = UCase("RepAvgInvCTB") Then mRep.TxtHead.SetText "Average CTB"
      mRep.PhAvg.Suppress = False
      mRep.PHDET.Suppress = True
      mRep.AvgDet.Suppress = False
      mRep.RmHead.Suppress = True
      mRep.RmDet.Suppress = True
      mRep.LABHEAD.Suppress = True
      mRep.LABDET.Suppress = True
      mRep.FldDiaQty.Suppress = True
      mRep.FldDiaWt.Suppress = True
      mRep.FldColQty.Suppress = True
      mRep.FldColWt.Suppress = True
      ' **** Manali 3.6.0
      If UCase(adc("UoInDet")) <> "TRNO/SR" Then
        Rep.FldTrayNo.Suppress = True: Rep.HDTrayNo.Suppress = True: Rep.HATrayNo.Suppress = True
      Else
        Rep.FldTrayNo.Suppress = False: Rep.HDTrayNo.Suppress = False: Rep.HATrayNo.Suppress = False
      End If
      ' **** Manali 3.6.0
      
    ElseIf UCase(adc.MenuCd) = UCase("RepInvCat") Then
      mRep.TxtHead.SetText "Invoice Catalogue"
    ElseIf UCase(adc.MenuCd) = UCase("RepInvCTBCat") Then
      mRep.TxtHead.SetText "CTB Catalogue"
    End If
    
    
    If adc("UoYN28") = "Y" Then
        Rep.TxtGldAsWtAvg.SetText "GldAsWt": Rep.TxtGldAsWtDet.SetText "GldAsWt"
    Else
        Rep.TxtGldAsWtAvg.SetText "Net Wt": Rep.TxtGldAsWtDet.SetText "Net Wt"
    End If
    
    '*** JAY 2.14 *** Rdo.OdSr Replaced with @OdSr
    If adc("UoYN1") = "Y" And (UCase(adc.MenuCd) = UCase("RepDetInv") Or UCase(adc.MenuCd) = UCase("RepAvgInv") Or _
       UCase(adc.MenuCd) = UCase("RepDetInvCTB") Or UCase(adc.MenuCd) = UCase("RepAvgInvCTB")) Then
      GRP_REP.SetFormula mRep, "g2_GrWt", "Sum ({@wRmWt}, {@OdSr})/ {rdo.qOrdQty}"
      'GRP_REP.SetFormula mRep, "wOdGldAsWt", "{rdo.OdGldAsWt}/ {rdo.qOrdQty}"
      ' ***** Manali 3.10.0 - 14/03/12 - Show Net Wt instead of Gld As Wt
      If adc("UoYN28") = "Y" Then
        GRP_REP.SetFormula Rep, "wOdGldAsWt", "{rdo.OdGldAsWt}/ {rdo.qOrdQty}"
      Else
        GRP_REP.SetFormula Rep, "wOdGldAsWt", "Sum({@wNetWt}, {@OdSr})/{rdo.qOrdQty}"
      End If
      ' ***** Manali 3.10.0 - 14/03/12 - Show Net Wt instead of Gld As Wt
      
      GRP_REP.SetFormula mRep, "wOdLabAsWt", "{rdo.OdLabAsWt}/ {rdo.qOrdQty}"
      GRP_REP.SetFormula mRep, "g2_DiaWt", "Sum ({@wDiaWt}, {@OdSr})/ {rdo.qOrdQty}"
      GRP_REP.SetFormula mRep, "g2_DiaQty", "Sum ({@wDiaQty}, {@OdSr})/ {rdo.qOrdQty}"
      GRP_REP.SetFormula mRep, "g2_ColWt", "Sum ({@wColWt}, {@OdSr})/ {rdo.qOrdQty}"
      GRP_REP.SetFormula mRep, "g2_ColQty", "Sum ({@wColQty}, {@OdSr})/ {rdo.qOrdQty}"
      GRP_REP.SetFormula mRep, "g2_LabVal", "(Sum ({rdo.OrSetSalVal}, {@OdSr})+Sum ({rdo.OlSalVal}, {@OdSr}))/ {rdo.qOrdQty}"
      GRP_REP.SetFormula mRep, "g3_RmWt", "Sum ({rdo.OrWt}, {@wGrp3})/ {rdo.qOrdQty}"
      GRP_REP.SetFormula mRep, "g3_RmQty", "Sum ({rdo.OrQty}, {@wGrp3})/ {rdo.qOrdQty}"
      GRP_REP.SetFormula mRep, "g3_RmVal", "Sum ({rdo.OrSalVal}, {@wGrp3})/ {rdo.qOrdQty}"
      
      GRP_REP.SetFormula mRep, "wOrQty", "{rdo.OrQty}/ {rdo.qOrdQty}"
      GRP_REP.SetFormula mRep, "wOrWt", "{rdo.OrWt}/ {rdo.qOrdQty}"
      GRP_REP.SetFormula mRep, "wOrSalVal", "{rdo.OrSalVal}/ {rdo.qOrdQty}"
      GRP_REP.SetFormula mRep, "wOrSetSalVal", "{rdo.OrSetSalVal}/ {rdo.qOrdQty}"
      GRP_REP.SetFormula mRep, "wOlSalVal", "{rdo.OlSalVal}/ {rdo.qOrdQty}"
      
      'GRP_REP.SetFormula mRep, "wTotLabVal", "({rdo.OrSetSalVal}+{rdo.OlSalVal})/ {rdo.qOrdQty}"
      GRP_REP.SetFormula mRep, "wTotLabVal", "({rdo.OrSetSalVal}+{rdo.OlSalVal})"
      
      mRep.TxtAvgWt.SetText "(Average Weights)"
    ElseIf adc("UoYN1") = "N" And (UCase(adc.MenuCd) = UCase("RepDetInv") Or UCase(adc.MenuCd) = UCase("RepAvgInv") Or _
       UCase(adc.MenuCd) = UCase("RepDetInvCTB") Or UCase(adc.MenuCd) = UCase("RepAvgInvCTB")) Then
      GRP_REP.SetFormula mRep, "g2_GrWt", "Sum ({@wRmWt}, {@OdSr})"
      'GRP_REP.SetFormula mRep, "wOdGldAsWt", "{rdo.OdGldAsWt}"
      ' ***** Manali 3.10.0 - 14/03/12 - Show Net Wt instead of Gld As Wt
      If adc("UoYN28") = "Y" Then
        GRP_REP.SetFormula Rep, "wOdGldAsWt", "{rdo.OdGldAsWt}"
      Else
        GRP_REP.SetFormula Rep, "wOdGldAsWt", "Sum({@wNetWt}, {@OdSr})"
      End If
      ' ***** Manali 3.10.0 - 14/03/12 - Show Net Wt instead of Gld As Wt
      
      
      GRP_REP.SetFormula mRep, "wOdLabAsWt", "{rdo.OdLabAsWt}"
      GRP_REP.SetFormula mRep, "g2_DiaWt", "Sum ({@wDiaWt}, {@OdSr})"
      GRP_REP.SetFormula mRep, "g2_DiaQty", "Sum ({@wDiaQty}, {@OdSr})"
      GRP_REP.SetFormula mRep, "g2_ColWt", "Sum ({@wColWt}, {@OdSr})"
      GRP_REP.SetFormula mRep, "g2_ColQty", "Sum ({@wColQty}, {@OdSr})"
      GRP_REP.SetFormula mRep, "g2_LabVal", "Sum ({rdo.OrSetSalVal}, {@OdSr})+Sum ({rdo.OlSalVal}, {@OdSr})"
      GRP_REP.SetFormula mRep, "g3_RmWt", "Sum ({rdo.OrWt}, {@wGrp3})"
      GRP_REP.SetFormula mRep, "g3_RmQty", "Sum ({rdo.OrQty}, {@wGrp3})"
      GRP_REP.SetFormula mRep, "g3_RmVal", "Sum ({rdo.OrSalVal}, {@wGrp3})"
      
      GRP_REP.SetFormula mRep, "wOrQty", "{rdo.OrQty}"
      GRP_REP.SetFormula mRep, "wOrWt", "{rdo.OrWt}"
      GRP_REP.SetFormula mRep, "wOrSalVal", "{rdo.OrSalVal}"
      GRP_REP.SetFormula mRep, "wOrSetSalVal", "{rdo.OrSetSalVal}"
      GRP_REP.SetFormula mRep, "wOlSalVal", "{rdo.OlSalVal}"
      GRP_REP.SetFormula mRep, "wTotLabVal", "({rdo.OrSetSalVal}+{rdo.OlSalVal})"
      
      mRep.TxtAvgWt.SetText "(Total Weights)"
    End If
    '*** JAY 2.14 *** Rdo.OdSr Replaced with @OdSr
    '****************************Geeta****************************
    '************Setting Group2 & Group3 Formula***********************
    If (UCase(adc.MenuCd) = UCase("RepDetInv") Or UCase(adc.MenuCd) = UCase("RepAvgInv")) Then  'Or UCase(adc.MenuCd) = UCase("RepInvCat")) Then
     ' If Trim(adc("UoInGrp")) <> "" Then
     
     
     
     If wGrpFormula1 <> "" Then
        GRP_REP.SetFormula mRep, "wGrp4", wGrpFormula1
        GRP_REP.SetFormula mRep, "wGrp4Desc", wGrpDescFormula1
      End If
      
  ' **** Zubin 211 **** '
  '''''    If Trim(ADC("UoInGrp1")) <> "" Or ADC("UoYN11") = "Y" Then
      '(Bef 20/06/05) If Trim(adc("UoInGrp1")) <> "" Then
      If Trim(adc("UoInGrp1")) <> "" Or wDetSrt <> "" Then
  ' **** Zubin 211 **** '
        GRP_REP.SetFormula mRep, "wGrp5", wGrpFormula2
        GRP_REP.SetFormula mRep, "wGrp5Desc", wGrpDescFormula2
      End If
    End If
    
    '*** jay 2.14 ***
    GRP_REP.SetFormula mRep, "wGrpDesc", IIF(wGrpDescFormula1 = "", "", "{@wGrp4Desc}") + _
                                         IIF(wGrpDescFormula1 = "" Or wGrpDescFormula2 = "", "", "+' / '+") + _
                                         IIF(wGrpDescFormula2 = "", "", "{@wGrp5Desc}")
    '*** jay 2.14 ***
  
    
    'Geeta***2.10
    If UCase(adc.MenuCd) = UCase("RepInvCat") And adc("UOYN3") = "Y" Then
        GRP_REP.SetFormula mRep, "wOrder", ""
        mRep.PhOrd.Suppress = True
        mRep.PHGRP.Suppress = True
    End If
    
    
    ' **** Manali 3.03 - 07/07/08 - RmSzDesc
    ' ***** Manali 3.9.0 - New Formula for OrLme and wLnSv is changed too - as wLnSv was string n thousand separator is not suppressed in that

    GRP_REP.SetFormula mRep, "wOrLmeSal", " If ({rdo.RmCtg} ='G' or {rdo.RmCtg} ='P' or {rdo.RmCtg} ='S' or {rdo.RmCtg} ='M' or {rdo.RmCtg} ='L') " + _
                                      " and Round({rdo.OrLmeSal}, 3) <> 0.00 and '" + adc("UoYn14") + "'='Y' Then Round({rdo.OrLmeSal}, 3) else 0.00 "
    
    GRP_REP.SetFormula mRep, "wLnSv", " If ({rdo.RmCtg} ='G' or {rdo.RmCtg} ='P' or {rdo.RmCtg} ='S' or {rdo.RmCtg} ='M' or {rdo.RmCtg} ='L') " + _
                                      " then (if Round({rdo.OrLmeSal}, 3) = 0.00 Then Trim (ToText ({rdo.OrLn1}, 3)) )  " + _
                                      " else (If Round({rdo.OrLn1}, 3)= 0.01 Then '+0 ' " + _
                                      " else if Round({rdo.OrLn1}, 3)= 0.02 Then '+00 '  " + _
                                      " else if Round({rdo.OrLn1}, 3)= 0.03 Then '+000 '  " + _
                                      " else if Round({rdo.OrLn1}, 3)= 0 Then ' '  " + _
                                      " else Trim (ToText ({rdo.OrLn1}, 3))+'*'+Trim (ToText ({rdo.OrLn2}, 3))+'*'+Trim (ToText ({rdo.OrLn3}, 3)) ) "

'                                      Trim (ToText ({rdo.OrLn1}, 3)) ) "
'    GRP_REP.SetFormula mRep, "wOrLmeSal", "If ({rdo.RmCtg} ='G' or {rdo.RmCtg} ='P' or {rdo.RmCtg} ='S' or {rdo.RmCtg} ='L') or {rdo.RmCtg} ='M') " + _
                                      "and Round({rdo.OrLmeSal}, 3) <> 0.00 and '" + adc("UoYn14") + "'='Y' Then {rdo.OrLmeSal} else 0.00 "
    
'    GRP_REP.SetFormula mRep, "wLnSv", "If ({rdo.RmCtg} ='G' or {rdo.RmCtg} ='P' or {rdo.RmCtg} ='S' or {rdo.RmCtg} ='L' or {rdo.RmCtg} ='M') " + _
                                      "Then If Round({rdo.OrLmeSal}, 3) = 0.00  Then Trim (ToText ({rdo.OrLn1}, 3)) else '' " + _
                                      "else (If {rdo.qRmSzDesc}<> '' Then {rdo.qRmSzDesc} Else " + _
                                      "(If Round({rdo.OrLn1}, 3)= 0.01 Then '+0 ' " + _
                                      "else if Round({rdo.OrLn1}, 3)= 0.02 Then '+00 '  " + _
                                      "else if Round({rdo.OrLn1}, 3)= 0.03 Then '+000 '  " + _
                                      "else if Round({rdo.OrLn1}, 3)= 0 Then ' '  " + _
                                      "else Trim (ToText ({rdo.OrLn1}, 3)) )) "
    ' Bef 3.9.0
'    GRP_REP.SetFormula mRep, "wLnSv", "If ({rdo.RmCtg} ='G' or {rdo.RmCtg} ='P' or {rdo.RmCtg} ='S' or {rdo.RmCtg} ='L') " + _
'                                      "and Round({rdo.OrLmeSal}, 3) <> 0.00 and '" + adc("UoYn14") + "'='Y' Then Trim (ToText ({rdo.OrLmeSal}, 3)) " + _
'                                      "else (If {rdo.qRmSzDesc}<> '' Then {rdo.qRmSzDesc} Else " + _
'                                      "(If Round({rdo.OrLn1}, 3)= 0.01 Then '+0 ' " + _
'                                      "else if Round({rdo.OrLn1}, 3)= 0.02 Then '+00 '  " + _
'                                      "else if Round({rdo.OrLn1}, 3)= 0.03 Then '+000 '  " + _
'                                      "else if Round({rdo.OrLn1}, 3)= 0 Then ' '  " + _
'                                      "else Trim (ToText ({rdo.OrLn1}, 3)) )) "
'    If Not (mRep.OrLn1.Value = "+0 " Or mRep.OrLn1.Value = "+00 " Or mRep.OrLn1.Value = "+000 ") Then
'      mRep.OrLn1.ThousandsSeparators = False
'      Rep.OrLn1.ZeroValueString.
'    End If
    ' ***** Manali 3.03 - 29/06/08 - Lme Rt
    'GRP_REP.SetFormula mRep, "wLnSv", "If ({rdo.RmCtg} ='G' or {rdo.RmCtg} ='P' or {rdo.RmCtg} ='S' or {rdo.RmCtg} ='L') " + _
                                      "and Round({rdo.OrLmeSal}, 3) <> 0.00  Then Trim (ToText ({rdo.OrLmeSal}, 3)) " + _
                                      "else (If Round({rdo.OrLn1}, 3)= 0.01 Then '+0 ' " + _
                                      "else if Round({rdo.OrLn1}, 3)= 0.02 Then '+00 '  " + _
                                      "else if Round({rdo.OrLn1}, 3)= 0.03 Then '+000 '  " + _
                                      "else if Round({rdo.OrLn1}, 3)= 0 Then ' '  " + _
                                      "else Trim (ToText ({rdo.OrLn1}, 3)) ) "
    ' **** Bef 3.03 - Manali *****
    'GRP_REP.SetFormula mRep, "wLnSv", "If Round({rdo.OrLn1}, 3)= 0.01 Then '+0 ' " + _
                                         "else if Round({rdo.OrLn1}, 3)= 0.02 Then '+00 '  " + _
                                         "else if Round({rdo.OrLn1}, 3)= 0.03 Then '+000 '  " + _
                                         "else if Round({rdo.OrLn1}, 3)= 0 Then ' '  " + _
                                         "else Trim (ToText ({rdo.OrLn1}, 4)) "
     '****************************Geeta****************************
     '***********geeta*****emr207******
     If (UCase(adc.MenuCd) = UCase("RepDetInv") Or UCase(adc.MenuCd) = UCase("RepAvgInv") Or _
          UCase(adc.MenuCd) = UCase("RepDetInvCTB") Or UCase(adc.MenuCd) = UCase("RepAvgInvCTB")) Then
          'GRP_REP.SetFormula mRep, "wTotLabVal", "({rdo.OrSetSalVal}+{rdo.OlSalVal})"   ' ***** Manali 3.9.0
      If UCase(adc("UoYN7")) = "N" Then
          Rep.FldOdDmCd.Suppress = True
      Else
          Rep.FldOdDmCd.Suppress = False
      End If
     End If
     ' '********************************************************************
    
    '*************************** ZUBIN **************************
    ' 09th Dec 2003, EMR206
    If adc("UoYN7") = "Y" Then
      Rep.hBOdDmCd.SetText "Customer Design"
      Rep.Text20.SetText "Customer Design"
      
      Rep1.Text1.SetText "CmDsg/Sz"
      Rep1.TxtDmSz2.SetText "CmDsg/Sz"
      Rep1.TxtDmSz3.SetText "CmDsg/Sz"
    Else
      Rep.hBOdDmCd.SetText "Design"
      Rep.Text20.SetText "Design"
      
      Rep1.Text1.SetText "Design/Sz"
      Rep1.TxtDmSz2.SetText "Design/Sz"
      Rep1.TxtDmSz3.SetText "Design/Sz"
    End If
    
    'If UCase(adc.MenuCd) <> UCase("RepDetOrd") Then
      If adc("UoYN6") = "Y" Then
        Rep.Text45.SetText "Cust Quality"
      Else
        Rep.Text45.SetText "Quality"
      End If
    'End If
    '*************************** ZUBIN **************************
  
    If UCase(adc.MenuCd) = UCase("RepDetInv") Or UCase(adc.MenuCd) = UCase("RepAvgInv") Then    '4.1.2
        Rep.LblWaxWt.Suppress = True: Rep.DmWaxWt.Suppress = True
    End If
  
    ' Manoj 2.10.0
    If adc("UoYN9") = "N" And _
      (UCase(adc.MenuCd) = UCase("RepDetInv") Or UCase(adc.MenuCd) = UCase("RepAvgInv") Or _
      UCase(adc.MenuCd) = UCase("RepDetInvCTB") Or UCase(adc.MenuCd) = UCase("RepAvgInvCTB")) Then
      Rep.SecSumm.Suppress = True
      Rep.SecLabSumm.Suppress = True ' ***** Manali 3.6.0
      
      ' ***** Manali 3.8.0 - CMName
      If GetUsrLevel > 5 Then Rep.PhCmName.Suppress = True
    Else
      Rep.SecSumm.Suppress = False
      Rep.SecLabSumm.Suppress = False   ' ***** Manali 3.6.0
    End If
  ' **** Zubin 211 **** '
  '''''  If ADC("UoYN8") = "Y" Then
    If UCase(adc.MenuCd) <> UCase("RepInvCat") And (adc("UoInDet") = "INVSR/SR" Or adc("UoInDet") = "DMCD/SR" Or adc("UoInDet") = "DMSFX/SR" Or adc("UoInDet") = "TRNO/SR") Then
  ' **** Zubin 211 **** '
     Rep.FldRunSr.Suppress = False
     Rep.FldOdSr.Suppress = True
    Else
     Rep.FldRunSr.Suppress = True
     Rep.FldOdSr.Suppress = False
    End If
    ' Manoj 2.10.0
    
      '***** Manali 3.5.0 - 14/11/08 - Inv Cat fields suppressed/displayed
    If UCase(adc.MenuCd) = UCase("RepInvCat") Or UCase(adc.MenuCd) = UCase("RepInvCTBCat") Then
      ' Bef 3.6.1
      '      If adc("UoYN15") = "Y" Then Rep1.DetOdCol.Suppress = False Else Rep1.DetOdCol.Suppress = True
      ' **** Manali 3.6.1 - Sr in Catalogue
      If adc("UoYN15") = "Y" Or adc("UoYN24") = "Y" Then
          Rep1.DetOdCol.Suppress = False
          ' ***** Manali 3.8.0 - Running Sr in RepInvCat for TrayNo/Sr
          If adc("UoYN24") = "Y" Then
            If (adc("UoInDet") = "INVSR/SR" Or adc("UoInDet") = "DMCD/SR" Or adc("UoInDet") = "DMSFX/SR" Or adc("UoInDet") = "TRNO/SR") Then
              Rep1.FldRunSr.Suppress = False:     Rep1.FldOdSr.Suppress = True
              Rep1.FldRunSr2.Suppress = False:    Rep1.FldOdSr2.Suppress = True     '4.1.2
              Rep1.FldRunSr3.Suppress = False:    Rep1.FldOdSr3.Suppress = True     '4.1.2
            Else
              Rep1.FldRunSr.Suppress = True:     Rep1.FldOdSr.Suppress = False
              Rep1.FldRunSr2.Suppress = True:    Rep1.FldOdSr2.Suppress = False    '4.1.2
              Rep1.FldRunSr3.Suppress = True:    Rep1.FldOdSr3.Suppress = False    '4.1.2
            End If
          End If
          
          ' ***** Manali 3.8.0 - Running Sr in RepInvCat for TrayNo/Sr
          If adc("UoYN15") = "N" And adc("UoYN24") = "Y" Then
            Rep1.lblOdCol.Suppress = True: Rep1.lblColOdCol.Suppress = True: Rep1.fldOdDmCol.Suppress = True
            Rep1.lblOdCol2.Suppress = True: Rep1.lblColOdCol2.Suppress = True: Rep1.fldOdDmCol2.Suppress = True
            Rep1.lblOdCol3.Suppress = True: Rep1.lblColOdCol3.Suppress = True: Rep1.fldOdDmCol3.Suppress = True
            
            Rep1.lblOdSr.Suppress = False: Rep1.lblColOdSr.Suppress = False
            Rep1.lblOdSr2.Suppress = False: Rep1.lblColOdSr2.Suppress = False
            Rep1.lblOdSr3.Suppress = False: Rep1.lblColOdSr3.Suppress = False
            
            Rep1.lblOdSr.Left = 70: Rep1.lblColOdSr.Left = 1090: Rep1.FldOdSr.Left = 1240: Rep1.FldRunSr.Left = 1240
            
          ElseIf Not (adc("UoYN15") = "Y" And adc("UoYN24") = "Y") Then
            Rep1.FldOdSr.Suppress = True: Rep1.lblOdSr.Suppress = True: Rep1.lblColOdSr.Suppress = True
            Rep1.FldOdSr2.Suppress = True: Rep1.lblOdSr2.Suppress = True: Rep1.lblColOdSr2.Suppress = True
            Rep1.FldOdSr3.Suppress = True: Rep1.lblOdSr3.Suppress = True: Rep1.lblColOdSr3.Suppress = True
            
            Rep1.lblOdCol.Suppress = False: Rep1.lblColOdCol.Suppress = False: Rep1.fldOdDmCol.Suppress = False
            Rep1.lblOdCol2.Suppress = False: Rep1.lblColOdCol2.Suppress = False: Rep1.fldOdDmCol2.Suppress = False
            Rep1.lblOdCol3.Suppress = False: Rep1.lblColOdCol3.Suppress = False: Rep1.fldOdDmCol3.Suppress = False
            
            Rep1.FldRunSr.Suppress = True: Rep1.FldRunSr2.Suppress = True: Rep1.FldRunSr3.Suppress = True
          End If
      Else
          Rep1.DetOdCol.Suppress = True
      End If
      ' **** Manali 3.6.1 - Sr in Catalogue
      If adc("UoYN16") = "Y" Then Rep1.DetOdWt.Suppress = False Else Rep1.DetOdWt.Suppress = True
      
      If adc("UoYN17") = "Y" Then
        Rep1.Subreport1_SecRmCd.Suppress = False
        Rep1.Subreport2_SecRmCd.Suppress = False        '4.1.2
        Rep1.Subreport3_SecRmCd.Suppress = False        '4.1.2
      Else
        Rep1.Subreport1_SecRmCd.Suppress = True
        Rep1.Subreport2_SecRmCd.Suppress = True         '4.1.2
        Rep1.Subreport3_SecRmCd.Suppress = True         '4.1.2
      End If
      
      If adc("UoYN18") = "Y" Then
        Rep1.Subreport1_SecRmSCtg.Suppress = False
        Rep1.Subreport2_SecRmSCtg.Suppress = False
        Rep1.Subreport3_SecRmSCtg.Suppress = False
      Else
        Rep1.Subreport1_SecRmSCtg.Suppress = True
        Rep1.Subreport2_SecRmSCtg.Suppress = True
        Rep1.Subreport3_SecRmSCtg.Suppress = True
      End If
      
      If adc("UoYN19") = "Y" Then
        Rep1.Subreport1_SecRmWt.Suppress = False
        Rep1.Subreport2_SecRmWt.Suppress = False
        Rep1.Subreport3_SecRmWt.Suppress = False
      Else
        Rep1.Subreport1_SecRmWt.Suppress = True
        Rep1.Subreport2_SecRmWt.Suppress = True
        Rep1.Subreport3_SecRmWt.Suppress = True
      End If
        
      If adc("UoYN20") = "Y" Then
        Rep1.Subreport1_SecRmQty.Suppress = False
        Rep1.Subreport2_SecRmQty.Suppress = False
        Rep1.Subreport3_SecRmQty.Suppress = False
      Else
        Rep1.Subreport1_SecRmQty.Suppress = True
        Rep1.Subreport2_SecRmQty.Suppress = True
        Rep1.Subreport3_SecRmQty.Suppress = True
      End If
      ' ***** Manali 3.10.0 - 09/03/12 - Different Suppress options for Unit Pricce and Qty
      
      Rep1.TxtUnitPrcCurCol.Suppress = True
      Rep1.TxtUnitPrcCurCol2.Suppress = True
      Rep1.TxtUnitPrcCurCol3.Suppress = True
      
      If adc("UoYN21") = "Y" Then
        Rep1.TxtUnitPrcCur.Suppress = False: Rep1.FldOdSalPrc.Suppress = False: Rep1.TxtUnitPrcCurCol.Suppress = False
        Rep1.TxtUnitPrcCur2.Suppress = False: Rep1.FldOdSalPrc2.Suppress = False: Rep1.TxtUnitPrcCurCol2.Suppress = False
        Rep1.TxtUnitPrcCur3.Suppress = False: Rep1.FldOdSalPrc3.Suppress = False: Rep1.TxtUnitPrcCurCol3.Suppress = False
      Else
        Rep1.TxtUnitPrcCur.Suppress = True:  Rep1.FldOdSalPrc.Suppress = True:   'Rep1.TxtUnitPrcCurCol.Suppress = True
        Rep1.TxtUnitPrcCur2.Suppress = True:  Rep1.FldOdSalPrc2.Suppress = True:   'Rep1.TxtUnitPrcCurCol2.Suppress = True
        Rep1.TxtUnitPrcCur3.Suppress = True:  Rep1.FldOdSalPrc3.Suppress = True:   'Rep1.TxtUnitPrcCurCol3.Suppress = True
      End If
      
      If adc("UoYN27") = "Y" Then
        Rep1.TxtQty.Suppress = False: Rep1.FldqOrdQty.Suppress = False
        Rep1.TxtQty2.Suppress = False: Rep1.FldqOrdQty2.Suppress = False
        Rep1.TxtQty3.Suppress = False: Rep1.FldqOrdQty3.Suppress = False
        
        If adc("UoYN21") = "N" Then
            Rep1.TxtQty.Left = Rep1.TxtUnitPrcCur.Left: Rep1.FldqOrdQty.Left = Rep1.FldOdSalPrc.Left
            Rep1.TxtQty2.Left = Rep1.TxtUnitPrcCur2.Left: Rep1.FldqOrdQty2.Left = Rep1.FldOdSalPrc2.Left
            Rep1.TxtQty3.Left = Rep1.TxtUnitPrcCur3.Left: Rep1.FldqOrdQty3.Left = Rep1.FldOdSalPrc3.Left
        End If
      Else
        Rep1.TxtQty.Suppress = True: Rep1.FldqOrdQty.Suppress = True
        Rep1.TxtQty2.Suppress = True: Rep1.FldqOrdQty2.Suppress = True
        Rep1.TxtQty3.Suppress = True: Rep1.FldqOrdQty3.Suppress = True
      End If
      
      'If adc("UoYN21") = "N" And adc("UoYN27") = "N" Then Rep1.DetOdPrc.Suppress = True
      ' ***** Manali 3.10.0 - 09/03/12 - Different Suppress options for Unit Pricce and Qty
      ' Bef 3.10.0
      'If adc("UoYN21") = "Y" Then Rep1.DetOdPrc.Suppress = False Else Rep1.DetOdPrc.Suppress = True
      
      If adc("UoYN22") = "Y" Then
        Rep1.DetodPoNo.Suppress = False
      Else
        Rep1.DetodPoNo.Suppress = True
      End If
      
      If adc("UoYN17") = "N" And adc("UoYN18") = "N" And adc("UoYN18") = "N" And adc("UoYN20") = "N" Then _
            Rep1.DetCatSubRep.Suppress = True Else Rep1.DetCatSubRep.Suppress = False
    End If
    CmdDispSc.Top = Me.ScaleHeight - CmdDispSc.Height
    '***** Manali 3.5.0 - 14/11/08 - Inv Cat fields suppressed/displayed
    
    '***** Manali 3.4.1 - 12/11/08 - OdOmPoNo to be displayed only in Invoice Printing
    Rep.LblOdPoNo.SetText "PO No ": Rep.TxtOdPoNo.Suppress = True: Rep.TxtOdomPoNo.Suppress = False
    ' ***** Manali 3.4.1 - 12/11/08 - OdOmPoNo to be displayed only in Invoice Printing
    
    'Rep1.TxtUnitPrcCur.Suppress
    'pg.5 -MRP detail will not display for invoice reports
    Rep.TxtMrp.Suppress = True: Rep.OdMrp1.Suppress = True: Rep.OdMrpDisc1.Suppress = True
    Rep.LblExpOdNoMrp.Suppress = True: Rep.TxtExpOdNoMrp.Suppress = True
    
    'BnM.3- if SHow Bag/StkId option is selected then will display Bag /stk id detials
    If adc("UoYN10") = "Y" Then
      'BnM.3- stock id details will display for  Detail and Average costing reports
      Rep.wStkId1.Suppress = True: Rep.wBagNo1.Suppress = False
      Rep.wStkId2.Suppress = True: Rep.wBagNo2.Suppress = False
    Else
      Rep.BagStkIdSec.Suppress = True
      Rep.BagStkIdSec1.Suppress = True
    End If
    
    'sv.24 StnCls summary visible only for inv detail report
    If UCase(adc.MenuCd) = UCase("RepDetInv") Then Rep.Subreport3.Suppress = False
    '62 HuidNo will come oly for detail invoice printing
    If UCase(adc.MenuCd) <> UCase("RepDetInv") Then
      Rep.qHuidNo1.Suppress = True: Rep.TxtHUIDNo.Suppress = True
    End If
    '578.19 calling function to change Rmcode font size
    GRP_REP.SetFormula Rep, "wRmCdFontSz", GetRmcdFontSz("{rdo.OrRmCd}", 7.5)
    GRP_REP.SetFormula Rep, "wRmCdFontSzCT", GetRmcdFontSz("{rdo.OctRmCd}", 7.5)
    
  CRV_REP.DisplayGroupTree = False
      
  '*** (Jen 2.12)
  
  
  
Call DispReport(adc) '****** Sachin 3.02 - Display Report ******
  
  Exit Sub
RepErr:
  Call ChkSelScope(Err.Number, Err.Description)
  '6.1_27 if copy pic option is selected and pic path is not valid it will show this error message
  If adc("UOYN31") = "Y" And adc("wTarPicPath") <> "" And Err.Number = 53 Then _
    DispMsg "Picture Path is Invalid. Enter Proper Picture Path ", etError
  '*** (Jen 2.12)
End Sub
Private Sub ChkBoxArr_Click(Index As Integer)
With ChkBoxArr(Index)
  Select Case Index
  Case Is = 0           'Tag = "ShowPic"
    If .Value = Checked Then adc("UoYN") = "Y" Else adc("UoYN") = "N"
  Case Is = 1           'Tag = "ShowAvgWt"
    If .Value = Checked Then adc("UoYN1") = "Y" Else adc("UoYN1") = "N"
  Case Is = 2           'Tag = "ShowDsgVal"
    If .Value = Checked Then adc("UoYN2") = "Y" Else adc("UoYN2") = "N"
  Case Is = 3          'Tag = "ShowDistDsg"
    ' **** Manali 3.6.1 - Sr in Catalogue - ChkBoxArr(24).Value
    If .Value = Checked Then adc("UoYN3") = "Y": ChkBoxArr(24).Value = Unchecked: adc("UoYN24") = "N" Else adc("UoYN3") = "N"
    'If .Value = Checked Then adc("UoYN3") = "Y" Else adc("UoYN3") = "N"
  Case Is = 4           'Tag = "ShowLabVal"
    If .Value = Checked Then adc("UoYN4") = "Y" Else adc("UoYN4") = "N"
  Case Is = 5         'Tag = "SortCust"
    If .Value = Checked Then adc("UoYN5") = "Y" Else adc("UoYN5") = "N"
  Case Is = 6          'Tag = "ShowCustRm"
    If .Value = Checked Then adc("UoYN6") = "Y" Else adc("UoYN6") = "N"
  Case Is = 7          'Tag = "ShowCustDsg"
    If .Value = Checked Then adc("UoYN7") = "Y" Else adc("UoYN7") = "N"
  Case Is = 8           'Tag = "ShowRnSr"
    If .Value = Checked Then adc("UoYN8") = "Y" Else adc("UoYN8") = "N"
  Case Is = 9           'Tag = "ShowSumm"
    If .Value = Checked Then adc("UoYN9") = "Y" Else adc("UoYN9") = "N"
  'BnM.3 UoYN10 will store YN value as per check box Show Bag/StkId is selected
  Case Is = 10           'Tag = "ShowBag/StkId"
    If .Value = Checked Then adc("UoYN10") = "Y" Else adc("UoYN10") = "N"
  Case Is = 11        'Tag = "FinalSrt"
    If .Value = Checked Then adc("UoYN11") = "Y" Else adc("UoYN11") = "N"
  Case Is = 13         'Tag = "ShowGrpDesc"
    If .Value = Checked Then adc("UoYN13") = "Y" Else adc("UoYN13") = "N"
  Case Is = 14          'Tag = "ShowRmVal"
    If .Value = Checked Then adc("UoYN14") = "Y" Else adc("UoYN14") = "N"
  Case Is = 15          'Tag = "ShowDsgCol"
    If .Value = Checked Then adc("UoYN15") = "Y" Else adc("UoYN15") = "N"
  Case Is = 16         'Tag = "ShowGrsWt"
    If .Value = Checked Then adc("UoYN16") = "Y" Else adc("UoYN16") = "N"
  Case Is = 17          'Tag = "ShowRmCd"
    If .Value = Checked Then adc("UoYN17") = "Y" Else adc("UoYN17") = "N"
  Case Is = 18          'Tag = "ShowRmSCtg"
    If .Value = Checked Then adc("UoYN18") = "Y" Else adc("UoYN18") = "N"
  Case Is = 19          'Tag = "ShowRmWt"
    If .Value = Checked Then adc("UoYN19") = "Y" Else adc("UoYN19") = "N"
  Case Is = 20         'Tag = "ShowRmQty"
    If .Value = Checked Then adc("UoYN20") = "Y" Else adc("UoYN20") = "N"
  Case Is = 21         'Tag = "ShowUnitPrc"
    If .Value = Checked Then adc("UoYN21") = "Y" Else adc("UoYN21") = "N"
  Case Is = 22          'Tag = "ShowOdPO"
    If .Value = Checked Then adc("UoYN22") = "Y" Else adc("UoYN22") = "N"
  Case Is = 23          'Tag = "ShowExcel"
    If .Value = Checked Then adc("UoYN23") = "Y" Else adc("UoYN23") = "N"
  ' **** Manali 3.6.1 - Sr in Catalogue
  Case Is = 24          'Tag = "ShowVchSr"
    If .Value = Checked Then adc("UoYN24") = "Y": ChkBoxArr(3).Value = Unchecked: adc("UoYN3") = "N" Else adc("UoYN24") = "N"
  ' **** Manali 3.6.1 - Sr in Catalogue
  Case Is = 26          'Tag = "ShowVchSr"
    If .Value = Checked Then adc("UoYN26") = "Y" Else adc("UoYN26") = "N"
  Case Is = 27          'Tag = "ShowDsgQty"   ' ***** Manali 3.10.0 - 09/03/12 - Different Suppress options for Unit Pricce and Qty
    If .Value = Checked Then adc("UoYN27") = "Y" Else adc("UoYN27") = "N"
  Case Is = 29          'Tag = "SortOrdCust"   ' 4.1.2
    If .Value = Checked Then adc("UoYN29") = "Y" Else adc("UoYN29") = "N"
  Case Is = 30          'Tag = "SortOrdCust"   ' 4.1.2
    If .Value = Checked Then adc("UoYN30") = "Y" Else adc("UoYN30") = "N"
  Case Is = 31          '6.1_27
    If .Value = Checked Then adc("UoYN31") = "Y" Else adc("UoYN31") = "N"
  Case Is = 32          '6.1b-15
    If .Value = Checked Then adc("UoYN32") = "Y" Else adc("UoYN32") = "N"
  End Select
End With
End Sub

Private Sub ChkBoxArr_GotFocus(Index As Integer)
DispMsg ChkBoxArr(Index).ToolTipText, etInfo
End Sub

Private Sub ChkBoxLme_Click(Index As Integer)
' ***** Manali 3.9.0 - Lme Selection
With ChkBoxLme(Index)
  If .Value = Checked Then wChkSel = wChkSel + 1 Else wChkSel = wChkSel - 1
  If wChkSel > 2 Then DispMsg "Cannot Select More Than 2 Lmes", etError: .Value = Unchecked: Exit Sub

  Dim wVal As String, i As Integer
  wVal = .Caption

  If .Value = Checked Then
    If adc("UoLme1") <> wVal And adc("UoLme2") <> wVal Then
'      If adc("UoLme1") = "" Then
'        adc("UoLme1") = wVal
'      Else
'        adc("UoLme2") = wVal
'      End If
      
      If adc("UoLme1") <> "" Then
        For i = Index + 1 To 3
          If adc("UoLme1") = ChkBoxLme(i).Caption Then
            adc("UoLme2") = adc("UoLme1")
            adc("UoLme1") = wVal
            Exit For
          End If
        Next i
      End If
      If adc("UoLme1") = "" Then
        adc("UoLme1") = wVal
      ElseIf adc("UoLme2") = "" Then
        adc("UoLme2") = wVal
      End If
    End If
  Else
    If adc("UoLme1") = wVal Then
      adc("UoLme1") = adc("UoLme2")
      adc("UoLme2") = ""
    ElseIf adc("UoLme2") = wVal Then
      adc("UoLme2") = ""
    End If
  End If
End With
End Sub
Private Sub ChkBoxLme_GotFocus(Index As Integer)
DispMsg FraLme.ToolTipText, etInfo
End Sub

Private Sub CmdChgImg_Click()
If FraChgImg.Visible = False Then
    If CRV_REP.Visible = True Then Exit Sub
    FraChgImg.Enabled = True
    FraChgImg.Visible = True
    TAB_REP.Enabled = False
    adc.AllowSave = False
    FraChgImg.ZOrder
  Else
    TAB_REP.Enabled = True
    adc.AllowSave = True
    FraChgImg.Enabled = False
    FraChgImg.Visible = False
  End If
End Sub

Private Sub CmdChgImgGo_Click()
On Error GoTo ErrHd
Dim wPicHt As Single, wPicWidth As Single
If UCase(adc.MenuCd) = UCase("RepInvCat") And UCase(adc.MenuCd) = UCase("RepInvCTBCat") Then
    wPicHt = 110: wPicWidth = 110
Else
    wPicHt = 84: wPicWidth = 84
End If
If adc("wXlFilePath") = "" Then
    DispMsg "Please Provide an Excel File to change Images", etError
    Exit Sub
End If
If ReplXlImg(adc("wXlFilePath"), wo_rsPic, wPicHt, wPicWidth) Then
    TAB_REP.Enabled = True
    adc.AllowSave = True
    FraChgImg.Visible = False
    CmdChgImg.Visible = False
    DispMsg "Images Replaced Successfully", etInfo
End If
Exit Sub
ErrHd:
    DispMsg Err.Description, etError
End Sub

Private Sub CmdSelFile_Click()
'Sachin - 3.11.2
    On Error GoTo ErrHd
    Dim wExtn As String
    wExtn = ".xls;*.xlsx"
    With Cd1
      .FileName = ""
      .DialogTitle = "Open Report Excel"
      .Filter = "Excel Files (*" + wExtn + ") |*" + wExtn
      .CancelError = False
      .ShowOpen
      If .FileName <> "" Then adc("WXLFILEPATH") = .FileName
    End With
    Exit Sub
ErrHd:
    DispMsg Err.Description, etError
End Sub



' ***** Manali 3.10.0 - 14/03/12 - Show Net Wt instead of Gld As Wt
Private Sub OptNetWt_Click(Index As Integer)
Select Case Index
Case Is = 0
  adc("UoYN28") = "Y"
Case Is = 1
  adc("UoYN28") = "N"
End Select
End Sub
' ***** Manali 3.10.0 - 14/03/12 - Show Net Wt instead of Gld As Wt
Private Sub OptNetWt_GotFocus(Index As Integer)
  DispMsg FraOptNetWt.ToolTipText, etInfo
End Sub

Private Sub CmdDispSc_Click()
' **** Manali 3.5.0 - 14/11/08 - Supress or Display fields in Ord Cat
  If FraDispSc.Visible = False Then
    If CRV_REP.Visible = True Then Exit Sub
    FraDispSc.Enabled = True
    FraDispSc.Visible = True
    TAB_REP.Enabled = False
    adc.AllowSave = False
    FraDispSc.ZOrder
  Else
    TAB_REP.Enabled = True
    adc.AllowSave = True
    FraDispSc.Enabled = False
    FraDispSc.Visible = False
    CmdDispSc.SetFocus
  End If
' **** Manali 3.5.0 - 14/11/08 - Supress or Display fields in Ord Cat
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
    gb_CoCdFor = True  '****** Sachin 2.14.0
    
    Call InitProp(Me)
    
   '****************************Geeta**********Emr206********************
   ' Zubin 212 (RepInvCTBCat added)
   If UCase(adc.MenuCd) = UCase("RepDetInvCTB") Or UCase(adc.MenuCd) = UCase("RepAvgInvCTB") Or UCase(adc.MenuCd) = UCase("RepInvCTBCat") Then _
        adc("UOINGRP").Visible = False: adc("UOINGRP1").Visible = False: LblGrpCombo.Visible = False: LblGrpCombo1.Visible = False

  '************************Geeta***********************************************************
    If UCase(adc.MenuCd) = UCase("RepDetInv") Or UCase(adc.MenuCd) = UCase("RepDetInvCTB") Then
      gs_RmGrp = "ORDDET"
      gs_LabGrp = "ORDDET"
      Me.Caption = "Detail Invoice Printing"
    '  LblCustCd.Enabled = False: LblCustCd.Visible = False
      adc("UoYN5").Visible = False: adc("UoYN5").Visible = False
      LblSortCust.Visible = False: LblSortCust.Enabled = False
      ChkBoxArr(5).Enabled = False: ChkBoxArr(5).Visible = False    ' ****** Manali 3.5.0 - 19/11/08 - Check Boxes added
     ' adc("UoCmCdFr").Enabled = False: adc("UoCmCdFr").Visible = False
      'adc("UoCmCdTo").Enabled = False: adc("UoCmCdTo").Visible = False
      CmdDispSc.Enabled = False: CmdDispSc.Visible = False    '***** Manali 3.5.0 - 14/11/08 - Ord Cat fields suppressed/displayed
      If UCase(adc.MenuCd) = UCase("RepDetInv") Then gs_InTcTyp = "IN": Me.Caption = "Detail Invoice Printing"
      If UCase(adc.MenuCd) = UCase("RepDetInvCTB") Then gs_InTcTyp = "CTB": Me.Caption = "Detail CTB Printing"
      ChkBoxArr(24).Enabled = False: ChkBoxArr(24).Visible = False    ' **** Manali 3.6.1 - Sr in Catalogue
    ElseIf UCase(adc.MenuCd) = UCase("RepAvgInv") Or UCase(adc.MenuCd) = UCase("RepAvgInvCTB") Then
      gs_RmGrp = "ORDAVG"
      gs_LabGrp = "ORDAVG"
      Me.Caption = "Average Invoice"
      adc("UoRmGrp").Visible = False: adc("UoLabGrp").Visible = False
      adc("UoYN5").Visible = False: adc("UoYN5").Visible = False
      LblSortCust.Visible = False: LblSortCust.Enabled = False
      ChkBoxArr(5).Enabled = False: ChkBoxArr(5).Visible = False    ' ****** Manali 3.5.0 - 19/11/08 - Check Boxes added
      adc("UoRmGrp").Enabled = False: adc("UoLabGrp").Enabled = False
      LblRmGrp.Visible = False: LblLabGrp.Visible = False
      'LblCustCd.Enabled = False: LblCustCd.Visible = False
      'adc("UoCmCdFr").Enabled = False: adc("UoCmCdFr").Visible = False
      'adc("UoCmCdTo").Enabled = False: adc("UoCmCdTo").Visible = False
      CmdDispSc.Enabled = False: CmdDispSc.Visible = False    '***** Manali 3.5.0 - 14/11/08 - Ord Cat fields suppressed/displayed
      If UCase(adc.MenuCd) = UCase("RepAvgInv") Then gs_InTcTyp = "IN": Me.Caption = "Average Invoice"
      If UCase(adc.MenuCd) = UCase("RepAvgInvCTB") Then gs_InTcTyp = "CTB": Me.Caption = "Average CTB"
      
      ChkBoxArr(23).Enabled = False: ChkBoxArr(23).Visible = False    ' ***** Manali 3.6.0 - 24/09/09 - Show Report in Excel Format
      ChkBoxArr(24).Enabled = False: ChkBoxArr(24).Visible = False    ' **** Manali 3.6.1 - Sr in Catalogue
            
    ElseIf UCase(adc.MenuCd) = UCase("RepInvCat") Or UCase(adc.MenuCd) = UCase("RepInvCTBCat") Then
      'adc("UoYN5").Visible = True: adc("UoYN5").Visible = True   ' ****** Manali 3.5.0 - 19/11/08 - Check Boxes added
      'LblSortCust.Visible = True: LblSortCust.Enabled = True     ' ****** Manali 3.5.0 - 19/11/08 - Check Boxes added
      ChkBoxArr(5).Enabled = True: ChkBoxArr(5).Visible = True    ' ****** Manali 3.5.0 - 19/11/08 - Check Boxes added
      adc("UoRmGrp").Visible = False: adc("UoLabGrp").Visible = False
      adc("UoRmGrp").Enabled = False: adc("UoLabGrp").Enabled = False
      adc("UoYN").Visible = False: adc("UoYN").Visible = False
      ChkBoxArr(0).Enabled = False: ChkBoxArr(0).Visible = False    ' ****** Manali 3.5.0 - 19/11/08 - Check Boxes added
      adc("UoYN1").Visible = False: adc("UoYN1").Visible = False
      ChkBoxArr(1).Enabled = False: ChkBoxArr(1).Visible = False    ' ****** Manali 3.5.0 - 19/11/08 - Check Boxes added
      LblRmGrp.Visible = False: LblLabGrp.Visible = False
      LblPic.Visible = False: LblAvgWt.Visible = False
      adc("UoYN4").Visible = False: adc("UoYN4").Enabled = False: LblLabVal.Visible = False
      ChkBoxArr(4).Enabled = False: ChkBoxArr(4).Visible = False    ' ****** Manali 3.5.0 - 19/11/08 - Check Boxes added
      adc("UoYN2").Visible = False: adc("UoYN2").Enabled = False: LblShwVal.Visible = False
      ChkBoxArr(2).Enabled = False: ChkBoxArr(2).Visible = False    ' ****** Manali 3.5.0 - 19/11/08 - Check Boxes added
      '*** Jay 2.14Next [Show Rm Value]
      adc("UoYN14").Visible = False: adc("UoYN14").Enabled = False: LblRmVal.Visible = False
      ChkBoxArr(14).Enabled = False: ChkBoxArr(14).Visible = False    ' ****** Manali 3.5.0 - 19/11/08 - Check Boxes added
      '*** Jay 2.14Next [Show Rm Value]
      ' Manoj 2.10.0
      adc("UoYN9").Visible = False: adc("UoYN9").Enabled = False: LblShowSumm.Visible = False
      ChkBoxArr(9).Enabled = False: ChkBoxArr(9).Visible = False    ' ****** Manali 3.5.0 - 19/11/08 - Check Boxes added
      adc("UoYN8").Visible = False: adc("UoYN8").Enabled = False: LblShowRunSr.Visible = False
      ChkBoxArr(8).Enabled = False: ChkBoxArr(8).Visible = False    ' ****** Manali 3.5.0 - 19/11/08 - Check Boxes added
      ' Manoj 2.10.0
      '*** JAY 2.14 ***
      adc("UOYN13").Visible = False: LblShowGrpDesc.Visible = False
      ChkBoxArr(13).Enabled = False: ChkBoxArr(13).Visible = False    ' ****** Manali 3.5.0 - 19/11/08 - Check Boxes added
      '*** JAY 2.14 ***
      If UCase(adc.MenuCd) = UCase("RepInvCat") Then gs_InTcTyp = "IN": Me.Caption = "Invoice Catalogue"
      If UCase(adc.MenuCd) = UCase("RepInvCTBCat") Then gs_InTcTyp = "CTB": Me.Caption = "CTB Catalogue"
      
      ChkBoxArr(23).Enabled = False: ChkBoxArr(23).Visible = False    ' ***** Manali 3.6.0 - 24/09/09 - Show Report in Excel Format
      
      FraLme.Enabled = False: FraLme.Visible = False      ' ***** Manali 3.9.0 - Lme Selection
      LblNetWt.Visible = False: FraOptNetWt.Enabled = False: FraOptNetWt.Visible = False    ' ***** Manali 3.10.0 - 14/03/12 - Show Net Wt instead of Gld As Wt
      
      ChkBoxArr(30).Visible = True      '4.1.2
      ChkBoxArr(26).Enabled = False: ChkBoxArr(26).Visible = False
    End If

' **** Zubin 211 **** '
' Zubin 212 (RepDetInvCTB, RepAvgInvCTB and RepInvCTBCat added)
    If UCase(adc.MenuCd) = UCase("RepDetInv") Or UCase(adc.MenuCd) = UCase("RepAvgInv") _
      Or UCase(adc.MenuCd) = UCase("RepInvCat") Or UCase(adc.MenuCd) = UCase("RepDetInvCTB") _
      Or UCase(adc.MenuCd) = UCase("RepAvgInvCTB") Or UCase(adc.MenuCd) = UCase("RepInvCTBCat") Then
        LblFinalSrt.Visible = False: adc("UOYN11").Visible = False
        LblShowRunSr.Visible = False: adc("UOYN8").Visible = False
        ChkBoxArr(11).Enabled = False: ChkBoxArr(11).Visible = False    ' ****** Manali 3.5.0 - 19/11/08 - Check Boxes added
        ChkBoxArr(8).Enabled = False: ChkBoxArr(8).Visible = False    ' ****** Manali 3.5.0 - 19/11/08 - Check Boxes added
    End If
' **** Zubin 211 **** '
    
    
    If UCase(adc.MenuCd) = UCase("RepDetInvCTB") Or UCase(adc.MenuCd) = UCase("RepAvgInvCTB") Or _
      UCase(adc.MenuCd) = UCase("RepInvCTBCat") Then _
        LblDet.Visible = False: adc("UOINDET").Visible = False
' **** Zubin 211 **** '

    ' ***** Manali 3.6.0 - Order Printing in Excel Format Module
    If moCn.RecSeek("Select 'x' from Head Where HCoCd='" + ctSelfCoCd + "' and HCd='" + ctSelfCmCd + "' and HOrdExcelRepYn<>'Y' ") Then _
      ChkBoxArr(23).Visible = False: ChkBoxArr(23).Enabled = False
    'BnM.3 -Show Bag/StkId option in scope will show for Detail and Average invoice printing
    If UCase(adc.MenuCd) = UCase("RepDetInv") Or UCase(adc.MenuCd) = UCase("RepAvgInv") Then
      ChkBoxArr(10).Visible = True: ChkBoxArr(10).Enabled = True
    Else
      ChkBoxArr(10).Visible = False: ChkBoxArr(10).Enabled = False
    End If
    '6.1_27 copy picture option and pic path will not visible except detail invoice and catalogue report
    If UCase(adc.MenuCd) <> UCase("RepDetInv") And UCase(adc.MenuCd) <> UCase("RepInvCat") Then
      adc("wTarPicPath").Visible = False: LblPicPath.Visible = False
      ChkBoxArr(31).Enabled = False: ChkBoxArr(31).Visible = False
    End If
    If UCase(adc.MenuCd) = UCase("RepInvCat") Then ChkBoxArr(32).Visible = True   '6.1B-15 this option will be visible only for invoice catalog
    
    Call SetGroupSort
End Sub


Private Sub ATXT_FldChange(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String, ByVal pv_NewValue As Variant)
  If UCase(adc.MenuCd) = UCase("RepDetInv") Or UCase(adc.MenuCd) = UCase("RepAvgInv") Or _
     UCase(adc.MenuCd) = UCase("RepInvCat") Then
    gs_InTcTyp = "IN"
  Else
    gs_InTcTyp = "CTB"
  End If
  
  Select Case UCase(IdName)
  Case Is = UCase("UoYN2")
      Cancel = pv_NewValue <> "" And (Not moCn.RecSeek("select PMCd from Param where " + _
                           " PTyp= 'YN' and PMCd = '" + pv_NewValue + "'"))
      ErrMsg = "Invalid Show Value Option": Exit Sub
    

  Case Is = UCase("UoYN4"), UCase("UoYN5")
      Cancel = pv_NewValue <> "" And (Not moCn.RecSeek("select PMCd from Param where " + _
                           " PTyp= 'YN' and PMCd = '" + pv_NewValue + "'"))
      ErrMsg = "Invalid Labour Value Option": Exit Sub
  Case Is = UCase("UoYN5"), UCase("UoYN3")
      Cancel = Not moCn.RecSeek("select PMCd from Param where " + _
                           " PTyp= 'YN' and PMCd = '" + pv_NewValue + "'")
      ErrMsg = "Invalid Sort Option": Exit Sub
    
  End Select

  Call SetProp(Me, IdName, Valid, Cancel, ErrMsg)
End Sub
Private Sub ATXT_FldWhen(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String)
  gb_CoCdFor = True         '*** (Jen 2.14 Next Patch 1)
  
  If UCase(adc.MenuCd) = UCase("RepDetInv") Or UCase(adc.MenuCd) = UCase("RepAvgInv") Or _
     UCase(adc.MenuCd) = UCase("RepInvCat") Then
    gs_InTcTyp = "IN"
  Else
    gs_InTcTyp = "CTB"
  End If
  
  ' Zubin 213
  Select Case UCase(IdName)
  Case Is = UCase("UoCurCdFr")
    If UCase(adc("UoInVal")) = "S" Then Cancel = True: Exit Sub
  End Select
  ' Zubin 213

  SetProp Me, IdName, When
End Sub
Private Sub atxt_ReCalculate(Index As Integer, ByVal IdName As String, ByVal OldVal As Variant, ByVal RowNum As Integer, ByVal SourceName As String, ByVal SourceColName As String)
 Select Case UCase(IdName)
 ' Zubin 212 (Below UoYN4 Recalc not req)
'  Case Is = UCase("UoYN4")
'    If adc("UoYN2") = "Y" Then
'      adc("UoYN4") = ""
'    End If
  ' Zubin 213
  Case Is = UCase("UoCurCdFr")
    If UCase(adc("UoInVal")) = "S" Then adc("UoCurCdFr") = "": Exit Sub
  End Select
End Sub

Private Sub tab_REP_Click(PreviousTab As Integer)
    fra_tabrep(TAB_REP.Tab).Enabled = True
    fra_tabrep(PreviousTab).Enabled = False
End Sub
Private Sub ADC_KeyWhen()
    adc("UOUSRCD") = adc.UsrCd
    adc("UOMNUCD") = adc.MenuCd
    '*************************** ZUBIN **************************
  ' 08th Dec 2003, EMR206
  If UCase(adc.MenuCd) = UCase("RepAvgInv") Or UCase(adc.MenuCd) = UCase("RepInvCat") _
        Or UCase(adc.MenuCd) = UCase("RepAvgInvCTB") Or UCase(adc.MenuCd) = UCase("RepInvCTBCat") Then
    lblCustRm.Visible = False
    adc("UoYN6").Visible = False
    ChkBoxArr(6).Enabled = False: ChkBoxArr(6).Visible = False    ' ****** Manali 3.5.0 - 19/11/08 - Check Boxes added
  End If
  If UCase(adc.MenuCd) <> UCase("RepInvCat") Then
      adc("UOYN3").Visible = False
      LblDistDsg.Visible = False
      ChkBoxArr(3).Enabled = False: ChkBoxArr(3).Visible = False    ' ****** Manali 3.5.0 - 19/11/08 - Check Boxes added
  End If

  '*************************** ZUBIN **************************
End Sub
Private Sub ADC_NonKeyWhen(pr_Cancel As Boolean, pr_ErrMsg As String)
 adc("UoCoCdFr") = gs_CoCd
 
' **** Zubin 211 **** '
    If (UCase(adc.MenuCd) = UCase("RepDetInv") Or UCase(adc.MenuCd) = UCase("RepAvgInv") Or _
      UCase(adc.MenuCd) = UCase("RepInvCat")) And adc("UoInDet") = "" Then _
        adc("UoInDet") = "INVSR"
 ' **** Zubin 211 **** '
  ' Zubin 212
  If adc("UOYN4") = "" And UCase(adc.MenuCd) <> UCase("RepInvCat") And UCase(adc.MenuCd) <> UCase("RepInvCTBCat") Then _
    adc("UOYN4") = "Y"
  ' Zubin 212
  
  ' ***** Manali 3.10.0 - 14/03/12 - Show Net Wt instead of Gld As Wt
  If adc("UoYN28") = "Y" Then OptNetWt(0).Value = True
  If adc("UoYN28") = "N" Then OptNetWt(1).Value = True
  ' ***** Manali 3.10.0 - 14/03/12 - Show Net Wt instead of Gld As Wt
  
  ' ****** Manali 3.5.0 - 19/11/08 - Check Boxes added
  If adc("UoYN") = "" Then adc("UoYN") = "N"
  If adc("UoYN1") = "" Then adc("UoYN1") = "N"
  If adc("UoYN2") = "" Then adc("UoYN2") = "N"
  If adc("UoYN3") = "" Then adc("UoYN3") = "N"
  If adc("UoYN4") = "" Then adc("UoYN4") = "N"
  If adc("UoYN5") = "" Then adc("UoYN5") = "N"
  If adc("UoYN6") = "" Then adc("UoYN6") = "N"
  If adc("UoYN7") = "" Then adc("UoYN7") = "N"
  If adc("UoYN8") = "" Then adc("UoYN8") = "N"
  If adc("UoYN9") = "" Then adc("UoYN9") = "N"
  If adc("UoYN11") = "" Then adc("UoYN11") = "N"
  If adc("UoYN13") = "" Then adc("UoYN13") = "N"
  If adc("UoYN14") = "" Then adc("UoYN14") = "N"
  If adc("UoYN15") = "" Then adc("UoYN15") = "N"
  If adc("UoYN16") = "" Then adc("UoYN16") = "N"
  If adc("UoYN17") = "" Then adc("UoYN17") = "N"
  If adc("UoYN18") = "" Then adc("UoYN18") = "N"
  If adc("UoYN19") = "" Then adc("UoYN19") = "N"
  If adc("UoYN20") = "" Then adc("UoYN20") = "N"
  If adc("UoYN21") = "" Then adc("UoYN21") = "N"
  If adc("UoYN22") = "" Then adc("UoYN22") = "N"
  ' ***** Manali 3.6.0 - Order Printing in Excel Format Module
  'If ADC("UoYN23") = "" Or UCase(ADC.MenuCd) <> UCase("RepInvDet") Then ADC("UoYN23") = "N"   ' ***** Manali 3.6.0 - 24/09/09 - Show Report in Excel Format
  If moCn.RecSeek("Select 'x' from Head Where HCoCd='" + ctSelfCoCd + "' and HCd='" + ctSelfCmCd + "' and HOrdExcelRepYn<>'Y' ") Or _
    adc("UoYN23") = "" Or UCase(adc.MenuCd) <> UCase("RepInvDet") Then adc("UoYN23") = "N"
  ' ***** Manali 3.6.0 - Order Printing in Excel Format Module
  If adc("UoYN24") = "" Then adc("UoYN24") = "N"
  If adc("UoYN27") = "" Then adc("UoYN27") = "N"  ' ***** Manali 3.10.0 - 09/03/12 - Different Suppress options for Unit Pricce and Qty
  If adc("UoYN29") = "" Then adc("UoYN29") = "N"  ' 4.1.2
  If adc("UoYN30") = "" Then adc("UoYN30") = "N"  ' 4.1.2
  If adc("UoYN31") = "" Then adc("UoYN31") = "N"  '6.1_27
  If adc("UoYN32") = "" Then adc("UoYN32") = "N"  '6.1b-15
  
  If adc("UoYN") = "Y" Then ChkBoxArr(0).Value = Checked Else ChkBoxArr(0).Value = Unchecked
  If adc("UoYN1") = "Y" Then ChkBoxArr(1).Value = Checked Else ChkBoxArr(1).Value = Unchecked
  If adc("UoYN2") = "Y" Then ChkBoxArr(2).Value = Checked Else ChkBoxArr(2).Value = Unchecked
  If adc("UoYN3") = "Y" Then ChkBoxArr(3).Value = Checked Else ChkBoxArr(3).Value = Unchecked
  If adc("UoYN4") = "Y" Then ChkBoxArr(4).Value = Checked Else ChkBoxArr(4).Value = Unchecked
  If adc("UoYN5") = "Y" Then ChkBoxArr(5).Value = Checked Else ChkBoxArr(5).Value = Unchecked
  If adc("UoYN6") = "Y" Then ChkBoxArr(6).Value = Checked Else ChkBoxArr(6).Value = Unchecked
  If adc("UoYN7") = "Y" Then ChkBoxArr(7).Value = Checked Else ChkBoxArr(7).Value = Unchecked
  If adc("UoYN8") = "Y" Then ChkBoxArr(8).Value = Checked Else ChkBoxArr(8).Value = Unchecked
  If adc("UoYN9") = "Y" Then ChkBoxArr(9).Value = Checked Else ChkBoxArr(9).Value = Unchecked
  'BnM.3 check box Show Bag/StkId show value from UsrOpt
  If adc("UoYN10") = "Y" Then ChkBoxArr(10).Value = Checked Else ChkBoxArr(10).Value = Unchecked
  If adc("UoYN11") = "Y" Then ChkBoxArr(11).Value = Checked Else ChkBoxArr(11).Value = Unchecked
  If adc("UoYN13") = "Y" Then ChkBoxArr(13).Value = Checked Else ChkBoxArr(13).Value = Unchecked
  If adc("UoYN14") = "Y" Then ChkBoxArr(14).Value = Checked Else ChkBoxArr(14).Value = Unchecked
  If adc("UoYN15") = "Y" Then ChkBoxArr(15).Value = Checked Else ChkBoxArr(15).Value = Unchecked
  If adc("UoYN16") = "Y" Then ChkBoxArr(16).Value = Checked Else ChkBoxArr(16).Value = Unchecked
  If adc("UoYN17") = "Y" Then ChkBoxArr(17).Value = Checked Else ChkBoxArr(17).Value = Unchecked
  If adc("UoYN18") = "Y" Then ChkBoxArr(18).Value = Checked Else ChkBoxArr(18).Value = Unchecked
  If adc("UoYN19") = "Y" Then ChkBoxArr(19).Value = Checked Else ChkBoxArr(19).Value = Unchecked
  If adc("UoYN20") = "Y" Then ChkBoxArr(20).Value = Checked Else ChkBoxArr(20).Value = Unchecked
  If adc("UoYN21") = "Y" Then ChkBoxArr(21).Value = Checked Else ChkBoxArr(21).Value = Unchecked
  If adc("UoYN22") = "Y" Then ChkBoxArr(22).Value = Checked Else ChkBoxArr(22).Value = Unchecked
  If adc("UoYN23") = "Y" Then ChkBoxArr(23).Value = Checked Else ChkBoxArr(23).Value = Unchecked      ' ***** Manali 3.6.0 - 24/09/09 - Show Report in Excel Format
  If adc("UoYN24") = "Y" Then ChkBoxArr(24).Value = Checked Else ChkBoxArr(24).Value = Unchecked
  If adc("UoYN27") = "Y" Then ChkBoxArr(27).Value = Checked Else ChkBoxArr(27).Value = Unchecked      ' ***** Manali 3.10.0 - 09/03/12 - Different Suppress options for Unit Pricce and Qty
  If adc("UoYN29") = "Y" Then ChkBoxArr(29).Value = Checked Else ChkBoxArr(29).Value = Unchecked      ' 4.1.2
  If adc("UoYN30") = "Y" Then ChkBoxArr(30).Value = Checked Else ChkBoxArr(30).Value = Unchecked      ' 4.1.2
  If adc("UoYN31") = "Y" Then ChkBoxArr(31).Value = Checked Else ChkBoxArr(31).Value = Unchecked      '6.1_27
  If adc("UoYN32") = "Y" Then ChkBoxArr(32).Value = Checked Else ChkBoxArr(32).Value = Unchecked      '6.1b-15
  
  ' ****** Manali 3.5.0 - 19/11/08 - Check Boxes added
  
  ' ***** Manali 3.9.0 - Lme Selection
  Dim i As Integer
  For i = 0 To 3
    If adc("UoLme1") = ChkBoxLme(i).Caption Or adc("UoLme2") = ChkBoxLme(i).Caption Then
      ChkBoxLme(i).Value = Checked
    End If
  Next i
  ' ***** Manali 3.9.0 - Lme Selection
  
End Sub
Private Sub DispCoNm()
  'mRep.wCoCd.SetText gs_CoNm
  Dim ws_HName As String
  ws_HName = moCn.GetFldVal("Select HName from Head where HCoCd = '" + adc("UoCoCdFr") + "' and " + _
                                                    "HCd = '" + ctSelfCmCd + "' ")
  mRep.wCoCd.SetText ws_HName
  
  ' ***** Manali 3.8.0 - Comp Logo in Report
  mRep.wCoCdLogo.SetText ws_HName
  If moCn.RecSeek("Select 'x' From Head Where HCoCd='" + gs_CoCd + "' and HCd='" + ctSelfCmCd + "' and HShowRepLogo='Y' ") Then
    mRep.RHMain.UnderlaySection = True: mRep.wCoCd.Suppress = True
  Else
    mRep.RHLogo.Suppress = True
  End If
    mRep.RHLogo.Suppress = True
  ' ***** Manali 3.8.0 - Comp Logo in Report
  
' ********* Manali 3.2.2 - Before 3.2.2
'  Dim ws_DtTm As String
'  ws_DtTm = Format$(moCn.SrvrDate, "DD-MMM-YYYY")
'  mRep.wSrvrDate.SetText ws_DtTm
'  ws_DtTm = Replace(Format(IIF(CDbl(moCn.SrvrTime) >= 13#, moCn.SrvrTime - 12, IIF(CDbl(moCn.SrvrTime) < 1#, moCn.SrvrTime + 12, moCn.SrvrTime)), "00.00"), ".", ":") + IIF(CDbl(moCn.SrvrTime) >= 12#, " PM", " AM")
'  mRep.wSrvrTime.SetText ws_DtTm

' ******** Manali 3.2.2 - User date and time in reports
  Dim ws_Dt As String, ws_Tm As String
    ws_Dt = Format$(moCn.SrvrDate, IIF(UCase(MWLib.DtSetting) = "MM/DD/YY", "MMM-DD-YYYY", "DD-MMM-YYYY"))
    ws_Tm = Replace(Format(IIF(CDbl(moCn.SrvrTime) >= 13#, moCn.SrvrTime - 12, IIF(CDbl(moCn.SrvrTime) < 1#, moCn.SrvrTime + 12, moCn.SrvrTime)), "00.00"), ".", ":") + IIF(CDbl(moCn.SrvrTime) >= 12#, " PM", " AM")
  mRep.TxtRepUsrDtTime.SetText Trim(adc.UsrCd) + " / " + ws_Dt + " / " + ws_Tm
  
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
    
    Set moCn = Nothing    '*** (09/08/05)    '*** (Bef speed) bef 24/11/06
    Set mRep = Nothing
    Set Rep = Nothing
    Set Rep1 = Nothing
End Sub

Private Sub SetExcel1(ByVal pv_rsOrd As MDORowSet)
Dim sFile, sFilePath As String
Dim wOdSr As Integer
Dim wOdRow As Integer, wOrMRow As Integer, wOrDCRow As Integer, wOrXRow As Integer, wOlRow As Integer, wOmRow As Integer, wLn1 As Integer, wLn2 As Integer
Dim wCoName As String, wSht As Integer
Dim wPictPath As String
Dim wMetLs As Single, wGmFactor As Double, wCmPurityRt As Double
Dim wGldAs As String, wGldAsWt As String
Dim wCurrRow As Integer   ' ***** Manali 3.8.0
Dim wOrMFRow As Integer, wOrDCFRow As Integer, wOrXfRow As Integer, wOlFRow As Integer

  Set oExcel = CreateObject("Excel.Application")
  
  sFile = gs_CoCd + "-" + gs_UsrCd + "-" + Format(Now, "ddMMyyHHmmss") + ".xls"
  sFilePath = App.Path + "\Output\" + sFile
  
  If Dir(sFilePath, vbNormal) <> "" Then
    Kill sFilePath
  End If
  
  FileCopy App.Path + "\Template\OrderPrint.xls", sFilePath
  
  oExcel.Visible = False: oExcel.DisplayAlerts = False
  
  wCoName = moCn.GetFldVal("Select HName from Head Where HCoCd='" + gs_CoCd + "' and HCd='" + ctSelfCmCd + "'")
  
  wOdSr = 0: wOdRow = 14: wOmRow = 1: wSht = 1
  
  Set oBook = oExcel.Workbooks.Open(sFilePath)
  
  With pv_rsOrd
    Do While Not (.EOF Or .BOF)
      If wOdRow <= 285 Then
        If !OdSr <> wOdSr Then
        
          'wiSht = oExcel.Worksheets.Count
          'oExcel.Worksheets(2).Copy after:=oExcel.ActiveSheet
          oExcel.worksheets(2).Copy after:=oBook.worksheets(wSht + 1)
    
          Set oSheet = oBook.worksheets(1)
          
          If wOmRow = 1 Then
            oSheet.Name = !OmCoCd + !OmYy + !OmTc + !OmChr + CStr(!OmNo)
            oSheet.Cells(1, 1) = wCoName
            oSheet.Cells(5, 3) = !OmCoCd + "/" + !OmYy + "/" + !OmTc + "/" + !OmChr + "/" + CStr(!OmNo)
            oSheet.Cells(5, 6) = !OmDt
            oSheet.Cells(5, 9) = !OmLmgSal
            oSheet.Cells(5, 13) = !OmLmsSal
            oSheet.Cells(6, 9) = !OmLmpSal
            oSheet.Cells(6, 13) = !OmLmlSal
            oSheet.Cells(6, 3) = !OmCmCd
            oSheet.Cells(7, 3) = !CmName
            oSheet.Cells(8, 3) = !OmPoNo
            oSheet.Cells(8, 6) = !OmPoDt
            oSheet.Cells(9, 3) = !OmDelDt
            oSheet.Cells(10, 1) = "Total (" + IIF(wb_InBaseCur, ws_BaseCurCd, adc("UoCurCdFr")) + ")"
            oSheet.Cells(13, 8) = "Price (" + IIF(wb_InBaseCur, ws_BaseCurCd, adc("UoCurCdFr")) + ")"
            oSheet.Cells(13, 9) = "Value (" + IIF(wb_InBaseCur, ws_BaseCurCd, adc("UoCurCdFr")) + ")"
            
            If adc("UoYn6") = "Y" Then
              oSheet.Cells(315, 1) = !OmCmDelTerms
              oSheet.Cells(320, 1) = !OmCmPayTerms
              oSheet.Cells(325, 1) = !OmSalRem
            Else
              oSheet.Range("A315", "D328").UnMerge
              oSheet.Range("A314", "D328") = ""
            End If
          End If
          oSheet.Cells(wOdRow, 1) = !OdSr
          oSheet.Cells(wOdRow, 2) = !OdDmCd
          oSheet.Cells(wOdRow, 4) = !OdSfx
          oSheet.Cells(wOdRow, 6) = !OdDmSz
          oSheet.Cells(wOdRow, 7) = !qOrdQty
          oSheet.Cells(wOdRow, 8) = !OdSalPrc
          oSheet.Cells(wOdRow, 9).Font.Color = vbBlack
          wOdRow = wOdRow + 1
          wOmRow = wOmRow + 1
          wSht = wSht + 1
        End If
        
        'Set oSheet = oBook.Worksheets(Worksheets.Count)
        Set oSheet = oBook.worksheets(wSht + 1)
        
        If !OdSr <> wOdSr Then
          oSheet.Name = !OdDmCd + "-" + "(" + CStr(IIF(adc("UoInDet") = "INVSR/SR" Or adc("UoInDet") = "DMCD/SR" Or adc("UoInDet") = "TRNO/SR", wSht - 1, !OdSr)) + ")"
          'oSheet.Name = !OdDmCd + "-" + "(" + CStr(!OdSr) + ")"
          oSheet.Cells(4, 3) = !OdDmCd
          oSheet.Cells(4, 6) = !OdSfx
          oSheet.Cells(4, 9) = !OdDmSz
          oSheet.Cells(5, 3) = !CdOurDmCd
          oSheet.Cells(13, 13) = !OdDmCd
          oSheet.Cells(14, 13) = !IdKt
          oSheet.Cells(14, 15) = !OdDmCol
          oSheet.Cells(4, 13) = !OdPicNm
          oSheet.Cells(6, 3) = !qOrdQty
          oSheet.Cells(7, 3) = !OdPONo
          oSheet.Cells(8, 3) = !OdDelDt
          oSheet.Cells(9, 7) = !OdSalPrc
          oSheet.Cells(31, 5) = !OdLabAsWt
          oSheet.Cells(39, 10) = !OdMulBy
          
          wCurrRow = 18
          wOrMRow = 18: wOrDCRow = 23: wOrXRow = 28: wOlRow = 33
          'wOrMRow = 18: wOrDCRow = 26: wOrXRow = 45: wOlRow = 53
          
          oSheet.Cells(9, 1) = "Price (" + IIF(wb_InBaseCur, ws_BaseCurCd, adc("UoCurCdFr")) + ")"
          oSheet.Cells(9, 5) = "Sales Price (" + IIF(wb_InBaseCur, ws_BaseCurCd, adc("UoCurCdFr")) + ")"
          oSheet.Cells(wOrMRow - 1, 8) = "LME (" + IIF(wb_InBaseCur, ws_BaseCurCd, adc("UoCurCdFr")) + ")"
          oSheet.Cells(wOrMRow - 1, 9) = "Rate (" + IIF(wb_InBaseCur, ws_BaseCurCd, adc("UoCurCdFr")) + ")"
          oSheet.Cells(wOrMRow - 1, 10) = "Value (" + IIF(wb_InBaseCur, ws_BaseCurCd, adc("UoCurCdFr")) + ")"
          oSheet.Cells(wOrDCRow - 1, 10) = "Rate (" + IIF(wb_InBaseCur, ws_BaseCurCd, adc("UoCurCdFr")) + ")"
          oSheet.Cells(wOrDCRow - 1, 11) = "Value (" + IIF(wb_InBaseCur, ws_BaseCurCd, adc("UoCurCdFr")) + ")"
          oSheet.Cells(wOrDCRow - 1, 13) = "Rate (" + IIF(wb_InBaseCur, ws_BaseCurCd, adc("UoCurCdFr")) + ")"
          oSheet.Cells(wOrDCRow - 1, 14) = "Value (" + IIF(wb_InBaseCur, ws_BaseCurCd, adc("UoCurCdFr")) + ")"
          oSheet.Cells(wOrXRow - 1, 8) = "Rate (" + IIF(wb_InBaseCur, ws_BaseCurCd, adc("UoCurCdFr")) + ")"
          oSheet.Cells(wOrXRow - 1, 9) = "Value (" + IIF(wb_InBaseCur, ws_BaseCurCd, adc("UoCurCdFr")) + ")"
          oSheet.Cells(wOlRow - 1, 6) = "Rate (" + IIF(wb_InBaseCur, ws_BaseCurCd, adc("UoCurCdFr")) + ")"
          oSheet.Cells(wOlRow - 1, 7) = "Value (" + IIF(wb_InBaseCur, ws_BaseCurCd, adc("UoCurCdFr")) + ")"
          oSheet.Cells(37, 5) = "Value (" + IIF(wb_InBaseCur, ws_BaseCurCd, adc("UoCurCdFr")) + ")"
          oSheet.Cells(37, 10) = "Value (" + IIF(wb_InBaseCur, ws_BaseCurCd, adc("UoCurCdFr")) + ")"
          
          ' **** Manali 3.8.0
          wOrMRow = 1: wOrDCRow = 1: wOrXRow = 1: wOlRow = 1
          
          If adc("UoYn") = "Y" Then
            ' **** Manali 3.6.1 - Pic Reference and Dm/SM
            wPictPath = GetPictPath(IIF(UCase(adc("UoTcTypFr")) = "QS", "SM", "DM"), IIF(UCase(adc("UoPMCdFr")) = "3D", False, True), !OdDmCd, !OdDmSz, !OmCoCd, !OmTc, !OmYy, !OmChr, !OmNo, !OdSr, !OdPicNm)
            If wPictPath = "" Or Dir(wPictPath) = "" Then
              wPictPath = GetPictPath(IIF(UCase(adc("UoTcTypFr")) = "QS", "SM", "DM"), IIF(UCase(adc("UoPMCdFr")) = "3D", True, False), !OdDmCd, !OdDmSz, !OmCoCd, !OmTc, !OmYy, !OmChr, !OmNo, !OdSr, !OdPicNm)
            End If
            'wPictPath = GetPictPath("DM", False, !OdDmCd, !OdDmSz, !OmCoCd, !OmTc, !OmYy, !OmChr, !OmNo, !OdSr, !OdPicNm)
            If Dir(wPictPath) <> "" And wPictPath <> "" Then
              oSheet.Pictures.Insert(wPictPath).Select
              oSheet.Pictures(1).Name = oSheet.Name
              oSheet.Pictures(1).Height = 110: oSheet.Pictures(1).Width = 140
              oSheet.Pictures(1).Left = 572: oSheet.Pictures(1).Top = 45
              oSheet.Pictures(1).Border.LineStyle = 1
            End If
          End If
          
          If adc("UoYn6") = "Y" Then
            oSheet.Cells(48, 1) = !OdDmPrdInst
            oSheet.Cells(55, 1) = !OdCmPrdInst
            oSheet.Cells(62, 1) = !OdCmStmpInst
            oSheet.Cells(67, 1) = !OdSzInst
            oSheet.Cells(72, 1) = !OdSalRem
          Else
            oSheet.Range("A47", "G75").UnMerge
            oSheet.Range("A47", "G75") = ""
          End If
        End If
          
          If !qTag = "1RM" Then
            Select Case !RmCtg
            Case Is = "G", "P", "S", "L", "M"
              ' ****** Manali 3.8.0 -
              If wOrMRow = 1 Then wOrMFRow = wCurrRow 'Else
              oSheet.Rows(CStr(wCurrRow) + ":" + CStr(wCurrRow)).Insert
              wOrMRow = wCurrRow
              'If wOrMRow <= 21 Then
                oSheet.Cells(wOrMRow - 1, 1) = !OrSrNo
                oSheet.Cells(wOrMRow - 1, 2) = !RmCtg
                oSheet.Cells(wOrMRow - 1, 3) = !OrRmSCtg
                oSheet.Cells(wOrMRow - 1, 4) = !OrRmCd
                oSheet.Cells(wOrMRow - 1, 5) = !OrQty
                oSheet.Cells(wOrMRow - 1, 6) = IIF(adc("UoYn1") = "Y", MWLib.Div(!OrWt, !qOrdQty), !OrWt) '!OrWt
                oSheet.Cells(wOrMRow - 1, 7) = !RmQw
                oSheet.Cells(wOrMRow - 1, 8) = !OrLMESal
                
                wMetLs = moCn.GetFldVal("Select  Top 1 RrSalRt from RmRt " + _
                          "where RrTcTyp='LS' and RrCtg='" + !RmCtg + "' and RrCmCurCd='' and RrDmCd='' and RrSTWGrp='' and " + _
                          "((RrCmCtg = 'P' and RrCmCd = '" + !CmLkUpMetLs + "') Or (RrCmCtg='" + gs_CmCtg + "' and RrCmCd in ('" + !OmCmCd + "', '" + ctSelfCmCd + "'))) " + _
                          "and ((RrDmCtg = '" + !DmCtg + "' and RrLsCtg In ('" + !DmLsCtg + "', '')) Or (RrDmCtg='' and RrLsCtg='')) " + _
                          "Order By (Case When (RrCmCtg='P' and RrCmCd = '" + !CmLkUpMetLs + "') Then 1 " + _
                          "When RrCmCtg='" + gs_CmCtg + "' Then (Case RrCmCd When '" + !OmCmCd + "' Then 2 When '" + ctSelfCmCd + "' Then 3  End)  End), " + _
                          "(Case RrDmCtg When '" + !DmCtg + "' Then (Case RrLsCtg When '" + !DmLsCtg + "' Then 1 When '' Then 2  End) When '' Then 3  End) ")
  
                
                wGmFactor = moCn.GetFldVal("Select PNum from Param Where PTyp='GMFACTOR' and PMCd='" + !RmCtg + "' and PSCd='" + IIF(wb_InBaseCur, ws_BaseCurCd, adc("UoCurCdFr")) + "'")
                If !RmCtg = "M" Then
                  wGmFactor = 1
                End If
                
                
                wCmPurityRt = moCn.GetFldVal("Select RrSalRt from RmRt Where RrTcTyp='CRP' and RrCmCtg='" + gs_CmCtg + "' and RrCmCd='" + !OmCmCd + "' and RrCd='" + !OrRmCd + "' ")
                wCmPurityRt = IIF(wCmPurityRt = 0, !RmPurityRt, wCmPurityRt)
                
'                If Not (!RmCtg = "M") Then
 '                 oSheet.Cells(wOrMRow - 1, 9).formula = "=(H" + CStr(wOrMRow - 1) + "/" + CStr(wGmFactor) + "*" + CStr(wCmPurityRt) + "*(1 + " + CStr(wMetLs) + "/100)) + " + _
  '                                              "((1-" + CStr(wCmPurityRt) + ")*" + CStr(!OrAlySalRt) + ")"
   '             Else
    '              oSheet.Cells(wOrRmRow - 1, 9) = !OrSalRt
     '           End If
  
                ' ***** Manali 3.6.1 - GldAs Wt
                wGldAsWt = "$F" + CStr(wOrMRow - 1)
                If !OrMainMet = "Y" Then
                  wGldAs = moCn.GetFldVal("Select PDesc225 from Param Where Ptyp='GLDAS' and PMCd='" + !OdGldAs + "'")
                  
                  If InStr(1, "," + wGldAs + ",", ",D,") <> 0 Then wGldAsWt = wGldAsWt + IIF(wGldAsWt <> "", "+", "") + "(D63/5)"
                  If InStr(1, "," + wGldAs + ",", ",C,") <> 0 Then wGldAsWt = wGldAsWt + IIF(wGldAsWt <> "", "+", "") + "(D64/5)"
                  If InStr(1, "," + wGldAs + ",", ",X,") <> 0 Then wGldAsWt = wGldAsWt + IIF(wGldAsWt <> "", "+", "") + "(D65)"
                  
                End If
                wGldAsWt = "(" + wGldAsWt + ")"
                
                
      '          If Not (!RmCtg = "M") Then
       '           oSheet.Cells(wOrMRow - 1, 10).formula = "=$I" + CStr(wOrMRow - 1) + "*(IF($G" + CStr(wOrMRow - 1) + "=""Q"",$E" + CStr(wOrMRow - 1) + ",(IF($G" + CStr(wOrMRow - 1) + "=""W""," + wGldAsWt + ",0))))"
        '        Else
         '         oSheet.Cells(wOrRmRow - 1, 10) = !OrSalVal
          '      End If
           '     oSheet.Cells(wOrMRow - 1, 10).Font.Color = vbBlack
                
            '    oSheet.Cells(wOrMRow + 1, 5).formula = "=SUM($E" + CStr(wOrMFRow) + ":$E" + CStr(wOrMRow) + ")"
             '   oSheet.Cells(wOrMRow + 1, 6).formula = "=SUM($F" + CStr(wOrMFRow) + ":$F" + CStr(wOrMRow) + ")"
              '  oSheet.Cells(wOrMRow + 1, 10).formula = "=SUM($J" + CStr(wOrMFRow) + ":$J" + CStr(wOrMRow) + ")"
                
              ' ***** Bef 3.8.0
              'ElseIf wOrMRow = 22 Then
              '  oSheet.Cells(16, 1).AddComment "#Metal Details InComplete"
              'End If
              'wOrMRow = wOrMRow + 1
              
            Case Is = "D", "C"
                If wOrDCRow = 1 Then
                  wCurrRow = wCurrRow + 4 + IIF(wOrMRow = 1, 1, 0)
                  wOrDCFRow = wCurrRow
'                Else
                End If
                  oSheet.Rows(CStr(wCurrRow) + ":" + CStr(wCurrRow)).Insert
                wOrDCRow = wCurrRow
'              If wOrDCRow <= 40 Then
                oSheet.Cells(wOrDCRow - 1, 1) = !OrSrNo
                oSheet.Cells(wOrDCRow - 1, 2) = !RmCtg
                oSheet.Cells(wOrDCRow - 1, 3) = !OrRmSCtg
                oSheet.Cells(wOrDCRow - 1, 4) = !OrRmCd
                oSheet.Cells(wOrDCRow - 1, 5) = !OrLn1
                oSheet.Cells(wOrDCRow - 1, 6) = !OrRmPtr
                oSheet.Cells(wOrDCRow - 1, 7) = !OrQty
                oSheet.Cells(wOrDCRow - 1, 8) = IIF(adc("UoYn1") = "Y", MWLib.Div(!OrWt, !qOrdQty), !OrWt) '!OrWt
                oSheet.Cells(wOrDCRow - 1, 9) = !RmQw
                oSheet.Cells(wOrDCRow - 1, 10) = !OrSalRt
                oSheet.Cells(wOrDCRow - 1, 11).formula = "=$J" + CStr(wOrDCRow - 1) + "*(IF($I" + CStr(wOrDCRow - 1) + "=""Q"",$G" + CStr(wOrDCRow - 1) + ",(IF($I" + CStr(wOrDCRow - 1) + "=""W"",$H" + CStr(wOrDCRow - 1) + ",0))))"
                oSheet.Cells(wOrDCRow - 1, 11).Font.Color = vbBlack
                oSheet.Cells(wOrDCRow - 1, 12) = !OrSetSCd
                oSheet.Cells(wOrDCRow - 1, 13) = !OrSetSalRt
                oSheet.Cells(wOrDCRow - 1, 14).formula = "=$M" + CStr(wOrDCRow - 1) + "*$G" + CStr(wOrDCRow - 1) + ""
                oSheet.Cells(wOrDCRow - 1, 14).Font.Color = vbBlack
                
'                oSheet.Cells(wOrDCRow + 1, 7).formula = "=SUM($G" + CStr(wOrDCFRow) + ":$G" + CStr(wOrDCRow) + ")"
'                oSheet.Cells(wOrDCRow + 1, 8).formula = "=SUM($H" + CStr(wOrDCFRow) + ":$H" + CStr(wOrDCRow) + ")"
'                oSheet.Cells(wOrDCRow + 1, 11).formula = "=SUM($K" + CStr(wOrDCFRow) + ":$K" + CStr(wOrDCRow) + ")"
'                oSheet.Cells(wOrDCRow + 1, 14).formula = "=SUM($N" + CStr(wOrDCFRow) + ":$N" + CStr(wOrDCRow) + ")"
'
'                oSheet.Cells(wOrDCRow + 16, 4).formula = "=SUMIF(B" + CStr(wOrDCFRow) + ":B" + CStr(wOrDCRow) + ",""D"", H" + CStr(wOrDCFRow) + ":H" + CStr(wOrDCRow) + ")"
'                oSheet.Cells(wOrDCRow + 17, 4).formula = "=SUMIF(B" + CStr(wOrDCFRow) + ":B" + CStr(wOrDCRow) + ",""C"", H" + CStr(wOrDCFRow) + ":H" + CStr(wOrDCRow) + ")"
'                oSheet.Cells(wOrDCRow + 16, 3).formula = "=SUMIF(B" + CStr(wOrDCFRow) + ":B" + CStr(wOrDCRow) + ",""D"", G" + CStr(wOrDCFRow) + ":G" + CStr(wOrDCRow) + ")"
'                oSheet.Cells(wOrDCRow + 17, 3).formula = "=SUMIF(B" + CStr(wOrDCFRow) + ":B" + CStr(wOrDCRow) + ",""C"", G" + CStr(wOrDCFRow) + ":G" + CStr(wOrDCRow) + ")"
                
'              ElseIf wOrDCRow = 41 Then
'                oSheet.Cells(24, 1).AddComment "#Studding Details Incomplete"
'              End If
'              wOrDCRow = wOrDCRow + 1
            
            Case Is = "X"
                If wOrXRow = 1 Then
                  wCurrRow = wCurrRow + IIF(wOrDCRow = 1, 5 + IIF(wOrMRow = 1, 5, 4), 4)
                  wOrXfRow = wCurrRow
'                Else
                End If
                  oSheet.Rows(CStr(wCurrRow) + ":" + CStr(wCurrRow)).Insert
                wOrXRow = wCurrRow
'              If wOrXRow <= 48 Then
                oSheet.Cells(wOrXRow - 1, 1) = !OrSrNo
                oSheet.Cells(wOrXRow - 1, 2) = !RmCtg
                oSheet.Cells(wOrXRow - 1, 3) = !OrRmSCtg
                oSheet.Cells(wOrXRow - 1, 4) = !OrRmCd
                oSheet.Cells(wOrXRow - 1, 5) = !OrQty
                oSheet.Cells(wOrXRow - 1, 6) = IIF(adc("UoYn1") = "Y", MWLib.Div(!OrWt, !qOrdQty), !OrWt) '!OrWt
                oSheet.Cells(wOrXRow - 1, 7) = !RmQw
                oSheet.Cells(wOrXRow - 1, 8) = !OrSalRt
                oSheet.Cells(wOrXRow - 1, 9).formula = "=$H" + CStr(wOrXRow - 1) + "*(IF($G" + CStr(wOrXRow - 1) + "=""Q"",$E" + CStr(wOrXRow - 1) + ",(IF($G" + CStr(wOrXRow - 1) + "=""W"",$F" + CStr(wOrXRow - 1) + ",0))))"
                oSheet.Cells(wOrXRow - 1, 9).Font.Color = vbBlack
                
'                oSheet.Cells(wOrXRow + 1, 5).formula = "=SUM($E" + CStr(wOrXfRow) + ":$E" + CStr(wOrXRow) + ")"
'                oSheet.Cells(wOrXRow + 1, 6).formula = "=SUM($F" + CStr(wOrXfRow) + ":$F" + CStr(wOrXRow) + ")"
'                oSheet.Cells(wOrXRow + 1, 9).formula = "=SUM($I" + CStr(wOrXfRow) + ":$I" + CStr(wOrXRow) + ")"
                
'              ElseIf wOrXRow = 49 Then
'                oSheet.Cells(43, 1).AddComment = "#Acc Details InComplete"
'              End If
'              wOrXRow = wOrXRow + 1
            End Select
          
          ElseIf !qTag = "2LAB" Then
                If wOlRow = 1 Then
                  wCurrRow = wCurrRow + IIF(wOrXRow = 1, 5 + IIF(wOrDCRow = 1, 5 + IIF(wOrMRow = 1, 5, 4), 4), 4)
                  wOlFRow = wCurrRow
'                Else
                End If
                  oSheet.Rows(CStr(wCurrRow) + ":" + CStr(wCurrRow)).Insert
                wOlRow = wCurrRow
'              If wOlRow <= 57 Then
                oSheet.Cells(wOlRow, 1) = !OlSrNo
                oSheet.Cells(wOlRow, 2) = !OlMCd
                oSheet.Cells(wOlRow, 3) = !OlSCd
                oSheet.Cells(wOlRow, 4) = !OlQw
                oSheet.Cells(wOlRow, 5) = !OlQty
                oSheet.Cells(wOlRow, 6) = !OlSalRt
                ' ***** Manali 3.6.1 - Lab Val based on min/max val or OlVal
                oSheet.Cells(wOlRow, 7) = "=IF($F" + CStr(wOlRow) + " * (IF($D" + CStr(wOlRow) + "=""Q"",$E" + CStr(wOlRow) + ",(IF($D" + CStr(wOlRow) + "=""W"",$E$51,0))))=" + IIF(!OlQw = "Q", CStr(!OlQty), CStr(!OdLabAsWt)) + "*" + CStr(!OlSalRt) + "," + _
                          CStr(!OlSalVal) + ", $F" + CStr(wOlRow) + " * (IF($D" + CStr(wOlRow) + "=""Q"",$E" + CStr(wOlRow) + ",(IF($D" + CStr(wOlRow) + "=""W"",$E$51,0)))))"
                oSheet.Cells(wOlRow, 7).Font.Color = vbBlack
                
                oSheet.Cells(wOlRow + 1, 5).formula = "=SUM($E" + CStr(wOlFRow) + ":$E" + CStr(wOlRow) + ")"
                oSheet.Cells(wOlRow + 1, 7).formula = "=SUM($G" + CStr(wOlFRow) + ":$G" + CStr(wOlRow) + ")"
                
'              ElseIf wOlRow = 58 Then
'                oSheet.Cells(51, 1).AddComment = "#Labour Details InComplete"
'              End If
'              wOlRow = wOlRow + 1
          End If
          wCurrRow = wCurrRow + 1   ' **** Manali 3.8.0
        wOdSr = !OdSr
        .MoveNext
      
      Else: Exit Do
      End If
    Loop
  End With
  Set oSheet = oBook.worksheets(1)
  oSheet.Rows(CStr(wOdRow) + ":299").Hidden = True
  
  oExcel.worksheets(2).Delete

  Set pv_rsOrd = Nothing
  
     oExcel.Visible = True

  
'
'  Cd1.DialogTitle = "Save Excel File As"
'  Cd1.FileName = ADC("UoInTcFr") + "-" + ADC("UoInYyFr") + "-" + ADC("UoInChrFr") + "-" + CStr(ADC("UoInNoFr")) + ".xls"
'  Cd1.Filter = "Excel Files (*.xls) |*.xls"
'
'  Cd1.Flags = cdlOFNOverwritePrompt Or cdlOFNLongNames Or cdlOFNPathMustExist Or cdlOFNExtensionDifferent
'
'  Cd1.CancelError = True
'  Cd1.ShowSave
'
'  If Len(Cd1.FileName) > 4 Then
'    If UCase(Right(Cd1.FileName, 4)) = UCase(".xls") Then
'      oExcel.ActiveWorkbook.SaveAs Cd1.FileName
'    Else
'      wErrMsg = "File Extension Is Wrong": Exit Sub
'    End If
'  End If
'
'  If oExcel.ActiveWorkbook.Path + "\" + oExcel.ActiveWorkbook.Name <> sFilePath Then
'    Kill sFilePath
'  End If
'
'  DispMsg "Excel Generated Successfully !!", etInfo
'  oExcel.ActiveWorkbook.Close True
'  oExcel.QUIT
  
  oExcel.DisplayAlerts = True
  Set oExcel = Nothing
  Set oBook = Nothing
  Set oSheet = Nothing
Exit Sub
Err_Hndlr:
  If Err.Number = 32755 Then
    wErrMsg = "Operation Cancelled"
    oExcel.ActiveWorkbook.Close
    oExcel.QUIT
  Else
    wErrMsg = Err.Description
  End If
  If Dir(sFilePath) <> "" Then Kill sFilePath
  Err.Clear
  
End Sub
Private Sub SetExcel(ByVal pv_rsOrd As MDORowSet)

On Error GoTo Err_Hndlr:    '4.1.4

Dim sFile, sFilePath As String
Dim wOdSr As Integer
Dim wOdRow As Integer, wOrMRow As Integer, wOrDCRow As Integer, wOrXRow As Integer, wOlRow As Integer, wOmRow As Integer, wLn1 As Integer, wLn2 As Integer
Dim wCoName As String, wSht As Integer
Dim wPictPath As String
Dim wMetLs As Single, wGmFactor As Double, wCmPurityRt As Double
Dim wGldAs As String, wGldAsWt As String
Dim wCurrRow As Integer   ' ***** Manali 3.8.0
Dim wOrMFRow As Integer, wOrDCFRow As Integer, wOrXfRow As Integer, wOlFRow As Integer
Dim wLabAs As String, wLabAsWt As String
Dim wShtNm As String

  ' ***** Manali 3.9.0 - Excel For 2007 n next
  Dim wExcVer As String, wExtn As String
  Set oExcel = GetExcelObj()     'CreateObject("Excel.Application")
  
  ' ***** Manali 3.9.0 - Excel 2007
  wExcVer = oExcel.version
  wExtn = IIF(wExcVer >= "12.0", ".xlsx", ".xls")
  sFile = gs_CoCd + "-" + gs_UsrCd + "-" + Format(Now, DtFullStr()) + wExtn   '6.1
  ' ***** Manali 3.9.0 - Excel 2007
  
  'Bef 3.9.0
  'sFile = gs_CoCd + "-" + gs_UsrCd + "-" + Format(Now, "ddMMyyHHmmss") + ".xls"
  'sFilePath = App.Path + "\Output\" + sFile
  
  ' ***** Manali 3.9.0 - Output folder Path From Param
  Dim wOutputPath As String
  
  wOutputPath = moCn.GetFldVal("Select PDesc225 from Param Where PTyp='OUTPUT' and PMCd='OUTPUT' ")
  sFilePath = wOutputPath + "\" + sFile
  ' ***** Manali 3.9.0 - Output folder Path From Param
  
  If Dir(sFilePath, vbNormal) <> "" Then
    Kill sFilePath
  End If
  
  ' ***** Manali 3.9.0 - Excel 2007
  FileCopy App.Path + "\Template\OrderPrint" + wExtn, sFilePath
  'Bef 3.9.0
  'FileCopy App.Path + "\Template\OrderPrint.xls", sFilePath
  
  oExcel.Visible = False: oExcel.DisplayAlerts = False
  
  wCoName = moCn.GetFldVal("Select HName from Head Where HCoCd='" + gs_CoCd + "' and HCd='" + ctSelfCmCd + "'")
  
  wOdSr = 0: wOdRow = 14: wOmRow = 1: wSht = 1

  '4.1.0.0 - Earlier
  'Set oBook = oExcel.Workbooks.Open(sFilePath)
  
  'Sachin 4.1.0.0 - Opening the Template file with Predefined Password
  Set oBook = oExcel.Workbooks.Open(sFilePath, Password:=ctXCelPwd, WriteResPassword:=ctXCelWritePwd)

  
  With pv_rsOrd
    Do While Not (.EOF Or .BOF)
      If wOdRow <= 285 Then
        If !OdSr <> wOdSr Then
        
          'wiSht = oExcel.Worksheets.Count
          'oExcel.Worksheets(2).Copy after:=oExcel.ActiveSheet
          oExcel.worksheets(2).Copy after:=oBook.worksheets(wSht + 1)
    
          Set oSheet = oBook.worksheets(1)
          
          If wOmRow = 1 Then
            oSheet.Name = !OmCoCd + !OmYy + !OmTc + !OmChr + CStr(!OmNo)
            oSheet.Cells(1, 1) = wCoName
            oSheet.Cells(1, 7) = "DETAIL INVOICE PRINTING"
            oSheet.Cells(1, 20) = Format(Date, DtShortStr())    '6.1
            oSheet.Cells(5, 3) = !OmCoCd + "/" + !OmYy + "/" + !OmTc + "/" + !OmChr + "/" + CStr(!OmNo)
            oSheet.Cells(5, 6) = !OmDt
            oSheet.Cells(5, 9) = !OmLmgSal
            oSheet.Cells(7, 9) = !OmLmsSal           'oSheet.Cells(5, 13) = !OmLmsSal
            oSheet.Cells(6, 9) = !OmLmpSal
            oSheet.Cells(8, 9) = !OmLmlSal           'oSheet.Cells(6, 13) = !OmLmsSal
            oSheet.Cells(6, 3) = !OmCmCd
            oSheet.Cells(7, 3) = !CmName
            oSheet.Cells(8, 3) = !OmPoNo
            oSheet.Cells(8, 6) = !OmPoDt
            oSheet.Cells(9, 3) = !OmDelDt
            oSheet.Cells(10, 1) = "Total (" + IIF(wb_InBaseCur, ws_BaseCurCd, adc("UoCurCdFr")) + ")"
            oSheet.Cells(13, 18) = "Prc (" + IIF(wb_InBaseCur, ws_BaseCurCd, adc("UoCurCdFr")) + ")"       'oSheet.Cells(13, 8)
            oSheet.Cells(13, 19) = "Val (" + IIF(wb_InBaseCur, ws_BaseCurCd, adc("UoCurCdFr")) + ")"       'oSheet.Cells(13, 9)
            
            If adc("UoYn6") = "Y" Then
              oSheet.Cells(315, 1) = !OmCmDelTerms
              oSheet.Cells(320, 1) = !OmCmPayTerms
              oSheet.Cells(325, 1) = !OmSalRem
            Else
              oSheet.Range("A315", "D328").UnMerge
              oSheet.Range("A314", "D328") = ""
            End If
          End If
          
          oSheet.Cells(wOdRow, 1) = IIF(adc("UoInDet") = "INVSR/SR" Or adc("UoInDet") = "DMCD/SR" Or adc("UoInDet") = "TRNO/SR", wSht, !OdSr)
          ''oSheet.Cells(wOdRow, 1) = !OdSr
          oSheet.Cells(wOdRow, 2) = !OdDmCd
          oSheet.Cells(wOdRow, 4) = !OdSfx
          oSheet.Cells(wOdRow, 6) = !OdDmSz
          oSheet.Cells(wOdRow, 17) = !qOrdQty            'oSheet.Cells(wOdRow, 7)
          oSheet.Cells(wOdRow, 18) = !OdSalPrc           'oSheet.Cells(wOdRow, 7)       'Sales Price
          oSheet.Cells(wOdRow, 19).Font.Color = vbBlack  'oSheet.Cells(wOdRow, 7)
          wOdRow = wOdRow + 1
          wOmRow = wOmRow + 1
          wSht = wSht + 1
        End If
        
        'Set oSheet = oBook.Worksheets(Worksheets.Count)
        Set oSheet = oBook.worksheets(wSht + 1)
        
        If !OdSr <> wOdSr Then
          ' ***** Manali 3.9.0 - DC Wt, Grs Wt in 1st sheet
          wShtNm = !OdDmCd + "-" + "(" + CStr(IIF(adc("UoInDet") = "INVSR/SR" Or adc("UoInDet") = "DMCD/SR" Or adc("UoInDet") = "TRNO/SR", wSht - 1, !OdSr)) + ")"
          oSheet.Name = wShtNm
          Set oSheet = oBook.worksheets(1)
          oSheet.Cells(wOdRow - 1, 13) = "='" + wShtNm + "'!D43"     '+'" + wShtNm + "'!D44" Diamond Wt
          oSheet.Cells(wOdRow - 1, 14) = "='" + wShtNm + "'!D44"     '+'" + wShtNm + "'!D44" Color Stone Wt
          oSheet.Cells(wOdRow - 1, 16) = "='" + wShtNm + "'!D47"     ' Gross Wt
          
          '***4.1.0.0***
          oSheet.Cells(wOdRow - 1, 11) = "='" + wShtNm + "'!C43"                        'DiaQty
          oSheet.Cells(wOdRow - 1, 12) = "='" + wShtNm + "'!C44"                        'CSQty
          oSheet.Cells(wOdRow - 1, 9) = "='" + wShtNm + "'!M14"                         'OdKT
          oSheet.Cells(wOdRow - 1, 10) = "='" + wShtNm + "'!O14"                        'OdCol
          oSheet.Cells(wOdRow - 1, 7) = "='" + wShtNm + "'!C7"                          'CustDmCd
          oSheet.Cells(wOdRow - 1, 15) = "='" + wShtNm + "'!D45+'" + wShtNm + "'!D42"   'NetWt
          oSheet.Cells(wOdRow - 1, 8) = "='" + wShtNm + "'!C9"                          'OdSubPO
          oSheet.Cells(wOdRow - 1, 20) = moCn.GetFldVal("Select  dbo.MwFn_GetColStnStr('" + !OmCoCd + "','" + !OmTc + "','" + !OmYy + "','" + !OmChr + "'," + CStr(!OmNo) + "," + CStr(!OdSr) + ", '','" + adc("UOYN6") + "')") 'RmCd Str
          
          Set oSheet = oBook.worksheets(wSht + 1)
          ' ***** Manali 3.9.0 - DC Wt, Grs Wt in 1st sheet
        
          'oSheet.Name = !OdDmCd + "-" + "(" + CStr(IIF(adc("UoInDet") = "INVSR/SR" Or adc("UoInDet") = "DMCD/SR" Or adc("UoInDet") = "TRNO/SR", wSht - 1, !OdSr)) + ")"
          'oSheet.Name = !OdDmCd + "-" + "(" + CStr(!OdSr) + ")"
          oSheet.Cells(1, 1) = wCoName
          oSheet.Cells(1, 5) = "DETAIL INVOICE PRINTING"
          oSheet.Cells(1, 14) = Format(Date, DtShortStr())    '6.1

          oSheet.Cells(4, 3) = !CmName
          oSheet.Cells(5, 3) = !OmCoCd + "/" + !OmYy + "/" + !OmTc + "/" + !OmChr + "/" + CStr(!OmNo)
          oSheet.Cells(5, 6) = !OmDt
          oSheet.Cells(6, 3) = !OdDmCd
          oSheet.Cells(6, 6) = !OdSfx
          oSheet.Cells(6, 9) = !OdDmSz
          oSheet.Cells(7, 3) = !CdOurDmCd
          oSheet.Cells(13, 13) = !OdDmCd
          oSheet.Cells(14, 13) = !IdKt
          oSheet.Cells(14, 15) = !OdDmCol
          oSheet.Cells(6, 13) = !OdPicNm
          oSheet.Cells(8, 3) = !qOrdQty
          oSheet.Cells(9, 3) = !OdPONo
          oSheet.Cells(10, 3) = !OdDelDt
          oSheet.Cells(11, 7) = !OdSalPrc
          'oSheet.Cells(34, 5) = !OdLabAsWt
          oSheet.Cells(43, 10) = !OdMulBy
          
          wLabAs = moCn.GetFldVal("Select PDesc225 from Param Where Ptyp='GLDAS' and PMCd='" + !OdLabAs + "'")
          wLabAsWt = ""
          If InStr(1, "," + wLabAs + ",", ",G,") <> 0 Then wLabAsWt = wLabAsWt + IIF(wLabAsWt <> "", "+", "") + "SUMIF(B18:B19,""G"", F18:F19)"
          If InStr(1, "," + wLabAs + ",", ",P,") <> 0 Then wLabAsWt = wLabAsWt + IIF(wLabAsWt <> "", "+", "") + "SUMIF(B18:B19,""P"", F18:F19)"
          If InStr(1, "," + wLabAs + ",", ",S,") <> 0 Then wLabAsWt = wLabAsWt + IIF(wLabAsWt <> "", "+", "") + "SUMIF(B18:B19,""S"", F18:F19)"
          If InStr(1, "," + wLabAs + ",", ",L,") <> 0 Then wLabAsWt = wLabAsWt + IIF(wLabAsWt <> "", "+", "") + "SUMIF(B18:B19,""L"", F18:F19)"
          If InStr(1, "," + wLabAs + ",", ",M,") <> 0 Then wLabAsWt = wLabAsWt + IIF(wLabAsWt <> "", "+", "") + "SUMIF(B18:B19,""M"", F18:F19)"
          If InStr(1, "," + wLabAs + ",", ",D,") <> 0 Then wLabAsWt = wLabAsWt + IIF(wLabAsWt <> "", "+", "") + "(D43/5)"
          If InStr(1, "," + wLabAs + ",", ",C,") <> 0 Then wLabAsWt = wLabAsWt + IIF(wLabAsWt <> "", "+", "") + "(D44/5)"
          If InStr(1, "," + wLabAs + ",", ",X,") <> 0 Then wLabAsWt = wLabAsWt + IIF(wLabAsWt <> "", "+", "") + "(D45)"
          wLabAsWt = "=(" + wLabAsWt + ")"
          oSheet.Cells(34, 5) = wLabAsWt    '!OdLabAsWt
          
          wCurrRow = 18
          'wOrMRow = 18: wOrDCRow = 24: wOrXRow = 30: wOlRow = 36
          'wOrMRow = 18: wOrDCRow = 26: wOrXRow = 45: wOlRow = 53
          
          oSheet.Cells(11, 1) = "Price (" + IIF(wb_InBaseCur, ws_BaseCurCd, adc("UoCurCdFr")) + ")"
          oSheet.Cells(11, 5) = "Sales Price (" + IIF(wb_InBaseCur, ws_BaseCurCd, adc("UoCurCdFr")) + ")"
          oSheet.Cells(17, 8) = "LME (" + IIF(wb_InBaseCur, ws_BaseCurCd, adc("UoCurCdFr")) + ")"
          oSheet.Cells(17, 9) = "Rate (" + IIF(wb_InBaseCur, ws_BaseCurCd, adc("UoCurCdFr")) + ")"
          oSheet.Cells(17, 10) = "Value (" + IIF(wb_InBaseCur, ws_BaseCurCd, adc("UoCurCdFr")) + ")"
          oSheet.Cells(23, 10) = "Rate (" + IIF(wb_InBaseCur, ws_BaseCurCd, adc("UoCurCdFr")) + ")"
          oSheet.Cells(23, 11) = "Value (" + IIF(wb_InBaseCur, ws_BaseCurCd, adc("UoCurCdFr")) + ")"
          oSheet.Cells(23, 13) = "Rate (" + IIF(wb_InBaseCur, ws_BaseCurCd, adc("UoCurCdFr")) + ")"
          oSheet.Cells(23, 14) = "Value (" + IIF(wb_InBaseCur, ws_BaseCurCd, adc("UoCurCdFr")) + ")"
          oSheet.Cells(29, 8) = "Rate (" + IIF(wb_InBaseCur, ws_BaseCurCd, adc("UoCurCdFr")) + ")"
          oSheet.Cells(29, 9) = "Value (" + IIF(wb_InBaseCur, ws_BaseCurCd, adc("UoCurCdFr")) + ")"
          oSheet.Cells(35, 6) = "Rate (" + IIF(wb_InBaseCur, ws_BaseCurCd, adc("UoCurCdFr")) + ")"
          oSheet.Cells(35, 7) = "Value (" + IIF(wb_InBaseCur, ws_BaseCurCd, adc("UoCurCdFr")) + ")"
          oSheet.Cells(41, 5) = "Value (" + IIF(wb_InBaseCur, ws_BaseCurCd, adc("UoCurCdFr")) + ")"
          oSheet.Cells(41, 10) = "Value (" + IIF(wb_InBaseCur, ws_BaseCurCd, adc("UoCurCdFr")) + ")"
          
          ' **** Manali 3.8.0
          wOrMRow = 1: wOrDCRow = 1: wOrXRow = 1: wOlRow = 1
          
          If adc("UoYn") = "Y" Then
            ' **** Manali 3.6.1 - Pic Reference and Dm/SM
            wPictPath = GetPictPath("DM", False, !OdDmCd, !OdDmSz, !OmCoCd, !OmTc, !OmYy, !OmChr, !OmNo, !OdSr, !OdPicNm)
            'wPictPath = GetPictPath("DM", False, !OdDmCd, !OdDmSz, !OmCoCd, !OmTc, !OmYy, !OmChr, !OmNo, !OdSr, !OdPicNm)
            If Dir(wPictPath) <> "" And wPictPath <> "" Then
              ' ***** Manali 3.9.0 - Change Done For Excel 2010 to embed image
              ' ***** Help - expression.AddPicture(FileName, LinkToFile, SaveWithDocument, Left, Top, Width, Height)
              oSheet.Shapes.AddPicture wPictPath, False, True, 577, 45, 140, 110
              oSheet.Pictures(1).ShapeRange.LockAspectRatio = False
              oSheet.Pictures(1).Name = oSheet.Name
              oSheet.Pictures(1).Border.LineStyle = 1
              
              ' Bef 3.9.0
              'oSheet.Pictures.Insert(wPictPath).Select
              'oSheet.Pictures(1).Left = 577: oSheet.Pictures(1).Top = 45
              'oSheet.Pictures(1).ShapeRange.LockAspectRatio = False
            End If
          End If
          
          If adc("UoYn6") = "Y" Then
            oSheet.Cells(52, 1) = !OdDmPrdInst
            oSheet.Cells(59, 1) = !OdCmPrdInst
            oSheet.Cells(66, 1) = !OdCmStmpInst
            oSheet.Cells(71, 1) = !OdSzInst
            oSheet.Cells(76, 1) = !OdSalRem
          Else
            oSheet.Range("A51", "G79").UnMerge
            oSheet.Range("A51", "G79") = ""
          End If
        End If
          
          If !qTag = "1RM" Then
            Select Case !RmCtg
            Case Is = "G", "P", "S", "L", "M"
              ' ****** Manali 3.8.0 -
              'wOrMRow = wOrMRow + 1
              If wOrMRow = 1 Then
                wOrMFRow = wCurrRow
              Else
                oSheet.Rows(CStr(wCurrRow) + ":" + CStr(wCurrRow)).Insert
                oSheet.Range("A" + CStr(wCurrRow - 1) + ":J" + CStr(wCurrRow - 1)).Copy
                oSheet.Range("A" + CStr(wCurrRow) + ":J" + CStr(wCurrRow)).Select
                oSheet.Paste
              End If
              
              'If wOrMRow <= 21 Then
                oSheet.Cells(wCurrRow, 1) = !OrSrNo
                oSheet.Cells(wCurrRow, 2) = !RmCtg
                oSheet.Cells(wCurrRow, 3) = !OrRmSCtg
                oSheet.Cells(wCurrRow, 4) = !OrRmCd
                oSheet.Cells(wCurrRow, 5) = !OrQty
                oSheet.Cells(wCurrRow, 6) = IIF(adc("UoYn1") = "Y", MWLib.Div(!OrWt, !qOrdQty), !OrWt)  '!OrWt
                oSheet.Cells(wCurrRow, 7) = !RmQw
                oSheet.Cells(wCurrRow, 8) = !OrLMESal
                               
                'indigo.162- getting voucher currency
                Dim wVchCur As String, wGldLsPer  As Single
                wVchCur = moCn.GetFldVal("Select CmCurCd from CustMst where CmCtg= '" + gs_CmCtg + "' and CmCd='" + !OmCmCd + "'")
                                              
                'indigo.162- calling GetLsPer function and getting Ls % from there.
                'first Ls calculation was in GetGldRt function. now that part cut and added in a new function GetLsPer
                wMetLs = GetGldRtLsPer(!RmCtg, gs_CmCtg, !OmCmCd, !OdDmCd, !IdKt, !DmCtg, !DmLsCtg, "RrSalRt", wVchCur, !CmLkUpMetLs)
                                                     
                wGmFactor = moCn.GetFldVal("Select PNum from Param Where PTyp='GMFACTOR' and PMCd='" + !RmCtg + "' and PSCd='" + IIF(wb_InBaseCur, ws_BaseCurCd, adc("UoCurCdFr")) + "'")
                
                If !RmCtg = "M" Then
                  wGmFactor = 1
                End If
                
                
                wCmPurityRt = moCn.GetFldVal("Select RrSalRt from RmRt Where RrTcTyp='CRP' and RrCmCtg='" + gs_CmCtg + "' and RrCmCd='" + !OmCmCd + "' and RrCd='" + !OrRmCd + "' ")
                wCmPurityRt = IIF(wCmPurityRt = 0, !RmPurityRt, wCmPurityRt)
                
                If Not (!RmCtg = "M") Then
                  oSheet.Cells(wCurrRow, 9).formula = "=(H" + CStr(wCurrRow) + "/" + CStr(wGmFactor) + "*" + CStr(wCmPurityRt) + "*(1 + " + CStr(wMetLs) + "/100)) + " + _
                                                "((1-" + CStr(wCmPurityRt) + ")*" + CStr(!OrAlySalRt) + ")"
                Else
                  oSheet.Cells(wCurrRow, 9) = !OrSalRt
                End If
  
                
'                wCmPurityRt = moCn.GetFldVal("Select RrSalRt from RmRt Where RrTcTyp='CRP' and RrCmCtg='" + gs_CmCtg + "' and RrCmCd='" + !OmCmCd + "' and RrCd='" + !OrRmCd + "' ")
 '               wCmPurityRt = IIF(wCmPurityRt = 0, !RmPurityRt, wCmPurityRt)
                
  '              oSheet.Cells(wCurrRow, 9).formula = "=(H" + CStr(wCurrRow) + "/" + CStr(wGmFactor) + "*" + CStr(wCmPurityRt) + "*(1 + " + CStr(wMetLs) + "/100)) + " + _
   '                                             "((1-" + CStr(wCmPurityRt) + ")*" + CStr(!OrAlySalRt) + ")"
  
                ' ***** Manali 3.6.1 - GldAs Wt
                wGldAsWt = "$F" + CStr(wCurrRow)
                If !OrMainMet = "Y" Then
                  wGldAs = moCn.GetFldVal("Select PDesc225 from Param Where Ptyp='GLDAS' and PMCd='" + !OdGldAs + "'")
                  
                  If InStr(1, "," + wGldAs + ",", ",D,") <> 0 Then wGldAsWt = wGldAsWt + IIF(wGldAsWt <> "", "+", "") + "(D43/5)"
                  If InStr(1, "," + wGldAs + ",", ",C,") <> 0 Then wGldAsWt = wGldAsWt + IIF(wGldAsWt <> "", "+", "") + "(D44/5)"
                  If InStr(1, "," + wGldAs + ",", ",X,") <> 0 Then wGldAsWt = wGldAsWt + IIF(wGldAsWt <> "", "+", "") + "(D45)"
                  
                End If
                wGldAsWt = "(" + wGldAsWt + ")"
                                        
                oSheet.Cells(wCurrRow, 10).formula = "=$I" + CStr(wCurrRow) + "*(IF($G" + CStr(wCurrRow) + "=""Q"",$E" + CStr(wCurrRow) + ",(IF($G" + CStr(wCurrRow) + "=""W""," + wGldAsWt + ",0))))"
                oSheet.Cells(wCurrRow, 10).Font.Color = vbBlack
                
                
              ' ***** Bef 3.8.0
              'ElseIf wOrMRow = 22 Then
              '  oSheet.Cells(16, 1).AddComment "#Metal Details InComplete"
              'End If
              wOrMRow = wOrMRow + 1
              
            Case Is = "D", "C"
              'wOrDCRow = wOrDCRow + 1
                If wOrDCRow = 1 Then
                  wCurrRow = wCurrRow + IIF(wOrMRow = 1, 6, IIF(wOrMRow = 2, 5, 5))
                  wOrDCFRow = wCurrRow
                Else  'If wOrDCRow <> 2 Then
                  oSheet.Rows(CStr(wCurrRow) + ":" + CStr(wCurrRow)).Insert
                  oSheet.Range("A" + CStr(wCurrRow - 1) + ":N" + CStr(wCurrRow - 1)).Copy
                  oSheet.Range("A" + CStr(wCurrRow) + ":N" + CStr(wCurrRow)).Select
                  oSheet.Paste
                End If
                
'              If wOrDCRow <= 40 Then
                oSheet.Cells(wCurrRow, 1) = !OrSrNo
                oSheet.Cells(wCurrRow, 2) = !RmCtg
                oSheet.Cells(wCurrRow, 3) = !OrRmSCtg
                oSheet.Cells(wCurrRow, 4) = !OrRmCd
                oSheet.Cells(wCurrRow, 5) = !OrLn1
                oSheet.Cells(wCurrRow, 6) = !OrRmPtr
                oSheet.Cells(wCurrRow, 7) = !OrQty
                oSheet.Cells(wCurrRow, 8) = IIF(adc("UoYn1") = "Y", MWLib.Div(!OrWt, !qOrdQty), !OrWt)  '!OrWt
                oSheet.Cells(wCurrRow, 9) = !RmQw
                oSheet.Cells(wCurrRow, 10) = !OrSalRt
                oSheet.Cells(wCurrRow, 11).formula = "=$J" + CStr(wCurrRow) + "*(IF($I" + CStr(wCurrRow) + "=""Q"",$G" + CStr(wCurrRow) + ",(IF($I" + CStr(wCurrRow) + "=""W"",$H" + CStr(wCurrRow) + ",0))))"
                oSheet.Cells(wCurrRow, 11).Font.Color = vbBlack
                oSheet.Cells(wCurrRow, 12) = !OrSetSCd
                oSheet.Cells(wCurrRow, 13) = !OrSetSalRt
                oSheet.Cells(wCurrRow, 14).formula = "=$M" + CStr(wCurrRow) + "*$G" + CStr(wCurrRow) + ""
                oSheet.Cells(wCurrRow, 14).Font.Color = vbBlack
                
                
              wOrDCRow = wOrDCRow + 1
            
            Case Is = "X"
                If wOrXRow = 1 Then
                  wCurrRow = wCurrRow + IIF(wOrDCRow = 1, 6 + IIF(wOrMRow = 1, 6, IIF(wOrMRow = 2, 5, 5)), IIF(wOrDCRow = 2, 5, 5))
                  wOrXfRow = wCurrRow
                Else    'If wOrXRow <> 2 Then
                  oSheet.Rows(CStr(wCurrRow) + ":" + CStr(wCurrRow)).Insert
                  oSheet.Range("A" + CStr(wCurrRow - 1) + ":I" + CStr(wCurrRow - 1)).Copy
                  oSheet.Range("A" + CStr(wCurrRow) + ":I" + CStr(wCurrRow)).Select
                  oSheet.Paste
                End If
'              If wOrXRow <= 48 Then
                oSheet.Cells(wCurrRow, 1) = !OrSrNo
                oSheet.Cells(wCurrRow, 2) = !RmCtg
                oSheet.Cells(wCurrRow, 3) = !OrRmSCtg
                oSheet.Cells(wCurrRow, 4) = !OrRmCd
                oSheet.Cells(wCurrRow, 5) = !OrQty
                oSheet.Cells(wCurrRow, 6) = IIF(adc("UoYn1") = "Y", MWLib.Div(!OrWt, !qOrdQty), !OrWt)    '!OrWt
                oSheet.Cells(wCurrRow, 7) = !RmQw
                oSheet.Cells(wCurrRow, 8) = !OrSalRt
                oSheet.Cells(wCurrRow, 9).formula = "=$H" + CStr(wCurrRow) + "*(IF($G" + CStr(wCurrRow) + "=""Q"",$E" + CStr(wCurrRow) + ",(IF($G" + CStr(wCurrRow) + "=""W"",$F" + CStr(wCurrRow) + ",0))))"
                oSheet.Cells(wCurrRow, 9).Font.Color = vbBlack
                
                
                
'              ElseIf wOrXRow = 49 Then
'                oSheet.Cells(43, 1).AddComment = "#Acc Details InComplete"
'              End If
              wOrXRow = wOrXRow + 1
            End Select
          
          ElseIf !qTag = "2LAB" Then
                If wOlRow = 1 Then
                  wCurrRow = wCurrRow + IIF(wOrXRow = 1, 6 + IIF(wOrDCRow = 1, 6 + IIF(wOrMRow = 1, 6, IIF(wOrMRow = 2, 5, 5)), IIF(wOrDCRow = 2, 5, 5)), _
                                        IIF(wOrXRow = 2, 5, 5))
                  wOlFRow = wCurrRow
                Else    'If wOlRow <> 2 Then
                  oSheet.Rows(CStr(wCurrRow) + ":" + CStr(wCurrRow)).Insert
                  oSheet.Range("A" + CStr(wCurrRow - 1) + ":G" + CStr(wCurrRow - 1)).Copy
                  oSheet.Range("A" + CStr(wCurrRow) + ":G" + CStr(wCurrRow)).Select
                  oSheet.Paste
                End If
'              If wOlRow <= 57 Then
                oSheet.Cells(wCurrRow, 1) = !OlSrNo
                oSheet.Cells(wCurrRow, 2) = !OlMCd
                oSheet.Cells(wCurrRow, 3) = !OlSCd
                oSheet.Cells(wCurrRow, 4) = !OlQw
                oSheet.Cells(wCurrRow, 5) = !OlQty
                oSheet.Cells(wCurrRow, 6) = !OlSalRt
                ' ***** Manali 3.6.1 - Lab Val based on min/max val or OlVal
                oSheet.Cells(wCurrRow, 7) = "=IF($F" + CStr(wCurrRow) + " * (IF($D" + CStr(wCurrRow) + "=""Q"",$E" + CStr(wCurrRow) + ",(IF($D" + CStr(wCurrRow) + "=""W"",$E$" + CStr(wOlFRow - 2) + ",0))))=" + IIF(!OlQw = "Q", CStr(!OlQty), CStr(!OdLabAsWt)) + "*" + CStr(!OlSalRt) + "," + _
                          CStr(!OlSalVal) + ", $F" + CStr(wCurrRow) + " * (IF($D" + CStr(wCurrRow) + "=""Q"",$E" + CStr(wCurrRow) + ",(IF($D" + CStr(wCurrRow) + "=""W"",$E$" + CStr(wOlFRow - 2) + ",0)))))"
                oSheet.Cells(wCurrRow, 7).Font.Color = vbBlack
                
'              ElseIf wOlRow = 58 Then
'                oSheet.Cells(51, 1).AddComment = "#Labour Details InComplete"
'              End If
              wOlRow = wOlRow + 1
          End If
          wCurrRow = wCurrRow + 1   ' **** Manali 3.8.0
        wOdSr = !OdSr
        .MoveNext
      
      Else: Exit Do
      End If
    Loop
  End With
  Set oSheet = oBook.worksheets(1)
  oSheet.Rows(CStr(wOdRow) + ":299").Hidden = True
  
  oExcel.worksheets(2).Delete

  Set pv_rsOrd = Nothing
  
  
  Cd1.DialogTitle = "Save Excel File As"
  ' ***** Manali 3.9.0 - Excel For 2007 n next
  Cd1.FileName = adc("UoInTcFr") + "-" + adc("UoInYyFr") + "-" + adc("UoInChrFr") + "-" + CStr(adc("UoInNoFr")) + wExtn
  Cd1.Filter = "Excel Files (*" + wExtn + ") |*" + wExtn
  
  ' Bef 3.9.0
  'CD1.FileName = adc("UoInTcFr") + "-" + adc("UoInYyFr") + "-" + adc("UoInChrFr") + "-" + CStr(adc("UoInNoFr")) + ".xls"
  'CD1.Filter = "Excel Files (*.xls) |*.xls"
  
  Cd1.Flags = cdlOFNOverwritePrompt Or cdlOFNLongNames Or cdlOFNPathMustExist Or cdlOFNExtensionDifferent
  
  Cd1.CancelError = True
  Cd1.ShowSave

  If Len(Cd1.FileName) > 4 Then
    ' ***** Manali 3.9.0 - Excel For 2007 n next
    If UCase(Right(Cd1.FileName, 4)) = UCase(wExtn) Or UCase(Right(Cd1.FileName, 5)) = UCase(wExtn) Then
    ' bef  3.9.0
    'If UCase(Right(CD1.FileName, 4)) = UCase(".xls") Then
      oExcel.ActiveWorkbook.SaveAs Cd1.FileName, , "", "" '4.1.0.0
    Else
      wErrMsg = "File Extension Is Wrong": Exit Sub
    End If
  End If
  
  If oExcel.ActiveWorkbook.Path + "\" + oExcel.ActiveWorkbook.Name <> sFilePath Then
    Kill sFilePath
  End If
  
  DispMsg "Excel Generated Successfully !!", etInfo
  oExcel.ActiveWorkbook.Close True
  oExcel.QUIT
  
  oExcel.DisplayAlerts = True
  Set oExcel = Nothing
  Set oBook = Nothing
  Set oSheet = Nothing
Exit Sub
Err_Hndlr:
  If Err.Number = 32755 Then
    wErrMsg = "Operation Cancelled"
    oExcel.ActiveWorkbook.Close
    oExcel.QUIT
  ElseIf Err.Number = 70 Then       '4.1.4
    wErrMsg = "Cannot Access Template File. Close If Already Open and Continue."
  Else
    wErrMsg = Err.Description
  End If
  If Dir(sFilePath) <> "" Then Kill sFilePath
  Err.Clear
  

End Sub

Private Sub SetExcel_bef380(ByVal pv_rsOrd As MDORowSet)
Dim sFile, sFilePath As String
Dim wOdSr As Integer
Dim wOdRow As Integer, wOrMRow As Integer, wOrDCRow As Integer, wOrXRow As Integer, wOlRow As Integer, wOmRow As Integer, wLn1 As Integer, wLn2 As Integer
Dim wCoName As String, wSht As Integer
Dim wPictPath As String
Dim wMetLs As Single, wGmFactor As Double, wCmPurityRt As Double
Dim wGldAs As String, wGldAsWt As String

  Set oExcel = CreateObject("Excel.Application")
  
  sFile = gs_CoCd + "-" + gs_UsrCd + "-" + Format(Now, "ddMMyyHHmmss") + ".xls"
  sFilePath = App.Path + "\Output\" + sFile
  
  If Dir(sFilePath, vbNormal) <> "" Then
    Kill sFilePath
  End If
  
  FileCopy App.Path + "\Template\OrderPrint.xls", sFilePath
  
  oExcel.Visible = False: oExcel.DisplayAlerts = False
  
  wCoName = moCn.GetFldVal("Select HName from Head Where HCoCd='" + gs_CoCd + "' and HCd='" + ctSelfCmCd + "'")
  
  wOdSr = 0: wOdRow = 14: wOmRow = 1: wSht = 1
  
  Set oBook = oExcel.Workbooks.Open(sFilePath)
  
  With pv_rsOrd
    Do While Not (.EOF Or .BOF)
      If wOdRow <= 285 Then
        If !OdSr <> wOdSr Then
        
          'wiSht = oExcel.Worksheets.Count
          'oExcel.Worksheets(2).Copy after:=oExcel.ActiveSheet
          oExcel.worksheets(2).Copy after:=oBook.worksheets(wSht + 1)
    
          Set oSheet = oBook.worksheets(1)
          
          If wOmRow = 1 Then
            oSheet.Name = !OmCoCd + !OmYy + !OmTc + !OmChr + CStr(!OmNo)
            oSheet.Cells(1, 1) = wCoName
            oSheet.Cells(5, 3) = !OmCoCd + "/" + !OmYy + "/" + !OmTc + "/" + !OmChr + "/" + CStr(!OmNo)
            oSheet.Cells(5, 6) = !OmDt
            oSheet.Cells(5, 9) = !OmLmgSal
            oSheet.Cells(5, 13) = !OmLmsSal
            oSheet.Cells(6, 9) = !OmLmpSal
            oSheet.Cells(6, 13) = !OmLmlSal
            oSheet.Cells(6, 3) = !OmCmCd
            oSheet.Cells(7, 3) = !CmName
            oSheet.Cells(8, 1) = "Export No "
            oSheet.Cells(8, 3) = !OmPoNo
            oSheet.Cells(8, 6) = !OmPoDt
            oSheet.Cells(9, 3) = !OmDelDt
            oSheet.Cells(10, 1) = "Total (" + IIF(wb_InBaseCur, ws_BaseCurCd, adc("UoCurCdFr")) + ")"
            oSheet.Cells(13, 8) = "Price (" + IIF(wb_InBaseCur, ws_BaseCurCd, adc("UoCurCdFr")) + ")"
            oSheet.Cells(13, 9) = "Value (" + IIF(wb_InBaseCur, ws_BaseCurCd, adc("UoCurCdFr")) + ")"
            
            oSheet.Range("A315", "D328").UnMerge
            oSheet.Range("A314", "D328") = ""
          End If
          oSheet.Cells(wOdRow, 1) = IIF(adc("UoInDet") = "INVSR/SR" Or adc("UoInDet") = "DMCD/SR" Or adc("UoInDet") = "TRNO/SR", wSht, !OdSr)
          oSheet.Cells(wOdRow, 2) = !OdDmCd
          oSheet.Cells(wOdRow, 4) = !OdSfx
          oSheet.Cells(wOdRow, 6) = !OdDmSz
          oSheet.Cells(wOdRow, 7) = !qOrdQty
          oSheet.Cells(wOdRow, 8) = !OdSalPrc
          oSheet.Cells(wOdRow, 9).Font.Color = vbBlack
          wOdRow = wOdRow + 1
          wOmRow = wOmRow + 1
          wSht = wSht + 1
        End If
        
        'Set oSheet = oBook.Worksheets(Worksheets.Count)
        Set oSheet = oBook.worksheets(wSht + 1)
        
        If !OdSr <> wOdSr Then
          oSheet.Name = !OdDmCd + "-" + "(" + CStr(IIF(adc("UoInDet") = "INVSR/SR" Or adc("UoInDet") = "DMCD/SR" Or adc("UoInDet") = "TRNO/SR", wSht - 1, !OdSr)) + ")"
          oSheet.Cells(1, 1) = wCoName
          oSheet.Cells(4, 3) = !OdDmCd
          oSheet.Cells(4, 6) = !OdSfx
          oSheet.Cells(4, 9) = !OdDmSz
          oSheet.Cells(5, 3) = !CdOurDmCd
          oSheet.Cells(13, 13) = !OdDmCd
          oSheet.Cells(14, 13) = !IdKt
          oSheet.Cells(14, 15) = !OdDmCol
          oSheet.Cells(4, 13) = !OdPicNm
          oSheet.Cells(6, 3) = !qOrdQty
          oSheet.Cells(7, 3) = !OdPONo
          oSheet.Cells(8, 3) = !OdDelDt
          oSheet.Cells(9, 7) = !OdSalPrc
          oSheet.Cells(51, 5) = IIF(adc("UoYn1") = "Y", MWLib.Div(!OdLabAsWt, !qOrdQty), !OdLabAsWt)
          oSheet.Cells(63, 10) = !OdMulBy
          
          wOrMRow = 18: wOrDCRow = 26: wOrXRow = 45: wOlRow = 53
          
          oSheet.Cells(9, 1) = "Price (" + IIF(wb_InBaseCur, ws_BaseCurCd, adc("UoCurCdFr")) + ")"
          oSheet.Cells(9, 5) = "Sales Price (" + IIF(wb_InBaseCur, ws_BaseCurCd, adc("UoCurCdFr")) + ")"
          oSheet.Cells(wOrMRow - 1, 8) = "LME (" + IIF(wb_InBaseCur, ws_BaseCurCd, adc("UoCurCdFr")) + ")"
          oSheet.Cells(wOrMRow - 1, 9) = "Rate (" + IIF(wb_InBaseCur, ws_BaseCurCd, adc("UoCurCdFr")) + ")"
          oSheet.Cells(wOrMRow - 1, 10) = "Value (" + IIF(wb_InBaseCur, ws_BaseCurCd, adc("UoCurCdFr")) + ")"
          oSheet.Cells(wOrDCRow - 1, 10) = "Rate (" + IIF(wb_InBaseCur, ws_BaseCurCd, adc("UoCurCdFr")) + ")"
          oSheet.Cells(wOrDCRow - 1, 11) = "Value (" + IIF(wb_InBaseCur, ws_BaseCurCd, adc("UoCurCdFr")) + ")"
          oSheet.Cells(wOrDCRow - 1, 13) = "Rate (" + IIF(wb_InBaseCur, ws_BaseCurCd, adc("UoCurCdFr")) + ")"
          oSheet.Cells(wOrDCRow - 1, 14) = "Value (" + IIF(wb_InBaseCur, ws_BaseCurCd, adc("UoCurCdFr")) + ")"
          oSheet.Cells(wOrXRow - 1, 8) = "Rate (" + IIF(wb_InBaseCur, ws_BaseCurCd, adc("UoCurCdFr")) + ")"
          oSheet.Cells(wOrXRow - 1, 9) = "Value (" + IIF(wb_InBaseCur, ws_BaseCurCd, adc("UoCurCdFr")) + ")"
          oSheet.Cells(wOlRow - 1, 6) = "Rate (" + IIF(wb_InBaseCur, ws_BaseCurCd, adc("UoCurCdFr")) + ")"
          oSheet.Cells(wOlRow - 1, 7) = "Value (" + IIF(wb_InBaseCur, ws_BaseCurCd, adc("UoCurCdFr")) + ")"
          oSheet.Cells(61, 5) = "Value (" + IIF(wb_InBaseCur, ws_BaseCurCd, adc("UoCurCdFr")) + ")"
          oSheet.Cells(61, 10) = "Value (" + IIF(wb_InBaseCur, ws_BaseCurCd, adc("UoCurCdFr")) + ")"
          
          If adc("UoYn") = "Y" Then
            wPictPath = GetPictPath("DM", False, !OdDmCd, !OdDmSz, !OmCoCd, !OmTc, !OmYy, !OmChr, !OmNo, !OdSr, !OdPicNm)
            If Dir(wPictPath) <> "" And wPictPath <> "" Then
              oSheet.Pictures.Insert(wPictPath).Select
              oSheet.Pictures(1).Name = oSheet.Name
              oSheet.Pictures(1).Height = 110: oSheet.Pictures(1).Width = 140
              oSheet.Pictures(1).Left = 572: oSheet.Pictures(1).Top = 45
              oSheet.Pictures(1).Border.LineStyle = 1
            End If
          End If
          
          oSheet.Range("A71", "G100").UnMerge
          oSheet.Range("A71", "G100") = ""
          
        End If
          
          
          If !qTag = "1RM" Then
'            If adc("UoYn1") = "Y" Then !OrWt = MWLib.Div(!OrWt, !qOrdQty)
            
            Select Case !RmCtg
            Case Is = "G", "P", "S", "L", "M"
              If wOrMRow <= 21 Then
                oSheet.Cells(wOrMRow, 1) = !OrSrNo
                oSheet.Cells(wOrMRow, 2) = !RmCtg
                oSheet.Cells(wOrMRow, 3) = !OrRmSCtg
                oSheet.Cells(wOrMRow, 4) = !OrRmCd
                oSheet.Cells(wOrMRow, 5) = !OrQty
                oSheet.Cells(wOrMRow, 6) = IIF(adc("UoYn1") = "Y", MWLib.Div(!OrWt, !qOrdQty), !OrWt)
                oSheet.Cells(wOrMRow, 7) = !RmQw
                oSheet.Cells(wOrMRow, 8) = !OrLMESal
                
                wMetLs = moCn.GetFldVal("Select  Top 1 " + IIF(adc("UoInVal") = "A" Or adc("UoInVal") = "A", "RrCstRt", "RrSalRt") + " from RmRt " + _
                          "where RrTcTyp='LS' and RrCtg='" + !RmCtg + "' and RrCmCurCd='' and RrDmCd='' and RrSTWGrp='' and " + _
                          "((RrCmCtg = 'P' and RrCmCd = '" + !CmLkUpMetLs + "') Or (RrCmCtg='C' and RrCmCd in ('" + !OmCmCd + "', '" + ctSelfCmCd + "'))) " + _
                          "and ((RrDmCtg = '" + !DmCtg + "' and RrLsCtg In ('" + !DmLsCtg + "', '')) Or (RrDmCtg='' and RrLsCtg='')) " + _
                          "Order By (Case When (RrCmCtg='P' and RrCmCd = '" + !CmLkUpMetLs + "') Then 1 " + _
                          "When RrCmCtg='C' Then (Case RrCmCd When '" + !OmCmCd + "' Then 2 When '" + ctSelfCmCd + "' Then 3  End)  End), " + _
                          "(Case RrDmCtg When '" + !DmCtg + "' Then (Case RrLsCtg When '" + !DmLsCtg + "' Then 1 When '' Then 2  End) When '' Then 3  End) ")
  
                
                wGmFactor = moCn.GetFldVal("Select PNum from Param Where PTyp='GMFACTOR' and PMCd='" + !RmCtg + "' and PSCd='" + IIF(wb_InBaseCur, ws_BaseCurCd, adc("UoCurCdFr")) + "'")
                wCmPurityRt = IIF(adc("UoInVal") = "I", moCn.GetFldVal("Select RrSalRt from RmRt Where RrTcTyp='CRP' and RrCmCtg='C' and RrCmCd='" + !OmCmCd + "' and RrCd='" + !OrRmCd + "' "), 0)
                wCmPurityRt = IIF(wCmPurityRt = 0, !RmPurityRt, wCmPurityRt)
                
                oSheet.Cells(wOrMRow, 9).formula = "=(H" + CStr(wOrMRow) + "/" + CStr(wGmFactor) + "*" + CStr(wCmPurityRt) + "*(1 + " + CStr(wMetLs) + "/100)) + " + _
                                                "((1-" + CStr(wCmPurityRt) + ")*" + CStr(!OrAlySalRt) + ")"
  
                ' ***** Manali 3.6.1 - GldAs Wt
                wGldAsWt = "$F" + CStr(wOrMRow)
                If !OrMainMet = "Y" Then
                  wGldAs = moCn.GetFldVal("Select PDesc225 from Param Where Ptyp='GLDAS' and PMCd='" + !OdGldAs + "'")
                  
                  If InStr(1, "," + wGldAs + ",", ",D,") <> 0 Then wGldAsWt = wGldAsWt + IIF(wGldAsWt <> "", "+", "") + "(D63/5)"
                  If InStr(1, "," + wGldAs + ",", ",C,") <> 0 Then wGldAsWt = wGldAsWt + IIF(wGldAsWt <> "", "+", "") + "(D64/5)"
                  If InStr(1, "," + wGldAs + ",", ",X,") <> 0 Then wGldAsWt = wGldAsWt + IIF(wGldAsWt <> "", "+", "") + "(D65)"
                  
                End If
                wGldAsWt = "(" + wGldAsWt + ")"
                                        
                oSheet.Cells(wOrMRow, 10).formula = "=$I" + CStr(wOrMRow) + "*(IF($G" + CStr(wOrMRow) + "=""Q"",$E" + CStr(wOrMRow) + ",(IF($G" + CStr(wOrMRow) + "=""W""," + wGldAsWt + ",0))))"
                oSheet.Cells(wOrMRow, 10).Font.Color = vbBlack
              ElseIf wOrMRow = 22 Then
                oSheet.Cells(16, 1).AddComment "#Metal Details InComplete"
              End If
              wOrMRow = wOrMRow + 1
              
              
            Case Is = "D", "C"
              If wOrDCRow <= 40 Then
                oSheet.Cells(wOrDCRow, 1) = !OrSrNo
                oSheet.Cells(wOrDCRow, 2) = !RmCtg
                oSheet.Cells(wOrDCRow, 3) = !OrRmSCtg
                oSheet.Cells(wOrDCRow, 4) = !OrRmCd
                oSheet.Cells(wOrDCRow, 5) = !OrLn1
                oSheet.Cells(wOrDCRow, 6) = !OrRmPtr
                oSheet.Cells(wOrDCRow, 7) = !OrQty
                oSheet.Cells(wOrDCRow, 8) = IIF(adc("UoYn1") = "Y", MWLib.Div(!OrWt, !qOrdQty), !OrWt)
                oSheet.Cells(wOrDCRow, 9) = !RmQw
                oSheet.Cells(wOrDCRow, 10) = !OrSalRt
                oSheet.Cells(wOrDCRow, 11).Font.Color = vbBlack
                oSheet.Cells(wOrDCRow, 12) = !OrSetSCd
                oSheet.Cells(wOrDCRow, 13) = !OrSetSalRt
                oSheet.Cells(wOrDCRow, 14).Font.Color = vbBlack
              ElseIf wOrDCRow = 41 Then
                oSheet.Cells(24, 1).AddComment "#Studding Details Incomplete"
              End If
              wOrDCRow = wOrDCRow + 1
              
            Case Is = "X"
              If wOrXRow <= 48 Then
                oSheet.Cells(wOrXRow, 1) = !OrSrNo
                oSheet.Cells(wOrXRow, 2) = !RmCtg
                oSheet.Cells(wOrXRow, 3) = !OrRmSCtg
                oSheet.Cells(wOrXRow, 4) = !OrRmCd
                oSheet.Cells(wOrXRow, 5) = !OrQty
                oSheet.Cells(wOrXRow, 6) = IIF(adc("UoYn1") = "Y", MWLib.Div(!OrWt, !qOrdQty), !OrWt)
                oSheet.Cells(wOrXRow, 7) = !RmQw
                oSheet.Cells(wOrXRow, 8) = !OrSalRt
                oSheet.Cells(wOrXRow, 9).Font.Color = vbBlack
              ElseIf wOrXRow = 49 Then
                oSheet.Cells(43, 1).AddComment = "#Acc Details InComplete"
              End If
              wOrXRow = wOrXRow + 1
            End Select
          
          ElseIf !qTag = "2LAB" Then
            If wOlRow <= 57 Then
              oSheet.Cells(wOlRow, 1) = !OlSrNo
              oSheet.Cells(wOlRow, 2) = !OlMCd
              oSheet.Cells(wOlRow, 3) = !OlSCd
              oSheet.Cells(wOlRow, 4) = !OlQw
              oSheet.Cells(wOlRow, 5) = !OlQty
              oSheet.Cells(wOlRow, 6) = !OlSalRt
              ' ***** Manali 3.6.1 - Lab Val based on min/max val or OlVal
              oSheet.Cells(wOlRow, 7) = "=IF($F" + CStr(wOlRow) + " * (IF($D" + CStr(wOlRow) + "=""Q"",$E" + CStr(wOlRow) + ",(IF($D" + CStr(wOlRow) + "=""W"",$E$51,0))))=" + IIF(!OlQw = "Q", CStr(!OlQty), CStr(!OdLabAsWt)) + "*" + CStr(!OlSalRt) + "," + _
                        CStr(!OlSalVal) + ", $F" + CStr(wOlRow) + " * (IF($D" + CStr(wOlRow) + "=""Q"",$E" + CStr(wOlRow) + ",(IF($D" + CStr(wOlRow) + "=""W"",$E$51,0)))))"
              oSheet.Cells(wOlRow, 7).Font.Color = vbBlack
            ElseIf wOlRow = 58 Then
              oSheet.Cells(51, 1).AddComment = "#Labour Details InComplete"
            End If
            wOlRow = wOlRow + 1
          End If
        
        wOdSr = !OdSr
        .MoveNext
      
      Else: Exit Do
      End If
    Loop
  End With
  Set oSheet = oBook.worksheets(1)
  oSheet.Rows(CStr(wOdRow) + ":299").Hidden = True
  
  oExcel.worksheets(2).Delete

  Set pv_rsOrd = Nothing
  
  Cd1.DialogTitle = "Save Travel Excel File As"
  Cd1.FileName = adc("UoInTcFr") + "-" + adc("UoInYyFr") + "-" + adc("UoInChrFr") + "-" + CStr(adc("UoInNoFr")) + ".xls"
  Cd1.Filter = "Excel Files (*.xls) |*.xls"
  
  Cd1.Flags = cdlOFNOverwritePrompt Or cdlOFNLongNames Or cdlOFNPathMustExist Or cdlOFNExtensionDifferent
  
  Cd1.CancelError = True
  Cd1.ShowSave

  If Len(Cd1.FileName) > 4 Then
    If UCase(Right(Cd1.FileName, 4)) = UCase(".xls") Then
      oExcel.ActiveWorkbook.SaveAs Cd1.FileName
    Else
      wErrMsg = "File Extension Is Wrong": Exit Sub
    End If
  End If
  
  If oExcel.ActiveWorkbook.Path + "\" + oExcel.ActiveWorkbook.Name <> sFilePath Then
    Kill sFilePath
  End If
  
  DispMsg "Excel Generated Successfully !!", etInfo
  oExcel.ActiveWorkbook.Close True
  oExcel.QUIT
  
  oExcel.DisplayAlerts = True
  Set oExcel = Nothing
  Set oBook = Nothing
  Set oSheet = Nothing
Exit Sub
Err_Hndlr:
  If Err.Number = 32755 Then
    wErrMsg = "Operation Cancelled"
    oExcel.ActiveWorkbook.Close
    oExcel.QUIT
  Else
    wErrMsg = Err.Description
  End If
  If Dir(sFilePath) <> "" Then Kill sFilePath
  Err.Clear
  
End Sub





