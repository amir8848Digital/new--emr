VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{AB2CC32A-E652-42A2-B341-26ECE02F6386}#2.1#0"; "mwfctlforspeed9.ocx"
Begin VB.Form EmrFrpSalAnaDsg 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sales Analysis (Design Level)"
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
   Begin MwfCtl.MWCTL_BTN1 CmdSc 
      Height          =   390
      Left            =   2160
      TabIndex        =   210
      ToolTipText     =   "Analysis Scope Options"
      Top             =   9720
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   688
      BackColor       =   -2147483633
      MaskColor       =   8454143
      Caption         =   "Det S&cope"
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
      Left            =   11640
      TabIndex        =   123
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
      TabIndex        =   121
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
      Left            =   5745
      TabIndex        =   127
      Top             =   9630
      Width           =   3555
      Begin MwfCtl.MWCTL_MED aktxt 
         Height          =   345
         Index           =   0
         Left            =   210
         TabIndex        =   128
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
         TabIndex        =   129
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
      Left            =   6960
      TabIndex        =   126
      Top             =   9645
      Visible         =   0   'False
      Width           =   855
      Begin CRVIEWER9LibCtl.CRViewer9 CRV_REP 
         Height          =   495
         Left            =   480
         TabIndex        =   125
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
      Left            =   0
      TabIndex        =   122
      Top             =   9600
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
      TabIndex        =   124
      Top             =   60
      Width           =   15255
      Begin TabDlg.SSTab TAB_REP 
         Height          =   9465
         Left            =   15
         TabIndex        =   3
         Top             =   60
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   16695
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "&Data Range"
         TabPicture(0)   =   "EmrFrpSalAnaDsg.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "fra_tabrep(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "&Group Sort"
         TabPicture(1)   =   "EmrFrpSalAnaDsg.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "fra_tabrep(1)"
         Tab(1).ControlCount=   1
         Begin VB.Frame fra_tabrep 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   9070
            Index           =   1
            Left            =   -74940
            TabIndex        =   98
            Top             =   360
            Width           =   15105
            Begin VB.Frame FraGrph 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   7920
               Left            =   150
               TabIndex        =   163
               Top             =   1050
               Width           =   11160
               Begin VB.Frame FraGrphDet 
                  BorderStyle     =   0  'None
                  Enabled         =   0   'False
                  Height          =   7425
                  Left            =   30
                  TabIndex        =   164
                  Top             =   450
                  Width           =   11085
                  Begin VB.ComboBox cboGrphValue 
                     Height          =   315
                     ItemData        =   "EmrFrpSalAnaDsg.frx":0038
                     Left            =   2100
                     List            =   "EmrFrpSalAnaDsg.frx":004E
                     Sorted          =   -1  'True
                     Style           =   2  'Dropdown List
                     TabIndex        =   105
                     ToolTipText     =   "Select Graph Value"
                     Top             =   915
                     Width           =   2865
                  End
                  Begin TabDlg.SSTab TAB_GRPH 
                     Height          =   3015
                     Left            =   0
                     TabIndex        =   109
                     Top             =   1650
                     Width           =   11115
                     _ExtentX        =   19606
                     _ExtentY        =   5318
                     _Version        =   393216
                     Tabs            =   2
                     TabsPerRow      =   2
                     TabHeight       =   520
                     TabCaption(0)   =   "T&itles"
                     TabPicture(0)   =   "EmrFrpSalAnaDsg.frx":0081
                     Tab(0).ControlEnabled=   -1  'True
                     Tab(0).Control(0)=   "FraTabGrph(0)"
                     Tab(0).Control(0).Enabled=   0   'False
                     Tab(0).ControlCount=   1
                     TabCaption(1)   =   "Top &N"
                     TabPicture(1)   =   "EmrFrpSalAnaDsg.frx":009D
                     Tab(1).ControlEnabled=   0   'False
                     Tab(1).Control(0)=   "FraTabGrph(1)"
                     Tab(1).ControlCount=   1
                     Begin VB.Frame FraTabGrph 
                        BorderStyle     =   0  'None
                        Height          =   2630
                        Index           =   0
                        Left            =   60
                        TabIndex        =   173
                        Top             =   360
                        Width           =   11030
                        Begin MwfCtl.MWCTL_MED atxt 
                           Height          =   285
                           Index           =   58
                           Left            =   1680
                           TabIndex        =   110
                           ToolTipText     =   "Enter Graph Title"
                           Top             =   435
                           Width           =   6975
                           _ExtentX        =   12303
                           _ExtentY        =   503
                           MaxLength       =   30
                           DataField       =   "UOGRPHMTITLE"
                           IdName          =   "UOGRPHMTITLE"
                        End
                        Begin MwfCtl.MWCTL_MED atxt 
                           Height          =   285
                           Index           =   59
                           Left            =   1680
                           TabIndex        =   111
                           ToolTipText     =   "Enter Graph Sub Ttile"
                           Top             =   720
                           Width           =   6975
                           _ExtentX        =   12303
                           _ExtentY        =   503
                           MaxLength       =   30
                           DataField       =   "UOGRPHSTITLE"
                           IdName          =   "UOGRPHSTITLE"
                        End
                        Begin MwfCtl.MWCTL_MED atxt 
                           Height          =   285
                           Index           =   61
                           Left            =   1680
                           TabIndex        =   113
                           ToolTipText     =   "Enter Group Title"
                           Top             =   1290
                           Width           =   6975
                           _ExtentX        =   12303
                           _ExtentY        =   503
                           MaxLength       =   30
                           DataField       =   "UOGRPHGTITLE"
                           IdName          =   "UOGRPHGTITLE"
                        End
                        Begin MwfCtl.MWCTL_MED atxt 
                           Height          =   285
                           Index           =   62
                           Left            =   1680
                           TabIndex        =   114
                           ToolTipText     =   "Enter Data Title"
                           Top             =   1575
                           Width           =   6975
                           _ExtentX        =   12303
                           _ExtentY        =   503
                           MaxLength       =   30
                           DataField       =   "UOGRPHDTITLE"
                           IdName          =   "UOGRPHDTITLE"
                        End
                        Begin MwfCtl.MWCTL_MED atxt 
                           Height          =   285
                           Index           =   63
                           Left            =   1680
                           TabIndex        =   115
                           ToolTipText     =   "Enter Data2 Title"
                           Top             =   1860
                           Width           =   6975
                           _ExtentX        =   12303
                           _ExtentY        =   503
                           MaxLength       =   30
                           DataField       =   "UOGRPHDTITLE2"
                           IdName          =   "UOGRPHDTITLE2"
                        End
                        Begin MwfCtl.MWCTL_MED atxt 
                           Height          =   285
                           Index           =   60
                           Left            =   1680
                           TabIndex        =   112
                           ToolTipText     =   "Enter Graph Foot Note"
                           Top             =   1005
                           Width           =   6975
                           _ExtentX        =   12303
                           _ExtentY        =   503
                           MaxLength       =   30
                           DataField       =   "UOGRPHFTNOTE"
                           IdName          =   "UOGRPHFTNOTE"
                        End
                        Begin VB.Label LblGrph 
                           AutoSize        =   -1  'True
                           BackStyle       =   0  'Transparent
                           Caption         =   "Foot Note"
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
                           Left            =   330
                           TabIndex        =   179
                           Top             =   1005
                           Width           =   1035
                        End
                        Begin VB.Label LblGrph 
                           AutoSize        =   -1  'True
                           BackStyle       =   0  'Transparent
                           Caption         =   "Data2 Title"
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
                           Left            =   330
                           TabIndex        =   178
                           Top             =   1860
                           Width           =   1155
                        End
                        Begin VB.Label LblGrph 
                           AutoSize        =   -1  'True
                           BackStyle       =   0  'Transparent
                           Caption         =   "Data Title"
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
                           Left            =   330
                           TabIndex        =   177
                           Top             =   1575
                           Width           =   1035
                        End
                        Begin VB.Label LblGrph 
                           AutoSize        =   -1  'True
                           BackStyle       =   0  'Transparent
                           Caption         =   "Group Title"
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
                           Left            =   330
                           TabIndex        =   176
                           Top             =   1290
                           Width           =   1170
                        End
                        Begin VB.Label LblGrph 
                           AutoSize        =   -1  'True
                           BackStyle       =   0  'Transparent
                           Caption         =   "Sub Title"
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
                           Left            =   330
                           TabIndex        =   175
                           Top             =   720
                           Width           =   945
                        End
                        Begin VB.Label LblGrph 
                           AutoSize        =   -1  'True
                           BackStyle       =   0  'Transparent
                           Caption         =   "Graph Title"
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
                           Left            =   330
                           TabIndex        =   174
                           Top             =   435
                           Width           =   1170
                        End
                     End
                     Begin VB.Frame FraTabGrph 
                        BorderStyle     =   0  'None
                        Height          =   2630
                        Index           =   1
                        Left            =   -74940
                        TabIndex        =   165
                        Top             =   360
                        Width           =   11030
                        Begin MwfCtl.MWCTL_MED atxt 
                           Height          =   285
                           Index           =   64
                           Left            =   2310
                           TabIndex        =   117
                           ToolTipText     =   "Enter Sorting Type for Group1"
                           Top             =   1110
                           Width           =   2115
                           _ExtentX        =   3731
                           _ExtentY        =   503
                           MaxLength       =   8
                           DataType        =   4
                           DataField       =   "UOGRPHG1TOPN"
                           IdName          =   "UOGRPHG1TOPN"
                        End
                        Begin MwfCtl.MWCTL_MED atxt 
                           Height          =   285
                           Index           =   65
                           Left            =   5580
                           TabIndex        =   118
                           ToolTipText     =   "Enter Top N Value For Group1"
                           Top             =   1110
                           Width           =   945
                           _ExtentX        =   1667
                           _ExtentY        =   503
                           Mask            =   "####0"
                           MaxLength       =   5
                           DataType        =   1
                           DataField       =   "UOGRPHG1NVAL"
                           IdName          =   "UOGRPHG1NVAL"
                        End
                        Begin MwfCtl.MWCTL_MED atxt 
                           Height          =   285
                           Index           =   69
                           Left            =   1230
                           TabIndex        =   116
                           ToolTipText     =   "Enter Top N Based On Summary"
                           Top             =   480
                           Width           =   2115
                           _ExtentX        =   3731
                           _ExtentY        =   503
                           MaxLength       =   8
                           DataType        =   4
                           DataField       =   "UOGRPHSUMM"
                           IdName          =   "UOGRPHSUMM"
                        End
                        Begin MwfCtl.MWCTL_MED atxt 
                           Height          =   285
                           Index           =   66
                           Left            =   2310
                           TabIndex        =   119
                           ToolTipText     =   "Enter Sorting Type for Group2"
                           Top             =   1740
                           Width           =   2115
                           _ExtentX        =   3731
                           _ExtentY        =   503
                           MaxLength       =   8
                           DataType        =   4
                           DataField       =   "UOGRPHG2TOPN"
                           IdName          =   "UOGRPHG2TOPN"
                        End
                        Begin MwfCtl.MWCTL_MED atxt 
                           Height          =   285
                           Index           =   67
                           Left            =   5580
                           TabIndex        =   120
                           ToolTipText     =   "Enter Top N Value For Group2"
                           Top             =   1740
                           Width           =   945
                           _ExtentX        =   1667
                           _ExtentY        =   503
                           Mask            =   "####0"
                           MaxLength       =   5
                           DataType        =   1
                           DataField       =   "UOGRPHG2NVAL"
                           IdName          =   "UOGRPHG2NVAL"
                        End
                        Begin VB.Label LblGrphG1Val 
                           AutoSize        =   -1  'True
                           BackStyle       =   0  'Transparent
                           Caption         =   "Value"
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
                           Left            =   4920
                           TabIndex        =   172
                           Top             =   1125
                           Width           =   615
                        End
                        Begin VB.Label LblGrph 
                           AutoSize        =   -1  'True
                           BackStyle       =   0  'Transparent
                           Caption         =   "Based On"
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
                           Left            =   150
                           TabIndex        =   171
                           Top             =   495
                           Width           =   1035
                        End
                        Begin VB.Label LblGrphG1Srt 
                           AutoSize        =   -1  'True
                           BackStyle       =   0  'Transparent
                           Caption         =   "Sort Type"
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
                           Left            =   1230
                           TabIndex        =   170
                           Top             =   1125
                           Width           =   1035
                        End
                        Begin VB.Label LblGrphGrp1 
                           AutoSize        =   -1  'True
                           BackStyle       =   0  'Transparent
                           Caption         =   "Group 1"
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
                           Left            =   150
                           TabIndex        =   169
                           Top             =   1125
                           Width           =   825
                        End
                        Begin VB.Label LblGrphGrp2 
                           AutoSize        =   -1  'True
                           BackStyle       =   0  'Transparent
                           Caption         =   "Group 2"
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
                           Left            =   150
                           TabIndex        =   168
                           Top             =   1770
                           Width           =   825
                        End
                        Begin VB.Label LblGrphG2Srt 
                           AutoSize        =   -1  'True
                           BackStyle       =   0  'Transparent
                           Caption         =   "Sort Type"
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
                           Left            =   1230
                           TabIndex        =   167
                           Top             =   1755
                           Width           =   1035
                        End
                        Begin VB.Label LblGrphG2Val 
                           AutoSize        =   -1  'True
                           BackStyle       =   0  'Transparent
                           Caption         =   "Value"
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
                           Left            =   4920
                           TabIndex        =   166
                           Top             =   1755
                           Width           =   615
                        End
                     End
                  End
                  Begin MwfCtl.MWCTL_MED atxt 
                     Height          =   285
                     Index           =   56
                     Left            =   2100
                     TabIndex        =   101
                     ToolTipText     =   "Enter Graph Type"
                     Top             =   345
                     Width           =   1995
                     _ExtentX        =   3519
                     _ExtentY        =   503
                     MaxLength       =   8
                     DataType        =   4
                     DataField       =   "UOGRPHTYP"
                     IdName          =   "UOGRPHTYP"
                  End
                  Begin MwfCtl.MWCTL_MED atxt 
                     Height          =   285
                     Index           =   57
                     Left            =   7620
                     TabIndex        =   102
                     ToolTipText     =   "Enter Graph irection"
                     Top             =   345
                     Width           =   1995
                     _ExtentX        =   3519
                     _ExtentY        =   503
                     MaxLength       =   8
                     DataType        =   4
                     DataField       =   "UOGRPHDIR"
                     IdName          =   "UOGRPHDIR"
                  End
                  Begin MwfCtl.MWCTL_MED atxt 
                     Height          =   285
                     Index           =   68
                     Left            =   2100
                     TabIndex        =   103
                     ToolTipText     =   "Enter Data Label"
                     Top             =   630
                     Width           =   1995
                     _ExtentX        =   3519
                     _ExtentY        =   503
                     MaxLength       =   8
                     DataType        =   4
                     DataField       =   "UOGRPHDATALB"
                     IdName          =   "UOGRPHDATALB"
                  End
                  Begin MwfCtl.MWCTL_MED atxt 
                     Height          =   285
                     Index           =   70
                     Left            =   7620
                     TabIndex        =   104
                     ToolTipText     =   "Enter Graph View Angle"
                     Top             =   630
                     Width           =   1995
                     _ExtentX        =   3519
                     _ExtentY        =   503
                     MaxLength       =   8
                     DataType        =   4
                     DataField       =   "UOGRPHVWANGL"
                     IdName          =   "UOGRPHVWANGL"
                  End
                  Begin MwfCtl.MWCTL_MED atxt 
                     Height          =   315
                     Index           =   71
                     Left            =   4950
                     TabIndex        =   180
                     ToolTipText     =   "Enter From Customer"
                     Top             =   915
                     Visible         =   0   'False
                     Width           =   405
                     _ExtentX        =   714
                     _ExtentY        =   556
                     Mask            =   "#0"
                     MaxLength       =   2
                     DataType        =   1
                     DataField       =   "UOGRPHVALIND"
                     IdName          =   "UOGRPHVALIND"
                  End
                  Begin MwfCtl.MWCTL_MED atxt 
                     Height          =   285
                     Index           =   53
                     Left            =   7620
                     TabIndex        =   106
                     ToolTipText     =   "Enter Graph in Color (Y / N) And Blank For Default"
                     Top             =   915
                     Width           =   315
                     _ExtentX        =   556
                     _ExtentY        =   503
                     DataType        =   4
                     DataField       =   "UOGRPHCLR"
                     IdName          =   "UOGRPHCLR"
                  End
                  Begin MwfCtl.MWCTL_MED atxt 
                     Height          =   285
                     Index           =   54
                     Left            =   7620
                     TabIndex        =   108
                     ToolTipText     =   "Enter Graph View Angle"
                     Top             =   1200
                     Width           =   1995
                     _ExtentX        =   3519
                     _ExtentY        =   503
                     MaxLength       =   8
                     DataType        =   4
                     DataField       =   "UOGRPHORIENT"
                     IdName          =   "UOGRPHORIENT"
                  End
                  Begin MwfCtl.MWCTL_MED atxt 
                     Height          =   285
                     Index           =   72
                     Left            =   2100
                     TabIndex        =   107
                     ToolTipText     =   "Enter Data Label"
                     Top             =   1210
                     Width           =   1995
                     _ExtentX        =   3519
                     _ExtentY        =   503
                     MaxLength       =   8
                     DataType        =   4
                     DataField       =   "UOGRPHPSZ"
                     IdName          =   "UOGRPHPSZ"
                  End
                  Begin VB.Label LblGrph 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "Value"
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
                     Index           =   10
                     Left            =   300
                     TabIndex        =   189
                     Top             =   915
                     Width           =   615
                  End
                  Begin VB.Label LblGrph 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "Graph Type"
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
                     Left            =   300
                     TabIndex        =   188
                     Top             =   345
                     Width           =   1245
                  End
                  Begin VB.Label LblGrph 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "Data Label"
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
                     Left            =   300
                     TabIndex        =   187
                     Top             =   630
                     Width           =   1155
                  End
                  Begin VB.Label LblGrph 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "Graph Direction"
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
                     TabIndex        =   186
                     Top             =   345
                     Width           =   1635
                  End
                  Begin VB.Label LblGrph 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "View Angle"
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
                     Left            =   5760
                     TabIndex        =   185
                     Top             =   630
                     Width           =   1170
                  End
                  Begin VB.Label LblGrph 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "Graph in Color"
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
                     Left            =   5760
                     TabIndex        =   184
                     Top             =   915
                     Width           =   1500
                  End
                  Begin VB.Label LblGrph 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "Paper Orientation"
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
                     Index           =   19
                     Left            =   5760
                     TabIndex        =   183
                     Top             =   1200
                     Width           =   1830
                  End
                  Begin VB.Label LblGrph 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "Paper Size"
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
                     Index           =   20
                     Left            =   300
                     TabIndex        =   182
                     Top             =   1210
                     Width           =   1155
                  End
                  Begin VB.Label LblMain 
                     BackStyle       =   0  'Transparent
                     BorderStyle     =   1  'Fixed Single
                     Caption         =   "  Graph Details"
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
                     TabIndex        =   181
                     Top             =   0
                     Width           =   11205
                  End
               End
               Begin VB.CheckBox ChkGrph 
                  Caption         =   "Show Graph"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   12
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Left            =   120
                  MaskColor       =   &H8000000F&
                  TabIndex        =   100
                  Top             =   0
                  UseMaskColor    =   -1  'True
                  Width           =   1875
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   74
                  Left            =   2010
                  TabIndex        =   190
                  ToolTipText     =   "Enter From Customer"
                  Top             =   0
                  Visible         =   0   'False
                  Width           =   315
                  _ExtentX        =   556
                  _ExtentY        =   503
                  MaxLength       =   8
                  DataType        =   4
                  DataField       =   "UOSHOWGRPH"
                  IdName          =   "UOSHOWGRPH"
               End
            End
            Begin MwfCtl.MWCTL_GRP GRP_REP 
               Height          =   360
               Left            =   240
               TabIndex        =   99
               Top             =   360
               Width           =   5595
               _ExtentX        =   9869
               _ExtentY        =   635
            End
         End
         Begin VB.Frame fra_tabrep 
            BorderStyle     =   0  'None
            Height          =   9070
            Index           =   0
            Left            =   120
            TabIndex        =   4
            Top             =   320
            Width           =   14940
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Include Sec SalCtg"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   460
               Index           =   1
               Left            =   60
               TabIndex        =   41
               Tag             =   "DmHold"
               ToolTipText     =   "Check To Include SalCtg2 and SalCtg3"
               Top             =   3960
               Width           =   2115
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Best Seller Option "
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
               TabIndex        =   81
               Tag             =   "ShowPic"
               ToolTipText     =   "Check to use Best Seller Option"
               Top             =   8760
               Width           =   2130
            End
            Begin VB.Frame FraBestSeller 
               Caption         =   "Best Seller Options"
               Enabled         =   0   'False
               Height          =   1575
               Left            =   9960
               TabIndex        =   199
               Top             =   7440
               Width           =   4935
               Begin VB.Frame Frame2 
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   585
                  Left            =   165
                  TabIndex        =   200
                  Top             =   930
                  Width           =   4635
                  Begin MwfCtl.MWCTL_MED atxt 
                     Height          =   285
                     Index           =   223
                     Left            =   2340
                     TabIndex        =   96
                     ToolTipText     =   "Enter Best Seller YY (Both YY And Qty/Val Has to be Entered For Best Seller Design)"
                     Top             =   210
                     Width           =   345
                     _ExtentX        =   609
                     _ExtentY        =   503
                     MaxLength       =   2
                     DataType        =   4
                     IdName          =   "WBSYY"
                  End
                  Begin MwfCtl.MWCTL_MED atxt 
                     Height          =   285
                     Index           =   224
                     Left            =   4140
                     TabIndex        =   97
                     ToolTipText     =   "Enter Whether Best Seller Is By Quantity Or Value"
                     Top             =   210
                     Width           =   285
                     _ExtentX        =   503
                     _ExtentY        =   503
                     DataType        =   4
                     IdName          =   "WBSQV"
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Best Seller Design"
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
                     Index           =   188
                     Left            =   60
                     TabIndex        =   203
                     Top             =   210
                     Width           =   2055
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Qty(Y)/Val(N)"
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
                     Index           =   187
                     Left            =   2790
                     TabIndex        =   202
                     Top             =   240
                     Width           =   1305
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "YY."
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
                     Index           =   186
                     Left            =   1890
                     TabIndex        =   201
                     Top             =   240
                     Width           =   375
                  End
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   219
                  Left            =   1065
                  TabIndex        =   92
                  ToolTipText     =   "Enter From Order Quantity"
                  Top             =   285
                  Width           =   930
                  _ExtentX        =   1640
                  _ExtentY        =   503
                  Alignment       =   2
                  Mask            =   "#####0"
                  MaxLength       =   6
                  DataType        =   1
                  IdName          =   "WEXPQTYFR"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   220
                  Left            =   2730
                  TabIndex        =   93
                  ToolTipText     =   "Enter To Order Quantity"
                  Top             =   285
                  Width           =   930
                  _ExtentX        =   1640
                  _ExtentY        =   503
                  Alignment       =   2
                  Mask            =   "#####0"
                  MaxLength       =   6
                  DataType        =   1
                  IdName          =   "WEXPQTYTO"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   221
                  Left            =   1065
                  TabIndex        =   94
                  ToolTipText     =   "Enter From Gross Weight"
                  Top             =   570
                  Width           =   1380
                  _ExtentX        =   2434
                  _ExtentY        =   503
                  Alignment       =   2
                  Mask            =   "#######0.00"
                  MaxLength       =   11
                  DataType        =   2
                  IdName          =   "WINVFOBFR"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   222
                  Left            =   2730
                  TabIndex        =   95
                  ToolTipText     =   "Enter To Gross Weight"
                  Top             =   570
                  Width           =   1380
                  _ExtentX        =   2434
                  _ExtentY        =   503
                  Alignment       =   2
                  Mask            =   "#######0.00"
                  MaxLength       =   11
                  DataType        =   2
                  IdName          =   "WINVFOBTO"
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Inv Val"
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
                  Index           =   182
                  Left            =   120
                  TabIndex        =   205
                  Top             =   570
                  Width           =   1230
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Inv. Qty"
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
                  Index           =   181
                  Left            =   120
                  TabIndex        =   204
                  Top             =   285
                  Width           =   1230
               End
            End
            Begin VB.Frame FraOptDt 
               BorderStyle     =   0  'None
               Height          =   285
               Left            =   6675
               TabIndex        =   196
               ToolTipText     =   "Select if to show Invoice Date Or Invoice Printing Date"
               Top             =   8475
               Width           =   3180
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
                  Left            =   0
                  TabIndex        =   86
                  ToolTipText     =   "Select if to show Invoice Date Or Invoice Printing Date"
                  Top             =   0
                  Width           =   1635
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
                  Left            =   1680
                  TabIndex        =   87
                  ToolTipText     =   "Select if to show Invoice Date Or Invoice Printing Date"
                  Top             =   0
                  Width           =   1365
               End
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Show Picture          "
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
               TabIndex        =   80
               Tag             =   "ShowPic"
               ToolTipText     =   "Check To Show Picture"
               Top             =   8460
               Width           =   2130
            End
            Begin VB.Frame FraOptCust 
               BorderStyle     =   0  'None
               Height          =   285
               Left            =   6675
               TabIndex        =   193
               ToolTipText     =   "Select eihter Invoice Customer Or Order Customer"
               Top             =   8190
               Width           =   3225
               Begin VB.OptionButton OptCust 
                  Caption         =   "Invoice Cust"
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
                  TabIndex        =   84
                  Top             =   0
                  Width           =   1545
               End
               Begin VB.OptionButton OptCust 
                  Caption         =   "Order Cust"
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
                  Left            =   1680
                  TabIndex        =   85
                  Top             =   0
                  Width           =   1425
               End
            End
            Begin VB.Frame FraOptValWt 
               BorderStyle     =   0  'None
               Height          =   285
               Left            =   6675
               TabIndex        =   192
               ToolTipText     =   "Select if to show Value or Weight"
               Top             =   7905
               Width           =   3180
               Begin VB.OptionButton OptValWt 
                  Caption         =   "Weight"
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
                  Left            =   1680
                  TabIndex        =   83
                  Top             =   0
                  Width           =   1425
               End
               Begin VB.OptionButton OptValWt 
                  Caption         =   "Value"
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
                  TabIndex        =   82
                  Top             =   0
                  Width           =   1425
               End
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   18
               Left            =   5780
               TabIndex        =   9
               ToolTipText     =   "Enter To Invoice Year"
               Top             =   795
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
               Left            =   2760
               TabIndex        =   6
               ToolTipText     =   "Enter From Invoice Year"
               Top             =   795
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
               Left            =   1995
               TabIndex        =   5
               ToolTipText     =   "Enter From Invoice Tc"
               Top             =   795
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
               Left            =   3210
               TabIndex        =   7
               ToolTipText     =   "Enter From Invoice Character"
               Top             =   795
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
               Left            =   5000
               TabIndex        =   8
               ToolTipText     =   "Enter To Invoice Tc"
               Top             =   795
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
               Left            =   6200
               TabIndex        =   10
               ToolTipText     =   "Enter To Invoice Character"
               Top             =   795
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
               Left            =   1995
               TabIndex        =   11
               ToolTipText     =   "Enter From Invoice Number"
               Top             =   1080
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
               Left            =   5000
               TabIndex        =   13
               ToolTipText     =   "Enter To Invoice Number"
               Top             =   1080
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
               Index           =   3
               Left            =   1995
               TabIndex        =   78
               ToolTipText     =   "Specify Whether Value/ Weight To Be Shown Is Actual, Invoice Or Customs "
               Top             =   7860
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
               Left            =   9675
               TabIndex        =   89
               ToolTipText     =   "Show Value Or Weight Figures ('Y' for Value / 'N' For Weight)"
               Top             =   7905
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
               Index           =   28
               Left            =   5000
               TabIndex        =   36
               ToolTipText     =   "Enter To Design Category"
               Top             =   3390
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
               Index           =   0
               Left            =   7980
               TabIndex        =   37
               ToolTipText     =   "Enter Design Category Selection"
               Top             =   3390
               Width           =   6930
               _ExtentX        =   12224
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
               Left            =   1995
               TabIndex        =   35
               ToolTipText     =   "Enter From Design Category"
               Top             =   3390
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
               Left            =   4995
               TabIndex        =   43
               ToolTipText     =   "Enter To Design Code"
               Top             =   4440
               Width           =   2385
               _ExtentX        =   4207
               _ExtentY        =   503
               MaxLength       =   15
               DataField       =   "UODMCDTO"
               IdName          =   "UODMCDTO"
               CmpStr          =   "IdDmCd <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   32
               Left            =   7980
               TabIndex        =   44
               ToolTipText     =   "Enter Design Code Selection"
               Top             =   4440
               Width           =   6930
               _ExtentX        =   12224
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UODMCDSEL"
               IdName          =   "UODMCDSEL"
               CmpStr          =   "IdDmCd In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   33
               Left            =   1995
               TabIndex        =   42
               ToolTipText     =   "Enter From Design Code"
               Top             =   4440
               Width           =   2385
               _ExtentX        =   4207
               _ExtentY        =   503
               MaxLength       =   15
               DataField       =   "UODMCDFR"
               IdName          =   "UODMCDFR"
               CmpStr          =   "IdDmCd >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   40
               Left            =   5000
               TabIndex        =   18
               ToolTipText     =   "Enter To Customer"
               Top             =   1650
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               DataField       =   "UOCMCDTO"
               IdName          =   "UOCMCDTO"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   41
               Left            =   7980
               TabIndex        =   19
               ToolTipText     =   "Enter Customer Selection"
               Top             =   1650
               Width           =   6930
               _ExtentX        =   12224
               _ExtentY        =   503
               MaxLength       =   255
               DataType        =   4
               DataField       =   "UOCMCDSEL"
               IdName          =   "UOCMCDSEL"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   42
               Left            =   1995
               TabIndex        =   17
               ToolTipText     =   "Enter From Customer"
               Top             =   1650
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               DataField       =   "UOCMCDFR"
               IdName          =   "UOCMCDFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   15
               Left            =   4995
               TabIndex        =   50
               ToolTipText     =   "Enter To Karat"
               Top             =   5295
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataField       =   "UOKTTO"
               IdName          =   "UOKTTO"
               CmpStr          =   "IdKt <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   16
               Left            =   7980
               TabIndex        =   51
               ToolTipText     =   "Enter Karat Selection"
               Top             =   5295
               Width           =   6930
               _ExtentX        =   12224
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOKTSEL"
               IdName          =   "UOKTSEL"
               CmpStr          =   "IdKt In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   2
               Left            =   1995
               TabIndex        =   49
               ToolTipText     =   "Enter From Karat"
               Top             =   5295
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataField       =   "UOKTFR"
               IdName          =   "UOKTFR"
               CmpStr          =   "IdKt >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   11
               Left            =   5000
               TabIndex        =   39
               ToolTipText     =   "Enter To Design Sales Category"
               Top             =   3675
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataType        =   4
               DataField       =   "UODMSALCTGTO"
               IdName          =   "UODMSALCTGTO"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   12
               Left            =   7980
               TabIndex        =   40
               ToolTipText     =   "Enter Design Sales Category Selection"
               Top             =   3675
               Width           =   6930
               _ExtentX        =   12224
               _ExtentY        =   503
               MaxLength       =   255
               DataType        =   4
               DataField       =   "UODMSALCTGSEL"
               IdName          =   "UODMSALCTGSEL"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   14
               Left            =   1995
               TabIndex        =   38
               ToolTipText     =   "Enter From Design Sales Category"
               Top             =   3675
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataType        =   4
               DataField       =   "UODMSALCTGFR"
               IdName          =   "UODMSALCTGFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   7
               Left            =   5000
               TabIndex        =   33
               ToolTipText     =   "Enter To Export Number"
               Top             =   3105
               Width           =   2625
               _ExtentX        =   4630
               _ExtentY        =   503
               MaxLength       =   20
               DataField       =   "UOINEXPNOTO"
               IdName          =   "UOINEXPNOTO"
               CmpStr          =   "InExpNo <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   8
               Left            =   7980
               TabIndex        =   34
               ToolTipText     =   "Enter Export Number Selection"
               Top             =   3105
               Width           =   6930
               _ExtentX        =   12224
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOINEXPNOSEL"
               IdName          =   "UOINEXPNOSEL"
               CmpStr          =   "InExpNo In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   9
               Left            =   1995
               TabIndex        =   32
               ToolTipText     =   "Enter From Export Number"
               Top             =   3105
               Width           =   2625
               _ExtentX        =   4630
               _ExtentY        =   503
               MaxLength       =   20
               DataField       =   "UOINEXPNOFR"
               IdName          =   "UOINEXPNOFR"
               CmpStr          =   "InExpNo >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   34
               Left            =   4995
               TabIndex        =   46
               ToolTipText     =   "Enter To Design Size"
               Top             =   4725
               Width           =   585
               _ExtentX        =   1032
               _ExtentY        =   503
               MaxLength       =   5
               DataField       =   "UODMSZTO"
               IdName          =   "UODMSZTO"
               CmpStr          =   "IdDmSz <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   36
               Left            =   1995
               TabIndex        =   45
               ToolTipText     =   "Enter From Design Size"
               Top             =   4725
               Width           =   585
               _ExtentX        =   1032
               _ExtentY        =   503
               MaxLength       =   5
               DataField       =   "UODMSZFR"
               IdName          =   "UODMSZFR"
               CmpStr          =   "IdDmSz >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   37
               Left            =   4995
               TabIndex        =   48
               ToolTipText     =   "Enter To Suffix"
               Top             =   5010
               Width           =   1605
               _ExtentX        =   2831
               _ExtentY        =   503
               MaxLength       =   15
               DataField       =   "UOODSFXTO"
               IdName          =   "UOODSFXTO"
               CmpStr          =   "IdSfx <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   39
               Left            =   1995
               TabIndex        =   47
               ToolTipText     =   "Enter From Suffix"
               Top             =   5010
               Width           =   1605
               _ExtentX        =   2831
               _ExtentY        =   503
               MaxLength       =   15
               DataField       =   "UOODSFXFR"
               IdName          =   "UOODSFXFR"
               CmpStr          =   "IdSfx >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   21
               Left            =   5880
               TabIndex        =   65
               ToolTipText     =   "Enter To Order Serial"
               Top             =   6435
               Width           =   525
               _ExtentX        =   926
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOODSRTO"
               IdName          =   "UOODSRTO"
               CmpStr          =   "IdExpOdSr <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   23
               Left            =   2895
               TabIndex        =   63
               ToolTipText     =   "Enter From Order Serial"
               Top             =   6435
               Width           =   525
               _ExtentX        =   926
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOODSRFR"
               IdName          =   "UOODSRFR"
               CmpStr          =   "IdExpOdSr >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   19
               Left            =   1995
               TabIndex        =   62
               ToolTipText     =   "Enter From Order Number"
               Top             =   6435
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOOMNOFR"
               IdName          =   "UOOMNOFR"
               CmpStr          =   "IdExpOdNo >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   22
               Left            =   4995
               TabIndex        =   64
               ToolTipText     =   "Enter To Order Number"
               Top             =   6435
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOOMNOTO"
               IdName          =   "UOOMNOTO"
               CmpStr          =   "IdExpOdNo <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   24
               Left            =   9675
               TabIndex        =   90
               ToolTipText     =   "Show Invoice Or Order Customer ('Y' for Invoice / 'N' For Order)"
               Top             =   8190
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
               Index           =   25
               Left            =   4995
               TabIndex        =   75
               ToolTipText     =   "Enter To Gross Weight Per Piece"
               Top             =   7290
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   503
               Alignment       =   2
               MaxLength       =   0
               DataType        =   2
               DataField       =   "UOGRSWTTO"
               IdName          =   "UOGRSWTTO"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   27
               Left            =   1995
               TabIndex        =   74
               ToolTipText     =   "Enter From Gross Weight Per Piece"
               Top             =   7290
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   503
               Alignment       =   2
               MaxLength       =   0
               DataType        =   2
               DataField       =   "UOGRSWTFR"
               IdName          =   "UOGRSWTFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   35
               Left            =   4995
               TabIndex        =   77
               ToolTipText     =   "Enter To Dia/ColStn Weight Per Piece"
               Top             =   7575
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   503
               Alignment       =   2
               MaxLength       =   0
               DataType        =   2
               DataField       =   "UORMWTTO"
               IdName          =   "UORMWTTO"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   38
               Left            =   1995
               TabIndex        =   76
               ToolTipText     =   "Enter From Dia/ColStn Weight Per Piece"
               Top             =   7575
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   503
               Alignment       =   2
               MaxLength       =   0
               DataType        =   2
               DataField       =   "UORMWTFR"
               IdName          =   "UORMWTFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   44
               Left            =   5000
               TabIndex        =   16
               ToolTipText     =   "Enter To Invoice Voucher Date"
               Top             =   1365
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   3
               DataField       =   "UODTTO"
               IdName          =   "UODTTO"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   48
               Left            =   1995
               TabIndex        =   15
               ToolTipText     =   "Enter From Invoice Voucher Date"
               Top             =   1365
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   3
               DataField       =   "UODTFR"
               IdName          =   "UODTFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   49
               Left            =   1995
               TabIndex        =   0
               ToolTipText     =   "Enter Company Code"
               Top             =   540
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOCOCDFR"
               IdName          =   "UOCOCDFR"
               CmpStr          =   "IdCoCd >= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   50
               Left            =   7980
               TabIndex        =   28
               ToolTipText     =   "Enter Sales Executive Selection"
               Top             =   2535
               Width           =   6930
               _ExtentX        =   12224
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOSALPERSEL"
               IdName          =   "UOSALPERSEL"
               CmpStr          =   "OmCmSalPer In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   51
               Left            =   5000
               TabIndex        =   27
               ToolTipText     =   "Enter To Sales Executive  (Order Cust)"
               Top             =   2535
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
               Index           =   52
               Left            =   1995
               TabIndex        =   26
               ToolTipText     =   "Enter From  Sales Executive (Order Cust)"
               Top             =   2535
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UOSALPERFR"
               IdName          =   "UOSALPERFR"
               CmpStr          =   "OmCmSalPer  >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   10
               Left            =   3195
               TabIndex        =   58
               ToolTipText     =   "Enter From Order Character"
               Top             =   6150
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOOMCHRFR"
               IdName          =   "UOOMCHRFR"
               CmpStr          =   "IdExpOdChr >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   5
               Left            =   2775
               TabIndex        =   57
               ToolTipText     =   "Enter From Order Year"
               Top             =   6150
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               MaxLength       =   2
               DataField       =   "UOOMYYFR"
               IdName          =   "UOOMYYFR"
               CmpStr          =   "IdExpOdYy >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   17
               Left            =   6180
               TabIndex        =   61
               ToolTipText     =   "Enter To Order Character"
               Top             =   6150
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOOMCHRTO"
               IdName          =   "UOOMCHRTO"
               CmpStr          =   "IdExpOdChr <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   1
               Left            =   5760
               TabIndex        =   60
               ToolTipText     =   "Enter To Order Year"
               Top             =   6150
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               MaxLength       =   2
               DataField       =   "UOOMYYTO"
               IdName          =   "UOOMYYTO"
               CmpStr          =   "IdExpOdYy <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   13
               Left            =   4995
               TabIndex        =   59
               ToolTipText     =   "Enter To Order Tc"
               Top             =   6150
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataField       =   "UOOMTCTO"
               IdName          =   "UOOMTCTO"
               CmpStr          =   "IdExpOdTc <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   6
               Left            =   1995
               TabIndex        =   56
               ToolTipText     =   "Enter From Order Tc"
               Top             =   6150
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataField       =   "UOOMTCFR"
               IdName          =   "UOOMTCFR"
               CmpStr          =   "IdExpOdTc >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   55
               Left            =   1995
               TabIndex        =   79
               ToolTipText     =   "Enter Currency Code Whose Invoice You Wish To View Or Leave It Blank If You Wish To View All Invoices In Base Currency"
               Top             =   8175
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
               Index           =   73
               Left            =   5000
               TabIndex        =   1
               ToolTipText     =   "Enter To Company Code"
               Top             =   510
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOCOCDTO"
               IdName          =   "UOCOCDTO"
               CmpStr          =   "IdCoCd <= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   75
               Left            =   7980
               TabIndex        =   2
               ToolTipText     =   "Enter Design Code Selection"
               Top             =   540
               Width           =   6930
               _ExtentX        =   12224
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOCOCDSEL"
               IdName          =   "UOCOCDSEL"
               CmpStr          =   "IdCoCd In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   76
               Left            =   2505
               TabIndex        =   88
               ToolTipText     =   "Show Picture (Yes / No)"
               Top             =   8460
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
               Index           =   77
               Left            =   7980
               TabIndex        =   31
               ToolTipText     =   "Enter Customer Region Selection"
               Top             =   2820
               Width           =   6930
               _ExtentX        =   12224
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOREGNCDSEL"
               IdName          =   "UOREGNCDSEL"
               CmpStr          =   "CE.CmRegnCd In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   78
               Left            =   5000
               TabIndex        =   30
               ToolTipText     =   "Enter To Customer Region  (Order Cust)"
               Top             =   2820
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UOREGNCDTO"
               IdName          =   "UOREGNCDTO"
               CmpStr          =   "CE.CmRegnCd <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   79
               Left            =   1995
               TabIndex        =   29
               ToolTipText     =   "Enter From Customer Region  (Order Cust)"
               Top             =   2820
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UOREGNCDFR"
               IdName          =   "UOREGNCDFR"
               CmpStr          =   "CE.CmRegnCd >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   80
               Left            =   9675
               TabIndex        =   91
               ToolTipText     =   "Show invoice Date/ Printng Dat (Yes / No)"
               Top             =   8475
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
               Index           =   81
               Left            =   1995
               TabIndex        =   54
               ToolTipText     =   "Enter From Design Date"
               Top             =   5865
               Width           =   930
               _ExtentX        =   1640
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   3
               DataField       =   "UODT1FR"
               IdName          =   "UODT1FR"
               CmpStr          =   "DmDsgDt>="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   82
               Left            =   4995
               TabIndex        =   55
               ToolTipText     =   "Enter To Design Date"
               Top             =   5865
               Width           =   930
               _ExtentX        =   1640
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   3
               DataField       =   "UODT1TO"
               IdName          =   "UODT1TO"
               CmpStr          =   "DmDsgDt<="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   83
               Left            =   4995
               TabIndex        =   53
               ToolTipText     =   "Enter To Design By"
               Top             =   5580
               Width           =   2355
               _ExtentX        =   4154
               _ExtentY        =   503
               MaxLength       =   20
               DataField       =   "UODMDSGBYTO"
               IdName          =   "UODMDSGBYTO"
               CmpStr          =   "DmDsgBy <= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   84
               Left            =   1995
               TabIndex        =   52
               ToolTipText     =   "Enter From Design By"
               Top             =   5580
               Width           =   2355
               _ExtentX        =   4154
               _ExtentY        =   503
               MaxLength       =   20
               DataField       =   "UODMDSGBYFR"
               IdName          =   "UODMDSGBYFR"
               CmpStr          =   "DmDsgBy >= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   85
               Left            =   7980
               TabIndex        =   22
               ToolTipText     =   "Enter Customer Group Selection"
               Top             =   1935
               Width           =   6930
               _ExtentX        =   12224
               _ExtentY        =   503
               MaxLength       =   255
               DataType        =   4
               DataField       =   "UOCMGRPSEL"
               IdName          =   "UOCMGRPSEL"
               CmpStr          =   "CE.CmGrp In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   86
               Left            =   5000
               TabIndex        =   21
               ToolTipText     =   "Enter To Customer Group  (Order Cust)"
               Top             =   1935
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               DataField       =   "UOCMGRPTO"
               IdName          =   "UOCMGRPTO"
               CmpStr          =   "CE.CmGrp <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   87
               Left            =   1995
               TabIndex        =   20
               ToolTipText     =   "Enter From Customer Group  (Order Cust)"
               Top             =   1935
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               DataField       =   "UOCMGRPFR"
               IdName          =   "UOCMGRPFR"
               CmpStr          =   "CE.CmGrp >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   88
               Left            =   7980
               TabIndex        =   25
               ToolTipText     =   "Enter Customer Type Selection"
               Top             =   2235
               Width           =   6930
               _ExtentX        =   12224
               _ExtentY        =   503
               MaxLength       =   255
               DataType        =   4
               DataField       =   "UOCMTYPSEL"
               IdName          =   "UOCMTYPSEL"
               CmpStr          =   "CE.CmTyp In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   89
               Left            =   5000
               TabIndex        =   24
               ToolTipText     =   "Enter To Customer Type  (Order Cust)"
               Top             =   2235
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               DataField       =   "UOCMTYPTO"
               IdName          =   "UOCMTYPTO"
               CmpStr          =   "CE.CmTyp <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   90
               Left            =   1995
               TabIndex        =   23
               ToolTipText     =   "Enter From Customer Type  (Order Cust)"
               Top             =   2235
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               DataField       =   "UOCMTYPFR"
               IdName          =   "UOCMTYPFR"
               CmpStr          =   "CE.CmTyp >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   91
               Left            =   2950
               TabIndex        =   12
               ToolTipText     =   "Enter From Invoice Serial"
               Top             =   1080
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
               Index           =   92
               Left            =   5900
               TabIndex        =   14
               ToolTipText     =   "Enter From Invoice Serial"
               Top             =   1080
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOIDSRTO"
               IdName          =   "UOIDSRTO"
               CmpStr          =   "IdSr <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   93
               Left            =   4995
               TabIndex        =   67
               ToolTipText     =   "Enter To Purchase Order "
               Top             =   6720
               Width           =   2625
               _ExtentX        =   4630
               _ExtentY        =   503
               MaxLength       =   20
               DataField       =   "UOPONOTO"
               IdName          =   "UOPONOTO"
               CmpStr          =   "OmPoNo <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   94
               Left            =   7995
               TabIndex        =   68
               ToolTipText     =   "Enter Purchase Order Selection"
               Top             =   6735
               Width           =   6915
               _ExtentX        =   12197
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOPONOSEL"
               IdName          =   "UOPONOSEL"
               CmpStr          =   "OmPoNo In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   95
               Left            =   1995
               TabIndex        =   66
               ToolTipText     =   "Enter From Purchase Order "
               Top             =   6720
               Width           =   2625
               _ExtentX        =   4630
               _ExtentY        =   503
               MaxLength       =   20
               DataField       =   "UOPONOFR"
               IdName          =   "UOPONOFR"
               CmpStr          =   "OmPoNo >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   96
               Left            =   3285
               TabIndex        =   70
               ToolTipText     =   "Enter From Sub Purchase Order"
               Top             =   7005
               Width           =   1665
               _ExtentX        =   2937
               _ExtentY        =   503
               MaxLength       =   20
               DataField       =   "UOODPONOFR"
               IdName          =   "UOODPONOFR"
               CmpStr          =   "OdPoNo >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   97
               Left            =   6285
               TabIndex        =   72
               ToolTipText     =   "Enter To Sub Purchase Order "
               Top             =   7005
               Width           =   1665
               _ExtentX        =   2937
               _ExtentY        =   503
               MaxLength       =   20
               DataField       =   "UOODPONOTO"
               IdName          =   "UOODPONOTO"
               CmpStr          =   "OdPoNo <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   98
               Left            =   7995
               TabIndex        =   73
               ToolTipText     =   "Enter Sub Purchase Order Selection"
               Top             =   7020
               Visible         =   0   'False
               Width           =   6915
               _ExtentX        =   12197
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOODPONOSEL"
               IdName          =   "UOODPONOSEL"
               CmpStr          =   "OdPoNo in"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   115
               Left            =   1995
               TabIndex        =   69
               ToolTipText     =   "Enter From Sub Customer"
               Top             =   7005
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               DataField       =   "UOODSUBCUSTFR"
               IdName          =   "UOODSUBCUSTFR"
               CmpStr          =   "OdSubCust >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   116
               Left            =   4995
               TabIndex        =   71
               ToolTipText     =   "Enter To Sub Customer"
               Top             =   7005
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               DataField       =   "UOODSUBCUSTTO"
               IdName          =   "UOODSUBCUSTTO"
               CmpStr          =   "OdSubCust <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   99
               Left            =   2730
               TabIndex        =   217
               ToolTipText     =   "Check To Include SalCtg2 and SalCtg3"
               Top             =   4020
               Visible         =   0   'False
               Width           =   315
               _ExtentX        =   556
               _ExtentY        =   503
               Enabled         =   0   'False
               DataType        =   4
               IdName          =   "WADDALLSALCTG"
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
               Left            =   6210
               TabIndex        =   216
               Top             =   7005
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
               Index           =   38
               Left            =   3225
               TabIndex        =   215
               Top             =   7005
               Width           =   105
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "PO No"
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
               Left            =   90
               TabIndex        =   214
               Top             =   6720
               Width           =   1425
            End
            Begin VB.Label ALBL 
               BackColor       =   &H00C0C0C0&
               BackStyle       =   0  'Transparent
               Caption         =   "Sub PO No"
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
               Index           =   36
               Left            =   90
               TabIndex        =   213
               Top             =   7005
               Width           =   1425
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
               Index           =   35
               Left            =   5840
               TabIndex        =   209
               Top             =   1080
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
               Index           =   34
               Left            =   2860
               TabIndex        =   208
               Top             =   1080
               Width           =   105
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Customer Grp"
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
               Index           =   33
               Left            =   90
               TabIndex        =   207
               Top             =   1935
               Width           =   1425
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Customer Typ"
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
               Index           =   32
               Left            =   90
               TabIndex        =   206
               Top             =   2235
               Width           =   1425
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Design Date"
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
               Index           =   29
               Left            =   90
               TabIndex        =   198
               Top             =   5865
               Width           =   1410
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Design By"
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
               Left            =   90
               TabIndex        =   197
               Top             =   5580
               Width           =   1425
            End
            Begin VB.Label Label3 
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
               Left            =   5010
               TabIndex        =   195
               Top             =   8475
               Width           =   2055
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Cust Region"
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
               Index           =   23
               Left            =   90
               TabIndex        =   194
               Top             =   2820
               Width           =   1965
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
               Left            =   90
               TabIndex        =   191
               Top             =   8175
               Width           =   1005
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Order Tc/Yy/Chr"
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
               Left            =   90
               TabIndex        =   162
               Top             =   6150
               Width           =   1785
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
               Left            =   2715
               TabIndex        =   161
               Top             =   6165
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
               Left            =   3135
               TabIndex        =   160
               Top             =   6165
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
               Left            =   5700
               TabIndex        =   159
               Top             =   6165
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
               Index           =   8
               Left            =   6120
               TabIndex        =   158
               Top             =   6165
               Width           =   105
            End
            Begin VB.Label ALBL 
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
               Height          =   285
               Index           =   22
               Left            =   90
               TabIndex        =   157
               Top             =   2535
               Width           =   1965
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
               Index           =   21
               Left            =   90
               TabIndex        =   156
               Top             =   510
               Width           =   1335
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
               Index           =   20
               Left            =   7980
               TabIndex        =   155
               Top             =   60
               Width           =   1575
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Inv Date/Print Date"
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
               Left            =   90
               TabIndex        =   154
               Top             =   1365
               Width           =   1845
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "DC Wt Per Pc"
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
               Index           =   18
               Left            =   90
               TabIndex        =   153
               Top             =   7575
               Width           =   1485
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Grs Wt Per Pc"
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
               TabIndex        =   152
               Top             =   7290
               Width           =   1545
            End
            Begin VB.Label Label2 
               BackStyle       =   0  'Transparent
               Caption         =   "Show Customer"
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
               Left            =   5010
               TabIndex        =   151
               Top             =   8190
               Width           =   1575
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Order No/Sr"
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
               TabIndex        =   150
               Top             =   6435
               Width           =   1485
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
               Index           =   30
               Left            =   5820
               TabIndex        =   149
               Top             =   6450
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
               Index           =   31
               Left            =   2835
               TabIndex        =   148
               Top             =   6450
               Width           =   105
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Suffix"
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
               Index           =   16
               Left            =   90
               TabIndex        =   147
               Top             =   5010
               Width           =   1425
            End
            Begin VB.Label LblDmSz 
               BackStyle       =   0  'Transparent
               Caption         =   "Design Size"
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
               TabIndex        =   146
               Top             =   4725
               Width           =   1365
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Inv Exp No"
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
               TabIndex        =   145
               Top             =   3105
               Width           =   1425
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Kt"
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
               Left            =   90
               TabIndex        =   144
               Top             =   5295
               Width           =   1425
            End
            Begin VB.Label ALBL 
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
               Index           =   17
               Left            =   90
               TabIndex        =   143
               Top             =   1650
               Width           =   1425
            End
            Begin VB.Label ALBL 
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
               Height          =   285
               Index           =   14
               Left            =   90
               TabIndex        =   142
               Top             =   4440
               Width           =   1425
            End
            Begin VB.Label ALBL 
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
               Height          =   285
               Index           =   13
               Left            =   90
               TabIndex        =   141
               Top             =   3390
               Width           =   1425
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Sal Ctg"
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
               TabIndex        =   140
               Top             =   3675
               Width           =   1365
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Value As"
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
               TabIndex        =   139
               Top             =   7890
               Width           =   1845
            End
            Begin VB.Label LblVal 
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
               Left            =   5010
               TabIndex        =   138
               Top             =   7905
               Width           =   1665
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
               Left            =   1995
               TabIndex        =   137
               Top             =   60
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
               Left            =   5000
               TabIndex        =   136
               Top             =   60
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
               Left            =   2700
               TabIndex        =   135
               Top             =   795
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
               Left            =   5720
               TabIndex        =   134
               Top             =   795
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
               Left            =   3120
               TabIndex        =   133
               Top             =   795
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
               Left            =   6140
               TabIndex        =   132
               Top             =   795
               Width           =   105
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Inv No/Sr"
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
               TabIndex        =   131
               Top             =   1080
               Width           =   1605
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
               Left            =   90
               TabIndex        =   130
               Top             =   795
               Width           =   1725
            End
         End
      End
      Begin VB.Frame FraSc 
         Height          =   2925
         Left            =   1920
         TabIndex        =   211
         Top             =   6600
         Visible         =   0   'False
         Width           =   3945
         Begin MwfCtl.MWCTL_FLX GrdDsgAna 
            Height          =   2430
            Left            =   120
            TabIndex        =   212
            Top             =   240
            Width           =   3555
            _ExtentX        =   6271
            _ExtentY        =   4286
            Cols            =   4
            colname1        =   "UDANASR"
            heading1        =   "Sr. "
            datafld1        =   "UdAnaSr"
            datatype1       =   4
            recalcparent1   =   "GrdDsgAna(UDANACDFR),GrdDsgAna(UDANACDTO)"
            maxlength1      =   2
            tooltiptext1    =   "Enter Analysis Sr. Number"
            colname2        =   "UDANACDFR"
            heading2        =   "Fr Ana Cd      "
            datafld2        =   "UdAnaCdFr"
            datatype2       =   4
            recalcon2       =   "GrdDsgAna(UDANASR)"
            maxlength2      =   8
            tooltiptext2    =   "Enter From Analysis Code"
            colname3        =   "UDANACDTO"
            heading3        =   "To Ana Cd      "
            datafld3        =   "UdAnaCdTo"
            datatype3       =   4
            recalcon3       =   "GrdDsgAna(UDANASR)"
            maxlength3      =   8
            tooltiptext3    =   "Enter To Analysis Code"
         End
      End
   End
End
Attribute VB_Name = "EmrFrpSalAnaDsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rep As New EmrRepSalAnaDsg
Dim moCn As MwfLib.MDOConnection   '*** (Bef speed) bef 24/11/06

' ########################  Manoj #### Ver: 2.11.0 #### Date: 26/05/2005  ########################
Dim moGrph As EmrClsGrph
' ###########################################  Manoj  ###########################################

'****** Sachin 2.14.0 - [25-11-2006] - [Multi-Company] ******
Dim ws_MultiCoMod As Boolean
'****** Sachin 2.14.0 - [25-11-2006] - [Multi-Company] ******

Private Sub SetGroupSort()
  GRP_REP.Add "Design Ctg", "max(DmCtg)", "", "DmCtg", "hDmCtg", "Sales Ctg", "", "", "(Select PDesc From Param Where PTyp= 'DMCTG' and PmCd= max(DmCtg)) "
  'Geeta***Crm***212
  'uni.4 max(DmSalCtg) replaced by qDmSalCtgRepl,qDmSalCtgDescRepl to get either SalCtg,SalCTg2 or SalCTg3
  GRP_REP.Add "Sales Ctg", "||qDmSalCtgRepl||", "", "", "", "Design Ctg", "", "", "(Select PDesc From Param Where PTyp= 'SALCTG' and PmCd=max(DmCtg) and PSCd= ||qDmSalCtgDescRepl|| ) "
  GRP_REP.Add "Inv Vch No", "IdCoCd+IdTc+IdYy+IdChr+str(IdNo)", "IdCoCd+'/'+IdTc+'/'+IdYy+'/'+IdChr+'/'+LTrim(Str(IdNo))", "", "", "Company Code,Invoice Cust,Inv Date,Inv Exp No,Customer Grp,Customer Typ", "", "", "max(InExpNo) "
  GRP_REP.Add "Inv Exp No", "IdCoCd+max(InExpNo)", "'['+IdCoCd+'] '+max(InExpNo)", "InExpNo", "hInExpNo", "Company Code,Inv Date,Inv Vch No,Invoice Cust,Customer Grp,Customer Typ", "", "", ""
  'GRP_REP.Add "Inv Date", "convert(varchar(10), max(InDt), 112)", "convert(varchar(8), max(InDt), 3)", "InDt", "hInDt", "Quarter,Month,Week", "", "", ""
  ' ****** Manali 3.5.0 - 26/11/08 - Option for InDt / InPrnDt added
''  ' Bef 3.10.0
''  GRP_REP.Add "Inv Date", "convert(varchar(10), max(InDt), 112)", "convert(varchar(8), max(InDt), 3)", "", "", "Quarter,Month,Week", "", "", ""
''  GRP_REP.Add "Prn Date", "convert(varchar(10), max(InPrnDt), 112)", "convert(varchar(8), max(InPrnDt), 3)", "", "", "Quarter,Month,Week", "", "", ""
''  ' ****** Manali 3.5.0 - 26/11/08 - Option for InDt / InPrnDt added
''  GRP_REP.Add "Week", "LTrim(Str(DatePart(yyyy,max(InDt))))+' - '+DateName(WW,max(InDt))", "", "", "", "Quarter,Month", "", "", ""
''  GRP_REP.Add "Month", "Str(DatePart(yyyy,max(InDt)))+Str(DatePart(mm,max(InDt)))", "LTrim(Str(DatePart(yyyy,max(InDt))))+' '+DateName(MM,max(InDt))", "", "", "Quarter", "", "", ""
''  GRP_REP.Add "Quarter", "LTrim(Str(DatePart(yyyy,max(InDt))))+' - '+DateName(QQ,max(InDt))", "", "", "", "", "", "", ""
  
  ' ***** Manali  3.10.0 - 10/03/12 - Sort Based on InDt/PrnDt
  '6.1
  GRP_REP.Add "Inv/Prn Date", "convert(varchar(10), max(Case When qDtOpt='InDt'Then InDt Else InPrnDt End), 112)", "convert(varchar(8), max(Case When qDtOpt='InDt'Then InDt Else InPrnDt End), " + CStr(DtFmtSlahVal()) + ")", "", "", "Quarter,Month,Week", "", "", ""
  GRP_REP.Add "Week", "LTrim(Str(DatePart(yyyy,max(Case When qDtOpt='InDt'Then InDt Else InPrnDt End))))+' - '+DateName(WW,max(Case When qDtOpt='InDt'Then InDt Else InPrnDt End))", "", "", "", "Quarter,Month", "", "", ""
  GRP_REP.Add "Month", "Str(DatePart(yyyy,max(Case When qDtOpt='InDt'Then InDt Else InPrnDt End)))+Str(DatePart(mm,max(Case When qDtOpt='InDt'Then InDt Else InPrnDt End)))", "LTrim(Str(DatePart(yyyy,max(Case When qDtOpt='InDt'Then InDt Else InPrnDt End))))+' '+DateName(MM,max(Case When qDtOpt='InDt'Then InDt Else InPrnDt End))", "", "", "Quarter", "", "", ""
  'MW.134 - group sort Year added.
  GRP_REP.Add "Year", "LTrim(Str(DatePart(yyyy,max(Case When qDtOpt='InDt'Then InDt Else InPrnDt End))))", "", "", "", "", "", "", ""
  GRP_REP.Add "Quarter", "LTrim(Str(DatePart(yyyy,max(Case When qDtOpt='InDt'Then InDt Else InPrnDt End))))+' - '+DateName(QQ,max(Case When qDtOpt='InDt'Then InDt Else InPrnDt End))", "", "", "", "", "", "", ""
  ' ***** Manali  3.10.0 - 10/03/12 - Sort Based on InDt/PrnDt
  
  
  GRP_REP.Add "Design Code", "max(IdDmCd)", "", "IdDmCd", "hIdDmCd", "Design Ctg,Sales Ctg", "", "", "max(DmDesc) "
  GRP_REP.Add "Design Size", "max(IdDmSz)", "", "IdDmSz", "hIdDmSz", "", "", "", "(Select PDesc From Param Where PTyp= 'DMSZ' and PmCd= max(IdDmSz))"
  GRP_REP.Add "Suffix", "max(IdSfx)", "", "IdSfx", "hIdSfx", "", "", "", ""
  '******Geeta*****************Emr206****06/04/04*******
  '**********If UserLevel>6 then don't display the Customer Name
  GRP_REP.Add "ExpOrd Cust", "max(IdExpCmCd)", "", "", "", "Sales Exec", "", "", "case when (Select PNum from Param where Ptyp = 'USR' and PmCd = '" + Trim(EmrFrmMDI.TXT_UCD.text) + "')  <= 5 then  max(CE.CmName) else max(IdExpCmCd) end "  'From CustMst Where CmCtg= 'C' and CmCd= max(IdExpCmCd))
  
  'Sachin - 4.1.0.0
  GRP_REP.Add "Invoice Cust", "max(InCmCd)", "", "", "", "Customer Grp,Customer Typ", "", "", "(Select (case when (Select PNum from Param where Ptyp = 'USR' and PmCd = '" + Trim(EmrFrmMDI.TXT_UCD.text) + "')  <= 5 then  CmName else  max(InCmCd) end) From CustMst Where CmCtg= 'C' and CmCd= max(InCmCd)) "
  
  'Sachin - 4.1.0.0
  GRP_REP.Add "Customer Grp", "max(CE.CmGrp)", "", "", "", "", "", "", "(Select PDesc from Param where Ptyp = 'CMGRP' and PmCd = Max(CE.CmGrp))"
  GRP_REP.Add "Customer Typ", "max(CE.CmTyp)", "", "", "", "", "", "", "(Select PDesc from Param where Ptyp = 'CMTYP' and PmCd = Max(CE.CmTyp))"
  
  GRP_REP.Add "Order", "IdCoCd+max(IdExpOdTc)+max(IdExpOdYy)+max(IdExpOdChr)+str(max(IdExpOdNo))", "IdCoCd+'/'+max(IdExpOdTc)+'/'+max(IdExpOdYy)+'/'+max(IdExpOdChr)+'/'+LTrim(Str(max(IdExpOdNo)))", "", "", "ExpOrd Cust,Sales Exec", "", "", "max(IdExpCmCd) "
  
  GRP_REP.Add "Cust PO", "max(OmPoNo)", "max(OmPoNo)", "", "", "", "", "", "max(OmPoNo) "
  'vk.19 sub customer added
  GRP_REP.Add "Cust SUB PO", "max(OdSubCust)+max(OdPoNo)", "max(OdSubCust)+(case when max(OdPoNo)<> '' then ' : ' else '' end)+max(OdPoNo)", "", "", "", "", "", "max(OdSubCust)+(case when max(OdPoNo)<> '' then ' : ' else '' end)+max(OdPoNo) "
  
'  GRP_REP.Add "Cust PO", "(Select max(OmPoNo) From OrdMst Where OmCoCd=IdCoCd and OmTc=Max(IdExpOdTc) and OmYy=Max(IdExpOdYy) and OmChr=Max(IdExpOdChr)  and OmNo=Max(IdExpOdNo)) ", "(Select max(OmPoNo) From OrdMst Where OmCoCd=IdCoCd and OmTc=Max(IdExpOdTc) and OmYy=Max(IdExpOdYy) and OmChr=Max(IdExpOdChr)  and OmNo=Max(IdExpOdNo)) ", "", "", "", "", "", "(Select max(OmPoNo) From OrdMst Where OmCoCd=IdCoCd and OmTc=Max(IdExpOdTc) and OmYy=Max(IdExpOdYy) and OmChr=Max(IdExpOdChr)  and OmNo=Max(IdExpOdNo)) "
'  GRP_REP.Add "Cust Sub PO", "(Select max(OdPoNo) From OrdDsg Where OdCoCd=IdCoCd and OdTc=Max(IdExpOdTc) and OdYy=Max(IdExpOdYy) and OdChr=Max(IdExpOdChr)  and OdNo=Max(IdExpOdNo) and OdSr=Max(IdExpOdSr))", "(Select max(OdPoNo) From OrdDsg Where OdCoCd=IdCoCd and OdTc=Max(IdExpOdTc) and OdYy=Max(IdExpOdYy) and OdChr=Max(IdExpOdChr)  and OdNo=Max(IdExpOdNo) and OdSr=Max(IdExpOdSr))", "", "", "", "", "", "(Select max(OdPoNo) From OrdDsg Where OdCoCd=IdCoCd and OdTc=Max(IdExpOdTc) and OdYy=Max(IdExpOdYy) and OdChr=Max(IdExpOdChr)  and OdNo=Max(IdExpOdNo) and OdSr=Max(IdExpOdSr))"

  
  GRP_REP.Add "Karat", "max(IdKt)", "", "IdKt", "hIdKt", "", "", "", "(Select PDesc From Param Where PTyp= 'KT' and PMCd = max(IdKt)) "
  GRP_REP.Add "PCS/PRS", "max(DmUom)", "", "", "", "", "", "", "(Select PDesc From Param Where PTyp= 'DMUOM' and PmCd= max(DmUom)) "
  GRP_REP.Add "VaCtg", "max(IdVaCtg)", "", "", "", "", "", "", "(Select PDesc From Param Where PTyp= 'VACTG' and PMCd = max(IdVaCtg)) "
  GRP_REP.Add "(DsgCd)Colour", "max(IdDmCd+OdDmCol)", "'('+max(IdDmCd)+')'+max(OdDmCol)", "IdDmCd", "hIdDmCd", "Design Ctg,Sales Ctg,Design Code,Design Colour", "", "", "(Select PDesc From Param Where PTyp= 'DMCOL' and PMCd = max(OdDmCol)) "
  GRP_REP.Add "Design Colour", "max(OdDmCol)", "", "", "", "(DsgCd)Colour", "", "", "(Select PDesc From Param Where PTyp= 'DMCOL' and PMCd = max(OdDmCol)) "
  GRP_REP.Add "Sales Exec", "max(OmCmSalPer)", "", "", "", "", "", "", "(Select PDesc from Param Where PTyp = 'SP' and PMCd= max(OmCmSalPer)) "
  

  '****** Sachin 2.14.0 - [23-11-2006] - [Multi-Company] - ,Process,OrderNo,OrderSr,Customer PO,FgSubLoc ******
  If ws_MultiCoMod = True Then _
    GRP_REP.Add "Company Code", "IdCoCd", "IdCoCd", "", "", "", "", "", "(Select HName from Head where HCd= '" + ctSelfCmCd + "' and HCoCd=IdCoCd) "
  '****** Sachin 2.14.0 - [23-11-2006] - [Multi-Company] - ******

' ########################  Manoj #### Ver: 2.11.0 #### Date: 26/05/2005  ########################
  moGrph.AddItem "Invoice FOB", "IdIFOB", "{rdo.IdIFOB}", ""
  moGrph.AddItem "Design Quantity", "IdQty", "{rdo.IdQty}", ""
  moGrph.AddItem "Gross Weight", "QGrWt", "{rdo.QGrWt}", ""
  moGrph.AddItem "Labour Value", "QLabVal", "{rdo.QLabVal}", ""
  moGrph.AddItem "Dia CS Weight", "QDCWt", "{rdo.QDCWt}", ""
  moGrph.AddItem "(Lab + Set) Val", "QLabVal + QSetVal", "{rdo.QLabVal} + {rdo.QSetVal}", ""
' ###########################################  Manoj  ###########################################
End Sub
Private Sub adc_setreprecsource()
  '*** Report Sql ***
  'Select space(30) as Grp1, space(30) as Grp2, space(30) as Grp3,
  'space(30) as DGrp1, space(30) as DGrp2, space(30) as DGrp3,
  'space(30) as G1Desc, space(30) as G2Desc, space(30) as G3Desc,
  'IdCoCd, IdTc, IdYy, IdChr, IdNo, IdSr, InExpNo, DmCtg, IdDmCd, IdSfx, IdDmSz, IdKt, InDt,
  'InCmCd, IdExpOdTc, IdExpOdYy, IdExpOdChr, IdExpOdNo, IdExpOdSr, DmSalCtg, IdVaCtg,
  'IdQty, IdIFob, IrRmIVal as qMetVal, IrRmIVal as qDCVal, IrRmIVal as qXVal,
  'IrSetIVal as qSetVal, iLiVal as qLabVal, IrRmIWt as qGrWt, IrRmIWt as qMetWt,
  'IrRmIWt As qDCWt, IrRmIWt As qXWt,
  'OdPicNm
  'from InvDsg, InvHd, DsgMst, InvRm, RmMst, InvLab, OrdDsg where 1= 2
  '*** Report Sql ***
  
  Dim grpflds As String, wSqlStrg As String, wStr1 As String, wStr2 As String, wStr3 As String
  Dim wrepcnd As String, wCnd As String, wHavingCnd As String, wQtyStr As String
  Dim wOpt As String, wIdFOBOpt As String, wIrRmWtOpt As String, wIrSetValOpt As String
  Dim wIrRmValOpt As String, wiLValOpt As String, wCust As String, wOrdJoin As String
  Dim ws_DsgAnaCnd As String
  
  ' Zubin 213
  Dim wb_InBaseCur As Boolean
  wb_InBaseCur = IIF(ADC("UoCurCdFr") = "", True, False)
  
  Set Rep = Nothing
  Set ADC.RepSource = Rep
  Call DispCoNm
  Call SetMwName(Rep)   ' ***** Manali 3.6.1 - Microway footer
  Call GRP_REP.Gen3LRep
  grpflds = GRP_REP.GrpFldLst
   
    Rep.hExpDt.Suppress = True
      Rep.qExpDt.Suppress = True
      Rep.hDelDt.Suppress = True
      Rep.qDelDt.Suppress = True
  
  If ADC("UoYN1") = "Y" Then
    wCust = "InCmCd"
    Rep.hInCmCd.SetText "Inv Cust"
  Else
    wCust = "IdExpCmCd"
    Rep.hInCmCd.SetText "Ord Cust"
  End If
  ADC("UoCmCdFr").CmpStr = wCust + ">= "
  ADC("UoCmCdTo").CmpStr = wCust + "<= "
  ADC("UoCmCdSel").CmpStr = wCust + " In "
  
'  adc("UoSalPerFr").CmpStr = IIF(adc("UoYN1") = "Y", "ce.", "c.") + "CmSalPer >= "
'  adc("UoSalPerTo").CmpStr = IIF(adc("UoYN1") = "Y", "ce.", "c.") + "CmSalPer <= "
'  adc("UoSalPerSel").CmpStr = IIF(adc("UoYN1") = "Y", "ce.", "c.") + "CmSalPer In "
'
  ' ****** Manali 3.5.0 - 25/11/08 - Option for InDt / InPrnDt added
  Dim wInvDt As String
  If ADC("UoYN3") = "Y" Then
    wInvDt = "InDt "
    Rep.hInDt.SetText "Inv Dt"
  Else
    wInvDt = "InPrnDt "
    Rep.hInDt.SetText "Prn Dt"
  End If
    ADC("UoDtFr").CmpStr = wInvDt + ">= "
    ADC("UoDtTo").CmpStr = wInvDt + "<= "
  ' ****** Manali 3.5.0 - 25/11/08 - Option for InDt / InPrnDt added
  
  wrepcnd = ADC.RepCond
  wCnd = IIF(wrepcnd <> "", " where " + wrepcnd, "")
  
  ' Zubin 213
  wCnd = wCnd + IIF(wb_InBaseCur = False, " and CE.CmCurCd= '" + ADC("UoCurCdFr") + "'", "")
  
  '*** To check if 'Design Colour' group is selected ***
  Dim wGrp(6) As String, i As Integer, wDsgColGrpPos As Integer, wDetPos As Integer
  wGrp(1) = GRP_REP.Value(gltGroup1)
  wGrp(2) = GRP_REP.Value(gltGroup2)
  wGrp(3) = GRP_REP.Value(gltGroup3)
  wDsgColGrpPos = 0
  For i = 1 To 3
    If wGrp(i) = "(DsgCd)Colour" Or wGrp(i) = "Design Colour" Then wDsgColGrpPos = i: Exit For
  Next i
  '*** To 'Design Colour' group not selected then ***
  
  '*** To check if 'Detail' group is selected ***
  ' ****** Manali 3.5.0 - 11/11/08 - Show Picture Option in report
  i = 1
  wDetPos = 0
  For i = 1 To 6
    If wGrp(i) = "(Detail)" Then wDetPos = i: Exit For
  Next i
  
  '*** If Detail not selected then suppress Box Header, Detail and Footer ***
  If wDetPos = 0 Then Rep.PICSEC.Suppress = True
  ' ****** Manali 3.5.0 - 11/11/08 - Show Picture Option in report

  ' ****** Manali 3.5.0 - 11/11/08 - Show Picture Option in report - Condition added for Picture
  If wDsgColGrpPos = 0 And (wDetPos = 0 Or ADC("UoYN2") = "N") Then
    wOrdJoin = ""
  Else
    wOrdJoin = " Join OrdDsg on " + IIF(gs_Partition = ctCurrPrtn, " OdPrtKey=IdPrtKey and ", "") + " OdIdNo=IdOdIdNo and OdCoCd= IdCoCd and OdTc= IdExpOdTc and OdYy= IdExpOdYy " + _
               " and OdChr= IdExpOdChr and OdNo= IdExpOdNo and OdSr= IdExpOdSr "
  End If
  wOrdJoin = ""
  
  If UCase(ADC("UoInVal")) = "A" Then
    wOpt = "Actual"   '''"Actual"
    wIdFOBOpt = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then IdAVal/InCnvFct else IdAVal*InCnvFct end) ", "IdAVal")
    wIrRmWtOpt = "IrRmAWt"
    wIrSetValOpt = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then IrSetAVal/InCnvFct else IrSetAVal*InCnvFct end) ", "IrSetAVal")
    wIrRmValOpt = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then IrRmAVal/InCnvFct else IrRmAVal*InCnvFct end) ", "IrRmAVal")
    wiLValOpt = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then iLAVal/InCnvFct else iLAVal*InCnvFct end) ", "iLAVal")
  ElseIf UCase(ADC("UoInVal")) = "I" Then
    wOpt = "Invoice"      '''"Invoice"
    wIdFOBOpt = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then IdIFob/InCnvFct else IdIFob*InCnvFct end) ", "IdIFob")
    wIrRmWtOpt = "IrRmIWt"
    wIrSetValOpt = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then IrSetIVal/InCnvFct else IrSetIVal*InCnvFct end) ", "IrSetIVal")
    wIrRmValOpt = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then IrRmIVal/InCnvFct else IrRmIVal*InCnvFct end) ", "IrRmIVal")
    wiLValOpt = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then iLiVal/InCnvFct else iLiVal*InCnvFct end) ", "iLiVal")
  ElseIf UCase(ADC("UoInVal")) = "Z" Then
    wOpt = "Customs"    '''"Customs"
    wIdFOBOpt = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then IdZFob/InCnvFct else IdZFob*InCnvFct end) ", "IdZFob")
    wIrRmWtOpt = "IrRmZWt"
    wIrSetValOpt = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then IrSetZVal/InCnvFct else IrSetZVal*InCnvFct end) ", "IrSetZVal")
    wIrRmValOpt = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then IrRmZVal/InCnvFct else IrRmZVal*InCnvFct end) ", "IrRmZVal")
    wiLValOpt = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then iLZVal/InCnvFct else iLZVal*InCnvFct end) ", "iLZVal")
  End If
  
  wQtyStr = " (case when max(IdQty)> 0 then max(IdQty) else 1 end) "
  If ADC("UoRmWtFr") > 0 Then wHavingCnd = " having sum(case when IsNull(RmCtg, '') in ('C', 'D') then " + wIrRmWtOpt + " else 0 end)/ " + wQtyStr + ">= " + CStr(ADC("UoRmWtFr")) + " "
  If ADC("UoRmWtTo") > 0 Then wHavingCnd = wHavingCnd + IIF(wHavingCnd <> "", " and ", " having ") + " sum(case when IsNull(RmCtg, '') in ('C', 'D') then " + wIrRmWtOpt + " else 0 end)/ " + wQtyStr + "<= " + CStr(ADC("UoRmWtTo")) + " "
  If ADC("UoGrsWtFr") > 0 Then wHavingCnd = wHavingCnd + IIF(wHavingCnd <> "", " and ", " having ") + " sum(case when IsNull(RmCtg, '') in ('C', 'D') then " + wIrRmWtOpt + "/5 else " + wIrRmWtOpt + " end)/ " + wQtyStr + ">= " + CStr(ADC("UoGrsWtFr")) + " "
  If ADC("UoGrsWtTo") > 0 Then wHavingCnd = wHavingCnd + IIF(wHavingCnd <> "", " and ", " having ") + " sum(case when IsNull(RmCtg, '') in ('C', 'D') then " + wIrRmWtOpt + "/5 else " + wIrRmWtOpt + " end)/ " + wQtyStr + "<= " + CStr(ADC("UoGrsWtTo")) + " "
 
  '3.11.0
  Dim wBSCnd As String
  If ChkBoxArr(0).Value = Checked Then
    wBSCnd = IIF(ADC("wExpQtyFr") <> 0, " And IdQty>=" + CStr(ADC("wExpQtyFr")), "") + _
             IIF(ADC("wExpQtyTo") <> 0, " And IdQty<=" + CStr(ADC("wExpQtyTo")), "") + _
             IIF(ADC("wInvFobFr") <> 0, " And " + wIdFOBOpt + ">=" + CStr(ADC("wInvFobFr")), "") + _
             IIF(ADC("wInvFobTo") <> 0, " And " + wIdFOBOpt + "<=" + CStr(ADC("wInvFobTo")), "") + _
             IIF(Trim(ADC("wBSYy")) <> "", " And IdYy='" + CStr(ADC("wBSYy")) + "'", "")
    wBSCnd = IIF(wCnd <> "", wBSCnd, " Where " + wBSCnd)
    wStr3 = " Order by (Case When '" + Trim(ADC("wBSQV")) + "'='Y' Then Max(IdQty) Else max(" + wIdFOBOpt + ") End) Desc"
  Else
    wBSCnd = ""
    wStr3 = " Order by IdCoCd, IdTc, IdYy, IdChr, IdNo, IdSr"
  End If
 
  ws_DsgAnaCnd = " And (Select Count(*) From UsrOptDsgAna " + _
                 "Join DsgAna on " + IIF(gs_Partition = ctCurrPrtn, " DaPrtKey='" + ctCurrPrtn + "' and ", "") + " DaAnaSr = UdAnaSr " + _
                 "where (DaAnaCd between UdAnaCdFr and UdAnaCdTo) " + _
                 "and DaTcTyp = DmTcTyp and DaCd= DmCd and DaSz= DmSz and " + _
                 "UdUsrCd = '" + ADC("UoUsrCd") + "' and " + _
                 "UdMnuCd= '" + ADC("UoMnuCd") + "')= " & (GrdDsgAna.Rows - 1) & " "
 
  'Uni.4 - SalCtg scope will check all of 3 Sal Ctg(SalCtg, SalCtg2 & SalCtg3).
  'if include SalCtg2 and SalCtg3 is YEs then checks all SalCtg else behave as it is
  'calling function to get where clause to join DsgMst
  Dim wDmSalCtgCnd As String, ws_DmSalCtg As String
  If ADC("UoDmSalCtgFr") <> "" Or ADC("UoDmSalCtgTo") <> "" Or ADC("UoDmSalCtgSel") <> "" Then
    wDmSalCtgCnd = GetDmSalCtgWhrCnd(ADC("UoDmSalCtgFr"), ADC("UoDmSalCtgTo"), ADC("UoDmSalCtgSel"), ADC("wAddAllSalCtg"))
  End If
    
  'uni.4 selecting Dmsalctg either from SalCtg,SalCtg2 or SalCtg3
  'if include SalCtg2 and SalCtg3 is YEs then checks all SalCtg else behave as it is
  'calling function to get SalCtg value to show in detail section and group sort
  ws_DmSalCtg = GetDmSalCtg(ADC("UoDmSalCtgFr"), ADC("UoDmSalCtgTo"), ADC("UoDmSalCtgSel"), ADC("wAddAllSalCtg"), "Y")
  
  
  '" + IIF(wb_InBaseCur, " Join InvHd on InCoCd= IrCoCd and InTc= IrTc and InYy= IrYy and InChr= IrChr and InNo= IrNo ", "")
  '****** Sachin 3.02 - Id fields in Join
  ' ***** Manali 3.03 - 24/06/08 - Pd Mod - 'L' added
  ' ****** Manali 3.5.0 - 11/11/08 - Show Picture Option in report - OdPicNm added
  ' ****** Manali 3.5.0 - 25/11/08 - Option for InDt / InPrnDt added
  ' ***** Manali  3.10.0 - 10/03/12 - Sort Based on InDt/PrnDt
  'uni.4 max(DmSalCtg) replaced by ws_DmSalCtg
  '62 qHuidNo added
  wStr1 = "Select " + grpflds + ", " + _
          "IdCoCd, IdTc, IdYy, IdChr, IdNo, IdSr, max(InExpNo) as InExpNo, max(DmCtg) as DmCtg, max(IdDmCd) as IdDmCd, " + _
          "max(IdSfx) as IdSfx, max(IdDmSz) as IdDmSz, max(IdKt) as IdKt, /* max(InDt) as InDt,*/ max(" + wInvDt + ") as InDt, " + _
          "max(" + wCust + ") as InCmCd, max(IdExpOdTc) as IdExpOdTc, " + _
          "max(IdExpOdYy) as IdExpOdYy, max(IdExpOdChr) as IdExpOdChr, max(IdExpOdNo) as IdExpOdNo, " + _
          "max(IdExpOdSr) as IdExpOdSr, " + ws_DmSalCtg + " as DmSalCtg, " + _
          "max(IdVaCtg) as IdVaCtg, max(IdQty) as IdQty, 0 as OdExpQty, max(" + wIdFOBOpt + ") as IdIFob, " + _
          "sum(case when IsNull(RmCtg, '') in ('G', 'P', 'S', 'L') then IsNull(" + wIrRmValOpt + ", 0) else 0 end) as qMetVal, " + _
          "sum(case when IsNull(RmCtg, '') in ('D') then IsNull(" + wIrRmValOpt + ", 0) else 0 end) as qDVal, " + _
          "sum(case when IsNull(RmCtg, '') in ('C') then IsNull(" + wIrRmValOpt + ", 0) else 0 end) as qCVal, " + _
          "sum(case when IsNull(RmCtg, '') in ('X') then IsNull(" + wIrRmValOpt + ", 0) else 0 end) as qXVal, " + _
          "sum(IsNull(" + wIrSetValOpt + ", 0)) as qSetVal, " + _
          "IsNull((select sum(" + wiLValOpt + ") from InvLab " + _
          IIF(wb_InBaseCur, " Join InvHd on " + IIF(gs_Partition = ctCurrPrtn, " IlPrtKey=IdPrtKey and ", "") + " InCoCd= IlCoCd and InTc= IlTc and InYy= IlYy and InChr= IlChr and InNo= IlNo ", "") + _
          "where iLCoCd= IdCoCd and iLTc= IdTc " + _
          "and iLYy= IdYy and iLChr= IdChr and iLNo= IdNo and iLSr= IdSr), 0) as qLabVal, " + _
          "sum(case when IsNull(RmCtg, '') in ('D', 'C') then IsNull(" + wIrRmWtOpt + ", 0)/ 5 else IsNull(" + wIrRmWtOpt + ", 0) end) as qGrWt, " + _
          "sum(case when IsNull(RmCtg, '') in ('G', 'P', 'S', 'L') then IsNull(" + wIrRmWtOpt + ", 0) else 0 end) as qMetWt, " + _
          "sum(case when IsNull(RmCtg, '') in ('D') then IsNull(" + wIrRmWtOpt + ", 0) else 0 end) as qDWt, " + _
          "sum(case when IsNull(RmCtg, '') in ('C') then IsNull(" + wIrRmWtOpt + ", 0) else 0 end) as qCWt, " + _
          "sum(case when IsNull(RmCtg, '') in ('X') then IsNull(" + wIrRmWtOpt + ", 0) else 0 end) as qXWt " + _
          IIF(ADC("UoYN2") = "Y" And wDetPos > 0, ", max(OdPicNm) as OdPicNm ", ", '' As OdPicNm ") + _
          ",(select isnull(max(CdCertNo),'') from CertDet join InvFgd on CdBCoCd =IfCoCd AND CdBYy =IfFdBYy and CdBChr =IfFdBChr " + _
          "and CdBNo =IfFdBNo where CdCtg='H' and IfCoCd= IdCoCd and IfTc= IdTc and IfYy= IdYy and IfChr= IdChr and IfNo= IdNo And IfSr =IdSr) as qHuidNo "
  'uni.4 wDmSalCtgCnd added to DsgMst join, to get SalCtg value from scope.
  'To avoid SalCtg scope condition in where clause
  wStr2 = "From InvDsg join InvHd on InIdNo=IdInIdNo and InCoCd= IdCoCd and InTc= IdTc and InYy= IdYy " + _
          " and InChr= IdChr and InNo= IdNo " + IIF(gs_Partition = ctCurrPrtn, " and InPrtKey=IdPrtKey ", "") + _
          " Join OrdDsg on OdCoCd= IdCoCd and OdTc=IdExpOdTc and OdYy=IdExpOdYy and OdChr =IdExpOdChr and OdNo=IdExpOdNo and OdSr = IdExpOdSr and OdIdNo=IdOdIdNo " + IIF(gs_Partition = ctCurrPrtn, " and OdPrtKey=IdPrtKey ", "") + _
          " Join OrdMst On OmCoCd=OdCoCd and OmTc=OdTc and OmYy=OdYy and OmChr=OdChr and OmNo=OdNo  " + IIF(gs_Partition = ctCurrPrtn, " and OmPrtKey=IdPrtKey ", "") + _
          " Join CustMst CE on CE.CmCtg= 'C' and CE.CmCd = IdExpCmCd " + _
          " Join CustMst C on C.CmCtg= 'C' and C.CmCd =InCmCd " + _
          " Join DsgMst on DmIdNo=IdDmIdNo and DmTcTyp= 'DM' and " + _
          " DmCd= IdDmCd and DmSz= '' " + IIF(gs_Partition = ctCurrPrtn, " and DmPrtKey=IdPrtKey ", "") + wDmSalCtgCnd + wOrdJoin + _
          " Join (Select '" + Trim(wInvDt) + "' as qDtOpt) A on 1=1 " + _
          " left outer join InvRm on IrIdIdNo=IdIdNo and IrCoCd= IdCoCd and IrTc= IdTc and IrYy= IdYy and " + _
          " IrChr= IdChr and IrNo= IdNo and IrSr= IdSr " + IIF(gs_Partition = ctCurrPrtn, " and IrPrtKey=IdPrtKey ", "") + " left outer join RmMst on RmCd= IrRmCd " + IIF(gs_Partition = ctCurrPrtn, " and RmPrtKey=IrPrtKey ", "") + _
          wCnd + wBSCnd + ws_DsgAnaCnd + IIF(gs_Partition = ctCurrPrtn, " and IdPrtKey='" + ctCurrPrtn + "' ", "") + " group by IdPrtKey, IdCoCd, IdTc, IdYy, IdChr, IdNo, IdSr " + wHavingCnd
  
  wSqlStrg = wStr1 + wStr2 + wStr3
  'uni.4 replacing DmSalCtg value and its description in group sort of Sales Ctg
  wSqlStrg = Replace(wSqlStrg, "||qDmSalCtgRepl||", ws_DmSalCtg)
  wSqlStrg = Replace(wSqlStrg, "||qDmSalCtgDescRepl||", ws_DmSalCtg)
  
  ADC.RepRecSource = wSqlStrg + ctMaxDopOpt '****** Sachin 3.01 added ctMaxDopOpt
  
  
  If ADC("UoYN") = "Y" Then
    Rep.TxtHead.SetText "Design Sales Analysis (Value)"
    GRP_REP.SetFormula Rep, "wRndOffDecPlaces", "2"
    GRP_REP.SetFormula Rep, "wTotVal", "{rdo.IdIFOB}"
    GRP_REP.SetFormula Rep, "wMetVal", "{rdo.qMetVal}"
    GRP_REP.SetFormula Rep, "wDVal", "{rdo.qDVal}"
    GRP_REP.SetFormula Rep, "wCVal", "{rdo.qCVal}"
    GRP_REP.SetFormula Rep, "wXVal", "{rdo.qXVal}"
    GRP_REP.SetFormula Rep, "wSetVal", "{rdo.qSetVal}"
    GRP_REP.SetFormula Rep, "wLabVal", "{rdo.qLabVal}"
    
''    Rep.h1Tot.SetText wOpt + Chr(13) + "Value": Rep.h2Tot.SetText wOpt + Chr(13) + "Value"
''    Rep.h1Met.SetText wOpt + Chr(13) + "Met Value": Rep.h2Met.SetText wOpt + Chr(13) + "Met Value"
''    Rep.h1DC.SetText wOpt + Chr(13) + "DC Value": Rep.h2DC.SetText wOpt + Chr(13) + "DC Value"
''    Rep.h1X.SetText wOpt + Chr(13) + "Acc Value": Rep.h2X.SetText wOpt + Chr(13) + "Acc Value"
''    Rep.h1SetVal.SetText wOpt + Chr(13) + "Set Value": Rep.h2SetVal.SetText wOpt + Chr(13) + "Set Value"
''    Rep.h1LabVal.SetText wOpt + Chr(13) + "Lab Value": Rep.h2LabVal.SetText wOpt + Chr(13) + "Lab Value"

    
   Rep.TxtValTyp.SetText wOpt + " Values"

    Rep.h1Tot.SetText " Val": Rep.h2Tot.SetText " Val"
    Rep.h1Met.SetText " Met Val": Rep.h2Met.SetText " Met Val"
    Rep.h1D.SetText " Dia Val": Rep.h2D.SetText " Dia Val"
    Rep.h1C.SetText " Cs Val": Rep.h2C.SetText " Cs Val"
    Rep.h1X.SetText " Acc Val": Rep.h2X.SetText " Acc Val"
    Rep.h1SetVal.SetText " Set Val": Rep.h2SetVal.SetText " Set Val"
    Rep.h1LabVal.SetText " Lab Val": Rep.h2LabVal.SetText " Lab Val"

    ' Zubin 213
    If wb_InBaseCur Then
      Rep.TxtValIn.SetText "Value In " + moCn.GetFldVal("Select HBaseCurCd from Head " + _
                                          "where HCd= '" + ctSelfCmCd + "' and HCoCd= '" + ADC("UoCoCdFr") + "'") + ""
    Else
      Rep.TxtValIn.SetText "Value In " + UCase(ADC("UoCurCdFr"))
    End If
    '62 for value option HUID number not showing
    Rep.hHuidNo.Suppress = True:  Rep.qHuidNo1.Suppress = True
  ElseIf ADC("UoYN") = "N" Then
    Rep.TxtHead.SetText "Design Sales Analysis (Weights)"
    Rep.TxtValTyp.SetText wOpt + " Weights"
    GRP_REP.SetFormula Rep, "wRndOffDecPlaces", "3"
    GRP_REP.SetFormula Rep, "wTotVal", "{rdo.qGrWt}"
    GRP_REP.SetFormula Rep, "wMetVal", "{rdo.qMetWt}"
    GRP_REP.SetFormula Rep, "wDVal", "{rdo.qDWt}"
    GRP_REP.SetFormula Rep, "wCVal", "{rdo.qCWt}"
    GRP_REP.SetFormula Rep, "wXVal", "{rdo.qXWt}"
    GRP_REP.SetFormula Rep, "wSetVal", "0"
    GRP_REP.SetFormula Rep, "wLabVal", "0"
  
'    Rep.h1Tot.SetText wOpt + Chr(13) + "Grs Wt": Rep.h2Tot.SetText wOpt + Chr(13) + "Grs Wt"
'    Rep.h1Met.SetText wOpt + Chr(13) + "Met Wt": Rep.h2Met.SetText wOpt + Chr(13) + "Met Wt"
'    Rep.h1DC.SetText wOpt + Chr(13) + "DC Wt": Rep.h2DC.SetText wOpt + Chr(13) + "DC Wt"
'    Rep.h1X.SetText wOpt + Chr(13) + "Acc Wt": Rep.h2X.SetText wOpt + Chr(13) + "Acc Wt"

    Rep.h1Tot.SetText " Grs Wt": Rep.h2Tot.SetText " Grs Wt"
    Rep.h1Met.SetText " Met Wt": Rep.h2Met.SetText " Met Wt"
    Rep.h1D.SetText " Dia Wt": Rep.h2D.SetText " Dia Wt"
    Rep.h1C.SetText " Cs Wt": Rep.h2C.SetText " Cs Wt"
    Rep.h1X.SetText " Acc Wt": Rep.h2X.SetText " Acc Wt"

    Rep.h1SetVal.SetText " ": Rep.h2SetVal.SetText " "
    Rep.h1LabVal.SetText " ": Rep.h2LabVal.SetText " "
    ' Zubin 213
    Rep.TxtValIn.SetText ""
    '62 for wt option HUID number will display
    Rep.hHuidNo.Suppress = False:  Rep.qHuidNo1.Suppress = False
  End If
  
  ' ****** Manali 3.5.0 - 11/11/08 - Show Picture Option in report
  If ADC("UoYN2") = "N" Or wDetPos = 0 Then
    Rep.PICSEC.Suppress = True
    gb_ShowPic = False
  ElseIf ADC("UoYN2") = "Y" And wDetPos <> 0 Then
    Rep.PICSEC.Suppress = False
    gb_ShowPic = True
  End If
  ' ****** Manali 3.5.0 - 11/11/08 - Show Picture Option in report
  
  ' ***** Manali 3.9.0 - Fr and To Date Scope Report
  Rep.TxtFrDt.SetText ADC("UoDtFr")
  Rep.TxtToDt.SetText ADC("UoDtTo")
  ' ***** Manali 3.9.0 - Fr and To Date Scope Report
  ' ***** Manali 3.10.0 - Date Scope based on Date Selected
  If ADC("UoYN3") = "Y" Then
    Rep.LblFrDt.SetText "Fr Inv Dt :": Rep.LblToDt.SetText "To Inv Dt :"
  Else
    Rep.LblFrDt.SetText "Fr Prn Dt :": Rep.LblToDt.SetText "To Prn Dt :"
  End If
  ' ***** Manali 3.10.0 - Date Scope based on Date Selected
  
  Rep.Text6.Suppress = True
  Rep.Text13.Suppress = True
  Rep.wBalQty1.Suppress = True
  Rep.SumofwBalQty1.Suppress = True
  Rep.SumofwBalQty2.Suppress = True
  Rep.SumofwBalQty3.Suppress = True
  Rep.SumofwBalQty4.Suppress = True
  
  
  
  CRV_REP.DisplayGroupTree = False
  
' ########################  Manoj #### Ver: 2.11.0 #### Date: 26/05/2005  ########################
  Dim wsQryStruct As String
  wsQryStruct = "  Select space(30) as Grp1, space(30) as Grp2, space(30) as Grp3, " + _
    "space(30) as DGrp1, space(30) as DGrp2, space(30) as DGrp3, " + _
    "space(50) as G1Desc, space(50) as G2Desc, space(50) as G3Desc, " + _
    "IdCoCd, IdTc, IdYy, IdChr, IdNo, IdSr, InExpNo, DmCtg, IdDmCd, IdSfx, IdDmSz, IdKt, InDt, " + _
    "InCmCd, IdExpOdTc, IdExpOdYy, IdExpOdChr, IdExpOdNo, IdExpOdSr, DmSalCtg, IdVaCtg, " + _
    "IdQty, IdIFob, IrRmIVal as qMetVal, IrRmIVal as qDCVal, IrRmIVal as qXVal, " + _
    "IrSetIVal as qSetVal, iLiVal as qLabVal, IrRmIWt as qGrWt, IrRmIWt as qMetWt, " + _
    "IrRmIWt As qDCWt, IrRmIWt As qXWt " + _
    "from InvDsg, InvHd, DsgMst, InvRm, RmMst, InvLab where 1= 2"
  Call moGrph.SetGraphProp(Rep, wsQryStruct, wSqlStrg)
' ###########################################  Manoj  ###########################################

  Call DispReport(ADC) '****** Sachin 3.02 - Display Report ******

End Sub

' Zubin 213
Private Sub atxt_ReCalculate(Index As Integer, ByVal IdName As String, ByVal OldVal As Variant, ByVal RowNum As Integer, ByVal SourceName As String, ByVal SourceColName As String)
  Select Case UCase(IdName)
    Case Is = UCase("UoCurCdFr")
    If UCase(ADC("UoInVal")) = "S" Then ADC("UoCurCdFr") = "": Exit Sub
  End Select
End Sub

' ****** Manali 3.5.0 - 21/11/08 - Check Boxes added
Private Sub ChkBoxArr_Click(Index As Integer)
With ChkBoxArr(Index)
  Select Case Index
  Case Is = 0     '3.11.0 - Tag = "BestSeller"
    If .Value = Checked Then
        FraBestSeller.Enabled = True
        ADC("wBSQV") = IIF(Trim(ADC("wBSQV")) <> "", ADC("wBSQV"), "N")
    Else
        FraBestSeller.Enabled = False
    End If
  'uni.4  checking Include SalCtg2,SalCtg3 option
  Case Is = 1
    If .Value = Checked Then ADC("wAddAllSalCtg") = "Y" Else ADC("wAddAllSalCtg") = "N"
  Case Is = 2     'Tag = "ShowPic"
    If .Value = Checked Then ADC("UoYN2") = "Y" Else ADC("UoYN2") = "N"
  End Select
End With
End Sub
' ****** Manali 3.5.0 - 21/11/08 - Check Boxes added

' ****** Manali 3.5.0 - 21/11/08 - Check Boxes added
Private Sub ChkBoxArr_GotFocus(Index As Integer)
  DispMsg ChkBoxArr(Index).ToolTipText, etInfo
End Sub
' ****** Manali 3.5.0 - 21/11/08 - Check Boxes added

Private Sub CmdSc_Click()
  If FraSC.Visible = False Then
    If CRV_REP.Visible = True Then Exit Sub
    FraSC.Enabled = True
    FraSC.Visible = True
    TAB_REP.Enabled = False
    ADC.AllowSave = False
    FraSC.ZOrder
  Else
    TAB_REP.Enabled = True
    ADC.AllowSave = True
    FraSC.Enabled = False
    FraSC.Visible = False
  End If
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
  Set moCn = ADC.Connection    '*** (Bef speed) bef 24/11/06
  
  
  '****** Sachin 2.14.0 - [23-11-2006] - [Multi-Company] ******
  ws_MultiCoMod = IIF(moCn.GetFldVal("Select hImpComLoc From Head Where HCd='" + ctSelfCmCd + "' And HCoCd='" + ctSelfCoCd + "'") = "Y", True, False)
  gb_CoCdFor = False
  
  ADC.Child = "GrdDsgAna"
  Set GrdDsgAna.PrevCtl = ADC
  Set GrdDsgAna.NextCtl = ADC
  GrdDsgAna.HotKey = "A"
    
  
  '****** Sachin 2.14.0 - [23-11-2006] - [Multi-Company] ******

  '****** Sachin 2.14.0 - [23-11-2006] - [Multi-Company] ******
  If ws_MultiCoMod = True Then
    Set ADC.FirNKeyCtl = ADC("UoCoCdFr")
  Else
    Set ADC.FirNKeyCtl = ADC("UOInTcFr")
  End If
  '****** Sachin 2.14.0 - [23-11-2006] - [Multi-Company] ******

  Call InitProp(Me)
  
  '*** Parameters Used For Setting Properties Of UsrOpt Fields ***
' ########################  Manoj #### Ver: 2.11.0 #### Date: 26/05/2005  ########################
  Set moGrph = New EmrClsGrph
  Call moGrph.SetFormDets(Me, moCn)    '*** (Bef speed) bef 24/11/06
  'Call moGrph.SetFormDets(Me)  '*** (Jen speed) 24/11/06
' ###########################################  Manoj  ###########################################

  '****** Sachin 2.14.0 - [23-11-2006] - [Multi-Company] ******
  If ws_MultiCoMod = True Then
      ADC("UOCOCDFR").Enabled = True: ADC("UOCOCDTO").Visible = True: ADC("UOCOCDSEL").Visible = True
  Else
      ADC("UOCOCDFR").Enabled = False: ADC("UOCOCDTO").Visible = False: ADC("UOCOCDSEL").Visible = False
  End If
  '****** Sachin 2.14.0 - [23-11-2006] - [Multi-Company] ******
  
  Call SetGroupSort
End Sub

Private Sub ATXT_FldChange(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String, ByVal pv_NewValue As Variant)
  gb_CoCdFor = False  '****** Sachin 2.14.0 - [Multi-Company]
  Call SetProp(Me, IdName, Valid, Cancel, ErrMsg)
  
  Select Case Trim(IdName)
  Case Is = UCase("WBSQV")      '3.11.0
    Cancel = Not moCn.RecSeek(" Select PMCd from Param where PTyp= 'YN' and PMCd = '" + pv_NewValue + "' ")
    If Cancel = True Then ErrMsg = "Invalid Best Seller QV Option. Enter(Y/N)": Exit Sub
  End Select
  
End Sub
Private Sub ATXT_FldWhen(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String)
' ########################  Manoj #### Ver: 2.11.0 #### Date: 26/05/2005  ########################
  If UCase$(IdName) = UCase$("UoGrphTyp") Then gs_GrphGrpNo = moGrph.SetGrphGrpNo
' ###########################################  Manoj  ###########################################

  gb_CoCdFor = False  '****** Sachin 2.14.0 - [Multi-Company]

  ' Zubin 213
  Select Case UCase(IdName)
  Case Is = UCase("UoCurCdFr")
    If UCase(ADC("UoInVal")) = "S" Then Cancel = True: Exit Sub
  Case Is = UCase("wBSYy")      '3.11.0
    Call HlpList.PMCd("YY")
  Case Is = UCase("wBSQv")      '3.11.0
    Call HlpList.PMCd("YN")
  End Select
  ' Zubin 213

  SetProp Me, IdName, When
End Sub


Private Sub GrdDsgAna_ColChange(ByVal RowNum As Integer, ByVal ColName As String, ByVal NewValue As Variant, Cancel As Boolean, ErrMsg As String)
' ****** Manali 3.5.0 - 07/11/08 - Dsg Anlaysis Scope added
  With GrdDsgAna
    Select Case UCase(ColName)
    Case Is = UCase("UdAnaSr")
      Cancel = Not moCn.RecSeek("Select PMCd From Param Where " + _
                                "PTyp= 'DAANAFLD' And PMCd = '" + NewValue + "'")
      ErrMsg = "Invalid Analysis Sr.": Exit Sub
    End Select
  End With
End Sub

Private Sub GrdDsgAna_ColWhen(ByVal RowNum As Integer, ByVal ColName As String, Cancel As Boolean, ErrMsg As String, ByVal Action As MwfCtl.en_fgAction)
' ****** Manali 3.5.0 - 07/11/08 - Dsg Anlaysis Scope added
  Select Case ColName
  Case Is = UCase("UdAnaSr")
    Call HlpList.PMCd("DAANAFLD")
  Case Is = UCase("UdAnaCdFr"), UCase("UdAnaCdTo")
    If moCn.GetFldVal("Select PValue from Param where PTyp= 'DAANAFLD' and " + _
       "PMCd= '" + GrdDsgAna.Value(RowNum, "UdAnaSr") + "' ") = "N" Then _
       Cancel = True: ErrMsg = "Cannot Enter Code For This Analysis Field": Exit Sub
  
    Call HlpList.PSCd("DAANACD", GrdDsgAna.Value(RowNum, "UdAnaSr"))
  End Select
End Sub

Private Sub GrdDsgAna_InitKey(ByVal RowNum As Integer)
' ****** Manali 3.5.0 - 07/11/08 - Dsg Anlaysis Scope added
  '*** Initialise Key Fields
      'The fields of the Base table which are not existing in the grid are stored in the
      'table using Grid_InitKey
  
  With GrdDsgAna
    If .IsNew(RowNum) Then
      .Store "UdUsrCd", ADC("UoUsrCd")
      .Store "UdMnuCd", ADC("UoMnuCd")
    End If
  End With

End Sub

Private Sub GrdDsgAna_ReCalc(ByVal RowNum As Integer, ByVal ColName As String, ByVal OldVal As Variant, ByVal SourceName As String, ByVal SourceColName As String)
' ****** Manali 3.5.0 - 07/11/08 - Dsg Anlaysis Scope added
  With GrdDsgAna
    Select Case UCase(ColName)
    Case Is = UCase("UdAnaCdFr")
      If .Mode = fgmnorm Then .Value(RowNum, "UdAnaCdFr") = ""
    Case Is = UCase("UdAnaCdTo")
      If .Mode = fgmnorm Then .Value(RowNum, "UdAnaCdTo") = ""
    End Select
  End With
End Sub

Private Sub GrdDsgAna_SetRecSource()
  '*** Set the Record Source of the Grid GrdDsgAna
  ' ****** Manali 3.5.0 - 07/11/08 - Dsg Anlaysis Scope added
  GrdDsgAna.RecSource = "Select * from UsrOptDsgAna where UdUsrCd= '" + ADC("UoUsrCd") + "' and " + _
                        "UdMnuCd= '" + ADC("UoMnuCd") + "'"
End Sub

Private Sub GrdDsgAna_Validate(Cancel As Boolean)
' ****** Manali 3.5.0 - 07/11/08 - Dsg Anlaysis Scope added
  Cancel = GrdDsgAna.Validate
End Sub
' ****** Manali 3.5.0 - 03/11/08 - Radio Buttons Added
Private Sub OptCust_Click(Index As Integer)
Select Case Index
Case Is = 0
  ADC("UoYN1") = "Y"
Case Is = 1
  ADC("UoYN1") = "N"
End Select
End Sub
' ****** Manali 3.5.0 - 03/11/08 - Radio Buttons Added

' ****** Manali 3.5.0 - 03/11/08 - Radio Buttons Added
Private Sub OptCust_GotFocus(Index As Integer)
  DispMsg FraOptCust.ToolTipText, etInfo
End Sub
' ****** Manali 3.5.0 - 03/11/08 - Radio Buttons Added

' ****** Manali 3.5.0 - 25/11/08 - Option for InDt / InPrnDt added
Private Sub OptInvDt_Click(Index As Integer)
Select Case Index
Case Is = 0
  ADC("UoYN3") = "Y"
Case Is = 1
  ADC("UoYN3") = "N"
End Select
End Sub

' ****** Manali 3.5.0 - 25/11/08 - Option for InDt / InPrnDt added
Private Sub OptInvDt_GotFocus(Index As Integer)
  DispMsg FraOptDt.ToolTipText, etInfo
End Sub

' ****** Manali 3.5.0 - 03/11/08 - Radio Buttons Added
Private Sub OptValWt_Click(Index As Integer)
Select Case Index
Case Is = 0
  ADC("UoYN") = "Y"
Case Is = 1
  ADC("UoYN") = "N"
End Select
End Sub
' ****** Manali 3.5.0 - 03/11/08 - Radio Buttons Added

Private Sub OptValWt_GotFocus(Index As Integer)
  DispMsg FraOptValWt.ToolTipText, etInfo
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

'****** Sachin 2.14.0 - [25-11-2006] - [Multi-Company] ******
  If ws_MultiCoMod = False Then _
      ADC("UoCoCdFr") = gs_CoCd: ADC("UoCoCdTo") = gs_CoCd: ADC("UoCoCdSel") = "": ADC("UoCoCdFr").Enabled = False
  '****** Sachin 2.14.0 - [25-11-2006] - [Multi-Company] ******

' ########################  Manoj #### Ver: 2.11.0 #### Date: 26/05/2005  ########################
  If ADC("UoShowGrph") = "" Then ADC("UoShowGrph") = "N"
  Call moGrph.NonKeyWhen
' ###########################################  Manoj  ###########################################

' ****** Manali 3.5.0 - 05/11/08 - Radio Buttons Added
If ADC("UoYN") = "Y" Then OptValWt(0).Value = True
If ADC("UoYN") = "N" Then OptValWt(1).Value = True
If ADC("UoYN1") = "Y" Then OptCust(0).Value = True
If ADC("UoYN1") = "N" Then OptCust(1).Value = True
If ADC("UoYN3") = "Y" Then OptInvDt(0).Value = True
If ADC("UoYN3") = "N" Then OptInvDt(1).Value = True
' ****** Manali 3.5.0 - 05/11/08 - Radio Buttons Added

' ****** Manali 3.5.0 - 21/11/08 - Check Boxes added
If ADC("UoYN2") = "" Then ADC("UoYN2") = "N"

If ADC("UoYN2") = "Y" Then ChkBoxArr(2).Value = Checked Else ChkBoxArr(2).Value = Unchecked
' ****** Manali 3.5.0 - 21/11/08 - Check Boxes added
'uni.4 default value for Include SalCtg2,SalCtg3 is Yes
ChkBoxArr(1).Value = Checked
ADC("wAddAllSalCtg") = "Y"
  
End Sub
Private Sub DispCoNm()
'Rep.wCoCd.SetText gs_CoNm
  Dim ws_HName As String
  
  ws_HName = moCn.GetFldVal("Select HName from Head where HCoCd ='" + IIF(ws_MultiCoMod = True, gs_CoCd, ADC("UOCOCDFR")) + "' and " + _
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
  
  '****** Sachin 2.14.0 - [25-11-2006] - [Multi-Company] ******
  If ws_MultiCoMod = True Then
      Rep.TxtCoCdToLbl.Suppress = False: Rep.TxtCoCdSelLbl.Suppress = False
      Rep.TxtCoCdFr.SetText ADC("UOCOCDFR"): Rep.TxtCoCdTo.SetText ADC("UOCOCDTO"): Rep.TxtCoCdSel.SetText ADC("UOCOCDSEL")
  Else
      Rep.TxtCoCdTo.Suppress = True: Rep.TxtCoCdSel.Suppress = True
      Rep.TxtCoCdToLbl.Suppress = True: Rep.TxtCoCdSelLbl.Suppress = True
  End If
  '****** Sachin 2.14.0 - [25-11-2006] - [Multi-Company] ******
  
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
  Rep.TxtRepUsrDtTime.SetText Trim(ADC.UsrCd) + " / " + ws_Dt + " / " + ws_Tm
  
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
    Set Rep = Nothing
    '*** (09/08/05)
End Sub

' ########################  Manoj #### Ver: 2.11.0 #### Date: 26/05/2005  ########################
Private Sub GRP_REP_LostFocus()
  Call moGrph.SetDefaVal
End Sub
Private Sub ADC_NonKeyValid(pr_Cancel As Boolean, pr_ErrMsg As String)
  moGrph.NonKeyValid pr_Cancel, pr_ErrMsg
End Sub
Private Sub TAB_GRPH_Click(PreviousTab As Integer)
    FraTabGrph(TAB_GRPH.Tab).Enabled = True
    FraTabGrph(PreviousTab).Enabled = False
End Sub
Private Sub ChkGrph_Click()
  If ChkGrph.Value = vbChecked Then
    ADC("UoShowGrph") = "Y"
    FraGrphDet.Enabled = True
    moGrph.SetDefaVal
  ElseIf ChkGrph.Value = vbUnchecked Then
    ADC("UoShowGrph") = "N"
    FraGrphDet.Enabled = False
  End If
End Sub
Private Sub cboGrphValue_Click()
  ADC("UoGrphValInd") = cboGrphValue.ListIndex
End Sub
' ###########################################  Manoj  ###########################################
