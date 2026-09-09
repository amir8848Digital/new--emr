VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{AB2CC32A-E652-42A2-B341-26ECE02F6386}#2.1#0"; "mwfctlforspeed9.ocx"
Begin VB.Form EmrFrpPrd 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Production Report"
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
      Left            =   11670
      TabIndex        =   88
      TabStop         =   0   'False
      Top             =   9600
      Visible         =   0   'False
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
   End
   Begin MwfCtl.MWCTL_DST DST_REP 
      Height          =   570
      Left            =   4455
      TabIndex        =   89
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
      Left            =   5625
      TabIndex        =   97
      Top             =   9630
      Width           =   3555
      Begin MwfCtl.MWCTL_MED aktxt 
         Height          =   345
         Index           =   0
         Left            =   210
         TabIndex        =   98
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
         TabIndex        =   99
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
      Left            =   6840
      TabIndex        =   96
      Top             =   9645
      Visible         =   0   'False
      Width           =   855
      Begin CRVIEWER9LibCtl.CRViewer9 CRV_REP 
         Height          =   495
         Left            =   480
         TabIndex        =   93
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
      TabIndex        =   90
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
      TabIndex        =   92
      Top             =   60
      Width           =   15255
      Begin TabDlg.SSTab TAB_REP 
         Height          =   9465
         Left            =   15
         TabIndex        =   91
         Top             =   30
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   16695
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "&Data Range"
         TabPicture(0)   =   "EmrFrpPrd.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "fra_tabrep(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "&Group Sort"
         TabPicture(1)   =   "EmrFrpPrd.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "fra_tabrep(1)"
         Tab(1).ControlCount=   1
         Begin VB.Frame fra_tabrep 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   9070
            Index           =   1
            Left            =   -74940
            TabIndex        =   95
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
               TabIndex        =   132
               Top             =   1050
               Width           =   11160
               Begin VB.Frame FraGrphDet 
                  BorderStyle     =   0  'None
                  Enabled         =   0   'False
                  Height          =   7425
                  Left            =   30
                  TabIndex        =   133
                  Top             =   450
                  Width           =   11095
                  Begin VB.ComboBox cboGrphValue 
                     Height          =   315
                     ItemData        =   "EmrFrpPrd.frx":0038
                     Left            =   2100
                     List            =   "EmrFrpPrd.frx":004E
                     Sorted          =   -1  'True
                     Style           =   2  'Dropdown List
                     TabIndex        =   72
                     ToolTipText     =   "Select Graph Value"
                     Top             =   915
                     Width           =   2865
                  End
                  Begin TabDlg.SSTab TAB_GRPH 
                     Height          =   5760
                     Left            =   -30
                     TabIndex        =   76
                     Top             =   1670
                     Width           =   11130
                     _ExtentX        =   19632
                     _ExtentY        =   10160
                     _Version        =   393216
                     Tabs            =   2
                     TabsPerRow      =   2
                     TabHeight       =   520
                     TabCaption(0)   =   "T&itles"
                     TabPicture(0)   =   "EmrFrpPrd.frx":0081
                     Tab(0).ControlEnabled=   -1  'True
                     Tab(0).Control(0)=   "FraTabGrph(0)"
                     Tab(0).Control(0).Enabled=   0   'False
                     Tab(0).ControlCount=   1
                     TabCaption(1)   =   "Top &N"
                     TabPicture(1)   =   "EmrFrpPrd.frx":009D
                     Tab(1).ControlEnabled=   0   'False
                     Tab(1).Control(0)=   "FraTabGrph(1)"
                     Tab(1).ControlCount=   1
                     Begin VB.Frame FraTabGrph 
                        BorderStyle     =   0  'None
                        Height          =   5370
                        Index           =   0
                        Left            =   60
                        TabIndex        =   142
                        Top             =   360
                        Width           =   11040
                        Begin MwfCtl.MWCTL_MED atxt 
                           Height          =   285
                           Index           =   58
                           Left            =   1680
                           TabIndex        =   77
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
                           TabIndex        =   78
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
                           TabIndex        =   80
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
                           TabIndex        =   81
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
                           TabIndex        =   82
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
                           TabIndex        =   79
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
                           TabIndex        =   148
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
                           TabIndex        =   147
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
                           TabIndex        =   146
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
                           TabIndex        =   145
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
                           TabIndex        =   144
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
                           TabIndex        =   143
                           Top             =   435
                           Width           =   1170
                        End
                     End
                     Begin VB.Frame FraTabGrph 
                        BorderStyle     =   0  'None
                        Height          =   5370
                        Index           =   1
                        Left            =   -74940
                        TabIndex        =   134
                        Top             =   360
                        Width           =   11040
                        Begin MwfCtl.MWCTL_MED atxt 
                           Height          =   285
                           Index           =   64
                           Left            =   2310
                           TabIndex        =   84
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
                           TabIndex        =   85
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
                           TabIndex        =   83
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
                           TabIndex        =   86
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
                           TabIndex        =   87
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
                           TabIndex        =   141
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
                           TabIndex        =   140
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
                           TabIndex        =   139
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
                           TabIndex        =   138
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
                           TabIndex        =   137
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
                           TabIndex        =   136
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
                           TabIndex        =   135
                           Top             =   1755
                           Width           =   615
                        End
                     End
                  End
                  Begin MwfCtl.MWCTL_MED atxt 
                     Height          =   285
                     Index           =   56
                     Left            =   2100
                     TabIndex        =   68
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
                     TabIndex        =   69
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
                     TabIndex        =   70
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
                     TabIndex        =   71
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
                     TabIndex        =   149
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
                     Index           =   72
                     Left            =   7620
                     TabIndex        =   73
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
                     Index           =   73
                     Left            =   7620
                     TabIndex        =   75
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
                     Index           =   74
                     Left            =   2100
                     TabIndex        =   74
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
                     TabIndex        =   158
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
                     TabIndex        =   157
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
                     TabIndex        =   156
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
                     TabIndex        =   155
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
                     TabIndex        =   154
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
                     TabIndex        =   153
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
                     TabIndex        =   152
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
                     TabIndex        =   151
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
                     TabIndex        =   150
                     Top             =   0
                     Width           =   11325
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
                  TabIndex        =   67
                  Top             =   0
                  UseMaskColor    =   -1  'True
                  Width           =   1875
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   75
                  Left            =   2010
                  TabIndex        =   159
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
               TabIndex        =   66
               Top             =   360
               Width           =   9945
               _ExtentX        =   17542
               _ExtentY        =   635
            End
         End
         Begin VB.Frame fra_tabrep 
            BorderStyle     =   0  'None
            Height          =   9070
            Index           =   0
            Left            =   60
            TabIndex        =   94
            Top             =   360
            Width           =   15105
            Begin VB.CheckBox ChkBoxWrkr 
               Alignment       =   1  'Right Justify
               Caption         =   "Worker Change Adds to BagQty YN"
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
               Left            =   90
               TabIndex        =   165
               Tag             =   "ShowPic"
               ToolTipText     =   "Worker Change Adds to BagQty YN"
               Top             =   7200
               Width           =   4305
            End
            Begin VB.Frame FraLoc 
               BorderStyle     =   0  'None
               Height          =   285
               Left            =   1440
               TabIndex        =   168
               Top             =   840
               Width           =   4305
               Begin VB.OptionButton OptLoc 
                  Caption         =   "Production Loc"
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
                  Left            =   80
                  TabIndex        =   170
                  ToolTipText     =   "Select Production or Process"
                  Top             =   0
                  Value           =   -1  'True
                  Width           =   2085
               End
               Begin VB.OptionButton OptLoc 
                  Caption         =   "Process loc"
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
                  Left            =   2780
                  TabIndex        =   169
                  ToolTipText     =   "Select Production or Process"
                  Top             =   0
                  Width           =   1545
               End
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Show Picture"
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
               TabIndex        =   164
               Tag             =   "ShowPic"
               ToolTipText     =   "Check To Show Picture"
               Top             =   6840
               Width           =   1635
            End
            Begin VB.Frame FraPtsVal 
               BorderStyle     =   0  'None
               Height          =   285
               Left            =   1530
               TabIndex        =   162
               Top             =   6330
               Visible         =   0   'False
               Width           =   5385
               Begin VB.OptionButton OptPtsVal 
                  Caption         =   "Rs. Val"
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
                  Left            =   2670
                  TabIndex        =   64
                  Top             =   0
                  Width           =   1365
               End
               Begin VB.OptionButton OptPtsVal 
                  Caption         =   "Prd Pts"
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
                  Left            =   30
                  TabIndex        =   63
                  Top             =   0
                  Width           =   1365
               End
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   1
               Left            =   4215
               TabIndex        =   7
               ToolTipText     =   "Enter To Process Location"
               Top             =   1200
               Width           =   1830
               _ExtentX        =   3228
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UOISSLOCTO"
               IdName          =   "UOISSLOCTO"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   2
               Left            =   6900
               TabIndex        =   8
               ToolTipText     =   "Enter Prrocess Location Selection"
               Top             =   1185
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOISSLOCSEL"
               IdName          =   "UOISSLOCSEL"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   0
               Left            =   1515
               TabIndex        =   6
               ToolTipText     =   "Enter From Process Location"
               Top             =   1185
               Width           =   1830
               _ExtentX        =   3228
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UOISSLOCFR"
               IdName          =   "UOISSLOCFR"
               Data            =   "ABCDW"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   28
               Left            =   4215
               TabIndex        =   16
               ToolTipText     =   "Enter To Design Category"
               Top             =   2040
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
               Index           =   29
               Left            =   6900
               TabIndex        =   17
               ToolTipText     =   "Enter Design Category Selection"
               Top             =   2040
               Width           =   7950
               _ExtentX        =   14023
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
               Left            =   1515
               TabIndex        =   15
               ToolTipText     =   "Enter From Design Category"
               Top             =   2040
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
               Left            =   4215
               TabIndex        =   22
               ToolTipText     =   "Enter To Design Code"
               Top             =   2610
               Width           =   2385
               _ExtentX        =   4207
               _ExtentY        =   503
               MaxLength       =   15
               DataField       =   "UODMCDTO"
               IdName          =   "UODMCDTO"
               CmpStr          =   "BOdDmCd <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   32
               Left            =   6900
               TabIndex        =   23
               ToolTipText     =   "Enter Design Code Selection"
               Top             =   2610
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UODMCDSEL"
               IdName          =   "UODMCDSEL"
               CmpStr          =   "BOdDmCd In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   33
               Left            =   1515
               TabIndex        =   21
               ToolTipText     =   "Enter From Design Code"
               Top             =   2610
               Width           =   2385
               _ExtentX        =   4207
               _ExtentY        =   503
               MaxLength       =   15
               DataField       =   "UODMCDFR"
               IdName          =   "UODMCDFR"
               CmpStr          =   "BOdDmCd >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   40
               Left            =   4215
               TabIndex        =   48
               ToolTipText     =   "Enter To Customer"
               Top             =   4605
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UOCMCDTO"
               IdName          =   "UOCMCDTO"
               CmpStr          =   "OmCmCd <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   41
               Left            =   6900
               TabIndex        =   49
               ToolTipText     =   "Enter Customer Selection"
               Top             =   4605
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOCMCDSEL"
               IdName          =   "UOCMCDSEL"
               CmpStr          =   "OmCmCd In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   42
               Left            =   1515
               TabIndex        =   47
               ToolTipText     =   "Enter From Customer"
               Top             =   4605
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UOCMCDFR"
               IdName          =   "UOCMCDFR"
               CmpStr          =   "OmCmCd >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   43
               Left            =   4215
               TabIndex        =   51
               ToolTipText     =   "Enter To Purchase Order "
               Top             =   4890
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
               Index           =   44
               Left            =   6900
               TabIndex        =   52
               ToolTipText     =   "Enter Purchase Order Selection"
               Top             =   4890
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOPONOSEL"
               IdName          =   "UOPONOSEL"
               CmpStr          =   "OmPoNo In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   45
               Left            =   1515
               TabIndex        =   50
               ToolTipText     =   "Enter From Purchase Order "
               Top             =   4890
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
               Index           =   46
               Left            =   4215
               TabIndex        =   28
               ToolTipText     =   "Enter To Voucher Date"
               Top             =   3180
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   3
               DataField       =   "UODTTO"
               IdName          =   "UODTTO"
               CmpStr          =   "TDt <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   48
               Left            =   1515
               TabIndex        =   27
               ToolTipText     =   "Enter From Voucher Date"
               Top             =   3180
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   3
               DataField       =   "UODTFR"
               IdName          =   "UODTFR"
               CmpStr          =   "TDt >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   5
               Left            =   4995
               TabIndex        =   43
               ToolTipText     =   "Enter To Order Year"
               Top             =   4035
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               MaxLength       =   2
               DataField       =   "UOOMYYTO"
               IdName          =   "UOOMYYTO"
               CmpStr          =   "OdYy <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   20
               Left            =   2265
               TabIndex        =   40
               ToolTipText     =   "Enter From Order Year"
               Top             =   4035
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               MaxLength       =   2
               DataField       =   "UOOMYYFR"
               IdName          =   "UOOMYYFR"
               CmpStr          =   "OdYy >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   10
               Left            =   1515
               TabIndex        =   39
               ToolTipText     =   "Enter From Order Tc"
               Top             =   4035
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataField       =   "UOOMTCFR"
               IdName          =   "UOOMTCFR"
               CmpStr          =   "OdTc >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   13
               Left            =   2685
               TabIndex        =   41
               ToolTipText     =   "Enter From Order Character"
               Top             =   4035
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOOMCHRFR"
               IdName          =   "UOOMCHRFR"
               CmpStr          =   "OdChr >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   17
               Left            =   4215
               TabIndex        =   42
               ToolTipText     =   "Enter To Order Tc"
               Top             =   4035
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataField       =   "UOOMTCTO"
               IdName          =   "UOOMTCTO"
               CmpStr          =   "OdTc <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   23
               Left            =   5415
               TabIndex        =   44
               ToolTipText     =   "Enter To Order Character"
               Top             =   4035
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOOMCHRTO"
               IdName          =   "UOOMCHRTO"
               CmpStr          =   "OdChr <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   53
               Left            =   1515
               TabIndex        =   45
               ToolTipText     =   "Enter From Order Number"
               Top             =   4320
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOOMNOFR"
               IdName          =   "UOOMNOFR"
               CmpStr          =   "OdNo >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   54
               Left            =   4215
               TabIndex        =   46
               ToolTipText     =   "Enter To Order Number"
               Top             =   4320
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOOMNOTO"
               IdName          =   "UOOMNOTO"
               CmpStr          =   "OdNo <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   3
               Left            =   4215
               TabIndex        =   10
               ToolTipText     =   "Enter To Sub Process"
               Top             =   1470
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UOSUBPRCTO"
               IdName          =   "UOSUBPRCTO"
               CmpStr          =   "TpSubPrc <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   4
               Left            =   6900
               TabIndex        =   11
               ToolTipText     =   "Enter Sub Process Selection"
               Top             =   1470
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOSUBPRCSEL"
               IdName          =   "UOSUBPRCSEL"
               CmpStr          =   "TpSubPrc In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   6
               Left            =   1515
               TabIndex        =   9
               ToolTipText     =   "Enter From Sub Process"
               Top             =   1470
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UOSUBPRCFR"
               IdName          =   "UOSUBPRCFR"
               CmpStr          =   "TpSubPrc >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   7
               Left            =   4215
               TabIndex        =   13
               ToolTipText     =   "Enter To By Worker"
               Top             =   1755
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UOWRKTO"
               IdName          =   "UOWRKTO"
               CmpStr          =   "TpByWrk <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   8
               Left            =   6900
               TabIndex        =   14
               ToolTipText     =   "Enter By Worker Selection"
               Top             =   1755
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOWRKSEL"
               IdName          =   "UOWRKSEL"
               CmpStr          =   "TpByWrk In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   9
               Left            =   1515
               TabIndex        =   12
               ToolTipText     =   "Enter From By Worker"
               Top             =   1755
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UOWRKFR"
               IdName          =   "UOWRKFR"
               Data            =   "ABCDW"
               CmpStr          =   "TpByWrk >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   11
               Left            =   4215
               TabIndex        =   19
               ToolTipText     =   "Enter To Design Production Category"
               Top             =   2325
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataType        =   4
               DataField       =   "UODMPRDCTGTO"
               IdName          =   "UODMPRDCTGTO"
               CmpStr          =   "DmPrdCtg <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   12
               Left            =   6900
               TabIndex        =   20
               ToolTipText     =   "Enter Design Production Category Selection"
               Top             =   2325
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataType        =   4
               DataField       =   "UODMPRDCTGSEL"
               IdName          =   "UODMPRDCTGSEL"
               CmpStr          =   "DmPrdCtg In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   14
               Left            =   1515
               TabIndex        =   18
               ToolTipText     =   "Enter From Design Production Category"
               Top             =   2325
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataType        =   4
               DataField       =   "UODMPRDCTGFR"
               IdName          =   "UODMPRDCTGFR"
               CmpStr          =   "DmPrdCtg >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   15
               Left            =   4215
               TabIndex        =   25
               ToolTipText     =   "Enter To Karat"
               Top             =   2895
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataField       =   "UOKTTO"
               IdName          =   "UOKTTO"
               CmpStr          =   "OdKt <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   16
               Left            =   6900
               TabIndex        =   26
               ToolTipText     =   "Enter Karat Selection"
               Top             =   2895
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOKTSEL"
               IdName          =   "UOKTSEL"
               CmpStr          =   "OdKt In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   18
               Left            =   1515
               TabIndex        =   24
               ToolTipText     =   "Enter From Karat"
               Top             =   2895
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataField       =   "UOKTFR"
               IdName          =   "UOKTFR"
               Data            =   "ABCDW"
               CmpStr          =   "OdKt >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   24
               Left            =   4995
               TabIndex        =   33
               ToolTipText     =   "Enter To Transaction Year"
               Top             =   3465
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               MaxLength       =   2
               DataField       =   "UOTYYTO"
               IdName          =   "UOTYYTO"
               CmpStr          =   "TpYy <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   27
               Left            =   2295
               TabIndex        =   30
               ToolTipText     =   "Enter From Transaction Year"
               Top             =   3465
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               MaxLength       =   2
               DataField       =   "UOTYYFR"
               IdName          =   "UOTYYFR"
               CmpStr          =   "TpYy >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   19
               Left            =   1515
               TabIndex        =   29
               ToolTipText     =   "Enter From Transaction Tc"
               Top             =   3465
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataField       =   "UOTTCFR"
               IdName          =   "UOTTCFR"
               CmpStr          =   "TpTc >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   49
               Left            =   2715
               TabIndex        =   31
               ToolTipText     =   "Enter From Transaction Character"
               Top             =   3465
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOTCHRFR"
               IdName          =   "UOTCHRFR"
               CmpStr          =   "TpChr >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   50
               Left            =   4215
               TabIndex        =   32
               ToolTipText     =   "Enter To Transaction Tc"
               Top             =   3465
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataField       =   "UOTTCTO"
               IdName          =   "UOTTCTO"
               CmpStr          =   "TpTc <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   51
               Left            =   5415
               TabIndex        =   34
               ToolTipText     =   "Enter To Transaction Character"
               Top             =   3465
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOTCHRTO"
               IdName          =   "UOTCHRTO"
               CmpStr          =   "TpChr <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   21
               Left            =   5115
               TabIndex        =   38
               ToolTipText     =   "Enter To Transaction Serial"
               Top             =   3750
               Width           =   525
               _ExtentX        =   926
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOTDSRTO"
               IdName          =   "UOTDSRTO"
               CmpStr          =   "TpSr <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   22
               Left            =   2415
               TabIndex        =   36
               ToolTipText     =   "Enter From Transaction Serial"
               Top             =   3750
               Width           =   525
               _ExtentX        =   926
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOTDSRFR"
               IdName          =   "UOTDSRFR"
               CmpStr          =   "TpSr >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   25
               Left            =   1515
               TabIndex        =   35
               ToolTipText     =   "Enter From Transaction Number"
               Top             =   3750
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOTNOFR"
               IdName          =   "UOTNOFR"
               CmpStr          =   "TpNo >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   55
               Left            =   4215
               TabIndex        =   37
               ToolTipText     =   "Enter To Transaction Number"
               Top             =   3750
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOTNOTO"
               IdName          =   "UOTNOTO"
               CmpStr          =   "TpNo <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   26
               Left            =   4215
               TabIndex        =   57
               ToolTipText     =   "Enter Raw Material Sub Category To Range"
               Top             =   5460
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataField       =   "UORMSCTGTO"
               IdName          =   "UORMSCTGTO"
               CmpStr          =   "TpRmSCtg <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   34
               Left            =   6900
               TabIndex        =   58
               ToolTipText     =   "Enter Raw Material Sub Category Selection"
               Top             =   5460
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UORMSCTGSEL"
               IdName          =   "UORMSCTGSEL"
               CmpStr          =   "TpRmSCtg In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   35
               Left            =   1515
               TabIndex        =   53
               ToolTipText     =   "Enter From Raw Material Category"
               Top             =   5175
               Width           =   465
               _ExtentX        =   820
               _ExtentY        =   503
               MaxLength       =   2
               DataField       =   "UORMCTGFR"
               IdName          =   "UORMCTGFR"
               CmpStr          =   "TpRmCtg >= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   36
               Left            =   1515
               TabIndex        =   56
               ToolTipText     =   "Enter Raw Material Sub Category From Range"
               Top             =   5460
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataField       =   "UORMSCTGFR"
               IdName          =   "UORMSCTGFR"
               Data            =   "ABCDW"
               CmpStr          =   "TpRmSCtg >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   37
               Left            =   4215
               TabIndex        =   54
               ToolTipText     =   "Enter To Raw Material Category"
               Top             =   5175
               Width           =   465
               _ExtentX        =   820
               _ExtentY        =   503
               MaxLength       =   2
               DataField       =   "UORMCTGTO"
               IdName          =   "UORMCTGTO"
               CmpStr          =   "TpRmCtg <= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   38
               Left            =   6900
               TabIndex        =   55
               ToolTipText     =   "Enter Raw Material Category Selection"
               Top             =   5175
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UORMCTGSEL"
               IdName          =   "UORMCTGSEL"
               CmpStr          =   "TpRmCtg In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   39
               Left            =   4215
               TabIndex        =   60
               ToolTipText     =   "Enter To Setting Type"
               Top             =   5745
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UOLSCDTO"
               IdName          =   "UOLSCDTO"
               CmpStr          =   "TpSetTyp <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   47
               Left            =   1515
               TabIndex        =   59
               ToolTipText     =   "Enter From Setting Type"
               Top             =   5745
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UOLSCDFR"
               IdName          =   "UOLSCDFR"
               CmpStr          =   "TpSetTyp >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   52
               Left            =   1515
               TabIndex        =   0
               ToolTipText     =   "Enter Company Code"
               Top             =   280
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOCOCDFR"
               IdName          =   "UOCOCDFR"
               CmpStr          =   "TdCoCd >= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   76
               Left            =   1515
               TabIndex        =   61
               ToolTipText     =   "Specify Bag Type or ' ' for All"
               Top             =   6030
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               DataType        =   4
               DataField       =   "UOBAGTYP"
               IdName          =   "UOBAGTYP"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   77
               Left            =   4215
               TabIndex        =   62
               ToolTipText     =   "Enter Bag Character"
               Top             =   6030
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataField       =   "UOBCHRFR"
               IdName          =   "UOBCHRFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   78
               Left            =   1515
               TabIndex        =   65
               ToolTipText     =   "Enter ""Y"" for Production Points and ""N"" for Rs. Value"
               Top             =   6075
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
               Index           =   79
               Left            =   4215
               TabIndex        =   1
               ToolTipText     =   "Enter To Company Code"
               Top             =   280
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOCOCDTO"
               IdName          =   "UOCOCDTO"
               CmpStr          =   "TdCoCd <= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   80
               Left            =   6900
               TabIndex        =   2
               ToolTipText     =   "Enter Design Code Selection"
               Top             =   280
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOCOCDSEL"
               IdName          =   "UOCOCDSEL"
               CmpStr          =   "TdCoCd In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   81
               Left            =   4215
               TabIndex        =   4
               ToolTipText     =   "Enter To Parent Code"
               Top             =   585
               Width           =   1530
               _ExtentX        =   2699
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UOPRNTCDTO"
               IdName          =   "UOPRNTCDTO"
               CmpStr          =   "LocPrntCd <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   82
               Left            =   6900
               TabIndex        =   5
               ToolTipText     =   "Enter Parent Code Selection"
               Top             =   580
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOPRNTCDSEL"
               IdName          =   "UOPRNTCDSEL"
               CmpStr          =   "LocPrntCd In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   83
               Left            =   1515
               TabIndex        =   3
               ToolTipText     =   "Enter From Parent Code"
               Top             =   580
               Width           =   1530
               _ExtentX        =   2699
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UOPRNTCDFR"
               IdName          =   "UOPRNTCDFR"
               CmpStr          =   "LocPrntCd >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   84
               Left            =   2400
               TabIndex        =   166
               ToolTipText     =   "Show Picture? ( Enter 'Y'/'N') ('Y' will slow down the report)"
               Top             =   6840
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
               Index           =   90
               Left            =   6960
               TabIndex        =   171
               ToolTipText     =   "Show Production or Process Location"
               Top             =   840
               Visible         =   0   'False
               Width           =   315
               _ExtentX        =   556
               _ExtentY        =   503
               Enabled         =   0   'False
               DataField       =   "UOLOCTYPFR"
               IdName          =   "UOLOCTYPFR"
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Note: Selecting the option as ‘Y’es , may slow down the processing of report."
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
               Height          =   315
               Index           =   20
               Left            =   1920
               TabIndex        =   167
               Top             =   6840
               Width           =   7455
            End
            Begin VB.Label LblPrntCd 
               BackStyle       =   0  'Transparent
               Caption         =   "Parent Loc"
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
               TabIndex        =   163
               Top             =   580
               Width           =   1425
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Show Prd Pnts/ Rs. Val"
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
               Index           =   16
               Left            =   90
               TabIndex        =   161
               Top             =   6300
               Visible         =   0   'False
               Width           =   1545
            End
            Begin VB.Label LblBagTyp 
               BackStyle       =   0  'Transparent
               Caption         =   "Enter Bag Typ"
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
               TabIndex        =   160
               ToolTipText     =   "Location"
               Top             =   6030
               Width           =   1395
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
               Left            =   3330
               TabIndex        =   131
               Top             =   6030
               Width           =   885
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
               TabIndex        =   130
               Top             =   280
               Width           =   1335
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Setting Typ"
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
               TabIndex        =   129
               Top             =   5745
               Width           =   1425
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Rm SubCtg"
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
               TabIndex        =   128
               Top             =   5460
               Width           =   1425
            End
            Begin VB.Label ALBL 
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
               Index           =   8
               Left            =   90
               TabIndex        =   127
               ToolTipText     =   "Location"
               Top             =   5175
               Width           =   1425
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
               Index           =   11
               Left            =   90
               TabIndex        =   126
               Top             =   3465
               Width           =   1605
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
               Index           =   12
               Left            =   90
               TabIndex        =   125
               Top             =   3750
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
               Index           =   28
               Left            =   5355
               TabIndex        =   124
               Top             =   3465
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
               Index           =   29
               Left            =   2655
               TabIndex        =   123
               Top             =   3465
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
               Index           =   32
               Left            =   4935
               TabIndex        =   122
               Top             =   3465
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
               Index           =   33
               Left            =   2235
               TabIndex        =   121
               Top             =   3465
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
               Left            =   5055
               TabIndex        =   120
               Top             =   3750
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
               Index           =   35
               Left            =   2355
               TabIndex        =   119
               Top             =   3750
               Width           =   105
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
               TabIndex        =   118
               Top             =   2895
               Width           =   1425
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Prd Ctg"
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
               TabIndex        =   117
               Top             =   2325
               Width           =   1425
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "By Worker"
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
               TabIndex        =   116
               Top             =   1755
               Width           =   1425
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Sub Process"
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
               TabIndex        =   115
               Top             =   1470
               Width           =   1425
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Ord Tc/Yy/Chr"
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
               TabIndex        =   114
               Top             =   4035
               Width           =   1725
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Ord No"
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
               TabIndex        =   113
               Top             =   4320
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
               Left            =   5355
               TabIndex        =   112
               Top             =   4035
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
               Left            =   2625
               TabIndex        =   111
               Top             =   4035
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
               Left            =   4935
               TabIndex        =   110
               Top             =   4035
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
               Left            =   2205
               TabIndex        =   109
               Top             =   4035
               Width           =   105
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Vch Date"
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
               TabIndex        =   108
               Top             =   3180
               Width           =   1425
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
               Index           =   18
               Left            =   90
               TabIndex        =   107
               Top             =   4890
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
               TabIndex        =   106
               Top             =   4605
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
               TabIndex        =   105
               Top             =   2610
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
               TabIndex        =   104
               Top             =   2040
               Width           =   1425
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Location"
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
               TabIndex        =   103
               Top             =   1185
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
               Left            =   6900
               TabIndex        =   102
               Top             =   50
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
               Left            =   4215
               TabIndex        =   101
               Top             =   50
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
               TabIndex        =   100
               Top             =   50
               Width           =   1335
            End
         End
      End
   End
End
Attribute VB_Name = "EmrFrpPrd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rep As New EmrRepPrd
Dim moCn As MwfLib.MDOConnection   '*** (Bef speed) bef 24/11/06

' ########################  Manoj #### Ver: 2.11.0 #### Date: 26/05/2005  ########################
Dim moGrph As EmrClsGrph
' ###########################################  Manoj  ###########################################

'****** Sachin 2.14.0 - [25-11-2006] - [Multi-Company] ******
Dim ws_MultiCoMod As Boolean
'****** Sachin 2.14.0 - [25-11-2006] - [Multi-Company] ******

Private Sub SetGroupSort()
  ' ***** Manali 3.10.0 - 12/03/12 - Multi Copmpany Allowed in Scope
  GRP_REP.Add "Location", "TdByLoc", "", "TdFrBLoc", "hTdFrBLoc", "Company Code,Parent Code", "", "", "IsNull((Select LocDesc From Loc Where LocCoCd = TdCoCd and LocTyp= 'P' and LocCd= TdByLoc), '') "
  GRP_REP.Add "Process", "LocPrcs", "", "TdFrBLoc", "hTdFrBLoc", "Company Code,Parent Code", "", "", "IsNull((Select LocDesc From Loc Where LocCoCd = TdCoCd and LocTyp= 'R' and LocCd= Lc.LocPrcs), '') "
  'Process for Sub Process Added
  GRP_REP.Add "(Location)Sub Process", "TdByLoc+TpSubPrc", "'('+ RTrim(TdByLoc)+ ') '+ TpSubPrc", "TdByLoc,TpSubPrc", "hTdFrBLoc,hTpSubPrc", "Parent Code,Process", "", "", "IsNull((Select PDesc From Param Where PTyp= 'SUBPRC' and PMCd= TdByLoc and PSCd= TpSubPrc), '') "
  GRP_REP.Add "(Process)Sub Process", "LocPrcs+TpSubPrc", "'('+ RTrim(LocPrcs)+ ') '+ TpSubPrc", "TdByLoc,TpSubPrc", "hTdFrBLoc,hTpSubPrc", "Parent Code,Process", "", "", "IsNull((Select PDesc From Param Where PTyp= 'SUBPRC' and PMCd= Lc.LocPrcs and PSCd= TpSubPrc), '') "
  
  GRP_REP.Add "(Location)Worker", "TdByLoc+TpByWrk", "'('+ RTrim(TdByLoc)+ ') '+ TpByWrk", "TdByLoc,TpByWrk", "hTdFrBLoc,hTpByWrk", "Parent Code,Process", "", "", Array("IsNull((Select '(TgtPts:' + LTrim(Str(vPNum, 7, 2)) + ') '+ vPDesc From vParam Where vPCoCd = TpCoCd and vPTyp= 'WORK' and vPMCd= TpByWrk), '') ", "IsNull((Select '(Rs Val:' + LTrim(Str(vPNum* vPNum1, 7, 2)) + ') '+ vPDesc From vParam Where vPCoCd = TpCoCd and vPTyp= 'WORK' and vPMCd= TpByWrk), '') ")
  GRP_REP.Add "(Process)Worker", "LocPrcs+TpByWrk", "'('+ RTrim(LocPrcs)+ ') '+ TpByWrk", "TdByLoc,TpByWrk", "hTdFrBLoc,hTpByWrk", "Parent Code,Process", "", "", Array("IsNull((Select '(TgtPts:' + LTrim(Str(vPNum, 7, 2)) + ') '+ vPDesc From vParam Where vPCoCd = TpCoCd and vPTyp= 'WORK' and vPMCd= TpByWrk), '') ", "IsNull((Select '(Rs Val:' + LTrim(Str(vPNum* vPNum1, 7, 2)) + ') '+ vPDesc From vParam Where vPCoCd = TpCoCd and vPTyp= 'WORK' and vPMCd= TpByWrk), '') ")
  ' ***** Manali 3.10.0 - 12/03/12 - Multi Copmpany Allowed in Scope
  
  ' Bef 3.10.0
  'GRP_REP.Add "Process", "TdFrBLoc", "", "TdFrBLoc", "hTdFrBLoc", "", "", "", "IsNull((Select LocDesc From Loc Where LocCoCd = TdCoCd and LocTyp= 'P' and LocCd= TdFrBLoc), '') "
  'GRP_REP.Add "Sub Process", "TdFrBLoc+TpSubPrc", "'('+ RTrim(TdFrBLoc)+ ') '+ TpSubPrc", "TdFrBLoc,TpSubPrc", "hTdFrBLoc,hTpSubPrc", "Process", "", "", "IsNull((Select PDesc From Param Where PTyp= 'SUBPRC' and PMCd= TdFrBLoc and PSCd= TpSubPrc), '') "
  
  '*** (Bef 2.14) GRP_REP.Add "(Process)Worker", "TdFrBLoc+TpByWrk", "'('+ RTrim(TdFrBLoc)+ ') '+ TpByWrk", "TdFrBLoc,TpByWrk", "hTdFrBLoc,hTpByWrk", "Process", "", "", "IsNull((Select " + wTpPts + "+ vPDesc From vParam Where vPCoCd = TpCoCd and vPTyp= 'WORK' and vPMCd= TpByWrk), '') "
  
  '*** Jay 2.14 ***
  ' bef 3.10.0
  '''GRP_REP.Add "(Process)Worker", "TdFrBLoc+TpByWrk", "'('+ RTrim(TdFrBLoc)+ ') '+ TpByWrk", "TdFrBLoc,TpByWrk", "hTdFrBLoc,hTpByWrk", "Process", "", "", Array("IsNull((Select '(TgtPts:' + LTrim(Str(vPNum, 7, 2)) + ') '+ vPDesc From vParam Where vPCoCd = TpCoCd and vPTyp= 'WORK' and vPMCd= TpByWrk), '') ", "IsNull((Select '(Rs Val:' + LTrim(Str(vPNum* vPNum1, 7, 2)) + ') '+ vPDesc From vParam Where vPCoCd = TpCoCd and vPTyp= 'WORK' and vPMCd= TpByWrk), '') ")
  '*** Jay 2.14 ***
  
  GRP_REP.Add "Design Ctg", "DmCtg", "", "DmCtg", "hDmCtg", "", "", "", "IsNull((Select PDesc from Param where PTyp= 'DMCTG' and PMCd= DmCtg), '') "
  GRP_REP.Add "Prd Ctg", "DmPrdCtg", "", "DmPrdCtg", "hDmPrdCtg", "", "", "", "IsNull((Select PDesc from Param where PTyp= 'PRDCTG' and PMCd= DmPrdCtg), '') "
  
  'Sachin 3.11.0
  GRP_REP.Add "Sal Ctg", "DmSalCtg", "", "DmSalCtg", "", "", "", "", "IsNull((Select PDesc from Param where PTyp= 'SALCTG' And PMCd=DmCtg And PSCd=DmSalCtg), '') "
  'indigo.164 DmTcTyp= 'DM' changed as DmTcTyp= BOdDmTcTyp in DsgMst where condition
  GRP_REP.Add "Design Code", "BOdDmCd", "", "BOdDmCd", "hBOdDmCd", "Design Ctg,Prd Ctg,(DsgCd)Colour", "", "", "IsNull((Select Substring(DmDesc,1,120) From DsgMst Where " + IIF(gs_Partition = ctCurrPrtn, " DmPrtKey='" + ctCurrPrtn + "' and ", "") + " DmTcTyp= BOdDmTcTyp and DmCd= BOdDmCd and DmSz= ''), '') "
  GRP_REP.Add "Ord Kt", "OdKt", "", "OdKt", "hOdKt", "", "", "", "IsNull((Select PDesc from Param where PTyp= 'KT' and PMCd= OdKt), '') "
  GRP_REP.Add "OrderNo", "BOdTc+BOdYy+BOdChr+str(BOdNo)", "BOdTc+'/'+BOdYy+'/'+BOdChr+'/'+LTrim(Str(BOdNo))", "wOrdNo,OmCmCd,OmPoNo", "hOrdNo,hOmCmCd,hOmPoNo", "Customer Code,Customer PO", "", "", "'('+OmCmCd+ '/'+ OmPONo+ ')' "
  '******Geeta*****************Emr206****06/04/04*******
  '**********If UserLevel>6 then don't display the Customer Name
  GRP_REP.Add "Customer Code", "OmCmCd", "", "OmCmCd", "hOmCmCd", "", "", "", "IsNull((Select (case when (Select PNum from Param where Ptyp = 'USR' and PmCd = '" + Trim(EmrFrmMDI.TXT_UCD.text) + "')  <= 5 then  CmName else CmCd end) From CustMst Where CmCtg= 'C' and CmCd= OmCmCd), '') "
  GRP_REP.Add "Customer PO", "OmCmCd+OmPONo", "'('+ OmCmCd+ ') '+ LTrim(OmPONo)", "OmCmCd,OmPoNo", "hOmCmCd,hOmPONo", "", "", "", "BOdTc+'/'+BOdYy+'/'+BOdChr+'/'+LTrim(Str(BOdNo))"
  GRP_REP.Add "Rm Ctg", "TpRmCtg", "", "TpRmCtg", "hTpRmCtg", "", "", "", "IsNull((Select PDesc From Param Where PTyp= 'RMCTG' and PMCd = TpRmCtg), '') "
  GRP_REP.Add "Set Type", "TpSetTyp", "TpSetTyp", "TpSetTyp", "hSetTyp", "", "", "", "IsNull((Select PDesc From Param Where PTyp= 'LABSCD' and PMCd= 'SET' and PSCd = TpSetTyp), '') "
  GRP_REP.Add "Rm SubCtg", "TpRmCtg+TpRmSCtg", "'('+ RTrim(TpRmCtg)+ ') '+ TpRmSCtg", "TpRmCtg,TpRmSCtg", "hTpRmCtg,hTpRmSCtg", "Rm Ctg", "", "", "IsNull((Select PDesc From Param Where PTyp= 'RMSCTG' and PMCd = TpRmCtg and PSCd= TpRmSCtg), '') "
  GRP_REP.Add "(DsgCd) Colour", "BOdDmCd+OdDmCol", "'('+BOdDmCd+') '+OdDmCol", "BOdDmCd,OdDmCol", "hBOdDmCd,hOdDmCol", "Design Ctg,Prd Ctg,Design Code", "", "", ""
  '6.1
  GRP_REP.Add "Date", "convert(char(10),TDt,112)", "convert(varchar, TDt, " + CStr(DtFmtSlahVal()) + ")", "", "", "", "", "", ""
  
  '4.1.0.0 (Sachin) - Addl. Group Sorts based on Date(TDt)
  GRP_REP.Add "Week", "LTrim(Str(DatePart(yyyy,TDt)))+' - '+DateName(WW,TDt)", "", "", "", "Quarter,Month", "", "", ""
  GRP_REP.Add "Month", "Str(DatePart(yyyy,TDt))+Str(DatePart(mm,TDt))", "LTrim(Str(DatePart(yyyy,TDt)))+' '+DateName(MM,TDt)", "", "", "Quarter", "", "", ""
  GRP_REP.Add "Quarter", "LTrim(Str(DatePart(yyyy,TDt)))+' - '+DateName(QQ,TDt)", "", "", "", "", "", "", ""
  GRP_REP.Add "Yearly", "Str(DatePart(yyyy,TDt))", "LTrim(Str(DatePart(yyyy,TDt)))", "", "", "", "", "", ""
  
  '### GRP_REP.Add "Customer PO", "OmCmCd+OmPONo", "'('+ OmCmCd+ ') '+ LTrim(OmPONo)", "wOrdNo,OmCmCd,OmPoNo", "hOrdNo,hOmCmCd,hOmPONo", "Customer Code,OrderNo", "", "", "'('+ OmCmCd+ ')'+ BOdTc+'/'+BOdYy+'/'+BOdChr+'/'+LTrim(Str(BOdNo))"
  ' ***** Manali 3.10.0 - 12/03/12 - Multi Copmpany Allowed in Scope
  If ws_MultiCoMod = True Then
    GRP_REP.Add "Company Code", "TdCoCd", "TdCoCd", "", "", "", "", "", "(Select HName from Head where HCd= '" + ctSelfCmCd + "' and HCoCd=TdCoCd) "
    GRP_REP.Add "Parent Location", "LocPrntCd", "LocPrntCd", "", "", "", "", "", "(Select Ln.LocDesc from Loc Ln where Ln.LocCd=Lc.LocPrntCd and Ln.LocCoCd='" + ctSelfCoCd + "') "
  End If
  ' ***** Manali 3.10.0 - 12/03/12 - Multi Copmpany Allowed in Scope
  
' ########################  Manoj #### Ver: 2.11.0 #### Date: 26/05/2005  ########################
  moGrph.AddItem "Prd Quantity", "TpRmQty", "{rdo.TpRmQty}", ""
  moGrph.AddItem "Prd Points", "TpPts", "{rdo.TpPts}", ""
' ###########################################  Manoj  ###########################################

'*** Jay 2.13(CT) ***
    GRP_REP.Add "Bag Type", "(Select (case PValue when 'F' then 'Flute Bag' when 'P' then 'Primary Bag' when 'C' then 'Component Bag' end) from Param where PTyp= 'BCHR' and PMCd= BChr and PSCd= '')", _
      "(Select case PValue when 'F' then 'Flute Bag' when 'P' then 'Primary Bag' when 'C' then 'Component Bag' end from Param where PTyp= 'BCHR' and PMCd= BChr and PSCd= '')", "", "", "Bag Typ+Chr", "", "", "(Select case PValue when 'F' then 'Flute Bag' when 'P' then 'Primary Bag' when 'C' then 'Component Bag' end from Param where PTyp= 'BCHR' and PMCd= BChr and PSCd= '')"
    GRP_REP.Add "Bag Character", "BChr", "BChr", "", "", "Bag Type,Bag Typ+Chr", "", "", "(Select PDesc from Param where PTyp ='BCHR' and PMCd= BChr and PSCd= '')"
    GRP_REP.Add "Bag Typ+Chr", "(Select (case PValue when 'F' then 'Flute Bag' when 'P' then 'Primary Bag' when 'C' then 'Component Bag' end) from Param where PTyp= 'BCHR' and PMCd= BChr and PSCd= '')+BChr", _
      "(Select (case PValue when 'F' then 'Flute Bag' when 'P' then 'Primary Bag' when 'C' then 'Component Bag' end) from Param where PTyp= 'BCHR' and PMCd= BChr and PSCd= '')+'/ '+ BChr", "", "", "Bag Type,Bag Character", "", "", "(Select PDesc from Param where PTyp ='BCHR' and PMCd= BChr and PSCd= '')"
'' **** Zubin 211 **** '
'  GRP_REP.Add "Bag Type", "PValue", "(case PValue when 'Y' then 'Flute Bag' when 'N' then 'Primary Bag' end )", "", "", "Bag Typ+Chr", "", "", "(case PValue when 'Y' then 'Flute Bag' when 'N' then 'Primary Bag' end)"
'  GRP_REP.Add "Bag Character", "BChr", "BChr", "", "", "Bag Type,Bag Typ+Chr", "", "", "(Select PDesc from param where PTyp ='BCHR' and PMCd= BChr and PSCd= '')"
'  GRP_REP.Add "Bag Typ+Chr", "PValue+BChr", "(case PValue when 'Y' then 'Flute Bag' when 'N' then 'Primary Bag' end)+ '/' + BChr", "", "", "Bag Type,Bag Character", "", "", "(Select PDesc from param where PTyp ='BCHR' and PMCd= BChr and PSCd= '')"
'' **** Zubin 211 **** '
'*** Jay 2.13(CT) ***
End Sub
Private Sub adc_setreprecsource()
  '*** Report Sql ***
'  Select space(30) as Grp1, space(30) as Grp2, space(30) as Grp3,
'  space(30) as DGrp1, space(30) as DGrp2, space(30) as DGrp3,
'  space(30) as G1Desc, space(30) as G2Desc, space(30) as G3Desc,
'  TdFrBLoc, TpSubPrc, TpByWrk, TpSetTyp, TpBQty, TpRmCtg, TpRmSctg, TpRmQty, TpPts, TpPts * vPNum1 as qTpVal, TpTc, TpYy, TpChr, TpNo,
'  TpSr, TpSrNo, TDt, TdBYy, TdBChr, TdBNo, BOdTc, BOdYy, BOdChr, BOdNo, BOdDmCd, DmCtg ,
'  DmPrdCtg , OdKt, OdDmCol, OmCmCd, OmPoNo
'  From TxndPrd, Txnd, Txn, Bag, OrdDsg, OrdMst, DsgMst where 1=2
  '*** Report Sql ***
  'Debug.Print "adc_setreprecsource" + adc("UOYN3").Value
  Dim grpflds As String, wrepcnd As String, wCnd As String, wSqlStrg As String, wDetPos As Integer, i As Integer
      
'*** Jay 2.13(CT) ***
' **** Zubin 211 **** '
  Dim ws_BagJoin As String
'  Dim ws_FluteBagJoin As String
' **** Zubin 211 **** '
'*** Jay 2.13(CT) ***

  '*** Jay 2.14 ***
  Dim ws_TpPts As String
  Dim ws_vParamJoin As String
  '*** Jay 2.14 ***
  
  Dim wAddlCnd As String   ' ****** Sachin 3.02 - New Fields in Where Clause
    
  Set Rep = Nothing
  Set ADC.RepSource = Rep
  Call DispCoNm
  Call SetMwName(Rep)   ' ***** Manali 3.6.1 - Microway footer
  Call GRP_REP.Gen3LRep
  
  'grpflds = GRP_REP.GrpFldLst    '*** (Bef 2.14)
    
  '*** Jay 2.14 ***
  If ADC("UoYN3") = "Y" Then
    grpflds = GRP_REP.GrpFldLst(0)
  Else
    grpflds = GRP_REP.GrpFldLst(1)
  End If
  '*** Jay 2.14 ***
  
  '***********Bhavna added pic y/n
  Dim wGrp(3) As String
  wGrp(1) = GRP_REP.Value(gltGroup1)
  wGrp(2) = GRP_REP.Value(gltGroup2)
  wGrp(3) = GRP_REP.Value(gltGroup3)
  
  wDetPos = 0
  For i = 1 To 3
    If wGrp(i) = "Design Code" Then
        wDetPos = i: Exit For
    End If
  Next i
  '***********Bhavna added pic y/n
    
  wrepcnd = ADC.RepCond
  wCnd = IIF(wrepcnd <> "", " where " + wrepcnd, "")
  If OptLoc(0).Value = True Then
      If ADC("UoIssLocFr") <> "" Then wCnd = wCnd + " and TdByLoc >= '" + ADC("UoIssLocFr") + "'"
      If ADC("UoIssLocTo") <> "" Then wCnd = wCnd + " and TdByLoc<= '" + ADC("UoIssLocTo") + "'"
      If ADC("UoIssLocSel") <> "" Then wCnd = wCnd + " and TdByLoc In (" + ADC("UoIssLocSel") + ")"
    ElseIf OptLoc(1).Value = True Then
      If ADC("UoIssLocFr") <> "" Then wCnd = wCnd + " and LocPrcs >= '" + ADC("UoIssLocFr") + "'"
      If ADC("UoIssLocTo") <> "" Then wCnd = wCnd + " and LocPrcs <= '" + ADC("UoIssLocTo") + "'"
      If ADC("UoIssLocSel") <> "" Then wCnd = wCnd + " and LocPrcs In (" + ADC("UoIssLocSel") + ")"
    End If
   
'*** Jay 2.13(CT) ***
  ws_BagJoin = "Join Param On PTyp= 'BCHR' and PMCd= BChr and PSCd= '' " + _
                    IIF(ADC("UoBagTyp") <> "", " and PValue= '" + ADC("UoBagTyp") + "' ", " ") + _
                    IIF(ADC("UOBCHRFR") <> "", " and PMCd= '" + ADC("UOBCHRFR") + "' ", " ")
'  ' **** Zubin 211 **** '
'  If UCase(adc("UOYN2")) = "N" Then
'    ws_FluteBagJoin = " Join Param On PTyp= 'BCHR' and PMCd= BChr and PSCd= '' and PValue = 'N' " + _
'              IIF(adc("UOBCHRFR") <> "", " and PMCd= '" + adc("UOBCHRFR") + "' ", " ")
'  ElseIf UCase(adc("UOYN2")) = "Y" Then
'    ws_FluteBagJoin = " Join Param On PTyp= 'BCHR' and PMCd= BChr and PSCd= '' and PValue = 'Y' " + _
'              IIF(adc("UOBCHRFR") <> "", " and PMCd= '" + adc("UOBCHRFR") + "' ", " ")
'  Else
'    ws_FluteBagJoin = " Join Param On PTyp= 'BCHR' and PMCd= BChr and PSCd= '' " + _
'              IIF(adc("UOBCHRFR") <> "", " and PMCd= '" + adc("UOBCHRFR") + "' ", " ")
'  End If
'  ' **** Zubin 211 **** '
'*** Jay 2.13(CT) ***

  '****** Sachin 3.02.0 - Additional Conditions For New Fields
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(ADC("UOCMCDFR")) <> "", " And ", "") + IIF(Trim(ADC("UOCMCDFR")) <> "", " OdOmCmCd >= '" + Trim(ADC("UOCMCDFR")) + "'", "")
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(ADC("UOCMCDTO")) <> "", " And ", "") + IIF(Trim(ADC("UOCMCDTO")) <> "", " OdOmCmCd <= '" + Trim(ADC("UOCMCDTO")) + "'", "")
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(ADC("UOCMCDSEL")) <> "", " And ", "") + IIF(Trim(ADC("UOCMCDSEL")) <> "", " OdOmCmCd In (" + Trim(ADC("UOCMCDSEL")) + ")", "")
  
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(ADC("UODTFR")) <> "01/01/80", " And ", "") + IIF(Trim(ADC("UODTFR")) <> "01/01/80", " TpDt >= '" + Trim(ADC("UODTFR")) + "'", "")
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(ADC("UODTTO")) <> "01/01/80", " And ", "") + IIF(Trim(ADC("UODTTO")) <> "01/01/80", " TpDt <= '" + Trim(ADC("UODTTO")) + "'", "")
  '****** Sachin 3.02.0 - Additional Conditions For New Fields

  ' **** Manali 3.8.0 - Both Shown in Report
  '*** Jay 2.14 ***
'  If UCase(adc("UOYN3")) = "Y" Then
'    'prd points
'    ws_TpPts = " TpPts "
'    ws_vParamJoin = " "
'    Rep.h1txtTpPts.SetText "Points"
'    Rep.h2txtTpPts.SetText "Points"
'  ElseIf UCase(adc("UOYN3")) = "N" Then
'    'Rs value
'    ws_TpPts = " vPNum1 * TpPts "
'    ws_vParamJoin = " Join vParam on vPTyp= 'WORK' and vPMCd= TpByWrk  and vPSCd= '' and TpCoCd= vPCoCd "
'    Rep.h1txtTpPts.SetText "Rs. Val"
'    Rep.h2txtTpPts.SetText "Rs. Val"
'  End If
  '*** Jay 2.14 ***
  ' **** Manali 3.8.0 - Both Shown in Report
  ws_vParamJoin = " Join vParam on vPTyp= 'WORK' and vPMCd= TpByWrk  and vPSCd= '' and TpCoCd= vPCoCd "
  
  Rep.TxtFrVchDt.SetText Format(ADC("UODTFR"), DtShortStrYyyy())    '6.1
  Rep.TxtToVchDt.SetText Format(ADC("UODTTO"), DtShortStrYyyy())
  
  '*** Jay 2.14 *** (Rs fector for worker is added)
  ' **** Zubin 211 (ws_FluteBagJoin added) **** '
  '****** Sachin 3.02 - Id fields in Join
  ' ***** Manali 3.8.0 - Value added
  ' ***** Manali 3.8.0 - Order By added in conctenated String (otherwise speed issue in case of Clustered indexes)
  ' ***** Manali 3.10.0 - 12/03/12 - Multi Copmpany Allowed in Scope  - Loc Join added
  ' ***** Manali 3.10.0 - Order By removed, Change reverted
  '**********Bhavna added dsg sz,sfx and picnm
  'indigo.164 DmTcTyp= 'DM' changed as DmTcTyp= BOdDmTcTyp in DsgMst join
  wSqlStrg = "Select " + grpflds + ", " + _
             "TdCoCd," + IIF(OptLoc(0).Value = True, "TdByLoc", "LocPrcs") + " as TdFrBLoc, TpSubPrc, TpByWrk, TpSetTyp, TpBQty, TpRmCtg, TpRmSCtg, TpRmQty, /*" + ws_TpPts + "*/ TpPts as TpPts, TpPts * vPNum1 as qTpVal, " + _
             "TpTc, TpYy, TpChr, TpNo, TpSr, TpSrNo, TDt, TdBYy, TdBChr, TdBNo, (case when TdBQty<>0 then (TdBGrWt/TdBQty)*TpBQty else 0 end) as TdBGrWt,  " + _
             "BOdTc, BOdYy, BOdChr, BOdNo, BOdDmCd, DmCtg, DmPrdCtg, OdKt, OdDmCol, " + _
             "OdDmSz, OdSfx, OdPicNm, DmTcTyp, " + _
             "OmCmCd, OmPoNo, LocPrcs From TxndPrd " + _
             "Join Txnd on TdIdNo=TpTdIdNo and TdCoCd= TpCoCd and TdTc= TpTc and TdYy= TpYy and TdChr= TpChr " + _
             "and TdNo= TpNo and TpSr= TdSr and TdSrNo= 0 " + IIF(gs_Partition = ctCurrPrtn, " and TdPrtKey=TpPrtKey ", "") + _
             "Join Txn on TCoCd= TpCoCd and TTc= TpTc and TYy= TpYy and TChr= TpChr and TNo= TpNo " + IIF(gs_Partition = ctCurrPrtn, " and TPrtKey=TpPrtKey ", "") + _
             "Join Bag on BIdNo=TdBIdNo and BCoCd= TdCoCd and BYy= TdBYy and BChr= TdBChr and BNo= TdBNo " + IIF(gs_Partition = ctCurrPrtn, " and BPrtKey=TdPrtKey ", "") + ws_BagJoin + _
             "Join Loc Lc on LocCocd= TdCoCd and LocCd=TdByLoc " + _
             "Join OrdDsg on OdIdNo=BOdIdNo and OdCoCd= BCoCd and OdTc= BOdTc and OdYy= BOdYy " + _
             "and OdChr= BOdChr and OdNo= BOdNo and OdSr= BOdSr " + IIF(gs_Partition = ctCurrPrtn, " and OdPrtKey=BPrtKey ", "") + _
             "Join OrdMst on OmIdNo=BOmIdNo and OmCoCd= BCoCd and OmTc= BOdTc and OmYy= BOdYy " + _
             "and OmChr= BOdChr and OmNo= BOdNo " + IIF(gs_Partition = ctCurrPrtn, " and OmPrtKey=BPrtKey ", "") + _
             "Join DsgMst on DmIdNo=BDmIdNo and DmTcTyp= BOdDmTcTyp and DmCd= BOdDmCd and DmSz= '' " + IIF(gs_Partition = ctCurrPrtn, " and BPrtKey=DmPrtKey ", "") + ws_vParamJoin + _
             wCnd + IIF(gs_Partition = ctCurrPrtn, " and TpPrtKey='" + ctCurrPrtn + "' ", "") + IIF(wAddlCnd <> "", IIF(wCnd <> "", " And ", " Where ") + wAddlCnd, "")
             '''" Order By Grp1, Grp2, Grp3, TdCoCd+TdBYy+TdBChr+Str(TdBNo, 10) "
             
  ADC.RepRecSource = wSqlStrg + ctMaxDopOpt  '****** Sachin 3.01 added ctMaxDopOpt

'***********Bhavna added show pic y/n
  If ADC("UoYN1") = "N" Or wDetPos = 0 Then
    Rep.PICSEC.Suppress = True
  End If
'***********Bhavna added show pic y/n

'*** Jay 2.13(CT) ***
  If ADC("UoBagTyp") = "F" Then
    Rep.TxtBagTypAndChr.SetText "(Flute Bags" & IIF(ADC("UoBChrFr") <> "", ",  Chr = " & ADC("UoBChrFr"), "") & ")"
  ElseIf ADC("UoBagTyp") = "P" Then
    Rep.TxtBagTypAndChr.SetText "(Primary Bags" & IIF(ADC("UoBChrFr") <> "", ",  Chr = " & ADC("UoBChrFr"), "") & ")"
  ElseIf ADC("UoBagTyp") = "C" Then
    Rep.TxtBagTypAndChr.SetText "(Component Bags" & IIF(ADC("UoBChrFr") <> "", ",  Chr = " & ADC("UoBChrFr"), "") & ")"
  End If
  
  GRP_REP.SetFormula Rep, "wWrkrChngAddToBQty", IIF(ChkBoxWrkr.Value = Checked, "'Y'", "'N'")

  
  
'  '***(Jen 2.11)
'  If adc("UoYn2") <> "" Then Rep.TxtBagTypAndChr.SetText IIF(adc("UoYn2") = "Y", "(Flute Bags", "(Primary Bags") & IIF(adc("UoBChrFr") <> "", ",  Chr = " & adc("UoBChrFr"), "") & ")"
'  '***(Jen 2.11)
'*** Jay 2.13(CT) ***
  
  CRV_REP.DisplayGroupTree = False

' ########################  Manoj #### Ver: 2.11.0 #### Date: 26/05/2005  ########################
  Dim wsQryStruct As String
  wsQryStruct = "  Select space(30) as Grp1, space(30) as Grp2, space(30) as Grp3, " + _
    "space(30) as DGrp1, space(30) as DGrp2, space(30) as DGrp3, " + _
    "space(50) as G1Desc, space(50) as G2Desc, space(50) as G3Desc, " + _
    "TdFrBLoc, TpSubPrc, TpByWrk, TpSetTyp, TpBQty, TpRmCtg, TpRmSctg, TpRmQty, TpPts, TpTc, TpYy, TpChr, TpNo, " + _
    "TpSr, TpSrNo, TDt, TdBYy, TdBChr, TdBNo, BOdTc, BOdYy, BOdChr, BOdNo, BOdDmCd, DmCtg , " + _
    "DmPrdCtg , OdKt, OdDmCol, OmCmCd, OmPoNo " + _
    "From TxndPrd, Txnd, Txn, Bag, OrdDsg, OrdMst, DsgMst where 1=2"
  Call moGrph.SetGraphProp(Rep, wsQryStruct, wSqlStrg + ctMaxDopOpt)  '****** Sachin 3.01 added ctMaxDopOpt
' ###########################################  Manoj  ###########################################
  
  Call DispReport(ADC) '****** Sachin 3.02 - Display Report ******
  
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
  gb_CoCdFor = False      ' ***** Manali 3.10.0 - 12/03/12 - Multi Copmpany Allowed in Scope
  '''gb_CoCdFor = True
  '****** Sachin 2.14.0 - [23-11-2006] - [Multi-Company] ******

  '****** Sachin 2.14.0 - [23-11-2006] - [Multi-Company] ******
  If ws_MultiCoMod = True Then
    Set ADC.FirNKeyCtl = ADC("UoCoCdFr")
  Else
    Set ADC.FirNKeyCtl = ADC("UoIssLocFr")
  End If
  '****** Sachin 2.14.0 - [23-11-2006] - [Multi-Company] ******
  
  Call InitProp(Me)
  
  '*** Parameters Used For Setting Properties Of UsrOpt Fields ***
  'gs_LocTyp = "'P'"
  'gb_RmCtgFor = True
  'gb_RmZ = False
  'indigo.164 sketch design tctype and order tc added
  gs_DmTcTyp = "DM','SM"
  gs_OmTcTyp = "SO','OS"
  
  gs_TxnTcTyp = "DT"
  gs_CmCtg = "C"
  gs_LabMCd = "SET"
  'gs_RmGrp = "ORDDET"
  'gs_LabGrp = "ORDDET"
  '*** Parameters Used For Setting Properties Of UsrOpt Fields ***

' ########################  Manoj #### Ver: 2.11.0 #### Date: 26/05/2005  ########################
  Set moGrph = New EmrClsGrph
  Call moGrph.SetFormDets(Me, moCn)  '*** (Bef speed) bef 24/11/06
  'Call moGrph.SetFormDets(Me)  '*** (Jen speed) 24/11/06
' ###########################################  Manoj  ###########################################
  
  
  '****** Sachin 2.14.0 - [Multi-Company] ******
  ADC("UOCOCDFR").Enabled = ws_MultiCoMod
  '****** Sachin 2.14.0 - [Multi-Company] ******
  
  ' ***** Manali 3.10.0 - 12/03/12 - Multi Copmpany Allowed in Scope
  LblPrntCd.Visible = ws_MultiCoMod
  ADC("UOCOCDFR").Enabled = ws_MultiCoMod: ADC("UOCOCDTO").Visible = ws_MultiCoMod
  ADC("UOCOCDSEL").Enabled = ws_MultiCoMod: ADC("UOCOCDSEL").Visible = ws_MultiCoMod
  ADC("UoPrntCdFr").Enabled = ws_MultiCoMod: ADC("UoPrntCdTo").Enabled = ws_MultiCoMod: ADC("UoPrntCdSel").Enabled = ws_MultiCoMod
  ADC("UoPrntCdFr").Visible = ws_MultiCoMod: ADC("UoPrntCdTo").Visible = ws_MultiCoMod: ADC("UoPrntCdSel").Visible = ws_MultiCoMod
  ' ***** Manali 3.10.0 - 12/03/12 - Multi Copmpany Allowed in Scope
  
  Call SetGroupSort
End Sub

Private Sub ATXT_FldChange(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String, ByVal pv_NewValue As Variant)
  
' ***** Manali 3.10.0 - 12/03/12 - Multi Copmpany Allowed in Scope
''  '****** Sachin 2.14.0 - [24-11-2006] - [Multi-Company] ******
''  gb_CoCdFor = True
''
''  Select Case UCase(IdName)
''    Case Is = UCase("UoCoCdFr")
''    If Not moCn.RecSeek("Select 'x' From Head Where HCd='" + ctSelfCmCd + "' And HCoCd='" + pv_NewValue + "'") Then _
''          Cancel = True: ErrMsg = "Invalid Company Code": Exit Sub
''     Me.Caption = GetFrmCaption(pv_NewValue, "Production Report")
''  End Select
''  '****** Sachin 2.14.0 - [24-11-2006] - [Multi-Company] ******
  Dim wValidLocSeq As String, wLocSel() As String, i As Integer
  '6.1B-16 checking valid location to the user
  Select Case UCase(IdName)
  Case Is = UCase("UOISSLOCFR"), UCase("UOISSLOCTO")
    If pv_NewValue <> "" Then
      If ADC("UOLOCTYPFR") = "P" Then
        wValidLocSeq = moCn.GetFldVal("Select PDesc225 from param where PTyp= 'USR' and PMCd= '" + Trim(ADC.UsrCd) + "'")
        If wValidLocSeq = "" Then Exit Sub
      Else
        wValidLocSeq = moCn.GetFldVal("Select PValue1 from param where PTyp= 'USR' and PMCd= '" + Trim(ADC.UsrCd) + "'")
        If wValidLocSeq = "" Then Exit Sub
      End If
      If InStr(1, "," + wValidLocSeq + ",", "," + pv_NewValue + ",") = 0 Then _
        Cancel = True: ErrMsg = "User Does Not Have Access to This Location ": Exit Sub
    End If
  Case Is = UCase("UOISSLOCSEL")
    If pv_NewValue <> "" Then
      If ADC("UOLOCTYPFR") = "P" Then
        wValidLocSeq = moCn.GetFldVal("Select PDesc225 from param where PTyp= 'USR' and PMCd= '" + Trim(ADC.UsrCd) + "'")
        If wValidLocSeq = "" Then Exit Sub
      Else
        wValidLocSeq = moCn.GetFldVal("Select PValue1 from param where PTyp= 'USR' and PMCd= '" + Trim(ADC.UsrCd) + "'")
        If wValidLocSeq = "" Then Exit Sub
      End If
      pv_NewValue = Replace(pv_NewValue, "'", "")
      wLocSel = Split(pv_NewValue, ",")
      If IsArray(wLocSel) Then
        For i = 0 To UBound(wLocSel)
          If InStr(1, "," + wValidLocSeq + ",", "," + wLocSel(i) + ",") = 0 Then _
            Cancel = True: ErrMsg = "User Does Not Have Access to " + wLocSel(i) + " Location ": Exit Sub
        Next i
      End If
    End If
  '6.1B-16 checking valid character access to the user
  Case Is = UCase("UOTCHRFR"), UCase("UOTCHRTO"), UCase("UOOMCHRFR"), UCase("UOOMCHRTO")
    Dim wValidUsr As String, wPMCd As String, w_CoCd As String
    If UCase(IdName) = UCase("UOTCHRFR") Then
      wPMCd = ADC("UOTTCFR"):   w_CoCd = ADC("UOCOCDFR")
    ElseIf UCase(IdName) = UCase("UOTCHRTO") Then
      wPMCd = ADC("UOTTCTO"):   w_CoCd = ADC("UOCOCDFR")
    ElseIf UCase(IdName) = UCase("UOOMCHRFR") Then
      wPMCd = ADC("UOOMTCFR"):   w_CoCd = ADC("UOCOCDFR")
    ElseIf UCase(IdName) = UCase("UOOMCHRTO") Then
      wPMCd = ADC("UOOMTCTO"):   w_CoCd = IIF(gb_CoCdFor = False, ADC("UOCOCDTO"), ADC("UOCOCDFR"))
    End If
    If pv_NewValue <> "" Then
      wValidUsr = moCn.GetFldVal("Select vPValue4 From vParam Where vPCoCd ='" + w_CoCd + _
                    "' and vPTyp ='CHR' and vPMCd ='" + wPMCd + "' and vPSCd ='" + pv_NewValue + "'")
      If wValidUsr = "" Then Exit Sub
      If InStr(1, "," + wValidUsr + ",", "," + Trim(ADC.UsrCd) + ",") = 0 Then _
        Cancel = True: ErrMsg = "User Does Not Have Access to This Character ": Exit Sub
    End If
  End Select
    
  Call SetProp(Me, IdName, Valid, Cancel, ErrMsg)
  
End Sub

'**********Bhavna added pic y/n checkbox
Private Sub ChkBoxArr_Click(Index As Integer)
With ChkBoxArr(Index)
  Select Case Index
  Case Is = 1     'Tag = "ShowPic"
    If .Value = Checked Then ADC("UoYN1") = "Y" Else ADC("UoYN1") = "N"
  End Select
End With
End Sub
'**********Bhavna added pic y/n checkbox


Private Sub ATXT_FldWhen(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String)
' ########################  Manoj #### Ver: 2.11.0 #### Date: 26/05/2005  ########################
  If UCase$(IdName) = UCase$("UoGrphTyp") Then gs_GrphGrpNo = moGrph.SetGrphGrpNo
' ###########################################  Manoj  ###########################################
  gs_LocTyp = "'P'"
  ' ***** Manali 3.10.0 - 12/03/12 - Multi Copmpany Allowed in Scope
  'gb_CoCdFor = True
  'indigo.164 sketch design tctype and order tc added
  gs_DmTcTyp = "DM','SM"
  gs_OmTcTyp = "SO','OS"
  '*** (Jen 2.11)
  Select Case UCase(IdName)
    Case Is = UCase("UoBChrFr")
      '*** Jay 2.13(CT) ***
      If ADC.Mode = xNorm Then
        If ADC("UoBagTyp") = "" Then Cancel = True: ErrMsg = "First Specify Bag Type": Exit Sub
        HlpList.PMCd "BCHR", "'" + ADC("UoBagTyp") + "'": Exit Sub
      End If
      '*** Jay 2.13(CT) ***
    Case Is = UCase("UOISSLOCFR"), UCase("UOISSLOCTO"), UCase("UOISSLOCSEL")
      If ADC("UOLOCTYPFR") = "R" Then gs_LocTyp = "'R'"
      If ADC("UOLOCTYPFR") = "P" Then gs_LocTyp = "'P'"

  End Select
  '*** (Jen 2.11)
  
  SetProp Me, IdName, When
End Sub
Private Sub atxt_ReCalculate(Index As Integer, ByVal IdName As String, ByVal OldVal As Variant, ByVal RowNum As Integer, ByVal SourceName As String, ByVal SourceColName As String)
' **** Zubin 211 **** '
  Select Case UCase(IdName)
    Case Is = UCase("UOBCHRFR")
      'If ADC("UOYN2") = "N" Then
      ADC("UOBCHRFR") = ""
  End Select
' **** Zubin 211 **** '
End Sub

Private Sub OptLoc_GotFocus(Index As Integer)
  DispMsg FraLoc.ToolTipText, etInfo
Select Case Index
Case Is = 0
  ADC("UOISSLOCFR") = "":   ADC("UOISSLOCTO") = "":   ADC("UOISSLOCsel") = ""
  ADC("UoLocTypFr") = "P"
Case Is = 1
  ADC("UOISSLOCFR") = "":   ADC("UOISSLOCTO") = "":   ADC("UOISSLOCsel") = ""
  ADC("UoLocTypFr") = "R"
End Select

End Sub

' ****** Manali 3.5.0 - 03/11/08 - Radio Buttons Added
Private Sub OptPtsVal_Click(Index As Integer)
Select Case Index
  Case Is = 0
    ADC("UoYN3") = "Y"
  Case Is = 1
    ADC("UoYN3") = "N"
End Select
End Sub
' ****** Manali 3.5.0 - 03/11/08 - Radio Buttons Added

Private Sub tab_REP_Click(PreviousTab As Integer)
  fra_tabrep(TAB_REP.Tab).Enabled = True
  fra_tabrep(PreviousTab).Enabled = False
End Sub
Private Sub ADC_KeyWhen()
  ADC("UOUSRCD") = ADC.UsrCd
  ADC("UOMNUCD") = ADC.MenuCd
  
End Sub
Private Sub ADC_NonKeyWhen(pr_Cancel As Boolean, pr_ErrMsg As String)
  ' ***** Manali 3.10.0 - 12/03/12 - Multi Copmpany Allowed in Scope
  If ws_MultiCoMod = False Then _
      ADC("UoCoCdFr") = gs_CoCd: ADC("UoCoCdTo") = gs_CoCd: ADC("UoCoCdSel") = "": ADC("UoCoCdFr").Enabled = False

 ''adc("UoCoCdFr") = gs_CoCd
 
 '****** Sachin 2.14.0 - [24-11-2006] - [Multi-Company] ******
  Me.Caption = GetFrmCaption("", "Production Report")     ' ***** Manali 3.10.0 - 12/03/12 - Multi Copmpany Allowed in Scope
  '''Me.Caption = GetFrmCaption(adc("UOCOCDFR"), "Production Report")
  '****** Sachin 2.14.0 - [24-11-2006] - [Multi-Company] ******
 
' ########################  Manoj #### Ver: 2.11.0 #### Date: 26/05/2005  ########################
  If ADC("UoShowGrph") = "" Then ADC("UoShowGrph") = "N"
  Call moGrph.NonKeyWhen
' ###########################################  Manoj  ###########################################

'************bhavna added pic y/n
  ADC("UoYN1") = "N"
'  If adc("UoYN1") = "" Then adc("UoYN1") = "N"
  
  If ADC("UoYN1") = "Y" Then ChkBoxArr(1).Value = Checked Else ChkBoxArr(1).Value = Unchecked
'************bhavna added pic y/n
  '*** Jay 2.13(CT) ***
  Dim wHeadChk As MwfLib.MDORowSet
  Set wHeadChk = moCn.OpenRes("Select hFluteBagYN, hCompBagsYn from Head where HCoCd= '" + ctSelfCoCd + "' and " + _
                                                    "HCd= '" + ctSelfCmCd + "' ")
  If wHeadChk!hFluteBagYN = "N" And wHeadChk!hCompBagsYn = "N" Then
    LblBagTyp.Visible = False
    ADC("UoBagTyp").Visible = False
    LblUoBChrFr.Visible = False
    ADC("UoBChrFr").Visible = False
    ADC("UoBagTyp") = "P": ADC("UoBChrFr") = ""
  End If
  Set wHeadChk = Nothing
  '*** Jay 2.13(CT) ***
  
'*** Jay 2.13(CT) ***
'  '*** (Jen 2.11)
'  If moCn.GetFldVal("Select hFluteBagYN from Head where HCoCd= '" + adc("UoCoCdFr") + "' and " + _
'                                                    "HCd= '" + ctSelfCmCd + "' ") = "N" Then
'    LblPrimFluteBags.Visible = False
'    adc("UoYN2").Visible = False
'    LblUoBChrFr.Visible = False
'    adc("UoBChrFr").Visible = False
'    adc("UOYN2") = "N": adc("UoBChrFr") = ""
'  End If
'  '*** (Jen 2.11)
'*** Jay 2.13(CT) ***

' ****** Manali 3.5.0 - 03/11/08 - Radio Buttons Added
' ***** Manali 3.8.0 - Following Option Removed
'If adc("UoYN3") = "Y" Then OptPtsVal(0).Value = True
'If adc("UoYN3") = "N" Then OptPtsVal(1).Value = True
' ****** Manali 3.5.0 - 03/11/08 - Radio Buttons Added
'***********Bhavna show picture default = 'N'
  If ADC("UOLOCTYPFR") = "P" Then
    OptLoc(0).Value = True
  ElseIf ADC("UOLOCTYPFR") = "R" Then
    OptLoc(1).Value = True
  End If


End Sub
Private Sub DispCoNm()
'Rep.wCoCd.SetText gs_CoNm
  Dim ws_HName As String
    ' ***** Manali 3.10.0 - 12/03/12 - Multi Copmpany Allowed in Scope
    ws_HName = moCn.GetFldVal("Select HName from Head where HCoCd ='" + IIF(ws_MultiCoMod = False Or UCase(ADC.MenuCd) = UCase("RepWipFgBPcs"), ADC("UOCOCDFR"), ctSelfCoCd) + "' and " + _
                                                      "HCd = '" + ctSelfCmCd + "' ")
  
''  ws_HName = moCn.GetFldVal("Select HName from Head where HCoCd = '" + adc("UoCoCdFr") + "' and " + _
''                                                    "HCd = '" + ctSelfCmCd + "' ")
  Rep.wCoCd.SetText ws_HName
  
  ' ***** Manali 3.8.0 - Comp Logo in Report
  Rep.wCoCdLogo.SetText ws_HName
  If moCn.RecSeek("Select 'x' From Head Where HCoCd='" + gs_CoCd + "' and HCd='" + ctSelfCmCd + "' and HShowRepLogo='Y' ") Then
    Rep.RHMain.UnderlaySection = True: Rep.wCoCd.Suppress = True
  Else
    Rep.RHLogo.Suppress = True
  End If
  ' ***** Manali 3.8.0 - Comp Logo in Report
  
  ' ***** Manali 3.10.0 - 12/03/12 - Multi Copmpany Allowed in Scope
  Rep.TxtCoCdFr.SetText ADC("UOCOCDFR")
  If ws_MultiCoMod = False Then
      Rep.TxtCoCdTo.Suppress = True: Rep.TxtCoCdSel.Suppress = True
      Rep.TxtCoCdToLbl.Suppress = True: Rep.TxtCoCdSelLbl.Suppress = True
  Else
      Rep.TxtCoCdToLbl.Suppress = False: Rep.TxtCoCdSelLbl.Suppress = False
      Rep.TxtCoCdTo.SetText ADC("UOCOCDTO"): Rep.TxtCoCdSel.SetText ADC("UOCOCDSEL")
  End If
  ' ***** Manali 3.10.0 - 12/03/12 - Multi Copmpany Allowed in Scope
  
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

Private Sub xxx()
'    wStr = GRP_REP.GrpOpts
'    If InStr(1, "," + wStr + ",", "," + "Kt" + ",") = 0 Then
'        t3_netwt.Suppress = True: t2_netwt.Suppress = True: t1_netwt.Suppress = True: g_netwt.Suppress = True
'    Else
'        warr = Split(wStr)
'        For i = 0 To 2
'           If warr(i) = "Kt" Then
'              OptPos = i
'              Exit For
'            End If
'        Next i
'        wSupGrand = False
'        For i = 0 To 2
'          If i < OptPos Then
'            Dim xxx As CRAXDRT.Report
'            xxx.Sections("G" + CStr(i) + "F").ReportObjects("T" + CStr(i) + "NetWt").Suppress = True
'            wSupGrand = True
'          End If
'        Next i
'        g_netwt.Suppress = True
'    End If
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

