VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{AB2CC32A-E652-42A2-B341-26ECE02F6386}#2.1#0"; "MwfCtl.ocx"
Begin VB.Form EmrFrmHead 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Head Master"
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
   Begin MwfCtl.MWCTL_BTN1 CmdRefFileLst 
      Height          =   465
      Left            =   7050
      TabIndex        =   304
      ToolTipText     =   "Refresh Rates Thourgh ILC"
      Top             =   9330
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   820
      BackColor       =   -2147483633
      MaskColor       =   8454143
      Caption         =   "&Refreh Pict File List"
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
      Height          =   435
      Left            =   5430
      TabIndex        =   284
      TabStop         =   0   'False
      Top             =   9420
      Visible         =   0   'False
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   767
   End
   Begin MwfCtl.MWCTL_ADC ADC 
      CausesValidation=   0   'False
      Height          =   555
      Left            =   60
      TabIndex        =   303
      Top             =   9330
      Width           =   15045
      _ExtentX        =   26538
      _ExtentY        =   979
      BtnWidth        =   885
      BtnHeight       =   435
      BtnMaskColor    =   8454143
      FontSz          =   8.25
   End
   Begin VB.Frame FRA_KEYS 
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   -36
      TabIndex        =   286
      Top             =   30
      Width           =   4380
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   1
         Left            =   1455
         TabIndex        =   331
         ToolTipText     =   "Enter Head Code (Only One Entry Allowed)"
         Top             =   -15
         Visible         =   0   'False
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   503
         Enabled         =   0   'False
         MaxLength       =   8
         DataType        =   4
         DataField       =   "HCd"
         IdName          =   "HCD"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   73
         Left            =   735
         TabIndex        =   1
         ToolTipText     =   "Enter Company Code"
         Top             =   0
         Width           =   690
         _ExtentX        =   1217
         _ExtentY        =   503
         MaxLength       =   3
         DataType        =   4
         DataField       =   "HCoCd"
         IdName          =   "HCOCD"
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
         Height          =   285
         Index           =   0
         Left            =   60
         TabIndex        =   287
         Top             =   0
         Width           =   1065
      End
   End
   Begin VB.Frame FRA_NKEYS 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   9660
      Left            =   -90
      TabIndex        =   285
      Top             =   90
      Width           =   15285
      Begin VB.Frame FraNKeyAll 
         BorderStyle     =   0  'None
         Height          =   9345
         Left            =   120
         TabIndex        =   288
         Top             =   -60
         Width           =   15090
         Begin TabDlg.SSTab TabHd 
            Height          =   8640
            Left            =   30
            TabIndex        =   3
            Top             =   600
            Width           =   15045
            _ExtentX        =   26538
            _ExtentY        =   15240
            _Version        =   393216
            Tabs            =   7
            Tab             =   4
            TabsPerRow      =   7
            TabHeight       =   520
            TabCaption(0)   =   "<&1> Inv Entry Defaults"
            TabPicture(0)   =   "EmrFrmHead.frx":0000
            Tab(0).ControlEnabled=   0   'False
            Tab(0).Control(0)=   "FraTabHd(0)"
            Tab(0).ControlCount=   1
            TabCaption(1)   =   "<&2> Inv Entry Defaults"
            TabPicture(1)   =   "EmrFrmHead.frx":001C
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "FraTabHd(1)"
            Tab(1).ControlCount=   1
            TabCaption(2)   =   "<&3> Inv Entry Defaults"
            TabPicture(2)   =   "EmrFrmHead.frx":0038
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "FraTabHd(2)"
            Tab(2).ControlCount=   1
            TabCaption(3)   =   "<&4> Other Options"
            TabPicture(3)   =   "EmrFrmHead.frx":0054
            Tab(3).ControlEnabled=   0   'False
            Tab(3).Control(0)=   "FraTabHd(3)"
            Tab(3).ControlCount=   1
            TabCaption(4)   =   "<&5> Other Options"
            TabPicture(4)   =   "EmrFrmHead.frx":0070
            Tab(4).ControlEnabled=   -1  'True
            Tab(4).Control(0)=   "FraTabHd(4)"
            Tab(4).Control(0).Enabled=   0   'False
            Tab(4).ControlCount=   1
            TabCaption(5)   =   "<&6> Auto Custom Matching"
            TabPicture(5)   =   "EmrFrmHead.frx":008C
            Tab(5).ControlEnabled=   0   'False
            Tab(5).Control(0)=   "FraTabHd(5)"
            Tab(5).ControlCount=   1
            TabCaption(6)   =   "<&7> jEmr Settings"
            TabPicture(6)   =   "EmrFrmHead.frx":00A8
            Tab(6).ControlEnabled=   0   'False
            Tab(6).Control(0)=   "FraTabHd(6)"
            Tab(6).ControlCount=   1
            Begin VB.Frame FraTabHd 
               BorderStyle     =   0  'None
               Height          =   8160
               Index           =   5
               Left            =   -74940
               TabIndex        =   477
               Top             =   360
               Width           =   14950
               Begin VB.Frame Frame2 
                  Height          =   3705
                  Left            =   -30
                  TabIndex        =   478
                  Top             =   -120
                  Width           =   15045
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   158
                     Left            =   570
                     TabIndex        =   226
                     ToolTipText     =   "Enter Rate Variation % From Range For Diamonds"
                     Top             =   1500
                     Width           =   765
                     _ExtentX        =   1349
                     _ExtentY        =   503
                     Alignment       =   2
                     Mask            =   "-##0.00"
                     MaxLength       =   7
                     DataType        =   2
                     DataField       =   "HRtMinusD"
                     IdName          =   "HRTMINUSD"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   159
                     Left            =   1410
                     TabIndex        =   227
                     ToolTipText     =   "Enter Rate Variation % To Range For Diamonds"
                     Top             =   1500
                     Width           =   765
                     _ExtentX        =   1349
                     _ExtentY        =   503
                     Alignment       =   2
                     Mask            =   "-##0.00"
                     MaxLength       =   7
                     DataType        =   2
                     DataField       =   "HRtPlusD"
                     IdName          =   "HRTPLUSD"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   160
                     Left            =   2400
                     TabIndex        =   228
                     ToolTipText     =   "Enter Pc/ Ct Variation % From Range For Diamonds"
                     Top             =   1500
                     Width           =   765
                     _ExtentX        =   1349
                     _ExtentY        =   503
                     Alignment       =   2
                     Mask            =   "-##0.00"
                     MaxLength       =   7
                     DataType        =   2
                     DataField       =   "HPtrMinusD"
                     IdName          =   "HPTRMINUSD"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   161
                     Left            =   3240
                     TabIndex        =   229
                     ToolTipText     =   "Enter Pc/ Ct Variation % To Range For Diamonds"
                     Top             =   1500
                     Width           =   765
                     _ExtentX        =   1349
                     _ExtentY        =   503
                     Alignment       =   2
                     Mask            =   "-##0.00"
                     MaxLength       =   7
                     DataType        =   2
                     DataField       =   "HPtrPlusD"
                     IdName          =   "HPTRPLUSD"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   162
                     Left            =   570
                     TabIndex        =   233
                     ToolTipText     =   "Enter Rate Variation % From Range For Color Stones"
                     Top             =   1785
                     Width           =   765
                     _ExtentX        =   1349
                     _ExtentY        =   503
                     Alignment       =   2
                     Mask            =   "-##0.00"
                     MaxLength       =   7
                     DataType        =   2
                     DataField       =   "HRtMinusC"
                     IdName          =   "HRTMINUSC"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   163
                     Left            =   1410
                     TabIndex        =   234
                     ToolTipText     =   "Enter Rate Variation % To Range For Color Stones"
                     Top             =   1785
                     Width           =   765
                     _ExtentX        =   1349
                     _ExtentY        =   503
                     Alignment       =   2
                     Mask            =   "-##0.00"
                     MaxLength       =   7
                     DataType        =   2
                     DataField       =   "HRtPlusC"
                     IdName          =   "HRTPLUSC"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   164
                     Left            =   2400
                     TabIndex        =   235
                     ToolTipText     =   "Enter Pc/ Ct Variation % From Range For Color Stones"
                     Top             =   1785
                     Width           =   765
                     _ExtentX        =   1349
                     _ExtentY        =   503
                     Alignment       =   2
                     Mask            =   "-##0.00"
                     MaxLength       =   7
                     DataType        =   2
                     DataField       =   "HPtrMinusC"
                     IdName          =   "HPTRMINUSC"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   165
                     Left            =   3240
                     TabIndex        =   236
                     ToolTipText     =   "Enter Pc/ Ct Variation % To Range For Color Stones"
                     Top             =   1785
                     Width           =   765
                     _ExtentX        =   1349
                     _ExtentY        =   503
                     Alignment       =   2
                     Mask            =   "-##0.00"
                     MaxLength       =   7
                     DataType        =   2
                     DataField       =   "HPtrPlusC"
                     IdName          =   "HPTRPLUSC"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   166
                     Left            =   570
                     TabIndex        =   240
                     ToolTipText     =   "Enter Rate Variation % From Range For Gold"
                     Top             =   2070
                     Width           =   765
                     _ExtentX        =   1349
                     _ExtentY        =   503
                     Alignment       =   2
                     Mask            =   "-##0.00"
                     MaxLength       =   7
                     DataType        =   2
                     DataField       =   "HRtMinusG"
                     IdName          =   "HRTMINUSG"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   167
                     Left            =   1410
                     TabIndex        =   241
                     ToolTipText     =   "Enter Rate Variation % To Range For Gold"
                     Top             =   2070
                     Width           =   765
                     _ExtentX        =   1349
                     _ExtentY        =   503
                     Alignment       =   2
                     Mask            =   "-##0.00"
                     MaxLength       =   7
                     DataType        =   2
                     DataField       =   "HRtPlusG"
                     IdName          =   "HRTPLUSG"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   168
                     Left            =   2400
                     TabIndex        =   242
                     ToolTipText     =   "Enter Pc/ Ct Variation % From Range For Gold"
                     Top             =   2070
                     Width           =   765
                     _ExtentX        =   1349
                     _ExtentY        =   503
                     Alignment       =   2
                     Mask            =   "-##0.00"
                     MaxLength       =   7
                     DataType        =   2
                     DataField       =   "HPtrMinusG"
                     IdName          =   "HPTRMINUSG"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   169
                     Left            =   3240
                     TabIndex        =   243
                     ToolTipText     =   "Enter Pc/ Ct Variation % To Range For Gold"
                     Top             =   2070
                     Width           =   765
                     _ExtentX        =   1349
                     _ExtentY        =   503
                     Alignment       =   2
                     Mask            =   "-##0.00"
                     MaxLength       =   7
                     DataType        =   2
                     DataField       =   "HPtrPlusG"
                     IdName          =   "HPTRPLUSG"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   170
                     Left            =   570
                     TabIndex        =   247
                     ToolTipText     =   "Enter Rate Variation % From Range For Platinum"
                     Top             =   2355
                     Width           =   765
                     _ExtentX        =   1349
                     _ExtentY        =   503
                     Alignment       =   2
                     Mask            =   "-##0.00"
                     MaxLength       =   7
                     DataType        =   2
                     DataField       =   "HRtMinusP"
                     IdName          =   "HRTMINUSP"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   171
                     Left            =   1410
                     TabIndex        =   248
                     ToolTipText     =   "Enter Rate Variation % To Range For Platinum"
                     Top             =   2355
                     Width           =   765
                     _ExtentX        =   1349
                     _ExtentY        =   503
                     Alignment       =   2
                     Mask            =   "-##0.00"
                     MaxLength       =   7
                     DataType        =   2
                     DataField       =   "HRtPlusP"
                     IdName          =   "HRTPLUSP"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   172
                     Left            =   2400
                     TabIndex        =   249
                     ToolTipText     =   "Enter Pc/ Ct Variation % From Range For Platinum"
                     Top             =   2355
                     Width           =   765
                     _ExtentX        =   1349
                     _ExtentY        =   503
                     Alignment       =   2
                     Mask            =   "-##0.00"
                     MaxLength       =   7
                     DataType        =   2
                     DataField       =   "HPtrMinusP"
                     IdName          =   "HPTRMINUSP"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   173
                     Left            =   3240
                     TabIndex        =   250
                     ToolTipText     =   "Enter Pc/ Ct Variation % To Range For Platinum"
                     Top             =   2355
                     Width           =   765
                     _ExtentX        =   1349
                     _ExtentY        =   503
                     Alignment       =   2
                     Mask            =   "-##0.00"
                     MaxLength       =   7
                     DataType        =   2
                     DataField       =   "HPtrPlusP"
                     IdName          =   "HPTRPLUSP"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   174
                     Left            =   570
                     TabIndex        =   254
                     ToolTipText     =   "Enter Rate Variation % From Range For Silver"
                     Top             =   2640
                     Width           =   765
                     _ExtentX        =   1349
                     _ExtentY        =   503
                     Alignment       =   2
                     Mask            =   "-##0.00"
                     MaxLength       =   7
                     DataType        =   2
                     DataField       =   "HRtMinusS"
                     IdName          =   "HRTMINUSS"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   175
                     Left            =   1410
                     TabIndex        =   255
                     ToolTipText     =   "Enter Rate Variation % To Range For Silver"
                     Top             =   2640
                     Width           =   765
                     _ExtentX        =   1349
                     _ExtentY        =   503
                     Alignment       =   2
                     Mask            =   "-##0.00"
                     MaxLength       =   7
                     DataType        =   2
                     DataField       =   "HRtPlusS"
                     IdName          =   "HRTPLUSS"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   176
                     Left            =   2400
                     TabIndex        =   256
                     ToolTipText     =   "Enter Pc/ Ct Variation % From Range For Silver"
                     Top             =   2640
                     Width           =   765
                     _ExtentX        =   1349
                     _ExtentY        =   503
                     Alignment       =   2
                     Mask            =   "-##0.00"
                     MaxLength       =   7
                     DataType        =   2
                     DataField       =   "HPtrMinusS"
                     IdName          =   "HPTRMINUSS"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   177
                     Left            =   3240
                     TabIndex        =   257
                     ToolTipText     =   "Enter Pc/ Ct Variation % To Range For Silver"
                     Top             =   2640
                     Width           =   765
                     _ExtentX        =   1349
                     _ExtentY        =   503
                     Alignment       =   2
                     Mask            =   "-##0.00"
                     MaxLength       =   7
                     DataType        =   2
                     DataField       =   "HPtrPlusS"
                     IdName          =   "HPTRPLUSS"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   178
                     Left            =   570
                     TabIndex        =   268
                     ToolTipText     =   "Enter Rate Variation % From Range For Accessories"
                     Top             =   3210
                     Width           =   765
                     _ExtentX        =   1349
                     _ExtentY        =   503
                     Alignment       =   2
                     Mask            =   "-##0.00"
                     MaxLength       =   7
                     DataType        =   2
                     DataField       =   "HRtMinusX"
                     IdName          =   "HRTMINUSX"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   179
                     Left            =   1410
                     TabIndex        =   269
                     ToolTipText     =   "Enter Rate Variation % To Range For Accessories"
                     Top             =   3210
                     Width           =   765
                     _ExtentX        =   1349
                     _ExtentY        =   503
                     Alignment       =   2
                     Mask            =   "-##0.00"
                     MaxLength       =   7
                     DataType        =   2
                     DataField       =   "HRtPlusX"
                     IdName          =   "HRTPLUSX"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   180
                     Left            =   2400
                     TabIndex        =   270
                     ToolTipText     =   "Enter Pc/ Ct Variation % From Range For Accessories"
                     Top             =   3210
                     Width           =   765
                     _ExtentX        =   1349
                     _ExtentY        =   503
                     Alignment       =   2
                     Mask            =   "-##0.00"
                     MaxLength       =   7
                     DataType        =   2
                     DataField       =   "HPtrMinusX"
                     IdName          =   "HPTRMINUSX"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   181
                     Left            =   3240
                     TabIndex        =   271
                     ToolTipText     =   "Enter Pc/ Ct Variation % To Range For Accessories"
                     Top             =   3210
                     Width           =   765
                     _ExtentX        =   1349
                     _ExtentY        =   503
                     Alignment       =   2
                     Mask            =   "-##0.00"
                     MaxLength       =   7
                     DataType        =   2
                     DataField       =   "HPtrPlusX"
                     IdName          =   "HPTRPLUSX"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   182
                     Left            =   4440
                     TabIndex        =   230
                     ToolTipText     =   "Enter Import Types Not To Be Used For Diamond RM Category"
                     Top             =   1500
                     Width           =   7755
                     _ExtentX        =   13679
                     _ExtentY        =   503
                     MaxLength       =   60
                     DataType        =   4
                     DataField       =   "HImpTypNotInSelD"
                     IdName          =   "HIMPTYPNOTINSELD"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   183
                     Left            =   4440
                     TabIndex        =   237
                     ToolTipText     =   "Enter Import Types Not To Be Used For Color Stone RM Category"
                     Top             =   1785
                     Width           =   7755
                     _ExtentX        =   13679
                     _ExtentY        =   503
                     MaxLength       =   60
                     DataType        =   4
                     DataField       =   "HImpTypNotInSelC"
                     IdName          =   "HIMPTYPNOTINSELC"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   184
                     Left            =   4440
                     TabIndex        =   244
                     ToolTipText     =   "Enter Import Types Not To Be Used For Gold RM Category"
                     Top             =   2070
                     Width           =   7755
                     _ExtentX        =   13679
                     _ExtentY        =   503
                     MaxLength       =   60
                     DataType        =   4
                     DataField       =   "HImpTypNotInSelG"
                     IdName          =   "HIMPTYPNOTINSELG"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   185
                     Left            =   4440
                     TabIndex        =   251
                     ToolTipText     =   "Enter Import Types Not To Be Used For Platinum RM Category"
                     Top             =   2355
                     Width           =   7755
                     _ExtentX        =   13679
                     _ExtentY        =   503
                     MaxLength       =   60
                     DataType        =   4
                     DataField       =   "HImpTypNotInSelP"
                     IdName          =   "HIMPTYPNOTINSELP"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   186
                     Left            =   4440
                     TabIndex        =   258
                     ToolTipText     =   "Enter Import Types Not To Be Used For Silver RM Category"
                     Top             =   2640
                     Width           =   7755
                     _ExtentX        =   13679
                     _ExtentY        =   503
                     MaxLength       =   60
                     DataType        =   4
                     DataField       =   "HImpTypNotInSelS"
                     IdName          =   "HIMPTYPNOTINSELS"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   187
                     Left            =   4440
                     TabIndex        =   272
                     ToolTipText     =   "Enter Import Types Not To Be Used For Accessories RM Category"
                     Top             =   3210
                     Width           =   7755
                     _ExtentX        =   13679
                     _ExtentY        =   503
                     MaxLength       =   60
                     DataType        =   4
                     DataField       =   "HImpTypNotInSelX"
                     IdName          =   "HIMPTYPNOTINSELX"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   188
                     Left            =   12360
                     TabIndex        =   231
                     ToolTipText     =   "Enter Date From Which Purchase Vouchers Have To Be Considered For Diamonds"
                     Top             =   1500
                     Width           =   885
                     _ExtentX        =   1561
                     _ExtentY        =   503
                     MaxLength       =   8
                     DataType        =   3
                     DataField       =   "HImpAftD"
                     IdName          =   "HIMPAFTD"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   189
                     Left            =   12360
                     TabIndex        =   238
                     ToolTipText     =   "Enter Date From Which Purchase Vouchers Have To Be Considered For Color Stones"
                     Top             =   1785
                     Width           =   885
                     _ExtentX        =   1561
                     _ExtentY        =   503
                     MaxLength       =   8
                     DataType        =   3
                     DataField       =   "HImpAftC"
                     IdName          =   "HIMPAFTC"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   190
                     Left            =   12360
                     TabIndex        =   245
                     ToolTipText     =   "Enter Date From Which Purchase Vouchers Have To Be Considered For Gold"
                     Top             =   2070
                     Width           =   885
                     _ExtentX        =   1561
                     _ExtentY        =   503
                     MaxLength       =   8
                     DataType        =   3
                     DataField       =   "HImpAftG"
                     IdName          =   "HIMPAFTG"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   191
                     Left            =   12360
                     TabIndex        =   252
                     ToolTipText     =   "Enter Date From Which Purchase Vouchers Have To Be Considered For Platinum"
                     Top             =   2355
                     Width           =   885
                     _ExtentX        =   1561
                     _ExtentY        =   503
                     MaxLength       =   8
                     DataType        =   3
                     DataField       =   "HImpAftP"
                     IdName          =   "HIMPAFTP"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   192
                     Left            =   12360
                     TabIndex        =   259
                     ToolTipText     =   "Enter Date From Which Purchase Vouchers Have To Be Considered For Silver"
                     Top             =   2640
                     Width           =   885
                     _ExtentX        =   1561
                     _ExtentY        =   503
                     MaxLength       =   8
                     DataType        =   3
                     DataField       =   "HImpAftS"
                     IdName          =   "HIMPAFTS"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   193
                     Left            =   12360
                     TabIndex        =   273
                     ToolTipText     =   "Enter Date From Which Purchase Vouchers Have To Be Considered For Accessories"
                     Top             =   3210
                     Width           =   885
                     _ExtentX        =   1561
                     _ExtentY        =   503
                     MaxLength       =   8
                     DataType        =   3
                     DataField       =   "HImpAftX"
                     IdName          =   "HIMPAFTX"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   200
                     Left            =   13650
                     TabIndex        =   232
                     ToolTipText     =   "Enter Date Upto Which Purchase Vouchers Have To Be Considered For Diamonds"
                     Top             =   1500
                     Width           =   885
                     _ExtentX        =   1561
                     _ExtentY        =   503
                     MaxLength       =   8
                     DataType        =   3
                     DataField       =   "HImpBefD"
                     IdName          =   "HIMPBEFD"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   201
                     Left            =   13650
                     TabIndex        =   239
                     ToolTipText     =   "Enter Date Upto Which Purchase Vouchers Have To Be Considered For Color Stones"
                     Top             =   1785
                     Width           =   885
                     _ExtentX        =   1561
                     _ExtentY        =   503
                     MaxLength       =   8
                     DataType        =   3
                     DataField       =   "HImpBefC"
                     IdName          =   "HIMPBEFC"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   202
                     Left            =   13650
                     TabIndex        =   246
                     ToolTipText     =   "Enter Date Upto Which Purchase Vouchers Have To Be Considered For Gold"
                     Top             =   2070
                     Width           =   885
                     _ExtentX        =   1561
                     _ExtentY        =   503
                     MaxLength       =   8
                     DataType        =   3
                     DataField       =   "HImpBefG"
                     IdName          =   "HIMPBEFG"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   203
                     Left            =   13650
                     TabIndex        =   253
                     ToolTipText     =   "Enter Date Upto Which Purchase Vouchers Have To Be Considered For Platinum"
                     Top             =   2355
                     Width           =   885
                     _ExtentX        =   1561
                     _ExtentY        =   503
                     MaxLength       =   8
                     DataType        =   3
                     DataField       =   "HImpBefP"
                     IdName          =   "HIMPBEFP"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   204
                     Left            =   13650
                     TabIndex        =   260
                     ToolTipText     =   "Enter Date Upto Which Purchase Vouchers Have To Be Considered For Silver"
                     Top             =   2640
                     Width           =   885
                     _ExtentX        =   1561
                     _ExtentY        =   503
                     MaxLength       =   8
                     DataType        =   3
                     DataField       =   "HImpBefS"
                     IdName          =   "HIMPBEFS"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   205
                     Left            =   13650
                     TabIndex        =   274
                     ToolTipText     =   "Enter Date Upto Which Purchase Vouchers Have To Be Considered For Accessories"
                     Top             =   3210
                     Width           =   885
                     _ExtentX        =   1561
                     _ExtentY        =   503
                     MaxLength       =   8
                     DataType        =   3
                     DataField       =   "HImpBefX"
                     IdName          =   "HIMPBEFX"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   208
                     Left            =   570
                     TabIndex        =   261
                     ToolTipText     =   "Enter Rate Variation % From Range For Paladium"
                     Top             =   2925
                     Width           =   765
                     _ExtentX        =   1349
                     _ExtentY        =   503
                     Alignment       =   2
                     Mask            =   "-##0.00"
                     MaxLength       =   7
                     DataType        =   2
                     DataField       =   "HRtMinusL"
                     IdName          =   "HRTMINUSL"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   209
                     Left            =   1410
                     TabIndex        =   262
                     ToolTipText     =   "Enter Rate Variation % To Range For Paladium"
                     Top             =   2925
                     Width           =   765
                     _ExtentX        =   1349
                     _ExtentY        =   503
                     Alignment       =   2
                     Mask            =   "-##0.00"
                     MaxLength       =   7
                     DataType        =   2
                     DataField       =   "HRtPlusL"
                     IdName          =   "HRTPLUSL"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   210
                     Left            =   2400
                     TabIndex        =   263
                     ToolTipText     =   "Enter Pc/ Ct Variation % From Range For Paladium"
                     Top             =   2925
                     Width           =   765
                     _ExtentX        =   1349
                     _ExtentY        =   503
                     Alignment       =   2
                     Mask            =   "-##0.00"
                     MaxLength       =   7
                     DataType        =   2
                     DataField       =   "HPtrMinusL"
                     IdName          =   "HPTRMINUSL"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   211
                     Left            =   3240
                     TabIndex        =   264
                     ToolTipText     =   "Enter Pc/ Ct Variation % To Range For Paladium"
                     Top             =   2925
                     Width           =   765
                     _ExtentX        =   1349
                     _ExtentY        =   503
                     Alignment       =   2
                     Mask            =   "-##0.00"
                     MaxLength       =   7
                     DataType        =   2
                     DataField       =   "HPtrPlusL"
                     IdName          =   "HPTRPLUSL"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   212
                     Left            =   4440
                     TabIndex        =   265
                     ToolTipText     =   "Enter Import Types Not To Be Used For Paladium RM Category"
                     Top             =   2925
                     Width           =   7755
                     _ExtentX        =   13679
                     _ExtentY        =   503
                     MaxLength       =   60
                     DataType        =   4
                     DataField       =   "HImpTypNotInSelL"
                     IdName          =   "HIMPTYPNOTINSELL"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   213
                     Left            =   12360
                     TabIndex        =   266
                     ToolTipText     =   "Enter Date From Which Purchase Vouchers Have To Be Considered For Paladium"
                     Top             =   2925
                     Width           =   885
                     _ExtentX        =   1561
                     _ExtentY        =   503
                     MaxLength       =   8
                     DataType        =   3
                     DataField       =   "HImpAftL"
                     IdName          =   "HIMPAFTL"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   214
                     Left            =   13650
                     TabIndex        =   267
                     ToolTipText     =   "Enter Date Upto Which Purchase Vouchers Have To Be Considered For Paladium"
                     Top             =   2925
                     Width           =   885
                     _ExtentX        =   1561
                     _ExtentY        =   503
                     MaxLength       =   8
                     DataType        =   3
                     DataField       =   "HImpBefL"
                     IdName          =   "HIMPBEFL"
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "L"
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
                     Index           =   154
                     Left            =   150
                     TabIndex        =   504
                     Top             =   2925
                     Width           =   405
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Imports Before"
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   435
                     Index           =   152
                     Left            =   13650
                     TabIndex        =   501
                     Top             =   780
                     Width           =   885
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "X"
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
                     Index           =   147
                     Left            =   150
                     TabIndex        =   494
                     Top             =   3210
                     Width           =   405
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "S"
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
                     Index           =   146
                     Left            =   150
                     TabIndex        =   493
                     Top             =   2640
                     Width           =   405
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "P"
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
                     Index           =   145
                     Left            =   150
                     TabIndex        =   492
                     Top             =   2355
                     Width           =   405
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "G"
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
                     Index           =   144
                     Left            =   150
                     TabIndex        =   491
                     Top             =   2070
                     Width           =   405
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "C"
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
                     Index           =   143
                     Left            =   150
                     TabIndex        =   490
                     Top             =   1785
                     Width           =   405
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "D"
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
                     Index           =   142
                     Left            =   150
                     TabIndex        =   489
                     Top             =   1500
                     Width           =   405
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Upper"
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
                     Index           =   141
                     Left            =   3240
                     TabIndex        =   488
                     Top             =   1065
                     Width           =   765
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Lower"
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
                     Index           =   140
                     Left            =   2430
                     TabIndex        =   487
                     Top             =   1065
                     Width           =   765
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Upper"
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
                     Index           =   139
                     Left            =   1410
                     TabIndex        =   486
                     Top             =   1065
                     Width           =   735
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Lower"
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
                     Index           =   137
                     Left            =   570
                     TabIndex        =   485
                     Top             =   1065
                     Width           =   765
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Do Not Use Import Type(s)"
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
                     Index           =   136
                     Left            =   4440
                     TabIndex        =   484
                     Top             =   780
                     Width           =   2715
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Pc/ Ct Variation (%)"
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
                     Index           =   135
                     Left            =   2400
                     TabIndex        =   483
                     Top             =   780
                     Width           =   1815
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Rate Variation (%)"
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
                     Index           =   134
                     Left            =   570
                     TabIndex        =   482
                     Top             =   780
                     Width           =   1725
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Ctg"
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
                     Index           =   133
                     Left            =   150
                     TabIndex        =   481
                     Top             =   780
                     Width           =   345
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Imports After"
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   465
                     Index           =   138
                     Left            =   12360
                     TabIndex        =   480
                     Top             =   780
                     Width           =   885
                  End
                  Begin VB.Label LblMain 
                     BackStyle       =   0  'Transparent
                     BorderStyle     =   1  'Fixed Single
                     Caption         =   "  Auto Custom Matching Settings"
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
                     Index           =   20
                     Left            =   -30
                     TabIndex        =   479
                     Top             =   90
                     Width           =   15105
                  End
               End
            End
            Begin VB.Frame FraTabHd 
               BorderStyle     =   0  'None
               Height          =   8130
               Index           =   2
               Left            =   -74940
               TabIndex        =   379
               Top             =   360
               Width           =   14950
               Begin VB.Frame Picture1 
                  Height          =   5270
                  Left            =   0
                  TabIndex        =   465
                  Top             =   2880
                  Width           =   15015
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   148
                     Left            =   1665
                     TabIndex        =   78
                     ToolTipText     =   "Enter Local Invoice Footer"
                     Top             =   1245
                     Width           =   12825
                     _ExtentX        =   22622
                     _ExtentY        =   503
                     Alignment       =   3
                     MaxLength       =   90
                     DataField       =   "HLclInvFT"
                     IdName          =   "HLCLINVFT"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   855
                     Index           =   149
                     Left            =   1665
                     TabIndex        =   79
                     ToolTipText     =   "Enter VAT Footer"
                     Top             =   1530
                     Width           =   12825
                     _ExtentX        =   22622
                     _ExtentY        =   1508
                     Alignment       =   3
                     MaxLength       =   0
                     DataField       =   "HLclInvVATFT"
                     IdName          =   "HLCLINVVATFT"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   855
                     Index           =   150
                     Left            =   1665
                     TabIndex        =   80
                     ToolTipText     =   "Enter CST Footer"
                     Top             =   2385
                     Width           =   12825
                     _ExtentX        =   22622
                     _ExtentY        =   1508
                     Alignment       =   3
                     MaxLength       =   0
                     DataField       =   "HLclInvCSTFT"
                     IdName          =   "HLCLINVCSTFT"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   151
                     Left            =   1665
                     TabIndex        =   74
                     ToolTipText     =   "Enter VAT Number"
                     Top             =   675
                     Width           =   8340
                     _ExtentX        =   14711
                     _ExtentY        =   503
                     MaxLength       =   60
                     DataType        =   4
                     DataField       =   "HVATNO"
                     IdName          =   "HVATNO"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   152
                     Left            =   1665
                     TabIndex        =   76
                     ToolTipText     =   "Enter CST Number"
                     Top             =   960
                     Width           =   8340
                     _ExtentX        =   14711
                     _ExtentY        =   503
                     MaxLength       =   60
                     DataType        =   4
                     DataField       =   "HCSTNO"
                     IdName          =   "HCSTNO"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   153
                     Left            =   1665
                     TabIndex        =   71
                     ToolTipText     =   "Enter VAT Percentage"
                     Top             =   390
                     Width           =   585
                     _ExtentX        =   1032
                     _ExtentY        =   503
                     Alignment       =   2
                     Mask            =   "#0.00"
                     MaxLength       =   5
                     DataType        =   2
                     DataField       =   "HVATPER"
                     IdName          =   "HVATPER"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   154
                     Left            =   3975
                     TabIndex        =   72
                     ToolTipText     =   "Enter CST Percentage"
                     Top             =   390
                     Width           =   585
                     _ExtentX        =   1032
                     _ExtentY        =   503
                     Alignment       =   2
                     Mask            =   "#0.00"
                     MaxLength       =   5
                     DataType        =   2
                     DataField       =   "HCstPer"
                     IdName          =   "HCSTPER"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   855
                     Index           =   233
                     Left            =   1665
                     TabIndex        =   81
                     ToolTipText     =   "Enter Invoice Terms and Conditions"
                     Top             =   3240
                     Width           =   12825
                     _ExtentX        =   22622
                     _ExtentY        =   1508
                     Alignment       =   3
                     MaxLength       =   0
                     DataField       =   "HInInvTermCond"
                     IdName          =   "HININVTERMCOND"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   855
                     Index           =   234
                     Left            =   1665
                     TabIndex        =   82
                     ToolTipText     =   "Enter Memo Terms and Conditions"
                     Top             =   4095
                     Width           =   12825
                     _ExtentX        =   22622
                     _ExtentY        =   1508
                     Alignment       =   3
                     MaxLength       =   0
                     DataField       =   "HInMemoTermCond"
                     IdName          =   "HINMEMOTERMCOND"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   249
                     Left            =   6915
                     TabIndex        =   73
                     ToolTipText     =   "Enter Local Invoice Type"
                     Top             =   390
                     Width           =   3090
                     _ExtentX        =   5450
                     _ExtentY        =   503
                     Alignment       =   3
                     MaxLength       =   20
                     DataType        =   4
                     DataField       =   "HInLclInvTyp"
                     IdName          =   "HINLCLINVTYP"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   250
                     Left            =   11385
                     TabIndex        =   75
                     ToolTipText     =   "Enter Tin No"
                     Top             =   675
                     Width           =   3105
                     _ExtentX        =   5477
                     _ExtentY        =   503
                     MaxLength       =   20
                     DataType        =   4
                     DataField       =   "HTinNo"
                     IdName          =   "HTINNO"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   251
                     Left            =   11385
                     TabIndex        =   77
                     ToolTipText     =   "Enter PAN"
                     Top             =   960
                     Width           =   3105
                     _ExtentX        =   5477
                     _ExtentY        =   503
                     MaxLength       =   20
                     DataType        =   4
                     DataField       =   "HPAN"
                     IdName          =   "HPAN"
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "PAN"
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
                     Left            =   10530
                     TabIndex        =   548
                     Top             =   960
                     Width           =   855
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "TIN"
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
                     Index           =   185
                     Left            =   10530
                     TabIndex        =   547
                     Top             =   675
                     Width           =   855
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Local Invoice Type"
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
                     Index           =   184
                     Left            =   4920
                     TabIndex        =   546
                     Top             =   390
                     Width           =   1815
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Memo Terms && Cond"
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   735
                     Index           =   171
                     Left            =   90
                     TabIndex        =   529
                     Top             =   4095
                     Width           =   1365
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Invoice Terms && Cond"
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   735
                     Index           =   170
                     Left            =   90
                     TabIndex        =   528
                     Top             =   3240
                     Width           =   1365
                  End
                  Begin VB.Label LblMain 
                     BackStyle       =   0  'Transparent
                     BorderStyle     =   1  'Fixed Single
                     Caption         =   "  Local Invoice"
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
                     Index           =   19
                     Left            =   -30
                     TabIndex        =   473
                     Top             =   90
                     Width           =   15015
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "CST Footer"
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   405
                     Index           =   123
                     Left            =   90
                     TabIndex        =   472
                     Top             =   2385
                     Width           =   1275
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "VAT Footer"
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
                     Index           =   124
                     Left            =   90
                     TabIndex        =   471
                     Top             =   1530
                     Width           =   1275
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Invoice Footer"
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   315
                     Index           =   125
                     Left            =   90
                     TabIndex        =   470
                     Top             =   1245
                     Width           =   1425
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "VAT Number"
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   315
                     Index           =   126
                     Left            =   90
                     TabIndex        =   469
                     Top             =   690
                     Width           =   1425
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "CST Number"
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   315
                     Index           =   127
                     Left            =   90
                     TabIndex        =   468
                     Top             =   960
                     Width           =   1425
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "VAT %"
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   315
                     Index           =   128
                     Left            =   90
                     TabIndex        =   467
                     Top             =   390
                     Width           =   1425
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "CST %"
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   315
                     Index           =   129
                     Left            =   3180
                     TabIndex        =   466
                     Top             =   390
                     Width           =   765
                  End
               End
               Begin VB.Frame FraDef4 
                  Height          =   3045
                  Left            =   -30
                  TabIndex        =   380
                  Top             =   -120
                  Width           =   15015
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   570
                     Index           =   105
                     Left            =   1665
                     TabIndex        =   63
                     ToolTipText     =   "Enter Exim Policy Details"
                     Top             =   675
                     Width           =   12885
                     _ExtentX        =   22728
                     _ExtentY        =   1005
                     Alignment       =   3
                     MaxLength       =   180
                     DataField       =   "hEximPolicy"
                     IdName          =   "HEXIMPOLICY"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   855
                     Index           =   106
                     Left            =   1665
                     TabIndex        =   64
                     ToolTipText     =   "Enter GSP Eligibility Details"
                     Top             =   1245
                     Width           =   12885
                     _ExtentX        =   22728
                     _ExtentY        =   1508
                     Alignment       =   3
                     MaxLength       =   225
                     DataField       =   "hGspEligibility"
                     IdName          =   "HGSPELIGIBILITY"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   118
                     Left            =   1665
                     TabIndex        =   61
                     ToolTipText     =   "Enter GJEPC Number"
                     Top             =   390
                     Width           =   4170
                     _ExtentX        =   7355
                     _ExtentY        =   503
                     MaxLength       =   30
                     DataType        =   4
                     DataField       =   "hGjepcNo"
                     IdName          =   "HGJEPCNO"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   142
                     Left            =   7020
                     TabIndex        =   62
                     ToolTipText     =   "Enter GJEPC Valid UpTo"
                     Top             =   390
                     Width           =   4350
                     _ExtentX        =   7673
                     _ExtentY        =   503
                     MaxLength       =   30
                     DataField       =   "HGjepcValidUpTo"
                     IdName          =   "HGJEPCVALIDUPTO"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   125
                     Left            =   1665
                     TabIndex        =   65
                     ToolTipText     =   "Enter GSP Package Description"
                     Top             =   2100
                     Width           =   4170
                     _ExtentX        =   7355
                     _ExtentY        =   503
                     MaxLength       =   30
                     DataField       =   "HGspPkgDesc"
                     IdName          =   "HGSPPKGDESC"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   226
                     Left            =   7140
                     TabIndex        =   68
                     ToolTipText     =   "Enter Gsp Officer Address"
                     Top             =   2100
                     Width           =   7410
                     _ExtentX        =   13070
                     _ExtentY        =   503
                     MaxLength       =   45
                     DataField       =   "hInGspAdd1"
                     IdName          =   "HINGSPADD1"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   228
                     Left            =   7140
                     TabIndex        =   69
                     ToolTipText     =   "Enter Gsp Officer Address"
                     Top             =   2385
                     Width           =   7410
                     _ExtentX        =   13070
                     _ExtentY        =   503
                     MaxLength       =   45
                     DataField       =   "hInGspAdd2"
                     IdName          =   "HINGSPADD2"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   229
                     Left            =   7140
                     TabIndex        =   70
                     ToolTipText     =   "Enter Gsp Officer Address"
                     Top             =   2670
                     Width           =   7410
                     _ExtentX        =   13070
                     _ExtentY        =   503
                     MaxLength       =   45
                     DataField       =   "hInGspAdd3"
                     IdName          =   "HINGSPADD3"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   230
                     Left            =   1665
                     TabIndex        =   66
                     ToolTipText     =   "Enter Registration Number"
                     Top             =   2385
                     Width           =   1665
                     _ExtentX        =   2937
                     _ExtentY        =   503
                     MaxLength       =   10
                     DataType        =   4
                     DataField       =   "hInRegNo"
                     IdName          =   "HINREGNO"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   231
                     Left            =   1665
                     TabIndex        =   67
                     ToolTipText     =   "Enter Circular Number"
                     Top             =   2670
                     Width           =   4170
                     _ExtentX        =   7355
                     _ExtentY        =   503
                     MaxLength       =   30
                     DataField       =   "hInCircularNo"
                     IdName          =   "HINCIRCULARNO"
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Circular No"
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
                     Index           =   168
                     Left            =   90
                     TabIndex        =   526
                     Top             =   2670
                     Width           =   1485
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Reg No"
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
                     Index           =   166
                     Left            =   90
                     TabIndex        =   525
                     Top             =   2385
                     Width           =   1485
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "GSP Officer Address"
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   495
                     Index           =   167
                     Left            =   5940
                     TabIndex        =   524
                     Top             =   2100
                     Width           =   1185
                  End
                  Begin VB.Label LblMain 
                     BackStyle       =   0  'Transparent
                     BorderStyle     =   1  'Fixed Single
                     Caption         =   "  DTA Details"
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
                     TabIndex        =   423
                     Top             =   90
                     Width           =   15015
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "GSP Pkg Desc"
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
                     Index           =   111
                     Left            =   120
                     TabIndex        =   422
                     Top             =   2100
                     Width           =   1545
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Valid UpTo"
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
                     Index           =   120
                     Left            =   5910
                     TabIndex        =   416
                     Top             =   390
                     Width           =   1215
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "GJEPC No."
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
                     Index           =   106
                     Left            =   120
                     TabIndex        =   415
                     Top             =   390
                     Width           =   1365
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Exim Policy"
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
                     Index           =   99
                     Left            =   120
                     TabIndex        =   382
                     Top             =   720
                     Width           =   1275
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "GSP Eligibility"
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
                     Index           =   98
                     Left            =   120
                     TabIndex        =   381
                     Top             =   1410
                     Width           =   1275
                  End
               End
            End
            Begin VB.Frame FraTabHd 
               BorderStyle     =   0  'None
               Height          =   8220
               Index           =   3
               Left            =   -74940
               TabIndex        =   353
               Top             =   360
               Width           =   14950
               Begin VB.Frame FrapPpcDet 
                  Height          =   2805
                  Left            =   0
                  TabIndex        =   384
                  Top             =   5520
                  Width           =   5775
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   84
                     Left            =   2715
                     TabIndex        =   104
                     ToolTipText     =   "Specify Whether Ppc Module Has To Be Shown ? (Y/N)"
                     Top             =   390
                     Width           =   315
                     _ExtentX        =   556
                     _ExtentY        =   503
                     DataType        =   4
                     DataField       =   "hPpcYN"
                     IdName          =   "HPPCYN"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   87
                     Left            =   2715
                     TabIndex        =   109
                     ToolTipText     =   "Enter Line Code"
                     Top             =   1815
                     Width           =   885
                     _ExtentX        =   1561
                     _ExtentY        =   503
                     MaxLength       =   5
                     DataType        =   4
                     DataField       =   "hLine"
                     IdName          =   "HLINE"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   83
                     Left            =   2715
                     TabIndex        =   106
                     ToolTipText     =   "Enter Max Order Quantity"
                     Top             =   960
                     Width           =   885
                     _ExtentX        =   1561
                     _ExtentY        =   503
                     Mask            =   "#####0.0"
                     MaxLength       =   8
                     DataType        =   2
                     DataField       =   "hMaxOrdQty"
                     IdName          =   "HMAXORDQTY"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   86
                     Left            =   2715
                     TabIndex        =   107
                     ToolTipText     =   "Enter Min Order Quantity"
                     Top             =   1245
                     Width           =   885
                     _ExtentX        =   1561
                     _ExtentY        =   503
                     Mask            =   "#####0.0"
                     MaxLength       =   8
                     DataType        =   2
                     DataField       =   "hMinOrdQty"
                     IdName          =   "HMINORDQTY"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   88
                     Left            =   2715
                     TabIndex        =   110
                     ToolTipText     =   "Enter Setting Option (Wax Set/ Hand Set)"
                     Top             =   2100
                     Width           =   495
                     _ExtentX        =   873
                     _ExtentY        =   503
                     MaxLength       =   2
                     DataType        =   4
                     DataField       =   "hWh"
                     IdName          =   "HWH"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   82
                     Left            =   2715
                     TabIndex        =   105
                     ToolTipText     =   "Enter Max Line Utilization"
                     Top             =   675
                     Width           =   885
                     _ExtentX        =   1561
                     _ExtentY        =   503
                     Mask            =   "##0.000"
                     MaxLength       =   7
                     DataType        =   2
                     DataField       =   "hMaxLineUtil"
                     IdName          =   "HMAXLINEUTIL"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   85
                     Left            =   2715
                     TabIndex        =   108
                     ToolTipText     =   "Enter Min Order Percentage"
                     Top             =   1530
                     Width           =   885
                     _ExtentX        =   1561
                     _ExtentY        =   503
                     Mask            =   "##0.000"
                     MaxLength       =   7
                     DataType        =   2
                     DataField       =   "hMinOrdprcnt"
                     IdName          =   "HMINORDPRCNT"
                  End
                  Begin VB.Label LblMain 
                     BackStyle       =   0  'Transparent
                     BorderStyle     =   1  'Fixed Single
                     Caption         =   "  Ppc Defaults"
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
                     Left            =   -30
                     TabIndex        =   424
                     Top             =   90
                     Width           =   5895
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Min Ord % To Plan For"
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
                     Index           =   81
                     Left            =   150
                     TabIndex        =   391
                     Top             =   1530
                     Width           =   2445
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Max Ord Qty To Plan For "
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
                     Index           =   79
                     Left            =   150
                     TabIndex        =   390
                     Top             =   960
                     Width           =   2565
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Line"
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
                     Index           =   83
                     Left            =   150
                     TabIndex        =   389
                     Top             =   1815
                     Width           =   2445
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Min Ord Qty To Plan For"
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
                     Index           =   82
                     Left            =   150
                     TabIndex        =   388
                     Top             =   1245
                     Width           =   2505
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Ppc Module"
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
                     Index           =   80
                     Left            =   150
                     TabIndex        =   387
                     Top             =   390
                     Width           =   2445
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Max. Line Utilization %"
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
                     Index           =   78
                     Left            =   150
                     TabIndex        =   386
                     Top             =   675
                     Width           =   2445
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Setting Option"
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
                     Index           =   84
                     Left            =   150
                     TabIndex        =   385
                     Top             =   2100
                     Width           =   2445
                  End
               End
               Begin VB.Frame FraDsgDet 
                  Height          =   2595
                  Left            =   -30
                  TabIndex        =   360
                  Top             =   2910
                  Width           =   5775
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   6
                     Left            =   2715
                     TabIndex        =   94
                     ToolTipText     =   "Enter Default Design Production Sequence"
                     Top             =   780
                     Width           =   2970
                     _ExtentX        =   5239
                     _ExtentY        =   503
                     MaxLength       =   100
                     DataType        =   4
                     DataField       =   "hDmPrdSeq"
                     IdName          =   "HDMPRDSEQ"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   39
                     Left            =   2715
                     TabIndex        =   95
                     ToolTipText     =   "Enter Default Design Value Addition Category"
                     Top             =   1065
                     Width           =   855
                     _ExtentX        =   1508
                     _ExtentY        =   503
                     MaxLength       =   5
                     DataType        =   4
                     DataField       =   "hDmVaCtg"
                     IdName          =   "HDMVACTG"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   196
                     Left            =   2715
                     TabIndex        =   97
                     ToolTipText     =   "Enter Y for SetCd Validation / N for Current SetCd Help"
                     Top             =   1350
                     Width           =   315
                     _ExtentX        =   556
                     _ExtentY        =   503
                     DataType        =   4
                     DataField       =   "HValidDsgSetCd"
                     IdName          =   "HVALIDDSGSETCD"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   217
                     Left            =   2715
                     TabIndex        =   98
                     ToolTipText     =   "Enter Valid Size has to be checked from Masters"
                     Top             =   1635
                     Width           =   315
                     _ExtentX        =   556
                     _ExtentY        =   503
                     DataType        =   4
                     DataField       =   "HChkVldRmLnYN"
                     IdName          =   "HCHKVLDRMLNYN"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   220
                     Left            =   2715
                     TabIndex        =   99
                     ToolTipText     =   "Implement Design Catalogue Module"
                     Top             =   1920
                     Width           =   315
                     _ExtentX        =   556
                     _ExtentY        =   503
                     DataType        =   4
                     DataField       =   "HDsgCatYN"
                     IdName          =   "HDSGCATYN"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   223
                     Left            =   2715
                     TabIndex        =   100
                     ToolTipText     =   "Enter Maximum Records to be displayed in Detailed Design View and Order Zoom"
                     Top             =   2205
                     Width           =   600
                     _ExtentX        =   1058
                     _ExtentY        =   503
                     Alignment       =   2
                     Mask            =   "###0"
                     MaxLength       =   4
                     DataType        =   1
                     DataField       =   "HMaxDsgViewRec"
                     IdName          =   "HMAXDSGVIEWREC"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   258
                     Left            =   5325
                     TabIndex        =   101
                     ToolTipText     =   "Enter Whether to copy Actual Wt Or Calculated Wt from FG"
                     Top             =   2220
                     Width           =   315
                     _ExtentX        =   556
                     _ExtentY        =   503
                     DataType        =   4
                     DataField       =   "hCpyActCalcWtYN"
                     IdName          =   "HCPYACTCALCWTYN"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   297
                     Left            =   2715
                     TabIndex        =   93
                     ToolTipText     =   "Enter Default Design Process Sequence"
                     Top             =   480
                     Width           =   2970
                     _ExtentX        =   5239
                     _ExtentY        =   503
                     MaxLength       =   100
                     DataType        =   4
                     DataField       =   "hDmPrcsSeq"
                     IdName          =   "HDMPRCSSEQ"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   299
                     Left            =   4200
                     TabIndex        =   96
                     ToolTipText     =   "Enter Default Cell"
                     Top             =   1065
                     Width           =   1455
                     _ExtentX        =   2566
                     _ExtentY        =   503
                     MaxLength       =   5
                     DataType        =   4
                     DataField       =   "hCell"
                     IdName          =   "HCELL"
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Cell"
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
                     Index           =   232
                     Left            =   3720
                     TabIndex        =   632
                     Top             =   1080
                     Width           =   615
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Process Sequence"
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
                     Index           =   230
                     Left            =   150
                     TabIndex        =   630
                     Top             =   510
                     Width           =   2085
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Cpy Act/Calc Wt"
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
                     Index           =   192
                     Left            =   3600
                     TabIndex        =   555
                     Top             =   2220
                     Width           =   1725
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Max Limit for Dsg View"
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
                     Index           =   163
                     Left            =   150
                     TabIndex        =   513
                     Top             =   2205
                     Width           =   3075
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Dsg Catalogue Module"
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
                     Index           =   160
                     Left            =   150
                     TabIndex        =   510
                     Top             =   1920
                     Width           =   2475
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Check Valid RM Size"
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
                     Index           =   157
                     Left            =   150
                     TabIndex        =   506
                     Top             =   1635
                     Width           =   2565
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Dsg Set Code From Param"
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
                     Index           =   150
                     Left            =   150
                     TabIndex        =   497
                     Top             =   1350
                     Width           =   2565
                  End
                  Begin VB.Label LblMain 
                     BackStyle       =   0  'Transparent
                     BorderStyle     =   1  'Fixed Single
                     Caption         =   "  Design Defaults"
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
                     Index           =   2
                     Left            =   -30
                     TabIndex        =   425
                     Top             =   90
                     Width           =   5925
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Value Addn Ctg"
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
                     Index           =   35
                     Left            =   150
                     TabIndex        =   362
                     Top             =   1065
                     Width           =   2085
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Production Sequence"
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
                     Left            =   150
                     TabIndex        =   361
                     Top             =   780
                     Width           =   2085
                  End
               End
               Begin VB.Frame FraCustDet 
                  Height          =   3165
                  Left            =   0
                  TabIndex        =   363
                  Top             =   -120
                  Width           =   5775
                  Begin VB.Frame FraOrdAct 
                     Height          =   1695
                     Left            =   0
                     TabIndex        =   367
                     Top             =   1440
                     Width           =   5445
                     Begin MwfCtl.MWCTL_MED ATXT 
                        Height          =   285
                        Index           =   8
                        Left            =   2595
                        TabIndex        =   89
                        ToolTipText     =   "Enter All Weights From Order (Y / N)"
                        Top             =   480
                        Width           =   315
                        _ExtentX        =   556
                        _ExtentY        =   503
                        DataType        =   4
                        DataField       =   "HCmIWtEqOrd"
                        ReCalcParent    =   "HCMIWTFRORD"
                        IdName          =   "HCMIWTEQORD"
                     End
                     Begin MwfCtl.MWCTL_MED ATXT 
                        Height          =   285
                        Index           =   9
                        Left            =   2595
                        TabIndex        =   90
                        ToolTipText     =   "Enter Weights From Order"
                        Top             =   765
                        Width           =   885
                        _ExtentX        =   1561
                        _ExtentY        =   503
                        MaxLength       =   5
                        DataType        =   4
                        DataField       =   "HCmIWtFrOrd"
                        ReCalcOn        =   "HCMIWTEQORD"
                        IdName          =   "HCMIWTFRORD"
                     End
                     Begin MwfCtl.MWCTL_MED ATXT 
                        Height          =   285
                        Index           =   10
                        Left            =   2595
                        TabIndex        =   91
                        ToolTipText     =   "Enter Invoice Gold Rate From Order (Y/N)"
                        Top             =   1050
                        Width           =   315
                        _ExtentX        =   556
                        _ExtentY        =   503
                        DataType        =   4
                        DataField       =   "hCmIGldRtEqOrd"
                        IdName          =   "HCMIGLDRTEQORD"
                     End
                     Begin MwfCtl.MWCTL_MED ATXT 
                        Height          =   285
                        Index           =   198
                        Left            =   2595
                        TabIndex        =   92
                        ToolTipText     =   "Invoice Labor Weight From Order (Y/N)"
                        Top             =   1335
                        Width           =   315
                        _ExtentX        =   556
                        _ExtentY        =   503
                        DataType        =   4
                        DataField       =   "hCmILabWtFrOrd"
                        IdName          =   "HCMILABWTFRORD"
                     End
                     Begin VB.Label Label18 
                        BackStyle       =   0  'Transparent
                        Caption         =   "Lab Wt From Ord"
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
                        Left            =   180
                        TabIndex        =   500
                        Top             =   1335
                        Width           =   2295
                     End
                     Begin VB.Label LblMain 
                        BackStyle       =   0  'Transparent
                        BorderStyle     =   1  'Fixed Single
                        Caption         =   "  Inv Wt From Order / Actuals"
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
                        Index           =   4
                        Left            =   -30
                        TabIndex        =   427
                        Top             =   90
                        Width           =   5445
                     End
                     Begin VB.Label ALBL 
                        BackStyle       =   0  'Transparent
                        Caption         =   "All Wts From Ord "
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
                        Left            =   180
                        TabIndex        =   370
                        Top             =   480
                        Width           =   1755
                     End
                     Begin VB.Label ALBL 
                        BackStyle       =   0  'Transparent
                        Caption         =   "Wt From Ord "
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
                        Left            =   180
                        TabIndex        =   369
                        Top             =   765
                        Width           =   1755
                     End
                     Begin VB.Label ALBL 
                        BackStyle       =   0  'Transparent
                        Caption         =   "Gld Rt From Ord"
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
                        Left            =   180
                        TabIndex        =   368
                        Top             =   1050
                        Width           =   1725
                     End
                  End
                  Begin VB.Frame FraFavAvg 
                     Height          =   1695
                     Left            =   5520
                     TabIndex        =   364
                     Top             =   1350
                     Visible         =   0   'False
                     Width           =   345
                     Begin MwfCtl.MWCTL_MED ATXT 
                        Height          =   285
                        Index           =   12
                        Left            =   1275
                        TabIndex        =   103
                        ToolTipText     =   "Enter Colour Stone Rate Favourable  (Y / N)"
                        Top             =   765
                        Width           =   315
                        _ExtentX        =   556
                        _ExtentY        =   503
                        DataType        =   4
                        DataField       =   "HCmCsRtFA"
                        IdName          =   "HCMCSRTFA"
                     End
                     Begin MwfCtl.MWCTL_MED ATXT 
                        Height          =   285
                        Index           =   11
                        Left            =   1275
                        TabIndex        =   102
                        ToolTipText     =   "Enter Diamond Rate Favourable  (Y / N)"
                        Top             =   480
                        Width           =   315
                        _ExtentX        =   556
                        _ExtentY        =   503
                        DataType        =   4
                        DataField       =   "HCmDiaRtFA"
                        IdName          =   "HCMDIARTFA"
                     End
                     Begin VB.Label LblMain 
                        BackStyle       =   0  'Transparent
                        BorderStyle     =   1  'Fixed Single
                        Caption         =   "  Favourable / Average"
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
                        Index           =   5
                        Left            =   -30
                        TabIndex        =   428
                        Top             =   90
                        Width           =   1905
                     End
                     Begin VB.Label ALBL 
                        BackStyle       =   0  'Transparent
                        Caption         =   "Dia Rt "
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
                        Left            =   180
                        TabIndex        =   366
                        Top             =   480
                        Width           =   1155
                     End
                     Begin VB.Label ALBL 
                        BackStyle       =   0  'Transparent
                        Caption         =   "ColStn Rt "
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
                        Index           =   11
                        Left            =   180
                        TabIndex        =   365
                        Top             =   765
                        Width           =   1035
                     End
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   255
                     Index           =   14
                     Left            =   2595
                     TabIndex        =   84
                     ToolTipText     =   "Specify How The Value Of Gold Should Be Calculated As In Order "
                     Top             =   720
                     Width           =   855
                     _ExtentX        =   1508
                     _ExtentY        =   450
                     MaxLength       =   5
                     DataType        =   4
                     DataField       =   "hCmGldAs"
                     IdName          =   "HCMGLDAS"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   7
                     Left            =   2595
                     TabIndex        =   83
                     ToolTipText     =   "Specify Whether To Fix Inv/Customs Price = Order Sale Value During Invoicing Or Calculate The Price (Y / N)"
                     Top             =   420
                     Width           =   315
                     _ExtentX        =   556
                     _ExtentY        =   503
                     DataType        =   4
                     DataField       =   "hCmFixPrc"
                     IdName          =   "HCMFIXPRC"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   56
                     Left            =   5400
                     TabIndex        =   88
                     ToolTipText     =   "Sort Customer Code Help On Customer Name/ Customer Code  (Enter Y/ N)"
                     Top             =   360
                     Visible         =   0   'False
                     Width           =   285
                     _ExtentX        =   503
                     _ExtentY        =   503
                     Enabled         =   0   'False
                     DataType        =   4
                     DataField       =   "HCmCdHlp"
                     IdName          =   "HCMCDHLP"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   133
                     Left            =   4800
                     TabIndex        =   85
                     ToolTipText     =   "Specify How The Value Of Weight Wise Labour Should Be Calculated As In Order "
                     Top             =   720
                     Width           =   855
                     _ExtentX        =   1508
                     _ExtentY        =   503
                     MaxLength       =   5
                     DataType        =   4
                     DataField       =   "hCmLabAs"
                     IdName          =   "HCMLABAS"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   268
                     Left            =   2595
                     TabIndex        =   86
                     ToolTipText     =   "Specify How The Value Of Chain For Gold Should Be Calculated As In Order "
                     Top             =   1005
                     Width           =   855
                     _ExtentX        =   1508
                     _ExtentY        =   503
                     MaxLength       =   45
                     DataType        =   4
                     DataField       =   "hChnSubCtgForGld"
                     IdName          =   "HCHNSUBCTGFORGLD"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   269
                     Left            =   4800
                     TabIndex        =   87
                     ToolTipText     =   "Specify How The Value Of Chain For Weight Wise Labour Should Be Calculated As In Order "
                     Top             =   1005
                     Width           =   855
                     _ExtentX        =   1508
                     _ExtentY        =   503
                     MaxLength       =   45
                     DataType        =   4
                     DataField       =   "hChnSubCtgForLab"
                     IdName          =   "HCHNSUBCTGFORLAB"
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Chn Lab As"
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
                     Index           =   202
                     Left            =   3600
                     TabIndex        =   573
                     Top             =   1005
                     Width           =   1095
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Chain SCtg For Gold As"
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   315
                     Index           =   201
                     Left            =   120
                     TabIndex        =   572
                     Top             =   1005
                     Width           =   2535
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Lab As"
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
                     Index           =   115
                     Left            =   3840
                     TabIndex        =   453
                     Top             =   720
                     Width           =   735
                  End
                  Begin VB.Label LblMain 
                     BackStyle       =   0  'Transparent
                     BorderStyle     =   1  'Fixed Single
                     Caption         =   "  Customer Defaults AND Chain Logic"
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
                     Index           =   3
                     Left            =   -30
                     TabIndex        =   426
                     Top             =   90
                     Width           =   5925
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Fix Prc"
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
                     Left            =   120
                     TabIndex        =   373
                     Top             =   420
                     Width           =   2535
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Gold As"
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
                     Index           =   14
                     Left            =   120
                     TabIndex        =   372
                     Top             =   720
                     Width           =   2535
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Sort Cust Help On Name"
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
                     Index           =   55
                     Left            =   3000
                     TabIndex        =   371
                     Top             =   360
                     Visible         =   0   'False
                     Width           =   2355
                  End
               End
               Begin VB.Frame FraTxnDet 
                  Height          =   4665
                  Left            =   5700
                  TabIndex        =   354
                  Top             =   3660
                  Width           =   9285
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   41
                     Left            =   3945
                     TabIndex        =   132
                     ToolTipText     =   "Enter Default Bag Opening Location"
                     Top             =   390
                     Width           =   1215
                     _ExtentX        =   2143
                     _ExtentY        =   503
                     MaxLength       =   8
                     DataType        =   4
                     DataField       =   "hBOpnLoc"
                     IdName          =   "HBOPNLOC"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   42
                     Left            =   3945
                     TabIndex        =   136
                     ToolTipText     =   "Specify Whether Acknowledgement Is Required In Bag Movement"
                     Top             =   960
                     Width           =   285
                     _ExtentX        =   503
                     _ExtentY        =   503
                     DataType        =   4
                     DataField       =   "hTdAck"
                     IdName          =   "HTDACK"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   49
                     Left            =   3945
                     TabIndex        =   134
                     ToolTipText     =   "Enter Default Bag Character"
                     Top             =   675
                     Width           =   585
                     _ExtentX        =   1032
                     _ExtentY        =   503
                     MaxLength       =   3
                     DataType        =   4
                     DataField       =   "hBChr"
                     IdName          =   "HBCHR"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   50
                     Left            =   3945
                     TabIndex        =   138
                     ToolTipText     =   "Show Bags Of All Locations In The Help Of Bag (Enter Y/ N)"
                     Top             =   1245
                     Width           =   285
                     _ExtentX        =   503
                     _ExtentY        =   503
                     DataType        =   4
                     DataField       =   "hOthLocBag"
                     IdName          =   "HOTHLOCBAG"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   74
                     Left            =   3945
                     TabIndex        =   140
                     ToolTipText     =   "Show Current Gross Wt in Bag Split (Enter Y/ N)"
                     Top             =   1530
                     Width           =   285
                     _ExtentX        =   503
                     _ExtentY        =   503
                     DataType        =   4
                     DataField       =   "hGrsWtInBsYN"
                     IdName          =   "HGRSWTINBSYN"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   95
                     Left            =   3945
                     TabIndex        =   142
                     ToolTipText     =   "Allow Editing of Bag Number? (Y/N)"
                     Top             =   1815
                     Width           =   285
                     _ExtentX        =   503
                     _ExtentY        =   503
                     DataType        =   4
                     DataField       =   "hBagNoEditYN"
                     IdName          =   "HBAGNOEDITYN"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   101
                     Left            =   3945
                     TabIndex        =   144
                     ToolTipText     =   "Allow Editing of Production Points? (Y/N)"
                     Top             =   2100
                     Width           =   285
                     _ExtentX        =   503
                     _ExtentY        =   503
                     DataType        =   4
                     DataField       =   "HPrdPtsYN"
                     IdName          =   "HPRDPTSYN"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   122
                     Left            =   3945
                     TabIndex        =   146
                     ToolTipText     =   "Allow Flute Bags? (Y/N)"
                     Top             =   2385
                     Width           =   285
                     _ExtentX        =   503
                     _ExtentY        =   503
                     DataType        =   4
                     DataField       =   "HFLUTEBAGYN"
                     IdName          =   "HFLUTEBAGYN"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   144
                     Left            =   8355
                     TabIndex        =   147
                     ToolTipText     =   "Allow Component Bags (Yes/ No)"
                     Top             =   2400
                     Width           =   285
                     _ExtentX        =   503
                     _ExtentY        =   503
                     DataType        =   4
                     DataField       =   "hCompBagsYn"
                     IdName          =   "HCOMPBAGSYN"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   237
                     Left            =   8355
                     TabIndex        =   135
                     ToolTipText     =   "Enter Default Metler Code For All Users"
                     Top             =   675
                     Width           =   855
                     _ExtentX        =   1508
                     _ExtentY        =   503
                     MaxLength       =   5
                     DataType        =   4
                     DataField       =   "HMetlrCd"
                     IdName          =   "HMETLRCD"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   254
                     Left            =   8355
                     TabIndex        =   137
                     ToolTipText     =   "Whether Weight tolerance to be applied in Percentage. Enter D,C or D or C"
                     Top             =   960
                     Width           =   855
                     _ExtentX        =   1508
                     _ExtentY        =   503
                     Enabled         =   0   'False
                     MaxLength       =   3
                     DataType        =   4
                     DataField       =   "hTolWtPctDC"
                     IdName          =   "HTOLWTPCTDC"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   256
                     Left            =   8355
                     TabIndex        =   141
                     ToolTipText     =   "Whether tolerance to be applied in Percentage. Enter D,C or D or C"
                     Top             =   1530
                     Width           =   285
                     _ExtentX        =   503
                     _ExtentY        =   503
                     MaxLength       =   3
                     DataType        =   4
                     DataField       =   "hSetTypOnlyFrOrdYN"
                     IdName          =   "HSETTYPONLYFRORDYN"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   257
                     Left            =   8355
                     TabIndex        =   139
                     ToolTipText     =   "Whether Quantity tolerance to be applied in Percentage. Enter D,C or D or C"
                     Top             =   1245
                     Width           =   855
                     _ExtentX        =   1508
                     _ExtentY        =   503
                     Enabled         =   0   'False
                     MaxLength       =   3
                     DataType        =   4
                     DataField       =   "hTolQtyPctDC"
                     IdName          =   "HTOLQTYPCTDC"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   259
                     Left            =   8355
                     TabIndex        =   143
                     ToolTipText     =   "Show Time Details in Bag History/Bag Movement Pcs Report"
                     Top             =   1830
                     Width           =   285
                     _ExtentX        =   503
                     _ExtentY        =   503
                     DataType        =   4
                     DataField       =   "hShowBgHistDetsYN"
                     IdName          =   "HSHOWBGHISTDETSYN"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   261
                     Left            =   8355
                     TabIndex        =   145
                     ToolTipText     =   "Apply Checks for Holidays"
                     Top             =   2115
                     Width           =   285
                     _ExtentX        =   503
                     _ExtentY        =   503
                     DataType        =   4
                     DataField       =   "HHolChkYN"
                     IdName          =   "HHOLCHKYN"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   120
                     Left            =   3945
                     TabIndex        =   148
                     ToolTipText     =   "Restrict Production Points Quantity On Bag Rm Qty? (Y/N)"
                     Top             =   2670
                     Width           =   285
                     _ExtentX        =   503
                     _ExtentY        =   503
                     DataType        =   4
                     DataField       =   "HOrdPrdPts"
                     IdName          =   "HORDPRDPTS"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   330
                     Index           =   224
                     Left            =   3945
                     TabIndex        =   151
                     ToolTipText     =   "Enter Process Location for applying Minimum Tolerance on Diamonds"
                     Top             =   3315
                     Width           =   5265
                     _ExtentX        =   9287
                     _ExtentY        =   582
                     MaxLength       =   60
                     DataType        =   4
                     DataField       =   "HMinDiaProcess"
                     IdName          =   "HMINDIAPROCESS"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   330
                     Index           =   263
                     Left            =   3945
                     TabIndex        =   152
                     ToolTipText     =   "Enter Process Location for applying Wax Pullout Tolerance"
                     Top             =   3645
                     Width           =   5265
                     _ExtentX        =   9287
                     _ExtentY        =   582
                     MaxLength       =   60
                     DataType        =   4
                     DataField       =   "hChkWaxPullOutLocs"
                     IdName          =   "HCHKWAXPULLOUTLOCS"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   570
                     Index           =   264
                     Left            =   3945
                     TabIndex        =   153
                     ToolTipText     =   "Enter Process Location for applying Casting Wt Tolerance"
                     Top             =   3975
                     Width           =   5265
                     _ExtentX        =   9287
                     _ExtentY        =   1005
                     MaxLength       =   60
                     DataType        =   4
                     DataField       =   "hChkCastLocs"
                     IdName          =   "HCHKCASTLOCS"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   265
                     Left            =   8355
                     TabIndex        =   149
                     ToolTipText     =   "Restrict Production Points Quantity On Bag Qty? (Y/N)"
                     Top             =   2685
                     Width           =   285
                     _ExtentX        =   503
                     _ExtentY        =   503
                     DataType        =   4
                     DataField       =   "hPrdPtsOnBQtyYN"
                     IdName          =   "HPRDPTSONBQTYYN"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   294
                     Left            =   8355
                     TabIndex        =   133
                     ToolTipText     =   "Diamond Min / Max wt Tolerance to be rounded off to 2 digits Y/N?"
                     Top             =   360
                     Width           =   285
                     _ExtentX        =   503
                     _ExtentY        =   503
                     MaxLength       =   3
                     DataType        =   4
                     DataField       =   "hRndOffOdTolWtYN"
                     IdName          =   "HRNDOFFODTOLWTYN"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   298
                     Left            =   3945
                     TabIndex        =   150
                     ToolTipText     =   "Enter Default Bag Opening Location"
                     Top             =   3000
                     Width           =   1215
                     _ExtentX        =   2143
                     _ExtentY        =   503
                     MaxLength       =   8
                     DataType        =   4
                     DataField       =   "HDefRepBagLoc"
                     IdName          =   "HDEFREPBAGLOC"
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Bag Repair Loc"
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
                     Index           =   231
                     Left            =   150
                     TabIndex        =   631
                     Top             =   3000
                     Width           =   3855
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   " Dia wt Tol to be rnd off to 2 dgt"
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
                     Index           =   220
                     Left            =   5160
                     TabIndex        =   626
                     Top             =   360
                     Width           =   3015
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Restrict Prd Pts Qty On Bag RmQty"
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
                     Index           =   108
                     Left            =   150
                     TabIndex        =   624
                     Top             =   2670
                     Width           =   3795
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Process Loc for Min Dia Tolerance "
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
                     Index           =   164
                     Left            =   150
                     TabIndex        =   623
                     Top             =   3315
                     Width           =   3855
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Process Loc for Wax PullOut Tolerance"
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
                     Index           =   196
                     Left            =   150
                     TabIndex        =   622
                     Top             =   3690
                     Width           =   3855
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Process Loc for Cast Wt Tolerance"
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
                     Index           =   197
                     Left            =   150
                     TabIndex        =   621
                     Top             =   4020
                     Width           =   3855
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Restrict Prd Pts Qty on Bag Qty"
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
                     Index           =   198
                     Left            =   5280
                     TabIndex        =   620
                     Top             =   2685
                     Width           =   2925
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Apply Checks for Holidays"
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
                     Index           =   195
                     Left            =   5610
                     TabIndex        =   558
                     Top             =   2115
                     Width           =   2565
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Show Time Dets in Reports"
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
                     Index           =   193
                     Left            =   5610
                     TabIndex        =   556
                     Top             =   1830
                     Width           =   2565
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Qty Tolerance in % for"
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
                     Index           =   191
                     Left            =   5970
                     TabIndex        =   554
                     Top             =   1245
                     Width           =   2205
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Setting Typ Only fr Ord"
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
                     Index           =   190
                     Left            =   5970
                     TabIndex        =   553
                     Top             =   1530
                     Width           =   2205
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Wt Tolerance in % for"
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
                     Index           =   189
                     Left            =   5970
                     TabIndex        =   551
                     Top             =   960
                     Width           =   2205
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Default Metler Code"
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
                     Index           =   172
                     Left            =   6150
                     TabIndex        =   532
                     Top             =   675
                     Width           =   2355
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Allow Component Bags"
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   255
                     Index           =   121
                     Left            =   5610
                     TabIndex        =   462
                     Top             =   2400
                     Width           =   2475
                  End
                  Begin VB.Label LblMain 
                     BackStyle       =   0  'Transparent
                     BorderStyle     =   1  'Fixed Single
                     Caption         =   "  Bag/ Rm Transaction Details"
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
                     Index           =   6
                     Left            =   -30
                     TabIndex        =   429
                     Top             =   90
                     Width           =   9180
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Allow Flute Bags"
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   255
                     Index           =   109
                     Left            =   150
                     TabIndex        =   418
                     Top             =   2385
                     Width           =   3795
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Allow Editing of Production Points"
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
                     Index           =   95
                     Left            =   150
                     TabIndex        =   399
                     Top             =   2100
                     Width           =   3795
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Allow Editing of Bag No."
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
                     Index           =   89
                     Left            =   150
                     TabIndex        =   396
                     Top             =   1815
                     Width           =   3795
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Bag Opening Loc"
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
                     Index           =   36
                     Left            =   150
                     TabIndex        =   359
                     Top             =   390
                     Width           =   3855
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Acknowledgment Reqd In Bag Movmt"
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
                     Index           =   39
                     Left            =   150
                     TabIndex        =   358
                     Top             =   960
                     Width           =   4035
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Bag Chr In Bag Entries"
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
                     Index           =   44
                     Left            =   150
                     TabIndex        =   357
                     Top             =   675
                     Width           =   3855
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Allow Oth Loc Bags In Bag Txn Entries "
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
                     Index           =   46
                     Left            =   150
                     TabIndex        =   356
                     Top             =   1245
                     Width           =   3975
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Show Current Grs. Wt. In Bag Split"
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
                     Index           =   74
                     Left            =   150
                     TabIndex        =   355
                     Top             =   1530
                     Width           =   3975
                  End
               End
               Begin VB.Frame FraOrdDet 
                  Height          =   3810
                  Left            =   5760
                  TabIndex        =   374
                  Top             =   -120
                  Width           =   9165
                  Begin VB.Frame FraOrdEnt 
                     Height          =   1700
                     Left            =   5520
                     TabIndex        =   560
                     Top             =   360
                     Width           =   3615
                     Begin MwfCtl.MWCTL_MED ATXT 
                        Height          =   285
                        Index           =   46
                        Left            =   2145
                        TabIndex        =   123
                        ToolTipText     =   "Enter Default Order Design Group Option"
                        Top             =   765
                        Width           =   1335
                        _ExtentX        =   2355
                        _ExtentY        =   503
                        MaxLength       =   8
                        DataType        =   4
                        DataField       =   "hOrdDsgGr"
                        IdName          =   "HORDDSGGR"
                     End
                     Begin MwfCtl.MWCTL_MED ATXT 
                        Height          =   285
                        Index           =   47
                        Left            =   2145
                        TabIndex        =   124
                        ToolTipText     =   "Enter Default Order Raw Material Group Option"
                        Top             =   1050
                        Width           =   1335
                        _ExtentX        =   2355
                        _ExtentY        =   503
                        MaxLength       =   8
                        DataType        =   4
                        DataField       =   "hOrdRmGr"
                        IdName          =   "HORDRMGR"
                     End
                     Begin MwfCtl.MWCTL_MED ATXT 
                        Height          =   285
                        Index           =   48
                        Left            =   2145
                        TabIndex        =   125
                        ToolTipText     =   "Enter Default Order Labour Group Option"
                        Top             =   1335
                        Width           =   1335
                        _ExtentX        =   2355
                        _ExtentY        =   503
                        MaxLength       =   8
                        DataType        =   4
                        DataField       =   "hOrdLabGr"
                        IdName          =   "HORDLABGR"
                     End
                     Begin MwfCtl.MWCTL_MED ATXT 
                        Height          =   285
                        Index           =   51
                        Left            =   2145
                        TabIndex        =   122
                        ToolTipText     =   "Show Order Amended Qty Or Entered Qty In The Order Entry Summary Tab (Enter Y/ N for Amended/ Entered Respectively)"
                        Top             =   480
                        Width           =   285
                        _ExtentX        =   503
                        _ExtentY        =   503
                        DataType        =   4
                        DataField       =   "HOdAmdQty"
                        IdName          =   "HODAMDQTY"
                     End
                     Begin VB.Label ALBL 
                        BackStyle       =   0  'Transparent
                        Caption         =   "Show Ord Amd Qty"
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
                        Index           =   50
                        Left            =   120
                        TabIndex        =   565
                        Top             =   480
                        Width           =   2115
                     End
                     Begin VB.Label ALBL 
                        BackStyle       =   0  'Transparent
                        Caption         =   "Ord Rm Group Opt"
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
                        Index           =   42
                        Left            =   120
                        TabIndex        =   564
                        Top             =   1050
                        Width           =   2145
                     End
                     Begin VB.Label ALBL 
                        BackStyle       =   0  'Transparent
                        Caption         =   "Ord Dsg Group Opt"
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
                        Index           =   43
                        Left            =   120
                        TabIndex        =   563
                        Top             =   765
                        Width           =   2145
                     End
                     Begin VB.Label ALBL 
                        BackStyle       =   0  'Transparent
                        Caption         =   "Ord Lab Group Opt"
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
                        Index           =   45
                        Left            =   120
                        TabIndex        =   562
                        Top             =   1335
                        Width           =   2025
                     End
                     Begin VB.Label LblMain 
                        BackStyle       =   0  'Transparent
                        BorderStyle     =   1  'Fixed Single
                        Caption         =   "  Options in Order Entry Summary"
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
                        Index           =   23
                        Left            =   0
                        TabIndex        =   561
                        Top             =   120
                        Width           =   3780
                     End
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   40
                     Left            =   4095
                     TabIndex        =   111
                     ToolTipText     =   "Enter Default Order Priority Code"
                     Top             =   380
                     Width           =   435
                     _ExtentX        =   767
                     _ExtentY        =   503
                     MaxLength       =   2
                     DataType        =   4
                     DataField       =   "hOmPrtCd"
                     IdName          =   "HOMPRTCD"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   66
                     Left            =   4095
                     TabIndex        =   112
                     ToolTipText     =   "Enter Method Of Generating Suffix In Party-Dsg Combination Master"
                     Top             =   680
                     Width           =   1335
                     _ExtentX        =   2355
                     _ExtentY        =   503
                     MaxLength       =   8
                     DataType        =   4
                     DataField       =   "hPdcSfx"
                     IdName          =   "HPDCSFX"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   67
                     Left            =   4095
                     TabIndex        =   113
                     ToolTipText     =   "Enter Option By Which Rounding Off Takes Place For Sales Price"
                     Top             =   950
                     Width           =   1335
                     _ExtentX        =   2355
                     _ExtentY        =   503
                     MaxLength       =   8
                     DataType        =   4
                     DataField       =   "HRndOrdSalPrc"
                     IdName          =   "HRNDORDSALPRC"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   69
                     Left            =   4095
                     TabIndex        =   114
                     ToolTipText     =   "Enter Default Order Rm Copy Rate Option (i.e. Specify The RmCtg's For Which Rates Have To Be Copied From Source)"
                     Top             =   1235
                     Width           =   1335
                     _ExtentX        =   2355
                     _ExtentY        =   503
                     MaxLength       =   8
                     DataType        =   4
                     DataField       =   "HOrdCpyRtOpt"
                     IdName          =   "HORDCPYRTOPT"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   103
                     Left            =   4815
                     TabIndex        =   0
                     ToolTipText     =   "Enter Default Contractor Code"
                     Top             =   1520
                     Visible         =   0   'False
                     Width           =   615
                     _ExtentX        =   1085
                     _ExtentY        =   503
                     MaxLength       =   8
                     DataType        =   4
                     DataField       =   "HCtCd"
                     ReCalcOn        =   "HCTBMOD"
                     IdName          =   "HCTCD"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   126
                     Left            =   4095
                     TabIndex        =   116
                     ToolTipText     =   "Enter Gram Check Option"
                     Top             =   1805
                     Width           =   1335
                     _ExtentX        =   2355
                     _ExtentY        =   503
                     MaxLength       =   8
                     DataType        =   4
                     DataField       =   "HOdGmChk"
                     IdName          =   "HODGMCHK"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   157
                     Left            =   4095
                     TabIndex        =   117
                     ToolTipText     =   "Customise Order Entry According To User Rights (Yes/ No)"
                     Top             =   2090
                     Width           =   285
                     _ExtentX        =   503
                     _ExtentY        =   503
                     DataType        =   4
                     DataField       =   "HCustomiseOrd"
                     ReCalcParent    =   "HCUSTOMIZERMCTGEXP"
                     IdName          =   "HCUSTOMISEORD"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   216
                     Left            =   4095
                     TabIndex        =   119
                     ToolTipText     =   "Specify Whether Import From Customer Excel File to be allowed in Order Entry"
                     Top             =   2640
                     Width           =   285
                     _ExtentX        =   503
                     _ExtentY        =   503
                     DataType        =   4
                     DataField       =   "HExcelImpInOrdYN"
                     IdName          =   "HEXCELIMPINORDYN"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   218
                     Left            =   4095
                     TabIndex        =   120
                     ToolTipText     =   "Specify Whether Import From Customer Excel File to be allowed in Order Entry"
                     Top             =   2945
                     Width           =   285
                     _ExtentX        =   503
                     _ExtentY        =   503
                     DataType        =   4
                     DataField       =   "hTravelExlYN"
                     IdName          =   "HTRAVELEXLYN"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   221
                     Left            =   4095
                     TabIndex        =   118
                     ToolTipText     =   "Specify exception RmCtgs to the IDE Rule "
                     Top             =   2375
                     Width           =   1110
                     _ExtentX        =   1958
                     _ExtentY        =   503
                     MaxLength       =   5
                     DataType        =   4
                     DataField       =   "hCustomizeRmCtgExp"
                     ReCalcOn        =   "HCUSTOMISEORD"
                     IdName          =   "HCUSTOMIZERMCTGEXP"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   232
                     Left            =   8715
                     TabIndex        =   129
                     ToolTipText     =   "Specify Whether Order Printing Report to be taken in Excel Format"
                     Top             =   2945
                     Width           =   285
                     _ExtentX        =   503
                     _ExtentY        =   503
                     DataType        =   4
                     DataField       =   "HOrdExcelRepYn"
                     IdName          =   "HORDEXCELREPYN"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   247
                     Left            =   8715
                     TabIndex        =   127
                     ToolTipText     =   "Specify RfId Module"
                     Top             =   2375
                     Width           =   285
                     _ExtentX        =   503
                     _ExtentY        =   503
                     Enabled         =   0   'False
                     DataType        =   4
                     DataField       =   "HRfIdMod"
                     IdName          =   "HRFIDMOD"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   248
                     Left            =   8040
                     TabIndex        =   128
                     ToolTipText     =   "Enter RmCtg Combination For Rate Round Off in Order"
                     Top             =   2660
                     Width           =   960
                     _ExtentX        =   1693
                     _ExtentY        =   503
                     MaxLength       =   5
                     DataType        =   4
                     DataField       =   "HOrdRtRndOffRmCtg"
                     IdName          =   "HORDRTRNDOFFRMCTG"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   252
                     Left            =   8715
                     TabIndex        =   126
                     ToolTipText     =   "Specify Multi Price Quotation Module"
                     Top             =   2090
                     Width           =   285
                     _ExtentX        =   503
                     _ExtentY        =   503
                     Enabled         =   0   'False
                     DataType        =   4
                     DataField       =   "HMultiPrcQtMod"
                     IdName          =   "HMULTIPRCQTMOD"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   293
                     Left            =   4095
                     TabIndex        =   121
                     ToolTipText     =   "Check zero in Order Rm Rate Y/N?"
                     Top             =   3210
                     Width           =   285
                     _ExtentX        =   503
                     _ExtentY        =   503
                     DataType        =   4
                     DataField       =   "hChkOrRtYn"
                     IdName          =   "HCHKORRTYN"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   296
                     Left            =   8715
                     TabIndex        =   130
                     ToolTipText     =   "Specify Whether CHC and DHC charges to be applied on Customer Rm "
                     Top             =   3210
                     Width           =   285
                     _ExtentX        =   503
                     _ExtentY        =   503
                     DataType        =   4
                     DataField       =   "HCHCDHCOnCustRm"
                     IdName          =   "HCHCDHCONCUSTRM"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   303
                     Left            =   4095
                     TabIndex        =   115
                     ToolTipText     =   "When Sales rate coming from PDCM : ‘YES’ Cost from PDCM, “NO” – Cost from Current Rate charts"
                     Top             =   1520
                     Width           =   285
                     _ExtentX        =   503
                     _ExtentY        =   503
                     DataType        =   4
                     DataField       =   "hCostFrPdcmYN"
                     ReCalcParent    =   "HCUSTOMIZERMCTGEXP"
                     IdName          =   "HCOSTFRPDCMYN"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   306
                     Left            =   8715
                     TabIndex        =   131
                     ToolTipText     =   "Specify Implement OrdRm Lot YN"
                     Top             =   3480
                     Width           =   285
                     _ExtentX        =   503
                     _ExtentY        =   503
                     DataType        =   4
                     DataField       =   "HOrLotYN"
                     IdName          =   "HORLOTYN"
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Implement OrdRm Lot "
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   255
                     Index           =   239
                     Left            =   4620
                     TabIndex        =   640
                     Top             =   3480
                     Width           =   3675
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Cost from PDCM YN"
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
                     Index           =   236
                     Left            =   150
                     TabIndex        =   637
                     Top             =   1560
                     Width           =   1875
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "CHC and DHC On Customer Rm"
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   255
                     Index           =   228
                     Left            =   4620
                     TabIndex        =   627
                     Top             =   3210
                     Width           =   3675
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Chk Zero In Ord Rm Rate"
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   255
                     Index           =   216
                     Left            =   120
                     TabIndex        =   625
                     Top             =   3210
                     Width           =   3795
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Multi Price Quotation Module"
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   255
                     Index           =   187
                     Left            =   5580
                     TabIndex        =   549
                     Top             =   2090
                     Width           =   2835
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Round Off Rate For Ord RmCtg "
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   255
                     Index           =   183
                     Left            =   4620
                     TabIndex        =   545
                     Top             =   2660
                     Width           =   3315
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "RfId Module "
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   255
                     Index           =   182
                     Left            =   7230
                     TabIndex        =   544
                     Top             =   2375
                     Width           =   1185
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Enable Order Printing in Excel Format"
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   255
                     Index           =   169
                     Left            =   4620
                     TabIndex        =   527
                     Top             =   2945
                     Width           =   3675
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "RmCtg Exception to Customization"
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   255
                     Index           =   161
                     Left            =   150
                     TabIndex        =   511
                     Top             =   2375
                     Width           =   3795
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Enable Travel Excel"
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   255
                     Index           =   158
                     Left            =   150
                     TabIndex        =   508
                     Top             =   2945
                     Width           =   3795
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Allow Order Import From Excel"
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   255
                     Index           =   156
                     Left            =   150
                     TabIndex        =   505
                     Top             =   2640
                     Width           =   3795
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Customise Order Entry Screen"
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   255
                     Index           =   132
                     Left            =   150
                     TabIndex        =   476
                     Top             =   2090
                     Width           =   3795
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Gram Check Option"
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
                     Index           =   112
                     Left            =   150
                     TabIndex        =   441
                     Top             =   1805
                     Width           =   3585
                  End
                  Begin VB.Label LblMain 
                     BackStyle       =   0  'Transparent
                     BorderStyle     =   1  'Fixed Single
                     Caption         =   "  Order Details"
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
                     Index           =   7
                     Left            =   -30
                     TabIndex        =   430
                     Top             =   90
                     Width           =   9165
                  End
                  Begin VB.Label LblCtCd 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Contractor Code"
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
                     Left            =   2250
                     TabIndex        =   401
                     Top             =   1520
                     Visible         =   0   'False
                     Width           =   1635
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Ord Priority Cd"
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
                     Index           =   37
                     Left            =   150
                     TabIndex        =   378
                     Top             =   380
                     Width           =   2505
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Method Of Generating Suffix For PDC"
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
                     Index           =   64
                     Left            =   150
                     TabIndex        =   377
                     Top             =   680
                     Width           =   3945
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Round Off Option For Sales Price"
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
                     Index           =   65
                     Left            =   150
                     TabIndex        =   376
                     Top             =   950
                     Width           =   3945
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Rm Rate Copy Opt In Ord Entry "
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
                     Index           =   67
                     Left            =   150
                     TabIndex        =   375
                     Top             =   1235
                     Width           =   3945
                  End
               End
            End
            Begin VB.Frame FraTabHd 
               BorderStyle     =   0  'None
               Height          =   8220
               Index           =   4
               Left            =   60
               TabIndex        =   333
               Top             =   360
               Width           =   14950
               Begin VB.Frame FraOth 
                  Height          =   8415
                  Left            =   0
                  TabIndex        =   340
                  Top             =   -120
                  Width           =   5595
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   44
                     Left            =   3945
                     TabIndex        =   154
                     ToolTipText     =   "Enter Default Year"
                     Top             =   330
                     Width           =   435
                     _ExtentX        =   767
                     _ExtentY        =   503
                     MaxLength       =   2
                     DataType        =   4
                     DataField       =   "hYy"
                     IdName          =   "HYY"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   65
                     Left            =   3945
                     TabIndex        =   155
                     ToolTipText     =   "Specify Whether To Allow The User To Change The 'Ptr Y/N' Option In The Param Of Type 'RMSCTG'"
                     Top             =   615
                     Width           =   285
                     _ExtentX        =   503
                     _ExtentY        =   503
                     DataType        =   4
                     DataField       =   "HPtrYN"
                     IdName          =   "HPTRYN"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   71
                     Left            =   3945
                     TabIndex        =   156
                     ToolTipText     =   "Should The Production Rm Qty && Wt Be Shown In Dsg And Ord Entries"
                     Top             =   900
                     Width           =   285
                     _ExtentX        =   503
                     _ExtentY        =   503
                     DataType        =   4
                     DataField       =   "HShowPrdRmDets"
                     IdName          =   "HSHOWPRDRMDETS"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   99
                     Left            =   3945
                     TabIndex        =   163
                     ToolTipText     =   "Specify Whether To implement Average Stock Rate (Y/N)"
                     Top             =   2610
                     Width           =   285
                     _ExtentX        =   503
                     _ExtentY        =   503
                     DataType        =   4
                     DataField       =   "HAVGSRTYN"
                     ReCalcParent    =   "HSTKRTNOEDIT,HSTKRTNODSP,HSTKRTNOEDIT"
                     IdName          =   "HAVGSRTYN"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   100
                     Left            =   3945
                     TabIndex        =   157
                     ToolTipText     =   "Enter 'Y'  for FgRm and 'N' for OrdRm to be copied in InvRm"
                     Top             =   1185
                     Width           =   285
                     _ExtentX        =   503
                     _ExtentY        =   503
                     DataType        =   4
                     DataField       =   "HSIZESTKRT"
                     IdName          =   "HSIZESTKRT"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   102
                     Left            =   3945
                     TabIndex        =   158
                     ToolTipText     =   "Specify Whether The Contractor Module Is To Be Shown In The System"
                     Top             =   1470
                     Width           =   285
                     _ExtentX        =   503
                     _ExtentY        =   503
                     DataType        =   4
                     DataField       =   "HCtbMod"
                     ReCalcParent    =   "HCTCD"
                     IdName          =   "HCTBMOD"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   98
                     Left            =   3945
                     TabIndex        =   159
                     ToolTipText     =   "Should The Priority Allotment Be On Production Or On Export Date Or Should There Be No Priority Allotment"
                     Top             =   1755
                     Width           =   285
                     _ExtentX        =   503
                     _ExtentY        =   503
                     DataType        =   4
                     DataField       =   "HPrtAllotOnPrdYN"
                     IdName          =   "HPRTALLOTONPRDYN"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   104
                     Left            =   3945
                     TabIndex        =   167
                     ToolTipText     =   " Check Diamond/Color Stone Qty at Issue,Enter (Y/N)?"
                     Top             =   3465
                     Width           =   285
                     _ExtentX        =   503
                     _ExtentY        =   503
                     DataType        =   4
                     DataField       =   "HChkDcQtyAtIssYN"
                     IdName          =   "HCHKDCQTYATISSYN"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   108
                     Left            =   3945
                     TabIndex        =   169
                     ToolTipText     =   " Check Gold/Platinum  Wt  at Issue,Enter (Y/N)?"
                     Top             =   3750
                     Width           =   285
                     _ExtentX        =   503
                     _ExtentY        =   503
                     DataType        =   4
                     DataField       =   "HChkGPWtAtIssYN"
                     IdName          =   "HCHKGPWTATISSYN"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   97
                     Left            =   3945
                     TabIndex        =   161
                     ToolTipText     =   "Specify Whether To Allow The Negative Stocks. "
                     Top             =   2325
                     Width           =   285
                     _ExtentX        =   503
                     _ExtentY        =   503
                     DataType        =   4
                     DataField       =   "HNegStkYN"
                     IdName          =   "HNEGSTKYN"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   107
                     Left            =   5050
                     TabIndex        =   168
                     ToolTipText     =   " Check Diamond/Color Stone Wt at Issue,Enter (Y/N)?"
                     Top             =   3465
                     Width           =   285
                     _ExtentX        =   503
                     _ExtentY        =   503
                     DataType        =   4
                     DataField       =   "HChkDCWtAtIssYN"
                     IdName          =   "HCHKDCWTATISSYN"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   115
                     Left            =   3945
                     TabIndex        =   160
                     ToolTipText     =   " Check Fg  Raw Materials With Order RawMaterials ,Enter (Y/N)?"
                     Top             =   2040
                     Width           =   285
                     _ExtentX        =   503
                     _ExtentY        =   503
                     DataType        =   4
                     DataField       =   "hFgRmEqOrdRm"
                     IdName          =   "HFGRMEQORDRM"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   119
                     Left            =   3945
                     TabIndex        =   170
                     ToolTipText     =   " Specify Whether SEO Password Is Requre To Edit Stock Rate"
                     Top             =   4035
                     Width           =   285
                     _ExtentX        =   503
                     _ExtentY        =   503
                     DataType        =   4
                     DataField       =   "HStkRtPassReq"
                     IdName          =   "HSTKRTPASSREQ"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   123
                     Left            =   3945
                     TabIndex        =   171
                     ToolTipText     =   "Specify Whether To Issue Rm By Production Pointer"
                     Top             =   4320
                     Width           =   285
                     _ExtentX        =   503
                     _ExtentY        =   503
                     DataType        =   4
                     DataField       =   "HDTPrdPtr"
                     IdName          =   "HDTPRDPTR"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   121
                     Left            =   3945
                     TabIndex        =   172
                     ToolTipText     =   "Specify Permissible Alloy Wt. Difference % In Conversion Entry"
                     Top             =   4605
                     Width           =   795
                     _ExtentX        =   1402
                     _ExtentY        =   503
                     Mask            =   "##0.00"
                     MaxLength       =   6
                     DataType        =   2
                     DataField       =   "HAlyCnvDiff"
                     IdName          =   "HALYCNVDIFF"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   124
                     Left            =   3945
                     TabIndex        =   173
                     ToolTipText     =   " Effective Date For Permissible Weight Difference"
                     Top             =   4890
                     Width           =   885
                     _ExtentX        =   1561
                     _ExtentY        =   503
                     MaxLength       =   8
                     DataType        =   3
                     DataField       =   "HCNVDIFFEFFDT"
                     IdName          =   "HCNVDIFFEFFDT"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   127
                     Left            =   3945
                     TabIndex        =   166
                     ToolTipText     =   "Allow Lot No. For Dia/Cs (Y/N)"
                     Top             =   3180
                     Width           =   285
                     _ExtentX        =   503
                     _ExtentY        =   503
                     DataType        =   4
                     DataField       =   "HLotNoYN"
                     IdName          =   "HLOTNOYN"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   129
                     Left            =   3945
                     TabIndex        =   174
                     ToolTipText     =   "Implement Silver Module,Enter (Y/N)?"
                     Top             =   5175
                     Width           =   285
                     _ExtentX        =   503
                     _ExtentY        =   503
                     DataType        =   4
                     DataField       =   "HSLVMODYN"
                     IdName          =   "HSLVMODYN"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   131
                     Left            =   3945
                     TabIndex        =   175
                     ToolTipText     =   "Show Quotation For FG Bag, Enter (Y/N)."
                     Top             =   5460
                     Width           =   285
                     _ExtentX        =   503
                     _ExtentY        =   503
                     DataType        =   4
                     DataField       =   "HFgQuotYN"
                     IdName          =   "HFGQUOTYN"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   132
                     Left            =   3945
                     TabIndex        =   176
                     ToolTipText     =   "Show Derived Or Actual Gold Rate In Order/Invoice Printing, Enter Y For Derived Or N For Actual."
                     Top             =   5745
                     Width           =   285
                     _ExtentX        =   503
                     _ExtentY        =   503
                     DataType        =   4
                     DataField       =   "HGldRtDAYN"
                     IdName          =   "HGLDRTDAYN"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   137
                     Left            =   3945
                     TabIndex        =   177
                     ToolTipText     =   "QW Module,Enter (Y/N)?"
                     Top             =   6030
                     Width           =   285
                     _ExtentX        =   503
                     _ExtentY        =   503
                     DataType        =   4
                     DataField       =   "HQWMODYN"
                     IdName          =   "HQWMODYN"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   138
                     Left            =   3945
                     TabIndex        =   179
                     ToolTipText     =   "Enter Base Currency Code"
                     Top             =   6600
                     Width           =   885
                     _ExtentX        =   1561
                     _ExtentY        =   503
                     MaxLength       =   5
                     DataType        =   4
                     DataField       =   "HBaseCurCd"
                     IdName          =   "HBASECURCD"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   139
                     Left            =   3945
                     TabIndex        =   180
                     ToolTipText     =   "Show Model Availability (Y/N)"
                     Top             =   6885
                     Width           =   285
                     _ExtentX        =   503
                     _ExtentY        =   503
                     DataType        =   4
                     DataField       =   "HModAvlYN"
                     IdName          =   "HMODAVLYN"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   140
                     Left            =   3945
                     TabIndex        =   178
                     ToolTipText     =   "Enter 'Y' To Implement Multi Currency Module Or 'N' To Disable It"
                     Top             =   6315
                     Width           =   285
                     _ExtentX        =   503
                     _ExtentY        =   503
                     DataType        =   4
                     DataField       =   "hMultiCurYN"
                     IdName          =   "HMULTICURYN"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   145
                     Left            =   4260
                     TabIndex        =   162
                     ToolTipText     =   "Specify the RM Categories to be checked For Negative Stock or Blank for All"
                     Top             =   2325
                     Width           =   1155
                     _ExtentX        =   2037
                     _ExtentY        =   503
                     MaxLength       =   5
                     DataType        =   4
                     DataField       =   "hNegStkRmCtg"
                     IdName          =   "HNEGSTKRMCTG"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   146
                     Left            =   3945
                     TabIndex        =   181
                     ToolTipText     =   "Implement Common Location in Multi-Company Environment (Y/N)"
                     Top             =   7170
                     Width           =   285
                     _ExtentX        =   503
                     _ExtentY        =   503
                     DataType        =   4
                     DataField       =   "HImpComLoc"
                     IdName          =   "HIMPCOMLOC"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   199
                     Left            =   3945
                     TabIndex        =   182
                     ToolTipText     =   "Change Pointer On STW (Y/N)"
                     Top             =   7455
                     Width           =   285
                     _ExtentX        =   503
                     _ExtentY        =   503
                     DataType        =   4
                     DataField       =   "hChgPtrOnStwYN"
                     IdName          =   "HCHGPTRONSTWYN"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   206
                     Left            =   3945
                     TabIndex        =   183
                     ToolTipText     =   "Implement Palladium Module,Enter (Y/N)?"
                     Top             =   7740
                     Width           =   285
                     _ExtentX        =   503
                     _ExtentY        =   503
                     DataType        =   4
                     DataField       =   "HPdModYN"
                     IdName          =   "HPDMODYN"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   235
                     Left            =   3945
                     TabIndex        =   184
                     ToolTipText     =   "Show Company Logo In Reports"
                     Top             =   8040
                     Width           =   285
                     _ExtentX        =   503
                     _ExtentY        =   503
                     DataType        =   4
                     DataField       =   "HShowRepLogo"
                     IdName          =   "HSHOWREPLOGO"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   301
                     Left            =   3945
                     TabIndex        =   164
                     ToolTipText     =   "Enter No Display Yes\No"
                     Top             =   2895
                     Width           =   285
                     _ExtentX        =   503
                     _ExtentY        =   503
                     DataType        =   4
                     DataField       =   "HSTKRTNODSP"
                     ReCalcParent    =   "HSTKRTNOEDIT"
                     ReCalcOn        =   "HAVGSRTYN"
                     IdName          =   "HSTKRTNODSP"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   302
                     Left            =   5050
                     TabIndex        =   165
                     ToolTipText     =   "Enter No Display Yes\No"
                     Top             =   2895
                     Width           =   285
                     _ExtentX        =   503
                     _ExtentY        =   503
                     DataType        =   4
                     DataField       =   "HSTKRTNOEDIT"
                     ReCalcOn        =   "HAVGSRTYN,HSTKRTNODSP"
                     IdName          =   "HSTKRTNOEDIT"
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Stock Rate NoDisplay"
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
                     Index           =   235
                     Left            =   120
                     TabIndex        =   636
                     Top             =   2895
                     Width           =   3675
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "NoEdit"
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
                     Index           =   234
                     Left            =   4320
                     TabIndex        =   635
                     Top             =   2895
                     Width           =   675
                  End
                  Begin VB.Label Label21 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Show Company Logo in Report"
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
                     Left            =   120
                     TabIndex        =   531
                     Top             =   8040
                     Width           =   3435
                  End
                  Begin VB.Label Label20 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Implement Palladium Module"
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
                     Left            =   120
                     TabIndex        =   502
                     Top             =   7740
                     Width           =   3435
                  End
                  Begin VB.Label Label19 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Change Pointer On STW"
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
                     Left            =   120
                     TabIndex        =   499
                     Top             =   7455
                     Width           =   3435
                  End
                  Begin VB.Label Label17 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Implement Common Locations"
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
                     Left            =   120
                     TabIndex        =   463
                     Top             =   7170
                     Width           =   3435
                  End
                  Begin VB.Label Label16 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Implement Multi Currency Module"
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
                     Left            =   120
                     TabIndex        =   461
                     Top             =   6315
                     Width           =   3435
                  End
                  Begin VB.Label Label15 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Show Model Availability"
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
                     Left            =   120
                     TabIndex        =   459
                     Top             =   6885
                     Width           =   3435
                  End
                  Begin VB.Label Label14 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Base Currency Code"
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
                     Left            =   120
                     TabIndex        =   458
                     Top             =   6600
                     Width           =   3435
                  End
                  Begin VB.Label Label13 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Q/W For All RM Categories"
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
                     Left            =   120
                     TabIndex        =   457
                     Top             =   6030
                     Width           =   3435
                  End
                  Begin VB.Label Label12 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Derived/Actual Gld Rt In Ord/Inv Prnt"
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
                     Left            =   120
                     TabIndex        =   446
                     Top             =   5745
                     Width           =   3915
                  End
                  Begin VB.Label Label11 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Show Quotation For FG Bag"
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
                     Left            =   120
                     TabIndex        =   445
                     Top             =   5460
                     Width           =   3435
                  End
                  Begin VB.Label Label10 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Implement Silver Module"
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
                     Left            =   120
                     TabIndex        =   443
                     Top             =   5175
                     Width           =   3435
                  End
                  Begin VB.Label Label9 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "Implement Lot No.For D/C/X/G/P/S/L/M"
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
                     Left            =   120
                     TabIndex        =   442
                     Top             =   3180
                     Width           =   3660
                  End
                  Begin VB.Label LblMain 
                     BackStyle       =   0  'Transparent
                     BorderStyle     =   1  'Fixed Single
                     Caption         =   "  General And Param Options"
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
                     Index           =   8
                     Left            =   120
                     TabIndex        =   431
                     Top             =   60
                     Width           =   5625
                  End
                  Begin VB.Label Label8 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "Effective Dt For Permissible Wt. Diff"
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
                     Left            =   120
                     TabIndex        =   421
                     Top             =   4890
                     Width           =   3360
                  End
                  Begin VB.Label Label7 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "Permissible Wt. Difference % In Cnv"
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
                     Left            =   120
                     TabIndex        =   420
                     Top             =   4605
                     Width           =   3390
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "In DT Issue By Prd Ptr"
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
                     Index           =   110
                     Left            =   120
                     TabIndex        =   419
                     Top             =   4320
                     Width           =   3825
                  End
                  Begin VB.Label Label6 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "Stk Rt Require SEO Pass to Edit"
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
                     Left            =   120
                     TabIndex        =   417
                     Top             =   4035
                     Width           =   3345
                  End
                  Begin VB.Label Label5 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Check  Fg RmCd With Order RM"
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
                     Left            =   120
                     TabIndex        =   412
                     Top             =   2040
                     Width           =   3735
                  End
                  Begin VB.Label Label3 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Chk Wt"
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
                     Left            =   4320
                     TabIndex        =   408
                     Top             =   3465
                     Width           =   675
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Allow Negative Stocks / Rm Ctg."
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
                     Index           =   91
                     Left            =   120
                     TabIndex        =   407
                     Top             =   2325
                     Width           =   3825
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Implement Average Stock Rate"
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
                     Index           =   93
                     Left            =   120
                     TabIndex        =   406
                     Top             =   2610
                     Width           =   3825
                  End
                  Begin VB.Label Label2 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Check Dia/Cs At Issue - Chk Qty"
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
                     Left            =   120
                     TabIndex        =   405
                     Top             =   3465
                     Width           =   3675
                  End
                  Begin VB.Label Label1 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Priority Allotment On Production"
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
                     Left            =   120
                     TabIndex        =   404
                     Top             =   1755
                     Width           =   3435
                  End
                  Begin VB.Label Label4 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Check  Gold/Platinum  Wt At Issue"
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
                     Left            =   120
                     TabIndex        =   403
                     Top             =   3750
                     Width           =   3735
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Show Contractor Module"
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
                     Index           =   96
                     Left            =   120
                     TabIndex        =   400
                     Top             =   1470
                     Width           =   3825
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Bring RmDetails in Inv From FgRm"
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
                     Index           =   94
                     Left            =   120
                     TabIndex        =   398
                     Top             =   1185
                     Width           =   3825
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Default Year"
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
                     Index           =   41
                     Left            =   120
                     TabIndex        =   343
                     Top             =   330
                     Width           =   2055
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Allow Ptr Option At RmSub Ctg Level"
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
                     Index           =   63
                     Left            =   120
                     TabIndex        =   342
                     Top             =   615
                     Width           =   3825
                  End
                  Begin VB.Label lblShowBothWts 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Show Both Wghts (Prd && Sal) YN"
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
                     Left            =   120
                     TabIndex        =   341
                     Top             =   900
                     Width           =   3405
                  End
               End
               Begin VB.Frame FraInvRepDet 
                  Height          =   3555
                  Left            =   5520
                  TabIndex        =   334
                  Top             =   5130
                  Width           =   9495
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   61
                     Left            =   4065
                     TabIndex        =   212
                     ToolTipText     =   "Specify Whether To Show Invoice Or Customs Wt & Amt In Inv Printing Rep"
                     Top             =   420
                     Width           =   285
                     _ExtentX        =   503
                     _ExtentY        =   503
                     DataType        =   4
                     DataField       =   "hInRepWtValAs"
                     IdName          =   "HINREPWTVALAS"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   60
                     Left            =   4065
                     TabIndex        =   213
                     ToolTipText     =   "Specify Whether To Show Wt Column In Inv Printing Rep"
                     Top             =   705
                     Width           =   285
                     _ExtentX        =   503
                     _ExtentY        =   503
                     DataType        =   4
                     DataField       =   "hInRepWtFld"
                     IdName          =   "HINREPWTFLD"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   62
                     Left            =   4065
                     TabIndex        =   214
                     ToolTipText     =   "Specify Whether To Show Metal Loss Summary In Inv Printing Rep"
                     Top             =   990
                     Width           =   285
                     _ExtentX        =   503
                     _ExtentY        =   503
                     DataType        =   4
                     DataField       =   "hInRepMetLs"
                     IdName          =   "HINREPMETLS"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   63
                     Left            =   4065
                     TabIndex        =   216
                     ToolTipText     =   "Specify Whether To Show Conversion Rt & Rupee Equivalent In Inv Printing Rep"
                     Top             =   1575
                     Width           =   285
                     _ExtentX        =   503
                     _ExtentY        =   503
                     DataType        =   4
                     DataField       =   "hInRepCnvRt"
                     IdName          =   "HINREPCNVRT"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   64
                     Left            =   4065
                     TabIndex        =   217
                     ToolTipText     =   "Specify Whether To Show Rt Column In Inv Printing Rep"
                     Top             =   1860
                     Width           =   285
                     _ExtentX        =   503
                     _ExtentY        =   503
                     DataType        =   4
                     DataField       =   "hInRepRtFld"
                     IdName          =   "HINREPRTFLD"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   93
                     Left            =   4065
                     TabIndex        =   218
                     ToolTipText     =   "Specify Whether To Show Value Addition On Metal Column In Value Addition Report ? (Y/N)"
                     Top             =   2145
                     Width           =   285
                     _ExtentX        =   503
                     _ExtentY        =   503
                     DataType        =   4
                     DataField       =   "hValAddMtlColYN"
                     IdName          =   "HVALADDMTLCOLYN"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   94
                     Left            =   4065
                     TabIndex        =   219
                     ToolTipText     =   "Specify Whether To Show BOE Details in Value Addition Annexure Report? (Y/N)"
                     Top             =   2430
                     Width           =   285
                     _ExtentX        =   503
                     _ExtentY        =   503
                     DataType        =   4
                     DataField       =   "hValBOEDetYN"
                     IdName          =   "HVALBOEDETYN"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   134
                     Left            =   4065
                     TabIndex        =   220
                     ToolTipText     =   "Specify Whether To Show Menu Options Related To Loose RM Invoicing Reports? (Y/N)"
                     Top             =   2715
                     Width           =   285
                     _ExtentX        =   503
                     _ExtentY        =   503
                     DataType        =   4
                     DataField       =   "HLooseInv"
                     IdName          =   "HLOOSEINV"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   136
                     Left            =   8985
                     TabIndex        =   221
                     ToolTipText     =   "Specify Whether To Show Menu Options Related To Consolidated Invoicing Reports? (Y/N)"
                     Top             =   420
                     Width           =   285
                     _ExtentX        =   503
                     _ExtentY        =   503
                     DataType        =   4
                     DataField       =   "HConsolidatedInv"
                     IdName          =   "HCONSOLIDATEDINV"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   147
                     Left            =   4065
                     TabIndex        =   215
                     ToolTipText     =   "Specify Whether To Show Metal Loss Summary based on Rm Code for Non-Local Invoices"
                     Top             =   1290
                     Width           =   285
                     _ExtentX        =   503
                     _ExtentY        =   503
                     DataType        =   4
                     DataField       =   "HINREPMETLSRMCD"
                     IdName          =   "HINREPMETLSRMCD"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   295
                     Left            =   8985
                     TabIndex        =   224
                     ToolTipText     =   "Specify Max Number of Sizes to keep Rm Stock By"
                     Top             =   1800
                     Width           =   285
                     _ExtentX        =   503
                     _ExtentY        =   503
                     Mask            =   "#"
                     DataType        =   4
                     DataField       =   "HRMStkBy"
                     IdName          =   "HRMSTKBY"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   300
                     Left            =   8685
                     TabIndex        =   225
                     ToolTipText     =   "Enter Cluster Cd"
                     Top             =   2700
                     Width           =   555
                     _ExtentX        =   979
                     _ExtentY        =   503
                     MaxLength       =   2
                     DataType        =   4
                     DataField       =   "HClstrCd"
                     IdName          =   "HCLSTRCD"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   304
                     Left            =   8985
                     TabIndex        =   222
                     ToolTipText     =   "Enter YN"
                     Top             =   705
                     Width           =   285
                     _ExtentX        =   503
                     _ExtentY        =   503
                     DataType        =   4
                     DataField       =   "HOLORYN"
                     IdName          =   "HOLORYN"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   305
                     Left            =   8985
                     TabIndex        =   223
                     ToolTipText     =   "Enter YN"
                     Top             =   990
                     Width           =   285
                     _ExtentX        =   503
                     _ExtentY        =   503
                     DataType        =   4
                     DataField       =   "hAddBrkMsgYN"
                     IdName          =   "HADDBRKMSGYN"
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Add Broken Missing Wt to Customs Wt"
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
                     Index           =   238
                     Left            =   4920
                     TabIndex        =   639
                     Top             =   990
                     Width           =   3795
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "One Lot One Rate"
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
                     Index           =   237
                     Left            =   4920
                     TabIndex        =   638
                     Top             =   705
                     Width           =   3795
                  End
                  Begin VB.Label LblMain 
                     BackStyle       =   0  'Transparent
                     BorderStyle     =   1  'Fixed Single
                     Caption         =   "Ppc Options"
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
                     Index           =   25
                     Left            =   4860
                     TabIndex        =   634
                     Top             =   2310
                     Width           =   4395
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Cluster Cd"
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
                     Index           =   233
                     Left            =   4890
                     TabIndex        =   633
                     Top             =   2700
                     Width           =   3105
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Rm Stock By - No of Sizes"
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
                     Index           =   229
                     Left            =   4920
                     TabIndex        =   629
                     Top             =   1800
                     Width           =   3795
                  End
                  Begin VB.Label LblMain 
                     BackStyle       =   0  'Transparent
                     BorderStyle     =   1  'Fixed Single
                     Caption         =   "  Invoice Report Options"
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
                     Index           =   24
                     Left            =   0
                     TabIndex        =   628
                     Top             =   0
                     Width           =   9555
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "In Inv Prn Rep Met Loss Summ On RmCd"
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
                     Index           =   122
                     Left            =   120
                     TabIndex        =   464
                     Top             =   1290
                     Width           =   3975
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Show Consolidated Invoice Reports"
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
                     Index           =   118
                     Left            =   4920
                     TabIndex        =   456
                     Top             =   420
                     Width           =   3795
                  End
                  Begin VB.Label LblMain 
                     BackStyle       =   0  'Transparent
                     BorderStyle     =   1  'Fixed Single
                     Caption         =   "Rm Stock Options"
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
                     Index           =   9
                     Left            =   4890
                     TabIndex        =   432
                     Top             =   1410
                     Width           =   4395
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Show Loose Invoice Reports"
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
                     Index           =   116
                     Left            =   120
                     TabIndex        =   454
                     Top             =   2715
                     Width           =   3795
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Show BOE Details Y/N"
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
                     Index           =   88
                     Left            =   120
                     TabIndex        =   395
                     Top             =   2430
                     Width           =   3795
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Value Addition On Metal Column Y/N"
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
                     Index           =   87
                     Left            =   120
                     TabIndex        =   394
                     Top             =   2145
                     Width           =   3795
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "In Inv Prn Rep Show Wt/ Val As"
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
                     Index           =   59
                     Left            =   120
                     TabIndex        =   339
                     Top             =   420
                     Width           =   3735
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "In Inv Prn Rep Show Wt Column"
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
                     Index           =   58
                     Left            =   120
                     TabIndex        =   338
                     Top             =   705
                     Width           =   3735
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "In Inv Prn Rep Show Met Loss Summ"
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
                     Index           =   60
                     Left            =   120
                     TabIndex        =   337
                     Top             =   990
                     Width           =   3795
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "In Inv Prn Rep Show Cnv Rt"
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
                     Index           =   61
                     Left            =   120
                     TabIndex        =   336
                     Top             =   1560
                     Width           =   3795
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "In Inv Prn Rep Show Rt Column"
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
                     Index           =   62
                     Left            =   120
                     TabIndex        =   335
                     Top             =   1860
                     Width           =   3795
                  End
               End
               Begin VB.Frame FraInvEntryDet 
                  Height          =   4035
                  Left            =   5520
                  TabIndex        =   346
                  Top             =   1260
                  Width           =   9495
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   45
                     Left            =   4065
                     TabIndex        =   190
                     ToolTipText     =   "Enter Default Customs Weight Equal To Option"
                     Top             =   705
                     Width           =   285
                     _ExtentX        =   503
                     _ExtentY        =   503
                     DataType        =   4
                     DataField       =   "hZWtEq"
                     IdName          =   "HZWTEQ"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   53
                     Left            =   4065
                     TabIndex        =   191
                     ToolTipText     =   "Specify Whether By Default The Customs Value Should Be Equal To Invoice Value Or Not"
                     Top             =   990
                     Width           =   285
                     _ExtentX        =   503
                     _ExtentY        =   503
                     DataType        =   4
                     DataField       =   "hZValEqInv"
                     IdName          =   "HZVALEQINV"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   43
                     Left            =   4065
                     TabIndex        =   192
                     ToolTipText     =   "Specify Whether To Round Off The Inv & Custom FOB Values At The Design Level In Invoicing"
                     Top             =   1290
                     Width           =   285
                     _ExtentX        =   503
                     _ExtentY        =   503
                     DataType        =   4
                     DataField       =   "HRndOffFOB"
                     IdName          =   "HRNDOFFFOB"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   58
                     Left            =   4065
                     TabIndex        =   193
                     ToolTipText     =   "Specify Whether To Show The Actual Weight Field In The Rm Grid In Inv Dsg Entry"
                     Top             =   1575
                     Width           =   285
                     _ExtentX        =   503
                     _ExtentY        =   503
                     DataType        =   4
                     DataField       =   "HShowActWt"
                     IdName          =   "HSHOWACTWT"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   59
                     Left            =   4065
                     TabIndex        =   194
                     ToolTipText     =   "Specify Sort Option For Custom Purchase Vch Help In Inv Custm Match Entry"
                     Top             =   1860
                     Width           =   1275
                     _ExtentX        =   2249
                     _ExtentY        =   503
                     MaxLength       =   8
                     DataType        =   4
                     DataField       =   "HSrtBasePurVch"
                     IdName          =   "HSRTBASEPURVCH"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   75
                     Left            =   4065
                     TabIndex        =   195
                     ToolTipText     =   "Specify Whether The Export Is From SEEPZ Or Outside"
                     Top             =   2145
                     Width           =   285
                     _ExtentX        =   503
                     _ExtentY        =   503
                     DataType        =   4
                     DataField       =   "hExpLocal"
                     IdName          =   "HEXPLOCAL"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   92
                     Left            =   4065
                     TabIndex        =   196
                     ToolTipText     =   "Enter Default Invoice Type"
                     Top             =   2430
                     Width           =   945
                     _ExtentX        =   1667
                     _ExtentY        =   503
                     MaxLength       =   5
                     DataType        =   4
                     DataField       =   "HInvTyp"
                     IdName          =   "HINVTYP"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   135
                     Left            =   4065
                     TabIndex        =   189
                     ToolTipText     =   "Enter Default Customs Weight Equal To Option"
                     Top             =   420
                     Width           =   285
                     _ExtentX        =   503
                     _ExtentY        =   503
                     DataType        =   4
                     DataField       =   "hGenExpNoYN"
                     IdName          =   "HGENEXPNOYN"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   143
                     Left            =   4065
                     TabIndex        =   197
                     ToolTipText     =   "Specify Whether To Automate Invoice Returns"
                     Top             =   2715
                     Width           =   285
                     _ExtentX        =   503
                     _ExtentY        =   503
                     DataType        =   4
                     DataField       =   "HAutomateReturns"
                     IdName          =   "HAUTOMATERETURNS"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   155
                     Left            =   4065
                     TabIndex        =   198
                     ToolTipText     =   "Specify Local Invoice (Yes/No)"
                     Top             =   3000
                     Width           =   285
                     _ExtentX        =   503
                     _ExtentY        =   503
                     DataType        =   4
                     DataField       =   "HLclInvYn"
                     IdName          =   "HLCLINVYN"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   194
                     Left            =   9075
                     TabIndex        =   200
                     ToolTipText     =   "Specify Whether Automatic Custom Matching Option Has To Be Enabled (Yes/ No)"
                     Top             =   420
                     Width           =   285
                     _ExtentX        =   503
                     _ExtentY        =   503
                     DataType        =   4
                     DataField       =   "HAutoCustmMtch"
                     IdName          =   "HAUTOCUSTMMTCH"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   195
                     Left            =   4065
                     TabIndex        =   199
                     ToolTipText     =   "Specify Whether The Copy Options In Invoice Design Entry Will Keep The Invoice Gross Weight Equal To Actual Gross Weight (Yes/ No)"
                     Top             =   3285
                     Width           =   285
                     _ExtentX        =   503
                     _ExtentY        =   503
                     DataType        =   4
                     DataField       =   "HInInvGrsWtEqActual"
                     IdName          =   "HININVGRSWTEQACTUAL"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   197
                     Left            =   9075
                     TabIndex        =   201
                     ToolTipText     =   "Specify Whether Automatic Custom Matching Option Should Consider The Best Use Fields In Custom Purchase (Yes/ No)"
                     Top             =   705
                     Width           =   285
                     _ExtentX        =   503
                     _ExtentY        =   503
                     DataType        =   4
                     DataField       =   "hBestUse"
                     IdName          =   "HBESTUSE"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   215
                     Left            =   9075
                     TabIndex        =   202
                     ToolTipText     =   "Whether Invoice Rate of Dia and Col stones should be With Respect to Stock Rate"
                     Top             =   990
                     Width           =   285
                     _ExtentX        =   503
                     _ExtentY        =   503
                     DataType        =   4
                     DataField       =   "hInIRtWrtStkYN"
                     IdName          =   "HINIRTWRTSTKYN"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   222
                     Left            =   9075
                     TabIndex        =   205
                     ToolTipText     =   "Specify Whether Commission is Self Payable (Y/N)"
                     Top             =   1860
                     Width           =   285
                     _ExtentX        =   503
                     _ExtentY        =   503
                     DataType        =   4
                     DataField       =   "HHideCommYN"
                     IdName          =   "HHIDECOMMYN"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   225
                     Left            =   9075
                     TabIndex        =   203
                     ToolTipText     =   "Whether Invoice Rate of Accessories should be With Respect to Stock Rate"
                     Top             =   1290
                     Width           =   285
                     _ExtentX        =   503
                     _ExtentY        =   503
                     DataType        =   4
                     DataField       =   "hInXIRtwrtStkYN"
                     IdName          =   "HINXIRTWRTSTKYN"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   227
                     Left            =   9075
                     TabIndex        =   208
                     ToolTipText     =   "Enter 'Y' to Use Notional Rate In Custom Matching Instead of Purchase Rate"
                     Top             =   2715
                     Width           =   285
                     _ExtentX        =   503
                     _ExtentY        =   503
                     DataType        =   4
                     DataField       =   "hInNotionRtYN"
                     IdName          =   "HINNOTIONRTYN"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   236
                     Left            =   8415
                     TabIndex        =   209
                     ToolTipText     =   "Enter Default Invoice Design Sort by Option"
                     Top             =   3000
                     Width           =   945
                     _ExtentX        =   1667
                     _ExtentY        =   503
                     MaxLength       =   5
                     DataType        =   4
                     DataField       =   "HInSrt"
                     IdName          =   "HINSRT"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   238
                     Left            =   9075
                     TabIndex        =   206
                     ToolTipText     =   "Specify Whether Discount is Self Payable (Y/N)"
                     Top             =   2145
                     Width           =   285
                     _ExtentX        =   503
                     _ExtentY        =   503
                     DataType        =   4
                     DataField       =   "HSelfDscYN"
                     IdName          =   "HSELFDSCYN"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   239
                     Left            =   9075
                     TabIndex        =   207
                     ToolTipText     =   "Specify Whether Deduction is Self Payable (Y/N)"
                     Top             =   2430
                     Width           =   285
                     _ExtentX        =   503
                     _ExtentY        =   503
                     DataType        =   4
                     DataField       =   "HSelfDeductYN"
                     IdName          =   "HSELFDEDUCTYN"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   253
                     Left            =   9075
                     TabIndex        =   204
                     ToolTipText     =   "Whether Invoice Rate of Metals should be With Respect to Stock Rate"
                     Top             =   1575
                     Width           =   285
                     _ExtentX        =   503
                     _ExtentY        =   503
                     DataType        =   4
                     DataField       =   "hMetIRtwrtStkYN"
                     IdName          =   "HMETIRTWRTSTKYN"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   255
                     Left            =   9075
                     TabIndex        =   210
                     ToolTipText     =   "Enter 'Y' to Use Notional Rate In Custom Matching Instead of Purchase Rate"
                     Top             =   3285
                     Width           =   285
                     _ExtentX        =   503
                     _ExtentY        =   503
                     DataType        =   4
                     DataField       =   "hRtfrRtChrtYN"
                     IdName          =   "HRTFRRTCHRTYN"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   262
                     Left            =   9075
                     TabIndex        =   211
                     ToolTipText     =   "Enter 'Y' to Enforce From Bag Location in FG Entry"
                     Top             =   3570
                     Width           =   285
                     _ExtentX        =   503
                     _ExtentY        =   503
                     DataType        =   4
                     DataField       =   "HEnfFrBLocInFgYN"
                     IdName          =   "HENFFRBLOCINFGYN"
                  End
                  Begin VB.Label Label24 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Enforce Fr Bag Loc in FG"
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
                     Left            =   5640
                     TabIndex        =   559
                     Top             =   3645
                     Width           =   3240
                  End
                  Begin VB.Label Label23 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Inv Rt from Rt Chart"
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
                     Left            =   5640
                     TabIndex        =   552
                     Top             =   3360
                     Width           =   2400
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Metal Inv Rt Wrt Stock Rt"
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
                     Left            =   5640
                     TabIndex        =   550
                     Top             =   1590
                     Width           =   2535
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Deduction Self Payable"
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
                     Index           =   174
                     Left            =   5640
                     TabIndex        =   534
                     Top             =   2460
                     Width           =   3345
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Discount Self Payable"
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
                     Index           =   173
                     Left            =   5640
                     TabIndex        =   533
                     Top             =   2160
                     Width           =   3345
                  End
                  Begin VB.Label Label22 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Invoice Design Sort By "
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
                     Left            =   5640
                     TabIndex        =   530
                     Top             =   3045
                     Width           =   2760
                  End
                  Begin VB.Label LblNotionRt 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Use Notional Rate"
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
                     Left            =   5640
                     TabIndex        =   518
                     Top             =   2760
                     Width           =   1920
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Acc Inv Rt Wrt Stock Rt"
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
                     Index           =   165
                     Left            =   5640
                     TabIndex        =   517
                     Top             =   1290
                     Width           =   2505
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Commission Self Payable"
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
                     Index           =   162
                     Left            =   5640
                     TabIndex        =   512
                     Top             =   1875
                     Width           =   3345
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "DC Inv Rt Wrt Stock Rt"
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
                     Index           =   155
                     Left            =   5640
                     TabIndex        =   507
                     Top             =   1020
                     Width           =   2295
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Best Use For Custom Matching"
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
                     Index           =   151
                     Left            =   5640
                     TabIndex        =   498
                     Top             =   705
                     Width           =   3405
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Inv Grs Wt Equal To Actual"
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   255
                     Index           =   149
                     Left            =   120
                     TabIndex        =   496
                     Top             =   3345
                     Width           =   2655
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Enable Automatic Custom Matching"
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   255
                     Index           =   148
                     Left            =   5640
                     TabIndex        =   495
                     Top             =   420
                     Width           =   3435
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Local Invoice "
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
                     Index           =   130
                     Left            =   120
                     TabIndex        =   474
                     Top             =   3000
                     Width           =   3315
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Automate Returns"
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
                     Index           =   119
                     Left            =   120
                     TabIndex        =   460
                     Top             =   2715
                     Width           =   3765
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Generate Export No"
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
                     Index           =   117
                     Left            =   120
                     TabIndex        =   455
                     Top             =   420
                     Width           =   3315
                  End
                  Begin VB.Label LblMain 
                     BackStyle       =   0  'Transparent
                     BorderStyle     =   1  'Fixed Single
                     Caption         =   "  Invoice Entry Options"
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
                     Index           =   10
                     Left            =   -30
                     TabIndex        =   433
                     Top             =   90
                     Width           =   9495
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Invoice Type"
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
                     Index           =   86
                     Left            =   120
                     TabIndex        =   393
                     Top             =   2430
                     Width           =   3315
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Customs Val Equal To Inv"
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
                     Index           =   52
                     Left            =   120
                     TabIndex        =   352
                     Top             =   1005
                     Width           =   3315
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Customs Wt Equal To"
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
                     Index           =   40
                     Left            =   120
                     TabIndex        =   351
                     Top             =   720
                     Width           =   3315
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Round Off FOB($) In Invoice"
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
                     Index           =   38
                     Left            =   120
                     TabIndex        =   350
                     Top             =   1290
                     Width           =   3735
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Show Actual Wt Fld In InvDsg Entry"
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
                     Index           =   51
                     Left            =   120
                     TabIndex        =   349
                     Top             =   1575
                     Width           =   3765
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Sort Opt For Custom Purchase Help "
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
                     Index           =   57
                     Left            =   120
                     TabIndex        =   348
                     Top             =   1860
                     Width           =   3795
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "DTA"
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
                     Index           =   69
                     Left            =   120
                     TabIndex        =   347
                     Top             =   2145
                     Width           =   3315
                  End
               End
               Begin VB.Frame FraFgDet 
                  Height          =   1425
                  Left            =   5520
                  TabIndex        =   344
                  Top             =   -120
                  Width           =   9495
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   57
                     Left            =   4050
                     TabIndex        =   185
                     ToolTipText     =   "Enter Default Packing List Character"
                     Top             =   420
                     Width           =   585
                     _ExtentX        =   1032
                     _ExtentY        =   503
                     MaxLength       =   3
                     DataType        =   4
                     DataField       =   "hFgPlChr"
                     IdName          =   "HFGPLCHR"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   96
                     Left            =   4050
                     TabIndex        =   186
                     ToolTipText     =   "Enter Default Finished Goods Sub Location"
                     Top             =   705
                     Width           =   1365
                     _ExtentX        =   2408
                     _ExtentY        =   503
                     MaxLength       =   8
                     DataType        =   4
                     DataField       =   "HFgSubLoc"
                     IdName          =   "HFGSUBLOC"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   156
                     Left            =   4050
                     TabIndex        =   187
                     ToolTipText     =   "Fg Tag Module (Yes/No)"
                     Top             =   990
                     Width           =   285
                     _ExtentX        =   503
                     _ExtentY        =   503
                     DataType        =   4
                     DataField       =   "HFgTagMod"
                     IdName          =   "HFGTAGMOD"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   219
                     Left            =   9090
                     TabIndex        =   188
                     ToolTipText     =   "Fg Transaction Module (Yes/No)"
                     Top             =   1020
                     Width           =   285
                     _ExtentX        =   503
                     _ExtentY        =   503
                     DataType        =   4
                     DataField       =   "HFGIOYN"
                     IdName          =   "HFGIOYN"
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Finish Goods Transaction Module"
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
                     Index           =   159
                     Left            =   5640
                     TabIndex        =   509
                     Top             =   1020
                     Width           =   3435
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "FG Tag Module"
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
                     Index           =   131
                     Left            =   120
                     TabIndex        =   475
                     Top             =   1020
                     Width           =   2355
                  End
                  Begin VB.Label LblMain 
                     BackStyle       =   0  'Transparent
                     BorderStyle     =   1  'Fixed Single
                     Caption         =   "  Finished Goods Options"
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
                     Index           =   11
                     Left            =   -30
                     TabIndex        =   434
                     Top             =   60
                     Width           =   9555
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Default Fg Sub Location"
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
                     Index           =   90
                     Left            =   120
                     TabIndex        =   397
                     Top             =   705
                     Width           =   2655
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Packing List Chr"
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
                     Index           =   56
                     Left            =   120
                     TabIndex        =   345
                     Top             =   420
                     Width           =   2355
                  End
               End
            End
            Begin VB.Frame FraTabHd 
               BorderStyle     =   0  'None
               Height          =   8160
               Index           =   1
               Left            =   -74940
               TabIndex        =   283
               Top             =   360
               Width           =   14950
               Begin VB.Frame FraDef2 
                  Height          =   3105
                  Left            =   -30
                  TabIndex        =   326
                  Top             =   -120
                  Width           =   6675
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   23
                     Left            =   1695
                     TabIndex        =   47
                     ToolTipText     =   "Enter Export Officer"
                     Top             =   675
                     Width           =   4725
                     _ExtentX        =   8334
                     _ExtentY        =   503
                     MaxLength       =   30
                     DataField       =   "HExpOff"
                     IdName          =   "HEXPOFF"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   31
                     Left            =   1695
                     TabIndex        =   49
                     ToolTipText     =   "Enter Export Trade Control Zone"
                     Top             =   1245
                     Width           =   3105
                     _ExtentX        =   5477
                     _ExtentY        =   503
                     MaxLength       =   20
                     DataField       =   "HEtcZone"
                     IdName          =   "HETCZONE"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   32
                     Left            =   1695
                     TabIndex        =   50
                     ToolTipText     =   "Enter Export Trade Control Section"
                     Top             =   1530
                     Width           =   4725
                     _ExtentX        =   8334
                     _ExtentY        =   503
                     MaxLength       =   30
                     DataField       =   "HEtcSection"
                     IdName          =   "HETCSECTION"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   68
                     Left            =   1695
                     TabIndex        =   52
                     ToolTipText     =   "Enter Business Identification Number"
                     Top             =   2100
                     Width           =   3105
                     _ExtentX        =   5477
                     _ExtentY        =   503
                     MaxLength       =   20
                     DataType        =   4
                     DataField       =   "HBinNo"
                     IdName          =   "HBINNO"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   76
                     Left            =   1695
                     TabIndex        =   51
                     ToolTipText     =   "Enter Export Trade Control Saving"
                     Top             =   1815
                     Width           =   3105
                     _ExtentX        =   5477
                     _ExtentY        =   503
                     MaxLength       =   20
                     DataField       =   "HETCSvg"
                     IdName          =   "HETCSVG"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   91
                     Left            =   1695
                     TabIndex        =   46
                     ToolTipText     =   "Enter Authorised Signatory"
                     Top             =   390
                     Width           =   4725
                     _ExtentX        =   8334
                     _ExtentY        =   503
                     MaxLength       =   30
                     DataField       =   "HAuthSgn"
                     IdName          =   "HAUTHSGN"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   109
                     Left            =   1695
                     TabIndex        =   54
                     ToolTipText     =   "Enter MMTC Code"
                     Top             =   2670
                     Width           =   3105
                     _ExtentX        =   5477
                     _ExtentY        =   503
                     MaxLength       =   20
                     DataField       =   "HMMTC"
                     IdName          =   "HMMTC"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   110
                     Left            =   1695
                     TabIndex        =   48
                     ToolTipText     =   "Enter Export Trade Control Text for Reports "
                     Top             =   960
                     Width           =   4725
                     _ExtentX        =   8334
                     _ExtentY        =   503
                     MaxLength       =   30
                     DataField       =   "HEtcTxt"
                     IdName          =   "HETCTXT"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   270
                     Left            =   1695
                     TabIndex        =   53
                     ToolTipText     =   "Enter Central Excise Regn Number"
                     Top             =   2385
                     Width           =   3105
                     _ExtentX        =   5477
                     _ExtentY        =   503
                     MaxLength       =   20
                     DataType        =   4
                     DataField       =   "HCERegnNo"
                     IdName          =   "HCEREGNNO"
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "CE Regn No"
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
                     Index           =   203
                     Left            =   90
                     TabIndex        =   574
                     Top             =   2370
                     Width           =   1365
                  End
                  Begin VB.Label LblMain 
                     BackStyle       =   0  'Transparent
                     BorderStyle     =   1  'Fixed Single
                     Caption         =   "  Defaults"
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
                     Index           =   12
                     Left            =   -30
                     TabIndex        =   435
                     Top             =   90
                     Width           =   6705
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "ETC Text"
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
                     Index           =   97
                     Left            =   90
                     TabIndex        =   409
                     Top             =   945
                     Width           =   1365
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "MMTC Code"
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
                     Index           =   92
                     Left            =   90
                     TabIndex        =   402
                     Top             =   2655
                     Width           =   1365
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Auth. Signature"
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
                     Index           =   85
                     Left            =   90
                     TabIndex        =   392
                     Top             =   390
                     Width           =   1845
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "ETC Saving"
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
                     Index           =   71
                     Left            =   90
                     TabIndex        =   332
                     Top             =   1800
                     Width           =   1275
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Bin No."
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
                     Index           =   66
                     Left            =   90
                     TabIndex        =   330
                     Top             =   2085
                     Width           =   1365
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Exp Officer"
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
                     Left            =   90
                     TabIndex        =   329
                     Top             =   675
                     Width           =   1845
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "ETC Zone"
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
                     Index           =   31
                     Left            =   90
                     TabIndex        =   328
                     Top             =   1230
                     Width           =   1365
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "ETC Section"
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
                     Index           =   30
                     Left            =   90
                     TabIndex        =   327
                     Top             =   1515
                     Width           =   1365
                  End
               End
               Begin VB.Frame FraDef3 
                  Height          =   3120
                  Left            =   6600
                  TabIndex        =   319
                  Top             =   -120
                  Width           =   8385
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   77
                     Left            =   1695
                     TabIndex        =   276
                     ToolTipText     =   "Enter AR4/ AR4A No. And Date (GR & Shipping Bill)"
                     Top             =   690
                     Width           =   3255
                     _ExtentX        =   5741
                     _ExtentY        =   503
                     MaxLength       =   20
                     DataField       =   "HAR4No"
                     IdName          =   "HAR4NO"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   78
                     Left            =   1695
                     TabIndex        =   277
                     ToolTipText     =   "Enter Q/Cert. No & Date (GR & Shipping Bill)"
                     Top             =   975
                     Width           =   3255
                     _ExtentX        =   5741
                     _ExtentY        =   503
                     MaxLength       =   20
                     DataField       =   "HQCertNo"
                     IdName          =   "HQCERTNO"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   79
                     Left            =   1695
                     TabIndex        =   279
                     ToolTipText     =   "Enter Licence Number"
                     Top             =   1545
                     Width           =   3255
                     _ExtentX        =   5741
                     _ExtentY        =   503
                     MaxLength       =   20
                     DataField       =   "HLicNo"
                     IdName          =   "HLICNO"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   80
                     Left            =   1695
                     TabIndex        =   278
                     ToolTipText     =   "Enter Origin Criterion"
                     Top             =   1260
                     Width           =   3255
                     _ExtentX        =   5741
                     _ExtentY        =   503
                     MaxLength       =   20
                     DataField       =   "HGSPOrgCrit"
                     IdName          =   "HGSPORGCRIT"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   52
                     Left            =   1695
                     TabIndex        =   275
                     ToolTipText     =   "Enter Invoice Total Value As "
                     Top             =   390
                     Width           =   570
                     _ExtentX        =   1005
                     _ExtentY        =   503
                     MaxLength       =   8
                     DataType        =   4
                     DataField       =   "HTotAs"
                     IdName          =   "HTOTAS"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   570
                     Index           =   54
                     Left            =   1695
                     TabIndex        =   281
                     ToolTipText     =   "Enter Statistical Code"
                     Top             =   2115
                     Width           =   5745
                     _ExtentX        =   10134
                     _ExtentY        =   1005
                     MaxLength       =   60
                     DataField       =   "hStatCd"
                     IdName          =   "HSTATCD"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   111
                     Left            =   1695
                     TabIndex        =   280
                     ToolTipText     =   "Enter Import Value Text For Reports"
                     Top             =   1830
                     Width           =   3255
                     _ExtentX        =   5741
                     _ExtentY        =   503
                     MaxLength       =   30
                     DataField       =   "hImpValTxt"
                     IdName          =   "HIMPVALTXT"
                  End
                  Begin VB.Label LblMain 
                     BackStyle       =   0  'Transparent
                     BorderStyle     =   1  'Fixed Single
                     Caption         =   "  Defaults"
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
                     Index           =   13
                     Left            =   -30
                     TabIndex        =   436
                     Top             =   90
                     Width           =   8400
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Import Val Txt"
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
                     Index           =   100
                     Left            =   150
                     TabIndex        =   410
                     Top             =   1830
                     Width           =   1485
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Statistical Cd"
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
                     Index           =   53
                     Left            =   150
                     TabIndex        =   325
                     Top             =   2115
                     Width           =   1485
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Total Val As"
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
                     Index           =   49
                     Left            =   150
                     TabIndex        =   324
                     Top             =   390
                     Width           =   1335
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Q/CertNo && Dt"
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
                     Index           =   48
                     Left            =   150
                     TabIndex        =   323
                     Top             =   975
                     Width           =   1545
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "AR4No && Dt"
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
                     Index           =   47
                     Left            =   150
                     TabIndex        =   322
                     Top             =   690
                     Width           =   1395
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Lic Number"
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
                     Index           =   72
                     Left            =   150
                     TabIndex        =   321
                     Top             =   1545
                     Width           =   1275
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Origin Criteria"
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
                     Index           =   70
                     Left            =   150
                     TabIndex        =   320
                     Top             =   1260
                     Width           =   1485
                  End
               End
               Begin VB.Frame FraDef5 
                  Height          =   5415
                  Left            =   0
                  TabIndex        =   447
                  Top             =   2760
                  Width           =   15135
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   570
                     Index           =   90
                     Left            =   1695
                     TabIndex        =   55
                     ToolTipText     =   "Enter Invoice Footer Remarks"
                     Top             =   330
                     Width           =   13245
                     _ExtentX        =   23363
                     _ExtentY        =   1005
                     Alignment       =   3
                     MaxLength       =   360
                     DataField       =   "HFtRem"
                     IdName          =   "HFTREM"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   1065
                     Index           =   55
                     Left            =   1695
                     TabIndex        =   56
                     ToolTipText     =   "Enter Bank Instructions"
                     Top             =   900
                     Width           =   9765
                     _ExtentX        =   17224
                     _ExtentY        =   1879
                     Alignment       =   3
                     MaxLength       =   0
                     DataField       =   "HBnkInstr"
                     IdName          =   "HBNKINSTR"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   1065
                     Index           =   70
                     Left            =   1695
                     TabIndex        =   57
                     ToolTipText     =   "Enter Value Addition Footer"
                     Top             =   1965
                     Width           =   9765
                     _ExtentX        =   17224
                     _ExtentY        =   1879
                     Alignment       =   3
                     MaxLength       =   0
                     DataField       =   "HValAddnFt"
                     IdName          =   "HVALADDNFT"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   1065
                     Index           =   72
                     Left            =   1695
                     TabIndex        =   58
                     ToolTipText     =   "Enter United Nations Resolutions"
                     Top             =   3030
                     Width           =   9765
                     _ExtentX        =   17224
                     _ExtentY        =   1879
                     Alignment       =   3
                     MaxLength       =   0
                     DataField       =   "HUnResolutions"
                     IdName          =   "HUNRESOLUTIONS"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   1065
                     Index           =   128
                     Left            =   1695
                     TabIndex        =   59
                     ToolTipText     =   "Enter Customer Remark"
                     Top             =   4095
                     Width           =   9765
                     _ExtentX        =   17224
                     _ExtentY        =   1879
                     Alignment       =   3
                     MaxLength       =   600
                     DataField       =   "HCUSTREM"
                     IdName          =   "HCUSTREM"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   3345
                     Index           =   260
                     Left            =   11520
                     TabIndex        =   60
                     ToolTipText     =   "Enter GSP HS Code"
                     Top             =   1800
                     Width           =   3045
                     _ExtentX        =   5371
                     _ExtentY        =   5900
                     Alignment       =   3
                     MaxLength       =   250
                     DataField       =   "HHSCODE"
                     IdName          =   "HHSCODE"
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "GSP HS Code"
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
                     Index           =   194
                     Left            =   11520
                     TabIndex        =   557
                     Top             =   1440
                     Width           =   1635
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Bank Instr"
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
                     Index           =   54
                     Left            =   90
                     TabIndex        =   452
                     Top             =   900
                     Width           =   1635
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Inv Footer Rem"
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
                     Index           =   76
                     Left            =   90
                     TabIndex        =   451
                     Top             =   330
                     Width           =   1635
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "ValAddn Footer"
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
                     Index           =   68
                     Left            =   90
                     TabIndex        =   450
                     Top             =   1965
                     Width           =   1635
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "UN Resolutions"
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
                     Index           =   73
                     Left            =   90
                     TabIndex        =   449
                     Top             =   3030
                     Width           =   1635
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Customer Rmk"
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
                     Index           =   113
                     Left            =   90
                     TabIndex        =   448
                     Top             =   4095
                     Width           =   1635
                  End
               End
            End
            Begin VB.Frame FraTabHd 
               BorderStyle     =   0  'None
               Height          =   8160
               Index           =   0
               Left            =   -74940
               TabIndex        =   282
               Top             =   360
               Width           =   14950
               Begin VB.Frame Frame3 
                  Height          =   1515
                  Left            =   6480
                  TabIndex        =   535
                  Top             =   4560
                  Width           =   8565
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   240
                     Left            =   1695
                     TabIndex        =   36
                     ToolTipText     =   "Enter Unit Code for SB File (Shiping Bill)"
                     Top             =   390
                     Width           =   1365
                     _ExtentX        =   2408
                     _ExtentY        =   503
                     MaxLength       =   8
                     DataType        =   4
                     DataField       =   "HUnitCd"
                     IdName          =   "HUNITCD"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   241
                     Left            =   6690
                     TabIndex        =   37
                     ToolTipText     =   "Enter Customs House Code for SB File (Shiping Bill)"
                     Top             =   390
                     Width           =   1725
                     _ExtentX        =   3043
                     _ExtentY        =   503
                     MaxLength       =   6
                     DataType        =   4
                     DataField       =   "HCustomHouseCd"
                     IdName          =   "HCUSTOMHOUSECD"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   242
                     Left            =   5730
                     TabIndex        =   41
                     ToolTipText     =   "Enter Type Of Exporter Code for SB File (Shiping Bill)"
                     Top             =   960
                     Width           =   345
                     _ExtentX        =   609
                     _ExtentY        =   503
                     DataType        =   4
                     DataField       =   "HTypOfExporter"
                     IdName          =   "HTYPOFEXPORTER"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   243
                     Left            =   5730
                     TabIndex        =   39
                     ToolTipText     =   "Enter CHA License Code for SB File (Shiping Bill)"
                     Top             =   675
                     Width           =   2685
                     _ExtentX        =   4736
                     _ExtentY        =   503
                     MaxLength       =   15
                     DataType        =   4
                     DataField       =   "HCHALicenseCd"
                     IdName          =   "HCHALICENSECD"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   244
                     Left            =   1695
                     TabIndex        =   40
                     ToolTipText     =   "Enter Port Of Loading Code for SB File (Shiping Bill)"
                     Top             =   960
                     Width           =   795
                     _ExtentX        =   1402
                     _ExtentY        =   503
                     MaxLength       =   6
                     DataType        =   4
                     DataField       =   "HPoL"
                     IdName          =   "HPOL"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   245
                     Left            =   7650
                     TabIndex        =   42
                     ToolTipText     =   "Enter Unit Of Measurement for SB File (Shiping Bill)"
                     Top             =   960
                     Width           =   765
                     _ExtentX        =   1349
                     _ExtentY        =   503
                     MaxLength       =   6
                     DataType        =   4
                     DataField       =   "HSbUom"
                     IdName          =   "HSBUOM"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   246
                     Left            =   1695
                     TabIndex        =   38
                     ToolTipText     =   "Enter Total No of Packages for SB File (Shiping Bill)"
                     Top             =   675
                     Width           =   825
                     _ExtentX        =   1455
                     _ExtentY        =   503
                     Alignment       =   2
                     Mask            =   "#######0"
                     MaxLength       =   8
                     DataType        =   1
                     DataField       =   "HTotNoOfPack"
                     IdName          =   "HTOTNOOFPACK"
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Tot No of Pack"
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
                     Left            =   90
                     TabIndex        =   543
                     Top             =   675
                     Width           =   2085
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "UOM"
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
                     Index           =   180
                     Left            =   7020
                     TabIndex        =   542
                     Top             =   990
                     Width           =   795
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Port Of Loading"
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
                     Index           =   179
                     Left            =   90
                     TabIndex        =   541
                     Top             =   960
                     Width           =   1545
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "CHA License"
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
                     Index           =   178
                     Left            =   3990
                     TabIndex        =   540
                     Top             =   675
                     Width           =   1365
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Type Of Exporter"
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
                     Index           =   177
                     Left            =   3990
                     TabIndex        =   539
                     Top             =   960
                     Width           =   1635
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Unit Code"
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
                     Index           =   176
                     Left            =   120
                     TabIndex        =   538
                     Top             =   390
                     Width           =   1365
                  End
                  Begin VB.Label LblMain 
                     BackStyle       =   0  'Transparent
                     BorderStyle     =   1  'Fixed Single
                     Caption         =   "  Sb File Defaults"
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
                     Index           =   22
                     Left            =   -60
                     TabIndex        =   537
                     Top             =   90
                     Width           =   8790
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Customs House Cd"
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
                     Left            =   3960
                     TabIndex        =   536
                     Top             =   390
                     Width           =   1995
                  End
               End
               Begin VB.Frame Frame1 
                  Height          =   2445
                  Left            =   0
                  TabIndex        =   519
                  Top             =   6030
                  Width           =   15015
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   570
                     Index           =   113
                     Left            =   1665
                     TabIndex        =   43
                     ToolTipText     =   "Enter State Of Origin"
                     Top             =   330
                     Width           =   13215
                     _ExtentX        =   23310
                     _ExtentY        =   1005
                     Alignment       =   3
                     MaxLength       =   120
                     DataField       =   "hStatOfOrg"
                     IdName          =   "HSTATOFORG"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   570
                     Index           =   114
                     Left            =   1665
                     TabIndex        =   44
                     ToolTipText     =   "Enter Metal BOE Text"
                     Top             =   900
                     Width           =   13215
                     _ExtentX        =   23310
                     _ExtentY        =   1005
                     Alignment       =   3
                     MaxLength       =   120
                     DataField       =   "hMetalBOETxt"
                     IdName          =   "HMETALBOETXT"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   570
                     Index           =   141
                     Left            =   1665
                     TabIndex        =   45
                     ToolTipText     =   "Enter RBI Policy Details"
                     Top             =   1470
                     Width           =   13215
                     _ExtentX        =   23310
                     _ExtentY        =   1005
                     Alignment       =   3
                     MaxLength       =   70
                     DataField       =   "hRbiPolicy"
                     IdName          =   "HRBIPOLICY"
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "State Of Origin"
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
                     Index           =   102
                     Left            =   90
                     TabIndex        =   523
                     Top             =   330
                     Width           =   1515
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Metal BOE Text"
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   525
                     Index           =   103
                     Left            =   90
                     TabIndex        =   522
                     Top             =   900
                     Width           =   1275
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "RBI Policy"
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   375
                     Index           =   107
                     Left            =   90
                     TabIndex        =   521
                     Top             =   1470
                     Width           =   1275
                  End
                  Begin VB.Label LblMain 
                     BackStyle       =   0  'Transparent
                     BorderStyle     =   1  'Fixed Single
                     Caption         =   "  Defaults"
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
                     Index           =   18
                     Left            =   -30
                     TabIndex        =   520
                     Top             =   60
                     Width           =   15015
                  End
               End
               Begin VB.Frame FraIns 
                  Height          =   2595
                  Left            =   -30
                  TabIndex        =   306
                  Top             =   3480
                  Width           =   6465
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   0
                     Left            =   1635
                     TabIndex        =   12
                     ToolTipText     =   "Enter Insurance Number"
                     Top             =   675
                     Width           =   4575
                     _ExtentX        =   8070
                     _ExtentY        =   503
                     MaxLength       =   30
                     DataType        =   4
                     DataField       =   "HInsNo"
                     IdName          =   "HINSNO"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   16
                     Left            =   1635
                     TabIndex        =   11
                     ToolTipText     =   "Enter Insurance By"
                     Top             =   390
                     Width           =   4575
                     _ExtentX        =   8070
                     _ExtentY        =   503
                     MaxLength       =   30
                     DataField       =   "HInsBy"
                     IdName          =   "HINSBY"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   4
                     Left            =   1635
                     TabIndex        =   13
                     ToolTipText     =   "Enter Agent Name"
                     Top             =   960
                     Width           =   4575
                     _ExtentX        =   8070
                     _ExtentY        =   503
                     MaxLength       =   30
                     DataField       =   "HAgName"
                     IdName          =   "HAGNAME"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   5
                     Left            =   1635
                     TabIndex        =   14
                     ToolTipText     =   "Enter Agent Address"
                     Top             =   1245
                     Width           =   4575
                     _ExtentX        =   8070
                     _ExtentY        =   503
                     MaxLength       =   30
                     DataField       =   "HAgAdd1"
                     IdName          =   "HAGADD1"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   37
                     Left            =   1635
                     TabIndex        =   15
                     ToolTipText     =   "Enter Agent Address"
                     Top             =   1530
                     Width           =   4575
                     _ExtentX        =   8070
                     _ExtentY        =   503
                     MaxLength       =   30
                     DataField       =   "HAgAdd2"
                     IdName          =   "HAGADD2"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   38
                     Left            =   1635
                     TabIndex        =   16
                     ToolTipText     =   "Enter Agent Address"
                     Top             =   1815
                     Width           =   4575
                     _ExtentX        =   8070
                     _ExtentY        =   503
                     MaxLength       =   30
                     DataField       =   "HAgAdd3"
                     IdName          =   "HAGADD3"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   112
                     Left            =   1635
                     TabIndex        =   17
                     ToolTipText     =   "Enter Shipping Bill Signed By"
                     Top             =   2100
                     Width           =   4575
                     _ExtentX        =   8070
                     _ExtentY        =   503
                     MaxLength       =   45
                     DataField       =   "hShpBillSgnBy"
                     IdName          =   "HSHPBILLSGNBY"
                  End
                  Begin VB.Label LblMain 
                     BackStyle       =   0  'Transparent
                     BorderStyle     =   1  'Fixed Single
                     Caption         =   "  Insurance"
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
                     Index           =   15
                     Left            =   -30
                     TabIndex        =   438
                     Top             =   90
                     Width           =   6525
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Shipping Bill Signed By"
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   525
                     Index           =   101
                     Left            =   120
                     TabIndex        =   411
                     Top             =   1980
                     Width           =   1455
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Number"
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
                     Left            =   90
                     TabIndex        =   310
                     Top             =   675
                     Width           =   1275
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Insurance By"
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
                     Left            =   90
                     TabIndex        =   309
                     Top             =   390
                     Width           =   1395
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Agent Name"
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
                     Left            =   90
                     TabIndex        =   308
                     Top             =   960
                     Width           =   1455
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Agent Address"
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
                     Left            =   90
                     TabIndex        =   307
                     Top             =   1245
                     Width           =   1575
                  End
               End
               Begin VB.Frame FraFrgnBnk 
                  Height          =   1185
                  Left            =   -30
                  TabIndex        =   294
                  Top             =   2550
                  Width           =   6465
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   22
                     Left            =   1635
                     TabIndex        =   9
                     ToolTipText     =   "Enter Foreign Bank Name"
                     Top             =   390
                     Width           =   4575
                     _ExtentX        =   8070
                     _ExtentY        =   503
                     MaxLength       =   30
                     DataField       =   "HFrgnBkName"
                     IdName          =   "HFRGNBKNAME"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   26
                     Left            =   1635
                     TabIndex        =   10
                     ToolTipText     =   "Enter Foreign Bank Account"
                     Top             =   675
                     Width           =   3105
                     _ExtentX        =   5477
                     _ExtentY        =   503
                     MaxLength       =   20
                     DataType        =   4
                     DataField       =   "HFrgnBkAc"
                     IdName          =   "HFRGNBKAC"
                  End
                  Begin VB.Label LblMain 
                     BackStyle       =   0  'Transparent
                     BorderStyle     =   1  'Fixed Single
                     Caption         =   "  Foreign Bank"
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
                     Index           =   16
                     Left            =   -30
                     TabIndex        =   439
                     Top             =   90
                     Width           =   6525
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Name"
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
                     Left            =   90
                     TabIndex        =   296
                     Top             =   390
                     Width           =   1065
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "A/c"
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
                     Index           =   23
                     Left            =   90
                     TabIndex        =   295
                     Top             =   675
                     Width           =   1305
                  End
               End
               Begin VB.Frame FraBnk 
                  Height          =   2865
                  Left            =   -30
                  TabIndex        =   290
                  Top             =   -120
                  Width           =   6465
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   20
                     Left            =   1635
                     TabIndex        =   4
                     ToolTipText     =   "Enter Bank Name"
                     Top             =   390
                     Width           =   4575
                     _ExtentX        =   8070
                     _ExtentY        =   503
                     MaxLength       =   30
                     DataField       =   "HBkName"
                     IdName          =   "HBKNAME"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   570
                     Index           =   17
                     Left            =   1635
                     TabIndex        =   5
                     ToolTipText     =   "Enter Bank Address"
                     Top             =   675
                     Width           =   4575
                     _ExtentX        =   8070
                     _ExtentY        =   1005
                     Alignment       =   3
                     MaxLength       =   45
                     DataField       =   "HBkAdd1"
                     IdName          =   "HBKADD1"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   570
                     Index           =   18
                     Left            =   1635
                     TabIndex        =   6
                     ToolTipText     =   "Enter Bank Address"
                     Top             =   1245
                     Width           =   4575
                     _ExtentX        =   8070
                     _ExtentY        =   1005
                     Alignment       =   3
                     MaxLength       =   45
                     DataField       =   "HBkAdd2"
                     IdName          =   "HBKADD2"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   570
                     Index           =   19
                     Left            =   1635
                     TabIndex        =   7
                     ToolTipText     =   "Enter Bank Address"
                     Top             =   1815
                     Width           =   4575
                     _ExtentX        =   8070
                     _ExtentY        =   1005
                     Alignment       =   3
                     MaxLength       =   45
                     DataField       =   "HBkAdd3"
                     IdName          =   "HBKADD3"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   21
                     Left            =   1635
                     TabIndex        =   8
                     ToolTipText     =   "Enter Bank Account"
                     Top             =   2385
                     Width           =   3105
                     _ExtentX        =   5477
                     _ExtentY        =   503
                     MaxLength       =   20
                     DataType        =   4
                     DataField       =   "HBkAc"
                     IdName          =   "HBKAC"
                  End
                  Begin VB.Label LblMain 
                     BackStyle       =   0  'Transparent
                     BorderStyle     =   1  'Fixed Single
                     Caption         =   "  Bank"
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
                     Index           =   17
                     Left            =   -30
                     TabIndex        =   440
                     Top             =   90
                     Width           =   6465
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "A/c"
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
                     Index           =   18
                     Left            =   90
                     TabIndex        =   293
                     Top             =   2385
                     Width           =   1275
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Address"
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
                     Left            =   90
                     TabIndex        =   292
                     Top             =   675
                     Width           =   1275
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Name"
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
                     Left            =   90
                     TabIndex        =   291
                     Top             =   390
                     Width           =   1275
                  End
               End
               Begin VB.Frame FraDef1 
                  Height          =   4755
                  Left            =   6390
                  TabIndex        =   297
                  Top             =   -120
                  Width           =   8610
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   30
                     Left            =   1995
                     TabIndex        =   27
                     ToolTipText     =   "Enter ITC Number"
                     Top             =   2955
                     Width           =   3630
                     _ExtentX        =   6403
                     _ExtentY        =   503
                     MaxLength       =   20
                     DataType        =   4
                     DataField       =   "HItcNo"
                     IdName          =   "HITCNO"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   34
                     Left            =   7875
                     TabIndex        =   32
                     ToolTipText     =   "Enter Value Addition Percentage For Accessories"
                     Top             =   675
                     Width           =   585
                     _ExtentX        =   1032
                     _ExtentY        =   503
                     Alignment       =   2
                     Mask            =   "#0.00"
                     MaxLength       =   5
                     DataType        =   2
                     DataField       =   "HVaX"
                     IdName          =   "HVAX"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   570
                     Index           =   35
                     Left            =   1995
                     TabIndex        =   29
                     ToolTipText     =   "Enter Product Description"
                     Top             =   3525
                     Width           =   6465
                     _ExtentX        =   11404
                     _ExtentY        =   1005
                     Alignment       =   3
                     MaxLength       =   180
                     DataField       =   "HPrdDesc"
                     ReCalcOn        =   "WINVPRDDESC"
                     IdName          =   "HPRDDESC"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   36
                     Left            =   7875
                     TabIndex        =   31
                     ToolTipText     =   "Enter Value Addition Percentage For Dia And ColStn"
                     Top             =   390
                     Width           =   585
                     _ExtentX        =   1032
                     _ExtentY        =   503
                     Alignment       =   2
                     Mask            =   "#0.00"
                     MaxLength       =   5
                     DataType        =   2
                     DataField       =   "HVaDc"
                     IdName          =   "HVADC"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   29
                     Left            =   1995
                     TabIndex        =   25
                     ToolTipText     =   "Enter Marks"
                     Top             =   2385
                     Width           =   3630
                     _ExtentX        =   6403
                     _ExtentY        =   503
                     MaxLength       =   20
                     DataField       =   "HMarks"
                     IdName          =   "HMARKS"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   33
                     Left            =   1995
                     TabIndex        =   26
                     ToolTipText     =   "Enter Package Description"
                     Top             =   2670
                     Width           =   3630
                     _ExtentX        =   6403
                     _ExtentY        =   503
                     MaxLength       =   20
                     DataField       =   "HPkgDesc"
                     IdName          =   "HPKGDESC"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   24
                     Left            =   1995
                     TabIndex        =   18
                     ToolTipText     =   "Enter RBI Numer"
                     Top             =   390
                     Width           =   3630
                     _ExtentX        =   6403
                     _ExtentY        =   503
                     MaxLength       =   20
                     DataType        =   4
                     DataField       =   "HRbiNo"
                     IdName          =   "HRBINO"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   25
                     Left            =   1995
                     TabIndex        =   19
                     ToolTipText     =   "Enter Import Export Code Number"
                     Top             =   675
                     Width           =   3630
                     _ExtentX        =   6403
                     _ExtentY        =   503
                     MaxLength       =   20
                     DataType        =   4
                     DataField       =   "HIecNo"
                     IdName          =   "HIECNO"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   2
                     Left            =   1995
                     TabIndex        =   20
                     ToolTipText     =   "Enter Pre Carriage"
                     Top             =   960
                     Width           =   3630
                     _ExtentX        =   6403
                     _ExtentY        =   503
                     MaxLength       =   20
                     DataField       =   "HPreCarriage"
                     IdName          =   "HPRECARRIAGE"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   13
                     Left            =   1995
                     TabIndex        =   21
                     ToolTipText     =   "Enter Pre Carrier"
                     Top             =   1245
                     Width           =   3630
                     _ExtentX        =   6403
                     _ExtentY        =   503
                     MaxLength       =   20
                     DataField       =   "HPreCarrier"
                     IdName          =   "HPRECARRIER"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   15
                     Left            =   1995
                     TabIndex        =   22
                     ToolTipText     =   "Enter Vessel"
                     Top             =   1530
                     Width           =   3630
                     _ExtentX        =   6403
                     _ExtentY        =   503
                     MaxLength       =   20
                     DataField       =   "Hvessel"
                     IdName          =   "HVESSEL"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   27
                     Left            =   1995
                     TabIndex        =   23
                     ToolTipText     =   "Enter Port Of Landing"
                     Top             =   1815
                     Width           =   3630
                     _ExtentX        =   6403
                     _ExtentY        =   503
                     MaxLength       =   20
                     DataType        =   4
                     DataField       =   "HPortofLd"
                     IdName          =   "HPORTOFLD"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   28
                     Left            =   1995
                     TabIndex        =   24
                     ToolTipText     =   "Enter Country Of Origin"
                     Top             =   2100
                     Width           =   3630
                     _ExtentX        =   6403
                     _ExtentY        =   503
                     MaxLength       =   20
                     DataType        =   4
                     DataField       =   "HConOfOrg"
                     IdName          =   "HCONOFORG"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   570
                     Index           =   89
                     Left            =   1995
                     TabIndex        =   30
                     ToolTipText     =   "Enter Consignee"
                     Top             =   4095
                     Width           =   3630
                     _ExtentX        =   6403
                     _ExtentY        =   1005
                     Alignment       =   3
                     MaxLength       =   45
                     DataField       =   "HConsignee"
                     IdName          =   "HCONSIGNEE"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   81
                     Left            =   1995
                     TabIndex        =   28
                     ToolTipText     =   "Enter Invoice Product Description Code"
                     Top             =   3240
                     Width           =   3630
                     _ExtentX        =   6403
                     _ExtentY        =   503
                     MaxLength       =   20
                     DataType        =   4
                     ReCalcParent    =   "HPRDDESC"
                     IdName          =   "WINVPRDDESC"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   116
                     Left            =   7635
                     TabIndex        =   33
                     ToolTipText     =   " Enter Platinum  Fineness"
                     Top             =   1245
                     Width           =   825
                     _ExtentX        =   1455
                     _ExtentY        =   503
                     Alignment       =   2
                     Mask            =   "#0.000"
                     MaxLength       =   6
                     DataType        =   2
                     DataField       =   "hPlFiness"
                     IdName          =   "HPLFINESS"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   130
                     Left            =   7635
                     TabIndex        =   34
                     ToolTipText     =   " Enter Silver Fineness"
                     Top             =   1530
                     Width           =   825
                     _ExtentX        =   1455
                     _ExtentY        =   503
                     Alignment       =   2
                     Mask            =   "#0.000"
                     MaxLength       =   6
                     DataType        =   2
                     DataField       =   "hSlFiness"
                     IdName          =   "HSLFINESS"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   207
                     Left            =   7635
                     TabIndex        =   35
                     ToolTipText     =   " Enter Paladium  Fineness"
                     Top             =   1815
                     Width           =   825
                     _ExtentX        =   1455
                     _ExtentY        =   503
                     Alignment       =   2
                     Mask            =   "#0.000"
                     MaxLength       =   6
                     DataType        =   2
                     DataField       =   "HPdFiness"
                     IdName          =   "HPDFINESS"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   117
                     Left            =   7635
                     TabIndex        =   587
                     ToolTipText     =   "Enter Gold Fineness"
                     Top             =   960
                     Width           =   825
                     _ExtentX        =   1455
                     _ExtentY        =   503
                     Alignment       =   2
                     Mask            =   "#0.000"
                     MaxLength       =   6
                     DataType        =   2
                     DataField       =   "hGldFiness"
                     IdName          =   "HGLDFINESS"
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Pd Fness"
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
                     Index           =   153
                     Left            =   6000
                     TabIndex        =   503
                     Top             =   1815
                     Width           =   945
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Slv Fness"
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
                     Index           =   114
                     Left            =   6000
                     TabIndex        =   444
                     Top             =   1545
                     Width           =   945
                  End
                  Begin VB.Label LblMain 
                     BackStyle       =   0  'Transparent
                     BorderStyle     =   1  'Fixed Single
                     Caption         =   "  Defaults"
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
                     Index           =   14
                     Left            =   -30
                     TabIndex        =   437
                     Top             =   90
                     Width           =   8700
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Plt Fness"
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
                     Index           =   105
                     Left            =   6000
                     TabIndex        =   414
                     Top             =   1245
                     Width           =   885
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Gld Fness"
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
                     Index           =   104
                     Left            =   6000
                     TabIndex        =   413
                     Top             =   960
                     Width           =   1095
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Prd Desc Cd"
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
                     Index           =   77
                     Left            =   120
                     TabIndex        =   383
                     Top             =   3240
                     Width           =   1365
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Consignee"
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
                     Index           =   75
                     Left            =   120
                     TabIndex        =   318
                     Top             =   4080
                     Width           =   1245
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Pre Carriage"
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
                     Left            =   120
                     TabIndex        =   317
                     Top             =   960
                     Width           =   1545
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Pre Carrier"
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
                     Index           =   28
                     Left            =   120
                     TabIndex        =   316
                     Top             =   1245
                     Width           =   1545
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Vessel"
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
                     Index           =   27
                     Left            =   120
                     TabIndex        =   315
                     Top             =   1530
                     Width           =   1545
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Port Of Ld"
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
                     Index           =   26
                     Left            =   120
                     TabIndex        =   314
                     Top             =   1815
                     Width           =   1665
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Cnty Of Org"
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
                     Index           =   12
                     Left            =   120
                     TabIndex        =   313
                     Top             =   2100
                     Width           =   1785
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "IEC No."
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
                     Index           =   22
                     Left            =   120
                     TabIndex        =   312
                     Top             =   675
                     Width           =   1635
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "RBI No."
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
                     Index           =   21
                     Left            =   120
                     TabIndex        =   311
                     Top             =   390
                     Width           =   1725
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Marks"
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
                     Index           =   25
                     Left            =   120
                     TabIndex        =   305
                     Top             =   2385
                     Width           =   1365
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Pck Desc"
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
                     Index           =   24
                     Left            =   120
                     TabIndex        =   302
                     Top             =   2670
                     Width           =   1785
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Val Addn % DC"
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
                     Index           =   34
                     Left            =   6000
                     TabIndex        =   301
                     Top             =   390
                     Width           =   1605
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Product Desc"
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
                     Index           =   33
                     Left            =   120
                     TabIndex        =   300
                     Top             =   3525
                     Width           =   1455
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Val Addn % X"
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
                     Index           =   32
                     Left            =   6000
                     TabIndex        =   299
                     Top             =   675
                     Width           =   1605
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "ITC No."
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
                     Left            =   120
                     TabIndex        =   298
                     Top             =   2955
                     Width           =   1365
                  End
               End
            End
            Begin VB.Frame FraTabHd 
               BorderStyle     =   0  'None
               Caption         =   "Frame4"
               Height          =   7575
               Index           =   6
               Left            =   -74640
               TabIndex        =   566
               Top             =   480
               Width           =   14535
               Begin VB.Frame Frame4 
                  BorderStyle     =   0  'None
                  Height          =   5415
                  Left            =   0
                  TabIndex        =   567
                  Top             =   240
                  Width           =   14415
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   266
                     Left            =   3075
                     TabIndex        =   568
                     ToolTipText     =   "Specify Whether To Fix Inv/Customs Price = Order Sale Value During Invoicing Or Calculate The Price (Y / N)"
                     Top             =   0
                     Width           =   315
                     _ExtentX        =   556
                     _ExtentY        =   503
                     DataType        =   4
                     DataField       =   "HjStkNumjFrBagYN"
                     IdName          =   "HJSTKNUMJFRBAGYN"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   267
                     Left            =   3075
                     TabIndex        =   569
                     ToolTipText     =   "Enter Default Year"
                     Top             =   285
                     Width           =   915
                     _ExtentX        =   1614
                     _ExtentY        =   503
                     MaxLength       =   3
                     DataType        =   4
                     DataField       =   "hjStkDefChr"
                     IdName          =   "HJSTKDEFCHR"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   271
                     Left            =   3075
                     TabIndex        =   598
                     ToolTipText     =   "Enter Invoice Amount Description 1"
                     Top             =   570
                     Width           =   6660
                     _ExtentX        =   11748
                     _ExtentY        =   503
                     MaxLength       =   45
                     DataField       =   "HInAmtDesc1"
                     IdName          =   "HINAMTDESC1"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   272
                     Left            =   3075
                     TabIndex        =   599
                     ToolTipText     =   "Enter Invoice Amount Description 2"
                     Top             =   855
                     Width           =   6660
                     _ExtentX        =   11748
                     _ExtentY        =   503
                     MaxLength       =   45
                     DataField       =   "HInAmtDesc2"
                     IdName          =   "HINAMTDESC2"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   273
                     Left            =   3075
                     TabIndex        =   600
                     ToolTipText     =   "Enter Invoice Amount Description 3"
                     Top             =   1140
                     Width           =   6660
                     _ExtentX        =   11748
                     _ExtentY        =   503
                     MaxLength       =   45
                     DataField       =   "HInAmtDesc3"
                     IdName          =   "HINAMTDESC3"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   274
                     Left            =   3075
                     TabIndex        =   601
                     ToolTipText     =   "Enter Invoice Amount Description 4"
                     Top             =   1425
                     Width           =   6660
                     _ExtentX        =   11748
                     _ExtentY        =   503
                     MaxLength       =   45
                     DataField       =   "HInAmtDesc4"
                     IdName          =   "HINAMTDESC4"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   275
                     Left            =   3075
                     TabIndex        =   602
                     ToolTipText     =   "Enter Invoice Amount Description 5"
                     Top             =   1710
                     Width           =   6660
                     _ExtentX        =   11748
                     _ExtentY        =   503
                     MaxLength       =   45
                     DataField       =   "HInAmtDesc5"
                     IdName          =   "HINAMTDESC5"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   276
                     Left            =   3075
                     TabIndex        =   603
                     ToolTipText     =   "Enter Invoice Amount Description 6"
                     Top             =   1995
                     Width           =   6660
                     _ExtentX        =   11748
                     _ExtentY        =   503
                     MaxLength       =   45
                     DataField       =   "HInAmtDesc6"
                     IdName          =   "HINAMTDESC6"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   277
                     Left            =   3075
                     TabIndex        =   604
                     ToolTipText     =   "Enter Invoice Amount Description 7"
                     Top             =   2280
                     Width           =   6660
                     _ExtentX        =   11748
                     _ExtentY        =   503
                     MaxLength       =   45
                     DataField       =   "HInAmtDesc7"
                     IdName          =   "HINAMTDESC7"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   278
                     Left            =   3075
                     TabIndex        =   605
                     ToolTipText     =   "Enter Invoice Amount Description 8"
                     Top             =   2565
                     Width           =   6660
                     _ExtentX        =   11748
                     _ExtentY        =   503
                     MaxLength       =   45
                     DataField       =   "HInAmtDesc8"
                     IdName          =   "HINAMTDESC8"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   279
                     Left            =   3075
                     TabIndex        =   606
                     ToolTipText     =   "Enter Invoice Amount Description 9"
                     Top             =   2850
                     Width           =   6660
                     _ExtentX        =   11748
                     _ExtentY        =   503
                     MaxLength       =   45
                     DataField       =   "HInAmtDesc9"
                     IdName          =   "HINAMTDESC9"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   280
                     Left            =   3075
                     TabIndex        =   607
                     ToolTipText     =   "Enter Invoice Amount Description 10"
                     Top             =   3135
                     Width           =   6660
                     _ExtentX        =   11748
                     _ExtentY        =   503
                     MaxLength       =   45
                     DataField       =   "HInAmtDesc10"
                     IdName          =   "HINAMTDESC10"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   281
                     Left            =   3075
                     TabIndex        =   608
                     ToolTipText     =   "Enter Invoice Amount Description 11"
                     Top             =   3420
                     Width           =   6660
                     _ExtentX        =   11748
                     _ExtentY        =   503
                     MaxLength       =   45
                     DataField       =   "HInAmtDesc11"
                     IdName          =   "HINAMTDESC11"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   282
                     Left            =   3075
                     TabIndex        =   609
                     ToolTipText     =   "Enter Invoice Amount Description 12"
                     Top             =   3705
                     Width           =   6660
                     _ExtentX        =   11748
                     _ExtentY        =   503
                     MaxLength       =   45
                     DataField       =   "HInAmtDesc12"
                     IdName          =   "HINAMTDESC12"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   283
                     Left            =   12000
                     TabIndex        =   610
                     ToolTipText     =   "Enter Invoice Amount 1 %"
                     Top             =   600
                     Width           =   825
                     _ExtentX        =   1455
                     _ExtentY        =   503
                     Alignment       =   2
                     Mask            =   "#0.00"
                     MaxLength       =   5
                     DataType        =   2
                     DataField       =   "HInAmt1Per"
                     IdName          =   "HINVAMT1PER"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   284
                     Left            =   12000
                     TabIndex        =   611
                     ToolTipText     =   "Enter Invoice Amount 2 %"
                     Top             =   885
                     Width           =   825
                     _ExtentX        =   1455
                     _ExtentY        =   503
                     Alignment       =   2
                     Mask            =   "#0.00"
                     MaxLength       =   5
                     DataType        =   2
                     DataField       =   "HInAmt2Per"
                     IdName          =   "HINVAMT2PER"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   285
                     Left            =   12000
                     TabIndex        =   612
                     ToolTipText     =   "Enter Invoice Amount 3 %"
                     Top             =   1170
                     Width           =   825
                     _ExtentX        =   1455
                     _ExtentY        =   503
                     Alignment       =   2
                     Mask            =   "#0.00"
                     MaxLength       =   5
                     DataType        =   2
                     DataField       =   "HInAmt3Per"
                     IdName          =   "HINVAMT3PER"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   286
                     Left            =   12000
                     TabIndex        =   613
                     ToolTipText     =   "Enter Invoice Amount 4 %"
                     Top             =   1455
                     Width           =   825
                     _ExtentX        =   1455
                     _ExtentY        =   503
                     Alignment       =   2
                     Mask            =   "#0.00"
                     MaxLength       =   5
                     DataType        =   2
                     DataField       =   "HInAmt4Per"
                     IdName          =   "HINVAMT4PER"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   287
                     Left            =   12000
                     TabIndex        =   614
                     ToolTipText     =   "Enter Invoice Amount 5 %"
                     Top             =   1740
                     Width           =   825
                     _ExtentX        =   1455
                     _ExtentY        =   503
                     Alignment       =   2
                     Mask            =   "#0.00"
                     MaxLength       =   5
                     DataType        =   2
                     DataField       =   "HInAmt5Per"
                     IdName          =   "HINVAMT5PER"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   288
                     Left            =   12000
                     TabIndex        =   615
                     ToolTipText     =   "Enter Invoice Amount 6 %"
                     Top             =   2025
                     Width           =   825
                     _ExtentX        =   1455
                     _ExtentY        =   503
                     Alignment       =   2
                     Mask            =   "#0.00"
                     MaxLength       =   5
                     DataType        =   2
                     DataField       =   "HInAmt6Per"
                     IdName          =   "HINVAMT6PER"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   289
                     Left            =   12000
                     TabIndex        =   616
                     ToolTipText     =   "Enter Invoice Amount 8 %"
                     Top             =   2595
                     Width           =   825
                     _ExtentX        =   1455
                     _ExtentY        =   503
                     Alignment       =   2
                     Mask            =   "#0.00"
                     MaxLength       =   5
                     DataType        =   2
                     DataField       =   "HInAmt8Per"
                     IdName          =   "HINVAMT8PER"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   290
                     Left            =   12000
                     TabIndex        =   617
                     ToolTipText     =   "Enter Invoice Amount 9 %"
                     Top             =   2880
                     Width           =   825
                     _ExtentX        =   1455
                     _ExtentY        =   503
                     Alignment       =   2
                     Mask            =   "#0.00"
                     MaxLength       =   5
                     DataType        =   2
                     DataField       =   "HInAmt9Per"
                     IdName          =   "HINVAMT9PER"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   291
                     Left            =   12000
                     TabIndex        =   618
                     ToolTipText     =   "Enter Invoice Amount 10 %"
                     Top             =   3165
                     Width           =   825
                     _ExtentX        =   1455
                     _ExtentY        =   503
                     Alignment       =   2
                     Mask            =   "#0.00"
                     MaxLength       =   5
                     DataType        =   2
                     DataField       =   "HInAmt10Per"
                     IdName          =   "HINVAMT10PER"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   292
                     Left            =   12000
                     TabIndex        =   619
                     ToolTipText     =   "Enter Invoice Amount 11 %"
                     Top             =   3450
                     Width           =   825
                     _ExtentX        =   1455
                     _ExtentY        =   503
                     Alignment       =   2
                     Mask            =   "#0.00"
                     MaxLength       =   5
                     DataType        =   2
                     DataField       =   "HInAmt11Per"
                     IdName          =   "HINVAMT11PER"
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Inv Amt11 %"
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
                     Index           =   217
                     Left            =   10440
                     TabIndex        =   597
                     Top             =   3450
                     Width           =   1395
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Inv Amt10 %"
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
                     Index           =   218
                     Left            =   10440
                     TabIndex        =   596
                     Top             =   3165
                     Width           =   1395
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Inv Amt9 %"
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
                     Index           =   219
                     Left            =   10440
                     TabIndex        =   595
                     Top             =   2880
                     Width           =   1395
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Inv Amt5 %"
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
                     Index           =   221
                     Left            =   10440
                     TabIndex        =   594
                     Top             =   1740
                     Width           =   1395
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Inv Amt8 %"
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
                     Index           =   222
                     Left            =   10440
                     TabIndex        =   593
                     Top             =   2595
                     Width           =   1395
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Inv Amt6 %"
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
                     Index           =   223
                     Left            =   10440
                     TabIndex        =   592
                     Top             =   2025
                     Width           =   1395
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Inv Amt4 %"
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
                     Index           =   224
                     Left            =   10440
                     TabIndex        =   591
                     Top             =   1455
                     Width           =   1395
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Inv Amt3 %"
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
                     Index           =   225
                     Left            =   10440
                     TabIndex        =   590
                     Top             =   1170
                     Width           =   1395
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Inv Amt2 %"
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
                     Index           =   226
                     Left            =   10440
                     TabIndex        =   589
                     Top             =   885
                     Width           =   1395
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Inv Amt1 %"
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
                     Index           =   227
                     Left            =   10440
                     TabIndex        =   588
                     Top             =   600
                     Width           =   1395
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Inv Amt12 Desc"
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
                     Index           =   215
                     Left            =   240
                     TabIndex        =   586
                     Top             =   3735
                     Width           =   2055
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Inv Amt11 Desc"
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
                     Index           =   214
                     Left            =   240
                     TabIndex        =   585
                     Top             =   3450
                     Width           =   2055
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Inv Amt10 Desc"
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
                     Index           =   213
                     Left            =   240
                     TabIndex        =   584
                     Top             =   3165
                     Width           =   2055
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Inv Amt9 Desc"
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
                     Index           =   212
                     Left            =   240
                     TabIndex        =   583
                     Top             =   2880
                     Width           =   2055
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Inv Amt7 Desc"
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
                     Index           =   211
                     Left            =   240
                     TabIndex        =   582
                     Top             =   2310
                     Width           =   2055
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Inv Amt5 Desc"
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
                     Index           =   210
                     Left            =   240
                     TabIndex        =   581
                     Top             =   1740
                     Width           =   2055
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Inv Amt8 Desc"
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
                     Index           =   209
                     Left            =   240
                     TabIndex        =   580
                     Top             =   2595
                     Width           =   2055
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Inv Amt6 Desc"
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
                     Index           =   208
                     Left            =   240
                     TabIndex        =   579
                     Top             =   2025
                     Width           =   2055
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Inv Amt4 Desc"
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
                     Index           =   207
                     Left            =   240
                     TabIndex        =   578
                     Top             =   1455
                     Width           =   2055
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Inv Amt3 Desc"
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
                     Index           =   206
                     Left            =   240
                     TabIndex        =   577
                     Top             =   1170
                     Width           =   2055
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Inv Amt2 Desc"
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
                     Index           =   205
                     Left            =   240
                     TabIndex        =   576
                     Top             =   885
                     Width           =   2055
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Inv Amt1 Desc"
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
                     Index           =   204
                     Left            =   240
                     TabIndex        =   575
                     Top             =   600
                     Width           =   2055
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Copy DmStkNum from Bag"
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
                     Index           =   199
                     Left            =   240
                     TabIndex        =   571
                     Top             =   0
                     Width           =   2535
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Default DmStkChr"
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
                     Index           =   200
                     Left            =   240
                     TabIndex        =   570
                     Top             =   285
                     Width           =   2055
                  End
               End
            End
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   3
            Left            =   675
            TabIndex        =   2
            ToolTipText     =   "Enter Company Name"
            Top             =   270
            Width           =   4605
            _ExtentX        =   8123
            _ExtentY        =   503
            MaxLength       =   30
            DataField       =   "Hname"
            IdName          =   "HNAME"
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Name"
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
            Left            =   0
            TabIndex        =   289
            Top             =   270
            Width           =   855
         End
      End
      Begin VB.Frame FraRefFileLst 
         Height          =   1140
         Left            =   5175
         TabIndex        =   514
         Top             =   7950
         Width           =   4860
         Begin MwfCtl.MWCTL_BTN1 CmdRefFileLstGo 
            Height          =   465
            Left            =   4050
            TabIndex        =   515
            Top             =   600
            Width           =   690
            _ExtentX        =   1217
            _ExtentY        =   820
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
         Begin VB.Label LblMain 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Picture File List from MwPict will be refreshed. This may take few minutes."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1050
            Index           =   21
            Left            =   30
            TabIndex        =   516
            Top             =   90
            Width           =   4830
         End
      End
   End
End
Attribute VB_Name = "EmrFrmHead"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---- General Comments on the Entry --------'
' 1 Menu
' X-Head
'ss
' TabPage, TabHd, 5 Tabs
'   1. Invoice Details
'   2. Invoice Details (Continued)
'   3. Other Details
'
' Only one entry is allowed in the Head table i.e. HCd= 'ZSELF'
' The Invoice Details is used to enter the details of the company for Invoice Purposes
' The Other Details include the defaults used in the system for various entries
'---------------------

Option Explicit
Dim moCn As MwfLib.MDOConnection   '*** (Bef speed) bef 24/11/06

Dim mb_DsgCat As Boolean      ' ***** Manali 3.5.0 - 09/06/09 - Refresh Pic File List

'*** For the Toggle Options of Command Buttons
Enum en_HeadFra
  RefFileLst = 0    ' ***** Manali 3.5.0 - 09/06/09 - Refresh Pic File List
End Enum
Dim ms_ChkQtsIdName As String
Private Sub ADC_Load()
  
  '*** Set the First Key Control to Head Code
  '*** Set the First Non Key Control to Head Name
  '*** Set the current tab as 0

  Set moCn = ADC.Connection    '*** (Bef speed) bef 24/11/06
  Set ADC.FirKeyCtl = ADC("HCoCd")
  Set ADC.FirNKeyCtl = ADC("HName")
  TabHd.Tab = 0
  
  ' ***** Manali 3.5.0 - 09/06/09 - Refresh Pic File List
  mb_DsgCat = moCn.RecSeek("Select 'x' from Head Where HCd='" + ctSelfCmCd + "' And HCoCd='" + ctSelfCoCd + "' and HDsgCatYn='Y' ")
  If mb_DsgCat = True Then CmdRefFileLst.Visible = True Else CmdRefFileLst.Visible = False
  ' ***** Manali 3.5.0 - 09/06/09 - Refresh Pic File List
  ADC("HSTKRTNODSP").Enabled = False
  ADC("HSTKRTNOEDIT").Enabled = False
End Sub
Private Sub ADC_KeyWhen()
   '*** This code is no longer required since all the Key Controls
      '(Controls on the Key Frame) are automatically initialised to blank or  Zero when
      'the Key Frame gets the focus
  ADC("HCd") = ctSelfCmCd
  'ADC("HCoCd") = gs_CoCd
  
  ' ***** Manali 3.5.0 - 09/06/09
  Call EnaDisaCmds(True)
  Call HideAllFras
  'stellar.1- IdName of text boxes to be check quotes
  ms_ChkQtsIdName = ADC("Hname").IdName + "|" + ADC("HBkName").IdName + "|" + ADC("HBkAdd1").IdName + "|" + ADC("HBkAdd2").IdName + "|" + ADC("HBkAdd3").IdName + "|" + ADC("HBkAc").IdName + "|" + _
        ADC("HFrgnBkName").IdName + "|" + ADC("HFrgnBkAc").IdName + "|" + ADC("HInsBy").IdName + "|" + ADC("HInsNo").IdName + "|" + ADC("HAgName").IdName + "|" + ADC("HAgAdd1").IdName + "|" + _
        ADC("HAgAdd2").IdName + "|" + ADC("HAgAdd3").IdName + "|" + ADC("HShpBillSgnBy").IdName + "|" + ADC("HRbiNo").IdName + "|" + ADC("HIecNo").IdName + "|" + ADC("HPreCarriage").IdName + "|" + _
        ADC("HPreCarrier").IdName + "|" + ADC("Hvessel").IdName + "|" + ADC("HPortofLd").IdName + "|" + ADC("HConOfOrg").IdName + "|" + ADC("HMarks").IdName + "|" + ADC("HPkgDesc").IdName + "|" + _
        ADC("HItcNo").IdName + "|" + ADC("HPrdDesc").IdName + "|" + ADC("HConsignee").IdName + "|" + ADC("HUnitCd").IdName + "|" + ADC("HPoL").IdName + "|" + ADC("HCustomHouseCd").IdName + "|" + _
        ADC("HCHALicenseCd").IdName + "|" + ADC("HTypOfExporter").IdName + "|" + ADC("HSbUom").IdName + "|" + ADC("hStatOfOrg").IdName + "|" + ADC("hMetalBOETxt").IdName + "|" + ADC("HRbiPolicy").IdName + "|" + _
        ADC("HAuthSgn").IdName + "|" + ADC("HExpOff").IdName + "|" + ADC("HEtcTxt").IdName + "|" + ADC("HEtcZone").IdName + "|" + ADC("HEtcSection").IdName + "|" + ADC("HETCSvg").IdName + "|" + _
        ADC("HBinNo").IdName + "|" + ADC("HCERegnNo").IdName + "|" + ADC("HMMTC").IdName + "|" + ADC("HTotAs").IdName + "|" + ADC("HAR4No").IdName + "|" + ADC("HQCertNo").IdName + "|" + _
        ADC("HGSPOrgCrit").IdName + "|" + ADC("HLicNo").IdName + "|" + ADC("hImpValTxt").IdName + "|" + ADC("hStatCd").IdName + "|" + ADC("HFtRem").IdName + "|" + ADC("HBnkInstr").IdName + "|" + _
        ADC("HValAddnFt").IdName + "|" + ADC("HUnResolutions").IdName + "|" + ADC("HCustRem").IdName + "|" + ADC("HHSCODE").IdName + "|" + ADC("hGjepcNo").IdName + "|" + ADC("HGjepcValidUpTo").IdName + "|" + _
        ADC("hEximPolicy").IdName + "|" + ADC("hGspEligibility").IdName + "|" + ADC("HGspPkgDesc").IdName + "|" + ADC("hInRegNo").IdName + "|" + ADC("hInCircularNo").IdName + "|" + ADC("hInGspAdd1").IdName + "|" + _
        ADC("hInGspAdd2").IdName + "|" + ADC("hInGspAdd3").IdName + "|" + ADC("HInLclInvTyp").IdName + "|" + ADC("HVATNO").IdName + "|" + ADC("HCSTNO").IdName + "|" + ADC("HTinNo").IdName + "|" + _
        ADC("HPAN").IdName + "|" + ADC("HLclInvFT").IdName + "|" + ADC("HLclInvVATFT").IdName + "|" + ADC("HLclInvCSTFT").IdName + "|" + ADC("HInInvTermCond").IdName + "|" + ADC("HInMemoTermCond").IdName + "|" + _
        ADC("HInAmtDesc1").IdName + "|" + ADC("HInAmtDesc2").IdName + "|" + ADC("HInAmtDesc3").IdName + "|" + ADC("HInAmtDesc4").IdName + "|" + ADC("HInAmtDesc5").IdName + "|" + ADC("HInAmtDesc6").IdName + "|" + _
        ADC("HInAmtDesc7").IdName + "|" + ADC("HInAmtDesc8").IdName + "|" + ADC("HInAmtDesc9").IdName + "|" + ADC("HInAmtDesc10").IdName + "|" + ADC("HInAmtDesc11").IdName + "|" + ADC("HInAmtDesc12").IdName
  
End Sub
Private Sub ADC_NonKeyWhen(Cancel As Boolean, ErrMsg As String)
  '*** Disable all the tabs except the current tab
  Dim i As Single, wb_Ena As Boolean
  For i = 0 To TabHd.Tabs - 1
    FraTabHd(i).Enabled = False
  Next i
  
  FraTabHd(TabHd.Tab).Enabled = True
  
  Call EnaDisaCmds(False)       ' ***** Manali 3.5.0 - 09/06/09
  
  'urmi 2.04-3
  wb_Ena = IIF(ADC("HCoCd") <> "ZZZ", False, True)
  ADC("hCmFixPrc").Enabled = wb_Ena:
  ADC("hCmGldAs").Enabled = wb_Ena
  '**********Bhavna
  ADC("hChnSubCtgForGld").Enabled = wb_Ena
  ADC("hChnSubCtgForLab").Enabled = wb_Ena
  '**********Bhavna
  ADC("hCmIWtEqOrd").Enabled = wb_Ena:
  ADC("hCmIWtFrOrd").Enabled = wb_Ena
  ADC("hCmIGldRtEqOrd").Enabled = wb_Ena
  ADC("HInInvGrsWtEqActual").Enabled = wb_Ena      '*** (Jen 3.1.0_1)
  ADC("HValidDsgSetCd").Enabled = wb_Ena         '*** Jay 3.1.0_1 [Dsg SetCd From Param]
  
  ADC("HChkVldRmLnYN").Enabled = wb_Ena      '****** Sachin 3.3.0 (25-06-08) - [12. Excel Import From Excel] ******
  ADC("HCOSTFRPDCMYN").Enabled = wb_Ena 'CostRtPDCM-Copy Cost Rate YN only in ZZZ
  ADC("hTolWtPctDC").Enabled = wb_Ena           '4.4.0
  ADC("hTolQtyPctDC").Enabled = wb_Ena           '4.4.0
  ADC("hOLORYN").Enabled = wb_Ena
  ADC("HRMSTKBY").Enabled = wb_Ena
  ADC("HORLOTYN").Enabled = wb_Ena
  CmdRefFileLst.Enabled = wb_Ena And mb_DsgCat     ' ***** Manali 3.5.0 - 09/06/09 - Refresh Pic File List
  'Indigo.160 - HOrdRtRndOffRmCtg, hCmCdHlp, HHolChkYN - fields are editable only for "ZZZ"
  ADC("HCmCdHlp").Enabled = wb_Ena
  ADC("HOrdRtRndOffRmCtg").Enabled = wb_Ena
  ADC("HHolChkYN").Enabled = wb_Ena
  
  '*** (Bef 2.14)
  'ADC("hCmDiaRtFA").Enabled = wb_Ena: ADC("hCmCsRtFA").Enabled = wb_Ena
  '*** (Bef 2.14)
  
  '****** Sachin 2.13.0 - 10-05-2006
  ADC("hCmLabAs").Enabled = wb_Ena
  '****** Sachin 2.13.0 - 10-05-2006
  
  'Zubin 212 (25/02/06)
  'ADC("hPPCYN").Enabled = wb_Ena
  ADC("hPPCYN").Enabled = False
  
  ADC("HPDCSfx").Enabled = wb_Ena: ADC("HRndOrdSalPrc").Enabled = wb_Ena
  ADC("hPtrYn").Enabled = wb_Ena
  ' Zubin 212 (25/02/06)
  'ADC("HCtbMod").Enabled = wb_Ena
  ADC("HCtbMod").Enabled = False
  
  ADC("HShowPrdRmDets").Enabled = wb_Ena
  'urmi 28/02/04
  ADC("HMMTc").Enabled = Not wb_Ena
  
  ADC("HMaxDsgViewRec").Enabled = wb_Ena      ' ****** Manali 3.5.0 - Max Limit for Dsg View Records
  
  
  ' ****** Sachin - 2.12 - SJM
   ADC("HSlvModYN").Enabled = False
  ' ****** Sachin - 2.12 - SJM
  
  ' ****** Manali - 3.03 - 16/06/08 - Palladium Module
   ADC("HPdModYN").Enabled = False
  ' ****** Manali - 3.03 16/06/08 - Palladium Module
  
  '****** Sachin 2.13.0 - 05-06-2006 ******
  ADC("HQWModYN").Enabled = False
  If ADC("HCoCd") <> ctSelfCoCd Then ADC("HModAvlYN").Enabled = False Else ADC("HModAvlYN").Enabled = True
  '****** Sachin 2.13.0 - 05-06-2006 ******

  '****** Sachin 2.14.0 - 26-10-2006 - [Multi-Company] ******
  ADC("HImpComLoc").Enabled = False
  ADC("HINREPMETLSRMCD").Enabled = False
  '****** Sachin 2.14.0 - 26-10-2006 - [Multi-Company] ******

  '****** Sachin 3.3.0 (30-06-08) - [Copy From Excel Module]******
  ADC("HEXCELIMPINORDYN").Enabled = False
  '****** Sachin 3.3.0 (30-06-08) - [Copy From Excel Module]******
  ADC("HTravelExlYN").Enabled = False       '****** Sachin 3.3.0 (19-07-08) - [Travel Excel Added]
  ADC("HOrdExcelRepYn").Enabled = False   ' ***** Manali 3.6.0 - Order Printing in Excel Format Module
  'ADC("HCHCDHCOnCustRm").Enabled = False
  ' ****** Manali 3.4.1 - Design Catalogue Module
  ADC("HDsgCatYN").Enabled = False
  ' ****** Manali 3.4.1 - Design Catalogue Module

  ' **** Zubin 211 **** '
  ' Zubin 212 (25/02/06)
  'ADC("hFluteBagYN").Enabled = wb_Ena
  ADC("hFluteBagYN").Enabled = False
  If ADC("HCoCd") <> ctSelfCoCd Then ADC("HDTPrdPtr").Enabled = False Else ADC("HDTPrdPtr").Enabled = True
  ' **** Zubin 211 **** '
  
  ' Zubin 212 (25/02/06)
  ADC("hFgQuotYN").Enabled = False
  ADC("hLotNoYN").Enabled = False
  
  ' Zubin 213
  ADC("HMultiCurYN").Enabled = False
  If moCn.RecSeek("Select 'x' from Head where HCd= 'ZSELF' and HCoCd= '" + ADC("HCoCd") + "' and " + _
      "HMultiCurYN= 'N'") Then
    ADC("HBaseCurCd").Enabled = False
  Else
    ADC("HBaseCurCd").Enabled = True
  End If
  
  '*** (Jen 2.12 Patch 3)
  ADC("hLooseInv").Enabled = False
  '*** (Jen 2.12 Patch 3)
  
  '*** (Jen 2.13)
  ADC("hConsolidatedInv").Enabled = False
  ADC("HAutomateReturns").Enabled = False
  '*** (Jen 2.13)
 '*** JAY 2.13(ct) ***
  ADC("HCOMPBAGSYN").Enabled = False
 '*** JAY 2.13(ct) ***
  '*******************

  'If ADC("HCoCd") = ctSelfCoCd Then
  '  ADC("HShowPrdRmDets").Visible = True
  '  ADC("HShowPrdRmDets").Enabled = True
  '  lblShowBothWts.Visible = True
  'Else
  '  ADC("HShowPrdRmDets").Visible = False
  '  ADC("HShowPrdRmDets").Enabled = False
  '  lblShowBothWts.Visible = False
  'End If
  
  '*** (Jen 2.14)
  If ADC.AddFndMode = xaddmode Then
    ADC("HCmDiaRtFA") = "A"
    ADC("HCmCsRtFA") = "A"
  End If
  '*** (Jen 2.14)
  
  ADC("HLclInvYn").Enabled = False    '*** JAY 2.14(LclInv) ***
  ADC("HFgTagMod").Enabled = False    '*** Jay 3.01 *** [FgTag Module]
  ADC("HFGIOYN").Enabled = False      '*** Jay 3.3(FG)
  
  ADC("HMultiPrcQtMod").Enabled = False     ' **** Manali 3.10.0 - MultiPrcQt Module
  
  
  '*** (Jen 3.01)
  ADC("HCustomiseOrd").Enabled = False
  ADC("HAutoCustmMtch").Enabled = False
  If ADC("HAutoCustmMtch") = "Y" Then
    TabHd.TabVisible(5) = True
    TabHd.TabsPerRow = 7
  Else
    TabHd.TabVisible(5) = False
    TabHd.TabsPerRow = 6
  End If
  '*** (Jen 3.01)
  
  ADC("hBestUse").Enabled = False   '*** (Jen 3.1.2)
  
  ' **** Manali 3.6.0 - 30/10/09 - For DTA Use Notional Rate in Custom Matching
    'Sachin - 4.1.0.0 - Notional Rate not related to DTA anymore
    '  If adc.AddFndMode = xFndMode Then
    '    If adc("hExpLocal") <> "Y" Then
    '      LblNotionRt.Visible = False: adc("hInNotionRtYn").Enabled = False: adc("hInNotionRtYn").Visible = False
    '      adc("hInNotionRtYn") = "N"
    '    Else
    '      LblNotionRt.Visible = True: adc("hInNotionRtYn").Enabled = True: adc("hInNotionRtYn").Visible = True
    '    End If
    '  End If
  ' **** Manali 3.6.0 - 30/10/09 - For DTA Use Notional Rate in Custom Matching
  
  ADC("hRtfrRtChrtYN").Enabled = False      '***** Sachin 3.11.0 - Inv Rt from Rt Chart
  '*************Bhavna added default for inv amt desc 7 and 12
  ADC("HInAmtDesc7") = "Net Realization"
  ADC("HInAmtDesc7").Enabled = False
  ADC("HInAmtDesc12") = "Total Amount"
  ADC("HInAmtDesc12").Enabled = False
  '*************Bhavna added default for inv amt desc 7 and 12
End Sub
Private Sub ADC_NonKeyValid(pr_Cancel As Boolean, pr_ErrMsg As String)
    Dim ws_RmCtgArr() As String, ws_PrcsSeqArr() As String, j As Integer, ws_PrdSeqArr() As String
    Dim i As Integer
    '3.11.0
    TabHd.Tab = 0
    If ADC("hDmPrdSeq") = "" Then
        pr_Cancel = True
        pr_ErrMsg = "Prd Seq cannot be blank"
        Exit Sub
    End If
    If ADC("hDmPrdSeq") <> "" Then
      ws_PrcsSeqArr = Split(ADC("hDmPrcsSeq"), ",")
      ws_PrdSeqArr = Split(ADC("hDmPrdSeq"), ",")
      
      If (UBound(ws_PrdSeqArr) <> UBound(ws_PrcsSeqArr)) Or (UBound(ws_PrcsSeqArr) > 0) Or (UBound(ws_PrdSeqArr) > 0) Then
        pr_Cancel = True
        pr_ErrMsg = "No of values in Default PrcsSeq must match that in PrdSeq and there can only be one value in PrcsSeq and PrdSeq"
        Exit Sub
      End If
    
      For i = 0 To UBound(ws_PrdSeqArr)
      If IsArray(ws_PrdSeqArr) Then
         pr_Cancel = (ADC("hDmPrcsSeq") <> "") And (Not moCn.RecSeek("select vPMCd from vParam where " + _
                                          "vPCoCd = '" + ADC("HCoCd") + "' and vPTyp= 'PRDSEQ' and " + _
                                          "vPMCd = '" + ws_PrdSeqArr(i) + "' and vPValue1='" + ws_PrcsSeqArr(i) + "'"))
         If pr_Cancel = True Then pr_ErrMsg = "Validation failed for PrdSeq, Value(" + CStr(i + 1) + ") '" + ws_PrdSeqArr(i) + " ' is invalid for given Process Sequence": Exit Sub
      End If
      Next i
    End If
    'stellar.1- calling a function to check Quotes exists in the list of text boxes
    Dim wChkQtsErrMsg As String
    wChkQtsErrMsg = ChkQuotes(Me, ms_ChkQtsIdName, "")
    If wChkQtsErrMsg <> "" Then pr_Cancel = True: pr_ErrMsg = wChkQtsErrMsg
  
End Sub


Private Sub ATXT_FldWhen(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String)
  '*** Set Helps For Each Field
  Select Case IdName
  Case Is = UCase("HCoCd")
    Call HlpList.hCoCd
  Case Is = UCase("HCd")
    Call HlpList.HCd(ADC("HCoCd"))
  Case Is = UCase("HCmGldAs")
    Call HlpList.PMCd("GLDAS")
  
  '****** Sachin - 2.13.0 - 10-05-2006
  Case Is = UCase("HCmLabAs")
    Call HlpList.PMCd("GLDAS")
  
  '****** Bhavna 5.1.0
  Case Is = UCase("hChnSubCtgForGld")
    HlpList.MultiSelect = True
    Call HlpList.PSCd("RMSCTG", "", True)
  
  Case Is = UCase("hChnSubCtgForLab")
    HlpList.MultiSelect = True
    Call HlpList.PSCd("RMSCTG", "", True)
  '****** Bhavna
  
  Case Is = UCase("HCmIWtEqOrd")
    Call HlpList.PMCd("YN")
  Case Is = UCase("HCHCDHCOnCustRm")
    If ADC("HCHCDHCOnCustRm") = "" Then ADC("HCHCDHCOnCustRm") = "N"
    Call HlpList.PMCd("YN")

  '*** When HCmIWtEqOrd='y' then HCmIWtFrOrd is made blank in recalc,
      'and read only
  Case Is = UCase("HCmIWtFrOrd")
    If ADC("HCmIWtEqOrd") = "Y" Then Cancel = True: ErrMsg = "Cannot Edit This Value": Exit Sub
    Call HlpList.PMCd("IWTFRORD")
  Case Is = UCase("HCmIGldRtEqOrd")
    Call HlpList.PMCd("YN")
  Case Is = UCase("HCmDiaRtFA")
    Call HlpList.PMCd("FA")
  Case Is = UCase("HCmCSRtFA")
    Call HlpList.PMCd("FA")
  Case Is = UCase("HCmFixPrc")
    Call HlpList.PMCd("YN")
  'uma *** add R process
  Case Is = UCase("hDmPrdSeq")
    HlpList.MultiSelect = True  'EmrSv - Uma
    Call HlpList.vPMCd(ADC("HCoCd"), "PRDSEQ")
  Case Is = UCase("hDmPrcsSeq")
    HlpList.MultiSelect = True  'EmrSv - Uma
    Call HlpList.vPMCd(ADC("HCoCd"), "PRCSSEQ")
  Case Is = UCase("hDmVaCtg")
    Call HlpList.PMCd("VACTG")
  Case Is = UCase("hOmPrtCd")
    Call HlpList.PMCd("PRTCD")
  Case Is = UCase("hBOpnLoc")
    Call HlpList.Loc(ADC("HCoCd"), "'P'")
  ' ***** Manali 3.5.0 - Loaction for applying Minimum tolerance On Diamond
  Case Is = UCase("HMinDiaProcess"), UCase("hChkWaxPullOutLocs"), UCase("hChkCastLocs") '4.1.3.0
    HlpList.MultiSelect = True ' ***** Manali 3.5.0 - 08/06/09 - Multiple Loc Allowed
    Call HlpList.Loc(ADC("HCoCd"), "'R'")
  ' ***** Manali 3.5.0 - Loaction for applying Minimum tolerance On Diamond
  Case Is = UCase("hBChr")
    Call HlpList.PMCd("BCHR")
  Case Is = UCase("hTdAck")
    Call HlpList.PMCd("YN")
  Case Is = UCase("hFgPlChr")
    Call HlpList.PMCd("PLCHR")
  Case Is = UCase("hYy")
    Call HlpList.PMCd("YY")
  Case Is = UCase("hZWtEq")
    Call HlpList.PMCd("ZWTEQ")
  Case Is = UCase("hOrdDsgGr")
    Call HlpList.PMCd("ORDDSGGR")
  Case Is = UCase("hOrdRmGr")
    Call HlpList.PMCd("ORDRMGR")
  Case Is = UCase("hOrdLabGr")
    Call HlpList.PMCd("ORDLABGR")
  Case Is = UCase("hOthLocBag")
    Call HlpList.PMCd("YN")
  Case Is = UCase("hOdAmdQty")
    Call HlpList.PMCd("YN")
  Case Is = UCase("HTotAs")
    Call HlpList.PMCd("INVTOTAS")
  Case Is = UCase("hZValEqInv")
    Call HlpList.PMCd("YN")
  Case Is = UCase("HCmCdHlp")
    Call HlpList.PMCd("YN")
  Case Is = UCase("HRndOffFOB")
    Call HlpList.PMCd("YN")
  Case Is = UCase("HShowActWt")
    Call HlpList.PMCd("YN")
  Case Is = UCase("HSrtBasePurVch")
    Call HlpList.PMCd("SRTPRHLP")
  Case Is = UCase("hInRepWtValAs")
    Call HlpList.PMCd("INWT")
  Case Is = UCase("hInRepWtFld")
    Call HlpList.PMCd("YN")
  Case Is = UCase("hInRepMetLs")
    Call HlpList.PMCd("YN")
  Case Is = UCase("hInRepCnvRt")
    Call HlpList.PMCd("YN")
  Case Is = UCase("hInRepRtFld")
    Call HlpList.PMCd("YN")
  Case Is = UCase("hPtrYN")
    Call HlpList.PMCd("YN")
  Case Is = UCase("hPdcSfx")
    Call HlpList.PMCd("PDCSFX")
  Case Is = UCase("HRndOrdSalPrc")
    Call HlpList.PMCd("RNDOPT")
  Case Is = UCase("HOrdCpyRtOpt")
    Call HlpList.PMCd("CPYRT")
  Case Is = UCase("HShowPrdRmDets")
    Call HlpList.PMCd("YN")
  Case Is = UCase("hGrsWtInBsYN")
    Call HlpList.PMCd("YN")
  Case Is = UCase("hExpLocal"), UCase("hGenExpNoYN")
    Call HlpList.PMCd("YN")
  Case Is = UCase("WINVPRDDESC")
    Call HlpList.PMCd("INVDESC")
  Case Is = UCase("hPPCYN")
    Call HlpList.PMCd("YN")
  Case Is = UCase("hLine")
    Call HlpList.vPMCd(ADC("HCoCd"), "LINE")
  Case Is = UCase("hWH")
    Call HlpList.PMCd("WH")
  Case Is = UCase("HInvTyp")
    Call HlpList.PMCd("InvTyp")
  Case Is = UCase("hValAddMtlColYN"), UCase("hValBOEDetYN"), UCase("hBagNoEditYN"), UCase("HNegStkYn"), _
            UCase("HAvgSRtYn"), UCase("hSetTypOnlyFrOrdYN")     '3.11.0
    Call HlpList.PMCd("YN")
  Case Is = UCase("HFgSubLoc")
    Call HlpList.vPMCd(ADC("HCoCd"), "FGSUBLOC")
  'EmrSv - Uma
  Case Is = UCase("hCell")
    Call HlpList.PMCd("Cell")
    
  '****** Sachin - 2.14.0 - [07-10-2006]
  Case Is = UCase("hNegStkRmCtg")
    If ADC("HNegStkYn") = "Y" Then Cancel = True: ErrMsg = "Cannot Edit If Negative Stock Is Allowed": Exit Sub
    Call HlpList.PMCd("GLDAS")
  '****** Sachin - 2.14.0 - [07-10-2006]
    
  Case Is = UCase("HSizeStkRt"), UCase("hPrdPtsYN")
    'If ADC("HAvgSRtYn") = "Y" Then Cancel = True: DispMsg "Cannot Alter when Avg Stk Rt Option is Y", etError
    Call HlpList.PMCd("YN")
  Case Is = UCase("HCtbMod")
    Call HlpList.PMCd("YN")
  Case Is = UCase("HCtCd")
    If ADC("HCtbMod") = "N" Then Cancel = True: ErrMsg = "Contractor Module Option Should Be 'Y' First": Exit Sub
    Call HlpList.CustCd("T")
  '*************************** ZUBIN **************************
  ' 10th Dec 2003, EMR206
  Case Is = UCase("HPrtAllotOnPrdYN")
    Call HlpList.PMCd("YN")
  '*************************** ZUBIN **************************
  '*************************** Geeta**************************
  ' 15th Jan 2004, EMR206
    Case Is = UCase("HChkDCWtAtIssYN"), UCase("HChkDCQtyAtIssYN"), UCase("HChkGPWtAtIssYN"), UCase("hFgRmEqOrdRm")
    Call HlpList.PMCd("YN")
    '*************************** Geeta**************************
  
  ' **** Zubin 211 **** '
  Case Is = UCase("hFluteBagYN"), UCase("HOrdPrdPts"), UCase("HDTPrdPtr")
    Call HlpList.PMCd("YN")
    
  '*** 2.11 (Jen)
  Case Is = UCase("HOdGmChk")
    Call HlpList.PMCd("CPYRT")
  '*** 2.11 (Jen)
' **** Zubin 211 **** '
   
  ' **** Zubin 212 **** '
  ' ***** Manali 3.8.0 - Company Logo in Report
  Case Is = UCase("HLotNoYN"), UCase("HFgQuotYN"), UCase("HGldRtDAYN"), UCase("HShowRepLogo"), UCase("hRtfrRtChrtYN")
    Call HlpList.PMCd("YN")
  ' **** Zubin 212 **** '
   
  '****** (Jen 3.1.2) added hBestUse ******
  '****** (Jen 3.01.1 added HInInvGrsWtEqActual) ******
  '****** (Jen 3.01) added HCustomiseOrd, HAutoCustmMtch ******
  '****** JAY 2.14(LclInv) added hLclInvYn ******
  '****** JAY 2.13(CT) added hCompBagsYn ******
  '****** Zubin 213 added hMultiCurYN ******
  '****** (Jen 2.13 added hConsolidatedInv, hAutomateReturns) ******
  '****** (Jen 2.12 Patch 3 added HLooseInv) ******
  '****** Sachin 2.12 - 05/12/05 - SJM added HSlvModYN, HQWModYN, HModAvlYN, hImpComLoc, HINREPMETLSRMCD ******
  '****** Zubin 3.02.01 - 26/03/08 - added hChgPtrOnStwYN, hCmILabWtFrOrd
  '****** Manali 3.03 - 16/06/08 - added HPdModYN
  '****** Sachin 3.3.0 (23-06-08) - [14/12. Invoicing with Markup on Cost of Diamonds/Excel Import In Order] ******
  '****** Manali 3.4.1 - 22/10/08 - Design Catalogue module [added HDsgCatYN]
  ' ******** Manali 3.5.0 - 26/11/08 - Hide Commission field added [HHideCommYN]
  ' ***** Manali 3.6.0 - 27/10/09 - Invoicing with Markup on Cost of Accessories
  ' **** Manali 3.6.0 - 30/10/09 - For DTA Use Notional Rate in Custom Matching
  ' ***** Manali 3.8.1 - Self Payable Options In Invoice added - HSelfDscYN, HSelfDeductYN
  ' **** Manali 3.10.0 - 03/04/12- MultiPrcQt Module
  ' ***** Sachin 3.11.0 - Markup for Metal
  '4.1.0.0
  '4.1.2 - Holiday check
  Case Is = UCase("HSlvModYN"), UCase("HQWModYN"), UCase("HModAvlYN"), UCase("hImpComLoc"), UCase("HINREPMETLSRMCD"), _
            UCase("HLooseInv"), UCase("hConsolidatedInv"), UCase("hAutomateReturns"), UCase("hMultiCurYN"), _
            UCase("hCompBagsYn"), UCase("hLclInvYn"), UCase("HCustomiseOrd"), UCase("HAutoCustmMtch"), _
            UCase("HInInvGrsWtEqActual"), UCase("hBestUse"), UCase("hChgPtrOnStwYN"), UCase("hCmILabWtFrOrd"), _
            UCase("HPdModYN"), UCase("HINIRTWRTSTKYN"), UCase("HEXCELIMPINORDYN"), UCase("HCHKVLDRMLNYN"), UCase("hTravelExlYN"), _
            UCase("HDsgCatYN"), UCase("HHideCommYN"), UCase("HINXIRTWRTSTKYN"), UCase("HMETIRTWRTSTKYN"), UCase("hInNotionRtYn"), _
            UCase("HSelfDscYN"), UCase("HSelfDeductYN"), UCase("HMultiPrcQtMod"), UCase("hMetIRtwrtStkYN"), UCase("hSetTypOnlyFrOrdYN"), _
            UCase("HSHOWBGHISTDETSYN"), UCase("HHOLCHKYN"), UCase("HOLORYN"), UCase("HAddBrkMSGYN"), UCase("HOrLotYN")
    Call HlpList.PMCd("YN")
      
  ' Zubin 213
  Case Is = UCase("hBaseCurCd")
    Call HlpList.PMCd("CURNCY")
  ' Zubin 213
  
  '*** (Jen 3.01)
  'Manali 3.03 - 18/06/08 - Palladium added
  Case Is = UCase("HImpTypNotInSelD"), UCase("HImpTypNotInSelC"), UCase("HImpTypNotInSelG"), _
            UCase("HImpTypNotInSelP"), UCase("HImpTypNotInSelS"), UCase("HImpTypNotInSelL"), UCase("HImpTypNotInSelX")
    HlpList.MultiSelect = True
    Call HlpList.PMCd("IMPTYP")
  '*** (Jen 3.01)
  
  '*** Jay 3.1.0_1 [Dsg SetCd From Param]
  Case Is = UCase("HValidDsgSetCd")
    If ADC("HValidDsgSetCd") = "Y" Then Cancel = True: ErrMsg = "Cannot Edit This Value": Exit Sub
    Call HlpList.PMCd("YN")
  '*** Jay 3.1.0_1 [Dsg SetCd From Param]
  
  '****** Sachin - 3.5.0 - [15/11/08]
  Case Is = UCase("hCustomizeRmCtgExp")
    If ADC("HCustomiseOrd") = "N" Then ADC("hCustomizeRmCtgExp") = "": Cancel = True: ErrMsg = "Cannot Edit If Customise Module Is Disabled, Enable Customise Module First": Exit Sub
    Call HlpList.PMCd("GLDAS")
  '****** Sachin - 3.5.0 - [15/11/08]
  
  ' ***** Manali 3.8.0 - Invoice Design Sort By
  Case Is = UCase("HInSrt")
    Call HlpList.PMCd("INSRT")
  
  ' ***** Manali 3.8.0 - Metler Code
  Case Is = UCase("HMetlrCd")
    Call HlpList.PMCd("METLR")
    
 ' ***** Manali 3.9.1 - 20/02/12 - RfId
  Case Is = UCase("HRfIdMod")
    Call HlpList.PMCd("YN")
    
 ' ***** Manali 3.9.1 - 20/02/12 - Rate Round Off RmCtgs
  Case Is = UCase("HOrdRtRndOffRmCtg")
    Call HlpList.PMCd("CPYRT")
    
  Case Is = UCase("hCpyActCalcWtYN")       '3.11.0
    Call HlpList.PMCd("YN")
    
  'Sachin - 4.1.0.0
  Case Is = UCase("hShowBgHistDets"), UCase("HHOLCHKYN"), UCase("HEnfFrBLocInFgYN")    '4.1.2
    Call HlpList.PMCd("YN")
    
  Case Is = UCase("HPRDPTSONBQTYYN")        '4.3.0.0
    Call HlpList.PMCd("YN")
  Case Is = UCase("HjStkNumjFrBagYN")        '4.3.0.0
    Call HlpList.PMCd("YN")

  Case Is = UCase("HJSTKDEFCHR")        '4.3.0.0
    Call HlpList.PMCd("BCHR", "'P'")
  
  'uma**** repair bag location
  Case Is = UCase("HDefRepBagLoc")
    Call HlpList.Loc(ADC("HCoCd"), "'P'")
  
  'StkRt-if Average Stock Rate is NO then can't edit NoDisplay field else it will change into Y\N
  Case Is = UCase("HSTKRTNODSP")
    If ADC("HAVGSRTYN") = "N" Then Cancel = True: ErrMsg = "Cannot Edit If Average Stock Rate=No": Exit Sub
    Call HlpList.PMCd("YN")
    
  'StkRt-if Average Stock Rate is NO then can't edit NoEdit field else will change into Y\N.
  'If Stock Rate NoDisplay is Yes then NoEdit change into Y & cant chnage the value else can change the value
  Case Is = UCase("HSTKRTNOEDIT")
    If ADC("HAVGSRTYN") = "N" Or ADC("HSTKRTNODSP") = "Y" Then _
        Cancel = True: ErrMsg = "Cannot Edit If Average Stock Rate=No Or Stock Rate NoDisplay=Yes": Exit Sub
    Call HlpList.PMCd("YN")
  'CostRtPDCM-When Sales rate coming from PDCM:'Y'-Cost from PDCM,'N'–Cost from Current Rate charts
  Case Is = UCase("hCostFrPdcmYN")
    Call HlpList.PMCd("YN")
      
  End Select
End Sub

Private Sub atxt_ReCalculate(Index As Integer, ByVal IdName As String, ByVal OldVal As Variant, ByVal RowNum As Integer, ByVal SourceName As String, ByVal SourceColName As String)
  '*** HCmIWtFrOrd on HCmIWtEqOrd         Normal
      'If HCmIWtEqOrd= 'Y' then ''
  Select Case UCase(IdName)
  Case Is = UCase("HCmIWtFrOrd")
    If ADC.Mode = xNorm Then
      If ADC("HCmIWtEqOrd") = "Y" Then ADC("HCmIWtFrOrd") = ""
    End If
  Case Is = UCase("HPRDDESC")
    If ADC.Mode = xNorm Then
      ADC("HPRDDESC") = moCn.GetFldVal("Select PDesc from Param where Ptyp='INVDESC' and PmCd = '" + ADC("WINVPRDDESC") + "'")
    End If
  Case Is = UCase("HCtCd")
    If ADC("HCtbMod") = "N" Then ADC("HCtCd") = ""
  'Case Is = UCase("HSizeStkRt")
    'If ADC.Mode = xNorm And ADC("HAvgSRtYn") = "Y" Then ADC("HSizeStkRt") = "Y"
  
  ' Manoj 2.9.0 Patch
  Case Is = UCase("HStkRtPassReq")
    Call HlpList.PMCd("YN")
  ' Manoj 2.9.0 Patch
'******Bhavna
  Case Is = UCase("HChkOrRtYN")
    Call HlpList.PMCd("YN")
  Case Is = UCase("HRndOffOdTolWtYN")
    Call HlpList.PMCd("YN")
'******Bhavna
  '****** Sachin 3.5.0 [15/11/08]
  Case Is = UCase("HCustomizeRmCtgExp")
    If ADC.Mode = xNorm Then _
      If ADC("HCustomiseOrd") = "N" Then ADC("HCustomizeRmCtgExp") = ""
  '****** Sachin 3.5.0 [15/11/08]
  'StkRt-IF Average stock rate is No then NoEdit & NoDisplay will change to N
  'IF Average stock rate is Yes then user can change NoEdit & NoDisplay to Yes or No
  Case Is = UCase("HSTKRTNODSP")
    If ADC.Mode = xNorm Then If ADC("HAVGSRTYN") = "N" Then ADC("HSTKRTNODSP") = "N"
  Case Is = UCase("HSTKRTNOEDIT")
    If ADC.Mode = xNorm Then If ADC("HAVGSRTYN") = "N" Then ADC("HSTKRTNOEDIT") = "N"
    'StkRt-IF NoDisplay is Yes then NoEdit will change into Y else user wish
    If ADC.Mode = xNorm Then If ADC("HSTKRTNODSP") = "Y" Then ADC("HSTKRTNOEDIT") = "Y"

  End Select
End Sub

Private Sub atxt_FldChange(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String, ByVal pv_NewValue As Variant)
 '*** HCd              Not Blank
  '*** HCmGldAs         valid PMCd('GLDAS') or ''
  '*** HCmLabAs         valid PMCd('GLDAS') or ''
  '*** HCmIWtEqOrd      valid PMCd('YN') or ''
  '*** HCmIWtFrOrd      valid PMCd('IWTFRORD') or ''
  '*** HCmIGldRtEqOrd   valid PMCd('YN') or ''
  '*** HCmDiaRtFA       valid PMCd('FA') or ''
  '*** HCmCSRtFA        valid PMCd('FA') or ''
  '*** HCmFixPrc        valid PMCd('YN') or ''
  '*** HDmPrdSeq        valid PMCd('PRDSEQ') or ''
  '*** hDmVaCtg         valid PMCd('VACTG') or ''
  '*** hOmPrtCd         valid PMCd('PRTCD') or ''
  '*** hBOpnLoc         valid LocCd from Loc(LocTyp= 'P') or ''
  '*** hBChr            valid PMCd('BCHR') or ''
  '*** hTdAck           valid PMCd('YN') or ''
  '*** hFgPlChr         valid PMCd('PLCHR') or ''
  '*** hYy              valid PMCd('YY') or ''
  '*** hZWtEq           valid PMCd('ZWTEQ') or ''
  '*** hOrdDsgGr        valid PMCd('ORDDSGGR') or ''
  '*** hOrdRmGr         valid PMCd('ORDRMGR') or ''
  '*** hOrdLabGr        valid PMCd('ORDLABGR') or ''
  '*** hOthLocBag       valid PMCd('YN') or ''
  '*** hOdAmdQty        valid PMCd('YN') or ''
  '*** HTotAs           valid PMCd('INVTOTAS') or ''
  '*** hZValEqInv       valid PMCd('YN') or ''
  '*** HCmCdHlp         valid PMCd('YN') or ''
  '*** HRndOffFOB       valid PMCd('YN') or ''
  '*** HShowActWt       valid PMCd('YN') or ''
  '*** HSrtBasePurVch   valid PMCd('SRTPRHLP') or ''
  '*** hInRepWtValAs    valid PMCd('INWT') or ''
  '*** hInRepWtFld      valid PMCd('YN') or ''
  '*** hInRepMetLs      valid PMCd('YN') or ''
  '*** hInRepCnvRt      valid PMCd('YN') or ''
  '*** hInRepRtFld      valid PMCd('YN') or ''
  '*** hPtrYN           valid PMCd('YN') or ''
  
  Dim ws_RmCtgArr() As String, ws_PrcsSeqArr() As String, j As Integer, ws_PrdSeqArr() As String    '3.11.0
  'stellar.1- calling a function to check Quotes exists in given value
  Dim wChkQtsErrMsg As String
  If InStr(1, "|" + ms_ChkQtsIdName + "|", "|" + IdName + "|") > 0 Then
    wChkQtsErrMsg = ChkQuotes(Me, IdName, pv_NewValue)
    If wChkQtsErrMsg <> "" Then Cancel = True: ErrMsg = wChkQtsErrMsg: Exit Sub
  End If
  
  Select Case UCase(IdName)
  Case Is = UCase("HCoCd")
    If Trim(pv_NewValue) = "" Then
      Cancel = True: ErrMsg = "Head Code Cannot Be Blank": Exit Sub
    End If
  Case Is = UCase("HCd")
    If Trim(pv_NewValue) = "" Then
      Cancel = True: ErrMsg = "Head Code Cannot Be Blank": Exit Sub
    End If
  Case Is = UCase("HCmGldAs")
    '*** (Bef 2.14)
    'Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("select PMCd from Param where " + _
    '                                     "PTyp= 'GLDAS' and PMCd = '" + pv_NewValue + "'"))
    '*** (Bef 2.14)
    '*** (Jen 2.14)
    'Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("select PMCd from Param where PTyp= 'GLDAS' " + _
    '         "and PMCd = '" + pv_NewValue + "' and not (',' + PDesc225 + ',' like '%,G,%' Or ',' + PDesc225 + ',' like '%,P,%' Or ',' + PDesc225 + ',' like '%,S,%') "))
    '
    Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("select PMCd from Param where PTyp= 'GLDAS' " + _
             "and PMCd = '" + pv_NewValue + "' "))
    '*** (Jen 2.14)
    ErrMsg = "Invalid Gold As Value": Exit Sub
  
  '****** Sachin - 2.13.0 - 10-05-2006
  Case Is = UCase("HCmLabAs")
    Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("select PMCd from Param where " + _
                                         "PTyp= 'GLDAS' and PMCd = '" + pv_NewValue + "'"))
    ErrMsg = "Invalid Labour As Value": Exit Sub
  '****** Sachin - 2.13.0 - 10-05-2006
  
   '***********Bhavna 5.1.0
   Case Is = UCase("hChnSubCtgForGld"), UCase("hChnSubCtgForLab")
    Dim ws_GldAsArr() As String, i As Integer
    ws_GldAsArr = Split(pv_NewValue, ",")
    If IsArray(ws_GldAsArr) Then
      For i = 0 To UBound(ws_GldAsArr)
          Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("select PSCd from Param where PTyp= 'RMSCTG' " + _
                    "and PSCd = '" + ws_GldAsArr(i) + "' "))
          If Cancel = True And UCase(IdName) = UCase("hChnSubCtgForGld") Then ErrMsg = "Value(" + CStr(i + 1) + ") Is An Invalid Chain Gold As Value": Exit Sub
          If Cancel = True And UCase(IdName) = UCase("hChnSubCtgForLab") Then ErrMsg = "Value(" + CStr(i + 1) + ") Is An Invalid Chain Labour As Value": Exit Sub
      Next i
    End If
  
  Case Is = UCase("HChkOrRtYN")
    Cancel = Not moCn.RecSeek("select PMCd from Param where " + _
                                            "PTyp= 'YN' and PMCd = '" + pv_NewValue + "'")
    ErrMsg = "Check Zero in Order Rm Rate (Y/N)": Exit Sub
  Case Is = UCase("HOLORYN")
    Cancel = Not moCn.RecSeek("select PMCd from Param where " + _
                                            "PTyp= 'YN' and PMCd = '" + pv_NewValue + "'")
    ErrMsg = "Enter Yes or No": Exit Sub
  Case Is = UCase("HAddBrkMsgYN")
    Cancel = Not moCn.RecSeek("select PMCd from Param where " + _
                                            "PTyp= 'YN' and PMCd = '" + pv_NewValue + "'")
    ErrMsg = "Enter Yes or No": Exit Sub
  
  Case Is = UCase("HRndOffOdTolWtYN")
    Cancel = Not moCn.RecSeek("select PMCd from Param where " + _
                                            "PTyp= 'YN' and PMCd = '" + pv_NewValue + "'")
    ErrMsg = "Diamond Min / Max wt Tolerance to be rounded off to 2 digits (Y/N)": Exit Sub
  
  '****** Bhavna

  
  Case Is = UCase("HCmIWtEqOrd")
    Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("select PMCd from Param where " + _
                                            "PTyp= 'YN' and PMCd = '" + pv_NewValue + "'"))
    ErrMsg = "Invoice Wts From Order (Y/N)": Exit Sub
  Case Is = UCase("HCmIWtFrOrd")
    Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("select PMCd from Param where " + _
                                                  "PTyp= 'IWTFRORD' and " + _
                                                  "PMCd = '" + pv_NewValue + "'"))
    ErrMsg = "Invalid Wts From Order": Exit Sub
  Case Is = UCase("HCmIGldRtEqOrd")
    Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("select PMCd from Param where " + _
                                                  "PTyp= 'YN' and PMCd = '" + pv_NewValue + "'"))
    ErrMsg = "Invoice Gold Rate From Order (Y/N)": Exit Sub
  Case Is = UCase("HCmDiaRtFA")
    '*** (Bef 2.14)
    'Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("select PMCd from Param where " + _
    '                                              "PTyp= 'FA' and PMCd = '" + pv_NewValue + "'"))
    '*** (Bef 2.14)
    
    Cancel = pv_NewValue <> "A"   '*** (Jen 2.14)
    
    ErrMsg = "Enter Diamond Rate Favourable/Actual": Exit Sub
  Case Is = UCase("HCmCSRtFA")
    '*** (Bef 2.14)
    'Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("select PMCd from Param where " + _
    '                                       "PTyp= 'FA' and PMCd = '" + pv_NewValue + "'"))
    '*** (Bef 2.14)
    
    Cancel = pv_NewValue <> "A"   '*** (Jen 2.14)
    ErrMsg = "Enter Colour Stone Rate Favourable/Actual": Exit Sub
  Case Is = UCase("HCmFixPrc")
    Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("select PMCd from Param where " + _
                                          "PTyp= 'YN' and PMCd = '" + pv_NewValue + "'"))
    ErrMsg = "Fixed Price (Y/N)": Exit Sub
  'uma *** add R process
  Case Is = UCase("hDmPrdSeq")
    Cancel = pv_NewValue = ""
    If Cancel = True Then ErrMsg = "Production Sequence Cannot Be Blank": Exit Sub

    'If pv_NewValue <> "" Then
     ' ws_PrdSeqArr = Split(pv_NewValue, ",")
      'ws_PrcsSeqArr = Split(ADC("hDmPrcsSeq"), ",")
      'If UBound(ws_PrdSeqArr) <> UBound(ws_PrcsSeqArr) Then
'        Cancel = True
 '       ErrMsg = "No of values in Default PrcsSeq must match that in PrdSeq"
  '      Exit Sub
   '   End If
    '  For i = 0 To UBound(ws_PrdSeqArr)
     '     If IsArray(ws_PrdSeqArr) Then
      '       Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("select vPMCd from vParam where " + _
       '                                     "vPCoCd = '" + ADC("HCoCd") + "' and vPTyp= 'PRDSEQ' and " + _
        '                                    "vPMCd = '" + ws_PrdSeqArr(i) + "' and vPValue1='" + ws_PrcsSeqArr(i) + "'"))
         '    If Cancel = True Then ErrMsg = "Value(" + CStr(i + 1) + ") Is An Invalid Production Sequence or Invalid for given Process Sequence": Exit Sub
          'End If
'      Next i
 '   End If
  Case Is = UCase("hDmPrcsSeq")
    Cancel = pv_NewValue = ""
    If Cancel = True Then ErrMsg = "Process Sequence Cannot Be Blank": Exit Sub
    ws_PrcsSeqArr = Split(pv_NewValue, ",")
    For i = 0 To UBound(ws_PrcsSeqArr)
      If IsArray(ws_PrcsSeqArr) Then
      Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("select vPMCd from vParam where " + _
                                          "vPCoCd = '" + ADC("HCoCd") + "' and vPTyp= 'PRCSSEQ' and " + _
                                          "vPMCd = '" + ws_PrcsSeqArr(i) + "'"))
      If Cancel = True Then ErrMsg = "Value(" + CStr(i + 1) + ") Is An Invalid Process Sequence": Exit Sub
      End If
    Next i
      '*** Set the current tab as 0
    
  Case Is = UCase("hDmVaCtg")
    Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("select PMCd from Param where " + _
                                          "PTyp= 'VACTG' and PMCd = '" + pv_NewValue + "'"))
    ErrMsg = "Invalid Value Addition Category": Exit Sub
  Case Is = UCase("hOmPrtCd")
    Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("select PMCd from Param where " + _
                                         "PTyp= 'PRTCD' and PMCd = '" + pv_NewValue + "'"))
    ErrMsg = "Invalid Order Priority Code": Exit Sub
  Case Is = UCase("hBOpnLoc")
    Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("select LocCd from Loc where " + _
                                         "LocCoCd= '" + ADC("HCoCd") + "' and " + _
                                         "LocTyp= 'P' and LocCd= '" + pv_NewValue + "'"))
    ErrMsg = "Invalid Bag Opening Location": Exit Sub
    
  Case Is = UCase("HMinDiaProcess"), UCase("hChkWaxPullOutLocs"), UCase("hChkCastLocs")   '4.1.3.0
    ' **** Manali 3.5.0 - 08/06/09 - Multiple Locations allowed
    Dim ws_PLocArr() As String, k As Integer
    ws_PLocArr = Split(pv_NewValue, ",")
    If IsArray(ws_PLocArr) Then
      For k = 0 To UBound(ws_PLocArr)
        Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("select LocCd from Loc where " + _
                                             "LocCoCd= '" + ADC("HCoCd") + "' and " + _
                                             "LocTyp= 'R' and LocCd= '" + ws_PLocArr(k) + "'"))
        If Cancel = True Then ErrMsg = "Value(" + CStr(k + 1) + ") Is An Invalid Process Location": Exit Sub
      Next k
    End If
  
''    Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("select LocCd from Loc where " + _
''                                         "LocCoCd= '" + ADC("HCoCd") + "' and " + _
''                                         "LocTyp= 'P' and LocCd= '" + pv_NewValue + "'"))
''    ErrMsg = "Invalid Process Loaction for applying minimum tolerance on Diamond": Exit Sub

  Case Is = UCase("hBChr")
    Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("select PMCd from Param where " + _
                                      "PTyp= 'BCHR' and PMCd = '" + pv_NewValue + "'"))
    ErrMsg = "Invalid Bag Character": Exit Sub
  Case Is = UCase("hTdAck")
    Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("select PMCd from Param where " + _
                                       "PTyp= 'YN' and PMCd = '" + pv_NewValue + "'"))
    ErrMsg = "Enter Acknowledgment Required Yes Or No (Y/N)": Exit Sub
  Case Is = UCase("hFgPlChr")
    Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("select PMCd from Param where " + _
                                         "PTyp= 'PLCHR' and PMCd = '" + pv_NewValue + "'"))
    ErrMsg = "Invalid Packing List Character": Exit Sub
  Case Is = UCase("hBagNoEditYN")
    Cancel = (Not moCn.RecSeek("select PMCd from Param where " + _
                                      "PTyp= 'YN' and PMCd = '" + pv_NewValue + "'"))
    ErrMsg = "Allow Bag No. Edit? (Y/N)": Exit Sub
  Case Is = UCase("hYy")
    Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("select PMCd from Param where " + _
                                      "PTyp= 'YY' and PMCd = '" + pv_NewValue + "'"))
    ErrMsg = "Invalid Year": Exit Sub
  Case Is = UCase("hZWtEq")
    Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("select PMCd from Param where " + _
                                       "PTyp= 'ZWTEQ' and PMCd = '" + pv_NewValue + "'"))
    ErrMsg = "Invalid Customs Weight Equal To Option": Exit Sub
  Case Is = UCase("hOrdDsgGr")
    Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("select PMCd from Param where " + _
                                      "PTyp= 'ORDDSGGR' and PMCd = '" + pv_NewValue + "'"))
    ErrMsg = "Invalid Order Design Group Option": Exit Sub
  Case Is = UCase("hOrdRmGr")
    Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("select PMCd from Param where " + _
                                      "PTyp= 'ORDRMGR' and PMCd = '" + pv_NewValue + "'"))
    ErrMsg = "Invalid Order Raw Material Group Option": Exit Sub
  Case Is = UCase("hOrdLabGr")
    Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("select PMCd from Param where " + _
                                      "PTyp= 'ORDLABGR' and PMCd = '" + pv_NewValue + "'"))
    ErrMsg = "Invalid Order Labour Group Option": Exit Sub
  Case Is = UCase("hOthLocBag")
    Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("select PMCd from Param where " + _
                                       "PTyp= 'YN' and PMCd = '" + pv_NewValue + "'"))
    ErrMsg = "Show Bags Of All Locations In Help Of Bag? Enter Yes Or No (Y/N)": Exit Sub
  Case Is = UCase("hOdAmdQty")
    Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("select PMCd from Param where " + _
                                       "PTyp= 'YN' and PMCd = '" + pv_NewValue + "'"))
    ErrMsg = "Show Order Amended Qty Or Entered Qty In The Order Entry Summary Tab (Enter Y/ N)": Exit Sub
  Case Is = UCase("HTotAs")
    Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("select PMCd from Param where PTyp= 'INVTOTAS' and " + _
                         "PMCd = '" + pv_NewValue + "'"))
    If Cancel = True Then ErrMsg = "Invalid 'Invoice Total Value As' Option": Exit Sub
  Case Is = UCase("hZValEqInv")
    Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("select PMCd from Param where " + _
                                       "PTyp= 'YN' and PMCd = '" + pv_NewValue + "'"))
    ErrMsg = "Specify whether Customs Value Equal To Invoice Value (Enter Y/ N)": Exit Sub
  Case Is = UCase("HCmCdHlp")
    Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("select PMCd from Param where " + _
                                       "PTyp= 'YN' and PMCd = '" + pv_NewValue + "'"))
    ErrMsg = "Specify Whether Customer Should Be Sorted On Cust Name Or Cust Cd ('Y' for Name & 'N' for Cd)": Exit Sub
  Case Is = UCase("HRndOffFOB")
    Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("select PMCd from Param where " + _
                                       "PTyp= 'YN' and PMCd = '" + pv_NewValue + "'"))
    ErrMsg = "Specify Whether To Round Off The Inv & Custom FOB Values At The Dsg Level In Invoicing (Enter Y/ N)": Exit Sub
  Case Is = UCase("HShowActWt")
    Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("select PMCd from Param where " + _
                                       "PTyp= 'YN' and PMCd = '" + pv_NewValue + "'"))
    ErrMsg = "Specify Whether To Show The Actual Weight Field In The Rm Grid In Inv Dsg Entry (Enter Y/ N)": Exit Sub
  Case Is = UCase("HSrtBasePurVch")
    Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("select PMCd from Param where " + _
                                       "PTyp= 'SRTPRHLP' and PMCd = '" + pv_NewValue + "'"))
    ErrMsg = "Invalid Sort Option For Custom Purchase Vch Help Of Inv Custm Match Entry": Exit Sub
  Case Is = UCase("hInRepWtValAs")
    Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("select PMCd from Param where " + _
                                       "PTyp= 'INWT' and PMCd = '" + pv_NewValue + "'"))
    ErrMsg = "Specify Whether To Show Invoice Or Customs Wt & Amt In Inv Printing Rep": Exit Sub
  Case Is = UCase("hInRepWtFld")
    Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("select PMCd from Param where " + _
                                       "PTyp= 'YN' and PMCd = '" + pv_NewValue + "'"))
    ErrMsg = "Specify Whether To Show Wt Column In Inv Printing Rep (Enter Y/ N)": Exit Sub
  Case Is = UCase("hInRepMetLs")
    Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("select PMCd from Param where " + _
                                       "PTyp= 'YN' and PMCd = '" + pv_NewValue + "'"))
    ErrMsg = "Specify Whether To Show Metal Loss Summary In Inv Printing Rep (Enter Y/ N)": Exit Sub
  Case Is = UCase("hInRepCnvRt")
    Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("select PMCd from Param where " + _
                                       "PTyp= 'YN' and PMCd = '" + pv_NewValue + "'"))
    ErrMsg = "Specify Whether To Show Conversion Rt & Rupee Equivalent In Inv Printing Rep (Enter Y/ N)": Exit Sub
  Case Is = UCase("hInRepRtFld")
    Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("select PMCd from Param where " + _
                                       "PTyp= 'YN' and PMCd = '" + pv_NewValue + "'"))
    ErrMsg = "Specify Whether To Show Rt Column In Inv Printing Rep (Enter Y/ N)": Exit Sub
    
  ' ******** Manali 3.5.0 - 26/11/08 - Hide Commission field added
  Case Is = UCase("HHideCommYN")
    Cancel = Not moCn.RecSeek("select PMCd from Param where " + _
                                      "PTyp= 'YN' and PMCd = '" + pv_NewValue + "'")
    ErrMsg = "Commission Self Payable? Enter Yes Or No (Y/N)": Exit Sub
    'ErrMsg = "Hide Commission in Invoice Printing Report? Enter Yes Or No (Y/N)": Exit Sub
  
  ' ***** Manali 3.8.1 - Self Payable Options In Invoice added
  Case Is = UCase("HSelfDscYN")
    Cancel = Not moCn.RecSeek("select PMCd from Param where " + _
                                      "PTyp= 'YN' and PMCd = '" + pv_NewValue + "'")
    ErrMsg = "Discount Self Payable? Enter Yes Or No (Y/N)": Exit Sub
  
  Case Is = UCase("HSelfDeductYN")
    Cancel = Not moCn.RecSeek("select PMCd from Param where " + _
                                      "PTyp= 'YN' and PMCd = '" + pv_NewValue + "'")
    ErrMsg = "Deduction Self Payable? Enter Yes Or No (Y/N)": Exit Sub
  
  Case Is = UCase("hPtrYN")
    Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("select PMCd from Param where " + _
                                       "PTyp= 'YN' and PMCd = '" + pv_NewValue + "'"))
    ErrMsg = "Specify Whether To Allow User To Change The Value In The Column 'Ptr Y/N' Of Param Type 'RMSCTG' (Enter Y/ N)": Exit Sub
  Case Is = UCase("hPdcSfx")
    Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("select PMCd from Param where " + _
                                       "PTyp= 'PDCSFX' and PMCd = '" + pv_NewValue + "'"))
    ErrMsg = "Specify Method By Which Suffix Will Be Generated In The PDC Entry": Exit Sub
  Case Is = UCase("HRndOrdSalPrc")
    Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("select PMCd from Param where " + _
                                       "PTyp= 'RNDOPT' and PMCd = '" + pv_NewValue + "'"))
    ErrMsg = "Specify Method In Which Sales Price Will Be Rounded In The Order Entry": Exit Sub
  Case Is = UCase("HOrdCpyRtOpt")
    Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("select PMCd from Param where PTyp= 'CPYRT' and PMCd = '" + pv_NewValue + "'"))
    ErrMsg = "Invalid RmCtg Combination For Which Rates Have To Be Copied": Exit Sub
  Case Is = UCase("HShowPrdRmDets")
    Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("select PMCd from Param where PTyp= 'YN' and PMCd = '" + pv_NewValue + "'"))
    ErrMsg = "Specify Whether To Show Production Rm Qty && Wt In Design And Order Entries": Exit Sub
  Case Is = UCase("hGrsWtInBsYN")
    Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("select PMCd from Param where " + _
                                      "PTyp= 'YN' and PMCd = '" + pv_NewValue + "'"))
    ErrMsg = "Show Current Gross Wt in Bag Split? Enter Yes Or No (Y/N)": Exit Sub
  Case Is = UCase("hExpLocal")
    Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("select PMCd from Param where " + _
                                      "PTyp= 'YN' and PMCd = '" + pv_NewValue + "'"))
    ErrMsg = "Specify Whether Export Is From SEEPZ Or Outside? Enter Yes Or No (Y/N)": Exit Sub
  Case Is = UCase("hPpcYN")
    Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("select PMCd from Param where " + _
                                      "PTyp= 'YN' and PMCd = '" + pv_NewValue + "'"))
    ErrMsg = "Specify Whether PPC Module is to be Shown Or Not? Enter Yes Or No (Y/N)": Exit Sub
  Case Is = UCase("hLine")
    Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("select vPMCd from vParam where " + _
                                      "vPCoCd= '" + ADC("HCoCd") + "' and vPTyp= 'LINE' and " + _
                                      "vPMCd = '" + pv_NewValue + "'"))
    ErrMsg = "Invalid Line": Exit Sub
  Case Is = UCase("hWH")
    Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("select PMCd from Param where " + _
                                      "PTyp= 'WH' and PMCd = '" + pv_NewValue + "'"))
    ErrMsg = "Invalid Setting Option": Exit Sub
  '***** Shilpa ******
  Case Is = UCase("HInvTyp")
    Cancel = (Not moCn.RecSeek("select PMCd from Param where " + _
                                      "PTyp= 'INVTYP' and PMCd = '" + pv_NewValue + "'"))
    ErrMsg = "Invalid Invoice Type": Exit Sub
  Case Is = UCase("hValAddMtlColYN")
    Cancel = (Not moCn.RecSeek("select PMCd from Param where " + _
                                      "PTyp= 'YN' and PMCd = '" + pv_NewValue + "'"))
    ErrMsg = "Show Value Addition On Metal Column In Value Addition Report? Enter Yes Or No (Y/N)?": Exit Sub
  Case Is = UCase("hValBOEDetYN")
    Cancel = (Not moCn.RecSeek("select PMCd from Param where " + _
                                      "PTyp= 'YN' and PMCd = '" + pv_NewValue + "'"))
    ErrMsg = "Show BOE Details In Value Addition Annexur Report? Enter Yes Or No (Y/N)?": Exit Sub
  
  Case Is = UCase("HFgSubLoc")
    Cancel = (Not moCn.RecSeek("select vPMCd from vParam where vPCoCd= '" + ADC("HCoCd") + "' and " + _
                               "vPTyp= 'FGSUBLOC' and vPMCd = '" + pv_NewValue + "'"))
    ErrMsg = "Invalid Fg Sub Location Code": Exit Sub
  
  Case Is = UCase("hNegStkYN")
    Cancel = (Not moCn.RecSeek("select PMCd from Param where " + _
                                      "PTyp= 'YN' and PMCd = '" + pv_NewValue + "'"))
    If Cancel = True Then ErrMsg = "Allow Negative Stock without SEO Password? (Y/N)": Exit Sub
    
    ' ****** Sachin 2.14.0 - [07/10/2006] - Neg Stk RmCtg will be Blank if Neg Stk is allowed ******
    If pv_NewValue = "Y" Then ADC("HNegStkRmCtg") = ""
    ' ****** Sachin 2.14.0 - [07/10/2006] ******
  
  Case Is = UCase("HAvgSRtYn")
    Cancel = (Not moCn.RecSeek("select PMCd from Param where " + _
                                      "PTyp= 'YN' and PMCd = '" + pv_NewValue + "'"))
    ErrMsg = "Implement Average Stock Rate? (Y/N)": Exit Sub

  Case Is = UCase("HSizeStkRt")
    Cancel = (Not moCn.RecSeek("select PMCd from Param where " + _
                                      "PTyp= 'YN' and PMCd = '" + pv_NewValue + "'"))
    ErrMsg = "Invoice with size and Stock Rate? (Y/N)": Exit Sub
    
  Case Is = UCase("hPrdPtsYN")
    Cancel = (Not moCn.RecSeek("select PMCd from Param where " + _
                                      "PTyp= 'YN' and PMCd = '" + pv_NewValue + "'"))
    ErrMsg = "Allow Editing of Production Points? (Y/N)": Exit Sub
  Case Is = UCase("HCtbMod")
    Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("select PMCd from Param where " + _
                                       "PTyp= 'YN' and PMCd = '" + pv_NewValue + "'"))
    ErrMsg = "Show The Contractor Bill Module In The System (Enter Y/ N)": Exit Sub
  Case Is = UCase("HCtCd")
    Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("select CmCd from CustMst where " + _
                                       "CmCtg= 'T' and CmCd = '" + pv_NewValue + "'"))
    ErrMsg = "Invalid Contractor Code": Exit Sub
  '*************************** ZUBIN **************************
  ' 10th Dec 2003, EMR206
  Case Is = UCase("HPrtAllotOnPrdYN")
    Cancel = pv_NewValue <> vbNullString And (Not moCn.RecSeek("select PMCd from Param " + _
                          "where PTyp= 'YN' and PMCd = '" + pv_NewValue + "'"))
    ErrMsg = "Specify Whether The Priority Allotment Is Based On Production, On " + _
                  "Export Or Should There Be No Priority Allotment? " + _
                  "Enter Yes Or No Or Leave Blank (Y/N/'')?": Exit Sub
  '*************************** ZUBIN **************************
  '*************************** Geeta**************************
  ' 15th Jan 2004, EMR206
    Case Is = UCase("HChkDCWtAtIssYN")
    Cancel = (Not moCn.RecSeek("select PMCd from Param " + _
                          "where PTyp= 'YN' and PMCd = '" + pv_NewValue + "'"))
    ErrMsg = "Specify Whether Diamond/Color Stones Wt Checked At Issue, " + _
                  "Enter Yes Or No (Y/N)?": Exit Sub
    Case Is = UCase("HChkDCQtyAtIssYN")
    Cancel = (Not moCn.RecSeek("select PMCd from Param " + _
                          "where PTyp= 'YN' and PMCd = '" + pv_NewValue + "'"))
    ErrMsg = "Specify Whether Diamond/Color Stones Qty Checked At Issue, " + _
                  "Enter Yes Or No (Y/N)?": Exit Sub
   
    Case Is = UCase("HChkGPWtAtIssYN")
    Cancel = (Not moCn.RecSeek("select PMCd from Param " + _
                          "where PTyp= 'YN' and PMCd = '" + pv_NewValue + "'"))
    ErrMsg = "Specify Whether Gold/Platinum Wt Checked At Issue, " + _
                  "Enter Yes Or No (Y/N)?": Exit Sub
     
    Case Is = UCase("hFgRmEqOrdRm")
    Cancel = (Not moCn.RecSeek("select PMCd from Param " + _
                          "where PTyp= 'YN' and PMCd = '" + pv_NewValue + "'"))
    ErrMsg = "Specify Whether To Check All The RawMaterials In Fg With Order Rm?" + _
                  "(Enter Yes Or No (Y/N))": Exit Sub
                    
   '*************************** Geeta**************************
   '*****Geeta***Emr208***
   '***** Manali 3.03 - 16/06/08 - Palladium Fineness included
   Case Is = UCase("HGldFiness"), UCase("HPlFiness"), UCase("HSlFiness"), UCase("HPdFiness")
        Cancel = IIF(pv_NewValue <= 0, True, False)
    ErrMsg = "Gold/Platinum/Silver/Palladium Fineness Should be >0": Exit Sub
   ' Manoj 2.9.0 Patch
   Case Is = UCase("HStkRtPassReq")
    Cancel = (Not moCn.RecSeek("select PMCd from Param " + _
                          "where PTyp= 'YN' and PMCd = '" + pv_NewValue + "'"))
    ErrMsg = "Specify Whether SEO Password Is Requre To Edit Stock Rate? (Enter Y/N)": Exit Sub
  ' Manoj 2.9.0 Patch

  ' **** Zubin 211 **** '
  Case Is = UCase("hFluteBagYN")
    Cancel = (Not moCn.RecSeek("select 'x' from Param " + _
                          "where PTyp= 'YN' and PMCd = '" + pv_NewValue + "' and PSCd= ''"))
    ErrMsg = "Specify Whether To Use Flute Bags? (Enter Y/N)": Exit Sub
  Case Is = UCase("HOrdPrdPts")
    Cancel = Not moCn.RecSeek("select 'x' from Param " + _
                          "where PTyp= 'YN' and PMCd = '" + pv_NewValue + "' and PSCd= ''")
    ErrMsg = "Restrict Production Points On Order RmQty, Enter Yes Or No (Y/N)?": Exit Sub
  Case Is = UCase("HDTPrdPtr")
    Cancel = Not moCn.RecSeek("select 'x' from Param " + _
                          "where PTyp= 'YN' and PMCd = '" + pv_NewValue + "' and PSCd= ''")
    ErrMsg = "In DT Issue By Production Pointer, Enter Yes Or No (Y/N)?": Exit Sub
' **** Zubin 211 **** '

  '*** 2.11 (Jen)
  Case Is = UCase("HOdGmChk")
    Cancel = Not moCn.RecSeek("select 'x' from Param where PTyp= 'CPYRT' " + _
                              "and PMCd = '" + pv_NewValue + "' and PSCd= ''")
    ErrMsg = "Invalid Gram Check Option": Exit Sub
  '*** 2.11 (Jen)
  
  ' **** Zubin 212 **** '
  Case Is = UCase("HLotNoYN")
    Cancel = Not moCn.RecSeek("select 'x' from Param " + _
                          "where PTyp= 'YN' and PMCd = '" + pv_NewValue + "' and PSCd= ''")
    ErrMsg = "Implement Lot No. For DC/X/G/P/S/L/M, Enter Yes Or No (Y/N)?": Exit Sub
  
  ' ****** Sachin 2.12 - 06/12/05 - SJM - ******
  Case Is = UCase("HSlvModYN")
    Cancel = Not moCn.RecSeek("select 'x' from Param " + _
                          "where PTyp= 'YN' and PMCd = '" + pv_NewValue + "' and PSCd= ''")
    ErrMsg = "Implement Silver Module, Enter Yes Or No (Y/N)?": Exit Sub
  ' ****** Sachin 2.12 - 06/12/05 - SJM - ******
  
  ' ****** Manali 3.03 - 16/08/06 - Palladium Module
  Case Is = UCase("HPdModYN")
    Cancel = Not moCn.RecSeek("select 'x' from Param " + _
                          "where PTyp= 'YN' and PMCd = '" + pv_NewValue + "' and PSCd= ''")
    ErrMsg = "Implement Palladium Module, Enter Yes Or No (Y/N)?": Exit Sub
  ' ****** Manali 3.03 - 16/08/06 - Palladium Module
  
  ' Zubin 212 (23/02/06)
  Case Is = UCase("HFgQuotYN")
    Cancel = Not moCn.RecSeek("select 'x' from Param " + _
                          "where PTyp= 'YN' and PMCd = '" + pv_NewValue + "' and PSCd= ''")
    ErrMsg = "Show Quotation For FG Bag, Enter Yes Or No (Y/N)?": Exit Sub
  Case Is = UCase("HGldRtDAYN")
    Cancel = Not moCn.RecSeek("select 'x' from Param " + _
                          "where PTyp= 'YN' and PMCd = '" + pv_NewValue + "' and PSCd= ''")
    ErrMsg = "Derived/Actual GldRt In Ord/Inv Prnt, Enter Yes Or No (Y/N)?": Exit Sub
  ' Zubin 212 (23/02/06)
  
  '*** (Jen 2.12 Patch 3)
  Case Is = UCase("HLooseInv")
    Cancel = Not moCn.RecSeek("select 'x' from Param where PTyp= 'YN' " + _
             "and PMCd = '" + pv_NewValue + "' and PSCd= ''")
    ErrMsg = "Show Loose RM Invoice Menus, Enter Yes Or No (Y/N)?": Exit Sub
  '*** (Jen 2.12 Patch 3)
  
  '*** (Jen 2.13)
  Case Is = UCase("hConsolidatedInv")
    Cancel = Not moCn.RecSeek("select 'x' from Param where PTyp= 'YN' " + _
             "and PMCd = '" + pv_NewValue + "' and PSCd= ''")
    ErrMsg = "Show Consolidated Invoice Report Menus, Enter Yes Or No (Y/N)?": Exit Sub
  Case Is = UCase("hAutomateReturns")
    Cancel = Not moCn.RecSeek("select 'x' from Param where PTyp= 'YN' " + _
             "and PMCd = '" + pv_NewValue + "' and PSCd= ''")
    ErrMsg = "Automate Invoice Returns Yes Or No (Y/N)?": Exit Sub
  '*** (Jen 2.13)
  
  Case Is = UCase("hGenExpNoYN")
    Cancel = Not moCn.RecSeek("select PMCd from Param where " + _
                                      "PTyp= 'YN' and PMCd = '" + pv_NewValue + "'")
    ErrMsg = "Generate the Export Number? Enter Yes Or No (Y/N)": Exit Sub
  
  ' ****** Sachin 2.13.0 - 05/06/2006 ******
  Case Is = UCase("hQWModYN")
    Cancel = Not moCn.RecSeek("select 'x' from Param " + _
                          "where PTyp= 'YN' and PMCd = '" + pv_NewValue + "' and PSCd= ''")
    ErrMsg = "Implement QW Module, Enter Yes Or No (Y/N)?": Exit Sub
  ' ****** Sachin 2.13.0 - 05/06/2006 ******
  
  ' Zubin 213
  Case Is = UCase("HBaseCurCd")
    Cancel = Not moCn.RecSeek("select 'x' from Param " + _
                          "where PTyp= 'CURNCY' and PMCd = '" + pv_NewValue + "' and PSCd= ''")
    ErrMsg = "Invalid Base Currency Code": Exit Sub
  Case Is = UCase("hMultiCurYN")
    Cancel = Not moCn.RecSeek("select 'x' from Param " + _
                          "where PTyp= 'YN' and PMCd = '" + pv_NewValue + "' and PSCd= ''")
    ErrMsg = "Invalid 'Implement Multi Currency Module'": Exit Sub
  ' Zubin 213
  '*** JAY 2.13(CT) ***
  Case Is = UCase("hCompBagsYn")
    Cancel = Not moCn.RecSeek("select 'x' from Param " + _
                          "where PTyp= 'YN' and PMCd = '" + pv_NewValue + "' and PSCd= ''")
    ErrMsg = "Invalid 'Component Bags Module'": Exit Sub
  '*** JAY 2.13(CT) ***
  
  ' ****** Sachin 2.14.0 - [07/10/2006] ******
  Case Is = UCase("hNegStkRmCtg")
    '****** Blank New Value would mean that the Neg Check will apply to All RmCtgs
    Cancel = pv_NewValue <> "" And (Not moCn.RecSeek("select PMCd from Param where " + _
                                         "PTyp= 'GLDAS' and PMCd = '" + pv_NewValue + "'"))
    If Cancel = True Then ErrMsg = "Invalid Negative Stock RM Ctg. Value": Exit Sub
  ' ****** Sachin 2.14.0 - [07/10/2006] ******
  
  ' ****** Sachin 2.14.0 - 26/10/2006 - [Multi-Company] ******
  Case Is = UCase("hImpComLoc")
    Cancel = Not moCn.RecSeek("select 'x' from Param " + _
                          "where PTyp= 'YN' and PMCd = '" + pv_NewValue + "' and PSCd= ''")
    ErrMsg = "Implement Common Locations, Enter Yes Or No (Y/N)?": Exit Sub
  ' ****** Sachin 2.13.0 - 05/06/2006 - [Multi-Company]  ******
  
  '*** JAY 2.14(LclInv) ***
  Case Is = UCase("hLclInvYn")
    Cancel = Not moCn.RecSeek("select 'x' from Param " + _
                          "where PTyp= 'YN' and PMCd = '" + pv_NewValue + "' and PSCd= ''")
    ErrMsg = "Implement Local Invoice, Enter Yes Or No (Y/N)?": Exit Sub
  '*** JAY 2.14(LclInv) ***
  
  '*** (Jen 3.01)
  Case Is = UCase("HCustomiseOrd")
    Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("select PMCd from Param where " + _
                                      "PTyp= 'YN' and PMCd = '" + pv_NewValue + "'"))
    ErrMsg = "Specify Whether To Customise Order Entry Or Not? Enter Yes Or No (Y/N)": Exit Sub
  
  Case Is = UCase("HRtMinusD")
    Cancel = (pv_NewValue <> 0) And ADC("HRtPlusD") <> 0 And Val(pv_NewValue) > Val(ADC("HRtPlusD"))
    If Cancel = True Then ErrMsg = "From Rate Variation % For Diamonds Cannot Be Greater Than To Rate Variation %": Exit Sub
  Case Is = UCase("HRtMinusC")
    Cancel = (pv_NewValue <> 0) And ADC("HRtPlusC") <> 0 And Val(pv_NewValue) > Val(ADC("HRtPlusC"))
    If Cancel = True Then ErrMsg = "From Rate Variation % For Color Stones Cannot Be Greater Than To Rate Variation %": Exit Sub
  Case Is = UCase("HRtMinusG")
    Cancel = (pv_NewValue <> 0) And ADC("HRtPlusG") <> 0 And Val(pv_NewValue) > Val(ADC("HRtPlusG"))
    If Cancel = True Then ErrMsg = "From Rate Variation % For Gold Cannot Be Greater Than To Rate Variation %": Exit Sub
  Case Is = UCase("HRtMinusP")
    Cancel = (pv_NewValue <> 0) And ADC("HRtPlusP") <> 0 And Val(pv_NewValue) > Val(ADC("HRtPlusP"))
    If Cancel = True Then ErrMsg = "From Rate Variation % For Platinum Cannot Be Greater Than To Rate Variation %": Exit Sub
  Case Is = UCase("HRtMinusS")
    Cancel = (pv_NewValue <> 0) And ADC("HRtPlusS") <> 0 And Val(pv_NewValue) > Val(ADC("HRtPlusS"))
    If Cancel = True Then ErrMsg = "From Rate Variation % For Silver Cannot Be Greater Than To Rate Variation %": Exit Sub
  'Manali 3.03 - 18/06/08 - Palladium added
  Case Is = UCase("HRtMinusL")
    Cancel = (pv_NewValue <> 0) And ADC("HRtPlusL") <> 0 And Val(pv_NewValue) > Val(ADC("HRtPlusL"))
    If Cancel = True Then ErrMsg = "From Rate Variation % For Palladium Cannot Be Greater Than To Rate Variation %": Exit Sub
  'Manali 3.03 - 18/06/08 - Palladium added
  Case Is = UCase("HRtMinusX")
    Cancel = (pv_NewValue <> 0) And ADC("HRtPlusX") <> 0 And Val(pv_NewValue) > Val(ADC("HRtPlusX"))
    If Cancel = True Then ErrMsg = "From Rate Variation % For Accessories Cannot Be Greater Than To Rate Variation %": Exit Sub
  
  Case Is = UCase("HRtPlusD")
    Cancel = (pv_NewValue) <> 0 And ADC("HRtMinusD") <> 0 And Val(pv_NewValue) < Val(ADC("HRtMinusD"))
    If Cancel = True Then ErrMsg = "From Rate Variation % For Diamonds Cannot Be Greater Than To Rate Variation %": Exit Sub
  Case Is = UCase("HRtPlusC")
    Cancel = (pv_NewValue <> 0) And ADC("HRtMinusC") <> 0 And Val(pv_NewValue) < Val(ADC("HRtMinusC"))
    If Cancel = True Then ErrMsg = "From Rate Variation % For Color Stones Cannot Be Greater Than To Rate Variation %": Exit Sub
  Case Is = UCase("HRtPlusG")
    Cancel = (pv_NewValue <> 0) And ADC("HRtMinusG") <> 0 And Val(pv_NewValue) < Val(ADC("HRtMinusG"))
    If Cancel = True Then ErrMsg = "From Rate Variation % For Gold Cannot Be Greater Than To Rate Variation %": Exit Sub
  Case Is = UCase("HRtPlusP")
    Cancel = (pv_NewValue <> 0) And ADC("HRtMinusP") <> 0 And Val(pv_NewValue) < Val(ADC("HRtMinusP"))
    If Cancel = True Then ErrMsg = "From Rate Variation % For Platinum Cannot Be Greater Than To Rate Variation %": Exit Sub
  Case Is = UCase("HRtPlusS")
    Cancel = (pv_NewValue <> 0) And ADC("HRtMinusS") <> 0 And Val(pv_NewValue) < Val(ADC("HRtMinusS"))
    If Cancel = True Then ErrMsg = "From Rate Variation % For Silver Cannot Be Greater Than To Rate Variation %": Exit Sub
  'Manali 3.03 - 18/06/08 - Palladium added
  Case Is = UCase("HRtPlusL")
    Cancel = (pv_NewValue <> 0) And ADC("HRtMinusL") <> 0 And Val(pv_NewValue) < Val(ADC("HRtMinusL"))
    If Cancel = True Then ErrMsg = "From Rate Variation % For Palladium Cannot Be Greater Than To Rate Variation %": Exit Sub
  'Manali 3.03 - 18/06/08 - Palladium added
  Case Is = UCase("HRtPlusX")
    Cancel = (pv_NewValue <> 0) And ADC("HRtMinusX") <> 0 And Val(pv_NewValue) < Val(ADC("HRtMinusX"))
    If Cancel = True Then ErrMsg = "From Rate Variation % For Accessories Cannot Be Greater Than To Rate Variation %": Exit Sub
  
  Case Is = UCase("HPtrMinusD")
    Cancel = (pv_NewValue <> 0) And ADC("HPtrPlusD") <> 0 And Val(pv_NewValue) > Val(ADC("HPtrPlusD"))
    If Cancel = True Then ErrMsg = "From Pc/Ct Variation % For Diamonds Cannot Be Greater Than To Pc/Ct Variation %": Exit Sub
  Case Is = UCase("HPtrMinusC")
    Cancel = (pv_NewValue <> 0) And ADC("HPtrPlusC") <> 0 And Val(pv_NewValue) > Val(ADC("HPtrPlusC"))
    If Cancel = True Then ErrMsg = "From Pc/Ct Variation % For Color Stones Cannot Be Greater Than To Pc/Ct Variation %": Exit Sub
  Case Is = UCase("HPtrMinusG")
    Cancel = (pv_NewValue <> 0) And ADC("HPtrPlusG") <> 0 And Val(pv_NewValue) > Val(ADC("HPtrPlusG"))
    If Cancel = True Then ErrMsg = "From Pc/Ct Variation % For Gold Cannot Be Greater Than To Pc/Ct Variation %": Exit Sub
  Case Is = UCase("HPtrMinusP")
    Cancel = (pv_NewValue <> 0) And ADC("HPtrPlusP") <> 0 And Val(pv_NewValue) > Val(ADC("HPtrPlusP"))
    If Cancel = True Then ErrMsg = "From Pc/Ct Variation % For Platinum Cannot Be Greater Than To Pc/Ct Variation %": Exit Sub
  Case Is = UCase("HPtrMinusS")
    Cancel = (pv_NewValue <> 0) And ADC("HPtrPlusS") <> 0 And Val(pv_NewValue) > Val(ADC("HPtrPlusS"))
    If Cancel = True Then ErrMsg = "From Pc/Ct Variation % For Silver Cannot Be Greater Than To Pc/Ct Variation %": Exit Sub
  'Manali 3.03 - 18/06/08 - Palladium added
  Case Is = UCase("HPtrMinusL")
    Cancel = (pv_NewValue <> 0) And ADC("HPtrPlusL") <> 0 And Val(pv_NewValue) > Val(ADC("HPtrPlusL"))
    If Cancel = True Then ErrMsg = "From Pc/Ct Variation % For Palladium Cannot Be Greater Than To Pc/Ct Variation %": Exit Sub
  'Manali 3.03 - 18/06/08 - Palladium added
  Case Is = UCase("HPtrMinusX")
    Cancel = (pv_NewValue <> 0) And ADC("HPtrPlusX") <> 0 And Val(pv_NewValue) > Val(ADC("HPtrPlusX"))
    If Cancel = True Then ErrMsg = "From Pc/Ct Variation % For Accessories Cannot Be Greater Than To Pc/Ct Variation %": Exit Sub
  
  Case Is = UCase("HPtrPlusD")
    Cancel = (pv_NewValue) <> 0 And ADC("HPtrMinusD") <> 0 And Val(pv_NewValue) < Val(ADC("HPtrMinusD"))
    If Cancel = True Then ErrMsg = "From Pc/Ct Variation % For Diamonds Cannot Be Greater Than To Pc/Ct Variation %": Exit Sub
  Case Is = UCase("HPtrPlusC")
    Cancel = (pv_NewValue <> 0) And ADC("HPtrMinusC") <> 0 And Val(pv_NewValue) < Val(ADC("HPtrMinusC"))
    If Cancel = True Then ErrMsg = "From Pc/Ct Variation % For Color Stones Cannot Be Greater Than To Pc/Ct Variation %": Exit Sub
  Case Is = UCase("HPtrPlusG")
    Cancel = (pv_NewValue <> 0) And ADC("HPtrMinusG") <> 0 And Val(pv_NewValue) < Val(ADC("HPtrMinusG"))
    If Cancel = True Then ErrMsg = "From Pc/Ct Variation % For Gold Cannot Be Greater Than To Pc/Ct Variation %": Exit Sub
  Case Is = UCase("HPtrPlusP")
    Cancel = (pv_NewValue <> 0) And ADC("HPtrMinusP") <> 0 And Val(pv_NewValue) < Val(ADC("HPtrMinusP"))
    If Cancel = True Then ErrMsg = "From Pc/Ct Variation % For Platinum Cannot Be Greater Than To Pc/Ct Variation %": Exit Sub
  Case Is = UCase("HPtrPlusS")
    Cancel = (pv_NewValue <> 0) And ADC("HPtrMinusS") <> 0 And Val(pv_NewValue) < Val(ADC("HPtrMinusS"))
    If Cancel = True Then ErrMsg = "From Pc/Ct Variation % For Silver Cannot Be Greater Than To Pc/Ct Variation %": Exit Sub
  'Manali 3.03 - 18/06/08 - Palladium added
  Case Is = UCase("HPtrPlusL")
    Cancel = (pv_NewValue <> 0) And ADC("HPtrMinusL") <> 0 And Val(pv_NewValue) < Val(ADC("HPtrMinusL"))
    If Cancel = True Then ErrMsg = "From Pc/Ct Variation % For Palladium Cannot Be Greater Than To Pc/Ct Variation %": Exit Sub
  'Manali 3.03 - 18/06/08 - Palladium added
  Case Is = UCase("HPtrPlusX")
    Cancel = (pv_NewValue <> 0) And ADC("HPtrMinusX") <> 0 And Val(pv_NewValue) < Val(ADC("HPtrMinusX"))
    If Cancel = True Then ErrMsg = "From Pc/Ct Variation % For Accessories Cannot Be Greater Than To Pc/Ct Variation %": Exit Sub
  
  'Manali 3.03 - 18/06/08 - Palladium added
  Case Is = UCase("HImpTypNotInSelD"), UCase("HImpTypNotInSelC"), UCase("HImpTypNotInSelG"), _
            UCase("HImpTypNotInSelP"), UCase("HImpTypNotInSelS"), UCase("HImpTypNotInSelL"), UCase("HImpTypNotInSelX")
    Dim ws_ImpTypArr() As String, wi_i As Integer
    If pv_NewValue <> "" Then
      ws_ImpTypArr = Split(pv_NewValue, ",")
      If IsArray(ws_ImpTypArr) Then
        For wi_i = 0 To UBound(ws_ImpTypArr)
          Cancel = Not moCn.RecSeek("Select PMCd From Param where PTyp= 'IMPTYP' " + _
                   "and PMCd = '" + ws_ImpTypArr(wi_i) + "'")
          If Cancel = True Then ErrMsg = "Value(" + CStr(wi_i + 1) + ") Is An Invalid Import Type": Exit Sub
        Next wi_i
      End If
    End If
    
  'Manali 3.2.2 - Check appiled on import dates
  Case Is = UCase("HImpAftD")
    Cancel = pv_NewValue = ""
    If Cancel = True Then ErrMsg = "Imports After Date Cannot Be Blank": Exit Sub
    Cancel = pv_NewValue <> MWLib.EmptyDate And ADC("HImpBefD") <> MWLib.EmptyDate And pv_NewValue > ADC("HImpBefD")
    If Cancel = True Then ErrMsg = "Imports After Date Cannot Be Greater Than Imports Before Date For Diamonds": Exit Sub
  Case Is = UCase("HImpAftC")
    Cancel = pv_NewValue = ""
    If Cancel = True Then ErrMsg = "Imports After Date Cannot Be Blank": Exit Sub
    Cancel = pv_NewValue <> MWLib.EmptyDate And ADC("HImpBefC") <> MWLib.EmptyDate And pv_NewValue > ADC("HImpBefC")
    If Cancel = True Then ErrMsg = "Imports After Date Cannot Be Greater Than Imports Before Date For Colour Stones": Exit Sub
  Case Is = UCase("HImpAftG")
    Cancel = pv_NewValue = ""
    If Cancel = True Then ErrMsg = "Imports After Date Cannot Be Blank": Exit Sub
    Cancel = pv_NewValue <> MWLib.EmptyDate And ADC("HImpBefG") <> MWLib.EmptyDate And pv_NewValue > ADC("HImpBefG")
    If Cancel = True Then ErrMsg = "Imports After Date Cannot Be Greater Than Imports Before Date For Gold": Exit Sub
  Case Is = UCase("HImpAftP")
    Cancel = pv_NewValue = ""
    If Cancel = True Then ErrMsg = "Imports After Date Cannot Be Blank": Exit Sub
    Cancel = pv_NewValue <> MWLib.EmptyDate And ADC("HImpBefP") <> MWLib.EmptyDate And pv_NewValue > ADC("HImpBefP")
    If Cancel = True Then ErrMsg = "Imports After Date Cannot Be Greater Than Imports Before Date For Platinum": Exit Sub
  Case Is = UCase("HImpAftS")
    Cancel = pv_NewValue = ""
    If Cancel = True Then ErrMsg = "Imports After Date Cannot Be Blank": Exit Sub
    Cancel = pv_NewValue <> MWLib.EmptyDate And ADC("HImpBefS") <> MWLib.EmptyDate And pv_NewValue > ADC("HImpBefS")
    If Cancel = True Then ErrMsg = "Imports After Date Cannot Be Greater Than Imports Before Date For Silver": Exit Sub
  'Manali 3.03 - 18/06/08 - Palladium added
  Case Is = UCase("HImpAftL")
    Cancel = pv_NewValue = ""
    If Cancel = True Then ErrMsg = "Imports After Date Cannot Be Blank": Exit Sub
    Cancel = pv_NewValue <> MWLib.EmptyDate And ADC("HImpBefL") <> MWLib.EmptyDate And pv_NewValue > ADC("HImpBefL")
    If Cancel = True Then ErrMsg = "Imports After Date Cannot Be Greater Than Imports Before Date For Palladium": Exit Sub
  'Manali 3.03 - 18/06/08 - Palladium added
  Case Is = UCase("HImpAftX")
    Cancel = pv_NewValue = ""
    If Cancel = True Then ErrMsg = "Imports After Date Cannot Be Blank": Exit Sub
    Cancel = pv_NewValue <> MWLib.EmptyDate And ADC("HImpBefX") <> MWLib.EmptyDate And pv_NewValue > ADC("HImpBefX")
    If Cancel = True Then ErrMsg = "Imports After Date Cannot Be Greater Than Imports Before Date For Accessories": Exit Sub
    
  Case Is = UCase("HImpBefD")
    Cancel = pv_NewValue = ""
    If Cancel = True Then ErrMsg = "Imports Before Date Cannot Be Blank": Exit Sub
    Cancel = pv_NewValue <> MWLib.EmptyDate And ADC("HImpAftD") <> MWLib.EmptyDate And pv_NewValue < ADC("HImpAftD")
    If Cancel = True Then ErrMsg = "Imports After Date Cannot Be Greater Than Imports Before Date For Diamonds": Exit Sub
  Case Is = UCase("HImpBefC")
    Cancel = pv_NewValue = ""
    If Cancel = True Then ErrMsg = "Imports Before Date Cannot Be Blank": Exit Sub
    Cancel = pv_NewValue <> MWLib.EmptyDate And ADC("HImpAftC") <> MWLib.EmptyDate And pv_NewValue < ADC("HImpAftC")
    If Cancel = True Then ErrMsg = "Imports After Date Cannot Be Greater Than Imports Before Date For Colour Stones": Exit Sub
  Case Is = UCase("HImpBefG")
    Cancel = pv_NewValue = ""
    If Cancel = True Then ErrMsg = "Imports Before Date Cannot Be Blank": Exit Sub
    Cancel = pv_NewValue <> MWLib.EmptyDate And ADC("HImpAftG") <> MWLib.EmptyDate And pv_NewValue < ADC("HImpAftG")
    If Cancel = True Then ErrMsg = "Imports After Date Cannot Be Greater Than Imports Before Date For Gold": Exit Sub
  Case Is = UCase("HImpBefP")
    Cancel = pv_NewValue = ""
    If Cancel = True Then ErrMsg = "Imports Before Date Cannot Be Blank": Exit Sub
    Cancel = pv_NewValue <> MWLib.EmptyDate And ADC("HImpAftP") <> MWLib.EmptyDate And pv_NewValue < ADC("HImpAftP")
    If Cancel = True Then ErrMsg = "Imports After Date Cannot Be Greater Than Imports Before Date For Platinum": Exit Sub
  Case Is = UCase("HImpBefS")
    Cancel = pv_NewValue = ""
    If Cancel = True Then ErrMsg = "Imports Before Date Cannot Be Blank": Exit Sub
    Cancel = pv_NewValue <> MWLib.EmptyDate And ADC("HImpAftS") <> MWLib.EmptyDate And pv_NewValue < ADC("HImpAftS")
    If Cancel = True Then ErrMsg = "Imports After Date Cannot Be Greater Than Imports Before Date For Silver": Exit Sub
  'Manali 3.03 - 18/06/08 - Palladium added
  Case Is = UCase("HImpBefL")
    Cancel = pv_NewValue = ""
    If Cancel = True Then ErrMsg = "Imports Before Date Cannot Be Blank": Exit Sub
    Cancel = pv_NewValue <> MWLib.EmptyDate And ADC("HImpAftL") <> MWLib.EmptyDate And pv_NewValue < ADC("HImpAftL")
    If Cancel = True Then ErrMsg = "Imports After Date Cannot Be Less Than Imports Before Date For Palladium": Exit Sub
  'Manali 3.03 - 18/06/08 - Palladium added
  Case Is = UCase("HImpBefX")
    Cancel = pv_NewValue = ""
    If Cancel = True Then ErrMsg = "Imports Before Date Cannot Be Blank": Exit Sub
    Cancel = pv_NewValue <> MWLib.EmptyDate And ADC("HImpAftX") <> MWLib.EmptyDate And pv_NewValue < ADC("HImpAftX")
    If Cancel = True Then ErrMsg = "Imports After Date Cannot Be Less Than Imports Before Date For Accessories": Exit Sub
  'Manali 3.2.2 - Check appiled on import dates
  
  Case Is = UCase("HAutoCustmMtch")
    Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("select PMCd from Param where " + _
                                      "PTyp= 'YN' and PMCd = '" + pv_NewValue + "'"))
    If Cancel = True Then ErrMsg = "Specify Whether To Enable Automatic Custom Matching Or Not? Enter Yes Or No (Y/N)": Exit Sub
    
    If pv_NewValue = "Y" Then
      TabHd.TabVisible(5) = True
      TabHd.TabsPerRow = 7
    Else
      TabHd.TabVisible(5) = False
      TabHd.TabsPerRow = 6
    End If
  '*** (Jen 3.01)
  
  '*** (Jen 3.1.2)
  Case Is = UCase("hBestUse")
    Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("select PMCd from Param where " + _
                                      "PTyp= 'YN' and PMCd = '" + pv_NewValue + "'"))
    If Cancel = True Then ErrMsg = "Specify Whether To Enable Best Use Fields For Automatic Custom Matching In Custom Purchase? Enter Yes Or No (Y/N)": Exit Sub
  '*** (Jen 3.1.2)
  
  '*** (Jen 3.1.0_1)
  Case Is = UCase("HInInvGrsWtEqActual")
    Cancel = Not moCn.RecSeek("select 'x' from Param where PTyp= 'YN' " + _
             "and PMCd = '" + pv_NewValue + "' and PSCd= ''")
    ErrMsg = "Make Invoice Gross Weight Equal To Actual During Copy, Enter Yes Or No (Y/N)?": Exit Sub
  '*** (Jen 3.1.0_1)
  
  '*** Jay 3.1.0_1 [Dsg SetCd From Param]
  Case Is = UCase("HValidDsgSetCd")
    Cancel = Not moCn.RecSeek("select 'x' from Param where PTyp= 'YN' " + _
             "and PMCd = '" + pv_NewValue + "' and PSCd= ''")
    If ADC("HValidDsgSetCd") = "N" And pv_NewValue = "Y" Then
      If MsgBox("All Design Set Codes in Design / Sketch/ Component Master will be Updated By '-'. Proceed Further", vbOKCancel + vbDefaultButton2 + vbExclamation, "Emperor Check") = vbOK Then
        Dim ws_Err As String
        ws_Err = UpdSetCd
        If ws_Err <> "" Then Cancel = True: ErrMsg = ws_Err & " (Error Occurs While Updating Set Code In DsgMst)": ADC("HValidDsgSetCd").SetFocus
        'Dim wStr As String
        'wStr = "If not Exists (Select 'x' from Param where Ptyp = 'SETCD' and PMCd= '-') " + _
        '       "Begin " + _
        '       "     Insert into Param (PTyp, PMCd, PSCd, PDesc, PDesc225, PValue, PValue1, PValue2, PValue3, PNum, PNum1, PValidYn) " + _
        '       "         Values ('SETCD', '-', '', 'Default Design Set Code', '', '', '', '', '', 0, 0, '') " + _
        '       "End " + _
        '       "Update DsgMst Set DmSetCd = '-'"
        'moCn.Execute (wStr)
      Else
        Cancel = True: ErrMsg = "Cannot Set to ''Y''": ADC("HValidDsgSetCd").SetFocus
      End If
      Exit Sub
    End If
  '*** Jay 3.1.0_1 [Dsg SetCd From Param]
  
   '****** Zubin 3.02.01 - 26/03/08
  Case Is = UCase("hChgPtrOnStwYN")
    Cancel = Not moCn.RecSeek("select 'x' from Param where PTyp= 'YN' and PMCd = '" + pv_NewValue + "'")
    If Cancel = True Then ErrMsg = "Specify Whether To Change Pointer On STW? Enter Yes Or No (Y/N)": Exit Sub
  Case Is = UCase("hCmILabWtFrOrd")
    Cancel = Not moCn.RecSeek("select 'x' from Param where PTyp= 'YN' and PMCd = '" + pv_NewValue + "'")
    If Cancel = True Then ErrMsg = "Specify Whether To Fetch Invoice Labor Weight From Order? Enter Yes Or No (Y/N)": Exit Sub
  '****** Zubin 3.02.01 - 26/03/08

  '****** Sachin 3.3.0 (23-06-08) - [14. Invoicing with Markup on Cost of Diamonds] ******
  Case Is = UCase("HINIRTWRTSTKYN")
    Cancel = Not moCn.RecSeek("select 'x' from Param where PTyp= 'YN' and PMCd = '" + pv_NewValue + "'")
    If Cancel = True Then ErrMsg = "Specify Whether Invoicing for Diamonds and Color Stones Has to be Done with Respect to Stock Rate? Enter Yes Or No (Y/N)": Exit Sub
    
  ' ***** Manali 3.6.0 - 27/10/09 - Invoicing with Markup on Cost of Accessories
  Case Is = UCase("hInXIRtwrtStkYN")
    Cancel = Not moCn.RecSeek("select 'x' from Param where PTyp= 'YN' and PMCd = '" + pv_NewValue + "'")
    If Cancel = True Then ErrMsg = "Specify Whether Invoicing for Accessories Has to be Done with Respect to Stock Rate? Enter Yes Or No (Y/N)": Exit Sub
    
  ' ***** Sachin 3.11.0 - Invoicing with Markup on Cost of Metals
  Case Is = UCase("hMetIRtwrtStkYN")
    Cancel = Not moCn.RecSeek("select 'x' from Param where PTyp= 'YN' and PMCd = '" + pv_NewValue + "'")
    If Cancel = True Then ErrMsg = "Specify Whether Invoicing for Metals Has to be Done with Respect to Stock Rate? Enter Yes Or No (Y/N)": Exit Sub
        
  ' ***** Sachin 3.11.0 - Copy Actual/Calculated Wt from Fg
  Case Is = UCase("hCpyActCalcWtYN")
    Cancel = Not moCn.RecSeek("select 'x' from Param where PTyp= 'YN' and PMCd = '" + pv_NewValue + "'")
    If Cancel = True Then ErrMsg = "Specify Whether to Copy the Actual Or Calculated Wt from FG? Enter Y for Actual Or N for Calculated": Exit Sub
    
  '****** Sachin 3.3.0 (23-06-08) - [12. Excel Import In Order] ******
  Case Is = UCase("HEXCELIMPINORDYN"), UCase("HTRAVELEXLYN")
    Cancel = Not moCn.RecSeek("select 'x' from Param where PTyp= 'YN' and PMCd = '" + pv_NewValue + "'")
    If Cancel = True Then ErrMsg = "Specify Whether Excel Import should be allowed in Order Entry? Enter Yes Or No (Y/N)": Exit Sub
    
  ' ***** Manali 3.6.0 - Order Printing in Excel Format Module
  Case Is = UCase("HOrdExcelRepYn")
    Cancel = Not moCn.RecSeek("Select 'x' from Param where PTyp='YN' and PMCd='" + pv_NewValue + "'")
    If Cancel = True Then ErrMsg = "Specify Wheather Order Printing to be allowed in Excel Format": Exit Sub
      
  Case Is = UCase("HCHCDHCOnCustRm")
    Cancel = Not moCn.RecSeek("Select 'x' from Param where PTyp='YN' and PMCd='" + pv_NewValue + "'")
    If Cancel = True Then ErrMsg = "Specify Wheather CHC and DHC charges to be applied on Customer Rm": Exit Sub
        
  ' ****** Manali 3.4.1 - Design Catalogue Module
  Case Is = UCase("HDsgCatYN")
    Cancel = Not moCn.RecSeek("Select 'x' from Param where PTyp='YN' and PMCd='" + pv_NewValue + "'")
    If Cancel = True Then ErrMsg = "Specify Wheather Design Catalogue Should be Implemented or Not. Enter Yes Or No (Y/N)": Exit Sub
  ' ****** Manali 3.4.1 - Design Catalogue Module
  
  '****** Sachin - 3.5.0 - [15/11/08]
  Case Is = UCase("HCustomizeRmCtgExp")
    Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("select PMCd from Param where " + _
                                         "PTyp= 'GLDAS' and PMCd = '" + pv_NewValue + "'"))
    ErrMsg = "Invalid Customize RmCtg Exceptions Value": Exit Sub
  '****** Sachin - 3.5.0 - [15/11/08]
  
  ' ****** Manali 3.5.0 - Max Limit for Dsg View Records
  Case Is = UCase("HMaxDsgViewRec")
    Cancel = (pv_NewValue = 0) Or (pv_NewValue > 9999)
    ErrMsg = "Value can be between 1 to 9999 ": Exit Sub
  ' ****** Manali 3.5.0 - Max Limit for Dsg View Records
  
  Case Is = UCase("hInNotionRtYn")
    Cancel = (Not moCn.RecSeek("select PMCd from Param where " + _
                                      "PTyp= 'YN' and PMCd = '" + pv_NewValue + "'"))
    ErrMsg = "Use Notional Rate in Custom Matching? (Y/N)": Exit Sub
  
  
  ' ***** Manali 3.8.0 - Invoice Design Sort By
  Case Is = UCase("HInSrt")
    Cancel = pv_NewValue <> "" And (Not moCn.RecSeek("select PMCd from Param where " + _
                                      "PTyp= 'INSRT' and PMCd = '" + pv_NewValue + "'"))
    ErrMsg = "Enter Valid Invoice Design Sort By Option Or Keep Blank": Exit Sub

  ' ***** Manali 3.8.0 - Company Logo in Report
  Case Is = UCase("HShowRepLogo")
    Cancel = Not moCn.RecSeek("select PMCd from Param where " + _
                                      "PTyp= 'YN' and PMCd = '" + pv_NewValue + "'")
    ErrMsg = "Enter Y or N to Show Company Logo Reports": Exit Sub

  ' ***** Manali 3.8.0 - Metler Code
  Case Is = UCase("HMetlrCd")
    Cancel = pv_NewValue <> "" And Not moCn.RecSeek("select PMCd from Param where " + _
                                      "PTyp= 'METLR' and PMCd = '" + pv_NewValue + "'")
    ErrMsg = "Enter Valid Metler Code or Keep Blank": Exit Sub

 ' ***** Manali 3.9.1 - 20/02/12 - RfId
  Case Is = UCase("HRfIdMod")
    Cancel = Not moCn.RecSeek("select PMCd from Param where " + _
                               "PTyp= 'YN' and PMCd = '" + pv_NewValue + "'")
    ErrMsg = "Enable RfId Module. Enter 'Y' or 'N' ": Exit Sub
    
 ' ***** Manali 3.9.1 - 20/02/12 - Rate Round Off RmCtgs
  Case Is = UCase("HOrdRtRndOffRmCtg")
    Cancel = pv_NewValue <> "" And Not moCn.RecSeek("select 'x' from Param where " + _
                               "PTyp= 'CPYRT' and PMCd = '" + pv_NewValue + "'")
    ErrMsg = "RmCtg Combination For Order Rate Round Off Not Defined": Exit Sub

  ' **** Manali 3.10.0 - 03/04/12 - MultiPrcQt Module
  Case Is = UCase("HMultiPrcQtMod")
    Cancel = Not moCn.RecSeek("select PMCd from Param where " + _
                               "PTyp= 'YN' and PMCd = '" + pv_NewValue + "'")
    ErrMsg = "Enable Multi Price Quotation Module. Enter 'Y' or 'N' ": Exit Sub

  ' ****** Sachin 3.11.0 - Invoice Rate from Rate Chart ******
  Case Is = UCase("hRtfrRtChrtYN")
    Cancel = Not moCn.RecSeek("select 'x' from Param " + _
                          "where PTyp= 'YN' and PMCd = '" + pv_NewValue + "' and PSCd= ''")
    ErrMsg = "Invoice Rate from Rate Chart, Enter Yes Or No (Y/N)?": Exit Sub
  
  ' ****** Sachin 3.11.0 - Setting Type Only from Order YN ******
  Case Is = UCase("hSetTypOnlyFrOrdYN")
    Cancel = Not moCn.RecSeek("select 'x' from Param " + _
                          "where PTyp= 'YN' and PMCd = '" + pv_NewValue + "' and PSCd= ''")
    ErrMsg = "Setting Type only from Order, Enter Yes Or No (Y/N)?": Exit Sub

 Case Is = UCase("hTolWtPctDC")         ' Sachin 3.11.0 - Wt. Tolerance in Percent for D or C
    ws_RmCtgArr = Split(pv_NewValue, ",")
    If IsArray(ws_RmCtgArr) Then
      For j = 0 To UBound(ws_RmCtgArr)
        Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek(" Select PMCd from Param Where " + _
                                             " PTyp= 'RMCTG' and PMCd In ('D','C') " + _
                                             " and PMCd= '" + ws_RmCtgArr(j) + "'"))
        If Cancel = True Then ErrMsg = "Value(" + CStr(j + 1) + ") Is An Invalid Rm Ctg. Only 'D' Or 'C' Or 'D,C' can be Entered.": Exit Sub
      Next j
    End If

 Case Is = UCase("hTolQtyPctDC")        ' Sachin 3.11.0 - Qty Tolerance in Percent for D or C
    ws_RmCtgArr = Split(pv_NewValue, ",")
    If IsArray(ws_RmCtgArr) Then
      For j = 0 To UBound(ws_RmCtgArr)
        Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek(" Select PMCd from Param Where " + _
                                             " PTyp= 'RMCTG' and PMCd In ('D','C') " + _
                                             " and PMCd= '" + ws_RmCtgArr(j) + "'"))
        If Cancel = True Then ErrMsg = "Value(" + CStr(j + 1) + ") Is An Invalid Rm Ctg. Only 'D' Or 'C' Or 'D,C' can be Entered.": Exit Sub
      Next j
    End If
    
 'Sachin - 4.1.0.0
 Case Is = UCase("hShowBgHistDetsYN")
    Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("select PMCd from Param where " + _
                                                  "PTyp= 'YN' and PMCd = '" + pv_NewValue + "'"))
    ErrMsg = "Show Details in Bag Reports (Y/N)": Exit Sub
    
 'Sachin - 4.1.2.0
 Case Is = UCase("HHolChkYN")
    Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("select PMCd from Param where " + _
                                                  "PTyp= 'YN' and PMCd = '" + pv_NewValue + "'"))
    ErrMsg = "Apply Checks for Holidays (Y/N)": Exit Sub
    
 Case Is = UCase("HEnfFrBLocInFgYN")
    Cancel = (Not moCn.RecSeek("select PMCd from Param where " + _
                                                  "PTyp= 'YN' and PMCd = '" + pv_NewValue + "'"))
    ErrMsg = "Enforce From Bag Location in FG Entry (Y/N)": Exit Sub
    
 'Sachin - 4.1.3.0
 Case Is = UCase("hPrdPtsOnBQtyYN")
    Cancel = (Not moCn.RecSeek("select PMCd from Param where " + _
                                                  "PTyp= 'YN' and PMCd = '" + pv_NewValue + "'"))
    ErrMsg = "Restrict Production Points On Bag Qty, Enter Yes Or No (Y/N)?": Exit Sub
 
  Case Is = UCase("HJSTKDEFCHR")
    Cancel = (Not moCn.RecSeek("select PMCd from Param where " + _
                                                  "PTyp= 'BCHR' and PMCd = '" + pv_NewValue + "' and PValue='P'"))
    ErrMsg = "Please Enter valid Production Bag Character": Exit Sub
 
 Case Is = UCase("HJSTKNUMJFRBAGYN")
    Cancel = (Not moCn.RecSeek("select PMCd from Param where " + _
                                                  "PTyp= 'YN' and PMCd = '" + pv_NewValue + "'"))
    ErrMsg = "Please enter Y or N": Exit Sub
 Case Is = UCase("HRMSTKBY")
    If (pv_NewValue < 0 Or pv_NewValue > 3) Then
      Cancel = True
      ErrMsg = "Please enter a value greater than zero and less than 3"
    End If
    If moCn.RecSeek("Select 'a' from Param where PTyp = 'RMSCTG' and PNum1<>1 and PNum1 > " + CStr(pv_NewValue)) Then
      Cancel = True
      ErrMsg = "Cannot modify, there are Rm Sub Categories that have no of Sizes set to a greater number"
    End If
  'uma *** Repair Bag Loc
  Case Is = UCase("HDefRepBagLoc")
    Cancel = (pv_NewValue <> "" And (Not moCn.RecSeek("select LocCd from Loc where " + _
                                                  "LocTyp= 'P' and LocCd = '" + pv_NewValue + "'")))
    ErrMsg = "Enter Default Location for Repair Bag": Exit Sub
  'EmrSv - Uma
  Case Is = UCase("hCell")
    Cancel = (Not moCn.RecSeek("select PMCD from Param where " + _
                                                  "PTyp= 'Cell' and PMCd = '" + pv_NewValue + "'"))
    ErrMsg = "Enter Valid Cell": Exit Sub
  'StkRt-Validation for Stock rate NoEdit & NoDisplay. It allow only Y or N. Blank not allowed
  Case Is = UCase("HSTKRTNOEDIT")
    Cancel = (Not moCn.RecSeek("select PMCd from Param where " + _
                                                  "PTyp= 'YN' and PMCd = '" + pv_NewValue + "'"))
    ErrMsg = "Check Stock Rate NOEDIT (Y/N)": Exit Sub
  Case Is = UCase("HSTKRTNODSP")
    Cancel = (Not moCn.RecSeek("select PMCd from Param where " + _
                                                  "PTyp= 'YN' and PMCd = '" + pv_NewValue + "'"))
    ErrMsg = "Check Stock Rate NODISPLAY (Y/N)": Exit Sub
  'CostRtPDCM-Validation for Cost from PDCM Y/N, Blank not allowed
  Case Is = UCase("hCostFrPdcmYN")
    Cancel = (Not moCn.RecSeek("select PMCd from Param where " + _
                                                  "PTyp= 'YN' and PMCd = '" + pv_NewValue + "'"))
    ErrMsg = "Check Cost Rate from PDCM (Y/N)": Exit Sub
  'mw.84-validating OrdRm Lot YN
  Case Is = UCase("HOrLotYN")
    Cancel = Not moCn.RecSeek("select PMCd from Param where PTyp= 'YN' and PMCd = '" + pv_NewValue + "'")
    ErrMsg = "Check Implement OrdRm Lot (Y/N)": Exit Sub
End Select
End Sub
Private Sub ADC_SetRecSource()
  '*** Set the RecSource for ADC
  ADC.RecSource = "Select * from Head Where HCoCd ='" + ADC("HCoCd") + "' and HCd ='" + ADC("HCd") + "' "
End Sub

Private Sub CmdRefFileLst_Click()
  ' ***** Manali 3.5.0 - 09/06/09 - Refresh Pic File List
  '*** Display or hide the RefFileLst frame depending on the toggle satate
  Call DispFra(RefFileLst)
End Sub

Private Sub CmdRefFileLstGo_Click()
  ' ***** Manali 3.5.0 - 09/06/09 - Refresh Pic File List
  Call GenFileList
  Call DispFra(RefFileLst)
End Sub

Private Sub TabHd_Click(PreviousTab As Integer)
  '*** Whenever a tab is selected enable the frame on the clicked tab and disable the previous tab frame
  FraTabHd(TabHd.Tab).Enabled = True
  FraTabHd(PreviousTab).Enabled = False
End Sub
Private Function UpdSetCd() As String
  On Error GoTo ErrUpdSetCd
  
  UpdSetCd = ""
  Dim wStr As String
  wStr = "If not Exists (Select 'x' from Param where Ptyp = 'SETCD' and PMCd= '-') " + _
         "Begin " + _
         "     Insert into Param (PTyp, PMCd, PSCd, PDesc, PDesc225, PValue, PValue1, PValue2, PValue3, PNum, PNum1, PValidYn) " + _
         "         Values ('SETCD', '-', '', 'Default Design Set Code', '', '', '', '', '', 0, 0, '') " + _
         "End " + _
         "Update DsgMst Set DmSetCd = '-'"
  moCn.Execute (wStr)
  UpdSetCd = ""
  Exit Function
  
ErrUpdSetCd:
  'DispMsg Err.Description, etError
  UpdSetCd = Err.Description
End Function

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
  ws_BtnStr = ""
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
    Set moCn = Nothing    '*** (09/08/05)    '*** (Bef speed) bef 24/11/06
End Sub
Private Sub HideAllFras()
  ' ***** Manali 3.5.0 - 09/06/09
  '*** This procedure hides all the frames other than FraNKeyAll
  FraNKeyAll.Visible = True
  FraRefFileLst.Visible = False
  FraRefFileLst.Enabled = False
End Sub

Private Sub EnaDisaCmds(ByVal pv_ShowFra As Boolean, Optional ByVal pv_CmdBut As MWCTL_BTN1)
  ' ***** Manali 3.5.0 - 09/06/09
  '*** If second parameter is not given and pv_ShowFra = True then the frame FraNKeyAll
      'on Frame Fra_nkeys is disabled and allowadd and allowdelete options of adc are
      'set to false

  '*** If second parameter is not given and pv_ShowFra = False then the frame FraNKeyAll
      'on Frame Fra_nkeys is enabled and allowadd and allowdelete options of adc are
      'set to true

  '*** If second parameter is given and pv_ShowFra = True then the frame FraNKeyAll
      'on Frame Fra_nkeys is disabled and allowadd and allowdelete options of adc are
      'set to false and for the command button specified in the second parameter,
      'the BackColor is changed to white and the font is made Bold

  '*** If second parameter is given and pv_ShowFra = False then the frame FraNKeyAll
      'on Frame Fra_nkeys is enabled and allowadd and allowdelete options of adc are
      'set to true and for the command button specified in the second parameter,
      'the BackColor is changed to normal and the font is made Regular


  If Not pv_CmdBut Is Nothing Then
    If pv_ShowFra = True Then
      '*** Jenny (Old Code Bef Color)
      'pv_CmdBut.BackColor = vbWhite
      'pv_CmdBut.FontBold = True
      '*** Jenny (Old Code Bef Color)
      If pv_CmdBut.Pic = 0 Then
        pv_CmdBut.BackColor = vbWhite
        'pv_CmdBut.FontBold = True
      End If
      pv_CmdBut.CausesValidation = True
    Else
      '*** Jenny (Old Code Bef Color)
      'pv_CmdBut.BackColor = &H8000000F
      'pv_CmdBut.FontBold = False
      '*** Jenny (Old Code Bef Color)
      If pv_CmdBut.Pic = 0 Then
        pv_CmdBut.BackColor = &H8000000F
        'pv_CmdBut.FontBold = False
      End If
      'pv_CmdBut.CausesValidation = False
    End If
  End If
  FraNKeyAll.Enabled = Not pv_ShowFra
  ADC.AllowSave = Not pv_ShowFra
  ADC.AllowDelete = Not pv_ShowFra
  CmdRefFileLst.Enabled = Not pv_ShowFra And mb_DsgCat      ' ***** Manali 3.5.0 - 09/06/09 - Refresh Pic File List
  If Not pv_CmdBut Is Nothing Then pv_CmdBut.Enabled = True
End Sub

Private Sub DispFra(ByVal pv_HeadFra As en_HeadFra)
' ***** Manali 3.5.0 - 09/06/09
  Select Case pv_HeadFra
  Case Is = RefFileLst      ' ***** Manali 3.5.0 - 09/06/09 - Refresh Pic File List
    '*** If Frame FraRefFileLst is visible then make it invisible and enable all the command buttons
    '*** If Frame FraRefFileLst is invisible then make it visible and disable all the
        'command buttons except the one which is clicked. Change the background colour of
        'the clicked command button to white and make the font bold
    If FraRefFileLst.Visible = True Then
      FraRefFileLst.Visible = False
      FraRefFileLst.Enabled = False
      Call EnaDisaCmds(False, CmdRefFileLst)
      CmdRefFileLst.SetFocus
    Else
      FraRefFileLst.Visible = True
      FraRefFileLst.Enabled = True
      CmdRefFileLstGo.SetFocus
      FraRefFileLst.ZOrder
      Call EnaDisaCmds(True, CmdRefFileLst)
    End If
  End Select
End Sub

Private Sub GenFileList()
 '****** Manali 3.5.0- 09/06/09 - Filling the Picture Files List in Table  ******
 On Error GoTo errhld
 Dim wfso As New FileSystemObject, wFile As file, wRsDmCtg As MwfLib.MDORowSet, wPctPath As String
  
 MWLib.BeginProcess Me, "Refrshing Picture File Names ..."
  
 moCn.BeginTrans (mdoCommitRec)

 Me.MousePointer = vbHourglass
  
  moCn.Execute ("Delete from PictFileLst")
  wPctPath = moCn.GetFldVal("Select PDesc225 from Param Where PTyp='PCT' and PMCd='3D'")
  If Not wfso.FolderExists(wPctPath) Then
    moCn.CommitTrans mdoCommitNone
    MWLib.EndProcess Me
    Me.MousePointer = vbDefault
    DispMsg " Folder Does Not Exist ", etInfo:  Exit Sub  '****** Exit Sub if Path does not exist
  End If
  Set wRsDmCtg = moCn.OpenRes("Select PMCd From Param Where PTyp='DMCTG'")
    With wRsDmCtg
      Do While Not (.EOF Or .BOF)
        wPctPath = moCn.GetFldVal("Select PDesc225 from Param Where PTyp='PCT' and PMCd='3D'") + "\" + !PMCd
        If wfso.FolderExists(wPctPath) Then
            For Each wFile In wfso.GetFolder(wPctPath).Files
              If Right(wFile.Name, 4) = ".jpg" Then
                moCn.Execute (" Insert Into PictFileLst (PfDmCtg, PfFileName, ModUsr, ModDt, ModTime) Values " + _
                              " ('" + !PMCd + "','" + wFile.Name + "', '" + UCase(gs_UsrCd) + "', " + _
                              " getdate(), cast(DATEPART(hh, GetDate())as varchar)+'.' +cast(DATEPART(mi, GetDate())as varchar) ) ")
              End If
            Next
        End If
        .MoveNext
      Loop
    End With

  moCn.CommitTrans mdoCommitRec
  
  MWLib.EndProcess Me
  Me.MousePointer = vbDefault
  DispMsg "Picture File List is Refreshed", etInfo
  
Exit Sub
errhld:
    DispMsg Err.Description, etError
    MWLib.EndProcess Me
    Me.MousePointer = vbDefault
 '****** Manali 3.5.0- 09/06/09 - Filling the Picture Files List in Table  ******
End Sub


