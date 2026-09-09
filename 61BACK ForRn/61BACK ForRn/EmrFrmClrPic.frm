VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#2.0#0"; "Mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{B7333271-A4A6-4B6F-B70D-C597AB683282}#1.0#0"; "mwfctl.ocx"
Begin VB.Form EmrFrmClrPic 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Colour Selection"
   ClientHeight    =   8220
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12540
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8220
   ScaleMode       =   0  'User
   ScaleWidth      =   12540
   WindowState     =   2  'Maximized
   Begin MwfCtl.MWCTL_BTN1 CmdCpy 
      Height          =   435
      Left            =   5190
      TabIndex        =   45
      ToolTipText     =   "Copy From Another Template"
      Top             =   7530
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   767
      BackColor       =   -2147483633
      DisableColor    =   8421504
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
   Begin VB.Frame FRA_KEYS 
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   0
      TabIndex        =   106
      Top             =   30
      Width           =   5055
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   1
         Left            =   1185
         TabIndex        =   0
         ToolTipText     =   "Enter Template Code"
         Top             =   0
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   503
         MaxLength       =   5
         DataType        =   4
         DataField       =   "CsCd"
         IdName          =   "CSCD"
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Template"
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
         TabIndex        =   107
         Top             =   15
         Width           =   1065
      End
   End
   Begin MwfCtl.MWCTL_HLP MWCTL_HLP1 
      Height          =   405
      Left            =   8010
      TabIndex        =   103
      TabStop         =   0   'False
      Top             =   7635
      Visible         =   0   'False
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   714
   End
   Begin MwfCtl.MWCTL_ADC ADC 
      CausesValidation=   0   'False
      Height          =   555
      Left            =   60
      TabIndex        =   102
      Top             =   7500
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   979
      BtnWidth        =   1000
      BtnHeight       =   435
      BtnMaskColor    =   8454143
      BtnDisableColor =   8421504
      FontSz          =   8.25
   End
   Begin VB.Frame FRA_NKEYS 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   7755
      Left            =   -120
      TabIndex        =   105
      Top             =   0
      Width           =   12525
      Begin VB.Frame PicBoxCpy 
         Height          =   1065
         Left            =   3630
         TabIndex        =   108
         Top             =   6150
         Width           =   4515
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   7
            Left            =   1095
            TabIndex        =   104
            ToolTipText     =   "Enter Template Code"
            Top             =   510
            Width           =   1050
            _ExtentX        =   1852
            _ExtentY        =   503
            MaxLength       =   5
            DataType        =   4
            IdName          =   "WCPYCSCD"
         End
         Begin MwfCtl.MWCTL_BTN1 CmdCpyGo 
            Height          =   495
            Left            =   2520
            TabIndex        =   112
            ToolTipText     =   "Copy From Another Template"
            Top             =   420
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   873
            BackColor       =   -2147483633
            DisableColor    =   8421504
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
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "  Copy From Template"
            Height          =   255
            Left            =   -30
            TabIndex        =   111
            Top             =   90
            Width           =   4545
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Template"
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
            Index           =   63
            Left            =   90
            TabIndex        =   110
            Top             =   525
            Width           =   1065
         End
      End
      Begin VB.Frame FraNKeyAll 
         BorderStyle     =   0  'None
         Height          =   7335
         Left            =   150
         TabIndex        =   113
         Top             =   10
         Width           =   12015
         Begin TabDlg.SSTab TabDet 
            Height          =   6600
            Left            =   30
            TabIndex        =   3
            Top             =   600
            Width           =   11910
            _ExtentX        =   21008
            _ExtentY        =   11642
            _Version        =   393216
            TabHeight       =   520
            ForeColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            TabCaption(0)   =   "<&1> Color Details"
            TabPicture(0)   =   "EmrFrmClrPic.frx":0000
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "PicBoxTabDet(0)"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).ControlCount=   1
            TabCaption(1)   =   "<&2> Color Details Contd."
            TabPicture(1)   =   "EmrFrmClrPic.frx":001C
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "PicBoxTabDet(1)"
            Tab(1).Control(0).Enabled=   0   'False
            Tab(1).ControlCount=   1
            TabCaption(2)   =   "<&3> Color Details Contd."
            TabPicture(2)   =   "EmrFrmClrPic.frx":0038
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "PicBoxTabDet(2)"
            Tab(2).Control(0).Enabled=   0   'False
            Tab(2).ControlCount=   1
            Begin VB.Frame PicBoxTabDet 
               BorderStyle     =   0  'None
               Height          =   6210
               Index           =   0
               Left            =   60
               TabIndex        =   4
               Top             =   360
               Width           =   11825
               Begin VB.Frame Picture8 
                  Height          =   6375
                  Left            =   3900
                  TabIndex        =   182
                  Top             =   -120
                  Width           =   4005
                  Begin VB.PictureBox PicBoxCsFlxPic 
                     Height          =   705
                     Left            =   2040
                     ScaleHeight     =   645
                     ScaleWidth      =   1695
                     TabIndex        =   183
                     Top             =   5310
                     Width           =   1755
                  End
                  Begin ComctlLib.ImageCombo ImgCombo 
                     DataField       =   "CsFlxBckClr"
                     Height          =   330
                     Index           =   5
                     Left            =   2040
                     TabIndex        =   17
                     Top             =   450
                     Width           =   1755
                     _ExtentX        =   3096
                     _ExtentY        =   582
                     _Version        =   393216
                     ForeColor       =   -2147483640
                     BackColor       =   -2147483643
                  End
                  Begin ComctlLib.ImageCombo ImgCombo 
                     DataField       =   "CsFlxForClr"
                     Height          =   330
                     Index           =   6
                     Left            =   2040
                     TabIndex        =   18
                     Top             =   780
                     Width           =   1755
                     _ExtentX        =   3096
                     _ExtentY        =   582
                     _Version        =   393216
                     ForeColor       =   -2147483640
                     BackColor       =   -2147483643
                  End
                  Begin ComctlLib.ImageCombo ImgCombo 
                     DataField       =   "CsFlxBckClrSelRow"
                     Height          =   330
                     Index           =   7
                     Left            =   2040
                     TabIndex        =   19
                     Top             =   1110
                     Width           =   1755
                     _ExtentX        =   3096
                     _ExtentY        =   582
                     _Version        =   393216
                     ForeColor       =   -2147483640
                     BackColor       =   -2147483643
                  End
                  Begin ComctlLib.ImageCombo ImgCombo 
                     DataField       =   "CsFlxBckClrSelTxt"
                     Height          =   330
                     Index           =   8
                     Left            =   2040
                     TabIndex        =   20
                     Top             =   1440
                     Width           =   1755
                     _ExtentX        =   3096
                     _ExtentY        =   582
                     _Version        =   393216
                     ForeColor       =   -2147483640
                     BackColor       =   -2147483643
                  End
                  Begin ComctlLib.ImageCombo ImgCombo 
                     DataField       =   "CsFlxForClrSelTxt"
                     Height          =   330
                     Index           =   9
                     Left            =   2040
                     TabIndex        =   21
                     Top             =   1770
                     Width           =   1755
                     _ExtentX        =   3096
                     _ExtentY        =   582
                     _Version        =   393216
                     ForeColor       =   -2147483640
                     BackColor       =   -2147483643
                  End
                  Begin ComctlLib.ImageCombo ImgCombo 
                     DataField       =   "CsFlxBckClrFix"
                     Height          =   330
                     Index           =   10
                     Left            =   2040
                     TabIndex        =   22
                     Top             =   2100
                     Width           =   1755
                     _ExtentX        =   3096
                     _ExtentY        =   582
                     _Version        =   393216
                     ForeColor       =   -2147483640
                     BackColor       =   -2147483643
                  End
                  Begin ComctlLib.ImageCombo ImgCombo 
                     DataField       =   "CsFlxForClrFix"
                     Height          =   330
                     Index           =   11
                     Left            =   2040
                     TabIndex        =   23
                     Top             =   2430
                     Width           =   1755
                     _ExtentX        =   3096
                     _ExtentY        =   582
                     _Version        =   393216
                     ForeColor       =   -2147483640
                     BackColor       =   -2147483643
                  End
                  Begin ComctlLib.ImageCombo ImgCombo 
                     DataField       =   "CsFlxBckClrBkg"
                     Height          =   330
                     Index           =   12
                     Left            =   2040
                     TabIndex        =   24
                     Top             =   2760
                     Width           =   1755
                     _ExtentX        =   3096
                     _ExtentY        =   582
                     _Version        =   393216
                     ForeColor       =   -2147483640
                     BackColor       =   -2147483643
                  End
                  Begin ComctlLib.ImageCombo ImgCombo 
                     DataField       =   "CsFlxBckClrSel"
                     Height          =   330
                     Index           =   13
                     Left            =   2040
                     TabIndex        =   25
                     Top             =   3090
                     Width           =   1755
                     _ExtentX        =   3096
                     _ExtentY        =   582
                     _Version        =   393216
                     ForeColor       =   -2147483640
                     BackColor       =   -2147483643
                  End
                  Begin ComctlLib.ImageCombo ImgCombo 
                     DataField       =   "CsFlxForClrSel"
                     Height          =   330
                     Index           =   14
                     Left            =   2040
                     TabIndex        =   26
                     Top             =   3420
                     Width           =   1755
                     _ExtentX        =   3096
                     _ExtentY        =   582
                     _Version        =   393216
                     ForeColor       =   -2147483640
                     BackColor       =   -2147483643
                  End
                  Begin ComctlLib.ImageCombo ImgCombo 
                     DataField       =   "CsFlxGrdClr"
                     Height          =   330
                     Index           =   15
                     Left            =   2040
                     TabIndex        =   27
                     Top             =   3750
                     Width           =   1755
                     _ExtentX        =   3096
                     _ExtentY        =   582
                     _Version        =   393216
                     ForeColor       =   -2147483640
                     BackColor       =   -2147483643
                  End
                  Begin ComctlLib.ImageCombo ImgCombo 
                     DataField       =   "CsFlxGrdClrFix"
                     Height          =   330
                     Index           =   16
                     Left            =   2040
                     TabIndex        =   28
                     Top             =   4080
                     Width           =   1755
                     _ExtentX        =   3096
                     _ExtentY        =   582
                     _Version        =   393216
                     ForeColor       =   -2147483640
                     BackColor       =   -2147483643
                  End
                  Begin ComctlLib.ImageCombo ImgCombo 
                     DataField       =   "CsFlxGrdClrFix1"
                     Height          =   330
                     Index           =   17
                     Left            =   2040
                     TabIndex        =   29
                     Top             =   4410
                     Width           =   1755
                     _ExtentX        =   3096
                     _ExtentY        =   582
                     _Version        =   393216
                     ForeColor       =   -2147483640
                     BackColor       =   -2147483643
                  End
                  Begin MwfCtl.MWCTL_BTN1 CmdGrdPointPic 
                     Height          =   435
                     Left            =   60
                     TabIndex        =   32
                     ToolTipText     =   "Get Grid Pointer Picture"
                     Top             =   5310
                     Width           =   1185
                     _ExtentX        =   2090
                     _ExtentY        =   767
                     BackColor       =   -2147483633
                     Caption         =   "&Ptr Pic"
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
                     Index           =   15
                     Left            =   2040
                     TabIndex        =   30
                     ToolTipText     =   "Specify Whether Heading Of Grid Objects Should Appear As Inset Or Flat ('Y' For Inset)"
                     Top             =   4740
                     Width           =   375
                     _ExtentX        =   661
                     _ExtentY        =   503
                     DataType        =   4
                     DataField       =   "CsGrdInset"
                     IdName          =   "CSGRDINSET"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   13
                     Left            =   2040
                     TabIndex        =   31
                     ToolTipText     =   "Enter Grid Row Pointer Picture Path"
                     Top             =   5025
                     Width           =   1755
                     _ExtentX        =   3096
                     _ExtentY        =   503
                     MaxLength       =   500
                     DataField       =   "CsFlxPic"
                     IdName          =   "CSFLXPIC"
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Heading Line Color2"
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
                     Height          =   330
                     Index           =   19
                     Left            =   60
                     TabIndex        =   199
                     Top             =   4410
                     Width           =   1995
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Heading Line Color1"
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
                     Height          =   330
                     Index           =   18
                     Left            =   60
                     TabIndex        =   198
                     Top             =   4080
                     Width           =   1995
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Cell Line Color"
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
                     Height          =   330
                     Index           =   17
                     Left            =   60
                     TabIndex        =   197
                     Top             =   3750
                     Width           =   1995
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Fore Color Selection"
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
                     Height          =   330
                     Index           =   16
                     Left            =   60
                     TabIndex        =   196
                     Top             =   3420
                     Width           =   1995
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Back Color Selection"
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
                     Height          =   330
                     Index           =   15
                     Left            =   60
                     TabIndex        =   195
                     Top             =   3090
                     Width           =   1995
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Back Color"
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
                     Height          =   330
                     Index           =   14
                     Left            =   60
                     TabIndex        =   194
                     Top             =   2760
                     Width           =   1995
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Heading Fore Color"
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
                     Height          =   330
                     Index           =   13
                     Left            =   60
                     TabIndex        =   193
                     Top             =   2430
                     Width           =   1995
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Heading Back Color"
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
                     Height          =   330
                     Index           =   12
                     Left            =   60
                     TabIndex        =   192
                     Top             =   2100
                     Width           =   1995
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Cell Back Color"
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
                     Height          =   330
                     Index           =   11
                     Left            =   60
                     TabIndex        =   191
                     Top             =   450
                     Width           =   1995
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Cell Fore Color"
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
                     Height          =   330
                     Index           =   10
                     Left            =   60
                     TabIndex        =   190
                     Top             =   780
                     Width           =   1995
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Sel Row Back Color"
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
                     Height          =   330
                     Index           =   9
                     Left            =   60
                     TabIndex        =   189
                     Top             =   1110
                     Width           =   1995
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Sel Cell Back Color"
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
                     Height          =   330
                     Index           =   8
                     Left            =   60
                     TabIndex        =   188
                     Top             =   1440
                     Width           =   1995
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Sel Cell Fore Color"
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
                     Height          =   330
                     Index           =   7
                     Left            =   60
                     TabIndex        =   187
                     Top             =   1770
                     Width           =   1995
                  End
                  Begin VB.Label Label1 
                     BackStyle       =   0  'Transparent
                     BorderStyle     =   1  'Fixed Single
                     Caption         =   "  Color Setting For Grid Control"
                     Height          =   255
                     Left            =   -30
                     TabIndex        =   186
                     Top             =   90
                     Width           =   4035
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Grid Inset (Y/ N)"
                     BeginProperty Font 
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
                     Left            =   60
                     TabIndex        =   185
                     Top             =   4740
                     Width           =   1635
                  End
                  Begin VB.Label ALBL 
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
                     Index           =   68
                     Left            =   60
                     TabIndex        =   184
                     Top             =   5025
                     Width           =   1995
                  End
               End
               Begin VB.Frame Picture10 
                  BorderStyle     =   0  'None
                  Height          =   5955
                  Left            =   7860
                  TabIndex        =   217
                  Top             =   -120
                  Width           =   3975
                  Begin VB.PictureBox PicBoxCsHlpPic 
                     Height          =   675
                     Left            =   2070
                     ScaleHeight     =   615
                     ScaleWidth      =   1695
                     TabIndex        =   218
                     Top             =   4035
                     Width           =   1755
                  End
                  Begin ComctlLib.ImageCombo ImgCombo 
                     DataField       =   "CsHlpBckClr"
                     Height          =   330
                     Index           =   21
                     Left            =   2070
                     TabIndex        =   33
                     Top             =   450
                     Width           =   1755
                     _ExtentX        =   3096
                     _ExtentY        =   582
                     _Version        =   393216
                     ForeColor       =   -2147483640
                     BackColor       =   -2147483643
                  End
                  Begin ComctlLib.ImageCombo ImgCombo 
                     DataField       =   "CsHlpForClr"
                     Height          =   330
                     Index           =   22
                     Left            =   2070
                     TabIndex        =   34
                     Top             =   780
                     Width           =   1755
                     _ExtentX        =   3096
                     _ExtentY        =   582
                     _Version        =   393216
                     ForeColor       =   -2147483640
                     BackColor       =   -2147483643
                  End
                  Begin ComctlLib.ImageCombo ImgCombo 
                     DataField       =   "CsHlpBckClrSelRow"
                     Height          =   330
                     Index           =   23
                     Left            =   2070
                     TabIndex        =   35
                     Top             =   1110
                     Width           =   1755
                     _ExtentX        =   3096
                     _ExtentY        =   582
                     _Version        =   393216
                     ForeColor       =   -2147483640
                     BackColor       =   -2147483643
                  End
                  Begin ComctlLib.ImageCombo ImgCombo 
                     DataField       =   "CsHlpBckClrFix"
                     Height          =   330
                     Index           =   27
                     Left            =   2070
                     TabIndex        =   36
                     Top             =   1440
                     Width           =   1755
                     _ExtentX        =   3096
                     _ExtentY        =   582
                     _Version        =   393216
                     ForeColor       =   -2147483640
                     BackColor       =   -2147483643
                  End
                  Begin ComctlLib.ImageCombo ImgCombo 
                     DataField       =   "CsHlpForClrFix"
                     Height          =   330
                     Index           =   28
                     Left            =   2070
                     TabIndex        =   37
                     Top             =   1770
                     Width           =   1755
                     _ExtentX        =   3096
                     _ExtentY        =   582
                     _Version        =   393216
                     ForeColor       =   -2147483640
                     BackColor       =   -2147483643
                  End
                  Begin ComctlLib.ImageCombo ImgCombo 
                     DataField       =   "CsHlpBckClrBkg"
                     Height          =   330
                     Index           =   29
                     Left            =   2070
                     TabIndex        =   38
                     Top             =   2100
                     Width           =   1755
                     _ExtentX        =   3096
                     _ExtentY        =   582
                     _Version        =   393216
                     ForeColor       =   -2147483640
                     BackColor       =   -2147483643
                  End
                  Begin ComctlLib.ImageCombo ImgCombo 
                     DataField       =   "CsHlpBckClrSel"
                     Height          =   330
                     Index           =   30
                     Left            =   2070
                     TabIndex        =   39
                     Top             =   2430
                     Width           =   1755
                     _ExtentX        =   3096
                     _ExtentY        =   582
                     _Version        =   393216
                     ForeColor       =   -2147483640
                     BackColor       =   -2147483643
                  End
                  Begin ComctlLib.ImageCombo ImgCombo 
                     DataField       =   "CsHlpForClrSel"
                     Height          =   330
                     Index           =   31
                     Left            =   2070
                     TabIndex        =   40
                     Top             =   2760
                     Width           =   1755
                     _ExtentX        =   3096
                     _ExtentY        =   582
                     _Version        =   393216
                     ForeColor       =   -2147483640
                     BackColor       =   -2147483643
                  End
                  Begin ComctlLib.ImageCombo ImgCombo 
                     DataField       =   "CsHlpGrdClr"
                     Height          =   330
                     Index           =   32
                     Left            =   2070
                     TabIndex        =   41
                     Top             =   3090
                     Width           =   1755
                     _ExtentX        =   3096
                     _ExtentY        =   582
                     _Version        =   393216
                     ForeColor       =   -2147483640
                     BackColor       =   -2147483643
                  End
                  Begin ComctlLib.ImageCombo ImgCombo 
                     DataField       =   "CsHlpGrdClrFix"
                     Height          =   330
                     Index           =   33
                     Left            =   2070
                     TabIndex        =   42
                     Top             =   3420
                     Width           =   1755
                     _ExtentX        =   3096
                     _ExtentY        =   582
                     _Version        =   393216
                     ForeColor       =   -2147483640
                     BackColor       =   -2147483643
                  End
                  Begin MwfCtl.MWCTL_BTN1 CmdHlpPointPic 
                     Height          =   435
                     Left            =   60
                     TabIndex        =   44
                     ToolTipText     =   "Get Help Pointer Picture"
                     Top             =   4035
                     Width           =   1185
                     _ExtentX        =   2090
                     _ExtentY        =   767
                     BackColor       =   -2147483633
                     Caption         =   "P&tr Pic"
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
                     Index           =   25
                     Left            =   2070
                     TabIndex        =   43
                     ToolTipText     =   "Enter Help Grid Row Pointer Picture Path"
                     Top             =   3750
                     Width           =   1755
                     _ExtentX        =   3096
                     _ExtentY        =   503
                     MaxLength       =   500
                     DataField       =   "CsHlpPic"
                     IdName          =   "CSHLPPIC"
                  End
                  Begin VB.Label Label4 
                     BackStyle       =   0  'Transparent
                     BorderStyle     =   1  'Fixed Single
                     Caption         =   "  Color Setting For Help Control"
                     Height          =   255
                     Left            =   -30
                     TabIndex        =   230
                     Top             =   90
                     Width           =   4005
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Sel Row Back Color"
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
                     Height          =   330
                     Index           =   38
                     Left            =   60
                     TabIndex        =   229
                     Top             =   1110
                     Width           =   1995
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Cell Fore Color"
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
                     Height          =   330
                     Index           =   37
                     Left            =   60
                     TabIndex        =   228
                     Top             =   780
                     Width           =   1995
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Cell Back Color"
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
                     Height          =   330
                     Index           =   36
                     Left            =   60
                     TabIndex        =   227
                     Top             =   450
                     Width           =   1995
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Heading Back Color"
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
                     Height          =   330
                     Index           =   35
                     Left            =   60
                     TabIndex        =   226
                     Top             =   1440
                     Width           =   1995
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Heading Fore Color"
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
                     Height          =   330
                     Index           =   34
                     Left            =   60
                     TabIndex        =   225
                     Top             =   1770
                     Width           =   1995
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Back Color"
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
                     Height          =   330
                     Index           =   33
                     Left            =   60
                     TabIndex        =   224
                     Top             =   2100
                     Width           =   1995
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Back Color Selection"
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
                     Height          =   330
                     Index           =   32
                     Left            =   60
                     TabIndex        =   223
                     Top             =   2430
                     Width           =   1995
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Fore Color Selection"
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
                     Height          =   330
                     Index           =   31
                     Left            =   60
                     TabIndex        =   222
                     Top             =   2760
                     Width           =   1995
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Cell Line Color"
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
                     Height          =   330
                     Index           =   30
                     Left            =   60
                     TabIndex        =   221
                     Top             =   3090
                     Width           =   1995
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Heading Line Color"
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
                     Height          =   330
                     Index           =   29
                     Left            =   60
                     TabIndex        =   220
                     Top             =   3420
                     Width           =   1995
                  End
                  Begin VB.Label ALBL 
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
                     Index           =   80
                     Left            =   60
                     TabIndex        =   219
                     Top             =   3750
                     Width           =   1995
                  End
               End
               Begin VB.Frame Picture9 
                  Height          =   4035
                  Left            =   -30
                  TabIndex        =   207
                  Top             =   2220
                  Width           =   3975
                  Begin ComctlLib.ImageCombo ImgCombo 
                     DataField       =   "CsADCBtnForClr"
                     Height          =   330
                     Index           =   18
                     Left            =   2040
                     TabIndex        =   10
                     Top             =   450
                     Width           =   1755
                     _ExtentX        =   3096
                     _ExtentY        =   582
                     _Version        =   393216
                     ForeColor       =   -2147483640
                     BackColor       =   -2147483643
                  End
                  Begin ComctlLib.ImageCombo ImgCombo 
                     DataField       =   "CsADCBtnDisaClr"
                     Height          =   330
                     Index           =   19
                     Left            =   2040
                     TabIndex        =   11
                     Top             =   780
                     Width           =   1755
                     _ExtentX        =   3096
                     _ExtentY        =   582
                     _Version        =   393216
                     ForeColor       =   -2147483640
                     BackColor       =   -2147483643
                  End
                  Begin ComctlLib.ImageCombo ImgCombo 
                     DataField       =   "CsADCBtnMaskClr"
                     Height          =   330
                     Index           =   20
                     Left            =   2040
                     TabIndex        =   12
                     Top             =   1110
                     Width           =   1755
                     _ExtentX        =   3096
                     _ExtentY        =   582
                     _Version        =   393216
                     ForeColor       =   -2147483640
                     BackColor       =   -2147483643
                  End
                  Begin MwfCtl.MWCTL_BTN1 CmdAdcBtnPic 
                     Height          =   435
                     Left            =   60
                     TabIndex        =   16
                     ToolTipText     =   "Get ADC Button Picture"
                     Top             =   2430
                     Width           =   1185
                     _ExtentX        =   2090
                     _ExtentY        =   767
                     BackColor       =   -2147483633
                     Caption         =   "&Btn Pic"
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
                     Index           =   24
                     Left            =   2040
                     TabIndex        =   13
                     ToolTipText     =   "Enter ADC Button Height"
                     Top             =   1440
                     Width           =   735
                     _ExtentX        =   1296
                     _ExtentY        =   503
                     Alignment       =   2
                     Mask            =   "####0"
                     MaxLength       =   5
                     DataType        =   1
                     DataField       =   "CsADCBtnHeight"
                     ReCalcParent    =   "WCSADCBTNPIC"
                     IdName          =   "CSADCBTNHEIGHT"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   0
                     Left            =   2040
                     TabIndex        =   14
                     ToolTipText     =   "Enter ADC Button Width"
                     Top             =   1725
                     Width           =   735
                     _ExtentX        =   1296
                     _ExtentY        =   503
                     Alignment       =   2
                     Mask            =   "####0"
                     MaxLength       =   5
                     DataType        =   1
                     DataField       =   "CsADCBtnWidth"
                     ReCalcParent    =   "WCSADCBTNPIC"
                     IdName          =   "CSADCBTNWIDTH"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   12
                     Left            =   2040
                     TabIndex        =   15
                     ToolTipText     =   "Enter ADC Button Picture Path (Picture Set Folder Has To Be Present In \\AppServer\D\MwApp\MwClrSet)"
                     Top             =   2010
                     Width           =   1755
                     _ExtentX        =   3096
                     _ExtentY        =   503
                     MaxLength       =   500
                     DataField       =   "CsADCBtnPic"
                     IdName          =   "CSADCBTNPIC"
                  End
                  Begin MwfCtl.MWCTL_BTN1 CmdAdcBtnPicSample 
                     Height          =   465
                     Left            =   2040
                     TabIndex        =   208
                     ToolTipText     =   "Get ADC Button Picture"
                     Top             =   2430
                     Width           =   1215
                     _ExtentX        =   2143
                     _ExtentY        =   820
                     BackColor       =   -2147483633
                     Caption         =   "Caption"
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
                     Height          =   300
                     Index           =   20
                     Left            =   3450
                     TabIndex        =   209
                     Top             =   1470
                     Visible         =   0   'False
                     Width           =   345
                     _ExtentX        =   609
                     _ExtentY        =   529
                     Alignment       =   2
                     Enabled         =   0   'False
                     MaxLength       =   0
                     ReCalcOn        =   "CSADCBTNHEIGHT,CSADCBTNWIDTH"
                     IdName          =   "WCSADCBTNPIC"
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Button Fore Color"
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
                     Height          =   330
                     Index           =   24
                     Left            =   60
                     TabIndex        =   216
                     Top             =   450
                     Width           =   1995
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Button Disable Color"
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
                     Height          =   330
                     Index           =   23
                     Left            =   60
                     TabIndex        =   215
                     Top             =   780
                     Width           =   1995
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Button Mask Color"
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
                     Height          =   330
                     Index           =   22
                     Left            =   60
                     TabIndex        =   214
                     Top             =   1110
                     Width           =   1995
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Button Height"
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
                     Left            =   60
                     TabIndex        =   213
                     Top             =   1440
                     Width           =   1995
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Button Width"
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
                     Left            =   60
                     TabIndex        =   212
                     Top             =   1725
                     Width           =   1995
                  End
                  Begin VB.Label Label3 
                     BackStyle       =   0  'Transparent
                     BorderStyle     =   1  'Fixed Single
                     Caption         =   "  Color Setting For ADC Control"
                     Height          =   255
                     Left            =   -30
                     TabIndex        =   211
                     Top             =   90
                     Width           =   4005
                  End
                  Begin VB.Label ALBL 
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
                     Index           =   69
                     Left            =   60
                     TabIndex        =   210
                     Top             =   2010
                     Width           =   1995
                  End
               End
               Begin VB.Frame Picture7 
                  Height          =   2475
                  Left            =   -30
                  TabIndex        =   200
                  Top             =   -120
                  Width           =   3975
                  Begin ComctlLib.ImageCombo ImgCombo 
                     DataField       =   "CsMedBckClr"
                     Height          =   330
                     Index           =   0
                     Left            =   2040
                     TabIndex        =   5
                     Top             =   450
                     Width           =   1755
                     _ExtentX        =   3096
                     _ExtentY        =   582
                     _Version        =   393216
                     ForeColor       =   -2147483640
                     BackColor       =   -2147483643
                  End
                  Begin ComctlLib.ImageCombo ImgCombo 
                     DataField       =   "CsMedForClr"
                     Height          =   330
                     Index           =   1
                     Left            =   2040
                     TabIndex        =   6
                     Top             =   780
                     Width           =   1755
                     _ExtentX        =   3096
                     _ExtentY        =   582
                     _Version        =   393216
                     ForeColor       =   -2147483640
                     BackColor       =   -2147483643
                  End
                  Begin ComctlLib.ImageCombo ImgCombo 
                     DataField       =   "CsMedBckClrDisa"
                     Height          =   330
                     Index           =   2
                     Left            =   2040
                     TabIndex        =   7
                     Top             =   1110
                     Width           =   1755
                     _ExtentX        =   3096
                     _ExtentY        =   582
                     _Version        =   393216
                     ForeColor       =   -2147483640
                     BackColor       =   -2147483643
                  End
                  Begin ComctlLib.ImageCombo ImgCombo 
                     DataField       =   "CsSysGrayText"
                     Height          =   330
                     Index           =   3
                     Left            =   2040
                     TabIndex        =   8
                     Top             =   1440
                     Width           =   1755
                     _ExtentX        =   3096
                     _ExtentY        =   582
                     _Version        =   393216
                     ForeColor       =   -2147483640
                     BackColor       =   -2147483643
                  End
                  Begin ComctlLib.ImageCombo ImgCombo 
                     DataField       =   "CsMedBckClrSel"
                     Height          =   330
                     Index           =   4
                     Left            =   2040
                     TabIndex        =   9
                     Top             =   1770
                     Width           =   1755
                     _ExtentX        =   3096
                     _ExtentY        =   582
                     _Version        =   393216
                     ForeColor       =   -2147483640
                     BackColor       =   -2147483643
                  End
                  Begin VB.Label Label2 
                     BackStyle       =   0  'Transparent
                     BorderStyle     =   1  'Fixed Single
                     Caption         =   "  Color Setting For MED Control"
                     Height          =   255
                     Left            =   -30
                     TabIndex        =   206
                     Top             =   90
                     Width           =   3975
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Sel Field Back Color"
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
                     Height          =   330
                     Index           =   6
                     Left            =   60
                     TabIndex        =   205
                     Top             =   1770
                     Width           =   1995
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Fore Color Disabled"
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
                     Height          =   330
                     Index           =   5
                     Left            =   60
                     TabIndex        =   204
                     Top             =   1440
                     Width           =   1995
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Back Color Disabled"
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
                     Height          =   330
                     Index           =   4
                     Left            =   60
                     TabIndex        =   203
                     Top             =   1110
                     Width           =   1995
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Fore Color"
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
                     Height          =   330
                     Index           =   2
                     Left            =   60
                     TabIndex        =   202
                     Top             =   780
                     Width           =   1995
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Back Color"
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
                     Height          =   330
                     Index           =   1
                     Left            =   60
                     TabIndex        =   201
                     Top             =   450
                     Width           =   1995
                  End
               End
            End
            Begin VB.Frame PicBoxTabDet 
               BorderStyle     =   0  'None
               Height          =   6210
               Index           =   1
               Left            =   -74940
               TabIndex        =   46
               Top             =   360
               Width           =   11825
               Begin VB.Frame Picture4 
                  BorderStyle     =   0  'None
                  Height          =   6165
                  Left            =   7830
                  TabIndex        =   150
                  Top             =   -120
                  Width           =   3975
                  Begin VB.PictureBox PicBoxCsMDIPic 
                     Height          =   1215
                     Left            =   -30
                     ScaleHeight     =   1155
                     ScaleWidth      =   3915
                     TabIndex        =   151
                     Top             =   4380
                     Width           =   3975
                  End
                  Begin ComctlLib.ImageCombo ImgCombo 
                     DataField       =   "CsSysWindowFrame"
                     Height          =   330
                     Index           =   50
                     Left            =   2070
                     TabIndex        =   73
                     Top             =   450
                     Width           =   1755
                     _ExtentX        =   3096
                     _ExtentY        =   582
                     _Version        =   393216
                     ForeColor       =   -2147483640
                     BackColor       =   -2147483643
                  End
                  Begin ComctlLib.ImageCombo ImgCombo 
                     DataField       =   "CsSysMenu"
                     Height          =   330
                     Index           =   51
                     Left            =   2070
                     TabIndex        =   74
                     Top             =   780
                     Width           =   1755
                     _ExtentX        =   3096
                     _ExtentY        =   582
                     _Version        =   393216
                     ForeColor       =   -2147483640
                     BackColor       =   -2147483643
                  End
                  Begin ComctlLib.ImageCombo ImgCombo 
                     DataField       =   "CsSysMenuText"
                     Height          =   330
                     Index           =   52
                     Left            =   2070
                     TabIndex        =   75
                     Top             =   1110
                     Width           =   1755
                     _ExtentX        =   3096
                     _ExtentY        =   582
                     _Version        =   393216
                     ForeColor       =   -2147483640
                     BackColor       =   -2147483643
                  End
                  Begin ComctlLib.ImageCombo ImgCombo 
                     DataField       =   "CsSysActiveCaption"
                     Height          =   330
                     Index           =   53
                     Left            =   2070
                     TabIndex        =   76
                     Top             =   1440
                     Width           =   1755
                     _ExtentX        =   3096
                     _ExtentY        =   582
                     _Version        =   393216
                     ForeColor       =   -2147483640
                     BackColor       =   -2147483643
                  End
                  Begin ComctlLib.ImageCombo ImgCombo 
                     DataField       =   "CsSysCaptionText"
                     Height          =   330
                     Index           =   54
                     Left            =   2070
                     TabIndex        =   77
                     Top             =   1770
                     Width           =   1755
                     _ExtentX        =   3096
                     _ExtentY        =   582
                     _Version        =   393216
                     ForeColor       =   -2147483640
                     BackColor       =   -2147483643
                  End
                  Begin ComctlLib.ImageCombo ImgCombo 
                     DataField       =   "CsSysHighLight"
                     Height          =   330
                     Index           =   46
                     Left            =   2070
                     TabIndex        =   78
                     Top             =   2100
                     Width           =   1755
                     _ExtentX        =   3096
                     _ExtentY        =   582
                     _Version        =   393216
                     ForeColor       =   -2147483640
                     BackColor       =   -2147483643
                  End
                  Begin ComctlLib.ImageCombo ImgCombo 
                     DataField       =   "CsSysHighLightText"
                     Height          =   330
                     Index           =   47
                     Left            =   2070
                     TabIndex        =   79
                     Top             =   2430
                     Width           =   1755
                     _ExtentX        =   3096
                     _ExtentY        =   582
                     _Version        =   393216
                     ForeColor       =   -2147483640
                     BackColor       =   -2147483643
                  End
                  Begin ComctlLib.ImageCombo ImgCombo 
                     DataField       =   "CsSysBtnFace"
                     Height          =   330
                     Index           =   41
                     Left            =   2070
                     TabIndex        =   80
                     Top             =   2760
                     Width           =   1755
                     _ExtentX        =   3096
                     _ExtentY        =   582
                     _Version        =   393216
                     ForeColor       =   -2147483640
                     BackColor       =   -2147483643
                  End
                  Begin ComctlLib.ImageCombo ImgCombo 
                     DataField       =   "CsSysBtnShadow"
                     Height          =   330
                     Index           =   42
                     Left            =   2070
                     TabIndex        =   81
                     Top             =   3090
                     Width           =   1755
                     _ExtentX        =   3096
                     _ExtentY        =   582
                     _Version        =   393216
                     ForeColor       =   -2147483640
                     BackColor       =   -2147483643
                  End
                  Begin ComctlLib.ImageCombo ImgCombo 
                     DataField       =   "CsSysBtnText"
                     Height          =   330
                     Index           =   34
                     Left            =   2070
                     TabIndex        =   82
                     Top             =   3420
                     Width           =   1755
                     _ExtentX        =   3096
                     _ExtentY        =   582
                     _Version        =   393216
                     ForeColor       =   -2147483640
                     BackColor       =   -2147483643
                  End
                  Begin ComctlLib.ImageCombo ImgCombo 
                     DataField       =   "CsSysBtnHighLight"
                     Height          =   330
                     Index           =   35
                     Left            =   2070
                     TabIndex        =   83
                     Top             =   3750
                     Width           =   1755
                     _ExtentX        =   3096
                     _ExtentY        =   582
                     _Version        =   393216
                     ForeColor       =   -2147483640
                     BackColor       =   -2147483643
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   14
                     Left            =   2070
                     TabIndex        =   84
                     ToolTipText     =   "Specify Whether Selected Objects Should Appear In 3D Or Flat ('Y' for 3D)"
                     Top             =   4080
                     Width           =   375
                     _ExtentX        =   661
                     _ExtentY        =   503
                     DataType        =   4
                     DataField       =   "Cs3D"
                     IdName          =   "CS3D"
                  End
                  Begin MwfCtl.MWCTL_BTN1 CmdMDIPic 
                     Height          =   435
                     Left            =   60
                     TabIndex        =   85
                     Top             =   5700
                     Width           =   1185
                     _ExtentX        =   2090
                     _ExtentY        =   767
                     BackColor       =   -2147483633
                     Caption         =   "&MDI Pic"
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
                     Index           =   29
                     Left            =   2070
                     TabIndex        =   86
                     ToolTipText     =   "Enter MDI Picture Path"
                     Top             =   5700
                     Width           =   1755
                     _ExtentX        =   3096
                     _ExtentY        =   503
                     MaxLength       =   500
                     DataField       =   "CsMDIPic"
                     IdName          =   "CSMDIPIC"
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Control Highlight"
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
                     Height          =   330
                     Index           =   62
                     Left            =   60
                     TabIndex        =   165
                     Top             =   3750
                     Width           =   2085
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Control Text"
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
                     Height          =   330
                     Index           =   61
                     Left            =   60
                     TabIndex        =   164
                     Top             =   3420
                     Width           =   2085
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Control Shadow"
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
                     Height          =   330
                     Index           =   60
                     Left            =   60
                     TabIndex        =   163
                     Top             =   3090
                     Width           =   2085
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Control Face"
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
                     Height          =   330
                     Index           =   59
                     Left            =   60
                     TabIndex        =   162
                     Top             =   2760
                     Width           =   2085
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Menu HighLight Text"
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
                     Height          =   330
                     Index           =   58
                     Left            =   60
                     TabIndex        =   161
                     Top             =   2430
                     Width           =   2085
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Window Frame"
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
                     Height          =   330
                     Index           =   57
                     Left            =   60
                     TabIndex        =   160
                     Top             =   450
                     Width           =   2085
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Menu Back"
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
                     Height          =   330
                     Index           =   56
                     Left            =   60
                     TabIndex        =   159
                     Top             =   780
                     Width           =   2085
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Menu Text"
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
                     Height          =   330
                     Index           =   55
                     Left            =   60
                     TabIndex        =   158
                     Top             =   1110
                     Width           =   2085
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Form Caption Back"
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
                     Height          =   330
                     Index           =   27
                     Left            =   60
                     TabIndex        =   157
                     Top             =   1440
                     Width           =   2085
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Form Caption Text"
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
                     Height          =   330
                     Index           =   26
                     Left            =   60
                     TabIndex        =   156
                     Top             =   1770
                     Width           =   2085
                  End
                  Begin VB.Label Label9 
                     BackStyle       =   0  'Transparent
                     BorderStyle     =   1  'Fixed Single
                     Caption         =   "  System Color Settings"
                     Height          =   255
                     Left            =   -30
                     TabIndex        =   155
                     Top             =   90
                     Width           =   4065
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Menu HighLight"
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
                     Height          =   330
                     Index           =   25
                     Left            =   60
                     TabIndex        =   154
                     Top             =   2100
                     Width           =   2085
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "3D (Y/ N)"
                     BeginProperty Font 
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
                     Left            =   60
                     TabIndex        =   153
                     Top             =   4080
                     Width           =   1605
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Path"
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
                     Index           =   84
                     Left            =   1530
                     TabIndex        =   152
                     Top             =   5700
                     Width           =   495
                  End
               End
               Begin VB.Frame Picture14 
                  Height          =   3915
                  Left            =   3930
                  TabIndex        =   138
                  Top             =   2340
                  Width           =   3915
                  Begin ComctlLib.ImageCombo ImgCombo 
                     DataField       =   "CsBtnForClrSet2"
                     Height          =   330
                     Index           =   43
                     Left            =   2040
                     TabIndex        =   64
                     Top             =   450
                     Width           =   1755
                     _ExtentX        =   3096
                     _ExtentY        =   582
                     _Version        =   393216
                     ForeColor       =   -2147483640
                     BackColor       =   -2147483643
                  End
                  Begin ComctlLib.ImageCombo ImgCombo 
                     DataField       =   "CsBtnDisaClrSet2"
                     Height          =   330
                     Index           =   44
                     Left            =   2040
                     TabIndex        =   65
                     Top             =   780
                     Width           =   1755
                     _ExtentX        =   3096
                     _ExtentY        =   582
                     _Version        =   393216
                     ForeColor       =   -2147483640
                     BackColor       =   -2147483643
                  End
                  Begin ComctlLib.ImageCombo ImgCombo 
                     DataField       =   "CsBtnMaskClrSet2"
                     Height          =   330
                     Index           =   45
                     Left            =   2040
                     TabIndex        =   66
                     Top             =   1110
                     Width           =   1755
                     _ExtentX        =   3096
                     _ExtentY        =   582
                     _Version        =   393216
                     ForeColor       =   -2147483640
                     BackColor       =   -2147483643
                  End
                  Begin MwfCtl.MWCTL_BTN1 CmdBtnPicSet2 
                     Height          =   435
                     Left            =   60
                     TabIndex        =   72
                     Top             =   2970
                     Width           =   1185
                     _ExtentX        =   2090
                     _ExtentY        =   767
                     BackColor       =   -2147483633
                     Caption         =   "Btn &Pic"
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
                     Index           =   5
                     Left            =   2040
                     TabIndex        =   67
                     ToolTipText     =   "Enter Button Height (Set 2)"
                     Top             =   1440
                     Width           =   735
                     _ExtentX        =   1296
                     _ExtentY        =   503
                     Alignment       =   2
                     Mask            =   "####0"
                     MaxLength       =   5
                     DataType        =   1
                     DataField       =   "CsBtnHeightSet2"
                     ReCalcParent    =   "WCSBTNPICSET2"
                     IdName          =   "CSBTNHEIGHTSET2"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   6
                     Left            =   2040
                     TabIndex        =   68
                     ToolTipText     =   "Enter Button Width (Set 2)"
                     Top             =   1725
                     Width           =   735
                     _ExtentX        =   1296
                     _ExtentY        =   503
                     Alignment       =   2
                     Mask            =   "####0"
                     MaxLength       =   5
                     DataType        =   1
                     DataField       =   "CsBtnWidthSet2"
                     ReCalcParent    =   "WCSBTNPICSET2"
                     IdName          =   "CSBTNWIDTHSET2"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   10
                     Left            =   2040
                     TabIndex        =   69
                     ToolTipText     =   "Enter Button Top Difference (Set 2)"
                     Top             =   2010
                     Width           =   735
                     _ExtentX        =   1296
                     _ExtentY        =   503
                     Alignment       =   2
                     Mask            =   "-####0"
                     MaxLength       =   6
                     DataType        =   1
                     DataField       =   "CsBtnTopDiffSet2"
                     IdName          =   "CSBTNTOPDIFFSET2"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   11
                     Left            =   2040
                     TabIndex        =   70
                     ToolTipText     =   "Enter Button Left Difference (Set 2)"
                     Top             =   2295
                     Width           =   735
                     _ExtentX        =   1296
                     _ExtentY        =   503
                     Alignment       =   2
                     Mask            =   "-####0"
                     MaxLength       =   6
                     DataType        =   1
                     DataField       =   "CsBtnLeftDiffSet2"
                     IdName          =   "CSBTNLEFTDIFFSET2"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   330
                     Index           =   22
                     Left            =   3450
                     TabIndex        =   139
                     Top             =   1890
                     Visible         =   0   'False
                     Width           =   345
                     _ExtentX        =   609
                     _ExtentY        =   582
                     Alignment       =   2
                     Enabled         =   0   'False
                     MaxLength       =   0
                     ReCalcOn        =   "CSBTNHEIGHTSET2,CSBTNWIDTHSET2"
                     IdName          =   "WCSBTNPICSET2"
                  End
                  Begin MwfCtl.MWCTL_BTN1 CmdBtnPicSet2Sample 
                     Height          =   465
                     Left            =   2040
                     TabIndex        =   140
                     ToolTipText     =   "Get ADC Button Picture"
                     Top             =   2970
                     Width           =   1215
                     _ExtentX        =   2143
                     _ExtentY        =   820
                     BackColor       =   -2147483633
                     Caption         =   "Caption"
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
                     Index           =   28
                     Left            =   2040
                     TabIndex        =   71
                     ToolTipText     =   "Enter Set1 Button Picture Path"
                     Top             =   2580
                     Width           =   1755
                     _ExtentX        =   3096
                     _ExtentY        =   503
                     MaxLength       =   500
                     DataField       =   "CsBtnPicSet2"
                     IdName          =   "CSBTNPICSET2"
                  End
                  Begin VB.Label Label8 
                     BackStyle       =   0  'Transparent
                     BorderStyle     =   1  'Fixed Single
                     Caption         =   "  Color Setting For Button Controls (Set 2)"
                     Height          =   255
                     Left            =   -30
                     TabIndex        =   149
                     Top             =   90
                     Width           =   3975
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Button Width"
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
                     Index           =   52
                     Left            =   60
                     TabIndex        =   148
                     Top             =   1725
                     Width           =   1995
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Button Height"
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
                     Index           =   51
                     Left            =   60
                     TabIndex        =   147
                     Top             =   1440
                     Width           =   1995
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Button Mask Color"
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
                     Height          =   330
                     Index           =   50
                     Left            =   60
                     TabIndex        =   146
                     Top             =   1110
                     Width           =   1995
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Button Disable Color"
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
                     Height          =   330
                     Index           =   49
                     Left            =   60
                     TabIndex        =   145
                     Top             =   780
                     Width           =   1995
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Button Fore Color"
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
                     Height          =   330
                     Index           =   48
                     Left            =   60
                     TabIndex        =   144
                     Top             =   450
                     Width           =   1995
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Button Left Diff"
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
                     Index           =   66
                     Left            =   60
                     TabIndex        =   143
                     Top             =   2295
                     Width           =   1995
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Button Top Diff"
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
                     Index           =   67
                     Left            =   60
                     TabIndex        =   142
                     Top             =   2010
                     Width           =   1995
                  End
                  Begin VB.Label ALBL 
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
                     Index           =   64
                     Left            =   60
                     TabIndex        =   141
                     Top             =   2580
                     Width           =   1995
                  End
               End
               Begin VB.Frame Picture3 
                  Height          =   2565
                  Left            =   3930
                  TabIndex        =   130
                  Top             =   -120
                  Width           =   3915
                  Begin ComctlLib.ImageCombo ImgCombo 
                     DataField       =   "CsTabBckClr"
                     Height          =   330
                     Index           =   24
                     Left            =   2040
                     TabIndex        =   58
                     Top             =   450
                     Width           =   1755
                     _ExtentX        =   3096
                     _ExtentY        =   582
                     _Version        =   393216
                     ForeColor       =   -2147483640
                     BackColor       =   -2147483643
                  End
                  Begin ComctlLib.ImageCombo ImgCombo 
                     DataField       =   "CsTabPgClr"
                     Height          =   330
                     Index           =   25
                     Left            =   2040
                     TabIndex        =   59
                     Top             =   780
                     Width           =   1755
                     _ExtentX        =   3096
                     _ExtentY        =   582
                     _Version        =   393216
                     ForeColor       =   -2147483640
                     BackColor       =   -2147483643
                  End
                  Begin ComctlLib.ImageCombo ImgCombo 
                     DataField       =   "CsTabForClr"
                     Height          =   330
                     Index           =   26
                     Left            =   2040
                     TabIndex        =   60
                     Top             =   1110
                     Width           =   1755
                     _ExtentX        =   3096
                     _ExtentY        =   582
                     _Version        =   393216
                     ForeColor       =   -2147483640
                     BackColor       =   -2147483643
                  End
                  Begin ComctlLib.ImageCombo ImgCombo 
                     DataField       =   "CsTabBorderClr"
                     Height          =   330
                     Index           =   36
                     Left            =   2040
                     TabIndex        =   61
                     Top             =   1440
                     Width           =   1755
                     _ExtentX        =   3096
                     _ExtentY        =   582
                     _Version        =   393216
                     ForeColor       =   -2147483640
                     BackColor       =   -2147483643
                  End
                  Begin ComctlLib.ImageCombo ImgCombo 
                     DataField       =   "CsTabShadowClr"
                     Height          =   330
                     Index           =   37
                     Left            =   2040
                     TabIndex        =   62
                     Top             =   1770
                     Width           =   1755
                     _ExtentX        =   3096
                     _ExtentY        =   582
                     _Version        =   393216
                     ForeColor       =   -2147483640
                     BackColor       =   -2147483643
                  End
                  Begin ComctlLib.ImageCombo ImgCombo 
                     DataField       =   "CsTabHighLightClr"
                     Height          =   330
                     Index           =   48
                     Left            =   2040
                     TabIndex        =   63
                     Top             =   2100
                     Width           =   1755
                     _ExtentX        =   3096
                     _ExtentY        =   582
                     _Version        =   393216
                     ForeColor       =   -2147483640
                     BackColor       =   -2147483643
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "HighLight Color"
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
                     Height          =   330
                     Index           =   53
                     Left            =   60
                     TabIndex        =   137
                     Top             =   2100
                     Width           =   1995
                  End
                  Begin VB.Label Label5 
                     BackStyle       =   0  'Transparent
                     BorderStyle     =   1  'Fixed Single
                     Caption         =   "  Color Setting For Tab Page"
                     Height          =   255
                     Left            =   -30
                     TabIndex        =   136
                     Top             =   90
                     Width           =   3975
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Shadow Color"
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
                     Height          =   330
                     Index           =   43
                     Left            =   60
                     TabIndex        =   135
                     Top             =   1770
                     Width           =   1995
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Border Color"
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
                     Height          =   330
                     Index           =   42
                     Left            =   60
                     TabIndex        =   134
                     Top             =   1440
                     Width           =   1995
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Fore Color"
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
                     Height          =   330
                     Index           =   41
                     Left            =   60
                     TabIndex        =   133
                     Top             =   1110
                     Width           =   1995
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Tab Color"
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
                     Height          =   330
                     Index           =   40
                     Left            =   60
                     TabIndex        =   132
                     Top             =   780
                     Width           =   1995
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Back Color"
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
                     Height          =   330
                     Index           =   39
                     Left            =   60
                     TabIndex        =   131
                     Top             =   450
                     Width           =   1995
                  End
               End
               Begin VB.Frame Picture12 
                  Height          =   3915
                  Left            =   -30
                  TabIndex        =   170
                  Top             =   2340
                  Width           =   4005
                  Begin ComctlLib.ImageCombo ImgCombo 
                     DataField       =   "CsBtnForClrSet1"
                     Height          =   330
                     Index           =   38
                     Left            =   2040
                     TabIndex        =   49
                     Top             =   450
                     Width           =   1755
                     _ExtentX        =   3096
                     _ExtentY        =   582
                     _Version        =   393216
                     ForeColor       =   -2147483640
                     BackColor       =   -2147483643
                  End
                  Begin ComctlLib.ImageCombo ImgCombo 
                     DataField       =   "CsBtnDisaClrSet1"
                     Height          =   330
                     Index           =   39
                     Left            =   2040
                     TabIndex        =   50
                     Top             =   780
                     Width           =   1755
                     _ExtentX        =   3096
                     _ExtentY        =   582
                     _Version        =   393216
                     ForeColor       =   -2147483640
                     BackColor       =   -2147483643
                  End
                  Begin ComctlLib.ImageCombo ImgCombo 
                     DataField       =   "CsBtnMaskClrSet1"
                     Height          =   330
                     Index           =   40
                     Left            =   2040
                     TabIndex        =   51
                     Top             =   1110
                     Width           =   1755
                     _ExtentX        =   3096
                     _ExtentY        =   582
                     _Version        =   393216
                     ForeColor       =   -2147483640
                     BackColor       =   -2147483643
                  End
                  Begin MwfCtl.MWCTL_BTN1 CmdBtnPicSet1 
                     Height          =   435
                     Left            =   60
                     TabIndex        =   57
                     Top             =   2970
                     Width           =   1185
                     _ExtentX        =   2090
                     _ExtentY        =   767
                     BackColor       =   -2147483633
                     Caption         =   "&Btn Pic"
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
                     Index           =   2
                     Left            =   2040
                     TabIndex        =   52
                     ToolTipText     =   "Enter Button Height (Set 1)"
                     Top             =   1440
                     Width           =   735
                     _ExtentX        =   1296
                     _ExtentY        =   503
                     Alignment       =   2
                     Mask            =   "####0"
                     MaxLength       =   5
                     DataType        =   1
                     DataField       =   "CsBtnHeightSet1"
                     ReCalcParent    =   "WCSBTNPICSET1"
                     IdName          =   "CSBTNHEIGHTSET1"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   4
                     Left            =   2040
                     TabIndex        =   53
                     ToolTipText     =   "Enter Button Width (Set 1)"
                     Top             =   1725
                     Width           =   735
                     _ExtentX        =   1296
                     _ExtentY        =   503
                     Alignment       =   2
                     Mask            =   "####0"
                     MaxLength       =   5
                     DataType        =   1
                     DataField       =   "CsBtnWidthSet1"
                     ReCalcParent    =   "WCSBTNPICSET1"
                     IdName          =   "CSBTNWIDTHSET1"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   8
                     Left            =   2040
                     TabIndex        =   54
                     ToolTipText     =   "Enter Button Top Difference (Set 1)"
                     Top             =   2010
                     Width           =   735
                     _ExtentX        =   1296
                     _ExtentY        =   503
                     Alignment       =   2
                     Mask            =   "-####0"
                     MaxLength       =   6
                     DataType        =   1
                     DataField       =   "CsBtnTopDiffSet1"
                     IdName          =   "CSBTNTOPDIFFSET1"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   9
                     Left            =   2040
                     TabIndex        =   55
                     ToolTipText     =   "Enter Button Left Difference (Set 1)"
                     Top             =   2295
                     Width           =   735
                     _ExtentX        =   1296
                     _ExtentY        =   503
                     Alignment       =   2
                     Mask            =   "-####0"
                     MaxLength       =   6
                     DataType        =   1
                     DataField       =   "CsBtnLeftDiffSet1"
                     IdName          =   "CSBTNLEFTDIFFSET1"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   330
                     Index           =   21
                     Left            =   3420
                     TabIndex        =   171
                     Top             =   1890
                     Visible         =   0   'False
                     Width           =   345
                     _ExtentX        =   609
                     _ExtentY        =   582
                     Alignment       =   2
                     Enabled         =   0   'False
                     MaxLength       =   0
                     ReCalcOn        =   "CSBTNHEIGHTSET1,CSBTNWIDTHSET1"
                     IdName          =   "WCSBTNPICSET1"
                  End
                  Begin MwfCtl.MWCTL_BTN1 CmdBtnPicSet1Sample 
                     Height          =   465
                     Left            =   2040
                     TabIndex        =   172
                     ToolTipText     =   "Get ADC Button Picture"
                     Top             =   2970
                     Width           =   1215
                     _ExtentX        =   2143
                     _ExtentY        =   820
                     BackColor       =   -2147483633
                     Caption         =   "Caption"
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
                     Index           =   27
                     Left            =   2040
                     TabIndex        =   56
                     ToolTipText     =   "Enter Set1 Button Picture Path"
                     Top             =   2580
                     Width           =   1755
                     _ExtentX        =   3096
                     _ExtentY        =   503
                     MaxLength       =   500
                     DataField       =   "CsBtnPicSet1"
                     IdName          =   "CSBTNPICSET1"
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Button Fore Color"
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
                     Height          =   330
                     Index           =   47
                     Left            =   60
                     TabIndex        =   181
                     Top             =   450
                     Width           =   1995
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Button Disable Color"
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
                     Height          =   330
                     Index           =   46
                     Left            =   60
                     TabIndex        =   180
                     Top             =   780
                     Width           =   1995
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Button Mask Color"
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
                     Height          =   330
                     Index           =   45
                     Left            =   60
                     TabIndex        =   179
                     Top             =   1110
                     Width           =   1995
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Button Height"
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
                     Index           =   44
                     Left            =   60
                     TabIndex        =   178
                     Top             =   1440
                     Width           =   1995
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Button Width"
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
                     Left            =   60
                     TabIndex        =   177
                     Top             =   1725
                     Width           =   1995
                  End
                  Begin VB.Label Label7 
                     BackStyle       =   0  'Transparent
                     BorderStyle     =   1  'Fixed Single
                     Caption         =   "  Color Setting For Button Controls (Set 1)"
                     Height          =   255
                     Left            =   -30
                     TabIndex        =   176
                     Top             =   90
                     Width           =   4005
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Button Top Diff"
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
                     Index           =   65
                     Left            =   60
                     TabIndex        =   175
                     Top             =   2010
                     Width           =   1995
                  End
                  Begin VB.Label ALBL 
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
                     Index           =   82
                     Left            =   60
                     TabIndex        =   174
                     Top             =   2580
                     Width           =   1995
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Button Left Diff"
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
                     Index           =   83
                     Left            =   60
                     TabIndex        =   173
                     Top             =   2295
                     Width           =   1995
                  End
               End
               Begin VB.Frame Picture5 
                  Height          =   2505
                  Left            =   -30
                  TabIndex        =   166
                  Top             =   -120
                  Width           =   4005
                  Begin VB.PictureBox PicBoxCsFormPic 
                     Height          =   1575
                     Left            =   -30
                     ScaleHeight     =   1515
                     ScaleWidth      =   3945
                     TabIndex        =   167
                     Top             =   360
                     Width           =   4005
                  End
                  Begin MwfCtl.MWCTL_BTN1 CmdFrmPic 
                     Height          =   435
                     Left            =   30
                     TabIndex        =   47
                     Top             =   1950
                     Width           =   1185
                     _ExtentX        =   2090
                     _ExtentY        =   767
                     BackColor       =   -2147483633
                     Caption         =   "&Frm Pic"
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
                     Index           =   26
                     Left            =   2070
                     TabIndex        =   48
                     ToolTipText     =   "Enter Form Picture Path"
                     Top             =   1950
                     Width           =   1755
                     _ExtentX        =   3096
                     _ExtentY        =   503
                     MaxLength       =   500
                     DataField       =   "CsFormPic"
                     IdName          =   "CSFORMPIC"
                  End
                  Begin VB.Label Label6 
                     BackStyle       =   0  'Transparent
                     BorderStyle     =   1  'Fixed Single
                     Caption         =   "  Form Settings"
                     Height          =   255
                     Left            =   -30
                     TabIndex        =   169
                     Top             =   90
                     Width           =   4005
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Path"
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
                     Index           =   81
                     Left            =   1560
                     TabIndex        =   168
                     Top             =   1950
                     Width           =   495
                  End
               End
            End
            Begin VB.Frame PicBoxTabDet 
               BorderStyle     =   0  'None
               Height          =   6210
               Index           =   2
               Left            =   -74940
               TabIndex        =   87
               Top             =   360
               Width           =   11825
               Begin VB.Frame Picture2 
                  Height          =   5055
                  Left            =   -30
                  TabIndex        =   115
                  Top             =   1200
                  Width           =   3945
                  Begin ComctlLib.ImageCombo ImgCombo 
                     DataField       =   "CsTlBarBtnForClr"
                     Height          =   330
                     Index           =   55
                     Left            =   2040
                     TabIndex        =   90
                     Top             =   450
                     Width           =   1755
                     _ExtentX        =   3096
                     _ExtentY        =   582
                     _Version        =   393216
                     ForeColor       =   -2147483640
                     BackColor       =   -2147483643
                  End
                  Begin ComctlLib.ImageCombo ImgCombo 
                     DataField       =   "CsTlBarBtnDisaClr"
                     Height          =   330
                     Index           =   56
                     Left            =   2040
                     TabIndex        =   91
                     Top             =   780
                     Width           =   1755
                     _ExtentX        =   3096
                     _ExtentY        =   582
                     _Version        =   393216
                     ForeColor       =   -2147483640
                     BackColor       =   -2147483643
                  End
                  Begin ComctlLib.ImageCombo ImgCombo 
                     DataField       =   "CsTlBarBtnMaskClr"
                     Height          =   330
                     Index           =   57
                     Left            =   2040
                     TabIndex        =   92
                     Top             =   1110
                     Width           =   1755
                     _ExtentX        =   3096
                     _ExtentY        =   582
                     _Version        =   393216
                     ForeColor       =   -2147483640
                     BackColor       =   -2147483643
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   16
                     Left            =   2040
                     TabIndex        =   93
                     ToolTipText     =   "Enter Tool Bar Button Height "
                     Top             =   1440
                     Width           =   735
                     _ExtentX        =   1296
                     _ExtentY        =   503
                     Alignment       =   2
                     Mask            =   "####0"
                     MaxLength       =   5
                     DataType        =   1
                     DataField       =   "CsTlBarBtnHeight"
                     ReCalcParent    =   "WCSTLBARPIC"
                     IdName          =   "CSTLBARBTNHEIGHT"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   17
                     Left            =   2040
                     TabIndex        =   94
                     ToolTipText     =   "Enter Tool Bar Button Width"
                     Top             =   1725
                     Width           =   735
                     _ExtentX        =   1296
                     _ExtentY        =   503
                     Alignment       =   2
                     Mask            =   "####0"
                     MaxLength       =   5
                     DataType        =   1
                     DataField       =   "CsTlBarBtnWidth"
                     ReCalcParent    =   "WCSTLBARPIC"
                     IdName          =   "CSTLBARBTNWIDTH"
                  End
                  Begin MwfCtl.MWCTL_MED ATXT 
                     Height          =   285
                     Index           =   18
                     Left            =   2040
                     TabIndex        =   95
                     ToolTipText     =   "Enter Gap Between Tool Bar Buttons"
                     Top             =   2010
                     Width           =   735
                     _ExtentX        =   1296
                     _ExtentY        =   503
                     Alignment       =   2
                     Mask            =   "-####0"
                     MaxLength       =   6
                     DataType        =   1
                     DataField       =   "CsTlBarBtnGap"
                     IdName          =   "CSTLBARBTNGAP"
                  End
                  Begin MwfCtl.MWCTL_BTN1 CmdTlBarBtnPic 
                     Height          =   435
                     Left            =   60
                     TabIndex        =   97
                     ToolTipText     =   "Get Tool Bar Button Picture"
                     Top             =   2730
                     Width           =   1185
                     _ExtentX        =   2090
                     _ExtentY        =   767
                     BackColor       =   -2147483633
                     Caption         =   "T&ool Pic"
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
                     Height          =   330
                     Index           =   23
                     Left            =   3420
                     TabIndex        =   116
                     Top             =   1560
                     Visible         =   0   'False
                     Width           =   345
                     _ExtentX        =   609
                     _ExtentY        =   582
                     Alignment       =   2
                     Enabled         =   0   'False
                     MaxLength       =   0
                     ReCalcOn        =   "CSTLBARBTNHEIGHT,CSTLBARBTNWIDTH"
                     IdName          =   "WCSTLBARPIC"
                  End
                  Begin MwfCtl.MWCTL_BTN1 CmdTlBarBtnPicSample 
                     Height          =   465
                     Left            =   2040
                     TabIndex        =   117
                     ToolTipText     =   "Get ADC Button Picture"
                     Top             =   2730
                     Width           =   1215
                     _ExtentX        =   2143
                     _ExtentY        =   820
                     BackColor       =   -2147483633
                     Caption         =   "Caption"
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
                     Index           =   30
                     Left            =   2040
                     TabIndex        =   96
                     ToolTipText     =   "Enter Tool Bar Button Picture Path"
                     Top             =   2295
                     Width           =   1755
                     _ExtentX        =   3096
                     _ExtentY        =   503
                     MaxLength       =   500
                     DataField       =   "CsTlBarPic"
                     IdName          =   "CSTLBARPIC"
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Gap Between Btns"
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
                     Index           =   72
                     Left            =   60
                     TabIndex        =   125
                     Top             =   2010
                     Width           =   1995
                  End
                  Begin VB.Label Label12 
                     BackStyle       =   0  'Transparent
                     BorderStyle     =   1  'Fixed Single
                     Caption         =   "  Color Setting For Tool Bar Buttons"
                     Height          =   255
                     Left            =   -30
                     TabIndex        =   124
                     Top             =   90
                     Width           =   3945
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Button Width"
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
                     Index           =   74
                     Left            =   60
                     TabIndex        =   123
                     Top             =   1725
                     Width           =   1995
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Button Height"
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
                     Index           =   75
                     Left            =   60
                     TabIndex        =   122
                     Top             =   1440
                     Width           =   1995
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Button Mask Color"
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
                     Height          =   330
                     Index           =   76
                     Left            =   60
                     TabIndex        =   121
                     Top             =   1110
                     Width           =   1995
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Button Disable Color"
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
                     Height          =   330
                     Index           =   77
                     Left            =   60
                     TabIndex        =   120
                     Top             =   780
                     Width           =   1995
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Button Fore Color"
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
                     Height          =   330
                     Index           =   78
                     Left            =   60
                     TabIndex        =   119
                     Top             =   450
                     Width           =   1995
                  End
                  Begin VB.Label ALBL 
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
                     Index           =   85
                     Left            =   60
                     TabIndex        =   118
                     Top             =   2295
                     Width           =   1995
                  End
               End
               Begin VB.Frame Picture1 
                  BorderStyle     =   0  'None
                  Height          =   6135
                  Left            =   3870
                  TabIndex        =   114
                  Top             =   -120
                  Width           =   7785
                  Begin VB.Frame Frame1 
                     Height          =   1365
                     Left            =   0
                     TabIndex        =   240
                     Top             =   2640
                     Width           =   4035
                     Begin ComctlLib.ImageCombo ImgCombo 
                        DataField       =   "CsFraBckClr"
                        Height          =   330
                        Index           =   62
                        Left            =   2040
                        TabIndex        =   101
                        Top             =   540
                        Width           =   1755
                        _ExtentX        =   3096
                        _ExtentY        =   582
                        _Version        =   393216
                        ForeColor       =   0
                        BackColor       =   -2147483643
                     End
                     Begin VB.Label Label11 
                        BackStyle       =   0  'Transparent
                        BorderStyle     =   1  'Fixed Single
                        Caption         =   "  Color Setting For Frame Control"
                        Height          =   255
                        Left            =   -30
                        TabIndex        =   242
                        Top             =   90
                        Width           =   4065
                     End
                     Begin VB.Label ALBL 
                        BackStyle       =   0  'Transparent
                        Caption         =   "Frame Back Color"
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
                        Height          =   330
                        Index           =   89
                        Left            =   60
                        TabIndex        =   241
                        Top             =   540
                        Width           =   1995
                     End
                  End
                  Begin VB.Frame Picture11 
                     Height          =   1365
                     Left            =   0
                     TabIndex        =   237
                     Top             =   1320
                     Width           =   4035
                     Begin ComctlLib.ImageCombo ImgCombo 
                        DataField       =   "CsLineBdrClr"
                        Height          =   330
                        Index           =   58
                        Left            =   2040
                        TabIndex        =   100
                        Top             =   540
                        Width           =   1755
                        _ExtentX        =   3096
                        _ExtentY        =   582
                        _Version        =   393216
                        ForeColor       =   0
                        BackColor       =   -2147483643
                     End
                     Begin VB.Label ALBL 
                        BackStyle       =   0  'Transparent
                        Caption         =   "Line Border Color"
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
                        Height          =   330
                        Index           =   73
                        Left            =   60
                        TabIndex        =   239
                        Top             =   540
                        Width           =   1995
                     End
                     Begin VB.Label Label15 
                        BackStyle       =   0  'Transparent
                        BorderStyle     =   1  'Fixed Single
                        Caption         =   "  Color Setting For Line Control"
                        Height          =   255
                        Left            =   -30
                        TabIndex        =   238
                        Top             =   90
                        Width           =   4065
                     End
                  End
                  Begin VB.Frame Picture6 
                     Height          =   1455
                     Left            =   0
                     TabIndex        =   233
                     Top             =   0
                     Width           =   4035
                     Begin ComctlLib.ImageCombo ImgCombo 
                        DataField       =   "CsDstForClr"
                        Height          =   330
                        Index           =   60
                        Left            =   2040
                        TabIndex        =   99
                        Top             =   780
                        Width           =   1755
                        _ExtentX        =   3096
                        _ExtentY        =   582
                        _Version        =   393216
                        ForeColor       =   -2147483640
                        BackColor       =   -2147483643
                     End
                     Begin ComctlLib.ImageCombo ImgCombo 
                        DataField       =   "CsDstBckClr"
                        Height          =   330
                        Index           =   61
                        Left            =   2040
                        TabIndex        =   98
                        Top             =   450
                        Width           =   1755
                        _ExtentX        =   3096
                        _ExtentY        =   582
                        _Version        =   393216
                        ForeColor       =   0
                        BackColor       =   -2147483643
                     End
                     Begin VB.Label ALBL 
                        BackStyle       =   0  'Transparent
                        Caption         =   "Back Color"
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
                        Height          =   330
                        Index           =   87
                        Left            =   60
                        TabIndex        =   236
                        Top             =   450
                        Width           =   1995
                     End
                     Begin VB.Label Label13 
                        BackStyle       =   0  'Transparent
                        BorderStyle     =   1  'Fixed Single
                        Caption         =   "  Color Setting For Report DST Control"
                        Height          =   255
                        Left            =   -30
                        TabIndex        =   235
                        Top             =   90
                        Width           =   4065
                     End
                     Begin VB.Label ALBL 
                        BackStyle       =   0  'Transparent
                        Caption         =   "Fore Color"
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
                        Height          =   330
                        Index           =   86
                        Left            =   60
                        TabIndex        =   234
                        Top             =   780
                        Width           =   1995
                     End
                  End
               End
               Begin VB.Frame Picture15 
                  Height          =   1485
                  Left            =   -30
                  TabIndex        =   126
                  Top             =   -120
                  Width           =   3945
                  Begin ComctlLib.ImageCombo ImgCombo 
                     DataField       =   "CsLblMainClr"
                     Height          =   330
                     Index           =   59
                     Left            =   2040
                     TabIndex        =   89
                     Top             =   780
                     Width           =   1755
                     _ExtentX        =   3096
                     _ExtentY        =   582
                     _Version        =   393216
                     ForeColor       =   -2147483640
                     BackColor       =   -2147483643
                  End
                  Begin ComctlLib.ImageCombo ImgCombo 
                     DataField       =   "CsLblForClr"
                     Height          =   330
                     Index           =   49
                     Left            =   2040
                     TabIndex        =   88
                     Top             =   450
                     Width           =   1755
                     _ExtentX        =   3096
                     _ExtentY        =   582
                     _Version        =   393216
                     ForeColor       =   0
                     BackColor       =   -2147483643
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Main Label Color"
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
                     Height          =   330
                     Index           =   88
                     Left            =   60
                     TabIndex        =   129
                     Top             =   780
                     Width           =   1995
                  End
                  Begin VB.Label Label14 
                     BackStyle       =   0  'Transparent
                     BorderStyle     =   1  'Fixed Single
                     Caption         =   "  Color Setting For Label Control"
                     Height          =   255
                     Left            =   -30
                     TabIndex        =   128
                     Top             =   90
                     Width           =   3945
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Label Fore Color"
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
                     Height          =   330
                     Index           =   54
                     Left            =   60
                     TabIndex        =   127
                     Top             =   450
                     Width           =   1995
                  End
               End
            End
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   3
            Left            =   1155
            TabIndex        =   1
            ToolTipText     =   "Enter Template Description"
            Top             =   300
            Width           =   5985
            _ExtentX        =   10557
            _ExtentY        =   503
            MaxLength       =   45
            DataField       =   "CsName"
            IdName          =   "CSNAME"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   19
            Left            =   8415
            TabIndex        =   2
            ToolTipText     =   "Specify Whether This Recor Is Editable (Enter 'Y'/ 'N')"
            Top             =   300
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   503
            DataType        =   4
            DataField       =   "CsEdit"
            IdName          =   "CSEDIT"
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Description"
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
            Left            =   30
            TabIndex        =   232
            Top             =   300
            Width           =   1065
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Editable"
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
            Index           =   79
            Left            =   7590
            TabIndex        =   231
            Top             =   300
            Width           =   885
         End
      End
      Begin VB.PictureBox PicClr 
         Height          =   6345
         Left            =   4080
         ScaleHeight     =   6285
         ScaleWidth      =   4395
         TabIndex        =   109
         Top             =   630
         Visible         =   0   'False
         Width           =   4455
         Begin MSComDlg.CommonDialog CommDialog 
            Left            =   0
            Top             =   30
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin ComctlLib.ImageList ImgLst 
            Left            =   0
            Top             =   750
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            MaskColor       =   12632256
            _Version        =   393216
         End
      End
   End
End
Attribute VB_Name = "EmrFrmClrPic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----   General Comments on the Entry Form  -------'
'
'   One Menu 'C'ustomer Master, 'S'upplier, 'P'rice List
'   Adc(CmCtg), invisible, in the Key Frame
'   Tab Page, TabDet with 5 Tabs
'     1.Customer Detail   2. Instructions
'     3.Factors, 2 frames within for display purpose only.
'     4.Cust Design, Grid GrdCustDsg
'     5.Cust Rm , Grid GrdCustRm
'     6.Cust Alloy, Grid GrdCustAly with CmdCpy button to Copy into grid.
'   X - ClrSet
'   private sub CpyAlyDet for Copying Alloy details into grid.
'
'----


Option Explicit
Dim moCn As MwfLib.MDOConnection   '*** (Bef speed) bef 24/11/06

''*** For the Toggle Options of Command Buttons
Private Enum en_ClrPicFra
  Dets = 0
End Enum

'*** 14/05/05
Dim mColl_ImgCombo As Collection
'*** 14/05/05
Dim ms_Path As String
Private Sub ADC_Load()
  '*** Set the First Key Control to Template Code
  '*** Set the First Non Key Control to Template Description
  '*** Set the Child Property Of ADC to nothing
  '*** Set the Previous and Next Control Properties for ADC
  '*** Set the frame FraNKeyAll as the top most frame
  '*** Set the current tab as 0 (This is done so that whatever be the active tab during
      'design time, the first tab is set as the active tab at run time)
  
  Set moCn = ADC.Connection    '*** (Bef speed) bef 24/11/06
  Set ADC.FirKeyCtl = ADC("CsCd")
  Set ADC.FirNKeyCtl = ADC("CsName")
           
  ADC.Child = ""
           
  FraNKeyAll.ZOrder (0)
  TabDet.Tab = 0
    
  '*** 14/05/05
  Dim wi_i As Integer
  Set mColl_ImgCombo = New Collection
  For wi_i = 0 To ImgCombo.Count - 1
    mColl_ImgCombo.Add wi_i, ImgCombo(wi_i).DataField
  Next wi_i
  '*** 14/05/05
    
End Sub
Private Sub ADC_KeyWhen()
  '*** Disable all the command buttons (Because since the focus is on the key frame and
      'the user accidently clicks any of the command buttons then it might give a
      'run time error. So it is better to disable all the command buttons when the focus is
      'in the Key Frame.
  '*** Hide all frames except the frame FraNKeyAll
  
  Call EnaDisaCmds(True)
  Call HideAllFras
  
  '*** 16/05/05
  CmdAdcBtnPicSample.Pic = LoadPicture("")
  CmdBtnPicSet1Sample.Pic = LoadPicture("")
  CmdBtnPicSet2Sample.Pic = LoadPicture("")
  CmdTlBarBtnPicSample.Pic = LoadPicture("")
  '*** 16/05/05
  
  '*** 14/06/05
  'CmdAdcBtnPicSample.Visible = False
  'CmdBtnPicSet1Sample.Visible = False
  'CmdBtnPicSet2Sample.Visible = False
  'CmdTlBarBtnPicSample.Visible = False
  CmdAdcBtnPicSample.Caption = ""
  CmdBtnPicSet1Sample.Caption = ""
  CmdBtnPicSet2Sample.Caption = ""
  CmdTlBarBtnPicSample.Caption = ""
  '*** 14/06/05
  
  '*** 23/05/05
  PicBoxCsFlxPic.Picture = LoadPicture("")
  PicBoxCsHlpPic.Picture = LoadPicture("")
  PicBoxCsFormPic.Picture = LoadPicture("")
  PicBoxCsMDIPic.Picture = LoadPicture("")
  '*** 23/05/05
End Sub

'Private Sub ADC_KeyValid(pr_Cancel As Boolean, pr_ErrMsg As String)
'  If moCn.GetFldVal("Select HCtbMod From Head where HCoCd= '" + gs_CoCd + "' and " + _
'     "HCd= '" + ctSelfCmCd + "'") = "N" And ADC("CmCtg") = "T" Then _
'     pr_Cancel = True: pr_ErrMsg = "Cannot Enter Contractor Code As Contractor Module Option Is Not Selected In Head File"
'End Sub
Private Sub ADC_NonKeyWhen(Cancel As Boolean, ErrMsg As String)
  '*** Enable all the command buttons except CmdCpy
  '*** If CmCtg= 'C' then
      'Set the default value of CmMulBy as 1, CmFixPrc as hCmFixPrc from Head
      'also set the default values for the fields CmFixPrc, CmGldAs, CmIWtEqOrd, CmIWtFrOrd,
      'CmIGldRtEqOrd, CmDiaRtFA, CmCSRtFA from the Head file
  '*** Disable all the tab frames except the frame of the current tab of the Tab Page

  Call EnaDisaCmds(False)
  CmdCpy.Enabled = True

 
  Dim wo_RecSetClr As MwfLib.MDORowSet
  Dim ws_ClrCd As String, wi_i As Integer, wi_j As Integer
  
  ImgLst.ListImages.Clear
  
  '*** ### Tmp Code (21/03/05)
  If ADC.AddFndMode = xaddmode Then
     For wi_i = 0 To ImgCombo.Count - 1
       ImgCombo(wi_i).text = "&H00000000&"
     Next wi_i
  Else
    'CmdAdcBtnPicSample.Pic = LoadPicture(IIF(ADC("CsADCBtnPic") <> "" And Dir(ms_Path & ADC("CsADCBtnPic")) <> "", ms_Path & ADC("CsADCBtnPic"), ""))
    CmdAdcBtnPicSample.Pic = LoadPicture(IIF(ADC("CsADCBtnPic") <> "" And IsValidPath(ms_Path & ADC("CsADCBtnPic")) = True, ms_Path & ADC("CsADCBtnPic"), ""))
    CmdBtnPicSet1Sample.Pic = LoadPicture(IIF(ADC("CsBtnPicSet1") <> "" And IsValidPath(ms_Path & ADC("CsBtnPicSet1")) = True, ms_Path & ADC("CsBtnPicSet1"), ""))
    CmdBtnPicSet2Sample.Pic = LoadPicture(IIF(ADC("CsBtnPicSet2") <> "" And IsValidPath(ms_Path & ADC("CsBtnPicSet2")) = True, ms_Path & ADC("CsBtnPicSet2"), ""))
    CmdTlBarBtnPicSample.Pic = LoadPicture(IIF(ADC("CsTlBarPic") <> "" And IsValidPath(ms_Path & ADC("CsTlBarPic")) = True, ms_Path & ADC("CsTlBarPic"), ""))
    PicBoxCsFlxPic.Picture = LoadPicture(IIF(ADC("CsFlxPic") <> "" And IsValidPath(ms_Path & ADC("CsFlxPic")) = True, ms_Path & ADC("CsFlxPic"), ""))
    PicBoxCsHlpPic.Picture = LoadPicture(IIF(ADC("CsHlpPic") <> "" And IsValidPath(ms_Path & ADC("CsHlpPic")) = True, ms_Path & ADC("CsHlpPic"), ""))
    PicBoxCsFormPic.Picture = LoadPicture(IIF(ADC("CsFormPic") <> "" And IsValidPath(ms_Path & ADC("CsFormPic")) = True, ms_Path & ADC("CsFormPic"), ""))
    PicBoxCsMDIPic.Picture = LoadPicture(IIF(ADC("CsMDIPic") <> "" And IsValidPath(ms_Path & ADC("CsMDIPic")) = True, ms_Path & ADC("CsMDIPic"), ""))
    
'    CmdAdcBtnPicSample.Pic = LoadPicture(ADC("CsADCBtnPic"))
'    CmdBtnPicSet1Sample.Pic = LoadPicture(ADC("CsBtnPicSet1"))
'    CmdBtnPicSet2Sample.Pic = LoadPicture(ADC("CsBtnPicSet2"))
'    CmdTlBarBtnPicSample.Pic = LoadPicture(ADC("CsTlBarPic"))
'    PicBoxCsFlxPic.Picture = LoadPicture(ADC("CsFlxPic"))
'    PicBoxCsHlpPic.Picture = LoadPicture(ADC("CsHlpPic"))
'    PicBoxCsFormPic.Picture = LoadPicture(ADC("CsFormPic"))
'    PicBoxCsMDIPic.Picture = LoadPicture(ADC("CsMDIPic"))
    
    '*** 14/06/05
    If Not (CmdAdcBtnPicSample.Pic Is Nothing Or CmdAdcBtnPicSample.Pic = 0) Then CmdAdcBtnPicSample.Caption = "Caption"
    If Not (CmdBtnPicSet1Sample.Pic Is Nothing Or CmdBtnPicSet1Sample.Pic = 0) Then CmdBtnPicSet1Sample.Caption = "Caption"
    If Not (CmdBtnPicSet2Sample.Pic Is Nothing Or CmdBtnPicSet2Sample.Pic = 0) Then CmdBtnPicSet2Sample.Caption = "Caption"
    If Not (CmdTlBarBtnPicSample.Pic Is Nothing Or CmdTlBarBtnPicSample.Pic = 0) Then CmdTlBarBtnPicSample.Caption = "Caption"
    '*** 14/06/05
  End If
  '*** ### Tmp Code (21/03/05)
  
  For wi_i = 0 To ImgCombo.Count - 1
    ImgCombo(wi_i).ComboItems.Clear
    ImgCombo(wi_i).Tag = ImgCombo(wi_i).text    '*** 21/03/05
  Next wi_i
  
  Set wo_RecSetClr = moCn.OpenRes("Select CfCd, CfShtName from ClrDfn ")
  
  ImgLst.ListImages.Add , UCase("ZCustom ..."), CreateColorImage(PicClr, 16777215, 18, 14)
  Do While Not (wo_RecSetClr.EOF Or wo_RecSetClr.BOF)
    ws_ClrCd = UCase(wo_RecSetClr.FldValue("CfCd"))
    ImgLst.ListImages.Add , "Z" & ws_ClrCd, CreateColorImage(PicClr, Val(ws_ClrCd), 18, 14)
    wo_RecSetClr.MoveNext
  Loop
  
  For wi_i = 0 To ImgCombo.Count - 1
    If Trim(ImgCombo(wi_i).text) <> "" And GetClrKey(Trim(ImgCombo(wi_i).text)) = "" Then
      ImgLst.ListImages.Add , "Z" & UCase(ImgCombo(wi_i).text), CreateColorImage(PicClr, Val(ImgCombo(wi_i).text), 16, 16)
    End If
  Next wi_i
  
  For wi_i = 0 To ImgCombo.Count - 1
    ImgCombo(wi_i).ImageList = ImgLst
  Next wi_i
  
  For wi_i = 0 To ImgCombo.Count - 1
    For wi_j = 1 To ImgLst.ListImages.Count
      'ImgCombo(wi_i).ComboItems.Add wi_j, ImgLst.ListImages.Item(wi_j).Key, "aaa", ImgLst.ListImages.Item(wi_j).Key
      ImgCombo(wi_i).ComboItems.Add wi_j, UCase(ImgLst.ListImages.Item(wi_j).Key), Mid(ImgLst.ListImages.Item(wi_j).Key, 2, Len(ImgLst.ListImages.Item(wi_j).Key) - 1), ImgLst.ListImages.Item(wi_j).Key
    Next wi_j
    If Trim(ImgCombo(wi_i).text) <> "" Then ImgCombo(wi_i).ComboItems("Z" & UCase(ImgCombo(wi_i).text)).Selected = True '*** this code is also proper
  Next wi_i
  
  For wi_i = 0 To TabDet.Tabs - 1
    PicBoxTabDet(wi_i).Enabled = False
  Next wi_i
  PicBoxTabDet(TabDet.Tab).Enabled = True
  
'  '*** ### Tmp Code (Bef 21/03/05)
'  If ADC.AddFndMode = xaddmode Then
'     For wi_i = 0 To ImgCombo.Count - 1
'       ImgCombo(wi_i).Text = "&H00000000&"
'     Next wi_i
'  End If
'  '*** ### Tmp Code (Bef 21/03/05)
  
  
  
End Sub

Private Sub ATXT_FldWhen(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String)
  '*** Set Helps For Each Field
  Select Case UCase(IdName)
  Case Is = UCase("CsCd")
    Call HlpList.CsCd
  Case Is = UCase("wCpyCsCd")
    Call HlpList.CsCd
  Case Is = UCase("Cs3D")
    Call HlpList.PMCd("YN")
  Case Is = UCase("CsGrdInset")
    Call HlpList.PMCd("YN")
  Case Is = UCase("CsEdit")
    Call HlpList.PMCd("YN")
  End Select
End Sub
Private Sub atxt_ReCalculate(Index As Integer, ByVal IdName As String, ByVal OldVal As Variant, ByVal RowNum As Integer, ByVal SourceName As String, ByVal SourceColName As String)
  '*** 16/05/05
  Select Case UCase(IdName)
  Case Is = UCase("wCsADCBtnPic")
    If ADC.Mode = xNorm Or ADC.Mode = xInit Or ADC.Mode = xCopy Then
      Call ViewSampleBtn(CmdAdcBtnPicSample)
    End If
  Case Is = UCase("wCsBtnPicSet1")
    If ADC.Mode = xNorm Or ADC.Mode = xInit Or ADC.Mode = xCopy Then
      Call ViewSampleBtn(CmdBtnPicSet1Sample)
    End If
  Case Is = UCase("wCsBtnPicSet2")
    If ADC.Mode = xNorm Or ADC.Mode = xInit Or ADC.Mode = xCopy Then
      Call ViewSampleBtn(CmdBtnPicSet2Sample)
    End If
  Case Is = UCase("wCsTlBarPic")
    If ADC.Mode = xNorm Or ADC.Mode = xInit Or ADC.Mode = xCopy Then
      Call ViewSampleBtn(CmdTlBarBtnPicSample)
    End If
  End Select
  '*** 16/05/05
End Sub
Private Sub atxt_FldChange(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String, ByVal pv_NewValue As Variant)
  '*** CsCd             Not Blank
  '*** CsName           Not Blank

  Select Case UCase(IdName)
  Case Is = UCase("CsCd")
    If Trim(pv_NewValue) = "" Then Cancel = True: ErrMsg = "Template Code Cannot Be Blank": Exit Sub

    Cancel = moCn.RecSeek("select CsCd from ClrSet where CsCd = '" + pv_NewValue + "' and CsEdit= 'N' ")
    If Cancel = True Then ErrMsg = "This Color Cannot Be Edited": Exit Sub
  Case Is = UCase("CsName")
    If Trim(pv_NewValue) = "" Then Cancel = True: ErrMsg = "Template Description Cannot Be Blank": Exit Sub
  Case Is = UCase("wCpyCsCd")
    Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("select CsCd from ClrSet where CsCd = '" + pv_NewValue + "'"))
    ErrMsg = "Invalid Template Code": Exit Sub
  Case Is = UCase("Cs3D")
    Cancel = Not moCn.RecSeek("select PMCd from Param where PTyp= 'YN' and " + _
                         "PMCd = '" + pv_NewValue + "'")
    ErrMsg = "Invalid Object Style Option (3D or Flat)": Exit Sub
  Case Is = UCase("CsGrdInset")
    Cancel = Not moCn.RecSeek("select PMCd from Param where PTyp= 'YN' and " + _
                         "PMCd = '" + pv_NewValue + "'")
    ErrMsg = "Invalid Grid Heading Option (Inset or Flat)": Exit Sub
  Case Is = UCase("CsEdit")
    Cancel = Not moCn.RecSeek("select PMCd from Param where PTyp= 'YN' and " + _
                         "PMCd = '" + pv_NewValue + "'")
    ErrMsg = "Invalid Editable Option ('Y' For Editable/ 'N' For Noneditable)": Exit Sub
  Case Is = UCase("CsADCBtnPic")
    'If pv_NewValue <> "" And Dir(Trim(ms_Path & pv_NewValue)) = "" Then
    If pv_NewValue <> "" And IsValidPath(ms_Path & pv_NewValue) = False Then
      Cancel = True
      ErrMsg = "Invalid ADC Button Picture Path": Exit Sub
    End If
    'CmdAdcBtnPicSample.Pic = LoadPicture(pv_NewValue)
    'CmdAdcBtnPicSample.Pic = LoadPicture(IIF(pv_NewValue <> "" And Dir(ms_Path & pv_NewValue) <> "", ms_Path & pv_NewValue, ""))
    CmdAdcBtnPicSample.Pic = LoadPicture(IIF(pv_NewValue <> "", ms_Path & pv_NewValue, ""))
  Case Is = UCase("CsFlxPic")
    If pv_NewValue <> "" And IsValidPath(ms_Path & pv_NewValue) = False Then
      Cancel = True
      ErrMsg = "Invalid Grid Row Pointer Picture Path": Exit Sub
    End If
    'PicBoxCsFlxPic.Picture = LoadPicture(pv_NewValue)
    PicBoxCsFlxPic.Picture = LoadPicture(IIF(pv_NewValue <> "", ms_Path & pv_NewValue, ""))
  Case Is = UCase("CsHlpPic")
    If pv_NewValue <> "" And IsValidPath(ms_Path & pv_NewValue) = False Then
      Cancel = True
      ErrMsg = "Invalid Help Grid Row Pointer Picture Path": Exit Sub
    End If
    'PicBoxCsHlpPic.Picture = LoadPicture(pv_NewValue)
    PicBoxCsHlpPic.Picture = LoadPicture(IIF(pv_NewValue <> "", ms_Path & pv_NewValue, ""))
  Case Is = UCase("CsTlBarPic")
    If pv_NewValue <> "" And IsValidPath(ms_Path & pv_NewValue) = False Then
      Cancel = True
      ErrMsg = "Invalid Tool Bar Button Picture Path": Exit Sub
    End If
    'CmdTlBarBtnPicSample.Pic = LoadPicture(pv_NewValue)
    CmdTlBarBtnPicSample.Pic = LoadPicture(IIF(pv_NewValue <> "", ms_Path & pv_NewValue, ""))
  Case Is = UCase("CsFormPic")
    If pv_NewValue <> "" And IsValidPath(ms_Path & pv_NewValue) = False Then
      Cancel = True
      ErrMsg = "Invalid Form Picture Path": Exit Sub
    End If
    'PicBoxCsFormPic.Picture = LoadPicture(pv_NewValue)
    PicBoxCsFormPic.Picture = LoadPicture(IIF(pv_NewValue <> "", ms_Path & pv_NewValue, ""))
  Case Is = UCase("CsBtnPicSet1")
    If pv_NewValue <> "" And IsValidPath(ms_Path & pv_NewValue) = False Then
      Cancel = True
      ErrMsg = "Invalid Button (Set1) Picture Path": Exit Sub
    End If
    'CmdBtnPicSet1Sample.Pic = LoadPicture(pv_NewValue)
    CmdBtnPicSet1Sample.Pic = LoadPicture(IIF(pv_NewValue <> "", ms_Path & pv_NewValue, ""))
  Case Is = UCase("CsBtnPicSet2")
    If pv_NewValue <> "" And IsValidPath(ms_Path & pv_NewValue) = False Then
      Cancel = True
      ErrMsg = "Invalid Button (Set2) Picture Path": Exit Sub
    End If
    'CmdBtnPicSet2Sample.Pic = LoadPicture(pv_NewValue)
    CmdBtnPicSet2Sample.Pic = LoadPicture(IIF(pv_NewValue <> "", ms_Path & pv_NewValue, ""))
  Case Is = UCase("CsMDIPic")
    If pv_NewValue <> "" And IsValidPath(ms_Path & pv_NewValue) = False Then
      Cancel = True
      ErrMsg = "Invalid MDI Form Picture Path": Exit Sub
    End If
    'PicBoxCsMDIPic.Picture = LoadPicture(pv_NewValue)
    PicBoxCsMDIPic.Picture = LoadPicture(IIF(pv_NewValue <> "", ms_Path & pv_NewValue, ""))
  End Select
End Sub
Private Sub ImgCombo_Click(Index As Integer)
  Dim wi_i As Integer
  If UCase(ImgCombo(Index).SelectedItem.Key) = UCase("ZCustom ...") Then
    ImgCombo(Index).text = GetColor
    If Trim(ImgCombo(Index).text) <> "" Then
      If GetClrKey(ImgCombo(Index).text) = "" Then
        ImgLst.ListImages.Add , "Z" & UCase(ImgCombo(Index).text), CreateColorImage(PicClr, Val(ImgCombo(Index).text), 18, 14)
        For wi_i = 0 To ImgCombo.Count - 1
          ImgCombo(wi_i).ComboItems.Add , "Z" & UCase(ImgCombo(Index).text), ImgCombo(Index).text, "Z" & UCase(ImgCombo(Index).text)
        Next wi_i
        'ImgCombo(Index).ComboItems.Add , "Z" & UCase(ImgCombo(Index).Text), ImgCombo(Index).Text, "Z" & UCase(ImgCombo(Index).Text)
      End If
      ImgCombo(Index).ComboItems("Z" & UCase(ImgCombo(Index).text)).Selected = True
    Else
      ImgCombo(Index).text = ImgCombo(Index).Tag     '*** 21/03/05
    End If
  End If
  ImgCombo(Index).Tag = ImgCombo(Index).text     '*** 21/03/05
  

  '*** 16/05/05
  Call ComboSel(Index)
  '*** 16/05/05
End Sub
Private Sub ImgCombo_Change(Index As Integer)
  '*** 16/05/05
  Call ComboSel(Index)
  '*** 16/05/05
End Sub
Private Sub ADC_NonKeyValid(pr_Cancel As Boolean, pr_ErrMsg As String)
  '*** Set the current tab of the tabpage as 0
  TabDet.Tab = 0
End Sub
Private Sub ADC_SetRecSource()
  '*** Set the RecSource for ADC  (Base Table is ClrSet)
  ADC.RecSource = " Select * from ClrSet " + _
                  " Where CsCd = '" + ADC("CsCd") + "' "
  
  'ADC.RecSource = " Select * from ClrSet " + _
  '                " Where CsCd = '" + ADC("CsCd") + "' and CsEdit= 'Y' "
End Sub
Private Sub TabDet_Click(PreviousTab As Integer)
  '*** Whenever a tab is selected enable the frame on the clicked tab and disable the previous tab frame
  '*** For this to work the index of the frame and the Tab Number should be the same

  PicBoxTabDet(TabDet.Tab).Enabled = True
  PicBoxTabDet(PreviousTab).Enabled = False
End Sub
Private Sub CmdCpy_Click()
  '*** Display or hide the Copy frame depending on the toggle satate
  Call DispFra(Dets)
End Sub
Private Sub CmdCpyGo_Click()
  If ADC("wCpyCsCd") = "" Then DispMsg "A Valid Template Code Should Be Entered", etError: Exit Sub

  Call CpyDet
  Call CmdCpy_Click
End Sub
Private Sub CmdAdcBtnPic_Click()
  Dim ws_FilePth As String
  ws_FilePth = OpenPicFilePath
  
  'If Trim(ws_FilePth) <> "" And Dir(ms_Path & ws_FilePth) <> "" Then
  If Trim(ws_FilePth) <> "" And IsValidPath(ms_Path & ws_FilePth) = True Then
    ADC("CsADCBtnPic") = ws_FilePth
    CmdAdcBtnPicSample.Pic = LoadPicture(ms_Path & ws_FilePth)
    'PicBoxCsADCBtnPic.Picture = LoadPicture(ws_FilePth)
    Call ViewSampleBtn(CmdAdcBtnPicSample)    '*** 16/05/05
    
    '*** 14/06/05
    If CmdAdcBtnPicSample.Pic Is Nothing Or CmdAdcBtnPicSample.Pic = 0 Then
      CmdAdcBtnPicSample.Caption = ""
    Else
      CmdAdcBtnPicSample.Caption = "Caption"
    End If
    '*** 14/06/05
    
  End If
End Sub
Private Sub CmdGrdPointPic_Click()
  Dim ws_FilePth As String
  ws_FilePth = OpenPicFilePath
  'If Trim(ws_FilePth) <> "" And Dir(ms_Path & ws_FilePth) <> "" Then
  If Trim(ws_FilePth) <> "" And IsValidPath(ms_Path & ws_FilePth) = True Then
    ADC("CsFlxPic") = ws_FilePth
    PicBoxCsFlxPic.Picture = LoadPicture(ms_Path & ws_FilePth)
  End If
End Sub
Private Sub CmdHlpPointPic_Click()
  Dim ws_FilePth As String
  ws_FilePth = OpenPicFilePath
  If Trim(ws_FilePth) <> "" And IsValidPath(ms_Path & ws_FilePth) = True Then
    ADC("CsHlpPic") = ws_FilePth
    PicBoxCsHlpPic.Picture = LoadPicture(ms_Path & ws_FilePth)
  End If
End Sub
Private Sub CmdTlBarBtnPic_Click()
  Dim ws_FilePth As String
  ws_FilePth = OpenPicFilePath
  If Trim(ws_FilePth) <> "" And IsValidPath(ms_Path & ws_FilePth) = True Then
    ADC("CsTlBarPic") = ws_FilePth
    CmdTlBarBtnPicSample.Pic = LoadPicture(ms_Path & ws_FilePth)
    'PicBoxCsTlBarPic.Picture = LoadPicture(ws_FilePth)
    Call ViewSampleBtn(CmdTlBarBtnPicSample)    '*** 16/05/05
  
    '*** 14/06/05
    If CmdTlBarBtnPicSample.Pic Is Nothing Or CmdTlBarBtnPicSample.Pic = 0 Then
      CmdTlBarBtnPicSample.Caption = ""
    Else
      CmdTlBarBtnPicSample.Caption = "Caption"
    End If
    '*** 14/06/05
  
  End If
End Sub
Private Sub CmdFrmPic_Click()
  Dim ws_FilePth As String
  ws_FilePth = OpenPicFilePath
  If Trim(ws_FilePth) <> "" And IsValidPath(ms_Path & ws_FilePth) = True Then
    ADC("CsFormPic") = ws_FilePth
    PicBoxCsFormPic.Picture = LoadPicture(ms_Path & ws_FilePth)
  End If
End Sub
Private Sub CmdMDIPic_Click()
  Dim ws_FilePth As String
  ws_FilePth = OpenPicFilePath
  If Trim(ws_FilePth) <> "" And IsValidPath(ms_Path & ws_FilePth) = True Then
  
    'Clipboard.Clear                '*** Clear Clipboard.      '??????????
    'Clipboard.SetData LoadPicture(ms_Path & ws_FilePth), vbCFBitmap     '??????????
    'Image1.Stretch = False
    'Image1.Picture = Clipboard.GetData(vbCFBitmap) '*** Copy from Clipboard
  
    'Set Image1.Picture = LoadPicture(ws_FilePth, vbLPCustom, vbLPColor, 100, 100)
    'CsMDIPic.Picture = LoadPicture(ws_FilePth, , vbLPColor)
    ADC("CsMDIPic") = ws_FilePth
    PicBoxCsMDIPic.Picture = LoadPicture(ms_Path & ws_FilePth)
  End If
End Sub
Private Sub CmdBtnPicSet1_Click()
  Dim ws_FilePth As String
  ws_FilePth = OpenPicFilePath
  If Trim(ws_FilePth) <> "" And IsValidPath(ms_Path & ws_FilePth) = True Then
    ADC("CsBtnPicSet1") = ws_FilePth
    CmdBtnPicSet1Sample.Pic = LoadPicture(ms_Path & ws_FilePth)
    'PicBoxCsBtnPicSet1.Picture = LoadPicture(ws_FilePth)
    Call ViewSampleBtn(CmdBtnPicSet1Sample)    '*** 16/05/05
  
    '*** 14/06/05
    If CmdBtnPicSet1Sample.Pic Is Nothing Or CmdBtnPicSet1Sample.Pic = 0 Then
      CmdBtnPicSet1Sample.Caption = ""
    Else
      CmdBtnPicSet1Sample.Caption = "Caption"
    End If
    '*** 14/06/05
  
  End If
End Sub
Private Sub CmdBtnPicSet2_Click()
  Dim ws_FilePth As String
  ws_FilePth = OpenPicFilePath
  If Trim(ws_FilePth) <> "" And IsValidPath(ms_Path & ws_FilePth) = True Then
    ADC("CsBtnPicSet2") = ws_FilePth
    CmdBtnPicSet2Sample.Pic = LoadPicture(ms_Path & ws_FilePth)
    'PicBoxCsBtnPicSet2.Picture = LoadPicture(ws_FilePth)
    Call ViewSampleBtn(CmdBtnPicSet2Sample)    '*** 16/05/05
    
    '*** 14/06/05
    If CmdBtnPicSet2Sample.Pic Is Nothing Or CmdBtnPicSet2Sample.Pic = 0 Then
      CmdBtnPicSet2Sample.Caption = ""
    Else
      CmdBtnPicSet2Sample.Caption = "Caption"
    End If
    '*** 14/06/05
    
  End If
End Sub
Private Function GetClrKey(ByVal ps_ClrCd As String) As String
  GetClrKey = ""
  
  Dim wi_i As Integer
  For wi_i = 1 To ImgLst.ListImages.Count
    If UCase("Z" & ps_ClrCd) = UCase(ImgLst.ListImages.Item(wi_i).Key) Then GetClrKey = ImgLst.ListImages.Item(wi_i).Key: Exit For
  Next wi_i
End Function
Private Sub CpyDet()
  '*** Copy Option is used to copy information from another template

  Dim wi_i As Integer
  Dim wi_j As Integer     '*** (21/03/05)
  
  '*** Beginning the copy process
  ADC.StartCopy
  MWLib.BeginProcess Me, "Copying Details ..."
  Dim wRsTemplate As MwfLib.MDORowSet
  Set wRsTemplate = moCn.OpenRes("Select * from ClrSet where CsCd= '" + ADC("wCpyCsCd") + "' ")
  With wRsTemplate
    Do While Not (.EOF Or .BOF)
      MWLib.RunningProcess "Copying Details ..."

      '*** (Bef 23/05/05)
      'PicBoxCsADCBtnPic.Picture = wRsTemplate.FldPicValue("CsADCBtnPic")
      'PicBoxCsFlxPic.Picture = wRsTemplate.FldPicValue("CsFlxPic")
      'PicBoxCsHlpPic.Picture = wRsTemplate.FldPicValue("CsHlpPic")
      'PicBoxCsTlBarPic.Picture = wRsTemplate.FldPicValue("CsTlBarPic")
      'PicBoxCsFormPic.Picture = wRsTemplate.FldPicValue("CsFormPic")
      'PicBoxCsBtnPicSet1.Picture = wRsTemplate.FldPicValue("CsBtnPicSet1")
      'PicBoxCsBtnPicSet2.Picture = wRsTemplate.FldPicValue("CsBtnPicSet2")
      'PicBoxCsMDIPic.Picture = wRsTemplate.FldPicValue("CsMDIPic")
      '*** (Bef 23/05/05)

      '*** (23/05/05)
      ADC("CsADCBtnPic") = wRsTemplate!CsADCBtnPic
      ADC("CsFlxPic") = wRsTemplate!CsFlxPic
      ADC("CsHlpPic") = wRsTemplate!CsHlpPic
      ADC("CsTlBarPic") = wRsTemplate!CsTlBarPic
      ADC("CsFormPic") = wRsTemplate!CsFormPic
      ADC("CsBtnPicSet1") = wRsTemplate!CsBtnPicSet1
      ADC("CsBtnPicSet2") = wRsTemplate!CsBtnPicSet2
      ADC("CsMDIPic") = wRsTemplate!CsMDIPic
      
      
      'CmdAdcBtnPicSample.Pic = LoadPicture(IIF(wRsTemplate!CsADCBtnPic <> "" And Dir(ms_Path & wRsTemplate!CsADCBtnPic) <> "", ms_Path & wRsTemplate!CsADCBtnPic, ""))  'LoadPicture(wRsTemplate!CsADCBtnPic)
      CmdAdcBtnPicSample.Pic = LoadPicture(IIF(wRsTemplate!CsADCBtnPic <> "" And IsValidPath(ms_Path & wRsTemplate!CsADCBtnPic) = True, ms_Path & Trim(wRsTemplate!CsADCBtnPic), ""))     'LoadPicture(wRsTemplate!CsADCBtnPic)
      CmdTlBarBtnPicSample.Pic = LoadPicture(IIF(Trim(wRsTemplate!CsTlBarPic) <> "" And IsValidPath(ms_Path & Trim(wRsTemplate!CsTlBarPic)) = True, ms_Path & Trim(wRsTemplate!CsTlBarPic), ""))
      CmdBtnPicSet1Sample.Pic = LoadPicture(IIF(Trim(wRsTemplate!CsBtnPicSet1) <> "" And IsValidPath(ms_Path & Trim(wRsTemplate!CsBtnPicSet1)) = True, ms_Path & Trim(wRsTemplate!CsBtnPicSet1), ""))
      CmdBtnPicSet2Sample.Pic = LoadPicture(IIF(Trim(wRsTemplate!CsBtnPicSet2) <> "" And IsValidPath(ms_Path & Trim(wRsTemplate!CsBtnPicSet2)) = True, ms_Path & Trim(wRsTemplate!CsBtnPicSet2), ""))
      PicBoxCsFlxPic.Picture = LoadPicture(IIF(Trim(wRsTemplate!CsFlxPic) <> "" And IsValidPath(ms_Path & Trim(wRsTemplate!CsFlxPic)) = True, ms_Path & Trim(wRsTemplate!CsFlxPic), ""))
      PicBoxCsHlpPic.Picture = LoadPicture(IIF(Trim(wRsTemplate!CsHlpPic) <> "" And IsValidPath(ms_Path & Trim(wRsTemplate!CsHlpPic)) = True, ms_Path & Trim(wRsTemplate!CsHlpPic), ""))
      PicBoxCsFormPic.Picture = LoadPicture(IIF(Trim(wRsTemplate!CsFormPic) <> "" And IsValidPath(ms_Path & Trim(wRsTemplate!CsFormPic)) = True, ms_Path & Trim(wRsTemplate!CsFormPic), ""))
      PicBoxCsMDIPic.Picture = LoadPicture(IIF(Trim(wRsTemplate!CsMDIPic) <> "" And IsValidPath(ms_Path & Trim(wRsTemplate!CsMDIPic)) = True, ms_Path & Trim(wRsTemplate!CsMDIPic), ""))
      '*** (23/05/05)

      '*** 14/06/05
      CmdAdcBtnPicSample.Caption = IIF(CmdAdcBtnPicSample.Pic Is Nothing Or CmdAdcBtnPicSample.Pic = 0, "", "Caption")
      CmdBtnPicSet1Sample.Caption = IIF(CmdBtnPicSet1Sample.Pic Is Nothing Or CmdBtnPicSet1Sample.Pic = 0, "", "Caption")
      CmdBtnPicSet2Sample.Caption = IIF(CmdBtnPicSet2Sample.Pic Is Nothing Or CmdBtnPicSet2Sample.Pic = 0, "", "Caption")
      CmdTlBarBtnPicSample.Caption = IIF(CmdTlBarBtnPicSample.Pic Is Nothing Or CmdTlBarBtnPicSample.Pic = 0, "", "Caption")
      '*** 14/06/05

      For wi_i = 0 To ImgCombo.Count - 1
        ImgCombo(wi_i).text = wRsTemplate(ImgCombo(wi_i).DataField)
        ImgCombo(wi_i).Tag = ImgCombo(wi_i).text
        
        '*** (21/03/05)
        If GetClrKey(ImgCombo(wi_i).text) = "" Then
          ImgLst.ListImages.Add , "Z" & UCase(ImgCombo(wi_i).text), CreateColorImage(PicClr, Val(ImgCombo(wi_i).text), 18, 14)
          For wi_j = 0 To ImgCombo.Count - 1
            ImgCombo(wi_j).ComboItems.Add , "Z" & UCase(ImgCombo(wi_i).text), ImgCombo(wi_i).text, "Z" & UCase(ImgCombo(wi_i).text)
          Next wi_j
        End If
        ImgCombo(wi_i).ComboItems("Z" & UCase(ImgCombo(wi_i).text)).Selected = True
        '*** (21/03/05)
      Next wi_i

      
      ADC("CsADCBtnHeight") = wRsTemplate!CsADCBtnHeight
      ADC("CsADCBtnWidth") = wRsTemplate!CsADCBtnWidth
      'ADC("CsADCDiffTop") = wRsTemplate!CsADCDiffTop    '*** (Bef 23/05/05)
      'ADC("CsADCDiffLeft") = wRsTemplate!CsADCDiffLeft  '*** (Bef 23/05/05)
      
      ADC("CsBtnHeightSet1") = wRsTemplate!CsBtnHeightSet1
      ADC("CsBtnWidthSet1") = wRsTemplate!CsBtnWidthSet1
      ADC("CsBtnTopDiffSet1") = wRsTemplate!CsBtnTopDiffSet1
      ADC("CsBtnLeftDiffSet1") = wRsTemplate!CsBtnLeftDiffSet1

      ADC("CsBtnHeightSet2") = wRsTemplate!CsBtnHeightSet2
      ADC("CsBtnWidthSet2") = wRsTemplate!CsBtnWidthSet2
      ADC("CsBtnTopDiffSet2") = wRsTemplate!CsBtnTopDiffSet2
      ADC("CsBtnLeftDiffSet2") = wRsTemplate!CsBtnLeftDiffSet2
      
      ADC("Cs3D") = wRsTemplate!Cs3D
      ADC("CsGrdInset") = wRsTemplate!CsGrdInset
    
      ADC("CsTlBarBtnHeight") = wRsTemplate!CsTlBarBtnHeight
      ADC("CsTlBarBtnWidth") = wRsTemplate!CsTlBarBtnWidth
      ADC("CsTlBarBtnGap") = wRsTemplate!CsTlBarBtnGap
      
      ADC("CsEdit") = "Y"
    
      .MoveNext
    Loop
  End With
  MWLib.EndProcess Me
  ADC.EndCopy
  Set wRsTemplate = Nothing
  '*** End of the copy process
End Sub
Private Function GetColor() As String
   '*** Set Cancel to True.
   CommDialog.CancelError = True
   On Error GoTo ErrHandler
   '*** Set the Flags property.
   CommDialog.Flags = cdlCCRGBInit
   '*** Display the Color dialog box.
   CommDialog.ShowColor
   '*** Set the form's background color to the selected color.
   
   GetColor = CommDialog.Color
   
   Exit Function

ErrHandler:
   '*** User pressed Cancel button.
   Exit Function
   
End Function
Private Function OpenPicFilePath() As String
  On Error GoTo ErrHndlr    '##$$
  
  'If InStrRev(App.Path, "\") > 1 Then CommDialog.InitDir = ms_Path
  If InStrRev(ms_Path, "\") > 1 Then CommDialog.InitDir = Mid(ms_Path, 1, InStrRev(ms_Path, "\") - 1)
    
  OpenPicFilePath = ""
  CommDialog.FileName = ""
  
  CommDialog.DialogTitle = "Open Picture File As"
  'CommDialog.InitDir = moCn.GetFldVal("Select PDesc225 from Param where PTyp= 'TRFDIR' and PMCd= 'TRFDIR' ")
  'CommDialog.Filter = "XML Files (*.xml)"
  CommDialog.Filter = "Pictures(*.bmp;*.ico;*.jpg;*.gif;*.jpeg)|*.bmp;*.ico;*.jpg;*.gif;*.jpeg"
  CommDialog.FilterIndex = 1
  
  '##$$ CommDialog.CancelError = True '***
  
  '*** cdlOFNOverwritePrompt:     specifies whether the user should be alerted if he is overwriting and existing file
  '*** cdlOFNFileMustExist:       Specifies that the user can enter only names of existing files in the File Name text box
  '*** cdlOFNPathMustExist:       makes sure that the the user should specify an existing path
  '*** cdlOFNLongNames:           specifies that long names can be used for the files
  '*** cdlOFNExtensionDifferent:  indicates that the user has selected a different extension than the one specified through the DefaultExt property
  '***(Bef 18/06/05)
  CommDialog.Flags = cdlOFNLongNames Or cdlOFNFileMustExist Or cdlOFNPathMustExist Or cdlOFNExtensionDifferent
  '***(Bef 18/06/05)
  '***(18/06/05)
  'CommDialog.Flags = cdlOFNLongNames Or cdlOFNFileMustExist Or cdlOFNPathMustExist Or cdlOFNExtensionDifferent Or cdlOFNNoChangeDir
  '***(18/06/05)
  CommDialog.DefaultExt = "bmp"
  CommDialog.ShowOpen
    
  'If Len(CommDialog.FileName) > 4 And Trim(CommDialog.FileName) <> "" And Dir(CommDialog.FileName) <> "" Then
  '  If UCase(Right(CommDialog.FileName, 4)) = ".XML" Then
  '    OpenPicFilePath = CommDialog.FileName
  '  Else
  '    OpenPicFilePath = ""
  '    DispMsg "File Extension Is Wrong", etError: Exit Function
  '  End If
  'End If

  '***(Bef 18/06/05) OpenPicFilePath = CommDialog.FileName
  '***(18/06/05)
  Dim ws_Path As String, wi_Pos As Integer
  ws_Path = CommDialog.FileName
  If InStrRev(ws_Path, "\") > 1 Then
    wi_Pos = InStrRev(Mid(ws_Path, 1, InStrRev(ws_Path, "\") - 1), "\")
    If wi_Pos > 0 Then OpenPicFilePath = Mid(ws_Path, wi_Pos + 1)
  End If
  '***(18/06/05)
  
  
'##$$
ErrHndlr:
  If Err.Number = 32755 Then
    '*** This error indicates that the user has choosen the Cancel button.
    OpenPicFilePath = ""
    'ADC("wDisp") = "Operation Cancelled "
    'DispMsg "Operation Cancelled ", etInfo
    Err.Clear
  End If
'##$$
End Function
Private Sub DispFra(ByVal pv_CustFra As en_ClrPicFra)
  Select Case pv_CustFra
  Case Is = Dets
    '*** If Frame PicBoxCpy is visible then make it invisible and enable all the command buttons
    '*** If Frame PicBoxCpy is invisible then make it visible and disable all the
        'command buttons except the one which is clicked. Change the background colour of
        'the clicked command button to white and make the font bold
    If PicBoxCpy.Visible = True Then
      PicBoxCpy.Visible = False
      PicBoxCpy.Enabled = False
      Call EnaDisaCmds(False, CmdCpy)
      CmdCpy.SetFocus
    Else
      If ADC.AddFndMode = xaddmode Then
        PicBoxCpy.Visible = True
        PicBoxCpy.Enabled = True
        'ADC("wCustFr").SetFocus
        PicBoxCpy.ZOrder
        Call EnaDisaCmds(True, CmdCpy)
      Else
        DispMsg "Copy Is Allowed Only For A New Record", etWarning
      End If
    End If
  End Select
End Sub
Private Sub EnaDisaCmds(ByVal pv_ShowFra As Boolean, Optional ByVal pv_CmdBut As MWCTL_BTN1)
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
      '### pv_CmdBut.BackColor = vbWhite
      '### pv_CmdBut.FontBold = True
      If pv_CmdBut.Pic = 0 Then
        pv_CmdBut.BackColor = vbWhite
        'pv_CmdBut.FontBold = True
      End If
      pv_CmdBut.CausesValidation = True
    Else
      '### pv_CmdBut.BackColor = &H8000000F
      '### pv_CmdBut.FontBold = False
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
  CmdCpy.Enabled = Not pv_ShowFra
  If Not pv_CmdBut Is Nothing Then pv_CmdBut.Enabled = True
End Sub
Private Sub HideAllFras()
  '*** This procedure hides all the frames other than FraNKeyAll
  FraNKeyAll.Visible = True
  PicBoxCpy.Visible = False
  PicBoxCpy.Enabled = False
End Sub
Private Sub ComboSel(ByVal Index As Integer)
  '*** 16/05/05
  If (mColl_ImgCombo Is Nothing) Then Exit Sub
  If mColl_ImgCombo.Count = 0 Then Exit Sub
  If Val(mColl_ImgCombo("CsADCBtnForClr")) = Index Or Val(mColl_ImgCombo("CsADCBtnMaskClr")) = Index Then
    Call ViewSampleBtn(CmdAdcBtnPicSample)
  End If
  If Val(mColl_ImgCombo("CsBtnForClrSet1")) = Index Or Val(mColl_ImgCombo("CsBtnMaskClrSet1")) = Index Then
    Call ViewSampleBtn(CmdBtnPicSet1Sample)
  End If
  If Val(mColl_ImgCombo("CsBtnForClrSet2")) = Index Or Val(mColl_ImgCombo("CsBtnMaskClrSet2")) = Index Then
    Call ViewSampleBtn(CmdBtnPicSet2Sample)
  End If
  If Val(mColl_ImgCombo("CsTlBarBtnForClr")) = Index Or Val(mColl_ImgCombo("CsTlBarBtnMaskClr")) = Index Then
    Call ViewSampleBtn(CmdTlBarBtnPicSample)
  End If
  '*** 16/05/05
End Sub
Private Sub ViewSampleBtn(ByVal po_Btn As MWCTL_BTN1)
  '*** 14/05/05
  Select Case UCase(po_Btn.Name)
  Case Is = UCase("CmdAdcBtnPicSample")
    'CmdAdcBtnPicSample.Pic = PicBoxCsADCBtnPic.Picture     '*** (Bef 23/05/05)
    If ADC("CsADCBtnHeight") > 0 And ADC("CsADCBtnWidth") > 0 Then
      CmdAdcBtnPicSample.Height = ADC("CsADCBtnHeight")
      CmdAdcBtnPicSample.Width = ADC("CsADCBtnWidth")
    End If
    CmdAdcBtnPicSample.ForeColor = Val(ImgCombo(mColl_ImgCombo("CsADCBtnForClr")).text)
    CmdAdcBtnPicSample.MaskColor = Val(ImgCombo(mColl_ImgCombo("CsADCBtnMaskClr")).text)
'    If ADC("CsADCBtnHeight") = 0 Or ADC("CsADCBtnWidth") = 0 Then
'      Call AdjustBtnSize(CmdAdcBtnPicSample)
'    End If
  Case Is = UCase("CmdBtnPicSet1Sample")
    'CmdBtnPicSet1Sample.Pic = PicBoxCsBtnPicSet1.Picture   '*** (Bef 23/05/05)
    If ADC("CsBtnHeightSet1") > 0 And ADC("CsBtnWidthSet1") > 0 Then
      CmdBtnPicSet1Sample.Height = ADC("CsBtnHeightSet1")
      CmdBtnPicSet1Sample.Width = ADC("CsBtnWidthSet1")
    End If
    CmdBtnPicSet1Sample.ForeColor = Val(ImgCombo(mColl_ImgCombo("CsBtnForClrSet1")).text)
    CmdBtnPicSet1Sample.MaskColor = Val(ImgCombo(mColl_ImgCombo("CsBtnMaskClrSet1")).text)
  Case Is = UCase("CmdBtnPicSet2Sample")
    'CmdBtnPicSet2Sample.Pic = PicBoxCsBtnPicSet2.Picture   '*** (Bef 23/05/05)
    If ADC("CsBtnHeightSet2") > 0 And ADC("CsBtnWidthSet2") > 0 Then
      CmdBtnPicSet2Sample.Height = ADC("CsBtnHeightSet2")
      CmdBtnPicSet2Sample.Width = ADC("CsBtnWidthSet2")
    End If
    CmdBtnPicSet2Sample.ForeColor = Val(ImgCombo(mColl_ImgCombo("CsBtnForClrSet2")).text)
    CmdBtnPicSet2Sample.MaskColor = Val(ImgCombo(mColl_ImgCombo("CsBtnMaskClrSet2")).text)
  Case Is = UCase("CmdTlBarBtnPicSample")
    'CmdTlBarBtnPicSample.Pic = PicBoxCsTlBarPic.Picture    '*** (Bef 23/05/05)
    If ADC("CsTlBarBtnHeight") > 0 And ADC("CsTlBarBtnWidth") > 0 Then
      CmdTlBarBtnPicSample.Height = ADC("CsTlBarBtnHeight")
      CmdTlBarBtnPicSample.Width = ADC("CsTlBarBtnWidth")
    End If
    CmdTlBarBtnPicSample.ForeColor = Val(ImgCombo(mColl_ImgCombo("CsTlBarBtnForClr")).text)
    CmdTlBarBtnPicSample.MaskColor = Val(ImgCombo(mColl_ImgCombo("CsTlBarBtnMaskClr")).text)
  End Select
  '*** 14/05/05
End Sub
Private Sub SetPicBoxProp(ByVal po_PicBox As PictureBox)
  '*** 23/05/05
  po_PicBox.Appearance = 1
  po_PicBox.Picture = LoadPicture("")
  
  If po_PicBox.BorderStyle = 1 Then
    po_PicBox.Top = po_PicBox.Top - 15
    po_PicBox.Left = po_PicBox.Left - 15
  End If
  '*** 23/05/05
End Sub

'''  ***** Std Code not to be changed **********************************
Private Sub Form_GotFocus()
  '*** Jenny Color (For Setting System Colors)
  SetMDISysColorsAfterLogin
  '*** Jenny Color (For Setting System Colors)
End Sub
Private Sub Form_LostFocus()
  '*** Jenny Color (For Setting System Colors)
  'SetMDISysColorsAfterLogin
  '*** Jenny Color (For Setting System Colors)
End Sub
Private Sub Form_Load()
  ms_Path = ""
  If InStrRev(App.Path, "\") > 1 Then ms_Path = Mid(App.Path, 1, InStrRev(App.Path, "\") - 1)
  If ms_Path <> "" Then ms_Path = ms_Path & "\MwClrSet\"
  
  '*** Jenny Colour
  'FRA_KEYS.Picture = Me.Picture
  'FRA_KEYS.BorderStyle = 0
  FraNKeyAll.Height = FRA_NKEYS.Height + 500
  FraNKeyAll.Width = FRA_NKEYS.Width + 500
  
  '(Bef 13/08/05)
  'SetMDISysColorsAfterLogin   '*** Jenny Color (For Setting System Colors)
  
  '*** 20/05/05
  Dim ws_BtnStr As String
  ws_BtnStr = "CmdCpyGo"
  Call SetObjColors(Me, ws_BtnStr)
  '*** 23/05/05
  If GetRecClrSet!Cs3D = "N" Then
    Call SetPicBoxProp(PicBoxCsFlxPic)
    Call SetPicBoxProp(PicBoxCsHlpPic)
    Call SetPicBoxProp(PicBoxCsFormPic)
    Call SetPicBoxProp(PicBoxCsMDIPic)
  End If
  '*** 23/05/05
  
  'Call BtnSet2(CmdCpyGo)
  
  '*** 20/05/05
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
    
    '*** (09/08/05)
    Set moCn = Nothing   '*** (Bef speed) bef 24/11/06
    Set mColl_ImgCombo = Nothing
    '*** (09/08/05)
    '*** Jenny Color (For Setting System Colors)
    'SetMDISysColorsAfterLogin
    '*** Jenny Color (For Setting System Colors)
End Sub
