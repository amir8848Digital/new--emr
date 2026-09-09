VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{6949E11D-3FEA-49A2-8D98-05EF092A8A24}#1.0#0"; "mwfctl.ocx"
Begin VB.Form xEmrFrmTrfIn 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "XML To XEmr Transfer"
   ClientHeight    =   9810
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12090
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9810
   ScaleWidth      =   12090
   Begin VB.PictureBox FraMst 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1230
      Left            =   0
      ScaleHeight     =   1170
      ScaleWidth      =   11820
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   7605
      Visible         =   0   'False
      Width           =   11885
      Begin VB.CommandButton CmdMstTrfIn 
         Caption         =   "&Transfer In"
         Height          =   315
         Left            =   10620
         TabIndex        =   15
         ToolTipText     =   "Transfer In"
         Top             =   435
         Width           =   1170
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   570
         Index           =   34
         Left            =   1110
         TabIndex        =   31
         ToolTipText     =   "Enter File Path"
         Top             =   450
         Width           =   9405
         _ExtentX        =   16589
         _ExtentY        =   1005
         Alignment       =   3
         Enabled         =   0   'False
         MaxLength       =   0
         IdName          =   "WFILEPATH"
      End
      Begin VB.Label LblMain 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "  Transfer In"
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
         TabIndex        =   50
         Top             =   -30
         Width           =   11885
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "File Path"
         BeginProperty Font 
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
         TabIndex        =   39
         Top             =   450
         Width           =   1065
      End
   End
   Begin VB.PictureBox FraInv 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1410
      Left            =   0
      ScaleHeight     =   1350
      ScaleWidth      =   11910
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   7425
      Visible         =   0   'False
      Width           =   11970
      Begin VB.CommandButton CmdInvTrfIn 
         Caption         =   "&Transfer In"
         Height          =   315
         Left            =   10590
         TabIndex        =   17
         ToolTipText     =   "Invoice Transfer In"
         Top             =   630
         Width           =   1230
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   19
         Left            =   1740
         TabIndex        =   16
         ToolTipText     =   "Enter Target Supplier Code"
         Top             =   330
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         MaxLength       =   8
         DataType        =   4
         IdName          =   "WINCMCDTGT"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   570
         Index           =   11
         Left            =   1740
         TabIndex        =   40
         ToolTipText     =   "Enter File Path"
         Top             =   615
         Width           =   8745
         _ExtentX        =   15425
         _ExtentY        =   1005
         Alignment       =   3
         Enabled         =   0   'False
         MaxLength       =   0
         IdName          =   "WINFILEPATH"
      End
      Begin VB.Label LblMain 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "  Invoice Transfer"
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
         TabIndex        =   51
         Top             =   -30
         Width           =   11970
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "File Path"
         BeginProperty Font 
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
         Left            =   90
         TabIndex        =   41
         Top             =   615
         Width           =   1065
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Target Supplier"
         BeginProperty Font 
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
         Left            =   90
         TabIndex        =   37
         Top             =   330
         Width           =   1725
      End
   End
   Begin VB.CommandButton CmdAllMst 
      Caption         =   "<&m> All Masters"
      CausesValidation=   0   'False
      Height          =   600
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Option To Transfer All Masters"
      Top             =   1200
      Width           =   2400
   End
   Begin VB.CommandButton CmdShowErrLog 
      Caption         =   "<&S>how Log"
      Height          =   405
      Left            =   9870
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   8925
      Width           =   1155
   End
   Begin VB.CommandButton CmdClearErrLog 
      Caption         =   "Clea<&R> Log"
      Height          =   405
      Left            =   9870
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   9330
      Width           =   1155
   End
   Begin VB.CommandButton CmdPrm 
      Caption         =   "<&f> Param"
      Height          =   600
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Option To Transfer Customer Master"
      Top             =   600
      Width           =   2400
   End
   Begin VB.CommandButton CmdVPrm 
      Caption         =   "<&g> Variable Param"
      Height          =   600
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Option To Transfer Customer Master"
      Top             =   600
      Width           =   2400
   End
   Begin VB.CommandButton CmdLoc 
      Caption         =   "<&j> Location"
      Height          =   600
      Left            =   9600
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Option To Transfer Customer Master"
      Top             =   600
      Width           =   2400
   End
   Begin VB.CommandButton CmdBkAg 
      Caption         =   "<&i> Bank/ Agent"
      Height          =   600
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Option To Transfer Customer Master"
      Top             =   600
      Width           =   2400
   End
   Begin VB.CommandButton CmdRmRtHist 
      Caption         =   "<&n> Rm Rate History"
      Enabled         =   0   'False
      Height          =   600
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Option To Transfer Customer Master"
      Top             =   1800
      Visible         =   0   'False
      Width           =   2400
   End
   Begin VB.CommandButton CmdLabRtHist 
      Caption         =   "<&o> Labour Rate History"
      Enabled         =   0   'False
      Height          =   600
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Option To Transfer Customer Master"
      Top             =   1800
      Visible         =   0   'False
      Width           =   2400
   End
   Begin VB.CommandButton CmdLabRt 
      Caption         =   "<&e> Labour Rate"
      Height          =   600
      Left            =   9600
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Option To Transfer Customer Master"
      Top             =   0
      Width           =   2400
   End
   Begin VB.CommandButton CmdRmRt 
      Caption         =   "<&d> Rm Rate"
      Height          =   600
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Option To Transfer Customer Master"
      Top             =   0
      Width           =   2400
   End
   Begin VB.CommandButton CmdDsg 
      Caption         =   "<&c> Design"
      Height          =   600
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Option To Transfer Customer Master"
      Top             =   0
      Width           =   2400
   End
   Begin VB.CommandButton CmdRm 
      Caption         =   "<&a> Raw Material"
      Height          =   600
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Option To Transfer Customer Master"
      Top             =   0
      Width           =   2400
   End
   Begin MSComDlg.CommonDialog ComDialogOpn 
      Left            =   10290
      Top             =   1230
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton CmdCust 
      Caption         =   "<&b> Customer"
      Height          =   600
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Option To Transfer Customer Master"
      Top             =   0
      Width           =   2400
   End
   Begin VB.CommandButton CmdInv 
      Caption         =   "<&l> Invoice"
      CausesValidation=   0   'False
      Height          =   600
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Option To Transfer Customer Master"
      Top             =   1200
      Width           =   2400
   End
   Begin VB.PictureBox FRA_KEYS 
      BorderStyle     =   0  'None
      Height          =   450
      Left            =   10950
      ScaleHeight     =   450
      ScaleWidth      =   825
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   1320
      Width           =   825
   End
   Begin VB.PictureBox FRA_NKEYS 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   420
      Left            =   10935
      ScaleHeight     =   420
      ScaleWidth      =   1125
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   1860
      Width           =   1125
      Begin VB.PictureBox FraNKeyAll 
         BorderStyle     =   0  'None
         Height          =   3705
         Left            =   6030
         ScaleHeight     =   3705
         ScaleWidth      =   5790
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   1530
         Width           =   5790
      End
   End
   Begin MwfCtl.MWCTL_HLP MWCTL_HLP1 
      Height          =   405
      Left            =   10230
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   2190
      Visible         =   0   'False
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   714
   End
   Begin MwfCtl.MWCTL_ADC ADC 
      CausesValidation=   0   'False
      Height          =   555
      Left            =   11085
      TabIndex        =   30
      Top             =   9300
      Width           =   2625
      _ExtentX        =   4630
      _ExtentY        =   979
      AllowAdd        =   0   'False
      AllowFind       =   0   'False
      AllowSave       =   0   'False
      AllowDelete     =   0   'False
      BtnWidth        =   885
      BtnHeight       =   435
      BtnMaskColor    =   8454143
      FontSz          =   8.25
   End
   Begin MwfCtl.MWCTL_MED ATXT 
      Height          =   810
      Index           =   8
      Left            =   0
      TabIndex        =   35
      ToolTipText     =   "Message"
      Top             =   8925
      Width           =   9750
      _ExtentX        =   17198
      _ExtentY        =   1429
      Alignment       =   3
      Enabled         =   0   'False
      MaxLength       =   255
      DataType        =   4
      IdName          =   "WDISP"
   End
   Begin VB.CommandButton CmdDtTbl 
      Caption         =   "<&h> Date Table"
      Height          =   600
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Option To Transfer Customer Master"
      Top             =   600
      Width           =   2400
   End
   Begin VB.CommandButton CmdOrd 
      Caption         =   "<&k> Order"
      Height          =   600
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Option To Transfer Customer Master"
      Top             =   1200
      Width           =   2400
   End
   Begin VB.PictureBox FraErrLog 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4365
      Left            =   0
      ScaleHeight     =   4305
      ScaleWidth      =   11910
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   4455
      Width           =   11970
      Begin MwfCtl.MWCTL_BTN1 CmdErrLogGo 
         Height          =   285
         Left            =   5925
         TabIndex        =   26
         ToolTipText     =   "Press This Button To Copy"
         Top             =   3690
         Width           =   585
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483633
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
      Begin MwfCtl.MWCTL_FLX GrdErrLog 
         Height          =   3330
         Left            =   90
         TabIndex        =   21
         Top             =   300
         Width           =   11745
         _ExtentX        =   20717
         _ExtentY        =   5874
         Cols            =   5
         AllowAdd        =   0   'False
         AllowDelete     =   0   'False
         colname1        =   "WSRNO"
         heading1        =   "Sr."
         datatype1       =   1
         mask1           =   "###0"
         maxlength1      =   4
         colname2        =   "WELTABLE"
         heading2        =   "Table                     "
         maxlength2      =   30
         colname3        =   "WELKEY"
         heading3        =   "Key                                              "
         maxlength3      =   30
         colname4        =   "WELERRMSG"
         heading4        =   $"xEmrFrmTrfIn.frx":0000
         datatype4       =   4
         maxlength4      =   255
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   0
         Left            =   1245
         TabIndex        =   22
         ToolTipText     =   "Enter From Table"
         Top             =   3690
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   503
         MaxLength       =   0
         IdName          =   "WTBLFR"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   1
         Left            =   1245
         TabIndex        =   23
         ToolTipText     =   "Enter To Table "
         Top             =   3975
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   503
         MaxLength       =   0
         IdName          =   "WTBLTO"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   2
         Left            =   4635
         TabIndex        =   24
         ToolTipText     =   "Enter From Day And Month"
         Top             =   3690
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
         MaxLength       =   0
         IdName          =   "WDDMMFR"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   3
         Left            =   4635
         TabIndex        =   25
         ToolTipText     =   "Enter To Day And Month"
         Top             =   3975
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
         MaxLength       =   0
         IdName          =   "WDDMMTO"
      End
      Begin VB.Label LblMain 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "  Error Log"
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
         TabIndex        =   52
         Top             =   -30
         Width           =   12000
      End
      Begin VB.Label ALBL 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Table Fr"
         BeginProperty Font 
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
         TabIndex        =   46
         Top             =   3690
         Width           =   1095
      End
      Begin VB.Label ALBL 
         Alignment       =   1  'Right Justify
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
         Index           =   0
         Left            =   90
         TabIndex        =   45
         Top             =   3975
         Width           =   1095
      End
      Begin VB.Label ALBL 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "DdMm Fr"
         BeginProperty Font 
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
         Left            =   3480
         TabIndex        =   44
         Top             =   3690
         Width           =   1095
      End
      Begin VB.Label ALBL 
         Alignment       =   1  'Right Justify
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
         Left            =   3480
         TabIndex        =   43
         Top             =   3975
         Width           =   1095
      End
   End
   Begin VB.PictureBox FraSelFolder 
      Height          =   6795
      Left            =   0
      ScaleHeight     =   6735
      ScaleWidth      =   9675
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   2040
      Width           =   9735
      Begin VB.DriveListBox DriveLst 
         Height          =   315
         Left            =   150
         TabIndex        =   49
         Top             =   5745
         Width           =   9465
      End
      Begin ComctlLib.ImageList ImgList 
         Left            =   7230
         Top             =   6240
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   3
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "xEmrFrmTrfIn.frx":0094
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "xEmrFrmTrfIn.frx":04E6
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "xEmrFrmTrfIn.frx":0938
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.TextBox TxtFolderPath 
         Height          =   285
         Left            =   1470
         TabIndex        =   28
         Top             =   6270
         Width           =   6885
      End
      Begin VB.CommandButton CmdFolderOk 
         Caption         =   "&Transfer In"
         Height          =   315
         Left            =   8430
         TabIndex        =   29
         Top             =   6270
         Width           =   1185
      End
      Begin ComctlLib.TreeView TreeViewSel 
         Height          =   5415
         Left            =   150
         TabIndex        =   47
         Top             =   330
         Width           =   9465
         _ExtentX        =   16695
         _ExtentY        =   9551
         _Version        =   393217
         Indentation     =   635
         LabelEdit       =   1
         Style           =   7
         Appearance      =   1
      End
      Begin VB.Label LblMain 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "  Select Folder"
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
         TabIndex        =   53
         Top             =   -30
         Width           =   9735
      End
      Begin VB.Label ALBL 
         Caption         =   "Folder Path"
         BeginProperty Font 
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
         Left            =   150
         TabIndex        =   48
         Top             =   6270
         Width           =   1305
      End
   End
   Begin VB.Line Line1 
      Visible         =   0   'False
      X1              =   5940
      X2              =   5940
      Y1              =   1860
      Y2              =   2640
   End
End
Attribute VB_Name = "xEmrFrmTrfIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************************
'*** Data Porting From XML To xEmr
'*** Also from XML To Emr
'*****************************************************************************************

Option Explicit

Dim WithEvents Explorer As CExplorer2
Attribute Explorer.VB_VarHelpID = -1

Dim moCn As MwfLib.MDOConnection

'*** Declarations For The Target RecordSets (Add)
Dim Rs_TgtAddX As MwfLib.MDORowSet
Dim Rs_TgtAddA As MwfLib.MDORowSet
Dim Rs_TgtAddB As MwfLib.MDORowSet
Dim Rs_TgtAddC As MwfLib.MDORowSet
Dim Rs_TgtAddB1 As MwfLib.MDORowSet
Dim Rs_TgtAddC1 As MwfLib.MDORowSet
Dim Rs_TgtAddD As MwfLib.MDORowSet
Dim Rs_TgtAddE As MwfLib.MDORowSet
Dim Rs_TgtAddG As MwfLib.MDORowSet
' Zubin 213
Dim Rs_TgtAddH As MwfLib.MDORowSet

'*** Declarations For The Target RecordSets (Update)
Dim Rs_TgtUpdX As MwfLib.MDORowSet
Dim Rs_TgtUpdA As MwfLib.MDORowSet
Dim Rs_TgtUpdB As MwfLib.MDORowSet
Dim Rs_TgtUpdC As MwfLib.MDORowSet
Dim Rs_TgtUpdB1 As MwfLib.MDORowSet
Dim Rs_TgtUpdC1 As MwfLib.MDORowSet
Dim Rs_TgtUpdD As MwfLib.MDORowSet
Dim Rs_TgtUpdE As MwfLib.MDORowSet
Dim Rs_TgtUpdG As MwfLib.MDORowSet
' Zubin 213
Dim Rs_TgtUpdH As MwfLib.MDORowSet

'*** Declarations For The XML Document
Dim XmlDoc As MSXML2.DOMDocument
Dim xmlRootElement As MSXML2.IXMLDOMElement
Dim xmlChildX As MSXML2.IXMLDOMElement
Dim xmlChildLevel1 As MSXML2.IXMLDOMElement
Dim xmlChildLevel2 As MSXML2.IXMLDOMElement

'*** Arr_Cv stores the causes validation state, Arr_Ed the Ena/Disa state of Cmd Buttons
Dim mArr_Cv As Variant, mArr_Ed As Variant

Dim mf_i As Single, mf_j As Single, mf_k As Single
Dim mi_i As Integer
Dim ms_FilePth As String
Dim ms_Opt As String
Dim mf_Cancel As Boolean, ms_ErrMsg As String
Dim mf_SrcCurCnvRt As Single

Dim ms_ErrTbl As String
Dim xmlErrRecSet As MSXML2.IXMLDOMElement


Enum en_EmrTrfInDispFra
  Rm = 0
  Cust = 1
  Dsg = 2
  RmRt = 3
  LabRt = 4
  RmRtHist = 5
  LabRtHist = 6
  DtTbl = 7
  BkAg = 8
  Loc = 9
  Prm = 10
  VPrm = 11
  Ord = 12
  Inv = 13
  ErrLog = 14
  AllMst = 15
End Enum
Private Sub ADC_Load()
  ReDim mArr_Cv(16), mArr_Ed(16)
  Set moCn = ADC.Connection
  Set ADC.FirKeyCtl = ADC
  GrdErrLog.HotKey = "A"
  Call HideAllFras
  
  If UCase(ctProjectName) = "EMR" Then
    CmdOrd.Enabled = False
    CmdOrd.Visible = False
    CmdInv.Enabled = False
    CmdInv.Visible = False
  End If
End Sub
Private Sub ATXT_FldWhen(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String)
  Select Case UCase(IdName)
  '*** Invoice ***
  Case Is = UCase("wInCmCdTgt")
    Call HlpList.CustCd("S")
  '*** Invoice ***
  Case Is = UCase("wTblFr"), UCase("wTblTo")
    Call HlpList.Tables
  End Select
End Sub
Private Sub ATXT_FldChange(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String, ByVal pv_NewValue As Variant)
  Select Case UCase(IdName)
  Case Is = UCase("wInCmCdTgt")
    Cancel = (pv_NewValue <> "") And Not moCn.RecSeek("Select CmCd from CustMst where CmCtg= 'S' and CmCd = '" + pv_NewValue + "' ")
    If Cancel = True Then ErrMsg = "Invalid Target Supplier Code": Exit Sub
    Case Is = UCase("wDdMmFr"), UCase("wDdMmTo")
      Cancel = (pv_NewValue <> "" And _
                  (Not IsNumeric(pv_NewValue) _
                    Or Len(pv_NewValue) <> 4 _
                    Or (Mid(pv_NewValue, 1, 2) < 1 Or Mid(pv_NewValue, 1, 2) > 31) _
                    Or (Mid(pv_NewValue, 3, 2) < 1 Or Mid(pv_NewValue, 3, 2) > 12)))
      ErrMsg = "Invalid DdMm Name": Exit Sub
  End Select
End Sub
Private Sub CmdBkAg_Click()
  Call DispFra(BkAg)
  If FraMst.Visible = True Then
    ' After Clr Scrn
    'FraMst.Caption = "Bank/Agent Transfer In"
    LblMain(0).Caption = "Bank/Agent Transfer In"
    ' After Clr Scrn
    ms_FilePth = OpenDocFilePath("Bank/ Agent")
    ADC("wFilePath") = ms_FilePth
    ms_Opt = "BkAg"
  End If
End Sub
Private Sub CmdRm_Click()
  Call DispFra(Rm)
  If FraMst.Visible = True Then
    ' After Clr Scrn
    'FraMst.Caption = "Raw Material Transfer In"
    LblMain(0).Caption = "Raw Material Transfer In"
    ' After Clr Scrn
    
    ms_FilePth = OpenDocFilePath("Raw Material")
    ADC("wFilePath") = ms_FilePth
    ms_Opt = "Rm"
  End If
End Sub
Private Sub CmdCust_Click()
  Call DispFra(Cust)
  If FraMst.Visible = True Then
    ' After Clr Scrn
    'FraMst.Caption = "Customer Transfer In"
    LblMain(0).Caption = "Customer Transfer In"
    ' After Clr Scrn
    ms_FilePth = OpenDocFilePath("Customer Master")
    ADC("wFilePath") = ms_FilePth
    ms_Opt = "Cust"
  End If
End Sub
Private Sub CmdDsg_Click()
  Call DispFra(Dsg)
  If FraMst.Visible = True Then
    ' After Clr Scrn
    'FraMst.Caption = "Design Transfer In"
    LblMain(0).Caption = "Design Transfer In"
    ' After Clr Scrn
    ms_FilePth = OpenDocFilePath("Design Master")
    ADC("wFilePath") = ms_FilePth
    ms_Opt = "Dsg"
  End If
End Sub
Private Sub CmdLabRt_Click()
  Call DispFra(LabRt)
  If FraMst.Visible = True Then
    ' After Clr Scrn
    'FraMst.Caption = "Labour Rate Transfer In"
    LblMain(0).Caption = "Labour Rate Transfer In"
    ' After Clr Scrn
    ms_FilePth = OpenDocFilePath("Labour Rate")
    ADC("wFilePath") = ms_FilePth
    ms_Opt = "LabRt"
  End If
End Sub
Private Sub CmdRmRt_Click()
  Call DispFra(RmRt)
  If FraMst.Visible = True Then
    ' After Clr Scrn
    'FraMst.Caption = "Rm Rate Transfer In"
    LblMain(0).Caption = "Rm Rate Transfer In"
    ' After Clr Scrn
    ms_FilePth = OpenDocFilePath("Rm Rate")
    ADC("wFilePath") = ms_FilePth
    ms_Opt = "RmRt"
  End If
End Sub
Private Sub CmdLabRtHist_Click()
  Call DispFra(LabRtHist)
  If FraMst.Visible = True Then
    ' After Clr Scrn
    'FraMst.Caption = "Labour Rate History Transfer In"
    LblMain(0).Caption = "Labour Rate History Transfer In"
    ' After Clr Scrn
    ms_FilePth = OpenDocFilePath("Labour Rate History")
    ADC("wFilePath") = ms_FilePth
    ms_Opt = "LabRtHist"
  End If
End Sub
Private Sub CmdRmRtHist_Click()
  Call DispFra(RmRtHist)
  If FraMst.Visible = True Then
    ' After Clr Scrn
    'FraMst.Caption = "Rm Rate History Transfer In"
    LblMain(0).Caption = "Rm Rate History Transfer In"
    ' After Clr Scrn
    ms_FilePth = OpenDocFilePath("Rm Rate History")
    ADC("wFilePath") = ms_FilePth
    ms_Opt = "RmRtHist"
  End If
End Sub
Private Sub CmdDtTbl_Click()
  Call DispFra(DtTbl)
  If FraMst.Visible = True Then
    ' After Clr Scrn
    'FraMst.Caption = "Date Table Transfer In"
    LblMain(0).Caption = "Date Table Transfer In"
    ' After Clr Scrn
    ms_FilePth = OpenDocFilePath("Date Table")
    ADC("wFilePath") = ms_FilePth
    ms_Opt = "DtTbl"
  End If
End Sub
Private Sub CmdLoc_Click()
  Call DispFra(Loc)
  If FraMst.Visible = True Then
    ' After Clr Scrn
    'FraMst.Caption = "Location Transfer In"
    LblMain(0).Caption = "Location Transfer In"
    ' After Clr Scrn
    ms_FilePth = OpenDocFilePath("Location")
    ADC("wFilePath") = ms_FilePth
    ms_Opt = "Loc"
  End If
End Sub
Private Sub CmdPrm_Click()
  Call DispFra(Prm)
  If FraMst.Visible = True Then
    ' After Clr Scrn
    'FraMst.Caption = "Param Transfer In"
    LblMain(0).Caption = "Param Transfer In"
    ' After Clr Scrn
    ms_FilePth = OpenDocFilePath("Param")
    ADC("wFilePath") = ms_FilePth
    ms_Opt = "Prm"
  End If
End Sub
Private Sub CmdVPrm_Click()
  Call DispFra(VPrm)
  If FraMst.Visible = True Then
    ' After Clr Scrn
    'FraMst.Caption = "Variable Param Transfer In"
    LblMain(0).Caption = "Variable Param Transfer In"
    ' After Clr Scrn
    ms_FilePth = OpenDocFilePath("Variable Param")
    ADC("wFilePath") = ms_FilePth
    ms_Opt = "VPrm"
  End If
End Sub
Private Sub CmdOrd_Click()
  Call DispFra(Ord)
  If FraMst.Visible = True Then
    ' After Clr Scrn
    'FraMst.Caption = "Order Transfer In"
    LblMain(0).Caption = "Order Transfer In"
    ' After Clr Scrn
    ms_FilePth = OpenDocFilePath("Order")
    ADC("wFilePath") = ms_FilePth
    ms_Opt = "Ord"
  End If
End Sub
Private Sub CmdAllMst_Click()
  Call DispFra(AllMst)
  If FraSelFolder.Visible = True Then
    'FraSelFolder.Caption = "Masters Transfer In"
    
    If TxtFolderPath.text = "" Then TxtFolderPath.text = moCn.GetFldVal("Select PDesc225 from Param where PTyp= 'TRFDIR' and PMCd= 'TRFDIR' ")
    
    TreeViewSel.ImageList = ImgList
    RefreshTree
  End If
End Sub
Sub RefreshTree()
    Set Explorer = New CExplorer2
    Explorer.Init TreeViewSel, , 1, 2, 3
    
    On Error Resume Next
    
    '** add the "C:\" root
    Dim rootDir As New CDirectory2
    'rootDir.Path = "C:"
    'rootDir.Path = DriveLst.List(DriveLst.ListIndex)
    rootDir.Path = Mid(DriveLst.Drive, 1, InStr(1, DriveLst.Drive, ":"))
    Set Explorer.Root = rootDir
End Sub
Private Sub DriveLst_Change()
  RefreshTree
End Sub
Private Sub TreeViewSel_NodeClick(ByVal Node As ComctlLib.Node)
  TxtFolderPath.text = TreeViewSel.SelectedItem.FullPath
End Sub
Private Sub CmdFolderOk_Click()
  '***%%%
  'Dim fs, f, f1, fc, s
  'Set fs = CreateObject("Scripting.FileSystemObject")
  'Set f = fs.GetFolder(folderspec)
  'Set fc = f.SubFolders
  'For Each f1 In fc
  '    s = s & f1.Name
  '    s = s & vbCrLf
  'Next
  'MsgBox s
  '***%%%
  
  
  
  Dim wo_FileSysObj, wo_Folder, wo_File, wo_FileColl
  Dim ws_FolderPath As String
  Dim wo_XmlDoc As MSXML2.DOMDocument
  Dim wo_XmlRootElement As MSXML2.IXMLDOMElement
  
  On Error Resume Next
  
  ws_FolderPath = TxtFolderPath.text    '"C:\Temp\AllMst"
  If Trim(ws_FolderPath) = "" Then
    ADC("wDisp") = "Enter Proper Folder Name"
    DispMsg "Enter Proper Folder Name", etError: Exit Sub
  End If
  '*** Or Dir(ComDialogOpn.FileName) = ""  (This does not work for directories)
  Set wo_FileSysObj = CreateObject("Scripting.FileSystemObject")
  Set wo_Folder = wo_FileSysObj.GetFolder(ws_FolderPath)
  
  '*** If Incorrect Path then Exit Sub.
  If Err <> 0 Then
    ADC("wDisp") = "Folder Does Not Exist"
    DispMsg "Folder Does Not Exist", etError: Exit Sub
  End If
  On Error GoTo 0
  
  If UCase(wo_Folder.Type) <> UCase("File Folder") Then
    ADC("wDisp") = "Specified File Is Not A Folder"
    DispMsg "Specified File Is Not A Folder", etError: Exit Sub
  End If

  Set wo_FileColl = wo_Folder.Files
  For Each wo_File In wo_FileColl
    If UCase(wo_File.Type) = UCase("XML Document") Then
      ms_FilePth = ws_FolderPath + "\" + wo_File.Name
      ADC("wFilePath") = ms_FilePth
      Set wo_XmlDoc = New MSXML2.DOMDocument
      wo_XmlDoc.Load ms_FilePth
      Set wo_XmlRootElement = wo_XmlDoc.childNodes(0)
      Select Case UCase(wo_XmlRootElement.baseName)
      Case Is = UCase("EmrBnkAgMstTransfer")
        ms_Opt = "BkAg"
      Case Is = UCase("EmrDtTblTransfer")
        ms_Opt = "DtTbl"
      Case Is = UCase("EmrLocTransfer")
        ms_Opt = "Loc"
      Case Is = UCase("EmrParamTransfer")
        ms_Opt = "Prm"
      Case Is = UCase("EmrvParamTransfer")
        ms_Opt = "VPrm"
      Case Is = UCase("EmrRmMstTransfer")
        ms_Opt = "Rm"
      Case Is = UCase("EmrCustMstTransfer")
        ms_Opt = "Cust"
      Case Is = UCase("EmrDsgMstTransfer")
        ms_Opt = "Dsg"
      Case Else
        ms_Opt = ""
      End Select
      
      If ms_Opt <> "" Then Call CmdMstTrfIn_Click
    End If
  Next

  Call CmdAllMst_Click
  ADC("wDisp") = "Transfer-In Of Masters Completed Successfully. "
  DispMsg "Transfer-In Of Masters Completed Successfully. ", etInfo
End Sub
Private Sub CmdMstTrfIn_Click()
  If ms_FilePth = "" Then DispMsg "Invalid File Name/ File Name Cannot Be Blank", etError: Exit Sub
  DispMsg "", etInfo
  Select Case UCase(ms_Opt)
  Case Is = UCase("BkAg")
    Call TransferInRoutine("BnkAgMst")
    If FraMst.Visible = True Then Call CmdBkAg_Click
  Case Is = UCase("Rm")
    Call TransferInRoutine("RmMst", "RmAlyCmp")
    If FraMst.Visible = True Then Call CmdRm_Click
  Case Is = UCase("Cust")
    Call TransferInRoutine("CustMst", "CustDsg", , , , , "CustAly", , "CustRm")
    If FraMst.Visible = True Then Call CmdCust_Click
  Case Is = UCase("Dsg")
    Call TransferInRoutine("DsgMst", "DsgRm", , , , , "DsgLab", "DsgPrt", "DsgAna", , "DsgModAvl")
    If FraMst.Visible = True Then Call CmdDsg_Click
  Case Is = UCase("LabRt")
    Call TransferInRoutine("LabRt")
    If FraMst.Visible = True Then Call CmdLabRt_Click
  Case Is = UCase("RmRt")
    Call TransferInRoutine("RmRt")
    If FraMst.Visible = True Then Call CmdRmRt_Click
  Case Is = UCase("LabRtHist")
    Call TransferInRoutine("LabRtHist")
    If FraMst.Visible = True Then Call CmdLabRtHist_Click
  Case Is = UCase("RmRtHist")
    Call TransferInRoutine("RmRtHist")
    If FraMst.Visible = True Then Call CmdRmRtHist_Click
  Case Is = UCase("DtTbl")
    Call TransferInRoutine("DtTbl")
    If FraMst.Visible = True Then Call CmdDtTbl_Click
  Case Is = UCase("Loc")
    Call TransferInRoutine("Loc", , , , , , , , , "Loc")
    If FraMst.Visible = True Then Call CmdLoc_Click
  Case Is = UCase("Prm")
    Call TransferInRoutine("Param")
    If FraMst.Visible = True Then Call CmdPrm_Click
  Case Is = UCase("VPrm")
    Call TransferInRoutine("vParam", , , , , , , , , "VPrm")
    If FraMst.Visible = True Then Call CmdVPrm_Click
  Case Is = UCase("Ord")
    Call TransferInRoutine("xTxn", "xTxnDsg", "xTxnRm", "xTxnLab", "xTxnDsgAmd", "xTxnTgt", "xTxnTgt", , , "SORD")
    If FraMst.Visible = True Then Call CmdOrd_Click
  End Select
End Sub
Private Sub CmdInv_Click()
  Call DispFra(Inv)
  If FraInv.Visible = True Then
    ' Clr Scrn
    'FraInv.Caption = "Invoice Transfer In"
    LblMain(1).Caption = "Invoice Transfer In"
    ' Clr Scrn
    ms_FilePth = OpenDocFilePath("Invoice")
    ADC("wInFilePath") = ms_FilePth
    CmdInv.SetFocus
    ms_Opt = "Inv"
  End If
End Sub
Private Sub CmdInvTrfIn_Click()
  If ADC("wInCmCdTgt") = "" Then DispMsg "Supplier Code Cannot Be Blank", etError: Exit Sub
  
  If ms_FilePth = "" Then DispMsg "Invalid File Name/ File Name Cannot Be Blank", etError: Exit Sub
  DispMsg "", etInfo
  
  Call TransferInRoutine("xTxn", "xTxnDsg", "xTxnRm", "xTxnLab", "xTxnDsgAmd", "xTxnTgt", "xTxnTgt", , , "PUR")
  Call CmdInv_Click
End Sub
Private Sub CmdClearErrLog_Click()
  If MsgBox("Are you sure you want to clear Transfer-In Log?", vbYesNo + vbDefaultButton2, "XML To XEmr Transfer") = vbYes Then
    moCn.Execute ("Delete from ErrLog")
    GrdErrLog.Rows = 1
  End If
End Sub
Private Sub CmdShowErrLog_Click()
  Call DispFra(ErrLog)
End Sub
Private Sub CmdErrLogGo_Click()
  Call ShowErrLog
End Sub
Private Sub GrdErrLog_ColWhen(ByVal RowNum As Integer, ByVal ColName As String, Cancel As Boolean, ErrMsg As String, ByVal Action As MwfCtl.en_fgAction)
  Cancel = True
End Sub
Private Sub TransferInRoutine(ps_TblX As String, Optional ByVal ps_TblA As Variant, _
                              Optional ByVal ps_TblB As Variant, _
                              Optional ByVal ps_TblC As Variant, _
                              Optional ByVal ps_TblB1 As Variant, _
                              Optional ByVal ps_TblC1 As Variant, _
                              Optional ByVal ps_TblD As Variant, _
                              Optional ByVal ps_TblE As Variant, _
                              Optional ByVal ps_TblG As Variant, _
                              Optional ps_VchTyp As Variant, _
                              Optional ByVal ps_TblH As Variant)
  
  On Error GoTo errh
  
  ms_ErrTbl = ""
  Set xmlErrRecSet = Nothing
  
  Dim ws_TblA As String, ws_TblB As String, ws_TblC As String, ws_TblB1 As String
  Dim ws_TblC1 As String, ws_TblD As String, ws_TblE As String, ws_TblG As String
  Dim ws_TcTyp As String, ws_Cnd As String, wb_Flag As Boolean
  Dim wb_SourceAvlbl As Boolean, ws_KeyVal As String, ws_KeyStr As String
  Dim wf_xTdTotAmt As Single, ws_VchTyp As String
  
  ' Zubin 213
  Dim ws_TblH As String
  
  'Zubin 05-01-04
  '*** Keeps count of Record Count successfully transferred in and the total no. of records
  '    to be transfered
  Dim wi_Cnt As Integer, wi_TotCnt As Integer
  Dim ws_TrnfCntStr As String
  'Zubin 05-01-04
  
  mf_SrcCurCnvRt = 0 '*** For xTxn
  wb_SourceAvlbl = False: ws_KeyVal = "": ws_KeyStr = ""

  Call VerMatch
  
  Set XmlDoc = New MSXML2.DOMDocument
  '*** Load The XML Document InTo Our Document Object
  XmlDoc.Load ms_FilePth
  ADC("wDisp") = ps_TblX + " Transfer-In Is In Process ..."
  '*** Accessing THe RootElement Which Is A Child Of The Document Object
  Set xmlRootElement = XmlDoc.childNodes(0)
    
  '*** (As On 16/11/04)
  Dim ws_RootEle As String
  If Not IsMissing(ps_VchTyp) Then
    Select Case UCase(ps_VchTyp)
    Case Is = "SORD"
      ws_RootEle = "EmrOrdMstTransfer"
    Case Is = "PUR"
      ws_RootEle = "EmrInvHdTransfer"
    Case Is = "LOC"
      ws_RootEle = "Emr" + ps_TblX + "Transfer"
    Case Is = "VPRM"
      ws_RootEle = "Emr" + ps_TblX + "Transfer"
    End Select
    ws_VchTyp = ps_VchTyp
  Else
    ws_RootEle = "Emr" + ps_TblX + "Transfer"
    ws_VchTyp = ""
  End If
  If UCase(xmlRootElement.baseName) <> UCase(ws_RootEle) Then DispMsg "File Does Not Contain Proper Data", etError: ADC("wDisp") = "File Does Not Contain Proper Data": Exit Sub
  'If Len(xmlRootElement.baseName) < Len(ws_RootEle) Then DispMsg "File Does Not Contain Proper Data", etError: ADC("wDisp") = "File Does Not Contain Proper Data": Exit Sub
  'If Mid(UCase(xmlRootElement.baseName), 1, Len(ws_RootEle)) <> UCase(ws_RootEle) Then DispMsg "File Does Not Contain Proper Data", etError: ADC("wDisp") = "File Does Not Contain Proper Data": Exit Sub
  '*** (As On 16/11/04)
  
  
  '*** (Bef 16/11/04)
  'If Not IsMissing(ps_VchTyp) Then
  '  Select Case UCase(ps_VchTyp)
  '  Case Is = "SORD"
  '    If UCase(xmlRootElement.baseName) <> UCase("EmrOrdMstTransfer") Then DispMsg "File Does Not Contain Proper Data", etError: ADC("wDisp") = "File Does Not Contain Proper Data": Exit Sub
  '  Case Is = "PUR"
  '    If UCase(xmlRootElement.baseName) <> UCase("EmrInvHdTransfer") Then DispMsg "File Does Not Contain Proper Data", etError: ADC("wDisp") = "File Does Not Contain Proper Data": Exit Sub
  '  Case Is = "LOC"
  '    If UCase(xmlRootElement.baseName) <> UCase("Emr" + ps_TblX + "Transfer") Then DispMsg "File Does Not Contain Proper Data", etError: ADC("wDisp") = "File Does Not Contain Proper Data": Exit Sub
  '  Case Is = "VPRM"
  '    If UCase(xmlRootElement.baseName) <> UCase("Emr" + ps_TblX + "Transfer") Then DispMsg "File Does Not Contain Proper Data", etError: ADC("wDisp") = "File Does Not Contain Proper Data": Exit Sub
  '  End Select
  '  ws_VchTyp = ps_VchTyp
  'Else
  '  If UCase(xmlRootElement.baseName) <> UCase("Emr" + ps_TblX + "Transfer") Then DispMsg "File Does Not Contain Proper Data", etError: ADC("wDisp") = "File Does Not Contain Proper Data": Exit Sub
  '  ws_VchTyp = ""
  'End If
  '*** (Bef 16/11/04)
  
  If xmlRootElement.Attributes.length - 1 < 0 Then DispMsg "XML File Does Not Have A Proper Version", etError: ADC("wDisp") = "XML File Does Not Have A Proper Version": Exit Sub
  If UCase(xmlRootElement.Attributes.Item(0).nodeName) <> UCase("Ver") Then DispMsg "XML File Does Not Have A Proper Version", etError: ADC("wDisp") = "XML File Does Not Have A Proper Version": Exit Sub
  If Trim(xmlRootElement.Attributes.Item(0).nodeValue) <> gs_Ver Then DispMsg "XML Version Not Proper", etError: ADC("wDisp") = "XML Version Not Proper": Exit Sub
  
  '*** Opening Targer RecordSet And XML Doc(RecSource) For Transfer To Table
  Set Rs_TgtAddX = moCn.OpenResultset("Select * from " + ps_TblX + " where 1= 2")
  If Not IsMissing(ps_TblA) Then Set Rs_TgtAddA = moCn.OpenResultset("Select * from " + ps_TblA + " where 1= 2"): ws_TblA = ps_TblA
  If Not IsMissing(ps_TblB) Then Set Rs_TgtAddB = moCn.OpenResultset("Select * from " + ps_TblB + " where 1= 2"): ws_TblB = ps_TblB
  If Not IsMissing(ps_TblC) Then Set Rs_TgtAddC = moCn.OpenResultset("Select * from " + ps_TblC + " where 1= 2"): ws_TblC = ps_TblC
  If Not IsMissing(ps_TblB1) Then Set Rs_TgtAddB1 = moCn.OpenResultset("Select * from " + ps_TblB1 + " where 1= 2"): ws_TblB1 = ps_TblB1
  If Not IsMissing(ps_TblC1) Then Set Rs_TgtAddC1 = moCn.OpenResultset("Select * from " + ps_TblC1 + " where 1= 2"): ws_TblC1 = ps_TblC1
  If Not IsMissing(ps_TblD) Then Set Rs_TgtAddD = moCn.OpenResultset("Select * from " + ps_TblD + " where 1= 2"): ws_TblD = ps_TblD
  If Not IsMissing(ps_TblE) Then Set Rs_TgtAddE = moCn.OpenResultset("Select * from " + ps_TblE + " where 1= 2"): ws_TblE = ps_TblE
  If Not IsMissing(ps_TblG) Then Set Rs_TgtAddG = moCn.OpenResultset("Select * from " + ps_TblG + " where 1= 2"): ws_TblG = ps_TblG
  
  If Not IsMissing(ps_TblH) Then Set Rs_TgtAddH = moCn.OpenResultset("Select * from " + ps_TblH + " where 1= 2"): ws_TblH = ps_TblH
  
  'Zubin 05/01/04
  '*** Total Records to be transferred
  wi_TotCnt = xmlRootElement.childNodes.length
  'Zubin 05/01/04
  
  '*** Loop For Retrieving X Recs - Child Elements Of RootElement
  For mf_i = 0 To xmlRootElement.childNodes.length - 1
    DoEvents
    moCn.BeginTrans (mdoCommitRec)
    '*** Take A Reference Of The Current X Rec In The Document
    'Dim xmlChildX As IXMLDOMElement
    Set xmlChildX = xmlRootElement.childNodes(mf_i)
    
    '*** 25/12/03
    ms_ErrTbl = ps_TblX
    Set xmlErrRecSet = xmlChildX
    '*** 25/12/03
    
    If CheckForRecExists(ps_TblX, xmlChildX, ws_TcTyp, ws_Cnd, ws_KeyVal) = True Then
      If ws_VchTyp = "PUR" Or ws_VchTyp = "SORD" Then moCn.CommitTrans mdoCommitRec: GoTo XNext
      ADC("wDisp") = ps_TblX + ": " + ws_KeyVal
      Set Rs_TgtUpdX = moCn.OpenResultset("Select * from " + ps_TblX + " " + ws_Cnd)
      If Rs_TgtUpdX.RecCount > 0 Then
        wb_Flag = True
        Rs_TgtUpdX.edit
        MoveAttributesToFields Rs_TgtUpdX, xmlChildX
        If Not IsMissing(ps_VchTyp) Then Call GenFldVal(ws_TcTyp, ps_TblX, Rs_TgtUpdX)
        Rs_TgtUpdX.Update
      End If
      '''Set Rs_TgtUpdX = Nothing
    Else
      ADC("wDisp") = ps_TblX + ": " + ws_KeyVal
      wb_Flag = False
      '*** Add A New Record In The Target X RecordSet
      Rs_TgtAddX.Addnew
      '*** Read Attributes Of Current X Element And Move It To Fields Of RecordSet
      MoveAttributesToFields Rs_TgtAddX, xmlChildX
      If Not IsMissing(ps_VchTyp) Then Call GenFldVal(ws_TcTyp, ps_TblX, Rs_TgtAddX)
      Rs_TgtAddX.Update
    End If
    '*** Update The X Record To The Database
    
    '*** Loop For A_Level Records - ChildNodes Of Current X Element
    For mf_j = 0 To xmlChildX.childNodes.length - 1
      '*** Take A Reference Of The Current A Rec In The Document
      Set xmlChildLevel1 = xmlChildX.childNodes(mf_j)
      
      If UCase(xmlChildLevel1.baseName) = UCase(ws_TblA) Then
        '*** 25/12/03
        ms_ErrTbl = ws_TblA
        Set xmlErrRecSet = xmlChildLevel1
        '*** 25/12/03
        
        If wb_Flag = True Then
          If CheckForRecExists(ws_TblA, xmlChildLevel1, ws_TcTyp, ws_Cnd) = True Then
            Set Rs_TgtUpdA = moCn.OpenResultset("Select * from " + ws_TblA + " " + ws_Cnd)
            If Rs_TgtUpdA.RecCount > 0 Then
              Rs_TgtUpdA.edit
              MoveAttributesToFields Rs_TgtUpdA, xmlChildLevel1
              If Not IsMissing(ps_VchTyp) Then Call GenFldVal(ws_TcTyp, ws_TblA, Rs_TgtUpdA)
              Rs_TgtUpdA.Update
            End If
            '''Set Rs_TgtUpdA = Nothing
          Else
            Rs_TgtAddA.Addnew
            MoveAttributesToFields Rs_TgtAddA, xmlChildLevel1
            If Not IsMissing(ps_VchTyp) Then Call GenFldVal(ws_TcTyp, ws_TblA, Rs_TgtAddA)
            Rs_TgtAddA.Update
          End If
        Else
          Rs_TgtAddA.Addnew
          MoveAttributesToFields Rs_TgtAddA, xmlChildLevel1
          If Not IsMissing(ps_VchTyp) Then Call GenFldVal(ws_TcTyp, ws_TblA, Rs_TgtAddA)
          Rs_TgtAddA.Update
        End If
      ElseIf UCase(xmlChildLevel1.baseName) = UCase(ws_TblD) Then
        '*** 25/12/03
        ms_ErrTbl = ws_TblD
        Set xmlErrRecSet = xmlChildLevel1
        '*** 25/12/03
     
        If wb_Flag = True Then
          If CheckForRecExists(ws_TblD, xmlChildLevel1, ws_TcTyp, ws_Cnd) = True Then
            Set Rs_TgtUpdD = moCn.OpenResultset("Select * from " + ws_TblD + " " + ws_Cnd)
            If Rs_TgtUpdD.RecCount > 0 Then
              Rs_TgtUpdD.edit
              MoveAttributesToFields Rs_TgtUpdD, xmlChildLevel1
              If Not IsMissing(ps_VchTyp) Then Call GenFldVal(ws_TcTyp, ws_TblD, Rs_TgtUpdD)
              Rs_TgtUpdD.Update
            End If
            '''Set Rs_TgtUpdD = Nothing
          Else
            Rs_TgtAddD.Addnew
            MoveAttributesToFields Rs_TgtAddD, xmlChildLevel1
            If Not IsMissing(ps_VchTyp) Then Call GenFldVal(ws_TcTyp, ws_TblD, Rs_TgtAddD)
            Rs_TgtAddD.Update
          End If
        Else
          Rs_TgtAddD.Addnew
          MoveAttributesToFields Rs_TgtAddD, xmlChildLevel1
          If Not IsMissing(ps_VchTyp) Then Call GenFldVal(ws_TcTyp, ws_TblD, Rs_TgtAddD)
          Rs_TgtAddD.Update
        End If
      ElseIf UCase(xmlChildLevel1.baseName) = UCase(ws_TblE) Then
        '*** 25/12/03
        ms_ErrTbl = ws_TblE
        Set xmlErrRecSet = xmlChildLevel1
        '*** 25/12/03
      
        If wb_Flag = True Then
          If CheckForRecExists(ws_TblE, xmlChildLevel1, ws_TcTyp, ws_Cnd) = True Then
            Set Rs_TgtUpdE = moCn.OpenResultset("Select * from " + ws_TblE + " " + ws_Cnd)
            If Rs_TgtUpdE.RecCount > 0 Then
              Rs_TgtUpdE.edit
              MoveAttributesToFields Rs_TgtUpdE, xmlChildLevel1
              If Not IsMissing(ps_VchTyp) Then Call GenFldVal(ws_TcTyp, ws_TblE, Rs_TgtUpdE)
              Rs_TgtUpdE.Update
            End If
            '''Set Rs_TgtUpdE = Nothing
          Else
            Rs_TgtAddE.Addnew
            MoveAttributesToFields Rs_TgtAddE, xmlChildLevel1
            If Not IsMissing(ps_VchTyp) Then Call GenFldVal(ws_TcTyp, ws_TblE, Rs_TgtAddE)
            Rs_TgtAddE.Update
          End If
        Else
          Rs_TgtAddE.Addnew
          MoveAttributesToFields Rs_TgtAddE, xmlChildLevel1
          If Not IsMissing(ps_VchTyp) Then Call GenFldVal(ws_TcTyp, ws_TblE, Rs_TgtAddE)
          Rs_TgtAddE.Update
        End If
      ElseIf UCase(xmlChildLevel1.baseName) = UCase(ws_TblG) Then
        '*** 25/12/03
        ms_ErrTbl = ws_TblG
        Set xmlErrRecSet = xmlChildLevel1
        '*** 25/12/03
      
        If wb_Flag = True Then
          If CheckForRecExists(ws_TblG, xmlChildLevel1, ws_TcTyp, ws_Cnd) = True Then
            Set Rs_TgtUpdG = moCn.OpenResultset("Select * from " + ws_TblG + " " + ws_Cnd)
            If Rs_TgtUpdG.RecCount > 0 Then
              Rs_TgtUpdG.edit
              MoveAttributesToFields Rs_TgtUpdG, xmlChildLevel1
              If Not IsMissing(ps_VchTyp) Then Call GenFldVal(ws_TcTyp, ws_TblG, Rs_TgtUpdG)
              Rs_TgtUpdG.Update
            End If
            '''Set Rs_TgtUpdG = Nothing
          Else
            Rs_TgtAddG.Addnew
            MoveAttributesToFields Rs_TgtAddG, xmlChildLevel1
            If Not IsMissing(ps_VchTyp) Then Call GenFldVal(ws_TcTyp, ws_TblG, Rs_TgtAddG)
            Rs_TgtAddG.Update
          End If
        Else
          Rs_TgtAddG.Addnew
          MoveAttributesToFields Rs_TgtAddG, xmlChildLevel1
          If Not IsMissing(ps_VchTyp) Then Call GenFldVal(ws_TcTyp, ws_TblG, Rs_TgtAddG)
          Rs_TgtAddG.Update
        End If
      ' Zubin 213
      ElseIf UCase(xmlChildLevel1.baseName) = UCase(ws_TblH) Then
        '*** 25/12/03
        ms_ErrTbl = ws_TblH
        Set xmlErrRecSet = xmlChildLevel1
        '*** 25/12/03
        
        If wb_Flag = True Then
          If CheckForRecExists(ws_TblH, xmlChildLevel1, ws_TcTyp, ws_Cnd) = True Then
            Set Rs_TgtUpdH = moCn.OpenResultset("Select * from " + ws_TblH + " " + ws_Cnd)
            If Rs_TgtUpdH.RecCount > 0 Then
              Rs_TgtUpdH.edit
              MoveAttributesToFields Rs_TgtUpdH, xmlChildLevel1
              If Not IsMissing(ps_VchTyp) Then Call GenFldVal(ws_TcTyp, ws_TblH, Rs_TgtUpdH)
              Rs_TgtUpdH.Update
            End If
            '''Set Rs_TgtUpdH = Nothing
          Else
            Rs_TgtAddH.Addnew
            MoveAttributesToFields Rs_TgtAddH, xmlChildLevel1
            If Not IsMissing(ps_VchTyp) Then Call GenFldVal(ws_TcTyp, ws_TblH, Rs_TgtAddH)
            Rs_TgtAddH.Update
          End If
        Else
          Rs_TgtAddH.Addnew
          MoveAttributesToFields Rs_TgtAddH, xmlChildLevel1
          If Not IsMissing(ps_VchTyp) Then Call GenFldVal(ws_TcTyp, ws_TblH, Rs_TgtAddH)
          Rs_TgtAddH.Update
        End If
        ' Zubin 213
      End If
      
      For mf_k = 0 To xmlChildLevel1.childNodes.length - 1
        Set xmlChildLevel2 = xmlChildLevel1.childNodes(mf_k)
        
        If UCase(xmlChildLevel2.baseName) = UCase(ws_TblB) Then
          '*** 25/12/03
          ms_ErrTbl = ws_TblB
          Set xmlErrRecSet = xmlChildLevel2
          '*** 25/12/03
        
          If wb_Flag = True Then
            If CheckForRecExists(ws_TblB, xmlChildLevel2, ws_TcTyp, ws_Cnd) = True Then
              Set Rs_TgtUpdB = moCn.OpenResultset("Select * from " + ws_TblB + " " + ws_Cnd)
              If Rs_TgtUpdB.RecCount > 0 Then
                Rs_TgtUpdB.edit
                MoveAttributesToFields Rs_TgtUpdB, xmlChildLevel2
                If Not IsMissing(ps_VchTyp) Then Call GenFldVal(ws_TcTyp, ws_TblB, Rs_TgtUpdB)
                Rs_TgtUpdB.Update
              End If
              '''Set Rs_TgtUpdB = Nothing
            Else
              Rs_TgtAddB.Addnew
              MoveAttributesToFields Rs_TgtAddB, xmlChildLevel2
              If Not IsMissing(ps_VchTyp) Then Call GenFldVal(ws_TcTyp, ws_TblB, Rs_TgtAddB)
              Rs_TgtAddB.Update
            End If
          Else
            Rs_TgtAddB.Addnew
            MoveAttributesToFields Rs_TgtAddB, xmlChildLevel2
            If Not IsMissing(ps_VchTyp) Then Call GenFldVal(ws_TcTyp, ws_TblB, Rs_TgtAddB)
            Rs_TgtAddB.Update
          End If
        ElseIf UCase(xmlChildLevel2.baseName) = UCase(ws_TblC) Then
          '*** 25/12/03
          ms_ErrTbl = ws_TblC
          Set xmlErrRecSet = xmlChildLevel2
          '*** 25/12/03
          
          If wb_Flag = True Then
            If CheckForRecExists(ws_TblC, xmlChildLevel2, ws_TcTyp, ws_Cnd) = True Then
              Set Rs_TgtUpdC = moCn.OpenResultset("Select * from " + ws_TblC + " " + ws_Cnd)
              If Rs_TgtUpdC.RecCount > 0 Then
                Rs_TgtUpdC.edit
                MoveAttributesToFields Rs_TgtUpdC, xmlChildLevel2
                If Not IsMissing(ps_VchTyp) Then Call GenFldVal(ws_TcTyp, ws_TblC, Rs_TgtUpdC)
                Rs_TgtUpdC.Update
              End If
              '''Set Rs_TgtUpdC = Nothing
            Else
              Rs_TgtAddC.Addnew
              MoveAttributesToFields Rs_TgtAddC, xmlChildLevel2
              If Not IsMissing(ps_VchTyp) Then Call GenFldVal(ws_TcTyp, ws_TblC, Rs_TgtAddC)
              Rs_TgtAddC.Update
            End If
          Else
            Rs_TgtAddC.Addnew
            MoveAttributesToFields Rs_TgtAddC, xmlChildLevel2
            If Not IsMissing(ps_VchTyp) Then Call GenFldVal(ws_TcTyp, ws_TblC, Rs_TgtAddC)
            Rs_TgtAddC.Update
          End If
        ElseIf UCase(xmlChildLevel2.baseName) = UCase(ws_TblB1) Then
          '*** 25/12/03
          ms_ErrTbl = ws_TblB1
          Set xmlErrRecSet = xmlChildLevel2
          '*** 25/12/03
        
          If wb_Flag = True Then
            If CheckForRecExists(ws_TblB1, xmlChildLevel2, ws_TcTyp, ws_Cnd) = True Then
              Set Rs_TgtUpdB1 = moCn.OpenResultset("Select * from " + ws_TblB1 + " " + ws_Cnd)
              If Rs_TgtUpdB1.RecCount > 0 Then
                Rs_TgtUpdB1.edit
                MoveAttributesToFields Rs_TgtUpdB1, xmlChildLevel2
                If Not IsMissing(ps_VchTyp) Then Call GenFldVal(ws_TcTyp, ws_TblB1, Rs_TgtUpdB1)
                Rs_TgtUpdB1.Update
              End If
              '''Set Rs_TgtUpdB1 = Nothing
            Else
              Rs_TgtAddB1.Addnew
              MoveAttributesToFields Rs_TgtAddB1, xmlChildLevel2
              If Not IsMissing(ps_VchTyp) Then Call GenFldVal(ws_TcTyp, ws_TblB1, Rs_TgtAddB1)
              Rs_TgtAddB1.Update
            End If
          Else
            Rs_TgtAddB1.Addnew
            MoveAttributesToFields Rs_TgtAddB1, xmlChildLevel2
            If Not IsMissing(ps_VchTyp) Then Call GenFldVal(ws_TcTyp, ws_TblB1, Rs_TgtAddB1)
            Rs_TgtAddB1.Update
          End If
        ElseIf UCase(xmlChildLevel2.baseName) = UCase(ws_TblC1) Then
          '*** 25/12/03
          ms_ErrTbl = ws_TblC1
          Set xmlErrRecSet = xmlChildLevel2
          '*** 25/12/03
        
          If wb_Flag = True Then
            If CheckForRecExists(ws_TblC1, xmlChildLevel2, ws_TcTyp, ws_Cnd) = True Then
              Set Rs_TgtUpdC1 = moCn.OpenResultset("Select * from " + ws_TblC1 + " " + ws_Cnd)
              If Rs_TgtUpdC1.RecCount > 0 Then
                Rs_TgtUpdC1.edit
                MoveAttributesToFields Rs_TgtUpdC1, xmlChildLevel2
                If Not IsMissing(ps_VchTyp) Then Call GenFldVal(ws_TcTyp, ws_TblC1, Rs_TgtUpdC1)
                Rs_TgtUpdC1.Update
              End If
              '''Set Rs_TgtUpdC1 = Nothing
            Else
              Rs_TgtAddC1.Addnew
              MoveAttributesToFields Rs_TgtAddC1, xmlChildLevel2
              If Not IsMissing(ps_VchTyp) Then Call GenFldVal(ws_TcTyp, ws_TblC1, Rs_TgtAddC1)
              Rs_TgtAddC1.Update
            End If
          Else
            Rs_TgtAddC1.Addnew
            MoveAttributesToFields Rs_TgtAddC1, xmlChildLevel2
            If Not IsMissing(ps_VchTyp) Then Call GenFldVal(ws_TcTyp, ws_TblC1, Rs_TgtAddC1)
            Rs_TgtAddC1.Update
          End If
        End If
        Set xmlChildLevel2 = Nothing
        '*** GoTo Next Rec In 2nd-Level Element
      Next mf_k
      Set xmlChildLevel1 = Nothing
      '*** GoTo Next Rec In 1st-Level Element
    Next mf_j
    
    moCn.CommitTrans mdoCommitRec
        
    If UCase(ps_TblX) = UCase("xTxn") Then
      '*** 25/12/03
      ms_ErrTbl = ps_TblX
      Set xmlErrRecSet = xmlChildX
      '*** 25/12/03
    
      If CheckForRecExists(ps_TblX, xmlChildX, , ws_Cnd) = True Then
        wf_xTdTotAmt = moCn.GetFldVal("Select sum(xTdTotAmt) from xTxn " + _
                                      "join xTxnDsg On xTdCoCd= xTCoCd and xTdTc= xTTc " + _
                                      "and xTdYy= xTYy and xTdChr= xTChr and xTdNo= xTNo " + ws_Cnd)
        If wb_Flag = True Then
          Rs_TgtUpdX.edit
          Rs_TgtUpdX!xTTotAmt = wf_xTdTotAmt
          Rs_TgtUpdX!xTNetAmt = Rs_TgtUpdX!xTTotAmt + Rs_TgtUpdX!xTAmt01 + Rs_TgtUpdX!xTAmt02 + _
                                Rs_TgtUpdX!xTAmt03 + Rs_TgtUpdX!xTAmt04 + Rs_TgtUpdX!xTAmt05 + _
                                Rs_TgtUpdX!xTAmt06 + Rs_TgtUpdX!xTAmt07 + Rs_TgtUpdX!xTAmt08
          Rs_TgtUpdX.Update
        Else
          Rs_TgtAddX.edit
          Rs_TgtAddX!xTTotAmt = wf_xTdTotAmt
          Rs_TgtAddX!xTNetAmt = Rs_TgtAddX!xTTotAmt + Rs_TgtAddX!xTAmt01 + Rs_TgtAddX!xTAmt02 + _
                                Rs_TgtAddX!xTAmt03 + Rs_TgtAddX!xTAmt04 + Rs_TgtAddX!xTAmt05 + _
                                Rs_TgtAddX!xTAmt06 + Rs_TgtAddX!xTAmt07 + Rs_TgtAddX!xTAmt08
          Rs_TgtAddX.Update
        End If
      End If
    End If
    Set xmlChildX = Nothing
    
    'Zubin
    '*** Total Records Tranferred
    wi_Cnt = wi_Cnt + 1
    'Zubin
    
'*** GoTo Next X Element
XNext:
  Next mf_i
XClose:
  'Zubin 05/01/04
  ws_TrnfCntStr = "(" + CStr(wi_Cnt) + "/" + CStr(wi_TotCnt) + " Records Transferred)"

  If wi_TotCnt <> wi_Cnt Then
    ADC("wDisp") = ps_TblX + " Transfer-In Completed With Errors. " + ws_TrnfCntStr
  Else
    ADC("wDisp") = ps_TblX + " Transfer-In Completed Successfully. " + ws_TrnfCntStr
  End If
  'Zubin 05/01/04
  
  Set Rs_TgtAddX = Nothing: Set Rs_TgtAddA = Nothing
  Set Rs_TgtAddB = Nothing: Set Rs_TgtAddC = Nothing
  Set Rs_TgtAddB1 = Nothing: Set Rs_TgtAddC1 = Nothing
  Set Rs_TgtAddD = Nothing: Set Rs_TgtAddG = Nothing
  ' Zubin 213
  Set Rs_TgtAddH = Nothing
  
  Set Rs_TgtUpdX = Nothing: Set Rs_TgtUpdA = Nothing
  Set Rs_TgtUpdB = Nothing: Set Rs_TgtUpdC = Nothing
  Set Rs_TgtUpdB1 = Nothing: Set Rs_TgtUpdC1 = Nothing
  Set Rs_TgtUpdD = Nothing: Set Rs_TgtUpdG = Nothing

  ' Zubin 213
  Set Rs_TgtUpdH = Nothing

  '*** 25/12/03
  Set xmlErrRecSet = Nothing
  '*** 25/12/03
  
  Exit Sub
errh:
  If Not (Rs_TgtAddX Is Nothing) Then
    If Not (Rs_TgtAddX.EOF Or Rs_TgtAddX.BOF) Then Rs_TgtAddX.CancelUpdate
  End If
  If Not (Rs_TgtAddA Is Nothing) Then
    If Not (Rs_TgtAddA.EOF Or Rs_TgtAddA.BOF) Then Rs_TgtAddA.CancelUpdate
  End If
  If Not (Rs_TgtAddB Is Nothing) Then
    If Not (Rs_TgtAddB.EOF Or Rs_TgtAddB.BOF) Then Rs_TgtAddB.CancelUpdate
  End If
  If Not (Rs_TgtAddC Is Nothing) Then
    If Not (Rs_TgtAddC.EOF Or Rs_TgtAddC.BOF) Then Rs_TgtAddC.CancelUpdate
  End If
  If Not (Rs_TgtAddB1 Is Nothing) Then
    If Not (Rs_TgtAddB1.EOF Or Rs_TgtAddB1.BOF) Then Rs_TgtAddB1.CancelUpdate
  End If
  If Not (Rs_TgtAddC1 Is Nothing) Then
    If Not (Rs_TgtAddC1.EOF Or Rs_TgtAddC1.BOF) Then Rs_TgtAddC1.CancelUpdate
  End If
  If Not (Rs_TgtAddD Is Nothing) Then
    If Not (Rs_TgtAddD.EOF Or Rs_TgtAddD.BOF) Then Rs_TgtAddD.CancelUpdate
  End If
  If Not (Rs_TgtAddG Is Nothing) Then
    If Not (Rs_TgtAddG.EOF Or Rs_TgtAddG.BOF) Then Rs_TgtAddG.CancelUpdate
  End If
  ' Zubin 213
  If Not (Rs_TgtAddH Is Nothing) Then
    If Not (Rs_TgtAddH.EOF Or Rs_TgtAddH.BOF) Then Rs_TgtAddH.CancelUpdate
  End If
  
  If Not (Rs_TgtUpdX Is Nothing) Then
    If Not (Rs_TgtUpdX.EOF Or Rs_TgtUpdX.BOF) Then Rs_TgtUpdX.CancelUpdate
  End If
  If Not (Rs_TgtUpdA Is Nothing) Then
    If Not (Rs_TgtUpdA.EOF Or Rs_TgtUpdA.BOF) Then Rs_TgtUpdA.CancelUpdate
  End If
  If Not (Rs_TgtUpdB Is Nothing) Then
    If Not (Rs_TgtUpdB.EOF Or Rs_TgtUpdB.BOF) Then Rs_TgtUpdB.CancelUpdate
  End If
  If Not (Rs_TgtUpdC Is Nothing) Then
    If Not (Rs_TgtUpdC.EOF Or Rs_TgtUpdC.BOF) Then Rs_TgtUpdC.CancelUpdate
  End If
  If Not (Rs_TgtUpdB1 Is Nothing) Then
    If Not (Rs_TgtUpdB1.EOF Or Rs_TgtUpdB1.BOF) Then Rs_TgtUpdB1.CancelUpdate
  End If
  If Not (Rs_TgtUpdC1 Is Nothing) Then
    If Not (Rs_TgtUpdC1.EOF Or Rs_TgtUpdC1.BOF) Then Rs_TgtUpdC1.CancelUpdate
  End If
  If Not (Rs_TgtUpdD Is Nothing) Then
    If Not (Rs_TgtUpdD.EOF Or Rs_TgtUpdD.BOF) Then Rs_TgtUpdD.CancelUpdate
  End If
  If Not (Rs_TgtUpdG Is Nothing) Then
    If Not (Rs_TgtUpdG.EOF Or Rs_TgtUpdG.BOF) Then Rs_TgtUpdG.CancelUpdate
  End If
  
  '*** (Bef 25/12/03) wb_SourceAvlbl = Not (xmlChildX Is Nothing)
  '*** 25/12/03
  wb_SourceAvlbl = Not (xmlErrRecSet Is Nothing)
  '*** 25/12/03
  
  If wb_SourceAvlbl Then
    '*** (Bef 25/12/03)
    'Call CheckForRecExists(ps_TblX, xmlChildX, , , ws_KeyVal)
    'ws_KeyStr = ws_KeyVal
    '*** (Bef 25/12/03)
    
    '*** 25/12/03
    Call CheckForRecExists(ms_ErrTbl, xmlErrRecSet, , , ws_KeyVal)
    ws_KeyStr = ws_KeyVal
    '*** 25/12/03
  Else
    ws_KeyStr = "No Source Data"
  End If
  
  '*** (Bef 25/12/03) Call ErrHndlr(CStr(Format(moCn.SrvrDate, "DDMM")) + ps_TblX, ws_KeyStr)
  '*** 25/12/03
  Call ErrHndlr(CStr(Format(moCn.SrvrDate, "DDMM")) + ms_ErrTbl, ws_KeyStr)
  '*** 25/12/03
  
  If wb_SourceAvlbl Then
    Resume XNext
  Else
    Resume XClose
  End If
End Sub
Private Sub MoveAttributesToFields(ByRef pRs_Tgt As MwfLib.MDORowSet, po_ReqXmlElement As IXMLDOMElement)
  '*** Loop Through Attributes Of The Current Element
  '*** The Name Of THe Attribute Is Assumed To Be The Same As Field Of The RecordSet
  '*** Moving The Values
  For mi_i = 0 To po_ReqXmlElement.Attributes.length - 1
    pRs_Tgt(po_ReqXmlElement.Attributes.Item(mi_i).nodeName) = po_ReqXmlElement.Attributes.Item(mi_i).nodeValue
  Next mi_i
End Sub
Private Function CheckForRecExists(fs_Tbl As String, fo_ReqXmlElement As IXMLDOMElement, Optional ByRef fs_TcTyp As Variant, Optional ByRef fs_Cnd As Variant, Optional ByRef fs_KeyVal As Variant) As Boolean
  Dim ws_TcTyp As String, ws_Cnd As String, ws_KeyVal As String
  CheckForRecExists = False
  ws_TcTyp = "": ws_Cnd = "": ws_KeyVal = ""
  
  Select Case UCase(fs_Tbl)
  Case Is = UCase("BnkAgMst")
    Dim ws_BkCtg As String, ws_BkCd As String
    ws_BkCtg = "": ws_BkCd = ""
    For mi_i = 0 To fo_ReqXmlElement.Attributes.length - 1
      If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("BkCtg") Then ws_BkCtg = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
      If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("BkCd") Then ws_BkCd = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
      If ws_BkCtg <> "" And ws_BkCd <> "" Then Exit For
    Next mi_i
    ws_Cnd = " where BkCtg= '" + ws_BkCtg + "' and BkCd= '" + ws_BkCd + "' "
    If Not IsMissing(fs_Cnd) Then CheckForRecExists = moCn.RecSeek("Select 'x' from " + fs_Tbl + " " + ws_Cnd)
    ws_KeyVal = ws_BkCtg + "/" + ws_BkCd
  Case Is = UCase("CustMst")
    Dim ws_CmCtg As String, ws_CmCd As String
    ws_CmCtg = "": ws_CmCd = ""
    For mi_i = 0 To fo_ReqXmlElement.Attributes.length - 1
      If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("CmCtg") Then ws_CmCtg = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
      If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("CmCd") Then ws_CmCd = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
      If ws_CmCtg <> "" And ws_CmCd <> "" Then Exit For
    Next mi_i
    ws_Cnd = " where CmCtg= '" + ws_CmCtg + "' and CmCd= '" + ws_CmCd + "' "
    If Not IsMissing(fs_Cnd) Then CheckForRecExists = moCn.RecSeek("Select 'x' from " + fs_Tbl + " " + ws_Cnd)
    ws_KeyVal = ws_CmCtg + "/" + ws_CmCd
  Case Is = UCase("CustDsg")
    Dim ws_CdCtg As String, ws_CdCd As String, ws_CdOurDmCd As String, ws_CdOurSfx As String
    Dim ws_CdOurDmSz As String
    ws_CdCtg = "": ws_CdCd = "": ws_CdOurDmCd = "": ws_CdOurSfx = "": ws_CdOurDmSz = ""
    For mi_i = 0 To fo_ReqXmlElement.Attributes.length - 1
      If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("CdCtg") Then ws_CdCtg = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
      If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("CdCd") Then ws_CdCd = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
      If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("CdOurDmCd") Then ws_CdOurDmCd = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
      If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("CdOurSfx") Then ws_CdOurSfx = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
      If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("CdOurDmSz") Then ws_CdOurDmSz = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
      If ws_CdCtg <> "" And ws_CdCd <> "" And ws_CdOurDmCd <> "" And ws_CdOurSfx <> "" And ws_CdOurDmSz <> "" Then Exit For
    Next mi_i
    ws_Cnd = " where CdCtg= '" + ws_CdCtg + "' and CdCd= '" + ws_CdCd + "' and CdOurDmCd= '" + ws_CdOurDmCd + "' " + _
             "and CdOurSfx= '" + ws_CdOurSfx + "' and CdOurDmSz= '" + ws_CdOurDmSz + "' "
    If Not IsMissing(fs_Cnd) Then CheckForRecExists = moCn.RecSeek("Select 'x' from " + fs_Tbl + " " + ws_Cnd)
    ws_KeyVal = ws_CdCtg + "/" + ws_CdCd + "/" + ws_CdOurDmCd + "/" + ws_CdOurSfx + "/" + ws_CdOurDmSz
  Case Is = UCase("CustAly")
    Dim ws_CaCtg As String, ws_CaCd As String, ws_CaRmCd As String, ws_CaDmCol As String
    ws_CaCtg = "": ws_CaCd = "": ws_CaRmCd = "": ws_CaDmCol = ""
    For mi_i = 0 To fo_ReqXmlElement.Attributes.length - 1
      If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("CaCtg") Then ws_CaCtg = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
      If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("CaCd") Then ws_CaCd = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
      If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("CaRmCd") Then ws_CaRmCd = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
      If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("CaDmCol") Then ws_CaDmCol = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
      If ws_CaCtg <> "" And ws_CaCd <> "" And ws_CaRmCd <> "" And ws_CaDmCol <> "" Then Exit For
    Next mi_i
    ws_Cnd = " where CaCtg= '" + ws_CaCtg + "' and CaCd= '" + ws_CaCd + "' " + _
             "and CaRmCd= '" + ws_CaRmCd + "' and CaDmCol= '" + ws_CaDmCol + "' "
    If Not IsMissing(fs_Cnd) Then CheckForRecExists = moCn.RecSeek("Select 'x' from " + fs_Tbl + " " + ws_Cnd)
    ws_KeyVal = ws_CaCtg + "/" + ws_CaCd + "/" + ws_CaRmCd + "/" + ws_CaDmCol
  Case Is = UCase("CustRm")
    Dim ws_CrCtg As String, ws_CrCd As String, ws_CrOurRmCd As String, wf_CrFrLn As Single, wf_CrToLn As Single
    ws_CrCtg = "": ws_CrCd = "": ws_CrOurRmCd = "": wf_CrFrLn = 0: wf_CrToLn = 0
    For mi_i = 0 To fo_ReqXmlElement.Attributes.length - 1
      If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("CrCtg") Then ws_CrCtg = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
      If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("CrCd") Then ws_CrCd = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
      If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("CrOurRmCd") Then ws_CrOurRmCd = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
      If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("CrFrLn") Then wf_CrFrLn = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
      If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("CrToLn") Then wf_CrToLn = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
      If ws_CrCtg <> "" And ws_CrCd <> "" And ws_CrOurRmCd <> "" And wf_CrFrLn <> 0 And wf_CrToLn <> 0 Then Exit For
    Next mi_i
    ws_Cnd = " where CrCtg= '" + ws_CrCtg + "' and CrCd= '" + ws_CrCd + "' " + _
             "and CrOurRmCd= '" + ws_CrOurRmCd + "' and CrFrLn= " + CStr(wf_CrFrLn) + " " + _
             "and CrToLn= " + CStr(wf_CrToLn) + " "
    If Not IsMissing(fs_Cnd) Then CheckForRecExists = moCn.RecSeek("Select 'x' from " + fs_Tbl + " " + ws_Cnd)
    ws_KeyVal = ws_CrCtg + "/" + ws_CrCd + "/" + ws_CrOurRmCd + "/" + Trim(CStr(wf_CrFrLn)) + "/" + Trim(CStr(wf_CrToLn))
  Case Is = UCase("DsgMst")
    Dim ws_DmTcTyp As String, ws_DmCd As String, ws_DmSz As String
    ws_DmTcTyp = "": ws_DmCd = "": ws_DmSz = ""
    For mi_i = 0 To fo_ReqXmlElement.Attributes.length - 1
      If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("DmTcTyp") Then ws_DmTcTyp = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
      If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("DmCd") Then ws_DmCd = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
      If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("DmSz") Then ws_DmSz = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
      If ws_DmTcTyp <> "" And ws_DmCd <> "" And ws_DmSz <> "" Then Exit For
    Next mi_i
    ws_Cnd = " where DmTcTyp= '" + ws_DmTcTyp + "' and DmCd= '" + ws_DmCd + "' and DmSz= '" + ws_DmSz + "' "
    If Not IsMissing(fs_Cnd) Then CheckForRecExists = moCn.RecSeek("Select 'x' from " + fs_Tbl + " " + ws_Cnd)
    ws_KeyVal = ws_DmTcTyp + "/" + ws_DmCd + "/" + ws_DmSz
  Case Is = UCase("DsgRm")
    Dim ws_DrTcTyp As String, ws_DrCd As String, ws_DrSz As String, wf_DrSr As Single
    ws_DrTcTyp = "": ws_DrCd = "": ws_DrSz = "": wf_DrSr = 0
    For mi_i = 0 To fo_ReqXmlElement.Attributes.length - 1
      If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("DrTcTyp") Then ws_DrTcTyp = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
      If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("DrCd") Then ws_DrCd = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
      If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("DrSz") Then ws_DrSz = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
      If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("DrSr") Then wf_DrSr = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
      If ws_DrTcTyp <> "" And ws_DrCd <> "" And ws_DrSz <> "" And wf_DrSr <> 0 Then Exit For
    Next mi_i
    ws_Cnd = " where DrTcTyp= '" + ws_DrTcTyp + "' and DrCd= '" + ws_DrCd + "' and DrSz= '" + ws_DrSz + "' and DrSr= " + CStr(wf_DrSr) + " "
    If Not IsMissing(fs_Cnd) Then CheckForRecExists = moCn.RecSeek("Select 'x' from " + fs_Tbl + " " + ws_Cnd)
    ws_KeyVal = ws_DrTcTyp + "/" + ws_DrCd + "/" + ws_DrSz + "/" + Trim(CStr(wf_DrSr))
  Case Is = UCase("DsgLab")
    Dim ws_DlTcTyp As String, ws_DlCd As String, ws_DlSz As String, wf_DlSr As Single
    ws_DlTcTyp = "": ws_DlCd = "": ws_DlSz = "": wf_DlSr = 0
    For mi_i = 0 To fo_ReqXmlElement.Attributes.length - 1
      If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("DlTcTyp") Then ws_DlTcTyp = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
      If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("DlCd") Then ws_DlCd = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
      If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("DlSz") Then ws_DlSz = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
      If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("DlSr") Then wf_DlSr = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
      If ws_DlTcTyp <> "" And ws_DlCd <> "" And ws_DlSz <> "" And wf_DlSr <> 0 Then Exit For
    Next mi_i
    ws_Cnd = " where DlTcTyp= '" + ws_DlTcTyp + "' and DlCd= '" + ws_DlCd + "' and DlSz= '" + ws_DlSz + "' and DlSr= " + CStr(wf_DlSr) + " "
    If Not IsMissing(fs_Cnd) Then CheckForRecExists = moCn.RecSeek("Select 'x' from " + fs_Tbl + " " + ws_Cnd)
    ws_KeyVal = ws_DlTcTyp + "/" + ws_DlCd + "/" + ws_DlSz + "/" + Trim(CStr(wf_DlSr))
  Case Is = UCase("DsgAna")
    Dim ws_DaTcTyp As String, ws_DaCd As String, ws_DaSz As String, wf_DaAnaSr As Single
    ws_DaTcTyp = "": ws_DaCd = "": ws_DaSz = "": wf_DaAnaSr = 0
    For mi_i = 0 To fo_ReqXmlElement.Attributes.length - 1
      If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("DaTcTyp") Then ws_DaTcTyp = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
      If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("DaCd") Then ws_DaCd = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
      If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("DaSz") Then ws_DaSz = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
      If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("DaAnaSr") Then wf_DaAnaSr = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
      If ws_DaTcTyp <> "" And ws_DaCd <> "" And ws_DaSz <> "" And wf_DaAnaSr <> 0 Then Exit For
    Next mi_i
    ws_Cnd = " where DaTcTyp= '" + ws_DaTcTyp + "' and DaCd= '" + ws_DaCd + "' and DaSz= '" + ws_DaSz + "' and DaAnaSr= " + CStr(wf_DaAnaSr) + " "
    If Not IsMissing(fs_Cnd) Then CheckForRecExists = moCn.RecSeek("Select 'x' from " + fs_Tbl + " " + ws_Cnd)
    ws_KeyVal = ws_DaTcTyp + "/" + ws_DaCd + "/" + ws_DaSz + "/" + Trim(CStr(wf_DaAnaSr))
  Case Is = UCase("DsgPrt")
    Dim ws_DpTcTyp As String, ws_DpCd As String, ws_DpSz As String, wf_DpSr As Single
    ws_DpTcTyp = "": ws_DpCd = "": ws_DpSz = "": wf_DpSr = 0
    For mi_i = 0 To fo_ReqXmlElement.Attributes.length - 1
      If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("DpTcTyp") Then ws_DpTcTyp = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
      If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("DpCd") Then ws_DpCd = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
      If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("DpSz") Then ws_DpSz = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
      If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("DpSr") Then wf_DpSr = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
      If ws_DpTcTyp <> "" And ws_DpCd <> "" And ws_DpSz <> "" And wf_DpSr <> 0 Then Exit For
    Next mi_i
    ws_Cnd = " where DpTcTyp= '" + ws_DpTcTyp + "' and DpCd= '" + ws_DpCd + "' and DpSz= '" + ws_DpSz + "' and DpSr= " + CStr(wf_DpSr) + " "
    If Not IsMissing(fs_Cnd) Then CheckForRecExists = moCn.RecSeek("Select 'x' from " + fs_Tbl + " " + ws_Cnd)
    ws_KeyVal = ws_DpTcTyp + "/" + ws_DpCd + "/" + ws_DpSz + "/" + Trim(CStr(wf_DpSr))
  
''''''''''
  Case Is = UCase("DsgModAvl")
    Dim ws_DmaTcTyp As String, ws_DmaCd As String, ws_DmaSz As String, wf_DmaSr As Single
    ws_DmaTcTyp = "": ws_DmaCd = "": ws_DmaSz = "": wf_DmaSr = 0
    For mi_i = 0 To fo_ReqXmlElement.Attributes.length - 1
      If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("DmaTcTyp") Then ws_DmaTcTyp = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
      If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("DmaCd") Then ws_DmaCd = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
      If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("DmaSz") Then ws_DmaSz = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
      If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("DmaSr") Then wf_DmaSr = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
      If ws_DmaTcTyp <> "" And ws_DmaCd <> "" And ws_DmaSz <> "" And wf_DmaSr <> 0 Then Exit For
    Next mi_i
    ws_Cnd = " where DmaTcTyp= '" + ws_DmaTcTyp + "' and DmaCd= '" + ws_DmaCd + "' and DmaSz= '" + ws_DmaSz + "' and DmaSr= " + CStr(wf_DmaSr) + " "
    If Not IsMissing(fs_Cnd) Then CheckForRecExists = moCn.RecSeek("Select 'x' from " + fs_Tbl + " " + ws_Cnd)
    ws_KeyVal = ws_DmaTcTyp + "/" + ws_DmaCd + "/" + ws_DmaSz + "/" + Trim(CStr(wf_DmaSr))
'''''''''''


  Case Is = UCase("DtTbl")
    Dim ws_DtTag As String, ws_DtDt As String
    ws_DtTag = "": ws_DtDt = ""
    For mi_i = 0 To fo_ReqXmlElement.Attributes.length - 1
      If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("DtTag") Then ws_DtTag = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
      If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("DtDt") Then ws_DtDt = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
      If ws_DtTag <> "" And ws_DtDt <> "" Then Exit For
    Next mi_i
    ws_Cnd = " where DtTag= '" + ws_DtTag + "' and DtDt= '" + ws_DtDt + "' "
    If Not IsMissing(fs_Cnd) Then CheckForRecExists = moCn.RecSeek("Select 'x' from " + fs_Tbl + " " + ws_Cnd)
    ws_KeyVal = ws_DtTag + "/" + ws_DtDt
  Case Is = UCase("LabRt")
    Dim ws_LrCmCtg As String, ws_LrCmCd As String, ws_LrCmCurCd As String
    Dim ws_LrMCd As String, ws_LrSCd As String, wf_LrFrWt As Single, wf_LrToWt As Single
    ws_LrCmCtg = "": ws_LrCmCd = "": ws_LrCmCurCd = "": ws_LrMCd = "": ws_LrSCd = ""
    wf_LrFrWt = 0: wf_LrToWt = 0
    For mi_i = 0 To fo_ReqXmlElement.Attributes.length - 1
      If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("LrCmCtg") Then ws_LrCmCtg = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
      If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("LrCmCd") Then ws_LrCmCd = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
      If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("LrCmCurCd") Then ws_LrCmCurCd = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
      If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("LrMCd") Then ws_LrMCd = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
      If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("LrSCd") Then ws_LrSCd = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
      If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("LrFrWt") Then wf_LrFrWt = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
      If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("LrToWt") Then wf_LrToWt = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
      If ws_LrCmCtg <> "" And ws_LrCmCd <> "" And ws_LrCmCurCd <> "" And ws_LrMCd <> "" And ws_LrSCd <> "" And wf_LrFrWt <> 0 And wf_LrToWt <> 0 Then Exit For
    Next mi_i
    ws_Cnd = " where LrCmCtg= '" + ws_LrCmCtg + "' and LrCmCd= '" + ws_LrCmCd + "' " + _
             "and LrCmCurCd= '" + ws_LrCmCurCd + "' and LrMCd= '" + ws_LrMCd + "' " + _
             "and LrSCd= '" + ws_LrSCd + "' " + IIF(gs_Ver = "2.5.0" Or gs_Ver = "2.6.0" Or gs_Ver = "2.7.0" Or gs_Ver = "2.8.0" Or gs_Ver = "2.9.0", " and LrFrWt= " + CStr(wf_LrFrWt) + " " + _
             "and LrToWt= " + CStr(wf_LrToWt) + " ", " ")
    If Not IsMissing(fs_Cnd) Then CheckForRecExists = moCn.RecSeek("Select 'x' from " + fs_Tbl + " " + ws_Cnd)
    ws_KeyVal = ws_LrCmCtg + "/" + ws_LrCmCd + "/" + ws_LrCmCurCd + "/" + ws_LrMCd + "/" + _
                ws_LrSCd + IIF(gs_Ver = "2.5.0" Or gs_Ver = "2.6.0" Or gs_Ver = "2.7.0" Or gs_Ver = "2.8.0" Or gs_Ver = "2.9.0", "/" + CStr(wf_LrFrWt) + "/" + CStr(wf_LrToWt), "")
  Case Is = UCase("LabRtHist")
    Dim wf_LhYyyyMm As Single, ws_LhCmCtg As String, ws_LhCmCd As String, ws_LhCmCurCd As String
    Dim ws_LhMCd As String, ws_LhSCd As String, wf_LhFrWt As Single, wf_LhToWt As Single
    wf_LhYyyyMm = 0: ws_LhCmCtg = "": ws_LhCmCd = "": ws_LhCmCurCd = "": ws_LhMCd = "": ws_LhSCd = ""
    wf_LhFrWt = 0: wf_LhToWt = 0
    For mi_i = 0 To fo_ReqXmlElement.Attributes.length - 1
      If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("LhYyyyMm") Then wf_LhYyyyMm = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
      If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("LhCmCtg") Then ws_LhCmCtg = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
      If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("LhCmCd") Then ws_LhCmCd = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
      If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("LhCmCurCd") Then ws_LhCmCurCd = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
      If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("LhMCd") Then ws_LhMCd = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
      If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("LhSCd") Then ws_LhSCd = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
      If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("LhFrWt") Then wf_LhFrWt = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
      If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("LhToWt") Then wf_LhToWt = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
      If wf_LhYyyyMm <> 0 And ws_LhCmCtg <> "" And ws_LhCmCd <> "" And ws_LhCmCurCd <> "" And ws_LhMCd <> "" And ws_LhSCd <> "" And wf_LhFrWt <> 0 And wf_LhToWt <> 0 Then Exit For
    Next mi_i
    ws_Cnd = " where LhYyyyMm= " + CStr(wf_LhYyyyMm) + " and LhCmCtg= '" + ws_LhCmCtg + "' " + _
             "and LhCmCd= '" + ws_LhCmCd + "' and LhCmCurCd= '" + ws_LhCmCurCd + "' " + _
             "and LhMCd= '" + ws_LhMCd + "' and LhSCd= '" + ws_LhSCd + "' " + _
             IIF(gs_Ver = "2.5.0" Or gs_Ver = "2.6.0" Or gs_Ver = "2.7.0" Or gs_Ver = "2.8.0" Or gs_Ver = "2.9.0", "and LhFrWt= " + CStr(wf_LhFrWt) + " and LhToWt= " + CStr(wf_LhToWt) + " ", "")
    If Not IsMissing(fs_Cnd) Then CheckForRecExists = moCn.RecSeek("Select 'x' from " + fs_Tbl + " " + ws_Cnd)
    ws_KeyVal = CStr(wf_LhYyyyMm) + "/" + ws_LhCmCtg + "/" + ws_LhCmCd + "/" + ws_LhCmCurCd + _
                "/" + ws_LhMCd + "/" + ws_LhSCd + IIF(gs_Ver = "2.5.0" Or gs_Ver = "2.6.0" Or gs_Ver = "2.7.0" Or gs_Ver = "2.8.0" Or gs_Ver = "2.9.0", "/" + CStr(wf_LhFrWt) + "/" + CStr(wf_LhToWt), "")
  Case Is = UCase("Loc")
    Dim ws_LocCoCd As String, ws_LocCd As String
    ws_LocCoCd = "": ws_LocCd = ""
    For mi_i = 0 To fo_ReqXmlElement.Attributes.length - 1
      If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("LocCoCd") Then ws_LocCoCd = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
      If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("LocCd") Then ws_LocCd = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
      If ws_LocCoCd <> "" And ws_LocCd <> "" Then Exit For
    Next mi_i
    ws_Cnd = " where LocCoCd= '" + ws_LocCoCd + "' and LocCd= '" + ws_LocCd + "' "
    If Not IsMissing(fs_Cnd) Then CheckForRecExists = moCn.RecSeek("Select 'x' from " + fs_Tbl + " " + ws_Cnd)
    ws_KeyVal = ws_LocCoCd + "/" + ws_LocCd
  Case Is = UCase("Param")
    Dim ws_PTyp As String, ws_PMCd As String, ws_PSCd As String
    ws_PTyp = "": ws_PMCd = "": ws_PSCd = ""
    For mi_i = 0 To fo_ReqXmlElement.Attributes.length - 1
      If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("PTyp") Then ws_PTyp = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
      If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("PMCd") Then ws_PMCd = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
      If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("PSCd") Then ws_PSCd = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
      If ws_PTyp <> "" And ws_PMCd <> "" And ws_PSCd <> "" Then Exit For
    Next mi_i
    ws_Cnd = " where PTyp= '" + ws_PTyp + "' and PMCd= '" + ws_PMCd + "' and PSCd= '" + ws_PSCd + "' "
    If Not IsMissing(fs_Cnd) Then CheckForRecExists = moCn.RecSeek("Select 'x' from " + fs_Tbl + " " + ws_Cnd)
    ws_KeyVal = ws_PTyp + "/" + ws_PMCd + "/" + ws_PSCd
  Case Is = UCase("RmMst")
    Dim ws_RmCd As String
    ws_RmCd = ""
    For mi_i = 0 To fo_ReqXmlElement.Attributes.length - 1
      If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("RmCd") Then ws_RmCd = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
      If ws_RmCd <> "" Then Exit For
    Next mi_i
    ws_Cnd = " where RmCd= '" + ws_RmCd + "' "
    If Not IsMissing(fs_Cnd) Then CheckForRecExists = moCn.RecSeek("Select 'x' from " + fs_Tbl + " " + ws_Cnd)
    ws_KeyVal = ws_RmCd
  Case Is = UCase("RmAlyCmp")
    Dim ws_RaCd As String, ws_RaCmpCd As String
    ws_RaCd = "": ws_RaCmpCd = ""
    For mi_i = 0 To fo_ReqXmlElement.Attributes.length - 1
      If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("RaCd") Then ws_RaCd = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
      If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("RaCmpCd") Then ws_RaCmpCd = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
      If ws_RmCd <> "" And ws_RaCmpCd <> "" Then Exit For
    Next mi_i
    ws_Cnd = " where RmCd= '" + ws_RmCd + "' RaCmpCd= '" + ws_RaCmpCd + "' "
    If Not IsMissing(fs_Cnd) Then CheckForRecExists = moCn.RecSeek("Select 'x' from " + fs_Tbl + " " + ws_Cnd)
    ws_KeyVal = ws_RmCd + "/" + ws_RaCmpCd
  Case Is = UCase("RmRt")
    Dim ws_RrTcTyp As String, ws_RrCmCtg As String, ws_RrCmCd As String, ws_RrCmCurCd As String
    Dim ws_RrCtg As String, ws_RrSCtg As String, ws_RrSTWGrp As String, ws_RrCd As String
    Dim ws_RrDmCtg As String, ws_RrLsCtg As String, ws_RrLabMCd As String
    Dim wf_RrFrLn As Single, wf_RrToLn As Single
    ' Zubin 212
    Dim ws_RrDmCd As String
    
    ws_RrTcTyp = "": ws_RrCmCtg = "":  ws_RrCmCd = "": ws_RrCmCurCd = "": ws_RrCtg = ""
    ws_RrSCtg = "": ws_RrSTWGrp = "":  ws_RrCd = "": ws_RrDmCtg = ""
    ws_RrLsCtg = "": ws_RrLabMCd = "": wf_RrFrLn = 0: wf_RrToLn = 0
    
    For mi_i = 0 To fo_ReqXmlElement.Attributes.length - 1
      If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("RrTcTyp") Then ws_RrTcTyp = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
      If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("RrCmCtg") Then ws_RrCmCtg = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
      If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("RrCmCd") Then ws_RrCmCd = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
      If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("RrCmCurCd") Then ws_RrCmCurCd = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
      If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("RrCtg") Then ws_RrCtg = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
      If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("RrSCtg") Then ws_RrSCtg = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
      If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("RrSTWGrp") Then ws_RrSTWGrp = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
      If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("RrCd") Then ws_RrCd = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
      If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("RrDmCtg") Then ws_RrDmCtg = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
      If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("RrLsCtg") Then ws_RrLsCtg = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
      If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("RrLabMCd") Then ws_RrLabMCd = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
      If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("RrFrLn") Then wf_RrFrLn = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
      If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("RrToLn") Then wf_RrToLn = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
      
      '****** (Jen 2.13) ******
      ' Zubin 212
      If gs_Ver >= "2.12.0" And gs_Ver <> "2.6.0" Then
        If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("RrDmCd") Then ws_RrDmCd = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
      End If
      
      ' Zubin 212 (ws_RrDmCd added)
      If ws_RrTcTyp <> "" And ws_RrCmCtg <> "" And ws_RrCmCd <> "" And ws_RrCmCurCd <> "" And _
        ws_RrCtg <> "" And ws_RrSCtg <> "" And ws_RrSTWGrp <> "" And ws_RrCd <> "" And _
        ws_RrDmCtg <> "" And ws_RrLsCtg <> "" And ws_RrLabMCd <> "" And _
        wf_RrFrLn <> 0 And wf_RrToLn <> 0 And ws_RrDmCd <> "" Then Exit For
    Next mi_i
    ws_Cnd = " where RrTcTyp= '" + ws_RrTcTyp + "' and RrCmCtg= '" + ws_RrCmCtg + "' " + _
             "and RrCmCd= '" + ws_RrCmCd + "' and RrCmCurCd= '" + ws_RrCmCurCd + "' " + _
             "and RrCtg= '" + ws_RrCtg + "' and RrSCtg= '" + ws_RrSCtg + "' " + _
             "and RrSTWGrp= '" + ws_RrSTWGrp + "' and RrCd= '" + ws_RrCd + "' " + _
             "and RrDmCtg= '" + ws_RrDmCtg + "' and RrLsCtg= '" + ws_RrLsCtg + "' " + _
             "and RrLabMCd= '" + ws_RrLabMCd + "' and RrFrLn= " + CStr(wf_RrFrLn) + " " + _
             "and RrToLn= " + CStr(wf_RrToLn) + " "
    
    '****** (Jen 2.13) ******
    ' Zubin 212
    If gs_Ver >= "2.12.0" And gs_Ver <> "2.6.0" Then
      ws_Cnd = ws_Cnd + " and RrDmCd= '" + ws_RrDmCd + "' "
    End If
    
    If Not IsMissing(fs_Cnd) Then CheckForRecExists = moCn.RecSeek("Select 'x' from " + fs_Tbl + " " + ws_Cnd)
    ' Zubin 212 (ws_RrDmCd added)
    ws_KeyVal = ws_RrTcTyp + "/" + ws_RrCmCtg + "/" + ws_RrCmCd + "/" + ws_RrCmCurCd + "/" + ws_RrCtg + "/" + ws_RrSCtg + "/" + ws_RrDmCd + "/" + ws_RrSTWGrp + "/" + ws_RrCd + "/" + ws_RrDmCtg + "/" + ws_RrLsCtg + "/" + ws_RrLabMCd + "/" + Trim(CStr(wf_RrFrLn)) + "/" + Trim(CStr(wf_RrToLn))
  Case Is = UCase("RmRtHist")
    Dim wf_RhYyyyMm As Single, ws_RhTcTyp As String, ws_RhCmCtg As String, ws_RhCmCd As String, ws_RhCmCurCd As String
    Dim ws_RhCtg As String, ws_RhSCtg As String, ws_RhSTWGrp As String, ws_RhCd As String
    Dim ws_RhDmCtg As String, ws_RhLsCtg As String, ws_RhLabMCd As String
    Dim wf_RhFrLn As Single, wf_RhToLn As Single
    
    wf_RhYyyyMm = 0: ws_RhTcTyp = "": ws_RhCmCtg = "": ws_RhCmCd = "": ws_RhCmCurCd = ""
    ws_RhCtg = "": ws_RrSCtg = "": ws_RhSTWGrp = "": ws_RhCd = "": ws_RhDmCtg = ""
    ws_RrLsCtg = "": ws_RhLabMCd = "": wf_RhFrLn = 0: wf_RhToLn = 0
    
    For mi_i = 0 To fo_ReqXmlElement.Attributes.length - 1
      If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("RhYyyyMm") Then wf_RhYyyyMm = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
      If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("RhTcTyp") Then ws_RhTcTyp = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
      If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("RhCmCtg") Then ws_RhCmCtg = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
      If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("RhCmCd") Then ws_RhCmCd = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
      If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("RhCmCurCd") Then ws_RhCmCurCd = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
      If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("RhCtg") Then ws_RhCtg = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
      If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("RhSCtg") Then ws_RhSCtg = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
      If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("RhSTWGrp") Then ws_RhSTWGrp = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
      If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("RhCd") Then ws_RhCd = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
      If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("RhDmCtg") Then ws_RhDmCtg = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
      If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("RhLsCtg") Then ws_RhLsCtg = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
      If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("RhLabMCd") Then ws_RhLabMCd = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
      If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("RhFrLn") Then wf_RhFrLn = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
      If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("RhToLn") Then wf_RhToLn = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
      
      If wf_RhYyyyMm <> 0 And ws_RhTcTyp <> "" And ws_RhCmCtg <> "" And ws_RhCmCd <> "" And _
        ws_RhCmCurCd <> "" And ws_RhCtg <> "" And ws_RhSCtg <> "" And ws_RhSTWGrp <> "" And _
        ws_RhCd <> "" And ws_RhDmCtg <> "" And ws_RhLsCtg <> "" And ws_RhLabMCd <> "" And _
        wf_RhFrLn <> 0 And wf_RhToLn <> 0 Then Exit For
    Next mi_i
    ws_Cnd = " where RhTcTyp= '" + ws_RhTcTyp + "' and RhCmCtg= '" + ws_RhCmCtg + "' " + _
             "and RhCmCd= '" + ws_RhCmCd + "' and RhCmCurCd= '" + ws_RhCmCurCd + "' " + _
             "and RhCtg= '" + ws_RhCtg + "' and RhSCtg= '" + ws_RhSCtg + "' " + _
             "and RhSTWGrp= '" + ws_RhSTWGrp + "' and RhCd= '" + ws_RhCd + "' " + _
             "and RhDmCtg= '" + ws_RhDmCtg + "' and RhLsCtg= '" + ws_RhLsCtg + "' " + _
             "and RhLabMCd= '" + ws_RhLabMCd + "' and RhFrLn= " + CStr(wf_RhFrLn) + " " + _
             "and RhToLn= " + CStr(wf_RhToLn) + " "
    If Not IsMissing(fs_Cnd) Then CheckForRecExists = moCn.RecSeek("Select 'x' from " + fs_Tbl + " " + ws_Cnd)
    ws_KeyVal = CStr(wf_RhYyyyMm) + "/" + ws_RhTcTyp + "/" + ws_RhCmCtg + "/" + ws_RhCmCd + "/" + ws_RhCmCurCd + "/" + ws_RhCtg + "/" + ws_RhSCtg + "/" + ws_RhSTWGrp + "/" + ws_RhCd + "/" + ws_RhDmCtg + "/" + ws_RhLsCtg + "/" + ws_RhLabMCd + "/" + Trim(CStr(wf_RhFrLn)) + "/" + Trim(CStr(wf_RhToLn))
  Case Is = UCase("vParam")
    Dim ws_vPCoCd As String, ws_vPTyp As String, ws_vPMCd As String, ws_vPSCd As String
    ws_vPCoCd = "": ws_vPTyp = "": ws_vPMCd = "": ws_vPSCd = ""
    For mi_i = 0 To fo_ReqXmlElement.Attributes.length - 1
      If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("vPCoCd") Then ws_vPCoCd = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
      If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("vPTyp") Then ws_vPTyp = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
      If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("vPMCd") Then ws_vPMCd = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
      If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("vPSCd") Then ws_vPSCd = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
      If ws_vPCoCd <> "" And ws_vPTyp <> "" And ws_vPMCd <> "" And ws_vPSCd <> "" Then Exit For
    Next mi_i
    ws_Cnd = " where vPCoCd= '" + ws_vPCoCd + "' and vPTyp= '" + ws_vPTyp + "' and " + _
             "vPMCd= '" + ws_vPMCd + "' and vPSCd= '" + ws_vPSCd + "' "
    If Not IsMissing(fs_Cnd) Then CheckForRecExists = moCn.RecSeek("Select 'x' from " + fs_Tbl + " " + ws_Cnd)
    ws_KeyVal = ws_vPCoCd + "/" + ws_vPTyp + "/" + ws_vPMCd + "/" + ws_vPSCd
  Case Is = UCase("xTxn")
    Dim ws_xTCoCd As String, ws_xTTc As String, ws_xTYy As String, ws_xTChr As String, wf_xTNo As Single
    ws_xTCoCd = "": ws_xTTc = "": ws_xTYy = "": ws_xTChr = "": wf_xTNo = 0
    For mi_i = 0 To fo_ReqXmlElement.Attributes.length - 1
      If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("xTCoCd") Then ws_xTCoCd = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
      If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("xTTc") Then ws_xTTc = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
      If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("xTYy") Then ws_xTYy = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
      If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("xTChr") Then ws_xTChr = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
      If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("xTNo") Then wf_xTNo = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
      If ws_xTCoCd <> "" And ws_xTTc <> "" And ws_xTYy <> "" And ws_xTChr <> "" And wf_xTNo <> 0 Then Exit For
    Next mi_i
    ws_TcTyp = moCn.GetFldVal("Select xPValue from xParam where xPTyp= 'TC' and xPMCd= '" + ws_xTTc + "'")
    'ws_Cnd = " where xTCoCd= '" + ws_xTCoCd + "' and xTTc= '" + ws_xTTc + "' and " + _
              "xTYy= '" + ws_xTYy + "' and xTChr= '" + ws_xTChr + "' and xTNo= " + CStr(wf_xTNo) + " "
    ws_Cnd = " where xTCoCd= '" + gs_CoCd + "' and xTTc= '" + ws_xTTc + "' and " + _
             "xTYy= '" + ws_xTYy + "' and xTChr= '" + ws_xTChr + "' and xTNo= " + CStr(wf_xTNo) + " "
    If Not IsMissing(fs_Cnd) Then CheckForRecExists = moCn.RecSeek("Select 'x' from " + fs_Tbl + " " + ws_Cnd)
    ws_KeyVal = ws_xTCoCd + "/" + ws_xTTc + "/" + ws_xTYy + "/" + ws_xTChr + "/" + Trim(CStr(wf_xTNo))
  Case Is = UCase("xTxnDsg")
    Dim ws_xTdCoCd As String, ws_xTdTc As String, ws_xTdYy As String, ws_xTdChr As String, wf_xTdNo As Single
    Dim wf_xTdSr As Single, wf_xTdSrNo As Single
    ws_xTdCoCd = "": ws_xTdTc = "": ws_xTdYy = "": ws_xTdChr = "": wf_xTdNo = 0: wf_xTdSr = 0: wf_xTdSrNo = 0
    For mi_i = 0 To fo_ReqXmlElement.Attributes.length - 1
      If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("xTdCoCd") Then ws_xTdCoCd = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
      If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("xTdTc") Then ws_xTdTc = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
      If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("xTdYy") Then ws_xTdYy = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
      If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("xTdChr") Then ws_xTdChr = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
      If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("xTdNo") Then wf_xTdNo = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
      If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("xTdSr") Then wf_xTdSr = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
      If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("xTdSrNo") Then wf_xTdSrNo = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
      If ws_xTdCoCd <> "" And ws_xTdTc <> "" And ws_xTdYy <> "" And ws_xTdChr <> "" And wf_xTdNo <> 0 And wf_xTdSr <> 0 Then Exit For
      '''And wf_xTdSrNo <> 0
    Next mi_i
    ws_TcTyp = moCn.GetFldVal("Select xPValue from xParam where xPTyp= 'TC' and xPMCd= '" + ws_xTdTc + "'")
    ws_Cnd = " where xTdCoCd= '" + gs_CoCd + "' and xTdTc= '" + ws_xTdTc + "' " + _
             "and xTdYy= '" + ws_xTdYy + "' and xTdChr= '" + ws_xTdChr + "' and xTdNo= " + CStr(wf_xTdNo) + " " + _
             "and xTdSr= " + CStr(wf_xTdSr) + " and xTdSrNo= " + CStr(wf_xTdSrNo) + " "
    If Not IsMissing(fs_Cnd) Then CheckForRecExists = moCn.RecSeek("Select 'x' from " + fs_Tbl + " " + ws_Cnd)
    ws_KeyVal = ws_xTdCoCd + "/" + ws_xTdTc + "/" + ws_xTdYy + "/" + ws_xTdChr + "/" + Trim(CStr(wf_xTdNo)) + "/" + Trim(CStr(wf_xTdSr)) + "/" + Trim(CStr(wf_xTdSrNo))
  Case Is = UCase("xTxnRm")
    Dim ws_xTrCoCd As String, ws_xTrTc As String, ws_xTrYy As String, ws_xTrChr As String, wf_xTrNo As Single
    Dim wf_xTrSr As Single, wf_xTrSrNo As Single
    ws_xTrCoCd = "": ws_xTrTc = "": ws_xTrYy = "": ws_xTrChr = "": wf_xTrNo = 0: wf_xTrSr = 0: wf_xTrSrNo = 0
    For mi_i = 0 To fo_ReqXmlElement.Attributes.length - 1
      If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("xTrCoCd") Then ws_xTrCoCd = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
      If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("xTrTc") Then ws_xTrTc = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
      If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("xTrYy") Then ws_xTrYy = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
      If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("xTrChr") Then ws_xTrChr = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
      If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("xTrNo") Then wf_xTrNo = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
      If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("xTrSr") Then wf_xTrSr = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
      If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("xTrSrNo") Then wf_xTrSrNo = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
      If ws_xTrCoCd <> "" And ws_xTrTc <> "" And ws_xTrYy <> "" And ws_xTrChr <> "" And wf_xTrNo <> 0 And wf_xTrSr <> 0 And wf_xTrSrNo <> 0 Then Exit For
    Next mi_i
    ws_TcTyp = moCn.GetFldVal("Select xPValue from xParam where xPTyp= 'TC' and xPMCd= '" + ws_xTrTc + "'")
    ws_Cnd = " where xTrCoCd= '" + gs_CoCd + "' and xTrTc= '" + ws_xTrTc + "' " + _
             "and xTrYy= '" + ws_xTrYy + "' and xTrChr= '" + ws_xTrChr + "' and xTrNo= " + CStr(wf_xTrNo) + " " + _
             "and xTrSr= " + CStr(wf_xTrSr) + " and xTrSrNo= " + CStr(wf_xTrSrNo) + " "
    If Not IsMissing(fs_Cnd) Then CheckForRecExists = moCn.RecSeek("Select 'x' from " + fs_Tbl + " " + ws_Cnd)
    ws_KeyVal = ws_xTrCoCd + "/" + ws_xTrTc + "/" + ws_xTrYy + "/" + ws_xTrChr + "/" + Trim(CStr(wf_xTrNo)) + "/" + Trim(CStr(wf_xTrSr)) + "/" + Trim(CStr(wf_xTrSrNo))
  Case Is = UCase("xTxnLab")
    Dim ws_xTlCoCd As String, ws_xTlTc As String, ws_xTlYy As String, ws_xTlChr As String, wf_xTlNo As Single
    Dim wf_xTlSr As Single, wf_xTlSrNo As Single
    ws_xTlCoCd = "": ws_xTlTc = "": ws_xTlYy = "": ws_xTlChr = "": wf_xTlNo = 0: wf_xTlSr = 0: wf_xTlSrNo = 0
    For mi_i = 0 To fo_ReqXmlElement.Attributes.length - 1
      If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("xTlCoCd") Then ws_xTlCoCd = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
      If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("xTlTc") Then ws_xTlTc = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
      If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("xTlYy") Then ws_xTlYy = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
      If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("xTlChr") Then ws_xTlChr = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
      If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("xTlNo") Then wf_xTlNo = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
      If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("xTlSr") Then wf_xTlSr = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
      If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("xTlSrNo") Then wf_xTlSrNo = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
      If ws_xTlCoCd <> "" And ws_xTlTc <> "" And ws_xTlYy <> "" And ws_xTlChr <> "" And wf_xTlNo <> 0 And wf_xTlSr <> 0 And wf_xTlSrNo <> 0 Then Exit For
    Next mi_i
    ws_TcTyp = moCn.GetFldVal("Select xPValue from xParam where xPTyp= 'TC' and xPMCd= '" + ws_xTlTc + "'")
    ws_Cnd = " where xTlCoCd= '" + gs_CoCd + "' and xTlTc= '" + ws_xTlTc + "' " + _
             "and xTlYy= '" + ws_xTlYy + "' and xTlChr= '" + ws_xTlChr + "' and xTlNo= " + CStr(wf_xTlNo) + " " + _
             "and xTlSr= " + CStr(wf_xTlSr) + " and xTlSrNo= " + CStr(wf_xTlSrNo) + " "
    If Not IsMissing(fs_Cnd) Then CheckForRecExists = moCn.RecSeek("Select 'x' from " + fs_Tbl + " " + ws_Cnd)
    ws_KeyVal = ws_xTlCoCd + "/" + ws_xTlTc + "/" + ws_xTlYy + "/" + ws_xTlChr + "/" + Trim(CStr(wf_xTlNo)) + "/" + Trim(CStr(wf_xTlSr)) + "/" + Trim(CStr(wf_xTlSrNo))
  Case Is = UCase("xTxnDsgAmd")
    Dim ws_xTdaCoCd As String, ws_xTdaTc As String, ws_xTdaYy As String, ws_xTdaChr As String, wf_xTdaNo As Single
    Dim wf_xTdaSr As Single, wf_xTdaSrNo As Single
    ws_xTdaCoCd = "": ws_xTdaTc = "": ws_xTdaYy = "": ws_xTdaChr = "": wf_xTdaNo = 0: wf_xTdaSr = 0: wf_xTdaSrNo = 0
    For mi_i = 0 To fo_ReqXmlElement.Attributes.length - 1
      If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("xTdaCoCd") Then ws_xTdaCoCd = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
      If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("xTdaTc") Then ws_xTdaTc = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
      If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("xTdaYy") Then ws_xTdaYy = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
      If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("xTdaChr") Then ws_xTdaChr = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
      If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("xTdaNo") Then wf_xTdaNo = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
      If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("xTdaSr") Then wf_xTdaSr = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
      If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("xTdaSrNo") Then wf_xTdaSrNo = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
      If ws_xTdaCoCd <> "" And ws_xTdaTc <> "" And ws_xTdaYy <> "" And ws_xTdaChr <> "" And wf_xTdaNo <> 0 And wf_xTdaSr <> 0 And wf_xTdaSrNo <> 0 Then Exit For
    Next mi_i
    ws_TcTyp = moCn.GetFldVal("Select xPValue from xParam where xPTyp= 'TC' and xPMCd= '" + ws_xTdaTc + "'")
    ws_Cnd = " where xTdaCoCd= '" + gs_CoCd + "' and xTdaTc= '" + ws_xTdaTc + "' " + _
             "and xTdaYy= '" + ws_xTdaYy + "' and xTdaChr= '" + ws_xTdaChr + "' and xTdaNo= " + CStr(wf_xTdaNo) + " " + _
             "and xTdaSr= " + CStr(wf_xTdaSr) + " and xTdaSrNo= " + CStr(wf_xTdaSrNo) + " "
    If Not IsMissing(fs_Cnd) Then CheckForRecExists = moCn.RecSeek("Select 'x' from " + fs_Tbl + " " + ws_Cnd)
    ws_KeyVal = ws_xTdaCoCd + "/" + ws_xTdaTc + "/" + ws_xTdaYy + "/" + ws_xTdaChr + "/" + Trim(CStr(wf_xTdaNo)) + "/" + Trim(CStr(wf_xTdaSr)) + "/" + Trim(CStr(wf_xTdaSrNo))
  Case Is = UCase("xTxnTgt")
    Dim ws_xTtCoCd As String, ws_xTtTc As String, ws_xTtYy As String, ws_xTtChr As String, wf_xTtNo As Single
    Dim wf_xTtSr As Single, ws_xTtRmCtg As String, ws_xTtRmSCtg As String
    ws_xTtCoCd = "": ws_xTtTc = "": ws_xTtYy = "": ws_xTtChr = "": wf_xTtNo = 0: wf_xTtSr = 0: ws_xTtRmCtg = "": ws_xTtRmSCtg = ""
    For mi_i = 0 To fo_ReqXmlElement.Attributes.length - 1
      If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("xTtCoCd") Then ws_xTtCoCd = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
      If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("xTtTc") Then ws_xTtTc = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
      If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("xTtYy") Then ws_xTtYy = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
      If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("xTtChr") Then ws_xTtChr = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
      If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("xTtNo") Then wf_xTtNo = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
      If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("xTtSr") Then wf_xTtSr = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
      If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("xTtRmCtg") Then ws_xTtRmCtg = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
      If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("xTtRmSCtg") Then ws_xTtRmSCtg = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
      If ws_xTtCoCd <> "" And ws_xTtTc <> "" And ws_xTtYy <> "" And ws_xTtChr <> "" And wf_xTtNo <> 0 And wf_xTtSr <> 0 And ws_xTtRmCtg <> "" And ws_xTtRmSCtg <> "" Then Exit For
    Next mi_i
    ws_TcTyp = moCn.GetFldVal("Select xPValue from xParam where xPTyp= 'TC' and xPMCd= '" + ws_xTtTc + "'")
    ws_Cnd = " where xTtCoCd= '" + gs_CoCd + "' and xTtTc= '" + ws_xTtTc + "' " + _
             "and xTtYy= '" + ws_xTtYy + "' and xTtChr= '" + ws_xTtChr + "' and xTtNo= " + CStr(wf_xTtNo) + " " + _
             "and xTtSr= " + CStr(wf_xTtSr) + " and xTtRmCtg= " + CStr(ws_xTtRmCtg) + " " + _
             "and xTtRmSCtg= " + CStr(ws_xTtRmSCtg) + " "
    If Not IsMissing(fs_Cnd) Then CheckForRecExists = moCn.RecSeek("Select 'x' from " + fs_Tbl + " " + ws_Cnd)
    ws_KeyVal = ws_xTtCoCd + "/" + ws_xTtTc + "/" + ws_xTtYy + "/" + ws_xTtChr + "/" + Trim(CStr(wf_xTtNo)) + "/" + Trim(CStr(wf_xTtSr)) + "/" + ws_xTtRmCtg + "/" + ws_xTtRmSCtg
  End Select

  If Not IsMissing(fs_TcTyp) Then fs_TcTyp = ws_TcTyp
  If Not IsMissing(fs_Cnd) Then fs_Cnd = ws_Cnd
  If Not IsMissing(fs_KeyVal) Then fs_KeyVal = ws_KeyVal
End Function
Private Sub GenFldVal(fs_TcTyp As String, fs_Tbl As String, ByRef fRs_Tbl As MwfLib.MDORowSet)
  Dim Rs_TcTyp As MwfLib.MDORowSet, Rs_AnaFlds As MDORowSet, Rs_xHead As MDORowSet
  Dim Rs_LocDets As MDORowSet
  Dim ws_TgtCmCd As String, ws_TgtCurCd As String, wf_TgtCurCnvRt As Single
  
  ws_TgtCmCd = "": ws_TgtCurCd = "": wf_TgtCurCnvRt = 0
  Set Rs_xHead = moCn.OpenRes("Select * from xHead where xHCd= '" + gs_CoCd + "' ")
  Set Rs_TcTyp = moCn.OpenRes("Select * from xTcTyp where xTpTcTyp= '" + fs_TcTyp + "' ")
  
  If UCase(fs_TcTyp) = "PRD" Then
    ws_TgtCmCd = ADC("wInCmCdTgt")
    ws_TgtCurCd = moCn.GetFldVal("Select CmCurCd from CustMst where CmCtg= '" + Rs_TcTyp!xTpCmCtg + "' and CmCd= '" + ADC("wInCmCdTgt") + "'")
    wf_TgtCurCnvRt = moCn.GetFldVal("Select PNum from Param where PTyp= 'CURNCY' and PMCd= '" + ws_TgtCurCd + "' and PSCd= ''")
  End If

  Select Case UCase(fs_Tbl)
  Case Is = UCase("Loc")
    fRs_Tbl!LocCoCd = gs_CoCd
  Case Is = UCase("vParam")
    fRs_Tbl!vPCoCd = gs_CoCd
  Case Is = UCase("xTxn")
    Set Rs_AnaFlds = moCn.OpenRes("select xPValue1, xPValue2 from xParam where xPCoCd= '" + ctSelfCoCd + "' and " + _
                                  "xPTyp= 'TC' and xPMCd= '" + fRs_Tbl!xTTc + "' ")
    Set Rs_LocDets = moCn.OpenRes("select * from xParam where xPTyp= 'CHR' and " + _
                                  "xPMCd= '" + fRs_Tbl!xTTc + "' and xPSCd= '" + fRs_Tbl!xTChr + "' ")
    fRs_Tbl!xTCoCd = gs_CoCd
    fRs_Tbl!xTOwnBy = gs_CoCd
    fRs_Tbl!xTAppTo = gs_CoCd
    If gs_Ver = "2.5.0" Or gs_Ver = "2.6.0" Then
      fRs_Tbl!xTFrRmLoc = Rs_LocDets("xPValue2")
      fRs_Tbl!xTToRmLoc = Rs_LocDets("xPValue3")
      fRs_Tbl!xTFrRmDC = Rs_LocDets("xPValue4")
      fRs_Tbl!xTFrDsgLoc = Rs_LocDets("xPDesc225")
      fRs_Tbl!xTToDsgLoc = Rs_LocDets("xPValue")
      fRs_Tbl!xTFrDsgDC = Rs_LocDets("xPValue1")
    Else
      fRs_Tbl!xTFrRmLoc = Rs_TcTyp!xTpFrRmLoc
      fRs_Tbl!xTToRmLoc = Rs_TcTyp!xTpToRmLoc
      fRs_Tbl!xTFrRmDC = Rs_TcTyp!xTpFrRmDC
      fRs_Tbl!xTFrDsgLoc = Rs_TcTyp!xTpFrDsgLoc
      fRs_Tbl!xTToDsgLoc = Rs_TcTyp!xTpToDsgLoc
      fRs_Tbl!xTFrDsgDC = Rs_TcTyp!xTpFrDsgDC
    End If
    fRs_Tbl!xTCmCtg = Rs_TcTyp!xTpCmCtg
    
    mf_SrcCurCnvRt = fRs_Tbl!xTCnvRt
    If ws_TgtCmCd <> "" Then fRs_Tbl!xTCmCd = ws_TgtCmCd
    If ws_TgtCurCd <> "" Then fRs_Tbl!xTCurCd = ws_TgtCurCd
    If wf_TgtCurCnvRt <> 0 Then fRs_Tbl!xTCnvRt = wf_TgtCurCnvRt
    
    fRs_Tbl!xTRefCmCtg = Rs_TcTyp!xTpRefCmCtg
    fRs_Tbl!xTAna01 = IIF(UCase(Mid(Rs_AnaFlds!xPValue1, 1, 1)) = "Y", Rs_xHead!xHAna01, "")
    fRs_Tbl!xTAna02 = IIF(UCase(Mid(Rs_AnaFlds!xPValue1, 2, 1)) = "Y", Rs_xHead!xHAna02, "")
    fRs_Tbl!xTAna03 = IIF(UCase(Mid(Rs_AnaFlds!xPValue1, 3, 1)) = "Y", Rs_xHead!xHAna03, "")
    fRs_Tbl!xTAna04 = IIF(UCase(Mid(Rs_AnaFlds!xPValue1, 4, 1)) = "Y", Rs_xHead!xHAna04, "")
    fRs_Tbl!xTAna05 = IIF(UCase(Mid(Rs_AnaFlds!xPValue1, 5, 1)) = "Y", Rs_xHead!xHAna05, "")
    fRs_Tbl!xTAna06 = IIF(UCase(Mid(Rs_AnaFlds!xPValue1, 6, 1)) = "Y", Rs_xHead!xHAna06, "")
    fRs_Tbl!xTAna07 = IIF(UCase(Mid(Rs_AnaFlds!xPValue1, 7, 1)) = "Y", Rs_xHead!xHAna07, "")
    fRs_Tbl!xTAna08 = IIF(UCase(Mid(Rs_AnaFlds!xPValue1, 8, 1)) = "Y", Rs_xHead!xHAna08, "")
    fRs_Tbl!xTKey = moCn.AutoGen("xTxn", "xTKey", gs_CoCd + fRs_Tbl!xTYy)
    '### fRs_Tbl!xTAmt01 = IIf(UCase(Mid(Rs_AnaFlds!xPValue2, 1, 1)) = "Y", 1, 0)
    '### fRs_Tbl!xTAmt02 = IIf(UCase(Mid(Rs_AnaFlds!xPValue2, 2, 1)) = "Y", 1, 0)
    '### fRs_Tbl!xTAmt03 = IIf(UCase(Mid(Rs_AnaFlds!xPValue2, 3, 1)) = "Y", 1, 0)
    '### fRs_Tbl!xTAmt04 = IIf(UCase(Mid(Rs_AnaFlds!xPValue2, 4, 1)) = "Y", 1, 0)
    '### fRs_Tbl!xTAmt05 = IIf(UCase(Mid(Rs_AnaFlds!xPValue2, 5, 1)) = "Y", 1, 0)
    '### fRs_Tbl!xTAmt06 = IIf(UCase(Mid(Rs_AnaFlds!xPValue2, 6, 1)) = "Y", 1, 0)
    '### fRs_Tbl!xTAmt07 = IIf(UCase(Mid(Rs_AnaFlds!xPValue2, 7, 1)) = "Y", 1, 0)
    '### fRs_Tbl!xTAmt08 = IIf(UCase(Mid(Rs_AnaFlds!xPValue2, 8, 1)) = "Y", 1, 0)
  Case Is = UCase("xTxnDsg")
    Dim wi_xTdKeyLen As Integer
    Set Rs_AnaFlds = moCn.OpenRes("select xPValue1, xPValue2 from xParam where xPCoCd= '" + ctSelfCoCd + "' and " + _
                                  "xPTyp= 'TC' and xPMCd= '" + fRs_Tbl!xTdTc + "' ")
    
    Set Rs_LocDets = moCn.OpenRes("select * from xParam where xPTyp= 'CHR' and " + _
                                  "xPMCd= '" + fRs_Tbl!xTdTc + "' and xPSCd= '" + fRs_Tbl!xTdChr + "' ")
    
    fRs_Tbl!xTdCoCd = gs_CoCd
    fRs_Tbl!xTdOwnBy = gs_CoCd
    fRs_Tbl!xTdAppTo = gs_CoCd
    If InStr(1, Rs_TcTyp!xTpId, "N") >= 1 Then
      fRs_Tbl!xTdIdOwnBy = moCn.GetFldVal("Select xHIdOwnBy from xHead where xHCd= '" + gs_CoCd + "' ")
      fRs_Tbl!xTdIdNo = moCn.GetFldVal("Select IsNull(max(xImIdNo), 0) from xIdMst where " + _
                                       "xImCoCd= '" + gs_CoCd + "' and " + _
                                       "xImIdOwn= '" + fRs_Tbl!xTdIdOwnBy + "' and " + _
                                       "xImIdYy= '" + fRs_Tbl!xTdIdYy + "'") + 1
    End If
    
    
    If gs_Ver = "2.5.0" Or gs_Ver = "2.6.0" Then
      fRs_Tbl!xTdFrDsgLoc = Rs_LocDets("xPDesc225")
      fRs_Tbl!xTdToDsgLoc = Rs_LocDets("xPValue")
      fRs_Tbl!xTdFrDsgDc = Rs_LocDets("xPValue1")
    Else
      fRs_Tbl!xTdFrDsgLoc = Rs_TcTyp!xTpFrDsgLoc
      fRs_Tbl!xTdToDsgLoc = Rs_TcTyp!xTpToDsgLoc
      fRs_Tbl!xTdFrDsgDc = Rs_TcTyp!xTpFrDsgDC
    End If
    fRs_Tbl!xTdToDsgDc = IIF(fRs_Tbl!xTdFrDsgDc = "D", "C", IIF(fRs_Tbl!xTdFrDsgDc = "C", "D", ""))
    fRs_Tbl!xTdKey = moCn.AutoGen("xTxnDsg", "xTdKey", gs_CoCd + fRs_Tbl!xTdYy)
    wi_xTdKeyLen = Len(Trim(Str(fRs_Tbl!xTdKey)))
    fRs_Tbl!xTdYyKey = fRs_Tbl!xTdYy + Space(7 - wi_xTdKeyLen) + Trim(Str(fRs_Tbl!xTdKey))
    
    If UCase(Rs_TcTyp!xTpTxnBM) = "B" Then
      fRs_Tbl!xTdTxnRefYy = fRs_Tbl!xTdYy: fRs_Tbl!xTdTxnRefKey = fRs_Tbl!xTdKey
    End If
    
    If UCase(Rs_TcTyp!xTpMemBM) = "B" Then
      fRs_Tbl!xTdMemRefYy = fRs_Tbl!xTdYy: fRs_Tbl!xTdMemRefKey = fRs_Tbl!xTdKey
    End If
    
    fRs_Tbl!xTdDmTyp = Rs_TcTyp!xTpDmTyp
    fRs_Tbl!xTdAna01 = IIF(UCase(Mid(Rs_AnaFlds!xPValue1, 9, 1)) = "Y", Rs_xHead!xHdAna01, "")
    fRs_Tbl!xTdAna02 = IIF(UCase(Mid(Rs_AnaFlds!xPValue1, 10, 1)) = "Y", Rs_xHead!xHdAna02, "")
    fRs_Tbl!xTdAna03 = IIF(UCase(Mid(Rs_AnaFlds!xPValue1, 11, 1)) = "Y", Rs_xHead!xHdAna03, "")
    fRs_Tbl!xTdAna04 = IIF(UCase(Mid(Rs_AnaFlds!xPValue1, 12, 1)) = "Y", Rs_xHead!xHdAna04, "")
    '### fRs_Tbl!xTdAmt01 = IIf(UCase(Mid(Rs_AnaFlds!xPValue2, 9, 1)) = "Y", 1, 0)
    '### fRs_Tbl!xTdAmt02 = IIf(UCase(Mid(Rs_AnaFlds!xPValue2, 10, 1)) = "Y", 1, 0)
    '### fRs_Tbl!xTdAmt03 = IIf(UCase(Mid(Rs_AnaFlds!xPValue2, 11, 1)) = "Y", 1, 0)
    '### fRs_Tbl!xTdAmt04 = IIf(UCase(Mid(Rs_AnaFlds!xPValue2, 12, 1)) = "Y", 1, 0)
   
    If wf_TgtCurCnvRt = 0 Then wf_TgtCurCnvRt = mf_SrcCurCnvRt
    fRs_Tbl!xTdSalValPpc = MWLib.Div(fRs_Tbl!xTdSalValPpc, mf_SrcCurCnvRt) * wf_TgtCurCnvRt
    fRs_Tbl!xTdOffValPpc = MWLib.Div(fRs_Tbl!xTdOffValPpc, mf_SrcCurCnvRt) * wf_TgtCurCnvRt
    fRs_Tbl!xTdCstValPpc = MWLib.Div(fRs_Tbl!xTdCstValPpc, mf_SrcCurCnvRt) * wf_TgtCurCnvRt
    fRs_Tbl!xTdOffVal = MWLib.Div(fRs_Tbl!xTdOffVal, mf_SrcCurCnvRt) * wf_TgtCurCnvRt
    fRs_Tbl!xTdSalVal = MWLib.Div(fRs_Tbl!xTdSalVal, mf_SrcCurCnvRt) * wf_TgtCurCnvRt
    fRs_Tbl!xTdCstVal = MWLib.Div(fRs_Tbl!xTdCstVal, mf_SrcCurCnvRt) * wf_TgtCurCnvRt
    fRs_Tbl!xTdTotAmt = fRs_Tbl!xTdOffVal + fRs_Tbl!xTdAmt01 + fRs_Tbl!xTdAmt02 + _
                        fRs_Tbl!xTdAmt03 + fRs_Tbl!xTdAmt04
  Case Is = UCase("xTxnRm")
    Set Rs_AnaFlds = moCn.OpenRes("select xPValue1, xPValue2 from xParam where xPCoCd= '" + ctSelfCoCd + "' and " + _
                                  "xPTyp= 'TC' and xPMCd= '" + fRs_Tbl!xTrTc + "' ")
    
    Set Rs_LocDets = moCn.OpenRes("select * from xParam where xPTyp= 'CHR' and " + _
                                  "xPMCd= '" + fRs_Tbl!xTrTc + "' and xPSCd= '" + fRs_Tbl!xTrChr + "' ")
    
    fRs_Tbl!xTrCoCd = gs_CoCd
    fRs_Tbl!xTrOwnBy = gs_CoCd
    fRs_Tbl!xTrAppTo = gs_CoCd
    
    If gs_Ver = "2.5.0" Or gs_Ver = "2.6.0" Then
      fRs_Tbl!xTrFrRmLoc = Rs_LocDets("xPValue2")
      fRs_Tbl!xTrToRmLoc = Rs_LocDets("xPValue3")
      fRs_Tbl!xTrFrRmDc = Rs_LocDets("xPValue4")
    Else
      fRs_Tbl!xTrFrRmLoc = Rs_TcTyp!xTpFrRmLoc
      fRs_Tbl!xTrToRmLoc = Rs_TcTyp!xTpToRmLoc
      fRs_Tbl!xTrFrRmDc = Rs_TcTyp!xTpFrRmDC
    End If
    fRs_Tbl!xTrToRmDC = IIF(fRs_Tbl!xTrFrRmDc = "D", "C", IIF(fRs_Tbl!xTrFrRmDc = "C", "D", ""))
    
    If wf_TgtCurCnvRt = 0 Then wf_TgtCurCnvRt = mf_SrcCurCnvRt
    fRs_Tbl!xTrSalRt = MWLib.Div(fRs_Tbl!xTrSalRt, mf_SrcCurCnvRt) * wf_TgtCurCnvRt
    fRs_Tbl!xTrSalValPPc = MWLib.Div(fRs_Tbl!xTrSalValPPc, mf_SrcCurCnvRt) * wf_TgtCurCnvRt
    fRs_Tbl!xTrCstRt = MWLib.Div(fRs_Tbl!xTrCstRt, mf_SrcCurCnvRt) * wf_TgtCurCnvRt
    fRs_Tbl!xTrCstValPPc = MWLib.Div(fRs_Tbl!xTrCstValPPc, mf_SrcCurCnvRt) * wf_TgtCurCnvRt
    fRs_Tbl!xTrSetSalRt = MWLib.Div(fRs_Tbl!xTrSetSalRt, mf_SrcCurCnvRt) * wf_TgtCurCnvRt
    fRs_Tbl!xTrSetSalValPPc = MWLib.Div(fRs_Tbl!xTrSetSalValPPc, mf_SrcCurCnvRt) * wf_TgtCurCnvRt
    fRs_Tbl!xTrSetCstRt = MWLib.Div(fRs_Tbl!xTrSetCstRt, mf_SrcCurCnvRt) * wf_TgtCurCnvRt
    fRs_Tbl!xTrSetCstValPPc = MWLib.Div(fRs_Tbl!xTrSetCstValPPc, mf_SrcCurCnvRt) * wf_TgtCurCnvRt
    fRs_Tbl!xTrAlySalRt = MWLib.Div(fRs_Tbl!xTrAlySalRt, mf_SrcCurCnvRt) * wf_TgtCurCnvRt
    fRs_Tbl!xTrAlyCstRt = MWLib.Div(fRs_Tbl!xTrAlyCstRt, mf_SrcCurCnvRt) * wf_TgtCurCnvRt
    fRs_Tbl!xTrAlyCstRt = MWLib.Div(fRs_Tbl!xTrAlyCstRt, mf_SrcCurCnvRt) * wf_TgtCurCnvRt
  Case Is = UCase("xTxnLab")
    fRs_Tbl!xTLCoCd = gs_CoCd
    fRs_Tbl!xTLOwnBy = gs_CoCd
    fRs_Tbl!xTLAppTo = gs_CoCd
    If wf_TgtCurCnvRt = 0 Then wf_TgtCurCnvRt = mf_SrcCurCnvRt
    fRs_Tbl!xTLSalRt = MWLib.Div(fRs_Tbl!xTLSalRt, mf_SrcCurCnvRt) * wf_TgtCurCnvRt
    fRs_Tbl!xTLSalValPpc = MWLib.Div(fRs_Tbl!xTLSalValPpc, mf_SrcCurCnvRt) * wf_TgtCurCnvRt
    fRs_Tbl!xTLCstRt = MWLib.Div(fRs_Tbl!xTLCstRt, mf_SrcCurCnvRt) * wf_TgtCurCnvRt
    fRs_Tbl!xTLCstValPpc = MWLib.Div(fRs_Tbl!xTLCstValPpc, mf_SrcCurCnvRt) * wf_TgtCurCnvRt
  Case Is = UCase("xTxnDsgAmd")
    fRs_Tbl!xTdaCoCd = gs_CoCd
    fRs_Tbl!xTdaOwnBy = gs_CoCd
    fRs_Tbl!xTdaAppTo = gs_CoCd
  Case Is = UCase("xTxnTgt")
    fRs_Tbl!xTtCoCd = gs_CoCd
    fRs_Tbl!xTtOwnBy = gs_CoCd
    fRs_Tbl!xTtAppTo = gs_CoCd
  End Select
End Sub
Private Function OpenDocFilePath(fs_XTblName As String) As String
  '##$$ On Error GoTo ErrHndlr
  
  Dim wl_XMLSizeInKB As Single, wo_FileSysObj, wo_MyFile
  
  ComDialogOpn.DialogTitle = "Open " + fs_XTblName + " XML File As"
  ComDialogOpn.InitDir = moCn.GetFldVal("Select PDesc225 from Param where PTyp= 'TRFDIR' and PMCd= 'TRFDIR' ")
  ComDialogOpn.Filter = "XML Files (*.xml)"
  ComDialogOpn.FilterIndex = 1
  
  '##$$ ComDialogOpn.CancelError = True '***
  
  '*** cdlOFNOverwritePrompt:     specifies whether the user should be alerted if he is overwriting and existing file
  '*** cdlOFNFileMustExist:       Specifies that the user can enter only names of existing files in the File Name text box
  '*** cdlOFNPathMustExist:       makes sure that the the user should specify an existing path
  '*** cdlOFNLongNames:           specifies that long names can be used for the files
  '*** cdlOFNExtensionDifferent:  indicates that the user has selected a different extension than the one specified through the DefaultExt property
  ComDialogOpn.Flags = cdlOFNLongNames Or cdlOFNFileMustExist Or cdlOFNPathMustExist Or cdlOFNExtensionDifferent
  ComDialogOpn.DefaultExt = "xml"
  ComDialogOpn.ShowOpen
    
  If Len(ComDialogOpn.FileName) > 4 And Trim(ComDialogOpn.FileName) <> "" And Dir(ComDialogOpn.FileName) <> "" Then
    If UCase(Right(ComDialogOpn.FileName, 4)) = ".XML" Then
      OpenDocFilePath = ComDialogOpn.FileName
    Else
      OpenDocFilePath = ""
      DispMsg "File Extension Is Wrong", etError: Exit Function
    End If
  End If

  If OpenDocFilePath <> "" Then
    Set wo_FileSysObj = CreateObject("Scripting.FileSystemObject")
    Set wo_MyFile = wo_FileSysObj.GetFile(ComDialogOpn.FileName)
    wl_XMLSizeInKB = MWLib.Div(wo_MyFile.Size, 1024)
    '*** For Rounding Up By One
    wl_XMLSizeInKB = IIF(wl_XMLSizeInKB > Round(wl_XMLSizeInKB), 1, 0) + Round(wl_XMLSizeInKB)
    If wl_XMLSizeInKB > 1200 Then
      OpenDocFilePath = ""
      DispMsg "The File Size Was Exceeding The Maximum Permissible Limits Hence Cannot Be Opened", etError
    End If
  End If
  
'##$$
'ErrHndlr:
'  If Err.Number = 32755 Then
'    ADC("wDisp") = "Operation Cancelled "
'    DispMsg "Operation Cancelled ", etInfo
'    Err.Clear
'  End If
'##$$
End Function
Private Sub ShowErrLog()
  Dim wRs_ErrLog As MwfLib.MDORowSet, wi_Row As Single
  Dim ws_QuryStr As String, ws_Cnd As String
  Dim wi_SrNo As Single
  
  If ADC("wTblFr") <> "" Then ws_Cnd = " SubString(ElTable, 5, Len(ElTable)-4) >= '" + ADC("wTblFr") + "' and "
  If ADC("wTblTo") <> "" Then ws_Cnd = ws_Cnd + " substring(ElTable, 5, Len(ElTable)-4) <= '" + ADC("wTblTo") + "' and "
  If ADC("wDdMmFr") <> "" Then ws_Cnd = ws_Cnd + " substring(ElTable, 1, 4) >= '" + ADC("wDdMmFr") + "' and "
  If ADC("wDdMmTo") <> "" Then ws_Cnd = ws_Cnd + " substring(ElTable, 1, 4) <= '" + ADC("wDdMmTo") + "' and "
  
  If Right(ws_Cnd, 4) = "and " Then ws_Cnd = Mid(ws_Cnd, 1, Len(ws_Cnd) - 4)
  
  If ws_Cnd <> "" Then ws_Cnd = " where " + ws_Cnd
  ws_QuryStr = "Select * from ErrLog " + ws_Cnd + _
                " Order By ElTable, ElKey, ElErrMsg "
  Set wRs_ErrLog = moCn.OpenRes(ws_QuryStr)
    
  GrdErrLog.Rows = 1
  With wRs_ErrLog
    Do While Not (.EOF Or .BOF)
      GrdErrLog.AddItem
      wi_Row = GrdErrLog.Rows - 1
      wi_SrNo = wi_SrNo + 1
      GrdErrLog.Value(wi_Row, "wSrNo") = wi_SrNo
      GrdErrLog.Value(wi_Row, "wElTable") = Mid(!ElTable, 5)
      GrdErrLog.Value(wi_Row, "wElKey") = !ElKey
      GrdErrLog.Value(wi_Row, "wElErrMsg") = !ElErrMsg
      .MoveNext
    Loop
  End With
  
  Set wRs_ErrLog = Nothing
End Sub
Private Sub DispFra(ByVal pv_DispFra As en_EmrTrfInDispFra)
  '*** This is the common procedure which is used to toggle between hiding and displaying
      'the frames (i.e. if the frame is visible then make it invisible and viseversa)
      'associated with a command button

  Select Case pv_DispFra
  Case Is = Rm
    If FraMst.Visible = True Then
      FraMst.Visible = False
      '*** required to pass the array position Rm-0, to resotre the 'Causes Validation' State
      Call EnaDisaCmds(False, CmdRm, Rm)
      CmdRm.SetFocus
    Else
      FraMst.Visible = True
      CmdMstTrfIn.SetFocus
      FraMst.ZOrder
      Call EnaDisaCmds(True, CmdRm)
    End If
  Case Is = Cust
    If FraMst.Visible = True Then
      FraMst.Visible = False
      Call EnaDisaCmds(False, CmdCust, Cust)
      CmdCust.SetFocus
    Else
      FraMst.Visible = True
      CmdMstTrfIn.SetFocus
      FraMst.ZOrder
      Call EnaDisaCmds(True, CmdCust)
    End If
  Case Is = Dsg
    If FraMst.Visible = True Then
      FraMst.Visible = False
      Call EnaDisaCmds(False, CmdDsg, Dsg)
      CmdDsg.SetFocus
    Else
      FraMst.Visible = True
      CmdMstTrfIn.SetFocus
      FraMst.ZOrder
      Call EnaDisaCmds(True, CmdDsg)
    End If
  Case Is = LabRt
    If FraMst.Visible = True Then
      FraMst.Visible = False
      Call EnaDisaCmds(False, CmdLabRt, LabRt)
      CmdLabRt.SetFocus
    Else
      FraMst.Visible = True
      CmdMstTrfIn.SetFocus
      FraMst.ZOrder
      Call EnaDisaCmds(True, CmdLabRt)
    End If
  Case Is = RmRt
    If FraMst.Visible = True Then
      FraMst.Visible = False
      Call EnaDisaCmds(False, CmdRmRt, RmRt)
      CmdRmRt.SetFocus
    Else
      FraMst.Visible = True
      CmdMstTrfIn.SetFocus
      FraMst.ZOrder
      Call EnaDisaCmds(True, CmdRmRt)
    End If
  Case Is = LabRtHist
    If FraMst.Visible = True Then
      FraMst.Visible = False
      Call EnaDisaCmds(False, CmdLabRtHist, LabRtHist)
      CmdLabRtHist.SetFocus
    Else
      FraMst.Visible = True
      CmdMstTrfIn.SetFocus
      FraMst.ZOrder
      Call EnaDisaCmds(True, CmdLabRtHist)
    End If
  Case Is = RmRtHist
    If FraMst.Visible = True Then
      FraMst.Visible = False
      Call EnaDisaCmds(False, CmdRmRtHist, RmRtHist)
      CmdRmRtHist.SetFocus
    Else
      FraMst.Visible = True
      CmdMstTrfIn.SetFocus
      FraMst.ZOrder
      Call EnaDisaCmds(True, CmdRmRtHist)
    End If
  Case Is = DtTbl
    If FraMst.Visible = True Then
      FraMst.Visible = False
      Call EnaDisaCmds(False, CmdDtTbl, DtTbl)
      CmdDtTbl.SetFocus
    Else
      FraMst.Visible = True
      CmdMstTrfIn.SetFocus
      FraMst.ZOrder
      Call EnaDisaCmds(True, CmdDtTbl)
    End If
  Case Is = BkAg
    If FraMst.Visible = True Then
      FraMst.Visible = False
      Call EnaDisaCmds(False, CmdBkAg, BkAg)
      CmdBkAg.SetFocus
    Else
      FraMst.Visible = True
      'Before 030604
      'ADC("wBkCtg").SetFocus
      FraMst.ZOrder
      Call EnaDisaCmds(True, CmdBkAg)
    End If
  Case Is = Loc
    If FraMst.Visible = True Then
      FraMst.Visible = False
      Call EnaDisaCmds(False, CmdLoc, Loc)
      CmdLoc.SetFocus
    Else
      FraMst.Visible = True
      CmdMstTrfIn.SetFocus
      FraMst.ZOrder
      Call EnaDisaCmds(True, CmdLoc)
    End If
  Case Is = Prm
    If FraMst.Visible = True Then
      FraMst.Visible = False
      Call EnaDisaCmds(False, CmdPrm, Prm)
      CmdPrm.SetFocus
    Else
      FraMst.Visible = True
      CmdMstTrfIn.SetFocus
      FraMst.ZOrder
      Call EnaDisaCmds(True, CmdPrm)
    End If
  Case Is = VPrm
    If FraMst.Visible = True Then
      FraMst.Visible = False
      Call EnaDisaCmds(False, CmdVPrm, VPrm)
      CmdVPrm.SetFocus
    Else
      FraMst.Visible = True
      CmdMstTrfIn.SetFocus
      FraMst.ZOrder
      Call EnaDisaCmds(True, CmdVPrm)
    End If
  Case Is = Ord
    If FraMst.Visible = True Then
      FraMst.Visible = False
      Call EnaDisaCmds(False, CmdOrd, Ord)
      CmdOrd.SetFocus
    Else
      FraMst.Visible = True
      CmdMstTrfIn.SetFocus
      FraMst.ZOrder
      Call EnaDisaCmds(True, CmdOrd)
    End If
  Case Is = Inv
    If FraInv.Visible = True Then
      FraInv.Visible = False
      Call EnaDisaCmds(False, CmdInv, Inv)
      CmdInv.SetFocus
    Else
      FraInv.Visible = True
      'ADC("wInCmCdTgt").SetFocus
      CmdInv.SetFocus
      FraInv.ZOrder
      Call EnaDisaCmds(True, CmdInv)
    End If
  Case Is = ErrLog
    If FraErrLog.Visible = True Then
      FraErrLog.Visible = False
      Call EnaDisaCmds(False, CmdShowErrLog, ErrLog)
      CmdShowErrLog.SetFocus
    Else
      FraErrLog.Visible = True
      ADC("wTblFr").SetFocus
      FraErrLog.ZOrder
      Call EnaDisaCmds(True, CmdShowErrLog)
    End If
  Case Is = AllMst
    If FraSelFolder.Visible = True Then
      FraSelFolder.Visible = False
      Call EnaDisaCmds(False, CmdAllMst, AllMst)
      CmdAllMst.SetFocus
    Else
      FraSelFolder.Visible = True
      CmdAllMst.SetFocus
      FraSelFolder.ZOrder
      Call EnaDisaCmds(True, CmdAllMst)
    End If
  End Select
End Sub
Private Sub EnaDisaCmds(ByVal pv_ShowFra As Boolean, Optional ByVal pv_CmdBut As CommandButton, Optional ByVal pv_DispFra As en_EmrTrfInDispFra)
  '*** This procedure is used to give the toggle option (enable and disable) to
      'the command buttons

  '*** storing the state of all cmd buttons, when the frame is going to be visible.
  Dim X As Integer
  If pv_ShowFra = True Then
    Call StoreState
  End If

  If Not pv_CmdBut Is Nothing Then
    If pv_ShowFra = True Then
        pv_CmdBut.BackColor = vbWhite
        pv_CmdBut.FontBold = True
        pv_CmdBut.CausesValidation = True
    Else
        pv_CmdBut.BackColor = &H8000000F
        pv_CmdBut.FontBold = False
        '*** restoring the 'Causes Validation' state of the current pv_cmdbutton
        If Not IsMissing(pv_DispFra) Then
          pv_CmdBut.CausesValidation = mArr_Cv(pv_DispFra)
        End If
    End If
  End If
  
  ''FraNKeyAll.Enabled = Not pv_ShowFra
  ''ADC.AllowSave = Not pv_ShowFra
  ''ADC.AllowDelete = Not pv_ShowFra
  
  If pv_ShowFra = True Then
    '** All command buttons are disabled
    CmdRm.Enabled = Not pv_ShowFra
    CmdCust.Enabled = Not pv_ShowFra
    CmdDsg.Enabled = Not pv_ShowFra
    CmdRmRt.Enabled = Not pv_ShowFra
    CmdLabRt.Enabled = Not pv_ShowFra
    CmdRmRtHist.Enabled = Not pv_ShowFra
    CmdLabRtHist.Enabled = Not pv_ShowFra
    CmdDtTbl.Enabled = Not pv_ShowFra
    CmdBkAg.Enabled = Not pv_ShowFra
    CmdLoc.Enabled = Not pv_ShowFra
    CmdPrm.Enabled = Not pv_ShowFra
    CmdVPrm.Enabled = Not pv_ShowFra
    CmdOrd.Enabled = Not pv_ShowFra
    CmdInv.Enabled = Not pv_ShowFra
    CmdClearErrLog.Enabled = Not pv_ShowFra
    CmdShowErrLog.Enabled = Not pv_ShowFra
    CmdAllMst.Enabled = Not pv_ShowFra
  Else
    '*** Restoring Ena/Disa state of all Command Buttons
    CmdRm.Enabled = mArr_Ed(Rm)
    CmdCust.Enabled = mArr_Ed(Cust)
    CmdDsg.Enabled = mArr_Ed(Dsg)
    CmdRmRt.Enabled = mArr_Ed(RmRt)
    CmdLabRt.Enabled = mArr_Ed(LabRt)
    CmdRmRtHist.Enabled = mArr_Ed(RmRtHist)
    CmdLabRtHist.Enabled = mArr_Ed(LabRtHist)
    CmdDtTbl.Enabled = mArr_Ed(DtTbl)
    CmdBkAg.Enabled = mArr_Ed(BkAg)
    CmdLoc.Enabled = mArr_Ed(Loc)
    CmdPrm.Enabled = mArr_Ed(Prm)
    CmdVPrm.Enabled = mArr_Ed(VPrm)
    CmdOrd.Enabled = mArr_Ed(Ord)
    CmdInv.Enabled = mArr_Ed(Inv)
    CmdClearErrLog.Enabled = mArr_Ed(ErrLog)
    CmdShowErrLog.Enabled = mArr_Ed(ErrLog)
    CmdAllMst.Enabled = mArr_Ed(AllMst)
  End If

  If Not pv_CmdBut Is Nothing Then pv_CmdBut.Enabled = True
End Sub
Private Sub HideAllFras()
  '*** This procedure hides all the frames other than FraNKeyAll
  'FraNKeyAll.Visible = True
  FraMst.Visible = False
  FraInv.Visible = False
  FraErrLog.Visible = False
  FraSelFolder.Visible = False
End Sub
Private Sub StoreState()
  '*** This procedure is used to store the CausesValidation State and the Enabled/Disabled
      'states of all the Command Buttons
  '*** This is used for restoring the original states in the future
  mArr_Cv(Rm) = CmdRm.CausesValidation
  mArr_Cv(Cust) = CmdCust.CausesValidation
  mArr_Cv(Dsg) = CmdDsg.CausesValidation
  mArr_Cv(RmRt) = CmdRmRt.CausesValidation
  mArr_Cv(LabRt) = CmdLabRt.CausesValidation
  mArr_Cv(RmRtHist) = CmdRmRtHist.CausesValidation
  mArr_Cv(LabRtHist) = CmdLabRtHist.CausesValidation
  mArr_Cv(DtTbl) = CmdDtTbl.CausesValidation
  mArr_Cv(BkAg) = CmdBkAg.CausesValidation
  mArr_Cv(Loc) = CmdLoc.CausesValidation
  mArr_Cv(Prm) = CmdPrm.CausesValidation
  mArr_Cv(VPrm) = CmdVPrm.CausesValidation
  mArr_Cv(Ord) = CmdOrd.CausesValidation
  mArr_Cv(Inv) = CmdInv.CausesValidation
  mArr_Cv(ErrLog) = CmdShowErrLog.CausesValidation
  mArr_Cv(AllMst) = CmdAllMst.CausesValidation
  
  mArr_Ed(Rm) = CmdRm.Enabled
  mArr_Ed(Cust) = CmdCust.Enabled
  mArr_Ed(Dsg) = CmdDsg.Enabled
  mArr_Ed(RmRt) = CmdRmRt.Enabled
  mArr_Ed(LabRt) = CmdLabRt.Enabled
  mArr_Ed(RmRtHist) = CmdRmRtHist.Enabled
  mArr_Ed(LabRtHist) = CmdLabRtHist.Enabled
  mArr_Ed(DtTbl) = CmdDtTbl.Enabled
  mArr_Ed(BkAg) = CmdBkAg.Enabled
  mArr_Ed(Loc) = CmdLoc.Enabled
  mArr_Ed(Prm) = CmdPrm.Enabled
  mArr_Ed(VPrm) = CmdVPrm.Enabled
  mArr_Ed(Ord) = CmdOrd.Enabled
  mArr_Ed(Inv) = CmdInv.Enabled
  mArr_Ed(ErrLog) = CmdShowErrLog.Enabled
  mArr_Ed(AllMst) = CmdAllMst.Enabled
End Sub
Private Sub ErrHndlr(ByVal TableName As String, ByVal KeyVal As String)
  Dim ws_ReturnStr As String
  If InStr(1, Err.Description, "|") > 0 Then
    ws_ReturnStr = Mid(Err.Description, InStr(1, Err.Description, "|") + 1)
  Else
    ws_ReturnStr = Err.Description
    Dim wf_Ctr As Single, ws_ErrMsg As String, ws_ExecStr As String
    For wf_Ctr = 0 To moCn.ErrCount - 1
      If moCn.ErrObj(wf_Ctr).Number = -2147217900 Then
        ws_ReturnStr = ""
        ws_ErrMsg = "Record Already Exists"
      Else
        If InStr(1, moCn.ErrObj(wf_Ctr).Description, "|") > 0 Then
          ws_ErrMsg = Mid(moCn.ErrObj(wf_Ctr).Description, InStr(1, moCn.ErrObj(wf_Ctr).Description, "|") + 1)
        Else
          ws_ErrMsg = moCn.ErrObj(wf_Ctr).Description
        End If
      End If
      ws_ReturnStr = ws_ReturnStr + "::" + ws_ErrMsg
    Next wf_Ctr
  End If
  ws_ReturnStr = ws_ReturnStr + "   " + KeyVal

  ADC("wDisp") = ws_ReturnStr
  DoEvents

  '*** This statement is not needed here as code for Rollback exists in the trigger
  'moCn.RollbackTrans mdoCommitRec
  '*** This statement is not needed here as code for Rollback exists in the trigger
  
  If InStr(1, TableName + KeyVal + ws_ReturnStr, "'") <> 0 Then
    If InStr(1, KeyVal, "'") = 0 Then
      ws_ExecStr = "insert into ErrLog values('" + TableName + "','" + KeyVal + "','Cannot Display Msg As It Contains Single Quotes')"
    Else
      ws_ExecStr = "insert into ErrLog values('" + TableName + "','Key Contains Quotes','Cannot Display Msg As It Contains Single Quotes')"
    End If
  Else
    ws_ExecStr = "insert into ErrLog values('" + TableName + "','" + KeyVal + "','" + ws_ReturnStr + "')"
  End If

  moCn.Execute (ws_ExecStr)
  DispMsg ws_ReturnStr, etInfo
  moCn.ErrClear
End Sub


'*** Std Code Not To Be Changed
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
  ws_BtnStr = "CmdErrLogGo"
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
    
    '*** (09/08/05)
    Set moCn = Nothing
    Set XmlDoc = Nothing
    Set xmlRootElement = Nothing
    Set xmlChildX = Nothing
    Set xmlChildLevel1 = Nothing
    Set xmlChildLevel2 = Nothing
    '*** (09/08/05)

End Sub
'***


