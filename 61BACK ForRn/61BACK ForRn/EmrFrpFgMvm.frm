VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{AB2CC32A-E652-42A2-B341-26ECE02F6386}#2.1#0"; "mwfctlforspeed9.ocx"
Begin VB.Form EmrFrpFgMvm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Transaction Listing"
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
      Left            =   12030
      TabIndex        =   27
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
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   9630
      Width           =   6330
      _ExtentX        =   11165
      _ExtentY        =   1005
   End
   Begin VB.Frame fra_keys 
      BorderStyle     =   0  'None
      Height          =   465
      Left            =   5985
      TabIndex        =   37
      Top             =   9600
      Width           =   3555
      Begin MwfCtl.MWCTL_MED aktxt 
         Height          =   345
         Index           =   0
         Left            =   210
         TabIndex        =   38
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
         TabIndex        =   39
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
      Left            =   7200
      TabIndex        =   36
      Top             =   9615
      Visible         =   0   'False
      Width           =   855
      Begin CRVIEWER9LibCtl.CRViewer9 CRV_REP 
         Height          =   495
         Left            =   480
         TabIndex        =   33
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
      TabIndex        =   29
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
      TabIndex        =   32
      Top             =   60
      Width           =   15255
      Begin TabDlg.SSTab TAB_REP 
         Height          =   9465
         Left            =   60
         TabIndex        =   31
         Top             =   30
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   16695
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "&Data Range"
         TabPicture(0)   =   "EmrFrpFgMvm.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "fra_tabrep(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "&Group Sort"
         TabPicture(1)   =   "EmrFrpFgMvm.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "fra_tabrep(1)"
         Tab(1).ControlCount=   1
         Begin VB.Frame fra_tabrep 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   9070
            Index           =   1
            Left            =   -74940
            TabIndex        =   35
            Top             =   360
            Width           =   15105
            Begin MwfCtl.MWCTL_GRP GRP_REP 
               Height          =   360
               Left            =   240
               TabIndex        =   30
               Top             =   360
               Width           =   8805
               _ExtentX        =   15531
               _ExtentY        =   635
            End
         End
         Begin VB.Frame fra_tabrep 
            BorderStyle     =   0  'None
            Height          =   9070
            Index           =   0
            Left            =   60
            TabIndex        =   34
            Top             =   360
            Width           =   15105
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   46
               Left            =   4275
               TabIndex        =   14
               ToolTipText     =   "Enter To Voucher Date"
               Top             =   1095
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   3
               DataField       =   "UODTTO"
               IdName          =   "UODTTO"
               CmpStr          =   "FmDt <="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   48
               Left            =   1935
               TabIndex        =   13
               ToolTipText     =   "Enter From Voucher Date"
               Top             =   1095
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   3
               DataField       =   "UODTFR"
               IdName          =   "UODTFR"
               CmpStr          =   "FmDt >="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   24
               Left            =   5055
               TabIndex        =   7
               ToolTipText     =   "Enter To Transaction Year"
               Top             =   525
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               MaxLength       =   2
               DataField       =   "UOTYYTO"
               IdName          =   "UOTYYTO"
               CmpStr          =   "fmdYy <="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   27
               Left            =   2715
               TabIndex        =   4
               ToolTipText     =   "Enter From Transaction Year"
               Top             =   525
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               MaxLength       =   2
               DataField       =   "UOTYYFR"
               IdName          =   "UOTYYFR"
               CmpStr          =   "FmdYy >="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   19
               Left            =   1935
               TabIndex        =   3
               ToolTipText     =   "Enter From Transaction Tc"
               Top             =   525
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataField       =   "UOTTCFR"
               IdName          =   "UOTTCFR"
               CmpStr          =   "FmdTc >="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   49
               Left            =   3135
               TabIndex        =   5
               ToolTipText     =   "Enter From Transaction Character"
               Top             =   525
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOTCHRFR"
               IdName          =   "UOTCHRFR"
               CmpStr          =   "FmdChr >="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   50
               Left            =   4275
               TabIndex        =   6
               ToolTipText     =   "Enter To Transaction Tc"
               Top             =   525
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataField       =   "UOTTCTO"
               IdName          =   "UOTTCTO"
               CmpStr          =   "FmdTc <="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   51
               Left            =   5475
               TabIndex        =   8
               ToolTipText     =   "Enter To Transaction Character"
               Top             =   525
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOTCHRTO"
               IdName          =   "UOTCHRTO"
               CmpStr          =   "FmdChr <="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   21
               Left            =   5175
               TabIndex        =   12
               ToolTipText     =   "Enter To Transaction Serial"
               Top             =   810
               Width           =   525
               _ExtentX        =   926
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOTDSRTO"
               IdName          =   "UOTDSRTO"
               CmpStr          =   "FmdSr <="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   22
               Left            =   2835
               TabIndex        =   10
               ToolTipText     =   "Enter From Transaction Serial"
               Top             =   810
               Width           =   525
               _ExtentX        =   926
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOTDSRFR"
               IdName          =   "UOTDSRFR"
               CmpStr          =   "FmdSr >="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   25
               Left            =   1935
               TabIndex        =   9
               ToolTipText     =   "Enter From Transaction Number"
               Top             =   810
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOTNOFR"
               IdName          =   "UOTNOFR"
               CmpStr          =   "FmdNo >="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   55
               Left            =   4275
               TabIndex        =   11
               ToolTipText     =   "Enter To Transaction Number"
               Top             =   810
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOTNOTO"
               IdName          =   "UOTNOTO"
               CmpStr          =   "FmdNo <="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   35
               Left            =   4275
               TabIndex        =   18
               ToolTipText     =   "Enter To Bag Year"
               Top             =   1380
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               MaxLength       =   2
               DataField       =   "UOBYYTO"
               IdName          =   "UOBYYTO"
               CmpStr          =   "FmdBYy <="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   36
               Left            =   1935
               TabIndex        =   15
               ToolTipText     =   "Enter From Bag Year"
               Top             =   1380
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               MaxLength       =   2
               DataField       =   "UOBYYFR"
               IdName          =   "UOBYYFR"
               CmpStr          =   "FmdBYy >="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   38
               Left            =   2355
               TabIndex        =   16
               ToolTipText     =   "Enter From Bag Character"
               Top             =   1380
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOBCHRFR"
               IdName          =   "UOBCHRFR"
               CmpStr          =   "FmdBChr >="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   40
               Left            =   4725
               TabIndex        =   19
               ToolTipText     =   "Enter To Bag Character"
               Top             =   1380
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOBCHRTO"
               IdName          =   "UOBCHRTO"
               CmpStr          =   "FmdBChr <="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   41
               Left            =   3165
               TabIndex        =   17
               ToolTipText     =   "Enter From Bag Number"
               Top             =   1380
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOBNOFR"
               IdName          =   "UOBNOFR"
               CmpStr          =   "FmdBNo >="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   42
               Left            =   5535
               TabIndex        =   20
               ToolTipText     =   "Enter To Bag Number"
               Top             =   1380
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOBNOTO"
               IdName          =   "UOBNOTO"
               CmpStr          =   "FmdBNo <="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   68
               Left            =   4275
               TabIndex        =   22
               ToolTipText     =   "Enter To Issuing FgSubLoc Location"
               Top             =   1665
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataField       =   "UOFGSUBLOCTO"
               IdName          =   "UOFGSUBLOCTO"
               CmpStr          =   "FmdFrFgSubLoc <="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   69
               Left            =   6660
               TabIndex        =   23
               ToolTipText     =   "Enter Issuing FgSubLoc Location Selection"
               Top             =   1665
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOFGSUBLOCSEL"
               IdName          =   "UOFGSUBLOCSEL"
               CmpStr          =   "FmdFrFgSubLoc In"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   70
               Left            =   1935
               TabIndex        =   21
               ToolTipText     =   "Enter From Issuing FgSubLoc Location"
               Top             =   1665
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataField       =   "UOFGSUBLOCFR"
               IdName          =   "UOFGSUBLOCFR"
               CmpStr          =   "FmdFrFgSubLoc >="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   0
               Left            =   4275
               TabIndex        =   25
               ToolTipText     =   "Enter To Receiving FgSubLoc Location"
               Top             =   1950
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataField       =   "UORECFGSUBLOCTO"
               IdName          =   "UORECFGSUBLOCTO"
               CmpStr          =   "FmdToFgSubLoc <="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   1
               Left            =   6660
               TabIndex        =   26
               ToolTipText     =   "Enter Receiving FgSubLoc Location Selection"
               Top             =   1950
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UORECFGSUBLOCSEL"
               IdName          =   "UORECFGSUBLOCSEL"
               CmpStr          =   "FmdToFgSubLoc In"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   2
               Left            =   1935
               TabIndex        =   24
               ToolTipText     =   "Enter From Receiving FgSubLoc Location"
               Top             =   1950
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataField       =   "UORECFGSUBLOCFR"
               IdName          =   "UORECFGSUBLOCFR"
               CmpStr          =   "FmdToFgSubLoc >="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   74
               Left            =   4275
               TabIndex        =   1
               ToolTipText     =   "Enter To Company Code"
               Top             =   240
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOCOCDTO"
               IdName          =   "UOCOCDTO"
               CmpStr          =   "FmdCoCd <= "
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   75
               Left            =   6660
               TabIndex        =   2
               ToolTipText     =   "Enter Design Code Selection"
               Top             =   240
               Width           =   8520
               _ExtentX        =   15028
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOCOCDSEL"
               IdName          =   "UOCOCDSEL"
               CmpStr          =   "FmdCoCd In"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   76
               Left            =   1935
               TabIndex        =   0
               ToolTipText     =   "Enter Company Code"
               Top             =   240
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOCOCDFR"
               IdName          =   "UOCOCDFR"
               CmpStr          =   "FmdCoCd >= "
            End
            Begin VB.Label Label3 
               BackStyle       =   0  'Transparent
               Caption         =   "Receiving SubLoc"
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
               Left            =   30
               TabIndex        =   59
               Top             =   1950
               Width           =   1905
            End
            Begin VB.Label Label2 
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
               Left            =   5460
               TabIndex        =   58
               Top             =   1380
               Width           =   105
            End
            Begin VB.Label Label1 
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
               Left            =   3090
               TabIndex        =   57
               Top             =   1380
               Width           =   105
            End
            Begin VB.Label lblFgSubLoc 
               BackStyle       =   0  'Transparent
               Caption         =   "Issuing SubLoc"
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
               Left            =   30
               TabIndex        =   56
               Top             =   1665
               Width           =   1815
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
               Index           =   0
               Left            =   30
               TabIndex        =   55
               Top             =   240
               Width           =   1725
            End
            Begin VB.Label LblBagSlh1 
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
               Left            =   2295
               TabIndex        =   54
               Top             =   1380
               Width           =   105
            End
            Begin VB.Label LblBagSlh2 
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
               Left            =   4635
               TabIndex        =   53
               Top             =   1380
               Width           =   105
            End
            Begin VB.Label LblBagYyChr 
               BackStyle       =   0  'Transparent
               Caption         =   "Bag Yy/ Chr/ No"
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
               Left            =   30
               TabIndex        =   52
               Top             =   1380
               Width           =   1725
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
               Left            =   30
               TabIndex        =   51
               Top             =   525
               Width           =   1725
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
               Left            =   30
               TabIndex        =   50
               Top             =   810
               Width           =   1725
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
               Left            =   5415
               TabIndex        =   49
               Top             =   525
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
               Left            =   3075
               TabIndex        =   48
               Top             =   525
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
               Left            =   4995
               TabIndex        =   47
               Top             =   525
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
               Left            =   2655
               TabIndex        =   46
               Top             =   525
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
               Left            =   5115
               TabIndex        =   45
               Top             =   810
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
               Left            =   2775
               TabIndex        =   44
               Top             =   810
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
               Left            =   30
               TabIndex        =   43
               Top             =   1095
               Width           =   1725
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
               TabIndex        =   42
               Top             =   0
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
               TabIndex        =   41
               Top             =   0
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
               Left            =   1935
               TabIndex        =   40
               Top             =   0
               Width           =   1335
            End
         End
      End
   End
End
Attribute VB_Name = "EmrFrpFgMvm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rep As New EmrRepFgMvm
Dim moCn As MwfLib.MDOConnection   '*** (Bef speed) bef 24/11/06
Dim ws_MultiCoMod As Boolean

Private Sub SetGroupSort()
  GRP_REP.Add "FrFgSubLoc", "FmdFrFgSubLoc", "", "", "", "", "", "", ""
  GRP_REP.Add "ToFgSubLoc", "FmdToFgSubLoc", "", "", "", "", "", "", ""
  '6.1
  GRP_REP.Add "Date", "convert(char(10),FmDt,112)", "convert(varchar, FmDt, " + CStr(DtFmtSlahVal()) + ")", "", "", "", "", "", ""    '4.1.2
  If ws_MultiCoMod = True Then
        GRP_REP.Add "Company Code", "FmdCoCd", "FmdCoCd", "", "", "", "", "", "(Select HName from Head where HCd= '" + ctSelfCmCd + "' and HCoCd=FmdCoCd) "
  End If

End Sub
Private Sub adc_setreprecsource()
On Error GoTo RepErr
  '*** Report Sql ***
'  Select space(30) as Grp1, space(30) as Grp2, space(30) as Grp3,
'  space(30) as DGrp1, space(30) as DGrp2, space(30) as DGrp3,
'  space(30) as G1Desc, space(30) as G2Desc, space(30) as G3Desc,
'  FmTc, FmYY, FmChr, FmNo, FmDt, FmdSr, FmdByy, FmdBChr, FmdBno,
'  FmdFrFgSubLoc, FmdToFgSubLoc, FdQty, FdGrWt
'  From Fm, Fmd, Bag, Fgd Where 1 = 2

  Dim grpflds As String, wCnd As String, wSqlStrg As String, wrepcnd As String
  
  Set Rep = Nothing
   
  Set adc.RepSource = Rep
  
  Call DispCoNm
  Call SetMwName(Rep)   ' ***** Manali 3.6.1 - Microway footer
  Call GRP_REP.Gen3LRep
  grpflds = GRP_REP.GrpFldLst
    
  wrepcnd = adc.RepCond
  wCnd = IIF(wrepcnd <> "", " and " + wrepcnd, "")
             
  wSqlStrg = "Select " + grpflds + ", " + _
             "FmdCoCd, FmTc, FmYY, FmChr, FmNo, FmDt, FmdSr, FmdByy, FmdBChr, FmdBno, BOdDmCd,BOdSfx,BOdDmSz, " + _
             "FmdFrFgSubLoc, FmdToFgSubLoc, FdQty, FdGrWt " + _
             "From FM " + _
             "Join Fmd on FmdFmIdNo=FmIdNo and FmdCoCd = FmCoCd and FmdTc = FmTc and FmdYy = FmYy and FmdChr = FmChr and FmdNo = FmNo " + IIF(gs_Partition = ctCurrPrtn, " and FmdPrtKey=FmPrtKey ", "") + _
             "Join Bag on BcoCd = FmdCoCd and Byy = FmdByy and BChr = FmdBChr and Bno = FmdBno " + IIF(gs_Partition = ctCurrPrtn, " and BPrtKey=FmdPrtKey ", "") + _
             "Join Fgd on FdCoCd = FmdCoCd and FdYy = FmdFdYy and FdKey = FmdFdKey" + " " + IIF(gs_Partition = ctCurrPrtn, " and FdPrtKey=FmdPrtKey ", "") + _
             " Where " + IIF(gs_Partition = ctCurrPrtn, " FmPrtKey='" + ctCurrPrtn + "' and ", "") + " (Select PValue from Param where PTyp= 'TC' and PMCd= FmdTc)= '" + gs_TxnTcTyp + "' " + _
             wCnd + " Order By FmdTc, FmdYY, FmdChr, FmdNo, FmdSr, FmDt"

  adc.RepRecSource = wSqlStrg + ctMaxDopOpt  '****** Sachin 3.01 ctMaxDopOpt added
  CRV_REP.DisplayGroupTree = False
  
  Call DispReport(adc) '****** Sachin 3.02 - Display Report ******
  
Exit Sub
RepErr:
  Call ChkSelScope(Err.Number, Err.Description)
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
  ws_MultiCoMod = IIF(moCn.GetFldVal("Select hImpComLoc From Head Where HCd='" + ctSelfCmCd + "' And HCoCd='" + ctSelfCoCd + "'") = "Y", True, False)

  Call InitProp(Me)
  If ws_MultiCoMod = True Then
    Set adc.FirNKeyCtl = adc("UoCoCdFr")
    adc("UOCOCDFR").Enabled = True: adc("UOCOCDTO").Visible = True: adc("UOCOCDSEL").Visible = True
  Else
    Set adc.FirNKeyCtl = adc("UoTTcFr")
    adc("UOCOCDFR").Enabled = False: adc("UOCOCDTO").Visible = False: adc("UOCOCDSEL").Visible = False
  End If
  
  gs_TxnTcTyp = "FM"
  Call SetGroupSort
End Sub
Private Sub ADC_NonKeyWhen(pr_Cancel As Boolean, pr_ErrMsg As String)
 If ws_MultiCoMod = False Then _
      adc("UoCoCdFr") = gs_CoCd: adc("UoCoCdTo") = gs_CoCd: adc("UoCoCdSel") = "": adc("UoCoCdFr").Enabled = False

'  adc("UoCoCdFr") = gs_CoCd
End Sub
Private Sub ATXT_FldChange(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String, ByVal pv_NewValue As Variant)
  gs_TxnTcTyp = "FM"
  Call SetProp(Me, IdName, Valid, Cancel, ErrMsg)
End Sub
Private Sub ATXT_FldWhen(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String)
  gs_TxnTcTyp = "FM"
  SetProp Me, IdName, When
End Sub

Private Sub tab_REP_Click(PreviousTab As Integer)
  fra_tabrep(TAB_REP.Tab).Enabled = True
  fra_tabrep(PreviousTab).Enabled = False
End Sub
Private Sub ADC_KeyWhen()
  adc("UOUSRCD") = adc.UsrCd
  adc("UOMNUCD") = adc.MenuCd
End Sub
Private Sub DispCoNm()
  Dim ws_HName As String
  ws_HName = moCn.GetFldVal("Select HName from Head where HCoCd ='" + IIF(ws_MultiCoMod = True, gs_CoCd, adc("UOCOCDFR")) + "' and " + _
                                                      "HCd = '" + ctSelfCmCd + "' ")
  
'  ws_HName = moCn.GetFldVal("Select HName from Head where HCoCd ='" + adc("UoCoCdFr") + "' and " + _
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
  
' ********* Manali 3.2.2 - Before 3.2.2
'  Dim ws_DtTm As String
'  ws_DtTm = Format$(moCn.SrvrDate, "DD-MMM-YYYY")
'  Rep.wSrvrDate.SetText ws_DtTm
'  ws_DtTm = Replace(Format(IIF(CDbl(moCn.SrvrTime) >= 13#, moCn.SrvrTime - 12, IIF(CDbl(moCn.SrvrTime) < 1#, moCn.SrvrTime + 12, moCn.SrvrTime)), "00.00"), ".", ":") + IIF(CDbl(moCn.SrvrTime) >= 12#, " PM", " AM")
'  Rep.wSrvrTime.SetText ws_DtTm
 If ws_MultiCoMod = True Then
      Rep.TxtCoCdToLbl.Suppress = False: Rep.TxtCoCdSelLbl.Suppress = False
      Rep.TxtCoCdTo.SetText adc("UOCOCDTO"): Rep.TxtCoCdSel.SetText adc("UOCOCDSEL")
  Else
      Rep.TxtCoCdTo.Suppress = True: Rep.TxtCoCdSel.Suppress = True
      Rep.TxtCoCdToLbl.Suppress = True: Rep.TxtCoCdSelLbl.Suppress = True
  End If
  Rep.TxtCoCdFr.SetText adc("UOCOCDFR")
 
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
