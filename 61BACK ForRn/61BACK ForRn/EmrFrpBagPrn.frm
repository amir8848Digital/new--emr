VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{AB2CC32A-E652-42A2-B341-26ECE02F6386}#2.1#0"; "mwfctlforspeed9.ocx"
Begin VB.Form EmrFrpBagPrn 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bag Printing"
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
      Left            =   11070
      TabIndex        =   39
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
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   9630
      Width           =   6360
      _ExtentX        =   11218
      _ExtentY        =   1005
   End
   Begin VB.Frame fra_keys 
      BorderStyle     =   0  'None
      Height          =   465
      Left            =   5430
      TabIndex        =   48
      Top             =   9570
      Width           =   3555
      Begin MwfCtl.MWCTL_MED aktxt 
         Height          =   345
         Index           =   0
         Left            =   210
         TabIndex        =   49
         Top             =   105
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
         TabIndex        =   50
         Top             =   120
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
      Left            =   8730
      TabIndex        =   47
      Top             =   9660
      Visible         =   0   'False
      Width           =   855
      Begin CRVIEWER9LibCtl.CRViewer9 CRV_REP 
         Height          =   495
         Left            =   480
         TabIndex        =   44
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
      TabIndex        =   40
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
      Height          =   9615
      Left            =   0
      TabIndex        =   43
      Top             =   60
      Width           =   15225
      Begin TabDlg.SSTab TAB_REP 
         Height          =   9465
         Left            =   15
         TabIndex        =   42
         Top             =   45
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   16695
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "&Data Range"
         TabPicture(0)   =   "EmrFrpBagPrn.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "fra_tabrep(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "&Group Sort"
         TabPicture(1)   =   "EmrFrpBagPrn.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "fra_tabrep(1)"
         Tab(1).ControlCount=   1
         Begin VB.Frame fra_tabrep 
            Enabled         =   0   'False
            Height          =   9150
            Index           =   1
            Left            =   -75000
            TabIndex        =   46
            Top             =   300
            Width           =   15180
            Begin MwfCtl.MWCTL_GRP GRP_REP 
               Height          =   360
               Left            =   240
               TabIndex        =   41
               Top             =   360
               Width           =   9765
               _ExtentX        =   17224
               _ExtentY        =   635
            End
         End
         Begin VB.Frame fra_tabrep 
            BorderStyle     =   0  'None
            Height          =   9070
            Index           =   0
            Left            =   60
            TabIndex        =   45
            Top             =   360
            Width           =   15105
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Show Single Bag Per Page"
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
               Index           =   2
               Left            =   60
               TabIndex        =   37
               Tag             =   "ShowBagRec"
               ToolTipText     =   "Check To Show Single Bag Per Page"
               Top             =   4830
               Width           =   1890
            End
            Begin VB.Frame FraOptRm 
               BorderStyle     =   0  'None
               Height          =   285
               Left            =   1725
               TabIndex        =   77
               ToolTipText     =   "Select to Shoe either Order Wts Or Actual Wts"
               Top             =   4260
               Width           =   4485
               Begin VB.OptionButton OptRm 
                  Caption         =   "Order"
                  BeginProperty Font 
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
                  Left            =   2490
                  TabIndex        =   79
                  Top             =   0
                  Width           =   1305
               End
               Begin VB.OptionButton OptRm 
                  Caption         =   "Actual"
                  BeginProperty Font 
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
                  TabIndex        =   78
                  Top             =   0
                  Width           =   1305
               End
            End
            Begin VB.Frame FraOptPtr 
               BorderStyle     =   0  'None
               Height          =   285
               Left            =   1725
               TabIndex        =   76
               ToolTipText     =   "Select either Productuion Pointer or Sales Pointer"
               Top             =   4545
               Width           =   4485
               Begin VB.OptionButton OptPtr 
                  Caption         =   "Sales Ptr"
                  BeginProperty Font 
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
                  Left            =   2490
                  TabIndex        =   35
                  Top             =   0
                  Width           =   1305
               End
               Begin VB.OptionButton OptPtr 
                  Caption         =   "Prd Ptr"
                  BeginProperty Font 
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
                  TabIndex        =   34
                  Top             =   0
                  Width           =   1305
               End
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   1
               Left            =   4275
               TabIndex        =   26
               ToolTipText     =   "Enter To Current Bag Location"
               Top             =   3120
               Width           =   1320
               _ExtentX        =   2328
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UOISSLOCTO"
               IdName          =   "UOISSLOCTO"
               CmpStr          =   "BLoc<= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   2
               Left            =   6840
               TabIndex        =   27
               ToolTipText     =   "Enter Current Bag Location Selection"
               Top             =   3120
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOISSLOCSEL"
               IdName          =   "UOISSLOCSEL"
               CmpStr          =   "BLoc In "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   0
               Left            =   1725
               TabIndex        =   25
               ToolTipText     =   "Enter From Current Bag Location"
               Top             =   3120
               Width           =   1320
               _ExtentX        =   2328
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UOISSLOCFR"
               IdName          =   "UOISSLOCFR"
               CmpStr          =   "BLoc>= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   18
               Left            =   5055
               TabIndex        =   5
               ToolTipText     =   "Enter To Order Year"
               Top             =   1410
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               MaxLength       =   2
               DataField       =   "UOOMYYTO"
               IdName          =   "UOOMYYTO"
               CmpStr          =   "BOdYy <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   20
               Left            =   2505
               TabIndex        =   2
               ToolTipText     =   "Enter From Order Year"
               Top             =   1410
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               MaxLength       =   2
               DataField       =   "UOOMYYFR"
               IdName          =   "UOOMYYFR"
               CmpStr          =   "BOdYy >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   43
               Left            =   1725
               TabIndex        =   1
               ToolTipText     =   "Enter From Order Tc"
               Top             =   1410
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataField       =   "UOOMTCFR"
               IdName          =   "UOOMTCFR"
               CmpStr          =   "BOdTc >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   45
               Left            =   2925
               TabIndex        =   3
               ToolTipText     =   "Enter From Order Character"
               Top             =   1410
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOOMCHRFR"
               IdName          =   "UOOMCHRFR"
               CmpStr          =   "BOdChr >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   46
               Left            =   4275
               TabIndex        =   4
               ToolTipText     =   "Enter To Order Tc"
               Top             =   1410
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataField       =   "UOOMTCTO"
               IdName          =   "UOOMTCTO"
               CmpStr          =   "BOdTc <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   47
               Left            =   5475
               TabIndex        =   6
               ToolTipText     =   "Enter To Order Character"
               Top             =   1410
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOOMCHRTO"
               IdName          =   "UOOMCHRTO"
               CmpStr          =   "BOdChr <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   21
               Left            =   5175
               TabIndex        =   10
               ToolTipText     =   "Enter To Order Serial"
               Top             =   1695
               Width           =   525
               _ExtentX        =   926
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOODSRTO"
               IdName          =   "UOODSRTO"
               CmpStr          =   "BOdSr <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   23
               Left            =   2625
               TabIndex        =   8
               ToolTipText     =   "Enter From Order Serial"
               Top             =   1695
               Width           =   525
               _ExtentX        =   926
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOODSRFR"
               IdName          =   "UOODSRFR"
               CmpStr          =   "BOdSr >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   26
               Left            =   1725
               TabIndex        =   7
               ToolTipText     =   "Enter From Order Number"
               Top             =   1695
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOOMNOFR"
               IdName          =   "UOOMNOFR"
               CmpStr          =   "BOdNo >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   29
               Left            =   4275
               TabIndex        =   9
               ToolTipText     =   "Enter To Order Number"
               Top             =   1695
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOOMNOTO"
               IdName          =   "UOOMNOTO"
               CmpStr          =   "BOdNo <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   3
               Left            =   4275
               TabIndex        =   18
               ToolTipText     =   "Enter To Bag Opening Date"
               Top             =   2265
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   3
               DataField       =   "UODTTO"
               IdName          =   "UODTTO"
               CmpStr          =   "BOpnDt <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   5
               Left            =   1725
               TabIndex        =   17
               ToolTipText     =   "Enter From Bag Opening Date"
               Top             =   2265
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   3
               DataField       =   "UODTFR"
               IdName          =   "UODTFR"
               CmpStr          =   "BOpnDt >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   14
               Left            =   4275
               TabIndex        =   14
               ToolTipText     =   "Enter To Bag Year"
               Top             =   1980
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               MaxLength       =   2
               DataField       =   "UOBYYTO"
               IdName          =   "UOBYYTO"
               CmpStr          =   "BYy <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   17
               Left            =   1725
               TabIndex        =   11
               ToolTipText     =   "Enter From Bag Year"
               Top             =   1980
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               MaxLength       =   2
               DataField       =   "UOBYYFR"
               IdName          =   "UOBYYFR"
               CmpStr          =   "BYy >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   19
               Left            =   2145
               TabIndex        =   12
               ToolTipText     =   "Enter From Bag Character"
               Top             =   1980
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataField       =   "UOBCHRFR"
               IdName          =   "UOBCHRFR"
               CmpStr          =   "BChr >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   4
               Left            =   2925
               TabIndex        =   13
               ToolTipText     =   "Enter From Bag Number"
               Top             =   1980
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOBNOFR"
               IdName          =   "UOBNOFR"
               CmpStr          =   "BNo >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   6
               Left            =   4695
               TabIndex        =   15
               ToolTipText     =   "Enter To Bag Character"
               Top             =   1980
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataField       =   "UOBCHRTO"
               IdName          =   "UOBCHRTO"
               CmpStr          =   "BChr <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   22
               Left            =   5475
               TabIndex        =   16
               ToolTipText     =   "Enter To Bag Number"
               Top             =   1980
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOBNOTO"
               IdName          =   "UOBNOTO"
               CmpStr          =   "BNo <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   31
               Left            =   4275
               TabIndex        =   20
               ToolTipText     =   "Enter To Design Code"
               Top             =   2550
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
               Left            =   6840
               TabIndex        =   21
               ToolTipText     =   "Enter Design Code Selection"
               Top             =   2550
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
               Left            =   1725
               TabIndex        =   19
               ToolTipText     =   "Enter From Design Code"
               Top             =   2550
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
               Left            =   4275
               TabIndex        =   23
               ToolTipText     =   "Enter To Customer"
               Top             =   2835
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               DataField       =   "UOCMCDTO"
               IdName          =   "UOCMCDTO"
               CmpStr          =   "OmCmCd <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   41
               Left            =   6840
               TabIndex        =   24
               ToolTipText     =   "Enter Customer Selection"
               Top             =   2835
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataType        =   4
               DataField       =   "UOCMCDSEL"
               IdName          =   "UOCMCDSEL"
               CmpStr          =   "OmCmCd In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   42
               Left            =   1725
               TabIndex        =   22
               ToolTipText     =   "Enter From Customer"
               Top             =   2835
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               DataField       =   "UOCMCDFR"
               IdName          =   "UOCMCDFR"
               CmpStr          =   "OmCmCd >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   8
               Left            =   4275
               TabIndex        =   29
               ToolTipText     =   "Enter To Worker"
               Top             =   3405
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UOWRKTO"
               IdName          =   "UOWRKTO"
               CmpStr          =   "BWrk <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   9
               Left            =   6840
               TabIndex        =   30
               ToolTipText     =   "Enter Worker Selection"
               Top             =   3405
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOWRKSEL"
               IdName          =   "UOWRKSEL"
               CmpStr          =   "BWrk In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   10
               Left            =   1725
               TabIndex        =   28
               ToolTipText     =   "Enter From Worker"
               Top             =   3405
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UOWRKFR"
               IdName          =   "UOWRKFR"
               CmpStr          =   "BWrk >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   11
               Left            =   1725
               TabIndex        =   31
               ToolTipText     =   "Show 3D Picture Or Hand Sketch ? (Blank For No Picture)"
               Top             =   3690
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UOPMCDFR"
               IdName          =   "UOPMCDFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   12
               Left            =   1725
               TabIndex        =   0
               ToolTipText     =   "Enter Company Code"
               Top             =   1125
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               Enabled         =   0   'False
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOCOCDFR"
               IdName          =   "UOCOCDFR"
               CmpStr          =   "BCoCd = "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   58
               Left            =   6270
               TabIndex        =   33
               ToolTipText     =   "Enter 'Y' For Production Pointer, 'N' for Sales Pointer"
               Top             =   4545
               Visible         =   0   'False
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               Enabled         =   0   'False
               MaxLength       =   8
               DataType        =   3
               DataField       =   "UOYN1"
               IdName          =   "UOYN1"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   7
               Left            =   1725
               TabIndex        =   32
               ToolTipText     =   "Select Bag Report Format"
               Top             =   3975
               Width           =   1035
               _ExtentX        =   1826
               _ExtentY        =   503
               MaxLength       =   5
               DataType        =   4
               DataField       =   "UOPMCD"
               IdName          =   "UOPMCD"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   13
               Left            =   6270
               TabIndex        =   81
               ToolTipText     =   "Enter 'Y' For Production Pointer, 'N' for Sales Pointer"
               Top             =   4260
               Visible         =   0   'False
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               Enabled         =   0   'False
               MaxLength       =   8
               DataType        =   3
               DataField       =   "UOYN"
               IdName          =   "UOYN"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   15
               Left            =   3000
               TabIndex        =   36
               ToolTipText     =   "Show Box Detail (Yes / No)"
               Top             =   4830
               Visible         =   0   'False
               Width           =   315
               _ExtentX        =   556
               _ExtentY        =   503
               Enabled         =   0   'False
               DataType        =   4
               DataField       =   "UOYN2"
               IdName          =   "UOYN2"
            End
            Begin VB.Label LblShowWt 
               BackStyle       =   0  'Transparent
               Caption         =   "Show Rm Details "
               BeginProperty Font 
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
               Left            =   30
               TabIndex        =   80
               Top             =   4260
               Width           =   1755
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
               Index           =   11
               Left            =   60
               TabIndex        =   75
               Top             =   1125
               Width           =   1335
            End
            Begin VB.Label LblBgFmt 
               BackStyle       =   0  'Transparent
               Caption         =   "Bag Format"
               BeginProperty Font 
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
               TabIndex        =   74
               ToolTipText     =   "Location"
               Top             =   3975
               Width           =   1725
            End
            Begin VB.Label LblCalcPtr 
               BackStyle       =   0  'Transparent
               Caption         =   "Show Ptr"
               BeginProperty Font 
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
               Left            =   60
               TabIndex        =   73
               Top             =   4545
               Width           =   1755
            End
            Begin VB.Label LblPic 
               BackStyle       =   0  'Transparent
               Caption         =   "Pic (3D or HSk)"
               BeginProperty Font 
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
               TabIndex        =   72
               ToolTipText     =   "Location"
               Top             =   3690
               Width           =   1725
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Curr Worker"
               BeginProperty Font 
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
               Left            =   60
               TabIndex        =   71
               Top             =   3405
               Width           =   1425
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Bag Opn Date"
               BeginProperty Font 
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
               Left            =   60
               TabIndex        =   70
               Top             =   2265
               Width           =   1605
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Bag Yy/Chr/No"
               BeginProperty Font 
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
               Left            =   60
               TabIndex        =   69
               Top             =   1980
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
               Index           =   18
               Left            =   2085
               TabIndex        =   68
               Top             =   1980
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
               Index           =   19
               Left            =   4635
               TabIndex        =   67
               Top             =   1980
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
               Index           =   20
               Left            =   2865
               TabIndex        =   66
               Top             =   1980
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
               Index           =   21
               Left            =   5415
               TabIndex        =   65
               Top             =   1980
               Width           =   105
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
               Left            =   60
               TabIndex        =   64
               Top             =   2550
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
               Left            =   60
               TabIndex        =   63
               Top             =   2835
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
               Index           =   31
               Left            =   2565
               TabIndex        =   62
               Top             =   1695
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
               Index           =   30
               Left            =   5115
               TabIndex        =   61
               Top             =   1695
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
               Left            =   2445
               TabIndex        =   60
               Top             =   1410
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
               Left            =   4995
               TabIndex        =   59
               Top             =   1410
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
               Left            =   2865
               TabIndex        =   58
               Top             =   1410
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
               Left            =   5415
               TabIndex        =   57
               Top             =   1410
               Width           =   105
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
               Index           =   10
               Left            =   60
               TabIndex        =   56
               Top             =   1695
               Width           =   1485
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
               Index           =   9
               Left            =   60
               TabIndex        =   55
               Top             =   1410
               Width           =   1785
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Curr Location"
               BeginProperty Font 
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
               TabIndex        =   54
               Top             =   3120
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
               Left            =   6840
               TabIndex        =   53
               Top             =   480
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
               TabIndex        =   52
               Top             =   480
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
               Left            =   1725
               TabIndex        =   51
               Top             =   480
               Width           =   1335
            End
         End
      End
   End
End
Attribute VB_Name = "EmrFrpBagPrn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---- General Comments on the Report --------'
' This report gives all the details of a bag in a particular format. The output of this
' report is used as a sticker on the Bags. This report helps the Worker in knowing how much
' raw materials of a particular Rm Ctg has to be put in the bag. The report also guides the
' movement of the bags through the process locations as it specifies the sequence of
' locations. The report helps in identifying a bag (i.e. to which order or customer it belongs)
' It also gives specific instructions for production to the workers regarding the bag
' This report does not have any group options
' The report has a sub report which gives a Rm Category wise summary of the Rm Qty & Wt
'--------------------
Option Explicit
Dim Rep As New EmrRepBagPrn
Dim Rep2 As New EmrRepBagPrnFmt2
Dim Rep3 As New EmrRepBagPrnFmt3
Dim Rep5 As New EmrRepBagPrnFmt5  '*** Jay 2.13(CT) ***
Dim Rep4 As New EmrRepBagPrn4  '*** Jay 2.14 ***
Dim mRep As CRAXDRT.Report
Dim moCn As MwfLib.MDOConnection   '*** (Bef speed) bef 24/11/06

'****** Sachin 2.14.0 - [25-11-2006] - [Multi-Company] ******
Dim ws_MultiCoMod As Boolean
'****** Sachin 2.14.0 - [25-11-2006] - [Multi-Company] ******
Dim mFldName As String, mFontSz As Integer

Private Sub SetGroupSort()
End Sub
Private Sub adc_setreprecsource()
'*** (Jen 2.12)
'On Error GoTo RepErr
'*** (Jen 2.12)

'*** The Report Designer does not store the Sql written there, so it has to be written somewhere
'*** Report Sql (Used in the Report designer) ***
'  Select BYy, BChr, BNo, BOdTc, BOdYy, BOdChr, BOdNo, BOdSr, BOdDmCd, BOdSfx, BOdDmSz,
'  BOpnQty, BOpnPtQty, BOpnDt, BOpnDesc, BQty, BPtQty, OdDelDt, OdOrdQty, OdPrdQty,
'  OdDmPrdInst, OdCmPrdInst, OdCmStmpInst, OdSzInst, OdPrtCd, OdKt, OdDmCol, OdPicNm, OrRmCtg, OrRmSCtg,
'  OrRmCd, 0.0000 as OrLn1, 0.0000 as OrLn2, PValue3 as qPtrYN,
'  OrQty, OrWt, 0.0 as qMainMetWt, OrMainMet, OmCmCd , OmPoNo, CmName
'  From OrdRm, OrdDsg, OrdMst, Bag, CustMst, Param Where 1 = 2

'Manoj 2.11.0, Correction
'  Select BYy, BChr, BNo, BOdTc, BOdYy, BOdChr, BOdNo, BOdSr, BOdDmCd, BOdSfx, BOdDmSz,
'  BOpnQty, BOpnPtQty, BOpnDt, BOpnDesc, BQty, BPtQty, OdDelDt, OdOrdQty, OdPrdQty,
'  OdDmPrdInst, OdCmPrdInst, OdCmStmpInst, OdSzInst, OdPrtCd, OdKt, OdDmCol, OrSrNo, OrRmCtg, OrRmSCtg,
'  OrRmCd, 0.0000 as OrLn1, 0.0000 as OrLn2, OrRmPtr, PValue3 as qPtrYN,
'  OrQty, OrWt, 0.0 as qMainMetWt, OrMainMet, OrSetSCd, OrSubShp, OrHsQty, OrWsQty, OmCmCd , OmPoNo, CmName
'  From OrdRm, OrdDsg, OrdMst, Bag, CustMst, Param Where 1 = 2
'Manoj 2.11.0, Correction
  
  '*** Report Sql ***
  '*** The query for this report is
'  Select BYy, BChr, BNo, OdTc, OdYy, OdChr, OdNo, OdSr, OctCtChr, OctParts, OCtDmCol,
'(Select RmCtg from RmMst where RmCd= OctRmCd) as qRmCtg,  OctRmCd,
'OCtRmWt , OCtDc, OCtPrdSeq, OCtDesc
'From OrdCT
'Join OrdDsg On OdCoCd = OctCoCd And OdTc=OctTc and OdYy=OctYy and OdChr=OctChr and OdNo=OctNo and OdSr=OctSr
'Join Bag on BCoCd= OdCoCd and BOdTc= OdTc and BOdYy= OdYy and BOdChr= OdChr and BOdNo= OdNo and BOdSr= OdSr and OctCtChr= BChr
'Where 1 = 2

  
  ' **** Zubin 190205 **** '
  ' **** Manali 3.6.0 - 01/09/09 - BQty added below
  ' **** Sub-Report SQL
'  Select BYy, BChr, BNo, OrTc, OrYy, OrChr, OrNo, OrSr, OrSrNo,
'  OrRmCtg, OrRmSCtg, OrQty, OrWt From OrdRm, Bag where 1= 2
  ' **** Zubin 190205 **** '
  
  Dim wrepcnd As String, wCnd As String, wSqlStrg As String, ws_SubRep As String, ws_RmPtr As String
'Manoj 2.11.0, Correction
  Dim ws_OrLn1 As String
'Manoj 2.11.0, Correction

  Dim wAddlCnd As String            '****** Sachin 3.02.0 - New fields Condition
  
  Dim ws_SubRep1 As String     '*** Jay 2.14 ***
  Dim ws_SubRep2 As String     '*** Jay 2.13(CT) ***
  Dim wRowCnt As Integer
  Dim ws_RmDet As String, ws_ORmJoin As String, ws_TxndJoin As String, ws_OrdBy As String, wCustDsgCd As String
  
  Set Rep = Nothing
  
' Sachin 211 '
  Set Rep2 = Nothing
  Set Rep3 = Nothing
  
  '*** Jay 2.13(CT) ***
  Set Rep5 = Nothing
  '*** Jay 2.14 ***
  Set Rep4 = Nothing
  
'''''  Set adc.RepSource = Rep
'Manoj 2.11.0, Correction
  Select Case UCase(adc("UOPMCD"))
    Case UCase("BAG2"), UCase("BAG2A")
      Set adc.RepSource = Rep
      Set mRep = Rep
      If adc("UOPMCD") = "BAG2" Then wRowCnt = 28   '578-31 no of rows can print
      If adc("UOPMCD") = "BAG2A" Then wRowCnt = 15   '578-31 no of rows can print
    Case UCase("BAG6")
      Set adc.RepSource = Rep2
      Set mRep = Rep2
      wRowCnt = 11    '578-31 no of rows can print
    Case UCase("BAG3")
      Set adc.RepSource = Rep3
      Set mRep = Rep3
      wRowCnt = 22    '578-31 no of rows can print
    '*** Jay 2.14 ***
    ' ***** Manali 3.9.0 - BAG1 added (1 bag/page)
    Case UCase("BAG2B"), UCase("BAG1")
      Set adc.RepSource = Rep4
      Set mRep = Rep4
      If adc("UOPMCD") = "BAG2B" Then wRowCnt = 11    '578-31 no of rows can print
      If adc("UOPMCD") = "BAG1" Then wRowCnt = 45    '578-31 no of rows can print
    '*** Jay 2.14 ***
    '*** Jay 2.13(CT) ***
    Case UCase("BAGCT")
      Set adc.RepSource = Rep5
      Set mRep = Rep5
      wRowCnt = 12    '578-31 no of rows can print
    '*** Jay 2.13(CT) ***

'Dim Rep As New EmrRepBagPrn
'Dim Rep2 As New EmrRepBagPrnFmt2
'Dim Rep3 As New EmrRepBagPrnFmt3
'Dim Rep5 As New EmrRepBagPrnFmt5  '*** Jay 2.13(CT) ***
'Dim Rep4 As New EmrRepBagPrn4  '*** Jay 2.14 ***

'    Case UCase("BAG2A")
'      Set adc.RepSource = Rep2A
'      Set mRep = Rep2A
'Manoj 2.11.0, Correction
            
  End Select
' Sachin 211 '

  
  Call DispCoNm
  'Call GRP_REP.Gen3LRep
  'grpflds = GRP_REP.GrpFldLst
  'Manoj 2.10.0
  If adc("UoYN1") = "Y" Then
    ws_RmPtr = " (Case When OrPrdQty = 0 Then 0.0 Else (Case When OrRmCtg in ('C', 'D') Then OrPrdWt / (Case When OrPrdQty = 0 Then 1.0 Else OrPrdQty End) Else 0.0 End) End) "
  Else
  'Manoj 2.11.0, Correction
  '    ws_RmPtr = " OrRmPtr "
    ws_RmPtr = "(Case When P1.PValue2= 'Y' Then OrRmPtr Else (Case When OrQty = 0 Then 0.0 Else (Case When OrRmCtg in ('C', 'D') Then OrWt / (Case When OrQty = 0 Then 1.0 Else OrQty End) Else 0.0 End) End) End)"
  'Manoj 2.11.0, Correction
  End If
  'Manoj 2.10.0
  
 '4.1.4 - Needs to be checked thoroughly on every Operating system
 If Dir(CStr(Environ$("windir") & "\Fonts\Free3of9.TTF")) = "" Then _
    DispMsg "Barcode Font Not Installed. Cannot Continue.", etError: Exit Sub
  
  ' ***** Manali 3.6.0 - 26/09/09 - Changes done in query To Show Actual Wt (UoYN)
   Dim wRmJoin, wRmSumm As String
      If adc("UoYN") = "Y" Then
        wRmSumm = " Left Outer Join (Select TdPrtKey qTdPrtKey, TdCoCd qTdCoCd, TdBYy qTdBYy, TdBChr qTdBChr, TdBNo qTdBNo, Max(BGrWt) as qGrsWt, " + _
                  "     Sum((Case TdRmCtg When 'D' Then (TdRmWt * (Case When TdFrRmLoc='W' and TdFrRmDc='C' Then -1 " + _
                  "                                                     When TdFrRmLoc='W' and TdFrRmDc='D' Then 1 " + _
                  "                                                     When TdToRmLoc='W' and TdToRmDc='C' Then -1 " + _
                  "                                                     When TdToRmLoc='W' and TdToRmDc='D' Then 1 End)) Else 0 End)) as qDiaWt, " + _
                  "     Sum((Case TdRmCtg When 'C' Then (TdRmWt * (Case When TdFrRmLoc='W' and TdFrRmDc='C' Then -1 " + _
                  "                                                     When TdFrRmLoc='W' and TdFrRmDc='D' Then 1 " + _
                  "                                                     When TdToRmLoc='W' and TdToRmDc='C' Then -1 " + _
                  "                                                     When TdToRmLoc='W' and TdToRmDc='D' Then 1 End)) Else 0 End)) as qCsWt, " + _
                  "     Sum((Case TdRmCtg When OrRmCtg Then " + _
                  "                      (TdRmWt * (Case When TdFrRmLoc='W' and TdFrRmDc='C' Then -1 " + _
                  "                                      When TdFrRmLoc='W' and TdFrRmDc='D' Then 1 " + _
                  "                                      When TdToRmLoc='W' and TdToRmDc='C' Then -1 " + _
                  "                                      When TdToRmLoc='W' and TdToRmDc='D' Then 1 End)) Else 0 End)) as qMainMetWt " + _
                  "  From Txnd " + _
                  "     Join Bag On BCoCd=TdCoCd And BYy=TdBYy And BChr=TdBChr And BNo=TdBNo " + _
                  "     Join OrdRm On OrCoCd=BCoCd And OrTc=BOdTc And OrYy=BOdYy And OrChr=BOdChr And OrNo=BOdNo And OrSr=BOdSr And OrMainMet='Y' " + _
                  "  Group by TdPrtKey, TdCoCd, TdBYy, TdBChr, TdBNo) qTdRm " + _
                  "         On qTdPrtKey=BPrtKey and  qTdBYy=BYy and qTdBChr=BChr and qTdBNo=BNo "
        'kc.10- Stock by Ptr check is removed for OrLn1
        wRmJoin = " Left Outer Join (Select TdPrtKey, TdCoCd, TdBYy, TdBChr, TdBNo, Row_number() Over (order by TdPrtKey, TdCoCd, TdBYy, TdBChr, TdBNo, TdRmCtg, TdRmSCtg, TdRmCd, TdRmSz) as OrSrNo, TdRmCtg as OrRmCtg, TdRmSCtg as OrRmSCtg, TdRmCd as OrRmCd, " + _
                  " TdRmSz as OrLn1, 0.0 as OrLn2," + _
                  " (case when Max(PValue3)= 'Y' then  TdRmSz else 0.0 end) as OrRmPtr," + _
                  " Sum(TdRmWt * (Case When TdFrRmLoc='W' and TdFrRmDc='C' Then -1 When TdFrRmLoc='W' and TdFrRmDc='D' Then 1 " + _
                  "                   When TdToRmLoc='W' and TdToRmDc='C' Then -1 When TdToRmLoc='W' and TdToRmDc='D' Then 1 End)) as OrWt, " + _
                  " Sum(TdRmWt * (Case When TdFrRmLoc='W' and TdFrRmDc='C' Then -1 When TdFrRmLoc='W' and TdFrRmDc='D' Then 1 " + _
                  "                   When TdToRmLoc='W' and TdToRmDc='C' Then -1 When TdToRmLoc='W' and TdToRmDc='D' Then 1 End)) as OrPrdWt," + _
                  " Sum(TdRmQty * (Case When TdFrRmLoc='W' and TdFrRmDc='C' Then -1 When TdFrRmLoc='W' and TdFrRmDc='D' Then 1 " + _
                  "                   When TdToRmLoc='W' and TdToRmDc='C' Then -1 When TdToRmLoc='W' and TdToRmDc='D' Then 1 End)) as OrQty," + _
                  " Sum(TdRmQty * (Case When TdFrRmLoc='W' and TdFrRmDc='C' Then -1 When TdFrRmLoc='W' and TdFrRmDc='D' Then 1 " + _
                  "                   When TdToRmLoc='W' and TdToRmDc='C' Then -1 When TdToRmLoc='W' and TdToRmDc='D' Then 1 End)) as OrPrdQty," + _
                  " 0.0 as OrSalRt, " + _
                  " '' as OrMainMet,'' as OrSetSCd,'' as OrSubShp,0 as OrWsQty,0 as OrHsQty " + _
                  " From Txnd  " + _
                  " Join Param On PTyp='RMSCTG' And PMCd= TdRmCtg And PSCd = TdRmSCtg " + _
                  " Where TdSrNo<>0 Group By TdPrtKey, TdCoCd, TdBYy, TdBChr, TdBNo, TdRmCtg, TdRmSCtg, TdRmCd, TdRmSz) Txnd " + _
                  " On TdCoCd=BCoCd and TdBYy=BYy and TdBChr=BChr and TdBNo=BNo" + IIF(gs_Partition = ctCurrPrtn, " And TdPrtKey=BPrtKey ", "") + _
                  " and (Convert(Decimal(14,3), OrWt)>0.000 Or Convert(Decimal(14,1), OrQty)>0.0) "
                    
      Else
        wRmSumm = " Join (Select OrPrtKey qOrPrtKey,OrCoCd as qOrCoCd,OrTc as qOrTc,OrYy as qOrYy,OrChr as qOrChr,OrNo as qOrNo,OrSr as qOrSr, " + _
                      "Sum((Case OrRmCtg When 'D' Then OrPrdWt Else 0 End)) as qDiaWt, " + _
                      "Sum((Case OrRmCtg When 'C' Then OrPrdWt Else 0 End)) as qCsWt, " + _
                      "Sum(OrPrdWt/(Case When OrRmCtg in ('D','C') Then 5 Else 1 End)) as qGrsWt, " + _
                      "Sum((Case OrMainMet when 'Y' Then OrPrdWt Else 0 End)) as qMainMetWt " + _
                      "From OrdRm Group By OrPrtKey,OrCoCd,OrTc,OrYy,OrChr,OrNo,OrSr) OrSumm  " + _
                      " On qOrPrtKey=OdPrtKey and  qOrCoCd=BCoCd and qOrTc=BOdTc and qOrYy=BOdYy and qOrChr=BOdchr and qOrNo=BOdNo and qOrSr=BOdSr "
      
        wRmJoin = " Join Ordrm On  BOdIdNo=OrOdIdNo and OrCoCd = BCoCd And ortc=Bodtc and oryy=Bodyy and orchr=Bodchr and orno=Bodno and orsr=Bodsr " + IIF(gs_Partition = ctCurrPrtn, " And OrPrtKey=BPrtKey ", "")
        'wRmJoin = " Join Ordrm On  BOdIdNo=OrOdIdNo and OrCoCd = OdCoCd And ortc=odtc and oryy=odyy and orchr=odchr and orno=odno and orsr=odsr " + IIF(gs_Partition = ctCurrPrtn, " And OrPrtKey=OdPrtKey ", "")
      End If
      
      'If UCase(adc("UoYN7")) = "Y" Then
        wCustDsgCd = "IsNull((Select  cast(CdCustDmCd as Varchar(40)) from CustDsg " + _
                          "Where " + IIF(gs_Partition = ctCurrPrtn, " CdPrtKey='" + ctCurrPrtn + "' And  ", "") + " CdCtg = 'C' and CdCd = OdOmCmCd and " + _
                          "CdOurDmCd = OdDmCd and CdOurSfx = OdSfx and " + _
                          "CdOurDmSz = OdDmSz), " + _
                        "IsNull((Select cast(CdCustDmCd as Varchar(40)) from CustDsg " + _
                              "where " + IIF(gs_Partition = ctCurrPrtn, " CdPrtKey='" + ctCurrPrtn + "' and ", "") + " CdCtg = 'C' and CdCd = OdOmCmCd and " + _
                              "CdOurDmCd = OdDmCd and CdOurSfx = '' and " + _
                              "CdOurDmSz = OdDmSz), " + _
                            "IsNull((Select cast(CdCustDmCd as Varchar(40)) from CustDsg " + _
                                  "where " + IIF(gs_Partition = ctCurrPrtn, " CdPrtKey='" + ctCurrPrtn + "' and  ", "") + " CdCtg = 'C' and CdCd = OdOmCmCd and " + _
                                  "CdOurDmCd = OdDmCd and CdOurSfx = OdSfx and " + _
                                  "CdOurDmSz = '' ), " + _
                                "IsNull((Select cast(CdCustDmCd as Varchar(40)) from CustDsg " + _
                                      "where " + IIF(gs_Partition = ctCurrPrtn, " CdPrtKey='" + ctCurrPrtn + "' and  ", "") + " CdCtg = 'C' and CdCd = OdOmCmCd and " + _
                                      "CdOurDmCd = OdDmCd and CdOurSfx = '' and " + _
                                      "CdOurDmSz = ''), '{' + OdDmCd + '}')))) as CdCustDmCd"
'      Else
'        wOurDsgCd = "Max(IdDmCd) as CdOurDmCd"
'      End If
'
  
  
'  If adc("UoYN") = "Y" Then
'    ws_ORmJoin = " (Select OrPrtKey, OrCoCd, OrTc, OrYy, OrChr, OrNo, OrSr, OrRmCtg, OrRmSCtg, OrRmCd, Sum(OrPrdWt) as OrPrdWt, Sum(OrPrdQty) as OrPrdQty, " + _
'                 " Sum(Case When OrMainMet ='Y' Then OrPrdWt Else 0.0 End) as qMainMetWt, max(OrMainMet) as OrMainMet " + _
'                 " From OrdRm Group By OrPrtKey, OrCoCd, OrTc, OrYy, OrChr, OrNo, OrSr, OrRmCtg, OrRmSCtg, OrRmCd) "
'    ws_TxndJoin = " Left Outer Join (Select TdPrtKey, TdCoCd, TdBYy, TdBChr, TdBNo, TdRmCtg, TdRmSCtg, TdRmCd, TdRmSz, " + _
'                  "Sum(TdRmWt * (Case When TdFrRmLoc='W' and TdFrRmDc='C' Then -1 When TdFrRmLoc='W' and TdFrRmDc='D' Then 1 " + _
'                  "                   When TdToRmLoc='W' and TdToRmDc='C' Then -1 When TdToRmLoc='W' and TdToRmDc='D' Then 1 End)) as TdRmWt, " + _
'                  "Sum(TdRmQty * (Case When TdFrRmLoc='W' and TdFrRmDc='C' Then -1 When TdFrRmLoc='W' and TdFrRmDc='D' Then 1 " + _
'                  "                   When TdToRmLoc='W' and TdToRmDc='C' Then -1 When TdToRmLoc='W' and TdToRmDc='D' Then 1 End)) as TdRmQty " + _
'                  "From Txnd Where TdSrNo<>0 Group By TdPrtKey, TdCoCd, TdBYy, TdBChr, TdBNo, TdRmCtg, TdRmSCtg, TdRmCd, TdRmSz) Txnd " + _
'                  "On TdCoCd=BCoCd and TdBYy=BYy and TdBChr=BChr and TdBNo=BNo and TdRmCd=OrRmCd " + IIF(gs_Partition = ctCurrPrtn, " And TdPrtKey=OrPrtKey ", "") + _
'                  " and (Convert(Decimal(14,3), TdRmWt)>0.000 Or Convert(Decimal(14,1), TdRmQty)>0.0) "
'    ws_OrdBy = " Order By BCoCd, BYy, BChr, BNo, (Case OrRmCtg When 'G' Then 1 When 'P' Then 2 When 'S' Then 3 " + _
'              "When 'L' Then 4 When 'D' Then 5 When 'C' Then 6 When 'X' Then 7 When 'A' Then 8 When 'M' Then 9 Else 10 End) "
'
'  Else
'    ws_ORmJoin = "": ws_TxndJoin = "": ws_OrdBy = ""
'  End If
  ' ***** Manali 3.6.0 - 26/09/09 - Changes done in query To Show Actual Wt (UoYN)
  
  '****** Sachin 3.02.0 - New Fields Condition
  If Trim(adc("UOCMCDFR")) <> "" Then wAddlCnd = " OdOmCmCd >= '" + adc("UOCMCDFR") + "'"
  If Trim(adc("UOCMCDTO")) <> "" Then wAddlCnd = IIF(wAddlCnd <> "", wAddlCnd + " And ", "") + " OdOmCmCd <= '" + adc("UOCMCDTO") + "'"
  If Trim(adc("UOCMCDSEL")) <> "" Then wAddlCnd = IIF(wAddlCnd <> "", wAddlCnd + " And ", "") + " OdOmCmCd In (" + Trim(adc("UOCMCDSEL")) + ")"
  
  '******Sachin*****************************************Emr211-07/06/05*******
    If UCase(adc("UOPMCD")) = "BAG3" Or UCase(adc("UOPMCD")) = "BAG6" Then
      wrepcnd = adc.RepCond
      wCnd = IIF(wrepcnd <> "", " where " + wrepcnd, "")
      
'      ' ***** Manali 3.6.0 - 26/09/09 - Changes done in query To Show Actual Wt (UoYN)
' bef 3.8.0
'      If adc("UoYN") = "Y" Then
'        ws_RmDet = "orrmctg, IsNull(TdRmCd, '') as orrmcd, P1.PValue3 as PtrYN, " + _
'                  " (case when P1.PValue3= 'N' then IsNull(TdRmSz, 0.00) else 0.00 end) as OrLn1, " + _
'                  " 0.00 as orln2, (case when P1.PValue3= 'Y' then IsNull(TdRmSz, 0.00) else 0.00 end) as OrRmPtr, " + _
'                  " IsNull(TdRmQty, 0) as orqty, IsNull(TdRmWt, 0.00) as orwt, OrPrdWt as qOrWt, 0.0 as orsalrt,"
'      Else
'        ' *****  Order Rm Details as earlier
'        ws_RmDet = "orrmctg,orrmcd, P1.PValue3 as PtrYN, " + _
'                  " (case when P1.PValue3= 'N' then orln1 else 0 end) as OrLn1, " + _
'                  " orln2, " + ws_RmPtr + " as OrRmPtr, orqty, OrPrdWt as orwt, OrPrdWt as qOrWt, orsalrt,"
'      End If
'      ' ***** Manali 3.6.0 - 26/09/09 - Changes done in query To Show Actual Wt (UoYN)
      
      
'Manoj 2.11.0, Correction
'      wSqlStrg = " Select BYy, BChr, BNo, BagNo=(cast(byy as varchar)+'/'+bchr+'/'+ cast(bno as varchar)), " + _
'                 " boddmcd,bodsfx,boddmsz,bopnqty,bopnptqty,bopndt,bqty, " + _
'                 " OrderNo=(odtc+'/'+cast(odyy as varchar)+'/'+odchr+'/'+cast(odno as varchar)+'/'+cast(odsr as varchar)), " + _
'                 " omcmcd,omdt,ompono,odkt,oddmcol,odcmstmpinst, " + _
'                 " oddmprdinst,odcmprdinst,odszinst,odprtcd,orrmctg,orrmcd, P1.PValue3 as PtrYN, " + _
'                 " (case when P1.PValue3= 'N' then orln1 else 0 end) as OrLn1, " + _
'                 " orln2, (case when P1.PValue3= 'Y' then " + ws_RmPtr + " else 0 end) as OrRmPtr, " + _
'                 " orqty,orwt,orsalrt,dmctg " + _
'                 " From Bag " + _
'                 "      Join Ordmst On OmCoCd = BCoCd And Omtc=BOdTc and Omyy=BOdYy and Omchr=Bodchr and Omno=Bodno " + _
'                 "      Join Orddsg On OdCoCd = BCoCd And BOdTc=OdTc and BOdYy=OdYy and BOdChr=OdChr and BOdNo=OdNo and BOdSr=OdSr " + _
'                 "      Join Ordrm On  OrCoCd = OdCoCd And ortc=odtc and oryy=odyy and orchr=odchr and orno=odno and orsr=odsr " + _
'                 "      Join Dsgmst On  OdDmCd=DmCd And DmTcTyp='DM' And DmSz=''" + _
'                 "      Join Param P1 On P1.PTyp='RMSCTG' And P1.PMCd= OrRmCtg And P1.PSCd = OrRmSCtg " + wCnd
      ' Zubin 211 (Patch 03 OmDt changed to OdDelDt [Prd Del Dt])
      ' *** JAY 2.14NEXT *** [BagNo Sorting for Bag Format 3]
      '*** Jay 3.2.0 [PicNm]
      ' ***** Manali 3.6.0 - 26/09/09 - Changes done in query To Show Actual Wt (UoYN)
      'kc.10- Stock by Ptr check is removed for OrLn1
      'indigo.164 DmTcTyp= 'DM' changed as OdDmTcTyp=DmTcTyp in DsgMst join & DsgmstForWax join
      'sv.77 BOdTc added
      wSqlStrg = " Select BYy, BChr, BNo, BYy+ BChr+ space(3-Len(BChr)) + Str(BNo) as BagNo, " + _
                 " qBagNo=(cast(byy as varchar)+'/'+bchr+'/'+ cast(bno as varchar))," + wCustDsgCd + ", " + _
                 " boddmcd,bodsfx,boddmsz,bopnqty,bptqty,bopndt,bqty, " + _
                 " OrderNo=(odtc+'/'+cast(odyy as varchar)+'/'+odchr+'/'+cast(odno as varchar)+'/'+cast(odsr as varchar)), " + _
                 " omcmcd, OdDelDt,ompono,odkt,oddmcol, OdPicNm, odcmstmpinst, " + _
                 " oddmprdinst,odcmprdinst,odszinst,odprtcd, odpono, " + _
                  "IsNull(orrmctg, '') as orrmctg, IsNull(orrmcd, '') as orrmcd,  IsNull(P1.PValue3, '') as PtrYN, " + _
                  " IsNull(OrLn1, 0.0) as OrLn1, " + _
                  " IsNull(orln2, 0.0) as OrLn2, IsNull(" + ws_RmPtr + ", 0.0) as OrRmPtr, IsNull(orqty, 0) as orqty, IsNull(OrPrdWt, 0.0) as orwt, IsNull(OrPrdWt, 0.0) as qOrWt, IsNull(orsalrt, 0.0) as OrSalRt, " + _
                  " qDiaWt, qCsWt, qGrsWt,isnull(DsgmstForWax.DmWaxWt,DsgMst.DmWaxWt) as DmWaxWt, qMainMetWt, DsgMst.dmctg, DsgMst.dmprdctg, OdPoNo,BOdTc " + _
                 " From Bag " + _
                 "      Join Ordmst On OmIdNo=BOmIdNo and OmCoCd = BCoCd And Omtc=BOdTc and Omyy=BOdYy and Omchr=Bodchr and Omno=Bodno " + IIF(gs_Partition = ctCurrPrtn, " And OmPrtKey=BPrtKey ", "") + _
                 "      Join Orddsg On OdIdNo=BOdIdNo  and OdCoCd = BCoCd And BOdTc=OdTc and BOdYy=OdYy and BOdChr=OdChr and BOdNo=OdNo and BOdSr=OdSr " + IIF(gs_Partition = ctCurrPrtn, " And BPrtKey=OdPrtKey ", "") + _
                 wRmSumm + _
                 "      Join Dsgmst On OdDmIdNo=DmIdNo and OdDmCd=DmCd and OdDmTcTyp=DmTcTyp And DsgMst.DmSz='' " + IIF(gs_Partition = ctCurrPrtn, " And DsgMst.DmPrtKey=OdPrtKey ", "") + _
                " Left Outer Join DsgMst DsgmstForWax On OdDmCd=DsgmstForWax.DmCd and OdDmTcTyp=DsgmstForWax.DmTcTyp And OdDmSz=DsgmstForWax.DmSz " + IIF(gs_Partition = ctCurrPrtn, " And DsgMstForWax.DmPrtKey=OdPrtKey ", "") + _
                 wRmJoin + _
                 IIF(adc("UoYN") = "Y", " Left Outer ", "") + " Join Param P1 On P1.PTyp='RMSCTG' And P1.PMCd= OrRmCtg And P1.PSCd = OrRmSCtg " + _
                 wCnd + IIF(wAddlCnd <> "", IIF(wCnd <> "", " And ", " Where ") + wAddlCnd, "") + _
                 IIF(gs_Partition = ctCurrPrtn, " And BPrtKey='" + ctCurrPrtn + "' ", "") + ws_OrdBy
      
''      wSqlStrg = " Select BYy, BChr, BNo, BYy+ BChr+ space(3-Len(BChr)) + Str(BNo) as BagNo, " + _
''                 " qBagNo=(cast(byy as varchar)+'/'+bchr+'/'+ cast(bno as varchar)), " + _
''                 " boddmcd,bodsfx,boddmsz,bopnqty,bptqty,bopndt,bqty, " + _
''                 " OrderNo=(odtc+'/'+cast(odyy as varchar)+'/'+odchr+'/'+cast(odno as varchar)+'/'+cast(odsr as varchar)), " + _
''                 " omcmcd, OdDelDt,ompono,odkt,oddmcol, OdPicNm, odcmstmpinst, " + _
''                 " oddmprdinst,odcmprdinst,odszinst,odprtcd, " + ws_RmDet + " dmctg " + _
''                 " From Bag " + _
''                 "      Join Ordmst On OmIdNo=BOmIdNo and OmCoCd = BCoCd And Omtc=BOdTc and Omyy=BOdYy and Omchr=Bodchr and Omno=Bodno " + IIF(gs_Partition = ctCurrPrtn, " And OmPrtKey=BPrtKey ", "") + _
''                 "      Join Orddsg On OdIdNo=BOdIdNo  and OdCoCd = BCoCd And BOdTc=OdTc and BOdYy=OdYy and BOdChr=OdChr and BOdNo=OdNo and BOdSr=OdSr " + IIF(gs_Partition = ctCurrPrtn, " And BPrtKey=OdPrtKey ", "") + _
''                 "      Join " + ws_ORmJoin + " Ordrm On " + IIF(adc("UoYN") = "Y", "", "OdIdNo=OrOdIdNo and ") + " OrCoCd = OdCoCd And ortc=odtc and oryy=odyy and orchr=odchr and orno=odno and orsr=odsr " + IIF(gs_Partition = ctCurrPrtn, " And OrPrtKey=OdPrtKey ", "") + _
''                 "      Join Dsgmst On OdDmIdNo=DmIdNo and OdDmCd=DmCd And DmTcTyp='DM' And DmSz='' " + IIF(gs_Partition = ctCurrPrtn, " And DmPrtKey=OdPrtKey ", "") + _
''                 "      Join Param P1 On P1.PTyp='RMSCTG' And P1.PMCd= OrRmCtg And P1.PSCd = OrRmSCtg " + ws_TxndJoin + _
''                 wCnd + IIF(wAddlCnd <> "", IIF(wCnd <> "", " And ", " Where ") + wAddlCnd, "") + _
''                 IIF(gs_Partition = ctCurrPrtn, " And BPrtKey='" + ctCurrPrtn + "' ", "") + ws_OrdBy
      ' *** JAY 2.14NEXT *** [BagNo Sorting for Bag Format 3]
      'Manoj 2.11.0, Correction
      adc.RepRecSource = wSqlStrg + ctMaxDopOpt '****** Sachin 3.01 ctMaxDopOpt added
      '******Sachin*****************************************Emr211-07/06/05*******
    Else
      wrepcnd = adc.RepCond
      wCnd = IIF(wrepcnd <> "", " where " + wrepcnd, "")
      '******Geeta*****************Emr206****05/04/04*******
      '**********If UserLevel>6 then don't display the Customer Name
      'Replaced OdDmPrdInst,OdCmPrdInst  with Replace( ,char(13)+ char(10),' ')
      
      'Manoj 2.11.0, Correction
      If UCase(adc("UOPMCD")) = "BAG2" Then
        ws_OrLn1 = " IsNull((case when P1.PValue3= 'Y' then " + ws_RmPtr + " else OrLn1 end), 0) as OrLn1, "
      '*** Jay 2.13(CT) *** [Introduced UCase(adc("UOPMCD")) = "BAGCT"]
      '*** Jay 2.14 *** [Introduced UCase(adc("UOPMCD")) = "BAG2B"]
      ' ***** Manali 3.9.0 - BAG1 added (1 bag/page)
      ElseIf UCase(adc("UOPMCD")) = "BAG2A" Or UCase(adc("UOPMCD")) = "BAG2B" Or UCase(adc("UOPMCD")) = UCase("BAG1") Or UCase(adc("UOPMCD")) = ("BAGCT") Then
        ws_OrLn1 = " IsNull(OrLn1, 0) as OrLn1, "
      End If
      '*** Jay 2.14 ***
      '*** Jay 2.13(CT) ***
      'Manoj 2.11.0, Correction
            
       ' ***** Manali 3.6.0 - 26/09/09 - Changes done in query To Show Actual Wt (UoYN)
       ' Bef 3.8.0
'      If adc("UoYN") = "Y" Then
'          ws_RmDet = "  Row_number() Over (order by TdPrtKey, TdCoCd, TdBYy, TdBChr, TdBNo, TdRmCtg, TdRmSCtg, TdRmCd, TdRmSz) OrSrNo, " + _
'                  " OrRmCtg, OrRmSCtg, IsNull(TdRmcd, '') as OrRmCd, IsNull(TdRmSz, 0.0) as OrLn1, " + _
'                  "0.00 as OrLn2, 0.00 as OrRmPtr, P1.PValue3 as qPtrYN, IsNull(TdRmQty, 0) as OrQty, IsNull(TdRmWt, 0.0) as OrWt, " + _
'                  "qMainMetWt, OrMainMet, OrPrdWt as qOrWt, '' as OrSetSCd, '' as OrSubShp, 0 as OrWsQty, 0 as OrHsQty, "
'
'      Else
'          ' *****  Order Rm Details as earlier
'          ws_RmDet = " OrSrNo, OrRmCtg, OrRmSCtg, OrRmCd, " + ws_OrLn1 + _
'                  " OrLn2, " + ws_RmPtr + "as OrRmPtr, P1.PValue3 as qPtrYN, OrPrdQty as OrQty, OrPrdWt as OrWt, " + _
'                  "/*(case when OrMainMet= 'Y' then OrPrdWt else 0 end) as*/ qMainMetWt, OrMainMet, OrPrdWt as qOrWt, " + _
'                  "OrSetSCd, OrSubShp, OrWsQty, OrHsQty, "
'      End If
      ' ***** Manali 3.6.0 - 26/09/09 - Changes done in query To Show Actual Wt (UoYN)
            
            
      ' Zubin 212 (OrSubShp was missing)
      '****** Sachin 3.02 - Id fields added in Joins
      '*** Jay 3.2.0 [PicNm]
      ' ***** Manali 350Nxt - Join with DsgMst and DmCtg added
      ' ***** Manali 3.6.0 - 26/09/09 - Changes done in query To Show Actual Wt (UoYN)
      'indigo.164 DmTcTyp= 'DM' changed as OdDmTcTyp=DmTcTyp in DsgMst join & DsgmstForWax join
      wSqlStrg = "Select BYy, BChr, BNo, BOdTc, BOdYy, BOdChr, BOdNo, BOdSr," + wCustDsgCd + ", BOdDmCd, BOdSfx, " + _
                 "BOdDmSz, BOpnQty, BOpnPtQty, BOpnDt, BOpnDesc, BQty, BPtQty, OdDelDt, " + _
                 "OdOrdQty, OdPrdQty, Replace(OdDmPrdInst,char(13)+ char(10),' ') as OdDmPrdInst," + _
                 "Replace( OdCmPrdInst,char(13)+ char(10),' ') as OdCmPrdInst, OdCmStmpInst, OdSzInst, OdPrtCd, OdKt, " + _
                 "OdDmCol, OdPicNm, " + _
                 " IsNull(OrSrNo, 1) as OrSrNo, IsNull(OrRmCtg, '') as OrRmCtg, IsNull(OrRmSCtg, '') as OrRmSCtg, IsNull(OrRmCd, '') as OrRmCd, " + ws_OrLn1 + _
                 " IsNull(OrLn2, 0.0) as OrLn2, IsNull(" + ws_RmPtr + ", 0.0) as OrRmPtr, IsNull(P1.PValue3, '') as qPtrYN, Isnull(OrPrdQty, 0.0) as OrQty, IsNull(OrPrdWt, 0.0) as OrWt, " + _
                 " qMainMetWt, IsNull(OrMainMet, '') as OrMainMet, Isnull(OrPrdWt, 0) as qOrWt, " + _
                 "IsNull(OrSetSCd, '') as OrSetSCd, IsNUll(OrSubShp, '') as OrSubShp, IsNull(OrWsQty, 0) as OrWsQty, IsNull(OrHsQty, 0) as OrHsQty, " + _
                 "qDiaWt, qCSWt, qGrsWt, isnull(DsgmstForWax.DmWaxWt,DsgMst.DmWaxWt) as DmWaxWt,  " + _
                 " OmCmCd, OmPoNo, (Select (case when (Select PNum from Param where Ptyp = 'USR' and " + _
                 "PmCd = '" + Trim(gs_UsrCd) + "')  <= 5 then  CmName else CmCd end) " + _
                 "from CustMst where CmCtg= 'C' and CmCd= OmCmCd) as CmName, DsgMst.DmCtg, DsgMst.dmprdctg, OdPoNo " + _
                 "From OrdDsg Join OrdMst on OmIdNo=OdOmIdNo and OmCoCd= OdCoCd and OmTc= OdTc and OmYy= OdYy and OmChr= OdChr and OmNo= OdNo " + IIF(gs_Partition = ctCurrPrtn, " And OmPrtKey=OdPrtKey ", "") + _
                 "Join Bag On BOdIdNo=OdIdNo and BCoCd = OdCoCd and BOdTc= OdTc and BOdYy= OdYy and BOdChr= OdChr " + _
                 "and BOdNo= OdNo and BOdSr= OdSr " + IIF(gs_Partition = ctCurrPrtn, " And BPrtKey=OdPrtKey ", "") + _
                 wRmSumm + _
                 " Join Dsgmst On OdDmIdNo=DsgMst.DmIdNo and OdDmCd=DsgMst.DmCd and OdDmTcTyp=DsgMst.DmTcTyp And DsgMst.DmSz='' " + IIF(gs_Partition = ctCurrPrtn, " And DsgMst.DmPrtKey=OdPrtKey ", "") + _
                 " Left Outer Join DsgMst DsgmstForWax On OdDmCd=DsgmstForWax.DmCd and OdDmTcTyp=DsgmstForWax.DmTcTyp And OdDmSz=DsgmstForWax.DmSz " + IIF(gs_Partition = ctCurrPrtn, " And DsgMstForWax.DmPrtKey=OdPrtKey ", "") + _
                 IIF(UCase(adc("UOPMCD")) = ("BAGCT"), _
                 " Join Param P2 on P2.PTyp= 'BCHR' and P2.PMCd= BChr and P2.PSCd= '' and P2.PValue= 'C' ", " ") + _
                 wRmJoin + _
                 IIF(adc("UoYN") = "Y", " Left Outer ", "") + " Join RmMst on RmCd= OrRmCd " + IIF(gs_Partition = ctCurrPrtn, " And RmPrtKey=OdPrtKey ", "") + _
                 IIF(adc("UoYN") = "Y", " Left Outer ", "") + " Join Param P1 on P1.PTyp= 'RMSCTG' and P1.PMCd= RmCtg and P1.PSCd= RmSCtg " + ws_TxndJoin + _
                wCnd + IIF(wAddlCnd <> "", IIF(wCnd <> "", " And ", " Where ") + wAddlCnd, "") + IIF(gs_Partition = ctCurrPrtn, " And BPrtKey='" + ctCurrPrtn + "' ", "") + ws_OrdBy
      
       adc.RepRecSource = wSqlStrg + ctMaxDopOpt '****** Sachin 3.01 ctMaxDopOpt added
      
'       Bef  3.8.0
'      wSqlStrg = "Select BYy, BChr, BNo, BOdTc, BOdYy, BOdChr, BOdNo, BOdSr, BOdDmCd, BOdSfx, " + _
'                 "BOdDmSz, BOpnQty, BOpnPtQty, BOpnDt, BOpnDesc, BQty, BPtQty, OdDelDt, " + _
'                 "OdOrdQty, OdPrdQty, Replace(OdDmPrdInst,char(13)+ char(10),' ') as OdDmPrdInst," + _
'                 "Replace( OdCmPrdInst,char(13)+ char(10),' ') as OdCmPrdInst, OdCmStmpInst, OdSzInst, OdPrtCd, OdKt, " + _
'                 "OdDmCol, OdPicNm, " + ws_RmDet + " OmCmCd, OmPoNo, " + _
'                 "(Select (case when (Select PNum from Param where Ptyp = 'USR' and " + _
'                 "PmCd = '" + Trim(gs_UsrCd) + "')  <= 5 then  CmName else CmCd end) " + _
'                 "from CustMst where CmCtg= 'C' and CmCd= OmCmCd) as CmName, DmCtg " + _
'                 "From " + ws_ORmJoin + " OrdRm " + _
'                 "Join OrdDsg On " + IIF(adc("UoYN") = "Y", "", "OrOdIdNo=OdIdNo and ") + " OrCoCd= OdCoCd and OrTc= OdTc and OrYy= OdYy " + _
'                 "and OrChr= OdChr and OrNo= OdNo and OrSr= OdSr " + IIF(gs_Partition = ctCurrPrtn, " And OrPrtKey=OdPrtKey ", "") + _
'                 "Join OrdMst on OmIdNo=OdOmIdNo and OmCoCd= OdCoCd and OmTc= OdTc and OmYy= OdYy and OmChr= OdChr and OmNo= OdNo " + IIF(gs_Partition = ctCurrPrtn, " And OmPrtKey=OdPrtKey ", "") + _
'                 "Join Bag On BOdIdNo=OdIdNo and BCoCd = OdCoCd and BOdTc= OdTc and BOdYy= OdYy and BOdChr= OdChr " + _
'                 "and BOdNo= OdNo and BOdSr= OdSr " + IIF(gs_Partition = ctCurrPrtn, " And BPrtKey=OdPrtKey ", "") + _
'                 " Join Dsgmst On OdDmIdNo=DmIdNo and OdDmCd=DmCd And DmTcTyp='DM' And DmSz='' " + IIF(gs_Partition = ctCurrPrtn, " And DmPrtKey=OdPrtKey ", "") + _
'                 "Join RmMst on RmCd= OrRmCd " + IIF(gs_Partition = ctCurrPrtn, " And RmPrtKey=OrPrtKey ", "") + _
'                 "Join Param P1 on P1.PTyp= 'RMSCTG' and P1.PMCd= RmCtg and P1.PSCd= RmSCtg " + ws_TxndJoin + _
'                 IIF(UCase(adc("UOPMCD")) = ("BAGCT"), _
'                 " Join Param P2 on P2.PTyp= 'BCHR' and P2.PMCd= BChr and P2.PSCd= '' and P2.PValue= 'C' ", " ") + wCnd + IIF(wAddlCnd <> "", IIF(wCnd <> "", " And ", " Where ") + wAddlCnd, "") + IIF(gs_Partition = ctCurrPrtn, " And BPrtKey='" + ctCurrPrtn + "' ", "") + ws_OrdBy
      
       '****** Sachin 3.02 - Id fields added in Joins
       ' **** Manali 3.6.0 - 01/09/09 - BQty added below
       ' ***** Manali 3.6.0 - 26/09/09 - Changes done in query To Show Actual Wt (UoYN)
       ' ***** Manali 3.8.0 - Query changed
          ws_SubRep = "Select BYy, BChr, BNo, BQty, OmTc as OrTc, OmYy as OrYy, OmChr as OrChr, OmNo as OrNo, BOdSr as OrSr, IsNull(OrSrNo, 0) as OrSrno, IsNull(OrRmCtg, '')as OrRmCtg, " + _
                      "IsNull(OrRmSCtg, '') as OrRmSCtg, IsNUll(OrPrdQty, 0) as OrQty, IsNull(OrPrdWt, 0.0) as OrWt From Bag " + _
                      "Join OrdMst on OmCoCd= BCoCd and OmTc= BOdTc and OmYy= BODYy and OmChr= BOdChr and OmNo= BOdNo " + IIF(gs_Partition = ctCurrPrtn, " And OmPrtKey=BPrtKey And OmPrtKey='" + ctCurrPrtn + "' ", "") + _
                      wRmJoin + wCnd
       
'       If adc("UoYN") = "Y" Then
'          ws_SubRep = "Select BYy, BChr, BNo, BQty, OrTc, OrYy, OrChr, OrNo, OrSr, 0 as OrSrNo, IsNull(TdRmCtg, '') as OrRmCtg, " + _
'                      "IsNull(TdRmSCtg, '') as OrRmSCtg, TdRmQty as OrQty, TdRmWt as OrWt From " + ws_ORmJoin + " OrdRm Join Bag on BCoCd= OrCoCd and BOdTc= OrTc and " + _
'                      "BOdYy= OrYy and BOdChr= OrChr and BOdNo= OrNo and BOdSr= OrSr " + IIF(gs_Partition = ctCurrPrtn, " And BPrtKey=OrPrtKey And BPrtKey='" + ctCurrPrtn + "' ", "") + " Join OrdMst on " + _
'                      "OmCoCd= OrCoCd and OmTc= OrTc and OmYy= OrYy and OmChr= OrChr and OmNo= OrNo " + IIF(gs_Partition = ctCurrPrtn, " And OmPrtKey=OrPrtKey And OmPrtKey='" + ctCurrPrtn + "' ", "") + ws_TxndJoin + wCnd
'       Else
'          ws_SubRep = "Select BYy, BChr, BNo, BQty, OrTc, OrYy, OrChr, OrNo, OrSr, OrSrNo, OrRmCtg, " + _
'                      "OrRmSCtg, OrPrdQty as OrQty, OrPrdWt as OrWt From OrdRm Join Bag on BOdIdNo=OrOdIdNo and BCoCd= OrCoCd and BOdTc= OrTc and " + _
'                      "BOdYy= OrYy and BOdChr= OrChr and BOdNo= OrNo and BOdSr= OrSr " + IIF(gs_Partition = ctCurrPrtn, " And BPrtKey=OrPrtKey And BPrtKey='" + ctCurrPrtn + "' ", "") + " Join OrdMst on " + _
'                      "OmCoCd= OrCoCd and OmTc= OrTc and OmYy= OrYy and OmChr= OrChr and OmNo= OrNo " + IIF(gs_Partition = ctCurrPrtn, " And OmPrtKey=OrPrtKey And OmPrtKey='" + ctCurrPrtn + "' ", "") + wCnd
'       End If
        '*** Jay 2.14 ***
        '*** Jay 2.13(CT) ***
        
        Call moCn.RepRes(mRep.Subreport1.OpenSubreport, ws_SubRep + ctMaxDopOpt)  '****** Sachin 3.01 ctMaxDopOpt added
        If UCase(adc("UOPMCD")) = ("BAGCT") Then
          'Call moCn.RepRes(Rep5.Subreport1.OpenSubreport, ws_SubRep)
          '*** Jay 2.14Next *** [Multiple Comp Char]
          '****** Sachin 3.02 - New fields in Where Clause
          ws_SubRep2 = "Select BYy, BChr, BNo, OdTc, OdYy, OdChr, OdNo, OdSr, OctSrNo, OctCd, OctSz, OctCtChr, OctParts, " + _
                  "(Select RmCtg from RmMst Where " + IIF(gs_Partition = ctCurrPrtn, " RmPrtKey='" + ctCurrPrtn + "' and ", "") + " RmCd= OctRmCd) as qRmCtg, OctRmCd, OctRmWt, OctDmCol, OctDc, OctPrdSeq, OctMainPrtYn, OctDesc " + _
                  "From OrdCT " + _
                  "Join OrdDsg on OdIdNo=OctOdIdNo and OdCoCd= OctCoCd and OdTc= OctTc and OdYy= OctYy and OdChr= OctChr and OdNo= OctNo and " + _
                  "OdSr= OctSr " + IIF(gs_Partition = ctCurrPrtn, " And OdPrtKey=OctPrtKey ", "") + _
                  "Join Bag on " + IIF(gs_Partition = ctCurrPrtn, " BPrtKey=OdPrtKey and ", "") + " BOdIdNo=OdIdNo and BCoCd= OdCoCd and BOdTc= OdTc and  BOdYy= OdYy and BOdChr= OdChr and BOdNo= OdNo and " + _
                  "BOdSr= OdSr and BChr= OctCtChr " + wCnd + IIF(wAddlCnd <> "", IIF(wCnd <> "", " And ", " Where  ") + wAddlCnd + IIF(gs_Partition = ctCurrPrtn, " and BPrtKey='" + ctCurrPrtn + "' ", ""), "") + " " + _
                  "Order by OdTc, OdYy, OdChr, OdNo, OdSr, OctSrNo "
          '*** Jay 2.14Next *** [Multiple Comp Char]
          Call moCn.RepRes(Rep5.Subreport2.OpenSubreport, ws_SubRep2 + ctMaxDopOpt) '****** Sachin 3.01 ctMaxDopOpt added
        '****** Sachin 3.02 - Id fields in Where Clause
        ' ***** Manali 3.9.0 - BAG1 added (1 bag/page), OdDmCd, OdDmSz, OdPicNm added
        ElseIf UCase(adc("UOPMCD")) = ("BAG2B") Or UCase(adc("UOPMCD")) = UCase(("BAG1")) Then
          ws_SubRep1 = " Select LPos, BOdTc, BOdYy, BOdChr, BOdNo, BOdSr, BYy, BChr, BNo, PrdSeq, OdDmCd, OdDmSz, OdPicNm " + _
                     " From OrdDsg " + _
                     "   Join (Select LocCoCd, vPMCd, CHARINDEX(LocCd,','+vpdesc225+',') LPos, LocCd As PrdSeq " + _
                     "         From Loc " + _
                     "         Join vParam On vPCoCd=LocCoCd And vptyp='PRDSEQ' " + _
                     "         Where Loctyp='P' And ','+vPDesc225+',' Like '%,'+LocCd+',%') PSeq " + _
                     "            On LocCoCd=OdCoCd and vPMCd=OdPrdSeq " + _
                     "  Join Bag on " + IIF(gs_Partition = ctCurrPrtn, " BPrtKey=OdPrtKey and ", "") + " BOdIdNo=OdIdNo and BCoCd= OdCoCd and BOdTc= OdTc and BOdYy= OdYy and BOdChr= OdChr and BOdNo= OdNo and BOdSr= OdSr " + _
                     "  Join OrdMst on " + IIF(gs_Partition = ctCurrPrtn, " OmPrtKey=OdPrtKey and ", "") + " OmIdNo=OdOmIdNo and OmCoCd= OdCoCd and OmTc= OdTc and OmYy= OdYy and OmChr= OdChr and OmNo= OdNo " + wCnd + IIF(wAddlCnd <> "", IIF(wCnd <> "", " And ", " Where ") + wAddlCnd + IIF(gs_Partition = ctCurrPrtn, " and OdPrtKey='" + ctCurrPrtn + "' ", ""), "") + _
                     " Order By BYy, BChr, BNo, LPos"
          Call moCn.RepRes(Rep4.Subreport2.OpenSubreport, ws_SubRep1 + ctMaxDopOpt)  '****** Sachin 3.01 ctMaxDopOpt added
        End If
        'Call moCn.RepRes(Rep.Subreport1.OpenSubreport, ws_SubRep)
        '*** Jay 2.13(CT) ***
        '*** Jay 2.14 ***
    End If
    
    '578-31 checking max records in ordrm
    Dim wRsRmCnt As MDORowSet
    Dim wBagNo As String
    Set wRsRmCnt = moCn.OpenRes("Select BYy ,BChr,BNo, COUNT(OrRmCd) qRsCnt " + _
                  "From OrdDsg Join OrdMst on OmIdNo=OdOmIdNo and OmCoCd= OdCoCd and OmTc= OdTc and OmYy= OdYy and OmChr= OdChr and OmNo= OdNo  And OmPrtKey=OdPrtKey " + _
                  "Join Bag On BOdIdNo=OdIdNo and BCoCd = OdCoCd and BOdTc= OdTc and BOdYy= OdYy and BOdChr= OdChr and BOdNo= OdNo and BOdSr= OdSr  And BPrtKey=OdPrtKey " + _
                  "Join OrdRm On  BOdIdNo=OrOdIdNo and OrCoCd = BCoCd And ortc=Bodtc and oryy=Bodyy and orchr=Bodchr and orno=Bodno and orsr=Bodsr  And OrPrtKey=BPrtKey " + _
                  wCnd + IIF(wAddlCnd <> "", IIF(wCnd <> "", " And ", " Where ") + wAddlCnd, "") + IIF(gs_Partition = ctCurrPrtn, " And BPrtKey='" + ctCurrPrtn + "' ", "") + _
                  " group by BYy ,BChr,BNo ")
    
    '578-31 comparing max records in ordrm and dsr. if not matched giving error message
    With wRsRmCnt
      Do While Not (.EOF Or .BOF)
        If !qRsCnt > wRowCnt Then
          wBagNo = wBagNo + IIF(wBagNo = "", "", ", ") + !BYy + "/" + !BChr + "/" + CStr(!BNo)
        End If
        .MoveNext
      Loop
    End With
    '578-31 if yes report will open, if No report will not open for that format
    If wBagNo <> "" Then
      If MsgBox("The following Bags will have more Rm lines than what can be printed: " + _
        wBagNo + ". Do you want to continue Yes/No?", vbYesNo) = vbNo Then Exit Sub
    End If
  
      
    If UCase(adc("UoPMCdFr")) = "3D" Then GRP_REP.SetFormula mRep, "wPicOpt", "'3D'"
    If UCase(adc("UoPMCdFr")) = "HSK" Then GRP_REP.SetFormula mRep, "wPicOpt", "'HSK'"
    If UCase(adc("UoPMCdFr")) = "" Then GRP_REP.SetFormula mRep, "wPicOpt", "''": mRep.Pic.Suppress = True

    ' ****** Sachin 2.14.0 [13-10-2006] Bag RM Details Restriction ******
    GRP_REP.SetFormula mRep, "wBagFmt", "'" + UCase(adc("UOPMCD")) + "'"
    ' ****** Sachin 2.14.0 [13-10-2006] Bag RM Details Restriction ******

'Manoj 2.11.0, Correction
  Select Case UCase(adc("UOPMCD"))
    Case UCase("BAG2")
      mRep.Sections("G1HBAG2A").Suppress = True
      mRep.Sections("G2HBAG2A").Suppress = True
    Case UCase("BAG2A")
      mRep.Sections("G1HBAG2").Suppress = True
      mRep.Sections("DET").Suppress = True
      mRep.Line4.Suppress = True
    '*** Jay 2.13(CT) ***
    Case UCase("BAGCT")
      mRep.Line4.Suppress = True
    '*** Jay 2.13(CT) ***
    '*** Jay 2.14 ***
    ' ***** Manali 3.9.0 - BAG1 added (1 bag/page)
    Case UCase("BAG2B"), UCase("BAG1")
      GRP_REP.SetFormula mRep, "wBagFmt", "'" + UCase(adc("UOPMCD")) + "'"
      SetSubFormula mRep.Subreport1, "wBagFmt", "'" + UCase(adc("UOPMCD")) + "'"
      SetSubFormula mRep.Subreport2, "NoOfLocs", IIF(UCase(adc("UOPMCD")) = UCase("BAG1"), "25", "11")  '3.11.0
      If UCase(adc("UOPMCD")) = UCase("BAG1") Then
        Rep4.PICSEC.Height = 12100
        Rep4.Pic.Suppress = True
        Rep4.G1FA.Height = 1800
        SetSubFormula mRep.Subreport2, "wPicOpt", "'" + UCase(adc("UoPMCdFr")) + "'"
      Else
        Rep4.Subreport2_RFPicSecHd.Suppress = True
      End If
      mRep.Sections("DET").Suppress = True
      mRep.Line4.Suppress = True
      ' ***** Manali 3.6.0 - 03/12/09 - Actual Wt
      If adc("UoYn") = "Y" Then
        Rep4.Subreport1_Fldg1TotRmQty.Suppress = True: Rep4.Subreport1_Fldg1TotRmWt.Suppress = True
        Rep4.Subreport1_HTotQty.Suppress = True: Rep4.Subreport1_HTotWt.Suppress = True
      Else
        Rep4.Subreport1_Fldg1TotRmQty.Suppress = False: Rep4.Subreport1_Fldg1TotRmWt.Suppress = False
        Rep4.Subreport1_HTotQty.Suppress = False: Rep4.Subreport1_HTotWt.Suppress = False
      End If
      ' ***** Manali 3.6.0 - 03/12/09 - Actual Wt
    '*** Jay 2.14 ***
  End Select
'Manoj 2.11.0, Correction

  ' ***** Manali 3.8.0 - Single Bag Per Page
  If adc("UoYn2") = "Y" And UCase(adc("UOPMCD")) <> UCase("BAG6") Then
'    mRep.G1F.NewPageAfter = True
    GRP_REP.SetFormula mRep, "wOneBagPerPage", "'Y'"
  Else
    GRP_REP.SetFormula mRep, "wOneBagPerPage", "'N'"
  End If
  ' ***** Manali 3.8.0 - Single Bag Per Page
  '578.19 getting Rmcode name & its size for various format
  mFldName = "": mFontSz = 0
  Call GetRmcdNmFontSz(UCase(adc("UOPMCD")))
  '578.19 calling function to change Rmcode font size
  GRP_REP.SetFormula mRep, "wRmCdFontSz", GetRmcdFontSz(mFldName, mFontSz)

'*** If The Bag From Which Split Quantity Is Received Is Needed By The User Then Use This Query ***
'Select Ch.TdBYy, Ch.TdBChr, Ch.TdBNo,
'  '' as BOdTc, '' as BOdYy, '' as BOdChr, 0 as BOdNo, 0 as BOdSr, '' as BOdDmCd,
'  '' as BOdSfx, '' as BOdDmSz, 0 as BOpnQty, 0 as BOpnPtQty, '' as BOpnDt, '' as BOpnDesc,
'  0 as BQty, 0 as BPtQty, '' as OdDelDt, 0 as OdOrdQty, 0 as OdPrdQty, '' as OdDmPrdInst,
'  '' as OdCmPrdInst, '' as OdCmStmpInst, '' as OdSzInst, '' as OdKt, '' as OdDmCol,
'  '' as OrRmCtg, '' as OrRmSCtg, '' as OrRmCd, 0 as OrLn1, 0 as OrLn2, 0 as OrQty,
'  0 as OrWt, 0 as qMainMetWt, '' as OrMainMet, Pt.TdBYy as qSpFrBYy,
'  Pt.TdBChr as qSpFrBChr, Pt.TdBNo as qSpFrBNo
'From Txnd Ch
'join Txnd Pt on Pt.TdTc= Ch.TdTc and Pt.TdYy= Ch.TdYy and Pt.TdChr= Ch.TdChr and Pt.TdNo= Ch.TdNo
'where Ch.TdYyKey= (Select max(Mx.TdYyKey) from Txnd Mx join Bag on Mx.TdBYy= BYy
'      and Mx.TdBChr= BChr and Mx.TdBNo= BNo where
'      Mx.TdBYy= Ch.TdBYy and Mx.TdBChr= Ch.TdBChr and Mx.TdBNo= Ch.TdBNo and
'      (Select PValue from Param where PTyp= 'TC' and PMCd= Mx.TdTc)= 'BS' and
'      Mx.TdBDc= 'D' and Mx.TdSrNo= 0)
'  and Pt.TdBDc= 'C' and Pt.TdSrNo= 0 and Ch.TdSrNo= 0
  CRV_REP.DisplayGroupTree = False
  Call DispReport(adc) '****** Sachin 3.02 - Display Report ******
  
'*** (Jen 2.12)
  Exit Sub
RepErr:
  Call ChkSelScope(Err.Number, Err.Description)
'*** (Jen 2.12)
End Sub


Private Sub ChkBoxArr_Click(Index As Integer)
With ChkBoxArr(Index)
  Select Case Index
  ' ***** Manali 3.8.0 - Single Bag Per Page
  Case Is = 2               ' Tag =
    If .Value = Checked Then adc("UoYN2") = "Y" Else adc("UoYN2") = "N"
  End Select
End With

End Sub

Private Sub ChkBoxArr_GotFocus(Index As Integer)
  ' ***** Manali 3.8.0 - Single Bag Per Page
  DispMsg ChkBoxArr(Index).ToolTipText, etInfo

End Sub

'''' std code not to be changed
Private Sub Form_GotFocus()
  '*** Jenny Color (For Setting System Colors)
  SetMDISysColorsAfterLogin
  '*** Jenny Color (For Setting System Colors)
End Sub
Private Sub Form_Load()
  '*** Jenny Colour
  SetMDISysColorsAfterLogin   '*** Jenny Color (For Setting System Colors)
  
  Dim ws_BtnStr As String
  ws_BtnStr = ""
  Call SetObjColors(Me, ws_BtnStr)
  '*** Jenny Colour
End Sub
Private Sub ADC_Load()
    '*** Set the first Non Key Control as UoCoCdFr
    Set moCn = adc.Connection    '*** (Bef speed) bef 24/11/06
    
    '****** Sachin 2.14.0 - [23-11-2006] - [Multi-Company] ******
    ws_MultiCoMod = IIF(moCn.GetFldVal("Select hImpComLoc From Head Where HCd='" + ctSelfCmCd + "' And HCoCd='" + ctSelfCoCd + "'") = "Y", True, False)
    gb_CoCdFor = True
    Set adc.FirNKeyCtl = IIF(ws_MultiCoMod = True, adc("UoCoCdFr"), adc("UoOmTcFr"))
    adc("UOCOCDFR").Enabled = ws_MultiCoMod
    '****** Sachin 2.14.0 - [23-11-2006] - [Multi-Company] ******
 
    Call InitProp(Me)
    
    '*** Parameters Used For Setting Properties Of UsrOpt Fields ***
    gs_LocTyp = "'P'"
    'indigo.164 sketch design tctype and order tc added
    gs_DmTcTyp = "DM','SM"
    gs_OmTcTyp = "SO','OS"
    gs_CmCtg = "C"
    gs_PTyp = "PCT"
    gs_Tbl = "Param"
  ' Sachin 211 '
    gs_PTyp1 = "BAGFMT"
  ' Sachin 211 '
    '*** Parameters Used For Setting Properties Of UsrOpt Fields ***
    Call SetGroupSort
End Sub
Private Sub ATXT_FldChange(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String, ByVal pv_NewValue As Variant)

  '****** Sachin 2.14.0 - [24-11-2006] - [Multi-Company] ******
  gb_CoCdFor = True
  '*** Parameters Used For Setting Properties Of UsrOpt Fields ***
  gs_LocTyp = "'P'"
  'indigo.164 sketch design tctype and order tc added
  gs_DmTcTyp = "DM','SM"
  gs_OmTcTyp = "SO','OS"
  gs_CmCtg = "C"
  gs_PTyp = "PCT"
  gs_Tbl = "Param"
' Sachin 211 '
  gs_PTyp1 = "BAGFMT"
' Sachin 211 '
  '*** Parameters Used For Setting Properties Of UsrOpt Fields ***
  
  Select Case UCase(IdName)
    Case Is = UCase("UoCoCdFr")
    If Not moCn.RecSeek("Select 'x' From Head Where HCd='" + ctSelfCmCd + "' And HCoCd='" + pv_NewValue + "'") Then _
          Cancel = True: ErrMsg = "Invalid Company Code": Exit Sub
     Me.Caption = GetFrmCaption(pv_NewValue, "Bag Printing")
  End Select
  '****** Sachin 2.14.0 - [24-11-2006] - [Multi-Company] ******
  Call SetProp(Me, IdName, Valid, Cancel, ErrMsg)
  
End Sub
Private Sub ATXT_FldWhen(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String)
  gb_CoCdFor = True '***** Sachin 2.14.0 - [Multi-Company] ******
  '*** Parameters Used For Setting Properties Of UsrOpt Fields ***
  gs_LocTyp = "'P'"
  'indigo.164 sketch design tctype and order tc added
  gs_DmTcTyp = "DM','SM"
  gs_OmTcTyp = "SO','OS"
  gs_CmCtg = "C"
  gs_PTyp = "PCT"
  gs_Tbl = "Param"
' Sachin 211 '
  gs_PTyp1 = "BAGFMT"
' Sachin 211 '
  '*** Parameters Used For Setting Properties Of UsrOpt Fields ***
  
  SetProp Me, IdName, When
End Sub

' ****** Manali 3.5.0 - 05/11/08 - Radio Buttons Added
Private Sub OptPtr_Click(Index As Integer)
' ***** Manali 3.6.0 - 26/09/09 - Changes done To Show Actual Wt (UoYN)
If adc("UoYN") <> "Y" Then
  Select Case Index
  Case Is = 0
    adc("UoYN1") = "Y"
  Case Is = 1
    adc("UoYN1") = "N"
  End Select
Else
  adc("UoYN1") = "": OptPtr(0).Value = False: OptPtr(1).Value = False
  DispMsg "When Actual Rm Details Selected, This Option Cannot Be Selected ", etInfo
End If
' ***** Manali 3.6.0 - 26/09/09 - Changes done To Show Actual Wt (UoYN)
End Sub
' ****** Manali 3.5.0 - 05/11/08 - Radio Buttons Added

' ****** Manali 3.5.0 - 5/11/08 - Radio Buttons Added
Private Sub OptPtr_GotFocus(Index As Integer)
  DispMsg FraOptPtr.ToolTipText, etInfo
End Sub
' ****** Manali 3.5.0 - 05/11/08 - Radio Buttons Added

Private Sub OptRm_Click(Index As Integer)
' ***** Manali 3.6.0 - 26/09/09 - Changes done in query To Show Actual Wt (UoYN)
Select Case Index
Case Is = 0
  adc("UoYN") = "Y"
  adc("UoYN1") = "": OptPtr(0).Value = False: OptPtr(1).Value = False
Case Is = 1
  adc("UoYN") = "N"
End Select
' ***** Manali 3.6.0 - 26/09/09 - Changes done in query To Show Actual Wt (UoYN)
End Sub
' ***** Manali 3.6.0 - 26/09/09 - Changes done in query To Show Actual Wt (UoYN)
Private Sub OptRm_GotFocus(Index As Integer)
  DispMsg FraOptRm.ToolTipText, etInfo
End Sub
' ***** Manali 3.6.0 - 26/09/09 - Changes done in query To Show Actual Wt (UoYN)

Private Sub tab_REP_Click(PreviousTab As Integer)
    fra_tabrep(TAB_REP.Tab).Enabled = True
    fra_tabrep(PreviousTab).Enabled = False
End Sub
Private Sub ADC_KeyWhen()
    adc("UOUSRCD") = adc.UsrCd
    adc("UOMNUCD") = adc.MenuCd
End Sub
Private Sub ADC_NonKeyWhen(pr_Cancel As Boolean, pr_ErrMsg As String)
 adc("UoCoCdFr") = gs_CoCd
 
 '****** Sachin 2.14.0 - [24-11-2006] - [Multi-Company] ******
  Me.Caption = GetFrmCaption(adc("UOCOCDFR"), "Bag Printing")
  '****** Sachin 2.14.0 - [24-11-2006] - [Multi-Company] ******
 
' Sachin 211 '
 If adc("UoPmCd") = "" Then adc("UoPmCd") = "BAG2"
' Sachin 211 '

' **** Zubin 211 **** '
 If adc("UODTFR") = "" Then adc("UODTFR") = MWLib.EmptyDate
 If adc("UODTTO") = "" Then adc("UODTTO") = MWLib.EmptyDate
' **** Zubin 211 **** '

' ****** Manali 3.5.0 - 03/11/08 - Radio Buttons Added
If adc("UoYN1") = "Y" Then OptPtr(0).Value = True
If adc("UoYN1") = "N" Then OptPtr(1).Value = True
' ****** Manali 3.5.0 - 03/11/08 - Radio Buttons Added

' ***** Manali 3.6.0 - 26/09/09 - Changes done To Show Actual Wt (UoYN)
If adc("UoYN") = "Y" Then OptRm(0).Value = True
If adc("UoYN") = "N" Then OptRm(1).Value = True
' ***** Manali 3.6.0 - 26/09/09 - Changes done To Show Actual Wt (UoYN)

' ***** Manali 3.8.0
If adc("UoYn2") = "" Then adc("UoYn2") = "N"
If adc("UoYn2") = "Y" Then ChkBoxArr(2).Value = Checked Else ChkBoxArr(2).Value = Unchecked

End Sub
Private Sub DispCoNm()
Select Case UCase(adc("UOPMCD"))
  Case UCase("BAG2"), UCase("BAG2A"), UCase("BAG2B"), UCase("BAGCT"), UCase("BAG1")
    Dim ws_HName As String, ws_DtTm As String
    ws_HName = moCn.GetFldVal("Select HName from Head where HCoCd = '" + adc("UoCoCdFr") + "' and " + _
                                                      "HCd = '" + ctSelfCmCd + "' ")
    mRep.wCoCd.SetText ws_HName
    
'******** Manali - Before 3.2.2
'    ws_DtTm = Format$(moCn.SrvrDate, "DD-MMM-YYYY")
'    mRep.wSrvrDate.SetText ws_DtTm
'    ws_DtTm = Replace(Format(IIF(CDbl(moCn.SrvrTime) >= 13#, moCn.SrvrTime - 12, IIF(CDbl(moCn.SrvrTime) < 1#, moCn.SrvrTime + 12, moCn.SrvrTime)), "00.00"), ".", ":") + IIF(CDbl(moCn.SrvrTime) >= 12#, " PM", " AM")
'    mRep.wSrvrTime.SetText ws_DtTm

'******* Manali - User, Date and Time in reports
    Dim ws_Dt As String, ws_Tm As String
      ws_Dt = Format$(moCn.SrvrDate, IIF(UCase(MWLib.DtSetting) = "MM/DD/YY", "MMM-DD-YYYY", "DD-MMM-YYYY"))
      ws_Tm = Replace(Format(IIF(CDbl(moCn.SrvrTime) >= 13#, moCn.SrvrTime - 12, IIF(CDbl(moCn.SrvrTime) < 1#, moCn.SrvrTime + 12, moCn.SrvrTime)), "00.00"), ".", ":") + IIF(CDbl(moCn.SrvrTime) >= 12#, " PM", " AM")
    mRep.TxtRepUsrDtTime.SetText Trim(adc.UsrCd) + " / " + ws_Dt + " / " + ws_Tm

End Select
'  Dim ws_HName As String
'  ws_HName = moCn.GetFldVal("Select HName from Head where HCoCd = '" + adc("UoCoCdFr") + "' and " + _
'                                                      "HCd = '" + ctSelfCmCd + "' ")
'  Rep.wCoCd.SetText ws_HName
'  Dim ws_DtTm As String
'  ws_DtTm = Format$(moCn.SrvrDate, "DD-MMM-YYYY")
'  Rep.wSrvrDate.SetText ws_DtTm
'  ws_DtTm = Replace(Format(IIF(CDbl(moCn.SrvrTime) >= 13#, moCn.SrvrTime - 12, IIF(CDbl(moCn.SrvrTime) < 1#, moCn.SrvrTime + 12, moCn.SrvrTime)), "00.00"), ".", ":") + IIF(CDbl(moCn.SrvrTime) >= 12#, " PM", " AM")
'  Rep.wSrvrTime.SetText ws_DtTm
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
    Set mRep = Nothing
    Set Rep = Nothing
    Set Rep2 = Nothing
    Set Rep3 = Nothing
    Set Rep4 = Nothing
    Set Rep5 = Nothing
    '*** (09/08/05)
End Sub
'578.19 getting Rmcode name & its size for various format
Private Sub GetRmcdNmFontSz(ByVal pBagFmt As String)
Select Case UCase(pBagFmt)
  Case UCase("BAG2"), UCase("BAG2A")
    mFldName = "{rdo.OrRmCd}":  mFontSz = 8
  Case UCase("BAG2B"), UCase("BAG1")
    mFldName = "{rdo.OrRmCd}":  mFontSz = 7
  Case UCase("BAG3")
    mFldName = "{ado.OrRmCd}":  mFontSz = 7
  Case UCase("BAG6")
    mFldName = "{ado.OrRmCd}":  mFontSz = 7
  Case UCase("BAGCT")
    mFldName = "{rdo.OrRmCd}":  mFontSz = 8
    SetSubFormula Rep5.Subreport2, "wRmCdFontSzCT", GetRmcdFontSz("{ado.OctRmCd}", 8)
End Select

End Sub
