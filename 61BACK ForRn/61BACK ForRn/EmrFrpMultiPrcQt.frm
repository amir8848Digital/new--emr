VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{AB2CC32A-E652-42A2-B341-26ECE02F6386}#2.1#0"; "mwfctlforspeed9.ocx"
Begin VB.Form EmrFrpMultiPrcQt 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Multi Price Quotation Printing"
   ClientHeight    =   10185
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15270
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   10185
   ScaleWidth      =   15270
   WindowState     =   2  'Maximized
   Begin MwfCtl.MWCTL_HLP MWCTL_HLP1 
      Height          =   495
      Left            =   10710
      TabIndex        =   51
      TabStop         =   0   'False
      Top             =   9570
      Visible         =   0   'False
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
   End
   Begin MwfCtl.MWCTL_DST DST_REP 
      Height          =   570
      Left            =   4440
      TabIndex        =   50
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
      Left            =   4605
      TabIndex        =   60
      Top             =   9645
      Width           =   3555
      Begin MwfCtl.MWCTL_MED aktxt 
         Height          =   345
         Index           =   0
         Left            =   210
         TabIndex        =   61
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
         TabIndex        =   62
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
      Left            =   5820
      TabIndex        =   59
      Top             =   9600
      Visible         =   0   'False
      Width           =   855
      Begin CRVIEWER9LibCtl.CRViewer9 CRV_REP 
         Height          =   495
         Left            =   480
         TabIndex        =   56
         Top             =   -135
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
      TabIndex        =   52
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
      Height          =   9780
      Left            =   0
      TabIndex        =   55
      Top             =   45
      Width           =   15255
      Begin TabDlg.SSTab TAB_REP 
         Height          =   9465
         Left            =   15
         TabIndex        =   54
         Top             =   45
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   16695
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "&Data Range"
         TabPicture(0)   =   "EmrFrpMultiPrcQt.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "fra_tabrep(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "&Group Sort"
         TabPicture(1)   =   "EmrFrpMultiPrcQt.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "fra_tabrep(1)"
         Tab(1).ControlCount=   1
         Begin VB.Frame fra_tabrep 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   9070
            Index           =   1
            Left            =   -74940
            TabIndex        =   58
            Top             =   360
            Width           =   15105
            Begin MwfCtl.MWCTL_GRP GRP_REP 
               Height          =   360
               Left            =   240
               TabIndex        =   53
               Top             =   360
               Width           =   10335
               _ExtentX        =   18230
               _ExtentY        =   635
            End
         End
         Begin VB.Frame fra_tabrep 
            BorderStyle     =   0  'None
            Height          =   9070
            Index           =   0
            Left            =   60
            TabIndex        =   57
            Top             =   360
            Width           =   15105
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
               Left            =   30
               TabIndex        =   40
               Tag             =   "ShowCustDsg"
               ToolTipText     =   "Check To Show Customer Design Code"
               Top             =   5100
               Width           =   2100
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Show Cust RmCd   "
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
               Left            =   30
               TabIndex        =   42
               Tag             =   "ShowCustRm"
               ToolTipText     =   "Check To Show Customer Rm Code"
               Top             =   5385
               Width           =   2100
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Picture                   "
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
               Left            =   30
               TabIndex        =   32
               Tag             =   "ShowPic"
               ToolTipText     =   "Check To Show Picture"
               Top             =   3960
               Width           =   2100
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Show Dsg Values   "
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
               Left            =   30
               TabIndex        =   34
               Tag             =   "ShowDsgVal"
               ToolTipText     =   "Check To Show Design Value"
               Top             =   4245
               Width           =   2100
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Rm Value               "
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
               Left            =   30
               TabIndex        =   36
               Tag             =   "ShowRmVal"
               ToolTipText     =   "Check To Show Picture"
               Top             =   4530
               Width           =   2100
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
               Left            =   30
               TabIndex        =   38
               Tag             =   "ShowLabVal"
               ToolTipText     =   "Check To Show Labour Value"
               Top             =   4815
               Width           =   2100
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Final Sort On Dsg Cd                          "
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   540
               Index           =   11
               Left            =   30
               TabIndex        =   44
               Tag             =   "FinalSrt"
               ToolTipText     =   "Check To Show Picture"
               Top             =   5670
               Width           =   2085
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   42
               Left            =   2805
               TabIndex        =   43
               ToolTipText     =   "Final Groping on Design Code (Y/ N)"
               Top             =   5670
               Visible         =   0   'False
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               Enabled         =   0   'False
               DataType        =   4
               DataField       =   "UOYN11"
               ReCalcOn        =   "UOYN3"
               IdName          =   "UOYN11"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   2
               Left            =   2805
               TabIndex        =   31
               ToolTipText     =   "Show Picture (Yes / No)"
               Top             =   3960
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
               Index           =   5
               Left            =   2805
               TabIndex        =   33
               ToolTipText     =   "Show Price (Yes / No)"
               Top             =   4245
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
               Left            =   2805
               TabIndex        =   37
               ToolTipText     =   "Show Labor Value (Yes / No)"
               Top             =   4815
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
               Index           =   52
               Left            =   2805
               TabIndex        =   35
               ToolTipText     =   "Show Rm Value (Yes / No)"
               Top             =   4530
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
               Index           =   35
               Left            =   2805
               TabIndex        =   39
               ToolTipText     =   "Show Cust DsgCd (Yes / No)"
               Top             =   5100
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
               Index           =   38
               Left            =   2805
               TabIndex        =   41
               ToolTipText     =   "Show Cust RmCd (Yes / No)"
               Top             =   5385
               Visible         =   0   'False
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               Enabled         =   0   'False
               DataType        =   4
               DataField       =   "UOYN8"
               IdName          =   "UOYN8"
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Show Run Sr         "
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
               Index           =   25
               Left            =   30
               TabIndex        =   46
               Tag             =   "ShowRunSr"
               ToolTipText     =   "Check To Show Runing Serials"
               Top             =   6195
               Width           =   2070
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   18
               Left            =   5385
               TabIndex        =   5
               ToolTipText     =   "Enter To Order Year"
               Top             =   825
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               MaxLength       =   2
               DataType        =   4
               DataField       =   "UOMQMYYTO"
               IdName          =   "UOMQMYYTO"
               CmpStr          =   "MqdYy <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   20
               Left            =   2685
               TabIndex        =   2
               ToolTipText     =   "Enter From Order Year"
               Top             =   825
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               MaxLength       =   2
               DataType        =   4
               DataField       =   "UOMQMYYFR"
               IdName          =   "UOMQMYYFR"
               CmpStr          =   "MqdYy >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   43
               Left            =   1905
               TabIndex        =   1
               ToolTipText     =   "Enter From Order Tc"
               Top             =   825
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOMQMTCFR"
               IdName          =   "UOMQMTCFR"
               CmpStr          =   "MqdTc >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   45
               Left            =   3105
               TabIndex        =   3
               ToolTipText     =   "Enter From Order Character"
               Top             =   825
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOMQMCHRFR"
               IdName          =   "UOMQMCHRFR"
               CmpStr          =   "MqdChr >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   46
               Left            =   4590
               TabIndex        =   4
               ToolTipText     =   "Enter To Order Tc"
               Top             =   825
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOMQMTCTO"
               IdName          =   "UOMQMTCTO"
               CmpStr          =   "MqdTc <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   47
               Left            =   5805
               TabIndex        =   6
               ToolTipText     =   "Enter To Order Character"
               Top             =   825
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOMQMCHRTO"
               IdName          =   "UOMQMCHRTO"
               CmpStr          =   "MqdChr <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   21
               Left            =   5520
               TabIndex        =   10
               ToolTipText     =   "Enter To Order Serial"
               Top             =   1110
               Width           =   525
               _ExtentX        =   926
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOMQDSRTO"
               IdName          =   "UOMQDSRTO"
               CmpStr          =   "MqdSr <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   23
               Left            =   2805
               TabIndex        =   8
               ToolTipText     =   "Enter From Order Serial"
               Top             =   1110
               Width           =   525
               _ExtentX        =   926
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOMQDSRFR"
               IdName          =   "UOMQDSRFR"
               CmpStr          =   "MqdSr >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   26
               Left            =   1905
               TabIndex        =   7
               ToolTipText     =   "Enter From Order Number"
               Top             =   1110
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOMQMNOFR"
               IdName          =   "UOMQMNOFR"
               CmpStr          =   "MqdNo >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   29
               Left            =   4590
               TabIndex        =   9
               ToolTipText     =   "Enter To Order Number"
               Top             =   1110
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOMQMNOTO"
               IdName          =   "UOMQMNOTO"
               CmpStr          =   "MqdNo <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   0
               Left            =   1905
               TabIndex        =   29
               ToolTipText     =   "Enter Level Of Raw Material Detail To Be Shown"
               Top             =   3390
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
               Index           =   17
               Left            =   1905
               TabIndex        =   0
               ToolTipText     =   "Enter Company Code"
               Top             =   540
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               Enabled         =   0   'False
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOCOCDFR"
               ReCalcOn        =   "UOTCTYPFR"
               IdName          =   "UOCOCDFR"
               CmpStr          =   "MqdCoCd = "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   31
               Left            =   4590
               TabIndex        =   23
               ToolTipText     =   "Enter To Design Code"
               Top             =   2535
               Width           =   2385
               _ExtentX        =   4207
               _ExtentY        =   503
               MaxLength       =   15
               DataField       =   "UODMCDTO"
               IdName          =   "UODMCDTO"
               CmpStr          =   "MqdDmCd <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   32
               Left            =   7290
               TabIndex        =   24
               ToolTipText     =   "Enter Design Code Selection"
               Top             =   2535
               Width           =   7650
               _ExtentX        =   13494
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UODMCDSEL"
               IdName          =   "UODMCDSEL"
               CmpStr          =   "MqdDmCd In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   33
               Left            =   1905
               TabIndex        =   22
               ToolTipText     =   "Enter From Design Code"
               Top             =   2535
               Width           =   2385
               _ExtentX        =   4207
               _ExtentY        =   503
               MaxLength       =   15
               DataField       =   "UODMCDFR"
               IdName          =   "UODMCDFR"
               CmpStr          =   "MqdDmCd >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   34
               Left            =   4590
               TabIndex        =   26
               ToolTipText     =   "Enter To Design Size"
               Top             =   2820
               Width           =   585
               _ExtentX        =   1032
               _ExtentY        =   503
               MaxLength       =   5
               DataField       =   "UODMSZTO"
               IdName          =   "UODMSZTO"
               CmpStr          =   "MqdDmSz <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   36
               Left            =   1905
               TabIndex        =   25
               ToolTipText     =   "Enter From Design Size"
               Top             =   2820
               Width           =   585
               _ExtentX        =   1032
               _ExtentY        =   503
               MaxLength       =   5
               DataField       =   "UODMSZFR"
               IdName          =   "UODMSZFR"
               CmpStr          =   "MqdDmSz >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   28
               Left            =   4590
               TabIndex        =   20
               ToolTipText     =   "Enter To Design Category"
               Top             =   2250
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
               Index           =   30
               Left            =   1905
               TabIndex        =   19
               ToolTipText     =   "Enter From Design Category"
               Top             =   2250
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
               Index           =   8
               Left            =   7290
               TabIndex        =   21
               ToolTipText     =   "Enter Design Category Selection"
               Top             =   2250
               Width           =   7650
               _ExtentX        =   13494
               _ExtentY        =   503
               MaxLength       =   255
               DataType        =   4
               DataField       =   "UODMCTGSEL"
               IdName          =   "UODMCTGSEL"
               CmpStr          =   "DmCtg In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   22
               Left            =   4590
               TabIndex        =   28
               ToolTipText     =   "Enter To Customer"
               Top             =   3105
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   10
               DataField       =   "UOCMCDTO"
               IdName          =   "UOCMCDTO"
               CmpStr          =   "MqmCmCd <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   24
               Left            =   1905
               TabIndex        =   27
               ToolTipText     =   "Enter From Customer"
               Top             =   3105
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UOCMCDFR"
               IdName          =   "UOCMCDFR"
               CmpStr          =   "MqmCmCd >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   44
               Left            =   1905
               TabIndex        =   30
               ToolTipText     =   "Enter Currency Code Whose Orders You Wish To View Or Leave It Blank If You Wish To View All Orders In Base Currency"
               Top             =   3675
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
               Index           =   64
               Left            =   2805
               TabIndex        =   45
               ToolTipText     =   "Show Components Details (Yes / No)"
               Top             =   6135
               Visible         =   0   'False
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               Enabled         =   0   'False
               DataType        =   4
               DataField       =   "UOYN25"
               IdName          =   "UOYN25"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   71
               Left            =   4590
               TabIndex        =   12
               ToolTipText     =   "Enter To Order Date"
               Top             =   1395
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   3
               DataField       =   "UODTTO"
               IdName          =   "UODTTO"
               CmpStr          =   "MqmDt <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   72
               Left            =   1905
               TabIndex        =   11
               ToolTipText     =   "Enter From Order Date"
               Top             =   1395
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   3
               DataField       =   "UODTFR"
               IdName          =   "UODTFR"
               CmpStr          =   "MqmDt >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   3
               Left            =   4590
               TabIndex        =   14
               ToolTipText     =   "Enter To Design Category"
               Top             =   1680
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataType        =   4
               DataField       =   "UOMETGRDTO"
               IdName          =   "UOMETGRDTO"
               CmpStr          =   "MqdMetGrd <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   4
               Left            =   1905
               TabIndex        =   13
               ToolTipText     =   "Enter From Design Category"
               Top             =   1680
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataType        =   4
               DataField       =   "UOMETGRDFR"
               IdName          =   "UOMETGRDFR"
               CmpStr          =   "MqdMetGrd >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   6
               Left            =   7290
               TabIndex        =   15
               ToolTipText     =   "Enter Design Category Selection"
               Top             =   1680
               Width           =   7650
               _ExtentX        =   13494
               _ExtentY        =   503
               MaxLength       =   255
               DataType        =   4
               DataField       =   "UOMETGRDSEL"
               IdName          =   "UOMETGRDSEL"
               CmpStr          =   "MqdMetGrd In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   7
               Left            =   4590
               TabIndex        =   17
               ToolTipText     =   "Enter To Design Category"
               Top             =   1965
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataType        =   4
               DataField       =   "UODIAGRDTO"
               IdName          =   "UODIAGRDTO"
               CmpStr          =   "MqrDiaGrd <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   9
               Left            =   1905
               TabIndex        =   16
               ToolTipText     =   "Enter From Design Category"
               Top             =   1965
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataType        =   4
               DataField       =   "UODIAGRDFR"
               IdName          =   "UODIAGRDFR"
               CmpStr          =   "MqrDiaGrd >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   10
               Left            =   7290
               TabIndex        =   18
               ToolTipText     =   "Enter Design Category Selection"
               Top             =   1965
               Width           =   7650
               _ExtentX        =   13494
               _ExtentY        =   503
               MaxLength       =   255
               DataType        =   4
               DataField       =   "UODIAGRDSEL"
               IdName          =   "UODIAGRDSEL"
               CmpStr          =   "MqrDiaGrd In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   11
               Left            =   4590
               TabIndex        =   47
               ToolTipText     =   "Enter Lab Main Code"
               Top             =   3735
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               DataField       =   "UOLMCD1"
               IdName          =   "UOLMCD1"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   12
               Left            =   4590
               TabIndex        =   48
               ToolTipText     =   "Enter Lab Main Code"
               Top             =   4020
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               DataField       =   "UOLMCD2"
               IdName          =   "UOLMCD2"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   13
               Left            =   4590
               TabIndex        =   49
               ToolTipText     =   "Enter Lab Main Code"
               Top             =   4305
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               DataField       =   "UOLMCD3"
               IdName          =   "UOLMCD3"
            End
            Begin VB.Label Label3 
               BackStyle       =   0  'Transparent
               Caption         =   "Option 1"
               BeginProperty Font 
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
               Left            =   3240
               TabIndex        =   87
               Top             =   3735
               Width           =   1395
            End
            Begin VB.Label Label4 
               BackStyle       =   0  'Transparent
               Caption         =   "Option 2"
               BeginProperty Font 
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
               Left            =   3240
               TabIndex        =   86
               Top             =   4020
               Width           =   1395
            End
            Begin VB.Label Label5 
               BackStyle       =   0  'Transparent
               Caption         =   "Option 3"
               BeginProperty Font 
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
               Left            =   3240
               TabIndex        =   85
               Top             =   4305
               Width           =   1395
            End
            Begin VB.Label Label7 
               BackStyle       =   0  'Transparent
               Caption         =   "Labour Details"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   3930
               TabIndex        =   84
               Top             =   3450
               Width           =   1845
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Dia Grades"
               BeginProperty Font 
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
               Left            =   75
               TabIndex        =   83
               Top             =   1965
               Width           =   1845
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Metal Grades"
               BeginProperty Font 
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
               Left            =   75
               TabIndex        =   82
               Top             =   1680
               Width           =   1845
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Order Date"
               BeginProperty Font 
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
               TabIndex        =   81
               Top             =   1395
               Width           =   1425
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
               TabIndex        =   80
               Top             =   3675
               Width           =   1005
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
               Index           =   8
               Left            =   75
               TabIndex        =   79
               Top             =   3105
               Width           =   1845
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
               Left            =   75
               TabIndex        =   78
               Top             =   2250
               Width           =   1845
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
               Left            =   7290
               TabIndex        =   77
               Top             =   120
               Width           =   1575
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
               Left            =   75
               TabIndex        =   76
               Top             =   2535
               Width           =   1845
            End
            Begin VB.Label ALBL 
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
               Index           =   0
               Left            =   75
               TabIndex        =   75
               Top             =   2820
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
               Left            =   75
               TabIndex        =   74
               Top             =   540
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
               TabIndex        =   73
               ToolTipText     =   "Location"
               Top             =   3390
               Width           =   1845
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
               Left            =   75
               TabIndex        =   72
               Top             =   810
               Width           =   1845
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
               Left            =   75
               TabIndex        =   71
               Top             =   1110
               Width           =   1845
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
               Left            =   5745
               TabIndex        =   70
               Top             =   825
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
               Left            =   3045
               TabIndex        =   69
               Top             =   825
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
               Left            =   5325
               TabIndex        =   68
               Top             =   825
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
               Left            =   2625
               TabIndex        =   67
               Top             =   825
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
               Left            =   5445
               TabIndex        =   66
               Top             =   1110
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
               Left            =   2745
               TabIndex        =   65
               Top             =   1110
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
               Left            =   4290
               TabIndex        =   64
               Top             =   120
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
               Left            =   1905
               TabIndex        =   63
               Top             =   120
               Width           =   1335
            End
         End
      End
   End
End
Attribute VB_Name = "EmrFrpMultiPrcQt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rep As New EmrRepMultiPrcQt
Dim moCn As MwfLib.MDOConnection   '*** (Bef speed) bef 24/11/06
Private Sub SetGroupSort()
  'GRP_REP.Add "Sub Ctg", "RmSCtg", "", "RmSCtg", "hRmSCtg", "", "", "", "(Select PDesc From Param Where PTyp= 'RMSCTG' and PMCd = RmCtg and PSCd= RmSCtg) "
End Sub

Private Sub adc_setreprecsource()
  On Error GoTo RepErr
  
  Dim wrepcnd As String, wCnd As String, wSqlStrg As String
  Dim wCommStr As String, wRmStr As String, wRmGrp As String, wRmCommStr As String
  Dim wRmDet As String, wRmValOpt As String, wRmSetValOpt As String
  Dim wLabValOpt As String
  Dim wRtStr As String
  Dim wGrdPrc, wGrdPrcOpt As String
  Dim ws_BaseCurCd As String, wb_InBaseCur As Boolean
  Dim wOurDsgCd As String, wOurRmCd As String
  
  Set Rep = Nothing
  adc("UoCoCdFr") = gs_CoCd
  Set adc.RepSource = Rep
  
  ws_BaseCurCd = moCn.GetFldVal("Select HBaseCurCd from Head where HCd= '" + ctSelfCmCd + "' and " + _
                  "HCoCd= '" + adc("UoCoCdFr") + "'")
  If adc("UoCurCdFr") = "" Then
    wb_InBaseCur = True
    Rep.TxtValIn.SetText "Value In " + ws_BaseCurCd
  Else
    wb_InBaseCur = False
    Rep.TxtValIn.SetText "Value In " + adc("UoCurCdFr")
  End If
  
  Call DispCoNm
  Call SetMwName(Rep)   ' ***** Manali 3.6.1 - Microway footer
    
   '3.11.2 - Getting Encrypted values
    
   wrepcnd = adc.RepCond
   wCnd = IIF(wrepcnd <> "", " where " + wrepcnd, "")
   wCnd = wCnd + IIF(wb_InBaseCur = True, "", " and MqmCmCurCd= '" + adc("UoCurCdFr") + "' ")
   
   wGrdPrc = "(Case MqrDiaGrd When MqmDiaGrd1 Then Cast(Convert(varChar(100), DecryptByKey(Convert(varBinary(100), MqdDiaGrd1PrcBin))) as Decimal(16,2))  " + _
             "                When MqmDiaGrd2 Then Cast(Convert(varChar(100), DecryptByKey(Convert(varBinary(100), MqdDiaGrd2PrcBin))) as Decimal(16,2))  " + _
             "                When MqmDiaGrd3 Then Cast(Convert(varChar(100), DecryptByKey(Convert(varBinary(100), MqdDiaGrd3PrcBin))) as Decimal(16,2)) " + _
             "                When MqmDiaGrd4 Then Cast(Convert(varChar(100), DecryptByKey(Convert(varBinary(100), MqdDiaGrd4PrcBin))) as Decimal(16,2)) " + _
             "                When MqmDiaGrd5 Then Cast(Convert(varChar(100), DecryptByKey(Convert(varBinary(100), MqdDiaGrd5PrcBin))) as Decimal(16,2)) " + _
             "                When MqmDiaGrd6 Then Cast(Convert(varChar(100), DecryptByKey(Convert(varBinary(100), MqdDiaGrd6PrcBin))) as Decimal(16,2)) " + _
             "                When MqmDiaGrd7 Then Cast(Convert(varChar(100), DecryptByKey(Convert(varBinary(100), MqdDiaGrd7PrcBin))) as Decimal(16,2)) " + _
             "                When MqmDiaGrd8 Then Cast(Convert(varChar(100), DecryptByKey(Convert(varBinary(100), MqdDiaGrd8PrcBin))) as Decimal(16,2)) Else 0 End)"
   
   If adc("UoYN2") = "Y" Then
     wGrdPrcOpt = " Max(" + IIF(wb_InBaseCur, " (Case when MqmMulDiv= 'M' then " + wGrdPrc + "/MqmCnvFct Else " + wGrdPrc + "*MqmCnvFct end) ", wGrdPrc) + ") "
   ElseIf adc("UoYN2") = "N" Then
     wGrdPrcOpt = " 0.0 "
     Rep.TxtLmg.Suppress = True: Rep.FldLmg.Suppress = True
     Rep.TxtLmp.Suppress = True: Rep.FldLmp.Suppress = True
     Rep.TxtLms.Suppress = True: Rep.FldLms.Suppress = True
     Rep.TxtLml.Suppress = True: Rep.FldLml.Suppress = True
   End If
   
   If adc("UoYN14") = "Y" Then
     wRmValOpt = "sum(" + IIF(wb_InBaseCur, " (Case when MqmMulDiv= 'M' then Cast(Convert(varChar(100), DecryptByKey(Convert(varBinary(100), MqrSalValBin))) as Decimal(16,2))/MqmCnvFct else Cast(Convert(varChar(100), DecryptByKey(Convert(varBinary(100), MqrSalValBin))) as Decimal(16,2))*MqmCnvFct end) ", "Cast(Convert(varChar(100), DecryptByKey(Convert(varBinary(100), MqrSalValBin))) as Decimal(16,2))") + ")"
   ElseIf adc("UoYN14") = "N" Then
     wRmValOpt = "0.00"
   End If
   
   If (adc("UoYN4")) = "Y" Then
       wRmSetValOpt = "sum(" + IIF(wb_InBaseCur, " (Case when MqmMulDiv= 'M' then Cast(Convert(varChar(100), DecryptByKey(Convert(varBinary(100), MqrSetSalValBin))) as Decimal(16,2))/MqmCnvFct else Cast(Convert(varChar(100), DecryptByKey(Convert(varBinary(100), MqrSetSalValBin))) as Decimal(16,2))*MqmCnvFct end) ", "Cast(Convert(varChar(100), DecryptByKey(Convert(varBinary(100), MqrSetSalValBin))) as Decimal(16,2))") + ")"
       wLabValOpt = "Cast(Convert(varChar(100), DecryptByKey(Convert(varBinary(100), MqlSalValBin))) as Decimal(16,2))"
      Rep.hLabCd1.SetText adc("UoLmCd1")
      Rep.hLabCd2.SetText adc("UoLmCd2")
      Rep.hLabCd3.SetText adc("UoLmCd3")
   Else
       wRmSetValOpt = "0.00": wLabValOpt = "0.00"
      Rep.hLabCd1.SetText ""
      Rep.hLabCd2.SetText ""
      Rep.hLabCd3.SetText ""
      Rep.hLabOth.SetText ""
      Rep.hLabSet.SetText ""
   End If
   
   If UCase(adc("UoYN7")) = "Y" Then
     wOurDsgCd = "IsNull((Select cast(CdCustDmCd as Varchar(17)) from CustDsg " + _
                               "where " + IIF(gs_Partition = ctCurrPrtn, " CdPrtKey='" + ctCurrPrtn + "' and ", "") + " CdCtg = 'C' and CdCd = Max(MqmCmCd) and " + _
                               "CdOurDmCd = Max(MqdDmCd) and CdOurSfx = '' and " + _
                               "CdOurDmSz = Max(MqdDmSz) ), " + _
                             "IsNull((Select cast(CdCustDmCd as Varchar(17)) from CustDsg " + _
                                   "where " + IIF(gs_Partition = ctCurrPrtn, " CdPrtKey='" + ctCurrPrtn + "' and ", "") + " CdCtg = 'C' and CdCd = Max(MqmCmCd) and " + _
                                   "CdOurDmCd = Max(MqdDmCd) and CdOurSfx = '' and " + _
                                   "CdOurDmSz= ''), '{' + max(RTrim(MqdDmCd)) + '}')) as CdOurDmCd"
   Else
     'wOurDsgCd = "Max(MqdDmCd) as CdOurDmCd"
     wOurDsgCd = "'' as CdOurDmCd"
   End If
   
   
     
   wCommStr = " Open Symmetric Key DataEnde Decryption By Password = 'RMeWm'; " + _
              " Select MqmCoCd, MqmYy, MqmTc, MqmChr, MqmNo, MqdSr, MqrDiaGrd, max(MqmCmCd) as MqmCmCd, max(MqmDt) as MqmDt, " + _
              " max(" + IIF(wb_InBaseCur, " (Case when MqmMulDiv= 'M' then MqmLmgSal/MqmCnvFct else MqmLmgSal*MqmCnvFct end) ", "MqmLmgSal") + ") as MqmLmgSal, " + _
              " max(" + IIF(wb_InBaseCur, " (Case when MqmMulDiv= 'M' then MqmLmpSal/MqmCnvFct else MqmLmpSal*MqmCnvFct end) ", "MqmLmpSal") + ") as MqmLmpSal," + _
              " max(" + IIF(wb_InBaseCur, " (Case when MqmMulDiv= 'M' then MqmLmsSal/MqmCnvFct else MqmLmsSal*MqmCnvFct end) ", "MqmLmsSal") + ") as MqmLmsSal, " + _
              " max(" + IIF(wb_InBaseCur, " (Case when MqmMulDiv= 'M' then MqmLmlSal/MqmCnvFct else MqmLmlSal*MqmCnvFct end) ", "MqmLmlSal") + ") as MqmLmlSal," + _
              " max(CmName) as CmName, " + _
              " max(MqdMetGrd) as MqdMetGrd, max(MqdDmCd) as MqdDmCd," + wOurDsgCd + ", max(MqdDmSz) as MqdDmSz, " + _
              " max(DmCtg) as DmCtg, max(DmSalCtg) as DmSalCtg, max(DmTcTyp) as DmTcTyp, " + _
              " max(DmUom) as DmUom, " + _
              " Max(MqdDiaWt) as MqdDiaWt, max(MqdGrsWt) as MqdGrsWt, max(MqdGldAsWt) as MqdGldAsWt, max(MqdLabAsWt) as MqdLabAsWt," + wGrdPrcOpt + " as qGrdPrc, " + _
              " Sum(Case When MqrRmCtg In ('G', 'P', 'S', 'L', 'M') Then MqrWt Else 0 End) as qNetWt, " + _
              wRmValOpt + " as qTotRmVal, " + wRmSetValOpt + " as qTotSetVal, " + _
              " Max(" + IIF(wb_InBaseCur, " (Case when MqmMulDiv= 'M' then qLabVal1/MqmCnvFct else qLabVal1*MqmCnvFct end) ", "qLabVal1") + ") as qLabVal1, " + _
              " Max(" + IIF(wb_InBaseCur, " (Case when MqmMulDiv= 'M' then qLabVal2/MqmCnvFct else qLabVal2*MqmCnvFct end) ", "qLabVal2") + ") as qLabVal2, " + _
              " Max(" + IIF(wb_InBaseCur, " (Case when MqmMulDiv= 'M' then qLabVal3/MqmCnvFct else qLabVal3*MqmCnvFct end) ", "qLabVal3") + ") as qLabVal3, " + _
              " Max(" + IIF(wb_InBaseCur, " (Case when MqmMulDiv= 'M' then qTotLabVal/MqmCnvFct else qTotLabVal*MqmCnvFct end) ", "qTotLabVal") + ") as qTotLabVal "
  
   wCommStr = wCommStr + " From MultiPrcQtMst " + _
              "join MultiPrcQtDsg on MqdMqmIdNo=MqmIdNo and MqdCoCd= MqmCoCd and MqdTc= MqmTc and MqdYy= MqmYy " + _
              "and MqdChr= MqmChr and MqdNo= MqmNo " + IIF(gs_Partition = ctCurrPrtn, " and MqdPrtKey=MqmPrtKey ", "") + _
              "join DsgMst on DmIdNo=MqdDmIdNo and DmTcTyp='DM' and DmCd= MqdDmCd and DmSz='' " + IIF(gs_Partition = ctCurrPrtn, " and DmPrtKey=MqdPrtKey ", "") + _
              "join CustMst on CmCtg='C' and CmCd= MqmCmCd " + _
              "Join MultiPrcQtRm On MqdCoCd=MqrCoCd And MqdTc=MqrTc And MqdYy=MqrYy And MqdChr=MqrChr And MqdNo=MqrNo And MqdSr=MqrSr " + _
              "Join (Select MqlCoCd, MqlTc, MqlYy, MqlChr, MqlNo, MqlSr, " + _
              "Sum(Case When Convert(varChar(100), DecryptByKey(Convert(varBinary(100), MqlMCdBin)))='" + adc("UoLmCd1") + "' Then " + wLabValOpt + " Else 0 End) as qLabVal1, " + _
              "Sum(Case When Convert(varChar(100), DecryptByKey(Convert(varBinary(100), MqlMCdBin)))='" + adc("UoLmCd2") + "' Then " + wLabValOpt + " Else 0 End) as qLabVal2, " + _
              "Sum(Case When Convert(varChar(100), DecryptByKey(Convert(varBinary(100), MqlMCdBin)))='" + adc("UoLmCd3") + "' Then " + wLabValOpt + " Else 0 End) as qLabVal3, " + _
              "Sum(" + wLabValOpt + ") as qTotLabVal From MultiPrcQtLab " + _
              "Group By MqlCoCd, MqlTc, MqlYy, MqlChr, MqlNo, MqlSr) MultiPrcQtLab " + _
              "On MqlCoCd=MqdCoCd and MqlTc=MqdTc and MqlYy=MqdYy and MqlChr=MqdChr and MqlNo=MqdNo and MqlSr=MqdSr  " + _
              wCnd + "Group by MqmCoCd, MqmTc, MqmYy, MqmChr, MqmNo, MqdSr, MqrDiaGrd "
   
   wSqlStrg = wCommStr + " order by MqmCoCd, MqmYy, MqmTc, MqmChr, MqmNo, MqdSr, MqrDiaGrd"
  
   adc.RepRecSource = wSqlStrg + ctMaxDopOpt + ";Close Symmetric Key DataEnde;"
  
     
   If adc("UoYN") = "N" Then
     Rep.PICSEC.Suppress = True
   End If
   If GetUsrLevel > 5 Then GRP_REP.SetFormula Rep, "wCustCdNm", "{rdo.MqmCmCd}"
   
   If adc("UoRmGrp") = "N" Then
     Rep.SubRep.Suppress = True
     Rep.DetSubRep.Suppress = True
   Else
     If UCase(adc("UoYN8")) = "Y" Then
       Select Case UCase(adc("UoRmGrp"))
         Case Is = UCase("Sr"), UCase("RmRSzRt")
           wOurRmCd = "IsNull((Select cast(CrCustRmCd as VarChar(14)) from CustRm where " + IIF(gs_Partition = ctCurrPrtn, " CrPrtKey='" + ctCurrPrtn + "' and ", "") + " CrCtg = 'C' and " + _
                   " CrCd = Max(MqmCmCd) and CrRmCtg = Max(MqrRmCtg) and " + _
                   " CrOurRmCd = Max(Convert(varChar(100), DecryptByKey(Convert(varBinary(100), MqrRmCdBin)))) and Max(MqrLn1) Between CrFrLn and CrToLn)," + _
                   " '{' + Max(Convert(varChar(100), DecryptByKey(Convert(varBinary(100), MqrRmCdBin)))) + '}') as MqrRmCd"
         Case Is = UCase("RmCdRt"), UCase("RmCd")
           wOurRmCd = "IsNull((Select cast(CrCustRmCd as VarChar(14)) from CustRm where " + IIF(gs_Partition = ctCurrPrtn, " CrPrtKey='" + ctCurrPrtn + "' and ", "") + " CrCtg = 'C' and " + _
                   "CrCd = Max(MqmCmCd) and CrRmCtg = Max(MqrRmCtg) and " + _
                   "CrOurRmCd = Max(Convert(varChar(100), DecryptByKey(Convert(varBinary(100), MqrRmCdBin)))) and 0 Between CrFrLn and CrToLn)," + _
                   "'{' + Max(Convert(varChar(100), DecryptByKey(Convert(varBinary(100), MqrRmCdBin)))) + '}') as MqrRmCd"
       End Select
     Else
       wOurRmCd = "Max(Convert(varChar(100), DecryptByKey(Convert(varBinary(100), MqrRmCdBin)))) as MqrRmCd"
     End If
  
     wRtStr = " case when (RmQw= 'Q' and sum(MqrQty)<>0) then " + wRmValOpt + " / sum(MqrQty) " + _
                    "when (RmQw = 'W' and sum(MqrWt)<>0) then " + wRmValOpt + " / sum(MqrWt) " + _
                    "else 0 end "
  
     Select Case UCase(adc("UoRmGrp"))
     Case Is = UCase("N") 'MqrRmCtg to get RmVal and RmSetVal, No dets r shown
       wRmStr = " 0 as MqrSrNo, MqrRmCtg , '*' as MqrRmSCtg, '*' as MqrRmCd, '*' as RmQw, 0 as MqrLn1, " + _
                "0 as MqrLn2, 0 as MqrLn3, 0 as MqrRmPtr, sum(MqrQty) as MqrQty, sum(MqrWt) as MqrWt, 0.00 as MqrLmeSal, 0.00 as MqrSalRt, '*' as MqrSetSCd, " + _
                "0.00 as MqrSetSalRt, " + wRmSetValOpt + " as MqrSetSalVal, " + wRmValOpt + " as MqrSalVal, '' as RmRSzRt "
       wRmGrp = " Group by MqmCoCd, MqmYy, MqmTc, MqmChr, MqmNo, MqdSr, MqrDiaGrd, MqrRmCtg "
     Case Is = UCase("Sr")
       wRmStr = " MqrSrNo, max(MqrRmCtg) as MqrRmCtg, max(MqrRmSCtg) as MqrRmSCtg, " + wOurRmCd + ",max(RmQw) As RmQw, max(MqrLn1) as MqrLn1, " + _
                "max(MqrLn2) as MqrLn2, max(MqrLn3) as MqrLn3, " + _
                "max(MqrRmPtr) as MqrRmPtr , sum(MqrQty) as MqrQty, sum(MqrWt) as MqrWt, " + _
                "max(" + IIF(wb_InBaseCur, " (Case when MqmMulDiv= 'M' then MqrLmeSal/MqmCnvFct else MqrLmeSal*MqmCnvFct end) ", "MqrLmeSal") + ") as MqrLmeSal, " + _
                "max(" + IIF(wb_InBaseCur, " (Case when MqmMulDiv= 'M' then Cast(Convert(varChar(100), DecryptByKey(Convert(varBinary(100), MqrSalRtBin))) as Decimal(16,2))/MqmCnvFct else Cast(Convert(varChar(100), DecryptByKey(Convert(varBinary(100), MqrSalRtBin))) as Decimal(16,2))*MqmCnvFct end) ", "Cast(Convert(varChar(100), DecryptByKey(Convert(varBinary(100), MqrSalRtBin))) as Decimal(16,2))") + ") as MqrSalRt, max(MqrSetSCd) as MqrSetSCd, " + _
                "max(" + IIF(wb_InBaseCur, " (Case when MqmMulDiv= 'M' then Cast(Convert(varChar(100), DecryptByKey(Convert(varBinary(100), MqrSetSalRtBin))) as Decimal(16,2))/MqmCnvFct else Cast(Convert(varChar(100), DecryptByKey(Convert(varBinary(100), MqrSetSalRtBin))) as Decimal(16,2))*MqmCnvFct end) ", "Cast(Convert(varChar(100), DecryptByKey(Convert(varBinary(100), MqrSetSalRtBin))) as Decimal(16,2))") + ") as MqrSetSalRt, " + _
                wRmSetValOpt + " as MqrSetSalVal, " + wRmValOpt + " as MqrSalVal, '' as RmRSzRt "
       wRmGrp = " Group by MqmCoCd, MqmYy, MqmTc, MqmChr, MqmNo, MqdSr, MqrDiaGrd, MqrSrNo "
     Case Is = UCase("RmCdRt")
       wRmStr = " 0 as MqrSrNo, MqrRmCtg, MqrRmSCtg, " + wOurRmCd + ", RmQw, 0.0 as MqrLn1, 0.0 as MqrLn2, 0.0 as MqrLn3, 0.0 as MqrRmPtr, " + _
                "sum(MqrQty) as MqrQty, sum(MqrWt) as MqrWt, " + _
                IIF(wb_InBaseCur, " (Case when max(MqmMulDiv)= 'M' then MqrLmeSal/max(MqmCnvFct) else MqrLmeSal*max(MqmCnvFct) end) ", "MqrLmeSal") + " as MqrLmeSal, " + _
                IIF(wb_InBaseCur, " (Case when max(MqmMulDiv)= 'M' then Cast(Convert(varChar(100), DecryptByKey(Convert(varBinary(100), MqrSalRtBin))) as Decimal(16,2))/max(MqmCnvFct) else Cast(Convert(varChar(100), DecryptByKey(Convert(varBinary(100), MqrSalRtBin))) as Decimal(16,2))*max(MqmCnvFct) end) ", "Cast(Convert(varChar(100), DecryptByKey(Convert(varBinary(100), MqrSalRtBin))) as Decimal(16,2))") + " as MqrSalRt, '*' as MqrSetSCd, 0.00 as MqrSetSalRt, " + wRmSetValOpt + " as MqrSetSalVal, " + wRmValOpt + " as MqrSalVal, '' as RmRSzRt "
       wRmGrp = " Group by MqmCoCd, MqmYy, MqmTc, MqmChr, MqmNo, MqdSr, MqrDiaGrd, MqrRmCtg, MqrRmSCtg, Convert(varChar(100), DecryptByKey(Convert(varBinary(100), MqrRmCdBin))), RmQw, MqrLmeSal, Cast(Convert(varChar(100), DecryptByKey(Convert(varBinary(100), MqrSalRtBin))) as Decimal(16,2)) "
     Case Is = UCase("RmCd")
       wRmStr = " 0 as MqrSrNo, MqrRmCtg, MqrRmSCtg, " + wOurRmCd + ", RmQw, 0 as MqrLn1, 0 as MqrLn2,  0 as MqrLn3, 0 as MqrRmPtr, " + _
                "sum(MqrQty) as MqrQty, sum(MqrWt) as MqrWt, 0.00 as MqrLmeSal, " + wRtStr + " as MqrSalRt, '*' as MqrSetSCd, 0.00 as MqrSetSalRt, " + wRmSetValOpt + " as MqrSetSalVal, " + wRmValOpt + " as MqrSalVal, '' as RmRSzRt "
       wRmGrp = " Group by MqrPrtKey, MqmCoCd, MqmYy, MqmTc, MqmChr, MqmNo, MqdSr, MqrDiaGrd, MqrRmCtg, MqrRmSCtg, Convert(varChar(100), DecryptByKey(Convert(varBinary(100), MqrRmCdBin))), RmQw "
  
     Case Is = UCase("RmSCtg")
       wRmStr = " 0 as MqrSrNo, MqrRmCtg, MqrRmSCtg, '*' as MqrRmCd, '*' As RmQw, 0 as MqrLn1, 0 as MqrLn2, 0 as MqrLn3, 0 as MqrRmPtr, " + _
                "sum(MqrQty) as MqrQty, sum(MqrWt) as MqrWt, 0.00 as MqrLmeSal, 0.00 as MqrSalRt, '' as MqrSetSCd, 0.00 as MqrSetSalRt, " + wRmSetValOpt + " as MqrSetSalVal, " + wRmValOpt + " as MqrSalVal, '' as RmRSzRt "
       wRmGrp = " group by MqmCoCd, MqmYy, MqmTc, MqmChr, MqmNo, MqdSr, MqrDiaGrd, MqrRmCtg, MqrRmSCtg "
     Case Is = UCase("RmCtg")
       wRmStr = " 0 as MqrSrNo, MqrRmCtg, '*' as MqrRmSCtg, '*' as MqrRmCd, '*' As RmQw, 0 as MqrLn1, 0 as MqrLn2, 0 as MqrLn3, 0 as MqrRmPtr, " + _
               "sum(MqrQty) as MqrQty, sum(MqrWt) as MqrWt, 0.00 as MqrLmeSal, 0.00 as MqrSalRt, '*' as MqrSetSCd, 0.00 as MqrSetSalRt, " + wRmSetValOpt + " as MqrSetSalVal, " + wRmValOpt + " as MqrSalVal, '' as RmRSzRt "
       wRmGrp = " group by MqmCoCd, MqmYy, MqmTc, MqmChr, MqmNo, MqdSr, MqrDiaGrd, MqrRmCtg "
     Case Is = UCase("RmRSzRt")
        wRmStr = " 0 as MqrSrNo, MqrRmCtg, MqrRmSCtg, " + wOurRmCd + ",max(RmQw) As RmQw, max(MqrLn1) as MqrLn1, max(MqrLn2) as MqrLn2, max(MqrLn3) as MqrLn3, 0 as MqrRmPtr, " + _
                "sum(MqrQty) as   MqrQty, sum(MqrWt) as MqrWt, " + _
                IIF(wb_InBaseCur, " (Case when max(MqmMulDiv)= 'M' then MqrLmeSal/max(MqmCnvFct) else MqrLmeSal*max(MqmCnvFct) end) ", "MqrLmeSal") + " as MqrLmeSal, " + _
                IIF(wb_InBaseCur, " (Case when max(MqmMulDiv)= 'M' then Cast(Convert(varChar(100), DecryptByKey(Convert(varBinary(100), MqrSalRtBin))) as Decimal(16,2))/max(MqmCnvFct) else Cast(Convert(varChar(100), DecryptByKey(Convert(varBinary(100), MqrSalRtBin))) as Decimal(16,2))*max(MqmCnvFct) end) ", "Cast(Convert(varChar(100), DecryptByKey(Convert(varBinary(100), MqrSalRtBin))) as Decimal(16,2))") + " as MqrSalRt, " + _
                "MqrSetSCd , " + IIF(wb_InBaseCur, " (Case when max(MqmMulDiv)= 'M' then Cast(Convert(varChar(100), DecryptByKey(Convert(varBinary(100), MqrSetSalRtBin))) as Decimal(16,2))/max(MqmCnvFct) else Cast(Convert(varChar(100), DecryptByKey(Convert(varBinary(100), MqrSetSalRtBin))) as Decimal(16,2))*max(MqmCnvFct) end) ", "Cast(Convert(varChar(100), DecryptByKey(Convert(varBinary(100), MqrSetSalRtBin))) as Decimal(16,2))") + " as MqrSetSalRt, " + wRmSetValOpt + " as MqrSetSalVal, " + wRmValOpt + " as MqrSalVal, isnull(max(RrRngSz)," + GetRmSzDBFldNm("max(MqrLn1)", "max(MqrLn2)", "max(MqrLn3)") + ") as RmRSzRt "
       wRmGrp = " group by MqmCoCd, MqmYy, MqmTc, MqmChr, MqmNo, MqdSr, MqrDiaGrd, MqrRmCtg, MqrRmSCtg, Convert(varChar(100), DecryptByKey(Convert(varBinary(100), MqrRmCdBin))), RrRngSz, MqrLmeSal, Cast(Convert(varChar(100), DecryptByKey(Convert(varBinary(100), MqrSalRtBin))) as Decimal(16,2)), MqrSetSCd, Cast(Convert(varChar(100), DecryptByKey(Convert(varBinary(100), MqrSetSalRtBin))) as Decimal(16,2))"
     Case Is = UCase("RmSCtgRt")
       wRmStr = " 0 as MqrSrNo, MqrRmCtg, MqrRmSCtg, '*' as MqrRmCd, max(RmQw) As RmQw, 0 as MqrLn1, 0 as MqrLn2, 0 as MqrLn3, 0 as MqrRmPtr, " + _
                "sum(MqrQty) as MqrQty, sum(MqrWt) as MqrWt, " + _
                IIF(wb_InBaseCur, " (Case when max(MqmMulDiv)= 'M' then MqrLmeSal/max(MqmCnvFct) else MqrLmeSal*max(MqmCnvFct) end) ", "MqrLmeSal") + " as MqrLmeSal, " + _
                IIF(wb_InBaseCur, " (Case when max(MqmMulDiv)= 'M' then Cast(Convert(varChar(100), DecryptByKey(Convert(varBinary(100), MqrSalRtBin))) as Decimal(16,2))/max(MqmCnvFct) else Cast(Convert(varChar(100), DecryptByKey(Convert(varBinary(100), MqrSalRtBin))) as Decimal(16,2))*max(MqmCnvFct) end) ", "Cast(Convert(varChar(100), DecryptByKey(Convert(varBinary(100), MqrSalRtBin))) as Decimal(16,2))") + " as MqrSalRt, " + _
                "'' as MqrSetSCd , 0.0 as MqrSetSalRt, " + wRmSetValOpt + " as MqrSetSalVal, " + wRmValOpt + " as MqrSalVal, Space(10) as RmRSzRt "
       wRmGrp = " group by MqmCoCd, MqmYy, MqmTc, MqmChr, MqmNo, MqdSr, MqrDiaGrd, MqrRmCtg, MqrRmSCtg, RmQW, MqrLmeSal, Cast(Convert(varChar(100), DecryptByKey(Convert(varBinary(100), MqrSalRtBin))) as Decimal(16,2)) "
     End Select
  
     wRmCommStr = " Open Symmetric Key DataEnde Decryption By Password = 'RMeWm'; " + _
                " Select MqmCoCd, MqmYy, MqmTc, MqmChr, MqmNo, MqdSr, MqrDiaGrd, max(MqdDmCd) as MqdDmCd, Max(MqdDmSz) as MqdDmSz, " + _
                "(Select PValue3 from Param where PTyp= 'RMSCTG' and PMCd= max(RmCtg) and PSCd= max(RmSCtg)) as qPtrYn, " + _
                wRmStr + " " + _
                " From MultiPrcQtMst " + _
                "  join MultiPrcQtDsg on MqdMqmIdNo=MqmIdNo and MqdCoCd= MqmCoCd and MqdTc= MqmTc and MqdYy= MqmYy " + _
                " and MqdChr= MqmChr and MqdNo= MqmNo " + IIF(gs_Partition = ctCurrPrtn, " and MqdPrtKey=MqmPrtKey ", "") + _
                " join DsgMst on DmIdNo=MqdDmIdNo and DmTcTyp='DM' and DmCd= MqdDmCd and DmSz='' " + IIF(gs_Partition = ctCurrPrtn, " and DmPrtKey=MqdPrtKey ", "") + _
                " join CustMst on CmCtg='C' and CmCd= MqmCmCd " + _
                " join MultiPrcQtRm on MqrCoCd= MqdCoCd and MqrTc= MqdTc and MqrYy= MqdYy " + _
                " and MqrChr= MqdChr and MqrNo= MqdNo and MqrSr= MqdSr " + IIF(gs_Partition = ctCurrPrtn, " and MqrPrtKey=MqdPrtKey ", "") + _
                " Join RmMst On RmCd=Convert(varChar(100), DecryptByKey(Convert(varBinary(100), MqrRmCdBin))) " + IIF(gs_Partition = ctCurrPrtn, " and RmPrtKey=MqrPrtKey ", "") + _
                " Left Outer join RmRt on RrTCTyp= 'RSZ' and RrCmCtg = 'C' and RrCmCd = 'ZSELF' and RrCmCurCd= '' and " + _
                " RrCtg = MqrRmCtg and RrSCtg = MqrRmSCtg and RrSTWGrp = '' and RrCd = '' and " + _
                " RrDmCtg = '' and RrLsCtg = '' and RrLabMCd = '' and RrFrLn <= MqrLn1 and RrToLn >= MqrLn1 " + _
                wCnd + IIF(gs_Partition = ctCurrPrtn, " and MqmPrtKey='" + ctCurrPrtn + "' ", "") + wRmGrp
  
     Call moCn.RepRes(Rep.SubRep.OpenSubreport, wRmCommStr + ctMaxDopOpt)
  
     If (adc("UoYN2")) = "Y" Then
         SetSubFormula Rep.SubRep, "wActSalRt", "{rdo.MqrSalRt}"
     ElseIf (adc("UoYN2")) = "N" Then
        SetSubFormula Rep.SubRep, "wActSalRt", 0
     End If
'   RamRam
    SetSubFormula Rep.SubRep, "wLnSv", "If ({rdo.MqrRmCtg} ='G' or {rdo.MqrRmCtg} ='P' or {rdo.MqrRmCtg} ='S' or {rdo.MqrRmCtg} ='L') " + _
                                         "then (if Round({rdo.MqrLmeSal}, 3) = 0.00 Then Trim (ToText ({rdo.MqrLn1}, 3)) ) " + _
                                         "Else If {rdo.qPtrYN}='Y' Then (ToText ({rdo.MqrRmPtr}, 3)) " + _
                                         "     Else (If Round({rdo.MqrLn1}, 3)= 0.01 Then '+0 ' " + _
                                         "           else if Round({rdo.MqrLn1}, 3)= 0.02 Then '+00 ' " + _
                                         "           else if Round({rdo.MqrLn1}, 3)= 0.03 Then '+000 ' " + _
                                         "           else if Round({rdo.MqrLn1}, 3)= 0 Then ' ' " + _
                                         " else Trim (ToText ({rdo.MqrLn1}, 3))+'*'+Trim (ToText ({rdo.MqrLn2}, 3))+'*'+Trim (ToText ({rdo.MqrLn3}, 3)) ) "
' "     else Trim (ToText ({rdo.MqrLn1}, 3)) ) "
  
     SetSubFormula Rep.SubRep, "wMqrLmeSal", "If ({rdo.MqrRmCtg} ='G' or {rdo.MqrRmCtg} ='P' or {rdo.MqrRmCtg} ='S' or {rdo.MqrRmCtg} ='L') " + _
                                     "and Round({rdo.MqrLmeSal}, 3) <> 0.00 and '" + adc("UoYn14") + "'='Y' Then Round({rdo.MqrLmeSal}, 3) else 0.00 "
     
     SetSubFormula Rep.SubRep, "wGldRtDAFlg", IIF(moCn.GetFldVal("Select HGldRtDAYN from Head where HCd= '" + ctSelfCmCd + "' and HCoCd= '" + adc("UOCOCDFR") + "'") = "Y", "'D'", "'A'")
     If adc("UORMGRP") = UCase("RmRSzRt") Then
       Rep.SubRep_TxtFrLn.Suppress = True
       Rep.SubRep_TxtToLn.Suppress = True
'       Rep.SubRep_wRmSz.Suppress = True
       Rep.SubRep_MqrLn1.Suppress = True
       Rep.SubRep_MqrLn2.Suppress = True
'       Rep.SubRep_MqrLn3.Suppress = True   ' RamRam
       Rep.SubRep_TxtRngSz.Suppress = False
     Else
'       Rep.SubRep_wRmSz.Suppress = False
       Rep.SubRep_MqrLn1.Suppress = False
       Rep.SubRep_MqrLn2.Suppress = False
'       Rep.SubRep_MqrLn3.Suppress = False  ' RamRam
       Rep.SubRep_TxtRngSz.Suppress = True
       Rep.SubRep_RmRSzRt.Suppress = True
     End If
     If adc("UoYN8") = "Y" Then
       Rep.SubRep_hRmCd.SetText "Cust Quality"
     Else
       Rep.SubRep_hRmCd.SetText "Quality"
     End If
   
   End If
    
   If adc("UOYN11") = "Y" Then
     GRP_REP.SetFormula Rep, "wMqdSr", "Trim(UpperCase(ToText({rdo.MqdDmCd})))+ '/'+ ToText ({rdo.MqdSr}, '####0') "
     SetSubFormula Rep.SubRep, "wMqdSr", "Trim(UpperCase(ToText({rdo.MqdDmCd})))+ '/'+ ToText ({rdo.MqdSr}, '####0') "
   Else
    GRP_REP.SetFormula Rep, "wMqdSr", "ToText ({rdo.MqdSr}, '####0') "
    SetSubFormula Rep.SubRep, "wMqdSr", "ToText ({rdo.MqdSr}, '####0') "
   End If
  
    
  
  If adc("UoYn2") = "N" And adc("UoYn14") = "N" And adc("UoYn4") = "N" And adc("UoRmGrp") = "N" Then
    Rep.RmDet.Suppress = True
    Rep.DetSubRep.Suppress = True
    Rep.DetFt.Suppress = True
    Rep.GH2BlankSec.Suppress = True
  End If
  
  If adc("UoYN7") = "N" Then Rep.hCustDmCd.Suppress = True: Rep.FldCustDmCd.Suppress = True
  
  GRP_REP.SetFormula Rep, "wSr", IIF(adc("UoYn25") = "Y", "{@wRunSr}", "{rdo.MqdSr}")
  '578.19 calling function to change Rmcode font size
  SetSubFormula Rep.SubRep, "wRmCdFontSz", GetRmcdFontSz("{rdo.MqrRmCd}", 7)
  
  CRV_REP.DisplayGroupTree = False
      
  Call DispReport(adc)
  
  Exit Sub
RepErr:
  Call ChkSelScope(Err.Number, Err.Description)

End Sub


Private Sub ChkBoxArr_Click(Index As Integer)
With ChkBoxArr(Index)
  Select Case Index
  Case Is = 0     'Tag = "ShowPic"
    If .Value = Checked Then adc("UoYN") = "Y" Else adc("UoYN") = "N"
  Case Is = 2     'Tag = "ShowDsgVal"
    If .Value = Checked Then adc("UoYN2") = "Y" Else adc("UoYN2") = "N"
  Case Is = 4     'Tag = "ShowLabval"
    If .Value = Checked Then adc("UoYN4") = "Y" Else adc("UoYN4") = "N"
  Case Is = 7     'Tag = "ShowCustDsg"
    If .Value = Checked Then adc("UoYN7") = "Y" Else adc("UoYN7") = "N"
  Case Is = 8     'Tag = "ShowCustRm"
    If .Value = Checked Then adc("UoYN8") = "Y" Else adc("UoYN8") = "N"
  Case Is = 11     'Tag = "FinalSrt"
    If .Value = Checked Then adc("UoYN11") = "Y" Else adc("UoYN11") = "N"
  Case Is = 14          'Tag = "ShowRmVal"
    If .Value = Checked Then adc("UoYN14") = "Y" Else adc("UoYN14") = "N"
  Case Is = 25          'Tag = "ShowRunSr"
    If .Value = Checked Then adc("UoYN25") = "Y" Else adc("UoYN25") = "N"
End Select
End With
End Sub
Private Sub ChkBoxArr_GotFocus(Index As Integer)
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
  '(Bef 13/08/05)
  'SetMDISysColorsAfterLogin   '*** Jenny Color (For Setting System Colors)
  
  Dim ws_BtnStr As String
  ws_BtnStr = ""
  Call SetObjColors(Me, ws_BtnStr)
  '*** Jenny Colour
End Sub
Private Sub ADC_Load()
    Set moCn = adc.Connection    '*** (Bef speed) bef 24/11/06
    Set adc.FirNKeyCtl = adc("UoMqmTcFr")
    Call InitProp(Me)
    
    gs_CmCtg = "C"
    
      gs_RmGrp = "ORDDET"
      gs_LabGrp = "ORDDET"
      
'      gs_PTyp = "PCT"
'      gs_Tbl = "Param"
      
      
    
    
    Call SetGroupSort
End Sub
Private Sub ATXT_FldChange(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String, ByVal pv_NewValue As Variant)
  gs_CmCtg = "C"
'  If UCase(adc.MenuCd) = UCase("RepDetOrd") Or UCase(adc.MenuCd) = UCase("RepFgDet") Or UCase(adc.MenuCd) = UCase("RepJtDet") Or _
'      UCase(adc.MenuCd) = UCase("RepAvgOrd") Or UCase(adc.MenuCd) = UCase("RepFgAvg") Or UCase(adc.MenuCd) = UCase("RepJtAvg") Then
'    gs_PTyp = "PCT"
'    gs_Tbl = "Param"
'  End If
  
  Call SetProp(Me, IdName, Valid, Cancel, ErrMsg)
End Sub
Private Sub ATXT_FldWhen(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String)
  gb_CoCdFor = True
  gs_CmCtg = "C"
  
  If UCase(adc.MenuCd) = UCase("RepDetOrd") Or UCase(adc.MenuCd) = UCase("RepFgDet") Or UCase(adc.MenuCd) = UCase("RepJtDet") Or _
      UCase(adc.MenuCd) = UCase("RepAvgOrd") Or UCase(adc.MenuCd) = UCase("RepFgAvg") Or UCase(adc.MenuCd) = UCase("RepJtAvg") Then
    gs_PTyp = "PCT"
    gs_Tbl = "Param"
  End If
    
  SetProp Me, IdName, When
End Sub
Private Sub atxt_ReCalculate(Index As Integer, ByVal IdName As String, ByVal OldVal As Variant, ByVal RowNum As Integer, ByVal SourceName As String, ByVal SourceColName As String)
      'adc("UoCoCdFr") = gs_CoCd
End Sub


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
    
  If adc("UOYN4") = "" Then adc("UOYN4") = "Y"
  If adc("UoYN") = "" Then adc("UoYN") = "N"
  If adc("UoYN2") = "" Then adc("UoYN2") = "N"
  If adc("UoYN7") = "" Then adc("UoYN7") = "N"
  If adc("UoYN8") = "" Then adc("UoYN8") = "N"
  If adc("UoYN11") = "" Then adc("UoYN11") = "N"
  If adc("UoYN14") = "" Then adc("UoYN14") = "N"
  If adc("UoYN25") = "" Then adc("UoYN25") = "N"

  If adc("UoYN") = "Y" Then ChkBoxArr(0).Value = Checked Else ChkBoxArr(0).Value = Unchecked
  If adc("UoYN2") = "Y" Then ChkBoxArr(2).Value = Checked Else ChkBoxArr(2).Value = Unchecked
  If adc("UoYN4") = "Y" Then ChkBoxArr(4).Value = Checked Else ChkBoxArr(4).Value = Unchecked
  If adc("UoYN7") = "Y" Then ChkBoxArr(7).Value = Checked Else ChkBoxArr(7).Value = Unchecked
  If adc("UoYN8") = "Y" Then ChkBoxArr(8).Value = Checked Else ChkBoxArr(8).Value = Unchecked
  If adc("UoYN11") = "Y" Then ChkBoxArr(11).Value = Checked Else ChkBoxArr(11).Value = Unchecked
  If adc("UoYN14") = "Y" Then ChkBoxArr(14).Value = Checked Else ChkBoxArr(14).Value = Unchecked
  If adc("UoYN25") = "Y" Then ChkBoxArr(25).Value = Checked Else ChkBoxArr(25).Value = Unchecked
  
  
End Sub
Private Sub DispCoNm()
    Dim ws_HName As String
    ws_HName = moCn.GetFldVal("Select HName from Head where HCoCd = '" + adc("UoCoCdFr") + "' and " + _
                                                      "HCd = '" + ctSelfCmCd + "' ")
    
      Rep.wCoCd.SetText ws_HName
      Rep.wCoCdLogo.SetText ws_HName
  
  If moCn.RecSeek("Select 'x' From Head Where HCoCd='" + gs_CoCd + "' and HCd='" + ctSelfCmCd + "' and HShowRepLogo='Y' ") Then
    Rep.RHMain.UnderlaySection = True: Rep.wCoCd.Suppress = True
  Else
    Rep.RHLogo.Suppress = True
  End If
  
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
    
    Set moCn = Nothing
    Set Rep = Nothing
    Set Rep = Nothing

End Sub

Private Function MakeStr(wsPassStr As String) As String
Dim i As Integer, NewStr As String
  
  i = 1
  Do While i <= Len(wsPassStr)
    If Mid(wsPassStr, i, 1) <> Chr(13) Then
      NewStr = NewStr + Mid(wsPassStr, i, 1)
    Else
      NewStr = NewStr + " "
    End If
    i = i + 1
  Loop
  MakeStr = NewStr
End Function






