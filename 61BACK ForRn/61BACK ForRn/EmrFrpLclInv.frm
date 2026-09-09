VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{AB2CC32A-E652-42A2-B341-26ECE02F6386}#2.1#0"; "mwfctlforspeed9.ocx"
Begin VB.Form EmrFrpLclInv 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Local Invoice"
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
      TabIndex        =   32
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
      TabIndex        =   31
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
      TabIndex        =   39
      Top             =   9570
      Width           =   3555
      Begin MwfCtl.MWCTL_MED aktxt 
         Height          =   345
         Index           =   0
         Left            =   210
         TabIndex        =   40
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
         TabIndex        =   41
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
      TabIndex        =   38
      Top             =   9585
      Visible         =   0   'False
      Width           =   855
      Begin CRVIEWER9LibCtl.CRViewer9 CRV_REP 
         Height          =   495
         Left            =   480
         TabIndex        =   36
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
      TabIndex        =   0
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
      BackColor       =   &H00C0C0C0&
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
      TabIndex        =   35
      Top             =   0
      Width           =   15255
      Begin TabDlg.SSTab TAB_REP 
         Height          =   9225
         Left            =   -2400
         TabIndex        =   34
         Top             =   120
         Width           =   15375
         _ExtentX        =   27120
         _ExtentY        =   16272
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "&Data Range"
         TabPicture(0)   =   "EmrFrpLclInv.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "fra_tabrep(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "&Group&Sort"
         TabPicture(1)   =   "EmrFrpLclInv.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "fra_tabrep(1)"
         Tab(1).ControlCount=   1
         Begin VB.Frame fra_tabrep 
            BorderStyle     =   0  'None
            Height          =   8115
            Index           =   0
            Left            =   2640
            TabIndex        =   44
            Top             =   600
            Width           =   12465
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
               Index           =   12
               Left            =   135
               TabIndex        =   25
               Tag             =   "ShowHUID"
               ToolTipText     =   "Check To Show Bag / Stock Id "
               Top             =   5700
               Width           =   3240
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Show Value"
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
               Left            =   135
               TabIndex        =   24
               Tag             =   "ShowValue"
               ToolTipText     =   "Check To Show Value"
               Top             =   5400
               Width           =   3240
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Show HUID"
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
               Left            =   135
               TabIndex        =   23
               Tag             =   "ShowHUID"
               ToolTipText     =   "Check To Show HUID Number"
               Top             =   5100
               Width           =   3240
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Show IRN / QR Code"
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
               Left            =   135
               TabIndex        =   22
               Tag             =   "Show IRN/QR Code"
               ToolTipText     =   "Check To Show IRN and QR Code"
               Top             =   4800
               Width           =   3240
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Show MRP"
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
               Left            =   130
               TabIndex        =   21
               Tag             =   "ShowMRP"
               ToolTipText     =   "Check To Show MRP Details"
               Top             =   4500
               Width           =   3240
            End
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
               Index           =   7
               Left            =   130
               TabIndex        =   20
               Tag             =   "ShowDCWt"
               ToolTipText     =   "Check To Show Logos"
               Top             =   4200
               Width           =   3240
            End
            Begin VB.Frame Frame1 
               Height          =   615
               Left            =   3160
               TabIndex        =   63
               Top             =   5895
               Width           =   3255
               Begin VB.OptionButton OptOldRpt 
                  Caption         =   "Old Format"
                  Height          =   255
                  Left            =   1800
                  TabIndex        =   27
                  Top             =   240
                  Width           =   1335
               End
               Begin VB.OptionButton OptNewRpt 
                  Caption         =   "New Format"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   26
                  Top             =   240
                  Width           =   1455
               End
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Show Demarcation                   "
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
               Left            =   130
               TabIndex        =   10
               Tag             =   "ShowDemarc"
               ToolTipText     =   "Check To Show Demarcation"
               Top             =   1290
               Width           =   3240
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Show Run Sr                             "
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
               Left            =   130
               TabIndex        =   11
               Tag             =   "ShowRunSr"
               ToolTipText     =   "Check To Show Running Serial"
               Top             =   1575
               Width           =   3240
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Sort By Design                          "
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
               Left            =   130
               TabIndex        =   12
               Tag             =   "SortByDmCd"
               ToolTipText     =   "Check To Sort By Desing Code"
               Top             =   1860
               Width           =   3240
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Show Picture                            "
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
               Left            =   130
               TabIndex        =   13
               Tag             =   "SortByDmCd"
               ToolTipText     =   "Check To Show Picture"
               Top             =   2160
               Width           =   3240
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Show Rm Summary                  "
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
               Left            =   130
               TabIndex        =   14
               Tag             =   "ShowRmSumm"
               ToolTipText     =   "Check To Show Rm Summary"
               Top             =   2430
               Width           =   3240
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Show Cust PO/ Sub PO       "
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
               Left            =   130
               TabIndex        =   15
               Tag             =   "ShowOdPO"
               ToolTipText     =   "Check To Show Cust PO/Sub PO"
               Top             =   2730
               Width           =   3240
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Show Design Suffix"
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
               Left            =   130
               TabIndex        =   16
               Tag             =   "ShowDsgSfx"
               ToolTipText     =   "Check To Show Design Suffix"
               Top             =   3015
               Width           =   3240
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   26
               Left            =   5145
               TabIndex        =   5
               ToolTipText     =   "Enter Invoice Number"
               Top             =   1005
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOINNOFR"
               IdName          =   "UOINNOFR"
               CmpStr          =   "InNo >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   20
               Left            =   3945
               TabIndex        =   3
               ToolTipText     =   "Enter Invoice Year"
               Top             =   1005
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               MaxLength       =   2
               DataType        =   4
               DataField       =   "UOINYYFR"
               IdName          =   "UOINYYFR"
               CmpStr          =   "InYy >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   43
               Left            =   3165
               TabIndex        =   2
               ToolTipText     =   "Enter Invoice Tc"
               Top             =   1005
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOINTCFR"
               IdName          =   "UOINTCFR"
               CmpStr          =   "InTc >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   45
               Left            =   4365
               TabIndex        =   4
               ToolTipText     =   "Enter Invoice Character"
               Top             =   1005
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOINCHRFR"
               IdName          =   "UOINCHRFR"
               CmpStr          =   "InChr >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   17
               Left            =   3165
               TabIndex        =   1
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
               CmpStr          =   "InCoCd = "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   0
               Left            =   6780
               TabIndex        =   6
               ToolTipText     =   "Enter Invoice Tc"
               Top             =   1005
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOINTCTO"
               IdName          =   "UOINTCTO"
               CmpStr          =   "InTc <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   1
               Left            =   7560
               TabIndex        =   7
               ToolTipText     =   "Enter Invoice Year"
               Top             =   1005
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               MaxLength       =   2
               DataType        =   4
               DataField       =   "UOINYYTO"
               IdName          =   "UOINYYTO"
               CmpStr          =   "InYy <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   2
               Left            =   7980
               TabIndex        =   8
               ToolTipText     =   "Enter Invoice Character"
               Top             =   1005
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOINCHRTO"
               IdName          =   "UOINCHRTO"
               CmpStr          =   "InChr <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   3
               Left            =   8760
               TabIndex        =   9
               ToolTipText     =   "Enter Invoice Number"
               Top             =   1005
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOINNOTO"
               IdName          =   "UOINNOTO"
               CmpStr          =   "InNo <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   4
               Left            =   3555
               TabIndex        =   45
               ToolTipText     =   "Show Demarcation Line(Y/ N)"
               Top             =   1290
               Visible         =   0   'False
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               Enabled         =   0   'False
               DataType        =   4
               DataField       =   "UOYN3"
               IdName          =   "UOYN3"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   5
               Left            =   3555
               TabIndex        =   46
               ToolTipText     =   "Show Running Sr(Y/ N)"
               Top             =   1575
               Visible         =   0   'False
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               Enabled         =   0   'False
               DataType        =   4
               DataField       =   "UOYN"
               IdName          =   "UOYN"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   6
               Left            =   3555
               TabIndex        =   47
               ToolTipText     =   "Show Running Sr(Y/ N)"
               Top             =   1860
               Visible         =   0   'False
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               Enabled         =   0   'False
               DataType        =   4
               DataField       =   "UOYN1"
               IdName          =   "UOYN1"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   7
               Left            =   3555
               TabIndex        =   48
               ToolTipText     =   "Show Running Sr(Y/ N)"
               Top             =   2145
               Visible         =   0   'False
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               Enabled         =   0   'False
               DataType        =   4
               DataField       =   "UOYN2"
               IdName          =   "UOYN2"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   8
               Left            =   3555
               TabIndex        =   49
               ToolTipText     =   "ShowRm Summary (Y/ N)"
               Top             =   2430
               Visible         =   0   'False
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               Enabled         =   0   'False
               DataType        =   4
               DataField       =   "UOYN4"
               IdName          =   "UOYN4"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   9
               Left            =   3165
               TabIndex        =   17
               ToolTipText     =   "Enter Level Of Raw Material Detail To Be Shown"
               Top             =   3285
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
               Index           =   10
               Left            =   3555
               TabIndex        =   50
               ToolTipText     =   "Show Customer PO /Sub PO (Y/ N)"
               Top             =   2715
               Visible         =   0   'False
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               Enabled         =   0   'False
               DataType        =   4
               DataField       =   "UOYN5"
               IdName          =   "UOYN5"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   11
               Left            =   3555
               TabIndex        =   51
               ToolTipText     =   "Show Dsg Sfx  (Y/ N)"
               Top             =   3000
               Visible         =   0   'False
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               Enabled         =   0   'False
               DataType        =   4
               DataField       =   "UOYN6"
               IdName          =   "UOYN6"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   12
               Left            =   3165
               TabIndex        =   18
               ToolTipText     =   "Enter Value On Which Invoice Report Grouping Has To Be Done"
               Top             =   3600
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
               Index           =   13
               Left            =   3165
               TabIndex        =   19
               ToolTipText     =   "Enter Detail Sorting To Be Shown In Report"
               Top             =   3885
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
               Index           =   25
               Left            =   3600
               TabIndex        =   65
               ToolTipText     =   "Show Diamond/ Color Stone Weight (Yes / No)"
               Top             =   4200
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
               Index           =   84
               Left            =   3600
               TabIndex        =   66
               ToolTipText     =   "Show MRP Details (Yes / No)"
               Top             =   4500
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
               Index           =   14
               Left            =   3600
               TabIndex        =   67
               ToolTipText     =   "Show IRN and QR Code (Yes / No)"
               Top             =   4800
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
               Index           =   15
               Left            =   3600
               TabIndex        =   68
               ToolTipText     =   "Show HUID (Yes / No)"
               Top             =   5100
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
               Index           =   88
               Left            =   3600
               TabIndex        =   69
               ToolTipText     =   "Check To Show Value "
               Top             =   5400
               Visible         =   0   'False
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               Enabled         =   0   'False
               DataType        =   4
               IdName          =   "WSHOWVALYN"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   16
               Left            =   3600
               TabIndex        =   70
               ToolTipText     =   "Show Bag / Stock Id (Yes / No)"
               Top             =   5700
               Visible         =   0   'False
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               Enabled         =   0   'False
               DataType        =   4
               DataField       =   "UOYN12"
               IdName          =   "UOYN12"
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
               TabIndex        =   30
               Top             =   3885
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
               TabIndex        =   29
               Top             =   3600
               Width           =   1845
            End
            Begin VB.Label LblRepFormat 
               BackStyle       =   0  'Transparent
               Caption         =   "Show Report "
               BeginProperty Font 
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
               TabIndex        =   64
               ToolTipText     =   "Location"
               Top             =   6045
               Width           =   1845
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
               TabIndex        =   62
               Top             =   1005
               Width           =   105
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
               TabIndex        =   61
               Top             =   720
               Width           =   1845
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
               TabIndex        =   60
               Top             =   1005
               Width           =   1845
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
               TabIndex        =   59
               Top             =   1005
               Width           =   105
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
               TabIndex        =   58
               Top             =   1005
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
               Left            =   7500
               TabIndex        =   57
               Top             =   1005
               Width           =   105
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
               Left            =   7920
               TabIndex        =   56
               Top             =   1005
               Width           =   105
            End
            Begin VB.Label Label3 
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
               Left            =   8700
               TabIndex        =   55
               Top             =   1005
               Width           =   105
            End
            Begin VB.Label Label4 
               BackStyle       =   0  'Transparent
               Caption         =   "Show Demarcation Line"
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
               Left            =   150
               TabIndex        =   54
               Top             =   1260
               Visible         =   0   'False
               Width           =   1845
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
               Left            =   3165
               TabIndex        =   53
               Top             =   180
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
               Left            =   6780
               TabIndex        =   52
               Top             =   180
               Width           =   1275
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
               Index           =   0
               Left            =   165
               TabIndex        =   28
               ToolTipText     =   "Location"
               Top             =   3300
               Width           =   1845
            End
         End
         Begin VB.Frame fra_tabrep 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   9070
            Index           =   1
            Left            =   -72480
            TabIndex        =   37
            Top             =   480
            Width           =   8745
            Begin MwfCtl.MWCTL_GRP GRP_REP 
               Height          =   360
               Left            =   240
               TabIndex        =   33
               Top             =   360
               Width           =   8745
               _ExtentX        =   15425
               _ExtentY        =   635
            End
         End
      End
   End
   Begin VB.Label Label7 
      Caption         =   "New Report"
      Height          =   255
      Left            =   11280
      TabIndex        =   43
      Top             =   7440
      Width           =   1695
   End
   Begin VB.Label LblRmGrp 
      BackStyle       =   0  'Transparent
      Caption         =   " RM Det Level"
      BeginProperty Font 
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
      Left            =   9840
      TabIndex        =   42
      ToolTipText     =   "Location"
      Top             =   6720
      Width           =   1845
   End
End
Attribute VB_Name = "EmrFrpLclInv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mRep As CRAXDRT.Report
Dim Rep As New EmrRepLclInv
Dim moCn As MwfLib.MDOConnection   '*** (Bef speed) bef 24/11/06
Dim wFlagLocalExp As String
Dim mb_MultiInv As Boolean

Private Sub SetGroupSort()
  ''GRP_REP.Add "Sub Ctg", "RmSCtg", "", "RmSCtg", "hRmSCtg", "", "", "", "(Select PDesc From Param Where PTyp= 'RMSCTG' and PMCd = RmCtg and PSCd= RmSCtg) "
End Sub

Private Sub adc_setreprecsource()
'On Error GoTo RepErr
'*** Report Sql ***
'Select InTc, InYy, InChr, InNO, IdSr, InDt, InExpNo, InFOB,InTot, InCmName, InCmAdd1, InCmAdd2, InCmAdd3, InCmTel,
'InCmFax, InCocd as CmBillCo, InSelfName, InSelfAdd1, InSelfAdd2, InSelfAdd3, InSelfTel, InSelfFax, InHCustRem, InHUnResolutions,
'InLclInvTyp, InHCstNo, InHVatNo, InHVatPer, InLclVatRs, InHCstPer, InLclCstRs, InCmOctroiPer, InLclOctroiRs,
'InLclOthrValDesc, InLclOthrValRs, InHLclInvFT, InHLclInvVATFT, InHAuthSgn, CmCurCd as qCurCd, IdDmCd, IdQty, IdIFob, IdKt,
'IrRmIWt As qGldWt, IrRmQty as qGldQty, IrRmIWt As qPlWt, IrRmQty as qPlQty, IrRmIWt As qSlWt, IrRmQty as qSlQty,
'IrRmIWt As qPdWt, IrRmQty as qPdQty, IrRmIWt As qOthMetWt, IrRmQty as qOthMetQty,
'IrRmIWt As qDiaWt, IrRmQty as qDiaQty, IrRmIWt As qCSWt, IrRmQty as qCSQty, IrRmIWt As qAccWt, IrRmQty as qAccQty,
'IrRmIWt As qGrsWt
'From InvDsg, InvRm, InvHd, RmMst
'Where 1 = 2
'*** Report Sql ***
  
  Dim wrepcnd As String, wCnd As String, wSqlStrg As String, wJoinWhr As String
  '*** Jay 3.3(FG)
  Dim ws_IdSr, ws_IdQty As String, Ws_IdKt As String, ws_IdDmCd As String, ws_IdIFob As String, ws_IrRmIWt As String, ws_IrRmQty As String
  '*** Jay 3.3(FG)
  
  Dim wDsgSrt As String     ' ***** Manali 3.8.0 - Sort By Design
  
  Call GetCnd
  
  wrepcnd = adc.RepCond
  'wCnd = IIF(wrepcnd <> "", " and " + wrepcnd, " ")
  wCnd = wrepcnd
    
  Set Rep = Nothing
  
  Set adc.RepSource = Rep
  Set mRep = Rep
  
  '*** Jay 3.3(FG)
  ' *** Bef Trading Module
  'If UCase(adc.MenuCd) <> UCase("RepFgLclInv") Then
  ' **** Manali Trading Module
  Dim ws_IdVaCtg As String, ws_IdGldLs As String
  If UCase(adc.MenuCd) = UCase("RepLclInv") Or UCase(adc.MenuCd) = UCase("RepInvLclRM") Then
    ws_IdSr = "IdSr": ws_IdQty = "max(IdQty)": Ws_IdKt = "OdKt": ws_IdDmCd = "IdDmCd": ws_IdIFob = "max(IdIFob)"
    ws_IrRmIWt = "InvRm.IrRmIWt": ws_IrRmQty = "InvRm.IrRmQty"
    'hk.5 - Vactg and Glodls mentione for invoice
    ws_IdVaCtg = "IdVaCtg": ws_IdGldLs = "IdGldLs"
  Else
    ws_IdSr = "OdSr": ws_IdQty = "Max(OdOrdQty)": Ws_IdKt = "OdKt": ws_IdDmCd = "OdDmCd": ws_IdIFob = "Max(OdOrdQty * OdSalPrc)" 'ws_IdIFob = "OdSalPrc"
    ws_IrRmIWt = "OdOrdQty * OrdRm.OrWt": ws_IrRmQty = "OdOrdQty * OrdRm.OrQty"
    'ws_IrRmIWt = "OrWt": ws_IrRmQty = "OrQty"
    'hk.5 - Vactg and Glodls mentione for order
    ws_IdVaCtg = "OdVaCtg": ws_IdGldLs = "ParamForVACtg.PNum"
  End If
  '*** Jay 3.3(FG)
  
  ' **** Manali 3.7.1 - Base/Ref Vch Column in Report
  Dim wBaseRefIdKey As String, wBaseRefVch As String
  If UCase(adc.MenuCd) = UCase("RepLclInv") Or UCase(adc.MenuCd) = UCase("RepInvLclRM") Then
    wBaseRefVch = " '' as qBaseRefVch, "
  Else
    wBaseRefIdKey = moCn.GetFldVal("Select (Case When TmBaseTcs<> '' Then 'OdBaseIdKey' When TmRefTcs <> '' Then 'OdRefIdKey' Else '' End) as qBaseRefIdKey " + _
                                  "From TcMst Where TmCoCd='" + adc("UoCoCdFr") + "' and TmTc='" + adc("UoInTcFr") + "' ")
    
    If wBaseRefIdKey <> "" Then
      wBaseRefVch = " (Select B.OdTc+'/'+B.OdYy+'/'+B.OdChr+'/'+LTrim(Str(B.OdNo))+'/'+LTrim(Str(B.OdSr)) From OrdDsg B Where B.OdCoCd=max(Od.OdCoCd) and " + _
                  "B.OdIdNo=max(Od." + wBaseRefIdKey + ")) as qBaseRefVch, "
    Else
      wBaseRefVch = " '' as qBaseRefVch, "
    End If
  End If
  ' **** Manali 3.7.1 - Base/Ref Vch Column in Report
  
  Call DispCoNm
  
  
  Dim wDetArr() As String, wGrpArr() As String, wGrpDescFormula As String, i As Integer
  'Hk.5 Group Combination and Detail Level scope newly added for Local Invoice
  Dim wDetFormula As String, wDetGrp As String, wDetDescFormula As String, wGrpStr As String
  Dim wMainMetJoin As String
    
  'hk.5 if old format option is selected then no chnage in report output, so group sory by default serial no.
  If (UCase(adc.MenuCd) = UCase("RepInvLclRM") Or OptOldRpt.Value = True) Then wDetGrp = "IdSr"
  If (UCase(adc.MenuCd) = UCase("RepJtLclInv") And OptOldRpt.Value = True) Then wDetGrp = "OdSr"
  'hk.5 if new format is selected then looks for group sort and detail section
  If (UCase(adc.MenuCd) = UCase("RepLclInv") Or UCase(adc.MenuCd) = UCase("RepJtLclInv")) And OptNewRpt.Value = True Then
    wGrpArr = Split(moCn.GetFldVal("Select PDesc225 from Param where PTyp= 'INGR' and " + _
         "PMCd= '" + adc("UoInGrp") + "' "), ",")
     
    For i = 0 To UBound(wGrpArr)
      'hk.5 if option is rmctg then it gets description from param, and group by Rmctg,
      ' and looking for main metal and its description
      If UCase(wGrpArr(i)) = "RMCTG" Then
        wGrpDescFormula = wGrpDescFormula + IIF(wGrpDescFormula = "", "", "+', '+") + "max(ParamForRmCtg.PDesc)"
        wGrpStr = wGrpStr + IIF(wGrpStr = "", "", ",") + "ParamForRmCtg.PMCd"
        If UCase(adc.MenuCd) = UCase("RepLclInv") Then
          wMainMetJoin = "join InvRm IrMnMtl on IrMnMtl.IrCoCd= InvRm.IrCoCd and IrMnMtl.IrTc= InvRm.IrTc and IrMnMtl.IrYy= InvRm.IrYy and IrMnMtl.IrChr= InvRm.IrChr " + _
                       "and IrMnMtl.IrNo= InvRm.IrNo and IrMnMtl.IrSr= InvRm.IrSr  and IrMnMtl.IrPrtKey=InvRm.IrPrtKey and IrMnMtl.IrMainMet ='Y' " + _
                       "join Param ParamForRmCtg on ParamForRmCtg.PTyp ='RMCTG' and ParamForRmCtg.PMCd = IrMnMtl.IrRmCtg "
        Else
          wMainMetJoin = "join OrdRm OrMnMtl on OrMnMtl.OrCoCd= OrdRm.OrCoCd and OrMnMtl.OrTc= OrdRm.OrTc and OrMnMtl.OrYy= OrdRm.OrYy and OrMnMtl.OrChr= OrdRm.OrChr " + _
                        "and OrMnMtl.OrNo= OrdRm.OrNo and OrMnMtl.OrSr= OrdRm.OrSr  and OrMnMtl.OrPrtKey=OrdRm.OrPrtKey and OrMnMtl.OrMainMet ='Y' " + _
                        "join Param ParamForRmCtg on ParamForRmCtg.PTyp ='RMCTG' and ParamForRmCtg.PMCd = OrMnMtl.OrRmCtg "
        End If
      'hk.5 if option is Kt then it gets description from param, and group by OdKt,
      ElseIf UCase(wGrpArr(i)) = "KT" Then
        wGrpDescFormula = wGrpDescFormula + IIF(wGrpDescFormula = "", "", "+', '+") + "max(ParamForKt.PDesc)"
        wGrpStr = wGrpStr + IIF(wGrpStr = "", "", ",") + "OdKt"
      'hk.5 if option is DmCol then it gets description from param, and group by OdCol,
      ElseIf UCase(wGrpArr(i)) = "DMCOL" Then
        wGrpDescFormula = wGrpDescFormula + IIF(wGrpDescFormula = "", "", "+', '+") + "max(ParamForDmCol.PDesc)"
        wGrpStr = wGrpStr + IIF(wGrpStr = "", "", ",") + "OdDmCol"
      'hk.5 if option is Vactg then it gets description from param, and group by invoice vactg/order vactg,
      ElseIf UCase(wGrpArr(i)) = "VACTG" Then
         wGrpDescFormula = wGrpDescFormula + IIF(wGrpDescFormula = "", "", "+', '+") + "max(ParamForVACtg.PDesc)"
         wGrpStr = wGrpStr + IIF(wGrpStr = "", "", ",") + ws_IdVaCtg
      'hk.5 if option is GldLs then it gets value and group by from IdGldLs, for order it gets Gldls value from param
      ElseIf UCase(wGrpArr(i)) = "GLDLS" Then
         wGrpDescFormula = wGrpDescFormula + IIF(wGrpDescFormula = "", "", "+', '+") + "'@'+convert(varchar,(convert(decimal(7,3), " + ws_IdGldLs + ")))+ '%'"
         wGrpStr = wGrpStr + IIF(wGrpStr = "", "", ",") + ws_IdGldLs
      'hk.5 if option is Uom then it show the value and group by DmUom
      ElseIf UCase(wGrpArr(i)) = "PCS/PRS" Then
         wGrpDescFormula = wGrpDescFormula + IIF(wGrpDescFormula = "", "", "+', '+") + "DmUom"
         wGrpStr = wGrpStr + IIF(wGrpStr = "", "", ",") + "DmUom"
      'hk.5 if option is DmCtg then it gets description from param, and group by dmctg
      ElseIf UCase(wGrpArr(i)) = "DMCTG" Then
         wGrpDescFormula = wGrpDescFormula + IIF(wGrpDescFormula = "", "", "+', '+") + "max(ParamForDmctg.PDesc)"
         wGrpStr = wGrpStr + IIF(wGrpStr = "", "", ",") + "DmCtg"
      'hk.5 if option is SalCtg then it gets description from param, and group by SalCtg,
      ElseIf UCase(wGrpArr(i)) = "SALCTG" Then
         wGrpDescFormula = wGrpDescFormula + IIF(wGrpDescFormula = "", "", "+', '+") + "max(ParamForSalCtg.PDesc)"
         wGrpStr = wGrpStr + IIF(wGrpStr = "", "", ",") + "DmSalCtg"
      End If
    Next i
    
    'hk.5 it getting groupby fields from param wher ptyp=INDET
    If adc("UoInDet") <> "" Then
      wDetGrp = moCn.GetFldVal("Select PDesc225 from Param where PTyp= 'INDET' and " + _
            "PMCd= '" + adc("UoInDet") + "' ")
      wDetArr = Split(moCn.GetFldVal("Select PDesc225 from Param where PTyp= 'INDET' and " + _
            "PMCd= '" + adc("UoInDet") + "' "), ",")
      'hk.5 -Giving specific fields for order
      If UCase(adc.MenuCd) = UCase("RepJtLclInv") Then
        If adc("UoInDet") = "DMCD" Then
          wDetGrp = "OdDMCd"
        ElseIf adc("UoInDet") = "DMCD/SR" Then
          wDetGrp = "OdDMCd,OdSr"
        ElseIf adc("UoInDet") = "INVSR" Or adc("UoInDet") = "INVSR/SR" Or adc("UoInDet") = "TRNO/SR" Then
          wDetGrp = "OdSr"
        End If
      End If
    Else
      wDetGrp = ws_IdSr
    End If
  
    'hk.5 displaying fields for newly added detail section
    If UCase(adc("UoInDet")) = "DMCD" Then
      wDetDescFormula = "{rdo.IdDMCd}"
    ElseIf UCase(adc("UoInDet")) = "DMCD/SR" Then
      wDetDescFormula = "Trim(ToText({@wRunSr}, '###0'))+'   '+Trim({rdo.IdDmCd})+ " + _
                        "(If {rdo.IdDmSz}= '' then '' else ' / ') +Trim({rdo.IdDmSz})"
    
    ElseIf UCase(adc("UoInDet")) = "DMCTG" Then
      wDetDescFormula = "{rdo.DmCtgDesc}"
    ElseIf UCase(adc("UoInDet")) = "INVSR" Then
      wDetDescFormula = "Trim(ToText({rdo.IdSr}, '###0'))+'   '+ Trim({rdo.IdDmCd} )+ " + _
                        "(If {rdo.IdDmSz}= '' then '' else ' / ') +Trim({rdo.IdDmSz})"
    
    ElseIf UCase(adc("UoInDet")) = "INVSR/SR" Then
      wDetDescFormula = "Trim(ToText({rdo.IdSr}, '###0'))+'   ' +Trim(ToText({@wRunSr}, '###0'))+'   '+ " + "Trim({rdo.IdDmCd} )+ " + _
                        "(If {rdo.IdDmSz}= '' then '' else ' / ') +Trim({rdo.IdDmSz})"
    
    ElseIf UCase(adc("UoInDet")) = "PCS/PRS" Then
      wDetDescFormula = "{rdo.DMUom}"
    ElseIf UCase(adc("UoInDet")) = "TRNO/SR" Then
      wDetDescFormula = "Trim(ToText({rdo.IDTrayNo}, '###0'))+'   '+Trim(ToText({@wRunSr}, '###0'))+'   '+ Trim({rdo.IdDmCd} )+ " + "(If {rdo.IdDmSz}= '' then '' else ' / ') +Trim({rdo.IdDmSz})"
    End If
  
    'hk.5 if detail section <>invsr , it adding qty and value
    'sv.51 ws_IdQty, ws_IdIFob changed for jri,jro
    'ws_IdQty = IIF((UCase(ADC.MenuCd) = UCase("RepLclInv")), "sum(case when InvRm.IrMainMet ='Y' then IdQty else 0 end)", "sum(case when OrdRm.OrMainMet ='Y' then OdOrdQty else 0 end)")
    ws_IdQty = IIF((UCase(adc.MenuCd) = UCase("RepLclInv")), "sum(case when InvRm.IrMainMet ='Y' then IdQty else 0 end)", "case when OmTc='JRI' Or OmTc='JRO' then max(OdOrdQty) else (sum(case when OrdRm.OrMainMet ='Y' then OdOrdQty else 0 end)) end")
    'ws_IdIFob = IIF((UCase(ADC.MenuCd) = UCase("RepLclInv")), "sum(case when InvRm.IrMainMet ='Y' then IdIFob else 0 end)", "sum(case when OrdRm.OrMainMet ='Y' then (OdOrdQty * OdSalPrc) else 0 end)")
    ws_IdIFob = IIF((UCase(adc.MenuCd) = UCase("RepLclInv")), "sum(case when InvRm.IrMainMet ='Y' then IdIFob else 0 end)", "case when OmTc='JRI' Or OmTc='JRO' then max(OdOrdQty * OdSalPrc) else (sum(case when OrdRm.OrMainMet ='Y' then (OdOrdQty * OdSalPrc) else 0 end)) end")
      
  End If
  
  '*** Jay 3.3(FG) [Variables introduced]
  ' ****** Manali 3.03 - 05/07/08 - PdMod - 'L' included below
  ' ***** Manali 3.7.0 - Trading Module - OdDmCol, InPrdDesc addded
  ' ***** Manali 3.9.0 - InPrnDt instead of InDt
  ' ***** Manali 3.10.0 - 09/03/12 - Oth Met Wt, Qty flds added
  ' ****** Manali 3.10.0 - 24/03/12 - Self TinNo, PAN and Cust VAT, CST and PAN Added
  ' ****** Manali 3.10.0 - 30/03/12 - Picture in Report - IdDmSz, OdPicNm added
  'Diagold.2 printing state code and state name
  'sv.63 max(InIRN) as InIRN added
  If UCase(adc.MenuCd) = UCase("RepLclInv") Or UCase(adc.MenuCd) = UCase("RepInvLclRM") Then
  wSqlStrg = "Select Max(IdVaCtg+' '+ParamForVACtg.PDesc+' '+'HSN '+ParamForVACtg.PDesc225) as qGrp2, InTc, InYy, InChr, InNO, Max(" + ws_IdSr + ") as IdSr, Max(InPrnDt) as InDt,  " + _
    "Max(InExpNo) as InExpNo, Max(InFOB) as InFOB, Max(InTot) as InTot, Max(InCmCd) as InCmCd, Max(InCmName) as InCmName, " + _
    "Max(InCmAdd1) as InCmAdd1, Max(InCmAdd2) as InCmAdd2, Max(InCmAdd3) as InCmAdd3, Max(InCmTel) as InCmTel, " + _
    "Max(InCmFax) as InCmFax, Max(CmBillCo) as CmBillCo, Max(InSelfName) as InSelfName, Max(InSelfAdd1) as InSelfAdd1,  " + _
    "Max(InSelfAdd2) as InSelfAdd2, Max(InSelfAdd3) as InSelfAdd3, Max(InSelfTel) as InSelfTel,  " + _
    "Max(InSelfFax) as InSelfFax, Max(InHCustRem) as InHCustRem, Max(SubString(InHUnResolutions, 1, 750)) as  " + _
    "InHUnResolutions, Max(InLclInvTyp) as InLclInvTyp, Max(InHCstNo) as InHCstNo, Max(InHVatNo) as InHVatNo, " + _
    "Max(InHVatPer) as InHVatPer, Max(InLclVatRs) as InLclVatRs, Max(InHCstPer) as InHCstPer, " + _
    "Max(InLclCstRs) as InLclCstRs, Max(InCmOctroiPer) as InCmOctroiPer, Max(InLclOctroiRs) as InLclOctroiRs, " + _
    "Max(InLclOthrValDesc) as InLclOthrValDesc, Max(InLclOthrValRs) as InLclOthrValRs, Max(InHLclInvFT) as InHLclInvFT, " + _
    "Max(SubString((Case when InLclCstRs <> 0 then InHLclInvCstFT when InLclVatRs <> 0 then InHLclInvVatFT else '' end), 1, 750)) as InHLclInvVATFT, " + _
    "Max(InHAuthSgn) as InHAuthSgn, Max(CmCurCd) as qCurCd, " + _
    "(select PValue from Param Where PTYp = 'CURNCY' and PMCd= Max(CmCurCd) and PSCd= '') as qCurUnit,  " + _
    "/* (Select DmCtg from DsgMst where DmTcTyp = 'DM' and DmCd= Max(" + ws_IdDmCd + ") and DmSz= '') as qDmCtg, */" + _
    "Max(OdDmCol) as OdDmCol, Replace(Max(InPrdDesc),char(13)+ char(10),' ')  as InPrdDesc, Max(SubString(InTermCond, 1, 1500)) as InTermCond, Max(InCmPayTerms) as InCmPayTerms, max(InHVessel) as InHVessel, max (InAmtDesc1) as InAmtDesc1," + _
    "max (InAmtDesc2) as InAmtDesc2, max (InAmtDesc3) as InAmtDesc3, max (InAmtDesc4) as InAmtDesc4, max (InAmtDesc5) as InAmtDesc5, max (InAmtDesc6) as InAmtDesc6, max (InAmtDesc7) as InAmtDesc7, max (InAmtDesc8) as InAmtDesc8, max (InAmtDesc9) as InAmtDesc9, max (InAmtDesc10) as InAmtDesc10, " + _
   "max (InAmtDesc11) as InAmtDesc11, max (InAmtDesc12) as InAmtDesc12, max (InAmt1Per)  as InAmt1Per , max (InAmt2Per) as InAmt2Per, max (InAmt3Per) as InAmt3Per, max (InAmt4Per) as InAmt4Per, max (InAmt5Per) as InAmt5Per, max (InAmt6Per) as InAmt6Per, max (InAmt8Per) as InAmt8Per," + _
   "max (InAmt9Per) as InAmt9Per,  max (InAmt10Per) as InAmt10Per, max (InAmt11Per) as InAmt11Per, max(InAmt1) as InAmt1, max(InAmt2) as InAmt2, max(InAmt3) as InAmt3, max(InAmt4) as InAmt4, max(InAmt5) as InAmt5, max(InAmt6) as InAmt6, max(InAmt7) as InAmt7, max(InAmt8) as InAmt8," + _
   "max(InAmt9) as InAmt9, max(InAmt10) as InAmt10, max(InAmt11) as InAmt11,  max(InAmt12) as InAmt12, max(InGstCd) as InGstCd,  max(InBaseAmt) as InBaseAmt," + _
       "Max(InCmTinNo) as InCmTinNo, " + wBaseRefVch + _
    "Max(InHPAN) as InHPAN, Max(InHTinNo) as InHTinNo, Max(InCmVATNo) as InCmVATNo, Max(InCmCstNo) as InCmCstNo, Max(InCmPAN) as InCmPAN, " + _
    "Max(" + ws_IdDmCd + ") as IdDmCd, Max(IdDmSz) as IdDmSz, (" + ws_IdQty + ") as IdQty, (" + ws_IdIFob + ") as IdIFob, Max(" + Ws_IdKt + ") as IdKt, " + _
    "Max(OdPicNm) as OdPicNm,Max(IdSfx) as IdSfx,Max((Case When OdPoNo<>'' Then OdPoNo Else OmPoNo End)) as OdPoNo,Max(IDTrayNo) as IDTrayNo, Max(ParamForDmctg.PDesc) as DmCtgDesc, Max(DmUom) as DmUom, " + IIF(wGrpDescFormula = "", "''", wGrpDescFormula) + " as wGrpDescFormula,Max(OdMrp) as OdMrp,Max(OdMrpDisc) as OdMrpDisc," + _
    "LEFT(max(InCmFax), 2) as qStateCode,IsNull((select PDesc from Param Where PTYp = 'STATECD' and PMCd= LEFT(max(InCmFax), 2) and PSCd= ''),'') as qStateName,max(InIRN) as InIRN,"
  '62 qCertNo added
  '38 qBagStkNo added
  wSqlStrg = wSqlStrg + _
    "Sum(case when RmCtg= 'G' then " + ws_IrRmIWt + " else 0 end) as qGldWt, Sum(case when RmCtg= 'G' then " + ws_IrRmQty + " else 0 end) as qGldQty, " + _
    "sum(case when RmCtg= 'P' then " + ws_IrRmIWt + " else 0 end) as qPlWt, sum(case when RmCtg= 'P' then " + ws_IrRmQty + " else 0 end) as qPlQty, " + _
    "sum(case when RmCtg= 'S' then " + ws_IrRmIWt + " else 0 end) as qSlWt, sum(case when RmCtg= 'S' then " + ws_IrRmQty + " else 0 end) as qSlQty, " + _
    "sum(case when RmCtg= 'L' then " + ws_IrRmIWt + " else 0 end) as qPdWt, sum(case when RmCtg= 'L' then " + ws_IrRmQty + " else 0 end) as qPdQty, " + _
    "sum(case when RmCtg= 'M' then " + ws_IrRmIWt + " else 0 end) as qOthMetWt, sum(case when RmCtg= 'M' then " + ws_IrRmQty + " else 0 end) as qOthMetQty, " + _
    "sum(case when RmCtg= 'D' then " + ws_IrRmIWt + " else 0 end) as qDiaWt, sum(case when RmCtg= 'D' then " + ws_IrRmQty + " else 0 end) as qDiaQty, " + _
    "sum(case when RmCtg= 'C' then " + ws_IrRmIWt + " else 0 end) as qCSWt, sum(case when RmCtg= 'C' then " + ws_IrRmQty + " else 0 end) as qCSQty, " + _
    "sum(case when RmCtg= 'X' then " + ws_IrRmIWt + " else 0 end) as qAccWt, sum(case when RmCtg= 'X' then " + ws_IrRmQty + " else 0 end) as qAccQty, " + _
    "sum(" + ws_IrRmIWt + "/ (case when RmCtg= 'D' or RmCtg= 'C' then 5 else 1 end)) as qGrsWt, " + _
    "(select isnull(max(CdCertNo),'') from CertDet join InvFgd on CdCtg='H' and CdBYy =IfFdBYy and CdBChr =IfFdBChr and CdBNo =IfFdBNo " + _
    "where IfCoCd= InCoCd and IfTc= InTc and IfYy= InYy and IfChr= InChr and IfNo= InNo And IfSr =max(IdSr) ) as qCertNo, " + _
    "(select min(IfCoCd+'/'+IfFdBYy+'/'+IfFdBChr+'/'+LTrim(str(IfFdBNo))) from InvFgd where IfCoCd= InCoCd and IfTc= InTc " + _
    "and IfYy= InYy and IfChr= InChr and IfNo= InNo And IfSr =max(IdSr) ) as qBagStkNo "
  '*** Jay 3.3(FG) [Variables introduced]
  
  Else
  ' ***** Manali 3.7.1
  'Diagold.2 printing state code and state name
  'sv.51 OdDmCol,InPrdDesc,InTermCond,InCmPayTerms,InHVessel are double entry,
  'dont know why it is like this. these are removed
  'in the isnull of InPrdDesc '' replaced with HPrdDesc
  'sv.63 max(InIRN) as InIRN added
  wSqlStrg = "Select Max(OdVaCtg+' '+ParamForVACtg.PDesc+' '+'HSN '+ParamForVACtg.PDesc225) as qGrp2, OmTc as InTc, OmYy as InYY, OmChr as InChr, OmNO as InNO, max(" + ws_IdSr + ") as IdSr, Max(IsNull(InPrnDt, OmDt)) as InDt,  " + _
    "Max(OmTc+'/'+OmYy+'/'+OmChr+'/'+RTrim(LTrim(str(OmNo)))) as InExpNo, Max(IsNull(InFOB, 0)) as InFOB, Max(IsNUll(InTot, 0)) as InTot, Max(IsNull(InCmCd, C.CmCd)) as InCmCd, Max(IsNull(InCmName, C.CmName)) as InCmName, " + _
    "Max(IsNull(InCmAdd1, C.CmAdd1)) as InCmAdd1, Max(IsNull(InCmAdd2, C.CmAdd2)) as InCmAdd2, Max(IsNUll(InCmAdd3, C.CmAdd3)) as InCmAdd3, Max(IsNull(InCmTel, C.CmTel)) as InCmTel, " + _
    "Max(IsNUll(InCmFax, C.CmFax)) as InCmFax, Max(C.CmBillCo) as CmBillCo, Max(IsNull(InSelfName, Z.CmName)) as InSelfName, Max(IsNull(InSelfAdd1, Z.CmAdd1)) as InSelfAdd1,  " + _
    "Max(IsNull(InSelfAdd2, Z.CmAdd2)) as InSelfAdd2, Max(IsNull(InSelfAdd3, Z.CmAdd3)) as InSelfAdd3, Max(IsNull(InSelfTel, Z.CmTel)) as InSelfTel,  " + _
    "Max(IsNull(InSelfFax, Z.CmFax)) as InSelfFax, Max(IsNull(InHCustRem, HCustRem)) as InHCustRem, Max(SubString(IsNull(InHUnResolutions, HUnResolutions), 1, 750)) as  " + _
    "InHUnResolutions, Max(IsNull(InLclInvTyp, Param.PDesc)) as InLclInvTyp, Max(IsNull(InHCstNo, HCstNo)) as InHCstNo, Max(IsNull(InHVatNo, HVatNo)) as InHVatNo, " + _
    "Max(IsNull(InHVatPer, 0.0/*HVatPer*/)) as InHVatPer, Max(IsNull(InLclVatRs, 0)) as InLclVatRs, Max(IsNull(InHCstPer, 0)) as InHCstPer, " + _
    "Max(IsNull(InLclCstRs, 0)) as InLclCstRs, Max(IsNull(InCmOctroiPer, 0)) as InCmOctroiPer, Max(IsNull(InLclOctroiRs, 0)) as InLclOctroiRs, " + _
    "Max(IsNull(InLclOthrValDesc, '')) as InLclOthrValDesc, Max(IsNull(InLclOthrValRs, 0)) as InLclOthrValRs, Max(IsNull(InHLclInvFT, HLclInvFT)) as InHLclInvFT, " + _
    "Max(SubString((Case when IsNull(InLclCstRs, 0) <> 0 then IsNull(InHLclInvCstFT, HLclInvCstFT) when IsNull(InLclVatRs, 0) <> 0 then IsNUll(InHLclInvVatFT, HLclInvVatFT) else '' end), 1, 750)) as InHLclInvVATFT, " + _
    "Max(IsNull(InHAuthSgn, HAuthSgn)) as InHAuthSgn, Max(C.CmCurCd) as qCurCd, " + _
    "(select PValue from Param Where PTYp = 'CURNCY' and PMCd= Max(C.CmCurCd) and PSCd= '') as qCurUnit,   /* (Select DmCtg from DsgMst where DmTcTyp = 'DM' and DmCd= Max(" + ws_IdDmCd + ") and DmSz= '') as qDmCtg, */" + _
    "Max(OdDmCol) as OdDmCol, Replace(Max(IsNull(InPrdDesc, HPrdDesc)),char(13)+ char(10),' ')  as InPrdDesc, " + _
    "Max(SubString(IsNull(InTermCond, (Case When OmTc='JMI' Then HInMemoTermCond Else HInInvTermCond End)), 1, 1500)) as InTermCond, Max(IsNull(InCmPayTerms, OmCmPayTerms)) as InCmPayTerms, Max(IsNull(InHVessel, HVessel)) as inHVessel, " + _
    "/*Max(OdDmCol) as OdDmCol, Replace(Max(InPrdDesc),char(13)+ char(10),' ')  as InPrdDesc, Max(SubString(InTermCond, 1, 1500)) as InTermCond, Max(InCmPayTerms) as InCmPayTerms, max(InHVessel) as InHVessel,*/ max (InAmtDesc1) as InAmtDesc1," + _
    "max (InAmtDesc2) as InAmtDesc2, max (InAmtDesc3) as InAmtDesc3, max (InAmtDesc4) as InAmtDesc4, max (InAmtDesc5) as InAmtDesc5, max (InAmtDesc6) as InAmtDesc6, max (InAmtDesc7) as InAmtDesc7, max (InAmtDesc8) as InAmtDesc8, max (InAmtDesc9) as InAmtDesc9, max (InAmtDesc10) as InAmtDesc10, " + _
   "max (InAmtDesc11) as InAmtDesc11, max (InAmtDesc12) as InAmtDesc12, max (InAmt1Per)  as InAmt1Per , max (InAmt2Per) as InAmt2Per, max (InAmt3Per) as InAmt3Per, max (InAmt4Per) as InAmt4Per, max (InAmt5Per) as InAmt5Per, max (InAmt6Per) as InAmt6Per, max (InAmt8Per) as InAmt8Per," + _
   "max (InAmt9Per) as InAmt9Per,  max (InAmt10Per) as InAmt10Per, max (InAmt11Per) as InAmt11Per, max(InAmt1) as InAmt1, max(InAmt2) as InAmt2, max(InAmt3) as InAmt3, max(InAmt4) as InAmt4, max(InAmt5) as InAmt5, max(InAmt6) as InAmt6, max(InAmt7) as InAmt7, max(InAmt8) as InAmt8," + _
   "max(InAmt9) as InAmt9, max(InAmt10) as InAmt10, max(InAmt11) as InAmt11,  max(InAmt12) as InAmt12,  max(InGstCd) as InGstCd,  max(InBaseAmt) as InBaseAmt, " + _
    "Max(IsNull(InCmTinNo, C.CmTinNo)) as InCmTinNo, " + wBaseRefVch + _
    "Max(IsNull(InHPAN, HPAN)) as InHPAN, Max(IsNull(InHTinNo, HTinNo)) as InHTinNo, Max(IsNull(InCmVATNo, C.CmVATNo)) as InCmVATNo, Max(IsNull(InCmCstNo, C.CmCstNo)) as InCmCstNo, Max(IsNull(InCmPAN, C.CmPAN)) as InCmPAN, " + _
    "Max(" + ws_IdDmCd + ") as IdDmCd, Max(OdDmSz) as IdDmSz, (" + ws_IdQty + ") as IdQty, (" + ws_IdIFob + ") as IdIFob, Max(" + Ws_IdKt + ") as IdKt, " + _
    "Max(OdPicNm) as OdPicNm,Max(OdSfx) as IdSfx,Max((Case When OdPoNo<>'' Then OdPoNo Else OmPoNo End)) as OdPoNo,0 as IDTrayNo, Max(ParamForDmctg.PDesc) as DmCtgDesc, Max(DmUom) as DmUom, " + IIF(wGrpDescFormula = "", "''", wGrpDescFormula) + " as wGrpDescFormula,Max(OdMrp) as OdMrp,Max(OdMrpDisc) as OdMrpDisc, " + _
    "LEFT(Max(IsNUll(InCmFax, C.CmFax)), 2) as qStateCode,IsNull((select PDesc from Param Where PTYp = 'STATECD' and PMCd= LEFT(Max(IsNUll(InCmFax, C.CmFax)), 2) and PSCd= ''),'') as qStateName,max(InIRN) as InIRN,"
        
    'wSqlStrg = wSqlStrg + "(Select PNum from Param where PTyp='VACTG' and PMCd=max(OdVaCtg) and PSCd='') as IdGldLs,"
    
  'sv.51 isnull added to qGrsWt
  '62 '' as qCertNo added
  '38 qBagStkNo added
  '5 qCertNo changed
  wSqlStrg = wSqlStrg + _
    "Sum(case when RmCtg= 'G' then " + ws_IrRmIWt + " else 0 end) as qGldWt, Sum(case when RmCtg= 'G' then " + ws_IrRmQty + " else 0 end) as qGldQty, " + _
    "sum(case when RmCtg= 'P' then " + ws_IrRmIWt + " else 0 end) as qPlWt, sum(case when RmCtg= 'P' then " + ws_IrRmQty + " else 0 end) as qPlQty, " + _
    "sum(case when RmCtg= 'S' then " + ws_IrRmIWt + " else 0 end) as qSlWt, sum(case when RmCtg= 'S' then " + ws_IrRmQty + " else 0 end) as qSlQty, " + _
    "sum(case when RmCtg= 'L' then " + ws_IrRmIWt + " else 0 end) as qPdWt, sum(case when RmCtg= 'L' then " + ws_IrRmQty + " else 0 end) as qPdQty, " + _
    "sum(case when RmCtg= 'M' then " + ws_IrRmIWt + " else 0 end) as qOthMetWt, sum(case when RmCtg= 'M' then " + ws_IrRmQty + " else 0 end) as qOthMetQty, " + _
    "sum(case when RmCtg= 'D' then " + ws_IrRmIWt + " else 0 end) as qDiaWt, sum(case when RmCtg= 'D' then " + ws_IrRmQty + " else 0 end) as qDiaQty, " + _
    "sum(case when RmCtg= 'C' then " + ws_IrRmIWt + " else 0 end) as qCSWt, sum(case when RmCtg= 'C' then " + ws_IrRmQty + " else 0 end) as qCSQty, " + _
    "sum(case when RmCtg= 'X' then " + ws_IrRmIWt + " else 0 end) as qAccWt, sum(case when RmCtg= 'X' then " + ws_IrRmQty + " else 0 end) as qAccQty, " + _
    "sum(isnull(" + ws_IrRmIWt + "/ (case when RmCtg= 'D' or RmCtg= 'C' then 5 else 1 end),0)) as qGrsWt, " + _
    "(select isnull(max(CdCertNo),'') from CertDet where CdCtg='H' and CdBCoCd+'/'+ CdBYy+'/'+CdBChr+'/'+LTrim(str(CdBNo))= max(OdDmStkCoCd+'/'+ OdDmStkYy+'/'+OdDmStkChr+'/'+LTrim(str(OdDmStkNo))) ) as qCertNo, " + _
    "min(OdDmStkCoCd+'/'+ OdDmStkYy+'/'+OdDmStkChr+'/'+LTrim(str(OdDmStkNo))) as qBagStkNo "
  End If
  
  ' ***** Manali 3.8.0 - Sort By Design
  wDsgSrt = IIF(adc("UoYn1") = "Y", " IdDmCd, ", "")
  
  
  
  '*** Jay 3.3(FG)
  ' *** Bef Trading Module
  'If UCase(adc.MenuCd) <> UCase("RepFgLclInv") Then
  ' **** Manali Trading Module
  ' ***** Manali 3.8.0 - Sort By Design - wDsgSrt added in Order By
  If UCase(adc.MenuCd) = UCase("RepLclInv") Or UCase(adc.MenuCd) = UCase("RepInvLclRM") Then
  '*** Jay 3.3(FG)
    '****** Sachin 3.02 - Id fields in Join
    ' ***** Manali 3.7.0 - Trading Module - OrdDsg Join added (For OdDmCol)
    'hk.5 param joins added for Kt,DmCtg,DmCol,Salctg,rmmst
    wJoinWhr = "From InvRm " + _
      "Join RmMst On IrRmCd= RmCd " + IIF(gs_Partition = ctCurrPrtn, " and IrPrtkey=RmPrtKey ", "") + _
      "Join InvDsg On IdIdNo=IrIdIdNo and IdCoCd= IrCoCd and IdTc= IrTc and IdYy= IrYy and IdChr= IdChr and IdNo= IrNo and IdSr= IrSr " + IIF(gs_Partition = ctCurrPrtn, " and IdPrtKey=IrPrtKey ", "") + _
      "Join InvHd On InCoCd=IrCoCd and InTc= IrTc and InYy= IrYy and InChr = IrChr and InNO= IrNo " + IIF(gs_Partition = ctCurrPrtn, " and IrPrtKey=InPrtKey ", "") + _
      "Join OrdDsg On OdIdNo=IdOdIdNo and OdCoCd=IdCoCd and OdTc=IdExpOdTc and OdYy=IdExpOdYy and OdChr=IdExpOdChr and OdNo=IdExpOdNo and OdSr=IdExpOdSr " + IIF(gs_Partition = ctCurrPrtn, " and OdPrtKey=IdPrtKey ", "") + _
      "Join OrdMst On OdOmIdNo=OmIdNo and OdCoCd=OmCoCd and OdTc=OmTc and OdYy=OmYy and OdChr=OmChr and OdNo=OmNo " + IIF(gs_Partition = ctCurrPrtn, " and OmPrtKey=OdPrtKey ", "") + _
      "Join CustMst On CmCtg= 'C' and CmCd= InCmCd " + _
      "Join Param ParamForVACtg on ParamForVACtg.PTyp='VACTG' and ParamForVACtg.PMCd = IdVaCtg " + _
      "join DsgMst on DmIdNo=IdDmIdNo and DmTcTyp= 'DM' and DmCd = IdDmCd and DmSz= '' " + IIF(gs_Partition = ctCurrPrtn, " And DmPrtKey=IdPrtKey ", "") + _
      "join Param ParamForKt on ParamForKt.PTyp ='KT' and ParamForKt.PMCd = OdKt " + _
      "join Param ParamForDmctg on ParamForDmctg.PTyp ='DMCTG' and ParamForDmctg.PMCd = DmCtg " + _
      "join Param ParamForDmCol on ParamForDmCol.PTyp ='DMCOL' and ParamForDmCol.PMCd = OdDmCol " + _
      "join Param ParamForSalCtg on ParamForSalCtg.PTyp ='SALCTG' and ParamForSalCtg.PMCd = DmCtg and ParamForSalCtg.PSCd =DmSalCtg " + wMainMetJoin + _
      "where " + IIF(gs_Partition = ctCurrPrtn, " InPrtKey='" + ctCurrPrtn + "' and ", "") + " InLclInvYn= 'Y' " + IIF(wCnd <> "", " and ", "") + wCnd + " " + _
      " Group By InCoCd, InTc, InYy, InChr, InNO " + IIF(wGrpStr = "", "", ",") + wGrpStr + "," + wDetGrp
      
  '*** Jay 3.3(FG)
  wSqlStrg = wSqlStrg + wJoinWhr + " Order By InCoCd, InTc, InYy, InChr, InNO, " + wDsgSrt + "IdSr"

  Else
    Dim wLeftOutJn As String
    wLeftOutJn = IIF(UCase(adc.MenuCd) = UCase("RepJtLclInv"), "Left Outer ", "")
    ' Manali 3.7.0 - Trading Module - OrdMst Join added, CmCtg='C' replaced by OmCmCtg
    'hk.5 param joins added for Kt,DmCtg,DmCol,Salctg,rmmst
    'sv.51 OrdRm and OrdDsg Od interchanged to get jri/jro records even there is no rm details
    'Left Outer added to OrdRm,RmMst
    wJoinWhr = " From OrdDsg Od " + _
      " Left Outer Join OrdRm On OdIdNo=OrOdIdNo and OdCoCd= OrCoCd and OdTc= OrTc and OdYy= OrYy and OdChr= OdChr and OdNo= OrNo and OdSr= OrSr " + IIF(gs_Partition = ctCurrPrtn, " and OdPrtKey=OrPrtKey ", "") + _
      " Left Outer Join RmMst On OrRmCd= RmCd " + IIF(gs_Partition = ctCurrPrtn, " and OrPrtkey=RmPrtKey ", "") + _
      " Join OrdMst On OmIdNo=OdOmIdNo and OmCoCd=OdCoCd and OmTc=OdTc and OmYY=OdYY and OmChr=OdChr and OdNo=OmNo " + IIF(gs_Partition = ctCurrPrtn, " and OmPrtKey=OdPrtKey ", "") + _
      " Join CustMst C On C.CmCtg=OmCmCtg and C.CmCd= OmCmCd " + _
      " Join Head On HCoCd=OmCoCd and HCd='" + ctSelfCmCd + "' " + _
      " Join CustMst Z On Z.CmCtg='Z' and Z.CmCd=HCoCd  " + _
      " Join Param ParamForVACtg on ParamForVACtg.PTyp = 'VACTG' and ParamForVACtg.PMCd=OdVaCtg " + _
      " Join Param On Param.PTyp='TC' and Param.PMCd=OdTc " + _
      wLeftOutJn + " Join InvHd On InCoCd=OrCoCd and InTc= OrTc and InYy= OrYy and InChr = OrChr and InNO= OrNo " + IIF(gs_Partition = ctCurrPrtn, " and OrPrtKey=InPrtKey ", "") + _
      "join DsgMst on DmIdNo=OdDmIdNo and DmTcTyp= 'DM' and DmCd = OdDmCd and DmSz= '' " + IIF(gs_Partition = ctCurrPrtn, " And DmPrtKey=OdPrtKey ", "") + _
      "join Param ParamForKt on ParamForKt.PTyp ='KT' and ParamForKt.PMCd = OdKt " + _
      "join Param ParamForDmctg on ParamForDmctg.PTyp ='DMCTG' and ParamForDmctg.PMCd = DmCtg " + _
      "join Param ParamForDmCol on ParamForDmCol.PTyp ='DMCOL' and ParamForDmCol.PMCd = OdDmCol " + _
      "join Param ParamForSalCtg on ParamForSalCtg.PTyp ='SALCTG' and ParamForSalCtg.PMCd = DmCtg and ParamForSalCtg.PSCd =DmSalCtg " + wMainMetJoin + _
      " where /*(Select PValue from Param Where PTyp='TC' and PMCd=OdTc)*/ Param.PValue=" + IIF(UCase(adc.MenuCd) = UCase("RepJtLclInv"), "'JT'", "'FG'") + " " + IIF(gs_Partition = ctCurrPrtn, " and OmPrtKey='" + ctCurrPrtn + "' ", "") + IIF(wCnd <> "", " and ", "") + wCnd + " " + _
      " Group By OmCoCd, OmTc, OmYy, OmChr, OmNO " + IIF(wGrpStr = "", "", ",") + wGrpStr + "," + wDetGrp
    
    wSqlStrg = wSqlStrg + wJoinWhr + " Order By OmCoCd, OmTc, OmYy, OmChr, OmNO, " + wDsgSrt + "IdSr"
  
           
''    wSqlStrg = wSqlStrg + "From OrdRm " + _
''      "Join RmMst On OrRmCd= RmCd " + IIF(gs_Partition = ctCurrPrtn, " and OrPrtkey=RmPrtKey ", "") + _
''      "Join OrdDsg On OdIdNo=OrOdIdNo and OdCoCd= OrCoCd and OdTc= OrTc and OdYy= OrYy and OdChr= OdChr and OdNo= OrNo and OdSr= OrSr " + IIF(gs_Partition = ctCurrPrtn, " and OdPrtKey=OrPrtKey ", "") + _
''      "Join OrdMst On OmIdNo=OdOmIdNo and OmCoCd=OdCoCd and OmTc=OdTc and OmYY=OdYY and OmChr=OdChr and OdNo=OmNo " + IIF(gs_Partition = ctCurrPrtn, " and OmPrtKey=OdPrtKey ", "") + _
''      "Join InvHd On InCoCd=OrCoCd and InTc= OrTc and InYy= OrYy and InChr = OrChr and InNO= OrNo " + IIF(gs_Partition = ctCurrPrtn, " and OrPrtKey=InPrtKey ", "") + _
''      "Join CustMst On CmCtg=OmCmCtg /*'C'*/ and CmCd= InCmCd " + _
''      "where " + IIF(gs_Partition = ctCurrPrtn, " InPrtKey='" + ctCurrPrtn + "' and ", "") + " InLclInvYn= 'Y' " + wCnd + " " + _
''      " Group By InCoCd, InTc, InYy, InChr, InNO, " + ws_IdSr
  End If
  
  '*** Jay 3.3(FG)
  '  wSqlStrg = wSqlStrg + " Order By InCoCd, InTc, InYy, InChr, InNO, IdSr"
  '*** Jay 3.3(FG)
     
  'pg.11 checking base values match
  If UCase(adc.MenuCd) = UCase("RepLclInv") Or UCase(adc.MenuCd) = UCase("RepJtLclInv") Then
    Dim wRsSqlStrg As MDORowSet, BaseVal As Double
    If UCase(adc.MenuCd) = UCase("RepLclInv") Then
      Set wRsSqlStrg = moCn.OpenRes("Select distinct InCoCd, InTc, InYy, InChr, InNo ,max(InGstBase) as InGstBase, max(InBaseAmt) as InBaseAmt " + wJoinWhr)
    ElseIf UCase(adc.MenuCd) = UCase("RepJtLclInv") Then
      Set wRsSqlStrg = moCn.OpenRes("Select distinct OmCoCd as InCoCd, OmTc as InTc, OmYy as InYy, OmChr as InChr, OmNo as InNo, max(InGstBase) as InGstBase, max(InBaseAmt) as InBaseAmt, max(InGstCd) as InGstcd " + wJoinWhr + " having max(InGstCd) is not null")
    End If

    Do While Not (wRsSqlStrg.BOF Or wRsSqlStrg.EOF)
      If wRsSqlStrg!InGstBase <> "-" Then
        BaseVal = InBaseCalc(wRsSqlStrg!InGstBase, wRsSqlStrg!incocd, wRsSqlStrg!InTc, wRsSqlStrg!InYy, wRsSqlStrg!InChr, wRsSqlStrg!InNo)
        'If wRsSqlStrg!InBaseAmt <> BaseVal Then _
          DispMsg "Base Values are Not Matched in Inv No:" + wRsSqlStrg!incocd + "/" + wRsSqlStrg!InTc + "/" + wRsSqlStrg!InYy + "/" + wRsSqlStrg!InChr + "/" + CStr(wRsSqlStrg!InNo), etError: Exit Sub
      End If
      wRsSqlStrg.MoveNext
    Loop
  End If
  
  adc.RepRecSource = wSqlStrg + ctMaxDopOpt   '****** Sachin 3.01 added ctMaxDopOpt
  
  If UCase(adc.MenuCd) = UCase("RepInvLclRM") Then
  Dim ws_RmCtg As String, ws_RmSCtg As String, ws_RmCd As String, ws_RmRt As String, ws_RmGrpOpt As String, ws_RmSubRep As String, wRmValOpt As String
   Dim wRmRtOpt As String
   ws_RmCtg = "max(RmCtg) + max((Case When PSCd is Null Then '' Else '-' + IsNull(PSCd, '') End )) "
      ws_RmSCtg = "Max(RmSCtg) "
      ws_RmCd = "RmCd "
      ws_RmRt = "max(IrRmIRt)"
      ws_RmGrpOpt = ",RmCd"
      wRmValOpt = "IrRmIVal"
      wRmRtOpt = "IrRmIRt"
    Select Case UCase(adc("UoRmGrp"))
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
      ws_RmRt = "0"
      Rep.Subreport2_hSubRep2RmRt.Suppress = True: Rep.Subreport2_FldSubRep2RmRt.Suppress = True
    Case Is = "RMSCTG"
      ws_RmGrpOpt = ", RmCtg, RmSCtg "
      ws_RmCtg = "max(RmCtg) + max((Case When PSCd is Null Then '' Else '-' + IsNull(PSCd, '') End )) "
      ws_RmSCtg = "max(RmSCtg) "
      ws_RmCd = "'' "
      ws_RmRt = "0 "
      Rep.Subreport2_hSubRep2RmCd.Suppress = True: Rep.Subreport2_FldSubRep2RmCd.Suppress = True
      Rep.Subreport2_hSubRep2RmRt.Suppress = True: Rep.Subreport2_FldSubRep2RmRt.Suppress = True
    Case Is = "RMCTG"
      ws_RmGrpOpt = ", RmCtg"
      ws_RmCtg = "max(RmCtg) + max((Case When PSCd is Null Then '' Else '-' + IsNull(PSCd, '') End )) "
      ws_RmSCtg = "'' "
      ws_RmCd = "'' "
      ws_RmRt = "0 "
      Rep.Subreport2_hSubRep2RmsCtg.Suppress = True: Rep.Subreport2_FldSubRep2RmSCtg.Suppress = True
      Rep.Subreport2_hSubRep2RmCd.Suppress = True: Rep.Subreport2_FldSubRep2RmCd.Suppress = True
      Rep.Subreport2_hSubRep2RmRt.Suppress = True: Rep.Subreport2_FldSubRep2RmRt.Suppress = True
    End Select
  
   'JewelSmith.2 ',' + PDesc225 + ',' like '%,'+ RmSCtg +',%' changed as PSCd = RmStnCls
   ws_RmSubRep = "Select " + ws_RmCtg + " as qRmCtg, " + ws_RmSCtg + " as RmSCtg, " + _
          ws_RmCd + " as RmCd, " + ws_RmRt + " as qRmRt, sum(IrRmQty) as IrRmQty, " + _
          "sum(case when RmCtg in ('C', 'D') then " + ws_IrRmIWt + " else 0 end) as qWtInCts, " + _
          "sum(case when RmCtg in ('C', 'D') then " + ws_IrRmIWt + "/5 else " + ws_IrRmIWt + " end) as qWtInGms, " + _
          "sum(" + wRmValOpt + ") as qVal, " + _
          "sum(" + wRmValOpt + ")* max(case when InRsMulDiv= 'D' and convert (decimal (16, 4), InCnvRt)> 0 then 1/ InCnvRt else InCnvRt end) as qValRs " + _
          " from InvRm join RmMst on RmCd = IrRmCd  " + IIF(gs_Partition = ctCurrPrtn, " And RmPrtKey=IrPrtKey ", "") + _
          "join InvHd Ih on InCoCd= IrCoCd and InTc = IrTc and InYy= IrYy and InChr= IrChr and " + _
          "     InNo= IrNo " + IIF(gs_Partition = ctCurrPrtn, " And InPrtKey=IrPrtKey ", "") + _
          "left outer join Param on PTyp= 'STNCLS' and PMCD=RmCtg and PSCd=RmStnCls" + _
          " where /* RmCtg in ('D', 'C', 'X') */ " + wCnd + IIF(gs_Partition = ctCurrPrtn, " And IrPrtKey='" + ctCurrPrtn + "' ", "") + _
          "Group by IrTc, IrYy, IrChr, IrNo " + ws_RmGrpOpt + _
          " Order By IrTc, IrYy, IrChr, IrNo,Max(RmCtg) " + ws_RmGrpOpt
    
  Call moCn.RepRes(mRep.Subreport2.OpenSubreport, ws_RmSubRep)
  'mRep.Subreport2.Suppress = True
  mRep.PgHd.Suppress = True
  mRep.DetDsg.Suppress = True
  mRep.PICSEC.Suppress = True
  mRep.LinePicSec.Suppress = True
  mRep.DetGld.Suppress = True
  mRep.DetP.Suppress = True
  mRep.DetSl.Suppress = True
  mRep.DetPld.Suppress = True
  mRep.DetMet.Suppress = True
  mRep.DetX.Suppress = True
  mRep.LineSec.Suppress = True
  mRep.GfA.Suppress = True
  mRep.GfTot.Suppress = True
  
  mRep.Line1.Suppress = True
  mRep.Line16.Suppress = True
  'hk.5- renamed vertical line from Line2 to LineDescLeftD
  mRep.LineDescLeftD.Suppress = True
'  mRep.LineBRVchLeftD.Suppress = True
  mRep.LineBRVchRightD.Suppress = True
  'hk.5- renamed vertical line from Line11 to LineOdDmColLeftD
  mRep.LineOdDmColLeftD.Suppress = True
  mRep.Line48.Suppress = True
  mRep.Line12.Suppress = True
  mRep.Line3.Suppress = True
  mRep.Line44.Suppress = True
  mRep.Line4.Suppress = True
  mRep.Line45.Suppress = True
  mRep.Line46.Suppress = True
  mRep.Line5.Suppress = True
  '*** Jay 3.3(FG)
  mRep.GFInAddAmt.Suppress = True
  'hk.5 for invoice rm report not showing new group & detail section
  Rep.wGrpSec.Suppress = True: mRep.DetGrpSec.Suppress = True
   
   
  Else
      mRep.RmSubRepSec.Suppress = True
      
  If OptNewRpt.Value = True Then
    mRep.GFInAddAmt.Suppress = False
   ' mRep.wWordsNewFormat.Suppress = False
  '  mRep.wWords1.Suppress = True
 '   mRep.GFAmtWord.Suppress = True
    mRep.GFRmSumm.Suppress = True
  Else
'  mRep.wWordsNewFormat.Suppress = True
 ' mRep.wWords1.Suppress = False
   mRep.GFInAddAmt.Suppress = True
'   mRep.GFAmtWord.Suppress = False
  End If
   
  If OptOldRpt.Value = True Then
   mRep.GFInAddAmt.Suppress = True
 GRP_REP.SetFormula mRep, "wWords", "{rdo.qCurUnit} + ' '+ Replace (ToWords ({@wTot}), ' and xx / 100', '')+ ' only'"

   
'   mRep.wWordsNewFormat.Suppress = True
 '  mRep.GFAmtWord.Suppress = False
'   mRep.wWords1.Suppress = False
  Else
  GRP_REP.SetFormula mRep, "wWords", "{rdo.qCurUnit} + ' '+ Replace (ToWords ({rdo.InAmt12}), ' and xx / 100', '')+ ' only'"
   mRep.GFInAddAmt.Suppress = False
 '  mRep.wWordsNewFormat.Suppress = False
  ' mRep.GFAmtWord.Suppress = True
'    mRep.wWords1.Suppress = True
  End If
   
  End If
  
   

  GRP_REP.SetFormula mRep, "wGpWt", "If {rdo.qGldWt} = 0 then {rdo.qPlWt} Else {rdo.qGldWt}"
  GRP_REP.SetFormula mRep, "wGpCtg", "if {rdo.qGldWt} <> 0 then 'Gld' Else if {rdo.qGldWt} = 0 and {rdo.qPlWt} <> 0 then 'Plat' Else ''"
  ' ***** Manali
  GRP_REP.SetFormula mRep, "wDcWt", "If ({rdo.qDiaWt} = 0 AND {rdo.qDiaQty} = 0) then {rdo.qCsWt} Else {rdo.qDiaWt}"
  GRP_REP.SetFormula mRep, "wDcCtg", "if ({rdo.qDiaWt} <> 0 or {rdo.qDiaQty} <> 0) then 'Dia' Else if {rdo.qDiaWt} = 0 and {rdo.qDiaQty} = 0 and ({rdo.qCsWt} <> 0 or {rdo.qCsQty} <> 0) then  'Col Stn' Else ''"
  GRP_REP.SetFormula mRep, "wDcQty", "If ({rdo.qDiaWt} = 0 AND {rdo.qDiaQty} = 0) then {rdo.qCsQty} Else {rdo.qDiaQty}"
  
'  GRP_REP.SetFormula mRep, "wDcWt", "If {rdo.qDiaWt} = 0 then {rdo.qCsWt} Else {rdo.qDiaWt}"
'  GRP_REP.SetFormula mRep, "wDcQty", "If {rdo.qDiaQty} = 0 then {rdo.qCsQty} Else {rdo.qDiaQty}"
'  GRP_REP.SetFormula mRep, "wDcCtg", "if {rdo.qDiaWt} <> 0 then 'Dia' Else if {rdo.qDiaWt} = 0 and {rdo.qCsWt} <> 0 then  'Col Stn' Else ''"

  If UCase(adc.MenuCd) = UCase("RepLclInv") Then
    If adc("UOYN6") = "N" Then  '4.1.4
        Rep.hDsgSfx.Suppress = True
        Rep.FldDsgSfx.Suppress = True
    End If
  End If
  'hk.5 if show logo is selected then displays the logo in report
  If adc("UOYN7") = "Y" Then
    Rep.PicLogo1.Suppress = False
    Rep.PicDTCLogo.Suppress = False
  End If
  
  ' ***** Manali 3.7.1
  If UCase(adc.MenuCd) = UCase("RepLclInv") Or UCase(adc.MenuCd) = UCase("RepInvLclRM") Then
    GRP_REP.SetFormula Rep, "wInFob", "{rdo.InFob}"
    'Rep.HBaseRefVch.Suppress = True: Rep.FldBaseRefVch.Suppress = True
    'Rep.LineBRVchLeftD.Suppress = True
    Rep.LineBRVchLeftH.Suppress = True
    If adc("UOYN5") = "N" Then     '4.1.4
        Rep.TxtOdPoNo.Top = 0
        Rep.TxtOdPoNo.Suppress = True
        Rep.DetDsg.Height = 221
    End If
  Else
    GRP_REP.SetFormula Rep, "wInFob", "iif({rdo.InFOB}=0, {@wTotFob}, {rdo.InFOB})"
    GRP_REP.SetFormula Rep, "wTot", "IIF({rdo.InFOB}=0, {@wInFob}, Round ({rdo.InTot}, 0))"      '"{@wInFob}"
    Rep.LblInExpNo.SetText "Vch No"
    Rep.LblInDt.SetText "Vch Date"
    
  '  Rep.HBaseRefVch.SetText IIF(wBaseRefIdKey = "OdBaseIdKey", "Base Voucher", IIF(wBaseRefIdKey = "OdRefIdKey", "Reference Voucher", ""))
    If wBaseRefIdKey = "" Then
      'Rep.LineBRVchLeftD.Suppress = True:
      Rep.LineBRVchLeftH.Suppress = True
    End If
  End If
  ' ***** Manali 3.7.1
  
  If UCase(adc("UoYN3")) = "N" Then
''    mRep.Line16.Suppress = True
    ' ***** Manali 3.10.0 - 31/03/12 - Show Picture
    GRP_REP.SetFormula Rep, "wShowDemark", "'N'"
    ' ***** Manali 3.10.0 - 31/03/12 - Show Picture
  Else
    ' ***** Manali 3.10.0 - 31/03/12 - Show Picture - Code Shifted to Report VB Code
    GRP_REP.SetFormula Rep, "wShowDemark", "'Y'"
    'mRep.Line16.Suppress = False
    ' ***** Manali 3.10.0 - 31/03/12 - Show Picture
  End If

  ' ***** Manali 3.10.0 - 31/03/12 - Show Picture
  If UCase(adc("UoYN2")) = "Y" Then
    GRP_REP.SetFormula Rep, "wShowPic", "'Y'"
  Else
    GRP_REP.SetFormula Rep, "wShowPic", "'N'"
    'Rep.PicSec.Suppress = True
  End If
  ' ***** Manali 3.10.0 - 31/03/12 - Show Picture
  
  ' ***** Manali 3.6.0 - 24/11/09 - Show Run Sr
  If adc("UoYn") = "Y" Then
    Rep.FldRunSr.Suppress = False: Rep.FldIdSr.Suppress = True
  Else
    Rep.FldRunSr.Suppress = True: Rep.FldIdSr.Suppress = False
  End If
  ' ***** Manali 3.6.0 - 24/11/09 - Show Run Sr
  
  ' ***** Manali 3.10.0 - 12/04/12 - Show Rm Summary
  If adc("UoYn4") = "N" Then
    Rep.gfGldQty.Suppress = True: Rep.gfGldWt.Suppress = True: Rep.txtGfGld.Suppress = True
    Rep.gfPlQty.Suppress = True: Rep.gfPlWt.Suppress = True: Rep.txtGfPl.Suppress = True
    Rep.gfSlQty.Suppress = True: Rep.gfSlWt.Suppress = True: Rep.txtGfSl.Suppress = True
    Rep.gfPdQty.Suppress = True: Rep.gfPdWt.Suppress = True: Rep.txtGfPd.Suppress = True
    Rep.gfDiaQty.Suppress = True: Rep.gfDiaWt.Suppress = True: Rep.txtGfDia.Suppress = True
    Rep.gfCsQty.Suppress = True: Rep.gfCsWt.Suppress = True: Rep.txtGfCS.Suppress = True
    Rep.gfAccQty.Suppress = True: Rep.gfAccWt.Suppress = True: Rep.txtGfAcc.Suppress = True
    Rep.gfMetQty.Suppress = True: Rep.gfMetWt.Suppress = True: Rep.txtGfMet.Suppress = True
    Rep.LblGfRmCtg1.Suppress = True: Rep.LblGfRmQty1.Suppress = True: Rep.LblGfRmWt1.Suppress = True
    Rep.LblGfRmCtg2.Suppress = True: Rep.LblGfRmQty2.Suppress = True: Rep.LblGfRmWt2.Suppress = True
    Rep.BoxRmSumm.Suppress = True
    Rep.LineRmSumm1.Suppress = True: Rep.LineRmSumm2.Suppress = True: Rep.LineRmSumm3.Suppress = True
    Rep.LineRmSumm4.Suppress = True: Rep.LineRmSumm5.Suppress = True: Rep.LineRmSumm6.Suppress = True
    Rep.LineRmSumm7.Suppress = True
  End If
  ' ***** Manali 3.10.0 - 12/04/12 - Show Rm Summary
'  pg.11- Gstcode & base amount will display only for local invoice in Emr&jEmr
  If UCase(adc.MenuCd) = UCase("RepLclInv") Or UCase(adc.MenuCd) = UCase("RepJtLclInv") Then
    Rep.TxtGstCd.Suppress = False:    Rep.InGstCd.Suppress = False
    Rep.TxtBaseAmt.Suppress = False:  Rep.InBaseAmt.Suppress = False
  Else
    Rep.TxtGstCd.Suppress = True:    Rep.InGstCd.Suppress = True
    Rep.TxtBaseAmt.Suppress = True:  Rep.InBaseAmt.Suppress = True
  End If
    
  If UCase(adc.MenuCd) = UCase("RepLclInv") Or UCase(adc.MenuCd) = UCase("RepJtLclInv") Then
    If OptNewRpt.Value = True Then
      GRP_REP.SetFormula mRep, "wDetDesc", wDetDescFormula
      'hk.5 if group section is not selected then suppress that section
      If adc("UoInGrp") = "" Then Rep.wGrpSec.Suppress = True
      'hk.5 suppress the fields as per option selected in details group option
      If adc("UoInDet") <> "INVSR" And adc("UoInDet") <> "" Then
        Rep.IdDmCd1.Suppress = True: Rep.FldIdSr.Suppress = True
        Rep.FldRunSr.Suppress = True: Rep.FldDsgSfx.Suppress = True
        Rep.IdDmSz1.Suppress = True: Rep.IdKt1.Suppress = True: Rep.OdDmCol1.Suppress = True
        Rep.TxtOdPoNo.Suppress = True: Rep.TxtOdPoNo.Top = 0
        Rep.hDsgCd.Suppress = True: Rep.hDsgSfx.Suppress = True: Rep.hSrNo.Suppress = True
        Rep.hDsgSz.Suppress = True: Rep.hDsgKt.Suppress = True: Rep.hDsgCol.Suppress = True
        Rep.LineBRVchRightH.Suppress = True: Rep.LineBRVchRightD.Suppress = True: Rep.LineBRVchRightG1.Suppress = True
        Rep.LineOdDmColLeftH.Suppress = True: Rep.LineOdDmColLeftD.Suppress = True: Rep.LineOdDmColLeftG1.Suppress = True
        Rep.LineDescLeftH.Suppress = True: Rep.LineDescLeftD.Suppress = True: Rep.LineDescLeftG1.Suppress = True
        Rep.LineBRVchLeftH.Suppress = True
      Else
        Rep.DetGrpSec.Suppress = True: Rep.hDesc.Suppress = True
      End If
    Else
      Rep.wGrpSec.Suppress = True: Rep.DetGrpSec.Suppress = True: Rep.hDesc.Suppress = True
    End If
  End If
  'pg.5- If Show MRP=Y or detail level has Sr.level then can view MRP details else cannot show
  If adc("UOYN8") <> "Y" Or OptOldRpt.Value = True Or (adc("UoInDet") <> "INVSR" And adc("UoInDet") <> "" And adc("UoInDet") <> "DMCD/SR" And adc("UoInDet") <> "INVSR/SR" And adc("UoInDet") <> "TRNO/SR") Then
    Rep.hMrp.Suppress = True: Rep.hMrpDisc.Suppress = True
    Rep.OdMrp1.Suppress = True: Rep.wOdMrpDisc1.Suppress = True
    Rep.hMrp.Top = 0: Rep.hMrpDisc.Top = 0: Rep.OdMrp1.Top = 0: Rep.wOdMrpDisc1.Top = 0
    Rep.DetDsg.Height = 221
  End If
  'sv.63 QR code & IRN will display for local invoice report in jemr
  'QR code & IRn added in emr also
  If adc("UOYN9") = "Y" And (UCase(adc.MenuCd) = UCase("RepJtLclInv") Or UCase(adc.MenuCd) = UCase("RepLclInv")) Then
    Rep.PicQRCode.Suppress = False: Rep.wInIRN1.Suppress = False
  End If
  '62 HUID details will display only when show huid=Y and detail level is Sr
  '21 RepJtLclInv added to show HUID no
  If adc("UOYN10") = "Y" And adc("UoInDet") = "INVSR" And (UCase(adc.MenuCd) = UCase("RepLclInv") Or UCase(adc.MenuCd) = UCase("RepJtLclInv")) Then
    Rep.qCertNo1.Suppress = False: Rep.TxtHUIDNo.Suppress = False
    If adc("UOYN5") = "N" And adc("UOYN8") = "N" And UCase(adc.MenuCd) = UCase("RepLclInv") Then
      Rep.qCertNo1.Top = 217: Rep.TxtHUIDNo.Top = 217:  Rep.DetDsg.Height = 403
    End If
  Else
    Rep.TxtHUIDNo.Suppress = True:  Rep.TxtHUIDNo.Top = 0
    Rep.qCertNo1.Suppress = True: Rep.qCertNo1.Top = 0: Rep.DetDsg.Height = 221
  End If
  '37 if show value option is yes then amounts & value in text will be suppress
  If adc("wShowValYN") <> "Y" Then
    Rep.IdIFob1.Suppress = True:        Rep.wInFob1.Suppress = True
    Rep.InAmt11.Suppress = True:        Rep.InAmt21.Suppress = True:  Rep.InAmt31.Suppress = True
    Rep.InAmt41.Suppress = True:        Rep.InAmt51.Suppress = True:  Rep.InAmt61.Suppress = True
    Rep.InAmt71.Suppress = True:        Rep.InAmt81.Suppress = True:  Rep.InAmt91.Suppress = True
    Rep.InAmt101.Suppress = True:       Rep.InAmt111.Suppress = True: Rep.InAmt121.Suppress = True
    Rep.InBaseAmt.Suppress = True:      Rep.wWords1.Suppress = True
    Rep.InLclOctroiRs1.Suppress = True: Rep.InLclOthrValRs1.Suppress = True
    Rep.InLclCstRs1.Suppress = True:    Rep.FldInTotRs.Suppress = True:     Rep.InLclVatRs1.Suppress = True
  End If
  '38 when show bag/stkid is yes and  detail level is Sr then bag details will display
  If adc("UOYN12") = "Y" And adc("UoInDet") = "INVSR" Then
    If adc("UOYN5") <> "Y" And adc("UOYN8") <> "Y" And adc("UOYN10") <> "Y" Then
      If UCase(adc.MenuCd) = UCase("RepJtLclInv") Then
        Rep.wBagStkNo1.Top = 403: Rep.TxtBagStkId.Top = 403: Rep.DetBagStkId.Height = 689
        Rep.LineBagStkId.Top = 469: Rep.LineSecBagStkId.Height = 469
      Else
        Rep.wBagStkNo1.Top = 207: Rep.TxtBagStkId.Top = 207: Rep.DetBagStkId.Height = 413
        Rep.LineBagStkId.Top = 193: Rep.LineSecBagStkId.Height = 193
      End If
    ElseIf (adc("UOYN5") = "N" And adc("UOYN8") = "N") And adc("UOYN10") = "Y" Then
'21 RepJtLclInv part removed and RepLclInv to show HUID no
'      If UCase(ADC.MenuCd) = UCase("RepJtLclInv") Then
'        Rep.wBagStkNo1.Top = 207: Rep.TxtBagStkId.Top = 207: Rep.DetBagStkId.Height = 413
'        Rep.LineBagStkId.Top = 193: Rep.LineSecBagStkId.Height = 193
'      Else
      If UCase(adc.MenuCd) = UCase("RepLclInv") Then
        Rep.wBagStkNo1.Top = 403: Rep.TxtBagStkId.Top = 403: Rep.DetBagStkId.Height = 689
        Rep.LineBagStkId.Top = 469: Rep.LineSecBagStkId.Height = 469
      End If
    ElseIf (adc("UOYN5") = "Y" Or adc("UOYN8") = "Y") And adc("UOYN10") <> "Y" Then
      Rep.wBagStkNo1.Top = 403: Rep.TxtBagStkId.Top = 403: Rep.DetBagStkId.Height = 689
      Rep.LineBagStkId.Top = 469: Rep.LineSecBagStkId.Height = 469
    End If
  Else
    Rep.DetBagStkId.Suppress = True:  Rep.LineSecBagStkId.Suppress = True
  End If
  '578.19 calling function to change Rmcode font size
  SetSubFormula Rep.Subreport2, "wRmCdFontSz", GetRmcdFontSz("{rdo.RmCd}", 8)
  
  CRV_REP.DisplayGroupTree = False
  Call DispReport(adc) '****** Sachin 3.02 - Display Report ******
Exit Sub
RepErr:
  Call ChkSelScope(Err.Number, Err.Description)
End Sub

' ****** Manali 3.5.0 - 19/11/08 - Check Boxes added
Private Sub ChkBoxArr_Click(Index As Integer)
With ChkBoxArr(Index)
  Select Case Index
  Case Is = 3       ' Tag = "ShowDemarc"
    If .Value = Checked Then adc("UoYN3") = "Y" Else adc("UoYN3") = "N"
  ' ***** Manali 3.6.0 - 24/11/09 - Show Run Sr
  Case Is = 0       ' Tag = "ShowRunSr"
    If .Value = Checked Then adc("UoYN") = "Y" Else adc("UoYN") = "N"
  ' ***** Manali 3.6.0 - 24/11/09 - Show Run Sr
  ' ***** Manali 3.8.0 - Sort By Design
  Case Is = 1       ' Tag = "SortByDmCd"
    If .Value = Checked Then adc("UoYN1") = "Y" Else adc("UoYN1") = "N"
  Case Is = 2       ' Tag = ""
    If .Value = Checked Then adc("UoYN2") = "Y" Else adc("UoYN2") = "N"
  
  ' ***** Manali 3.10.0 - 12/04/12 - Show Rm Summary
  Case Is = 4       ' Tag = "ShowRmSumm"
    If .Value = Checked Then adc("UoYN4") = "Y" Else adc("UoYN4") = "N"
  
  Case Is = 5       ' Tag = ""      4.1.4
    If .Value = Checked Then adc("UoYN5") = "Y" Else adc("UoYN5") = "N"
  Case Is = 6       ' Tag = ""
    If .Value = Checked Then adc("UoYN6") = "Y" Else adc("UoYN6") = "N"
  Case Is = 7
    If .Value = Checked Then adc("UoYN7") = "Y" Else adc("UoYN7") = "N"
  Case Is = 8
    If .Value = Checked Then adc("UoYN8") = "Y" Else adc("UoYN8") = "N"
  Case Is = 9
    If .Value = Checked Then adc("UoYN9") = "Y" Else adc("UoYN9") = "N" 'sv.63 show QRcode/IRN
  Case Is = 10
    If .Value = Checked Then adc("UoYN10") = "Y" Else adc("UoYN10") = "N" '62 show HUID
  Case Is = 11
    If .Value = Checked Then adc("wShowValYN") = "Y" Else adc("wShowValYN") = "N" '37 show value
  Case Is = 12
    If .Value = Checked Then adc("UoYN12") = "Y" Else adc("UoYN12") = "N" '38 show Bag/StkId
    
  End Select
End With
End Sub
' ****** Manali 3.5.0 - 19/11/08 - Check Boxes added

' ****** Manali 3.5.0 - 19/11/08 - Check Boxes added
Private Sub ChkBoxArr_GotFocus(Index As Integer)
  DispMsg ChkBoxArr(Index).ToolTipText, etInfo
End Sub
' ****** Manali 3.5.0 - 19/11/08 - Check Boxes added

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
  OptNewRpt.Value = True
  OptOldRpt.Value = False
End Sub
Private Sub ADC_Load()
    Set moCn = adc.Connection    '*** (Bef speed) bef 24/11/06
    Set adc.FirNKeyCtl = adc("UoInTcFr")
    Call InitProp(Me)
    
    '*** Jay 3.3(FG)
    gs_InTcTyp = IIF(UCase(adc.MenuCd) <> UCase("RepFgLclInv"), "IN", "FG")
    '*** Jay 3.3(FG)
    
    ' ***** Manali Trading Module
    gs_InTcTyp = IIF(UCase(adc.MenuCd) = UCase("RepJtLclInv"), ctJtTcTyp, gs_InTcTyp)
    ' ***** Manali Trading Module
    
    ' **** Manali 3.7.1
    gb_InHlpFrOm = IIF(UCase(adc.MenuCd) = UCase("RepJtLclInv"), True, False)
    
    gs_RmGrp = "INV"
    gs_Tbl = "Param"
    gs_PTyp = "GRLSSUMM"
    
    '4.1.2
    If UCase(adc.MenuCd) = UCase("RepInvLclRM") Then
        ChkBoxArr(0).Visible = False
        ChkBoxArr(1).Visible = False
        ChkBoxArr(2).Visible = False
        ChkBoxArr(3).Visible = False
        ChkBoxArr(4).Visible = False
        LblRepFormat.Visible = False: Frame1.Visible = False
        'Hk.5 Group Combination and Detail Level scope newly added for Local Invoice
        LblGrpCombo.Visible = False:  adc("UOINGRP").Visible = False:
        LblDet.Visible = False: adc("UOINDET").Visible = False
        ChkBoxArr(7).Visible = False
        'pg.5-Show MRP option added for local invoice
        ChkBoxArr(8).Visible = False
        ChkBoxArr(11).Visible = False   '37 show value option not visible for localRM
        ChkBoxArr(12).Visible = False   '38 show Bag/StkId option not visible for localRM
    End If
    
    If UCase(adc.MenuCd) <> UCase("RepLclInv") Then
        ChkBoxArr(6).Visible = False
        'ChkBoxArr(10).Visible = False   '62 show huid option will display only in local invoice
    End If
    'sv.63 show QRcode/IRN option only for jt localinvoice
    'this option also added in emr
    If UCase(adc.MenuCd) <> UCase("RepJtLclInv") And UCase(adc.MenuCd) <> UCase("RepLclInv") Then ChkBoxArr(9).Visible = False
    
    '21 HUID option will display for emr & jemr local invoice
    If UCase(adc.MenuCd) <> UCase("RepLclInv") And UCase(adc.MenuCd) <> UCase("RepJtLclInv") Then _
      ChkBoxArr(10).Visible = False
 
    Call SetGroupSort
End Sub
Private Sub ATXT_FldChange(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String, ByVal pv_NewValue As Variant)
  '*** Jay 3.3(FG)
  gs_InTcTyp = IIF(UCase(adc.MenuCd) <> UCase("RepFgLclInv"), "IN", "FG")
  '*** Jay 3.3(FG)
  ' ***** Manali Trading Module
  gs_InTcTyp = IIF(UCase(adc.MenuCd) = UCase("RepJtLclInv"), ctJtTcTyp, gs_InTcTyp)
  ' ***** Manali Trading Module
  gs_PTyp = "GRLSSUMM"
  
  ' **** Manali 3.7.1
  gb_InHlpFrOm = IIF(UCase(adc.MenuCd) = UCase("RepJtLclInv"), True, False)
  
  Call SetProp(Me, IdName, Valid, Cancel, ErrMsg)
End Sub
Private Sub ATXT_FldWhen(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String)
  '*** Jay 3.3(FG)
  gs_InTcTyp = IIF(UCase(adc.MenuCd) <> UCase("RepFgLclInv"), "IN", "FG")
  '*** Jay 3.3(FG)
  ' ***** Manali Trading Module
  gs_InTcTyp = IIF(UCase(adc.MenuCd) = UCase("RepJtLclInv"), ctJtTcTyp, gs_InTcTyp)
  ' ***** Manali Trading Module
  gs_RmGrp = "INV"
  gs_Tbl = "Param"
  gs_PTyp = "GRLSSUMM"
 
  ' **** Manali 3.7.1
  gb_InHlpFrOm = IIF(UCase(adc.MenuCd) = UCase("RepJtLclInv"), True, False)
 
  SetProp Me, IdName, When
  

  Select Case UCase(IdName)
  Case Is = UCase("UoPMCdTo")
    Call HlpList.PMCd("INVDESC")
  End Select
End Sub



Private Sub OptNewRpt_Click()

OptNewRpt.Value = True
OptOldRpt.Value = False
End Sub

Private Sub OptOldRpt_Click()
OptOldRpt.Value = True
OptNewRpt.Value = False
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
  If UCase(adc.MenuCd) = UCase("RepInv") Or UCase(adc.MenuCd) = UCase("RepInvLoose") Then _
      adc("UOYN2") = moCn.GetFldVal("Select hInRepMetLs from Head where " + _
                   "HCoCd = '" + adc("UoCoCdFr") + "' and HCd= '" + ctSelfCmCd + "'")
                   
  ' ****** Manali 3.5.0 - 19/11/08 - Check Boxes added
  If adc("UoYN3") = "" Then adc("UoYN3") = "N"
  If adc("UoYN") = "" Then adc("UoYN") = "N"    ' ***** Manali 3.6.0 - 24/11/09 - Show Run Sr
  If adc("UoYN1") = "" Then adc("UoYN1") = "N"    ' ***** Manali 3.8.0 - Sort By Design
  If adc("UoYN2") = "" Then adc("UoYN2") = "N"
  If adc("UoYN5") = "" Then adc("UoYN5") = "N"
  If adc("UoYN6") = "" Then adc("UoYN6") = "N"
  'SRITMP - Always disable Demarcation based on discussion w Rn for danda breaking (part of 5.7 QC)
'  If adc("UoYN3") = "Y" Then ChkBoxArr(3).Value = Checked Else ChkBoxArr(3).Value = Unchecked
  ChkBoxArr(3).Value = Checked
  ChkBoxArr(3).Enabled = False
  If adc("UoYN") = "Y" Then ChkBoxArr(0).Value = Checked Else ChkBoxArr(0).Value = Unchecked    ' ***** Manali 3.6.0 - 24/11/09 - Show Run Sr
  If adc("UoYN1") = "Y" Then ChkBoxArr(1).Value = Checked Else ChkBoxArr(1).Value = Unchecked    ' ***** Manali 3.8.0 - Sort By Design
  If adc("UoYN2") = "Y" Then ChkBoxArr(2).Value = Checked Else ChkBoxArr(2).Value = Unchecked
  If adc("UoYN4") = "Y" Then ChkBoxArr(4).Value = Checked Else ChkBoxArr(4).Value = Unchecked     ' ***** Manali 3.10.0 - 12/04/12 - Show Rm Summary
  If adc("UoYN5") = "Y" Then ChkBoxArr(5).Value = Checked Else ChkBoxArr(5).Value = Unchecked     ' 4.1.4
  If adc("UoYN6") = "Y" Then ChkBoxArr(6).Value = Checked Else ChkBoxArr(6).Value = Unchecked
  If adc("UoYN7") = "Y" Then ChkBoxArr(7).Value = Checked Else ChkBoxArr(7).Value = Unchecked
  If adc("UoYN8") = "Y" Then ChkBoxArr(8).Value = Checked Else ChkBoxArr(8).Value = Unchecked
  If adc("UoYN9") = "Y" Then ChkBoxArr(9).Value = Checked Else ChkBoxArr(9).Value = Unchecked 'sv.63 show QRcode/IRN
  If adc("UoYN10") = "Y" Then ChkBoxArr(10).Value = Checked Else ChkBoxArr(10).Value = Unchecked  '62 show HUID
  adc("wShowValYN") = "Y": ChkBoxArr(11).Value = Checked    '37 default value for show values is yes
  If adc("UoYN12") = "Y" Then ChkBoxArr(12).Value = Checked Else ChkBoxArr(12).Value = Unchecked  '38 show Bag/StkId
  ' ****** Manali 3.5.0 - 19/11/08 - Check Boxes added
End Sub
Private Sub DispCoNm()
  Dim ws_HName As String
  ws_HName = moCn.GetFldVal("Select HName from Head where HCoCd = '" + adc("UoCoCdFr") + "' and " + _
                            "HCd = '" + ctSelfCmCd + "' ")
  'mRep.wCoCd.SetText ws_HName
  Dim ws_DtTm As String
  ws_DtTm = Format$(moCn.SrvrDate, IIF(UCase(MWLib.DtSetting) = "MM/DD/YY", "MMM-DD-YYYY", "DD-MMM-YYYY"))
  'Rep.wSrvrDate.SetText ws_DtTm
  ws_DtTm = Replace(Format(IIF(CDbl(moCn.SrvrTime) >= 13#, moCn.SrvrTime - 12, IIF(CDbl(moCn.SrvrTime) < 1#, moCn.SrvrTime + 12, moCn.SrvrTime)), "00.00"), ".", ":") + IIF(CDbl(moCn.SrvrTime) >= 12#, " PM", " AM")
  'Rep.wSrvrTime.SetText ws_DtTm
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
    '*** (09/08/05)
End Sub

Private Sub GetCnd()

Select Case UCase(adc.MenuCd)
Case Is = UCase("RepJtLclInv")
  adc("UoCoCdFr").CmpStr = "OmCoCd = "
  adc("UoInTcFr").CmpStr = "OmTc >= ":  adc("UoInYyFr").CmpStr = "OmYY >= "
  adc("UoInChrFr").CmpStr = "OmChr >= ":  adc("UoInNoFr").CmpStr = "OmNo >= "
  adc("UoInTcTo").CmpStr = "OmTc <= ":  adc("UoInYyTo").CmpStr = "OmYY <= "
  adc("UoInChrTo").CmpStr = "OmChr <= ":  adc("UoInNoTo").CmpStr = "OmNo <= "
Case Else
  adc("UoCoCdFr").CmpStr = "InCoCd = "
  adc("UoInTcFr").CmpStr = "InTc >= ":  adc("UoInYyFr").CmpStr = "InYy >= "
  adc("UoInChrFr").CmpStr = "InChr >= ":  adc("UoInNoFr").CmpStr = "InNo >= "
  adc("UoInTcTo").CmpStr = "InTc <= ":  adc("UoInYyTo").CmpStr = "InYy <= "
  adc("UoInChrTo").CmpStr = "InChr <= ":  adc("UoInNoTo").CmpStr = "InNo <= "
End Select
End Sub




