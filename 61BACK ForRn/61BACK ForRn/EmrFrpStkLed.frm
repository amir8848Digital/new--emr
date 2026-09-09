VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{AB2CC32A-E652-42A2-B341-26ECE02F6386}#2.1#0"; "mwfctlforspeed9.ocx"
Begin VB.Form EmrFrpStkLed 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Stock Ledger"
   ClientHeight    =   10185
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15270
   Enabled         =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10185
   ScaleWidth      =   15270
   WindowState     =   2  'Maximized
   Begin MwfCtl.MWCTL_HLP MWCTL_HLP1 
      Height          =   495
      Left            =   11010
      TabIndex        =   57
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
      TabIndex        =   58
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
      Left            =   4965
      TabIndex        =   67
      Top             =   9630
      Width           =   3555
      Begin MwfCtl.MWCTL_MED aktxt 
         Height          =   345
         Index           =   0
         Left            =   210
         TabIndex        =   68
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
         TabIndex        =   69
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
      Left            =   6180
      TabIndex        =   66
      Top             =   9645
      Visible         =   0   'False
      Width           =   855
      Begin CRVIEWER9LibCtl.CRViewer9 CRV_REP 
         Height          =   495
         Left            =   480
         TabIndex        =   63
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
      TabIndex        =   59
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
      Height          =   9705
      Left            =   0
      TabIndex        =   62
      Top             =   60
      Width           =   15255
      Begin TabDlg.SSTab TAB_REP 
         Height          =   9465
         Left            =   15
         TabIndex        =   61
         Top             =   30
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   16695
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "&Data Range"
         TabPicture(0)   =   "EmrFrpStkLed.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "fra_tabrep(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "&Group Sort"
         TabPicture(1)   =   "EmrFrpStkLed.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "fra_tabrep(1)"
         Tab(1).ControlCount=   1
         Begin VB.Frame fra_tabrep 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   9070
            Index           =   1
            Left            =   -74940
            TabIndex        =   65
            Top             =   360
            Width           =   15105
            Begin MwfCtl.MWCTL_GRP GRP_REP 
               Height          =   360
               Left            =   240
               TabIndex        =   60
               Top             =   360
               Width           =   10245
               _ExtentX        =   18071
               _ExtentY        =   635
            End
         End
         Begin VB.Frame fra_tabrep 
            BorderStyle     =   0  'None
            Height          =   9070
            Index           =   0
            Left            =   120
            TabIndex        =   64
            Top             =   480
            Width           =   14985
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Show Stock Rate "
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
               Left            =   0
               TabIndex        =   100
               Tag             =   "ShowPic"
               ToolTipText     =   "Check To Show Stock Rate"
               Top             =   6150
               Width           =   1995
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Show Custom Pure Wt"
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
               Index           =   1
               Left            =   30
               TabIndex        =   54
               Tag             =   "ShowZWt"
               ToolTipText     =   "Check To Show ZPurity Wt"
               Top             =   8160
               Width           =   1950
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Show Bal"
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
               TabIndex        =   55
               Tag             =   "ShowDesc"
               ToolTipText     =   "Check To Show Openning/Closing Balance"
               Top             =   7920
               Width           =   1950
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Show Desc       "
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
               TabIndex        =   53
               Tag             =   "ShowDesc"
               ToolTipText     =   "Check To Show Description"
               Top             =   7635
               Width           =   1950
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Reverse Logic  "
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
               TabIndex        =   51
               Tag             =   "RevLogic"
               ToolTipText     =   "Check To Select Reverse Logic"
               Top             =   7320
               Width           =   1950
            End
            Begin VB.Frame FraOptLastCol 
               BorderStyle     =   0  'None
               Height          =   285
               Left            =   1755
               TabIndex        =   56
               ToolTipText     =   "Select Value to be Displayed in Last Column"
               Top             =   6780
               Width           =   3615
               Begin VB.OptionButton OptLastCol 
                  Caption         =   "Blank"
                  BeginProperty Font 
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
                  Left            =   2520
                  TabIndex        =   49
                  ToolTipText     =   "Select Value to be Displayed in Last Column"
                  Top             =   0
                  Width           =   945
               End
               Begin VB.OptionButton OptLastCol 
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
                  Index           =   1
                  Left            =   1260
                  TabIndex        =   48
                  ToolTipText     =   "Select Value to be Displayed in Last Column"
                  Top             =   0
                  Width           =   885
               End
               Begin VB.OptionButton OptLastCol 
                  Caption         =   "Qty/Wt"
                  BeginProperty Font 
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
                  TabIndex        =   47
                  ToolTipText     =   "Select Value to be Displayed in Last Column"
                  Top             =   0
                  Width           =   975
               End
            End
            Begin VB.Frame FraOptQW 
               BorderStyle     =   0  'None
               Height          =   285
               Left            =   1755
               TabIndex        =   96
               ToolTipText     =   "Select either Quantity Or Weight"
               Top             =   6495
               Width           =   3195
               Begin VB.OptionButton OptQW 
                  Caption         =   "Qty"
                  BeginProperty Font 
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
                  Left            =   1260
                  TabIndex        =   45
                  ToolTipText     =   "Select either Quantity Or Weight"
                  Top             =   0
                  Width           =   825
               End
               Begin VB.OptionButton OptQW 
                  Caption         =   "Wt"
                  BeginProperty Font 
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
                  Left            =   0
                  TabIndex        =   44
                  ToolTipText     =   "Select either Quantity Or Weight"
                  Top             =   0
                  Width           =   765
               End
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   1
               Left            =   4155
               TabIndex        =   8
               ToolTipText     =   "Enter To FrBag Location"
               Top             =   1575
               Width           =   1830
               _ExtentX        =   3228
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UOISSLOCTO"
               IdName          =   "UOISSLOCTO"
               CmpStr          =   "LocCd<= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   2
               Left            =   6540
               TabIndex        =   9
               ToolTipText     =   "Enter FrBag Location Selection"
               Top             =   1575
               Width           =   8340
               _ExtentX        =   14711
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOISSLOCSEL"
               IdName          =   "UOISSLOCSEL"
               CmpStr          =   "LocCd In "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   40
               Left            =   4155
               TabIndex        =   37
               ToolTipText     =   "Enter To Customer (Customer As Mentioned In Rm Master)"
               Top             =   5010
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UOCMCDTO"
               IdName          =   "UOCMCDTO"
               CmpStr          =   "RmCmCd <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   41
               Left            =   6540
               TabIndex        =   38
               ToolTipText     =   "Enter Customer Selection (Customer As Mentioned In Rm Master)"
               Top             =   5010
               Width           =   8340
               _ExtentX        =   14711
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOCMCDSEL"
               IdName          =   "UOCMCDSEL"
               CmpStr          =   "RmCmCd In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   42
               Left            =   1755
               TabIndex        =   36
               ToolTipText     =   "Enter From Customer (Customer As Mentioned In Rm Master)"
               Top             =   4995
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UOCMCDFR"
               IdName          =   "UOCMCDFR"
               CmpStr          =   "RmCmCd >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   46
               Left            =   4155
               TabIndex        =   34
               ToolTipText     =   "Enter To Voucher Date"
               Top             =   4440
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   3
               DataField       =   "UODTTO"
               ReCalcParent    =   "UOYN5"
               IdName          =   "UODTTO"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   48
               Left            =   1755
               TabIndex        =   33
               ToolTipText     =   "Enter From Voucher Date"
               Top             =   4425
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   3
               DataField       =   "UODTFR"
               ReCalcParent    =   "UOYN5"
               IdName          =   "UODTFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   0
               Left            =   1755
               TabIndex        =   7
               ToolTipText     =   "Enter From FrBag Location"
               Top             =   1575
               Width           =   1830
               _ExtentX        =   3228
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UOISSLOCFR"
               IdName          =   "UOISSLOCFR"
               CmpStr          =   "LocCd>= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   8
               Left            =   4155
               TabIndex        =   21
               ToolTipText     =   "Enter Raw Material Code To Range"
               Top             =   3015
               Width           =   2385
               _ExtentX        =   4207
               _ExtentY        =   503
               MaxLength       =   16
               DataField       =   "UORMCDTO"
               IdName          =   "UORMCDTO"
               CmpStr          =   "RmCd <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   4
               Left            =   4155
               TabIndex        =   15
               ToolTipText     =   "Enter Raw Material Sub Category To Range"
               Top             =   2445
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataField       =   "UORMSCTGTO"
               IdName          =   "UORMSCTGTO"
               CmpStr          =   "RmSCtg <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   6
               Left            =   1755
               TabIndex        =   10
               ToolTipText     =   "Enter From Raw Material Category"
               Top             =   1860
               Width           =   465
               _ExtentX        =   820
               _ExtentY        =   503
               MaxLength       =   2
               DataField       =   "UORMCTGFR"
               IdName          =   "UORMCTGFR"
               CmpStr          =   "RmCtg ="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   7
               Left            =   1755
               TabIndex        =   14
               ToolTipText     =   "Enter Raw Material Sub Category From Range"
               Top             =   2430
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataField       =   "UORMSCTGFR"
               IdName          =   "UORMSCTGFR"
               Data            =   "ABCDW"
               CmpStr          =   "RmSCtg >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   9
               Left            =   1755
               TabIndex        =   20
               ToolTipText     =   "Enter Raw Material Code From Range"
               Top             =   3000
               Width           =   2385
               _ExtentX        =   4207
               _ExtentY        =   503
               MaxLength       =   16
               DataField       =   "UORMCDFR"
               IdName          =   "UORMCDFR"
               CmpStr          =   "RmCd >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   10
               Left            =   4155
               TabIndex        =   30
               ToolTipText     =   "Enter To Raw Material Len/Sv"
               Top             =   3870
               Width           =   1125
               _ExtentX        =   1984
               _ExtentY        =   503
               Alignment       =   2
               Mask            =   "#0.000"
               MaxLength       =   6
               DataType        =   2
               DataField       =   "UOLN1TO"
               IdName          =   "UOLN1TO"
               CmpStr          =   "TdRmSz<= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   11
               Left            =   1755
               TabIndex        =   29
               ToolTipText     =   "Enter From Raw Material Len/Sv"
               Top             =   3855
               Width           =   1005
               _ExtentX        =   1773
               _ExtentY        =   503
               Alignment       =   2
               Mask            =   "#0.000"
               MaxLength       =   6
               DataType        =   2
               DataField       =   "UOLN1FR"
               IdName          =   "UOLN1FR"
               CmpStr          =   "TdRmSz>= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   16
               Left            =   4155
               TabIndex        =   32
               ToolTipText     =   "Enter To Cost Rate"
               Top             =   4155
               Width           =   1245
               _ExtentX        =   2196
               _ExtentY        =   503
               Alignment       =   2
               MaxLength       =   0
               DataType        =   2
               DataField       =   "UOCSTRTTO"
               IdName          =   "UOCSTRTTO"
               CmpStr          =   "TdRmStkRt<= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   18
               Left            =   1755
               TabIndex        =   31
               ToolTipText     =   "Enter From Cost Rate"
               Top             =   4140
               Width           =   1245
               _ExtentX        =   2196
               _ExtentY        =   503
               Alignment       =   2
               MaxLength       =   0
               DataType        =   2
               DataField       =   "UOCSTRTFR"
               IdName          =   "UOCSTRTFR"
               CmpStr          =   "TdRmStkRt>= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   3
               Left            =   1755
               TabIndex        =   3
               ToolTipText     =   "Enter Location Type"
               Top             =   990
               Width           =   495
               _ExtentX        =   873
               _ExtentY        =   503
               MaxLength       =   2
               DataType        =   4
               DataField       =   "UOLOCTYPFR"
               IdName          =   "UOLOCTYPFR"
               CmpStr          =   "LocTyp = "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   12
               Left            =   1755
               TabIndex        =   39
               ToolTipText     =   "Enter Type Of Detail Record (Voucher, Voucher Sr, Voucher SrNo)"
               Top             =   5280
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               DataField       =   "UOVCHOPT"
               IdName          =   "UOVCHOPT"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   14
               Left            =   4155
               TabIndex        =   27
               ToolTipText     =   "Enter To Raw Material Range Size"
               Top             =   3585
               Width           =   1605
               _ExtentX        =   2831
               _ExtentY        =   503
               MaxLength       =   10
               DataField       =   "UORNGSZTO"
               IdName          =   "UORNGSZTO"
               CmpStr          =   "RrRngSz <= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   19
               Left            =   1755
               TabIndex        =   26
               ToolTipText     =   "Enter From Raw Material Range Size"
               Top             =   3570
               Width           =   1605
               _ExtentX        =   2831
               _ExtentY        =   503
               MaxLength       =   10
               DataField       =   "UORNGSZFR"
               IdName          =   "UORNGSZFR"
               CmpStr          =   "RrRngSz >= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   20
               Left            =   4155
               TabIndex        =   18
               ToolTipText     =   "Enter To Parent Code"
               Top             =   2730
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   4
               DataField       =   "UOPARENTCDTO"
               IdName          =   "UOPARENTCDTO"
               CmpStr          =   "RmParentCd <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   22
               Left            =   1755
               TabIndex        =   17
               ToolTipText     =   "Enter From Parent Code"
               Top             =   2715
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   4
               DataField       =   "UOPARENTCDFR"
               IdName          =   "UOPARENTCDFR"
               CmpStr          =   "RmParentCd >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   35
               Left            =   1755
               TabIndex        =   0
               ToolTipText     =   "Enter Company Code"
               Top             =   705
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOCOCDFR"
               IdName          =   "UOCOCDFR"
               CmpStr          =   "LocCoCd >= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   25
               Left            =   1755
               TabIndex        =   35
               ToolTipText     =   "Enter Open YearMonth"
               Top             =   4710
               Width           =   765
               _ExtentX        =   1349
               _ExtentY        =   503
               Mask            =   "#####0"
               MaxLength       =   6
               DataType        =   1
               DataField       =   "UOOPNYYMM"
               IdName          =   "UOOPNYYMM"
            End
            Begin VB.TextBox TxtPwd 
               Height          =   285
               IMEMode         =   3  'DISABLE
               Left            =   1755
               MaxLength       =   30
               PasswordChar    =   "*"
               TabIndex        =   40
               Top             =   5565
               Width           =   2340
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   29
               Left            =   1755
               TabIndex        =   41
               ToolTipText     =   "Specify If The Value To Be Shown Should Be Stock Value, Cost Value Or Sale Value ?"
               Top             =   5850
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UOPMCDFR"
               ReCalcParent    =   "UOYN"
               IdName          =   "UOPMCDFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   23
               Left            =   5595
               TabIndex        =   43
               ToolTipText     =   "Enter 'Y' for Quantity and 'N' for Weight"
               Top             =   6495
               Visible         =   0   'False
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               Enabled         =   0   'False
               DataType        =   4
               DataField       =   "UOYN3"
               ReCalcParent    =   "UOYN"
               IdName          =   "UOYN3"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   24
               Left            =   2355
               TabIndex        =   50
               ToolTipText     =   "Enter 'Y' to apply Reverse Logic else enter 'N' for normal logic"
               Top             =   7350
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
               Index           =   32
               Left            =   1755
               TabIndex        =   23
               ToolTipText     =   "Enter From Lot No."
               Top             =   3285
               Width           =   1905
               _ExtentX        =   3360
               _ExtentY        =   503
               MaxLength       =   12
               DataType        =   4
               DataField       =   "UOLOTNOFR"
               IdName          =   "UOLOTNOFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   26
               Left            =   4155
               TabIndex        =   24
               ToolTipText     =   "Enter To Lot No."
               Top             =   3300
               Width           =   1905
               _ExtentX        =   3360
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   4
               DataField       =   "UOLOTNOTO"
               IdName          =   "UOLOTNOTO"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   61
               Left            =   4155
               TabIndex        =   12
               ToolTipText     =   "Enter Raw Material Stone Class To Range"
               Top             =   2160
               Width           =   1365
               _ExtentX        =   2408
               _ExtentY        =   503
               MaxLength       =   5
               DataField       =   "UOSTNCLSTO"
               IdName          =   "UOSTNCLSTO"
               CmpStr          =   "ISNULL(sc.PSCd,'') <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   63
               Left            =   1755
               TabIndex        =   11
               ToolTipText     =   "Enter Raw Material Stone Class From Range"
               Top             =   2145
               Width           =   1365
               _ExtentX        =   2408
               _ExtentY        =   503
               MaxLength       =   5
               DataField       =   "UOSTNCLSFR"
               IdName          =   "UOSTNCLSFR"
               CmpStr          =   "ISNULL(sc.PSCd,'') >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   15
               Left            =   6540
               TabIndex        =   22
               ToolTipText     =   "Enter Raw Material Code Selection"
               Top             =   3000
               Width           =   8340
               _ExtentX        =   14711
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UORMCDSEL"
               IdName          =   "UORMCDSEL"
               CmpStr          =   "RmCd In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   5
               Left            =   6540
               TabIndex        =   16
               ToolTipText     =   "Enter Raw Material Sub Category Selection"
               Top             =   2445
               Width           =   8340
               _ExtentX        =   14711
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UORMSCTGSEL"
               IdName          =   "UORMSCTGSEL"
               CmpStr          =   "RmSCtg In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   17
               Left            =   6540
               TabIndex        =   28
               ToolTipText     =   "Enter Raw Material Range Size Selection"
               Top             =   3585
               Width           =   8340
               _ExtentX        =   14711
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UORNGSZSEL"
               IdName          =   "UORNGSZSEL"
               CmpStr          =   "RrRngSz In "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   21
               Left            =   6540
               TabIndex        =   19
               ToolTipText     =   "Enter Parent Code Selection"
               Top             =   2730
               Width           =   8340
               _ExtentX        =   14711
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOPARENTCDSEL"
               IdName          =   "UOPARENTCDSEL"
               CmpStr          =   "RmParentCd In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   13
               Left            =   6540
               TabIndex        =   25
               ToolTipText     =   "Enter Lot No. Selection"
               Top             =   3300
               Width           =   8340
               _ExtentX        =   14711
               _ExtentY        =   503
               MaxLength       =   255
               DataType        =   4
               DataField       =   "UOLOTNOSEL"
               IdName          =   "UOLOTNOSEL"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   62
               Left            =   6540
               TabIndex        =   13
               ToolTipText     =   "Enter Raw Material Stone Class Selection"
               Top             =   2160
               Width           =   8340
               _ExtentX        =   14711
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOSTNCLSSEL"
               IdName          =   "UOSTNCLSSEL"
               CmpStr          =   "ISNULL(sc.PSCd,'') In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   27
               Left            =   5595
               TabIndex        =   46
               ToolTipText     =   "Enter 'Y' to Show Qty/ Wt Instead Of Value Else Enter 'N' For Not Showing Anything In The Column"
               Top             =   6780
               Visible         =   0   'False
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               Enabled         =   0   'False
               DataType        =   4
               DataField       =   "UOYN"
               ReCalcOn        =   "UOPMCDFR,UOYN3"
               IdName          =   "UOYN"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   28
               Left            =   2355
               TabIndex        =   52
               ToolTipText     =   "Show Description (Yes/No)"
               Top             =   7635
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
               Index           =   47
               Left            =   4155
               TabIndex        =   1
               ToolTipText     =   "Enter To Company Code"
               Top             =   690
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOCOCDTO"
               IdName          =   "UOCOCDTO"
               CmpStr          =   "LocCoCd <= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   50
               Left            =   6570
               TabIndex        =   2
               ToolTipText     =   "Enter Design Code Selection"
               Top             =   690
               Width           =   8280
               _ExtentX        =   14605
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOCOCDSEL"
               IdName          =   "UOCOCDSEL"
               CmpStr          =   "LocCoCd In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   30
               Left            =   4155
               TabIndex        =   5
               ToolTipText     =   "Enter To Parent Code"
               Top             =   1290
               Width           =   1830
               _ExtentX        =   3228
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UOPRNTCDTO"
               IdName          =   "UOPRNTCDTO"
               CmpStr          =   "LocPrntCd <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   49
               Left            =   6540
               TabIndex        =   6
               ToolTipText     =   "Enter Parent Code Selection"
               Top             =   1290
               Width           =   8340
               _ExtentX        =   14711
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOPRNTCDSEL"
               IdName          =   "UOPRNTCDSEL"
               CmpStr          =   "LocPrntCd In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   51
               Left            =   1755
               TabIndex        =   4
               ToolTipText     =   "Enter From Parent Code"
               Top             =   1290
               Width           =   1830
               _ExtentX        =   3228
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UOPRNTCDFR"
               IdName          =   "UOPRNTCDFR"
               CmpStr          =   "LocPrntCd >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   31
               Left            =   2355
               TabIndex        =   97
               ToolTipText     =   "Show Description (Yes/No)"
               Top             =   7920
               Visible         =   0   'False
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               Enabled         =   0   'False
               DataType        =   4
               DataField       =   "UOYN5"
               ReCalcOn        =   "UODTFR,UODTTO"
               IdName          =   "UOYN5"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   84
               Left            =   6540
               TabIndex        =   42
               ToolTipText     =   "Specify If The Other Value To Be Shown Should Be Stock Value, Cost Value Or Sale Value"
               Top             =   5850
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UOPMCDTO"
               IdName          =   "UOPMCDTO"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   37
               Left            =   2355
               TabIndex        =   99
               ToolTipText     =   "Show Custom Pure Wt? (Y/N)"
               Top             =   8205
               Visible         =   0   'False
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               Enabled         =   0   'False
               DataField       =   "UOYN1"
               IdName          =   "UOYN1"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   240
               Index           =   89
               Left            =   2400
               TabIndex        =   101
               ToolTipText     =   "Show Custom Pure Wt? (Y/N)"
               Top             =   6120
               Visible         =   0   'False
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   423
               Enabled         =   0   'False
               DataField       =   "UOYN6"
               IdName          =   "UOYN6"
            End
            Begin VB.Label Label6 
               BackStyle       =   0  'Transparent
               Caption         =   "Other Value Type"
               BeginProperty Font 
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
               Left            =   4680
               TabIndex        =   98
               ToolTipText     =   "Location"
               Top             =   5850
               Width           =   1845
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
               Left            =   60
               TabIndex        =   95
               Top             =   1290
               Width           =   1425
            End
            Begin VB.Label Label5 
               BackStyle       =   0  'Transparent
               Caption         =   "Show Desc"
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
               TabIndex        =   94
               Top             =   7635
               Visible         =   0   'False
               Width           =   1215
            End
            Begin VB.Label Label4 
               BackStyle       =   0  'Transparent
               Caption         =   "Last Column Value"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   570
               Left            =   60
               TabIndex        =   93
               Top             =   6780
               Width           =   1455
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Stone Class"
               BeginProperty Font 
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
               Left            =   60
               TabIndex        =   92
               Top             =   2145
               Width           =   1425
            End
            Begin VB.Label LblLotNo 
               BackStyle       =   0  'Transparent
               Caption         =   "Lot No."
               BeginProperty Font 
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
               TabIndex        =   91
               Top             =   3285
               Width           =   1425
            End
            Begin VB.Label Label3 
               BackStyle       =   0  'Transparent
               Caption         =   "Reverse Logic"
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
               Left            =   30
               TabIndex        =   90
               Top             =   7320
               Visible         =   0   'False
               Width           =   1485
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "SEO Pwd"
               BeginProperty Font 
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
               TabIndex        =   89
               Top             =   5565
               Width           =   1455
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "OpenYyyyMm"
               BeginProperty Font 
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
               Left            =   60
               TabIndex        =   88
               Top             =   4710
               Width           =   1425
            End
            Begin VB.Label Label2 
               BackStyle       =   0  'Transparent
               Caption         =   "Show Wt/Qty"
               BeginProperty Font 
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
               TabIndex        =   87
               Top             =   6495
               Width           =   1455
            End
            Begin VB.Label Lbl2ndLine 
               BackStyle       =   0  'Transparent
               Caption         =   "Value Type"
               BeginProperty Font 
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
               TabIndex        =   86
               ToolTipText     =   "Location"
               Top             =   5850
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
               Index           =   25
               Left            =   60
               TabIndex        =   85
               Top             =   705
               Width           =   1335
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Parent Cd"
               BeginProperty Font 
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
               Left            =   60
               TabIndex        =   84
               Top             =   2715
               Width           =   1425
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Range Size"
               BeginProperty Font 
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
               Left            =   60
               TabIndex        =   83
               Top             =   3570
               Width           =   1425
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Vch Option"
               BeginProperty Font 
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
               TabIndex        =   82
               Top             =   5280
               Width           =   1425
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Location Typ"
               BeginProperty Font 
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
               TabIndex        =   81
               Top             =   990
               Width           =   1425
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Size/Ptr"
               BeginProperty Font 
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
               TabIndex        =   80
               Top             =   3855
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
               Index           =   6
               Left            =   60
               TabIndex        =   79
               ToolTipText     =   "Location"
               Top             =   1860
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
               Index           =   7
               Left            =   60
               TabIndex        =   78
               Top             =   2430
               Width           =   1425
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Rm Code"
               BeginProperty Font 
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
               Left            =   60
               TabIndex        =   77
               Top             =   3000
               Width           =   1425
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Stock Rate"
               BeginProperty Font 
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
               TabIndex        =   76
               Top             =   4140
               Width           =   1425
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
               Left            =   60
               TabIndex        =   75
               Top             =   4425
               Width           =   1425
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Rm Customer"
               BeginProperty Font 
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
               TabIndex        =   74
               Top             =   4995
               Width           =   1425
            End
            Begin VB.Label LblLocCd 
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
               Left            =   60
               TabIndex        =   73
               Top             =   1575
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
               Left            =   6540
               TabIndex        =   72
               Top             =   120
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
               Left            =   4155
               TabIndex        =   71
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
               Left            =   1755
               TabIndex        =   70
               Top             =   120
               Width           =   1335
            End
         End
      End
   End
End
Attribute VB_Name = "EmrFrpStkLed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rep As New EmrRepStkLed
Dim moCn As MwfLib.MDOConnection '*** (Bef speed) bef 24/11/06
'*** Jay 2.14 *** (SEO Change)
Dim ms_PwdFlag As Boolean
'Enum en_StkLdgFra
'  Pwd = 0
'End Enum
'****** Sachin 2.14.0 - [25-11-2006] - [Multi-Company] ******
Dim ws_MultiCoMod As Boolean
'****** Sachin 2.14.0 - [25-11-2006] - [Multi-Company] ******

Private Sub SetGroupSort()
  GRP_REP.Add "Process", "LocCoCd+LocCd", "'['+LocCoCd+'] '+LocCd", "qLoc", "hqLoc", "Company Code,Parent Code", "", "", "LocDesc"
  GRP_REP.Add "Rm Ctg", "Rtrim(RmCtg)", "", "RmCtg", "hRmCtg", "", "", "", "(Select PDesc From Param Where PTyp= 'RMCTG' and PMCd = RmCtg) "
  GRP_REP.Add "Rm SubCtg", "Rtrim(RmCtg)+Rtrim(RmSCtg)", "'('+ RTrim(RmCtg)+ ') '+ RmSCtg", "RmCtg,RmSCtg", "hRmCtg,hRmSCtg", "Rm Ctg", "", "", "(Select PDesc From Param Where PTyp= 'RMSCTG' and PMCd = RmCtg and PSCd= RmSCtg)"
  GRP_REP.Add "Rm Code", "Rtrim(RmCtg)+Rtrim(RmSCtg)+Rtrim(RmCd)", "'('+ RTrim(RmCtg)+'/'+ RmSCtg+ ') '+ RmCd", "RmCtg,RmSCtg,RmCd", "hRmCtg,hRmSCtg,hRmCd", "Rm Ctg,Rm SubCtg,BaseCd + RmCd,Parent Cd", "", "", "SubString(RmDesc, 1, 30)"
  GRP_REP.Add "BaseCd + RmCd", "Rtrim(RmBaseCd)+Rtrim(RmCd)", "'('+ RTrim(RmBaseCd) + ') '+ RmCd", "RmCd", "hRmCd", "Rm Ctg,Rm SubCtg,Rm Code,Parent Cd", "", "", "SubString(RmDesc, 1, 30)"
  GRP_REP.Add "Parent Cd", "Rtrim(RmCtg)+Rtrim(RmSCtg)+(case when RmParentCd= '' then RmCd else RmParentCd end)", "'('+ RTrim(RmCtg)+'/'+ RmSCtg+ ') '+ (case when RmParentCd= '' then RmCd else RmParentCd end)", "RmCtg,RmSCtg", "hRmCtg,hRmSCtg", "Rm Ctg,Rm SubCtg,BaseCd + RmCd", "", "", "(Select PDesc from Param where PTyp= 'PARENTCD' and PMCd= RmParentCd) "
  
  '******Geeta*****************Emr206****06/04/04*******
  '**********If UserLevel>6 then don't display the Customer Name
  ' ***** Manali 3.8.0 - EmrFrmMDI.TXT_UCD.text Replaced by gs_UsrCd
  GRP_REP.Add "Customer Code", "RTrim(RmCmCd)", "", "", "", "", "", "", "(Select (case when (Select PNum from Param where Ptyp = 'USR' and PmCd = '" + Trim(gs_UsrCd) + "')  <= 5 then  CmName else CmCd end) From CustMst Where CmCtg= 'C' and CmCd= RmCmCd) "
  'GRP_REP.Add "Customer Code", "RTrim(RmCmCd)", "", "", "", "", "", "", "(Select (case when (Select PNum from Param where Ptyp = 'USR' and PmCd = '" + Trim(EmrFrmMDI.TXT_UCD.text) + "')  <= 5 then  CmName else CmCd end) From CustMst Where CmCtg= 'C' and CmCd= RmCmCd) "
  '****************Geeta**************Emr206*********12/04/04
  'GRP_REP.Add "Rm Sv/Sz", Array("str(SRmSz,6,3)", "str(TdRmSz,6,3)", "str(FrRmSz,6,3)", "str(IrRmSz,6,3)"), "", "qRmSz", "hqRmSz", "Rm Ptr,Rm StkRt,(RmCd) Sv/Sz,(RmCd) StkRt", "", "", ""
  Dim wsPtrYN As String
' ########################  Manoj #### Ver: 2.0.11 #### Date: 16/08/2005  ########################
'  wsPtrYN = "(Select PValue3 from Param where PTyp= 'RMSCTG' and PMCd= RmCtg and PSCd= RmSCtg)"
'  GRP_REP.Add "Rm Sv/Sz", Array("case when " + wsPtrYN + "='Y' then str(SRmSz,7,4) else str(SRmSz,6,3) end ", _
'  "case when " + wsPtrYN + "='Y' then  str(TdRmSz,7,4) else str(TdRmSz,6,3) end ", "case when " + wsPtrYN + "='Y' then  str(FrRmSz,7,4) else str(FrRmSz,6,3) end ", _
'  "case when " + wsPtrYN + "='Y' then   str(IrRmSz,7,4) else str(IrRmSz,6,3) end"), "", "qRmSz", "hqRmSz", "Rm Ptr,Rm StkRt,(RmCd) Sv/Sz,(RmCd) StkRt", "", "", ""
  wsPtrYN = "(Select PValue3 from Param where PTyp= 'RMSCTG' and PMCd= RmCtg and PSCd= RmSCtg)"
'  GRP_REP.Add "Rm Sv/Sz", Array("case when " + wsPtrYN + "='Y' then str(SRmSz,7,4) else str(SRmSz,6,3) end ", _
'  "case when " + wsPtrYN + "='Y' then  str(TdRmSz,7,4) else str(TdRmSz,6,3) end ", "case when " + wsPtrYN + "='Y' then  str(FrRmSz,7,4) else str(FrRmSz,6,3) end ", _
'  "case when " + wsPtrYN + "='Y' then   str(IrRmSz,7,4) else str(IrRmSz,6,3) end"), "", "qRmSz", "hqRmSz", "Rm Ptr,(RmCd) Sv/Sz,Range Size", "", "", ""
   ' *** Manali 3.03 - RmSzDesc added

'  GRP_REP.Add "Rm Sv/Sz", Array("case when " + wsPtrYN + "='Y' then str(SRmSz,7,4) else IsNull(RisRmSzDesc, str(SRmSz,6,3)) end ", _
  "case when " + wsPtrYN + "='Y' then  str(TdRmSz,7,4) else IsNull(RisRmSzDesc, str(TdRmSz,6,3)) end ", "case when " + wsPtrYN + "='Y' then  str(FrRmSz,7,4) else IsNull(RisRmSzDesc, str(FrRmSz,6,3)) end ", _
  "case when " + wsPtrYN + "='Y' then   str(IrRmSz,7,4) else IsNull(RisRmSzDesc, str(IrRmSz,6,3)) end"), "", "wRmSz", "hqRmSz", "Rm Ptr,(RmCd) Sv/Sz,Range Size", "", "", ""
  
  GRP_REP.Add "Rm Sv/Sz", Array("case when " + wsPtrYN + "='Y' then " + GetRmSzDBFldNm("SRmSz", "SRmSz2", "SRmSz3") + " else (case when RisRmSzDesc is null then  " + GetRmSzDBFldNm("SRmSz", "SRmSz2", "SRmSz3") + " else RisRmSzDesc end) end ", _
                                "case when " + wsPtrYN + "='Y' then " + GetRmSzDBFldNm("TdRmSz", "TdRmSz2", "TdRmSz3") + " else (case when RisRmSzDesc is null then " + GetRmSzDBFldNm("TdRmSz", "TdRmSz2", "TdRmSz3") + " else RisRmSzDesc end) end ", _
                                "case when " + wsPtrYN + "='Y' then " + GetRmSzDBFldNm("FrRmSz", "FrRmSz2", "FrRmSz3") + " else (case when RisRmSzDesc is null then " + GetRmSzDBFldNm("FrRmSz", "FrRmSz2", "FrRmSz3") + " else RisRmSzDesc end) end ", _
                                "case when " + wsPtrYN + "='Y' then " + GetRmSzDBFldNm("IrRmSz", "IrRmSz2", "IrRmSz3") + " else (case when RisRmSzDesc is null then " + GetRmSzDBFldNm("IrRmSz", "IrRmSz2", "IrRmSz3") + " else RisRmSzDesc end) end"), _
                                "", "wRmSz", "hRmSz", "Rm Ptr,(RmCd) Sv/Sz,Range Size", "", "", ""
  
' GRP_REP.Add "Rm StkRt", Array("str(SRmStkRt,9,2)", "str(TdRmStkRt,9,2)", "str(FrRmStkRt,9,2)", "str(IrStkRt,9,2)"), "", "qStkRt", "hqStkRt", "Rm Ptr,Rm Sv/Sz,(RmCd) Sv/Sz,(RmCd) StkRt", "", "", ""
  
  GRP_REP.Add "Rm StkRt", Array("str(SRmStkRt,9,2)", "str(TdRmStkRt,9,2)", "str(FrRmStkRt,9,2)", "str(IrStkRt,9,2)"), "", "qStkRt", "hqStkRt", "(RmCd) StkRt", "", "", ""
  
  '4.1.4
  ' ****************************************************************************************************************************
  Dim wSRate As String, wTdRate As String, wFrRate As String, wIrRate As String
  
    wSRate = " (Case When RSCtg.PValue1='N' Then Str(" + RmRate("RrCstRt", "SRmSz", "RmQwCst") + ",9,2) Else Str(SRmStkRt,9,2) End)"
    wTdRate = " (Case When RSCtg.PValue1='N' Then Str(" + RmRate("RrCstRt", "TdRmSz", "RmQwCst") + ",9,2) Else Str(TdRmStkRt,9,2) End)"
    wFrRate = " (Case When RSCtg.PValue1='N' Then Str(" + RmRate("RrCstRt", "FrRmSz", "RmQwCst") + ",9,2) Else Str(FrRmStkRt,9,2) End)"
    wIrRate = " (Case When RSCtg.PValue1='N' Then Str(" + RmRate("RrCstRt", "IrRmSz", "RmQwCst") + ",9,2) Else Str(IrStkRt,9,2) End)"
  
  GRP_REP.Add "(RmCd) Stk-CstRt", Array(wSRate, wTdRate, wFrRate, wIrRate), "", "qStkRt", "hqStkRt", "(RmCd) StkRt", "", "", ""
  
  
    wSRate = " (Case When RSCtg.PValue1='N' Then Str(" + RmRate("RrSalRt", "SRmSz", "RmQw") + ",9,2) Else Str(SRmStkRt,9,2) End)"
    wTdRate = " (Case When RSCtg.PValue1='N' Then Str(" + RmRate("RrSalRt", "TdRmSz", "RmQw") + ",9,2) Else Str(TdRmStkRt,9,2) End)"
    wFrRate = " (Case When RSCtg.PValue1='N' Then Str(" + RmRate("RrSalRt", "FrRmSz", "RmQw") + ",9,2) Else Str(FrRmStkRt,9,2) End)"
    wIrRate = " (Case When RSCtg.PValue1='N' Then Str(" + RmRate("RrSalRt", "IrRmSz", "RmQw") + ",9,2) Else Str(IrStkRt,9,2) End)"
  
  GRP_REP.Add "(RmCd) Stk-SalRt", Array(wSRate, wTdRate, wFrRate, wIrRate), "", "qStkRt", "hqStkRt", "(RmCd) StkRt", "", "", ""
  ' ****************************************************************************************************************************
  
  
  'The check below is no more required as lot no will always be displayed for this report if multi-comp module is on
  If (moCn.GetFldVal("Select HPtrYN from Head where HCoCd = '" + gs_CoCd + "' and " + _
                                                  "HCd='" + ctSelfCmCd + "'") = "Y") Or (ws_MultiCoMod = True) Then
     'GRP_REP.Add "Rm Ptr", Array("str(SRmSz,6,3)", "str(TdRmSz,6,3)", "str(FrRmSz,6,3)", "str(IrRmSz,6,3)"), "", "qRmSz", "hqRmSz", "Rm StkRt,(RmCd) Sv/Sz,(RmCd) StkRt", "", "", ""
     '****************Geeta**************Emr206*********12/04/04
'     GRP_REP.Add "Rm Ptr", Array("case when " + wsPtrYN + "='Y' then str(SRmSz,7,4) else str(SRmSz,6,3) end ", _
'      "case when " + wsPtrYN + "='Y' then  str(TdRmSz,7,4) else str(TdRmSz,6,3) end ", _
'      "case when " + wsPtrYN + "='Y' then  str(FrRmSz,7,4) else str(FrRmSz,6,3) end ", _
'      "case when " + wsPtrYN + "='Y' then   str(IrRmSz,7,4) else str(IrRmSz,6,3) end"), "", "qRmSz", "hqRmSz", "Rm StkRt,(RmCd) Sv/Sz,(RmCd) StkRt", "", "", ""
'     GRP_REP.Add "Rm Ptr", Array("case when " + wsPtrYN + "='Y' then str(SRmSz,7,4) else str(SRmSz,6,3) end ", _
'      "case when " + wsPtrYN + "='Y' then  str(TdRmSz,7,4) else str(TdRmSz,6,3) end ", _
'      "case when " + wsPtrYN + "='Y' then  str(FrRmSz,7,4) else str(FrRmSz,6,3) end ", _
'      "case when " + wsPtrYN + "='Y' then   str(IrRmSz,7,4) else str(IrRmSz,6,3) end"), "", "qRmSz", "hqRmSz", "(RmCd) Sv/Sz,Range Size", "", "", ""
    '*****************

      GRP_REP.Add "Rm Ptr", Array("case when " + wsPtrYN + "='Y' then str(SRmSz,7,4) else IsNull(RisRmSzDesc, str(SRmSz,6,3)) end ", _
      "case when " + wsPtrYN + "='Y' then  str(TdRmSz,7,4) else IsNull(RisRmSzDesc, str(TdRmSz,6,3)) end ", _
      "case when " + wsPtrYN + "='Y' then  str(FrRmSz,7,4) else IsNull(RisRmSzDesc, str(FrRmSz,6,3)) end ", _
      "case when " + wsPtrYN + "='Y' then   str(IrRmSz,7,4) else IsNull(RisRmSzDesc, str(IrRmSz,6,3)) end"), "", "wRmSz", "hRmSz", "(RmCd) Sv/Len,Range Size", "", "", ""

'      GRP_REP.Add "Rm Ptr/Sz", Array("case when " + wsPtrYN + "='Y' then " + GetRmSzDBFldNm("SRmSz", "SRmSz2", "SRmSz3") + " else IsNull(RisRmSzDesc, " + GetRmSzDBFldNm("SRmSz", "SRmSz2", "SRmSz3") + ") end ", _
                                "case when " + wsPtrYN + "='Y' then " + GetRmSzDBFldNm("TdRmSz", "TdRmSz2", "TdRmSz3") + " else IsNull(RisRmSzDesc, " + GetRmSzDBFldNm("TdRmSz", "TdRmSz2", "TdRmSz3") + ") end ", _
                                "case when " + wsPtrYN + "='Y' then " + GetRmSzDBFldNm("FrRmSz", "FrRmSz2", "FrRmSz3") + " else IsNull(RisRmSzDesc, " + GetRmSzDBFldNm("FrRmSz", "FrRmSz2", "FrRmSz3") + ") end ", _
                                "case when " + wsPtrYN + "='Y' then " + GetRmSzDBFldNm("IrRmSz", "IrRmSz2", "IrRmSz3") + " else IsNull(RisRmSzDesc, " + GetRmSzDBFldNm("IrRmSz", "IrRmSz2", "IrRmSz3") + ") end"), _
                                "", "wRmSz", "hqRmSz", "Rm Ptr,(RmCd) Sv/Sz,Range Size", "", "", ""
  End If
  
  GRP_REP.Add "Range Size", "RmCtg+RmSCtg+IsNull(RrRngSz, '')", "'('+ RTrim(RmCtg)+ ', '+ RmSCtg+ ') '+IsNull(RrRngSz, '')", "RmCtg,RmSCtg", "hRmCtg,hRmSCtg", "Rm Ctg,Rm SubCtg", "", "", ""
 ' GRP_REP.Add "(RmCd) Sv/Sz", Array("RmCd+str(SRmSz,6,3)", "RmCd+str(TdRmSz,6,3)", "RmCd+str(FrRmSz,6,3)", "RmCd+str(IrRmSz,6,3)"), Array("'('+RmCd+') '+LTrim(str(SRmSz,6,3))", "'('+RmCd+') '+LTrim(str(TdRmSz,6,3))", "'('+RmCd+') '+LTrim(str(FrRmSz,6,3))", "'('+RmCd+') '+LTrim(str(IrRmSz,6,3))"), "RmCtg,RmSCtg,RmCd,qRmSz", "hRmCtg,hRmSCtg,hRmCd,hqRmSz", "Rm Ctg,Rm SubCtg,Rm Code,BaseCd + RmCd,(RmCd) StkRt,Rm Sv/Sz,Rm Ptr,Rm StkRt", "", "", " '('+ RTrim(RmCtg)+ '/'+ RmSCtg+ ')' "
  '****************Geeta**************Emr206*********12/04/04
'   GRP_REP.Add "(RmCd) Sv/Sz", Array("RmCd+ case when " + wsPtrYN + "='Y' then str(SRmSz,7,4) else str(SRmSz,6,3) end ", _
'     "RmCd+ case when " + wsPtrYN + "='Y' then  str(TdRmSz,7,4) else str(TdRmSz,6,3) end ", _
'     "RmCd+ case when " + wsPtrYN + "='Y' then  str(FrRmSz,7,4) else str(FrRmSz,6,3) end ", _
'     "RmCd+ case when " + wsPtrYN + "='Y' then   str(IrRmSz,7,4) else str(IrRmSz,6,3) end "), _
'     Array(" '('+RmCd+') '+LTrim(case when " + wsPtrYN + "='Y' then str(SRmSz,7,4) else str(SRmSz,6,3) end )", _
'     " '('+RmCd+') '+LTrim(case when " + wsPtrYN + "='Y' then str(TdRmSz,7,4) else str(TdRmSz,6,3) end )", _
'     " '('+RmCd+') '+LTrim(case when " + wsPtrYN + "='Y' then  str(FrRmSz,7,4) else  str(FrRmSz,6,3) end ) ", _
'     " '('+RmCd+') '+LTrim(case when " + wsPtrYN + "='Y' then  str(IrRmSz,7,4) else  str(IrRmSz,6,3) end )"), _
'     "RmCtg,RmSCtg,RmCd,qRmSz", "hRmCtg,hRmSCtg,hRmCd,hqRmSz", "Rm Ctg,Rm SubCtg,Rm Code,BaseCd + RmCd,(RmCd) StkRt,Rm Sv/Sz,Rm Ptr,Rm StkRt", "", "", " '('+ RTrim(RmCtg)+ '/'+ RmSCtg+ ')' "
'   GRP_REP.Add "(RmCd) Sv/Sz", Array("RmCd+ case when " + wsPtrYN + "='Y' then str(SRmSz,7,4) else str(SRmSz,6,3) end ", _
'     "RmCd+ case when " + wsPtrYN + "='Y' then  str(TdRmSz,7,4) else str(TdRmSz,6,3) end ", _
'     "RmCd+ case when " + wsPtrYN + "='Y' then  str(FrRmSz,7,4) else str(FrRmSz,6,3) end ", _
'     "RmCd+ case when " + wsPtrYN + "='Y' then   str(IrRmSz,7,4) else str(IrRmSz,6,3) end "), _
'     Array(" '('+RmCd+') '+LTrim(case when " + wsPtrYN + "='Y' then str(SRmSz,7,4) else str(SRmSz,6,3) end )", _
'     " '('+RmCd+') '+LTrim(case when " + wsPtrYN + "='Y' then str(TdRmSz,7,4) else str(TdRmSz,6,3) end )", _
'     " '('+RmCd+') '+LTrim(case when " + wsPtrYN + "='Y' then  str(FrRmSz,7,4) else  str(FrRmSz,6,3) end ) ", _
'     " '('+RmCd+') '+LTrim(case when " + wsPtrYN + "='Y' then  str(IrRmSz,7,4) else  str(IrRmSz,6,3) end )"), _
'     "RmCtg,RmSCtg,RmCd,qRmSz", "hRmCtg,hRmSCtg,hRmCd,hqRmSz", "Rm Ctg,Rm SubCtg,Rm Code,BaseCd + RmCd,Rm Sv/Sz,Rm Ptr,(RmCd) StkRt,Range Size", "", "", " '('+ RTrim(RmCtg)+ '/'+ RmSCtg+ ')' "
   ' ***** Manali 3.03 - RmSzDesc added
     
'   GRP_REP.Add "(RmCd) Sv/Sz", Array("RmCd + case when " + wsPtrYN + "='Y' then " + GetRmSzDBFldNm("SRmSz", "SRmSz2", "SRmSz3") + " else (case when RisRmSzDesc is null then " + GetRmSzDBFldNm("SRmSz", "SRmSz2", "SRmSz3") + " else RisRmSzDesc end) end ", _
                                "RmCd + case when " + wsPtrYN + "='Y' then " + GetRmSzDBFldNm("TdRmSz", "TdRmSz2", "TdRmSz3") + " else (case when RisRmSzDesc is null then " + GetRmSzDBFldNm("TdRmSz", "TdRmSz2", "TdRmSz3") + " else RisRmSzDesc end) end ", _
                                "RmCd + case when " + wsPtrYN + "='Y' then " + GetRmSzDBFldNm("FrRmSz", "FrRmSz2", "FrRmSz3") + " else (case when RisRmSzDesc is null then " + GetRmSzDBFldNm("FrRmSz", "FrRmSz2", "FrRmSz3") + " else RisRmSzDesc end) end ", _
                                "RmCd + case when " + wsPtrYN + "='Y' then " + GetRmSzDBFldNm("IrRmSz", "IrRmSz2", "IrRmSz3") + " else (case when RisRmSzDesc is null then " + GetRmSzDBFldNm("IrRmSz", "IrRmSz2", "IrRmSz3") + " else RisRmSzDesc end) end"), _
                               Array(" '('+RmCd+') '+LTrim(case when " + wsPtrYN + "='Y' then " + GetRmSzDispFldNm("SRmSz", "SRmSz2", "SRmSz3") + " else (case when RisRmSzDesc is null then " + GetRmSzDispFldNm("SRmSz", "SRmSz2", "SRmSz3") + " else RisRmSzDesc end) end )", _
                                " '('+RmCd+') '+LTrim(case when " + wsPtrYN + "='Y' then " + GetRmSzDispFldNm("TdRmSz", "TdRmSz2", "TdRmSz3") + " else (case when RisRmSzDesc is null then " + GetRmSzDispFldNm("TdRmSz", "TdRmSz2", "TdRmSz3") + " else RisRmSzDesc end) end )", _
                                " '('+RmCd+') '+LTrim(case when " + wsPtrYN + "='Y' then " + GetRmSzDispFldNm("FrRmSz", "FrRmSz2", "FrRmSz3") + " else (case when RisRmSzDesc is null then " + GetRmSzDispFldNm("FrRmSz", "FrRmSz2", "FrRmSz3") + " else RisRmSzDesc end) end )", _
                                " '('+RmCd+') '+LTrim(case when " + wsPtrYN + "='Y' then " + GetRmSzDispFldNm("IrRmSz", "IrRmSz2", "IrRmSz3") + " else (case when RisRmSzDesc is null then " + GetRmSzDispFldNm("IrRmSz", "IrRmSz2", "IrRmSz3") + " else RisRmSzDesc end) end )"), _
                                "RmCtg,RmSCtg,RmCd,wRmSz", "hRmCtg,hRmSCtg,hRmCd,hRmSz", "Rm Ctg,Rm SubCtg,Rm Code,BaseCd + RmCd,Rm Sv/Sz,Rm Ptr,(RmCd) StkRt,Range Size", "", "", " '('+ RTrim(RmCtg)+ '/'+ RmSCtg+ ')' "
                                
   GRP_REP.Add "(RmCd) Sv/Sz", Array("RmCd + case when " + wsPtrYN + "='Y' then " + GetRmSzDBFldNm("SRmSz", "SRmSz2", "SRmSz3") + " else (case when RisRmSzDesc is null then " + GetRmSzDBFldNm("SRmSz", "SRmSz2", "SRmSz3") + " else RisRmSzDesc end) end ", _
                                "RmCd + case when " + wsPtrYN + "='Y' then " + GetRmSzDBFldNm("TdRmSz", "TdRmSz2", "TdRmSz3") + " else (case when RisRmSzDesc is null then " + GetRmSzDBFldNm("TdRmSz", "TdRmSz2", "TdRmSz3") + " else RisRmSzDesc end) end ", _
                                "RmCd + case when " + wsPtrYN + "='Y' then " + GetRmSzDBFldNm("FrRmSz", "FrRmSz2", "FrRmSz3") + " else (case when RisRmSzDesc is null then " + GetRmSzDBFldNm("FrRmSz", "FrRmSz2", "FrRmSz3") + " else RisRmSzDesc end) end ", _
                                "RmCd + case when " + wsPtrYN + "='Y' then " + GetRmSzDBFldNm("IrRmSz", "IrRmSz2", "IrRmSz3") + " else (case when RisRmSzDesc is null then " + GetRmSzDBFldNm("IrRmSz", "IrRmSz2", "IrRmSz3") + " else RisRmSzDesc end) end"), _
                               Array(" '('+RmCd+') '+LTrim(case when " + wsPtrYN + "='Y' then " + GetRmSzDispFldNm("SRmSz", "SRmSz2", "SRmSz3") + " else (case when RisRmSzDesc is null then " + GetRmSzDispFldNm("SRmSz", "SRmSz2", "SRmSz3") + " else RisRmSzDesc end) end )", _
                                " '('+RmCd+') '+LTrim(case when " + wsPtrYN + "='Y' then " + GetRmSzDispFldNm("TdRmSz", "TdRmSz2", "TdRmSz3") + " else (case when RisRmSzDesc is null then " + GetRmSzDispFldNm("TdRmSz", "TdRmSz2", "TdRmSz3") + " else RisRmSzDesc end) end )", _
                                " '('+RmCd+') '+LTrim(case when " + wsPtrYN + "='Y' then " + GetRmSzDispFldNm("FrRmSz", "FrRmSz2", "FrRmSz3") + " else (case when RisRmSzDesc is null then " + GetRmSzDispFldNm("FrRmSz", "FrRmSz2", "FrRmSz3") + " else RisRmSzDesc end) end )", _
                                " '('+RmCd+') '+LTrim(case when " + wsPtrYN + "='Y' then " + GetRmSzDispFldNm("IrRmSz", "IrRmSz2", "IrRmSz3") + " else (case when RisRmSzDesc is null then " + GetRmSzDispFldNm("IrRmSz", "IrRmSz2", "IrRmSz3") + " else RisRmSzDesc end) end )"), _
                                "RmCtg,RmSCtg,RmCd,wRmSz", "hRmCtg,hRmSCtg,hRmCd,hRmSz", "Rm Ctg,Rm SubCtg,Rm Code,BaseCd + RmCd,Rm Sv/Sz,Rm Ptr,(RmCd) StkRt,Range Size", "", "", " '('+ RTrim(RmCtg)+ '/'+ RmSCtg+ ')' "

  '*****************
'  GRP_REP.Add "(RmCd) StkRt", Array("RmCd+str(SRmStkRt,9,2)", "RmCd+str(TdRmStkRt,9,2)", "RmCd+str(FrRmStkRt,9,2)", "RmCd+str(IrStkRt,9,2)"), Array("'('+RmCd+') '+LTrim(str(SRmStkRt,9,2))", "'('+RmCd+') '+LTrim(str(TdRmStkRt,9,2))", "'('+RmCd+') '+LTrim(str(FrRmStkRt,9,2))", "'('+RmCd+') '+LTrim(str(IrStkRt,9,2))"), "RmCtg,RmSCtg,RmCd,qStkRt", "hRmCtg,hRmSCtg,hRmCd,hqStkRt", "Rm Ctg,Rm SubCtg,Rm Code,BaseCd + RmCd,(RmCd) Sv/Sz,Rm Sv/Sz,Rm Ptr,Rm StkRt", "", "", " '('+ RTrim(RmCtg)+ '/'+ RmSCtg+ ')' "
  GRP_REP.Add "(RmCd) StkRt", Array("RmCd+str(SRmStkRt,9,2)", "RmCd+str(TdRmStkRt,9,2)", "RmCd+str(FrRmStkRt,9,2)", "RmCd+str(IrStkRt,9,2)"), Array("'('+RmCd+') '+LTrim(str(SRmStkRt,9,2))", "'('+RmCd+') '+LTrim(str(TdRmStkRt,9,2))", "'('+RmCd+') '+LTrim(str(FrRmStkRt,9,2))", "'('+RmCd+') '+LTrim(str(IrStkRt,9,2))"), "RmCtg,RmSCtg,RmCd,qStkRt", "hRmCtg,hRmSCtg,hRmCd,hqStkRt", "Rm Ctg,Rm SubCtg,Rm Code,BaseCd + RmCd,(RmCd) Sv/Sz,Rm StkRt", "", "", " '('+ RTrim(RmCtg)+ '/'+ RmSCtg+ ')' "
' ###########################################  Manoj  ###########################################
  
  '****** Sachin 2.13.0 - [07-07-2006] ******
  GRP_REP.Add "Stone Class", "RmCtg+IsNull(Sc.PSCd, '')", "'('+ RTrim(RmCtg)+') '+IsNull(sc.PSCd, '')", "RmCtg", "hRmCtg", "Rm Ctg", "", "", "sc.PDesc"
  '****** Sachin 2.13.0 - [07-07-2006] ******

  'GRP_REP.Add "(RmCd) Sv/Sz", Array("RmCd+str(SRmSz,4,2)", "RmCd+str(TdRmSz,4,2)", "RmCd+str(FrRmSz,4,2)"), Array("'('+RmCd+') '+LTrim(str(SRmSz,4,2))", "'('+RmCd+') '+LTrim(str(TdRmSz,4,2))", "'('+RmCd+') '+LTrim(str(FrRmSz,4,2))"), "RmCtg,RmSCtg,RmCd,qRmSz", "hRmCtg,hRmSCtg,hRmCd,hqRmSz", "Rm Ctg,Rm SubCtg,Rm Code,(RmCd) StkRt", "", "", " '('+ RTrim(RmCtg)+ '/'+ RmSCtg+ ')' "
  'GRP_REP.Add "(RmCd) StkRt", Array("RmCd+str(SRmStkRt,9,2)", "RmCd+str(TdRmStkRt,9,2)", "RmCd+str(FrRmStkRt,9,2)"), Array("'('+RmCd+') '+LTrim(str(SRmStkRt,9,2))", "'('+RmCd+') '+LTrim(str(TdRmStkRt,9,2))", "'('+RmCd+') '+LTrim(str(FrRmStkRt,9,2))"), "RmCtg,RmSCtg,RmCd,qStkRt", "hRmCtg,hRmSCtg,hRmCd,hqStkRt", "Rm Ctg,Rm SubCtg,Rm Code,(RmCd) Sv/Sz", "", "", " '('+ RTrim(RmCtg)+ '/'+ RmSCtg+ ')' "
  
  ' Zubin 212
  'The check below is no more required as lot no will always be displayed for this report
  'If moCn.GetFldVal("select Max(HLotNoYn) from Head where HCd= '" + ctSelfCmCd + "' " + IIF(adc("UoCoCdFr") <> "", " and HCoCd>= '" + adc("UoCoCdFr") + "' ", " ") + IIF(adc("UoCoCdTo") <> "", " and HCoCd<= '" + adc("UoCoCdTo") + "' ", " ") + IIF(adc("UoCoCdSel") <> "", " and HCoCd In (" + adc("UoCoCdSel") + ") ", " ")) = "Y" Then
  If (moCn.RecSeek("select 'x' from Head where HCd= '" + ctSelfCmCd + "' and HCoCd= '" + gs_CoCd + "' and HLotNoYN= 'Y' ") And ws_MultiCoMod = False) Or (ws_MultiCoMod = True) Then
    GRP_REP.Add "RM Lot No.", Array("SLotNo", "TdLotNo", "FrLotNo", "IrLotNo"), "", "qLotNo", "hLotNo", "", "", "", Array("SLotNo", "TdLotNo", "FrLotNo", "IrLotNo")
  End If
  ' Zubin 212
  
  '****** Sachin 2.14.0 - [23-11-2006] - [Multi-Company] - ,Process,OrderNo,OrderSr,Customer PO,FgSubLoc ******
  If ws_MultiCoMod = True Then
    GRP_REP.Add "Company Code", "LocCoCd", "LocCoCd", "", "", "", "", "", "(Select HName from Head where HCd= '" + ctSelfCmCd + "' and HCoCd=LocCoCd) "
    GRP_REP.Add "Parent Location", "LocPrntCd", "LocPrntCd", "", "", "", "", "", "(Select LocDesc From Loc Where LocCoCd= '" + ctSelfCoCd + "' And LocCd=Lc.LocPrntCd) "
  End If
  '****** Sachin 2.14.0 - [23-11-2006] - [Multi-Company] - ******
  
End Sub

Private Sub adc_setreprecsource()
'On Error GoTo errhld
Dim ws_Val As String

  '*** Report Sql ***
  '*** (Jen 2.12)
  '  Select space(30) as Grp1, space(30) as Grp2, space(30) as Grp3,
  '  space(30) as DGrp1, space(30) as DGrp2, space(30) as DGrp3,
  '  space(30) as G1Desc, space(30) as G2Desc, space(30) as G3Desc,
  '  space(1) as qFlag, RmCd, TdLotNo as qLotNo, TdRmSz as qRmSz, RrRngSz, RmCtg, RmSCtg,
  '  TdRmStkRt as qStkRt, PValue3 as qPtrYN, TdRmWt as qPureWt, TdCoCd, TdTc, TdYy,
  '  TdChr, TdNo, TdSr, TdSrNo, TdBYy, TdBChr, TdBNo, TDt as qDt,
  '  LocPrntCd, TdToRmLoc as qLoc, 0 as qOpnQty, TdRmQty as qDrQty,
  '  TdRmQty as qCrQty, 0.0 as qOpnWt, TdRmWt as qDrWt, TdRmWt as qCrWt,
  '  TdFrRmLoc As qFrRmLoc, TdToRmLoc As qToRmLoc, TdRmStkRt As qRmRt,
  '  space(1) As qQW, space(1) As qQWCst, TdRmStkRt as qVal
  '  From Txnd, Txn, RmMst, Loc, RmRt, Param where 1=2
  '*** (Jen 2.12)

  '*** (Bef 2.12)
  '  Select space(30) as Grp1, space(30) as Grp2, space(30) as Grp3,
  '  space(30) as DGrp1, space(30) as DGrp2, space(30) as DGrp3,
  '  space(30) as G1Desc, space(30) as G2Desc, space(30) as G3Desc,
  '  space(1) as qFlag, RmCd, TdLotNo as qLotNo, TdRmSz as qRmSz, RrRngSz, RmCtg, RmSCtg,
  '  TdRmStkRt as qStkRt, PValue3 as qPtrYN, RmPurityWt as qRmPurityWt, TdCoCd, TdTc, TdYy,
  '  TdChr, TdNo, TdSr, TdSrNo, TdBYy, TdBChr, TdBNo, TDt as qDt,
  '  LocPrntCd, TdToRmLoc as qLoc, 0 as qOpnQty, TdRmQty as qDrQty,
  '  TdRmQty as qCrQty, 0.0 as qOpnWt, TdRmWt as qDrWt, TdRmWt as qCrWt,
  '  TdFrRmLoc As qFrRmLoc, TdToRmLoc As qToRmLoc, TdRmStkRt As qRmRt,
  '  space(1) As qQW
  '  From Txnd, Txn, RmMst, Loc, RmRt, Param where 1=2
  '*** (Bef 2.12)
  '*** Report Sql ***
  
  '### 4. Ledger will be for a RmCtg, for a LocTyp, user may take 'W' also. 'FG' not permitted.
  
  Dim grpflds0 As String, grpflds1 As String, grpflds2 As String, grpflds3 As String, wrepcnd As String
  Dim wCnd As String, wSqlStrg As String, wPurityWt As String, wDtCndOpnYyMm As String
  Dim wsqlOpnYyMm As String, wFrDt As String, wDtCndOpnTd As String, wSqlOpnTdFr As String
  Dim wSqlOpnTdTo As String, wDtCndDet As String, wSqlDetFr As String, wSqlDetTo As String
  Dim wSqlOpnFgFr As String, wSqlOpnFgTo As String, wSqlDetFgFr As String, wSqlDetFgTo As String
  Dim wFrDt1 As String, wDtCndOpnFg As String, wDtCndDetFg As String, wFgStr As String
  Dim wTdRate As String, wFrRate, wSyRate As String
  Dim wQw As String, wOpnYyStr As String
  Dim wTmpDt As String
  Dim wPureWtOrZWt As String, wXPureWtOrZwt As String
  
  '*****sachin 25-02-05
  Dim wSetTmpTab As String, wQwChk As String
  Dim wb_ShowOpn, wb_MonEnd, wb_MonStart As Boolean
  
  '*********** urmi 26/8/03 ************
  Dim wRLQty As String, wRLWt As String
  '*************************************
  '********************** urmi 5/7/03 ch28********************************************'
  Dim wDtCndOpnInv As String, wDtCndDetInv As String
  Dim wSqlOpnInvFr As String, wSqlOpnInvTo As String, wSqlDetInvFr As String, wSqlDetInvTo As String
  Dim wInvStr As String, wIrRate As String
  '************************************************************************************'
  Dim wHidLocStr As String
  '*** Jay 2.14 ***
  Dim qDesc As String
  '*** Jay 2.14 ***
  
  '*** Manali 3.03 - 07/07/08 - RmSzDesc
  Dim wRmSzDesc As String
  '*** Manali 3.03 - 07/07/08 - RmSzDesc
    
  Dim wAddlCnd As String   '****** Sachin 3.02.0
  
  Dim wDtFr As String, wDtTo As String  ' ***** Manali 3.4.1 - 02/12/08
  
  ' ***** Manali Trading Module
  Dim wBagJoin As String
  Dim wSqlOpnXdFr As String, wSqlOpnXdTo As String, wSqlDetXdFr As String, wSqlDetXdTo As String
  Dim wDtCndOpnXd As String, wDtCndDetXd As String
  ' ***** Manali Trading Module
  
  '****** Sachin 3.02.0 - Check on YYYYMM Entered by user ******]
  ' **** Manali 3.4.1 - 02/12/08 - "01/01/80" excluded from check
  Dim ws_PurgeDt As String
  'sv.9 getting max(PurgeDate) from vParam for given company code
  ws_PurgeDt = moCn.GetFldVal("Select max(convert(date,vPDesc225,3)) from vParam " + _
                          " where vPCoCd >='" + ADC("UoCoCdFr") + "' and " + _
                          " vPCoCd <='" + IIF(ADC("UoCoCdTo") <> "", ADC("UoCoCdTo"), "ZZZ") + "'" + _
                          IIF(ADC("UoCoCdSel") <> "", " and vPCoCd In(" + ADC("UoCoCdSel") + ")", "") + _
                          " and vPTyp= 'PURGEDT' and vPMCd= 'PURGEDT' and vPSCd =''")
  
  If ws_PurgeDt = "" Then ws_PurgeDt = "01/01/80"
  
  'sv.9 gs_PurgeDt replaced with ws_PurgeDt
  If gs_Partition = ctCurrPrtn And CDate(ws_PurgeDt) <> "01/01/80" And _
      ((CDate(ADC("UOOPNYYMM")) <= Format(CDate(ws_PurgeDt), "YYYYMM") And ADC("UOOPNYYMM") <> 0) Or _
        (CDate(ADC("UODTFR")) <= CDate(ws_PurgeDt) And CDate(ADC("UoDtFr")) <> MWLib.EmptyDate) Or _
        (CDate(ADC("UODTTO")) <= CDate(ws_PurgeDt) And CDate(ADC("UoDtTo")) <> MWLib.EmptyDate)) Then
      DispMsg "From/To Date Or OpenYYYYMM Cannot Be Less Than Purge Date [" + CStr(CDate(ws_PurgeDt)) + "]", etError: Exit Sub
  End If
  '****** Sachin 3.02.0 - Check on YYYYMM Entered by user ******
  
  Set Rep = Nothing
  Set ADC.RepSource = Rep
  Call DispCoNm
  Call SetMwName(Rep)   ' ***** Manali 3.6.1 - Microway footer
  
  Call GRP_REP.Gen3LRep
  grpflds0 = GRP_REP.GrpFldLst(0)
  grpflds1 = GRP_REP.GrpFldLst(1)
  grpflds2 = GRP_REP.GrpFldLst(2)
  grpflds3 = GRP_REP.GrpFldLst(3)
  
  '*** Jay 2.14 ***
  '*** To check if Detail is selected ***
  Dim wGrp(3) As String, i As Integer, wDetPos As Integer
  wGrp(1) = GRP_REP.Value(gltGroup1)
  wGrp(2) = GRP_REP.Value(gltGroup2)
  wGrp(3) = GRP_REP.Value(gltGroup3)
  wDetPos = 0
  For i = 1 To 3
    If wGrp(i) = "(Detail)" Then
      wDetPos = i: Exit For
    End If
  Next i
  
  '*** If Detail not selected then suppress Section 'DetDesc' Of Detail ***
  If wDetPos = 0 Then Rep.DETDesc.Suppress = True: Rep.DetSecLine.Suppress = True
  '*** Jay 2.14 ***
  
  ' ****** Manali 3.4.1 - "01/01/80" converted to purge dt
  'sv.9 gs_PurgeDt replaced with ws_PurgeDt
  If gs_Partition = ctCurrPrtn And CDate(ws_PurgeDt) <> "01/01/80" And CDate(ADC("UoDtFr")) = MWLib.EmptyDate Then
    wDtFr = CDate(ws_PurgeDt)
  Else
    wDtFr = ADC("UoDtFr")
  End If
  'sv.9 gs_PurgeDt replaced with ws_PurgeDt
  If gs_Partition = ctCurrPrtn And CDate(ws_PurgeDt) <> "01/01/80" And CDate(ADC("UoDtTo")) = MWLib.EmptyDate Then
    wDtTo = CDate(ws_PurgeDt)
  Else
    wDtTo = ADC("UoDtTo")
  End If
  
  Dim wMonEndDt As String
  wb_MonEnd = False: wb_MonStart = False
  '6.1
   If moCn.RecSeek("Select 'x' where convert(varchar(15),DateAdd(Day,-1,DateAdd(Month ,1," + IIF(UCase(MWLib.DtSetting) = "MM/DD/YY", "Cast(Month('" + CStr(wDtTo) + "') as Varchar(2))+'/01'", "'01/'+Cast(Month('" + CStr(wDtTo) + "') as Varchar(2))") + "+'/'+Cast(Year('" + CStr(wDtTo) + "')   as Varchar(4))))," + CStr(DtFmtSlahVal()) + ")='" + wDtTo + "'") Then wb_MonEnd = True
   If CInt(Format(CStr(ADC("UODTFr")), "dd")) = 1 Then wb_MonStart = True
    wb_ShowOpn = IIF(ADC("UOYN5") = "Y", True, False)
    
    ADC("UOYN2") = IIF(CInt(Format(wDtFr, "dd")) <= 15, "N", "Y")
    If wb_MonEnd Then ADC("UOYN2") = "Y"
    
  
  ' ****** Manali 3.4.1 - "01/01/80" converted to purge dt

  If ADC("UOYN2") = "Y" Then
    ' ****** Manali 3.4.1 - 02/12/08 - [UoDtFr, UoDtTo respectively replaced by wDtFr, wDtTo]
    '6.1 added in if and else
    If CInt(Format(wDtFr, "mm")) > 0 And CInt(Format(wDtFr, "mm")) < 12 Then
        wTmpDt = Format$(DateAdd("d", -1, CDate(IIF(UCase(MWLib.DtSetting) = "MM/DD/YY", Format$(CInt(Format(wDtFr, "mm")) + 1, "00") + "/01", "01/" + Format$(CInt(Format(wDtFr, "mm")) + 1, "00")) + "/" + CStr(CInt(Format(ADC("UoDtFr"), "yy"))))), DtShortStr())
    ElseIf CInt(Format(wDtFr, "mm")) = 12 Then
        wTmpDt = Format$(DateAdd("d", -1, CDate("01/01/" + CStr(CInt(Format(wDtFr, "yy")) + 1))), DtShortStr())
    End If
    
'    Bef 3.4.1
'    If CInt(Format(adc("UoDtFr"), "mm")) > 0 And CInt(Format(adc("UoDtFr"), "mm")) < 12 Then
'        wTmpDt = Format$(DateAdd("d", -1, CDate("01/" + Format$(CInt(Format(adc("UoDtFr"), "mm")) + 1, "00") + "/" + CStr(CInt(Format(adc("UoDtFr"), "yy"))))), "dd/MM/yy")
'    ElseIf CInt(Format(adc("UoDtFr"), "mm")) = 12 Then
'        wTmpDt = Format$(DateAdd("d", -1, CDate("01/01/" + CStr(CInt(Format(adc("UoDtFr"), "yy")) + 1))), "dd/MM/yy")
'    End If
  End If
  
  If ADC("UOYN3") = "Y" Then
    Rep.TxtHead.SetText ("Stock Ledger (By Quantity)")
    'Rep.RecordSelectionFormula = "{rdo.qDrQty} <> 0 Or {rdo.qCrQty} <> 0 Or {rdo.qOpnQty} <> 0"
  ElseIf ADC("UOYN3") = "N" Then
    Rep.TxtHead.SetText ("Stock Ledger (By Weight)")
    'Rep.RecordSelectionFormula = "Round({rdo.qDrWt}, 3) <> 0 Or Round({rdo.qCrWt}, 3) <> 0 Or Round({rdo.qOpnWt}, 3) <> 0"
  End If
  
  If UCase(ADC("UoYN1")) = UCase("Y") Then
    Rep.h1PureWt.SetText ("Z Pure Wt")
    Rep.h2PureWt.SetText ("Z Pure Wt")
    wPureWtOrZWt = "RmPurityZ"
    wXPureWtOrZwt = "(Select XZ.RmPurityZ from RmMst XZ where XZ.RmCd=Rm.RmBaseCd " + IIF(gs_Partition = ctCurrPrtn, " And XZ.RmPrtKey=Rm.RmPrtKey)", ")")
  ElseIf UCase(ADC("UoYN1")) = UCase("N") Then
    wPureWtOrZWt = "RmPurityWt"
    wXPureWtOrZwt = "RmBasePurityWt"
  End If

  '*** Bef 1.11
  'wPurityWt = " /(case when Rm.RmCtg in ('G', 'P') then " + _
  '                        "IsNull((Select RmPurityWt from Param join RmMst on " + _
  '                        "RmCd= PValue where PTyp= 'RMCTG' and PMCd= Rm.RmCtg), 1) " + _
  '              "else 1 end) "
  
  If UCase(ADC("UoPMCdFr")) = "STKVAL" Then
    wSyRate = " 0 as qRmRt "
    wTdRate = " 0 as qRmRt "
    wFrRate = " 0 as qRmRt "
    wIrRate = " 0 as qRmRt "
    
    ws_Val = ", sum(convert (decimal (16, 3), (qOpnWt+ qDrWt- qCrWt)* qStkRt )) as qVal "
    
    GRP_REP.SetFormula Rep, "wRmVal", "{rdo.qVal}"
    GRP_REP.SetFormula Rep, "wRmClsVal", IIF(wb_MonEnd, "if {rdo.qFlag}='O' Then {rdo.qVal} Else 0", "{rdo.qVal}")
    GRP_REP.SetFormula Rep, "wQW", "''"
    Rep.HDtRmVal.SetText "StkVal"
    Rep.HGrpRmVal.SetText "StkVal"
    
    '4.1.4
'    If UCase(adc("UoPMCdTo")) <> "STKVAL" Then   '***bhavna changed for error correction
     If UCase(ADC("UoPMCdTo")) = "CSTVAL" Or UCase(ADC("UoPMCdTo")) = "SALVAL" Then
       wSyRate = IIF(UCase(ADC("UoPMCdTo")) = "CSTVAL", " (Case When RSCtg.PValue1='N' Then " + RmRate("RrCstRt", "SRmSz", "RmQwCst") + " Else 0 End) ", _
                                                        " (Case When RSCtg.PValue1='N' Then " + RmRate("RrSalRt", "SRmSz", "RmQw") + " Else 0 End) ") + " as qRmRt"
       wTdRate = IIF(UCase(ADC("UoPMCdTo")) = "CSTVAL", " (Case When RSCtg.PValue1='N' Then " + RmRate("RrCstRt", "TdRmSz", "RmQwCst") + " Else 0 End) ", _
                                                        " (Case When RSCtg.PValue1='N' Then " + RmRate("RrSalRt", "TdRmSz", "RmQw") + " Else 0 End) ") + " as qRmRt"
       wFrRate = IIF(UCase(ADC("UoPMCdTo")) = "CSTVAL", " (Case When RSCtg.PValue1='N' Then " + RmRate("RrCstRt", "FrRmSz", "RmQwCst") + " Else 0 End) ", _
                                                        " (Case When RSCtg.PValue1='N' Then " + RmRate("RrSalRt", "FrRmSz", "RmQw") + " Else 0 End) ") + " as qRmRt"
       wIrRate = IIF(UCase(ADC("UoPMCdTo")) = "CSTVAL", " (Case When RSCtg.PValue1='N' Then " + RmRate("RrCstRt", "IrRmSz", "RmQwCst") + " Else 0 End) ", _
                                                        " (Case When RSCtg.PValue1='N' Then " + RmRate("RrSalRt", "IrRmSz", "RmQw") + " Else 0 End) ") + " as qRmRt"
       
       ws_Val = ", Sum(Convert (Decimal (16, 3), (Case When " + IIF(UCase(ADC("UoPMCdTo")) = "CSTVAL", " (Case When qStkbyRt='N' Then qQWCst Else qQW End) ", " qQW ") + " = 'Q' then " + _
                " (qOpnQty+ qDrQty- qCrQty) else (Case When qRmPurityWt<> 0 and RmCtg<>'X' Then qRmPurityWt Else 1 End)*(qOpnWt+ qDrWt- qCrWt) end)*(Case When qStkbyRt='N' Then qRmRt Else qStkRt End))) as qVal "
       
'       If wb_MonEnd Then ws_Val = ", Sum(Convert (Decimal (16, 3), (Case When " + IIF(UCase(adc("UoPMCdTo")) = "CSTVAL", " (Case When qStkbyRt='N' Then qQWCst Else qQW End) ", " qQW ") + " = 'Q' then " + _
'                " ((Case when TdNo<>0 Then  (qOpnQty+ qDrQty- qCrQty) Else qOpnQty End)  else (Case When qRmPurityWt<> 0 Then qRmPurityWt Else 1 End)*( Case when TdNo<>0 Then  (qOpnWt+ qDrWt- qCrWt)  Else qOpnWt End)  end)*(Case When qStkbyRt='N' Then qRmRt Else qStkRt End))) as qVal "

       If ADC("UoRmCtgFr") <> "X" And (UCase(ADC("UoPMCdTo")) = "CSTVAL" Or UCase(ADC("UoPMCdTo")) = "SALVAL") Then
        Rep.HDtRmVal.SetText IIF(UCase(ADC("UoPMCdTo")) = "CSTVAL", "Stk-CstVal", "Stk-SalVal")
        Rep.HGrpRmVal.SetText IIF(UCase(ADC("UoPMCdTo")) = "CSTVAL", "Stk/CstVal", "Stk/SalVal")
       End If
    End If

  ElseIf UCase(ADC("UoPMCdFr")) = "CSTVAL" Then
    wSyRate = RmRate("RrCstRt", "SRmSz", "RmQwCst") + " as qRmRt"
    wTdRate = RmRate("RrCstRt", "TdRmSz", "RmQwCst") + " as qRmRt"
    wFrRate = RmRate("RrCstRt", "FrRmSz", "RmQwCst") + " as qRmRt"
    wIrRate = RmRate("RrCstRt", "IrRmSz", "RmQwCst") + " as qRmRt"
    
    ws_Val = ", sum(convert (decimal (16, 3), (case when qQWCst= 'Q' then (qOpnQty+ qDrQty- qCrQty) else (Case When qRmPurityWt<> 0 and RmCtg<>'X' Then qRmPurityWt Else 1 End)* (qOpnWt+ qDrWt- qCrWt) end)* qRmRt)) as qVal "
    
    'If wb_MonEnd Then ws_Val = ", sum(convert (decimal (16, 3), (case when qQWCst= 'Q' then (Case when TdNo<>0 Then  (qOpnQty+ qDrQty- qCrQty) Else qOpnQty End)  else (Case When qRmPurityWt<> 0 Then qRmPurityWt Else 1 End)* (Case when TdNo<>0 Then (qOpnWt+ qDrWt- qCrWt) Else qOpnWt End) end)* qRmRt)) as qVal "
    
    
    GRP_REP.SetFormula Rep, "wRmVal", "{rdo.qVal}"
    GRP_REP.SetFormula Rep, "wRmClsVal", IIF(wb_MonEnd, "if {rdo.qFlag}='O' Then {rdo.qVal} Else 0", "{rdo.qVal}")
    Rep.HDtRmVal.SetText "CstVal"
    Rep.HGrpRmVal.SetText "CstVal"
                                      
  ElseIf UCase(ADC("UoPMCdFr")) = "SALVAL" Then
    wSyRate = RmRate("RrSalRt", "SRmSz", "RmQw") + " as qRmRt"
    wTdRate = RmRate("RrSalRt", "TdRmSz", "RmQw") + " as qRmRt"
    wFrRate = RmRate("RrSalRt", "FrRmSz", "RmQw") + " as qRmRt"
    wIrRate = RmRate("RrSalRt", "IrRmSz", "RmQw") + " as qRmRt"
    
    ws_Val = ", sum(convert (decimal (16, 3), (case when qQW= 'Q' then (qOpnQty+ qDrQty- qCrQty) else (Case When qRmPurityWt<> 0  and RmCtg<>'X' Then qRmPurityWt Else 1 End)* (qOpnWt+ qDrWt- qCrWt) end)* qRmRt)) as qVal "
   ' If wb_MonEnd Then ws_Val = ", sum(convert (decimal (16, 3), (case when qQW= 'Q' then (Case when TdNo<>0 Then (qOpnQty+ qDrQty- qCrQty) Else qOpnQty End)  else (Case When qRmPurityWt<> 0 Then qRmPurityWt Else 1 End)* (Case when TdNo<>0 Then (qOpnWt+ qDrWt- qCrWt) Else qOpnWt End) end)* qRmRt)) as qVal "
    
    GRP_REP.SetFormula Rep, "wRmVal", "{rdo.qVal}"
    GRP_REP.SetFormula Rep, "wRmClsVal", IIF(wb_MonEnd, "if {rdo.qFlag}='O' Then {rdo.qVal} Else 0", "{rdo.qVal}")
    Rep.HDtRmVal.SetText "SalVal"
    Rep.HGrpRmVal.SetText "SalVal"
 ElseIf (UCase(ADC("UoPMCdFr")) = "LOTV-BAS" Or UCase(ADC("UoPMCdFr")) = "LOTV-INR") Then
    wSyRate = RmRate("RrCstRt", "SRmSz", "RmQwCst", IIF(UCase(ADC("UoPMCDFr")) = "LOTV-INR", "RS", ""), True) + " as qRmRt"
    wTdRate = RmRate("RrCstRt", "TdRmSz", "RmQwCst", IIF(UCase(ADC("UoPMCDFr")) = "LOTV-INR", "RS", ""), True) + " as qRmRt"
    wFrRate = RmRate("RrCstRt", "FrRmSz", "RmQwCst", IIF(UCase(ADC("UoPMCDFr")) = "LOTV-INR", "RS", ""), True) + " as qRmRt"
    wIrRate = RmRate("RrCstRt", "IrRmSz", "RmQwCst", IIF(UCase(ADC("UoPMCDFr")) = "LOTV-INR", "RS", ""), True) + " as qRmRt"
    
    ws_Val = ", sum(convert (decimal (16, 3), (case when qQWCst= 'Q' then (qOpnQty+ qDrQty- qCrQty) else (Case When qRmPurityWt<> 0 and RmCtg<>'X' Then qRmPurityWt Else 1 End)* (qOpnWt+ qDrWt- qCrWt) end)* qRmRt)) as qVal "
    
    'If wb_MonEnd Then ws_Val = ", sum(convert (decimal (16, 3), (case when qQWCst= 'Q' then (Case when TdNo<>0 Then  (qOpnQty+ qDrQty- qCrQty) Else qOpnQty End)  else (Case When qRmPurityWt<> 0 Then qRmPurityWt Else 1 End)* (Case when TdNo<>0 Then (qOpnWt+ qDrWt- qCrWt) Else qOpnWt End) end)* qRmRt)) as qVal "
    
    
    GRP_REP.SetFormula Rep, "wRmVal", "{rdo.qVal}"
    GRP_REP.SetFormula Rep, "wRmClsVal", IIF(wb_MonEnd, "if {rdo.qFlag}='O' Then {rdo.qVal} Else 0", "{rdo.qVal}")
    Rep.HDtRmVal.SetText IIF(UCase(ADC("UoPMCDFr")) = "LOTV-INR", "LotV-INR", "LotV-BAS")
    Rep.HGrpRmVal.SetText IIF(UCase(ADC("UoPMCDFr")) = "LOTV-INR", "LotV-INR", "LotV-BAS")
  ElseIf UCase(ADC("UoPMCdFr")) = "" Then
    wSyRate = "0 as qRmRt"
    wTdRate = "0 as qRmRt"
    wFrRate = "0 as qRmRt"
    wIrRate = "0 as qRmRt"
    
    ws_Val = ", 0 as qVal "         '*** (Jen 2.12 Patch 1)
        
    '*** (Jen 2.13)
    If ADC("UoYN") = "Y" And wb_ShowOpn Then
    '*** (Jen 2.13)
      If ADC("UoYn3") = "Y" Then
      
       GRP_REP.SetFormula Rep, "wRmVal", "{rdo.qOpnwt}+ {rdo.qDrWt}- {rdo.qCrWt}"
       GRP_REP.SetFormula Rep, "wRmClsVal", IIF(wb_MonEnd, "{rdo.qOpnwt}", "{rdo.qOpnwt}+ {rdo.qDrWt}- {rdo.qCrWt}")
        '*** (Bef 2.12) Rep.g4RmVal.DecimalPlaces = 3: Rep.g4RmVal.RoundingType = crRoundToThousandth
        Rep.wRmVal.DecimalPlaces = 3: Rep.wRmVal.RoundingType = crRoundToThousandth '*** (Jen 2.12)
        Rep.g3RmVal.DecimalPlaces = 3: Rep.g3RmVal.RoundingType = crRoundToThousandth
        Rep.g2RmVal.DecimalPlaces = 3: Rep.g2RmVal.RoundingType = crRoundToThousandth
        Rep.g1RmVal.DecimalPlaces = 3: Rep.g1RmVal.RoundingType = crRoundToThousandth
        Rep.gRmVal.DecimalPlaces = 3: Rep.gRmVal.RoundingType = crRoundToThousandth
        Rep.HDtRmVal.SetText "RmWt"
        Rep.HGrpRmVal.SetText "RmWt"
        
      ElseIf ADC("UoYn3") = "N" Then
        GRP_REP.SetFormula Rep, "wRmVal", "{rdo.qOpnQty}+ {rdo.qDrQty}- {rdo.qCrQty}"
        GRP_REP.SetFormula Rep, "wRmClsVal", IIF(wb_MonEnd, "{rdo.qOpnQty}", "{rdo.qOpnQty}+ {rdo.qDrQty}- {rdo.qCrQty}")
        Rep.wRmVal.DecimalPlaces = 0: Rep.wRmVal.RoundingType = crRoundToUnit '*** (Jen 2.12)
        '*** (Bef 2.12) Rep.g4RmVal.DecimalPlaces = 0: Rep.g4RmVal.RoundingType = crRoundToUnit
        Rep.g3RmVal.DecimalPlaces = 0: Rep.g3RmVal.RoundingType = crRoundToUnit
        Rep.g2RmVal.DecimalPlaces = 0: Rep.g2RmVal.RoundingType = crRoundToUnit
        Rep.g1RmVal.DecimalPlaces = 0: Rep.g1RmVal.RoundingType = crRoundToUnit
        Rep.gRmVal.DecimalPlaces = 0: Rep.gRmVal.RoundingType = crRoundToUnit
        Rep.HDtRmVal.SetText "RmQty"
        Rep.HGrpRmVal.SetText "RmQty"
      End If
      
    '*** (Jen 2.13)
    Else
      Rep.wRmVal.Suppress = True
      Rep.g3RmVal.Suppress = True
      Rep.g2RmVal.Suppress = True
      Rep.g1RmVal.Suppress = True
      Rep.gRmVal.Suppress = True
      Rep.HDtRmVal.Suppress = True
      Rep.HGrpRmVal.Suppress = True
    End If
    '*** (Jen 2.13)
  '********************************'
  End If
            
    ' ***** Manali 3.03 - 24/06/08 - Pd Mod - 'L' added
 '   wPurityWt = " /(Case When Rm.RmCtg In ('G', 'P', 'S', 'L', 'M') Then " + _
                          "IsNull((Select RmPurityWt from Param Join RmMst on " + _
                          IIF(gs_Partition = ctCurrPrtn, " RmPrtKey='" + ctCurrPrtn + "' and ", "") + " RmCd= PValue where PTyp= 'RMCTG' and PMCd= Rm.RmCtg), 1) " + _
                 "When Rm.RmCtg= 'X' then " + _
                          "IsNull((Select RmPurityWt from Param Join RmMst on " + _
                          IIF(gs_Partition = ctCurrPrtn, " RmPrtKey='" + ctCurrPrtn + "' and ", "") + " RmCd= PValue where PTyp= 'RMCTG' and " + _
                          "PMCd= IsNull((Select RmCtg From RmMst where " + IIF(gs_Partition = ctCurrPrtn, " RmPrtKey='" + ctCurrPrtn + "' and ", "") + " RmCd= Rm.RmBaseCd), '')), 1) " + _
                 "Else 1 end) "
  
  wPurityWt = " /(case when Rm.RmCtg in ('G', 'P', 'S', 'L', 'M') then " + _
                          "IsNull((Select " + wPureWtOrZWt + " from Param join RmMst on " + _
                          IIF(gs_Partition = ctCurrPrtn, " RmPrtKey='" + ctCurrPrtn + "' and ", "") + " RmCd= PValue where PTyp= 'RMCTG' and PMCd= Rm.RmCtg), 1) " + _
                 "when Rm.RmCtg= 'X' then " + _
                          "IsNull((Select " + wPureWtOrZWt + " from Param join RmMst on " + _
                          IIF(gs_Partition = ctCurrPrtn, " RmPrtKey='" + ctCurrPrtn + "' and ", "") + " RmCd= PValue where PTyp= 'RMCTG' and " + _
                          "PMCd= IsNull((Select RmCtg From RmMst where " + IIF(gs_Partition = ctCurrPrtn, " RmPrtKey='" + ctCurrPrtn + "' and ", "") + " RmCd= Rm.RmBaseCd), '')), 1) " + _
                 "else 1 end) "
 
  ' ****** Sachin 2.13.0 - 05/06/2006 ******
  ' ***** Manali 3.03 - 24/06/08 - Pd Mod - 'L' added
    wQwChk = " IsNull((Case when (RmCtg in ('G', 'P', 'S', 'L', 'M') And RmQw='W') Then (Select RmPr.RmQw From RmMst RmPr " + _
              "   Where " + IIF(gs_Partition = ctCurrPrtn, " RmPr.RmPrtKey='" + ctCurrPrtn + "' and ", "") + " RmPr.RmCd=(Select PValue From param " + _
              "       Where PTyp= 'RMCTG' and PMCd= Rm.RmCtg)) Else RmQw End),'')  as qQw," + _
              "IsNull((Case when (RmCtg In ('G', 'P', 'S', 'L', 'M') And RmQwCst = 'W') Then (Select RmPr.RmQwCst From RmMst RmPr " + _
              "   Where " + IIF(gs_Partition = ctCurrPrtn, " RmPr.RmPrtKey='" + ctCurrPrtn + "' and ", "") + " RmPr.RmCd=(Select PValue From param " + _
              "       Where PTyp= 'RMCTG' and PMCd = Rm.RmCtg)) Else RmQwCst End),'') as qQwCst "
              
             
  ' ****** Sachin 2.13.0 - 05/06/2006 ******
    
'  Select space(30) as Grp1, space(30) as Grp2, space(30) as Grp3,
'  space(30) as DGrp1, space(30) as DGrp2, space(30) as DGrp3,
'  space(30) as G1Desc, space(30) as G2Desc, space(30) as G3Desc,
'  space(1) as qFlag, RmCd, TdRmSz as qRmSz, RisRmSzDesc as qRmSzDesc, RrRngSz, RmCtg, PSCd As qStnCls, RmSCtg,
'  TdRmStkRt as qStkRt, PValue3 as qPtrYN, RmPurityWt as qRmPurityWt, TdCoCd, TdTc, TdYy,
'  TdChr, TdNo, TdSr, TdSrNo, TdBYy, TdBChr, TdBNo, TDt as qDt,
'  LocPrntCd, TdToRmLoc as qLoc, 0 as qOpnQty, TdRmQty as qDrQty,
'  TdRmQty as qCrQty, 0.0 as qOpnWt, TdRmWt as qDrWt, TdRmWt as qCrWt,
'  TdFrRmLoc As qFrRmLoc, TdToRmLoc As qToRmLoc, TdRmStkRt As qRmRt,
'  space(1) As qQW, space(1) As qQWCst
'  From Txnd, Txn, RmMst, Loc, RmRt, Param, RmIdSz where 1=2

'  "Select space(30) as Grp1, space(30) as Grp2, space(30) as Grp3, " + _
'  "space(30) as DGrp1, space(30) as DGrp2, space(30) as DGrp3, " + _
'  "space(30) as G1Desc, space(30) as G2Desc, space(30) as G3Desc, " + _
'  "space(1) as qFlag, RmCd, TdRmSz as qRmSz, RisRmSzDesc as qRmSzDesc, RrRngSz, RmCtg, PSCd As qStnCls, RmSCtg, " + _
'  "TdRmStkRt as qStkRt, PValue3 as qPtrYN, RmPurityWt as qRmPurityWt, TdCoCd, TdTc, TdYy, " + _
'  "TdChr, TdNo, TdSr, TdSrNo, TdBYy, TdBChr, TdBNo, TDt as qDt, " + _
'  "LocPrntCd, TdToRmLoc as qLoc, 0 as qOpnQty, TdRmQty as qDrQty, " + _
'  "TdRmQty as qCrQty, 0.0 as qOpnWt, TdRmWt as qDrWt, TdRmWt as qCrWt, " + _
'  "TdFrRmLoc As qFrRmLoc, TdToRmLoc As qToRmLoc, TdRmStkRt As qRmRt, " + _
'  "space(1) As qQW, space(1) As qQWCst " + _
'  "Into #TmpStkLed From Txnd, Txn, RmMst, Loc, RmRt, Param where 1=2 " + _

  'Setting the opening tmp table
  'wSetTmpTab = "Select space(30) as Grp1, space(30) as Grp2, space(30) as Grp3, " + _
            "space(30) as DGrp1, space(30) as DGrp2, space(30) as DGrp3,  " + _
            "space(30) as G1Desc, space(30) as G2Desc, space(30) as G3Desc, " + _
            "Space(1) as qFlag, RMCd, SRmSz, RisRmSzDesc as qRmSzDesc, IsNull(RrRngSz, '') RrRngSz, RmCtg, " + _
            "RmSCtg, SRmStkRt as qStkRt, PValue3 as qPtrYN, RmPurityWt as qRmPurityWt, " + _
            "TdCoCd, TdTc, TdYy, TdChr, TdNo, TdSr, TdSrNo, TdBYy, TdBChr, TdBNo, TDt, " + _
            "LocCd, (SRmDrQty - SRmCrQty) as qOpnQty, SRmDrQty as qDrQty, " + _
            "SRmCrQty as qCrQty, (SRmDrWt- SRmCrWt) as qOpnWt, SRmDrWt as qDrWt, SRmCrWt as qCrWt, " + _
            "TdFrRmLoc as qFrRmLoc, TdToRmLoc as qToRmLoc, TdRmStkRt as qRmRt, RmQW as qQw, RmQWCst as qQwCst  " + _
            "Into #TmpStkLed From SYyMm, Txn, Txnd, Loc, RmMst, RmRt, Param Where 1=2"
  ' Zubin 212
  
'*** Jay 2.14 *** (TDesc as qDesc added)
' *** Manali 3.03 - (RisRmSzDesc as qRmSzDesc) added
'*** REPORT SQL ***
'Select space(30) as Grp1, space(30) as Grp2, space(30) as Grp3,
'space(30) as DGrp1, space(30) as DGrp2, space(30) as DGrp3,
'space(30) as G1Desc, space(30) as G2Desc, space(30) as G3Desc,
'space(1) as qFlag, RmCd, TdLotNo as qLotNo, TdRmSz as qRmSz, RisRmSzDesc as qRmSzDesc, RrRngSz, RmCtg, PSCd As qStnCls, RmSCtg,
'TdRmStkRt as qStkRt, PValue3 as qPtrYN, TdRmWt as qPureWt, TdCoCd, TdTc, TdYy,
'TdChr, TdNo, TdSr, TdSrNo, TdBYy, TdBChr, TdBNo, TDt as qDt, TDesc as qDesc,
'LocPrntCd, TdToRmLoc as qLoc, 0 as qOpnQty, TdRmQty as qDrQty,
'TdRmQty as qCrQty, 0.0 as qOpnWt, TdRmWt as qDrWt, TdRmWt as qCrWt,
'TdFrRmLoc As qFrRmLoc, TdToRmLoc As qToRmLoc, TdRmStkRt As qRmRt,
'space(1) As qQW, space(1) As qQWCst, TdRmStkRt as qVal
'From Txnd, Txn, RmMst, Loc, RmRt, Param, RmIdSz where 1=2
'*** Jay 2.14 ***

'*** Jay 2.14 ***
If UCase(ADC("UoYn4")) = "Y" Then
  If UCase(ADC("UoVchOpt")) = UCase("Vch") Then qDesc = " TDesc " Else qDesc = " TdDesc "
Else
  qDesc = " '' "
  Rep.DETDesc.Suppress = True
End If
'*** Jay 2.14 ***
'*** Jay 2.14 ***(Tdesc as qDesc, added)
' **** Manali 3.03 - 09/07/08 - 'RisRmSzDesc as qRmSzDesc' addded
  wSetTmpTab = "Select space(60) as Grp1, space(60) as Grp2, space(60) as Grp3, " + _
            " space(60) as DGrp1, space(60) as DGrp2, space(60) as DGrp3, " + _
            " space(60) as G1Desc, space(60) as G2Desc, space(60) as G3Desc, " + _
            " space(1) as qFlag, RmCd, SLotNo as qLotNo, SRmSz as qRmSz, SRmSz2 as qRmSz2, SRmSz3 as qRmSz3, RisRmSzDesc as qRmSzDesc, RrRngSz, RmCtg, PSCd As qStnCls, RmSCtg, " + _
            " SRmStkRt as qStkRt, PValue3 as qPtrYN, RmPurityWt as qRmPurityWt, " + _
            " TdCoCd, TdTc, TdYy, TdChr, TdNo, TdSr, TdSrNo, TdBYy, TdBChr, TdBNo, TDt as qDt, Tdesc as qDesc, " + _
            " LocPrntCd as qPrntCd, LocCd as qLoc, (SRmDrQty - SRmCrQty) as qOpnQty, SRmDrQty as qDrQty, " + _
            " SRmCrQty as qCrQty, (SRmDrWt- SRmCrWt) as qOpnWt, SRmDrWt as qDrWt, SRmCrWt as qCrWt, " + _
            " TdFrRmLoc As qFrRmLoc, TdToRmLoc As qToRmLoc, TdRmStkRt As qRmRt, RmQw As qQw, RmQwCst As qQwCst, PValue1 as qStkbyRt " + _
            " Into #TmpStkLed From SYyMm, Txn, Txnd, Loc, RmMst, RmRt, Param, RmIdSz where 1=2 "

  moCn.CreateTmpTable "#TmpStkLed", wSetTmpTab
  
  ADC("UoLn1Fr").CmpStr = "SRmSz>= "
  ADC("UoLn1To").CmpStr = "SRmSz<= "
  ADC("UoCstRtFr").CmpStr = "SRmStkRt>= "
  ADC("UoCstRtTo").CmpStr = "SRmStkRt<= "
  ' Zubin 212
  ADC("UoLotNoFr").CmpStr = "SLotNo>= "
  ADC("UoLotNoTo").CmpStr = "SLotNo<= "
  ADC("UoLotNoSel").CmpStr = "SLotNo In "
  
  wrepcnd = ADC.RepCond
  
  'wCnd = IIF(wrepcnd <> "", " and " + wrepcnd , "")
  wCnd = IIF(wrepcnd <> "", " and " + wrepcnd + " Option (MaxDop 1)", " Option (MaxDop 1)")
  
  ' ****** Manali 3.4.1 - 02/12/08 - [UoDtFr, UoDtTo respectively replaced by wDtFr, wDtTo]
  wDtCndOpnYyMm = " where SYyMm  " + IIF(wb_MonEnd, "<=", "<") + Format(IIF(wb_MonEnd, wDtTo, wDtFr), "yyyymm")
  'wDtCndOpnYyMm = " where SYyMm < " + Format(adc("UoDtFr"), "yyyymm")
  
  If ADC("UoOpnYyMm") <> 0 Then
    wOpnYyStr = " and SyyMm >= " + CStr(ADC("UoOpnYyMM")) + ""
  Else
    wOpnYyStr = ""
  End If
  
  '*****************  Reverse Logic urmi 2.04 *******************
  ' eg. UoDtFr = 20/7/2002 UoDtTo 22/4/03 and Todays dt = 25/8/03
  ' wDtCndOpnYyMm = "Where SYyMm <= 200207 "
  ' ****** Manali 3.4.1 - 02/12/08 - [UoDtFr, UoDtTo respectively replaced by wDtFr, wDtTo]
  If ADC("UoYn2") = "Y" Then
    If wb_MonEnd Then
        wDtCndOpnYyMm = " where SYyMm <= " + Format(wDtTo, "yyyymm")
    Else
        wDtCndOpnYyMm = " where SYyMm <= " + Format(wDtFr, "yyyymm")
    End If
  End If
  'If adc("UoYn2") = "Y" Then wDtCndOpnYyMm = " where SYyMm <= " + Format(adc("UoDtFr"), "yyyymm")
  '**************************************************************
  
  
  '*** Manali 3.03 - 07/07/08 - RmSzDesc
  wRmSzDesc = "case when (Select PDesc225 from Param where PTyp= 'RMSCTG' and PMCd= IsNull(RisRmCtg, '') and PSCd= RmSCtg)= 'Y' then IsNull(RisRmSzDesc, '') Else '' End "
  '*** Manali 3.03 - 07/07/08 - RmSzDesc
  
  '"RmPurityWt" + wPurityWt + " as qRmPurityWt, "
  '1
  ' Zubin 212
  ' ****** Sachin 2.13.0 - StnCls added to all queries ******
  '*** jay 2.14 *** ('' as qDesc,)
  ' ***** Manali 3.03 - 24/06/08 - Pd Mod - 'L' added
  ' ****** Manali 3.03 - 09/07/08 - RmSzDesc added and join with RmIdSz added
  '4.1.4
  'JewelSmith.2 ',' + SC.PDesc225 + ',' like '%,'+ RmSCtg +',%' changed as SC.PSCd = RmStnCls
  wsqlOpnYyMm = "Insert Into #TmpStkLed Select " + grpflds0 + ", " + _
                "'O' as qFlag, RmCd, SLotNo as qLotNo, SRmSz as qRmSz, SRmSz2 as qRmSz2, SRmSz3 as qRmSz3, " + wRmSzDesc + " as qRmSzDesc, IsNull(RrRngSz, '') as RrRngSz, RmCtg, IsNull(SC.PSCd,'') As qStnCls, RmSCtg, " + _
                "SRmStkRt as qStkRt, " + _
                "(Select PValue3 from Param where PTyp= 'RMSCTG' and PMCd= RmCtg and PSCd= RmSCtg) as qPtrYN, " + _
                "(Case when RmCtg in ('G', 'P', 'S', 'L', 'M') then " + wPureWtOrZWt + " when RmCtg= 'X' then " + wXPureWtOrZwt + " else 0 end)" + wPurityWt + " as qRmPurityWt, " + _
                "SCoCd as TdCoCd, '' as TdTc, '' as TdYy, '' as TdChr, 0 as TdNo, 0 as TdSr, 0 as TdSrNo, " + _
                "'' as TdBYy, '' as TdBChr, 0 as TdBNo, '' as qDt, '' as qDesc, LocPrntCd as qPrntCd, SLoc as qLoc, " + _
                "SRmDrQty- SRmCrQty as qOpnQty, 0 as qDrQty, 0 as qCrQty, " + _
                "(SRmDrWt- SRmCrWt) as qOpnWt, 0 as qDrWt, 0 as qCrWt, '' as qFrRmLoc, '' as qToRmLoc, " + wSyRate + _
                ", " + wQwChk + ", RSCtg.PValue1 as qStkbyRt " + _
                "From SYyMm NoLock Join Loc Lc On LocCoCd = SCoCd and LocCd= SLoc and LocHideYN='N'  " + _
                " Join RmMst Rm On RmCd= SRmCd " + IIF(gs_Partition = ctCurrPrtn, " and RmPrtKey=SPrtKey ", "") + _
                " Join Param RSCtg On RSCtg.PTyp='RMSCTG' And RSCtg.PMCd=Rm.RmCtg And RSCtg.PSCd=Rm.RmSCtg " + _
                " Left Outer Join Param SC on SC.PTyp= 'STNCLS' and SC.PmCd= RmCtg and SC.PSCd = RmStnCls " + _
                " Left Outer Join RmRt RSZ on RrTcTyp= 'RSZ' and RrCmCtg= 'C' and RrCmCd= '" + ctSelfCmCd + "' and " + _
                " RrCtg= RmCtg and RrSCtg= RmSCtg and RrSTWGrp= '' and RrCd= '' and RrDmCtg= '' and " + _
                " RrLsCtg= '' and RrLabMCd= '' and RrFrLn<= SRmSz and RrToLn>= SRmSz " + _
                " Left Outer Join LotMst on LotMst.LotNo = SLotNo " + _
                " Left Outer Join RmIdSz on RisRmCtg=RmCtg and RisRmId=RmId and RisSz1=SRmSz " + _
                wDtCndOpnYyMm + wOpnYyStr + IIF(gs_Partition = ctCurrPrtn, " and SPrtKey='" + ctCurrPrtn + "' ", " and SYyMm<>'0' ") + wCnd
                
  If wb_ShowOpn Then moCn.Execute wsqlOpnYyMm
  
  ' ****** Manali 3.4.1 - 02/12/08 - [UoDtFr, UoDtTo respectively replaced by wDtFr, wDtTo]
  wFrDt = IIF(UCase(MWLib.DtSetting) = "MM/DD/YY", Format(wDtFr, "mm") + "/01", "01/" + Format(wDtFr, "mm")) + "/" + Format(wDtFr, "yy") '6.1
  'wFrDt = "01/" + Format(adc("UoDtFr"), "mm") + "/" + Format(adc("UoDtFr"), "yy")
  '****** Sachin 3.02.0 - New fields
  ' ****** Manali 3.4.1 - 02/12/08 - [UoDtFr, UoDtTo respectively replaced by wDtFr, wDtTo]
  wDtCndOpnTd = " where TDt >= '" + wFrDt + "' And TDt < '" + CStr(wDtFr) + "' And TdDt >= '" + wFrDt + "' And TdDt < '" + CStr(wDtFr) + "' "
  'wDtCndOpnTd = " where TDt >= '" + wFrDt + "' And TDt < '" + CStr(adc("UoDtFr")) + "' And TdDt >= '" + wFrDt + "' And TdDt < '" + CStr(adc("UoDtFr")) + "' "
  
  '************* Reverse Logic urmi 5/7/03 *************************************
  wRLQty = "(Case TdFrRmDc When 'D' Then TdRmQty When 'C' Then -TdRmQty Else 0 End) as qOpnQty, "
  wRLWt = "(Case TdFrRmDc When 'D' Then TdRmWt When 'C' Then -TdRmWt Else 0 End) as qOpnWt, "
  
  If ADC("Uoyn2") = "Y" Then
'    wDtCndOpnTd = " where TDt >= '" & ADC("UoDtFr") & "' And month(TDt) = " & _
'                  CInt(Format(ADC("UoDtFr"), "mm")) & " and Year(TDt) = " & _
'                  CInt(Format(ADC("UoDtFr"), "yyyy"))

    '****** Sachin 3.02.0 - New fields
    ' ****** Manali 3.4.1 - 02/12/08 - [UoDtFr, UoDtTo respectively replaced by wDtFr, wDtTo]
    wDtCndOpnTd = " where TDt >= '" & wDtFr & "' And TDt <= '" & wTmpDt & "' And TdDt >= '" & wDtFr & "' And TdDt <= '" & wTmpDt & "' "
    'wDtCndOpnTd = " where TDt >= '" & adc("UoDtFr") & "' And TDt <= '" & wTmpDt & "' And TdDt >= '" & adc("UoDtFr") & "' And TdDt <= '" & wTmpDt & "' "
    
    wRLQty = "(Case TdFrRmDc When 'D' Then -TdRmQty When 'C' Then TdRmQty Else 0 End) as qOpnQty, "
    wRLWt = "(Case TdFrRmDc When 'D' Then -TdRmWt When 'C' Then TdRmWt Else 0 End) as qOpnWt, "
  End If
  '*******************************************************************************
  
  ADC("UoLn1Fr").CmpStr = "TdRmSz>= "
  ADC("UoLn1To").CmpStr = "TdRmSz<= "
  ADC("UoCstRtFr").CmpStr = "TdRmStkRt>= "
  ADC("UoCstRtTo").CmpStr = "TdRmStkRt<= "
  
  ' Zubin 212
  ADC("UoLotNoFr").CmpStr = "TdLotNo>= "
  ADC("UoLotNoTo").CmpStr = "TdLotNo<= "
  ADC("UoLotNoSel").CmpStr = "TdLotNo In "
  
  wrepcnd = ADC.RepCond
  'wCnd = IIF(wrepcnd <> "", " and " + wrepcnd, "")
  wCnd = IIF(wrepcnd <> "", " and TdAck='Y' and " + wrepcnd, " and TdAck='Y' ")
     '****** Sachin 3.02.0 - Additional Conditions For New Fields
      wAddlCnd = ""
      wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(ADC("UORMCTGFR")) <> "", " And ", "") + IIF(Trim(ADC("UORMCTGFR")) <> "", " TdRmCtg = '" + Trim(ADC("UORMCTGFR")) + "'", "")
      
      wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(ADC("UORMSCTGFR")) <> "", " And ", "") + IIF(Trim(ADC("UORMSCTGFR")) <> "", " TdRmSCtg >= '" + Trim(ADC("UORMSCTGFR")) + "'", "")
      wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(ADC("UORMSCTGTO")) <> "", " And ", "") + IIF(Trim(ADC("UORMSCTGTO")) <> "", " TdRmSCtg <= '" + Trim(ADC("UORMSCTGTO")) + "'", "")
      wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(ADC("UORMSCTGSEL")) <> "", " And ", "") + IIF(Trim(ADC("UORMSCTGSEL")) <> "", " TdRmSCtg In (" + Trim(ADC("UORMSCTGSEL")) + ")", "")
    '****** Sachin 3.02.0 - Additional Conditions For New Fields

  '2
  ' Zubin 212
  '*** jay 2.14 *** ('' as qDesc,)
  ' ***** Manali 3.03 - 24/06/08 - Pd Mod - 'L' added
  ' ****** Manali 3.03 - 09/07/08 - RmSzDesc added and join with RmIdSz added
  '4.1.4
  'JewelSmith.2 ',' + SC.PDesc225 + ',' like '%,'+ RmSCtg +',%' changed as SC.PSCd = RmStnCls
  wSqlOpnTdFr = "Insert Into #TmpStkLed Select " + grpflds1 + "," + _
                "'O' as qFlag, RmCd, TdLotNo as qLotNo, TdRmSz as qRmSz, TdRmSz2 as qRmSz2, TdRmSz3 as qRmSz3, " + wRmSzDesc + " as qRmSzDesc,  IsNull(RrRngSz, '') as RrRngSz, RmCtg, IsNull(SC.PSCd,'') As qStnCls, RmSCtg, " + _
                "TdRmStkRt as qStkRt, (Select PValue3 from Param where PTyp= 'RMSCTG' and PMCd= RmCtg and PSCd= RmSCtg) as qPtrYN, " + _
                "(Case when RmCtg in ('G', 'P', 'S', 'L', 'M') then " + wPureWtOrZWt + " when RmCtg= 'X' then " + wXPureWtOrZwt + " else 0 end)" + wPurityWt + " as qRmPurityWt, " + _
                "TCoCd as TdCoCd, '' as TdTc,'' as TdYy, '' as TdChr, 0 as TdNo, 0 as TdSr, 0 as TdSrNo, " + _
                "'' as TdBYy, '' as TdBChr, 0 as TdBNo, '' as qDt, '' as qDesc, LocPrntCd as qPrntCd, TdFrRmLoc as qLoc, " + _
                wRLQty + _
                "0 as qDrQty, 0 as qCrQty, " + _
                wRLWt + _
                "0 as qDrWt, 0 as qCrWt, '' as qFrRmLoc, '' as qToRmLoc, " + wTdRate + _
                ", " + wQwChk + ", RSCtg.PValue1 as qStkbyRt " + _
                " From Txn NoLock " + _
                " Join Txnd on TIdNo=TdTIdNo and TCoCd= TdCoCd and TTc= TdTc and TYy= TdYy and TChr= TdChr and TNo= TdNo " + IIF(gs_Partition = ctCurrPrtn, " and TPrtKey=TdPrtKey ", "") + _
                " Join Loc Lc On TdCoCd= LocCoCd and TdFrRmLoc= LocCd and LocHideYN='N' " + _
                " Join RmMst Rm On TdRmCd= RmCd " + IIF(gs_Partition = ctCurrPrtn, " and RmPrtKey=TdPrtKey ", "") + _
                " Join Param RSCtg On RSCtg.PTyp='RMSCTG' And RSCtg.PMCd=Rm.RmCtg And RSCtg.PSCd=RmSCtg  " + _
                " Left Outer Join Param SC on SC.PTyp= 'STNCLS' and SC.PmCd= RmCtg and SC.PSCd = RmStnCls " + _
                " Left Outer Join LotMst on LotMst.LotNo = TdLotNo " + _
                " Left Outer Join RmRt RSZ on RrTcTyp= 'RSZ' and RrCmCtg= 'C' and RrCmCd= '" + ctSelfCmCd + "' and " + _
                " RrCtg= RmCtg and RrSCtg= RmSCtg and RrSTWGrp= '' and RrCd= '' and RrDmCtg= '' and " + _
                " RrLsCtg= '' and RrLabMCd= '' and RrFrLn<= TdRmSz and RrToLn>= TdRmSz " + _
                "Left Outer Join RmIdSz on RisRmCtg=RmCtg and RisRmId=RmId and RisSz1=TdRmSz " + _
                 wDtCndOpnTd + wCnd + IIF(gs_Partition = ctCurrPrtn, " and TPrtKey='" + ctCurrPrtn + "' ", "") + " And TdFrRmLocTyp= '" + ADC("UOLOCTYPFR") + "'" + IIF(wAddlCnd <> "", " And " + wAddlCnd, "") + " Option (MaxDop 1)"
  If wb_ShowOpn And Not wb_MonStart And Not wb_MonEnd Then moCn.Execute wSqlOpnTdFr

  ADC("UoLn1Fr").CmpStr = "TdRmSz>= "
  ADC("UoLn1To").CmpStr = "TdRmSz<= "
  ADC("UoCstRtFr").CmpStr = "TdRmStkRt>= "
  ADC("UoCstRtTo").CmpStr = "TdRmStkRt<= "
  ' Zubin 212 (This CmpStr block is not req as it is the same for the above and below query)
  ADC("UoLotNoFr").CmpStr = "TdLotNo>= "
  ADC("UoLotNoTo").CmpStr = "TdLotNo<= "
  ADC("UoLotNoSel").CmpStr = "TdLotNo In "
  
  wrepcnd = ADC.RepCond
  'wCnd = IIF(wrepcnd <> "", " and " + wrepcnd, "")
   wCnd = IIF(wrepcnd <> "", " and TdAck='Y' and " + wrepcnd, " and TdAck='Y' ")
  '**************urmi 2.04  Reverse Logic ****************
  wRLQty = "(Case TdToRmDc When 'D' Then TdRmQty When 'C' Then -TdRmQty Else 0 End) as qOpnQty, "
  wRLWt = "(Case TdToRmDc When 'D' Then TdRmWt When 'C' Then -TdRmWt Else 0 End) as qOpnWt, "
  
  If ADC("UoYn2") = "Y" Then
    wRLQty = "(Case TdToRmDc When 'D' Then -TdRmQty When 'C' Then TdRmQty Else 0 End) as qOpnQty, "
    wRLWt = "(Case TdToRmDc When 'D' Then -TdRmWt When 'C' Then TdRmWt Else 0 End) as qOpnWt, "
  End If
  '*******************************************************
  '3
  ' Zubin 212
  '*** jay 2.14 *** ('' as qDesc,)
  '****** Sachin 3.02 - Id fields in Joins
  ' ***** Manali 3.03 - 24/06/08 - Pd Mod - 'L' added
  ' ****** Manali 3.03 - 09/07/08 - RmSzDesc added and join with RmIdSz added
  '4.1.4
  'JewelSmith.2 ',' + SC.PDesc225 + ',' like '%,'+ RmSCtg +',%' changed as SC.PSCd = RmStnCls
  wSqlOpnTdTo = "Insert Into #TmpStkLed Select " + grpflds1 + "," + _
                "'O' as qFlag, RmCd, TdLotNo as qLotNo, TdRmSz as qRmSz, TdRmSz2 as qRmSz2, TdRmSz3 as qRmSz3, " + wRmSzDesc + " as qRmSzDesc, IsNull(RrRngSz, '') as RrRngSz, RmCtg, IsNull(SC.PSCd,'') As qStnCls, RmSCtg, " + _
                "TdRmStkRt as qStkRt, (Select PValue3 from Param where PTyp= 'RMSCTG' and PMCd= RmCtg and PSCd= RmSCtg) as qPtrYN, " + _
                "(Case when RmCtg in ('G', 'P', 'S', 'L', 'M') then " + wPureWtOrZWt + " when RmCtg= 'X' then " + wXPureWtOrZwt + " else 0 end)" + wPurityWt + " as qRmPurityWt, " + _
                "TCoCd as TdCoCd, '' as TdTc,'' as TdYy, '' as TdChr, 0 as TdNo, 0 as TdSr, 0 as TdSrNo, " + _
                "'' as TdBYy, '' as TdBChr, 0 as TdBNo, '' as qDt, '' as qDesc, LocPrntCd as qPrntCd, TdToRmLoc as qLoc, " + _
                wRLQty + _
                "0 as qDrQty, 0 as qCrQty, " + _
                wRLWt + _
                "0 as qDrWt, 0 as qCrWt, '' as qFrRmLoc, '' as qToRmLoc, " + wTdRate + _
                ", " + wQwChk + ", RSCtg.PValue1 as qStkbyRt " + _
                " From Txn NoLock " + _
                "   Join Txnd on TIdNo=TdTIdNo and TCoCd= TdCoCd and TTc= TdTc and TYy= TdYy and TChr= TdChr and TNo= TdNo " + IIF(gs_Partition = ctCurrPrtn, " and TPrtKey=TdPrtKey ", "") + _
                "   Join Loc Lc On TdCoCd= LocCoCd and TdToRmLoc= LocCd and LocHideYN='N' Join RmMst Rm On TdRmCd= RmCd " + IIF(gs_Partition = ctCurrPrtn, " and TdPrtKey=RmPrtKey ", "") + _
                "   Join Param RSCtg On RSCtg.PTyp='RMSCTG' And RSCtg.PMCd=Rm.RmCtg And RSCtg.PSCd=RmSCtg  " + _
                "   Left Outer Join Param SC on SC.PTyp= 'STNCLS' and SC.PmCd= RmCtg and SC.PSCd = RmStnCls " + _
                " Left Outer Join LotMst on LotMst.LotNo = TdLotNo " + _
                "   Left Outer Join RmRt RSZ on RrTcTyp= 'RSZ' and RrCmCtg= 'C' and RrCmCd= '" + ctSelfCmCd + "' and " + _
                "       RrCtg= RmCtg and RrSCtg= RmSCtg and RrSTWGrp= '' and RrCd= '' and RrDmCtg= '' and " + _
                "       RrLsCtg= '' and RrLabMCd= '' and RrFrLn<= TdRmSz and RrToLn>= TdRmSz " + _
                " Left Outer Join RmIdSz on RisRmCtg=RmCtg and RisRmId=RmId and RisSz1=TdRmSz " + _
                wDtCndOpnTd + wCnd + IIF(gs_Partition = ctCurrPrtn, " and TdPrtKey='" + ctCurrPrtn + "' ", "") + " And TdToRmLocTyp= '" + ADC("UOLOCTYPFR") + "'" + IIF(wAddlCnd <> "", " And " + wAddlCnd, "") + " Option (MaxDop 1)"
  If wb_ShowOpn And Not wb_MonStart And Not wb_MonEnd Then moCn.Execute wSqlOpnTdTo
  
  '****** Sachin 3.02 - New fields
  ' ****** Manali 3.4.1 - 02/12/08 - [UoDtFr, UoDtTo respectively replaced by wDtFr, wDtTo]
  wDtCndDet = " where TDt between '" + CStr(wDtFr) + "'" + _
              " And '" + CStr(wDtTo) + "' And TdDt between '" + CStr(wDtFr) + "'" + _
              " And '" + CStr(wDtTo) + "'"
  'wDtCndDet = " where TDt between '" + CStr(adc("UoDtFr")) + "'" + _
              " And '" + CStr(adc("UoDtTo")) + "' And TdDt between '" + CStr(adc("UoDtFr")) + "'" + _
              " And '" + CStr(adc("UoDtTo")) + "'"
  
  ADC("UoLn1Fr").CmpStr = "TdRmSz>= "
  ADC("UoLn1To").CmpStr = "TdRmSz<= "
  ADC("UoCstRtFr").CmpStr = "TdRmStkRt>= "
  ADC("UoCstRtTo").CmpStr = "TdRmStkRt<= "
  ' Zubin 212 (This CmpStr block is not req as it is the same for the above and below query)
  ADC("UoLotNoFr").CmpStr = "TdLotNo>= "
  ADC("UoLotNoTo").CmpStr = "TdLotNo<= "
  ADC("UoLotNoSel").CmpStr = "TdLotNo In "
  
  wrepcnd = ADC.RepCond
'  wCnd = IIF(wrepcnd <> "", " and " + wrepcnd, "")
  wCnd = IIF(wrepcnd <> "", " and TdAck='Y' and " + wrepcnd, " and TdAck='Y' ")
  '4
  ' Zubin 212
  '*** jay 2.14 *** (" + qDesc + " as qDesc, added)
  ' ***** Manali 3.03 - 24/06/08 - Pd Mod - 'L' added
  ' ****** Manali 3.03 - 09/07/08 - RmSzDesc added and join with RmIdSz added
  '4.1.4
  'JewelSmith.2 ',' + SC.PDesc225 + ',' like '%,'+ RmSCtg +',%' changed as SC.PSCd = RmStnCls
  wSqlDetFr = " Insert Into #TmpStkLed Select " + grpflds1 + "," + _
              " 'T' as qFlag, RmCd, TdLotNo as qLotNo, TdRmSz as qRmSz, TdRmSz2 as qRmSz2, TdRmSz3 as qRmSz3, " + wRmSzDesc + " as qRmSzDesc,  IsNull(RrRngSz, '') as RrRngSz, RmCtg, IsNull(SC.PSCd,'') As qStnCls, RmSCtg, " + _
              " TdRmStkRt as qStkRt, (Select PValue3 from Param where PTyp= 'RMSCTG' and PMCd= RmCtg and PSCd= RmSCtg) as qPtrYN, " + _
              " (Case when RmCtg in ('G', 'P', 'S', 'L', 'M') then " + wPureWtOrZWt + " when RmCtg= 'X' then " + wXPureWtOrZwt + " else 0 end)" + wPurityWt + " as qRmPurityWt, " + _
              " TdCoCd, TdTc, TdYy, TdChr, TdNo, TdSr, TdSrNo, TdBYy, TdBChr, TdBNo, " + _
              " TDt as qDt, " + qDesc + " as qDesc, LocPrntCd as qPrntCd, TdFrRmLoc as qLoc, 0 as qOpnQty, " + _
              " Case TdFrRmDc When 'D' Then TdRmQty Else 0 End  as qDrQty, " + _
              " Case TdFrRmDc When 'C' Then TdRmQty else 0 End  as qCrQty, " + _
              " 0 as qOpnWt, Case TdFrRmDc When 'D' Then TdRmWt Else 0 End  as qDrWt, " + _
              " Case TdFrRmDc When 'C' Then TdRmWt else 0 End  as qCrWt, " + _
              " (case when TdFrRmDc= 'C' then TdFrRmLoc else TdToRmLoc end) as qFrRmLoc, " + _
              " (case when TdFrRmDc= 'D' then TdFrRmLoc else TdToRmLoc end) as qToRmLoc, " + wTdRate + _
              " , " + wQwChk + ", RSCtg.PValue1 as qStkbyRt " + _
              " From Txn NoLock join Txnd on TIdNo=TdTIdNo and TCoCd= TdCoCd and TTc= TdTc and TYy= TdYy and TChr= TdChr and TNo= TdNo " + IIF(gs_Partition = ctCurrPrtn, " and TPrtKey=TdPrtKey ", "") + _
              " Join Loc Lc On LocCoCd= TdCoCd and LocCd= TdFrRmLoc and LocHideYN='N' " + _
              " Join RmMst Rm On TdRmCd= RmCd " + IIF(gs_Partition = ctCurrPrtn, " and RmPrtKey=TdPrtKey ", "") + _
              "   Join Param RSCtg On RSCtg.PTyp='RMSCTG' And RSCtg.PMCd=Rm.RmCtg And RSCtg.PSCd=Rm.RmSCtg  " + _
              "   Left Outer Join Param SC on SC.PTyp= 'STNCLS' and SC.PmCd= RmCtg and SC.PSCd = RmStnCls " + _
              " Left Outer Join LotMst on LotMst.LotNo = TdLotNo " + _
              "   Left Outer Join RmRt RSZ on RrTcTyp= 'RSZ' and RrCmCtg= 'C' and RrCmCd= '" + ctSelfCmCd + "' and " + _
              " RrCtg= RmCtg and RrSCtg= RmSCtg and RrSTWGrp= '' and RrCd= '' and RrDmCtg= '' and " + _
              " RrLsCtg= '' and RrLabMCd= '' and RrFrLn<= TdRmSz and RrToLn>= TdRmSz " + _
              " Left Outer Join RmIdSz on RisRmCtg=RmCtg and RisRmId=RmId and RisSz1=TdRmSz " + _
              wDtCndDet + wCnd + IIF(gs_Partition = ctCurrPrtn, " and TPrtKey='" + ctCurrPrtn + "' ", "") + " And TdFrRmLocTyp= '" + ADC("UOLOCTYPFR") + "'" + IIF(wAddlCnd <> "", " And " + wAddlCnd, "") + " Option (MaxDop 1)"
    moCn.Execute wSqlDetFr

  ADC("UoLn1Fr").CmpStr = "TdRmSz>= "
  ADC("UoLn1To").CmpStr = "TdRmSz<= "
  ADC("UoCstRtFr").CmpStr = "TdRmStkRt>= "
  ADC("UoCstRtTo").CmpStr = "TdRmStkRt<= "
  ' Zubin 212 (This CmpStr block is not req as it is the same for the above and below query)
  ADC("UoLotNoFr").CmpStr = "TdLotNo>= "
  ADC("UoLotNoTo").CmpStr = "TdLotNo<= "
  ADC("UoLotNoSel").CmpStr = "TdLotNo In "
  
  wrepcnd = ADC.RepCond
  'wCnd = IIF(wrepcnd <> "", " and " + wrepcnd, "")
    wCnd = IIF(wrepcnd <> "", " and TdAck='Y' and " + wrepcnd, " and TdAck='Y' ")
  '5
  ' Zubin 212
  ' *** jay 2.14 ***(" + qDesc + " as qDesc, added)
  ' ***** Manali 3.03 - 24/06/08 - Pd Mod - 'L' added
  ' ****** Manali 3.03 - 09/07/08 - RmSzDesc added and join with RmIdSz added
  ' 4.1.4
  'JewelSmith.2 ',' + SC.PDesc225 + ',' like '%,'+ RmSCtg +',%' changed as SC.PSCd = RmStnCls
  wSqlDetTo = " Insert Into #TmpStkLed Select " + grpflds1 + "," + _
              " 'T' as qFlag, RmCd, TdLotNo as qLotNo, TdRmSz as qRmSz, TdRmSz2 as qRmSz2, TdRmSz3 as qRmSz3, " + wRmSzDesc + " as qRmSzDesc, IsNull(RrRngSz, '') as RrRngSz, RmCtg, IsNull(SC.PSCd,'') As qStnCls, RmSCtg, " + _
              " TdRmStkRt as qStkRt, (Select PValue3 from Param where PTyp= 'RMSCTG' and PMCd= RmCtg and PSCd= RmSCtg) as qPtrYN, " + _
              " (Case when RmCtg in ('G', 'P', 'S', 'L', 'M') then " + wPureWtOrZWt + " when RmCtg= 'X' then " + wXPureWtOrZwt + " else 0 end)" + wPurityWt + " as qRmPurityWt, " + _
              " TdCoCd, TdTc, TdYy, TdChr, TdNo, TdSr, TdSrNo, TdBYy, TdBChr, TdBNo, " + _
              " TDt as qDt, " + qDesc + " as qDesc, LocPrntCd as qPrntCd, TdToRmLoc as qLoc, 0 as qOpnQty, " + _
              " Case TdToRmDc When 'D' Then TdRmQty Else 0 End  as qDrQty, " + _
              " Case TdToRmDc When 'C' Then TdRmQty else 0 End  as qCrQty, " + _
              " 0 as qOpnWt, Case TdToRmDc When 'D' Then TdRmWt Else 0 End  as qDrWt, " + _
              " Case TdToRmDc When 'C' Then TdRmWt else 0 End  as qCrWt, " + _
              " (case when TdFrRmDc= 'C' then TdFrRmLoc else TdToRmLoc end) as qFrRmLoc, " + _
              " (case when TdFrRmDc= 'D' then TdFrRmLoc else TdToRmLoc end) as qToRmLoc, " + wTdRate + _
              " , " + wQwChk + ", RSCtg.PValue1 as qStkbyRt " + _
              " From Txn NoLock join Txnd on TIdNo=TdTIdNo and TCoCd= TdCoCd and TTc= TdTc and TYy= TdYy and TChr= TdChr and TNo= TdNo " + IIF(gs_Partition = ctCurrPrtn, " and TPrtKey=TdPrtKey ", "") + _
              " Join Loc Lc On LocCoCd= TdCoCd and LocCd= TdToRmLoc and LocHideYN='N' Join RmMst Rm On TdRmCd= RmCd " + IIF(gs_Partition = ctCurrPrtn, " and TdPrtKey=RmPrtKey ", "") + _
              "   Join Param RSCtg On RSCtg.PTyp='RMSCTG' And RSCtg.PMCd=Rm.RmCtg And RSCtg.PSCd=Rm.RmSCtg  " + _
              " Left Outer Join Param SC on SC.PTyp= 'STNCLS' and SC.PmCd= RmCtg and SC.PSCd = RmStnCls " + _
              " Left Outer Join LotMst on LotMst.LotNo = TdLotNo " + _
              " Left Outer Join RmRt RSZ on RrTcTyp= 'RSZ' and RrCmCtg= 'C' and RrCmCd= '" + ctSelfCmCd + "' and " + _
              " RrCtg= RmCtg and RrSCtg= RmSCtg and RrSTWGrp= '' and RrCd= '' and RrDmCtg= '' and " + _
              " RrLsCtg= '' and RrLabMCd= '' and RrFrLn<= TdRmSz and RrToLn>= TdRmSz " + _
              " Left Outer Join RmIdSz on RisRmCtg=RmCtg and RisRmId=RmId and RisSz1=TdRmSz " + _
              wDtCndDet + wCnd + IIF(gs_Partition = ctCurrPrtn, " and TPrtKey='" + ctCurrPrtn + "' ", "") + " And TdToRmLocTyp= '" + ADC("UOLOCTYPFR") + "'" + IIF(wAddlCnd <> "", " And " + wAddlCnd, "") + " Option (MaxDop 1)"
  moCn.Execute wSqlDetTo

  '*** If Location 'W' is selected, an additional Unions FgRm FrLoc, FgRm ToLoc is required.
  wSqlOpnFgFr = "": wSqlOpnFgTo = "": wSqlDetFgFr = "": wSqlDetFgTo = ""
  ' ****** Manali 3.4.1 - 02/12/08 - [UoDtFr, UoDtTo respectively replaced by wDtFr, wDtTo]
  wFrDt1 = IIF(UCase(MWLib.DtSetting) = "MM/DD/YY", Format(wDtFr, "mm") + "/01", "01/" + Format(wDtFr, "mm")) + "/" + Format(wDtFr, "yy")   '6.1
  'wFrDt1 = "01/" + Format(adc("UoDtFr"), "mm") + "/" + Format(adc("UoDtFr"), "yy")
  
  '****** SAchin 3.02.0 - New fields in Where Clause
  ' ****** Manali 3.4.1 - 02/12/08 - [UoDtFr, UoDtTo respectively replaced by wDtFr, wDtTo]
  wDtCndOpnFg = " where FgDt >= '" + wFrDt1 + "' And FgDt < '" + CStr(wDtFr) + "' And FdDt >= '" + wFrDt1 + "' And FdDt < '" + CStr(wDtFr) + "' "
  'wDtCndOpnFg = " where FgDt >= '" + wFrDt1 + "' And FgDt < '" + CStr(adc("UoDtFr")) + "' And FdDt >= '" + wFrDt1 + "' And FdDt < '" + CStr(adc("UoDtFr")) + "' "
  
  '************* Reverse Logic urmi 5/7/03 *************************************
  wRLQty = "(Case FrFrRmDc When 'D' Then FrRmQty When 'C' Then -FrRmQty Else 0 End) as qOpnQty, "
  wRLWt = "(Case FrFrRmDc When 'D' Then FrRmWt When 'C' Then -FrRmWt Else 0 End) as qOpnWt, "
  
  
  If ADC("Uoyn2") = "Y" Then
'    wDtCndOpnFg = " where FgDt >= '" & ADC("UoDtFr") & "' And month(FgDt) = " & _
'                  CInt(Format(ADC("UoDtFr"), "mm")) & " and Year(FgDt) = " & _
'                  CInt(Format(ADC("UoDtFr"), "yyyy"))

    '****** Sachin 3.02.0 - New fields in Where Clause
    ' ****** Manali 3.4.1 - 02/12/08 - [UoDtFr, UoDtTo respectively replaced by wDtFr, wDtTo]
    wDtCndOpnFg = " where FgDt >= '" & wDtFr & "' And FgDt <= '" & wTmpDt & "' And  FdDt >= '" & wDtFr & "' And FdDt <= '" & wTmpDt & "' "
    'wDtCndOpnFg = " where FgDt >= '" & adc("UoDtFr") & "' And FgDt <= '" & wTmpDt & "' And  FdDt >= '" & adc("UoDtFr") & "' And FdDt <= '" & wTmpDt & "' "
    
    
    wRLQty = "(Case FrFrRmDc When 'D' Then -FrRmQty When 'C' Then FrRmQty Else 0 End) as qOpnQty, "
    wRLWt = "(Case FrFrRmDc When 'D' Then -FrRmWt When 'C' Then FrRmWt Else 0 End) as qOpnWt, "
    
  End If
  '*******************************************************************************
  
  ADC("UoLn1Fr").CmpStr = "FrRmSz>= "
  ADC("UoLn1To").CmpStr = "FrRmSz<= "
  ADC("UoCstRtFr").CmpStr = "FrRmStkRt>= "
  ADC("UoCstRtTo").CmpStr = "FrRmStkRt<= "
  ' Zubin 212
  ADC("UoLotNoFr").CmpStr = "FrLotNo>= "
  ADC("UoLotNoTo").CmpStr = "FrLotNo<= "
  ADC("UoLotNoSel").CmpStr = "FrLotNo In "
  wrepcnd = ADC.RepCond
  wCnd = IIF(wrepcnd <> "", " and " + wrepcnd, "")
  
  '****** Sachin 3.02.0 - Additional Conditions For New Fields
  wAddlCnd = ""
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(ADC("UORMCTGFR")) <> "", " And ", "") + IIF(Trim(ADC("UORMCTGFR")) <> "", " FrRmCtg <= '" + Trim(ADC("UORMCTGFR")) + "'", "")
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(ADC("UORMSCTGFR")) <> "", " And ", "") + IIF(Trim(ADC("UORMSCTGFR")) <> "", " FrRmSCtg >= '" + Trim(ADC("UORMSCTGFR")) + "'", "")
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(ADC("UORMSCTGTO")) <> "", " And ", "") + IIF(Trim(ADC("UORMSCTGTO")) <> "", " FrRmSCtg <= '" + Trim(ADC("UORMSCTGTO")) + "'", "")
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(ADC("UORMSCTGSEL")) <> "", " And ", "") + IIF(Trim(ADC("UORMSCTGSEL")) <> "", " FrRmSCtg In (" + Trim(ADC("UORMSCTGSEL")) + ")", "")
  '****** Sachin 3.02.0 - Additional Conditions For New Fields

  '6
  ' Zubin 212
  '*** jay 2.14 *** ('' as qDesc,)
  ' ***** Manali 3.03 - 24/06/08 - Pd Mod - 'L' added
  ' ****** Manali 3.03 - 09/07/08 - RmSzDesc added and join with RmIdSz added
  '4.1.4
  'JewelSmith.2 ',' + SC.PDesc225 + ',' like '%,'+ RmSCtg +',%' changed as SC.PSCd = RmStnCls
  wSqlOpnFgFr = " Insert Into #TmpStkLed Select " + grpflds2 + "," + _
                " 'O' as qFlag, RmCd, FrLotNo as qLotNo, FrRmSz as qRmSz, FrRmSz2 as qRmSz2, FrRmSz3 as qRmSz3, " + wRmSzDesc + " as qRmSzDesc, IsNull(RrRngSz, '') as RrRngSz, RmCtg, IsNull(SC.PSCd,'') As qStnCls, RmSCtg, " + _
                " FrRmStkRt as qStkRt, (Select PValue3 from Param where PTyp= 'RMSCTG' and PMCd= RmCtg and PSCd= RmSCtg) as qPtrYN, " + _
                " (Case when RmCtg in ('G', 'P', 'S', 'L', 'M') then " + wPureWtOrZWt + " when RmCtg= 'X' then " + wXPureWtOrZwt + " else 0 end)" + wPurityWt + " as qRmPurityWt, " + _
                " FgCoCd as TdCoCd, '' as TdTc, '' as TdYy, '' as TdChr, 0 as TdNo, 0 as TdSr, " + _
                " 0 as TdSrNo, '' as TdBYy, '' as TdBChr, 0 as TdBNo, '' as qDt, '' as qDesc, " + _
                " LocPrntCd as qPrntCd, FrFrRmLoc as qLoc, " + _
                wRLQty + _
                " 0 as qDrQty, 0 as qCrQty, " + _
                wRLWt + _
                " 0 as qDrWt, 0 as qCrWt, '' as qFrRmLoc, '' as qToRmLoc, " + wFrRate + _
                " , " + wQwChk + ", RSCtg.PValue1 as qStkbyRt " + _
                " From Fg NoLock " + _
                " join Fgd on FgIdNo=FdFgIdNo and FgCoCd= FdCoCd and FgTc= FdTc and FgYy= FdYy and FgChr= FdChr and FgNo= FdNo " + IIF(gs_Partition = ctCurrPrtn, " and FgPrtKey=FdPrtKey ", "") + _
                " join FgRm on FdIdNo=FrFdIdNo and FdCoCd= FrCoCd and FdTc= FrTc and FdYy= FrYy and " + _
                " FdChr= FrChr and FdNo= FrNo and FdSr= FrSr " + IIF(gs_Partition = ctCurrPrtn, " and FdPrtKey=FrPrtKey ", "") + _
                " Join Loc Lc On FrCoCd= LocCoCd and FrFrRmLoc= LocCd and LocHideYN='N' Join RmMst Rm On FrRmCd= RmCd " + IIF(gs_Partition = ctCurrPrtn, " and FrPrtKey=RmPrtKey ", "") + _
                "   Join Param RSCtg On RSCtg.PTyp='RMSCTG' And RSCtg.PMCd=Rm.RmCtg And RSCtg.PSCd=Rm.RmSCtg  " + _
                " Left Outer Join Param SC on SC.PTyp= 'STNCLS' and SC.PmCd= RmCtg and SC.PSCd = RmStnCls " + _
                " Left Outer Join LotMst on LotMst.LotNo = FrLotNo " + _
                " Left Outer Join RmRt RSZ on RrTcTyp= 'RSZ' and RrCmCtg= 'C' and RrCmCd= '" + ctSelfCmCd + "' and " + _
                " RrCtg= RmCtg and RrSCtg= RmSCtg and RrSTWGrp= '' and RrCd= '' and RrDmCtg= '' and " + _
                " RrLsCtg= '' and RrLabMCd= '' and RrFrLn<= FrRmSz and RrToLn>= FrRmSz " + _
                " Left Outer Join RmIdSz on RisRmCtg=RmCtg and RisRmId=RmId and RisSz1=FrRmSz " + _
                wDtCndOpnFg + wCnd + IIF(gs_Partition = ctCurrPrtn, " and FgPrtKey='" + ctCurrPrtn + "' ", "") + " And FrFrRmLocTyp= '" + ADC("UOLOCTYPFR") + "'" + IIF(wAddlCnd <> "", " And " + wAddlCnd, "") + " Option (MaxDop 1)"
                '"Join Loc On FrFrRmLoc= LocCd where FrFrRmLoc= 'W' " + wDtCndOpnFg + wCnd
  If wb_ShowOpn And Not wb_MonStart And Not wb_MonEnd Then moCn.Execute wSqlOpnFgFr

  ADC("UoLn1Fr").CmpStr = "FrRmSz>= "
  ADC("UoLn1To").CmpStr = "FrRmSz<= "
  ADC("UoCstRtFr").CmpStr = "FrRmStkRt>= "
  ADC("UoCstRtTo").CmpStr = "FrRmStkRt<= "
  ' Zubin 212 (This CmpStr block is not req as it is the same for the above and below query)
  ADC("UoLotNoFr").CmpStr = "FrLotNo>= "
  ADC("UoLotNoTo").CmpStr = "FrLotNo<= "
  ADC("UoLotNoSel").CmpStr = "FrLotNo In "
  
  wrepcnd = ADC.RepCond
  wCnd = IIF(wrepcnd <> "", " and " + wrepcnd, "")
  
   '**************urmi 2.04  Reverse Logic ****************
  wRLQty = "(Case FrToRmDc When 'D' Then FrRmQty When 'C' Then -FrRmQty Else 0 End) as qOpnQty, "
  wRLWt = "(Case FrToRmDc When 'D' Then FrRmWt When 'C' Then -FrRmWt Else 0 End) as qOpnWt, "
  
  If ADC("UoYn2") = "Y" Then
    wRLQty = "(Case FrToRmDc When 'D' Then -FrRmQty When 'C' Then FrRmQty Else 0 End) as qOpnQty, "
    wRLWt = "(Case FrToRmDc When 'D' Then -FrRmWt When 'C' Then FrRmWt Else 0 End) as qOpnWt, "
  End If
  '*******************************************************
  '7
  '*** jay 2.14 *** ('' as qDesc,)
  '****** Sachin 3.02 - Id fields in Joins
  ' ***** Manali 3.03 - 24/06/08 - Pd Mod - 'L' added
  ' ****** Manali 3.03 - 09/07/08 - RmSzDesc added and join with RmIdSz added
  '4.1.4
  'JewelSmith.2 ',' + SC.PDesc225 + ',' like '%,'+ RmSCtg +',%' changed as SC.PSCd = RmStnCls
  wSqlOpnFgTo = "Insert Into #TmpStkLed Select " + grpflds2 + ", " + _
                "'O' as qFlag, RmCd, FrLotNo as qLotNo, FrRmSz as qRmSz, FrRmSz2 as qRmSz2, FrRmSz3 as qRmSz3, " + wRmSzDesc + " as qRmSzDesc, IsNull(RrRngSz, '') as RrRngSz, RmCtg, IsNull(SC.PSCd,'') As qStnCls, RmSCtg, " + _
                "FrRmStkRt as qStkRt, (Select PValue3 from Param where PTyp= 'RMSCTG' and PMCd= RmCtg and PSCd= RmSCtg) as qPtrYN, " + _
                "(Case when RmCtg in ('G', 'P', 'S', 'L', 'M') then " + wPureWtOrZWt + " when RmCtg= 'X' then " + wXPureWtOrZwt + " else 0 end)" + wPurityWt + " as qRmPurityWt, " + _
                "FgCoCd as TdCoCd, '' as TdTc, '' as TdYy, '' as TdChr, 0 as TdNo, 0 as TdSr, " + _
                "0 as TdSrNo, '' as TdBYy, '' as TdBChr, 0 as TdBNo, '' as qDt, '' as qDesc, " + _
                "LocPrntCd as qPrntCd, FrToRmLoc as qLoc, " + _
                wRLQty + _
                "0 as qDrQty, 0 as qCrQty, " + _
                wRLWt + _
                "0 as qDrWt, 0 as qCrWt, '' as qFrRmLoc, '' as qToRmLoc, " + wFrRate + _
                ", " + wQwChk + ", RSCtg.PValue1 as qStkbyRt " + _
                "From Fg NoLock " + _
                "join Fgd on FgIdNo=FdFgIdNo and FgCoCd= FdCoCd and FgTc= FdTc and FgYy= FdYy and FgChr= FdChr and FgNo= FdNo " + IIF(gs_Partition = ctCurrPrtn, " and FgPrtKey=FdPrtKey ", "") + _
                "join FgRm on FdIdNo=FrFdIdNo and FdCoCd= FrCoCd and FdTc= FrTc and FdYy= FrYy and " + _
                "FdChr= FrChr and FdNo= FrNo and FdSr= FrSr " + IIF(gs_Partition = ctCurrPrtn, " and FdPrtKey=FrPrtKey ", "") + _
                "Join Loc Lc On FrCoCd= LocCoCd and FrToRmLoc= LocCd and LocHideYN='N' Join RmMst Rm On FrRmCd= RmCd " + IIF(gs_Partition = ctCurrPrtn, " and RmPrtKey=FrPrtKey ", "") + _
                "   Join Param RSCtg On RSCtg.PTyp='RMSCTG' And RSCtg.PMCd=Rm.RmCtg And RSCtg.PSCd=Rm.RmSCtg  " + _
                "Left Outer Join Param SC on SC.PTyp= 'STNCLS' and SC.PmCd= RmCtg and SC.PSCd = RmStnCls " + _
                " Left Outer Join LotMst on LotMst.LotNo = FrLotNo " + _
                "Left Outer Join RmRt RSZ on RrTcTyp= 'RSZ' and RrCmCtg= 'C' and RrCmCd= '" + ctSelfCmCd + "' and " + _
                "RrCtg= RmCtg and RrSCtg= RmSCtg and RrSTWGrp= '' and RrCd= '' and RrDmCtg= '' and " + _
                "RrLsCtg= '' and RrLabMCd= '' and RrFrLn<= FrRmSz and RrToLn>= FrRmSz " + _
                "Left Outer Join RmIdSz on RisRmCtg=RmCtg and RisRmId=RmId and RisSz1=FrRmSz " + _
                wDtCndOpnFg + wCnd + IIF(gs_Partition = ctCurrPrtn, " and FgPrtKey='" + ctCurrPrtn + "' ", "") + " And FrToRmLocTyp= '" + ADC("UOLOCTYPFR") + "'" + IIF(wAddlCnd <> "", " And " + wAddlCnd, "") + " Option (MaxDop 1)"
                '"Join Loc On FrToRmLoc= LocCd where FrToRmLoc= 'W' " + wDtCndOpnFg + wCnd
  If wb_ShowOpn And Not wb_MonStart And Not wb_MonEnd Then moCn.Execute wSqlOpnFgTo

  '****** Sachin 3.02.0 - New fields in Where Clause
  ' ****** Manali 3.4.1 - 02/12/08 - [UoDtFr, UoDtTo respectively replaced by wDtFr, wDtTo]
  wDtCndDetFg = " where FgDt between '" + CStr(wDtFr) + "'" + _
              " And '" + CStr(wDtTo) + "' And FdDt between '" + CStr(wDtFr) + "'" + _
              " And '" + CStr(wDtTo) + "'"
'  wDtCndDetFg = " where FgDt between '" + CStr(adc("UoDtFr")) + "'" + _
'              " And '" + CStr(adc("UoDtTo")) + "' And FdDt between '" + CStr(adc("UoDtFr")) + "'" + _
'              " And '" + CStr(adc("UoDtTo")) + "'"

    
  ADC("UoLn1Fr").CmpStr = "FrRmSz>= "
  ADC("UoLn1To").CmpStr = "FrRmSz<= "
  ADC("UoCstRtFr").CmpStr = "FrRmStkRt>= "
  ADC("UoCstRtTo").CmpStr = "FrRmStkRt<= "
  ' Zubin 212 (This CmpStr block is not req as it is the same for the above and below query)
  ADC("UoLotNoFr").CmpStr = "FrLotNo>= "
  ADC("UoLotNoTo").CmpStr = "FrLotNo<= "
  ADC("UoLotNoSel").CmpStr = "FrLotNo In "
  wrepcnd = ADC.RepCond
  wCnd = IIF(wrepcnd <> "", " and " + wrepcnd, "")
  '8
  ' Zubin 212
  '*** jay 2.14 *** ('' as qDesc,)
  ' ***** Manali 3.03 - 24/06/08 - Pd Mod - 'L' added
  ' ****** Manali 3.03 - 09/07/08 - RmSzDesc added and join with RmIdSz added
  ' ***** Manali Trading Module - Bag Join added
  '4.1.4
  'JewelSmith.2 ',' + SC.PDesc225 + ',' like '%,'+ RmSCtg +',%' changed as SC.PSCd = RmStnCls
  wSqlDetFgFr = " Insert Into #TmpStkLed Select " + grpflds2 + "," + _
              " 'T' as qFlag, RmCd, FrLotNo as qLotNo, FrRmSz as qRmSz, FrRmSz2 as qRmSz2, FrRmSz3 as qRmSz3, " + wRmSzDesc + " as qRmSzDesc, IsNull(RrRngSz, '') as RrRngSz, RmCtg, IsNull(SC.PSCd,'') As qStnCls, RmSCtg, " + _
              " FrRmStkRt as qStkRt, (Select PValue3 from Param where PTyp= 'RMSCTG' and PMCd= RmCtg and PSCd= RmSCtg) as qPtrYN, " + _
              " (Case when RmCtg in ('G', 'P', 'S', 'L', 'M') then " + wPureWtOrZWt + " when RmCtg= 'X' then " + wXPureWtOrZwt + " else 0 end)" + wPurityWt + " as qRmPurityWt, " + _
              " FgCoCd as TdCoCd, FrTc as TdTc, FrYy as TdYy, FrChr as TdChr, FrNo as TdNo, FrSr as TdSr, FrSrNo as TdSrNo, " + _
              " FdBYy as TdBYy, FdBChr as TdBChr, FdBNo as TdBNo, FgDt as qDt, '' as qDesc, " + _
              " LocPrntCd as qPrntCd, FrFrRmLoc as qLoc, 0 as qOpnQty, " + _
              " (Case FrFrRmDc When 'D' Then FrRmQty Else 0 End) as qDrQty, " + _
              " (Case FrFrRmDc When 'C' Then FrRmQty else 0 End) as qCrQty, 0 as qOpnWt, " + _
              " (Case FrFrRmDc When 'D' Then FrRmWt Else 0 End) as qDrWt, " + _
              " (Case FrFrRmDc When 'C' Then FrRmWt else 0 End) as qCrWt, " + _
              " (case when FrFrRmDc= 'C' then FrFrRmLoc else FrToRmLoc end) as qFrRmLoc, " + _
              " (case when FrFrRmDc= 'D' then FrFrRmLoc else FrToRmLoc end) as qToRmLoc, " + wFrRate + ", " + wQwChk + ", RSCtg.PValue1 as qStkbyRt " + _
              " From Fg NoLock " + _
              " join Fgd on FgIdNo=FdFgIdNo and FgCoCd= FdCoCd and FgTc= FdTc and FgYy= FdYy and FgChr= FdChr and FgNo= FdNo " + IIF(gs_Partition = ctCurrPrtn, " and FgPrtKey=FdPrtKey ", "") + _
              " join FgRm on FdIdNo=FrFdIdNo and FdCoCd= FrCoCd and FdTc= FrTc and FdYy= FrYy and " + _
              " FdChr= FrChr and FdNo= FrNo and FdSr= FrSr " + IIF(gs_Partition = ctCurrPrtn, " and FdPrtKey=FrPrtKey ", "") + _
              " Join Loc Lc On LocCoCd= FrCoCd and LocCd= FrFrRmLoc and LocHideYN='N' Join RmMst Rm On FrRmCd= RmCd " + IIF(gs_Partition = ctCurrPrtn, " and RmPrtKey=FrPrtKey ", "") + _
              "    Join Param RSCtg On RSCtg.PTyp='RMSCTG' And RSCtg.PMCd=Rm.RmCtg And RSCtg.PSCd=Rm.RmSCtg  " + _
              " Left Outer Join Param SC on SC.PTyp= 'STNCLS' and SC.PmCd= RmCtg and SC.PSCd = RmStnCls " + _
              " Left Outer Join LotMst on LotMst.LotNo = FrLotNo  Left Outer Join RmRt RSZ on RrTcTyp= 'RSZ' and RrCmCtg= 'C' and RrCmCd= '" + ctSelfCmCd + "' and " + _
              " RrCtg= RmCtg and RrSCtg= RmSCtg and RrSTWGrp= '' and RrCd= '' and RrDmCtg= '' and " + _
              " RrLsCtg= '' and RrLabMCd= '' and RrFrLn<= FrRmSz and RrToLn>= FrRmSz " + _
              " Left Outer Join RmIdSz on RisRmCtg=RmCtg and RisRmId=RmId and RisSz1=FrRmSz " + _
              wDtCndDetFg + wCnd + IIF(gs_Partition = ctCurrPrtn, " and FgPrtKey='" + ctCurrPrtn + "' ", "") + " And FrFrRmLocTyp= '" + ADC("UOLOCTYPFR") + "'" + IIF(wAddlCnd <> "", " And " + wAddlCnd, "") + " Option (MaxDop 1)"
              '"where FrFrRmLoc= 'W' " + wDtCndDetFg + wCnd
              
  moCn.Execute wSqlDetFgFr

  ADC("UoLn1Fr").CmpStr = "FrRmSz>= "
  ADC("UoLn1To").CmpStr = "FrRmSz<= "
  ADC("UoCstRtFr").CmpStr = "FrRmStkRt>= "
  ADC("UoCstRtTo").CmpStr = "FrRmStkRt<= "
  ' Zubin 212 (This CmpStr block is not req as it is the same for the above and below query)
  ADC("UoLotNoFr").CmpStr = "FrLotNo>= "
  ADC("UoLotNoTo").CmpStr = "FrLotNo<= "
  ADC("UoLotNoSel").CmpStr = "FrLotNo In "
  
  wrepcnd = ADC.RepCond
  wCnd = IIF(wrepcnd <> "", " and " + wrepcnd, "")
  
  '9
  ' Zubin 212
  '*** jay 2.14 *** ('' as qDesc, )
  ' ***** Manali 3.03 - 24/06/08 - Pd Mod - 'L' added
  ' ****** Manali 3.03 - 09/07/08 - RmSzDesc added and join with RmIdSz added
  ' ***** Manali Trading Module - Bag Join added
  '4.1.4
  'JewelSmith.2 ',' + SC.PDesc225 + ',' like '%,'+ RmSCtg +',%' changed as SC.PSCd = RmStnCls
  wSqlDetFgTo = "Insert Into #TmpStkLed Select " + grpflds2 + "," + _
              " 'T' as qFlag, RmCd, FrLotNo as qLotNo, FrRmSz as qRmSz, FrRmSz2 as qRmSz2, FrRmSz3 as qRmSz3, " + wRmSzDesc + " as qRmSzDesc, IsNull(RrRngSz, '') as RrRngSz, RmCtg, IsNull(SC.PSCd,'') As qStnCls, RmSCtg, " + _
              " FrRmStkRt as qStkRt, (Select PValue3 from Param where PTyp= 'RMSCTG' and PMCd= RmCtg and PSCd= RmSCtg) as qPtrYN, " + _
              " (Case when RmCtg in ('G', 'P', 'S', 'L', 'M') then " + wPureWtOrZWt + " when RmCtg= 'X' then " + wXPureWtOrZwt + " else 0 end)" + wPurityWt + " as qRmPurityWt, " + _
              " FgCoCd as TdCoCd, FrTc as TdTc, FrYy as TdYy, FrChr as TdChr, FrNo as TdNo, FrSr as TdSr, FrSrNo as TdSrNo, " + _
              " FdBYy as TdBYy, FdBChr as TdBChr, FdBNo as TdBNo, FgDt as qDt, '' as qDesc, " + _
              " LocPrntCd as qPrntCd, FrToRmLoc as qLoc, 0 as qOpnQty, " + _
              " (Case FrToRmDc When 'D' Then FrRmQty Else 0 End) as qDrQty, " + _
              " (Case FrToRmDc When 'C' Then FrRmQty else 0 End) as qCrQty, 0 as qOpnWt, " + _
              " (Case FrToRmDc When 'D' Then FrRmWt Else 0 End) as qDrWt, " + _
              " (Case FrToRmDc When 'C' Then FrRmWt else 0 End) as qCrWt, " + _
              " (case when FrFrRmDc= 'C' then FrFrRmLoc else FrToRmLoc end) as qFrRmLoc, " + _
              " (case when FrFrRmDc= 'D' then FrFrRmLoc else FrToRmLoc end) as qToRmLoc, " + wFrRate + ", " + wQwChk + ", RSCtg.PValue1 as qStkbyRt " + _
              " From Fg NoLock " + _
              " join Fgd on FgIdNo=FdFgIdNo and FgCoCd= FdCoCd and FgTc= FdTc and FgYy= FdYy and FgChr= FdChr and FgNo= FdNo " + IIF(gs_Partition = ctCurrPrtn, " and FgPrtKey=FdPrtKey ", "") + _
              " join FgRm on FdIdNo=FrFdIdNo and FdCoCd= FrCoCd and FdTc= FrTc and FdYy= FrYy and " + _
              "     FdChr= FrChr and FdNo= FrNo and FdSr= FrSr " + IIF(gs_Partition = ctCurrPrtn, " and FdPrtKey=FrPrtKey ", "") + _
              " Join Loc Lc On LocCoCd= FrCoCd and LocCd= FrToRmLoc and LocHideYN='N' Join RmMst Rm On FrRmCd= RmCd " + IIF(gs_Partition = ctCurrPrtn, " and FrPrtKey=RmPrtKey ", "") + _
" Left Outer Join LotMst on LotMst.LotNo = FrLotNo  Join Param RSCtg On RSCtg.PTyp='RMSCTG' And RSCtg.PMCd=Rm.RmCtg And RSCtg.PSCd=Rm.RmSCtg  " + _
              " Left Outer Join Param SC on SC.PTyp= 'STNCLS' and SC.PmCd= RmCtg and SC.PSCd = RmStnCls " + _
              " Left Outer Join RmRt RSZ on RrTcTyp= 'RSZ' and RrCmCtg= 'C' and RrCmCd= '" + ctSelfCmCd + "' and " + _
              "     RrCtg= RmCtg and RrSCtg= RmSCtg and RrSTWGrp= '' and RrCd= '' and RrDmCtg= '' and " + _
              "     RrLsCtg= '' and RrLabMCd= '' and RrFrLn<= FrRmSz and RrToLn>= FrRmSz " + _
              " Left Outer Join RmIdSz on RisRmCtg=RmCtg and RisRmId=RmId and RisSz1=FrRmSz " + _
              wDtCndDetFg + wCnd + IIF(gs_Partition = ctCurrPrtn, " and FgPrtKey='" + ctCurrPrtn + "' ", "") + " And FrToRmLocTyp= '" + ADC("UOLOCTYPFR") + "'" + IIF(wAddlCnd <> "", " And " + wAddlCnd, "") + " Option (MaxDop 1)"
              '"where FrToRmLoc= 'W' " + wDtCndDetFg + wCnd
  moCn.Execute wSqlDetFgTo

  'wFgStr = " Union All " + wSqlOpnFgFr + " Union All " + wSqlOpnFgTo + _
           " Union All " + wSqlDetFgFr + " Union All " + wSqlDetFgTo
  '*** If Location 'W' is selected, an additional Unions FgRm FrLoc, FgRm ToLoc is required.
   

'********************* urmi 5/7/03 *********************************'
  ' ****** Manali 3.4.1 - 02/12/08 - [UoDtFr, UoDtTo respectively replaced by wDtFr, wDtTo]
  wFrDt = IIF(UCase(MWLib.DtSetting) = "MM/DD/YY", Format(wDtFr, "mm") + "/01", "01/" + Format(wDtFr, "mm")) + "/" + Format(wDtFr, "yy") '6.1
  'wFrDt = "01/" + Format(adc("UoDtFr"), "mm") + "/" + Format(adc("UoDtFr"), "yy")
  
  '****** Sachin 3.02.0 - New fields in Where Clause
  ' ****** Manali 3.4.1 - 02/12/08 - [UoDtFr, UoDtTo respectively replaced by wDtFr, wDtTo]
  wDtCndOpnInv = " where InDt >= '" + wFrDt + "' And InDt < '" + CStr(wDtFr) + "' And IrDt >= '" + wFrDt + "' And IrDt < '" + CStr(wDtFr) + "' "
  'wDtCndOpnInv = " where InDt >= '" + wFrDt + "' And InDt < '" + CStr(adc("UoDtFr")) + "' And IrDt >= '" + wFrDt + "' And IrDt < '" + CStr(adc("UoDtFr")) + "' "
  
  '************* Reverse Logic urmi 5/7/03 *************************************
  wRLQty = "(Case IrFrRmDc When 'D' Then IrRmQty When 'C' Then -IrRmQty Else 0 End) as qOpnQty, "
  wRLWt = "(Case IrFrRmDc When 'D' Then IrRmAWt When 'C' Then -IrRmAWt Else 0 End) as qOpnWt, "
  
  If ADC("Uoyn2") = "Y" Then
'    wDtCndOpnTd = " where InDt >= '" & ADC("UoDtFr") & "' And month(InDt) = " & _
'                  CInt(Format(ADC("UoDtFr"), "mm")) & " and Year(InDt) = " & _
'                  CInt(Format(ADC("UoDtFr"), "yyyy"))

    '****** Sachin 3.02.0 - New fields in Where Clause
    ' ****** Manali 3.4.1 - 02/12/08 - [UoDtFr, UoDtTo respectively replaced by wDtFr, wDtTo]
    wDtCndOpnInv = " where InDt >= '" & wDtFr & "' And InDt <= '" & wTmpDt & "' And IrDt >= '" & wDtFr & "' And IrDt <= '" & wTmpDt & "' "
    'wDtCndOpnInv = " where InDt >= '" & adc("UoDtFr") & "' And InDt <= '" & wTmpDt & "' And IrDt >= '" & adc("UoDtFr") & "' And IrDt <= '" & wTmpDt & "' "
    
    
    wRLQty = "(Case IrFrRmDc When 'D' Then -IrRmQty When 'C' Then IrRmQty Else 0 End) as qOpnQty, "
    wRLWt = "(Case IrFrRmDc When 'D' Then -IrRmAWt When 'C' Then IrRmAWt Else 0 End) as qOpnWt, "
  End If
  '*******************************************************************************
  
  ADC("UoLn1Fr").CmpStr = "IrRmSz>= "
  ADC("UoLn1To").CmpStr = "IrRmSz<= "
  ADC("UoCstRtFr").CmpStr = "IrStkRt>= "
  ADC("UoCstRtTo").CmpStr = "IrStkRt<= "
  ' Zubin 212
  ADC("UoLotNoFr").CmpStr = "IrLotNo>= "
  ADC("UoLotNoTo").CmpStr = "IrLotNo<= "
  ADC("UoLotNoSel").CmpStr = "IrLotNo In "
  wrepcnd = ADC.RepCond
  wCnd = IIF(wrepcnd <> "", " and " + wrepcnd, "")
  
  '****** Sachin 3.02.0 - Additional Conditions For New Fields
  wAddlCnd = ""
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(ADC("UORMCTGFR")) <> "", " And ", "") + IIF(Trim(ADC("UORMCTGFR")) <> "", " IrRmCtg <= '" + Trim(ADC("UORMCTGFR")) + "'", "")
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(ADC("UORMSCTGFR")) <> "", " And ", "") + IIF(Trim(ADC("UORMSCTGFR")) <> "", " IrRmSCtg >= '" + Trim(ADC("UORMSCTGFR")) + "'", "")
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(ADC("UORMSCTGTO")) <> "", " And ", "") + IIF(Trim(ADC("UORMSCTGTO")) <> "", " IrRmSCtg <= '" + Trim(ADC("UORMSCTGTO")) + "'", "")
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(ADC("UORMSCTGSEL")) <> "", " And ", "") + IIF(Trim(ADC("UORMSCTGSEL")) <> "", " IrRmSCtg In (" + Trim(ADC("UORMSCTGSEL")) + ")", "")
  '****** Sachin 3.02.0 - Additional Conditions For New Fields
 
  '10
  ' Zubin 212
  '*** jay 2.14 *** ('' as qDesc, )
  ' ***** Manali 3.03 - 24/06/08 - Pd Mod - 'L' added
  ' ****** Manali 3.03 - 09/07/08 - RmSzDesc added and join with RmIdSz added
  '4.1.4
  'JewelSmith.2 ',' + SC.PDesc225 + ',' like '%,'+ RmSCtg +',%' changed as SC.PSCd = RmStnCls
  wSqlOpnInvFr = "Insert Into #TmpStkLed Select " + grpflds3 + ", " + _
                 " 'O' as qFlag, RmCd, IrLotNo as qLotNo, IrRmSz as qRmSz, IrRmSz2 as qRmSz2, IrRmSz3 as qRmSz3, " + wRmSzDesc + " as qRmSzDesc, IsNull(RrRngSz, '') as RrRngSz, RmCtg, IsNull(SC.PSCd,'') As qStnCls, RmSCtg, " + _
                 " IrStkRt as qStkRt, (Select PValue3 from Param where PTyp= 'RMSCTG' and PMCd= RmCtg and PSCd= RmSCtg) as qPtrYN, " + _
                 " (Case when RmCtg in ('G', 'P', 'S', 'L', 'M') then " + wPureWtOrZWt + " when RmCtg= 'X' then " + wXPureWtOrZwt + " else 0 end)" + wPurityWt + " as qRmPurityWt, " + _
                 " InCoCd as TdCoCd, '' as TdTc, '' as TdYy, '' as TdChr, 0 as TdNo, 0 as TdSr, 0 as TdSrNo, " + _
                 " '' as TdBYy, '' as TdBChr, 0 as TdBNo, '' as qDt, '' as qDesc, LocPrntCd as qPrntCd,  IrFrRmLoc as qLoc, " + _
                 wRLQty + _
                 " 0 as qDrQty, 0 as qCrQty, " + _
                 wRLWt + _
                 " 0 as qDrWt, 0 as qCrWt, '' as qFrRmLoc, '' as qToRmLoc, " + _
                 wIrRate + "," + wQwChk + ", RSCtg.PValue1 as qStkbyRt " + _
                 " From InvHd NoLock " + _
                 " join InvRm on InCoCd= IrCoCd and InTc= IrTc and InYy= IrYy and InChr= IrChr and InNo= IrNo " + IIF(gs_Partition = ctCurrPrtn, " and InPrtKey=IrPrtKey ", "") + _
                 " Join Loc Lc On LocCoCd= IrCoCd and LocCd= IrFrRmLoc and LocHideYN='N' Join RmMst Rm On RmCd= IrRmCd " + IIF(gs_Partition = ctCurrPrtn, " and RmPrtKey=IrPrtKey ", "") + _
                 " Join Param RSCtg On RSCtg.PTyp='RMSCTG' And RSCtg.PMCd=Rm.RmCtg And RSCtg.PSCd=Rm.RmSCtg  " + _
                 " Left Outer Join Param SC on SC.PTyp= 'STNCLS' and SC.PmCd= RmCtg and SC.PSCd = RmStnCls " + _
                " Left Outer Join LotMst on LotMst.LotNo = IrLotNo " + _
                 " Left Outer Join RmRt RSZ on RrTcTyp= 'RSZ' and RrCmCtg= 'C' and RrCmCd= '" + ctSelfCmCd + "' and " + _
                 " RrCtg= RmCtg and RrSCtg= RmSCtg and RrSTWGrp= '' and RrCd= '' and RrDmCtg= '' and " + _
                 " RrLsCtg= '' and RrLabMCd= '' and RrFrLn<= IrRmSz and RrToLn>= IrRmSz " + _
                 " Left Outer Join RmIdSz on RisRmCtg=RmCtg and RisRmId=RmId and RisSz1=IrRmSz " + _
                 wDtCndOpnInv + wCnd + IIF(gs_Partition = ctCurrPrtn, " and InPrtKey='" + ctCurrPrtn + "' ", "") + " And IrFrRmLocTyp= '" + ADC("UOLOCTYPFR") + "'" + IIF(wAddlCnd <> "", " And " + wAddlCnd, "") + " Option (MaxDop 1)"
  If wb_ShowOpn And Not wb_MonStart And Not wb_MonEnd Then moCn.Execute wSqlOpnInvFr

  ADC("UoLn1Fr").CmpStr = "IrRmSz>= "
  ADC("UoLn1To").CmpStr = "IrRmSz<= "
  ADC("UoCstRtFr").CmpStr = "IrStkRt>= "
  ADC("UoCstRtTo").CmpStr = "IrStkRt<= "
  ' Zubin 212 (This CmpStr block is not req as it is the same for the above and below query)
  ADC("UoLotNoFr").CmpStr = "IrLotNo>= "
  ADC("UoLotNoTo").CmpStr = "IrLotNo<= "
  ADC("UoLotNoSel").CmpStr = "IrLotNo In "
  wrepcnd = ADC.RepCond
  wCnd = IIF(wrepcnd <> "", " and " + wrepcnd, "")
  
  '*********** urmi 2.04 Reverse Logic *****************
  wRLQty = "(Case IrToRmDc When 'D' Then IrRmQty When 'C' Then -IrRmQty Else 0 End) as qOpnQty, "
  wRLWt = "(Case IrToRmDc When 'D' Then IrRmAWt When 'C' Then -IrRmAWt Else 0 End) as qOpnWt, "
  If ADC("UoYN2") = "Y" Then
    wRLQty = "(Case IrToRmDc When 'D' Then -IrRmQty When 'C' Then IrRmQty Else 0 End) as qOpnQty, "
    wRLWt = "(Case IrToRmDc When 'D' Then -IrRmAWt When 'C' Then IrRmAWt Else 0 End) as qOpnWt, "
  End If
  '**************************************************
  '11
  ' Zubin 212
  '*** jay 2.14 *** ('' as qDesc, )
  ' ***** Manali 3.03 - 24/06/08 - Pd Mod - 'L' added
  ' ****** Manali 3.03 - 09/07/08 - RmSzDesc added and join with RmIdSz added
  '4.1.4
  'JewelSmith.2 ',' + SC.PDesc225 + ',' like '%,'+ RmSCtg +',%' changed as SC.PSCd = RmStnCls
  wSqlOpnInvTo = "Insert Into #TmpStkLed Select " + grpflds3 + ", " + _
                "'O' as qFlag, RmCd, IrLotNo as qLotNo, IrRmSz as qRmSz, IrRmSz2 as qRmSz2, IrRmSz3 as qRmSz3, " + wRmSzDesc + " as qRmSzDesc, IsNull(RrRngSz, '') as RrRngSz, RmCtg, IsNull(SC.PSCd,'') As qStnCls, RmSCtg, " + _
                "IrStkRt as qStkRt, (Select PValue3 from Param where PTyp= 'RMSCTG' and PMCd= RmCtg and PSCd= RmSCtg) as qPtrYN, " + _
                "(Case when RmCtg in ('G', 'P', 'S', 'L', 'M') then " + wPureWtOrZWt + " when RmCtg= 'X' then " + wXPureWtOrZwt + " else 0 end)" + wPurityWt + " as qRmPurityWt, " + _
                "InCoCd as TdCoCd, '' as TdTc,'' as TdYy, '' as TdChr, 0 as TdNo, 0 as TdSr, 0 as TdSrNo, " + _
                "'' as TdBYy, '' as TdBChr, 0 as TdBNo, '' as qDt, '' as qDesc, LocPrntCd As qPrntCd, IrToRmLoc as qLoc, " + _
                wRLQty + _
                "0 as qDrQty, 0 as qCrQty, " + _
                wRLWt + _
                "0 as qDrWt, 0 as qCrWt, '' as qFrRmLoc, '' as qToRmLoc, " + wIrRate + _
                ", " + wQwChk + ", RSCtg.PValue1 as qStkbyRt " + _
                "From InvHd NoLock " + _
                "Join InvRm on InCoCd= IrCoCd and InTc= IrTc and InYy= IrYy and InChr= IrChr and InNo= IrNo " + IIF(gs_Partition = ctCurrPrtn, " and InPrtKey=IrPrtKey ", "") + _
                "Join Loc Lc On IrCoCd= LocCoCd and IrToRmLoc= LocCd and LocHideYN='N' Join RmMst Rm On IrRmCd= RmCd " + IIF(gs_Partition = ctCurrPrtn, " and IrPrtKey=RmPrtKey ", "") + _
                " Join Param RSCtg On RSCtg.PTyp='RMSCTG' And RSCtg.PMCd=Rm.RmCtg And RSCtg.PSCd=Rm.RmSCtg  " + _
                "Left Outer Join Param SC on SC.PTyp= 'STNCLS' and SC.PmCd= RmCtg and SC.PSCd = RmStnCls " + _
                " Left Outer Join LotMst on LotMst.LotNo = IrLotNo " + _
                "Left Outer Join RmRt RSZ on RrTcTyp= 'RSZ' and RrCmCtg= 'C' and RrCmCd= '" + ctSelfCmCd + "' and " + _
                "RrCtg= RmCtg and RrSCtg= RmSCtg and RrSTWGrp= '' and RrCd= '' and RrDmCtg= '' and " + _
                "RrLsCtg= '' and RrLabMCd= '' " + _
                "and RrFrLn<= IrRmSz and RrToLn>= IrRmSz " + _
                "Left Outer Join RmIdSz on RisRmCtg=RmCtg and RisRmId=RmId and RisSz1=IrRmSz " + _
                wDtCndOpnInv + wCnd + IIF(gs_Partition = ctCurrPrtn, " and InPrtKey='" + ctCurrPrtn + "' ", "") + " And IrToRmLocTyp= '" + ADC("UOLOCTYPFR") + "'" + IIF(wAddlCnd <> "", " And " + wAddlCnd, "") + " Option (MaxDop 1)"
  If wb_ShowOpn And Not wb_MonStart And Not wb_MonEnd Then moCn.Execute wSqlOpnInvTo

  '****** Sachin 3.02 - new fields
  ' ****** Manali 3.4.1 - 02/12/08 - [UoDtFr, UoDtTo respectively replaced by wDtFr, wDtTo]
  wDtCndDetInv = " where InDt between '" + CStr(wDtFr) + "'" + _
              " And '" + CStr(wDtTo) + "' And IrDt between '" + CStr(wDtFr) + "'" + _
              " And '" + CStr(wDtTo) + "'"
'  wDtCndDetInv = " where InDt between '" + CStr(adc("UoDtFr")) + "'" + _
'              " And '" + CStr(adc("UoDtTo")) + "' And IrDt between '" + CStr(adc("UoDtFr")) + "'" + _
'              " And '" + CStr(adc("UoDtTo")) + "'"
  
  ADC("UoLn1Fr").CmpStr = "IrRmSz>= "
  ADC("UoLn1To").CmpStr = "IrRmSz<= "
  ADC("UoCstRtFr").CmpStr = "IrStkRt>= "
  ADC("UoCstRtTo").CmpStr = "IrStkRt<= "
  ' Zubin 212 (This CmpStr block is not req as it is the same for the above and below query)
  ADC("UoLotNoFr").CmpStr = "IrLotNo>= "
  ADC("UoLotNoTo").CmpStr = "IrLotNo<= "
  ADC("UoLotNoSel").CmpStr = "IrLotNo In "
  wrepcnd = ADC.RepCond
  wCnd = IIF(wrepcnd <> "", " and " + wrepcnd, "")
  '12
  ' Zubin 212
  '*** jay 2.14 *** ('' as qDesc, )
  ' ***** Manali 3.03 - 24/06/08 - Pd Mod - 'L' added
  ' ****** Manali 3.03 - 09/07/08 - RmSzDesc added and join with RmIdSz added
  '4.1.4
  'JewelSmith.2 ',' + SC.PDesc225 + ',' like '%,'+ RmSCtg +',%' changed as SC.PSCd = RmStnCls
  wSqlDetInvFr = "Insert Into #TmpStkLed Select " + grpflds3 + ", " + _
              "'T' as qFlag, RmCd, IrLotNo as qLotNo, IrRmSz as qRmSz, IrRmSz2 as qRmSz2, IrRmSz3 as qRmSz3, " + wRmSzDesc + " as qRmSzDesc, IsNull(RrRngSz, '') as RrRngSz, RmCtg, IsNull(SC.PSCd,'') As qStnCls, RmSCtg, " + _
              "IrStkRt as qStkRt, (Select PValue3 from Param where PTyp= 'RMSCTG' and PMCd= RmCtg and PSCd= RmSCtg) as qPtrYN, " + _
              "(Case when RmCtg in ('G', 'P', 'S', 'L', 'M') then " + wPureWtOrZWt + " when RmCtg= 'X' then " + wXPureWtOrZwt + " else 0 end)" + wPurityWt + " as qRmPurityWt, " + _
              "InCoCd as TdCoCd, IrTc as TdTc, IrYy as TdYy, IrChr as TdChr, IrNo as TdNo, IrSr as TdSr, IrSrNo as TdSrNo, " + _
              "'' as TdBYy, '' as TdBChr, 0 as TdBNo, " + _
              "InDt as qDt, '' as qDesc, LocPrntCd as qPrntCd, IrFrRmLoc as qLoc, 0 as qOpnQty, " + _
              "(Case IrFrRmDc When 'D' Then IrRmQty Else 0 End) as qDrQty, " + _
              "(Case IrFrRmDc When 'C' Then IrRmQty else 0 End) as qCrQty, " + _
              "0 as qOpnWt, (Case IrFrRmDc When 'D' Then IrRmAWt Else 0 End) as qDrWt, " + _
              "(Case IrFrRmDc When 'C' Then IrRmAWt else 0 End) as qCrWt, " + _
              "(case when IrFrRmDc= 'C' then IrFrRmLoc else IrToRmLoc end) as qFrRmLoc, " + _
              "(case when IrFrRmDc= 'D' then IrFrRmLoc else IrToRmLoc end) as qToRmLoc, " + wIrRate + _
              ", " + wQwChk + ", RSCtg.PValue1 as qStkbyRt " + _
              " From InvHd NoLock join InvRm on IrCoCd= InCoCd and IrTc= InTc and IrYy= InYy and IrChr= InChr and IrNo= InNo " + IIF(gs_Partition = ctCurrPrtn, " and IrPrtKey=InPrtKey ", "") + _
              "Join Loc Lc On LocCoCd= IrCoCd and LocCd= IrFrRmLoc and LocHideYN='N' Join RmMst Rm On IrRmCd= RmCd " + IIF(gs_Partition = ctCurrPrtn, " and IrPrtKey=RmPrtKey ", "") + _
              " Join Param RSCtg On RSCtg.PTyp='RMSCTG' And RSCtg.PMCd=Rm.RmCtg And RSCtg.PSCd=Rm.RmSCtg  " + _
              "Left Outer Join Param SC on SC.PTyp= 'STNCLS' and SC.PmCd= RmCtg and SC.PSCd = RmStnCls " + _
              " Left Outer Join LotMst on LotMst.LotNo = IrLotNo " + _
              "Left Outer Join RmRt RSZ on RrTcTyp= 'RSZ' and RrCmCtg= 'C' and RrCmCd= '" + ctSelfCmCd + "' and " + _
              "RrCtg= RmCtg and RrSCtg= RmSCtg and RrSTWGrp= '' and RrCd= '' and RrDmCtg= '' and " + _
              "RrLsCtg= '' and RrLabMCd= '' and RrFrLn<= IrRmSz and RrToLn>= IrRmSz " + _
              "Left Outer Join RmIdSz on RisRmCtg=RmCtg and RisRmId=RmId and RisSz1=IrRmSz " + _
              wDtCndDetInv + wCnd + IIF(gs_Partition = ctCurrPrtn, " and InPrtKey='" + ctCurrPrtn + "' ", "") + " And IrFrRmLocTyp= '" + ADC("UOLOCTYPFR") + "'" + IIF(wAddlCnd <> "", " And " + wAddlCnd, "") + " Option (MaxDop 1)"
   moCn.Execute wSqlDetInvFr
  
  ' Zubin 212 (Below 4 cmpstr are not required)
  'adc("UoLn1Fr").CmpStr = "FrRmSz>= "
  'adc("UoLn1To").CmpStr = "FrRmSz<= "
  'adc("UoCstRtFr").CmpStr = "FrRmStkRt>= "
  'adc("UoCstRtTo").CmpStr = "FrRmStkRt<= "
  
  ADC("UoLn1Fr").CmpStr = "IrRmSz>= "
  ADC("UoLn1To").CmpStr = "IrRmSz<= "
  ADC("UoCstRtFr").CmpStr = "IrStkRt>= "
  ADC("UoCstRtTo").CmpStr = "IrStkRt<= "
  ' Zubin 212 (This CmpStr block is not req as it is the same for the above and below query)
  ADC("UoLotNoFr").CmpStr = "IrLotNo>= "
  ADC("UoLotNoTo").CmpStr = "IrLotNo<= "
  ADC("UoLotNoSel").CmpStr = "IrLotNo In "
  
  wrepcnd = ADC.RepCond
  wCnd = IIF(wrepcnd <> "", " and " + wrepcnd, "")
  '13
  ' Zubin 212
  '*** jay 2.14 *** ('' as qDesc, )
  ' ***** Manali 3.03 - 24/06/08 - Pd Mod - 'L' added
  ' ****** Manali 3.03 - 09/07/08 - RmSzDesc added and join with RmIdSz added
  '4.1.4
  'JewelSmith.2 ',' + SC.PDesc225 + ',' Like '%,'+ RmSCtg +',%' changed as SC.PSCd = RmStnCls
  wSqlDetInvTo = "Insert Into #TmpStkLed Select " + grpflds3 + ", " + _
              "'T' as qFlag, RmCd, IrLotNo as qLotNo, IrRmSz as qRmSz, IrRmSz2 as qRmSz2, IrRmSz3 as qRmSz3, " + wRmSzDesc + " as qRmSzDesc, IsNull(RrRngSz, '') as RrRngSz, RmCtg, IsNull(SC.PSCd,'') As qStnCls, RmSCtg, " + _
              "IrStkRt as qStkRt, (Select PValue3 from Param where PTyp= 'RMSCTG' and PMCd= RmCtg and PSCd= RmSCtg) as qPtrYN, " + _
              "(Case when RmCtg in ('G', 'P', 'S', 'L', 'M') then " + wPureWtOrZWt + " when RmCtg= 'X' then " + wXPureWtOrZwt + " else 0 end)" + wPurityWt + " as qRmPurityWt, " + _
              "InCoCd as TdCoCd, IrTc as TdTc, IrYy as TdYy, IrChr as TdChr, IrNo as TdNo, IrSr as TdSr, IrSrNo as TdSrNo, " + _
              "'' as TdBYy, '' as TdBChr, 0 as TdBNo, " + _
              "InDt as qDt, '' as qDesc, LocPrntCd as qPrntCd, IrToRmLoc as qLoc, 0 as qOpnQty, " + _
              "Case IrToRmDc When 'D' Then IrRmQty Else 0 End  as qDrQty, " + _
              "Case IrToRmDc When 'C' Then IrRmQty else 0 End  as qCrQty, " + _
              "0 as qOpnWt, Case IrToRmDc When 'D' Then IrRmAWt Else 0 End  as qDrWt, " + _
              "Case IrToRmDc When 'C' Then IrRmAWt else 0 End  as qCrWt, " + _
              "(case when IrFrRmDc= 'C' then IrFrRmLoc else IrToRmLoc end) as qFrRmLoc, " + _
              "(case when IrFrRmDc= 'D' then IrFrRmLoc else IrToRmLoc end) as qToRmLoc, " + wIrRate + _
              ", " + wQwChk + ", RSCtg.PValue1 as qStkbyRt " + _
              "From InvHd NoLock join InvRm on IrCoCd= InCoCd and IrTc= InTc and IrYy= InYy and IrChr= InChr and IrNo= InNo " + IIF(gs_Partition = ctCurrPrtn, " and IrPrtKey=InPrtKey ", "") + _
              "Join Loc Lc On LocCoCd= IrCoCd and LocCd= IrToRmLoc and LocHideYN='N' Join RmMst Rm On IrRmCd= RmCd " + IIF(gs_Partition = ctCurrPrtn, " and IrPrtKey=RmPrtKey ", "") + _
            " Left Outer Join LotMst on LotMst.LotNo = IrLotNo " + _
              "Join Param RSCtg On RSCtg.PTyp='RMSCTG' And RSCtg.PMCd=RmCtg And RSCtg.PSCd=RmSCtg " + _
              "Left Outer Join Param SC on SC.PTyp= 'STNCLS' and SC.PmCd= RmCtg and SC.PSCd = RmStnCls " + _
              "Left Outer Join RmRt RSZ on RrTcTyp= 'RSZ' and RrCmCtg= 'C' and RrCmCd= '" + ctSelfCmCd + "' and " + _
              "RrCtg= RmCtg and RrSCtg= RmSCtg and RrSTWGrp= '' and RrCd= '' and RrDmCtg= '' and " + _
              "RrLsCtg= '' and RrLabMCd= '' and RrFrLn<= IrRmSz and RrToLn>= IrRmSz " + _
              "Left Outer Join RmIdSz on RisRmCtg=RmCtg and RisRmId=RmId and RisSz1=IrRmSz " + _
              wDtCndDetInv + wCnd + IIF(gs_Partition = ctCurrPrtn, " and InPrtKey='" + ctCurrPrtn + "' ", "") + " And IrToRmLocTyp= '" + ADC("UOLOCTYPFR") + "'" + IIF(wAddlCnd <> "", " And " + wAddlCnd, "") + " Option (MaxDop 1)"
  moCn.Execute wSqlDetInvTo
  
  
  
  wSqlOpnXdFr = "": wSqlOpnXdTo = "": wSqlDetXdFr = "": wSqlDetXdTo = ""
  wFrDt1 = IIF(UCase(MWLib.DtSetting) = "MM/DD/YY", Format(wDtFr, "mm") + "/01", "01/" + Format(wDtFr, "mm")) + "/" + Format(wDtFr, "yy") '6.1
  
  wDtCndOpnXd = " where BClsDt >= '" + wFrDt1 + "' And BClsDt < '" + CStr(wDtFr) + "' and OmDt >= '" + wFrDt1 + "' And OmDt < '" + CStr(wDtFr) + "' "
  
  wRLQty = "(Case OmFrRmDc When 'D' Then FrRmQty When 'C' Then -FrRmQty Else 0 End) as qOpnQty, "
  wRLWt = "(Case OmFrRmDc When 'D' Then FrRmWt When 'C' Then -FrRmWt Else 0 End) as qOpnWt, "
  '578-34 Str(FdKey, 7) changed to Str(FdKey, 8)
  wBagJoin = "Join (Select BIdNo, BPrtKey, BCoCd, BYy, BChr, BNo, max(BCnvJtYn) as BCnvJtYn, max(BClsDt) as BClsDt, max(FdYy+Str(FdKey, 8)) as qFdYyKey " + _
             "From Bag Join Fgd On  FdCoCd=BCoCd and FdBYy=BYy and FdBChr=BChr and FdBNo=BNo and FdBIdNo=BIdNo " + IIF(gs_Partition = ctCurrPrtn, " and FdPrtKey=BPrtKey ", "") + _
             "Group By BPrtKey, BIdNo, BCoCd, BYy, BChr, BNo) Bag On " + _
             "BCoCd = FdCoCd And BYy = FdBYy And BChr = FdBChr And BNo = FdBNo and BidNo=FdBIdNo " + IIF(gs_Partition = ctCurrPrtn, " and BPrtKey=FdPrtKey ", "") + _
             "and BCnvJtYn='Y' and qFdYyKey=(FdYy+str(FdKey, 8)) "
  
  If ADC("Uoyn2") = "Y" Then
    wDtCndOpnXd = " where BClsDt >= '" & wDtFr & "' And BClsDt <= '" & wTmpDt & "' and OmDt >= '" & wDtFr & "' And OmDt <= '" & wTmpDt & "' "
    
    wRLQty = "(Case OmFrRmDc When 'D' Then -FrRmQty When 'C' Then FrRmQty Else 0 End) as qOpnQty, "
    wRLWt = "(Case OmFrRmDc When 'D' Then -FrRmWt When 'C' Then FrRmWt Else 0 End) as qOpnWt, "
  End If
  
  Dim wOrdJoin As String
  ADC("UoLn1Fr").CmpStr = "FrRmSz>= "
  ADC("UoLn1To").CmpStr = "FrRmSz<= "
  ADC("UoCstRtFr").CmpStr = "FrRmStkRt>= "
  ADC("UoCstRtTo").CmpStr = "FrRmStkRt<= "
  ADC("UoLotNoFr").CmpStr = "FrLotNo>= "
  ADC("UoLotNoTo").CmpStr = "FrLotNo<= "
  ADC("UoLotNoSel").CmpStr = "FrLotNo In "
  wrepcnd = ADC.RepCond
  wCnd = IIF(wrepcnd <> "", " and " + wrepcnd, "")
  
  wAddlCnd = ""
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(ADC("UORMCTGFR")) <> "", " And ", "") + IIF(Trim(ADC("UORMCTGFR")) <> "", " FrRmCtg <= '" + Trim(ADC("UORMCTGFR")) + "'", "")
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(ADC("UORMSCTGFR")) <> "", " And ", "") + IIF(Trim(ADC("UORMSCTGFR")) <> "", " FrRmSCtg >= '" + Trim(ADC("UORMSCTGFR")) + "'", "")
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(ADC("UORMSCTGTO")) <> "", " And ", "") + IIF(Trim(ADC("UORMSCTGTO")) <> "", " FrRmSCtg <= '" + Trim(ADC("UORMSCTGTO")) + "'", "")
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(ADC("UORMSCTGSEL")) <> "", " And ", "") + IIF(Trim(ADC("UORMSCTGSEL")) <> "", " FrRmSCtg In (" + Trim(ADC("UORMSCTGSEL")) + ")", "")
  
 
  '14
  wSqlOpnXdFr = "Insert Into #TmpStkLed Select " + grpflds2 + "," + _
                "'O' as qFlag, RmCd, FrLotNo as qLotNo, FrRmSz as qRmSz, FrRmSz2 as qRmSz2, FrRmSz3 as qRmSz3, " + wRmSzDesc + " as qRmSzDesc, IsNull(RrRngSz, '') as RrRngSz, RmCtg, IsNull(SC.PSCd,'') As qStnCls, RmSCtg, " + _
                "FrRmStkRt as qStkRt, (Select PValue3 from Param where PTyp= 'RMSCTG' and PMCd= RmCtg and PSCd= RmSCtg) as qPtrYN, " + _
                "(Case when RmCtg in ('G', 'P', 'S', 'L', 'M') then " + wPureWtOrZWt + " when RmCtg= 'X' then " + wXPureWtOrZwt + " else 0 end)" + wPurityWt + " as qRmPurityWt, " + _
                "OmCoCd as TdCoCd, '' as TdTc, '' as TdYy, '' as TdChr, 0 as TdNo, 0 as TdSr, " + _
                "0 as TdSrNo, '' as TdBYy, '' as TdBChr, 0 as TdBNo, '' as qDt, '' as qDesc, " + _
                "LocPrntCd as qPrntCd, OmFrRmLoc as qLoc, " + _
                wRLQty + _
                "0 as qDrQty, 0 as qCrQty, " + _
                wRLWt + _
                "0 as qDrWt, 0 as qCrWt, '' as qFrRmLoc, '' as qToRmLoc, " + wFrRate + _
                ", " + wQwChk + ", RSCtg.PValue1 as qStkbyRt "
'JewelSmith.2 ',' + SC.PDesc225 + ',' like '%,'+ RmSCtg +',%' changed as SC.PSCd = RmStnCls
  wSqlOpnXdFr = wSqlOpnXdFr + _
                " From OrdMst  NoLock " + _
                "Join OrdDsg On  OdCoCd=OmCoCd and OdTc=OmTc and OdYy=OmYy and OdChr=OmChr and OdNo=OmNo and OdOmIdNo=OmIdNo and OdPrtKey=OmPrtKey " + IIF(gs_Partition = ctCurrPrtn, " and OmPrtKey=OdPrtKey ", "") + _
                "join Fgd On FdCoCd=OdCoCd and FdBYy=OdBYy and FdBChr=OdBChr and FdBNo=OdBNo " + IIF(gs_Partition = ctCurrPrtn, " and FdPrtKey=OdPrtKey ", "") + _
                "join FgRm on FdIdNo=FrFdIdNo and FdCoCd= FrCoCd and FdTc= FrTc and FdYy= FrYy and " + _
                "FdChr= FrChr and FdNo= FrNo and FdSr= FrSr " + IIF(gs_Partition = ctCurrPrtn, " and FdPrtKey=FrPrtKey ", "") + _
                wBagJoin + _
                "Join Loc Lc On OmCoCd= LocCoCd and OmFrRmLoc= LocCd and LocHideYN='N' Join RmMst Rm On FrRmCd= RmCd " + IIF(gs_Partition = ctCurrPrtn, " and FrPrtKey=RmPrtKey ", "") + _
                "Join Param RSCtg On RSCtg.PTyp='RMSCTG' And RSCtg.PMCd=RmCtg And RSCtg.PSCd=RmSCtg " + _
                "Left Outer Join Param SC on SC.PTyp= 'STNCLS' and SC.PmCd= RmCtg and SC.PSCd = RmStnCls " + _
                " Left Outer Join LotMst on LotMst.LotNo = FrLotNo " + _
                "Left Outer Join RmRt RSZ on RrTcTyp= 'RSZ' and RrCmCtg= 'C' and RrCmCd= '" + ctSelfCmCd + "' and " + _
                "RrCtg= RmCtg and RrSCtg= RmSCtg and RrSTWGrp= '' and RrCd= '' and RrDmCtg= '' and " + _
                "RrLsCtg= '' and RrLabMCd= '' and RrFrLn<= FrRmSz and RrToLn>= FrRmSz " + _
                "Left Outer Join RmIdSz on RisRmCtg=RmCtg and RisRmId=RmId and RisSz1=FrRmSz " + _
                wDtCndOpnXd + wCnd + " and OdTc=(Select PMCd from Param Where PTyp='TC' and PMCd='JCN' and PValue='JT') " + IIF(gs_Partition = ctCurrPrtn, " and OmPrtKey='" + ctCurrPrtn + "' ", "") + _
                IIF(wAddlCnd <> "", " And " + wAddlCnd, "") + " Option (MaxDop 1)"
  If wb_ShowOpn And Not wb_MonStart And Not wb_MonEnd Then moCn.Execute wSqlOpnXdFr
  
  ADC("UoLn1Fr").CmpStr = "FrRmSz>= "
  ADC("UoLn1To").CmpStr = "FrRmSz<= "
  ADC("UoCstRtFr").CmpStr = "FrRmStkRt>= "
  ADC("UoCstRtTo").CmpStr = "FrRmStkRt<= "
  ' Zubin 212 (This CmpStr block is not req as it is the same for the above and below query)
  ADC("UoLotNoFr").CmpStr = "FrLotNo>= "
  ADC("UoLotNoTo").CmpStr = "FrLotNo<= "
  ADC("UoLotNoSel").CmpStr = "FrLotNo In "
  
  wrepcnd = ADC.RepCond
  wCnd = IIF(wrepcnd <> "", " and " + wrepcnd, "")
  'wCnd = wCnd + IIF(wCnd <> "", " and ", " ") + wLoctyp
  
  
  wRLQty = "(Case OmToRmDc When 'D' Then FrRmQty When 'C' Then -FrRmQty Else 0 End) as qOpnQty, "
  wRLWt = "(Case OmToRmDc When 'D' Then FrRmWt When 'C' Then -FrRmWt Else 0 End) as qOpnWt, "
  
  If ADC("UoYn2") = "Y" Then
    wRLQty = "(Case OmToRmDc When 'D' Then -FrRmQty When 'C' Then FrRmQty Else 0 End) as qOpnQty, "
    wRLWt = "(Case OmToRmDc When 'D' Then -FrRmWt When 'C' Then FrRmWt Else 0 End) as qOpnWt, "
  End If
  
  '15
  'JewelSmith.2 ',' + SC.PDesc225 + ',' like '%,'+ RmSCtg +',%' changed as SC.PSCd = RmStnCls
  wSqlOpnXdTo = "Insert Into #TmpStkLed Select " + grpflds2 + ", " + _
                "'O' as qFlag, RmCd, FrLotNo as qLotNo, FrRmSz as qRmSz, FrRmSz2 as qRmSz2, FrRmSz3 as qRmSz3, " + wRmSzDesc + " as qRmSzDesc, IsNull(RrRngSz, '') as RrRngSz, RmCtg, IsNull(SC.PSCd,'') As qStnCls, RmSCtg, " + _
                "FrRmStkRt as qStkRt, (Select PValue3 from Param where PTyp= 'RMSCTG' and PMCd= RmCtg and PSCd= RmSCtg) as qPtrYN, " + _
                "(Case when RmCtg in ('G', 'P', 'S', 'L', 'M') then " + wPureWtOrZWt + " when RmCtg= 'X' then " + wXPureWtOrZwt + " else 0 end)" + wPurityWt + " as qRmPurityWt, " + _
                "OmCoCd as TdCoCd, '' as TdTc, '' as TdYy, '' as TdChr, 0 as TdNo, 0 as TdSr, " + _
                "0 as TdSrNo, '' as TdBYy, '' as TdBChr, 0 as TdBNo, '' as qDt, '' as qDesc, " + _
                "LocPrntCd as qPrntCd, OmToRmLoc as qLoc, " + _
                wRLQty + _
                "0 as qDrQty, 0 as qCrQty, " + _
                wRLWt + _
                "0 as qDrWt, 0 as qCrWt, '' as qFrRmLoc, '' as qToRmLoc, " + wFrRate + _
                ", " + wQwChk + ", RSCtg.PValue1 as qStkbyRt "
  wSqlOpnXdTo = wSqlOpnXdTo + " From OrdMst  NoLock " + _
                "Join OrdDsg On  OdCoCd=OmCoCd and OdTc=OmTc and OdYy=OmYy and OdChr=OmChr and OdNo=OmNo and OdOmIdNo=OmIdNo and OdPrtKey=OmPrtKey " + IIF(gs_Partition = ctCurrPrtn, " and OmPrtKey=OdPrtKey ", "") + _
                "join Fgd On FdCoCd=OdCoCd and FdBYy=OdBYy and FdBChr=OdBChr and FdBNo=OdBNo " + IIF(gs_Partition = ctCurrPrtn, " and FdPrtKey=OdPrtKey ", "") + _
                "join FgRm on FdIdNo=FrFdIdNo and FdCoCd= FrCoCd and FdTc= FrTc and FdYy= FrYy and " + _
                "FdChr= FrChr and FdNo= FrNo and FdSr= FrSr " + IIF(gs_Partition = ctCurrPrtn, " and FdPrtKey=FrPrtKey ", "") + _
                wBagJoin + _
                "Join Loc Lc On OmCoCd= LocCoCd and OmToRmLoc= LocCd and LocHideYN='N' Join RmMst Rm On FrRmCd= RmCd " + IIF(gs_Partition = ctCurrPrtn, " and RmPrtKey=FrPrtKey ", "") + _
                "Join Param RSCtg On RSCtg.PTyp='RMSCTG' And RSCtg.PMCd=RmCtg And RSCtg.PSCd=RmSCtg " + _
              " Left Outer Join LotMst on LotMst.LotNo = FrLotNo " + _
                "Left Outer Join Param SC on SC.PTyp= 'STNCLS' and SC.PmCd= RmCtg and SC.PSCd = RmStnCls " + _
                "Left Outer Join RmRt RSZ on RrTcTyp= 'RSZ' and RrCmCtg= 'C' and RrCmCd= '" + ctSelfCmCd + "' and " + _
                "RrCtg= RmCtg and RrSCtg= RmSCtg and RrSTWGrp= '' and RrCd= '' and RrDmCtg= '' and " + _
                "RrLsCtg= '' and RrLabMCd= '' and RrFrLn<= FrRmSz and RrToLn>= FrRmSz " + _
                "Left Outer Join RmIdSz on RisRmCtg=RmCtg and RisRmId=RmId and RisSz1=FrRmSz " + _
                wDtCndOpnXd + wCnd + " and OdTc=(Select PMCd from Param Where PTyp='TC' and PMCd='JCN' and PValue='JT') " + IIF(gs_Partition = ctCurrPrtn, " and OmPrtKey='" + ctCurrPrtn + "' ", "") + _
                IIF(wAddlCnd <> "", " And " + wAddlCnd, "") + " Option (MaxDop 1)"
  If wb_ShowOpn And Not wb_MonStart And Not wb_MonEnd Then moCn.Execute wSqlOpnXdTo
  
  wDtCndDetXd = " where BClsDt between '" + CStr(wDtFr) + "'" + _
              " And '" + CStr(wDtTo) + "' And OmDt between '" + CStr(wDtFr) + "'" + _
              " And '" + CStr(wDtTo) + "'"
  '578-34 Str(FdKey, 7) changed to Str(FdKey, 8)
  wBagJoin = "Join (Select BIdNo, BPrtKey, BCoCd, BYy, BChr, BNo, max(BCnvJtYn) as BCnvJtYn, max(BClsDt) as BClsDt, max(FdYy+Str(FdKey, 8)) as qFdYyKey " + _
             "From Bag Join Fgd On  FdCoCd=BCoCd and FdBYy=BYy and FdBChr=BChr and FdBNo=BNo and FdBIdNo=BIdNo " + IIF(gs_Partition = ctCurrPrtn, " and FdPrtKey=BPrtKey ", "") + _
             "Group By BPrtKey, BIdNo, BCoCd, BYy, BChr, BNo) Bag On " + _
             "BCoCd = FdCoCd And BYy = FdBYy And BChr = FdBChr And BNo = FdBNo and BidNo=FdBIdNo " + IIF(gs_Partition = ctCurrPrtn, " and BPrtKey=FdPrtKey ", "") + _
             "and BCnvJtYn='Y' and qFdYyKey=(FdYy+str(FdKey, 8)) "
    
  ADC("UoLn1Fr").CmpStr = "FrRmSz>= "
  ADC("UoLn1To").CmpStr = "FrRmSz<= "
  ADC("UoCstRtFr").CmpStr = "FrRmStkRt>= "
  ADC("UoCstRtTo").CmpStr = "FrRmStkRt<= "
  ' Zubin 212 (This CmpStr block is not req as it is the same for the above and below query)
  ADC("UoLotNoFr").CmpStr = "FrLotNo>= "
  ADC("UoLotNoTo").CmpStr = "FrLotNo<= "
  ADC("UoLotNoSel").CmpStr = "FrLotNo In "
  wrepcnd = ADC.RepCond
  wCnd = IIF(wrepcnd <> "", " and " + wrepcnd, "")
  '16
  'JewelSmith.2 ',' + SC.PDesc225 + ',' like '%,'+ RmSCtg +',%' changed as SC.PSCd = RmStnCls
  wSqlDetXdFr = "Insert Into #TmpStkLed Select " + grpflds2 + "," + _
              "'T' as qFlag, RmCd, FrLotNo as qLotNo, FrRmSz as qRmSz, FrRmSz2 as qRmSz2, FrRmSz3 as qRmSz3, " + wRmSzDesc + " as qRmSzDesc, IsNull(RrRngSz, '') as RrRngSz, RmCtg, IsNull(SC.PSCd,'') As qStnCls, RmSCtg, " + _
              "FrRmStkRt as qStkRt, (Select PValue3 from Param where PTyp= 'RMSCTG' and PMCd= RmCtg and PSCd= RmSCtg) as qPtrYN, " + _
              "(Case when RmCtg in ('G', 'P', 'S', 'L', 'M') then " + wPureWtOrZWt + " when RmCtg= 'X' then " + wXPureWtOrZwt + " else 0 end)" + wPurityWt + " as qRmPurityWt, " + _
              "OmCoCd as TdCoCd, OdTc as TdTc, OdYy as TdYy, OdChr as TdChr, OdNo as TdNo, OdSr as TdSr, FrSrNo as TdSrNo, " + _
              "FdBYy as TdBYy, FdBChr as TdBChr, FdBNo as TdBNo, OmDt as qDt, '' as qDesc, " + _
              "LocPrntCd as qPrntCd, OmFrRmLoc as qLoc, 0 as qOpnQty, " + _
              "(Case OmFrRmDc When 'D' Then FrRmQty Else 0 End) as qDrQty, " + _
              "(Case OmFrRmDc When 'C' Then FrRmQty else 0 End) as qCrQty, 0 as qOpnWt, " + _
              "(Case OmFrRmDc When 'D' Then FrRmWt Else 0 End) as qDrWt, " + _
              "(Case OmFrRmDc When 'C' Then FrRmWt else 0 End) as qCrWt, " + _
              "(case when OmFrRmDc= 'C' then OmFrRmLoc else OmToRmLoc end) as qFrRmLoc, " + _
              "(case when OmFrRmDc= 'D' then OmFrRmLoc else OmToRmLoc end) as qToRmLoc, " + wFrRate + _
              ", " + wQwChk + ", RSCtg.PValue1 as qStkbyRt "
  wSqlDetXdFr = wSqlDetXdFr + " From OrdMst  NoLock " + _
              "Join OrdDsg On  OdCoCd=OmCoCd and OdTc=OmTc and OdYy=OmYy and OdChr=OmChr and OdNo=OmNo and OdOmIdNo=OmIdNo and OdPrtKey=OmPrtKey " + IIF(gs_Partition = ctCurrPrtn, " and OmPrtKey=OdPrtKey ", "") + _
              "join Fgd On FdCoCd=OdCoCd and FdBYy=OdBYy and FdBChr=OdBChr and FdBNo=OdBNo " + IIF(gs_Partition = ctCurrPrtn, " and FdPrtKey=OdPrtKey ", "") + _
              "join FgRm on FdIdNo=FrFdIdNo and FdCoCd= FrCoCd and FdTc= FrTc and FdYy= FrYy and " + _
              "FdChr= FrChr and FdNo= FrNo and FdSr= FrSr " + IIF(gs_Partition = ctCurrPrtn, " and FdPrtKey=FrPrtKey ", "") + wBagJoin + _
              "Join Loc Lc On LocCoCd= OmCoCd and LocCd= OmFrRmLoc and LocHideYN='N' Join RmMst Rm On FrRmCd= RmCd " + IIF(gs_Partition = ctCurrPrtn, " and RmPrtKey=FrPrtKey ", "") + _
              "Join Param RSCtg On RSCtg.PTyp='RMSCTG' And RSCtg.PMCd=RmCtg And RSCtg.PSCd=RmSCtg " + _
              " Left Outer Join LotMst on LotMst.LotNo = FrLotNo " + _
              "Left Outer Join Param SC on SC.PTyp= 'STNCLS' and SC.PmCd= RmCtg and SC.PSCd = RmStnCls " + _
              "Left Outer Join RmRt RSZ on RrTcTyp= 'RSZ' and RrCmCtg= 'C' and RrCmCd= '" + ctSelfCmCd + "' and " + _
              "RrCtg= RmCtg and RrSCtg= RmSCtg and RrSTWGrp= '' and RrCd= '' and RrDmCtg= '' and " + _
              "RrLsCtg= '' and RrLabMCd= '' and RrFrLn<= FrRmSz and RrToLn>= FrRmSz " + _
              "Left Outer Join RmIdSz on RisRmCtg=RmCtg and RisRmId=RmId and RisSz1=FrRmSz " + _
              wDtCndDetXd + wCnd + " and OdTc=(Select PMCd from Param Where PTyp='TC' and PMCd='JCN' and PValue='JT') " + IIF(gs_Partition = ctCurrPrtn, " and OmPrtKey='" + ctCurrPrtn + "' ", "") + _
              IIF(wAddlCnd <> "", " And " + wAddlCnd, "") + " Option (MaxDop 1)"
  moCn.Execute wSqlDetXdFr
  
  ADC("UoLn1Fr").CmpStr = "FrRmSz>= "
  ADC("UoLn1To").CmpStr = "FrRmSz<= "
  ADC("UoCstRtFr").CmpStr = "FrRmStkRt>= "
  ADC("UoCstRtTo").CmpStr = "FrRmStkRt<= "
  ' Zubin 212 (This CmpStr block is not req as it is the same for the above and below query)
  ADC("UoLotNoFr").CmpStr = "FrLotNo>= "
  ADC("UoLotNoTo").CmpStr = "FrLotNo<= "
  ADC("UoLotNoSel").CmpStr = "FrLotNo In "
  
  wrepcnd = ADC.RepCond
  wCnd = IIF(wrepcnd <> "", " and " + wrepcnd, "")
  'wCnd = wCnd + IIF(wCnd <> "", " and ", " ") + wLoctyp
  
  '17
  '"(case when IrFrRmDc= 'C' then IrFrRmLoc else IrToRmLoc end) as qFrRmLoc, "
  '"(case when FrFrRmDc= 'C' then (Case FrFrRmLocTyp When 'F' Then 'XD' When 'W' Then FrToRmLoc End) else (Case FrToRmLocTyp When 'F' Then 'XD' When 'W' Then FrFrRmLoc End) end) as qFrRmLoc, " + _
  '"(case when FrFrRmDc= 'D' then (Case FrFrRmLocTyp When 'F' Then 'XD' When 'W' Then FrToRmLoc End) else (Case FrToRmLocTyp When 'F' Then 'XD' When 'W' Then FrFrRmLoc End) end) as qToRmLoc, " + wFrRate
  'JewelSmith.2 ',' + SC.PDesc225 + ',' like '%,'+ RmSCtg +',%' changed as SC.PSCd = RmStnCls
  wSqlDetXdTo = "Insert Into #TmpStkLed Select " + grpflds2 + "," + _
              "'T' as qFlag, RmCd, FrLotNo as qLotNo, FrRmSz as qRmSz, FrRmSz2 as qRmSz2, FrRmSz3 as qRmSz3, " + wRmSzDesc + " as qRmSzDesc, IsNull(RrRngSz, '') as RrRngSz, RmCtg, IsNull(SC.PSCd,'') As qStnCls, RmSCtg, " + _
              "FrRmStkRt as qStkRt, (Select PValue3 from Param where PTyp= 'RMSCTG' and PMCd= RmCtg and PSCd= RmSCtg) as qPtrYN, " + _
              "(Case when RmCtg in ('G', 'P', 'S', 'L', 'M') then " + wPureWtOrZWt + " when RmCtg= 'X' then " + wXPureWtOrZwt + " else 0 end)" + wPurityWt + " as qRmPurityWt, " + _
              "OmCoCd as TdCoCd, OdTc as TdTc, OdYy as TdYy, OdChr as TdChr, OdNo as TdNo, OdSr as TdSr, FrSrNo as TdSrNo, " + _
              "FdBYy as TdBYy, FdBChr as TdBChr, FdBNo as TdBNo, OmDt as qDt, '' as qDesc, " + _
              "LocPrntCd as qPrntCd, OmToRmLoc as qLoc, 0 as qOpnQty, " + _
              "(Case OmToRmDc When 'D' Then FrRmQty Else 0 End) as qDrQty, " + _
              "(Case OmToRmDc When 'C' Then FrRmQty else 0 End) as qCrQty, 0 as qOpnWt, " + _
              "(Case OmToRmDc When 'D' Then FrRmWt Else 0 End) as qDrWt, " + _
              "(Case OmToRmDc When 'C' Then FrRmWt else 0 End) as qCrWt, " + _
              "(case when OmFrRmDc= 'C' then OmFrRmLoc else OmToRmLoc end) as qFrRmLoc, " + _
              "(case when OmFrRmDc= 'D' then OmFrRmLoc else OmToRmLoc end) as qToRmLoc, " + wFrRate + _
              ", " + wQwChk + ", RSCtg.PValue1 as qStkbyRt "
              
  wSqlDetXdTo = wSqlDetXdTo + " From OrdMst  NoLock " + _
              "Join OrdDsg On  OdCoCd=OmCoCd and OdTc=OmTc and OdYy=OmYy and OdChr=OmChr and OdNo=OmNo and OdOmIdNo=OmIdNo and OdPrtKey=OmPrtKey " + IIF(gs_Partition = ctCurrPrtn, " and OmPrtKey=OdPrtKey ", "") + _
              "join Fgd On FdCoCd=OdCoCd and FdBYy=OdBYy and FdBChr=OdBChr and FdBNo=OdBNo " + IIF(gs_Partition = ctCurrPrtn, " and FdPrtKey=OdPrtKey ", "") + _
              "join FgRm on FdIdNo=FrFdIdNo and FdCoCd= FrCoCd and FdTc= FrTc and FdYy= FrYy and " + _
              "FdChr= FrChr and FdNo= FrNo and FdSr= FrSr " + IIF(gs_Partition = ctCurrPrtn, " and FdPrtKey=FrPrtKey ", "") + wOrdJoin + wBagJoin + _
              "Join Loc Lc On LocCoCd= OmCoCd and LocCd= OmToRmLoc and LocHideYN='N' Join RmMst Rm On FrRmCd= RmCd " + IIF(gs_Partition = ctCurrPrtn, " and FrPrtKey=RmPrtKey ", "") + _
              "Join Param RSCtg On RSCtg.PTyp='RMSCTG' And RSCtg.PMCd=RmCtg And RSCtg.PSCd=RmSCtg " + _
              " Left Outer Join LotMst on LotMst.LotNo = FrLotNo " + _
              "Left Outer Join Param SC on SC.PTyp= 'STNCLS' and SC.PmCd= RmCtg and SC.PSCd = RmStnCls " + _
              "Left Outer Join RmRt RSZ on RrTcTyp= 'RSZ' and RrCmCtg= 'C' and RrCmCd= '" + ctSelfCmCd + "' and " + _
              "RrCtg= RmCtg and RrSCtg= RmSCtg and RrSTWGrp= '' and RrCd= '' and RrDmCtg= '' and " + _
              "RrLsCtg= '' and RrLabMCd= '' and RrFrLn<= FrRmSz and RrToLn>= FrRmSz " + _
              "Left Outer Join RmIdSz on RisRmCtg=RmCtg and RisRmId=RmId and RisSz1=FrRmSz " + _
              wDtCndDetXd + wCnd + " and OdTc=(Select PMCd from Param Where PTyp='TC' and PMCd='JCN' and PValue='JT') " + IIF(gs_Partition = ctCurrPrtn, " and OmPrtKey='" + ctCurrPrtn + "' ", "") + _
              IIF(wAddlCnd <> "", " And " + wAddlCnd, "") + " Option (MaxDop 1)"
  moCn.Execute wSqlDetXdTo
  

'  wInvStr = " Union All " + wSqlOpnInvFr + " Union All " + wSqlOpnInvTo + _
'            " Union All " + wSqlDetInvFr + " Union All " + wSqlDetInvTo
''********************* urmi 5/7/03 *********************************'
'
'
'  wSqlStrg = wsqlOpnYyMm + " Union All " + wSqlOpnTdFr + " Union All " + wSqlOpnTdTo + _
'             " Union All " + wSqlDetFr + " Union All " + wSqlDetTo + wFgStr + wInvStr '+ " Order By qDt"
'
'  ADC.RepRecSource = wSqlStrg
  
  '*** (Bef 2.12)
  'adc.RepRecSource = "Select * From #TmpStkLed"
  '*** (Bef 2.12)
  
  
  '*** (Jen 2.12)
  Dim ws_DetGrp As String
  If UCase(ADC("UoVchOpt")) = UCase("VchSrNo") Then
    ws_DetGrp = ",TdCoCd, TdTc, TdYy, TdChr, TdNo, TdSr, TdSrNo "
  ElseIf UCase(ADC("UoVchOpt")) = UCase("VchSr") Then
    ws_DetGrp = ",TdCoCd, TdTc, TdYy, TdChr, TdNo, TdSr "
  ElseIf UCase(ADC("UoVchOpt")) = UCase("Vch") Then
    ws_DetGrp = ",TdCoCd, TdTc, TdYy, TdChr, TdNo "
  End If
  
  '*** jay 2.14 *** (Dt: 28/12/06)
  Dim ws_Desc As String
  ws_Desc = IIF(UCase(ADC("UoVchOpt")) = UCase("VchSr"), _
            " (Select TdDesc from Txnd Td where " + IIF(gs_Partition = ctCurrPrtn, " Td.TdPrtKey='" + ctCurrPrtn + "' and ", "") + " Td.TdCoCd= Tmp.TdCoCd and Td.TdTc= Tmp.TdTc and Td.TdYy= Tmp.TdYy and " + _
            " Td.TdChr= Tmp.TdChr and Td.TdNo= Tmp.TdNO and Td.TdSr= Tmp.TdSr and  Td.TdSrNo= 0) ", _
            " Max(qDesc) ")
  '*** jay 2.14 *** (Dt: 28/12/06)
  
  
  '*** jay 2.14 *** (max(qDesc) as qDesc,)
  '*** jay 2.14 *** (Dt: 28/12/06)
  ' ****** Manali 3.03 - 09/07/08 - RmSzDesc added
  ADC.RepRecSource = "Select max(Grp1) as Grp1, max(Grp2) as Grp2, max(Grp3) as Grp3, " + _
      "DGrp1, DGrp2, DGrp3, max(G1Desc) as G1Desc, max(G2Desc) as G2Desc, max(G3Desc) as G3Desc, " + _
      "max(qFlag) as qFlag, max(RmCd) as RmCd, max(qLotNo) as qLotNo, " + _
      "max(qRmSz) as qRmSz, max(qRmSz2) as qRmSz2, max(qRmSz3) as qRmSz3, " + _
      "max(qRmSzDesc) as qRmSzDesc, max(RrRngSz) as RrRngSz, max(RmCtg) as RmCtg, max(qStnCls) as qStnCls, " + _
      "max(RmSCtg) as RmSCtg, max(qStkRt) as qStkRt, max(qPtrYN) as qPtrYN,Sum(qOpnWt*qRmPurityWt) as qOpnPureWt, " + _
      "sum((qDrWt- qCrWt)* qRmPurityWt) as qPureWt, TdCoCd, TdTc, TdYy, TdChr, TdNo, max(TdSr) as TdSr, " + _
      "max(TdSrNo) as TdSrNo, max(TdBYy) as TdBYy, max(TdBChr) as TdBChr, " + _
      "max(TdBNo) as TdBNo, max(qDt) as qDt, " + ws_Desc + " as qDesc, max(qPrntCd) as qPrntCd, max(qLoc) as qLoc, sum(qOpnQty) as qOpnQty, " + _
      "sum(qDrQty) as qDrQty, sum(qCrQty) as qCrQty, sum(qOpnWt) as qOpnWt, " + _
      "sum(qDrWt) as qDrWt, sum(qCrWt) as qCrWt, max(qFrRmLoc) as qFrRmLoc, " + _
      "max(qToRmLoc) as qToRmLoc, max(qRmRt) as qRmRt, max(qQW) as qQW, max(qQWCst) as qQWCst " + ws_Val + _
      " From #TmpStkLed Tmp Group By DGrp1, DGrp2, DGrp3 " + ws_DetGrp
             
  Rep.TxtDtFr.SetText ADC("UoDtFr")
  Rep.TxtDtTo.SetText ADC("UoDtTo")
             
  Rep.TxtRmCtg.SetText ADC("UoRmCtgFr")
  Rep.TxtLocTyp.SetText ADC("UoLocTypFr")
  Rep.TxtLocFr.SetText ADC("UoIssLocFr")
  Rep.TxtLocTo.SetText ADC("UoIssLocTo")
  Rep.TxtLocSel.SetText ADC("UoIssLocSel")
  
  If UCase(ADC("UoYN3")) = UCase("Y") Then
    wQw = "Qty"
    
    Rep.Fldg3Opn.DecimalPlaces = 0: Rep.Fldg3Opn.RoundingType = crRoundToUnit
    Rep.Fldg2Opn.DecimalPlaces = 0: Rep.Fldg2Opn.RoundingType = crRoundToUnit
    Rep.Fldg1Opn.DecimalPlaces = 0: Rep.Fldg1Opn.RoundingType = crRoundToUnit
    Rep.FldgOpn.DecimalPlaces = 0: Rep.FldgOpn.RoundingType = crRoundToUnit
  
    Rep.Fldg4Cr.DecimalPlaces = 0: Rep.Fldg4Cr.RoundingType = crRoundToUnit
    Rep.Fldg3Cr.DecimalPlaces = 0: Rep.Fldg3Cr.RoundingType = crRoundToUnit
    Rep.Fldg2Cr.DecimalPlaces = 0: Rep.Fldg2Cr.RoundingType = crRoundToUnit
    Rep.Fldg1Cr.DecimalPlaces = 0: Rep.Fldg1Cr.RoundingType = crRoundToUnit
    Rep.FldgCr.DecimalPlaces = 0: Rep.FldgCr.RoundingType = crRoundToUnit
  
    Rep.Fldg4Dr.DecimalPlaces = 0: Rep.Fldg4Dr.RoundingType = crRoundToUnit
    Rep.Fldg3Dr.DecimalPlaces = 0: Rep.Fldg3Dr.RoundingType = crRoundToUnit
    Rep.Fldg2Dr.DecimalPlaces = 0: Rep.Fldg2Dr.RoundingType = crRoundToUnit
    Rep.Fldg1Dr.DecimalPlaces = 0: Rep.Fldg1Dr.RoundingType = crRoundToUnit
    Rep.FldgDr.DecimalPlaces = 0: Rep.FldgDr.RoundingType = crRoundToUnit
  
    Rep.Fldg3Cls.DecimalPlaces = 0: Rep.Fldg3Cls.RoundingType = crRoundToUnit
    Rep.Fldg2Cls.DecimalPlaces = 0: Rep.Fldg2Cls.RoundingType = crRoundToUnit
    Rep.Fldg1Cls.DecimalPlaces = 0: Rep.Fldg1Cls.RoundingType = crRoundToUnit
    Rep.FldgCls.DecimalPlaces = 0: Rep.FldgCls.RoundingType = crRoundToUnit
  ElseIf UCase(ADC("UoYN3")) = UCase("N") Then
    wQw = "Wt"
  End If
  
  '*** (Bef 2.12)
  'GRP_REP.SetFormula Rep, "g4_CrWt", " Sum ({rdo.qCr" + wQw + "}, {@wVchNo})"
  'GRP_REP.SetFormula Rep, "g4_DrWt", " Sum ({rdo.qDr" + wQw + "}, {@wVchNo})"
  '*** (Bef 2.12)
  
  '*** (Jen 2.12)
  GRP_REP.SetFormula Rep, "g4_CrWt", "{rdo.qCr" + wQw + "}"
  GRP_REP.SetFormula Rep, "g4_DrWt", "{rdo.qDr" + wQw + "}"
  '*** (Jen 2.12)
  
 If wb_MonEnd Then
            GRP_REP.SetFormula Rep, "g3_ClsWt", "Sum ({rdo.qOpn" + wQw + "}, {rdo.Grp3})"
            GRP_REP.SetFormula Rep, "g2_ClsWt", "Sum ({rdo.qOpn" + wQw + "}, {rdo.Grp2})"
            GRP_REP.SetFormula Rep, "g1_ClsWt", "Sum ({rdo.qOpn" + wQw + "}, {rdo.Grp1})"
            GRP_REP.SetFormula Rep, "g_ClsWt", "Sum ({rdo.qOpn" + wQw + "})"
            
          If ADC("UOYN5") = "Y" Then
            GRP_REP.SetFormula Rep, "g3_PureWt", "Sum ({rdo.qOpnPureWt}, {rdo.Grp3})"
            GRP_REP.SetFormula Rep, "g2_PureWt", "Sum ({rdo.qOpnPureWt}, {rdo.Grp2})"
            GRP_REP.SetFormula Rep, "g1_PureWt", "Sum ({rdo.qOpnPureWt}, {rdo.Grp1})"
            GRP_REP.SetFormula Rep, "g_PureWt", "Sum ({rdo.qOpnPureWt})"
          End If
  
  Else
            GRP_REP.SetFormula Rep, "g3_OpnWt", "Sum ({rdo.qOpn" + wQw + "}, {rdo.Grp3})"
            GRP_REP.SetFormula Rep, "g2_OpnWt", "Sum ({rdo.qOpn" + wQw + "}, {rdo.Grp2})"
            GRP_REP.SetFormula Rep, "g1_OpnWt", "Sum ({rdo.qOpn" + wQw + "}, {rdo.Grp1})"
            GRP_REP.SetFormula Rep, "g_OpnWt", "Sum ({rdo.qOpn" + wQw + "})"
            
           If ADC("UOYN5") = "Y" Then
            GRP_REP.SetFormula Rep, "g3_PureWt", "Sum ({rdo.qPureWt}, {rdo.Grp3})+Sum ({rdo.qOpnPureWt}, {rdo.Grp3})"
            GRP_REP.SetFormula Rep, "g2_PureWt", "Sum ({rdo.qPureWt}, {rdo.Grp2})+Sum ({rdo.qOpnPureWt}, {rdo.Grp2})"
            GRP_REP.SetFormula Rep, "g1_PureWt", "Sum ({rdo.qPureWt}, {rdo.Grp1})+Sum ({rdo.qOpnPureWt}, {rdo.Grp1})"
            GRP_REP.SetFormula Rep, "g_PureWt", "Sum ({rdo.qPureWt})+Sum ({rdo.qOpnPureWt})"
           End If
  End If
  
  
  GRP_REP.SetFormula Rep, "g3_CrWt", "Sum ({rdo.qCr" + wQw + "}, {rdo.Grp3})"
  GRP_REP.SetFormula Rep, "g2_CrWt", "Sum ({rdo.qCr" + wQw + "}, {rdo.Grp2})"
  GRP_REP.SetFormula Rep, "g1_CrWt", "Sum ({rdo.qCr" + wQw + "}, {rdo.Grp1})"
  GRP_REP.SetFormula Rep, "g_CrWt", "Sum ({rdo.qCr" + wQw + "})"
  GRP_REP.SetFormula Rep, "g3_DrWt", "Sum ({rdo.qDr" + wQw + "}, {rdo.Grp3})"
  GRP_REP.SetFormula Rep, "g2_DrWt", "Sum ({rdo.qDr" + wQw + "}, {rdo.Grp2})"
  GRP_REP.SetFormula Rep, "g1_DrWt", "Sum ({rdo.qDr" + wQw + "}, {rdo.Grp1})"
  GRP_REP.SetFormula Rep, "g_DrWt", "Sum ({rdo.qDr" + wQw + "})"
  
 If wb_ShowOpn Then
      If wb_MonEnd Then
            
            GRP_REP.SetFormula Rep, "g3_OpnWt", "{@g3_ClsWt}-{@g3_DrWt} + {@g3_CrWt}"
            GRP_REP.SetFormula Rep, "g2_OpnWt", "{@g2_ClsWt}- {@g2_DrWt}+ {@g2_CrWt}"
            GRP_REP.SetFormula Rep, "g1_OpnWt", "{@g1_ClsWt}- {@g1_DrWt}+ {@g1_CrWt}"
            GRP_REP.SetFormula Rep, "g_OpnWt", "{@g_ClsWt}- {@g_DrWt}+ {@g_CrWt}"
    
       Else
            GRP_REP.SetFormula Rep, "g3_ClsWt", "{@g3_OpnWt}+ {@g3_DrWt}- {@g3_CrWt}"
            GRP_REP.SetFormula Rep, "g2_ClsWt", "{@g2_OpnWt}+ {@g2_DrWt}- {@g2_CrWt}"
            GRP_REP.SetFormula Rep, "g1_ClsWt", "{@g1_OpnWt}+ {@g1_DrWt}- {@g1_CrWt}"
            GRP_REP.SetFormula Rep, "g_ClsWt", "{@g_OpnWt}+ {@g_DrWt}- {@g_CrWt}"
     End If
  Else
    GRP_REP.SetFormula Rep, "g3_ClsWt", "0"
  GRP_REP.SetFormula Rep, "g2_ClsWt", "0"
  GRP_REP.SetFormula Rep, "g1_ClsWt", "0"
  GRP_REP.SetFormula Rep, "g_ClsWt", "0"
  End If
  
  'KC.3- StkRt-when Show Stock Rate=YES then Stock Rate is visible in report
  'if it is NO then Stock Rate is invisible in report
 
  If ADC("UoYN6") = "Y" Then
    Rep.hqStkRt.SetText "Stock Rt"
    Rep.FldqStkRt.Suppress = False
  Else
    Rep.hqStkRt.SetText ""
    Rep.FldqStkRt.Suppress = True
  End If
  

  '*** Set The Formula For Group4 According To The User Selection (Group4 is the Detail Record)
  If UCase(ADC("UoVchOpt")) = UCase("VchSrNo") Then
    '*** (Bef 15/05/07)
    'GRP_REP.SetFormula Rep, "wVchNo", "Trim ({rdo.TdCoCd})+'/'+({rdo.TdTc})+'/'+" + _
    '"Trim ({rdo.TdYy})+'/'+Trim ({rdo.TdChr})+'/'+Trim (ToText ({rdo.TdNo}, '######0'))+'/'+" + _
    '"Trim (ToText ({rdo.TdSr}, '##0'))+'/'+Trim (ToText ({rdo.TdSrNo}, '#0'))"
    '*** (Bef 15/05/07)
    
    '*** (Jen 2.14 Next Patch 1) (15/05/07)
    GRP_REP.SetFormula Rep, "wVchNo", "Trim ({rdo.TdTc})+'/'+" + _
    "Trim ({rdo.TdYy})+'/'+Trim ({rdo.TdChr})+'/'+Trim (ToText ({rdo.TdNo}, '######0'))+'/'+" + _
    "Trim (ToText ({rdo.TdSr}, '##0'))+'/'+Trim (ToText ({rdo.TdSrNo}, '#0'))"
    '*** (Jen 2.14 Next Patch 1) (15/05/07)
    
  ElseIf UCase(ADC("UoVchOpt")) = UCase("VchSr") Then
  
    '*** (Bef 15/05/07)
    'GRP_REP.SetFormula Rep, "wVchNo", "Trim ({rdo.TdCoCd})+'/'+({rdo.TdTc})+'/'+" + _
    '"Trim ({rdo.TdYy})+'/'+Trim ({rdo.TdChr})+'/'+Trim (ToText ({rdo.TdNo}, '######0'))+'/'+" + _
    '"Trim (ToText ({rdo.TdSr}, '##0'))"
    '*** (Bef 15/05/07)
    
    '*** (Jen 2.14 Next Patch 1) (15/05/07)
    GRP_REP.SetFormula Rep, "wVchNo", "Trim ({rdo.TdTc})+'/'+" + _
    "Trim ({rdo.TdYy})+'/'+Trim ({rdo.TdChr})+'/'+Trim (ToText ({rdo.TdNo}, '######0'))+'/'+" + _
    "Trim (ToText ({rdo.TdSr}, '##0'))"
    '*** (Jen 2.14 Next Patch 1) (15/05/07)
      
    Rep.hqLoc.Suppress = True
    Rep.hRmCtg.Suppress = True
    Rep.hRmSCtg.Suppress = True
    Rep.hRmCd.Suppress = True
    Rep.hRmSz.Suppress = True
    'Rep.hRrRngSz.Suppress = True
    Rep.hqStkRt.Suppress = True
    Rep.hFrRmLoc.Suppress = True
    Rep.hToRmLoc.Suppress = True
      
    Rep.FldqLoc.Suppress = True
    Rep.FldRmCtg.Suppress = True
    Rep.FldRmSCtg.Suppress = True
    Rep.FldRmCd.Suppress = True
    Rep.wRmSz.Suppress = True
    'Rep.FldRrRngSz.Suppress = True
    Rep.FldqStkRt.Suppress = True
    Rep.FldFrRmLoc.Suppress = True
    Rep.FldToRmLoc.Suppress = True
  ElseIf UCase(ADC("UoVchOpt")) = UCase("Vch") Then
    '*** (Bef 15/05/07)
    ''****** Sachin 2.14.0 - [25-11-2006] - [Multi-Company] - Company Code added in the Formula Below
    'GRP_REP.SetFormula Rep, "wVchNo", "Trim ({rdo.TdCoCd})+'/'+({rdo.TdTc})+'/'+" + _
    '"Trim ({rdo.TdYy})+'/'+Trim ({rdo.TdChr})+'/'+Trim (ToText ({rdo.TdNo}, '######0'))"
    '*** (Bef 15/05/07)
    
    '*** (Jen 2.14 Next Patch 1) (15/05/07)
    GRP_REP.SetFormula Rep, "wVchNo", "Trim ({rdo.TdTc})+'/'+" + _
    "Trim ({rdo.TdYy})+'/'+Trim ({rdo.TdChr})+'/'+Trim (ToText ({rdo.TdNo}, '######0'))"
    '*** (Jen 2.14 Next Patch 1) (15/05/07)
  
    Rep.hqLoc.Suppress = True
    Rep.hRmCtg.Suppress = True
    Rep.hRmSCtg.Suppress = True
    Rep.hRmCd.Suppress = True
    Rep.hRmSz.Suppress = True
    'Rep.hRrRngSz.Suppress = True
    Rep.hqStkRt.Suppress = True
    Rep.hFrRmLoc.Suppress = True
    Rep.hToRmLoc.Suppress = True
    Rep.hBagNo.Suppress = True
      
    Rep.FldqLoc.Suppress = True
    Rep.FldRmCtg.Suppress = True
    Rep.FldRmSCtg.Suppress = True
    Rep.FldRmCd.Suppress = True
    Rep.wRmSz.Suppress = True
    'Rep.FldRrRngSz.Suppress = True
    Rep.FldqStkRt.Suppress = True
    Rep.FldFrRmLoc.Suppress = True
    Rep.FldToRmLoc.Suppress = True
    Rep.FldBagNo.Suppress = True
    ' Zubin 212
    Rep.HLotNo.Suppress = True
    Rep.FldLotNo.Suppress = True
    ' Zubin 212
  End If
   
  '2.04-3 urmi RmWt/Qty Closing Val not Proper so change the Selection Formula
  ' also thr Det Section Suppress fromula has been changed
' ########################  Manoj #### Ver: 2.0.6 #### Date: 24/01/2004  ########################
  If ADC("UoPMCdFr") <> "" Then
'  If adc("UoYN1") = "Y" Then
' ###########################################  Manoj  ###########################################
    '****** (Jen 2.12 Patch 1) put the additional check " Or {rdo.qVal} <> 0 " in the RecordSelectionFormula ******'
    If ADC("UoYN3") = "Y" Then
      Rep.RecordSelectionFormula = "{rdo.qDrQty} <> 0 Or {rdo.qCrQty} <> 0 Or {rdo.qOpnQty} <> 0 Or {rdo.qVal} <> 0 "
      
      '*** (Jen 2.14 Next)
      Rep.h1PureWt.Suppress = True
      Rep.h2PureWt.Suppress = True
      Rep.FldqPureWt.Suppress = True
      Rep.FldG1PureWt.Suppress = True
      Rep.FldG2PureWt.Suppress = True
      Rep.FldG3PureWt.Suppress = True
      Rep.FldGPureWt.Suppress = True
      '*** (Jen 2.14 Next)
      
    Else
      '*** (Bef 2.12)
      'Rep.RecordSelectionFormula = "Round({rdo.qDrWt}, 3) <> 0 Or Round({rdo.qCrWt}, 3) <> 0 " + _
      '                         "Or Round({rdo.qOpnWt}, 3) <> 0 "
      '*** (Bef 2.12)
    
      '*** (Jen 2.12)
      Rep.RecordSelectionFormula = "Round({rdo.qDrWt}, 3) <> 0 Or Round({rdo.qCrWt}, 3) <> 0 " + _
                                   "Or Round({rdo.qOpnWt}, 3) <> 0 Or Round({rdo.qPureWt}, 3)<> 0 Or {rdo.qVal} <> 0 "
      '*** (Jen 2.12)
    End If
  Else
    If ADC("UoYN3") = "Y" Then
      '*** (Bef 2.14 Next)
      'Rep.RecordSelectionFormula = " Round({rdo.qDrWt}- {rdo.qCrWt} + {rdo.qOpnWt}, 3) <> 0 " + _
                                   "Or {rdo.qDrQty} <> 0 Or {rdo.qCrQty} <> 0 Or {rdo.qOpnQty} <> 0 Or {rdo.qVal} <> 0 "
      '*** (Bef 2.14 Next)
      
      '*** (Jen 2.14 Next)
      If ADC("UoYN") = "Y" Then
        Rep.RecordSelectionFormula = " Round({rdo.qDrWt}- {rdo.qCrWt} + {rdo.qOpnWt}, 3) <> 0 " + _
                                     "Or {rdo.qDrQty} <> 0 Or {rdo.qCrQty} <> 0 Or {rdo.qOpnQty} <> 0 "
      Else
        Rep.RecordSelectionFormula = " {rdo.qDrQty} <> 0 Or {rdo.qCrQty} <> 0 Or {rdo.qOpnQty} <> 0 "
        
        Rep.h1PureWt.Suppress = True
        Rep.h2PureWt.Suppress = True
        Rep.FldqPureWt.Suppress = True
        Rep.FldG1PureWt.Suppress = True
        Rep.FldG2PureWt.Suppress = True
        Rep.FldG3PureWt.Suppress = True
        Rep.FldGPureWt.Suppress = True
      End If
      '*** (Jen 2.14 Next)
    
    Else
      '*** (Bef 2.12)
      'Rep.RecordSelectionFormula = " Round({rdo.qDrWt}, 3) <> 0 Or Round({rdo.qCrWt}, 3) <> 0 " + _
      '                             "Or Round({rdo.qOpnWt}, 3) <> 0 " + _
      '                             "Or ({rdo.qDrQty} - {rdo.qCrQty} + {rdo.qOpnQty}) <> 0 "
      '*** (Bef 2.12)
    

      '*** (Bef 2.13)
      ''*** (Bef 2.12 Patch 3)
      '''*** (Jen 2.12)
      ''Rep.RecordSelectionFormula = " Round({rdo.qDrWt}, 3) <> 0 Or Round({rdo.qCrWt}, 3) <> 0 " + _
      ''                             "Or Round({rdo.qOpnWt}, 3) <> 0 Or Round({rdo.qPureWt}, 3)<> 0 " + _
      ''                             "Or ({rdo.qDrQty} - {rdo.qCrQty} + {rdo.qOpnQty}) <> 0 Or {rdo.qVal} <> 0 "
      '''*** (Jen 2.12)
      ''*** (Bef 2.12 Patch 3)
      '
      ''*** (Jen 2.12 Patch 3)
      'Rep.RecordSelectionFormula = " Round({rdo.qDrWt}, 3) <> 0 Or Round({rdo.qCrWt}, 3) <> 0 " + _
      '                             "Or Round({rdo.qOpnWt}, 3) <> 0 Or Round({rdo.qPureWt}, 3)<> 0 Or {rdo.qVal} <> 0 "
      ''*** (Jen 2.12 Patch 3)
      '*** (Bef 2.13)


      '*** (Jen 2.13)
      If ADC("UoYN") = "Y" Then
        Rep.RecordSelectionFormula = " Round({rdo.qDrWt}, 3) <> 0 Or Round({rdo.qCrWt}, 3) <> 0 " + _
                                     "Or Round({rdo.qOpnWt}, 3) <> 0 Or Round({rdo.qPureWt}, 3)<> 0 " + _
                                     "Or ({rdo.qDrQty} - {rdo.qCrQty} + {rdo.qOpnQty}) <> 0 "
      Else
        Rep.RecordSelectionFormula = " Round({rdo.qDrWt}, 3) <> 0 Or Round({rdo.qCrWt}, 3) <> 0 " + _
                                     "Or Round({rdo.qOpnWt}, 3) <> 0 Or Round({rdo.qPureWt}, 3)<> 0 "
      End If
      '*** (Jen 2.13)
      
    End If
  End If
 '********************************************************************************
  
  ' Zubin 212
  If Not moCn.RecSeek("select 'x' from Head where HCd= '" + ctSelfCmCd + "' and HCoCd= '" + gs_CoCd + "' and HLotNoYN= 'Y' ") Then
    Rep.HLotNo.Suppress = True: Rep.FldLotNo.Suppress = True
  End If
  ' Zubin 212
  '578.19 calling function to change Rmcode font size
  GRP_REP.SetFormula Rep, "wRmCdFontSz", GetRmcdFontSz("{rdo.RmCd}", 7)
  
  ' ****** Manali 3.03 - 09/07/08 - RmSzDesc
'  GRP_REP.SetFormula Rep, "wRmSz", "If {rdo.qRmSzDesc}<> '' Then {rdo.qRmSzDesc} " + _
                                     "Else (If {rdo.qPtrYN}= 'Y' then ToText({rdo.qRmSz}, 4) Else ToText({rdo.qRmSz}, 3))"
  ' ****** Manali 3.03 - 09/07/08 - RmSzDesc
  
    
  CRV_REP.DisplayGroupTree = False
  
  Call DispReport(ADC) '****** Sachin 3.02 - Display Report ******
  
Exit Sub

errhld:
    moCn.Execute "Delete From #TmpStkLed"
    DispMsg Err.Description, etError
    ADC.RepRecSource = "Select * From #TmpStkLed"
  
End Sub

' ****** Manali 3.5.0 - 21/11/08 - Check Boxes added
Private Sub ChkBoxArr_Click(Index As Integer)
With ChkBoxArr(Index)
  Select Case Index
  Case Is = 1     'Tag = "ShowZWt"
    If .Value = Checked Then ADC("UoYN1") = "Y" Else ADC("UoYN1") = "N"
  Case Is = 2     'Tag = "RevLogic"
    If .Value = Checked Then ADC("UoYN2") = "Y" Else ADC("UoYN2") = "N"
   Case Is = 4     'Tag = "ShowDesc"
    If .Value = Checked Then ADC("UoYN4") = "Y" Else ADC("UoYN4") = "N"
  Case Is = 0     'Tag = "Show Bal"
   
  'If Not (moCn.RecSeek("Select 'x' where convert(varchar(15),DateAdd(Day,-1,DateAdd(Month ,1,'01/'+Cast(Month('" + CStr(adc("UODTTO")) + "') as Varchar(2))+'/'+Cast(Year('" + CStr(adc("UODTTO")) + "') as Varchar(4)))),3)='" + CStr(adc("UODTTO")) + "'") Or CInt(Format(CStr(adc("UODTFr")), "dd") = 1)) Or
'  If Trim(TxtPwd.text) <> moCn.GetFldVal("Select PValue1 from Param where Ptyp = 'SEOPWD' and PmCd = 'SEOPWD'") Then
'     .Value = Unchecked:  Exit Sub
'  End If
  
  If .Value = Checked Then ADC("UoYN5") = "Y" Else ADC("UoYN5") = "N"
  'StkRt-To select this option(Show Stock rate) either Seo user can give their password or
  ' can enter SEOPWD password otherwise this option can't not select
  Case Is = 5
    If .Value = Checked Then
       If Trim(TxtPwd.text) <> moCn.GetFldVal("select IsNull ((select PValue from Param Join UsrAccess on UaUsrCd= PMCd " + _
            "where PTyp= 'USR' and PMCd= '" + gs_UsrCd + "'  and UaMnuCd= '" + ADC.MenuCd + "' and " + _
            "UaSeoYn= 'Y' and PValue<> '' and PValue= '" + Trim(TxtPwd.text) + "' ), " + _
            "Isnull ((Select PValue1 from Param where " + _
            "Ptyp= 'SEOPWD' and PmCd= 'SEOPWD' ), ''))") Then _
            .Value = Unchecked
    End If
    If .Value = Checked Then ADC("UoYN6") = "Y" Else ADC("UoYN6") = "N"
    
  End Select
  
End With
End Sub
' ****** Manali 3.5.0 - 21/11/08 - Check Boxes added

' ****** Manali 3.5.0 - 21/11/08 - Check Boxes added
Private Sub ChkBoxArr_GotFocus(Index As Integer)
  DispMsg ChkBoxArr(Index).ToolTipText, etInfo
End Sub
' ****** Manali 3.5.0 - 21/11/08 - Check Boxes added

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
  '****** Sachin 2.14.0 - [23-11-2006] - [Multi-Company] ******
  
  '****** Sachin 2.14.0 - [23-11-2006] - [Multi-Company] ******
  If ws_MultiCoMod = True Then
    Set ADC.FirNKeyCtl = ADC("UoCoCdFr")
  Else
    Set ADC.FirNKeyCtl = ADC("UoLocTypFr")
  End If
  '****** Sachin 2.14.0 - [23-11-2006] - [Multi-Company] ******
  
  'Set adc.FirNKeyCtl = adc("UoLocTypFr") 'adc("UoCoCdFr")
  Call InitProp(Me)
  
  '*** Parameters Used For Setting Properties Of UsrOpt Fields ***
  gb_RmCtgFor = True
  gb_RmZ = False
  'gs_DmTcTyp = "DM"
  'gs_TxnTcTyp = "BV"
  'gs_OmTcTyp = "SO"
  gs_CmCtg = "C"
  gs_Tbl = "Param"
  gs_PTyp = "WIPVAL"
  'gs_RmGrp = "ORDDET"
  'gs_LabGrp = "ORDDET"
  '*** Parameters Used For Setting Properties Of UsrOpt Fields ***
  
  ' Zubin 212
  If Not moCn.RecSeek("select 'x' from Head where HCd= '" + ctSelfCmCd + "' and HCoCd= '" + gs_CoCd + "' and HLotNoYN= 'Y' ") Then
    LblLotNo.Visible = False: ADC("UoLotNoFr").Visible = False
    ADC("UoLotNoTo").Visible = False: ADC("UoLotNoSel").Visible = False
    ADC("UoLotNoFr") = "": ADC("UoLotNoTo") = "": ADC("UoLotNoSel") = ""
  End If
  ' Zubin 212
  
  '****** Sachin 2.14.0 - [23-11-2006] - [Multi-Company] ******
  If ws_MultiCoMod = True Then
      ADC("UOCOCDFR").Enabled = True: ADC("UOCOCDTO").Visible = True: ADC("UOCOCDSEL").Visible = True
      LblPrntCd.Visible = True: ADC("UOPRNTCDFR").Visible = True: ADC("UOPRNTCDTO").Visible = True: ADC("UOPRNTCDSEL").Visible = True
  Else
      ADC("UOCOCDFR").Enabled = False: ADC("UOCOCDTO").Visible = False: ADC("UOCOCDSEL").Visible = False
      LblPrntCd.Visible = False: ADC("UOPRNTCDFR").Visible = False: ADC("UOPRNTCDTO").Visible = False: ADC("UOPRNTCDSEL").Visible = False
  End If
  '****** Sachin 2.14.0 - [23-11-2006] - [Multi-Company] ******
  
  Call SetGroupSort
  '*** Jay 2.14 *** (SEO Change)
  ms_PwdFlag = moCn.RecSeek("Select PValue1 from Param where PTyp= 'SEOPWD' and PmCd= 'SEOPWD' and PValue1= '' ")
  '*** Jay 2.14 *** (SEO Change)

  '4.1.0.0
  ChkBoxArr(2).Visible = False

End Sub

Private Sub ATXT_FldChange(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String, ByVal pv_NewValue As Variant)
  Dim ws_NewVal As String
  gb_RmCtgFor = True
  
  Select Case UCase(IdName)
  Case Is = UCase("UoPMCdFr"), UCase("UoPmCdTo")                                                '4.1.4
    If UCase(IdName) = UCase("UoPmCdFr") And pv_NewValue <> "STKVAL" Then ADC("UoPMCdTo") = ""  '4.1.4
    gs_Tbl = "Param"
    gs_PTyp = "WIPVAL"
  
  Case Is = UCase("UoLocTypFr")
    ws_NewVal = pv_NewValue
    gs_LocTyp = IIF(ws_NewVal = "", "", "'" + ws_NewVal + "'")
  
  End Select
  
  gb_CoCdFor = False '****** Sachin 2.14.0 - [Multi-Company]
  Call SetProp(Me, IdName, Valid, Cancel, ErrMsg)
End Sub
Private Sub ATXT_FldWhen(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String)
gb_RmCtgFor = True
  Select Case UCase(IdName)
  Case Is = UCase("UoPMCdFr")
'*** Jay 2.14 *** (SEO Change)
' ########################  Manoj #### Ver: 2.0.6 #### Date: 24/01/2004  ########################
'    If TxtPwd.text = "" Then Cancel = True: ErrMsg = "Enter SEO Password to View Values ": Exit Sub
'    If adc("UoYn1") = "N" Then Cancel = True: ErrMsg = "Cannot edit this value as Show Val Option is 'N' ": Exit Sub
' ###########################################  Manoj  ###########################################
    If ms_PwdFlag = False Then Cancel = True: ErrMsg = "Enter SEO Password to View Values ": Exit Sub
'*** Jay 2.14 *** (SEO Change)

    gs_Tbl = "Param"
    gs_PTyp = "WIPVAL"
    gb_CoCdFor = False '****** Sachin 2.14.0 - [Multi-Company]
  
  '*** (Jen 2.13)
  Case Is = UCase("UoYN")
    '*** (Bef 2.14 Next)
    'If adc("UoPMCdFr") <> "" Or adc("UoYN3") = "Y" Then Cancel = True: ErrMsg = "Cannot Edit This Value": Exit Sub
    '*** (Bef 2.14 Next)
    '*** (Jen 2.14 Next)
    If ADC("UoPMCdFr") <> "" Then Cancel = True: ErrMsg = "Cannot Edit This Value": Exit Sub
    '*** (Jen 2.14 Next)
  '*** (Jen 2.13)
  '*** Jay 3.01 **** [LotNo Err]
  Case Is = UCase("UoLotNoFr"), UCase("UoLotNoSel")
    gs_RmLocTyp = ADC("UoLocTypFr")
    gs_Loc = ADC("UoIssLocFr")
  Case Is = UCase("UoLotNoTo")
    gs_RmLocTyp = ADC("UoLocTypFr")
    gs_Loc = ADC("UoIssLocTo")
  ' ******* Manali 350Nxt
  Case Is = UCase("UoLotNoSel")
    'gs_RmLocTyp = adc("UoLocTypFr")
    gs_Loc = ADC("UoIssLocFr")
  ' ******* Manali 350Nxt
  '*** Jay 3.01 **** [LotNo Err]
  
  Case Is = UCase("UoPMCdTo")     '4.1.4
      If ADC.Mode = xNorm Then _
        If UCase(ADC("UoPMCdFr")) <> UCase("STKVAL") Then Cancel = True: ErrMsg = "Can Enter Only if Value Type is 'STKVAL'": Exit Sub
  
  End Select
  
  SetProp Me, IdName, When
End Sub
Private Sub atxt_ReCalculate(Index As Integer, ByVal IdName As String, ByVal OldVal As Variant, ByVal RowNum As Integer, ByVal SourceName As String, ByVal SourceColName As String)
  '*** (Jen 2.13)
  Select Case UCase(IdName)
  Case Is = UCase("UoYN")
    If ADC("UoPMCdFr") = "" Then
      'adc("UoYN") = IIF(adc("UoYN3") = "N", "Y", "N")    '*** (Bef 2.14 Next)
      ADC("UoYN") = "Y"                                  '*** (Jen 2.14 Next)
      OptLastCol(0).Value = True
    Else
      ADC("UoYN") = "N"
      OptLastCol(1).Value = True
    End If
    
'  Case Is = UCase("UoYN5")      '4.1.0.0
'    If Not (moCn.RecSeek("Select 'x' where convert(varchar(15),DateAdd(Day,-1,DateAdd(Month ,1,'01/'+Cast(Month('" + CStr(adc("UODTTO")) + "') as Varchar(2))+'/'+Cast(Year('" + CStr(adc("UODTTO")) + "') as Varchar(4)))),3)='" + CStr(adc("UODTTO")) + "'") Or CInt(Format(CStr(adc("UODTFr")), "dd") = 1)) Then
'        adc("UoYn5") = "N"
'        ChkBoxArr(0).Value = Unchecked
'    End If
    
  End Select
  '*** (Jen 2.13)
End Sub


' ****** Manali 3.5.0 - 21/11/08 - Radio Buttons Added
Private Sub OptLastCol_Click(Index As Integer)
Select Case Index
Case Is = 0
  If ADC("UoPMCdFr") <> "" Then
    ADC("UoYN") = "N": OptLastCol(1).Value = True: DispMsg "Value Type Is Selected", etError:
  Else
    ADC("UoYN") = "Y"
  End If
Case Is = 1
  ADC("UoYN") = "N"
  If ADC("UoPMCdFr") = "" Then _
    OptLastCol(2).Value = True: DispMsg "Value Type Is Blank", etError: Exit Sub
Case Is = 2
  ADC("UoYN") = "N"
  If ADC("UoPMCdFr") <> "" Then
    OptLastCol(1).Value = True: DispMsg "Value Type Is Selected", etError: Exit Sub
  Else
    OptLastCol(2).Value = True
  End If
End Select
End Sub
' ****** Manali 3.5.0 - 21/11/08 - Radio Buttons Added

' ****** Manali 3.5.0 - 21/11/08 - Radio Buttons Added
Private Sub OptLastCol_GotFocus(Index As Integer)
  DispMsg FraOptLastCol.ToolTipText, etInfo
End Sub
' ****** Manali 3.5.0 - 21/11/08 - Radio Buttons Added

' ****** Manali 3.5.0 - 06/11/08 - Radio Buttons Added
Private Sub OptQW_Click(Index As Integer)
Select Case Index
Case Is = 0
   ADC("UoYN3") = "Y"
Case Is = 1
   ADC("UoYN3") = "N"
End Select
End Sub
' ****** Manali 3.5.0 - 06/11/08 - Radio Buttons Added

' ****** Manali 3.5.0 - 06/11/08 - Radio Buttons Added
Private Sub OptQW_GotFocus(Index As Integer)
  DispMsg FraOptQW.ToolTipText, etInfo
End Sub
' ****** Manali 3.5.0 - 06/11/08 - Radio Buttons Added

Private Sub tab_REP_Click(PreviousTab As Integer)
  fra_tabrep(TAB_REP.Tab).Enabled = True
  fra_tabrep(PreviousTab).Enabled = False
End Sub
Private Sub ADC_KeyWhen()
  ADC("UOUSRCD") = ADC.UsrCd
  ADC("UOMNUCD") = ADC.MenuCd
  
  Call EnaDisaCmds(False)
' ########################  Manoj #### Ver: 2.0.6 #### Date: 24/01/2004  ########################
'  Call HideAllFras
' ###########################################  Manoj  ###########################################
End Sub
Private Sub ADC_NonKeyWhen(pr_Cancel As Boolean, pr_ErrMsg As String)

  '****** Sachin 2.14.0 - [25-11-2006] - [Multi-Company] ******
  If ws_MultiCoMod = False Then _
      ADC("UoCoCdFr") = gs_CoCd: ADC("UoCoCdTo") = gs_CoCd: ADC("UoCoCdSel") = "": ADC("UoCoCdFr").Enabled = False
  '****** Sachin 2.14.0 - [25-11-2006] - [Multi-Company] ******

 ' Call SetGroupSort     '4.1.4

  'adc("UoCoCdFr") = gs_CoCd
  ADC("UoOpnYyMm") = 0
  gs_LocTyp = "'" + ADC("UoLocTypFr") + "'"
' ########################  Manoj #### Ver: 2.0.6 #### Date: 24/01/2004  ########################
  TxtPwd.text = ""
  ADC("UoPMCdFr") = ""
  ADC("UoYn5") = "N"    '4.1.0.0
  
  ADC("UOYN1") = "N"
' ###########################################  Manoj  ###########################################

  ADC("UOYN2") = "Y"  '*** (Jen 3.01)
  
  ' ****** Manali 3.5.0 - 06/11/08 - Radio Buttons Added
  If ADC("UoYN3") = "Y" Then OptQW(0).Value = True
  If ADC("UoYN3") = "N" Then OptQW(1).Value = True
  
  If ADC("UoYN") = "Y" Then OptLastCol(0).Value = True
  If ADC("UoYN") = "N" Then OptLastCol(2).Value = True
  ' ****** Manali 3.5.0 - 06/11/08 - Radio Buttons Added
  
  ' ****** Manali 3.5.0 - 21/11/08 - Check Boxes added
  If ADC("UoYN4") = "" Then ADC("UoYN4") = "N"
  If ADC("UoYN1") = "Y" Then ChkBoxArr(1).Value = Checked Else ChkBoxArr(1).Value = Unchecked
  If ADC("UoYN2") = "Y" Then ChkBoxArr(2).Value = Checked Else ChkBoxArr(2).Value = Unchecked
  If ADC("UoYN4") = "Y" Then ChkBoxArr(4).Value = Checked Else ChkBoxArr(4).Value = Unchecked
  If ADC("UoYN5") = "Y" Then ChkBoxArr(0).Value = Checked Else ChkBoxArr(0).Value = Unchecked
  ' ****** Manali 3.5.0 - 21/11/08 - Check Boxes added
  
  'StkRt-Set Show StockRate option is No
  ADC("UoYN6") = "N"
  
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
' ########################  Manoj #### Ver: 2.0.6 #### Date: 24/01/2004  ########################
'**** Shilpa Emr 206 30/10/03 *****/
'Private Sub CmdPwd_Click()
'  Call DispFra(Pwd)
'End Sub

'Private Sub DispFra(ByVal pv_StkLdgFra As en_StkLdgFra)
'  '*** This is the common procedure which is used to toggle between hiding and displaying
'      'the frames (i.e. if the frame is visible then make it invisible and viseversa)
'      'associated with a command button
'  Select Case pv_StkLdgFra
'  Case Is = Pwd
'    If FraPwd.Visible = True Then
'      FraPwd.Visible = False
'      FraPwd.Enabled = False
'      Call EnaDisaCmds(False, CmdPwd)
'      CmdPwd.SetFocus
'    Else
'      FraPwd.Visible = True
'      FraPwd.Enabled = True
'      CmdPwd.SetFocus
'      FraPwd.ZOrder
'      Call EnaDisaCmds(True, CmdPwd)
'    End If
'  End Select
'End Sub
' ###########################################  Manoj  ###########################################
Private Sub EnaDisaCmds(ByVal pv_ShowFra As Boolean, Optional ByVal pv_CmdBut As CommandButton)
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
      pv_CmdBut.BackColor = vbWhite
      pv_CmdBut.FontBold = True
      pv_CmdBut.CausesValidation = True
    Else
      pv_CmdBut.BackColor = &H8000000F
      pv_CmdBut.FontBold = False
      'pv_CmdBut.CausesValidation = False
    End If
  End If
  'FraNKeyAll.Enabled = Not pv_ShowFra
' ########################  Manoj #### Ver: 2.0.6 #### Date: 24/01/2004  ########################
'  CmdPwd.Enabled = Not pv_ShowFra
' ###########################################  Manoj  ###########################################
  If Not pv_CmdBut Is Nothing Then pv_CmdBut.Enabled = True
End Sub
' ########################  Manoj #### Ver: 2.0.6 #### Date: 24/01/2004  ########################
'Private Sub HideAllFras()
  '*** This procedure hides all the frames
'  FraPwd.Visible = False
'  FraPwd.Enabled = False
'End Sub
' ###########################################  Manoj  ###########################################

Private Sub TxtPwd_Validate(Cancel As Boolean)
'*** Jay 2.14 *** (SEO Change)
  If Trim(TxtPwd.text) = "" Then
      ADC("UoPMCdFr") = ""
      If moCn.GetFldVal("Select PValue1 from Param where Ptyp = 'SEOPWD' and PmCd = 'SEOPWD'") = "" Then _
         ms_PwdFlag = True Else ms_PwdFlag = False
  Else
      If Trim(TxtPwd.text) = moCn.GetFldVal("select IsNull ((select PValue from Param Join UsrAccess on UaUsrCd = PMCd " + _
                     "where PTyp= 'USR' and PMCd= '" + gs_UsrCd + "' and UaMnuCd = '" + ADC.MenuCd + "' and " + _
                     "UaSeoYn= 'Y' and PValue= '" + Trim(TxtPwd.text) + "'), " + _
                     "(Select PValue1 from Param where Ptyp= 'SEOPWD' and PmCd= 'SEOPWD' )) ") Then
          ms_PwdFlag = True: Exit Sub
      Else
          ms_PwdFlag = False
          Cancel = True:  DispMsg "Invalid Password / User Does Not Have Seo Rights ", etError
      End If
  End If
'*** Jay 2.14 *** (SEO Change)
' Bef 2.14
'' ########################  Manoj #### Ver: 2.0.6 #### Date: 24/01/2004  ########################
'    If Trim(TxtPwd.text) = "" Then adc("UoPMCdFr") = ""
''    If Trim(TxtPwd.Text) = "" Then ADC("UoYn1") = ""
'' ###########################################  Manoj  ###########################################
'
'    If Trim(TxtPwd.text) = "" Or Trim(TxtPwd.text) = moCn.GetFldVal("Select PValue1 from Param where " + _
'                           "Ptyp = 'SEOPWD' and PmCd = 'SEOPWD'") Then
'      Exit Sub
'    Else
'      Cancel = True
'      DispMsg "Invalid Password", etError
'    End If
' Bef 2.14
End Sub
'**** Shilpa Emr 206 30/10/03 *****/
Private Function RmRate(ByVal mRtFld As String, ByVal mSzFld As String, ByVal mQw As String, Optional ByVal mCurrenCy As String = "", Optional ByVal LotValYN As Boolean = False) As String
  Dim wPureGldCd As String, wPurePlCd As String, wPureSlCd As String
  Dim wGRt As String, wPRt As String, wSRt As String, wDCXARt As String
  ' ***** Manali 3.03 - 24/06/08 - Pd Mod - 'L' added
  Dim wLRt As String, wPurePdCd As String
  ' ***** Manali 3.03 - 24/06/08 - Pd Mod - 'L' added
  
  wPureGldCd = moCn.GetFldVal("Select PValue from param where PTyp= 'RMCTG' and PMCd= 'G'")
  wPurePlCd = moCn.GetFldVal("Select PValue from param where PTyp= 'RMCTG' and PMCd= 'P'")
  wPureSlCd = moCn.GetFldVal("Select PValue from param where PTyp= 'RMCTG' and PMCd= 'S'")
  ' ***** Manali 3.03 - 24/06/08 - Pd Mod - 'L' added
  wPurePdCd = moCn.GetFldVal("Select PValue from param where PTyp= 'RMCTG' and PMCd= 'L'")
  ' ***** Manali 3.03 - 24/06/08 - Pd Mod - 'L' added
  
  'Prev 01/02/05 Geeta  2.10
  ' Goes Wrong when RmCtg='G' and RmSCtg='GX'
  ' And RmCtg='G' and RmSCtg='-' for RmCd='G24' PureMetal Because RmSCtg does not Match
  
'  wGRt = " isnull((select " + mRtFld + " from RmRt where RrTcTyp= 'RM' and RrCmCtg= 'C' and " + _
'         "RrCtg= RmCtg and RrSCtg= RmSCtg and RrCd= '" + wPureGldCd + "' and " + _
'         "RrFrLn <= " + mSzFld + " And RrToLn >= " + mSzFld + " and RrCmCd = '" + ctValueCmCd + "' and " + _
'         "RrDmCtg= '' and RrLabMCd= '') , 0) "
'  wPRt = " isnull((select " + mRtFld + " from RmRt where RrTcTyp= 'RM' and RrCmCtg= 'C' and " + _
'         "RrCtg= RmCtg and RrSCtg= RmSCtg and RrCd= '" + wPurePlCd + "' and " + _
'         "RrFrLn <= " + mSzFld + " And RrToLn >= " + mSzFld + " and RrCmCd = '" + ctValueCmCd + "' and " + _
'         "RrDmCtg= '' and RrLabMCd= '') , 0) "
'  wSRt = " isnull((select " + mRtFld + " from RmRt where RrTcTyp= 'RM' and RrCmCtg= 'C' and " + _
'         "RrCtg= RmCtg and RrSCtg= RmSCtg and RrCd= '" + wPureSlCd + "' and " + _
'         "RrFrLn <= " + mSzFld + " And RrToLn >= " + mSzFld + " and RrCmCd = '" + ctValueCmCd + "' and " + _
'         "RrDmCtg= '' and RrLabMCd= '') , 0) "

  ' Zubin 213
  Dim ws_BaseCurCd As String
  If mCurrenCy = "" Then
    ws_BaseCurCd = moCn.GetFldVal("Select HBaseCurCd from Head where HCd= '" + ctSelfCmCd + "' and " + _
                  "HCoCd= '" + ADC("UoCoCdFr") + "'")
  Else
    ws_BaseCurCd = mCurrenCy
  End If
  ' Zubin 213 (ws_BaseCurCd cnd added)
  wGRt = " isnull((select " + mRtFld + " from RmRt where " + IIF(gs_Partition = ctCurrPrtn, " RrPrtKey='" + ctCurrPrtn + "' and ", "") + " RrTcTyp= 'RM' and RrCmCtg= 'C' and " + _
         "RrCd= '" + wPureGldCd + "' and " + _
         "RrFrLn <= " + mSzFld + " And RrToLn >= " + mSzFld + " and RrCmCd = '" + ctValueCmCd + "' and " + _
         "RrCmCurCd= '" + ws_BaseCurCd + "' and RrDmCtg= '' and RrLabMCd= '') , 0) "
  wPRt = " isnull((select " + mRtFld + " from RmRt where " + IIF(gs_Partition = ctCurrPrtn, " RrPrtKey='" + ctCurrPrtn + "' and ", "") + " RrTcTyp= 'RM' and RrCmCtg= 'C' and " + _
         "RrCd= '" + wPurePlCd + "' and " + _
         "RrFrLn <= " + mSzFld + " And RrToLn >= " + mSzFld + " and RrCmCd = '" + ctValueCmCd + "' and " + _
         "RrCmCurCd= '" + ws_BaseCurCd + "' and RrDmCtg= '' and RrLabMCd= '') , 0) "
  wSRt = " isnull((select " + mRtFld + " from RmRt where " + IIF(gs_Partition = ctCurrPrtn, " RrPrtKey='" + ctCurrPrtn + "' and ", "") + " RrTcTyp= 'RM' and RrCmCtg= 'C' and " + _
         "RrCd= '" + wPureSlCd + "' and " + _
         "RrFrLn <= " + mSzFld + " And RrToLn >= " + mSzFld + " and RrCmCd = '" + ctValueCmCd + "' and " + _
         "RrCmCurCd= '" + ws_BaseCurCd + "' and RrDmCtg= '' and RrLabMCd= '') , 0) "
  ' ********* Manali 3.03 - Pd mod
  wLRt = " isnull((select " + mRtFld + " from RmRt where " + IIF(gs_Partition = ctCurrPrtn, " RrPrtKey='" + ctCurrPrtn + "' and ", "") + " RrTcTyp= 'RM' and RrCmCtg= 'C' and " + _
         "RrCd= '" + wPurePdCd + "' and " + _
         "RrFrLn <= " + mSzFld + " And RrToLn >= " + mSzFld + " and RrCmCd = '" + ctValueCmCd + "' and " + _
         "RrCmCurCd= '" + ws_BaseCurCd + "' and RrDmCtg= '' and RrLabMCd= '') , 0) "
  ' ********* Manali 3.03 - Pd mod
  '*********
  ' Zubin 213
  
  Dim wLotValFldNm As String
  If ws_BaseCurCd = "RS" Then
    wLotValFldNm = "(case when RmQWCst='Q' then LotRsRtByQ  when RmQWCst='W' then LotRsRtByW end) "
  Else
    wLotValFldNm = "(case when RmQWCst='Q' then LotBaseRtByQ when RmQWCst='W' THEN LotBaseRtByW end) "
  End If
  
  ' ***** Manali 3.8.0 - min added - [rectified error - subquery return more than 1 value - in case of rate entered for both sides]
  If LotValYN = True Then
    wDCXARt = " isnull( " + wLotValFldNm + ",isnull((Select min(" + mRtFld + ") from RmRt where " + IIF(gs_Partition = ctCurrPrtn, " RrPrtKey='" + ctCurrPrtn + "' and ", "") + " RrTcTyp='RM' and RrCmCtg= 'C' and " + _
         "RrCtg= RmCtg and RrSCtg= RmSCtg and RrCd= RmCd and RrCmCd = '" + ctValueCmCd + "' and " + _
         "RrCmCurCd= '" + ws_BaseCurCd + "' and " + _
         "RrFrLn <= " + mSzFld + " And RrToLn >= " + mSzFld + " and RrDmCtg= '' and RrLabMCd= '' and " + _
         mRtFld + " <> 0), 0)) "
  Else
      wDCXARt = " isnull((Select min(" + mRtFld + ") from RmRt where " + IIF(gs_Partition = ctCurrPrtn, " RrPrtKey='" + ctCurrPrtn + "' and ", "") + " RrTcTyp='RM' and RrCmCtg= 'C' and " + _
         "RrCtg= RmCtg and RrSCtg= RmSCtg and RrCd= RmCd and RrCmCd = '" + ctValueCmCd + "' and " + _
         "RrCmCurCd= '" + ws_BaseCurCd + "' and " + _
         "RrFrLn <= " + mSzFld + " And RrToLn >= " + mSzFld + " and RrDmCtg= '' and RrLabMCd= '' and " + _
         mRtFld + " <> 0), 0) "
  End If
  'Before 2.13.0
  'RmRate = " (case when RmCtg= 'G' then " + wGRt + " when RmCtg= 'P' then " + wPRt + " " + _
           " when RmCtg= 'S' then " + wSRt + " when RmCtg in ('D', 'C', 'X', 'A') then " + wDCXARt + " else 0 end) "
           
  '****** Sachin 2.13.0 [12-08-2006] [Q/W Module]
  ' ***** Manali 3.03 - 24/06/08 - Pd Mod - 'L' added
  RmRate = " (case when (RmCtg= 'G' And " + mQw + "='W')  then " + wGRt + " when (RmCtg= 'P' And " + mQw + "='W') then " + wPRt + " " + _
           " when (RmCtg= 'S' And " + mQw + "='W') then " + wSRt + " when (RmCtg= 'L' And " + mQw + "='W') then " + wLRt + " " + _
           "when (RmCtg In ('G','P','S', 'L') And " + mQw + "='Q') Or  (RmCtg in ('D', 'C', 'X', 'A')) then " + wDCXARt + " else 0 end) "
  '****** Sachin 2.13.0 [12-08-2006] [Q/W Module]
    
End Function

'StkRt-if password is change to blank, Show StockRate check box will turn into blank
Private Sub TxtPwd_Change()
    If Trim(TxtPwd.text) = "" Then
        ChkBoxArr(5).Value = Unchecked
    End If
End Sub
