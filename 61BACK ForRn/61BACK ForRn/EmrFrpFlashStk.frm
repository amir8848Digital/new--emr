VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{AB2CC32A-E652-42A2-B341-26ECE02F6386}#2.1#0"; "mwfctlforspeed9.ocx"
Begin VB.Form EmrFrpFlashStk 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Flash Stock Report"
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
      Left            =   10920
      TabIndex        =   50
      TabStop         =   0   'False
      Top             =   9660
      Visible         =   0   'False
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
   End
   Begin MwfCtl.MWCTL_DST DST_REP 
      Height          =   570
      Left            =   4440
      TabIndex        =   51
      TabStop         =   0   'False
      Top             =   9630
      Width           =   6360
      _ExtentX        =   11218
      _ExtentY        =   1005
   End
   Begin VB.Frame fra_keys 
      BorderStyle     =   0  'None
      Height          =   465
      Left            =   4875
      TabIndex        =   60
      Top             =   9630
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
      Left            =   6090
      TabIndex        =   59
      Top             =   9645
      Visible         =   0   'False
      Width           =   855
      Begin CRVIEWER9LibCtl.CRViewer9 CRV_REP 
         Height          =   495
         Left            =   480
         TabIndex        =   56
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
      Height          =   9765
      Left            =   0
      TabIndex        =   55
      Top             =   60
      Width           =   15255
      Begin TabDlg.SSTab TAB_REP 
         Height          =   9465
         Left            =   60
         TabIndex        =   54
         Top             =   30
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   16695
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "&Data Range"
         TabPicture(0)   =   "EmrFrpFlashStk.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "fra_tabrep(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "&Group Sort"
         TabPicture(1)   =   "EmrFrpFlashStk.frx":001C
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
               Width           =   8385
               _ExtentX        =   14790
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
               Index           =   2
               Left            =   0
               TabIndex        =   91
               Tag             =   "ShowStkRt"
               ToolTipText     =   "Check To Show Stock Rate"
               Top             =   5810
               Width           =   1935
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
               Left            =   0
               TabIndex        =   82
               Tag             =   "ShowZWt"
               ToolTipText     =   "Check To Show ZPurity Wt"
               Top             =   7560
               Width           =   1905
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "RM Break Up (Box Detail)     "
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
               Index           =   0
               Left            =   45
               TabIndex        =   46
               Tag             =   "ShowBoxDet"
               ToolTipText     =   "Check To Show Raw Material Break Up"
               Top             =   6690
               Width           =   1905
            End
            Begin VB.Frame FraOptQW 
               BorderStyle     =   0  'None
               Height          =   285
               Left            =   1755
               TabIndex        =   88
               ToolTipText     =   "Select either Quantity Or Weight"
               Top             =   7215
               Width           =   4785
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
                  Left            =   2100
                  TabIndex        =   49
                  ToolTipText     =   "Select either Quantity Or Weight"
                  Top             =   0
                  Width           =   1065
               End
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
                  Left            =   0
                  TabIndex        =   48
                  ToolTipText     =   "Select either Quantity Or Weight"
                  Top             =   0
                  Width           =   1365
               End
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   1
               Left            =   4155
               TabIndex        =   10
               ToolTipText     =   "Enter To FrBag Location"
               Top             =   1485
               Width           =   1830
               _ExtentX        =   3228
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UOISSLOCTO"
               IdName          =   "UOISSLOCTO"
               CmpStr          =   "SLoc<= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   2
               Left            =   6540
               TabIndex        =   11
               ToolTipText     =   "Enter FrBag Location Selection"
               Top             =   1485
               Width           =   8400
               _ExtentX        =   14817
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOISSLOCSEL"
               IdName          =   "UOISSLOCSEL"
               CmpStr          =   "SLoc in "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   40
               Left            =   4155
               TabIndex        =   39
               ToolTipText     =   "Enter To Customer (Customer As Mentioned In Rm Master)"
               Top             =   4905
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   4
               DataField       =   "UOCMCDTO"
               IdName          =   "UOCMCDTO"
               CmpStr          =   "RmCmCd <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   41
               Left            =   6540
               TabIndex        =   40
               ToolTipText     =   "Enter Customer Selection (Customer As Mentioned In Rm Master)"
               Top             =   4905
               Width           =   8400
               _ExtentX        =   14817
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
               TabIndex        =   38
               ToolTipText     =   "Enter From Customer (Customer As Mentioned In Rm Master)"
               Top             =   4920
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   4
               DataField       =   "UOCMCDFR"
               IdName          =   "UOCMCDFR"
               CmpStr          =   "RmCmCd >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   46
               Left            =   4155
               TabIndex        =   36
               ToolTipText     =   "Enter To YearMonth"
               Top             =   4350
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               Mask            =   "#####0"
               MaxLength       =   6
               DataType        =   1
               DataField       =   "UOYYMMTO"
               IdName          =   "UOYYMMTO"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   48
               Left            =   1755
               TabIndex        =   35
               ToolTipText     =   "Enter From YearMonth"
               Top             =   4350
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               Mask            =   "#####0"
               MaxLength       =   6
               DataType        =   1
               DataField       =   "UOYYMMFR"
               IdName          =   "UOYYMMFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   0
               Left            =   1755
               TabIndex        =   9
               ToolTipText     =   "Enter From FrBag Location"
               Top             =   1485
               Width           =   1830
               _ExtentX        =   3228
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UOISSLOCFR"
               IdName          =   "UOISSLOCFR"
               CmpStr          =   "SLoc>= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   8
               Left            =   4155
               TabIndex        =   23
               ToolTipText     =   "Enter Raw Material Code To Range"
               Top             =   2925
               Width           =   2385
               _ExtentX        =   4207
               _ExtentY        =   503
               MaxLength       =   16
               DataType        =   4
               DataField       =   "UORMCDTO"
               IdName          =   "UORMCDTO"
               CmpStr          =   "RmCd <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   15
               Left            =   6540
               TabIndex        =   24
               ToolTipText     =   "Enter Raw Material Code Selection"
               Top             =   2925
               Width           =   8400
               _ExtentX        =   14817
               _ExtentY        =   503
               MaxLength       =   255
               DataType        =   4
               DataField       =   "UORMCDSEL"
               IdName          =   "UORMCDSEL"
               CmpStr          =   "RmCd In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   4
               Left            =   4155
               TabIndex        =   17
               ToolTipText     =   "Enter Raw Material Sub Category To Range"
               Top             =   2355
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
               Index           =   5
               Left            =   6540
               TabIndex        =   18
               ToolTipText     =   "Enter Raw Material Sub Category Selection"
               Top             =   2355
               Width           =   8400
               _ExtentX        =   14817
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UORMSCTGSEL"
               IdName          =   "UORMSCTGSEL"
               CmpStr          =   "RmSCtg In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   6
               Left            =   1755
               TabIndex        =   12
               ToolTipText     =   "Enter From Raw Material Category"
               Top             =   1770
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
               TabIndex        =   16
               ToolTipText     =   "Enter Raw Material Sub Category From Range"
               Top             =   2355
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataField       =   "UORMSCTGFR"
               IdName          =   "UORMSCTGFR"
               CmpStr          =   "RmSCtg >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   9
               Left            =   1755
               TabIndex        =   22
               ToolTipText     =   "Enter Raw Material Code From Range"
               Top             =   2925
               Width           =   2385
               _ExtentX        =   4207
               _ExtentY        =   503
               MaxLength       =   16
               DataType        =   4
               DataField       =   "UORMCDFR"
               IdName          =   "UORMCDFR"
               CmpStr          =   "RmCd >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   10
               Left            =   4155
               TabIndex        =   32
               ToolTipText     =   "Enter To Raw Material Len/Sv"
               Top             =   3780
               Width           =   1035
               _ExtentX        =   1826
               _ExtentY        =   503
               Alignment       =   2
               Mask            =   "#0.000"
               MaxLength       =   6
               DataType        =   2
               DataField       =   "UOLN1TO"
               IdName          =   "UOLN1TO"
               CmpStr          =   "SRmSz <= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   11
               Left            =   1755
               TabIndex        =   31
               ToolTipText     =   "Enter From Raw Material Len/Sv"
               Top             =   3780
               Width           =   1005
               _ExtentX        =   1773
               _ExtentY        =   503
               Alignment       =   2
               Mask            =   "#0.000"
               MaxLength       =   6
               DataType        =   2
               DataField       =   "UOLN1FR"
               IdName          =   "UOLN1FR"
               CmpStr          =   "SRmSz >= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   16
               Left            =   4155
               TabIndex        =   34
               ToolTipText     =   "Enter To Cost Rate"
               Top             =   4065
               Width           =   1005
               _ExtentX        =   1773
               _ExtentY        =   503
               Alignment       =   2
               MaxLength       =   0
               DataType        =   2
               DataField       =   "UOCSTRTTO"
               IdName          =   "UOCSTRTTO"
               CmpStr          =   "SRmStkRt<= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   18
               Left            =   1755
               TabIndex        =   33
               ToolTipText     =   "Enter From Cost Rate"
               Top             =   4065
               Width           =   1005
               _ExtentX        =   1773
               _ExtentY        =   503
               Alignment       =   2
               MaxLength       =   0
               DataType        =   2
               DataField       =   "UOCSTRTFR"
               IdName          =   "UOCSTRTFR"
               CmpStr          =   "SRmStkRt>= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   3
               Left            =   1755
               TabIndex        =   6
               ToolTipText     =   "Enter Location Type"
               Top             =   1200
               Width           =   495
               _ExtentX        =   873
               _ExtentY        =   503
               MaxLength       =   2
               DataType        =   4
               DataField       =   "UOLOCTYPFR"
               IdName          =   "UOLOCTYPFR"
               CmpStr          =   "LocTyp>= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   12
               Left            =   2445
               TabIndex        =   45
               ToolTipText     =   "Show Raw Material Break Up? (Y/N)"
               Top             =   6690
               Visible         =   0   'False
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               Enabled         =   0   'False
               DataField       =   "UOYN"
               IdName          =   "UOYN"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   13
               Left            =   1755
               TabIndex        =   42
               ToolTipText     =   "Specify If The Value To Be Shown Should Be Stock Value, Cost Value Or Sale Value ?"
               Top             =   5490
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               DataField       =   "UOPMCDFR"
               IdName          =   "UOPMCDFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   14
               Left            =   4155
               TabIndex        =   29
               ToolTipText     =   "Enter To Raw Material Range Size"
               Top             =   3495
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
               Index           =   17
               Left            =   6540
               TabIndex        =   30
               ToolTipText     =   "Enter Raw Material Range Size Selection"
               Top             =   3495
               Width           =   8400
               _ExtentX        =   14817
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UORNGSZSEL"
               IdName          =   "UORNGSZSEL"
               CmpStr          =   "RrRngSz In "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   19
               Left            =   1755
               TabIndex        =   28
               ToolTipText     =   "Enter From Raw Material Range Size"
               Top             =   3495
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
               TabIndex        =   20
               ToolTipText     =   "Enter To Parent Code"
               Top             =   2640
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
               Index           =   21
               Left            =   6540
               TabIndex        =   21
               ToolTipText     =   "Enter Parent Code Selection"
               Top             =   2640
               Width           =   8400
               _ExtentX        =   14817
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOPARENTCDSEL"
               IdName          =   "UOPARENTCDSEL"
               CmpStr          =   "RmParentCd In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   22
               Left            =   1755
               TabIndex        =   19
               ToolTipText     =   "Enter From Parent Code"
               Top             =   2640
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
               Index           =   30
               Left            =   1755
               TabIndex        =   44
               ToolTipText     =   "Enter The Percentage By Which Value Has To Be Incremented"
               Top             =   6135
               Width           =   855
               _ExtentX        =   1508
               _ExtentY        =   503
               Alignment       =   2
               Mask            =   "-##0.000"
               MaxLength       =   8
               DataType        =   2
               DataField       =   "UOPRC"
               IdName          =   "UOPRC"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   23
               Left            =   1755
               TabIndex        =   0
               ToolTipText     =   "Enter Company Code"
               Top             =   585
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOCOCDFR"
               IdName          =   "UOCOCDFR"
               CmpStr          =   "SCoCd >= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   24
               Left            =   1515
               TabIndex        =   47
               ToolTipText     =   "Enter 'Y' for Quantity and 'N' for Weight"
               Top             =   6855
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
               Index           =   25
               Left            =   1755
               TabIndex        =   37
               ToolTipText     =   "Enter Open YearMonth"
               Top             =   4635
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               Mask            =   "#####0"
               MaxLength       =   6
               DataType        =   1
               DataField       =   "UOOPNYYMM"
               IdName          =   "UOOPNYYMM"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   27
               Left            =   4155
               TabIndex        =   7
               ToolTipText     =   "Enter To Location Type"
               Top             =   1200
               Width           =   495
               _ExtentX        =   873
               _ExtentY        =   503
               MaxLength       =   2
               DataType        =   4
               DataField       =   "UOLOCTYPTO"
               IdName          =   "UOLOCTYPTO"
               CmpStr          =   "LocTyp<= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   28
               Left            =   6540
               TabIndex        =   8
               ToolTipText     =   "Enter FrBag Location Selection"
               Top             =   1200
               Width           =   8400
               _ExtentX        =   14817
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOLOCTYPSEL"
               IdName          =   "UOLOCTYPSEL"
               CmpStr          =   "LocTyp  in "
            End
            Begin VB.TextBox TxtPwd 
               Height          =   285
               IMEMode         =   3  'DISABLE
               Left            =   1755
               MaxLength       =   30
               PasswordChar    =   "*"
               TabIndex        =   41
               Top             =   5205
               Width           =   1905
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   32
               Left            =   1755
               TabIndex        =   25
               ToolTipText     =   "Enter From Lot No."
               Top             =   3210
               Width           =   1905
               _ExtentX        =   3360
               _ExtentY        =   503
               MaxLength       =   12
               DataType        =   4
               DataField       =   "UOLOTNOFR"
               IdName          =   "UOLOTNOFR"
               CmpStr          =   "SLotNo>= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   26
               Left            =   4155
               TabIndex        =   26
               ToolTipText     =   "Enter To Lot No."
               Top             =   3210
               Width           =   1905
               _ExtentX        =   3360
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   4
               DataField       =   "UOLOTNOTO"
               IdName          =   "UOLOTNOTO"
               CmpStr          =   "SLotNo<= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   29
               Left            =   6540
               TabIndex        =   27
               ToolTipText     =   "Enter Lot No. Selection"
               Top             =   3210
               Width           =   8400
               _ExtentX        =   14817
               _ExtentY        =   503
               MaxLength       =   255
               DataType        =   4
               DataField       =   "UOLOTNOSEL"
               IdName          =   "UOLOTNOSEL"
               CmpStr          =   "SLotNo In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   31
               Left            =   4155
               TabIndex        =   14
               ToolTipText     =   "Enter Raw Material Stone Class To Range"
               Top             =   2070
               Width           =   1365
               _ExtentX        =   2408
               _ExtentY        =   503
               MaxLength       =   5
               DataField       =   "UOSTNCLSTO"
               IdName          =   "UOSTNCLSTO"
               CmpStr          =   "IsNull(ST.PSCd,'') <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   33
               Left            =   6540
               TabIndex        =   15
               ToolTipText     =   "Enter Raw Material Stone Class Selection"
               Top             =   2070
               Width           =   8400
               _ExtentX        =   14817
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOSTNCLSSEL"
               IdName          =   "UOSTNCLSSEL"
               CmpStr          =   "IsNull(ST.PSCd,'') In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   34
               Left            =   1755
               TabIndex        =   13
               ToolTipText     =   "Enter Raw Material Stone Class From Range"
               Top             =   2070
               Width           =   1365
               _ExtentX        =   2408
               _ExtentY        =   503
               MaxLength       =   5
               DataField       =   "UOSTNCLSFR"
               IdName          =   "UOSTNCLSFR"
               CmpStr          =   "IsNull(ST.PSCd,'') >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   47
               Left            =   4155
               TabIndex        =   1
               ToolTipText     =   "Enter To Company Code"
               Top             =   600
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOCOCDTO"
               IdName          =   "UOCOCDTO"
               CmpStr          =   "SCoCd <= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   50
               Left            =   6540
               TabIndex        =   2
               ToolTipText     =   "Enter Design Code Selection"
               Top             =   600
               Width           =   8400
               _ExtentX        =   14817
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOCOCDSEL"
               IdName          =   "UOCOCDSEL"
               CmpStr          =   "SCoCd In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   35
               Left            =   4155
               TabIndex        =   4
               ToolTipText     =   "Enter To Parent Code"
               Top             =   900
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
               TabIndex        =   5
               ToolTipText     =   "Enter Parent Code Selection"
               Top             =   900
               Width           =   8400
               _ExtentX        =   14817
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
               TabIndex        =   3
               ToolTipText     =   "Enter From Parent Code"
               Top             =   900
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
               Index           =   36
               Left            =   6540
               TabIndex        =   43
               ToolTipText     =   "Specify If The Value To Be Shown Should Be Stock Value, Cost Value Or Sale Value ?"
               Top             =   5490
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               DataField       =   "UOPMCDTO"
               IdName          =   "UOPMCDTO"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   37
               Left            =   2400
               TabIndex        =   90
               ToolTipText     =   "Show Custom Pure Wt? (Y/N)"
               Top             =   7560
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
               TabIndex        =   92
               ToolTipText     =   "Show Custom Pure Wt? (Y/N)"
               Top             =   5810
               Visible         =   0   'False
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   423
               Enabled         =   0   'False
               DataField       =   "UOYN4"
               IdName          =   "UOYN4"
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Oth Value Type"
               BeginProperty Font 
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
               Left            =   4800
               TabIndex        =   89
               Top             =   5520
               Width           =   1545
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
               TabIndex        =   87
               Top             =   900
               Width           =   1425
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
               TabIndex        =   86
               Top             =   2070
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
               TabIndex        =   85
               Top             =   3210
               Width           =   1425
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
               Height          =   315
               Left            =   60
               TabIndex        =   84
               Top             =   5205
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
               TabIndex        =   83
               Top             =   4635
               Width           =   1425
            End
            Begin VB.Label Label2 
               BackStyle       =   0  'Transparent
               Caption         =   "Show Qty/ Wt"
               BeginProperty Font 
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
               Left            =   60
               TabIndex        =   81
               Top             =   7215
               Width           =   1455
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
               Left            =   60
               TabIndex        =   80
               Top             =   585
               Width           =   1335
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Value Incremental %"
               BeginProperty Font 
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
               Index           =   14
               Left            =   60
               TabIndex        =   79
               Top             =   6165
               Width           =   1515
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
               TabIndex        =   78
               Top             =   2640
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
               TabIndex        =   77
               Top             =   3495
               Width           =   1425
            End
            Begin VB.Label ALBL 
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
               Index           =   11
               Left            =   60
               TabIndex        =   76
               Top             =   5520
               Width           =   1425
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "RM Break Up (Box Detail)"
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
               Height          =   525
               Index           =   10
               Left            =   60
               TabIndex        =   75
               ToolTipText     =   "Location"
               Top             =   6690
               Visible         =   0   'False
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
               TabIndex        =   74
               Top             =   1200
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
               TabIndex        =   73
               Top             =   3780
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
               TabIndex        =   72
               ToolTipText     =   "Location"
               Top             =   1770
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
               TabIndex        =   71
               Top             =   2355
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
               TabIndex        =   70
               Top             =   2925
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
               TabIndex        =   69
               Top             =   4065
               Width           =   1425
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "YyyyMm"
               BeginProperty Font 
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
               TabIndex        =   68
               Top             =   4350
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
               TabIndex        =   67
               Top             =   4920
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
               Left            =   60
               TabIndex        =   66
               Top             =   1485
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
               Left            =   5700
               TabIndex        =   65
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
               Left            =   1755
               TabIndex        =   63
               Top             =   120
               Width           =   1335
            End
         End
      End
   End
End
Attribute VB_Name = "EmrFrpFlashStk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rep As New EmrRepFlashStk
Dim moCn As MwfLib.MDOConnection   '*** (Bef speed) bef 24/11/06
'*** Jay 2.14 *** (SEO Change)
Dim ms_PwdFlag As Boolean
'Enum en_FlshStkFra
'  Pwd = 0
'End Enum
'*** Jay 2.14 *** (SEO Change)

'****** Sachin 2.14.0 - [25-11-2006] - [Multi-Company] ******
Dim ws_MultiCoMod As Boolean
'****** Sachin 2.14.0 - [25-11-2006] - [Multi-Company] ******

Private Sub SetGroupSort()
' ########################  Manoj #### Ver: 2.0.6 #### Date: 24/01/2004  ########################
  GRP_REP.Add "Loc Typ", "LocTyp", "", "", "", "", "", "", "(Select PDesc From Param Where PTyp = 'LOCTYP' And PMCD = LocTyp) "
  GRP_REP.Add "(Typ) Process", "RTrim(LocTyp)+RTrim(SCoCd)+RTrim(SLoc)", "'(' + RTrim(LocTyp) + '/'+RTrim(SCoCd)+') '+ RTrim(SLoc)", "wLoc", "hSLoc", "Company Code", "", "", " LocDesc"
' ###########################################  Manoj  ###########################################
  GRP_REP.Add "Rm Ctg", "Rtrim(RmCtg)", "", "RmCtg", "hRmCtg", "", "", "", "(Select PDesc From Param Where PTyp= 'RMCTG' and PMCd = RmCtg) "
  GRP_REP.Add "Rm SubCtg", "Rtrim(RmCtg)+Rtrim(RmSCtg)", "'('+ RTrim(RmCtg)+ ') '+ RmSCtg", "RmCtg,RmSCtg", "hRmCtg,hRmSCtg", "Rm Ctg", "", "", "(Select PDesc From Param Where PTyp= 'RMSCTG' and PMCd = RmCtg and PSCd= RmSCtg)"
  GRP_REP.Add "Rm Code", "Rtrim(RmCtg)+Rtrim(RmSCtg)+Rtrim(RmCd)", "'('+ RTrim(RmCtg)+'/'+ RmSCtg+ ') '+ RmCd", "RmCtg,RmSCtg,RmCd", "hRmCtg,hRmSCtg,hRmCd", "Rm Ctg,Rm SubCtg,BaseCd + RmCd,Parent Cd", "", "", "SubString(RmDesc, 1, 30) "
  
  GRP_REP.Add "Parent Cd", "Rtrim(RmCtg)+Rtrim(RmSCtg)+(case when RmParentCd= '' then RmCd else RmParentCd end)", "'('+ RTrim(RmCtg)+'/'+ RmSCtg+ ') '+ (case when RmParentCd= '' then RmCd else RmParentCd end)", "RmCtg,RmSCtg", "hRmCtg,hRmSCtg", "Rm Ctg,Rm SubCtg,BaseCd + RmCd", "", "", "(Select PDesc from Param where PTyp= 'PARENTCD' and PMCd= RmParentCd) "
  
  GRP_REP.Add "BaseCd + RmCd", "Rtrim(RmBaseCd)+Rtrim(RmCd)", "'('+ RTrim(RmBaseCd) + ') '+ RmCd", "RmCd", "hRmCd", "Rm Ctg,Rm SubCtg,Rm Code,Parent Cd", "", "", "SubString(RmDesc, 1, 30)"
  'GRP_REP.Add "Rm Sv/Sz", "str(SRmSz,6,3)", "", "SRmSz", "hSRmSz", "Rm Ptr,Rm StkRt,(RmCd) Sv/Sz,(RmCd) StkRt", "", "", ""
  '****************Geeta**************Emr206*********12/04/04
  Dim wsPtrYN As String
  wsPtrYN = "(Select PValue3 from Param where PTyp= 'RMSCTG' and PMCd= RmCtg and PSCd= RmSCtg)"
  'GRP_REP.Add "Rm Sv/Sz", "case when " + wsPtrYN + "='Y' then str(SRmSz,7,4) else str(SRmSz,6,3) end ", "", "SRmSz", "hSRmSz", "Rm Ptr,Rm StkRt,(RmCd) Sv/Sz,(RmCd) StkRt", "", "", ""
  ' ****** Manali 3.03 - 10/07/08 - RmSzDesc added
  GRP_REP.Add "Rm Sv/Sz", "case when " + wsPtrYN + "='Y' then " + GetRmSzDBFldNm("SRmSz", "SRmSz2", "SRmSz3") + " else (case when RisRmSzDesc is null then " + GetRmSzDBFldNm("SRmSz", "SRmSz2", "SRmSz3") + " else RisRmSzDesc end) end ", "", "wRmSz", "hSRmSz", "Rm Ptr,Rm StkRt,(RmCd) Sv/Sz,(RmCd) StkRt", "", "", ""
  '**************
  
  GRP_REP.Add "Rm StkRt", "Str(SRmStkRt,9,2)", "", "SRmStkRt", "hSRmStkRt", "Rm Ptr,Rm Sv/Sz,(RmCd) Sv/Sz,(RmCd) StkRt", "", "", ""
  
  ' *******************************************************************************************************************************************
  '4.1.4
  Dim wPureGldCd As String, wPurePlCd As String, wPureSlCd As String, wPurePdCd As String
  Dim wGRt As String, wPRt As String, wSRt As String, wLRt As String, wDCXARt As String
  Dim ws_BaseCurCd As String, wRate As String
  
  wPureGldCd = moCn.GetFldVal("Select PValue from param where PTyp= 'RMCTG' and PMCd= 'G'")
  wPurePlCd = moCn.GetFldVal("Select PValue from param where PTyp= 'RMCTG' and PMCd= 'P'")
  wPureSlCd = moCn.GetFldVal("Select PValue from param where PTyp= 'RMCTG' and PMCd= 'S'")
  wPurePdCd = moCn.GetFldVal("Select PValue from param where PTyp= 'RMCTG' and PMCd= 'L'")
  
  ws_BaseCurCd = moCn.GetFldVal("Select HBaseCurCd from Head where HCd= '" + ctSelfCmCd + "' and HCoCd= '" + adc("UoCoCdFr") + "'")

  wGRt = " isnull((select RrCstRt from RmRt where " + IIF(gs_Partition = ctCurrPrtn, " RrPrtKey='" + ctCurrPrtn + "' and ", "") + " RrTcTyp= 'RM' and RrCmCtg= 'C' and " + _
                "RrCtg= RmCtg and RrSCtg= RmSCtg and RrCd= '" + wPureGldCd + "' and " + _
                "RrFrLn <= SRmSz And RrToLn >= SRmSz and RrCmCd = '" + ctValueCmCd + "' and " + _
                "RrCmCurCd= '" + ws_BaseCurCd + "' and RrDmCtg= '' and RrLabMCd= '') , 0) "
          
          wPRt = " isnull((select RrCstRt from RmRt where " + IIF(gs_Partition = ctCurrPrtn, " RrPrtKey='" + ctCurrPrtn + "' and ", "") + " RrTcTyp= 'RM' and RrCmCtg= 'C' and " + _
               "RrCtg= RmCtg and RrSCtg= RmSCtg and RrCd= '" + wPurePlCd + "' and " + _
               "RrFrLn <= SRmSz And RrToLn >= SRmSz and RrCmCd = '" + ctValueCmCd + "' and " + _
               "RrCmCurCd= '" + ws_BaseCurCd + "' and RrDmCtg= '' and RrLabMCd= '') , 0) "
          
          wSRt = " isnull((select RrCstRt from RmRt where " + IIF(gs_Partition = ctCurrPrtn, " RrPrtKey='" + ctCurrPrtn + "' and ", "") + " RrTcTyp= 'RM' and RrCmCtg= 'C' and " + _
               "RrCtg= RmCtg and RrSCtg= RmSCtg and RrCd= '" + wPureSlCd + "' and " + _
               "RrFrLn <= SRmSz And RrToLn >= SRmSz and RrCmCd = '" + ctValueCmCd + "' and " + _
               "RrCmCurCd= '" + ws_BaseCurCd + "' and RrDmCtg= '' and RrLabMCd= '') , 0) "
          
          wLRt = " isnull((select RrCstRt from RmRt where " + IIF(gs_Partition = ctCurrPrtn, " RrPrtKey='" + ctCurrPrtn + "' and ", "") + " RrTcTyp= 'RM' and RrCmCtg= 'C' and " + _
               "RrCtg= RmCtg and RrSCtg= RmSCtg and RrCd= '" + wPurePdCd + "' and " + _
               "RrFrLn <= SRmSz And RrToLn >= SRmSz and RrCmCd = '" + ctValueCmCd + "' and " + _
               "RrCmCurCd= '" + ws_BaseCurCd + "' and RrDmCtg= '' and RrLabMCd= '') , 0) "
               
           
          wDCXARt = " isnull((Select RrCstRt from RmRt where " + IIF(gs_Partition = ctCurrPrtn, " RrPrtKey='" + ctCurrPrtn + "' and ", "") + " RrTcTyp='RM' and RrCmCtg= 'C' and " + _
               "RrCtg= RmCtg and RrSCtg= RmSCtg and RrCd = RmCd and RrCmCd = '" + ctValueCmCd + "' and " + _
               "RrCmCurCd= '" + ws_BaseCurCd + "' and RrFrLn <= SRmSz And RrToLn >= SRmSz and RrDmCtg= '' and RrLabMCd= '' and RrCstRt <> 0), 0) "
       
          wGRt = " (Case When RSCtg.PValue1='N' Then " + wGRt + " Else SRmStkRt End) "
          wPRt = " (Case When RSCtg.PValue1='N' Then " + wPRt + " Else SRmStkRt End) "
          wSRt = " (Case When RSCtg.PValue1='N' Then " + wSRt + " Else SRmStkRt End) "
          wLRt = " (Case When RSCtg.PValue1='N' Then " + wLRt + " Else SRmStkRt End) "
          wDCXARt = " (Case When RSCtg.PValue1='N' Then " + wDCXARt + " Else SRmStkRt End) "
        
          wRate = " Str((case when (RmCtg= 'G' And RmQwCst='W') then " + wGRt + _
                  "       when (RmCtg= 'P'  And RmQwCst='W') then " + wPRt + " " + _
                  "       when (RmCtg= 'S' And RmQwCst='W') then " + wSRt + _
                  "       when (RmCtg= 'L' And RmQwCst='W') then " + wLRt + " " + _
                  "       when (RmCtg In ('G','P','S', 'L') And RmQwCst='Q') Or (RmCtg in ('D', 'C', 'X', 'A')) then  " + _
                  wDCXARt + " else 0 end),9,2)"
  
  GRP_REP.Add "Rm StkRt-Cst", wRate, "", "SRmStkRt", "hSRmStkRt", "Rm Ptr,Rm Sv/Sz,(RmCd) Sv/Sz,(RmCd) StkRt", "", "", ""
  ' *******************************************************************************************************************************************
  
  wGRt = " isnull((select RrSalRt from RmRt where " + IIF(gs_Partition = ctCurrPrtn, " RrPrtKey='" + ctCurrPrtn + "' and ", "") + " RrTcTyp= 'RM' and RrCmCtg= 'C' and " + _
                "RrCtg= RmCtg and RrSCtg= RmSCtg and RrCd= '" + wPureGldCd + "' and " + _
                "RrFrLn <= SRmSz And RrToLn >= SRmSz and RrCmCd = '" + ctValueCmCd + "' and " + _
                "RrCmCurCd= '" + ws_BaseCurCd + "' and RrDmCtg= '' and RrLabMCd= '') , 0) "
          
          wPRt = " isnull((select RrSalRt from RmRt where " + IIF(gs_Partition = ctCurrPrtn, " RrPrtKey='" + ctCurrPrtn + "' and ", "") + " RrTcTyp= 'RM' and RrCmCtg= 'C' and " + _
               "RrCtg= RmCtg and RrSCtg= RmSCtg and RrCd= '" + wPurePlCd + "' and " + _
               "RrFrLn <= SRmSz And RrToLn >= SRmSz and RrCmCd = '" + ctValueCmCd + "' and " + _
               "RrCmCurCd= '" + ws_BaseCurCd + "' and RrDmCtg= '' and RrLabMCd= '') , 0) "
          
          wSRt = " isnull((select RrSalRt from RmRt where " + IIF(gs_Partition = ctCurrPrtn, " RrPrtKey='" + ctCurrPrtn + "' and ", "") + " RrTcTyp= 'RM' and RrCmCtg= 'C' and " + _
               "RrCtg= RmCtg and RrSCtg= RmSCtg and RrCd= '" + wPureSlCd + "' and " + _
               "RrFrLn <= SRmSz And RrToLn >= SRmSz and RrCmCd = '" + ctValueCmCd + "' and " + _
               "RrCmCurCd= '" + ws_BaseCurCd + "' and RrDmCtg= '' and RrLabMCd= '') , 0) "
          
          wLRt = " isnull((select RrSalRt from RmRt where " + IIF(gs_Partition = ctCurrPrtn, " RrPrtKey='" + ctCurrPrtn + "' and ", "") + " RrTcTyp= 'RM' and RrCmCtg= 'C' and " + _
               "RrCtg= RmCtg and RrSCtg= RmSCtg and RrCd= '" + wPurePdCd + "' and " + _
               "RrFrLn <= SRmSz And RrToLn >= SRmSz and RrCmCd = '" + ctValueCmCd + "' and " + _
               "RrCmCurCd= '" + ws_BaseCurCd + "' and RrDmCtg= '' and RrLabMCd= '') , 0) "
               
          wDCXARt = " isnull((Select RrSalRt from RmRt where " + IIF(gs_Partition = ctCurrPrtn, " RrPrtKey='" + ctCurrPrtn + "' and ", "") + " RrTcTyp='RM' and RrCmCtg= 'C' and " + _
               "RrCtg= RmCtg and RrSCtg= RmSCtg and RrCd= RmCd and RrCmCd = '" + ctValueCmCd + "' and " + _
               "RrCmCurCd= '" + ws_BaseCurCd + "' and RrFrLn <= SRmSz And RrToLn >= SRmSz and RrDmCtg= '' and RrLabMCd= '' and RrSalRt <> 0), 0) "
       
          wGRt = " (Case When RSCtg.PValue1='N' Then " + wGRt + " Else SRmStkRt End) "
          wPRt = " (Case When RSCtg.PValue1='N' Then " + wPRt + " Else SRmStkRt End) "
          wSRt = " (Case When RSCtg.PValue1='N' Then " + wSRt + " Else SRmStkRt End) "
          wLRt = " (Case When RSCtg.PValue1='N' Then " + wLRt + " Else SRmStkRt End) "
          wDCXARt = " (Case When RSCtg.PValue1='N' Then " + wDCXARt + " Else SRmStkRt End) "
        
          wRate = " Str((case when (RmCtg= 'G' And RmQw='W') then " + wGRt + _
                  "       when (RmCtg= 'P'  And RmQw='W') then " + wPRt + " " + _
                  "       when (RmCtg= 'S' And RmQw='W') then " + wSRt + _
                  "       when (RmCtg= 'L' And RmQw='W') then " + wLRt + " " + _
                  "       when (RmCtg In ('G','P','S', 'L') And RmQwCst='Q') Or (RmCtg in ('D', 'C', 'X', 'A')) then  " + _
                  wDCXARt + " else 0 end),9,2)"
  
  GRP_REP.Add "Rm StkRt-Sal", wRate, "", "SRmStkRt", "hSRmStkRt", "Rm Ptr,Rm Sv/Sz,(RmCd) Sv/Sz,(RmCd) StkRt", "", "", ""
  ' *******************************************************************************************************************************************
  
  '******Geeta*****************Emr206****06/04/04*******
  '**********If UserLevel>6 then don't display the Customer Name
  ' ***** Manali 3.8.0 - EmrFrmMDI.TXT_UCD.text Replaced by gs_UsrCd
  GRP_REP.Add "Cust Code", "RmCmCd", "", "", "", "", "", "", "(Select (case when (Select PNum from Param where Ptyp = 'USR' and PmCd = '" + Trim(gs_UsrCd) + "')  <= 5 then  CmName else CmCd end) From CustMst Where CmCtg= 'C' and CmCd= RmCmCd) "
  'GRP_REP.Add "Cust Code", "RmCmCd", "", "", "", "", "", "", "(Select (case when (Select PNum from Param where Ptyp = 'USR' and PmCd = '" + Trim(EmrFrmMDI.TXT_UCD.text) + "')  <= 5 then  CmName else CmCd end) From CustMst Where CmCtg= 'C' and CmCd= RmCmCd) "
  '****** Sachin 2.14.0 ******
  ' ****** Manali 3.03 - 10/07/08 - RmSzDesc added
  If (moCn.GetFldVal("Select HPtrYN from Head where HCoCd= '" + gs_CoCd + "' and HCd='" + ctSelfCmCd + "'") = "Y" And ws_MultiCoMod = False) Or (ws_MultiCoMod = True) Then _
       GRP_REP.Add "Rm Ptr", "case when " + wsPtrYN + "='Y' then str(SRmSz,7,4) else IsNull(RisRmSzDesc, str(SRmSz,6,3)) end ", "", "wRmSz", "hSRmSz", "Rm Sv/Ln,Rm StkRt,(RmCd) Sv/Sz,(RmCd) StkRt", "", "", ""
     
     ' GRP_REP.Add "Rm Ptr", "case when " + wsPtrYN + "='Y' then str(SRmSz,7,4) else str(SRmSz,6,3) end ", "", "SRmSz", "hSRmSz", "Rm Sv/Sz,Rm StkRt,(RmCd) Sv/Sz,(RmCd) StkRt", "", "", ""
     ' GRP_REP.Add "Rm Ptr", "str(SRmSz,6,3)", "", "SRmSz", "hSRmSz", "Rm Sv/Sz,Rm StkRt,(RmCd) Sv/Sz,(RmCd) StkRt", "", "", ""

  GRP_REP.Add "Range Size", "RmCtg+RmSCtg+IsNull(RrRngSz, '')", "'('+ RTrim(RmCtg)+ ', '+ RmSCtg+ ') '+IsNull(RrRngSz, '')", "RmCtg,RmSCtg,RrRngSz", "hRmCtg,hRmSCtg,hRrRngSz", "Rm Ctg,Rm SubCtg", "", "", ""
  '****************Geeta**************Emr206*********12/04/04
  'GRP_REP.Add "(RmCd) Sv/Sz", "RmCd+str(SRmSz,6,3)", "'('+RmCd+') '+LTrim(str(SRmSz,6,3))", "RmCtg,RmSCtg,RmCd,SRmSz", "hRmCtg,hRmSCtg,hRmCd,hSRmSz", "Rm Code,Rm Ctg,Rm SubCtg,BaseCd + RmCd,(RmCd) StkRt,Rm Sv/Sz,Rm StkRt,Rm Ptr", "", "", " '('+ RTrim(RmCtg)+ '/'+ RmSCtg+ ')'"
  'GRP_REP.Add "(RmCd) Sv/Sz", "RmCd+case when " + wsPtrYN + "='Y' then  str(SRmSz,7,4) else  str(SRmSz,6,3) end ", "'('+RmCd+') '+LTrim(case when " + wsPtrYN + "='Y' then  str(SRmSz,7,4) else  str(SRmSz,6,3) end)", "RmCtg,RmSCtg,RmCd,SRmSz", "hRmCtg,hRmSCtg,hRmCd,hSRmSz", "Rm Code,Rm Ctg,Rm SubCtg,BaseCd + RmCd,(RmCd) StkRt,Rm Sv/Sz,Rm StkRt,Rm Ptr", "", "", " '('+ RTrim(RmCtg)+ '/'+ RmSCtg+ ')'"
  ' ****** Manali 3.03 - 10/07/08 - RmSzDesc added

  GRP_REP.Add "(RmCd) Sv/Sz", "RmCd+case when " + wsPtrYN + "='Y' then " + GetRmSzDBFldNm("SRmSz", "SRmSz2", "SRmSz3") + " else (case when RisRmSzDesc is null then " + GetRmSzDBFldNm("SRmSz", "SRmSz2", "SRmSz3") + " else RisRmSzDesc end) end ", _
                              "'('+RmCd+') '+LTrim(case when " + wsPtrYN + "='Y' then " + GetRmSzDispFldNm("SRmSz", "SRmSz2", "SRmSz3") + " else (case when RisRmSzDesc is null then " + GetRmSzDispFldNm("SRmSz", "SRmSz2", "SRmSz3") + " else RisRmSzDesc end) end)", _
                              "RmCtg,RmSCtg,RmCd,wRmSz", "hRmCtg,hRmSCtg,hRmCd,hSRmSz", "Rm Code,Rm Ctg,Rm SubCtg,BaseCd + RmCd,(RmCd) StkRt,Rm Sv/Sz,Rm StkRt,Rm Ptr", "", "", " '('+ RTrim(RmCtg)+ '/'+ RmSCtg+ ')'"
  GRP_REP.Add "(RmCd) StkRt", "RmCd+str(SRmStkRt,9,2)", "'('+RmCd+') '+LTrim(str(SRmStkRt,9,2))", "RmCtg,RmSCtg,RmCd,SRmStkRt", "hRmCtg,hRmSCtg,hRmCd,hSRmStkRt", "Rm Code,Rm Ctg,Rm SubCtg,BaseCd + RmCd,(RmCd) Sv/Sz,Rm Sv/Sz,Rm StkRt,Rm Ptr", "", "", " '('+ RTrim(RmCtg)+ '/'+ RmSCtg+ ')'"
  ' Zubin 212
  'GRP_REP.Add "YyMm", "SYyMm", "", "SYyMm", "", "", "", "", "SYyMm"

  '****** Sachin 2.13.0 [07-07-2006] ******
  GRP_REP.Add "Stone Class", "RmCtg+IsNull(ST.PSCd, '')", "'('+ RTrim(RmCtg)+') '+IsNull(ST.PSCd, '')", "RmCtg", "hRmCtg", "Rm Ctg", "", "", "ST.PDesc"
  '****** Sachin 2.13.0 [07-07-2006] ******
  
  '****** Sachin 2.14.0 ******
  If (moCn.RecSeek("select 'x' from Head where HCd= '" + ctSelfCmCd + "' and HCoCd= '" + gs_CoCd + "' and HLotNoYN= 'Y' ") And ws_MultiCoMod = False) Or (ws_MultiCoMod = True) Then
  'If moCn.GetFldVal("select Max(HLotNoYn) from Head where HCd= '" + ctSelfCmCd + "' " + IIF(adc("UoCoCdFr") <> "", " and HCoCd>= '" + adc("UoCoCdFr") + "' ", " ") + IIF(adc("UoCoCdTo") <> "", " and HCoCd<= '" + adc("UoCoCdTo") + "' ", " ") + IIF(adc("UoCoCdSel") <> "", " and HCoCd In (" + adc("UoCoCdSel") + ") ", " ")) = "Y" Then
    GRP_REP.Add "RM Lot No.", "SLotNo", "", "SLotNo", "hSLotNo", "", "", "", "SLotNo"
  End If
  ' Zubin 212
  '****** Sachin 2.14.0 ******
  
  '****** Sachin 2.14.0 - [23-11-2006] - [Multi-Company] - ,Process,OrderNo,OrderSr,Customer PO,FgSubLoc ******
  If ws_MultiCoMod = True Then
    GRP_REP.Add "Company Code", "SCoCd", "SCoCd", "", "", "", "", "", "(Select HName from Head where HCd= '" + ctSelfCmCd + "' and HCoCd=SCoCd) "
    GRP_REP.Add "Parent Location", "LocPrntCd", "LocPrntCd", "", "", "", "", "", "(Select LocDesc From Loc Where LocCoCd= '" + ctSelfCoCd + "' and LocCd=Lc.LocPrntCd) "
  End If
  '****** Sachin 2.14.0 - [23-11-2006] - [Multi-Company] - ******
  '2 YyyyMm group sort added
  GRP_REP.Add "YyyyMm", "LTrim(str(SYyMm))", "", "", "", "", "", "", ""

End Sub

Private Sub adc_setreprecsource()
  '*** Report Sql ***
' Zubin 212 (Lot No. added)
' ****Manali - qShowQtyYN added
'  Select space(30) as Grp1, space(30) as Grp2, space(30) as Grp3,
'  space(30) as DGrp1, space(30) as DGrp2, space(30) as DGrp3,
'  space(30) as G1Desc, space(30) as G2Desc, space(30) as G3Desc,
'  space(1) as qFlag, RMCd, SLotNo, SRmSz, RrRngSz, RmCtg, RmSCtg,
'  SRmStkRt, PValue3 as qPtrYN, RmPurityWt,
'  SCoCd, SLoc, 0.0 as qOpnWt, SRmDrWt as qDrWt, SRmCrWt as qCrWt, SYyMm, RrSalRt,
'  SRmDrQty As qOpnQty, SRmDrQty As qDrQty, SRmCrQty As qCrQty, space(1) As qQW, space(1) As qQWCst, Space(1) as qShowQtyYN
'  From SYyMm, RmMst, RmRt, Param where 1=2

  '****** Sachin 3.02.0 - Check on YYYYMM Entered by user ******
  Dim ws_PurgeDt As String
  'sv.9 getting max(PurgeDate) from vParam for given company code
  ws_PurgeDt = moCn.GetFldVal("Select max(convert(date,vPDesc225,3)) from vParam " + _
                          " where vPCoCd >='" + adc("UoCoCdFr") + "' and " + _
                          " vPCoCd <='" + IIF(adc("UoCoCdTo") <> "", adc("UoCoCdTo"), "ZZZ") + "'" + _
                          IIF(adc("UoCoCdSel") <> "", " and vPCoCd In(" + adc("UoCoCdSel") + ")", "") + _
                          " and vPTyp= 'PURGEDT' and vPMCd= 'PURGEDT' and vPSCd =''")
  
  If ws_PurgeDt = "" Then ws_PurgeDt = "01/01/80"
  
  'sv.9 gs_PurgeDt replaced with ws_PurgeDt
  If gs_Partition = ctCurrPrtn And CDate(ws_PurgeDt) <> "01/01/80" And (CDate(adc("UOYYMMFR")) <= Format(CDate(ws_PurgeDt), "YYYYMM") Or CDate(adc("UOYYMMTO")) <= Format(CDate(ws_PurgeDt), "YYYYMM")) Then
      DispMsg "From/To YYYYMM Cannot Be Less Than Purge Date [" + CStr(CDate(ws_PurgeDt)) + "]", etError:      Exit Sub
  End If
  '****** Sachin 3.02.0 - Check on YYYYMM Entered by user ******

  '*** Report Sql ***
  Dim grpflds As String, wrepcnd As String, wCnd As String, wSqlStrg As String, wPurityWt As String
  Dim wPureWtOrZWt As String, wXPureWtOrZwt As String
  Dim wRt As String, wPureGldCd As String, wPurePlCd As String, wPureSlCd As String
  Dim wPurePdCd As String, wLRt As String ' ***** Manali 3.03 - 23/06/08 - Pd Mod - 'L'
  Dim wGRt As String, wPRt As String, wSRt As String, wDCXARt As String, wRate As String
  Dim wQw As String, wOpnYyStr As String, wQwChk As String
  ' ****** Manali 3.03 - 10/07/08 - RmSzDesc
  Dim wRmSzDesc As String, wRmIdSzJoin As String
  ' ****** Manali 3.03 - 10/07/08 - RmSzDesc
  Set Rep = Nothing
  Set adc.RepSource = Rep
  Call DispCoNm
  Call SetMwName(Rep)   ' ***** Manali 3.6.1 - Microway footer
  Call GRP_REP.Gen3LRep
  grpflds = GRP_REP.GrpFldLst
 
  
  If adc("UOYN3") = "Y" Then
    Rep.TxtHead.SetText ("Flash Stock (By Quantity)")
    
    Rep.Fldg4Opn.DecimalPlaces = 0: Rep.Fldg4Opn.RoundingType = crRoundToUnit
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
  
    Rep.Fldg4Cls.DecimalPlaces = 0: Rep.Fldg4Cls.RoundingType = crRoundToUnit
    Rep.Fldg3Cls.DecimalPlaces = 0: Rep.Fldg3Cls.RoundingType = crRoundToUnit
    Rep.Fldg2Cls.DecimalPlaces = 0: Rep.Fldg2Cls.RoundingType = crRoundToUnit
    Rep.Fldg1Cls.DecimalPlaces = 0: Rep.Fldg1Cls.RoundingType = crRoundToUnit
    Rep.FldgCls.DecimalPlaces = 0: Rep.FldgCls.RoundingType = crRoundToUnit
  ElseIf adc("UOYN3") = "N" Then
    Rep.TxtHead.SetText ("Flash Stock (By Weight)")
  End If
  
  ' Zubin 213
  Dim ws_BaseCurCd As String
  ws_BaseCurCd = moCn.GetFldVal("Select HBaseCurCd from Head where HCd= '" + ctSelfCmCd + "' and " + _
                   "HCoCd= '" + adc("UoCoCdFr") + "'")
  
  If UCase(adc("UoYN3")) = UCase("Y") Then
    wQw = "Qty"
  ElseIf UCase(adc("UoYN3")) = UCase("N") Then
    wQw = "Wt"
  End If
  
  If UCase(adc("UoYN1")) = UCase("Y") Then
    Rep.h1PureWt.SetText ("Z Pure Wt")
    Rep.h2PureWt.SetText ("Z Pure Wt")
    wPureWtOrZWt = "RmPurityZ"
    wXPureWtOrZwt = "(Select XZ.RmPurityZ from RmMst XZ where XZ.RmCd=Rm.RmBaseCd " + IIF(gs_Partition = ctCurrPrtn, " And XZ.RmPrtKey=Rm.RmPrtKey)", ")")
  ElseIf UCase(adc("UoYN1")) = UCase("N") Then
    wPureWtOrZWt = "RmPurityWt"
    wXPureWtOrZwt = "RmBasePurityWt"
  End If

  GRP_REP.SetFormula Rep, "wCR", "{rdo.qCr" + wQw + "}"
  GRP_REP.SetFormula Rep, "wDr", "{rdo.qDr" + wQw + "}"
  GRP_REP.SetFormula Rep, "g4_OpnWt", " Sum ({rdo.qOpn" + wQw + "}, {@wGrp4})"
  GRP_REP.SetFormula Rep, "g4_CrWt", " Sum ({rdo.qCr" + wQw + "}, {@wGrp4})"
  GRP_REP.SetFormula Rep, "g4_DrWt", " Sum ({rdo.qDr" + wQw + "}, {@wGrp4})"
  GRP_REP.SetFormula Rep, "g4_CrWt", " Sum ({rdo.qCr" + wQw + "}, {@wGrp4})"
  GRP_REP.SetFormula Rep, "g4_DrWt", " Sum ({rdo.qDr" + wQw + "}, {@wGrp4})"
  GRP_REP.SetFormula Rep, "g3_OpnWt", "Sum ({rdo.qOpn" + wQw + "}, {rdo.Grp3})"
  GRP_REP.SetFormula Rep, "g2_OpnWt", "Sum ({rdo.qOpn" + wQw + "}, {rdo.Grp2})"
  GRP_REP.SetFormula Rep, "g1_OpnWt", "Sum ({rdo.qOpn" + wQw + "}, {rdo.Grp1})"
  GRP_REP.SetFormula Rep, "g_OpnWt", "Sum ({rdo.qOpn" + wQw + "})"
  GRP_REP.SetFormula Rep, "g3_CrWt", "Sum ({rdo.qCr" + wQw + "}, {rdo.Grp3})"
  GRP_REP.SetFormula Rep, "g2_CrWt", "Sum ({rdo.qCr" + wQw + "}, {rdo.Grp2})"
  GRP_REP.SetFormula Rep, "g1_CrWt", "Sum ({rdo.qCr" + wQw + "}, {rdo.Grp1})"
  GRP_REP.SetFormula Rep, "g_CrWt", "Sum ({rdo.qCr" + wQw + "})"
  GRP_REP.SetFormula Rep, "g3_DrWt", "Sum ({rdo.qDr" + wQw + "}, {rdo.Grp3})"
  GRP_REP.SetFormula Rep, "g2_DrWt", "Sum ({rdo.qDr" + wQw + "}, {rdo.Grp2})"
  GRP_REP.SetFormula Rep, "g1_DrWt", "Sum ({rdo.qDr" + wQw + "}, {rdo.Grp1})"
  GRP_REP.SetFormula Rep, "g_DrWt", "Sum ({rdo.qDr" + wQw + "})"
  GRP_REP.SetFormula Rep, "g3_ClsWt", "{@g3_OpnWt}+ {@g3_DrWt}- {@g3_CrWt}"
  GRP_REP.SetFormula Rep, "g2_ClsWt", "{@g2_OpnWt}+ {@g2_DrWt}- {@g2_CrWt}"
  GRP_REP.SetFormula Rep, "g1_ClsWt", "{@g1_OpnWt}+ {@g1_DrWt}- {@g1_CrWt}"
  GRP_REP.SetFormula Rep, "g_ClsWt", "{@g_OpnWt}+ {@g_DrWt}- {@g_CrWt}"
      
  ' ***** Manali 3.03 - 23/06/08 - Pd Mod - 'L' added
  
 wPurityWt = " /(case when Rm.RmCtg in ('G', 'P', 'S', 'L', 'M') then " + _
                          "IsNull((Select " + wPureWtOrZWt + " from Param join RmMst on " + _
                          IIF(gs_Partition = ctCurrPrtn, " RmPrtKey='" + ctCurrPrtn + "' and ", "") + " RmCd= PValue where PTyp= 'RMCTG' and PMCd= Rm.RmCtg), 1) " + _
                 "when Rm.RmCtg= 'X' then " + _
                          "IsNull((Select " + wPureWtOrZWt + " from Param join RmMst on " + _
                          IIF(gs_Partition = ctCurrPrtn, " RmPrtKey='" + ctCurrPrtn + "' and ", "") + " RmCd= PValue where PTyp= 'RMCTG' and " + _
                          "PMCd= IsNull((Select RmCtg From RmMst where " + IIF(gs_Partition = ctCurrPrtn, " RmPrtKey='" + ctCurrPrtn + "' and ", "") + " RmCd= Rm.RmBaseCd), '')), 1) " + _
                 "else 1 end) "
  
  wrepcnd = adc.RepCond
  wCnd = IIF(wrepcnd <> "", " and " + wrepcnd, "")
  
  '*** To check if Detail is selected ***
  Dim wGrp(6) As String, i As Integer, wDetPos As Integer
  wGrp(1) = GRP_REP.Value(gltGroup1)
  wGrp(2) = GRP_REP.Value(gltGroup2)
  wGrp(3) = GRP_REP.Value(gltGroup3)
  wDetPos = 0
  For i = 1 To 6
    If wGrp(i) = "(Detail)" Then
      wDetPos = i: Exit For
    End If
  Next i
  
  '*** If Detail not selected then suppress Box Header, Detail and Footer ***
  If wDetPos = 0 Then Rep.BoxHd.Suppress = True: Rep.BoxDet.Suppress = True
  
  'wRt = IIf(adc("UoYN1") = "N", "RrSalRt", "RrCstRt")
  '*** Code Changed by Shilpa.
  If adc("UOPMCDFR") = "CSTVAL" Then
    wRt = "RrCstRt"
'***urmi 2.04-3 Change in Col Headings
    Rep.lblPgHVal.SetText "CstVal"
    Rep.lblGrpHVal.SetText "CstVal"
  ElseIf adc("UOPMCDFR") = "LOTV-BAS" Then
      wRt = "RrCstRt"
  
    Rep.lblPgHVal.SetText "LotV-BAS"
    Rep.lblGrpHVal.SetText "LotV-BAS"
  ElseIf UCase(adc("UoPMCdFr")) = "LOTV-INR" Then
      wRt = "RrCstRt"
      ws_BaseCurCd = "RS"
    Rep.lblPgHVal.SetText "LotV-INR"
    Rep.lblGrpHVal.SetText "LotV-INR"
  ElseIf adc("UOPMCDFR") = "SALVAL" Then
    wRt = "RrSalRt"
'***urmi 2.04-3 Change in Col Headings
    Rep.lblPgHVal.SetText "SalVal"
    Rep.lblGrpHVal.SetText "SalVal"
    
   '***** urmi 2.04 *************'
  ElseIf UCase(adc("UoPMCdFr")) = "" Then
    wRt = "0"
    If adc("UoYn3") = "Y" Then
      GRP_REP.SetFormula Rep, "wVal", "{rdo.qOpnwt}+ {rdo.qDrWt}- {rdo.qCrWt}"
      Rep.fldG4Val.DecimalPlaces = 3: Rep.fldG4Val.RoundingType = crRoundToThousandth
      Rep.fldG3Val.DecimalPlaces = 3: Rep.fldG3Val.RoundingType = crRoundToThousandth
      Rep.fldG2Val.DecimalPlaces = 3: Rep.fldG2Val.RoundingType = crRoundToThousandth
      Rep.fldG1Val.DecimalPlaces = 3: Rep.fldG1Val.RoundingType = crRoundToThousandth
      Rep.fldGVal.DecimalPlaces = 3: Rep.fldGVal.RoundingType = crRoundToThousandth
      Rep.lblPgHVal.SetText "RM Wt"
      Rep.lblGrpHVal.SetText "RM Wt"
    ElseIf adc("UoYn3") = "N" Then
      GRP_REP.SetFormula Rep, "wVal", "{rdo.qOpnQty}+ {rdo.qDrQty}- {rdo.qCrQty}"
      Rep.fldG4Val.DecimalPlaces = 0: Rep.fldG4Val.RoundingType = crRoundToUnit
      Rep.fldG3Val.DecimalPlaces = 0: Rep.fldG3Val.RoundingType = crRoundToUnit
      Rep.fldG2Val.DecimalPlaces = 0: Rep.fldG2Val.RoundingType = crRoundToUnit
      Rep.fldG1Val.DecimalPlaces = 0: Rep.fldG1Val.RoundingType = crRoundToUnit
      Rep.fldGVal.DecimalPlaces = 0: Rep.fldGVal.RoundingType = crRoundToUnit
      Rep.lblPgHVal.SetText "RM Qty"
      Rep.lblGrpHVal.SetText "RM Qty"
    End If
  '********************************'
  End If
  
  If adc("UOPMCDFR") <> "STKVAL" And adc("UOPMCDFR") <> "" Then
    wPureGldCd = moCn.GetFldVal("Select PValue from param where PTyp= 'RMCTG' and PMCd= 'G'")
    wPurePlCd = moCn.GetFldVal("Select PValue from param where PTyp= 'RMCTG' and PMCd= 'P'")
    wPureSlCd = moCn.GetFldVal("Select PValue from param where PTyp= 'RMCTG' and PMCd= 'S'")
    ' ***** Manali 3.03 - 23/06/08 - Pd Mod - 'L' added below
    wPurePdCd = moCn.GetFldVal("Select PValue from param where PTyp= 'RMCTG' and PMCd= 'L'")
    
    ' Zubin 213 (Cur cd Cnd added)
    wGRt = " isnull((select " + wRt + " from RmRt where " + IIF(gs_Partition = ctCurrPrtn, " RrPrtKey='" + ctCurrPrtn + "' and ", "") + " RrTcTyp= 'RM' and RrCmCtg= 'C' and " + _
         "RrCtg= RmCtg and RrSCtg= RmSCtg and RrCd= '" + wPureGldCd + "' and " + _
         "RrFrLn <= SRmSz And RrToLn >= SRmSz and RrCmCd = '" + ctValueCmCd + "' and " + _
         "RrCmCurCd= '" + ws_BaseCurCd + "' and RrDmCtg= '' and RrLabMCd= '') , 0) "
    wPRt = " isnull((select " + wRt + " from RmRt where " + IIF(gs_Partition = ctCurrPrtn, " RrPrtKey='" + ctCurrPrtn + "' and ", "") + " RrTcTyp= 'RM' and RrCmCtg= 'C' and " + _
         "RrCtg= RmCtg and RrSCtg= RmSCtg and RrCd= '" + wPurePlCd + "' and " + _
         "RrFrLn <= SRmSz And RrToLn >= SRmSz and RrCmCd = '" + ctValueCmCd + "' and " + _
         "RrCmCurCd= '" + ws_BaseCurCd + "' and RrDmCtg= '' and RrLabMCd= '') , 0) "
    wSRt = " isnull((select " + wRt + " from RmRt where " + IIF(gs_Partition = ctCurrPrtn, " RrPrtKey='" + ctCurrPrtn + "' and ", "") + " RrTcTyp= 'RM' and RrCmCtg= 'C' and " + _
         "RrCtg= RmCtg and RrSCtg= RmSCtg and RrCd= '" + wPureSlCd + "' and " + _
         "RrFrLn <= SRmSz And RrToLn >= SRmSz and RrCmCd = '" + ctValueCmCd + "' and " + _
         "RrCmCurCd= '" + ws_BaseCurCd + "' and RrDmCtg= '' and RrLabMCd= '') , 0) "
    wLRt = " isnull((select " + wRt + " from RmRt where " + IIF(gs_Partition = ctCurrPrtn, " RrPrtKey='" + ctCurrPrtn + "' and ", "") + " RrTcTyp= 'RM' and RrCmCtg= 'C' and " + _
         "RrCtg= RmCtg and RrSCtg= RmSCtg and RrCd= '" + wPurePdCd + "' and " + _
         "RrFrLn <= SRmSz And RrToLn >= SRmSz and RrCmCd = '" + ctValueCmCd + "' and " + _
         "RrCmCurCd= '" + ws_BaseCurCd + "' and RrDmCtg= '' and RrLabMCd= '') , 0) "
         
    Dim wLotValFldNm As String
    
    If adc("UOPMCDFR") = "LOTV-BAS" Then
      wLotValFldNm = "(Select (case when RmQWCst='Q' then LotBaseRtByQ when RmQWCst='W' then LotBaseRtByW else 0 end) from LotMst where LotNo = SLotNo) "
    End If
    If adc("UOPMCDFR") = "LOTV-INR" Then
      wLotValFldNm = "(Select (case when RmQWCst='Q' then LotRsRtByQ when RmQWCst='W' then LotRsRtByW else 0 end) from LotMst where LotNo = SLotNo) "
    End If
    
  
    If adc("UOPMCDFR") = "LOTV-BAS" Or UCase(adc("UoPMCdFr")) = "LOTV-INR" Then
        wDCXARt = "isnull(" + wLotValFldNm + ", isnull((Select " + wRt + " from RmRt where " + IIF(gs_Partition = ctCurrPrtn, " RrPrtKey='" + ctCurrPrtn + "' and ", "") + " RrTcTyp='RM' and RrCmCtg= 'C' and " + _
         "RrCtg= RmCtg and RrSCtg= RmSCtg and RrCd= RmCd and RrCmCd = '" + ctValueCmCd + "' and " + _
         "RrCmCurCd= '" + ws_BaseCurCd + "' and RrFrLn <= SRmSz And RrToLn >= SRmSz and RrDmCtg= '' and RrLabMCd= '' and " + _
         wRt + " <> 0), 0)) "
    
    Else
        wDCXARt = " isnull((Select " + wRt + " from RmRt where " + IIF(gs_Partition = ctCurrPrtn, " RrPrtKey='" + ctCurrPrtn + "' and ", "") + " RrTcTyp='RM' and RrCmCtg= 'C' and " + _
         "RrCtg= RmCtg and RrSCtg= RmSCtg and RrCd= RmCd and RrCmCd = '" + ctValueCmCd + "' and " + _
         "RrCmCurCd= '" + ws_BaseCurCd + "' and RrFrLn <= SRmSz And RrToLn >= SRmSz and RrDmCtg= '' and RrLabMCd= '' and " + _
         wRt + " <> 0), 0) "
   End If
    '****** Before 2.13.0
    'wRate = " (case when RmCtg= 'G' then " + wGRt + " when RmCtg= 'P' then " + wPRt + " " + _
          " when RmCtg= 'S' then " + wSRt + " when RmCtg in ('D', 'C', 'X', 'A') then " + wDCXARt + " else 0 end) as RrSalRt "
          
    '****** Sachin 2.13.0 [14/08/2006] - [Q/W Module]
    ' ***** Manali 3.03 - 23/06/08 - Pd Mod - 'L' added
    wRate = " (case when (RmCtg= 'G' And " + IIF(adc("UOPMCDFR") = "CSTVAL", "RmQwCst", "RmQw") + "='W') then " + wGRt + " when (RmCtg= 'P'  And " + IIF(adc("UOPMCDFR") = "CSTVAL", "RmQwCst", "RmQw") + "='W') then " + wPRt + " " + _
            " when (RmCtg= 'S' And " + IIF(adc("UOPMCDFR") = "CSTVAL", "RmQwCst", "RmQw") + "='W') then " + wSRt + " when (RmCtg= 'L' And " + IIF(adc("UOPMCDFR") = "CSTVAL", "RmQwCst", "RmQw") + "='W') then " + wLRt + " " + _
            " when (RmCtg In ('G','P','S', 'L') And " + IIF(adc("UOPMCDFR") = "CSTVAL", "RmQwCst", "RmQw") + "='Q') Or (RmCtg in ('D', 'C', 'X', 'A')) then " + wDCXARt + " else 0 end) as RrSalRt "
    '****** Sachin 2.13.0 [14/08/2006] - [Q/W Module]
          
  ElseIf adc("UOPMCDFR") = "STKVAL" Then
    wRate = "SRmStkRt as RrSalRt "
'*** urmi 2.04-3 Change in Col Headings
    Rep.lblPgHVal.SetText "StkVal"
    Rep.lblGrpHVal.SetText "StkVal"
    
    '4.1.4
    If Trim(adc("UoPMCdTo")) <> "STKVAL" Then
        Dim wQwStr As String
        
          wRt = IIF(UCase(adc("UOPMCDTO")) = "CSTVAL", "RrCstRt", "RrSalRt")
          Rep.lblPgHVal.SetText IIF(UCase(adc("UOPMCDTO")) = "CSTVAL", "Stk-CstVal", "Stk-SalVal")
          Rep.lblGrpHVal.SetText IIF(UCase(adc("UOPMCDTO")) = "CSTVAL", "Stk-CstVal", "Stk-SalVal")

          wGRt = " isnull((select " + wRt + " from RmRt where " + IIF(gs_Partition = ctCurrPrtn, " RrPrtKey='" + ctCurrPrtn + "' and ", "") + " RrTcTyp= 'RM' and RrCmCtg= 'C' and " + _
                "RrCtg= RmCtg and RrSCtg= RmSCtg and RrCd= '" + wPureGldCd + "' and " + _
                "RrFrLn <= SRmSz And RrToLn >= SRmSz and RrCmCd = '" + ctValueCmCd + "' and " + _
                "RrCmCurCd= '" + ws_BaseCurCd + "' and RrDmCtg= '' and RrLabMCd= '') , 0) "
          
          wPRt = " isnull((select " + wRt + " from RmRt where " + IIF(gs_Partition = ctCurrPrtn, " RrPrtKey='" + ctCurrPrtn + "' and ", "") + " RrTcTyp= 'RM' and RrCmCtg= 'C' and " + _
               "RrCtg= RmCtg and RrSCtg= RmSCtg and RrCd= '" + wPurePlCd + "' and " + _
               "RrFrLn <= SRmSz And RrToLn >= SRmSz and RrCmCd = '" + ctValueCmCd + "' and " + _
               "RrCmCurCd= '" + ws_BaseCurCd + "' and RrDmCtg= '' and RrLabMCd= '') , 0) "
          
          wSRt = " isnull((select " + wRt + " from RmRt where " + IIF(gs_Partition = ctCurrPrtn, " RrPrtKey='" + ctCurrPrtn + "' and ", "") + " RrTcTyp= 'RM' and RrCmCtg= 'C' and " + _
               "RrCtg= RmCtg and RrSCtg= RmSCtg and RrCd= '" + wPureSlCd + "' and " + _
               "RrFrLn <= SRmSz And RrToLn >= SRmSz and RrCmCd = '" + ctValueCmCd + "' and " + _
               "RrCmCurCd= '" + ws_BaseCurCd + "' and RrDmCtg= '' and RrLabMCd= '') , 0) "
          
          wLRt = " isnull((select " + wRt + " from RmRt where " + IIF(gs_Partition = ctCurrPrtn, " RrPrtKey='" + ctCurrPrtn + "' and ", "") + " RrTcTyp= 'RM' and RrCmCtg= 'C' and " + _
               "RrCtg= RmCtg and RrSCtg= RmSCtg and RrCd= '" + wPurePdCd + "' and " + _
               "RrFrLn <= SRmSz And RrToLn >= SRmSz and RrCmCd = '" + ctValueCmCd + "' and " + _
               "RrCmCurCd= '" + ws_BaseCurCd + "' and RrDmCtg= '' and RrLabMCd= '') , 0) "
               
          wDCXARt = " isnull((Select " + wRt + " from RmRt where " + IIF(gs_Partition = ctCurrPrtn, " RrPrtKey='" + ctCurrPrtn + "' and ", "") + " RrTcTyp='RM' and RrCmCtg= 'C' and " + _
               "RrCtg= RmCtg and RrSCtg= RmSCtg and RrCd= RmCd and RrCmCd = '" + ctValueCmCd + "' and " + _
               "RrCmCurCd= '" + ws_BaseCurCd + "' and RrFrLn <= SRmSz And RrToLn >= SRmSz and RrDmCtg= '' and RrLabMCd= '' and " + _
               wRt + " <> 0), 0) "
       
          wGRt = " (Case When RSCtg.PValue1='N' Then " + wGRt + " Else SRmStkRt End) "
          wPRt = " (Case When RSCtg.PValue1='N' Then " + wPRt + " Else SRmStkRt End) "
          wSRt = " (Case When RSCtg.PValue1='N' Then " + wSRt + " Else SRmStkRt End) "
          wLRt = " (Case When RSCtg.PValue1='N' Then " + wLRt + " Else SRmStkRt End) "
          wDCXARt = " (Case When RSCtg.PValue1='N' Then " + wDCXARt + " Else SRmStkRt End) "
        
          wQwStr = IIF(adc("UOPMCDTO") = "CSTVAL", "RmQwCst", "RmQw")
       
         wRate = " (case when (RmCtg= 'G' And " + wQwStr + "='W') then " + wGRt + _
                  " when (RmCtg= 'P'  And " + wQwStr + "='W') then " + wPRt + " " + _
                  " when (RmCtg= 'S' And " + wQwStr + "='W') then " + wSRt + _
                  " when (RmCtg= 'L' And " + wQwStr + "='W') then " + wLRt + " " + _
                  " when (RmCtg In ('G','P','S', 'L') And " + wQwStr + "='Q') Or (RmCtg in ('D', 'C', 'X', 'A')) then  " + _
                  wDCXARt + " else 0 end) as RrSalRt "
    End If
  
  ElseIf adc("UOPMCDFR") = "" Then
    wGRt = "0"
    wPRt = "0"
    wSRt = "0"
    wDCXARt = "0"
    wRate = "0 as RrSalRt"
  End If
  
  If adc("UoOpnYyMm") <> 0 Then
    wOpnYyStr = " and SyyMm >= " + CStr(adc("UoOpnYyMM")) + ""
  Else
    wOpnYyStr = ""
  End If
  
  ' ****** Sachin 2.13.0 - 05/06/2006 ******
  ' ***** Manali 3.03 - 23/06/08 - Pd Mod - 'L' added
    wQwChk = " IsNull((Case when (RmCtg in ('G', 'P', 'S', 'L', 'M') And RmQw='W') then (Select RmPr.RmQw From RmMst RmPr " + _
              "   Where " + IIF(gs_Partition = ctCurrPrtn, " RmPr.RmPrtKey='" + ctCurrPrtn + "' and ", "") + " RmPr.RmCd=(Select PValue from param " + _
              "       where PTyp= 'RMCTG' and PMCd= Rm.RmCtg)) Else RmQw End),'')  as qQw," + _
              " IsNull((Case when (RmCtg in ('G', 'P', 'S', 'L', 'M') And RmQwCst='W') then (Select RmPr.RmQwCst From RmMst RmPr " + _
              "   Where " + IIF(gs_Partition = ctCurrPrtn, " RmPr.RmPrtKey='" + ctCurrPrtn + "' and ", "") + " RmPr.RmCd=(Select PValue from param " + _
              "       Where PTyp= 'RMCTG' and PMCd= Rm.RmCtg)) Else RmQwCst End),'') as qQwCst "
  ' ****** Sachin 2.13.0 - 05/06/2006 ******
  ' ****** Manali 3.03 - 10/07/08 - RmSzDesc
  wRmSzDesc = "(case when (Select PDesc225 from Param where PTyp= 'RMSCTG' and PMCd= IsNull(RisRmCtg, '') and PSCd= RmSCtg)= 'Y' then IsNull(RisRmSzDesc, '') Else '' End )"
  wRmIdSzJoin = "Left Outer Join RmIdSz on RisRmCtg=RmCtg and RisRmId=RmId and RisSz1=SRmSz "
  ' ****** Manali 3.03 - 10/07/08 - RmSzDesc
  ' Zubin 212
  ' ***** Manali 3.03 - 23/06/08 - Pd Mod - 'L' added
  ' ****** Manali 3.03 - 10/07/08 - RmSzDesc added
  ' ****** Manali 3.4.1 - 28/03/09 - Condition added [SRmDrQty<> 0 or SRmCrQty<> 0] and adc("UoYN3") passed in query as qShowQtyYN
  'JewelSmith.2 ',' + ST.PDesc225 + ',' like '%,'+ RmSCtg +',%' changed as ST.PSCd = RmStnCls
  '2 qGrpHd1,qGrpHd2,qGrpHd3 added
  wSqlStrg = "Select " + grpflds + ", " + _
             "'O' as qFlag, RMCd, SLotNo, SRmSz, SRmSz2, SRmSz3, " + wRmSzDesc + " as qRmSzDesc, IsNull(RrRngSz, '') as RrRngSz, RmCtg, Isnull(ST.PSCd,'') As StnCls, RmSCtg, SRmStkRt, " + _
             "(Select PValue3 from Param where PTyp= 'RMSCTG' and PMCd= RmCtg and PSCd= RmSCtg) as qPtrYN, " + _
             "(Case when RmCtg in ('G', 'P', 'S', 'L', 'M') then " + wPureWtOrZWt + " when RmCtg= 'X' then " + wXPureWtOrZwt + " else 0 end)" + wPurityWt + " as RmPurityWt, SCoCd, SLoc, LocPrntCd, " + _
             "(SRmDrWt- SRmCrWt) as qOpnWt, 0.0 as qDrWt, 0.0 as qCrWt, 0 as SYyMm, " + _
             wRate + ", (SRmDrQty- SRmCrQty) as qOpnQty, 0 as qDrQty, 0 as qCrQty, " + wQwChk + ", '" + adc("UoYN3") + "' as qShowQtyYN " + _
             ",'" + Rep.tGH1.text + "' as qGrpHd1,'" + Rep.tGH2.text + "' as qGrpHd2,'" + Rep.tGH3.text + "' as qGrpHd3 " + _
             "From SYyMm Join RmMst Rm On RmCd= SRmCd " + IIF(gs_Partition = ctCurrPrtn, " and RmPrtKey=SPrtKey ", "") + _
             " Join Param RSCtg On RSCtg.PTyp='RMSCTG' And RSCtg.PMCd=RmCtg And RSCtg.PSCd=RmSCtg " + _
             "Join Loc Lc on LocCoCd= SCoCd and LocCd= SLoc and LocHideYN='N' Left Outer Join Param ST on ST.PTyp= 'STNCLS' and ST.PMCD=RmCtg and ST.PSCd=RmStnCls  " + _
             "Left Outer Join RmRt RSZ on RrTcTyp= 'RSZ' and RrCmCtg= 'C' and RrCmCd= '" + ctSelfCmCd + "' and " + _
             "RrCtg= RmCtg and RrSCtg= RmSCtg and RrSTWGrp= '' and RrCd= '' and RrDmCtg = '' and " + _
             "RrLsCtg= '' and RrLabMCd= '' and RrFrLn<= SRmSz and RrToLn>= SRmSz " + wRmIdSzJoin + _
             "Where " + IIF(gs_Partition = ctCurrPrtn, " SPrtKey='" + ctCurrPrtn + "' and ", " SYyMm<>0 and ") + " SYyMm< " + CStr(adc("UoYyMmFr")) + " And (SRmDrWt <> 0 Or SRmCrWt <> 0 Or SRmDrQty<> 0 or SRmCrQty<> 0) " + wOpnYyStr + wCnd
  '2 qGrpHd1,qGrpHd2,qGrpHd3 added
  wSqlStrg = wSqlStrg + " Union All Select " + grpflds + ", " + _
             "'T' as qFlag, RMCd, SLotNo, SRmSz, SRmSz2, SRmSz3, " + wRmSzDesc + " as qRmSzDesc, IsNull(RrRngSz, '') as RrRngSz, RmCtg, Isnull(ST.PSCd,'') As StnCls, RmSCtg, SRmStkRt, " + _
             "(Select PValue3 from Param where PTyp= 'RMSCTG' and PMCd= RmCtg and PSCd= RmSCtg) as qPtrYN, " + _
             "(Case when RmCtg in ('G', 'P', 'S', 'L', 'M') then " + wPureWtOrZWt + " when RmCtg= 'X' then " + wXPureWtOrZwt + " else 0 end)" + wPurityWt + " as RmPurityWt, SCoCd, SLoc, LocPrntCd, " + _
             "0.0 as qOpnWt, SRmDrWt as qDrWt, SRmCrWt as qCrWt, SYyMm, " + wRate + _
             ", 0 as qOpnQty, SRmDrQty as qDrQty, SRmCrQty as qCrQty , " + wQwChk + ", '" + adc("UoYN3") + "' as qShowQtyYN " + _
             ",'" + Rep.tGH1.text + "' as qGrpHd1,'" + Rep.tGH2.text + "' as qGrpHd2,'" + Rep.tGH3.text + "' as qGrpHd3 " + _
             " From SYyMm Join RmMst Rm On RmCd= SRmCd " + IIF(gs_Partition = ctCurrPrtn, " and RmPrtKey=SPrtKey ", "") + _
             " Join Param RSCtg On RSCtg.PTyp='RMSCTG' And RSCtg.PMCd=RmCtg And RSCtg.PSCd=RmSCtg " + _
             "Join Loc Lc on LocCoCd= SCoCd and LocCd= SLoc and LocHideYN='N' Left Outer Join Param ST on ST.PTyp= 'STNCLS' and ST.PmCd= RmCtg and ST.PSCd=RmStnCls  " + _
             "Left Outer Join RmRt RSZ on RrTcTyp= 'RSZ' and RrCmCtg= 'C' and RrCmCd= '" + ctSelfCmCd + "' and " + _
             "RrCtg= RmCtg and RrSCtg= RmSCtg and RrSTWGrp= '' and RrCd= '' and RrDmCtg= '' and " + _
             "RrLsCtg= '' and RrLabMCd= '' and RrFrLn<= SRmSz and RrToLn>= SRmSz " + wRmIdSzJoin + _
             "Where " + IIF(gs_Partition = ctCurrPrtn, " SPrtKey='" + ctCurrPrtn + "' and ", " SYyMm<>0 and ") + " SYyMm>=" + CStr(adc("UoYyMmFr")) + " and " + _
             "SYyMm<=" + CStr(adc("UoYyMmTo")) + " and (SRmDrWt<> 0 or SRmCrWt<> 0 Or SRmDrQty<> 0 or SRmCrQty<> 0) " + wOpnYyStr + wCnd
  
  
  
  'Debug.Print ws
'  wSqlStrg = "Select " + grpflds + ", " + _
'             "'O' as qFlag, RMCd, SRmSz, RmCtg, RmSCtg, SRmStkRt, RmPurityWt" + wPurityWt + " , SLoc, " + _
'             "(SRmDrWt- SRmCrWt) as qOpnWt, 0.0 as qDrWt, 0.0 as qCrWt, 0 as SYyMm, " + _
'             wRate + " From SYyMm Join RmMst Rm On RmCd= SRmCd " + _
'             "Join Loc on LocCd= SLoc " + _
'             "Where SYyMm< " + CStr(adc("UoYyMmFr")) + _
'             " And (SRmDrWt <> 0 Or SRmCrWt <> 0) " + wCnd + _
'             " Union All " + _
'             " Select " + grpflds + ", " + _
'             "'T' as qFlag, RMCd, SRmSz, RmCtg, RmSCtg, SRmStkRt, RmPurityWt" + wPurityWt + " , SLoc, " + _
'             "0.0 as qOpnWt, SRmDrWt as qDrWt, SRmCrWt as qCrWt, SYyMm, " + wRate + _
'             " From SYyMm Join RmMst Rm On RmCd= SRmCd " + _
'             "Join Loc on LocCd= SLoc " + _
'             "Where SYyMm>=" + CStr(adc("UoYyMmFr")) + " and " + _
'             "SYyMm<=" + CStr(adc("UoYyMmTo")) + _
'             " and (SRmDrWt<> 0 or SRmCrWt<> 0) " + wCnd

'*** Old Code
'  wSqlStrg = "Select " + grpflds + ", " + _
'             "'O' as qFlag, RMCd, SRmSz, RmCtg, RmSCtg, SRmStkRt, RmPurityWt" + wPurityWt + " , SLoc, " + _
'             "(SRmDrWt- SRmCrWt) as qOpnWt, 0.0 as qDrWt, 0.0 as qCrWt, 0 as SYyMm " + _
'             "From SYyMm Join RmMst Rm On RmCd= SRmCd " + _
'             "Join Loc on LocCd= SLoc " + _
'             "Where SYyMm< " + CStr(ADC("UoYyMmFr")) + _
'             " And (SRmDrWt <> 0 Or SRmCrWt <> 0) " + wCnd + _
'             "Union All " + _
'             "Select " + grpflds + ", " + _
'             "'T' as qFlag, RMCd, SRmSz, RmCtg, RmSCtg, SRmStkRt, RmPurityWt" + wPurityWt + " , SLoc, " + _
'             "0.0 as qOpnWt, SRmDrWt as qDrWt, SRmCrWt as qCrWt, SYyMm " + _
'             "From SYyMm Join RmMst Rm On RmCd= SRmCd " + _
'             "Join Loc on LocCd= SLoc " + _
'             "Where SYyMm>=" + CStr(ADC("UoYyMmFr")) + " and SYyMm<=" + CStr(ADC("UoYyMmTo")) + _
'             " and (SRmDrWt<> 0 or SRmCrWt<> 0) " + wCnd
  
  adc.RepRecSource = wSqlStrg + ctMaxDopOpt  '****** Sachin 3.01 ctMaxDopOpt added
  
' ########################  Manoj #### Ver: 2.0.6 #### Date: 24/01/2004  ########################
  If adc("UoPMCdFr") <> "" Then
'  If ADC("UoYn1") = "Y" Then
' ###########################################  Manoj  ###########################################
    If Round(adc("UoPrc"), 3) = 0 Then
      If adc("UoPMCdFr") <> "STKVAL" Then
        GRP_REP.SetFormula Rep, "wVal", "If " + IIF(adc("UoPMCdFr") = "CSTVAL", "{rdo.qQwCst}", "{rdo.qQw}") + " = 'Q' then {@wQty} * {rdo.RrSalRt} else " + _
                                          "(If Round({@wPureWt}, 3)<> 0 then {@wPureWt}* {rdo.RrSalRt} Else {@wWt}* {rdo.RrSalRt})"
      
      Else
        '*** (Bef 05/05/07)
        ''****** Sachin 2.13.0 - 05-06/2006 - Stock Value will be always by Weight Wise ******
        'GRP_REP.SetFormula Rep, "wVal", "If Round({@wPureWt}, 3)<> 0 then {@wPureWt}* {rdo.RrSalRt} Else {@wWt}* {rdo.RrSalRt}"
        ''****** Sachin 2.13.0 - 05-06/2006 - Stock Value will be always by Weight Wise ******
        '*** (Bef 05/05/07)
        
        '*** (Jen 2.13 05/05/07)
        GRP_REP.SetFormula Rep, "wVal", "({rdo.qOpnWt}+ {rdo.qDrWt}- {rdo.qCrWt})* {rdo.RrSalRt}"
        '*** (Jen 2.13 05/05/07)
      End If
    Else
      If adc("UoPMCdFr") <> "STKVAL" Then
        GRP_REP.SetFormula Rep, "wVal", "If " + IIF(adc("UoPMCdFr") = "CSTVAL", "{rdo.qQwCst}", "{rdo.qQw}") + " = 'Q' then {@wQty} * {rdo.RrSalRt} else " + _
                                    "(If Round({@wPureWt}, 3)<> 0 then {@wPureWt}* {rdo.RrSalRt} + {@wPureWt}* {rdo.RrSalRt}* " + _
                                    CStr(adc("UoPrc")) + "/ 100 Else {@wWt}* {rdo.RrSalRt} + {@wWt}* {rdo.RrSalRt}* " + _
                                    CStr(adc("UoPrc")) + "/ 100)"
                                    '"{@wWt}* {rdo.RrSalRt}+ {@wWt}* {rdo.RrSalRt}* " + CStr(adc("UoPrc")) + "/ 100"
      Else
        '*** (Bef 05/05/07)
        ''****** Sachin 2.13.0 - 05-06/2006 - Stock Value will be always by Weight Wise ******
        'GRP_REP.SetFormula Rep, "wVal", "If Round({@wPureWt}, 3)<> 0 then {@wPureWt}* {rdo.RrSalRt} + {@wPureWt}* {rdo.RrSalRt}* " + _
        '                                CStr(adc("UoPrc")) + "/ 100 Else {@wWt}* {rdo.RrSalRt} + {@wWt}* {rdo.RrSalRt}* " + _
        '                                CStr(adc("UoPrc")) + "/ 100"
        ''****** Sachin 2.13.0 - 05-06/2006 - Stock Value will be always by Weight Wise ******
        '*** (Bef 05/05/07)
        
        '*** (Jen 2.13 05/05/07)
        GRP_REP.SetFormula Rep, "wVal", "({rdo.qOpnWt}+ {rdo.qDrWt}- {rdo.qCrWt})* {rdo.RrSalRt} + " + _
                                        "({rdo.qOpnWt}+ {rdo.qDrWt}- {rdo.qCrWt})* {rdo.RrSalRt}* " + CStr(adc("UoPrc")) + "/ 100"
        '*** (Jen 2.13 05/05/07)
      End If
    End If
  End If
  
  Rep.TxtRmCtg.SetText adc("UoRmCtgFr")
  Rep.txtLocTypFr.SetText adc("UoLocTypFr")
  '***************Geeta************Emr206**********
  Rep.TxtLocTypTo.SetText adc("UoLocTypTo")
  '***************Geeta************Emr206**********
  Rep.TxtYyMmFr.SetText adc("UoYyMmFr")
  Rep.TxtYyMmTo.SetText adc("UoYyMmTo")
  If adc("UoYN") = "N" Then
    Rep.BoxDet.Suppress = True
    Rep.BoxHd.Suppress = True
  End If
  
  ' Zubin 212
  If Not moCn.RecSeek("select 'x' from Head where HCd= '" + ctSelfCmCd + "' and HCoCd= '" + gs_CoCd + "' and HLotNoYN= 'Y' ") Then
    Rep.HSLotNo.Suppress = True: Rep.FldSLotNo.Suppress = True
  End If
  ' Zubin 212
  
  ' ******** Manali 3.03 - 10/07/08
  
'  GRP_REP.SetFormula Rep, "wRmSz", "If {rdo.qRmSzDesc}<> '' Then {rdo.qRmSzDesc} " + _
                                   "Else (If {rdo.qPtrYN}= 'Y' then ToText({rdo.SRmSz}, 4) Else ToText({rdo.SRmSz}, 3))"
  '2 else part added to print size as Sz1*Sz2*Sz3
  GRP_REP.SetFormula Rep, "wRmSz", "If {rdo.qRmSzDesc}<> '' Then {rdo.qRmSzDesc} " + _
                                      "Else CustRmSzDisp ({rdo.SRmSz},{rdo.SRmSz2},{rdo.SRmSz3}) "
                                   
   ' ******** Manali 3.03 - 10/07/08
               
  '*** (Bef 30/09/05)
  'If adc("UoRmCtgFr") <> "G" And adc("UoRmCtgFr") <> "P" And adc("UoRmCtgFr") <> "X" Then
    'REP.h1PureWt.Suppress = True
    'REP.h2PureWt.Suppress = True
    'REP.FldG1PureWt.Suppress = True
    'REP.FldG2PureWt.Suppress = True
    'REP.FldG3PureWt.Suppress = True
    'REP.FldG4PureWt.Suppress = True
    'REP.FldGPureWt.Suppress = True
  'End If
  '*** (Bef 30/09/05)
  
  'KC.3- StkRt-when Show Stock Rate=YES then Stock Rate is visible in report
  'if it is NO then Stock Rate is invisible in report
 
  If adc("UoYN4") = "Y" Then
    Rep.hSRmStkRt.SetText "Stock Rt"
    Rep.FldRmStkRt.Suppress = False
  Else
    Rep.hSRmStkRt.SetText ""
    Rep.FldRmStkRt.Suppress = True
  End If
  '2 if YyyyMm group sort selected then display only issue & receipt
  'opening, closing, purewt, Rmqty/wt columns will be suppress
  Dim n As Integer
  For n = 1 To 2
    If wGrp(3) = "YyyyMm" Or (wGrp(n) = "YyyyMm" And wGrp(n + 1) = "(None)") Then
      Rep.Fldg3Opn.Suppress = True: Rep.Fldg3Cls.Suppress = True: Rep.FldG3PureWt.Suppress = True: Rep.fldG3Val.Suppress = True:
    ElseIf wGrp(n) = "YyyyMm" And wGrp(n + 1) = "(Detail)" Then
      Rep.Fldg4Opn.Suppress = True: Rep.Fldg4Cls.Suppress = True: Rep.FldG4PureWt.Suppress = True: Rep.fldG4Val.Suppress = True:
      Rep.Fldg2Opn.Suppress = True: Rep.Fldg2Cls.Suppress = True: Rep.FldG2PureWt.Suppress = True: Rep.fldG2Val.Suppress = True
    ElseIf wGrp(1) = "YyyyMm" And (wGrp(2) = "(Detail)" Or wGrp(2) = "(None)") Then
      Rep.FldgOpn.Suppress = True:  Rep.FldgCls.Suppress = True:  Rep.FldGPureWt.Suppress = True:  Rep.fldGVal.Suppress = True
    ElseIf wGrp(1) = "YyyyMm" And wGrp(3) <> "(None)" Then
      Rep.Fldg1Opn.Suppress = True: Rep.Fldg1Cls.Suppress = True: Rep.FldG1PureWt.Suppress = True: Rep.fldG1Val.Suppress = True:
    ElseIf (wGrp(1) = "YyyyMm" And wGrp(3) = "(None)") Or (wGrp(2) = "YyyyMm" And wGrp(3) <> "(None)" And wGrp(3) <> "(Detail)") Then
      Rep.Fldg2Opn.Suppress = True: Rep.Fldg2Cls.Suppress = True: Rep.FldG2PureWt.Suppress = True: Rep.fldG2Val.Suppress = True
    End If
  Next n
  '578.19 calling function to change Rmcode font size
  GRP_REP.SetFormula Rep, "wRmCdFontSz", GetRmcdFontSz("{rdo.RmCd}", 8)
  
  CRV_REP.DisplayGroupTree = False
  
  Call DispReport(adc) '****** Sachin 3.02 - Display Report ******
  
End Sub

' ***** Manali 3.5.0 - 18/11/08 - Check Boxes Added
Private Sub ChkBoxArr_Click(Index As Integer)
With ChkBoxArr(Index)
  Select Case .Tag
  Case Is = "ShowBoxDet"    'Index = 0
    If .Value = Checked Then adc("UoYN") = "Y" Else adc("UoYN") = "N"
  Case Is = "ShowZWt"    'Index = 1
    If .Value = Checked Then adc("UoYN1") = "Y" Else adc("UoYN1") = "N"

  'StkRt-To select this option(Show Stock rate) either Seo user can give their password or
  ' can enter SEOPWD password otherwise this option can't not select
  Case Is = "ShowStkRt"
    If .Value = Checked Then
       If Trim(TxtPwd.text) <> moCn.GetFldVal("select IsNull ((select PValue from Param Join UsrAccess on UaUsrCd= PMCd " + _
            "where PTyp= 'USR' and PMCd= '" + gs_UsrCd + "'  and UaMnuCd= '" + adc.MenuCd + "' and " + _
            "UaSeoYn= 'Y' and PValue<> '' and PValue= '" + Trim(TxtPwd.text) + "' ), " + _
            "Isnull ((Select PValue1 from Param where " + _
            "Ptyp= 'SEOPWD' and PmCd= 'SEOPWD' ), ''))") Then _
            .Value = Unchecked
    End If
    If .Value = Checked Then adc("UoYN4") = "Y" Else adc("UoYN4") = "N"
  End Select
End With
End Sub
' ***** Manali 3.5.0 - 18/11/08 - Check Boxes Added

' ***** Manali 3.5.0 - 18/11/08 - Check Boxes Added
Private Sub ChkBoxArr_GotFocus(Index As Integer)
  DispMsg ChkBoxArr(Index).ToolTipText, etInfo
End Sub
' ***** Manali 3.5.0 - 18/11/08 - Check Boxes Added

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
  
  '****** Sachin 2.14.0 - [23-11-2006] - [Multi-Company] ******
  ws_MultiCoMod = IIF(moCn.GetFldVal("Select hImpComLoc From Head Where HCd='" + ctSelfCmCd + "' And HCoCd='" + ctSelfCoCd + "'") = "Y", True, False)
  gb_CoCdFor = False
  '****** Sachin 2.14.0 - [23-11-2006] - [Multi-Company] ******
  
  '****** Sachin 2.14.0 - [23-11-2006] - [Multi-Company] ******
  If ws_MultiCoMod = True Then
    Set adc.FirNKeyCtl = adc("UoCoCdFr")
  Else
    Set adc.FirNKeyCtl = adc("UoLocTypFr")
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
    LblLotNo.Visible = False: adc("UoLotNoFr").Visible = False
    adc("UoLotNoTo").Visible = False: adc("UoLotNoSel").Visible = False
    adc("UoLotNoFr") = "": adc("UoLotNoTo") = "": adc("UoLotNoSel") = ""
  End If
  ' Zubin 212
  'Call SetGroupSort        4.1.4
  '*** Jay 2.14 *** (SEO Change)
  ms_PwdFlag = moCn.RecSeek("Select PValue1 from Param where PTyp= 'SEOPWD' and PmCd= 'SEOPWD' and PValue1= '' ")
  '*** Jay 2.14 *** (SEO Change)
  
  '****** Sachin 2.14.0 - [23-11-2006] - [Multi-Company] ******
  If ws_MultiCoMod = True Then
      adc("UOCOCDFR").Enabled = True: adc("UOCOCDTO").Visible = True: adc("UOCOCDSEL").Visible = True
      LblPrntCd.Visible = True: adc("UOPRNTCDFR").Visible = True: adc("UOPRNTCDTO").Visible = True: adc("UOPRNTCDSEL").Visible = True
  Else
      adc("UOCOCDFR").Enabled = False: adc("UOCOCDTO").Visible = False: adc("UOCOCDSEL").Visible = False
      LblPrntCd.Visible = False: adc("UOPRNTCDFR").Visible = False: adc("UOPRNTCDTO").Visible = False: adc("UOPRNTCDSEL").Visible = False
  End If
  '****** Sachin 2.14.0 - [23-11-2006] - [Multi-Company] ******

End Sub


Private Sub ATXT_FldChange(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String, ByVal pv_NewValue As Variant)
  Dim ws_NewVal As String
  
  gb_CoCdFor = False '****** Sachin 2.14.0 - [25-11-2006] - [Multi-Company] ******
  
  Select Case UCase(IdName)
  Case Is = UCase("UoPMCdFr"), UCase("UoPmCdTo")                                        '4.1.4
    If UCase(IdName) = UCase("UoPmCdFr") And pv_NewValue <> "STKVAL" Then adc("UoPMCdTo") = ""  '4.1.4
    gs_Tbl = "Param"
    gs_PTyp = "WIPVAL"
  
  Case Is = UCase("UoLocTypFr")
    ws_NewVal = pv_NewValue
    
    'If ws_NewVal = "" Then ws_NewVal = "P"
    'gs_LocTyp = "'" + ws_NewVal + "'"
    
    gs_LocTyp = IIF(ws_NewVal = "", "", "'" + ws_NewVal + "'")
' ########################  Manoj #### Ver: 2.0.6 #### Date: 24/01/2004  ########################
'  Case Is = UCase("UoYn1")
'    If ADC("UoYN1") = "N" And Trim(TxtPwd.Text) <> moCn.GetFldVal("Select PValue1 from Param where " + _
'                                   "Ptyp = 'SEOPWD' and PmCd = 'SEOPWD'") Then
'      Cancel = True: DispMsg "Enter SEO Password to View Values", etError
'    End If
' ###########################################  Manoj  ###########################################
  End Select
  Call SetProp(Me, IdName, Valid, Cancel, ErrMsg)
End Sub
Private Sub ATXT_FldWhen(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String)
  
  gb_CoCdFor = False '****** Sachin 2.14.0 - [25-11-2006] - [Multi-Company] ******
  
  Select Case UCase(IdName)
  Case Is = UCase("UoPMCdFr")
' ########################  Manoj #### Ver: 2.0.6 #### Date: 24/01/2004  ########################
    '*** Jay 2.14 *** (SEO Change)
    'If TxtPwd.text = "" Then Cancel = True: ErrMsg = "Enter SEO Password to View Values ": Exit Sub
    If ms_PwdFlag = False Then Cancel = True: ErrMsg = "Enter SEO Password to View Values ": Exit Sub
    '*** Jay 2.14 *** (SEO Change)
'    If adc("UoYn1") = "N" Then Cancel = True: ErrMsg = "Enter SEO Password to View Values ": Exit Sub
' ###########################################  Manoj  ###########################################
    gs_Tbl = "Param"
    gs_PTyp = "WIPVAL"
' ########################  Manoj #### Ver: 2.0.6 #### Date: 08/12/2003  ########################
' Change: Help for Location fields depending on Location Type
  Case Is = UCase$("UoIssLocFr")
    gs_LocTyp = IIF(adc("UoLocTypFr") = "", "", "'" + adc("UoLocTypFr") + "'")
  Case Is = UCase$("UoIssLocTo")
    gs_LocTyp = IIF(adc("UoLocTypTo") = "", "", "'" + adc("UoLocTypTo") + "'")
  Case Is = UCase$("UoIssLocSel")
'    gs_LocTyp = adc("UoLocTypSel")
    gs_LocTyp = IIF(adc("UoLocTypFr") = "", "", "'" + adc("UoLocTypFr") + "'")
' ###########################################  Manoj  ###########################################
  '*** Jay 3.01 **** [LotNo Err]
  Case Is = UCase("UoLotNoFr"), UCase("UoLotNoSel")
    gs_RmLocTyp = adc("UoLocTypFr")
    gs_Loc = adc("UoIssLocFr")
  Case Is = UCase("UoLotNoTo")
    gs_RmLocTyp = adc("UoLocTypTo")
    gs_Loc = adc("UoIssLocTo")
  ' ******* Manali 350Nxt
  Case Is = UCase("UoLotNoSel")
    'gs_RmLocTyp = adc("UoLocTypFr")
    gs_Loc = adc("UoIssLocFr")
  ' ******* Manali 350Nxt
  '*** Jay 3.01 **** [LotNo Err]
  
  Case Is = UCase("UoPMCdTo")     '4.1.4
      If adc.Mode = xNorm Then _
        If UCase(adc("UoPMCdFr")) <> UCase("STKVAL") Then Cancel = True: ErrMsg = "Can Enter Only if Value Type is 'STKVAL'": Exit Sub

  End Select
  
  SetProp Me, IdName, When
End Sub

Private Sub atxt_ReCalculate(Index As Integer, ByVal IdName As String, ByVal OldVal As Variant, ByVal RowNum As Integer, ByVal SourceName As String, ByVal SourceColName As String)
' ########################  Manoj #### Ver: 2.0.6 #### Date: 24/01/2004  ########################
'  Select Case UCase(IdName)
'  Case Is = UCase("UoPMCdFr")
'    If ADC("UoYn1") = "N" Then ADC("UoPMCdFr") = ""
'  End Select
' ###########################################  Manoj  ###########################################
End Sub

' ****** Manali 3.5.0 - 07/11/08 - Radio Buttons Added
Private Sub OptQW_Click(Index As Integer)
Select Case Index
Case Is = 0
   adc("UoYN3") = "Y"
Case Is = 1
   adc("UoYN3") = "N"
End Select
End Sub
' ****** Manali 3.5.0 - 07/11/08 - Radio Buttons Added

' ****** Manali 3.5.0 - 07/11/08 - Radio Buttons Added
Private Sub OptQW_GotFocus(Index As Integer)
  DispMsg FraOptQW.ToolTipText, etInfo
End Sub
' ****** Manali 3.5.0 - 07/11/08 - Radio Buttons Added

Private Sub tab_REP_Click(PreviousTab As Integer)
  fra_tabrep(TAB_REP.Tab).Enabled = True
  fra_tabrep(PreviousTab).Enabled = False
End Sub
Private Sub ADC_KeyWhen()
  adc("UOUSRCD") = adc.UsrCd
  adc("UOMNUCD") = adc.MenuCd
  Call EnaDisaCmds(False)
' ########################  Manoj #### Ver: 2.0.6 #### Date: 24/01/2004  ########################
'  Call HideAllFras
' ###########################################  Manoj  ###########################################
End Sub
Private Sub ADC_NonKeyWhen(pr_Cancel As Boolean, pr_ErrMsg As String)
  
  '****** Sachin 2.14.0 - [25-11-2006] - [Multi-Company] ******
  If ws_MultiCoMod = False Then _
      adc("UoCoCdFr") = gs_CoCd: adc("UoCoCdTo") = gs_CoCd: adc("UoCoCdSel") = "": adc("UoCoCdFr").Enabled = False
  '****** Sachin 2.14.0 - [25-11-2006] - [Multi-Company] ******
  
  Call SetGroupSort     '4.1.4
  
  'adc("UoCoCdFr") = gs_CoCd
  adc("UoOpnYyMm") = 0
  gs_LocTyp = "'" + adc("UoLocTypFr") + "'"
' ########################  Manoj #### Ver: 2.0.6 #### Date: 24/01/2004  ########################
  TxtPwd.text = ""
  adc("UoPMCdFr") = ""
'  adc("UOYN1") = "N"
' ###########################################  Manoj  ###########################################

If adc("UoYN") = "" Then adc("UoYN") = "N"
If adc("UoYN1") = "" Then adc("UoYN1") = "N"

' ****** Manali 3.5.0 - 07/11/08 - Radio Buttons Added
  If adc("UoYN3") = "Y" Then OptQW(0).Value = True
  If adc("UoYN3") = "N" Then OptQW(1).Value = True
' ****** Manali 3.5.0 - 07/11/08 - Radio Buttons Added

 ' ***** Manali 3.5.0 - 18/11/08 - Check Boxes Added
 If adc("UoYN") = "Y" Then ChkBoxArr(0).Value = Checked Else ChkBoxArr(0).Value = Unchecked
 If adc("UoYN1") = "Y" Then ChkBoxArr(1).Value = Checked Else ChkBoxArr(1).Value = Unchecked
 ' ***** Manali 3.5.0 - 18/11/08 - Check Boxes Added
 
   'StkRt-Set Show StockRate option is No
  adc("UoYN4") = "N"


End Sub
Private Sub DispCoNm()
  Dim ws_HName As String
  
  ws_HName = moCn.GetFldVal("Select HName from Head where HCoCd ='" + IIF(ws_MultiCoMod = True, gs_CoCd, adc("UOCOCDFR")) + "' and " + _
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
  Rep.TxtCoCdFr.SetText adc("UOCOCDFR")
  
  If ws_MultiCoMod = True Then
      Rep.TxtCoCdToLbl.Suppress = False: Rep.TxtCoCdSelLbl.Suppress = False
      Rep.TxtCoCdTo.SetText adc("UOCOCDTO"): Rep.TxtCoCdSel.SetText adc("UOCOCDSEL")
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
' ########################  Manoj #### Ver: 2.0.6 #### Date: 24/01/2004  ########################
'**** Shilpa Emr 206 30/10/03 *****/
'Private Sub CmdPwd_Click()
'  Call DispFra(Pwd)
'End Sub
'Private Sub DispFra(ByVal pv_FlshStkFra As en_FlshStkFra)
'  '*** This is the common procedure which is used to toggle between hiding and displaying
'      'the frames (i.e. if the frame is visible then make it invisible and viseversa)
'      'associated with a command button
'  Select Case pv_FlshStkFra
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

'StkRt-if password is change to blank, Show StockRate check box will turn into blank
Private Sub TxtPwd_Change()
    If Trim(TxtPwd.text) = "" Then
        ChkBoxArr(2).Value = Unchecked
    End If
End Sub

Private Sub TxtPwd_GotFocus()
  DispMsg "Enter SEO Password", etInfo
End Sub

Private Sub TxtPwd_Validate(Cancel As Boolean)
'*** Jay 2.14 *** (SEO Change)
  If Trim(TxtPwd.text) = "" Then
    adc("UoPMCdFr") = ""
    ms_PwdFlag = moCn.RecSeek("Select PValue1 from Param where PTyp= 'SEOPWD' and PmCd= 'SEOPWD' and PValue1= '' ")
  Else
    If Trim(TxtPwd.text) = moCn.GetFldVal("select IsNull ((select PValue from Param Join UsrAccess on UaUsrCd = PMCd " + _
                   "where PTyp= 'USR' and PMCd= '" + gs_UsrCd + "' and UaMnuCd = '" + adc.MenuCd + "' and " + _
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
'    If Trim(TxtPwd.text) = "" Or Trim(TxtPwd.text) = moCn.GetFldVal("Select PValue1 from Param where " + _
'                           "Ptyp = 'SEOPWD' and PmCd = 'SEOPWD'") Then
'      Exit Sub
'    Else
'      Cancel = True
'      DispMsg "Invalid Edit Password", etError
'    End If
' Bef 2.14
End Sub
'**** Shilpa Emr 206 30/10/03 *****/
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
