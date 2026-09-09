VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{AB2CC32A-E652-42A2-B341-26ECE02F6386}#1.0#0"; "mwfctl.ocx"
Begin VB.Form EmrFrpOrd 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Order Printing"
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
   Begin MwfCtl.MWCTL_BTN1 CmdChgImg 
      Height          =   390
      Left            =   10800
      TabIndex        =   193
      ToolTipText     =   "Analysis Scope Options"
      Top             =   9660
      Visible         =   0   'False
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   688
      BackColor       =   -2147483633
      MaskColor       =   8454143
      Caption         =   "Chg &Image"
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
   Begin VB.Frame FraChgImg 
      Height          =   1275
      Left            =   5880
      TabIndex        =   190
      Top             =   8280
      Visible         =   0   'False
      Width           =   6135
      Begin MwfCtl.MWCTL_BTN1 CmdChgImgGo 
         Height          =   390
         Left            =   5400
         TabIndex        =   196
         ToolTipText     =   "Analysis Scope Options"
         Top             =   675
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   688
         BackColor       =   -2147483633
         MaskColor       =   8454143
         Caption         =   "Go"
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
      Begin MwfCtl.MWCTL_BTN1 CmdSelFile 
         Height          =   285
         Left            =   4920
         TabIndex        =   195
         ToolTipText     =   "Analysis Scope Options"
         Top             =   720
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   503
         BackColor       =   -2147483633
         MaskColor       =   8454143
         Caption         =   "..."
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
         Index           =   78
         Left            =   1200
         TabIndex        =   194
         ToolTipText     =   "Enter Sort Option"
         Top             =   720
         Width           =   3780
         _ExtentX        =   6668
         _ExtentY        =   503
         MaxLength       =   0
         IdName          =   "WXLFILEPATH"
      End
      Begin VB.Label LblMain 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "  Select Report Excel File"
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
         Left            =   0
         TabIndex        =   192
         Top             =   240
         Width           =   6180
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Report File"
         BeginProperty Font 
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
         Left            =   120
         TabIndex        =   191
         Top             =   720
         Width           =   1245
      End
   End
   Begin MwfCtl.MWCTL_BTN1 CmdDispSc 
      Height          =   390
      Left            =   3210
      TabIndex        =   166
      ToolTipText     =   "Display Scope Options"
      Top             =   9660
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   688
      BackColor       =   -2147483633
      MaskColor       =   8454143
      Caption         =   "Disp&lay "
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
   Begin MwfCtl.MWCTL_BTN1 CmdSc 
      Height          =   390
      Left            =   1890
      TabIndex        =   88
      ToolTipText     =   "Analysis Scope Options"
      Top             =   9660
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   688
      BackColor       =   -2147483633
      MaskColor       =   8454143
      Caption         =   "Extra Sc&ope"
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
      Height          =   495
      Left            =   10710
      TabIndex        =   90
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
      TabIndex        =   89
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
      TabIndex        =   99
      Top             =   9645
      Width           =   3555
      Begin MwfCtl.MWCTL_MED aktxt 
         Height          =   345
         Index           =   0
         Left            =   210
         TabIndex        =   100
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
         TabIndex        =   101
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
      TabIndex        =   98
      Top             =   9600
      Visible         =   0   'False
      Width           =   855
      Begin CRVIEWER9LibCtl.CRViewer9 CRV_REP 
         Height          =   495
         Left            =   480
         TabIndex        =   95
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
      TabIndex        =   91
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
      TabIndex        =   94
      Top             =   45
      Width           =   15255
      Begin VB.Frame FraSc 
         Height          =   3915
         Left            =   6120
         TabIndex        =   126
         Top             =   5580
         Visible         =   0   'False
         Width           =   3885
         Begin VB.CheckBox ChkBoxArr 
            Alignment       =   1  'Right Justify
            Caption         =   "Show Design Remarks                    "
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
            Index           =   26
            Left            =   90
            TabIndex        =   175
            Tag             =   "ShowDsgRem"
            ToolTipText     =   "Check To Show Order Design Level Remark"
            Top             =   1560
            Width           =   3600
         End
         Begin VB.CheckBox ChkBoxArr 
            Alignment       =   1  'Right Justify
            Caption         =   "Show Summary                              "
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
            Left            =   90
            TabIndex        =   178
            Tag             =   "ShowSumm"
            ToolTipText     =   "Check To Show Summary"
            Top             =   2415
            Width           =   3600
         End
         Begin VB.CheckBox ChkBoxArr 
            Alignment       =   1  'Right Justify
            Caption         =   "Show Cust RmCd                            "
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
            Left            =   90
            TabIndex        =   177
            Tag             =   "ShowCustRm"
            ToolTipText     =   "Check To Show Customer Rm Code"
            Top             =   2130
            Width           =   3600
         End
         Begin VB.CheckBox ChkBoxArr 
            Alignment       =   1  'Right Justify
            Caption         =   "Show Cust DsgCd                            "
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
            Left            =   90
            TabIndex        =   176
            Tag             =   "ShowCustDsg"
            ToolTipText     =   "Check To Show Customer Design Code"
            Top             =   1845
            Width           =   3600
         End
         Begin VB.CheckBox ChkBoxArr 
            Alignment       =   1  'Right Justify
            Caption         =   "Show Order Remarks                    "
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
            Left            =   90
            TabIndex        =   174
            Tag             =   "ShowRem"
            ToolTipText     =   "Check To Show Order Master Level Remark"
            Top             =   1290
            Width           =   3600
         End
         Begin VB.CheckBox ChkBoxArr 
            Alignment       =   1  'Right Justify
            Caption         =   "Labour Value                                  "
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
            Left            =   90
            TabIndex        =   173
            Tag             =   "ShowLabVal"
            ToolTipText     =   "Check To Show Labour Value"
            Top             =   1005
            Width           =   3600
         End
         Begin VB.CheckBox ChkBoxArr 
            Alignment       =   1  'Right Justify
            Caption         =   "Rm Value                                        "
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
            Left            =   90
            TabIndex        =   172
            Tag             =   "ShowRmVal"
            ToolTipText     =   "Check To Show Picture"
            Top             =   720
            Width           =   3600
         End
         Begin VB.CheckBox ChkBoxArr 
            Alignment       =   1  'Right Justify
            Caption         =   "Show Design values                        "
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
            Left            =   90
            TabIndex        =   171
            Tag             =   "ShowDsgVal"
            ToolTipText     =   "Check To Show Design Value"
            Top             =   435
            Width           =   3600
         End
         Begin VB.Frame FraOptOrdQty 
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   2010
            TabIndex        =   170
            ToolTipText     =   "Select Amended + Entered Qty O r Entered Qty"
            Top             =   2700
            Width           =   1785
            Begin VB.OptionButton OptOrdQty 
               Caption         =   "Amd"
               BeginProperty Font 
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
               TabIndex        =   77
               ToolTipText     =   "Select Amended + Entered Qty O r Entered Qty"
               Top             =   0
               Width           =   735
            End
            Begin VB.OptionButton OptOrdQty 
               Caption         =   " Enter"
               BeginProperty Font 
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
               Left            =   900
               TabIndex        =   78
               ToolTipText     =   "Select Amended + Entered Qty O r Entered Qty"
               Top             =   0
               Width           =   885
            End
         End
         Begin VB.CheckBox ChkBoxArr 
            Alignment       =   1  'Right Justify
            Caption         =   "Picture                                            "
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
            Left            =   90
            TabIndex        =   168
            Tag             =   "ShowPic"
            ToolTipText     =   "Check To Show Picture"
            Top             =   150
            Width           =   3600
         End
         Begin VB.Frame FraOptPtr 
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   2010
            TabIndex        =   146
            ToolTipText     =   "Select Calculated Ptr Or Actual Pointer"
            Top             =   3600
            Width           =   1785
            Begin VB.OptionButton OptPtr 
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
               Index           =   1
               Left            =   900
               TabIndex        =   87
               ToolTipText     =   "Select Calculated Ptr Or Actual Pointer"
               Top             =   0
               Width           =   1005
            End
            Begin VB.OptionButton OptPtr 
               Caption         =   "Calc"
               BeginProperty Font 
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
               TabIndex        =   86
               ToolTipText     =   "Select Calculated Ptr Or Actual Pointer"
               Top             =   0
               Width           =   795
            End
         End
         Begin VB.Frame FraOptDelDt 
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   2010
            TabIndex        =   145
            ToolTipText     =   "Select Exp Del Date/ Prd Del Date"
            Top             =   3000
            Width           =   1785
            Begin VB.OptionButton OptDelDt 
               Caption         =   "    Prd"
               BeginProperty Font 
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
               Left            =   900
               TabIndex        =   81
               ToolTipText     =   "Select Exp Del Date/ Prd Del Date"
               Top             =   0
               Width           =   885
            End
            Begin VB.OptionButton OptDelDt 
               Caption         =   "Exp"
               BeginProperty Font 
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
               TabIndex        =   80
               ToolTipText     =   "Select Exp Del Date/ Prd Del Date"
               Top             =   0
               Width           =   735
            End
         End
         Begin VB.Frame FraOptRm 
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   2010
            TabIndex        =   144
            ToolTipText     =   "Select Production Rm Qty, Wt or Order Rm Qty, Wt"
            Top             =   3300
            Width           =   1755
            Begin VB.OptionButton OptRmPrd 
               Caption         =   "    Ord"
               BeginProperty Font 
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
               Left            =   900
               TabIndex        =   84
               ToolTipText     =   "Select Production Rm Qty, Wt or Order Rm Qty, Wt"
               Top             =   0
               Width           =   945
            End
            Begin VB.OptionButton OptRmPrd 
               Caption         =   "Prd"
               BeginProperty Font 
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
               TabIndex        =   83
               ToolTipText     =   "Select Production Rm Qty, Wt or Order Rm Qty, Wt"
               Top             =   0
               Width           =   645
            End
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   2
            Left            =   3540
            TabIndex        =   68
            ToolTipText     =   "Show Picture (Yes / No)"
            Top             =   150
            Visible         =   0   'False
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   503
            Enabled         =   0   'False
            DataType        =   4
            DataField       =   "UOYN"
            IdName          =   "UOYN"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   5
            Left            =   3540
            TabIndex        =   69
            ToolTipText     =   "Show Price (Yes / No)"
            Top             =   435
            Visible         =   0   'False
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   503
            Enabled         =   0   'False
            DataType        =   4
            DataField       =   "UOYN2"
            IdName          =   "UOYN2"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   14
            Left            =   3540
            TabIndex        =   71
            ToolTipText     =   "Show Labor Value (Yes / No)"
            Top             =   1005
            Visible         =   0   'False
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   503
            Enabled         =   0   'False
            DataType        =   4
            DataField       =   "UOYN4"
            IdName          =   "UOYN4"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   98
            Left            =   3540
            TabIndex        =   72
            ToolTipText     =   "Show Remarks in  Report? (Y/N)"
            Top             =   1290
            Visible         =   0   'False
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   503
            Enabled         =   0   'False
            DataType        =   4
            DataField       =   "UOYN6"
            IdName          =   "UOYN6"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   35
            Left            =   3540
            TabIndex        =   73
            ToolTipText     =   "Show Cust DsgCd (Yes / No)"
            Top             =   1845
            Visible         =   0   'False
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   503
            Enabled         =   0   'False
            DataType        =   4
            DataField       =   "UOYN7"
            IdName          =   "UOYN7"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   38
            Left            =   3540
            TabIndex        =   74
            ToolTipText     =   "Show Cust RmCd (Yes / No)"
            Top             =   2130
            Visible         =   0   'False
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   503
            Enabled         =   0   'False
            DataType        =   4
            DataField       =   "UOYN8"
            IdName          =   "UOYN8"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   40
            Left            =   3540
            TabIndex        =   75
            ToolTipText     =   "Show Summary (Yes / No)"
            Top             =   2415
            Visible         =   0   'False
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   503
            Enabled         =   0   'False
            DataType        =   4
            DataField       =   "UOYN9"
            IdName          =   "UOYN9"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   52
            Left            =   3540
            TabIndex        =   70
            ToolTipText     =   "Show Rm Value (Yes / No)"
            Top             =   720
            Visible         =   0   'False
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   503
            Enabled         =   0   'False
            DataType        =   4
            DataField       =   "UOYN14"
            IdName          =   "UOYN14"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   6
            Left            =   3210
            TabIndex        =   82
            ToolTipText     =   "Show Production Rm Qty, Wt ('Y' For Prd/ 'N' For Ord)"
            Top             =   3300
            Visible         =   0   'False
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   503
            Enabled         =   0   'False
            DataType        =   4
            DataField       =   "UOYN3"
            ReCalcParent    =   "UOYN10"
            IdName          =   "UOYN3"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   19
            Left            =   3240
            TabIndex        =   79
            ToolTipText     =   "Show Export/ Production Delivery Date? (Enter ""Y"" for Export and ""N"" for Production)"
            Top             =   3000
            Visible         =   0   'False
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   503
            Enabled         =   0   'False
            DataType        =   4
            DataField       =   "UOYN5"
            IdName          =   "UOYN5"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   41
            Left            =   3240
            TabIndex        =   85
            ToolTipText     =   "Show Calculated Pointer (Yes / No)"
            Top             =   3600
            Visible         =   0   'False
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   503
            Enabled         =   0   'False
            DataType        =   4
            DataField       =   "UOYN10"
            ReCalcOn        =   "UOYN3"
            IdName          =   "UOYN10"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   4
            Left            =   3540
            TabIndex        =   76
            ToolTipText     =   "Entered + Amendment Qty (Yes / No)"
            Top             =   2460
            Visible         =   0   'False
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   503
            Enabled         =   0   'False
            DataType        =   4
            DataField       =   "UOYN1"
            IdName          =   "UOYN1"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   68
            Left            =   3540
            TabIndex        =   182
            ToolTipText     =   "Show Remarks in  Report? (Y/N)"
            Top             =   1560
            Visible         =   0   'False
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   503
            Enabled         =   0   'False
            DataType        =   4
            DataField       =   "UOYN26"
            IdName          =   "UOYN26"
         End
         Begin VB.Label lblAmdQty 
            BackStyle       =   0  'Transparent
            Caption         =   "Amended Qty"
            BeginProperty Font 
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
            Left            =   120
            TabIndex        =   169
            Top             =   2700
            Width           =   1305
         End
         Begin VB.Label LblCalcPtr 
            BackStyle       =   0  'Transparent
            Caption         =   "Show Pointer"
            BeginProperty Font 
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
            Left            =   120
            TabIndex        =   143
            ToolTipText     =   "Location"
            Top             =   3600
            Width           =   1395
         End
         Begin VB.Label lblDelDt 
            BackStyle       =   0  'Transparent
            Caption         =   "Show Del Date"
            BeginProperty Font 
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
            Left            =   120
            TabIndex        =   142
            Top             =   3000
            Width           =   1440
         End
         Begin VB.Label lblPrdRmQtyWt 
            BackStyle       =   0  'Transparent
            Caption         =   "Show Rm Qty && Wt"
            BeginProperty Font 
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
            Left            =   120
            TabIndex        =   141
            Top             =   3300
            Width           =   1800
         End
         Begin VB.Label LblRmVal 
            BackStyle       =   0  'Transparent
            Caption         =   "Rm Value"
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
            Left            =   120
            TabIndex        =   140
            ToolTipText     =   "Location"
            Top             =   720
            Visible         =   0   'False
            Width           =   1845
         End
         Begin VB.Label LblShowSumm 
            BackStyle       =   0  'Transparent
            Caption         =   "Show Summary"
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
            Left            =   120
            TabIndex        =   135
            ToolTipText     =   "Location"
            Top             =   2415
            Visible         =   0   'False
            Width           =   1845
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Show Cust DsgCd"
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
            Left            =   120
            TabIndex        =   134
            ToolTipText     =   "Location"
            Top             =   1815
            Visible         =   0   'False
            Width           =   1845
         End
         Begin VB.Label lblCustRm 
            BackStyle       =   0  'Transparent
            Caption         =   "Show Cust RmCd"
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
            Left            =   120
            TabIndex        =   133
            ToolTipText     =   "Location"
            Top             =   2130
            Visible         =   0   'False
            Width           =   1845
         End
         Begin VB.Label lblShowRem 
            BackStyle       =   0  'Transparent
            Caption         =   "Show Remarks "
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
            Height          =   285
            Left            =   120
            TabIndex        =   130
            Top             =   1290
            Visible         =   0   'False
            Width           =   2310
         End
         Begin VB.Label LblPic 
            BackStyle       =   0  'Transparent
            Caption         =   "Picture"
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
            Left            =   120
            TabIndex        =   129
            ToolTipText     =   "Location"
            Top             =   150
            Visible         =   0   'False
            Width           =   1845
         End
         Begin VB.Label LblShwVal 
            BackStyle       =   0  'Transparent
            Caption         =   "Show Design Values"
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
            Left            =   120
            TabIndex        =   128
            ToolTipText     =   "Location"
            Top             =   435
            Visible         =   0   'False
            Width           =   2265
         End
         Begin VB.Label LblLabVal 
            BackStyle       =   0  'Transparent
            Caption         =   "Labour Value"
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
            Left            =   120
            TabIndex        =   127
            ToolTipText     =   "Location"
            Top             =   1005
            Visible         =   0   'False
            Width           =   1845
         End
      End
      Begin TabDlg.SSTab TAB_REP 
         Height          =   9465
         Left            =   15
         TabIndex        =   93
         Top             =   45
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   16695
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "&Data Range"
         TabPicture(0)   =   "EmrFrpOrdBeforeChange.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "fra_tabrep(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "&Group Sort"
         TabPicture(1)   =   "EmrFrpOrdBeforeChange.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "fra_tabrep(1)"
         Tab(1).ControlCount=   1
         Begin VB.Frame fra_tabrep 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   9070
            Index           =   1
            Left            =   -74940
            TabIndex        =   97
            Top             =   360
            Width           =   15105
            Begin MwfCtl.MWCTL_GRP GRP_REP 
               Height          =   360
               Left            =   240
               TabIndex        =   92
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
            TabIndex        =   96
            Top             =   360
            Width           =   15105
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Show Intl Qly"
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
               Index           =   30
               Left            =   3000
               TabIndex        =   60
               Tag             =   "ShowRunSr"
               ToolTipText     =   "Check To Show Intl Quality"
               Top             =   8115
               Visible         =   0   'False
               Width           =   2040
            End
            Begin VB.Frame FraOptNetWt 
               BorderStyle     =   0  'None
               Height          =   285
               Left            =   1905
               TabIndex        =   188
               ToolTipText     =   "Select if to show GldAs Wt or Net Wt"
               Top             =   8400
               Width           =   2895
               Begin VB.OptionButton OptNetWt 
                  Caption         =   "GldAs Wt"
                  BeginProperty Font 
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
                  TabIndex        =   62
                  Top             =   0
                  Width           =   1365
               End
               Begin VB.OptionButton OptNetWt 
                  Caption         =   "Net Wt"
                  BeginProperty Font 
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
                  Left            =   1560
                  TabIndex        =   63
                  Top             =   0
                  Width           =   1365
               End
            End
            Begin VB.Frame FraLme 
               BorderStyle     =   0  'None
               Height          =   375
               Left            =   120
               TabIndex        =   183
               ToolTipText     =   "Select Any 2 LMEs"
               Top             =   8700
               Width           =   8685
               Begin VB.CheckBox ChkBoxLme 
                  Caption         =   "LML"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   270
                  Index           =   3
                  Left            =   4800
                  TabIndex        =   67
                  Tag             =   "ShowVchSr"
                  ToolTipText     =   "Check To Display LML in Report"
                  Top             =   0
                  Width           =   810
               End
               Begin VB.CheckBox ChkBoxLme 
                  Caption         =   "LMS"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   270
                  Index           =   2
                  Left            =   3800
                  TabIndex        =   66
                  Tag             =   "ShowVchSr"
                  ToolTipText     =   "Check To Display LMS in Report"
                  Top             =   0
                  Width           =   810
               End
               Begin VB.CheckBox ChkBoxLme 
                  Caption         =   "LMP"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   270
                  Index           =   1
                  Left            =   2800
                  TabIndex        =   65
                  Tag             =   "ShowVchSr"
                  ToolTipText     =   "Check To Display LMP in Report"
                  Top             =   0
                  Width           =   810
               End
               Begin VB.CheckBox ChkBoxLme 
                  Caption         =   "LMG"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   270
                  Index           =   0
                  Left            =   1800
                  TabIndex        =   64
                  Tag             =   "ShowVchSr"
                  ToolTipText     =   "Check To Display LMG in Report"
                  Top             =   0
                  Width           =   810
               End
               Begin MwfCtl.MWCTL_MED ATXT 
                  Height          =   285
                  Index           =   69
                  Left            =   5640
                  TabIndex        =   185
                  ToolTipText     =   "Enter From Design Category"
                  Top             =   0
                  Visible         =   0   'False
                  Width           =   885
                  _ExtentX        =   1561
                  _ExtentY        =   503
                  Enabled         =   0   'False
                  MaxLength       =   5
                  DataType        =   4
                  DataField       =   "UOLME1"
                  IdName          =   "UOLME1"
               End
               Begin MwfCtl.MWCTL_MED ATXT 
                  Height          =   285
                  Index           =   70
                  Left            =   6540
                  TabIndex        =   186
                  ToolTipText     =   "Enter From Design Category"
                  Top             =   0
                  Visible         =   0   'False
                  Width           =   885
                  _ExtentX        =   1561
                  _ExtentY        =   503
                  Enabled         =   0   'False
                  MaxLength       =   5
                  DataType        =   4
                  DataField       =   "UOLME2"
                  IdName          =   "UOLME2"
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Select Any 2 LMEs"
                  BeginProperty Font 
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
                  Left            =   -30
                  TabIndex        =   184
                  Top             =   0
                  Width           =   1815
               End
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
               Left            =   75
               TabIndex        =   59
               Tag             =   "ShowRunSr"
               ToolTipText     =   "Check To Show Voucher Serials"
               Top             =   8115
               Width           =   2040
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Show Vch Sr         "
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
               Index           =   24
               Left            =   75
               TabIndex        =   57
               Tag             =   "ShowVchSr"
               ToolTipText     =   "Check To Show Voucher Serials"
               Top             =   7800
               Width           =   2040
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Show Excel Report"
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
               Index           =   23
               Left            =   75
               TabIndex        =   55
               Tag             =   "ShowExcel"
               ToolTipText     =   "Check To Show Report in Excel Format"
               Top             =   7290
               Width           =   2040
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Show Grp Comb Desc "
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
               Index           =   13
               Left            =   4590
               TabIndex        =   17
               Tag             =   "ShowGrpDesc"
               ToolTipText     =   "Check To Show Group Description"
               Top             =   1965
               Width           =   2415
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Final Sort On Dsg Cd   "
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
               Left            =   4590
               TabIndex        =   19
               Tag             =   "FinalSrt"
               ToolTipText     =   "Check To Show Picture"
               Top             =   2250
               Width           =   2415
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Show Components"
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
               Height          =   480
               Index           =   12
               Left            =   75
               TabIndex        =   53
               Tag             =   "ShowComp"
               ToolTipText     =   "Check To Show Components"
               Top             =   6810
               Visible         =   0   'False
               Width           =   2040
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   3
               Left            =   1905
               TabIndex        =   1
               ToolTipText     =   "Enter Tc Type"
               Top             =   825
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOTCTYPFR"
               ReCalcParent    =   "UOCOCDFR"
               IdName          =   "UOTCTYPFR"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   18
               Left            =   5385
               TabIndex        =   6
               ToolTipText     =   "Enter To Order Year"
               Top             =   1110
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               MaxLength       =   2
               DataType        =   4
               DataField       =   "UOOMYYTO"
               IdName          =   "UOOMYYTO"
               CmpStr          =   "OdYy <="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   20
               Left            =   2685
               TabIndex        =   3
               ToolTipText     =   "Enter From Order Year"
               Top             =   1110
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               MaxLength       =   2
               DataType        =   4
               DataField       =   "UOOMYYFR"
               IdName          =   "UOOMYYFR"
               CmpStr          =   "OdYy >="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   43
               Left            =   1905
               TabIndex        =   2
               ToolTipText     =   "Enter From Order Tc"
               Top             =   1110
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOOMTCFR"
               IdName          =   "UOOMTCFR"
               CmpStr          =   "OdTc >="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   45
               Left            =   3105
               TabIndex        =   4
               ToolTipText     =   "Enter From Order Character"
               Top             =   1110
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOOMCHRFR"
               IdName          =   "UOOMCHRFR"
               CmpStr          =   "OdChr >="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   46
               Left            =   4590
               TabIndex        =   5
               ToolTipText     =   "Enter To Order Tc"
               Top             =   1110
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOOMTCTO"
               IdName          =   "UOOMTCTO"
               CmpStr          =   "OdTc <="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   47
               Left            =   5805
               TabIndex        =   7
               ToolTipText     =   "Enter To Order Character"
               Top             =   1110
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOOMCHRTO"
               IdName          =   "UOOMCHRTO"
               CmpStr          =   "OdChr <="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   21
               Left            =   5520
               TabIndex        =   11
               ToolTipText     =   "Enter To Order Serial"
               Top             =   1395
               Width           =   525
               _ExtentX        =   926
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOODSRTO"
               IdName          =   "UOODSRTO"
               CmpStr          =   "OdSr <="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   23
               Left            =   2805
               TabIndex        =   9
               ToolTipText     =   "Enter From Order Serial"
               Top             =   1395
               Width           =   525
               _ExtentX        =   926
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOODSRFR"
               IdName          =   "UOODSRFR"
               CmpStr          =   "OdSr >="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   26
               Left            =   1905
               TabIndex        =   8
               ToolTipText     =   "Enter From Order Number"
               Top             =   1395
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOOMNOFR"
               IdName          =   "UOOMNOFR"
               CmpStr          =   "OdNo >="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   29
               Left            =   4590
               TabIndex        =   10
               ToolTipText     =   "Enter To Order Number"
               Top             =   1395
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOOMNOTO"
               IdName          =   "UOOMNOTO"
               CmpStr          =   "OdNo <="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   0
               Left            =   1905
               TabIndex        =   48
               ToolTipText     =   "Enter Level Of Raw Material Detail To Be Shown"
               Top             =   5670
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               DataField       =   "UORMGRP"
               IdName          =   "UORMGRP"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   1
               Left            =   1905
               TabIndex        =   49
               ToolTipText     =   "Enter Level Of Labour Detail To Be Shown"
               Top             =   5955
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               DataField       =   "UOLABGRP"
               IdName          =   "UOLABGRP"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
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
               CmpStr          =   "OdCoCd = "
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   31
               Left            =   4590
               TabIndex        =   24
               ToolTipText     =   "Enter To Design Code"
               Top             =   2820
               Width           =   2385
               _ExtentX        =   4207
               _ExtentY        =   503
               MaxLength       =   15
               DataField       =   "UODMCDTO"
               IdName          =   "UODMCDTO"
               CmpStr          =   "OdDmCd <="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   32
               Left            =   7290
               TabIndex        =   25
               ToolTipText     =   "Enter Design Code Selection"
               Top             =   2820
               Width           =   7650
               _ExtentX        =   13494
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UODMCDSEL"
               IdName          =   "UODMCDSEL"
               CmpStr          =   "OdDmCd In"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   33
               Left            =   1905
               TabIndex        =   23
               ToolTipText     =   "Enter From Design Code"
               Top             =   2820
               Width           =   2385
               _ExtentX        =   4207
               _ExtentY        =   503
               MaxLength       =   15
               DataField       =   "UODMCDFR"
               IdName          =   "UODMCDFR"
               CmpStr          =   "OdDmCd >="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   34
               Left            =   4590
               TabIndex        =   27
               ToolTipText     =   "Enter To Design Size"
               Top             =   3105
               Width           =   585
               _ExtentX        =   1032
               _ExtentY        =   503
               MaxLength       =   5
               DataField       =   "UODMSZTO"
               IdName          =   "UODMSZTO"
               CmpStr          =   "OdDmSz <="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   36
               Left            =   1905
               TabIndex        =   26
               ToolTipText     =   "Enter From Design Size"
               Top             =   3105
               Width           =   585
               _ExtentX        =   1032
               _ExtentY        =   503
               MaxLength       =   5
               DataField       =   "UODMSZFR"
               IdName          =   "UODMSZFR"
               CmpStr          =   "OdDmSz >="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   37
               Left            =   4590
               TabIndex        =   29
               ToolTipText     =   "Enter To Suffix"
               Top             =   3390
               Width           =   1605
               _ExtentX        =   2831
               _ExtentY        =   503
               MaxLength       =   10
               DataField       =   "UOODSFXTO"
               IdName          =   "UOODSFXTO"
               CmpStr          =   "OdSfx <="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   39
               Left            =   1905
               TabIndex        =   28
               ToolTipText     =   "Enter From Suffix"
               Top             =   3390
               Width           =   1605
               _ExtentX        =   2831
               _ExtentY        =   503
               MaxLength       =   10
               DataField       =   "UOODSFXFR"
               IdName          =   "UOODSFXFR"
               CmpStr          =   "OdSfx >="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   28
               Left            =   4590
               TabIndex        =   21
               ToolTipText     =   "Enter To Design Category"
               Top             =   2535
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataType        =   4
               DataField       =   "UODMCTGTO"
               IdName          =   "UODMCTGTO"
               CmpStr          =   "DmCtg <="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   30
               Left            =   1905
               TabIndex        =   20
               ToolTipText     =   "Enter From Design Category"
               Top             =   2535
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataType        =   4
               DataField       =   "UODMCTGFR"
               IdName          =   "UODMCTGFR"
               CmpStr          =   "DmCtg >="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   15
               Left            =   4590
               TabIndex        =   39
               ToolTipText     =   "Enter To Karat"
               Top             =   4530
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataField       =   "UOKTTO"
               IdName          =   "UOKTTO"
               CmpStr          =   "OdKt <="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   16
               Left            =   7290
               TabIndex        =   40
               ToolTipText     =   "Enter Karat Selection"
               Top             =   4530
               Width           =   7650
               _ExtentX        =   13494
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOKTSEL"
               IdName          =   "UOKTSEL"
               CmpStr          =   "OdKt In"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   7
               Left            =   1905
               TabIndex        =   38
               ToolTipText     =   "Enter From Karat"
               Top             =   4530
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataField       =   "UOKTFR"
               IdName          =   "UOKTFR"
               Data            =   "ABCDW"
               CmpStr          =   "OdKt >="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   8
               Left            =   7290
               TabIndex        =   22
               ToolTipText     =   "Enter Design Category Selection"
               Top             =   2535
               Width           =   7650
               _ExtentX        =   13494
               _ExtentY        =   503
               MaxLength       =   255
               DataType        =   4
               DataField       =   "UODMCTGSEL"
               IdName          =   "UODMCTGSEL"
               CmpStr          =   "DmCtg In"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   49
               Left            =   4590
               TabIndex        =   42
               ToolTipText     =   "Enter To Export Delivery Date"
               Top             =   4815
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   3
               DataField       =   "UODELDTTO"
               IdName          =   "UODELDTTO"
               CmpStr          =   "OdExpDelDt <="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   51
               Left            =   1905
               TabIndex        =   41
               ToolTipText     =   "Enter From Export Delivery Date"
               Top             =   4815
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   3
               DataField       =   "UODELDTFR"
               IdName          =   "UODELDTFR"
               CmpStr          =   "OdExpDelDt >="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   9
               Left            =   4590
               TabIndex        =   44
               ToolTipText     =   "Enter To Export Delivery Date"
               Top             =   5100
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   3
               DataField       =   "UODELDT1TO"
               IdName          =   "UODELDT1TO"
               CmpStr          =   "OdDelDt <="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   10
               Left            =   1905
               TabIndex        =   43
               ToolTipText     =   "Enter From Export Delivery Date"
               Top             =   5100
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   3
               DataField       =   "UODELDT1FR"
               IdName          =   "UODELDT1FR"
               CmpStr          =   "OdDelDt >="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   11
               Left            =   4590
               TabIndex        =   46
               ToolTipText     =   "Enter To Sales Executive"
               Top             =   5385
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UOSALPERTO"
               IdName          =   "UOSALPERTO"
               CmpStr          =   "CmSalPer<="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   12
               Left            =   7290
               TabIndex        =   47
               ToolTipText     =   "Enter Sales Executive Selection"
               Top             =   5385
               Width           =   7650
               _ExtentX        =   13494
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOSALPERSEL"
               IdName          =   "UOSALPERSEL"
               CmpStr          =   "CmSalPer In"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   13
               Left            =   1905
               TabIndex        =   45
               ToolTipText     =   "Enter From Sales Executive"
               Top             =   5385
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UOSALPERFR"
               IdName          =   "UOSALPERFR"
               CmpStr          =   "CmSalPer >="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   22
               Left            =   4590
               TabIndex        =   31
               ToolTipText     =   "Enter To Customer"
               Top             =   3675
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   10
               DataField       =   "UOCMCDTO"
               IdName          =   "UOCMCDTO"
               CmpStr          =   "OmCmCd <="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   24
               Left            =   1905
               TabIndex        =   30
               ToolTipText     =   "Enter From Customer"
               Top             =   3675
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UOCMCDFR"
               IdName          =   "UOCMCDFR"
               CmpStr          =   "OmCmCd >="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   25
               Left            =   1905
               TabIndex        =   14
               ToolTipText     =   "Enter First Group"
               Top             =   1965
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               DataField       =   "UOINGRP"
               IdName          =   "UOINGRP"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   27
               Left            =   1905
               TabIndex        =   15
               ToolTipText     =   "Enter Second Group"
               Top             =   2250
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               DataField       =   "UOINGRP1"
               IdName          =   "UOINGRP1"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   42
               Left            =   7335
               TabIndex        =   18
               ToolTipText     =   "Final Groping on Design Code (Y/ N)"
               Top             =   2250
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
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   44
               Left            =   1905
               TabIndex        =   50
               ToolTipText     =   "Enter Currency Code Whose Orders You Wish To View Or Leave It Blank If You Wish To View All Orders In Base Currency"
               Top             =   6240
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataType        =   4
               DataField       =   "UOCURCDFR"
               ReCalcOn        =   "UOINVAL"
               IdName          =   "UOCURCDFR"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   48
               Left            =   2205
               TabIndex        =   52
               ToolTipText     =   "Show Components Details (Yes / No)"
               Top             =   6810
               Visible         =   0   'False
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               Enabled         =   0   'False
               DataType        =   4
               DataField       =   "UOYN12"
               IdName          =   "UOYN12"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   50
               Left            =   7335
               TabIndex        =   16
               ToolTipText     =   "Show Picture (Yes / No)"
               Top             =   1965
               Visible         =   0   'False
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               Enabled         =   0   'False
               DataType        =   4
               DataField       =   "UOYN13"
               IdName          =   "UOYN13"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   60
               Left            =   2205
               TabIndex        =   54
               ToolTipText     =   "Show Components Details (Yes / No)"
               Top             =   7320
               Visible         =   0   'False
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               Enabled         =   0   'False
               DataType        =   4
               DataField       =   "UOYN23"
               IdName          =   "UOYN23"
            End
            Begin MSComDlg.CommonDialog Cd1 
               Left            =   2730
               Top             =   7590
               _ExtentX        =   847
               _ExtentY        =   847
               _Version        =   393216
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   62
               Left            =   2205
               TabIndex        =   56
               ToolTipText     =   "Show Components Details (Yes / No)"
               Top             =   7800
               Visible         =   0   'False
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               Enabled         =   0   'False
               DataType        =   4
               DataField       =   "UOYN24"
               IdName          =   "UOYN24"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   63
               Left            =   1905
               TabIndex        =   51
               ToolTipText     =   "Show 3D Picture Or Hand Sketch ? (Blank For No Picture)"
               Top             =   6525
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UOPMCDFR"
               IdName          =   "UOPMCDFR"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   64
               Left            =   2205
               TabIndex        =   58
               ToolTipText     =   "Show Components Details (Yes / No)"
               Top             =   8085
               Visible         =   0   'False
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               Enabled         =   0   'False
               DataType        =   4
               DataField       =   "UOYN25"
               IdName          =   "UOYN25"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   65
               Left            =   4590
               TabIndex        =   33
               ToolTipText     =   "Enter To Purchase Order "
               Top             =   3960
               Width           =   2625
               _ExtentX        =   4630
               _ExtentY        =   503
               MaxLength       =   20
               DataField       =   "UOPONOTO"
               IdName          =   "UOPONOTO"
               CmpStr          =   "OmPoNo <="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   66
               Left            =   7290
               TabIndex        =   34
               ToolTipText     =   "Enter Purchase Order Selection"
               Top             =   3960
               Width           =   7650
               _ExtentX        =   13494
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOPONOSEL"
               IdName          =   "UOPONOSEL"
               CmpStr          =   "OmPoNo In"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   67
               Left            =   1905
               TabIndex        =   32
               ToolTipText     =   "Enter From Purchase Order "
               Top             =   3960
               Width           =   2625
               _ExtentX        =   4630
               _ExtentY        =   503
               MaxLength       =   20
               DataField       =   "UOPONOFR"
               IdName          =   "UOPONOFR"
               CmpStr          =   "OmPoNo >="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   75
               Left            =   4590
               TabIndex        =   36
               ToolTipText     =   "Enter To Purchase Order "
               Top             =   4245
               Width           =   2625
               _ExtentX        =   4630
               _ExtentY        =   503
               MaxLength       =   20
               DataField       =   "UOODPONOTO"
               IdName          =   "UOODPONOTO"
               CmpStr          =   "OdPoNo <="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   76
               Left            =   7290
               TabIndex        =   37
               ToolTipText     =   "Enter Purchase Order Selection"
               Top             =   4245
               Width           =   7650
               _ExtentX        =   13494
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOODPONOSEL"
               IdName          =   "UOODPONOSEL"
               CmpStr          =   "OdPoNo In"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   77
               Left            =   1905
               TabIndex        =   35
               ToolTipText     =   "Enter From Purchase Order "
               Top             =   4245
               Width           =   2625
               _ExtentX        =   4630
               _ExtentY        =   503
               MaxLength       =   20
               DataField       =   "UOODPONOFR"
               IdName          =   "UOODPONOFR"
               CmpStr          =   "OdPoNo >="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   71
               Left            =   4590
               TabIndex        =   13
               ToolTipText     =   "Enter To Order Date"
               Top             =   1680
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   3
               DataField       =   "UODTTO"
               IdName          =   "UODTTO"
               CmpStr          =   "OdOmDt <="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   72
               Left            =   1905
               TabIndex        =   12
               ToolTipText     =   "Enter From Order Date"
               Top             =   1680
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   3
               DataField       =   "UODTFR"
               IdName          =   "UODTFR"
               CmpStr          =   "OdOmDt >="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   74
               Left            =   5130
               TabIndex        =   61
               ToolTipText     =   "Show Components Details (Yes / No)"
               Top             =   8400
               Visible         =   0   'False
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               Enabled         =   0   'False
               DataType        =   4
               DataField       =   "UOYN28"
               IdName          =   "UOYN28"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   79
               Left            =   5130
               TabIndex        =   197
               ToolTipText     =   "Show Intl RM Quality (Yes / No)"
               Top             =   8085
               Visible         =   0   'False
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               Enabled         =   0   'False
               DataType        =   4
               DataField       =   "UOYN30"
               IdName          =   "UOYN30"
            End
            Begin VB.Label LblNetWt 
               BackStyle       =   0  'Transparent
               Caption         =   "Show "
               BeginProperty Font 
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
               TabIndex        =   189
               Top             =   8400
               Width           =   1665
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
               TabIndex        =   187
               Top             =   1680
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
               Left            =   75
               TabIndex        =   181
               Top             =   3960
               Width           =   1425
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Sub PO No"
               BeginProperty Font 
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
               Left            =   75
               TabIndex        =   180
               Top             =   4245
               Width           =   1425
            End
            Begin VB.Label LblPicOpt 
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
               Left            =   75
               TabIndex        =   179
               ToolTipText     =   "Location"
               Top             =   6525
               Width           =   1725
            End
            Begin VB.Label LblShowGrpDesc 
               BackStyle       =   0  'Transparent
               Caption         =   "Show Group Comb Desc"
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
               Left            =   4590
               TabIndex        =   139
               ToolTipText     =   "Location"
               Top             =   1965
               Visible         =   0   'False
               Width           =   2385
            End
            Begin VB.Label LblShwComp 
               BackStyle       =   0  'Transparent
               Caption         =   "Show Components"
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
               Height          =   285
               Left            =   75
               TabIndex        =   138
               Top             =   6930
               Visible         =   0   'False
               Width           =   1785
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
               TabIndex        =   137
               Top             =   6240
               Width           =   1005
            End
            Begin VB.Label Label4 
               BackStyle       =   0  'Transparent
               Caption         =   "Final Sort on Dsg Cd"
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
               Left            =   4590
               TabIndex        =   136
               ToolTipText     =   "Location"
               Top             =   2250
               Visible         =   0   'False
               Width           =   2145
            End
            Begin VB.Label LblGrpCombo 
               BackStyle       =   0  'Transparent
               Caption         =   "Grp Combination1"
               BeginProperty Font 
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
               TabIndex        =   132
               ToolTipText     =   "Location"
               Top             =   1965
               Width           =   1875
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Grp Combination2"
               BeginProperty Font 
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
               Left            =   75
               TabIndex        =   131
               ToolTipText     =   "Location"
               Top             =   2250
               Width           =   1875
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
               TabIndex        =   125
               Top             =   3675
               Width           =   1845
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Sales Executive"
               BeginProperty Font 
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
               TabIndex        =   124
               Top             =   5385
               Width           =   1845
            End
            Begin VB.Label lblPrdDelDt 
               BackStyle       =   0  'Transparent
               Caption         =   "Prd Del Dt"
               BeginProperty Font 
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
               TabIndex        =   123
               Top             =   5100
               Width           =   1845
            End
            Begin VB.Label lblExpDelDt 
               BackStyle       =   0  'Transparent
               Caption         =   "Exp Del Dt"
               BeginProperty Font 
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
               TabIndex        =   122
               Top             =   4815
               Width           =   1845
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Order Kt"
               BeginProperty Font 
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
               Left            =   75
               TabIndex        =   121
               Top             =   4530
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
               TabIndex        =   120
               Top             =   2535
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
               TabIndex        =   119
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
               TabIndex        =   118
               Top             =   2820
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
               TabIndex        =   117
               Top             =   3105
               Width           =   1845
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Suffix"
               BeginProperty Font 
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
               Left            =   75
               TabIndex        =   116
               Top             =   3390
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
               TabIndex        =   115
               Top             =   540
               Width           =   1845
            End
            Begin VB.Label LblLabGrp 
               BackStyle       =   0  'Transparent
               Caption         =   "Labour Det Level"
               BeginProperty Font 
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
               TabIndex        =   114
               ToolTipText     =   "Location"
               Top             =   5955
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
               TabIndex        =   113
               ToolTipText     =   "Location"
               Top             =   5670
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
               TabIndex        =   112
               Top             =   1110
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
               TabIndex        =   111
               Top             =   1395
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
               TabIndex        =   110
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
               Index           =   25
               Left            =   3045
               TabIndex        =   109
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
               Index           =   26
               Left            =   5325
               TabIndex        =   108
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
               Index           =   27
               Left            =   2625
               TabIndex        =   107
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
               Index           =   30
               Left            =   5445
               TabIndex        =   106
               Top             =   1395
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
               TabIndex        =   105
               Top             =   1395
               Width           =   105
            End
            Begin VB.Label lblTcTyp 
               BackStyle       =   0  'Transparent
               Caption         =   "Type"
               BeginProperty Font 
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
               TabIndex        =   104
               ToolTipText     =   "Location"
               Top             =   825
               Width           =   1845
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
               TabIndex        =   103
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
               TabIndex        =   102
               Top             =   120
               Width           =   1335
            End
         End
      End
      Begin VB.Frame FraDispSc 
         Height          =   3105
         Left            =   4200
         TabIndex        =   165
         Top             =   6390
         Visible         =   0   'False
         Width           =   2895
         Begin VB.CheckBox ChkBoxArr 
            Alignment       =   1  'Right Justify
            Caption         =   "Ord Qty                        "
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
            Index           =   27
            Left            =   120
            TabIndex        =   162
            Tag             =   "ShowOrdQty"
            ToolTipText     =   "Check To Show Ord Qty"
            Top             =   2445
            Width           =   2445
         End
         Begin VB.CheckBox ChkBoxArr 
            Alignment       =   1  'Right Justify
            Caption         =   "Design Col              "
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
            Index           =   15
            Left            =   120
            TabIndex        =   148
            Tag             =   "ShowDsgCol"
            ToolTipText     =   "Check To Show Color"
            Top             =   465
            Width           =   2445
         End
         Begin VB.CheckBox ChkBoxArr 
            Alignment       =   1  'Right Justify
            Caption         =   "Design Grs Wt         "
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
            Index           =   16
            Left            =   120
            TabIndex        =   150
            Tag             =   "ShowGrsWt"
            ToolTipText     =   "Check To Show Design Gross Wt"
            Top             =   750
            Width           =   2445
         End
         Begin VB.CheckBox ChkBoxArr 
            Alignment       =   1  'Right Justify
            Caption         =   "Rm Code Details     "
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
            Index           =   17
            Left            =   120
            TabIndex        =   152
            Tag             =   "ShowRmCd"
            ToolTipText     =   "Check To Show Code Details"
            Top             =   1035
            Width           =   2445
         End
         Begin VB.CheckBox ChkBoxArr 
            Alignment       =   1  'Right Justify
            Caption         =   "Rm SubCtg             "
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
            Index           =   18
            Left            =   120
            TabIndex        =   154
            Tag             =   "ShowRmSCtg"
            ToolTipText     =   "Check To Show Rm SubCtg Details"
            Top             =   1320
            Width           =   2445
         End
         Begin VB.CheckBox ChkBoxArr 
            Alignment       =   1  'Right Justify
            Caption         =   "Rm Wt Details         "
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
            Index           =   19
            Left            =   120
            TabIndex        =   156
            Tag             =   "ShowRmWt"
            ToolTipText     =   "Check To Show Rm Weight Details"
            Top             =   1605
            Width           =   2445
         End
         Begin VB.CheckBox ChkBoxArr 
            Alignment       =   1  'Right Justify
            Caption         =   "Rm Quantity           "
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
            Index           =   20
            Left            =   120
            TabIndex        =   158
            Tag             =   "ShowRmQty"
            ToolTipText     =   "Check To Show Rm Quantity"
            Top             =   1890
            Width           =   2445
         End
         Begin VB.CheckBox ChkBoxArr 
            Alignment       =   1  'Right Justify
            Caption         =   "Unit Price               "
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
            Index           =   21
            Left            =   120
            TabIndex        =   160
            Tag             =   "ShowUnitPrc"
            ToolTipText     =   "Check To Show Unit Price"
            Top             =   2175
            Width           =   2445
         End
         Begin VB.CheckBox ChkBoxArr 
            Alignment       =   1  'Right Justify
            Caption         =   "Sub PO                   "
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
            Index           =   22
            Left            =   120
            TabIndex        =   164
            Tag             =   "ShowOdPO"
            ToolTipText     =   "Check To Show Sub PO"
            Top             =   2730
            Width           =   2445
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   55
            Left            =   2490
            TabIndex        =   155
            ToolTipText     =   "Show Rm Weight Details (Yes/No)?"
            Top             =   1590
            Visible         =   0   'False
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   503
            Enabled         =   0   'False
            DataType        =   4
            DataField       =   "UOYN19"
            ReCalcParent    =   "UOYN10"
            IdName          =   "UOYN19"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   56
            Left            =   2490
            TabIndex        =   153
            ToolTipText     =   "Show Rm Sub Category (Yes/No) ?"
            Top             =   1305
            Visible         =   0   'False
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   503
            Enabled         =   0   'False
            DataType        =   4
            DataField       =   "UOYN18"
            IdName          =   "UOYN18"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   57
            Left            =   2490
            TabIndex        =   157
            ToolTipText     =   "Show Rm Quantity (Yes/No) ?"
            Top             =   1875
            Visible         =   0   'False
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   503
            Enabled         =   0   'False
            DataType        =   4
            DataField       =   "UOYN20"
            IdName          =   "UOYN20"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   58
            Left            =   2490
            TabIndex        =   159
            ToolTipText     =   "Show Unit Price (Yes/No)?"
            Top             =   2160
            Visible         =   0   'False
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   503
            Enabled         =   0   'False
            DataType        =   4
            DataField       =   "UOYN21"
            IdName          =   "UOYN21"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   59
            Left            =   2490
            TabIndex        =   163
            ToolTipText     =   "Show Sub PO (Yes/No) ?"
            Top             =   2730
            Visible         =   0   'False
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   503
            Enabled         =   0   'False
            DataType        =   4
            DataField       =   "UOYN22"
            IdName          =   "UOYN22"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   61
            Left            =   2490
            TabIndex        =   151
            ToolTipText     =   "Show Rm Code Details (Yes / No)"
            Top             =   1020
            Visible         =   0   'False
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   503
            Enabled         =   0   'False
            DataType        =   4
            DataField       =   "UOYN17"
            IdName          =   "UOYN17"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   53
            Left            =   2490
            TabIndex        =   147
            ToolTipText     =   "Show Design Color? (Yes/No)"
            Top             =   465
            Visible         =   0   'False
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   503
            Enabled         =   0   'False
            DataType        =   4
            DataField       =   "UOYN15"
            IdName          =   "UOYN15"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   54
            Left            =   2490
            TabIndex        =   149
            ToolTipText     =   "Show Design Grs Wt? (Yes / No)"
            Top             =   750
            Visible         =   0   'False
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   503
            Enabled         =   0   'False
            DataType        =   4
            DataField       =   "UOYN16"
            IdName          =   "UOYN16"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   73
            Left            =   2490
            TabIndex        =   161
            ToolTipText     =   "Show Ord Qty (Yes/No) ?"
            Top             =   2445
            Visible         =   0   'False
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   503
            Enabled         =   0   'False
            DataType        =   4
            DataField       =   "UOYN27"
            IdName          =   "UOYN27"
         End
         Begin VB.Label LblMain 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "  Show Design Details"
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
            Left            =   0
            TabIndex        =   167
            Top             =   150
            Width           =   2910
         End
      End
   End
End
Attribute VB_Name = "EmrFrpOrd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** Frp for the Detail Order Report, Average Order Report and Catalogue Report
Option Explicit
Dim mRep As CRAXDRT.Report
Dim Rep As New EmrRepOrd
Dim Rep1 As New EmrRepOrdCat
Dim moCn As MwfLib.MDOConnection    ' *** (Bef speed) bef 24/11/06
Dim oExcel, oBook, oSheet           ' ***** Manali 3.6.0 - Excel Format Report
Dim wErrMsg As String               ' ***** Manali 3.6.0 - Excel Format Report
Dim ws_BaseCurCd As String, wb_InBaseCur As Boolean
Dim wChkSel As Integer
Dim wo_rsPic As MDORowSet
Private Sub SetGroupSort()
  'GRP_REP.Add "Sub Ctg", "RmSCtg", "", "RmSCtg", "hRmSCtg", "", "", "", "(Select PDesc From Param Where PTyp= 'RMSCTG' and PMCd = RmCtg and PSCd= RmSCtg) "
End Sub
Private Sub adc_setreprecsource()
On Error GoTo RepErr
  
  Dim wAddlCnd As String '****** Sachin 3.02.0 - New fields in where clause
  Dim wExcelOrdBy As String
  
  If FraSc.Visible = True Then CmdSc_Click
  '*** Jay 3.3(FG)
  ' ***** Manali Trading Module
  If UCase(ADC.MenuCd) = UCase("RepAvgOrd") Or UCase(ADC.MenuCd) = UCase("RepFgAvg") Or UCase(ADC.MenuCd) = UCase("RepJtAvg") Then
  'If UCase(ADC.MenuCd) = UCase("RepAvgOrd") Then
  '*** Jay 3.3(FG)
    ADC("UoRmGrp") = "SR"
    ADC("UoLabGrp") = "SR"
  End If

    '*** Jay 2.13(CT) ***
    '*** Jay 3.2.0 [OdPicNm]   '*** Jay 3.2.0 [OdPoNo]
    '*** Jay 3.3(FG) [OdDmId, ]
    ' ***** Manali 3.03 - 28/06/08 - OrLmeSal added
    ' ***** Manali 3.6.0 - OmPoDt, OmLmsSal,OmLmlSal, OdDelDt, 0 as OdTrayNo, ' ' as CmLkUpMetLs, DmLsCtg, OdMulBy, ' ' as CmName, OrAlySalRt, 0 as RmPuritRt added
    ''    Select OmCoCd, OmYy, OmTc, OmChr, OmNo, OmCmCd, OmDt, OmPoNo, OmPoDt, OmLmgCst, OmLmgSal,OmLmpSal, OmLmsSal,OmLmlSal, OmDelDt, OdDelDt
    ''    OmHCtCd, OdSr,
    ''    ' ' as IdExpOdTc, ' ' as IdExpOdYy, ' ' as IdExpOdChr, 0 as IdExpOdNo, 0 as IdExpOdSr,
    ''    OdDmCd, Space(17) as CdOurDmCd, OdDmSz, OdSfx, OdDmId, DmCtg, DmSalCtg, DmTcTyp, OdDmCol, 0 as OdTrayNo, OdPicNm, OdPoNo, OdKt as IdKt, DmUom, OdVaCtg as IdVaCtg, 0.00 as IdGldLs, Space(254) as qRmCtgDesc,
    ''    space(254) as qDmCtgDesc, space(254) as qDmSalCtgDesc, space(254) as qDmColDesc,
    ''    space(254) as qVaCtgDesc,
    ''    OrRmCtg /*as qMainMetRmCtg*/, OdDmCol as OdDmCol3, OdGldAsWt, OdLabAsWt, OdSalPrc,
    ''    OdOrdEnt as qOrdQty, OdSalPrc as qSalVal,
    ''    OdDmPrdInst, OdCmPrdInst, OdCmStmpInst, OdSalRem, OdSzInst, OmSalRem, OmCmDelTerms, OmCmPayTerms, ' '  as qOdomPoNo,
    ''    OdDelDt, ' ' as CmLkUpMetLs, DmLsCtg, OdMulBy, ' ' as CmName,
    ''    space(4) as qTag, 0 as qMainMet,Space(2) as qPtrYN,
    ''    OrSrNo, OrRmCtg RmCtg, OrRmSCtg,space(14) as OrRmCd, RmQw, 0.0000 as OrLn1, ' ' as qRmSzDesc, 0.000 as OrLn2,
    ''    0.0000 as OrRmPtr, OrQty, OrWt, OrLmeSal, OrSalRt, OrSetSCd,
    ''    OrSetSalRt, OrSetSalVal, OrSalVal, Space(10) as RmRSzRt,
    ''     OrAlySalRt, 0 as RmPuritRt,
    ''    OlSrNo, OlMcd, OlSCd, PDesc as qLabDesc,
    ''    OlQw , OlQty, OlSalRt, OlSalVal,
    ''    OctSrNo , OctCd, OctSz, OctCtChr, OctParts, OctRmCd, OctDmCol, OctRmWt, OctDc, OctPrdSeq, OctMainPrtYn, OctDesc
    ''    From OrdMst, OrdDsg, OrdRm, OrdLab, OrdCT, DsgMst, Param, RmMst
    ''    Where 1 = 2
    
    '*** Jay 2.13(CT) ***
  
  
    ' Zubin 212 (Sales Ctg and qDmSalCtgDesc added)
    ' ***** Manali 3.03 - 28/06/08 - OrLmeSal added
    '*** Report Sql for Det & Avg Ord ***
    '    Select OmCoCd, OmYy, OmTc, OmChr, OmNo, OmCmCd, OmDt, OmPoNo, OmPoDt, OmLmgCst, OmLmgSal,OmLmpSal, OmLmsSal,OmLmlSal, OmDelDt, OdDelDt
    '    OmHCtCd, OdSr,
    '    ' ' as IdExpOdTc, ' ' as IdExpOdYy, ' ' as IdExpOdChr, 0 as IdExpOdNo, 0 as IdExpOdSr,
    '    OdDmCd, Space(17) as CdOurDmCd, OdDmSz, OdSfx, OdDmId, DmCtg, DmSalCtg, DmTcTyp, OdDmCol, 0 as OdTrayNo, OdPicNm, OdPoNo, OdKt as IdKt, DmUom, OdVaCtg as IdVaCtg, 0.00 as IdGldLs, Space(254) as qRmCtgDesc,
    '    space(254) as qDmCtgDesc, space(254) as qDmSalCtgDesc, space(254) as qDmColDesc,
    '    space(254) as qVaCtgDesc,
    '    OrRmCtg /*as qMainMetRmCtg*/, OdDmCol as OdDmCol3, OdGldAsWt, OdLabAsWt, OdSalPrc,
    '    OdOrdEnt as qOrdQty, OdSalPrc as qSalVal,
    '    OdDmPrdInst, OdCmPrdInst, OdCmStmpInst, OdSalRem, OdSzInst, OmSalRem, OmCmDelTerms, OmCmPayTerms, ' '  as qOdomPoNo,
    '    OdDelDt, ' ' as CmLkUpMetLs, DmLsCtg, OdMulBy, ' ' as CmName,
    '    space(4) as qTag, 0 as qMainMet,Space(2) as qPtrYN,
    '    OrSrNo, OrRmCtg RmCtg, OrRmSCtg,space(14) as OrRmCd, RmQw, 0.0000 as OrLn1, ' ' as qRmSzDesc, 0.000 as OrLn2,
    '    0.0000 as OrRmPtr, OrQty, OrWt, OrLmeSal, OrSalRt, OrSetSCd,
    '    OrSetSalRt, OrSetSalVal, OrSalVal, Space(10) as RmRSzRt,
    '     OrAlySalRt, 0 as RmPuritRt,
    '    OlSrNo, OlMcd, OlSCd, PDesc as qLabDesc,
    '    OlQw , OlQty, OlSalRt, OlSalVal
    '    From OrdMst, OrdDsg, OrdRm, OrdLab, DsgMst, Param, RmMst
    '    Where 1 = 2
    '*******************Geeta****************************************************
  
  '*** Sub Report Sql for Det & Avg Ord ***
  '    Select OdYy, OdTc, OdChr, OdNo, OdSr, DmCtg, PDesc as qDesc, OdOrdEnt as qOrdQty,
  '    OdSalPrc as qSalVal From OrdDsg, DsgMst, Param Where 1= 2
  '*** Sub Report Sql for Det & Avg Ord ***
  
    ' **** Zubin 211 **** '
    ' **** Zubin (Sub-Report RmDet - 210205) **** '
    '''''  select OdTc, OdYy, OdChr, OdNo, OrRmCtg, OrPrdQty as qRmQty, OrPrdWt as qRmWt,
    '''''  OrSalVal As qRmVal
    '''''  from OrdDsg, OrdRm where 1= 2

'  select OdTc, OdYy, OdChr, OdNo, OrRmCtg, OrPrdQty as qRmQty, OrWsQty as qWSetQty,
'  OrHsQty as qHSetQty, OrPrdWt as qRmWt, OrSalVal As qRmVal
'  from OrdDsg, OrdRm where 1= 2

  ' **** Zubin *********************** '
  ' **** Zubin 211 **** '

  '*******************Geeta****************************************************
  '***************Emr206**********************************************************
  '*** Report Sql for Catalogue ***
  
  ' Zubin 212 (Sales Ctg and qDmSalCtgDesc added)
  '*** Jay 3.3(FG)
  '    Select OmCoCd, OmYy, OmTc, OmChr, OmNo, OmCmCd, OmDt, OmPoNo, OdSr, Space(17) as CdOurDmCd, OdDmCd, OdDmSz, OdSfx,OdDmId,
  '    DmCtg, DmSalCtg, DmTcTyp,OdDmCol, OdKt as IdKt, OdVaCtg as IdVaCtg, DmUom, 0.00 as IdGldLs,
  '    Space(254) as qRmCtgDesc,
  '    space(254) as qDmCtgDesc, space(254) as qDmSalCtgDesc, space(254) as qDmColDesc,space(254) as qVaCtgDesc,
  '    OrRmctg /*as qMainMetRmCtg*/, OdGldAsWt, OdLabAsWt, OdSalPrc, OdOrdEnt as qOrdQty, PDesc as qColDesc,
  '    OdGldAsWt as qGrsWt From OrdMst,OrdRm, OrdDsg, DsgMst, Param Where 1 = 2
  '********************************************************************************

  '  Select OdYy, OdTc, OdChr, OdNo, OdSr, OrMainMet,
  '  OrRmCtg, 0 as qTag, OrRmSCtg, space(14) as OrRmCd, OrQty, OrWt
  '  From OrdDsg, OrdRm Where 1 = 2
  '*** Report Sql for Catalogue ***
  
  '''If UCase(adc.MenuCd) = UCase("RepAvgOrd") Then
  
  Dim wrepcnd As String, wCnd As String, wSqlStrg As String, ws_RmQty As String, ws_RmWt As String
  Dim wCommStr As String, wRmStr As String, wRmGrp As String, wRmCommStr As String
  Dim wLabStr As String, wLabGrp As String, wLabCommStr As String, wDsgCtgStr As String
  Dim wRmDet As String, wOdSalPrcOpt As String, wRmValOpt As String, wRmSetValOpt As String
  Dim wLabValOpt As String, wTcTyp As String, wDmTcTyp As String, wRmDetSql As String
  Dim wLabRtOpt As String, wDelDt As String, wDelStr As String, ws_RmPtr As String
  Dim wRtStr As String 'urmi 2.04-3 to get the rate when grouping is by RMCd (TotVal/Wt)
  Dim wInGrp2 As String       ' **** Manali 3.6.0
  
  '*************************** ZUBIN **************************
  ' 08th Dec 2003, EMR206
  Dim wOurDsgCd As String, wOurRmCd As String
  '*************************** ZUBIN **************************
  '*** Jay 2.13(CT) ***
  Dim wOrdCTCommStr As String
  '*** Jay 2.13(CT) ***
  
  ' **** manali 3.6.0
  Dim ws_OrdBy As String
  
  Dim ms_TcDesc As String
  
  Set Rep = Nothing
  Set Rep1 = Nothing

  If UCase(ADC("UoTcTypFr")) = "PL" Then
      ADC("UoCoCdFr") = ctSelfCoCd
  Else
    ADC("UoCoCdFr") = gs_CoCd
  End If

  If UCase(ADC("UoTcTypFr")) = "QS" Then
    wDmTcTyp = "SM"
  Else
    wDmTcTyp = "DM"
  End If
  
  '*** Jay 3.3(FG)
  ' ***** Manali Trading Module
  If UCase(ADC.MenuCd) = UCase("RepOrdCat") Or UCase(ADC.MenuCd) = UCase("RepFgCat") Or UCase(ADC.MenuCd) = UCase("RepJtCat") Then
  'If UCase(ADC.MenuCd) = UCase("RepOrdCat") Then
  '*** Jay 3.3(FG)
    GRP_REP.SetFormula Rep1, "wGrpOpt", ""
  End If

  Select Case UCase(ADC.MenuCd)
  '*** Jay 3.3(FG)
  ' ***** Manali Trading Module - "JT" added below
  Case UCase("RepAvgOrd"), UCase("RepDetOrd"), UCase("RepFgAvg"), UCase("RepFgDet"), UCase("RepJtAvg"), UCase("RepJtDet")
  'Case UCase("RepAvgOrd"), UCase("RepDetOrd")
  '*** Jay 3.3(FG)
    Set ADC.RepSource = Rep
    Set mRep = Rep
  '*** Jay 3.3(FG)
  ' ***** Manali Trading Module - "JT" added below
  Case UCase("RepOrdCat"), UCase("RepFgCat"), UCase("RepJtCat")
  'Case UCase("RepOrdCat")
  '*** Jay 3.3(FG)
    Set ADC.RepSource = Rep1
    Set mRep = Rep1
  End Select
  'Set adc.RepSource = Rep
  
  ' Zubin 213
  ' Dim ws_BaseCurCd As String, wb_InBaseCur As Boolean
  ws_BaseCurCd = moCn.GetFldVal("Select HBaseCurCd from Head where HCd= '" + ctSelfCmCd + "' and " + _
                  "HCoCd= '" + ADC("UoCoCdFr") + "'")
  If ADC("UoCurCdFr") = "" Then
    wb_InBaseCur = True
    mRep.TxtValIn.SetText "Value In " + ws_BaseCurCd
    'If UCase(adc.MenuCd) = UCase("RepOrdCat") Then mRep.TxtUnitPrcCur.SetText "Unit Price" + vbCrLf + "(" + ws_BaseCurCd + ")"
    If UCase(ADC.MenuCd) = UCase("RepOrdCat") Then
        mRep.TxtUnitPrcCur.SetText "Price" + " (" + ws_BaseCurCd + ")"
        mRep.TxtUnitPrcCur2.SetText "Price" + " (" + ws_BaseCurCd + ")"
        mRep.TxtUnitPrcCur3.SetText "Price" + " (" + ws_BaseCurCd + ")"
    End If
  Else
    wb_InBaseCur = False
    mRep.TxtValIn.SetText "Value In " + ADC("UoCurCdFr")
    'If UCase(adc.MenuCd) = UCase("RepOrdCat") Then mRep.TxtUnitPrcCur.SetText "Unit Price" + vbCrLf + "(" + adc("UoCurCdFr") + ")"
    If UCase(ADC.MenuCd) = UCase("RepOrdCat") Then
        mRep.TxtUnitPrcCur.SetText "Price" + " (" + ADC("UoCurCdFr") + ")"
        mRep.TxtUnitPrcCur2.SetText "Price" + " (" + ADC("UoCurCdFr") + ")"
        mRep.TxtUnitPrcCur3.SetText "Price" + " (" + ADC("UoCurCdFr") + ")"
    End If
  End If
  ' Zubin 213
    Call DispCoNm
    Call SetMwName(mRep)   ' ***** Manali 3.6.1 - Microway footer
    ''Call GRP_REP.Gen3LRep
    ''grpflds = GRP_REP.GrpFldLst
      
    wrepcnd = ADC.RepCond
    wCnd = IIF(wrepcnd <> "", " where " + wrepcnd, "")
    ' Zubin 213
    wCnd = wCnd + IIF(wb_InBaseCur = True, "", " and OmCmCurCd= '" + ADC("UoCurCdFr") + "' ")
   
    ' **** Crm 2.12 - 23-02-06 **** '
    Dim wCrmCnd As String
    wCrmCnd = " And OmSubmitYN <> 'N' "
    ' **** Crm 2.12 - 23-02-06 **** '
  
    ''wCommStr = "Select " + grpflds + ", "
  
  ' *** Jay 2.14Next *** [Show Rm Value]
  ''*****************************************urmi*************************************
  '  If (adc("UoYN2")) = "Y" Or UCase(adc.MenuCd) = UCase("RepOrdCat") Then
  '    wOdSalPrcOpt = IIF(wb_InBaseCur, " (Case when OmMulDiv= 'M' then OdSalPrc/OmCnvFct else OdSalPrc*OmCnvFct end) ", "OdSalPrc")
  '    wRmValOpt = "sum(" + IIF(wb_InBaseCur, " (Case when OmMulDiv= 'M' then OrSalVal/OmCnvFct else OrSalVal*OmCnvFct end) ", "OrSalVal") + ")"
  '    ' Zubin 212 (Before 31/12/05)
  ''    wRmSetValOpt = "sum(OrSetSalVal) as OrSetSalVal"
  ''    wLabValOpt = "sum(OlSalVal) as OlSalVal"
  ''    wLabRtOpt = "max(OlSalRt) as OlSalRt"
  ''    If moCn.RecSeek("Select PmCd from Param where PTyp = 'SYS' and PmCd = 'ECH'") <> False Then
  ''      GRP_REP.SetFormula REP, "wLabRtYN", "'N'"
  ''    End If
  '    ' Zubin 212 (Before 31/12/05)
  '    ' Zubin 212
  '    If (adc("UoYN4")) = "Y" Or UCase(adc.MenuCd) = UCase("RepOrdCat") Then
  '      wRmSetValOpt = "sum(" + IIF(wb_InBaseCur, " (Case when OmMulDiv= 'M' then OrSetSalVal/OmCnvFct else OrSetSalVal*OmCnvFct end) ", "OrSetSalVal") + ") as OrSetSalVal"
  '      wLabValOpt = "sum(" + IIF(wb_InBaseCur, " (Case when OmMulDiv= 'M' then OlSalVal/OmCnvFct else OlSalVal*OmCnvFct end) ", "OlSalVal") + ") as OlSalVal"
  '      wLabRtOpt = "max(" + IIF(wb_InBaseCur, " (Case when OmMulDiv= 'M' then OlSalRt/OmCnvFct else OlSalRt*OmCnvFct end) ", "OlSalRt") + ") as OlSalRt"
  '      If moCn.RecSeek("Select PmCd from Param where PTyp = 'SYS' and PmCd = 'ECH'") <> False Then GRP_REP.SetFormula Rep, "wLabRtYN", "'N'"
  '    Else
  '      wRmSetValOpt = "0 as OrSetSalVal": wLabValOpt = "0 as OlSalVal": wLabRtOpt = "0 as OlSalRt"
  '      If moCn.RecSeek("Select PmCd from Param where PTyp = 'SYS' and PmCd = 'ECH'") <> False Then GRP_REP.SetFormula Rep, "wLabRtYN", "'Y'"
  '    End If
  '    ' Zubin 212
  '  ElseIf (adc("UoYN2")) = "N" And UCase(adc.MenuCd) <> UCase("RepOrdCat") Then
  '    wOdSalPrcOpt = "0"
  '    wRmValOpt = "0"
  '    If (adc("UoYN4")) = "Y" Then
  '      wRmSetValOpt = "sum(" + IIF(wb_InBaseCur, " (Case when OmMulDiv= 'M' then OrSetSalVal/OmCnvFct else OrSetSalVal*OmCnvFct end) ", "OrSetSalVal") + ") as OrSetSalVal"
  '      wLabValOpt = "sum(" + IIF(wb_InBaseCur, " (Case when OmMulDiv= 'M' then OlSalVal/OmCnvFct else OlSalVal*OmCnvFct end) ", "OlSalVal") + ") as OlSalVal"
  '      wLabRtOpt = "max(" + IIF(wb_InBaseCur, " (Case when OmMulDiv= 'M' then OlSalRt/OmCnvFct else OlSalRt*OmCnvFct end) ", "OlSalRt") + ") as OlSalRt"
  '      If moCn.RecSeek("Select PmCd from Param where PTyp = 'SYS' and PmCd = 'ECH'") <> False Then
  '        GRP_REP.SetFormula Rep, "wLabRtYN", "'N'"
  '      End If
  '    Else
  '        wRmSetValOpt = "0 as OrSetSalVal"
  '        wLabValOpt = "0 as OlSalVal"
  '        wLabRtOpt = "0 as OlSalRt"
  '        If moCn.RecSeek("Select PmCd from Param where PTyp = 'SYS' and PmCd = 'ECH'") <> False Then
  '          GRP_REP.SetFormula Rep, "wLabRtYN", "'Y'"
  '        End If
  '    End If
  '  End If
  '
  '*****************************************urmi*************************************
    
  ''''  If (adc("UoYN2")) = "Y" Or UCase(adc.MenuCd) = UCase("RepOrdCat") Then
  ''''    wOdSalPrcOpt = IIF(wb_InBaseCur, " (Case when OmMulDiv= 'M' then OdSalPrc/OmCnvFct else OdSalPrc*OmCnvFct end) ", "OdSalPrc")
  ''''    wRmValOpt = "sum(" + IIF(wb_InBaseCur, " (Case when OmMulDiv= 'M' then OrSalVal/OmCnvFct else OrSalVal*OmCnvFct end) ", "OrSalVal") + ")"
  ''''  ElseIf (adc("UoYN2")) = "N" And UCase(adc.MenuCd) <> UCase("RepOrdCat") Then
  ''''    wOdSalPrcOpt = "0"
  ''''    wRmValOpt = "0"
  ''''  End If
  ''''
  
    '*** Jay 3.3(FG)
    ' ***** Manali Trading Module
    If (ADC("UoYN2")) = "Y" Or UCase(ADC.MenuCd) = UCase("RepOrdCat") Or UCase(ADC.MenuCd) = UCase("RepFgCat") Or UCase(ADC.MenuCd) = UCase("RepJtCat") Then
      wOdSalPrcOpt = IIF(wb_InBaseCur, " (Case when OmMulDiv= 'M' then OdSalPrc/OmCnvFct else OdSalPrc*OmCnvFct end) ", "OdSalPrc")
    ElseIf (ADC("UoYN2")) = "N" And (UCase(ADC.MenuCd) <> UCase("RepOrdCat") And UCase(ADC.MenuCd) <> UCase("RepFgCat") And UCase(ADC.MenuCd) <> UCase("RepJtCat")) Then
      wOdSalPrcOpt = "0"
    End If
    
    If (ADC("UoYN14")) = "Y" Or UCase(ADC.MenuCd) = UCase("RepOrdCat") Or UCase(ADC.MenuCd) = UCase("RepFgCat") Or UCase(ADC.MenuCd) = UCase("RepJtCat") Then
      wRmValOpt = "sum(" + IIF(wb_InBaseCur, " (Case when OmMulDiv= 'M' then OrSalVal/OmCnvFct else OrSalVal*OmCnvFct end) ", "OrSalVal") + ")"
    ElseIf (ADC("UoYN14")) = "N" And (UCase(ADC.MenuCd) <> UCase("RepOrdCat") And UCase(ADC.MenuCd) <> UCase("RepFgCat") And UCase(ADC.MenuCd) <> UCase("RepJtCat")) Then
      wRmValOpt = "0"
    End If
    '*** Jay 3.3(FG)
    
    '*** Jay 3.3(FG)(Bef)
    'If (ADC("UoYN2")) = "Y" Or UCase(ADC.MenuCd) = UCase("RepOrdCat") Then
    '  wOdSalPrcOpt = IIF(wb_InBaseCur, " (Case when OmMulDiv= 'M' then OdSalPrc/OmCnvFct else OdSalPrc*OmCnvFct end) ", "OdSalPrc")
    'ElseIf (ADC("UoYN2")) = "N" And UCase(ADC.MenuCd) <> UCase("RepOrdCat") Then
    '  wOdSalPrcOpt = "0"
    'End If
    'If (ADC("UoYN14")) = "Y" Or UCase(ADC.MenuCd) = UCase("RepOrdCat") Then
    '  wRmValOpt = "sum(" + IIF(wb_InBaseCur, " (Case when OmMulDiv= 'M' then OrSalVal/OmCnvFct else OrSalVal*OmCnvFct end) ", "OrSalVal") + ")"
    'ElseIf (ADC("UoYN14")) = "N" And UCase(ADC.MenuCd) <> UCase("RepOrdCat") Then
    '  wRmValOpt = "0"
    'End If
    '*** Jay 3.3(FG)(Bef)
    
    If (ADC("UoYN4")) = "Y" Then
        wRmSetValOpt = "sum(" + IIF(wb_InBaseCur, " (Case when OmMulDiv= 'M' then OrSetSalVal/OmCnvFct else OrSetSalVal*OmCnvFct end) ", "OrSetSalVal") + ") as OrSetSalVal"
        wLabValOpt = "sum(" + IIF(wb_InBaseCur, " (Case when OmMulDiv= 'M' then OlSalVal/OmCnvFct else OlSalVal*OmCnvFct end) ", "OlSalVal") + ") as OlSalVal"
        wLabRtOpt = "max(" + IIF(wb_InBaseCur, " (Case when OmMulDiv= 'M' then OlSalRt/OmCnvFct else OlSalRt*OmCnvFct end) ", "OlSalRt") + ") as OlSalRt"
        If moCn.RecSeek("Select PmCd from Param where PTyp = 'SYS' and PmCd = 'ECH'") <> False Then GRP_REP.SetFormula Rep, "wLabRtYN", "'N'"
    Else
      wRmSetValOpt = "0.00 as OrSetSalVal": wLabValOpt = "0.00 as OlSalVal": wLabRtOpt = "0.00 as OlSalRt"
      If moCn.RecSeek("Select PmCd from Param where PTyp = 'SYS' and PmCd = 'ECH'") <> False Then GRP_REP.SetFormula Rep, "wLabRtYN", "'Y'"
    End If
    
   ' If adc("UoLabGrp") = "N" Then wRmSetValOpt = "0.00 as OrSetSalVal"
   
 
     GRP_REP.SetFormula Rep, "wSuppresSetVal", IIF(ADC("UoLabGrp") = "N", "'Y'", "'N'")
    
   
   ' *** Jay 2.14Next *** [Show Rm Value]
    
    If ADC("UoYN2") = "Y" Then
    ' bef 3.9.0
    '        Rep.TxtLmg.Suppress = False
    '        Rep.TxtLmp.Suppress = False
    '        Rep.FldLmg.Suppress = False
    '        Rep.FldLmp.Suppress = False
        
    ' ***** Manali 3.9.0 - Lme Selection
        Dim wLme1 As String, wLme2 As String
        wLme1 = ADC("UoLme1"): wLme2 = ADC("UoLme2")
        If wLme1 <> "" Then
          Rep.TxtLmg.SetText wLme1 + " Sales"
          GRP_REP.SetFormula Rep, "wLme1", "{rdo.Om" + wLme1 + "Sal}"
        Else
          Rep.TxtLmg.Suppress = True
        End If
        If wLme2 <> "" Then
          Rep.TxtLmp.SetText wLme2 + " Sales"
          GRP_REP.SetFormula Rep, "wLme2", "{rdo.Om" + wLme2 + "Sal}"
        Else
          Rep.TxtLmp.Suppress = True
        End If
        ' ***** Manali 3.9.0 - Lme Selection
        ' ***** Manali 3.10.0 - 02/04/12 - When No Lme Selected G,P Displayed (Original Behaviour)
        If wLme1 = "" And wLme2 = "" Then
          Rep.TxtLmg.Suppress = False: Rep.TxtLmp.Suppress = False
          GRP_REP.SetFormula Rep, "wLme1", "{rdo.OmLmgSal}"
          GRP_REP.SetFormula Rep, "wLme2", "{rdo.OmLmPSal}"
        End If
        ' ***** Manali 3.10.0 - 02/04/12 - When No Lme Selected G,P Displayed (Original Behaviour)
    Else
        Rep.TxtLmg.Suppress = True
        Rep.TxtLmp.Suppress = True
        Rep.FldLmg.Suppress = True
        Rep.FldLmp.Suppress = True
    End If
      
    ' **** Manali 3.6.0 - 04/11/09 - [OdDelDt] added for Excel Report
    If ADC("UoYN5") = "Y" Then
      wDelDt = "max(OmExpDelDt) as OmDelDt, max(OdExpDelDt) as OdDelDt "
      wDelStr = "Exp"
    Else
      wDelDt = "max(OmDelDt) as OmDelDt, max(OdDelDt) as OdDelDt"
      wDelStr = "Prd"
    End If
     
  '  If UCase(adc.MenuCd) = UCase("RepAvgOrd") Then
  '    wOdSalPrcOpt = IIf(adc("UoYN2") = "Y", "OdSalPrc", "0")
  '    wRmValOpt = IIf(adc("UoYN2") = "Y", "sum(OrSalVal)", "0")
  '    wRmSetValOpt = IIf(adc("UoYN2") = "Y", "sum(OrSetSalVal)", "0")
  '    wLabValOpt = IIf(adc("UoYN2") = "Y", "sum(OlSalVal)", "0")
  '  Else
  '    wOdSalPrcOpt = "OdSalPrc"
  '    wRmValOpt = "sum(OrSalVal)"
  '    wRmSetValOpt = "sum(OrSetSalVal)"
  '    wLabValOpt = "sum(OlSalVal)"
  '  End If
  '*****************************************urmi*************************************
    
    If ADC("UoYN3") = "Y" Then
      ws_RmQty = "OrPrdQty": ws_RmWt = "OrPrdWt"
    Else
      ws_RmQty = "OrQty": ws_RmWt = "OrWt"
    End If
    
    '****** Sachin 3.02.0 - Additional Conditions For New Fields
    wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(ADC("UOCMCDFR")) <> "", " And ", "") + IIF(Trim(ADC("UOCMCDFR")) <> "", " OdOmCmCd >= '" + Trim(ADC("UOCMCDFR")) + "'", "")
    wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(ADC("UOCMCDTO")) <> "", " And ", "") + IIF(Trim(ADC("UOCMCDTO")) <> "", " OdOmCmCd <= '" + Trim(ADC("UOCMCDTO")) + "'", "")
    '****** Sachin 3.02.0 - Additional Conditions For New Fields
    
    '*************************** ZUBIN **************************
    ' 08th Dec 2003, EMR206
  '   wOurRmCd = "Max(OrRmCd) as OrRmCd"
  '   wOurDsgCd = "OdDmCd as CdOurDmCd"
    '*** Jay 3.3(FG)
    ' ***** Manali Trading Module
    If UCase(ADC.MenuCd) = UCase("RepDetOrd") Or UCase(ADC.MenuCd) = UCase("RepAvgOrd") Or _
       UCase(ADC.MenuCd) = UCase("RepFgDet") Or UCase(ADC.MenuCd) = UCase("RepFgAvg") Or _
       UCase(ADC.MenuCd) = UCase("RepJtDet") Or UCase(ADC.MenuCd) = UCase("RepJtAvg") Then
    'If UCase(ADC.MenuCd) = UCase("RepDetOrd") Or UCase(ADC.MenuCd) = UCase("RepAvgOrd") Then
    '*** Jay 3.3(FG)
      If UCase(ADC("UoYN7")) = "Y" Then
        wOurDsgCd = "IsNull((Select cast(CdCustDmCd as Varchar(17)) from CustDsg " + _
                          "Where " + IIF(gs_Partition = ctCurrPrtn, " CdPrtKey='" + ctCurrPrtn + "' and ", "") + " CdCtg = 'C' and CdCd = Max(OmCmCd) and " + _
                          "CdOurDmCd = Max(OdDmCd) and CdOurSfx = Max(OdSfx) and " + _
                          "CdOurDmSz = Max(OdDmSz)), " + _
                        "IsNull((Select cast(CdCustDmCd as Varchar(17)) from CustDsg " + _
                              "where " + IIF(gs_Partition = ctCurrPrtn, " CdPrtKey='" + ctCurrPrtn + "' and ", "") + " CdCtg = 'C' and CdCd = Max(OmCmCd) and " + _
                              "CdOurDmCd = Max(OdDmCd) and CdOurSfx = Max(OdSfx) and " + _
                              "CdOurDmSz = '' ), " + _
                            "IsNull((Select cast(CdCustDmCd as Varchar(17)) from CustDsg " + _
                                  "where " + IIF(gs_Partition = ctCurrPrtn, " CdPrtKey='" + ctCurrPrtn + "' and ", "") + " CdCtg = 'C' and CdCd = Max(OmCmCd) and " + _
                                  "CdOurDmCd = Max(OdDmCd) and CdOurSfx = '' and " + _
                                  "CdOurDmSz = Max(OdDmSz) ), " + _
                                "IsNull((Select cast(CdCustDmCd as Varchar(17)) from CustDsg " + _
                                      "where " + IIF(gs_Partition = ctCurrPrtn, " CdPrtKey='" + ctCurrPrtn + "' and ", "") + " CdCtg = 'C' and CdCd = Max(OmCmCd) and " + _
                                      "CdOurDmCd = Max(OdDmCd) and CdOurSfx = '' and " + _
                                      "CdOurDmSz= ''), '{' + max(RTrim(OdDmCd)) + '}')))) as CdOurDmCd"
      Else
        wOurDsgCd = "Max(OdDmCd) as CdOurDmCd"
      End If
    '*** Jay 3.3(FG)
    ' ***** Manali Trading Module
    ElseIf UCase(ADC.MenuCd) = UCase("RepOrdCat") Or UCase(ADC.MenuCd) = UCase("RepFgCat") Or UCase(ADC.MenuCd) = UCase("RepJtCat") Then
    'ElseIf UCase(ADC.MenuCd) = UCase("RepOrdCat") Then
    '*** Jay 3.3(FG)
      If UCase(ADC("UoYN7")) = "Y" Then
        wOurDsgCd = "IsNull((Select cast(CdCustDmCd as Varchar(17)) from CustDsg " + _
                          "Where " + IIF(gs_Partition = ctCurrPrtn, " CdPrtKey='" + ctCurrPrtn + "' and ", "") + " CdCtg = 'C' and CdCd = OmCmCd and " + _
                          "CdOurDmCd = OdDmCd and CdOurSfx = OdSfx and " + _
                          "CdOurDmSz = OdDmSz), " + _
                        "IsNull((Select cast(CdCustDmCd as Varchar(17)) from CustDsg " + _
                              "where " + IIF(gs_Partition = ctCurrPrtn, " CdPrtKey='" + ctCurrPrtn + "' and ", "") + " CdCtg = 'C' and CdCd = OmCmCd and " + _
                              "CdOurDmCd = OdDmCd and CdOurSfx = OdSfx and " + _
                              "CdOurDmSz = ''), " + _
                            "IsNull((Select cast(CdCustDmCd as Varchar(17)) from CustDsg " + _
                                  "where " + IIF(gs_Partition = ctCurrPrtn, " CdPrtKey='" + ctCurrPrtn + "' and ", "") + " CdCtg = 'C' and CdCd = OmCmCd and " + _
                                  "CdOurDmCd = OdDmCd and CdOurSfx = '' and " + _
                                  "CdOurDmSz = OdDmSz), " + _
                                "IsNull((Select cast(CdCustDmCd as Varchar(17)) from CustDsg " + _
                                      "where " + IIF(gs_Partition = ctCurrPrtn, " CdPrtKey='" + ctCurrPrtn + "' and ", "") + " CdCtg = 'C' and CdCd = OmCmCd and " + _
                                      "CdOurDmCd = OdDmCd and CdOurSfx = '' and " + _
                                      "CdOurDmSz = ''), '{' + OdDmCd + '}')))) as CdOurDmCd"
      Else
        wOurDsgCd = "OdDmCd as CdOurDmCd"
      End If
    End If
    
    '*** Jay 3.3(FG)
    ' ***** Manali Trading Module
    If (UCase(ADC.MenuCd) = UCase("RepDetOrd") Or UCase(ADC.MenuCd) = UCase("RepFgDet") Or UCase(ADC.MenuCd) = UCase("RepJtDet")) And UCase(ADC("UoYN8")) = "Y" Then
    'If UCase(ADC.MenuCd) = UCase("RepDetOrd") And UCase(ADC("UoYN8")) = "Y" Then
    '*** Jay 3.3(FG)
      Select Case UCase(ADC("UoRmGrp"))
        Case Is = UCase("Sr"), UCase("RmRSzRt")
'          wOurRmCd = "IsNull((Select cast(CrCustRmCd as VarChar(14)) from CustRm where " + IIF(gs_Partition = ctCurrPrtn, " CrPrtKey='" + ctCurrPrtn + "' and ", "") + " CrCtg = 'C' and " + _
'                  "CrCd = Max(OmCmCd) and CrRmCtg = Max(OrRmCtg) and " + _
'                  "CrOurRmCd = Max(OrRmCd) and Max(OrLn1) Between CrFrLn and CrToLn)," + _
'                  "'{' + Max(OrRmCd) + '}') as OrRmCd"
            wOurRmCd = " (Case when Max(OrCustRmCd) ='' Then '{'+Max(OrRmCd)+'}' Else Max(Cast(OrCustRmCd as varChar(14))) End) as OrRmCd "
        Case Is = UCase("RmCdRt"), UCase("RmCd")
'            wOurRmCd = "IsNull((Select cast(CrCustRmCd as VarChar(14)) from CustRm where " + IIF(gs_Partition = ctCurrPrtn, " CrPrtKey='" + ctCurrPrtn + "' and ", "") + " CrCtg = 'C' and " + _
'                  "CrCd = Max(OmCmCd) and CrRmCtg = Max(OrRmCtg) and " + _
'                  "CrOurRmCd = Max(OrRmCd) and 0 Between CrFrLn and CrToLn)," + _
'                  "'{' + Max(OrRmCd) + '}') as OrRmCd"
            wOurRmCd = " (Case when Max(OrCustRmCd) ='' Then '{'+Max(OrRmCd)+'}' Else Max(Cast(OrCustRmCd as varChar(14))) End) as OrRmCd "
      End Select
    Else
      wOurRmCd = "Max(OrRmCd) as OrRmCd"
    End If
    
    ' ***** Manali 3.9.0 - For Trading - Base/Ref Vch displayed in report
    Dim ws_BaseRefJoin As String, ws_BaseRefFlds As String, ws_BaseRef As String
    ws_BaseRefJoin = "": ws_BaseRef = ""
    If UCase(ADC.MenuCd) = UCase("RepJtDet") Then
      ws_BaseRef = moCn.GetFldVal(" Select (Case When TmBaseTcs<> '' Then 'OdBaseIdKey' Else (Case When TmRefTcs<>'' Then  'OdRefIdKey' Else '' End) End) From TcMst Where TmCoCd='" + ADC("UoCoCdFr") + "' and TmTc='" + ADC("UoOmTcFr") + "' ")
    End If
    If ws_BaseRef <> "" Then
        ws_BaseRefJoin = "Left Outer Join (Select OdCoCd as qOdCoCd, OdTc as qOdTc, OdYy as qOdYY, OdChr as qOdChr, OdNo as qOdNo, OdSr as qOdSr, " + _
                       " OdPrtKey as qOdPrtKey, OdIdNo as qOdIdNo From OrdDsg) BOd On qOdCoCd=OdCoCd and qOdIdNo=" + ws_BaseRef + " "
      
        ws_BaseRefFlds = " Max(IsNull(qOdTc, '')) as IdExpOdTc, max(IsNull(qOdYY, '')) as IdExpOdYy, max(IsNull(qOdChr, '')) as IdExpOdChr, Max(IsNull(qOdNo, 0)) as IdExpOdNo, Max(IsNull(qOdSr, 0)) as IdExpOdSr, "
    Else
      ws_BaseRefFlds = " '' as IdExpOdTc, '' as IdExpOdYy, '' as IdExpOdChr, 0 as IdExpOdNo, 0 as IdExpOdSr, "
    End If
    ' ***** Manali 3.9.0 - For Trading - Base/Ref Vch displayed in report
      
    '*********************************************Geeta****************************************
    '*********************Assigning Group Desc & group formula ************************************************************
    '*** For the Group Formula of the report
    Dim wGrpArr1() As String, wGrpFormula1 As String, wGrpDescFormula1 As String
    Dim wGrpArr2() As String, wGrpFormula2 As String, wGrpDescFormula2 As String
    Dim i As Integer, qDmCtgDesc As String, qRmCtgDesc As String, qDmColDesc As String
    Dim qVaCtgDesc As String, qMainMet As String, wOrJoin As String
    ' Zubin 212
    Dim qDmSalCtgDesc As String
    
    qDmCtgDesc = "'' as qDmCtgDesc": qRmCtgDesc = "'' as qRmCtgDesc"
    qDmColDesc = "'' as qDmColDesc": qVaCtgDesc = "'' as qVaCtgDesc"
    ' Zubin 212
    qDmSalCtgDesc = "'' as qDmSalCtgDesc"
    
    qMainMet = "space(2) as OrRmCtg": wOrJoin = ""
  
    wGrpArr1 = Split(moCn.GetFldVal("Select PDesc225 from Param where PTyp= 'INGR' and " + _
           "PMCd= '" + ADC("UoInGrp") + "' "), ",")
    For i = 0 To UBound(wGrpArr1)
      Select Case UCase(wGrpArr1(i))
       Case Is = "RMCTG"
          '*** Jay 3.3(FG)
          ' ***** Manali Trading Module
          If UCase(ADC.MenuCd) = UCase("RepDetOrd") Or UCase(ADC.MenuCd) = UCase("RepAvgOrd") Or UCase(ADC.MenuCd) = UCase("RepFgDet") Or UCase(ADC.MenuCd) = UCase("RepFgAvg") _
              Or UCase(ADC.MenuCd) = UCase("RepJtDet") Or UCase(ADC.MenuCd) = UCase("RepJtAvg") Then
          'If UCase(ADC.MenuCd) = UCase("RepDetOrd") Or UCase(ADC.MenuCd) = UCase("RepAvgOrd") Then
          '*** Jay 3.3(FG)
            qRmCtgDesc = "IsNull((Select PDesc from Param where PTyp= 'RMCTG' and PMCd= " + _
                       "(select RmCtg from RmMst join OrdRm on RmCd=OrRmCd " + IIF(gs_Partition = ctCurrPrtn, " and RmPrtKey=OrPrtKey ", "") + " and  OrMainMet='Y' " + _
                       "Where OrCocd = Max(OmCoCd) And OrTc = OmTc  and OrYy = OmYy and OrChr=OmChr " + _
                       "and OrNo=OmNo and OrSr=OdSr " + IIF(gs_Partition = ctCurrPrtn, " and OrPrtKey= max(OmPrtKey) ", "") + " )),'') as qRmCtgDesc"
          Else
            qRmCtgDesc = "IsNull((Select PDesc from Param where PTyp= 'RMCTG' and PMCd= " + _
                       "(select RmCtg from RmMst,OrdRm where RmCd=OrRmCd " + IIF(gs_Partition = ctCurrPrtn, " and RmPrtKey=OrPrtKey ", "") + "  and  OrMainMet='Y'  " + _
                       "and OrCocd = OmCocd And OrTc = OmTc  and OrYy = OmYy and OrChr=OmChr  " + _
                       "and OrNo=OmNo and OrSr=OdSr " + IIF(gs_Partition = ctCurrPrtn, " and OrPrtKey=OdPrtKey ", "") + ")),'') as qRmCtgDesc"
          End If
            wGrpDescFormula1 = wGrpDescFormula1 + IIF(wGrpDescFormula1 = "", "", "+', '+") + "Trim({rdo.qRmCtgDesc})"
            'qMainMet = "max(OrRmCtg) as OrRmCtg"
            'wOrJoin = " join OrdRm on OrCoCd= OdCoCd and OrTc= OdTc and OrYy= OdYy and OrChr= OdChr and OrNo= OdNo and OrSr= OdSr and OrMainMet= 'Y' "
       Case Is = "KT"
            wGrpDescFormula1 = wGrpDescFormula1 + IIF(wGrpDescFormula1 = "", "", "+', '+") + "{rdo.IdKt}"
       Case Is = "DMCOL"
          '*** Jay 3.3(FG)
          ' ***** Manali Trading Module
          If UCase(ADC.MenuCd) = UCase("RepDetOrd") Or UCase(ADC.MenuCd) = UCase("RepAvgOrd") Or UCase(ADC.MenuCd) = UCase("RepFgDet") Or UCase(ADC.MenuCd) = UCase("RepFgAvg") _
            Or UCase(ADC.MenuCd) = UCase("RepJtDet") Or UCase(ADC.MenuCd) = UCase("RepJtAvg") Then
          'If UCase(ADC.MenuCd) = UCase("RepDetOrd") Or UCase(ADC.MenuCd) = UCase("RepAvgOrd") Then
          '*** Jay 3.3(FG)
            qDmColDesc = "(Select PDesc from Param where PTyp= 'DMCOL' and PMCd= max(OdDmCol)) as qDmColDesc"
          Else
            qDmColDesc = "(Select PDesc from Param where PTyp= 'DMCOL' and PMCd= OdDmCol) as qDmColDesc"
          End If
            wGrpDescFormula1 = wGrpDescFormula1 + IIF(wGrpDescFormula1 = "", "", "+', '+") + "Trim({rdo.qDmColDesc})"
       Case Is = "VACTG"
           '*** Jay 3.3(FG)
           ' ***** Manali Trading Module
          If UCase(ADC.MenuCd) = UCase("RepDetOrd") Or UCase(ADC.MenuCd) = UCase("RepAvgOrd") Or UCase(ADC.MenuCd) = UCase("RepFgDet") Or UCase(ADC.MenuCd) = UCase("RepFgAvg") _
            Or UCase(ADC.MenuCd) = UCase("RepJtDet") Or UCase(ADC.MenuCd) = UCase("RepJtAvg") Then
          'If UCase(ADC.MenuCd) = UCase("RepDetOrd") Or UCase(ADC.MenuCd) = UCase("RepAvgOrd") Then
          '*** Jay 3.3(FG)
            qVaCtgDesc = "(Select PDesc from Param where PTyp= 'VACTG' and PMCd= max(OdVaCtg)) as qVaCtgDesc"
          Else
            qVaCtgDesc = "(Select PDesc from Param where PTyp= 'VACTG' and PMCd= OdVaCtg) as qVaCtgDesc"
          End If
            wGrpDescFormula1 = wGrpDescFormula1 + IIF(wGrpDescFormula1 = "", "", "+', '+") + "Trim({rdo.qVaCtgDesc})"
       Case Is = "PCS/PRS"
          wGrpDescFormula1 = wGrpDescFormula1 + IIF(wGrpDescFormula1 = "", "", "+', '+") + "{rdo.DmUom}"
       Case Is = "DMCTG"
           '*** Jay 3.3(FG)
           ' ***** Manali Trading Module
          If UCase(ADC.MenuCd) = UCase("RepDetOrd") Or UCase(ADC.MenuCd) = UCase("RepAvgOrd") Or UCase(ADC.MenuCd) = UCase("RepFgDet") Or UCase(ADC.MenuCd) = UCase("RepFgAvg") _
            Or UCase(ADC.MenuCd) = UCase("RepJtDet") Or UCase(ADC.MenuCd) = UCase("RepJtAvg") Then
          'If UCase(ADC.MenuCd) = UCase("RepDetOrd") Or UCase(ADC.MenuCd) = UCase("RepAvgOrd") Then
          '*** Jay 3.3(FG)
            qDmCtgDesc = " (Select PDesc from Param where PTyp= 'DMCTG' and PMCd= max(DmCtg)) as qDmCtgDesc"
          Else
            qDmCtgDesc = " (Select PDesc from Param where PTyp= 'DMCTG' and PMCd=DmCtg) as qDmCtgDesc"
          End If
          wGrpDescFormula1 = wGrpDescFormula1 + IIF(wGrpDescFormula1 = "", "", "+', '+") + "Trim({rdo.qDmCtgDesc})"
      Case Is = "GLDLS"
           wGrpDescFormula1 = wGrpDescFormula1 + IIF(wGrpDescFormula1 = "", "", "+', '+") + "'@'+ToText({rdo.IdGldLs}, '##0.000')+ '%'"
      ' Zubin 212
      Case Is = "SALCTG"
        '*** Jay 3.3(FG)
        ' ***** Manali Trading Module
        If UCase(ADC.MenuCd) = UCase("RepDetOrd") Or UCase(ADC.MenuCd) = UCase("RepAvgOrd") Or UCase(ADC.MenuCd) = UCase("RepFgDet") Or UCase(ADC.MenuCd) = UCase("RepFgAvg") _
          Or UCase(ADC.MenuCd) = UCase("RepJtDet") Or UCase(ADC.MenuCd) = UCase("RepJtAvg") Then
        'If UCase(ADC.MenuCd) = UCase("RepDetOrd") Or UCase(ADC.MenuCd) = UCase("RepAvgOrd") Then
        '*** Jay 3.3(FG)
          qDmSalCtgDesc = " (Select PDesc from Param where PTyp= 'SALCTG' and PMCd= max(DmCtg) and PSCd= max(DmSalCtg)) as qDmSalCtgDesc "
        Else
          qDmSalCtgDesc = " (Select PDesc from Param where PTyp= 'SALCTG' and PMCd= DmCtg and PSCd= DmSalCtg) as qDmSalCtgDesc "
        End If
        wGrpDescFormula1 = wGrpDescFormula1 + IIF(wGrpDescFormula1 = "", "", "+', '+") + "Trim({rdo.qDmSalCtgDesc})"
      ' Zubin 212
      End Select
      wGrpFormula1 = wGrpFormula1 + IIF(wGrpFormula1 = "", "", "+") + "Trim(UpperCase(ToText({rdo." + moCn.GetFldVal("Select PDesc225 from Param where PTyp= 'INGRPRCP' and PMCd= '" + wGrpArr1(i) + "'") + "})))"
      ' **** Manali 3.6.0 - 04/11/09 - Order By clause for Excel Report
      If ADC("UoYN23") = "Y" Then ws_OrdBy = ws_OrdBy + IIF(ws_OrdBy = "", "", ",") + moCn.GetFldVal("Select PDesc225 from Param where PTyp= 'INGRPRCP' and PMCd= '" + wGrpArr1(i) + "'")
      
        If UCase(ADC.MenuCd) = UCase("RepOrdCat") Or UCase(ADC.MenuCd) = UCase("RepFgCat") Or UCase(ADC.MenuCd) = UCase("RepJtCat") Then
           wExcelOrdBy = wExcelOrdBy + IIF(wExcelOrdBy = "", "", ",") + moCn.GetFldVal("Select (Case PDesc225 when 'IdKt' Then 'OdKT' when 'IdGldLS' " + _
                       " Then '(Select PNum from Param where PTyp=''VACTG'' and PMCd=OdVaCtg and PSCd='''')' " + _
                       " When 'IdVaCtg' Then 'OdVaCtg' " + _
                       " When 'OrRmCtg' Then 'IsNull((select max(RmCtg) from RmMst ,OrdRm where " + IIF(gs_Partition = ctCurrPrtn, " RmPrtKey=OrPrtKey and ", "") + " RmCd=OrRmCd  and  " + IIF(gs_Partition = ctCurrPrtn, " OrPrtKey=OmPrtKey and ", "") + _
                       " OrMainMet=''Y'' and OrCocd = OmCocd And OrTc = OmTc and OrYy = OmYy and OrChr=OmChr and OrNo=OmNo and OrSr=OdSr),'''')' Else PDesc225 End)   " + _
                       " from Param where PTyp= 'INGRPRCP' and PMCd= '" + wGrpArr1(i) + "'")
        Else
            wExcelOrdBy = wExcelOrdBy + IIF(wExcelOrdBy = "", "", ",") + moCn.GetFldVal("Select 'Max('+PDesc225+')' from Param where PTyp= 'INGRPRCP' and PMCd= '" + wGrpArr1(i) + "'")
        End If
         
      ' **** Manali 3.6.0 - 04/11/09 - Order By clause for Excel Report
    Next i
  
    wGrpArr2 = Split(moCn.GetFldVal("Select PDesc225 from Param where PTyp= 'INGR' and " + _
           "PMCd= '" + ADC("UoInGrp1") + "' "), ",")
    For i = 0 To UBound(wGrpArr2)
     Select Case UCase(wGrpArr2(i))
      Case Is = "RMCTG"
        '*** Jay 3.3(FG)
        ' ***** Manali Trading Module
        If UCase(ADC.MenuCd) = UCase("RepDetOrd") Or UCase(ADC.MenuCd) = UCase("RepAvgOrd") Or UCase(ADC.MenuCd) = UCase("RepFgDet") Or UCase(ADC.MenuCd) = UCase("RepFgAvg") _
          Or UCase(ADC.MenuCd) = UCase("RepJtDet") Or UCase(ADC.MenuCd) = UCase("RepJtAvg") Then
        'If UCase(ADC.MenuCd) = UCase("RepDetOrd") Or UCase(ADC.MenuCd) = UCase("RepAvgOrd") Then
        '*** Jay 3.3(FG)
           qRmCtgDesc = "IsNull((Select PDesc from Param where PTyp= 'RMCTG' and PMCd= " + _
                     "(select RmCtg from RmMst join OrdRm on RmCd=OrRmCd " + IIF(gs_Partition = ctCurrPrtn, " and RmPrtKey=OrPrtKey ", "") + " and  OrMainMet='Y' " + _
                     "Where OrCocd = Max(OmCocd) And OrTc = OmTc  and OrYy = OmYy and OrChr=OmChr " + _
                     "and OrNo=OmNo and OrSr=OdSr " + IIF(gs_Partition = ctCurrPrtn, " and OrPrtKey=Max(OdPrtKey) ", "") + " )),'') as qRmCtgDesc"
        Else
           qRmCtgDesc = "IsNull((Select PDesc from Param where PTyp= 'RMCTG' and PMCd= " + _
                     "(select RmCtg from RmMst,OrdRm where RmCd=OrRmCd " + IIF(gs_Partition = ctCurrPrtn, " and RmPrtKey=OrPrtKey ", "") + "  and  OrMainMet='Y' " + _
                     "and OrCocd = OmCocd And OrTc = OmTc  and OrYy = OmYy and OrChr=OmChr " + _
                     "and OrNo=OmNo and OrSr=OdSr " + IIF(gs_Partition = ctCurrPrtn, " and OrPrtKey=OmPrtKey ", "") + ")),'') as qRmCtgDesc"
        End If
        wGrpDescFormula2 = wGrpDescFormula2 + IIF(wGrpDescFormula2 = "", "", "+', '+") + "Trim({rdo.qRmCtgDesc})"
        'If adc("UoYN23") = "Y" Then ws_OrdBy = ws_OrdBy + IIF(ws_OrdBy = "", "", ",") + " qRmCtgDesc"
        'qMainMet = "max(OrRmCtg) as OrRmCtg"
     Case Is = "KT"
        wGrpDescFormula2 = wGrpDescFormula2 + IIF(wGrpDescFormula2 = "", "", "+', '+") + "{rdo.IdKt}"
     Case Is = "DMCOL"
        '*** Jay 3.3(FG)
        ' ***** Manali Trading Module
        If UCase(ADC.MenuCd) = UCase("RepDetOrd") Or UCase(ADC.MenuCd) = UCase("RepAvgOrd") Or UCase(ADC.MenuCd) = UCase("RepFgDet") Or UCase(ADC.MenuCd) = UCase("RepFgAvg") _
          Or UCase(ADC.MenuCd) = UCase("RepJtDet") Or UCase(ADC.MenuCd) = UCase("RepJtAvg") Then
        'If UCase(ADC.MenuCd) = UCase("RepDetOrd") Or UCase(ADC.MenuCd) = UCase("RepAvgOrd") Then
        '*** Jay 3.3(FG)
            qDmColDesc = "(Select PDesc from Param where PTyp= 'DMCOL' and PMCd= max(OdDmCol)) as qDmColDesc"
        Else
          qDmColDesc = "(Select PDesc from Param where PTyp= 'DMCOL' and PMCd= OdDmCol) as qDmColDesc"
        End If
        wGrpDescFormula2 = wGrpDescFormula2 + IIF(wGrpDescFormula2 = "", "", "+', '+") + "Trim({rdo.qDmColDesc})"
     Case Is = "VACTG"
        '*** Jay 3.3(FG)
        ' ***** Manali Trading Module
        If UCase(ADC.MenuCd) = UCase("RepDetOrd") Or UCase(ADC.MenuCd) = UCase("RepAvgOrd") Or UCase(ADC.MenuCd) = UCase("RepFgDet") Or UCase(ADC.MenuCd) = UCase("RepFgAvg") _
          Or UCase(ADC.MenuCd) = UCase("RepJtDet") Or UCase(ADC.MenuCd) = UCase("RepJtAvg") Then
        'If UCase(ADC.MenuCd) = UCase("RepDetOrd") Or UCase(ADC.MenuCd) = UCase("RepAvgOrd") Then
        '*** Jay 3.3(FG)
            qVaCtgDesc = "(Select PDesc from Param where PTyp= 'VACTG' and PMCd= max(OdVaCtg)) as qVaCtgDesc"
        Else
            qVaCtgDesc = "(Select PDesc from Param where PTyp= 'VACTG' and PMCd= OdVaCtg) as qVaCtgDesc"
        End If
        wGrpDescFormula2 = wGrpDescFormula2 + IIF(wGrpDescFormula2 = "", "", "+', '+") + "Trim({rdo.qVaCtgDesc})"
    Case Is = "PCS/PRS"
        wGrpDescFormula2 = wGrpDescFormula2 + IIF(wGrpDescFormula2 = "", "", "+', '+") + "{rdo.DmUom}"
        'If adc("UoYN23") = "Y" Then ws_OrdBy = ws_OrdBy + IIF(ws_OrdBy = "", "", ",") + " DmUom"
    Case Is = "DMCTG"
        '*** Jay 3.3(FG)
        ' ***** Manali Trading Module
        If UCase(ADC.MenuCd) = UCase("RepDetOrd") Or UCase(ADC.MenuCd) = UCase("RepAvgOrd") Or UCase(ADC.MenuCd) = UCase("RepFgDet") Or UCase(ADC.MenuCd) = UCase("RepFgAvg") _
          Or UCase(ADC.MenuCd) = UCase("RepJtDet") Or UCase(ADC.MenuCd) = UCase("RepJtAvg") Then
        'If UCase(ADC.MenuCd) = UCase("RepDetOrd") Or UCase(ADC.MenuCd) = UCase("RepAvgOrd") Then
        '*** Jay 3.3(FG)
          qDmCtgDesc = " (Select PDesc from Param where PTyp= 'DMCTG' and PMCd= max(DmCtg)) as qDmCtgDesc"
        Else
          qDmCtgDesc = " (Select PDesc from Param where PTyp= 'DMCTG' and PMCd= DmCtg) as qDmCtgDesc"
        End If
        wGrpDescFormula2 = wGrpDescFormula2 + IIF(wGrpDescFormula2 = "", "", "+', '+") + "Trim({rdo.qDmCtgDesc})"
    Case Is = "GLDLS"
        wGrpDescFormula2 = wGrpDescFormula2 + IIF(wGrpDescFormula2 = "", "", "+', '+") + "'@'+ToText({rdo.IdGldLs}, '##0.000')+ '%'"
    ' Zubin 212
    Case Is = "SALCTG"
      '*** Jay 3.3(FG)
      ' ***** Manali Trading Module
      If UCase(ADC.MenuCd) = UCase("RepDetOrd") Or UCase(ADC.MenuCd) = UCase("RepAvgOrd") Or UCase(ADC.MenuCd) = UCase("RepFgDet") Or UCase(ADC.MenuCd) = UCase("RepFgAvg") _
        Or UCase(ADC.MenuCd) = UCase("RepJtDet") Or UCase(ADC.MenuCd) = UCase("RepJtAvg") Then
      'If UCase(ADC.MenuCd) = UCase("RepDetOrd") Or UCase(ADC.MenuCd) = UCase("RepAvgOrd") Then
      '*** Jay 3.3(FG)
        qDmSalCtgDesc = " (Select PDesc from Param where PTyp= 'SALCTG' and PMCd= max(DmCtg) and PSCd= max(DmSalCtg)) as qDmSalCtgDesc "
      Else
        qDmSalCtgDesc = " (Select PDesc from Param where PTyp= 'SALCTG' and PMCd= DmCtg and PSCd= DmSalCtg) as qDmSalCtgDesc "
      End If
      wGrpDescFormula2 = wGrpDescFormula2 + IIF(wGrpDescFormula2 = "", "", "+', '+") + "Trim({rdo.qDmSalCtgDesc})"
    ' Zubin 212
    End Select
    wGrpFormula2 = wGrpFormula2 + IIF(wGrpFormula2 = "", "", "+") + "Trim(UpperCase(ToText({rdo." + moCn.GetFldVal("Select PDesc225 from Param where PTyp= 'INGRPRCP' and PMCd= '" + wGrpArr2(i) + "'") + "})))"
     ' **** Manali 3.6.0 - 04/11/09 - Order By clause for Excel Report
     If UCase(ADC.MenuCd) = UCase("RepOrdCat") Or UCase(ADC.MenuCd) = UCase("RepFgCat") Or UCase(ADC.MenuCd) = UCase("RepJtCat") Then
        wInGrp2 = moCn.GetFldVal("Select (Case PDesc225 when 'IdKt' Then 'OdKT' when 'IdGldLS' " + _
                       " Then '(Select PNum from Param where PTyp=''VACTG'' and PMCd=OdVaCtg and PSCd='''')' " + _
                       " When 'IdVaCtg' Then 'OdVaCtg' " + _
                       " When 'OrRmCtg' Then 'IsNull((select max(RmCtg) from RmMst ,OrdRm where " + IIF(gs_Partition = ctCurrPrtn, " RmPrtKey=OrPrtKey and ", "") + " RmCd=OrRmCd  and  " + IIF(gs_Partition = ctCurrPrtn, " OrPrtKey=OmPrtKey and ", "") + _
                       " OrMainMet=''Y'' and OrCocd = OmCocd And OrTc = OmTc and OrYy = OmYy and OrChr=OmChr and OrNo=OmNo and OrSr=OdSr),'''')' Else PDesc225 End)  from Param where PTyp= 'INGRPRCP' and PMCd= '" + wGrpArr2(i) + "'")
     Else
       If ADC("UoYN23") = "Y" Then
            wInGrp2 = moCn.GetFldVal("Select PDesc225 from Param where PTyp= 'INGRPRCP' and PMCd= '" + wGrpArr2(i) + "'")
            ws_OrdBy = ws_OrdBy + IIF(ws_OrdBy <> "" And InStr(1, ws_OrdBy, wInGrp2) = 0, ",", "") + IIF(InStr(1, ws_OrdBy, wInGrp2) = 0, wInGrp2, "")
       Else
            wInGrp2 = moCn.GetFldVal("Select 'Max('+PDesc225+')' from Param where PTyp= 'INGRPRCP' and PMCd= '" + wGrpArr2(i) + "'")
       End If
     End If
     ' **** Manali 3.6.0 - 04/11/09 - Order By clause for Excel Report
     'wExcelOrdBy = wExcelOrdBy + IIF(wExcelOrdBy <> "" And InStr(1, wExcelOrdBy, "Max(" + wInGrp2 + ")") = 0, ",", "") + IIF(InStr(1, wExcelOrdBy, "Max(" + wInGrp2 + ")") = 0, "Max(" + wInGrp2 + ")", "")
     wExcelOrdBy = wExcelOrdBy + IIF(wExcelOrdBy <> "" And InStr(1, wExcelOrdBy, wInGrp2) = 0, ",", "") + IIF(InStr(1, wExcelOrdBy, wInGrp2) = 0, wInGrp2, "")
     
    Next i
    '****************************************************************************************
  
    '*** Jay 2.14 ***
    '*** Jay 3.3(FG)
    ' ***** Manali Trading Module
    If UCase(ADC.MenuCd) <> UCase("RepOrdCat") And UCase(ADC.MenuCd) <> UCase("RepFgCat") And UCase(ADC.MenuCd) <> UCase("RepJtCat") Then
    'If UCase(ADC.MenuCd) <> UCase("RepOrdCat") Then
    '*** Jay 3.3(FG)
      If ADC("UOYN13") = "Y" And Not (ADC("UoInGrp") = "" And ADC("UoInGrp1") = "") Then
        mRep.Grp5H.Suppress = False
      Else
        mRep.Grp5H.Suppress = True
      End If
    End If
    
    If ADC("UOYN11") = "Y" Then
      'wGrpFormula2= wGrpFormula2 + IIF(wGrpFormula2 = "", "", "+") + "Trim(UpperCase(ToText({rdo.OdDmCd})))"
      GRP_REP.SetFormula mRep, "OdSr", "Trim(UpperCase(ToText({rdo.OdDmCd})))+ '/'+ ToText ({rdo.OdSr}, '####0') "
      
      ' **** Manali 3.6.0 - 04/11/09 - Order By clause for Excel Report
      If ADC("UoYN23") = "Y" Then ws_OrdBy = ws_OrdBy + IIF(ws_OrdBy = "", "", ",") + " OdDmCd"
      wExcelOrdBy = wExcelOrdBy + IIF(wExcelOrdBy = "", "", ",") + IIF(UCase(ADC.MenuCd) = UCase("RepOrdCat") Or UCase(ADC.MenuCd) = UCase("RepFgCat") Or UCase(ADC.MenuCd) = UCase("RepJtCat"), " OdDmCd ", " Max(OdDmCd)")
      ' **** Manali 3.6.0 - 04/11/09 - Order By clause for Excel Report
    '*** 2.11 (Jen)
    Else
      ''''wGrpFormula2 = wGrpFormula2 + IIF(wGrpFormula2 = "", "", "+") + "ToText({rdo.OdSr})"
      'wGrpFormula2 = wGrpFormula2 + IIF(wGrpFormula2 = "", "", "+") + "{rdo.OmYy} +'/'+ {rdo.OmTc} +'/'+ {rdo.OmChr} +'/'+ ToText ({rdo.OmNo}, '######0') +'/'+ ToText ({rdo.OdSr}, '####0')"
      GRP_REP.SetFormula mRep, "OdSr", "ToText ({rdo.OdSr}, '####0') "
    '*** 2.11 (Jen)
    End If
    '*** Jay 2.14 ***
   
    '*** Jay 3.3(FG)
    ' ***** Manali Trading Module
    If UCase(ADC.MenuCd) = UCase("RepDetOrd") Or UCase(ADC.MenuCd) = UCase("RepAvgOrd") Or UCase(ADC.MenuCd) = UCase("RepFgDet") Or UCase(ADC.MenuCd) = UCase("RepFgAvg") _
      Or UCase(ADC.MenuCd) = UCase("RepJtDet") Or UCase(ADC.MenuCd) = UCase("RepJtAvg") Then
    'If UCase(ADC.MenuCd) = UCase("RepDetOrd") Or UCase(ADC.MenuCd) = UCase("RepAvgOrd") Then
    '*** Jay 3.3(FG)
    '*******************************Geeta**********************************
    ' 23rd Dec,2003 substituted Group Combination fields
    'OdVaCtg,OdKt,DmUom,IdVactg,qRmCtgDesc, qDmCtgDesc,qDmColDesc , qVaCtgDesc,qMainMetRmCtg
    '  max(OdDmCd) Replaced with " + wOurDsgCd + "
      ' Zubin 212 (DmSalCtg and qDmSalCtgDesc added)
      ' Zubin 213 (BaseCur added)
      '*** Jay 3.2.0 [PicNm] '*** Jay 3.2.0 [OdPoNo]
      ' **** Manali 3.03 - 29/08/06 - Exp Ord fields added
      '*** Jay 3.3(FG) [OdDmId]
      ' ***** Manali 3.4.1 - 12/11/08 - [' ' AS qOdOmPoNo] added
      ' ***** Manali 3.6.0 - OmPoDt added
      ' ***** Manali 3.6.1 - OdGldAs added
      ' ***** Manali 3.9.0 - For Trading - Base/Ref Vch displayed in report - ws_BaseRefFlds
          Dim wTblStr As String
    
    wTblStr = "Select OmCoCd, OmYy, OmTc, OmChr, OmNo,OmCmCd,OmDt, OmPONo,OmPoDt,OmLmgCst,OmLmgSal,OmLmpSal,OmLmsSal,OmLmlSal,OmDelDt,OdDelDt," + _
              " OmHCtCd, OdSr,  OmTc as IdExpOdTc, OmYy as IdExpOdYy, OmChr as IdExpOdChr, 0 as IdExpOdNo, 0 as IdExpOdSr, " + _
              " OdDmCd,Space(17) as CdOurDmCd,OdDmSz,OdSfx,OdDmId,DmCtg,DmSalCtg,DmTcTyp,OdDmCol, 0 as OdTrayNo,OdPicNm,OdPoNo,OdKt as IdKt,DmUom, OdVaCtg as IdVaCtg," + _
              " PNum as IdGldLs,space(254) as qRmCtgDesc,space(254) as qDmCtgDesc,space(254) as qDmSalCtgDesc,space(254) as qDmColDesc,space(254) as qVaCtgDesc, " + _
              " OrRmctg, OdDmCol as OdDmCol3,OdGldAs,OdLabAs,OdGldAsWt,OdLabAsWt,OdSalPrc,OdOrdEnt as qOrdQty, OdSalPrc as qSalVal,OdDmPrdInst, " + _
              " OdCmPrdInst,OdCmStmpInst, OdSalRem, OdSzInst,OmSalRem,OmCmDelTerms, OmCmPayTerms, OmPoNo AS qOdOmPoNo, " + _
              " CmLkUpMetLs,DmLsCtg,OdMulby,CmName, OdOmCmCd, DmWaxWt, Space(4) as qTag,OrMainMet,0 as qMainMet,Space(1) as qPtrYN,  OrSrNo, OrRmCtg as RmCtg,OrRmSCtg,Space(14) as OrRmCd ,RmQw,OrLn1,OrLn2,OrRmPtr  , " + _
              " OrPrdQty as OrQty,OrPrdWt as OrWt,OrLmeSal,OrSalRt,OrSetSCd,OrSetSalRt,OrSetSalVal,OrSalVal,Space(10) as RmRSzRt,OrAlySalRt,RmPurityRt,0 as qRmCtgSr, " + _
              " 0 as OlSrNo,OlMCd,OlSCd, PDesc as qLabDesc,RmQW as OlQw,OlQty,OlSalRt,OlSalVal, 0 as OctSrNo, OctCd ,OctSz,OctCtChr, 0 as OctParts,OctRmCd,OctDmCol,OctRmWt,OctDc, " + _
              " OctPrdSeq , OctMainPrtYn, OctDesc Into #TmpA From OrdMst,OrdDsg,OrdRm,OrdLab,DsgMst,CustMst,RmMst,RmRt,OrdCt,Param where 1=2"
         
      moCn.CreateTmpTable "#TmpA", wTblStr
      
      wCommStr = "Select /* row_Number() over(Partition By OmIdNo Order By OmCoCd,OmTc,OmYy,OmChr,OmNo" + wExcelOrdBy + ",OdSr) as qRow,*/ OmCoCd, OmYy, OmTc, OmChr, OmNo, max(OmCmCd) as OmCmCd, max(OmDt) as OmDt, MAX(OmPONo) as OmPONo, max(OmPoDt) as OmPoDt, " + _
                 "max(" + IIF(wb_InBaseCur, " (Case when OmMulDiv= 'M' then OmLmgCst/OmCnvFct else OmLmgCst*OmCnvFct end) ", "OmLmgCst") + ") as OmLmgCst, " + _
                 "max(" + IIF(wb_InBaseCur, " (Case when OmMulDiv= 'M' then OmLmgSal/OmCnvFct else OmLmgSal*OmCnvFct end) ", "OmLmgSal") + ") as OmLmgSal, " + _
                 "max(" + IIF(wb_InBaseCur, " (Case when OmMulDiv= 'M' then OmLmpSal/OmCnvFct else OmLmpSal*OmCnvFct end) ", "OmLmpSal") + ") as OmLmpSal," + _
                 "max(" + IIF(wb_InBaseCur, " (Case when OmMulDiv= 'M' then OmLmsSal/OmCnvFct else OmLmsSal*OmCnvFct end) ", "OmLmsSal") + ") as OmLmsSal, " + _
                 "max(" + IIF(wb_InBaseCur, " (Case when OmMulDiv= 'M' then OmLmlSal/OmCnvFct else OmLmlSal*OmCnvFct end) ", "OmLmlSal") + ") as OmLmlSal," + _
                 wDelDt + ", max(OmHCtCd) as OmHCtCd, " + _
                 "OdSr, " + _
                 ws_BaseRefFlds + " /*'' as IdExpOdTc, '' as IdExpOdYy, '' as IdExpOdChr, 0 as IdExpOdNo, 0 as IdExpOdSr,*/ " + _
                 "max(OdDmCd) as OdDmCd," + wOurDsgCd + ", max(OdDmSz) as OdDmSz, max(OdSfx) as OdSfx, max(OdDmId) as OdDmId, " + _
                 "max(DmCtg) as DmCtg, max(DmSalCtg) as DmSalCtg, max(DmTcTyp) as DmTcTyp, max(OdDmCol) as OdDmCol, 0 as OdTrayNo, Max(OdPicNm) as OdPicNm, Max(OdPoNo) as OdPoNo, " + _
                 "max(OdKt) as IdKt, max(DmUom) as DmUom, max(OdVaCtg) as IdVaCtg, " + _
                 "(Select PNum from Param where PTyp='VACTG' and PMCd=max(OdVaCtg) and PSCd='') as IdGldLs" + _
                 "," + qRmCtgDesc + "," + qDmCtgDesc + "," + qDmSalCtgDesc + "," + qDmColDesc + "," + qVaCtgDesc + "," + _
                 "IsNull((select RmCtg from RmMst join OrdRm on RmCd=OrRmCd " + IIF(gs_Partition = ctCurrPrtn, " and RmPrtKey=OrPrtKey ", "") + " and  " + _
                 "OrMainMet='Y' Where " + IIF(gs_Partition = ctCurrPrtn, " OrPrtKey='" + ctCurrPrtn + "' and ", "") + " OrCocd = Max(OmCocd) And OrTc = OmTc  " + _
                 "and OrYy = OmYy and OrChr=OmChr and OrNo=OmNo and OrSr=OdSr),'') as OrRmctg /*qMainMetRmCtg*/, max(OdDmCol) as OdDmCol3, " + _
                 "Max(OdGldAs) as OdGldAs, Max(OdLabAs) as OdLabAs, max(OdGldAsWt) as OdGldAsWt, max(OdLabAsWt) as OdLabAsWt, max(" + wOdSalPrcOpt + ") as OdSalPrc, "
  '    wCommStr = wCommStr + IIf(adc("UoYN1") = "Y", _
  '               " max(OdOrdQty) as qOrdQty, max(OdOrdQty* OdSalPrc) as qSalVal, ", _
  '               " max(OdOrdEnt) as qOrdQty, max(OdOrdEnt* OdSalPrc) as qSalVal, ")
      
      wCommStr = wCommStr + IIF(ADC("UoYN1") = "Y", _
                 " max(OdOrdQty) as qOrdQty, max(OdOrdQty* " + wOdSalPrcOpt + ") as qSalVal, ", _
                 " max(OdOrdEnt) as qOrdQty, max(OdOrdEnt* " + wOdSalPrcOpt + ") as qSalVal, ")
      
  '    OdDmPrdInst, OdCmPrdInst, OdCmStmpInst, OdSalRem, OdSzInst, OmSalRem, OmCmDelTerms, OmCmPayTerms,
      '****** Manali 3.6.0 - CmLkUpMetLs, DmLsCtg, OdMulBy, CmName added for Excel
      wCommStr = wCommStr + " max(Replace(OdDmPrdInst,char(13)+ char(10),' ')) as OdDmPrdInst, " + _
                            " max(Replace(OdCmPrdInst,char(13)+ char(10),' ')) as OdCmPrdInst, " + _
                            " max(Replace(OdCmStmpInst,char(13)+ char(10),' ')) as OdCmStmpInst, " + _
                            " max(Replace(OdSalRem,char(13)+ char(10),' ')) as OdSalRem, " + _
                            " max(Replace(OdSzInst,char(13)+ char(10),' ')) as OdSzInst, " + _
                            " max(Replace(OmSalRem,char(13)+ char(10),' ')) as OmSalRem, " + _
                            " max(Replace(OmCmDelTerms,char(13)+ char(10),' ')) as OmCmDelTerms, " + _
                            " max(Replace(OmCmPayTerms,char(13)+ char(10),' ')) as OmCmPayTerms, " + _
                            " ' ' AS qOdOmPoNo, max(Case When OmLkUpMetLs <> '' Then OmLkUpMetLs Else CmLkUpMetLs End) as CmLkUpMetLs, max(DmLsCtg) as DmLsCtg, max(OdMulby) as OdMulby, max(CmName) as CmName, max(OdOmCmCd) as OdOmCmCd, max(DmWaxWt) as DmWaxWt, "
          
  '   urmi 2.04-3 RmRt when RmGrp <> Sr, and at RmCd and Lower level
      wRtStr = " case when ((select RmQw from RmMst where " + IIF(gs_Partition = ctCurrPrtn, " RmPrtKey=OrPrtKey and ", "") + " RmCd= OrRmCd) = 'Q' and " + _
                            "sum(" + ws_RmQty + ")<>0) then " + _
                            wRmValOpt + " / sum(" + ws_RmQty + ") " + _
                     "when ((select RmQw from RmMst where " + IIF(gs_Partition = ctCurrPrtn, " RmPrtKey=OrPrtKey and ", "") + " RmCd= OrRmCd)<> 'W' and " + _
                            "sum(" + ws_RmWt + ")<>0) then " + _
                            wRmValOpt + " / sum(" + ws_RmWt + ") " + _
                     "else 0 end "
  '   *********************************
    
      
    'Manoj 2.10.0
    If ADC("UoYN3") = "Y" And ADC("UoYN10") = "Y" Then
      ws_RmPtr = "( Case When Max(OrPrdQty) = 0 Then 0 Else (Case When Max(OrRmCtg) in ('C', 'D') Then Sum(OrPrdWt / (Case When OrPrdQty = 0 Then 1 Else OrPrdQty End)) Else 0 End) End ) as OrRmPtr "
    Else
      ws_RmPtr = " max(OrRmPtr) as OrRmPtr "
    End If
    'Manoj 2.10.0
      
      Select Case UCase(ADC("UoRmGrp"))
      ' ***** Manali 3.6.0 - 04/11/09 - OrAlySalRt added for Excel Report
      Case Is = UCase("N") 'OrRmCtg to get RmVal and RmSetVal, No dets r shown
        ' ***** Manali 3.03 - 28/06/08 - OrLmeSal added
        wRmStr = " 0 as OrSrNo, OrRmCtg as RmCtg, '*' as OrRmSCtg, '*' as OrRmCd, '*' as RmQw, 0 as OrLn1, " + _
                 "0 as OrLn2, 0 as OrRmPtr, sum(" + ws_RmQty + ") as OrQty, sum(" + ws_RmWt + ") as OrWt, 0.00 as OrLmeSal, 0.00 as OrSalRt, '*' as OrSetSCd, " + _
                 "0.00 as OrSetSalRt, " + wRmSetValOpt + ", " + wRmValOpt + " as OrSalVal, '' as RmRSzRt, 0.00 as OrAlySalRt, "
        wRmGrp = " group by OmCoCd, OmYy, OmTc, OmChr, OmNo, OdSr, OrRmCtg "
      Case Is = UCase("Sr")
       '*************************** ZUBIN **************************
        ' 08th Dec 2003, EMR206 - Used var. wOurRmCd
        ' Zubin 213 (BaseCur added)
        ' ***** Manali 3.03 - 28/06/08 - OrLmeSal added
        wRmStr = " OrSrNo, max(OrRmCtg) as RmCtg, max(OrRmSCtg) as OrRmSCtg, " + wOurRmCd + ",max(RmQw) As RmQw, max(OrLn1) as OrLn1, " + _
                 "max(OrLn2) as OrLn2, " + _
                 ws_RmPtr + " , sum(" + ws_RmQty + ") as OrQty, sum(" + ws_RmWt + ") as OrWt, " + _
                 "max(" + IIF(wb_InBaseCur, " (Case when OmMulDiv= 'M' then OrLmeSal/OmCnvFct else OrLmeSal*OmCnvFct end) ", "OrLmeSal") + ") as OrLmeSal, " + _
                 "max(" + IIF(wb_InBaseCur, " (Case when OmMulDiv= 'M' then OrSalRt/OmCnvFct else OrSalRt*OmCnvFct end) ", "OrSalRt") + ") as OrSalRt, max(OrSetSCd) as OrSetSCd, " + _
                 "max(" + IIF(wb_InBaseCur, " (Case when OmMulDiv= 'M' then OrSetSalRt/OmCnvFct else OrSetSalRt*OmCnvFct end) ", "OrSetSalRt") + ") as OrSetSalRt, " + _
                 wRmSetValOpt + ", " + wRmValOpt + " as OrSalVal, '' as RmRSzRt, " + _
                 "max(" + IIF(wb_InBaseCur, " (Case when OmMulDiv= 'M' then OrAlySalRt/OmCnvFct else OrAlySalRt*OmCnvFct end) ", "OrAlySalRt") + ") as OrAlySalRt, "
  '      wRmStr = " OrSrNo, max(OrRmCtg), max(OrRmSCtg), " + wOurRmCd + ", max(OrLn1), max(OrLn2), " + _
                 "max(OrRmPtr), sum(" + ws_RmQty + "), sum(" + ws_RmWt + "), max(OrSalRt), max(OrSetSCd), max(OrSetSalRt), " + _
                 wRmSetValOpt ", " + wRmValOpt + " as OrSalVal, '' as RmRSzRt, "
        wRmGrp = " group by OmCoCd, OmYy, OmTc, OmChr, OmNo, OdSr, OrSrNo "
      Case Is = UCase("RmCdRt")
        ' Zubin 213 (BaseCur added)
        ' ***** Manali 3.03 - 28/06/08 - OrLmeSal added
        wRmStr = " 0 as OrSrNo, OrRmCtg as RmCtg, OrRmSCtg, " + wOurRmCd + ", RmQw, 0 as OrLn1, 0 as OrLn2, 0 as OrRmPtr, " + _
                 "sum(" + ws_RmQty + ") as OrQty, sum(" + ws_RmWt + ") as OrWt, " + _
                 IIF(wb_InBaseCur, " (Case when max(OmMulDiv)= 'M' then OrLmeSal/max(OmCnvFct) else OrLmeSal*max(OmCnvFct) end) ", "OrLmeSal") + " as OrLmeSal, " + _
                 IIF(wb_InBaseCur, " (Case when max(OmMulDiv)= 'M' then OrSalRt/max(OmCnvFct) else OrSalRt*max(OmCnvFct) end) ", "OrSalRt") + " as OrSalRt, '*' as OrSetSCd, 0.00 as OrSetSalRt, " + wRmSetValOpt + ", " + wRmValOpt + " as OrSalVal, '' as RmRSzRt, " + _
                 "max(" + IIF(wb_InBaseCur, " (Case when OmMulDiv= 'M' then OrAlySalRt/OmCnvFct else OrAlySalRt*OmCnvFct end) ", "OrAlySalRt") + ") as OrAlySalRt, "
        ' ***** Manali 3.03 - 28/06/08 - OrLmeSal added
        wRmGrp = " group by OmCoCd, OmYy, OmTc, OmChr, OmNo, OdSr, OrRmCtg, OrRmSCtg, OrRmCd, RmQw, OrLmeSal, OrSalRt "
      Case Is = UCase("RmCd")
  '      wRmStr = " 0 as OrSrNo, OrRmCtg, OrRmSCtg, OrRmCd, 0 as OrLn1, 0 as OrLn2, 0 as OrRmPtr, " + _
  '               "sum(" + ws_RmQty + "), sum(" + ws_RmWt + "), 0 as OrSalRt, '*' as OrSetSCd, 0 as OrSetSalRt, " + wRmSetValOpt + ", " + wRmValOpt + " as OrSalVal, '' as RmRSzRt, "
  '      wRmGrp = " group by OmYy, OmTc, OmChr, OmNo, OdSr, OrRmCtg, OrRmSCtg, OrRmCd "
  
  '     URmi 2.04-3 Change is  wRtStr has been used to calculate the rates
        ' ***** Manali 3.03 - 28/06/08 - OrLmeSal added, OrPrtKey added in Group By
        wRmStr = " 0 as OrSrNo, OrRmCtg as RmCtg, OrRmSCtg, " + wOurRmCd + ", RmQw, 0 as OrLn1, 0 as OrLn2, 0 as OrRmPtr, " + _
                 "sum(" + ws_RmQty + ") as OrQty, sum(" + ws_RmWt + ") as OrWt, 0.00 as OrLmeSal, " + wRtStr + " as OrSalRt, '*' as OrSetSCd, 0.00 as OrSetSalRt, " + wRmSetValOpt + ", " + wRmValOpt + " as OrSalVal, '' as RmRSzRt, 0.00 as OrAlySalRt, "
        wRmGrp = " group by OrPrtKey, OmCoCd, OmYy, OmTc, OmChr, OmNo, OdSr, OrRmCtg, OrRmSCtg, OrRmCd, RmQw "
  
      Case Is = UCase("RmSCtg")
        ' ***** Manali 3.03 - 28/06/08 - OrLmeSal added
        wRmStr = " 0 as OrSrNo, OrRmCtg as RmCtg, OrRmSCtg, '*' as OrRmCd, '*' As RmQw, 0 as OrLn1, 0 as OrLn2, 0 as OrRmPtr, " + _
                 "sum(" + ws_RmQty + ") as OrQty, sum(" + ws_RmWt + ") as OrWt, 0.00 as OrLmeSal, 0.00 as OrSalRt, '' as OrSetSCd, 0.00 as OrSetSalRt, " + wRmSetValOpt + ", " + wRmValOpt + " as OrSalVal, '' as RmRSzRt, 0.00 as OrAlySalRt, "
        wRmGrp = " group by OmCoCd, OmYy, OmTc, OmChr, OmNo, OdSr, OrRmCtg, OrRmSCtg "
      Case Is = UCase("RmCtg")
        ' ***** Manali 3.03 - 28/06/08 - OrLmeSal added
        wRmStr = " 0 as OrSrNo, OrRmCtg as RmCtg, '*' as OrRmSCtg, '*' as OrRmCd, '*' As RmQw, 0 as OrLn1, 0 as OrLn2, 0 as OrRmPtr, " + _
                "sum(" + ws_RmQty + ") as OrQty, sum(" + ws_RmWt + ") as OrWt, 0.00 as OrLmeSal, 0.00 as OrSalRt, '*' as OrSetSCd, 0.00 as OrSetSalRt, " + wRmSetValOpt + ", " + wRmValOpt + " as OrSalVal, '' as RmRSzRt, 0.00 as OrAlySalRt, "
        wRmGrp = " group by OmCoCd, OmYy, OmTc, OmChr, OmNo, OdSr, OrRmCtg "
      Case Is = UCase("RmRSzRt")
  '     Code by Shilp Changed by urmi
  '     wRmStr = " 0 as OrSrNo, max(OrRmCtg), max(OrRmSCtg), max(OrRmCd), max(OrLn1), max(OrLn2), 0 as OrRmPtr, " + _
  '              "sum(" + ws_RmQty + "), sum(" + ws_RmWt + "), max(OrSalRt), max(OrSetSCd) , max(OrSetSalRt), " + wRmSetValOpt + ", " + wRmValOpt + " as OrSalVal, isnull(max(RrRngSz),Str(max(OrLn1), 6,3)) as RmRSzRt, "
  '     wRmGrp = " group by OmYy, OmTc, OmChr, OmNo, OdSr, OrRmCtg, OrRmSCtg, RrRngSz, OrSalRt, OrSetSalRt"
        
  '     urmi 2.04-3
        ' Zubin 213 (BaseCur added)
        ' ***** Manali 3.03 - 28/06/08 - OrLmeSal added
        wRmStr = " 0 as OrSrNo, OrRmCtg as RmCtg, OrRmSCtg, " + wOurRmCd + ",max(RmQw) As RmQw, max(OrLn1) as OrLn1, max(OrLn2) as OrLn2, 0 as OrRmPtr, " + _
                 "sum(" + ws_RmQty + ") as   OrQty, sum(" + ws_RmWt + ") as OrWt, " + _
                 IIF(wb_InBaseCur, " (Case when max(OmMulDiv)= 'M' then OrLmeSal/max(OmCnvFct) else OrLmeSal*max(OmCnvFct) end) ", "OrLmeSal") + " as OrLmeSal, " + _
                 IIF(wb_InBaseCur, " (Case when max(OmMulDiv)= 'M' then OrSalRt/max(OmCnvFct) else OrSalRt*max(OmCnvFct) end) ", "OrSalRt") + " as OrSalRt, " + _
                 "OrSetSCd , " + IIF(wb_InBaseCur, " (Case when max(OmMulDiv)= 'M' then OrSetSalRt/max(OmCnvFct) else OrSetSalRt*max(OmCnvFct) end) ", "OrSetSalRt") + " as OrSetSalRt, " + wRmSetValOpt + ", " + wRmValOpt + " as OrSalVal, isnull(max(RrRngSz),Str(max(OrLn1), 6,3)) as RmRSzRt, " + _
                 "max(" + IIF(wb_InBaseCur, " (Case when OmMulDiv= 'M' then OrAlySalRt/OmCnvFct else OrAlySalRt*OmCnvFct end) ", "OrAlySalRt") + ") as OrAlySalRt, "
                 
  '     urmi 2.04-3 Change is RMCd & OrSetSCd has been mentioned
        ' ***** Manali 3.03 - 28/06/08 - OrLmeSal added
        wRmGrp = " group by OmCoCd, OmYy, OmTc, OmChr, OmNo, OdSr, OrRmCtg, OrRmSCtg, OrRMCd, RrRngSz, OrLmeSal, OrSalRt, OrSetSCd, OrSetSalRt"
  '     ??wRmGrp = " group by OmYy, OmTc, OmChr, OmNo, OdSr, OrRmCtg, OrRmSCtg, OrRMCd, RrRngSz, OrSalRt, (OrSetSCd, OrSetSalRt)?????"
        
      ' ***** Manali 3.6.0 - 25/09/09 - New RmGrp Added - RmSCtgRt
      Case Is = UCase("RmSCtgRt")
        wRmStr = " 0 as OrSrNo, OrRmCtg as RmCtg, OrRmSCtg, '*' as OrRmCd, max(RmQw) As RmQw, 0 as OrLn1, 0 as OrLn2, 0 as OrRmPtr, " + _
                 "sum(" + ws_RmQty + ") as   OrQty, sum(" + ws_RmWt + ") as OrWt, " + _
                 IIF(wb_InBaseCur, " (Case when max(OmMulDiv)= 'M' then OrLmeSal/max(OmCnvFct) else OrLmeSal*max(OmCnvFct) end) ", "OrLmeSal") + " as OrLmeSal, " + _
                 IIF(wb_InBaseCur, " (Case when max(OmMulDiv)= 'M' then OrSalRt/max(OmCnvFct) else OrSalRt*max(OmCnvFct) end) ", "OrSalRt") + " as OrSalRt, " + _
                 "'' as OrSetSCd , 0 as OrSetSalRt, " + wRmSetValOpt + ", " + wRmValOpt + " as OrSalVal, Space(10) as RmRSzRt, " + _
                 "max(" + IIF(wb_InBaseCur, " (Case when OmMulDiv= 'M' then OrAlySalRt/OmCnvFct else OrAlySalRt*OmCnvFct end) ", "OrAlySalRt") + ") as OrAlySalRt, "
        wRmGrp = " group by OmCoCd, OmYy, OmTc, OmChr, OmNo, OdSr, OrRmCtg, OrRmSCtg, RmQW, OrLmeSal, OrSalRt "
      ' ***** Manali 3.6.0 - 25/09/09 - New RmGrp Added - RmSCtgRt
      End Select
   
   '******Geeta *****Emr207** RmsCtg StkByPtr 'N' as qPtrYN is included As Underline is Not reqd for Ptr in Order (Reqd in Invoice )
      '*** Jay 2.13(CT) ***(Ordct flds added)
      '****** Sachin 3.02 - Id fields in Join
      ' Manali 3.6.0 - 04/11/09 - max(RmPurityRt) as RmPurityRt added for Excel Report
      ' ***** Manali Trading Module - CmCtg='C' replaced by gs_CmCtg
      ' ***** Manali 3.6.1 - OrMainMet added
      ' ***** Manali 3.9.0 - For Trading - Base/Ref Vch displayed in report - ws_BaseRefJoin added
      wRmCommStr = "'1RM' as qTag, Max(OrMainMet) as OrMainMet, (case when max(OrMainMet)= 'Y' or max(OrRmCd)= " + _
               "IsNull((Select OrRmCd From OrdRm where " + IIF(gs_Partition = ctCurrPrtn, " OrPrtKey='" + ctCurrPrtn + "' and ", "") + " OrCoCd= max(OmCoCd) and OrTc= OmTc and " + _
               "OrYy= OmYy and OrChr= OmChr and OrNo= OmNo and OrSr= OdSr and OrMainMet= 'Y'), '') " + _
               "then 1 else 2 end) as qMainMet, 'N' as qPtrYN, " + wRmStr + " max(RmPurityRt) as RmPurityRt, " + _
               "Max(Case When RmCtg in ('G', 'P', 'S', 'L') Then 1 When RmCtg In ('D', 'C') Then 2 When RmCtg = 'X' Then 3 Else 4 End) as qRmCtgSr, " + _
               "0 as OlSrNo,'' as OlMCd, '' as OlSCd, " + _
               "IsNull((select rtrim(PDesc) from Param where PTyp= 'LABSCD' and PMCd= 'SET' and " + _
               "PSCd= max(OrSetSCd)),'') as qLabDesc, " + _
               "'' as OlQw, 0 as OlQty, 0 as OlSalRt, 0 as OlSalVal, " + _
               "0 as OctSrNo, '' as OctCd , '' as OctSz, '' as OctCtChr, 0 as OctParts, '' as OctRmCd, '' as OctDmCol, 0 as OctRmWt, " + _
               "'' as OctDc, '' as OctPrdSeq, '' as OctMainPrtYn, '' as OctDesc " + _
               "From OrdMst " + _
               "  join OrdDsg on OdOmIdNo=OmIdNo and OdCoCd= OmCoCd and OdTc= OmTc and OdYy= OmYy " + _
               " and OdChr= OmChr and OdNo= OmNo " + IIF(gs_Partition = ctCurrPrtn, " and OdPrtKey=OmPrtKey ", "") + _
               " join DsgMst on DmIdNo=OdDmIdNo and DmTcTyp='" + wDmTcTyp + "' and DmCd= OdDmCd and DmSz='' " + IIF(gs_Partition = ctCurrPrtn, " and DmPrtKey=OdPrtKey ", "") + _
               " join CustMst on CmCtg='" + gs_CmCtg + "' and CmCd= OmCmCd " + _
               " join OrdRm on OrOdIdNo=OdIdNo and OrCoCd= OdCoCd and OrTc= OdTc and OrYy= OdYy " + _
               " and OrChr= OdChr and OrNo= OdNo and OrSr= OdSr " + IIF(gs_Partition = ctCurrPrtn, " and OrPrtKey=OdPrtKey ", "") + _
               " Join RmMst On RmCd=OrRmCd " + IIF(gs_Partition = ctCurrPrtn, " and RmPrtKey=OrPrtKey ", "") + _
                ws_BaseRefJoin + _
               " Left Outer join RmRt on RrTCTyp= 'RSZ' and RrCmCtg = 'C' and RrCmCd = 'ZSELF' and RrCmCurCd= '' and " + _
               " RrCtg = OrRmCtg and RrSCtg = OrRmSCtg and RrSTWGrp = '' and RrCd = '' and " + _
               " RrDmCtg = '' and RrLsCtg = '' and RrLabMCd = '' and RrFrLn <= OrLn1 and RrToLn >= OrLn1 " + _
               wCnd + IIF(gs_Partition = ctCurrPrtn, " and OmPrtKey='" + ctCurrPrtn + "' ", "") + _
               IIF(wAddlCnd <> "", IIF(wCnd <> "", " And ", " Where ") + wAddlCnd, "") + wCrmCnd + wRmGrp
               '(index= OrdDsg) '(index= OrdRm)
               'and RrFrLn < OrLn2
      '*** Jay 2.13(CT) ***(Ordct flds added)
      Select Case UCase(ADC("UoLabGrp"))
      Case Is = UCase("N")
        wLabStr = "0 as OlSrNo, '*' as OlMCd, '*' as OlSCd, " + _
                  "'*' as qLabDesc, '*' as OlQw , sum(OlQty) as OlQty, 0.00 as OlSalRt, " + wLabValOpt + " "
        wLabGrp = " group by OmCoCd, OmYy, OmTc, OmChr, OmNo, OdSr, OlSrNo "
      Case Is = UCase("Sr")
        wLabStr = "OlSrNo, max(OlMCd) as OlMCd, max(OlSCd) as OlSCd, " + _
                  "(select rtrim(PDesc) from Param where PTyp= 'LABSCD' and PMCd= max(OlMCd) and " + _
                  "PSCd= max(OLSCd)) as qLabDesc, max(OlQw) as OlQw, sum(OlQty) as OlQty, " + wLabRtOpt + ", " + wLabValOpt + " "
        wLabGrp = " group by OmCoCd, OmYy, OmTc, OmChr, OmNo, OdSr, OlSrNo "
      Case Is = UCase("LabMCd")
        wLabStr = "0 as OlSrNo, OlMCd, '*' as OlSCd, " + _
                  "'*' as qLabDesc, '*' as OlQw, sum(OlQty) as OlQty, 0.00 as OlSalRt, " + wLabValOpt + " "
        wLabGrp = " group by OmCoCd, OmYy, OmTc, OmChr, OmNo, OdSr, OlMCd "
      Case Is = UCase("LabSCd")
        wLabStr = "0 as OlSrNo, max(OlMCd) as OlMCd, OlSCd, " + _
                  "(select rtrim(PDesc) from Param where PTyp= 'LABSCD' and PMCd= max(OlMCd) and " + _
                  "PSCd= OLSCd) as qLabDesc, OlQw, sum(OlQty) as OlQty, 0.00 as OlSalRt, " + wLabValOpt + " "
        wLabGrp = " group by OmCoCd, OmYy, OmTc, OmChr, OmNo, OdSr, OlMCd, OlSCd, OlQw "
      End Select
      '*** Jay 2.13(CT) ***(Ordct flds added)
      '****** Sachin 3.02 - Id fields in Join
      ' ***** Manali 3.03 - 28/06/08 - OrLmeSal added
      ' ***** Manali Trading Module - CmCtg='C' replaced by gs_CmCtg
      ' ***** Manali 3.9.0 - For Trading - Base/Ref Vch displayed in report - ws_BaseRefJoin added
      wLabCommStr = "'2LAB' as qTag, '' as OrMainMet, 0 as qMainMet,'' as qPtrYN, 0 as OrSrNo, '' as RmCtg, '' as OrRmSCtg, " + _
                "'' as OrRmCd, '' as RmQw, 0 as OrLn1, 0 as OrLn2, 0 as OrRmPtr, 0 as OrQty, 0 as OrWt, 0 as OrLmeSal, 0 as OrSalRt, " + _
                "'' as OrSetSCd, 0 as OrSetSalRt, 0 as OrSetSalVal, 0 as OrSalVal, '' as RmRSzRt, 0 as OrAlySalRt, 0 as RmPurityRt, 10 as qRmCtgSr, " + wLabStr + _
                ", 0 as OctSrNo, '' as OctCd , '' as OctSz, '' as OctCtChr, 0 as OctParts,  '' as OctRmCd, '' as OctDmCol, 0 as OctRmWt, " + _
                "'' as OctDc, '' as OctPrdSeq, '' as OctMainPrtYn, '' as OctDesc " + _
                "From OrdMst " + _
                "join OrdDsg on OdOmIdNo=OmIdNo and OdCoCd= OmCoCd and OdTc= OmTc and OdYy= OmYy and OdChr= OmChr and OdNo= OmNo " + IIF(gs_Partition = ctCurrPrtn, " and OdPrtKey=OmPrtKey ", "") + _
                "join DsgMst on DmIdNo=OdDmIdNo and DmTcTyp='" + wDmTcTyp + "' and DmCd= OdDmCd and DmSz='' " + IIF(gs_Partition = ctCurrPrtn, " and DmPrtKey=OdPrtKey ", "") + _
                "join CustMst on CmCtg='" + gs_CmCtg + "' and CmCd= OmCmCd " + _
                "join OrdLab on OlOdIdNo=OdIdNo and OlCoCd= OdCoCd and OlTc= OdTc and OlYy= OdYy " + _
                "and OlChr= OdChr and OlNo= OdNo and OlSr= OdSr " + IIF(gs_Partition = ctCurrPrtn, " and OlPrtKey=OdPrtKey ", "") + ws_BaseRefJoin + wCnd + _
                IIF(gs_Partition = ctCurrPrtn, " and OmPrtKey='" + ctCurrPrtn + "' ", "") + _
                IIF(wAddlCnd <> "", IIF(wCnd <> "", " And ", " Where ") + wAddlCnd, "") + wCrmCnd + wLabGrp
                '(index= OrdDsg) '(index= OrdLab)
    '*** Jay 2.13(CT) ***(Ordct flds added)
  
  '*** Jay 2.13(CT) ***
     '****** Sachin 3.02 - Id fields in Joins
     ' ***** Manali Trading Module - CmCtg='C' replaced by gs_CmCtg
     If moCn.RecSeek("Select hCompBagsYn from Head where HCd='" + ctSelfCmCd + "' And " + _
                        " HCoCd='" + ctSelfCoCd + "' and hCompBagsYn= 'Y'") And _
        (ADC("UoYN12")) = "Y" And UCase(ADC.MenuCd) = UCase("RepDetOrd") Then
        wOrdCTCommStr = "'3CT' as qTag, '' as OrMainMet, 0 as qMainMet,'' as qPtrYN, 0 as OrSrNo, '' as RmCtg, '' as OrRmSCtg, " + _
                "'' as OrRmCd, '' as RmQw, 0 as OrLn1, 0 as OrLn2, 0 as OrRmPtr, 0 as OrQty, 0 as OrWt,  " + _
                "0 as OrLmeSal, 0 as OrSalRt, '' as OrSetSCd, 0 as OrSetSalRt, 0.00 as OrSetSalVal, 0.00 as OrSalVal, '' as RmRSzRt, 0.00 as OrAlySalRt, 0.00 as RmPurityRt, 10 as qRmCtgSr, " + _
                "0 as OlSrNo,'' as OlMCd, '' as OlSCd, '' as qLabDesc, '' as OlQw, 0 as OlQty, 0 as OlSalRt, " + _
                "0 as OlSalVal, OctSrNo, Max(OctCd) as OctCd , Max(OctSz) as OctSz, Max(OctCtChr) as OctCtChr, Max(OctParts) as OctParts,  Max(OctRmCd) as OctRmCd, " + _
                "Max(OctDmCol) as OctDmCol, Max(OctRmWt) as OctRmWt, Max(OctDc) as OctDc, " + _
                "Max(OctPrdSeq) as OctPrdSeq, Max(OctMainPrtYn) as OctMainPrtYn, Max(OctDesc) as OctDesc " + _
                "From OrdMst " + _
                "join OrdDsg on OdOmIdNo=OmIdNo and OdCoCd= OmCoCd and OdTc= OmTc and OdYy= OmYy and OdChr= OmChr and OdNo= OmNo " + IIF(gs_Partition = ctCurrPrtn, " and OdPrtKey=OmPrtKey ", "") + _
                "join DsgMst on DmIdNo=OdDmIdNo and DmTcTyp='" + wDmTcTyp + "' and DmCd= OdDmCd and DmSz='' " + IIF(gs_Partition = ctCurrPrtn, " and DmPrtKey=OdPrtKey ", "") + _
                "join CustMst on CmCtg='" + gs_CmCtg + "' and CmCd= OmCmCd " + _
                "join OrdCT on OctOdIdNo=OdIdNo and OctCoCd= OdCoCd and OctTc= OdTc and OctYy= OdYy and OctChr= OdChr and OctNo= OdNo and " + _
                "OctSr= OdSr " + IIF(gs_Partition = ctCurrPrtn, " and OctPrtKey=OdPrtKey ", "") + _
                IIF(UCase(ADC.MenuCd) = UCase("RepDetOrd"), wCnd + IIF(gs_Partition = ctCurrPrtn, " and OmPrtKey='" + ctCurrPrtn + "' ", "") + _
                IIF(wAddlCnd <> "", IIF(wCnd <> "", " And ", " Where ") + wAddlCnd, "") + wCrmCnd, " where 1= 2 ") + _
                " group by OmCoCd, OmYy, OmTc, OmChr, OmNo, OdSr, OctSrNo"
          
        wSqlStrg = wCommStr + wRmCommStr + " Union All " + _
                 wCommStr + wLabCommStr + " Union All " + _
                 wCommStr + wOrdCTCommStr + _
                 " order by OmCoCd, OmYy, OmTc, OmChr, OmNo, OdSr, qTag, qMainMet, OrSrNo, OlSrNo, OctSrNo "
      Else
        If ADC("UoYn23") = "Y" Then
          ws_OrdBy = ws_OrdBy + IIF(ws_OrdBy <> "", ",", "")
          wSqlStrg = wCommStr + wRmCommStr + " Union All " + _
                     wCommStr + wLabCommStr + " order by OmCoCd, OmYy, OmTc, OmChr, OmNo, " + ws_OrdBy + " OdSr, qTag, /*qMainMet,*/ qRmCtgSr, OrSrNo, OlSrNo "
        Else
          wSqlStrg = wCommStr + wRmCommStr + " Union All " + _
                     wCommStr + wLabCommStr + " order by OmCoCd, OmYy, OmTc, OmChr, OmNo, OdSr, qTag, qMainMet, OrSrNo, OlSrNo "

        End If
        mRep.ORDCTHEAD.Suppress = True
        mRep.ORDCTDET.Suppress = True
      End If
  '*** Jay 2.13(CT) ***
      
      moCn.Execute ("Insert Into #TmpA " + wSqlStrg)
      wSqlStrg = "Select * From #TmpA " + IIF(ADC("UoYn23") = "Y", " order by OmCoCd, OmYy, OmTc, OmChr, OmNo, " + ws_OrdBy + " OdSr, qTag, /*qMainMet,*/ qRmCtgSr, OrSrNo, OlSrNo ", "")
      
      If ADC("UoYn23") <> "Y" Then
      wExcelOrdBy = wExcelOrdBy + IIF(wExcelOrdBy <> "", ",", "")
        Set wo_rsPic = moCn.OpenRes("Select Max(DmTcTyp) as DmTcTyp,Max(OdDmCd) as DmCd,Max(OdDmSz) as DmSz,OmCoCd,OmTc,OmYy,OmChr,OmNo,OdSr,Max(OdPicNm) as PicNm from #TmpA" + _
                                    " Group By OmCoCd,OmTc,OmYy,OmChr,OmNo,OdSr " + _
                                    " order by OmCoCd, OmYy, OmTc, OmChr, OmNo, " + wExcelOrdBy + " OdSr ")
        If wo_rsPic.RecCount > 0 Then CmdChgImg.Visible = True
      End If
      
      If ADC("UoYn23") = "Y" Then
        Dim wOdSrCnt As Integer
        Dim wRsOrd As MDORowSet
        ' ***** Manali Trading Module - CmCtg='C' replaced by gs_CmCtg
        wOdSrCnt = moCn.GetFldVal("Select Count(OdSr) from OrdDsg " + _
                  "join OrdMst on OdOmIdNo=OmIdNo and OdCoCd= OmCoCd and OdTc= OmTc and OdYy= OmYy and OdChr= OmChr and OdNo= OmNo " + IIF(gs_Partition = ctCurrPrtn, " and OdPrtKey=OmPrtKey ", "") + _
                  "join CustMst on CmCtg='" + gs_CmCtg + "' and CmCd= OmCmCd " + _
                  "join DsgMst on DmIdNo=OdDmIdNo and DmTcTyp= '" + wDmTcTyp + "' and DmCd= OdDmCd and DmSz= '' " + IIF(gs_Partition = ctCurrPrtn, " and DmPrtKey=OdPrtKey ", "") + _
                  wCnd + IIF(gs_Partition = ctCurrPrtn, " and OdPrtKey='" + ctCurrPrtn + "' ", "") + wCrmCnd + " ")
        If wOdSrCnt = 0 Then
          DispMsg "No Data Found for given scope", etError: Exit Sub
        ElseIf wOdSrCnt > 285 Then
          DispMsg "No of Designs is greater than 285, Please Reduce the scope", etError: Exit Sub
        Else
          Set wRsOrd = moCn.OpenRes(wSqlStrg + ctMaxDopOpt)
          wErrMsg = ""
          Call SetExcel(wRsOrd)
          ADC.RepRecSource = "Select * from Param Where 1=2"
          
          If wErrMsg <> "" Then
            DispMsg wErrMsg, etError
          End If
          Set wRsOrd = Nothing
        End If
      Else
        ADC.RepRecSource = wSqlStrg + ctMaxDopOpt  '****** Sachin 3.01 added ctMaxDopOpt
      End If
      
  '    wDsgCtgStr = "Select OdYy, OdTc, OdChr, OdNo, OdSr, DmCtg, (Select PDesc from Param where " + _
  '           "PTyp= 'DMCTG' and PMCd= DmCtg) as qDesc, " + _
  '           IIf(adc("UoYN1") = "Y", " OdOrdQty as qOrdQty, OdOrdQty* " + wOdSalPrcOpt + " as qSalVal ", " OdOrdEnt as qSalVal, OdOrdEnt* " + wOdSalPrcOpt + " as qSalVal ") + _
  '           " From OrdDsg join DsgMst on DmTcTyp= 'DM' and DmCd= OdDmCd and " + _
  '           "DmSz= isnull((Select DmSz from DsgMst where DmTcTyp='DM' " + _
  '           "and DmCd= OdDmCd and DmSz= OdDmSz), '') " + _
  '           "join OrdMst on OdCoCd= OmCoCd and OdTc= OmTc and OdYy= OmYy and OdChr= OmChr and OdNo= OmNo " + _
  '           "join CustMst on CmCtg='C' and CmCd= OmCmCd " + wCnd
             '(index= DsgMst)
    
    'OdYy, OdTc, OdChr, OdNo, OdSr, DmCtg, qDesc, qOrdQty, qSalVal
      '****** Sachin 3.02 - Id fields in Joins
      ' ***** Manali Trading Module - CmCtg='C' replaced by gs_CmCtg
      wDsgCtgStr = "Select OdYy, OdTc, OdChr, OdNo, OdSr, DmCtg, (Select PDesc from Param where " + _
             "PTyp= 'DMCTG' and PMCd= DmCtg) as qDesc, " + _
             IIF(ADC("UoYN1") = "Y", " OdOrdQty as qOrdQty, OdOrdQty* " + wOdSalPrcOpt + " as qSalVal ", " OdOrdEnt as qOrdQty, OdOrdEnt* " + wOdSalPrcOpt + " as qSalVal ") + _
             " From OrdDsg join DsgMst On DmTcTyp='" + wDmTcTyp + "' and DmCd= OdDmCd and " + _
             "DmSz= isnull((Select DmSz from DsgMst where " + IIF(gs_Partition = ctCurrPrtn, " DmPrtKey=OdPrtKey and ", "") + " DmIdNo=OdDmIdNo and DmTcTyp='DM' " + _
             "and DmCd= OdDmCd and DmSz= OdDmSz), '') " + _
             "join OrdMst on OdOmIdNo=OmIdNo and OdCoCd= OmCoCd and OdTc= OmTc and OdYy= OmYy and OdChr= OmChr and OdNo= OmNo " + IIF(gs_Partition = ctCurrPrtn, " and OdPrtKey=OmPrtKey ", "") + _
             "join CustMst on CmCtg='" + gs_CmCtg + "' and CmCd= OmCmCd " + wCnd + IIF(gs_Partition = ctCurrPrtn, " and OdPrtKey='" + ctCurrPrtn + "' ", "") + _
             IIF(wAddlCnd <> "", IIF(wCnd <> "", " And ", " Where ") + wAddlCnd, "") + wCrmCnd
  
      Call moCn.RepRes(mRep.Subreport1.OpenSubreport, wDsgCtgStr + ctMaxDopOpt)  '****** Sachin 3.01 added ctMaxDopOpt
  
  
  
  
  
      Dim wOdQty As String
      
      'urmi Error solved in 2.05 : instead of sum Of (*) it was * of sum
      'wOdQty = IIf(ADC("UoYN1") = "Y", "max(OdOrdQty)", "max(OdOrdEnt)")
      wOdQty = IIF(ADC("UoYN1") = "Y", "OdOrdQty", "OdOrdEnt")
    
      If (ADC("UoYN2")) = "Y" Then
          ' Zubin 213 (BaseCur added)
          wRmValOpt = IIF(wb_InBaseCur, " (Case when OmMulDiv= 'M' then OrSalVal/OmCnvFct else OrSalVal*OmCnvFct end) ", "OrSalVal")
          ' Zubin 212 (02/03/06)
          GRP_REP.SetFormula mRep, "wActSalRt", "{rdo.OrSalRt}"
      ElseIf (ADC("UoYN2")) = "N" Then
          wRmValOpt = "0"
          ' Zubin 212 (02/03/06)
          GRP_REP.SetFormula mRep, "wActSalRt", 0
      End If
      
  ' **** Zubin 211 **** '
  '''''    wRmDetSql = "Select OdTc, OdYy, OdChr, OdNo, OrRmCtg, sum( " + wOdQty + " * " + ws_RmQty + ") as qRmQty, sum( " + wOdQty + _
  '''''    " * " + ws_RmWt + ") as qRmWt, sum( " + _
  '''''    wOdQty + " * " + wRmValOpt + ") as qRmVal from OrdDsg " + _
  '''''    "join OrdMst on OdCoCd= OmCoCd and OdTc= OmTc and OdYy= OmYy and OdChr= OmChr and OdNo= OmNo " + _
  '''''    "join CustMst on CmCtg='C' and CmCd= OmCmCd " + _
  '''''    "join DsgMst on DmTcTyp= '" + wDmTcTyp + "' and DmCd= OdDmCd and DmSz= '' " + _
  '''''    "join OrdRm on OrCoCd= OdCoCd and OrTc= OdTc and OrYy= OdYy and OrChr= OdChr and " + _
  '''''    "OrNo= OdNo And OrSr= OdSr" + wCnd + " group by OdCoCd, OdTc, OdYy, OdChr, OdNo, OrRmCtg"
  
      '****** Sachin 3.02 - Id fields in Joins
      ' ***** Manali Trading Module - CmCtg='C' replaced by gs_CmCtg
      wRmDetSql = "Select OdTc, OdYy, OdChr, OdNo, OrRmCtg, " + _
                  "sum( " + wOdQty + " * " + ws_RmQty + ") as qRmQty, " + _
                  "sum(" + wOdQty + " * OrWsQty ) as qWSetQty, sum(" + wOdQty + " * OrHsQty ) as qHSetQty, " + _
                  "sum( " + wOdQty + " * " + ws_RmWt + ") as qRmWt, sum( " + wOdQty + " * " + wRmValOpt + ") as qRmVal " + _
                  "from OrdDsg " + _
                  "join OrdMst on OdOmIdNo=OmIdNo and OdCoCd= OmCoCd and OdTc= OmTc and OdYy= OmYy and OdChr= OmChr and OdNo= OmNo " + IIF(gs_Partition = ctCurrPrtn, " and OdPrtKey=OmPrtKey ", "") + _
                  "join CustMst on CmCtg='" + gs_CmCtg + "' and CmCd= OmCmCd " + _
                  "join DsgMst on DmIdNo=OdDmIdNo and DmTcTyp= '" + wDmTcTyp + "' and DmCd= OdDmCd and DmSz= '' " + IIF(gs_Partition = ctCurrPrtn, " and DmPrtKey=OdPrtKey ", "") + _
                  "join OrdRm on OrOdIdNo=OdIdNo and OrCoCd= OdCoCd and OrTc= OdTc and OrYy= OdYy and OrChr= OdChr and " + _
                  "OrNo= OdNo And OrSr= OdSr " + IIF(gs_Partition = ctCurrPrtn, " and OrPrtKey=OdPrtKey ", "") + wCnd + _
                  IIF(gs_Partition = ctCurrPrtn, " and OdPrtKey='" + ctCurrPrtn + "' ", "") + _
                   wCrmCnd + " group by OdCoCd, OdTc, OdYy, OdChr, OdNo, OrRmCtg "
  ' **** Zubin 211 **** '
  
      Call moCn.RepRes(mRep.Subreport2.OpenSubreport, wRmDetSql + ctMaxDopOpt)  '****** Sachin 3.01 added ctMaxDopOpt
  
      
      If ADC("UoYN") = "N" Then
        mRep.PICSEC.Suppress = True
      End If
    '*** Jay 3.3(FG)
    ' ***** Manali Trading Module
    ElseIf UCase(ADC.MenuCd) = UCase("RepOrdCat") Or UCase(ADC.MenuCd) = UCase("RepFgCat") Or UCase(ADC.MenuCd) = UCase("RepJtCat") Then
    'ElseIf UCase(ADC.MenuCd) = UCase("RepOrdCat") Then
    '*** Jay 3.3(FG)
  '    wSqlStrg = Select OmYy, OmTc, OmChr, OmNo, OmCmCd, OmDt, OdSr, OdDmCd, OdDmSz, OdSfx, " + _
  '               "DmCtg, DmTcTyp, OdDmCol, OdKt, OdGldAsWt, OdLabAsWt, OdSalPrc, OdOrdEnt as qOrdQty, OrMainMet, " + _
  '               "OrRmCtg, OrRmSCtg, OrRmCd, OrQty, OrWt, (Select PDesc from Param " + _
  '               "where PTyp= 'DMCOL' and PMCd= OdDmCol) as qColDesc " + _
  '               "From OrdMst join OrdDsg (index= OrdDsg) on OdTc= OmTc and OdYy= OmYy " + _
  '               "and OdChr= OmChr and OdNo= OmNo " + _
  '               "join DsgMst on DmTcTyp= 'DM' and DmCd= OdDmCd and DmSz= '' " + _
  '               "join OrdRm (index= OrdRm) on OrTc= OdTc and OrYy= OdYy and OrChr= OdChr and " + _
  '               "OrNo= OdNo and OrSr= OdSr " + wCnd
  '    adc.RepRecSource = wSqlStrg
  '**********Previous-Emr205**********************************************
    '  wSqlStrg = "Select OmYy, OmTc, OmChr, OmNo, OmCmCd, OmDt, OmPoNo, OdSr, OdDmCd, OdDmSz, OdSfx, " + _
                 "DmCtg, DmTcTyp, OdDmCol, OdKt, OdGldAsWt, OdSalPrc, OdOrdEnt as qOrdQty, " + _
                 "(Select PDesc from Param where PTyp= 'DMCOL' and PMCd= OdDmCol) as qColDesc, " + _
                 "(Select sum(" + ws_RmWt + "/ (case when OrRmCtg in ('C', 'D') then 5 else 1 end)) " + _
                 "from OrdRm where OrCoCd= OdCoCd and OrTc= OdTc and OrYy= OdYy and " + _
                 "OrChr= OdChr and OrNo= OdNo and OrSr= OdSr) as qGrsWt " + _
                 "From OrdMst join OrdDsg on OdCoCd= OmCoCd and OdTc= OmTc and OdYy= OmYy " + _
                 "and OdChr= OmChr and OdNo= OmNo " + _
                 "join CustMst on CmCtg='C' and CmCd= OmCmCd " + _
                 "join DsgMst on DmTcTyp= '" + wDmTcTyp + "' and DmCd= OdDmCd and DmSz= '' " + wCnd
                 '(index= OrdDsg)
    
    '*******************************Geeta**********************************
    ' 23rd Dec added Group Combination fields
    'OdVaCtg,DmUom,IdVaCtg,qRmCtgDesc, qDmCtgDesc,qDmColDesc , qVaCtgDesc,qMainMetRmCtg
    ' Zubin 212 (DmSalCtg and qDmSalCtgDesc added)
      ' Zubin 213 (Basecur added)
      '****** Sachin 3.02 - Id fields In Joins
      '*** Jay 3.2.0 [OdPicNm] '*** Jay 3.2.0 [OdPoNo]
      '*** Jay 3.3(FG)[OdDmId]
      ' ***** Manali Trading Module - CmCtg='C' replaced by gs_CmCtg
      ' ***** Manali 3.8.0 - OdTrayNo added
'      wSqlStrg = "Select OmCoCd, OmYy, OmTc, OmChr, OmNo, OmCmCd, OmDt, OmPoNo, OdSr," + wOurDsgCd + ", OdDmCd, OdDmSz, OdSfx, OdDmId, " + _
'                 "DmCtg, DmSalCtg, DmTcTyp, OdDmCol, OdPicNm, OdPoNo, OdKt as IdKt, OdVaCtg as IdVaCtg, DmUom," + _
'                 "(Select PNum from Param where PTyp='VACTG' and PMCd=OdVaCtg and PSCd='') as IdGldLs" + _
'                 "," + qRmCtgDesc + "," + qDmCtgDesc + "," + qDmSalCtgDesc + "," + qDmColDesc + "," + qVaCtgDesc + "," + _
'                 "(select max(RmCtg) from RmMst ,OrdRm where " + IIF(gs_Partition = ctCurrPrtn, " RmPrtKey=OrPrtKey and ", "") + " RmCd=OrRmCd  and  " + IIF(gs_Partition = ctCurrPrtn, " OrPrtKey=OmPrtKey and ", "") + _
'                 "OrMainMet='Y' and OrCocd = OmCocd And OrTc = OmTc  " + _
'                 "and OrYy = OmYy and OrChr=OmChr and OrNo=OmNo and OrSr=OdSr) as OrRmctg /*qMainMetRmCtg*/," + _
'                 "OdGldAsWt, OdLabAsWt, " + _
'                 IIF(wb_InBaseCur, " (Case when OmMulDiv= 'M' then OdSalPrc/OmCnvFct else OdSalPrc*OmCnvFct end) ", "OdSalPrc") + " as OdSalPrc, " + _
'                 "OdOrdEnt as qOrdQty, (Select PDesc from Param where PTyp= 'DMCOL' and PMCd= OdDmCol) as qColDesc, " + _
'                 "(Select sum(" + ws_RmWt + "/ (case when OrRmCtg in ('C', 'D') then 5 else 1 end)) " + _
'                 "from OrdRm where OrOdIdNo=OdIdNo and OrCoCd= OdCoCd and OrTc= OdTc and OrYy= OdYy and " + _
'                 " OrChr= OdChr and OrNo= OdNo and OrSr= OdSr " + IIF(gs_Partition = ctCurrPrtn, " and OrPrtKey=OdPrtKey ", "") + " ) as qGrsWt, 0 as OdTrayNo " + _
'                 " From OrdMst join OrdDsg on OdOmIdNo=OmIdNo and OdCoCd= OmCoCd and OdTc= OmTc and OdYy= OmYy " + _
'                 " and OdChr= OmChr and OdNo= OmNo " + IIF(gs_Partition = ctCurrPrtn, " and OdPrtKey=OmPrtKey ", "") + _
'                 " join CustMst on CmCtg='" + gs_CmCtg + "' and CmCd= OmCmCd " + _
'                 " join DsgMst on DmIdNo=OdDmIdNo and DmTcTyp= '" + wDmTcTyp + "' and DmCd= OdDmCd and DmSz= ''  " + IIF(gs_Partition = ctCurrPrtn, " and DmPrtKey=OdPrtKey ", "") + _
'                 wCnd + IIF(gs_Partition = ctCurrPrtn, " and OmPrtKey='" + ctCurrPrtn + "' ", "") + wCrmCnd
                 '(index= OrdDsg)
                 
                 
      wSqlStrg = "Select 0 as qRow, 0 as qCol,OmCoCd, OmYy, OmTc, OmChr, OmNo, OmCmCd,OdOmCmCd,OmDt, OmPoNo, OdSr ,Space(17) as CdOurDmCd, OdDmCd, OdDmSz, OdSfx, OdDmId, DmCtg, DmSalCtg," + _
               "DmTcTyp, OdDmCol, OdPicNm, OdPoNo, OdKt as IdKt, OdVaCtg as IdVaCtg, DmUom ,PNum as IdGldLs, " + _
               "Space(254) as qRmCtgDesc,Space(254) as  qDmCtgDesc,Space(254) as  qDmSalCtgDesc,Space(254) as  qDmColDesc,Space(254) as  qVaCtgDesc, " + _
               "OrRmCtg  as OrRmctg ,OdGldAsWt, OdLabAsWt,  OdSalPrc  as OdSalPrc, OdOrdEnt as qOrdQty, Space(254) as qColDesc, " + _
               "OrPrdWt as qGrsWt, 0 as OdTrayNo   Into #TmpA " + _
               " From OrdMst,OrdDsg,Param,CustMst,DsgMst,OrdRm  where 1=2 "
      moCn.CreateTmpTable "#TmpA", wSqlStrg
      
      wExcelOrdBy = IIF(wExcelOrdBy <> "", ",", "") + wExcelOrdBy
      
      wSqlStrg = "Insert Into #TmpA " + _
                 "Select (row_Number() over(Partition By OmIdNo Order By OmCoCd,OmTc,OmYy,OmChr,OmNo" + wExcelOrdBy + ",OdSr)-1)/3 as qRow," + _
                 " row_Number() over(Partition By OmIdNo Order By OmCoCd,OmTc,OmYy,OmChr,OmNo" + wExcelOrdBy + ",OdSr)%3 as qCol,OmCoCd, OmYy, OmTc, OmChr, OmNo, OmCmCd, OdOmCmCd, OmDt, OmPoNo, OdSr ," + wOurDsgCd + ", OdDmCd, OdDmSz, OdSfx, OdDmId, " + _
                 "DmCtg, DmSalCtg, DmTcTyp, OdDmCol, OdPicNm, OdPoNo, OdKt as IdKt, OdVaCtg as IdVaCtg, DmUom ," + _
                 "PNum as IdGldLs" + _
                 "," + qRmCtgDesc + "," + qDmCtgDesc + "," + qDmSalCtgDesc + "," + qDmColDesc + "," + qVaCtgDesc + "," + _
                 "IsNull((select max(RmCtg) from RmMst ,OrdRm where " + IIF(gs_Partition = ctCurrPrtn, " RmPrtKey=OrPrtKey and ", "") + " RmCd=OrRmCd  and  " + IIF(gs_Partition = ctCurrPrtn, " OrPrtKey=OmPrtKey and ", "") + _
                 "OrMainMet='Y' and OrCocd = OmCocd And OrTc = OmTc  " + _
                 "and OrYy = OmYy and OrChr=OmChr and OrNo=OmNo and OrSr=OdSr),'') as OrRmctg /*qMainMetRmCtg*/," + _
                 "OdGldAsWt, OdLabAsWt, " + _
                 IIF(wb_InBaseCur, " (Case when OmMulDiv= 'M' then OdSalPrc/OmCnvFct else OdSalPrc*OmCnvFct end) ", "OdSalPrc") + " as OdSalPrc, " + _
                 "OdOrdEnt as qOrdQty, (Select PDesc from Param where PTyp= 'DMCOL' and PMCd= OdDmCol) as qColDesc, " + _
                 "(Select sum(" + ws_RmWt + "/ (case when OrRmCtg in ('C', 'D') then 5 else 1 end)) " + _
                 "from OrdRm where OrOdIdNo=OdIdNo and OrCoCd= OdCoCd and OrTc= OdTc and OrYy= OdYy and " + _
                 " OrChr= OdChr and OrNo= OdNo and OrSr= OdSr " + IIF(gs_Partition = ctCurrPrtn, " and OrPrtKey=OdPrtKey ", "") + " ) as qGrsWt, 0 as OdTrayNo  " + _
                 " From OrdMst join OrdDsg on OdOmIdNo=OmIdNo and OdCoCd= OmCoCd and OdTc= OmTc and OdYy= OmYy " + _
                 " and OdChr= OmChr and OdNo= OmNo " + IIF(gs_Partition = ctCurrPrtn, " and OdPrtKey=OmPrtKey ", "") + _
                 " Join Param a On a.PTyp='VACTG' and a.PMCd=OdVaCtg and a.PSCd=''" + _
                 " join CustMst on CmCtg='" + gs_CmCtg + "' and CmCd= OmCmCd " + _
                 " join DsgMst on DmIdNo=OdDmIdNo and DmTcTyp= '" + wDmTcTyp + "' and DmCd= OdDmCd and DmSz= ''  " + IIF(gs_Partition = ctCurrPrtn, " and DmPrtKey=OdPrtKey ", "") + _
                 wCnd + IIF(gs_Partition = ctCurrPrtn, " and OmPrtKey='" + ctCurrPrtn + "' ", "") + wCrmCnd
                 '(index= OrdDsg)
    moCn.Execute (wSqlStrg)
    
    wSqlStrg = " Select  OmCoCd,OmYy,OmTc,OmChr,OmNo,qRow,Max(OmCmCd) as OmCmCd ,Max(OmDt) as OmDt,Max(OdOmCmCd) as OdOmCmCd,Max(OmPoNo) as OmPoNo, " + _
               " Max((Case when qCol=1 Then   OdSr else 0 End)) as OdSr1,Max((Case when qCol=1 Then   CdOurDmCd else '' End)) as CdOurDmCd1, " + _
               " Max((Case when qCol=1 Then   OdDmCd else '' End)) as OdDmCd1,Max((Case when qCol=1 Then   OdDmSz else '' End)) as OdDmSz1, " + _
               " Max((Case when qCol=1 Then   OdSfx else '' End)) as OdSfx1,Max((Case when qCol=1 Then   OdDmId else '' End)) as OdDmId1, " + _
               " Max((Case when qCol=1 Then   DmCtg else '' End)) as DmCtg1,Max((Case when qCol=1 Then   DmSalCtg else '' End)) as DmSalCtg1, " + _
               " Max((Case when qCol=1 Then   DmTcTyp else '' End)) as DmTcTyp1,Max((Case when qCol=1 Then   OdDmCol else '' End)) as OdDmCol1, " + _
               " Max((Case when qCol=1 Then   OdPicNm else '' End)) as OdPicNm1,Max((Case when qCol=1 Then   OdPoNo else '' End)) as OdPoNo1, " + _
               " Max((Case when qCol=1 Then   IdKt else '' End)) as IdKt1,Max((Case when qCol=1 Then   IdVaCtg else '' End)) as IdVaCtg1, " + _
               " Max((Case when qCol=1 Then   DmUom else '' End)) as DmUom1,Max((Case when qCol=1 Then   IdGldLs else '' End)) as IdVaCtg1, " + _
               " Max((Case when qCol=1 Then   qRmCtgDesc else '' End)) as qRmCtgDesc1,Max((Case when qCol=1 Then   qDmCtgDesc else '' End)) as qDmCtgDesc1, " + _
               " Max((Case when qCol=1 Then   qDmSalCtgDesc else '' End)) as qDmSalCtgDesc1,Max((Case when qCol=1 Then   qDmColDesc else '' End)) as qDmColDesc1," + _
               " Max((Case when qCol=1 Then   qVaCtgDesc else '' End)) as qVaCtgDesc1,Max((Case when qCol=1 Then   OrRmctg else '' End)) as OrRmctg1, " + _
               " Max((Case when qCol=1 Then   OdGldAsWt else 0 End)) as OdGldAsWt1,Max((Case when qCol=1 Then   OdLabAsWt else 0 End)) as OdLabAsWt1, " + _
               " Max((Case when qCol=1 Then   OdSalPrc else 0 End)) as OdSalPrc1,Max((Case when qCol=1 Then   OdLabAsWt else 0 End)) as OdLabAsWt1, " + _
               " Max((Case when qCol=1 Then   qOrdQty else 0 End)) as qOrdQty1,Max((Case when qCol=1 Then   qColDesc else '' End)) as qColDesc1, " + _
               " Max((Case when qCol=1 Then   qGrsWt else 0 End)) as qGrsWt1,Max((Case when qCol=1 Then   OdTrayNo else '' End)) as OdTrayNo1, "
      wSqlStrg = wSqlStrg + "Max((Case when qCol=2 Then   OdSr else 0 End)) as OdSr2,Max((Case when qCol=2 Then   CdOurDmCd else '' End)) as CdOurDmCd2, " + _
                        " Max((Case when qCol=2 Then   OdDmCd else '' End)) as OdDmCd2,Max((Case when qCol=2 Then   OdDmSz else '' End)) as OdDmSz2," + _
                        " Max((Case when qCol=2 Then   OdSfx else '' End)) as OdSfx2,Max((Case when qCol=2 Then   OdDmId else '' End)) as OdDmId2, " + _
                        " Max((Case when qCol=2 Then   DmCtg else '' End)) as DmCtg2,Max((Case when qCol=2 Then   DmSalCtg else '' End)) as DmSalCtg2, " + _
                        " Max((Case when qCol=2 Then   DmTcTyp else '' End)) as DmTcTyp2,Max((Case when qCol=2 Then   OdDmCol else '' End)) as OdDmCol2, " + _
                        " Max((Case when qCol=2 Then   OdPicNm else '' End)) as OdPicNm2,Max((Case when qCol=2 Then   OdPoNo else '' End)) as OdPoNo2, " + _
                        " Max((Case when qCol=2 Then   IdKt else '' End)) as IdKt2,Max((Case when qCol=2 Then   IdVaCtg else '' End)) as IdVaCtg2, " + _
                        " Max((Case when qCol=2 Then   DmUom else '' End)) as DmUom2,Max((Case when qCol=2 Then   IdGldLs else '' End)) as IdVaCtg2, " + _
                        " Max((Case when qCol=2 Then   qRmCtgDesc else '' End)) as qRmCtgDesc2,Max((Case when qCol=2 Then   qDmCtgDesc else '' End)) as qDmCtgDesc2, " + _
                        " Max((Case when qCol=2 Then   qDmSalCtgDesc else '' End)) as qDmSalCtgDesc2,Max((Case when qCol=2 Then   qDmColDesc else '' End)) as qDmColDesc2, " + _
                        " Max((Case when qCol=2 Then   qVaCtgDesc else '' End)) as qVaCtgDesc2,Max((Case when qCol=2 Then   OrRmctg else '' End)) as OrRmctg2, " + _
                        " Max((Case when qCol=2 Then   OdGldAsWt else 0 End)) as OdGldAsWt2,Max((Case when qCol=2 Then   OdLabAsWt else 0 End)) as OdLabAsWt2, " + _
                        " Max((Case when qCol=2 Then   OdSalPrc else 0 End)) as OdSalPrc2,Max((Case when qCol=2 Then   OdLabAsWt else 0 End)) as OdLabAsWt2," + _
                        " Max((Case when qCol=2 Then   qOrdQty else 0 End)) as qOrdQty2,Max((Case when qCol=2 Then   qColDesc else '' End)) as qColDesc2, " + _
                        " Max((Case when qCol=2 Then   qGrsWt else 0 End)) as qGrsWt2,Max((Case when qCol=2 Then   OdTrayNo else '' End)) as OdTrayNo2, "
      wSqlStrg = wSqlStrg + " Max((Case when qCol=0 Then   OdSr else 0 End)) as OdSr3,Max((Case when qCol=0 Then   CdOurDmCd else '' End)) as CdOurDmCd3, " + _
                            " Max((Case when qCol=0 Then   OdDmCd else '' End)) as OdDmCd3,Max((Case when qCol=0 Then   OdDmSz else '' End)) as OdDmSz3, " + _
                            " Max((Case when qCol=0 Then   OdSfx else '' End)) as OdSfx3,Max((Case when qCol=0 Then   OdDmId else '' End)) as OdDmId3, " + _
                            " Max((Case when qCol=0 Then   DmCtg else '' End)) as DmCtg3,Max((Case when qCol=0 Then   DmSalCtg else '' End)) as DmSalCtg3, " + _
                            " Max((Case when qCol=0 Then   DmTcTyp else '' End)) as DmTcTyp3,Max((Case when qCol=0 Then   OdDmCol else '' End)) as OdDmCol3, " + _
                            " Max((Case when qCol=0 Then   OdPicNm else '' End)) as OdPicNm3,Max((Case when qCol=0 Then   OdPoNo else '' End)) as OdPoNo3, " + _
                            " Max((Case when qCol=0 Then   IdKt else '' End)) as IdKt3,Max((Case when qCol=0 Then   IdVaCtg else '' End)) as IdVaCtg3, " + _
                            " Max((Case when qCol=0 Then   DmUom else '' End)) as DmUom3,Max((Case when qCol=0 Then   IdGldLs else '' End)) as IdVaCtg3, " + _
                            " Max((Case when qCol=0 Then   qRmCtgDesc else '' End)) as qRmCtgDesc3,Max((Case when qCol=0 Then   qDmCtgDesc else '' End)) as qDmCtgDesc3, " + _
                            " Max((Case when qCol=0 Then   qDmSalCtgDesc else '' End)) as qDmSalCtgDesc3,Max((Case when qCol=0 Then   qDmColDesc else '' End)) as qDmColDesc3, " + _
                            " Max((Case when qCol=0 Then   qVaCtgDesc else '' End)) as qVaCtgDesc3,Max((Case when qCol=0 Then   OrRmctg else '' End)) as OrRmctg3, " + _
                            " Max((Case when qCol=0 Then   OdGldAsWt else 0 End)) as OdGldAsWt3,Max((Case when qCol=0 Then   OdLabAsWt else 0 End)) as OdLabAsWt3, " + _
                            " Max((Case when qCol=0 Then   OdSalPrc else 0 End)) as OdSalPrc3,Max((Case when qCol=0 Then   OdLabAsWt else 0 End)) as OdLabAsWt3, " + _
                            " Max((Case when qCol=0 Then   qOrdQty else 0 End)) as qOrdQty3,Max((Case when qCol=0 Then   qColDesc else '' End)) as qColDesc3, " + _
                            " Max((Case when qCol=0 Then   qGrsWt else 0 End)) as qGrsWt3,Max((Case when qCol=0 Then   OdTrayNo else '' End)) as OdTrayNo3 " + _
                            "  From #TmpA Group By OmCoCd,OmYy,OmTc,OmChr,OmNo,qRow Order By OmCoCd,OmYy,OmTc,OmChr,OmNo,qRow "
     
      '****** Sachin 3.02 - Id fields In Joins
      ADC.RepRecSource = wSqlStrg + ctMaxDopOpt  '****** Sachin 3.01 added ctMaxDopOpt
      
      ' ***** Manali 3.03 -17/06/08 - Pd Mod - 'L' added
      ' ***** Manali Trading Module - CmCtg='C' replaced by gs_CmCtg
      wRmDet = "Select OdYy, OdTc, OdChr, OdNo, OdSr, OrMainMet, " + _
               "OrRmCtg, (case OrRmCtg when 'G' then 1 when 'P' then 2 when 'S' then 8 when 'L' then 9 when 'D' then 3 " + _
               "when 'C' then 4 when 'X' then 5 when 'A' then 6 when 'M' then 7 else 0 end) " + _
               "as qTag, OrRmSCtg, " + IIF(ADC("UOYN30") = "Y", "(Case When RmIntQly='' Then OrRmCd Else RmIntQly End)", "OrRmCd") + " as OrRmCd, " + ws_RmQty + " as OrQty, " + ws_RmWt + " as OrWt " + _
               "From OrdDsg " + _
               " join OrdMst on OdOmIdNo=OmIdNo and OdCoCd= OmCoCd and OdTc= OmTc and OdYy= OmYy and OdChr= OmChr and OdNo= OmNo " + IIF(gs_Partition = ctCurrPrtn, " and OdPrtKey=OmPrtKey ", "") + _
               " join OrdRm on OrOdIdNo=OdIdNo and OrCoCd= OdCoCd and OrTc= OdTc and OrYy= OdYy and OrChr= OdChr and " + _
               " OrNo= OdNo and OrSr= OdSr " + IIF(gs_Partition = ctCurrPrtn, " and OrPrtKey=OdPrtKey ", "") + _
               " Join RmMst On RmCd=OrRmCd " + IIF(gs_Partition = ctCurrPrtn, " And RmPrtKey=OrPrtKey ", "") + _
               " join CustMst on CmCtg='" + gs_CmCtg + "' and CmCd= OmCmCd " + _
               " join DsgMst on DmIdNo=OdDmIdNo and DmTcTyp= '" + wDmTcTyp + "' and DmCd= OdDmCd and DmSz= '' " + IIF(gs_Partition = ctCurrPrtn, " and OdPrtKey=DmPrtKey ", "") + _
               wCnd + IIF(gs_Partition = ctCurrPrtn, " and OdPrtKey='" + ctCurrPrtn + "' ", "") + wCrmCnd
               '(index= OrdRm)
        
      Call moCn.RepRes(mRep.Subreport1.OpenSubreport, wRmDet + ctMaxDopOpt)  '****** Sachin 3.01 added ctMaxDopOpt
      Call moCn.RepRes(mRep.Subreport2.OpenSubreport, wRmDet + ctMaxDopOpt)  '****** Sachin 3.01 added ctMaxDopOpt
      Call moCn.RepRes(mRep.Subreport3.OpenSubreport, wRmDet + ctMaxDopOpt)  '****** Sachin 3.01 added ctMaxDopOpt
      
      Set wo_rsPic = moCn.OpenRes("Select DmTcTyp as DmTcTyp,OdDmCd as DmCd,OdDmSz as DmSz,OmCoCd,OmTc,OmYy,OmChr,OmNo,OdSr,OdPicNm as PicNm from #TmpA" + _
                                  " order by OmCoCd, OmYy, OmTc, OmChr, OmNo,qRow,(Case when qCol=0 Then 3 Else qCol End)")
      If wo_rsPic.RecCount > 0 Then CmdChgImg.Visible = True
    End If
    
  '  Rep.Subreport1.OpenSubreport.Database.SetDataSource
     
     '****************************Geeta****************************
    '************Setting Group4 & Group5 Formula***********************
    If (UCase(ADC.MenuCd) = UCase("RepDetOrd") Or UCase(ADC.MenuCd) = UCase("RepAvgOrd") Or UCase(ADC.MenuCd) = UCase("RepFgDet") Or UCase(ADC.MenuCd) = UCase("RepFgAvg") _
        Or UCase(ADC.MenuCd) = UCase("RepJtDet") Or UCase(ADC.MenuCd) = UCase("RepJtAvg")) Then
    
            If Trim(ADC("UoInGrp")) <> "" Then
              GRP_REP.SetFormula mRep, "wGrp4", wGrpFormula1
              GRP_REP.SetFormula mRep, "wGrp4Desc", wGrpDescFormula1
            End If
            'If Trim(adc("UoInGrp1")) <> "" Or adc("UoYN11") = "Y" Then   '*** (Bef 2.11)
            If Trim(ADC("UoInGrp1")) <> "" Or ADC("UoYN11") <> "" Then
              GRP_REP.SetFormula mRep, "wGrp5", wGrpFormula2
              GRP_REP.SetFormula mRep, "wGrp5Desc", wGrpDescFormula2
            End If
            
            GRP_REP.SetFormula mRep, "wGrpDesc", IIF(wGrpDescFormula1 = "", "", "{@wGrp4Desc}") + _
                                                 IIF(wGrpDescFormula1 = "" Or wGrpDescFormula2 = "", "", "+' / '+") + _
                                                 IIF(wGrpDescFormula2 = "", "", "{@wGrp5Desc}")
    End If
    
    
    GRP_REP.SetFormula mRep, "wOrLmeSal", " If ({rdo.RmCtg} ='G' or {rdo.RmCtg} ='P' or {rdo.RmCtg} ='S' or {rdo.RmCtg} ='L') " + _
                                      " and Round({rdo.OrLmeSal}, 3) <> 0.00 and '" + ADC("UoYn14") + "'='Y' Then Round({rdo.OrLmeSal}, 3) else 0.00 "
    
    GRP_REP.SetFormula mRep, "wLnSv", " If ({rdo.RmCtg} ='G' or {rdo.RmCtg} ='P' or {rdo.RmCtg} ='S' or {rdo.RmCtg} ='L') " + _
                                      " then (if Round({rdo.OrLmeSal}, 3) = 0.00 Then Trim (ToText ({rdo.OrLn1}, 3)) )  " + _
                                      " else (If Round({rdo.OrLn1}, 3)= 0.01 Then '+0 ' " + _
                                      " else if Round({rdo.OrLn1}, 3)= 0.02 Then '+00 '  " + _
                                      " else if Round({rdo.OrLn1}, 3)= 0.03 Then '+000 '  " + _
                                      " else if Round({rdo.OrLn1}, 3)= 0 Then ' '  " + _
                                      " else Trim (ToText ({rdo.OrLn1}, 3)) ) "
    
    ' ***** Manali 3.10.0 - 14/03/12 - Show Net Wt instead of Gld As Wt
    If ADC("UoYN28") = "Y" Then
      GRP_REP.SetFormula Rep, "wOdGldAsWt", "{rdo.OdGldAsWt}"
      Rep.TxtGldAsWtAvg.SetText "GldAsWt": Rep.TxtGldAsWtDet.SetText "GldAsWt"
    Else
      GRP_REP.SetFormula Rep, "wOdGldAsWt", "Sum ({@wNetWt},{@OdSr})"
      Rep.TxtGldAsWtAvg.SetText "Net Wt": Rep.TxtGldAsWtDet.SetText "Net Wt"
    End If
    ' ***** Manali 3.10.0 - 14/03/12 - Show Net Wt instead of Gld As Wt
    
' Bef 3.9.0
'    GRP_REP.SetFormula mRep, "wLnSv", "If ({rdo.RmCtg} ='G' or {rdo.RmCtg} ='P' or {rdo.RmCtg} ='S' or {rdo.RmCtg} ='L') " + _
'                                      "and Round({rdo.OrLmeSal}, 3) <> 0.00 and '" + adc("UoYn14") + "'='Y' Then Trim (ToText ({rdo.OrLmeSal}, 3)) " + _
'                                      "else (If Round({rdo.OrLn1}, 3)= 0.01 Then '+0 ' " + _
'                                      "else if Round({rdo.OrLn1}, 3)= 0.02 Then '+00 '  " + _
'                                      "else if Round({rdo.OrLn1}, 3)= 0.03 Then '+000 '  " + _
'                                      "else if Round({rdo.OrLn1}, 3)= 0 Then ' '  " + _
'                                      "else Trim (ToText ({rdo.OrLn1}, 3)) ) "
    ' **** Bef 3.03 - Manali *****
'    GRP_REP.SetFormula mRep, "wLnSv", "If Round({rdo.OrLn1}, 3)= 0.01 Then '+0 ' " + _
'                                       "else if Round({rdo.OrLn1}, 3)= 0.02 Then '+00 '  " + _
'                                       "else if Round({rdo.OrLn1}, 3)= 0.03 Then '+000 '  " + _
'                                       "else if Round({rdo.OrLn1}, 3)= 0 Then ' '  " + _
'                                       "else Trim (ToText ({rdo.OrLn1}, 3)) "
     '****************************Geeta****************************
      '***********Geeta*****Emr207******
    '*** Jay 3.3(FG)
    ' ***** Manali Trading Module
    If UCase(ADC.MenuCd) = UCase("RepDetOrd") Or UCase(ADC.MenuCd) = UCase("RepAvgOrd") Or UCase(ADC.MenuCd) = UCase("RepFgDet") Or UCase(ADC.MenuCd) = UCase("RepFgAvg") _
      Or UCase(ADC.MenuCd) = UCase("RepJtDet") Or UCase(ADC.MenuCd) = UCase("RepJtAvg") Then
    'If UCase(ADC.MenuCd) = UCase("RepDetOrd") Or UCase(ADC.MenuCd) = UCase("RepAvgOrd") Then
    '*** Jay 3.3(FG)
      If UCase(ADC("UoYN7")) = "N" Then _
          Rep.FldOdDmCd.Suppress = True
      
      ' ***** Manali 3.8.0 - CmName
      If GetUsrLevel > 5 Then Rep.PhCmName.Suppress = True
    End If  '*** Jay 3.3(FG)
    '********************************
    '*** Jay 3.3(FG)
    ' ***** Manali Trading Module
    If UCase(ADC.MenuCd) = UCase("RepDetOrd") Or UCase(ADC.MenuCd) = UCase("RepFgDet") Or UCase(ADC.MenuCd) = UCase("RepJtDet") Then
    'If UCase(ADC.MenuCd) = UCase("RepDetOrd") Then
    '*** Jay 3.3(FG)
      GRP_REP.SetFormula mRep, "wGrp3", "{rdo.qTag}+ToText({rdo.qMainMet}, '#')+ToText({rdo.OrSrNo}, '##0')+{rdo.RmCtg}+{rdo.OrRmSCtg}"
      Select Case UCase(ADC("UoTcTypFr"))
      Case Is = "SO"
        mRep.TxtHead.SetText "Detail Order Printing"
        mRep.TxtVchNo.SetText "Order": mRep.TxtVchDt.SetText "Ord Dt/" + wDelStr + " Del Dt"
      Case Is = "QT"
        mRep.TxtHead.SetText "Detail Quotation Printing"
        mRep.TxtVchNo.SetText "Quotation": mRep.TxtVchDt.SetText "Date"
      Case Is = "PL"
        mRep.TxtHead.SetText "Detail PDC Printing"
        mRep.TxtVchNo.SetText "PDC": mRep.TxtVchDt.SetText "PDC Date"
      Case Is = "QS"
        mRep.TxtHead.SetText "Detail Sketch Quotation Printing"
        mRep.TxtVchNo.SetText "Sketch Vch": mRep.TxtVchDt.SetText "Date"
      Case Is = "QB"
        mRep.TxtHead.SetText "Detail Bag Quotation Printing"
        mRep.TxtVchNo.SetText "Bag Quot": mRep.TxtVchDt.SetText "Date"
      '*** Jay 3.3(FG)
      Case Is = "FG"
        mRep.TxtHead.SetText "Detail Finish Goods Printing"
        mRep.TxtVchNo.SetText "Finish Goods": mRep.TxtVchDt.SetText "Date"
      '*** Jay 3.3(FG)
      ' ***** Manali Trading Module
      Case Is = "JT"
        mRep.TxtHead.SetText "Detail Trading Voucher Printing"
        mRep.TxtVchNo.SetText "Voucher": mRep.TxtVchDt.SetText "Date"
        mRep.TxtPgVchNo.SetText "Voucher": mRep.TxtPgVchDt.SetText "Date"
      ' ***** Manali Trading Module
      End Select
      
      mRep.PHDET.Suppress = False
      mRep.PhAvg.Suppress = True
      mRep.AvgDet.Suppress = True
      
      If ADC("UoRmGrp") = "N" Then
        mRep.RMHEAD.Suppress = True
        mRep.RMDET.Suppress = True
      End If
      
      If ADC("UoLabGrp") = "N" Then
        mRep.LABHEAD.Suppress = True
        mRep.LABDET.Suppress = True
      End If
     '************ urmi 2.05 Show Remarks Y/N
     
     ' ***** Manali 3.9.0 - Different Option for Design Level Remrks
      If ADC("UoYn6") = "N" Then
        Rep.SecOmRemDet1.Suppress = True
        Rep.SecOmRemDet2.Suppress = True
        Rep.SecOmRemDet3.Suppress = True
      End If
      
      ' ***** Manali 3.9.0 - Different Option for Design Level Remrks
      If ADC("UoYn26") = "N" Then
        Rep.SecRemDet1.Suppress = True
        Rep.SecRemDet2.Suppress = True
        Rep.SecRemDet3.Suppress = True
        Rep.SecRemDet4.Suppress = True
        Rep.SecRemDet5.Suppress = True
      End If
      
      '**************************************
      If UCase(ADC("UoTcTypFr")) <> "SO" Then
        Rep.FldOmDelDt.Suppress = True
        Rep.TxtSlash.Suppress = True
      End If
      
      ' Zubin 212 (26/02/06)
      GRP_REP.SetFormula mRep, "wGldRtDAFlg", IIF(moCn.GetFldVal("Select HGldRtDAYN from Head where HCd= '" + ctSelfCmCd + "' and HCoCd= '" + ADC("UOCOCDFR") + "'") = "Y", "'D'", "'A'")
      ' Zubin 212 (26/02/06)
      
      Rep.FldTrayNo.Suppress = True     ' ****** Manali 3.6.0 - 05/11/09 - Tray No in Invoice Design
    
    '*** Jay 3.3(FG)
    ' ***** Manali Trading Module
    ElseIf UCase(ADC.MenuCd) = UCase("RepAvgOrd") Or UCase(ADC.MenuCd) = UCase("RepFgAvg") Or UCase(ADC.MenuCd) = UCase("RepJtAvg") Then
    '*** Jay 3.3(FG)
      GRP_REP.SetFormula mRep, "wGrp3", "{rdo.qTag}+ToText({rdo.qMainMet}, '#')+{rdo.RmCtg}+{rdo.OrRmSCtg}"
      If UCase(ADC("UoTcTypFr")) <> "SO" Then
        Rep.FldOmDelDt.Suppress = True
        Rep.TxtSlash.Suppress = True
      End If
        
        Rep.SecRemDet1.Suppress = True
        Rep.SecRemDet2.Suppress = True
        Rep.SecRemDet3.Suppress = True
        Rep.SecRemDet4.Suppress = True
        Rep.SecRemDet5.Suppress = True
        Rep.SecOmRemDet1.Suppress = True
        Rep.SecOmRemDet2.Suppress = True
        Rep.SecOmRemDet3.Suppress = True
      
      Select Case UCase(ADC("UoTcTypFr"))
      Case Is = "SO"
        mRep.TxtHead.SetText "Average Order"
        mRep.TxtVchNo.SetText "Order": mRep.TxtVchDt.SetText "Ord Dt/" + wDelStr + " Del Dt"
      Case Is = "QT"
        mRep.TxtHead.SetText "Average Quotation"
        mRep.TxtVchNo.SetText "Quotation": mRep.TxtVchDt.SetText "Date"
      Case Is = "PL"
        mRep.TxtHead.SetText "Average PDC"
        mRep.TxtVchNo.SetText "PDC": mRep.TxtVchDt.SetText "PDC Date"
      Case Is = "QS"
        mRep.TxtHead.SetText "Average Sketch Quotation"
        mRep.TxtVchNo.SetText "Sketch Vch": mRep.TxtVchDt.SetText "Date"
      Case Is = "QB"
        mRep.TxtHead.SetText "Average Bag Quotation"
        mRep.TxtVchNo.SetText "Bag Quot": mRep.TxtVchDt.SetText "Date"
      Case Is = "FG"
        mRep.TxtHead.SetText "Average Finish Goods Printing"
        mRep.TxtVchNo.SetText "Finish Goods": mRep.TxtVchDt.SetText "Date"
      Case Is = "JT"
        mRep.TxtHead.SetText "Average Trading Voucher Printing"
        mRep.TxtVchNo.SetText "Voucher": mRep.TxtVchDt.SetText "Date"
        mRep.TxtPgVchNo.SetText "Voucher": mRep.TxtPgVchDt.SetText "Date"
      End Select
      
      mRep.PhAvg.Suppress = False
      mRep.PHDET.Suppress = True
      mRep.AvgDet.Suppress = False
      mRep.RMHEAD.Suppress = True
      mRep.RMDET.Suppress = True
      mRep.LABHEAD.Suppress = True
      mRep.LABDET.Suppress = True
      mRep.FldDiaQty.Suppress = True
      mRep.FldDiaWt.Suppress = True
      mRep.FldColQty.Suppress = True
      mRep.FldColWt.Suppress = True
    
    ElseIf UCase(ADC.MenuCd) = UCase("RepOrdCat") Or UCase(ADC.MenuCd) = UCase("RepFgCat") Or UCase(ADC.MenuCd) = UCase("RepJtCat") Then
      
      Select Case UCase(ADC("UoTcTypFr"))
      Case Is = "SO"
        mRep.TxtHead.SetText "Order Catalogue"
        mRep.TxtVchNo.SetText "Order": mRep.TxtVchDt.SetText "Order Date"
      Case Is = "QT"
        mRep.TxtHead.SetText "Quotation Catalogue"
        mRep.TxtVchNo.SetText "Quotation": mRep.TxtVchDt.SetText "Date"
      Case Is = "PL"
        mRep.TxtHead.SetText "PDC Catalogue"
        mRep.TxtVchNo.SetText "PDC": mRep.TxtVchDt.SetText "PDC Date"
      Case Is = "QS"
        mRep.TxtHead.SetText "Sketch Quotation Catalogue"
        mRep.TxtVchNo.SetText "Sketch": mRep.TxtVchDt.SetText "Date"
      Case Is = "QT"
        mRep.TxtHead.SetText "Bag Quotation Catalogue"
        mRep.TxtVchNo.SetText "Bag Quot": mRep.TxtVchDt.SetText "Date"
      Case Is = "FG"
        mRep.TxtHead.SetText "Finish Goods Catalogue"
        mRep.TxtVchNo.SetText "Finish Goods": mRep.TxtVchDt.SetText "Date"
      Case Is = "JT"
        mRep.TxtHead.SetText "Trading Catalogue"
        mRep.TxtVchNo.SetText "Voucher": mRep.TxtVchDt.SetText "Date"
        mRep.TxtPgVchNo.SetText "Voucher": mRep.TxtPgVchDt.SetText "Date"
      End Select

      If ADC("UoYN15") = "Y" Or ADC("UoYN24") = "Y" Then
          Rep1.DetOdCol.Suppress = False
          If ADC("UoYN15") = "N" And ADC("UoYN24") = "Y" Then
            Rep1.lblOdCol.Suppress = True: Rep1.lblColOdCol.Suppress = True: Rep1.fldOdDmCol.Suppress = True
            Rep1.lblOdCol2.Suppress = True: Rep1.lblColOdCol2.Suppress = True: Rep1.fldOdDmCol2.Suppress = True
            Rep1.lblOdCol3.Suppress = True: Rep1.lblColOdCol3.Suppress = True: Rep1.fldOdDmCol3.Suppress = True
            
            Rep1.FldOdSr.Suppress = False: Rep1.lblOdSr.Suppress = False: Rep1.lblColOdSr.Suppress = False
            Rep1.FldOdSr2.Suppress = False: Rep1.lblOdSr2.Suppress = False: Rep1.lblColOdSr2.Suppress = False
            Rep1.FldOdSr3.Suppress = False: Rep1.lblOdSr3.Suppress = False: Rep1.lblColOdSr3.Suppress = False
            Rep1.lblOdSr.Left = 70: Rep1.lblColOdSr.Left = 1090: Rep1.FldOdSr.Left = 1240
            
          ElseIf Not (ADC("UoYN15") = "Y" And ADC("UoYN24") = "Y") Then
            Rep1.FldOdSr.Suppress = True: Rep1.lblOdSr.Suppress = True: Rep1.lblColOdSr.Suppress = True
            Rep1.FldOdSr2.Suppress = True: Rep1.lblOdSr2.Suppress = True: Rep1.lblColOdSr2.Suppress = True
            Rep1.FldOdSr3.Suppress = True: Rep1.lblOdSr3.Suppress = True: Rep1.lblColOdSr3.Suppress = True
            
            Rep1.lblOdCol.Suppress = False: Rep1.lblColOdCol.Suppress = False: Rep1.fldOdDmCol.Suppress = False
            Rep1.lblOdCol2.Suppress = False: Rep1.lblColOdCol2.Suppress = False: Rep1.fldOdDmCol2.Suppress = False
            Rep1.lblOdCol3.Suppress = False: Rep1.lblColOdCol3.Suppress = False: Rep1.fldOdDmCol3.Suppress = False
          End If
      Else
        Rep1.DetOdCol.Suppress = True
      End If

      If ADC("UoYN16") = "Y" Then Rep1.DetOdWt.Suppress = False Else Rep1.DetOdWt.Suppress = True
      
      If ADC("UoYN17") = "Y" Then
            Rep1.Subreport1_SecRmCd.Suppress = False
            Rep1.Subreport2_SecRmCd.Suppress = False
            Rep1.Subreport3_SecRmCd.Suppress = False
      Else
            Rep1.Subreport1_SecRmCd.Suppress = True
            Rep1.Subreport2_SecRmCd.Suppress = True
            Rep1.Subreport3_SecRmCd.Suppress = True
      End If
      
      If ADC("UoYN18") = "Y" Then
        Rep1.Subreport1_SecRmSCtg.Suppress = False
        Rep1.Subreport2_SecRmSCtg.Suppress = False
        Rep1.Subreport3_SecRmSCtg.Suppress = False
      Else
        Rep1.Subreport1_SecRmSCtg.Suppress = True
        Rep1.Subreport2_SecRmSCtg.Suppress = True
        Rep1.Subreport3_SecRmSCtg.Suppress = True
      End If
        
        
      If ADC("UoYN19") = "Y" Then
         Rep1.Subreport1_SecRmWt.Suppress = False
         Rep1.Subreport2_SecRmWt.Suppress = False
         Rep1.Subreport3_SecRmWt.Suppress = False
      Else
         Rep1.Subreport1_SecRmWt.Suppress = True
         Rep1.Subreport2_SecRmWt.Suppress = True
         Rep1.Subreport3_SecRmWt.Suppress = True
      End If
         
         
      If ADC("UoYN20") = "Y" Then
        Rep1.Subreport1_SecRmQty.Suppress = False
        Rep1.Subreport2_SecRmQty.Suppress = False
        Rep1.Subreport3_SecRmQty.Suppress = False
      Else
        Rep1.Subreport1_SecRmQty.Suppress = True
        Rep1.Subreport2_SecRmQty.Suppress = True
        Rep1.Subreport3_SecRmQty.Suppress = True
      End If

      If ADC("UoYN21") = "Y" Then
        If Rep1.FldDmCd.Value <> "" Then _
            Rep1.TxtUnitPrcCur.Suppress = False: Rep1.FldOdSalPrc.Suppress = False: Rep1.TxtUnitPrcCurCol.Suppress = False
        
        If Rep1.FldDmCd2.Value <> "" Then _
            Rep1.TxtUnitPrcCur2.Suppress = False: Rep1.FldOdSalPrc2.Suppress = False: Rep1.TxtUnitPrcCurCol2.Suppress = False
        
        If Rep1.FldDmCd3.Value <> "" Then _
            Rep1.TxtUnitPrcCur3.Suppress = False: Rep1.FldOdSalPrc3.Suppress = False: Rep1.TxtUnitPrcCurCol3.Suppress = False
      Else
        Rep1.TxtUnitPrcCur.Suppress = True:  Rep1.FldOdSalPrc.Suppress = True   'Rep1.TxtUnitPrcCurCol.Suppress = True:
        Rep1.TxtUnitPrcCur2.Suppress = True:  Rep1.FldOdSalPrc2.Suppress = True   'Rep1.TxtUnitPrcCurCol.Suppress = True:
        Rep1.TxtUnitPrcCur3.Suppress = True:  Rep1.FldOdSalPrc3.Suppress = True   'Rep1.TxtUnitPrcCurCol3.Suppress = True:
      End If
      
      
      If ADC("UoYN27") = "Y" Then
        
        If Rep1.FldDmCd.Value <> "" Then _
            Rep1.TxtQty.Suppress = False: Rep1.FldqOrdQty.Suppress = False
        
        If Rep1.FldDmCd2.Value <> "" Then _
            Rep1.TxtQty2.Suppress = False: Rep1.FldqOrdQty2.Suppress = False
        
        If Rep1.FldDmCd3.Value <> "" Then _
            Rep1.TxtQty3.Suppress = False: Rep1.FldqOrdQty3.Suppress = False
        
        If ADC("UoYN21") = "N" Then
            Rep1.TxtQty.Left = Rep1.TxtUnitPrcCur.Left: Rep1.FldqOrdQty.Left = Rep1.FldOdSalPrc.Left
            Rep1.TxtQty2.Left = Rep1.TxtUnitPrcCur2.Left: Rep1.FldqOrdQty2.Left = Rep1.FldOdSalPrc2.Left
            Rep1.TxtQty3.Left = Rep1.TxtUnitPrcCur3.Left: Rep1.FldqOrdQty3.Left = Rep1.FldOdSalPrc3.Left
        End If
      Else
        Rep1.TxtQty.Suppress = True: Rep1.FldqOrdQty.Suppress = True
        Rep1.TxtQty2.Suppress = True: Rep1.FldqOrdQty2.Suppress = True
        Rep1.TxtQty3.Suppress = True: Rep1.FldqOrdQty3.Suppress = True
      End If

      If ADC("UoYN21") = "N" And ADC("UoYN27") = "N" Then Rep1.DetOdPrc.Suppress = True

      If ADC("UoYN22") = "Y" Then Rep1.DetodPoNo.Suppress = False Else Rep1.DetodPoNo.Suppress = True
      If ADC("UoYN17") = "N" And ADC("UoYN18") = "N" And ADC("UoYN19") = "N" And ADC("UoYN20") = "N" Then _
      Rep1.DetCatSubRep.Suppress = True Else Rep1.DetCatSubRep.Suppress = False
    End If

    If ADC("UORMGRP") = UCase("RmRSzRt") Then
      Rep.TxtFrLn.Suppress = True
      Rep.TxtToLn.Suppress = True
      Rep.OrLn1.Suppress = True
      Rep.OrLn2.Suppress = True
      Rep.TxtRngSz.Suppress = False
    Else
      Rep.OrLn1.Suppress = False
      Rep.OrLn2.Suppress = False
      Rep.TxtRngSz.Suppress = True
      Rep.RmRSzRt.Suppress = True
    End If

    If ADC("UoYN1") = "Y" Then
      Rep.TxtDsgQty.SetText "Amd Qty"
      Rep.TxtDsgQty1.SetText "Amd Qty"
      Rep1.TxtQty.SetText "Amd Qty"
      Rep1.TxtQty2.SetText "Amd Qty"
      Rep1.TxtQty3.SetText "Amd Qty"
    ElseIf ADC("UoYN1") = "N" Then
      Rep.TxtDsgQty.SetText "Ord Qty"
      Rep.TxtDsgQty1.SetText "Ord Qty"
      Rep1.TxtQty.SetText "Ord Qty"
      Rep1.TxtQty2.SetText "Ord Qty"
      Rep1.TxtQty3.SetText "Ord Qty"
    End If

    If ADC("UoYN7") = "Y" Then
      Rep.hBOdDmCd.SetText "Customer Design"
      Rep.Text20.SetText "Customer Design"
      Rep1.Text1.SetText "CmDsg/Sz"
    Else
      Rep.hBOdDmCd.SetText "Design"
      Rep.Text20.SetText "Design"
      Rep1.Text1.SetText "Design/Sz"
    End If

    If ADC("UoYN8") = "Y" Then
        Rep.Text45.SetText "Cust Quality"
      Else
        Rep.Text45.SetText "Quality"
      End If

    If ADC("UoYN9") = "N" And (UCase(ADC.MenuCd) = UCase("RepDetOrd") Or UCase(ADC.MenuCd) = UCase("RepAvgOrd") Or UCase(ADC.MenuCd) = UCase("RepFgDet") Or UCase(ADC.MenuCd) = UCase("RepFgAvg") _
                              Or UCase(ADC.MenuCd) = UCase("RepJtDet") Or UCase(ADC.MenuCd) = UCase("RepJtAvg")) Then
      Rep.SecSumm.Suppress = True
      Rep.SecLabSumm.Suppress = True    ' ***** Manali 3.6.0 - 31/10/09 - Lab Summary
    Else
      Rep.SecSumm.Suppress = False
      Rep.SecLabSumm.Suppress = False   ' ***** Manali 3.6.0 - 31/10/09 - Lab Summary
    End If

    If UCase(ADC.MenuCd) <> UCase("RepJtDet") Then
      GRP_REP.SetFormula mRep, "wExpOdNo", ""
      Rep.LblExpOdNo.Suppress = True
    End If

    Rep.TxtOdomPoNo.Suppress = True: Rep.TxtOdPoNo.Suppress = False

    If moCn.GetFldVal("Select HPtrYN from Head where HCoCd = '" + ADC("UoCoCdFr") + "' and " + _
                      "HCd= '" + ctSelfCmCd + "'") = "N" Then _
       Rep.hOrRmPtr.Suppress = True: Rep.FldOrRmPtr.Suppress = True
    
    CRV_REP.DisplayGroupTree = False
    CmdSc.Top = Me.ScaleHeight - CmdSc.Height
    CmdDispSc.Top = Me.ScaleHeight - CmdDispSc.Height   '***** Manali 3.5.0 - 14/11/08 - Ord Cat fields suppressed/displayed
      
    ' ***** Manali 3.8.0
    If ADC("UoYn25") = "Y" Then
      Rep.FldOdSr.Suppress = True: Rep.FldRunSr.Suppress = False
    Else
      Rep.FldOdSr.Suppress = False: Rep.FldRunSr.Suppress = True
    End If
    
    ' Bef 3.8.0
'    Rep.FldOdSr.Suppress = False
'    Rep.FldRunSr.Suppress = True

    'Manoj 2.10.0
    '*** Jay 3.3(FG)
    ' ***** Manali Trading Module
    If UCase(ADC.MenuCd) = UCase("RepFgDet") Or UCase(ADC.MenuCd) = UCase("RepJtDet") Then
      Rep.txtOdDmId1.Suppress = False
      Rep.FldOdDmId.Suppress = False
    ElseIf UCase(ADC.MenuCd) = UCase("RepFgAvg") Or UCase(ADC.MenuCd) = UCase("RepJtAvg") Then
      Rep.TxtOdDmId2.Suppress = False
      Rep.FldOdDmId.Suppress = False
    End If
    '*** Jay 3.3(FG)
    
    ' ***** Manali 3.7.1
    Dim wTcDesc As String
    wTcDesc = moCn.GetFldVal("Select PDesc From Param Where PTyp='TC' and PMCd='" + ADC("UoOmTcFr") + "'")
    If UCase(ADC.MenuCd) = UCase("RepJtDet") Or UCase(ADC.MenuCd) = UCase("RepJtAvg") Then
      Rep.RhJtSec.Suppress = False: Rep.RhLineSec.Suppress = True: Rep.RHMain.UnderlaySection = True
      Rep.TxtHeadDesc.SetText "(" + wTcDesc + ")"
    ElseIf UCase(ADC.MenuCd) = UCase("RepJtCat") Then
      Rep1.TxtHeadDesc.SetText "(" + wTcDesc + ")"
    Else
      Rep.RhJtSec.Suppress = True
    End If
    ' ***** Manali 3.7.1
    
    ' ***** Manali 3.6.1 - Pic Preference
    ' ***** Manali Trading Module - "RepJtCat" added
    If UCase(ADC.MenuCd) <> UCase("RepOrdCat") And UCase(ADC.MenuCd) <> UCase("RepFgCat") And UCase(ADC.MenuCd) <> UCase("RepJtCat") Then
      If UCase(ADC("UoPMCdFr")) = "3D" Then GRP_REP.SetFormula mRep, "wPicOpt", "'3D'"
      If UCase(ADC("UoPMCdFr")) = "HSK" Then GRP_REP.SetFormula mRep, "wPicOpt", "'HSK'"
      'If UCase(adc("UoPMCdFr")) = "" Then GRP_REP.SetFormula mRep, "wPicOpt", "''": mRep.Pic.Suppress = True
    End If
    
  Call DispReport(ADC) '****** Sachin 3.02 - Display Report ******
  Exit Sub
RepErr:
  Call ChkSelScope(Err.Number, Err.Description)
End Sub


Private Sub ChkBoxArr_Click(Index As Integer)
With ChkBoxArr(Index)
  Select Case Index
  Case Is = 0     'Tag = "ShowPic"
    If .Value = Checked Then ADC("UoYN") = "Y" Else ADC("UoYN") = "N"
  Case Is = 2     'Tag = "ShowDsgVal"
    If .Value = Checked Then ADC("UoYN2") = "Y" Else ADC("UoYN2") = "N"
  Case Is = 4     'Tag = "ShowLabval"
    If .Value = Checked Then ADC("UoYN4") = "Y" Else ADC("UoYN4") = "N"
  Case Is = 6     'Tag = "ShowRem"
    If .Value = Checked Then ADC("UoYN6") = "Y" Else ADC("UoYN6") = "N"
  Case Is = 7     'Tag = "ShowCustDsg"
    If .Value = Checked Then ADC("UoYN7") = "Y" Else ADC("UoYN7") = "N"
  Case Is = 8     'Tag = "ShowCustRm"
    If .Value = Checked Then ADC("UoYN8") = "Y" Else ADC("UoYN8") = "N"
  Case Is = 9     'Tag = "ShowSumm"
    If .Value = Checked Then ADC("UoYN9") = "Y" Else ADC("UoYN9") = "N"
  Case Is = 11     'Tag = "FinalSrt"
    If .Value = Checked Then ADC("UoYN11") = "Y" Else ADC("UoYN11") = "N"
  Case Is = 12     'Tag = "ShowComp"
    If .Value = Checked Then ADC("UoYN12") = "Y" Else ADC("UoYN12") = "N"
  Case Is = 13     'Tag = "ShowGrpDesc"
    If .Value = Checked Then ADC("UoYN13") = "Y" Else ADC("UoYN13") = "N"
  Case Is = 14          'Tag = "ShowRmVal"
    If .Value = Checked Then ADC("UoYN14") = "Y" Else ADC("UoYN14") = "N"
  Case Is = 15          'Tag = "ShowDsgCol"
    If .Value = Checked Then ADC("UoYN15") = "Y" Else ADC("UoYN15") = "N"
  Case Is = 16         'Tag = "ShowGrsWt"
    If .Value = Checked Then ADC("UoYN16") = "Y" Else ADC("UoYN16") = "N"
  Case Is = 17          'Tag = "ShowRmCd"
    If .Value = Checked Then ADC("UoYN17") = "Y" Else ADC("UoYN17") = "N"
  Case Is = 18          'Tag = "ShowRmSCtg"
    If .Value = Checked Then ADC("UoYN18") = "Y" Else ADC("UoYN18") = "N"
  Case Is = 19          'Tag = "ShowRmWt"
    If .Value = Checked Then ADC("UoYN19") = "Y" Else ADC("UoYN19") = "N"
  Case Is = 20         'Tag = "ShowRmQty"
    If .Value = Checked Then ADC("UoYN20") = "Y" Else ADC("UoYN20") = "N"
  Case Is = 21         'Tag = "ShowUnitPrc"
    If .Value = Checked Then ADC("UoYN21") = "Y" Else ADC("UoYN21") = "N"
  Case Is = 22          'Tag = "ShowOdPO"
    If .Value = Checked Then ADC("UoYN22") = "Y" Else ADC("UoYN22") = "N"
  Case Is = 23          'Tag = "ShowExcel"
    If .Value = Checked Then ADC("UoYN23") = "Y" Else ADC("UoYN23") = "N"
  Case Is = 24          'Tag = "ShowVchSr"
    If .Value = Checked Then ADC("UoYN24") = "Y" Else ADC("UoYN24") = "N"
  Case Is = 25          'Tag = "ShowVchSr"    ' ***** Manali 3.8.0
    If .Value = Checked Then ADC("UoYN25") = "Y" Else ADC("UoYN25") = "N"
  Case Is = 26          'Tag = "ShowMstRem"    ' ***** Manali 3.9.0
    If .Value = Checked Then ADC("UoYN26") = "Y" Else ADC("UoYN26") = "N"
  Case Is = 27          'Tag = "ShowOrdQty"    ' ***** Manali 3.10.0 - 09/03/12 - Different Suppres option for Unit Pricce and Qty
    If .Value = Checked Then ADC("UoYN27") = "Y" Else ADC("UoYN27") = "N"
  Case Is = 30          'Tag = "ShowIntlQty"    ' 4.1.2
    If .Value = Checked Then ADC("UoYN30") = "Y" Else ADC("UoYN30") = "N"
End Select
End With
End Sub
Private Sub ChkBoxArr_GotFocus(Index As Integer)
DispMsg ChkBoxArr(Index).ToolTipText, etInfo
End Sub

Private Sub ChkBoxLme_Click(Index As Integer)
' ***** Manali 3.9.0 - Lme Selection
With ChkBoxLme(Index)
  If .Value = Checked Then wChkSel = wChkSel + 1 Else wChkSel = wChkSel - 1
  If wChkSel > 2 Then DispMsg "Cannot Select More Than 2 Lmes", etError: .Value = Unchecked: Exit Sub

  Dim wVal As String, i As Integer
  wVal = .Caption

  If .Value = Checked Then
    If ADC("UoLme1") <> wVal And ADC("UoLme2") <> wVal Then
'      If adc("UoLme1") = "" Then
'        adc("UoLme1") = wVal
'      Else
'        adc("UoLme2") = wVal
'      End If
      
      If ADC("UoLme1") <> "" Then
        For i = Index + 1 To 3
          If ADC("UoLme1") = ChkBoxLme(i).Caption Then
            ADC("UoLme2") = ADC("UoLme1")
            ADC("UoLme1") = wVal
            Exit For
          End If
        Next i
      End If
      If ADC("UoLme1") = "" Then
        ADC("UoLme1") = wVal
      ElseIf ADC("UoLme2") = "" Then
        ADC("UoLme2") = wVal
      End If
    End If
  Else
    If ADC("UoLme1") = wVal Then
      ADC("UoLme1") = ADC("UoLme2")
      ADC("UoLme2") = ""
    ElseIf ADC("UoLme2") = wVal Then
      ADC("UoLme2") = ""
    End If
  End If
End With
End Sub

Private Sub ChkBoxLme_GotFocus(Index As Integer)
DispMsg FraLme.ToolTipText, etInfo
End Sub

Private Sub CmdChgImg_Click()
If FraChgImg.Visible = False Then
    If CRV_REP.Visible = True Then Exit Sub
    FraChgImg.Enabled = True
    FraChgImg.Visible = True
    TAB_REP.Enabled = False
    ADC.AllowSave = False
    FraChgImg.ZOrder
  Else
    TAB_REP.Enabled = True
    ADC.AllowSave = True
    FraChgImg.Enabled = False
    FraChgImg.Visible = False
  End If
End Sub
Private Sub CmdChgImgGo_Click()
On Error GoTo ErrHd
Dim wPicHeight As Single, wPicWidth As Single
If UCase(ADC.MenuCd) = UCase("RepOrdCat") Or UCase(ADC.MenuCd) = UCase("RepFgCat") Or UCase(ADC.MenuCd) = UCase("RepJtCat") Then
    wPicHeight = 110: wPicWidth = 110
Else
    wPicHeight = 84: wPicWidth = 84
End If

If ADC("wXlFilePath") = "" Then
    DispMsg "Please Provide an Excel File to change Images", etError
    Exit Sub
End If
 
If ReplXlImg(ADC("wXlFilePath"), wo_rsPic, wPicHeight, wPicWidth) Then
    TAB_REP.Enabled = True
    ADC.AllowSave = True
    FraChgImg.Visible = False
    CmdChgImg.Visible = False
    DispMsg "Images Replaced Successfully", etInfo
End If
  Exit Sub
ErrHd:
    DispMsg Err.Description, etError
End Sub
Private Sub CmdDispSc_Click()
' **** Manali 3.5.0 - 14/11/08 - Supress or Display fields in Ord Cat
  If FraDispSc.Visible = False Then
    If CRV_REP.Visible = True Then Exit Sub
    FraDispSc.Enabled = True
    FraDispSc.Visible = True
    TAB_REP.Enabled = False
    ADC.AllowSave = False
    FraDispSc.ZOrder
    'adc("UoYN15").SetFocus
    CmdSc.Enabled = False
  Else
    TAB_REP.Enabled = True
    ADC.AllowSave = True
    FraDispSc.Enabled = False
    FraDispSc.Visible = False
    CmdDispSc.SetFocus
    CmdSc.Enabled = True
  End If
' **** Manali 3.5.0 - 14/11/08 - Supress or Display fields in Ord Cat
End Sub

Private Sub CmdSelFile_Click()
'Sachin - 3.11.2
    On Error GoTo ErrHd
    Dim wExtn As String
    wExtn = ".xls;*.xlsx"
    With Cd1
      .FileName = ""
      .DialogTitle = "Open Report Excel"
      .Filter = "Excel Files (*" + wExtn + ") |*" + wExtn
      .CancelError = False
      .ShowOpen
      If .FileName <> "" Then ADC("WXLFILEPATH") = .FileName
    End With
    Exit Sub
ErrHd:
    DispMsg Err.Description, etError
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
    '*** Jay 3.3(FG)
    If UCase(ADC.MenuCd) = UCase("RepFgDet") Or UCase(ADC.MenuCd) = UCase("RepFgAvg") Or UCase(ADC.MenuCd) = UCase("RepFgCat") Or _
        UCase(ADC.MenuCd) = UCase("RepJtDet") Or UCase(ADC.MenuCd) = UCase("RepJtAvg") Or UCase(ADC.MenuCd) = UCase("RepJtCat") Then
      Set ADC.FirNKeyCtl = ADC("UoOmTcFr") 'adc("UoCoCdFr")
    Else
      Set ADC.FirNKeyCtl = ADC("UoTcTypFr") 'adc("UoCoCdFr")
    End If
    '*** Jay 3.3(FG)
    Call InitProp(Me)
    
    ' ***** Manali Trading Module
    gs_CmCtg = "C"
    gs_EntTyp = "OM"
    
    '*** Jay 3.3(FG)
    If UCase(ADC.MenuCd) = UCase("RepDetOrd") Or UCase(ADC.MenuCd) = UCase("RepFgDet") Or UCase(ADC.MenuCd) = UCase("RepJtDet") Then
    'If UCase(ADC.MenuCd) = UCase("RepDetOrd") Then
    '*** Jay 3.3(FG)
      gs_RmGrp = "ORDDET"
      gs_LabGrp = "ORDDET"
      
      ' **** Manali 3.6.1 - Pic Preference
      gs_PTyp = "PCT"
      gs_Tbl = "Param"
      ' **** Manali 3.6.1 - Pic Preference
      
      'Me.Caption = "Detail Order Printing"
      '*****************************************urmi*************************************
      'adc("UoYN2").Visible = False: adc("UoYN2").Enabled = False: LblShwVal.Visible = False
      '*****************************************urmi*************************************
      '*** Jay 3.3(FG)
      If UCase(ADC.MenuCd) = UCase("RepDetOrd") Then
      '*** Jay 2.13(CT) ***
        'adc("UoYN12").Visible = True: adc("UoYN12").Enabled = True: LblShwComp.Visible = True
        ChkBoxArr(12).Enabled = True: ChkBoxArr(12).Visible = True    ' ****** Manali 3.5.0 - 19/11/08 - Check Boxes added
      '*** Jay 2.13(CT) ***
      End If
      '*** Jay 3.3(FG)
      
      ' ***** Manali 3.5.0 - 14/11/08 - Suppress/Display fields in Ord Cat
      CmdDispSc.Enabled = False: CmdDispSc.Visible = False
      ' ***** Manali 3.5.0 - 14/11/08 - Suppress/Display fields in Ord Cat

      ChkBoxArr(24).Enabled = False: ChkBoxArr(24).Visible = False    ' **** Manali 3.6.1 - Sr in Catalogue

    '*** Jay 3.3(FG)
    ElseIf UCase(ADC.MenuCd) = UCase("RepAvgOrd") Or UCase(ADC.MenuCd) = UCase("RepFgAvg") Or UCase(ADC.MenuCd) = UCase("RepJtAvg") Then
    'ElseIf UCase(ADC.MenuCd) = UCase("RepAvgOrd") Then
    '*** Jay 3.3(FG)
      gs_RmGrp = "ORDAVG"
      gs_LabGrp = "ORDAVG"
      
      ' **** Manali 3.6.1 - Pic Preference
      gs_PTyp = "PCT"
      gs_Tbl = "Param"
      ' **** Manali 3.6.1 - Pic Preference
      
      'Me.Caption = "Average Order"   '*** Jay 3.3(FG)
      ADC("UoRmGrp").Visible = False: ADC("UoLabGrp").Visible = False
      ADC("UoRmGrp").Enabled = False: ADC("UoLabGrp").Enabled = False
      LblRmGrp.Visible = False: LblLabGrp.Visible = False
      ' ****** Manali 3.5.0 - 07/11/08 - Radio Buttons added
      lblDelDt.Enabled = True: 'adc("UoYN5").Enabled = True
      lblDelDt.Visible = True: 'adc("UoYN5").Visible = True
      ADC("UoYN5").Enabled = False: ADC("UoYN5").Visible = False
      FraOptDelDt.Enabled = True: FraOptDelDt.Visible = True
      ' ****** Manali 3.5.0 - 07/11/08 - Radio Buttons added
      ' urmi show remark
      ADC("UoYN6").Visible = False: ADC("UoYN6").Enabled = False: lblShowRem.Visible = False
      ChkBoxArr(6).Enabled = False: ChkBoxArr(6).Visible = False    ' ****** Manali 3.5.0 - 19/11/08 - Check Boxes added
      ChkBoxArr(26).Enabled = False: ChkBoxArr(26).Visible = False    ' ****** Manali 3.9.0 - Different Option For Design level Remark
      ' Manoj 2.10.0
      ADC("UoYn10").Visible = False: ADC("UoYN10").Enabled = False: LblCalcPtr.Visible = False
      ' Manoj 2.10.0
      ' ****** Manali 3.5.0 - 07/11/08 - Radio Buttons added
      FraOptPtr.Visible = False: FraOptPtr.Visible = False
      ' ****** Manali 3.5.0 - 07/11/08 - Radio Buttons added
       
      ' ***** Manali 3.5.0 - 14/11/08 - Suppress/Display fields in Ord Cat
      CmdDispSc.Enabled = False: CmdDispSc.Visible = False
      ' ***** Manali 3.5.0 - 14/11/08 - Suppress/Display fields in Ord Cat
      
      ChkBoxArr(23).Enabled = False: ChkBoxArr(23).Visible = False    ' ***** Manali 3.6.0 - 24/09/09 - Show Report in Excel Format
      
      ChkBoxArr(24).Enabled = False: ChkBoxArr(24).Visible = False    ' **** Manali 3.6.1 - Sr in Catalogue
      
    '*** Jay 3.3(FG)
    ElseIf UCase(ADC.MenuCd) = UCase("RepOrdCat") Or UCase(ADC.MenuCd) = UCase("RepFgCat") Or UCase(ADC.MenuCd) = UCase("RepJtCat") Then
    'ElseIf UCase(ADC.MenuCd) = UCase("RepOrdCat") Then
    '*** Jay 3.3(FG)
      'Me.Caption = "Order Catalogue"   '*** Jay 3.3(FG)
      
      ChkBoxArr(30).Enabled = True: ChkBoxArr(30).Visible = True    '4.1.2
      ADC("UoRmGrp").Visible = False: ADC("UoLabGrp").Visible = False
      ADC("UoRmGrp").Enabled = False: ADC("UoLabGrp").Enabled = False
      ADC("UoYN").Visible = False: ADC("UoYN").Enabled = False
      ChkBoxArr(0).Enabled = False: ChkBoxArr(0).Visible = False    ' ****** Manali 3.5.0 - 19/11/08 - Check Boxes added
      LblRmGrp.Visible = False: LblLabGrp.Visible = False
      LblPic.Visible = False
      ADC("UoYN2").Visible = False: ADC("UoYN2").Enabled = False: LblShwVal.Visible = False
      ChkBoxArr(2).Enabled = False: ChkBoxArr(2).Visible = False    ' ****** Manali 3.5.0 - 19/11/08 - Check Boxes added
      ADC("UoYN4").Visible = False: ADC("UoYN4").Enabled = False: LblLabVal.Visible = False
      ChkBoxArr(4).Enabled = False: ChkBoxArr(4).Visible = False    ' ****** Manali 3.5.0 - 19/11/08 - Check Boxes added
      '*** Jay 2.14Next [Show Rm Value]
      ADC("UoYN14").Visible = False: ADC("UoYN14").Enabled = False: LblRmVal.Visible = False
      ChkBoxArr(14).Enabled = False: ChkBoxArr(14).Visible = False    ' ****** Manali 3.5.0 - 19/11/08 - Check Boxes added
      '*** Jay 2.14Next [Show Rm Value]
      lblDelDt.Enabled = False: ADC("UoYN5").Enabled = False
      lblDelDt.Visible = False: ADC("UoYN5").Visible = False
      ' ****** Manali 3.5.0 - 07/11/08 - Radio Buttons added
      FraOptDelDt.Enabled = False: FraOptDelDt.Visible = False
      ' ****** Manali 3.5.0 - 07/11/08 - Radio Buttons added
      
      ' urmi show remark
      ADC("UoYN6").Visible = False: ADC("UoYN6").Enabled = False: lblShowRem.Visible = False
      ChkBoxArr(6).Enabled = False: ChkBoxArr(6).Visible = False    ' ****** Manali 3.5.0 - 19/11/08 - Check Boxes added
      ChkBoxArr(26).Enabled = False: ChkBoxArr(26).Visible = False    ' ****** Manali 3.9.0 - Different Option For Design level Remark
      ' Manoj 2.10.0
      ADC("UoYN9").Visible = False: ADC("UoYN9").Enabled = False: LblShowSumm.Visible = False
      ChkBoxArr(9).Enabled = False: ChkBoxArr(9).Visible = False    ' ****** Manali 3.5.0 - 19/11/08 - Check Boxes added
      ADC("UoYn10").Visible = False: ADC("UoYN10").Enabled = False: LblCalcPtr.Visible = False
      ' Manoj 2.10.0
      ' ****** Manali 3.5.0 - 07/11/08 - Radio Buttons added
      FraOptPtr.Visible = False: FraOptPtr.Visible = False
       ' ****** Manali 3.5.0 - 07/11/08 - Radio Buttons added
      '*** JAY 2.14 ***
      ADC("UOYN13").Visible = False: LblShowGrpDesc.Visible = False
      ChkBoxArr(13).Enabled = False: ChkBoxArr(13).Visible = False    ' ****** Manali 3.5.0 - 19/11/08 - Check Boxes added
      '*** JAY 2.14 ***
      
      ChkBoxArr(23).Enabled = False: ChkBoxArr(23).Visible = False    ' ***** Manali 3.6.0 - 24/09/09 - Show Report in Excel Format
      
      ' ***** Manali 3.6.1 - Pic Preference
      ADC("UoPMCdFr").Enabled = False: ADC("UoPMCdFr").Visible = False
      LblPicOpt.Visible = False
      ' ***** Manali 3.6.1 - Pic Preference
      
      ' ***** Manali 3.8.0
      ChkBoxArr(25).Enabled = False: ChkBoxArr(25).Visible = False
       
      FraLme.Enabled = False: FraLme.Visible = False      ' ***** Manali 3.9.0 - Lme Selection
      LblNetWt.Visible = False: FraOptNetWt.Enabled = False: FraOptNetWt.Visible = False    ' ***** Manali 3.10.0 - 14/03/12 - Show Net Wt instead of Gld As Wt
      
    End If
    '*** Jay 3.3(FG)
    ' **** Manali Trading Module
    If UCase(ADC.MenuCd) = UCase("RepFgDet") Or UCase(ADC.MenuCd) = UCase("RepFgAvg") Or _
       UCase(ADC.MenuCd) = UCase("RepFgCat") Or _
       UCase(ADC.MenuCd) = UCase("RepJtDet") Or UCase(ADC.MenuCd) = UCase("RepJtAvg") Or _
       UCase(ADC.MenuCd) = UCase("RepJtCat") Then
       gs_OmTcTyp = IIF(UCase(ADC.MenuCd) = UCase("RepJtDet") Or UCase(ADC.MenuCd) = UCase("RepJtAvg") Or _
                        UCase(ADC.MenuCd) = UCase("RepJtCat"), ctJtTcTyp, "FG")
        gs_EntTyp = ADC("UOTCTYPFR")
        
       'adc("UOTCTYPFR").Visible = False: adc("UOTCTYPFR").Enabled = False: lblTcTyp.Visible = False
       lblExpDelDt.Visible = False: lblPrdDelDt.Visible = False: LblShwComp.Visible = False
       ADC("UODELDTFR").Visible = False: ADC("UODELDTTo").Visible = False
       ADC("UODELDT1FR").Visible = False: ADC("UODELDT1To").Visible = False
       lblAmdQty.Visible = False:  ADC("UOYn1").Visible = False
       FraOptOrdQty.Enabled = False: FraOptOrdQty.Visible = False    ' ****** Manali 3.5.0 - 19/11/08 - Check Boxes added
       lblPrdRmQtyWt.Visible = False: ADC("UOYn3").Visible = False
       LblCalcPtr.Visible = False: ADC("UOYn10").Visible = False
       ' ****** Manali 3.5.0 - 07/11/08 - Radio Buttons added
       FraOptRm.Enabled = False: FraOptRm.Visible = False
       FraOptPtr.Visible = False: FraOptPtr.Visible = False
       FraOptOrdQty.Enabled = False: FraOptOrdQty.Visible = False
       ' ****** Manali 3.5.0 - 07/11/08 - Radio Buttons added
    End If
    '*** Jay 3.3(FG)
    
    ' ***** Manali 3.6.0 - Order Printing in Excel Format Module
    If moCn.RecSeek("Select 'x' from Head Where HCoCd='" + ctSelfCoCd + "' and HCd='" + ctSelfCmCd + "' and HOrdExcelRepYn<>'Y' ") Then _
      ChkBoxArr(23).Visible = False: ChkBoxArr(23).Enabled = False
    
    Call SetGroupSort
End Sub
Private Sub atxt_FldChange(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String, ByVal pv_NewValue As Variant)
   '*** Jay 3.3(FG)
  gs_EntTyp = "OM"
  If UCase(ADC.MenuCd) = UCase("RepFgDet") Or UCase(ADC.MenuCd) = UCase("RepFgAvg") Or _
     UCase(ADC.MenuCd) = UCase("RepFgCat") Then
     gs_OmTcTyp = "FG"
     gs_EntTyp = "FG" 'adc("UOTCTYPFR")
  ' ***** Manali Trading Module
  ElseIf UCase(ADC.MenuCd) = UCase("RepJtDet") Or UCase(ADC.MenuCd) = UCase("RepJtAvg") Or _
     UCase(ADC.MenuCd) = UCase("RepJtCat") Then
     gs_OmTcTyp = ctJtTcTyp
     gs_EntTyp = ctJtTcTyp    'adc("UOTCTYPFR")
  End If
  '*** Jay 3.3(FG)
  Select Case UCase(IdName)
'  Case Is = UCase("UoYN2")
'    Cancel = pv_NewValue <> "" And (Not moCn.RecSeek("select PMCd from Param where " + _
'                          " PTyp= 'YN' and PMCd = '" + pv_NewValue + "'"))
'    ErrMsg = "Invalid Show Value Option": Exit Sub
'
' *** Jay 2.14Next ***  [Show Rm Value]
'  Case Is = UCase("UoYN4")
'    Cancel = pv_NewValue <> "" And (Not moCn.RecSeek("select PMCd from Param where " + _
'                          " PTyp= 'YN' and PMCd = '" + pv_NewValue + "'"))
'    ErrMsg = "Invalid Labour Value Option": Exit Sub
' *** Jay 2.14Next ***  [Show Rm Value]
  End Select
  
  ' **** Manali 3.6.1 - Pic Preference
  If UCase(ADC.MenuCd) = UCase("RepDetOrd") Or UCase(ADC.MenuCd) = UCase("RepFgDet") Or UCase(ADC.MenuCd) = UCase("RepJtDet") Or _
      UCase(ADC.MenuCd) = UCase("RepAvgOrd") Or UCase(ADC.MenuCd) = UCase("RepFgAvg") Or UCase(ADC.MenuCd) = UCase("RepJtAvg") Then
    gs_PTyp = "PCT"
    gs_Tbl = "Param"
  End If
  ' **** Manali 3.6.1 - Pic Preference
  
  Call SetProp(Me, IdName, Valid, Cancel, ErrMsg)
End Sub
Private Sub ATXT_FldWhen(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String)
  gb_CoCdFor = True         '*** (Jen 2.14 Next Patch 1)
  
  Select Case UCase(IdName)
  Case Is = UCase("UoOmTcFr"), UCase("UoOmTcTo")
    '*** Parameters Used For Setting Properties Of UsrOpt Fields ***
    '*** Jay 3.3(FG)
    If UCase(ADC.MenuCd) = UCase("RepFgDet") Or UCase(ADC.MenuCd) = UCase("RepFgAvg") Or _
     UCase(ADC.MenuCd) = UCase("RepFgCat") Then
      gs_OmTcTyp = "FG"
      gs_EntTyp = "FG" ' adc("UOTCTYPFR")
    ElseIf UCase(ADC.MenuCd) = UCase("RepJtDet") Or UCase(ADC.MenuCd) = UCase("RepJtAvg") Or _
     UCase(ADC.MenuCd) = UCase("RepJtCat") Then
      gs_OmTcTyp = ctJtTcTyp
      gs_CmCtg = moCn.GetFldVal("Select TmCmCtg from TcMst Where TmCoCd='" + ADC("UoCoCdFr") + "' and TmTc='" + ADC("UoOmTcFr") + "'")
      gs_CmCtg = IIF(gs_CmCtg = "", "C", gs_CmCtg)
      gs_EntTyp = ctJtTcTyp 'adc("UOTCTYPFR")
    Else
      gs_OmTcTyp = ADC("UoTcTypFr")
      gs_EntTyp = "OM"
    End If
    '*** Jay 3.3(FG)
  
    ' **** Manali 3.6.1 - Pic Preference
    If UCase(ADC.MenuCd) = UCase("RepDetOrd") Or UCase(ADC.MenuCd) = UCase("RepFgDet") Or UCase(ADC.MenuCd) = UCase("RepJtDet") Or _
        UCase(ADC.MenuCd) = UCase("RepAvgOrd") Or UCase(ADC.MenuCd) = UCase("RepFgAvg") Or UCase(ADC.MenuCd) = UCase("RepJtAvg") Then
      gs_PTyp = "PCT"
      gs_Tbl = "Param"
    End If
    ' **** Manali 3.6.1 - Pic Preference
    
   '*** Parameters Used For Setting Properties Of UsrOpt Fields ***
  Case Is = UCase("UoYN4")
    'If adc("UoYN2") = "N" Then
    '   Cancel = False
    'Else
' Zubin 212 (not req)
'    If adc("UoYN2") = "Y" Then
'      Cancel = True: ErrMsg = "Cannot Edit This Value": Exit Sub
'    End If
  Case Is = UCase("UoYN10")
    Cancel = IIF(ADC("UoYN3") = "N", True, False)
    ErrMsg = "Can be Used Only When Prd Wt & Prd Qty is Used": Exit Sub
  End Select
  SetProp Me, IdName, When
End Sub
Private Sub atxt_ReCalculate(Index As Integer, ByVal IdName As String, ByVal OldVal As Variant, ByVal RowNum As Integer, ByVal SourceName As String, ByVal SourceColName As String)
  Select Case UCase(IdName)
' Zubin 212 (Recalc not req for UoYN4)
'  Case Is = UCase("UoYN4")
'    If adc("UoYN2") = "Y" Then
'      adc("UoYN4") = ""
'    End If
  '************Geeta*******Emr205******
   Case Is = UCase("UOCOCDFR")
    If UCase(ADC("UoTcTypFr")) = "PL" Then
      ADC("UoCoCdFr") = ctSelfCoCd
    Else
      ADC("UoCoCdFr") = gs_CoCd
    End If
  '****************************
  Case Is = UCase("UoYN10")
    If ADC("UoYN3") = "N" Then ADC("UoYN10") = "N"
  End Select
End Sub
Private Sub CmdSc_Click()
  If FraSc.Visible = False Then
    '*** Jenny (Old Code Bef Color)
'    CmdSc.BackColor = vbWhite
'    CmdSc.FontBold = True
    '*** Jenny (Old Code Bef Color)
    FraSc.Enabled = True
    FraSc.Visible = True
    TAB_REP.Enabled = False
    ADC.AllowSave = False
    CmdDispSc.Enabled = False   ' ******** Manali 3.5.0 - 14/11/08
    FraSc.ZOrder
  Else
    '*** Jenny (Old Code Bef Color)
'    CmdSc.BackColor = &H8000000F
'    CmdSc.FontBold = False
    '*** Jenny (Old Code Bef Color)
    TAB_REP.Enabled = True
    FraSc.Enabled = False
    FraSc.Visible = False
    ADC.AllowSave = True
    CmdSc.SetFocus
    CmdDispSc.Enabled = True   ' ******** Manali 3.5.0 - 14/11/08
  End If
End Sub


' ****** Manali 3.5.0 - 07/11/08 - Radio Buttons Added
Private Sub OptDelDt_Click(Index As Integer)
Select Case Index
Case Is = 0
  ADC("UoYN5") = "Y"
Case Is = 1
  ADC("UoYN5") = "N"
End Select
End Sub
' ****** Manali 3.5.0 - 07/11/08 - Radio Buttons Added

' ****** Manali 3.5.0 - 07/11/08 - Radio Buttons Added
Private Sub OptDelDt_GotFocus(Index As Integer)
  DispMsg FraOptDelDt.ToolTipText, etInfo
End Sub
' ****** Manali 3.5.0 - 07/11/08 - Radio Buttons Added

' ***** Manali 3.10.0 - 14/03/12 - Show Net Wt instead of Gld As Wt
Private Sub OptNetWt_Click(Index As Integer)
Select Case Index
Case Is = 0
  ADC("UoYN28") = "Y"
Case Is = 1
  ADC("UoYN28") = "N"
End Select

End Sub
' ***** Manali 3.10.0 - 14/03/12 - Show Net Wt instead of Gld As Wt
Private Sub OptNetWt_GotFocus(Index As Integer)
  DispMsg FraOptNetWt.ToolTipText, etInfo
End Sub

' ****** Manali 3.5.0 - 19/11/08 - Radio Buttons Added
Private Sub OptOrdQty_Click(Index As Integer)
Select Case Index
Case Is = 0
  ADC("UoYN1") = "Y"
Case Is = 1
  ADC("UoYN1") = "N"
End Select
End Sub
' ****** Manali 3.5.0 - 19/11/08 - Radio Buttons Added

' ****** Manali 3.5.0 - 07/11/08 - Radio Buttons Added
Private Sub OptOrdQty_GotFocus(Index As Integer)
  DispMsg FraOptOrdQty.ToolTipText, etInfo
End Sub
' ****** Manali 3.5.0 - 07/11/08 - Radio Buttons Added

' ****** Manali 3.5.0 - 07/11/08 - Radio Buttons Added
Private Sub OptPtr_Click(Index As Integer)
Select Case Index
Case Is = 0
  If OptRmPrd(1).Value = True Then
    OptPtr(0).Value = False: OptPtr(1).Value = True: ADC("UoYN10") = "N"
    DispMsg "Calculated Pointer can be selected only when Prd Wt & Qty is Used", etError
  Else
    ADC("UoYN10") = "Y"
  End If
Case Is = 1
  ADC("UoYN10") = "N"
End Select
End Sub
' ****** Manali 3.5.0 - 07/11/08 - Radio Buttons Added

' ****** Manali 3.5.0 - 07/11/08 - Radio Buttons Added
Private Sub OptPtr_GotFocus(Index As Integer)
  DispMsg FraOptPtr.ToolTipText, etInfo
End Sub
' ****** Manali 3.5.0 - 07/11/08 - Radio Buttons Added

' ****** Manali 3.5.0 - 07/11/08 - Radio Buttons Added
Private Sub OptRmPrd_Click(Index As Integer)
Select Case Index
Case Is = 0
  ADC("UoYN3") = "Y"
Case Is = 1
  ADC("UoYN3") = "N"
  OptPtr(1).Value = True: ADC("UoYN10") = "N"
End Select
End Sub
' ****** Manali 3.5.0 - 07/11/08 - Radio Buttons Added

' ****** Manali 3.5.0 - 07/11/08 - Radio Buttons Added
Private Sub OptRmPrd_GotFocus(Index As Integer)
  DispMsg FraOptRm.ToolTipText, etInfo
End Sub
' ****** Manali 3.5.0 - 07/11/08 - Radio Buttons Added

Private Sub tab_REP_Click(PreviousTab As Integer)
    fra_tabrep(TAB_REP.Tab).Enabled = True
    fra_tabrep(PreviousTab).Enabled = False
End Sub
Private Sub ADC_KeyWhen()
    ADC("UOUSRCD") = ADC.UsrCd
    ADC("UOMNUCD") = ADC.MenuCd
     '*************************** ZUBIN **************************
  ' 08th Dec 2003, EMR206
  '*** Jay 3.3(FG)
  ' **** Manali Trading Module
  If UCase(ADC.MenuCd) = UCase("RepAvgOrd") Or UCase(ADC.MenuCd) = UCase("RepOrdCat") Or UCase(ADC.MenuCd) = UCase("RepFgAvg") Or UCase(ADC.MenuCd) = UCase("RepFgCat") Or _
    UCase(ADC.MenuCd) = UCase("RepJtAvg") Or UCase(ADC.MenuCd) = UCase("RepJtCat") Then
  'If UCase(ADC.MenuCd) = UCase("RepAvgOrd") Or UCase(ADC.MenuCd) = UCase("RepOrdCat") Then
  '*** Jay 3.3(FG)
    lblCustRm.Visible = False
    ADC("UoYN8").Visible = False
    ChkBoxArr(8).Enabled = False: ChkBoxArr(8).Visible = False    ' ****** Manali 3.5.0 - 19/11/08 - Check Boxes added
  End If
  '*************************** ZUBIN **************************
End Sub
Private Sub ADC_NonKeyWhen(pr_Cancel As Boolean, pr_ErrMsg As String)
    '*************Geeta*********************
    If UCase(ADC("UoTcTypFr")) = "PL" Then
      ADC("UoCoCdFr") = ctSelfCoCd
    Else
      ADC("UoCoCdFr") = gs_CoCd
    End If
    '*** Jay 3.3(FG)
    If UCase(ADC.MenuCd) = UCase("RepFgDet") Or UCase(ADC.MenuCd) = UCase("RepFgAvg") Or UCase(ADC.MenuCd) = UCase("RepFgCat") Or _
      UCase(ADC.MenuCd) = UCase("RepJtDet") Or UCase(ADC.MenuCd) = UCase("RepJtAvg") Or UCase(ADC.MenuCd) = UCase("RepJtCat") Then
      ADC("UODELDTFR") = MWLib.EmptyDate: ADC("UODELDTTo") = MWLib.EmptyDate
      ADC("UODELDT1FR") = MWLib.EmptyDate: ADC("UODELDT1To") = MWLib.EmptyDate
      ADC("UOYn1") = "Y": ADC("UOYn3") = "N": ADC("UOYn10") = "N"
      ADC("UoTcTypFr") = IIF(UCase(ADC.MenuCd) = UCase("RepJtDet") Or UCase(ADC.MenuCd) = UCase("RepJtAvg") Or UCase(ADC.MenuCd) = UCase("RepJtCat"), ctJtTcTyp, "FG")
    End If
    '*** Jay 3.3(FG)
 '*******************************
  ' adc("UoCoCdFr") = gs_CoCd
  ' Zubin 212
  '*** Jay 3.3(FG)
  If ADC("UOYN4") = "" And (UCase(ADC.MenuCd) <> UCase("RepOrdCat") And UCase(ADC.MenuCd) <> UCase("RepFgCat") And UCase(ADC.MenuCd) <> UCase("RepJtCat")) Then ADC("UOYN4") = "Y"
  'If ADC("UOYN4") = "" And UCase(ADC.MenuCd) <> UCase("RepOrdCat") Then ADC("UOYN4") = "Y"
  '*** Jay 3.3(FG)
  ' Zubin 212
  
  ' ****** Manali 3.5.0 - 07/11/08 - Radio Buttons Added
  If ADC("UoYN5") = "Y" Then OptDelDt(0).Value = True
  If ADC("UoYN5") = "N" Then OptDelDt(1).Value = True
  If ADC("UoYN3") = "Y" Then OptRmPrd(0).Value = True
  If ADC("UoYN3") = "N" Then OptRmPrd(1).Value = True
  If ADC("UoYN10") = "Y" Then OptPtr(0).Value = True
  If ADC("UoYN10") = "N" Then OptPtr(1).Value = True
  If ADC("UoYN1") = "Y" Then OptOrdQty(0).Value = True
  If ADC("UoYN1") = "N" Then OptOrdQty(1).Value = True
  ' ***** Manali 3.10.0 - 14/03/12 - Show Net Wt instead of Gld As Wt
  If ADC("UoYN28") = "Y" Then OptNetWt(0).Value = True
  If ADC("UoYN28") = "N" Then OptNetWt(1).Value = True
  ' ***** Manali 3.10.0 - 14/03/12 - Show Net Wt instead of Gld As Wt
  ' ****** Manali 3.5.0 - 07/11/08 - Radio Buttons Added
  
  
  ' ****** Manali 3.5.0 - 19/11/08 - Check Boxes added
  If ADC("UoYN") = "" Then ADC("UoYN") = "N"
  If ADC("UoYN2") = "" Then ADC("UoYN2") = "N"
  'If adc("UoYN4") = "" Then adc("UoYN4") = "N"
  If ADC("UoYN6") = "" Then ADC("UoYN6") = "N"
  If ADC("UoYN7") = "" Then ADC("UoYN7") = "N"
  If ADC("UoYN8") = "" Then ADC("UoYN8") = "N"
  If ADC("UoYN9") = "" Then ADC("UoYN9") = "N"
  If ADC("UoYN11") = "" Then ADC("UoYN11") = "N"
  If ADC("UoYN12") = "" Then ADC("UoYN12") = "N"
  If ADC("UoYN13") = "" Then ADC("UoYN13") = "N"
  If ADC("UoYN14") = "" Then ADC("UoYN14") = "N"
  If ADC("UoYN15") = "" Then ADC("UoYN15") = "N"
  If ADC("UoYN16") = "" Then ADC("UoYN16") = "N"
  If ADC("UoYN17") = "" Then ADC("UoYN17") = "N"
  If ADC("UoYN18") = "" Then ADC("UoYN18") = "N"
  If ADC("UoYN19") = "" Then ADC("UoYN19") = "N"
  If ADC("UoYN20") = "" Then ADC("UoYN20") = "N"
  If ADC("UoYN21") = "" Then ADC("UoYN21") = "N"
  If ADC("UoYN22") = "" Then ADC("UoYN22") = "N"
  ' ***** Manali 3.6.0 - Order Printing in Excel Format Module
  'If adc("UoYN23") = "" Or UCase(adc.MenuCd) <> UCase("RepDetOrd") Then adc("UoYN23") = "N"    ' ***** Manali 3.6.0 - 24/09/09 - Show Report in Excel Format
  If moCn.RecSeek("Select 'x' from Head Where HCoCd='" + ctSelfCoCd + "' and HCd='" + ctSelfCmCd + "' and HOrdExcelRepYn<>'Y' ") Or _
    ADC("UoYN23") = "" Or UCase(ADC.MenuCd) <> UCase("RepDetOrd") Then ADC("UoYN23") = "N"
  ' ***** Manali 3.6.0 - Order Printing in Excel Format Module
  If ADC("UoYN24") = "" Then ADC("UoYN24") = "N"
  If ADC("UoYN25") = "" Then ADC("UoYN25") = "N"      ''''' Manali 3.8.0
  If ADC("UoYN26") = "" Then ADC("UoYN26") = "N"      ''''' Manali 3.9.0 Design Level Remarks
  If ADC("UoYN27") = "" Then ADC("UoYN27") = "N"      ' ***** Manali 3.10.0 - 09/03/12 - Different Suppres option for Unit Pricce and Qty
  If ADC("UoYN30") = "" Then ADC("UoYN30") = "N"      ' 4.1.2

  If ADC("UoYN") = "Y" Then ChkBoxArr(0).Value = Checked Else ChkBoxArr(0).Value = Unchecked
  If ADC("UoYN2") = "Y" Then ChkBoxArr(2).Value = Checked Else ChkBoxArr(2).Value = Unchecked
  If ADC("UoYN4") = "Y" Then ChkBoxArr(4).Value = Checked Else ChkBoxArr(4).Value = Unchecked
  If ADC("UoYN6") = "Y" Then ChkBoxArr(6).Value = Checked Else ChkBoxArr(6).Value = Unchecked
  If ADC("UoYN7") = "Y" Then ChkBoxArr(7).Value = Checked Else ChkBoxArr(7).Value = Unchecked
  If ADC("UoYN8") = "Y" Then ChkBoxArr(8).Value = Checked Else ChkBoxArr(8).Value = Unchecked
  If ADC("UoYN9") = "Y" Then ChkBoxArr(9).Value = Checked Else ChkBoxArr(9).Value = Unchecked
  If ADC("UoYN11") = "Y" Then ChkBoxArr(11).Value = Checked Else ChkBoxArr(11).Value = Unchecked
  If ADC("UoYN12") = "Y" Then ChkBoxArr(12).Value = Checked Else ChkBoxArr(12).Value = Unchecked
  If ADC("UoYN13") = "Y" Then ChkBoxArr(13).Value = Checked Else ChkBoxArr(13).Value = Unchecked
  If ADC("UoYN14") = "Y" Then ChkBoxArr(14).Value = Checked Else ChkBoxArr(14).Value = Unchecked
  If ADC("UoYN15") = "Y" Then ChkBoxArr(15).Value = Checked Else ChkBoxArr(15).Value = Unchecked
  If ADC("UoYN16") = "Y" Then ChkBoxArr(16).Value = Checked Else ChkBoxArr(16).Value = Unchecked
  If ADC("UoYN17") = "Y" Then ChkBoxArr(17).Value = Checked Else ChkBoxArr(17).Value = Unchecked
  If ADC("UoYN18") = "Y" Then ChkBoxArr(18).Value = Checked Else ChkBoxArr(18).Value = Unchecked
  If ADC("UoYN19") = "Y" Then ChkBoxArr(19).Value = Checked Else ChkBoxArr(19).Value = Unchecked
  If ADC("UoYN20") = "Y" Then ChkBoxArr(20).Value = Checked Else ChkBoxArr(20).Value = Unchecked
  If ADC("UoYN21") = "Y" Then ChkBoxArr(21).Value = Checked Else ChkBoxArr(21).Value = Unchecked
  If ADC("UoYN22") = "Y" Then ChkBoxArr(22).Value = Checked Else ChkBoxArr(22).Value = Unchecked
  If ADC("UoYN23") = "Y" Then ChkBoxArr(23).Value = Checked Else ChkBoxArr(23).Value = Unchecked      ' ***** Manali 3.6.0 - 24/09/09 - Show Report in Excel Format
  If ADC("UoYN24") = "Y" Then ChkBoxArr(24).Value = Checked Else ChkBoxArr(24).Value = Unchecked
  If ADC("UoYN25") = "Y" Then ChkBoxArr(25).Value = Checked Else ChkBoxArr(25).Value = Unchecked    ''''' Manali 3.8.0
  If ADC("UoYN26") = "Y" Then ChkBoxArr(26).Value = Checked Else ChkBoxArr(26).Value = Unchecked    ' ***** Manali 3.9.0 - Design Level Remarks
  If ADC("UoYN27") = "Y" Then ChkBoxArr(27).Value = Checked Else ChkBoxArr(27).Value = Unchecked    ' ***** Manali 3.10.0 - 09/03/12 - Different Suppres option for Unit Pricce and Qty
  If ADC("UoYN30") = "Y" Then ChkBoxArr(30).Value = Checked Else ChkBoxArr(30).Value = Unchecked    ' 4.1.2
  ' ****** Manali 3.5.0 - 19/11/08 - Check Boxes added
  
  'adc("UoPMCdFr") = "3D"
  
  
  ' ***** Manali 3.9.0 - Lme Selection
  Dim i As Integer
  For i = 0 To 3
    If ADC("UoLme1") = ChkBoxLme(i).Caption Or ADC("UoLme2") = ChkBoxLme(i).Caption Then
      ChkBoxLme(i).Value = Checked
    End If
  Next i
  ' ***** Manali 3.9.0 - Lme Selection
End Sub
Private Sub DispCoNm()
'mRep.wCoCd.SetText gs_CoNm
    Dim ws_HName As String
    ws_HName = moCn.GetFldVal("Select HName from Head where HCoCd = '" + ADC("UoCoCdFr") + "' and " + _
                                                      "HCd = '" + ctSelfCmCd + "' ")
    
    If ADC("UoTcTypFr") = "PL" Then
      mRep.wCoCd.SetText gs_CoNm
      mRep.wCoCdLogo.SetText gs_CoNm
    Else
      mRep.wCoCd.SetText ws_HName
      mRep.wCoCdLogo.SetText ws_HName
    End If
        
  
  ' ***** Manali 3.8.0 - Comp Logo in Report
  If moCn.RecSeek("Select 'x' From Head Where HCoCd='" + gs_CoCd + "' and HCd='" + ctSelfCmCd + "' and HShowRepLogo='Y' ") Then
    mRep.RHMain.UnderlaySection = True: mRep.wCoCd.Suppress = True
    If UCase(ADC.MenuCd) = UCase("RepJtDet") Or UCase(ADC.MenuCd) = UCase("RepJtAvg") Then
      Rep.RHLogo.UnderlaySection = True
    End If
  Else
    mRep.RHLogo.Suppress = True
  End If
    mRep.RHLogo.Suppress = True
  ' ***** Manali 3.8.0 - Comp Logo in Report
  
' ********* Manali 3.2.2 - Before 3.2.2
'  Dim ws_DtTm As String
'  ws_DtTm = Format$(moCn.SrvrDate, "DD-MMM-YYYY")
'  mRep.wSrvrDate.SetText ws_DtTm
'  ws_DtTm = Replace(Format(IIF(CDbl(moCn.SrvrTime) >= 13#, moCn.SrvrTime - 12, IIF(CDbl(moCn.SrvrTime) < 1#, moCn.SrvrTime + 12, moCn.SrvrTime)), "00.00"), ".", ":") + IIF(CDbl(moCn.SrvrTime) >= 12#, " PM", " AM")
'  mRep.wSrvrTime.SetText ws_DtTm

' ******** Manali 3.2.2 - User date and time in reports
  Dim ws_Dt As String, ws_Tm As String
    ws_Dt = Format$(moCn.SrvrDate, "DD-MMM-YYYY")
    ws_Tm = Replace(Format(IIF(CDbl(moCn.SrvrTime) >= 13#, moCn.SrvrTime - 12, IIF(CDbl(moCn.SrvrTime) < 1#, moCn.SrvrTime + 12, moCn.SrvrTime)), "00.00"), ".", ":") + IIF(CDbl(moCn.SrvrTime) >= 12#, " PM", " AM")
    mRep.TxtRepUsrDtTime.SetText Trim(ADC.UsrCd) + " / " + ws_Dt + " / " + ws_Tm

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
    Cancel = atxt(Index).Validate
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode <> vbFormCode Then Cancel = True: Exit Sub
    
    '*** (09/08/05)
    Set moCn = Nothing   '*** (Bef speed) bef 24/11/06
    Set mRep = Nothing
    Set Rep = Nothing
    Set Rep1 = Nothing
    '*** (09/08/05)
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

Private Sub SetExcel(ByVal pv_rsOrd As MDORowSet)

On Error GoTo Err_Hndlr

Dim sFile, sFilePath As String
Dim wOdSr As Integer
Dim wOdRow As Integer, wOrMRow As Integer, wOrDCRow As Integer, wOrXRow As Integer, wOlRow As Integer, wOmRow As Integer, wLn1 As Integer, wLn2 As Integer
Dim wCoName As String, wSht As Integer
Dim wPictPath As String
Dim wMetLs As Single, wGmFactor As Double, wCmPurityRt As Double
Dim wGldAs As String, wGldAsWt As String
Dim wCurrRow As Integer   ' ***** Manali 3.8.0
Dim wOrMFRow As Integer, wOrDCFRow As Integer, wOrXfRow As Integer, wOlFRow As Integer
Dim wLabAs As String, wLabAsWt As String
Dim wShtNm As String

' ***** Manali 3.9.0 - Excel 2007
Dim wExcVer As String, wExtn As String

  Set oExcel = GetExcelObj()                            'Sachin - 4.4.2
  'Set oExcel = CreateObject("Excel.Application.12")
  'Set oExcel = CreateObject("Excel.Application")
  
  ' ***** Manali 3.9.0 - Excel 2007
  wExcVer = oExcel.version
  wExtn = IIF(wExcVer >= "12.0", ".xlsx", ".xls")
  sFile = gs_CoCd + "-" + gs_UsrCd + "-" + Format(Now, "ddMMyyHHmmss") + wExtn
  ' ***** Manali 3.9.0 - Excel 2007
  
  'Bef 3.9.0
  'sFile = gs_CoCd + "-" + gs_UsrCd + "-" + Format(Now, "ddMMyyHHmmss") + ".xls"
  'sFilePath = App.Path + "\Output\" + sFile
  
  ' ***** Manali 3.9.0 - Output folder Path From Param
  Dim wOutputPath As String
  
  wOutputPath = moCn.GetFldVal("Select PDesc225 from Param Where PTyp='OUTPUT' and PMCd='OUTPUT' ")
  sFilePath = wOutputPath + "\" + sFile
  ' ***** Manali 3.9.0 - Output folder Path From Param
  
  If Dir(sFilePath, vbNormal) <> "" Then
    Kill sFilePath
  End If
  
  ' ***** Manali 3.9.0 - Excel 2007
  FileCopy App.Path + "\Template\OrderPrint" + wExtn, sFilePath
  ' bef  3.9.0
  'FileCopy App.Path + "\Template\OrderPrint.xls", sFilePath
  
  oExcel.Visible = False: oExcel.DisplayAlerts = False
  
  wCoName = moCn.GetFldVal("Select HName from Head Where HCoCd='" + gs_CoCd + "' and HCd='" + ctSelfCmCd + "'")
  
  wOdSr = 0: wOdRow = 14: wOmRow = 1: wSht = 1
    
  '4.1.0.0 - Earlier
  'Set oBook = oExcel.Workbooks.Open(sFilePath)
  
  'Sachin 4.1.0.0 - Opening the Template file with Predefined Password
  Set oBook = oExcel.Workbooks.Open(sFilePath, Password:=ctXCelPwd, WriteResPassword:=ctXCelWritePwd)

  
  With pv_rsOrd
    Do While Not (.EOF Or .BOF)
      If wOdRow <= 285 Then
        If !OdSr <> wOdSr Then
        
          'wiSht = oExcel.Worksheets.Count
          'oExcel.Worksheets(2).Copy after:=oExcel.ActiveSheet
          oExcel.worksheets(2).Copy after:=oBook.worksheets(wSht + 1)
    
          Set oSheet = oBook.worksheets(1)
          
          If wOmRow = 1 Then
            oSheet.Name = !OmCoCd + !OmYy + !OmTc + !OmChr + CStr(!OmNo)
            oSheet.Cells(1, 1) = wCoName
            oSheet.Cells(1, 7) = "DETAIL ORDER PRINTING"
            'oSheet.Cells(1, 20) = Format(Date, "DD/MM/YY")

            oSheet.Cells(5, 3) = !OmCoCd + "/" + !OmYy + "/" + !OmTc + "/" + !OmChr + "/" + CStr(!OmNo)
            oSheet.Cells(5, 6) = !OmDt
            oSheet.Cells(5, 9) = !OmLmgSal
            oSheet.Cells(7, 9) = !OmLmsSal       'oSheet.Cells(5, 13)
            oSheet.Cells(6, 9) = !OmLmpSal
            oSheet.Cells(8, 9) = !OmLmlSal       'oSheet.Cells(6, 13)
            oSheet.Cells(6, 3) = !OmCmCd
            oSheet.Cells(7, 3) = !CmName
            oSheet.Cells(8, 3) = !OmPoNo
            oSheet.Cells(8, 6) = !OmPoDt
            oSheet.Cells(9, 6) = !OmDelDt
            oSheet.Cells(10, 1) = "Total (" + IIF(wb_InBaseCur, ws_BaseCurCd, ADC("UoCurCdFr")) + ")"
            oSheet.Cells(13, 18) = "Prc (" + IIF(wb_InBaseCur, ws_BaseCurCd, ADC("UoCurCdFr")) + ")"      'oSheet.Cells(13, 8)
            oSheet.Cells(13, 19) = "Val (" + IIF(wb_InBaseCur, ws_BaseCurCd, ADC("UoCurCdFr")) + ")"      'oSheet.Cells(13, 9)
            
            If ADC("UoYn6") = "Y" Then
              oSheet.Cells(315, 1) = !OmCmDelTerms
              oSheet.Cells(320, 1) = !OmCmPayTerms
              oSheet.Cells(325, 1) = !OmSalRem
            Else
              oSheet.Range("A315", "D328").UnMerge
              oSheet.Range("A314", "D328") = ""
            End If
          End If
          oSheet.Cells(wOdRow, 1) = IIF(ADC("UoYn25") = "Y", wOdRow - 13, !OdSr)
          oSheet.Cells(wOdRow, 2) = !OdDmCd
          oSheet.Cells(wOdRow, 4) = !OdSfx
          oSheet.Cells(wOdRow, 6) = !OdDmSz
          oSheet.Cells(wOdRow, 17) = !qOrdQty               'oSheet.Cells(wOdRow, 7)
          oSheet.Cells(wOdRow, 18) = !OdSalPrc              'oSheet.Cells(wOdRow, 8)
          oSheet.Cells(wOdRow, 19).Font.Color = vbBlack     'oSheet.Cells(wOdRow, 9)
          
          wOdRow = wOdRow + 1
          wOmRow = wOmRow + 1
          wSht = wSht + 1
        End If
        
        'Set oSheet = oBook.Worksheets(Worksheets.Count)
        Set oSheet = oBook.worksheets(wSht + 1)
        
        If !OdSr <> wOdSr Then
          ' ***** Manali 3.9.0 - DC Wt, Grs Wt in 1st sheet
          wShtNm = !OdDmCd + "-" + "(" + CStr(IIF(ADC("UoYn25") = "Y", wOdRow - 14, !OdSr)) + ")"
          oSheet.Name = wShtNm
          Set oSheet = oBook.worksheets(1)
          'oSheet.Cells(wOdRow - 1, 7) = "='" + wShtNm + "'!D43+'" + wShtNm + "'!D44"
          'oSheet.Cells(wOdRow - 1, 8) = "='" + wShtNm + "'!D47"
          oSheet.Cells(wOdRow - 1, 13) = "='" + wShtNm + "'!D43"     '+'" + wShtm + "'!D44" Diamond Wt
          oSheet.Cells(wOdRow - 1, 14) = "='" + wShtNm + "'!D44"     '+'" + wShtNm + "'!D44" Color Stone Wt
          oSheet.Cells(wOdRow - 1, 16) = "='" + wShtNm + "'!D47"     ' Gross Wt
          
          '***4.1.0.0***
          oSheet.Cells(wOdRow - 1, 11) = "='" + wShtNm + "'!C43"                        'DiaQty
          oSheet.Cells(wOdRow - 1, 12) = "='" + wShtNm + "'!C44"                        'CSQty
          oSheet.Cells(wOdRow - 1, 9) = "='" + wShtNm + "'!M14"                         'OdKT
          oSheet.Cells(wOdRow - 1, 10) = "='" + wShtNm + "'!O14"                        'OdCol
          oSheet.Cells(wOdRow - 1, 7) = "='" + wShtNm + "'!C7"                          'CustDmCd
          oSheet.Cells(wOdRow - 1, 15) = "='" + wShtNm + "'!D45+'" + wShtNm + "'!D42"   'NetWt
          oSheet.Cells(wOdRow - 1, 8) = "='" + wShtNm + "'!C9"                          'OdSubPO
          oSheet.Cells(wOdRow - 1, 20) = moCn.GetFldVal("Select  dbo.MwFn_GetColStnStr('" + !OmCoCd + "','" + !OmTc + "','" + !OmYy + "','" + !OmChr + "'," + CStr(!OmNo) + "," + CStr(!OdSr) + ", '','" + ADC("UOYN8") + "')") 'RmCd Str
          
          Set oSheet = oBook.worksheets(wSht + 1)
          ' ***** Manali 3.9.0 - DC Wt, Grs Wt in 1st sheet
          
          'oSheet.Name = !OdDmCd + "-" + "(" + CStr(IIF(adc("UoYn25") = "Y", wOdRow - 14, !OdSr)) + ")"    '!OdSr
          oSheet.Cells(1, 1) = wCoName
          oSheet.Cells(1, 5) = "DETAIL ORDER PRINTING"
          oSheet.Cells(1, 14) = Format(Date, "DD/MM/YY")
          oSheet.Cells(4, 3) = !CmName
          oSheet.Cells(5, 3) = !OmCoCd + "/" + !OmYy + "/" + !OmTc + "/" + !OmChr + "/" + CStr(!OmNo)
          oSheet.Cells(5, 6) = !OmDt
          oSheet.Cells(6, 3) = !OdDmCd
          oSheet.Cells(6, 6) = !OdSfx
          oSheet.Cells(6, 9) = !OdDmSz
          oSheet.Cells(7, 3) = !CdOurDmCd
          oSheet.Cells(13, 13) = !OdDmCd
          oSheet.Cells(14, 13) = !IdKt
          oSheet.Cells(14, 15) = !OdDmCol
          oSheet.Cells(6, 13) = !OdPicNm
          oSheet.Cells(8, 3) = !qOrdQty
          oSheet.Cells(9, 3) = IIF(!OdPoNo <> "", !OmPoNo + "/" + !OdPoNo, !OmPoNo)   'Sangeeta 3.11.0
          oSheet.Cells(10, 3) = !OdDelDt
          oSheet.Cells(11, 7) = !OdSalPrc
          'oSheet.Cells(34, 5) = !OdLabAsWt
          oSheet.Cells(43, 10) = !OdMulBy
          
          wLabAs = moCn.GetFldVal("Select PDesc225 from Param Where Ptyp='GLDAS' and PMCd='" + !OdLabAs + "'")
          wLabAsWt = ""
          If InStr(1, "," + wLabAs + ",", ",G,") <> 0 Then wLabAsWt = wLabAsWt + IIF(wLabAsWt <> "", "+", "") + "SUMIF(B18:B19,""G"", F18:F19)"
          If InStr(1, "," + wLabAs + ",", ",P,") <> 0 Then wLabAsWt = wLabAsWt + IIF(wLabAsWt <> "", "+", "") + "SUMIF(B18:B19,""P"", F18:F19)"
          If InStr(1, "," + wLabAs + ",", ",S,") <> 0 Then wLabAsWt = wLabAsWt + IIF(wLabAsWt <> "", "+", "") + "SUMIF(B18:B19,""S"", F18:F19)"
          If InStr(1, "," + wLabAs + ",", ",L,") <> 0 Then wLabAsWt = wLabAsWt + IIF(wLabAsWt <> "", "+", "") + "SUMIF(B18:B19,""L"", F18:F19)"
          If InStr(1, "," + wLabAs + ",", ",D,") <> 0 Then wLabAsWt = wLabAsWt + IIF(wLabAsWt <> "", "+", "") + "(D43/5)"
          If InStr(1, "," + wLabAs + ",", ",C,") <> 0 Then wLabAsWt = wLabAsWt + IIF(wLabAsWt <> "", "+", "") + "(D44/5)"
          If InStr(1, "," + wLabAs + ",", ",X,") <> 0 Then wLabAsWt = wLabAsWt + IIF(wLabAsWt <> "", "+", "") + "(D45)"
          wLabAsWt = "=(" + wLabAsWt + ")"
          oSheet.Cells(34, 5) = wLabAsWt    '!OdLabAsWt
          
          wCurrRow = 18
          'wOrMRow = 18: wOrDCRow = 24: wOrXRow = 30: wOlRow = 36
          'wOrMRow = 18: wOrDCRow = 26: wOrXRow = 45: wOlRow = 53
          
          oSheet.Cells(11, 1) = "Price (" + IIF(wb_InBaseCur, ws_BaseCurCd, ADC("UoCurCdFr")) + ")"
          oSheet.Cells(11, 5) = "Sales Price (" + IIF(wb_InBaseCur, ws_BaseCurCd, ADC("UoCurCdFr")) + ")"
          oSheet.Cells(17, 8) = "LME (" + IIF(wb_InBaseCur, ws_BaseCurCd, ADC("UoCurCdFr")) + ")"
          oSheet.Cells(17, 9) = "Rate (" + IIF(wb_InBaseCur, ws_BaseCurCd, ADC("UoCurCdFr")) + ")"
          oSheet.Cells(17, 10) = "Value (" + IIF(wb_InBaseCur, ws_BaseCurCd, ADC("UoCurCdFr")) + ")"
          oSheet.Cells(23, 10) = "Rate (" + IIF(wb_InBaseCur, ws_BaseCurCd, ADC("UoCurCdFr")) + ")"
          oSheet.Cells(23, 11) = "Value (" + IIF(wb_InBaseCur, ws_BaseCurCd, ADC("UoCurCdFr")) + ")"
          oSheet.Cells(23, 13) = "Rate (" + IIF(wb_InBaseCur, ws_BaseCurCd, ADC("UoCurCdFr")) + ")"
          oSheet.Cells(23, 14) = "Value (" + IIF(wb_InBaseCur, ws_BaseCurCd, ADC("UoCurCdFr")) + ")"
          oSheet.Cells(29, 8) = "Rate (" + IIF(wb_InBaseCur, ws_BaseCurCd, ADC("UoCurCdFr")) + ")"
          oSheet.Cells(29, 9) = "Value (" + IIF(wb_InBaseCur, ws_BaseCurCd, ADC("UoCurCdFr")) + ")"
          oSheet.Cells(35, 6) = "Rate (" + IIF(wb_InBaseCur, ws_BaseCurCd, ADC("UoCurCdFr")) + ")"
          oSheet.Cells(35, 7) = "Value (" + IIF(wb_InBaseCur, ws_BaseCurCd, ADC("UoCurCdFr")) + ")"
          oSheet.Cells(41, 5) = "Value (" + IIF(wb_InBaseCur, ws_BaseCurCd, ADC("UoCurCdFr")) + ")"
          oSheet.Cells(41, 10) = "Value (" + IIF(wb_InBaseCur, ws_BaseCurCd, ADC("UoCurCdFr")) + ")"
          
          ' **** Manali 3.8.0
          wOrMRow = 1: wOrDCRow = 1: wOrXRow = 1: wOlRow = 1
          
          If ADC("UoYn") = "Y" Then
            ' **** Manali 3.6.1 - Pic Reference and Dm/SM
            wPictPath = GetPictPath(IIF(UCase(ADC("UoTcTypFr")) = "QS", "SM", "DM"), IIF(UCase(ADC("UoPMCdFr")) = "3D", False, True), !OdDmCd, !OdDmSz, !OmCoCd, !OmTc, !OmYy, !OmChr, !OmNo, !OdSr, !OdPicNm)
            If wPictPath = "" Or Dir(wPictPath) = "" Then
              wPictPath = GetPictPath(IIF(UCase(ADC("UoTcTypFr")) = "QS", "SM", "DM"), IIF(UCase(ADC("UoPMCdFr")) = "3D", True, False), !OdDmCd, !OdDmSz, !OmCoCd, !OmTc, !OmYy, !OmChr, !OmNo, !OdSr, !OdPicNm)
            End If
            'wPictPath = GetPictPath("DM", False, !OdDmCd, !OdDmSz, !OmCoCd, !OmTc, !OmYy, !OmChr, !OmNo, !OdSr, !OdPicNm)
            If Dir(wPictPath) <> "" And wPictPath <> "" Then
              ' ***** Manali 3.9.0 - Change Done For Excel 2010 to embed image
              ' ***** Help - expression.AddPicture(FileName, LinkToFile, SaveWithDocument, Left, Top, Width, Height)
              oSheet.Shapes.AddPicture wPictPath, False, True, 577, 45, 140, 110
              oSheet.Pictures(1).ShapeRange.LockAspectRatio = False
              oSheet.Pictures(1).Name = oSheet.Name
              oSheet.Pictures(1).Border.LineStyle = 1
              
              ' Bef 3.9.0
              'oSheet.Pictures.Insert(wPictPath).Select
              'oSheet.Pictures(1).Left = 577: oSheet.Pictures(1).Top = 45
              'oSheet.Pictures(1).ShapeRange.LockAspectRatio = False
            End If
          End If
          
          ' bef  3.9.0
          'If adc("UoYn6") = "Y" Then
          ' ***** Manali 3.9.0 - Different Option for Design Level Remrks
          If ADC("UoYn26") = "Y" Then
            oSheet.Cells(52, 1) = !OdDmPrdInst
            oSheet.Cells(59, 1) = !OdCmPrdInst
            oSheet.Cells(66, 1) = !OdCmStmpInst
            oSheet.Cells(71, 1) = !OdSzInst
            oSheet.Cells(76, 1) = !OdSalRem
          Else
            oSheet.Range("A51", "G79").UnMerge
            oSheet.Range("A51", "G79") = ""
          End If
        End If
          
          If !qTag = "1RM" Then
            Select Case !RmCtg
            Case Is = "G", "P", "S", "L"
              ' ****** Manali 3.8.0 -
              'wOrMRow = wOrMRow + 1
              If wOrMRow = 1 Then
                wOrMFRow = wCurrRow
              Else
                oSheet.Rows(CStr(wCurrRow) + ":" + CStr(wCurrRow)).Insert
                oSheet.Range("A" + CStr(wCurrRow - 1) + ":J" + CStr(wCurrRow - 1)).Copy
                oSheet.Range("A" + CStr(wCurrRow) + ":J" + CStr(wCurrRow)).Select
                oSheet.Paste
              End If
              
              'If wOrMRow <= 21 Then
                oSheet.Cells(wCurrRow, 1) = !OrSrNo
                oSheet.Cells(wCurrRow, 2) = !RmCtg
                oSheet.Cells(wCurrRow, 3) = !OrRmSCtg
                oSheet.Cells(wCurrRow, 4) = !OrRmCd
                oSheet.Cells(wCurrRow, 5) = !OrQty
                oSheet.Cells(wCurrRow, 6) = !OrWt
                oSheet.Cells(wCurrRow, 7) = !RmQw
                oSheet.Cells(wCurrRow, 8) = !OrLMESal
                
                wMetLs = moCn.GetFldVal("Select  Top 1 RrSalRt from RmRt " + _
                          "where RrTcTyp='LS' and RrCtg='" + !RmCtg + "' and RrCmCurCd='' and RrDmCd='' and RrSTWGrp='' and " + _
                          "((RrCmCtg = 'P' and RrCmCd = '" + !CmLkUpMetLs + "') Or (RrCmCtg='" + gs_CmCtg + "' and RrCmCd in ('" + !OmCmCd + "', '" + ctSelfCmCd + "'))) " + _
                          "and ((RrDmCtg = '" + !DmCtg + "' and RrLsCtg In ('" + !DmLsCtg + "', '')) Or (RrDmCtg='' and RrLsCtg='')) " + _
                          "Order By (Case When (RrCmCtg='P' and RrCmCd = '" + !CmLkUpMetLs + "') Then 1 " + _
                          "When RrCmCtg='" + gs_CmCtg + "' Then (Case RrCmCd When '" + !OmCmCd + "' Then 2 When '" + ctSelfCmCd + "' Then 3  End)  End), " + _
                          "(Case RrDmCtg When '" + !DmCtg + "' Then (Case RrLsCtg When '" + !DmLsCtg + "' Then 1 When '' Then 2  End) When '' Then 3  End) ")
  
                
                wGmFactor = moCn.GetFldVal("Select PNum from Param Where PTyp='GMFACTOR' and PMCd='" + !RmCtg + "' and PSCd='" + IIF(wb_InBaseCur, ws_BaseCurCd, ADC("UoCurCdFr")) + "'")
                wCmPurityRt = moCn.GetFldVal("Select RrSalRt from RmRt Where RrTcTyp='CRP' and RrCmCtg='" + gs_CmCtg + "' and RrCmCd='" + !OmCmCd + "' and RrCd='" + !OrRmCd + "' ")
                wCmPurityRt = IIF(wCmPurityRt = 0, !RmPurityRt, wCmPurityRt)
                
                oSheet.Cells(wCurrRow, 9).formula = "=(H" + CStr(wCurrRow) + "/" + CStr(wGmFactor) + "*" + CStr(wCmPurityRt) + "*(1 + " + CStr(wMetLs) + "/100)) + " + _
                                                "((1-" + CStr(wCmPurityRt) + ")*" + CStr(!OrAlySalRt) + ")"
  
                ' ***** Manali 3.6.1 - GldAs Wt
                wGldAsWt = "$F" + CStr(wCurrRow)
                If !OrMainMet = "Y" Then
                  wGldAs = moCn.GetFldVal("Select PDesc225 from Param Where Ptyp='GLDAS' and PMCd='" + !OdGldAs + "'")
                  
                  If InStr(1, "," + wGldAs + ",", ",D,") <> 0 Then wGldAsWt = wGldAsWt + IIF(wGldAsWt <> "", "+", "") + "(D43/5)"
                  If InStr(1, "," + wGldAs + ",", ",C,") <> 0 Then wGldAsWt = wGldAsWt + IIF(wGldAsWt <> "", "+", "") + "(D44/5)"
                  If InStr(1, "," + wGldAs + ",", ",X,") <> 0 Then wGldAsWt = wGldAsWt + IIF(wGldAsWt <> "", "+", "") + "(D45)"
                  
                End If
                wGldAsWt = "(" + wGldAsWt + ")"
                                        
                oSheet.Cells(wCurrRow, 10).formula = "=$I" + CStr(wCurrRow) + "*(IF($G" + CStr(wCurrRow) + "=""Q"",$E" + CStr(wCurrRow) + ",(IF($G" + CStr(wCurrRow) + "=""W""," + wGldAsWt + ",0))))"
                oSheet.Cells(wCurrRow, 10).Font.Color = vbBlack
                
                
              ' ***** Bef 3.8.0
              'ElseIf wOrMRow = 22 Then
              '  oSheet.Cells(16, 1).AddComment "#Metal Details InComplete"
              'End If
              wOrMRow = wOrMRow + 1
              
            Case Is = "D", "C"
              'wOrDCRow = wOrDCRow + 1
                If wOrDCRow = 1 Then
                  wCurrRow = wCurrRow + IIF(wOrMRow = 1, 6, IIF(wOrMRow = 2, 5, 5))
                  wOrDCFRow = wCurrRow
                Else  'If wOrDCRow <> 2 Then
                  oSheet.Rows(CStr(wCurrRow) + ":" + CStr(wCurrRow)).Insert
                  oSheet.Range("A" + CStr(wCurrRow - 1) + ":N" + CStr(wCurrRow - 1)).Copy
                  oSheet.Range("A" + CStr(wCurrRow) + ":N" + CStr(wCurrRow)).Select
                  oSheet.Paste
                End If
                
'              If wOrDCRow <= 40 Then
                oSheet.Cells(wCurrRow, 1) = !OrSrNo
                oSheet.Cells(wCurrRow, 2) = !RmCtg
                oSheet.Cells(wCurrRow, 3) = !OrRmSCtg
                oSheet.Cells(wCurrRow, 4) = !OrRmCd
                oSheet.Cells(wCurrRow, 5) = !OrLn1
                oSheet.Cells(wCurrRow, 6) = !OrRmPtr
                oSheet.Cells(wCurrRow, 7) = !OrQty
                oSheet.Cells(wCurrRow, 8) = !OrWt
                oSheet.Cells(wCurrRow, 9) = !RmQw
                oSheet.Cells(wCurrRow, 10) = !OrSalRt
                oSheet.Cells(wCurrRow, 11).formula = "=$J" + CStr(wCurrRow) + "*(IF($I" + CStr(wCurrRow) + "=""Q"",$G" + CStr(wCurrRow) + ",(IF($I" + CStr(wCurrRow) + "=""W"",$H" + CStr(wCurrRow) + ",0))))"
                oSheet.Cells(wCurrRow, 11).Font.Color = vbBlack
                oSheet.Cells(wCurrRow, 12) = !OrSetSCd
                oSheet.Cells(wCurrRow, 13) = !OrSetSalRt
                oSheet.Cells(wCurrRow, 14).formula = "=$M" + CStr(wCurrRow) + "*$G" + CStr(wCurrRow) + ""
                oSheet.Cells(wCurrRow, 14).Font.Color = vbBlack
                
                
              wOrDCRow = wOrDCRow + 1
            
            Case Is = "X"
                If wOrXRow = 1 Then
                  wCurrRow = wCurrRow + IIF(wOrDCRow = 1, 6 + IIF(wOrMRow = 1, 6, IIF(wOrMRow = 2, 5, 5)), IIF(wOrDCRow = 2, 5, 5))
                  wOrXfRow = wCurrRow
                Else    'If wOrXRow <> 2 Then
                  oSheet.Rows(CStr(wCurrRow) + ":" + CStr(wCurrRow)).Insert
                  oSheet.Range("A" + CStr(wCurrRow - 1) + ":I" + CStr(wCurrRow - 1)).Copy
                  oSheet.Range("A" + CStr(wCurrRow) + ":I" + CStr(wCurrRow)).Select
                  oSheet.Paste
                End If
'              If wOrXRow <= 48 Then
                oSheet.Cells(wCurrRow, 1) = !OrSrNo
                oSheet.Cells(wCurrRow, 2) = !RmCtg
                oSheet.Cells(wCurrRow, 3) = !OrRmSCtg
                oSheet.Cells(wCurrRow, 4) = !OrRmCd
                oSheet.Cells(wCurrRow, 5) = !OrQty
                oSheet.Cells(wCurrRow, 6) = !OrWt
                oSheet.Cells(wCurrRow, 7) = !RmQw
                oSheet.Cells(wCurrRow, 8) = !OrSalRt
                oSheet.Cells(wCurrRow, 9).formula = "=$H" + CStr(wCurrRow) + "*(IF($G" + CStr(wCurrRow) + "=""Q"",$E" + CStr(wCurrRow) + ",(IF($G" + CStr(wCurrRow) + "=""W"",$F" + CStr(wCurrRow) + ",0))))"
                oSheet.Cells(wCurrRow, 9).Font.Color = vbBlack
                
                
                
'              ElseIf wOrXRow = 49 Then
'                oSheet.Cells(43, 1).AddComment = "#Acc Details InComplete"
'              End If
              wOrXRow = wOrXRow + 1
            End Select
          
          ElseIf !qTag = "2LAB" Then
                If wOlRow = 1 Then
                  wCurrRow = wCurrRow + IIF(wOrXRow = 1, 6 + IIF(wOrDCRow = 1, 6 + IIF(wOrMRow = 1, 6, IIF(wOrMRow = 2, 5, 5)), IIF(wOrDCRow = 2, 5, 5)), _
                                        IIF(wOrXRow = 2, 5, 5))
                  wOlFRow = wCurrRow
                Else    'If wOlRow <> 2 Then
                  oSheet.Rows(CStr(wCurrRow) + ":" + CStr(wCurrRow)).Insert
                  oSheet.Range("A" + CStr(wCurrRow - 1) + ":G" + CStr(wCurrRow - 1)).Copy
                  oSheet.Range("A" + CStr(wCurrRow) + ":G" + CStr(wCurrRow)).Select
                  oSheet.Paste
                End If
'              If wOlRow <= 57 Then
                oSheet.Cells(wCurrRow, 1) = !OlSrNo
                oSheet.Cells(wCurrRow, 2) = !OlMCd
                oSheet.Cells(wCurrRow, 3) = !OlSCd
                oSheet.Cells(wCurrRow, 4) = !OlQw
                oSheet.Cells(wCurrRow, 5) = !OlQty
                oSheet.Cells(wCurrRow, 6) = !OlSalRt
                ' ***** Manali 3.6.1 - Lab Val based on min/max val or OlVal
                oSheet.Cells(wCurrRow, 7) = "=IF($F" + CStr(wCurrRow) + " * (IF($D" + CStr(wCurrRow) + "=""Q"",$E" + CStr(wCurrRow) + ",(IF($D" + CStr(wCurrRow) + "=""W"",$E$" + CStr(wOlFRow - 2) + ",0))))=" + IIF(!OlQw = "Q", CStr(!OlQty), CStr(!OdLabAsWt)) + "*" + CStr(!OlSalRt) + "," + _
                          CStr(!OlSalVal) + ", $F" + CStr(wCurrRow) + " * (IF($D" + CStr(wCurrRow) + "=""Q"",$E" + CStr(wCurrRow) + ",(IF($D" + CStr(wCurrRow) + "=""W"",$E$" + CStr(wOlFRow - 2) + ",0)))))"
                oSheet.Cells(wCurrRow, 7).Font.Color = vbBlack
                
'              ElseIf wOlRow = 58 Then
'                oSheet.Cells(51, 1).AddComment = "#Labour Details InComplete"
'              End If
              wOlRow = wOlRow + 1
          End If
          wCurrRow = wCurrRow + 1   ' **** Manali 3.8.0
        wOdSr = !OdSr
        .MoveNext
      
      Else: Exit Do
      End If
    Loop
  End With
  Set oSheet = oBook.worksheets(1)
  oSheet.Rows(CStr(wOdRow) + ":299").Hidden = True
  
  oExcel.worksheets(2).Delete

  Set pv_rsOrd = Nothing
  
  '  oExcel.Visible = True
    

  Cd1.DialogTitle = "Save Excel File As"
  ' ***** Manali 3.9.0 - Excel 2007
  Cd1.FileName = ADC("UoOmTcFr") + "-" + ADC("UoOmYyFr") + "-" + ADC("UoOmChrFr") + "-" + CStr(ADC("UoOmNoFr")) + wExtn
  Cd1.Filter = "Excel Files (*" + wExtn + ") |*" + wExtn

  ' Bef  3.9.0
  'CD1.FileName = adc("UoOmTcFr") + "-" + adc("UoOmYyFr") + "-" + adc("UoOmChrFr") + "-" + CStr(adc("UoOmNoFr")) + ".xls"
  'CD1.Filter = "Excel Files (*.xls) |*.xls"

  Cd1.Flags = cdlOFNOverwritePrompt Or cdlOFNLongNames Or cdlOFNPathMustExist Or cdlOFNExtensionDifferent

  Cd1.CancelError = True
  Cd1.ShowSave

  If Len(Cd1.FileName) > 4 Then
    ' ***** Manali 3.9.0 - Excel 2007
    If UCase(Right(Cd1.FileName, 4)) = UCase(wExtn) Or UCase(Right(Cd1.FileName, 5)) = UCase(wExtn) Then
    ' Bef  3.9.0
    ''If UCase(Right(CD1.FileName, 4)) = UCase(".xls") Then
      'oExcel.ActiveWorkbook.SaveAs CD1.FileName
      oExcel.ActiveWorkbook.SaveAs Cd1.FileName, , "", ""       '4.1.0.0 (Pwd Removed)
    Else
      wErrMsg = "File Extension Is Wrong": Exit Sub
    End If
  End If

  If oExcel.ActiveWorkbook.Path + "\" + oExcel.ActiveWorkbook.Name <> sFilePath Then
    Kill sFilePath
  End If

  DispMsg "Excel Generated Successfully !!", etInfo
  oExcel.ActiveWorkbook.Close True
  oExcel.QUIT

  oExcel.DisplayAlerts = True
  Set oExcel = Nothing
  Set oBook = Nothing
  Set oSheet = Nothing
Exit Sub
Err_Hndlr:
  If Err.Number = 32755 Then
    wErrMsg = "Operation Cancelled"
    oExcel.ActiveWorkbook.Close
    oExcel.QUIT
  ElseIf Err.Number = 70 Then       '4.1.4
    DispMsg "Cannot Access Template File. Close If Already Open and Continue.", etError
  Else
    wErrMsg = Err.Description
  End If
  If Dir(sFilePath) <> "" Then Kill sFilePath
  Err.Clear
  
End Sub


Private Sub SetExcel_2(ByVal pv_rsOrd As MDORowSet)
Dim sFile, sFilePath As String
Dim wOdSr As Integer
Dim wOdRow As Integer, wOrMRow As Integer, wOrDCRow As Integer, wOrXRow As Integer, wOlRow As Integer, wOmRow As Integer, wLn1 As Integer, wLn2 As Integer
Dim wCoName As String, wSht As Integer
Dim wPictPath As String
Dim wMetLs As Single, wGmFactor As Double, wCmPurityRt As Double
Dim wGldAs As String, wGldAsWt As String
Dim wCurrRow As Integer   ' ***** Manali 3.8.0
Dim wOrMFRow As Integer, wOrDCFRow As Integer, wOrXfRow As Integer, wOlFRow As Integer

  Set oExcel = CreateObject("Excel.Application")
  
  sFile = gs_CoCd + "-" + gs_UsrCd + "-" + Format(Now, "ddMMyyHHmmss") + ".xls"
  sFilePath = App.Path + "\Output\" + sFile
  
  If Dir(sFilePath, vbNormal) <> "" Then
    Kill sFilePath
  End If
  
  FileCopy App.Path + "\Template\OrderPrint.xls", sFilePath
  
  oExcel.Visible = False: oExcel.DisplayAlerts = False
  
  wCoName = moCn.GetFldVal("Select HName from Head Where HCoCd='" + gs_CoCd + "' and HCd='" + ctSelfCmCd + "'")
  
  wOdSr = 0: wOdRow = 14: wOmRow = 1: wSht = 1
  
  Set oBook = oExcel.Workbooks.Open(sFilePath)
  
  With pv_rsOrd
    Do While Not (.EOF Or .BOF)
      If wOdRow <= 285 Then
        If !OdSr <> wOdSr Then
        
          'wiSht = oExcel.Worksheets.Count
          'oExcel.Worksheets(2).Copy after:=oExcel.ActiveSheet
          oExcel.worksheets(2).Copy after:=oBook.worksheets(wSht + 1)
    
          Set oSheet = oBook.worksheets(1)
          
          If wOmRow = 1 Then
            oSheet.Name = !OmCoCd + !OmYy + !OmTc + !OmChr + CStr(!OmNo)
            oSheet.Cells(1, 1) = wCoName
            oSheet.Cells(5, 3) = !OmCoCd + "/" + !OmYy + "/" + !OmTc + "/" + !OmChr + "/" + CStr(!OmNo)
            oSheet.Cells(5, 6) = !OmDt
            oSheet.Cells(5, 9) = !OmLmgSal
            oSheet.Cells(5, 13) = !OmLmsSal
            oSheet.Cells(6, 9) = !OmLmpSal
            oSheet.Cells(6, 13) = !OmLmlSal
            oSheet.Cells(6, 3) = !OmCmCd
            oSheet.Cells(7, 3) = !CmName
            oSheet.Cells(8, 3) = !OmPoNo
            oSheet.Cells(8, 6) = !OmPoDt
            oSheet.Cells(9, 3) = !OmDelDt
            oSheet.Cells(10, 1) = "Total (" + IIF(wb_InBaseCur, ws_BaseCurCd, ADC("UoCurCdFr")) + ")"
            oSheet.Cells(13, 8) = "Price (" + IIF(wb_InBaseCur, ws_BaseCurCd, ADC("UoCurCdFr")) + ")"
            oSheet.Cells(13, 9) = "Value (" + IIF(wb_InBaseCur, ws_BaseCurCd, ADC("UoCurCdFr")) + ")"
            
            If ADC("UoYn6") = "Y" Then
              oSheet.Cells(315, 1) = !OmCmDelTerms
              oSheet.Cells(320, 1) = !OmCmPayTerms
              oSheet.Cells(325, 1) = !OmSalRem
            Else
              oSheet.Range("A315", "D328").UnMerge
              oSheet.Range("A314", "D328") = ""
            End If
          End If
          oSheet.Cells(wOdRow, 1) = !OdSr
          oSheet.Cells(wOdRow, 2) = !OdDmCd
          oSheet.Cells(wOdRow, 4) = !OdSfx
          oSheet.Cells(wOdRow, 6) = !OdDmSz
          oSheet.Cells(wOdRow, 7) = !qOrdQty
          oSheet.Cells(wOdRow, 8) = !OdSalPrc
          oSheet.Cells(wOdRow, 9).Font.Color = vbBlack
          wOdRow = wOdRow + 1
          wOmRow = wOmRow + 1
          wSht = wSht + 1
        End If
        
        'Set oSheet = oBook.Worksheets(Worksheets.Count)
        Set oSheet = oBook.worksheets(wSht + 1)
        
        If !OdSr <> wOdSr Then
          oSheet.Name = !OdDmCd + "-" + "(" + CStr(!OdSr) + ")"
          oSheet.Cells(4, 3) = !OdDmCd
          oSheet.Cells(4, 6) = !OdSfx
          oSheet.Cells(4, 9) = !OdDmSz
          oSheet.Cells(5, 3) = !CdOurDmCd
          oSheet.Cells(13, 13) = !OdDmCd
          oSheet.Cells(14, 13) = !IdKt
          oSheet.Cells(14, 15) = !OdDmCol
          oSheet.Cells(4, 13) = !OdPicNm
          oSheet.Cells(6, 3) = !qOrdQty
          oSheet.Cells(7, 3) = !OdPoNo
          oSheet.Cells(8, 3) = !OdDelDt
          oSheet.Cells(9, 7) = !OdSalPrc
          oSheet.Cells(34, 5) = !OdLabAsWt
          oSheet.Cells(43, 10) = !OdMulBy
          
          wCurrRow = 18
          'wOrMRow = 18: wOrDCRow = 24: wOrXRow = 30: wOlRow = 36
          'wOrMRow = 18: wOrDCRow = 26: wOrXRow = 45: wOlRow = 53
          
          oSheet.Cells(9, 1) = "Price (" + IIF(wb_InBaseCur, ws_BaseCurCd, ADC("UoCurCdFr")) + ")"
          oSheet.Cells(9, 5) = "Sales Price (" + IIF(wb_InBaseCur, ws_BaseCurCd, ADC("UoCurCdFr")) + ")"
          oSheet.Cells(17, 8) = "LME (" + IIF(wb_InBaseCur, ws_BaseCurCd, ADC("UoCurCdFr")) + ")"
          oSheet.Cells(17, 9) = "Rate (" + IIF(wb_InBaseCur, ws_BaseCurCd, ADC("UoCurCdFr")) + ")"
          oSheet.Cells(17, 10) = "Value (" + IIF(wb_InBaseCur, ws_BaseCurCd, ADC("UoCurCdFr")) + ")"
          oSheet.Cells(23, 10) = "Rate (" + IIF(wb_InBaseCur, ws_BaseCurCd, ADC("UoCurCdFr")) + ")"
          oSheet.Cells(23, 11) = "Value (" + IIF(wb_InBaseCur, ws_BaseCurCd, ADC("UoCurCdFr")) + ")"
          oSheet.Cells(23, 13) = "Rate (" + IIF(wb_InBaseCur, ws_BaseCurCd, ADC("UoCurCdFr")) + ")"
          oSheet.Cells(23, 14) = "Value (" + IIF(wb_InBaseCur, ws_BaseCurCd, ADC("UoCurCdFr")) + ")"
          oSheet.Cells(29, 8) = "Rate (" + IIF(wb_InBaseCur, ws_BaseCurCd, ADC("UoCurCdFr")) + ")"
          oSheet.Cells(29, 9) = "Value (" + IIF(wb_InBaseCur, ws_BaseCurCd, ADC("UoCurCdFr")) + ")"
          oSheet.Cells(35, 6) = "Rate (" + IIF(wb_InBaseCur, ws_BaseCurCd, ADC("UoCurCdFr")) + ")"
          oSheet.Cells(35, 7) = "Value (" + IIF(wb_InBaseCur, ws_BaseCurCd, ADC("UoCurCdFr")) + ")"
          oSheet.Cells(41, 5) = "Value (" + IIF(wb_InBaseCur, ws_BaseCurCd, ADC("UoCurCdFr")) + ")"
          oSheet.Cells(41, 10) = "Value (" + IIF(wb_InBaseCur, ws_BaseCurCd, ADC("UoCurCdFr")) + ")"
          
          ' **** Manali 3.8.0
          wOrMRow = 1: wOrDCRow = 1: wOrXRow = 1: wOlRow = 1
          
          If ADC("UoYn") = "Y" Then
            ' **** Manali 3.6.1 - Pic Reference and Dm/SM
            wPictPath = GetPictPath(IIF(UCase(ADC("UoTcTypFr")) = "QS", "SM", "DM"), IIF(UCase(ADC("UoPMCdFr")) = "3D", False, True), !OdDmCd, !OdDmSz, !OmCoCd, !OmTc, !OmYy, !OmChr, !OmNo, !OdSr, !OdPicNm)
            If wPictPath = "" Or Dir(wPictPath) = "" Then
              wPictPath = GetPictPath(IIF(UCase(ADC("UoTcTypFr")) = "QS", "SM", "DM"), IIF(UCase(ADC("UoPMCdFr")) = "3D", True, False), !OdDmCd, !OdDmSz, !OmCoCd, !OmTc, !OmYy, !OmChr, !OmNo, !OdSr, !OdPicNm)
            End If
            'wPictPath = GetPictPath("DM", False, !OdDmCd, !OdDmSz, !OmCoCd, !OmTc, !OmYy, !OmChr, !OmNo, !OdSr, !OdPicNm)
            If Dir(wPictPath) <> "" And wPictPath <> "" Then
              oSheet.Pictures.Insert(wPictPath).Select
              oSheet.Pictures(1).Name = oSheet.Name
              oSheet.Pictures(1).Height = 110: oSheet.Pictures(1).Width = 140
              oSheet.Pictures(1).Left = 572: oSheet.Pictures(1).Top = 45
              oSheet.Pictures(1).Border.LineStyle = 1
            End If
          End If
          
          If ADC("UoYn6") = "Y" Then
            oSheet.Cells(52, 1) = !OdDmPrdInst
            oSheet.Cells(59, 1) = !OdCmPrdInst
            oSheet.Cells(66, 1) = !OdCmStmpInst
            oSheet.Cells(71, 1) = !OdSzInst
            oSheet.Cells(76, 1) = !OdSalRem
          Else
            oSheet.Range("A51", "G79").UnMerge
            oSheet.Range("A51", "G79") = ""
          End If
        End If
          
          If !qTag = "1RM" Then
            Select Case !RmCtg
            Case Is = "G", "P", "S", "L"
              ' ****** Manali 3.8.0 -
              wOrMRow = wOrMRow + 1
              If wOrMRow = 1 Or wOrMRow = 2 Then
                wOrMFRow = wCurrRow
              Else
                oSheet.Rows(CStr(wCurrRow) + ":" + CStr(wCurrRow)).Insert
                oSheet.Range("A" + CStr(wCurrRow) + ":J" + CStr(wCurrRow)).Copy
                oSheet.Range("A" + CStr(wCurrRow + 1) + ":J" + CStr(wCurrRow + 1)).Select
                oSheet.Paste
              End If
              
              'If wOrMRow <= 21 Then
                oSheet.Cells(wCurrRow, 1) = !OrSrNo
                oSheet.Cells(wCurrRow, 2) = !RmCtg
                oSheet.Cells(wCurrRow, 3) = !OrRmSCtg
                oSheet.Cells(wCurrRow, 4) = !OrRmCd
                oSheet.Cells(wCurrRow, 5) = !OrQty
                oSheet.Cells(wCurrRow, 6) = !OrWt
                oSheet.Cells(wCurrRow, 7) = !RmQw
                oSheet.Cells(wCurrRow, 8) = !OrLMESal
                
                wMetLs = moCn.GetFldVal("Select  Top 1 RrSalRt from RmRt " + _
                          "where RrTcTyp='LS' and RrCtg='" + !RmCtg + "' and RrCmCurCd='' and RrDmCd='' and RrSTWGrp='' and " + _
                          "((RrCmCtg = 'P' and RrCmCd = '" + !CmLkUpMetLs + "') Or (RrCmCtg='" + gs_CmCtg + "' and RrCmCd in ('" + !OmCmCd + "', '" + ctSelfCmCd + "'))) " + _
                          "and ((RrDmCtg = '" + !DmCtg + "' and RrLsCtg In ('" + !DmLsCtg + "', '')) Or (RrDmCtg='' and RrLsCtg='')) " + _
                          "Order By (Case When (RrCmCtg='P' and RrCmCd = '" + !CmLkUpMetLs + "') Then 1 " + _
                          "When RrCmCtg='" + gs_CmCtg + "' Then (Case RrCmCd When '" + !OmCmCd + "' Then 2 When '" + ctSelfCmCd + "' Then 3  End)  End), " + _
                          "(Case RrDmCtg When '" + !DmCtg + "' Then (Case RrLsCtg When '" + !DmLsCtg + "' Then 1 When '' Then 2  End) When '' Then 3  End) ")
  
                
                wGmFactor = moCn.GetFldVal("Select PNum from Param Where PTyp='GMFACTOR' and PMCd='" + !RmCtg + "' and PSCd='" + IIF(wb_InBaseCur, ws_BaseCurCd, ADC("UoCurCdFr")) + "'")
                wCmPurityRt = moCn.GetFldVal("Select RrSalRt from RmRt Where RrTcTyp='CRP' and RrCmCtg='" + gs_CmCtg + "' and RrCmCd='" + !OmCmCd + "' and RrCd='" + !OrRmCd + "' ")
                wCmPurityRt = IIF(wCmPurityRt = 0, !RmPurityRt, wCmPurityRt)
                
                oSheet.Cells(wCurrRow, 9).formula = "=(H" + CStr(wCurrRow) + "/" + CStr(wGmFactor) + "*" + CStr(wCmPurityRt) + "*(1 + " + CStr(wMetLs) + "/100)) + " + _
                                                "((1-" + CStr(wCmPurityRt) + ")*" + CStr(!OrAlySalRt) + ")"
  
                ' ***** Manali 3.6.1 - GldAs Wt
                wGldAsWt = "$F" + CStr(wCurrRow)
                If !OrMainMet = "Y" Then
                  wGldAs = moCn.GetFldVal("Select PDesc225 from Param Where Ptyp='GLDAS' and PMCd='" + !OdGldAs + "'")
                  
                  If InStr(1, "," + wGldAs + ",", ",D,") <> 0 Then wGldAsWt = wGldAsWt + IIF(wGldAsWt <> "", "+", "") + "(D63/5)"
                  If InStr(1, "," + wGldAs + ",", ",C,") <> 0 Then wGldAsWt = wGldAsWt + IIF(wGldAsWt <> "", "+", "") + "(D64/5)"
                  If InStr(1, "," + wGldAs + ",", ",X,") <> 0 Then wGldAsWt = wGldAsWt + IIF(wGldAsWt <> "", "+", "") + "(D65)"
                  
                End If
                wGldAsWt = "(" + wGldAsWt + ")"
                                        
                oSheet.Cells(wCurrRow, 10).formula = "=$I" + CStr(wCurrRow) + "*(IF($G" + CStr(wCurrRow) + "=""Q"",$E" + CStr(wCurrRow) + ",(IF($G" + CStr(wCurrRow) + "=""W""," + wGldAsWt + ",0))))"
                oSheet.Cells(wCurrRow, 10).Font.Color = vbBlack
                
                
              ' ***** Bef 3.8.0
              'ElseIf wOrMRow = 22 Then
              '  oSheet.Cells(16, 1).AddComment "#Metal Details InComplete"
              'End If
              'wOrMRow = wOrMRow + 1
              
            Case Is = "D", "C"
              'wOrDCRow = wOrDCRow + 1
                If wOrDCRow = 1 Then
                  wCurrRow = wCurrRow + IIF(wOrMRow = 1, 6, IIF(wOrMRow = 2, 5, 4))
                  wOrDCFRow = wCurrRow
                ElseIf wOrDCRow <> 2 Then
                  oSheet.Rows(CStr(wCurrRow) + ":" + CStr(wCurrRow)).Insert
                  oSheet.Range("A" + CStr(wCurrRow) + ":N" + CStr(wCurrRow)).Copy
                  oSheet.Range("A" + CStr(wCurrRow + 1) + ":N" + CStr(wCurrRow + 1)).Select
                  oSheet.Paste
                End If
                
'              If wOrDCRow <= 40 Then
                oSheet.Cells(wCurrRow, 1) = !OrSrNo
                oSheet.Cells(wCurrRow, 2) = !RmCtg
                oSheet.Cells(wCurrRow, 3) = !OrRmSCtg
                oSheet.Cells(wCurrRow, 4) = !OrRmCd
                oSheet.Cells(wCurrRow, 5) = !OrLn1
                oSheet.Cells(wCurrRow, 6) = !OrRmPtr
                oSheet.Cells(wCurrRow, 7) = !OrQty
                oSheet.Cells(wCurrRow, 8) = !OrWt
                oSheet.Cells(wCurrRow, 9) = !RmQw
                oSheet.Cells(wCurrRow, 10) = !OrSalRt
                oSheet.Cells(wCurrRow, 11).formula = "=$J" + CStr(wCurrRow) + "*(IF($I" + CStr(wCurrRow) + "=""Q"",$G" + CStr(wCurrRow) + ",(IF($I" + CStr(wCurrRow) + "=""W"",$H" + CStr(wCurrRow) + ",0))))"
                oSheet.Cells(wCurrRow, 11).Font.Color = vbBlack
                oSheet.Cells(wCurrRow, 12) = !OrSetSCd
                oSheet.Cells(wCurrRow, 13) = !OrSetSalRt
                oSheet.Cells(wCurrRow, 14).formula = "=$M" + CStr(wCurrRow) + "*$G" + CStr(wCurrRow) + ""
                oSheet.Cells(wCurrRow, 14).Font.Color = vbBlack
                
              wOrDCRow = wOrDCRow + 1
            
            Case Is = "X"
                If wOrXRow = 1 Then
                  wCurrRow = wCurrRow + IIF(wOrDCRow = 1, 6 + IIF(wOrMRow = 1, 6, IIF(wOrMRow = 2, 5, 4)), IIF(wOrDCRow = 2, 5 + IIF(wOrMRow = 1, 6, IIF(wOrMRow = 2, 5, 4)), 4))
                  wOrXfRow = wCurrRow
                ElseIf wOrXRow <> 2 Then
                  oSheet.Rows(CStr(wCurrRow) + ":" + CStr(wCurrRow)).Insert
                  oSheet.Range("A" + CStr(wCurrRow) + ":I" + CStr(wCurrRow)).Copy
                  oSheet.Range("A" + CStr(wCurrRow + 1) + ":I" + CStr(wCurrRow + 1)).Select
                  oSheet.Paste
                End If
'              If wOrXRow <= 48 Then
                oSheet.Cells(wCurrRow, 1) = !OrSrNo
                oSheet.Cells(wCurrRow, 2) = !RmCtg
                oSheet.Cells(wCurrRow, 3) = !OrRmSCtg
                oSheet.Cells(wCurrRow, 4) = !OrRmCd
                oSheet.Cells(wCurrRow, 5) = !OrQty
                oSheet.Cells(wCurrRow, 6) = !OrWt
                oSheet.Cells(wCurrRow, 7) = !RmQw
                oSheet.Cells(wCurrRow, 8) = !OrSalRt
                oSheet.Cells(wCurrRow, 9).formula = "=$H" + CStr(wCurrRow) + "*(IF($G" + CStr(wCurrRow) + "=""Q"",$E" + CStr(wCurrRow) + ",(IF($G" + CStr(wCurrRow) + "=""W"",$F" + CStr(wCurrRow) + ",0))))"
                oSheet.Cells(wCurrRow, 9).Font.Color = vbBlack
                
                
'              ElseIf wOrXRow = 49 Then
'                oSheet.Cells(43, 1).AddComment = "#Acc Details InComplete"
'              End If
              wOrXRow = wOrXRow + 1
            End Select
          
          ElseIf !qTag = "2LAB" Then
                If wOlRow = 1 Then
                  wCurrRow = wCurrRow + IIF(wOrXRow = 1, 6 + IIF(wOrDCRow = 1, 6 + IIF(wOrMRow = 1, 6, IIF(wOrMRow = 2, 5, 4)), IIF(wOrDCRow = 2, 5, 4)), _
                                        IIF(wOrXRow = 2, 5, 4))
                                        'IIF(wOrDCRow = 1, 6 + IIF(wOrMRow = 1, 6, IIF(wOrMRow = 2, 5, 0)), IIF(wOrDCRow = 2, 5 + IIF(wOrMRow = 1, 6, IIF(wOrMRow = 2, 5, 0)), 0))
                  wOlFRow = wCurrRow
                ElseIf wOlRow <> 2 Then
                  oSheet.Rows(CStr(wCurrRow) + ":" + CStr(wCurrRow)).Insert
                  oSheet.Range("A" + CStr(wCurrRow) + ":G" + CStr(wCurrRow)).Copy
                  oSheet.Range("A" + CStr(wCurrRow + 1) + ":G" + CStr(wCurrRow + 1)).Select
                  oSheet.Paste
                End If
'              If wOlRow <= 57 Then
                oSheet.Cells(wCurrRow, 1) = !OlSrNo
                oSheet.Cells(wCurrRow, 2) = !OlMCd
                oSheet.Cells(wCurrRow, 3) = !OlSCd
                oSheet.Cells(wCurrRow, 4) = !OlQw
                oSheet.Cells(wCurrRow, 5) = !OlQty
                oSheet.Cells(wCurrRow, 6) = !OlSalRt
                ' ***** Manali 3.6.1 - Lab Val based on min/max val or OlVal
                oSheet.Cells(wCurrRow, 7) = "=IF($F" + CStr(wCurrRow) + " * (IF($D" + CStr(wCurrRow) + "=""Q"",$E" + CStr(wCurrRow) + ",(IF($D" + CStr(wCurrRow) + "=""W"",$E$51,0))))=" + IIF(!OlQw = "Q", CStr(!OlQty), CStr(!OdLabAsWt)) + "*" + CStr(!OlSalRt) + "," + _
                          CStr(!OlSalVal) + ", $F" + CStr(wCurrRow) + " * (IF($D" + CStr(wCurrRow) + "=""Q"",$E" + CStr(wCurrRow) + ",(IF($D" + CStr(wCurrRow) + "=""W"",$E$51,0)))))"
                oSheet.Cells(wCurrRow, 7).Font.Color = vbBlack
                
                
'              ElseIf wOlRow = 58 Then
'                oSheet.Cells(51, 1).AddComment = "#Labour Details InComplete"
'              End If
              wOlRow = wOlRow + 1
          End If
          wCurrRow = wCurrRow + 1   ' **** Manali 3.8.0
        wOdSr = !OdSr
        .MoveNext
      
      Else: Exit Do
      End If
    Loop
  End With
  Set oSheet = oBook.worksheets(1)
  oSheet.Rows(CStr(wOdRow) + ":299").Hidden = True
  
  oExcel.worksheets(2).Delete

  Set pv_rsOrd = Nothing
  
  Cd1.DialogTitle = "Save Excel File As"
  Cd1.FileName = ADC("UoOmTcFr") + "-" + ADC("UoOmYyFr") + "-" + ADC("UoOmChrFr") + "-" + CStr(ADC("UoOmNoFr")) + ".xls"
  Cd1.Filter = "Excel Files (*.xls) |*.xls"
  
  Cd1.Flags = cdlOFNOverwritePrompt Or cdlOFNLongNames Or cdlOFNPathMustExist Or cdlOFNExtensionDifferent
  
  Cd1.CancelError = True
  Cd1.ShowSave

  If Len(Cd1.FileName) > 4 Then
    If UCase(Right(Cd1.FileName, 4)) = UCase(".xls") Then
      'oExcel.ActiveWorkbook.SaveAs Cd1.FileName
      oExcel.ActiveWorkbook.SaveAs Cd1.FileName, , "", ""       '4.1.0.0 (Pwd Removed)
    Else
      wErrMsg = "File Extension Is Wrong": Exit Sub
    End If
  End If
  
  If oExcel.ActiveWorkbook.Path + "\" + oExcel.ActiveWorkbook.Name <> sFilePath Then
    Kill sFilePath
  End If
  
  DispMsg "Excel Generated Successfully !!", etInfo
  oExcel.ActiveWorkbook.Close True
  oExcel.QUIT
  
  oExcel.DisplayAlerts = True
  Set oExcel = Nothing
  Set oBook = Nothing
  Set oSheet = Nothing
Exit Sub
Err_Hndlr:
  If Err.Number = 32755 Then
    wErrMsg = "Operation Cancelled"
    oExcel.ActiveWorkbook.Close
    oExcel.QUIT
  Else
    wErrMsg = Err.Description
  End If
  If Dir(sFilePath) <> "" Then Kill sFilePath
  Err.Clear
  
End Sub


Private Sub SetExcel_old(ByVal pv_rsOrd As MDORowSet)
Dim sFile, sFilePath As String
Dim wOdSr As Integer
Dim wOdRow As Integer, wOrMRow As Integer, wOrDCRow As Integer, wOrXRow As Integer, wOlRow As Integer, wOmRow As Integer, wLn1 As Integer, wLn2 As Integer
Dim wCoName As String, wSht As Integer
Dim wPictPath As String
Dim wMetLs As Single, wGmFactor As Double, wCmPurityRt As Double
Dim wGldAs As String, wGldAsWt As String
Dim wCurrRow As Integer   ' ***** Manali 3.8.0
Dim wOrMFRow As Integer, wOrDCFRow As Integer, wOrXfRow As Integer, wOlFRow As Integer

  Set oExcel = CreateObject("Excel.Application")
  
  sFile = gs_CoCd + "-" + gs_UsrCd + "-" + Format(Now, "ddMMyyHHmmss") + ".xls"
  sFilePath = App.Path + "\Output\" + sFile
  
  If Dir(sFilePath, vbNormal) <> "" Then
    Kill sFilePath
  End If
  
  FileCopy App.Path + "\Template\OrderPrint.xls", sFilePath
  
  oExcel.Visible = False: oExcel.DisplayAlerts = False
  
  wCoName = moCn.GetFldVal("Select HName from Head Where HCoCd='" + gs_CoCd + "' and HCd='" + ctSelfCmCd + "'")
  
  wOdSr = 0: wOdRow = 14: wOmRow = 1: wSht = 1
  
  Set oBook = oExcel.Workbooks.Open(sFilePath)
  
  With pv_rsOrd
    Do While Not (.EOF Or .BOF)
      If wOdRow <= 285 Then
        If !OdSr <> wOdSr Then
        
          'wiSht = oExcel.Worksheets.Count
          'oExcel.Worksheets(2).Copy after:=oExcel.ActiveSheet
          oExcel.worksheets(2).Copy after:=oBook.worksheets(wSht + 1)
    
          Set oSheet = oBook.worksheets(1)
          
          If wOmRow = 1 Then
            oSheet.Name = !OmCoCd + !OmYy + !OmTc + !OmChr + CStr(!OmNo)
            oSheet.Cells(1, 1) = wCoName
            oSheet.Cells(5, 3) = !OmCoCd + "/" + !OmYy + "/" + !OmTc + "/" + !OmChr + "/" + CStr(!OmNo)
            oSheet.Cells(5, 6) = !OmDt
            oSheet.Cells(5, 9) = !OmLmgSal
            oSheet.Cells(5, 13) = !OmLmsSal
            oSheet.Cells(6, 9) = !OmLmpSal
            oSheet.Cells(6, 13) = !OmLmlSal
            oSheet.Cells(6, 3) = !OmCmCd
            oSheet.Cells(7, 3) = !CmName
            oSheet.Cells(8, 3) = !OmPoNo
            oSheet.Cells(8, 6) = !OmPoDt
            oSheet.Cells(9, 3) = !OmDelDt
            oSheet.Cells(10, 1) = "Total (" + IIF(wb_InBaseCur, ws_BaseCurCd, ADC("UoCurCdFr")) + ")"
            oSheet.Cells(13, 8) = "Price (" + IIF(wb_InBaseCur, ws_BaseCurCd, ADC("UoCurCdFr")) + ")"
            oSheet.Cells(13, 9) = "Value (" + IIF(wb_InBaseCur, ws_BaseCurCd, ADC("UoCurCdFr")) + ")"
            
            If ADC("UoYn6") = "Y" Then
              oSheet.Cells(315, 1) = !OmCmDelTerms
              oSheet.Cells(320, 1) = !OmCmPayTerms
              oSheet.Cells(325, 1) = !OmSalRem
            Else
              oSheet.Range("A315", "D328").UnMerge
              oSheet.Range("A314", "D328") = ""
            End If
          End If
          oSheet.Cells(wOdRow, 1) = !OdSr
          oSheet.Cells(wOdRow, 2) = !OdDmCd
          oSheet.Cells(wOdRow, 4) = !OdSfx
          oSheet.Cells(wOdRow, 6) = !OdDmSz
          oSheet.Cells(wOdRow, 7) = !qOrdQty
          oSheet.Cells(wOdRow, 8) = !OdSalPrc
          oSheet.Cells(wOdRow, 9).Font.Color = vbBlack
          wOdRow = wOdRow + 1
          wOmRow = wOmRow + 1
          wSht = wSht + 1
        End If
        
        'Set oSheet = oBook.Worksheets(Worksheets.Count)
        Set oSheet = oBook.worksheets(wSht + 1)
        
        If !OdSr <> wOdSr Then
          oSheet.Name = !OdDmCd + "-" + "(" + CStr(!OdSr) + ")"
          oSheet.Cells(4, 3) = !OdDmCd
          oSheet.Cells(4, 6) = !OdSfx
          oSheet.Cells(4, 9) = !OdDmSz
          oSheet.Cells(5, 3) = !CdOurDmCd
          oSheet.Cells(13, 13) = !OdDmCd
          oSheet.Cells(14, 13) = !IdKt
          oSheet.Cells(14, 15) = !OdDmCol
          oSheet.Cells(4, 13) = !OdPicNm
          oSheet.Cells(6, 3) = !qOrdQty
          oSheet.Cells(7, 3) = !OdPoNo
          oSheet.Cells(8, 3) = !OdDelDt
          oSheet.Cells(9, 7) = !OdSalPrc
          oSheet.Cells(31, 5) = !OdLabAsWt
          oSheet.Cells(39, 10) = !OdMulBy
          
          wCurrRow = 18
          wOrMRow = 18: wOrDCRow = 23: wOrXRow = 28: wOlRow = 33
          'wOrMRow = 18: wOrDCRow = 26: wOrXRow = 45: wOlRow = 53
          
          oSheet.Cells(11, 1) = "Price (" + IIF(wb_InBaseCur, ws_BaseCurCd, ADC("UoCurCdFr")) + ")"
          oSheet.Cells(11, 5) = "Sales Price (" + IIF(wb_InBaseCur, ws_BaseCurCd, ADC("UoCurCdFr")) + ")"
          oSheet.Cells(wOrMRow - 1, 8) = "LME (" + IIF(wb_InBaseCur, ws_BaseCurCd, ADC("UoCurCdFr")) + ")"
          oSheet.Cells(wOrMRow - 1, 9) = "Rate (" + IIF(wb_InBaseCur, ws_BaseCurCd, ADC("UoCurCdFr")) + ")"
          oSheet.Cells(wOrMRow - 1, 10) = "Value (" + IIF(wb_InBaseCur, ws_BaseCurCd, ADC("UoCurCdFr")) + ")"
          oSheet.Cells(wOrDCRow - 1, 10) = "Rate (" + IIF(wb_InBaseCur, ws_BaseCurCd, ADC("UoCurCdFr")) + ")"
          oSheet.Cells(wOrDCRow - 1, 11) = "Value (" + IIF(wb_InBaseCur, ws_BaseCurCd, ADC("UoCurCdFr")) + ")"
          oSheet.Cells(wOrDCRow - 1, 13) = "Rate (" + IIF(wb_InBaseCur, ws_BaseCurCd, ADC("UoCurCdFr")) + ")"
          oSheet.Cells(wOrDCRow - 1, 14) = "Value (" + IIF(wb_InBaseCur, ws_BaseCurCd, ADC("UoCurCdFr")) + ")"
          oSheet.Cells(wOrXRow - 1, 8) = "Rate (" + IIF(wb_InBaseCur, ws_BaseCurCd, ADC("UoCurCdFr")) + ")"
          oSheet.Cells(wOrXRow - 1, 9) = "Value (" + IIF(wb_InBaseCur, ws_BaseCurCd, ADC("UoCurCdFr")) + ")"
          oSheet.Cells(wOlRow - 1, 6) = "Rate (" + IIF(wb_InBaseCur, ws_BaseCurCd, ADC("UoCurCdFr")) + ")"
          oSheet.Cells(wOlRow - 1, 7) = "Value (" + IIF(wb_InBaseCur, ws_BaseCurCd, ADC("UoCurCdFr")) + ")"
          oSheet.Cells(37, 5) = "Value (" + IIF(wb_InBaseCur, ws_BaseCurCd, ADC("UoCurCdFr")) + ")"
          oSheet.Cells(37, 10) = "Value (" + IIF(wb_InBaseCur, ws_BaseCurCd, ADC("UoCurCdFr")) + ")"
          
          ' **** Manali 3.8.0
          wOrMRow = 1: wOrDCRow = 1: wOrXRow = 1: wOlRow = 1
          
          If ADC("UoYn") = "Y" Then
            ' **** Manali 3.6.1 - Pic Reference and Dm/SM
            wPictPath = GetPictPath(IIF(UCase(ADC("UoTcTypFr")) = "QS", "SM", "DM"), IIF(UCase(ADC("UoPMCdFr")) = "3D", False, True), !OdDmCd, !OdDmSz, !OmCoCd, !OmTc, !OmYy, !OmChr, !OmNo, !OdSr, !OdPicNm)
            If wPictPath = "" Or Dir(wPictPath) = "" Then
              wPictPath = GetPictPath(IIF(UCase(ADC("UoTcTypFr")) = "QS", "SM", "DM"), IIF(UCase(ADC("UoPMCdFr")) = "3D", True, False), !OdDmCd, !OdDmSz, !OmCoCd, !OmTc, !OmYy, !OmChr, !OmNo, !OdSr, !OdPicNm)
            End If
            'wPictPath = GetPictPath("DM", False, !OdDmCd, !OdDmSz, !OmCoCd, !OmTc, !OmYy, !OmChr, !OmNo, !OdSr, !OdPicNm)
            If Dir(wPictPath) <> "" And wPictPath <> "" Then
              oSheet.Pictures.Insert(wPictPath).Select
              oSheet.Pictures(1).Name = oSheet.Name
              oSheet.Pictures(1).Height = 110: oSheet.Pictures(1).Width = 140
              oSheet.Pictures(1).Left = 572: oSheet.Pictures(1).Top = 45
              oSheet.Pictures(1).Border.LineStyle = 1
            End If
          End If
          
          If ADC("UoYn6") = "Y" Then
            oSheet.Cells(48, 1) = !OdDmPrdInst
            oSheet.Cells(55, 1) = !OdCmPrdInst
            oSheet.Cells(62, 1) = !OdCmStmpInst
            oSheet.Cells(67, 1) = !OdSzInst
            oSheet.Cells(72, 1) = !OdSalRem
          Else
            oSheet.Range("A47", "G75").UnMerge
            oSheet.Range("A47", "G75") = ""
          End If
        End If
          
          If !qTag = "1RM" Then
            Select Case !RmCtg
            Case Is = "G", "P", "S", "L"
              ' ****** Manali 3.8.0 -
              If wOrMRow = 1 Then wOrMFRow = wCurrRow Else oSheet.Rows(CStr(wCurrRow) + ":" + CStr(wCurrRow)).Insert
              wOrMRow = wCurrRow
              'If wOrMRow <= 21 Then
                oSheet.Cells(wOrMRow, 1) = !OrSrNo
                oSheet.Cells(wOrMRow, 2) = !RmCtg
                oSheet.Cells(wOrMRow, 3) = !OrRmSCtg
                oSheet.Cells(wOrMRow, 4) = !OrRmCd
                oSheet.Cells(wOrMRow, 5) = !OrQty
                oSheet.Cells(wOrMRow, 6) = !OrWt
                oSheet.Cells(wOrMRow, 7) = !RmQw
                oSheet.Cells(wOrMRow, 8) = !OrLMESal
                
                wMetLs = moCn.GetFldVal("Select  Top 1 RrSalRt from RmRt " + _
                          "where RrTcTyp='LS' and RrCtg='" + !RmCtg + "' and RrCmCurCd='' and RrDmCd='' and RrSTWGrp='' and " + _
                          "((RrCmCtg = 'P' and RrCmCd = '" + !CmLkUpMetLs + "') Or (RrCmCtg='" + gs_CmCtg + "' and RrCmCd in ('" + !OmCmCd + "', '" + ctSelfCmCd + "'))) " + _
                          "and ((RrDmCtg = '" + !DmCtg + "' and RrLsCtg In ('" + !DmLsCtg + "', '')) Or (RrDmCtg='' and RrLsCtg='')) " + _
                          "Order By (Case When (RrCmCtg='P' and RrCmCd = '" + !CmLkUpMetLs + "') Then 1 " + _
                          "When RrCmCtg='" + gs_CmCtg + "' Then (Case RrCmCd When '" + !OmCmCd + "' Then 2 When '" + ctSelfCmCd + "' Then 3  End)  End), " + _
                          "(Case RrDmCtg When '" + !DmCtg + "' Then (Case RrLsCtg When '" + !DmLsCtg + "' Then 1 When '' Then 2  End) When '' Then 3  End) ")
  
                
                wGmFactor = moCn.GetFldVal("Select PNum from Param Where PTyp='GMFACTOR' and PMCd='" + !RmCtg + "' and PSCd='" + IIF(wb_InBaseCur, ws_BaseCurCd, ADC("UoCurCdFr")) + "'")
                wCmPurityRt = moCn.GetFldVal("Select RrSalRt from RmRt Where RrTcTyp='CRP' and RrCmCtg='" + gs_CmCtg + "' and RrCmCd='" + !OmCmCd + "' and RrCd='" + !OrRmCd + "' ")
                wCmPurityRt = IIF(wCmPurityRt = 0, !RmPurityRt, wCmPurityRt)
                
                oSheet.Cells(wOrMRow, 9).formula = "=(H" + CStr(wOrMRow) + "/" + CStr(wGmFactor) + "*" + CStr(wCmPurityRt) + "*(1 + " + CStr(wMetLs) + "/100)) + " + _
                                                "((1-" + CStr(wCmPurityRt) + ")*" + CStr(!OrAlySalRt) + ")"
  
                ' ***** Manali 3.6.1 - GldAs Wt
                wGldAsWt = "$F" + CStr(wOrMRow)
                If !OrMainMet = "Y" Then
                  wGldAs = moCn.GetFldVal("Select PDesc225 from Param Where Ptyp='GLDAS' and PMCd='" + !OdGldAs + "'")
                  
                  If InStr(1, "," + wGldAs + ",", ",D,") <> 0 Then wGldAsWt = wGldAsWt + IIF(wGldAsWt <> "", "+", "") + "(D63/5)"
                  If InStr(1, "," + wGldAs + ",", ",C,") <> 0 Then wGldAsWt = wGldAsWt + IIF(wGldAsWt <> "", "+", "") + "(D64/5)"
                  If InStr(1, "," + wGldAs + ",", ",X,") <> 0 Then wGldAsWt = wGldAsWt + IIF(wGldAsWt <> "", "+", "") + "(D65)"
                  
                End If
                wGldAsWt = "(" + wGldAsWt + ")"
                                        
                oSheet.Cells(wOrMRow, 10).formula = "=$I" + CStr(wOrMRow) + "*(IF($G" + CStr(wOrMRow) + "=""Q"",$E" + CStr(wOrMRow) + ",(IF($G" + CStr(wOrMRow) + "=""W""," + wGldAsWt + ",0))))"
                oSheet.Cells(wOrMRow, 10).Font.Color = vbBlack
                
                
                
                oSheet.Cells(wOrMRow + 1, 5).formula = "=SUM($E" + CStr(wOrMFRow) + ":$E" + CStr(wOrMRow) + ")"
                oSheet.Cells(wOrMRow + 1, 6).formula = "=SUM($F" + CStr(wOrMFRow) + ":$F" + CStr(wOrMRow) + ")"
                oSheet.Cells(wOrMRow + 1, 10).formula = "=SUM($J" + CStr(wOrMFRow) + ":$J" + CStr(wOrMRow) + ")"
                
              ' ***** Bef 3.8.0
              'ElseIf wOrMRow = 22 Then
              '  oSheet.Cells(16, 1).AddComment "#Metal Details InComplete"
              'End If
              'wOrMRow = wOrMRow + 1
              
            Case Is = "D", "C"
                If wOrDCRow = 1 Then
                  wCurrRow = wCurrRow + 4 + IIF(wOrMRow = 1, 1, 0)
                  wOrDCFRow = wCurrRow
                Else
                  oSheet.Rows(CStr(wCurrRow) + ":" + CStr(wCurrRow)).Insert
                End If
                wOrDCRow = wCurrRow
'              If wOrDCRow <= 40 Then
                oSheet.Cells(wOrDCRow, 1) = !OrSrNo
                oSheet.Cells(wOrDCRow, 2) = !RmCtg
                oSheet.Cells(wOrDCRow, 3) = !OrRmSCtg
                oSheet.Cells(wOrDCRow, 4) = !OrRmCd
                oSheet.Cells(wOrDCRow, 5) = !OrLn1
                oSheet.Cells(wOrDCRow, 6) = !OrRmPtr
                oSheet.Cells(wOrDCRow, 7) = !OrQty
                oSheet.Cells(wOrDCRow, 8) = !OrWt
                oSheet.Cells(wOrDCRow, 9) = !RmQw
                oSheet.Cells(wOrDCRow, 10) = !OrSalRt
                oSheet.Cells(wOrDCRow, 11).formula = "=$J" + CStr(wOrDCRow) + "*(IF($I" + CStr(wOrDCRow) + "=""Q"",$G" + CStr(wOrDCRow) + ",(IF($I" + CStr(wOrDCRow) + "=""W"",$H" + CStr(wOrDCRow) + ",0))))"
                oSheet.Cells(wOrDCRow, 11).Font.Color = vbBlack
                oSheet.Cells(wOrDCRow, 12) = !OrSetSCd
                oSheet.Cells(wOrDCRow, 13) = !OrSetSalRt
                oSheet.Cells(wOrDCRow, 14).formula = "=$M" + CStr(wOrDCRow) + "*$G" + CStr(wOrDCRow) + ""
                oSheet.Cells(wOrDCRow, 14).Font.Color = vbBlack
                
                oSheet.Cells(wOrDCRow + 1, 7).formula = "=SUM($G" + CStr(wOrDCFRow) + ":$G" + CStr(wOrDCRow) + ")"
                oSheet.Cells(wOrDCRow + 1, 8).formula = "=SUM($H" + CStr(wOrDCFRow) + ":$H" + CStr(wOrDCRow) + ")"
                oSheet.Cells(wOrDCRow + 1, 11).formula = "=SUM($K" + CStr(wOrDCFRow) + ":$K" + CStr(wOrDCRow) + ")"
                oSheet.Cells(wOrDCRow + 1, 14).formula = "=SUM($N" + CStr(wOrDCFRow) + ":$N" + CStr(wOrDCRow) + ")"

                oSheet.Cells(wOrDCRow + 16, 4).formula = "=SUMIF(B" + CStr(wOrDCFRow) + ":B" + CStr(wOrDCRow) + ",""D"", H" + CStr(wOrDCFRow) + ":H" + CStr(wOrDCRow) + ")"
                oSheet.Cells(wOrDCRow + 17, 4).formula = "=SUMIF(B" + CStr(wOrDCFRow) + ":B" + CStr(wOrDCRow) + ",""C"", H" + CStr(wOrDCFRow) + ":H" + CStr(wOrDCRow) + ")"
                oSheet.Cells(wOrDCRow + 16, 3).formula = "=SUMIF(B" + CStr(wOrDCFRow) + ":B" + CStr(wOrDCRow) + ",""D"", G" + CStr(wOrDCFRow) + ":G" + CStr(wOrDCRow) + ")"
                oSheet.Cells(wOrDCRow + 17, 3).formula = "=SUMIF(B" + CStr(wOrDCFRow) + ":B" + CStr(wOrDCRow) + ",""C"", G" + CStr(wOrDCFRow) + ":G" + CStr(wOrDCRow) + ")"
                
'              ElseIf wOrDCRow = 41 Then
'                oSheet.Cells(24, 1).AddComment "#Studding Details Incomplete"
'              End If
'              wOrDCRow = wOrDCRow + 1
            
            Case Is = "X"
                If wOrXRow = 1 Then
                  wCurrRow = wCurrRow + IIF(wOrDCRow = 1, 5 + IIF(wOrMRow = 1, 5, 4), 4)
                  wOrXfRow = wCurrRow
                Else
                  oSheet.Rows(CStr(wCurrRow) + ":" + CStr(wCurrRow)).Insert
                End If
                wOrXRow = wCurrRow
'              If wOrXRow <= 48 Then
                oSheet.Cells(wOrXRow, 1) = !OrSrNo
                oSheet.Cells(wOrXRow, 2) = !RmCtg
                oSheet.Cells(wOrXRow, 3) = !OrRmSCtg
                oSheet.Cells(wOrXRow, 4) = !OrRmCd
                oSheet.Cells(wOrXRow, 5) = !OrQty
                oSheet.Cells(wOrXRow, 6) = !OrWt
                oSheet.Cells(wOrXRow, 7) = !RmQw
                oSheet.Cells(wOrXRow, 8) = !OrSalRt
                oSheet.Cells(wOrXRow, 9).formula = "=$H" + CStr(wOrXRow) + "*(IF($G" + CStr(wOrXRow) + "=""Q"",$E" + CStr(wOrXRow) + ",(IF($G" + CStr(wOrXRow) + "=""W"",$F" + CStr(wOrXRow) + ",0))))"
                oSheet.Cells(wOrXRow, 9).Font.Color = vbBlack
                
                oSheet.Cells(wOrXRow + 1, 5).formula = "=SUM($E" + CStr(wOrXfRow) + ":$E" + CStr(wOrXRow) + ")"
                oSheet.Cells(wOrXRow + 1, 6).formula = "=SUM($F" + CStr(wOrXfRow) + ":$F" + CStr(wOrXRow) + ")"
                oSheet.Cells(wOrXRow + 1, 9).formula = "=SUM($I" + CStr(wOrXfRow) + ":$I" + CStr(wOrXRow) + ")"
                
'              ElseIf wOrXRow = 49 Then
'                oSheet.Cells(43, 1).AddComment = "#Acc Details InComplete"
'              End If
'              wOrXRow = wOrXRow + 1
            End Select
          
          ElseIf !qTag = "2LAB" Then
                If wOlRow = 1 Then
                  wCurrRow = wCurrRow + IIF(wOrXRow = 1, 5 + IIF(wOrDCRow = 1, 5 + IIF(wOrMRow = 1, 5, 4), 4), 4)
                  wOlFRow = wCurrRow
                Else
                  oSheet.Rows(CStr(wCurrRow) + ":" + CStr(wCurrRow)).Insert
                End If
                wOlRow = wCurrRow
'              If wOlRow <= 57 Then
                oSheet.Cells(wOlRow, 1) = !OlSrNo
                oSheet.Cells(wOlRow, 2) = !OlMCd
                oSheet.Cells(wOlRow, 3) = !OlSCd
                oSheet.Cells(wOlRow, 4) = !OlQw
                oSheet.Cells(wOlRow, 5) = !OlQty
                oSheet.Cells(wOlRow, 6) = !OlSalRt
                ' ***** Manali 3.6.1 - Lab Val based on min/max val or OlVal
                oSheet.Cells(wOlRow, 7) = "=IF($F" + CStr(wOlRow) + " * (IF($D" + CStr(wOlRow) + "=""Q"",$E" + CStr(wOlRow) + ",(IF($D" + CStr(wOlRow) + "=""W"",$E$51,0))))=" + IIF(!OlQw = "Q", CStr(!OlQty), CStr(!OdLabAsWt)) + "*" + CStr(!OlSalRt) + "," + _
                          CStr(!OlSalVal) + ", $F" + CStr(wOlRow) + " * (IF($D" + CStr(wOlRow) + "=""Q"",$E" + CStr(wOlRow) + ",(IF($D" + CStr(wOlRow) + "=""W"",$E$51,0)))))"
                oSheet.Cells(wOlRow, 7).Font.Color = vbBlack
                
                oSheet.Cells(wOlRow + 1, 5).formula = "=SUM($E" + CStr(wOlFRow) + ":$E" + CStr(wOlRow) + ")"
                oSheet.Cells(wOlRow + 1, 7).formula = "=SUM($G" + CStr(wOlFRow) + ":$G" + CStr(wOlRow) + ")"
                
'              ElseIf wOlRow = 58 Then
'                oSheet.Cells(51, 1).AddComment = "#Labour Details InComplete"
'              End If
'              wOlRow = wOlRow + 1
          End If
          wCurrRow = wCurrRow + 1   ' **** Manali 3.8.0
        wOdSr = !OdSr
        .MoveNext
      
      Else: Exit Do
      End If
    Loop
  End With
  Set oSheet = oBook.worksheets(1)
  oSheet.Rows(CStr(wOdRow) + ":299").Hidden = True
  
  oExcel.worksheets(2).Delete

  Set pv_rsOrd = Nothing
  
  Cd1.DialogTitle = "Save Excel File As"
  Cd1.FileName = ADC("UoOmTcFr") + "-" + ADC("UoOmYyFr") + "-" + ADC("UoOmChrFr") + "-" + CStr(ADC("UoOmNoFr")) + ".xls"
  Cd1.Filter = "Excel Files (*.xls) |*.xls"
  
  Cd1.Flags = cdlOFNOverwritePrompt Or cdlOFNLongNames Or cdlOFNPathMustExist Or cdlOFNExtensionDifferent
  
  Cd1.CancelError = True
  Cd1.ShowSave

  If Len(Cd1.FileName) > 4 Then
    If UCase(Right(Cd1.FileName, 4)) = UCase(".xls") Then
      'oExcel.ActiveWorkbook.SaveAs Cd1.FileName
      oExcel.ActiveWorkbook.SaveAs Cd1.FileName, , "", ""       '4.1.0.0 (Pwd Removed)
    Else
      wErrMsg = "File Extension Is Wrong": Exit Sub
    End If
  End If
  
  If oExcel.ActiveWorkbook.Path + "\" + oExcel.ActiveWorkbook.Name <> sFilePath Then
    Kill sFilePath
  End If
  
  DispMsg "Excel Generated Successfully !!", etInfo
  oExcel.ActiveWorkbook.Close True
  oExcel.QUIT
  
  oExcel.DisplayAlerts = True
  Set oExcel = Nothing
  Set oBook = Nothing
  Set oSheet = Nothing
Exit Sub
Err_Hndlr:
  If Err.Number = 32755 Then
    wErrMsg = "Operation Cancelled"
    oExcel.ActiveWorkbook.Close
    oExcel.QUIT
  Else
    wErrMsg = Err.Description
  End If
  If Dir(sFilePath) <> "" Then Kill sFilePath
  Err.Clear
  
End Sub



Private Sub SetExcel_bef380(ByVal pv_rsOrd As MDORowSet)
Dim sFile, sFilePath As String
Dim wOdSr As Integer
Dim wOdRow As Integer, wOrMRow As Integer, wOrDCRow As Integer, wOrXRow As Integer, wOlRow As Integer, wOmRow As Integer, wLn1 As Integer, wLn2 As Integer
Dim wCoName As String, wSht As Integer
Dim wPictPath As String
Dim wMetLs As Single, wGmFactor As Double, wCmPurityRt As Double
Dim wGldAs As String, wGldAsWt As String


  Set oExcel = CreateObject("Excel.Application")
  
  sFile = gs_CoCd + "-" + gs_UsrCd + "-" + Format(Now, "ddMMyyHHmmss") + ".xls"
  sFilePath = App.Path + "\Output\" + sFile
  
  If Dir(sFilePath, vbNormal) <> "" Then
    Kill sFilePath
  End If
  
  FileCopy App.Path + "\Template\OrderPrint.xls", sFilePath
  
  oExcel.Visible = False: oExcel.DisplayAlerts = False
  
  wCoName = moCn.GetFldVal("Select HName from Head Where HCoCd='" + gs_CoCd + "' and HCd='" + ctSelfCmCd + "'")
  
  wOdSr = 0: wOdRow = 14: wOmRow = 1: wSht = 1
  
  Set oBook = oExcel.Workbooks.Open(sFilePath)
  
  With pv_rsOrd
    Do While Not (.EOF Or .BOF)
      If wOdRow <= 285 Then
        If !OdSr <> wOdSr Then
        
          'wiSht = oExcel.Worksheets.Count
          'oExcel.Worksheets(2).Copy after:=oExcel.ActiveSheet
          oExcel.worksheets(2).Copy after:=oBook.worksheets(wSht + 1)
    
          Set oSheet = oBook.worksheets(1)
          
          If wOmRow = 1 Then
            oSheet.Name = !OmCoCd + !OmYy + !OmTc + !OmChr + CStr(!OmNo)
            oSheet.Cells(1, 1) = wCoName
            oSheet.Cells(5, 3) = !OmCoCd + "/" + !OmYy + "/" + !OmTc + "/" + !OmChr + "/" + CStr(!OmNo)
            oSheet.Cells(5, 6) = !OmDt
            oSheet.Cells(5, 9) = !OmLmgSal
            oSheet.Cells(5, 13) = !OmLmsSal
            oSheet.Cells(6, 9) = !OmLmpSal
            oSheet.Cells(6, 13) = !OmLmlSal
            oSheet.Cells(6, 3) = !OmCmCd
            oSheet.Cells(7, 3) = !CmName
            oSheet.Cells(8, 3) = !OmPoNo
            oSheet.Cells(8, 6) = !OmPoDt
            oSheet.Cells(9, 3) = !OmDelDt
            oSheet.Cells(10, 1) = "Total (" + IIF(wb_InBaseCur, ws_BaseCurCd, ADC("UoCurCdFr")) + ")"
            oSheet.Cells(13, 8) = "Price (" + IIF(wb_InBaseCur, ws_BaseCurCd, ADC("UoCurCdFr")) + ")"
            oSheet.Cells(13, 9) = "Value (" + IIF(wb_InBaseCur, ws_BaseCurCd, ADC("UoCurCdFr")) + ")"
            
            If ADC("UoYn6") = "Y" Then
              oSheet.Cells(315, 1) = !OmCmDelTerms
              oSheet.Cells(320, 1) = !OmCmPayTerms
              oSheet.Cells(325, 1) = !OmSalRem
            Else
              oSheet.Range("A315", "D328").UnMerge
              oSheet.Range("A314", "D328") = ""
            End If
          End If
          oSheet.Cells(wOdRow, 1) = !OdSr
          oSheet.Cells(wOdRow, 2) = !OdDmCd
          oSheet.Cells(wOdRow, 4) = !OdSfx
          oSheet.Cells(wOdRow, 6) = !OdDmSz
          oSheet.Cells(wOdRow, 7) = !qOrdQty
          oSheet.Cells(wOdRow, 8) = !OdSalPrc
          oSheet.Cells(wOdRow, 9).Font.Color = vbBlack
          wOdRow = wOdRow + 1
          wOmRow = wOmRow + 1
          wSht = wSht + 1
        End If
        
        'Set oSheet = oBook.Worksheets(Worksheets.Count)
        Set oSheet = oBook.worksheets(wSht + 1)
        
        If !OdSr <> wOdSr Then
          oSheet.Name = !OdDmCd + "-" + "(" + CStr(!OdSr) + ")"
          oSheet.Cells(4, 3) = !OdDmCd
          oSheet.Cells(4, 6) = !OdSfx
          oSheet.Cells(4, 9) = !OdDmSz
          oSheet.Cells(5, 3) = !CdOurDmCd
          oSheet.Cells(13, 13) = !OdDmCd
          oSheet.Cells(14, 13) = !IdKt
          oSheet.Cells(14, 15) = !OdDmCol
          oSheet.Cells(4, 13) = !OdPicNm
          oSheet.Cells(6, 3) = !qOrdQty
          oSheet.Cells(7, 3) = !OdPoNo
          oSheet.Cells(8, 3) = !OdDelDt
          oSheet.Cells(9, 7) = !OdSalPrc
          oSheet.Cells(51, 5) = !OdLabAsWt
          oSheet.Cells(63, 10) = !OdMulBy
          
          wOrMRow = 18: wOrDCRow = 26: wOrXRow = 45: wOlRow = 53
          
          oSheet.Cells(9, 1) = "Price (" + IIF(wb_InBaseCur, ws_BaseCurCd, ADC("UoCurCdFr")) + ")"
          oSheet.Cells(9, 5) = "Sales Price (" + IIF(wb_InBaseCur, ws_BaseCurCd, ADC("UoCurCdFr")) + ")"
          oSheet.Cells(wOrMRow - 1, 8) = "LME (" + IIF(wb_InBaseCur, ws_BaseCurCd, ADC("UoCurCdFr")) + ")"
          oSheet.Cells(wOrMRow - 1, 9) = "Rate (" + IIF(wb_InBaseCur, ws_BaseCurCd, ADC("UoCurCdFr")) + ")"
          oSheet.Cells(wOrMRow - 1, 10) = "Value (" + IIF(wb_InBaseCur, ws_BaseCurCd, ADC("UoCurCdFr")) + ")"
          oSheet.Cells(wOrDCRow - 1, 10) = "Rate (" + IIF(wb_InBaseCur, ws_BaseCurCd, ADC("UoCurCdFr")) + ")"
          oSheet.Cells(wOrDCRow - 1, 11) = "Value (" + IIF(wb_InBaseCur, ws_BaseCurCd, ADC("UoCurCdFr")) + ")"
          oSheet.Cells(wOrDCRow - 1, 13) = "Rate (" + IIF(wb_InBaseCur, ws_BaseCurCd, ADC("UoCurCdFr")) + ")"
          oSheet.Cells(wOrDCRow - 1, 14) = "Value (" + IIF(wb_InBaseCur, ws_BaseCurCd, ADC("UoCurCdFr")) + ")"
          oSheet.Cells(wOrXRow - 1, 8) = "Rate (" + IIF(wb_InBaseCur, ws_BaseCurCd, ADC("UoCurCdFr")) + ")"
          oSheet.Cells(wOrXRow - 1, 9) = "Value (" + IIF(wb_InBaseCur, ws_BaseCurCd, ADC("UoCurCdFr")) + ")"
          oSheet.Cells(wOlRow - 1, 6) = "Rate (" + IIF(wb_InBaseCur, ws_BaseCurCd, ADC("UoCurCdFr")) + ")"
          oSheet.Cells(wOlRow - 1, 7) = "Value (" + IIF(wb_InBaseCur, ws_BaseCurCd, ADC("UoCurCdFr")) + ")"
          oSheet.Cells(61, 5) = "Value (" + IIF(wb_InBaseCur, ws_BaseCurCd, ADC("UoCurCdFr")) + ")"
          oSheet.Cells(61, 10) = "Value (" + IIF(wb_InBaseCur, ws_BaseCurCd, ADC("UoCurCdFr")) + ")"
          
          If ADC("UoYn") = "Y" Then
            ' **** Manali 3.6.1 - Pic Reference and Dm/SM
            wPictPath = GetPictPath(IIF(UCase(ADC("UoTcTypFr")) = "QS", "SM", "DM"), IIF(UCase(ADC("UoPMCdFr")) = "3D", False, True), !OdDmCd, !OdDmSz, !OmCoCd, !OmTc, !OmYy, !OmChr, !OmNo, !OdSr, !OdPicNm)
            If wPictPath = "" Or Dir(wPictPath) = "" Then
              wPictPath = GetPictPath(IIF(UCase(ADC("UoTcTypFr")) = "QS", "SM", "DM"), IIF(UCase(ADC("UoPMCdFr")) = "3D", True, False), !OdDmCd, !OdDmSz, !OmCoCd, !OmTc, !OmYy, !OmChr, !OmNo, !OdSr, !OdPicNm)
            End If
            'wPictPath = GetPictPath("DM", False, !OdDmCd, !OdDmSz, !OmCoCd, !OmTc, !OmYy, !OmChr, !OmNo, !OdSr, !OdPicNm)
            If Dir(wPictPath) <> "" And wPictPath <> "" Then
              oSheet.Pictures.Insert(wPictPath).Select
              oSheet.Pictures(1).Name = oSheet.Name
              oSheet.Pictures(1).Height = 110: oSheet.Pictures(1).Width = 140
              oSheet.Pictures(1).Left = 572: oSheet.Pictures(1).Top = 45
              oSheet.Pictures(1).Border.LineStyle = 1
            End If
          End If
          
          If ADC("UoYn6") = "Y" Then
            oSheet.Cells(72, 1) = !OdDmPrdInst
            oSheet.Cells(79, 1) = !OdCmPrdInst
            oSheet.Cells(86, 1) = !OdCmStmpInst
            oSheet.Cells(91, 1) = !OdSzInst
            oSheet.Cells(96, 1) = !OdSalRem
          Else
            oSheet.Range("A71", "G100").UnMerge
            oSheet.Range("A71", "G100") = ""
          End If
        End If
          
          If !qTag = "1RM" Then
            Select Case !RmCtg
            Case Is = "G", "P", "S", "L"
              If wOrMRow <= 21 Then
                oSheet.Cells(wOrMRow, 1) = !OrSrNo
                oSheet.Cells(wOrMRow, 2) = !RmCtg
                oSheet.Cells(wOrMRow, 3) = !OrRmSCtg
                oSheet.Cells(wOrMRow, 4) = !OrRmCd
                oSheet.Cells(wOrMRow, 5) = !OrQty
                oSheet.Cells(wOrMRow, 6) = !OrWt
                oSheet.Cells(wOrMRow, 7) = !RmQw
                oSheet.Cells(wOrMRow, 8) = !OrLMESal
                
                wMetLs = moCn.GetFldVal("Select  Top 1 RrSalRt from RmRt " + _
                          "where RrTcTyp='LS' and RrCtg='" + !RmCtg + "' and RrCmCurCd='' and RrDmCd='' and RrSTWGrp='' and " + _
                          "((RrCmCtg = 'P' and RrCmCd = '" + !CmLkUpMetLs + "') Or (RrCmCtg='" + gs_CmCtg + "' and RrCmCd in ('" + !OmCmCd + "', '" + ctSelfCmCd + "'))) " + _
                          "and ((RrDmCtg = '" + !DmCtg + "' and RrLsCtg In ('" + !DmLsCtg + "', '')) Or (RrDmCtg='' and RrLsCtg='')) " + _
                          "Order By (Case When (RrCmCtg='P' and RrCmCd = '" + !CmLkUpMetLs + "') Then 1 " + _
                          "When RrCmCtg='" + gs_CmCtg + "' Then (Case RrCmCd When '" + !OmCmCd + "' Then 2 When '" + ctSelfCmCd + "' Then 3  End)  End), " + _
                          "(Case RrDmCtg When '" + !DmCtg + "' Then (Case RrLsCtg When '" + !DmLsCtg + "' Then 1 When '' Then 2  End) When '' Then 3  End) ")
  
                
                wGmFactor = moCn.GetFldVal("Select PNum from Param Where PTyp='GMFACTOR' and PMCd='" + !RmCtg + "' and PSCd='" + IIF(wb_InBaseCur, ws_BaseCurCd, ADC("UoCurCdFr")) + "'")
                wCmPurityRt = moCn.GetFldVal("Select RrSalRt from RmRt Where RrTcTyp='CRP' and RrCmCtg='" + gs_CmCtg + "' and RrCmCd='" + !OmCmCd + "' and RrCd='" + !OrRmCd + "' ")
                wCmPurityRt = IIF(wCmPurityRt = 0, !RmPurityRt, wCmPurityRt)
                
                oSheet.Cells(wOrMRow, 9).formula = "=(H" + CStr(wOrMRow) + "/" + CStr(wGmFactor) + "*" + CStr(wCmPurityRt) + "*(1 + " + CStr(wMetLs) + "/100)) + " + _
                                                "((1-" + CStr(wCmPurityRt) + ")*" + CStr(!OrAlySalRt) + ")"
  
                ' ***** Manali 3.6.1 - GldAs Wt
                wGldAsWt = "$F" + CStr(wOrMRow)
                If !OrMainMet = "Y" Then
                  wGldAs = moCn.GetFldVal("Select PDesc225 from Param Where Ptyp='GLDAS' and PMCd='" + !OdGldAs + "'")
                  
                  If InStr(1, "," + wGldAs + ",", ",D,") <> 0 Then wGldAsWt = wGldAsWt + IIF(wGldAsWt <> "", "+", "") + "(D63/5)"
                  If InStr(1, "," + wGldAs + ",", ",C,") <> 0 Then wGldAsWt = wGldAsWt + IIF(wGldAsWt <> "", "+", "") + "(D64/5)"
                  If InStr(1, "," + wGldAs + ",", ",X,") <> 0 Then wGldAsWt = wGldAsWt + IIF(wGldAsWt <> "", "+", "") + "(D65)"
                  
                End If
                wGldAsWt = "(" + wGldAsWt + ")"
                                        
                oSheet.Cells(wOrMRow, 10).formula = "=$I" + CStr(wOrMRow) + "*(IF($G" + CStr(wOrMRow) + "=""Q"",$E" + CStr(wOrMRow) + ",(IF($G" + CStr(wOrMRow) + "=""W""," + wGldAsWt + ",0))))"
                oSheet.Cells(wOrMRow, 10).Font.Color = vbBlack
              ElseIf wOrMRow = 22 Then
                oSheet.Cells(16, 1).AddComment "#Metal Details InComplete"
              End If
              wOrMRow = wOrMRow + 1
              
            Case Is = "D", "C"
              If wOrDCRow <= 40 Then
                oSheet.Cells(wOrDCRow, 1) = !OrSrNo
                oSheet.Cells(wOrDCRow, 2) = !RmCtg
                oSheet.Cells(wOrDCRow, 3) = !OrRmSCtg
                oSheet.Cells(wOrDCRow, 4) = !OrRmCd
                oSheet.Cells(wOrDCRow, 5) = !OrLn1
                oSheet.Cells(wOrDCRow, 6) = !OrRmPtr
                oSheet.Cells(wOrDCRow, 7) = !OrQty
                oSheet.Cells(wOrDCRow, 8) = !OrWt
                oSheet.Cells(wOrDCRow, 9) = !RmQw
                oSheet.Cells(wOrDCRow, 10) = !OrSalRt
                oSheet.Cells(wOrDCRow, 11).Font.Color = vbBlack
                oSheet.Cells(wOrDCRow, 12) = !OrSetSCd
                oSheet.Cells(wOrDCRow, 13) = !OrSetSalRt
                oSheet.Cells(wOrDCRow, 14).Font.Color = vbBlack
              ElseIf wOrDCRow = 41 Then
                oSheet.Cells(24, 1).AddComment "#Studding Details Incomplete"
              End If
              wOrDCRow = wOrDCRow + 1
            
            Case Is = "X"
              If wOrXRow <= 48 Then
                oSheet.Cells(wOrXRow, 1) = !OrSrNo
                oSheet.Cells(wOrXRow, 2) = !RmCtg
                oSheet.Cells(wOrXRow, 3) = !OrRmSCtg
                oSheet.Cells(wOrXRow, 4) = !OrRmCd
                oSheet.Cells(wOrXRow, 5) = !OrQty
                oSheet.Cells(wOrXRow, 6) = !OrWt
                oSheet.Cells(wOrXRow, 7) = !RmQw
                oSheet.Cells(wOrXRow, 8) = !OrSalRt
                oSheet.Cells(wOrXRow, 9).Font.Color = vbBlack
              ElseIf wOrXRow = 49 Then
                oSheet.Cells(43, 1).AddComment = "#Acc Details InComplete"
              End If
              wOrXRow = wOrXRow + 1
            End Select
          
          ElseIf !qTag = "2LAB" Then
              If wOlRow <= 57 Then
                oSheet.Cells(wOlRow, 1) = !OlSrNo
                oSheet.Cells(wOlRow, 2) = !OlMCd
                oSheet.Cells(wOlRow, 3) = !OlSCd
                oSheet.Cells(wOlRow, 4) = !OlQw
                oSheet.Cells(wOlRow, 5) = !OlQty
                oSheet.Cells(wOlRow, 6) = !OlSalRt
                ' ***** Manali 3.6.1 - Lab Val based on min/max val or OlVal
                oSheet.Cells(wOlRow, 7) = "=IF($F" + CStr(wOlRow) + " * (IF($D" + CStr(wOlRow) + "=""Q"",$E" + CStr(wOlRow) + ",(IF($D" + CStr(wOlRow) + "=""W"",$E$51,0))))=" + IIF(!OlQw = "Q", CStr(!OlQty), CStr(!OdLabAsWt)) + "*" + CStr(!OlSalRt) + "," + _
                          CStr(!OlSalVal) + ", $F" + CStr(wOlRow) + " * (IF($D" + CStr(wOlRow) + "=""Q"",$E" + CStr(wOlRow) + ",(IF($D" + CStr(wOlRow) + "=""W"",$E$51,0)))))"
                oSheet.Cells(wOlRow, 7).Font.Color = vbBlack
              ElseIf wOlRow = 58 Then
                oSheet.Cells(51, 1).AddComment = "#Labour Details InComplete"
              End If
              wOlRow = wOlRow + 1
          End If
        
        wOdSr = !OdSr
        .MoveNext
      
      Else: Exit Do
      End If
    Loop
  End With
  Set oSheet = oBook.worksheets(1)
  oSheet.Rows(CStr(wOdRow) + ":299").Hidden = True
  
  oExcel.worksheets(2).Delete

  Set pv_rsOrd = Nothing
  
  ' ***** Manali 3.9.0 - Excel For 2007 n next
  Dim wExcVer As String
  wExcVer = oExcel.version
  
  
  Cd1.DialogTitle = "Save Excel File As"
  Cd1.FileName = ADC("UoOmTcFr") + "-" + ADC("UoOmYyFr") + "-" + ADC("UoOmChrFr") + "-" + CStr(ADC("UoOmNoFr")) + ".xls"
  ' ***** Manali 3.9.0 - Excel For 2007 n next
  Cd1.Filter = "Excel Files (*.xls" + IIF(wExcVer >= "12.0", ";*.xlsx", "") + ") |*.xls" + IIF(wExcVer >= "12.0", ";*.xlsx", "") + ""
  ' Bef 3.9.0
  'CD1.Filter = "Excel Files (*.xls) |*.xls"
  
  Cd1.Flags = cdlOFNOverwritePrompt Or cdlOFNLongNames Or cdlOFNPathMustExist Or cdlOFNExtensionDifferent
  
  Cd1.CancelError = True
  Cd1.ShowSave

  ' bef  3.9.0
''  If Len(CD1.FileName) > 4 Then
''    If UCase(Right(CD1.FileName, 4)) = UCase(".xls") Then
  
  ' ***** Manali 3.9.0 - Excel For 2007 n next
  If Len(Cd1.FileName) > 5 Then
    If UCase(Right(Cd1.FileName, 4)) = UCase(".xls") Or (wExcVer >= "12.0" And UCase(Right(Cd1.FileName, 5)) = UCase(".xlsx")) Then
      'oExcel.ActiveWorkbook.SaveAs Cd1.FileName
      oExcel.ActiveWorkbook.SaveAs Cd1.FileName, , "", ""       '4.1.0.0 (Pwd Removed)
    Else
      wErrMsg = "File Extension Is Wrong": Exit Sub
    End If
  End If
  
  If oExcel.ActiveWorkbook.Path + "\" + oExcel.ActiveWorkbook.Name <> sFilePath Then
    Kill sFilePath
  End If
  
  DispMsg "Excel Generated Successfully !!", etInfo
  oExcel.ActiveWorkbook.Close True
  oExcel.QUIT
  
  oExcel.DisplayAlerts = True
  Set oExcel = Nothing
  Set oBook = Nothing
  Set oSheet = Nothing
Exit Sub
Err_Hndlr:
  If Err.Number = 32755 Then
    wErrMsg = "Operation Cancelled"
    oExcel.ActiveWorkbook.Close
    oExcel.QUIT
  Else
    wErrMsg = Err.Description
  End If
  If Dir(sFilePath) <> "" Then Kill sFilePath
  Err.Clear
  
End Sub







