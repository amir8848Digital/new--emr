VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{AB2CC32A-E652-42A2-B341-26ECE02F6386}#2.1#0"; "mwfctlforspeed9.ocx"
Begin VB.Form EmrFrmOrdZoom 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Order Zoom"
   ClientHeight    =   9810
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15210
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   13440
   ScaleWidth      =   25890
   WindowState     =   2  'Maximized
   Begin VB.Frame FRA_KEYS 
      BorderStyle     =   0  'None
      Height          =   9255
      Left            =   0
      TabIndex        =   43
      Top             =   0
      Width           =   15210
      Begin VB.OptionButton OptDelDt 
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
         Height          =   285
         Index           =   1
         Left            =   4440
         TabIndex        =   40
         ToolTipText     =   "Select Export Delivery Date Or Production Delivery Date"
         Top             =   4275
         Width           =   1725
      End
      Begin VB.OptionButton OptDelDt 
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
         Height          =   285
         Index           =   0
         Left            =   2880
         TabIndex        =   39
         ToolTipText     =   "Select Export Delivery Date Or Production Delivery Date"
         Top             =   4275
         Width           =   1695
      End
      Begin VB.CheckBox ChkBoxArr 
         Alignment       =   1  'Right Justify
         Caption         =   "Sort By Dsg Cd                       "
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
         TabIndex        =   38
         Tag             =   "SortDmCd"
         ToolTipText     =   "Check To Show Sort By Design Cd"
         Top             =   3990
         Width           =   3045
      End
      Begin VB.CheckBox ChkBoxArr 
         Alignment       =   1  'Right Justify
         Caption         =   "Show Value                         "
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
         Left            =   30
         TabIndex        =   36
         Tag             =   "ShowVal"
         ToolTipText     =   "Check To Show Value"
         Top             =   3705
         Width           =   3045
      End
      Begin VB.CheckBox ChkBoxArr 
         Alignment       =   1  'Right Justify
         Caption         =   "Show Pending                     "
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
         TabIndex        =   31
         Tag             =   "ShowPnd"
         ToolTipText     =   "Check To Show Pending Records Only"
         Top             =   2850
         Width           =   3045
      End
      Begin VB.TextBox TxtPwd 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   2880
         MaxLength       =   30
         PasswordChar    =   "*"
         TabIndex        =   34
         Top             =   3420
         Width           =   2160
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   0
         Left            =   2880
         TabIndex        =   2
         ToolTipText     =   "Enter Fr Voucher Tc"
         Top             =   270
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   503
         MaxLength       =   3
         DataType        =   4
         IdName          =   "WODTCFR"
         CmpStr          =   "OmTc>="
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   1
         Left            =   3660
         TabIndex        =   3
         ToolTipText     =   "Enter Fr Voucher Year"
         Top             =   270
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   503
         MaxLength       =   2
         DataType        =   4
         IdName          =   "WODYYFR"
         CmpStr          =   "OmYy>="
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   2
         Left            =   4110
         TabIndex        =   4
         ToolTipText     =   "Enter Fr Voucher Chr"
         Top             =   270
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   503
         MaxLength       =   3
         DataType        =   4
         IdName          =   "WODCHRFR"
         CmpStr          =   "OmChr>="
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   3
         Left            =   6810
         TabIndex        =   7
         ToolTipText     =   "Enter To Voucher Tc "
         Top             =   270
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   503
         MaxLength       =   3
         DataType        =   4
         IdName          =   "WODTCTO"
         CmpStr          =   "OmTc<="
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   7
         Left            =   5700
         TabIndex        =   6
         ToolTipText     =   "Enter Fr Voucher Sr"
         Top             =   270
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   503
         Alignment       =   2
         Mask            =   "#####0"
         MaxLength       =   6
         DataType        =   1
         IdName          =   "WODSRFR"
         CmpStr          =   "OdSr>="
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   6
         Left            =   4850
         TabIndex        =   5
         ToolTipText     =   "Enter Fr Voucher No"
         Top             =   270
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   503
         Alignment       =   2
         Mask            =   "#####0"
         MaxLength       =   6
         DataType        =   1
         IdName          =   "WODNOFR"
         CmpStr          =   "OmNo>="
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   4
         Left            =   7590
         TabIndex        =   8
         ToolTipText     =   "Enter To Voucher Year"
         Top             =   270
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   503
         MaxLength       =   2
         DataType        =   4
         IdName          =   "WODYYTO"
         CmpStr          =   "OmYy<="
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   5
         Left            =   8040
         TabIndex        =   9
         ToolTipText     =   "Enter To Voucher Chr"
         Top             =   270
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   503
         MaxLength       =   3
         DataType        =   4
         IdName          =   "WODCHRTO"
         CmpStr          =   "OmChr<="
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   9
         Left            =   9710
         TabIndex        =   11
         ToolTipText     =   "Enter To Voucher Sr"
         Top             =   270
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   503
         Alignment       =   2
         Mask            =   "#####0"
         MaxLength       =   6
         DataType        =   1
         IdName          =   "WODSRTO"
         CmpStr          =   "OdSr<="
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   8
         Left            =   8820
         TabIndex        =   10
         ToolTipText     =   "EnterTo Voucher No"
         Top             =   270
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   503
         Alignment       =   2
         Mask            =   "#####0"
         MaxLength       =   6
         DataType        =   1
         IdName          =   "WODNOTO"
         CmpStr          =   "OmNo<="
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   18
         Left            =   2880
         TabIndex        =   18
         ToolTipText     =   "Enter Document No  From Scope"
         Top             =   1410
         Width           =   3105
         _ExtentX        =   5477
         _ExtentY        =   503
         MaxLength       =   20
         IdName          =   "WOMPONOFR"
         CmpStr          =   "OMPONO>="
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   19
         Left            =   6810
         TabIndex        =   19
         ToolTipText     =   "Enter Document No To Scope"
         Top             =   1410
         Width           =   3105
         _ExtentX        =   5477
         _ExtentY        =   503
         MaxLength       =   20
         IdName          =   "WOMPONOTO"
         CmpStr          =   "OMPONO<="
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   20
         Left            =   6810
         TabIndex        =   25
         ToolTipText     =   "Enter To Design Category"
         Top             =   1995
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   503
         MaxLength       =   5
         DataType        =   4
         IdName          =   "WODDSGCTGTO"
         CmpStr          =   "DmCtg<="
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   21
         Left            =   2880
         TabIndex        =   24
         ToolTipText     =   "Enter From Design Category"
         Top             =   1995
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   503
         MaxLength       =   5
         DataType        =   4
         IdName          =   "WODDSGCTGFR"
         CmpStr          =   "DmCtg>="
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   22
         Left            =   2880
         TabIndex        =   26
         ToolTipText     =   "Enter From Dsg Cd"
         Top             =   2280
         Width           =   2355
         _ExtentX        =   4154
         _ExtentY        =   503
         MaxLength       =   15
         DataType        =   4
         IdName          =   "WODDMCDFR"
         CmpStr          =   "OdDmCd>="
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   23
         Left            =   6810
         TabIndex        =   27
         ToolTipText     =   "Enter To Dsg Cd"
         Top             =   2280
         Width           =   2385
         _ExtentX        =   4207
         _ExtentY        =   503
         MaxLength       =   15
         DataType        =   4
         IdName          =   "WODDMCDTO"
         CmpStr          =   "OdDmCd<="
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   24
         Left            =   3420
         TabIndex        =   30
         ToolTipText     =   "Enter 'Y' to Show Only Pending Orders/ 'N'  To show all orders"
         Top             =   2850
         Visible         =   0   'False
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   503
         Enabled         =   0   'False
         DataType        =   4
         IdName          =   "WYN"
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   37
         Left            =   6810
         TabIndex        =   13
         ToolTipText     =   "Enter To Order Date"
         Top             =   555
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   503
         MaxLength       =   8
         DataType        =   3
         IdName          =   "WOMDTTO"
         CmpStr          =   "OmDt <="
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   48
         Left            =   2880
         TabIndex        =   12
         ToolTipText     =   "Enter From Order  Date"
         Top             =   555
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   503
         MaxLength       =   8
         DataType        =   3
         IdName          =   "WOMDTFR"
         CmpStr          =   "OmDt >="
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   38
         Left            =   2880
         TabIndex        =   16
         ToolTipText     =   "Enter From Customer (Customer As Mentioned In Transactions)"
         Top             =   1125
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   503
         MaxLength       =   8
         DataType        =   4
         IdName          =   "WOMCMCDFR"
         CmpStr          =   "OmCmCd >="
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   39
         Left            =   6810
         TabIndex        =   17
         ToolTipText     =   "Enter To Customer (Customer As Mentioned In Transactions)"
         Top             =   1125
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   503
         MaxLength       =   8
         DataType        =   4
         IdName          =   "WOMCMCDTO"
         CmpStr          =   "OmCmCd <="
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   15
         Left            =   6810
         TabIndex        =   15
         ToolTipText     =   "Enter To Delivery Date"
         Top             =   840
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   503
         MaxLength       =   8
         DataType        =   3
         IdName          =   "WOMDELDTTO"
         CmpStr          =   "OdDelDt <="
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   16
         Left            =   2880
         TabIndex        =   14
         ToolTipText     =   "Enter From Delivery  Date"
         Top             =   840
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   503
         MaxLength       =   8
         DataType        =   3
         IdName          =   "WOMDELDTFR"
         CmpStr          =   "OdDelDt >="
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   17
         Left            =   2880
         TabIndex        =   32
         ToolTipText     =   "Specify Bag Type or ' ' for All"
         Top             =   3135
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   503
         DataType        =   4
         ReCalcParent    =   "WBAGCHR"
         IdName          =   "WBAGTYP"
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   25
         Left            =   4320
         TabIndex        =   33
         ToolTipText     =   "Enter Bag Character"
         Top             =   3135
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   503
         MaxLength       =   3
         DataType        =   4
         ReCalcOn        =   "WYNFLTBAG"
         IdName          =   "WBAGCHR"
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   42
         Left            =   3420
         TabIndex        =   35
         ToolTipText     =   "Enter 'Y' to Show Value"
         Top             =   3705
         Visible         =   0   'False
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   503
         Enabled         =   0   'False
         DataType        =   4
         IdName          =   "WPWDYN"
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   29
         Left            =   3420
         TabIndex        =   37
         ToolTipText     =   "Enter 'Y' to Show Only Pending Orders/ 'N'  To show all orders"
         Top             =   3990
         Visible         =   0   'False
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   503
         Enabled         =   0   'False
         DataType        =   4
         IdName          =   "WSORTDMCD"
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   31
         Left            =   2880
         TabIndex        =   0
         ToolTipText     =   "Enter Company Code"
         Top             =   0
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   503
         MaxLength       =   3
         DataType        =   4
         DataField       =   "UOCOCDFR"
         IdName          =   "WCOCDFR"
         CmpStr          =   "OmCoCd >= "
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   47
         Left            =   6810
         TabIndex        =   1
         ToolTipText     =   "Enter To Company Code"
         Top             =   0
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   503
         MaxLength       =   3
         DataType        =   4
         DataField       =   "UOCOCDTO"
         IdName          =   "WCOCDTO"
         CmpStr          =   "OmCoCd <= "
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   32
         Left            =   6240
         TabIndex        =   41
         ToolTipText     =   "Show Export/ Production Delivery Date (Enter ""Y"" for Export and ""N"" for Production)"
         Top             =   4275
         Visible         =   0   'False
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   503
         Enabled         =   0   'False
         DataType        =   4
         IdName          =   "WDELDT"
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   33
         Left            =   2880
         TabIndex        =   28
         ToolTipText     =   "Enter From Cust Dsg Cd (Filter on Blank Suffix, Blank Size record of Customer Design table)"
         Top             =   2565
         Width           =   3500
         _ExtentX        =   6165
         _ExtentY        =   503
         MaxLength       =   40
         DataType        =   4
         IdName          =   "WCDCUSTDMCDFR"
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   34
         Left            =   6810
         TabIndex        =   29
         ToolTipText     =   "Enter To Cust Dsg Cd (Enter From Cust Dsg Cd (Filter on Blank Suffix, Blank Size record of Customer Design table))"
         Top             =   2565
         Width           =   3500
         _ExtentX        =   6165
         _ExtentY        =   503
         MaxLength       =   40
         DataType        =   4
         IdName          =   "WCDCUSTDMCDTO"
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   26
         Left            =   4200
         TabIndex        =   21
         ToolTipText     =   "Enter From Sub PO"
         Top             =   1700
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   503
         MaxLength       =   20
         DataType        =   4
         IdName          =   "WODPONOFR"
         CmpStr          =   "OdPoNo >="
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   85
         Left            =   2880
         TabIndex        =   20
         ToolTipText     =   "Enter From Sub Customer"
         Top             =   1700
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         MaxLength       =   8
         DataType        =   4
         IdName          =   "WODSUBCUSTFR"
         CmpStr          =   "OdSubCust >="
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   27
         Left            =   8130
         TabIndex        =   23
         ToolTipText     =   "Enter To Sub PO"
         Top             =   1700
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   503
         MaxLength       =   20
         DataType        =   4
         IdName          =   "WODPONOTO"
         CmpStr          =   "OdPoNo <="
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   86
         Left            =   6810
         TabIndex        =   22
         ToolTipText     =   "Enter To Sub Customer"
         Top             =   1700
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         MaxLength       =   8
         DataType        =   4
         IdName          =   "WODSUBCUSTTO"
         CmpStr          =   "OdSubCust <="
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
         Index           =   6
         Left            =   8025
         TabIndex        =   210
         Top             =   1710
         Width           =   60
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
         Index           =   5
         Left            =   4110
         TabIndex        =   209
         Top             =   1710
         Width           =   60
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cust Dsg Cd"
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
         Left            =   60
         TabIndex        =   208
         Top             =   2565
         Width           =   1125
      End
      Begin VB.Label LblPRdExpDt 
         BackStyle       =   0  'Transparent
         Caption         =   "Consider Delivery Date"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   30
         TabIndex        =   207
         ToolTipText     =   "Location"
         Top             =   4275
         Width           =   2205
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Company Code"
         BeginProperty Font 
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
         Left            =   60
         TabIndex        =   206
         Top             =   0
         Width           =   2145
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
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
         Height          =   240
         Index           =   9
         Left            =   60
         TabIndex        =   191
         Top             =   1710
         Width           =   1035
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
         Index           =   3
         Left            =   60
         TabIndex        =   90
         Top             =   3420
         Width           =   1455
      End
      Begin VB.Label LblBagTyp 
         BackStyle       =   0  'Transparent
         Caption         =   "Bag Type"
         BeginProperty Font 
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
         ToolTipText     =   "Location"
         Top             =   3135
         Width           =   1035
      End
      Begin VB.Label LblUoBChrFr 
         BackStyle       =   0  'Transparent
         Caption         =   "Bag Chr"
         BeginProperty Font 
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
         Left            =   3480
         TabIndex        =   86
         Top             =   3135
         Width           =   885
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Delivery Date"
         BeginProperty Font 
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
         Left            =   60
         TabIndex        =   81
         Top             =   840
         Width           =   1455
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
         Height          =   285
         Index           =   2
         Left            =   60
         TabIndex        =   64
         Top             =   555
         Width           =   1185
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Cust Cd"
         BeginProperty Font 
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
         TabIndex        =   63
         Top             =   1125
         Width           =   825
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
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
         Height          =   285
         Index           =   1
         Left            =   60
         TabIndex        =   62
         Top             =   1410
         Width           =   645
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dsg Ctg"
         BeginProperty Font 
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
         Left            =   60
         TabIndex        =   61
         Top             =   1995
         Width           =   720
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dsg Cd"
         BeginProperty Font 
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
         Left            =   60
         TabIndex        =   60
         Top             =   2280
         Width           =   660
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
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
         Height          =   240
         Index           =   3
         Left            =   6480
         TabIndex        =   59
         Top             =   0
         Width           =   240
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fr"
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
         Left            =   2520
         TabIndex        =   58
         Top             =   0
         Width           =   195
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
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
         Height          =   240
         Index           =   0
         Left            =   9615
         TabIndex        =   55
         Top             =   270
         Width           =   60
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
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
         Height          =   240
         Index           =   2
         Left            =   8700
         TabIndex        =   54
         Top             =   270
         Width           =   60
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
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
         Height          =   240
         Left            =   7470
         TabIndex        =   53
         Top             =   270
         Width           =   60
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
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
         Height          =   240
         Left            =   7920
         TabIndex        =   52
         Top             =   270
         Width           =   60
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
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
         Height          =   240
         Left            =   5645
         TabIndex        =   51
         Top             =   270
         Width           =   60
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
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
         Height          =   240
         Index           =   1
         Left            =   4770
         TabIndex        =   50
         Top             =   270
         Width           =   60
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
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
         Height          =   240
         Index           =   0
         Left            =   3990
         TabIndex        =   48
         Top             =   270
         Width           =   60
      End
      Begin VB.Label LblSlash 
         AutoSize        =   -1  'True
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
         Height          =   240
         Left            =   3540
         TabIndex        =   47
         Top             =   270
         Width           =   60
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vch Tc/ Yy/ Chr/ No/ Sr"
         BeginProperty Font 
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
         TabIndex        =   46
         Top             =   285
         Width           =   2145
      End
   End
   Begin MwfCtl.MWCTL_BTN1 CmdViewSel 
      Height          =   420
      Left            =   8205
      TabIndex        =   194
      Top             =   9270
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   741
      BackColor       =   -2147483633
      MaskColor       =   8454143
      Caption         =   "&View Sel"
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
   Begin MwfCtl.MWCTL_BTN1 CmdDsgDet 
      Height          =   420
      Left            =   6300
      TabIndex        =   89
      Top             =   9270
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   741
      BackColor       =   -2147483633
      MaskColor       =   8454143
      Caption         =   "DsgD&et"
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
   Begin MwfCtl.MWCTL_BTN1 cmd_InvDet 
      Height          =   420
      Left            =   7260
      TabIndex        =   45
      Top             =   9270
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   741
      BackColor       =   -2147483633
      MaskColor       =   8454143
      Caption         =   "&InvDet"
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
      Left            =   9750
      TabIndex        =   44
      TabStop         =   0   'False
      Top             =   9270
      Visible         =   0   'False
      Width           =   645
      _ExtentX        =   1138
      _ExtentY        =   767
   End
   Begin MwfCtl.MWCTL_ADC ADC 
      CausesValidation=   0   'False
      Height          =   555
      Left            =   30
      TabIndex        =   42
      Top             =   9270
      Width           =   14985
      _ExtentX        =   26432
      _ExtentY        =   979
      AllowAdd        =   0   'False
      AllowDelete     =   0   'False
      BtnWidth        =   885
      BtnHeight       =   435
      BtnMaskColor    =   8454143
      FontSz          =   8.25
   End
   Begin VB.Frame FraCpySel 
      Height          =   2415
      Left            =   7680
      TabIndex        =   197
      Top             =   6810
      Visible         =   0   'False
      Width           =   2655
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   2205
         Index           =   28
         Left            =   60
         TabIndex        =   198
         ToolTipText     =   "ZOOM Selection"
         Top             =   120
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   3889
         Alignment       =   3
         MaxLength       =   1000
         IdName          =   "WZOOMSEL"
      End
   End
   Begin VB.Frame FRA_NKEYS 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   9030
      Left            =   -150
      TabIndex        =   49
      Top             =   495
      Width           =   15315
      Begin VB.Frame FraNKeyAll 
         BorderStyle     =   0  'None
         Height          =   8745
         Left            =   60
         TabIndex        =   56
         Top             =   30
         Width           =   15210
         Begin MwfCtl.MWCTL_FLX GrdOrdWIW 
            Height          =   2520
            Left            =   150
            TabIndex        =   65
            Top             =   6180
            Width           =   2955
            _ExtentX        =   5212
            _ExtentY        =   4445
            Cols            =   4
            AllowAdd        =   0   'False
            AllowDelete     =   0   'False
            colname1        =   "WOWLOC"
            heading1        =   "Loc                   "
            maxlength1      =   5
            tooltiptext1    =   "Location Code"
            colname2        =   "WOWTOTBAGS"
            heading2        =   "Bags"
            datatype2       =   1
            mask2           =   "###0"
            maxlength2      =   4
            tooltiptext2    =   "Total No. Of Bags of the Order Sr. in the Location"
            colname3        =   "WOWPCS"
            heading3        =   "BQty      "
            datatype3       =   2
            mask3           =   "####0.0"
            maxlength3      =   7
            tooltiptext3    =   "Total No. of Pieces in all the Bags in the Loc"
         End
         Begin MwfCtl.MWCTL_FLX GrdOrdBagRm 
            Height          =   2520
            Left            =   10680
            TabIndex        =   203
            Top             =   6180
            Width           =   4515
            _ExtentX        =   7964
            _ExtentY        =   4445
            Cols            =   7
            AllowAdd        =   0   'False
            AllowDelete     =   0   'False
            colname1        =   "WOBRRMCTG"
            heading1        =   "Ctg"
            datatype1       =   4
            maxlength1      =   1
            style1          =   2
            colname2        =   "WOBRRMSCTG"
            heading2        =   "SCtg"
            datatype2       =   4
            maxlength2      =   5
            style2          =   2
            colname3        =   "WOBRRMCD"
            heading3        =   "Rm Cd                                      "
            datatype3       =   4
            maxlength3      =   16
            style3          =   2
            colname4        =   "WOBRRMSZ"
            datatype4       =   2
            mask4           =   "##0.0000"
            maxlength4      =   8
            style4          =   0
            colname5        =   "WOBRRMQTY"
            heading5        =   "Qty      "
            datatype5       =   1
            mask5           =   "####0"
            maxlength5      =   5
            style5          =   2
            colname6        =   "WOBRRMWT"
            heading6        =   "Rm Wt         "
            datatype6       =   2
            mask6           =   "####0.000"
            maxlength6      =   9
            style6          =   2
         End
         Begin VB.Frame FraPicHsk 
            BorderStyle     =   0  'None
            Height          =   1965
            Left            =   10710
            TabIndex        =   85
            Top             =   30
            Width           =   1965
            Begin VB.Image PicHsk 
               Height          =   1905
               Left            =   0
               Top             =   0
               Width           =   1905
            End
         End
         Begin VB.Frame FraPic3D 
            BorderStyle     =   0  'None
            Height          =   1965
            Left            =   12900
            TabIndex        =   84
            Top             =   30
            Width           =   1965
            Begin VB.Image Pic3D 
               Height          =   1905
               Left            =   30
               Top             =   0
               Width           =   1905
            End
         End
         Begin MwfCtl.MWCTL_FLX GrdOrdBag 
            Height          =   2145
            Left            =   3120
            TabIndex        =   66
            Top             =   6180
            Width           =   7545
            _ExtentX        =   13309
            _ExtentY        =   3784
            Cols            =   13
            AllowAdd        =   0   'False
            AllowDelete     =   0   'False
            colname1        =   "WOBYY"
            heading1        =   "BYy"
            datatype1       =   4
            maxlength1      =   2
            tooltiptext1    =   "Bag YY"
            colname2        =   "WOBCHR"
            heading2        =   "BChr   "
            datatype2       =   4
            maxlength2      =   3
            tooltiptext2    =   "Bag Chr"
            colname3        =   "WOBNO"
            heading3        =   "BNo           "
            datatype3       =   1
            mask3           =   "######0"
            maxlength3      =   7
            tooltiptext3    =   "Bag No"
            colname4        =   "WOBPCS"
            heading4        =   "BQty      "
            datatype4       =   2
            mask4           =   "####0.0"
            maxlength4      =   7
            tooltiptext4    =   "Number of Pieces of the Design In the BAg"
            colname5        =   "WOBPTPCS"
            heading5        =   "Parts"
            datatype5       =   2
            mask5           =   "####0"
            maxlength5      =   5
            style5          =   2
            tooltiptext5    =   "Number of Part Pieces of the Design In the BAg"
            colname6        =   "WOBGRSWT"
            heading6        =   "GrossWt     "
            datatype6       =   2
            mask6           =   "####0.##0"
            maxlength6      =   9
            tooltiptext6    =   "Grs Wt Of the Pieces in the Bag"
            colname7        =   "WOBRECVDT"
            heading7        =   "Rcvd Dt "
            datatype7       =   3
            maxlength7      =   8
            style7          =   2
            colname8        =   "WOBWRK"
            heading8        =   "Worker         "
            maxlength8      =   8
            style8          =   2
            colname9        =   "WOBWRKNAME"
            heading9        =   "Worker Name                         "
            maxlength9      =   20
            style9          =   0
            tooltiptext9    =   "Worker Name"
            colname10       =   "WOBREJYN"
            heading10       =   "Rej YN"
            maxlength10     =   1
            tooltiptext10   =   "Rejected Bag YN"
            colname11       =   "WOBFGSUBLOC"
            heading11       =   "Fg Sub Loc"
            maxlength11     =   8
            style11         =   2
            tooltiptext11   =   "Fg Sub Loc"
            colname12       =   "WOBCOCD"
            maxlength12     =   3
            style12         =   0
         End
         Begin MwfCtl.MWCTL_FLX GrdOrdPnd 
            Height          =   3735
            Left            =   150
            TabIndex        =   57
            Top             =   2430
            Width           =   15060
            _ExtentX        =   26564
            _ExtentY        =   6588
            Cols            =   26
            AllowAdd        =   0   'False
            AllowDelete     =   0   'False
            colname1        =   "WOPTC"
            maxlength1      =   3
            style1          =   0
            colname2        =   "WOPYY"
            maxlength2      =   2
            style2          =   0
            colname3        =   "WOPCHR"
            maxlength3      =   3
            style3          =   0
            colname4        =   "WOPNO"
            datatype4       =   1
            mask4           =   "#########0"
            maxlength4      =   10
            style4          =   0
            tooltiptext4    =   "Order No."
            colname5        =   "WOPSR"
            heading5        =   "Sr           "
            datatype5       =   1
            mask5           =   "#####0"
            maxlength5      =   6
            colname6        =   "WOPDSGCD"
            heading6        =   "Design   Code             "
            maxlength6      =   15
            tooltiptext6    =   "Design Code "
            colname7        =   "WOPSFX"
            heading7        =   "Suffix                  "
            maxlength7      =   15
            tooltiptext7    =   "Design Suffix"
            colname8        =   "WOPSZ"
            heading8        =   "Size      "
            maxlength8      =   5
            tooltiptext8    =   "Design Size"
            colname9        =   "WOPBALPCS"
            datatype9       =   2
            mask9           =   "####0.0"
            maxlength9      =   7
            style9          =   0
            tooltiptext9    =   "Number Of Pieces Balanced"
            colname10       =   "WOPORDQTY"
            heading10       =   "Ord Qty   "
            datatype10      =   1
            mask10          =   "#####0.0"
            maxlength10     =   8
            style10         =   2
            colname11       =   "WOPPRDQTY"
            heading11       =   "Prd Qty   "
            datatype11      =   1
            mask11          =   "#####0.0"
            maxlength11     =   8
            style11         =   2
            colname12       =   "WOPBALEXPQTY"
            heading12       =   "Bal Qty   "
            datatype12      =   1
            mask12          =   "#####0.0"
            maxlength12     =   8
            style12         =   2
            tooltiptext12   =   "Balance Export Qty"
            colname13       =   "WOPVALPPC"
            heading13       =   "ValuePpc "
            datatype13      =   2
            mask13          =   "#####0.#0"
            maxlength13     =   9
            tooltiptext13   =   "PerPc Val Of the Design"
            colname14       =   "WOPTOTVAL"
            heading14       =   "Bal Value     "
            datatype14      =   2
            mask14          =   "########0.#0"
            maxlength14     =   12
            tooltiptext14   =   "Total Value Of the Balance Pieces"
            colname15       =   "WOPPRTCD"
            heading15       =   "Prt "
            maxlength15     =   5
            tooltiptext15   =   "Order Priority"
            colname16       =   "WOPEXPDT"
            heading16       =   "Exp Dt     "
            datatype16      =   3
            maxlength16     =   8
            tooltiptext16   =   "Export Del Dt"
            colname17       =   "WOPKARAT"
            heading17       =   "Karat   "
            maxlength17     =   5
            tooltiptext17   =   "Design Karat"
            colname18       =   "WOPORDENT"
            datatype18      =   1
            mask18          =   "#####0"
            maxlength18     =   6
            style18         =   0
            colname19       =   "WOPPRDENT"
            datatype19      =   1
            mask19          =   "#####0"
            maxlength19     =   6
            style19         =   0
            colname20       =   "WOPDMCOL"
            maxlength20     =   5
            style20         =   0
            tooltiptext20   =   "Design Color"
            colname21       =   "WOPPICNM"
            maxlength21     =   45
            style21         =   0
            colname22       =   "WOPSUBPONO"
            heading22       =   "SUB PO                                "
            datatype22      =   4
            maxlength22     =   20
            colname23       =   "WOPCUSTDMCD"
            heading23       =   "Cust Dsg Cd                          "
            datatype23      =   4
            maxlength23     =   40
            tooltiptext23   =   "Customer Design Code (With depth of Suffix, Size)"
            colname24       =   "WCPYODSR"
            heading24       =   "Cpy"
            datatype24      =   4
            maxlength24     =   1
            tooltiptext24   =   "Enter 'Y' if you want to copy design to Order"
            colname25       =   "WOPCOCD"
            maxlength25     =   3
            style25         =   0
         End
         Begin MwfCtl.MWCTL_FLX GrdOrdMst 
            Height          =   2415
            Left            =   180
            TabIndex        =   88
            Top             =   0
            Width           =   10350
            _ExtentX        =   18256
            _ExtentY        =   4260
            Cols            =   12
            AllowAdd        =   0   'False
            AllowDelete     =   0   'False
            colname1        =   "WOMTC"
            maxlength1      =   3
            style1          =   0
            colname2        =   "WOMYY"
            maxlength2      =   2
            style2          =   0
            colname3        =   "WOMCHR"
            maxlength3      =   3
            style3          =   0
            colname4        =   "WOMNO"
            datatype4       =   1
            mask4           =   "#########0"
            maxlength4      =   10
            style4          =   0
            tooltiptext4    =   "Order No."
            colname5        =   "WORDNO"
            heading5        =   "Order No                                         "
            colname6        =   "WOMEPDELDT"
            heading6        =   "ExpDelDt"
            datatype6       =   3
            maxlength6      =   8
            tooltiptext6    =   "Export Delivery Date"
            colname7        =   "WOMCUST"
            heading7        =   "Customer   "
            maxlength7      =   8
            tooltiptext7    =   "Customer Code"
            colname8        =   "WOMPONO"
            heading8        =   "Purchase  Order  Number    "
            maxlength8      =   20
            tooltiptext8    =   "PO Number"
            colname9        =   "WOMPODT"
            heading9        =   "PO Date"
            datatype9       =   3
            maxlength9      =   8
            tooltiptext9    =   "PO Date"
            colname10       =   "WPRT"
            heading10       =   "Priority"
            maxlength10     =   5
            tooltiptext10   =   "Order Priority"
            colname11       =   "WOMCOCD"
            maxlength11     =   3
            style11         =   0
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   30
            Left            =   4650
            TabIndex        =   204
            ToolTipText     =   "Enter Fr Voucher Tc"
            Top             =   8370
            Width           =   6015
            _ExtentX        =   10610
            _ExtentY        =   503
            Enabled         =   0   'False
            MaxLength       =   20
            DataType        =   4
            IdName          =   "WWRKNAME"
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Worker Name :"
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
            Left            =   3150
            TabIndex        =   205
            Top             =   8400
            Width           =   1425
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Line Sketch"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   3
            Left            =   11190
            TabIndex        =   83
            Top             =   2010
            Width           =   885
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "3D Image"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   23
            Left            =   13560
            TabIndex        =   82
            Top             =   2010
            Width           =   795
         End
      End
      Begin VB.Frame FraOrdDsg 
         BorderStyle     =   0  'None
         Height          =   6480
         Left            =   90
         TabIndex        =   91
         Top             =   2250
         Width           =   15270
         Begin TabDlg.SSTab TabOrdDsgDet 
            Height          =   6405
            Left            =   0
            TabIndex        =   92
            Top             =   180
            Width           =   15135
            _ExtentX        =   26696
            _ExtentY        =   11298
            _Version        =   393216
            Tabs            =   4
            Tab             =   3
            TabsPerRow      =   4
            TabHeight       =   520
            TabCaption(0)   =   "<&1> General"
            TabPicture(0)   =   "EmrFrmOrdZoom.frx":0000
            Tab(0).ControlEnabled=   0   'False
            Tab(0).Control(0)=   "FraOrdDsgDet(0)"
            Tab(0).ControlCount=   1
            TabCaption(1)   =   "<&2>Bill Of Materials"
            TabPicture(1)   =   "EmrFrmOrdZoom.frx":001C
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "FraOrdDsgDet(1)"
            Tab(1).ControlCount=   1
            TabCaption(2)   =   "<&3> Lab Details"
            TabPicture(2)   =   "EmrFrmOrdZoom.frx":0038
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "FraOrdDsgDet(2)"
            Tab(2).ControlCount=   1
            TabCaption(3)   =   "<&4> Component Details"
            TabPicture(3)   =   "EmrFrmOrdZoom.frx":0054
            Tab(3).ControlEnabled=   -1  'True
            Tab(3).Control(0)=   "FraOrdDsgDet(3)"
            Tab(3).Control(0).Enabled=   0   'False
            Tab(3).ControlCount=   1
            Begin VB.Frame FraOrdDsgDet 
               Enabled         =   0   'False
               Height          =   3465
               Index           =   3
               Left            =   0
               TabIndex        =   189
               Top             =   420
               Width           =   14925
               Begin MwfCtl.MWCTL_FLX GrdOrdCT 
                  Height          =   2835
                  Left            =   60
                  TabIndex        =   190
                  Top             =   180
                  Width           =   14745
                  _ExtentX        =   26009
                  _ExtentY        =   5001
                  Cols            =   13
                  colname1        =   "OCTSRNO"
                  heading1        =   "Sr "
                  datafld1        =   "OctSrNo"
                  datatype1       =   1
                  mask1           =   "#0"
                  maxlength1      =   2
                  style1          =   2
                  tooltiptext1    =   "Compnent Serial Number"
                  colname2        =   "OCTCD"
                  heading2        =   "Comp Code                     "
                  datafld2        =   "OctCd"
                  datatype2       =   4
                  maxlength2      =   15
                  tooltiptext2    =   "Enter Component Code"
                  colname3        =   "OCTSZ"
                  heading3        =   "Size   "
                  datafld3        =   "OctSz"
                  datatype3       =   4
                  maxlength3      =   5
                  tooltiptext3    =   "Enter Component Size"
                  colname4        =   "OCTCTCHR"
                  heading4        =   "Comp Chr"
                  datafld4        =   "OctCTChr"
                  datatype4       =   4
                  maxlength4      =   3
                  tooltiptext4    =   "Enter Component Code"
                  colname5        =   "OCTPARTS"
                  heading5        =   "Prt Qty"
                  datafld5        =   "OCtParts"
                  datatype5       =   1
                  mask5           =   "##0"
                  maxlength5      =   3
                  tooltiptext5    =   "Enter Component Quantity"
                  colname6        =   "OCTRMCD"
                  heading6        =   "Main RmCd                             "
                  datafld6        =   "OctRmCd"
                  datatype6       =   4
                  recalcparent6   =   "GrdOrdCT(OCTRMWT)"
                  maxlength6      =   16
                  tooltiptext6    =   "Enter Main Raw Material Code"
                  colname7        =   "OCTDMCOL"
                  heading7        =   "Dm Color"
                  datafld7        =   "OctDmCol"
                  datatype7       =   4
                  maxlength7      =   5
                  tooltiptext7    =   "Enter Dm Color"
                  colname8        =   "OCTRMWT"
                  heading8        =   "Rm Weight"
                  datafld8        =   "OctRmWt"
                  datatype8       =   2
                  mask8           =   "###0.000"
                  recalcon8       =   "GrdOrdCT(OCTRMCD)"
                  maxlength8      =   8
                  tooltiptext8    =   "Enter Rm Weight"
                  colname9        =   "OCTDC"
                  heading9        =   "DC  "
                  datafld9        =   "OctDc"
                  datatype9       =   4
                  maxlength9      =   2
                  tooltiptext9    =   "Specify 'C' Only Clr Stn/ 'D' Only Diamond/ 'CD' for Both/ Blank if No Stone is used"
                  colname10       =   "OCTPRDSEQ"
                  heading10       =   "PrdSeq     "
                  datafld10       =   "OctPrdSeq"
                  datatype10      =   4
                  maxlength10     =   5
                  tooltiptext10   =   "Enter Production Sequence"
                  colname11       =   "OCTMAINPRTYN"
                  heading11       =   "MainPrt"
                  datafld11       =   "OctMainPrtYn"
                  datatype11      =   4
                  maxlength11     =   1
                  tooltiptext11   =   "Main Part Yes/No"
                  colname12       =   "OCTDESC"
                  heading12       =   "Description                                                                                           "
                  datafld12       =   "OctDesc"
                  maxlength12     =   45
                  tooltiptext12   =   "Enter Discription for Compnent"
               End
            End
            Begin VB.Frame FraOrdDsgDet 
               Enabled         =   0   'False
               Height          =   3465
               Index           =   2
               Left            =   -75000
               TabIndex        =   131
               Top             =   360
               Width           =   7305
               Begin MwfCtl.MWCTL_FLX GrdOrdLab 
                  Height          =   3210
                  Left            =   90
                  TabIndex        =   132
                  Top             =   180
                  Width           =   7125
                  _ExtentX        =   12568
                  _ExtentY        =   5662
                  Cols            =   11
                  colname1        =   "OLSRNO"
                  heading1        =   "Sr. "
                  datatype1       =   2
                  mask1           =   "#0"
                  maxlength1      =   2
                  style1          =   2
                  tooltiptext1    =   "Enter Labour Sr No."
                  colname2        =   "OLMCD"
                  heading2        =   "Main Cd"
                  datatype2       =   4
                  maxlength2      =   5
                  tooltiptext2    =   "Enter Labour Main Code"
                  colname3        =   "OLSCD"
                  heading3        =   "Sub Cd        "
                  datatype3       =   4
                  maxlength3      =   8
                  tooltiptext3    =   "Enter Labour Sub Code"
                  colname4        =   "OLQW"
                  heading4        =   "By Q/W"
                  datatype4       =   4
                  maxlength4      =   1
                  tooltiptext4    =   "Enter whether Labour Sales Value Is Calculated By Qty Or Wt"
                  colname5        =   "OLQTY"
                  heading5        =   "Qty"
                  datatype5       =   1
                  mask5           =   "##0"
                  maxlength5      =   3
                  tooltiptext5    =   "Enter Labour Quantity"
                  colname6        =   "OLSALRT"
                  heading6        =   "Rate        "
                  datatype6       =   2
                  mask6           =   "#####0.00"
                  maxlength6      =   9
                  tooltiptext6    =   "Enter Labour Sales Rate"
                  colname7        =   "OLSALVAL"
                  heading7        =   "Value        "
                  datatype7       =   2
                  mask7           =   "######0.00"
                  maxlength7      =   10
                  tooltiptext7    =   "Enter Labour Sales Value"
                  colname8        =   "OLCSTQW"
                  datatype8       =   4
                  maxlength8      =   1
                  style8          =   0
                  tooltiptext8    =   "Enter whether Labour Cost Value Is Calculated By Qty Or Wt"
                  colname9        =   "OLCSTRT"
                  datatype9       =   2
                  mask9           =   "#######0.00"
                  maxlength9      =   11
                  style9          =   0
                  tooltiptext9    =   "Enter Labour Cost Rate"
                  colname10       =   "OLCSTVAL"
                  datatype10      =   2
                  mask10          =   "#######0.00"
                  maxlength10     =   11
                  style10         =   0
                  tooltiptext10   =   "Enter Labour Cost Value"
               End
            End
            Begin VB.Frame FraOrdDsgDet 
               Enabled         =   0   'False
               Height          =   5565
               Index           =   1
               Left            =   -75000
               TabIndex        =   129
               Top             =   390
               Width           =   15015
               Begin MwfCtl.MWCTL_FLX GrdOrdRm 
                  Height          =   5250
                  Left            =   120
                  TabIndex        =   130
                  Top             =   210
                  Width           =   14775
                  _ExtentX        =   26061
                  _ExtentY        =   9260
                  Cols            =   31
                  colname1        =   "ORSRNO"
                  heading1        =   "Sr.   "
                  datatype1       =   1
                  mask1           =   "##0"
                  maxlength1      =   3
                  style1          =   2
                  tooltiptext1    =   "Enter Raw Material Sr No."
                  colname2        =   "ORRMCTG"
                  heading2        =   "Ctg"
                  datatype2       =   4
                  maxlength2      =   2
                  tooltiptext2    =   "Enter Raw Material Category"
                  colname3        =   "ORRMSCTG"
                  heading3        =   "SubCtg"
                  datatype3       =   4
                  maxlength3      =   5
                  tooltiptext3    =   "Enter Raw Material Sub Category"
                  colname4        =   "ORRMCD"
                  heading4        =   "Rm Code                                 "
                  datatype4       =   4
                  maxlength4      =   16
                  tooltiptext4    =   "Enter Raw Material Code"
                  colname5        =   "ORLN1"
                  heading5        =   "Sv/Ln  "
                  datatype5       =   2
                  mask5           =   "#0.000"
                  maxlength5      =   6
                  tooltiptext5    =   "Enter RM Sieve Sz / Length"
                  colname6        =   "ORLN2"
                  heading6        =   "Brdth  "
                  datatype6       =   2
                  mask6           =   "#0.000"
                  maxlength6      =   6
                  tooltiptext6    =   "Enter RM Breadth"
                  colname7        =   "ORLN3"
                  heading7        =   "Rm Ht"
                  datatype7       =   2
                  mask7           =   "#0.000"
                  maxlength7      =   6
                  tooltiptext7    =   "Enter Rm Height"
                  colname8        =   "ORQTY"
                  heading8        =   "Qty   "
                  datatype8       =   1
                  mask8           =   "###0"
                  maxlength8      =   4
                  tooltiptext8    =   "Enter Raw Material Quantity"
                  colname9        =   "ORRMPTR"
                  heading9        =   "Rm Ptr     "
                  datatype9       =   2
                  mask9           =   "##0.0000"
                  maxlength9      =   8
                  tooltiptext9    =   "Enter Rm Pointer"
                  colname10       =   "ORWT"
                  heading10       =   "Wt          "
                  datatype10      =   2
                  mask10          =   "-##0.000"
                  maxlength10     =   8
                  tooltiptext10   =   "Enter Raw Material Weigth"
                  colname11       =   "ORLMESAL"
                  heading11       =   "LME Rt   "
                  datatype11      =   2
                  mask11          =   "####0.00"
                  maxlength11     =   8
                  tooltiptext11   =   "Enter LMG/ LMP/LMS Sale Rate"
                  colname12       =   "ORSALRT"
                  heading12       =   "Sale Rt       "
                  datatype12      =   2
                  mask12          =   "#######0.000"
                  maxlength12     =   12
                  tooltiptext12   =   "Enter Raw Material Sales Rate"
                  colname13       =   "WQWSAL"
                  heading13       =   "Q/W"
                  datatype13      =   4
                  maxlength13     =   1
                  style13         =   2
                  colname14       =   "ORSALVAL"
                  heading14       =   "Sale Value"
                  datatype14      =   2
                  mask14          =   "-#######0.000"
                  maxlength14     =   13
                  tooltiptext14   =   "Enter Raw Material Sales Value"
                  colname15       =   "ORPRDQTY"
                  heading15       =   "Prd Qty"
                  datatype15      =   1
                  mask15          =   "###0"
                  maxlength15     =   4
                  tooltiptext15   =   "Enter Prd Rm Qty"
                  colname16       =   "ORPRDWT"
                  heading16       =   "Prd Wt   "
                  datatype16      =   2
                  mask16          =   "-##0.000"
                  maxlength16     =   8
                  tooltiptext16   =   "Enter Prd Rm Weight"
                  colname17       =   "ORCSTRT"
                  datatype17      =   2
                  mask17          =   "#######0.000"
                  maxlength17     =   12
                  style17         =   0
                  tooltiptext17   =   "Enter Cost Rate"
                  colname18       =   "ORCSTVAL"
                  datatype18      =   2
                  mask18          =   "-#######0.00"
                  maxlength18     =   12
                  style18         =   0
                  tooltiptext18   =   "Enter Cost Value"
                  colname19       =   "ORSETSCD"
                  heading19       =   "Setting        "
                  datatype19      =   4
                  maxlength19     =   8
                  tooltiptext19   =   "Enter Setting Code"
                  colname20       =   "ORSETSALRT"
                  heading20       =   "Set Rt   "
                  datatype20      =   2
                  mask20          =   "#####0.00"
                  maxlength20     =   9
                  tooltiptext20   =   "Enter Setting Rate"
                  colname21       =   "ORSETSALVAL"
                  heading21       =   "Set Value"
                  datatype21      =   2
                  mask21          =   "######0.00"
                  maxlength21     =   10
                  tooltiptext21   =   "Enter Setting Value"
                  colname22       =   "ORSETCSTRT"
                  datatype22      =   2
                  mask22          =   "#####0.00"
                  maxlength22     =   9
                  style22         =   0
                  tooltiptext22   =   "Enter Set Cost Rate"
                  colname23       =   "ORSETCSTVAL"
                  datatype23      =   2
                  mask23          =   "#####0.00"
                  maxlength23     =   9
                  style23         =   0
                  tooltiptext23   =   "Enter Set Cost Value"
                  colname24       =   "ORALYCD"
                  heading24       =   "Alloy                                        "
                  datatype24      =   4
                  maxlength24     =   16
                  tooltiptext24   =   "Enter Alloy Code"
                  colname25       =   "ORALYSALRT"
                  heading25       =   "Alloy Rt  "
                  datatype25      =   2
                  mask25          =   "######0.00"
                  maxlength25     =   10
                  tooltiptext25   =   "Enter Alloy Sale Rate"
                  colname26       =   "ORALYCSTRT"
                  datatype26      =   2
                  mask26          =   "######0.00"
                  maxlength26     =   10
                  style26         =   0
                  tooltiptext26   =   "Enter Alloy Cost Rate"
                  colname27       =   "ORWSQTY"
                  heading27       =   "WSet"
                  datatype27      =   1
                  mask27          =   "###0"
                  maxlength27     =   4
                  tooltiptext27   =   "Enter Wax Setting Quantity"
                  colname28       =   "ORHSQTY"
                  heading28       =   "HSet"
                  datatype28      =   1
                  mask28          =   "###0"
                  maxlength28     =   4
                  style28         =   2
                  tooltiptext28   =   "Enter Hand Setting Quantity"
                  colname29       =   "ORSUBSHP"
                  heading29       =   "SShp                 "
                  datatype29      =   4
                  maxlength29     =   8
                  tooltiptext29   =   "Enter Sub Shape"
                  colname30       =   "ORMAINMET"
                  heading30       =   "MMet"
                  datatype30      =   4
                  maxlength30     =   1
                  tooltiptext30   =   "Enter Main Metal (Y/N)"
               End
            End
            Begin VB.Frame FraOrdDsgDet 
               Enabled         =   0   'False
               Height          =   6045
               Index           =   0
               Left            =   -74910
               TabIndex        =   133
               Top             =   330
               Width           =   14955
               Begin VB.Frame Fra2 
                  Height          =   1665
                  Left            =   2670
                  TabIndex        =   144
                  Top             =   360
                  Width           =   6945
                  Begin MwfCtl.MWCTL_MED atxtDisp 
                     Height          =   285
                     Index           =   34
                     Left            =   885
                     TabIndex        =   145
                     ToolTipText     =   "Enter Order Quantity"
                     Top             =   150
                     Width           =   825
                     _ExtentX        =   1455
                     _ExtentY        =   503
                     Alignment       =   2
                     Enabled         =   0   'False
                     Mask            =   "-####0.0"
                     MaxLength       =   8
                     DataType        =   2
                     IdName          =   "WODORDQTY"
                  End
                  Begin MwfCtl.MWCTL_MED atxtDisp 
                     Height          =   285
                     Index           =   35
                     Left            =   885
                     TabIndex        =   146
                     ToolTipText     =   "Enter Export Quantity"
                     Top             =   435
                     Width           =   825
                     _ExtentX        =   1455
                     _ExtentY        =   503
                     Alignment       =   2
                     Enabled         =   0   'False
                     Mask            =   "####0.0"
                     MaxLength       =   7
                     DataType        =   2
                     IdName          =   "WODEXPQTY"
                  End
                  Begin MwfCtl.MWCTL_MED atxtDisp 
                     Height          =   285
                     Index           =   36
                     Left            =   885
                     TabIndex        =   147
                     ToolTipText     =   "Enter Production Quantity"
                     Top             =   1005
                     Width           =   825
                     _ExtentX        =   1455
                     _ExtentY        =   503
                     Alignment       =   2
                     Enabled         =   0   'False
                     Mask            =   "-####0.0"
                     MaxLength       =   8
                     DataType        =   2
                     ReCalcOn        =   "WODPRDENT"
                     IdName          =   "WODPRDQTY"
                  End
                  Begin MwfCtl.MWCTL_MED atxtDisp 
                     Height          =   285
                     Index           =   40
                     Left            =   885
                     TabIndex        =   97
                     ToolTipText     =   "Entered Production Quantity "
                     Top             =   720
                     Width           =   825
                     _ExtentX        =   1455
                     _ExtentY        =   503
                     Alignment       =   2
                     Mask            =   "####0.0"
                     MaxLength       =   7
                     DataType        =   2
                     ReCalcParent    =   "WODPRDQTY"
                     IdName          =   "WODPRDENT"
                  End
                  Begin MwfCtl.MWCTL_MED atxtDisp 
                     Height          =   285
                     Index           =   41
                     Left            =   3015
                     TabIndex        =   148
                     ToolTipText     =   "Enter Finished Goods Quantity"
                     Top             =   150
                     Width           =   825
                     _ExtentX        =   1455
                     _ExtentY        =   503
                     Alignment       =   2
                     Enabled         =   0   'False
                     Mask            =   "####0.0"
                     MaxLength       =   7
                     DataType        =   2
                     IdName          =   "WODFGQTY"
                  End
                  Begin MwfCtl.MWCTL_MED atxtDisp 
                     Height          =   285
                     Index           =   45
                     Left            =   3015
                     TabIndex        =   149
                     ToolTipText     =   "Enter Cost Price"
                     Top             =   720
                     Width           =   1005
                     _ExtentX        =   1773
                     _ExtentY        =   503
                     Alignment       =   2
                     Enabled         =   0   'False
                     Mask            =   "#####0.00"
                     MaxLength       =   9
                     DataType        =   2
                     IdName          =   "WODCSTPRC"
                  End
                  Begin MwfCtl.MWCTL_MED atxtDisp 
                     Height          =   285
                     Index           =   53
                     Left            =   3015
                     TabIndex        =   98
                     ToolTipText     =   "Enter Bag Opening Pieces"
                     Top             =   435
                     Width           =   465
                     _ExtentX        =   820
                     _ExtentY        =   503
                     Alignment       =   2
                     Mask            =   "##0"
                     MaxLength       =   3
                     DataType        =   1
                     IdName          =   "WODBAGPCS"
                  End
                  Begin MwfCtl.MWCTL_MED atxtDisp 
                     Height          =   285
                     Index           =   75
                     Left            =   3015
                     TabIndex        =   150
                     ToolTipText     =   "Design Category"
                     Top             =   1005
                     Width           =   885
                     _ExtentX        =   1561
                     _ExtentY        =   503
                     Enabled         =   0   'False
                     MaxLength       =   5
                     DataType        =   4
                     IdName          =   "WODDMCTG"
                  End
                  Begin MwfCtl.MWCTL_MED atxtDisp 
                     Height          =   285
                     Index           =   138
                     Left            =   5505
                     TabIndex        =   100
                     ToolTipText     =   "Enter Minimium Acceptable Gross Gram Wt (eg: Gold,Platinum,Accesserioes)"
                     Top             =   720
                     Width           =   915
                     _ExtentX        =   1614
                     _ExtentY        =   503
                     Alignment       =   2
                     Mask            =   "###0.000"
                     MaxLength       =   8
                     DataType        =   2
                     ReCalcParent    =   "WODPRDQTY"
                     IdName          =   "WODMINWT"
                  End
                  Begin MwfCtl.MWCTL_MED atxtDisp 
                     Height          =   285
                     Index           =   139
                     Left            =   5505
                     TabIndex        =   101
                     ToolTipText     =   "Enter Maximum Acceptable Gross Gram Wt (eg: for Gold,Platinum,Accesserioes)"
                     Top             =   1005
                     Width           =   915
                     _ExtentX        =   1614
                     _ExtentY        =   503
                     Alignment       =   2
                     Mask            =   "###0.000"
                     MaxLength       =   8
                     DataType        =   2
                     ReCalcParent    =   "WODPRDQTY"
                     IdName          =   "WODMAXWT"
                  End
                  Begin MwfCtl.MWCTL_MED atxtDisp 
                     Height          =   285
                     Index           =   141
                     Left            =   5505
                     TabIndex        =   151
                     ToolTipText     =   "Gross Gram Wt  (eg:Wt of Gold,Platinum,Accesserioes)"
                     Top             =   150
                     Width           =   915
                     _ExtentX        =   1614
                     _ExtentY        =   503
                     Alignment       =   2
                     Enabled         =   0   'False
                     Mask            =   "####0.000"
                     MaxLength       =   9
                     DataType        =   2
                     IdName          =   "WODGRSWT"
                  End
                  Begin MwfCtl.MWCTL_MED atxtDisp 
                     Height          =   285
                     Index           =   152
                     Left            =   5505
                     TabIndex        =   99
                     ToolTipText     =   "Enter Gram Check Code"
                     Top             =   435
                     Width           =   1365
                     _ExtentX        =   2408
                     _ExtentY        =   503
                     MaxLength       =   8
                     DataType        =   4
                     IdName          =   "WODGMCHK"
                  End
                  Begin MwfCtl.MWCTL_MED atxtDisp 
                     Height          =   285
                     Index           =   268
                     Left            =   5505
                     TabIndex        =   103
                     ToolTipText     =   "Enter Maximum Diamond Tolerance Wt"
                     Top             =   1290
                     Width           =   915
                     _ExtentX        =   1614
                     _ExtentY        =   503
                     Alignment       =   2
                     Mask            =   "###0.000"
                     MaxLength       =   8
                     DataType        =   2
                     IdName          =   "WODMAXDIATOLWT"
                  End
                  Begin MwfCtl.MWCTL_MED atxtDisp 
                     Height          =   285
                     Index           =   269
                     Left            =   3015
                     TabIndex        =   102
                     ToolTipText     =   "Enter Minimum Diamond Tolerance Wt"
                     Top             =   1290
                     Width           =   915
                     _ExtentX        =   1614
                     _ExtentY        =   503
                     Alignment       =   2
                     Mask            =   "###0.000"
                     MaxLength       =   8
                     DataType        =   2
                     ReCalcOn        =   "WODIWTEQORD,WODIWTFRORD"
                     IdName          =   "WODMINDIATOLWT"
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Max Dia Tol"
                     BeginProperty Font 
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
                     Left            =   4170
                     TabIndex        =   202
                     Top             =   1290
                     Width           =   1275
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Min Dia Tol"
                     BeginProperty Font 
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
                     Left            =   1860
                     TabIndex        =   201
                     Top             =   1290
                     Width           =   1095
                  End
                  Begin VB.Label Label1 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Gram Check"
                     BeginProperty Font 
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
                     Left            =   4140
                     TabIndex        =   163
                     Top             =   435
                     Width           =   1395
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Gram Wt"
                     BeginProperty Font 
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
                     Left            =   4140
                     TabIndex        =   162
                     Top             =   150
                     Width           =   915
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Gram Max Wt"
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   285
                     Index           =   123
                     Left            =   4140
                     TabIndex        =   161
                     Top             =   1005
                     Width           =   1335
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Gram Min Wt"
                     BeginProperty Font 
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
                     Left            =   4140
                     TabIndex        =   160
                     Top             =   720
                     Width           =   1335
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Dsg Ctg"
                     BeginProperty Font 
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
                     Left            =   1890
                     TabIndex        =   159
                     Top             =   1005
                     Width           =   945
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Bag Pcs"
                     BeginProperty Font 
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
                     Left            =   1890
                     TabIndex        =   158
                     Top             =   435
                     Width           =   945
                  End
                  Begin VB.Label LblCstPrc 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Cost Prc"
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   285
                     Left            =   1890
                     TabIndex        =   157
                     Top             =   720
                     Width           =   945
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Fg Qty"
                     BeginProperty Font 
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
                     Left            =   1890
                     TabIndex        =   156
                     Top             =   150
                     Width           =   945
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Prd Ent"
                     BeginProperty Font 
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
                     Left            =   60
                     TabIndex        =   155
                     Top             =   720
                     Width           =   795
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Prd Qty"
                     BeginProperty Font 
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
                     Left            =   60
                     TabIndex        =   154
                     Top             =   1005
                     Width           =   825
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Exp Qty"
                     BeginProperty Font 
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
                     Left            =   60
                     TabIndex        =   153
                     Top             =   435
                     Width           =   855
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Ord Qty"
                     BeginProperty Font 
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
                     Left            =   60
                     TabIndex        =   152
                     Top             =   150
                     Width           =   825
                  End
               End
               Begin VB.Frame Fr1 
                  Height          =   1335
                  Left            =   60
                  TabIndex        =   169
                  Top             =   360
                  Width           =   2565
                  Begin MwfCtl.MWCTL_MED atxtDisp 
                     Height          =   285
                     Index           =   26
                     Left            =   1485
                     TabIndex        =   93
                     ToolTipText     =   "Enter Production Sequence"
                     Top             =   150
                     Width           =   885
                     _ExtentX        =   1561
                     _ExtentY        =   503
                     MaxLength       =   5
                     DataType        =   4
                     IdName          =   "WODPRDSEQ"
                  End
                  Begin MwfCtl.MWCTL_MED atxtDisp 
                     Height          =   285
                     Index           =   52
                     Left            =   1485
                     TabIndex        =   94
                     ToolTipText     =   "Enter Value Addition Category"
                     Top             =   435
                     Width           =   885
                     _ExtentX        =   1561
                     _ExtentY        =   503
                     MaxLength       =   5
                     DataType        =   4
                     IdName          =   "WODVACTG"
                  End
                  Begin MwfCtl.MWCTL_MED atxtDisp 
                     Height          =   285
                     Index           =   73
                     Left            =   1485
                     TabIndex        =   95
                     ToolTipText     =   "Enter Design Colour"
                     Top             =   720
                     Width           =   885
                     _ExtentX        =   1561
                     _ExtentY        =   503
                     MaxLength       =   5
                     DataType        =   4
                     IdName          =   "WODDMCOL"
                  End
                  Begin MwfCtl.MWCTL_MED atxtDisp 
                     Height          =   285
                     Index           =   153
                     Left            =   1485
                     TabIndex        =   96
                     ToolTipText     =   "Specify Whether Tolerence Is To Be Checked"
                     Top             =   1005
                     Visible         =   0   'False
                     Width           =   345
                     _ExtentX        =   609
                     _ExtentY        =   503
                     DataType        =   4
                     IdName          =   "WODCHKTOL"
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Chk Tolerance"
                     BeginProperty Font 
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
                     Left            =   60
                     TabIndex        =   173
                     Top             =   1005
                     Visible         =   0   'False
                     Width           =   1425
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Colour"
                     BeginProperty Font 
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
                     Left            =   60
                     TabIndex        =   172
                     Top             =   720
                     Width           =   915
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Val Addn Ctg"
                     BeginProperty Font 
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
                     Left            =   60
                     TabIndex        =   171
                     Top             =   435
                     Width           =   1395
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Prd Seq"
                     BeginProperty Font 
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
                     Left            =   60
                     TabIndex        =   170
                     Top             =   150
                     Width           =   915
                  End
               End
               Begin VB.Frame FraDsgFact 
                  Height          =   1665
                  Left            =   9630
                  TabIndex        =   138
                  Top             =   360
                  Width           =   2535
                  Begin MwfCtl.MWCTL_MED atxtDisp 
                     Height          =   285
                     Index           =   27
                     Left            =   1455
                     TabIndex        =   104
                     ToolTipText     =   "Enter Multiplying Factor"
                     Top             =   150
                     Width           =   675
                     _ExtentX        =   1191
                     _ExtentY        =   503
                     Alignment       =   2
                     Mask            =   "#0.000"
                     MaxLength       =   6
                     DataType        =   2
                     IdName          =   "WODMULBY"
                  End
                  Begin MwfCtl.MWCTL_MED atxtDisp 
                     Height          =   285
                     Index           =   28
                     Left            =   1455
                     TabIndex        =   105
                     ToolTipText     =   "Enter Fixed Price (Y / N)"
                     Top             =   435
                     Width           =   315
                     _ExtentX        =   556
                     _ExtentY        =   503
                     DataType        =   4
                     IdName          =   "WODFIXPRC"
                  End
                  Begin MwfCtl.MWCTL_MED atxtDisp 
                     Height          =   285
                     Index           =   29
                     Left            =   1455
                     TabIndex        =   106
                     ToolTipText     =   "Enter Gold As"
                     Top             =   720
                     Width           =   855
                     _ExtentX        =   1508
                     _ExtentY        =   503
                     MaxLength       =   5
                     DataType        =   4
                     IdName          =   "WODGLDAS"
                  End
                  Begin MwfCtl.MWCTL_MED atxtDisp 
                     Height          =   285
                     Index           =   156
                     Left            =   1455
                     TabIndex        =   108
                     ToolTipText     =   "Enter Crm Fixed Price (Y / N)"
                     Top             =   1305
                     Visible         =   0   'False
                     Width           =   315
                     _ExtentX        =   556
                     _ExtentY        =   503
                     DataType        =   4
                     IdName          =   "WODCRMFIXPRCYN"
                  End
                  Begin MwfCtl.MWCTL_MED atxtDisp 
                     Height          =   285
                     Index           =   206
                     Left            =   1455
                     TabIndex        =   107
                     ToolTipText     =   "Enter Lab As"
                     Top             =   1020
                     Width           =   855
                     _ExtentX        =   1508
                     _ExtentY        =   503
                     MaxLength       =   5
                     DataType        =   4
                     IdName          =   "WODLABAS"
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
                     Index           =   169
                     Left            =   30
                     TabIndex        =   143
                     Top             =   1020
                     Width           =   1215
                  End
                  Begin VB.Label LblCrmFixPrcYN 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Crm Fixed Prc"
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   285
                     Left            =   30
                     TabIndex        =   142
                     Top             =   1305
                     Visible         =   0   'False
                     Width           =   1365
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Multiply By"
                     BeginProperty Font 
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
                     Left            =   30
                     TabIndex        =   141
                     Top             =   150
                     Width           =   1215
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Fixed Prc"
                     BeginProperty Font 
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
                     Left            =   30
                     TabIndex        =   140
                     Top             =   435
                     Width           =   1215
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
                     Index           =   13
                     Left            =   30
                     TabIndex        =   139
                     Top             =   720
                     Width           =   1215
                  End
               End
               Begin MwfCtl.MWCTL_MED atxtDisp 
                  Height          =   570
                  Index           =   46
                  Left            =   1365
                  TabIndex        =   112
                  ToolTipText     =   "Enter Design Production Instruction"
                  Top             =   2625
                  Width           =   9700
                  _ExtentX        =   17119
                  _ExtentY        =   1005
                  Alignment       =   3
                  MaxLength       =   250
                  IdName          =   "WODDMPRDINST"
               End
               Begin MwfCtl.MWCTL_MED atxtDisp 
                  Height          =   570
                  Index           =   47
                  Left            =   1365
                  TabIndex        =   113
                  ToolTipText     =   "Enter Customer Production Instruction"
                  Top             =   3180
                  Width           =   9700
                  _ExtentX        =   17119
                  _ExtentY        =   1005
                  Alignment       =   3
                  MaxLength       =   250
                  IdName          =   "WODCMPRDINST"
               End
               Begin MwfCtl.MWCTL_MED atxtDisp 
                  Height          =   285
                  Index           =   30
                  Left            =   1365
                  TabIndex        =   114
                  ToolTipText     =   "Enter Stamping Instruction"
                  Top             =   3735
                  Width           =   9700
                  _ExtentX        =   17119
                  _ExtentY        =   503
                  MaxLength       =   90
                  IdName          =   "WODCMSTMPINST"
               End
               Begin MwfCtl.MWCTL_MED atxtDisp 
                  Height          =   285
                  Index           =   49
                  Left            =   1365
                  TabIndex        =   115
                  ToolTipText     =   "Enter Size Instructions"
                  Top             =   4020
                  Width           =   9700
                  _ExtentX        =   17119
                  _ExtentY        =   503
                  MaxLength       =   90
                  IdName          =   "WODSZINST"
               End
               Begin MwfCtl.MWCTL_MED atxtDisp 
                  Height          =   285
                  Index           =   50
                  Left            =   1365
                  TabIndex        =   109
                  ToolTipText     =   "Enter Design On Hold (Y/N)"
                  Top             =   2340
                  Width           =   315
                  _ExtentX        =   556
                  _ExtentY        =   503
                  DataType        =   4
                  ReCalcParent    =   "WODHLDDESC"
                  IdName          =   "WODHLD"
               End
               Begin MwfCtl.MWCTL_MED atxtDisp 
                  Height          =   285
                  Index           =   51
                  Left            =   2925
                  TabIndex        =   110
                  ToolTipText     =   "Enter Holding Description"
                  Top             =   2340
                  Width           =   7980
                  _ExtentX        =   14076
                  _ExtentY        =   503
                  MaxLength       =   45
                  ReCalcOn        =   "WODHLD"
                  IdName          =   "WODHLDDESC"
               End
               Begin MwfCtl.MWCTL_MED atxtDisp 
                  Height          =   570
                  Index           =   88
                  Left            =   1365
                  TabIndex        =   116
                  ToolTipText     =   "Enter Sales Person's Remark"
                  Top             =   4305
                  Width           =   9700
                  _ExtentX        =   17119
                  _ExtentY        =   1005
                  Alignment       =   3
                  MaxLength       =   100
                  IdName          =   "WODSALREM"
               End
               Begin MwfCtl.MWCTL_MED atxtDisp 
                  Height          =   285
                  Index           =   122
                  Left            =   10725
                  TabIndex        =   111
                  ToolTipText     =   "Enter Contractor Code"
                  Top             =   2040
                  Width           =   1365
                  _ExtentX        =   2408
                  _ExtentY        =   503
                  MaxLength       =   8
                  DataType        =   4
                  IdName          =   "WODOMCTCD"
               End
               Begin MwfCtl.MWCTL_MED atxtDisp 
                  Height          =   570
                  Index           =   204
                  Left            =   1365
                  TabIndex        =   117
                  ToolTipText     =   "Enter Additional Remarks"
                  Top             =   4860
                  Width           =   9700
                  _ExtentX        =   17119
                  _ExtentY        =   1005
                  MaxLength       =   120
                  IdName          =   "WODSUBREM"
               End
               Begin MwfCtl.MWCTL_MED atxtDisp 
                  Height          =   285
                  Index           =   234
                  Left            =   1365
                  TabIndex        =   193
                  ToolTipText     =   "Enter Picture name w/o Prefix ""DM 3D"" or ""DM LD"""
                  Top             =   2040
                  Width           =   6885
                  _ExtentX        =   12144
                  _ExtentY        =   503
                  MaxLength       =   45
                  IdName          =   "WODPICNM"
               End
               Begin VB.Frame FraDsgOrdAct 
                  Height          =   1935
                  Left            =   11085
                  TabIndex        =   164
                  Top             =   2580
                  Width           =   3840
                  Begin MwfCtl.MWCTL_MED atxtDisp 
                     Height          =   285
                     Index           =   31
                     Left            =   2265
                     TabIndex        =   118
                     ToolTipText     =   "Invoice Weight Equal To Order (Y / N)"
                     Top             =   390
                     Width           =   315
                     _ExtentX        =   556
                     _ExtentY        =   503
                     DataType        =   4
                     ReCalcParent    =   "WODIWTFRORD"
                     IdName          =   "WODIWTEQORD"
                  End
                  Begin MwfCtl.MWCTL_MED atxtDisp 
                     Height          =   285
                     Index           =   32
                     Left            =   2265
                     TabIndex        =   119
                     ToolTipText     =   "Enter Invoice Weights From Order"
                     Top             =   675
                     Width           =   885
                     _ExtentX        =   1561
                     _ExtentY        =   503
                     MaxLength       =   5
                     DataType        =   4
                     ReCalcOn        =   "WODIWTEQORD"
                     IdName          =   "WODIWTFRORD"
                  End
                  Begin MwfCtl.MWCTL_MED atxtDisp 
                     Height          =   285
                     Index           =   33
                     Left            =   2265
                     TabIndex        =   120
                     ToolTipText     =   "Invoice Gold Rate From Order (Y / N)"
                     Top             =   960
                     Width           =   315
                     _ExtentX        =   556
                     _ExtentY        =   503
                     DataType        =   4
                     IdName          =   "WODIGLDRTEQORD"
                  End
                  Begin MwfCtl.MWCTL_MED atxtDisp 
                     Height          =   285
                     Index           =   237
                     Left            =   2265
                     TabIndex        =   121
                     ToolTipText     =   "Change Pointer On Size To Weight Group (Y / N)"
                     Top             =   1260
                     Width           =   315
                     _ExtentX        =   556
                     _ExtentY        =   503
                     DataType        =   4
                     IdName          =   "WODCHGPTRONSTWYN"
                  End
                  Begin MwfCtl.MWCTL_MED atxtDisp 
                     Height          =   285
                     Index           =   239
                     Left            =   2265
                     TabIndex        =   122
                     ToolTipText     =   "Labor Weight From Order (Y / N)"
                     Top             =   1545
                     Width           =   315
                     _ExtentX        =   556
                     _ExtentY        =   503
                     DataType        =   4
                     IdName          =   "WODILABWTFRORD"
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Chg Ptr On STW"
                     BeginProperty Font 
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
                     Left            =   120
                     TabIndex        =   196
                     Top             =   1260
                     Width           =   1815
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Labor Wt. From Ord"
                     BeginProperty Font 
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
                     Left            =   120
                     TabIndex        =   195
                     Top             =   1545
                     Width           =   2085
                  End
                  Begin VB.Label LblMain 
                     BackStyle       =   0  'Transparent
                     BorderStyle     =   1  'Fixed Single
                     Caption         =   "  Invoice Wt && Rate From Order / Actuals"
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
                     TabIndex        =   168
                     Top             =   90
                     Width           =   3870
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
                     Index           =   21
                     Left            =   120
                     TabIndex        =   167
                     Top             =   420
                     Width           =   1815
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Wts From Ord "
                     BeginProperty Font 
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
                     TabIndex        =   166
                     Top             =   690
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
                     Index           =   11
                     Left            =   120
                     TabIndex        =   165
                     Top             =   990
                     Width           =   1725
                  End
               End
               Begin VB.Frame FraPpc 
                  Height          =   1350
                  Left            =   11085
                  TabIndex        =   174
                  Top             =   4455
                  Width           =   3840
                  Begin MwfCtl.MWCTL_MED atxtDisp 
                     Height          =   285
                     Index           =   94
                     Left            =   2265
                     TabIndex        =   125
                     ToolTipText     =   "Enter Diamond Availability Date For Wax Setting "
                     Top             =   720
                     Width           =   855
                     _ExtentX        =   1508
                     _ExtentY        =   503
                     MaxLength       =   8
                     DataType        =   3
                     IdName          =   "WODWDIAAVLBLDT"
                  End
                  Begin MwfCtl.MWCTL_MED atxtDisp 
                     Height          =   285
                     Index           =   95
                     Left            =   2265
                     TabIndex        =   126
                     ToolTipText     =   "Enter Enter Diamond Availability Date For Hand Setting"
                     Top             =   1005
                     Width           =   855
                     _ExtentX        =   1508
                     _ExtentY        =   503
                     MaxLength       =   8
                     DataType        =   3
                     IdName          =   "WODHDIAAVLBLDT"
                  End
                  Begin MwfCtl.MWCTL_MED atxtDisp 
                     Height          =   285
                     Index           =   96
                     Left            =   2265
                     TabIndex        =   123
                     ToolTipText     =   "Invoice Weight Equal To Order (Y / N)"
                     Top             =   150
                     Width           =   495
                     _ExtentX        =   873
                     _ExtentY        =   503
                     MaxLength       =   2
                     DataType        =   4
                     IdName          =   "WODWH"
                  End
                  Begin MwfCtl.MWCTL_MED atxtDisp 
                     Height          =   285
                     Index           =   93
                     Left            =   2265
                     TabIndex        =   124
                     ToolTipText     =   "Enter Production Line"
                     Top             =   435
                     Width           =   855
                     _ExtentX        =   1508
                     _ExtentY        =   503
                     MaxLength       =   5
                     DataType        =   4
                     IdName          =   "WODLINE"
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Production Line"
                     BeginProperty Font 
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
                     Left            =   90
                     TabIndex        =   178
                     Top             =   435
                     Width           =   1665
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "W/H"
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
                     Index           =   86
                     Left            =   60
                     TabIndex        =   177
                     Top             =   150
                     Width           =   825
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "WaxSet Dia Avlbl Dt"
                     BeginProperty Font 
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
                     TabIndex        =   176
                     Top             =   720
                     Width           =   2205
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "HndSet Dia Avlbl Dt"
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   195
                     Index           =   84
                     Left            =   60
                     TabIndex        =   175
                     Top             =   1005
                     Width           =   2205
                  End
               End
               Begin VB.Frame FraGrdDet 
                  Height          =   540
                  Left            =   60
                  TabIndex        =   134
                  Top             =   5475
                  Visible         =   0   'False
                  Width           =   8220
                  Begin MwfCtl.MWCTL_MED atxtDisp 
                     Height          =   285
                     Index           =   155
                     Left            =   2055
                     TabIndex        =   127
                     ToolTipText     =   "Enter Metal Grade"
                     Top             =   150
                     Width           =   1575
                     _ExtentX        =   2778
                     _ExtentY        =   503
                     MaxLength       =   8
                     DataType        =   4
                     IdName          =   "WODGRMET"
                  End
                  Begin MwfCtl.MWCTL_MED atxtDisp 
                     Height          =   285
                     Index           =   154
                     Left            =   4800
                     TabIndex        =   128
                     ToolTipText     =   "Enter Diamond Grade"
                     Top             =   150
                     Width           =   1575
                     _ExtentX        =   2778
                     _ExtentY        =   503
                     MaxLength       =   8
                     DataType        =   4
                     IdName          =   "WODGRDIA"
                  End
                  Begin VB.Label LblMain 
                     BackStyle       =   0  'Transparent
                     BorderStyle     =   1  'Fixed Single
                     Caption         =   "Grade Details"
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   9
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   450
                     Index           =   15
                     Left            =   0
                     TabIndex        =   137
                     Top             =   90
                     Width           =   1350
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Diamond"
                     BeginProperty Font 
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
                     Left            =   3825
                     TabIndex        =   136
                     Top             =   150
                     Width           =   1035
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Metal"
                     BeginProperty Font 
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
                     Left            =   1440
                     TabIndex        =   135
                     Top             =   150
                     Width           =   795
                  End
               End
               Begin MwfCtl.MWCTL_MED atxtDisp 
                  Height          =   285
                  Index           =   262
                  Left            =   1365
                  TabIndex        =   199
                  ToolTipText     =   "Inserted Date"
                  Top             =   1755
                  Width           =   885
                  _ExtentX        =   1561
                  _ExtentY        =   503
                  Enabled         =   0   'False
                  MaxLength       =   8
                  DataType        =   3
                  IdName          =   "WODINSDT"
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Ins Date"
                  BeginProperty Font 
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
                  Left            =   60
                  TabIndex        =   200
                  Top             =   1755
                  Width           =   1185
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Picture Name"
                  BeginProperty Font 
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
                  Left            =   60
                  TabIndex        =   192
                  Top             =   2040
                  Width           =   1335
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Sub Rem*"
                  BeginProperty Font 
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
                  Left            =   60
                  TabIndex        =   188
                  Top             =   4845
                  Width           =   1185
               End
               Begin VB.Label LblwOdOmCtCd 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Contractor"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Left            =   9630
                  TabIndex        =   187
                  Top             =   2040
                  Width           =   1095
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Hold Desc"
                  BeginProperty Font 
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
                  Left            =   1800
                  TabIndex        =   186
                  Top             =   2340
                  Width           =   1155
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "On Hold"
                  BeginProperty Font 
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
                  Left            =   60
                  TabIndex        =   185
                  Top             =   2340
                  Width           =   915
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Size Inst"
                  BeginProperty Font 
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
                  Left            =   60
                  TabIndex        =   184
                  Top             =   4020
                  Width           =   1455
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Stmp Inst"
                  BeginProperty Font 
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
                  Left            =   60
                  TabIndex        =   183
                  Top             =   3735
                  Width           =   1455
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Cust Prd Inst"
                  BeginProperty Font 
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
                  Left            =   60
                  TabIndex        =   182
                  Top             =   3180
                  Width           =   1455
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Dsg Prd Inst"
                  BeginProperty Font 
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
                  Left            =   60
                  TabIndex        =   181
                  Top             =   2625
                  Width           =   1455
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Special Remark"
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
                  Index           =   78
                  Left            =   60
                  TabIndex        =   180
                  Top             =   4305
                  Width           =   1245
               End
               Begin VB.Label LblMain 
                  BackStyle       =   0  'Transparent
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "  Design Details"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   300
                  Index           =   10
                  Left            =   30
                  TabIndex        =   179
                  Top             =   90
                  Width           =   14850
               End
            End
         End
      End
      Begin VB.Frame FraInvDet 
         Enabled         =   0   'False
         Height          =   8235
         Left            =   150
         TabIndex        =   67
         Top             =   510
         Width           =   14985
         Begin MwfCtl.MWCTL_FLX GrdInv 
            Height          =   5550
            Left            =   60
            TabIndex        =   68
            Top             =   450
            Width           =   8310
            _ExtentX        =   14658
            _ExtentY        =   9790
            Cols            =   11
            AllowAdd        =   0   'False
            AllowDelete     =   0   'False
            colname1        =   "WIDTC"
            heading1        =   "InTc"
            maxlength1      =   3
            tooltiptext1    =   "Invoice Tc"
            colname2        =   "WIDYY"
            heading2        =   "Yy"
            maxlength2      =   2
            tooltiptext2    =   "Invoice Yy"
            colname3        =   "WIDCHR"
            heading3        =   "Chr  "
            maxlength3      =   3
            tooltiptext3    =   "Invoice Chr"
            colname4        =   "WIDNO"
            heading4        =   "No        "
            datatype4       =   1
            mask4           =   "#####0"
            maxlength4      =   6
            tooltiptext4    =   "Invoice No."
            colname5        =   "WIDSR"
            heading5        =   "Sr     "
            datatype5       =   1
            mask5           =   "###0"
            maxlength5      =   4
            tooltiptext5    =   "Invoice Sr"
            colname6        =   "WINDT"
            heading6        =   "InvDate "
            datatype6       =   3
            maxlength6      =   8
            tooltiptext6    =   "Invoice Date"
            colname7        =   "WINEXPNO"
            heading7        =   "InExpNo                                     "
            maxlength7      =   20
            tooltiptext7    =   "Invoice Export No"
            colname8        =   "WIDQTY"
            heading8        =   "Qty         "
            datatype8       =   2
            mask8           =   "####0.0"
            maxlength8      =   7
            tooltiptext8    =   "Tot Qty in Exp Vch"
            colname9        =   "WIDIFOBPC"
            heading9        =   "Prc             "
            datatype9       =   2
            mask9           =   "#####0.#0"
            maxlength9      =   9
            tooltiptext9    =   "Per Pc Price"
            colname10       =   "WINCOCD"
            maxlength10     =   3
            style10         =   0
         End
         Begin MwfCtl.MWCTL_FLX GrdInvBag 
            Height          =   5550
            Left            =   8370
            TabIndex        =   69
            Top             =   450
            Width           =   4380
            _ExtentX        =   7726
            _ExtentY        =   9790
            Cols            =   6
            AllowAdd        =   0   'False
            AllowDelete     =   0   'False
            colname1        =   "WIBYY"
            heading1        =   "BYy"
            datatype1       =   4
            maxlength1      =   2
            tooltiptext1    =   "Inv Bag YY"
            colname2        =   "WIBCHR"
            heading2        =   "BChr"
            datatype2       =   4
            maxlength2      =   3
            tooltiptext2    =   "Bag Chr"
            colname3        =   "WIBNO"
            heading3        =   "Bag No  "
            datatype3       =   1
            mask3           =   "######0"
            maxlength3      =   7
            tooltiptext3    =   "Inv Bag No"
            colname4        =   "WIBPCS"
            heading4        =   "Pcs         "
            datatype4       =   2
            mask4           =   "####0.0"
            maxlength4      =   7
            tooltiptext4    =   "Number of Pieces of the Design In the Bag"
            colname5        =   "WIBGRSWT"
            heading5        =   "Gross Wt   "
            datatype5       =   2
            mask5           =   "####0.##0"
            maxlength5      =   9
            tooltiptext5    =   "Inv Grs Wt Of the Pieces in the Bag"
         End
         Begin MwfCtl.MWCTL_FLX GrdInvRm 
            Height          =   2700
            Left            =   60
            TabIndex        =   70
            Top             =   6000
            Width           =   11250
            _ExtentX        =   19844
            _ExtentY        =   4763
            Cols            =   13
            AllowAdd        =   0   'False
            AllowDelete     =   0   'False
            colname1        =   "WRMCTG"
            heading1        =   "RmCtg"
            datatype1       =   4
            maxlength1      =   2
            tooltiptext1    =   "Rm Ctg"
            colname2        =   "WRMSCTG"
            heading2        =   "RmSCtg"
            datatype2       =   4
            maxlength2      =   5
            tooltiptext2    =   "Rm Sub Ctg"
            colname3        =   "WRMCD"
            heading3        =   "RmCd                                      "
            datatype3       =   4
            maxlength3      =   16
            tooltiptext3    =   "RM Code"
            colname4        =   "WRMQTY"
            heading4        =   "RmQty"
            datatype4       =   1
            mask4           =   "###0"
            maxlength4      =   4
            tooltiptext4    =   "Rm Qty"
            colname5        =   "WRMIWT"
            heading5        =   "RmIWt    "
            datatype5       =   2
            mask5           =   "##0.##0"
            maxlength5      =   7
            tooltiptext5    =   "Rm Inv Wt"
            colname6        =   "WQW"
            heading6        =   "Q/W"
            datatype6       =   4
            maxlength6      =   1
            colname7        =   "WRMIRT"
            heading7        =   "RmIRt         "
            datatype7       =   2
            mask7           =   "#####0.#0"
            maxlength7      =   9
            tooltiptext7    =   "Rm I Rt"
            colname8        =   "WRMIVAL"
            heading8        =   "RmIVal       "
            datatype8       =   2
            mask8           =   "#####0.#0"
            maxlength8      =   9
            tooltiptext8    =   "Rm Inv val"
            colname9        =   "WSETSCD"
            heading9        =   "SetScd          "
            maxlength9      =   8
            tooltiptext9    =   "Set sub Code"
            colname10       =   "WSETIRT"
            heading10       =   "SetIRt    "
            datatype10      =   2
            mask10          =   "###0.#0"
            maxlength10     =   7
            tooltiptext10   =   "Inv Set Rt"
            colname11       =   "WSETIVAL"
            heading11       =   "SetIVal    "
            datatype11      =   2
            mask11          =   "####0.#0"
            maxlength11     =   8
            tooltiptext11   =   "Inv Set Val"
            colname12       =   "WMAINMET"
            heading12       =   "MainMet"
            datatype12      =   4
            maxlength12     =   1
            tooltiptext12   =   "MainMetal Y/N "
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   10
            Left            =   840
            TabIndex        =   71
            ToolTipText     =   "Enter Fr Voucher Tc"
            Top             =   150
            Width           =   675
            _ExtentX        =   1191
            _ExtentY        =   503
            Enabled         =   0   'False
            MaxLength       =   3
            DataType        =   4
            IdName          =   "WODTC"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   11
            Left            =   1620
            TabIndex        =   72
            ToolTipText     =   "Enter Fr Voucher Year"
            Top             =   150
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   503
            Enabled         =   0   'False
            MaxLength       =   2
            DataType        =   4
            IdName          =   "WODYY"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   12
            Left            =   2070
            TabIndex        =   73
            ToolTipText     =   "Enter Fr Voucher Chr"
            Top             =   150
            Width           =   675
            _ExtentX        =   1191
            _ExtentY        =   503
            Enabled         =   0   'False
            MaxLength       =   3
            DataType        =   4
            IdName          =   "WODCHR"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   13
            Left            =   3780
            TabIndex        =   74
            ToolTipText     =   "Enter Fr Voucher Sr"
            Top             =   150
            Width           =   795
            _ExtentX        =   1402
            _ExtentY        =   503
            Enabled         =   0   'False
            Mask            =   "#####0"
            MaxLength       =   6
            DataType        =   1
            IdName          =   "WODSR"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   14
            Left            =   2850
            TabIndex        =   75
            ToolTipText     =   "Enter Fr Voucher No"
            Top             =   150
            Width           =   795
            _ExtentX        =   1402
            _ExtentY        =   503
            Enabled         =   0   'False
            Mask            =   "#####0"
            MaxLength       =   6
            DataType        =   1
            IdName          =   "WODNO"
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "/"
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
            Height          =   240
            Left            =   3645
            TabIndex        =   80
            Top             =   195
            Width           =   135
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "/"
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
            Height          =   240
            Index           =   3
            Left            =   2730
            TabIndex        =   79
            Top             =   195
            Width           =   135
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "/"
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
            Height          =   240
            Index           =   1
            Left            =   1950
            TabIndex        =   78
            Top             =   195
            Width           =   135
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "/"
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
            Height          =   240
            Left            =   1500
            TabIndex        =   77
            Top             =   195
            Width           =   135
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Vch No"
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
            TabIndex        =   76
            Top             =   150
            Width           =   750
         End
      End
   End
End
Attribute VB_Name = "EmrFrmOrdZoom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---- General Comments on the Entry --------'
Option Explicit
Dim wb_WiwFlag As Boolean, wb_BagFlag As Boolean
Dim moCn As MwfLib.MDOConnection   '*** (Bef speed) bef 24/11/06
Dim ws_TcFr As String, ws_TcTo As String, ws_YyFr As String, ws_YyTo As String, ws_CoCdFr As String, ws_CoCdTo As String
Dim ws_ChrFr As String, ws_ChrTo As String, ws_CmCdFr As String, ws_CmCdTo As String
Dim ws_CdCustDmCdFr As String, ws_CdCustDmCdTo As String
Dim wi_NoFr As Long, wi_NoTo As Long, wi_SrFr As Long, wi_SrTo As Long
Dim w_OmDtFr As Date, w_OmDtTo As Date, w_DelDtFr As Date, w_DelDtTo As Date
Dim ws_OmPoNoFr As String, ws_OmPoNoTo As String, ws_DmCtgFr As String, ws_DmCtgTo As String
Dim ws_OdSubCustFr As String, ws_OdSubCustTo As String, ws_OdPoNoFr As String, ws_OdPoNoTo As String
Dim ws_DmCdFr As String, ws_DmCdTo As String, ws_ShPnd As String, ws_ShVal As String, ws_Pwd As String
Dim ws_SortOnDmCd As String       ' ***** Manali 3.10.0 - 09/03/12 - Sort on DmCd in GrdOrdPnd
Dim ws_CoCd As String       ' ***** Manali 3.10.0 - Allow PDCM in Zoom
Dim wb_OdSrOpt As Boolean   '****** Sachin 3.01
Dim wPic3DHt As Single
Dim wPic3DWdt As Single
Dim wPicHskHt As Single
Dim wPicHskWdt As Single
Dim wFraPic3DHt As Single
Dim wFraPic3DWdt As Single
Dim wFraPicHskHt As Single
Dim wFraPicHskWdt As Single
Dim wFraPic3DLft As Single
Dim wFraPic3DTop As Single
Dim wFraPicHskLft As Single
Dim wFraPicHskTop As Single
Dim wPic3DDblClk As Boolean
Dim wPicHskDblClk As Boolean
Dim wHskPth As String
Dim w3DPth As String
Dim ms_PwdFlag As Boolean

Dim wb_BagRmFlag As Boolean     ' ***** Manali 3.8.0 - Bag Rm Details

'Enum en_OrdZm
'  DsgDet = 0
'End Enum

'*** 2.11 (Jen)
Dim ms_BagJoin As String
Dim mi_FraKeysHt As Integer
'*** 2.11 (Jen)

'*** Jay 3.01 *** [Custom Views]
Enum en_OrdDsgFra
  DsgDet = 0
  InvDet = 1
  ViewSel = 2     '****** Sachin 3.3.0 (26-06-08) - [21. Copy From Order Zoom]
End Enum

Dim mf_OrdMstNo As Single, mf_OrdNo As Single, mf_Sr As Single
Dim mb_CrmDb As Boolean, mb_CTMod As Boolean
'*** Jay 3.01 *** [Custom Views]

Dim wLoad As Boolean

Private Sub ADC_Load()
  '*** Set the First Key Control property to the working field wDtTag
  '*** Set the First Non Key Control property to the grid GrdDt
  '*** Set the Child Property Of ADC to GrdDt
  '*** Set the Previous and Next Control Properties for the Grid GrdDt
  '*** Set the Hot Key for the Grid GrdDt as 'A'
  Set moCn = ADC.Connection    '*** (Bef speed) bef 24/11/06
  
  FRA_KEYS.ZOrder (0)
  Set ADC.FirKeyCtl = ADC("wCoCdFr")
  Set ADC.FirNKeyCtl = GrdOrdMst
  GrdOrdMst.HotKey = "A"
  GrdOrdPnd.HotKey = "B"
  GrdInv.HotKey = "B"
  GrdOrdWIW.HotKey = "C"
  GrdInvBag.HotKey = "C"
  GrdOrdBag.HotKey = "D"
  GrdInvRm.HotKey = "D"
  GrdOrdBagRm.HotKey = "F"  ' ***** Manali 3.8.0 - Bag Rm Details

  If mi_FraKeysHt = 0 Then mi_FraKeysHt = FRA_KEYS.Height
  
  '****** Sachin 3.01 - 01-06-2007 - Custom Views
  GrdOrdRm.AllowAdd = False: GrdOrdRm.AllowDelete = False
  GrdOrdLab.AllowAdd = False: GrdOrdLab.AllowDelete = False
  wb_OdSrOpt = False
  '****** Sachin 3.01 - 01-06-2007 - Custom Views
  
  'FRA_KEYS.Height = 2595
  wb_WiwFlag = True
  wb_BagFlag = True
  wb_BagRmFlag = True ' ***** Manali 3.8.0
  w_OmDtFr = "01/01/80"
  w_OmDtTo = "01/01/80"
  w_DelDtFr = "01/01/80"
  w_DelDtTo = "01/01/80"

 '*** Storing The Original Size And Position Of The Picture Controls ***
  wPic3DHt = Pic3D.Height
  wPic3DWdt = Pic3D.Width
  wPicHskHt = PicHsk.Height
  wPicHskWdt = PicHsk.Width
  
  wFraPic3DHt = FraPic3D.Height
  wFraPic3DWdt = FraPic3D.Width
  wFraPicHskHt = FraPicHsk.Height
  wFraPicHskWdt = FraPicHsk.Width
 
  wFraPic3DLft = FraPic3D.Left
  wFraPic3DTop = FraPic3D.Top
  wFraPicHskLft = FraPicHsk.Left
  wFraPicHskTop = FraPicHsk.Top
  
  wPic3DDblClk = False
  wPicHskDblClk = False
   
  '*** Jay 2.14 *** (SEO Change)
  ms_PwdFlag = moCn.RecSeek("Select PValue1 from Param where PTyp= 'SEOPWD' and PmCd= 'SEOPWD' and PValue1= '' ")
  '*** Jay 2.14 *** (SEO Change)
   
  '*** (Bef 2.11)
  'Dim wFineNm As Boolean
  'wFineNm = IIF(InStr(UCase(gs_CoNm), "FINE ") <= 0, False, True)
  'cmd_InvDet.Visible = wFineNm
  '*** (Bef 2.11)
  '*** Jay 3.01 *** [Custom View]
  mb_CTMod = moCn.RecSeek("Select 'x' From Head Where HCd='" + ctSelfCmCd + "' And HCoCd='" + gs_CoCd + "' " + _
              "and hCompBagsYn = 'Y'")
  mb_CrmDb = moCn.RecSeek("Select 'x' from SysDb where SysId= 'CRM' And SysDb ='WEB'")
  '*** Jay 3.01 *** [Custom View]
  
  '******** Manali 3.2.2 - Default value for OmTc, Yy and Chr
  ws_CoCdFr = gs_CoCd: ws_CoCdTo = gs_CoCd
  ws_TcFr = "SO": ws_TcTo = "SO"
  ws_YyFr = moCn.GetFldVal("Select HYy From Head Where HCd='" + ctSelfCmCd + "' And HCoCd='" + gs_CoCd + "'")
  ws_YyTo = moCn.GetFldVal("Select HYy From Head Where HCd='" + ctSelfCmCd + "' And HCoCd='" + gs_CoCd + "'")
  ws_ChrFr = moCn.GetFldVal("Select vPValue from vParam where vPCoCd='" + gs_CoCd + "' and vPTyp='DEFCHR' and vPMCd='" + ws_TcFr + "' and vPValue<> ''")
  ws_ChrTo = moCn.GetFldVal("Select vPValue from vParam where vPCoCd='" + gs_CoCd + "' and vPTyp='DEFCHR' and vPMCd='" + ws_TcTo + "' and vPValue<> ''")
  '******** Manali 3.2.2 - Default value for OmTc, Yy and Chr

  '      IG.211 - By default, Del Dt option to be set as "P".
  ADC("wDelDt") = "P"
  OptDelDt(1).Value = True

End Sub
Private Sub ADC_KeyWhen()

  '*** Set the frame FraNKeyAll as the top most frame
  
  '****** Sachin 2.14.0 - [24-11-2006] - [Multi-Company] ******
  Me.Caption = GetFrmCaption(gs_CoCd, "Order Zoom")
  '****** Sachin 2.14.0 - [24-11-2006] - [Multi-Company] ******
  
  FraNKeyAll.ZOrder (0)
  GrdOrdMst.Rows = 1        '****** Sachin 3.01 - Custom Views
  GrdOrdPnd.Rows = 1
  GrdOrdWIW.Rows = 1
  GrdOrdBag.Rows = 1
  GrdOrdBagRm.Rows = 1      ' ***** Manali 3.8.0 - Bag Rm Details
  GrdInv.Rows = 1
  GrdInvBag.Rows = 1
  GrdInvRm.Rows = 1
  'FRA_KEYS.Height = 2595
  FRA_KEYS.Height = mi_FraKeysHt
  
  cmd_InvDet.Enabled = False
  CmdDsgDet.Enabled = False
  
  '****** Sachin 3.3.0 (27-06-08) - [21. Copy From Order Zoom] ******
  CmdViewSel.Enabled = False
  '****** Sachin 3.3.0 (27-06-08) - [21. Copy From Order Zoom] ******
  
  wb_OdSrOpt = False
    
  Call ShowPic(0, 0, "", PicHsk)
  Call ShowPic(0, 0, "", Pic3D)

  FraPic3D.Height = wFraPic3DHt
  FraPic3D.Width = wFraPic3DWdt
  FraPicHsk.Height = wFraPicHskHt
  FraPicHsk.Width = wFraPicHskWdt
    
  FraPic3D.Left = wFraPic3DLft
  FraPic3D.Top = wFraPic3DTop
  FraPicHsk.Left = wFraPicHskLft
  FraPicHsk.Top = wFraPicHskTop
  
  ADC("wCoCdFr") = ws_CoCdFr
  ADC("wCoCdTo") = ws_CoCdTo
  ADC("wOdTcFr") = ws_TcFr
  ADC("wOdTcTo") = ws_TcTo
  ADC("wOdYyFr") = ws_YyFr
  ADC("wOdYyTo") = ws_YyTo
  ADC("wOdChrFr") = ws_ChrFr
  ADC("wOdChrTo") = ws_ChrTo
  ADC("wOdNoFr") = wi_NoFr
  ADC("wOdNoTo") = wi_NoTo
  ADC("wOdSrFr") = wi_SrFr
  ADC("wOdSrTo") = wi_SrTo
  
  ADC("wOmDtFr") = w_OmDtFr
  ADC("wOmDtTo") = w_OmDtTo
  ADC("wOmDelDtFr") = w_DelDtFr
  ADC("wOmDelDtTo") = w_DelDtTo
     
  ADC("wOmCmCdFr") = ws_CmCdFr
  ADC("wOmCmCdTo") = ws_CmCdTo
    
  ADC("wCdCustDmCdFr") = ws_CdCustDmCdFr
  ADC("wCdCustDmCdTo") = ws_CdCustDmCdTo
  
  ADC("wOmPoNoFr") = ws_OmPoNoFr
  ADC("wOmPoNoTo") = ws_OmPoNoTo
  
  'VK.19 - Restore values for Sub Cust and Sub PO
  ADC("wOdSubCustFr") = ws_OdSubCustFr
  ADC("wOdSubCustTo") = ws_OdSubCustTo
  ADC("wOdPoNoFr") = ws_OdPoNoFr
  ADC("wOdPoNoTo") = ws_OdPoNoTo
  
  ADC("wOdDsgCtgFr") = ws_DmCtgFr
  ADC("wOdDsgCtgTo") = ws_DmCtgTo
  ADC("wOdDmCdFr") = ws_DmCdFr
  ADC("wOdDmCdTo") = ws_DmCdTo
  
  ADC("wYn") = IIF(ws_ShPnd <> "", ws_ShPnd, "Y")
  ADC("wPwdYn") = IIF(ws_ShVal <> "", ws_ShVal, "N")
  ADC("wDelDt") = IIF(OptDelDt(0).Value = True, "E", "P")
  TxtPwd.text = IIF(ws_Pwd <> "", ws_Pwd, "")
  
  ADC("wSortDmCd") = IIF(ws_SortOnDmCd <> "", ws_SortOnDmCd, "N") ' ***** Manali 3.10.0 - 09/03/12 - Sort on DmCd in GrdOrdPnd
      
  ' ***** Manali 3.5.0 - 06/05/09 - Check Boxes Added
  If ADC("wYn") = "Y" Then ChkBoxArr(0).Value = Checked Else ChkBoxArr(0).Value = Unchecked
  If ADC("wPwdYn") = "Y" Then ChkBoxArr(1).Value = Checked Else ChkBoxArr(1).Value = Unchecked
  If ADC("wSortDmCd") = "Y" Then ChkBoxArr(2).Value = Checked Else ChkBoxArr(2).Value = Unchecked       ' ***** Manali 3.10.0 - 09/03/12 - Sort on DmCd in GrdOrdPnd
  'If adc("wRejYN") = "Y" Then ChkBoxArr(3).Value = Checked Else ChkBoxArr(3).Value = Unchecked       ' ***** Manali 3.10.0 - 09/03/12 - Sort on DmCd in GrdOrdPnd
  ' ***** Manali 3.5.0 - 06/05/09 - Check Boxes Added
    
  '*** Jay 2.13(CT) ***
  '  '*** (Jen 2.11)
  '  If moCn.GetFldVal("Select hFluteBagYN from Head where HCoCd= '" + gs_CoCd + "' and " + _
  '                                                    "HCd= '" + ctSelfCmCd + "' ") = "N" Then
  '    LblPrimFluteBags.Visible = False
  '    ADC("wBagTyp").Visible = False
  '    LblUoBChrFr.Visible = False
  '    ADC("wBagChr").Visible = False
  '    ' Zubin 212
  '    'ADC("wBagTyp") = "N": ADC("wBagChr") = ""
  '    ' Zubin 212
  '  End If
  '  '*** (Jen 2.11)
  ' Zubin 212
  ADC("wBagTyp") = "P": ADC("wBagChr") = ""
  ' Zubin 212
  '*** Jay 2.13(CT) ***
  mf_OrdMstNo = 0: mf_OrdNo = 0: mf_Sr = 0       '*** Jay 3.01 *** [Custom View]
End Sub
Private Sub ADC_KeyValid(pr_Cancel As Boolean, pr_ErrMsg As String)

  'IG.211 - User level checking to be done. Header and Tool tip to be changed based on scope.
  
  If GetUsrLevel > 5 And ADC("wDelDt") = "E" Then
    pr_Cancel = True: pr_ErrMsg = "User don't have access to view Export Delivery Date": Exit Sub
  End If
  If ADC("wDelDt") = "E" Then
    GrdOrdMst.ColProp("wOmEPDelDt").Heading = "ExpDelDt"
    GrdOrdMst.ColProp("wOmEPDelDt").ToolTipText = "Export Delivery Date"
  Else
    GrdOrdMst.ColProp("wOmEPDelDt").Heading = "PrdDelDt"
    GrdOrdMst.ColProp("wOmEPDelDt").ToolTipText = "Production Delivery Date"
  End If
  If ADC("wDelDt") = "E" Then
    GrdOrdPnd.ColProp("wOpExpDt").Heading = "ExpDelDt"
    GrdOrdPnd.ColProp("wOpExpDt").ToolTipText = "Export Delivery Date"
  Else
    GrdOrdPnd.ColProp("wOpExpDt").Heading = "PrdDelDt"
    GrdOrdPnd.ColProp("wOpExpDt").ToolTipText = "Production Delivery Date"
  End If
  
  '*** Check whether valid
  ' ***** Manali 3.10.0 - 09/03/12 - Allow PDCM in Zoom
  If ADC("wCoCdFr") = "" Or ADC("wCoCdTo") = "" Then pr_Cancel = True: pr_ErrMsg = "Company Codes Cannot Be Blank": Exit Sub   '**Bhavna
  If ADC("wOdTcFr") = "" Or ADC("wOdTcTo") = "" Then pr_Cancel = True: pr_ErrMsg = "Tcs Cannot Be Blank": Exit Sub
  If ADC("wOdTcTo") <> ADC("wOdTcFr") Then pr_Cancel = True: pr_ErrMsg = "Both Tcs Should Be Same": Exit Sub
  ' ***** Manali 3.10.0 - 09/03/12 - Allow PDCM in Zoom
    
  If (ADC("wOdSrFr") <> 0 Or ADC("wOdSrTo") <> 0) And _
     ((ADC("wOdTcFr") = "" Or ADC("wOdYyFr") = "" Or ADC("wOdChrFr") = "" Or ADC("wOdNoFr") = 0) _
       Or _
       (ADC("wOdTcFr") <> ADC("wOdTcTo") Or ADC("wOdYyFr") <> ADC("wOdYyTo") Or ADC("wOdChrFr") <> ADC("wOdChrTo") _
            Or ADC("wOdNoFr") <> ADC("wOdNoTo")) _
     ) Then pr_Cancel = True: pr_ErrMsg = "Design Sr Can Be Specified Only For Single Order": Exit Sub
     
End Sub

Private Sub ADC_NonKeyValid(pr_Cancel As Boolean, pr_ErrMsg As String)
  '****** Sachin 3.3.0 (26-06-08) - [21. Copy From Order Zoom] ******
  gs_OrdZoomSel = ""
  '****** Sachin 3.3.0 (26-06-08) - [21. Copy From Order Zoom] ******
End Sub

Private Sub ADC_NonKeyWhen(pr_Cancel As Boolean, pr_ErrMsg As String)

  FRA_KEYS.Height = 560
  ws_CoCdFr = ADC("wCoCdFr")
  ws_CoCdTo = ADC("wCoCdTo")
  ws_TcFr = ADC("wOdTcFr")
  ws_TcTo = ADC("wOdTcTo")
  ws_YyFr = ADC("wOdYyFr")
  ws_YyTo = ADC("wOdYyTo")
  ws_ChrFr = ADC("wOdChrFr")
  ws_ChrTo = ADC("wOdChrTo")
  wi_NoFr = ADC("wOdNoFr")
  wi_NoTo = ADC("wOdNoTo")
  wi_SrFr = ADC("wOdSrFr")
  wi_SrTo = ADC("wOdSrTo")
  
  w_OmDtFr = IIF(ADC("wOmDtFr") <> "", ADC("wOmDtFr"), "01/01/80")
  w_OmDtTo = IIF(ADC("wOmDtTo") <> "", ADC("wOmDtFr"), "01/01/80")
  w_DelDtFr = IIF(ADC("wOmDelDtFr") <> "", ADC("wOmDelDtFr"), "01/01/80")
  w_DelDtTo = IIF(ADC("wOmDelDtTo") <> "", ADC("wOmDelDtFr"), "01/01/80")
  
  ws_CmCdFr = ADC("wOmCmCdFr")
  ws_CmCdTo = ADC("wOmCmCdTo")
  
  ws_CdCustDmCdFr = ADC("wCdCustDmCdFr")
  ws_CdCustDmCdTo = ADC("wCdCustDmCdTo")
  
  ws_OmPoNoFr = ADC("wOmPONoFr")
  ws_OmPoNoTo = ADC("wOmPoNoTo")
  
  'VK.19 - Store values for Sub Cust and Sub PO
  ws_OdSubCustFr = ADC("wOdSubCustFr")
  ws_OdSubCustTo = ADC("wOdSubCustTo")
  ws_OdPoNoFr = ADC("wOdPoNoFr")
  ws_OdPoNoTo = ADC("wOdPoNoTo")
  
  ws_DmCtgFr = ADC("wOdDsgctgFr")
  ws_DmCtgTo = ADC("wOdDsgctgTo")
  ws_DmCdFr = ADC("wOdDmCdFr")
  ws_DmCdTo = ADC("wOdDmCdTo")
  ws_ShPnd = ADC("wYn")
  ws_ShVal = ADC("wPwdYN")
  ws_Pwd = TxtPwd.text
  ws_SortOnDmCd = ADC("wSortDmCd")       ' ***** Manali 3.10.0 - 09/03/12 - Sort on DmCd in GrdOrdPnd
  
  
  'Bef *** Jay 3.01 *** [Custom View]
  'If ADC("WODTCFR") = ADC("WODTCTO") And ADC("WODYYFR") = ADC("WODYYTO") And ADC("WODCHRFR") = ADC("WODCHRTO") And ADC("WODNOFR") = ADC("WODNOTO") Then wb_OdSrOpt = True
  'Bef *** Jay 3.01 *** [Custom View]
  Call FillGrdOrdMst
'****** Before 3.01 ******
'  If GrdOrdPnd.Rows > 1 Then
'    GrdOrdPnd.Row = 1: GrdOrdPnd.SetFocus
'    cmd_InvDet.Enabled = True
'  Else
'    GrdOrdWIW.Rows = 1: GrdOrdBag.Rows = 1
'  End If
  
  If GrdOrdMst.Rows > 1 Then
    GrdOrdMst.Row = 1: GrdOrdMst.SetFocus
    cmd_InvDet.Enabled = True
    CmdDsgDet.Enabled = True
    '****** Sachin 3.3.0 (27-06-08) - [21. Copy From Order Zoom] ******
    CmdViewSel.Enabled = True
    '****** Sachin 3.3.0 (27-06-08) - [21. Copy From Order Zoom] ******
  Else
    GrdOrdPnd.Rows = 1: GrdOrdWIW.Rows = 1: GrdOrdBag.Rows = 1
    GrdOrdBagRm.Rows = 1      ' ***** Manali 3.8.0 - Bag Rm Details
  End If
  
  '*** Jay 2.13(CT) ***
  Dim wHeadChk As MwfLib.MDORowSet
  Set wHeadChk = moCn.OpenRes("Select hFluteBagYN, hCompBagsYn from Head where HCoCd= '" + gs_CoCd + "' and " + _
                                                    "HCd= '" + ctSelfCmCd + "' ")
  If wHeadChk!hFluteBagYN = "N" And wHeadChk!hCompBagsYn = "N" Then
    LblBagTyp.Visible = False
    ADC("wBagTyp").Visible = False
    LblUoBChrFr.Visible = False
    ADC("wBagChr").Visible = False
    ADC("wBagTyp") = "P": ADC("wBagChr") = ""
  End If
  Set wHeadChk = Nothing
  '*** Jay 2.13(CT) ***
  
    '*** Jay 2.13(CT) ***
    ms_BagJoin = " Join Param On PTyp= 'BCHR' and PMCd= BChr and PSCd= '' " + _
                      IIF(ADC("wBagTyp") <> "", " and PValue= '" + ADC("wBagTyp") + "' ", " ") + _
                      IIF(ADC("wBagChr") <> "", " and PMCd= '" + ADC("wBagChr") + "' ", " ")
    If ADC("wBagTyp") = "" Then ms_BagJoin = " "
  
    '  '*** Jenny 2.11 ***
    '  ms_BagJoin = " "
    '  If UCase(ADC("wBagTyp")) = "N" Then
    '    ms_BagJoin = " Join Param On PTyp= 'BCHR' and PMCd= BChr and PSCd= '' and PValue = 'N' " + _
    '              IIF(ADC("wBagChr") <> "", " and PMCd= '" + ADC("wBagChr") + "' ", " ")
    '  ElseIf UCase(ADC("wBagTyp")) = "Y" Then
    '    ms_BagJoin = " Join Param On PTyp= 'BCHR' and PMCd= BChr and PSCd= '' and PValue = 'Y' " + _
    '              IIF(ADC("wBagChr") <> "", " and PMCd= '" + ADC("wBagChr") + "' ", " ")
    '  'Else
    '    'ms_BagJoin = " Join Param On PTyp= 'BCHR' and PMCd= BChr and PSCd= '' " + _
    '    '          IIF(ADC("wBagChr") <> "", " and PMCd= '" + ADC("wBagChr") + "' ", " ")
    '  End If
    '  '*** Jenny 2.11 ***
    '*** Jay 2.13(CT) ***
  '*** Jay 3.01 *** [Custom View]
  ' **** Manali 3.03
  'If mb_CTMod And (adc("wOdTcFr") = "SO" Or adc("wOdTcFr") = "PL") Then
  If mb_CTMod Then
    TabOrdDsgDet.TabVisible(3) = True:    TabOrdDsgDet.TabsPerRow = 4
  Else
    TabOrdDsgDet.TabVisible(3) = False:   TabOrdDsgDet.TabsPerRow = 3
  End If
  If mb_CrmDb Then LblCrmFixPrcYN.Visible = True: ADC("wOdCrmFixPrcYN").Visible = True: FraGrdDet.Visible = True
  '*** Jay 3.01 *** [Custom View]
End Sub
Private Sub ADC_SetRecSource()
    '*** Since this is an 'A' Entry, there is no RecSource associated with ADC. So set the
    'DataMode property of ADC control as xdatUnbound
  ADC.DataMode = xdatUnbound
  
  'IG.206 - When Show Value is checked in Scope, show the cost price Value in DsgDet frame. Otherwise make invisible of Costprice value.
  If ADC("wPwdYN") <> "Y" Then
    LblCstPrc.Visible = False
    ADC("wOdCstPrc").Visible = False
  Else
    LblCstPrc.Visible = True
    ADC("wOdCstPrc").Visible = True
  End If
    
End Sub

'IG.211 - Y - Exp Del Dt and N - Prd Del Dt
Private Sub OptDelDt_Click(Index As Integer)
Select Case Index
Case Is = 0
  ADC("wDelDt") = "E"
Case Is = 1
  ADC("wDelDt") = "P"
End Select
End Sub

Private Sub ATXT_FldChange(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String, ByVal pv_NewValue As Variant)
 Select Case UCase(IdName)
  Case Is = UCase("wYN")
    '*** Jay 3.01 *** [Custom View]
    Cancel = (Not moCn.RecSeek("select PMCd from Param where " + _
                         " PTyp= 'YN' and PMCd = '" + pv_NewValue + "'"))
    'Cancel = pv_NewValue <> "" And (Not moCn.RecSeek("select PMCd from Param where " + _
                         " PTyp= 'YN' and PMCd = '" + pv_NewValue + "'"))
    '*** Jay 3.01 *** [Custom View]
    ErrMsg = "Show Pending Orders/All Orders (Y/N)": Exit Sub
  Case Is = UCase("wOdTcFr")
    'If ADC("wOdTcTo") = "" Then   *********  Manali 3.2.2
    ADC("wOdTcTo") = pv_NewValue
    '***** Manali 3.10.0 - 09/03/12 - Allow PDCM in Zoom
    ws_CoCd = IIF(moCn.RecSeek("Select 'x' from Param Where PTyp='TC' and PMCd='" + pv_NewValue + "' and PValue='PL'"), ctSelfCoCd, gs_CoCd)
  
  Case Is = UCase("wOdYyFr")
    'If ADC("wOdYyTo") = "" Then   *********  Manali 3.2.2
    ADC("wOdYyTo") = pv_NewValue
  Case Is = UCase("wOdChrFr")
    'If ADC("wOdChrTo") = "" Then   *********  Manali 3.2.2
    ADC("wOdChrTo") = pv_NewValue
  Case Is = UCase("wOdNoFr")
    'If ADC("wOdNoTo") = 0 Then   *********  Manali 3.2.2
    ADC("wOdNoTo") = pv_NewValue
  Case Is = UCase("wOdSrFr")
    'If ADC("wOdSrTo") = 0 Then   *********  Manali 3.2.2
    ADC("wOdSrTo") = pv_NewValue
  Case Is = UCase("wOmDtFr")
    'If ADC("wOmDtTo") = "01/01/80" Then   *********  Manali 3.2.2
    ADC("wOmDtTo") = pv_NewValue
  Case Is = UCase("wOmDelDtFr")
    'If ADC("wOmDelDtTo") = "01/01/80" Then   *********  Manali 3.2.2
    ADC("wOmDelDtTo") = pv_NewValue
  Case Is = UCase("wOmCmCdFr")
    'If ADC("wOmCmCdTo") = "" Then   *********  Manali 3.2.2
    ADC("wOmCmCdTo") = pv_NewValue
  Case Is = UCase("wOmCmCdTo")
    'Creation.12.a - Customer Design Code should be empty when different Customer Code
    If UCase(Trim(ADC("wOmCmCdFr"))) <> UCase(Trim(pv_NewValue)) Then
      ADC("wCdCustDmCdFr") = ""
      ADC("wCdCustDmCdTo") = ""
    End If
  Case Is = UCase("wOmPoNoFr")
    'If ADC("wOmPoNoTo") = "" Then   *********  Manali 3.2.2
    ADC("wOmPoNoTo") = pv_NewValue
  Case Is = UCase("wOdDsgCtgFr")
    'If ADC("wOdDsgCtgTo") = "" Then   *********  Manali 3.2.2
    ADC("wOdDsgCtgTo") = pv_NewValue
  Case Is = UCase("wOdDmCdFr")
    'If ADC("wOdDmCdTo") = "" Then   *********  Manali 3.2.2
    ADC("wOdDmCdTo") = pv_NewValue
  
  '****** Sachin 3.01 - 02-06-07 - Custom Views ******
  Case Is = UCase("wPwdYN")
    If pv_NewValue = "Y" And ms_PwdFlag = False Then Cancel = True: ErrMsg = "Enter SEO Password to View Values ": Exit Sub
    Cancel = pv_NewValue <> "" And (Not moCn.RecSeek("select PMCd from Param where " + _
                         " PTyp= 'YN' and PMCd = '" + pv_NewValue + "'"))
    ErrMsg = "Show Value (Y/N)": Exit Sub
  '****** Sachin 3.01 - 02-06-07 - Custom Views ******
    
  '*** 2.11 (Jen)
  Case Is = UCase("wBagTyp")
    '*** Jay 2.13(CT) ***
    Cancel = pv_NewValue <> "" And (Not moCn.RecSeek("select PMCd from Param where " + _
                         " PTyp= 'BTYP' and PMCd = '" + pv_NewValue + "'"))
    If Cancel = True Then ErrMsg = "Invalid Bag Type": Exit Sub
    '*** Jay 2.13(CT) ***
  Case Is = UCase("wBagChr")
    Cancel = pv_NewValue <> "" And (Not moCn.RecSeek("select PMCd from Param where " + _
                         " PTyp= 'BCHR' and PMCd = '" + pv_NewValue + "' and PValue= '" + ADC("wBagTyp") + "'"))
    If Cancel = True Then ErrMsg = "Invalid Bag Character": Exit Sub
  '*** 2.11 (Jen)
  'VK.19 - While entering Sub Cust From, it copied to Sub Cust To
  Case Is = UCase("wOdSubCustFr")
    ADC("wOdSubCustTo") = pv_NewValue
    
  Case Is = UCase("wOdPoNoFr")
    ADC("wOdPoNoTo") = pv_NewValue
  End Select
End Sub
'*** Jay 3.01 *** [Custom View]
Private Sub atxtDisp_FldWhen(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String)
  Cancel = True
End Sub
Private Sub ATXT_FldWhen(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String)
  
  '****** Sachin 3.3.0 (26-06-08) - [21. Copy From Order Zoom]******
  If UCase(IdName) = "WZOOMSEL" Then Cancel = True
  '****** Sachin 3.3.0 (26-06-08) - [21. Copy From Order Zoom]******
  
  Select Case UCase(IdName)
  '********Bhavna added co scope
  Case Is = UCase("wCoCdFr"), UCase("wCoCdTo")     '***** Bhavna added fr, to co cd
    Call HlpList.hCoCd
  '********Bhavna added co scope
  Case Is = UCase("wOdTcFr"), UCase("wOdTcTo")
    Call HlpList.PMCd("TC", "'SO','QT','PL','QS'")
  Case Is = UCase("wOdYyFr"), UCase("wOdYyTo")
    Call HlpList.PMCd("YY")
  Case Is = UCase("wOdChrFr")
    Call HlpList.vPSCd(ADC("wCoCdFr"), "CHR", ADC("wOdTcFr"))    ' ***** Manali 3.10.0 - 09/03/12 -  Allow PDCM in Zoom
    'Call HlpList.vPSCd(gs_CoCd, "CHR", ADC("wOdTcFr"))
  Case Is = UCase("wOdChrTo")
    Call HlpList.vPSCd(ADC("wCoCdTo"), "CHR", ADC("wOdTcTo"))    ' ***** Manali 3.10.0 - 09/03/12 -  Allow PDCM in Zoom
    'Call HlpList.vPSCd(gs_CoCd, "CHR", ADC("wOdTcTo"))
  Case Is = UCase("wOdNoFr")
    Call HlpList.OmNo(ADC("wCoCdFr"), ADC("wOdTcFr"), ADC("wOdYyFr"), ADC("wOdChrFr"))    ' ***** Manali 3.10.0 - 09/03/12 -  Allow PDCM in Zoom
    'Call HlpList.OmNo(gs_CoCd, ADC("wOdTcFr"), ADC("wOdYyFr"), ADC("wOdChrFr"))
  Case Is = UCase("wOdNoTo")
    Call HlpList.OmNo(ADC("wCoCdTo"), ADC("wOdTcTo"), ADC("wOdYyTo"), ADC("wOdChrTo"))     ' ***** Manali 3.10.0 - 09/03/12 -  Allow PDCM in Zoom
    'Call HlpList.OmNo(gs_CoCd, ADC("wOdTcTo"), ADC("wOdYyTo"), ADC("wOdChrTo"))
  Case Is = UCase("wOdSrFr")
    Call HlpList.OdSr(ADC("wCoCdFr"), ADC("wOdTcFr"), ADC("wOdYyFr"), ADC("wOdChrFr"), ADC("wOdNoFr"))   ' ***** Manali 3.10.0 - 09/03/12 -  Allow PDCM in Zoom
    'Call HlpList.OdSr(gs_CoCd, ADC("wOdTcFr"), ADC("wOdYyFr"), ADC("wOdChrFr"), ADC("wOdNoFr"))
  Case Is = UCase("wOdSrTo")
    Call HlpList.OdSr(ADC("wCoCdTo"), ADC("wOdTcTo"), ADC("wOdYyTo"), ADC("wOdChrTo"), ADC("wOdNoTo"))   ' ***** Manali 3.10.0 - 09/03/12 -  Allow PDCM in Zoom
    'Call HlpList.OdSr(gs_CoCd, ADC("wOdTcTo"), ADC("wOdYyTo"), ADC("wOdChrTo"), ADC("wOdNoTo"))
  Case Is = UCase("wOmCmCdFr"), UCase("wOmCmCdTo")
    Call HlpList.CustCd("C")
  Case Is = UCase("wOdDsgCtgFr"), UCase("wOdDsgCtgTo")
    Call HlpList.PMCd("DmCtg")
  Case Is = UCase("wOdDmCdFr"), UCase("wOdDmCdTo")
    Call HlpList.DmCd("DM")
  Case Is = UCase("wYN")
    Call HlpList.PMCd("YN")
  'Creation.12.a - Help for CustDmCdFrom & To
  Case Is = UCase("wCdCustDmCdFr"), UCase("wCdCustDmCdTo")
    If ((ADC("wOmCmCdFr") = "") Or ADC("wOmCmCdTo") = "") Or ADC("wOmCmCdFr") <> ADC("wOmCmCdTo") Then Cancel = True: ErrMsg = "Cannot Enter Customer Design Code as From & To Customer code is either blank or not same": Exit Sub
      Call HlpList.CustDmCd(ADC("wOmCmCdFr"))
  '*** Jenny 2.11 ***
  Case Is = UCase("wBagTyp")
    HlpList.PMCd "BTYP"       '*** Jay 2.13(CT) ***
  Case Is = UCase("wBagChr")
    If ADC.Mode = xNorm Then
      If ADC("wBagTyp") = "" Then Cancel = True: ErrMsg = "First Specify Bag Type": Exit Sub  '*** Jay 2.13(CT) ***
      HlpList.PMCd "BCHR", "'" + ADC("wBagTyp") + "'": Exit Sub
    End If
  '*** Jenny 2.11 ***
  'VK.19 - Help for Sub Cust
  Case Is = UCase("wOdSubCustFr"), UCase("wOdSubCustTo")
    HlpList.PMCd "SUBCUST"
  End Select
End Sub

Private Sub atxt_ReCalculate(Index As Integer, ByVal IdName As String, ByVal OldVal As Variant, ByVal RowNum As Integer, ByVal SourceName As String, ByVal SourceColName As String)
  '*** Jenny 2.11 ***
  Select Case UCase(IdName)
    Case Is = UCase("wBagChr")
      ADC("wBagChr") = ""
  End Select
  '*** Jenny 2.11 ***
End Sub

' ***** Manali 3.5.0 - 06/05/09 - Check Boxes Added
Private Sub ChkBoxArr_Click(Index As Integer)
With ChkBoxArr(Index)
  Select Case Index
  Case Is = 0     ' Tag = "ShowPnd"
    If .Value = Checked Then ADC("wYN") = "Y" Else ADC("wYN") = "N"
  Case Is = 1     ' Tag = "ShowVal"
    If .Value = Checked Then
      If ms_PwdFlag = False Then
        ChkBoxArr(1).Value = Unchecked: ADC("wPwdYN") = "N": DispMsg "Enter SEO Password To Show Value", etError
      Else
        ADC("wPwdYN") = "Y"
      End If
    Else
      ADC("wPwdYN") = "N"
    End If
    
  ' ***** Manali 3.10.0 - 09/03/12 - Sort on DmCd in GrdOrdPnd
  Case Is = 2     ' Tag = "SortDmCd"
    If .Value = Checked Then ADC("wSortDmCd") = "Y" Else ADC("wSortDmCd") = "N"
    
  End Select
End With
End Sub
' ***** Manali 3.5.0 - 06/05/09 - Check Boxes Added

' ***** Manali 3.5.0 - 06/05/09- Check Boxes Added
Private Sub ChkBoxArr_GotFocus(Index As Integer)
  DispMsg ChkBoxArr(Index).ToolTipText, etInfo
End Sub
' ***** Manali 3.5.0 - 06/05/09 - Check Boxes Added

Private Sub cmd_InvDet_Click()
  'open a frame fraOrdInv containing the grdOrdInv if closed else open it
  Call DispFra(InvDet)
End Sub
Private Sub CmdDsgDet_Click()
'*** Jay 3.01 *** [Custom View]
  Call DispFra(DsgDet)
  If mf_OrdNo <> GrdOrdMst.Row Or mf_Sr <> GrdOrdPnd.Row Then
    Call FillDsgDets(GrdOrdPnd.Row)
    Call FillOrdRm(GrdOrdPnd.Row)
    Call FillOrdLab(GrdOrdPnd.Row)
    If mb_CTMod Then Call FillOrdCt(GrdOrdPnd.Row)
    mf_OrdNo = GrdOrdMst.Row: mf_Sr = GrdOrdPnd.Row
  End If
'*** Jay 3.01 *** [Custom View]
End Sub

Private Sub CmdViewSel_Click()
  '****** Sachin 3.3.0 (26-06-08) - [21. Copy From Order Zoom]
  ADC("WZOOMSEL") = Replace(gs_OrdZoomSel, ",", vbCrLf)
  Call DispFra(ViewSel)
  '****** Sachin 3.3.0 (26-06-08) - [21. Copy From Order Zoom]
End Sub

Private Sub GrdInv_ColWhen(ByVal RowNum As Integer, ByVal ColName As String, Cancel As Boolean, ErrMsg As String, ByVal Action As MwfCtl.en_fgAction)
  Cancel = True
End Sub
Private Sub GrdInvBag_ColWhen(ByVal RowNum As Integer, ByVal ColName As String, Cancel As Boolean, ErrMsg As String, ByVal Action As MwfCtl.en_fgAction)
  Cancel = True
End Sub
Private Sub GrdInvRm_ColWhen(ByVal RowNum As Integer, ByVal ColName As String, Cancel As Boolean, ErrMsg As String, ByVal Action As MwfCtl.en_fgAction)
  Cancel = True
End Sub
Private Sub GrdOrdBag_ColWhen(ByVal RowNum As Integer, ByVal ColName As String, Cancel As Boolean, ErrMsg As String, ByVal Action As MwfCtl.en_fgAction)
  Cancel = True
End Sub

Private Sub GrdOrdBag_RowWhen(ByVal RowNum As Integer)
If RowNum > 0 Then
  Call FillGrdOrdBagRm(RowNum)   ' ***** Manali 3.8.0 - Bag Rm Details
  ADC("WWRKNAME") = GrdOrdBag.Value(RowNum, "WOBWRKNAME")   '***** Sachin 3.11.0
End If
End Sub

Private Sub GrdOrdBagRm_ColWhen(ByVal RowNum As Integer, ByVal ColName As String, Cancel As Boolean, ErrMsg As String, ByVal Action As MwfCtl.en_fgAction)
Cancel = True   ' ***** Manali 3.8.0 - Bag Rm Details
End Sub

Private Sub GrdOrdCT_ColWhen(ByVal RowNum As Integer, ByVal ColName As String, Cancel As Boolean, ErrMsg As String, ByVal Action As MwfCtl.en_fgAction)
  Cancel = True     ' ***** Manali 3.03
End Sub

Private Sub GrdOrdLab_ColWhen(ByVal RowNum As Integer, ByVal ColName As String, Cancel As Boolean, ErrMsg As String, ByVal Action As MwfCtl.en_fgAction)
  Cancel = True
End Sub
Private Sub GrdOrdMst_ColWhen(ByVal RowNum As Integer, ByVal ColName As String, Cancel As Boolean, ErrMsg As String, ByVal Action As MwfCtl.en_fgAction)
  Cancel = True
End Sub
Private Sub GrdOrdMst_RowWhen(ByVal RowNum As Integer)
  '***  Jay 3.01 *** [Custom View]
  If RowNum > 0 And mf_OrdMstNo <> GrdOrdMst.Row Then
      Call FillGrdOrdPnd(RowNum)
      mf_OrdMstNo = GrdOrdMst.Row
  End If
  '***  Jay 3.01 *** [Custom View]
End Sub

Private Sub GrdOrdPnd_ColChange(ByVal RowNum As Integer, ByVal ColName As String, ByVal NewValue As Variant, Cancel As Boolean, ErrMsg As String)
  '****** Sachin 3.3.0 (26-06-08) - [21. Order Copy From Zoom] ******
  Dim wZoomTxt As String, wChkSr As Integer
  With GrdOrdPnd
    Select Case ColName
      Case Is = UCase("WCPYODSR")
            If NewValue <> "Y" And NewValue <> "" Then Cancel = True: ErrMsg = "You Can Enter only 'Y' Or '' In this field ": Exit Sub
            wZoomTxt = .Value(RowNum, "WOPCOCD") + "/" + .Value(RowNum, "WOPTC") + "/" + .Value(RowNum, "WOPYY") + "/" + .Value(RowNum, "WOPCHR") + "/" + CStr(.Value(RowNum, "WOPNO")) + "/" + CStr(.Value(RowNum, "WOPSR")) + ","
            If NewValue = "Y" Then
              wChkSr = InStr(1, gs_OrdZoomSel, wZoomTxt, vbTextCompare)
              If wChkSr = 0 Then gs_OrdZoomSel = gs_OrdZoomSel + wZoomTxt
            Else
                gs_OrdZoomSel = Replace(gs_OrdZoomSel, wZoomTxt, "")
            End If
    End Select
  End With
  '****** Sachin 3.3.0 (26-06-08) - [21. Order Copy From Zoom] ******
End Sub

Private Sub GrdOrdPnd_ColWhen(ByVal RowNum As Integer, ByVal ColName As String, Cancel As Boolean, ErrMsg As String, ByVal Action As MwfCtl.en_fgAction)
  '****** Sachin 3.3.0 (26-06-08) - [21. Order Copy From Zoom] ******
  If Not ColName = "WCPYODSR" Then Cancel = True
  '****** Sachin 3.3.0 (26-06-08) - [21. Order Copy From Zoom] ******
End Sub

Private Sub GrdOrdPnd_RowWhen(ByVal RowNum As Integer)
  If RowNum > 0 Then Call FillGrdOrdWiw(RowNum)
  'Manali 3.2.2
  FraPic3D.Left = wFraPic3DLft: FraPic3D.Top = wFraPic3DTop
  FraPic3D.Height = wFraPic3DHt: FraPic3D.Width = wFraPic3DWdt
  FraPicHsk.Left = wFraPicHskLft: FraPicHsk.Top = wFraPicHskTop
  FraPicHsk.Height = wFraPicHskHt: FraPicHsk.Width = wFraPicHskWdt
  
  '*** Jay 3.2.0 [PicNm]
  wHskPth = GetPictPath("DM", True, GrdOrdPnd.Value(RowNum, "wOpDsgCd"), GrdOrdPnd.Value(RowNum, "wOpSz"), , , , , , , GrdOrdPnd.Value(RowNum, "wOpPicNm"))
  w3DPth = GetPictPath("DM", False, GrdOrdPnd.Value(RowNum, "wOpDsgCd"), GrdOrdPnd.Value(RowNum, "wOpSz"), , , , , , , GrdOrdPnd.Value(RowNum, "wOpPicNm"))
  
  'wHskPth = GetPictPath("DM", True, GrdOrdPnd.Value(RowNum, "wOpDsgCd"), "")
  'w3DPth = GetPictPath("DM", False, GrdOrdPnd.Value(RowNum, "wOpDsgCd"), "")
  '*** Jay 3.2.0 [PicNm]
  ' **** Zubin 211 **** '
  'wHskPth = GetPictPath("DM", GrdOrdPnd.Value(RowNum, "wOpDsgCd"), "", True)
  'w3DPth = GetPictPath("DM", GrdOrdPnd.Value(RowNum, "wOpDsgCd"), "", False)
  ' **** Zubin 211 **** '
  Call ShowPic(wPicHskHt, wPicHskWdt, wHskPth, PicHsk)
  Call ShowPic(wPic3DHt, wPic3DWdt, w3DPth, Pic3D)
End Sub
Private Sub GrdInv_RowWhen(ByVal RowNum As Integer)
  If RowNum > 0 Then Call fillGrdInvBag(RowNum)
End Sub

Private Sub GrdOrdRm_ColWhen(ByVal RowNum As Integer, ByVal ColName As String, Cancel As Boolean, ErrMsg As String, ByVal Action As MwfCtl.en_fgAction)
    Cancel = True
End Sub

Private Sub GrdOrdWIW_ColWhen(ByVal RowNum As Integer, ByVal ColName As String, Cancel As Boolean, ErrMsg As String, ByVal Action As MwfCtl.en_fgAction)
  Cancel = True
End Sub
Private Sub GrdOrdWIW_RowWhen(ByVal RowNum As Integer)
  If RowNum > 0 Then Call FillGrdOrdBag(RowNum)
End Sub
Public Sub fillGrdInv(ByVal RowNum As Integer)
  Dim wCnd As String, wsql As String
  Dim wo_rsInv As MwfLib.MDORowSet, wi_Row As Integer
  
  ' ***** Manali 3.10.0 - 09/03/12 -  Allow PDCM in Zoom - ws_CoCd replace gs_CoCd
  'Creation.7-InCoCd added, InCoCd in where condition is changed
  wsql = "Select IdTc, IdYy, IdChr, IdNo, IdSr, InDt, InExpNo, IdQty, " + _
         "IdIFob/(case when IdQty>0 then IdQty else 1 end) as qPrc, InCoCd " + _
         "From InvDsg join InvHd on InCocd= IdCoCd and InTc= IdTc and InYy= IdYy and InChr= IdChr and " + _
         "InNo= IdNo " + IIF(gs_Partition = ctCurrPrtn, " And InPrtKey=IdPrtKey ", "") + " Where " + IIF(gs_Partition = ctCurrPrtn, " InPrtKey='" + ctCurrPrtn + "' And ", "") + " IdCocd = '" + GrdOrdPnd.Value(RowNum, "WOPCOCD") + "' /*'" + gs_CoCd + "'*/ and IdExpOdTc = '" + GrdOrdPnd.Value(RowNum, "wOpTc") + "' and " + _
         "IdExpOdYy = '" + GrdOrdPnd.Value(RowNum, "wOpYy") + "' and IdExpOdChr = '" + _
         GrdOrdPnd.Value(RowNum, "wOpChr") + "' and IdExpOdNo = " + _
         CStr(GrdOrdPnd.Value(RowNum, "wOpNo")) + " and IdExpOdSr = " + _
         CStr(GrdOrdPnd.Value(RowNum, "wOpSr")) + " Order By IdTc, IdYy, IdChr, IdNo, IdSr "
     
  GrdInv.Rows = 1
     
  Set wo_rsInv = moCn.OpenRes(wsql)
  GrdInv.AllowAdd = True
    
  If (wo_rsInv.EOF Or wo_rsInv.BOF) Then GrdInv.Rows = 1
        
  With GrdInv
    Do While Not (wo_rsInv.EOF Or wo_rsInv.BOF)
      .AddItem
      wi_Row = .Rows - 1
      .Value(wi_Row, "wIdTc") = Trim(wo_rsInv!IdTc)
      .Value(wi_Row, "wIdYy") = Trim(wo_rsInv!IdYy)
      .Value(wi_Row, "wIdChr") = Trim(wo_rsInv!IdChr)
      .Value(wi_Row, "wIdNo") = Trim(wo_rsInv!IdNo)
      .Value(wi_Row, "wIdSr") = Trim(wo_rsInv!IdSr)
      .Value(wi_Row, "wInDt") = Trim(wo_rsInv!InDt)
      .Value(wi_Row, "wInExpNo") = Trim(wo_rsInv!InExpNo)
      .Value(wi_Row, "wIdQty") = Trim(wo_rsInv!IdQty)
      .Value(wi_Row, "wIdIFobPc") = Trim(wo_rsInv!qPrc)
      'Creation.7-InCoCd value added in grid
      .Value(wi_Row, "WINCOCD") = Trim(wo_rsInv!incocd)
      wo_rsInv.MoveNext
    Loop
  End With
  
  Set wo_rsInv = Nothing
       
  If ADC("WPWDYN") <> "Y" Then
    GrdInv.ColProp("wIdIFobPc").Style = fgcinvisible
  Else
    GrdInv.ColProp("wIdIFobPc").Style = fgcReadOnly
  End If
       
  If GrdInv.Rows > 1 Then
    GrdInv.Row = 1
    GrdInv.Col = 0
    Call fillGrdInvBag(1)
  End If
End Sub
Public Sub fillGrdInvBag(ByVal RowNum As Integer)
  Dim wsql As String
  Dim wo_rsInvBag As MwfLib.MDORowSet, wi_Row As Integer
  
  ' ***** Manali 3.10.0 - 09/03/12 -  Allow PDCM in Zoom - ws_CoCd replace gs_CoCd
  'Creation.7- IfCoCd in where condition is changed, given invoice CoCd
  wsql = "SELECT IfFdBYy, IfFdBChr, IfFdBNo, IfQty, IfGrsWt " + _
          "from InvFgd " + _
          "where " + IIF(gs_Partition = ctCurrPrtn, " IfPrtKey='" + ctCurrPrtn + "' And ", "") + " IfCoCd= '" + GrdInv.Value(RowNum, "WINCOCD") + "' /*'" + gs_CoCd + "'*/ and IfTc = '" + GrdInv.Value(RowNum, "wIdTc") + "' and " + _
          "IfYy='" + GrdInv.Value(RowNum, "wIdYy") + "' and " + _
          "IfChr = '" + GrdInv.Value(RowNum, "wIdChr") + "' and " + _
          "IfNo = " & GrdInv.Value(RowNum, "wIdNo") & " and " + _
          "IfSr = " & GrdInv.Value(RowNum, "wIdSr") & " Order By IfFdByy, IfFdBChr, IfFdBNo "
   
  GrdInvBag.Rows = 1
  
  Set wo_rsInvBag = moCn.OpenRes(wsql)
  If (wo_rsInvBag.EOF = True) Then
     GrdInvRm.Rows = 1
  End If
  
  GrdInvBag.AllowAdd = True
  With GrdInvBag
    Do While Not (wo_rsInvBag.EOF Or wo_rsInvBag.BOF)
      .AddItem
      wi_Row = .Rows - 1
      .Value(wi_Row, "wIBYy") = Trim(wo_rsInvBag!IfFdBYy)
      .Value(wi_Row, "wIBChr") = Trim(wo_rsInvBag!IfFdBChr)
      .Value(wi_Row, "wIBNo") = Trim(wo_rsInvBag!IfFdBNo)
      .Value(wi_Row, "wIBPcs") = Trim(wo_rsInvBag!IfQty)
      .Value(wi_Row, "wIBGrsWt") = Trim(wo_rsInvBag!IfGrsWt)
      wo_rsInvBag.MoveNext
    Loop
  End With
  Set wo_rsInvBag = Nothing
  
  If GrdInvBag.Rows > 1 Then
    GrdInvBag.Row = 1
    Call FillGrdInvRm(1)
  End If
  
  'GrdInvBag.AllowAdd = False
  
End Sub
Public Sub FillGrdInvRm(ByVal RowNum As Integer)
  'Incomplete
  Dim wsql As String
  Dim wo_rsInvRm As MwfLib.MDORowSet, wi_Row As Integer
  
  If GrdInv.Row = 0 Then Exit Sub       '*** Jay 3.01 *** [Custom View]
  
  'IrCoCd, IrTc, IrYy, IrChr, IrNo, IrSr, IrSrNo
  ' ***** Manali 3.10.0 - 09/03/12 -  Allow PDCM in Zoom - ws_CoCd replace gs_CoCd
  'Creation.7- IrCoCd in where condition is changed, given invoice CoCd
  wsql = "Select RmCtg, RmSCtg, IrRmCd, IrRmQty, IrRmIWt, RmQw, IrRmIRt, IrRmIVal, IrSetSCd, IrSetIRt, " + _
           "IrSetIVal, IrMainMet from InvRm Join RMMst on RMCd = IrRmCd " + IIF(gs_Partition = ctCurrPrtn, " And RmPrtKey=IrPrtKey ", "") + _
           "where " + IIF(gs_Partition = ctCurrPrtn, " IrPrtKey='" + ctCurrPrtn + "' And ", "") + " IrCoCd= '" + GrdInv.Value(RowNum, "WINCOCD") + "' /*'" + gs_CoCd + "'*/ and IrTc= '" + GrdInv.Value(GrdInv.Row, "wIdTc") + "' and " + _
           "IrYy='" + GrdInv.Value(GrdInv.Row, "wIdYy") + "' and " + _
           "IrChr = '" + GrdInv.Value(GrdInv.Row, "wIdChr") + "' and " + _
           "IrNo = " & GrdInv.Value(GrdInv.Row, "wIdNo") & " and " + _
           "IrSr = " & GrdInv.Value(GrdInv.Row, "wIdSr") & " Order By IrSrNo "
  
  
  Set wo_rsInvRm = moCn.OpenRes(wsql)
  
  GrdInvRm.AllowAdd = True
  GrdInvRm.Rows = 1
  With GrdInvRm
    Do While Not (wo_rsInvRm.EOF Or wo_rsInvRm.BOF)
      .AddItem
      wi_Row = .Rows - 1
      .Value(wi_Row, "wRmCtg") = Trim(wo_rsInvRm!RmCtg)
      .Value(wi_Row, "wRmSCtg") = Trim(wo_rsInvRm!RmSCtg)
      .Value(wi_Row, "wRmCd") = Trim(wo_rsInvRm!IrRmCd)
      .Value(wi_Row, "wRmQty") = wo_rsInvRm!IrRmQty
      .Value(wi_Row, "wRmIWt") = wo_rsInvRm!IrRmIWt
      .Value(wi_Row, "wQw") = wo_rsInvRm!RmQw
      .Value(wi_Row, "wRmIRt") = wo_rsInvRm!IrRmIRt
      .Value(wi_Row, "wRmIVal") = wo_rsInvRm!IrRmIVal
      .Value(wi_Row, "wSetSCd") = Trim(wo_rsInvRm!IrSetSCd)
      .Value(wi_Row, "wSetIRt") = wo_rsInvRm!IrSetIRt
      .Value(wi_Row, "wSetIVal") = wo_rsInvRm!IrSetIVal
      .Value(wi_Row, "wMainMet") = Trim(wo_rsInvRm!IrMainMet)
      wo_rsInvRm.MoveNext
    Loop
  End With
  Set wo_rsInvRm = Nothing
  
  If ADC("WPWDYN") <> "Y" Then
    GrdInvRm.ColProp("wRmIRt").Style = fgcinvisible
    GrdInvRm.ColProp("wRmIVal").Style = fgcinvisible
    GrdInvRm.ColProp("wSetIRt").Style = fgcinvisible
    GrdInvRm.ColProp("wSetIVal").Style = fgcinvisible
  Else
    GrdInvRm.ColProp("wRmIRt").Style = fgcReadOnly
    GrdInvRm.ColProp("wRmIVal").Style = fgcReadOnly
    GrdInvRm.ColProp("wSetIRt").Style = fgcReadOnly
    GrdInvRm.ColProp("wSetIVal").Style = fgcReadOnly
  End If
  
  If GrdInvRm.Rows > 1 Then
    GrdInvRm.Row = 1
  End If
  'GrdInvRm.AllowAdd = False
  
End Sub
Public Function GetCustDsgCnd()
  Dim ws_CustDsgCnd As String
  Dim ws_CustDsgWhere  As String
  ws_CustDsgCnd = GetCustDsgCnd
  ws_CustDsgWhere = ""
  ws_CustDsgCnd = ""
  If ADC("WCDCUSTDMCDFR") <> "" And ADC("WCDCUSTDMCDTO") = "" Then ADC("WCDCUSTDMCDTO") = String(40, "Z")
  If ADC("WCDCUSTDMCDTO") <> "" Or ADC("WCDCUSTDMCDFR") <> "" Then
    ws_CustDsgWhere = " IsNull(CdCustDmCd,'')>='" + ADC("wCdCustDmCdFr") + "' and IsNull(CdCustDmCd,'')<='" + ADC("wCdCustDmCdTo") + "'"
    ws_CustDsgCnd = " Exists (Select 'a' from CustDsg where CdCtg='C' and CdCd=OmCmCd and CdOurDmCd =OdDmCd and " + _
                                      " ((CdOurSfx='' and CdOurDmSz='') or (CdOurSfx=OdSfx and CdOurDmSz=OdDmSz) or (CdOurSfx='' and CdOurDmSz=OdDmSz) or (CdOurSfx=OdSfx and CdOurDmSz='')) and " + ws_CustDsgWhere + " )"
  End If
  GetCustDsgCnd = ws_CustDsgCnd
End Function
Public Sub FillGrdOrdMst()
  '****** Sachin 3.01 - 01-06-07 - Custom Views ******
  Dim wCnd As String, wrepcnd As String, wsql As String
  Dim wo_rsOrdMst As MwfLib.MDORowSet, wi_Row As Integer

  Dim ws_CustDsgCnd As String
  ws_CustDsgCnd = GetCustDsgCnd
  wb_WiwFlag = True
  
  wrepcnd = ADC.RepCond
  If ADC("wYN") = "Y" Then
    wCnd = "Where convert(decimal (14,1), OdOrdQty)<> convert(decimal (14,1), OdExpQty) and OmCoCd>= '" + ws_CoCdFr + "' and OmCoCd<= '" + ws_CoCdTo + "' " + IIF(wrepcnd <> "", " and " + wrepcnd, "")
  ElseIf ADC("wYN") = "N" Then
    wCnd = IIF(wrepcnd <> "", " Where OmCoCd>= '" + ws_CoCdFr + "' and OmCoCd<= '" + ws_CoCdTo + "' and " + wrepcnd, "")
  End If
  
  wsql = "Select OmCoCd, OmTc, OmYy, OmChr, OmNo, Max(OmExpDelDt) As OmExpDelDt,Max(OmDelDt) As OmDelDt, " + _
         "Max(OmCmCd) As OmCmCd, Max(OmPoNo) As OmPoNo, Max(OmPoDt) As OmPoDt, Max(OmPrtCd) as OmPrtCd  " + _
         "From OrdMst " + _
         "Join OrdDsg on OmIdNo=OdOmIdNo And OmCoCd= OdCoCd and OmTc= OdTc and OmYy= OdYy and OmChr= OdChr and OmNo= OdNo " + IIF(gs_Partition = ctCurrPrtn, " And OmPrtKey=OdPrtKey ", "") + _
         "Join DsgMst on DmIdNo=OdDmIdNo And /*DmTcTyp = 'DM' and*/ DmCd = OdDmCd and DmSz = '' " + IIF(gs_Partition = ctCurrPrtn, " and DmPrtKey='" + ctCurrPrtn + "' ", "") + _
         wCnd + _
         IIF(gs_Partition = ctCurrPrtn, IIF(wCnd = "", " where ", " and ") + _
         " OmPrtKey='" + ctCurrPrtn + "' ", "") + _
         IIF(ws_CustDsgCnd <> "", " and " + ws_CustDsgCnd, "") + _
         " Group By OmCoCd, OmTc, OmYy, OmChr, OmNo Order By OmCoCd, OmTc, OmYy, OmChr, OmNo "
  GrdOrdMst.Rows = 1
  Set wo_rsOrdMst = moCn.OpenRes(wsql)
  If wo_rsOrdMst.RecCount <= 2500 Then
    GrdOrdMst.AllowAdd = True
    With GrdOrdMst
      Do While Not (wo_rsOrdMst.EOF Or wo_rsOrdMst.BOF)
        .AddItem
        wi_Row = .Rows - 1
        .Value(wi_Row, "wOmTc") = Trim(wo_rsOrdMst!OmTc)
        .Value(wi_Row, "wOmYy") = Trim(wo_rsOrdMst!OmYy)
        .Value(wi_Row, "wOmChr") = Trim(wo_rsOrdMst!OmChr)
        .Value(wi_Row, "wOmNo") = Trim(wo_rsOrdMst!OmNo)
        .Value(wi_Row, "wOrdNo") = Trim(wo_rsOrdMst!OmCoCd) + "/" + Trim(wo_rsOrdMst!OmTc) + "/" + Trim(wo_rsOrdMst!OmYy) + "/" + Trim(wo_rsOrdMst!OmChr) + "/" + CStr(Trim(wo_rsOrdMst!OmNo)) '*** Jay 3.01 *** [Custom View]
        
         'IG.211 - To be displayed either Exp Del Dt or Prd Del Dt based on scope.
          
        If ADC("wDelDt") = "E" Then
          .Value(wi_Row, "wOmEPDelDt") = Trim(wo_rsOrdMst!OmExpDelDt)
        Else
          .Value(wi_Row, "wOmEPDelDt") = Trim(wo_rsOrdMst!OmDelDt)
        End If
        
        .Value(wi_Row, "wOmCust") = Trim(wo_rsOrdMst!OmCmCd)
        .Value(wi_Row, "wOmPoNo") = Trim(wo_rsOrdMst!OmPoNo)
        .Value(wi_Row, "wOmPoDt") = Trim(wo_rsOrdMst!OmPoDt)
        .Value(wi_Row, "wPrt") = Trim(wo_rsOrdMst!OmPrtCd)    '*** Jay 3.01 *** [Custom View]
        'Creation.7-ImCoCd added to grid
        .Value(wi_Row, "WOMCOCD") = Trim(wo_rsOrdMst!OmCoCd)
        
        wo_rsOrdMst.MoveNext
      Loop
    End With
  Else
    DispMsg "No. of Records are more than 2500, Plz Limit the Scope", etError
  End If
  Set wo_rsOrdMst = Nothing
  wb_WiwFlag = False
  If GrdOrdMst.Rows > 1 Then
    GrdOrdMst.Row = 1
    Call FillGrdOrdPnd(1)
  End If
  mf_OrdMstNo = GrdOrdMst.Row
  '****** Sachin 3.01 - 01-06-07 - Custom Views ******
End Sub

Public Sub FillGrdOrdPnd(ByVal RowNum As Integer)
  Dim wCnd As String, wrepcnd As String
  Dim wsql As String
  Dim wo_rsOrdPnd As MwfLib.MDORowSet, wi_Row As Integer
  Dim wd_OrdQty As Double, wd_PrdQty As Double, wPndYn As String, wOdCpySr As String, wSrStr As Integer
  
  Dim ws_OrdBy As String    ' ***** Manali 3.10.0 - 09/03/12 - Sort on DmCd in GrdOrdPnd
  If ADC("wSortDmCd") = "Y" Then
    ws_OrdBy = " Order By OdTc, OdYy, OdChr, OdNo, OdDmCd, OdSr  "
  Else
    ws_OrdBy = " Order By OdTc, OdYy, OdChr, OdNo, OdSr  "
  End If
  
  
  wb_WiwFlag = True
  wd_OrdQty = 0: wd_PrdQty = 0
  
  wrepcnd = ADC.RepCond
  wCnd = IIF(wrepcnd <> "", " and OmCoCd>= '" + ws_CoCdFr + "' and OmCoCd <= '" + ws_CoCdTo + "' and " + wrepcnd, "")
  If ADC("wYN") = "Y" Then
    wPndYn = " and convert(decimal (14,1), OdOrdQty)<> convert(decimal (14,1), OdExpQty) "
  ElseIf ADC("wYN") = "N" Then
    wPndYn = ""
  End If
  Dim ws_CustDsgCnd As String
  ws_CustDsgCnd = GetCustDsgCnd
         
  wsql = "Select OdCoCd, OdTc, OdYy, OdChr, OdNo, OdSr, OdDmCd, OdSfx, OdDmSz, OdOrdEnt, OdPrdEnt, OdOrdQty, OdPrdQty, " + _
         "OdExpQty, (OdOrdQty-OdExpQty) as qBalQty, OdPrtCd, OdExpDelDt, OdDelDt, OdKt, OdDmCol, OdPicNm, OdSalPrc, " + _
         "(OdSalPrc * (OdOrdQty-OdExpQty)) as qTotVal, OdSubCust, OdPoNo," + _
         "IsNull((Select CdCustDmCd from CustDsg Where CdCtg = 'C' and CdCd = OmCmCd and " + _
         " CdOurDmCd =OdDmCd and CdOurSfx=OdSfx and CdOurDmSz=OdDmSz), " + _
         "IsNull((Select CdCustDmCd from CustDsg Where CdCtg = 'C' and CdCd = OmCmCd and " + _
         " CdOurDmCd =OdDmCd and CdOurSfx=OdSfx and CdOurDmSz = '' ), " + _
         "IsNull((Select CdCustDmCd from CustDsg Where CdCtg = 'C' and CdCd = OmCmCd and " + _
         " CdOurDmCd =OdDmCd and CdOurSfx='' and CdOurDmSz=OdDmSz), " + _
         "IsNull((Select CdCustDmCd from CustDsg Where CdCtg = 'C' and CdCd = OmCmCd and " + _
         " CdOurDmCd =OdDmCd and CdOurSfx='' and CdOurDmSz=''), '')))) as CdCustDmCd " + _
         "From OrdMst join OrdDsg on OmIdNo=OdOmIdNo And OmCoCd= OdCoCd and OmTc= OdTc and OmYy= OdYy and OmChr= OdChr and OmNo= OdNo " + IIF(gs_Partition = ctCurrPrtn, " And OmPrtKey=OdPrtKey ", "") + _
         "              join DsgMst on DmIdNo=OdDmIdNo And /*DmTcTyp = 'DM' and*/ DmCd = OdDmCd and DmSz = '' " + IIF(gs_Partition = ctCurrPrtn, " And DmPrtKey=OdPrtKey ", "") + _
         "Left Outer Join CustDsg on CdCtg='C' and CdCd=OmCmCd and CdOurDmCd =OdDmCd and CdOurSfx='' and CdOurDmSz='' " + _
         "Where " + IIF(gs_Partition = ctCurrPrtn, " OdPrtKey='" + ctCurrPrtn + "' And ", "") + " OdCoCd='" + GrdOrdMst.Value(RowNum, "WOMCOCD") + "' /*'" + gs_CoCd + "'*/ And OdTc='" + GrdOrdMst.Value(RowNum, "WOMTC") + "' " + _
         "  And OdYy='" + GrdOrdMst.Value(RowNum, "WOMYY") + "' And OdChr='" + GrdOrdMst.Value(RowNum, "WOMCHR") + "' And OdNo=" + CStr(GrdOrdMst.Value(RowNum, "WOMNO")) + _
         wCnd + wPndYn + IIF(ws_CustDsgCnd <> "", " and " + ws_CustDsgCnd, "") + ws_OrdBy
  
  GrdOrdPnd.Rows = 1
  Set wo_rsOrdPnd = moCn.OpenRes(wsql)
    GrdOrdPnd.AllowAdd = True
    With GrdOrdPnd
      Do While Not (wo_rsOrdPnd.EOF Or wo_rsOrdPnd.BOF)
        'if CdCustDsgFr (from of CustDsg filter) is filled and To is blank To is automatically filled by code with ZZZZ
        If (ADC("wCdCustDmCdFr") <> "" Or ADC("wCdCustDmCdTo") <> "") And Not (Trim(wo_rsOrdPnd!CdCustDmCd) >= ADC("wCdCustDmCdFr") And Trim(wo_rsOrdPnd!CdCustDmCd) <= ADC("wCdCustDmCdTo")) Then
          GoTo ContinueLoop
        End If
        .AddItem
        wi_Row = .Rows - 1
        .Value(wi_Row, "wOpTc") = Trim(wo_rsOrdPnd!OdTc)
        .Value(wi_Row, "wOpYy") = Trim(wo_rsOrdPnd!OdYy)
        .Value(wi_Row, "wOpChr") = Trim(wo_rsOrdPnd!OdChr)
        .Value(wi_Row, "wOpNo") = Trim(wo_rsOrdPnd!OdNo)
        .Value(wi_Row, "wOpSr") = Trim(wo_rsOrdPnd!OdSr)
        .Value(wi_Row, "wOpExpDt") = Trim(wo_rsOrdPnd!OdExpDelDt)
        If ADC("wDelDt") = "E" Then
          .Value(wi_Row, "wOpExpDt") = Trim(wo_rsOrdPnd!OdExpDelDt)
        Else
          .Value(wi_Row, "wOpExpDt") = Trim(wo_rsOrdPnd!OdDelDt)
        End If
        
        .Value(wi_Row, "wOpDsgCd") = Trim(wo_rsOrdPnd!OdDmCd)
        .Value(wi_Row, "wOpBalPcs") = Trim(wo_rsOrdPnd!qBalQty)
        .Value(wi_Row, "wOpKarat") = Trim(wo_rsOrdPnd!OdKt)
        .Value(wi_Row, "wOpValPpc") = Trim(wo_rsOrdPnd!OdSalPrc)
        .Value(wi_Row, "wOpTotVal") = Trim(wo_rsOrdPnd!qTotVal)
        .Value(wi_Row, "wOpOrdEnt") = Trim(wo_rsOrdPnd!OdOrdEnt)
        .Value(wi_Row, "wOpPrdEnt") = Trim(wo_rsOrdPnd!OdPrdEnt)
        .Value(wi_Row, "wOpSfx") = Trim(wo_rsOrdPnd!OdSfx)
        .Value(wi_Row, "wOpSz") = Trim(wo_rsOrdPnd!OdDmSz)
        .Value(wi_Row, "wOpDmCol") = Trim(wo_rsOrdPnd!OdDmCol)
        .Value(wi_Row, "wOpPrtCd") = Trim(wo_rsOrdPnd!OdPrtCd)
        .Value(wi_Row, "wOpOrdQty") = Trim(wo_rsOrdPnd!OdOrdQty)
        .Value(wi_Row, "wOpPrdQty") = Trim(wo_rsOrdPnd!OdPrdQty)
        .Value(wi_Row, "wOpBalExpQty") = Trim(wo_rsOrdPnd!qBalQty)
        .Value(wi_Row, "wOpPicNm") = Trim(wo_rsOrdPnd!OdPicNm)
        If Trim(wo_rsOrdPnd!OdSubCust) = "" And Trim(wo_rsOrdPnd!OdPONo) = "" Then
          .Value(wi_Row, "wOpSubPoNo") = ""
        Else
          .Value(wi_Row, "wOpSubPoNo") = Trim(wo_rsOrdPnd!OdSubCust) + " / " + Trim(wo_rsOrdPnd!OdPONo)
        End If
        wOdCpySr = .Value(wi_Row, "wOpTc") + "/" + .Value(wi_Row, "wOpYy") + "/" + .Value(wi_Row, "wOpChr") + "/" + CStr(.Value(wi_Row, "wOpNo")) + "/" + CStr(.Value(wi_Row, "wOpSr"))
        wSrStr = InStr(1, gs_OrdZoomSel, wOdCpySr, vbTextCompare)
        .Value(wi_Row, "WCPYODSR") = IIF(gs_OrdZoomSel = "", "", IIF(wSrStr > 0, "Y", ""))
        .Value(wi_Row, "WOPCOCD") = Trim(wo_rsOrdPnd!OdCoCd)
        .Value(wi_Row, "wOpCustDmCd") = Trim(wo_rsOrdPnd!CdCustDmCd)
ContinueLoop:
        wo_rsOrdPnd.MoveNext
      Loop
    End With
  Set wo_rsOrdPnd = Nothing
  wb_WiwFlag = False
  If ADC("WPWDYN") <> "Y" Then
    GrdOrdPnd.ColProp("wOpValPpc").Style = fgcinvisible:    GrdOrdPnd.ColProp("wOpTotVal").Style = fgcinvisible
  Else
    GrdOrdPnd.ColProp("wOpValPpc").Style = fgcReadOnly:     GrdOrdPnd.ColProp("wOpTotVal").Style = fgcReadOnly
  End If
  If GrdOrdPnd.Rows > 1 Then
    GrdOrdPnd.Row = 1
    Call FillGrdOrdWiw(1)
  End If
  'mf_OrdNo = GrdOrdMst.Row   ' ***** Manali 3.5.0 -  Commented this line as sometimes design details were not getting refreshed
End Sub
 Public Sub FillGrdOrdWiw(ByVal RowNum As Integer)
  Dim wsql As String, ws_Seq As String, ws_PrdSeq As String
  Dim wo_rsOrdWiw As MwfLib.MDORowSet, wi_Row As Integer
  Dim wCnd As String
  If wb_WiwFlag = False Then
    wb_BagFlag = True
    ws_Seq = moCn.GetFldVal("Select hDmPrdSeq from Head where hCd = '" + ctSelfCmCd + "'")
    ws_PrdSeq = moCn.GetFldVal("Select PDesc225 from Param where PTyp= 'PRDSEQ' and " + _
                "PMCd= '" + ws_Seq + "'")
    If ADC("wYN") = "Y" Then
      wCnd = " and BCls= 'N' "
    Else
      wCnd = " "
    End If
    ' 2.11 (Jen) Introduced join for Flute Bag option
    ' ***** Manali 3.9.0 - Melted bag skipped
    ' ***** Manali 3.10.0 - 09/03/12 -  Allow PDCM in Zoom - ws_CoCd replace gs_CoCd
    'Creation.7-BCoCd in where condition is changed, given OrdDSg Cocd
    wsql = "SELECT BLoc, count(BNo) as BNo, sum(BQty) as BQty, " + _
           "str(case charindex(BLoc, '" + ws_PrdSeq + "') when 0 then 99 else " + _
           "charindex(BLoc, '" + ws_PrdSeq + "' ) end)+ BLoc as qLoc " + _
           "from Bag " + ms_BagJoin + _
           " where " + IIF(gs_Partition = ctCurrPrtn, " BPrtKey='" + ctCurrPrtn + "' and ", "") + " BCoCd = '" + GrdOrdPnd.Value(RowNum, "WOPCOCD") + "' /*'" + gs_CoCd + "'*/ and BOdTc= '" + GrdOrdPnd.Value(RowNum, "wOpTc") + "' and " + _
           "BOdYy='" + GrdOrdPnd.Value(RowNum, "wOpYy") + "' and " + _
           "BOdChr = '" + GrdOrdPnd.Value(RowNum, "wOpChr") + "' and " + _
           "BOdNo = " & GrdOrdPnd.Value(RowNum, "wOpNo") & "and " + _
           "BOdSr = " & GrdOrdPnd.Value(RowNum, "wOpSr") & wCnd + " and (IsNull((Select TdTc from Txnd where BLstYyKey= TdYyKey and  TdPrtKey=BPrtKey and TdBIdNo=BIdno and TdCoCd=BCoCd and  TdByy=Byy and TdBChr=Bchr and TdBNo=BNo and TdSrNo=0),'') <> (Select PValue from Param where PTyp='TCTyp' and PMCd='MLT')) " + _
           " Group By BLoc " + _
           " Order By Case When BLoc='PFG' Then '' Else BLoc End"
    GrdOrdWIW.Rows = 1
    Set wo_rsOrdWiw = moCn.OpenRes(wsql)
    If (wo_rsOrdWiw.EOF = True) Then
      GrdOrdBag.Rows = 1
      GrdOrdBagRm.Rows = 1      ' ***** Manali 3.8.0 - Bag Rm Details
    End If
    GrdOrdWIW.AllowAdd = True
    With GrdOrdWIW
      Do While Not (wo_rsOrdWiw.EOF Or wo_rsOrdWiw.BOF)
        .AddItem
        wi_Row = .Rows - 1
        .Value(wi_Row, "wOwLoc") = Trim(wo_rsOrdWiw!BLoc)
        .Value(wi_Row, "WOWTOTBAGS") = Trim(wo_rsOrdWiw!BNo)
        .Value(wi_Row, "WOWPCS") = Trim(wo_rsOrdWiw!BQty)
        wo_rsOrdWiw.MoveNext
      Loop
    End With
    Set wo_rsOrdWiw = Nothing
    wb_BagFlag = False
    If GrdOrdWIW.Rows > 1 Then
        GrdOrdWIW.Row = 1
        Call FillGrdOrdBag(1)
    End If
    'GrdOrdWIW.AllowAdd = False
  End If
End Sub
Public Sub FillGrdOrdBag(ByVal RowNum As Integer)
  Dim wsql As String
  Dim wo_rsOrdBag As MwfLib.MDORowSet, wi_Row As Integer
  Dim wCnd As String
  
  If wb_BagFlag = False Then
    wb_BagRmFlag = True ' ***** Manali 3.8.0 - Bag Rm Details
    If ADC("wYN") = "Y" Then
      wCnd = " and BCls= 'N' "
    Else
      wCnd = " "
    End If
    
    ' 2.11 (Jen) Introduced join for Flute Bag option
    '*** Jay 2.13(CT) *** (Part Qty Added)
    ' **** Manali 3.5.0 - 08/11/08 - bag Recvd date added in OrdBag Grid
    ' ***** Manali 3.9.0 - Melted bag skipped
    ' ***** Manali - 3.10.0 - 09/03/12 - Bag Wrk added
    ' ***** Manali 3.10.0 - 09/03/12 -  Allow PDCM in Zoom - ws_CoCd replace gs_CoCd
    'Creation.7-BCoCd is added, BCoCd in where condition is changed, given OrdDSg Cocd
    wsql = " Select BCoCd, BYy, BChr, BNo, (Case When BRjQty<>0 Then BRjQty Else BQty End) as BQty, " + _
           "       (Case When BRjPtQty<>0 Then BRjPtQty Else BPtQty End) as BPtQty, BGrWt, BRecvDt, BWrk, BFgSubLoc, " + _
           "       vPDesc as qWrkName, (Case When BRjQty<>0 Then 'Y' Else 'N' End) as qRejYN from Bag " + ms_BagJoin + _
           "    Join vParam On vPCoCd=BCoCd And vPTyp='WORK' And vPMCd=BWrk And vPSCd='' " + _
           " Where " + IIF(gs_Partition = ctCurrPrtn, " BPrtKey='" + ctCurrPrtn + "' And ", "") + " BCoCd = '" + GrdOrdPnd.Value(GrdOrdPnd.Row, "WOPCOCD") + "' /*'" + gs_CoCd + "'*/ and BOdTc= '" + GrdOrdPnd.Value(GrdOrdPnd.Row, "wOpTc") + "' and " + _
           " BOdYy='" + GrdOrdPnd.Value(GrdOrdPnd.Row, "wOpYy") + "' and " + _
           " BOdChr = '" + GrdOrdPnd.Value(GrdOrdPnd.Row, "wOpChr") + "' and " + _
           " BOdNo = " & GrdOrdPnd.Value(GrdOrdPnd.Row, "wOpNo") & " and " + _
           " BOdSr = " & GrdOrdPnd.Value(GrdOrdPnd.Row, "wOpSr") & " and BLoc = '" + _
           GrdOrdWIW.Value(RowNum, "wOwLoc") + "' " + wCnd + " and (IsNull((Select TdTc from Txnd where BLstYyKey= TdYyKey and  TdPrtKey=BPrtKey and TdBIdNo=BIdno and TdCoCd=BCoCd and  TdByy=Byy and TdBChr=Bchr and TdBNo=BNo and TdSrNo=0),'') <> (Select PValue from Param where PTyp='TCTyp' and PMCd='MLT')) "
    
    GrdOrdBag.Rows = 1
    Set wo_rsOrdBag = moCn.OpenRes(wsql)
    GrdOrdBag.AllowAdd = True
    With GrdOrdBag
      Do While Not (wo_rsOrdBag.EOF Or wo_rsOrdBag.BOF)
        .AddItem
        wi_Row = .Rows - 1
        .Value(wi_Row, "wObYY") = Trim(wo_rsOrdBag!BYy)
        .Value(wi_Row, "wObChr") = Trim(wo_rsOrdBag!BChr)
        .Value(wi_Row, "WObNo") = Trim(wo_rsOrdBag!BNo)
        .Value(wi_Row, "WObPcs") = Trim(wo_rsOrdBag!BQty)
        .Value(wi_Row, "WObPtPcs") = Trim(wo_rsOrdBag!BPtQty)     '*** Jay 2.13(CT) ***
        .Value(wi_Row, "wObGrsWt") = Trim(wo_rsOrdBag!BGrWt)
        .Value(wi_Row, "wObRecvDt") = Trim(wo_rsOrdBag!BRecvDt)   ' ***** Manali 3.5.0 - 08/11/08 - bag Recvd date added in OrdBag Grid
        .Value(wi_Row, "wObWrk") = Trim(wo_rsOrdBag!BWrk)         ' ***** Manali - 3.10.0 - 09/03/12 - Bag Wrk
        .Value(wi_Row, "wObFgSubLoc") = Trim(wo_rsOrdBag!BFgSubLoc)
        .Value(wi_Row, "wObWrkName") = Trim(wo_rsOrdBag!qWrkName) ' ***** Sachin 3.11.0
        .Value(wi_Row, "wObRejYN") = Trim(wo_rsOrdBag!qRejYN)     ' ***** Sachin 3.11.0
        'Creation.7-BCoCd is added to grid
        .Value(wi_Row, "wObCoCd") = Trim(wo_rsOrdBag!BCoCd)
        wo_rsOrdBag.MoveNext
      Loop
    End With
    Set wo_rsOrdBag = Nothing
    
    ' ***** Manali 3.8.0
    wb_BagRmFlag = False
    If GrdOrdBag.Rows > 1 Then
        GrdOrdBag.Row = 1
        Call FillGrdOrdBagRm(1)
    End If
    ' ***** Manali 3.8.0
  End If
  'GrdOrdBag.AllowAdd = False
End Sub
' ***** Manali 3.8.0 - Bag Rm Details
Public Sub FillGrdOrdBagRm(ByVal RowNum As Integer)
  Dim wsql As String
  Dim wo_rsOrdBagRm As MwfLib.MDORowSet, wi_Row As Integer
  Dim wCnd As String
  
  If wb_BagRmFlag = False Then
    ' ***** Manali 3.10.0 - 09/03/12 -  Allow PDCM in Zoom - ws_CoCd replace gs_CoCd
    'Creation.7-TdCoCd in where condition is changed, given OrdBag Cocd
    wsql = "Select max(TdRmCtg) as TdRmCtg, max(TdRmSCtg) as TdRmSCtg, TdRmCd , " + _
            "sum((case (case when TdFrRmLoc='W' then TdFrRmDc " + _
            "when TdToRmLoc='W' then TdToRmDc else '' end) when 'C' then -1 " + _
            "when 'D' then +1 else 0 end)* TdRmQty) as TdRmQty, " + _
            "sum((case (case when TdFrRmLoc='W' then TdFrRmDc when TdToRmLoc='W' then " + _
            "TdToRmDc end) when 'C' then -1 when 'D' then +1 else 0 end)* " + _
            "TdRmWt) as TdRmWt From Txnd " + _
            "Where TdSrNo > 0 and TdRmCd <> '' " + IIF(gs_Partition = ctCurrPrtn, " And TdPrtKey='" + ctCurrPrtn + "' ", "") + _
            " and TdCoCd= '" + GrdOrdBag.Value(GrdOrdBag.Row, "wObCoCd") + "' /*'" + gs_CoCd + "'*/ and " + _
            "TdBYy= '" + GrdOrdBag.Value(GrdOrdBag.Row, "wObYY") + "' and " + _
            "TdBChr= '" + GrdOrdBag.Value(GrdOrdBag.Row, "wObChr") + "' and " + _
            "TdBNo= " + CStr(GrdOrdBag.Value(GrdOrdBag.Row, "wObNo")) + " " + _
            "group By TdBYy, TdBChr, TdBNo, TdRmCd  " + _
            "Having " + _
            "sum((case (case when TdFrRmLoc='W' then TdFrRmDc " + _
            "when TdToRmLoc='W' then TdToRmDc else '' end) when 'C' then -1 " + _
            "when 'D' then +1 else 0 end)* TdRmQty)> 0 or " + _
            "sum((case (case when TdFrRmLoc='W' then TdFrRmDc when TdToRmLoc='W' then " + _
            "TdToRmDc end) when 'C' then -1 when 'D' then +1 else 0 end)* " + _
            "TdRmWt) > 0 "
                
    GrdOrdBagRm.Rows = 1
    Set wo_rsOrdBagRm = moCn.OpenRes(wsql)
    GrdOrdBagRm.AllowAdd = True
    With GrdOrdBagRm
      Do While Not (wo_rsOrdBagRm.EOF Or wo_rsOrdBagRm.BOF)
        .AddItem
        wi_Row = .Rows - 1
        .Value(wi_Row, "wObrRmCtg") = Trim(wo_rsOrdBagRm!TdRmCtg)
        .Value(wi_Row, "wObrRmSCtg") = Trim(wo_rsOrdBagRm!TdRmSCtg)
        .Value(wi_Row, "wObrRmCd") = Trim(wo_rsOrdBagRm!TdRmCd)
        .Value(wi_Row, "WObrRmSz") = ""
        .Value(wi_Row, "WObrRmQty") = Trim(wo_rsOrdBagRm!TdRmQty)
        .Value(wi_Row, "WObrRmWt") = Trim(wo_rsOrdBagRm!TdRmWt)
        wo_rsOrdBagRm.MoveNext
      Loop
    End With
    Set wo_rsOrdBagRm = Nothing
    GrdOrdBagRm.AllowAdd = False
  End If
End Sub
' ***** Manali 3.8.0 - Bag Rm Details


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
Private Sub ATXTDisp_Validate(Index As Integer, Cancel As Boolean)
  Cancel = atxtDisp(Index).Validate
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If UnloadMode <> vbFormCode Then Cancel = True: Exit Sub
    Set moCn = Nothing    '*** (09/08/05)    '*** (Bef speed) bef 24/11/06
End Sub
Private Sub GrdOrdPnd_Validate(Cancel As Boolean)
  Cancel = GrdOrdPnd.Validate
End Sub
Private Sub GrdOrdBag_Validate(Cancel As Boolean)
  Cancel = GrdOrdBag.Validate
End Sub
Private Sub GrdOrdWIW_Validate(Cancel As Boolean)
  Cancel = GrdOrdWIW.Validate
End Sub
Private Sub GrdInv_Validate(Cancel As Boolean)
  Cancel = GrdInv.Validate
End Sub
Private Sub GrdInvBag_Validate(Cancel As Boolean)
  Cancel = GrdInvBag.Validate
End Sub
Private Sub GrdInvRm_Validate(Cancel As Boolean)
  Cancel = GrdInvRm.Validate
End Sub
Private Sub FillDsgDets(ByVal RowNum As Integer)
  '****** Sachin 3.01 - 01-06-07 - Custom Views ******
  Dim wsql As String
  Dim wo_rsDsgDets As MwfLib.MDORowSet, wi_Row As Integer
  Dim wCnd As String
    ' 2.11 (Jen) Introduced join for Flute Bag option
    '***     Jay 2.13(CT) *** (Part Qty Added)
    ' ***** Manali 3.10.0 - 09/03/12 -  Allow PDCM in Zoom - ws_CoCd replace gs_CoCd
  'Creation.7- OdCoCd in where condition is changed, given OrdDSg Cocd
  wsql = "Select * from OrdDsg " + _
     " Where " + IIF(gs_Partition = ctCurrPrtn, " OdPrtKey='" + ctCurrPrtn + "' And ", "") + " OdCoCd = '" + GrdOrdPnd.Value(RowNum, "WOPCOCD") + "' /*'" + gs_CoCd + "'*/ and OdTc= '" + GrdOrdPnd.Value(RowNum, "wOpTc") + "' and " + _
     "OdYy='" + GrdOrdPnd.Value(RowNum, "wOpYy") + "' and OdChr = '" + GrdOrdPnd.Value(RowNum, "wOpChr") + "' and " + _
     "OdNo = " & GrdOrdPnd.Value(RowNum, "wOpNo") & " and OdSr = " & GrdOrdPnd.Value(RowNum, "wOpSr")
  Set wo_rsDsgDets = moCn.OpenRes(wsql)
  If Not (wo_rsDsgDets.EOF Or wo_rsDsgDets.BOF) Then
      ADC("wOdPrdSeq") = Trim(wo_rsDsgDets!OdPrdSeq)
      ADC("wOdOrdQty") = GrdOrdPnd.Value(RowNum, "wOpOrdQty")
      ADC("wOdExpQty") = Trim(wo_rsDsgDets!OdExpQty)            '*** Jay 3.01 *** [Custom Views]
      ADC("wOdPrdEnt") = Trim(wo_rsDsgDets!OdPrdEnt)
      ADC("wOdPrdQty") = Trim(wo_rsDsgDets!OdPrdQty)
      ADC("wOdFgQty") = Trim(wo_rsDsgDets!OdFGQty)              '*** Jay 3.01 *** [Custom Views]
      ADC("wOdCstPrc") = Trim(wo_rsDsgDets!OdCstPrc)
      ADC("wOdDmPrdInst") = Trim(wo_rsDsgDets!OdDmPrdInst)
      ADC("wOdCmPrdInst") = Trim(wo_rsDsgDets!OdCmPrdInst)
      ADC("wOdCmStmpInst") = Trim(wo_rsDsgDets!OdCmStmpInst)
      ADC("wOdSzInst") = Trim(wo_rsDsgDets!OdSzInst)
      ADC("wOdHld") = Trim(wo_rsDsgDets!OdHld)
      ADC("wOdHldDesc") = Trim(wo_rsDsgDets!OdHldDesc)
      ADC("wOdVaCtg") = Trim(wo_rsDsgDets!OdVaCtg)
      ADC("wOdMulBy") = Trim(wo_rsDsgDets!OdMulBy)
      ADC("wOdFixPrc") = Trim(wo_rsDsgDets!ODFixPrc)      ' ****** Manali 3.03
      ADC("wOdGldAs") = Trim(wo_rsDsgDets!OdGldAs)
      ADC("wOdLabAs") = Trim(wo_rsDsgDets!OdLabAs)
      ADC("wOdIWtEqOrd") = Trim(wo_rsDsgDets!OdIwtEqOrd)
      ADC("wOdIWtFrOrd") = Trim(wo_rsDsgDets!OdIwtFrOrd)
      ADC("wOdIGldRtEqOrd") = Trim(wo_rsDsgDets!OdIGldRtEqOrd)
      ' ****** Manali 3.03
      ADC("wOdChgPtrOnStwYn") = Trim(wo_rsDsgDets!OdChgPtrOnStwYN)
      ADC("wOdILabWtFrOrd") = Trim(wo_rsDsgDets!OdILabWtFrOrd)
      ' ****** Manali 3.03
      ADC("wOdBagPcs") = Trim(wo_rsDsgDets!OdBagPcs)
      ADC("wOdDmCol") = Trim(wo_rsDsgDets!OdDmCol)
      ADC("wOdDmCtg") = moCn.GetFldVal("Select DmCtg from DsgMst where " + _
                                    "DmCd= '" + GrdOrdPnd.Value(RowNum, "wOpDsgCd") + "' and " + _
                                    "DmTcTyp= 'DM' and DmSz= '' " + IIF(gs_Partition = ctCurrPrtn, " And DmPrtKey='" + ctCurrPrtn + "' ", ""))
      ADC("wOdSalRem") = Trim(wo_rsDsgDets!OdSalRem)
      ADC("wOdWH") = Trim(wo_rsDsgDets!OdWH)
      ADC("wOdLine") = Trim(wo_rsDsgDets!OdLine)
      ADC("wOdWDiaAvlblDt") = Trim(wo_rsDsgDets!OdWDiaAvlblDt)
      ADC("wOdHDiaAvlblDt") = Trim(wo_rsDsgDets!OdHDiaAvlblDt)
      ADC("wOdOmCtCd") = Trim(wo_rsDsgDets!OdOmCtCd)
      ADC("wOdSubRem") = Trim(wo_rsDsgDets!OdSubRem)
      ADC("wOdChkTol") = Trim(wo_rsDsgDets!OdChkTol)
      ADC("wOdGmChk") = Trim(wo_rsDsgDets!OdGmChk)
      ADC("wOdMinWt") = Trim(wo_rsDsgDets!OdMinWt)
      ADC("wOdMaxWt") = Trim(wo_rsDsgDets!OdMaxWt)
  '            ADC("wOdGrsWt") = moCn.GetFldVal("Select cast(IsNull(sum(OrPrdWt/ (case when OrRmCtg in ('C', 'D') then 5 else 1 end)), 0) as Decimal(8, 3)) " + _
  '                                   " from OrdRm " + _
  '                                   " Join OrdDsg On OdCoCd=OrCoCd And OdTc=OrTc And OdYy=OrYy And OdChr=OrChr And OdNo=OrNo And OdSr=OrSr" + _
  '                                   " Where OdCoCd = '" + gs_CoCd + "' and OdTc= '" + GrdOrdPnd.Value(RowNum, "wOpTc") + "' and " + _
  '                                   "    OdYy='" + GrdOrdPnd.Value(RowNum, "wOpYy") + "' and " + _
  '                                   "    OdChr = '" + GrdOrdPnd.Value(RowNum, "wOpChr") + "' and " + _
  '                                   "    OdNo = " & GrdOrdPnd.Value(RowNum, "wOpNo") & " and " + _
  '                                   "    OdSr = " & GrdOrdPnd.Value(RowNum, "wOpSr") & _
  '                                   "        And ','+ (Select PDesc225 From Param where PTyp= 'CPYRT' " + _
  '                                   "           and PmCd= '" + wo_rsDsgDets!OdGmChk + "')+ ',' like '%,'+ OrRmCtg+ ',%' ")
      ADC("wOdGrMet") = Trim(wo_rsDsgDets!OdGrMet)
      ADC("wOdGrDia") = Trim(wo_rsDsgDets!OdGrDia)
      ADC("wOdCrmFixPrcYN") = Trim(wo_rsDsgDets!OdCrmFixPrcYN)
      ADC("wOdPicNm") = Trim(wo_rsDsgDets!OdPicNm)    '*** Jay 3.2.0 [PicNm]
      ' ****** Manali 3.5.0 - 17/06/09 - Shifted to grid
      '''ADC("wOdPoNo") = Trim(wo_rsDsgDets!OdPoNo)      '*** Jay 3.2.0 [OdPoNo]
      ' ****** Manali 3.5.0 - 17/06/09 - Shifted to grid
      ' ***** Manali 3.5.0 - New fields
      ADC("wOdMinDiaTolWt") = Trim(wo_rsDsgDets!OdMinDiaTolWt)
      ADC("wOdMaxDiaTolWt") = Trim(wo_rsDsgDets!OdMaxDiaTolWt)
      ADC("wOdInsDt") = Trim(wo_rsDsgDets!InsDt)
      ' ***** Manali 3.5.0 - New fields
  End If
  Set wo_rsDsgDets = Nothing
'****** Sachin 3.01 - 01-06-07 - Custom Views ******
End Sub
Private Sub FillOrdRm(ByVal RowNum As Integer)
'****** Sachin 3.01 - 01-06-07 - Custom Views ******
  Dim wsql As String
  Dim wo_rsOrdRm As MwfLib.MDORowSet, wi_Row As Integer
  Dim wCnd As String
    ' ***** Manali 3.10.0 - 09/03/12 -  Allow PDCM in Zoom - ws_CoCd replace gs_CoCd
    'Creation.7- OrCoCd in where condition is changed, given OrdDSg Cocd
    wsql = "Select * from OrdRm " + _
      " Where " + IIF(gs_Partition = ctCurrPrtn, " OrPrtKey='" + ctCurrPrtn + "' And ", "") + " OrCoCd = '" + GrdOrdPnd.Value(RowNum, "WOPCOCD") + "' /*'" + gs_CoCd + "'*/ and OrTc= '" + GrdOrdPnd.Value(RowNum, "wOpTc") + "' and " + _
      "OrYy='" + GrdOrdPnd.Value(RowNum, "wOpYy") + "' and OrChr = '" + GrdOrdPnd.Value(RowNum, "wOpChr") + "' and " + _
      "OrNo = " & GrdOrdPnd.Value(RowNum, "wOpNo") & " and OrSr = " & GrdOrdPnd.Value(RowNum, "wOpSr")
    GrdOrdRm.Rows = 1
    Set wo_rsOrdRm = moCn.OpenRes(wsql)
    With GrdOrdRm
      .AllowAdd = True
      Do While Not (wo_rsOrdRm.EOF Or wo_rsOrdRm.BOF)
        .AddItem
        wi_Row = .Rows - 1
        .Value(wi_Row, "OrSrNo") = Trim(wo_rsOrdRm!OrSrNo)
        .Value(wi_Row, "OrRmCtg") = Trim(wo_rsOrdRm!OrRmCtg)
        .Value(wi_Row, "OrRmSCtg") = Trim(wo_rsOrdRm!OrRmSCtg)
        .Value(wi_Row, "OrRmCd") = Trim(wo_rsOrdRm!OrRmCd)
        .Value(wi_Row, "OrLn1") = Trim(wo_rsOrdRm!OrLn1)
        .Value(wi_Row, "OrLn2") = Trim(wo_rsOrdRm!OrLn2)
        .Value(wi_Row, "OrLn3") = Trim(wo_rsOrdRm!OrLn3)
        .Value(wi_Row, "OrQty") = Trim(wo_rsOrdRm!OrQty)
        .Value(wi_Row, "OrRmPtr") = Trim(wo_rsOrdRm!OrRmPtr)
        .Value(wi_Row, "OrWt") = Trim(wo_rsOrdRm!OrWt)
        .Value(wi_Row, "OrLMESal") = Trim(wo_rsOrdRm!OrLMESal)
        .Value(wi_Row, "OrSalRt") = Trim(wo_rsOrdRm!OrSalRt)
        .Value(wi_Row, "wQWSal") = moCn.GetFldVal("Select RmQw From RmMst Where " + IIF(gs_Partition = ctCurrPrtn, " RmPrtKey='" + ctCurrPrtn + "' and ", "") + " RmCd='" + Trim(wo_rsOrdRm!OrRmCd) + "'")
        .Value(wi_Row, "OrSalVal") = Trim(wo_rsOrdRm!OrSalVal)
        .Value(wi_Row, "OrPrdQty") = Trim(wo_rsOrdRm!OrPrdQty)
        .Value(wi_Row, "OrPrdWt") = Trim(wo_rsOrdRm!OrPrdWt)
        .Value(wi_Row, "OrCstRt") = Trim(wo_rsOrdRm!OrCstRt)
        .Value(wi_Row, "OrCstVal") = Trim(wo_rsOrdRm!OrCstVal)
        .Value(wi_Row, "OrSetSCd") = Trim(wo_rsOrdRm!OrSetSCd)
        .Value(wi_Row, "OrSetSalRt") = Trim(wo_rsOrdRm!OrSetSalRt)
        .Value(wi_Row, "OrSetSalVal") = Trim(wo_rsOrdRm!OrSetSalVal)
        .Value(wi_Row, "OrSetCstRt") = Trim(wo_rsOrdRm!OrSetCstRt)
        .Value(wi_Row, "OrSetCstVal") = Trim(wo_rsOrdRm!OrSetCstVal)
        .Value(wi_Row, "OrAlyCd") = Trim(wo_rsOrdRm!OrAlyCd)
        .Value(wi_Row, "OrAlySalRt") = Trim(wo_rsOrdRm!OrAlySalRt)
        .Value(wi_Row, "OrAlyCstRt") = Trim(wo_rsOrdRm!OrAlyCstRt)
        .Value(wi_Row, "OrWsQty") = Trim(wo_rsOrdRm!OrWsQty)
        .Value(wi_Row, "OrHsQty") = Trim(wo_rsOrdRm!OrHsQty)
        .Value(wi_Row, "OrSubShp") = Trim(wo_rsOrdRm!OrSubShp)
        .Value(wi_Row, "OrMainMet") = Trim(wo_rsOrdRm!OrMainMet)
        wo_rsOrdRm.MoveNext
      Loop
    End With
    Set wo_rsOrdRm = Nothing
    GrdOrdRm.AllowAdd = False

    If ADC("WPWDYN") <> "Y" Then
      GrdOrdRm.ColProp("OrSalRt").Style = fgcinvisible
      GrdOrdRm.ColProp("OrSalVal").Style = fgcinvisible
      GrdOrdRm.ColProp("OrSetSalRt").Style = fgcinvisible
      GrdOrdRm.ColProp("OrSetSalVal").Style = fgcinvisible
      GrdOrdRm.ColProp("OrAlySalRt").Style = fgcinvisible
    Else
      GrdOrdRm.ColProp("OrSalRt").Style = fgcReadOnly
      GrdOrdRm.ColProp("OrSalVal").Style = fgcReadOnly
      GrdOrdRm.ColProp("OrSetSalRt").Style = fgcReadOnly
      GrdOrdRm.ColProp("OrSetSalVal").Style = fgcReadOnly
      GrdOrdRm.ColProp("OrAlySalRt").Style = fgcReadOnly
    End If
'****** Sachin 3.01 - 01-06-07 - Custom Views ******
End Sub

Private Sub FillOrdLab(ByVal RowNum As Integer)
  '****** Sachin 3.01 - 01-06-07 - Custom Views ******
  Dim wsql As String
  Dim wo_rsOrdLab As MwfLib.MDORowSet, wi_Row As Integer
  Dim wCnd As String
    ' ***** Manali 3.10.0 - 09/03/12 -  Allow PDCM in Zoom - ws_CoCd replace gs_CoCd
    'Creation.7-OlCoCd in where condition is changed, given OrdDSg Cocd
    wsql = "Select * from OrdLab " + _
      " Where " + IIF(gs_Partition = ctCurrPrtn, " OlPrtKey='" + ctCurrPrtn + "' And ", "") + " OlCoCd = '" + GrdOrdPnd.Value(RowNum, "WOPCOCD") + "' /*'" + gs_CoCd + "'*/ and OlTc= '" + GrdOrdPnd.Value(RowNum, "wOpTc") + "' and " + _
      "OlYy='" + GrdOrdPnd.Value(RowNum, "wOpYy") + "' and OlChr = '" + GrdOrdPnd.Value(RowNum, "wOpChr") + "' and " + _
      "OlNo = " & GrdOrdPnd.Value(RowNum, "wOpNo") & " and OlSr = " & GrdOrdPnd.Value(RowNum, "wOpSr")
    GrdOrdLab.Rows = 1
    Set wo_rsOrdLab = moCn.OpenRes(wsql)
    With GrdOrdLab
      .AllowAdd = True
      Do While Not (wo_rsOrdLab.EOF Or wo_rsOrdLab.BOF)
        .AddItem
        wi_Row = .Rows - 1
        .Value(wi_Row, "OlSrNo") = Trim(wo_rsOrdLab!OlSrNo)
        .Value(wi_Row, "OlMCd") = Trim(wo_rsOrdLab!OlMCd)
        .Value(wi_Row, "OlSCd") = Trim(wo_rsOrdLab!OlSCd)
        .Value(wi_Row, "OlQW") = Trim(wo_rsOrdLab!OlQw)
        .Value(wi_Row, "OlQty") = Trim(wo_rsOrdLab!OlQty)
        .Value(wi_Row, "OlSalRt") = Trim(wo_rsOrdLab!OlSalRt)
        .Value(wi_Row, "OlSalVal") = Trim(wo_rsOrdLab!OlSalVal)
        .Value(wi_Row, "OlCstQW") = Trim(wo_rsOrdLab!OlCstQw)
        .Value(wi_Row, "OlCstRt") = Trim(wo_rsOrdLab!OlCstRt)
        .Value(wi_Row, "OlCstVal") = Trim(wo_rsOrdLab!OlCstVal)
        wo_rsOrdLab.MoveNext
      Loop
    End With
    Set wo_rsOrdLab = Nothing
    GrdOrdLab.AllowAdd = False
    If ADC("WPWDYN") <> "Y" Then
      GrdOrdLab.ColProp("OlSalRt").Style = fgcinvisible
      GrdOrdLab.ColProp("OlSalVal").Style = fgcinvisible
    Else
      GrdOrdLab.ColProp("OlSalRt").Style = fgcReadOnly
      GrdOrdLab.ColProp("OlSalVal").Style = fgcReadOnly
    End If
'****** Sachin 3.01 - 01-06-07 - Custom Views ******
End Sub
'*** Jay 3.01 *** [Custom Views]
Private Sub FillOrdCt(ByVal RowNum As Integer)
  Dim wsql As String
  Dim wo_rsOrdCt As MwfLib.MDORowSet, wi_Row As Integer
  Dim wCnd As String
  ' ***** Manali 3.10.0 - 09/03/12 -  Allow PDCM in Zoom - ws_CoCd replace gs_CoCd
  wsql = "Select * from OrdCt " + _
    " Where " + IIF(gs_Partition = ctCurrPrtn, " OctPrtKey='" + ctCurrPrtn + "' And ", "") + " OctCoCd = '" + GrdOrdPnd.Value(RowNum, "WOPCOCD") + "' /*'" + gs_CoCd + "'*/ and OctTc= '" + GrdOrdPnd.Value(RowNum, "wOpTc") + "' and " + _
    "OctYy='" + GrdOrdPnd.Value(RowNum, "wOpYy") + "' and OctChr = '" + GrdOrdPnd.Value(RowNum, "wOpChr") + "' and " + _
    "OctNo = " & GrdOrdPnd.Value(RowNum, "wOpNo") & " and OctSr = " & GrdOrdPnd.Value(RowNum, "wOpSr")
   GrdOrdCT.Rows = 1
   Set wo_rsOrdCt = moCn.OpenRes(wsql)
   With GrdOrdCT
     .AllowAdd = True
     Do While Not (wo_rsOrdCt.EOF Or wo_rsOrdCt.BOF)
       .AddItem
       wi_Row = .Rows - 1
       .Value(wi_Row, "OctSrNo") = Trim(wo_rsOrdCt!OctSrNo)
       .Value(wi_Row, "OctCd") = Trim(wo_rsOrdCt!OctCd)
       .Value(wi_Row, "OctSz") = Trim(wo_rsOrdCt!OctSz)
       .Value(wi_Row, "OctCtChr") = Trim(wo_rsOrdCt!OctCtChr)
       .Value(wi_Row, "OctParts") = Trim(wo_rsOrdCt!OctParts)
       .Value(wi_Row, "OctRmCd") = Trim(wo_rsOrdCt!OctRmCd)
       .Value(wi_Row, "OctDmCol") = Trim(wo_rsOrdCt!OctDmCol)
       .Value(wi_Row, "OctRmWt") = Trim(wo_rsOrdCt!OctRmWt)
       .Value(wi_Row, "OctDc") = Trim(wo_rsOrdCt!OctDc)
       .Value(wi_Row, "OctPrdSeq") = Trim(wo_rsOrdCt!OctPrdSeq)
       .Value(wi_Row, "OctMainPrtYn") = Trim(wo_rsOrdCt!OctMainPrtYn)
       .Value(wi_Row, "OctDesc") = Trim(wo_rsOrdCt!OctDesc)
       wo_rsOrdCt.MoveNext
     Loop
   End With
  Set wo_rsOrdCt = Nothing
  GrdOrdCT.AllowAdd = False
End Sub

'*** Jay 3.01 *** [Custom Views]
Private Sub TabOrdDsgDet_Click(PreviousTab As Integer)
 '*** Whenever a tab is selected enable the frame on the clicked tab and disable the previous tab frame
  FraOrdDsgDet(TabOrdDsgDet.Tab).Enabled = True:  FraOrdDsgDet(PreviousTab).Enabled = False
  'Call DispFra(FraOrdDsgDet(TabOrdDsgDet.Tab))
End Sub
'*** Jay 3.01 *** [Custom Views]
Private Sub DispFra(ByVal pv_OrdDsgFra As en_OrdDsgFra)
  Select Case pv_OrdDsgFra
  Case Is = DsgDet
    If FraOrdDsg.Visible = True Then
      FraOrdDsg.Visible = False:      FraOrdDsg.Enabled = False
      Call EnaDisaCmds(False, CmdDsgDet)
      CmdDsgDet.SetFocus
    Else
      FraOrdDsg.Visible = True:      FraOrdDsg.Enabled = True
      TabOrdDsgDet.SetFocus:         TabOrdDsgDet.Tab = 0
      FraOrdDsgDet(TabOrdDsgDet.Tab).Enabled = True
      FraOrdDsg.ZOrder
      Call EnaDisaCmds(True, CmdDsgDet)
    End If
  Case Is = InvDet
    If FraInvDet.Visible = False Then
      FraInvDet.ZOrder (0)
      FraInvDet.Visible = True
      cmd_InvDet.Caption = "&OrdDet"
      ADC.AllowSave = False
      FraInvDet.Enabled = True
      FraNKeyAll.Enabled = False
      CmdDsgDet.Enabled = False
      If GrdOrdPnd.Rows > 1 Then
        ADC("wOdTc") = GrdOrdPnd.Value(GrdOrdPnd.Row, "wOpTc")
        ADC("wOdYy") = GrdOrdPnd.Value(GrdOrdPnd.Row, "wOpYy")
        ADC("wOdChr") = GrdOrdPnd.Value(GrdOrdPnd.Row, "wOpChr")
        ADC("wOdNo") = GrdOrdPnd.Value(GrdOrdPnd.Row, "wOpNo")
        ADC("wOdSr") = GrdOrdPnd.Value(GrdOrdPnd.Row, "wOpSr")
        Call fillGrdInv(GrdOrdPnd.Row)
        GrdInv.SetFocus
      End If
    Else
      ADC("wOdTc") = "": ADC("wOdYy") = "":  ADC("wOdChr") = "":  ADC("wOdNo") = "":  ADC("wOdSr") = ""
      FraInvDet.Visible = False
      cmd_InvDet.Caption = "&InvDet"
      ADC.AllowSave = True
      FraInvDet.Enabled = False
      FraNKeyAll.Enabled = True
      FraNKeyAll.ZOrder (0)
      CmdDsgDet.Enabled = True
      GrdOrdPnd.SetFocus
    End If
  
  '****** Sachin 3.3.0 (26-06-08) - [21. Copy From Order Zoom]******
    Case Is = ViewSel
      If FraCpySel.Visible = True Then
        FraCpySel.Visible = False
        Call EnaDisaCmds(False, CmdViewSel)
        CmdViewSel.SetFocus
      Else
        FraCpySel.Visible = True
        FraCpySel.ZOrder
        Call EnaDisaCmds(True, CmdViewSel)
      End If
  '****** Sachin 3.3.0 (26-06-08) - [21. Copy From Order Zoom]******
    
  End Select
End Sub
'*** Jay 3.01 *** [Custom Views]
Private Sub EnaDisaCmds(ByVal pv_ShowFra As Boolean, Optional ByVal pv_CmdBut As MWCTL_BTN1)
  If Not pv_CmdBut Is Nothing Then
    If pv_ShowFra = True Then
      If pv_CmdBut.Pic = 0 Then pv_CmdBut.BackColor = vbWhite
      pv_CmdBut.CausesValidation = True
    Else
      If pv_CmdBut.Pic = 0 Then pv_CmdBut.BackColor = &H8000000F
    End If
  End If
  FraNKeyAll.Enabled = Not pv_ShowFra
  ADC.AllowSave = Not pv_ShowFra
  'ADC.AllowDelete = Not pv_ShowFra
  cmd_InvDet.Enabled = Not pv_ShowFra
  CmdDsgDet.Enabled = Not pv_ShowFra
  CmdViewSel.Enabled = Not pv_ShowFra     '****** Sachin 3.3.0 (27-06-08) - [21. Copy From Order Zoom] ******
  If Not pv_CmdBut Is Nothing Then pv_CmdBut.Enabled = True
End Sub

Private Sub TxtPwd_Validate(Cancel As Boolean)
'****** Sachin 3.01 - 01-06-07 - Custom Views ******
  If Trim(TxtPwd.text) = "" Then
    ADC("WPWDYN") = ""
    ChkBoxArr(1).Value = Unchecked      ' ***** Manali 3.5.0 - 06/05/09 - Check Boxes Added
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
'****** Sachin 3.01 - 01-06-07 - Custom Views ******
End Sub
Private Sub Pic3D_DblClick()
'****** Manali 3.2.2 Enlarge Pic
  '*** If wPic3DDblClk flag option is False then Enlarge the 3-Dimensional picture to fit the FraNKeyAll Frame
  '*** If wPic3DDblClk flag option is True then bring the 3-Dimensional picture to the original pic control size

    If GrdOrdPnd.Rows - 1 = 0 Then Exit Sub
    ADC.SetFocus
    Dim w3DPth As String
    If wPic3DDblClk = False Then
        '*** Enlarge The Picture ***
        w3DPth = GetPictPath("DM", False, GrdOrdPnd.Value(GrdOrdPnd.Row, "wOpDsgCd"), GrdOrdPnd.Value(GrdOrdPnd.Row, "wOpSz"), , , , , , , GrdOrdPnd.Value(GrdOrdPnd.Row, "wOpPicNm"))

        Call ShowLargePic(IIF(FraNKeyAll.Height - 195 > 0, FraNKeyAll.Height - 195, 0), FraNKeyAll.Width, w3DPth, Pic3D)

        FraPic3D.Left = 0: FraPic3D.Top = 195
        FraPic3D.Height = Pic3D.Height + 100: FraPic3D.Width = Pic3D.Width + 100
    FraNKeyAll.ZOrder
    FraPic3D.ZOrder (0)
    Else
        '*** Make The Picture Small ***
        w3DPth = GetPictPath("DM", False, GrdOrdPnd.Value(GrdOrdPnd.Row, "wOpDsgCd"), GrdOrdPnd.Value(GrdOrdPnd.Row, "wOpSz"), , , , , , , GrdOrdPnd.Value(GrdOrdPnd.Row, "wOpPicNm"))
        Call ShowPic(wPic3DHt, wPic3DWdt, w3DPth, Pic3D)
        FraPic3D.Left = wFraPic3DLft: FraPic3D.Top = wFraPic3DTop
        FraPic3D.Height = Pic3D.Height + 100: FraPic3D.Width = Pic3D.Width + 100
        FraNKeyAll.ZOrder
    End If
    If wPic3DDblClk = False Then
        wPic3DDblClk = True
    Else
        wPic3DDblClk = False
    End If
End Sub
Private Sub PicHsk_DblClick()
'****** Manali 3.2.2 Enlarge Pic
  '*** If wPicHskDblClk flag option is False then Enlarge the Line Drawing to fit the FraNKeyAll Frame
  '*** If wPicHskDblClk flag option is True then bring the Line Drawing to the original pic control size

    If GrdOrdPnd.Rows - 1 = 0 Then Exit Sub
    ADC.SetFocus
    Dim wHskPth As String
    If wPicHskDblClk = False Then
        '*** Enlarge The Picture ***
        wHskPth = GetPictPath("DM", True, GrdOrdPnd.Value(GrdOrdPnd.Row, "wOpDsgCd"), GrdOrdPnd.Value(GrdOrdPnd.Row, "wOpSz"), , , , , , , GrdOrdPnd.Value(GrdOrdPnd.Row, "wOpPicNm"))
        Call ShowLargePic(IIF(FraNKeyAll.Height - 195 > 0, FraNKeyAll.Height - 195, 0), FraNKeyAll.Width, wHskPth, PicHsk)
        FraPicHsk.Left = 0: FraPicHsk.Top = 195
        FraPicHsk.Height = PicHsk.Height + 100: FraPicHsk.Width = PicHsk.Width + 100
        FraNKeyAll.ZOrder
        FraPicHsk.ZOrder (0)
    Else
        '*** Make The Picture Small ***
    wHskPth = GetPictPath("DM", True, GrdOrdPnd.Value(GrdOrdPnd.Row, "wOpDsgCd"), GrdOrdPnd.Value(GrdOrdPnd.Row, "wOpSz"), , , , , , , GrdOrdPnd.Value(GrdOrdPnd.Row, "wOpPicNm"))
        Call ShowPic(wPicHskHt, wPicHskWdt, wHskPth, PicHsk)
        FraPicHsk.Left = wFraPicHskLft: FraPicHsk.Top = wFraPicHskTop
        FraPicHsk.Height = PicHsk.Height + 100: FraPicHsk.Width = PicHsk.Width + 100
        FraNKeyAll.ZOrder
    End If
    If wPicHskDblClk = False Then
        wPicHskDblClk = True
    Else
        wPicHskDblClk = False
    End If
End Sub

