VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{AB2CC32A-E652-42A2-B341-26ECE02F6386}#2.1#0"; "mwfctlforspeed9.ocx"
Begin VB.Form EmrFrmJtZoom 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Jewellery Trading Zoom"
   ClientHeight    =   9810
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15210
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11040
   ScaleWidth      =   20400
   WindowState     =   2  'Maximized
   Begin VB.Frame FRA_KEYS 
      BorderStyle     =   0  'None
      Height          =   9255
      Left            =   0
      TabIndex        =   52
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
         TabIndex        =   45
         ToolTipText     =   "Select Export Delivery Date Or Production Delivery Date"
         Top             =   4875
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
         Left            =   2850
         TabIndex        =   44
         ToolTipText     =   "Select Export Delivery Date Or Production Delivery Date"
         Top             =   4875
         Width           =   1695
      End
      Begin VB.CheckBox ChkBoxArr 
         Alignment       =   1  'Right Justify
         Caption         =   "Sort By Customer"
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
         TabIndex        =   43
         Tag             =   "SortCmCd"
         ToolTipText     =   "Check To Show Sort By Customer Cd"
         Top             =   4500
         Width           =   3045
      End
      Begin VB.CheckBox ChkBoxArr 
         Alignment       =   1  'Right Justify
         Caption         =   "Show Price                        "
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
         TabIndex        =   42
         Tag             =   "ShowVal"
         ToolTipText     =   "Check To Show Value"
         Top             =   4185
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
         TabIndex        =   40
         Tag             =   "ShowPnd"
         ToolTipText     =   "Check To Show Pending Records Only"
         Top             =   3570
         Width           =   3045
      End
      Begin VB.TextBox TxtPwd 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   2880
         MaxLength       =   30
         PasswordChar    =   "*"
         TabIndex        =   41
         Top             =   3900
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
         IdName          =   "WOMTCFR"
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
         IdName          =   "WOMYYFR"
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
         IdName          =   "WOMCHRFR"
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
         IdName          =   "WOMTCTO"
         CmpStr          =   "OmTc<="
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   7
         Left            =   5820
         TabIndex        =   6
         ToolTipText     =   "Enter Fr Voucher Sr"
         Top             =   270
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   503
         Alignment       =   2
         Mask            =   "###0"
         MaxLength       =   4
         DataType        =   1
         IdName          =   "WODSRFR"
         CmpStr          =   "OdSr>="
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   6
         Left            =   4890
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
         IdName          =   "WOMNOFR"
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
         IdName          =   "WOMYYTO"
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
         IdName          =   "WOMCHRTO"
         CmpStr          =   "OmChr<="
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   9
         Left            =   9750
         TabIndex        =   11
         ToolTipText     =   "Enter To Voucher Sr"
         Top             =   270
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   503
         Alignment       =   2
         Mask            =   "###0"
         MaxLength       =   4
         DataType        =   1
         IdName          =   "WODSRTO"
         CmpStr          =   "OdSr<="
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   8
         Left            =   8820
         TabIndex        =   10
         ToolTipText     =   "Enter To Voucher No"
         Top             =   270
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   503
         Alignment       =   2
         Mask            =   "#####0"
         MaxLength       =   6
         DataType        =   1
         IdName          =   "WOMNOTO"
         CmpStr          =   "OmNo<="
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   18
         Left            =   2880
         TabIndex        =   20
         ToolTipText     =   "Enter From PO No "
         Top             =   1720
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
         TabIndex        =   21
         ToolTipText     =   "Enter To PO No "
         Top             =   1720
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
         TabIndex        =   27
         ToolTipText     =   "Enter To Design Category"
         Top             =   2330
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   503
         MaxLength       =   5
         DataType        =   4
         IdName          =   "WDSGCTGTO"
         CmpStr          =   "DmCtg<="
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   21
         Left            =   2880
         TabIndex        =   26
         ToolTipText     =   "Enter From Design Category"
         Top             =   2330
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   503
         MaxLength       =   5
         DataType        =   4
         IdName          =   "WDSGCTGFR"
         CmpStr          =   "DmCtg>="
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   22
         Left            =   2880
         TabIndex        =   28
         ToolTipText     =   "Enter From Design Code"
         Top             =   2620
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
         TabIndex        =   29
         ToolTipText     =   "Enter To Design Code"
         Top             =   2620
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
         TabIndex        =   47
         ToolTipText     =   "Enter 'Y' to Show Only Pending Orders/ 'N'  To show all orders"
         Top             =   3570
         Visible         =   0   'False
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   503
         Enabled         =   0   'False
         DataType        =   4
         IdName          =   "WPNDYN"
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
         TabIndex        =   18
         ToolTipText     =   "Enter From Customer (Customer As Mentioned In Transactions)"
         Top             =   1420
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
         TabIndex        =   19
         ToolTipText     =   "Enter To Customer (Customer As Mentioned In Transactions)"
         Top             =   1420
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
         IdName          =   "WODDELDTTO"
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
         IdName          =   "WODDELDTFR"
         CmpStr          =   "OdDelDt >="
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   42
         Left            =   3420
         TabIndex        =   48
         ToolTipText     =   "Enter 'Y' to Show Value"
         Top             =   4185
         Visible         =   0   'False
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   503
         Enabled         =   0   'False
         DataType        =   4
         IdName          =   "WVALYN"
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   27
         Left            =   8085
         TabIndex        =   25
         ToolTipText     =   "Enter To Sub PO No "
         Top             =   2025
         Width           =   2595
         _ExtentX        =   4577
         _ExtentY        =   503
         MaxLength       =   20
         IdName          =   "WODPONOTO"
         CmpStr          =   "OdPONO<="
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   26
         Left            =   4150
         TabIndex        =   23
         ToolTipText     =   "Enter From Sub PO No "
         Top             =   2025
         Width           =   2595
         _ExtentX        =   4577
         _ExtentY        =   503
         MaxLength       =   20
         IdName          =   "WODPONOFR"
         CmpStr          =   "OdPONO>="
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   29
         Left            =   3420
         TabIndex        =   49
         ToolTipText     =   "Enter 'Y' to Sort By Customer Code"
         Top             =   4500
         Visible         =   0   'False
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   503
         Enabled         =   0   'False
         DataType        =   4
         IdName          =   "WSORTCMCD"
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
         IdName          =   "WCOCDTO"
         CmpStr          =   "OmCoCd <= "
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   32
         Left            =   6240
         TabIndex        =   50
         ToolTipText     =   "Show Export/ Production Delivery Date (Enter ""Y"" for Export and ""N"" for Production)"
         Top             =   4875
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
         TabIndex        =   30
         ToolTipText     =   "Enter From Cust Dsg Cd"
         Top             =   2920
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   503
         MaxLength       =   40
         DataType        =   4
         IdName          =   "WCDCUSTDMCDFR"
         CmpStr          =   "isnull(CdCustDmCd,'')>="
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   34
         Left            =   6810
         TabIndex        =   31
         ToolTipText     =   "Enter To Cust Dsg Cd"
         Top             =   2920
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   503
         MaxLength       =   40
         DataType        =   4
         IdName          =   "WCDCUSTDMCDTO"
         CmpStr          =   "isnull(CdCustDmCd,'')<="
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   35
         Left            =   4580
         TabIndex        =   35
         ToolTipText     =   "Enter From Stock Number"
         Top             =   3220
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   503
         Mask            =   "######0"
         MaxLength       =   7
         DataType        =   1
         IdName          =   "WODDMSTKNOFR"
         CmpStr          =   "OdDmStkNo >="
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   36
         Left            =   8480
         TabIndex        =   39
         ToolTipText     =   "Enter To Stock Number"
         Top             =   3225
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   503
         Mask            =   "######0"
         MaxLength       =   7
         DataType        =   1
         IdName          =   "WODDMSTKNOTO"
         CmpStr          =   "OdDmStkNo <="
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   40
         Left            =   3850
         TabIndex        =   34
         ToolTipText     =   "Enter From Stock Character"
         Top             =   3220
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   503
         MaxLength       =   3
         DataType        =   4
         IdName          =   "WODDMSTKCHRFR"
         CmpStr          =   "OdDmStkChr >="
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   41
         Left            =   3480
         TabIndex        =   33
         ToolTipText     =   "Enter From Stock Year"
         Top             =   3220
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   503
         MaxLength       =   2
         DataType        =   4
         IdName          =   "WODDMSTKYYFR"
         CmpStr          =   "OdDmStkYy >="
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   43
         Left            =   2880
         TabIndex        =   32
         ToolTipText     =   "Enter From Stock Company Code"
         Top             =   3220
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   503
         MaxLength       =   3
         DataType        =   4
         IdName          =   "WODDMSTKCOCDFR"
         CmpStr          =   "OdDmStkCoCd >="
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   44
         Left            =   7755
         TabIndex        =   38
         ToolTipText     =   "Enter To Stock Character"
         Top             =   3220
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   503
         MaxLength       =   3
         DataType        =   4
         IdName          =   "WODDMSTKCHRTO"
         CmpStr          =   "OdDmStkChr <="
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   45
         Left            =   7400
         TabIndex        =   37
         ToolTipText     =   "Enter To Stock Year"
         Top             =   3225
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   503
         MaxLength       =   2
         DataType        =   4
         IdName          =   "WODDMSTKYYTO"
         CmpStr          =   "OdDmStkYy <="
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   46
         Left            =   6810
         TabIndex        =   36
         ToolTipText     =   "Enter To Stock Company Code"
         Top             =   3220
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   503
         MaxLength       =   3
         DataType        =   4
         IdName          =   "WODDMSTKCOCDTO"
         CmpStr          =   "OdDmStkCoCd <="
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   49
         Left            =   2880
         TabIndex        =   16
         ToolTipText     =   "Enter From Customer Category"
         Top             =   1125
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   503
         MaxLength       =   2
         DataType        =   4
         IdName          =   "WOMCMCTGFR"
         CmpStr          =   "OmCmCtg>="
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   50
         Left            =   6810
         TabIndex        =   17
         ToolTipText     =   "Enter To Customer Category"
         Top             =   1125
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   503
         MaxLength       =   2
         DataType        =   4
         IdName          =   "WOMCMCTGTO"
         CmpStr          =   "OmCmCtg <="
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   86
         Left            =   2880
         TabIndex        =   22
         ToolTipText     =   "Enter From Sub Customer"
         Top             =   2025
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
         Index           =   10
         Left            =   6810
         TabIndex        =   24
         ToolTipText     =   "Enter To Sub Customer"
         Top             =   2025
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         MaxLength       =   8
         DataType        =   4
         IdName          =   "WODSUBCUSTTO"
         CmpStr          =   "OdSubCust <="
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   31
         Left            =   2880
         TabIndex        =   0
         ToolTipText     =   "Enter From Company Code"
         Top             =   0
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   503
         MaxLength       =   3
         DataType        =   4
         IdName          =   "WCOCDFR"
         CmpStr          =   "OmCoCd >= "
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   11
         Left            =   2880
         TabIndex        =   46
         ToolTipText     =   "Enter Currency Code Whose Orders You Wish To View Or Leave It Blank If You Wish To View All Orders In Base Currency"
         Top             =   5280
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   503
         MaxLength       =   5
         DataType        =   4
         IdName          =   "WCURCD"
         CmpStr          =   "OmCmCurCd =  "
      End
      Begin VB.Label LblCurCd 
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
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   60
         TabIndex        =   125
         Top             =   5280
         Width           =   1185
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
         Index           =   2
         Left            =   8020
         TabIndex        =   103
         Top             =   2025
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
         Index           =   1
         Left            =   4100
         TabIndex        =   102
         Top             =   2040
         Width           =   60
      End
      Begin VB.Label LblDmId 
         BackStyle       =   0  'Transparent
         Caption         =   "StkNum"
         BeginProperty Font 
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
         Top             =   3220
         Width           =   1425
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Cust Ctg"
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
         TabIndex        =   86
         Top             =   1120
         Width           =   825
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
         TabIndex        =   85
         Top             =   2920
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
         Left            =   60
         TabIndex        =   84
         ToolTipText     =   "Location"
         Top             =   4875
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
         TabIndex        =   83
         Top             =   0
         Width           =   2145
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SUB PO No"
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
         TabIndex        =   82
         Top             =   2025
         Width           =   1065
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
         TabIndex        =   80
         Top             =   3900
         Width           =   1455
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
         TabIndex        =   73
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
         TabIndex        =   72
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
         TabIndex        =   71
         Top             =   1420
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
         TabIndex        =   70
         Top             =   1720
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
         TabIndex        =   69
         Top             =   2330
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
         TabIndex        =   68
         Top             =   2620
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
         TabIndex        =   67
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
         TabIndex        =   66
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
         Left            =   9630
         TabIndex        =   63
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
         Left            =   8720
         TabIndex        =   62
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
         Left            =   7500
         TabIndex        =   61
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
         Left            =   7950
         TabIndex        =   60
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
         Left            =   5685
         TabIndex        =   59
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
         Left            =   4800
         TabIndex        =   58
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
         Left            =   4020
         TabIndex        =   56
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
         Left            =   3580
         TabIndex        =   55
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
         TabIndex        =   54
         Top             =   285
         Width           =   2145
      End
   End
   Begin MwfCtl.MWCTL_BTN1 cmdBasChild 
      Height          =   420
      Left            =   6030
      TabIndex        =   105
      Top             =   9270
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   741
      BackColor       =   -2147483633
      MaskColor       =   8454143
      Caption         =   "&Base Child"
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
   Begin MwfCtl.MWCTL_BTN1 cmdRefChild 
      Height          =   420
      Left            =   7065
      TabIndex        =   104
      Top             =   9270
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   741
      BackColor       =   -2147483633
      MaskColor       =   8454143
      Caption         =   "&Reference Child"
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
   Begin MwfCtl.MWCTL_BTN1 CmdShowPic 
      Height          =   420
      Left            =   9045
      TabIndex        =   94
      Top             =   9270
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   741
      BackColor       =   -2147483633
      MaskColor       =   8454143
      Caption         =   "Show &Pic"
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
      Left            =   8100
      TabIndex        =   79
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
   Begin MwfCtl.MWCTL_HLP MWCTL_HLP1 
      Height          =   435
      Left            =   9750
      TabIndex        =   53
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
      TabIndex        =   51
      Top             =   9270
      Width           =   15135
      _ExtentX        =   26696
      _ExtentY        =   979
      AllowAdd        =   0   'False
      AllowDelete     =   0   'False
      BtnWidth        =   885
      BtnHeight       =   435
      BtnMaskColor    =   8454143
      FontSz          =   8.25
   End
   Begin VB.Frame FRA_NKEYS 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   9720
      Left            =   -150
      TabIndex        =   57
      Top             =   15
      Width           =   15315
      Begin VB.Frame FraNKeyAll 
         BorderStyle     =   0  'None
         Height          =   9225
         Left            =   90
         TabIndex        =   64
         Top             =   0
         Width           =   15210
         Begin MwfCtl.MWCTL_FLX GrdBasChild 
            Height          =   2565
            Left            =   90
            TabIndex        =   106
            Top             =   6720
            Width           =   15030
            _ExtentX        =   26511
            _ExtentY        =   4524
            Cols            =   18
            AllowAdd        =   0   'False
            AllowDelete     =   0   'False
            colname1        =   "WREFBSCOCD"
            maxlength1      =   3
            style1          =   0
            colname2        =   "WREFBSTC"
            maxlength2      =   3
            style2          =   0
            colname3        =   "WREFBSYY"
            maxlength3      =   2
            style3          =   0
            colname4        =   "WREFBSCHR"
            maxlength4      =   3
            style4          =   0
            colname5        =   "WREFBSNO"
            datatype5       =   1
            mask5           =   "#########0"
            maxlength5      =   10
            style5          =   0
            tooltiptext5    =   "Order No."
            colname6        =   "WREFBSORDNO"
            heading6        =   "Voucher No                                           "
            colname7        =   "WREFBSORDDT"
            heading7        =   "Date         "
            datatype7       =   3
            maxlength7      =   8
            colname8        =   "WREFBSCUST"
            heading8        =   "Customer        "
            datatype8       =   4
            maxlength8      =   8
            colname9        =   "WREFBSSR"
            heading9        =   "Sr       "
            datatype9       =   1
            mask9           =   "####0"
            maxlength9      =   5
            colname10       =   "WREFBSSTKCOCD"
            maxlength10     =   3
            style10         =   0
            colname11       =   "WREFBSSTKYY"
            maxlength11     =   2
            style11         =   0
            colname12       =   "WREFBSSTKCHR"
            maxlength12     =   3
            style12         =   0
            colname13       =   "WREFBSSTKNO"
            datatype13      =   1
            mask13          =   "######0"
            maxlength13     =   7
            style13         =   0
            colname14       =   "WREFBSSTKID"
            heading14       =   "StkId                                        "
            colname15       =   "WREFBSORDQTY"
            heading15       =   "Qty           "
            datatype15      =   1
            mask15          =   "#####0.0"
            maxlength15     =   8
            style15         =   2
            colname16       =   "WREFBSSALPRC"
            heading16       =   "Price                      "
            datatype16      =   2
            mask16          =   "#######0.#0"
            maxlength16     =   11
            colname17       =   "WREFBSSUBPONO"
            heading17       =   "SUB PO                                                                "
            datatype17      =   4
            maxlength17     =   35
         End
         Begin MwfCtl.MWCTL_FLX GrdRefChild 
            Height          =   2595
            Left            =   120
            TabIndex        =   88
            Top             =   6690
            Width           =   15060
            _ExtentX        =   26564
            _ExtentY        =   4577
            Cols            =   18
            AllowAdd        =   0   'False
            AllowDelete     =   0   'False
            colname1        =   "WREFBSCOCD"
            maxlength1      =   3
            style1          =   0
            colname2        =   "WREFBSTC"
            maxlength2      =   3
            style2          =   0
            colname3        =   "WREFBSYY"
            maxlength3      =   2
            style3          =   0
            colname4        =   "WREFBSCHR"
            maxlength4      =   3
            style4          =   0
            colname5        =   "WREFBSNO"
            datatype5       =   1
            mask5           =   "#########0"
            maxlength5      =   10
            style5          =   0
            tooltiptext5    =   "Order No."
            colname6        =   "WREFBSORDNO"
            heading6        =   "Voucher No                                           "
            colname7        =   "WREFBSORDDT"
            heading7        =   "Date         "
            datatype7       =   3
            maxlength7      =   8
            colname8        =   "WREFBSCUST"
            heading8        =   "Customer        "
            datatype8       =   4
            maxlength8      =   8
            colname9        =   "WREFBSSR"
            heading9        =   "Sr       "
            datatype9       =   1
            mask9           =   "####0"
            maxlength9      =   5
            colname10       =   "WREFBSSTKCOCD"
            maxlength10     =   3
            style10         =   0
            colname11       =   "WREFBSSTKYY"
            maxlength11     =   2
            style11         =   0
            colname12       =   "WREFBSSTKCHR"
            maxlength12     =   3
            style12         =   0
            colname13       =   "WREFBSSTKNO"
            datatype13      =   1
            mask13          =   "######0"
            maxlength13     =   7
            style13         =   0
            colname14       =   "WREFBSSTKID"
            heading14       =   "StkId                                        "
            colname15       =   "WREFBSORDQTY"
            heading15       =   "Qty           "
            datatype15      =   1
            mask15          =   "#####0.0"
            maxlength15     =   8
            style15         =   2
            colname16       =   "WREFBSSALPRC"
            heading16       =   "Price                     "
            datatype16      =   2
            mask16          =   "#######0.#0"
            maxlength16     =   11
            colname17       =   "WREFBSSUBPONO"
            heading17       =   "SUB PO                                                                "
            datatype17      =   4
            maxlength17     =   35
         End
         Begin VB.Frame FraPicHsk 
            BorderStyle     =   0  'None
            Height          =   1965
            Left            =   10710
            TabIndex        =   77
            Top             =   30
            Width           =   1965
            Begin VB.Image PicHsk 
               Height          =   1905
               Left            =   0
               Top             =   30
               Width           =   1905
            End
         End
         Begin VB.Frame FraPic3D 
            BorderStyle     =   0  'None
            Height          =   1965
            Left            =   12900
            TabIndex        =   76
            Top             =   30
            Width           =   1965
            Begin VB.Image Pic3D 
               Height          =   1905
               Left            =   30
               Top             =   30
               Width           =   1905
            End
         End
         Begin MwfCtl.MWCTL_FLX GrdOrdMst 
            Height          =   2415
            Left            =   150
            TabIndex        =   78
            Top             =   0
            Width           =   10230
            _ExtentX        =   18045
            _ExtentY        =   4260
            Cols            =   11
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
            heading5        =   "Voucher No                                          "
            colname6        =   "WOMPRDEXPDELDT"
            heading6        =   "PrdExpDelDt"
            datatype6       =   3
            maxlength6      =   8
            tooltiptext6    =   "Export Delivery Date"
            colname7        =   "WOMCUST"
            heading7        =   "Customer        "
            datatype7       =   4
            maxlength7      =   8
            tooltiptext7    =   "Customer Code"
            colname8        =   "WOMPONO"
            heading8        =   "Purchase  Order  Number             "
            maxlength8      =   20
            tooltiptext8    =   "PO Number"
            colname9        =   "WOMPODT"
            heading9        =   "Vchr Date"
            datatype9       =   3
            maxlength9      =   8
            tooltiptext9    =   "PO Date"
            colname10       =   "WOMCOCD"
            maxlength10     =   3
            style10         =   0
         End
         Begin MwfCtl.MWCTL_FLX GrdOrdDsg 
            Height          =   3615
            Left            =   120
            TabIndex        =   65
            Top             =   2460
            Width           =   15060
            _ExtentX        =   26564
            _ExtentY        =   6376
            Cols            =   31
            AllowAdd        =   0   'False
            AllowDelete     =   0   'False
            colname1        =   "WODCOCD"
            maxlength1      =   3
            style1          =   0
            colname2        =   "WODTC"
            maxlength2      =   3
            style2          =   0
            colname3        =   "WODYY"
            maxlength3      =   2
            style3          =   0
            colname4        =   "WODCHR"
            maxlength4      =   3
            style4          =   0
            colname5        =   "WODNO"
            datatype5       =   1
            mask5           =   "#########0"
            maxlength5      =   10
            style5          =   0
            tooltiptext5    =   "Order No."
            colname6        =   "WODSR"
            heading6        =   "Sr       "
            datatype6       =   1
            mask6           =   "####0"
            maxlength6      =   5
            colname7        =   "WODPRDEXPDELDT"
            heading7        =   "ExpDelDt  "
            datatype7       =   3
            maxlength7      =   8
            colname8        =   "WODDSGCD"
            heading8        =   "Design   Code                     "
            maxlength8      =   15
            tooltiptext8    =   "Design Code "
            colname9        =   "WODSFX"
            heading9        =   "Suffix                        "
            maxlength9      =   10
            tooltiptext9    =   "Design Suffix"
            colname10       =   "WODSZ"
            heading10       =   "Size        "
            maxlength10     =   5
            tooltiptext10   =   "Design Size"
            colname11       =   "WODSTKCOCD"
            maxlength11     =   3
            style11         =   0
            colname12       =   "WODSTKYY"
            maxlength12     =   2
            style12         =   0
            colname13       =   "WODSTKCHR"
            maxlength13     =   3
            style13         =   0
            colname14       =   "WODSTKNO"
            datatype14      =   1
            mask14          =   "######0"
            maxlength14     =   7
            style14         =   0
            colname15       =   "WODSTKID"
            heading15       =   "StkId                                        "
            colname16       =   "WODORDQTY"
            heading16       =   "Qty           "
            datatype16      =   1
            mask16          =   "#####0.0"
            maxlength16     =   8
            style16         =   2
            colname17       =   "WODSALPRC"
            heading17       =   "Price                      "
            datatype17      =   2
            mask17          =   "#######0.#0"
            maxlength17     =   11
            colname18       =   "WODBALQTY"
            heading18       =   "Bal Qty   "
            datatype18      =   1
            mask18          =   "#####0.0"
            maxlength18     =   8
            style18         =   2
            tooltiptext18   =   "Balance Export Qty"
            colname19       =   "WODSUBPONO"
            datatype19      =   4
            maxlength19     =   35
            style19         =   0
            colname20       =   "WODBASEVCH"
            style20         =   0
            colname21       =   "WODREFVCH"
            style21         =   0
            colname22       =   "WODREFQTY"
            datatype22      =   1
            mask22          =   "#####0.0"
            maxlength22     =   8
            style22         =   0
            colname23       =   "WODGRSWT"
            datatype23      =   2
            mask23          =   "####0.##0"
            maxlength23     =   9
            style23         =   0
            colname24       =   "WODMETWT"
            datatype24      =   2
            mask24          =   "-####0.##0"
            maxlength24     =   10
            style24         =   0
            colname25       =   "WODDIAWT"
            datatype25      =   2
            mask25          =   "####0.##0"
            maxlength25     =   9
            style25         =   0
            colname26       =   "WODCSWT"
            datatype26      =   2
            mask26          =   "####0.##0"
            maxlength26     =   9
            style26         =   0
            colname27       =   "WODCHNWT"
            datatype27      =   2
            mask27          =   "####0.##0"
            maxlength27     =   9
            style27         =   0
            colname28       =   "WODPICNM"
            maxlength28     =   45
            style28         =   0
            colname29       =   "WODXWT"
            datatype29      =   2
            mask29          =   "####0.##0"
            maxlength29     =   9
            style29         =   0
            colname30       =   "WODIDNO"
            datatype30      =   1
            mask30          =   "#########0"
            maxlength30     =   10
            style30         =   0
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   12
            Left            =   1290
            TabIndex        =   107
            Top             =   6120
            Width           =   3825
            _ExtentX        =   6747
            _ExtentY        =   503
            Enabled         =   0   'False
            MaxLength       =   32
            DataType        =   4
            IdName          =   "WODSUBPONO"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   87
            Left            =   6000
            TabIndex        =   109
            ToolTipText     =   "Base Voucher"
            Top             =   6090
            Width           =   2940
            _ExtentX        =   5186
            _ExtentY        =   503
            Enabled         =   0   'False
            MaxLength       =   30
            DataType        =   4
            IdName          =   "WODBASEVCH"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   25
            Left            =   6000
            TabIndex        =   110
            ToolTipText     =   "Reference Voucher"
            Top             =   6375
            Width           =   2940
            _ExtentX        =   5186
            _ExtentY        =   503
            Enabled         =   0   'False
            MaxLength       =   30
            DataType        =   4
            IdName          =   "WODREFVCH"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   485
            Left            =   11700
            TabIndex        =   113
            ToolTipText     =   "Enter From Diamond Weight"
            Top             =   6120
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   503
            Alignment       =   2
            Enabled         =   0   'False
            Mask            =   "####0.000"
            MaxLength       =   9
            DataType        =   2
            IdName          =   "WODMETWT"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   110
            Left            =   13890
            TabIndex        =   115
            ToolTipText     =   "Enter From Colour Stone Weight"
            Top             =   6120
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   503
            Alignment       =   2
            Enabled         =   0   'False
            Mask            =   "####0.000"
            MaxLength       =   9
            DataType        =   2
            IdName          =   "WODDIAWT"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   54
            Left            =   13890
            TabIndex        =   117
            Top             =   6405
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   503
            Alignment       =   2
            Enabled         =   0   'False
            Mask            =   "####0.000"
            MaxLength       =   9
            DataType        =   2
            IdName          =   "WODCHNWT"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   55
            Left            =   9660
            TabIndex        =   119
            ToolTipText     =   "Enter From Colour Stone Weight"
            Top             =   6405
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   503
            Alignment       =   2
            Enabled         =   0   'False
            Mask            =   "####0.000"
            MaxLength       =   9
            DataType        =   2
            IdName          =   "WODCSWT"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   13
            Left            =   9660
            TabIndex        =   121
            ToolTipText     =   "Enter From Diamond Weight"
            Top             =   6120
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   503
            Alignment       =   2
            Enabled         =   0   'False
            Mask            =   "####0.000"
            MaxLength       =   9
            DataType        =   2
            IdName          =   "WODGRSWT"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   14
            Left            =   11700
            TabIndex        =   123
            ToolTipText     =   "Enter From Diamond Weight"
            Top             =   6405
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   503
            Alignment       =   2
            Enabled         =   0   'False
            Mask            =   "####0.000"
            MaxLength       =   9
            DataType        =   2
            IdName          =   "WODXWT"
         End
         Begin VB.Label ALBL 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "X Wt"
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
            Left            =   11010
            TabIndex        =   124
            Top             =   6420
            Width           =   450
         End
         Begin VB.Label ALBL 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Met Wt"
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
            Index           =   8
            Left            =   11010
            TabIndex        =   122
            Top             =   6120
            Width           =   660
         End
         Begin VB.Label ALBL 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Col Wt"
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
            Index           =   14
            Left            =   8970
            TabIndex        =   120
            Top             =   6420
            Width           =   630
         End
         Begin VB.Label ALBL 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CHN Wt"
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
            Left            =   13050
            TabIndex        =   118
            Top             =   6420
            Width           =   720
         End
         Begin VB.Label ALBL 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Grs Wt"
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
            Index           =   99
            Left            =   8970
            TabIndex        =   116
            Top             =   6120
            Width           =   630
         End
         Begin VB.Label ALBL 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Dia Wt"
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
            Index           =   367
            Left            =   13050
            TabIndex        =   114
            Top             =   6120
            Width           =   630
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Ref Vch"
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
            Left            =   5160
            TabIndex        =   112
            Top             =   6420
            Width           =   1065
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "BaseVch"
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
            Left            =   5160
            TabIndex        =   111
            Top             =   6120
            Width           =   1065
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "SUB PO No"
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
            Left            =   150
            TabIndex        =   108
            Top             =   6120
            Width           =   1065
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
            TabIndex        =   75
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
            Left            =   13440
            TabIndex        =   74
            Top             =   2010
            Width           =   795
         End
      End
      Begin VB.Frame FraOrdDsg 
         BorderStyle     =   0  'None
         Height          =   5190
         Left            =   60
         TabIndex        =   81
         Top             =   4080
         Width           =   15210
         Begin TabDlg.SSTab TabOrdDsgDet 
            Height          =   5085
            Left            =   120
            TabIndex        =   89
            Top             =   60
            Width           =   15015
            _ExtentX        =   26485
            _ExtentY        =   8969
            _Version        =   393216
            Tabs            =   2
            TabsPerRow      =   2
            TabHeight       =   520
            TabCaption(0)   =   "<&1>Bill Of Materials"
            TabPicture(0)   =   "EmrFrmJTZoom.frx":0000
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "FraOrdDsgDet(0)"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).ControlCount=   1
            TabCaption(1)   =   "<&2> Lab Details"
            TabPicture(1)   =   "EmrFrmJTZoom.frx":001C
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "FraOrdDsgDet(1)"
            Tab(1).ControlCount=   1
            Begin VB.Frame FraOrdDsgDet 
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               Height          =   4665
               Index           =   1
               Left            =   -74940
               TabIndex        =   90
               Top             =   360
               Width           =   14895
               Begin MwfCtl.MWCTL_FLX GrdOrdLab 
                  Height          =   4710
                  Left            =   -30
                  TabIndex        =   91
                  Top             =   -30
                  Width           =   14925
                  _ExtentX        =   26326
                  _ExtentY        =   8308
                  Cols            =   11
                  colname1        =   "OLSRNO"
                  heading1        =   "Sr. "
                  datatype1       =   2
                  mask1           =   "#0"
                  maxlength1      =   2
                  style1          =   2
                  tooltiptext1    =   "Enter Labour Sr No."
                  colname2        =   "OLMCD"
                  heading2        =   "Main Cd   "
                  datatype2       =   4
                  maxlength2      =   5
                  tooltiptext2    =   "Enter Labour Main Code"
                  colname3        =   "OLSCD"
                  heading3        =   "Sub Cd              "
                  datatype3       =   4
                  maxlength3      =   8
                  tooltiptext3    =   "Enter Labour Sub Code"
                  colname4        =   "OLQW"
                  heading4        =   "By Q/W"
                  datatype4       =   4
                  maxlength4      =   1
                  tooltiptext4    =   "Enter whether Labour Sales Value Is Calculated By Qty Or Wt"
                  colname5        =   "OLQTY"
                  heading5        =   "Qty   "
                  datatype5       =   1
                  mask5           =   "##0"
                  maxlength5      =   3
                  tooltiptext5    =   "Enter Labour Quantity"
                  colname6        =   "OLSALRT"
                  heading6        =   "Rate            "
                  datatype6       =   2
                  mask6           =   "#####0.00"
                  maxlength6      =   9
                  tooltiptext6    =   "Enter Labour Sales Rate"
                  colname7        =   "OLSALVAL"
                  heading7        =   "Value             "
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
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               Height          =   4695
               Index           =   0
               Left            =   30
               TabIndex        =   92
               Top             =   360
               Width           =   14925
               Begin MwfCtl.MWCTL_FLX GrdOrdRm 
                  Height          =   4650
                  Left            =   0
                  TabIndex        =   93
                  Top             =   0
                  Width           =   14895
                  _ExtentX        =   26273
                  _ExtentY        =   8202
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
                  heading3        =   "SubCtg    "
                  datatype3       =   4
                  maxlength3      =   5
                  tooltiptext3    =   "Enter Raw Material Sub Category"
                  colname4        =   "ORRMCD"
                  heading4        =   "Rm Code                                  "
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
                  heading6        =   "Brdth   "
                  datatype6       =   2
                  mask6           =   "#0.000"
                  maxlength6      =   6
                  tooltiptext6    =   "Enter RM Breadth"
                  colname7        =   "ORLN3"
                  heading7        =   "Rm Ht "
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
                  heading10       =   "Wt            "
                  datatype10      =   2
                  mask10          =   "-##0.000"
                  maxlength10     =   8
                  tooltiptext10   =   "Enter Raw Material Weigth"
                  colname11       =   "ORLMESAL"
                  heading11       =   "LME Rt    "
                  datatype11      =   2
                  mask11          =   "####0.00"
                  maxlength11     =   8
                  tooltiptext11   =   "Enter LMG/ LMP/LMS Sale Rate"
                  colname12       =   "ORSALRT"
                  heading12       =   "Sale Rt                "
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
                  heading14       =   "Sale Value          "
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
                  heading19       =   "Setting             "
                  datatype19      =   4
                  maxlength19     =   8
                  tooltiptext19   =   "Enter Setting Code"
                  colname20       =   "ORSETSALRT"
                  heading20       =   "Set Rt           "
                  datatype20      =   2
                  mask20          =   "#####0.00"
                  maxlength20     =   9
                  tooltiptext20   =   "Enter Setting Rate"
                  colname21       =   "ORSETSALVAL"
                  heading21       =   "Set Value       "
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
                  heading24       =   "Alloy                                         "
                  datatype24      =   4
                  maxlength24     =   16
                  tooltiptext24   =   "Enter Alloy Code"
                  colname25       =   "ORALYSALRT"
                  heading25       =   "Alloy Rt          "
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
                  heading29       =   "SShp"
                  datatype29      =   4
                  maxlength29     =   3
                  tooltiptext29   =   "Enter Sub Shape"
                  colname30       =   "ORMAINMET"
                  heading30       =   "MMet"
                  datatype30      =   4
                  maxlength30     =   1
                  tooltiptext30   =   "Enter Main Metal (Y/N)"
               End
            End
         End
      End
      Begin VB.Frame FraShowPic 
         Height          =   5745
         Left            =   180
         TabIndex        =   95
         Top             =   3000
         Width           =   15105
         Begin VB.Frame FraPic1 
            BorderStyle     =   0  'None
            Height          =   4725
            Left            =   200
            TabIndex        =   98
            Top             =   300
            Width           =   4380
            Begin VB.Image Pic1 
               Height          =   4440
               Left            =   75
               Top             =   120
               Width           =   4200
            End
         End
         Begin VB.Frame FraPic2 
            BorderStyle     =   0  'None
            Height          =   4725
            Left            =   5280
            TabIndex        =   97
            Top             =   300
            Width           =   4380
            Begin VB.Image Pic2 
               Height          =   4440
               Left            =   75
               Top             =   120
               Width           =   4200
            End
         End
         Begin VB.Frame FraPic3 
            BorderStyle     =   0  'None
            Height          =   4725
            Left            =   10320
            TabIndex        =   96
            Top             =   300
            Width           =   4380
            Begin VB.Image Pic3 
               Height          =   4440
               Left            =   75
               Top             =   120
               Width           =   4200
            End
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Image 3"
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
            Index           =   67
            Left            =   12000
            TabIndex        =   101
            Top             =   5145
            Width           =   885
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Image 2"
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
            Index           =   66
            Left            =   7035
            TabIndex        =   100
            Top             =   5145
            Width           =   885
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Image 1"
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
            Index           =   7
            Left            =   1845
            TabIndex        =   99
            Top             =   5145
            Width           =   885
         End
      End
   End
End
Attribute VB_Name = "EmrFrmJtZoom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mi_FraKeysHt As Integer
Dim moCn As MwfLib.MDOConnection
Dim ms_TcFr As String, ms_TcTo As String
Dim ms_YyFr As String, ms_YyTo As String
Dim ms_ChrFr As String, ms_ChrTo As String
Dim ms_CoCdFr As String, ms_CoCdTo As String
Dim mi_NoFr As Long, mi_NoTo As Long
Dim mi_SrFr As Long, mi_SrTo As Long
Dim m_OmDtFr As Date, m_OmDtTo As Date
Dim m_DelDtFr As Date, m_DelDtTo As Date
Dim ms_CmCtgFr As String, ms_CmCtgTo As String
Dim ms_CmCdFr As String, ms_CmCdTo As String
Dim mb_PwdFlag As Boolean

Dim ms_OmPoNoFr As String, ms_OmPoNoTo As String
Dim ms_OdPoNoFr As String, ms_OdPoNoTo As String
Dim ms_OdSubCustFr As String, ms_OdSubCustTo As String  'vk.19 sub customer added
Dim ms_DmCtgFr As String, ms_DmCtgTo As String
Dim ms_DmCdFr As String, ms_DmCdTo As String
Dim ms_CdCustDmCdFr As String, ms_CdCustDmCdTo As String
Dim ms_StkCoCdFr As String, ms_StkCoCdTo As String
Dim ms_StkYyFr As String, ms_StkYyTo As String
Dim ms_StkChrFr As String, ms_StkChrTo As String
Dim mi_StkNoFr As Long, mi_StkNoTo As Long
Dim ms_ShPnd As String, ms_ShVal As String
Dim ms_Pwd As String, ms_SortOnCmCd As String

Dim mi_FraPicHskHt As Single
Dim mi_FraPicHskWdt As Single
Dim mi_FraPicHskLft As Single
Dim mi_FraPicHskTop As Single
Dim mi_FraPic3DHt As Single
Dim mi_FraPic3DWdt As Single
Dim mi_FraPic3DLft As Single
Dim mi_FraPic3DTop As Single
Dim mi_PicHskHt As Single
Dim mi_PicHskWdt As Single
Dim mi_Pic3DHt As Single
Dim mi_Pic3DWdt As Single
Dim mb_Pic3DDblClk As Boolean
Dim mb_PicHskDblClk As Boolean
Dim mi_OldGrdOrdMstRw As Single, mi_OldGrdOrdDsgRw As Single
Dim ms_BaseCurCd As String, mb_BaseCur As Boolean, ms_CurCd As String

Enum en_OrdDsgFra
  DsgDet = 0
  DispPic = 1
End Enum

Private Sub ADC_Load()
  '*** Set the First Key Control property to From Cocd
  '*** Set the First Non Key Control property to the grid
  '*** Set the Hot Key for the Grid
  Set moCn = ADC.Connection
  FRA_KEYS.ZOrder (0)
  Set ADC.FirKeyCtl = ADC("wCoCdFr")
  Set ADC.FirNKeyCtl = GrdOrdMst
  GrdOrdMst.HotKey = "A"
  GrdOrdDsg.HotKey = "B"
  GrdBasChild.HotKey = "C"
  GrdRefChild.HotKey = "D"

  If mi_FraKeysHt = 0 Then mi_FraKeysHt = FRA_KEYS.Height
  GrdOrdMst.AllowAdd = False: GrdOrdMst.AllowDelete = False
  GrdOrdDsg.AllowAdd = False: GrdOrdDsg.AllowDelete = False
  GrdBasChild.AllowAdd = False: GrdBasChild.AllowDelete = False
  GrdRefChild.AllowAdd = False: GrdRefChild.AllowDelete = False
  GrdOrdRm.AllowAdd = False: GrdOrdRm.AllowDelete = False
  GrdOrdLab.AllowAdd = False: GrdOrdLab.AllowDelete = False
  'Setting Default values
  m_OmDtFr = "01/01/80"
  m_OmDtTo = "01/01/80"
  m_DelDtFr = "01/01/80"
  m_DelDtTo = "01/01/80"
  ms_CoCdFr = gs_CoCd: ms_CoCdTo = gs_CoCd
  ADC("wDelDt") = "P"
  OptDelDt(1).Value = True
  mb_PwdFlag = moCn.RecSeek("Select PValue1 from Param where PTyp= 'SEOPWD' and PmCd= 'SEOPWD' and PValue1= '' ")
  
  '*** Storing The Original Size And Position Of The Picture Controls ***
  mi_FraPicHskHt = FraPicHsk.Height
  mi_FraPicHskWdt = FraPicHsk.Width
  mi_FraPicHskLft = FraPicHsk.Left
  mi_FraPicHskTop = FraPicHsk.Top
  mi_FraPic3DHt = FraPic3D.Height
  mi_FraPic3DWdt = FraPic3D.Width
  mi_FraPic3DLft = FraPic3D.Left
  mi_FraPic3DTop = FraPic3D.Top

  mi_PicHskHt = PicHsk.Height
  mi_PicHskWdt = PicHsk.Width
  mi_Pic3DHt = Pic3D.Height
  mi_Pic3DWdt = Pic3D.Width
  mb_Pic3DDblClk = False
  mb_PicHskDblClk = False

End Sub
Private Sub ADC_KeyWhen()
  
  FraNKeyAll.ZOrder (0)
  GrdOrdMst.Rows = 1
  GrdOrdDsg.Rows = 1
  GrdBasChild.Rows = 1
  GrdRefChild.Rows = 1
  
  FRA_KEYS.Height = mi_FraKeysHt
  CmdDsgDet.Enabled = False
  CmdShowPic.Enabled = False
    
  ADC("wCoCdFr") = ms_CoCdFr
  ADC("wCoCdTo") = ms_CoCdTo
  ADC("wOmTcFr") = ms_TcFr
  ADC("wOmTcTo") = ms_TcTo
  ADC("wOmYyFr") = ms_YyFr
  ADC("wOmYyTo") = ms_YyTo
  ADC("wOmChrFr") = ms_ChrFr
  ADC("wOmChrTo") = ms_ChrTo
  ADC("wOmNoFr") = mi_NoFr
  ADC("wOmNoTo") = mi_NoTo
  ADC("wOdSrFr") = mi_SrFr
  ADC("wOdSrTo") = mi_SrTo
  ADC("wOmDtFr") = m_OmDtFr
  ADC("wOmDtTo") = m_OmDtTo
  ADC("wOdDelDtFr") = m_DelDtFr
  ADC("wOdDelDtTo") = m_DelDtTo

  ADC("wOmCmCtgFr") = ms_CmCtgFr
  ADC("wOmCmCtgTo") = ms_CmCtgTo
  ADC("wOmCmCdFr") = ms_CmCdFr
  ADC("wOmCmCdTo") = ms_CmCdTo
  ADC("wOmPoNoFr") = ms_OmPoNoFr
  ADC("wOmPoNoTo") = ms_OmPoNoTo
  ADC("wOdPoNoFr") = ms_OdPoNoFr
  ADC("wOdPoNoTo") = ms_OdPoNoTo
  'vk.19 sub customer added
  ADC("wOdSubCustFr") = ms_OdSubCustFr
  ADC("wOdSubCustTo") = ms_OdSubCustTo
  
  ADC("wDsgCtgFr") = ms_DmCtgFr
  ADC("wDsgCtgTo") = ms_DmCtgTo
  ADC("wOdDmCdFr") = ms_DmCdFr
  ADC("wOdDmCdTo") = ms_DmCdTo
  ADC("wCdCustDmCdFr") = ms_CdCustDmCdFr
  ADC("wCdCustDmCdTo") = ms_CdCustDmCdTo

  ADC("wOdDmStkCoCdFr") = ms_StkCoCdFr
  ADC("wOdDmStkCoCdTo") = ms_StkCoCdTo
  ADC("wOdDmStkYyFr") = ms_StkYyFr
  ADC("wOdDmStkYyTo") = ms_StkYyTo
  ADC("wOdDmStkChrFr") = ms_StkChrFr
  ADC("wOdDmStkChrTo") = ms_StkChrTo
  ADC("wOdDmStkNoFr") = mi_StkNoFr
  ADC("wOdDmStkNoTo") = mi_StkNoTo
    
  ADC("wPndYn") = IIF(ms_ShPnd <> "", ms_ShPnd, "Y")
  ADC("wValYn") = IIF(ms_ShVal <> "", ms_ShVal, "N")
  TxtPwd.Text = IIF(ms_Pwd <> "", ms_Pwd, "")
  ADC("wSortCmCd") = IIF(ms_SortOnCmCd <> "", ms_SortOnCmCd, "N")
  ADC("wDelDt") = IIF(OptDelDt(0).Value = True, "E", "P")
  ADC("wCurCd") = ms_CurCd    '578.22
  
  If ADC("wPndYn") = "Y" Then ChkBoxArr(0).Value = Checked Else ChkBoxArr(0).Value = Unchecked
  If ADC("wValYn") = "Y" Then ChkBoxArr(1).Value = Checked Else ChkBoxArr(1).Value = Unchecked
  If ADC("wSortCmCd") = "Y" Then ChkBoxArr(2).Value = Checked Else ChkBoxArr(2).Value = Unchecked
    
  Call ShowPic(0, 0, "", PicHsk)
  Call ShowPic(0, 0, "", Pic3D)
  Call ShowPic(0, 0, "", Pic1)
  Call ShowPic(0, 0, "", Pic2)
  Call ShowPic(0, 0, "", Pic3)

  FraPicHsk.Height = mi_FraPicHskHt
  FraPicHsk.Width = mi_FraPicHskWdt
  FraPicHsk.Left = mi_FraPicHskLft
  FraPicHsk.Top = mi_FraPicHskTop
  FraPic3D.Height = mi_FraPic3DHt
  FraPic3D.Width = mi_FraPic3DWdt
  FraPic3D.Left = mi_FraPic3DLft
  FraPic3D.Top = mi_FraPic3DTop

  mi_OldGrdOrdMstRw = 0: mi_OldGrdOrdDsgRw = 0
  
End Sub
Private Sub ADC_KeyValid(pr_Cancel As Boolean, pr_ErrMsg As String)
  'Checking user level access  to view export delivery date
  If GetUsrLevel > 5 And ADC("wDelDt") = "E" Then
    pr_Cancel = True: pr_ErrMsg = "User don't have access to view Export Delivery Date": Exit Sub
  End If
  If ADC("wDelDt") = "E" Then
    GrdOrdMst.ColProp("wOmPrdExpDelDt").Heading = "Exp DelDt"
  Else
    GrdOrdMst.ColProp("wOmPrdExpDelDt").Heading = "Prd DelDt"
  End If
  If ADC("wDelDt") = "E" Then
    GrdOrdDsg.ColProp("wOdPrdExpDelDt").Heading = "Exp DelDt"
  Else
    GrdOrdDsg.ColProp("wOdPrdExpDelDt").Heading = "Prd DelDt"
  End If

  'Checking valid company code and Tc
  If ADC("wCoCdFr") = "" Or ADC("wCoCdTo") = "" Then pr_Cancel = True: pr_ErrMsg = "Company Codes Cannot Be Blank": Exit Sub
  If ADC("wOmTcFr") = "" Or ADC("wOmTcTo") = "" Then pr_Cancel = True: pr_ErrMsg = "Tcs Cannot Be Blank": Exit Sub
  
  'Sr no cannot enter for different orders
  If (ADC("wOdSrFr") <> 0 Or ADC("wOdSrTo") <> 0) And _
     ((ADC("wOmTcFr") = "" Or ADC("wOmYyFr") = "" Or ADC("wOmChrFr") = "" Or ADC("wOmNoFr") = 0) Or _
       (ADC("wOmTcFr") <> ADC("wOmTcTo") Or ADC("wOmYyFr") <> ADC("wOmYyTo") Or ADC("wOmChrFr") <> ADC("wOmChrTo") Or ADC("wOmNoFr") <> ADC("wOmNoTo")) _
     ) Then pr_Cancel = True: pr_ErrMsg = "Design Sr Can Be Specified Only For Single Order": Exit Sub

  '578.22 if currency code not given it takes base currency from head else takes from scope
  ms_BaseCurCd = moCn.GetFldVal("Select HBaseCurCd from Head where HCd= '" + ctSelfCmCd + "' and " + _
                  "HCoCd= '" + ADC("wCurCd") + "'")
  mb_BaseCur = IIF(ADC("wCurCd") = "", True, False)
  
End Sub

Private Sub ADC_NonKeyWhen(pr_Cancel As Boolean, pr_ErrMsg As String)
  FRA_KEYS.Height = 0
  'FRA_NKEYS.ZOrder (0)
  'ADC.ZOrder (0)
  ms_CoCdFr = ADC("wCoCdFr")
  ms_CoCdTo = ADC("wCoCdTo")
  ms_TcFr = ADC("wOmTcFr")
  ms_TcTo = ADC("wOmTcTo")
  ms_YyFr = ADC("wOmYyFr")
  ms_YyTo = ADC("wOmYyTo")
  ms_ChrFr = ADC("wOmChrFr")
  ms_ChrTo = ADC("wOmChrTo")
  mi_NoFr = ADC("wOmNoFr")
  mi_NoTo = ADC("wOmNoTo")
  mi_SrFr = ADC("wOdSrFr")
  mi_SrTo = ADC("wOdSrTo")

  m_OmDtFr = IIF(ADC("wOmDtFr") <> "", ADC("wOmDtFr"), "01/01/80")
  m_OmDtTo = IIF(ADC("wOmDtTo") <> "", ADC("wOmDtFr"), "01/01/80")
  m_DelDtFr = IIF(ADC("wOdDelDtFr") <> "", ADC("wOdDelDtFr"), "01/01/80")
  m_DelDtTo = IIF(ADC("wOdDelDtTo") <> "", ADC("wOdDelDtFr"), "01/01/80")
  ms_CmCtgFr = ADC("wOmCmCtgFr")
  ms_CmCtgTo = ADC("wOmCmCtgTo")
  ms_CmCdFr = ADC("wOmCmCdFr")
  ms_CmCdTo = ADC("wOmCmCdTo")

  ms_OmPoNoFr = ADC("wOmPoNoFr")
  ms_OmPoNoTo = ADC("wOmPoNoTo")
  ms_OdPoNoFr = ADC("wOdPoNoFr")
  ms_OdPoNoTo = ADC("wOdPoNoTo")
  'vk.19 sub customer added
  ms_OdSubCustFr = ADC("wOdSubCustFr")
  ms_OdSubCustTo = ADC("wOdSubCustTo")
  
  ms_DmCtgFr = ADC("wDsgCtgFr")
  ms_DmCtgTo = ADC("wDsgCtgTo")
  ms_DmCdFr = ADC("wOdDmCdFr")
  ms_DmCdTo = ADC("wOdDmCdTo")
  ms_CdCustDmCdFr = ADC("wCdCustDmCdFr")
  ms_CdCustDmCdTo = ADC("wCdCustDmCdTo")

  ms_StkCoCdFr = ADC("wOdDmStkCoCdFr")
  ms_StkCoCdTo = ADC("wOdDmStkCoCdTo")
  ms_StkYyFr = ADC("wOdDmStkYyFr")
  ms_StkYyTo = ADC("wOdDmStkYyTo")
  ms_StkChrFr = ADC("wOdDmStkChrFr")
  ms_StkChrTo = ADC("wOdDmStkChrTo")
  mi_StkNoFr = ADC("wOdDmStkNoFr")
  mi_StkNoTo = ADC("wOdDmStkNoTo")
    
  ms_ShPnd = ADC("wPndYn")
  ms_ShVal = ADC("wValYn")
  ms_Pwd = TxtPwd.Text
  ms_SortOnCmCd = ADC("wSortCmCd")
  ms_CurCd = ADC("wCurCd")    '578.22

  Call FillGrdOrdMst

  If GrdOrdMst.Rows > 1 Then
    GrdOrdMst.Row = 1: GrdOrdMst.SetFocus
    CmdDsgDet.Enabled = True
    CmdShowPic.Enabled = True
  Else
    GrdOrdDsg.Rows = 1: GrdBasChild.Rows = 1: GrdRefChild.Rows = 1
  End If
  Call InitGrdBasRefchild
End Sub
Private Sub InitGrdBasRefchild()
  Call cmdBasChild_Click
End Sub
Private Sub ADC_SetRecSource()
  '*** Since this is an 'A' Entry, there is no RecSource associated with ADC. So set the
  'DataMode property of ADC control as xdatUnbound
  ADC.DataMode = xdatUnbound

End Sub
Private Sub cmdRefChild_Click()
  GrdRefChild.ZOrder (0)
  cmdRefChild.Enabled = False
  cmdBasChild.Enabled = True
'  GrdBasChild.ZOrder = 1
End Sub

Private Sub cmdBasChild_Click()
'  GrdRefChild.ZOrder = 1
  GrdBasChild.ZOrder (0)
  cmdBasChild.Enabled = False
  cmdRefChild.Enabled = True
End Sub

Private Sub CmdShowPic_Click()
  Call DispFra(DispPic)
  
  'This will show 3 images of a selected design code
  If FraShowPic.Visible = True Then
    Dim w3DPth As String
    w3DPth = GetPictPath("DM", False, GrdOrdDsg.Value(GrdOrdDsg.Row, "wOdDsgCd"), GrdOrdDsg.Value(GrdOrdDsg.Row, "wOdSz"), , , , , , , , "1")
    Call ShowPic(4440, 4200, w3DPth, Pic1)
    w3DPth = GetPictPath("DM", False, GrdOrdDsg.Value(GrdOrdDsg.Row, "wOdDsgCd"), GrdOrdDsg.Value(GrdOrdDsg.Row, "wOdSz"), , , , , , , , "2")
    Call ShowPic(4440, 4200, w3DPth, Pic2)
    w3DPth = GetPictPath("DM", False, GrdOrdDsg.Value(GrdOrdDsg.Row, "wOdDsgCd"), GrdOrdDsg.Value(GrdOrdDsg.Row, "wOdSz"), , , , , , , , "3")
    Call ShowPic(4440, 4200, w3DPth, Pic3)
        
  End If

End Sub

Private Sub GrdBasChild_ColWhen(ByVal RowNum As Integer, ByVal ColName As String, Cancel As Boolean, ErrMsg As String, ByVal Action As MwfCtl.en_fgAction)
  Cancel = True
End Sub

Private Sub GrdBasChild_Validate(Cancel As Boolean)
  Cancel = GrdBasChild.Validate
End Sub

Private Sub GrdOrdDsg_ColWhen(ByVal RowNum As Integer, ByVal ColName As String, Cancel As Boolean, ErrMsg As String, ByVal Action As MwfCtl.en_fgAction)
  Cancel = True
End Sub

Private Sub GrdOrdDsg_RowWhen(ByVal RowNum As Integer)
  Dim wHskPth As String, w3DPth As String
  'Calling child records
  If RowNum > 0 Then
    Call FillFooterFields(RowNum)
    Call FillGrdBasChild(RowNum)
    Call FillGrdRefChild(RowNum)
  End If
  'Calling Picture design for the row
  FraPicHsk.Height = mi_FraPicHskHt: FraPicHsk.Width = mi_FraPicHskWdt
  FraPicHsk.Left = mi_FraPicHskLft: FraPicHsk.Top = mi_FraPicHskTop
  
  wHskPth = GetPictPath("DM", True, GrdOrdDsg.Value(RowNum, "wOdDsgCd"), GrdOrdDsg.Value(RowNum, "wOdSz"), , , , , , , GrdOrdDsg.Value(RowNum, "wOdPicNm"))
  w3DPth = GetPictPath("DM", False, GrdOrdDsg.Value(RowNum, "wOdDsgCd"), GrdOrdDsg.Value(RowNum, "wOdSz"), , , , , , , GrdOrdDsg.Value(RowNum, "wOdPicNm"))
  Call ShowPic(mi_PicHskHt, mi_PicHskWdt, wHskPth, PicHsk)
  Call ShowPic(mi_Pic3DHt, mi_Pic3DWdt, w3DPth, Pic3D)

End Sub


Private Sub GrdOrdLab_ColWhen(ByVal RowNum As Integer, ByVal ColName As String, Cancel As Boolean, ErrMsg As String, ByVal Action As MwfCtl.en_fgAction)
  Cancel = True
End Sub

Private Sub GrdOrdMst_ColWhen(ByVal RowNum As Integer, ByVal ColName As String, Cancel As Boolean, ErrMsg As String, ByVal Action As MwfCtl.en_fgAction)
  Cancel = True
End Sub

Private Sub GrdOrdMst_RowWhen(ByVal RowNum As Integer)
  If RowNum > 0 Then Call FillGrdOrdDsg(RowNum)
  
End Sub

Private Sub GrdOrdRm_ColWhen(ByVal RowNum As Integer, ByVal ColName As String, Cancel As Boolean, ErrMsg As String, ByVal Action As MwfCtl.en_fgAction)
  Cancel = True
End Sub

Private Sub GrdRefChild_ColWhen(ByVal RowNum As Integer, ByVal ColName As String, Cancel As Boolean, ErrMsg As String, ByVal Action As MwfCtl.en_fgAction)
  Cancel = True
End Sub

Private Sub GrdRefChild_Validate(Cancel As Boolean)
  Cancel = GrdRefChild.Validate
End Sub

Private Sub OptDelDt_Click(Index As Integer)
  'If export deldate is selected it store E or P
  Select Case Index
  Case Is = 0
    ADC("wDelDt") = "E"
  Case Is = 1
    ADC("wDelDt") = "P"
  End Select
End Sub

Private Sub ATXT_FldChange(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String, ByVal pv_NewValue As Variant)
 'Stores From text box values to To textbox
 Select Case UCase(IdName)
  Case Is = UCase("wOmTcFr")
    ADC("wOmTcTo") = pv_NewValue
  Case Is = UCase("wOmYyFr")
    ADC("wOmYyTo") = pv_NewValue
  Case Is = UCase("wOmChrFr")
    ADC("wOmChrTo") = pv_NewValue
  Case Is = UCase("wOmNoFr")
    ADC("wOmNoTo") = pv_NewValue
  Case Is = UCase("wOdSrFr")
    ADC("wOdSrTo") = pv_NewValue
  Case Is = UCase("wOmDtFr")
    ADC("wOmDtTo") = pv_NewValue
  Case Is = UCase("wOdDelDtFr")
    ADC("wOdDelDtTo") = pv_NewValue
  Case Is = UCase("wOmCmCtgFr")
    ADC("wOmCmCtgTo") = pv_NewValue
  Case Is = UCase("wOmCmCdFr")
    ADC("wOmCmCdTo") = pv_NewValue
  Case Is = UCase("wOmCmCdTo")
    'Customer Design Code should be empty for different Customer Code
    If UCase(Trim(ADC("wOmCmCdFr"))) <> UCase(Trim(pv_NewValue)) Then
      ADC("wCdCustDmCdFr") = ""
      ADC("wCdCustDmCdTo") = ""
    End If
  Case Is = UCase("wOmPoNoFr")
    ADC("wOmPoNoTo") = pv_NewValue
  Case Is = UCase("wOdPoNoFr")
    ADC("wOdPoNoTo") = pv_NewValue
  'vk.19 sub customer added
  Case Is = UCase("wOdSubCustFr")
    ADC("wOdSubCustTo") = pv_NewValue
  Case Is = UCase("wDsgCtgFr")
    ADC("wDsgCtgTo") = pv_NewValue
  
      
  Case Is = UCase("wOdDmCdFr")
    ADC("wOdDmCdTo") = pv_NewValue
  Case Is = UCase("wCdCustDmCdFr")
    ADC("wCdCustDmCdTo") = pv_NewValue
    
  Case Is = UCase("wOdDmStkCoCdFr")
    ADC("wOdDmStkCoCdTo") = pv_NewValue
  Case Is = UCase("wOdDmStkYyFr")
    ADC("wOdDmStkYyTo") = pv_NewValue
  Case Is = UCase("wOdDmStkChrFr")
    ADC("wOdDmStkChrTo") = pv_NewValue
  Case Is = UCase("wOdDmStkNoFr")
    ADC("wOdDmStkNoTo") = pv_NewValue
  End Select
  
End Sub

Private Sub ATXT_FldWhen(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String)
  'Calling help for fields
  Select Case UCase(IdName)
  Case Is = UCase("wCoCdFr"), UCase("wCoCdTo")
    Call HlpList.hCoCd
  Case Is = UCase("wOmTcFr"), UCase("wOmTcTo")
    Call HlpList.PMCd("TC", "'JT'")
  Case Is = UCase("wOmYyFr"), UCase("wOmYyTo")
    Call HlpList.PMCd("YY")
  Case Is = UCase("wOmChrFr")
    Call HlpList.vPSCd(ADC("wCoCdFr"), "CHR", ADC("wOmTcFr"))
  Case Is = UCase("wOmChrTo")
    Call HlpList.vPSCd(ADC("wCoCdTo"), "CHR", ADC("wOmTcTo"))
  Case Is = UCase("wOmNoFr")
    Call HlpList.OmNo(ADC("wCoCdFr"), ADC("wOmTcFr"), ADC("wOmYyFr"), ADC("wOmChrFr"))
  Case Is = UCase("wOmNoTo")
    Call HlpList.OmNo(ADC("wCoCdTo"), ADC("wOmTcTo"), ADC("wOmYyTo"), ADC("wOmChrTo"))
  Case Is = UCase("wOdSrFr")
    Call HlpList.OdSr(ADC("wCoCdFr"), ADC("wOmTcFr"), ADC("wOmYyFr"), ADC("wOmChrFr"), ADC("wOmNoFr"))
  Case Is = UCase("wOdSrTo")
    Call HlpList.OdSr(ADC("wCoCdTo"), ADC("wOmTcTo"), ADC("wOmYyTo"), ADC("wOmChrTo"), ADC("wOmNoTo"))
  Case Is = UCase("wOmCmCtgFr"), UCase("wOmCmCtgTo")
    Call HlpList.PMCd("CMCTG")
  Case Is = UCase("wOmCmCdFr")
    Call HlpList.CustCd(ADC("wOmCmCtgFr"))
  Case Is = UCase("wOmCmCdTo")
    Call HlpList.CustCd(ADC("wOmCmCtgTo"))
  Case Is = UCase("wDsgCtgFr"), UCase("wDsgCtgTo")
    Call HlpList.PMCd("DmCtg")
  Case Is = UCase("wOdDmCdFr"), UCase("wOdDmCdTo")
    Call HlpList.DmCd("DM")
  Case Is = UCase("wCdCustDmCdFr"), UCase("wCdCustDmCdTo")
    If ADC("wOmCmCdFr") <> ADC("wOmCmCdTo") Then Cancel = True: ErrMsg = "Cannot Enter Customer Design Code as From & To Customer code not same": Exit Sub
    Call HlpList.CustDmCd(ADC("wOmCmCdFr"))
  Case Is = UCase("wOdDmStkCoCdFr"), UCase("wOdDmStkCoCdTo")
    Call HlpList.hCoCd
  Case Is = UCase("wOdDmStkYyFr"), UCase("wOdDmStkYyTo")
    HlpList.PMCd "YY"
  Case Is = UCase("wOdDmStkChrFr"), UCase("wOdDmStkChrTo")
    HlpList.PMCd "BCHR"
  Case Is = UCase("wOdDmStkNoFr")
    HlpList.OdDmStkNo ADC("wCoCdFr"), ADC("wOdDmCdFr"), "", "", ADC("wOdDmStkCoCdFr"), ADC("wOdDmStkYyFr"), ADC("wOdDmStkChrFr")
  Case Is = UCase("wOdDmStkNoTo")
    HlpList.OdDmStkNo ADC("wCoCdTo"), ADC("wOdDmCdTo"), "", "", ADC("wOdDmStkCoCdTo"), ADC("wOdDmStkYyTo"), ADC("wOdDmStkChrTo")
  'vk.19 sub customer added
  Case Is = UCase("wOdSubCustFr"), UCase("wOdSubCustTo")
    HlpList.PMCd "SUBCUST"
  '578.22 calling help for currency
  Case Is = UCase("wCurCd")
    HlpList.PMCd "Curncy"
  End Select
  
    
End Sub

Private Sub ChkBoxArr_Click(Index As Integer)
  With ChkBoxArr(Index)
    Select Case Index
    'Show Pending
    Case Is = 0
      If .Value = Checked Then ADC("wPndYn") = "Y" Else ADC("wPndYn") = "N"
    'Show values
    Case Is = 1
      If .Value = Checked Then
        If mb_PwdFlag = False Then
          ChkBoxArr(1).Value = Unchecked: ADC("wValYn") = "N": DispMsg "Enter SEO Password To Show Value", etError
        Else
          ADC("wValYn") = "Y"
        End If
      Else
        ADC("wValYn") = "N"
      End If
    'Order by customer code
    Case Is = 2
      If .Value = Checked Then ADC("wSortCmCd") = "Y" Else ADC("wSortCmCd") = "N"
    End Select
  End With
End Sub

Private Sub ChkBoxArr_GotFocus(Index As Integer)
  DispMsg ChkBoxArr(Index).ToolTipText, etInfo
End Sub
Private Sub CmdVchDet_Click()

End Sub
Private Sub CmdDsgDet_Click()
  Call DispFra(DsgDet)
  'when click DegDet button will call OrdRm , Lab details. if grid A or B row change will call rm and lab details
  If mi_OldGrdOrdMstRw <> GrdOrdMst.Row Or mi_OldGrdOrdDsgRw <> GrdOrdDsg.Row Then
    Call FillOrdRm(GrdOrdDsg.Row)
    Call FillOrdLab(GrdOrdDsg.Row)
    mi_OldGrdOrdMstRw = GrdOrdMst.Row: mi_OldGrdOrdDsgRw = GrdOrdDsg.Row
  End If
End Sub

Public Sub FillGrdOrdMst()
  Dim wCnd As String, wrepcnd As String, wsql As String
  Dim wo_rsOrdMst As MwfLib.MDORowSet, wi_Row As Integer
  Dim ws_OrdBy As String
  
  wrepcnd = ADC.RepCond
  'If Order by selected then include CmCtg, Cmcd
  If ADC("wSortCmCd") = "Y" Then
    ws_OrdBy = " Order By OmCmCtg, OmCmCd, OmCoCd, OmTc, OmYy, OmChr, OmNo "
  Else
    ws_OrdBy = " Order By OmCoCd, OmTc, OmYy, OmChr, OmNo "
  End If
  'condition for pending qty
  If ADC("wPndYn") = "Y" Then
    wCnd = "Where convert(decimal (14,1), OdOrdQty)<> convert(decimal (14,1), OdExpQty) " + IIF(wrepcnd <> "", " and " + wrepcnd, "")
  ElseIf ADC("wPndYn") = "N" Then
    wCnd = "Where " + IIF(wrepcnd <> "", wrepcnd, "")
  End If
  
  wsql = "Select OmCoCd, OmTc, OmYy, OmChr, OmNo, Max(OmExpDelDt) As OmExpDelDt,Max(OmDelDt) As OmDelDt, " + _
         "Max(OmCmCtg) As OmCmCtg, Max(OmCmCd) As OmCmCd, Max(OmPoNo) As OmPoNo, Max(OmDt) As OmDt " + _
         "From OrdMst " + _
         "Join OrdDsg on OmIdNo=OdOmIdNo And OmCoCd= OdCoCd and OmTc= OdTc and OmYy= OdYy and OmChr= OdChr and OmNo= OdNo " + IIF(gs_Partition = ctCurrPrtn, " And OmPrtKey=OdPrtKey ", "") + _
         "Join DsgMst on DmIdNo=OdDmIdNo And DmTcTyp='DM' and DmCd = OdDmCd and DmSz = '' " + IIF(gs_Partition = ctCurrPrtn, " and DmPrtKey='" + ctCurrPrtn + "' ", "") + _
         "Left Outer Join CustDsg on CdCtg='C' and CdCd=OmCmCd and CdOurDmCd =OdDmCd and CdOurSfx='' and CdOurDmSz='' " + _
         wCnd + IIF(gs_Partition = ctCurrPrtn, " and OmPrtKey='" + ctCurrPrtn + "' ", "") + _
         " Group By OmCoCd, OmTc, OmYy, OmChr, OmNo " + ws_OrdBy
  
  GrdOrdMst.Rows = 1
  Set wo_rsOrdMst = moCn.OpenRes(wsql)
  If wo_rsOrdMst.RecCount <= 2500 Then
    GrdOrdMst.AllowAdd = True
    With GrdOrdMst
      Do While Not (wo_rsOrdMst.EOF Or wo_rsOrdMst.BOF)
        .AddItem
        wi_Row = .Rows - 1
        .Value(wi_Row, "wOmCoCd") = Trim(wo_rsOrdMst!OmCoCd)
        .Value(wi_Row, "wOmTc") = Trim(wo_rsOrdMst!OmTc)
        .Value(wi_Row, "wOmYy") = Trim(wo_rsOrdMst!OmYy)
        .Value(wi_Row, "wOmChr") = Trim(wo_rsOrdMst!OmChr)
        .Value(wi_Row, "wOmNo") = Trim(wo_rsOrdMst!OmNo)
        .Value(wi_Row, "wOrdNo") = Trim(wo_rsOrdMst!OmCoCd) + "/" + Trim(wo_rsOrdMst!OmTc) + "/" + Trim(wo_rsOrdMst!OmYy) + "/" + Trim(wo_rsOrdMst!OmChr) + "/" + CStr(Trim(wo_rsOrdMst!OmNo))
        .Value(wi_Row, "wOmPrdExpDelDt") = IIF(ADC("wDelDt") = "E", wo_rsOrdMst!OmExpDelDt, wo_rsOrdMst!OmDelDt)
        .Value(wi_Row, "wOmCust") = Trim(wo_rsOrdMst!OmCmCd)
        .Value(wi_Row, "wOmPoNo") = Trim(wo_rsOrdMst!OmPoNo)
        .Value(wi_Row, "wOmPoDt") = Trim(wo_rsOrdMst!OmDt)  '5716 PO date changed as voucher date
        wo_rsOrdMst.MoveNext
      Loop
    End With
  Else
    DispMsg "No. of Records are more than 2500, Plz Limit the Scope", etError
  End If
  Set wo_rsOrdMst = Nothing
  GrdOrdMst.AllowAdd = False
  'calling grid B details
  If GrdOrdMst.Rows > 1 Then
    GrdOrdMst.Row = 1
    Call FillGrdOrdDsg(1)
  End If

End Sub

Public Sub FillGrdOrdDsg(ByVal RowNum As Integer)
  Dim wCnd As String, wrepcnd As String, wsql As String
  Dim wo_rsOrdDsg As MwfLib.MDORowSet, wi_Row As Integer
    
  wrepcnd = ADC.RepCond
  'condition for pending qty
  If ADC("wPndYn") = "Y" Then
    wCnd = " and convert(decimal (14,1), OdOrdQty)<> convert(decimal (14,1), OdExpQty) " + IIF(wrepcnd <> "", " and " + wrepcnd, "")
  ElseIf ADC("wPndYn") = "N" Then
    wCnd = IIF(wrepcnd <> "", " and " + wrepcnd, "")
  End If
  '578.22 conversion factor added to OdSalPrc
  wsql = "Select OdCoCd, OdTc, OdYy, OdChr, OdNo, OdSr, max(OdDmCd) as OdDmCd, max(OdSfx) as OdSfx, max(OdDmSz) as OdDmSz ," + _
         "max(OdDmStkCoCd) as OdDmStkCoCd, max(OdDmStkYy) as OdDmStkYy, max(OdDmStkChr) as OdDmStkChr, max(OdDmStkNo) as OdDmStkNo, " + _
         "max(OdOrdQty) as OdOrdQty, max(OdExpDelDt) as OdExpDelDt, max(OdDelDt) as OdDelDt, max(" + IIF(mb_BaseCur, "(Case when OmMulDiv= 'M' then OdSalPrc/OmCnvFct else OdSalPrc*OmCnvFct end) ", "OdSalPrc") + ") as OdSalPrc, max(OdOrdQty-OdExpQty) as qBalQty, max(OdSubCust) as OdSubCust, max(OdPoNo) as OdPoNo, max(OdPicNm) as OdPicNm, " + _
         "isnull((select OdB.OdCoCd+'/'+OdB.OdTc+'/'+OdB.OdYy+'/'+OdB.OdChr+'/'+Ltrim(Str(OdB.OdNo))+'/'+Ltrim(Str(OdB.OdSr)) from OrdDsg OdB where OdB.OdIdNo=max(OrdDsg.OdBaseIdKey)),'') as qBaseVch, " + _
         "isnull((select OdR.OdCoCd+'/'+OdR.OdTc+'/'+OdR.OdYy+'/'+OdR.OdChr+'/'+Ltrim(Str(OdR.OdNo))+'/'+Ltrim(Str(OdR.OdSr)) from OrdDsg OdR where OdR.OdIdNo=max(OrdDsg.OdRefIdKey)),'') as qRefVch, " + _
         " " + _
         "sum(case when IsNull(OrdRm.OrRmCtg,'') in ('D','C') then IsNull(OrdRm.OrWt,0)/5 else IsNull(OrdRm.OrWt,0) end) as qGrsWt," + _
         "sum(case when IsNull(OrdRm.OrRmCtg,'') in ('G','P','S','L','M') and (charindex(','+OrdRm.OrRmSCtg+',',','+hChnSubCtgForGld+',')=0) then IsNull(OrdRm.OrWt,0) else 0 end) as qMetWt ," + _
         "sum(case when IsNull(OrdRm.OrRmCtg,'') ='D' then IsNull(OrdRm.OrWt,0) else 0 end) as qDiaWt," + _
         "sum(case when IsNull(OrdRm.OrRmCtg,'') ='C' then IsNull(OrdRm.OrWt,0) else 0 end) as qCsWt," + _
         "sum(case when IsNull(OrdRm.OrRmCtg,'') in ('X') then IsNull(OrdRm.OrWt,0) else 0 end) as qXWt ," + _
         "sum(case when " + _
         "(charindex(','+OrdRm.OrRmSCtg+',',','+hChnSubCtgForGld+',')<>0) " + _
         "then (case when IsNull(OrdRm.OrRmCtg,'') in ('D','C') then IsNull(OrdRm.OrWt,0)/5 else IsNull(OrdRm.OrWt,0) end) " + _
         "else 0 end ) as qCHNWt, max(OdIdNo) as OdIdNo " + _
         " From OrdMst join OrdDsg on OmIdNo=OdOmIdNo And OmCoCd= OdCoCd and OmTc= OdTc and OmYy= OdYy and OmChr= OdChr and OmNo= OdNo " + IIF(gs_Partition = ctCurrPrtn, " And OmPrtKey=OdPrtKey ", "") + _
         " Join DsgMst on DmIdNo=OdDmIdNo And DmTcTyp='DM' and DmCd = OdDmCd and DmSz = '' " + IIF(gs_Partition = ctCurrPrtn, " And DmPrtKey=OdPrtKey ", "") + _
         " Left Outer Join CustDsg on CdCtg='C' and CdCd=OmCmCd and CdOurDmCd =OdDmCd and CdOurSfx='' and CdOurDmSz='' " + _
         " Join OrdRm on OrCoCd=OdCoCd and OrTc=OdTc and OrYy=OdYy and OrChr=OdChr and OrNo=OdNo and OrSr=OdSr " + IIF(gs_Partition = ctCurrPrtn, " And OrPrtKey=OdPrtKey ", "") + _
         " Join Head on HCoCd =OdCoCd "
         
  wsql = wsql + _
         " Where " + IIF(gs_Partition = ctCurrPrtn, " OdPrtKey='" + ctCurrPrtn + "' And ", "") + " OdCoCd='" + GrdOrdMst.Value(RowNum, "wOmCoCd") + "' And OdTc='" + GrdOrdMst.Value(RowNum, "wOmTc") + "' " + _
         " And OdYy='" + GrdOrdMst.Value(RowNum, "wOmYy") + "' And OdChr='" + GrdOrdMst.Value(RowNum, "wOmChr") + "' And OdNo=" + CStr(GrdOrdMst.Value(RowNum, "wOmNo")) + _
         wCnd + " group by OdCoCd, OdTc, OdYy, OdChr, OdNo, OdSr Order By OdCoCd, OdTc, OdYy, OdChr, OdNo, OdSr "
                                      
  GrdOrdDsg.Rows = 1
  Set wo_rsOrdDsg = moCn.OpenRes(wsql)
    GrdOrdDsg.AllowAdd = True
    With GrdOrdDsg
      Do While Not (wo_rsOrdDsg.EOF Or wo_rsOrdDsg.BOF)
        .AddItem
        wi_Row = .Rows - 1
        .Value(wi_Row, "wOdCoCd") = Trim(wo_rsOrdDsg!OdCoCd)
        .Value(wi_Row, "wOdTc") = Trim(wo_rsOrdDsg!OdTc)
        .Value(wi_Row, "wOdYy") = Trim(wo_rsOrdDsg!OdYy)
        .Value(wi_Row, "wOdChr") = Trim(wo_rsOrdDsg!OdChr)
        .Value(wi_Row, "wOdNo") = Trim(wo_rsOrdDsg!OdNo)
        .Value(wi_Row, "wOdSr") = Trim(wo_rsOrdDsg!OdSr)
        .Value(wi_Row, "wOdDsgCd") = Trim(wo_rsOrdDsg!OdDmCd)
        .Value(wi_Row, "wOdSfx") = Trim(wo_rsOrdDsg!OdSfx)
        .Value(wi_Row, "wOdSz") = Trim(wo_rsOrdDsg!OdDmSz)
        .Value(wi_Row, "wOdStkCoCd") = Trim(wo_rsOrdDsg!OdDmStkCoCd)
        .Value(wi_Row, "wOdStkYy") = Trim(wo_rsOrdDsg!OdDmStkYy)
        .Value(wi_Row, "wOdStkChr") = Trim(wo_rsOrdDsg!OdDmStkChr)
        .Value(wi_Row, "wOdStkNo") = Trim(wo_rsOrdDsg!OdDmStkNo)
        .Value(wi_Row, "wOdPrdExpDelDt") = IIF(ADC("wDelDt") = "E", wo_rsOrdDsg!OdExpDelDt, wo_rsOrdDsg!OdDelDt)
        If Trim(wo_rsOrdDsg!OdDmStkCoCd) = "" And Trim(wo_rsOrdDsg!OdDmStkYy) = "" And Trim(wo_rsOrdDsg!OdDmStkChr) = "" And Trim(wo_rsOrdDsg!OdDmStkNo) = 0 Then
          .Value(wi_Row, "wOdStkId") = ""
        Else
          .Value(wi_Row, "wOdStkId") = Trim(wo_rsOrdDsg!OdDmStkCoCd) + "/" + Trim(wo_rsOrdDsg!OdDmStkYy) + "/" + Trim(wo_rsOrdDsg!OdDmStkChr) + "/" + CStr(Trim(wo_rsOrdDsg!OdDmStkNo))
        End If
        .Value(wi_Row, "wOdOrdQty") = Trim(wo_rsOrdDsg!OdOrdQty)
        .Value(wi_Row, "wOdSalPrc") = Trim(wo_rsOrdDsg!OdSalPrc)
        .Value(wi_Row, "wOdBalQty") = Trim(wo_rsOrdDsg!qBalQty)
        If Trim(wo_rsOrdDsg!OdSubCust) = "" And Trim(wo_rsOrdDsg!OdPONo) = "" Then
          .Value(wi_Row, "wOdSubPoNo") = ""
        Else
          .Value(wi_Row, "wOdSubPoNo") = Trim(wo_rsOrdDsg!OdSubCust) + " / " + Trim(wo_rsOrdDsg!OdPONo)
        End If
        .Value(wi_Row, "wOdBaseVch") = Trim(wo_rsOrdDsg!qBaseVch)
        .Value(wi_Row, "wOdRefVch") = Trim(wo_rsOrdDsg!qRefVch)
        
        .Value(wi_Row, "wOdGrsWt") = Trim(wo_rsOrdDsg!qGrsWt)
        .Value(wi_Row, "wOdMetWt") = Trim(wo_rsOrdDsg!qMetWt)
        .Value(wi_Row, "wOdDiaWt") = Trim(wo_rsOrdDsg!qDiaWt)
        .Value(wi_Row, "wOdCSWt") = Trim(wo_rsOrdDsg!qCsWt)
        .Value(wi_Row, "wOdCHNWt") = Trim(wo_rsOrdDsg!qCHNWt)
        .Value(wi_Row, "wOdXWt") = Trim(wo_rsOrdDsg!qXWt)
        .Value(wi_Row, "wOdPicNm") = Trim(wo_rsOrdDsg!OdPicNm)
        .Value(wi_Row, "wOdIdNo") = Trim(wo_rsOrdDsg!OdIdNo)
        wo_rsOrdDsg.MoveNext
      Loop
    End With
  Set wo_rsOrdDsg = Nothing
  GrdOrdDsg.AllowAdd = False
  'if Show value is yes then will show salprice else it will invisible
  If ADC("wValYn") <> "Y" Then
    GrdOrdDsg.ColProp("wOdSalPrc").Style = fgcinvisible
  Else
    GrdOrdDsg.ColProp("wOdSalPrc").Style = fgcReadOnly
  End If
  'calling Grid base & refbase details
  If GrdOrdDsg.Rows > 1 Then
    GrdOrdDsg.Row = 1
    Call FillGrdBasChild(1)
    Call FillGrdRefChild(1)
  End If

End Sub
Private Sub FillFooterFields(ByVal RowNum As Integer)
        ADC("wOdSUBPoNo") = GrdOrdDsg.Value(RowNum, "wOdSUBPoNo")
        ADC("wOdBaseVch") = GrdOrdDsg.Value(RowNum, "wOdBaseVch")
        ADC("wOdRefVch") = GrdOrdDsg.Value(RowNum, "wOdRefVch")
        ADC("wOdGrsWt") = GrdOrdDsg.Value(RowNum, "wOdGrsWt")
        ADC("wOdMetWt") = GrdOrdDsg.Value(RowNum, "wOdMetWt")
        ADC("wOdDiaWt") = GrdOrdDsg.Value(RowNum, "wOdDiaWt")
        ADC("wOdCsWt") = GrdOrdDsg.Value(RowNum, "wOdCSWt")
        ADC("wOdChnWt") = GrdOrdDsg.Value(RowNum, "wOdCHNWt")
        ADC("wOdXWt") = GrdOrdDsg.Value(RowNum, "wOdXWt")
End Sub
Public Sub FillGrdBasChild(ByVal RowNum As Integer)
  Dim wsql As String
  Dim wo_rsOrdBase As MwfLib.MDORowSet, wi_Row As Integer
  'if Show value is yes then will show salprice else it will invisible
  If ADC("wValYn") <> "Y" Then
     GrdBasChild.ColProp("wRefBsSalPrc").Style = fgcinvisible
  Else
    GrdBasChild.ColProp("wRefBsSalPrc").Style = fgcReadOnly
  End If
  '578.22 conversion factor added to OdSalPrc
  wsql = "Select OdCoCd, OdTc, OdYy, OdChr, OdNo, OdSr, OmDt,OmCmCd,OdDmStkCoCd ,OdDmStkYy ," + _
         "OdDmStkChr ,OdDmStkNo ,OdOrdQty ," + IIF(mb_BaseCur, "(Case when OmMulDiv= 'M' then OdSalPrc/OmCnvFct else OdSalPrc*OmCnvFct end) ", "OdSalPrc") + " as OdSalPrc,OdSubCust,OdPoNo From OrdDsg " + _
         "Join OrdMst on OmIdNo=OdOmIdNo And OmCoCd= OdCoCd and OmTc= OdTc and OmYy= OdYy and OmChr= OdChr and OmNo= OdNo " + IIF(gs_Partition = ctCurrPrtn, " And OmPrtKey=OdPrtKey ", "") + _
         "Join DsgMst on DmIdNo=OdDmIdNo And DmTcTyp='DM' and DmCd = OdDmCd and DmSz = '' " + IIF(gs_Partition = ctCurrPrtn, " and DmPrtKey='" + ctCurrPrtn + "' ", "") + _
         "Left Outer Join CustDsg on CdCtg='C' and CdCd=OmCmCd and CdOurDmCd =OdDmCd and CdOurSfx='' and CdOurDmSz='' " + _
         " Where " + IIF(gs_Partition = ctCurrPrtn, " OdPrtKey='" + ctCurrPrtn + "' And ", "") + _
         " OdBaseIdKey=" + CStr(GrdOrdDsg.Value(RowNum, "wOdIdNo")) + " Order By OmDt, OdCoCd, OdTc, OdYy, OdChr, OdNo, OdSr "

  GrdBasChild.Rows = 1
  Set wo_rsOrdBase = moCn.OpenRes(wsql)
    GrdBasChild.AllowAdd = True
    With GrdBasChild
      Do While Not (wo_rsOrdBase.EOF Or wo_rsOrdBase.BOF)
        .AddItem
        wi_Row = .Rows - 1
        .Value(wi_Row, "wRefBsCoCd") = Trim(wo_rsOrdBase!OdCoCd)
        .Value(wi_Row, "wRefBsTc") = Trim(wo_rsOrdBase!OdTc)
        .Value(wi_Row, "wRefBsYy") = Trim(wo_rsOrdBase!OdYy)
        .Value(wi_Row, "wRefBsChr") = Trim(wo_rsOrdBase!OdChr)
        .Value(wi_Row, "wRefBsNo") = Trim(wo_rsOrdBase!OdNo)
        .Value(wi_Row, "wRefBsOrdNo") = Trim(wo_rsOrdBase!OdCoCd) + "/" + Trim(wo_rsOrdBase!OdTc) + "/" + Trim(wo_rsOrdBase!OdYy) + "/" + Trim(wo_rsOrdBase!OdChr) + "/" + CStr(Trim(wo_rsOrdBase!OdNo))
        .Value(wi_Row, "wRefBsOrdDt") = Trim(wo_rsOrdBase!OmDt)
        .Value(wi_Row, "wRefBsCust") = Trim(wo_rsOrdBase!OmCmCd)
        .Value(wi_Row, "wRefBsSr") = Trim(wo_rsOrdBase!OdSr)
        .Value(wi_Row, "wRefBsStkCoCd") = Trim(wo_rsOrdBase!OdDmStkCoCd)
        .Value(wi_Row, "wRefBsStkYy") = Trim(wo_rsOrdBase!OdDmStkYy)
        .Value(wi_Row, "wRefBsStkChr") = Trim(wo_rsOrdBase!OdDmStkChr)
        .Value(wi_Row, "wRefBsStkNo") = Trim(wo_rsOrdBase!OdDmStkNo)
        If Trim(wo_rsOrdBase!OdDmStkCoCd) = "" And Trim(wo_rsOrdBase!OdDmStkYy) = "" And Trim(wo_rsOrdBase!OdDmStkChr) = "" And Trim(wo_rsOrdBase!OdDmStkNo) = 0 Then
          .Value(wi_Row, "wRefBsStkId") = ""
        Else
          .Value(wi_Row, "wRefBsStkId") = Trim(wo_rsOrdBase!OdDmStkCoCd) + "/" + Trim(wo_rsOrdBase!OdDmStkYy) + "/" + Trim(wo_rsOrdBase!OdDmStkChr) + "/" + CStr(Trim(wo_rsOrdBase!OdDmStkNo))
        End If
        .Value(wi_Row, "wRefBsOrdQty") = Trim(wo_rsOrdBase!OdOrdQty)
        .Value(wi_Row, "wRefBsSalPrc") = Trim(wo_rsOrdBase!OdSalPrc)
        If Trim(wo_rsOrdBase!OdSubCust) = "" And Trim(wo_rsOrdBase!OdPONo) = "" Then
          .Value(wi_Row, "wRefBsSubPoNo") = ""
        Else
          .Value(wi_Row, "wRefBsSubPoNo") = Trim(wo_rsOrdBase!OdSubCust) + " / " + Trim(wo_rsOrdBase!OdPONo)
        End If
        wo_rsOrdBase.MoveNext
      Loop
    End With
    
  Set wo_rsOrdBase = Nothing
  GrdBasChild.AllowAdd = False
  
End Sub
Public Sub FillGrdRefChild(ByVal RowNum As Integer)
  Dim wsql As String
  Dim wo_rsOrdRefBase As MwfLib.MDORowSet, wi_Row As Integer

  'if Show value is yes then will show salprice else it will invisible
  If ADC("wValYn") <> "Y" Then
     GrdRefChild.ColProp("wRefBsSalPrc").Style = fgcinvisible
  Else
    GrdRefChild.ColProp("wRefBsSalPrc").Style = fgcReadOnly
  End If
  '578.22 conversion factor added to OdSalPrc
  wsql = "Select OdCoCd, OdTc, OdYy, OdChr, OdNo, OdSr, OmDt,OmCmCd,OdDmStkCoCd ,OdDmStkYy ," + _
         "OdDmStkChr ,OdDmStkNo ,OdOrdQty ," + IIF(mb_BaseCur, "(Case when OmMulDiv= 'M' then OdSalPrc/OmCnvFct else OdSalPrc*OmCnvFct end) ", "OdSalPrc") + " AS OdSalPrc,OdSubCust,OdPoNo From OrdDsg " + _
         "Join OrdMst on OmIdNo=OdOmIdNo And OmCoCd= OdCoCd and OmTc= OdTc and OmYy= OdYy and OmChr= OdChr and OmNo= OdNo " + IIF(gs_Partition = ctCurrPrtn, " And OmPrtKey=OdPrtKey ", "") + _
         "Join DsgMst on DmIdNo=OdDmIdNo And DmTcTyp='DM' and DmCd = OdDmCd and DmSz = '' " + IIF(gs_Partition = ctCurrPrtn, " and DmPrtKey='" + ctCurrPrtn + "' ", "") + _
         "Left Outer Join CustDsg on CdCtg='C' and CdCd=OmCmCd and CdOurDmCd =OdDmCd and CdOurSfx='' and CdOurDmSz='' " + _
         " Where " + IIF(gs_Partition = ctCurrPrtn, " OdPrtKey='" + ctCurrPrtn + "' And ", "") + _
         " OdRefIdKey=" + CStr(GrdOrdDsg.Value(RowNum, "wOdIdNo")) + " Order By OmDt, OdCoCd, OdTc, OdYy, OdChr, OdNo, OdSr "

  GrdRefChild.Rows = 1
  Set wo_rsOrdRefBase = moCn.OpenRes(wsql)
    GrdRefChild.AllowAdd = True
    With GrdRefChild
      Do While Not (wo_rsOrdRefBase.EOF Or wo_rsOrdRefBase.BOF)
        .AddItem
        wi_Row = .Rows - 1
        .Value(wi_Row, "wRefBsCoCd") = Trim(wo_rsOrdRefBase!OdCoCd)
        .Value(wi_Row, "wRefBsTc") = Trim(wo_rsOrdRefBase!OdTc)
        .Value(wi_Row, "wRefBsYy") = Trim(wo_rsOrdRefBase!OdYy)
        .Value(wi_Row, "wRefBsChr") = Trim(wo_rsOrdRefBase!OdChr)
        .Value(wi_Row, "wRefBsNo") = Trim(wo_rsOrdRefBase!OdNo)
        .Value(wi_Row, "wRefBsOrdNo") = Trim(wo_rsOrdRefBase!OdCoCd) + "/" + Trim(wo_rsOrdRefBase!OdTc) + "/" + Trim(wo_rsOrdRefBase!OdYy) + "/" + Trim(wo_rsOrdRefBase!OdChr) + "/" + CStr(Trim(wo_rsOrdRefBase!OdNo))
        .Value(wi_Row, "wRefBsOrdDt") = Trim(wo_rsOrdRefBase!OmDt)
        .Value(wi_Row, "wRefBsCust") = Trim(wo_rsOrdRefBase!OmCmCd)
        .Value(wi_Row, "wRefBsSr") = Trim(wo_rsOrdRefBase!OdSr)
        .Value(wi_Row, "wRefBsStkCoCd") = Trim(wo_rsOrdRefBase!OdDmStkCoCd)
        .Value(wi_Row, "wRefBsStkYy") = Trim(wo_rsOrdRefBase!OdDmStkYy)
        .Value(wi_Row, "wRefBsStkChr") = Trim(wo_rsOrdRefBase!OdDmStkChr)
        .Value(wi_Row, "wRefBsStkNo") = Trim(wo_rsOrdRefBase!OdDmStkNo)
        If Trim(wo_rsOrdRefBase!OdDmStkCoCd) = "" And Trim(wo_rsOrdRefBase!OdDmStkYy) = "" And Trim(wo_rsOrdRefBase!OdDmStkChr) = "" And Trim(wo_rsOrdRefBase!OdDmStkNo) = 0 Then
          .Value(wi_Row, "wRefBsStkId") = ""
        Else
          .Value(wi_Row, "wRefBsStkId") = Trim(wo_rsOrdRefBase!OdDmStkCoCd) + "/" + Trim(wo_rsOrdRefBase!OdDmStkYy) + "/" + Trim(wo_rsOrdRefBase!OdDmStkChr) + "/" + CStr(Trim(wo_rsOrdRefBase!OdDmStkNo))
        End If
        .Value(wi_Row, "wRefBsOrdQty") = Trim(wo_rsOrdRefBase!OdOrdQty)
        .Value(wi_Row, "wRefBsSalPrc") = Trim(wo_rsOrdRefBase!OdSalPrc)
        If Trim(wo_rsOrdRefBase!OdSubCust) = "" And Trim(wo_rsOrdRefBase!OdPONo) = "" Then
          .Value(wi_Row, "wRefBsSubPoNo") = ""
        Else
          .Value(wi_Row, "wRefBsSubPoNo") = Trim(wo_rsOrdRefBase!OdSubCust) + " / " + Trim(wo_rsOrdRefBase!OdPONo)
        End If
        wo_rsOrdRefBase.MoveNext
      Loop
    End With
  Set wo_rsOrdRefBase = Nothing
  GrdRefChild.AllowAdd = False
End Sub
''''  ***** Std Code not to be changed **********************************
Private Sub Form_GotFocus()
  '*** Jenny Color (For Setting System Colors)
  SetMDISysColorsAfterLogin
  '*** Jenny Color (For Setting System Colors)
End Sub
Private Sub Form_Load()
  'SetMDISysColorsAfterLogin
  Dim ws_BtnStr As String
  ws_BtnStr = ""
  Call SetObjColors(Me, ws_BtnStr)
  
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
    Set moCn = Nothing    '*** (09/08/05)    '*** (Bef speed) bef 24/11/06
End Sub
Private Sub GrdOrdDsg_Validate(Cancel As Boolean)
  Cancel = GrdOrdDsg.Validate
End Sub
Private Sub GrdOrdMst_Validate(Cancel As Boolean)
  Cancel = GrdOrdMst.Validate
End Sub

Private Sub FillOrdRm(ByVal RowNum As Integer)
  Dim wsql As String
  Dim wo_rsOrdRm As MwfLib.MDORowSet, wi_Row As Integer
  '578.22 conversion factor added to all rate and value and OrdMst join added
  wsql = "Select *," + _
    IIF(mb_BaseCur, "(Case when OmMulDiv= 'M' then OrSalRt/OmCnvFct else OrSalRt*OmCnvFct end) ", "OrSalRt") + " as OrSalRtForCur, " + _
    IIF(mb_BaseCur, "(Case when OmMulDiv= 'M' then OrSalVal/OmCnvFct else OrSalVal*OmCnvFct end) ", "OrSalVal") + " as OrSalValForCur," + _
    IIF(mb_BaseCur, "(Case when OmMulDiv= 'M' then OrSetSalRt/OmCnvFct else OrSetSalRt*OmCnvFct end) ", "OrSetSalRt") + " as OrSetSalRtForCur," + _
    IIF(mb_BaseCur, "(Case when OmMulDiv= 'M' then OrSetSalVal/OmCnvFct else OrSetSalVal*OmCnvFct end) ", "OrSetSalVal") + " as OrSetSalValForCur," + _
    IIF(mb_BaseCur, "(Case when OmMulDiv= 'M' then OrAlySalRt/OmCnvFct else OrAlySalRt*OmCnvFct end) ", "OrAlySalRt") + " as OrAlySalRtForCur, " + _
    IIF(mb_BaseCur, "(Case when OmMulDiv= 'M' then OrCstRt/OmCnvFct else OrCstRt*OmCnvFct end) ", "OrCstRt") + " as OrCstRtForCur, " + _
    IIF(mb_BaseCur, "(Case when OmMulDiv= 'M' then OrCstVal/OmCnvFct else OrCstVal*OmCnvFct end) ", "OrCstVal") + " as OrCstValForCur, " + _
    IIF(mb_BaseCur, "(Case when OmMulDiv= 'M' then OrSetCstRt/OmCnvFct else OrSetCstRt*OmCnvFct end) ", "OrSetCstRt") + " as OrSetCstRtForCur, " + _
    IIF(mb_BaseCur, "(Case when OmMulDiv= 'M' then OrSetCstVal/OmCnvFct else OrSetCstVal*OmCnvFct end) ", "OrSetCstVal") + " as OrSetCstValForCur, " + _
    IIF(mb_BaseCur, "(Case when OmMulDiv= 'M' then OrAlyCstRt/OmCnvFct else OrAlyCstRt*OmCnvFct end) ", "OrAlyCstRt") + " as OrAlyCstRtForCur " + _
    " from OrdRm " + _
    "Join OrdMst on OmCoCd= OrCoCd and OmTc= OrTc and OmYy= OrYy and OmChr= OrChr and OmNo= OrNo " + IIF(gs_Partition = ctCurrPrtn, " And OmPrtKey=OrPrtKey ", "") + _
    " Where " + IIF(gs_Partition = ctCurrPrtn, " OrPrtKey='" + ctCurrPrtn + "' And ", "") + _
    " OrCoCd = '" + GrdOrdDsg.Value(RowNum, "wOdCoCd") + "' and OrTc= '" + GrdOrdDsg.Value(RowNum, "wOdTc") + "' and " + _
    "OrYy='" + GrdOrdDsg.Value(RowNum, "wOdYy") + "' and OrChr = '" + GrdOrdDsg.Value(RowNum, "wOdChr") + "' and " + _
    "OrNo = " + CStr(GrdOrdDsg.Value(RowNum, "wOdNo")) + " and OrSr = " + CStr(GrdOrdDsg.Value(RowNum, "wOdSr")) + _
    " Order By OrCoCd, OrTc, OrYy, OrChr, OrNo, OrSr, OrSrNo "
  
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
      .Value(wi_Row, "OrSalRt") = Trim(wo_rsOrdRm!OrSalRtForCur)    '578.22 OrSalRt changed to OrSalRtForCur
      .Value(wi_Row, "wQWSal") = moCn.GetFldVal("Select RmQw From RmMst Where " + IIF(gs_Partition = ctCurrPrtn, " RmPrtKey='" + ctCurrPrtn + "' and ", "") + " RmCd='" + Trim(wo_rsOrdRm!OrRmCd) + "'")
      .Value(wi_Row, "OrSalVal") = Trim(wo_rsOrdRm!OrSalValForCur)  '578.22 OrSalVal changed to OrSalValForCur
      .Value(wi_Row, "OrPrdQty") = Trim(wo_rsOrdRm!OrPrdQty)
      .Value(wi_Row, "OrPrdWt") = Trim(wo_rsOrdRm!OrPrdWt)
      .Value(wi_Row, "OrCstRt") = Trim(wo_rsOrdRm!OrCstRtForCur)    '578.22 OrCstRt changed to OrCstRtForCur
      .Value(wi_Row, "OrCstVal") = Trim(wo_rsOrdRm!OrCstValForCur)  '578.22 OrCstVal changed to OrCstValForCur
      .Value(wi_Row, "OrSetSCd") = Trim(wo_rsOrdRm!OrSetSCd)
      .Value(wi_Row, "OrSetSalRt") = Trim(wo_rsOrdRm!OrSetSalRtForCur)    '578.22 OrSetSalRt changed to OrSetSalRtForCur
      .Value(wi_Row, "OrSetSalVal") = Trim(wo_rsOrdRm!OrSetSalValForCur)  '578.22 OrSetSalVal changed to OrSetSalValForCur
      .Value(wi_Row, "OrSetCstRt") = Trim(wo_rsOrdRm!OrSetCstRtForCur)  '578.22 OrSetCstRt changed to OrSetCstRtForCur
      .Value(wi_Row, "OrSetCstVal") = Trim(wo_rsOrdRm!OrSetCstValForCur) '578.22 OrSetCstVal changed to OrSetCstValForCur
      .Value(wi_Row, "OrAlyCd") = Trim(wo_rsOrdRm!OrAlyCd)
      .Value(wi_Row, "OrAlySalRt") = Trim(wo_rsOrdRm!OrAlySalRtForCur)  '578.22 OrAlySalRt changed to OrAlySalRtForCur
      .Value(wi_Row, "OrAlyCstRt") = Trim(wo_rsOrdRm!OrAlyCstRtForCur)  '578.22 OrAlyCstRt changed to OrAlyCstRtForCur
      .Value(wi_Row, "OrWsQty") = Trim(wo_rsOrdRm!OrWsQty)
      .Value(wi_Row, "OrHsQty") = Trim(wo_rsOrdRm!OrHsQty)
      .Value(wi_Row, "OrSubShp") = Trim(wo_rsOrdRm!OrSubShp)
      .Value(wi_Row, "OrMainMet") = Trim(wo_rsOrdRm!OrMainMet)
      wo_rsOrdRm.MoveNext
    Loop
  End With
  Set wo_rsOrdRm = Nothing
  GrdOrdRm.AllowAdd = False
  'If show value option is selected, shows rate & value
  If ADC("wValYn") <> "Y" Then
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

End Sub

Private Sub FillOrdLab(ByVal RowNum As Integer)
  Dim wsql As String
  Dim wo_rsOrdLab As MwfLib.MDORowSet, wi_Row As Integer
  '578.22 conversion factor added to all rate, value and OrdMst join added
  wsql = "Select *," + _
    IIF(mb_BaseCur, "(Case when OmMulDiv= 'M' then OlSalRt/OmCnvFct else OlSalRt*OmCnvFct end) ", "OlSalRt") + " as OlSalRtForCur, " + _
    IIF(mb_BaseCur, "(Case when OmMulDiv= 'M' then OlSalVal/OmCnvFct else OlSalVal*OmCnvFct end) ", "OlSalVal") + " as OlSalValForCur, " + _
    IIF(mb_BaseCur, "(Case when OmMulDiv= 'M' then OlCstRt/OmCnvFct else OlCstRt*OmCnvFct end) ", "OlCstRt") + " as OlCstRtForCur, " + _
    IIF(mb_BaseCur, "(Case when OmMulDiv= 'M' then OlCstVal/OmCnvFct else OlCstVal*OmCnvFct end) ", "OlCstVal") + " as OlCstValForCur " + _
    " from OrdLab " + _
    " Join OrdMst on OmCoCd= OlCoCd and OmTc= OlTc and OmYy= OlYy and OmChr= OlChr and OmNo= OlNo " + IIF(gs_Partition = ctCurrPrtn, " And OmPrtKey=OlPrtKey ", "") + _
    " Where " + IIF(gs_Partition = ctCurrPrtn, " OlPrtKey='" + ctCurrPrtn + "' And ", "") + _
    " OlCoCd = '" + GrdOrdDsg.Value(RowNum, "wOdCoCd") + "' and OlTc= '" + GrdOrdDsg.Value(RowNum, "wOdTc") + "' and " + _
    "OlYy='" + GrdOrdDsg.Value(RowNum, "wOdYy") + "' and OlChr = '" + GrdOrdDsg.Value(RowNum, "wOdChr") + "' and " + _
    "OlNo = " + CStr(GrdOrdDsg.Value(RowNum, "wOdNo")) + " and OlSr = " + CStr(GrdOrdDsg.Value(RowNum, "wOdSr")) + _
    " Order By OlCoCd, OlTc, OlYy, OlChr, OlNo, OlSr, OlSrNo "
  
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
      .Value(wi_Row, "OlSalRt") = Trim(wo_rsOrdLab!OlSalRtForCur)   '578.22 OlSalRt changed to OlSalRtForCur
      .Value(wi_Row, "OlSalVal") = Trim(wo_rsOrdLab!OlSalValForCur) '578.22 OlSalVal changed to OlSalValForCur
      .Value(wi_Row, "OlCstQW") = Trim(wo_rsOrdLab!OlCstQw)
      .Value(wi_Row, "OlCstRt") = Trim(wo_rsOrdLab!OlCstRtForCur)   '578.22 OlCstRt changed to OlCstRtForCur
      .Value(wi_Row, "OlCstVal") = Trim(wo_rsOrdLab!OlCstValForCur) '578.22 OlCstVal changed to OlCstValForCur
      wo_rsOrdLab.MoveNext
    Loop
  End With
  Set wo_rsOrdLab = Nothing
  GrdOrdLab.AllowAdd = False
  
  'If show value option is selected, shows rate & value
  If ADC("wValYn") <> "Y" Then
    GrdOrdLab.ColProp("OlSalRt").Style = fgcinvisible
    GrdOrdLab.ColProp("OlSalVal").Style = fgcinvisible
  Else
    GrdOrdLab.ColProp("OlSalRt").Style = fgcReadOnly
    GrdOrdLab.ColProp("OlSalVal").Style = fgcReadOnly
  End If

End Sub

Private Sub OptDelDt_GotFocus(Index As Integer)
  DispMsg OptDelDt(Index).ToolTipText, etInfo
End Sub

Private Sub TabOrdDsgDet_Click(PreviousTab As Integer)
 '*** Whenever a tab is selected enable the frame on the clicked tab and disable the previous tab frame
  FraOrdDsgDet(TabOrdDsgDet.Tab).Enabled = True:  FraOrdDsgDet(PreviousTab).Enabled = False

End Sub

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
  Case Is = DispPic
    If FraShowPic.Visible = True Then
      FraShowPic.Visible = False:      FraShowPic.Enabled = False
      Call EnaDisaCmds(False, CmdShowPic)
      CmdShowPic.SetFocus
    Else
      FraShowPic.Visible = True:      FraShowPic.Enabled = True
      FraShowPic.ZOrder
      Call EnaDisaCmds(True, CmdShowPic)
    End If
  End Select
  
End Sub

Private Sub EnaDisaCmds(ByVal pv_ShowFra As Boolean, Optional ByVal pv_CmdBut As MWCTL_BTN1)
  If Not pv_CmdBut Is Nothing Then
    If pv_ShowFra = True Then
      If pv_CmdBut.Pic = 0 Then pv_CmdBut.BackColor = vbWhite
      pv_CmdBut.CausesValidation = True
      DisableGrdBasRefChild
    Else
      If pv_CmdBut.Pic = 0 Then pv_CmdBut.BackColor = &H8000000F
      InitGrdBasRefchild
    End If
  End If
  FraNKeyAll.Enabled = Not pv_ShowFra
  ADC.AllowSave = Not pv_ShowFra
  CmdDsgDet.Enabled = Not pv_ShowFra
  CmdShowPic.Enabled = Not pv_ShowFra
  If Not pv_CmdBut Is Nothing Then pv_CmdBut.Enabled = True
End Sub
Private Sub DisableGrdBasRefChild()
      cmdBasChild.Enabled = False
      cmdRefChild.Enabled = False
End Sub
Private Sub TxtPwd_Validate(Cancel As Boolean)
  If Trim(TxtPwd.Text) = "" Then
    ADC("wValYn") = ""
    ChkBoxArr(1).Value = Unchecked
    If moCn.GetFldVal("Select PValue1 from Param where Ptyp = 'SEOPWD' and PmCd = 'SEOPWD'") = "" Then _
       mb_PwdFlag = True Else mb_PwdFlag = False
  Else
    If Trim(TxtPwd.Text) = moCn.GetFldVal("select IsNull ((select PValue from Param Join UsrAccess on UaUsrCd = PMCd " + _
                   "where PTyp= 'USR' and PMCd= '" + gs_UsrCd + "' and UaMnuCd = '" + ADC.MenuCd + "' and " + _
                   "UaSeoYn= 'Y' and PValue= '" + Trim(TxtPwd.Text) + "'), " + _
                   "(Select PValue1 from Param where Ptyp= 'SEOPWD' and PmCd= 'SEOPWD' )) ") Then
      mb_PwdFlag = True: Exit Sub
    Else
      mb_PwdFlag = False
      Cancel = True:  DispMsg "Invalid Password / User Does Not Have Seo Rights ", etError
    End If
  End If
End Sub
Private Sub Pic3D_DblClick()
'*** If mb_Pic3DDblClk flag option is False then Enlarge the 3-Dimensional picture to fit the FraNKeyAll Frame
'*** If mb_Pic3DDblClk flag option is True then bring the 3-Dimensional picture to the original pic control size
  If GrdOrdDsg.Rows - 1 = 0 Then Exit Sub
  ADC.SetFocus
  Dim w3DPth As String
  If mb_Pic3DDblClk = False Then
    '*** Enlarge The Picture ***
    w3DPth = GetPictPath("DM", False, GrdOrdDsg.Value(GrdOrdDsg.Row, "wOdDsgCd"), GrdOrdDsg.Value(GrdOrdDsg.Row, "wOdSz"), , , , , , , GrdOrdDsg.Value(GrdOrdDsg.Row, "wOdPicNm"))
    Call ShowLargePic(IIF(FraNKeyAll.Height - 195 > 0, FraNKeyAll.Height - 195, 0), FraNKeyAll.Width, w3DPth, Pic3D)

    FraPic3D.Left = 0: FraPic3D.Top = 195
    FraPic3D.Height = Pic3D.Height + 100: FraPic3D.Width = Pic3D.Width + 100
    FraNKeyAll.ZOrder
    FraPic3D.ZOrder (0)
  Else
      '*** Make The Picture Small ***
      w3DPth = GetPictPath("DM", False, GrdOrdDsg.Value(GrdOrdDsg.Row, "wOdDsgCd"), GrdOrdDsg.Value(GrdOrdDsg.Row, "wOdSz"), , , , , , , GrdOrdDsg.Value(GrdOrdDsg.Row, "wOdPicNm"))
      Call ShowPic(mi_Pic3DHt, mi_Pic3DWdt, w3DPth, Pic3D)
      FraPic3D.Left = mi_FraPic3DLft: FraPic3D.Top = mi_FraPic3DTop
      FraPic3D.Height = Pic3D.Height + 100: FraPic3D.Width = Pic3D.Width + 100
      FraNKeyAll.ZOrder
  End If
  If mb_Pic3DDblClk = False Then
      mb_Pic3DDblClk = True
  Else
      mb_Pic3DDblClk = False
  End If
End Sub
Private Sub PicHsk_DblClick()
'*** If mb_PicHskDblClk flag option is False then Enlarge the Line Drawing to fit the FraNKeyAll Frame
'*** If mb_PicHskDblClk flag option is True then bring the Line Drawing to the original pic control size
  If GrdOrdDsg.Rows - 1 = 0 Then Exit Sub
  ADC.SetFocus
  Dim wHskPth As String
  If mb_PicHskDblClk = False Then
      '*** Enlarge The Picture ***
      wHskPth = GetPictPath("DM", True, GrdOrdDsg.Value(GrdOrdDsg.Row, "wOdDsgCd"), GrdOrdDsg.Value(GrdOrdDsg.Row, "wOdSz"), , , , , , , GrdOrdDsg.Value(GrdOrdDsg.Row, "wOdPicNm"))
      Call ShowLargePic(IIF(FraNKeyAll.Height - 195 > 0, FraNKeyAll.Height - 195, 0), FraNKeyAll.Width, wHskPth, PicHsk)
      FraPicHsk.Left = 0: FraPicHsk.Top = 195
      FraPicHsk.Height = PicHsk.Height + 100: FraPicHsk.Width = PicHsk.Width + 100
      FraNKeyAll.ZOrder
      FraPicHsk.ZOrder (0)
  Else
      '*** Make The Picture Small ***
  wHskPth = GetPictPath("DM", True, GrdOrdDsg.Value(GrdOrdDsg.Row, "wOdDsgCd"), GrdOrdDsg.Value(GrdOrdDsg.Row, "wOdSz"), , , , , , , GrdOrdDsg.Value(GrdOrdDsg.Row, "wOdPicNm"))
      Call ShowPic(mi_PicHskHt, mi_PicHskWdt, wHskPth, PicHsk)
      FraPicHsk.Left = mi_FraPicHskLft: FraPicHsk.Top = mi_FraPicHskTop
      FraPicHsk.Height = PicHsk.Height + 100: FraPicHsk.Width = PicHsk.Width + 100
      FraNKeyAll.ZOrder
  End If
  If mb_PicHskDblClk = False Then
      mb_PicHskDblClk = True
  Else
      mb_PicHskDblClk = False
  End If
End Sub



