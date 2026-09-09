VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{AB2CC32A-E652-42A2-B341-26ECE02F6386}#2.1#0"; "MwfCtl.ocx"
Begin VB.Form EmrFrpJtTxnLst2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
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
      Height          =   9675
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   15255
      Begin TabDlg.SSTab TAB_REP 
         Height          =   9465
         Left            =   15
         TabIndex        =   9
         Top             =   45
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   16695
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "&Data Range"
         TabPicture(0)   =   "EmrFrpJtTxnLst2.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "fra_tabrep(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "&Group Sort"
         TabPicture(1)   =   "EmrFrpJtTxnLst2.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "fra_tabrep(1)"
         Tab(1).ControlCount=   1
         Begin VB.Frame fra_tabrep 
            BorderStyle     =   0  'None
            Height          =   8475
            Index           =   0
            Left            =   60
            TabIndex        =   12
            Top             =   600
            Width           =   15105
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Show Picture            "
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
               Left            =   75
               TabIndex        =   15
               Tag             =   "ShowPic"
               ToolTipText     =   "Check To Show Picture"
               Top             =   4350
               Width           =   2265
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Add Value"
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
               TabIndex        =   14
               Tag             =   "AddVal2"
               ToolTipText     =   "Add Value"
               Top             =   1395
               Width           =   1410
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Add Value   "
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
               Left            =   4800
               TabIndex        =   13
               Tag             =   "AddVal3"
               ToolTipText     =   "Check To Add Value"
               Top             =   1680
               Width           =   1410
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   40
               Left            =   4800
               TabIndex        =   16
               ToolTipText     =   "Enter To Customer"
               Top             =   3435
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               DataField       =   "UOCMCDTO"
               IdName          =   "UOCMCDTO"
               CmpStr          =   "OmCmCd <="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   41
               Left            =   7500
               TabIndex        =   17
               ToolTipText     =   "Enter Customer Selection"
               Top             =   3435
               Width           =   7410
               _ExtentX        =   13070
               _ExtentY        =   503
               MaxLength       =   255
               DataType        =   4
               DataField       =   "UOCMCDSEL"
               IdName          =   "UOCMCDSEL"
               CmpStr          =   "OmCmCd In"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   42
               Left            =   2115
               TabIndex        =   18
               ToolTipText     =   "Enter From Customer"
               Top             =   3435
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               DataField       =   "UOCMCDFR"
               IdName          =   "UOCMCDFR"
               CmpStr          =   "OmCmCd >="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   19
               Left            =   2115
               TabIndex        =   19
               ToolTipText     =   "Enter From Order Number"
               Top             =   2250
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOOMNOFR"
               IdName          =   "UOOMNOFR"
               CmpStr          =   "OmNo >="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   22
               Left            =   4800
               TabIndex        =   20
               ToolTipText     =   "Enter To Order Number"
               Top             =   2250
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOOMNOTO"
               IdName          =   "UOOMNOTO"
               CmpStr          =   "OmNo <="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   24
               Left            =   6405
               TabIndex        =   21
               ToolTipText     =   "Show Invoice Or Order Customer ('Y' for Invoice / 'N' For Order)"
               Top             =   1395
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
               Index           =   49
               Left            =   2115
               TabIndex        =   22
               ToolTipText     =   "Enter Company Code"
               Top             =   540
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOCOCDFR"
               IdName          =   "UOCOCDFR"
               CmpStr          =   "OmCoCd >= "
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   10
               Left            =   3315
               TabIndex        =   23
               ToolTipText     =   "Enter From Order Character"
               Top             =   1965
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOOMCHRFR"
               IdName          =   "UOOMCHRFR"
               CmpStr          =   "OmChr >="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   5
               Left            =   2895
               TabIndex        =   24
               ToolTipText     =   "Enter From Order Year"
               Top             =   1965
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               MaxLength       =   2
               DataField       =   "UOOMYYFR"
               IdName          =   "UOOMYYFR"
               CmpStr          =   "OmYy >="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   17
               Left            =   5985
               TabIndex        =   25
               ToolTipText     =   "Enter To Order Character"
               Top             =   1965
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOOMCHRTO"
               IdName          =   "UOOMCHRTO"
               CmpStr          =   "OmChr <="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   1
               Left            =   5565
               TabIndex        =   26
               ToolTipText     =   "Enter To Order Year"
               Top             =   1965
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               MaxLength       =   2
               DataField       =   "UOOMYYTO"
               IdName          =   "UOOMYYTO"
               CmpStr          =   "OmYy <="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   13
               Left            =   4800
               TabIndex        =   27
               ToolTipText     =   "Enter To Order Tc"
               Top             =   1965
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataField       =   "UOOMTCTO"
               IdName          =   "UOOMTCTO"
               CmpStr          =   "OmTc <="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   6
               Left            =   2115
               TabIndex        =   28
               ToolTipText     =   "Enter From Order Tc"
               Top             =   1965
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataField       =   "UOOMTCFR"
               IdName          =   "UOOMTCFR"
               CmpStr          =   "OmTc >="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   55
               Left            =   2115
               TabIndex        =   29
               ToolTipText     =   "Enter Currency Code Whose Invoice You Wish To View Or Leave It Blank If You Wish To View All Invoices In Base Currency"
               Top             =   4065
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataType        =   4
               DataField       =   "UOCURCDFR"
               IdName          =   "UOCURCDFR"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   73
               Left            =   4800
               TabIndex        =   30
               ToolTipText     =   "Enter To Company Code"
               Top             =   540
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOCOCDTO"
               IdName          =   "UOCOCDTO"
               CmpStr          =   "OmCoCd <= "
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   75
               Left            =   7500
               TabIndex        =   31
               ToolTipText     =   "Enter Design Code Selection"
               Top             =   540
               Width           =   7410
               _ExtentX        =   13070
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOCOCDSEL"
               IdName          =   "UOCOCDSEL"
               CmpStr          =   "OmCoCd In"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   76
               Left            =   2505
               TabIndex        =   32
               ToolTipText     =   "Show Picture (Yes / No)"
               Top             =   4350
               Visible         =   0   'False
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               Enabled         =   0   'False
               DataType        =   4
               DataField       =   "UOYN3"
               IdName          =   "UOYN3"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   80
               Left            =   6405
               TabIndex        =   33
               ToolTipText     =   "Show invoice Date/ Printng Dat (Yes / No)"
               Top             =   1680
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
               Index           =   3
               Left            =   2115
               TabIndex        =   34
               ToolTipText     =   "Enter Customer Category"
               Top             =   825
               Width           =   495
               _ExtentX        =   873
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               DataField       =   "UOCMCTGFR"
               IdName          =   "UOCMCTGFR"
               CmpStr          =   "OmCmCtg ="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   7
               Left            =   2115
               TabIndex        =   35
               ToolTipText     =   "Enter Tc 1"
               Top             =   1110
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataField       =   "UOOMTC"
               ReCalcParent    =   "UOOMTCFR"
               IdName          =   "UOOMTC"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   8
               Left            =   2115
               TabIndex        =   36
               ToolTipText     =   "Enter Tc 2"
               Top             =   1395
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOOMTC1"
               ReCalcParent    =   "UOOMTCFR"
               IdName          =   "UOOMTC1"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   9
               Left            =   2115
               TabIndex        =   37
               ToolTipText     =   "Enter Tc 3"
               Top             =   1680
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOOMTC2"
               ReCalcParent    =   "UOOMTCFR"
               IdName          =   "UOOMTC2"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   44
               Left            =   4800
               TabIndex        =   38
               ToolTipText     =   "Enter To Voucher Date"
               Top             =   2535
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   3
               DataField       =   "UODTTO"
               IdName          =   "UODTTO"
               CmpStr          =   "OmDt <="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   48
               Left            =   2115
               TabIndex        =   39
               ToolTipText     =   "Enter From Voucher Date"
               Top             =   2535
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   3
               DataField       =   "UODTFR"
               IdName          =   "UODTFR"
               CmpStr          =   "OmDt >="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   18
               Left            =   4800
               TabIndex        =   40
               ToolTipText     =   "Enter To Po Date"
               Top             =   3150
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   3
               DataField       =   "UODT1TO"
               IdName          =   "UODT1TO"
               CmpStr          =   "OmPoDt <="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   20
               Left            =   2115
               TabIndex        =   41
               ToolTipText     =   "Enter From Po Date"
               Top             =   3150
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   3
               DataField       =   "UODT1FR"
               IdName          =   "UODT1FR"
               CmpStr          =   "OmPoDt >="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   43
               Left            =   4800
               TabIndex        =   42
               ToolTipText     =   "Enter To Purchase Order "
               Top             =   2820
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
               Index           =   25
               Left            =   7500
               TabIndex        =   43
               ToolTipText     =   "Enter Purchase Order Selection"
               Top             =   2820
               Width           =   7410
               _ExtentX        =   13070
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOPONOSEL"
               IdName          =   "UOPONOSEL"
               CmpStr          =   "OmPoNo In"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   26
               Left            =   2115
               TabIndex        =   44
               ToolTipText     =   "Enter From Purchase Order "
               Top             =   2820
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
               Index           =   68
               Left            =   4800
               TabIndex        =   45
               ToolTipText     =   "Enter To Sales Executive"
               Top             =   3720
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UOSALPERTO"
               IdName          =   "UOSALPERTO"
               CmpStr          =   "OmCmSalPer <="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   69
               Left            =   7500
               TabIndex        =   46
               ToolTipText     =   "Enter Sales Executive Selection"
               Top             =   3720
               Width           =   7650
               _ExtentX        =   13494
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOSALPERSEL"
               IdName          =   "UOSALPERSEL"
               CmpStr          =   "OmCmSalPer In"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   70
               Left            =   2115
               TabIndex        =   47
               ToolTipText     =   "Enter From Sales Executive"
               Top             =   3720
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UOSALPERFR"
               IdName          =   "UOSALPERFR"
               CmpStr          =   "OmCmSalPer >="
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
               Left            =   4800
               TabIndex        =   67
               Top             =   60
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
               Left            =   2115
               TabIndex        =   66
               Top             =   60
               Width           =   1335
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
               Left            =   90
               TabIndex        =   65
               Top             =   3435
               Width           =   1425
            End
            Begin VB.Label ALBL 
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
               ForeColor       =   &H00000000&
               Height          =   285
               Index           =   11
               Left            =   90
               TabIndex        =   64
               Top             =   2250
               Width           =   1485
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
               Index           =   20
               Left            =   7500
               TabIndex        =   63
               Top             =   60
               Width           =   1575
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
               Index           =   21
               Left            =   90
               TabIndex        =   62
               Top             =   540
               Width           =   1335
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
               Index           =   8
               Left            =   5925
               TabIndex        =   61
               Top             =   1965
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
               Index           =   3
               Left            =   5505
               TabIndex        =   60
               Top             =   1965
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
               Index           =   4
               Left            =   3255
               TabIndex        =   59
               Top             =   1965
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
               Index           =   0
               Left            =   2835
               TabIndex        =   58
               Top             =   1965
               Width           =   105
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
               Index           =   12
               Left            =   90
               TabIndex        =   57
               Top             =   1965
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
               Left            =   90
               TabIndex        =   56
               Top             =   4065
               Width           =   1005
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Customer Ctg"
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
               TabIndex        =   55
               Top             =   825
               Width           =   1425
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Tc 1"
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
               Left            =   90
               TabIndex        =   54
               Top             =   1110
               Width           =   1425
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Tc 2"
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
               Left            =   90
               TabIndex        =   53
               Top             =   1395
               Width           =   1425
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Tc 3"
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
               Left            =   90
               TabIndex        =   52
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
               Left            =   90
               TabIndex        =   51
               Top             =   2820
               Width           =   1425
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Vch Dt"
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
               Left            =   90
               TabIndex        =   50
               Top             =   2520
               Width           =   1485
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "PO Dt"
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
               Index           =   22
               Left            =   90
               TabIndex        =   49
               Top             =   3150
               Width           =   1485
            End
            Begin VB.Label LblSalExec 
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
               Left            =   90
               TabIndex        =   48
               Top             =   3720
               Width           =   1845
            End
         End
         Begin VB.Frame fra_tabrep 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   9070
            Index           =   1
            Left            =   -74940
            TabIndex        =   10
            Top             =   360
            Width           =   15105
            Begin MwfCtl.MWCTL_GRP GRP_REP 
               Height          =   360
               Left            =   240
               TabIndex        =   11
               Top             =   360
               Width           =   5595
               _ExtentX        =   9869
               _ExtentY        =   635
            End
         End
      End
   End
   Begin VB.Frame FRA_REPOBJ 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   6960
      TabIndex        =   5
      Top             =   9585
      Visible         =   0   'False
      Width           =   855
      Begin CRVIEWER9LibCtl.CRViewer9 CRV_REP 
         Height          =   495
         Left            =   480
         TabIndex        =   6
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
   Begin VB.Frame fra_keys 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   465
      Left            =   5745
      TabIndex        =   2
      Top             =   9570
      Width           =   3555
      Begin MwfCtl.MWCTL_MED aktxt 
         Height          =   345
         Index           =   0
         Left            =   210
         TabIndex        =   3
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
         TabIndex        =   4
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
   Begin MwfCtl.MWCTL_HLP MWCTL_HLP1 
      Height          =   495
      Left            =   11640
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   9540
      Visible         =   0   'False
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
   End
   Begin MwfCtl.MWCTL_DST DST_REP 
      Height          =   570
      Left            =   4410
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   9570
      Width           =   6420
      _ExtentX        =   11324
      _ExtentY        =   1005
   End
   Begin MwfCtl.MWCTL_ADC adc 
      Height          =   555
      Left            =   60
      TabIndex        =   7
      Top             =   9600
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
End
Attribute VB_Name = "EmrFrpJtTxnLst2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rep As New EmrRepJtTxnLst
Dim moCn As MwfLib.MDOConnection

Private Sub ADC_NonKeyWhen(pr_Cancel As Boolean, pr_ErrMsg As String)
' If adc("UoYN") = "" Then adc("UoYN") = "N"
' If adc("UoYN") = "Y" Then ChkBoxArr(0).Value = Checked Else ChkBoxArr(0).Value = Unchecked
End Sub

Private Sub adc_setreprecsource()
On Error GoTo RepErr
  Dim grpflds As String, wrepcnd As String, wCnd As String, wSqlStrg As String, ws_OrdBy As String
  Dim ws_DetOrdBy As String
  Set Rep = Nothing
  Set adc.RepSource = Rep
  Call DispCoNm
  Call SetMwName(Rep)

  wrepcnd = adc.RepCond
  wCnd = IIF(wrepcnd <> "", " and " + wrepcnd, "")
  

'  If adc("UoPMCdFr") = "CDBAGNO" Then
'    ws_DetOrdBy = " ,CdRmCtg, CdRmLotNo,CdCertNo"
'  ElseIf adc("UoPMCdFr") = "CDCERT" Then
'    ws_DetOrdBy = " ,CdRmCtg , CdRmLotNo , CdBCoCd, CdBYy ,CdBChr, CdBNo"
'  ElseIf adc("UoPMCdFr") = "CDDT" Then
'    ws_DetOrdBy = " ,CdRmCtg , CdRmLotNo , CdBCoCd, CdBYy ,CdBChr, CdBNo, CdCertNo"
'  ElseIf adc("UoPMCdFr") = "CDRMLOT" Then
'    ws_DetOrdBy = " , CdBCoCd, CdBYy ,CdBChr, CdBNo, CdCertNo"
'  End If
  
'  If adc("UoPMCdFr") <> "" Then
'    ws_OrdBy = "Order By " + moCn.GetFldVal("Select PDesc225 from Param " + _
'             "where PTyp='SRTCD' and PMCd='" + adc("UoPMCdFr") + "'") + ws_DetOrdBy
'  ElseIf adc("UoPMCdFr") = "" Then
'    ws_OrdBy = "Order By CdRmCtg,CdRmLotNo,CdBCoCd, CdBYy ,CdBChr, CdBNo,CdCertNo"
'  End If
'
'
'  If adc("UoCdCtg") = "L" Then
 
'   wSqlStrg = "Select  " + _
'             "CdDt,CdRmLotNo,CdCertNo,CdRmCtg, CdRmQty,CdRmCol,CdRmQlty, " + _
'             " CdRmCd,CdRmWt,CdFMBarCd,CdComment,CdDocNo,CdDocDt, CdRem,  " + _
'             " '' as CdBCoCd, '' as CdBYy, '' as CdBChr, 0 as CdBNo " + _
'             "from CertDet where CdCtg='L' " + wCnd + ws_OrdBy
             
'  ElseIf adc("UoCdCtg") = "B" Then
'
'  wSqlStrg = "Select  " + _
'             "CdDt,CdBCoCd, CdBYy, CdBChr, CdBNo,CdCertNo, CdComment, CdDocNo,CdDocDt,CdRem, " + _
'             " '' as CdRmLotNo, '' as CdRmCtg, 0 as CdRmQty, '' as CdRmCol,  '' as CdRmQlty, " + _
'             "'' as CdRmCd, 0 as CdRmWt, '' as CdFMBarCd " + _
'             "from CertDet where CdCtg='B'" + wCnd + ws_OrdBy
'
' End If
  wSqlStrg = "Select OmCoCd,OmTc,OmYy,OmChr,OmNo,max(OmDt) as OmDt,max(OmFrDmLoc) as OmFrDmLoc,max(OmToDmLoc) as OmToDmLoc," + _
            "max(OmCmCd) as OmCmCd,max(OmPoNo) as OmPoNo,max(OmPoDt) as OmPoDt,OdSr,max(OdDmCd) as OdDmCd,max(OdSfx) as OdSfx," + _
            "max(OdDmSz) as OdDmSz,max(OdDmStkCoCd) as OdDmStkCoCd,max(OdDmStkYy) as OdDmStkYy, max(OdDmStkChr) as OdDmStkChr,max(OdDmStkNo) as OdDmStkNo," + _
            "max(OdDmCol) as OdDmCol,max(OdKt) as OdKt ,max(OdOrdQty) as OdOrdQty, max(Convert(Decimal(14,3), OdSalPrc)) as OdSalPrc, " + _
            "max(OdSalPrc*OdOrdQty) as qSalVal,sum(case when IsNull(OrRmCtg, '') in ('D','C') then IsNull(OrWt, 0)/5 else  IsNull(OrWt, 0) end) as qGrsWt ," + _
            "sum((case when    IsNull(OrRmCtg, '') in ('G','P','S','L','M')  then  IsNull( OrWt, 0) else 0 end)* OdOrdQty) as qMetWt ," + _
            "sum((case when OrRmCtg in ('G', 'P', 'S', 'L', 'M') then RmPurityWt Else 0 end) /(case when OrRmCtg in ('G', 'P', 'S', 'L', 'M') " + _
            "then IsNull(qPurityWt, 1) Else 1 End)) as qPureWt,sum((case when    IsNull(OrRmCtg, '') ='D' then  IsNull( OrWt, 0) else 0 end) * OdOrdQty )  as qDiaWt," + _
            "sum((case when    IsNull(OrRmCtg, '') ='D' then OrQty else 0 end)* OdOrdQty)  as qDiaQty ," + _
            "sum((case when    IsNull(OrRmCtg, '') ='C' then IsNull( OrWt, 0) else 0 end)* OdOrdQty)  as qCsWt," + _
            "sum((case when    IsNull(OrRmCtg, '') ='C' then OrQty else 0 end)* OdOrdQty)  as qCsQty, " + _
            "sum((case when    IsNull(OrRmCtg, '') ='X' then  IsNull( OrWt, 0) else 0 end)* OdOrdQty) as qXWt " + _
            "From OrdMst Join OrdDsg On OdCoCd=OmCoCd and OdTc=OmTc and OdYy=OmYy and OdChr=OmChr and OdNo=OmNo and OdOmIdNo=OmIdNo and OdPrtKey=OmPrtKey " + _
            "Join OrdRm On OrCoCd=OdCoCd and OrTc=OdTc and OrYy=OdYy and OrChr=OdChr and OrNo=OdNo and OrSr=OdSr and OrOdIdNo=OdIdNo and OrPrtKey=OdPrtKey " + _
            "Join DsgMst on DmIdNo=OdDmIdNo and DmTcTyp= 'DM' and DmCd= OdDmCd and DmSz= ''  and DmPrtKey=OdPrtKey " + _
            "Join RmMst On RmCd=OrRmCd  and RmPrtKey=OrPrtKey " + _
            "Left Outer Join (Select PMCd as qPMCd, RmPurityWt as qPurityWt from Param join RmMst on  RmPrtKey='C' and  RmCd= PValue " + _
            "where PTyp='RMCTG')PRm On qPMCd= OrRmCtg " + wCnd + " group by OmCoCd, OmTc, OmYy, OmChr, OmNo, OdSr "
  
  adc.RepRecSource = wSqlStrg + ctMaxDopOpt
  
'  If adc("UoCdCtg") = "L" Then
'    Rep.TxtHead.SetText "Certificate Listing - RM Lot"
'    Rep.Hd2.Suppress = True
'    Rep.DET2.Suppress = True
'    Rep.DET2A.Suppress = True
'
'     If adc("UoYN") = "N" Then
'        Rep.TxtCmnt1.Suppress = True
'        Rep.TxtDocNo1.Suppress = True
'        Rep.TxtDocDt1.Suppress = True
'        Rep.TxtRem1.Suppress = True
'        Rep.DET1A.Suppress = True
'     End If
'
'  ElseIf adc("UoCdCtg") = "B" Then
'    Rep.TxtHead.SetText "Certificate Listing - Bag"
'    Rep.Hd1.Suppress = True
'    Rep.DET1.Suppress = True
'    Rep.DET1A.Suppress = True
'
'    If adc("UoYN") = "N" Then
'        Rep.TxtCmnt2.Suppress = True
'        Rep.TxtRem2.Suppress = True
'        Rep.DET2A.Suppress = True
'    End If
'
'  End If
  
  CRV_REP.DisplayGroupTree = False
  Call DispReport(adc)
  
Exit Sub
RepErr:
  Call ChkSelScope(Err.Number, Err.Description)
End Sub

'Private Sub ChkBoxArr_Click(Index As Integer)
' With ChkBoxArr(Index)
' Select Case Index
'   Case Is = 0    'Tag = "ShowLine"
'    If .Value = Checked Then adc("UoYN") = "Y" Else adc("UoYN") = "N"
'  End Select
' End With
'End Sub
'Private Sub ChkBoxArr_GotFocus(Index As Integer)
'  DispMsg ChkBoxArr(Index).ToolTipText, etInfo
'End Sub
Private Sub Form_GotFocus()
  SetMDISysColorsAfterLogin
End Sub
Private Sub Form_Load()
  Dim ws_BtnStr As String
  ws_BtnStr = ""
  Call SetObjColors(Me, ws_BtnStr)
End Sub
Private Sub ADC_Load()
    Set moCn = adc.Connection
'    Set adc.FirNKeyCtl = adc("UoCdCtg")
    Call InitProp(Me)
End Sub
Private Sub atxt_FldChange(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String, ByVal pv_NewValue As Variant)
  Call SetProp(Me, IdName, Valid, Cancel, ErrMsg)
End Sub
Private Sub ATXT_FldWhen(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String)
'  Select Case UCase(IdName)
'    Case Is = UCase("UoLotNoFr"), UCase("UoLotNoTo")
'      If adc.Mode = xNorm Then
'        HlpList.RmLotNo (adc("UoRmCtgFr")): Exit Sub
'      End If
'      Case Is = UCase("UoPMCdFr")
'      If adc.Mode = xNorm Then
'        HlpList.PMCd "SRTCD": Exit Sub
'      End If
'      Case Is = UCase("UoRmCdFr"), UCase("UoRmCdTo")
'      If adc.Mode = xNorm Then
'      HlpList.RmCd "N", "'" + adc("UoRmCtgFr") + "'": Exit Sub
'      End If
'
'  End Select
'  SetProp Me, IdName, When
End Sub

Private Sub tab_REP_Click(PreviousTab As Integer)
    fra_tabrep(TAB_REP.Tab).Enabled = True
    fra_tabrep(PreviousTab).Enabled = False
End Sub
Private Sub ADC_KeyWhen()
    adc("UoUsrCd") = adc.UsrCd
    adc("UoMnuCd") = adc.MenuCd
   
End Sub
Private Sub DispCoNm()
'  Rep.wCoCd.SetText gs_CoNm
'
'  Rep.wCoCdLogo.SetText gs_CoNm
'  If moCn.RecSeek("Select 'x' From Head Where HCoCd='" + gs_CoCd + "' and HCd='" + ctSelfCmCd + "' and HShowRepLogo='Y' ") Then
'    Rep.RHMain.UnderlaySection = True: Rep.wCoCd.Suppress = True
'  Else
'    Rep.RHLogo.Suppress = True
'  End If
'
'  Dim ws_Dt As String, ws_Tm As String
'    ws_Dt = Format$(moCn.SrvrDate, "DD-MMM-YYYY")
'    ws_Tm = Replace(Format(IIF(CDbl(moCn.SrvrTime) >= 13#, moCn.SrvrTime - 12, IIF(CDbl(moCn.SrvrTime) < 1#, moCn.SrvrTime + 12, moCn.SrvrTime)), "00.00"), ".", ":") + IIF(CDbl(moCn.SrvrTime) >= 12#, " PM", " AM")
'  Rep.TxtRepUsrDtTime.SetText Trim(adc.UsrCd) + " / " + ws_Dt + " / " + ws_Tm
'
'  Rep.TxtCdCtg.SetText adc("UoCdCtg"):  Rep.TxtRmCtg.SetText adc("UoRmCtgFr")
'  Rep.TxtFrCertDt.SetText adc("UoDtFr"):   Rep.TxtToCertDt.SetText adc("UoDtTo")
'  Rep.TxtFrCertNo.SetText adc("UoCdCertNoFr"):  Rep.TxtToCertNo.SetText adc("UoCdCertNoTo")
'  Rep.TxtFrRmLotNo.SetText adc("UoLotNoFr"):  Rep.TxtToRmLotNo.SetText adc("UoLotNoTo")
'  Rep.TxtFrBCoCd.SetText adc("UoCoCdFr"):  Rep.TxtFrBYy.SetText adc("UoBYyFr")
'  Rep.TxtFrBChr.SetText adc("UoBChrFr"):  Rep.TxtFrBNo.SetText adc("UoBNoFr")
'  Rep.TxtToBCoCd.SetText adc("UoCoCdTo"):  Rep.TxtToBYy.SetText adc("UoBYyTo")
'  Rep.TxtToBChr.SetText adc("UoBChrTo"):  Rep.TxtToBagNo.SetText adc("UoBNoTo")
'  Rep.TxtFrRmCode.SetText adc("UoRmCdFr"):  Rep.TxtToRmCode.SetText adc("UoRmCdTo")
'  Rep.TxtFrFMBarCd.SetText adc("UoCdFMBarCdFr"):  Rep.TxtToFMBarCd.SetText adc("UoCdFMBarCdTo")

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
    Cancel = ATXT(Index).Validate
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode <> vbFormCode Then Cancel = True: Exit Sub
    Set moCn = Nothing
    Set Rep = Nothing
End Sub




