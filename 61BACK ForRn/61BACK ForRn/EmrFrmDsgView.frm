VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{AB2CC32A-E652-42A2-B341-26ECE02F6386}#2.1#0"; "MwfCtlForSpeed7.ocx"
Begin VB.Form EmrFrmDsgView 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   9810
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15210
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9810
   ScaleWidth      =   15210
   WindowState     =   2  'Maximized
   Begin MwfCtl.MWCTL_HLP MWCTL_HLP1 
      Height          =   375
      Left            =   9630
      TabIndex        =   80
      TabStop         =   0   'False
      Top             =   9480
      Visible         =   0   'False
      Width           =   435
      _ExtentX        =   767
      _ExtentY        =   661
   End
   Begin MwfCtl.MWCTL_BTN1 CmdDetScope 
      Height          =   435
      Left            =   6870
      TabIndex        =   77
      ToolTipText     =   "Option To See The Summary Details"
      Top             =   9300
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   767
      BackColor       =   -2147483633
      MaskColor       =   8454143
      Caption         =   "&Det Scope"
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
   Begin MwfCtl.MWCTL_ADC ADC 
      CausesValidation=   0   'False
      Height          =   555
      Left            =   150
      TabIndex        =   81
      Top             =   9300
      Width           =   14865
      _ExtentX        =   26220
      _ExtentY        =   979
      AllowAdd        =   0   'False
      AllowDelete     =   0   'False
      BtnWidth        =   885
      BtnHeight       =   435
      BtnMaskColor    =   8454143
      FontSz          =   8.25
   End
   Begin MwfCtl.MWCTL_BTN1 CmdCpyCat 
      Height          =   435
      Left            =   8040
      TabIndex        =   224
      ToolTipText     =   "Option To See The Summary Details"
      Top             =   9300
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   767
      BackColor       =   -2147483633
      MaskColor       =   8454143
      Caption         =   "&Cpy Catalog"
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
   Begin VB.Frame FRA_KEYS 
      BorderStyle     =   0  'None
      Height          =   9195
      Left            =   150
      TabIndex        =   82
      Top             =   0
      Width           =   14940
      Begin VB.Frame FraScop 
         Height          =   8925
         Left            =   0
         TabIndex        =   161
         Top             =   0
         Width           =   14835
         Begin VB.Frame Frame1 
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   2970
            TabIndex        =   250
            ToolTipText     =   "Select if to Show Design On Hold - 'Yes', 'No', 'Both'"
            Top             =   6020
            Width           =   6945
            Begin VB.OptionButton OptDmValid 
               Caption         =   "No"
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
               Left            =   2550
               TabIndex        =   51
               ToolTipText     =   "Select if to Show Valid Designs - 'Yes', 'No', 'Both'"
               Top             =   0
               Width           =   2055
            End
            Begin VB.OptionButton OptDmValid 
               Caption         =   "Both"
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
               Left            =   5085
               TabIndex        =   52
               ToolTipText     =   "Select if to Show Valid Designs - 'Yes', 'No', 'Both'"
               Top             =   0
               Width           =   1665
            End
            Begin VB.OptionButton OptDmValid 
               Caption         =   "Yes"
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
               TabIndex        =   50
               ToolTipText     =   "Select if to Show Valid Designs - 'Yes', 'No', 'Both'"
               Top             =   0
               Width           =   1545
            End
         End
         Begin VB.Frame FraOptRm 
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   2970
            TabIndex        =   245
            ToolTipText     =   "Select Production Rm Qty, Wt or Sales Rm Qty, Wt"
            Top             =   6295
            Width           =   3975
            Begin VB.OptionButton OptRmPrd 
               Caption         =   "Sales Rm"
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
               Left            =   2550
               TabIndex        =   54
               ToolTipText     =   "Select Production Rm Qty, Wt or Sales Rm Qty, Wt"
               Top             =   0
               Width           =   1455
            End
            Begin VB.OptionButton OptRmPrd 
               Caption         =   "Prd Rm"
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
               TabIndex        =   53
               ToolTipText     =   "Select Production Rm Qty, Wt or Sales Rm Qty, Wt"
               Top             =   0
               Width           =   1245
            End
         End
         Begin VB.Frame FraOptDsgExc 
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   2970
            TabIndex        =   244
            ToolTipText     =   "Select if to Show Free Design or Region Exclusive Or Not Exclusive Designs Or All Designs"
            Top             =   6590
            Visible         =   0   'False
            Width           =   8565
            Begin VB.OptionButton OptDsgExc 
               Caption         =   "Free Design"
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
               Left            =   5085
               TabIndex        =   57
               ToolTipText     =   "Select if to Show Free Design or Region Exclusive Or Not Exclusive Designs Or All Designs"
               Top             =   0
               Width           =   1545
            End
            Begin VB.OptionButton OptDsgExc 
               Caption         =   "Region  Exclusive"
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
               TabIndex        =   55
               ToolTipText     =   "Select if to Show Free Design or Region Exclusive Or Not Exclusive Designs Or All Designs"
               Top             =   0
               Width           =   2055
            End
            Begin VB.OptionButton OptDsgExc 
               Caption         =   "Not Exclusive"
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
               Left            =   2550
               TabIndex        =   56
               ToolTipText     =   "Select if to Show Free Design or Region Exclusive Or Not Exclusive Designs Or All Designs"
               Top             =   0
               Width           =   1665
            End
            Begin VB.OptionButton OptDsgExc 
               Caption         =   "All Designs"
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
               Index           =   3
               Left            =   7200
               TabIndex        =   58
               ToolTipText     =   "Select if to Show Free Design or Region Exclusive Or Not Exclusive Designs Or All Designs"
               Top             =   0
               Width           =   1665
            End
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   28
            Left            =   5535
            TabIndex        =   2
            ToolTipText     =   "Enter To Design Category"
            Top             =   555
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   503
            MaxLength       =   5
            DataType        =   4
            DataField       =   "UODMCTGTO"
            IdName          =   "WDMCTGTO"
            CmpStr          =   "DmCtg <="
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   30
            Left            =   2985
            TabIndex        =   1
            ToolTipText     =   "Enter From Design Category"
            Top             =   555
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   503
            MaxLength       =   5
            DataType        =   4
            DataField       =   "UODMCTGFR"
            IdName          =   "WDMCTGFR"
            CmpStr          =   "DmCtg >="
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   27
            Left            =   5535
            TabIndex        =   5
            ToolTipText     =   "Enter To Design Production Category"
            Top             =   840
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   503
            MaxLength       =   5
            DataType        =   4
            IdName          =   "WDMPRDCTGTO"
            CmpStr          =   "DmPrdCtg <="
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   29
            Left            =   2985
            TabIndex        =   4
            ToolTipText     =   "Enter From Design Production Category"
            Top             =   840
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   503
            MaxLength       =   5
            DataType        =   4
            IdName          =   "WDMPRDCTGFR"
            CmpStr          =   "DmPrdCtg >="
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   31
            Left            =   5535
            TabIndex        =   8
            ToolTipText     =   "Enter To Sales Category"
            Top             =   1125
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   503
            MaxLength       =   5
            DataType        =   4
            IdName          =   "WDMSALCTGTO"
            CmpStr          =   "DmSalCtg <="
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   32
            Left            =   2985
            TabIndex        =   7
            ToolTipText     =   "Enter From Sales Category"
            Top             =   1125
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   503
            MaxLength       =   5
            DataType        =   4
            IdName          =   "WDMSALCTGFR"
            CmpStr          =   "DmSalCtg >="
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   0
            Left            =   5535
            TabIndex        =   11
            ToolTipText     =   "Enter To Design Code"
            Top             =   1410
            Width           =   2385
            _ExtentX        =   4207
            _ExtentY        =   503
            MaxLength       =   15
            IdName          =   "WDMCDTO"
            CmpStr          =   "DmCd <="
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   33
            Left            =   2985
            TabIndex        =   10
            ToolTipText     =   "Enter From Design Code"
            Top             =   1410
            Width           =   2385
            _ExtentX        =   4207
            _ExtentY        =   503
            MaxLength       =   15
            IdName          =   "WDMCDFR"
            CmpStr          =   "DmCd >="
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   34
            Left            =   5535
            TabIndex        =   14
            ToolTipText     =   "Enter To Design Size"
            Top             =   1695
            Width           =   585
            _ExtentX        =   1032
            _ExtentY        =   503
            MaxLength       =   5
            IdName          =   "WDMSZTO"
            CmpStr          =   "DmSz <="
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   36
            Left            =   2985
            TabIndex        =   13
            ToolTipText     =   "Enter From Design Size"
            Top             =   1695
            Width           =   585
            _ExtentX        =   1032
            _ExtentY        =   503
            MaxLength       =   5
            IdName          =   "WDMSZFR"
            CmpStr          =   "DmSz >="
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   1
            Left            =   11715
            TabIndex        =   60
            ToolTipText     =   "Show Production Rm Qty, Wt ('Y' For Prd/ 'N' For Ord)"
            Top             =   6295
            Visible         =   0   'False
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   503
            Enabled         =   0   'False
            DataType        =   4
            IdName          =   "WPRDRMQTYWTYN"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   110
            Left            =   2985
            TabIndex        =   35
            ToolTipText     =   "Enter From Diamond Weight"
            Top             =   3990
            Width           =   1020
            _ExtentX        =   1799
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "#####0.000"
            MaxLength       =   10
            DataType        =   2
            IdName          =   "WRMDIAWTFR"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   2
            Left            =   5535
            TabIndex        =   36
            ToolTipText     =   "Enter To Diamond Weight"
            Top             =   3975
            Width           =   1020
            _ExtentX        =   1799
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "#####0.000"
            MaxLength       =   10
            DataType        =   2
            IdName          =   "WRMDIAWTTO"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   3
            Left            =   2985
            TabIndex        =   16
            ToolTipText     =   "Enter From Design Set Family Code"
            Top             =   1980
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   503
            MaxLength       =   15
            DataType        =   4
            IdName          =   "WSETCDFR"
            CmpStr          =   "DmSetCd >="
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   4
            Left            =   5535
            TabIndex        =   17
            ToolTipText     =   "Enter  To Design Set Family Code"
            Top             =   1980
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   503
            MaxLength       =   15
            DataType        =   4
            IdName          =   "WSETCDTO"
            CmpStr          =   "DmSetCd <="
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   5
            Left            =   2985
            TabIndex        =   25
            ToolTipText     =   "Enter From Design Color"
            Top             =   2850
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   503
            MaxLength       =   5
            DataType        =   4
            IdName          =   "WDMCOLFR"
            CmpStr          =   "DmCol >="
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   6
            Left            =   5535
            TabIndex        =   26
            ToolTipText     =   "Enter  To Design Color"
            Top             =   2835
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   503
            MaxLength       =   5
            DataType        =   4
            IdName          =   "WDMCOLTO"
            CmpStr          =   "DmCol <="
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   7
            Left            =   11715
            TabIndex        =   59
            ToolTipText     =   "Show Design On Hold (Y/N/' ')"
            Top             =   5730
            Visible         =   0   'False
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   503
            Enabled         =   0   'False
            DataType        =   4
            IdName          =   "WONHLDYN"
            CmpStr          =   "DmHld="
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   8
            Left            =   2985
            TabIndex        =   31
            ToolTipText     =   "Enter From Main Metal Weight"
            Top             =   3420
            Width           =   1020
            _ExtentX        =   1799
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "#####0.000"
            MaxLength       =   10
            DataType        =   2
            IdName          =   "WMAINMETWTFR"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   114
            Left            =   5535
            TabIndex        =   32
            ToolTipText     =   "Enter To Main Metal Weight"
            Top             =   3405
            Width           =   1020
            _ExtentX        =   1799
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "#####0.000"
            MaxLength       =   10
            DataType        =   2
            IdName          =   "WMAINMETWTTO"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   128
            Left            =   2985
            TabIndex        =   33
            ToolTipText     =   "Enter From Gross Weight"
            Top             =   3705
            Width           =   1020
            _ExtentX        =   1799
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "#####0.000"
            MaxLength       =   10
            DataType        =   2
            IdName          =   "WRMGRSWTFR"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   129
            Left            =   5535
            TabIndex        =   34
            ToolTipText     =   "Enter To Gross Weight"
            Top             =   3690
            Width           =   1020
            _ExtentX        =   1799
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "#####0.000"
            MaxLength       =   10
            DataType        =   2
            IdName          =   "WRMGRSWTTO"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   9
            Left            =   2985
            TabIndex        =   45
            ToolTipText     =   "Enter From Design Date"
            Top             =   5415
            Width           =   930
            _ExtentX        =   1640
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   3
            IdName          =   "WODTFR"
            CmpStr          =   "DmDsgDt>="
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   15
            Left            =   5535
            TabIndex        =   46
            ToolTipText     =   "Enter To Design Date"
            Top             =   5400
            Width           =   930
            _ExtentX        =   1640
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   3
            IdName          =   "WODTTO"
            CmpStr          =   "DmDsgDt<="
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   21
            Left            =   5535
            TabIndex        =   23
            ToolTipText     =   "Enter To Design Production Sequence"
            Top             =   2550
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   4
            IdName          =   "WDMPRDSEQTO"
            CmpStr          =   "DmPrdSeq<="
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   23
            Left            =   2985
            TabIndex        =   22
            ToolTipText     =   "Enter From Design Production Sequence"
            Top             =   2565
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   4
            IdName          =   "WDMPRDSEQFR"
            CmpStr          =   "DmPrdSeq >="
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   10
            Left            =   2985
            TabIndex        =   39
            ToolTipText     =   "Enter From Colour Stone Weight"
            Top             =   4560
            Width           =   1020
            _ExtentX        =   1799
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "#####0.000"
            MaxLength       =   10
            DataType        =   2
            IdName          =   "WRMCSWTFR"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   11
            Left            =   5535
            TabIndex        =   40
            ToolTipText     =   "Enter To Colour Stone Weight"
            Top             =   4545
            Width           =   1020
            _ExtentX        =   1799
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "#####0.000"
            MaxLength       =   10
            DataType        =   2
            IdName          =   "WRMCSWTTO"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   12
            Left            =   2985
            TabIndex        =   41
            ToolTipText     =   "Enter From Colour Stone Quantity"
            Top             =   4845
            Width           =   1020
            _ExtentX        =   1799
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "#####0"
            MaxLength       =   6
            DataType        =   1
            IdName          =   "WRMCSQTYFR"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   13
            Left            =   5535
            TabIndex        =   42
            ToolTipText     =   "Enter To Colour Stone Quantity"
            Top             =   4830
            Width           =   1020
            _ExtentX        =   1799
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "#####0"
            MaxLength       =   6
            DataType        =   1
            IdName          =   "WRMCSQTYTO"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   14
            Left            =   2985
            TabIndex        =   0
            ToolTipText     =   "Enter Design Type"
            Top             =   270
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   503
            MaxLength       =   2
            DataType        =   4
            IdName          =   "WDMTYP"
            CmpStr          =   "DmTcTyp = "
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   16
            Left            =   11715
            TabIndex        =   61
            ToolTipText     =   "Show Free Designs Or Region Exclusive Designs Or Not Exclusive Designs (Blank For Free, 'Y' For Exclusive, 'N' For Not Exclusive)"
            Top             =   6590
            Visible         =   0   'False
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   503
            Enabled         =   0   'False
            DataType        =   4
            IdName          =   "WDMREGEXC"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   18
            Left            =   5535
            TabIndex        =   64
            ToolTipText     =   "Enter To Customer"
            Top             =   7110
            Visible         =   0   'False
            Width           =   1065
            _ExtentX        =   1879
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   4
            IdName          =   "WCMCDTO"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   26
            Left            =   2985
            TabIndex        =   37
            ToolTipText     =   "Enter From Diamond Quantity"
            Top             =   4275
            Width           =   1020
            _ExtentX        =   1799
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "#####0"
            MaxLength       =   6
            DataType        =   1
            IdName          =   "WRMDIAQTYFR"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   60
            Left            =   5535
            TabIndex        =   38
            ToolTipText     =   "Enter To Diamond Quantity"
            Top             =   4260
            Width           =   1020
            _ExtentX        =   1799
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "#####0"
            MaxLength       =   6
            DataType        =   1
            IdName          =   "WRMDIAQTYTO"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   75
            Left            =   8070
            TabIndex        =   3
            ToolTipText     =   "Enter Design Category Selection"
            Top             =   555
            Width           =   6540
            _ExtentX        =   11536
            _ExtentY        =   503
            MaxLength       =   255
            DataType        =   4
            IdName          =   "WDMCTGSEL"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   76
            Left            =   8070
            TabIndex        =   6
            ToolTipText     =   "Enter Production Category Selection"
            Top             =   840
            Width           =   6540
            _ExtentX        =   11536
            _ExtentY        =   503
            MaxLength       =   255
            DataType        =   4
            IdName          =   "WPRDCTGSEL"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   77
            Left            =   8070
            TabIndex        =   9
            ToolTipText     =   "Enter Sales Category Selection"
            Top             =   1125
            Width           =   6540
            _ExtentX        =   11536
            _ExtentY        =   503
            MaxLength       =   255
            DataType        =   4
            IdName          =   "WSALCTGSEL"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   78
            Left            =   8070
            TabIndex        =   12
            ToolTipText     =   "Enter Design Code Selection"
            Top             =   1410
            Width           =   6540
            _ExtentX        =   11536
            _ExtentY        =   503
            MaxLength       =   255
            DataType        =   4
            IdName          =   "WDMCDSEL"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   79
            Left            =   8070
            TabIndex        =   15
            ToolTipText     =   "Enter Design Size Selection"
            Top             =   1695
            Width           =   6540
            _ExtentX        =   11536
            _ExtentY        =   503
            MaxLength       =   255
            DataType        =   4
            IdName          =   "WDMSZSEL"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   80
            Left            =   8070
            TabIndex        =   18
            ToolTipText     =   "Enter Set Family Code Selection"
            Top             =   1980
            Width           =   6540
            _ExtentX        =   11536
            _ExtentY        =   503
            MaxLength       =   255
            DataType        =   4
            IdName          =   "WSETCDSEL"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   81
            Left            =   8070
            TabIndex        =   24
            ToolTipText     =   "Enter Production Sequence Selection"
            Top             =   2550
            Width           =   6540
            _ExtentX        =   11536
            _ExtentY        =   503
            MaxLength       =   255
            DataType        =   4
            IdName          =   "WPRDSEQSEL"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   82
            Left            =   8070
            TabIndex        =   27
            ToolTipText     =   "Enter Design Colour Selection"
            Top             =   2835
            Width           =   6540
            _ExtentX        =   11536
            _ExtentY        =   503
            MaxLength       =   255
            DataType        =   4
            IdName          =   "WDMCOLSEL"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   83
            Left            =   2985
            TabIndex        =   43
            ToolTipText     =   "Enter From Design Created Date"
            Top             =   5130
            Width           =   930
            _ExtentX        =   1640
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   3
            IdName          =   "WOCRDTFR"
            CmpStr          =   "DmCreatedDt>="
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   85
            Left            =   5535
            TabIndex        =   44
            ToolTipText     =   "Enter To Design Created Date"
            Top             =   5115
            Width           =   930
            _ExtentX        =   1640
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   3
            IdName          =   "WOCRDTTO"
            CmpStr          =   "DmCreatedDt<="
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   98
            Left            =   2985
            TabIndex        =   28
            ToolTipText     =   "Enter From Customer "
            Top             =   3135
            Width           =   1065
            _ExtentX        =   1879
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   4
            IdName          =   "WDRSETSCDFR"
            CmpStr          =   "DrSetSCd >="
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   99
            Left            =   5535
            TabIndex        =   29
            ToolTipText     =   "Enter From Customer "
            Top             =   3120
            Width           =   1065
            _ExtentX        =   1879
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   4
            IdName          =   "WDRSETSCDTO"
            CmpStr          =   "DrSetSCd <="
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   100
            Left            =   8070
            TabIndex        =   30
            ToolTipText     =   "Enter Catalog Selection"
            Top             =   3120
            Width           =   6540
            _ExtentX        =   11536
            _ExtentY        =   503
            MaxLength       =   0
            DataType        =   4
            IdName          =   "WDRSETSCDSEL"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   86
            Left            =   5535
            TabIndex        =   66
            ToolTipText     =   "Enter To Catalog"
            Top             =   7395
            Width           =   2505
            _ExtentX        =   4419
            _ExtentY        =   503
            MaxLength       =   16
            DataType        =   4
            IdName          =   "WDPCOLLTO"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   88
            Left            =   8070
            TabIndex        =   67
            ToolTipText     =   "Enter Catalog Selection"
            Top             =   7395
            Width           =   6540
            _ExtentX        =   11536
            _ExtentY        =   503
            MaxLength       =   0
            DataType        =   4
            IdName          =   "WDPCOLLSEL"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   89
            Left            =   5535
            TabIndex        =   72
            ToolTipText     =   "Enter To Set"
            Top             =   7965
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "####0"
            MaxLength       =   5
            DataType        =   1
            IdName          =   "WDPSETTO"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   92
            Left            =   5535
            TabIndex        =   69
            ToolTipText     =   "Enter To Series"
            Top             =   7680
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "####0"
            MaxLength       =   5
            DataType        =   1
            IdName          =   "WDPSERTO"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   94
            Left            =   5535
            TabIndex        =   75
            ToolTipText     =   "Enter To Version"
            Top             =   8250
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "####0"
            MaxLength       =   5
            DataType        =   1
            IdName          =   "WDPVERTO"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   95
            Left            =   8070
            TabIndex        =   76
            ToolTipText     =   "Enter Version Selection"
            Top             =   8250
            Width           =   6540
            _ExtentX        =   11536
            _ExtentY        =   503
            MaxLength       =   0
            DataType        =   4
            IdName          =   "WDPVERSEL"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   96
            Left            =   8070
            TabIndex        =   73
            ToolTipText     =   "Enter Set Selection"
            Top             =   7965
            Width           =   6540
            _ExtentX        =   11536
            _ExtentY        =   503
            MaxLength       =   0
            DataType        =   4
            IdName          =   "WDPSETSEL"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   97
            Left            =   8070
            TabIndex        =   70
            ToolTipText     =   "Enter Series Selection"
            Top             =   7680
            Width           =   6540
            _ExtentX        =   11536
            _ExtentY        =   503
            MaxLength       =   0
            DataType        =   4
            IdName          =   "WDPSERSEL"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   20
            Left            =   2985
            TabIndex        =   63
            ToolTipText     =   "Enter From Customer "
            Top             =   7110
            Visible         =   0   'False
            Width           =   1065
            _ExtentX        =   1879
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   4
            IdName          =   "WCMCDFR"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   87
            Left            =   2985
            TabIndex        =   74
            ToolTipText     =   "Enter From Version"
            Top             =   8250
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "####0"
            MaxLength       =   5
            DataType        =   1
            IdName          =   "WDPVERFR"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   90
            Left            =   2985
            TabIndex        =   71
            ToolTipText     =   "Enter From Set"
            Top             =   7965
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "####0"
            MaxLength       =   5
            DataType        =   1
            IdName          =   "WDPSETFR"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   91
            Left            =   2985
            TabIndex        =   68
            ToolTipText     =   "Enter From Series"
            Top             =   7680
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "####0"
            MaxLength       =   5
            DataType        =   1
            IdName          =   "WDPSERFR"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   93
            Left            =   2985
            TabIndex        =   65
            ToolTipText     =   "Enter From Catalog"
            Top             =   7395
            Width           =   2505
            _ExtentX        =   4419
            _ExtentY        =   503
            MaxLength       =   16
            DataType        =   4
            IdName          =   "WDPCOLLFR"
         End
         Begin VB.Frame FraOptDmHld 
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   2970
            TabIndex        =   246
            ToolTipText     =   "Select if to Show Design On Hold - 'Yes', 'No', 'Both'"
            Top             =   5730
            Width           =   6945
            Begin VB.OptionButton OptDmHld 
               Caption         =   "Yes"
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
               ToolTipText     =   "Select if to Show Design On Hold - 'Yes', 'No', 'Both'"
               Top             =   0
               Width           =   1545
            End
            Begin VB.OptionButton OptDmHld 
               Caption         =   "Both"
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
               Left            =   5085
               TabIndex        =   49
               ToolTipText     =   "Select if to Show Design On Hold - 'Yes', 'No', 'Both'"
               Top             =   0
               Width           =   1665
            End
            Begin VB.OptionButton OptDmHld 
               Caption         =   "No"
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
               Left            =   2550
               TabIndex        =   48
               ToolTipText     =   "Select if to Show Design On Hold - 'Yes', 'No', 'Both'"
               Top             =   0
               Width           =   2055
            End
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   105
            Left            =   11715
            TabIndex        =   249
            ToolTipText     =   "Show Valid Designs (Y/N/' ')"
            Top             =   6020
            Visible         =   0   'False
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   503
            Enabled         =   0   'False
            DataType        =   4
            IdName          =   "WVALIDYN"
            CmpStr          =   "DmValidYN="
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   107
            Left            =   5535
            TabIndex        =   20
            ToolTipText     =   "Enter To Old Design Code"
            Top             =   2265
            Width           =   2385
            _ExtentX        =   4207
            _ExtentY        =   503
            MaxLength       =   20
            IdName          =   "WDMOLDCDTO"
            CmpStr          =   "DmOldCd <="
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   108
            Left            =   2985
            TabIndex        =   19
            ToolTipText     =   "Enter From Old Design Code"
            Top             =   2265
            Width           =   2385
            _ExtentX        =   4207
            _ExtentY        =   503
            MaxLength       =   20
            IdName          =   "WDMOLDCDFR"
            CmpStr          =   "DmOldCd >="
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   109
            Left            =   8070
            TabIndex        =   21
            ToolTipText     =   "EnterOld  Design Code Selection"
            Top             =   2265
            Width           =   6540
            _ExtentX        =   11536
            _ExtentY        =   503
            MaxLength       =   255
            IdName          =   "WDMOLDCDSEL"
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Old Design Code"
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
            Index           =   66
            Left            =   90
            TabIndex        =   253
            Top             =   2265
            Width           =   1605
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Design Valid"
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
            Left            =   90
            TabIndex        =   251
            Top             =   6020
            Width           =   1455
         End
         Begin VB.Label LblVer 
            BackStyle       =   0  'Transparent
            Caption         =   "Version"
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
            TabIndex        =   238
            Top             =   8295
            Width           =   1365
         End
         Begin VB.Label LblSet 
            BackStyle       =   0  'Transparent
            Caption         =   "Set"
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
            TabIndex        =   237
            Top             =   8010
            Width           =   1365
         End
         Begin VB.Label LblSer 
            BackStyle       =   0  'Transparent
            Caption         =   "Series"
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
            TabIndex        =   236
            Top             =   7725
            Width           =   1365
         End
         Begin VB.Label LblColl 
            BackStyle       =   0  'Transparent
            Caption         =   "Collection"
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
            TabIndex        =   235
            Top             =   7440
            Width           =   1365
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Setting Code"
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
            Index           =   69
            Left            =   90
            TabIndex        =   231
            Top             =   3135
            Width           =   1380
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Design Created Date"
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
            Index           =   63
            Left            =   90
            TabIndex        =   230
            Top             =   5145
            Width           =   2190
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
            Left            =   2640
            TabIndex        =   184
            Top             =   270
            Width           =   195
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
            Left            =   5520
            TabIndex        =   183
            Top             =   240
            Width           =   240
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Dm Typ"
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
            Left            =   90
            TabIndex        =   182
            Top             =   270
            Width           =   1875
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Dm Ctg"
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
            Left            =   90
            TabIndex        =   181
            Top             =   555
            Width           =   1965
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Prd Ctg"
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
            Left            =   90
            TabIndex        =   180
            Top             =   840
            Width           =   1935
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Sal Ctg"
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
            Index           =   7
            Left            =   90
            TabIndex        =   179
            Top             =   1125
            Width           =   1845
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
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   10
            Left            =   90
            TabIndex        =   178
            Top             =   2565
            Width           =   1365
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
            Index           =   9
            Left            =   90
            TabIndex        =   177
            Top             =   7110
            Visible         =   0   'False
            Width           =   1365
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Free/Regn Exclusive/ Not Exclusive"
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
            Left            =   90
            TabIndex        =   176
            Top             =   6590
            Visible         =   0   'False
            Width           =   2805
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Design Date"
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
            Left            =   90
            TabIndex        =   175
            Top             =   5445
            Width           =   1410
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Main Met Wt"
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
            Left            =   90
            TabIndex        =   174
            Top             =   3435
            Width           =   1380
         End
         Begin VB.Label ALBL 
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
            Index           =   115
            Left            =   90
            TabIndex        =   173
            Top             =   3720
            Width           =   1230
         End
         Begin VB.Label Label1 
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
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   0
            Left            =   90
            TabIndex        =   172
            Top             =   5730
            Width           =   1455
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Design Color"
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
            Left            =   90
            TabIndex        =   171
            Top             =   2850
            Width           =   1365
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Set Family Code"
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
            Index           =   2
            Left            =   90
            TabIndex        =   170
            Top             =   1980
            Width           =   1605
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Show Prd Rm Qty && Wt"
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
            TabIndex        =   169
            Top             =   6295
            Width           =   2325
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Dia Wt."
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
            Left            =   90
            TabIndex        =   168
            Top             =   4005
            Width           =   1365
         End
         Begin VB.Label LblDmSz 
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
            Left            =   90
            TabIndex        =   167
            Top             =   1695
            Width           =   1365
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
            Left            =   90
            TabIndex        =   166
            Top             =   1410
            Width           =   1365
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "CS Wt."
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
            Left            =   90
            TabIndex        =   165
            Top             =   4575
            Width           =   1365
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "CS Qty."
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
            Left            =   90
            TabIndex        =   164
            Top             =   4860
            Width           =   1365
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Dia Qty."
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
            Left            =   90
            TabIndex        =   163
            Top             =   4290
            Width           =   1365
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
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
            Height          =   240
            Index           =   0
            Left            =   8070
            TabIndex        =   162
            Top             =   270
            Width           =   900
         End
      End
      Begin VB.Frame FraSc 
         Height          =   2895
         Left            =   2340
         TabIndex        =   160
         Top             =   6180
         Visible         =   0   'False
         Width           =   9855
         Begin VB.CheckBox ChkBoxArr 
            Alignment       =   1  'Right Justify
            Caption         =   "Show Dsgs With Below 'D','C' SCtg Only"
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
            Left            =   4920
            TabIndex        =   247
            ToolTipText     =   "Check To Show Show Designs Which Have Diamonds Given In Scope Only"
            Top             =   240
            Width           =   4185
         End
         Begin MwfCtl.MWCTL_FLX GrdDsgSz 
            Height          =   2130
            Left            =   4755
            TabIndex        =   79
            Top             =   600
            Width           =   4920
            _ExtentX        =   8678
            _ExtentY        =   3757
            Cols            =   6
            colname1        =   "UODRMCTG"
            heading1        =   "Ctg"
            datafld1        =   "UodRmCtg"
            datatype1       =   4
            recalcparent1   =   "GrdDsgSz(UODPTRYN),GrdDsgSz(UODRMSCTG)"
            maxlength1      =   1
            colname2        =   "UODRMSCTG"
            heading2        =   "SubCtg "
            datafld2        =   "UodRmSCtg"
            datatype2       =   4
            recalcon2       =   "GrdDsgSz(UODRMCTG)"
            recalcparent2   =   "GrdDsgSz(UODPTRYN)"
            maxlength2      =   5
            tooltiptext2    =   "Enter Rm Sub Category"
            colname3        =   "UODPTRYN"
            heading3        =   "Ptr "
            datafld3        =   "UodPtrYN"
            datatype3       =   4
            recalcon3       =   "GrdDsgSz(UODRMSCTG),GrdDsgSz(UODRMCTG)"
            maxlength3      =   1
            style3          =   2
            colname4        =   "UODLNSVPTRFR"
            heading4        =   "Fr Ln/Sv/Ptr "
            datafld4        =   "UodLnSvPtrFr"
            datatype4       =   2
            mask4           =   "##0.0000"
            maxlength4      =   8
            tooltiptext4    =   "Enter From Length/ Sieve/ Pointer"
            colname5        =   "UODLNSVPTRTO"
            heading5        =   "To Ln/Sv/Ptr "
            datafld5        =   "UodLnSvPtrTo"
            datatype5       =   2
            mask5           =   "##0.0000"
            maxlength5      =   8
            tooltiptext5    =   "Enter To Length/ Sieve/ Pointer"
         End
         Begin MwfCtl.MWCTL_FLX GrdDsgAnaScop 
            Height          =   2130
            Left            =   150
            TabIndex        =   78
            Top             =   600
            Width           =   4485
            _ExtentX        =   7911
            _ExtentY        =   3757
            Cols            =   4
            colname1        =   "UDANASR"
            heading1        =   "Sr. "
            datafld1        =   "UdAnaSr"
            datatype1       =   4
            recalcparent1   =   "GrdDsgAnaScop(UDANACDFR),GrdDsgAnaScop(UDANACDTO)"
            maxlength1      =   2
            tooltiptext1    =   "Enter Analysis Sr. Number"
            colname2        =   "UDANACDFR"
            heading2        =   "Fr Ana Cd      "
            datafld2        =   "UdAnaCdFr"
            datatype2       =   4
            recalcon2       =   "GrdDsgAnaScop(UDANASR)"
            maxlength2      =   8
            tooltiptext2    =   "Enter From Analysis Code"
            colname3        =   "UDANACDTO"
            heading3        =   "To Ana Cd      "
            datafld3        =   "UdAnaCdTo"
            datatype3       =   4
            recalcon3       =   "GrdDsgAnaScop(UDANASR)"
            maxlength3      =   8
            tooltiptext3    =   "Enter To Analysis Code"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   104
            Left            =   9090
            TabIndex        =   248
            ToolTipText     =   "Show Designs Which Have Diamonds Given In Scope Only (Y/N)"
            Top             =   240
            Visible         =   0   'False
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   503
            Enabled         =   0   'False
            DataType        =   4
            DataField       =   "UOYN4"
            IdName          =   "UOYN"
         End
      End
   End
   Begin VB.Frame FRA_NKEYS 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   9180
      Left            =   60
      TabIndex        =   83
      Top             =   30
      Width           =   15135
      Begin VB.Frame FraNKeyAll 
         BorderStyle     =   0  'None
         Height          =   9135
         Left            =   -150
         TabIndex        =   84
         Top             =   30
         Width           =   15150
         Begin MwfCtl.MWCTL_FLX GrdDsgLst 
            Height          =   4305
            Left            =   270
            TabIndex        =   85
            Top             =   120
            Width           =   12870
            _ExtentX        =   22701
            _ExtentY        =   7594
            Cols            =   40
            AllowAdd        =   0   'False
            AllowDelete     =   0   'False
            colname1        =   "DMTCTYP"
            heading1        =   "Dm Typ."
            datatype1       =   4
            maxlength1      =   3
            style1          =   0
            colname2        =   "DMCTG"
            heading2        =   "Dsg Ctg  "
            datatype2       =   4
            maxlength2      =   5
            colname3        =   "DMCD"
            heading3        =   "Design Code                      "
            datatype3       =   4
            maxlength3      =   15
            colname4        =   "DMSZ"
            heading4        =   "Size       "
            datatype4       =   4
            maxlength4      =   5
            colname5        =   "DMUOM"
            heading5        =   "UOM          "
            maxlength5      =   5
            style5          =   0
            colname6        =   "WDIAWT"
            heading6        =   "Dia Wt.   "
            datatype6       =   2
            mask6           =   "#0.000"
            maxlength6      =   6
            tooltiptext6    =   "Diamond Weight"
            colname7        =   "WGRSWT"
            heading7        =   "Grs Wt.           "
            datatype7       =   2
            mask7           =   "####0.000"
            maxlength7      =   9
            colname8        =   "DMDESC"
            heading8        =   "Description                                                                                                  "
            maxlength8      =   45
            colname9        =   "DMPRTKEY"
            heading9        =   "Prtn "
            maxlength9      =   1
            style9          =   0
            colname10       =   "DMPRDCTG"
            heading10       =   " "
            datatype10      =   4
            maxlength10     =   5
            style10         =   0
            colname11       =   "DMSALCTG"
            heading11       =   " "
            datatype11      =   4
            maxlength11     =   5
            style11         =   0
            colname12       =   "DMVACTG"
            heading12       =   " "
            datatype12      =   4
            maxlength12     =   5
            style12         =   0
            colname13       =   "DMLSCTG"
            heading13       =   " "
            datatype13      =   4
            maxlength13     =   5
            style13         =   0
            colname14       =   "DMKT"
            heading14       =   " "
            datatype14      =   4
            maxlength14     =   5
            style14         =   0
            colname15       =   "DMBAGPCS"
            heading15       =   " "
            datatype15      =   1
            mask15          =   "###0"
            maxlength15     =   4
            style15         =   0
            colname16       =   "DMSETCD"
            heading16       =   " "
            datatype16      =   4
            maxlength16     =   15
            style16         =   0
            colname17       =   "DMOLDCD"
            heading17       =   " "
            datatype17      =   4
            maxlength17     =   20
            style17         =   0
            colname18       =   "DMCOL"
            heading18       =   " "
            datatype18      =   4
            maxlength18     =   5
            style18         =   0
            colname19       =   "DMCMCD"
            datatype19      =   4
            maxlength19     =   8
            style19         =   0
            colname20       =   "DMREGNEXC"
            datatype20      =   4
            maxlength20     =   1
            style20         =   0
            colname21       =   "DMPARTS"
            datatype21      =   1
            mask21          =   "##0"
            maxlength21     =   3
            style21         =   0
            colname22       =   "DMPARTDESC"
            datatype22      =   4
            maxlength22     =   45
            style22         =   0
            colname23       =   "DMHLD"
            datatype23      =   4
            maxlength23     =   1
            style23         =   0
            colname24       =   "DMHLDDESC"
            datatype24      =   4
            maxlength24     =   45
            style24         =   0
            colname25       =   "DMPRDSEQ"
            datatype25      =   4
            maxlength25     =   5
            style25         =   0
            colname26       =   "DMPRDINST"
            datatype26      =   4
            maxlength26     =   250
            style26         =   0
            colname27       =   "DMDSGDT"
            datatype27      =   3
            maxlength27     =   8
            style27         =   0
            colname28       =   "DMDSGBY"
            datatype28      =   4
            maxlength28     =   20
            style28         =   0
            colname29       =   "DMMODMKR"
            datatype29      =   4
            maxlength29     =   20
            style29         =   0
            colname30       =   "DMLSTMDF"
            datatype30      =   4
            maxlength30     =   45
            style30         =   0
            colname31       =   "DMWAXWT"
            datatype31      =   2
            mask31          =   "####0.000"
            maxlength31     =   9
            style31         =   0
            colname32       =   "DMMODRUNWT"
            datatype32      =   2
            mask32          =   "####0.000"
            maxlength32     =   9
            style32         =   0
            colname33       =   "DMSILMODWT"
            datatype33      =   2
            mask33          =   "####0.000"
            maxlength33     =   9
            style33         =   0
            colname34       =   "DMCASPCWT"
            datatype34      =   2
            mask34          =   "####0.000"
            maxlength34     =   9
            style34         =   0
            colname35       =   "DMHISTORY"
            maxlength35     =   500
            style35         =   0
            colname36       =   "DMFAULTS"
            maxlength36     =   500
            style36         =   0
            colname37       =   "DMCREATEDDT"
            datatype37      =   3
            maxlength37     =   8
            style37         =   0
            colname38       =   "WSEL"
            heading38       =   "Sel"
            datatype38      =   4
            maxlength38     =   1
            tooltiptext38   =   "Y for Select design and Copy into Catalog"
            colname39       =   "DMVALIDYN"
            datatype39      =   4
            maxlength39     =   1
            style39         =   0
         End
         Begin VB.Frame FraPic3D 
            BorderStyle     =   0  'None
            Height          =   1635
            Left            =   13380
            TabIndex        =   88
            Top             =   120
            Width           =   1725
            Begin VB.Image Pic3D 
               Height          =   1575
               Left            =   0
               Top             =   0
               Width           =   1665
            End
         End
         Begin TabDlg.SSTab TabDesgnDet 
            Height          =   4635
            Left            =   240
            TabIndex        =   90
            Top             =   4500
            Width           =   14865
            _ExtentX        =   26220
            _ExtentY        =   8176
            _Version        =   393216
            Tabs            =   11
            Tab             =   6
            TabsPerRow      =   11
            TabHeight       =   520
            TabCaption(0)   =   "<&1> General"
            TabPicture(0)   =   "EmrFrmDsgView.frx":0000
            Tab(0).ControlEnabled=   0   'False
            Tab(0).Control(0)=   "FraDesgnDet(0)"
            Tab(0).ControlCount=   1
            TabCaption(1)   =   "<&2>Bill Of Material"
            TabPicture(1)   =   "EmrFrmDsgView.frx":001C
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "FraDesgnDet(1)"
            Tab(1).ControlCount=   1
            TabCaption(2)   =   "<&3> Lab Details"
            TabPicture(2)   =   "EmrFrmDsgView.frx":0038
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "FraDesgnDet(2)"
            Tab(2).ControlCount=   1
            TabCaption(3)   =   "<&4> Analysis"
            TabPicture(3)   =   "EmrFrmDsgView.frx":0054
            Tab(3).ControlEnabled=   0   'False
            Tab(3).Control(0)=   "FraDesgnDet(3)"
            Tab(3).ControlCount=   1
            TabCaption(4)   =   "<&5> History"
            TabPicture(4)   =   "EmrFrmDsgView.frx":0070
            Tab(4).ControlEnabled=   0   'False
            Tab(4).Control(0)=   "FraDesgnDet(4)"
            Tab(4).ControlCount=   1
            TabCaption(5)   =   "<&6> Faults"
            TabPicture(5)   =   "EmrFrmDsgView.frx":008C
            Tab(5).ControlEnabled=   0   'False
            Tab(5).Control(0)=   "FraDesgnDet(5)"
            Tab(5).ControlCount=   1
            TabCaption(6)   =   "<&7> Comp Details"
            TabPicture(6)   =   "EmrFrmDsgView.frx":00A8
            Tab(6).ControlEnabled=   -1  'True
            Tab(6).Control(0)=   "FraDesgnDet(6)"
            Tab(6).Control(0).Enabled=   0   'False
            Tab(6).ControlCount=   1
            TabCaption(7)   =   "<&8> Model Avl."
            TabPicture(7)   =   "EmrFrmDsgView.frx":00C4
            Tab(7).ControlEnabled=   0   'False
            Tab(7).Control(0)=   "FraDesgnDet(7)"
            Tab(7).ControlCount=   1
            TabCaption(8)   =   "<&9> RM Summary"
            TabPicture(8)   =   "EmrFrmDsgView.frx":00E0
            Tab(8).ControlEnabled=   0   'False
            Tab(8).Control(0)=   "FraDesgnDet(8)"
            Tab(8).ControlCount=   1
            TabCaption(9)   =   "<1&0> Param"
            TabPicture(9)   =   "EmrFrmDsgView.frx":00FC
            Tab(9).ControlEnabled=   0   'False
            Tab(9).Control(0)=   "FraDesgnDet(9)"
            Tab(9).ControlCount=   1
            TabCaption(10)  =   "<&a> Allow"
            TabPicture(10)  =   "EmrFrmDsgView.frx":0118
            Tab(10).ControlEnabled=   0   'False
            Tab(10).Control(0)=   "FraDesgnDet(10)"
            Tab(10).ControlCount=   1
            Begin VB.Frame FraDesgnDet 
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               Height          =   4215
               Index           =   10
               Left            =   -74940
               TabIndex        =   254
               Top             =   360
               Width           =   14685
               Begin MwfCtl.MWCTL_FLX GrdDsgAllow 
                  Height          =   4110
                  Left            =   75
                  TabIndex        =   255
                  Top             =   60
                  Width           =   7155
                  _ExtentX        =   12621
                  _ExtentY        =   7250
                  Cols            =   6
                  colname1        =   "DMWTYP"
                  heading1        =   "Type          "
                  datafld1        =   "DmwTyp"
                  datatype1       =   4
                  recalcparent1   =   "GrdDsgAllow(DMWCMCD),GrdDsgAllow(DMWCMGRP),GrdDsgAllow(DMWREG),GrdDsgAllow(DMWYN)"
                  maxlength1      =   8
                  tooltiptext1    =   "Enter Type"
                  colname2        =   "DMWCMCD"
                  heading2        =   "Customer                      "
                  datafld2        =   "DmwCmCd"
                  datatype2       =   4
                  recalcon2       =   "GrdDsgAllow(DMWTYP)"
                  maxlength2      =   8
                  tooltiptext2    =   "Enter Customer Code"
                  colname3        =   "DMWCMGRP"
                  heading3        =   "Group                       "
                  datafld3        =   "DmwCmGrp"
                  datatype3       =   4
                  recalcon3       =   "GrdDsgAllow(DMWTYP)"
                  maxlength3      =   8
                  tooltiptext3    =   "Enter Group"
                  colname4        =   "DMWREG"
                  heading4        =   "Region              "
                  datafld4        =   "DmwReg"
                  datatype4       =   4
                  recalcon4       =   "GrdDsgAllow(DMWTYP)"
                  maxlength4      =   5
                  tooltiptext4    =   "Enter Region"
                  colname5        =   "DMWYN"
                  heading5        =   "YN  "
                  datafld5        =   "DmwYN"
                  datatype5       =   4
                  recalcon5       =   "GrdDsgAllow(DMWTYP)"
                  maxlength5      =   1
                  tooltiptext5    =   "Enter Allow Yes/No"
               End
            End
            Begin VB.Frame FraDesgnDet 
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               Height          =   4215
               Index           =   9
               Left            =   -74970
               TabIndex        =   221
               Top             =   330
               Width           =   14265
               Begin VB.Frame FraPicDP 
                  BorderStyle     =   0  'None
                  Height          =   4035
                  Left            =   7380
                  TabIndex        =   232
                  Top             =   0
                  Width           =   4305
                  Begin VB.Image PicDP 
                     Height          =   3990
                     Left            =   60
                     Top             =   30
                     Width           =   4020
                  End
               End
               Begin MwfCtl.MWCTL_FLX GrdDsgPrm 
                  Height          =   4230
                  Left            =   0
                  TabIndex        =   222
                  Top             =   0
                  Width           =   3795
                  _ExtentX        =   6694
                  _ExtentY        =   7461
                  Cols            =   3
                  colname1        =   "DPTYP"
                  heading1        =   "Type     "
                  datafld1        =   "DpTyp"
                  datatype1       =   4
                  maxlength1      =   5
                  tooltiptext1    =   "Enter Design Param Type"
                  colname2        =   "DPCD"
                  heading2        =   "Code                                          "
                  datafld2        =   "DpCd"
                  datatype2       =   4
                  maxlength2      =   16
                  tooltiptext2    =   "Enter Code"
               End
               Begin MwfCtl.MWCTL_FLX GrdDsgPrmDet 
                  Height          =   4230
                  Left            =   4020
                  TabIndex        =   223
                  Top             =   0
                  Width           =   3375
                  _ExtentX        =   5953
                  _ExtentY        =   7461
                  Cols            =   3
                  colname1        =   "DPDMCD"
                  heading1        =   "Design Code                     "
                  datafld1        =   "DpDmCd"
                  datatype1       =   4
                  recalcparent1   =   "GrdDsgPrmDet(WDPSEL)"
                  maxlength1      =   15
                  tooltiptext1    =   "Enter Design Code"
                  colname2        =   "WDPSEL"
                  heading2        =   "Sel"
                  datatype2       =   4
                  recalcon2       =   "GrdDsgPrmDet(DPDMCD)"
                  maxlength2      =   1
                  tooltiptext2    =   "Y for Select design and Copy into Catalog"
               End
            End
            Begin VB.Frame FraDesgnDet 
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               Height          =   4215
               Index           =   8
               Left            =   -74940
               TabIndex        =   185
               Top             =   360
               Width           =   14685
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   61
                  Left            =   1695
                  TabIndex        =   186
                  ToolTipText     =   "Gross Quantity"
                  Top             =   1320
                  Width           =   645
                  _ExtentX        =   1138
                  _ExtentY        =   503
                  Alignment       =   2
                  Enabled         =   0   'False
                  Mask            =   "####0"
                  MaxLength       =   5
                  DataType        =   2
                  IdName          =   "WGRSQTY"
               End
               Begin MwfCtl.MWCTL_FLX GrdSumm 
                  Height          =   3360
                  Left            =   4050
                  TabIndex        =   187
                  Top             =   510
                  Width           =   3975
                  _ExtentX        =   7011
                  _ExtentY        =   5927
                  Cols            =   5
                  colname1        =   "WRMCTG"
                  heading1        =   "Ctg"
                  datatype1       =   4
                  maxlength1      =   2
                  style1          =   2
                  colname2        =   "WRMSCTG"
                  heading2        =   "Sub Ctg"
                  datatype2       =   4
                  maxlength2      =   5
                  style2          =   2
                  colname3        =   "WQTY"
                  heading3        =   "Qty     "
                  datatype3       =   1
                  mask3           =   "####0"
                  maxlength3      =   5
                  style3          =   2
                  colname4        =   "WWT"
                  heading4        =   "Wt.           "
                  datatype4       =   2
                  mask4           =   "-###0.000"
                  maxlength4      =   9
                  style4          =   2
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   62
                  Left            =   2370
                  TabIndex        =   188
                  ToolTipText     =   "Gold Weight"
                  Top             =   1605
                  Width           =   975
                  _ExtentX        =   1720
                  _ExtentY        =   503
                  Alignment       =   2
                  Enabled         =   0   'False
                  Mask            =   "-###0.000"
                  MaxLength       =   9
                  DataType        =   2
                  IdName          =   "WGLDWT"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   63
                  Left            =   2370
                  TabIndex        =   189
                  ToolTipText     =   "Other Metal Weight"
                  Top             =   2745
                  Width           =   975
                  _ExtentX        =   1720
                  _ExtentY        =   503
                  Alignment       =   2
                  Enabled         =   0   'False
                  Mask            =   "-###0.000"
                  MaxLength       =   9
                  DataType        =   2
                  IdName          =   "WOTHMTWT"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   64
                  Left            =   2370
                  TabIndex        =   190
                  ToolTipText     =   "Diamond Weight"
                  Top             =   3030
                  Width           =   975
                  _ExtentX        =   1720
                  _ExtentY        =   503
                  Alignment       =   2
                  Enabled         =   0   'False
                  Mask            =   "-###0.000"
                  MaxLength       =   9
                  DataType        =   2
                  IdName          =   "WDIAWT"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   65
                  Left            =   2370
                  TabIndex        =   191
                  ToolTipText     =   "Col Stone Weight"
                  Top             =   3315
                  Width           =   975
                  _ExtentX        =   1720
                  _ExtentY        =   503
                  Alignment       =   2
                  Enabled         =   0   'False
                  Mask            =   "-###0.000"
                  MaxLength       =   9
                  DataType        =   2
                  IdName          =   "WCSWT"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   66
                  Left            =   2370
                  TabIndex        =   192
                  ToolTipText     =   "Accessory Weight"
                  Top             =   3600
                  Width           =   975
                  _ExtentX        =   1720
                  _ExtentY        =   503
                  Alignment       =   2
                  Enabled         =   0   'False
                  Mask            =   "-###0.000"
                  MaxLength       =   9
                  DataType        =   2
                  IdName          =   "WACCWT"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   67
                  Left            =   2370
                  TabIndex        =   193
                  ToolTipText     =   "Platinum Weight"
                  Top             =   1890
                  Width           =   975
                  _ExtentX        =   1720
                  _ExtentY        =   503
                  Alignment       =   2
                  Enabled         =   0   'False
                  Mask            =   "-###0.000"
                  MaxLength       =   9
                  DataType        =   2
                  IdName          =   "WPLWT"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   68
                  Left            =   2370
                  TabIndex        =   194
                  ToolTipText     =   "Gross Weight"
                  Top             =   1320
                  Width           =   975
                  _ExtentX        =   1720
                  _ExtentY        =   503
                  Alignment       =   2
                  Enabled         =   0   'False
                  Mask            =   "-###0.000"
                  MaxLength       =   9
                  DataType        =   2
                  IdName          =   "WGRSWT"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   69
                  Left            =   1695
                  TabIndex        =   195
                  ToolTipText     =   "Diamond Quantity"
                  Top             =   3030
                  Width           =   645
                  _ExtentX        =   1138
                  _ExtentY        =   503
                  Alignment       =   2
                  Enabled         =   0   'False
                  Mask            =   "####0"
                  MaxLength       =   5
                  DataType        =   2
                  IdName          =   "WDIAQTY"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   70
                  Left            =   1695
                  TabIndex        =   196
                  ToolTipText     =   "Col Stone Quantity"
                  Top             =   3315
                  Width           =   645
                  _ExtentX        =   1138
                  _ExtentY        =   503
                  Alignment       =   2
                  Enabled         =   0   'False
                  Mask            =   "####0"
                  MaxLength       =   5
                  DataType        =   2
                  IdName          =   "WCSQTY"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   71
                  Left            =   1695
                  TabIndex        =   197
                  ToolTipText     =   "Accessory Quantity"
                  Top             =   3600
                  Width           =   645
                  _ExtentX        =   1138
                  _ExtentY        =   503
                  Alignment       =   2
                  Enabled         =   0   'False
                  Mask            =   "####0"
                  MaxLength       =   5
                  DataType        =   2
                  IdName          =   "WACCQTY"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   72
                  Left            =   1695
                  TabIndex        =   198
                  ToolTipText     =   "Karatage"
                  Top             =   510
                  Width           =   855
                  _ExtentX        =   1508
                  _ExtentY        =   503
                  Enabled         =   0   'False
                  MaxLength       =   5
                  DataType        =   4
                  IdName          =   "WKT"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   73
                  Left            =   10860
                  TabIndex        =   199
                  ToolTipText     =   "Should The Production Rm Qty && Wt Be Shown In Dsg And Ord Entries"
                  Top             =   3570
                  Width           =   345
                  _ExtentX        =   609
                  _ExtentY        =   503
                  DataType        =   4
                  IdName          =   "WSHOWPRDRMDETS"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   74
                  Left            =   2370
                  TabIndex        =   200
                  ToolTipText     =   "Platinum Weight"
                  Top             =   2175
                  Width           =   975
                  _ExtentX        =   1720
                  _ExtentY        =   503
                  Alignment       =   2
                  Enabled         =   0   'False
                  Mask            =   "-###0.000"
                  MaxLength       =   9
                  DataType        =   2
                  IdName          =   "WSILWT"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   103
                  Left            =   2370
                  TabIndex        =   242
                  ToolTipText     =   "Palladium Weight"
                  Top             =   2460
                  Width           =   975
                  _ExtentX        =   1720
                  _ExtentY        =   503
                  Alignment       =   2
                  Enabled         =   0   'False
                  Mask            =   "-###0.000"
                  MaxLength       =   9
                  DataType        =   2
                  IdName          =   "WPDWT"
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "gms"
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
                  Index           =   72
                  Left            =   3390
                  TabIndex        =   243
                  Top             =   2190
                  Width           =   525
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Palladium"
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
                  Index           =   29
                  Left            =   300
                  TabIndex        =   241
                  Top             =   2460
                  Width           =   1425
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Gross"
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
                  Index           =   62
                  Left            =   300
                  TabIndex        =   220
                  Top             =   1320
                  Width           =   1425
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Gold"
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
                  Index           =   61
                  Left            =   300
                  TabIndex        =   219
                  Top             =   1605
                  Width           =   1425
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Oth Mt"
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
                  Left            =   300
                  TabIndex        =   218
                  Top             =   2745
                  Width           =   1425
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Dia"
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
                  Index           =   59
                  Left            =   300
                  TabIndex        =   217
                  Top             =   3030
                  Width           =   1425
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Col Stn"
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
                  Left            =   300
                  TabIndex        =   216
                  Top             =   3315
                  Width           =   1425
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Accessories"
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
                  Left            =   300
                  TabIndex        =   215
                  Top             =   3600
                  Width           =   1425
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
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
                  Index           =   33
                  Left            =   1830
                  TabIndex        =   214
                  Top             =   1020
                  Width           =   675
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
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
                  Index           =   34
                  Left            =   2670
                  TabIndex        =   213
                  Top             =   1020
                  Width           =   555
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Platinum"
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
                  Index           =   35
                  Left            =   300
                  TabIndex        =   212
                  Top             =   1890
                  Width           =   1425
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Karat"
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
                  Index           =   38
                  Left            =   300
                  TabIndex        =   211
                  Top             =   510
                  Width           =   855
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "gms"
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
                  Index           =   42
                  Left            =   3390
                  TabIndex        =   210
                  Top             =   1320
                  Width           =   525
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "gms"
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
                  Index           =   44
                  Left            =   3390
                  TabIndex        =   209
                  Top             =   1605
                  Width           =   525
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "gms"
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
                  Index           =   45
                  Left            =   3390
                  TabIndex        =   208
                  Top             =   3600
                  Width           =   525
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "cts"
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
                  Left            =   3390
                  TabIndex        =   207
                  Top             =   3315
                  Width           =   525
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "cts"
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
                  Left            =   3390
                  TabIndex        =   206
                  Top             =   3030
                  Width           =   525
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "gms"
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
                  Left            =   3390
                  TabIndex        =   205
                  Top             =   2745
                  Width           =   525
               End
               Begin VB.Label lblShBothWts 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Show Prd/ Sal wghts (Y/N)"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   345
                  Left            =   8070
                  TabIndex        =   204
                  Top             =   3540
                  Width           =   2745
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Silver"
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
                  Index           =   54
                  Left            =   300
                  TabIndex        =   203
                  Top             =   2175
                  Width           =   1425
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "gms"
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
                  Left            =   3390
                  TabIndex        =   202
                  Top             =   1890
                  Width           =   525
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "gms"
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
                  Left            =   3390
                  TabIndex        =   201
                  Top             =   2490
                  Width           =   525
               End
            End
            Begin VB.Frame FraDesgnDet 
               BorderStyle     =   0  'None
               Height          =   4215
               Index           =   0
               Left            =   -74940
               TabIndex        =   134
               Top             =   360
               Width           =   14685
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   19
                  Left            =   12765
                  TabIndex        =   96
                  Top             =   675
                  Width           =   465
                  _ExtentX        =   820
                  _ExtentY        =   503
                  Alignment       =   2
                  Mask            =   "##0"
                  MaxLength       =   3
                  DataType        =   1
                  IdName          =   "WDMBAGPCS"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   22
                  Left            =   6135
                  TabIndex        =   98
                  ToolTipText     =   "Enter Old Design Code"
                  Top             =   960
                  Width           =   2355
                  _ExtentX        =   4154
                  _ExtentY        =   503
                  MaxLength       =   20
                  DataType        =   4
                  IdName          =   "WDMOLDCD"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   24
                  Left            =   12375
                  TabIndex        =   99
                  ToolTipText     =   "Enter Colour"
                  Top             =   960
                  Width           =   855
                  _ExtentX        =   1508
                  _ExtentY        =   503
                  MaxLength       =   5
                  DataType        =   4
                  IdName          =   "WDMCOL"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   25
                  Left            =   1455
                  TabIndex        =   104
                  ToolTipText     =   "Enter The Number Of Parts"
                  Top             =   1260
                  Width           =   465
                  _ExtentX        =   820
                  _ExtentY        =   503
                  Alignment       =   2
                  Mask            =   "##0"
                  MaxLength       =   3
                  DataType        =   1
                  IdName          =   "WDMPARTS"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   35
                  Left            =   6135
                  TabIndex        =   105
                  ToolTipText     =   "Enter Description For Parts Of The Design"
                  Top             =   1260
                  Width           =   7095
                  _ExtentX        =   12515
                  _ExtentY        =   503
                  MaxLength       =   45
                  IdName          =   "WDMPARTDESC"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   37
                  Left            =   1455
                  TabIndex        =   106
                  ToolTipText     =   "Design On Hold Yes/No?"
                  Top             =   1545
                  Width           =   315
                  _ExtentX        =   556
                  _ExtentY        =   503
                  DataType        =   4
                  ReCalcParent    =   "WDMHLDDESC"
                  IdName          =   "WDMHLD"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   38
                  Left            =   6135
                  TabIndex        =   107
                  ToolTipText     =   "Enter The Reason For Keeping Design On Hold"
                  Top             =   1575
                  Width           =   7095
                  _ExtentX        =   12515
                  _ExtentY        =   503
                  MaxLength       =   45
                  ReCalcOn        =   "WDMHLD"
                  IdName          =   "WDMHLDDESC"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   795
                  Index           =   39
                  Left            =   6135
                  TabIndex        =   109
                  ToolTipText     =   "Enter Production Instruction"
                  Top             =   1860
                  Width           =   7095
                  _ExtentX        =   12515
                  _ExtentY        =   1402
                  Alignment       =   3
                  MaxLength       =   250
                  IdName          =   "WDMPRDINST"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   40
                  Left            =   6135
                  TabIndex        =   111
                  ToolTipText     =   "Enter The Name/Code Of Designer"
                  Top             =   2670
                  Width           =   2355
                  _ExtentX        =   4154
                  _ExtentY        =   503
                  MaxLength       =   20
                  IdName          =   "WDMDSGBY"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   41
                  Left            =   6135
                  TabIndex        =   113
                  ToolTipText     =   "Last Modified Date"
                  Top             =   2955
                  Width           =   7095
                  _ExtentX        =   12515
                  _ExtentY        =   503
                  MaxLength       =   45
                  IdName          =   "WDMLSTMDF"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   42
                  Left            =   6135
                  TabIndex        =   95
                  ToolTipText     =   "Enter Karatage"
                  Top             =   645
                  Width           =   855
                  _ExtentX        =   1508
                  _ExtentY        =   503
                  MaxLength       =   5
                  DataType        =   4
                  IdName          =   "WDMKT"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   43
                  Left            =   1455
                  TabIndex        =   97
                  ToolTipText     =   "Enter Set Code"
                  Top             =   960
                  Width           =   2130
                  _ExtentX        =   3757
                  _ExtentY        =   503
                  MaxLength       =   15
                  DataType        =   4
                  IdName          =   "WDMSETCD"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   44
                  Left            =   12900
                  TabIndex        =   103
                  ToolTipText     =   "Specify Whether The Design Is Exclusive For The Customer For A Region Or For All Regions (Enter 'Y' or 'N')"
                  Top             =   4005
                  Visible         =   0   'False
                  Width           =   315
                  _ExtentX        =   556
                  _ExtentY        =   503
                  DataType        =   4
                  IdName          =   "WDMREGNEXC"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   45
                  Left            =   1455
                  TabIndex        =   91
                  ToolTipText     =   "Enter Production Category"
                  Top             =   360
                  Width           =   855
                  _ExtentX        =   1508
                  _ExtentY        =   503
                  MaxLength       =   5
                  DataType        =   4
                  IdName          =   "WDMPRDCTG"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   47
                  Left            =   6135
                  TabIndex        =   92
                  ToolTipText     =   "Enter Sales Category"
                  Top             =   360
                  Width           =   855
                  _ExtentX        =   1508
                  _ExtentY        =   503
                  MaxLength       =   5
                  DataType        =   4
                  IdName          =   "WDMSALCTG"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   48
                  Left            =   1455
                  TabIndex        =   94
                  ToolTipText     =   "Enter Loss Category"
                  Top             =   645
                  Width           =   855
                  _ExtentX        =   1508
                  _ExtentY        =   503
                  MaxLength       =   5
                  DataType        =   4
                  IdName          =   "WDMLSCTG"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   49
                  Left            =   12375
                  TabIndex        =   93
                  ToolTipText     =   "Enter Value Addition Category"
                  Top             =   390
                  Width           =   855
                  _ExtentX        =   1508
                  _ExtentY        =   503
                  MaxLength       =   5
                  DataType        =   4
                  IdName          =   "WDMVACTG"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   50
                  Left            =   10815
                  TabIndex        =   102
                  ToolTipText     =   "Enter Customer Name"
                  Top             =   4005
                  Visible         =   0   'False
                  Width           =   165
                  _ExtentX        =   291
                  _ExtentY        =   503
                  MaxLength       =   45
                  ReCalcParent    =   "WDMCMCD"
                  ReCalcOn        =   "WDMCMCD"
                  IdName          =   "WCMNAME"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   51
                  Left            =   1455
                  TabIndex        =   100
                  ToolTipText     =   "Enter Customer Code"
                  Top             =   3975
                  Visible         =   0   'False
                  Width           =   8415
                  _ExtentX        =   14843
                  _ExtentY        =   503
                  MaxLength       =   80
                  DataType        =   4
                  ReCalcParent    =   "WCMREGNCD,WCMNAME"
                  ReCalcOn        =   "WCMNAME"
                  IdName          =   "WDMCMCD"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   52
                  Left            =   9870
                  TabIndex        =   101
                  ToolTipText     =   "Customer Region Code"
                  Top             =   3975
                  Visible         =   0   'False
                  Width           =   855
                  _ExtentX        =   1508
                  _ExtentY        =   503
                  Enabled         =   0   'False
                  MaxLength       =   5
                  DataType        =   4
                  ReCalcOn        =   "WDMCMCD"
                  IdName          =   "WCMREGNCD"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   53
                  Left            =   1455
                  TabIndex        =   108
                  ToolTipText     =   "Enter Production Sequence"
                  Top             =   1830
                  Width           =   1440
                  _ExtentX        =   2540
                  _ExtentY        =   503
                  MaxLength       =   8
                  DataType        =   4
                  IdName          =   "WDMPRDSEQ"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   54
                  Left            =   1455
                  TabIndex        =   110
                  ToolTipText     =   "Enter Design Date"
                  Top             =   2670
                  Width           =   885
                  _ExtentX        =   1561
                  _ExtentY        =   503
                  MaxLength       =   8
                  DataType        =   3
                  IdName          =   "WDMDSGDT"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   55
                  Left            =   1455
                  TabIndex        =   112
                  ToolTipText     =   "Enter Model Maker"
                  Top             =   2955
                  Width           =   2355
                  _ExtentX        =   4154
                  _ExtentY        =   503
                  MaxLength       =   20
                  IdName          =   "WDMMODMKR"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   56
                  Left            =   1455
                  TabIndex        =   114
                  ToolTipText     =   "Enter Wax Weight"
                  Top             =   3240
                  Width           =   1425
                  _ExtentX        =   2514
                  _ExtentY        =   503
                  Alignment       =   2
                  Mask            =   "###0.000"
                  MaxLength       =   8
                  DataType        =   2
                  IdName          =   "WDMWAXWT"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   57
                  Left            =   6135
                  TabIndex        =   115
                  ToolTipText     =   "Enter Model Weight With Runner"
                  Top             =   3240
                  Width           =   1425
                  _ExtentX        =   2514
                  _ExtentY        =   503
                  Alignment       =   2
                  Mask            =   "###0.000"
                  MaxLength       =   8
                  DataType        =   2
                  IdName          =   "WDMMODRUNWT"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   58
                  Left            =   11805
                  TabIndex        =   116
                  ToolTipText     =   "Enter Silver/Model Weight"
                  Top             =   3240
                  Width           =   1425
                  _ExtentX        =   2514
                  _ExtentY        =   503
                  Alignment       =   2
                  Mask            =   "###0.000"
                  MaxLength       =   8
                  DataType        =   2
                  IdName          =   "WDMSILMODWT"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   59
                  Left            =   1455
                  TabIndex        =   117
                  ToolTipText     =   "Enter Casting Piece Weight"
                  Top             =   3525
                  Width           =   1425
                  _ExtentX        =   2514
                  _ExtentY        =   503
                  Alignment       =   2
                  Mask            =   "###0.000"
                  MaxLength       =   8
                  DataType        =   2
                  IdName          =   "WDMCASPCWT"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   101
                  Left            =   12345
                  TabIndex        =   234
                  ToolTipText     =   "Enter Design Date"
                  Top             =   2670
                  Width           =   885
                  _ExtentX        =   1561
                  _ExtentY        =   503
                  Enabled         =   0   'False
                  MaxLength       =   8
                  DataType        =   3
                  IdName          =   "WDMCREATEDDT"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   102
                  Left            =   1455
                  TabIndex        =   239
                  ToolTipText     =   "Enter Production Category"
                  Top             =   60
                  Width           =   855
                  _ExtentX        =   1508
                  _ExtentY        =   503
                  MaxLength       =   5
                  DataType        =   4
                  IdName          =   "WDMUOM"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   106
                  Left            =   12915
                  TabIndex        =   62
                  ToolTipText     =   "Design Valid Yes/No?"
                  Top             =   3525
                  Width           =   315
                  _ExtentX        =   556
                  _ExtentY        =   503
                  DataType        =   4
                  ReCalcParent    =   "WDMHLDDESC"
                  IdName          =   "WDMVALIDYN"
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Valid YN"
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
                  Index           =   65
                  Left            =   11520
                  TabIndex        =   252
                  Top             =   3555
                  Width           =   1485
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "UOM"
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
                  Index           =   71
                  Left            =   60
                  TabIndex        =   240
                  Top             =   60
                  Width           =   855
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Created Date"
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
                  Index           =   70
                  Left            =   11040
                  TabIndex        =   233
                  Top             =   2670
                  Width           =   1485
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Old Cd"
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
                  Index           =   28
                  Left            =   4650
                  TabIndex        =   159
                  Top             =   960
                  Width           =   1485
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
                  Index           =   27
                  Left            =   11040
                  TabIndex        =   158
                  Top             =   945
                  Width           =   1485
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Parts"
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
                  Left            =   60
                  TabIndex        =   157
                  Top             =   1260
                  Width           =   1485
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Part Desc"
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
                  Index           =   25
                  Left            =   4650
                  TabIndex        =   156
                  Top             =   1290
                  Width           =   1545
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
                  Index           =   24
                  Left            =   60
                  TabIndex        =   155
                  Top             =   1545
                  Width           =   1485
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Holding Desc"
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
                  Index           =   22
                  Left            =   4650
                  TabIndex        =   154
                  Top             =   1575
                  Width           =   1545
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
                  Index           =   20
                  Left            =   60
                  TabIndex        =   153
                  Top             =   1830
                  Width           =   1485
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Prd Instruction"
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
                  Left            =   4650
                  TabIndex        =   152
                  Top             =   1860
                  Width           =   1545
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Designed By"
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
                  Left            =   4650
                  TabIndex        =   151
                  Top             =   2670
                  Width           =   1545
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Design Date"
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
                  TabIndex        =   150
                  Top             =   2670
                  Width           =   1485
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Model Maker"
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
                  Index           =   18
                  Left            =   60
                  TabIndex        =   149
                  Top             =   2955
                  Width           =   1485
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Last Modified"
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
                  Left            =   4650
                  TabIndex        =   148
                  Top             =   2955
                  Width           =   1545
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Karat"
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
                  Left            =   4650
                  TabIndex        =   147
                  Top             =   675
                  Width           =   1545
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Set Cd"
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
                  Index           =   40
                  Left            =   60
                  TabIndex        =   146
                  Top             =   930
                  Width           =   1485
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
                  Index           =   13
                  Left            =   11040
                  TabIndex        =   145
                  Top             =   645
                  Width           =   885
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Region Exc"
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
                  Index           =   41
                  Left            =   11040
                  TabIndex        =   144
                  Top             =   3990
                  Visible         =   0   'False
                  Width           =   1485
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
                  Index           =   36
                  Left            =   11040
                  TabIndex        =   143
                  Top             =   390
                  Width           =   1395
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Loss Ctg"
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
                  Index           =   39
                  Left            =   60
                  TabIndex        =   142
                  Top             =   675
                  Width           =   1005
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Prd Ctg"
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
                  Left            =   60
                  TabIndex        =   141
                  Top             =   360
                  Width           =   855
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Sales Ctg"
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
                  Left            =   4650
                  TabIndex        =   140
                  Top             =   390
                  Width           =   1065
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Cust/ Region"
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
                  Index           =   43
                  Left            =   60
                  TabIndex        =   139
                  Top             =   3975
                  Visible         =   0   'False
                  Width           =   1545
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Wax Wt"
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
                  Index           =   55
                  Left            =   60
                  TabIndex        =   138
                  Top             =   3240
                  Width           =   1485
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Model Wt With Runner"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   465
                  Index           =   56
                  Left            =   4650
                  TabIndex        =   137
                  Top             =   3240
                  Width           =   1455
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Silver/Model Wt"
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
                  Index           =   57
                  Left            =   10200
                  TabIndex        =   136
                  Top             =   3240
                  Width           =   1635
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Casting Pc Wt"
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
                  Index           =   58
                  Left            =   60
                  TabIndex        =   135
                  Top             =   3525
                  Width           =   1485
               End
            End
            Begin VB.Frame FraDesgnDet 
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               Height          =   4215
               Index           =   1
               Left            =   -74940
               TabIndex        =   132
               Top             =   360
               Width           =   14685
               Begin MwfCtl.MWCTL_FLX GrdDsgRm 
                  Height          =   4110
                  Left            =   75
                  TabIndex        =   133
                  Top             =   60
                  Width           =   14505
                  _ExtentX        =   25585
                  _ExtentY        =   7250
                  Cols            =   18
                  colname1        =   "DRSR"
                  heading1        =   "Sr.  "
                  datafld1        =   "DrSr"
                  datatype1       =   1
                  mask1           =   "##0"
                  maxlength1      =   3
                  colname2        =   "DRRMCTG"
                  heading2        =   "Ctg"
                  datafld2        =   "DrRmCtg"
                  datatype2       =   4
                  maxlength2      =   2
                  tooltiptext2    =   "Enter Raw Material Category"
                  colname3        =   "DRRMSCTG"
                  heading3        =   "SubCtg"
                  datafld3        =   "DrRmSCtg"
                  datatype3       =   4
                  maxlength3      =   5
                  tooltiptext3    =   "Enter Raw Material Sub Category"
                  colname4        =   "DRRMCD"
                  heading4        =   "Rm Code                                   "
                  datafld4        =   "DrRmCd"
                  datatype4       =   4
                  maxlength4      =   16
                  tooltiptext4    =   "Enter Rm Code"
                  colname5        =   "DRLN1"
                  heading5        =   "Len/Sv"
                  datafld5        =   "DrLn1"
                  datatype5       =   2
                  mask5           =   "#0.000"
                  maxlength5      =   6
                  tooltiptext5    =   "Enter the Length or Sieve Sz of RmCode"
                  colname6        =   "DRLN2"
                  heading6        =   "Brdth  "
                  datafld6        =   "DrLn2"
                  datatype6       =   2
                  mask6           =   "#0.000"
                  maxlength6      =   6
                  tooltiptext6    =   "Enter the Breadth of RmCode"
                  colname7        =   "DRPRDQTY"
                  heading7        =   "Prd Qty"
                  datafld7        =   "DrPrdQty"
                  datatype7       =   1
                  mask7           =   "###0"
                  maxlength7      =   4
                  tooltiptext7    =   "Enter Prd Qty for RmCode"
                  colname8        =   "DRRMPTR"
                  heading8        =   "Rm Ptr   "
                  datafld8        =   "DrRmPtr"
                  datatype8       =   2
                  mask8           =   "##0.0000"
                  maxlength8      =   8
                  tooltiptext8    =   "Enter Pointer Weight"
                  colname9        =   "DRPRDWT"
                  heading9        =   "Prd Wt     "
                  datafld9        =   "DrPrdWt"
                  datatype9       =   2
                  mask9           =   "-##0.000"
                  maxlength9      =   8
                  tooltiptext9    =   "Enter Prd Weight"
                  colname10       =   "DRQTY"
                  heading10       =   "Sal Qty"
                  datafld10       =   "DrQty"
                  datatype10      =   1
                  mask10          =   "###0"
                  maxlength10     =   4
                  tooltiptext10   =   "Enter Qty for RmCode"
                  colname11       =   "DRWT"
                  heading11       =   "Sal Wt     "
                  datafld11       =   "DrWt"
                  datatype11      =   2
                  mask11          =   "-##0.000"
                  maxlength11     =   8
                  tooltiptext11   =   "Enter Weight for RmCode"
                  colname12       =   "DRSETSCD"
                  heading12       =   "Set Cd           "
                  datafld12       =   "DrSetSCd"
                  datatype12      =   4
                  maxlength12     =   8
                  tooltiptext12   =   "Enter Setting Sub Code"
                  colname13       =   "DRWSQTY"
                  heading13       =   "WSet"
                  datafld13       =   "DrWsQty"
                  datatype13      =   1
                  mask13          =   "###0"
                  maxlength13     =   4
                  tooltiptext13   =   "Enter Wax Set Quantity"
                  colname14       =   "DRHSQTY"
                  heading14       =   "HSet"
                  datafld14       =   "DrHsQty"
                  datatype14      =   1
                  mask14          =   "###0"
                  maxlength14     =   4
                  style14         =   2
                  tooltiptext14   =   "Enter Hand Set Quantity"
                  colname15       =   "DRSUBSHP"
                  heading15       =   "SShp "
                  datafld15       =   "DrSubShp"
                  datatype15      =   4
                  maxlength15     =   3
                  tooltiptext15   =   "Enter Sub Shape"
                  colname16       =   "DRALYCD"
                  heading16       =   "Alloy                                         "
                  datafld16       =   "DrAlyCd"
                  datatype16      =   4
                  maxlength16     =   16
                  tooltiptext16   =   "Enter Alloy Code"
                  colname17       =   "DRMAINMET"
                  heading17       =   "MMet"
                  datafld17       =   "DrMainMet"
                  datatype17      =   4
                  maxlength17     =   1
                  tooltiptext17   =   "Enter Main Metal (Y/N)"
               End
            End
            Begin VB.Frame FraDesgnDet 
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               Height          =   4215
               Index           =   2
               Left            =   -74940
               TabIndex        =   130
               Top             =   360
               Width           =   14685
               Begin MwfCtl.MWCTL_FLX GrdDsgLab 
                  Height          =   4110
                  Left            =   75
                  TabIndex        =   131
                  Top             =   60
                  Width           =   5115
                  _ExtentX        =   9022
                  _ExtentY        =   7250
                  Cols            =   6
                  colname1        =   "DLSR"
                  heading1        =   "Sr.  "
                  datafld1        =   "DlSr"
                  datatype1       =   1
                  mask1           =   "#0"
                  maxlength1      =   2
                  style1          =   2
                  colname2        =   "DLMCD"
                  heading2        =   "Lab Cd   "
                  datafld2        =   "DlMCd"
                  datatype2       =   4
                  recalcparent2   =   "GrdDsgLab(WQW)"
                  maxlength2      =   5
                  tooltiptext2    =   "Enter Labour Main Code"
                  colname3        =   "DLSCD"
                  heading3        =   "Lab SCd        "
                  datafld3        =   "DLSCd"
                  datatype3       =   4
                  recalcparent3   =   "GrdDsgLab(WQW)"
                  maxlength3      =   8
                  tooltiptext3    =   "Enter Labour Sub Code"
                  colname4        =   "WQW"
                  heading4        =   "By Q/W"
                  datatype4       =   4
                  recalcon4       =   "GrdDsgLab(DLMCD),GrdDsgLab(DLSCD)"
                  maxlength4      =   5
                  style4          =   2
                  colname5        =   "DLQTY"
                  heading5        =   "Qty  "
                  datafld5        =   "DLQty"
                  datatype5       =   1
                  mask5           =   "##0"
                  maxlength5      =   3
                  tooltiptext5    =   "Enter Labour Quantity"
               End
            End
            Begin VB.Frame FraDesgnDet 
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               Height          =   4215
               Index           =   3
               Left            =   -74940
               TabIndex        =   128
               Top             =   360
               Width           =   14685
               Begin MwfCtl.MWCTL_FLX GrdDsgAna 
                  Height          =   4110
                  Left            =   75
                  TabIndex        =   129
                  Top             =   60
                  Width           =   8835
                  _ExtentX        =   15584
                  _ExtentY        =   7250
                  Cols            =   5
                  colname1        =   "DAANASR"
                  heading1        =   "Sr. "
                  datafld1        =   "DaAnaSr"
                  datatype1       =   4
                  recalcparent1   =   "GrdDsgAna(WANANAME)"
                  maxlength1      =   2
                  tooltiptext1    =   "Enter Analysis Serial No."
                  colname2        =   "WANANAME"
                  heading2        =   "Ana Name     "
                  recalcon2       =   "GrdDsgAna(DAANASR)"
                  maxlength2      =   10
                  style2          =   2
                  tooltiptext2    =   "Analysis Name"
                  colname3        =   "DAANACD"
                  heading3        =   "Ana Cd        "
                  datafld3        =   "DaAnaCd"
                  datatype3       =   4
                  maxlength3      =   8
                  tooltiptext3    =   "Enter Analysis Code"
                  colname4        =   "DAANADESC"
                  heading4        =   "Analysis Desc                                                                        "
                  datafld4        =   "DaAnaDesc"
                  maxlength4      =   45
                  tooltiptext4    =   "Enter Analysis Description"
               End
            End
            Begin VB.Frame FraDesgnDet 
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               Height          =   4215
               Index           =   4
               Left            =   -74940
               TabIndex        =   125
               Top             =   360
               Width           =   14685
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   3795
                  Index           =   17
                  Left            =   75
                  TabIndex        =   126
                  ToolTipText     =   "Enter History Of The Design"
                  Top             =   360
                  Width           =   14505
                  _ExtentX        =   25585
                  _ExtentY        =   6694
                  Alignment       =   3
                  MaxLength       =   0
                  IdName          =   "DMHISTORY"
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "History"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Index           =   21
                  Left            =   90
                  TabIndex        =   127
                  Top             =   60
                  Width           =   1095
               End
            End
            Begin VB.Frame FraDesgnDet 
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               Height          =   4215
               Index           =   5
               Left            =   -74940
               TabIndex        =   122
               Top             =   360
               Width           =   14685
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   3795
                  Index           =   46
                  Left            =   75
                  TabIndex        =   123
                  ToolTipText     =   "Enter Fault Description"
                  Top             =   360
                  Width           =   14505
                  _ExtentX        =   25585
                  _ExtentY        =   6694
                  Alignment       =   3
                  MaxLength       =   0
                  IdName          =   "DMFAULTS"
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Faults"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Index           =   50
                  Left            =   90
                  TabIndex        =   124
                  Top             =   60
                  Width           =   1095
               End
            End
            Begin VB.Frame FraDesgnDet 
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               Height          =   4215
               Index           =   6
               Left            =   60
               TabIndex        =   120
               Top             =   360
               Width           =   14685
               Begin MwfCtl.MWCTL_FLX GrdDsgCT 
                  Height          =   4110
                  Left            =   75
                  TabIndex        =   121
                  Top             =   60
                  Width           =   14505
                  _ExtentX        =   25585
                  _ExtentY        =   7250
                  Cols            =   13
                  colname1        =   "DCTSR"
                  heading1        =   "Sr  "
                  datafld1        =   "DctSr"
                  datatype1       =   1
                  mask1           =   "#0"
                  maxlength1      =   2
                  style1          =   2
                  tooltiptext1    =   "Serial Number"
                  colname2        =   "DCTCD"
                  heading2        =   "Comp Cd                     "
                  datafld2        =   "DctCd"
                  datatype2       =   4
                  maxlength2      =   15
                  tooltiptext2    =   "Enter Component Code"
                  colname3        =   "DCTSZ"
                  heading3        =   "Size    "
                  datafld3        =   "DctSz"
                  datatype3       =   4
                  maxlength3      =   5
                  tooltiptext3    =   "Enter Component Size"
                  colname4        =   "DCTCHR"
                  heading4        =   "Comp Chr"
                  datafld4        =   "DctChr"
                  datatype4       =   4
                  maxlength4      =   3
                  tooltiptext4    =   "Enter Component Character"
                  colname5        =   "DCTPARTS"
                  heading5        =   "Prt Qty"
                  datafld5        =   "DctParts"
                  datatype5       =   1
                  mask5           =   "##0"
                  maxlength5      =   3
                  tooltiptext5    =   "Enter Component Quantity"
                  colname6        =   "DCTRMCD"
                  heading6        =   "Main RmCd                               "
                  datafld6        =   "DctRmCd"
                  datatype6       =   4
                  maxlength6      =   16
                  tooltiptext6    =   "Enter Main Raw Material Code"
                  colname7        =   "DCTDMCOL"
                  heading7        =   "Dm Color"
                  datafld7        =   "DctDmCol"
                  datatype7       =   4
                  maxlength7      =   5
                  tooltiptext7    =   "Enter Main Rm Color"
                  colname8        =   "DCTRMWT"
                  heading8        =   "Rm Weight"
                  datafld8        =   "DctRmWt"
                  datatype8       =   2
                  mask8           =   "###0.000"
                  maxlength8      =   8
                  tooltiptext8    =   "Enter Rm Weight"
                  colname9        =   "DCTDC"
                  heading9        =   "DC  "
                  datafld9        =   "DctDc"
                  datatype9       =   4
                  maxlength9      =   2
                  tooltiptext9    =   "Specify 'C' Only Clr Stn/ 'D' Only Diamond/ 'CD' for Both/ Blank if No Stone is used"
                  colname10       =   "DCTPRDSEQ"
                  heading10       =   "PrdSeq     "
                  datafld10       =   "DctPrdSeq"
                  datatype10      =   4
                  maxlength10     =   5
                  tooltiptext10   =   "Enter Production Sequence"
                  colname11       =   "DCTMAINPRTYN"
                  heading11       =   "MainPrt"
                  datafld11       =   "DctMainPrtYn"
                  datatype11      =   4
                  maxlength11     =   1
                  tooltiptext11   =   "Main Part Yes/ No"
                  colname12       =   "DCTDESC"
                  heading12       =   "Description                                                                                           "
                  datafld12       =   "DctDesc"
                  maxlength12     =   45
                  tooltiptext12   =   "Enter Discription for Compnent"
               End
            End
            Begin VB.Frame FraDesgnDet 
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               Height          =   4215
               Index           =   7
               Left            =   -74940
               TabIndex        =   118
               Top             =   360
               Width           =   14685
               Begin MwfCtl.MWCTL_FLX GrdDsgModAvl 
                  Height          =   4110
                  Left            =   75
                  TabIndex        =   119
                  Top             =   60
                  Width           =   8115
                  _ExtentX        =   14314
                  _ExtentY        =   7250
                  Cols            =   4
                  colname1        =   "DMASR"
                  heading1        =   "Sr.     "
                  datafld1        =   "DmaSr"
                  datatype1       =   1
                  mask1           =   "#0"
                  maxlength1      =   2
                  style1          =   2
                  colname2        =   "DMAAVLSZ"
                  heading2        =   "Avl. Sizes               "
                  datafld2        =   "DmaAvlSz"
                  datatype2       =   4
                  maxlength2      =   5
                  tooltiptext2    =   "Enter Available Sizes"
                  colname3        =   "DMADESC"
                  heading3        =   "Description                                                                     "
                  datafld3        =   "DmaDesc"
                  maxlength3      =   45
                  tooltiptext3    =   "Enter Description"
               End
            End
         End
         Begin VB.Frame FraPicHsk 
            BorderStyle     =   0  'None
            Height          =   1635
            Left            =   13380
            TabIndex        =   89
            Top             =   2130
            Width           =   1725
            Begin VB.Image PicHsk 
               Height          =   1575
               Left            =   0
               Top             =   0
               Width           =   1665
            End
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
            Left            =   13800
            TabIndex        =   87
            Top             =   3820
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
            Left            =   13890
            TabIndex        =   86
            Top             =   1780
            Width           =   765
         End
      End
      Begin VB.Frame FraCpyCat 
         Height          =   945
         Left            =   5400
         TabIndex        =   225
         Top             =   8160
         Width           =   4425
         Begin MwfCtl.MWCTL_BTN1 CmdCpyCatGo 
            Height          =   495
            Left            =   3780
            TabIndex        =   226
            ToolTipText     =   "Click Here To Copy"
            Top             =   390
            Width           =   555
            _ExtentX        =   979
            _ExtentY        =   873
            BackColor       =   -2147483633
            MaskColor       =   8454143
            Caption         =   "&Go"
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
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   84
            Left            =   1185
            TabIndex        =   227
            ToolTipText     =   "Enter Source Design Code"
            Top             =   525
            Width           =   2355
            _ExtentX        =   4154
            _ExtentY        =   503
            MaxLength       =   15
            DataType        =   4
            IdName          =   "WDPCD"
         End
         Begin VB.Label LblMain 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "  Copy To Catalog"
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
            Index           =   1
            Left            =   -30
            TabIndex        =   229
            Top             =   90
            Width           =   4485
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Catalog Cd"
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
            Index           =   64
            Left            =   60
            TabIndex        =   228
            Top             =   525
            Width           =   1125
         End
      End
   End
End
Attribute VB_Name = "EmrFrmDsgView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---- General Comments on the Entry --------'
Option Explicit
Dim wb_OnHold As String, wb_Valid As String, wb_PrdRmQty As String, wb_RegExcl As String
Dim moCn As MwfLib.MDOConnection   '*** (Bef speed) bef 24/11/06
Dim ws_DmTcTyp As String
Dim ws_DmCtgFr As String, ws_DmCtgTo As String, ws_DmCtgSel As String, ws_PrdCtgFr As String, ws_PrdCtgTo As String, ws_PrdCtgSel As String
Dim ws_SalCtgFr As String, ws_SalCtgTo As String, ws_SalCtgSel As String, ws_DmCdFr As String, ws_DmCdTo As String, ws_DmCdSel As String
Dim ws_DmSzFr As String, ws_DmSzTo As String, ws_DmSzSel As String, ws_SetCdFr As String, ws_SetCdTo As String, ws_SetCdSel As String
Dim ws_PrdSeqFr As String, ws_PrdSeqTo As String, ws_PrdSeqSel As String, ws_DmColFr As String, ws_DmColTo As String, ws_DmColSel As String
Dim wi_MMtlWtFr As Double, wi_MMtlWtTo As Double, wi_GrsWtFr As Double, wi_GrsWtTo As Double
Dim wi_DiaWtFr As Double, wi_DiaWtTo As Double, wi_CSWtFr As Double, wi_CSWtTo As Double
Dim wi_CSQtyFr As Double, wi_CSQtyTo As Double, w_DsgDtFr As Date, w_DsgDtTo As Date
Dim wi_DiaQtyFr As Double, wi_DiaQtyTo As Double
Dim ws_CmCdFr As String, ws_CmCdTo As String
Dim ws_DmOldCdFr As String, ws_DmOldCdTo As String, ws_DmOldCdSel As String
'*** Jay 3.4[DP]
Dim w_DsgCrDtFr As Date, w_DsgCrDtTo As Date
Dim ws_DrSetSCdFr As String, ws_DrSetSCdTo As String, ws_DrSetSCdSel As String
Dim ws_DpCollFr As String, ws_DpCollTo As String, wi_DpSerFr As Integer, wi_DpSerTo As Integer, wi_DpSetFr As Integer, wi_DpSetTo As Integer, wi_DpVerFr As Integer, wi_DpVerTo As Integer
Dim ws_DpCollSel As String, ws_DpSerSel As String, ws_DpSetSel As String, ws_DpVerSel As String
Dim ms_DmCdStr As String
'*** Jay 3.4[DP]

Dim wPic3DHt As Single
Dim wPic3DWdt As Single
Dim wPicHskHt As Single
Dim wPicHskWdt As Single

Dim wPicDPHt As Single, wPicDPWdt As Single     '*** Jay 3.4[DP]

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

Enum en_DsgVwFra
  Summ = 0
  CatCpy = 1 '*** Jay 3.4[DP]
End Enum

'*** 2.11 (Jen)
Dim mi_FraKeysHt As Integer
'*** 2.11 (Jen)

'*** (Jen 08/06/07)
Dim mb_ModAvl As Boolean, mb_CTMod As Boolean
'*** (Jen 08/06/07)

' ***** Manali - 22/10/08 - Design Catalogue Module
Dim mb_DsgCat As Boolean
' ***** Manali - 22/10/08 - Design Catalogue Module

Private Sub ADC_Load()
  '*** Set the First Key Control property to the working field wDtTag
  '*** Set the First Non Key Control property to the grid GrdDt
  '*** Set the Child Property Of ADC to GrdDt
  '*** Set the Previous and Next Control Properties for the Grid GrdDt
  '*** Set the Hot Key for the Grid GrdDt as 'A'
  Set moCn = ADC.Connection    '*** (Bef speed) bef 24/11/06
  Set ADC.FirKeyCtl = ADC("WDMTYP")
  Set ADC.FirNKeyCtl = GrdDsgLst
  
  '*** (Jen 08/06/07)
  ADC.Child = "GrdDsgLst"
  Call GrdDsgAnaScop.RefreshData
  Call GrdDsgSz.RefreshData
  '*** (Jen 08/06/07)
  '578-12 GrdDsgAllow added
  GrdDsgLst.Child = "GrdDsgRm,GrdDsgLab,GrdDsgAna,GrdDsgCT,GrdDsgModAvl,GrdDsgPrm,GrdDsgAllow" '*** Jay 3.4[DP]
  GrdDsgPrm.Child = "GrdDsgPrmDet" '*** Jay 3.4[DP]
  
  GrdDsgLst.HotKey = "A"
  If mi_FraKeysHt = 0 Then mi_FraKeysHt = FRA_KEYS.Height
  TabDesgnDet.Tab = 0
  FRA_KEYS.ZOrder (0)
    
  Set GrdDsgRm.PrevCtl = GrdDsgLst
  Set GrdDsgRm.NextCtl = GrdDsgLst
  Set GrdDsgLab.PrevCtl = GrdDsgLst
  Set GrdDsgLab.NextCtl = GrdDsgLst
  Set GrdDsgCT.PrevCtl = GrdDsgLst
  Set GrdDsgCT.NextCtl = GrdDsgLst
  Set GrdDsgModAvl.PrevCtl = GrdDsgLst
  Set GrdDsgModAvl.NextCtl = GrdDsgLst
    
  ''adc.Child = "GrdDsgAnaScop"
  Set GrdDsgAnaScop.PrevCtl = CmdDetScope   'ADC
  Set GrdDsgAnaScop.NextCtl = GrdDsgSz   'ADC
  Set GrdDsgSz.PrevCtl = GrdDsgAnaScop
  Set GrdDsgSz.NextCtl = CmdDetScope
  '*** Jay 3.4[DP]
  Set GrdDsgPrm.PrevCtl = GrdDsgLst
  Set GrdDsgPrm.NextCtl = GrdDsgPrmDet
  Set GrdDsgPrmDet.PrevCtl = GrdDsgPrm
  Set GrdDsgPrmDet.NextCtl = GrdDsgLst
  '*** Jay 3.4[DP]
  Set GrdDsgAllow.PrevCtl = GrdDsgLst   '578-12
  Set GrdDsgAllow.NextCtl = GrdDsgLst   '578-12
  
  GrdDsgLst.AllowAdd = False: GrdDsgLst.AllowDelete = False
  GrdDsgRm.AllowAdd = False: GrdDsgRm.AllowDelete = False
  GrdDsgLab.AllowAdd = False: GrdDsgLab.AllowDelete = False
  GrdDsgCT.AllowAdd = False: GrdDsgCT.AllowDelete = False
  GrdDsgAna.AllowAdd = False: GrdDsgAna.AllowDelete = False
  GrdDsgModAvl.AllowAdd = False: GrdDsgModAvl.AllowDelete = False
  GrdDsgPrm.AllowAdd = False: GrdDsgPrm.AllowDelete = False         '*** Jay 3.4[DP]
  GrdDsgPrmDet.AllowAdd = False: GrdDsgPrmDet.AllowDelete = False   '*** Jay 3.4[DP]
  GrdDsgAllow.AllowAdd = False: GrdDsgAllow.AllowDelete = False   '578-12
  'FraDesgnDet(0).Enabled = False
  
  'wb_OnHold = False:  wb_PrdRmQty = False:  wb_RegExcl = False
  w_DsgDtFr = "01/01/80"
  w_DsgDtTo = "01/01/80"
  '*** Jay 3.4[DP]
  w_DsgCrDtFr = MWLib.EmptyDate:  w_DsgCrDtTo = MWLib.EmptyDate
  '*** Jay 3.4[DP]
 '*** Storing The Original Size And Position Of The Picture Controls ***
  wPic3DHt = Pic3D.Height
  wPic3DWdt = Pic3D.Width
  wPicHskHt = PicHsk.Height
  wPicHskWdt = PicHsk.Width
  
  wPicDPHt = PicDP.Height: wPicDPWdt = PicDP.Width  '*** Jay 3.4[DP]
  
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
  
  '*** (Jen 08/06/07)
  mb_ModAvl = moCn.RecSeek("Select HModAvlYN From Head Where HCd='" + ctSelfCmCd + "' And HCoCd='" + ctSelfCoCd + "' and HModAvlYN= 'Y'")
  mb_CTMod = moCn.RecSeek("Select hCompBagsYn From Head Where HCd='" + ctSelfCmCd + "' And HCoCd='" + ctSelfCoCd + "' and hCompBagsYn= 'Y'")
  '*** (Jen 08/06/07)
  
  '****** Sachin 3.02.0 - Hiding the Partition Key Column based on Partition Selected
  GrdDsgLst.ColProp("DmPrtKey").Style = IIF(gs_Partition = ctCurrPrtn, fgcinvisible, fgcReadOnly)
  
  ' ***** Manali - 22/10/08 - Design Catalogue Module
  mb_DsgCat = moCn.RecSeek("Select 'x' From Head Where HCd='" + ctSelfCmCd + "' And HCoCd='" + ctSelfCoCd + "' and HDsgCatYn= 'Y'")
  
  If mb_DsgCat = False Then
    LblColl.Visible = False: LblSer.Visible = False: LblSet.Visible = False: LblVer.Visible = False
    ADC("WDPCOLLFR").Enabled = False: ADC("WDPCOLLFR").Visible = False
    ADC("WDPCOLLTO").Enabled = False: ADC("WDPCOLLTO").Visible = False
    ADC("WDPCOLLSEL").Enabled = False: ADC("WDPCOLLSEL").Visible = False
    ADC("WDPSERFR").Enabled = False: ADC("WDPSERFR").Visible = False
    ADC("WDPSERTO").Enabled = False: ADC("WDPSERTO").Visible = False
    ADC("WDPSERSEL").Enabled = False: ADC("WDPSERSEL").Visible = False
    ADC("WDPSETFR").Enabled = False: ADC("WDPSETFR").Visible = False
    ADC("WDPSETTO").Enabled = False: ADC("WDPSETTO").Visible = False
    ADC("WDPSETSEL").Enabled = False: ADC("WDPSETSEL").Visible = False
    ADC("WDPVERFR").Enabled = False: ADC("WDPVERFR").Visible = False
    ADC("WDPVERTO").Enabled = False: ADC("WDPVERTO").Visible = False
    ADC("WDPVERSEL").Enabled = False: ADC("WDPVERSEL").Visible = False
    ''' CmdCpyCat.Enabled = False: CmdCpyCat.Visible = False  ' ***** Manali 3.6.0 - 30/09/09 - Catalogue Entry removed from Module Check
  End If
  ' ***** Manali - 22/10/08 - Design Catalogue Module
  
End Sub

Private Sub ADC_KeyWhen()
  '*** Set the frame FraNKeyAll as the top most frame
  '****** Sachin 2.14.0 - [24-11-2006] - [Multi-Company] ******
  Me.Caption = GetFrmCaption(gs_CoCd, "Detailed Design View")
  '****** Sachin 2.14.0 - [24-11-2006] - [Multi-Company] ******
  
  FraNKeyAll.ZOrder (0)
  TabDesgnDet.Tab = 0   '*** (Jen 08/06/07)
  
  GrdDsgLst.Rows = 1
  
  FRA_KEYS.Height = mi_FraKeysHt
  'CmdSumm.Visible = False      '*** (Bef 08/06/07)
  '*** Jay 3.4[DP]
  Call EnaDisaCmds(True)
  Call HideAllFras
  
  CmdDetScope.Visible = True
  CmdDetScope.Enabled = True
  
  '*** Jay 3.4[DP]
  ws_DmTcTyp = IIF(Trim(ws_DmTcTyp) <> "", ws_DmTcTyp, "DM")    'Setting the Default TcTyp to Design Master
  
  ADC("wDmTyp") = ws_DmTcTyp
  ADC("wDmCtgFr") = ws_DmCtgFr:  ADC("wDmCtgTo") = ws_DmCtgTo: ADC("wDmCtgSel") = ws_DmCtgSel
  ADC("wDmPrdCtgFr") = ws_PrdCtgFr:  ADC("wDmPrdCtgTo") = ws_PrdCtgTo:  ADC("wPrdCtgSel") = ws_PrdCtgSel
  ADC("wDmSalCtgFr") = ws_SalCtgFr:  ADC("wDmSalCtgTo") = ws_SalCtgTo:  ADC("wSalCtgSel") = ws_SalCtgSel
  ADC("wDmCdFr") = ws_DmCdFr:    ADC("wDmCdTo") = ws_DmCdTo:    ADC("wDmCdSel") = ws_DmCdSel
  ADC("wDmOldCdFr") = ws_DmOldCdFr:    ADC("wDmOldCdTo") = ws_DmOldCdTo:    ADC("wDmOldCdSel") = ws_DmOldCdSel
  ADC("wDmSzFr") = ws_DmSzFr: ADC("wDmSzTo") = ws_DmSzTo: ADC("wDmSzSel") = ws_DmSzSel
  ADC("wSetCdFr") = ws_SetCdFr:  ADC("wSetCdTo") = ws_SetCdTo:  ADC("wSetCdSel") = ws_SetCdSel
  ADC("wDmPrdSeqFr") = ws_PrdSeqFr:  ADC("wDmPrdSeqTo") = ws_PrdSeqTo:  ADC("wPrdSeqSel") = ws_PrdSeqSel
  ADC("wDmColFr") = ws_DmColFr: ADC("wDmColTo") = ws_DmColTo: ADC("wDmColSel") = ws_DmColSel
  ADC("wMainMetWtFr") = wi_MMtlWtFr: ADC("wMainMetWtTo") = wi_MMtlWtTo
  ADC("wRmGrsWtFr") = wi_GrsWtFr:  ADC("wRmGrsWtTo") = wi_GrsWtTo
  ADC("wRmDiaWtFr") = wi_DiaWtFr: ADC("wRmDiaWtTo") = wi_DiaWtTo
  ADC("wRmCSWtFr") = wi_CSWtFr: ADC("wRmCSWtTo") = wi_CSWtTo
  ADC("wRmCSQtyFr") = wi_CSQtyFr: ADC("wRmCSQtyTo") = wi_CSQtyTo
  ADC("wRmDiaQtyFr") = wi_DiaQtyFr: ADC("wRmDiaQtyTo") = wi_DiaQtyTo
  'ADC("wCmCdFr") = ws_CmCdFr: ADC("wCmCdTo") = ws_CmCdTo    '578-12 customer code disabled
  ADC("wODtFr") = w_DsgDtFr: ADC("wODtTo") = w_DsgDtTo
  ADC("wOnHldYN") = wb_OnHold
  ADC("wValidYN") = wb_Valid                                            '4.1.3.0
  ADC("wPrdRmQtyWtYN") = IIF(wb_PrdRmQty <> "", wb_PrdRmQty, "N")
  'adc("wDmRegExc") = IIF(wb_RegExcl <> "", wb_RegExcl, "A")  '578-12 exlusive option disabled
           
  '****** Manali 3.5.0 - 05/05/09 - Radio Buttons Added
  If ADC("wOnHldYN") = "Y" Then
    OptDmHld(0).Value = True
  ElseIf ADC("wOnHldYN") = "N" Then
    OptDmHld(1).Value = True
  Else
     OptDmHld(2).Value = True
  End If
  
  '4.1.3.0
  If ADC("wValidYN") = "Y" Then
    OptDmValid(0).Value = True
  ElseIf ADC("wValidYN") = "N" Then
    OptDmValid(1).Value = True
  Else
    OptDmValid(2).Value = True
  End If
  
  If ADC("wPrdRmQtyWtYN") = "Y" Then OptRmPrd(0).Value = True
  If ADC("wPrdRmQtyWtYN") = "N" Then OptRmPrd(1).Value = True
  
  ' ****** Manali 350Nxt - Valid value for wDmRegExc changed from Parameter 'YN' to 'DSGEXC'
'578-12 exclusive option disabled
'  If ADC("wDmRegExc") = "R" Then
'    OptDsgExc(0).Value = True
'  ElseIf ADC("wDmRegExc") = "C" Then
'    OptDsgExc(1).Value = True
'  ElseIf ADC("wDmRegExc") = "F" Then
'    OptDsgExc(2).Value = True
'  ElseIf ADC("wDmRegExc") = "A" Then
'    OptDsgExc(3).Value = True
'  End If
  ' ****** Manali 350Nxt - Valid value for wDmRegExc changed from Parameter 'YN' to 'DSGEXC'

  ''  If adc("wDmRegExc") = "Y" Then
  ''    OptDsgExc(0).Value = True
  ''  ElseIf adc("wDmRegExc") = "N" Then
  ''    OptDsgExc(1).Value = True
  ''  Else
  ''    OptDsgExc(2).Value = True
  ''  End If
  ' ****** Manali 3.5.0 - 05/05/09 - Radio Buttons Added
      
  '*** Jay 3.4[DP]
  ADC("wOCrDtFr") = w_DsgCrDtFr: ADC("wOCrDtTo") = w_DsgCrDtTo
  ADC("wDrSetSCdFr") = ws_DrSetSCdFr: ADC("wDrSetSCdTo") = ws_DrSetSCdTo: ADC("wDrSetSCdSel") = ws_DrSetSCdSel
  ADC("wDpCollFr") = ws_DpCollFr: ADC("wDpCollTo") = ws_DpCollTo: ADC("wDpCollSel") = ws_DpCollSel
  ADC("wDpSerFr") = wi_DpSerFr: ADC("wDpSerTo") = wi_DpSerTo: ADC("wDpSerSel") = ws_DpSerSel
  ADC("wDpSetFr") = wi_DpSetFr: ADC("wDpSetTo") = wi_DpSetTo: ADC("wDpSetSel") = ws_DpSetSel
  ADC("wDpVerFr") = wi_DpVerFr: ADC("wDpVerTo") = wi_DpVerTo: ADC("wDpVerSel") = ws_DpVerSel
  '*** Jay 3.4[DP]
  
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
End Sub

Private Sub ADC_KeyValid(pr_Cancel As Boolean, pr_ErrMsg As String)
  '*** Check whether valid
  
  '*** (Jen 09/06/07)
  ''' If ADC("wDmRegExc") = "" Then
  ' ****** Manali 350Nxt - Valid value for wDmRegExc changed from Parameter 'YN' to 'DSGEXC'
'578-12 customer code disabled
'  If ADC("wDmRegExc") = "F" Then
'    If ADC("wCmCdFr") <> "" Or ADC("wCmCdTo") <> "" Then pr_Cancel = True: pr_ErrMsg = "Customer Code Range Cannot Be Specified For Free Designs"
'  Else
'    'If ADC("wCmCdFr") = "" And ADC("wCmCdTo") = "" Then Cancel = True: pr_ErrMsg = "Some Customer Code Range Should Be Specified For Region Exclusivity Check"
'  End If
  '*** (Jen 09/06/07)
End Sub
Private Sub ADC_NonKeyWhen(pr_Cancel As Boolean, pr_ErrMsg As String)
  FRA_KEYS.Height = 200
  ws_DmTcTyp = ADC("wDmTyp")
  ws_DmCtgFr = ADC("wDmCtgFr"):  ws_DmCtgTo = ADC("wDmCtgTo"):  ws_DmCtgSel = ADC("wDmCtgSel")
  ws_PrdCtgFr = ADC("wDmPrdCtgFr"):  ws_PrdCtgTo = ADC("wDmPrdCtgTo"):  ws_PrdCtgSel = ADC("wPrdCtgSel")
  ws_SalCtgFr = ADC("wDmSalCtgFr"):  ws_SalCtgTo = ADC("wDmSalCtgTo"):  ws_SalCtgSel = ADC("wSalCtgSel")
  ws_DmCdFr = ADC("wDmCdFr"):  ws_DmCdTo = ADC("wDmCdTo"):  ws_DmCdSel = ADC("wDmCdSel")
  ws_DmOldCdFr = ADC("wDmOldCdFr"):  ws_DmOldCdTo = ADC("wDmOldCdTo"):  ws_DmOldCdSel = ADC("wDmOldCdSel")
  ws_DmSzFr = ADC("wDmSzFr"):  ws_DmSzTo = ADC("wDmSzTo"):  ws_DmSzSel = ADC("wDmSzSel")
  ws_SetCdFr = ADC("wSetCdFr"):  ws_SetCdTo = ADC("wSetCdTo"):  ws_SetCdSel = ADC("wSetCdSel")
  ws_PrdSeqFr = ADC("wDmPrdSeqFr"):  ws_PrdSeqTo = ADC("wDmPrdSeqTo"):  ws_PrdSeqSel = ADC("wPrdSeqSel")
  ws_DmColFr = ADC("wDmColFr"): ws_DmColTo = ADC("wDmColTo"): ws_DmColSel = ADC("wDmColSel")
  wi_MMtlWtFr = ADC("wMainMetWtFr"):  wi_MMtlWtTo = ADC("wMainMetWtTo")
  wi_GrsWtFr = ADC("wRmGrsWtFr"):  wi_GrsWtTo = ADC("wRmGrsWtTo")
  wi_DiaWtFr = ADC("wRmDiaWtFr"):  wi_DiaWtTo = ADC("wRmDiaWtTo")
  wi_CSWtFr = ADC("wRmCSWtFr"):  wi_CSWtTo = ADC("wRmCSWtTo")
  wi_CSQtyFr = ADC("wRmCSQtyFr"):  wi_CSQtyTo = ADC("wRmCSQtyTo")
  wi_DiaQtyFr = ADC("wRmDiaQtyFr"):  wi_DiaQtyTo = ADC("wRmDiaQtyTo")
  w_DsgDtFr = CDate(ADC("wODtFr")): w_DsgDtTo = CDate(ADC("wODtTo"))
  'ws_CmCdFr = ADC("wCmCdFr"): ws_CmCdTo = ADC("wCmCdTo")     '578-12 customer code disabled
  w_DsgDtFr = IIF(ADC("wODtFr") <> "", ADC("wODtFr"), "01/01/80")
  w_DsgDtTo = IIF(ADC("wODtTo") <> "", ADC("wODtTo"), "01/01/80")
  wb_OnHold = ADC("wOnHldYN")
  wb_Valid = ADC("wValidYN")        '4.1.3.0
  wb_PrdRmQty = ADC("wPrdRmQtyWtYN")
  'wb_RegExcl = adc("wDmRegExc")    '578-12 exlusive option disabled
  

  '*** Jay 3.4[DP]
  w_DsgCrDtFr = IIF(ADC("wOCrDtFr") <> "", ADC("wOCrDtFr"), MWLib.EmptyDate)
  w_DsgCrDtTo = IIF(ADC("wOCrDtTo") <> "", ADC("wOCrDtTo"), MWLib.EmptyDate)
  ws_DrSetSCdFr = ADC("wDrSetSCdFr"): ws_DrSetSCdTo = ADC("wDrSetSCdTo"): ws_DrSetSCdSel = ADC("wDrSetSCdSel")
  ws_DpCollFr = ADC("wDpCollFr"): ws_DpCollTo = ADC("wDpCollTo"): ws_DpCollSel = ADC("wDpCollSel")
  wi_DpSerFr = ADC("wDpSerFr"): wi_DpSerTo = ADC("wDpSerTo"): ws_DpSerSel = ADC("wDpSerSel")
  wi_DpSetFr = ADC("wDpSetFr"): wi_DpSetTo = ADC("wDpSetTo"): ws_DpSetSel = ADC("wDpSetSel")
  wi_DpVerFr = ADC("wDpVerFr"): wi_DpVerTo = ADC("wDpVerTo"): ws_DpVerSel = ADC("wDpVerSel")
  ms_DmCdStr = ""
  '*** Jay 3.4[DP]
  
  '*** Jay 3.4[DP]
  Call EnaDisaCmds(False)
  CmdDetScope.Visible = False
  CmdDetScope.Enabled = False
  '*** Jay 3.4[DP]
  
  FraSC.Enabled = False:  FraSC.Visible = False
  FraScop.Enabled = True
  'FraScopDet.Enabled = True
  
  Call SetSelection
  Call FillGrdDsgLst
  
  '*** Jay 3.4[DP]
  ' ***** Manali - 22/10/08 - Design Catalogue Module cond added
  ' ***** Manali 3.6.0 - 30/09/09 - Catalogue Entry removed from Module Check
  '''  If ADC("WDMTYP") = "DM" And mb_DsgCat = True Then
  '''    GrdDsgLst.ColProp("wSel").Style = fgcNormal
  '''    CmdCpyCat.Visible = True
  '''    CmdCpyCat.Enabled = True
  '''  Else
  '''    GrdDsgLst.ColProp("wSel").Style = fgcinvisible
  '''    CmdCpyCat.Visible = False
  '''    CmdCpyCat.Enabled = False
  '''  End If
  '*** Jay 3.4[DP]
  
End Sub

Private Sub ADC_SetRecSource()
    '*** Since this is an 'A' Entry, there is no RecSource associated with ADC. So set the
    'DataMode property of ADC control as xdatUnbound
  ADC.DataMode = xdatUnbound
End Sub

Private Sub atxt_FldChange(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String, ByVal pv_NewValue As Variant)
    Select Case UCase(IdName)
        Case Is = UCase("wDmTyp")
            Cancel = Not moCn.RecSeek("Select 'x' From Param Where PTyp= 'DMTCTYP' And PMCd='" + pv_NewValue + "'")
            If Cancel = True Then ErrMsg = "Enter a Valid DmTcTyp": Exit Sub
            
        Case Is = UCase("wShowPrdRmDets")
            Cancel = Not moCn.RecSeek("select PMCd from Param where " + _
                                 "PTyp= 'YN' and PMCd = '" + pv_NewValue + "'")
            If Cancel = True Then ErrMsg = "Enter 'Y' Or 'N' For Prd Or Sales": Exit Sub
            Call Summary(pv_NewValue)
            
        Case Is = UCase("wOnHldYn")
            Cancel = pv_NewValue <> "" And (Not moCn.RecSeek("select PMCd from Param where " + _
                                 " PTyp= 'YN' and PMCd = '" + pv_NewValue + "'"))
            ErrMsg = "Enter On Hold Value As Y, N Or Blank": Exit Sub
            
        'DmValidYN - 4.1.3.0
        Case Is = UCase("wValidYn")
            Cancel = pv_NewValue <> "" And (Not moCn.RecSeek("select PMCd from Param where " + _
                                 " PTyp= 'YN' and PMCd = '" + pv_NewValue + "'"))
            ErrMsg = "Enter Design Valid Value As Y, N Or Blank": Exit Sub
            
            
        Case Is = UCase("wPrdRmQtyWtYN")
            Cancel = (Not moCn.RecSeek("select PMCd from Param where " + _
                                 " PTyp= 'YN' and PMCd = '" + pv_NewValue + "'"))
            ErrMsg = "Enter Whether Production/Sales RM Quantity Or Weight As Y, N ": Exit Sub
        '578-12 exclusive option disabled
'        Case Is = UCase("WDMREGEXC")
'            Cancel = Not moCn.RecSeek("select PMCd from Param where " + _
'                                 " PTyp= 'DSGEXC' and PMCd = '" + pv_NewValue + "'")
'            ErrMsg = "Enter Design Exclusivity": Exit Sub
        
''            Cancel = pv_NewValue <> "" And (Not moCn.RecSeek("select PMCd from Param where " + _
''                                 " PTyp= 'YN' and PMCd = '" + pv_NewValue + "'"))
''            ErrMsg = "Enter Region Exclusive Value As Y, N Or Blank": Exit Sub
        '*** Jay 3.4[DP]
        Case Is = UCase("wDpCd")
            Cancel = pv_NewValue <> "" And (Not moCn.RecSeek("select 'x' from DsgPrm where " + _
                                 " DpTyp= 'CAT' and DpCd = '" + pv_NewValue + "' and DpDmCd=''"))
            ErrMsg = "Catalogue Code Not Define Or Invalid": Exit Sub
        '*** Jay 3.4[DP]
    End Select
End Sub

Private Sub ATXT_FldWhen(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String)
  
  '****** Sachin 3.01 - [28-05-2007] ******
  Select Case UCase(IdName)
    Case Is = UCase("wDmTyp")
      Call HlpList.PMCd("DmTcTyp")
      
    Case Is = UCase("wDmCtgFr"), UCase("wDmCtgTo"), UCase("wDmCtgSel")
      If IdName = UCase("wDmCtgSel") Then HlpList.MultiSelect = True
      Call HlpList.PMCd("DmCtg")
      
    Case Is = UCase("wDmPrdCtgFr"), UCase("wDmPrdCtgTo"), UCase("wPrdCtgSel")
      If IdName = UCase("wPrdCtgSel") Then HlpList.MultiSelect = True
      Call HlpList.PMCd("PrdCtg")
      
    Case Is = UCase("wDmSalCtgFr"), UCase("wDmSalCtgTo"), UCase("wSalCtgSel")
      If IdName = UCase("wSalCtgSel") Then HlpList.MultiSelect = True
      'If IdName = UCase("wDmSalCtgFr") And IdName = UCase("wSalCtgSel") Then     '*** (Bef 08/06/07)
      '*** (Jen 08/06/07)
      If IdName = UCase("wDmSalCtgFr") Or IdName = UCase("wSalCtgSel") Then
      '*** (Jen 08/06/07)
      '*** Jay 3.4[DP]
        Call HlpList.PSCd("SalCtg", ADC(("wDmCtgFr")), True)
        'Call HlpList.PSCd("SalCtg", ADC(("wDmCtgFr")))
      Else
        Call HlpList.PSCd("SalCtg", ADC(("wDmCtgTo")), True)
        'Call HlpList.PSCd("SalCtg", ADC(("wDmCtgTo")))
      '*** Jay 3.4[DP]
      End If
      
    Case Is = UCase("wDmCdFr"), UCase("wDmCdTo"), UCase("wDmCdSel")
      If IdName = UCase("wDmCdSel") Then HlpList.MultiSelect = True
      Call HlpList.DmCd(ADC("wDmTyp"))
      
    Case Is = UCase("wDmOldCdFr"), UCase("wDmOldCdTo"), UCase("wDmOldCdSel")
      If IdName = UCase("wDmOldCdSel") Then HlpList.MultiSelect = True
      Call HlpList.DmOldCd(ADC("wDmTyp"))
    
    Case Is = UCase("wDmSzFr"), UCase("wDmSzTo"), UCase("wDmSzSel")
      If IdName = UCase("wDmSzSel") Then HlpList.MultiSelect = True
      Call HlpList.PMCd("DmSz")
      
    Case Is = UCase("wSetCdFr"), UCase("wSetCdTo"), UCase("wSetCdSel")
      If IdName = UCase("wSetCdSel") Then HlpList.MultiSelect = True
      Call HlpList.DmSetCd
      
    Case Is = UCase("wDmPrdSeqFr"), UCase("wDmPrdSeqTo"), UCase("wPrdSeqSel")
      If IdName = UCase("wPrdSeqSel") Then HlpList.MultiSelect = True
      Call HlpList.vPMCd(gs_CoCd, "PrdSeq")
      
    Case Is = UCase("wDmColFr"), UCase("wDmColTo"), UCase("wDmColSel")
      If IdName = UCase("wDmColSel") Then HlpList.MultiSelect = True
      Call HlpList.PMCd("DmCol")
      
    Case Is = UCase("wOnHldYN"), UCase("wPrdRmQtyWtYN"), UCase("wValidYN") ', UCase("wDmRegExc")    '4.1.3
      Call HlpList.PMCd("YN")
      
    Case Is = UCase("wCmCdFr"), UCase("wCmCdTo")
      HlpList.CustCd "C", True
    '*** Jay 3.4[DP]   {added ("wDmUom")}
    '4.1.3.0
    Case Is = UCase("wDmPrdCtg"), UCase("wDmSalCtg"), UCase("wDmVaCtg"), UCase("wDmLsCtg"), _
              UCase("wDmKt"), UCase("wDmBagPcs"), UCase("wDmSetCd"), _
              UCase("wDmOldCd"), UCase("wDmCol"), UCase("wDmCmCd"), _
              UCase("wDmRegnExc"), UCase("wDmParts"), UCase("wDmPartDesc"), _
              UCase("wDmHld"), UCase("wDmValidYN"), UCase("wDmHldDesc"), UCase("wDmPrdSeq"), _
              UCase("wDmPrdInst"), UCase("wDmDsgDt"), UCase("wDmDsgBy"), _
              UCase("wDmModMkr"), UCase("wDmLstMdf"), UCase("wDmWaxWt"), _
              UCase("wDmModRunWt"), UCase("wDmSilModWt"), UCase("wDmCasPcWt"), _
              UCase("DmHistory"), UCase("DmFaults"), UCase("wCmName"), UCase("wDmUom")
      Cancel = True
    
    '*** Jay 3.4[DP]
     Case Is = UCase("wDrSetSCdFr"), UCase("wDrSetSCdTo"), UCase("wDrSetSCdSel")
      If IdName = UCase("wDrSetSCdSel") Then HlpList.MultiSelect = True
      Call HlpList.PSCd("LABSCD", "SET")
    Case Is = UCase("wDpCd")
      HlpList.DsgPrmCd ("CAT")
    Case Is = UCase("wDpCollFr"), UCase("wDpCollTo"), UCase("wDpCollSel")
      If IdName = UCase("wDpCollSel") Then HlpList.MultiSelect = True
      Call HlpList.DsgPrmCd("COLL")
    Case Is = UCase("wDpSerFr"), UCase("wDpSerTo")
      Call HlpList.DsgPrmCd("SER", True)
    Case Is = UCase("wDpSerSel")
      HlpList.MultiSelect = True
      Call HlpList.DsgPrmCd("SER", True)
    Case Is = UCase("wDpSetFr"), UCase("wDpSetTo")
      Call HlpList.DsgPrmCd("SET", True)
    Case Is = UCase("wDpSetSel")
      HlpList.MultiSelect = True
      Call HlpList.DsgPrmCd("SET", True)
    Case Is = UCase("wDpVerFr"), UCase("wDpVerTo")
      Call HlpList.DsgPrmCd("VER", True)
    Case Is = UCase("wDpVerSel")
      HlpList.MultiSelect = True
      Call HlpList.DsgPrmCd("VER", True)
    '*** Jay 3.4[DP]
  End Select
  '****** Sachin 3.01 - [28-05-2007] ******
End Sub

Private Sub FillGrdDsgLst()
  Dim wo_rsDsgLst As MwfLib.MDORowSet
  Dim wCnd As String, wrepcnd As String, wsql As String, ws_DsgAnaCnd As String, ws_HDTPrdPtr As String
  Dim wi_Row As Integer, wHaveCnd As String, ws_RmWt As String, ws_RmQty As String, ws_DsgLnSvPtr As String
  
  Dim ws_DsgPrmCnd As String      '*** Jay 3.4[DP]
  
  Dim wi_MaxRecCnt As Integer
  
  Dim ws_DsgExcCnd As String      ' **** Manali 350Nxt - Valid value for wDmRegExc changed from Parameter 'YN' to 'DSGEXC'
  Dim ws_DsgLnSvPtrOnly As String
  
  wrepcnd = ADC.RepCond
  wCnd = IIF(wrepcnd <> "", " Where " + wrepcnd, "")
  
  ' ****** Manali 350Nxt - Valid value for wDmRegExc changed from Parameter 'YN' to 'DSGEXC'
  '578-12 exlusive option disabled
'  If adc("wDmRegExc") = "R" Then
'    ws_DsgExcCnd = " DmRegnExc = 'Y' "
'  ElseIf adc("wDmRegExc") = "C" Then
'    ws_DsgExcCnd = " DmRegnExc = 'N' "
'  ElseIf adc("wDmRegExc") = "F" Then
'    ws_DsgExcCnd = " DmRegnExc = '' "
'  ElseIf adc("wDmRegExc") = "A" Then
'    ws_DsgExcCnd = " DmRegnExc In  ('Y', 'N', '') "
'  End If

  'wCnd = wCnd + IIF(wCnd <> "", " and ", " Where ") + " DmRegnExc= '" + ADC("wDmRegExc") + "' "   '*** (Jen 11/06/07)
  '578-12 exlusive option disabled
  'wCnd = wCnd + IIF(wCnd <> "", " and ", " Where ") + ws_DsgExcCnd
  ' ****** Manali 350Nxt - Valid value for wDmRegExc changed from Parameter 'YN' to 'DSGEXC'
  
  If ADC("wPrdRmQtyWtYN") = "Y" Then
    ws_RmWt = "DrPrdWt"
    ws_RmQty = "DrPrdQty"
  Else
    ws_RmWt = "DrWt"
    ws_RmQty = "DrQty"
  End If
  
  '*** (Bef 09/06/07)
  'moCn.Execute ("Select UdAnaSr, UdAnaCdFr, UdAnaCdTo Into #TmpDsgAna " + _
  '                             " From UsrOptDsgAna Where 1= 2 ")
  'Dim i As Integer
  'i = 1
  'With GrdDsgAnaScop
  '  Do While i <= .Rows - 1
  '      If CStr(Trim(.Value(i, "UDANASR"))) <> "" Then _
  '      moCn.Execute ("Insert Into #TmpDsgAna values ('" + CStr(.Value(i, "UDANASR")) + "','" + CStr(.Value(i, "UDANACDFR")) + "','" + CStr(.Value(i, "UDANACDTO")) + "')")
  '      i = i + 1
  '  Loop
  'End With
  '
  'ws_DsgAnaCnd = " And (Select Count(*) From #TmpDsgAna " + _
  '               "Join DsgAna on DaAnaSr = UdAnaSr " + _
  '               "where (DaAnaCd between UdAnaCdFr and UdAnaCdTo) " + _
  '               "and DaTcTyp = DmTcTyp and DaCd= DmCd and DaSz= DmSz /* and " + _
  '               " UdUsrCd = '" + gs_UsrCd + "' and " + _
  '               "UdMnuCd= '" + ADC.MenuCd + "'*/)= " & (GrdDsgAnaScop.Rows - 1) & " "
  '
  '
  'moCn.Execute ("Select UodRmSCtg, UodPtrYN, UodLnSvPtrFr, UodLnSvPtrTo Into #TmpDsgSz From UsrOptDet Where 1=2 ")
  '
  'i = 1
  'With GrdDsgSz
  '  Do While i <= .Rows - 1
  '      If CStr(Trim(.Value(i, "UodRmSCtg"))) <> "" Then _
  '      moCn.Execute ("Insert Into #TmpDsgSz values ('" + CStr(.Value(i, "UodRmSCtg")) + "','" + CStr(.Value(i, "UodPtrYN")) + "'," + CStr(.Value(i, "UodLnSvPtrFr")) + "," + CStr(.Value(i, "UodLnSvPtrTo")) + ")")
  '      i = i + 1
  '  Loop
  'End With
  '
  'ws_HDTPrdPtr = UCase(moCn.GetFldVal("Select HDTPrdPtr from Head where HCoCd= '" + gs_CoCd + "'"))
  'ws_DsgLnSvPtr = " And (Select Count(*) From #TmpDsgSz where " + _
  '                "not exists(Select 'x' From DsgRm " + _
  '                "     join Param on PTyp= 'RMSCTG' and PMCd= DrRmCtg and PSCd= DrRmSCtg " + _
  '                "     where DrTcTyp= DmTcTyp and DrCd= DmCd and DrSz= DmSz " + _
  '                "     and DrRmCtg= 'D' and DrRmSCtg= UodRmSCtg " + _
  '                "     and (  (PValue3= 'Y' and (case when DrPrdQty> 0 and '" + ws_HDTPrdPtr + "'= 'Y' and PValue2= 'N' " + _
  '                "                              then DrPrdWt/DrPrdQty else DrRmPtr end) between UodLnSvPtrFr and UodLnSvPtrTo) " + _
  '                "            Or (PValue3= 'N' and DrLn1 between UodLnSvPtrFr and UodLnSvPtrTo) " + _
  '                "         ) " + _
  '                "     ))= 0 "
  '*** (Bef 09/06/07)
  
  '*** (Jen 09/06/07)
  ws_DsgAnaCnd = " And (Select Count('x') From UsrOptDsgAna " + _
                 "Join DsgAna on DaAnaSr = UdAnaSr " + _
                 "where (DaAnaCd between UdAnaCdFr and UdAnaCdTo) " + _
                 "and DaTcTyp = DmTcTyp and DaCd= DmCd and DaSz= DmSz and " + IIF(gs_Partition = ctCurrPrtn, " DaPrtKey=DmPrtKey And ", "") + _
                 " UdUsrCd = '" + gs_UsrCd + "' and " + _
                 "UdMnuCd= '" + ADC.MenuCd + "')= " & (GrdDsgAnaScop.Rows - 1) & " "
  
  ws_HDTPrdPtr = UCase(moCn.GetFldVal("Select HDTPrdPtr from Head where HCoCd= '" + gs_CoCd + "'"))
  ' ***** Manali 3.6.0 - 31/10/09 - Detail Scope applied for Colour Stones - RmCtg Condition added in Join (UodRmCtg)
  ws_DsgLnSvPtr = " And (Select Count('x') From UsrOptDet where " + _
                  "not exists(Select 'x' From DsgRm " + _
                  "     join Param on PTyp= 'RMSCTG' and PMCd= DrRmCtg and PSCd= DrRmSCtg " + _
                  "     where DrPrtKey=DmPrtKey and DrDmIdNo=DmIdNo and DrTcTyp= DmTcTyp and DrCd= DmCd and DrSz= DmSz " + _
                  "     and /*DrRmCtg= 'D'*/ DrRmCtg=UodRmCtg and DrRmSCtg= UodRmSCtg " + _
                  "     and (  (PValue3= 'Y' and (case when DrPrdQty> 0 and '" + ws_HDTPrdPtr + "'= 'Y' and PValue2= 'N' " + _
                  "                              then DrPrdWt/DrPrdQty else DrRmPtr end) between UodLnSvPtrFr and UodLnSvPtrTo) " + _
                  "            Or (PValue3= 'N' and DrLn1 between UodLnSvPtrFr and UodLnSvPtrTo) " + _
                  "         ) " + _
                  "     ) and UodUsrCd= '" + gs_UsrCd + "' and UodMnuCd= '" + ADC.MenuCd + "' )= 0 "
  '*** (Jen 09/06/07)
   If ADC("UOYN") = "Y" Then
    '****** Sachin 3.02 - Id fields in Join
    ' ***** Manali 3.6.0 - 31/10/09 - Detail Scope applied for Colour Stones - RmCtg Condition added in Join (UodRmCtg)
    ws_DsgLnSvPtrOnly = " and (Select count('x') From DsgRm " + _
                    " join Param on PTyp= 'RMSCTG' and PMCd= DrRmCtg and PSCd= DrRmSCtg " + _
                    " where DrDmIdNo=DmIdNo and DrTcTyp= DmTcTyp and DrCd= DmCd and DrSz= DmSz and DrRmCtg In  ('D','C') " + IIF(gs_Partition = ctCurrPrtn, " and DrPrtKey=DmPrtKey ", "") + _
                    " and not Exists(Select 'x' From UsrOptDet where DrRmCtg= UodRmCtg and DrRmSCtg= UodRmSCtg and " + _
                    " UodUsrCd = '" + gs_UsrCd + "' and UodMnuCd= '" + ADC.MenuCd + "' and " + _
                    " ( (PValue3= 'Y' and (case when DrPrdQty> 0 and '" + ws_HDTPrdPtr + "'= 'Y' and PValue2= 'N' " + _
                    "     then DrPrdWt/DrPrdQty else DrRmPtr end) between UodLnSvPtrFr and UodLnSvPtrTo) Or " + _
                    "   (PValue3= 'N' and DrLn1 between UodLnSvPtrFr and UodLnSvPtrTo)) )) = 0 "
  End If
  
  
  
  '*** Jay 3.4[DP]
  If Trim(ADC("wDpCollSel")) <> "" Then ADC("wDpCollSel") = "'" + Replace(ADC("wDpCollSel"), ",", "','") + "'"
  If Trim(ADC("wDpSerSel")) <> "" Then ADC("wDpSerSel") = "'" + Replace(ADC("wDpSerSel"), ",", "','") + "'"
  If Trim(ADC("wDpSetSel")) <> "" Then ADC("wDpSetSel") = "'" + Replace(ADC("wDpSetSel"), ",", "','") + "'"
  If Trim(ADC("wDpVerSel")) <> "" Then ADC("wDpVerSel") = "'" + Replace(ADC("wDpVerSel"), ",", "','") + "'"
  
  If ADC("wDpCollFr") <> "" Or ADC("wDpCollTo") <> "" Or ADC("wDpCollSel") <> "" Then
    ws_DsgPrmCnd = " And (Select count('x') from DsgPrm where DpTyp='COLL' and DmTcTyp='DM' and DpDmCd = DmCd and DmSz='' " + _
                IIF(ADC("wDpCollFr") <> "", " and DpCd>='" + ADC("wDpCollFr") + "'", "") + _
                IIF(ADC("wDpCollTo") <> "", " and DpCd<='" + ADC("wDpCollTo") + "'", "") + _
                IIF(ADC("wDpCollSel") <> "", " and DpCd in (" + ADC("wDpCollSel") + ") ", "") + ")> 0"
  End If
  If ADC("wDpSerFr") <> 0 Or ADC("wDpSerTo") <> 0 Or ADC("wDpSerSel") <> "" Then
    ws_DsgPrmCnd = ws_DsgPrmCnd + " And (Select count('x') from DsgPrm where DpTyp='SER' and DmTcTyp='DM' and DpDmCd = DmCd and DmSz='' " + _
                IIF(ADC("wDpSerFr") <> 0, " and Convert(Int, DpCd)>='" + CStr(ADC("wDpSerFr")) + "'", "") + _
                IIF(ADC("wDpSerTo") <> 0, " and Convert(Int, DpCd)<='" + CStr(ADC("wDpSerTo")) + "'", "") + _
                IIF(ADC("wDpSerSel") <> "", " and DpCd in (Select DpCd from DsgPrm where DpTyp= 'SER' and DpDmCd in (" + ADC("wDpSerSel") + ")) and DpDmCd<>''", "") + ")> 0"
  End If
  If ADC("wDpSetFr") <> 0 Or ADC("wDpSetTo") <> 0 Or ADC("wDpSetSel") <> "" Then
    ws_DsgPrmCnd = ws_DsgPrmCnd + " And (Select count('x') from DsgPrm where DpTyp='SET' and DmTcTyp='DM' and DpDmCd = DmCd and DmSz='' " + _
                IIF(ADC("wDpSetFr") <> 0, " and Convert(Int, DpCd)>='" + CStr(ADC("wDpSetFr")) + "'", "") + _
                IIF(ADC("wDpSetTo") <> 0, " and Convert(Int, DpCd)<='" + CStr(ADC("wDpSetTo")) + "'", "") + _
                IIF(ADC("wDpSetSel") <> "", " and DpCd in (Select DpCd from DsgPrm where DpTyp= 'SET' and DpDmCd in (" + ADC("wDpSetSel") + ")) and DpDmCd<>''", "") + ")> 0"
  End If
  If ADC("wDpVerFr") <> 0 Or ADC("wDpVerTo") <> 0 Or ADC("wDpVerSel") <> "" Then
    ws_DsgPrmCnd = ws_DsgPrmCnd + " And (Select count('x') from DsgPrm where DpTyp='VER' and DmTcTyp='DM' and DpDmCd = DmCd and DmSz='' " + _
                IIF(ADC("wDpVerFr") <> 0, " and Convert(Int, DpCd)>='" + CStr(ADC("wDpVerFr")) + "'", "") + _
                IIF(ADC("wDpVerTo") <> 0, " and Convert(Int, DpCd)<='" + CStr(ADC("wDpVerTo")) + "'", "") + _
                IIF(ADC("wDpVerSel") <> "", " and DpCd in (Select DpCd from DsgPrm where DpTyp= 'VER' and DpDmCd in (" + ADC("wDpVerSel") + ")) and DpDmCd<>'' ", "") + ")> 0"
  End If
  'Debug.Print ws_DsgPrmCnd
  '*** Jay 3.4[DP]
  '578-12 condition for customer code in DmAllow table added
  'Dim wCustCdCnd As String, wDmAllowCnd As String
  'wCustCdCnd = IIF(ADC("wCmCdFr") <> "", " and DmwCmCd >= '" + ADC("wCmCdFr") + "' ", "")
  'wCustCdCnd = wCustCdCnd + IIF(ADC("wCmCdTo") <> "", " and DmwCmCd <= '" + ADC("wCmCdTo") + "' ", "")
    
  'wDmAllowCnd = IIF(wCustCdCnd <> "", " and (Select isnull((Select top 1 DmwYN from DmAllow " + _
                  " Where DmwDmTyp = DmTcTyp and DmwDmCd= DmCd and DmwDmSz = DmSz and DmwTyp='CUST' " + _
                  wCustCdCnd + " and DmwYN ='Y'),isnull((Select top 1 DmwYN from DmAllow " + _
                  " Where DmwDmTyp = DmTcTyp and DmwDmCd= DmCd and DmwDmSz = DmSz and DmwTyp='CUST' " + _
                  wCustCdCnd + " and DmwYN ='N'), (Select DmwYN from DmAllow " + _
                  " Where DmwDmTyp = DmTcTyp and DmwDmCd= DmCd and DmwDmSz = DmSz and DmwTyp='CUST' " + _
                  " and DmwCmCd='ZSELF' and DmwYN ='Y') ))) = 'Y' ", "")
               
  
  If UCase(ADC("wMainMetWtFr")) > 0 Then wHaveCnd = wHaveCnd + " Having Max(Case when IsNull(DrMainMet, '') ='Y' then IsNull(" + ws_RmWt + ", 0) Else 0 End)>= " & ADC("wMainMetWtFr")
  If UCase(ADC("wMainMetWtTo")) > 0 Then wHaveCnd = wHaveCnd + IIF(wHaveCnd <> "", " and ", " Having ") + " Max(Case when IsNull(DrMainMet, '') ='Y' then IsNull(" + ws_RmWt + ", 0) Else 0 End)<= " & ADC("wMainMetWtTo")
  If UCase(ADC("wRmGrsWtFr")) > 0 Then wHaveCnd = wHaveCnd + IIF(wHaveCnd <> "", " and ", " Having ") + " convert(decimal(16, 3), Sum (IsNull(" + ws_RmWt + ",0)/ (case when IsNull(DrRmCtg,'') in ('C', 'D') then 5 else 1 end))) >= " & ADC("wRmGrsWtFr")
  If UCase(ADC("wRmGrsWtTo")) > 0 Then wHaveCnd = wHaveCnd + IIF(wHaveCnd <> "", " and ", " Having ") + " convert(decimal(16, 3), Sum (IsNull(" + ws_RmWt + ",0)/ (case when IsNull(DrRmCtg,'') in ('C', 'D') then 5 else 1 end))) <=" & ADC("wRmGrsWtTo")
  If UCase(ADC("wRmDiaWtFr")) > 0 Then wHaveCnd = wHaveCnd + IIF(wHaveCnd <> "", " and ", " Having ") + " convert(decimal(16, 3), Sum (" + ws_RmWt + "* (case when DrRmCtg = 'D' then 1 else 0 end))) >= " & ADC("wRmDiaWtFr")
  If UCase(ADC("wRmDiaWtTo")) > 0 Then wHaveCnd = wHaveCnd + IIF(wHaveCnd <> "", " and ", " Having ") + " convert(decimal(16, 3), Sum (" + ws_RmWt + "* (case when DrRmCtg = 'D' then 1 else 0 end))) <= " & ADC("wRmDiaWtTo")
  If UCase(ADC("wRmCSWtFr")) > 0 Then wHaveCnd = wHaveCnd + IIF(wHaveCnd <> "", " and ", " Having ") + " convert(decimal(16, 3), Sum (" + ws_RmWt + "* (case when DrRmCtg = 'C' then 1 else 0 end))) >= " & ADC("wRmCSWtFr")
  If UCase(ADC("wRmCSWtTo")) > 0 Then wHaveCnd = wHaveCnd + IIF(wHaveCnd <> "", " and ", " Having ") + " convert(decimal(16, 3), Sum (" + ws_RmWt + "* (case when DrRmCtg = 'C' then 1 else 0 end))) <= " & ADC("wRmCSWtTo")
  If UCase(ADC("wRmDiaQtyFr")) > 0 Then wHaveCnd = wHaveCnd + IIF(wHaveCnd <> "", " and ", " Having ") + " convert(decimal(16, 3), Sum (" + ws_RmQty + "* (case when DrRmCtg = 'D' then 1 else 0 end))) >= " & ADC("wRmDiaQtyFr")
  If UCase(ADC("wRmDiaQtyTo")) > 0 Then wHaveCnd = wHaveCnd + IIF(wHaveCnd <> "", " and ", " Having ") + " convert(decimal(16, 3), Sum (" + ws_RmQty + "* (case when DrRmCtg = 'D' then 1 else 0 end))) <= " & ADC("wRmDiaQtyTo")
  If UCase(ADC("wRmCSQtyFr")) > 0 Then wHaveCnd = wHaveCnd + IIF(wHaveCnd <> "", " and ", " Having ") + " convert(decimal(16, 3), Sum (" + ws_RmQty + "* (case when DrRmCtg = 'C' then 1 else 0 end))) >= " & ADC("wRmCSQtyFr")
  If UCase(ADC("wRmCSQtyTo")) > 0 Then wHaveCnd = wHaveCnd + IIF(wHaveCnd <> "", " and ", " Having ") + " convert(decimal(16, 3), Sum (" + ws_RmQty + "* (case when DrRmCtg = 'C' then 1 else 0 end))) <= " & ADC("wRmCSQtyTo")

  '*** Jay 3.4[DP] {DmCreatedDt, ws_DsgPrmCnd - Variable added}
  '4.1.3.0
  wsql = "Select Max(DmPrtKey) As DmPrtKey, Max(DmTcTyp) As DmTcTyp, Max(DmCtg) As DmCtg, DmCd, DmSz, Max(DmUom) As DmUom, " + _
         " Max(DmDesc) As DmDesc, Max(DmPrdCtg) As DmPrdCtg, Max(DmSalCtg) As DmSalCtg, " + _
         " Max(DmVaCtg) As DmVaCtg, Max(DmLsCtg) As DmLsCtg, Max(DmKt) As DmKt, Max(DmBagPcs) As DmBagPcs, " + _
         " Max(DmSetCd) As DmSetCd, Max(DmOldCd) As DmOldCd, Max(DmCol) As DmCol, Max(DmCmCd) As DmCmCd, " + _
         " Max(DmRegnExc) As DmRegnExc, Max(DmParts) As DmParts, Max(DmPartDesc) As DmPartDesc, " + _
         " Max(DmValidYn) As DmValidYn, Max(DmHld) As DmHld, Max(DmHldDesc) As DmHldDesc, Max(DmPrdSeq) As DmPrdSeq, Max(DmPrdInst) As DmPrdInst, " + _
         " Max(DmDsgDt) As DmDsgDt, Max(DmDsgBy) As DmDsgBy, Max(DmModMkr) As DmModMkr, Max(DmLstMdf) As DmLstMdf, " + _
         " Max(DmWaxWt) As DmWaxWt, Max(DmModRunWt) As DmModRunWt, Max(DmSilModWt) As DmSilModWt, Max(DmCasPcWt) As DmCasPcWt, Max(DmCreatedDt) as DmCreatedDt " + _
         " From DsgMst " + _
         " Left Outer Join DsgRm Dr on DmTcTyp = DrTcTyp and DmCd = DrCd and DmSz = DrSz " + IIF(gs_Partition = ctCurrPrtn, " And DmPrtKey=DrPrtKey ", "") + _
         wCnd + IIF(gs_Partition = ctCurrPrtn, " And DmPrtKey='" + ctCurrPrtn + "' ", "") + ws_DsgAnaCnd + ws_DsgLnSvPtr + ws_DsgLnSvPtrOnly + ws_DsgPrmCnd + " Group By DmTcTyp, DmCd, DmSz " + wHaveCnd + " Order By DmCtg, DmCd, DmSz /* ; Drop Table #TmpDsgAna; Drop Table #TmpDsgSz */ "
  'Debug.Print wsql
  
  GrdDsgLst.Rows = 1
  Set wo_rsDsgLst = moCn.OpenRes(wsql)
  
  ' ****** Manali 3.5.0 - 17/12/08 - Max limit for records checked from Head
  wi_MaxRecCnt = moCn.GetFldVal("Select HMaxDsgViewRec from Head where HCoCd='" + ctSelfCoCd + "' and HCd='" + ctSelfCmCd + "' ")
  
  If wo_rsDsgLst.RecCount <= wi_MaxRecCnt Then
  'If wo_rsDsgLst.RecCount <= 10000 Then     ' ****** Manali 3.4.1 - Limit extended to 9999  [For Shrenuj Only]
  'If wo_rsDsgLst.RecCount <= 1000 Then
    GrdDsgLst.AllowAdd = False
    With GrdDsgLst
      Do While Not (wo_rsDsgLst.EOF Or wo_rsDsgLst.BOF)
        .AddItem
        wi_Row = .Rows - 1
        .Value(wi_Row, "DmPrtKey") = Trim(wo_rsDsgLst!DmPrtKey)   '****** Sachin 3.02.0
        .Value(wi_Row, "DmTcTyp") = Trim(wo_rsDsgLst!DmTcTyp)
        .Value(wi_Row, "DmCtg") = Trim(wo_rsDsgLst!DmCtg)
        .Value(wi_Row, "DmCd") = Trim(wo_rsDsgLst!DmCd)
        .Value(wi_Row, "DmSz") = Trim(wo_rsDsgLst!DmSz)
        .Value(wi_Row, "DmUom") = Trim(wo_rsDsgLst!DmUom)
        '*** Jay 3.4 [DP]
        .Value(wi_Row, "wDiaWt") = moCn.GetFldVal("Select cast(IsNull(sum(case when DrRmCtg = 'D' then DrPrdWt else 0 end), 0) as Decimal(8, 3)) " + _
                           " From DsgRm Where DrTcTyp ='" + .Value(wi_Row, "DMTCTYP") + "'" + _
                           " And DrCd = '" + .Value(wi_Row, "DMCD") + "'  And DrSz ='" + .Value(wi_Row, "DMSZ") + "' " + IIF(gs_Partition = ctCurrPrtn, " And DrPrtKey='" + ctCurrPrtn + "' ", ""))
        '*** Jay 3.4 [DP]
        .Value(wi_Row, "wGrsWt") = moCn.GetFldVal("Select cast(IsNull(sum(DrPrdWt/ (case when DrRmCtg in ('C', 'D') then 5 else 1 end)), 0) as Decimal(8, 3)) " + _
                           " From DsgRm Where DrTcTyp ='" + .Value(wi_Row, "DMTCTYP") + "'" + _
                           " And DrCd = '" + .Value(wi_Row, "DMCD") + "'  And DrSz ='" + .Value(wi_Row, "DMSZ") + "' " + IIF(gs_Partition = ctCurrPrtn, " And DrPrtKey='" + ctCurrPrtn + "' ", ""))
        
        .Value(wi_Row, "DmDesc") = Trim(wo_rsDsgLst!DmDesc)
        .Value(wi_Row, "DmPrdCtg") = Trim(wo_rsDsgLst!DmPrdCtg)
        .Value(wi_Row, "DmSalCtg") = Trim(wo_rsDsgLst!DmSalCtg)
        .Value(wi_Row, "DmVaCtg") = Trim(wo_rsDsgLst!DmVaCtg)
        .Value(wi_Row, "DmLsCtg") = Trim(wo_rsDsgLst!DmLsCtg)
        .Value(wi_Row, "DmKt") = Trim(wo_rsDsgLst!DmKt)
        .Value(wi_Row, "DmBagPcs") = Trim(wo_rsDsgLst!DmBagPcs)
        .Value(wi_Row, "DmSetCd") = Trim(wo_rsDsgLst!DmSetCd)
        .Value(wi_Row, "DmOldCd") = Trim(wo_rsDsgLst!DmOldCd)
        .Value(wi_Row, "DmCol") = Trim(wo_rsDsgLst!DmCol)
        .Value(wi_Row, "DmCmCd") = Trim(wo_rsDsgLst!DmCmCd)
        .Value(wi_Row, "DmRegnExc") = Trim(wo_rsDsgLst!DmRegnExc)
        .Value(wi_Row, "DmParts") = Trim(wo_rsDsgLst!DmParts)
        .Value(wi_Row, "DmPartDesc") = Trim(wo_rsDsgLst!DmPartDesc)
        .Value(wi_Row, "DmValidYN") = Trim(wo_rsDsgLst!DmValidYN)       '4.1.3.0
        .Value(wi_Row, "DmHld") = Trim(wo_rsDsgLst!DmHld)
        .Value(wi_Row, "DmHldDesc") = Trim(wo_rsDsgLst!DmHldDesc)
        .Value(wi_Row, "DmPrdSeq") = Trim(wo_rsDsgLst!DmPrdSeq)
        .Value(wi_Row, "DmPrdInst") = Trim(wo_rsDsgLst!DmPrdInst)
        .Value(wi_Row, "DmDsgDt") = Trim(wo_rsDsgLst!DmDsgDt)
        .Value(wi_Row, "DmDsgBy") = Trim(wo_rsDsgLst!DmDsgBy)
        .Value(wi_Row, "DmModMkr") = Trim(wo_rsDsgLst!DmModMkr)
        .Value(wi_Row, "DmLstMdf") = Trim(wo_rsDsgLst!DmLstMdf)
        .Value(wi_Row, "DmWaxWt") = Trim(wo_rsDsgLst!DmWaxWt)
        .Value(wi_Row, "DmModRunWt") = Trim(wo_rsDsgLst!DmModRunWt)
        .Value(wi_Row, "DmSilModWt") = Trim(wo_rsDsgLst!DmSilModWt)
        .Value(wi_Row, "DmCasPcWt") = Trim(wo_rsDsgLst!DmCasPcWt)
        .Value(wi_Row, "DmCreatedDt") = Trim(wo_rsDsgLst!DmCreatedDt)   '*** Jay 3.4[DP]
        wo_rsDsgLst.MoveNext
      Loop
    End With
  Else
    ' ****** Manali 3.5.0 - 17/12/08 - Max limit for records checked from Head
    DispMsg "No. of Records are more than " + CStr(wi_MaxRecCnt) + ", Plz Limit the Scope", etError
    'DispMsg "No. of Records are more than 10000, Plz Limit the Scope", etError      ' ****** Manali 3.4.1 - Limit extended to 9999
    'DispMsg "No. of Records are more than 1000, Plz Limit the Scope", etError
  End If
  Set wo_rsDsgLst = Nothing

  If GrdDsgLst.Rows > 1 Then
    GrdDsgLst.Row = 1: GrdDsgLst.Col = 1: GrdDsgLst.SetFocus
    'CmdSumm.Visible = True   '*** (Bef 08/06/07)
  End If
End Sub

Private Sub FillDsgDets(ByVal RowNum As Integer)

        Dim wo_rsDsgDet As MwfLib.MDORowSet, wsql As String
    
        wsql = "Select DmHistory, DmFaults  From DsgMst " + _
                 " Where DmTcTyp='" + GrdDsgLst.Value(RowNum, "DMTCTYP") + "'" + _
                 "  And DmCd='" + GrdDsgLst.Value(RowNum, "DMCD") + "' And DmSz='" + GrdDsgLst.Value(RowNum, "DMSZ") + "' " + IIF(gs_Partition = ctCurrPrtn, " And DmPrtKey='" + ctCurrPrtn + "'", "")
        Set wo_rsDsgDet = moCn.OpenRes(wsql)
      
        If Not (wo_rsDsgDet.EOF Or wo_rsDsgDet.BOF) Then
              ADC("DmHistory") = wo_rsDsgDet!DmHistory
              ADC("DmFaults") = wo_rsDsgDet!DmFaults
        End If
            
        Set wo_rsDsgDet = Nothing

        With GrdDsgLst
            ADC("wDmUom") = .Value(.Row, "DmUom")       '*** Jay 3.4[DP]
            ADC("wDmPrdCtg") = .Value(.Row, "DmPrdCtg")
            ADC("wDmSalCtg") = .Value(.Row, "DmSalCtg")
            ADC("wDmVaCtg") = .Value(.Row, "DmVaCtg")
            ADC("wDmLsCtg") = .Value(.Row, "DmLsCtg")
            ADC("wDmKt") = .Value(.Row, "DmKt")
            ADC("wDmBagPcs") = .Value(.Row, "DmBagPcs")
            ADC("wDmSetCd") = .Value(.Row, "DmSetCd")
            ADC("wDmOldCd") = .Value(.Row, "DmOldCd")
            ADC("wDmCol") = .Value(.Row, "DmCol")
            ADC("wDmCmCd") = .Value(.Row, "DmCmCd")
            ADC("wDmRegnExc") = .Value(.Row, "DmRegnExc")
            ADC("wDmParts") = .Value(.Row, "DmParts")
            ADC("wDmPartDesc") = .Value(.Row, "DmPartDesc")
            ADC("wDmHld") = .Value(.Row, "DmHld")
            ADC("wDmValidYn") = .Value(.Row, "DmValidYn")       '4.1.3.0
            ADC("wDmHldDesc") = .Value(.Row, "DmHldDesc")
            ADC("wDmPrdSeq") = .Value(.Row, "DmPrdSeq")
            ADC("wDmPrdInst") = .Value(.Row, "DmPrdInst")
            ADC("wDmDsgDt") = .Value(.Row, "DmDsgDt")
            ADC("wDmDsgBy") = .Value(.Row, "DmDsgBy")
            ADC("wDmModMkr") = .Value(.Row, "DmModMkr")
            ADC("wDmLstMdf") = .Value(.Row, "DmLstMdf")
            ADC("wDmWaxWt") = .Value(.Row, "DmWaxWt")
            ADC("wDmModRunWt") = .Value(.Row, "DmModRunWt")
            ADC("wDmSilModWt") = .Value(.Row, "DmSilModWt")
            ADC("wDmCasPcWt") = .Value(.Row, "DmCasPcWt")
            ADC("wDmCreatedDt") = .Value(.Row, "DmCreatedDt") '*** Jay 3.4[DP]
        End With
        
End Sub

Private Sub atxt_ReCalculate(Index As Integer, ByVal IdName As String, ByVal OldVal As Variant, ByVal RowNum As Integer, ByVal SourceName As String, ByVal SourceColName As String)

Dim wCmCd As String, wCustName As String

Select Case UCase(IdName)
    Case Is = UCase("DmHldDesc")
        If ADC.Mode = xNorm Then
          If ADC("wDmHld") = "N" Or Trim(ADC("wDmHld")) = "" Then ADC("wDmHldDesc") = ""
        End If
    
    Case Is = UCase("wCmRegnCd")
      '****** Sachin 3.3.0 (27-06-08) - [37. Design Exclusivity for Multiple Customers] ******
      Dim wCustRegn() As String
      wCustRegn = Split(ADC("wDmCmCd"), ",")
        If (ADC.Mode = xNorm Or ADC.Mode = xInit) And Trim(ADC("wDmCmCd")) <> "" Then _
           ADC("wCmRegnCd") = moCn.GetFldVal("Select CmRegnCd From CustMst " + _
                              "where CmCtg= 'C' and CmCd='" + wCustRegn(0) + "'")
      '****** Sachin 3.3.0 (27-06-08) - [37. Design Exclusivity for Multiple Customers] ******
        
    Case Is = UCase("wDmRegnExc")
      If ADC.Mode = xNorm Then _
         If ADC("wDmCmCd") = "" Then ADC("wDmRegnExc") = ""
    
    '****** Sachin 3.3.0 (27-06-08) - [37. Design Exclusivity for Multiple Customers] ******
    '    Case Is = UCase("wDmCmCd")
    '      If adc.Mode = xNorm Then
    '        wCmCd = moCn.GetFldVal("Select CmCd from CustMst where CmCtg='C' and " + _
    '                               "CmName= '" + adc("wCmName") + "'")
    '        If wCmCd <> adc("wDmCmCd") Then adc("wDmCmCd") = wCmCd
    '      End If
    '
    '    Case Is = UCase("wCmName")
    '      If adc.Mode = xNorm Then
    '         wCustName = moCn.GetFldVal("Select CmName from CustMst where CmCtg='C' and " + _
    '                                    "CmCd='" + adc("wDmCmCd") + "'")
    '         If wCustName <> adc("wCmName") Then adc("wCmName") = wCustName
    '      ElseIf adc.Mode = xInit Then
    '         adc("wCmName") = moCn.GetFldVal("Select CmName from CustMst where CmCtg='C' and " + _
    '                                    "CmCd='" + adc("wDmCmCd") + "'")
    '      End If
    '****** Sachin 3.3.0 (27-06-08) - [37. Design Exclusivity for Multiple Customers] ******

End Select

End Sub

Private Sub ChkBoxArr_Click(Index As Integer)
With ChkBoxArr(Index)
  Select Case Index
  Case Is = 0
    If .Value = Checked Then ADC("UoYN") = "Y" Else ADC("UoYN") = "N"
  End Select
End With
End Sub
' ***** Manali 3.5.0 - 18/11/08 - Check Boxes Added

' ***** Manali 3.5.0 - 18/11/08 - Check Boxes Added
Private Sub ChkBoxArr_GotFocus(Index As Integer)
  DispMsg ChkBoxArr(Index).ToolTipText, etInfo
End Sub

Private Sub CmdCpyCat_Click()
  Call DispFra(CatCpy)
End Sub

Private Sub CmdCpyCatGo_Click()
  If ADC("wDpCd") = "" Then DispMsg "Catalogue Code Can Not Be Blank", etError: Exit Sub
  Call CpyCat
  Call CmdCpyCat_Click
End Sub

Private Sub CmdDetScope_Click()
  If FraSC.Visible = False Then
    FraSC.Enabled = True
    FraSC.Visible = True
    FraScop.Enabled = False
    'GrdDsgAnaScop.AllowDelete = True   '*** (Bef 09/06/07)
    FraSC.ZOrder
    ADC.AllowFind = False     '*** (Jen 09/06/07)
  Else
    FraSC.Enabled = False
    FraSC.Visible = False
    FraScop.Enabled = True
    ADC.AllowFind = True      '*** (Jen 09/06/07)
  End If
End Sub

'*** (Bef 08/06/07)
'Private Sub CmdSumm_Click()
'  ADC("wShowPrdRmDets") = "Y"
'  Call DispFra(Summ)
'  If FraSumm.Visible = True Then Call Summary(ADC("wShowPrdRmDets"))
'End Sub
'*** (Bef 08/06/07)

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

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If UnloadMode <> vbFormCode Then Cancel = True: Exit Sub
  Set moCn = Nothing    '*** (09/08/05)    '*** (Bef speed) bef 24/11/06
End Sub
'578-12
Private Sub GrdDsgAllow_ColWhen(ByVal RowNum As Integer, ByVal ColName As String, Cancel As Boolean, ErrMsg As String, ByVal Action As MwfCtl.en_fgAction)
  Cancel = True
End Sub

Private Sub GrdDsgAllow_SetRecSource()
 ''578-12*** Set the RecSource for GrdDsgAllow
  If GrdDsgLst.Row > 0 Then _
  GrdDsgAllow.RecSource = " Select * from DmAllow " + _
                        " Where DmwDmTyp ='" + GrdDsgLst.Value(GrdDsgLst.Row, "DMTCTYP") + "'" + _
                        " And DmwDmCd = '" + GrdDsgLst.Value(GrdDsgLst.Row, "DMCD") + "'" + _
                        " And DmwDmSz ='" + GrdDsgLst.Value(GrdDsgLst.Row, "DMSZ") + "' "
                        
End Sub

Private Sub GrdDsgAna_ColWhen(ByVal RowNum As Integer, ByVal ColName As String, Cancel As Boolean, ErrMsg As String, ByVal Action As MwfCtl.en_fgAction)
    Cancel = True
End Sub
Private Sub GrdDsgAna_ReCalc(ByVal RowNum As Integer, ByVal ColName As String, ByVal OldVal As Variant, ByVal SourceName As String, ByVal SourceColName As String)
  '*** wAnaName on DaAnaSr             Normal
      'PDesc225('DAANAFLD', DaAnaSr)
  
  '*** (Jen 11/06/07)
  With GrdDsgAna
    Select Case UCase(ColName)
    Case Is = UCase("wAnaName")
      If .Mode = fgminit Then
        .Value(RowNum, "wAnaName") = moCn.GetFldVal("Select PDesc225 from Param where " + _
                                     "PTyp= 'DAANAFLD' and PMCd= '" + .Value(RowNum, "DaAnaSr") + "' ")
      End If
    End Select
  End With
  '*** (Jen 11/06/07)
End Sub
Private Sub GrdDsgAnaScop_ColChange(ByVal RowNum As Integer, ByVal ColName As String, ByVal NewValue As Variant, Cancel As Boolean, ErrMsg As String)
  With GrdDsgAnaScop
    Select Case UCase(ColName)
    Case Is = UCase("UdAnaSr")
      Cancel = NewValue <> "" And Not moCn.RecSeek("Select PMCd From Param Where " + _
                                "PTyp= 'DAANAFLD' And PMCd = '" + NewValue + "'")
      ErrMsg = "Invalid Analysis Sr.": Exit Sub
    End Select
  End With
End Sub

Private Sub GrdDsgAnaScop_ColWhen(ByVal RowNum As Integer, ByVal ColName As String, Cancel As Boolean, ErrMsg As String, ByVal Action As MwfCtl.en_fgAction)
  Select Case ColName
  Case Is = UCase("UdAnaSr")
    Call HlpList.PMCd("DAANAFLD")
  Case Is = UCase("UdAnaCdFr"), UCase("UdAnaCdTo")
    If moCn.GetFldVal("Select PValue from Param where PTyp= 'DAANAFLD' and " + _
       "PMCd= '" + GrdDsgAnaScop.Value(RowNum, "UdAnaSr") + "' ") = "N" Then _
       Cancel = True: ErrMsg = "Cannot Enter Code For This Analysis Field": Exit Sub
    Call HlpList.PSCd("DAANACD", GrdDsgAnaScop.Value(RowNum, "UdAnaSr"))
  End Select

End Sub
Private Sub GrdDsgAnaScop_InitKey(ByVal RowNum As Integer)
  With GrdDsgAnaScop
    If .IsNew(RowNum) Then
      .Store "UdUsrCd", gs_UsrCd
      .Store "UdMnuCd", ADC.MenuCd
    End If
  End With
End Sub
Private Sub GrdDsgAnaScop_ReCalc(ByVal RowNum As Integer, ByVal ColName As String, ByVal OldVal As Variant, ByVal SourceName As String, ByVal SourceColName As String)
  With GrdDsgAnaScop
    Select Case UCase(ColName)
    Case Is = UCase("UdAnaCdFr")
      If .Mode = fgmnorm Then .Value(RowNum, "UdAnaCdFr") = ""
    Case Is = UCase("UdAnaCdTo")
      If .Mode = fgmnorm Then .Value(RowNum, "UdAnaCdTo") = ""
    End Select
  End With
End Sub
Private Sub GrdDsgAnaScop_SetRecSource()
  '*** Set the Record Source of the Grid GrdDsgAna
  GrdDsgAnaScop.RecSource = "Select * from UsrOptDsgAna where UdUsrCd= '" + gs_UsrCd + "' and " + _
                            "UdMnuCd= '" + ADC.MenuCd + "' Order By UdAnaSr"
End Sub
Private Sub GrdDsgAnaScop_Validate(Cancel As Boolean)
  Cancel = GrdDsgAnaScop.Validate
End Sub
Private Sub GrdDsgCT_ColWhen(ByVal RowNum As Integer, ByVal ColName As String, Cancel As Boolean, ErrMsg As String, ByVal Action As MwfCtl.en_fgAction)
    Cancel = True
End Sub
Private Sub GrdDsgLab_ColWhen(ByVal RowNum As Integer, ByVal ColName As String, Cancel As Boolean, ErrMsg As String, ByVal Action As MwfCtl.en_fgAction)
  Cancel = True
End Sub
Private Sub GrdDsgLab_ReCalc(ByVal RowNum As Integer, ByVal ColName As String, ByVal OldVal As Variant, ByVal SourceName As String, ByVal SourceColName As String)
  With GrdDsgLab
    Select Case UCase(ColName)
      Case Is = UCase("wQW")
        '*** (Jen 08/06/07)
        If .Mode = fgminit Then
          .Value(RowNum, "wQW") = moCn.GetFldVal("Select PValue from Param where PTyp= 'LABSCD' " + _
                                  "and PMCd= '" + .Value(RowNum, "DlMCd") + "' and " + _
                                  "PSCd= '" + .Value(RowNum, "DlSCd") + "'")
        End If
        '*** (Jen 08/06/07)
    End Select
  End With
End Sub
'*** Jay 3.4[DP]
Private Sub GrdDsgLst_ColChange(ByVal RowNum As Integer, ByVal ColName As String, ByVal NewValue As Variant, Cancel As Boolean, ErrMsg As String)
  Select Case UCase(ColName)
    '*** Jay 3.4[DP]
    Case Is = UCase("WSel")
      Cancel = NewValue <> "" And (Not moCn.RecSeek("select PMCd from Param where " + _
                             " PTyp= 'YN' and PMCd = '" + NewValue + "'"))
      If NewValue <> "Y" Then
        ms_DmCdStr = Replace(ms_DmCdStr, "," + GrdDsgLst.Value(RowNum, "DmCd"), "")
        ErrMsg = "Enter Selection As Y to Selcting Design Or N/Blank for Not To Select": Exit Sub
      Else
        If InStr(1, ms_DmCdStr, GrdDsgLst.Value(RowNum, "DmCd")) <= 0 Then _
            ms_DmCdStr = ms_DmCdStr + "," + GrdDsgLst.Value(RowNum, "DmCd")
      End If
    '*** Jay 3.4[DP]
  End Select
End Sub
'*** Jay 3.4[DP]

Private Sub GrdDsgLst_ColWhen(ByVal RowNum As Integer, ByVal ColName As String, Cancel As Boolean, ErrMsg As String, ByVal Action As MwfCtl.en_fgAction)
  '*** Jay 3.4[DP]
  If UCase(ColName) <> "WSEL" Then
    Cancel = True: ErrMsg = "Total Records : " + CStr(GrdDsgLst.Rows - 1)
  Else
    GrdDsgLst.Value(RowNum, "WSEL") = IIF(Trim(GrdDsgLst.Value(RowNum, "WSEL")) = "", "Y", "")
  End If
  '*** Jay 3.4[DP]
End Sub

Private Sub GrdDsgLst_RowWhen(ByVal RowNum As Integer)
    If RowNum > 0 Then Call FillDsgDets(RowNum)
    
    'Manali 3.2.2 - Pic large/small
    FraPic3D.Left = wFraPic3DLft: FraPic3D.Top = wFraPic3DTop
    FraPic3D.Height = wFraPic3DHt: FraPic3D.Width = wFraPic3DWdt
    FraPicHsk.Left = wFraPicHskLft: FraPicHsk.Top = wFraPicHskTop
    FraPicHsk.Height = wFraPicHskHt: FraPicHsk.Width = wFraPicHskWdt
    'sv.75 "DM" changed as ADC("wDmTyp")
    wHskPth = GetPictPath(ADC("wDmTyp"), True, GrdDsgLst.Value(RowNum, "DMCD"), "")
    w3DPth = GetPictPath(ADC("wDmTyp"), False, GrdDsgLst.Value(RowNum, "DMCD"), "")
  
    Call ShowPic(wPicHskHt, wPicHskWdt, wHskPth, PicHsk)
    Call ShowPic(wPic3DHt, wPic3DWdt, w3DPth, Pic3D)
    
    '*** (Jen) 08/06/07
    If ADC("wShowPrdRmDets") = "" Then ADC("wShowPrdRmDets") = "Y"
    If TabDesgnDet.Tab = 8 Then Call Summary(ADC("wShowPrdRmDets"))
        
    With TabDesgnDet
      If ADC("wDmTyp") = "DM" And GrdDsgLst.Value(RowNum, "DmSz") = "" Then
        Dim wi_Tabs As Integer
        wi_Tabs = .Tabs
        
        .TabVisible(6) = True: .TabEnabled(6) = True
        .TabVisible(9) = True: .TabEnabled(9) = True    '*** Jay 3.4[DP]
        If mb_CTMod = False Then
          GrdDsgCT.ColProp("DctChr").Style = fgcinvisible
          GrdDsgCT.ColProp("DctRmCd").Style = fgcinvisible
          GrdDsgCT.ColProp("DctDmCol").Style = fgcinvisible
          GrdDsgCT.ColProp("DctRmWt").Style = fgcinvisible
          GrdDsgCT.ColProp("DctDc").Style = fgcinvisible
          GrdDsgCT.ColProp("DctPrdSeq").Style = fgcinvisible
          GrdDsgCT.ColProp("DctDesc").Style = fgcinvisible
        Else
          GrdDsgCT.ColProp("DctChr").Style = fgcNormal
          GrdDsgCT.ColProp("DctRmCd").Style = fgcNormal
          GrdDsgCT.ColProp("DctDmCol").Style = fgcNormal
          GrdDsgCT.ColProp("DctRmWt").Style = fgcNormal
          GrdDsgCT.ColProp("DctDc").Style = fgcNormal
          GrdDsgCT.ColProp("DctPrdSeq").Style = fgcNormal
          GrdDsgCT.ColProp("DctDesc").Style = fgcNormal
        End If
        
        If mb_ModAvl = True Then
          .TabVisible(7) = True: .TabEnabled(7) = True
        Else
          .TabVisible(7) = False: .TabEnabled(7) = False
          wi_Tabs = wi_Tabs - 1
        End If
        
        ' ***** Manali - 22/10/08 - Design Catalogue Module
        If mb_DsgCat = True Then
          .TabVisible(9) = True: .TabEnabled(9) = True
        Else
          .TabVisible(9) = False: .TabEnabled(9) = False
          wi_Tabs = wi_Tabs - 1
        End If
        ' ***** Manali - 22/10/08 - Design Catalogue Module

        .TabsPerRow = wi_Tabs
      Else
        .TabVisible(6) = False: .TabEnabled(6) = False
        .TabVisible(7) = False: .TabEnabled(7) = False
        .TabVisible(9) = False: .TabEnabled(9) = False    '*** Jay 3.4[DP]
        .TabsPerRow = 8   '578-12 changed 7 to 8, included Allow tab
      End If
    End With
    '*** (Jen) 08/06/07
    
End Sub
Private Sub GrdDsgLst_Validate(Cancel As Boolean)
  Cancel = GrdDsgLst.Validate
End Sub
Private Sub GrdDsgModAvl_ColWhen(ByVal RowNum As Integer, ByVal ColName As String, Cancel As Boolean, ErrMsg As String, ByVal Action As MwfCtl.en_fgAction)
  Cancel = True
End Sub
Private Sub GrdDsgModAvl_SetRecSource() ' ***** Manali - 22/10/08 - Design Catalogue Module

  If GrdDsgLst.Row > 0 Then _
  GrdDsgModAvl.RecSource = " Select * from DsgModAvl " + _
                        " Where DmaTcTyp ='" + GrdDsgLst.Value(GrdDsgLst.Row, "DMTCTYP") + "'" + _
                        " And DmaCd = '" + GrdDsgLst.Value(GrdDsgLst.Row, "DMCD") + "'" + _
                        " And DmaSz ='" + GrdDsgLst.Value(GrdDsgLst.Row, "DMSZ") + "' "
End Sub

Private Sub GrdDsgPrm_ColWhen(ByVal RowNum As Integer, ByVal ColName As String, Cancel As Boolean, ErrMsg As String, ByVal Action As MwfCtl.en_fgAction)
'*** Jay 3.4[DP]
  Cancel = True
End Sub
'*** Jay 3.4[DP]

Private Sub GrdDsgPrm_SetRecSource()
'*** Jay 3.4[DP]
  If GrdDsgLst.Row > 0 And GrdDsgLst.Value(GrdDsgLst.Row, "DMTCTYP") = "DM" Then
  'If GrdDsgLst.Row > 0 Then
    GrdDsgPrm.RecSource = " Select * from DsgPrm " + _
                       " Where DpDmCd = '" + GrdDsgLst.Value(GrdDsgLst.Row, "DMCD") + "' and DpTyp <>'CAT' Order By DpTyp, DpCd"
  End If
End Sub
'*** Jay 3.4[DP]

Private Sub GrdDsgPrmDet_ReCalc(ByVal RowNum As Integer, ByVal ColName As String, ByVal OldVal As Variant, ByVal SourceName As String, ByVal SourceColName As String)
'*** Jay 3.4[DP]
  Select Case UCase(ColName)
  Case Is = UCase("wDpSel")
    If ms_DmCdStr <> "" And GrdDsgPrmDet.Mode = fgminit Then
      GrdDsgPrmDet.Value(RowNum, "wDpSel") = IIF(InStr(1, ms_DmCdStr, GrdDsgPrmDet.Value(RowNum, "DpDmCd")) > 0, "Y", "")
    End If
  End Select
'*** Jay 3.4[DP]
End Sub

Private Sub GrdDsgPrmDet_RowWhen(ByVal RowNum As Integer)
'*** Jay 3.4[DP]
  Dim wDpPth As String
  wDpPth = GetPictPath("DM", False, GrdDsgPrmDet.Value(RowNum, "DpDmCd"))
  Call ShowPic(wPicDPHt, wPicDPWdt, wDpPth, PicDP)
End Sub
'*** Jay 3.4[DP]

Private Sub GrdDsgPrmDet_SetRecSource()
'*** Jay 3.4[DP]
  If GrdDsgPrm.Row > 0 Then
    GrdDsgPrmDet.RecSource = " Select * from DsgPrm " + _
                         " Where DpTyp ='" + GrdDsgPrm.Value(GrdDsgPrm.Row, "DpTyp") + "' " + _
                         " And DpCd ='" + GrdDsgPrm.Value(GrdDsgPrm.Row, "DpCd") + "' " + _
                         " And DpDmCd not In ('','" + GrdDsgLst.Value(GrdDsgLst.Row, "DMCD") + "') Order By DpDmCd"
  End If
End Sub
'*** Jay 3.4[DP]

'*** Jay 3.4[DP]
Private Sub GrdDsgPrmDet_ColChange(ByVal RowNum As Integer, ByVal ColName As String, ByVal NewValue As Variant, Cancel As Boolean, ErrMsg As String)
  Select Case UCase(ColName)
    '*** Jay 3.4[DP]
    Case Is = UCase("WDpSel")
      Cancel = NewValue <> "" And (Not moCn.RecSeek("select PMCd from Param where " + _
                     " PTyp= 'YN' and PMCd = '" + NewValue + "'"))
      If NewValue <> "Y" Then
        ms_DmCdStr = Replace(ms_DmCdStr, "," + GrdDsgPrmDet.Value(RowNum, "DpDmCd"), "")
        ErrMsg = "Enter Selection As Y to Selcting Design Or N/Blank for Not To Select": Exit Sub
      Else
        If InStr(1, ms_DmCdStr, GrdDsgPrmDet.Value(RowNum, "DpDmCd")) <= 0 Then _
            ms_DmCdStr = ms_DmCdStr + "," + GrdDsgPrmDet.Value(RowNum, "DpDmCd")
      End If
    '*** Jay 3.4[DP]
  End Select
End Sub
'*** Jay 3.4[DP]

'*** Jay 3.4[DP]
Private Sub GrdDsgPrmDet_ColWhen(ByVal RowNum As Integer, ByVal ColName As String, Cancel As Boolean, ErrMsg As String, ByVal Action As MwfCtl.en_fgAction)
  '*** Jay 3.4[DP]
  If UCase(ColName) <> "WDPSEL" Then
    Cancel = True
  Else
    GrdDsgPrmDet.Value(RowNum, "WDpSEL") = IIF(Trim(GrdDsgPrmDet.Value(RowNum, "WDpSEL")) = "", "Y", "")
  End If
  '*** Jay 3.4[DP]
End Sub
'*** Jay 3.4[DP]

Private Sub GrdDsgPrmDet_LostFocus()
'*** Jay 3.4[DP]
  Call ShowPic(wPicDPHt, wPicDPWdt, "", PicDP)
End Sub
'*** Jay 3.4[DP]
Private Sub GrdDsgRm_ColWhen(ByVal RowNum As Integer, ByVal ColName As String, Cancel As Boolean, ErrMsg As String, ByVal Action As MwfCtl.en_fgAction)
    Cancel = True
End Sub
Private Sub GrdDsgRm_SetRecSource()
  '*** Set the RecSource for GrdDsgRm
  If GrdDsgLst.Row > 0 Then _
  GrdDsgRm.RecSource = " Select * from DsgRm " + _
                       " Where DrTcTyp ='" + GrdDsgLst.Value(GrdDsgLst.Row, "DMTCTYP") + "'" + _
                       " And DrCd = '" + GrdDsgLst.Value(GrdDsgLst.Row, "DMCD") + "'" + _
                       " And DrSz ='" + GrdDsgLst.Value(GrdDsgLst.Row, "DMSZ") + "' " + IIF(gs_Partition = ctCurrPrtn, " And DrPrtKey='" + ctCurrPrtn + "' ", "")
End Sub

Private Sub GrdDsgLab_SetRecSource()
  '*** Set the RecSource for GrdDsgLab
  If GrdDsgLst.Row > 0 Then _
  GrdDsgLab.RecSource = " Select * from DsgLab " + _
                        " Where DLTcTyp ='" + GrdDsgLst.Value(GrdDsgLst.Row, "DMTCTYP") + "'" + _
                        " And DlCd = '" + GrdDsgLst.Value(GrdDsgLst.Row, "DMCD") + "'" + _
                        " And DLSz ='" + GrdDsgLst.Value(GrdDsgLst.Row, "DMSZ") + "' " + IIF(gs_Partition = ctCurrPrtn, " And DlPrtKey='" + ctCurrPrtn + "' ", "")
End Sub

Private Sub GrdDsgAna_SetRecSource()
  '*** Set the RecSource for GrdDsgAna
  If GrdDsgLst.Row > 0 Then _
  GrdDsgAna.RecSource = " Select * from DsgAna " + _
                        " Where DaTcTyp ='" + GrdDsgLst.Value(GrdDsgLst.Row, "DMTCTYP") + "'" + _
                        " And DaCd = '" + GrdDsgLst.Value(GrdDsgLst.Row, "DMCD") + "'" + _
                        " And DaSz ='" + GrdDsgLst.Value(GrdDsgLst.Row, "DMSZ") + "' " + IIF(gs_Partition = ctCurrPrtn, " And DaPrtKey='" + ctCurrPrtn + "' ", "") + " Order By DaAnaSr "
End Sub
'*** Jay 2.13(CT) ***
Private Sub GrdDsgCT_SetRecSource()
  '*** Set the RecSource for GrdDsgCT
  If GrdDsgLst.Row > 0 Then _
  GrdDsgCT.RecSource = " Select * from DsgCT " + _
                        " Where DctTcTyp='" + GrdDsgLst.Value(GrdDsgLst.Row, "DMTCTYP") + "'" + _
                        " And DctDmCd= '" + GrdDsgLst.Value(GrdDsgLst.Row, "DMCD") + "'" + _
                        " And DctDmSz='" + GrdDsgLst.Value(GrdDsgLst.Row, "DMSZ") + "'" + IIF(gs_Partition = ctCurrPrtn, " And DctPrtKey='" + ctCurrPrtn + "' ", "") + " Order By DctSr"
End Sub
Private Sub GrdDsgRm_Validate(Cancel As Boolean)
  Cancel = GrdDsgRm.Validate
End Sub
Private Sub GrdDsgPrm_Validate(Cancel As Boolean)
'*** Jay 3.4[DP]
  Cancel = GrdDsgPrm.Validate
End Sub
'*** Jay 3.4[DP]

Private Sub GrdDsgPrmDet_Validate(Cancel As Boolean)
'*** Jay 3.4[DP]
  Cancel = GrdDsgPrmDet.Validate
End Sub
'*** Jay 3.4[DP]

Private Sub OptDmHld_Click(Index As Integer)
' ****** Manali 3.5.0 - 05/05/09 - Radio Buttons Added
Select Case Index
  Case Is = 0
    ADC("wOnHldYN") = "Y"
  Case Is = 1
    ADC("wOnHldYN") = "N"
  Case Is = 2
    ADC("wOnHldYN") = ""
End Select
' ****** Manali 3.5.0 - 05/05/09 - Radio Buttons Added
End Sub
' ***** Manali 3.5.0 - 06/05/09- Radio Buttons Added
Private Sub OptDmHld_GotFocus(Index As Integer)
  DispMsg OptDmHld(Index).ToolTipText, etInfo
End Sub
' ***** Manali 3.5.0 - 06/05/09 - Radio Buttons Added

Private Sub OptDmValid_Click(Index As Integer)
'4.1.3.0
Select Case Index
  Case Is = 0
    ADC("wValidYN") = "Y"
  Case Is = 1
    ADC("wValidYN") = "N"
  Case Is = 2
    ADC("wValidYN") = ""
End Select
End Sub

Private Sub OptDsgExc_Click(Index As Integer)
' ****** Manali 3.5.0 - 05/05/09 - Radio Buttons Added
' ****** Manali 350Nxt - Valid values for wDmRegExc changed from 'YN' to 'DSGEXC'
'578-12 exclusive option disabled
'Select Case Index
'  Case Is = 0
'''    ADC("wDmRegExc") = "Y"
'      ADC("wDmRegExc") = "R"
'  Case Is = 1
'''    ADC("wDmRegExc") = "N"
'    ADC("wDmRegExc") = "C"
'  Case Is = 2
'    ADC("wDmRegExc") = "F"
'''    ADC("wDmRegExc") = ""
'  Case Is = 3
'    ADC("wDmRegExc") = "A"
'End Select
' ****** Manali 350Nxt - Valid values for wDmRegExc changed from 'YN' to 'DSGEXC'
' ****** Manali 3.5.0 - 05/05/09 - Radio Buttons Added
End Sub
' ***** Manali 3.5.0 - 06/05/09- Radio Buttons Added
Private Sub OptDsgExc_GotFocus(Index As Integer)
  DispMsg OptDsgExc(Index).ToolTipText, etInfo
End Sub
' ***** Manali 3.5.0 - 06/05/09 - Radio Buttons Added

' ****** Manali 3.5.0 - 05/05/09 - Radio Buttons Added
Private Sub OptRmPrd_Click(Index As Integer)
Select Case Index
  Case Is = 0
    ADC("wPrdRmQtyWtYN") = "Y"
  Case Is = 1
    ADC("wPrdRmQtyWtYN") = "N"
End Select
End Sub
' ****** Manali 3.5.0 - 05/05/09 - Radio Buttons Added

' ***** Manali 3.5.0 - 06/05/09- Radio Buttons Added
Private Sub OptRmPrd_GotFocus(Index As Integer)
  DispMsg OptRmPrd(Index).ToolTipText, etInfo
End Sub
' ***** Manali 3.5.0 - 06/05/09 - Radio Buttons Added

Private Sub TabDesgnDet_Click(PreviousTab As Integer)
  '*** (Jen) 08/06/07
  FraDesgnDet(TabDesgnDet.Tab).Enabled = True:  FraDesgnDet(PreviousTab).Enabled = False
  If ADC("wShowPrdRmDets") = "" Then ADC("wShowPrdRmDets") = "Y"
  If TabDesgnDet.Tab = 8 Then Call Summary(ADC("wShowPrdRmDets"))
  '*** (Jen) 08/06/07
End Sub
Private Sub DispFra(ByVal pv_DsgVwFra As en_DsgVwFra)
  '*** This is the common procedure which is used to toggle between hiding and displaying
      'the frames (i.e. if the frame is visible then make it invisible and viseversa)
      'associated with a command button
  Select Case pv_DsgVwFra
  Case Is = Summ
    '*** (Bef 08/06/07)
    'If FraSumm.Visible = True Then
    '  FraSumm.Visible = False
    '  FraSumm.Enabled = False
    '  Call EnaDisaCmds(False, CmdSumm)
    '  GrdDsgLst.SetFocus
    'Else
    '  FraSumm.Visible = True
    '  FraSumm.Enabled = True
    '  CmdSumm.SetFocus
    '  FraSumm.ZOrder
    '  Call EnaDisaCmds(True, CmdSumm)
    'End If
    '*** (Bef 08/06/07)
  '*** Jay 3.4[DP]
  Case Is = CatCpy
    If FraCpyCat.Visible = True Then
      FraCpyCat.Visible = False
      FraCpyCat.Enabled = False
      Call EnaDisaCmds(False, CmdCpyCat)
      CmdCpyCat.SetFocus
    Else
      FraCpyCat.Visible = True
      FraCpyCat.Enabled = True
      FraCpyCat.ZOrder
      ADC("wDpCd").SetFocus
      Call EnaDisaCmds(True, CmdCpyCat)
    End If
    '*** Jay 3.4[DP]
  End Select
End Sub
Private Sub EnaDisaCmds(ByVal pv_ShowFra As Boolean, Optional ByVal pv_CmdBut As MWCTL_BTN1)
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
      '*** Old Code
      'pv_CmdBut.BackColor = vbWhite
      'pv_CmdBut.FontBold = True
      '*** Old Code
      If pv_CmdBut.Pic = 0 Then
        pv_CmdBut.BackColor = vbWhite
        'pv_CmdBut.FontBold = True
      End If
      pv_CmdBut.CausesValidation = True
    Else
      '*** Old Code
      'pv_CmdBut.BackColor = &H8000000F
      'pv_CmdBut.FontBold = False
      '*** Old Code
      If pv_CmdBut.Pic = 0 Then
        pv_CmdBut.BackColor = &H8000000F
        'pv_CmdBut.FontBold = False
      End If
      'pv_CmdBut.CausesValidation = False
    End If
  End If
  FraNKeyAll.Enabled = Not pv_ShowFra
  'CmdSumm.Enabled = Not pv_ShowFra     '*** (Bef 08/06/07)
  CmdCpyCat.Enabled = Not pv_ShowFra     '*** Jay 3.4[DP]
  CmdDetScope.Enabled = Not pv_ShowFra
  If Not pv_CmdBut Is Nothing Then pv_CmdBut.Enabled = True

End Sub

Private Sub Summary(ByVal ms_ShowPrdRmDets As String)
  '*** Summary option is used to show the Raw Material Qty and Wt totals at the Rm Sub Ctg Level
  '*** It also shows the Gross and Category wise Totals
  '*** While calculating the Gross Wt, the Diamond & Colour Stone Wts have to be converted to Gms
  
  ' ******** Manali 3.03 - 16/07/08 - Pd Mod - Pd added
  Dim wGrsWt As Single, wGldWt As Single, wPlWt As Single, wSilWt As Single, wPdWt As Single, wOthMtWt As Single
  Dim wDiaWt As Single, wDiaQty As Single, wCSWt As Single, wCSQty As Single
  Dim wAccWt As Single, wAccQty As Single, wGrsQty As Single, wRow As Single
  Dim wWt As String, wQty As String
  
  If GrdDsgLst.Rows - 1 = 0 Or GrdDsgLst.Row = 0 Then Exit Sub '*** (Jen 08/06/07)
  
  wGrsWt = 0: wGldWt = 0: wPlWt = 0: wOthMtWt = 0
  wDiaWt = 0: wDiaQty = 0: wCSWt = 0: wCSQty = 0
  wAccWt = 0: wAccQty = 0: wSilWt = 0: wPdWt = 0        ' **** Manali 3.03 - 16/07/08 - Pd Mod - Pd added
    
    If ms_ShowPrdRmDets = "Y" Then
      wWt = "DrPrdWt": wQty = "DrPrdQty"
    ElseIf ms_ShowPrdRmDets = "N" Then
      wWt = "DrWt": wQty = "DrQty"
    End If
  '*** Open The ResultSet
  Dim wSummRs As MwfLib.MDORowSet
  Set wSummRs = moCn.OpenRes("Select DrRmCtg, DrRmSCtg, sum(" + wQty + ") as qQty, " + _
                             "sum(" + wWt + ") as qWt " + _
                             "from DsgRm where DrTcTyp='" + GrdDsgLst.Value(GrdDsgLst.Row, "DMTCTYP") + "' and " + _
                             "DrCd='" + GrdDsgLst.Value(GrdDsgLst.Row, "DMCD") + "' and " + _
                             "DrSz='" + GrdDsgLst.Value(GrdDsgLst.Row, "DMSZ") + "' " + IIF(gs_Partition = ctCurrPrtn, " And DrPrtKey='" + ctCurrPrtn + "' ", "") + _
                             "Group By DrRmCtg, DrRmSCtg")
  '*** Set the allow add and allow delete options of the grid to true
  GrdSumm.AllowAdd = True
  GrdSumm.AllowDelete = True
  
  '*** The DelAllRecs option is not to be used for deleting existing records in the grid instead set the dot rows property to 1
  GrdSumm.Rows = 1
  ''GrdSumm.DelAllRecs
  
  With GrdSumm
    Do While Not (wSummRs.EOF Or wSummRs.BOF)
      .AddItem
      wRow = .Rows - 1
      .Value(wRow, "wRmCtg") = wSummRs!DrRmCtg
      .Value(wRow, "wRmSCtg") = wSummRs!DrRmSCtg
      .Value(wRow, "wQty") = wSummRs!qQty
      .Value(wRow, "wWt") = wSummRs!qWt
      Select Case Trim(wSummRs!DrRmCtg)
      Case Is = "G"
        wGldWt = wGldWt + wSummRs!qWt
      Case Is = "P"
        wPlWt = wPlWt + wSummRs!qWt
      
      ' ****** sachin 2.12 ****** '
      Case Is = "S"
        wSilWt = wSilWt + wSummRs!qWt
      ' ****** sachin 2.12 ****** '
        
      ' ****** Manali 3.03 - 16/07/08 - Pd Mod
      Case Is = "L"
        wPdWt = wPdWt + wSummRs!qWt
      ' ****** Manali 3.03 - 16/07/08 - Pd Mod
    
      Case Is = "M"
        wOthMtWt = wOthMtWt + wSummRs!qWt
      Case Is = "D"
        'wDiaWt = wDiaWt + (wSummRs!qWt / 5)
        wDiaWt = wDiaWt + wSummRs!qWt
        wDiaQty = wDiaQty + wSummRs!qQty
      Case Is = "C"
        'wCSWt = wCSWt + (wSummRs!qWt / 5)
        wCSWt = wCSWt + wSummRs!qWt
        wCSQty = wCSQty + wSummRs!qQty
      Case Is = "X"
        wAccWt = wAccWt + wSummRs!qWt
        wAccQty = wAccQty + wSummRs!qQty
      End Select
      If Trim(wSummRs!DrRmCtg) = "C" Or Trim(wSummRs!DrRmCtg) = "D" Then
        wGrsWt = wGrsWt + (wSummRs!qWt / 5)
      Else
        wGrsWt = wGrsWt + wSummRs!qWt
      End If
      wGrsQty = wGrsQty + wSummRs!qQty
      ADC("wGrsWt") = wGrsWt: ADC("wGrsQty") = wGrsQty
      ADC("wGldWt") = wGldWt
      ADC("wPlWt") = wPlWt
      ADC("wSilWt") = wSilWt
      ADC("wPdWt") = wPdWt      ' ****** Manali 3.03 - 16/07/08 - Pd Mod
      ADC("wOthMtWt") = wOthMtWt
      ADC("wDiaWt") = wDiaWt: ADC("wDiaQty") = wDiaQty
      ADC("wCSWt") = wCSWt: ADC("wCSQty") = wCSQty
      ADC("wAccWt") = wAccWt: ADC("wAccQty") = wAccQty
      ADC("wKt") = ADC("WDmKt")
      wSummRs.MoveNext
    Loop
  End With
  
  GrdSumm.AllowAdd = False
  GrdSumm.AllowDelete = False
  Set wSummRs = Nothing
End Sub
Private Sub SetSelection()
  If Trim(ADC("wDmCtgSel")) <> "" Then ADC("wDmCtgSel") = "'" + Replace(ADC("wDmCtgSel"), ",", "','") + "'"
  If Trim(ADC("wPrdCtgSel")) <> "" Then ADC("wPrdCtgSel") = "'" + Replace(ADC("wPrdCtgSel"), ",", "','") + "'"
  If Trim(ADC("wSalCtgSel")) <> "" Then ADC("wSalCtgSel") = "'" + Replace(ADC("wSalCtgSel"), ",", "','") + "'"
  If Trim(ADC("wDmCdSel")) <> "" Then ADC("wDmCdSel") = "'" + Replace(ADC("wDmCdSel"), ",", "','") + "'"
  If Trim(ADC("wDmOldCdSel")) <> "" Then ADC("wDmOldCdSel") = "'" + Replace(ADC("wDmOldCdSel"), ",", "','") + "'"
  If Trim(ADC("wDmSzSel")) <> "" Then ADC("wDmSzSel") = "'" + Replace(ADC("wDmSzSel"), ",", "','") + "'"
  If Trim(ADC("wSetCdSel")) <> "" Then ADC("wSetCdSel") = "'" + Replace(ADC("wSetCdSel"), ",", "','") + "'"
  If Trim(ADC("wPrdSeqSel")) <> "" Then ADC("wPrdSeqSel") = "'" + Replace(ADC("wPrdSeqSel"), ",", "','") + "'"
  If Trim(ADC("wDmColSel")) <> "" Then ADC("wDmColSel") = "'" + Replace(ADC("wDmColSel"), ",", "','") + "'"
  If Trim(ADC("wDrSetSCdSel")) <> "" Then ADC("wDrSetSCdSel") = "'" + Replace(ADC("wDrSetSCdSel"), ",", "','") + "'"    '*** Jay 3.4[DP]
  
  ADC("wDmCtgSel").CmpStr = "DmCtg In "
  ADC("wPrdCtgSel").CmpStr = "DmPrdCtg In "
  ADC("wSalCtgSel").CmpStr = "DmSalCtg In "
  ADC("wDmCdSel").CmpStr = "DmCd In "
  ADC("wDmOldCdSel").CmpStr = "DmOldCd In "
  ADC("wDmSzSel").CmpStr = "DmSz In "
  ADC("wSetCdSel").CmpStr = "DmSetCd In "
  ADC("wPrdSeqSel").CmpStr = "DmPrdSeq In "
  ADC("wDmColSel").CmpStr = "DmCol In "
  ADC("wDrSetSCdSel").CmpStr = "DrSetSCd In "   '*** Jay 3.4[DP]
End Sub
Private Sub GrdDsgSz_ColChange(ByVal RowNum As Integer, ByVal ColName As String, ByVal NewValue As Variant, Cancel As Boolean, ErrMsg As String)
  With GrdDsgSz
    Select Case UCase(ColName)
    ' ***** Manali 3.6.0 - 31/10/09 - Detail Scope applied for Colour Stones
    Case Is = UCase("UodRmCtg")
      Cancel = Not (NewValue = "D" Or NewValue = "C")
      ErrMsg = "Rm Ctg 'D' and 'C' allowed": Exit Sub
    
    Case Is = UCase("UodRmSCtg")
      Cancel = NewValue <> "" And Not moCn.RecSeek("Select PMCd From Param Where PTyp= 'RMSCTG' and PMCd = '" + GrdDsgSz.Value(RowNum, "UodRmCtg") + "' and PSCd= '" + NewValue + "'")
'      Cancel = NewValue <> "" And Not moCn.RecSeek("Select PMCd From Param Where PTyp= 'RMSCTG' and PMCd = 'D' and PSCd= '" + NewValue + "'")
      ErrMsg = "Invalid Rm Sub Ctg": Exit Sub
    End Select
  End With
End Sub
Private Sub GrdDsgSz_ColWhen(ByVal RowNum As Integer, ByVal ColName As String, Cancel As Boolean, ErrMsg As String, ByVal Action As MwfCtl.en_fgAction)
  Select Case ColName
  ' ***** Manali 3.6.0 - 31/10/09 - Detail Scope applied for Colour Stones
  Case Is = UCase("UodRmCtg")
    Call HlpList.PMCd("RMCTG")
  
  Case Is = UCase("UodRmSCtg")
    Call HlpList.PSCd("RMSCTG", GrdDsgSz.Value(RowNum, "UodRmCtg"))
'    Call HlpList.PSCd("RMSCTG", "D")
  End Select
End Sub
Private Sub GrdDsgSz_ReCalc(ByVal RowNum As Integer, ByVal ColName As String, ByVal OldVal As Variant, ByVal SourceName As String, ByVal SourceColName As String)
  With GrdDsgSz
    Select Case UCase(ColName)
    ' ***** Manali 3.6.0 - 31/10/09 - Detail Scope applied for Colour Stones
    Case Is = UCase("UodRmSCtg")
      If .Mode = fgmnorm Then .Value(RowNum, "UodRmSCtg") = ""
    
    Case Is = UCase("UodPtrYN")
      If .Mode = fgmnorm Or .Mode = fgminit Then
        .Value(RowNum, "UodPtrYN") = IIF(moCn.GetFldVal("Select PValue3 from Param where PTyp= 'RMSCTG'  " + _
                                     "and PMCd= '" + GrdDsgSz.Value(RowNum, "UodRmCtg") + "' and PSCd= '" + .Value(RowNum, "UodRmSCtg") + "'") = "N", "N", "Y")

'      If .Mode = fgmnorm Or .Mode = fgminit Then
'        .Value(RowNum, "UodPtrYN") = IIF(moCn.GetFldVal("Select PValue3 from Param where PTyp= 'RMSCTG'  " + _
'                                     "and PMCd= 'D' and PSCd= '" + .Value(RowNum, "UodRmSCtg") + "'") = "N", "N", "Y")
      End If
    End Select
  End With
End Sub
Private Sub GrdDsgSz_InitKey(ByVal RowNum As Integer)
  '*** Initialise Key Fields
      'The fields of the Base table which are not existing in the grid are stored in the
      'table using Grid_InitKey
  
  With GrdDsgSz
    If .IsNew(RowNum) Then
      .Store "UodUsrCd", gs_UsrCd
      .Store "UodMnuCd", ADC.MenuCd
      .Store "UodTbl", "RmSz"
    End If
  End With
End Sub
Private Sub GrdDsgSz_SetRecSource()
  '*** Set the Record Source of the Grid GrdDsgAna
  GrdDsgSz.RecSource = "Select * from UsrOptDet where UodUsrCd= '" + gs_UsrCd + "' and " + _
                       "UodMnuCd= '" + ADC.MenuCd + "' and UodTbl= 'RmSz' Order By UodRmCtg, UodRmSCtg, UodLnSvPtrFr"
End Sub
Private Sub GrdDsgSz_Validate(Cancel As Boolean)
  Cancel = GrdDsgSz.Validate
End Sub


Private Sub Pic3D_DblClick()
'****** Manali 3.2.2 Enlarge Pic
  '*** If wPic3DDblClk flag option is False then Enlarge the 3-Dimensional picture to fit the FraNKeyAll Frame
  '*** If wPic3DDblClk flag option is True then bring the 3-Dimensional picture to the original pic control size

    If GrdDsgLst.Rows - 1 = 0 Then Exit Sub
    ADC.SetFocus
    Dim w3DPth As String
    If wPic3DDblClk = False Then
        '*** Enlarge The Picture ***
        'sv.75 "DM" changed as ADC("wDmTyp")
        w3DPth = GetPictPath(ADC("wDmTyp"), False, GrdDsgLst.Value(GrdDsgLst.Row, "DMCD"), "")
        Call ShowLargePic(IIF(FraNKeyAll.Height - 195 > 0, FraNKeyAll.Height - 195, 0), FraNKeyAll.Width, w3DPth, Pic3D)

        FraPic3D.Left = 0: FraPic3D.Top = 195
        FraPic3D.Height = Pic3D.Height + 100: FraPic3D.Width = Pic3D.Width + 100
    FraNKeyAll.ZOrder
    FraPic3D.ZOrder (0)
    Else
        '*** Make The Picture Small ***
        'sv.75 "DM" changed as ADC("wDmTyp")
        w3DPth = GetPictPath(ADC("wDmTyp"), False, GrdDsgLst.Value(GrdDsgLst.Row, "DMCD"), "")
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

    If GrdDsgLst.Rows - 1 = 0 Then Exit Sub
    ADC.SetFocus
    Dim wHskPth As String
    If wPicHskDblClk = False Then
        '*** Enlarge The Picture ***
        'sv.75 "DM" changed as ADC("wDmTyp")
        wHskPth = GetPictPath(ADC("wDmTyp"), True, GrdDsgLst.Value(GrdDsgLst.Row, "DMCD"), "")
        Call ShowLargePic(IIF(FraNKeyAll.Height - 195 > 0, FraNKeyAll.Height - 195, 0), FraNKeyAll.Width, wHskPth, PicHsk)
        FraPicHsk.Left = 0: FraPicHsk.Top = 195
        FraPicHsk.Height = PicHsk.Height + 100: FraPicHsk.Width = PicHsk.Width + 100
        FraNKeyAll.ZOrder
        FraPicHsk.ZOrder (0)
    Else
        '*** Make The Picture Small ***
        'sv.75 "DM" changed as ADC("wDmTyp")
        wHskPth = GetPictPath(ADC("wDmTyp"), True, GrdDsgLst.Value(GrdDsgLst.Row, "DMCD"), "")
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
'*** Jay 3.4[DP]
Private Sub HideAllFras()
  '*** This procedure hides all the frames other than FraNKeyAll
  FraNKeyAll.Visible = True
  FraCpyCat.Visible = False
  FraCpyCat.Enabled = False
End Sub
'*** Jay 3.4[DP]

'*** Jay 3.4[DP]
Public Sub CpyCat()
Dim i As Integer, wi_MaxDpDmNo As Integer, wDmCd As String, wi_MaxDpNo As Integer, ws_DmCdStr As String
  With GrdDsgLst
    If .Rows - 1 < 0 Then DispMsg "Design List Grid Empty Can Not Proceed", etError: Exit Sub
    DispMsg "Adding Designs To Catalogue :" + ADC("wDpCd"), etInfo
    
    wi_MaxDpDmNo = moCn.GetFldVal("Select Max(DpDmNo) from DsgPrm where DpTyp = 'CAT' and DpCd= '" + ADC("wDpCd") + "' ")
    wi_MaxDpNo = moCn.GetFldVal("Select Max(DpNo) from DsgPrm where DpTyp = 'CAT' and DpCd= '" + ADC("wDpCd") + "' ")
    
    On Error GoTo ErrH
    
    For i = 1 To GrdDsgLst.Rows - 1
      If .Value(i, "wSel") = "Y" And InStr(1, ws_DmCdStr, "," + .Value(i, "DmCd")) <= 0 Then
        ws_DmCdStr = ws_DmCdStr + .Value(i, "DmCd")
      End If
    Next
    If ms_DmCdStr = "" Then DispMsg "No Designs To Add Into Catalogue :" + ADC("wDpCd"), etWarning: Exit Sub
    
    ws_DmCdStr = "'" + Replace(Mid(ms_DmCdStr, 2), ",", "','") + "'"
    moCn.Execute (" Insert Into DsgPrm " + _
                  " (DpTyp,DpCd,DpDmCd,DpNo,DpDesc,DpDt,DpDmNo,DpPctNm,DpRem1,DpRem2,ModUsr,ModDt,ModTime) " + _
                  " " + _
                  " Select 'CAT','" + ADC("wDpCd") + "',DmCd," + CStr(wi_MaxDpNo) + " + ROW_NUMBER() OVER(ORDER BY DmCd ASC), " + _
                  " '','" + "01/01/80" + "'," + CStr(wi_MaxDpDmNo) + " + ROW_NUMBER() OVER(ORDER BY DmCd ASC), " + _
                  " '','','','" + UCase(gs_UsrCd) + "',GetDate(), cast(DATEPART(hh, GetDate())as varchar)+'.' +cast(DATEPART(mi, GetDate())as varchar) " + _
                  " from DsgMst where DmTcTyp='DM' and DmSz='' and DmCd in (" + ws_DmCdStr + ") and DmHld= 'N' and DmValidYn='Y'")      '4.1.3.0
    
    '---------
    'For i = 1 To GrdDsgLst.Rows - 1
    '  If .Value(i, "wSel") = "Y" Then
    '    '.Value(i, "wSel") = ""
    '    moCn.Execute ("Insert Into DsgPrm " + _
    '                "(DpTyp,DpCd,DpDmCd,DpNo,DpDesc,DpDt,DpDmNo,DpPctNm,DpRem1,DpRem2,ModUsr,ModDt,ModTime) " + _
    '                "Values " + _
    '                "('CAT','" + ADC("wDpCd") + "','" + .Value(i, "DmCd") + "'," + CStr(wi_MaxDpNo) + ",'','" + "01/01/80" + "'," + _
    '                CStr(wi_MaxDpDmNo) + ",'','','','" + UCase(gs_UsrCd) + "',GetDate(), cast(DATEPART(hh, GetDate())as varchar)+'.' +cast(DATEPART(mi, GetDate())as varchar) )")
    '    wi_MaxDpDmNo = wi_MaxDpDmNo + 1
    '    wi_MaxDpNo = wi_MaxDpNo + 1
    '  End If
    'Next
    'For i = 1 To GrdDsgPrmDet.Rows - 1
    '  If GrdDsgPrmDet.Value(i, "wDpSel") = "Y" Then
    '    'GrdDsgPrmDet.Value(i, "wDpSel") = ""
    '    moCn.Execute ("Insert Into DsgPrm " + _
    '                "(DpTyp,DpCd,DpDmCd,DpNo,DpDesc,DpDt,DpDmNo,DpPctNm,DpRem1,DpRem2,ModUsr,ModDt,ModTime) " + _
    '                "Values " + _
    '                "('CAT','" + ADC("wDpCd") + "','" + GrdDsgPrmDet.Value(i, "DpDmCd") + "'," + CStr(wi_MaxDpNo) + ",'','" + "01/01/80" + "'," + _
    '                CStr(wi_MaxDpDmNo) + ",'','','','" + UCase(gs_UsrCd) + "',GetDate(), cast(DATEPART(hh, GetDate())as varchar)+'.' +cast(DATEPART(mi, GetDate())as varchar) )")
    '    wi_MaxDpDmNo = wi_MaxDpDmNo + 1
    '    wi_MaxDpNo = wi_MaxDpNo + 1
    '  End If
    'Next
    '------------
    If wDmCd = "" Then
      DispMsg "Selected Designs Are Added Into Catalogue :" + ADC("wDpCd"), etInfo
    Else
      'DispMsg "Following Designs Are Not Added Into Catalogue :" + wDmCd, etInfo
    End If
    Exit Sub
  End With
ErrH:
  'wDmCd = wDmCd + IIF(wDmCd <> "", ",", "") + .Value(i, "DmCd")
  'Resume Next
  MsgBox Err.Description
End Sub
'*** Jay 3.4[DP]


