VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{AB2CC32A-E652-42A2-B341-26ECE02F6386}#2.1#0"; "MwfCtlForSpeed610.ocx"
Begin VB.Form EmrFrmInvHd 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Invoice Header"
   ClientHeight    =   9810
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14880
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9810
   ScaleWidth      =   14880
   WindowState     =   2  'Maximized
   Begin VB.Frame FraAddAmt 
      BorderStyle     =   0  'None
      Height          =   5295
      Left            =   120
      TabIndex        =   460
      Top             =   3800
      Width           =   14535
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   271
         Left            =   1515
         TabIndex        =   463
         ToolTipText     =   "Enter Invoice Amount Description 1"
         Top             =   690
         Width           =   6660
         _ExtentX        =   11748
         _ExtentY        =   503
         MaxLength       =   45
         DataField       =   "InAmtDesc1"
         IdName          =   "INAMTDESC1"
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   272
         Left            =   1515
         TabIndex        =   466
         ToolTipText     =   "Enter Invoice Amount Description 2"
         Top             =   975
         Width           =   6660
         _ExtentX        =   11748
         _ExtentY        =   503
         MaxLength       =   45
         DataField       =   "InAmtDesc2"
         IdName          =   "INAMTDESC2"
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   273
         Left            =   1515
         TabIndex        =   469
         ToolTipText     =   "Enter Invoice Amount Description 3"
         Top             =   1260
         Width           =   6660
         _ExtentX        =   11748
         _ExtentY        =   503
         MaxLength       =   45
         DataField       =   "InAmtDesc3"
         IdName          =   "INAMTDESC3"
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   274
         Left            =   1515
         TabIndex        =   472
         ToolTipText     =   "Enter Invoice Amount Description 4"
         Top             =   1545
         Width           =   6660
         _ExtentX        =   11748
         _ExtentY        =   503
         MaxLength       =   45
         DataField       =   "InAmtDesc4"
         IdName          =   "INAMTDESC4"
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   275
         Left            =   1515
         TabIndex        =   475
         ToolTipText     =   "Enter Invoice Amount Description 5"
         Top             =   1830
         Width           =   6660
         _ExtentX        =   11748
         _ExtentY        =   503
         MaxLength       =   45
         DataField       =   "InAmtDesc5"
         IdName          =   "INAMTDESC5"
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   276
         Left            =   1515
         TabIndex        =   478
         ToolTipText     =   "Enter Invoice Amount Description 6"
         Top             =   2115
         Width           =   6660
         _ExtentX        =   11748
         _ExtentY        =   503
         MaxLength       =   45
         DataField       =   "InAmtDesc6"
         IdName          =   "INAMTDESC6"
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   277
         Left            =   1515
         TabIndex        =   481
         ToolTipText     =   "Enter Invoice Amount Description 7"
         Top             =   2400
         Width           =   6660
         _ExtentX        =   11748
         _ExtentY        =   503
         MaxLength       =   45
         DataField       =   "InAmtDesc7"
         IdName          =   "INAMTDESC7"
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   278
         Left            =   1515
         TabIndex        =   484
         ToolTipText     =   "Enter Invoice Amount Description 8"
         Top             =   2685
         Width           =   6660
         _ExtentX        =   11748
         _ExtentY        =   503
         MaxLength       =   45
         DataField       =   "InAmtDesc8"
         IdName          =   "INAMTDESC8"
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   279
         Left            =   1515
         TabIndex        =   487
         ToolTipText     =   "Enter Invoice Amount Description 9"
         Top             =   2970
         Width           =   6660
         _ExtentX        =   11748
         _ExtentY        =   503
         MaxLength       =   45
         DataField       =   "InAmtDesc9"
         IdName          =   "INAMTDESC9"
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   280
         Left            =   1515
         TabIndex        =   490
         ToolTipText     =   "Enter Invoice Amount Description 10"
         Top             =   3255
         Width           =   6660
         _ExtentX        =   11748
         _ExtentY        =   503
         MaxLength       =   45
         DataField       =   "InAmtDesc10"
         IdName          =   "INAMTDESC10"
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   281
         Left            =   1515
         TabIndex        =   493
         ToolTipText     =   "Enter Invoice Amount Description 11"
         Top             =   3540
         Width           =   6660
         _ExtentX        =   11748
         _ExtentY        =   503
         MaxLength       =   45
         DataField       =   "InAmtDesc11"
         IdName          =   "INAMTDESC11"
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   282
         Left            =   1515
         TabIndex        =   496
         ToolTipText     =   "Enter Invoice Amount Description 12"
         Top             =   3825
         Width           =   6660
         _ExtentX        =   11748
         _ExtentY        =   503
         MaxLength       =   45
         DataField       =   "InAmtDesc12"
         IdName          =   "INAMTDESC12"
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   283
         Left            =   10320
         TabIndex        =   464
         ToolTipText     =   "Enter Invoice Amount 1 %"
         Top             =   720
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   503
         Alignment       =   2
         Mask            =   "-##0.000"
         MaxLength       =   8
         DataType        =   2
         DataField       =   "InAmt1Per"
         IdName          =   "INAMT1PER"
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   284
         Left            =   10320
         TabIndex        =   467
         ToolTipText     =   "Enter Invoice Amount 2 %"
         Top             =   1005
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   503
         Alignment       =   2
         Mask            =   "-##0.000"
         MaxLength       =   8
         DataType        =   2
         DataField       =   "InAmt2Per"
         IdName          =   "INAMT2PER"
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   285
         Left            =   10320
         TabIndex        =   470
         ToolTipText     =   "Enter Invoice Amount 3 %"
         Top             =   1290
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   503
         Alignment       =   2
         Mask            =   "-##0.000"
         MaxLength       =   8
         DataType        =   2
         DataField       =   "InAmt3Per"
         IdName          =   "INAMT3PER"
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   286
         Left            =   10320
         TabIndex        =   473
         ToolTipText     =   "Enter Invoice Amount 4 %"
         Top             =   1575
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   503
         Alignment       =   2
         Mask            =   "-##0.000"
         MaxLength       =   8
         DataType        =   2
         DataField       =   "InAmt4Per"
         IdName          =   "INAMT4PER"
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   287
         Left            =   10320
         TabIndex        =   476
         ToolTipText     =   "Enter Invoice Amount 5 %"
         Top             =   1860
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   503
         Alignment       =   2
         Mask            =   "-##0.000"
         MaxLength       =   8
         DataType        =   2
         DataField       =   "InAmt5Per"
         IdName          =   "INAMT5PER"
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   288
         Left            =   10320
         TabIndex        =   479
         ToolTipText     =   "Enter Invoice Amount 6 %"
         Top             =   2145
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   503
         Alignment       =   2
         Mask            =   "-##0.000"
         MaxLength       =   8
         DataType        =   2
         DataField       =   "InAmt6Per"
         IdName          =   "INAMT6PER"
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   289
         Left            =   10320
         TabIndex        =   485
         ToolTipText     =   "Enter Invoice Amount 8 %"
         Top             =   2715
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   503
         Alignment       =   2
         Mask            =   "-##0.000"
         MaxLength       =   8
         DataType        =   2
         DataField       =   "InAmt8Per"
         IdName          =   "INAMT8PER"
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   290
         Left            =   10320
         TabIndex        =   488
         ToolTipText     =   "Enter Invoice Amount 9 %"
         Top             =   3000
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   503
         Alignment       =   2
         Mask            =   "-##0.000"
         MaxLength       =   8
         DataType        =   2
         DataField       =   "InAmt9Per"
         IdName          =   "INAMT9PER"
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   291
         Left            =   10320
         TabIndex        =   491
         ToolTipText     =   "Enter Invoice Amount 10 %"
         Top             =   3285
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   503
         Alignment       =   2
         Mask            =   "-##0.000"
         MaxLength       =   8
         DataType        =   2
         DataField       =   "InAmt10Per"
         IdName          =   "INAMT10PER"
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   292
         Left            =   10320
         TabIndex        =   494
         ToolTipText     =   "Enter Invoice Amount 11 %"
         Top             =   3570
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   503
         Alignment       =   2
         Mask            =   "-##0.000"
         MaxLength       =   8
         DataType        =   2
         DataField       =   "InAmt11Per"
         IdName          =   "INAMT11PER"
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   213
         Left            =   12840
         TabIndex        =   480
         ToolTipText     =   "Enter Invoice Amount 6 "
         Top             =   2145
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   503
         Alignment       =   2
         Mask            =   "-########0.00"
         MaxLength       =   13
         DataType        =   2
         DataField       =   "InAmt6"
         IdName          =   "INAMT6"
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   214
         Left            =   12840
         TabIndex        =   465
         ToolTipText     =   "Enter Invoice Amount 1"
         Top             =   720
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   503
         Alignment       =   2
         Mask            =   "-########0.00"
         MaxLength       =   13
         DataType        =   2
         DataField       =   "InAmt1"
         IdName          =   "INAMT1"
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   215
         Left            =   12840
         TabIndex        =   468
         ToolTipText     =   "Enter Invoice Amount 2"
         Top             =   1005
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   503
         Alignment       =   2
         Mask            =   "-########0.00"
         MaxLength       =   13
         DataType        =   2
         DataField       =   "InAmt2"
         IdName          =   "INAMT2"
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   216
         Left            =   12840
         TabIndex        =   471
         ToolTipText     =   "Enter Invoice Amount 3"
         Top             =   1290
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   503
         Alignment       =   2
         Mask            =   "-########0.00"
         MaxLength       =   13
         DataType        =   2
         DataField       =   "InAmt3"
         IdName          =   "INAMT3"
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   217
         Left            =   12840
         TabIndex        =   474
         ToolTipText     =   "Enter Invoice Amount 4"
         Top             =   1575
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   503
         Alignment       =   2
         Mask            =   "-########0.00"
         MaxLength       =   13
         DataType        =   2
         DataField       =   "InAmt4"
         IdName          =   "INAMT4"
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   218
         Left            =   12840
         TabIndex        =   477
         ToolTipText     =   "Enter Invoice Amount 5"
         Top             =   1860
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   503
         Alignment       =   2
         Mask            =   "-########0.00"
         MaxLength       =   13
         DataType        =   2
         DataField       =   "InAmt5"
         IdName          =   "INAMT5"
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   219
         Left            =   12840
         TabIndex        =   498
         ToolTipText     =   "Enter Invoice Amount 12"
         Top             =   3855
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   503
         Alignment       =   2
         Enabled         =   0   'False
         Mask            =   "-#########0.00"
         MaxLength       =   14
         DataType        =   2
         DataField       =   "InAmt12"
         IdName          =   "INAMT12"
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   220
         Left            =   12840
         TabIndex        =   483
         ToolTipText     =   "Enter Invoice Amount 7"
         Top             =   2430
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   503
         Alignment       =   2
         Enabled         =   0   'False
         Mask            =   "-#########0.00"
         MaxLength       =   14
         DataType        =   2
         DataField       =   "InAmt7"
         IdName          =   "INAMT7"
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   221
         Left            =   12840
         TabIndex        =   486
         ToolTipText     =   "Enter Invoice Amount 8"
         Top             =   2715
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   503
         Alignment       =   2
         Mask            =   "-########0.00"
         MaxLength       =   13
         DataType        =   2
         DataField       =   "InAmt8"
         IdName          =   "INAMT8"
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   222
         Left            =   12840
         TabIndex        =   489
         ToolTipText     =   "Enter Invoice Amount 9"
         Top             =   3000
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   503
         Alignment       =   2
         Mask            =   "-########0.00"
         MaxLength       =   13
         DataType        =   2
         DataField       =   "InAmt9"
         IdName          =   "INAMT9"
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   223
         Left            =   12840
         TabIndex        =   492
         ToolTipText     =   "Enter Invoice Amount 10"
         Top             =   3285
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   503
         Alignment       =   2
         Mask            =   "-########0.00"
         MaxLength       =   13
         DataType        =   2
         DataField       =   "InAmt10"
         IdName          =   "INAMT10"
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   224
         Left            =   12840
         TabIndex        =   495
         ToolTipText     =   "Enter Invoice Amount 11 "
         Top             =   3570
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   503
         Alignment       =   2
         Mask            =   "-########0.00"
         MaxLength       =   13
         DataType        =   2
         DataField       =   "InAmt11"
         IdName          =   "INAMT11"
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   225
         Left            =   1515
         TabIndex        =   499
         ToolTipText     =   "Enter Invoice Payment  Description 1"
         Top             =   4440
         Width           =   6660
         _ExtentX        =   11748
         _ExtentY        =   503
         MaxLength       =   45
         DataField       =   "InPayDesc1"
         IdName          =   "INPAYDESC1"
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   227
         Left            =   1515
         TabIndex        =   502
         ToolTipText     =   "Enter Invoice Payment  Description 2"
         Top             =   4725
         Width           =   6660
         _ExtentX        =   11748
         _ExtentY        =   503
         MaxLength       =   45
         DataField       =   "InPayDesc2"
         IdName          =   "INPAYDESC2"
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   232
         Left            =   1515
         TabIndex        =   505
         ToolTipText     =   "Enter Invoice Payment Description 3"
         Top             =   5010
         Width           =   6660
         _ExtentX        =   11748
         _ExtentY        =   503
         MaxLength       =   45
         DataField       =   "InPayDesc3"
         IdName          =   "INPAYDESC3"
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   233
         Left            =   10320
         TabIndex        =   500
         ToolTipText     =   "Enter Invoice Payment Days 1"
         Top             =   4440
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   503
         Alignment       =   2
         Mask            =   "##0"
         MaxLength       =   3
         DataType        =   1
         DataField       =   "InPayDays1"
         IdName          =   "INPAYDAYS1"
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   234
         Left            =   10320
         TabIndex        =   503
         ToolTipText     =   "Enter Invoice Payment Days 2"
         Top             =   4725
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   503
         Alignment       =   2
         Mask            =   "##0"
         MaxLength       =   3
         DataType        =   1
         DataField       =   "InPayDays2"
         IdName          =   "INPAYDAYS2"
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   235
         Left            =   10320
         TabIndex        =   506
         ToolTipText     =   "Enter Invoice Payment Days 3"
         Top             =   5010
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   503
         Alignment       =   2
         Mask            =   "##0"
         MaxLength       =   3
         DataType        =   1
         DataField       =   "InPayDays3"
         IdName          =   "INPAYDAYS3"
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   236
         Left            =   12840
         TabIndex        =   501
         ToolTipText     =   "Enter Invoice Payment Amount 1"
         Top             =   4440
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   503
         Alignment       =   2
         Mask            =   "-########0.00"
         MaxLength       =   13
         DataType        =   2
         DataField       =   "InPayAmt1"
         ReCalcParent    =   "INAMT7"
         IdName          =   "INPAYAMT1"
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   237
         Left            =   12840
         TabIndex        =   504
         ToolTipText     =   "Enter Invoice Payment Amount 2"
         Top             =   4725
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   503
         Alignment       =   2
         Mask            =   "-########0.00"
         MaxLength       =   13
         DataType        =   2
         DataField       =   "InPayAmt2"
         ReCalcParent    =   "INAMT7"
         IdName          =   "INPAYAMT2"
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   238
         Left            =   12840
         TabIndex        =   507
         ToolTipText     =   "Enter Invoice Payment Amount 3"
         Top             =   5010
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   503
         Alignment       =   2
         Mask            =   "-########0.00"
         MaxLength       =   13
         DataType        =   2
         DataField       =   "InPayAmt3"
         ReCalcParent    =   "INAMT7"
         IdName          =   "INPAYAMT3"
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   240
         Left            =   1515
         TabIndex        =   461
         ToolTipText     =   "Enter Gst Code"
         Top             =   360
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   503
         MaxLength       =   8
         DataType        =   4
         DataField       =   "InGstCd"
         IdName          =   "INGSTCD"
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   239
         Left            =   6795
         TabIndex        =   462
         ToolTipText     =   "Enter Base Amount"
         Top             =   360
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   503
         Alignment       =   2
         Mask            =   "########0.00"
         MaxLength       =   12
         DataType        =   2
         DataField       =   "InBaseAmt"
         ReCalcParent    =   "INFOBRS,INCOMMRS,INFRTRS,ININSRS,INDSCRS,INDEDUCTRS"
         IdName          =   "INBASEAMT"
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   241
         Left            =   8280
         TabIndex        =   558
         ToolTipText     =   "Enter GstOn Value"
         Top             =   690
         Visible         =   0   'False
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   503
         Alignment       =   2
         MaxLength       =   30
         DataType        =   4
         DataField       =   "InGstOn1"
         ReCalcParent    =   "INAMT1"
         IdName          =   "INGSTON1"
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   242
         Left            =   8280
         TabIndex        =   559
         ToolTipText     =   "Enter GstOn Value"
         Top             =   960
         Visible         =   0   'False
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   503
         Alignment       =   2
         MaxLength       =   30
         DataType        =   4
         DataField       =   "InGstOn2"
         ReCalcParent    =   "INAMT1"
         IdName          =   "INGSTON2"
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   243
         Left            =   8280
         TabIndex        =   560
         ToolTipText     =   "Enter GstOn Value"
         Top             =   1260
         Visible         =   0   'False
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   503
         Alignment       =   2
         MaxLength       =   30
         DataType        =   4
         DataField       =   "InGstOn3"
         ReCalcParent    =   "INAMT1"
         IdName          =   "INGSTON3"
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   244
         Left            =   8280
         TabIndex        =   561
         ToolTipText     =   "Enter GstOn Value"
         Top             =   1560
         Visible         =   0   'False
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   503
         Alignment       =   2
         MaxLength       =   30
         DataType        =   4
         DataField       =   "InGstOn4"
         ReCalcParent    =   "INAMT1"
         IdName          =   "INGSTON4"
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   245
         Left            =   8280
         TabIndex        =   562
         ToolTipText     =   "Enter GstOn Value"
         Top             =   1830
         Visible         =   0   'False
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   503
         Alignment       =   2
         MaxLength       =   30
         DataType        =   4
         DataField       =   "InGstOn5"
         ReCalcParent    =   "INAMT1"
         IdName          =   "INGSTON5"
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   246
         Left            =   8280
         TabIndex        =   563
         ToolTipText     =   "Enter GstOn Value"
         Top             =   2115
         Visible         =   0   'False
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   503
         Alignment       =   2
         MaxLength       =   30
         DataType        =   4
         DataField       =   "InGstOn6"
         ReCalcParent    =   "INAMT1"
         IdName          =   "INGSTON6"
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   247
         Left            =   8280
         TabIndex        =   564
         ToolTipText     =   "Enter GstOn Value"
         Top             =   2685
         Visible         =   0   'False
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   503
         Alignment       =   2
         MaxLength       =   30
         DataType        =   4
         DataField       =   "InGstOn8"
         ReCalcParent    =   "INAMT1"
         IdName          =   "INGSTON8"
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   248
         Left            =   8280
         TabIndex        =   565
         ToolTipText     =   "Enter GstOn Value"
         Top             =   2970
         Visible         =   0   'False
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   503
         Alignment       =   2
         MaxLength       =   30
         DataType        =   4
         DataField       =   "InGstOn9"
         ReCalcParent    =   "INAMT1"
         IdName          =   "INGSTON9"
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   249
         Left            =   8280
         TabIndex        =   566
         ToolTipText     =   "Enter GstOn Value"
         Top             =   3255
         Visible         =   0   'False
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   503
         Alignment       =   2
         MaxLength       =   30
         DataType        =   4
         DataField       =   "InGstOn10"
         ReCalcParent    =   "INAMT1"
         IdName          =   "INGSTON10"
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   250
         Left            =   8280
         TabIndex        =   567
         ToolTipText     =   "Enter GstOn Value"
         Top             =   3540
         Visible         =   0   'False
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   503
         Alignment       =   2
         MaxLength       =   30
         DataType        =   4
         DataField       =   "InGstOn11"
         ReCalcParent    =   "INAMT1"
         IdName          =   "INGSTON11"
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   261
         Left            =   8280
         TabIndex        =   578
         ToolTipText     =   "Enter GstOn Value"
         Top             =   2400
         Visible         =   0   'False
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   503
         Alignment       =   2
         MaxLength       =   30
         DataType        =   4
         DataField       =   "InGstOn7"
         ReCalcParent    =   "INAMT1"
         IdName          =   "INGSTON7"
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   262
         Left            =   8280
         TabIndex        =   579
         ToolTipText     =   "Enter GstOn Value"
         Top             =   3840
         Visible         =   0   'False
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   503
         Alignment       =   2
         MaxLength       =   30
         DataType        =   4
         DataField       =   "InGstOn12"
         ReCalcParent    =   "INAMT1"
         IdName          =   "INGSTON12"
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   263
         Left            =   10320
         TabIndex        =   482
         ToolTipText     =   "Enter Invoice Amount 7 %"
         Top             =   2430
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   503
         Alignment       =   2
         Mask            =   "-##0.000"
         MaxLength       =   8
         DataType        =   2
         DataField       =   "InAmt7Per"
         IdName          =   "INAMT7PER"
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   264
         Left            =   10320
         TabIndex        =   497
         ToolTipText     =   "Enter Invoice Amount 12 %"
         Top             =   3840
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   503
         Alignment       =   2
         Mask            =   "-##0.000"
         MaxLength       =   8
         DataType        =   2
         DataField       =   "InAmt12Per"
         IdName          =   "INAMT12PER"
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Inv Amt12 %"
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
         Index           =   200
         Left            =   8760
         TabIndex        =   581
         Top             =   3840
         Width           =   1395
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Inv Amt7 %"
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
         Index           =   199
         Left            =   8760
         TabIndex        =   580
         Top             =   2430
         Width           =   1395
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Base Amount"
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
         Index           =   192
         Left            =   5400
         TabIndex        =   557
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Gst Code"
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
         Index           =   191
         Left            =   0
         TabIndex        =   556
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Pay Amt 3"
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
         Index           =   190
         Left            =   11520
         TabIndex        =   553
         Top             =   5010
         Width           =   1305
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Pay Amt 2"
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
         Index           =   189
         Left            =   11520
         TabIndex        =   552
         Top             =   4725
         Width           =   1305
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Pay Amt 1"
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
         Index           =   188
         Left            =   11520
         TabIndex        =   551
         Top             =   4440
         Width           =   1305
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Pay Days 3"
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
         Index           =   187
         Left            =   8655
         TabIndex        =   550
         Top             =   5010
         Width           =   1500
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Pay Days 2"
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
         Index           =   185
         Left            =   8655
         TabIndex        =   549
         Top             =   4725
         Width           =   1500
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Pay Days 1"
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
         Index           =   184
         Left            =   8655
         TabIndex        =   548
         Top             =   4440
         Width           =   1500
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Pay Desc 3"
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
         Index           =   183
         Left            =   0
         TabIndex        =   547
         Top             =   5010
         Width           =   1545
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Pay Desc 2"
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
         Index           =   182
         Left            =   0
         TabIndex        =   546
         Top             =   4725
         Width           =   1545
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Pay Desc 1"
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
         Index           =   181
         Left            =   0
         TabIndex        =   545
         Top             =   4440
         Width           =   1545
      End
      Begin VB.Label LblMain 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "  Invoice Payment Details"
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
         Index           =   14
         Left            =   0
         TabIndex        =   544
         Top             =   4155
         Width           =   14445
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Inv Amt1"
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
         Index           =   180
         Left            =   11520
         TabIndex        =   543
         Top             =   720
         Width           =   1305
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Inv Amt2"
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
         Index           =   179
         Left            =   11520
         TabIndex        =   542
         Top             =   1005
         Width           =   1305
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Inv Amt3"
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
         Index           =   178
         Left            =   11520
         TabIndex        =   541
         Top             =   1290
         Width           =   1305
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Inv Amt4"
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
         Index           =   177
         Left            =   11520
         TabIndex        =   540
         Top             =   1575
         Width           =   1305
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Inv Amt6"
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
         Index           =   176
         Left            =   11520
         TabIndex        =   539
         Top             =   2145
         Width           =   1305
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Inv Amt8"
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
         Left            =   11520
         TabIndex        =   538
         Top             =   2715
         Width           =   1305
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Inv Amt5"
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
         Index           =   174
         Left            =   11520
         TabIndex        =   537
         Top             =   1860
         Width           =   1305
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Inv Amt7"
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
         Index           =   173
         Left            =   11520
         TabIndex        =   536
         Top             =   2430
         Width           =   1305
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Inv Amt9"
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
         Index           =   172
         Left            =   11520
         TabIndex        =   535
         Top             =   3000
         Width           =   1305
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Inv Amt10"
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
         Index           =   171
         Left            =   11520
         TabIndex        =   534
         Top             =   3285
         Width           =   1305
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Inv Amt11"
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
         Index           =   170
         Left            =   11520
         TabIndex        =   533
         Top             =   3570
         Width           =   1305
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Inv Amt12"
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
         Index           =   165
         Left            =   11520
         TabIndex        =   532
         Top             =   3855
         Width           =   1305
      End
      Begin VB.Label LblMain 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "  Invoice Additional Amounts"
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
         Index           =   13
         Left            =   0
         TabIndex        =   531
         Top             =   30
         Width           =   14565
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Inv Amt1 Desc"
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
         Index           =   204
         Left            =   0
         TabIndex        =   530
         Top             =   720
         Width           =   1545
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Inv Amt2 Desc"
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
         Left            =   0
         TabIndex        =   529
         Top             =   1005
         Width           =   1545
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Inv Amt3 Desc"
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
         Index           =   206
         Left            =   0
         TabIndex        =   528
         Top             =   1290
         Width           =   1545
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Inv Amt4 Desc"
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
         Index           =   207
         Left            =   0
         TabIndex        =   527
         Top             =   1575
         Width           =   1545
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Inv Amt6 Desc"
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
         Index           =   208
         Left            =   0
         TabIndex        =   526
         Top             =   2145
         Width           =   1545
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Inv Amt8 Desc"
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
         Index           =   209
         Left            =   0
         TabIndex        =   525
         Top             =   2715
         Width           =   1545
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Inv Amt5 Desc"
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
         Index           =   210
         Left            =   0
         TabIndex        =   524
         Top             =   1860
         Width           =   1545
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Inv Amt7 Desc"
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
         Index           =   211
         Left            =   0
         TabIndex        =   523
         Top             =   2430
         Width           =   1545
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Inv Amt9 Desc"
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
         Index           =   212
         Left            =   0
         TabIndex        =   522
         Top             =   3000
         Width           =   1545
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Inv Amt10 Desc"
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
         Index           =   213
         Left            =   0
         TabIndex        =   521
         Top             =   3285
         Width           =   1545
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Inv Amt11 Desc"
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
         Index           =   214
         Left            =   0
         TabIndex        =   520
         Top             =   3570
         Width           =   1545
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Inv Amt12 Desc"
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
         Index           =   215
         Left            =   0
         TabIndex        =   519
         Top             =   3855
         Width           =   1545
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Inv Amt1 %"
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
         Index           =   227
         Left            =   8760
         TabIndex        =   518
         Top             =   720
         Width           =   1395
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Inv Amt2 %"
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
         Index           =   226
         Left            =   8760
         TabIndex        =   517
         Top             =   1005
         Width           =   1395
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Inv Amt3 %"
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
         Index           =   225
         Left            =   8760
         TabIndex        =   516
         Top             =   1290
         Width           =   1395
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Inv Amt4 %"
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
         Index           =   224
         Left            =   8760
         TabIndex        =   515
         Top             =   1575
         Width           =   1395
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Inv Amt6 %"
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
         Index           =   223
         Left            =   8760
         TabIndex        =   514
         Top             =   2145
         Width           =   1395
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Inv Amt8 %"
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
         Index           =   222
         Left            =   8760
         TabIndex        =   513
         Top             =   2715
         Width           =   1395
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Inv Amt5 %"
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
         Index           =   221
         Left            =   8760
         TabIndex        =   512
         Top             =   1860
         Width           =   1395
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Inv Amt9 %"
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
         Index           =   219
         Left            =   8760
         TabIndex        =   511
         Top             =   3000
         Width           =   1395
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Inv Amt10 %"
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
         Index           =   218
         Left            =   8760
         TabIndex        =   510
         Top             =   3285
         Width           =   1395
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Inv Amt11 %"
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
         Index           =   217
         Left            =   8760
         TabIndex        =   508
         Top             =   3570
         Width           =   1395
      End
   End
   Begin MwfCtl.MWCTL_BTN1 CmdAddAmt 
      Height          =   465
      Left            =   8880
      TabIndex        =   509
      ToolTipText     =   "Invoice Additional Amounts"
      Top             =   9300
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   820
      BackColor       =   -2147483633
      MaskColor       =   8454143
      Caption         =   "&Additional Amounts"
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
   Begin MwfCtl.MWCTL_BTN1 CmdChgDt 
      Height          =   465
      Left            =   10080
      TabIndex        =   454
      ToolTipText     =   "Copy Option"
      Top             =   9300
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   820
      BackColor       =   -2147483633
      MaskColor       =   8454143
      Caption         =   "&Change Date"
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
      Height          =   405
      Left            =   6900
      TabIndex        =   216
      TabStop         =   0   'False
      Top             =   9330
      Visible         =   0   'False
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   714
   End
   Begin MwfCtl.MWCTL_ADC ADC 
      CausesValidation=   0   'False
      Height          =   555
      Left            =   30
      TabIndex        =   217
      Top             =   9300
      Width           =   14715
      _ExtentX        =   25956
      _ExtentY        =   979
      BtnWidth        =   885
      BtnHeight       =   435
      BtnMaskColor    =   8454143
      FontSz          =   8.25
   End
   Begin VB.Frame FRA_KEYS 
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   -30
      TabIndex        =   220
      Top             =   15
      Width           =   5040
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   40
         Left            =   2600
         TabIndex        =   2
         ToolTipText     =   "Enter Invoice Voucher Year"
         Top             =   0
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   503
         MaxLength       =   2
         DataType        =   4
         DataField       =   "InYy"
         IdName          =   "INYY"
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   42
         Left            =   3850
         TabIndex        =   4
         ToolTipText     =   "Enter Invoice Voucher No."
         Top             =   0
         Width           =   930
         _ExtentX        =   1640
         _ExtentY        =   503
         Alignment       =   2
         Mask            =   "#####0"
         MaxLength       =   6
         DataType        =   1
         DataField       =   "InNo"
         ReCalcParent    =   "INEXPNO"
         IdName          =   "INNO"
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   1
         Left            =   1845
         TabIndex        =   1
         ToolTipText     =   "Enter Invoice Voucher Tc"
         Top             =   0
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   503
         MaxLength       =   3
         DataType        =   4
         DataField       =   "InTc"
         ReCalcParent    =   "INCHR"
         IdName          =   "INTC"
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   2
         Left            =   3100
         TabIndex        =   3
         ToolTipText     =   "Enter Invoice Voucher Character"
         Top             =   0
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   503
         MaxLength       =   3
         DataType        =   4
         DataField       =   "InChr"
         ReCalcOn        =   "INTC"
         IdName          =   "INCHR"
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   116
         Left            =   1120
         TabIndex        =   0
         ToolTipText     =   "Enter Invoice Voucher Company Code"
         Top             =   0
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   503
         MaxLength       =   3
         DataType        =   4
         DataField       =   "InCoCd"
         IdName          =   "INCOCD"
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   173
         Left            =   4830
         TabIndex        =   405
         Top             =   0
         Visible         =   0   'False
         Width           =   120
         _ExtentX        =   212
         _ExtentY        =   503
         Enabled         =   0   'False
         DataField       =   "InPrtKey"
         IdName          =   "INPRTKEY"
      End
      Begin VB.Label ALBL 
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
         Height          =   285
         Index           =   193
         Left            =   1750
         TabIndex        =   570
         Top             =   0
         Width           =   105
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Voucher"
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
         Left            =   60
         TabIndex        =   232
         Top             =   0
         Width           =   1065
      End
      Begin VB.Label ALBL 
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
         Height          =   285
         Index           =   38
         Left            =   2490
         TabIndex        =   231
         Top             =   0
         Width           =   105
      End
      Begin VB.Label ALBL 
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
         Height          =   285
         Index           =   39
         Left            =   3000
         TabIndex        =   229
         Top             =   0
         Width           =   105
      End
      Begin VB.Label ALBL 
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
         Height          =   285
         Index           =   40
         Left            =   3750
         TabIndex        =   227
         Top             =   0
         Width           =   105
      End
   End
   Begin VB.Frame FRA_NKEYS 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   9630
      Left            =   -120
      TabIndex        =   218
      Top             =   0
      Width           =   15315
      Begin VB.Frame FraNKeyAll 
         BorderStyle     =   0  'None
         Height          =   9315
         Left            =   120
         TabIndex        =   222
         Top             =   0
         Width           =   15210
         Begin VB.TextBox TxtPwd 
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   1095
            MaxLength       =   30
            PasswordChar    =   "*"
            TabIndex        =   20
            TabStop         =   0   'False
            ToolTipText     =   "Enter SEO Password "
            Top             =   855
            Width           =   825
         End
         Begin TabDlg.SSTab TabInvDet 
            Height          =   7965
            Left            =   0
            TabIndex        =   28
            Top             =   1230
            Width           =   14745
            _ExtentX        =   26009
            _ExtentY        =   14049
            _Version        =   393216
            Tabs            =   7
            TabsPerRow      =   7
            TabHeight       =   520
            TabCaption(0)   =   "<&1>Customer Dets"
            TabPicture(0)   =   "EmrFrmInvHd.frx":0000
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "FraTabInvDet(0)"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).ControlCount=   1
            TabCaption(1)   =   "<&2>Company Dets"
            TabPicture(1)   =   "EmrFrmInvHd.frx":001C
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "FraTabInvDet(1)"
            Tab(1).ControlCount=   1
            TabCaption(2)   =   "<&3>Amounts"
            TabPicture(2)   =   "EmrFrmInvHd.frx":0038
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "FraTabInvDet(2)"
            Tab(2).ControlCount=   1
            TabCaption(3)   =   "<&4>Doc Dets(1)"
            TabPicture(3)   =   "EmrFrmInvHd.frx":0054
            Tab(3).ControlEnabled=   0   'False
            Tab(3).Control(0)=   "FraTabInvDet(3)"
            Tab(3).ControlCount=   1
            TabCaption(4)   =   "<&5>Doc Dets(2)"
            TabPicture(4)   =   "EmrFrmInvHd.frx":0070
            Tab(4).ControlEnabled=   0   'False
            Tab(4).Control(0)=   "FraTabInvDet(4)"
            Tab(4).ControlCount=   1
            TabCaption(5)   =   "<&6>Doc Dets(3)"
            TabPicture(5)   =   "EmrFrmInvHd.frx":008C
            Tab(5).ControlEnabled=   0   'False
            Tab(5).Control(0)=   "FraTabInvDet(5)"
            Tab(5).ControlCount=   1
            TabCaption(6)   =   "<&7>Defaults"
            TabPicture(6)   =   "EmrFrmInvHd.frx":00A8
            Tab(6).ControlEnabled=   0   'False
            Tab(6).Control(0)=   "FraTabInvDet(6)"
            Tab(6).ControlCount=   1
            Begin VB.Frame FraTabInvDet 
               BorderStyle     =   0  'None
               Height          =   7580
               Index           =   6
               Left            =   -74940
               TabIndex        =   230
               Top             =   360
               Width           =   14625
               Begin VB.Frame Frame3 
                  Height          =   3015
                  Left            =   -30
                  TabIndex        =   334
                  Top             =   2910
                  Width           =   14745
                  Begin MwfCtl.MWCTL_MED atxt 
                     Height          =   570
                     Index           =   133
                     Left            =   1725
                     TabIndex        =   211
                     ToolTipText     =   "Enter State Of Origin"
                     Top             =   420
                     Width           =   9555
                     _ExtentX        =   16854
                     _ExtentY        =   1005
                     Alignment       =   3
                     MaxLength       =   120
                     DataField       =   "InHStatOfOrg"
                     IdName          =   "INHSTATOFORG"
                  End
                  Begin MwfCtl.MWCTL_MED atxt 
                     Height          =   570
                     Index           =   134
                     Left            =   1725
                     TabIndex        =   212
                     ToolTipText     =   "Enter Metal BOE Text"
                     Top             =   990
                     Width           =   9555
                     _ExtentX        =   16854
                     _ExtentY        =   1005
                     Alignment       =   3
                     MaxLength       =   120
                     DataField       =   "InHMetalBOETxt"
                     IdName          =   "INHMETALBOETXT"
                  End
                  Begin MwfCtl.MWCTL_MED atxt 
                     Height          =   630
                     Index           =   141
                     Left            =   1725
                     TabIndex        =   213
                     ToolTipText     =   "Enter RBI Policy Details"
                     Top             =   1560
                     Width           =   9555
                     _ExtentX        =   16854
                     _ExtentY        =   1111
                     Alignment       =   3
                     MaxLength       =   70
                     DataField       =   "InHRbiPolicy"
                     IdName          =   "INHRBIPOLICY"
                  End
                  Begin MwfCtl.MWCTL_MED atxt 
                     Height          =   630
                     Index           =   149
                     Left            =   1725
                     TabIndex        =   214
                     ToolTipText     =   "Enter BOE Details For Loose RM Export"
                     Top             =   2190
                     Width           =   9555
                     _ExtentX        =   16854
                     _ExtentY        =   1111
                     Alignment       =   3
                     MaxLength       =   250
                     DataField       =   "InBOELooseExp"
                     IdName          =   "INBOELOOSEEXP"
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "BOE (Loose Exp)"
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   375
                     Index           =   125
                     Left            =   90
                     TabIndex        =   366
                     Top             =   2190
                     Width           =   1575
                  End
                  Begin VB.Label LblMain 
                     BackStyle       =   0  'Transparent
                     BorderStyle     =   1  'Fixed Single
                     Caption         =   "  Defaults"
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
                     Left            =   -30
                     TabIndex        =   344
                     Top             =   90
                     Width           =   15135
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "RBI Policy"
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   375
                     Index           =   120
                     Left            =   90
                     TabIndex        =   342
                     Top             =   1560
                     Width           =   1275
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "State Of Origin"
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
                     Index           =   113
                     Left            =   90
                     TabIndex        =   336
                     Top             =   420
                     Width           =   1515
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Metal BOE Text"
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   525
                     Index           =   112
                     Left            =   90
                     TabIndex        =   335
                     Top             =   990
                     Width           =   1275
                  End
               End
               Begin VB.Frame Fra4 
                  Height          =   1290
                  Left            =   -30
                  TabIndex        =   378
                  Top             =   1650
                  Width           =   14730
                  Begin MwfCtl.MWCTL_MED atxt 
                     Height          =   285
                     Index           =   97
                     Left            =   6855
                     TabIndex        =   207
                     ToolTipText     =   "Enter Port Of Discharge"
                     Top             =   600
                     Width           =   4105
                     _ExtentX        =   7250
                     _ExtentY        =   503
                     Enabled         =   0   'False
                     MaxLength       =   20
                     DataType        =   4
                     DataField       =   "InCmPortofDisch"
                     ReCalcOn        =   "INCMCD,INCMPOD"
                     IdName          =   "INCMPORTOFDISCH"
                  End
                  Begin MwfCtl.MWCTL_MED atxt 
                     Height          =   285
                     Index           =   98
                     Left            =   6855
                     TabIndex        =   210
                     ToolTipText     =   "Enter Final Destination"
                     Top             =   885
                     Width           =   4105
                     _ExtentX        =   7250
                     _ExtentY        =   503
                     Enabled         =   0   'False
                     MaxLength       =   30
                     DataType        =   4
                     DataField       =   "InCmFinDest"
                     ReCalcOn        =   "INCMCD,INCMFINDSTCD"
                     IdName          =   "INCMFINDEST"
                  End
                  Begin MwfCtl.MWCTL_MED atxt 
                     Height          =   285
                     Index           =   99
                     Left            =   1575
                     TabIndex        =   205
                     ToolTipText     =   "Enter Port Of Origin"
                     Top             =   600
                     Width           =   3105
                     _ExtentX        =   5477
                     _ExtentY        =   503
                     MaxLength       =   20
                     DataType        =   4
                     DataField       =   "InHPortofLd"
                     IdName          =   "INHPORTOFLD"
                  End
                  Begin MwfCtl.MWCTL_MED atxt 
                     Height          =   285
                     Index           =   100
                     Left            =   1575
                     TabIndex        =   208
                     ToolTipText     =   "Enter Country Of Origin"
                     Top             =   885
                     Width           =   3105
                     _ExtentX        =   5477
                     _ExtentY        =   503
                     MaxLength       =   20
                     DataType        =   4
                     DataField       =   "InHConOfOrg"
                     IdName          =   "INHCONOFORG"
                  End
                  Begin MwfCtl.MWCTL_MED atxt 
                     Height          =   285
                     Index           =   209
                     Left            =   5880
                     TabIndex        =   206
                     ToolTipText     =   "Specify Customer Port of Destination"
                     Top             =   600
                     Width           =   960
                     _ExtentX        =   1693
                     _ExtentY        =   503
                     MaxLength       =   6
                     DataType        =   4
                     DataField       =   "InCmPod"
                     ReCalcParent    =   "INCMPORTOFDISCH"
                     ReCalcOn        =   "INCMCD"
                     IdName          =   "INCMPOD"
                  End
                  Begin MwfCtl.MWCTL_MED atxt 
                     Height          =   285
                     Index           =   210
                     Left            =   5880
                     TabIndex        =   209
                     ToolTipText     =   "Specify Customer Final Destination Code"
                     Top             =   885
                     Width           =   495
                     _ExtentX        =   873
                     _ExtentY        =   503
                     MaxLength       =   2
                     DataType        =   4
                     DataField       =   "InCmFinDstCd"
                     ReCalcParent    =   "INCMFINDEST"
                     ReCalcOn        =   "INCMCD"
                     IdName          =   "INCMFINDSTCD"
                  End
                  Begin VB.Label ALBL 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "Destination"
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
                     Index           =   86
                     Left            =   7860
                     TabIndex        =   383
                     Top             =   360
                     Width           =   1065
                  End
                  Begin VB.Label ALBL 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "Origin"
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
                     Index           =   85
                     Left            =   2760
                     TabIndex        =   382
                     Top             =   360
                     Width           =   585
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Country"
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
                     Index           =   84
                     Left            =   90
                     TabIndex        =   381
                     Top             =   885
                     Width           =   1725
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Port"
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
                     TabIndex        =   380
                     Top             =   600
                     Width           =   1845
                  End
                  Begin VB.Label LblMain 
                     BackStyle       =   0  'Transparent
                     BorderStyle     =   1  'Fixed Single
                     Caption         =   "  Place Details"
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
                     Index           =   2
                     Left            =   -30
                     TabIndex        =   379
                     Top             =   90
                     Width           =   15135
                  End
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   125
                  Left            =   2535
                  TabIndex        =   195
                  ToolTipText     =   "Enter Default From Rm Location"
                  Top             =   120
                  Width           =   1830
                  _ExtentX        =   3228
                  _ExtentY        =   503
                  MaxLength       =   8
                  DataType        =   4
                  DataField       =   "InFrRmLoc"
                  IdName          =   "INFRRMLOC"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   126
                  Left            =   2535
                  TabIndex        =   197
                  ToolTipText     =   "Enter Default To Rm Location"
                  Top             =   405
                  Width           =   1830
                  _ExtentX        =   3228
                  _ExtentY        =   503
                  MaxLength       =   8
                  DataType        =   4
                  DataField       =   "InToRmLoc"
                  IdName          =   "INTORMLOC"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   121
                  Left            =   5625
                  TabIndex        =   196
                  ToolTipText     =   "Enter Default Rm Issue/ Receipt Option"
                  Top             =   120
                  Width           =   300
                  _ExtentX        =   529
                  _ExtentY        =   503
                  DataType        =   4
                  ReCalcParent    =   "INFRRMDC,INTORMDC"
                  IdName          =   "WINFRRMDC"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   122
                  Left            =   6315
                  TabIndex        =   325
                  ToolTipText     =   "Enter Default D/C Option For To Location"
                  Top             =   120
                  Visible         =   0   'False
                  Width           =   300
                  _ExtentX        =   529
                  _ExtentY        =   503
                  Enabled         =   0   'False
                  DataType        =   4
                  DataField       =   "InToRmDc"
                  ReCalcOn        =   "WINFRRMDC"
                  IdName          =   "INTORMDC"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   123
                  Left            =   6015
                  TabIndex        =   326
                  ToolTipText     =   "Enter Default D/C Option For From Location"
                  Top             =   120
                  Visible         =   0   'False
                  Width           =   300
                  _ExtentX        =   529
                  _ExtentY        =   503
                  Enabled         =   0   'False
                  DataType        =   4
                  DataField       =   "InFrRmDc"
                  ReCalcOn        =   "WINFRRMDC"
                  IdName          =   "INFRRMDC"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   130
                  Left            =   2535
                  TabIndex        =   198
                  ToolTipText     =   "Enter Whether Value Addion On Metal Column Should Be Displayed in Value Addition Report?"
                  Top             =   690
                  Width           =   300
                  _ExtentX        =   529
                  _ExtentY        =   503
                  DataType        =   4
                  DataField       =   "InHValAddMtlColYN"
                  IdName          =   "INHVALADDMTLCOLYN"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   300
                  Index           =   131
                  Left            =   2535
                  TabIndex        =   199
                  ToolTipText     =   "Enter Whether BOE DEtails Should Be Displayed in Value Addition Annexure ?"
                  Top             =   975
                  Width           =   300
                  _ExtentX        =   529
                  _ExtentY        =   529
                  DataType        =   4
                  DataField       =   "InHValBOEDetYN"
                  IdName          =   "INHVALBOEDETYN"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   143
                  Left            =   1725
                  TabIndex        =   215
                  ToolTipText     =   "Enter GSP Package Description"
                  Top             =   6090
                  Width           =   4470
                  _ExtentX        =   7885
                  _ExtentY        =   503
                  MaxLength       =   30
                  DataField       =   "InHGspPkgDesc"
                  IdName          =   "INHGSPPKGDESC"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   161
                  Left            =   9930
                  TabIndex        =   201
                  ToolTipText     =   "Specify Local Invoice (Yes/No)"
                  Top             =   120
                  Width           =   300
                  _ExtentX        =   529
                  _ExtentY        =   503
                  Alignment       =   3
                  DataType        =   4
                  DataField       =   "InLclInvYn"
                  ReCalcParent    =   $"EmrFrmInvHd.frx":00C4
                  ReCalcOn        =   "INCMCD"
                  IdName          =   "INLCLINVYN"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   162
                  Left            =   9930
                  TabIndex        =   202
                  ToolTipText     =   "Enter Local Invoice Type"
                  Top             =   405
                  Width           =   3090
                  _ExtentX        =   5450
                  _ExtentY        =   503
                  Alignment       =   3
                  MaxLength       =   20
                  DataType        =   4
                  DataField       =   "InLclInvTyp"
                  IdName          =   "INLCLINVTYP"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   300
                  Index           =   170
                  Left            =   2535
                  TabIndex        =   200
                  ToolTipText     =   "Specify Whether Designs Are On Hold. (Enter Y/ N)"
                  Top             =   1275
                  Width           =   300
                  _ExtentX        =   529
                  _ExtentY        =   529
                  DataType        =   4
                  DataField       =   "InDsgOnHld"
                  IdName          =   "INDSGONHLD"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   171
                  Left            =   9930
                  TabIndex        =   203
                  ToolTipText     =   "Specify Remake Rm Ctg"
                  Top             =   690
                  Width           =   900
                  _ExtentX        =   1588
                  _ExtentY        =   503
                  MaxLength       =   5
                  DataType        =   4
                  DataField       =   "InRemakingCtgs"
                  IdName          =   "INREMAKINGCTGS"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   172
                  Left            =   9930
                  TabIndex        =   204
                  ToolTipText     =   "Specify Whether The Copy Options In Invoice Design Entry Will Keep The Invoice Gross Weight Equal To Actual Gross Weight (Yes/ No)"
                  Top             =   975
                  Width           =   300
                  _ExtentX        =   529
                  _ExtentY        =   503
                  DataType        =   4
                  DataField       =   "InInvGrsWtEqActual"
                  IdName          =   "ININVGRSWTEQACTUAL"
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Inv Grs Wt Equal To Actual"
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
                  Left            =   7320
                  TabIndex        =   404
                  Top             =   975
                  Width           =   2670
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Remake RmCtg"
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
                  Index           =   144
                  Left            =   7320
                  TabIndex        =   403
                  Top             =   690
                  Width           =   1890
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Designs On Hold"
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
                  Left            =   90
                  TabIndex        =   402
                  Top             =   1260
                  Width           =   1890
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Local Invoice Type"
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
                  Index           =   135
                  Left            =   7320
                  TabIndex        =   398
                  Top             =   405
                  Width           =   1815
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Local Invoice"
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
                  Index           =   129
                  Left            =   7320
                  TabIndex        =   397
                  Top             =   120
                  Width           =   1395
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "GSP Pkg Desc"
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
                  Index           =   117
                  Left            =   90
                  TabIndex        =   343
                  Top             =   6090
                  Width           =   1545
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "From Rm Loc"
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
                  Index           =   102
                  Left            =   90
                  TabIndex        =   331
                  Top             =   120
                  Width           =   1560
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "I/R"
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
                  Index           =   103
                  Left            =   5280
                  TabIndex        =   330
                  Top             =   120
                  Width           =   345
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "To Rm Loc"
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
                  Index           =   104
                  Left            =   90
                  TabIndex        =   329
                  Top             =   390
                  Width           =   1350
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Value Addition on Metal "
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
                  Index           =   109
                  Left            =   90
                  TabIndex        =   328
                  Top             =   675
                  Width           =   2370
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Show BOE Details"
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
                  Index           =   110
                  Left            =   90
                  TabIndex        =   327
                  Top             =   960
                  Width           =   1890
               End
            End
            Begin VB.Frame FraTabInvDet 
               BorderStyle     =   0  'None
               Height          =   7580
               Index           =   5
               Left            =   -74940
               TabIndex        =   228
               Top             =   360
               Width           =   14655
               Begin VB.Frame FraDef5 
                  Height          =   3465
                  Left            =   -30
                  TabIndex        =   373
                  Top             =   4260
                  Width           =   14775
                  Begin MwfCtl.MWCTL_MED atxt 
                     Height          =   285
                     Index           =   155
                     Left            =   1545
                     TabIndex        =   192
                     ToolTipText     =   "Enter Local Invoice Footer"
                     Top             =   960
                     Width           =   12585
                     _ExtentX        =   22199
                     _ExtentY        =   503
                     Alignment       =   3
                     MaxLength       =   90
                     DataField       =   "InHLclInvFT"
                     IdName          =   "INHLCLINVFT"
                  End
                  Begin MwfCtl.MWCTL_MED atxt 
                     Height          =   1005
                     Index           =   156
                     Left            =   1545
                     TabIndex        =   193
                     ToolTipText     =   "Enter VAT Footer"
                     Top             =   1245
                     Width           =   12585
                     _ExtentX        =   22199
                     _ExtentY        =   1773
                     Alignment       =   3
                     MaxLength       =   750
                     DataField       =   "InHLclInvVATFT"
                     IdName          =   "INHLCLINVVATFT"
                  End
                  Begin MwfCtl.MWCTL_MED atxt 
                     Height          =   1005
                     Index           =   157
                     Left            =   1545
                     TabIndex        =   194
                     ToolTipText     =   "Enter CST Footer"
                     Top             =   2250
                     Width           =   12585
                     _ExtentX        =   22199
                     _ExtentY        =   1773
                     Alignment       =   3
                     MaxLength       =   750
                     DataField       =   "InHLclInvCSTFT"
                     IdName          =   "INHLCLINVCSTFT"
                  End
                  Begin MwfCtl.MWCTL_MED atxt 
                     Height          =   285
                     Index           =   200
                     Left            =   1545
                     TabIndex        =   190
                     ToolTipText     =   "Enter Customer CST Number"
                     Top             =   675
                     Width           =   8340
                     _ExtentX        =   14711
                     _ExtentY        =   503
                     MaxLength       =   60
                     DataType        =   4
                     DataField       =   "InCmCSTNO"
                     IdName          =   "INCMCSTNO"
                  End
                  Begin MwfCtl.MWCTL_MED atxt 
                     Height          =   285
                     Index           =   201
                     Left            =   1545
                     TabIndex        =   188
                     ToolTipText     =   "Enter Customer VAT Number"
                     Top             =   390
                     Width           =   8340
                     _ExtentX        =   14711
                     _ExtentY        =   503
                     MaxLength       =   60
                     DataType        =   4
                     DataField       =   "InCmVATNO"
                     IdName          =   "INCMVATNO"
                  End
                  Begin MwfCtl.MWCTL_MED atxt 
                     Height          =   285
                     Index           =   202
                     Left            =   11025
                     TabIndex        =   189
                     ToolTipText     =   "Enter Customer Tin No"
                     Top             =   390
                     Width           =   3105
                     _ExtentX        =   5477
                     _ExtentY        =   503
                     MaxLength       =   20
                     DataType        =   4
                     DataField       =   "InCmTinNo"
                     IdName          =   "INCMTINNO"
                  End
                  Begin MwfCtl.MWCTL_MED atxt 
                     Height          =   285
                     Index           =   203
                     Left            =   11025
                     TabIndex        =   191
                     ToolTipText     =   "Enter Customer PAN"
                     Top             =   675
                     Width           =   3105
                     _ExtentX        =   5477
                     _ExtentY        =   503
                     MaxLength       =   20
                     DataType        =   4
                     DataField       =   "InCmPAN"
                     IdName          =   "INCMPAN"
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Cust PAN"
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
                     Index           =   161
                     Left            =   9960
                     TabIndex        =   446
                     Top             =   675
                     Width           =   1065
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Cust CST No"
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
                     Index           =   160
                     Left            =   60
                     TabIndex        =   445
                     Top             =   675
                     Width           =   1425
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Cust VAT No"
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
                     Index           =   159
                     Left            =   60
                     TabIndex        =   444
                     Top             =   390
                     Width           =   1425
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Cust TIN"
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
                     Index           =   158
                     Left            =   9960
                     TabIndex        =   443
                     Top             =   390
                     Width           =   1065
                  End
                  Begin VB.Label LblMain 
                     BackStyle       =   0  'Transparent
                     BorderStyle     =   1  'Fixed Single
                     Caption         =   "  Local Invoice"
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
                     Index           =   19
                     Left            =   -30
                     TabIndex        =   377
                     Top             =   90
                     Width           =   14805
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "CST Footer"
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
                     Left            =   60
                     TabIndex        =   376
                     Top             =   2250
                     Width           =   1275
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "VAT Footer"
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
                     Index           =   133
                     Left            =   60
                     TabIndex        =   375
                     Top             =   1245
                     Width           =   1275
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Invoice Footer"
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
                     Index           =   132
                     Left            =   60
                     TabIndex        =   374
                     Top             =   960
                     Width           =   1425
                  End
               End
               Begin VB.Frame FraDef4 
                  Height          =   2805
                  Left            =   -30
                  TabIndex        =   312
                  Top             =   1530
                  Width           =   14790
                  Begin MwfCtl.MWCTL_MED atxt 
                     Height          =   600
                     Index           =   117
                     Left            =   1575
                     TabIndex        =   185
                     ToolTipText     =   "Enter Exim Policy Details"
                     Top             =   420
                     Width           =   11205
                     _ExtentX        =   19764
                     _ExtentY        =   1058
                     Alignment       =   3
                     MaxLength       =   180
                     DataField       =   "InHEximPolicy"
                     IdName          =   "INHEXIMPOLICY"
                  End
                  Begin MwfCtl.MWCTL_MED atxt 
                     Height          =   840
                     Index           =   118
                     Left            =   1575
                     TabIndex        =   187
                     ToolTipText     =   "Enter Metal Utilization Details"
                     Top             =   1860
                     Width           =   11205
                     _ExtentX        =   19764
                     _ExtentY        =   1482
                     Alignment       =   3
                     MaxLength       =   240
                     DataField       =   "InMetalUtil"
                     IdName          =   "INMETALUTIL"
                  End
                  Begin MwfCtl.MWCTL_MED atxt 
                     Height          =   840
                     Index           =   119
                     Left            =   1575
                     TabIndex        =   186
                     ToolTipText     =   "Enter GSP Eligibility Details"
                     Top             =   1020
                     Width           =   11205
                     _ExtentX        =   19764
                     _ExtentY        =   1482
                     Alignment       =   3
                     MaxLength       =   225
                     DataField       =   "InHGspEligibility"
                     IdName          =   "INHGSPELIGIBILITY"
                  End
                  Begin VB.Label LblMain 
                     BackStyle       =   0  'Transparent
                     BorderStyle     =   1  'Fixed Single
                     Caption         =   "  DTA Details"
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
                     TabIndex        =   345
                     Top             =   90
                     Width           =   15135
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Exim Policy"
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
                     Left            =   30
                     TabIndex        =   315
                     Top             =   420
                     Width           =   1275
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "GSP Eligibility"
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
                     Index           =   98
                     Left            =   30
                     TabIndex        =   314
                     Top             =   1110
                     Width           =   1515
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Metal Utilization"
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   495
                     Index           =   100
                     Left            =   30
                     TabIndex        =   313
                     Top             =   1965
                     Width           =   1635
                  End
               End
               Begin VB.Frame Frame1 
                  BorderStyle     =   0  'None
                  Height          =   1605
                  Left            =   -30
                  TabIndex        =   316
                  Top             =   -30
                  Width           =   14640
                  Begin MwfCtl.MWCTL_MED atxt 
                     Height          =   285
                     Index           =   101
                     Left            =   6375
                     TabIndex        =   183
                     ToolTipText     =   "Enter ETC Zone"
                     Top             =   945
                     Width           =   3105
                     _ExtentX        =   5477
                     _ExtentY        =   503
                     MaxLength       =   20
                     DataType        =   4
                     DataField       =   "InHEtcZone"
                     IdName          =   "INHETCZONE"
                  End
                  Begin MwfCtl.MWCTL_MED atxt 
                     Height          =   285
                     Index           =   102
                     Left            =   6375
                     TabIndex        =   184
                     ToolTipText     =   "Enter ETC Section"
                     Top             =   1230
                     Width           =   4725
                     _ExtentX        =   8334
                     _ExtentY        =   503
                     MaxLength       =   30
                     DataType        =   4
                     DataField       =   "InHEtcSection"
                     IdName          =   "INHETCSECTION"
                  End
                  Begin MwfCtl.MWCTL_MED atxt 
                     Height          =   285
                     Index           =   103
                     Left            =   1575
                     TabIndex        =   181
                     ToolTipText     =   "Enter Marks"
                     Top             =   945
                     Width           =   3105
                     _ExtentX        =   5477
                     _ExtentY        =   503
                     MaxLength       =   20
                     DataType        =   4
                     DataField       =   "InHMarks"
                     IdName          =   "INHMARKS"
                  End
                  Begin MwfCtl.MWCTL_MED atxt 
                     Height          =   285
                     Index           =   104
                     Left            =   6375
                     TabIndex        =   177
                     ToolTipText     =   "Enter Package Description"
                     Top             =   90
                     Width           =   3105
                     _ExtentX        =   5477
                     _ExtentY        =   503
                     MaxLength       =   20
                     DataField       =   "InHPkgDesc"
                     IdName          =   "INHPKGDESC"
                  End
                  Begin MwfCtl.MWCTL_MED atxt 
                     Height          =   285
                     Index           =   105
                     Left            =   1575
                     TabIndex        =   176
                     ToolTipText     =   "Enter Pre Carriage"
                     Top             =   90
                     Width           =   3105
                     _ExtentX        =   5477
                     _ExtentY        =   503
                     MaxLength       =   20
                     DataField       =   "InHPreCarriage"
                     IdName          =   "INHPRECARRIAGE"
                  End
                  Begin MwfCtl.MWCTL_MED atxt 
                     Height          =   285
                     Index           =   106
                     Left            =   1575
                     TabIndex        =   178
                     ToolTipText     =   "Enter Pre Carrier"
                     Top             =   375
                     Width           =   3105
                     _ExtentX        =   5477
                     _ExtentY        =   503
                     MaxLength       =   20
                     DataType        =   4
                     DataField       =   "InHPreCarrier"
                     IdName          =   "INHPRECARRIER"
                  End
                  Begin MwfCtl.MWCTL_MED atxt 
                     Height          =   285
                     Index           =   107
                     Left            =   1575
                     TabIndex        =   180
                     ToolTipText     =   "Enter Vessel"
                     Top             =   660
                     Width           =   3105
                     _ExtentX        =   5477
                     _ExtentY        =   503
                     MaxLength       =   20
                     DataType        =   4
                     DataField       =   "InHvessel"
                     ReCalcOn        =   "INCMCD"
                     IdName          =   "INHVESSEL"
                  End
                  Begin MwfCtl.MWCTL_MED atxt 
                     Height          =   285
                     Index           =   108
                     Left            =   6375
                     TabIndex        =   179
                     ToolTipText     =   "Enter Package Weight"
                     Top             =   375
                     Width           =   945
                     _ExtentX        =   1667
                     _ExtentY        =   503
                     Alignment       =   2
                     Mask            =   "####0.000"
                     MaxLength       =   9
                     DataType        =   2
                     DataField       =   "InPkgWt"
                     IdName          =   "INPKGWT"
                  End
                  Begin MwfCtl.MWCTL_MED atxt 
                     Height          =   285
                     Index           =   136
                     Left            =   6375
                     TabIndex        =   182
                     ToolTipText     =   "Enter Export Trade Control Text for Reports "
                     Top             =   660
                     Width           =   4725
                     _ExtentX        =   8334
                     _ExtentY        =   503
                     MaxLength       =   30
                     DataField       =   "InHEtcTxt"
                     IdName          =   "INHETCTXT"
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "ETC Text"
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
                     Left            =   5040
                     TabIndex        =   338
                     Top             =   660
                     Width           =   1365
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "gms."
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
                     Left            =   7440
                     TabIndex        =   333
                     Top             =   375
                     Width           =   585
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Pkg Wt"
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
                     Index           =   88
                     Left            =   5040
                     TabIndex        =   324
                     Top             =   375
                     Width           =   1185
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Pkg Desc"
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
                     Index           =   89
                     Left            =   5040
                     TabIndex        =   323
                     Top             =   90
                     Width           =   1785
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Marks"
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
                     Index           =   90
                     Left            =   90
                     TabIndex        =   322
                     Top             =   945
                     Width           =   1365
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Vessel"
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
                     Index           =   91
                     Left            =   90
                     TabIndex        =   321
                     Top             =   660
                     Width           =   1545
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Pre Carrier"
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
                     Index           =   92
                     Left            =   90
                     TabIndex        =   320
                     Top             =   375
                     Width           =   1545
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Pre Carriage"
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
                     Index           =   93
                     Left            =   90
                     TabIndex        =   319
                     Top             =   90
                     Width           =   1545
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "ETC Zone"
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
                     Index           =   94
                     Left            =   5040
                     TabIndex        =   318
                     Top             =   945
                     Width           =   1365
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "ETC Section"
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
                     Index           =   95
                     Left            =   5040
                     TabIndex        =   317
                     Top             =   1230
                     Width           =   1365
                  End
               End
            End
            Begin VB.Frame FraTabInvDet 
               BorderStyle     =   0  'None
               Height          =   7580
               Index           =   4
               Left            =   -74940
               TabIndex        =   226
               Top             =   360
               Width           =   14655
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   825
                  Index           =   90
                  Left            =   1725
                  TabIndex        =   166
                  ToolTipText     =   "Enter Invoice Footer Remarks"
                  Top             =   1455
                  Width           =   9765
                  _ExtentX        =   17224
                  _ExtentY        =   1455
                  Alignment       =   3
                  MaxLength       =   360
                  DataField       =   "InHFtRem"
                  IdName          =   "INHFTREM"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   1065
                  Index           =   6
                  Left            =   1725
                  TabIndex        =   167
                  ToolTipText     =   "Enter Bank Instructions"
                  Top             =   2280
                  Width           =   9765
                  _ExtentX        =   17224
                  _ExtentY        =   1879
                  Alignment       =   3
                  MaxLength       =   650
                  DataField       =   "InHBnkInstr"
                  ReCalcOn        =   "WBKCD"
                  IdName          =   "INHBNKINSTR"
                  Data            =   $"EmrFrmInvHd.frx":0158
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   825
                  Index           =   110
                  Left            =   1725
                  TabIndex        =   165
                  ToolTipText     =   "Enter Product Description"
                  Top             =   630
                  Width           =   9765
                  _ExtentX        =   17224
                  _ExtentY        =   1455
                  Alignment       =   3
                  MaxLength       =   180
                  DataField       =   "InPrdDesc"
                  ReCalcOn        =   "WINVPRDDESC"
                  IdName          =   "INPRDDESC"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   89
                  Left            =   1725
                  TabIndex        =   174
                  ToolTipText     =   "Enter Business Identification Number"
                  Top             =   7110
                  Width           =   3105
                  _ExtentX        =   5477
                  _ExtentY        =   503
                  MaxLength       =   20
                  DataType        =   4
                  DataField       =   "InHBinNo"
                  IdName          =   "INHBINNO"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   795
                  Index           =   115
                  Left            =   1725
                  TabIndex        =   169
                  ToolTipText     =   "Enter United Nations Resolutions"
                  Top             =   4140
                  Width           =   9765
                  _ExtentX        =   17224
                  _ExtentY        =   1402
                  Alignment       =   3
                  MaxLength       =   0
                  DataField       =   "InHUnResolutions"
                  IdName          =   "INHUNRESOLUTIONS"
                  Data            =   $"EmrFrmInvHd.frx":02E6
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   120
                  Left            =   1725
                  TabIndex        =   164
                  ToolTipText     =   "Enter Invoice Product Description Code"
                  Top             =   345
                  Width           =   1335
                  _ExtentX        =   2355
                  _ExtentY        =   503
                  MaxLength       =   8
                  DataType        =   4
                  ReCalcParent    =   "INPRDDESC"
                  IdName          =   "WINVPRDDESC"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   795
                  Index           =   114
                  Left            =   1725
                  TabIndex        =   168
                  ToolTipText     =   "Enter Value Addition Footer"
                  Top             =   3345
                  Width           =   9765
                  _ExtentX        =   17224
                  _ExtentY        =   1402
                  Alignment       =   3
                  MaxLength       =   290
                  DataField       =   "InHValAddnFt"
                  IdName          =   "INHVALADDNFT"
                  Data            =   $"EmrFrmInvHd.frx":03BB
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   132
                  Left            =   1725
                  TabIndex        =   163
                  ToolTipText     =   "Enter MMTC Code"
                  Top             =   60
                  Width           =   2970
                  _ExtentX        =   5239
                  _ExtentY        =   503
                  MaxLength       =   20
                  DataField       =   "InMMTC"
                  IdName          =   "INMMTC"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   140
                  Left            =   1725
                  TabIndex        =   172
                  ToolTipText     =   "Enter GJEPC Number"
                  Top             =   6525
                  Width           =   4920
                  _ExtentX        =   8678
                  _ExtentY        =   503
                  MaxLength       =   30
                  DataType        =   4
                  DataField       =   "InHGjepcNo"
                  IdName          =   "INHGJEPCNO"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   142
                  Left            =   1725
                  TabIndex        =   173
                  ToolTipText     =   "Enter GJEPC Valid UpTo"
                  Top             =   6825
                  Width           =   4920
                  _ExtentX        =   8678
                  _ExtentY        =   503
                  MaxLength       =   30
                  DataField       =   "InHGjepcValidUpTo"
                  IdName          =   "INHGJEPCVALIDUPTO"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   795
                  Index           =   144
                  Left            =   1725
                  TabIndex        =   170
                  ToolTipText     =   "Enter Customer Remark"
                  Top             =   4935
                  Width           =   9765
                  _ExtentX        =   17224
                  _ExtentY        =   1402
                  Alignment       =   3
                  MaxLength       =   800
                  DataField       =   "INHCUSTREM"
                  IdName          =   "INHCUSTREM"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   795
                  Index           =   193
                  Left            =   1725
                  TabIndex        =   171
                  ToolTipText     =   "Enter Invoice/Memo Terms & Conditions"
                  Top             =   5730
                  Width           =   9765
                  _ExtentX        =   17224
                  _ExtentY        =   1402
                  Alignment       =   3
                  MaxLength       =   1500
                  DataField       =   "InTermCond"
                  IdName          =   "INTERMCOND"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   3465
                  Index           =   260
                  Left            =   11520
                  TabIndex        =   175
                  ToolTipText     =   "Enter HS Code for GSP"
                  Top             =   630
                  Width           =   3045
                  _ExtentX        =   5371
                  _ExtentY        =   6112
                  Alignment       =   3
                  MaxLength       =   250
                  DataField       =   "INHSCODE"
                  IdName          =   "INHSCODE"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   1000
                  Left            =   6720
                  TabIndex        =   554
                  ToolTipText     =   "Enter Central Excise Regn Number"
                  Top             =   7110
                  Width           =   3105
                  _ExtentX        =   5477
                  _ExtentY        =   503
                  MaxLength       =   20
                  DataType        =   4
                  DataField       =   "InCERegnNo"
                  IdName          =   "INCEREGNNO"
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "(Prints max 750 characters)"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   495
                  Index           =   197
                  Left            =   90
                  TabIndex        =   576
                  Top             =   4320
                  Width           =   1425
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "CE Regn No"
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
                  Index           =   1000
                  Left            =   5280
                  TabIndex        =   555
                  Top             =   7110
                  Width           =   1365
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "GSP HS Code"
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
                  Index           =   194
                  Left            =   11520
                  TabIndex        =   452
                  Top             =   345
                  Width           =   1635
               End
               Begin VB.Label LblInTermCond 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Terms && Cond"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   765
                  Left            =   60
                  TabIndex        =   434
                  Top             =   5730
                  Width           =   1635
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Customer Rmk"
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
                  Index           =   118
                  Left            =   90
                  TabIndex        =   361
                  Top             =   4935
                  Width           =   1635
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "ValAddn Footer"
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
                  Index           =   74
                  Left            =   90
                  TabIndex        =   360
                  Top             =   3345
                  Width           =   1785
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "MMTC Code"
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
                  Index           =   106
                  Left            =   90
                  TabIndex        =   359
                  Top             =   60
                  Width           =   1365
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Prd Desc Cd"
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
                  Index           =   101
                  Left            =   90
                  TabIndex        =   358
                  Top             =   345
                  Width           =   1365
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "UN Resolutions"
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
                  Index           =   97
                  Left            =   90
                  TabIndex        =   357
                  Top             =   4140
                  Width           =   1785
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Bin No."
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
                  Left            =   90
                  TabIndex        =   356
                  Top             =   7110
                  Width           =   1365
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Prod Desc"
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
                  Index           =   96
                  Left            =   90
                  TabIndex        =   355
                  Top             =   630
                  Width           =   1455
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Bank Instr"
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
                  TabIndex        =   354
                  Top             =   2280
                  Width           =   1785
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Inv Footer Rem"
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
                  Index           =   76
                  Left            =   90
                  TabIndex        =   353
                  Top             =   1455
                  Width           =   1635
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "GJEPC No."
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
                  Index           =   119
                  Left            =   90
                  TabIndex        =   352
                  Top             =   6540
                  Width           =   1365
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Valid UpTo"
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
                  Index           =   121
                  Left            =   90
                  TabIndex        =   351
                  Top             =   6825
                  Width           =   1395
               End
            End
            Begin VB.Frame FraTabInvDet 
               BorderStyle     =   0  'None
               Height          =   7580
               Index           =   3
               Left            =   -74940
               TabIndex        =   225
               Top             =   360
               Width           =   14655
               Begin VB.Frame FraDefDta 
                  Height          =   1905
                  Left            =   0
                  TabIndex        =   421
                  Top             =   5700
                  Width           =   12945
                  Begin MwfCtl.MWCTL_MED atxt 
                     Height          =   285
                     Index           =   226
                     Left            =   1605
                     TabIndex        =   158
                     ToolTipText     =   "Enter Gsp Officer Address"
                     Top             =   390
                     Width           =   7410
                     _ExtentX        =   13070
                     _ExtentY        =   503
                     MaxLength       =   45
                     DataField       =   "InGspAdd1"
                     IdName          =   "INGSPADD1"
                  End
                  Begin MwfCtl.MWCTL_MED atxt 
                     Height          =   285
                     Index           =   228
                     Left            =   1605
                     TabIndex        =   159
                     ToolTipText     =   "Enter Gsp Officer Address"
                     Top             =   675
                     Width           =   7410
                     _ExtentX        =   13070
                     _ExtentY        =   503
                     MaxLength       =   45
                     DataField       =   "InGspAdd2"
                     IdName          =   "INGSPADD2"
                  End
                  Begin MwfCtl.MWCTL_MED atxt 
                     Height          =   285
                     Index           =   229
                     Left            =   1605
                     TabIndex        =   160
                     ToolTipText     =   "Enter Gsp Officer Address"
                     Top             =   960
                     Width           =   7410
                     _ExtentX        =   13070
                     _ExtentY        =   503
                     MaxLength       =   45
                     DataField       =   "InGspAdd3"
                     IdName          =   "INGSPADD3"
                  End
                  Begin MwfCtl.MWCTL_MED atxt 
                     Height          =   285
                     Index           =   230
                     Left            =   1605
                     TabIndex        =   161
                     ToolTipText     =   "Enter Registration No"
                     Top             =   1260
                     Width           =   1665
                     _ExtentX        =   2937
                     _ExtentY        =   503
                     MaxLength       =   10
                     DataType        =   4
                     DataField       =   "InRegNo"
                     IdName          =   "INREGNO"
                  End
                  Begin MwfCtl.MWCTL_MED atxt 
                     Height          =   285
                     Index           =   231
                     Left            =   1605
                     TabIndex        =   162
                     ToolTipText     =   "Enter Circular No"
                     Top             =   1560
                     Width           =   4170
                     _ExtentX        =   7355
                     _ExtentY        =   503
                     MaxLength       =   30
                     DataField       =   "InCircularNo"
                     IdName          =   "INCIRCULARNO"
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Reg No"
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
                     Index           =   166
                     Left            =   30
                     TabIndex        =   425
                     Top             =   1260
                     Width           =   1485
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Circular No"
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
                     Index           =   168
                     Left            =   30
                     TabIndex        =   424
                     Top             =   1560
                     Width           =   1485
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "GSP Officer Address"
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   495
                     Index           =   167
                     Left            =   60
                     TabIndex        =   423
                     Top             =   390
                     Width           =   1455
                  End
                  Begin VB.Label LblMain 
                     BackStyle       =   0  'Transparent
                     BorderStyle     =   1  'Fixed Single
                     Caption         =   "  DTA Details"
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
                     Left            =   0
                     TabIndex        =   422
                     Top             =   90
                     Width           =   13215
                  End
               End
               Begin VB.Frame FraIns 
                  Height          =   2475
                  Left            =   -30
                  TabIndex        =   289
                  Top             =   3270
                  Width           =   6405
                  Begin MwfCtl.MWCTL_MED atxt 
                     Height          =   285
                     Index           =   65
                     Left            =   1455
                     TabIndex        =   147
                     ToolTipText     =   "Enter Insurance Number"
                     Top             =   705
                     Width           =   4725
                     _ExtentX        =   8334
                     _ExtentY        =   503
                     MaxLength       =   30
                     DataType        =   4
                     DataField       =   "InHInsNo"
                     IdName          =   "INHINSNO"
                  End
                  Begin MwfCtl.MWCTL_MED atxt 
                     Height          =   285
                     Index           =   66
                     Left            =   1455
                     TabIndex        =   146
                     ToolTipText     =   "Enter Insurance By"
                     Top             =   420
                     Width           =   4725
                     _ExtentX        =   8334
                     _ExtentY        =   503
                     MaxLength       =   30
                     DataField       =   "InCmInsBy"
                     ReCalcOn        =   "INCMCD"
                     IdName          =   "INCMINSBY"
                  End
                  Begin MwfCtl.MWCTL_MED atxt 
                     Height          =   285
                     Index           =   67
                     Left            =   1455
                     TabIndex        =   148
                     ToolTipText     =   "Enter Agent Name"
                     Top             =   990
                     Width           =   4725
                     _ExtentX        =   8334
                     _ExtentY        =   503
                     MaxLength       =   30
                     DataType        =   4
                     DataField       =   "InHAgName"
                     IdName          =   "INHAGNAME"
                  End
                  Begin MwfCtl.MWCTL_MED atxt 
                     Height          =   285
                     Index           =   68
                     Left            =   1455
                     TabIndex        =   149
                     ToolTipText     =   "Enter Agent Address"
                     Top             =   1545
                     Width           =   4725
                     _ExtentX        =   8334
                     _ExtentY        =   503
                     MaxLength       =   30
                     DataType        =   4
                     DataField       =   "InHAgAdd1"
                     IdName          =   "INHAGADD1"
                  End
                  Begin MwfCtl.MWCTL_MED atxt 
                     Height          =   285
                     Index           =   69
                     Left            =   1455
                     TabIndex        =   150
                     ToolTipText     =   "Enter Agent Address"
                     Top             =   1830
                     Width           =   4725
                     _ExtentX        =   8334
                     _ExtentY        =   503
                     MaxLength       =   30
                     DataType        =   4
                     DataField       =   "InHAgAdd2"
                     IdName          =   "INHAGADD2"
                  End
                  Begin MwfCtl.MWCTL_MED atxt 
                     Height          =   285
                     Index           =   70
                     Left            =   1455
                     TabIndex        =   151
                     ToolTipText     =   "Enter Agent Address"
                     Top             =   2115
                     Width           =   4725
                     _ExtentX        =   8334
                     _ExtentY        =   503
                     MaxLength       =   30
                     DataType        =   4
                     DataField       =   "InHAgAdd3"
                     IdName          =   "INHAGADD3"
                  End
                  Begin VB.Label LblMain 
                     BackStyle       =   0  'Transparent
                     BorderStyle     =   1  'Fixed Single
                     Caption         =   "  Insurance Details"
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
                     Index           =   3
                     Left            =   -30
                     TabIndex        =   346
                     Top             =   90
                     Width           =   6405
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Number"
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
                     Left            =   90
                     TabIndex        =   293
                     Top             =   705
                     Width           =   1275
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Insurance By"
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
                     Left            =   90
                     TabIndex        =   292
                     Top             =   420
                     Width           =   1395
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Shipping Bill Agent Name"
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
                     Index           =   61
                     Left            =   90
                     TabIndex        =   291
                     Top             =   990
                     Width           =   1455
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Agent Addr"
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
                     Left            =   90
                     TabIndex        =   290
                     Top             =   1575
                     Width           =   1575
                  End
               End
               Begin VB.Frame Fra3 
                  Height          =   2475
                  Left            =   6330
                  TabIndex        =   282
                  Top             =   3270
                  Width           =   6615
                  Begin MwfCtl.MWCTL_MED atxt 
                     Height          =   285
                     Index           =   13
                     Left            =   1605
                     TabIndex        =   157
                     ToolTipText     =   "Enter Export Trade Control Saving"
                     Top             =   1995
                     Width           =   3105
                     _ExtentX        =   5477
                     _ExtentY        =   503
                     MaxLength       =   20
                     DataField       =   "InHETCSvg"
                     IdName          =   "INHETCSVG"
                  End
                  Begin MwfCtl.MWCTL_MED atxt 
                     Height          =   285
                     Index           =   92
                     Left            =   1605
                     TabIndex        =   153
                     ToolTipText     =   "Enter AR4/ AR4A No. And Date (GR & Shipping Bill)"
                     Top             =   855
                     Width           =   3105
                     _ExtentX        =   5477
                     _ExtentY        =   503
                     MaxLength       =   20
                     DataField       =   "InHAR4No"
                     IdName          =   "INHAR4NO"
                  End
                  Begin MwfCtl.MWCTL_MED atxt 
                     Height          =   285
                     Index           =   93
                     Left            =   1605
                     TabIndex        =   154
                     ToolTipText     =   "Enter Q/Cert. No & Date (GR & Shipping Bill)"
                     Top             =   1140
                     Width           =   3105
                     _ExtentX        =   5477
                     _ExtentY        =   503
                     MaxLength       =   20
                     DataField       =   "InHQCertNo"
                     IdName          =   "INHQCERTNO"
                  End
                  Begin MwfCtl.MWCTL_MED atxt 
                     Height          =   285
                     Index           =   94
                     Left            =   1605
                     TabIndex        =   156
                     ToolTipText     =   "Enter Licence Number"
                     Top             =   1710
                     Width           =   3105
                     _ExtentX        =   5477
                     _ExtentY        =   503
                     MaxLength       =   20
                     DataField       =   "InHLicNo"
                     IdName          =   "INHLICNO"
                  End
                  Begin MwfCtl.MWCTL_MED atxt 
                     Height          =   285
                     Index           =   95
                     Left            =   1605
                     TabIndex        =   155
                     ToolTipText     =   "Enter Origin Criteria"
                     Top             =   1425
                     Width           =   3105
                     _ExtentX        =   5477
                     _ExtentY        =   503
                     MaxLength       =   20
                     DataField       =   "InHGSPOrgCrit"
                     ReCalcOn        =   "INCMCD"
                     IdName          =   "INHGSPORGCRIT"
                  End
                  Begin MwfCtl.MWCTL_MED atxt 
                     Height          =   600
                     Index           =   96
                     Left            =   1605
                     TabIndex        =   152
                     ToolTipText     =   "Enter Statistical Code"
                     Top             =   270
                     Width           =   4935
                     _ExtentX        =   8705
                     _ExtentY        =   1058
                     MaxLength       =   60
                     DataField       =   "InHStatCd"
                     IdName          =   "INHSTATCD"
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Q/CertNo && Dt"
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
                     Left            =   90
                     TabIndex        =   288
                     Top             =   1140
                     Width           =   1545
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "AR4No && Dt"
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
                     Index           =   78
                     Left            =   90
                     TabIndex        =   287
                     Top             =   855
                     Width           =   1395
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Lic Number"
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
                     Left            =   90
                     TabIndex        =   286
                     Top             =   1710
                     Width           =   1275
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "ETC Saving"
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
                     Index           =   80
                     Left            =   90
                     TabIndex        =   285
                     Top             =   1995
                     Width           =   1275
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Origin Criteria"
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
                     Index           =   81
                     Left            =   90
                     TabIndex        =   284
                     Top             =   1425
                     Width           =   1485
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Statistical Cd"
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
                     Index           =   82
                     Left            =   90
                     TabIndex        =   283
                     Top             =   270
                     Width           =   1485
                  End
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   51
                  Left            =   1455
                  TabIndex        =   141
                  ToolTipText     =   "Enter RBI Number"
                  Top             =   2145
                  Width           =   3105
                  _ExtentX        =   5477
                  _ExtentY        =   503
                  MaxLength       =   20
                  DataType        =   4
                  DataField       =   "InHRbiNo"
                  IdName          =   "INHRBINO"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   52
                  Left            =   1455
                  TabIndex        =   143
                  ToolTipText     =   "Enter Import Export Code Number"
                  Top             =   2430
                  Width           =   3105
                  _ExtentX        =   5477
                  _ExtentY        =   503
                  MaxLength       =   20
                  DataType        =   4
                  DataField       =   "InHIecNo"
                  ReCalcParent    =   "INLCLCSTRS"
                  IdName          =   "INHIECNO"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   53
                  Left            =   1455
                  TabIndex        =   127
                  ToolTipText     =   "Enter EDF Number"
                  Top             =   150
                  Width           =   3105
                  _ExtentX        =   5477
                  _ExtentY        =   503
                  MaxLength       =   20
                  DataType        =   4
                  DataField       =   "InGrNo"
                  IdName          =   "INGRNO"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   54
                  Left            =   6330
                  TabIndex        =   128
                  ToolTipText     =   "Enter GR Date"
                  Top             =   150
                  Width           =   885
                  _ExtentX        =   1561
                  _ExtentY        =   503
                  MaxLength       =   8
                  DataType        =   3
                  DataField       =   "InGrDt"
                  IdName          =   "INGRDT"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   55
                  Left            =   1455
                  TabIndex        =   129
                  ToolTipText     =   "Enter Shipping Bill Number"
                  Top             =   435
                  Width           =   3105
                  _ExtentX        =   5477
                  _ExtentY        =   503
                  MaxLength       =   20
                  DataType        =   4
                  DataField       =   "InSbNo"
                  IdName          =   "INSBNO"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   56
                  Left            =   6330
                  TabIndex        =   130
                  ToolTipText     =   "Enter Shipping Bill Date"
                  Top             =   435
                  Width           =   885
                  _ExtentX        =   1561
                  _ExtentY        =   503
                  MaxLength       =   8
                  DataType        =   3
                  DataField       =   "InSbDt"
                  IdName          =   "INSBDT"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   57
                  Left            =   1455
                  TabIndex        =   131
                  ToolTipText     =   "Enter Master Air-Way Bill Number"
                  Top             =   720
                  Width           =   3105
                  _ExtentX        =   5477
                  _ExtentY        =   503
                  MaxLength       =   20
                  DataType        =   4
                  DataField       =   "InAwbNo"
                  IdName          =   "INAWBNO"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   58
                  Left            =   6330
                  TabIndex        =   132
                  ToolTipText     =   "Enter Master Air-Way Bill Date"
                  Top             =   720
                  Width           =   885
                  _ExtentX        =   1561
                  _ExtentY        =   503
                  MaxLength       =   8
                  DataType        =   3
                  DataField       =   "InAwbDt"
                  IdName          =   "INAWBDT"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   59
                  Left            =   1455
                  TabIndex        =   135
                  ToolTipText     =   "Enter GSP Form Number"
                  Top             =   1290
                  Width           =   3105
                  _ExtentX        =   5477
                  _ExtentY        =   503
                  MaxLength       =   20
                  DataType        =   4
                  DataField       =   "InGspNo"
                  IdName          =   "INGSPNO"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   60
                  Left            =   6330
                  TabIndex        =   136
                  ToolTipText     =   "Enter GSP Form Date"
                  Top             =   1290
                  Width           =   885
                  _ExtentX        =   1561
                  _ExtentY        =   503
                  MaxLength       =   8
                  DataType        =   3
                  DataField       =   "InGspDt"
                  IdName          =   "INGSPDT"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   61
                  Left            =   1455
                  TabIndex        =   137
                  ToolTipText     =   "Enter Purchase Order Number"
                  Top             =   1575
                  Width           =   3105
                  _ExtentX        =   5477
                  _ExtentY        =   503
                  MaxLength       =   20
                  DataType        =   4
                  DataField       =   "InPoNo"
                  IdName          =   "INPONO"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   62
                  Left            =   6330
                  TabIndex        =   138
                  ToolTipText     =   "Enter Purchase Order Date"
                  Top             =   1575
                  Width           =   885
                  _ExtentX        =   1561
                  _ExtentY        =   503
                  MaxLength       =   8
                  DataType        =   3
                  DataField       =   "InPoDt"
                  IdName          =   "INPODT"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   63
                  Left            =   1455
                  TabIndex        =   139
                  ToolTipText     =   "Enter Cheque Number"
                  Top             =   1860
                  Width           =   3105
                  _ExtentX        =   5477
                  _ExtentY        =   503
                  MaxLength       =   20
                  DataType        =   4
                  DataField       =   "InChqNo"
                  IdName          =   "INCHQNO"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   64
                  Left            =   6330
                  TabIndex        =   140
                  ToolTipText     =   "Enter Cheque Amount"
                  Top             =   1860
                  Width           =   885
                  _ExtentX        =   1561
                  _ExtentY        =   503
                  Alignment       =   2
                  Mask            =   "#####.00"
                  MaxLength       =   8
                  DataType        =   2
                  DataField       =   "InChqAmt"
                  IdName          =   "INCHQAMT"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   14
                  Left            =   1455
                  TabIndex        =   144
                  ToolTipText     =   "Enter ITC Number"
                  Top             =   2715
                  Width           =   3105
                  _ExtentX        =   5477
                  _ExtentY        =   503
                  MaxLength       =   20
                  DataType        =   4
                  DataField       =   "InHItcNo"
                  IdName          =   "INHITCNO"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   82
                  Left            =   1455
                  TabIndex        =   133
                  ToolTipText     =   "Enter House Air-Way Bill Number"
                  Top             =   1005
                  Width           =   3105
                  _ExtentX        =   5477
                  _ExtentY        =   503
                  MaxLength       =   20
                  DataType        =   4
                  DataField       =   "InHAwbNo"
                  IdName          =   "INHAWBNO"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   88
                  Left            =   6330
                  TabIndex        =   134
                  ToolTipText     =   "Enter House Air-Way Bill Date"
                  Top             =   1005
                  Width           =   885
                  _ExtentX        =   1561
                  _ExtentY        =   503
                  MaxLength       =   8
                  DataType        =   3
                  DataField       =   "InHAwbDt"
                  IdName          =   "INHAWBDT"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   128
                  Left            =   6330
                  TabIndex        =   142
                  ToolTipText     =   "Enter Authorised Signatory"
                  Top             =   2145
                  Width           =   4215
                  _ExtentX        =   7435
                  _ExtentY        =   503
                  MaxLength       =   30
                  DataField       =   "InhAuthSgn"
                  IdName          =   "INHAUTHSGN"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   135
                  Left            =   6330
                  TabIndex        =   145
                  ToolTipText     =   "Enter Shipping Bill Signed By"
                  Top             =   2640
                  Width           =   4725
                  _ExtentX        =   8334
                  _ExtentY        =   503
                  MaxLength       =   45
                  DataField       =   "InHShpBillSgnBy"
                  IdName          =   "INHSHPBILLSGNBY"
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Shp Bill Signed By"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   525
                  Index           =   114
                  Left            =   5160
                  TabIndex        =   337
                  Top             =   2640
                  Width           =   1215
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "IEC No."
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
                  Left            =   90
                  TabIndex        =   311
                  Top             =   2430
                  Width           =   1635
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "RBI No."
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
                  Left            =   90
                  TabIndex        =   310
                  Top             =   2145
                  Width           =   1725
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "EDF No"
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
                  Left            =   90
                  TabIndex        =   309
                  Top             =   150
                  Width           =   1395
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Date"
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
                  Left            =   5160
                  TabIndex        =   308
                  Top             =   150
                  Width           =   555
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "S/B No"
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
                  Left            =   90
                  TabIndex        =   307
                  Top             =   435
                  Width           =   1725
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Date"
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
                  Left            =   5160
                  TabIndex        =   306
                  Top             =   435
                  Width           =   555
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "MAWB No"
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
                  Left            =   90
                  TabIndex        =   305
                  Top             =   720
                  Width           =   1395
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Date"
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
                  Left            =   5160
                  TabIndex        =   304
                  Top             =   720
                  Width           =   555
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "GSP No"
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
                  Left            =   90
                  TabIndex        =   303
                  Top             =   1290
                  Width           =   1725
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Date"
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
                  Left            =   5160
                  TabIndex        =   302
                  Top             =   1290
                  Width           =   555
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Date"
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
                  Left            =   5160
                  TabIndex        =   301
                  Top             =   1575
                  Width           =   555
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
                  Height          =   285
                  Index           =   56
                  Left            =   90
                  TabIndex        =   300
                  Top             =   1575
                  Width           =   1725
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Chq No"
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
                  Left            =   90
                  TabIndex        =   299
                  Top             =   1860
                  Width           =   1725
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Chq Amt"
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
                  Left            =   5160
                  TabIndex        =   298
                  Top             =   1860
                  Width           =   945
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "ITC No."
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
                  Left            =   90
                  TabIndex        =   297
                  Top             =   2715
                  Width           =   1365
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Date"
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
                  Left            =   5160
                  TabIndex        =   296
                  Top             =   1005
                  Width           =   555
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "HAWB No"
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
                  Left            =   90
                  TabIndex        =   295
                  Top             =   1005
                  Width           =   1425
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Authorised Signatory"
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
                  Index           =   107
                  Left            =   5160
                  TabIndex        =   294
                  Top             =   2145
                  Width           =   1095
               End
            End
            Begin VB.Frame FraTabInvDet 
               BorderStyle     =   0  'None
               Height          =   7580
               Index           =   2
               Left            =   -74940
               TabIndex        =   223
               Top             =   360
               Width           =   14655
               Begin VB.Frame FraCal1 
                  Height          =   2565
                  Left            =   60
                  TabIndex        =   384
                  Top             =   2130
                  Width           =   14550
                  Begin MwfCtl.MWCTL_MED atxt 
                     Height          =   285
                     Index           =   16
                     Left            =   1320
                     TabIndex        =   92
                     ToolTipText     =   "Enter Commission (In US Dollars)"
                     Top             =   705
                     Width           =   1155
                     _ExtentX        =   2037
                     _ExtentY        =   503
                     Alignment       =   2
                     Mask            =   "####0.00"
                     MaxLength       =   8
                     DataType        =   2
                     DataField       =   "InComm"
                     ReCalcParent    =   "INCOMMRS,INTOT,WRUNNINGINTOT"
                     ReCalcOn        =   "INLCLINVYN"
                     IdName          =   "INCOMM"
                  End
                  Begin MwfCtl.MWCTL_MED atxt 
                     Height          =   285
                     Index           =   71
                     Left            =   3120
                     TabIndex        =   93
                     ToolTipText     =   "Enter Commission (In Rupees)"
                     Top             =   705
                     Width           =   1395
                     _ExtentX        =   2461
                     _ExtentY        =   503
                     Alignment       =   2
                     Mask            =   "#######0.00"
                     MaxLength       =   11
                     DataType        =   2
                     DataField       =   "InCommRs"
                     ReCalcParent    =   "WRUNNINGINTOTRS"
                     ReCalcOn        =   "INCNVRT,INCOMM,INRSMULDIV,INLCLINVYN"
                     IdName          =   "INCOMMRS"
                  End
                  Begin MwfCtl.MWCTL_MED atxt 
                     Height          =   285
                     Index           =   72
                     Left            =   1320
                     TabIndex        =   95
                     ToolTipText     =   "Enter Freight (In US Dollars)"
                     Top             =   990
                     Width           =   1155
                     _ExtentX        =   2037
                     _ExtentY        =   503
                     Alignment       =   2
                     Mask            =   "####0.00"
                     MaxLength       =   8
                     DataType        =   2
                     DataField       =   "InFrt"
                     ReCalcParent    =   "INFRTRS,INTOTAS,INTOT,WRUNNINGINTOT"
                     ReCalcOn        =   "INLCLINVYN"
                     IdName          =   "INFRT"
                  End
                  Begin MwfCtl.MWCTL_MED atxt 
                     Height          =   285
                     Index           =   73
                     Left            =   3120
                     TabIndex        =   96
                     ToolTipText     =   "Enter Freight (In Rupees)"
                     Top             =   990
                     Width           =   1395
                     _ExtentX        =   2461
                     _ExtentY        =   503
                     Alignment       =   2
                     Mask            =   "#######0.00"
                     MaxLength       =   11
                     DataType        =   2
                     DataField       =   "InFrtRs"
                     ReCalcParent    =   "INTOTAS,WRUNNINGINTOTRS"
                     ReCalcOn        =   "INCNVRT,INFRT,INRSMULDIV,INLCLINVYN"
                     IdName          =   "INFRTRS"
                  End
                  Begin MwfCtl.MWCTL_MED atxt 
                     Height          =   285
                     Index           =   74
                     Left            =   1320
                     TabIndex        =   97
                     ToolTipText     =   "Enter Insurance (In US Dollars)"
                     Top             =   1275
                     Width           =   1155
                     _ExtentX        =   2037
                     _ExtentY        =   503
                     Alignment       =   2
                     Mask            =   "####0.00"
                     MaxLength       =   8
                     DataType        =   2
                     DataField       =   "InIns"
                     ReCalcParent    =   "ININSRS,INTOTAS,INTOT,WRUNNINGINTOT"
                     ReCalcOn        =   "INLCLINVYN"
                     IdName          =   "ININS"
                  End
                  Begin MwfCtl.MWCTL_MED atxt 
                     Height          =   285
                     Index           =   75
                     Left            =   3120
                     TabIndex        =   98
                     ToolTipText     =   "Enter Insurance (In Rupees)"
                     Top             =   1275
                     Width           =   1395
                     _ExtentX        =   2461
                     _ExtentY        =   503
                     Alignment       =   2
                     Mask            =   "#######0.00"
                     MaxLength       =   11
                     DataType        =   2
                     DataField       =   "InInsRs"
                     ReCalcParent    =   "INTOTAS,WRUNNINGINTOTRS"
                     ReCalcOn        =   "INCNVRT,ININS,INRSMULDIV,INLCLINVYN"
                     IdName          =   "ININSRS"
                  End
                  Begin MwfCtl.MWCTL_MED atxt 
                     Height          =   285
                     Index           =   175
                     Left            =   1320
                     TabIndex        =   99
                     ToolTipText     =   "Enter Discount (In US Dollars)"
                     Top             =   1560
                     Width           =   1155
                     _ExtentX        =   2037
                     _ExtentY        =   503
                     Alignment       =   2
                     Mask            =   "-####0.00"
                     MaxLength       =   9
                     DataType        =   2
                     DataField       =   "InDsc"
                     ReCalcParent    =   "INDSCRS,WRUNNINGINTOT"
                     ReCalcOn        =   "INLCLINVYN,INHSELFDSCYN"
                     IdName          =   "INDSC"
                  End
                  Begin MwfCtl.MWCTL_MED atxt 
                     Height          =   285
                     Index           =   176
                     Left            =   3120
                     TabIndex        =   100
                     ToolTipText     =   "Enter Discount (In Rupees)"
                     Top             =   1560
                     Width           =   1395
                     _ExtentX        =   2461
                     _ExtentY        =   503
                     Alignment       =   2
                     Mask            =   "-#######0.00"
                     MaxLength       =   12
                     DataType        =   2
                     DataField       =   "InDscRs"
                     ReCalcParent    =   "WRUNNINGINTOTRS"
                     ReCalcOn        =   "INLCLINVYN,INDSC,INCNVRT,INRSMULDIV,INHSELFDSCYN"
                     IdName          =   "INDSCRS"
                  End
                  Begin MwfCtl.MWCTL_MED atxt 
                     Height          =   285
                     Index           =   177
                     Left            =   6030
                     TabIndex        =   102
                     ToolTipText     =   "Enter Discount Description"
                     Top             =   1560
                     Width           =   5505
                     _ExtentX        =   9710
                     _ExtentY        =   503
                     MaxLength       =   45
                     DataField       =   "InDscDesc"
                     IdName          =   "INDSCDESC"
                  End
                  Begin MwfCtl.MWCTL_MED atxt 
                     Height          =   285
                     Index           =   189
                     Left            =   4875
                     TabIndex        =   94
                     ToolTipText     =   "Specify Whether Commission is Self Payable (Y/N)"
                     Top             =   705
                     Width           =   300
                     _ExtentX        =   529
                     _ExtentY        =   503
                     DataType        =   4
                     DataField       =   "InHHideCommYN"
                     ReCalcParent    =   "WRUNNINGINTOT,WRUNNINGINTOTRS"
                     IdName          =   "INHHIDECOMMYN"
                  End
                  Begin MwfCtl.MWCTL_MED atxt 
                     Height          =   285
                     Index           =   194
                     Left            =   1320
                     TabIndex        =   103
                     ToolTipText     =   "Enter Deduction (In US Dollars)"
                     Top             =   1845
                     Width           =   1155
                     _ExtentX        =   2037
                     _ExtentY        =   503
                     Alignment       =   2
                     Mask            =   "-####0.00"
                     MaxLength       =   9
                     DataType        =   2
                     DataField       =   "InDeduct"
                     ReCalcParent    =   "INDEDUCTRS,WRUNNINGINTOT"
                     ReCalcOn        =   "INLCLINVYN,INHSELFDEDUCTYN"
                     IdName          =   "INDEDUCT"
                  End
                  Begin MwfCtl.MWCTL_MED atxt 
                     Height          =   285
                     Index           =   195
                     Left            =   3120
                     TabIndex        =   104
                     ToolTipText     =   "Enter Deduction (In Rupees)"
                     Top             =   1845
                     Width           =   1395
                     _ExtentX        =   2461
                     _ExtentY        =   503
                     Alignment       =   2
                     Mask            =   "-#######0.00"
                     MaxLength       =   12
                     DataType        =   2
                     DataField       =   "InDeductRs"
                     ReCalcParent    =   "WRUNNINGINTOTRS"
                     ReCalcOn        =   "INLCLINVYN,INDEDUCT,INCNVRT,INRSMULDIV,INHSELFDEDUCTYN"
                     IdName          =   "INDEDUCTRS"
                  End
                  Begin MwfCtl.MWCTL_MED atxt 
                     Height          =   285
                     Index           =   196
                     Left            =   6030
                     TabIndex        =   106
                     ToolTipText     =   "Enter Deduction Description"
                     Top             =   1845
                     Width           =   5505
                     _ExtentX        =   9710
                     _ExtentY        =   503
                     MaxLength       =   45
                     DataField       =   "InDeductDesc"
                     IdName          =   "INDEDUCTDESC"
                  End
                  Begin MwfCtl.MWCTL_MED atxt 
                     Height          =   285
                     Index           =   197
                     Left            =   4875
                     TabIndex        =   101
                     ToolTipText     =   "Specify Whether Discount is Self Payable (Y/N)"
                     Top             =   1560
                     Width           =   300
                     _ExtentX        =   529
                     _ExtentY        =   503
                     DataType        =   4
                     DataField       =   "InHSelfDscYN"
                     ReCalcParent    =   "INDSC,INDSCRS,WRUNNINGINTOT,WRUNNINGINTOTRS"
                     IdName          =   "INHSELFDSCYN"
                  End
                  Begin MwfCtl.MWCTL_MED atxt 
                     Height          =   285
                     Index           =   198
                     Left            =   4875
                     TabIndex        =   105
                     ToolTipText     =   "Specify Whether Deduction is Self Payable (Y/N)"
                     Top             =   1845
                     Width           =   300
                     _ExtentX        =   529
                     _ExtentY        =   503
                     DataType        =   4
                     DataField       =   "InHSelfDeductYN"
                     ReCalcParent    =   "INDEDUCT,INDEDUCTRS,WRUNNINGINTOT,WRUNNINGINTOTRS"
                     IdName          =   "INHSELFDEDUCTYN"
                  End
                  Begin MwfCtl.MWCTL_MED atxt 
                     Height          =   285
                     Index           =   207
                     Left            =   3720
                     TabIndex        =   91
                     ToolTipText     =   "Enter Total As Text"
                     Top             =   420
                     Width           =   570
                     _ExtentX        =   1005
                     _ExtentY        =   503
                     MaxLength       =   3
                     DataType        =   4
                     DataField       =   "INTOTASTXT"
                     ReCalcOn        =   "INTOTAS"
                     IdName          =   "INTOTASTXT"
                  End
                  Begin MwfCtl.MWCTL_MED atxt 
                     Height          =   285
                     Index           =   23
                     Left            =   1320
                     TabIndex        =   90
                     ToolTipText     =   "Enter Total As"
                     Top             =   420
                     Width           =   570
                     _ExtentX        =   1005
                     _ExtentY        =   503
                     MaxLength       =   3
                     DataType        =   4
                     DataField       =   "INTOTAS"
                     ReCalcParent    =   "INTOTASTXT,WRUNNINGINTOT,WRUNNINGINTOTRS"
                     ReCalcOn        =   "INFRT,INFRTRS,ININS,ININSRS"
                     IdName          =   "INTOTAS"
                  End
                  Begin MwfCtl.MWCTL_MED atxt 
                     Height          =   285
                     Index           =   254
                     Left            =   1320
                     TabIndex        =   107
                     ToolTipText     =   "Enter TCS (In US Dollars)"
                     Top             =   2130
                     Width           =   1155
                     _ExtentX        =   2037
                     _ExtentY        =   503
                     Alignment       =   2
                     Mask            =   "####0.00"
                     MaxLength       =   8
                     DataType        =   2
                     DataField       =   "InTCS"
                     ReCalcParent    =   "INDEDUCTRS,INTCSRS,INTOT,INTOT"
                     ReCalcOn        =   "WRUNNINGINTOT,INTCSPER"
                     IdName          =   "INTCS"
                  End
                  Begin MwfCtl.MWCTL_MED atxt 
                     Height          =   285
                     Index           =   255
                     Left            =   3120
                     TabIndex        =   108
                     ToolTipText     =   "Enter TCS (In Rupees)"
                     Top             =   2130
                     Width           =   1395
                     _ExtentX        =   2461
                     _ExtentY        =   503
                     Alignment       =   2
                     Mask            =   "#######0.00"
                     MaxLength       =   11
                     DataType        =   2
                     DataField       =   "InTCSRs"
                     ReCalcParent    =   "INTOTRS"
                     ReCalcOn        =   "WRUNNINGINTOTRS,INTCSPER"
                     IdName          =   "INTCSRS"
                  End
                  Begin MwfCtl.MWCTL_MED atxt 
                     Height          =   285
                     Index           =   256
                     Left            =   510
                     TabIndex        =   572
                     ToolTipText     =   "Enter TCS Percentage"
                     Top             =   2130
                     Width           =   765
                     _ExtentX        =   1349
                     _ExtentY        =   503
                     Alignment       =   2
                     Enabled         =   0   'False
                     Mask            =   "##0.000"
                     MaxLength       =   7
                     DataType        =   2
                     DataField       =   "InTcsPer"
                     ReCalcParent    =   "INLCLOCTROIRS,INTCS,INTCSRS"
                     ReCalcOn        =   "INLCLINVYN,INCMCD"
                     IdName          =   "INTCSPER"
                  End
                  Begin MwfCtl.MWCTL_MED atxt 
                     Height          =   285
                     Index           =   257
                     Left            =   7200
                     TabIndex        =   574
                     ToolTipText     =   "Enter Total Amount (In US Dollars)"
                     Top             =   2130
                     Visible         =   0   'False
                     Width           =   1155
                     _ExtentX        =   2037
                     _ExtentY        =   503
                     Alignment       =   2
                     Enabled         =   0   'False
                     Mask            =   "#######0.00"
                     MaxLength       =   11
                     DataType        =   2
                     ReCalcParent    =   "INTOTRS,INTCSRS,INTCS,INTOT,INTOT"
                     ReCalcOn        =   $"EmrFrmInvHd.frx":0490
                     IdName          =   "WRUNNINGINTOT"
                  End
                  Begin MwfCtl.MWCTL_MED atxt 
                     Height          =   285
                     Index           =   258
                     Left            =   8970
                     TabIndex        =   575
                     ToolTipText     =   "Enter Total Amount (In Rupees)"
                     Top             =   2130
                     Visible         =   0   'False
                     Width           =   1395
                     _ExtentX        =   2461
                     _ExtentY        =   503
                     Alignment       =   2
                     Enabled         =   0   'False
                     Mask            =   "########0.00"
                     MaxLength       =   12
                     DataType        =   2
                     ReCalcParent    =   "INTOTRS,INTCSRS"
                     ReCalcOn        =   "INCOMMRS,INFRTRS,ININSRS,INFOBRS,INTOTAS,INDSCRS,INHHIDECOMMYN,INDEDUCTRS,INHSELFDSCYN,INHSELFDEDUCTYN"
                     IdName          =   "WRUNNINGINTOTRS"
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "TCS "
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
                     Left            =   30
                     TabIndex        =   571
                     Top             =   2145
                     Width           =   855
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Total Val Text"
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
                     Index           =   163
                     Left            =   2280
                     TabIndex        =   450
                     Top             =   420
                     Width           =   1455
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Description"
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
                     Index           =   157
                     Left            =   6030
                     TabIndex        =   437
                     Top             =   450
                     Width           =   2805
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Self Payable"
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
                     Index           =   156
                     Left            =   4470
                     TabIndex        =   436
                     Top             =   450
                     Width           =   1305
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Deduction"
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
                     Index           =   155
                     Left            =   30
                     TabIndex        =   435
                     Top             =   1860
                     Width           =   1335
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Discount"
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
                     Index           =   146
                     Left            =   30
                     TabIndex        =   407
                     Top             =   1560
                     Width           =   1335
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Total Val As"
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
                     Left            =   30
                     TabIndex        =   401
                     Top             =   420
                     Width           =   1215
                  End
                  Begin VB.Label LblMain 
                     BackStyle       =   0  'Transparent
                     BorderStyle     =   1  'Fixed Single
                     Caption         =   "  Export Invoice"
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
                     Index           =   8
                     Left            =   0
                     TabIndex        =   399
                     Top             =   90
                     Width           =   14685
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Insurance"
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
                     Left            =   30
                     TabIndex        =   387
                     Top             =   1275
                     Width           =   1335
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Freight"
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
                     Left            =   30
                     TabIndex        =   386
                     Top             =   990
                     Width           =   1335
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Commission"
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
                     Left            =   30
                     TabIndex        =   385
                     Top             =   705
                     Width           =   1335
                  End
               End
               Begin VB.Frame Frame2 
                  Height          =   2085
                  Left            =   60
                  TabIndex        =   271
                  Top             =   5430
                  Width           =   11640
                  Begin MwfCtl.MWCTL_MED atxt 
                     Height          =   285
                     Index           =   24
                     Left            =   1305
                     TabIndex        =   121
                     ToolTipText     =   "Enter Diamond Cost"
                     Top             =   210
                     Width           =   1155
                     _ExtentX        =   2037
                     _ExtentY        =   503
                     Alignment       =   2
                     Mask            =   "#####0.00"
                     MaxLength       =   9
                     DataType        =   2
                     DataField       =   "InDiaCost"
                     IdName          =   "INDIACOST"
                  End
                  Begin MwfCtl.MWCTL_MED atxt 
                     Height          =   570
                     Index           =   78
                     Left            =   2475
                     TabIndex        =   122
                     ToolTipText     =   "Enter Diamond Cost Description"
                     Top             =   210
                     Width           =   8655
                     _ExtentX        =   15266
                     _ExtentY        =   1005
                     MaxLength       =   150
                     DataField       =   "InDiaCostDesc"
                     IdName          =   "INDIACOSTDESC"
                  End
                  Begin MwfCtl.MWCTL_MED atxt 
                     Height          =   285
                     Index           =   83
                     Left            =   1305
                     TabIndex        =   123
                     ToolTipText     =   "Enter Colour Stone Cost"
                     Top             =   780
                     Width           =   1155
                     _ExtentX        =   2037
                     _ExtentY        =   503
                     Alignment       =   2
                     Mask            =   "#####0.00"
                     MaxLength       =   9
                     DataType        =   2
                     DataField       =   "InCSCost"
                     IdName          =   "INCSCOST"
                  End
                  Begin MwfCtl.MWCTL_MED atxt 
                     Height          =   570
                     Index           =   84
                     Left            =   2475
                     TabIndex        =   124
                     ToolTipText     =   "Enter Colour Stone Cost Description"
                     Top             =   780
                     Width           =   8655
                     _ExtentX        =   15266
                     _ExtentY        =   1005
                     MaxLength       =   150
                     DataField       =   "InCSCostDesc"
                     IdName          =   "INCSCOSTDESC"
                  End
                  Begin MwfCtl.MWCTL_MED atxt 
                     Height          =   285
                     Index           =   85
                     Left            =   1305
                     TabIndex        =   125
                     ToolTipText     =   "Enter Advance Remittance"
                     Top             =   1350
                     Width           =   1155
                     _ExtentX        =   2037
                     _ExtentY        =   503
                     Alignment       =   2
                     Mask            =   "#####0.00"
                     MaxLength       =   9
                     DataType        =   2
                     DataField       =   "InAdv"
                     IdName          =   "INADV"
                  End
                  Begin MwfCtl.MWCTL_MED atxt 
                     Height          =   570
                     Index           =   86
                     Left            =   2475
                     TabIndex        =   126
                     ToolTipText     =   "Enter Advance Remittance Description"
                     Top             =   1350
                     Width           =   8655
                     _ExtentX        =   15266
                     _ExtentY        =   1005
                     MaxLength       =   150
                     DataField       =   "InAdvDesc"
                     IdName          =   "INADVDESC"
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Dia Cost"
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
                     Left            =   90
                     TabIndex        =   274
                     Top             =   210
                     Width           =   1335
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "ColStn Cost"
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
                     Left            =   90
                     TabIndex        =   273
                     Top             =   780
                     Width           =   1335
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Adv Remit"
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
                     Left            =   90
                     TabIndex        =   272
                     Top             =   1350
                     Width           =   1335
                  End
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   0
                  Left            =   1905
                  TabIndex        =   73
                  ToolTipText     =   "Enter Currency Conversion Rate"
                  Top             =   435
                  Width           =   1065
                  _ExtentX        =   1879
                  _ExtentY        =   503
                  Alignment       =   2
                  Mask            =   "###0.0000"
                  MaxLength       =   9
                  DataType        =   2
                  DataField       =   "InCnvRt"
                  ReCalcParent    =   "INFOBRS,INCOMMRS,INFRTRS,ININSRS,INDSCRS,INDEDUCTRS,INTCSRS"
                  ReCalcOn        =   "INDT,INCMCD"
                  IdName          =   "INCNVRT"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   4
                  Left            =   1365
                  TabIndex        =   78
                  ToolTipText     =   "FOB Value (In US Dollars)"
                  Top             =   1425
                  Width           =   1155
                  _ExtentX        =   2037
                  _ExtentY        =   503
                  Alignment       =   2
                  Enabled         =   0   'False
                  Mask            =   "#######0.00"
                  MaxLength       =   11
                  DataType        =   2
                  DataField       =   "InFOB"
                  ReCalcParent    =   "INFOBRS,INLCLOCTROIRS,INLCLCSTRS,INLCLVATRS,INTOT,WRUNNINGINTOT"
                  IdName          =   "INFOB"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   5
                  Left            =   3165
                  TabIndex        =   79
                  ToolTipText     =   "FOB Value (In Rupees)"
                  Top             =   1425
                  Width           =   1395
                  _ExtentX        =   2461
                  _ExtentY        =   503
                  Alignment       =   2
                  Enabled         =   0   'False
                  Mask            =   "########0.00"
                  MaxLength       =   12
                  DataType        =   2
                  DataField       =   "InFOBRs"
                  ReCalcParent    =   "WRUNNINGINTOTRS"
                  ReCalcOn        =   "INCNVRT,INFOB,INRSMULDIV"
                  IdName          =   "INFOBRS"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   76
                  Left            =   1380
                  TabIndex        =   117
                  ToolTipText     =   "Enter Total Amount (In US Dollars)"
                  Top             =   4770
                  Width           =   1155
                  _ExtentX        =   2037
                  _ExtentY        =   503
                  Alignment       =   2
                  Enabled         =   0   'False
                  Mask            =   "#######0.00"
                  MaxLength       =   11
                  DataType        =   2
                  DataField       =   "InTot"
                  ReCalcOn        =   "WRUNNINGINTOT,INTCS"
                  IdName          =   "INTOT"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   77
                  Left            =   3150
                  TabIndex        =   118
                  ToolTipText     =   "Enter Total Amount (In Rupees)"
                  Top             =   4770
                  Width           =   1395
                  _ExtentX        =   2461
                  _ExtentY        =   503
                  Alignment       =   2
                  Enabled         =   0   'False
                  Mask            =   "########0.00"
                  MaxLength       =   12
                  DataType        =   2
                  DataField       =   "InTotRs"
                  ReCalcOn        =   "WRUNNINGINTOTRS,INTCSRS"
                  IdName          =   "INTOTRS"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   87
                  Left            =   3150
                  TabIndex        =   120
                  ToolTipText     =   "Enter Payment Terms"
                  Top             =   5070
                  Width           =   7035
                  _ExtentX        =   12409
                  _ExtentY        =   503
                  MaxLength       =   45
                  DataField       =   "InCmPayTerms"
                  ReCalcOn        =   "INCMCD"
                  IdName          =   "INCMPAYTERMS"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   129
                  Left            =   7785
                  TabIndex        =   75
                  ToolTipText     =   "Enter VA Comission Percentage"
                  Top             =   150
                  Width           =   765
                  _ExtentX        =   1349
                  _ExtentY        =   503
                  Alignment       =   2
                  Mask            =   "##0.00"
                  MaxLength       =   6
                  DataType        =   2
                  DataField       =   "InComPer"
                  ReCalcParent    =   "INFOBRS,INCOMMRS,INFRTRS,ININSRS"
                  IdName          =   "INCOMPER"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   127
                  Left            =   7785
                  TabIndex        =   76
                  ToolTipText     =   "Invoice/Custom  FOB Percentage"
                  Top             =   435
                  Width           =   765
                  _ExtentX        =   1349
                  _ExtentY        =   503
                  Alignment       =   2
                  Mask            =   "-#0.00"
                  MaxLength       =   6
                  DataType        =   2
                  DataField       =   "INFOBPER"
                  IdName          =   "INFOBPER"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   137
                  Left            =   7785
                  TabIndex        =   77
                  ToolTipText     =   "Enter Import Value Text For Reports"
                  Top             =   720
                  Width           =   3105
                  _ExtentX        =   5477
                  _ExtentY        =   503
                  MaxLength       =   30
                  DataField       =   "InHImpValTxt"
                  IdName          =   "INHIMPVALTXT"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   81
                  Left            =   7785
                  TabIndex        =   80
                  ToolTipText     =   "Specify Whether To Round Off The Inv & Custom FOB Values At The Dsg Level In Invoicing (Enter Y/N)"
                  Top             =   1005
                  Width           =   300
                  _ExtentX        =   529
                  _ExtentY        =   503
                  MaxLength       =   2
                  DataType        =   4
                  DataField       =   "InHRndOffFOB"
                  IdName          =   "INHRNDOFFFOB"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   150
                  Left            =   1905
                  TabIndex        =   71
                  ToolTipText     =   "Enter Customer Currency Conversion Factor"
                  Top             =   150
                  Width           =   1065
                  _ExtentX        =   1879
                  _ExtentY        =   503
                  Alignment       =   2
                  Mask            =   "###0.0000"
                  MaxLength       =   9
                  DataType        =   2
                  DataField       =   "InCnvFct"
                  ReCalcOn        =   "INCMCD,INDT"
                  IdName          =   "INCNVFCT"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   151
                  Left            =   4125
                  TabIndex        =   72
                  ToolTipText     =   "Multiplication/Division"
                  Top             =   150
                  Width           =   300
                  _ExtentX        =   529
                  _ExtentY        =   503
                  Alignment       =   2
                  Enabled         =   0   'False
                  DataType        =   4
                  DataField       =   "InMulDiv"
                  ReCalcOn        =   "INCMCD"
                  IdName          =   "INMULDIV"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   152
                  Left            =   4125
                  TabIndex        =   74
                  ToolTipText     =   "Multiplication/Division"
                  Top             =   435
                  Width           =   300
                  _ExtentX        =   529
                  _ExtentY        =   503
                  Alignment       =   2
                  Enabled         =   0   'False
                  DataType        =   4
                  DataField       =   "InRsMulDiv"
                  ReCalcParent    =   "INFOBRS,INCOMMRS,INFRTRS,ININSRS,INDSCRS,INDEDUCTRS,INTCSRS"
                  ReCalcOn        =   "INCMCD"
                  IdName          =   "INRSMULDIV"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   154
                  Left            =   7785
                  TabIndex        =   81
                  ToolTipText     =   "Specify Whether To Show The Gold Loss Percentage Field In The Invoice Reports (Enter Y/N)"
                  Top             =   1320
                  Width           =   300
                  _ExtentX        =   529
                  _ExtentY        =   503
                  DataType        =   4
                  DataField       =   "InShowGldLs"
                  IdName          =   "INSHOWGLDLS"
               End
               Begin VB.Frame FraCal2 
                  Height          =   2115
                  Left            =   60
                  TabIndex        =   388
                  Top             =   2250
                  Width           =   5265
                  Begin MwfCtl.MWCTL_MED atxt 
                     Height          =   285
                     Index           =   163
                     Left            =   1320
                     TabIndex        =   109
                     ToolTipText     =   "Enter Octroi Percentage"
                     Top             =   480
                     Width           =   765
                     _ExtentX        =   1349
                     _ExtentY        =   503
                     Alignment       =   2
                     Mask            =   "#0.00"
                     MaxLength       =   5
                     DataType        =   2
                     DataField       =   "InCmOctroiPer"
                     ReCalcParent    =   "INLCLOCTROIRS"
                     ReCalcOn        =   "INCMCD,INLCLINVYN"
                     IdName          =   "INCMOCTROIPER"
                  End
                  Begin MwfCtl.MWCTL_MED atxt 
                     Height          =   285
                     Index           =   164
                     Left            =   1320
                     TabIndex        =   110
                     ToolTipText     =   "Enter VAT Percentage"
                     Top             =   765
                     Width           =   765
                     _ExtentX        =   1349
                     _ExtentY        =   503
                     Alignment       =   2
                     Mask            =   "#0.00"
                     MaxLength       =   5
                     DataType        =   2
                     DataField       =   "InHVatPer"
                     ReCalcParent    =   "INLCLVATRS"
                     ReCalcOn        =   "INCMCD,INLCLINVYN"
                     IdName          =   "INHVATPER"
                  End
                  Begin MwfCtl.MWCTL_MED atxt 
                     Height          =   285
                     Index           =   165
                     Left            =   1320
                     TabIndex        =   111
                     ToolTipText     =   "Enter CST Percentage"
                     Top             =   1050
                     Width           =   765
                     _ExtentX        =   1349
                     _ExtentY        =   503
                     Alignment       =   2
                     Mask            =   "#0.00"
                     MaxLength       =   5
                     DataType        =   2
                     DataField       =   "InHCstPer"
                     ReCalcParent    =   "INLCLCSTRS"
                     ReCalcOn        =   "INCMCD,INLCLINVYN"
                     IdName          =   "INHCSTPER"
                  End
                  Begin MwfCtl.MWCTL_MED atxt 
                     Height          =   285
                     Index           =   166
                     Left            =   3810
                     TabIndex        =   112
                     ToolTipText     =   "Enter Octroi Value"
                     Top             =   480
                     Width           =   1275
                     _ExtentX        =   2249
                     _ExtentY        =   503
                     Alignment       =   2
                     Mask            =   "######0.00"
                     MaxLength       =   10
                     DataType        =   2
                     DataField       =   "InLclOctroiRs"
                     ReCalcParent    =   "INLCLCSTRS,INLCLVATRS,WRUNNINGINTOT"
                     ReCalcOn        =   "INCMOCTROIPER,INFOB"
                     IdName          =   "INLCLOCTROIRS"
                  End
                  Begin MwfCtl.MWCTL_MED atxt 
                     Height          =   285
                     Index           =   167
                     Left            =   3810
                     TabIndex        =   113
                     ToolTipText     =   "Enter VAT Value"
                     Top             =   765
                     Width           =   1275
                     _ExtentX        =   2249
                     _ExtentY        =   503
                     Alignment       =   2
                     Mask            =   "######0.00"
                     MaxLength       =   10
                     DataType        =   2
                     DataField       =   "InLclVatRs"
                     ReCalcParent    =   "WRUNNINGINTOT"
                     ReCalcOn        =   "INHVATPER,INLCLOCTROIRS,INFOB,INLCLOTHRVALRS"
                     IdName          =   "INLCLVATRS"
                  End
                  Begin MwfCtl.MWCTL_MED atxt 
                     Height          =   285
                     Index           =   168
                     Left            =   1320
                     TabIndex        =   116
                     ToolTipText     =   "Enter Other Value Description"
                     Top             =   1620
                     Width           =   3765
                     _ExtentX        =   6641
                     _ExtentY        =   503
                     MaxLength       =   20
                     DataField       =   "InLclOthrValDesc"
                     ReCalcOn        =   "INCMCD,INLCLINVYN"
                     IdName          =   "INLCLOTHRVALDESC"
                  End
                  Begin MwfCtl.MWCTL_MED atxt 
                     Height          =   285
                     Index           =   169
                     Left            =   3810
                     TabIndex        =   115
                     ToolTipText     =   "Enter Other Values"
                     Top             =   1335
                     Width           =   1275
                     _ExtentX        =   2249
                     _ExtentY        =   503
                     Alignment       =   2
                     Mask            =   "-#####0.00"
                     MaxLength       =   10
                     DataType        =   2
                     DataField       =   "InLclOthrValRs"
                     ReCalcParent    =   "INLCLCSTRS,INLCLVATRS,WRUNNINGINTOT"
                     ReCalcOn        =   "INCMCD,INLCLINVYN"
                     IdName          =   "INLCLOTHRVALRS"
                  End
                  Begin MwfCtl.MWCTL_MED atxt 
                     Height          =   285
                     Index           =   160
                     Left            =   3810
                     TabIndex        =   114
                     ToolTipText     =   "Enter CST Value"
                     Top             =   1050
                     Width           =   1275
                     _ExtentX        =   2249
                     _ExtentY        =   503
                     Alignment       =   2
                     Mask            =   "######0.00"
                     MaxLength       =   10
                     DataType        =   2
                     DataField       =   "InLclCstRs"
                     ReCalcParent    =   "WRUNNINGINTOT"
                     ReCalcOn        =   "INHCSTPER,INLCLOCTROIRS,INFOB,INHIECNO,INLCLOTHRVALRS"
                     IdName          =   "INLCLCSTRS"
                  End
                  Begin VB.Label LblMain 
                     BackStyle       =   0  'Transparent
                     BorderStyle     =   1  'Fixed Single
                     Caption         =   "  Local Invoice"
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
                     Index           =   9
                     Left            =   -30
                     TabIndex        =   400
                     Top             =   90
                     Width           =   5295
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "CST Val"
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
                     Index           =   128
                     Left            =   2370
                     TabIndex        =   396
                     Top             =   1050
                     Width           =   1305
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Other Values"
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
                     Index           =   142
                     Left            =   2370
                     TabIndex        =   395
                     Top             =   1320
                     Width           =   1305
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Othr Value Desc"
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
                     Index           =   141
                     Left            =   90
                     TabIndex        =   394
                     Top             =   1470
                     Width           =   1065
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "VAT Val"
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
                     Left            =   2370
                     TabIndex        =   393
                     Top             =   765
                     Width           =   1305
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Octroi Val"
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
                     Index           =   139
                     Left            =   2370
                     TabIndex        =   392
                     Top             =   480
                     Width           =   1305
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "CST %"
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
                     Index           =   138
                     Left            =   90
                     TabIndex        =   391
                     Top             =   1050
                     Width           =   915
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "VAT %"
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
                     Index           =   137
                     Left            =   90
                     TabIndex        =   390
                     Top             =   765
                     Width           =   915
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Octroi %"
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
                     Index           =   136
                     Left            =   90
                     TabIndex        =   389
                     Top             =   480
                     Width           =   915
                  End
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   182
                  Left            =   13695
                  TabIndex        =   83
                  ToolTipText     =   "Enter 'Y' If Invoice rate of Dia and Col Stones has to be Calc By Stock "
                  Top             =   30
                  Width           =   300
                  _ExtentX        =   529
                  _ExtentY        =   503
                  DataType        =   4
                  DataField       =   "INIRTWRTSTKYN"
                  IdName          =   "INIRTWRTSTKYN"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   190
                  Left            =   13695
                  TabIndex        =   84
                  ToolTipText     =   "Enter 'Y' If Invoice rate of Accessories has to be Calc By Stock "
                  Top             =   315
                  Width           =   300
                  _ExtentX        =   529
                  _ExtentY        =   503
                  DataType        =   4
                  DataField       =   "INXIRTWRTSTKYN"
                  IdName          =   "INXIRTWRTSTKYN"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   192
                  Left            =   13695
                  TabIndex        =   86
                  ToolTipText     =   "Enter 'Y' to Use Notional Rate In Custom Matching Instead of Purchase Rate"
                  Top             =   900
                  Width           =   300
                  _ExtentX        =   529
                  _ExtentY        =   503
                  DataType        =   4
                  DataField       =   "InNotionRtYN"
                  IdName          =   "INNOTIONRTYN"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   204
                  Left            =   13695
                  TabIndex        =   85
                  ToolTipText     =   "Enter 'Y' If Invoice rate of Accessories has to be Calc By Stock "
                  Top             =   600
                  Width           =   300
                  _ExtentX        =   529
                  _ExtentY        =   503
                  DataType        =   4
                  DataField       =   "InMetIRtwrtStkYN"
                  IdName          =   "INMETIRTWRTSTKYN"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   205
                  Left            =   13695
                  TabIndex        =   87
                  ToolTipText     =   "Enter 'Y' to Use RM Rate from Rate Chart"
                  Top             =   1200
                  Width           =   300
                  _ExtentX        =   529
                  _ExtentY        =   503
                  DataType        =   4
                  DataField       =   "InRmRtfrRtChrt"
                  IdName          =   "INRMRTFRRTCHRT"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   206
                  Left            =   13695
                  TabIndex        =   88
                  ToolTipText     =   "Enter 'Y' to Use Labor Rate from Rate Chart"
                  Top             =   1500
                  Width           =   300
                  _ExtentX        =   529
                  _ExtentY        =   503
                  DataType        =   4
                  DataField       =   "InLabRtfrRtChrt"
                  IdName          =   "INLABRTFRRTCHRT"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   208
                  Left            =   7785
                  TabIndex        =   82
                  ToolTipText     =   "Specify Whether To Apply Order Mulby Factor. Enter (Y/N)"
                  Top             =   1620
                  Width           =   300
                  _ExtentX        =   529
                  _ExtentY        =   503
                  DataType        =   4
                  DataField       =   "InOrdMulbyYN"
                  IdName          =   "INORDMULBYYN"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   211
                  Left            =   13695
                  TabIndex        =   89
                  ToolTipText     =   "Enter 'Y' to apply Rate from Invoice Customer"
                  Top             =   1800
                  Width           =   300
                  _ExtentX        =   529
                  _ExtentY        =   503
                  DataType        =   4
                  DataField       =   "InRtFrInvCust"
                  IdName          =   "INRTFRINVCUST"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   345
                  Left            =   1380
                  TabIndex        =   119
                  ToolTipText     =   "Enter Payment Term Code"
                  Top             =   5070
                  Width           =   1140
                  _ExtentX        =   2011
                  _ExtentY        =   503
                  MaxLength       =   5
                  DataType        =   4
                  IdName          =   "WPAYTERMSCD"
               End
               Begin VB.Label Label3 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Rate from Inv Cust"
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
                  Left            =   11130
                  TabIndex        =   453
                  Top             =   1800
                  Width           =   2460
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Apply Ord Mulby Factor"
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
                  Index           =   164
                  Left            =   5220
                  TabIndex        =   451
                  Top             =   1620
                  Width           =   2325
               End
               Begin VB.Label Label2 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Lab Rate from Rate Chart"
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
                  Left            =   11130
                  TabIndex        =   449
                  Top             =   1500
                  Width           =   2460
               End
               Begin VB.Label Label1 
                  BackStyle       =   0  'Transparent
                  Caption         =   "RM Rate from Rate Chart"
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
                  Left            =   11130
                  TabIndex        =   448
                  Top             =   1200
                  Width           =   2370
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Metal Inv Rt By Stk"
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
                  Index           =   162
                  Left            =   11130
                  TabIndex        =   447
                  Top             =   600
                  Width           =   2070
               End
               Begin VB.Label LblNotionRt 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Notional Rate"
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
                  Left            =   11130
                  TabIndex        =   420
                  Top             =   900
                  Width           =   2370
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Acc Inv Rt By Stk"
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
                  Index           =   153
                  Left            =   11130
                  TabIndex        =   418
                  Top             =   315
                  Width           =   1650
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "DC Inv Rt By Stk"
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
                  Index           =   147
                  Left            =   11130
                  TabIndex        =   417
                  Top             =   30
                  Width           =   1650
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Show Loss % In Reports"
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
                  Index           =   127
                  Left            =   5220
                  TabIndex        =   372
                  Top             =   1320
                  Width           =   2325
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Mul/Div"
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
                  Index           =   124
                  Left            =   3300
                  TabIndex        =   369
                  Top             =   435
                  Width           =   855
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Mul/Div"
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
                  Left            =   3300
                  TabIndex        =   368
                  Top             =   150
                  Width           =   855
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Cust Curr Conv Fct"
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
                  Left            =   90
                  TabIndex        =   367
                  Top             =   150
                  Width           =   1905
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Rnd FOB"
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
                  Index           =   67
                  Left            =   5220
                  TabIndex        =   364
                  Top             =   1020
                  Width           =   1035
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Import Val  Text"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   405
                  Index           =   116
                  Left            =   5220
                  TabIndex        =   339
                  Top             =   720
                  Width           =   1845
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "FOB Percentage"
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
                  Left            =   5220
                  TabIndex        =   332
                  Top             =   435
                  Width           =   2145
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Rupee Conv Rt"
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
                  Left            =   90
                  TabIndex        =   281
                  Top             =   435
                  Width           =   1815
               End
               Begin VB.Label LblLocalCurncyHd 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "in Ind Rs"
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
                  Left            =   3165
                  TabIndex        =   280
                  Top             =   1140
                  Width           =   1005
               End
               Begin VB.Label LblInUSD 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "in Cust Currency"
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
                  Left            =   1365
                  TabIndex        =   279
                  Top             =   1140
                  Width           =   1560
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "FOB"
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
                  TabIndex        =   278
                  Top             =   1425
                  Width           =   855
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Total Amt"
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
                  Index           =   66
                  Left            =   90
                  TabIndex        =   277
                  Top             =   4770
                  Width           =   1335
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Pay Terms"
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
                  TabIndex        =   276
                  Top             =   5100
                  Width           =   1245
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "VA Commision Percentage"
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
                  Index           =   108
                  Left            =   5205
                  TabIndex        =   275
                  Top             =   150
                  Width           =   2550
               End
            End
            Begin VB.Frame FraTabInvDet 
               BorderStyle     =   0  'None
               Height          =   7545
               Index           =   1
               Left            =   -74940
               TabIndex        =   221
               Top             =   390
               Width           =   14625
               Begin VB.Frame FraFrgnBnk 
                  Height          =   1245
                  Left            =   -30
                  TabIndex        =   257
                  Top             =   4530
                  Width           =   14760
                  Begin MwfCtl.MWCTL_MED atxt 
                     Height          =   285
                     Index           =   49
                     Left            =   1815
                     TabIndex        =   65
                     ToolTipText     =   "Enter Foreign Bank Name"
                     Top             =   420
                     Width           =   4725
                     _ExtentX        =   8334
                     _ExtentY        =   503
                     MaxLength       =   30
                     DataField       =   "InHFrgnBkName"
                     IdName          =   "INHFRGNBKNAME"
                  End
                  Begin MwfCtl.MWCTL_MED atxt 
                     Height          =   285
                     Index           =   50
                     Left            =   1815
                     TabIndex        =   66
                     ToolTipText     =   "Enter Foreign Bank Account"
                     Top             =   705
                     Width           =   3105
                     _ExtentX        =   5477
                     _ExtentY        =   503
                     MaxLength       =   20
                     DataType        =   4
                     DataField       =   "InHFrgnBkAc"
                     IdName          =   "INHFRGNBKAC"
                  End
                  Begin VB.Label LblMain 
                     BackStyle       =   0  'Transparent
                     BorderStyle     =   1  'Fixed Single
                     Caption         =   "  Foreign Bank"
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
                     Left            =   -30
                     TabIndex        =   347
                     Top             =   90
                     Width           =   15135
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Name"
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
                     Left            =   90
                     TabIndex        =   259
                     Top             =   420
                     Width           =   1065
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "A/c"
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
                     Left            =   90
                     TabIndex        =   258
                     Top             =   705
                     Width           =   1305
                  End
               End
               Begin VB.Frame FraBnk 
                  Height          =   2025
                  Left            =   -30
                  TabIndex        =   260
                  Top             =   2610
                  Width           =   14730
                  Begin MwfCtl.MWCTL_MED atxt 
                     Height          =   285
                     Index           =   44
                     Left            =   1815
                     TabIndex        =   61
                     ToolTipText     =   "Enter Bank Address"
                     Top             =   705
                     Width           =   7035
                     _ExtentX        =   12409
                     _ExtentY        =   503
                     MaxLength       =   45
                     DataField       =   "InHBkAdd1"
                     ReCalcOn        =   "WBKCD"
                     IdName          =   "INHBKADD1"
                  End
                  Begin MwfCtl.MWCTL_MED atxt 
                     Height          =   285
                     Index           =   45
                     Left            =   1815
                     TabIndex        =   62
                     ToolTipText     =   "Enter Bank Address"
                     Top             =   990
                     Width           =   7035
                     _ExtentX        =   12409
                     _ExtentY        =   503
                     MaxLength       =   45
                     DataField       =   "InHBkAdd2"
                     ReCalcOn        =   "WBKCD"
                     IdName          =   "INHBKADD2"
                  End
                  Begin MwfCtl.MWCTL_MED atxt 
                     Height          =   285
                     Index           =   46
                     Left            =   1815
                     TabIndex        =   63
                     ToolTipText     =   "Enter Bank Address"
                     Top             =   1275
                     Width           =   7035
                     _ExtentX        =   12409
                     _ExtentY        =   503
                     MaxLength       =   45
                     DataField       =   "InHBkAdd3"
                     ReCalcOn        =   "WBKCD"
                     IdName          =   "INHBKADD3"
                  End
                  Begin MwfCtl.MWCTL_MED atxt 
                     Height          =   285
                     Index           =   47
                     Left            =   3135
                     TabIndex        =   60
                     ToolTipText     =   "Enter Bank Name"
                     Top             =   420
                     Width           =   4725
                     _ExtentX        =   8334
                     _ExtentY        =   503
                     MaxLength       =   30
                     DataField       =   "InHBkName"
                     ReCalcOn        =   "WBKCD"
                     IdName          =   "INHBKNAME"
                  End
                  Begin MwfCtl.MWCTL_MED atxt 
                     Height          =   285
                     Index           =   48
                     Left            =   1815
                     TabIndex        =   64
                     ToolTipText     =   "Enter Bank Account"
                     Top             =   1560
                     Width           =   3105
                     _ExtentX        =   5477
                     _ExtentY        =   503
                     MaxLength       =   20
                     DataType        =   4
                     DataField       =   "InHBkAc"
                     ReCalcOn        =   "WBKCD"
                     IdName          =   "INHBKAC"
                  End
                  Begin MwfCtl.MWCTL_MED atxt 
                     Height          =   285
                     Index           =   79
                     Left            =   1815
                     TabIndex        =   59
                     ToolTipText     =   "Enter Bank Code"
                     Top             =   420
                     Width           =   1245
                     _ExtentX        =   2196
                     _ExtentY        =   503
                     MaxLength       =   8
                     DataType        =   4
                     ReCalcParent    =   "INHBKNAME,INHBKADD1,INHBKADD2,INHBKADD3,INHBNKINSTR,INHBKAC"
                     IdName          =   "WBKCD"
                  End
                  Begin VB.Label LblMain 
                     BackStyle       =   0  'Transparent
                     BorderStyle     =   1  'Fixed Single
                     Caption         =   "  Bank"
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
                     Index           =   5
                     Left            =   -30
                     TabIndex        =   348
                     Top             =   90
                     Width           =   15135
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "A/c"
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
                     Left            =   90
                     TabIndex        =   264
                     Top             =   1560
                     Width           =   675
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Address"
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
                     Left            =   90
                     TabIndex        =   263
                     Top             =   705
                     Width           =   1275
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Cd/ Name"
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
                     Left            =   90
                     TabIndex        =   262
                     Top             =   420
                     Width           =   1755
                  End
                  Begin VB.Label ALBL 
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
                     Height          =   285
                     Index           =   10
                     Left            =   3060
                     TabIndex        =   261
                     Top             =   420
                     Width           =   75
                  End
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   32
                  Left            =   1815
                  TabIndex        =   52
                  ToolTipText     =   "Enter Company Address"
                  Top             =   555
                  Width           =   7035
                  _ExtentX        =   12409
                  _ExtentY        =   503
                  MaxLength       =   45
                  DataField       =   "InSelfAdd1"
                  IdName          =   "INSELFADD1"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   33
                  Left            =   1815
                  TabIndex        =   53
                  ToolTipText     =   "Enter Company Address"
                  Top             =   840
                  Width           =   7035
                  _ExtentX        =   12409
                  _ExtentY        =   503
                  MaxLength       =   45
                  DataField       =   "InSelfAdd2"
                  IdName          =   "INSELFADD2"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   35
                  Left            =   1815
                  TabIndex        =   54
                  ToolTipText     =   "Enter Company Address"
                  Top             =   1125
                  Width           =   7035
                  _ExtentX        =   12409
                  _ExtentY        =   503
                  MaxLength       =   45
                  DataField       =   "InSelfAdd3"
                  IdName          =   "INSELFADD3"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   37
                  Left            =   1815
                  TabIndex        =   55
                  ToolTipText     =   "Enter Company Telephone No"
                  Top             =   1410
                  Width           =   3105
                  _ExtentX        =   5477
                  _ExtentY        =   503
                  MaxLength       =   20
                  DataField       =   "InSelfTel"
                  IdName          =   "INSELFTEL"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   38
                  Left            =   1815
                  TabIndex        =   56
                  ToolTipText     =   "Enter Company GST No"
                  Top             =   1695
                  Width           =   3105
                  _ExtentX        =   5477
                  _ExtentY        =   503
                  MaxLength       =   20
                  DataField       =   "InSelfFax"
                  IdName          =   "INSELFFAX"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   39
                  Left            =   1815
                  TabIndex        =   57
                  ToolTipText     =   "Enter Company E-Mail"
                  Top             =   1980
                  Width           =   7035
                  _ExtentX        =   12409
                  _ExtentY        =   503
                  MaxLength       =   45
                  DataField       =   "InSelfEMail"
                  IdName          =   "INSELFEMAIL"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   41
                  Left            =   1815
                  TabIndex        =   58
                  ToolTipText     =   "Enter Name Of The Export Officer"
                  Top             =   2265
                  Width           =   4725
                  _ExtentX        =   8334
                  _ExtentY        =   503
                  MaxLength       =   30
                  DataField       =   "InHExpOff"
                  IdName          =   "INHEXPOFF"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   43
                  Left            =   1815
                  TabIndex        =   51
                  ToolTipText     =   "Enter Company Name"
                  Top             =   270
                  Width           =   7035
                  _ExtentX        =   12409
                  _ExtentY        =   503
                  MaxLength       =   45
                  DataField       =   "InSelfName"
                  IdName          =   "INSELFNAME"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   159
                  Left            =   1815
                  TabIndex        =   69
                  ToolTipText     =   "Enter Company CST Number"
                  Top             =   6405
                  Width           =   8340
                  _ExtentX        =   14711
                  _ExtentY        =   503
                  MaxLength       =   60
                  DataType        =   4
                  DataField       =   "InHCSTNO"
                  IdName          =   "INHCSTNO"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   158
                  Left            =   1815
                  TabIndex        =   67
                  ToolTipText     =   "Enter Company VAT Number"
                  Top             =   6120
                  Width           =   8340
                  _ExtentX        =   14711
                  _ExtentY        =   503
                  MaxLength       =   60
                  DataType        =   4
                  DataField       =   "InHVATNO"
                  IdName          =   "INHVATNO"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   199
                  Left            =   11205
                  TabIndex        =   68
                  ToolTipText     =   "Enter Company Tin Number"
                  Top             =   6120
                  Width           =   3105
                  _ExtentX        =   5477
                  _ExtentY        =   503
                  MaxLength       =   20
                  DataType        =   4
                  DataField       =   "InHTinNo"
                  IdName          =   "INHTINNO"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   251
                  Left            =   11205
                  TabIndex        =   70
                  ToolTipText     =   "Enter Company PAN"
                  Top             =   6405
                  Width           =   3105
                  _ExtentX        =   5477
                  _ExtentY        =   503
                  MaxLength       =   20
                  DataType        =   4
                  DataField       =   "InHPAN"
                  IdName          =   "INHPAN"
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "PAN"
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
                  Index           =   186
                  Left            =   10440
                  TabIndex        =   442
                  Top             =   6405
                  Width           =   735
               End
               Begin VB.Label LblMain 
                  BackStyle       =   0  'Transparent
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "  Local Invoice"
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
                  Index           =   12
                  Left            =   0
                  TabIndex        =   441
                  Top             =   5760
                  Width           =   14805
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "CST No"
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
                  Index           =   130
                  Left            =   90
                  TabIndex        =   440
                  Top             =   6405
                  Width           =   1425
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "VAT No"
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
                  Index           =   131
                  Left            =   90
                  TabIndex        =   439
                  Top             =   6120
                  Width           =   1425
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "TIN"
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
                  Index           =   152
                  Left            =   10440
                  TabIndex        =   438
                  Top             =   6120
                  Width           =   735
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Export Officer"
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
                  Left            =   90
                  TabIndex        =   270
                  Top             =   2265
                  Width           =   1725
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "E-Mail"
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
                  Left            =   90
                  TabIndex        =   269
                  Top             =   1980
                  Width           =   1545
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "GST"
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
                  Left            =   90
                  TabIndex        =   268
                  Top             =   1695
                  Width           =   1545
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Tel"
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
                  Left            =   90
                  TabIndex        =   267
                  Top             =   1410
                  Width           =   1545
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Address"
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
                  Index           =   30
                  Left            =   90
                  TabIndex        =   266
                  Top             =   555
                  Width           =   1545
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Name"
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
                  Index           =   31
                  Left            =   90
                  TabIndex        =   265
                  Top             =   270
                  Width           =   855
               End
            End
            Begin VB.Frame FraTabInvDet 
               BorderStyle     =   0  'None
               Height          =   7580
               Index           =   0
               Left            =   60
               TabIndex        =   219
               Top             =   360
               Width           =   14625
               Begin VB.Frame FraAg 
                  Height          =   1695
                  Left            =   -30
                  TabIndex        =   412
                  Top             =   5910
                  Width           =   14685
                  Begin MwfCtl.MWCTL_MED atxt 
                     Height          =   285
                     Index           =   184
                     Left            =   3165
                     TabIndex        =   47
                     ToolTipText     =   "Enter Agent Name"
                     Top             =   420
                     Width           =   4725
                     _ExtentX        =   8334
                     _ExtentY        =   503
                     MaxLength       =   30
                     DataField       =   "InCmAgName"
                     ReCalcOn        =   "WAGCD"
                     IdName          =   "INCMAGNAME"
                  End
                  Begin MwfCtl.MWCTL_MED atxt 
                     Height          =   285
                     Index           =   185
                     Left            =   1815
                     TabIndex        =   48
                     ToolTipText     =   "Enter Agent Address"
                     Top             =   705
                     Width           =   8295
                     _ExtentX        =   14631
                     _ExtentY        =   503
                     MaxLength       =   45
                     DataField       =   "InCmAgAdd1"
                     ReCalcOn        =   "WAGCD"
                     IdName          =   "INCMAGADD1"
                  End
                  Begin MwfCtl.MWCTL_MED atxt 
                     Height          =   285
                     Index           =   186
                     Left            =   1815
                     TabIndex        =   49
                     ToolTipText     =   "Enter Agent Address"
                     Top             =   990
                     Width           =   8295
                     _ExtentX        =   14631
                     _ExtentY        =   503
                     MaxLength       =   45
                     DataField       =   "InCmAgAdd2"
                     ReCalcOn        =   "WAGCD"
                     IdName          =   "INCMAGADD2"
                  End
                  Begin MwfCtl.MWCTL_MED atxt 
                     Height          =   285
                     Index           =   187
                     Left            =   1815
                     TabIndex        =   50
                     ToolTipText     =   "Enter Agent Address"
                     Top             =   1275
                     Width           =   8295
                     _ExtentX        =   14631
                     _ExtentY        =   503
                     MaxLength       =   45
                     DataField       =   "InCmAgAdd3"
                     ReCalcOn        =   "WAGCD"
                     IdName          =   "INCMAGADD3"
                  End
                  Begin MwfCtl.MWCTL_MED atxt 
                     Height          =   285
                     Index           =   188
                     Left            =   1815
                     TabIndex        =   46
                     ToolTipText     =   "Enter Agent Code"
                     Top             =   420
                     Width           =   1245
                     _ExtentX        =   2196
                     _ExtentY        =   503
                     MaxLength       =   5
                     DataType        =   4
                     ReCalcParent    =   "INCMAGNAME,INCMAGADD1,INCMAGADD2,INCMAGADD3"
                     ReCalcOn        =   "INCMCD"
                     IdName          =   "WAGCD"
                  End
                  Begin VB.Label ALBL 
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
                     Height          =   285
                     Index           =   151
                     Left            =   3090
                     TabIndex        =   416
                     Top             =   420
                     Width           =   75
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Address"
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
                     Index           =   150
                     Left            =   90
                     TabIndex        =   415
                     Top             =   690
                     Width           =   1545
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Agent Cd/ Name"
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
                     Index           =   149
                     Left            =   90
                     TabIndex        =   414
                     Top             =   420
                     Width           =   1665
                  End
                  Begin VB.Label LblMain 
                     BackStyle       =   0  'Transparent
                     BorderStyle     =   1  'Fixed Single
                     Caption         =   "  Agent"
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
                     Index           =   10
                     Left            =   -30
                     TabIndex        =   413
                     Top             =   90
                     Width           =   15165
                  End
               End
               Begin VB.Frame FraCon 
                  Height          =   1695
                  Left            =   -30
                  TabIndex        =   245
                  Top             =   4320
                  Width           =   14685
                  Begin MwfCtl.MWCTL_MED atxt 
                     Height          =   285
                     Index           =   109
                     Left            =   4095
                     TabIndex        =   42
                     ToolTipText     =   "Enter Consignee"
                     Top             =   420
                     Width           =   7035
                     _ExtentX        =   12409
                     _ExtentY        =   503
                     MaxLength       =   45
                     DataField       =   "InHConsignee"
                     ReCalcOn        =   "INCSTCD"
                     IdName          =   "INHCONSIGNEE"
                  End
                  Begin MwfCtl.MWCTL_MED atxt 
                     Height          =   285
                     Index           =   10
                     Left            =   1815
                     TabIndex        =   43
                     ToolTipText     =   "Enter Consignee Address"
                     Top             =   705
                     Width           =   9300
                     _ExtentX        =   16404
                     _ExtentY        =   503
                     MaxLength       =   55
                     DataField       =   "InConAdd1"
                     ReCalcOn        =   "INCSTCD"
                     IdName          =   "INCONADD1"
                  End
                  Begin MwfCtl.MWCTL_MED atxt 
                     Height          =   285
                     Index           =   11
                     Left            =   1815
                     TabIndex        =   44
                     ToolTipText     =   "Enter Consignee Address"
                     Top             =   990
                     Width           =   9300
                     _ExtentX        =   16404
                     _ExtentY        =   503
                     MaxLength       =   55
                     DataField       =   "InConAdd2"
                     ReCalcOn        =   "INCSTCD"
                     IdName          =   "INCONADD2"
                  End
                  Begin MwfCtl.MWCTL_MED atxt 
                     Height          =   285
                     Index           =   12
                     Left            =   1815
                     TabIndex        =   45
                     ToolTipText     =   "Enter Consignee Address"
                     Top             =   1275
                     Width           =   9300
                     _ExtentX        =   16404
                     _ExtentY        =   503
                     MaxLength       =   55
                     DataField       =   "InConAdd3"
                     ReCalcOn        =   "INCSTCD"
                     IdName          =   "INCONADD3"
                  End
                  Begin MwfCtl.MWCTL_MED atxt 
                     Height          =   285
                     Index           =   259
                     Left            =   1815
                     TabIndex        =   41
                     ToolTipText     =   "Enter Customer Store Code"
                     Top             =   420
                     Width           =   2205
                     _ExtentX        =   3889
                     _ExtentY        =   503
                     MaxLength       =   12
                     DataType        =   4
                     DataField       =   "INCSTCD"
                     ReCalcParent    =   "INCMAGNAME,INCMAGADD1,INCMAGADD2,INCMAGADD3,INHCONSIGNEE,INCONADD1,INCONADD2,INCONADD3"
                     IdName          =   "INCSTCD"
                  End
                  Begin VB.Label ALBL 
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
                     Height          =   285
                     Index           =   198
                     Left            =   4030
                     TabIndex        =   577
                     Top             =   420
                     Width           =   75
                  End
                  Begin VB.Label LblMain 
                     BackStyle       =   0  'Transparent
                     BorderStyle     =   1  'Fixed Single
                     Caption         =   "  Consignee"
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
                     Index           =   6
                     Left            =   -30
                     TabIndex        =   349
                     Top             =   90
                     Width           =   15165
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Consignee"
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
                     Index           =   87
                     Left            =   90
                     TabIndex        =   247
                     Top             =   420
                     Width           =   1245
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Address"
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
                     TabIndex        =   246
                     Top             =   690
                     Width           =   1545
                  End
               End
               Begin VB.Frame FraBuy 
                  Height          =   1725
                  Left            =   -30
                  TabIndex        =   248
                  Top             =   2670
                  Width           =   14685
                  Begin MwfCtl.MWCTL_MED atxt 
                     Height          =   285
                     Index           =   25
                     Left            =   1815
                     TabIndex        =   37
                     ToolTipText     =   "Enter Buyer"
                     Top             =   450
                     Width           =   7035
                     _ExtentX        =   12409
                     _ExtentY        =   503
                     MaxLength       =   45
                     DataType        =   4
                     DataField       =   "InCmBuyer"
                     IdName          =   "INCMBUYER"
                  End
                  Begin MwfCtl.MWCTL_MED atxt 
                     Height          =   285
                     Index           =   7
                     Left            =   1815
                     TabIndex        =   38
                     ToolTipText     =   "Enter Buyer Address"
                     Top             =   735
                     Width           =   8295
                     _ExtentX        =   14631
                     _ExtentY        =   503
                     MaxLength       =   55
                     DataField       =   "InBuyAdd1"
                     IdName          =   "INBUYADD1"
                  End
                  Begin MwfCtl.MWCTL_MED atxt 
                     Height          =   285
                     Index           =   8
                     Left            =   1815
                     TabIndex        =   39
                     ToolTipText     =   "Enter Buyer Address"
                     Top             =   1020
                     Width           =   8295
                     _ExtentX        =   14631
                     _ExtentY        =   503
                     MaxLength       =   55
                     DataField       =   "InBuyAdd2"
                     IdName          =   "INBUYADD2"
                  End
                  Begin MwfCtl.MWCTL_MED atxt 
                     Height          =   285
                     Index           =   9
                     Left            =   1815
                     TabIndex        =   40
                     ToolTipText     =   "Enter Buyer Address"
                     Top             =   1305
                     Width           =   8295
                     _ExtentX        =   14631
                     _ExtentY        =   503
                     MaxLength       =   55
                     DataField       =   "InBuyAdd3"
                     IdName          =   "INBUYADD3"
                  End
                  Begin VB.Label LblMain 
                     BackStyle       =   0  'Transparent
                     BorderStyle     =   1  'Fixed Single
                     Caption         =   "  Buyer"
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
                     Index           =   7
                     Left            =   -30
                     TabIndex        =   350
                     Top             =   90
                     Width           =   15165
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Buyer (Invoice)"
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
                     Left            =   90
                     TabIndex        =   250
                     Top             =   450
                     Width           =   1635
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Address"
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
                     TabIndex        =   249
                     Top             =   735
                     Width           =   1545
                  End
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   17
                  Left            =   1815
                  TabIndex        =   30
                  ToolTipText     =   "Enter Customer Address"
                  Top             =   810
                  Width           =   7035
                  _ExtentX        =   12409
                  _ExtentY        =   503
                  MaxLength       =   45
                  DataField       =   "InCmAdd1"
                  IdName          =   "INCMADD1"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   18
                  Left            =   1815
                  TabIndex        =   31
                  ToolTipText     =   "Enter Customer Address"
                  Top             =   1095
                  Width           =   7035
                  _ExtentX        =   12409
                  _ExtentY        =   503
                  MaxLength       =   45
                  DataField       =   "InCmAdd2"
                  IdName          =   "INCMADD2"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   19
                  Left            =   1815
                  TabIndex        =   32
                  ToolTipText     =   "Enter Customer Address"
                  Top             =   1380
                  Width           =   7035
                  _ExtentX        =   12409
                  _ExtentY        =   503
                  MaxLength       =   45
                  DataField       =   "InCmAdd3"
                  IdName          =   "INCMADD3"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   20
                  Left            =   1815
                  TabIndex        =   33
                  ToolTipText     =   "Enter Customer Telephone No"
                  Top             =   1665
                  Width           =   3105
                  _ExtentX        =   5477
                  _ExtentY        =   503
                  MaxLength       =   20
                  DataField       =   "InCmTel"
                  ReCalcOn        =   "INCMCD,INHINVTYP"
                  IdName          =   "INCMTEL"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   21
                  Left            =   5745
                  TabIndex        =   34
                  ToolTipText     =   "Enter Customer GST No"
                  Top             =   1665
                  Width           =   3105
                  _ExtentX        =   5477
                  _ExtentY        =   503
                  MaxLength       =   20
                  DataField       =   "InCmFax"
                  ReCalcOn        =   "INCMCD,INHINVTYP"
                  IdName          =   "INCMFAX"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   22
                  Left            =   1815
                  TabIndex        =   35
                  ToolTipText     =   "Enter Customer E-Mail"
                  Top             =   1950
                  Width           =   7035
                  _ExtentX        =   12409
                  _ExtentY        =   503
                  MaxLength       =   45
                  DataField       =   "InCmeMail"
                  ReCalcOn        =   "INCMCD,INHINVTYP"
                  IdName          =   "INCMEMAIL"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   26
                  Left            =   1815
                  TabIndex        =   36
                  ToolTipText     =   "Enter Contact Person Name"
                  Top             =   2235
                  Width           =   7035
                  _ExtentX        =   12409
                  _ExtentY        =   503
                  MaxLength       =   45
                  DataField       =   "InCmCpName"
                  ReCalcOn        =   "INCMCD,INHINVTYP"
                  IdName          =   "INCMCPNAME"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   3
                  Left            =   1815
                  TabIndex        =   29
                  ToolTipText     =   "Enter Customer Name"
                  Top             =   525
                  Width           =   7035
                  _ExtentX        =   12409
                  _ExtentY        =   503
                  MaxLength       =   45
                  DataField       =   "InCmName"
                  ReCalcParent    =   "INCMCD"
                  ReCalcOn        =   "INCMCD,INHINVTYP"
                  IdName          =   "INCMNAME"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   153
                  Left            =   1815
                  TabIndex        =   370
                  ToolTipText     =   "Customer Currency Code"
                  Top             =   240
                  Width           =   885
                  _ExtentX        =   1561
                  _ExtentY        =   503
                  Enabled         =   0   'False
                  MaxLength       =   5
                  DataType        =   4
                  ReCalcOn        =   "INCMCD"
                  IdName          =   "WCURCD"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   174
                  Left            =   13230
                  TabIndex        =   406
                  Top             =   120
                  Visible         =   0   'False
                  Width           =   1320
                  _ExtentX        =   2328
                  _ExtentY        =   503
                  Enabled         =   0   'False
                  MaxLength       =   8
                  DataType        =   3
                  DataField       =   "InMaxBgDt_Pg"
                  IdName          =   "INMAXBGDT_PG"
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
                  Index           =   126
                  Left            =   90
                  TabIndex        =   371
                  Top             =   240
                  Width           =   1635
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Cont Pers Name"
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
                  Index           =   23
                  Left            =   90
                  TabIndex        =   256
                  Top             =   2235
                  Width           =   1695
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "E-Mail"
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
                  Left            =   90
                  TabIndex        =   255
                  Top             =   1950
                  Width           =   1545
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "GST"
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
                  Left            =   5310
                  TabIndex        =   254
                  Top             =   1665
                  Width           =   465
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Tel"
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
                  Left            =   90
                  TabIndex        =   253
                  Top             =   1665
                  Width           =   1545
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Address"
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
                  Left            =   90
                  TabIndex        =   252
                  Top             =   810
                  Width           =   1545
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Name"
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
                  TabIndex        =   251
                  Top             =   525
                  Width           =   855
               End
            End
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   15
            Left            =   5955
            TabIndex        =   5
            ToolTipText     =   "Enter Invoice Voucher Date"
            Top             =   0
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   3
            DataField       =   "InDt"
            ReCalcParent    =   "INCNVRT,INCNVFCT,INLMGCST,INLMPCST,INLMSCST,INLMGDT,INLMPDT,INLMSDT,INLMLCST,INLMLDT"
            IdName          =   "INDT"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   27
            Left            =   1095
            TabIndex        =   10
            ToolTipText     =   "Enter Invoice Export Number"
            Top             =   285
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   503
            MaxLength       =   20
            DataType        =   4
            DataField       =   "InExpNo"
            ReCalcOn        =   "INNO"
            IdName          =   "INEXPNO"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   28
            Left            =   7920
            TabIndex        =   6
            ToolTipText     =   "Enter London Metal Exchange Rate For Gold"
            Top             =   0
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "###0.00"
            MaxLength       =   7
            DataType        =   2
            DataField       =   "InLmgCst"
            ReCalcParent    =   "INLMGSAL"
            ReCalcOn        =   "INDT,INCMCD"
            IdName          =   "INLMGCST"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   30
            Left            =   7920
            TabIndex        =   16
            ToolTipText     =   "Enter LME Sales Rate For Gold"
            Top             =   285
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "###0.00"
            MaxLength       =   7
            DataType        =   2
            DataField       =   "InLmgSal"
            ReCalcOn        =   "INLMGCST,INCMCD"
            IdName          =   "INLMGSAL"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   31
            Left            =   1095
            TabIndex        =   12
            ToolTipText     =   "Enter Customer Code"
            Top             =   570
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   4
            DataField       =   "InCmCd"
            ReCalcParent    =   $"EmrFrmInvHd.frx":0521
            ReCalcOn        =   "INCMNAME"
            IdName          =   "INCMCD"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   34
            Left            =   2835
            TabIndex        =   22
            ToolTipText     =   "Enter Value Addition Percentage For Accessories"
            Top             =   855
            Width           =   585
            _ExtentX        =   1032
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "#0.00"
            MaxLength       =   5
            DataType        =   2
            DataField       =   "InVaX"
            IdName          =   "INVAX"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   36
            Left            =   4365
            TabIndex        =   23
            ToolTipText     =   "Enter Value Addition Percentage For Dia And ColStn"
            Top             =   855
            Width           =   585
            _ExtentX        =   1032
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "#0.00"
            MaxLength       =   5
            DataType        =   2
            DataField       =   "InVaDc"
            IdName          =   "INVADC"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   91
            Left            =   4650
            TabIndex        =   14
            ToolTipText     =   "Enter If Dispatch Has Taken Place For The Invoice (Y/N)"
            Top             =   570
            Width           =   300
            _ExtentX        =   529
            _ExtentY        =   503
            DataType        =   4
            DataField       =   "InCls"
            IdName          =   "INCLS"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   80
            Left            =   5955
            TabIndex        =   11
            ToolTipText     =   "Enter Date To Be Printed In The Invoice Reports"
            Top             =   285
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   503
            Enabled         =   0   'False
            MaxLength       =   8
            DataType        =   3
            DataField       =   "InPrnDt"
            IdName          =   "INPRNDT"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   111
            Left            =   9885
            TabIndex        =   7
            ToolTipText     =   "Enter LMP Cost"
            Top             =   0
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "###0.00"
            MaxLength       =   7
            DataType        =   2
            DataField       =   "InLmpCst"
            ReCalcParent    =   "INLMPSAL"
            ReCalcOn        =   "INDT,INCMCD"
            IdName          =   "INLMPCST"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   112
            Left            =   9885
            TabIndex        =   17
            ToolTipText     =   "Enter LMP Sales"
            Top             =   285
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "###0.00"
            MaxLength       =   7
            DataType        =   2
            DataField       =   "InLmpSal"
            ReCalcOn        =   "INLMPCST,INCMCD"
            IdName          =   "INLMPSAL"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   124
            Left            =   5955
            TabIndex        =   21
            ToolTipText     =   "Enter Invoice Type"
            Top             =   855
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   503
            MaxLength       =   5
            DataType        =   4
            DataField       =   "InhInvTyp"
            ReCalcParent    =   "INCMNAME,INCMTEL,INCMFAX,INCMEMAIL,INCMCPNAME,INCMNAME,INCMTEL,INCMFAX,INCMEMAIL,INCMCPNAME"
            IdName          =   "INHINVTYP"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   138
            Left            =   7920
            TabIndex        =   24
            ToolTipText     =   "Enter Gold Fineness"
            Top             =   855
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "#0.000"
            MaxLength       =   6
            DataType        =   2
            DataField       =   "InHGldFiness"
            IdName          =   "INHGLDFINESS"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   139
            Left            =   9885
            TabIndex        =   25
            ToolTipText     =   "Enter Platinum Fineness"
            Top             =   855
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "#0.000"
            MaxLength       =   6
            DataType        =   2
            DataField       =   "InHPlFiness"
            IdName          =   "INHPLFINESS"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   146
            Left            =   11895
            TabIndex        =   18
            ToolTipText     =   "Enter LMS Sales"
            Top             =   285
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "###0.00"
            MaxLength       =   7
            DataType        =   2
            DataField       =   "InLmsSal"
            ReCalcOn        =   "INLMSCST,INCMCD"
            IdName          =   "INLMSSAL"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   145
            Left            =   11895
            TabIndex        =   8
            ToolTipText     =   "Enter LMP Cost"
            Top             =   0
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "###0.00"
            MaxLength       =   7
            DataType        =   2
            DataField       =   "InLmsCst"
            ReCalcParent    =   "INLMSSAL"
            ReCalcOn        =   "INDT,INCMCD"
            IdName          =   "INLMSCST"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   148
            Left            =   11895
            TabIndex        =   26
            ToolTipText     =   "Enter Silver Fineness"
            Top             =   855
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "#0.000"
            MaxLength       =   6
            DataType        =   2
            DataField       =   "InHSlFiness"
            IdName          =   "INHSLFINESS"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   178
            Left            =   13770
            TabIndex        =   9
            ToolTipText     =   "Enter LML Cost"
            Top             =   0
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "###0.00"
            MaxLength       =   7
            DataType        =   2
            DataField       =   "InLmlCst"
            ReCalcParent    =   "INLMLSAL"
            ReCalcOn        =   "INDT,INCMCD"
            IdName          =   "INLMLCST"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   180
            Left            =   13770
            TabIndex        =   19
            ToolTipText     =   "Enter LML Sales"
            Top             =   285
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "###0.00"
            MaxLength       =   7
            DataType        =   2
            DataField       =   "InLmlSal"
            ReCalcOn        =   "INLMLCST,INCMCD"
            IdName          =   "INLMLSAL"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   181
            Left            =   13770
            TabIndex        =   27
            ToolTipText     =   "Enter Palladoim Fineness"
            Top             =   855
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "#0.000"
            MaxLength       =   6
            DataType        =   2
            DataField       =   "InHPdFiness"
            IdName          =   "INHPDFINESS"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   183
            Left            =   3550
            TabIndex        =   13
            ToolTipText     =   "Enter 'Y' If Invoicing has to be done By One Bag per Sr."
            Top             =   570
            Width           =   300
            _ExtentX        =   529
            _ExtentY        =   503
            DataType        =   4
            DataField       =   "INBYBAGYN"
            IdName          =   "INBYBAGYN"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   191
            Left            =   5955
            TabIndex        =   15
            ToolTipText     =   "Rate Refresh Date For Metals"
            Top             =   570
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   503
            Enabled         =   0   'False
            MaxLength       =   8
            DataType        =   3
            DataField       =   "InRefreshDt"
            IdName          =   "INREFRESHDT"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   29
            Left            =   7920
            TabIndex        =   426
            ToolTipText     =   "Enter LME Date For Gold"
            Top             =   570
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   3
            DataField       =   "InLmgDt"
            ReCalcOn        =   "INDT,INCMCD"
            IdName          =   "INLMGDT"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   113
            Left            =   9885
            TabIndex        =   427
            ToolTipText     =   "Enter LMP Date"
            Top             =   570
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   3
            DataField       =   "InLmpDt"
            ReCalcOn        =   "INDT,INCMCD"
            IdName          =   "INLMPDT"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   147
            Left            =   11895
            TabIndex        =   428
            ToolTipText     =   "Enter LMS Date"
            Top             =   570
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   3
            DataField       =   "InLmsDt"
            ReCalcOn        =   "INDT,INCMCD"
            IdName          =   "INLMSDT"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   179
            Left            =   13770
            TabIndex        =   429
            ToolTipText     =   "Enter LML Date"
            Top             =   570
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   503
            Alignment       =   2
            MaxLength       =   8
            DataType        =   3
            DataField       =   "InLmlDt"
            ReCalcOn        =   "INDT,INCMCD"
            IdName          =   "INLMLDT"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   252
            Left            =   5250
            TabIndex        =   568
            ToolTipText     =   "Enter 'Y' If Invoicing has to be done By One Bag per Sr."
            Top             =   570
            Visible         =   0   'False
            Width           =   300
            _ExtentX        =   529
            _ExtentY        =   503
            DataType        =   4
            DataField       =   "InGstBase"
            IdName          =   "INGSTBASE"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   253
            Left            =   5340
            TabIndex        =   569
            ToolTipText     =   "Enter 'Y' If Invoicing has to be done By One Bag per Sr."
            Top             =   90
            Visible         =   0   'False
            Width           =   300
            _ExtentX        =   529
            _ExtentY        =   503
            DataType        =   4
            DataField       =   "InCmCtg"
            IdName          =   "INCMCTG"
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "LMG Dt"
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
            Left            =   6870
            TabIndex        =   433
            Top             =   570
            Width           =   1035
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "LMP Dt"
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
            Index           =   73
            Left            =   8895
            TabIndex        =   432
            Top             =   570
            Width           =   1035
         End
         Begin VB.Label LblLmsDt 
            BackStyle       =   0  'Transparent
            Caption         =   "LMS Dt"
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
            Left            =   10845
            TabIndex        =   431
            Top             =   570
            Width           =   1035
         End
         Begin VB.Label LblLmlDt 
            BackStyle       =   0  'Transparent
            Caption         =   "LML Dt"
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
            Left            =   12825
            TabIndex        =   430
            Top             =   570
            Width           =   1035
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Ref Rt Dt"
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
            Index           =   154
            Left            =   5010
            TabIndex        =   419
            Top             =   570
            Width           =   825
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Inv By Bag"
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
            Index           =   148
            Left            =   2505
            TabIndex        =   411
            Top             =   570
            Width           =   1110
         End
         Begin VB.Label LblLmlSal 
            BackStyle       =   0  'Transparent
            Caption         =   "LML Sal"
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
            Left            =   12810
            TabIndex        =   410
            Top             =   270
            Width           =   1035
         End
         Begin VB.Label LblLmlCst 
            BackStyle       =   0  'Transparent
            Caption         =   "LML Cst"
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
            Left            =   12810
            TabIndex        =   409
            Top             =   0
            Width           =   1035
         End
         Begin VB.Label LblPdFiness 
            BackStyle       =   0  'Transparent
            Caption         =   "Pd Finess"
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
            Left            =   12810
            TabIndex        =   408
            Top             =   855
            Width           =   1035
         End
         Begin VB.Label LblSlFiness 
            BackStyle       =   0  'Transparent
            Caption         =   "Slv Finess"
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
            Left            =   10830
            TabIndex        =   365
            Top             =   855
            Width           =   1035
         End
         Begin VB.Label LblLmsCst 
            BackStyle       =   0  'Transparent
            Caption         =   "LMS Cst"
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
            Left            =   10830
            TabIndex        =   363
            Top             =   0
            Width           =   1035
         End
         Begin VB.Label LblLmsSal 
            BackStyle       =   0  'Transparent
            Caption         =   "LMS Sal"
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
            Left            =   10830
            TabIndex        =   362
            Top             =   270
            Width           =   1035
         End
         Begin VB.Label LblPlFiness 
            BackStyle       =   0  'Transparent
            Caption         =   "Plt Finess"
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
            Left            =   8880
            TabIndex        =   341
            Top             =   855
            Width           =   1035
         End
         Begin VB.Label LblGldFiness 
            BackStyle       =   0  'Transparent
            Caption         =   "Gld Finess"
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
            Left            =   6855
            TabIndex        =   340
            Top             =   855
            Width           =   1035
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Password"
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
            Index           =   111
            Left            =   0
            TabIndex        =   244
            Top             =   855
            Width           =   1005
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Inv Type"
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
            Index           =   105
            Left            =   5010
            TabIndex        =   243
            Top             =   855
            Width           =   915
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "LMP Sal"
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
            Index           =   75
            Left            =   8880
            TabIndex        =   242
            Top             =   270
            Width           =   1035
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "LMP Cst"
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
            Left            =   8880
            TabIndex        =   241
            Top             =   0
            Width           =   1035
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Prn Dt"
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
            Left            =   5010
            TabIndex        =   240
            Top             =   285
            Width           =   915
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Inv Cls"
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
            Index           =   77
            Left            =   3930
            TabIndex        =   239
            Top             =   570
            Width           =   765
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "VA % X"
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
            Index           =   32
            Left            =   1950
            TabIndex        =   238
            Top             =   855
            Width           =   915
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "VA % DC"
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
            Left            =   3540
            TabIndex        =   237
            Top             =   855
            Width           =   1125
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "LMG Cst"
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
            Left            =   6855
            TabIndex        =   236
            Top             =   0
            Width           =   1035
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "LMG Sal"
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
            Left            =   6855
            TabIndex        =   235
            Top             =   270
            Width           =   1035
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Inv Exp No"
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
            Left            =   0
            TabIndex        =   234
            Top             =   360
            Width           =   1125
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Date"
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
            Left            =   5010
            TabIndex        =   233
            Top             =   0
            Width           =   915
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
            Height          =   285
            Index           =   3
            Left            =   0
            TabIndex        =   224
            Top             =   570
            Width           =   1005
         End
      End
   End
   Begin VB.Frame FraCD 
      Height          =   1035
      Left            =   7800
      TabIndex        =   455
      ToolTipText     =   "#0.0000"
      Top             =   8280
      Visible         =   0   'False
      Width           =   3255
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   212
         Left            =   1350
         TabIndex        =   456
         ToolTipText     =   "Enter New Invoice Date"
         Top             =   465
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   503
         MaxLength       =   8
         DataType        =   3
         IdName          =   "WINDT"
      End
      Begin MwfCtl.MWCTL_BTN1 CmdCDGo 
         Height          =   495
         Left            =   2520
         TabIndex        =   459
         ToolTipText     =   "Press This Button To Copy"
         Top             =   360
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
      Begin VB.Label LblMain 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "  Change Date"
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
         Index           =   16
         Left            =   -30
         TabIndex        =   458
         Top             =   90
         Width           =   3285
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "New Date"
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
         Left            =   150
         TabIndex        =   457
         Top             =   465
         Width           =   1245
      End
   End
   Begin VB.Label ALBL 
      BackStyle       =   0  'Transparent
      Caption         =   "TCS "
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
      Left            =   0
      TabIndex        =   573
      Top             =   0
      Width           =   855
   End
End
Attribute VB_Name = "EmrFrmInvHd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'----   General Comments on the Entry Form  -------'
'   One Menu 'Invoice Header'
'   Tab Page, TabInvDet with 5 Tabs
'     1.Customer Detail
'     2.Company Details
'     3.Amounts
'     4.Document Details
'     5.Transportation Details
'
'----
'*** Declare the variables ms_LockDt, mf_DtLocked for the Lock Date Check
Dim ms_LockDt As String
Dim ms_FutureDt As String
Dim ms_PurgeDt As String        '****** Sachn 3.02 Checking the Purge Date

Dim mf_DtLocked As Boolean
Dim mLocalExp As String
Dim moCn As MwfLib.MDOConnection   '*** (Bef speed) bef 24/11/06
Dim ms_TcTyp As String, ms_InCoCd As String
Dim ms_InTc As String, ms_InYy As String, ms_InChr As String
Dim ms_CmCtg As String, mf_SlvMod As Boolean, mf_GenExpNo As Boolean
Dim wDsgExists As Boolean
Dim mb_ChgInDt As Boolean

Dim mf_PdMod As Boolean ' *** Manali 3.03 - 21/06/08 - Pd Mod

'4.3.0.0
Private Enum en_TxnRmFra
  ChgDt = 0
  AddAmt = 1     '***********Bhavna
End Enum
Dim mArr_Cv As Variant, mArr_Ed As Variant  '4.3.0.0

Dim ms_CmCurCd As String
Dim ms_PwdFlag As Boolean '*** Jay 2.14 *** (SEO Change)
Dim mb_IzRmExists As Boolean  ' **** Manali 3.6.0 - 30/10/09 - For DTA Use Notional Rate in Custom Matching
Dim mb_RtfrRtChrt As Boolean  '***** Sachin 3.11.0 - Invoice Rate from Rate Chart
Dim mb_InvHdGST As Boolean
Dim ms_ChkQtsIdName As String

Private Sub ADC_Load()
  '*** Set the First Key Control property to the field InChr
  '*** Set the First Non Key Control property to InDt
  '*** Bring the FraNKeyAll frame to the top
  '*** Set the default Tab Option for the TabInvDet Tab Page as 'Customer Details'
  '*** Store the defa values for Inv Tc,Yy in variables
  
  Set moCn = ADC.Connection    '*** (Bef speed) bef 24/11/06
  Set ADC.FirKeyCtl = ADC("InChr")
  Set ADC.FirNKeyCtl = ADC("InDt")
             
  FraNKeyAll.ZOrder (0)
  TabInvDet.Tab = 0
  
  ReDim mArr_Cv(1), mArr_Ed(1)    '4.3.0.0
      
  Select Case UCase(ADC.MenuCd)
   '*** Jay 3.3(FG)
   Case Is = UCase("In"), UCase("InvHdFot"), UCase("InvHdJT")
     ms_TcTyp = IIF(UCase(ADC.MenuCd) = UCase("InvHdFot"), ctFotTcTyp, ctInTcTyp)
     ms_TcTyp = IIF(UCase(ADC.MenuCd) = UCase("InvHdJT"), "JT", ms_TcTyp)
     'Case Is = UCase("In")
     '  ms_TcTyp = ctInTcTyp
   '*** Jay 3.3(FG)
    ms_CmCtg = "C"
    ADC("InCmCtg") = ms_CmCtg
    '*** (Bef 2.13) LblInUSD.Caption = "in USD"
    LblInUSD.Caption = "in Cust Currency"   '*** (Jen 2.13)
    LblLocalCurncyHd.Caption = "in Ind Rs"
    LblLocalCurncyHd.Visible = True
    ADC("InFOBRs").Visible = True: ADC("InFOBRs").Enabled = True
    ADC("InCommRs").Visible = True: ADC("InCommRs").Enabled = True
    ADC("InFrtRs").Visible = True: ADC("InFrtRs").Enabled = True
    ADC("InInsRs").Visible = True: ADC("InInsRs").Enabled = True
    ADC("InTotRs").Visible = True: ADC("InTotRs").Enabled = True
'    If ms_TcTyp = "JT" Then
'        ms_InCoCd = gs_CoCd: ms_InTc = gs_InTc: ms_InYy = gs_InYy: ms_InChr = gs_InChr ': ms_InNo = gs_InNo
'    End If
'gst
     'If gs_InTc = "PR" Or gs_InTc = "SA" Then ms_TcTyp = "PR"
     If gs_InTc <> "" Then
       ms_TcTyp = moCn.GetFldVal("Select PValue from Param where PTyp ='TC' and PMCd='" + gs_InTc + "'")
     End If
  Case Is = UCase("CTBIn")
    ms_TcTyp = ctCTBTcTyp
    ms_CmCtg = "T"
    ADC("InCmCtg") = ms_CmCtg
    LblInUSD.Caption = "in Local Currency"
    LblLocalCurncyHd.Visible = False
    ADC("InFOBRs").Visible = False: ADC("InFOBRs").Enabled = False
    ADC("InCommRs").Visible = False: ADC("InCommRs").Enabled = False
    ADC("InFrtRs").Visible = False: ADC("InFrtRs").Enabled = False
    ADC("InInsRs").Visible = False: ADC("InInsRs").Enabled = False
    ADC("InTotRs").Visible = False: ADC("InTotRs").Enabled = False
    ' **** Manali 3.6.0 - 30/10/09 - For DTA Use Notional Rate in Custom Matching
    LblNotionRt.Visible = False: ADC("InNotionRtYn").Visible = False: ADC("InNotionRtYn").Enabled = False
    ' **** Manali 3.6.0 - 30/10/09 - For DTA Use Notional Rate in Custom Matching
  End Select
 '****Geeta***Emr210
  mLocalExp = moCn.GetFldVal("Select hExpLocal from Head where HCoCd= '" + ms_InCoCd + "' and HCd= '" + ctSelfCmCd + "' ")

  If UCase$(mLocalExp) = "N" Then
     LblGldFiness.Visible = False: ADC("InhGldFiness").Visible = False: ADC("InhGldFiness").Enabled = False
     LblPlFiness.Visible = False: ADC("InhPlFiness").Visible = False: ADC("InhPlFiness").Enabled = False
     ' ****** Sachin 2.12 - 06/12/05 - SJM
     LblSlFiness.Visible = False: ADC("InhSlFiness").Visible = False: ADC("InhSlFiness").Enabled = False
     ' ****** Sachin 2.12 - 06/12/05 - SJM
     ' **** Manali 3.03 - 21/06/08 - Pd Mod
     LblPdFiness.Visible = False: ADC("InhPdFiness").Visible = False: ADC("InhPdFiness").Enabled = False
     ' **** Manali 3.03 - 21/06/08 - Pd Mod
     ' **** Manali 3.6.0 - 30/10/09 - For DTA Use Notional Rate in Custom Matching
     
     'Line below commented in 4.1.0.0
     'LblNotionRt.Visible = False: adc("InNotionRtYn").Visible = False: adc("InNotionRtYn").Enabled = False
     
     ' **** Manali 3.6.0 - 30/10/09 - For DTA Use Notional Rate in Custom Matching
     FraDefDta.Enabled = False: FraDefDta.Visible = False
  End If
      
       
 'ms_TcTyp = ADC.MenuCd
  ms_InCoCd = gs_CoCd
  ms_InTc = moCn.GetFldVal("Select PValue from Param where PTyp='TCTYP' and " + _
            "PMCd='" + ms_TcTyp + "' ")
  ms_InYy = moCn.GetFldVal("Select hYy from Head where HCoCd= '" + ms_InCoCd + "' and HCd='" + ctSelfCmCd + "'")
  mf_SlvMod = IIF(moCn.GetFldVal("Select HSlvModYN From Head Where HCoCd = '" + ms_InCoCd + "' And HCd = '" + ctSelfCmCd + "'") = "Y", True, False)
  
  ' **** Manali 3.03 - 21/06/08 - Pd Mod
  mf_PdMod = IIF(moCn.GetFldVal("Select HPdModYN From Head Where HCoCd = '" + ms_InCoCd + "' And HCd = '" + ctSelfCmCd + "'") = "Y", True, False)
  ' **** Manali 3.03 - 21/06/08 - Pd Mod
  
  ' ****** Sachin 2.12.0 Patch 3 - 21/04/06 ******
  If mf_SlvMod = False Then
    LblLmsCst.Visible = False: LblLmsSal.Visible = False: LblLmsDt.Visible = False: LblSlFiness.Visible = False
    ADC("InLmsCst").Visible = False: ADC("InLmsSal").Visible = False: ADC("InLmsDt").Visible = False
    ADC("InHSlFiness").Visible = False
  End If
  ' ****** Sachin 2.12.0 Patch 3 - 21/04/06 ******
  
  ' **** Manali 3.03 - 21/06/08 - Pd Mod
  If mf_PdMod = False Then
    LblLmlCst.Visible = False: LblLmlSal.Visible = False: LblLmlDt.Visible = False: LblPdFiness.Visible = False
    ADC("InLmlCst").Visible = False: ADC("InLmlSal").Visible = False: ADC("InLmlDt").Visible = False
    ADC("InHPdFiness").Visible = False
  End If
  ' **** Manali 3.03 - 21/06/08 - Pd Mod
  
  '****** Sachin 2.13.0 - 23-05-2006 ******
  mf_GenExpNo = IIF(moCn.GetFldVal("Select HGenExpNoYN From Head Where HCoCd = '" + ms_InCoCd + "' And HCd = '" + ctSelfCmCd + "'") = "Y", True, False)
  If mf_GenExpNo = True Then
    ADC("InExpNo").Enabled = False
  Else
    ADC("InExpNo").Enabled = True
  End If
  '****** Sachin 2.13.0 - 23-05-2006 ******
  
  LblInTermCond.Caption = IIF(ms_InTc = "JMI", "Memo Terms and Conditions", "Invoice Terms and Conditions")
  
  '***** Sachin 3.11.0 - Rate from Rate Chart Module
  ADC("InRmRtfrRtChrt") = "N": ADC("InLabRtfrRtChrt") = "N": ADC("InOrdMulbyYN") = "N"
  mb_RtfrRtChrt = IIF(moCn.GetFldVal("Select hRtfrRtChrtYN From Head Where HCoCd = '" + ms_InCoCd + "' And HCd = '" + ctSelfCmCd + "'") = "Y", True, False)
  ADC("InRmRtfrRtChrt").Enabled = mb_RtfrRtChrt
  ADC("InLabRtfrRtChrt").Enabled = mb_RtfrRtChrt
  mb_ChgInDt = False
  
End Sub
Private Sub ADC_KeyWhen()
  '*** set the default values for InTc,InYy and InChr from the previous entry
  mf_DtLocked = False
  
  '4.3.0.0
  Call EnaDisaCmds(True)
  Call HideAllFras
  
  If ms_InCoCd <> "" Then ADC("InCoCd") = ms_InCoCd
  If ms_InTc <> "" Then ADC("InTc") = ms_InTc
  If ms_InYy <> "" Then ADC("InYy") = ms_InYy
  If ms_InChr <> "" Then ADC("InChr") = ms_InChr
  
  
  '****** Sachin 2.14.0 - [24-11-2006] - [Multi-Company] ******
  Me.Caption = GetFrmCaption(ms_InCoCd, "Invoice Header")
  '****** Sachin 2.14.0 - [24-11-2006] - [Multi-Company] ******
  
' ###################################  Manoj #### Ver: 2.0.6  ###################################
' Date:   20/01/2004
  txtPwd = ""
' ###########################################  Manoj  ###########################################
'SRI GST 'removed ms_TcTyp = "JT" from below condition
'indigo.160- company code changed from gs_CoCd to gs_InCoCd
 If gs_InNo <> 0 Then
   ADC("InCoCd") = gs_InCoCd: ADC("InTc") = gs_InTc: ADC("InYy") = gs_InYy: ADC("InChr") = gs_InChr: ADC("InNo") = gs_InNo ': SendKeys ("^f")
   ms_InCoCd = gs_InCoCd
   ADC("InCoCd").Enabled = False
   ADC("inTc").Enabled = False
   ADC("InYy").Enabled = False
   ADC("InChr").Enabled = False
   ADC("InNo").Enabled = False
   ADC("InDt").Enabled = False  'sv.12 invoice date can't editable for other voucher
 Else
   ADC("InCoCd").Enabled = True
   ADC("inTc").Enabled = True
   ADC("InYy").Enabled = True
   ADC("InChr").Enabled = True
   ADC("InNo").Enabled = True
   ADC("InDt").Enabled = True   'sv.12 invoice date can editable for invoice
 End If
'gst
 mb_InvHdGST = False
 'pg.11 TZ tctype added
 If (ms_TcTyp = "PR" Or ms_TcTyp = "CNV" Or ms_TcTyp = "MV" Or ms_TcTyp = "DT" Or ms_TcTyp = "BV" Or ms_TcTyp = "BS" Or ms_TcTyp = "BM" Or ms_TcTyp = "MLT" Or ms_TcTyp = "REJ" Or ms_TcTyp = "FB" Or ms_TcTyp = "FWB" Or ms_TcTyp = "TZ") And gs_InNo <> 0 Then
  mb_InvHdGST = True
'    ADC("InCoCd") = gs_CoCd: ADC("InTc") = gs_InTc: ADC("InYy") = gs_InYy: ADC("InChr") = gs_InChr: ADC("InNo") = gs_InNo
End If
'stellar.1- IdName of text boxes to be check quotes
 ms_ChkQtsIdName = ADC("InCmName").IdName + "|" + ADC("InCmAdd1").IdName + "|" + ADC("InCmAdd2").IdName + "|" + ADC("InCmAdd3").IdName + "|" + ADC("InCmTel").IdName + "|" + ADC("InCmFax").IdName + "|" + _
      ADC("InCmeMail").IdName + "|" + ADC("InCmCpName").IdName + "|" + ADC("InCmBuyer").IdName + "|" + ADC("InBuyAdd1").IdName + "|" + ADC("InBuyAdd2").IdName + "|" + ADC("InBuyAdd3").IdName + "|" + _
      ADC("InHConsignee").IdName + "|" + ADC("InConAdd1").IdName + "|" + ADC("InConAdd2").IdName + "|" + ADC("InConAdd3").IdName + "|" + ADC("InCmAgName").IdName + "|" + ADC("InCmAgAdd1").IdName + "|" + _
      ADC("InCmAgAdd2").IdName + "|" + ADC("InCmAgAdd3").IdName + "|" + ADC("InSelfName").IdName + "|" + ADC("InSelfAdd1").IdName + "|" + ADC("InSelfAdd2").IdName + "|" + ADC("InSelfAdd3").IdName + "|" + _
      ADC("InSelfTel").IdName + "|" + ADC("InSelfFax").IdName + "|" + ADC("InSelfEMail").IdName + "|" + ADC("InHExpOff").IdName + "|" + ADC("InHBkName").IdName + "|" + ADC("InHBkAdd1").IdName + "|" + _
      ADC("InHBkAdd2").IdName + "|" + ADC("InHBkAdd3").IdName + "|" + ADC("InHBkAc").IdName + "|" + ADC("InHFrgnBkName").IdName + "|" + ADC("InHFrgnBkAc").IdName + "|" + ADC("InHVATNO").IdName + "|" + _
      ADC("InHCSTNO").IdName + "|" + ADC("InHTinNo").IdName + "|" + ADC("InHPAN").IdName + "|" + ADC("InHImpValTxt").IdName + "|" + ADC("InTotAsTxt").IdName + "|" + ADC("InDscDesc").IdName + "|" + _
      ADC("InDeductDesc").IdName + "|" + ADC("InDiaCostDesc").IdName + "|" + ADC("InCSCostDesc").IdName + "|" + ADC("InAdvDesc").IdName + "|" + ADC("InGrNo").IdName + "|" + ADC("InSbNo").IdName + "|" + _
      ADC("InAwbNo").IdName + "|" + ADC("InHAwbNo").IdName + "|" + ADC("InGspNo").IdName + "|" + ADC("InPoNo").IdName + "|" + ADC("InChqNo").IdName + "|" + ADC("InHRbiNo").IdName + "|" + ADC("InHIecNo").IdName + "|" + _
      ADC("InHItcNo").IdName + "|" + ADC("InhAuthSgn").IdName + "|" + ADC("InHShpBillSgnBy").IdName + "|" + ADC("InCmInsBy").IdName + "|" + ADC("InHInsNo").IdName + "|" + ADC("InHAgName").IdName + "|" + _
      ADC("InHAgAdd1").IdName + "|" + ADC("InHAgAdd2").IdName + "|" + ADC("InHAgAdd3").IdName + "|" + ADC("InHStatCd").IdName + "|" + ADC("InHAR4No").IdName + "|" + ADC("InHQCertNo").IdName + "|" + _
      ADC("InHGSPOrgCrit").IdName + "|" + ADC("InHLicNo").IdName + "|" + ADC("InHETCSvg").IdName + "|" + ADC("InGspAdd1").IdName + "|" + ADC("InGspAdd2").IdName + "|" + ADC("InGspAdd3").IdName + "|" + _
      ADC("InRegNo").IdName + "|" + ADC("InCircularNo").IdName + "|" + ADC("InMMTC").IdName + "|" + ADC("InPrdDesc").IdName + "|" + ADC("InHFtRem").IdName + "|" + ADC("InHBnkInstr").IdName + "|" + _
      ADC("InHValAddnFt").IdName + "|" + ADC("InHUnResolutions").IdName + "|" + ADC("InHCustRem").IdName + "|" + ADC("InTermCond").IdName + "|" + ADC("InHGjepcNo").IdName + "|" + ADC("InHGjepcValidUpTo").IdName + "|" + _
      ADC("InHBinNo").IdName + "|" + ADC("InCERegnNo").IdName + "|" + ADC("InHsCode").IdName + "|" + ADC("InHPreCarriage").IdName + "|" + ADC("InHPreCarrier").IdName + "|" + ADC("InHvessel").IdName + "|" + _
      ADC("InHMarks").IdName + "|" + ADC("InHPkgDesc").IdName + "|" + ADC("InHEtcTxt").IdName + "|" + ADC("InHEtcZone").IdName + "|" + ADC("InHEtcSection").IdName + "|" + ADC("InHEximPolicy").IdName + "|" + _
      ADC("InHGspEligibility").IdName + "|" + ADC("InMetalUtil").IdName + "|" + ADC("InCmVATNO").IdName + "|" + ADC("InCmCSTNO").IdName + "|" + ADC("InHLclInvFT").IdName + "|" + ADC("InHLclInvVATFT").IdName + "|" + _
      ADC("InHLclInvCSTFT").IdName + "|" + ADC("InCmTinNo").IdName + "|" + ADC("InCmPAN").IdName + "|" + ADC("InLclInvTyp").IdName + "|" + ADC("InHPortofLd").IdName + "|" + ADC("InHConOfOrg").IdName + "|" + _
      ADC("InHStatOfOrg").IdName + "|" + ADC("InHMetalBOETxt").IdName + "|" + ADC("InHRbiPolicy").IdName + "|" + ADC("InBOELooseExp").IdName + "|" + ADC("InHGspPkgDesc").IdName + "|" + _
      ADC("InAmtDesc1").IdName + "|" + ADC("InAmtDesc2").IdName + "|" + ADC("InAmtDesc3").IdName + "|" + ADC("InAmtDesc4").IdName + "|" + ADC("InAmtDesc5").IdName + "|" + ADC("InAmtDesc6").IdName + "|" + _
      ADC("InAmtDesc7").IdName + "|" + ADC("InAmtDesc8").IdName + "|" + ADC("InAmtDesc9").IdName + "|" + ADC("InAmtDesc10").IdName + "|" + ADC("InAmtDesc11").IdName + "|" + ADC("InAmtDesc12").IdName + "|" + _
      ADC("InPayDesc1").IdName + "|" + ADC("InPayDesc2").IdName + "|" + ADC("InPayDesc3").IdName + "|" + _
      ADC("InExpNo").IdName + "|" + ADC("InCmPayTerms").IdName + "|" + ADC("InLclOthrValDesc").IdName

End Sub
Private Sub ADC_KeyValid(Cancel As Boolean, ErrMsg As String)
  '*** Check whether the InYy,InChr fields are valid
  If ADC("InYy") = "" Then Cancel = True: ErrMsg = "Enter Voucher Year": Exit Sub
  If ADC("InChr") = "" Then Cancel = True: ErrMsg = "Enter Voucher Character": Exit Sub
End Sub
Private Sub adc_AutoGen()
  
  '*** Autogenerate the Invoice Number
      'If the InNo is Zero then generate the number
      'If the user himself enters the number then write this number in the autogen table (if it is greater than the number in the autogen table)
  
  If ADC("InNo") = 0 Then
    ADC("InNo") = moCn.AutoGen("InvHd", "InNo", ms_InCoCd + ADC("InTc") + ADC("InYy") + ADC("InChr"))
  Else
    Call moCn.AutoGen("InvHd", "InNo", ms_InCoCd + ADC("InTc") + ADC("InYy") + ADC("InChr"), ADC("InNo"))
  End If
End Sub

Private Sub ADC_NonKeyWhen(Cancel As Boolean, ErrMsg As String)
  '*** Store the InTc,InYy and InChr values for the next entry
  '*** Give default values for the fields InDt, InVaDC, InVaX, InPrdDesc, InCls,
  '    InGrDt,InSBDt,InAWBDt,InGSPDt,InPODt,
  '    InHBkName, InHBkAdd1, InHBkAdd2, InHBkAdd3, InHBkAc, InHFrgnBkName, InHFrgnBkAc,
  '    InHExpOff, InHRbiNo, InHIecNo, InHPreCarriage, InHPreCarrier, InHVessel, InHPortOfLd,
  '    InHConOfOrg, InHMarks, InHPkgDesc, InHItcNo, InHEtcZone, InHEtcSection, InHConsignee,
  '    InHInsNo, InHAgName, InHAgAdd1, InHAgAdd2, InHAgAdd3, InHLicNo, InHETCSvg,
  '    InHAR4No, InHQCertNo, InHGSPOrgCrit, InTotAs, InHStatCd, InSelfName, InSelfAdd1,
  '    InSelfAdd2, InSelfAdd3, InSelfTel, InSelfFax, InSelfEMail
  '*** Enable the frame on the current tab and disable all other frames of tab page 'TabInvDet'
  '*** Check if Rec exists in InvDsg and store the result in a variable (wDsgExists)
  
  
      If ADC.IsNew And (Not moCn.RecSeek("select vPSCd from vParam where " + _
                         "vPCoCd= '" + ms_InCoCd + "' and vPTyp= 'CHR' and " + _
                         "vPMCd = '" + ADC("InTc") + "' and vPSCd='" + ADC("InChr") + "' and vPValidYn<>'N'")) Then
       Cancel = True: ErrMsg = "Voucher Character No More Valid": ADC.AllowSave = False: Exit Sub
      Else
        ADC.AllowSave = True
      End If

  ' ***** Manali 3.4.1 - 19/12/08 - Following Code is Shifted here from the middle of the routine
  '*** Jay 3.3(FG)

  Call EnaDisaCmds(False)   '4.3.0.0
  CmdAddAmt.Visible = True      '**********Bhavna
  
  '****** Sachin 3.02 28-12-07 - Current Partition
  If UCase(ms_TcTyp) <> UCase(ctFotTcTyp) Then
    wDsgExists = moCn.RecSeek("Select IdTc from InvDsg where " + _
                           "IdCoCd= '" + ms_InCoCd + "' and IdTc= '" + ADC("InTc") + "' and " + _
                           "IdYy= '" + ADC("InYy") + "' and IdChr= '" + ADC("InChr") + "' and " + _
                           "IdNo= " + CStr(ADC("InNo")) + " and IdPrtKey=(Case When '" + gs_Partition + "'='" + ctCurrPrtn + "' Then '" + ctCurrPrtn + "' Else IdPrtKey End) ")
       CmdChgDt.Visible = wDsgExists
        
  Else
     wDsgExists = moCn.RecSeek("Select 'x' from OrdMst where " + _
                            "OmCoCd= '" + ms_InCoCd + "' and OmTc= '" + ADC("InTc") + "' and " + _
                            "OmYy= '" + ADC("InYy") + "' and OmChr= '" + ADC("InChr") + "' and " + _
                            "OmNo= " + CStr(ADC("InNo")) + " and OmPrtKey=(Case When '" + gs_Partition + "'='" + ctCurrPrtn + "' Then '" + ctCurrPrtn + "' Else OmPrtKey End) ")
     If wDsgExists Then
       ADC("InCmCd").Enabled = False
       ADC("InFrRmLoc").Enabled = False: ADC("InToRmLoc").Enabled = False:  ADC("wInFrRmDc").Enabled = False
     Else
       ADC("InCmCd").Enabled = True
       ADC("InFrRmLoc").Enabled = True: ADC("InToRmLoc").Enabled = True:  ADC("wInFrRmDc").Enabled = True
     End If
   End If
   '*** Jay 3.3(FG)

   ' **** Manali 3.6.0 - 30/10/09 - For DTA Use Notional Rate in Custom Matching
   mb_IzRmExists = moCn.RecSeek("Select 'x' from InvZRm where IzCoCd= '" + ms_InCoCd + "' and IzTc= '" + ADC("InTc") + "' and " + _
                    "IzYy=  '" + ADC("InYy") + "' and IzChr=  '" + ADC("InChr") + "' and " + _
                    "IzNo=  " + CStr(ADC("InNo")) + " and IzPrtKey='" + ctCurrPrtn + "' ")

  '****** Sachin 3.02 28-12-07 - Current Partition
  If ADC.IsNew Then _
        ADC("InPrtKey") = ctCurrPrtn: ADC("INMAXBGDT_PG") = "01/01/80"
  
  ms_LockDt = moCn.GetFldVal("Select PDesc225 from Param where PTyp= 'LOCKDT' and PMCd= 'LOCKDT'")
  
  '****** Sachin 3.02 28-12-07 - Fetching the Purge Date
'  ms_PurgeDt = moCn.GetFldVal("Select PDesc225 from Param where PTyp= 'PURGEDT' and PMCd= 'PURGEDT'")
  
  '*** For the time being do not put Lock Date Check in the Invoice Entry
  '### If ms_LockDt = "" Then ms_LockDt = "01/01/80"
 ' ms_LockDt = "01/01/80"
  
  '****** Sachin 3.02 28-12-07 - Def Purge Date
'  If ms_PurgeDt = "" Then ms_PurgeDt = "01/01/80"
  ms_FutureDt = "01/01/80"
  '6.1_51 wVchrNo added to GetLockDt to check that given voucher number is exist in unlock ptyp,
  'if given voucher is in unlock ptyp it will not check lock date
  Dim wVchrNo As String
  wVchrNo = ADC("InTc") + "/" + ADC("InYy") + "/" + ADC("InChr") + "/" + CStr(ADC("InNo"))

   'indigo.160- company code changed from ms_InCoCd to ms_InCoCd
   Call GetLockDt(ms_InCoCd, "TC", ms_InTc, ms_LockDt, ms_FutureDt, wVchrNo)
  
  Dim wRsHead As MwfLib.MDORowSet, wRsSelfCustMst As MwfLib.MDORowSet
  ms_InTc = ADC("InTc")
  ms_InYy = ADC("InYy")
  ms_InChr = ADC("InChr")

  '  FraNKeyAll.Enabled = True
  
  If ADC.AddFndMode = xaddmode Then

    '**************** urmi 4Jul 03*********************************************************'
    '*** Jay 3.3(FG)
    'If UCase(ms_TcTyp) = UCase(ctInTcTyp) Then
    If UCase(ms_TcTyp) = UCase(ctInTcTyp) Or UCase(ms_TcTyp) = UCase(ctFotTcTyp) Or UCase(ms_TcTyp) = UCase("JT") Then
    '*** Jay 3.3(FG)
      ADC("InFrRmLoc") = moCn.GetFldVal("Select vPValue from vParam where " + _
                                       "vPCoCd= '" + ms_InCoCd + "' and vPTyp= 'CHR' " + _
                                       "and vPMCd= '" + ADC("InTc") + "' and vPSCd= '" + ADC("InChr") + "'")
      ADC("InToRmLoc") = moCn.GetFldVal("Select vPDesc225 from vParam where " + _
                                       "vPCoCd= '" + ms_InCoCd + "' and vPTyp= 'CHR' " + _
                                       "and vPMCd= '" + ADC("InTc") + "' and vPSCd= '" + ADC("InChr") + "'")
       ADC("InFrRmDc") = moCn.GetFldVal("Select vPValue2 from vParam where " + _
                                       "vPCoCd= '" + ms_InCoCd + "' and vPTyp= 'CHR' " + _
                                       "and vPMCd= '" + ADC("InTc") + "' and vPSCd= '" + ADC("InChr") + "'")
    End If
    '****************************************************************************************'
    '53 company code parameter added
    ADC("InDt") = GetNextDate(ADC("InCoCd")) 'moCn.SrvrDate       'Date
    
    ADC("InPrnDt") = moCn.SrvrDate
    ADC("InRefreshDt") = moCn.SrvrDate    ' ***** Manali 3.6.0 - 30/10/09 - Refresh Rate for Metals
    ADC("InCls") = "N"
    
    '****** Sachin 3.3.0 (23-06-08) - [42. Make Invoice By Bag] ******
    ADC("InByBagYN") = "N"
    '****** Sachin 3.3.0 (23-06-08) - [42. Make Invoice By Bag] ******
    
    ADC("InGrDt") = MWLib.EmptyDate
    ADC("InSBDt") = MWLib.EmptyDate
    ADC("InAWBDt") = MWLib.EmptyDate
    ADC("InHAWBDt") = MWLib.EmptyDate
    ADC("InGSPDt") = MWLib.EmptyDate
    ADC("InPODt") = MWLib.EmptyDate
    ADC("InTotAs") = "FOB"
    ADC("InRtFrInvCust") = "N"   '4.1.2
    
        '*************Bhavna added default for inv amt desc 7 and 12
      ADC("InAmtDesc7") = "Net Realization"
      ADC("InAmtDesc7").Enabled = False
      ADC("InAmtDesc12") = "Total Amount"
      ADC("InAmtDesc12").Enabled = False
      ADC("InAmt7").Enabled = False
      ADC("InAmt12").Enabled = False
  '*************Bhavna added default for inv amt desc 7 and 12
    
    Set wRsHead = moCn.OpenRes("Select * From Head where HCoCd= '" + ms_InCoCd + "' and HCd='" + ctSelfCmCd + "'")
    If Not (wRsHead.EOF Or wRsHead.BOF) Then
      ADC("InVaDC") = wRsHead!hVaDc
      ADC("InVaX") = wRsHead!hVaX
      ADC("InPrdDesc") = wRsHead!HPRDDESC
      ADC("InHBkName") = wRsHead!HBkName
      ADC("InHBkAdd1") = wRsHead!HBkAdd1
      ADC("InHBkAdd2") = wRsHead!HBkAdd2
      ADC("InHBkAdd3") = wRsHead!HBkAdd3
      ADC("InHBkAc") = wRsHead!HBkAc
      ADC("InHFrgnBkName") = wRsHead!HFrgnBkName
      ADC("InHFrgnBkAc") = wRsHead!HFrgnBkAc
      ADC("InHExpOff") = wRsHead!HExpOff
      ADC("InHRbiNo") = wRsHead!HRbiNo
      ADC("InHIecNo") = wRsHead!HIecNo
      ADC("InHPreCarriage") = wRsHead!HPreCarriage
      ADC("InHPreCarrier") = wRsHead!HPreCarrier
     'ADC("InHVessel") = wRsHead!HVessel
      ADC("InHPortOfLd") = wRsHead!HPortOfLd
      ADC("InHConOfOrg") = wRsHead!HConOfOrg
      ADC("InHMarks") = wRsHead!HMarks
      ADC("InHPkgDesc") = wRsHead!HPkgDesc
      ADC("InHGspPkgDesc") = wRsHead!HGspPkgDesc
      ADC("InHItcNo") = wRsHead!HItcNo
      ADC("InHEtcZone") = wRsHead!HEtcZone
      ADC("InHEtcSection") = wRsHead!HEtcSection
      ADC("InHConsignee") = wRsHead!HConsignee
      ADC("InHInsNo") = wRsHead!HInsNo
      ADC("InHAgName") = wRsHead!HAgName
      ADC("InHAgAdd1") = wRsHead!HAgAdd1
      ADC("InHAgAdd2") = wRsHead!HAgAdd2
      ADC("InHAgAdd3") = wRsHead!HAgAdd3
      ADC("InHLicNo") = wRsHead!HLicNo
      ADC("InHETCSvg") = wRsHead!HETCSvg
      ADC("InHAR4No") = wRsHead!HAR4No
      ADC("InHQCertNo") = wRsHead!HQCertNo
'Manoj
      ADC("InHGSPOrgCrit") = GetOriginCriteria
'Manoj
'      adc("InTotAs") = wRsHead!HTotAs
      ADC("InHStatCd") = wRsHead!HStatCd
      'adc("InHFtRem") = wRsHead!hFtRem         '4.1.2 - Brought from CustMst
      ADC("InHBnkInstr") = wRsHead!HBnkInstr
      ADC("InHRndOffFOB") = wRsHead!HRndOffFOB
      ADC("InHBinNo") = wRsHead!HBinNo
   ADC("InCERegnNo") = wRsHead!HCERegnNo       '**********Bhavna
      ADC("InHValAddnFt") = wRsHead!HValAddnFt
      ADC("InHUnResolutions") = wRsHead!HUnResolutions
      
      ADC("InHSCode") = wRsHead!hHSCode    '4.1.0.0
      
      ' Zubin 212
      ADC("InHCustRem") = wRsHead!HCustRem
      ' Zubin 212
      ADC("InHEximPolicy") = wRsHead!HEximPolicy
      ADC("InHGspEligibility") = wRsHead!HGspEligibility
      ADC("InhInvTyp") = wRsHead!HInvTyp
      ADC("InhAuthSgn") = wRsHead!HAuthSgn
      ADC("InHValAddMtlColYN") = wRsHead!HValAddMtlColYN
      ADC("InHValBOEDetYN") = wRsHead!HValBOEDetYN
      'urmi 28/02/04
      ADC("InMMTc") = wRsHead!HMMTc
      '*********Geeta******Emr207*****
      ADC("InHEtcTxt") = wRsHead!HEtcTxt
      ADC("InHStatOfOrg") = wRsHead!HStatOfOrg
      ADC("InHImpValTxt") = wRsHead!HImpValTxt
      ADC("InHShpBillSgnBy") = wRsHead!HShpBillSgnBy
      ADC("InHMetalBOETxt") = wRsHead!HMetalBOETxt
      '*********Geeta******Emr207*****
      '*********Geeta******Emr208*****
      
      '**********Bhavna added additional invoice amounts 1-12
'      ADC("InAmtDesc1") = wRsHead!HInAmtDesc1
'      ADC("InAmtDesc2") = wRsHead!HInAmtDesc2
'      ADC("InAmtDesc3") = wRsHead!HInAmtDesc3
'      ADC("InAmtDesc4") = wRsHead!HInAmtDesc4
'      ADC("InAmtDesc5") = wRsHead!HInAmtDesc5
'      ADC("InAmtDesc6") = wRsHead!HInAmtDesc6
'      ADC("InAmtDesc7") = wRsHead!HInAmtDesc7
'      ADC("InAmtDesc8") = wRsHead!HInAmtDesc8
'      ADC("InAmtDesc9") = wRsHead!HInAmtDesc9
'      ADC("InAmtDesc10") = wRsHead!HInAmtDesc10
'      ADC("InAmtDesc11") = wRsHead!HInAmtDesc11
'      ADC("InAmtDesc12") = wRsHead!HInAmtDesc12
'      ADC("InAmt1Per") = wRsHead!HInAmt1Per
'      ADC("InAmt2Per") = wRsHead!HInAmt2Per
'      ADC("InAmt3Per") = wRsHead!HInAmt3Per
'      ADC("InAmt4Per") = wRsHead!HInAmt4Per
'      ADC("InAmt5Per") = wRsHead!HInAmt5Per
'      ADC("InAmt6Per") = wRsHead!HInAmt6Per
'      ADC("InAmt8Per") = wRsHead!HInAmt8Per
'      ADC("InAmt9Per") = wRsHead!HInAmt9Per
'      ADC("InAmt10Per") = wRsHead!HInAmt10Per
'      ADC("InAmt11Per") = wRsHead!HInAmt11Per
        
      '*** (Bef 2.14 Next)
      'ADC("InHGldFiness") = IIF(wRsHead!hGldFiness <> 0, wRsHead!hGldFiness, 1)
      'ADC("InHPlFiness") = IIF(wRsHead!hPlFiness <> 0, wRsHead!hPlFiness, 1)
      '' ****** Sachin 2.12 - 06/12/05 - SJM
      'ADC("InHSlFiness") = IIF(wRsHead!hSlFiness <> 0, wRsHead!hSlFiness, 1)
      '*** (Bef 2.14 Next)
      
      '*** (Jen 2.14 Next)
      If UCase$(mLocalExp) = "N" Then
        ADC("InHGldFiness") = 1
        ADC("InHPlFiness") = 1
        ADC("InHSlFiness") = 1
        ADC("InHPdFiness") = 1    ' **** Manali 3.03 - 21/06/08 - Pd Mod
        
        'Sachin - 4.1.0.0 - Line Below Commented
        'adc("InNotionRtYN") = "N" ' **** Manali 3.6.0 - 30/10/09 - For DTA Use Notional Rate in Custom Matching
      Else
        ADC("InHGldFiness") = IIF(wRsHead!hGldFiness <> 0, wRsHead!hGldFiness, 1)
        ADC("InHPlFiness") = IIF(wRsHead!hPlFiness <> 0, wRsHead!hPlFiness, 1)
        ADC("InHSlFiness") = IIF(wRsHead!hSlFiness <> 0, wRsHead!hSlFiness, 1)
        ADC("InHPdFiness") = IIF(wRsHead!hPdFiness <> 0, wRsHead!hPdFiness, 1)    ' **** Manali 3.03 - 21/06/08 - Pd Mod
        
        'Sachin - 4.1.0.0 - Line Below Commented
        'adc("InNotionRtYN") = wRsHead!hInNotionRtYN   ' **** Manali 3.6.0 - 30/10/09 - For DTA Use Notional Rate in Custom Matching
      End If
      '*** (Jen 2.14 Next)
      
      'Sachin - 4.1.0.0 - Notional Rate Applicability irrespective of DTA or Normal Invoice
      ADC("InNotionRtYN") = wRsHead!hInNotionRtYN
      
      ' ****** Sachin 2.12 - 06/12/05 - SJM
      ADC("InHGjepcNo") = wRsHead!hGjepcNo
      ADC("InHGjepcValidUpTo") = wRsHead!hGjepcValidUpTo
      ADC("InHRbiPolicy") = wRsHead!HRbiPolicy
      '*********Geeta******Emr208*****
      
      '****** Sachin 2.13.0 - 23-05-2006 ******
      'gst
      If mf_GenExpNo = True Then
          ADC("InExpNo") = ms_InCoCd + "/" + ADC("InTc") + "/" + CStr(ADC("InYy")) + "/" + ADC("InChr") + "/" + CStr(ADC("InNo"))
      End If
      '****** Sachin 2.13.0 - 23-05-2006 ******
      
      '****** Sachin 3.3.0 (23-06-08) - [14. Cost Markup On Diamonds/Cs]
      ADC("InIRtWrtStkYN") = wRsHead!HInIRtWrtStkYN
      '****** Sachin 3.3.0 (23-06-08) - [14. Cost Markup On Diamonds/Cs]
      
      '****** Sachin 3.11.0 - [Cost Markup On Metal]
      ADC("InMetIRtwrtStkYN") = wRsHead!hMetIRtwrtStkYN
      ADC("InRmRtfrRtChrt") = "N": ADC("InLabRtfrRtChrt") = "N": ADC("InOrdMulbyYN") = "N"
      '****** Sachin 3.11.0 - [Cost Markup On Metal]
      
      ' ***** Manali 3.6.0 - 27/10/09 - Invoicing with Markup on Cost of Accessories
      ADC("InXIRtWrtStkYN") = wRsHead!HInXIRtWrtStkYN
      ' ***** Manali 3.6.0 - 27/10/09 - Invoicing with Markup on Cost of Accessories
      
      ADC("InInvGrsWtEqActual") = wRsHead!HInInvGrsWtEqActual   '*** (Jen 3.01.0_1)
      
      ADC("InHHideCommYN") = wRsHead!HHideCommYN    ' ***** Manali 3.5.0 - 28/11/08
      
      ' ***** Manali 3.8.1 - Self Payable Option In Invoice
      ADC("InHSelfDscYN") = wRsHead!HSelfDscYN
      ADC("InHSelfDeductYN") = wRsHead!HSelfDeductYN
      ' ***** Manali 3.8.1 - Self Payable Option In Invoice
      
      ADC("InGspAdd1") = wRsHead!hInGspAdd1
      ADC("InGspAdd2") = wRsHead!hInGspAdd2
      ADC("InGspAdd3") = wRsHead!hInGspAdd3
      ADC("InCircularNo") = wRsHead!hInCircularNo
      ADC("InRegNo") = wRsHead!hInRegNo
      ADC("InTermCond") = IIF(ms_InTc = "JMI", wRsHead!HInMemoTermCond, wRsHead!HInInvTermCond)    '  Manali Trading Module
      
    End If
    
    ADC("InShowGldLs") = "Y"    '*** (Jen 2.13)
    ADC("InDsgOnHld") = "N"       '*** (Jen 3.01.0_1)
    
  '*** Jay 2.14(LclInv) ***
'*** Bef 2.14(LclInv)
    'Set wRsSelfCustMst = moCn.OpenRes("Select * From CustMst where CmCtg= 'C' and CmCd= '" + ctSelfCmCd + "'")
'    Set wRsSelfCustMst = moCn.OpenRes("Select * From CustMst where CmCtg= 'Z' and CmCd= '" + ms_InCoCd  + "'")
'    If Not (wRsSelfCustMst.EOF Or wRsSelfCustMst.BOF) Then
'      ADC("InSelfName") = wRsSelfCustMst!CmName
'      ADC("InSelfAdd1") = wRsSelfCustMst!CmAdd1
'      ADC("InSelfAdd2") = wRsSelfCustMst!CmAdd2
'      ADC("InSelfAdd3") = wRsSelfCustMst!CmAdd3
'      ADC("InSelfTel") = wRsSelfCustMst!CmTel
'      ADC("InSelfFax") = wRsSelfCustMst!CmFax
'      ADC("InSelfEMail") = wRsSelfCustMst!CmEMail
'    End If
'*** Bef 2.14(LclInv)
      Call BillCoDets(ms_InCoCd)
  '*** Jay 2.14(LclInv) ***
  '*** (Jen 2.13)
    'pg.11- getting default Gst code. it seach GstCode in vParam table for given voucher Tc & chr.
    'if no record then look for given Tc with blank chr. if no record then set as "-" as default code
    Dim wGstTcChr As String
    wGstTcChr = moCn.GetFldVal("Select vPValue from vParam where " + _
                                       "vPCoCd= '" + ms_InCoCd + "' and vPTyp= 'GSTTCCHR' " + _
                                       "and vPMCd= '" + ADC("InTc") + "' and vPSCd= '" + ADC("InChr") + "'")
    If wGstTcChr = "" Then _
    wGstTcChr = moCn.GetFldVal("Select vPValue from vParam where " + _
                                       "vPCoCd= '" + ms_InCoCd + "' and vPTyp= 'GSTTCCHR' " + _
                                         "and vPMCd= '" + ADC("InTc") + "' and vPSCd= ''")
    If wGstTcChr <> "" Then
      ADC("InGstCd") = wGstTcChr
      Call InvAdditionalAmtDetsFill(ADC("InGstCd"))
    Else
      ADC("InGstCd") = "-"
      Call InvAdditionalAmtDetsFill(ADC("InGstCd"))
    End If
  Else
    ms_CmCurCd = moCn.GetFldVal("Select CmCurCd from CustMst where CmCtg= '" + ms_CmCtg + "' and CmCd= '" + ADC("InCmCd") + "'")
  End If
  
  'mf_DtLocked = IIF(adc("InDt") < CDate(ms_LockDt), True, False)
    
  '6.1_50 if date is blank user cant enter anything in this field. to avoid this adc("InDt")<>"" added
  ' and user can enter date like other entries
   mf_DtLocked = IIF((ADC("InDt") < CDate(ms_LockDt)) _
                               Or _
                   (ADC("InDt") > CDate(ms_FutureDt)), _
                True, False)
  Dim i As Single
  For i = 0 To TabInvDet.Tabs - 1
    FraTabInvDet(i).Enabled = False
  Next i
  
  FraTabInvDet(TabInvDet.Tab).Enabled = True
  
' ***** Manali 3.4.1 - 19/12/08 - Following Code is Shifted to top
'   '*** Jay 3.3(FG)
'   '****** Sachin 3.02 28-12-07 - Current Partition
'  If UCase(ms_TcTyp) <> UCase(ctFotTcTyp) Then
'    wDsgExists = moCn.RecSeek("Select IdTc from InvDsg where " + _
'                           "IdCoCd= '" + ms_InCoCd  + "' and IdTc= '" + ADC("InTc") + "' and " + _
'                           "IdYy= '" + ADC("InYy") + "' and IdChr= '" + ADC("InChr") + "' and " + _
'                           "IdNo= " + CStr(ADC("InNo")) + " and IdPrtKey=(Case When '" + gs_Partition + "'='" + ctCurrPrtn + "' Then '" + ctCurrPrtn + "' Else IdPrtKey End) ")
'  Else
'     wDsgExists = moCn.RecSeek("Select 'x' from OrdMst where " + _
'                            "OmCoCd= '" + ms_InCoCd  + "' and OmTc= '" + ADC("InTc") + "' and " + _
'                            "OmYy= '" + ADC("InYy") + "' and OmChr= '" + ADC("InChr") + "' and " + _
'                            "OmNo= " + CStr(ADC("InNo")) + " and OmPrtKey=(Case When '" + gs_Partition + "'='" + ctCurrPrtn + "' Then '" + ctCurrPrtn + "' Else OmPrtKey End) ")
'     If wDsgExists Then
'       ADC("InCmCd").Enabled = False
'       ADC("InFrRmLoc").Enabled = False: ADC("InToRmLoc").Enabled = False:  ADC("wInFrRmDc").Enabled = False
'     Else
'       ADC("InCmCd").Enabled = True
'       ADC("InFrRmLoc").Enabled = True: ADC("InToRmLoc").Enabled = True:  ADC("wInFrRmDc").Enabled = True
'     End If
'   End If
'   '*** Jay 3.3(FG)

  '**************** urmi 5Jul 03*********************************************************'
  If ADC("InFrRmDc") = "D" Then
    ADC("wInFrRmDc") = "R"
  ElseIf ADC("InFrRmDc") = "C" Then
    ADC("wInFrRmDc") = "I"
    Else
      ADC("wInFrRmDc") = ""
  End If
   '**************** urmi 5Jul 03*********************************************************'
  Set wRsHead = Nothing
  '****************Geeta******Emr206*****02/04/04*********
'*** Jay 2.14 *** (SEO Change)
  If ADC("InCls") = "Y" Or mf_DtLocked Then
    txtPwd.Enabled = True
    ms_PwdFlag = False
    ADC.AllowDelete = False
  Else
    txtPwd.Enabled = False
    ms_PwdFlag = True
    ADC.AllowDelete = True
  End If
  ms_PwdFlag = moCn.RecSeek("Select PValue1 from Param where PTyp= 'SEOPWD' and PmCd= 'SEOPWD' and PValue1= '' ")
 
  
  
  ADC("InRtFrInvCust").Enabled = mb_RtfrRtChrt          '4.1.2
  
'  If ADC("InCls") = "Y" Then
'    ADC.AllowDelete = False
'  Else
'    ADC.AllowDelete = True
'  End If
'*** Jay 2.14 *** (SEO Change)
If UCase(ADC.MenuCd) = UCase("In") Or UCase(ADC.MenuCd) = UCase("InvHdFot") Then
  ADC("InCmCtg") = ms_CmCtg
End If
If (ADC("InTc") = "JOT" Or ADC("InTc") = "JIN" Or ADC("InTc") = "LOT" Or ADC("InTc") = "LIN" Or ADC("InTc") = "JOT" Or ADC("InTc") = "JIN" Or ADC("InTc") = "JMV" Or ADC("InTc") = "JCN" Or ADC("InTc") = "JMT" Or ADC("InTc") = "LMV" Or ADC("InTc") = "LCN") Then
  ms_CmCtg = "Z"
  ADC("InCmCtg") = ms_CmCtg
End If

If ms_TcTyp = "JT" Then
    Dim wRsOd As MDORowSet
    Dim wStrIns As String
    'indigo.160- company code changed from gs_CoCd to ms_InCoCd
    wStrIns = "Select * from OrdMst Where OmCoCd='" + ms_InCoCd + "' and OmTc='" + ADC("InTc") + "' " + _
              " and OmYy='" + ADC("InYy") + "' and OmChr='" + ADC("InChr") + "' and OmNo='" + CStr(ADC("InNo")) + "' " + _
              " and OmPrtKey=(Case When '" + gs_Partition + "'='" + ctCurrPrtn + "' Then '" + ctCurrPrtn + "' Else OmPrtKey End) "
    Set wRsOd = moCn.OpenRes(wStrIns)
    With wRsOd
    If Not (.EOF Or .BOF) Then
          ms_CmCtg = !OmCmCtg
          ADC("InCmCtg") = !OmCmCtg
          ADC("InCmCd") = !OmCmCd
          ADC("InDt") = !OmDt
          ADC("InLmgCst") = !OmLmgCst
          ADC("InLmgDt") = !OmLmgDt
          ADC("InLmgSal") = !OmLmgSal
          ADC("InLmpCst") = !OmLmpCst
          ADC("InLmpDt") = !OmLmpDt
          ADC("InLmpSal") = !OmLmpSal
          ADC("InLmsCst") = !OmLmsCst
          ADC("InLmsDt") = !OmLmsDt
          ADC("InLmsSal") = !OmLmsSal
          ADC("InLmlCst") = !OmLmlCst
          ADC("InLmlDt") = !OmLmlDt
          ADC("InLmlSal") = !OmLmlSal
          ADC("InFrRmDc") = !OmFrDmDc
          ADC("InToRmDc") = !OmToDmDc
          ADC("WINFRRMDC") = IIF(!OmFrDmDc = "C", "I", IIF(!OmFrDmDc = "D", "R", ""))
          ADC("InFrRmLoc") = !OmFrDmLoc
          ADC("InToRmLoc") = !OmToDmLoc
          
    End If
  End With
End If
'gst
'indigo.160- company code changed from gs_CoCd to ms_InCoCd
If ms_TcTyp = "SO" Then
    wStrIns = "Select * from OrdMst Where OmCoCd='" + ms_InCoCd + "' and OmTc='" + ADC("InTc") + "' " + _
              " and OmYy='" + ADC("InYy") + "' and OmChr='" + ADC("InChr") + "' and OmNo='" + CStr(ADC("InNo")) + "' " + _
              " and OmPrtKey=(Case When '" + gs_Partition + "'='" + ctCurrPrtn + "' Then '" + ctCurrPrtn + "' Else OmPrtKey End) "
    Set wRsOd = moCn.OpenRes(wStrIns)
    With wRsOd
    If Not (.EOF Or .BOF) Then
          ms_CmCtg = !OmCmCtg
          ADC("InCmCtg") = !OmCmCtg
          ADC("InCmCd") = !OmCmCd
          ADC("InDt") = !OmDt
          ADC("InLmgCst") = !OmLmgCst
          ADC("InLmgDt") = !OmLmgDt
          ADC("InLmgSal") = !OmLmgSal
          ADC("InLmpCst") = !OmLmpCst
          ADC("InLmpDt") = !OmLmpDt
          ADC("InLmpSal") = !OmLmpSal
          ADC("InLmsCst") = !OmLmsCst
          ADC("InLmsDt") = !OmLmsDt
          ADC("InLmsSal") = !OmLmsSal
          ADC("InLmlCst") = !OmLmlCst
          ADC("InLmlDt") = !OmLmlDt
          ADC("InLmlSal") = !OmLmlSal
          ADC("InFrRmDc") = !OmFrDmDc
          ADC("InToRmDc") = !OmToDmDc
          ADC("WINFRRMDC") = IIF(!OmFrDmDc = "C", "I", IIF(!OmFrDmDc = "D", "R", ""))
          ADC("InFrRmLoc") = !OmFrDmLoc
          ADC("InToRmLoc") = !OmToDmLoc
    End If
  End With
End If
'indigo.160- company code changed from gs_CoCd to ms_InCoCd
If ms_TcTyp = "PR" Or ms_TcTyp = "CNV" Or ms_TcTyp = "MV" Or ms_TcTyp = "DT" Or ms_TcTyp = "BV" Or ms_TcTyp = "BS" Or ms_TcTyp = "BM" Or ms_TcTyp = "MLT" Or ms_TcTyp = "REJ" Then
    Dim wRsTxn As MDORowSet
    Set wRsTxn = moCn.OpenRes("Select * from Txn Where TCoCd='" + ms_InCoCd + "' and TTc='" + ADC("InTc") + "' " + _
              " and TYy='" + ADC("InYy") + "' and TChr='" + ADC("InChr") + "' and TNo='" + CStr(ADC("InNo")) + "' " + _
              " and TPrtKey=(Case When '" + gs_Partition + "'='" + ctCurrPrtn + "' Then '" + ctCurrPrtn + "' Else TPrtKey End) ")
    If Not (wRsTxn.EOF Or wRsTxn.BOF) Then
      ms_CmCtg = IIF(ADC("InTc") = "PR", "S", "C")
      ADC("InCmCd") = IIF(ADC("InTc") = "PR" Or ADC("InTc") = "MV" Or ADC("InTc") = "SA" Or ADC("InTc") = "CNV", wRsTxn!TSuppCd, wRsTxn!TCmCd)
      If ADC("InCmCd") = "" Then
        ms_CmCtg = "Z"
        
        ADC("InCmCd") = ms_InCoCd
      End If
      ADC("InCmCtg") = ms_CmCtg
      ADC("InDt") = wRsTxn!TDt
      ADC("InFrRmDc") = "C"
      ADC("InToRmDc") = "D"
      ADC("WINFRRMDC") = "I"
      ADC("InFrRmLoc") = "XC"
      ADC("InToRmLoc") = "XC"
    End If
End If
'indigo.160- company code changed from gs_CoCd to ms_InCoCd
If ms_TcTyp = "FB" Or ms_TcTyp = "FWB" Then
    ADC("InCmCd") = ms_InCoCd
    ms_CmCtg = "Z"
    ADC("InCmCtg") = ms_CmCtg
    ADC("InDt") = moCn.GetFldVal("Select FgDt from Fg Where FgCoCd='" + ms_InCoCd + "' and FgTc='" + ADC("InTc") + "' " + _
            " and FgYy='" + ADC("InYy") + "' and FgChr='" + ADC("InChr") + "' and FgNo='" + CStr(ADC("InNo")) + "' " + _
            " and FgPrtKey=(Case When '" + gs_Partition + "'='" + ctCurrPrtn + "' Then '" + ctCurrPrtn + "' Else FgPrtKey End) ")
    
    ADC("InFrRmDc") = "C"
    ADC("InToRmDc") = "D"
    ADC("WINFRRMDC") = "I"
    ADC("InFrRmLoc") = "XC"
    ADC("InToRmLoc") = "XC"
'    End If
End If
'pg.11 custom purchase details added
If ms_TcTyp = "TZ" Then
    Dim wRsTxnZ As MDORowSet
    Set wRsTxnZ = moCn.OpenRes("Select * from Txn Where TCoCd='" + ms_InCoCd + "' and TTc='" + ADC("InTc") + "' " + _
              " and TYy='" + ADC("InYy") + "' and TChr='" + ADC("InChr") + "' and TNo='" + CStr(ADC("InNo")) + "' " + _
              " and TPrtKey=(Case When '" + gs_Partition + "'='" + ctCurrPrtn + "' Then '" + ctCurrPrtn + "' Else TPrtKey End) ")
    If Not (wRsTxnZ.EOF Or wRsTxnZ.BOF) Then
      ms_CmCtg = "S"
      ADC("InCmCd") = wRsTxnZ!TSuppCd
      ADC("InCmCtg") = ms_CmCtg
      ADC("InDt") = wRsTxnZ!TDt
      ADC("InFrRmDc") = "C"
      ADC("InToRmDc") = "D"
      ADC("WINFRRMDC") = "I"
      ADC("InFrRmLoc") = "XC"
      ADC("InToRmLoc") = "XC"
    End If
End If
    'Dim wGstTcChr As String
   ' wGstTcChr = moCn.GetFldVal("Select vPValue from vParam where " + _
                                       "vPCoCd= '" + ms_InCoCd + "' and vPTyp= 'GSTTCCHR' " + _
                                       "and vPMCd= '" + ADC("InTc") + "' and vPSCd= '" + ADC("InChr") + "'")
   ' If wGstTcChr = "" Then _
    'wGstTcChr = moCn.GetFldVal("Select vPValue from vParam where " + _
                                       "vPCoCd= '" + ms_InCoCd + "' and vPTyp= 'GSTTCCHR' " + _
                                         "and vPMCd= '" + ADC("InTc") + "' and vPSCd= ''")
   ' If wGstTcChr <> "" Then
   '   ADC("InGstCd") = wGstTcChr
   '   Call InvAdditionalAmtDetsFill(ADC("InGstCd"))
   ' Else
   '   ADC("InGstCd") = "-"
   '   Call InvAdditionalAmtDetsFill(ADC("InGstCd"))
   ' End If
 ' Else
  '  ms_CmCurCd = moCn.GetFldVal("Select CmCurCd from CustMst where CmCtg= '" + ms_CmCtg + "' and CmCd= '" + ADC("InCmCd") + "'")
  'End If
If ms_TcTyp = "JT" And ADC.AddFndMode = xFndMode Then
Call InvAdditionalAmtsReCalc(True)
End If
End Sub

Private Sub atxt_FldValid(Index As Integer, ByVal IdName As String, ByVal OldVal As Variant)
 If ADC(IdName) = OldVal Then Exit Sub
 Select Case UCase(IdName)
 'IG 91- If Paytermcode is not blank & its description is blank then will het the description from Param else no change
  Case Is = UCase("wPayTermsCd")
    Dim wPDesc As String
    wPDesc = moCn.GetFldVal("Select PDesc225 from Param where PTyp = 'PYMTRM' and PmCd= '" + ADC("WPAYTERMSCD") + "' and PSCd= ''")
    If wPDesc <> "" Then ADC("InCmPayTerms") = wPDesc
    Exit Sub
  Case Is = UCase("InGstCd")
    If (Not moCn.RecSeek("Select PMCd from Param where PTyp= 'GSTFLDS' And " + _
                         " PMCd = '" + ADC("InGstCd") + "'")) Then
      MsgBox ("There are no GST fields definition for this given GST Cd. All Amount values will be cleared. Please make sure before you save this record")
    End If
    Call InvAdditionalAmtDetsFill(ADC("InGstCd"))
  Case Is = UCase("InBaseAmt"), UCase("InAmt1Per"), UCase("InAmt1"), UCase("InAmt2Per"), UCase("InAmt2"), UCase("InAmt3Per"), UCase("InAmt3"), _
  UCase("InAmt4Per"), UCase("InAmt4"), UCase("InAmt5Per"), UCase("InAmt5"), UCase("InAmt6Per"), UCase("InAmt6"), UCase("InAmt7Per"), UCase("InAmt7"), _
  UCase("InAmt8Per"), UCase("InAmt8"), UCase("InAmt9Per"), UCase("InAmt9"), UCase("InAmt10Per"), UCase("InAmt10"), UCase("InAmt11Per"), UCase("InAmt11"), _
  UCase("InAmt12Per"), UCase("InAmt12")
    Call InvAdditionalAmtsReCalc(False)
  End Select
End Sub



Private Sub ATXT_FldWhen(Index As Integer, ByVal IdName As String, _
                          Cancel As Boolean, ErrMsg As String)
  '*** Do not allow the user to edit the voucher date, Customer, Lmg Cost, Lmg Date,
  '    Lmg Sales, Value Addn % for DC & X  if record exists in InvDsg
  '*** Do not allow the user to edit the fields
  '    InExpNo, InCmCd, InDt,InLmgCst, InLmgSal, InVaDc, InVaX, InCnvRt, InFOB$, InFOBRs,
  '    InComm$, InCommRs, InFrt$, InFrtRs, InIns$, InInsRs, InTot$, InTotRs, InTotAs,
  '    InDiaCost, InCsCost, InAdv when InCls= 'Y'
  '*** Give help for InYy, InChr, InNo, InCmCd, InCls, InTotAs
  
  '****** Sachin 3.02 28-12-07 - Locking the Previous Partition Entries
  If UCase(IdName) <> UCase("InCoCd") And UCase(IdName) <> UCase("InTc") And UCase(IdName) <> UCase("InYy") _
    And UCase(IdName) <> UCase("InChr") And UCase(IdName) <> UCase("InNo") _
    And ADC.AddFndMode = xFndMode And (ADC("InPrtKey") <> ctCurrPrtn Or mf_DtLocked) Then _
    Cancel = True: ErrMsg = "Cannot Edit Previous Partition Entry": Exit Sub
  
  If mf_DtLocked = True Then
    Cancel = True: ErrMsg = "Cannot Edit Locked Entry": Exit Sub
  End If
' ########################  Manoj #### Ver: 2.0.6 #### Date: 25/12/2003  ########################
  If UCase$(ADC("InCls")) = UCase$("Y") And UCase$(IdName) <> UCase$("InCls") Then
    Cancel = True: ErrMsg = "Cannot Edit Locked Entry": Exit Sub
  End If
' ###########################################  Manoj  ###########################################
  
  Select Case UCase(IdName)
  'indigo.160- calling help for company code
  Case Is = UCase("InCoCd")
    Call HlpList.hCoCd
  Case Is = UCase("InYy")
    Call HlpList.PMCd("YY")
  Case Is = UCase("InChr")
' **** Zubin 211 **** '
    'Call HlpList.vPSCd(ms_InCoCd , "CHR", ADC("InTc"))
    Call HlpList.vPSCd(ms_InCoCd, "CHR", ADC("InTc"), "Y")
' **** Zubin 211 **** '
  Case Is = UCase("InNo")
    Call HlpList.InNo(ms_InCoCd, ADC("InTc"), ADC("InYy"), ADC("InChr"))
  Case Is = UCase("InDt")
    If ADC("InCls") = "Y" Then Cancel = True: ErrMsg = "Cannot Edit As Invoice Is Closed": Exit Sub
    If wDsgExists = True Then Cancel = True: ErrMsg = "Cannot Edit As Inv Design Exists": Exit Sub
  Case Is = UCase("InPrnDt")
    If ADC("InCls") = "Y" Then Cancel = True: ErrMsg = "Cannot Edit As Invoice Is Closed": Exit Sub
  Case Is = UCase("InExpNo")
    If ADC("InCls") = "Y" Then Cancel = True: ErrMsg = "Cannot Edit As Invoice Is Closed": Exit Sub
  Case Is = UCase("InCmCd")
    If ADC("InCls") = "Y" Then Cancel = True: ErrMsg = "Cannot Edit As Invoice Is Closed": Exit Sub
    If Trim(CStr(ADC("InFOB"))) = "0.00" Then Cancel = True: ErrMsg = "Cannot Edit As Invoice FOB Value Is Non Zero": Exit Sub      '*** Jay 2.14(LclInv) ***
    If UCase(ms_TcTyp) = UCase(ctCTBTcTyp) And wDsgExists = True Then Cancel = True: ErrMsg = "Cannot Edit As Inv Design Exists": Exit Sub
    If (mb_InvHdGST Or ms_TcTyp = "JT") Then Cancel = True: ErrMsg = "Cannot Edit Customer Code for GST Header": Exit Sub
    Call HlpList.CustCd(ms_CmCtg)
  Case Is = UCase("InCls")
' ###################################  Manoj #### Ver: 2.0.6  ###################################
' Date:   20/01/2004
    If ADC("InCls") = "Y" Then
      '*** Jay 2.14 *** (SEO Change)
      If ms_PwdFlag = False And Trim$(txtPwd.text) = "" Then _
        Cancel = True: ErrMsg = "Enter SEO Password to Edit Value": Exit Sub
      Call HlpList.PMCd("YN")
      'Cancel = Not moCn.RecSeek("Select 'x' from Param where " + _
                           "Ptyp = 'SEOPWD' and PmCd = 'SEOPWD' And PValue1='" & TxtPwd.text & "'")
      'If Cancel = True Then ErrMsg = "Enter SEO Password for Opening Closed Invoice."
      Exit Sub
      '*** Jay 2.14 *** (SEO Change)
    End If
' ###########################################  Manoj  ###########################################
    Call HlpList.PMCd("YN")
    '*********Geeta**********Emr207****Check for UCase("InFOBPer")
    '****Geeta****Emr208***UCase("InHGldFiness"), UCase("InHPlFiness")
  ' ***** Manali 3.03 - 21/06/08 - Pd Mod - InLmlCst, InLmlDt, InLmlSal, InHPdFiness added
  Case Is = UCase("InLmgCst"), UCase("InLmgDt"), UCase("InLmgSal"), UCase("InLmpCst"), _
            UCase("InLmpDt"), UCase("InLmpSal"), UCase("InLmsCst"), UCase("InLmsDt"), UCase("InLmsSal"), _
            UCase("InFOBPer"), UCase("InHGldFiness"), UCase("InHPlFiness"), UCase("InHSlFiness"), _
            UCase("InLmlCst"), UCase("InLmlDt"), UCase("InLmlSal"), UCase("InHPdFiness")
            ' Sachin 2.12 - 02/12/05 - SJM
    If ADC("InCls") = "Y" Then Cancel = True: ErrMsg = "Cannot Edit As Invoice Is Closed": Exit Sub
    If wDsgExists = True Then Cancel = True: ErrMsg = "Cannot Edit As Inv Design Exists": Exit Sub
  Case Is = UCase("InVaDC"), UCase("InVaX")
    If ADC("InCls") = "Y" Then Cancel = True: ErrMsg = "Cannot Edit As Invoice Is Closed": Exit Sub
    If wDsgExists = True Then Cancel = True: ErrMsg = "Cannot Edit As Inv Design Exists": Exit Sub
  Case Is = UCase("InTotAs")
    If ADC("InCls") = "Y" Then Cancel = True: ErrMsg = "Cannot Edit As Invoice Is Closed": Exit Sub
    Call HlpList.PMCd("INVTOTAS")
  Case Is = UCase("InCnvRt")
    If ADC("InCls") = "Y" Then Cancel = True: ErrMsg = "Cannot Edit As Invoice Is Closed": Exit Sub
    If wDsgExists = True And UCase(ms_TcTyp) = UCase(ctCTBTcTyp) Then Cancel = True: ErrMsg = "Cannot Edit As Inv Design Exists": Exit Sub
    ' **** Manali 3.6.0 - 30/10/09 - For DTA Use Notional Rate in Custom Matching
    If mb_IzRmExists And UCase(ms_TcTyp) = UCase(ctInTcTyp) And ADC("InNotionRtYn") = "Y" Then Cancel = True: ErrMsg = "Cannot Edit as Records exist for Custom Matching": Exit Sub
    ' **** Manali 3.6.0 - 30/10/09 - For DTA Use Notional Rate in Custom Matching
  
  '*** (Jen 2.13)
  Case Is = UCase("InCnvFct")
    If ADC("InCls") = "Y" Then Cancel = True: ErrMsg = "Cannot Edit As Invoice Is Closed": Exit Sub
    If wDsgExists = True Then Cancel = True: ErrMsg = "Cannot Edit As Inv Design Exists": Exit Sub
  '*** (Jen 2.13)
  
  '*** Jay 3.1.1 [Discount] (Added Flds: InDsc, InDscRs, InDscDesc)
  ' ***** Manali - 3.8.1 - Self Payable Options Added - InDeduct, InDeductRs, InDeductDesc added
  'sv.44 InTCS, InTCSRs added
  Case Is = UCase("InFOB"), UCase("InComm"), UCase("InFrt"), UCase("InIns"), _
            UCase("InCommRs"), UCase("InFrtRs"), UCase("InInsRs"), UCase("InDiaCst"), _
            UCase("InCSCost"), UCase("InAdv"), UCase("InDsc"), UCase("InDscRs"), UCase("InDscDesc"), _
            UCase("InDeduct"), UCase("InDeductRs"), UCase("InDeductDesc"), UCase("InTCS"), UCase("InTCSRs")
    If ADC("InCls") = "Y" Then Cancel = True: ErrMsg = "Cannot Edit As Invoice Is Closed": Exit Sub
  Case Is = UCase("wBkCd")
    If ADC("InCls") = "Y" Then Cancel = True: ErrMsg = "Cannot Edit As Invoice Is Closed": Exit Sub
    Call HlpList.BkCd("B")
  Case Is = UCase("InHRndOffFOB")
    If ADC("InCls") = "Y" Then Cancel = True: ErrMsg = "Cannot Edit As Invoice Is Closed": Exit Sub
    Call HlpList.PMCd("YN")
  Case Is = UCase("InCmName")
    If ADC("InCls") = "Y" Then Cancel = True: ErrMsg = "Cannot Edit As Invoice Is Closed": Exit Sub
    If UCase(ms_TcTyp) = UCase(ctCTBTcTyp) And wDsgExists = True Then Cancel = True: ErrMsg = "Cannot Edit As Inv Design Exists": Exit Sub
    Call HlpList.CmName(ms_CmCtg)
  Case Is = UCase("WINVPRDDESC")
    Call HlpList.PMCd("INVDESC")
    
' ****** Manali 3.03 - 16/08/08 - Agent Code, Name and Address ******************************
  Case Is = UCase("wAgCd")
    Call HlpList.BkCd("A")
' ****** Manali 3.03 - 16/08/08 - Agent Code, Name and Address ******************************
    
'********************* urmi 4Jul03 ******************************************************
  Case Is = UCase("InFrRmLoc")
    If UCase(ms_TcTyp) = UCase(ctCTBTcTyp) Then Cancel = True: ErrMsg = "Cannot Edit This Field In CTB": Exit Sub
    '*** Jay 3.3(FG)
    If wDsgExists Then Cancel = True: ErrMsg = "Cannot Edit As Inv Design Exists": Exit Sub
    '*** Jay 3.3(FG)
    Call HlpList.Loc(ms_InCoCd, "'F', 'XS'")
  Case Is = UCase("InToRmLoc")
    If UCase(ms_TcTyp) = UCase(ctCTBTcTyp) Then Cancel = True: ErrMsg = "Cannot Edit This Field In CTB": Exit Sub
    '*** Jay 3.3(FG)
    If wDsgExists Then Cancel = True: ErrMsg = "Cannot Edit As Inv Design Exists": Exit Sub
    '*** Jay 3.3(FG)
    Call HlpList.Loc(ms_InCoCd, "'F', 'XS'")
  Case Is = UCase("wInFrRmDc")
    If UCase(ms_TcTyp) = UCase(ctCTBTcTyp) Then Cancel = True: ErrMsg = "Cannot Edit This Field In CTB": Exit Sub
    '*** Jay 3.3(FG)
    If wDsgExists Then Cancel = True: ErrMsg = "Cannot Edit As Inv Design Exists": Exit Sub
    '*** Jay 3.3(FG)
    Call HlpList.PMCd("IR")
'********************* urmi 4Jul03 ******************************************************
'***** Shilpa ******
  Case Is = UCase("InhInvTyp")
    If ADC("InCls") = "Y" Then Cancel = True: ErrMsg = "Cannot Edit As Invoice Is Closed": Exit Sub
    Call HlpList.PMCd("InvTyp")
  Case Is = UCase("InHValAddMtlColYN"), UCase("InHValBOEDetYN")
    Call HlpList.PMCd("YN")
  Case Is = UCase("InComPer")
    If ADC("InCls") = "Y" Then Cancel = True: ErrMsg = "Cannot Edit As Invoice Is Closed": Exit Sub
'***** Shilpa ******
 
  '****** (Jen 3.01.0_1 added InDsgOnHld, UCase("InInvGrsWtEqActual")) ******
  '*** (Jen 2.13)
  Case Is = UCase("InShowGldLs"), UCase("InDsgOnHld"), UCase("InInvGrsWtEqActual")
    If wDsgExists = True And UCase(IdName) = UCase("InInvGrsWtEqActual") Then Cancel = True: ErrMsg = "Cannot Edit As Inv Design Exists": Exit Sub
    Call HlpList.PMCd("YN")
  '*** (Jen 2.13)
  '*** Jay 2.14(LclInv) ***
  Case Is = UCase("InLclInvYn")
    Call HlpList.PMCd("YN")
  '*** Jay 2.14(LclInv) ***
  '*** (Jen 3.01.0_1)
  Case Is = UCase("InRemakingCtgs")
    Call HlpList.PMCd("REMAKE")
  '*** (Jen 3.01.0_1)
  '****** Sachin 3.3.0 (23-06-08) - [42. Make Invoice By Bag] ******  'Sachin 3.11.0 - Metal Included
  ' **** Manali 3.6.0 - 27/10/09 - Invoicing with Markup on Cost of Accessories
  Case Is = UCase("InByBagYN"), UCase("InIRtWrtStkYN"), UCase("InXIRtWrtStkYN"), UCase("InMetIRtWrtStkYN"), UCase("InRmRtfrRtChrt"), UCase("InLabRtfrRtChrt"), UCase("InOrdMulbyYN")    '3.11.0
    If wDsgExists = True Then Cancel = True: ErrMsg = "Cannot Edit As Inv Design Exists": Exit Sub
    Call HlpList.PMCd("YN")
  '****** Sachin 3.3.0 (23-06-08) - [42. Make Invoice By Bag] ******]
  
  ' ***** Manali 3.5.0 - 28/11/08 - Hide Commission added
  ' ***** Manali 3.8.1 - Self Payable Option In Invoice
  Case Is = UCase("InHHideCommYN"), UCase("InHSelfDscYN"), UCase("InHSelfDeductYN")
    If ADC("InCls") = "Y" Then Cancel = True: ErrMsg = "Cannot Edit as Invoice is Closed": Exit Sub
    Call HlpList.PMCd("YN")
  ' ***** Manali 3.5.0 - 28/11/08 - Hide Commission added

    ' **** Manali 3.6.0 - 30/10/09 - For DTA Use Notional Rate in Custom Matching
    Case Is = UCase("InNotionRtYn")
       If ADC("InCls") = "Y" Then Cancel = True: ErrMsg = "Cannot Edit as Invoice is Closed": Exit Sub
       If mb_IzRmExists Then Cancel = True: ErrMsg = "Cannot Edit as Records exist for Custom Matching": Exit Sub
      Call HlpList.PMCd("YN")
    ' **** Manali 3.6.0 - 30/10/09 - For DTA Use Notional Rate in Custom Matching
  
    Case Is = UCase("InRtFrInvCust")  '4.1.2
      Call HlpList.PMCd("YN")
  
    'Sachin 4.1.0.0
    Case Is = UCase("InCmPod")
        Call HlpList.PMCd("PORTCD")

    Case Is = UCase("InCmFinDstCd")
        Call HlpList.PMCd("CONTRYCD")
    'Sachin 4.1.0.0
   'IG 91- Calling help of Payterm code
    Case Is = UCase("wPayTermsCd")
      Call HlpList.PMCd("PYMTRM")
    'pg.11-help for Gst code
    Case Is = UCase("InGstCd")
      Call HlpList.PMCd("GSTCD")
    'pg.11-Amount fields non editable when % <>0
    Case Is = UCase("InAmt1")
      If ADC("InAmt1Per") <> 0 Then Cancel = True: ErrMsg = "Cannot Edit As Percentage1 <>0 ": Exit Sub
    Case Is = UCase("InAmt2")
      If ADC("InAmt2Per") <> 0 Then Cancel = True: ErrMsg = "Cannot Edit As Percentage2 <>0 ": Exit Sub
    Case Is = UCase("InAmt3")
      If ADC("InAmt3Per") <> 0 Then Cancel = True: ErrMsg = "Cannot Edit As Percentage3 <>0 ": Exit Sub
    Case Is = UCase("InAmt4")
      If ADC("InAmt4Per") <> 0 Then Cancel = True: ErrMsg = "Cannot Edit As Percentage4 <>0 ": Exit Sub
    Case Is = UCase("InAmt5")
      If ADC("InAmt5Per") <> 0 Then Cancel = True: ErrMsg = "Cannot Edit As Percentage5 <>0 ": Exit Sub
    Case Is = UCase("InAmt6")
      If ADC("InAmt6Per") <> 0 Then Cancel = True: ErrMsg = "Cannot Edit As Percentage6 <>0 ": Exit Sub
    Case Is = UCase("InAmt8")
      If ADC("InAmt8Per") <> 0 Then Cancel = True: ErrMsg = "Cannot Edit As Percentage8 <>0 ": Exit Sub
    Case Is = UCase("InAmt9")
      If ADC("InAmt9Per") <> 0 Then Cancel = True: ErrMsg = "Cannot Edit As Percentage9 <>0 ": Exit Sub
    Case Is = UCase("InAmt10")
      If ADC("InAmt10Per") <> 0 Then Cancel = True: ErrMsg = "Cannot Edit As Percentage10 <>0 ": Exit Sub
    Case Is = UCase("InAmt11")
      If ADC("InAmt11Per") <> 0 Then Cancel = True: ErrMsg = "Cannot Edit As Percentage11 <>0 ": Exit Sub
    'pg.11-Base amount field non editable when GstBase <> "-"
    Case Is = UCase("InBaseAmt")
      If ADC("InGstBase") <> "-" Then _
        Cancel = True: ErrMsg = "Base Amount Can Be Edited Only in '-' GstBase ": Exit Sub
    '578-20 calling help of customer store code
    Case Is = UCase("InCstCd")
      Call HlpList.CustStoreCd(ADC("InCmCtg"), ADC("InCmCd"))
    
  End Select
End Sub

Private Sub atxt_ReCalculate(Index As Integer, ByVal IdName As String, ByVal OldVal As Variant, ByVal RowNum As Integer, ByVal SourceName As String, ByVal SourceColName As String)
  '*** InCmName on InCmCd                           Normal
      'CmName from CustMst(CmCtg= 'C', CmCd= InCmCd)
  '*** InChr on InTc                                Normal
      'PDesc225('TC', adc(InTc))
  '*** InLmgCst on InDt                             Normal
      'GetDtVal("LMG", InDt, wDtVal)
      'InLmgCst= wDtVal
      '   This procedure gets DtVal from DtTbl(DtTag= 'LMG' and max(DtDt)<= InDt)
  '*** InLmgDt on InDt                              Normal
      'GetDtVal("LMG", InDt, ,wDtDt)
      'InLmgDt= wDtDt
      '   This procedure gets max(DtDt) from DtTbl(DtTag= 'LMG' and DtDt<= InDt)
  '*** InLmgSal on InCmCd, InLmgCst                 Normal
      'Check if SalRt exists in RmRt for RrTcTyp= 'LMG', RrCtg= 'G', RrFrLn<= InLmgCst,
      '     RrToLn >= InLmgCst, RrSalRt<> 0 for the specified customer.
      'If Not, Check if SalRt exists in RmRt for RrTcTyp= 'LMG', RrCtg= 'G' and RrFrLn<=
      '    InLmgCst, RrToLn >= InLmgCst, RrCmCd= InCmCd and RrSalRt<> 0 For ZSelf.
      '    Also consider the Incremental % for gold for the specified customer
      'InLmgSal= GetLmgSal(InCmCd, InLmgCst)
      'GetLmgSal = if exists (RrSalRt from RmRt(RrTcTyp = 'LMG', RrCtg= 'G',
      '                                         RrFrLn <= InLmgCst, RrToLn >= InLmgCst,
      '                                         RrCmCd= InCmCd, RrSalRt<> 0)
      '
      '            else (RrSalRt from RmRt(RrTcTyp = 'LMG', RrCtg= 'G', RrFrLn <= InLmgCst,
      '                                    RrToLn >= InLmgCst, RrCmCd= 'ZSELF', RrSalRt<> 0)
      '                 *(1+ (Select RrSalRt from RmRt(RrTcTyp = 'INC', RrCtg = 'G',
      '                   RrCmCd = InCmCd)/100)
  '*** InCnvRt on InDt                              Normal
      'GetDtVal("USD", InDt, ,wDtDt)
      'InCnvRt= wDtDt
      '   This procedure gets max(DtDt) from DtTbl(DtTag= 'LMG' and DtDt<= InDt)
  '*** InPoDt on InPoNo                             Normal      '### Not To Be Done
      '
  '*** InFOBRs on InFOB, InCnvRt                    Normal
      'If InFOB <> 0 then InFOBRs= InFOB * InCnvRt
  '*** InCommRs on InComm, InCnvRt                  Normal
      'If InComm <> 0 then InCommRs= InComm * InCnvRt
  '*** InFrtRs on InFrt, InCnvRt                    Normal
      'If InFrt <> 0 InFrtRs= InFrt * InCnvRt
  '*** InInsRs on InIns, InCnvRt                    Normal
      'If InIns <> 0 InInsRs= InIns * InCnvRt
  '*** InTot on InTotAs,InFOB,InComm,InFrt,InIns    Normal
      'InTot= InFOB + InComm + InFrt + InIns    (depending on InTotAs)
  '*** InTotRs on InTotAs,InFOBRs,InCommRs,InFrtRs,InInsRs    Normal
      'InTotRs= InFOBRs + InCommRs + InFrtRs + InInsRs    (depending on InTotAs)
  '''''''''''*** InTotRs on InTot, InCnvRt                    Normal
  ''''''''''    'InTotRs= InTot * InCnvRt
  '*** InCmAdd1 on InCmCd                           Normal
      'InCmAdd1= CmAdd1 from CustMst(CmCtg= 'C', CmCd= InCmCd)
  '*** InCmAdd2 on InCmCd                           Normal
      'InCmAdd2= CmAdd2 from CustMst(CmCtg= 'C', CmCd= InCmCd)
  '*** InCmAdd3 on InCmCd                           Normal
      'InCmAdd3= CmAdd3 from CustMst(CmCtg= 'C', CmCd= InCmCd)
  '*** InCmTel on InCmCd                            Normal
      'InCmTel= CmTel from CustMst(CmCtg= 'C', CmCd= InCmCd)
  '*** InCmFax on InCmCd                            Normal
      'InCmFax= CmFax from CustMst(CmCtg= 'C', CmCd= InCmCd)
  '*** InCmEMail on InCmCd                          Normal
      'InCmEMail= CmEMail from CustMst(CmCtg= 'C', CmCd= InCmCd)
  '*** InCmCpName on InCmCd                         Normal
      'InCmCpName= CmCpName from CustMst(CmCtg= 'C', CmCd= InCmCd)
  '*** InCmPortOfDisch on InCmCd                    Normal
      'InCmPortOfDisch= CmPortOfDisch from CustMst(CmCtg= 'C', CmCd= InCmCd)
  '*** InCmFinDest on InCmCd                        Normal
      'InCmFinDest= CmFinDest from CustMst(CmCtg= 'C', CmCd= InCmCd)
  '*** InCmBuyer on InCmCd                          Normal
      'InCmBuyer= CmBuyer from CustMst(CmCtg= 'C', CmCd= InCmCd)
  '*** InCmPayTerms on InCmCd                       Normal
      'InCmPayTerms= CmPayTerms from CustMst(CmCtg= 'C', CmCd= InCmCd)
  '*** InCmInsBy on InCmCd                           Normal
      'InCmInsBy= CmHInsBy from CustMst(CmCtg= 'C', CmCd= InCmCd)
  '*** InHBkName on wBkCd                           Normal
      'InHBkName= BkName from BnkAgMst(BkCtg='B' and BkCd= wBkCd)
  '*** InHBkAdd1 on wBkCd                           Normal
      'InHBkAdd1= BkAdd1 from Bank( BkCtg='B' and BkCd= wBkCd)
  '*** InHBkAdd2 on wBkCd                           Normal
      'InHBkAdd2= BkAdd2 from Bank(BkCtg='B' and BkCd= wBkCd)
  '*** InHBkAdd3 on wBkCd                           Normal
      'InHBkAdd3= BkAdd3 from Bank(BkCtg='B' and BkCd= wBkCd)
  '*** InHBkInstr on wBkCd                          Normal
      'InHBkInstr= BkInstr from Bank(BkCtg='B' and BkCd= wBkCd)
      
  '**********Bhavna
  '*** InAmt7 = InFOB + (InAmt1 + InAmt2 + InAmt3 + InAmt4 + InAmt5 + InAmt6)
  '*** InAmt12 = InAmt7 + InAmt8 + InAmt9 + InAmt10 + InAmt11
  '**********Bhavna
  
  Dim wCmCd As String, wCustName As String, wDtVal As Double, wDtDt As String
  Select Case UCase(IdName)
  Case Is = UCase("InCmName")
    If ADC.Mode = xNorm Then
        wCustName = moCn.GetFldVal("Select CmName from CustMst where CmCtg='" + ms_CmCtg + "' and " + _
                                  "CmCd='" + ADC("InCmCd") + "'")
        If wCustName <> ADC("InCmName") Then ADC("InCmName") = wCustName
        Call FillShpDet(ADC("InhInvTyp"))
    End If
  Case Is = UCase("InChr")
    If ADC.Mode = xNorm Then
     ADC("InChr") = moCn.GetFldVal("Select vPValue from vParam where " + _
                                "vPCoCd= '" + ms_InCoCd + "' and vPTyp='DEFCHR' " + _
                                "and vPMCd='" + ADC("InTc") + "'")
     '("Select PDesc225 from Param where PTyp='TC' " + _
                                   "and PMCd='" + ADC("InTc") + "'")
    End If
  Case Is = UCase("InLmgCst")
    If ADC.Mode = xNorm And mb_ChgInDt = False Then
      '*** (Bef 2.13) Call GetDtVal("LMG", ADC("InDt"), wDtVal)
      '*** (Jen 2.13)
      Call GetDtVal("LMG", ADC("InDt"), wDtVal, , ms_CmCurCd)
      '*** (Jen 2.13)
      ADC("InLmgCst") = wDtVal
    End If
  Case Is = UCase("InLmgDt")
    If ADC.Mode = xNorm And mb_ChgInDt = False Then
      '*** (Bef 2.13) Call GetDtVal("LMG", ADC("InDt"), , wDtDt)
      '*** (Jen 2.13)
      Call GetDtVal("LMG", ADC("InDt"), , wDtDt, ms_CmCurCd)
      '*** (Jen 2.13)
      ADC("InLmgDt") = wDtDt
    End If
  Case Is = UCase("InLmgSal")
    If ADC.Mode = xNorm And wDsgExists = False And mb_ChgInDt = False Then
      '*** (Bef 2.13) ADC("InLmgSal") = GetLmgSal("LMG", ADC("InCmCd"), ADC("InLmgCst"), ms_CmCtg)
      
      '*** (Jen 2.13)
      'Debug.Print ADC("InCmCd") + "/" + CStr(ADC("InLmgCst")) + "/" + ms_CmCtg + "/" + ms_CmCurCd
      ADC("InLmgSal") = GetLmgSal("LMG", ADC("InCmCd"), ADC("InLmgCst"), ms_CmCtg, ms_CmCurCd)
      '*** (Jen 2.13)
      
      If ADC("InLmgSal") = 0 Then ADC("InLmgSal") = ADC("InLmgCst")
    End If
  Case Is = UCase("InLmpCst")
    If ADC.Mode = xNorm And mb_ChgInDt = False Then
      '*** (Bef 2.13) Call GetDtVal("LMP", ADC("InDt"), wDtVal)
      '*** (Jen 2.13)
      Call GetDtVal("LMP", ADC("InDt"), wDtVal, , ms_CmCurCd)
      '*** (Jen 2.13)
      ADC("InLmpCst") = wDtVal
    End If
  Case Is = UCase("InLmpDt")
    If ADC.Mode = xNorm And mb_ChgInDt = False Then
      '*** (Bef 2.13) Call GetDtVal("LMP", ADC("InDt"), , wDtDt)
      '*** (Jen 2.13)
      Call GetDtVal("LMP", ADC("InDt"), , wDtDt, ms_CmCurCd)
      '*** (Jen 2.13)
      ADC("InLmpDt") = wDtDt
    End If
  Case Is = UCase("InLmpSal")
    If ADC.Mode = xNorm And wDsgExists = False And mb_ChgInDt = False Then
      '*** (Bef 2.13) ADC("InLmpSal") = GetLmgSal("LMP", ADC("InCmCd"), ADC("InLmpCst"), ms_CmCtg)
      '*** (Jen 2.13)
      'DispMsg "Cur:" + ms_CmCurCd, etError
      ADC("InLmpSal") = GetLmgSal("LMP", ADC("InCmCd"), ADC("InLmpCst"), ms_CmCtg, ms_CmCurCd)
      '*** (Jen 2.13)
      
      If ADC("InLmpSal") = 0 Then ADC("InLmpSal") = ADC("InLmpCst")
    End If
  ' ****** Sachin 2.12 - 02/12/05 - SJM
  Case Is = UCase("InLmsCst")
    If ADC.Mode = xNorm And mb_ChgInDt = False Then
      If mf_SlvMod = True Then
        '*** (Bef 2.13) Call GetDtVal("LMS", ADC("InDt"), wDtVal)
        '*** (Jen 2.13)
        Call GetDtVal("LMS", ADC("InDt"), wDtVal, , ms_CmCurCd)
        '*** (Jen 2.13)
        ADC("InLmsCst") = wDtVal
      Else
        ADC("InLmsCst") = 1
      End If
    End If
  Case Is = UCase("InLmsDt")
    If ADC.Mode = xNorm And mb_ChgInDt = False Then
      If mf_SlvMod = True Then
        '*** (Bef 2.13) Call GetDtVal("LMS", ADC("InDt"), , wDtDt)
        '*** (Jen 2.13)
        Call GetDtVal("LMS", ADC("InDt"), , wDtDt, ms_CmCurCd)
        '*** (Jen 2.13)
        ADC("InLmsDt") = wDtDt
      Else
        ADC("InLmsDt") = ADC("InDt")
      End If
    End If
  Case Is = UCase("InLmsSal")
    If ADC.Mode = xNorm And wDsgExists = False And mb_ChgInDt = False Then
      If mf_SlvMod = True Then
        '*** (Bef 2.13) ADC("InLmsSal") = GetLmgSal("LMS", ADC("InCmCd"), ADC("InLmsCst"), ms_CmCtg)
        '*** (Jen 2.13)
        'DispMsg "Cur1:" + ms_CmCurCd, etError
        ADC("InLmsSal") = GetLmgSal("LMS", ADC("InCmCd"), ADC("InLmsCst"), ms_CmCtg, ms_CmCurCd)
        '*** (Jen 2.13)
        If ADC("InLmsSal") = 0 Then ADC("InLmsSal") = ADC("InLmsCst")
      Else
        ADC("InLmsSal") = 1
      End If
    End If
    ' ****** Sachin 2.12 - 02/12/05 - SJM
  ' ***** Manali 3.03 - 21/06/08 - Pd Mod
  Case Is = UCase("InLmlCst")
    If ADC.Mode = xNorm And mb_ChgInDt = False Then
      If mf_PdMod = True Then
        Call GetDtVal("LML", ADC("InDt"), wDtVal, , ms_CmCurCd)
        ADC("InLmlCst") = wDtVal
      Else
        ADC("InLmlCst") = 1
      End If
    End If
  Case Is = UCase("InLmlDt")
    If ADC.Mode = xNorm And mb_ChgInDt = False Then
      If mf_PdMod = True Then
        Call GetDtVal("LML", ADC("InDt"), , wDtDt, ms_CmCurCd)
        ADC("InLmlDt") = wDtDt
      Else
        ADC("InLmlDt") = ADC("InDt")
      End If
    End If
  Case Is = UCase("InLmlSal")
    If ADC.Mode = xNorm And wDsgExists = False And mb_ChgInDt = False Then
      If mf_PdMod = True Then
        ADC("InLmlSal") = GetLmgSal("LML", ADC("InCmCd"), ADC("InLmlCst"), ms_CmCtg, ms_CmCurCd)
        If ADC("InLmlSal") = 0 Then ADC("InLmlSal") = ADC("InLmlCst")
      Else
        ADC("InLmlSal") = 1
      End If
    End If
  ' ***** Manali 3.03 - 21/06/08 - Pd Mod
  Case Is = UCase("InCnvRt")
    '*** (Bef 2.13) Dim wCnvRt As Double
    If ADC.Mode = xNorm Then
      '*** (Bef 2.13)
      'If UCase(ms_TcTyp) = UCase(ctInTcTyp) Then
      '  Call GetDtVal("USD", ADC("InDt"), wCnvRt)
      '  ADC("InCnvRt") = wCnvRt
      'ElseIf UCase(ms_TcTyp) = UCase(ctCTBTcTyp) And wDsgExists = False Then
      '  ADC("InCnvRt") = moCn.GetFldVal("Select PNum from Param where PTyp= 'CURNCY' and " + _
      '                   "PMCd= (Select CmCurCd from CustMst where CmCtg= '" + ms_CmCtg + "' and " + _
      '                   "CmCd= '" + ADC("InCmCd") + "') ")
      'End If
      '*** (Bef 2.13)
    
      '*** (Jen 2.13)
      '*** Jay 3.3(FG)
      'If UCase(ms_TcTyp) = UCase(ctInTcTyp) Then
      'gst
      If UCase(ms_TcTyp) = UCase(ctInTcTyp) Or UCase(ms_TcTyp) = UCase(ctFotTcTyp) Or UCase(ms_TcTyp) = UCase("JT") _
          Or mb_InvHdGST = True Then
      '*** Jay 3.3(FG)
      
        '*** (Bef 22/08/06)
        'ADC("InCnvRt") = CnvRt(moCn.GetFldVal("Select HBaseCurCd from Head where HCoCd= '" + ms_InCoCd  + "' and HCd= '" + ctSelfCmCd + "' "), _
        '                 ctIndCurncy, ADC("InDt"))
        '*** (Bef 22/08/06)
        ADC("InCnvRt") = CnvRt(ms_CmCurCd, ctIndCurncy, ADC("InDt"))    ''ADC("wCurCd")
      ElseIf UCase(ms_TcTyp) = UCase(ctCTBTcTyp) And wDsgExists = False Then
        ADC("InCnvRt") = CnvRt(moCn.GetFldVal("Select HBaseCurCd from Head where HCoCd= '" + ms_InCoCd + "' and HCd= '" + ctSelfCmCd + "' "), _
                         ms_CmCurCd, ADC("InDt"))
      End If
      '*** (Jen 2.13)
    End If
  
  '*** (Jen 2.13)
  Case Is = UCase("InCnvFct")
    If ADC.Mode = xNorm Then
      '*** Jay 3.3(FG)
      'If UCase(ms_TcTyp) = UCase(ctInTcTyp) And wDsgExists = False Then
      'gst
      If (UCase(ms_TcTyp) = UCase(ctInTcTyp) Or UCase(ms_TcTyp) = UCase(ctFotTcTyp) Or UCase(ms_TcTyp) = UCase("JT") _
         Or mb_InvHdGST = True) And wDsgExists = False Then
      '*** Jay 3.3(FG)
        ADC("InCnvFct") = CnvRt(moCn.GetFldVal("Select HBaseCurCd from Head where HCoCd= '" + ms_InCoCd + "' and HCd= '" + ctSelfCmCd + "' "), _
                          ms_CmCurCd, ADC("InDt"))
      
      End If
    End If
  Case Is = UCase("InRsMulDiv")
    If ADC.Mode = xNorm Then
      '*** (Bef 22/08/06)
      'ADC("InRsMulDiv") = MulDiv(moCn.GetFldVal("Select HBaseCurCd from Head where HCd= '" + ctSelfCmCd + "' and HCoCd= '" + ms_InCoCd  + "'"), _
                          ctIndCurncy)
      '*** (Bef 22/08/06)
      ADC("InRsMulDiv") = MulDiv(ms_CmCurCd, ctIndCurncy)
    End If
  Case Is = UCase("InMulDiv")
    If ADC.Mode = xNorm Then
      ADC("InMulDiv") = MulDiv(moCn.GetFldVal("Select HBaseCurCd from Head where HCd= '" + ctSelfCmCd + "' and HCoCd= '" + ms_InCoCd + "'"), _
                        ms_CmCurCd)
    End If
  '*** (Jen 2.13)
 
  Case Is = UCase("InFOBRs")
    If ADC.Mode = xNorm Then
      ' urmi 2.04-3 Round of VB and Sql is not same for .5, so get it from SQL
      'If adc("InFOB") <> 0 Then adc("InFOBRs") = Round(adc("InFOB") * adc("InCnvRt"))
      
      '*** (Bef 2.13)
      'If ADC("InFOB") <> 0 Then ADC("InFOBRs") = moCn.GetFldVal("Select Round(" + CStr(ADC("InFOB") * ADC("InCnvRt")) + " ,2)")
      '*** (Bef 2.13)
      
      '*** (Jen 2.13)
      'ADC("InFOBRs") = moCn.GetFldVal("Select convert( decimal (16, 0), " + CStr(ADC("InFOB")) + IIF(ADC("InCnvRt") > 0 And ADC("InRsMulDiv") = "/", "/ ", "* ") + CStr(ADC("InCnvRt")) + ")")
      
      '*** (Bef 22/08/06) proper for definition of CnvRt= Base to rupee
      'If ADC("InFOB") <> 0 Then _
      '   ADC("InFOBRs") = moCn.GetFldVal("Select convert( decimal (16, 0), " + CStr(ADC("InFOB")) + IIF(ADC("InCnvFct") > 0 And ADC("InMulDiv") = "M", "/ ", "* ") + CStr(ADC("InCnvFct")) + IIF(ADC("InCnvRt") > 0 And ADC("InRsMulDiv") = "M", "* ", "/ ") + CStr(ADC("InCnvRt")) + ")")
      '*** (Bef 22/08/06) proper for definition of CnvRt= Base to rupee
         
      If ADC("InFOB") <> 0 Then _
         ADC("InFOBRs") = moCn.GetFldVal("Select convert( decimal (16, 0), " + CStr(ADC("InFOB")) + IIF(ADC("InCnvRt") > 0 And ADC("InRsMulDiv") = "M", "* ", "/ ") + CStr(ADC("InCnvRt")) + ")")
      '*** (Jen 2.13)
      
    End If
' Bef Jay 3.1.1
'  '*** Jay 2.14(LclInv) ***
'  Case Is = UCase("InComm")
'    If ADC.Mode = xNorm And UCase(ms_TcTyp) = UCase(ctInTcTyp) Then
'      If ADC("InLclInvYn") = "Y" Then ADC("InComm") = 0
'    End If
'  Case Is = UCase("InFrt")
'    If ADC.Mode = xNorm And UCase(ms_TcTyp) = UCase(ctInTcTyp) Then
'      If ADC("InLclInvYn") = "Y" Then ADC("InFrt") = 0
'    End If
'  Case Is = UCase("InIns")
'    If ADC.Mode = xNorm And UCase(ms_TcTyp) = UCase(ctInTcTyp) Then
'      If ADC("InLclInvYn") = "Y" Then ADC("InIns") = 0
'    End If
'  '*** Jay 2.14(LclInv) ***
' Bef Jay 3.1.1
  '*** Jay 3.1.1 [Discount]
  ' ***** Manali 3.8.1 - Self Payable Options Added -InDeduct
  'sv.44 InTCS added
  Case Is = UCase("InComm"), UCase("InFrt"), UCase("InIns"), UCase("InDsc"), UCase("InDeduct")
    '*** Jay 3.3(FG)
    'If ADC.Mode = xNorm And UCase(ms_TcTyp) = UCase(ctInTcTyp) Then
    If ADC.Mode = xNorm And (UCase(ms_TcTyp) = UCase(ctInTcTyp) Or UCase(ms_TcTyp) = UCase(ctFotTcTyp) Or UCase(ms_TcTyp) = UCase("JT")) Then
    '*** Jay 3.3(FG)
      If ADC("InLclInvYn") = "Y" Then
        ADC(IdName) = 0
      Else
        ' ***** Manali 3.8.1 - Self Payable Option In Invoice
        If (UCase(IdName) = UCase("InDsc") And UCase(SourceName) = UCase("INHSELFDSCYN")) Or _
            (UCase(IdName) = UCase("InDeduct") And UCase(SourceName) = UCase("INHSELFDEDUCTYN")) Then
          If ADC(SourceName) = "Y" And ADC(IdName) < 0 Then ADC(IdName) = Abs(ADC(IdName))
        End If
        ' ***** Manali 3.8.1 - Self Payable Option In Invoice
      End If
    End If
  '*** Jay 3.1.1 [Discount]
  '*** Jay 3.1.1 [Discount] (Simplified)
  ' ***** Manali 3.8.1 - Self Payable Options Added -InDeduct
  'sv.44 InTCSRs added to get TCS amount in Rs
  Case Is = UCase("InCommRs"), UCase("InFrtRs"), UCase("InInsRs"), UCase("InDscRs"), UCase("InDeductRs")
    '*** Jay 3.3(FG)
    'If ADC.Mode = xNorm And UCase(ms_TcTyp) = UCase(ctInTcTyp) Then
    If ADC.Mode = xNorm And (UCase(ms_TcTyp) = UCase(ctInTcTyp) Or UCase(ms_TcTyp) = UCase(ctFotTcTyp) Or UCase(ms_TcTyp) = UCase(ctFotTcTyp)) Then
    '*** Jay 3.3(FG)
      Dim wPrntId As String '##### Setting Customer Currency Values for Respective Fields ######
      Dim wAbs As String
      wAbs = ""
      wPrntId = Mid(IdName, 1, Len(IdName) - 2)
      If ADC("InLclInvYn") = "Y" Then ADC(IdName) = 0: Exit Sub
      
      ' ***** Manali 3.8.1 - Self Payable Option In Invoice
      If (UCase(IdName) = UCase("InDscRs") And UCase(SourceName) = UCase("INHSELFDSCYN")) Or _
            (UCase(IdName) = UCase("InDeductRs") And UCase(SourceName) = UCase("INHSELFDEDUCTYN")) Then
         If ADC(SourceName) = "Y" Then wAbs = "Abs"
      End If
      If ADC(wPrntId) <> 0 And Not ((UCase(SourceName) = UCase("INHSELFDSCYN") Or UCase(SourceName) = UCase("INHSELFDEDUCTYN")) And ADC(SourceName) = "N") Then
        ADC(IdName) = moCn.GetFldVal("Select " + wAbs + "(" + CStr(ADC(wPrntId)) + IIF(ADC("InCnvRt") > 0 And ADC("InRsMulDiv") = "M", "* ", "/ ") + CStr(ADC("InCnvRt")) + ") ")
      End If
      ' ***** Manali 3.8.1 - Self Payable Option In Invoice
      ' Bef  3.8.1
      'If ADC(wPrntId) <> 0 Then ADC(IdName) = moCn.GetFldVal("Select "  + CStr(ADC(wPrntId)) + IIF(ADC("InCnvRt") > 0 And ADC("InRsMulDiv") = "M", "* ", "/ ") + CStr(ADC("InCnvRt")) + " ")
    End If
  '*** Jay 3.1.1 [Discount] (Simplified)
  
  '*** Bef 3.1.1 (*** Jay 3.1.1)
  'Case Is = UCase("InCommRs")
  '  If ADC.Mode = xNorm And UCase(ms_TcTyp) = UCase(ctInTcTyp) Then
  '    '*** (Bef 2.13) If ADC("InComm") <> 0 Then ADC("InCommRs") = ADC("InComm") * ADC("InCnvRt")
  '    '*** (Jen 2.13)
  '    'If ADC("InComm") <> 0 Then ADC("InCommRs") = moCn.GetFldVal("Select " + CStr(ADC("InComm")) + IIF(ADC("InCnvRt") > 0 And ADC("InRsMulDiv") = "/", "/ ", "* ") + CStr(ADC("InCnvRt")) + " ")
  '    '*** (Bef 22/08/06) If ADC("InComm") <> 0 Then ADC("InCommRs") = moCn.GetFldVal("Select " + CStr(ADC("InComm")) + IIF(ADC("InCnvFct") > 0 And ADC("InMulDiv") = "M", "/ ", "* ") + CStr(ADC("InCnvFct")) + IIF(ADC("InCnvRt") > 0 And ADC("InRsMulDiv") = "M", "* ", "/ ") + CStr(ADC("InCnvRt")) + " ")
  '    '*** Jay 2.14(LclInv) ***
  '    If ADC("InLclInvYn") = "Y" Then ADC("InCommRs") = 0: Exit Sub
  '    '*** Jay 2.14(LclInv) ***
  '    If ADC("InComm") <> 0 Then ADC("InCommRs") = moCn.GetFldVal("Select " + CStr(ADC("InComm")) + IIF(ADC("InCnvRt") > 0 And ADC("InRsMulDiv") = "M", "* ", "/ ") + CStr(ADC("InCnvRt")) + " ")
  '    '*** (Jen 2.13)
  '  End If
  'Case Is = UCase("InFrtRs")
  '  If ADC.Mode = xNorm And UCase(ms_TcTyp) = UCase(ctInTcTyp) Then
  '    '*** (Bef 2.13) If ADC("InFrt") <> 0 Then ADC("InFrtRs") = ADC("InFrt") * ADC("InCnvRt")
  '    '*** (Jen 2.13)
  '    'If ADC("InFrt") <> 0 Then ADC("InFrtRs") = moCn.GetFldVal("Select " + CStr(ADC("InFrt")) + IIF(ADC("InCnvRt") > 0 And ADC("InRsMulDiv") = "/", "/ ", "* ") + CStr(ADC("InCnvRt")) + " ")
  '    '*** (Bef 22/08/06) If ADC("InFrt") <> 0 Then ADC("InFrtRs") = moCn.GetFldVal("Select " + CStr(ADC("InFrt")) + IIF(ADC("InCnvFct") > 0 And ADC("InRsMulDiv") = "M", "/ ", "* ") + CStr(ADC("InCnvFct")) + IIF(ADC("InCnvRt") > 0 And ADC("InRsMulDiv") = "M", "* ", "/ ") + CStr(ADC("InCnvRt")) + " ")
  '    '*** Jay 2.14(LclInv) ***
  '    If ADC("InLclInvYn") = "Y" Then ADC("InFrtRs") = 0: Exit Sub
  '    '*** Jay 2.14(LclInv) ***
  '    If ADC("InFrt") <> 0 Then ADC("InFrtRs") = moCn.GetFldVal("Select " + CStr(ADC("InFrt")) + IIF(ADC("InCnvRt") > 0 And ADC("InRsMulDiv") = "M", "* ", "/ ") + CStr(ADC("InCnvRt")) + " ")
  '    '*** (Jen 2.13)
  '  End If
  'Case Is = UCase("InInsRs")
  '  If ADC.Mode = xNorm And UCase(ms_TcTyp) = UCase(ctInTcTyp) Then
  '    '*** (Bef 2.13) If ADC("InIns") <> 0 Then ADC("InInsRs") = ADC("InIns") * ADC("InCnvRt")
  '    '*** (Jen 2.13)
  '    'If ADC("InIns") <> 0 Then ADC("InInsRs") = moCn.GetFldVal("Select " + CStr(ADC("InIns")) + IIF(ADC("InCnvRt") > 0 And ADC("InRsMulDiv") = "/", "/ ", "* ") + CStr(ADC("InCnvRt")) + " ")
  '    '*** (Bef 22/08/06) If ADC("InIns") <> 0 Then ADC("InInsRs") = moCn.GetFldVal("Select " + CStr(ADC("InIns")) + IIF(ADC("InCnvFct") > 0 And ADC("InMulDiv") = "M", "/ ", "* ") + CStr(ADC("InCnvFct")) + IIF(ADC("InCnvRt") > 0 And ADC("InRsMulDiv") = "M", "* ", "/ ") + CStr(ADC("InCnvRt")) + " ")
  '    '*** Jay 2.14(LclInv) ***
  '    If ADC("InLclInvYn") = "Y" Then ADC("InInsRs") = 0: Exit Sub
  '    '*** Jay 2.14(LclInv) ***
  '    If ADC("InIns") <> 0 Then ADC("InInsRs") = moCn.GetFldVal("Select " + CStr(ADC("InIns")) + IIF(ADC("InCnvRt") > 0 And ADC("InRsMulDiv") = "M", "* ", "/ ") + CStr(ADC("InCnvRt")) + " ")
  '    '*** (Jen 2.13)
  '  End If
  '*** Bef 3.1.1 (*** Jay 3.1.1)
  
'??????? *** Added For 2.01
  Case Is = UCase("InTotAs")
    'Debug.Print adc.Mode
    If ADC.Mode = xNorm Then
      If (ADC("InFrt") <> 0 Or ADC("InFrtRs") <> 0) And (ADC("InIns") = 0 And ADC("InInsRs") = 0) Then
        ADC("InTotAs") = "C&F"
      ElseIf (ADC("InFrt") = 0 And ADC("InFrtRs") = 0) And (ADC("InIns") <> 0 Or ADC("InInsRs") <> 0) Then
        ADC("InTotAs") = "C&I"
      ElseIf (ADC("InFrt") <> 0 Or ADC("InFrtRs") <> 0) And (ADC("InIns") <> 0 Or ADC("InInsRs") <> 0) Then
        ADC("InTotAs") = "CIF"
      ElseIf ADC("InFrt") = 0 And ADC("InFrtRs") = 0 And ADC("InIns") = 0 And ADC("InInsRs") = 0 Then
        ADC("InTotAs") = "FOB"
      End If

      
'      If (adc("InComm") <> 0 Or adc("InCommRs") <> 0) And (adc("InFrt") <> 0 Or adc("InFrtRs") <> 0) And (adc("InIns") = 0 And adc("InInsRs") = 0) Then
'        adc("InTotAs") = "C&F"
'      ElseIf (adc("InComm") <> 0 Or adc("InCommRs") <> 0) And (adc("InFrt") = 0 And adc("InFrtRs") = 0) And (adc("InIns") <> 0 Or adc("InInsRs") <> 0) Then
'        adc("InTotAs") = "C&I"
'      ElseIf (adc("InComm") <> 0 Or adc("InCommRs") <> 0) Or (adc("InFrt") <> 0 Or adc("InFrtRs") <> 0) Or (adc("InIns") <> 0 Or adc("InInsRs") <> 0) Then
'        adc("InTotAs") = "CIF"
'      ElseIf adc("InComm") = 0 And adc("InCommRs") = 0 And adc("InFrt") = 0 And adc("InFrtRs") = 0 And adc("InIns") = 0 And adc("InInsRs") = 0 Then
'        adc("InTotAs") = "FOB"
'      End If
    End If
'*****
  Case Is = UCase("InTCS")
    If ADC.Mode = xNorm Then
        ADC("InTcs") = ADC("InTcsPer") * ADC("wRunningInTot") / 100
    End If
  Case Is = UCase("InTCSRs")
    If ADC.Mode = xNorm Then
        ADC("InTcsRs") = ADC("InTcsPer") * ADC("wRunningInTotRs") / 100
    End If
  Case Is = UCase("wRunningInTot")
    Dim wComm As Single       ' ***** Manali 3.5.0 - 11/12/08 - HideCommission Introduced
    ' ***** Manali 3.8.1 -  Self Payable Option In Invoice
    Dim wDsc As Single, wDeduct As Single
    wComm = IIF(ADC("InHHideCommYN") = "Y", 0, ADC("InComm"))
    wDsc = IIF(ADC("InHSelfDscYN") = "Y", 0, ADC("InDsc"))
    wDeduct = IIF(ADC("InHSelfDeductYN") = "Y", 0, ADC("InDeduct"))
    'sv.44 TCS amount added
    'If ADC.Mode = xNorm Then
      If ADC("InLclInvYn") = "N" Then
        If ADC("InTotAs") = "FOB" Then
          ADC("wRunningInTot") = ADC("InFOB") + wDsc + wDeduct
        ElseIf ADC("InTotAs") = "CIF" Then
          ADC("wRunningInTot") = ADC("InFOB") + wComm + ADC("InFrt") + ADC("InIns") + wDsc + wDeduct
        ElseIf ADC("InTotAs") = "C&F" Then
          ADC("wRunningInTot") = ADC("InFOB") + wComm + ADC("InFrt") + wDsc + wDeduct
        ElseIf ADC("InTotAs") = "C&I" Then
          ADC("wRunningInTot") = ADC("InFOB") + wComm + ADC("InIns") + wDsc + wDeduct
        Else
          ADC("wRunningInTot") = 0
        End If
      ElseIf ADC("InLclInvYn") = "Y" Then
        ADC("wRunningInTot") = ADC("InFOB") + ADC("InLclOctroiRs") + ADC("InLclVatRs") + ADC("InLclCstRs") + ADC("InLclOthrValRs")
      End If
      
    'End If
  Case Is = UCase("InTot")
      If ADC.Mode = xNorm Then
    ADC("InTot") = ADC("wRunningInTot") + ADC("InTCS")
    End If
' ***** Manali 3.8.1 -  Self Payable Option In Invoice
'    Bef  3.8.1
'    If ADC("InHHideCommYN") = "Y" Then
'      wComm = 0
'    Else
'      wComm = ADC("InComm")
'    End If
'    If ADC.Mode = xNorm Then
'      '*** Jay 2.14(LclInv) *** [InLclInvYn introduced]
'      If ADC("InLclInvYn") = "N" Then
'        '*** Jay 3.1.1 [Discount]
'        ' **** Manali 3.5.0 - 11/12/08 - InComm Replaced By wComm
'        If ADC("InTotAs") = "FOB" Then
'          ADC("InTot") = ADC("InFOB") + ADC("InDsc")
'        ElseIf ADC("InTotAs") = "CIF" Then
'          ADC("InTot") = ADC("InFOB") + wComm + ADC("InFrt") + ADC("InIns") + ADC("InDsc")
'          'ADC("InTot") = ADC("InFOB") + ADC("InComm") + ADC("InFrt") + ADC("InIns") + ADC("InDsc")
'        ElseIf ADC("InTotAs") = "C&F" Then
'          ADC("InTot") = ADC("InFOB") + wComm + ADC("InFrt") + ADC("InDsc")
'          'ADC("InTot") = ADC("InFOB") + ADC("InComm") + ADC("InFrt") + ADC("InDsc")
'        ElseIf ADC("InTotAs") = "C&I" Then
'          ADC("InTot") = ADC("InFOB") + wComm + ADC("InIns") + ADC("InDsc")
'          'ADC("InTot") = ADC("InFOB") + ADC("InComm") + ADC("InIns") + ADC("InDsc")
'        '*** Jay 3.1.1 [Discount]
'        Else
'          ADC("InTot") = 0
'        End If
'      ElseIf ADC("InLclInvYn") = "Y" Then
'        ADC("InTot") = ADC("InFOB") + ADC("InLclOctroiRs") + ADC("InLclVatRs") + ADC("InLclCstRs") + ADC("InLclOthrValRs")
'      End If
'      '*** Jay 2.14(LclInv) ***
'    End If

'  Case Is = UCase("InTotRs")
'    If adc.Mode = xNorm Then
'      adc("InTotRs") = adc("InTot") * adc("InCnvRt")
'    End If

  Case Is = UCase("wRunningInTotRs")
    Dim wCommRs As Double
    ' ***** Manali 3.8.1 -  Self Payable Option In Invoice
    Dim wDscRs As Double, wDeductRs As Double
    wCommRs = IIF(ADC("InHHideCommYN") = "Y", 0, ADC("InCommRs"))
    wDscRs = IIF(ADC("InHSelfDscYN") = "Y", 0, ADC("InDscRs"))
    wDeductRs = IIF(ADC("InHSelfDeductYN") = "Y", 0, ADC("InDeductRs"))
    'sv.44 TCS amount in Rs added
    If (UCase(ms_TcTyp) = UCase(ctInTcTyp) Or UCase(ms_TcTyp) = UCase(ctFotTcTyp) Or UCase(ms_TcTyp) = UCase("JT")) Then
      If ADC("InLclInvYn") = "N" Then
        If ADC("InTotAs") = "FOB" Then
          ADC("wRunningInTotRs") = ADC("InFOBRs") + wDscRs + wDeductRs
        ElseIf ADC("InTotAs") = "CIF" Then
          ADC("wRunningInTotRs") = ADC("InFOBRs") + wCommRs + ADC("InFrtRs") + ADC("InInsRs") + wDscRs + wDeductRs
        ElseIf ADC("InTotAs") = "C&F" Then
          ADC("wRunningInTotRs") = ADC("InFOBRs") + wCommRs + ADC("InFrtRs") + wDscRs + wDeductRs
        ElseIf ADC("InTotAs") = "C&I" Then
          ADC("wRunningInTotRs") = ADC("InFOBRs") + wCommRs + ADC("InInsRs") + wDscRs + wDeductRs
        Else
          ADC("wRunningInTotRs") = 0
        End If
      ElseIf ADC("InLclInvYn") = "Y" Then
        If ADC("InTot") <> 0 Then _
         ADC("wRunningInTotRs") = moCn.GetFldVal("Select convert( decimal (16, 0), " + CStr(ADC("InTot")) + IIF(ADC("InCnvRt") > 0 And ADC("InRsMulDiv") = "M", "* ", "/ ") + CStr(ADC("InCnvRt")) + ")")
      End If
    End If
  Case Is = UCase("InTotRs")
      If ADC.Mode = xNorm Then
    ADC("InTotRs") = ADC("wRunningInTotRs") + ADC("InTcsRs")
    End If
    ' Bef 3.8.1
''    If ADC("InHHideCommYN") = "Y" Then
''      wCommRs = 0
''    Else
''      wCommRs = ADC("InCommRs")
''    End If
''    '*** Jay 3.3(FG)
''    'If ADC.Mode = xNorm And UCase(ms_TcTyp) = UCase(ctInTcTyp) Then
''    If ADC.Mode = xNorm And (UCase(ms_TcTyp) = UCase(ctInTcTyp) Or UCase(ms_TcTyp) = UCase(ctFotTcTyp)) Then
''    '*** Jay 3.3(FG)
''      If ADC("InLclInvYn") = "N" Then
''        '*** Jay 3.1.1 [Discount]
''        ' **** Manali 3.5.0 - 11/12/08 - InCommRs Replaced By wCommRs
''        If ADC("InTotAs") = "FOB" Then
''          ADC("InTotRs") = ADC("InFOBRs") + ADC("InDscRs")
''        ElseIf ADC("InTotAs") = "CIF" Then
''          ADC("InTotRs") = ADC("InFOBRs") + wCommRs + ADC("InFrtRs") + ADC("InInsRs") + ADC("InDscRs")
''          'adc("InTotRs") = adc("InFOBRs") + adc("InCommRs") + adc("InFrtRs") + adc("InInsRs") + adc("InDscRs")
''        ElseIf ADC("InTotAs") = "C&F" Then
''          ADC("InTotRs") = ADC("InFOBRs") + wCommRs + ADC("InFrtRs") + ADC("InDscRs")
''          'adc("InTotRs") = adc("InFOBRs") + adc("InCommRs") + adc("InFrtRs") + adc("InDscRs")
''        ElseIf ADC("InTotAs") = "C&I" Then
''          ADC("InTotRs") = ADC("InFOBRs") + wCommRs + ADC("InInsRs") + ADC("InDscRs")
''          'adc("InTotRs") = adc("InFOBRs") + adc("InCommRs") + adc("InInsRs") + adc("InDscRs")
''        '*** Jay 3.1.1 [Discount]
''        Else
''          ADC("InTotRs") = 0
''        End If
''      '*** Jay 2.14(LclInv) ***
''      ElseIf ADC("InLclInvYn") = "Y" Then
''        If ADC("InTot") <> 0 Then _
''         ADC("InTotRs") = moCn.GetFldVal("Select convert( decimal (16, 0), " + CStr(ADC("InTot")) + IIF(ADC("InCnvRt") > 0 And ADC("InRsMulDiv") = "M", "* ", "/ ") + CStr(ADC("InCnvRt")) + ")")
''      End If
''      '*** Jay 2.14(LclInv) ***
''    End If

'  Case Is = UCase("InCmAdd1")
'    If ADC.Mode = xNorm Then
'      ADC("InCmAdd1") = moCn.GetFldVal("Select CmAdd1 From CustMst where CmCtg= 'C' and CmCd= '" + ADC("InCmCd") + "'")
'    End If
'  Case Is = UCase("InCmAdd2")
'    If ADC.Mode = xNorm Then
'      ADC("InCmAdd2") = moCn.GetFldVal("Select CmAdd2 From CustMst where CmCtg= 'C' and CmCd= '" + ADC("InCmCd") + "'")
'    End If
'  Case Is = UCase("InCmAdd3")
'    If ADC.Mode = xNorm Then
'      ADC("InCmAdd3") = moCn.GetFldVal("Select CmAdd3 From CustMst where CmCtg= 'C' and CmCd= '" + ADC("InCmCd") + "'")
'    End If
  Case Is = UCase("InCmTel")
    If ADC.Mode = xNorm Then
      If ADC("InhInvTyp") <> "DR" And ADC("InhInvTyp") <> "BS" Then
        ADC("InCmTel") = moCn.GetFldVal("Select CmTel From CustMst where CmCtg= '" + ms_CmCtg + "' and CmCd= '" + ADC("InCmCd") + "'")
      End If
    End If
  Case Is = UCase("InCmFax")
    If ADC.Mode = xNorm Then
      If ADC("InhInvTyp") <> "DR" And ADC("InhInvTyp") <> "BS" Then
        ADC("InCmFax") = moCn.GetFldVal("Select CmFax From CustMst where CmCtg= '" + ms_CmCtg + "' and CmCd= '" + ADC("InCmCd") + "'")
      End If
    End If
  Case Is = UCase("InCmEMail")
    If ADC.Mode = xNorm Then
     If ADC("InhInvTyp") <> "DR" And ADC("InhInvTyp") <> "BS" Then
      ADC("InCmEMail") = moCn.GetFldVal("Select CmEMail From CustMst where CmCtg= '" + ms_CmCtg + "' and CmCd= '" + ADC("InCmCd") + "'")
     End If
    End If
  Case Is = UCase("InCmCpName")
    If ADC.Mode = xNorm Then
        If ADC("InhInvTyp") <> "DR" And ADC("InhInvTyp") <> "BS" Then
          ADC("InCmCpName") = moCn.GetFldVal("Select CmCpName From CustMst where CmCtg= '" + ms_CmCtg + "' and CmCd= '" + ADC("InCmCd") + "'")
        End If
    End If
    
  'Sachin 4.1.0.0
  Case Is = UCase("InCmPortOfDisch")
    If ADC.Mode = xNorm Then
        If SourceName = UCase("InCmCd") Then
            ADC("InCmPortOfDisch") = moCn.GetFldVal("Select CmPortOfDisch From CustMst where CmCtg= '" + ms_CmCtg + "' and CmCd= '" + ADC("InCmCd") + "'")
        ElseIf SourceName = UCase("InCmPod") Then
            ADC("InCmPortOfDisch") = moCn.GetFldVal("Select PDesc From Param where PTyp= 'PORTCD' and PMCd= '" + ADC("InCmPod") + "' And PSCd=''")
        End If
    End If
    
  Case Is = UCase("InCmFinDest")
    If ADC.Mode = xNorm Then
        If SourceName = UCase("InCmCd") Then
            ADC("InCmFinDest") = moCn.GetFldVal("Select CmFinDest From CustMst where CmCtg= '" + ms_CmCtg + "' and CmCd= '" + ADC("InCmCd") + "' ")
        ElseIf SourceName = UCase("InCmFinDstCd") Then
            ADC("InCmFinDest") = moCn.GetFldVal("Select PDesc From Param where PTyp= 'CONTRYCD' and PMCd= '" + ADC("InCmFinDstCd") + "' And PSCd=''")
        End If
    End If
    
  Case Is = UCase("InCmPod")
    If ADC.Mode = xNorm Then
       ADC("InCmPod") = moCn.GetFldVal("Select CmPOd From CustMst where CmCtg= '" + ms_CmCtg + "' and CmCd= '" + ADC("InCmCd") + "'")
    End If
    
  Case Is = UCase("InCmFinDstCd")
    If ADC.Mode = xNorm Then
      ADC("InCmFinDstCd") = moCn.GetFldVal("Select CmFinDstCd From CustMst where CmCtg= '" + ms_CmCtg + "' and CmCd= '" + ADC("InCmCd") + "'")
    End If
  'Sachin 4.1.0.0
    
'  Case Is = UCase("InCmBuyer")
'    If ADC.Mode = xNorm Then
'      ADC("InCmBuyer") = moCn.GetFldVal("Select CmBuyer From CustMst where CmCtg= 'C' and CmCd= '" + ADC("InCmCd") + "'")
'    End If
  Case Is = UCase("InCmPayTerms")
    If ADC.Mode = xNorm Then
      ADC("InCmPayTerms") = moCn.GetFldVal("Select CmPayTerms From CustMst where CmCtg= '" + ms_CmCtg + "' and CmCd= '" + ADC("InCmCd") + "'")
    End If
  '**************Geeta*******Emr207*********
  Case Is = UCase("InHVessel")
    If ADC.Mode = xNorm Then
      ADC("InHVessel") = moCn.GetFldVal("Select CmVessel From CustMst where CmCtg= '" + ms_CmCtg + "' and CmCd= '" + ADC("InCmCd") + "'")
    End If
  '**************Geeta*******Emr207*********
  Case Is = UCase("InCmInsBy")
    If ADC.Mode = xNorm Then
      ADC("InCmInsBy") = moCn.GetFldVal("Select CmHInsBy From CustMst where CmCtg= '" + ms_CmCtg + "' and CmCd= '" + ADC("InCmCd") + "'")
    End If
  '*** Jay 3.1.1 [CmOurBkCd]
  Case Is = UCase("wBkCd")
    If ADC.Mode = xNorm And ADC("InCmCd") <> "" Then
      ADC("wBkCd") = moCn.GetFldVal("Select CmOurBkCd From CustMst where CmCtg='" + ms_CmCtg + "' and CmCd= '" + ADC("InCmCd") + "'")
    End If
  '*** Jay 3.1.1 [CmOurBkCd]
  Case Is = UCase("InHBkName")
    If ADC.Mode = xNorm And ADC("wBkCd") <> "" Then
      ADC("InHBkName") = moCn.GetFldVal("Select BkName From BnkAgMst where BkCtg='B' and BkCd= '" + ADC("wBkCd") + "'")
    End If
  Case Is = UCase("InHBkAdd1")
    If ADC.Mode = xNorm And ADC("wBkCd") <> "" Then
      ADC("InHBkAdd1") = moCn.GetFldVal("Select BkAdd1 From BnkAgMst where BkCtg='B' and BkCd= '" + ADC("wBkCd") + "'")
    End If
  Case Is = UCase("InHBkAdd2")
    If ADC.Mode = xNorm And ADC("wBkCd") <> "" Then
      ADC("InHBkAdd2") = moCn.GetFldVal("Select BkAdd2 From BnkAgMst where BkCtg='B' and BkCd= '" + ADC("wBkCd") + "'")
    End If
  Case Is = UCase("InHBkAdd3")
    If ADC.Mode = xNorm And ADC("wBkCd") <> "" Then
      ADC("InHBkAdd3") = moCn.GetFldVal("Select BkAdd3 From BnkAgMst where BkCtg='B' and BkCd= '" + ADC("wBkCd") + "'")
    End If
  Case Is = UCase("InHBnkInstr")
    If ADC.Mode = xNorm And ADC("wBkCd") <> "" Then
      ADC("InHBnkInstr") = moCn.GetFldVal("Select BkInstr From BnkAgMst where BkCtg='B' and  BkCd= '" + ADC("wBkCd") + "'")
    End If
   Case Is = UCase("InHBkAc")
    If ADC.Mode = xNorm And ADC("wBkCd") <> "" Then
      ADC("InHBkAc") = moCn.GetFldVal("Select BkAc From BnkAgMst where BkCtg='B' and BkCd= '" + ADC("wBkCd") + "'")
    End If
  Case Is = UCase("InCmCd")
    If ADC.Mode = xNorm Then
        If ADC("InCmName") <> "" Then
            wCmCd = moCn.GetFldVal("Select CmCd from CustMst where CmCtg= '" + ms_CmCtg + "' and " + _
                                 "CmName= '" + ADC("InCmName") + "'")
            If wCmCd <> ADC("InCmCd") Then ADC("InCmCd") = wCmCd
        End If
    End If
  Case Is = UCase("InPrdDesc")
    If ADC.Mode = xNorm Then
      If UCase(ADC("wInvPrdDesc")) <> UCase(OldVal) Then _
      ADC("InPrdDesc") = moCn.GetFldVal("Select PDesc225 from Param where PTyp= 'INVDESC' and PMCd = '" + ADC("wInvPrdDesc") + "'")
    End If
    
 ' ****** Manali 3.03 - 16/08/08 - Agent Code, Name and Address ******************************
  Case Is = UCase("wAgCd")
    If ADC.Mode = xNorm And ADC("InCmCd") <> "" Then
        ADC("wAgCd") = moCn.GetFldVal("Select CmAgCd From CustMst where CmCtg='" + ms_CmCtg + "' and CmCd= '" + ADC("InCmCd") + "'")
    End If

  Case Is = UCase("InCmAgName")
    If ADC.Mode = xNorm And ADC("wAgCd") <> "" Then
        ADC("InCmAgName") = moCn.GetFldVal("Select BkName From BnkAgMst where BkCtg='A' and BkCd= '" + ADC("wAgCd") + "'")
    End If
  Case Is = UCase("InCmAgAdd1")
    If ADC.Mode = xNorm And ADC("wAgCd") <> "" Then
        ADC("InCmAgAdd1") = moCn.GetFldVal("Select BkAdd1 From BnkAgMst where BkCtg='A' and BkCd= '" + ADC("wAgCd") + "'")
    End If
  Case Is = UCase("InCmAgAdd2")
    If ADC.Mode = xNorm And ADC("wAgCd") <> "" Then
        ADC("InCmAgAdd2") = moCn.GetFldVal("Select BkAdd2 From BnkAgMst where BkCtg='A' and BkCd= '" + ADC("wAgCd") + "'")
    End If
  Case Is = UCase("InCmAgAdd3")
    If ADC.Mode = xNorm And ADC("wAgCd") <> "" Then
        ADC("InCmAgAdd3") = moCn.GetFldVal("Select BkAdd3 From BnkAgMst where BkCtg='A' and BkCd= '" + ADC("wAgCd") + "'")
    End If
 ' ****** Manali 3.03 - 16/08/08 - Agent Code, Name and Address ******************************
 
 
'************************************** urmi 4Jul03  ***************************************
  Case Is = UCase("InFrRmDc")
    If ADC.Mode = xNorm Then
      If ADC("wInFrRmDc") = "I" Then
        ADC("InFrRmDc") = "C"
      ElseIf ADC("wInFrRmDc") = "R" Then
        ADC("InFrRmDc") = "D"
      Else
        ADC("InFrRmDc") = ""
      End If
    End If
  Case Is = UCase("InToRmDc")
    If ADC.Mode = xNorm Then
      If ADC("wInFrRmDc") = "I" Then
         ADC("InToRmDc") = "D"
      ElseIf ADC("wInFrRmDc") = "R" Then
         ADC("InToRmDc") = "C"
      Else
         ADC("InToRmDc") = ""
      End If
    End If
'*******************************************************************************************
  'Manoj
  Case Is = UCase("InHGSPOrgCrit")
    If ADC.Mode = xNorm Then
      ADC("InHGSPOrgCrit") = GetOriginCriteria
    End If
  'Manoj

  '*** (Jen 2.13)
  Case Is = UCase("wCurCd")
    If ADC.Mode = xNorm Or ADC.Mode = xInit Then
      ADC("wCurCd") = moCn.GetFldVal("Select CmCurCd from CustMst where CmCtg= '" + ms_CmCtg + "' and CmCd= '" + ADC("InCmCd") + "'")
    End If
  '*** (Jen 2.13)
  '*** Jay 2.14(LclInv) ***
  Case Is = UCase("InLclInvYn")
    If ADC.Mode = xNorm And (UCase(ms_TcTyp) = UCase(ctInTcTyp) Or UCase(ms_TcTyp) = UCase(ctFotTcTyp) Or UCase(ms_TcTyp) = UCase("JT") Or mb_InvHdGST) Then
      ADC("InLclInvYn") = IIF(moCn.RecSeek("Select 'x' from Head Join CustMst On CmCtg= '" + ADC("InCmCtg") + "' and " + _
                          "CmCd = '" + ADC("InCmCd") + "' where HCocd= '" + ms_InCoCd + "' and " + _
                          "HLclInvYN= 'Y' and CmCurCd= 'RS' "), _
                          "Y", "N")
    End If
  Case Is = UCase("InCmOctroiPer")
    If ADC.Mode = xNorm And ADC("InLclInvYn") <> "" And ADC("InCmCd") <> "" Then
      ADC("InCmOctroiPer") = IIF(ADC("InLclInvYn") = "N", 0, moCn.GetFldVal("Select CmOctroiPer from CustMst " + _
                                        "where CmCtg= '" + ms_CmCtg + "' and CmCd= '" + ADC("InCmCd") + "' and CmOctroiYn = 'Y'"))
    End If
  Case Is = UCase("InHVatPer")
    If ADC.Mode = xNorm And ADC("InLclInvYn") <> "" And ADC("InCmCd") <> "" Then
      ADC("InHVatPer") = IIF(ADC("InLclInvYn") = "N", 0, _
                           moCn.GetFldVal("Select IsNull ((Select HVatPer from Head Join CustMst On HCocd= CmBillCo " + _
                           "where CmCtg= '" + ms_CmCtg + "' and CmCd= '" + ADC("InCmCd") + "' and CmVatYn = 'Y'),0)"))
    End If
  Case Is = UCase("InHCstPer")
    If ADC.Mode = xNorm And ADC("InLclInvYn") <> "" And ADC("InCmCd") <> "" Then
      ADC("InHCstPer") = IIF(ADC("InLclInvYn") = "N", 0, _
                           moCn.GetFldVal("Select IsNull ((Select HCstPer from Head Join CustMst On HCocd= CmBillCo " + _
                           "where CmCtg= '" + ms_CmCtg + "' and CmCd= '" + ADC("InCmCd") + "' and CmCstYn = 'Y'),0)"))
    End If
  Case Is = UCase("InLclOthrValRs")
    '*** Jay 3.3(FG)
    'If ADC.Mode = xNorm And UCase(ms_TcTyp) = UCase(ctInTcTyp) Then
    If ADC.Mode = xNorm And (UCase(ms_TcTyp) = UCase(ctInTcTyp) Or UCase(ms_TcTyp) = UCase(ctFotTcTyp) Or UCase(ms_TcTyp) = UCase("JT")) Then
    '*** Jay 3.3(FG)
      If ADC("InLclInvYn") = "N" Then ADC("InLclOthrValRs") = 0
    End If
  Case Is = UCase("InLclOctroiRs")
    '*** Jay 3.3(FG)
    'If ADC.Mode = xNorm And UCase(ms_TcTyp) = UCase(ctInTcTyp) Then
    If ADC.Mode = xNorm And (UCase(ms_TcTyp) = UCase(ctInTcTyp) Or UCase(ms_TcTyp) = UCase(ctFotTcTyp) Or UCase(ms_TcTyp) = UCase("JT")) Then
    '*** Jay 3.3(FG)
      ADC("InLclOctroiRs") = Round(ADC("InFOB") * ADC("InCmOctroiPer") / 100, 2)
    End If
  Case Is = UCase("InLclVatRs")
    '*** Jay 3.3(FG)
    'If ADC.Mode = xNorm And UCase(ms_TcTyp) = UCase(ctInTcTyp) Then
    If ADC.Mode = xNorm And (UCase(ms_TcTyp) = UCase(ctInTcTyp) Or UCase(ms_TcTyp) = UCase(ctFotTcTyp) Or UCase(ms_TcTyp) = UCase("JT")) Then
    '*** Jay 3.3(FG)
      ADC("InLclVatRs") = Round((ADC("InFOB") + ADC("InLclOctroiRs") + ADC("InLclOthrValRs")) * ADC("InHVatPer") / 100, 2)
    End If
  Case Is = UCase("InLclCstRs")
    '*** Jay 3.3(FG)
    'If ADC.Mode = xNorm And UCase(ms_TcTyp) = UCase(ctInTcTyp) Then
    If ADC.Mode = xNorm And (UCase(ms_TcTyp) = UCase(ctInTcTyp) Or UCase(ms_TcTyp) = UCase(ctFotTcTyp) Or UCase(ms_TcTyp) = UCase("JT")) Then
    '*** Jay 3.3(FG)
      ADC("InLclCstRs") = Round((ADC("InFOB") + ADC("InLclOctroiRs") + ADC("InLclOthrValRs")) * ADC("InHCstPer") / 100, 2)
    End If
  Case Is = UCase("INLCLOTHRVALDESC")
    '*** Jay 3.3(FG)
    'If ADC.Mode = xNorm And UCase(ms_TcTyp) = UCase(ctInTcTyp) And ADC("InLclInvYn") = "N" Then
    If ADC.Mode = xNorm And (UCase(ms_TcTyp) = UCase(ctInTcTyp) Or UCase(ms_TcTyp) = UCase(ctFotTcTyp) Or UCase(ms_TcTyp) = UCase("JT")) And ADC("InLclInvYn") = "N" Then
    '*** Jay 3.3(FG)
      ADC("InLclOthrValDesc") = ""
    End If
  '*** Jay 2.14(LclInv) ***
  
  '3.11.0   '4.1.0.0
  Case Is = UCase("INTOTASTXT")
    If ADC.Mode = xNorm Then ADC("InTotAsTxt") = ADC("InTotAs")
  
  Case Is = UCase("InTCSPer")
    Dim wInTcsPer As Double
    wInTcsPer = IIF(ADC("InLclInvYn") = "Y", 0, _
                           moCn.GetFldVal("Select IsNull ((Select HTCSPer from Head Join CustMst On  HCocd= '" + ADC("InCoCd") + "' and HCd='" + ctSelfCmCd + "'" + _
                           " and CmCtg= '" + ms_CmCtg + "' and CmCd= '" + ADC("InCmCd") + "' and CmTCSYN = 'Y'),0)"))
    ADC("InTcsPer") = wInTcsPer
  '578-20 updating consignee name and address from customer store atble
  Case Is = UCase("InHConsignee")
    If ADC.Mode = xNorm And ADC("InCstCd") <> "" Then
        ADC("InHConsignee") = moCn.GetFldVal("Select CstName From CustStore where CstCmCtg='" + ADC("InCmCtg") + "' and CstCmCd= '" + ADC("InCmCd") + "' and CstCd= '" + ADC("InCstCd") + "'")
    End If
  Case Is = UCase("InConAdd1")
    If ADC.Mode = xNorm And ADC("InCstCd") <> "" Then
        ADC("InConAdd1") = moCn.GetFldVal("Select CstAdd1 From CustStore where CstCmCtg='" + ADC("InCmCtg") + "' and CstCmCd= '" + ADC("InCmCd") + "' and CstCd= '" + ADC("InCstCd") + "'")
    End If
  Case Is = UCase("InConAdd2")
    If ADC.Mode = xNorm And ADC("InCstCd") <> "" Then
        ADC("InConAdd2") = moCn.GetFldVal("Select CstAdd2 From CustStore where CstCmCtg='" + ADC("InCmCtg") + "' and CstCmCd= '" + ADC("InCmCd") + "' and CstCd= '" + ADC("InCstCd") + "'")
    End If
  Case Is = UCase("InConAdd3")
    If ADC.Mode = xNorm And ADC("InCstCd") <> "" Then
        ADC("InConAdd3") = moCn.GetFldVal("Select CstAdd3 From CustStore where CstCmCtg='" + ADC("InCmCtg") + "' and CstCmCd= '" + ADC("InCmCd") + "' and CstCd= '" + ADC("InCstCd") + "'")
    End If

 End Select
End Sub
Private Sub ATXT_FldChange(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String, ByVal pv_NewValue As Variant)
  '*** InTc         PMCd('TC', PValue= ms_TcTyp) and the type of entry should be 'IN'
  '*** InYy         PMCd('YY')
  '*** InChr        PSCd('CHR', adc(InTc)) or ''
  '*** InDt         Not Blank
  '*** InCmCd       CmCd from CustMst(CmCtg= 'C')
  '*** InCls        PMCd('YN')
  '*** InTotAs      PMCd('INVTOTAS')
  'stellar.1- calling a function to check Quotes exists in given value
  Dim wChkQtsErrMsg As String
  If InStr(1, "|" + ms_ChkQtsIdName + "|", "|" + IdName + "|") > 0 Then
    wChkQtsErrMsg = ChkQuotes(Me, IdName, pv_NewValue)
    If wChkQtsErrMsg <> "" Then Cancel = True: ErrMsg = wChkQtsErrMsg: Exit Sub
  End If
  
  Select Case UCase(Trim(IdName))
  'indigo.160- validating company code
  Case Is = UCase("InCoCd")
    Cancel = Not moCn.RecSeek("Select 'x' From Head Where HCd='" + ctSelfCmCd + "' And HCoCd='" + pv_NewValue + "'")
    If Cancel = True Then ErrMsg = "Invalid Company Code": Exit Sub
    If Not ADC.Connection.RecSeek("Select * from UsrAccess where UaMnuCd='" + ADC.MenuCd + "' And " + _
                        " UaUsrCd ='" + gs_UsrCd + "' And UaMnuForm = 'MDI'" + _
                      " And ',' + UaValidCoCd + ',' Like Case When (UaValidCoCd = '') Then ',' + UaValidCoCd + ',' Else '%," & UCase$(pv_NewValue) & ",%' End") Then
    
        Cancel = True
        ErrMsg = "User does not have permissions to access this Company Code": Exit Sub
    End If
    ms_InCoCd = pv_NewValue
    Me.Caption = GetFrmCaption(ms_InCoCd, Me.Caption)
  
  Case Is = UCase("InTc")
    Cancel = Not moCn.RecSeek("select PMCd from Param where PTyp= 'TC' and " + _
                         "PMCd = '" + pv_NewValue + "' and " + _
                         "PValue= '" + ms_TcTyp + "' ")
    If Cancel = True Then ErrMsg = "Invalid Invoice Voucher Tc": Exit Sub
  Case Is = UCase("InYy")
    Cancel = Not moCn.RecSeek("select PMCd from Param where PTyp= 'YY' and " + _
                         "PMCd = '" + pv_NewValue + "'")
    If Cancel = True Then ErrMsg = "Invalid Invoice Voucher Year": Exit Sub
  Case Is = UCase("InChr")
    Cancel = pv_NewValue <> "" And (Not moCn.RecSeek("select vPSCd from vParam where " + _
                         "vPCoCd= '" + ms_InCoCd + "' and vPTyp= 'CHR' and " + _
                         "vPMCd = '" + ADC("InTc") + "' and vPSCd='" + pv_NewValue + "'"))
    If Cancel = True Then ErrMsg = "Invalid Invoice Voucher Character": Exit Sub
' **** Zubin 211 **** '
    If Cancel = True Then ErrMsg = "Invalid Voucher Character": Exit Sub
    Dim ws_UsrArr() As String, wi_Index As Integer
    ws_UsrArr = Split(moCn.GetFldVal("Select vPValue4 from vParam where vPCoCd= '" + ms_InCoCd + "' and " + _
                                        "vPTyp='CHR' and vPMCd='" + ADC("InTc") + "' and vPSCd= '" + pv_NewValue + "'"), ",")
    If IsArray(ws_UsrArr) Then
      For wi_Index = 0 To UBound(ws_UsrArr)
        Cancel = IIF(UCase(gs_UsrCd) = UCase(ws_UsrArr(wi_Index)), False, True)
        If Cancel = False Then: Exit For
      Next wi_Index
    End If
    If Cancel = True Then ErrMsg = "Access To This Invoice Voucher Character Is Denied To User " + gs_UsrCd: Exit Sub
' **** Zubin 211 **** '
  
  Case Is = UCase("InDt")
    If pv_NewValue = "" Then Cancel = True: ErrMsg = "Invoice Date Cannot Be Blank": Exit Sub
    If pv_NewValue > CDate(ms_FutureDt) Then Cancel = True: ErrMsg = "Invoice Date Cannot Be Greater Than Future Date": Exit Sub
    If pv_NewValue < CDate(ms_LockDt) Then Cancel = True: ErrMsg = "Invoice Date Cannot Be Less Than Lock Date": Exit Sub
    'CC.999.13 check for purge date added
    If pv_NewValue <= CDate(ms_PurgeDt) Then Cancel = True: ErrMsg = "Invoice Date Cannot Be Less Than Or Equal To Purge Date": Exit Sub
    
    If moCn.GetFldVal("Select hHolChkYN From Head where HCd='ZSELF' and HCoCd='" + ADC("InCoCd") + "'") = "Y" Then
    'indigo.160- company code changed from gs_CoCd to ms_InCoCd
      If IsHOLYN(ADC("InCoCd"), pv_NewValue) Then Cancel = True: ErrMsg = "Voucher Cannot be Added on Holidays": Exit Sub
    End If
    ADC("InPrnDt") = pv_NewValue
  Case Is = UCase("InPrnDt")
    If pv_NewValue = "" Then Cancel = True: ErrMsg = "Invoice Print Date Cannot Be Blank": Exit Sub
  Case Is = UCase("InCmCd")
    Cancel = Not moCn.RecSeek("select CmCd from CustMst where CmCtg= '" + ms_CmCtg + "' and " + _
                         " CmCd= '" + pv_NewValue + "'")
    If Cancel = True Then ErrMsg = "Invalid Customer Code": Exit Sub
    
    '*** (Jen 2.13)
    ms_CmCurCd = moCn.GetFldVal("Select CmCurCd from CustMst where CmCtg= '" + ms_CmCtg + "' and CmCd= '" + pv_NewValue + "'")
    '*** (Jen 2.13)
    '*** Jay 2.14(LclInv) ***
    ' ***** Manali 3.10.0 - InCmCd passed
    'GST PG.11 PER Rajeshbhai removing this IF CoDets will always come from Company Code record
'    If ms_CmCurCd = ctIndCurncy And moCn.RecSeek("Select 'x' From Head where " + _
 '          "HCoCd= '" + ms_InCoCd  + "' and HLclInvYn= 'Y'") Then
  '    Call BillCoDets(moCn.GetFldVal("Select (Case CmBillCo when '' then '" + ms_InCoCd  + "' else CmBillCo end ) " + _
   '                                         "from CustMst where CmCtg= 'C' and CmCd = '" + pv_NewValue + "'"), pv_NewValue)
   ' Else
    Call BillCoDets(ms_InCoCd, pv_NewValue)
    'End If
    '*** Jay 2.14(LclInv) ***
  Case Is = UCase("InCls")
    Cancel = Not moCn.RecSeek("select PMCd from Param where PTyp= 'YN' and " + _
                         "PMCd = '" + pv_NewValue + "'")
    If Cancel = True Then ErrMsg = "Invoice Closed Yes/No? (Enter Y/N)": Exit Sub
    '********Geeta***********Emr206****02/04/04
    If pv_NewValue = "Y" Then
      txtPwd.Enabled = True: ms_PwdFlag = False                     '*** Jay 2.14 *** (SEO Change)
      ADC.AllowDelete = False
    Else
      txtPwd.text = "": txtPwd.Enabled = False: ms_PwdFlag = True   '*** Jay 2.14 *** (SEO Change)
      ADC.AllowDelete = True
    End If
    '*************

  Case Is = UCase("wInvPrdDesc")
    Cancel = pv_NewValue <> "" And Not moCn.RecSeek("select PMCd from Param where PTyp= 'INVDESC' and " + _
                         "PMCd = '" + pv_NewValue + "'")
    If Cancel = True Then ErrMsg = "Invalid Product Description": Exit Sub
      
  '***** Shilpa ******
  Case Is = UCase("InhInvTyp")
    If pv_NewValue = "" Then Cancel = True: ErrMsg = "Invoice Type Cannot Be Blank": Exit Sub
    Cancel = Not moCn.RecSeek("select PMCd from Param where PTyp= 'INVTYP' and " + _
                         "PMCd = '" + pv_NewValue + "'")
    If Cancel = True Then ErrMsg = "Invalid Invoice Type": Exit Sub
  '***** Shilpa ******
  Case Is = UCase("InTotAs")
    Cancel = Not moCn.RecSeek("select PMCd from Param where PTyp= 'INVTOTAS' and " + _
                         "PMCd = '" + pv_NewValue + "'")
    If Cancel = True Then ErrMsg = "Invalid 'Total Value As' Option": Exit Sub
  Case Is = UCase("wBkCd")
    Cancel = pv_NewValue <> "" And (Not moCn.RecSeek(" Select BkCd From BnkAgMst where BkCtg='B' and BkCd= '" + pv_NewValue + "'"))
    If Cancel = True Then ErrMsg = "Invalid Bank Code": Exit Sub
  Case Is = UCase("InHRndOffFOB")
    Cancel = pv_NewValue <> "" And (Not moCn.RecSeek("select PMCd from Param where PTyp= 'YN' and " + _
                         "PMCd = '" + pv_NewValue + "'"))
    If Cancel = True Then ErrMsg = "Invalid FOB Round Off Option": Exit Sub
  Case Is = UCase("InCmName")
    Cancel = pv_NewValue <> "" And (Not moCn.RecSeek("select CmName from CustMst where CmCtg= '" + ms_CmCtg + "' and CmName = '" + pv_NewValue + "'"))
    If Cancel = True Then ErrMsg = "Invalid Customer Name": Exit Sub
  
  ' ****** Manali 3.03 - 16/08/08 - Agent Code, Name and Address ******************************
  Case Is = UCase("wAgCd")
    Cancel = pv_NewValue <> "" And (Not moCn.RecSeek(" Select BkCd From BnkAgMst where BkCtg='A' and BkCd= '" + pv_NewValue + "'"))
    If Cancel = True Then ErrMsg = "Invalid Agent Code": Exit Sub
  ' ****** Manali 3.03 - 16/08/08 - Agent Code, Name and Address ******************************
  
  '********************** urmi 4Jul03 ****************************************************
  Case Is = UCase("wInFrRmDc")
    Cancel = pv_NewValue <> "" And (Not moCn.RecSeek("select PmCd from Param where PTyp= 'IR' and PMCd = '" + pv_NewValue + "'"))
    If Cancel = True Then ErrMsg = "Invalid Issue/Receipt Option": Exit Sub
    
    '*** Exceptional check where previous value is checked
    '****** Sachin 3.02 28-12-07 - Current Partition
    '*** Jay 3.3(FG)
    Cancel = ADC("wInFrRmDc") <> "" And wDsgExists
    'Cancel = ADC("wInFrRmDc") <> "" And moCn.RecSeek("Select 'x' from InvDsg where " + _
             "IdCoCd= '" + ms_InCoCd  + "' and IdTc = '" + ADC("InTc") + "' and " + _
             "IdYy = '" + ADC("InYy") + "' and IdChr = '" + ADC("InChr") + "' and IdNo = " + _
             CStr(ADC("InNo")) + " and IdPrtKey=(Case When '" + gs_Partition + "'='" + ctCurrPrtn + "' Then '" + ctCurrPrtn + "' Else IdPrtKey End) ")
    '*** Jay 3.3(FG)
    If Cancel = True Then ErrMsg = "Cannot Change From Rm Loc I/R When InvDsg Records exists": Exit Sub
    '*** Exceptional check where previous value is checked
  Case Is = UCase("InFrRmDc")
    Cancel = pv_NewValue <> "" And (Not moCn.RecSeek("select PmCd from Param where PTyp= 'DC' and PMCd = '" + pv_NewValue + "'"))
    If Cancel = True Then ErrMsg = "Invalid From Debit/Credit Option": Exit Sub
  Case Is = UCase("InToRmDc")
    Cancel = pv_NewValue <> "" And (Not moCn.RecSeek("select PmCd from Param where PTyp= 'DC' and PMCd = '" + pv_NewValue + "'"))
    If Cancel = True Then ErrMsg = "Invalid To Debit/Credit Option": Exit Sub
  Case Is = UCase("InFrRmLoc")
  'gst
    Dim wFrLocTyp As String
    wFrLocTyp = " and LocTyp IN (" + IIF(ms_TcTyp = "IN", "'F','XP','XS','XD','XM')", "'S','MG','XT','XP','XS','XD','Y','XM','XC')")
    
'    If ms_TcTyp = "PR" Or ms_TcTyp = "CNV" Or ms_TcTyp = "MV" Or ms_TcTyp = "DT" Or ms_TcTyp = "BV" Or ms_TcTyp = "BS" Or ms_TcTyp = "BM" Or ms_TcTyp = "MLT" Or ms_TcTyp = "REJ" Or ms_TcTyp = "FB" Or ms_TcTyp = "FWB" Then
 '     wFrLocTyp = " and LocTyp ='XC'"
 '   Else
  '    wFrLocTyp = " and LocTyp IN (" + IIF(ms_TcTyp = "JT", "'S','MG','XT'", "'F'") + " , 'XP','XS','XD','XM')"
'    End If
    Cancel = pv_NewValue <> "" And (Not moCn.RecSeek("Select LocCd from Loc where " + _
                              "LocCoCd= '" + ms_InCoCd + "' and " + _
                              "LocCd = '" + pv_NewValue + "' " + wFrLocTyp))
    If Cancel = True Then ErrMsg = "Invalid Raw Material From Location": Exit Sub
    '*** Jay 3.3(FG) (Bef)
    Cancel = wDsgExists
    'Cancel = moCn.RecSeek("Select 'x' from InvDsg where IdPrtKey='" + ADC("InPrtKey") + "' And IdCoCd= '" + ms_InCoCd  + "' " + _
             "and IdTc = '" + ADC("InTc") + "' and IdYy = '" + ADC("InYy") + "' " + _
             "and IdChr = '" + ADC("InChr") + "' and IdNo = " + _
            CStr(ADC("InNo")))
    '*** Jay 3.3(FG) (Bef)
    If Cancel = True Then ErrMsg = "Cannot Change From RM Loc When InvDsg Records exists": Exit Sub
  Case Is = UCase("InToRmLoc")
  'gst
    Dim wToLocTyp As String
'    If ms_TcTyp = "PR" Or ms_TcTyp = "CNV" Or ms_TcTyp = "MV" Or ms_TcTyp = "DT" Or ms_TcTyp = "BV" Or ms_TcTyp = "BS" Or ms_TcTyp = "BM" Or ms_TcTyp = "MLT" Or ms_TcTyp = "REJ" Or ms_TcTyp = "FB" Or ms_TcTyp = "FWB" Then
 '     wToLocTyp = " and LocTyp ='XC'"
 '   Else
  '    wToLocTyp = " and LocTyp IN (" + IIF(ms_TcTyp = "JT", "'S','MG','XT','XD','XMLT'", "'F'") + ", 'XS','XD','XM')"
  '  End If
    wToLocTyp = " and LocTyp IN (" + IIF(ms_TcTyp = "IN", "'F','XP','XS','XD','XM')", "'S','Y','MG','XT','XD','XMLT','XP','XS','XD','XM','XC')")
  
    Cancel = pv_NewValue <> "" And (Not moCn.RecSeek("Select LocCd from Loc where " + _
                              "LocCoCd= '" + ms_InCoCd + "' and " + _
                              "LocCd = '" + pv_NewValue + "' " + wToLocTyp))
    If Cancel = True Then ErrMsg = "Invalid Raw Material To Location": Exit Sub
    '*** Jay 3.3(FG)
    Cancel = wDsgExists
    'Cancel = moCn.RecSeek("Select 'x' from InvDsg where IdPrtKey='" + ADC("InPrtKey") + "' And IdCoCd= '" + ms_InCoCd  + "' " + _
             "and IdTc = '" + ADC("InTc") + "' and IdYy = '" + ADC("InYy") + "' " + _
             "and IdChr = '" + ADC("InChr") + "' and IdNo = " + _
             CStr(ADC("InNo")))
    '*** Jay 3.3(FG)
    If Cancel = True Then ErrMsg = "Cannot Change To RM Loc When InvDsg Records exists": Exit Sub
  '********************** urmi 4Jul03 ****************************************************
  '***** Shilpa ******
  Case Is = UCase("InHValAddMtlColYN")
    Cancel = (Not moCn.RecSeek("select PMCd from Param where PTyp= 'YN' and " + _
                         "PMCd = '" + pv_NewValue + "'"))
    ErrMsg = "Show Value Addition On Metal Column In Value Addition Report? Enter Yes Or No (Y/N)?": Exit Sub
  Case Is = UCase("InHValBOEDetYN")
    Cancel = (Not moCn.RecSeek("select PMCd from Param where PTyp= 'YN' and " + _
                         "PMCd = '" + pv_NewValue + "'"))
        ErrMsg = "Show BOE Details In Value Addition Annexur Report? Enter Yes Or No (Y/N)?": Exit Sub
    '***** Shilpa ******
    '*********Geeta*********Emr207
  Case Is = UCase("InFOBPer")
    Cancel = IIF(pv_NewValue > 99.99 Or pv_NewValue < -99.99, True, False)
    ErrMsg = "FOB Percentage Cannot be greater than 99.99 and Less than -99.99"
   '*****Geeta***Emr208***
   ' **** Manali 3.03 - 21/06/08 - Pd Mod
  Case Is = UCase("InHGldFiness"), UCase("InHPlFiness"), UCase("InHSlFiness"), UCase("InHPdFiness")
        Cancel = IIF(pv_NewValue <= 0, True, False)
    ErrMsg = "Gold/Platinum/Silver/Palladium Fineness Should be >0": Exit Sub
 
  '*** (Jen 2.13)
  Case Is = UCase("InShowGldLs")
     Cancel = (Not moCn.RecSeek("select PMCd from Param where PTyp= 'YN' and " + _
                          "PMCd = '" + pv_NewValue + "'"))
     If Cancel = True Then ErrMsg = "Show Gold Loss % Field In Invoice Reports? Enter Yes Or No (Y/N)?": Exit Sub
  '*** (Jen 2.13)
  
  '*** Jay 2.14(LclInv) ***
  Case Is = UCase("InLclInvYn")
    Cancel = (Not moCn.RecSeek("select PMCd from Param where PTyp= 'YN' and " + _
                       "PMCd = '" + pv_NewValue + "'"))
    If Cancel = True Then ErrMsg = "Invalid Local Invoice (Y/N)?": Exit Sub
    If ADC("InFOB") <> 0 Then Cancel = True: ErrMsg = "Can Not Change Local Invoice YN as Invoice Designs Exists": Exit Sub
    If pv_NewValue = "Y" Then
      If moCn.RecSeek("Select 'x' From Head where " + _
           "HCoCd= '" + ms_InCoCd + "' and HLclInvYn = 'Y'") Then
        Call BillCoDets(moCn.GetFldVal("Select (Case CmBillCo when '' then '" + ms_InCoCd + "' else CmBillCo end ) " + _
                                            "from CustMst where CmCtg= 'C' and CmCd = '" + ADC("InCmCd") + "'"))
''        Call BillCoDets(moCn.GetFldVal("Select (Case CmBillCo when '' then '" + ms_InCoCd  + "' else CmBillCo end ) " + _
''                                            "from CustMst where CmCtg= 'C' and CmCd = '" + pv_NewValue + "'"))
      Else
        Cancel = True: ErrMsg = "Local Invoice Module Is Not Available"
      End If
    ElseIf pv_NewValue = "N" Then
      Call BillCoDets(ms_InCoCd)
    End If
  '*** Jay 2.14(LclInv) ***
  
  '*** (Jen 3.01.0_1)
  Case Is = UCase("InDsgOnHld")
    Cancel = (Not moCn.RecSeek("select PMCd from Param where PTyp= 'YN' and " + _
                               "PMCd= '" + pv_NewValue + "' and PSCd= '' "))
    ErrMsg = "Allow More Designs To Be Copied In The Invoice? Enter Yes Or No (Y/N)?": Exit Sub
  Case Is = UCase("InInvGrsWtEqActual")
    Cancel = Not moCn.RecSeek("select 'x' from Param where PTyp= 'YN' " + _
             "and PMCd = '" + pv_NewValue + "' and PSCd= ''")
    ErrMsg = "Make Invoice Gross Weight Equal To Actual During Copy, Enter Yes Or No (Y/N)?": Exit Sub
  Case Is = UCase("InRemakingCtgs")
    Cancel = pv_NewValue <> "" And (Not moCn.RecSeek("select PMCd from Param where PTyp= 'REMAKE' and " + _
                                    "PMCd= '" + pv_NewValue + "' and PSCd= '' "))
    If Cancel = True Then ErrMsg = "Invalid Remake Option": Exit Sub
    
    '****** Sachin 3.02 28-12-07 - Current Partition
    '*** Jay 3.3(FG)
    Cancel = wDsgExists
    'Cancel = moCn.RecSeek("Select 'x' from InvDsg where IdCoCd= '" + ms_InCoCd  + "' " + _
             "and IdTc = '" + ADC("InTc") + "' and IdYy = '" + ADC("InYy") + "' " + _
             "and IdChr = '" + ADC("InChr") + "' and IdNo = " + _
            CStr(ADC("InNo")) + " and IdPrtKey=(Case When '" + gs_Partition + "'='" + ctCurrPrtn + "' Then '" + ctCurrPrtn + "' Else IdPrtKey End)")
    '*** Jay 3.3(FG)
    If Cancel = True Then ErrMsg = "Cannot Change Remake Option When InvDsg Records exists": Exit Sub
  '*** (Jen 3.01.0_1)
  
  '****** Sachin 3.3.0 (23-06-08) ******
  '****** [42. Make Invoice By Bag] ******
  Case Is = UCase("InByBagYN")
    Cancel = (Not moCn.RecSeek("select PMCd from Param where PTyp= 'YN' and " + _
                         "PMCd = '" + pv_NewValue + "'"))
    ErrMsg = "Make Invoice By Bag? Enter Yes Or No (Y/N)?": Exit Sub
  '****** [42. Make Invoice By Bag] ******
  
  '****** [14. Invoice Rate With Respect To Stock YN] ******
  Case Is = UCase("InIRtWrtStkYN")
    Cancel = (Not moCn.RecSeek("select PMCd from Param where PTyp= 'YN' and " + _
                         "PMCd = '" + pv_NewValue + "'"))
    ErrMsg = "Invoice Rate od Diamond and Color Stones With Respect to Stock Rt? Enter Yes Or No (Y/N)?": Exit Sub
    '****** [14. Invoice Rate With Respect To Stock YN] ******
  '****** Sachin 3.3.0 (23-06-08) ******
  
  ' ***** Manali 3.6.0 - 27/10/09 - Invoicing with Markup on Cost of Accessories
  Case Is = UCase("InXIRtWrtStkYN")
    Cancel = (Not moCn.RecSeek("select PMCd from Param where PTyp= 'YN' and " + _
                         "PMCd = '" + pv_NewValue + "'"))
    ErrMsg = "Invoice Rate of Accessories With Respect to Stock Rt? Enter Yes Or No (Y/N)?": Exit Sub
  ' ***** Manali 3.6.0 - 27/10/09 - Invoicing with Markup on Cost of Accessories
  
  ' ***** Sachin 3.11.0 - Invoicing with Markup on Metals
  Case Is = UCase("InMetIRtWrtStkYN")
    Cancel = (Not moCn.RecSeek("select PMCd from Param where PTyp= 'YN' and " + _
                         "PMCd = '" + pv_NewValue + "'"))
    ErrMsg = "Invoice Rate of Metals With Respect to Stock Rt? Enter Yes Or No (Y/N)?": Exit Sub
    
  Case Is = UCase("InRmRtfrRtChrt")
    Cancel = (Not moCn.RecSeek("select PMCd from Param where PTyp= 'YN' and " + _
                         "PMCd = '" + pv_NewValue + "'"))
    ErrMsg = "Invoice RM Rate from Rate Chart? Enter Yes Or No (Y/N)?": Exit Sub

  Case Is = UCase("InLabRtfrRtChrt")
    Cancel = (Not moCn.RecSeek("select PMCd from Param where PTyp= 'YN' and " + _
                         "PMCd = '" + pv_NewValue + "'"))
    ErrMsg = "Invoice Labor Rate from Rate Chart? Enter Yes Or No (Y/N)?": Exit Sub
  
  '4.1.2
  Case Is = UCase("InRtFrInvCust")
    Cancel = (Not moCn.RecSeek("select PMCd from Param where PTyp= 'YN' and " + _
                         "PMCd = '" + pv_NewValue + "'"))
    ErrMsg = "Invoice Rate from Invoice Customer? Enter Yes Or No (Y/N)?": Exit Sub
  
  '3.11.0
  Case Is = UCase("InOrdMulbyYN")
    Cancel = (Not moCn.RecSeek("select PMCd from Param where PTyp= 'YN' and " + _
                         "PMCd = '" + pv_NewValue + "'"))
    ErrMsg = "Apply Order Mulby factor? Enter Yes Or No (Y/N)?": Exit Sub
  ' ***** Sachin 3.11.0 - Invoicing with Markup on Metals
    
  ' **** Manali 3.5.0 - 28/11/08 - Hide Commission
  Case Is = UCase("InHHideCommYN")
    Cancel = (Not moCn.RecSeek("select PMCd from Param where PTyp= 'YN' and " + _
                         "PMCd = '" + pv_NewValue + "'"))
    ErrMsg = "Commssion Self Payable? Enter Yes Or No (Y/N)?": Exit Sub
    'ErrMsg = "Hide Commssion In Invoice Printing? Enter Yes Or No (Y/N)?": Exit Sub
  ' **** Manali 3.5.0 - 28/11/08 - Hide Commission
  
  ' ***** Manali 3.8.1 - Self Payable Options In Invoice added
  Case Is = UCase("InHSelfDscYN")
    Cancel = Not moCn.RecSeek("select PMCd from Param where " + _
                                      "PTyp= 'YN' and PMCd = '" + pv_NewValue + "'")
    ErrMsg = "Discount Self Payable? Enter Yes Or No (Y/N)": Exit Sub
  
  Case Is = UCase("InHSelfDeductYN")
    Cancel = Not moCn.RecSeek("select PMCd from Param where " + _
                                      "PTyp= 'YN' and PMCd = '" + pv_NewValue + "'")
    ErrMsg = "Deduction Self Payable? Enter Yes Or No (Y/N)": Exit Sub
  
  
  ' **** Manali 3.6.0 - 30/10/09 - For DTA Use Notional Rate in Custom Matching
  Case Is = UCase("InNotionRtYn")
    Cancel = (Not moCn.RecSeek("select PMCd from Param where PTyp= 'YN' and " + _
                         "PMCd = '" + pv_NewValue + "'"))
    ErrMsg = "Use Notional Rate in Custom Matching ? Enter Yes Or No (Y/N)?": Exit Sub
  ' **** Manali 3.6.0 - 30/10/09 - For DTA Use Notional Rate in Custom Matching
  
  ' ***** Manali 3.8.1 - Self Payable Option In Invoice
  Case Is = UCase("InDsc"), UCase("InDscRs")
    Cancel = pv_NewValue < 0 And ADC("InHSelfDscYn") = "Y"
    ErrMsg = "Discount Cannot Be Entered In Negative When Self Payable Is Set To Yes For Discount": Exit Sub
    
  Case Is = UCase("InDeduct"), UCase("InDeductRs")
    Cancel = pv_NewValue < 0 And ADC("InHSelfDeductYn") = "Y"
    ErrMsg = "Deduction Cannot Be Entered In Negative When Self Payable Is Set To Yes For Deduction": Exit Sub
    
  '3.11.0
  Case Is = UCase("InTotAsTxt")
    If Trim(pv_NewValue) = "" Then
        Cancel = True
        ErrMsg = "Invoice Total As Text Cannot be Blank. Please enter some value."
    End If
    
  'Sachin 4.1.0.0
  Case Is = UCase("InCmPod")
    Cancel = (Not moCn.RecSeek("Select PMCd from Param where PTyp= 'PORTCD' And " + _
                         " PMCd = '" + pv_NewValue + "'"))
    ErrMsg = "Invalid Port of Discharge code": Exit Sub
    
  Case Is = UCase("InCmFinDstCd")
    Cancel = (Not moCn.RecSeek("Select PMCd from Param where PTyp= 'CONTRYCD' And " + _
                         " PMCd = '" + pv_NewValue + "'"))
    ErrMsg = "Invalid Final Destination code": Exit Sub
  'Sachin 4.1.0.0
  'IG 91- Checking Valid Paytermcode, Blank allowed
  Case Is = UCase("wPayTermsCd")
      Cancel = pv_NewValue <> "" And Not moCn.RecSeek("select PMCd from Param where PTyp= 'PYMTRM' and PMCd = '" + pv_NewValue + "'")
      ErrMsg = "Invalid Payment Term Code"
      Exit Sub
  'pg.11-checking valid GstCode
  Case Is = UCase("InGstCd")
    Cancel = (Not moCn.RecSeek("Select PMCd from Param where PTyp= 'GSTCD' And " + _
                         " PMCd = '" + pv_NewValue + "'"))
                        
    If Cancel = True Then ErrMsg = "Invalid Gst code": Exit Sub
  '578-20 validating customer store code
  Case Is = UCase("InCstCd")
    Cancel = pv_NewValue <> "" And (Not moCn.RecSeek("Select CstCd from CustStore " + _
                         "where CstCmCtg= '" + ADC("InCmCtg") + "' and CstCmCd= '" + ADC("InCmCd") + "' " + _
                         "and CstCd= '" + pv_NewValue + "'"))
    If Cancel = True Then ErrMsg = "Invalid Customer Store code": Exit Sub
      
  End Select
  
End Sub

Private Sub CmdCDGo_Click()
    '4.3.0.0
    If ADC("wInDt") <> "01/01/80" And ADC("wInDt") <> ADC("InDt") Then
        mb_ChgInDt = True
        ADC("InDt") = ADC("wInDt")
        ADC("InPrnDt") = ADC("wInDt")
        mb_ChgInDt = False
    End If
    Call DispFra(ChgDt)
End Sub

'**********Bhavna added addl inv amts
Private Sub CmdAddAmt_Click()
    Call DispFra(AddAmt)
End Sub
'**********Bhavna added addl inv amts
Private Sub RecalcAmountFields()
    Call atxt_ReCalculate(0, "INAMT7", 0, 0, "", "")
    Call atxt_ReCalculate(0, "INAMT12", 0, 0, "", "")
End Sub
Private Sub CmdChgDt_Click()
    '4.3.0.0
    If ADC("InCls") = "N" Then
        Call DispFra(ChgDt)
    Else
        DispMsg "Cannot Change Date as Invoice is Closed", etError
    End If
End Sub

Private Sub DispFra(ByVal pv_TxnRmFra As en_TxnRmFra)
  '*** This is the common procedure which is used to toggle between hiding and displaying
      'the frames (i.e. if the frame is visible then make it invisible and viseversa)
      'associated with a command button

  '4.3.0.0
  Select Case pv_TxnRmFra
  Case Is = ChgDt
    If FraCD.Visible = True Then
      FraCD.Visible = False
      FraCD.Enabled = False
      Call EnaDisaCmds(False, CmdChgDt, ChgDt)
      CmdChgDt.SetFocus
    Else
      FraCD.Visible = True
      FraCD.Enabled = True
      CmdChgDt.SetFocus
      FraCD.ZOrder
      ADC("wInDt") = ADC("InDt")
      Call EnaDisaCmds(True, CmdChgDt)
    End If
    '**********Bhavna added add inv amt frame
    Case Is = AddAmt
    If FraAddAmt.Visible = True Then
      Dim wTotPay As Double
      wTotPay = ADC("InPayAmt1") + ADC("InPayAmt2") + ADC("InPayAmt3")
      If wTotPay <> 0 And wTotPay <> ADC("InAmt12") Then
        DispMsg "Total Payment Amount <> Total Invoice Amount", etError: Exit Sub
      End If
      FraAddAmt.Visible = False
      FraAddAmt.Enabled = False
      Call EnaDisaCmds(False, CmdAddAmt, AddAmt)
      CmdAddAmt.SetFocus
    Else
      FraAddAmt.Visible = True
      FraAddAmt.Enabled = True
      ADC("InGstCd").SetFocus
      FraAddAmt.ZOrder
      Call EnaDisaCmds(True, CmdAddAmt)
    End If
    '**********Bhavna added add inv amt frame
  End Select
End Sub

Private Sub EnaDisaCmds(ByVal pv_ShowFra As Boolean, Optional ByVal pv_CmdBut As MWCTL_BTN1, Optional ByVal pv_DispFra As en_DispFra)
  '*** This procedure is used to give the toggle option (enable and disable) to
      'the command buttons
  'Added in 4.3.0.0
  If pv_ShowFra = True Then
    Call StoreState
  End If
  If Not pv_CmdBut Is Nothing Then
    If pv_ShowFra = True Then
      If pv_CmdBut.Pic = 0 Then
        pv_CmdBut.BackColor = vbWhite
      End If
      pv_CmdBut.CausesValidation = True
    Else
      If pv_CmdBut.Pic = 0 Then
        pv_CmdBut.BackColor = &H8000000F
      End If
      '*** restoring the 'Causes Validation' state of the current pv_cmdbutton
      If Not IsMissing(pv_DispFra) Then
        pv_CmdBut.CausesValidation = mArr_Cv(pv_DispFra)
      End If
      ' Zubin 212
    End If
  End If
  FraNKeyAll.Enabled = Not pv_ShowFra
  ADC.AllowSave = Not pv_ShowFra
  ADC.AllowDelete = Not pv_ShowFra
  ADC.AllowCancel = Not pv_ShowFra
  
  If pv_ShowFra = True Then
    '** All command buttons are disabled
    CmdChgDt.Enabled = Not pv_ShowFra
    CmdAddAmt.Enabled = Not pv_ShowFra        '***********Bhavna
  Else
    '*** Restoring Ena/Disa state of all Command Buttons
    CmdChgDt.Enabled = mArr_Ed(ChgDt)
    CmdAddAmt.Enabled = mArr_Ed(AddAmt)       '**********Bhavna
  End If
  If Not pv_CmdBut Is Nothing Then pv_CmdBut.Enabled = True
End Sub
Private Sub HideAllFras()
  '*** This procedure hides all the frames other than FraNKeyAll
  'Added in 4.3.0.0
  FraNKeyAll.Visible = True
  FraCD.Visible = False
  FraAddAmt.Visible = False
End Sub
Private Sub StoreState()
  '*** This procedure is used to store the CausesValidation State and the Enabled/Disabled
      'states of all the Command Buttons
  '*** This is used for restoring the original states in the future
  'Added in 4.3.0.0
  mArr_Cv(ChgDt) = CmdChgDt.CausesValidation
  mArr_Cv(AddAmt) = CmdAddAmt.CausesValidation      '**********Bhavna
  mArr_Ed(ChgDt) = CmdChgDt.Enabled
  mArr_Ed(AddAmt) = CmdAddAmt.Enabled                 '**********Bhavna

End Sub

' ###################################  Manoj #### Ver: 2.0.6  ###################################
' Date:   20/01/2004
Private Sub TxtPwd_Validate(Cancel As Boolean)
'*** Jay 2.14 *** (SEO Change)
  If ADC.AddFndMode = xFndMode And Trim(txtPwd.text) <> "" Then
    If Trim(txtPwd.text) = moCn.GetFldVal("select IsNull ((select PValue from Param Join UsrAccess on UaUsrCd = PMCd " + _
                   "where PTyp= 'USR' and PMCd= '" + gs_UsrCd + "' and UaMnuCd = '" + ADC.MenuCd + "' and " + _
                   "UaSeoYn= 'Y' and PValue= '" + Trim(txtPwd.text) + "'), " + _
                   "(Select PValue1 from Param where Ptyp= 'SEOPWD' and PmCd= 'SEOPWD' )) ") Then
      If UCase$(ADC("InCls")) = UCase$("N") Then ms_PwdFlag = True Else ms_PwdFlag = False
      Exit Sub
    Else
       ms_PwdFlag = False: Cancel = True: DispMsg "Invalid Edit Password / User Does Not Have SEO Rights ", etError
    End If
'    If Trim(TxtPwd.text) = moCn.GetFldVal("Select PValue1 from Param where " + _
'                           "Ptyp = 'SEOPWD' and PmCd = 'SEOPWD'") Then
'      Exit Sub
'    Else
'      Cancel = True
'      DispMsg "Invalid Edit Password", etError
'    End If
  Else
    If moCn.GetFldVal("Select PValue1 from Param where Ptyp = 'SEOPWD' and PmCd = 'SEOPWD'") = "" Then _
       ms_PwdFlag = True Else ms_PwdFlag = False
  End If
'*** Jay 2.14 *** (SEO Change)
End Sub
' ###########################################  Manoj  ###########################################
Private Sub ADC_NonKeyValid(pr_Cancel As Boolean, pr_ErrMsg As String)
  
  '*** reset the first tab as the current tab of the tab page 'TabInvDet'
  '*** Jay 2.14(LclInv) ***
  If ADC.Mode = xNorm And ADC("InHVatPer") <> 0 And ADC("InHCstPer") <> 0 Then _
    pr_Cancel = True: pr_ErrMsg = "VAT and CST, Both Percentage Can Not Be Aplicable": Exit Sub
  
  If ADC.Mode = xNorm And ADC("InLclVatRs") <> 0 And ADC("InLclCstRs") <> 0 Then _
    pr_Cancel = True: pr_ErrMsg = "VAT and CST, Both Values Can Not Be Aplicable": Exit Sub
 
  '*** Jay 2.14(LclInv) ***
  TabInvDet.Tab = 0
   InvAdditionalAmtsReCalc (True)
  'stellar.1- calling a function to check Quotes exists in the list of text boxes
  Dim wChkQtsErrMsg As String
  wChkQtsErrMsg = ChkQuotes(Me, ms_ChkQtsIdName, "")
  If wChkQtsErrMsg <> "" Then pr_Cancel = True: pr_ErrMsg = wChkQtsErrMsg
End Sub
Private Sub ADC_SetRecSource()
  '*** Set the RecSource for ADC (Table InvHd)
  'sv.9 getting purge date from vparam for emr & jemr
  If ms_TcTyp = "JT" Then
    ms_PurgeDt = GetPurgeDt("ZZZ", "JT")
  Else
    ms_PurgeDt = GetPurgeDt(ms_InCoCd, "")
  End If

  '****** Sachin 3.02 28-12-07 - Current Partition
  ADC.RecSource = " Select * from InvHd " + _
                  " Where InCoCd= '" + ms_InCoCd + "'" + _
                  " And InTc= '" + ADC("InTc") + "'" + _
                  " And InYy= '" + ADC("InYy") + "'" + _
                  " And InChr= '" + ADC("InChr") + "'" + _
                  " And InNo= " + CStr(ADC("InNo")) + " And InPrtKey=(Case When '" + gs_Partition + "'='" + ctCurrPrtn + "' Then '" + ctCurrPrtn + "' Else InPrtKey End) " + _
                  IIF(gs_Partition = ctCurrPrtn, " And Indt>'" + ms_PurgeDt + "' ", "")
End Sub

Private Sub TabInvDet_Click(PreviousTab As Integer)
  '*** Whenever a tab is selected enable the frame on the clicked tab and disable the previous tab frame
  FraTabInvDet(TabInvDet.Tab).Enabled = True
  FraTabInvDet(PreviousTab).Enabled = False
  If TabInvDet.Tab = 2 Then
    If ADC("InLclInvYn") = "N" Then
      FraCal1.Visible = True
      FraCal2.Visible = False
    ElseIf ADC("InLclInvYn") = "Y" Then
      FraCal1.Visible = False
      FraCal2.Visible = True
    End If
  End If
End Sub
''***** Shilpa ******
'Private Sub FillShpDet(mInvTyp As String)
'    Dim wRsCustMst As MwfLib.MDORowSet, wRsBnkAgMst As MwfLib.MDORowSet, wRsBnkMst As MwfLib.MDORowSet
'
'    Set wRsCustMst = moCn.OpenRes("Select * From CustMst where CmCtg = 'C' and CmCd = '" + ADC("InCmCd") + "'")
'    If wRsCustMst.EOF Or wRsCustMst.BOF Then Exit Sub
'    ADC("InCmName") = ""
'    ADC("InCmAdd1") = ""
'    ADC("InCmAdd2") = ""
'    ADC("InCmAdd3") = ""
'    ADC("InCmTel") = ""
'    ADC("InCmFax") = ""
'    ADC("InCmEmail") = ""
'    ADC("InCmCpName") = ""
'    ADC("InCmBuyer") = ""
'    ADC("InBuyAdd1") = ""
'    ADC("InBuyAdd2") = ""
'    ADC("InBuyAdd3") = ""
'    ADC("InHConsignee") = ""
'    ADC("InConAdd1") = ""
'    ADC("InConAdd2") = ""
'    ADC("InConAdd3") = ""
'
'    Select Case ADC("InhInvTyp")
'    Case Is = "DR"
'        ADC("InCmName") = ""
'        ADC("InCmAdd1") = ""
'        ADC("InCmAdd2") = ""
'        ADC("InCmAdd3") = ""
'        ADC("InCmBuyer") = ""
'        ADC("InCmTel") = ""
'        ADC("InCmFax") = ""
'        ADC("InCmEmail") = ""
'        ADC("InCmCpName") = ""
'        ADC("InBuyAdd1") = ""
'        ADC("InBuyAdd2") = ""
'        ADC("InBuyAdd3") = ""
'        ADC("InHConsignee") = wRsCustMst!CmName
'        ADC("InConAdd1") = wRsCustMst!CmAdd1
'        ADC("InConAdd2") = wRsCustMst!CmAdd2
'        ADC("InConAdd3") = wRsCustMst!CmAdd3
'    Case Is = "BS"
'        Set wRsBnkMst = moCn.OpenRes("Select * From BnkAgMst where BkCtg = 'B' and BkCd = '" + wRsCustMst!CmBkCd + "'")
'        ADC("InCmName") = ""
'        ADC("InCmAdd1") = ""
'        ADC("InCmAdd2") = ""
'        ADC("InCmAdd3") = ""
'        ADC("InCmTel") = ""
'        ADC("InCmFax") = ""
'        ADC("InCmEmail") = ""
'        ADC("InCmCpName") = ""
'        ADC("InCmBuyer") = wRsCustMst!CmName
'        ADC("InBuyAdd1") = wRsCustMst!CmAdd1
'        ADC("InBuyAdd2") = wRsCustMst!CmAdd2
'        ADC("InBuyAdd3") = wRsCustMst!CmAdd3
'        If Not wRsBnkMst.EOF Or wRsBnkMst.BOF Then
'            ADC("InHConsignee") = wRsBnkMst!BkName
'            ADC("InConAdd1") = wRsBnkMst!BkAdd1
'            ADC("InConAdd2") = wRsBnkMst!BkAdd2
'            ADC("InConAdd3") = wRsBnkMst!BkAdd3
'        End If
'    Case Is = "DA"
'        Set wRsBnkMst = moCn.OpenRes("Select * From BnkAgMst where BkCtg = 'B' and BkCd = '" + wRsCustMst!CmBkCd + "'")
'        Set wRsBnkAgMst = moCn.OpenRes("Select * From BnkAgMst where BkCtg = 'A' and BkCd = '" + wRsCustMst!CmAgCd + "'")
'            ADC("InCmName") = wRsCustMst!CmName
'            ADC("InCmAdd1") = wRsCustMst!CmAdd1
'            ADC("InCmAdd2") = wRsCustMst!CmAdd2
'            ADC("InCmAdd3") = wRsCustMst!CmAdd3
'        If Not wRsBnkAgMst.EOF Or wRsBnkAgMst.BOF Then
'            ADC("InCmBuyer") = wRsBnkAgMst!BkName
'            ADC("InBuyAdd1") = wRsBnkAgMst!BkAdd1
'            ADC("InBuyAdd2") = wRsBnkAgMst!BkAdd2
'            ADC("InBuyAdd3") = wRsBnkAgMst!BkAdd3
'        End If
'            ADC("InHConsignee") = ""
'            ADC("InConAdd1") = ""
'            ADC("InConAdd2") = ""
'            ADC("InConAdd3") = ""
'    Case Is = "BA"
'        Set wRsBnkMst = moCn.OpenRes("Select * From BnkAgMst where BkCtg = 'B' and BkCd = '" + wRsCustMst!CmBkCd + "'")
'        Set wRsBnkAgMst = moCn.OpenRes("Select * From BnkAgMst where BkCtg = 'A' and BkCd = '" + wRsCustMst!CmAgCd + "'")
'            ADC("InCmName") = wRsCustMst!CmName
'            ADC("InCmAdd1") = wRsCustMst!CmAdd1
'            ADC("InCmAdd2") = wRsCustMst!CmAdd2
'            ADC("InCmAdd3") = wRsCustMst!CmAdd3
'        If Not wRsBnkAgMst.EOF Or wRsBnkAgMst.BOF Then
'            ADC("InCmBuyer") = wRsBnkAgMst!BkName
'            ADC("InBuyAdd1") = wRsBnkAgMst!BkAdd1
'            ADC("InBuyAdd2") = wRsBnkAgMst!BkAdd2
'            ADC("InBuyAdd3") = wRsBnkAgMst!BkAdd3
'        End If
'        If Not wRsBnkMst.EOF Or wRsBnkMst.BOF Then
'            ADC("InHConsignee") = wRsBnkMst!BkName
'            ADC("InConAdd1") = wRsBnkMst!BkAdd1
'            ADC("InConAdd2") = wRsBnkMst!BkAdd2
'            ADC("InConAdd3") = wRsBnkMst!BkAdd3
'        End If
'End Select
'Set wRsCustMst = Nothing
'Set wRsBnkMst = Nothing
'Set wRsBnkAgMst = Nothing
'End Sub
''***** Shilpa Emr 2.06 ******

'***** Shilpa ******
Private Sub FillShpDet(mInvTyp As String)
    Dim wRsCustMst As MwfLib.MDORowSet, wRsBnkAgMst As MwfLib.MDORowSet, wRsBnkMst As MwfLib.MDORowSet
    
    Set wRsCustMst = moCn.OpenRes("Select * From CustMst where CmCtg = '" + ms_CmCtg + "' and CmCd = '" + ADC("InCmCd") + "'")
    If wRsCustMst.EOF Or wRsCustMst.BOF Then Exit Sub
    ADC("InCmName") = wRsCustMst!CmName
    ADC("InCmAdd1") = wRsCustMst!CmAdd1
    ADC("InCmAdd2") = wRsCustMst!CMADD2
    ADC("InCmAdd3") = wRsCustMst!CMADD3
    ADC("InCmTel") = wRsCustMst!CMTEL
    ADC("InCmFax") = wRsCustMst!CmFax
    ADC("InCmEmail") = wRsCustMst!CmEMail
    ADC("InCmCpName") = wRsCustMst!CmCpName
    ' ***** Manali 3.03 - 26/06/08
    ADC("InCmBuyer") = wRsCustMst!CmBuyer
    ADC("InBuyAdd1") = wRsCustMst!CmBuyAdd1
    ADC("InBuyAdd2") = wRsCustMst!CmBuyAdd2
    ADC("InBuyAdd3") = wRsCustMst!CmBuyAdd3
    ' ***** Manali 3.03 - 26/06/08
    '***** Bef 3.03 - Manali ****
'    ADC("InCmBuyer") = ""
'    ADC("InBuyAdd1") = ""
'    ADC("InBuyAdd2") = ""
'    ADC("InBuyAdd3") = ""
    '***** Bef 3.03 - Manali ****
    ADC("InCstCd") = ""         '578-20
    ADC("InHConsignee") = ""
    ADC("InConAdd1") = ""
    ADC("InConAdd2") = ""
    ADC("InConAdd3") = ""
    ' ***** Manali 3.03 - 26/06/08 - Agent Code, Name and Address
    ADC("wAgCd") = ""
    ADC("InCmAgName") = ""
    ADC("InCmAgAdd1") = ""
    ADC("InCmAgAdd2") = ""
    ADC("InCmAgAdd3") = ""
    ' ***** Manali 3.03 - 26/06/08 - Agent Code, Name and Address
    ADC("InHFtRem") = wRsCustMst!CmInvFtRem     '4.1.2
    'vm.3 Getting Multiply Factor to invoice from custmst
    ADC("InOrdMulbyYN") = wRsCustMst!CmOrdMulByYN
    'vk.24 getting value from customer
    ADC("InByBagYN") = wRsCustMst!CmInvByBagYN
    Select Case ADC("InhInvTyp")
    Case Is = "DR"
    '***** Bef 3.03 - Manali ****
'        ADC("InCmBuyer") = ""
'        ADC("InBuyAdd1") = ""
'        ADC("InBuyAdd2") = ""
'        ADC("InBuyAdd3") = ""]
    '***** Bef 3.03 - Manali ****
        ADC("InHConsignee") = "Direct Shipment"
        ADC("InConAdd1") = ""
        ADC("InConAdd2") = ""
        ADC("InConAdd3") = ""
    ' ***** Manali 3.03 - 26/06/08 - Agent Code, Name and Address
        ADC("wAgCd") = ""
        ADC("InCmAgName") = ""
        ADC("InCmAgAdd1") = ""
        ADC("InCmAgAdd2") = ""
        ADC("InCmAgAdd3") = ""
    ' ***** Manali 3.03 - 26/06/08 - Agent Code, Name and Address
    Case Is = "BS"
        Set wRsBnkMst = moCn.OpenRes("Select * From BnkAgMst where BkCtg = 'B' and BkCd = '" + wRsCustMst!CmBkCd + "'")
        '***** Bef 3.03 - Manali ****
'        ADC("InCmBuyer") = ""
'        ADC("InBuyAdd1") = ""
'        ADC("InBuyAdd2") = ""
'        ADC("InBuyAdd3") = ""
        '***** Bef 3.03 - Manali ****
        If Not wRsBnkMst.EOF Or wRsBnkMst.BOF Then
            ADC("InHConsignee") = wRsBnkMst!BkName
            ADC("InConAdd1") = wRsBnkMst!BkAdd1
            ADC("InConAdd2") = wRsBnkMst!BkAdd2
            ADC("InConAdd3") = wRsBnkMst!BkAdd3
        End If
    ' ***** Manali 3.03 - 26/06/08 - Agent Code, Name and Address
        ADC("wAgCd") = ""
        ADC("InCmAgName") = ""
        ADC("InCmAgAdd1") = ""
        ADC("InCmAgAdd2") = ""
        ADC("InCmAgAdd3") = ""
    ' ***** Manali 3.03 - 26/06/08 - Agent Code, Name and Address
    Case Is = "DA"
        Set wRsBnkMst = moCn.OpenRes("Select * From BnkAgMst where BkCtg = 'B' and BkCd = '" + wRsCustMst!CmBkCd + "'")
        Set wRsBnkAgMst = moCn.OpenRes("Select * From BnkAgMst where BkCtg = 'A' and BkCd = '" + wRsCustMst!CmAgCd + "'")
        '***** Bef 3.03 - Manali ****
'        If Not wRsBnkAgMst.EOF Or wRsBnkAgMst.BOF Then
'            ADC("InCmBuyer") = wRsBnkAgMst!BkName
'            ADC("InBuyAdd1") = wRsBnkAgMst!BkAdd1
'            ADC("InBuyAdd2") = wRsBnkAgMst!BkAdd2
'            ADC("InBuyAdd3") = wRsBnkAgMst!BkAdd3
'        End If
        '***** Bef 3.03 - Manali ****
        ADC("InHConsignee") = "Direct Shipment"
        ADC("InConAdd1") = ""
        ADC("InConAdd2") = ""
        ADC("InConAdd3") = ""
        ' ***** Manali 3.03 - 26/06/08 - Agent Code, Name and Address
        If Not wRsBnkAgMst.EOF Or wRsBnkAgMst.BOF Then
            ADC("wAgCd") = wRsBnkAgMst!BkCd
            ADC("InCmAgName") = wRsBnkAgMst!BkName
            ADC("InCmAgAdd1") = wRsBnkAgMst!BkAdd1
            ADC("InCmAgAdd2") = wRsBnkAgMst!BkAdd2
            ADC("InCmAgAdd3") = wRsBnkAgMst!BkAdd3
        End If
        ' ***** Manali 3.03 - 26/06/08 - Agent Code, Name and Address
    Case Is = "BA"
        Set wRsBnkAgMst = moCn.OpenRes("Select * From BnkAgMst where BkCtg = 'A' and BkCd = '" + wRsCustMst!CmAgCd + "'")
        Set wRsBnkMst = moCn.OpenRes("Select * From BnkAgMst where BkCtg = 'B' and BkCd = '" + wRsCustMst!CmBkCd + "'")
       '***** Bef 3.03 - Manali ****
'        If Not wRsBnkAgMst.EOF Or wRsBnkAgMst.BOF Then
'            ADC("InCmBuyer") = wRsBnkAgMst!BkName
'            ADC("InBuyAdd1") = wRsBnkAgMst!BkAdd1
'            ADC("InBuyAdd2") = wRsBnkAgMst!BkAdd2
'            ADC("InBuyAdd3") = wRsBnkAgMst!BkAdd3
'        End If
        '***** Bef 3.03 - Manali ****
        If Not wRsBnkMst.EOF Or wRsBnkMst.BOF Then
            ADC("InHConsignee") = wRsBnkMst!BkName
            ADC("InConAdd1") = wRsBnkMst!BkAdd1
            ADC("InConAdd2") = wRsBnkMst!BkAdd2
            ADC("InConAdd3") = wRsBnkMst!BkAdd3
        End If
        ' ***** Manali 3.03 - 26/06/08 - Agent Code, Name and Address
        If Not wRsBnkAgMst.EOF Or wRsBnkAgMst.BOF Then
            ADC("wAgCd") = wRsBnkAgMst!BkCd
            ADC("InCmAgName") = wRsBnkAgMst!BkName
            ADC("InCmAgAdd1") = wRsBnkAgMst!BkAdd1
            ADC("InCmAgAdd2") = wRsBnkAgMst!BkAdd2
            ADC("InCmAgAdd3") = wRsBnkAgMst!BkAdd3
        End If
        ' ***** Manali 3.03 - 26/06/08 - Agent Code, Name and Address
End Select
Set wRsCustMst = Nothing
Set wRsBnkMst = Nothing
Set wRsBnkAgMst = Nothing
End Sub
'***** Shilpa ******

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
    gs_InCoCd = "": gs_InTc = "": gs_InYy = "": gs_InChr = "": gs_InNo = 0
    
End Sub

'InTc             InYy            InChr          InNo
'InExpNo          InCmCd          InDt           InLmgCst
'InLmgDt          InLmgSal        InVaDC         InVaX
'InCnvRt          InPrdDesc       InPkgWt        InGrNo
'InGrDt           InSbNo          InSbDt         InAwbNo
'InAwbDt          InGspNo         InGspDt        InChqNo
'InChqAmt         InPoNo          InPoDt         InCmPayTerms
'InFOB            InFOBRs         InComm         InCommRs
'InFrt            InFrtRs         InIns          InInsRs
'InTot            InTotRs         InTotAs        InDiaCost
'InCSCost         InAdv           InDiaCostDesc  InCSCostDesc
'InAdvDesc        InCls           InHBkName      InHBkAdd1
'InHBkAdd2        InHBkAdd3       InHBkAc        InHFrgnBkName
'InHFrgnBkAc      InHExpOff       InHRbiNo       InHIecNo
'InHPreCarriage   InHPreCarrier   InHVessel      InHPortOfLd
'InHConOfOrg      InHMarks        InHPkgDesc     InHItcNo
'InHEtcZone       InHEtcSection   InHConsignee   InCmInsBy
'InHInsNo         InHAgName       InHAgAdd1      InHAgAdd2
'InHAgAdd3        InCmName        InCmAdd1       InCmAdd2
'InCmAdd3         InCmTel         InCmFax        InCmEMail
'InCmCpName       InCmPortOfDisch InCmFinDest    InCmBuyer
'InSelfName       InSelfAdd1      InSelfAdd2     InSelfAdd3
'InSelfTel        InSelfFax       InSelfEMail

Private Function GetOriginCriteria() As String
  GetOriginCriteria = moCn.GetFldVal("Select Case isNull(PValue,'') When '' Then HGspOrgCrit  Else isNull(PValue,'') End" & _
                      " From (Select HCd , HCoCd, HGspOrgCrit from Head) H Left Outer Join Param On HCd ='" & ctSelfCmCd & "' " & _
                      " And PTyp = 'Region' And PmCd = (Select CmRegnCd From CustMst Where CmCtg = 'C' " & _
                      " And CmCd = '" & ADC("InCmCd") & "') And PsCd = '' Where HCoCd = '" & ms_InCoCd & "'")
End Function
Private Sub BillCoDets(mCmCd As String, Optional mInvCmCd As Variant)
    Dim wRsSelfCustMst As MwfLib.MDORowSet
    Dim wRsSelfHCoCd As MwfLib.MDORowSet
    Dim wRsCustMst As MwfLib.MDORowSet
    If IsMissing(mCmCd) Then Exit Sub
    If IsMissing(mInvCmCd) Then mInvCmCd = ADC("InCmCd")
    Set wRsSelfCustMst = moCn.OpenRes("Select * From CustMst where CmCtg= 'Z' and CmCd= '" + mCmCd + "'")
    If Not (wRsSelfCustMst.EOF Or wRsSelfCustMst.BOF) Then
      ADC("InSelfName") = wRsSelfCustMst!CmName
      ADC("InSelfAdd1") = wRsSelfCustMst!CmAdd1
      ADC("InSelfAdd2") = wRsSelfCustMst!CMADD2
      ADC("InSelfAdd3") = wRsSelfCustMst!CMADD3
      ADC("InSelfTel") = wRsSelfCustMst!CMTEL
      ADC("InSelfFax") = wRsSelfCustMst!CmFax
      ADC("InSelfEMail") = wRsSelfCustMst!CmEMail
    End If
    
    Set wRsSelfHCoCd = moCn.OpenRes("Select * From Head where HCoCd= '" + mCmCd + "' and HCd= '" + ctSelfCmCd + "' ")
    If Not (wRsSelfHCoCd.EOF Or wRsSelfHCoCd.BOF) Then
      ADC("InHVatNo") = wRsSelfHCoCd!HVATNo
      ADC("InHCstNo") = wRsSelfHCoCd!HCstNo
      ADC("InHLclInvFT") = wRsSelfHCoCd!HLclInvFT
      ADC("InHLclInvVatFT") = wRsSelfHCoCd!HLclInvVATFT
      ADC("InHLclInvCstFT") = wRsSelfHCoCd!HLclInvCSTFT
      ' ****** Manali 3.10.0 - 23/03/12 - Self TinNo and PAN Added
      ADC("InHTINNo") = wRsSelfHCoCd!HTinNo
      ADC("InHPAN") = wRsSelfHCoCd!HPAN
      
      ADC("InLclInvTyp") = wRsSelfHCoCd!HInLclInvTyp      ' ***** Manali 3.10.0 - 24/05/12 - Default from Head
    End If
    
    
    ' ****** Manali 3.10.0 - 23/03/12 - Cust VAT, CST and PAN Added
    Set wRsCustMst = moCn.OpenRes("Select * From CustMst where CmCtg= 'C' and CmCd= '" + mInvCmCd + "'")
    If Not (wRsCustMst.EOF Or wRsCustMst.BOF) Then
      ADC("InCmVatNo") = wRsCustMst!CmVATNo
      ADC("InCmCstNo") = wRsCustMst!CmCstNo
      ADC("InCmTINNo") = wRsCustMst!CmTinNo
      ADC("InCmPAN") = wRsCustMst!CmPAN
    
    End If
    ' Bef 3.10.0
    '''ADC("InCmTinNo") = moCn.GetFldVal("Select CmTinNo From CustMst where CmCtg= 'C' and CmCd= '" + ADC("InCmCd") + "'")      ' ***** Manali 3.9.0 - Cust Tin No For Local
    
End Sub
Private Sub InvAdditionalAmtDetsFill(ps_GstCd As Variant)
  Dim wRsAddAmt As MwfLib.MDORowSet
  Dim wGstForJoin As String, wCustFax As String, wZSelfFax As String, wCmTcsYN As String, wTCSCnd As String
  'pg.11-compares first 2 char of fax in customer master for zself & voucher customer.
  'if both are match then local customer else outside customer. only those fields(L/O in GstFor) in GstFlds ptyp can display
  wCustFax = moCn.GetFldVal("Select SUBSTRING(CmFax, 1, 2) From CustMst where CmCtg='" + ms_CmCtg + "' and CmCd= '" + ADC("InCmCd") + "'")
  wZSelfFax = moCn.GetFldVal("Select SUBSTRING(CmFax, 1, 2) From CustMst where CmCtg='Z' and CmCd= '" + ms_InCoCd + "'")
  wCmTcsYN = moCn.GetFldVal("Select CmTCSYN From CustMst where CmCtg='" + ms_CmCtg + "' and CmCd= '" + ADC("InCmCd") + "'")
  wTCSCnd = IIF(wCmTcsYN = "Y", " or PValue1='T'", "")
  If wCustFax = wZSelfFax Then
    wGstForJoin = " and (PValue1='-' or PValue1='R' or PValue1='L' " + wTCSCnd + " )"
  Else
    wGstForJoin = " and (PValue1='-' or PValue1='R' or PValue1='O' " + wTCSCnd + " )"
  End If
  
  Set wRsAddAmt = moCn.OpenRes("select * from Param where PTyp ='GSTFLDS' and PMCd ='" + ps_GstCd + "' " + wGstForJoin)
  'pg.11-Getting Description,% & GstOn from Param, ptyp=GSTFLDS and displaying values to text boxes
    
  ADC("InAmtDesc1") = "": ADC("InAmtDesc2") = "": ADC("InAmtDesc3") = "": ADC("InAmtDesc4") = ""
  ADC("InAmtDesc5") = "": ADC("InAmtDesc6") = "": ADC("InAmtDesc7") = "": ADC("InAmtDesc8") = ""
  ADC("InAmtDesc9") = "": ADC("InAmtDesc10") = "": ADC("InAmtDesc11") = "": ADC("InAmtDesc12") = ""
  ADC("InAmt1Per") = 0:  ADC("InAmt2Per") = 0:  ADC("InAmt3Per") = 0:  ADC("InAmt4Per") = 0
  ADC("InAmt5Per") = 0:  ADC("InAmt6Per") = 0:  ADC("InAmt8Per") = 0:  ADC("InAmt9Per") = 0
  ADC("InAmt10Per") = 0:  ADC("InAmt11Per") = 0
  ADC("InGstOn1") = "":  ADC("InGstOn2") = "":  ADC("InGstOn3") = "":  ADC("InGstOn4") = ""
  ADC("InGstOn5") = "":  ADC("InGstOn6") = "":  ADC("InGstOn8") = "":  ADC("InGstOn9") = ""
  ADC("InGstOn10") = "":  ADC("InGstOn11") = "": ADC("InAmt1") = 0: ADC("InAmt2") = 0:
  ADC("InAmt3") = 0: ADC("InAmt4") = 0: ADC("InAmt5") = 0: ADC("InAmt6") = 0: ADC("InAmt7") = 0:
  ADC("InAmt8") = 0: ADC("InAmt9") = 0: ADC("InAmt10") = 0: ADC("InAmt11") = 0: ADC("InAmt12") = 0
  ADC("InGstBase") = moCn.GetFldVal("select PValue from Param where PTyp ='GSTCD' and PMCd='" + ADC("InGstCd") + "'")
      
  Do While Not (wRsAddAmt.EOF Or wRsAddAmt.BOF)
    If wRsAddAmt!PSCd = 1 Then
      ADC("InAmtDesc1") = wRsAddAmt!PDesc: ADC("InGstOn1") = wRsAddAmt!pValue: ADC("InAmt1Per") = wRsAddAmt!PNum
    End If
    If wRsAddAmt!PSCd = 2 Then
      ADC("InAmtDesc2") = wRsAddAmt!PDesc: ADC("InGstOn2") = wRsAddAmt!pValue: ADC("InAmt2Per") = wRsAddAmt!PNum
    End If
    If wRsAddAmt!PSCd = 3 Then
      ADC("InAmtDesc3") = wRsAddAmt!PDesc: ADC("InGstOn3") = wRsAddAmt!pValue: ADC("InAmt3Per") = wRsAddAmt!PNum
    End If
    If wRsAddAmt!PSCd = 4 Then
      ADC("InAmtDesc4") = wRsAddAmt!PDesc: ADC("InGstOn4") = wRsAddAmt!pValue: ADC("InAmt4Per") = wRsAddAmt!PNum
    End If
    If wRsAddAmt!PSCd = 5 Then
      ADC("InAmtDesc5") = wRsAddAmt!PDesc: ADC("InGstOn5") = wRsAddAmt!pValue: ADC("InAmt5Per") = wRsAddAmt!PNum
    End If
    If wRsAddAmt!PSCd = 6 Then
      ADC("InAmtDesc6") = wRsAddAmt!PDesc: ADC("InGstOn6") = wRsAddAmt!pValue: ADC("InAmt6Per") = wRsAddAmt!PNum
    End If
    If wRsAddAmt!PSCd = 7 Then  '6.1B Gston and amount is added
      ADC("InAmtDesc7") = wRsAddAmt!PDesc: ADC("InGstOn7") = wRsAddAmt!pValue: ADC("InAmt7Per") = wRsAddAmt!PNum
    End If
    
    If wRsAddAmt!PSCd = 8 Then
      ADC("InAmtDesc8") = wRsAddAmt!PDesc: ADC("InGstOn8") = wRsAddAmt!pValue: ADC("InAmt8Per") = wRsAddAmt!PNum
    End If
    If wRsAddAmt!PSCd = 9 Then
      ADC("InAmtDesc9") = wRsAddAmt!PDesc: ADC("InGstOn9") = wRsAddAmt!pValue: ADC("InAmt9Per") = wRsAddAmt!PNum
    End If
    If wRsAddAmt!PSCd = 10 Then
      ADC("InAmtDesc10") = wRsAddAmt!PDesc: ADC("InGstOn10") = wRsAddAmt!pValue: ADC("InAmt10Per") = wRsAddAmt!PNum
    End If
    If wRsAddAmt!PSCd = 11 Then
      ADC("InAmtDesc11") = wRsAddAmt!PDesc: ADC("InGstOn11") = wRsAddAmt!pValue: ADC("InAmt11Per") = wRsAddAmt!PNum
    End If
    If wRsAddAmt!PSCd = 12 Then '6.1B Gston and amount is added
      ADC("InAmtDesc12") = wRsAddAmt!PDesc: ADC("InGstOn12") = wRsAddAmt!pValue: ADC("InAmt12Per") = wRsAddAmt!PNum
    End If
    wRsAddAmt.MoveNext
  Loop
  Call InvAdditionalAmtsReCalc(False)
End Sub
Private Sub InvAdditionalAmtsReCalc(pRecalcOnlyIfBaseAmtChange As Boolean)
      Dim wGstBase As String
      Dim wOdOrdQty As Long
      
      wGstBase = ADC("InGstBase")
      Dim wInBaseAmt As Double
      If wGstBase = "INFOB" Then
        wInBaseAmt = moCn.GetFldVal("select InFOB from InvHd where InCoCd ='" + ms_InCoCd + _
                                "' and InTc ='" + ADC("InTc") + "' and InYy ='" + ADC("InYy") + "' and InChr ='" + _
                                ADC("InChr") + "' and InNo =" + CStr(ADC("InNo")) + "")
      ElseIf wGstBase = "OMCS" Then
      
        wOdOrdQty = moCn.GetFldVal("select OdOrdQty from OrdDsg join OrdMst on OmCoCd =OdCoCd and OmTc =OdTc " + _
                                "and OmYy =OdYy and OmChr =OdChr and OmNo =OdNo where OmCoCd='" + ms_InCoCd + _
                                "' and OmTc ='" + ADC("InTc") + "' and OmYy='" + ADC("InYy") + "' and OmChr ='" + _
                                ADC("InChr") + "' and OmNo =" + CStr(ADC("InNo")))

        wInBaseAmt = wOdOrdQty * moCn.GetFldVal("select Sum(OrSalVal) from OrdRm join OrdMst on OmCoCd =OrCoCd and OmTc =OrTc " + _
                                "and OmYy =OrYy and OmChr =OrChr and OmNo =OrNo where OmCoCd='" + ms_InCoCd + _
                                "' and OmTc ='" + ADC("InTc") + "' and OmYy='" + ADC("InYy") + "' and OmChr ='" + _
                                ADC("InChr") + "' and OmNo =" + CStr(ADC("InNo")) + " and OrRmCtg ='C'")
                                
      ElseIf wGstBase = "OMDIA" Then
        wInBaseAmt = wOdOrdQty * moCn.GetFldVal("select Sum(OrSalVal) from OrdRm join OrdMst on OmCoCd =OrCoCd and OmTc =OrTc " + _
                                "and OmYy =OrYy and OmChr =OrChr and OmNo =OrNo where OmCoCd='" + ms_InCoCd + _
                                "' and OmTc ='" + ADC("InTc") + "' and OmYy='" + ADC("InYy") + "' and OmChr ='" + _
                                ADC("InChr") + "' and OmNo =" + CStr(ADC("InNo")) + " and OrRmCtg ='D'")
                       
      ElseIf wGstBase = "OMDIACS" Then

        wOdOrdQty = moCn.GetFldVal("select OdOrdQty from OrdDsg join OrdMst on OmCoCd =OdCoCd and OmTc =OdTc " + _
                                "and OmYy =OdYy and OmChr =OdChr and OmNo =OdNo where OmCoCd='" + ms_InCoCd + _
                                "' and OmTc ='" + ADC("InTc") + "' and OmYy='" + ADC("InYy") + "' and OmChr ='" + _
                                ADC("InChr") + "' and OmNo =" + CStr(ADC("InNo")))
        
        wInBaseAmt = wOdOrdQty * moCn.GetFldVal("select Sum(OrSalVal) from OrdRm join OrdMst on OmCoCd =OrCoCd and OmTc =OrTc " + _
                                "and OmYy =OrYy and OmChr =OrChr and OmNo =OrNo where OmCoCd='" + ms_InCoCd + _
                                "' and OmTc ='" + ADC("InTc") + "' and OmYy='" + ADC("InYy") + "' and OmChr ='" + _
                                ADC("InChr") + "' and OmNo =" + CStr(ADC("InNo")) + " and (OrRmCtg ='D' or OrRmCtg ='C')")
                      
      ElseIf wGstBase = "OMJLVAL" Then
        wInBaseAmt = moCn.GetFldVal("select Sum(OdJLVchVal) from OrdDsg join OrdMst on OmCoCd =OdCoCd and OmTc =OdTc " + _
                                "and OmYy =OdYy and OmChr =OdChr and OmNo =OdNo where OmCoCd='" + ms_InCoCd + _
                                "' and OmTc ='" + ADC("InTc") + "' and OmYy='" + ADC("InYy") + "' and OmChr ='" + _
                                ADC("InChr") + "' and OmNo =" + CStr(ADC("InNo")))
                 
      ElseIf wGstBase = "OMLABVAL" Then
      
        wOdOrdQty = moCn.GetFldVal("select OdOrdQty from OrdDsg join OrdMst on OmCoCd =OdCoCd and OmTc =OdTc " + _
                                "and OmYy =OdYy and OmChr =OdChr and OmNo =OdNo where OmCoCd='" + ms_InCoCd + _
                                "' and OmTc ='" + ADC("InTc") + "' and OmYy='" + ADC("InYy") + "' and OmChr ='" + _
                                ADC("InChr") + "' and OmNo =" + CStr(ADC("InNo")))


        wInBaseAmt = moCn.GetFldVal("select Sum(OlSalVal) from OrdLab join OrdMst on OmCoCd =OlCoCd and OmTc =OlTc " + _
                                "and OmYy =OlYy and OmChr =OlChr and OmNo =OlNo where OmCoCd='" + ms_InCoCd + _
                                "' and OmTc ='" + ADC("InTc") + "' and OmYy='" + ADC("InYy") + "' and OmChr ='" + _
                                ADC("InChr") + "' and OmNo =" + CStr(ADC("InNo")))
        wInBaseAmt = wInBaseAmt + moCn.GetFldVal("select Sum(OrSetSalVal) from OrdRm join OrdMst on OmCoCd =OrCoCd and OmTc =OrTc " + _
                                "and OmYy =OrYy and OmChr =OrChr and OmNo =OrNo where OmCoCd='" + ms_InCoCd + _
                                "' and OmTc ='" + ADC("InTc") + "' and OmYy='" + ADC("InYy") + "' and OmChr ='" + _
                                ADC("InChr") + "' and OmNo =" + CStr(ADC("InNo")))
        wInBaseAmt = wInBaseAmt * wOdOrdQty
      ElseIf wGstBase = "OMRMVAL" Then
        wOdOrdQty = moCn.GetFldVal("select OdOrdQty from OrdDsg join OrdMst on OmCoCd =OdCoCd and OmTc =OdTc " + _
                                "and OmYy =OdYy and OmChr =OdChr and OmNo =OdNo where OmCoCd='" + ms_InCoCd + _
                                "' and OmTc ='" + ADC("InTc") + "' and OmYy='" + ADC("InYy") + "' and OmChr ='" + _
                                ADC("InChr") + "' and OmNo =" + CStr(ADC("InNo")))

        wInBaseAmt = wOdOrdQty * moCn.GetFldVal("select Sum(OrSalVal) from OrdRm join OrdMst on OmCoCd =OrCoCd and OmTc =OrTc " + _
                                "and OmYy =OrYy and OmChr =OrChr and OmNo =OrNo where OmCoCd='" + ms_InCoCd + _
                                "' and OmTc ='" + ADC("InTc") + "' and OmYy='" + ADC("InYy") + "' and OmChr ='" + _
                                ADC("InChr") + "' and OmNo =" + CStr(ADC("InNo")))
      ElseIf wGstBase = "OMSALVAL" Then
        wInBaseAmt = moCn.GetFldVal("select Sum(OdSalPrc*OdOrdQty) from OrdDsg join OrdMst on OmCoCd =OdCoCd and OmTc =OdTc " + _
                                "and OmYy =OdYy and OmChr =OdChr and OmNo =OdNo where OmCoCd='" + ms_InCoCd + _
                                "' and OmTc ='" + ADC("InTc") + "' and OmYy='" + ADC("InYy") + "' and OmChr ='" + _
                                ADC("InChr") + "' and OmNo =" + CStr(ADC("InNo")))
      ElseIf wGstBase = "TTOTAMT" Then
        wInBaseAmt = moCn.GetFldVal("select TTotAmt from Txn where TCoCd='" + ms_InCoCd + _
                                "' and TTc ='" + ADC("InTc") + "' and TYy='" + ADC("InYy") + "' and TChr ='" + _
                                ADC("InChr") + "' and TNo =" + CStr(ADC("InNo")))
      End If
      
      If wInBaseAmt = ADC("InBaseAmt") And pRecalcOnlyIfBaseAmtChange Then
        Exit Sub
      End If
      If ADC("InGstBase") <> "-" Then
        ADC("InBaseAmt") = wInBaseAmt
      End If
      Dim W As Integer
      W = 0
      For W = 1 To 12
        ADC("InAmt" + CStr(W)) = InvAmtCalc(W)
      Next W
End Sub
'pg.11- InAmt calculation based on Percentage, GstOn values
Private Function InvAmtCalc(pInAmtPos As Integer) As Double
  Dim wInAmt As Double
  Dim W As Integer
  Dim wRunningAmtTot As Double
  wRunningAmtTot = 0
  Dim wAmtBasedOnGstOn As Double
  Dim wAmtForCalc As Double
  Dim wGstOn As String, wInAmtPer As Double
  InvAmtCalc = 0
  If pInAmtPos <= 0 Or pInAmtPos > 12 Then Exit Function
  Dim wGstOnInDB As String, wGstFor As String, wRndOffAmt As Double, wRndOffSr As String
  Dim wGstOnArr() As String, i As Integer, wInvAmtCalc As Double
  Dim w_GstOn As String, wGstOnSign As String
  
  wGstOnInDB = ADC("InGstOn" + CStr(pInAmtPos))
  
  wGstOnArr = Split(wGstOnInDB, ",")
  If IsArray(wGstOnArr) Then
    For i = 0 To UBound(wGstOnArr)
      w_GstOn = IIF(Left(wGstOnArr(i), 1) = "-", Mid(wGstOnArr(i), 2, Len(wGstOnArr(i))), wGstOnArr(i))
      If IsNumeric(wGstOnArr(i)) Then
          wGstOnSign = IIF(Left(wGstOnArr(i), 1) = "-", "-", "")
          wAmtBasedOnGstOn = wAmtBasedOnGstOn + (wGstOnSign + CStr(ADC("InAmt" + CStr(w_GstOn))))
        Else
          wAmtBasedOnGstOn = wAmtBasedOnGstOn + IIF(wGstOnArr(i) = "B", ADC("InBaseAmt"), 0)
      End If
    Next i
  End If
  
  wInAmtPer = ADC("InAmt" + CStr(pInAmtPos) + "Per")
  
  If wInAmtPer = 0 Then
    wAmtForCalc = ADC("InAmt" + CStr(pInAmtPos))
  Else
    wAmtForCalc = wAmtBasedOnGstOn * (wInAmtPer / 100)
  End If
   
  If pInAmtPos = 12 Then
    wRndOffSr = moCn.GetFldVal("select top 1 PSCd from Param where PTyp ='GSTFLDS' and PMCd ='" + ADC("INGSTCD") + "' and PValue1 ='R'")
    If wRndOffSr <> "" Then
      wRndOffAmt = (Round(wAmtForCalc) - wAmtForCalc)
      ADC("InAmt" + CStr(wRndOffSr)) = ADC("InAmt" + CStr(wRndOffSr)) + wRndOffAmt
      InvAmtCalc = Round(wAmtForCalc)
    Else
      InvAmtCalc = wAmtForCalc
    End If
  Else
    InvAmtCalc = wAmtForCalc
  End If

End Function
Private Sub InRunAmt(pChgAmtNo As Integer, pInAmt1 As Double, pInAmt2 As Double, pInAmt3 As Double, pInAmt4 As Double, pInAmt5 As Double, pInAmt6 As Double, pInAmt8 As Double, pInAmt9 As Double, pInAmt10 As Double, pInAmt11 As Double)
'pg.11-when invoice amount is chnaged the rest will change automatically if it is related
  Select Case UCase(pChgAmtNo)
  Case Is = "1"
    If ADC("InGstOn2") = "R" Then ADC("InAmt2") = ((ADC("InBaseAmt") + pInAmt1) * ADC("InAmt2Per")) / 100
    If ADC("InGstOn3") = "R" Then ADC("InAmt3") = ((ADC("InBaseAmt") + pInAmt1 + ADC("InAmt2")) * ADC("InAmt3Per")) / 100
    If ADC("InGstOn4") = "R" Then ADC("InAmt4") = ((ADC("InBaseAmt") + pInAmt1 + ADC("InAmt2") + ADC("InAmt3")) * ADC("InAmt4Per")) / 100
    If ADC("InGstOn5") = "R" Then ADC("InAmt5") = ((ADC("InBaseAmt") + pInAmt1 + ADC("InAmt2") + ADC("InAmt3") + ADC("InAmt4")) * ADC("InAmt5Per")) / 100
    If ADC("InGstOn6") = "R" Then ADC("InAmt6") = ((ADC("InBaseAmt") + pInAmt1 + ADC("InAmt2") + ADC("InAmt3") + ADC("InAmt4") + ADC("InAmt5")) * ADC("InAmt6Per")) / 100
    If ADC("InGstOn8") = "R" Then ADC("InAmt8") = ((ADC("InBaseAmt") + pInAmt1 + ADC("InAmt2") + ADC("InAmt3") + ADC("InAmt4") + ADC("InAmt5") + ADC("InAmt6")) * ADC("InAmt8Per")) / 100
    If ADC("InGstOn9") = "R" Then ADC("InAmt9") = ((ADC("InBaseAmt") + pInAmt1 + ADC("InAmt2") + ADC("InAmt3") + ADC("InAmt4") + ADC("InAmt5") + ADC("InAmt6") + ADC("InAmt8")) * ADC("InAmt9Per")) / 100
    If ADC("InGstOn10") = "R" Then ADC("InAmt10") = ((ADC("InBaseAmt") + pInAmt1 + ADC("InAmt2") + ADC("InAmt3") + ADC("InAmt4") + ADC("InAmt5") + ADC("InAmt6") + ADC("InAmt8") + ADC("InAmt9")) * ADC("InAmt10Per")) / 100
    If ADC("InGstOn11") = "R" Then ADC("InAmt11") = (((ADC("InBaseAmt") + pInAmt1 + ADC("InAmt2") + ADC("InAmt3") + ADC("InAmt4") + ADC("InAmt5") + ADC("InAmt6")) + ADC("InAmt8") + ADC("InAmt9") + ADC("InAmt10")) * ADC("InAmt11Per")) / 100
  Case Is = "2"
    If ADC("InGstOn3") = "R" Then ADC("InAmt3") = ((ADC("InBaseAmt") + ADC("InAmt1") + pInAmt2) * ADC("InAmt3Per")) / 100
    If ADC("InGstOn4") = "R" Then ADC("InAmt4") = ((ADC("InBaseAmt") + ADC("InAmt1") + pInAmt2 + ADC("InAmt3")) * ADC("InAmt4Per")) / 100
    If ADC("InGstOn5") = "R" Then ADC("InAmt5") = ((ADC("InBaseAmt") + ADC("InAmt1") + pInAmt2 + ADC("InAmt3") + ADC("InAmt4")) * ADC("InAmt5Per")) / 100
    If ADC("InGstOn6") = "R" Then ADC("InAmt6") = ((ADC("InBaseAmt") + ADC("InAmt1") + pInAmt2 + ADC("InAmt3") + ADC("InAmt4") + ADC("InAmt5")) * ADC("InAmt6Per")) / 100
    If ADC("InGstOn8") = "R" Then ADC("InAmt8") = ((ADC("InBaseAmt") + ADC("InAmt1") + pInAmt2 + ADC("InAmt3") + ADC("InAmt4") + ADC("InAmt5") + ADC("InAmt6")) * ADC("InAmt8Per")) / 100
    If ADC("InGstOn9") = "R" Then ADC("InAmt9") = ((ADC("InBaseAmt") + ADC("InAmt1") + pInAmt2 + ADC("InAmt3") + ADC("InAmt4") + ADC("InAmt5") + ADC("InAmt6") + ADC("InAmt8")) * ADC("InAmt9Per")) / 100
    If ADC("InGstOn10") = "R" Then ADC("InAmt10") = ((ADC("InBaseAmt") + ADC("InAmt1") + pInAmt2 + ADC("InAmt3") + ADC("InAmt4") + ADC("InAmt5") + ADC("InAmt6") + ADC("InAmt8") + ADC("InAmt9")) * ADC("InAmt10Per")) / 100
    If ADC("InGstOn11") = "R" Then ADC("InAmt11") = (((ADC("InBaseAmt") + ADC("InAmt1") + pInAmt2 + ADC("InAmt3") + ADC("InAmt4") + ADC("InAmt5") + ADC("InAmt6")) + ADC("InAmt8") + ADC("InAmt9") + ADC("InAmt10")) * ADC("InAmt11Per")) / 100
  Case Is = "3"
    If ADC("InGstOn4") = "R" Then ADC("InAmt4") = ((ADC("InBaseAmt") + ADC("InAmt1") + ADC("InAmt2") + pInAmt3) * ADC("InAmt4Per")) / 100
    If ADC("InGstOn5") = "R" Then ADC("InAmt5") = ((ADC("InBaseAmt") + ADC("InAmt1") + ADC("InAmt2") + pInAmt3 + ADC("InAmt4")) * ADC("InAmt5Per")) / 100
    If ADC("InGstOn6") = "R" Then ADC("InAmt6") = ((ADC("InBaseAmt") + ADC("InAmt1") + ADC("InAmt2") + pInAmt3 + ADC("InAmt4") + ADC("InAmt5")) * ADC("InAmt6Per")) / 100
    If ADC("InGstOn8") = "R" Then ADC("InAmt8") = ((ADC("InBaseAmt") + ADC("InAmt1") + ADC("InAmt2") + pInAmt3 + ADC("InAmt4") + ADC("InAmt5") + ADC("InAmt6")) * ADC("InAmt8Per")) / 100
    If ADC("InGstOn9") = "R" Then ADC("InAmt9") = ((ADC("InBaseAmt") + ADC("InAmt1") + ADC("InAmt2") + pInAmt3 + ADC("InAmt4") + ADC("InAmt5") + ADC("InAmt6") + ADC("InAmt8")) * ADC("InAmt9Per")) / 100
    If ADC("InGstOn10") = "R" Then ADC("InAmt10") = ((ADC("InBaseAmt") + ADC("InAmt1") + ADC("InAmt2") + pInAmt3 + ADC("InAmt4") + ADC("InAmt5") + ADC("InAmt6") + ADC("InAmt8") + ADC("InAmt9")) * ADC("InAmt10Per")) / 100
    If ADC("InGstOn11") = "R" Then ADC("InAmt11") = (((ADC("InBaseAmt") + ADC("InAmt1") + ADC("InAmt2") + pInAmt3 + ADC("InAmt4") + ADC("InAmt5") + ADC("InAmt6")) + ADC("InAmt8") + ADC("InAmt9") + ADC("InAmt10")) * ADC("InAmt11Per")) / 100
   
  Case Is = "4"
    If ADC("InGstOn5") = "R" Then ADC("InAmt5") = ((ADC("InBaseAmt") + ADC("InAmt1") + ADC("InAmt2") + ADC("InAmt3") + pInAmt4) * ADC("InAmt5Per")) / 100
    If ADC("InGstOn6") = "R" Then ADC("InAmt6") = ((ADC("InBaseAmt") + ADC("InAmt1") + ADC("InAmt2") + ADC("InAmt3") + pInAmt4 + ADC("InAmt5")) * ADC("InAmt6Per")) / 100
    If ADC("InGstOn8") = "R" Then ADC("InAmt8") = ((ADC("InBaseAmt") + ADC("InAmt1") + ADC("InAmt2") + ADC("InAmt3") + pInAmt4 + ADC("InAmt5") + ADC("InAmt6")) * ADC("InAmt8Per")) / 100
    If ADC("InGstOn9") = "R" Then ADC("InAmt9") = ((ADC("InBaseAmt") + ADC("InAmt1") + ADC("InAmt2") + ADC("InAmt3") + pInAmt4 + ADC("InAmt5") + ADC("InAmt6") + ADC("InAmt8")) * ADC("InAmt9Per")) / 100
    If ADC("InGstOn10") = "R" Then ADC("InAmt10") = ((ADC("InBaseAmt") + ADC("InAmt1") + ADC("InAmt2") + ADC("InAmt3") + pInAmt4 + ADC("InAmt5") + ADC("InAmt6") + ADC("InAmt8") + ADC("InAmt9")) * ADC("InAmt10Per")) / 100
    If ADC("InGstOn11") = "R" Then ADC("InAmt11") = (((ADC("InBaseAmt") + ADC("InAmt1") + ADC("InAmt2") + ADC("InAmt3") + pInAmt4 + ADC("InAmt5") + ADC("InAmt6")) + ADC("InAmt8") + ADC("InAmt9") + ADC("InAmt10")) * ADC("InAmt11Per")) / 100
  Case Is = "5"
    If ADC("InGstOn6") = "R" Then ADC("InAmt6") = ((ADC("InBaseAmt") + ADC("InAmt1") + ADC("InAmt2") + ADC("InAmt3") + ADC("InAmt4") + pInAmt5) * ADC("InAmt6Per")) / 100
    If ADC("InGstOn8") = "R" Then ADC("InAmt8") = ((ADC("InBaseAmt") + ADC("InAmt1") + ADC("InAmt2") + ADC("InAmt3") + ADC("InAmt4") + pInAmt5 + ADC("InAmt6")) * ADC("InAmt8Per")) / 100
    If ADC("InGstOn9") = "R" Then ADC("InAmt9") = ((ADC("InBaseAmt") + ADC("InAmt1") + ADC("InAmt2") + ADC("InAmt3") + ADC("InAmt4") + pInAmt5 + ADC("InAmt6") + ADC("InAmt8")) * ADC("InAmt9Per")) / 100
    If ADC("InGstOn10") = "R" Then ADC("InAmt10") = ((ADC("InBaseAmt") + ADC("InAmt1") + ADC("InAmt2") + ADC("InAmt3") + ADC("InAmt4") + pInAmt5 + ADC("InAmt6") + ADC("InAmt8") + ADC("InAmt9")) * ADC("InAmt10Per")) / 100
    If ADC("InGstOn11") = "R" Then ADC("InAmt11") = (((ADC("InBaseAmt") + ADC("InAmt1") + ADC("InAmt2") + ADC("InAmt3") + ADC("InAmt4") + pInAmt5 + ADC("InAmt6")) + ADC("InAmt8") + ADC("InAmt9") + ADC("InAmt10")) * ADC("InAmt11Per")) / 100
  Case Is = "6"
    If ADC("InGstOn8") = "R" Then ADC("InAmt8") = ((ADC("InBaseAmt") + ADC("InAmt1") + ADC("InAmt2") + ADC("InAmt3") + ADC("InAmt4") + ADC("InAmt5") + pInAmt6) * ADC("InAmt8Per")) / 100
    If ADC("InGstOn9") = "R" Then ADC("InAmt9") = ((ADC("InBaseAmt") + ADC("InAmt1") + ADC("InAmt2") + ADC("InAmt3") + ADC("InAmt4") + ADC("InAmt5") + pInAmt6 + ADC("InAmt8")) * ADC("InAmt9Per")) / 100
    If ADC("InGstOn10") = "R" Then ADC("InAmt10") = ((ADC("InBaseAmt") + ADC("InAmt1") + ADC("InAmt2") + ADC("InAmt3") + ADC("InAmt4") + ADC("InAmt5") + pInAmt6 + ADC("InAmt8") + ADC("InAmt9")) * ADC("InAmt10Per")) / 100
    If ADC("InGstOn11") = "R" Then ADC("InAmt11") = ((ADC("InBaseAmt") + ADC("InAmt1") + ADC("InAmt2") + ADC("InAmt3") + ADC("InAmt4") + ADC("InAmt5") + pInAmt6 + ADC("InAmt8") + ADC("InAmt9") + ADC("InAmt10")) * ADC("InAmt11Per")) / 100
  Case Is = "8"
    If ADC("InGstOn9") = "R" Then ADC("InAmt9") = ((ADC("InBaseAmt") + ADC("InAmt1") + ADC("InAmt2") + ADC("InAmt3") + ADC("InAmt4") + ADC("InAmt5") + ADC("InAmt6") + pInAmt8) * ADC("InAmt9Per")) / 100
    If ADC("InGstOn10") = "R" Then ADC("InAmt10") = ((ADC("InBaseAmt") + ADC("InAmt1") + ADC("InAmt2") + ADC("InAmt3") + ADC("InAmt4") + ADC("InAmt5") + ADC("InAmt6") + pInAmt8 + ADC("InAmt9")) * ADC("InAmt10Per")) / 100
    If ADC("InGstOn11") = "R" Then ADC("InAmt11") = (((ADC("InBaseAmt") + ADC("InAmt1") + ADC("InAmt2") + ADC("InAmt3") + ADC("InAmt4") + ADC("InAmt5") + ADC("InAmt6")) + pInAmt8 + ADC("InAmt9") + ADC("InAmt10")) * ADC("InAmt11Per")) / 100
  Case Is = "9"
    If ADC("InGstOn10") = "R" Then ADC("InAmt10") = ((ADC("InBaseAmt") + ADC("InAmt1") + ADC("InAmt2") + ADC("InAmt3") + ADC("InAmt4") + ADC("InAmt5") + ADC("InAmt6") + ADC("InAmt8") + pInAmt9) * ADC("InAmt10Per")) / 100
    If ADC("InGstOn11") = "R" Then ADC("InAmt11") = (((ADC("InBaseAmt") + ADC("InAmt1") + ADC("InAmt2") + ADC("InAmt3") + ADC("InAmt4") + ADC("InAmt5") + ADC("InAmt6")) + ADC("InAmt8") + pInAmt9 + ADC("InAmt10")) * ADC("InAmt11Per")) / 100
  Case Is = "10"
    If ADC("InGstOn11") = "R" Then ADC("InAmt11") = (((ADC("InBaseAmt") + ADC("InAmt1") + ADC("InAmt2") + ADC("InAmt3") + ADC("InAmt4") + ADC("InAmt5") + ADC("InAmt6")) + ADC("InAmt8") + ADC("InAmt9") + pInAmt10) * ADC("InAmt11Per")) / 100
    
  End Select

End Sub

Private Function InBaseChg(pGstOn As String, pGstPer As Double, pInAmtNo As Integer, pBase As Double) As Double
Dim wInAmt As Double
  Select Case UCase(pGstOn)
  Case Is = "1"
    wInAmt = (ADC("InAmt1") * pGstPer) / 100
  Case Is = "2"
    wInAmt = (ADC("InAmt2") * pGstPer) / 100
  Case Is = "3"
    wInAmt = (ADC("InAmt3") * pGstPer) / 100
  Case Is = "4"
    wInAmt = (ADC("InAmt4") * pGstPer) / 100
  Case Is = "5"
    wInAmt = (ADC("InAmt5") * pGstPer) / 100
  Case Is = "6"
    wInAmt = (ADC("InAmt6") * pGstPer) / 100
  Case Is = "7"
    wInAmt = (ADC("InAmt7") * pGstPer) / 100
  Case Is = "8"
    wInAmt = (ADC("InAmt8") * pGstPer) / 100
  Case Is = "9"
    wInAmt = (ADC("InAmt9") * pGstPer) / 100
  Case Is = "10"
    wInAmt = (ADC("InAmt10") * pGstPer) / 100
  Case Is = "11"
    wInAmt = (ADC("InAmt11") * pGstPer) / 100
  Case Is = "B"
    wInAmt = (pBase * pGstPer) / 100
  Case Is = "R"
    If pInAmtNo = 1 Then wInAmt = (pBase * pGstPer) / 100
    If pInAmtNo = 2 Then wInAmt = ((pBase + ADC("InAmt1")) * pGstPer) / 100
    If pInAmtNo = 3 Then wInAmt = ((pBase + ADC("InAmt1") + ADC("InAmt2")) * pGstPer) / 100
    If pInAmtNo = 4 Then wInAmt = ((pBase + ADC("InAmt1") + ADC("InAmt2") + ADC("InAmt3")) * pGstPer) / 100
    If pInAmtNo = 5 Then wInAmt = ((pBase + ADC("InAmt1") + ADC("InAmt2") + ADC("InAmt3") + ADC("InAmt4")) * pGstPer) / 100
    If pInAmtNo = 6 Then wInAmt = ((pBase + ADC("InAmt1") + ADC("InAmt2") + ADC("InAmt3") + ADC("InAmt4") + ADC("InAmt5")) * pGstPer) / 100
    If pInAmtNo = 8 Then wInAmt = ((pBase + ADC("InAmt1") + ADC("InAmt2") + ADC("InAmt3") + ADC("InAmt4") + ADC("InAmt5") + ADC("InAmt6")) * pGstPer) / 100
    If pInAmtNo = 9 Then wInAmt = ((pBase + ADC("InAmt1") + ADC("InAmt2") + ADC("InAmt3") + ADC("InAmt4") + ADC("InAmt5") + ADC("InAmt6") + ADC("InAmt8")) * pGstPer) / 100
    If pInAmtNo = 10 Then wInAmt = ((pBase + ADC("InAmt1") + ADC("InAmt2") + ADC("InAmt3") + ADC("InAmt4") + ADC("InAmt5") + ADC("InAmt6") + ADC("InAmt8") + ADC("InAmt9")) * pGstPer) / 100
    If pInAmtNo = 11 Then wInAmt = ((pBase + ADC("InAmt1") + ADC("InAmt2") + ADC("InAmt3") + ADC("InAmt4") + ADC("InAmt5") + ADC("InAmt6") + ADC("InAmt8") + ADC("InAmt9") + ADC("InAmt10")) * pGstPer) / 100
  Case Is = "-"
    wInAmt = (pBase * pGstPer) / 100
  End Select
  InBaseChg = wInAmt
  
End Function


