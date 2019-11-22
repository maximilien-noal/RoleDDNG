VERSION 5.00
Object = "{D76D7130-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7d.ocx"
Begin VB.Form frmDonsComp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Affecter les dons et les points de compétences"
   ClientHeight    =   6735
   ClientLeft      =   4095
   ClientTop       =   3480
   ClientWidth     =   16650
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6735
   ScaleWidth      =   16650
   StartUpPosition =   1  'CenterOwner
   Begin VB.HScrollBar Spincomp 
      Height          =   255
      Index           =   46
      Left            =   10530
      TabIndex        =   343
      Top             =   435
      Width           =   360
   End
   Begin VB.HScrollBar Spincomp 
      Height          =   255
      Index           =   47
      Left            =   10530
      TabIndex        =   342
      Top             =   705
      Width           =   360
   End
   Begin VB.HScrollBar Spincomp 
      Height          =   255
      Index           =   48
      Left            =   10530
      TabIndex        =   341
      Top             =   975
      Width           =   360
   End
   Begin VB.HScrollBar Spincomp 
      Height          =   255
      Index           =   49
      Left            =   10530
      TabIndex        =   340
      Top             =   1245
      Width           =   360
   End
   Begin VB.HScrollBar Spincomp 
      Height          =   255
      Index           =   50
      Left            =   10530
      TabIndex        =   339
      Top             =   1515
      Width           =   360
   End
   Begin VB.HScrollBar Spincomp 
      Height          =   255
      Index           =   51
      Left            =   10530
      TabIndex        =   338
      Top             =   1785
      Width           =   360
   End
   Begin VB.HScrollBar Spincomp 
      Height          =   255
      Index           =   52
      Left            =   10530
      TabIndex        =   337
      Top             =   2055
      Width           =   360
   End
   Begin VB.HScrollBar Spincomp 
      Height          =   255
      Index           =   53
      Left            =   10530
      TabIndex        =   336
      Top             =   2325
      Width           =   360
   End
   Begin VB.HScrollBar Spincomp 
      Height          =   255
      Index           =   54
      Left            =   10530
      TabIndex        =   335
      Top             =   2595
      Width           =   360
   End
   Begin VB.HScrollBar Spincomp 
      Height          =   300
      Index           =   55
      Left            =   10530
      TabIndex        =   334
      Top             =   2865
      Width           =   360
   End
   Begin VB.HScrollBar Spincomp 
      Height          =   255
      Index           =   56
      Left            =   10530
      TabIndex        =   333
      Top             =   3135
      Width           =   360
   End
   Begin VB.HScrollBar Spincomp 
      Height          =   255
      Index           =   57
      Left            =   10530
      TabIndex        =   332
      Top             =   3405
      Width           =   360
   End
   Begin VB.HScrollBar Spincomp 
      Height          =   255
      Index           =   58
      Left            =   10530
      TabIndex        =   331
      Top             =   3675
      Width           =   360
   End
   Begin VB.HScrollBar Spincomp 
      Height          =   300
      Index           =   59
      Left            =   10530
      TabIndex        =   330
      Top             =   3945
      Width           =   360
   End
   Begin VB.HScrollBar Spincomp 
      Height          =   255
      Index           =   60
      Left            =   10530
      TabIndex        =   329
      Top             =   4215
      Width           =   360
   End
   Begin VB.HScrollBar Spincomp 
      Height          =   255
      Index           =   61
      Left            =   10530
      TabIndex        =   328
      Top             =   4485
      Width           =   360
   End
   Begin VB.HScrollBar Spincomp 
      Height          =   255
      Index           =   62
      Left            =   10530
      TabIndex        =   327
      Top             =   4755
      Width           =   360
   End
   Begin VB.HScrollBar Spincomp 
      Height          =   255
      Index           =   63
      Left            =   10530
      TabIndex        =   326
      Top             =   5025
      Width           =   360
   End
   Begin VB.HScrollBar Spincomp 
      Height          =   255
      Index           =   64
      Left            =   10530
      TabIndex        =   325
      Top             =   5295
      Width           =   360
   End
   Begin VB.HScrollBar Spincomp 
      Height          =   255
      Index           =   65
      Left            =   10530
      TabIndex        =   324
      Top             =   5565
      Width           =   360
   End
   Begin VB.HScrollBar Spincomp 
      Height          =   255
      Index           =   66
      Left            =   10530
      TabIndex        =   323
      Top             =   5835
      Width           =   360
   End
   Begin VB.HScrollBar Spincomp 
      Height          =   255
      Index           =   67
      Left            =   10530
      TabIndex        =   322
      Top             =   6105
      Width           =   360
   End
   Begin VB.HScrollBar Spincomp 
      Height          =   255
      Index           =   68
      Left            =   10530
      TabIndex        =   321
      Top             =   6360
      Width           =   360
   End
   Begin VB.Frame Frame2 
      Caption         =   "Dons"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6165
      Left            =   10920
      TabIndex        =   195
      Top             =   0
      Width           =   5655
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1200
         ScaleHeight     =   240
         ScaleWidth      =   2535
         TabIndex        =   198
         Top             =   2880
         Width           =   2535
         Begin VB.CommandButton btnEnlever 
            Caption         =   "Enlever"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1440
            Picture         =   "frmDonsComp.frx":0000
            TabIndex        =   200
            Top             =   0
            Width           =   795
         End
         Begin VB.CommandButton btnAjouter 
            Caption         =   "Ajouter"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   480
            TabIndex        =   199
            Top             =   0
            Width           =   795
         End
      End
      Begin VSFlex7DAOCtl.VSFlexGrid vs1 
         Height          =   2625
         Left            =   120
         TabIndex        =   196
         Top             =   240
         Width           =   4665
         _cx             =   5080
         _cy             =   5080
         _ConvInfo       =   1
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483639
         BackColorBkg    =   -2147483636
         BackColorAlternate=   16777215
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   3
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   7
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   "Don du personnage                                          |Niveau"
         ScrollTrack     =   0   'False
         ScrollBars      =   2
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   0   'False
         AutoSizeMode    =   0
         AutoSearch      =   1
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   1
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   1
         ShowComboButton =   -1  'True
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
      End
      Begin VSFlex7DAOCtl.VSFlexGrid vs2 
         Height          =   2865
         Left            =   120
         TabIndex        =   197
         Top             =   3240
         Width           =   5490
         _cx             =   5080
         _cy             =   5080
         _ConvInfo       =   1
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483639
         BackColorBkg    =   -2147483636
         BackColorAlternate=   12648384
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   3
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   6
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   "Nom                                             |Genre                      |Source                  "
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   0   'False
         AutoSizeMode    =   0
         AutoSearch      =   1
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   1
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   1
         ShowComboButton =   -1  'True
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
      End
   End
   Begin VB.Timer Timer1 
      Left            =   11040
      Top             =   6240
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6720
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   10950
      Begin VB.HScrollBar Spincomp 
         Height          =   255
         Index           =   23
         Left            =   6930
         TabIndex        =   366
         Top             =   450
         Width           =   360
      End
      Begin VB.HScrollBar Spincomp 
         Height          =   255
         Index           =   24
         Left            =   6930
         TabIndex        =   365
         Top             =   720
         Width           =   360
      End
      Begin VB.HScrollBar Spincomp 
         Height          =   255
         Index           =   25
         Left            =   6930
         TabIndex        =   364
         Top             =   990
         Width           =   360
      End
      Begin VB.HScrollBar Spincomp 
         Height          =   255
         Index           =   26
         Left            =   6930
         TabIndex        =   363
         Top             =   1260
         Width           =   360
      End
      Begin VB.HScrollBar Spincomp 
         Height          =   255
         Index           =   27
         Left            =   6930
         TabIndex        =   362
         Top             =   1530
         Width           =   360
      End
      Begin VB.HScrollBar Spincomp 
         Height          =   255
         Index           =   28
         Left            =   6930
         TabIndex        =   361
         Top             =   1800
         Width           =   360
      End
      Begin VB.HScrollBar Spincomp 
         Height          =   255
         Index           =   29
         Left            =   6930
         TabIndex        =   360
         Top             =   2070
         Width           =   360
      End
      Begin VB.HScrollBar Spincomp 
         Height          =   255
         Index           =   30
         Left            =   6930
         TabIndex        =   359
         Top             =   2340
         Width           =   360
      End
      Begin VB.HScrollBar Spincomp 
         Height          =   255
         Index           =   31
         Left            =   6930
         TabIndex        =   358
         Top             =   2610
         Width           =   360
      End
      Begin VB.HScrollBar Spincomp 
         Height          =   255
         Index           =   32
         Left            =   6930
         TabIndex        =   357
         Top             =   2880
         Width           =   360
      End
      Begin VB.HScrollBar Spincomp 
         Height          =   255
         Index           =   33
         Left            =   6930
         TabIndex        =   356
         Top             =   3150
         Width           =   360
      End
      Begin VB.HScrollBar Spincomp 
         Height          =   255
         Index           =   34
         Left            =   6930
         TabIndex        =   355
         Top             =   3420
         Width           =   360
      End
      Begin VB.HScrollBar Spincomp 
         Height          =   255
         Index           =   35
         Left            =   6930
         TabIndex        =   354
         Top             =   3690
         Width           =   360
      End
      Begin VB.HScrollBar Spincomp 
         Height          =   255
         Index           =   36
         Left            =   6930
         TabIndex        =   353
         Top             =   3960
         Width           =   360
      End
      Begin VB.HScrollBar Spincomp 
         Height          =   255
         Index           =   37
         Left            =   6930
         TabIndex        =   352
         Top             =   4230
         Width           =   360
      End
      Begin VB.HScrollBar Spincomp 
         Height          =   255
         Index           =   38
         Left            =   6930
         TabIndex        =   351
         Top             =   4500
         Width           =   360
      End
      Begin VB.HScrollBar Spincomp 
         Height          =   255
         Index           =   39
         Left            =   6930
         TabIndex        =   350
         Top             =   4770
         Width           =   360
      End
      Begin VB.HScrollBar Spincomp 
         Height          =   255
         Index           =   40
         Left            =   6930
         TabIndex        =   349
         Top             =   5040
         Width           =   360
      End
      Begin VB.HScrollBar Spincomp 
         Height          =   255
         Index           =   41
         Left            =   6930
         TabIndex        =   348
         Top             =   5310
         Width           =   360
      End
      Begin VB.HScrollBar Spincomp 
         Height          =   255
         Index           =   42
         Left            =   6930
         TabIndex        =   347
         Top             =   5580
         Width           =   360
      End
      Begin VB.HScrollBar Spincomp 
         Height          =   255
         Index           =   43
         Left            =   6930
         TabIndex        =   346
         Top             =   5850
         Width           =   360
      End
      Begin VB.HScrollBar Spincomp 
         Height          =   255
         Index           =   44
         Left            =   6930
         TabIndex        =   345
         Top             =   6120
         Width           =   360
      End
      Begin VB.HScrollBar Spincomp 
         Height          =   255
         Index           =   45
         Left            =   6930
         TabIndex        =   344
         Top             =   6390
         Width           =   360
      End
      Begin VB.HScrollBar Spincomp 
         Height          =   255
         Index           =   22
         Left            =   3340
         TabIndex        =   320
         Top             =   6390
         Width           =   360
      End
      Begin VB.HScrollBar Spincomp 
         Height          =   255
         Index           =   21
         Left            =   3340
         TabIndex        =   319
         Top             =   6120
         Width           =   360
      End
      Begin VB.HScrollBar Spincomp 
         Height          =   255
         Index           =   20
         Left            =   3340
         TabIndex        =   318
         Top             =   5850
         Width           =   360
      End
      Begin VB.HScrollBar Spincomp 
         Height          =   255
         Index           =   19
         Left            =   3340
         TabIndex        =   317
         Top             =   5580
         Width           =   360
      End
      Begin VB.HScrollBar Spincomp 
         Height          =   255
         Index           =   18
         Left            =   3340
         TabIndex        =   316
         Top             =   5310
         Width           =   360
      End
      Begin VB.HScrollBar Spincomp 
         Height          =   255
         Index           =   17
         Left            =   3340
         TabIndex        =   315
         Top             =   5040
         Width           =   360
      End
      Begin VB.HScrollBar Spincomp 
         Height          =   255
         Index           =   16
         Left            =   3340
         TabIndex        =   314
         Top             =   4770
         Width           =   360
      End
      Begin VB.HScrollBar Spincomp 
         Height          =   255
         Index           =   15
         Left            =   3340
         TabIndex        =   313
         Top             =   4500
         Width           =   360
      End
      Begin VB.HScrollBar Spincomp 
         Height          =   255
         Index           =   14
         Left            =   3340
         TabIndex        =   312
         Top             =   4230
         Width           =   360
      End
      Begin VB.HScrollBar Spincomp 
         Height          =   255
         Index           =   13
         Left            =   3340
         TabIndex        =   311
         Top             =   3960
         Width           =   360
      End
      Begin VB.HScrollBar Spincomp 
         Height          =   255
         Index           =   12
         Left            =   3340
         TabIndex        =   310
         Top             =   3690
         Width           =   360
      End
      Begin VB.HScrollBar Spincomp 
         Height          =   255
         Index           =   11
         Left            =   3340
         TabIndex        =   309
         Top             =   3420
         Width           =   360
      End
      Begin VB.HScrollBar Spincomp 
         Height          =   255
         Index           =   10
         Left            =   3340
         TabIndex        =   308
         Top             =   3150
         Width           =   360
      End
      Begin VB.HScrollBar Spincomp 
         Height          =   255
         Index           =   9
         Left            =   3340
         TabIndex        =   307
         Top             =   2880
         Width           =   360
      End
      Begin VB.HScrollBar Spincomp 
         Height          =   255
         Index           =   8
         Left            =   3340
         TabIndex        =   306
         Top             =   2610
         Width           =   360
      End
      Begin VB.HScrollBar Spincomp 
         Height          =   255
         Index           =   7
         Left            =   3340
         TabIndex        =   305
         Top             =   2340
         Width           =   360
      End
      Begin VB.HScrollBar Spincomp 
         Height          =   255
         Index           =   6
         Left            =   3340
         TabIndex        =   304
         Top             =   2070
         Width           =   360
      End
      Begin VB.HScrollBar Spincomp 
         Height          =   255
         Index           =   5
         Left            =   3340
         TabIndex        =   303
         Top             =   1800
         Width           =   360
      End
      Begin VB.HScrollBar Spincomp 
         Height          =   255
         Index           =   4
         Left            =   3340
         TabIndex        =   302
         Top             =   1530
         Width           =   360
      End
      Begin VB.HScrollBar Spincomp 
         Height          =   255
         Index           =   3
         Left            =   3340
         TabIndex        =   301
         Top             =   1260
         Width           =   360
      End
      Begin VB.HScrollBar Spincomp 
         Height          =   255
         Index           =   2
         Left            =   3340
         TabIndex        =   300
         Top             =   990
         Width           =   360
      End
      Begin VB.HScrollBar Spincomp 
         Height          =   255
         Index           =   1
         Left            =   3340
         TabIndex        =   299
         Top             =   720
         Width           =   360
      End
      Begin VB.HScrollBar Spincomp 
         Height          =   255
         Index           =   0
         Left            =   3340
         TabIndex        =   298
         Top             =   450
         Width           =   360
      End
      Begin VB.Label LabPtsAttrib 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   68
         Left            =   9600
         TabIndex        =   295
         Top             =   6390
         Width           =   300
      End
      Begin VB.Label LabCompetence 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   68
         Left            =   7335
         TabIndex        =   294
         Top             =   6390
         Width           =   2265
      End
      Begin VB.Label labnModif 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   68
         Left            =   9900
         TabIndex        =   293
         Top             =   6390
         Width           =   300
      End
      Begin VB.Label labnFinal 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   255
         Index           =   68
         Left            =   10200
         TabIndex        =   292
         Top             =   6390
         Width           =   300
      End
      Begin VB.Label LabPtsAttrib 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   67
         Left            =   9600
         TabIndex        =   291
         Top             =   6120
         Width           =   300
      End
      Begin VB.Label LabCompetence 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   67
         Left            =   7335
         TabIndex        =   290
         Top             =   6120
         Width           =   2265
      End
      Begin VB.Label labnModif 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   67
         Left            =   9900
         TabIndex        =   289
         Top             =   6120
         Width           =   300
      End
      Begin VB.Label labnFinal 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   255
         Index           =   67
         Left            =   10200
         TabIndex        =   288
         Top             =   6120
         Width           =   300
      End
      Begin VB.Label LabPtsAttrib 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   66
         Left            =   9600
         TabIndex        =   287
         Top             =   5850
         Width           =   300
      End
      Begin VB.Label LabCompetence 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   66
         Left            =   7335
         TabIndex        =   286
         Top             =   5850
         Width           =   2265
      End
      Begin VB.Label labnModif 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   66
         Left            =   9900
         TabIndex        =   285
         Top             =   5850
         Width           =   300
      End
      Begin VB.Label labnFinal 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   255
         Index           =   66
         Left            =   10200
         TabIndex        =   284
         Top             =   5850
         Width           =   300
      End
      Begin VB.Label LabPtsAttrib 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   65
         Left            =   9600
         TabIndex        =   283
         Top             =   5580
         Width           =   300
      End
      Begin VB.Label LabCompetence 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   65
         Left            =   7335
         TabIndex        =   282
         Top             =   5580
         Width           =   2265
      End
      Begin VB.Label labnModif 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   65
         Left            =   9900
         TabIndex        =   281
         Top             =   5580
         Width           =   300
      End
      Begin VB.Label labnFinal 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   255
         Index           =   65
         Left            =   10200
         TabIndex        =   280
         Top             =   5580
         Width           =   300
      End
      Begin VB.Label LabPtsAttrib 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   64
         Left            =   9600
         TabIndex        =   279
         Top             =   5310
         Width           =   300
      End
      Begin VB.Label LabCompetence 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   64
         Left            =   7335
         TabIndex        =   278
         Top             =   5310
         Width           =   2265
      End
      Begin VB.Label labnModif 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   64
         Left            =   9900
         TabIndex        =   277
         Top             =   5310
         Width           =   300
      End
      Begin VB.Label labnFinal 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   255
         Index           =   64
         Left            =   10200
         TabIndex        =   276
         Top             =   5310
         Width           =   300
      End
      Begin VB.Label LabPtsAttrib 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   63
         Left            =   9600
         TabIndex        =   275
         Top             =   5040
         Width           =   300
      End
      Begin VB.Label LabCompetence 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   63
         Left            =   7335
         TabIndex        =   274
         Top             =   5040
         Width           =   2265
      End
      Begin VB.Label labnModif 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   63
         Left            =   9900
         TabIndex        =   273
         Top             =   5040
         Width           =   300
      End
      Begin VB.Label labnFinal 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   255
         Index           =   63
         Left            =   10200
         TabIndex        =   272
         Top             =   5040
         Width           =   300
      End
      Begin VB.Label LabPtsAttrib 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   62
         Left            =   9600
         TabIndex        =   271
         Top             =   4770
         Width           =   300
      End
      Begin VB.Label LabCompetence 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   62
         Left            =   7335
         TabIndex        =   270
         Top             =   4770
         Width           =   2265
      End
      Begin VB.Label labnModif 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   62
         Left            =   9900
         TabIndex        =   269
         Top             =   4770
         Width           =   300
      End
      Begin VB.Label labnFinal 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   255
         Index           =   62
         Left            =   10200
         TabIndex        =   268
         Top             =   4770
         Width           =   300
      End
      Begin VB.Label LabPtsAttrib 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   61
         Left            =   9600
         TabIndex        =   267
         Top             =   4500
         Width           =   300
      End
      Begin VB.Label LabCompetence 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   61
         Left            =   7335
         TabIndex        =   266
         Top             =   4500
         Width           =   2265
      End
      Begin VB.Label labnModif 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   61
         Left            =   9900
         TabIndex        =   265
         Top             =   4500
         Width           =   300
      End
      Begin VB.Label labnFinal 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   255
         Index           =   61
         Left            =   10200
         TabIndex        =   264
         Top             =   4500
         Width           =   300
      End
      Begin VB.Label LabPtsAttrib 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   60
         Left            =   9600
         TabIndex        =   263
         Top             =   4230
         Width           =   300
      End
      Begin VB.Label LabCompetence 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   60
         Left            =   7335
         TabIndex        =   262
         Top             =   4230
         Width           =   2265
      End
      Begin VB.Label labnModif 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   60
         Left            =   9900
         TabIndex        =   261
         Top             =   4230
         Width           =   300
      End
      Begin VB.Label labnFinal 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   255
         Index           =   60
         Left            =   10200
         TabIndex        =   260
         Top             =   4230
         Width           =   300
      End
      Begin VB.Label LabPtsAttrib 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   59
         Left            =   9600
         TabIndex        =   259
         Top             =   3960
         Width           =   300
      End
      Begin VB.Label LabCompetence 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   59
         Left            =   7335
         TabIndex        =   258
         Top             =   3960
         Width           =   2265
      End
      Begin VB.Label labnModif 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   59
         Left            =   9900
         TabIndex        =   257
         Top             =   3960
         Width           =   300
      End
      Begin VB.Label labnFinal 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   255
         Index           =   59
         Left            =   10200
         TabIndex        =   256
         Top             =   3960
         Width           =   300
      End
      Begin VB.Label LabPtsAttrib 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   58
         Left            =   9600
         TabIndex        =   255
         Top             =   3690
         Width           =   300
      End
      Begin VB.Label LabCompetence 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   58
         Left            =   7335
         TabIndex        =   254
         Top             =   3690
         Width           =   2265
      End
      Begin VB.Label labnModif 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   58
         Left            =   9900
         TabIndex        =   253
         Top             =   3690
         Width           =   300
      End
      Begin VB.Label labnFinal 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   255
         Index           =   58
         Left            =   10200
         TabIndex        =   252
         Top             =   3690
         Width           =   300
      End
      Begin VB.Label LabPtsAttrib 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   57
         Left            =   9600
         TabIndex        =   251
         Top             =   3420
         Width           =   300
      End
      Begin VB.Label LabCompetence 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   57
         Left            =   7335
         TabIndex        =   250
         Top             =   3420
         Width           =   2265
      End
      Begin VB.Label labnModif 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   57
         Left            =   9900
         TabIndex        =   249
         Top             =   3420
         Width           =   300
      End
      Begin VB.Label labnFinal 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   255
         Index           =   57
         Left            =   10200
         TabIndex        =   248
         Top             =   3420
         Width           =   300
      End
      Begin VB.Label LabPtsAttrib 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   56
         Left            =   9600
         TabIndex        =   247
         Top             =   3150
         Width           =   300
      End
      Begin VB.Label LabCompetence 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   56
         Left            =   7335
         TabIndex        =   246
         Top             =   3150
         Width           =   2265
      End
      Begin VB.Label labnModif 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   56
         Left            =   9900
         TabIndex        =   245
         Top             =   3150
         Width           =   300
      End
      Begin VB.Label labnFinal 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   255
         Index           =   56
         Left            =   10200
         TabIndex        =   244
         Top             =   3150
         Width           =   300
      End
      Begin VB.Label LabPtsAttrib 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   55
         Left            =   9600
         TabIndex        =   243
         Top             =   2880
         Width           =   300
      End
      Begin VB.Label LabCompetence 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   55
         Left            =   7335
         TabIndex        =   242
         Top             =   2880
         Width           =   2265
      End
      Begin VB.Label labnModif 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   55
         Left            =   9900
         TabIndex        =   241
         Top             =   2880
         Width           =   300
      End
      Begin VB.Label labnFinal 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   255
         Index           =   55
         Left            =   10200
         TabIndex        =   240
         Top             =   2880
         Width           =   300
      End
      Begin VB.Label LabPtsAttrib 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   54
         Left            =   9600
         TabIndex        =   239
         Top             =   2610
         Width           =   300
      End
      Begin VB.Label LabCompetence 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   54
         Left            =   7335
         TabIndex        =   238
         Top             =   2610
         Width           =   2265
      End
      Begin VB.Label labnModif 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   54
         Left            =   9900
         TabIndex        =   237
         Top             =   2610
         Width           =   300
      End
      Begin VB.Label labnFinal 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   255
         Index           =   54
         Left            =   10200
         TabIndex        =   236
         Top             =   2610
         Width           =   300
      End
      Begin VB.Label LabPtsAttrib 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   53
         Left            =   9600
         TabIndex        =   235
         Top             =   2340
         Width           =   300
      End
      Begin VB.Label LabCompetence 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   53
         Left            =   7335
         TabIndex        =   234
         Top             =   2340
         Width           =   2265
      End
      Begin VB.Label labnModif 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   53
         Left            =   9900
         TabIndex        =   233
         Top             =   2340
         Width           =   300
      End
      Begin VB.Label labnFinal 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   255
         Index           =   53
         Left            =   10200
         TabIndex        =   232
         Top             =   2340
         Width           =   300
      End
      Begin VB.Label LabPtsAttrib 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   52
         Left            =   9600
         TabIndex        =   231
         Top             =   2070
         Width           =   300
      End
      Begin VB.Label LabCompetence 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   52
         Left            =   7335
         TabIndex        =   230
         Top             =   2070
         Width           =   2265
      End
      Begin VB.Label labnModif 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   52
         Left            =   9900
         TabIndex        =   229
         Top             =   2070
         Width           =   300
      End
      Begin VB.Label labnFinal 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   255
         Index           =   52
         Left            =   10200
         TabIndex        =   228
         Top             =   2070
         Width           =   300
      End
      Begin VB.Label LabPtsAttrib 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   51
         Left            =   9600
         TabIndex        =   227
         Top             =   1800
         Width           =   300
      End
      Begin VB.Label LabCompetence 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   51
         Left            =   7335
         TabIndex        =   226
         Top             =   1800
         Width           =   2265
      End
      Begin VB.Label labnModif 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   51
         Left            =   9900
         TabIndex        =   225
         Top             =   1800
         Width           =   300
      End
      Begin VB.Label labnFinal 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   255
         Index           =   51
         Left            =   10200
         TabIndex        =   224
         Top             =   1800
         Width           =   300
      End
      Begin VB.Label LabPtsAttrib 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   50
         Left            =   9600
         TabIndex        =   223
         Top             =   1530
         Width           =   300
      End
      Begin VB.Label LabCompetence 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   50
         Left            =   7335
         TabIndex        =   222
         Top             =   1530
         Width           =   2265
      End
      Begin VB.Label labnModif 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   50
         Left            =   9900
         TabIndex        =   221
         Top             =   1530
         Width           =   300
      End
      Begin VB.Label labnFinal 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   255
         Index           =   50
         Left            =   10200
         TabIndex        =   220
         Top             =   1530
         Width           =   300
      End
      Begin VB.Label LabPtsAttrib 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   49
         Left            =   9600
         TabIndex        =   219
         Top             =   1260
         Width           =   300
      End
      Begin VB.Label LabCompetence 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   49
         Left            =   7335
         TabIndex        =   218
         Top             =   1260
         Width           =   2265
      End
      Begin VB.Label labnModif 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   49
         Left            =   9900
         TabIndex        =   217
         Top             =   1260
         Width           =   300
      End
      Begin VB.Label labnFinal 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   255
         Index           =   49
         Left            =   10200
         TabIndex        =   216
         Top             =   1260
         Width           =   300
      End
      Begin VB.Label LabPtsAttrib 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   48
         Left            =   9600
         TabIndex        =   215
         Top             =   990
         Width           =   300
      End
      Begin VB.Label LabCompetence 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   48
         Left            =   7335
         TabIndex        =   214
         Top             =   990
         Width           =   2265
      End
      Begin VB.Label labnModif 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   48
         Left            =   9900
         TabIndex        =   213
         Top             =   990
         Width           =   300
      End
      Begin VB.Label labnFinal 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   255
         Index           =   48
         Left            =   10200
         TabIndex        =   212
         Top             =   990
         Width           =   300
      End
      Begin VB.Label LabPtsAttrib 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   47
         Left            =   9600
         TabIndex        =   211
         Top             =   720
         Width           =   300
      End
      Begin VB.Label LabCompetence 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   47
         Left            =   7335
         TabIndex        =   210
         Top             =   720
         Width           =   2265
      End
      Begin VB.Label labnModif 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   47
         Left            =   9900
         TabIndex        =   209
         Top             =   720
         Width           =   300
      End
      Begin VB.Label labnFinal 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   255
         Index           =   47
         Left            =   10200
         TabIndex        =   208
         Top             =   720
         Width           =   300
      End
      Begin VB.Label LabPtsAttrib 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   46
         Left            =   9600
         TabIndex        =   207
         Top             =   450
         Width           =   300
      End
      Begin VB.Label LabCompetence 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   46
         Left            =   7335
         TabIndex        =   206
         Top             =   450
         Width           =   2265
      End
      Begin VB.Label labnModif 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   46
         Left            =   9900
         TabIndex        =   205
         Top             =   450
         Width           =   300
      End
      Begin VB.Label labnFinal 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   255
         Index           =   46
         Left            =   10200
         TabIndex        =   204
         Top             =   450
         Width           =   300
      End
      Begin VB.Label Label11 
         Caption         =   "Compétence"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7440
         TabIndex        =   203
         Top             =   195
         Width           =   975
      End
      Begin VB.Label Label10 
         Caption         =   "Pts"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   9660
         TabIndex        =   202
         Top             =   225
         Width           =   330
      End
      Begin VB.Label Label1 
         Caption         =   "Modif"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   10110
         TabIndex        =   201
         Top             =   225
         Width           =   465
      End
      Begin VB.Label labnFinal 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   255
         Index           =   45
         Left            =   6600
         TabIndex        =   194
         Top             =   6390
         Width           =   300
      End
      Begin VB.Label labnFinal 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   255
         Index           =   44
         Left            =   6600
         TabIndex        =   193
         Top             =   6120
         Width           =   300
      End
      Begin VB.Label labnFinal 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   255
         Index           =   43
         Left            =   6600
         TabIndex        =   192
         Top             =   5850
         Width           =   300
      End
      Begin VB.Label labnFinal 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   255
         Index           =   42
         Left            =   6600
         TabIndex        =   191
         Top             =   5580
         Width           =   300
      End
      Begin VB.Label labnFinal 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   255
         Index           =   41
         Left            =   6600
         TabIndex        =   190
         Top             =   5310
         Width           =   300
      End
      Begin VB.Label labnFinal 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   255
         Index           =   40
         Left            =   6600
         TabIndex        =   189
         Top             =   5040
         Width           =   300
      End
      Begin VB.Label labnFinal 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   255
         Index           =   39
         Left            =   6600
         TabIndex        =   188
         Top             =   4770
         Width           =   300
      End
      Begin VB.Label labnFinal 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   255
         Index           =   38
         Left            =   6600
         TabIndex        =   187
         Top             =   4500
         Width           =   300
      End
      Begin VB.Label labnFinal 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   255
         Index           =   37
         Left            =   6600
         TabIndex        =   186
         Top             =   4230
         Width           =   300
      End
      Begin VB.Label labnFinal 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   255
         Index           =   36
         Left            =   6600
         TabIndex        =   185
         Top             =   3960
         Width           =   300
      End
      Begin VB.Label labnFinal 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   255
         Index           =   35
         Left            =   6600
         TabIndex        =   184
         Top             =   3690
         Width           =   300
      End
      Begin VB.Label labnFinal 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   255
         Index           =   34
         Left            =   6600
         TabIndex        =   183
         Top             =   3420
         Width           =   300
      End
      Begin VB.Label labnFinal 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   255
         Index           =   33
         Left            =   6600
         TabIndex        =   182
         Top             =   3150
         Width           =   300
      End
      Begin VB.Label labnFinal 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   255
         Index           =   32
         Left            =   6600
         TabIndex        =   181
         Top             =   2880
         Width           =   300
      End
      Begin VB.Label labnFinal 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   255
         Index           =   31
         Left            =   6600
         TabIndex        =   180
         Top             =   2610
         Width           =   300
      End
      Begin VB.Label labnFinal 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   255
         Index           =   30
         Left            =   6600
         TabIndex        =   179
         Top             =   2340
         Width           =   300
      End
      Begin VB.Label labnFinal 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   255
         Index           =   29
         Left            =   6600
         TabIndex        =   178
         Top             =   2070
         Width           =   300
      End
      Begin VB.Label labnFinal 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   255
         Index           =   28
         Left            =   6600
         TabIndex        =   177
         Top             =   1800
         Width           =   300
      End
      Begin VB.Label labnFinal 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   255
         Index           =   27
         Left            =   6600
         TabIndex        =   176
         Top             =   1530
         Width           =   300
      End
      Begin VB.Label labnFinal 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   255
         Index           =   26
         Left            =   6600
         TabIndex        =   175
         Top             =   1260
         Width           =   300
      End
      Begin VB.Label labnFinal 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   255
         Index           =   25
         Left            =   6600
         TabIndex        =   174
         Top             =   990
         Width           =   300
      End
      Begin VB.Label labnFinal 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   255
         Index           =   24
         Left            =   6600
         TabIndex        =   173
         Top             =   720
         Width           =   300
      End
      Begin VB.Label labnFinal 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   255
         Index           =   23
         Left            =   6600
         TabIndex        =   172
         Top             =   450
         Width           =   300
      End
      Begin VB.Label labnFinal 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   255
         Index           =   22
         Left            =   3000
         TabIndex        =   171
         Top             =   6390
         Width           =   300
      End
      Begin VB.Label labnFinal 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   255
         Index           =   21
         Left            =   3000
         TabIndex        =   170
         Top             =   6120
         Width           =   300
      End
      Begin VB.Label labnFinal 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   255
         Index           =   20
         Left            =   3000
         TabIndex        =   169
         Top             =   5850
         Width           =   300
      End
      Begin VB.Label labnFinal 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   255
         Index           =   19
         Left            =   3000
         TabIndex        =   168
         Top             =   5580
         Width           =   300
      End
      Begin VB.Label labnFinal 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   255
         Index           =   18
         Left            =   3000
         TabIndex        =   167
         Top             =   5310
         Width           =   300
      End
      Begin VB.Label labnFinal 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   255
         Index           =   17
         Left            =   3000
         TabIndex        =   166
         Top             =   5040
         Width           =   300
      End
      Begin VB.Label labnFinal 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   255
         Index           =   16
         Left            =   3000
         TabIndex        =   165
         Top             =   4770
         Width           =   300
      End
      Begin VB.Label labnFinal 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   255
         Index           =   15
         Left            =   3000
         TabIndex        =   164
         Top             =   4500
         Width           =   300
      End
      Begin VB.Label labnFinal 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   255
         Index           =   14
         Left            =   3000
         TabIndex        =   163
         Top             =   4230
         Width           =   300
      End
      Begin VB.Label labnFinal 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   255
         Index           =   13
         Left            =   3000
         TabIndex        =   162
         Top             =   3960
         Width           =   300
      End
      Begin VB.Label labnFinal 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   255
         Index           =   12
         Left            =   3000
         TabIndex        =   161
         Top             =   3690
         Width           =   300
      End
      Begin VB.Label labnFinal 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   255
         Index           =   11
         Left            =   3000
         TabIndex        =   160
         Top             =   3420
         Width           =   300
      End
      Begin VB.Label labnFinal 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   255
         Index           =   10
         Left            =   3000
         TabIndex        =   159
         Top             =   3150
         Width           =   300
      End
      Begin VB.Label labnFinal 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   255
         Index           =   9
         Left            =   3000
         TabIndex        =   158
         Top             =   2880
         Width           =   300
      End
      Begin VB.Label labnFinal 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   255
         Index           =   8
         Left            =   3000
         TabIndex        =   157
         Top             =   2610
         Width           =   300
      End
      Begin VB.Label labnFinal 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   255
         Index           =   7
         Left            =   3000
         TabIndex        =   156
         Top             =   2340
         Width           =   300
      End
      Begin VB.Label labnFinal 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   255
         Index           =   6
         Left            =   3000
         TabIndex        =   155
         Top             =   2070
         Width           =   300
      End
      Begin VB.Label labnFinal 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   255
         Index           =   5
         Left            =   3000
         TabIndex        =   154
         Top             =   1800
         Width           =   300
      End
      Begin VB.Label labnFinal 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   255
         Index           =   4
         Left            =   3000
         TabIndex        =   153
         Top             =   1530
         Width           =   300
      End
      Begin VB.Label labnFinal 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   255
         Index           =   3
         Left            =   3000
         TabIndex        =   152
         Top             =   1260
         Width           =   300
      End
      Begin VB.Label labnFinal 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   255
         Index           =   2
         Left            =   3000
         TabIndex        =   151
         Top             =   990
         Width           =   300
      End
      Begin VB.Label labnFinal 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   255
         Index           =   1
         Left            =   3000
         TabIndex        =   150
         Top             =   720
         Width           =   300
      End
      Begin VB.Label labnFinal 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   255
         Index           =   0
         Left            =   3000
         TabIndex        =   149
         Top             =   450
         Width           =   300
      End
      Begin VB.Label labnModif 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   45
         Left            =   6300
         TabIndex        =   148
         Top             =   6390
         Width           =   300
      End
      Begin VB.Label labnModif 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   44
         Left            =   6300
         TabIndex        =   147
         Top             =   6120
         Width           =   300
      End
      Begin VB.Label labnModif 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   43
         Left            =   6300
         TabIndex        =   146
         Top             =   5850
         Width           =   300
      End
      Begin VB.Label labnModif 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   42
         Left            =   6300
         TabIndex        =   145
         Top             =   5580
         Width           =   300
      End
      Begin VB.Label labnModif 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   41
         Left            =   6300
         TabIndex        =   144
         Top             =   5310
         Width           =   300
      End
      Begin VB.Label labnModif 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   40
         Left            =   6300
         TabIndex        =   143
         Top             =   5040
         Width           =   300
      End
      Begin VB.Label labnModif 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   39
         Left            =   6300
         TabIndex        =   142
         Top             =   4770
         Width           =   300
      End
      Begin VB.Label labnModif 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   38
         Left            =   6300
         TabIndex        =   141
         Top             =   4500
         Width           =   300
      End
      Begin VB.Label labnModif 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   37
         Left            =   6300
         TabIndex        =   140
         Top             =   4230
         Width           =   300
      End
      Begin VB.Label labnModif 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   36
         Left            =   6300
         TabIndex        =   139
         Top             =   3960
         Width           =   300
      End
      Begin VB.Label labnModif 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   35
         Left            =   6300
         TabIndex        =   138
         Top             =   3690
         Width           =   300
      End
      Begin VB.Label labnModif 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   34
         Left            =   6300
         TabIndex        =   137
         Top             =   3420
         Width           =   300
      End
      Begin VB.Label labnModif 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   33
         Left            =   6300
         TabIndex        =   136
         Top             =   3150
         Width           =   300
      End
      Begin VB.Label labnModif 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   32
         Left            =   6300
         TabIndex        =   135
         Top             =   2880
         Width           =   300
      End
      Begin VB.Label labnModif 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   31
         Left            =   6300
         TabIndex        =   134
         Top             =   2610
         Width           =   300
      End
      Begin VB.Label labnModif 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   30
         Left            =   6300
         TabIndex        =   133
         Top             =   2340
         Width           =   300
      End
      Begin VB.Label labnModif 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   29
         Left            =   6300
         TabIndex        =   132
         Top             =   2070
         Width           =   300
      End
      Begin VB.Label labnModif 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   28
         Left            =   6300
         TabIndex        =   131
         Top             =   1800
         Width           =   300
      End
      Begin VB.Label labnModif 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   27
         Left            =   6300
         TabIndex        =   130
         Top             =   1530
         Width           =   300
      End
      Begin VB.Label labnModif 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   26
         Left            =   6300
         TabIndex        =   129
         Top             =   1260
         Width           =   300
      End
      Begin VB.Label labnModif 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   25
         Left            =   6300
         TabIndex        =   128
         Top             =   990
         Width           =   300
      End
      Begin VB.Label labnModif 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   24
         Left            =   6300
         TabIndex        =   127
         Top             =   720
         Width           =   300
      End
      Begin VB.Label labnModif 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   23
         Left            =   6300
         TabIndex        =   126
         Top             =   450
         Width           =   300
      End
      Begin VB.Label labnModif 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   22
         Left            =   2700
         TabIndex        =   125
         Top             =   6390
         Width           =   300
      End
      Begin VB.Label labnModif 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   21
         Left            =   2700
         TabIndex        =   124
         Top             =   6120
         Width           =   300
      End
      Begin VB.Label labnModif 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   20
         Left            =   2700
         TabIndex        =   123
         Top             =   5850
         Width           =   300
      End
      Begin VB.Label labnModif 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   19
         Left            =   2700
         TabIndex        =   122
         Top             =   5580
         Width           =   300
      End
      Begin VB.Label labnModif 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   18
         Left            =   2700
         TabIndex        =   121
         Top             =   5310
         Width           =   300
      End
      Begin VB.Label labnModif 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   17
         Left            =   2700
         TabIndex        =   120
         Top             =   5040
         Width           =   300
      End
      Begin VB.Label labnModif 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   16
         Left            =   2700
         TabIndex        =   119
         Top             =   4770
         Width           =   300
      End
      Begin VB.Label labnModif 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   15
         Left            =   2700
         TabIndex        =   118
         Top             =   4500
         Width           =   300
      End
      Begin VB.Label labnModif 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   14
         Left            =   2700
         TabIndex        =   117
         Top             =   4230
         Width           =   300
      End
      Begin VB.Label labnModif 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   13
         Left            =   2700
         TabIndex        =   116
         Top             =   3960
         Width           =   300
      End
      Begin VB.Label labnModif 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   12
         Left            =   2700
         TabIndex        =   115
         Top             =   3690
         Width           =   300
      End
      Begin VB.Label labnModif 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   11
         Left            =   2700
         TabIndex        =   114
         Top             =   3420
         Width           =   300
      End
      Begin VB.Label labnModif 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   10
         Left            =   2700
         TabIndex        =   113
         Top             =   3150
         Width           =   300
      End
      Begin VB.Label labnModif 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   9
         Left            =   2700
         TabIndex        =   112
         Top             =   2880
         Width           =   300
      End
      Begin VB.Label labnModif 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   8
         Left            =   2700
         TabIndex        =   111
         Top             =   2610
         Width           =   300
      End
      Begin VB.Label labnModif 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   7
         Left            =   2700
         TabIndex        =   110
         Top             =   2340
         Width           =   300
      End
      Begin VB.Label labnModif 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   6
         Left            =   2700
         TabIndex        =   109
         Top             =   2070
         Width           =   300
      End
      Begin VB.Label labnModif 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   5
         Left            =   2700
         TabIndex        =   108
         Top             =   1800
         Width           =   300
      End
      Begin VB.Label labnModif 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   4
         Left            =   2700
         TabIndex        =   107
         Top             =   1530
         Width           =   300
      End
      Begin VB.Label labnModif 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   3
         Left            =   2700
         TabIndex        =   106
         Top             =   1260
         Width           =   300
      End
      Begin VB.Label labnModif 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   2
         Left            =   2700
         TabIndex        =   105
         Top             =   990
         Width           =   300
      End
      Begin VB.Label labnModif 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   1
         Left            =   2700
         TabIndex        =   104
         Top             =   720
         Width           =   300
      End
      Begin VB.Label labnModif 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   2700
         TabIndex        =   103
         Top             =   450
         Width           =   300
      End
      Begin VB.Label LabPtsAttrib 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   45
         Left            =   6000
         TabIndex        =   102
         Top             =   6390
         Width           =   300
      End
      Begin VB.Label LabCompetence 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   45
         Left            =   3735
         TabIndex        =   101
         Top             =   6390
         Width           =   2265
      End
      Begin VB.Label LabCompetence 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   44
         Left            =   3735
         TabIndex        =   100
         Top             =   6120
         Width           =   2265
      End
      Begin VB.Label LabPtsAttrib 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   44
         Left            =   6000
         TabIndex        =   99
         Top             =   6120
         Width           =   300
      End
      Begin VB.Label LabCompetence 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   43
         Left            =   3735
         TabIndex        =   98
         Top             =   5850
         Width           =   2265
      End
      Begin VB.Label LabPtsAttrib 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   43
         Left            =   6000
         TabIndex        =   97
         Top             =   5850
         Width           =   300
      End
      Begin VB.Label LabCompetence 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   42
         Left            =   3735
         TabIndex        =   96
         Top             =   5580
         Width           =   2265
      End
      Begin VB.Label LabPtsAttrib 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   42
         Left            =   6000
         TabIndex        =   95
         Top             =   5580
         Width           =   300
      End
      Begin VB.Label LabCompetence 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   41
         Left            =   3735
         TabIndex        =   94
         Top             =   5310
         Width           =   2265
      End
      Begin VB.Label LabPtsAttrib 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   41
         Left            =   6000
         TabIndex        =   93
         Top             =   5310
         Width           =   300
      End
      Begin VB.Label LabCompetence 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   40
         Left            =   3735
         TabIndex        =   92
         Top             =   5040
         Width           =   2265
      End
      Begin VB.Label LabPtsAttrib 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   40
         Left            =   6000
         TabIndex        =   91
         Top             =   5040
         Width           =   300
      End
      Begin VB.Label LabCompetence 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   39
         Left            =   3735
         TabIndex        =   90
         Top             =   4770
         Width           =   2265
      End
      Begin VB.Label LabPtsAttrib 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   39
         Left            =   6000
         TabIndex        =   89
         Top             =   4770
         Width           =   300
      End
      Begin VB.Label LabCompetence 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   38
         Left            =   3735
         TabIndex        =   88
         Top             =   4500
         Width           =   2265
      End
      Begin VB.Label LabPtsAttrib 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   38
         Left            =   6000
         TabIndex        =   87
         Top             =   4500
         Width           =   300
      End
      Begin VB.Label LabCompetence 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   37
         Left            =   3735
         TabIndex        =   86
         Top             =   4230
         Width           =   2265
      End
      Begin VB.Label LabPtsAttrib 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   37
         Left            =   6000
         TabIndex        =   85
         Top             =   4230
         Width           =   300
      End
      Begin VB.Label LabCompetence 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   36
         Left            =   3735
         TabIndex        =   84
         Top             =   3960
         Width           =   2265
      End
      Begin VB.Label LabPtsAttrib 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   36
         Left            =   6000
         TabIndex        =   83
         Top             =   3960
         Width           =   300
      End
      Begin VB.Label LabCompetence 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   35
         Left            =   3735
         TabIndex        =   82
         Top             =   3690
         Width           =   2265
      End
      Begin VB.Label LabPtsAttrib 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   35
         Left            =   6000
         TabIndex        =   81
         Top             =   3690
         Width           =   300
      End
      Begin VB.Label LabCompetence 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   34
         Left            =   3735
         TabIndex        =   80
         Top             =   3420
         Width           =   2265
      End
      Begin VB.Label LabPtsAttrib 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   34
         Left            =   6000
         TabIndex        =   79
         Top             =   3420
         Width           =   300
      End
      Begin VB.Label LabCompetence 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   33
         Left            =   3735
         TabIndex        =   78
         Top             =   3150
         Width           =   2265
      End
      Begin VB.Label LabPtsAttrib 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   33
         Left            =   6000
         TabIndex        =   77
         Top             =   3150
         Width           =   300
      End
      Begin VB.Label LabCompetence 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   32
         Left            =   3735
         TabIndex        =   76
         Top             =   2880
         Width           =   2265
      End
      Begin VB.Label LabPtsAttrib 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   32
         Left            =   6000
         TabIndex        =   75
         Top             =   2880
         Width           =   300
      End
      Begin VB.Label LabCompetence 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   31
         Left            =   3735
         TabIndex        =   74
         Top             =   2610
         Width           =   2265
      End
      Begin VB.Label LabPtsAttrib 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   31
         Left            =   6000
         TabIndex        =   73
         Top             =   2610
         Width           =   300
      End
      Begin VB.Label LabCompetence 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   30
         Left            =   3735
         TabIndex        =   72
         Top             =   2340
         Width           =   2265
      End
      Begin VB.Label LabPtsAttrib 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   30
         Left            =   6000
         TabIndex        =   71
         Top             =   2340
         Width           =   300
      End
      Begin VB.Label LabCompetence 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   29
         Left            =   3735
         TabIndex        =   70
         Top             =   2070
         Width           =   2265
      End
      Begin VB.Label LabPtsAttrib 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   29
         Left            =   6000
         TabIndex        =   69
         Top             =   2070
         Width           =   300
      End
      Begin VB.Label LabCompetence 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   28
         Left            =   3735
         TabIndex        =   68
         Top             =   1800
         Width           =   2265
      End
      Begin VB.Label LabPtsAttrib 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   28
         Left            =   6000
         TabIndex        =   67
         Top             =   1800
         Width           =   300
      End
      Begin VB.Label LabCompetence 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   27
         Left            =   3735
         TabIndex        =   66
         Top             =   1530
         Width           =   2265
      End
      Begin VB.Label LabPtsAttrib 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   27
         Left            =   6000
         TabIndex        =   65
         Top             =   1530
         Width           =   300
      End
      Begin VB.Label LabCompetence 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   26
         Left            =   3735
         TabIndex        =   64
         Top             =   1260
         Width           =   2265
      End
      Begin VB.Label LabPtsAttrib 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   26
         Left            =   6000
         TabIndex        =   63
         Top             =   1260
         Width           =   300
      End
      Begin VB.Label LabCompetence 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   25
         Left            =   3735
         TabIndex        =   62
         Top             =   990
         Width           =   2265
      End
      Begin VB.Label LabPtsAttrib 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   25
         Left            =   6000
         TabIndex        =   61
         Top             =   990
         Width           =   300
      End
      Begin VB.Label LabCompetence 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   24
         Left            =   3735
         TabIndex        =   60
         Top             =   720
         Width           =   2265
      End
      Begin VB.Label LabPtsAttrib 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   24
         Left            =   6000
         TabIndex        =   59
         Top             =   720
         Width           =   300
      End
      Begin VB.Label LabCompetence 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   23
         Left            =   3735
         TabIndex        =   58
         Top             =   450
         Width           =   2265
      End
      Begin VB.Label LabPtsAttrib 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   23
         Left            =   6000
         TabIndex        =   57
         Top             =   450
         Width           =   300
      End
      Begin VB.Label Label7 
         Caption         =   "Modif"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   6510
         TabIndex        =   56
         Top             =   225
         Width           =   465
      End
      Begin VB.Label Label6 
         Caption         =   "Pts"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   6060
         TabIndex        =   55
         Top             =   225
         Width           =   330
      End
      Begin VB.Label Label5 
         Caption         =   "Compétence"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3840
         TabIndex        =   54
         Top             =   195
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Compétence"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   51
         Top             =   195
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Pts"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2400
         TabIndex        =   50
         Top             =   225
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Modif"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2760
         TabIndex        =   49
         Top             =   225
         Width           =   525
      End
      Begin VB.Label LabPtsAttrib 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   2400
         TabIndex        =   48
         Top             =   450
         Width           =   300
      End
      Begin VB.Label LabCompetence 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   135
         TabIndex        =   47
         Top             =   450
         Width           =   2265
      End
      Begin VB.Label LabPtsAttrib 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   1
         Left            =   2400
         TabIndex        =   46
         Top             =   720
         Width           =   300
      End
      Begin VB.Label LabCompetence 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   135
         TabIndex        =   45
         Top             =   720
         Width           =   2265
      End
      Begin VB.Label LabPtsAttrib 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   2
         Left            =   2400
         TabIndex        =   44
         Top             =   990
         Width           =   300
      End
      Begin VB.Label LabCompetence 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   135
         TabIndex        =   43
         Top             =   990
         Width           =   2265
      End
      Begin VB.Label LabPtsAttrib 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   3
         Left            =   2400
         TabIndex        =   42
         Top             =   1260
         Width           =   300
      End
      Begin VB.Label LabCompetence 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   135
         TabIndex        =   41
         Top             =   1260
         Width           =   2265
      End
      Begin VB.Label LabPtsAttrib 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   4
         Left            =   2400
         TabIndex        =   40
         Top             =   1530
         Width           =   300
      End
      Begin VB.Label LabCompetence 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   135
         TabIndex        =   39
         Top             =   1530
         Width           =   2265
      End
      Begin VB.Label LabPtsAttrib 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   5
         Left            =   2400
         TabIndex        =   38
         Top             =   1800
         Width           =   300
      End
      Begin VB.Label LabCompetence 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   135
         TabIndex        =   37
         Top             =   1800
         Width           =   2265
      End
      Begin VB.Label LabPtsAttrib 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   6
         Left            =   2400
         TabIndex        =   36
         Top             =   2070
         Width           =   300
      End
      Begin VB.Label LabCompetence 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   135
         TabIndex        =   35
         Top             =   2070
         Width           =   2265
      End
      Begin VB.Label LabPtsAttrib 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   7
         Left            =   2400
         TabIndex        =   34
         Top             =   2340
         Width           =   300
      End
      Begin VB.Label LabCompetence 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   135
         TabIndex        =   33
         Top             =   2340
         Width           =   2265
      End
      Begin VB.Label LabPtsAttrib 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   8
         Left            =   2400
         TabIndex        =   32
         Top             =   2610
         Width           =   300
      End
      Begin VB.Label LabCompetence 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   8
         Left            =   135
         TabIndex        =   31
         Top             =   2610
         Width           =   2265
      End
      Begin VB.Label LabPtsAttrib 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   9
         Left            =   2400
         TabIndex        =   30
         Top             =   2880
         Width           =   300
      End
      Begin VB.Label LabCompetence 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   9
         Left            =   135
         TabIndex        =   29
         Top             =   2880
         Width           =   2265
      End
      Begin VB.Label LabPtsAttrib 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   10
         Left            =   2400
         TabIndex        =   28
         Top             =   3150
         Width           =   300
      End
      Begin VB.Label LabCompetence 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   10
         Left            =   135
         TabIndex        =   27
         Top             =   3150
         Width           =   2265
      End
      Begin VB.Label LabPtsAttrib 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   11
         Left            =   2400
         TabIndex        =   26
         Top             =   3420
         Width           =   300
      End
      Begin VB.Label LabCompetence 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   11
         Left            =   135
         TabIndex        =   25
         Top             =   3420
         Width           =   2265
      End
      Begin VB.Label LabPtsAttrib 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   12
         Left            =   2400
         TabIndex        =   24
         Top             =   3690
         Width           =   300
      End
      Begin VB.Label LabCompetence 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   12
         Left            =   135
         TabIndex        =   23
         Top             =   3690
         Width           =   2265
      End
      Begin VB.Label LabPtsAttrib 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   13
         Left            =   2400
         TabIndex        =   22
         Top             =   3960
         Width           =   300
      End
      Begin VB.Label LabCompetence 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   13
         Left            =   135
         TabIndex        =   21
         Top             =   3960
         Width           =   2265
      End
      Begin VB.Label LabPtsAttrib 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   14
         Left            =   2400
         TabIndex        =   20
         Top             =   4230
         Width           =   300
      End
      Begin VB.Label LabCompetence 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   14
         Left            =   135
         TabIndex        =   19
         Top             =   4230
         Width           =   2265
      End
      Begin VB.Label LabPtsAttrib 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   15
         Left            =   2400
         TabIndex        =   18
         Top             =   4500
         Width           =   300
      End
      Begin VB.Label LabCompetence 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   15
         Left            =   135
         TabIndex        =   17
         Top             =   4500
         Width           =   2265
      End
      Begin VB.Label LabPtsAttrib 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   16
         Left            =   2400
         TabIndex        =   16
         Top             =   4770
         Width           =   300
      End
      Begin VB.Label LabCompetence 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   16
         Left            =   135
         TabIndex        =   15
         Top             =   4770
         Width           =   2265
      End
      Begin VB.Label LabPtsAttrib 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   17
         Left            =   2400
         TabIndex        =   14
         Top             =   5040
         Width           =   300
      End
      Begin VB.Label LabCompetence 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   17
         Left            =   135
         TabIndex        =   13
         Top             =   5040
         Width           =   2265
      End
      Begin VB.Label LabPtsAttrib 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   18
         Left            =   2400
         TabIndex        =   12
         Top             =   5310
         Width           =   300
      End
      Begin VB.Label LabCompetence 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   18
         Left            =   135
         TabIndex        =   11
         Top             =   5310
         Width           =   2265
      End
      Begin VB.Label LabPtsAttrib 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   19
         Left            =   2400
         TabIndex        =   10
         Top             =   5580
         Width           =   300
      End
      Begin VB.Label LabCompetence 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   19
         Left            =   135
         TabIndex        =   9
         Top             =   5580
         Width           =   2265
      End
      Begin VB.Label LabPtsAttrib 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   20
         Left            =   2400
         TabIndex        =   8
         Top             =   5850
         Width           =   300
      End
      Begin VB.Label LabCompetence 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   20
         Left            =   135
         TabIndex        =   7
         Top             =   5850
         Width           =   2265
      End
      Begin VB.Label LabPtsAttrib 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   21
         Left            =   2400
         TabIndex        =   6
         Top             =   6120
         Width           =   300
      End
      Begin VB.Label LabCompetence 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   21
         Left            =   135
         TabIndex        =   5
         Top             =   6120
         Width           =   2265
      End
      Begin VB.Label LabCompetence 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   22
         Left            =   135
         TabIndex        =   4
         Top             =   6390
         Width           =   2265
      End
      Begin VB.Label LabPtsAttrib 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   22
         Left            =   2400
         TabIndex        =   3
         Top             =   6390
         Width           =   300
      End
   End
   Begin VB.CommandButton btnAnnuler 
      Cancel          =   -1  'True
      Caption         =   "&Annuler"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   15120
      TabIndex        =   1
      Top             =   6330
      Width           =   1005
   End
   Begin VB.CommandButton btnOk 
      Caption         =   "&Ok"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   14160
      TabIndex        =   0
      Top             =   6330
      Width           =   915
   End
   Begin VB.Label Label9 
      Caption         =   "Restant"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   12960
      TabIndex        =   297
      Top             =   6360
      Width           =   615
   End
   Begin VB.Label Label8 
      Caption         =   "Points à répartir"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   11160
      TabIndex        =   296
      Top             =   6360
      Width           =   1200
   End
   Begin VB.Label LabnRestant 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   13560
      TabIndex        =   53
      Top             =   6330
      Width           =   600
   End
   Begin VB.Label LabnArepartir 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   12360
      TabIndex        =   52
      Top             =   6330
      Width           =   600
   End
End
Attribute VB_Name = "frmDonsComp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Niveau As Integer
Public NiveauGlobal As Integer
Public bCreation As Boolean

Dim m_race As S_RACE
Dim m_class As S_CLASS
Dim m_archetype As S_ARCHETYPE
Dim m_charge As Boolean
Dim NomDon As String
Dim Genre As String
Dim Source As String
Dim Multiple As Boolean

Function ExistInDomaine(Domaine As String) As Boolean
Dim i As Integer

ExistInDomaine = False

For i = 0 To 3
  If frmGestPerso.cmbDomaine(i) = Domaine Then
    ExistInDomaine = True
    Exit Function
  End If
Next i

End Function
Private Function ExistInDon(strDon As String, strLib As String, vsflex As VSFlexGrid, Trie As Boolean) As Boolean
Dim i As Long
Dim min As Long, max As Long
Dim ret As Long, ecart As Long
Dim strDonNom As String

  With vsflex
    If .Rows < 2 Then
      Exit Function
    End If
  
    ExistInDon = False
  
    If Trie Then
      vsflex.Col = L8_COL_DON_NOM
      vsflex.Sort = flexSortStringAscending
    End If
    
    strDonNom = strDon & IIf(strLib <> "", " [" & strLib & "]", "")
      
    min = 1: max = .Rows
    
    Do
      i = Fix((min + max) / 2)
      ecart = max - min
      ret = StrComp(strDonNom, .Cell(flexcpText, i, L8_COL_DON_NOM), vbTextCompare)
      If (ret > 0) Then min = i
      If (ret < 0) Then max = i
    Loop While ((ret <> 0) And (ecart > 1))
    
    If ret = 0 Then
      ExistInDon = True
      Exit Function
    End If
  
  End With

End Function
Function ExistInDonNiveau(NomDon As String, vsflex As VSFlexGrid) As Boolean
Dim i As Long
Dim min As Long, max As Long
Dim ret As Integer, ecart As Long
Dim strDonNom As String

  ExistInDonNiveau = False
  
  With vsflex
    If .Rows < 2 Then
      Exit Function
    End If
  
    min = 1: max = .Rows
    Do
      i = Fix((min + max) / 2)
      ecart = max - min
      ret = StrComp(NomDon, .Cell(flexcpText, i, L7_COL_DON), vbTextCompare)
      If (ret > 0) Then min = i
      If (ret < 0) Then max = i
    Loop While ((ret <> 0) And (ecart > 1))
    If ret = 0 Then
      ExistInDonNiveau = True
      Exit Function
    End If
  End With
End Function
Sub vsDerNiveau()
Dim i As Integer, Col As Integer
Dim Coul As Long
  
  With vs1
    For i = 1 To .Rows - 1
      Coul = IIf(.Cell(flexcpText, i, L7_COL_DON_NIV) = frmGestPerso.vsHistoPerso.Row, vbGreen, vbWhite)
      For Col = 0 To .Cols - 1
        .Cell(flexcpBackColor, i, Col) = Coul
      Next Col
    Next i
  End With
End Sub
Function CanUp(Index As Integer) As Boolean
Dim nbPoints As Integer
Dim limMax As Integer
Dim i As Integer
Dim DeuxPoint As Integer

DeuxPoint = PointDouble()

  CanUp = False
  If LabCompetence(Index).Enabled = False Then
    Exit Function
  End If
  
  nbPoints = Val(labnModif(Index)) + Val(LabPtsAttrib(Index)) + 1
  
  If LabCompetence(Index).ForeColor = COUL_COMPCLASSE Or LabCompetence(Index).ForeColor = COUL_COMPOLDCLASSE Then
    limMax = frmGestPerso.vsHistoPerso.Row + 3
  Else
    limMax = Int((frmGestPerso.vsHistoPerso.Row + 3) / 2)
  End If
  
  If nbPoints > limMax Then
    Exit Function
  End If
  
  nbPoints = IIf(LabCompetence(Index).ForeColor = COUL_COMPCLASSE, 1, DeuxPoint)
  For i = 0 To labnModif.Count - 1
    If Val(labnModif(i)) > 0 Then
      nbPoints = nbPoints + IIf(LabCompetence(i).ForeColor = COUL_COMPCLASSE, Val(labnModif(i)), Val(labnModif(i)) * DeuxPoint)
    End If
  Next i
  
  If nbPoints > Val(LabnArepartir) Then
    Exit Function
  End If
  
  CanUp = True
End Function
Private Sub btnAnnuler_Click()

  If g_ModeDebug = vbUnchecked Then On Error GoTo btnAnnuler_Click_Error
  Unload Me

btnAnnuler_Click_Exit:
  Exit Sub

btnAnnuler_Click_Error:
  MsgBox "L'erreur suivante s'est produite : " & Err.Description & vbCrLf & vbCrLf & " numéro d'erreur (" & Err.Number & " )", vbCritical, g_strNomApplication
  Print #g_FileDebug, Format(Now, "dd/MM/yyyy hh:mm:ss") & " Erreur n° : (" & Err.Number & ") " & Err.Description & " Procedure : btnAnnuler_Click  Module : frmDonsComp.frm"
  Resume btnAnnuler_Click_Exit
End Sub

Private Sub btnOk_Click()
Dim i As Integer, n As Integer
Dim strCompetence As String
Dim strDon As String
Dim strLib As String
Dim DeuxPoint As Integer

  If g_ModeDebug = vbUnchecked Then On Error GoTo BtnOK_Click_Error
  DeuxPoint = PointDouble()
  strCompetence = ""
  For i = 0 To labnModif.Count - 1
    If Val(labnModif(i)) > 0 Then
      strCompetence = strCompetence & CStr(LabCompetence(i).Tag) & "-"
      strCompetence = strCompetence & CStr(IIf(LabCompetence(i).ForeColor = COUL_COMPCLASSE, labnModif(i), Val(labnModif(i)) * DeuxPoint)) & "-"
      strCompetence = strCompetence & CStr(labnModif(i)) & "|"
    End If
  Next i
  frmGestPerso.vsHistoPerso.Cell(flexcpText, frmGestPerso.vsHistoPerso.Row, L1_COL_COMPETENCE) = strCompetence
  
  n = 0
  With vs1
    For i = 1 To vs1.Rows - 1
      If Val(.Cell(flexcpText, i, L7_COL_DON_NIV)) = frmGestPerso.vsHistoPerso.Row Then
        strDon = .Cell(flexcpText, i, L7_COL_DON)
        strLib = .Cell(flexcpText, i, L7_COL_DON_LIB)
        If n < L1_NB_DONS Then
          frmGestPerso.vsHistoPerso.Cell(flexcpText, frmGestPerso.vsHistoPerso.Row, L1_COL_DON_1 + n) = strDon
          frmGestPerso.vsHistoPerso.Cell(flexcpText, frmGestPerso.vsHistoPerso.Row, L1_COL_LIB_1 + n) = strLib
        End If
        n = n + 1
      End If
    Next i
  End With
  Do While n < L1_NB_DONS
    frmGestPerso.vsHistoPerso.Cell(flexcpText, frmGestPerso.vsHistoPerso.Row, L1_COL_DON_1 + n) = ""
    frmGestPerso.vsHistoPerso.Cell(flexcpText, frmGestPerso.vsHistoPerso.Row, L1_COL_LIB_1 + n) = ""
    n = n + 1
  Loop
  frmGestPerso.Modifperso
  Unload Me

BtnOK_Click_Exit:
  Exit Sub

BtnOK_Click_Error:
  MsgBox "L'erreur suivante s'est produite : " & Err.Description & vbCrLf & vbCrLf & " numéro d'erreur (" & Err.Number & " )", vbCritical, g_strNomApplication
  Print #g_FileDebug, Format(Now, "dd/MM/yyyy hh:mm:ss") & " Erreur n° : (" & Err.Number & ") " & Err.Description & " Procedure : BtnOK_Click  Module : frmDonsComp.frm"
  Resume BtnOK_Click_Exit
End Sub

Private Sub btnAjouter_Click()
Dim i As Integer
Dim curlig1 As Integer
Dim curlig2 As Integer
Dim rst As Recordset

  If g_ModeDebug = vbUnchecked Then On Error GoTo btnAjouter_Click_Error
  
  If vs2.SelectedRows < 1 Then
    Exit Sub
  End If
  
  curlig2 = vs2.SelectedRow(0)
  Set rst = g_bdRoles.OpenRecordset("select * from dons  where nom='" & FormatStrSQL(vs2.Cell(flexcpText, curlig2, L8_COL_DON)) & "'", dbOpenDynaset)
  
  If rst!Multiple = False Then
    For i = 1 To vs1.Rows - 1
      If UCase(vs1.Cell(flexcpText, i, L7_COL_DON)) = UCase(vs2.Cell(flexcpText, curlig2, L8_COL_DON)) _
          And UCase(vs1.Cell(flexcpText, i, L7_COL_DON_LIB)) = UCase(vs2.Cell(flexcpText, curlig2, L8_COL_DON_LIB)) Then
        Exit Sub
      End If
    Next i
  End If
  vs1.AddItem vs2.Cell(flexcpText, curlig2, L8_COL_DON_NOM)
  curlig1 = vs1.Rows - 1
  vs1.Cell(flexcpText, curlig1, L7_COL_DON_NIV) = frmGestPerso.vsHistoPerso.Row
  vs1.Cell(flexcpText, curlig1, L7_COL_DON) = vs2.Cell(flexcpText, curlig2, L8_COL_DON)
  vs1.Cell(flexcpText, curlig1, L7_COL_DON_LIB) = vs2.Cell(flexcpText, curlig2, L8_COL_DON_LIB)
  vs1.Cell(flexcpText, curlig1, L7_COL_DON_GENRE) = vs2.Cell(flexcpText, curlig2, L8_COL_DON_GENRE)
  vs1.Cell(flexcpText, curlig1, L7_COL_DON_SOURCE) = vs2.Cell(flexcpText, curlig2, L8_COL_DON_SOURCE)
  If rst!Multiple = False Then
    vs2.RemoveItem curlig2
  End If
  vsDerNiveau
  
  rst.Close
btnAjouter_Click_Exit:
  Exit Sub

btnAjouter_Click_Error:
  MsgBox "L'erreur suivante s'est produite : " & Err.Description & vbCrLf & vbCrLf & " numéro d'erreur (" & Err.Number & " )", vbCritical, g_strNomApplication
  Print #g_FileDebug, Format(Now, "dd/MM/yyyy hh:mm:ss") & " Erreur n° : (" & Err.Number & ") " & Err.Description & " Procedure : btnAjouter_Click  Module : frmDonsComp.frm"
  Resume btnAjouter_Click_Exit
End Sub

Private Sub btnEnlever_Click()
Dim curlig1 As Integer
Dim curlig2 As Integer

  If g_ModeDebug = vbUnchecked Then On Error GoTo btnEnlever_Click_Error
  
  If vs1.SelectedRows < 1 Then
    Exit Sub
  End If
  
  curlig1 = vs1.SelectedRow(0)
  
  If vs1.Cell(flexcpText, curlig1, L7_COL_DON_NIV) <> frmGestPerso.vsHistoPerso.Row Then
    Exit Sub
  End If
  
  vs2.AddItem vs1.Cell(flexcpText, curlig1, L7_COL_DON_NOM)
  curlig2 = vs2.Rows - 1
  vs2.Cell(flexcpText, curlig2, L8_COL_DON) = vs1.Cell(flexcpText, curlig1, L7_COL_DON)
  vs2.Cell(flexcpText, curlig2, L8_COL_DON_LIB) = vs1.Cell(flexcpText, curlig1, L7_COL_DON_LIB)
  vs2.Cell(flexcpText, curlig2, L8_COL_DON_GENRE) = vs1.Cell(flexcpText, curlig1, L7_COL_DON_GENRE)
  vs2.Cell(flexcpText, curlig2, L8_COL_DON_SOURCE) = vs1.Cell(flexcpText, curlig1, L7_COL_DON_SOURCE)

  vs1.RemoveItem curlig1

btnEnlever_Click_Exit:
  Exit Sub

btnEnlever_Click_Error:
  MsgBox "L'erreur suivante s'est produite : " & Err.Description & vbCrLf & vbCrLf & " numéro d'erreur (" & Err.Number & " )", vbCritical, g_strNomApplication
  Print #g_FileDebug, Format(Now, "dd/MM/yyyy hh:mm:ss") & " Erreur n° : (" & Err.Number & ") " & Err.Description & " Procedure : btnEnlever_Click  Module : frmDonsComp.frm"
  Resume btnEnlever_Click_Exit
End Sub

Private Sub Form_Load()
Dim i As Integer

  If g_ModeDebug = vbUnchecked Then On Error GoTo Form_Load_Error
  
  With vs1
    .ColHidden(L7_COL_DON_GENRE) = True
    .ColHidden(L7_COL_DON_SOURCE) = True
    .ColHidden(L7_COL_DON_AIDE) = True
    .ColHidden(L7_COL_DON) = True
    .ColHidden(L7_COL_DON_LIB) = True
  End With
  With vs2
    .ColHidden(L8_COL_DON) = True
    .ColHidden(L8_COL_DON_LIB) = True
    .ColHidden(L8_COL_DON_AIDE) = True
  End With
  Timer1.Interval = 1000

Form_Load_Exit:
  Exit Sub

Form_Load_Error:
  MsgBox "L'erreur suivante s'est produite : " & Err.Description & vbCrLf & vbCrLf & " numéro d'erreur (" & Err.Number & " )", vbCritical, g_strNomApplication
  Print #g_FileDebug, Format(Now, "dd/MM/yyyy hh:mm:ss") & " Erreur n° : (" & Err.Number & ") " & Err.Description & " Procedure : Form_Load  Module : frmDonsComp.frm"
  Resume Form_Load_Exit
End Sub


Private Sub LabCompetence_Click(Index As Integer)
Dim n As Integer, i As Integer
Dim Competence As String

n = InStr(1, LabCompetence(Index).Caption, "(", vbTextCompare)
If n <> 0 Then
  If InStr(1, LabCompetence(Index).Caption, "Connaissance", vbTextCompare) <> 0 Or InStr(1, LabCompetence(Index).Caption, "Représentation", vbTextCompare) <> 0 Then
    n = InStr(n + 1, LabCompetence(Index).Caption, "(", vbTextCompare)
    If n <> 0 Then
      Competence = Trim(left(LabCompetence(Index).Caption, n - 1))
    Else
      Competence = Trim(LabCompetence(Index).Caption)
    End If
  Else
    Competence = Trim(left(LabCompetence(Index).Caption, n - 1))
  End If
Else
  Competence = Trim(LabCompetence(Index).Caption)
End If
With frmGestPerso
  For i = 0 To 2
    If Competence = .CmbArtisanat(i) Then
      Competence = "Artisanat 1"
    End If
  Next i
  For i = 0 To 3
    If Competence = .CmbProfession(i) Then
      Competence = "Profession 1"
    End If
  Next i
End With

frmDescriptifCompetence.m_competence = Competence
frmDescriptifCompetence.Show vbModal
End Sub

Private Sub labnModif_Change(Index As Integer)
Dim i As Integer
Dim nbPoints As Integer
Dim DeuxPoint As Integer

DeuxPoint = PointDouble

  If g_ModeDebug = vbUnchecked Then On Error GoTo labnModif_Change_Error
  
  If m_charge = True Then
    Exit Sub
  End If
  
  nbPoints = 0
  For i = 0 To labnModif.Count - 1
    If Val(labnModif(i)) > 0 Then
      nbPoints = nbPoints + IIf(LabCompetence(i).ForeColor = COUL_COMPCLASSE, Val(labnModif(i)), Val(labnModif(i)) * DeuxPoint)
    End If
    If IsNumeric(LabPtsAttrib(i)) Or IsNumeric(labnModif(i)) Then
      labnFinal(i) = Val(labnModif(i)) + Val(LabPtsAttrib(i))
    End If
  Next i
  LabnRestant = Val(LabnArepartir) - nbPoints

labnModif_Change_Exit:
  Exit Sub

labnModif_Change_Error:
  MsgBox "L'erreur suivante s'est produite : " & Err.Description & vbCrLf & vbCrLf & " numéro d'erreur (" & Err.Number & " )", vbCritical, g_strNomApplication
  Print #g_FileDebug, Format(Now, "dd/MM/yyyy hh:mm:ss") & " Erreur n° : (" & Err.Number & ") " & Err.Description & " Procedure : labnModif_Change  Module : frmDonsComp.frm"
  Resume labnModif_Change_Exit
End Sub

Private Sub SpinComp_SpinDown(Index As Integer)

  If g_ModeDebug = vbUnchecked Then On Error GoTo SpinComp_SpinDown_Error
  
  If Val(labnModif(Index)) > 0 Then
    labnModif(Index) = Val(labnModif(Index)) - 1
  End If

SpinComp_SpinDown_Exit:
  Exit Sub

SpinComp_SpinDown_Error:
  MsgBox "L'erreur suivante s'est produite : " & Err.Description & vbCrLf & vbCrLf & " numéro d'erreur (" & Err.Number & " )", vbCritical, g_strNomApplication
  Print #g_FileDebug, Format(Now, "dd/MM/yyyy hh:mm:ss") & " Erreur n° : (" & Err.Number & ") " & Err.Description & " Procedure : SpinComp_SpinDown  Module : frmDonsComp.frm"
  Resume SpinComp_SpinDown_Exit
End Sub
Private Sub SpinComp_spinup(Index As Integer)

  If g_ModeDebug = vbUnchecked Then On Error GoTo SpinComp_SpinUp_Error
  
  If CanUp(Index) Then
    labnModif(Index) = Val(labnModif(Index)) + 1
  End If

SpinComp_SpinUp_Exit:
  Exit Sub

SpinComp_SpinUp_Error:
  MsgBox "L'erreur suivante s'est produite : " & Err.Description & vbCrLf & vbCrLf & " numéro d'erreur (" & Err.Number & " )", vbCritical, g_strNomApplication
  Print #g_FileDebug, Format(Now, "dd/MM/yyyy hh:mm:ss") & " Erreur n° : (" & Err.Number & ") " & Err.Description & " Procedure : SpinComp_SpinUp  Module : frmDonsComp.frm"
  Resume SpinComp_SpinUp_Exit
End Sub

Private Sub SpinComp_Change(Index As Integer)

If Spincomp(Index).Value <> Val(Spincomp(Index).Tag) Then
  If Spincomp(Index).Value < Val(Spincomp(Index).Tag) Then
    SpinComp_SpinDown Index
    Spincomp(Index).Tag = CStr(Val(Spincomp(Index).Tag) - 1)
  Else
    SpinComp_spinup Index
    Spincomp(Index).Tag = CStr(Val(Spincomp(Index).Tag) + 1)
  End If
End If
   
End Sub

Private Sub Timer1_Timer()
Dim rst As Recordset
Dim rst2 As Recordset, rstEnnemi As Recordset, rstAberration As Recordset, rstGrandArcane As Recordset, rstTerrain As Recordset, rstEnergie As Recordset, rstPersonnage As Recordset, rstAlignement As Recordset, rstTaille As Recordset, CompetenceRace As Recordset
Dim rstcompetence As Recordset
Dim RstProgression As Recordset
Dim i As Integer
Dim n As Integer, j As Integer, m As Integer, l As Integer, BonusComp As Integer, BonusCompNiveau As Integer
Dim nbCompetence As Integer
Dim bexist As Boolean
Dim benable As Boolean, bmonstre As Boolean
Dim nColor As Long
Dim nRepartir As Integer
Dim tabComp() As S_TABVAL
Dim tabDomaine() As String
Dim nbRepartis As Integer
Dim strDon As String
Dim strLibelle As String
Dim nRep As Integer
Dim tblObjet As Recordset
Dim maxInt As Integer
Dim curlig1 As Integer, NiveauDon As Integer
Dim curlig2 As Integer
Dim tabSortDrider() As String
Dim tabClassLanceurSort() As String
Dim nClassLanceurSort As Integer
Dim nbDomaines As Integer
Dim empl As Integer
Dim PointComp As Integer, NiveauPretre As Integer
Dim NbPointComp As Integer
Dim NomCompetence As String

  If g_ModeDebug = vbUnchecked Then On Error GoTo Timer1_Timer_Error
  Timer1.Interval = 0
  
  With frmGestPerso

    
    If .vsHistoPerso.Row < 1 Then
      Unload Me
      Exit Sub
    End If
    
    
    If ChercheRace(m_race, .CmbRace) = False Then
      MsgBox "Saisie obligatoire de la race du personnage", vbExclamation, g_strNomApplication
      Unload Me
      Exit Sub
    End If
    ChercheArchetype m_archetype, .Cmbarchetype
    m_race.Int = m_race.Int + m_archetype.Int
    
    vs1.Redraw = False
    
    Set rst = g_bdRoles.OpenRecordset("dons", dbOpenTable)
    rst.Index = "nom"
    For i = 1 To .vsHistoPerso.Row
      ''renseigne les dons
      For n = 0 To L1_NB_DONS - 1
        If .vsHistoPerso.Cell(flexcpText, i, L1_COL_DON_1 + n) <> "" Then
          strDon = .vsHistoPerso.Cell(flexcpText, i, L1_COL_DON_1 + n)
          If .vsHistoPerso.Cell(flexcpText, i, L1_COL_LIB_1 + n) <> "" Then
            strDon = strDon & " [" & .vsHistoPerso.Cell(flexcpText, i, L1_COL_LIB_1 + n) & "]"
          End If
          vs1.AddItem strDon
          curlig1 = vs1.Rows - 1
          vs1.Cell(flexcpText, curlig1, L7_COL_DON_NIV) = i ''retient le niveau
          vs1.Cell(flexcpText, curlig1, L7_COL_DON) = .vsHistoPerso.Cell(flexcpText, i, L1_COL_DON_1 + n) '' le don
          vs1.Cell(flexcpText, curlig1, L7_COL_DON_LIB) = .vsHistoPerso.Cell(flexcpText, i, L1_COL_LIB_1 + n) ''à quoi le don est associé
          rst.Seek "=", vs1.Cell(flexcpText, curlig1, L7_COL_DON)
          If Not rst.NoMatch Then
            vs1.Cell(flexcpText, curlig1, L7_COL_DON_GENRE) = rst!Genre
            vs1.Cell(flexcpText, curlig1, L7_COL_DON_SOURCE) = rst!Source
          End If
        End If
      Next n
    Next i
    rst.Close

    vs1.Col = L7_COL_DON_NOM
    vs1.Sort = flexSortStringAscending

    nRep = vbYes 'inutille
    
    ''intellect point de competence
    maxInt = 0
    nRepartir = 0
    
    For i = 1 To vs1.Rows - 1
      If vs1.Cell(flexcpText, i, L7_COL_DON) = COMPETENCE_SUPPLEMENTAIRE Then
        BonusComp = BonusComp + 1
      End If
    Next i
    
''On tourne sur les niveaux du personnage du 1 au niveau max
    For i = 1 To .vsHistoPerso.Row
      BonusCompNiveau = 0
      If .vsHistoPerso.Cell(flexcpText, i, L1_COL_CLASSE) = "" Then
        MsgBox "Saisi obligatoire de la classe de ce personnage", vbExclamation, g_strNomApplication
        Unload Me
        Exit Sub
      End If
      
      If Val(.vsHistoPerso.Cell(flexcpText, i, L1_COL_NIVEAU)) < 1 Then
        MsgBox "Saisie d'un niveau obligatoire", vbExclamation, g_strNomApplication
        Unload Me
        Exit Sub
      End If
      
      If ChercheClasse(m_class, _
          .vsHistoPerso.Cell(flexcpText, i, L1_COL_CLASSE), _
          .vsHistoPerso.Cell(flexcpText, i, L1_COL_NIVEAU)) = False Then
        MsgBox "Saisi obligatoire d'une classe existante", vbExclamation, g_strNomApplication
        Unload Me
        Exit Sub
      End If
        
      For j = 1 To vs1.Rows - 1
        NiveauDon = Val(vs1.Cell(flexcpText, j, L7_COL_DON_NIV))
        If i >= NiveauDon Then
          Select Case vs1.Cell(flexcpText, j, L7_COL_DON)
            Case OUVERT_D_ESPRIT
              If i = NiveauDon Then
                BonusCompNiveau = BonusCompNiveau + 5
              End If
            Case BAISER_DE_LA_NYMPHE
                BonusCompNiveau = BonusCompNiveau + 1
            Case Else
              If i = NiveauDon Then
                If EstSkillTriks(vs1.Cell(flexcpText, j, L7_COL_DON_GENRE)) Then
                  BonusCompNiveau = BonusCompNiveau - 2
                End If
              End If
          End Select
        End If
      Next j
      
      maxInt = IIf(m_class.BonusIntelligence > maxInt, m_class.BonusIntelligence, maxInt) ''changement de carac definitif
      
      m_race.Int = m_race.Int + Val(.vsHistoPerso.Cell(flexcpText, i, L1_COL_INT))
      NbPointComp = BonusComp + Modificateur(maxInt + m_race.Int + Val(.vsHistoPerso.Cell(flexcpText, i, L1_COL_MODIF_INT))) + m_race.Competence
      Select Case .Cmbarchetype
        Case SQUELETTE, ZOMBI
          PointComp = 0
        Case DEMI_CELESTE, DEMI_FIELON
          If EstClasseMonstrueuse(.vsHistoPerso.Cell(flexcpText, i, L1_COL_CLASSE)) Then
            PointComp = maxi_(1, NbPointComp + 8)
          Else
            PointComp = maxi_(1, NbPointComp + m_class.CompetenceNiveauSupp)
          End If
        Case DEMI_DRAGON, DEMI_FEE
          If EstClasseMonstrueuse(.vsHistoPerso.Cell(flexcpText, i, L1_COL_CLASSE)) Then
            PointComp = maxi_(1, NbPointComp + maxi_(6, m_class.CompetenceNiveauSupp))
          Else
            PointComp = maxi_(1, NbPointComp + m_class.CompetenceNiveauSupp)
          End If
        Case Else
          PointComp = maxi_(1, NbPointComp + m_class.CompetenceNiveauSupp)
      End Select
      If PasDePointDeCompetence(.CmbRace, .Cmbarchetype, .vsHistoPerso.Cell(flexcpText, 1, L1_COL_CLASSE)) Or m_class.CompetenceNiveauSupp = 0 Then
        PointComp = 0
      End If
      If i = 1 Then
        nRepartir = nRepartir + PointComp * 4
      Else
        nRepartir = nRepartir + PointComp
      End If
        nRepartir = nRepartir + BonusCompNiveau
    Next i
        
    If m_race.Int = 0 Then
      MsgBox "Saisie obligatoire de l'intelligence pour ce personnage", vbExclamation, g_strNomApplication
      Unload Me
      Exit Sub
    End If
     
    bexist = False
    Set rst = g_bdRoles.OpenRecordset("Competence", dbOpenTable)
    For i = 0 To rst.Fields.Count - 1
      If UCase(m_class.Classe) = UCase(rst.Fields(i).name) Then
        bexist = True
        Exit For
      End If
    Next i
    rst.Close
    
    If bexist = False Then
      MsgBox "Table competence mal parametrée", vbInformation, g_strNomApplication
      Unload Me
      Exit Sub
    End If
    
    i = 0
    bmonstre = False

    Set rst = g_bdRoles.OpenRecordset("select * from Competence order by nom", dbOpenDynaset)
    
    '' Compétences de race
    If EstClasseMonstrueuse(m_class.Classe) Then
      bmonstre = True
      Set CompetenceRace = g_bdRoles.OpenRecordset("racecompetence", dbOpenTable)
      CompetenceRace.Index = "index"
    End If
    
    
    NiveauPretre = ChercheNiveauMinClasse("Prêtre", "Prêtre cloîtré")
    
    Do While Not rst.EOF
      NomCompetence = rst.Fields("NOM")
      benable = True
      nColor = vbBlack
      If IsNumeric(Right(NomCompetence, 1)) Then   ''Si il y a un nombre à la fin du mot on change l'intituler de la compétence
        Select Case NomCompetence
          Case "Artisanat 1"
            LabCompetence(i).Caption = .CmbArtisanat(0)
          Case "Artisanat 2"
            LabCompetence(i).Caption = .CmbArtisanat(1)
          Case "Artisanat 3"
            LabCompetence(i).Caption = .CmbArtisanat(2)
          Case "Profession 1"
            LabCompetence(i).Caption = .CmbProfession(0)
          Case "Profession 2"
            LabCompetence(i).Caption = .CmbProfession(1)
          Case "Profession 3"
            LabCompetence(i).Caption = .CmbProfession(2)
          Case "Profession 4"
            LabCompetence(i).Caption = .CmbProfession(3)
        End Select
      Else
          LabCompetence(i).Caption = NomCompetence
      End If
      Select Case rst.Fields(m_class.Classe)
        Case "-1"
          benable = False
        Case "0"
          If frmGestPerso.CmbDragonTotem <> "" Then
            Select Case frmGestPerso.CmbDragonTotem
              Case "Blanc"
                If NomCompetence = "Déplacement silencieux" Or NomCompetence = "Natation" Or NomCompetence = "Discrétion" Then
                  nColor = COUL_COMPCLASSE
                  LabCompetence(i).Caption = LabCompetence(i).Caption & " (Totem)"
                End If
              Case "Bleu"
                If NomCompetence = "Bluff" Or NomCompetence = "Art de la magie" Or NomCompetence = "Discrétion" Then
                  nColor = COUL_COMPCLASSE
                  LabCompetence(i).Caption = LabCompetence(i).Caption & " (Totem)"
                End If
              Case "Noir"
                If NomCompetence = "Déplacement silencieux" Or NomCompetence = "Natation" Or NomCompetence = "Discrétion" Then
                  nColor = COUL_COMPCLASSE
                  LabCompetence(i).Caption = LabCompetence(i).Caption & " (Totem)"
                End If
              Case "Rouge"
                If NomCompetence = "Estimation" Or NomCompetence = "Bluff" Or NomCompetence = "Saut" Then
                  nColor = COUL_COMPCLASSE
                  LabCompetence(i).Caption = LabCompetence(i).Caption & " (Totem)"
                End If
              Case "Vert"
                If NomCompetence = "Déplacement silencieux" Or NomCompetence = "Bluff" Or NomCompetence = "Discrétion" Then
                  nColor = COUL_COMPCLASSE
                  LabCompetence(i).Caption = LabCompetence(i).Caption & " (Totem)"
                End If
              Case "Airain"
                If NomCompetence = "Bluff" Or NomCompetence = "Renseignements" Or NomCompetence = "Survie" Then
                  nColor = COUL_COMPCLASSE
                  LabCompetence(i).Caption = LabCompetence(i).Caption & " (Totem)"
                End If
              Case "Argent"
                If NomCompetence = "Bluff" Or NomCompetence = "Déguisement" Or NomCompetence = "Saut" Then
                  nColor = COUL_COMPCLASSE
                  LabCompetence(i).Caption = LabCompetence(i).Caption & " (Totem)"
                End If
              Case "Bronze"
                If NomCompetence = "Déguisement" Or NomCompetence = "Natation" Or NomCompetence = "Survie" Then
                  nColor = COUL_COMPCLASSE
                  LabCompetence(i).Caption = LabCompetence(i).Caption & " (Totem)"
                End If
              Case "Cuivre"
                If NomCompetence = "Bluff" Or NomCompetence = "Discrétion" Or NomCompetence = "Saut" Then
                  nColor = COUL_COMPCLASSE
                  LabCompetence(i).Caption = LabCompetence(i).Caption & " (Totem)"
                End If
              Case "Or"
                If NomCompetence = "Déguisement" Or NomCompetence = "Premiers secours" Or NomCompetence = "Natation" Then
                  nColor = COUL_COMPCLASSE
                  LabCompetence(i).Caption = LabCompetence(i).Caption & " (Totem)"
                End If
            End Select
          End If
          If bmonstre Then
            CompetenceRace.Seek "=", m_race.RACE, NomCompetence
            If Not CompetenceRace.NoMatch Then
              nColor = COUL_COMPCLASSE
              LabCompetence(i).Caption = LabCompetence(i).Caption & " (race)"
              If Not IsNull(CompetenceRace!Rang) Then
                If bCreation Then
                  labnModif(i) = CompetenceRace!Rang
                End If
              End If
            End If
          End If
          
          If NiveauPretre > 0 Then
            If InStr(1, NomCompetence, "Connaissance", vbTextCompare) > 0 Then
              If ExistInDomaine("Connaissance") Then
                CouleurDomaine m_class.Classe, nColor, i, NiveauPretre
              End If
            End If
            If NomCompetence = "Bluff" Or NomCompetence = "Déguisement" Or NomCompetence = "Discrétion" Then
              If ExistInDomaine("Duperie") Then
                CouleurDomaine m_class.Classe, nColor, i, NiveauPretre
              End If
            End If
            If LabCompetence(i).Caption = "Connaissances (nature)" Then
              If ExistInDomaine("Faune") Or ExistInDomaine("Flore") Then
                CouleurDomaine m_class.Classe, nColor, i, NiveauPretre
              End If
            End If
            If NomCompetence = "Survie" Then
              If ExistInDomaine("Voyage") Then
                CouleurDomaine m_class.Classe, nColor, i, NiveauPretre
              End If
            End If
          End If
          
          If NomCompetence = "Fouille" And m_class.Classe = "Magicien" Then
            If ExistInDonNiveau(MAGIE_GENERALISTE, vs1) Then
              If Niveau = 1 Then
                nColor = COUL_COMPCLASSE
                LabCompetence(i).Caption = LabCompetence(i).Caption & " (Elfe)"
              Else
                nColor = COUL_COMPOLDCLASSE
              End If
            End If
          End If
          
          If NomCompetence = "Connaissances (explo souterre)" And m_class.Classe = "Guerrier" Then
            If ExistInDonNiveau(PREDILECTION_POUR_LES_HACHES, vs1) Then
              If Niveau = 1 Then
                nColor = COUL_COMPCLASSE
                LabCompetence(i).Caption = LabCompetence(i).Caption & " (Nain)"
              Else
                nColor = COUL_COMPOLDCLASSE
              End If
            End If
          End If
          
          If m_class.Classe = "Magicien" Then
            If (NomCompetence = "Bluff" Or NomCompetence = "Diplomatie" Or NomCompetence = "Intimidation" Or NomCompetence = "Psychologie" Or NomCompetence = "Renseignements") Then
              If ExistInDonNiveau(COMPETENCE_SOCIALE, vs1) Then
                nColor = COUL_COMPCLASSE
                LabCompetence(i).Caption = LabCompetence(i).Caption & " (Enchanteur)"
              End If
            End If
          End If
          
          If InStr(1, .Cmbarchetype, "garou", vbTextCompare) Then
            If NomCompetence = "Contrôle de forme" Then
              nColor = COUL_COMPCLASSE
              LabCompetence(i).Caption = LabCompetence(i).Caption & " (garou)"
            End If
          End If
          
          If nColor <> COUL_COMPCLASSE Then
            Set RstProgression = g_bdPerso.OpenRecordset("select Classe,Niv from PersonnageProgression where nom='" & FormatStrSQL(.fldstrnom) & "'", dbOpenDynaset)
            Do While Not RstProgression.EOF
              If rst.Fields(RstProgression!Classe) = "1" Then
                If RstProgression!niv <= NiveauGlobal Then
                  nColor = COUL_COMPOLDCLASSE
                  Exit Do
                Else
                  nColor = COUL_COMPFUTURCLASSE
                End If
              End If
              RstProgression.MoveNext
            Loop
            RstProgression.Close
            
            Set CompetenceRace = g_bdRoles.OpenRecordset("racecompetence", dbOpenTable)
            CompetenceRace.Index = "index"
            CompetenceRace.Seek "=", m_race.RACE, NomCompetence
            If Not CompetenceRace.NoMatch Then
              If EstClasseMonstrueuse(.vsHistoPerso.Cell(flexcpText, 1, L1_COL_CLASSE)) Then
                nColor = COUL_COMPOLDCLASSE
              End If
            End If
          End If
        Case "1"
          nColor = COUL_COMPCLASSE
          LabCompetence(i).Caption = LabCompetence(i).Caption & " (classe)"
      End Select
      
      LabCompetence(i).Enabled = benable
      LabPtsAttrib(i).Enabled = benable
      LabCompetence(i).ForeColor = nColor
      labnModif(i).Enabled = benable
      Spincomp(i).Enabled = benable
      LabCompetence(i).Tag = rst![n°]
      rst.MoveNext
      i = i + 1
    Loop
    rst.Close
    
    If bmonstre Or True Then
      CompetenceRace.Close
    End If
    
    ''renseigne les points
    m_charge = True

    vs2.Redraw = False

    nbRepartis = 0
    For i = 1 To .vsHistoPerso.Row
      ''renseigne les competences
      nbCompetence = EclateCompetence(.vsHistoPerso.Cell(flexcpText, i, L1_COL_COMPETENCE), tabComp)
      For n = 0 To nbCompetence - 1
        For j = 0 To labnModif.Count - 1
          If Val(tabComp(n).indic(0)) = Val(LabCompetence(j).Tag) Then
            If i < .vsHistoPerso.Row Then
              nbRepartis = nbRepartis + tabComp(n).indic(1)
              LabPtsAttrib(j) = Val(LabPtsAttrib(j)) + tabComp(n).indic(2)
            Else
              labnModif(j) = tabComp(n).indic(2)
            End If
            Exit For
          End If
        Next j
      Next n
    Next i
 
    nbDomaines = ExtraitDomaineSort(tabDomaine, False, "", "", "", "")
    
    m_charge = False
    
    LabnArepartir = nRepartir - nbRepartis
    labnModif_Change (0)
    
    For l = 0 To Spincomp.Count - 1
      If Val(labnFinal(l)) <> 0 Then
        Spincomp(l).Tag = CStr(labnFinal(l))
        Spincomp(l).Value = Val(labnFinal(l))
      End If
    Next l
    
    ''Les compétences sont finies ainsi que la liste des dons du personnage
    
    ''remplie la liste des dons dispos
    
    ''alimente sous un tableau les classes de lanceur de sort
    nClassLanceurSort = 0
    Dim CurClasse As String
    
    ReDim tabClassLanceurSort(nClassLanceurSort)
    For i = 1 To .vsHistoPerso.Row
      CurClasse = .vsHistoPerso.Cell(flexcpText, i, L1_COL_CLASSE)
      If EstLanceurSort(CurClasse) Then
        bexist = False
        For m = 0 To nClassLanceurSort - 1
          If UCase(tabClassLanceurSort(m)) = UCase(CurClasse) Then
            bexist = True
            Exit For
          End If
        Next m
        If bexist = False Then
          ReDim Preserve tabClassLanceurSort(nClassLanceurSort + 1)
          tabClassLanceurSort(nClassLanceurSort) = CurClasse
          nClassLanceurSort = nClassLanceurSort + 1
        End If
      End If
    Next i
    
    Set tblObjet = g_bdPerso.OpenRecordset("objets", dbOpenTable)
    tblObjet.Index = "primaryKey"
    Set rst = g_bdRoles.OpenRecordset("select Nom, Genre, Source,Multiple,liearme from dons where prive=FALSE and genre<>'Objet magique' order by nom", dbOpenDynaset)
    
    vs1.Col = L8_COL_DON_NOM
    vs1.Sort = flexSortStringAscending
    
    Do While Not rst.EOF
      NomDon = rst!Nom
      Genre = rst!Genre
      Source = IIf(IsNull(rst!Source), "", rst!Source)
      Multiple = rst!Multiple
      
      Select Case NomDon
        Case DISCIPLE_PROFANE
          For i = 0 To nbDomaines - 1
           AjouteDonListe tabDomaine(i)
          Next i
        Case BONUS_ENNEMI_JURE, ENNEMI_JURE
          Set rstEnnemi = g_bdRoles.OpenRecordset("select Nom from EnnemiJure order by nom", dbOpenDynaset)
          Do While Not rstEnnemi.EOF
            AjouteDonListe rstEnnemi!Nom
            rstEnnemi.MoveNext
          Loop
          rstEnnemi.Close
        Case AJOUT_D_ENERGIE_DESTRUCTIVE, SUBSTITUTION_D_ENERGIE_DESTRUCTIVE
          Set rstEnergie = g_bdRoles.OpenRecordset("select Nom from Energie order by nom", dbOpenDynaset)
          Do While Not rstEnergie.EOF
            AjouteDonListe rstEnergie!Nom
            rstEnergie.MoveNext
          Loop
          rstEnergie.Close
        Case SANG_D_ABERRATION
          Set rstAberration = g_bdRoles.OpenRecordset("select Nom from Aberration order by nom", dbOpenDynaset)
          Do While Not rstAberration.EOF
            AjouteDonListe rstAberration!Nom
            rstAberration.MoveNext
          Loop
          rstAberration.Close
        Case MAITRISE_DE_TERRAIN
          Set rstTerrain = g_bdRoles.OpenRecordset("select nom from Terrain where type='Terrain' order by nom", dbOpenDynaset)
          Do While Not rstTerrain.EOF
            AjouteDonListe rstTerrain!Nom
            rstTerrain.MoveNext
          Loop
          rstTerrain.Close
        Case MAITRISE_DE_TERRAIN_PLANAIRE
          Set rstTerrain = g_bdRoles.OpenRecordset("select nom from Terrain where type='Planaire' order by nom", dbOpenDynaset)
          Do While Not rstTerrain.EOF
            AjouteDonListe rstTerrain!Nom
            rstTerrain.MoveNext
          Loop
          rstTerrain.Close
        Case GRAND_ARCANE
          Set rstGrandArcane = g_bdRoles.OpenRecordset("select Nom from GrandArcane order by nom", dbOpenDynaset)
          Do While Not rstGrandArcane.EOF
            AjouteDonListe rstGrandArcane!Nom
            rstGrandArcane.MoveNext
          Loop
          rstGrandArcane.Close
        Case MARQUE_FERALE, MARQUE_FERALE_SUPPLEMENTAIRE
          Set rstGrandArcane = g_bdRoles.OpenRecordset("select Nom from Marqueferal order by nom", dbOpenDynaset)
          Do While Not rstGrandArcane.EOF
            AjouteDonListe rstGrandArcane!Nom
            rstGrandArcane.MoveNext
          Loop
          rstGrandArcane.Close
        Case CHANGEMENT_DE_TAILLE
          Set rstTaille = g_bdRoles.OpenRecordset("select Nom from Taille order by nom", dbOpenDynaset)
          Do While Not rstTaille.EOF
            AjouteDonListe rstTaille!Nom
            rstTaille.MoveNext
          Loop
          rstTaille.Close
          
        Case ECOLE_RENFORCEE, ECOLE_INTERDITE_SUPPLEMENTAIRE, TATOUAGE_FOCALISATEUR, ECOLE_SUPERIEURE, PROTECTION_PROFANE
          For m = 1 To .CmbEcoleSpecialisation.ListCount - 1
            AjouteDonListe .CmbEcoleSpecialisation.List(m)
          Next m
        Case DOMAINE_DE_PREDILECTION, SPONTANEITE_DE_DOMAINE
          For m = 0 To 3
            If .cmbDomaine(m) = "" Then
              Exit For
            End If
            AjouteDonListe .cmbDomaine(m)
          Next m
        
        Case MAITRISE_DES_CONNAISSANCES
          Set rstcompetence = g_bdRoles.OpenRecordset("select Nom from Competence order by nom", dbOpenDynaset)
          Do While Not rstcompetence.EOF
            If InStr(1, rstcompetence!Nom, "Connaissances (", vbTextCompare) <> 0 Then
              AjouteDonListe rstcompetence!Nom
            End If
            rstcompetence.MoveNext
          Loop
          rstcompetence.Close
        Case TALENT
          Set rstcompetence = g_bdRoles.OpenRecordset("select Nom from Competence order by nom", dbOpenDynaset)
          Do While Not rstcompetence.EOF
            AjouteDonListe rstcompetence!Nom
            rstcompetence.MoveNext
          Loop
          rstcompetence.Close
        Case SORT_DRIDER
          If m_race.RACE = "Drider" Then
            ReDim tabSortDrider(3)
            tabSortDrider(0) = "Prêtre"
            tabSortDrider(1) = "Magicien"
            tabSortDrider(2) = "Ensorceleur"
            For m = 0 To 2
              AjouteDonListe tabSortDrider(m)
            Next m
          End If
        Case LANCEUR_DE_SORT_POLYVALENT, BLOQUE_PROGRESSION
          For m = 0 To nClassLanceurSort - 1
            AjouteDonListe tabClassLanceurSort(m)
          Next m
        
        Case METAMAGIE_DIVINE
          For m = 0 To vs1.Rows - 1
            If vs1.Cell(flexcpText, m, L8_COL_DON_GENRE) = "Métamagie" Then
              AjouteDonListe vs1.Cell(flexcpText, m, L8_COL_DON_LIB)
            End If
          Next m
        Case ATTAQUE_SPECIALE_RENFORCEE
          Set rst2 = g_bdRoles.OpenRecordset("Select Nom, pouvoir from Dons order by Nom", dbOpenDynaset)
          Do While Not rst2.EOF
            If rst2!pouvoir Or rst2!Nom = DECHARGE_FANTASTIQUE Then
              AjouteDonListe rst2!Nom
            End If
            rst2.MoveNext
          Loop
          rst2.Close
        
        Case ALIGNEMENT_RENFORCE, DEVOTION_EPIQUE
          Set rstAlignement = g_bdRoles.OpenRecordset("select Nom from Alignement order by nom", dbOpenDynaset)
          Do While Not rstAlignement.EOF
            AjouteDonListe rstAlignement!Nom
            rstAlignement.MoveNext
          Loop
          rstAlignement.Close
          
        Case Else
          If rst!liearme Then
            If Genre <> "Prouesse à l'arme exotique" Then
              For m = 1 To .vsEquip.Rows - 1
                If objet_estarme(.vsEquip.Cell(flexcpText, m, L3_COL_EQUIP_TYPE)) Then
                  tblObjet.Seek "=", .vsEquip.Cell(flexcpText, m, L3_COL_EQUIP_NOM_OBJET)
                  If Not tblObjet.NoMatch Then
                    AjouteDonListe tblObjet!GroupeObjet
                  End If
                End If
              Next m
            Else
              If ChercheNiveauMinClasse("Maître d'armes exotiques", "") <> 0 Then
                For m = 1 To .vsEquip.Rows - 1
                  If objet_estarme(.vsEquip.Cell(flexcpText, m, L3_COL_EQUIP_TYPE)) Then
                    tblObjet.Seek "=", .vsEquip.Cell(flexcpText, m, L3_COL_EQUIP_NOM_OBJET)
                    If Not tblObjet.NoMatch Then
                      AjouteDonListe tblObjet!GroupeObjet
                    End If
                  End If
                Next m
              End If
            End If
          Else
            If Not IsNull(rst!Nom) Then
              If NomDon = MANUEL_AUGMENTATION_CARACTERISTIQUE_FORCE Or _
                NomDon = MANUEL_AUGMENTATION_CARACTERISTIQUE_DEXTERITE Or _
                NomDon = MANUEL_AUGMENTATION_CARACTERISTIQUE_CONSTITUTION Or _
                NomDon = MANUEL_AUGMENTATION_CARACTERISTIQUE_INTELLIGENCE Or _
                NomDon = MANUEL_AUGMENTATION_CARACTERISTIQUE_SAGESSE Or _
                NomDon = MANUEL_AUGMENTATION_CARACTERISTIQUE_CHARISME Or _
                NomDon = FORCE_SURHUMAINE Or _
                NomDon = DEXTERITE_SURHUMAINE Or _
                NomDon = CONSTITUTION_SURHUMAINE Or _
                NomDon = INTELLIGENCE_SURHUMAINE Or _
                NomDon = SAGESSE_SURHUMAINE Or _
                NomDon = CHARISME_SURHUMAIN Or _
                NomDon = EMPLACEMENT_DE_SORT_SUPPLEMENTAIRE Or _
                NomDon = SORT_SUPPLEMENTAIRE Or _
                NomDon = MALUS_NLS Or _
                NomDon = CONNAISSANCE_MAGIQUE Then
                
                If NomDon = FORCE_SURHUMAINE Or _
                  NomDon = DEXTERITE_SURHUMAINE Or _
                  NomDon = CONSTITUTION_SURHUMAINE Or _
                  NomDon = INTELLIGENCE_SURHUMAINE Or _
                  NomDon = SAGESSE_SURHUMAINE Or _
                  NomDon = CHARISME_SURHUMAIN Or _
                  NomDon = EMPLACEMENT_DE_SORT_SUPPLEMENTAIRE Or _
                  NomDon = SORT_SUPPLEMENTAIRE Or _
                  NomDon = MALUS_NLS Or _
                  NomDon = CONNAISSANCE_MAGIQUE Then
                  n = 10
                Else
                  n = 5
                End If
                If NomDon = EMPLACEMENT_DE_SORT_SUPPLEMENTAIRE Or _
                  NomDon = SORT_SUPPLEMENTAIRE Or _
                  NomDon = CONNAISSANCE_MAGIQUE Then
                  empl = 1
                End If
                For m = 1 - empl To n - empl
                  AjouteDonListe m
                Next m
              Else
                If ExistInDonNiveau(NomDon, vs1) = False Or Multiple Then   ''Attention il n'y a pas de trie dans ExistInDonNiveau
                    vs2.AddItem NomDon
                    curlig2 = vs2.Rows - 1
                    vs2.Cell(flexcpText, curlig2, L8_COL_DON) = NomDon
                    vs2.Cell(flexcpText, curlig2, L8_COL_DON_LIB) = ""
                    vs2.Cell(flexcpText, curlig2, L8_COL_DON_GENRE) = Genre
                    vs2.Cell(flexcpText, curlig2, L8_COL_DON_SOURCE) = Source
                End If
              End If
            End If
          End If
      End Select
      rst.MoveNext
    Loop
    rst.Close
    tblObjet.Close
    vs2.Redraw = True
    vs1.Col = L7_COL_DON_NIV
    vs1.Sort = flexSortStringAscending
    vs1.Redraw = True
  End With
  vsDerNiveau


Timer1_Timer_Exit:
  Exit Sub

Timer1_Timer_Error:
  MsgBox "L'erreur suivante s'est produite : " & Err.Description & vbCrLf & vbCrLf & " numéro d'erreur (" & Err.Number & " )", vbCritical, g_strNomApplication
  Print #g_FileDebug, Format(Now, "dd/MM/yyyy hh:mm:ss") & " Erreur n° : (" & Err.Number & ") " & Err.Description & " Procedure : Timer1_Timer  Module : frmDonsComp.frm"
  Resume Timer1_Timer_Exit
End Sub
Private Sub CouleurDomaine(Classe As String, ByRef nColor As Long, IndexCompetence As Integer, NiveauPretre As Integer)

If Classe = "Prêtre" Or Classe = "Prêtre cloîtré" Then
  nColor = COUL_COMPCLASSE
  LabCompetence(IndexCompetence).Caption = LabCompetence(IndexCompetence).Caption & " (Domaine)"
Else
  If NiveauGlobal > NiveauPretre Then
    nColor = COUL_COMPOLDCLASSE
  Else
    nColor = COUL_COMPFUTURCLASSE
  End If
End If

End Sub
Private Function EstSkillTriks(Genre As String) As Boolean
EstSkillTriks = False

Select Case Genre
  Case "Mouvement", "Manipulation", "Intéraction"
    EstSkillTriks = True
End Select

End Function
Private Function ChercheNiveauMinClasse(Classe_1 As String, Optional Classe_2 As String) As Integer
Dim RstProgression As Recordset

ChercheNiveauMinClasse = 0
Set RstProgression = g_bdPerso.OpenRecordset("select Classe,Niv from PersonnageProgression where nom='" & FormatStrSQL(frmGestPerso.fldstrnom) & "' order by Niv", dbOpenDynaset)
  Do While Not RstProgression.EOF
    If RstProgression!Classe = Classe_1 Or RstProgression!Classe = Classe_2 Then
      ChercheNiveauMinClasse = RstProgression!niv
      RstProgression.Close
      Exit Function
    End If
    RstProgression.MoveNext
  Loop
  
RstProgression.Close

End Function


Private Sub vs1_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
  If Row > 0 And Col = L7_COL_DON_AIDE Then
    vs1.ComboList = "..."
  Else
    Cancel = True
    vs1.ComboList = ""
  End If
End Sub

Private Sub vs1_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
  If Row > 0 And Col = L7_COL_DON_AIDE Then
    frmDescriptifDon.LabDon = vs1.Cell(flexcpText, Row, L7_COL_DON)
    frmDescriptifDon.m_Descriptif = False
    frmDescriptifDon.Show vbModal
  End If
End Sub
Private Sub vs2_DblClick()
  If vs2.SelectedRow(0) >= 0 Then
    frmDescriptifDon.LabDon = vs2.Cell(flexcpText, vs2.SelectedRow(0), L8_COL_DON)
    frmDescriptifDon.m_Descriptif = False
    frmDescriptifDon.Show vbModal
  End If
End Sub
Private Sub vs1_DblClick()
  If vs1.SelectedRow(0) >= 0 Then
  frmDescriptifDon.LabDon = vs1.Cell(flexcpText, vs1.SelectedRow(0), L7_COL_DON)
  frmDescriptifDon.m_Descriptif = False
  frmDescriptifDon.Show vbModal
  End If
End Sub

Private Sub vs2_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
  If Row > 0 And Col = L8_COL_DON_AIDE Then
    vs2.ComboList = "..."
  Else
    Cancel = True
    vs1.ComboList = ""
  End If
End Sub

Private Sub vs2_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
  If Row > 0 And Col = L8_COL_DON_AIDE Then
    frmDescriptifDon.LabDon = vs2.Cell(flexcpText, Row, L8_COL_DON)
    frmDescriptifDon.m_Descriptif = False
    frmDescriptifDon.Show vbModal
  End If
End Sub
Public Function PointDouble() As Integer
Dim i As Integer

PointDouble = 2
  For i = 1 To vs1.Rows - 1
    If vs1.Cell(flexcpText, i, L7_COL_DON) = ABLE_LEARNER Then
      PointDouble = 1
      Exit For
    End If
  Next i
End Function
Private Sub AjouteDonListe(ByVal Libelle As String)
Dim strDon As String
Dim LigneCur2 As Long

If ExistInDon(NomDon, Libelle, vs1, False) = False Or Multiple Then
  strDon = NomDon & " [" & Libelle & "]"
  'If ExistInDon(NomDon, Libelle, vs2) = False Or Multiple Then
    vs2.AddItem strDon
    LigneCur2 = vs2.Rows - 1
    vs2.Cell(flexcpText, LigneCur2, L8_COL_DON) = NomDon
    vs2.Cell(flexcpText, LigneCur2, L8_COL_DON_LIB) = Libelle
    vs2.Cell(flexcpText, LigneCur2, L8_COL_DON_GENRE) = Genre
    vs2.Cell(flexcpText, LigneCur2, L8_COL_DON_SOURCE) = Source
  'End If
End If

End Sub
