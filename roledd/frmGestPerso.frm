VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{D76D7130-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7d.ocx"
Begin VB.Form frmGestPerso 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Gestion personnages"
   ClientHeight    =   9480
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13305
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9480
   ScaleWidth      =   13305
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton btnRenommer 
      Caption         =   "&Renommer"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10305
      TabIndex        =   67
      Top             =   45
      Width           =   1695
   End
   Begin VB.CommandButton btnBackground 
      Caption         =   "Background"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8595
      TabIndex        =   59
      Top             =   45
      Width           =   1695
   End
   Begin VB.CommandButton btnAnnuler 
      Caption         =   "Annuler (ESC)"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6885
      TabIndex        =   64
      Top             =   45
      Width           =   1695
   End
   Begin VB.CommandButton btnSupprimer 
      Caption         =   "Supprimer (F10)"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5175
      TabIndex        =   63
      Top             =   45
      Width           =   1695
   End
   Begin VB.CommandButton btnEnregistrer 
      Caption         =   "Enregistrer (F5)"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3465
      TabIndex        =   62
      Top             =   45
      Width           =   1695
   End
   Begin VB.CommandButton btnAjouter 
      Caption         =   "Ajouter (F9)"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1755
      TabIndex        =   61
      Top             =   45
      Width           =   1695
   End
   Begin VB.CommandButton btnFermer 
      Caption         =   "Fermer (ESC)"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   45
      TabIndex        =   60
      Top             =   45
      Width           =   1695
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   11925
      Top             =   0
   End
   Begin VB.Frame Frame3 
      Caption         =   "Equipement, richesse, langues,alphabets, artisanats et professions"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4440
      Left            =   5310
      TabIndex        =   38
      Top             =   5040
      Width           =   7995
      Begin VB.CommandButton btnEquipInserer 
         Caption         =   "&Acheter un objet"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   1
         Left            =   4200
         TabIndex        =   97
         Top             =   2512
         Width           =   1560
      End
      Begin VB.CommandButton btnEquipInserer 
         Caption         =   "&Inserer un objet (F8)"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   2400
         TabIndex        =   96
         Top             =   2512
         Width           =   1680
      End
      Begin VB.CommandButton btnEquipEnlever 
         Caption         =   "&Enlever (vendre) un objet"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   120
         TabIndex        =   95
         Top             =   2512
         Width           =   2160
      End
      Begin VB.TextBox fldnTotalPo 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6960
         MaxLength       =   7
         TabIndex        =   82
         Top             =   2520
         Width           =   855
      End
      Begin VB.ComboBox CmbProfession 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   3
         Left            =   6240
         TabIndex        =   78
         Text            =   "Profession 4"
         Top             =   3840
         Width           =   1700
      End
      Begin VB.ComboBox CmbProfession 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   2
         Left            =   4486
         TabIndex        =   77
         Text            =   "Profession 3"
         Top             =   3840
         Width           =   1700
      End
      Begin VB.ComboBox CmbProfession 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   1
         Left            =   2734
         TabIndex        =   76
         Text            =   "Profession 2"
         Top             =   3840
         Width           =   1700
      End
      Begin VB.ComboBox CmbProfession 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   0
         Left            =   982
         TabIndex        =   75
         Text            =   "Profession 1"
         Top             =   3840
         Width           =   1700
      End
      Begin VB.ComboBox CmbArtisanat 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   2
         Left            =   5115
         TabIndex        =   73
         Text            =   "Artisanat 3"
         Top             =   3375
         Width           =   2000
      End
      Begin VB.ComboBox CmbArtisanat 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   1
         Left            =   3030
         TabIndex        =   72
         Text            =   "Artisanat 2"
         Top             =   3375
         Width           =   2000
      End
      Begin VB.ComboBox CmbArtisanat 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   0
         Left            =   945
         TabIndex        =   70
         Text            =   "CmbArtisanat"
         Top             =   3375
         Width           =   1995
      End
      Begin VB.TextBox fldstrLangueConnue 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   765
         TabIndex        =   53
         Top             =   2985
         Width           =   3150
      End
      Begin VB.TextBox fldstrAlphabet 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4800
         TabIndex        =   52
         Top             =   3000
         Width           =   3150
      End
      Begin VSFlex7DAOCtl.VSFlexGrid vsEquip 
         Height          =   2265
         Left            =   45
         TabIndex        =   39
         Top             =   225
         Width           =   7860
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
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483646
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
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
         FormatString    =   $"frmGestPerso.frx":0000
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
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
         TabBehavior     =   1
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
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         Caption         =   "Richesse (PO)"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   5880
         TabIndex        =   81
         Top             =   2565
         Width           =   1035
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         Caption         =   "Professions"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   90
         TabIndex        =   74
         Top             =   3900
         Width           =   840
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "Artisanats"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   90
         TabIndex        =   71
         Top             =   3420
         Width           =   750
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "Langues"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   90
         TabIndex        =   55
         Top             =   3015
         Width           =   585
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Alphabets"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   3960
         TabIndex        =   54
         Top             =   3000
         Width           =   750
      End
   End
   Begin VB.Frame Frame2 
      Height          =   9000
      Left            =   45
      TabIndex        =   19
      Top             =   480
      Width           =   5250
      Begin VB.ComboBox CmbShugenja 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   1
         Left            =   3550
         Style           =   2  'Dropdown List
         TabIndex        =   109
         Top             =   8080
         Width           =   1620
      End
      Begin VB.ComboBox CmbShugenja 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   0
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   108
         Top             =   8080
         Width           =   900
      End
      Begin VB.ComboBox CmbDragonTotem 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   4035
         Style           =   2  'Dropdown List
         TabIndex        =   105
         Top             =   6600
         Width           =   1140
      End
      Begin VB.ComboBox cmbDomaine 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   3
         Left            =   3690
         Style           =   2  'Dropdown List
         TabIndex        =   103
         Top             =   6240
         Width           =   1500
      End
      Begin VB.ComboBox CmbEcole_interdite 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   3
         ItemData        =   "frmGestPerso.frx":0088
         Left            =   1320
         List            =   "frmGestPerso.frx":008A
         Style           =   2  'Dropdown List
         TabIndex        =   101
         Top             =   6600
         Width           =   1500
      End
      Begin VB.ComboBox CmbEcole_interdite 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   2
         ItemData        =   "frmGestPerso.frx":008C
         Left            =   1320
         List            =   "frmGestPerso.frx":008E
         Style           =   2  'Dropdown List
         TabIndex        =   100
         Top             =   6240
         Width           =   1500
      End
      Begin VB.ComboBox CmbWujen 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1300
         Style           =   2  'Dropdown List
         TabIndex        =   94
         Top             =   7680
         Width           =   1380
      End
      Begin VB.ComboBox CmbElu 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   2
         Left            =   3690
         Style           =   2  'Dropdown List
         TabIndex        =   93
         Top             =   7680
         Width           =   1500
      End
      Begin VB.ComboBox CmbElu 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   1
         Left            =   3690
         Style           =   2  'Dropdown List
         TabIndex        =   92
         Top             =   7320
         Width           =   1500
      End
      Begin VB.ComboBox CmbElu 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   0
         Left            =   3690
         Style           =   2  'Dropdown List
         TabIndex        =   91
         Top             =   6960
         Width           =   1500
      End
      Begin VB.ComboBox CmbSorcier 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   1
         Left            =   1300
         Style           =   2  'Dropdown List
         TabIndex        =   90
         Top             =   7320
         Width           =   1380
      End
      Begin VB.ComboBox CmbSorcier 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   0
         Left            =   1300
         Style           =   2  'Dropdown List
         TabIndex        =   89
         Top             =   6960
         Width           =   1380
      End
      Begin VB.ComboBox cmbAlignement 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3105
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   3555
         Width           =   2010
      End
      Begin VB.ComboBox Cmbarchetype 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   990
         Style           =   2  'Dropdown List
         TabIndex        =   65
         Top             =   2700
         Width           =   2055
      End
      Begin VB.ComboBox CmbEcole_interdite 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   1
         ItemData        =   "frmGestPerso.frx":0090
         Left            =   1300
         List            =   "frmGestPerso.frx":0092
         Style           =   2  'Dropdown List
         TabIndex        =   50
         Top             =   5865
         Width           =   1500
      End
      Begin VB.ComboBox CmbEcole_interdite 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   0
         ItemData        =   "frmGestPerso.frx":0094
         Left            =   1300
         List            =   "frmGestPerso.frx":0096
         Style           =   2  'Dropdown List
         TabIndex        =   48
         Top             =   5505
         Width           =   1500
      End
      Begin VB.ComboBox cmbDomaine 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   2
         Left            =   3690
         Style           =   2  'Dropdown List
         TabIndex        =   47
         Top             =   5865
         Width           =   1500
      End
      Begin VB.ComboBox cmbDomaine 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   1
         Left            =   3690
         Style           =   2  'Dropdown List
         TabIndex        =   46
         Top             =   5505
         Width           =   1500
      End
      Begin VB.ComboBox cmbDomaine 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   0
         Left            =   3690
         Style           =   2  'Dropdown List
         TabIndex        =   45
         Top             =   5145
         Width           =   1500
      End
      Begin VB.ComboBox CmbEcoleSpecialisation 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "frmGestPerso.frx":0098
         Left            =   1300
         List            =   "frmGestPerso.frx":009A
         Style           =   2  'Dropdown List
         TabIndex        =   41
         Top             =   5145
         Width           =   1500
      End
      Begin VB.TextBox fldnMalusXP 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4260
         MaxLength       =   3
         TabIndex        =   14
         Top             =   4740
         Width           =   855
      End
      Begin VB.TextBox fldnAge 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2580
         MaxLength       =   3
         TabIndex        =   13
         Top             =   4740
         Width           =   855
      End
      Begin VB.TextBox flddate 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   690
         MaxLength       =   50
         TabIndex        =   12
         Top             =   4740
         Width           =   885
      End
      Begin VB.TextBox fldnTaille 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4260
         MaxLength       =   3
         TabIndex        =   11
         Top             =   4380
         Width           =   855
      End
      Begin VB.TextBox fldnPoids 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2580
         MaxLength       =   5
         TabIndex        =   10
         Top             =   4380
         Width           =   855
      End
      Begin VB.TextBox fldstrSexe 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   690
         MaxLength       =   10
         TabIndex        =   9
         Top             =   4380
         Width           =   885
      End
      Begin VB.TextBox fldstryeux 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4260
         MaxLength       =   15
         TabIndex        =   8
         Top             =   4020
         Width           =   855
      End
      Begin VB.TextBox fldstrcheveux 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2580
         MaxLength       =   15
         TabIndex        =   7
         Top             =   4020
         Width           =   855
      End
      Begin VB.TextBox fldnbeaute 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   690
         MaxLength       =   2
         TabIndex        =   6
         Top             =   4020
         Width           =   885
      End
      Begin VB.TextBox fldstrdieu 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   660
         MaxLength       =   30
         TabIndex        =   4
         Top             =   3570
         Width           =   1455
      End
      Begin VB.TextBox fldstrtitre 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3090
         MaxLength       =   30
         TabIndex        =   3
         Top             =   3240
         Width           =   2010
      End
      Begin VB.TextBox fldstrprofil 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   660
         MaxLength       =   30
         TabIndex        =   2
         Top             =   3240
         Width           =   1875
      End
      Begin VB.TextBox fldstrnom 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   90
         MaxLength       =   50
         TabIndex        =   0
         Top             =   2295
         Width           =   2955
      End
      Begin VB.ComboBox CmbRace 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3105
         TabIndex        =   1
         Text            =   "Combo1"
         Top             =   2295
         Width           =   2055
      End
      Begin VSFlex7DAOCtl.VSFlexGrid vsPerso 
         Bindings        =   "frmGestPerso.frx":009C
         Height          =   2115
         Left            =   90
         TabIndex        =   20
         Top             =   135
         Width           =   5040
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
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   0
         SelectionMode   =   3
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   9
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   "Nom                                                                    |Race                 "
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   -1  'True
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   1
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
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         Caption         =   "Ordre Shugenja"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   8
         Left            =   2400
         TabIndex        =   107
         Top             =   8100
         Width           =   1095
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         Caption         =   "Elément Shugenja"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   7
         Left            =   90
         TabIndex        =   106
         Top             =   8100
         Width           =   1260
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         Caption         =   "Dragon Totem"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   6
         Left            =   2900
         TabIndex        =   104
         Top             =   6680
         Width           =   1050
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "Domaine 4"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   3
         Left            =   2880
         TabIndex        =   102
         Top             =   6310
         Width           =   780
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "Ecole interdite 4"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   3
         Left            =   90
         TabIndex        =   99
         Top             =   6600
         Width           =   1170
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "Ecole interdite 3"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   2
         Left            =   90
         TabIndex        =   98
         Top             =   6240
         Width           =   1170
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         Caption         =   "Elément Wujen"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   5
         Left            =   90
         TabIndex        =   88
         Top             =   7740
         Width           =   1095
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         Caption         =   "Energie Elu 3"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   4
         Left            =   2700
         TabIndex        =   87
         Top             =   7740
         Width           =   945
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         Caption         =   "Energie Elu 2"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   3
         Left            =   2700
         TabIndex        =   86
         Top             =   7380
         Width           =   945
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         Caption         =   "Energie Elu 1"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   2
         Left            =   2700
         TabIndex        =   85
         Top             =   7005
         Width           =   945
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         Caption         =   "Energie sorcier 2"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   90
         TabIndex        =   84
         Top             =   7380
         Width           =   1185
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         Caption         =   "Energie sorcier 1"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   90
         TabIndex        =   83
         Top             =   7020
         Width           =   1185
      End
      Begin VB.Label Label25 
         Caption         =   "Archétype"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   135
         TabIndex        =   66
         Top             =   2730
         Width           =   990
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "Domaine 1"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   2865
         TabIndex        =   58
         Top             =   5190
         Width           =   780
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "Domaine 2"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   2865
         TabIndex        =   57
         Top             =   5550
         Width           =   780
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "Domaine 3"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   2
         Left            =   2865
         TabIndex        =   56
         Top             =   5925
         Width           =   780
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "Ecole interdite 2"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   90
         TabIndex        =   51
         Top             =   5925
         Width           =   1170
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "Ecole interdite 1"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   90
         TabIndex        =   49
         Top             =   5565
         Width           =   1170
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Spécialisation"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   90
         TabIndex        =   40
         Top             =   5205
         Width           =   990
      End
      Begin VB.Label Label19 
         Caption         =   "MalusXP"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3480
         TabIndex        =   37
         Top             =   4770
         Width           =   735
      End
      Begin VB.Label Label18 
         Caption         =   "Age"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1650
         TabIndex        =   36
         Top             =   4770
         Width           =   855
      End
      Begin VB.Label Label17 
         Caption         =   "Date"
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
         Left            =   60
         TabIndex        =   35
         Top             =   4770
         Width           =   615
      End
      Begin VB.Line Line3 
         X1              =   60
         X2              =   5100
         Y1              =   3930
         Y2              =   3930
      End
      Begin VB.Label Label16 
         Caption         =   "Taille (cm)"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3450
         TabIndex        =   34
         Top             =   4410
         Width           =   825
      End
      Begin VB.Label Label3 
         Caption         =   "Poids (Kg)"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1650
         TabIndex        =   33
         Top             =   4410
         Width           =   855
      End
      Begin VB.Label Label9 
         Caption         =   "Sexe"
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
         Left            =   60
         TabIndex        =   32
         Top             =   4410
         Width           =   615
      End
      Begin VB.Label Label8 
         Caption         =   "Yeux"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3480
         TabIndex        =   31
         Top             =   4050
         Width           =   645
      End
      Begin VB.Label Label7 
         Caption         =   "Cheveux"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1650
         TabIndex        =   30
         Top             =   4050
         Width           =   855
      End
      Begin VB.Line Line2 
         X1              =   90
         X2              =   5130
         Y1              =   5100
         Y2              =   5100
      End
      Begin VB.Label Label6 
         Caption         =   "Beauté"
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
         Left            =   60
         TabIndex        =   29
         Top             =   4050
         Width           =   615
      End
      Begin VB.Label Label5 
         Caption         =   "Alignement:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2220
         TabIndex        =   28
         Top             =   3600
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Dieu"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   60
         TabIndex        =   27
         Top             =   3600
         Width           =   465
      End
      Begin VB.Label Label2 
         Caption         =   "Titre:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2670
         TabIndex        =   26
         Top             =   3285
         Width           =   405
      End
      Begin VB.Line Line1 
         X1              =   90
         X2              =   5160
         Y1              =   3120
         Y2              =   3120
      End
      Begin VB.Label Label1 
         Caption         =   "Profil"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   45
         TabIndex        =   25
         Top             =   3285
         Width           =   495
      End
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   10440
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   45
      Visible         =   0   'False
      Width           =   1230
   End
   Begin VB.Frame Frame1 
      Height          =   4560
      Left            =   5310
      TabIndex        =   16
      Top             =   450
      Width           =   7995
      Begin VB.TextBox fldnTotalXP 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6960
         MaxLength       =   7
         TabIndex        =   80
         Top             =   2880
         Width           =   855
      End
      Begin VB.CheckBox ChkDonsAPrendre 
         Caption         =   "Marquer les dons à prendre"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   5625
         TabIndex        =   69
         Top             =   1800
         Width           =   2310
      End
      Begin VB.CheckBox ChkDonsObtenus 
         Caption         =   "Marquer les dons obtenus"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   5625
         TabIndex        =   68
         Top             =   2280
         Width           =   2190
      End
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   690
         Left            =   5595
         ScaleHeight     =   690
         ScaleWidth      =   1980
         TabIndex        =   42
         Top             =   1080
         Width           =   1980
         Begin VB.CommandButton btnNivInserer 
            Caption         =   "&Inserer un niveau (F7)"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   0
            TabIndex        =   44
            Top             =   360
            Width           =   1995
         End
         Begin VB.CommandButton btnNivEnlever 
            Caption         =   "&Enlever un niveau"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   0
            TabIndex        =   43
            Top             =   0
            Width           =   1995
         End
      End
      Begin VB.PictureBox Picture1 
         Height          =   870
         Left            =   6120
         Picture         =   "frmGestPerso.frx":00B0
         ScaleHeight     =   810
         ScaleWidth      =   1035
         TabIndex        =   24
         Top             =   180
         Width           =   1095
      End
      Begin VB.CommandButton btnRedescendre 
         Height          =   285
         Left            =   7620
         Picture         =   "frmGestPerso.frx":30FA
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   1440
         Width           =   285
      End
      Begin VB.CommandButton btnRemonter 
         Height          =   330
         Left            =   7620
         Picture         =   "frmGestPerso.frx":36D0
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   1080
         Width           =   285
      End
      Begin VSFlex7DAOCtl.VSFlexGrid vsHistoPerso 
         Height          =   4305
         Left            =   45
         TabIndex        =   15
         Top             =   135
         Width           =   5550
         _cx             =   5080
         _cy             =   5080
         _ConvInfo       =   1
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   8.25
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
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483641
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   0
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   0
         SelectionMode   =   3
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   "Niv|Classe                        |PV|For|Dex|Con|Int|Sag|Cha |Comp/Don"
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   1
         OwnerDraw       =   0
         Editable        =   1
         ShowComboButton =   -1  'True
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   1
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
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         Caption         =   "Expérience (XP)"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   5640
         TabIndex        =   79
         Top             =   2880
         Width           =   1185
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   135
      Left            =   0
      TabIndex        =   21
      Top             =   9345
      Width           =   13305
      _ExtentX        =   23469
      _ExtentY        =   238
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   23416
         EndProperty
      EndProperty
   End
   Begin VB.Label Label12 
      Height          =   240
      Left            =   2610
      TabIndex        =   18
      Top             =   3915
      Width           =   825
   End
   Begin VB.Label Label31 
      Height          =   240
      Left            =   45
      TabIndex        =   17
      Top             =   3870
      Width           =   825
   End
End
Attribute VB_Name = "frmGestPerso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bModif As Boolean
Dim bRefrech As Boolean
Dim bChargePerso As Boolean
Dim m_tabButton() As S_TABVAL
Private Sub BloqueCombo()
Dim bOk As Boolean
Dim i As Integer
  
  ''debloque les combo specialisation si le personnage est mago
  bOk = False
  For i = 1 To vsHistoPerso.Rows - 1
    If UCase(vsHistoPerso.Cell(flexcpText, i, L1_COL_CLASSE)) = "MAGICIEN" Or UCase(vsHistoPerso.Cell(flexcpText, i, L1_COL_CLASSE)) = "INCANTATRIXE" _
    Or UCase(vsHistoPerso.Cell(flexcpText, i, L1_COL_CLASSE)) = "MAGICIEN ELFE" Then
      bOk = True
      Exit For
    End If
  Next i
  CmbEcoleSpecialisation.Enabled = bOk
  For i = 0 To 3
    CmbEcole_interdite(i).Enabled = bOk
  Next i
  
  If bOk = False Then
    CmbEcoleSpecialisation.ListIndex = -1
    For i = 0 To 3
      CmbEcole_interdite(i).ListIndex = -1
    Next i
  End If
  ''debloque les combo domaine si le personnage est pretre
  bOk = False
  For i = 1 To vsHistoPerso.Rows - 1
    If vsHistoPerso.Cell(flexcpText, i, L1_COL_CLASSE) = "Prêtre" Or UCase(vsHistoPerso.Cell(flexcpText, i, L1_COL_CLASSE)) = "DRUIDE" Or vsHistoPerso.Cell(flexcpText, i, L1_COL_CLASSE) = "Prêtre cloîtré" Or vsHistoPerso.Cell(flexcpText, i, L1_COL_CLASSE) = "Croisé divin" Then
      bOk = True
      Exit For
    End If
  Next i
  If CmbRace = "Ghaele" Or CmbRace = "Archon messager" Or CmbRace = "Ange planétar" Or CmbRace = "Ange solar" Then
    bOk = True
  End If
  cmbDomaine(0).Enabled = bOk
  cmbDomaine(1).Enabled = bOk
  cmbDomaine(2).Enabled = bOk
  cmbDomaine(3).Enabled = bOk
  
  If bOk = False Then
    cmbDomaine(0).ListIndex = -1
    cmbDomaine(1).ListIndex = -1
    cmbDomaine(2).ListIndex = -1
    cmbDomaine(3).ListIndex = -1
  End If
  ''debloque les combo Energie si le personnage est Elu divin
  bOk = False
  For i = 1 To vsHistoPerso.Rows - 1
    If UCase(vsHistoPerso.Cell(flexcpText, i, L1_COL_CLASSE)) = "ELU DIVIN" Then
      bOk = True
      Exit For
    End If
  Next i
  
  CmbElu(0).Enabled = bOk
  CmbElu(1).Enabled = bOk
  CmbElu(2).Enabled = bOk
  If bOk = False Then
    CmbElu(0).ListIndex = -1
    CmbElu(1).ListIndex = -1
    CmbElu(2).ListIndex = -1
  End If
  ''debloque les combo Energie si le personnage est Sorcier
  bOk = False
  For i = 1 To vsHistoPerso.Rows - 1
    If UCase(vsHistoPerso.Cell(flexcpText, i, L1_COL_CLASSE)) = "SORCIER" Or vsHistoPerso.Cell(flexcpText, i, L1_COL_CLASSE) = "Savant élémentaire" Then
      bOk = True
      Exit For
    End If
  Next i
  
  CmbSorcier(0).Enabled = bOk
  CmbSorcier(1).Enabled = bOk
  If bOk = False Then
    CmbSorcier(0).ListIndex = -1
    CmbSorcier(1).ListIndex = -1
  End If
  ''debloque la combo Element si le personnage est Wujen
  bOk = False
  For i = 1 To vsHistoPerso.Rows - 1
    If UCase(vsHistoPerso.Cell(flexcpText, i, L1_COL_CLASSE)) = "WUJEN" Then
      bOk = True
      Exit For
    End If
  Next i
  CmbWujen.Enabled = bOk
  If bOk = False Then
    CmbWujen.ListIndex = -1
  End If
  '' debloque la combo si le personnage est Shugenja
  bOk = False
  For i = 1 To vsHistoPerso.Rows - 1
    If UCase(vsHistoPerso.Cell(flexcpText, i, L1_COL_CLASSE)) = "SHUGENJA" Then
      bOk = True
      Exit For
    End If
  Next i
  CmbShugenja(0).Enabled = bOk
  CmbShugenja(1).Enabled = bOk
  If bOk = False Then
    CmbWujen.ListIndex = -1
  End If
  '' debloque la combo si le personnage est Dragon shaman
  bOk = False
  For i = 1 To vsHistoPerso.Rows - 1
    If UCase(vsHistoPerso.Cell(flexcpText, i, L1_COL_CLASSE)) = "DRAGON SHAMAN" Then
      bOk = True
      Exit For
    End If
  Next i
  CmbDragonTotem.Enabled = bOk
  If bOk = False Then
    CmbDragonTotem.ListIndex = -1
  End If
End Sub
Private Sub Fermer()

  If g_ModeDebug = vbUnchecked Then On Error GoTo Fermer_Error
  
  If BtnFermer.Visible = True Then
    Unload Me
  Else
    vsPerso.Enabled = True
    vsPerso_SelChange
  End If

Fermer_Exit:
  Exit Sub

Fermer_Error:
  MsgBox "L'erreur suivante s'est produite : " & Err.Description & vbCrLf & vbCrLf & " numéro d'erreur (" & Err.Number & " )", vbCritical, g_strNomApplication
  Print #g_FileDebug, Format(Now, "dd/MM/yyyy hh:mm:ss") & " Erreur n° : (" & Err.Number & ") " & Err.Description & " Procedure : Fermer  Module : frmGestPerso.frm"
  Resume Fermer_Exit
End Sub
Function Enregistre()
Dim bexist As Boolean
Dim j As Integer, n As Integer, niv As Integer, Col As Integer, nbcarac As Integer
Dim tblPerso As Recordset
Dim tblEquip As Recordset
Dim tblPersoDons As Recordset
Dim tabPerso() As S_PERSO
Dim i As Long
Dim pb As Boolean
Dim sclass As S_CLASS
Dim ClasseMaitre As String
Dim ListeClasse As String
Dim PersoClasse As String
Dim strDon As Variant, strLib As Variant
Dim rst As Recordset
Dim PVMax As Integer

Dim Nom As String

  Enregistre = False
  If Trim(fldstrnom) = "" Then
    Screen.MousePointer = vbNormal
    MsgBox "Saisie d'un nom valide", vbExclamation, Me.Caption
    fldstrnom.SetFocus
    Exit Function
  End If
  
  bexist = False
  For i = 0 To CmbRace.ListCount - 1
    If UCase(CmbRace.List(i)) = UCase(CmbRace) Then
      bexist = True
      Exit For
    End If
  Next i
  
  If bexist = False Then
    Screen.MousePointer = vbNormal
    MsgBox "Saisie d'une race existante au sein de la liste prédefinie", vbExclamation, Me.Caption
    CmbRace.SetFocus
    Exit Function
  End If
  
  If fldnBeaute <> "" Then
    If Not IsNumeric(fldnBeaute) Then
      Screen.MousePointer = vbNormal
      MsgBox "La beauté doit être entrée en tant que valeur numérique", vbExclamation, Me.name
      fldnBeaute.SetFocus
      Exit Function
    End If
  End If
  
  If fldnPoids <> "" Then
    If Not IsNumeric(fldnPoids) Then
      Screen.MousePointer = vbNormal
      MsgBox "Le poids doit être entrée en tant que valeur numérique", vbExclamation, Me.name
      fldnPoids.SetFocus
      Exit Function
    End If
  End If

  If fldnPoids <> "" Then
    If Not IsNumeric(fldnPoids) Then
      Screen.MousePointer = vbNormal
      MsgBox "Le poids doit être entré en tant que valeur numérique", vbExclamation, Me.name
      fldnPoids.SetFocus
      Exit Function
    End If
  End If
  
  If fldnTaille <> "" Then
    If Not IsNumeric(fldnTaille) Then
      Screen.MousePointer = vbNormal
      MsgBox "La taille doit être entrée en tant que valeur numérique", vbExclamation, Me.name
      fldnTaille.SetFocus
      Exit Function
    End If
  End If
  
  If fldnAge <> "" Then
    If Not IsNumeric(fldnAge) Then
      Screen.MousePointer = vbNormal
      MsgBox "L'age doit être entrée en tant que valeur numérique", vbExclamation, Me.name
      fldnAge.SetFocus
      Exit Function
    End If
  End If
  
  If fldnMalusXP <> "" Then
    If Not IsNumeric(fldnMalusXP) Then
      Screen.MousePointer = vbNormal
      MsgBox "Le Malux XP doit être entrée en tant que valeur numérique", vbExclamation, Me.name
      fldnMalusXP.SetFocus
      Exit Function
    End If
  End If
  
  If fldnTotalPo <> "" Then
    If Not IsNumeric(fldnTotalPo) Then
      Screen.MousePointer = vbNormal
      MsgBox "Le nombre de Po doit être entrée en tant que valeur numérique", vbExclamation, Me.name
      fldnTotalPo.SetFocus
      Exit Function
    End If
  End If
  
  If fldnTotalXP <> "" Then
    If Not IsNumeric(fldnTotalXP) Then
      Screen.MousePointer = vbNormal
      MsgBox "Le nombre de points d'expérience doit être entrée en tant que valeur numérique", vbExclamation, Me.name
      fldnTotalXP.SetFocus
      Exit Function
    End If
  End If
  
  If Trim(CmbEcoleSpecialisation) <> "" Then
    bexist = False
    For i = 0 To CmbEcoleSpecialisation.ListCount - 1
      If UCase(CmbEcoleSpecialisation.List(i)) = UCase(CmbEcoleSpecialisation) Then
        bexist = True
        Exit For
      End If
    Next i
    
    If bexist = False Then
      Screen.MousePointer = vbNormal
      MsgBox "Saisie d'un specialisation d'école obligatoire au sein de la liste", vbExclamation, Me.Caption
      CmbEcoleSpecialisation.SetFocus
      Exit Function
    End If
  End If
  
  ''controle la progression perso
  pb = False
  ReDim tabPerso(3)
  n = 0
  ''decompte toutes les classes
  With vsHistoPerso
    For i = 1 To .Rows - 1
      ''controle niveau
      If Val(.Cell(flexcpText, i, L1_COL_NIVEAU)) < 1 Then
        Screen.MousePointer = vbNormal
        MsgBox "Saisie d'un niveau supérieur à zero", vbExclamation, Me.Caption
        .Col = L1_COL_NIVEAU
        .Row = i
        .EditCell
        pb = True
        Exit For
      End If
      
      ''controle que la classe existe
      If ChercheClasse(sclass, .Cell(flexcpText, i, L1_COL_CLASSE), .Cell(flexcpText, i, L1_COL_NIVEAU)) = False Then
        Screen.MousePointer = vbNormal
        MsgBox "Saisie d'une classe existante obligatoire", vbExclamation, Me.Caption
        .Col = L1_COL_CLASSE
        .Row = i
        .EditCell
        pb = True
        Exit For
      End If
      
      ''controle validité PV
      PVMax = sclass.PointDeVie
      
      If EstClasseMonstrueuse(sclass.Classe) Then
        Select Case Cmbarchetype
          Case DEMI_DRAGON
            PVMax = mini_(PVMax + 2, 12)
          Case DEMI_FEE
            PVMax = 6
        End Select
      End If
      If EstMortVivant(Cmbarchetype) Then
        PVMax = 12
      End If
      If Val(.Cell(flexcpText, i, L1_COL_PV)) < 1 Or Val(.Cell(flexcpText, i, L1_COL_PV)) > PVMax Then
        Screen.MousePointer = vbNormal
        MsgBox "Saisie d'un nombre de point de vie compris entre 1 et " & PVMax, vbExclamation, Me.Caption
        .Col = L1_COL_PV
        .Row = i
        .EditCell
        pb = True
        Exit For
      End If
      
      ''controle les carac
      Select Case i
        Case 1
          For Col = L1_COL_FOR To L1_COL_CHA
            If Val(.Cell(flexcpText, i, Col)) < 1 Or Val(.Cell(flexcpText, i, Col)) > 30 Then
              Screen.MousePointer = vbNormal
              MsgBox "Pour le premier niveau les caracteristiques doivent être comprise entre 1 et 30 ", vbExclamation, Me.Caption
              .Col = Col
              .Row = i
              .EditCell
              pb = True
              Exit For
            End If
          Next Col
      End Select
      
      If pb = True Then
        Exit For
      End If
      
      bexist = False
      ''cherche la classe
      For j = 0 To n - 1
        If UCase(.Cell(flexcpText, i, L1_COL_CLASSE)) = UCase(tabPerso(j).Classe_1) Then
          bexist = True
          Exit For
        End If
      Next j
      
      If bexist = False Then
        If n + 1 > UBound(tabPerso) Then
          ReDim Preserve tabPerso(n + 1)
        End If
        j = n
        tabPerso(j).Classe_1 = .Cell(flexcpText, i, L1_COL_CLASSE)
        tabPerso(j).Niv_1 = .Cell(flexcpText, i, L1_COL_NIVEAU)
        n = n + 1
      End If
      
      If .Cell(flexcpText, i, L1_COL_NIVEAU) > tabPerso(j).Niv_1 Then
        tabPerso(j).Niv_1 = .Cell(flexcpText, i, L1_COL_NIVEAU)
      End If
    Next i
  End With
  If pb Then Exit Function
  
  If n = 0 Then
    Screen.MousePointer = vbNormal
    MsgBox "Veuillez au moins inserer une classe et un niveau dans la liste ", vbExclamation, Me.Caption
    vsHistoPerso.SetFocus
    Exit Function
  End If
      
  If n > 8 Then
    Screen.MousePointer = vbNormal
    MsgBox "Ce personnage comprend plus de 8 classes", vbExclamation, Me.Caption
    vsHistoPerso.SetFocus
    Exit Function
  End If
  ''coherence des niveaux
  For j = 0 To n - 1
    For niv = 1 To tabPerso(j).Niv_1
      bexist = False
      ClasseMaitre = sclass.ClasseMaitre
      
      With vsHistoPerso
        For i = 1 To .Rows - 1
          
          Set rst = g_bdRoles.OpenRecordset("classe", dbOpenTable)
          rst.Index = "PrimaryKey"
          rst.Seek "=", .Cell(flexcpText, i, L1_COL_CLASSE)
          ListeClasse = rst!ClasseMaitre
          rst.Seek "=", tabPerso(j).Classe_1
          PersoClasse = rst!ClasseMaitre
          rst.Close
 
          If UCase(ListeClasse) = UCase(PersoClasse) And _
            Val(.Cell(flexcpText, i, L1_COL_NIVEAU)) = niv Then
            bexist = True
            Exit For
          End If
        Next i
      End With
      
      If bexist = False Then
        Screen.MousePointer = vbNormal
        MsgBox "Il manque le niveau numero " & niv & " pour la classe de '" & tabPerso(j).Classe_1 & "'", vbExclamation, Me.Caption
        vsHistoPerso.SetFocus
        Exit Function
      End If
    Next niv
  Next j
  
  
  
  
  Set tblPerso = g_bdPerso.OpenRecordset("Personnage", dbOpenTable)
  With tblPerso
    .Index = "Primarykey"
    .Seek "=", fldstrnom
    If .NoMatch Then
      .AddNew
    Else
      .Edit
    End If
    !Nom = fldstrnom
    !RACE = CmbRace
    
    !Classe_1 = tabPerso(0).Classe_1
    !Niv_1 = tabPerso(0).Niv_1
    If n > 1 Then
      !Classe_2 = tabPerso(1).Classe_1
      !Niv_2 = tabPerso(1).Niv_1
      Else
      !Classe_2 = Null
      !Niv_2 = Null
    End If
    If n > 2 Then
      !Classe_3 = tabPerso(2).Classe_1
      !Niv_3 = tabPerso(2).Niv_1
      Else
      !Classe_3 = Null
      !Niv_3 = Null
    End If
    If n > 3 Then
      !Classe_4 = tabPerso(3).Classe_1
      !Niv_4 = tabPerso(3).Niv_1
      Else
      !Classe_4 = Null
      !Niv_4 = Null
    End If
    If n > 4 Then
      !Classe_5 = tabPerso(4).Classe_1
      !Niv_5 = tabPerso(4).Niv_1
      Else
      !Classe_5 = Null
      !Niv_5 = Null
    End If
    If n > 5 Then
      !Classe_6 = tabPerso(5).Classe_1
      !Niv_6 = tabPerso(5).Niv_1
      Else
      !Classe_6 = Null
      !Niv_6 = Null
    End If
    If n > 6 Then
      !Classe_7 = tabPerso(6).Classe_1
      !Niv_7 = tabPerso(6).Niv_1
      Else
      !Classe_7 = Null
      !Niv_7 = Null
    End If
    If n > 7 Then
      !Classe_8 = tabPerso(7).Classe_1
      !Niv_8 = tabPerso(7).Niv_1
      Else
      !Classe_8 = Null
      !Niv_8 = Null
    End If
    !Archetype = IIf(Trim(Cmbarchetype) <> "", Cmbarchetype, Null)
    !Profil = IIf(Trim(fldstrprofil) <> "", fldstrprofil, Null)
    !Titre = IIf(Trim(fldstrtitre) <> "", fldstrtitre, Null)
    !Dieu = IIf(Trim(fldstrdieu) <> "", fldstrdieu, Null)
    !Alignement = IIf(Trim(CmbAlignement) <> "", CmbAlignement, Null)
    !Beaute = IIf(IsNumeric(fldnBeaute), fldnBeaute, Null)
    !Cheveux = IIf(Trim(fldstrcheveux) <> "", fldstrcheveux, Null)
    !Yeux = IIf(Trim(fldstryeux) <> "", fldstryeux, Null)
    !Sexe = IIf(Trim(fldstrSexe) <> "", fldstrSexe, Null)
    !poids = IIf(IsNumeric(fldnPoids), fldnPoids, Null)
    !taille = IIf(IsNumeric(fldnTaille), fldnTaille, Null)
    !Date = IIf(Trim(flddate) <> "", flddate, Null)
    !Age = IIf(IsNumeric(fldnAge), fldnAge, Null)
    !MalusXP = IIf(IsNumeric(fldnMalusXP), fldnMalusXP, Null)
    !EcoleSpecialisation = IIf(Trim(CmbEcoleSpecialisation) <> "", CmbEcoleSpecialisation, Null)
    !LangueConnue = IIf(Trim(fldstrLangueConnue) <> "", fldstrLangueConnue, Null)
    !Alphabet = IIf(Trim(fldstrAlphabet) <> "", fldstrAlphabet, Null)
    !Domaine_1 = IIf(Trim(cmbDomaine(0)) <> "", cmbDomaine(0), Null)
    !Domaine_2 = IIf(Trim(cmbDomaine(1)) <> "", cmbDomaine(1), Null)
    !Domaine_3 = IIf(Trim(cmbDomaine(2)) <> "", cmbDomaine(2), Null)
    !Domaine_4 = IIf(Trim(cmbDomaine(3)) <> "", cmbDomaine(3), Null)
    !Ecole_interdite_1 = IIf(Trim(CmbEcole_interdite(0)) <> "", CmbEcole_interdite(0), Null)
    !Ecole_interdite_2 = IIf(Trim(CmbEcole_interdite(1)) <> "", CmbEcole_interdite(1), Null)
    !Ecole_interdite_3 = IIf(Trim(CmbEcole_interdite(2)) <> "", CmbEcole_interdite(2), Null)
    !Ecole_interdite_4 = IIf(Trim(CmbEcole_interdite(3)) <> "", CmbEcole_interdite(3), Null)
    !TotalPo = IIf(IsNumeric(fldnTotalPo), fldnTotalPo, 0)
    !TotalXP = IIf(IsNumeric(fldnTotalXP), fldnTotalXP, 0)
    !Profession_1 = IIf(Trim(CmbProfession(0)) <> "", CmbProfession(0), "Profession_1")
    !Profession_2 = IIf(Trim(CmbProfession(1)) <> "", CmbProfession(1), "Profession_2")
    !Profession_3 = IIf(Trim(CmbProfession(2)) <> "", CmbProfession(2), "Profession_3")
    !Profession_4 = IIf(Trim(CmbProfession(3)) <> "", CmbProfession(3), "Profession_4")
    !Artisanat_1 = IIf(Trim(CmbArtisanat(0)) <> "", CmbArtisanat(0), "Artisanat_1")
    !Artisanat_2 = IIf(Trim(CmbArtisanat(1)) <> "", CmbArtisanat(1), "Artisanat_2")
    !Artisanat_3 = IIf(Trim(CmbArtisanat(2)) <> "", CmbArtisanat(2), "Artisanat_3")
    !EnergieElu_1 = IIf(Trim(CmbElu(0)) <> "", CmbElu(0), Null)
    !EnergieElu_2 = IIf(Trim(CmbElu(1)) <> "", CmbElu(1), Null)
    !EnergieElu_3 = IIf(Trim(CmbElu(2)) <> "", CmbElu(2), Null)
    !EnergieSorcier_1 = IIf(Trim(CmbSorcier(0)) <> "", CmbSorcier(0), Null)
    !EnergieSorcier_2 = IIf(Trim(CmbSorcier(1)) <> "", CmbSorcier(1), Null)
    !DragonTotem = IIf(Trim(CmbDragonTotem) <> "", CmbDragonTotem, Null)
    !ElementWujen = IIf(Trim(CmbWujen) <> "", CmbWujen, Null)
    !ElementShugenja = IIf(Trim(CmbShugenja(0)) <> "", CmbShugenja(0), Null)
    !OrdreShugenja = IIf(Trim(CmbShugenja(1)) <> "", CmbShugenja(1), Null)
    .Update
    .Close
  End With
      
  g_bdPerso.Execute "update PersonnageProgression set niv=0 where nom='" & FormatStrSQL(fldstrnom) & "'", dbFailOnError
  g_bdPerso.Execute "update PersonnageDons set niv=0 where nom='" & FormatStrSQL(fldstrnom) & "'", dbFailOnError
  Set tblPerso = g_bdPerso.OpenRecordset("PersonnageProgression", dbOpenTable)
  tblPerso.Index = "index"
  Set tblPersoDons = g_bdPerso.OpenRecordset("PersonnageDons", dbOpenTable)
  tblPersoDons.Index = "index"
  With vsHistoPerso
    For i = 1 To .Rows - 1
      tblPerso.Seek "=", fldstrnom, .Cell(flexcpText, i, L1_COL_CLASSE), .Cell(flexcpText, i, L1_COL_NIVEAU)
      If tblPerso.NoMatch Then
        tblPerso.AddNew
      Else
        tblPerso.Edit
      End If
      tblPerso!Nom = fldstrnom
      tblPerso!Niveau = .Cell(flexcpText, i, L1_COL_NIVEAU)
      tblPerso!Classe = .Cell(flexcpText, i, L1_COL_CLASSE)
      tblPerso!PV = .Cell(flexcpText, i, L1_COL_PV)
      tblPerso!For = IIf(Val(.Cell(flexcpText, i, L1_COL_FOR)) <> 0, .Cell(flexcpText, i, L1_COL_FOR), Null)
      tblPerso!Dex = IIf(Val(.Cell(flexcpText, i, L1_COL_DEX)) <> 0, .Cell(flexcpText, i, L1_COL_DEX), Null)
      tblPerso!Con = IIf(Val(.Cell(flexcpText, i, L1_COL_CON)) <> 0, .Cell(flexcpText, i, L1_COL_CON), Null)
      tblPerso!Int = IIf(Val(.Cell(flexcpText, i, L1_COL_INT)) <> 0, .Cell(flexcpText, i, L1_COL_INT), Null)
      tblPerso!Sag = IIf(Val(.Cell(flexcpText, i, L1_COL_SAG)) <> 0, .Cell(flexcpText, i, L1_COL_SAG), Null)
      tblPerso!Cha = IIf(Val(.Cell(flexcpText, i, L1_COL_CHA)) <> 0, .Cell(flexcpText, i, L1_COL_CHA), Null)
      tblPerso!Competence = .Cell(flexcpText, i, L1_COL_COMPETENCE)
      For j = 0 To L1_NB_DONS - 1
        strDon = IIf(.Cell(flexcpText, i, L1_COL_DON_1 + j) <> "", .Cell(flexcpText, i, L1_COL_DON_1 + j), Null)
        strLib = IIf(.Cell(flexcpText, i, L1_COL_LIB_1 + j) <> "", .Cell(flexcpText, i, L1_COL_LIB_1 + j), Null)
        If Not IsNull(strDon) Then
          tblPersoDons.Seek "=", fldstrnom, .Cell(flexcpText, i, L1_COL_CLASSE), .Cell(flexcpText, i, L1_COL_NIVEAU), strDon, strLib
          If tblPersoDons.NoMatch Then
            tblPersoDons.AddNew
          Else
            tblPersoDons.Edit
          End If
          tblPersoDons!Nom = fldstrnom
          tblPersoDons!Niveau = .Cell(flexcpText, i, L1_COL_NIVEAU)
          tblPersoDons!Classe = .Cell(flexcpText, i, L1_COL_CLASSE)
          tblPersoDons!dons = strDon
          tblPersoDons!Libelle = strLib
          tblPersoDons!niv = i
          tblPersoDons.Update
        End If
      Next j
      tblPerso!modifint = IIf(Val(.Cell(flexcpText, i, L1_COL_MODIF_INT)) > 0, .Cell(flexcpText, i, L1_COL_MODIF_INT), Null)
      tblPerso!niv = i
      tblPerso.Update
    Next i
  End With
  tblPerso.Close
  g_bdPerso.Execute "delete from PersonnageProgression where nom='" & FormatStrSQL(fldstrnom) & "' and niv=0", dbFailOnError
  g_bdPerso.Execute "delete from PersonnageDons where nom='" & FormatStrSQL(fldstrnom) & "' and niv=0", dbFailOnError

  ''inventaire
  g_bdPerso.Execute "update Equipement set pos=0 where personnage='" & FormatStrSQL(fldstrnom) & "'", dbFailOnError
  Set tblEquip = g_bdPerso.OpenRecordset("Equipement", dbOpenTable)
  For i = 1 To vsEquip.Rows - 1
    If Trim(vsEquip.Cell(flexcpText, i, L3_COL_EQUIP_NOM_OBJET)) <> "" Then
      tblEquip.AddNew
      tblEquip!NomObjet = vsEquip.Cell(flexcpText, i, L3_COL_EQUIP_NOM_OBJET)
      tblEquip!Type = IIf(Trim(vsEquip.Cell(flexcpText, i, L3_COL_EQUIP_TYPE)) <> "", vsEquip.Cell(flexcpText, i, L3_COL_EQUIP_TYPE), Null)
      tblEquip!Personnage = fldstrnom
      tblEquip!Place = IIf(IsNumeric(vsEquip.Cell(flexcpText, i, L3_COL_EQUIP_PLACE)), vsEquip.Cell(flexcpText, i, L3_COL_EQUIP_PLACE), Null)
      tblEquip!Pos = i
      tblEquip!CHARGE = IIf(IsNumeric(vsEquip.Cell(flexcpText, i, L3_COL_EQUIP_CHARGE)), vsEquip.Cell(flexcpText, i, L3_COL_EQUIP_CHARGE), Null)
      tblEquip!Multiple = IIf(IsNumeric(vsEquip.Cell(flexcpText, i, L3_COL_EQUIP_MULTIPLE)), vsEquip.Cell(flexcpText, i, L3_COL_EQUIP_MULTIPLE), Null)
      tblEquip!surpersonnage = IIf(vsEquip.Cell(flexcpText, i, L3_COL_EQUIP_SUR_PERSO) = "Oui", True, False)
      tblEquip.Update
    End If
  Next i
  tblEquip.Close
  g_bdPerso.Execute "delete from Equipement where personnage='" & FormatStrSQL(fldstrnom) & "' and pos=0", dbFailOnError
  
  vsPerso.Enabled = True
  Nom = fldstrnom
  RempPerso
  
  RetrouvVSFLex vsPerso, Nom, L2_COL_NOM, 0
    
  Enregistre = True
 
End Function
Sub InverseCarac(carac1 As TextBox, carac2 As TextBox)
Dim nb As Integer

  If g_ModeDebug = vbUnchecked Then On Error GoTo InverseCarac_Error

  If bChargePerso = False Then
    bChargePerso = True
    carac2 = ""
    If IsNumeric(carac1) Then
      nb = CLng(carac1)
      If nb >= -2 And nb <= 2 Then
        If nb < 0 Then
          carac2 = CStr(Abs(nb))
        Else
          If nb > 0 Then
            carac2 = "-" & CStr(Abs(nb))
          Else
            carac2 = "0"
          End If
        End If
      End If
    End If
    If carac2 = "" Then
      carac1 = "0"
      carac2 = "0"
    End If
    bChargePerso = False
  End If

InverseCarac_Exit:
  Exit Sub

InverseCarac_Error:
  MsgBox "L'erreur suivante s'est produite : " & Err.Description & vbCrLf & vbCrLf & " numéro d'erreur (" & Err.Number & " )", vbCritical, g_strNomApplication
  Print #g_FileDebug, Format(Now, "dd/MM/yyyy hh:mm:ss") & " Erreur n° : (" & Err.Number & ") " & Err.Description & " Procedure : InverseCarac  Module : frmGestPerso.frm"
  Resume InverseCarac_Exit
End Sub
Public Sub Modifperso()
Dim ctl As Control

  If g_ModeDebug = vbUnchecked Then On Error GoTo Modifperso_Error
  
  If bChargePerso = False Then
    btnAjouter.Visible = False
    btnSupprimer.Visible = False
    btnBackground.Visible = False
    btnEnregistrer.Visible = True
    btnAnnuler.Visible = True
    BtnFermer.Visible = False
    btnRenommer.Visible = False
    vsPerso.Enabled = False
    AjusteBouton Me, m_tabButton
    bModif = True
  End If

Modifperso_Exit:
  Exit Sub

Modifperso_Error:
  MsgBox "L'erreur suivante s'est produite : " & Err.Description & vbCrLf & vbCrLf & " numéro d'erreur (" & Err.Number & " )", vbCritical, g_strNomApplication
  Print #g_FileDebug, Format(Now, "dd/MM/yyyy hh:mm:ss") & " Erreur n° : (" & Err.Number & ") " & Err.Description & " Procedure : Modifperso  Module : frmGestPerso.frm"
  Resume Modifperso_Exit
End Sub
Private Sub RempPerso()
Dim i As Integer

  If g_ModeDebug = vbUnchecked Then On Error GoTo RempPerso_Error
  
  Data1.DatabaseName = g_bdPerso.name
  Data1.RecordSource = "select  Nom,Race,Endurance,Puissance,Précision,Équilibre,Résistance,Vitalité,Intellect,Érudition,Intuition,Volonté,Magnétisme,Charme,profil,titre,dieu,Alignement,Beaute,Cheveux,Yeux,Sexe,Poids,Taille,Date,Age,MalusXP,EcoleSpecialisation,LangueConnue,alphabet,Domaine_1,Domaine_2,Domaine_3,Ecole_interdite_1,Ecole_interdite_2,Archetype,totalpo,totalxp,profession_1,profession_2,profession_3,profession_4,Artisanat_1,Artisanat_2,Artisanat_3,EnergieElu_1,EnergieElu_2,EnergieElu_3,EnergieSorcier_1,EnergieSorcier_2,ElementWujen,EnnemisJures, Image,Domaine_4,Ecole_interdite_3,Ecole_interdite_4, DragonTotem, ElementShugenja, OrdreShugenja from personnage order by nom"
  bRefrech = False
  Data1.Refresh
  bRefrech = True
  With vsPerso
    .Redraw = False
    .FormatString = "Nom                                                  |Race              "
    .Cols = 59
     ''cache les colonnes
    For i = 2 To .Cols - 1
      .ColHidden(i) = True
    Next i
    .Redraw = True
  End With

RempPerso_Exit:
  Exit Sub

RempPerso_Error:
  MsgBox "L'erreur suivante s'est produite : " & Err.Description & vbCrLf & vbCrLf & " numéro d'erreur (" & Err.Number & " )", vbCritical, g_strNomApplication
  Print #g_FileDebug, Format(Now, "dd/MM/yyyy hh:mm:ss") & " Erreur n° : (" & Err.Number & ") " & Err.Description & " Procedure : RempPerso  Module : frmGestPerso.frm"
  Resume RempPerso_Exit
End Sub
Private Sub RempHistoPerso(strnom As String)
Dim i As Integer, j As Integer
Dim rst As Recordset
Dim tblPersoDons As Recordset

  If g_ModeDebug = vbUnchecked Then On Error GoTo RempHistoPerso_Error

  
  Set rst = g_bdPerso.OpenRecordset("select  Niveau,classe,PV,For,Dex,Con,Int,Sag,Cha,competence,modifInt from personnageProgression where nom='" & FormatStrSQL(strnom) & "' order by niv", dbOpenDynaset)
  Set tblPersoDons = g_bdPerso.OpenRecordset("PersonnageDons", dbOpenTable)
  tblPersoDons.Index = "index"
  With vsHistoPerso
    .Cols = L1_COL_MODIF_INT + 1
    .Redraw = False
    For i = L1_COL_COMPETENCE To L1_COL_MODIF_INT
      .ColHidden(i) = True
    Next i
    i = 1
    .Rows = i
    Do While Not rst.EOF
      .Rows = i + 1
      .Cell(flexcpText, i, L1_COL_NIVEAU) = rst!Niveau
      .Cell(flexcpText, i, L1_COL_CLASSE) = rst!Classe
      .Cell(flexcpText, i, L1_COL_PV) = rst!PV
      .Cell(flexcpText, i, L1_COL_FOR) = rst!For
      .Cell(flexcpText, i, L1_COL_DEX) = rst!Dex
      .Cell(flexcpText, i, L1_COL_CON) = rst!Con
      .Cell(flexcpText, i, L1_COL_INT) = rst!Int
      .Cell(flexcpText, i, L1_COL_SAG) = rst!Sag
      .Cell(flexcpText, i, L1_COL_CHA) = rst!Cha
      .Cell(flexcpText, i, L1_COL_COMPETENCE) = rst!Competence
      tblPersoDons.Seek ">=", strnom, rst!Classe, rst!Niveau
      j = 0
      If Not tblPersoDons.NoMatch Then
        Do While UCase(tblPersoDons!Nom) = UCase(strnom) And _
            UCase(tblPersoDons!Classe) = UCase(rst!Classe) And _
            tblPersoDons!Niveau = rst!Niveau
          .Cell(flexcpText, i, j + L1_COL_DON_1) = tblPersoDons!dons
          .Cell(flexcpText, i, j + L1_COL_LIB_1) = tblPersoDons!Libelle
          j = j + 1
          tblPersoDons.MoveNext
          If tblPersoDons.EOF Then Exit Do
        Loop
      End If
      .Cell(flexcpText, i, L1_COL_MODIF_INT) = rst!modifint
      i = i + 1
      rst.MoveNext
    Loop
    .Redraw = True
  End With
  tblPersoDons.Close

RempHistoPerso_Exit:
  Exit Sub

RempHistoPerso_Error:
  MsgBox "L'erreur suivante s'est produite : " & Err.Description & vbCrLf & vbCrLf & " numéro d'erreur (" & Err.Number & " )", vbCritical, g_strNomApplication
  Print #g_FileDebug, Format(Now, "dd/MM/yyyy hh:mm:ss") & " Erreur n° : (" & Err.Number & ") " & Err.Description & " Procedure : RempHistoPerso  Module : frmGestPerso.frm"
  Resume RempHistoPerso_Exit
End Sub
Private Sub Supprimer()

  If g_ModeDebug = vbUnchecked Then On Error GoTo Supprimer_Error

  If MsgBox("Confirmation suppression du personnage " & vsPerso.Cell(flexcpText, vsPerso.SelectedRow(0), L2_COL_NOM), vbExclamation Or vbOKCancel Or vbDefaultButton2, Me.Caption) = vbOK Then
    g_bdPerso.Execute "delete from Equipement where personnage='" & FormatStrSQL(vsPerso.Cell(flexcpText, vsPerso.SelectedRow(0))) & "'", dbFailOnError
    g_bdPerso.Execute "delete from PersonnageProgression where nom='" & FormatStrSQL(vsPerso.Cell(flexcpText, vsPerso.SelectedRow(0))) & "'", dbFailOnError
    g_bdPerso.Execute "delete from PersonnageDons where nom='" & FormatStrSQL(vsPerso.Cell(flexcpText, vsPerso.SelectedRow(0))) & "'", dbFailOnError
    g_bdPerso.Execute "delete from Personnage where nom='" & FormatStrSQL(vsPerso.Cell(flexcpText, vsPerso.SelectedRow(0))) & "'", dbFailOnError
    bChargePerso = True
    RempPerso
    Efface Me
    bChargePerso = False
  End If

Supprimer_Exit:
  Exit Sub

Supprimer_Error:
  MsgBox "L'erreur suivante s'est produite : " & Err.Description & vbCrLf & vbCrLf & " numéro d'erreur (" & Err.Number & " )", vbCritical, g_strNomApplication
  Print #g_FileDebug, Format(Now, "dd/MM/yyyy hh:mm:ss") & " Erreur n° : (" & Err.Number & ") " & Err.Description & " Procedure : Supprimer  Module : frmGestPerso.frm"
  Resume Supprimer_Exit
End Sub
Private Sub btnAjouter_Click()
  'ajout
  
  frmAjouter.Show vbModal
  If frmAjouter.m_type = 0 Then
    Exit Sub
  End If
  'Ferme
  
  ''Err.Clear
  
  If frmAjouter.m_type = 2 Then
    Efface Me
    vsHistoPerso.Rows = 1
    vsEquip.Rows = 1
    frmAssistant.Show  'vbModal
    Exit Sub
  End If
  Efface Me
  vsHistoPerso.Rows = 1
  vsEquip.Rows = 1
  CmbRace.Enabled = True
  CmbEcoleSpecialisation.ListIndex = 0
  fldstrnom.SetFocus
  BloqueCombo
End Sub

Private Sub btnAnnuler_Click()
  Fermer
End Sub

Private Sub btnBackground_Click()
  frmBackHist.Show vbModal
End Sub

Private Sub btnEnregistrer_Click()
  If Enregistre = True Then
    If frmMain.TrouveChild("frmAcceder") <> -1 Then
      frmAcceder.RempPerso
    End If
    
    If frmMain.TrouveChild("frmFichePerso") <> -1 Then
      frmFichePerso.RetrouveFiche
    End If
  End If
End Sub

Private Sub btnEquipEnlever_Click()
Dim i As Integer
Dim rst As Recordset

  If g_ModeDebug = vbUnchecked Then On Error GoTo btnEquipEnlever_Click_Error
  
  With vsEquip
    For i = 1 To .Rows - 1
      If .IsSelected(i) = True Then
        Set rst = g_bdPerso.OpenRecordset("select CoutTotal,CHARGE from objets where nomObjet='" & FormatStrSQL(.Cell(flexcpText, i, L3_COL_EQUIP_NOM_OBJET)) & "'", dbOpenDynaset)
        If Not rst.EOF Then
          frmVendre.m_PrixUnitaire = IIf(IsNull(rst!CoutTotal), 0, rst!CoutTotal) * maxi_(1, Val(.Cell(flexcpText, i, L3_COL_EQUIP_CHARGE))) / maxi_(1, rst!CHARGE)
        Else
          frmVendre.m_PrixUnitaire = 0
        End If
        rst.Close
        frmVendre.m_NombreObjet = Val(.Cell(flexcpText, i, L3_COL_EQUIP_MULTIPLE))
        frmVendre.m_NomObjet = .Cell(flexcpText, i, L3_COL_EQUIP_NOM_OBJET)
        frmVendre.Show vbModal
        
        Exit For
      End If
    Next i
  End With
  Modifperso

btnEquipEnlever_Click_Exit:
  Exit Sub

btnEquipEnlever_Click_Error:
  MsgBox "L'erreur suivante s'est produite : " & Err.Description & vbCrLf & vbCrLf & " numéro d'erreur (" & Err.Number & " )", vbCritical, g_strNomApplication
  Print #g_FileDebug, Format(Now, "dd/MM/yyyy hh:mm:ss") & " Erreur n° : (" & Err.Number & ") " & Err.Description & " Procedure : btnEquipEnlever_Click  Module : frmGestPerso.frm"
  Resume btnEquipEnlever_Click_Exit
End Sub

Private Sub btnEquipInserer_Click(Index As Integer)

  If g_ModeDebug = vbUnchecked Then On Error GoTo btnEquipInserer_Click_Error
  vsEquip.Rows = vsEquip.Rows + 1
  Modifperso
  vsEquip.SetFocus
  
  vsEquip.Row = vsEquip.Rows - 1
  vsEquip.Col = L3_COL_EQUIP_TYPE
  FrmLstArticle.MultPrix = Index
  FrmLstArticle.Show vbModal
  
btnEquipInserer_Click_Exit:
  Exit Sub

btnEquipInserer_Click_Error:
  MsgBox "L'erreur suivante s'est produite : " & Err.Description & vbCrLf & vbCrLf & " numéro d'erreur (" & Err.Number & " )", vbCritical, g_strNomApplication
  Print #g_FileDebug, Format(Now, "dd/MM/yyyy hh:mm:ss") & " Erreur n° : (" & Err.Number & ") " & Err.Description & " Procedure : btnEquipInserer_Click  Module : frmGestPerso.frm"
  Resume btnEquipInserer_Click_Exit
End Sub

Private Sub btnFermer_Click()
  Fermer
End Sub

Private Sub btnnivEnlever_Click()
Dim i As Integer

  If g_ModeDebug = vbUnchecked Then On Error GoTo btnnivEnlever_Click_Error

  
  With vsHistoPerso
    For i = 1 To .Rows - 1
      If .IsSelected(i) = True Then
        .RemoveItem i
        ChkDonsAPrendre_Click
        Exit For
      End If
    Next i
  End With
  Modifperso

btnnivEnlever_Click_Exit:
  Exit Sub

btnnivEnlever_Click_Error:
  MsgBox "L'erreur suivante s'est produite : " & Err.Description & vbCrLf & vbCrLf & " numéro d'erreur (" & Err.Number & " )", vbCritical, g_strNomApplication
  Print #g_FileDebug, Format(Now, "dd/MM/yyyy hh:mm:ss") & " Erreur n° : (" & Err.Number & ") " & Err.Description & " Procedure : btnnivEnlever_Click  Module : frmGestPerso.frm"
  Resume btnnivEnlever_Click_Exit
End Sub

Private Sub btnnivInserer_Click()
Dim rst As Recordset

  If g_ModeDebug = vbUnchecked Then On Error GoTo btnnivInserer_Click_Error
  
  vsHistoPerso.Rows = vsHistoPerso.Rows + 1
  If vsHistoPerso.Rows > 2 Then
    vsHistoPerso.TextMatrix(vsHistoPerso.Rows - 1, L1_COL_CLASSE) = vsHistoPerso.TextMatrix(vsHistoPerso.Rows - 2, L1_COL_CLASSE)
    vsHistoPerso.TextMatrix(vsHistoPerso.Rows - 1, L1_COL_NIVEAU) = vsHistoPerso.TextMatrix(vsHistoPerso.Rows - 2, L1_COL_NIVEAU) + 1
    Set rst = g_bdRoles.OpenRecordset("Classe", dbOpenTable)
    rst.Index = "PrimaryKey"
    rst.Seek "=", vsHistoPerso.TextMatrix(vsHistoPerso.Rows - 2, L1_COL_CLASSE)
    If Not rst.NoMatch Then
      If Not EstMortVivant(Cmbarchetype) Then
        vsHistoPerso.TextMatrix(vsHistoPerso.Rows - 1, L1_COL_PV) = Int(rst!PointDeVie * Rnd(1000)) + 1
      Else
        vsHistoPerso.TextMatrix(vsHistoPerso.Rows - 1, L1_COL_PV) = Int(12 * Rnd(1000)) + 1
      End If
    End If
    rst.Close
  End If
  ' A VOIR
  vsHistoPerso.Cell(flexcpText, vsHistoPerso.Rows - 1, L1_COL_MODIF_INT) = 0
  
  Modifperso
  ChkDonsAPrendre_Click
  vsHistoPerso.SetFocus

btnnivInserer_Click_Exit:
  Exit Sub

btnnivInserer_Click_Error:
  MsgBox "L'erreur suivante s'est produite : " & Err.Description & vbCrLf & vbCrLf & " numéro d'erreur (" & Err.Number & " )", vbCritical, g_strNomApplication
  Print #g_FileDebug, Format(Now, "dd/MM/yyyy hh:mm:ss") & " Erreur n° : (" & Err.Number & ") " & Err.Description & " Procedure : btnnivInserer_Click  Module : frmGestPerso.frm"
  Resume btnnivInserer_Click_Exit
End Sub


Private Sub btnRedescendre_Click()
Dim i As Integer, n As Integer
Dim strinter As String

  If g_ModeDebug = vbUnchecked Then On Error GoTo btnRedescendre_Click_Error
 
  With vsHistoPerso
    For i = 1 To .Rows - 1
      If .IsSelected(i) Then
        If i < .Rows - 1 Then
          For n = 0 To .Cols - 1
            strinter = .Cell(flexcpText, i + 1, n)
            .Cell(flexcpText, i + 1, n) = .Cell(flexcpText, i, n)
            .Cell(flexcpText, i, n) = strinter
          Next n
          .IsSelected(i) = False
          .IsSelected(i + 1) = True
          Modifperso
        End If
        Exit For
      End If
    Next i
  End With

btnRedescendre_Click_Exit:
  Exit Sub

btnRedescendre_Click_Error:
  MsgBox "L'erreur suivante s'est produite : " & Err.Description & vbCrLf & vbCrLf & " numéro d'erreur (" & Err.Number & " )", vbCritical, g_strNomApplication
  Print #g_FileDebug, Format(Now, "dd/MM/yyyy hh:mm:ss") & " Erreur n° : (" & Err.Number & ") " & Err.Description & " Procedure : btnRedescendre_Click  Module : frmGestPerso.frm"
  Resume btnRedescendre_Click_Exit
End Sub

Private Sub btnRemonter_Click()
Dim i As Integer, n As Integer
Dim strinter As String

  If g_ModeDebug = vbUnchecked Then On Error GoTo btnRemonter_Click_Error
 
  With vsHistoPerso
    For i = 1 To .Rows - 1
      If .IsSelected(i) Then
        If i > 1 Then
          For n = 0 To .Cols - 1
            strinter = .Cell(flexcpText, i - 1, n)
            .Cell(flexcpText, i - 1, n) = .Cell(flexcpText, i, n)
            .Cell(flexcpText, i, n) = strinter
          Next n
          .IsSelected(i) = False
          .IsSelected(i - 1) = True
          Modifperso
        End If
        Exit For
      End If
    Next i
  End With

btnRemonter_Click_Exit:
  Exit Sub

btnRemonter_Click_Error:
  MsgBox "L'erreur suivante s'est produite : " & Err.Description & vbCrLf & vbCrLf & " numéro d'erreur (" & Err.Number & " )", vbCritical, g_strNomApplication
  Print #g_FileDebug, Format(Now, "dd/MM/yyyy hh:mm:ss") & " Erreur n° : (" & Err.Number & ") " & Err.Description & " Procedure : btnRemonter_Click  Module : frmGestPerso.frm"
  Resume btnRemonter_Click_Exit
End Sub

Private Sub btnRenommer_Click()
Dim n As Integer
Dim rst As Recordset
  
  If vsPerso.SelectedRows < 1 Then
    Exit Sub
  End If
  
  n = vsPerso.SelectedRow(0)
  frmRenommer.fldstrNouveauNom = vsPerso.TextMatrix(n, L2_COL_NOM)
  frmRenommer.fldstrNouveauNom.ToolTipText = "Permet de renommer le personnage, (retour en arriere possible)"
  frmRenommer.Show vbModal
  If frmRenommer.m_strnouveaunom <> "" Then
    If UCase(frmRenommer.m_strnouveaunom) <> UCase(vsPerso.TextMatrix(n, L2_COL_NOM)) Then
      Set rst = g_bdPerso.OpenRecordset("select * from personnage where nom='" & FormatStrSQL(frmRenommer.m_strnouveaunom) & "'", dbOpenDynaset)
      If Not rst.EOF Then
        MsgBox "Impossible de renommer vu que ce nouveau nom de personnage existe déjà", vbExclamation, g_strNomApplication
        rst.Close
        Exit Sub
      Else
        rst.Close
      End If
    End If
    
    bChargePerso = True
    g_bdPerso.Execute "update personnage set nom='" & FormatStrSQL(frmRenommer.m_strnouveaunom) & "' where nom='" & FormatStrSQL(vsPerso.TextMatrix(n, L2_COL_NOM)) & "'", dbFailOnError
    g_bdPerso.Execute "update personnageDons set nom='" & FormatStrSQL(frmRenommer.m_strnouveaunom) & "' where nom='" & FormatStrSQL(vsPerso.TextMatrix(n, L2_COL_NOM)) & "'", dbFailOnError
    g_bdPerso.Execute "update personnageProgression set nom='" & FormatStrSQL(frmRenommer.m_strnouveaunom) & "' where nom='" & FormatStrSQL(vsPerso.TextMatrix(n, L2_COL_NOM)) & "'", dbFailOnError
    g_bdPerso.Execute "update SortPersonnage set nomperso='" & FormatStrSQL(frmRenommer.m_strnouveaunom) & "' where nomperso='" & FormatStrSQL(vsPerso.TextMatrix(n, L2_COL_NOM)) & "'", dbFailOnError
    vsPerso.TextMatrix(n, L2_COL_NOM) = frmRenommer.m_strnouveaunom
    fldstrnom = frmRenommer.m_strnouveaunom
    
    If frmMain.TrouveChild("frmAcceder") <> -1 Then
      frmAcceder.RempPerso
    End If
    
    If frmMain.TrouveChild("frmFichePerso") <> -1 Then
      Unload frmFichePerso
    End If
    bChargePerso = False
  End If
End Sub

Private Sub btnSupprimer_Click()
  Supprimer
End Sub

Private Sub ChkDonsAPrendre_Click()
Dim i As Integer, c As Integer, j As Integer
Dim backcolor As Long
Dim nDons As Integer
Dim nDonsDu As Integer
Dim nDonsPris As Integer
Dim tblClasseprogress As Recordset

Const MANQUE_DONS = 1
Const A_PRIS_DONS = 2


  Set tblClasseprogress = g_bdRoles.OpenRecordset("ClasseProgression", dbOpenTable)
  tblClasseprogress.Index = "index"
  
  With vsHistoPerso
    For i = 1 To .Rows - 1
      nDons = 0
      nDonsDu = 0
      nDonsPris = 0
      If ChkDonsAPrendre = vbChecked Or ChkDonsObtenus = vbChecked Then
        ''decompte le nombre de dons pris
        For c = L1_COL_DON_1 To L1_COL_DON_10
          If Trim(.TextMatrix(i, c)) <> "" Then nDonsPris = nDonsPris + 1
        Next c
        ''decompte le nombre de dons à prendre
        tblClasseprogress.Seek "=", .TextMatrix(i, L1_COL_CLASSE), .TextMatrix(i, L1_COL_NIVEAU)
        If Not tblClasseprogress.NoMatch Then
          For j = 1 To 6
            If Dons_A_Selectionner(IIf(IsNull(tblClasseprogress.Fields("don_" & j)), "", tblClasseprogress.Fields("don_" & j)), "") Then
              nDonsDu = nDonsDu + 1
            End If
          Next j
        End If
        
        If (i Mod 3) = 0 Or i = 1 Then
          nDonsDu = nDonsDu + 1
        End If
        
        If ChkDonsAPrendre = vbChecked Then
          If nDonsPris < nDonsDu Then nDons = nDons + MANQUE_DONS
        End If
        
        If ChkDonsObtenus = vbChecked Then
          If nDonsPris > 0 Then nDons = nDons + A_PRIS_DONS
        End If
      End If
      
      If nDons > A_PRIS_DONS Then
        backcolor = &H400040
      ElseIf nDons = A_PRIS_DONS Then
        backcolor = &HFF&
      ElseIf nDons = MANQUE_DONS Then
        backcolor = &HFF0000
      Else
        backcolor = &HFFFFFF
      End If
      For c = 0 To L1_COL_DON_COMP
        .Cell(flexcpBackColor, i, c) = backcolor
      Next c
    Next i
  End With

End Sub

Private Sub ChkDonsObtenus_Click()
  ChkDonsAPrendre_Click

End Sub

Private Sub cmbAlignement_Click()
  Modifperso
End Sub
Private Sub CmbElu_Click(Index As Integer)
  Modifperso
End Sub

Private Sub CmbShugenja_Click(Index As Integer)
 Modifperso
End Sub

Private Sub CmbSorcier_Click(Index As Integer)
  Modifperso
End Sub
Private Sub CmbWujen_Click()
  Modifperso
End Sub
Private Sub CmbDragonTotem_Click()
  Modifperso
End Sub
Private Sub Cmbarchetype_Click()
  Modifperso
End Sub

Private Sub fldnTotalXP_Click()
  Modifperso
End Sub
Private Sub fldnTotalPo_Click()
  Modifperso
End Sub
Private Sub CmbArtisanat_Click(Index As Integer)
  Modifperso
End Sub
Private Sub CmbArtisanat_change(Index As Integer)
  Modifperso
End Sub
Private Sub CmbArtisanat_GotFocus(Index As Integer)
Selec
End Sub

Private Sub CmbProfession_Click(Index As Integer)
  Modifperso
End Sub
Private Sub CmbProfession_Change(Index As Integer)
  Modifperso
End Sub
Private Sub CmbProfession_GotFocus(Index As Integer)
Selec
End Sub

Private Sub cmbDomaine_Click(Index As Integer)
  Modifperso
End Sub
Private Sub cmbEcole_interdite_Click(Index As Integer)
  Modifperso
End Sub
Private Sub CmbRace_click()
  Modifperso
End Sub
Private Sub CmbRace_Change()

  If g_ModeDebug = vbUnchecked Then On Error GoTo CmbRace_Change_Error
  Modifperso

CmbRace_Change_Exit:
  Exit Sub

CmbRace_Change_Error:
  MsgBox "L'erreur suivante s'est produite : " & Err.Description & vbCrLf & vbCrLf & " numéro d'erreur (" & Err.Number & " )", vbCritical, g_strNomApplication
  Print #g_FileDebug, Format(Now, "dd/MM/yyyy hh:mm:ss") & " Erreur n° : (" & Err.Number & ") " & Err.Description & " Procedure : CmbRace_Change  Module : frmGestPerso.frm"
  Resume CmbRace_Change_Exit
End Sub
Private Sub CmbRace_GotFocus()

  If g_ModeDebug = vbUnchecked Then On Error GoTo CmbRace_GotFocus_Error
  Selec

CmbRace_GotFocus_Exit:
  Exit Sub

CmbRace_GotFocus_Error:
  MsgBox "L'erreur suivante s'est produite : " & Err.Description & vbCrLf & vbCrLf & " numéro d'erreur (" & Err.Number & " )", vbCritical, g_strNomApplication
  Print #g_FileDebug, Format(Now, "dd/MM/yyyy hh:mm:ss") & " Erreur n° : (" & Err.Number & ") " & Err.Description & " Procedure : CmbRace_GotFocus  Module : frmGestPerso.frm"
  Resume CmbRace_GotFocus_Exit
End Sub
Private Sub CmbEcoleSpecialisation_Change()
  Modifperso
End Sub

Private Sub CmbEcoleSpecialisation_Click()
  Modifperso
End Sub

Private Sub flddate_Change()

  If g_ModeDebug = vbUnchecked Then On Error GoTo flddate_Change_Error
  Modifperso

flddate_Change_Exit:
  Exit Sub

flddate_Change_Error:
  MsgBox "L'erreur suivante s'est produite : " & Err.Description & vbCrLf & vbCrLf & " numéro d'erreur (" & Err.Number & " )", vbCritical, g_strNomApplication
  Print #g_FileDebug, Format(Now, "dd/MM/yyyy hh:mm:ss") & " Erreur n° : (" & Err.Number & ") " & Err.Description & " Procedure : flddate_Change  Module : frmGestPerso.frm"
  Resume flddate_Change_Exit
End Sub

Private Sub flddate_GotFocus()

  If g_ModeDebug = vbUnchecked Then On Error GoTo flddate_GotFocus_Error
  Selec


flddate_GotFocus_Exit:
  Exit Sub

flddate_GotFocus_Error:
  MsgBox "L'erreur suivante s'est produite : " & Err.Description & vbCrLf & vbCrLf & " numéro d'erreur (" & Err.Number & " )", vbCritical, g_strNomApplication
  Print #g_FileDebug, Format(Now, "dd/MM/yyyy hh:mm:ss") & " Erreur n° : (" & Err.Number & ") " & Err.Description & " Procedure : flddate_GotFocus  Module : frmGestPerso.frm"
  Resume flddate_GotFocus_Exit
End Sub

Private Sub fldnAge_Change()

  If g_ModeDebug = vbUnchecked Then On Error GoTo fldnAge_Change_Error
  Modifperso

fldnAge_Change_Exit:
  Exit Sub

fldnAge_Change_Error:
  MsgBox "L'erreur suivante s'est produite : " & Err.Description & vbCrLf & vbCrLf & " numéro d'erreur (" & Err.Number & " )", vbCritical, g_strNomApplication
  Print #g_FileDebug, Format(Now, "dd/MM/yyyy hh:mm:ss") & " Erreur n° : (" & Err.Number & ") " & Err.Description & " Procedure : fldnAge_Change  Module : frmGestPerso.frm"
  Resume fldnAge_Change_Exit
End Sub

Private Sub fldnAge_GotFocus()

  If g_ModeDebug = vbUnchecked Then On Error GoTo fldnAge_GotFocus_Error
  Selec


fldnAge_GotFocus_Exit:
  Exit Sub

fldnAge_GotFocus_Error:
  MsgBox "L'erreur suivante s'est produite : " & Err.Description & vbCrLf & vbCrLf & " numéro d'erreur (" & Err.Number & " )", vbCritical, g_strNomApplication
  Print #g_FileDebug, Format(Now, "dd/MM/yyyy hh:mm:ss") & " Erreur n° : (" & Err.Number & ") " & Err.Description & " Procedure : fldnAge_GotFocus  Module : frmGestPerso.frm"
  Resume fldnAge_GotFocus_Exit
End Sub

Private Sub fldnbeaute_Change()

  If g_ModeDebug = vbUnchecked Then On Error GoTo fldnbeaute_Change_Error
  Modifperso

fldnbeaute_Change_Exit:
  Exit Sub

fldnbeaute_Change_Error:
  MsgBox "L'erreur suivante s'est produite : " & Err.Description & vbCrLf & vbCrLf & " numéro d'erreur (" & Err.Number & " )", vbCritical, g_strNomApplication
  Print #g_FileDebug, Format(Now, "dd/MM/yyyy hh:mm:ss") & " Erreur n° : (" & Err.Number & ") " & Err.Description & " Procedure : fldnbeaute_Change  Module : frmGestPerso.frm"
  Resume fldnbeaute_Change_Exit
End Sub

Private Sub fldnbeaute_GotFocus()

  If g_ModeDebug = vbUnchecked Then On Error GoTo fldnbeaute_GotFocus_Error
  Selec


fldnbeaute_GotFocus_Exit:
  Exit Sub

fldnbeaute_GotFocus_Error:
  MsgBox "L'erreur suivante s'est produite : " & Err.Description & vbCrLf & vbCrLf & " numéro d'erreur (" & Err.Number & " )", vbCritical, g_strNomApplication
  Print #g_FileDebug, Format(Now, "dd/MM/yyyy hh:mm:ss") & " Erreur n° : (" & Err.Number & ") " & Err.Description & " Procedure : fldnbeaute_GotFocus  Module : frmGestPerso.frm"
  Resume fldnbeaute_GotFocus_Exit
End Sub


Private Sub fldnMalusXP_Change()

  If g_ModeDebug = vbUnchecked Then On Error GoTo fldnMalusXP_Change_Error
  Modifperso

fldnMalusXP_Change_Exit:
  Exit Sub

fldnMalusXP_Change_Error:
  MsgBox "L'erreur suivante s'est produite : " & Err.Description & vbCrLf & vbCrLf & " numéro d'erreur (" & Err.Number & " )", vbCritical, g_strNomApplication
  Print #g_FileDebug, Format(Now, "dd/MM/yyyy hh:mm:ss") & " Erreur n° : (" & Err.Number & ") " & Err.Description & " Procedure : fldnMalusXP_Change  Module : frmGestPerso.frm"
  Resume fldnMalusXP_Change_Exit
End Sub

Private Sub fldnMalusXP_GotFocus()

  If g_ModeDebug = vbUnchecked Then On Error GoTo fldnMalusXP_GotFocus_Error
  Selec


fldnMalusXP_GotFocus_Exit:
  Exit Sub

fldnMalusXP_GotFocus_Error:
  MsgBox "L'erreur suivante s'est produite : " & Err.Description & vbCrLf & vbCrLf & " numéro d'erreur (" & Err.Number & " )", vbCritical, g_strNomApplication
  Print #g_FileDebug, Format(Now, "dd/MM/yyyy hh:mm:ss") & " Erreur n° : (" & Err.Number & ") " & Err.Description & " Procedure : fldnMalusXP_GotFocus  Module : frmGestPerso.frm"
  Resume fldnMalusXP_GotFocus_Exit
End Sub

Private Sub fldnPoids_Change()

  If g_ModeDebug = vbUnchecked Then On Error GoTo fldnPoids_Change_Error
  Modifperso

fldnPoids_Change_Exit:
  Exit Sub

fldnPoids_Change_Error:
  MsgBox "L'erreur suivante s'est produite : " & Err.Description & vbCrLf & vbCrLf & " numéro d'erreur (" & Err.Number & " )", vbCritical, g_strNomApplication
  Print #g_FileDebug, Format(Now, "dd/MM/yyyy hh:mm:ss") & " Erreur n° : (" & Err.Number & ") " & Err.Description & " Procedure : fldnPoids_Change  Module : frmGestPerso.frm"
  Resume fldnPoids_Change_Exit
End Sub

Private Sub fldnPoids_GotFocus()

  If g_ModeDebug = vbUnchecked Then On Error GoTo fldnPoids_GotFocus_Error
  Selec


fldnPoids_GotFocus_Exit:
  Exit Sub

fldnPoids_GotFocus_Error:
  MsgBox "L'erreur suivante s'est produite : " & Err.Description & vbCrLf & vbCrLf & " numéro d'erreur (" & Err.Number & " )", vbCritical, g_strNomApplication
  Print #g_FileDebug, Format(Now, "dd/MM/yyyy hh:mm:ss") & " Erreur n° : (" & Err.Number & ") " & Err.Description & " Procedure : fldnPoids_GotFocus  Module : frmGestPerso.frm"
  Resume fldnPoids_GotFocus_Exit
End Sub

Private Sub fldnTaille_Change()

  If g_ModeDebug = vbUnchecked Then On Error GoTo fldnTaille_Change_Error
  Modifperso

fldnTaille_Change_Exit:
  Exit Sub

fldnTaille_Change_Error:
  MsgBox "L'erreur suivante s'est produite : " & Err.Description & vbCrLf & vbCrLf & " numéro d'erreur (" & Err.Number & " )", vbCritical, g_strNomApplication
  Print #g_FileDebug, Format(Now, "dd/MM/yyyy hh:mm:ss") & " Erreur n° : (" & Err.Number & ") " & Err.Description & " Procedure : fldnTaille_Change  Module : frmGestPerso.frm"
  Resume fldnTaille_Change_Exit
End Sub

Private Sub fldnTaille_GotFocus()

  If g_ModeDebug = vbUnchecked Then On Error GoTo fldnTaille_GotFocus_Error
  Selec


fldnTaille_GotFocus_Exit:
  Exit Sub

fldnTaille_GotFocus_Error:
  MsgBox "L'erreur suivante s'est produite : " & Err.Description & vbCrLf & vbCrLf & " numéro d'erreur (" & Err.Number & " )", vbCritical, g_strNomApplication
  Print #g_FileDebug, Format(Now, "dd/MM/yyyy hh:mm:ss") & " Erreur n° : (" & Err.Number & ") " & Err.Description & " Procedure : fldnTaille_GotFocus  Module : frmGestPerso.frm"
  Resume fldnTaille_GotFocus_Exit
End Sub



Private Sub fldnVitalite_Change()

  If g_ModeDebug = vbUnchecked Then On Error GoTo fldnVitalite_Change_Error
  Modifperso

fldnVitalite_Change_Exit:
  Exit Sub

fldnVitalite_Change_Error:
  MsgBox "L'erreur suivante s'est produite : " & Err.Description & vbCrLf & vbCrLf & " numéro d'erreur (" & Err.Number & " )", vbCritical, g_strNomApplication
  Print #g_FileDebug, Format(Now, "dd/MM/yyyy hh:mm:ss") & " Erreur n° : (" & Err.Number & ") " & Err.Description & " Procedure : fldnVitalite_Change  Module : frmGestPerso.frm"
  Resume fldnVitalite_Change_Exit
End Sub

Private Sub fldnVitalite_GotFocus()

  If g_ModeDebug = vbUnchecked Then On Error GoTo fldnVitalite_GotFocus_Error
  Selec

fldnVitalite_GotFocus_Exit:
  Exit Sub

fldnVitalite_GotFocus_Error:
  MsgBox "L'erreur suivante s'est produite : " & Err.Description & vbCrLf & vbCrLf & " numéro d'erreur (" & Err.Number & " )", vbCritical, g_strNomApplication
  Print #g_FileDebug, Format(Now, "dd/MM/yyyy hh:mm:ss") & " Erreur n° : (" & Err.Number & ") " & Err.Description & " Procedure : fldnVitalite_GotFocus  Module : frmGestPerso.frm"
  Resume fldnVitalite_GotFocus_Exit
End Sub

Private Sub fldstrAlphabet_Change()
  Modifperso
End Sub

Private Sub fldstrcheveux_Change()

  If g_ModeDebug = vbUnchecked Then On Error GoTo fldstrcheveux_Change_Error
  Modifperso

fldstrcheveux_Change_Exit:
  Exit Sub

fldstrcheveux_Change_Error:
  MsgBox "L'erreur suivante s'est produite : " & Err.Description & vbCrLf & vbCrLf & " numéro d'erreur (" & Err.Number & " )", vbCritical, g_strNomApplication
  Print #g_FileDebug, Format(Now, "dd/MM/yyyy hh:mm:ss") & " Erreur n° : (" & Err.Number & ") " & Err.Description & " Procedure : fldstrcheveux_Change  Module : frmGestPerso.frm"
  Resume fldstrcheveux_Change_Exit
End Sub

Private Sub fldstrcheveux_GotFocus()

  If g_ModeDebug = vbUnchecked Then On Error GoTo fldstrcheveux_GotFocus_Error
  Selec


fldstrcheveux_GotFocus_Exit:
  Exit Sub

fldstrcheveux_GotFocus_Error:
  MsgBox "L'erreur suivante s'est produite : " & Err.Description & vbCrLf & vbCrLf & " numéro d'erreur (" & Err.Number & " )", vbCritical, g_strNomApplication
  Print #g_FileDebug, Format(Now, "dd/MM/yyyy hh:mm:ss") & " Erreur n° : (" & Err.Number & ") " & Err.Description & " Procedure : fldstrcheveux_GotFocus  Module : frmGestPerso.frm"
  Resume fldstrcheveux_GotFocus_Exit
End Sub

Private Sub fldstrdieu_Change()

  If g_ModeDebug = vbUnchecked Then On Error GoTo fldstrdieu_Change_Error
  Modifperso

fldstrdieu_Change_Exit:
  Exit Sub

fldstrdieu_Change_Error:
  MsgBox "L'erreur suivante s'est produite : " & Err.Description & vbCrLf & vbCrLf & " numéro d'erreur (" & Err.Number & " )", vbCritical, g_strNomApplication
  Print #g_FileDebug, Format(Now, "dd/MM/yyyy hh:mm:ss") & " Erreur n° : (" & Err.Number & ") " & Err.Description & " Procedure : fldstrdieu_Change  Module : frmGestPerso.frm"
  Resume fldstrdieu_Change_Exit
End Sub

Private Sub fldstrdieu_GotFocus()

  If g_ModeDebug = vbUnchecked Then On Error GoTo fldstrdieu_GotFocus_Error
  Selec


fldstrdieu_GotFocus_Exit:
  Exit Sub

fldstrdieu_GotFocus_Error:
  MsgBox "L'erreur suivante s'est produite : " & Err.Description & vbCrLf & vbCrLf & " numéro d'erreur (" & Err.Number & " )", vbCritical, g_strNomApplication
  Print #g_FileDebug, Format(Now, "dd/MM/yyyy hh:mm:ss") & " Erreur n° : (" & Err.Number & ") " & Err.Description & " Procedure : fldstrdieu_GotFocus  Module : frmGestPerso.frm"
  Resume fldstrdieu_GotFocus_Exit
End Sub

Private Sub fldstrLangueConnue_Change()
  Modifperso
End Sub

Private Sub fldstrnom_Change()

  Modifperso

End Sub

Private Sub fldstrnom_GotFocus()

  If g_ModeDebug = vbUnchecked Then On Error GoTo fldstrnom_GotFocus_Error
  Selec

fldstrnom_GotFocus_Exit:
  Exit Sub

fldstrnom_GotFocus_Error:
  MsgBox "L'erreur suivante s'est produite : " & Err.Description & vbCrLf & vbCrLf & " numéro d'erreur (" & Err.Number & " )", vbCritical, g_strNomApplication
  Print #g_FileDebug, Format(Now, "dd/MM/yyyy hh:mm:ss") & " Erreur n° : (" & Err.Number & ") " & Err.Description & " Procedure : fldstrnom_GotFocus  Module : frmGestPerso.frm"
  Resume fldstrnom_GotFocus_Exit
End Sub

Private Sub fldstrprofil_Change()

  If g_ModeDebug = vbUnchecked Then On Error GoTo fldstrprofil_Change_Error
  Modifperso

fldstrprofil_Change_Exit:
  Exit Sub

fldstrprofil_Change_Error:
  MsgBox "L'erreur suivante s'est produite : " & Err.Description & vbCrLf & vbCrLf & " numéro d'erreur (" & Err.Number & " )", vbCritical, g_strNomApplication
  Print #g_FileDebug, Format(Now, "dd/MM/yyyy hh:mm:ss") & " Erreur n° : (" & Err.Number & ") " & Err.Description & " Procedure : fldstrprofil_Change  Module : frmGestPerso.frm"
  Resume fldstrprofil_Change_Exit
End Sub

Private Sub fldstrprofil_GotFocus()

  If g_ModeDebug = vbUnchecked Then On Error GoTo fldstrprofil_GotFocus_Error
  Selec

fldstrprofil_GotFocus_Exit:
  Exit Sub

fldstrprofil_GotFocus_Error:
  MsgBox "L'erreur suivante s'est produite : " & Err.Description & vbCrLf & vbCrLf & " numéro d'erreur (" & Err.Number & " )", vbCritical, g_strNomApplication
  Print #g_FileDebug, Format(Now, "dd/MM/yyyy hh:mm:ss") & " Erreur n° : (" & Err.Number & ") " & Err.Description & " Procedure : fldstrprofil_GotFocus  Module : frmGestPerso.frm"
  Resume fldstrprofil_GotFocus_Exit
End Sub

Private Sub fldstrSexe_Change()

  If g_ModeDebug = vbUnchecked Then On Error GoTo fldstrSexe_Change_Error
  Modifperso

fldstrSexe_Change_Exit:
  Exit Sub

fldstrSexe_Change_Error:
  MsgBox "L'erreur suivante s'est produite : " & Err.Description & vbCrLf & vbCrLf & " numéro d'erreur (" & Err.Number & " )", vbCritical, g_strNomApplication
  Print #g_FileDebug, Format(Now, "dd/MM/yyyy hh:mm:ss") & " Erreur n° : (" & Err.Number & ") " & Err.Description & " Procedure : fldstrSexe_Change  Module : frmGestPerso.frm"
  Resume fldstrSexe_Change_Exit
End Sub

Private Sub fldstrSexe_GotFocus()

  If g_ModeDebug = vbUnchecked Then On Error GoTo fldstrSexe_GotFocus_Error
  Selec


fldstrSexe_GotFocus_Exit:
  Exit Sub

fldstrSexe_GotFocus_Error:
  MsgBox "L'erreur suivante s'est produite : " & Err.Description & vbCrLf & vbCrLf & " numéro d'erreur (" & Err.Number & " )", vbCritical, g_strNomApplication
  Print #g_FileDebug, Format(Now, "dd/MM/yyyy hh:mm:ss") & " Erreur n° : (" & Err.Number & ") " & Err.Description & " Procedure : fldstrSexe_GotFocus  Module : frmGestPerso.frm"
  Resume fldstrSexe_GotFocus_Exit
End Sub

Private Sub fldstrtitre_Change()

  If g_ModeDebug = vbUnchecked Then On Error GoTo fldstrtitre_Change_Error
  Modifperso

fldstrtitre_Change_Exit:
  Exit Sub

fldstrtitre_Change_Error:
  MsgBox "L'erreur suivante s'est produite : " & Err.Description & vbCrLf & vbCrLf & " numéro d'erreur (" & Err.Number & " )", vbCritical, g_strNomApplication
  Print #g_FileDebug, Format(Now, "dd/MM/yyyy hh:mm:ss") & " Erreur n° : (" & Err.Number & ") " & Err.Description & " Procedure : fldstrtitre_Change  Module : frmGestPerso.frm"
  Resume fldstrtitre_Change_Exit
End Sub

Private Sub fldstrtitre_GotFocus()

  If g_ModeDebug = vbUnchecked Then On Error GoTo fldstrtitre_GotFocus_Error
  Selec


fldstrtitre_GotFocus_Exit:
  Exit Sub

fldstrtitre_GotFocus_Error:
  MsgBox "L'erreur suivante s'est produite : " & Err.Description & vbCrLf & vbCrLf & " numéro d'erreur (" & Err.Number & " )", vbCritical, g_strNomApplication
  Print #g_FileDebug, Format(Now, "dd/MM/yyyy hh:mm:ss") & " Erreur n° : (" & Err.Number & ") " & Err.Description & " Procedure : fldstrtitre_GotFocus  Module : frmGestPerso.frm"
  Resume fldstrtitre_GotFocus_Exit
End Sub

Private Sub fldstryeux_Change()

  If g_ModeDebug = vbUnchecked Then On Error GoTo fldstryeux_Change_Error
  Modifperso

fldstryeux_Change_Exit:
  Exit Sub

fldstryeux_Change_Error:
  MsgBox "L'erreur suivante s'est produite : " & Err.Description & vbCrLf & vbCrLf & " numéro d'erreur (" & Err.Number & " )", vbCritical, g_strNomApplication
  Print #g_FileDebug, Format(Now, "dd/MM/yyyy hh:mm:ss") & " Erreur n° : (" & Err.Number & ") " & Err.Description & " Procedure : fldstryeux_Change  Module : frmGestPerso.frm"
  Resume fldstryeux_Change_Exit
End Sub

Private Sub fldstryeux_GotFocus()

  If g_ModeDebug = vbUnchecked Then On Error GoTo fldstryeux_GotFocus_Error
  Selec


fldstryeux_GotFocus_Exit:
  Exit Sub

fldstryeux_GotFocus_Error:
  MsgBox "L'erreur suivante s'est produite : " & Err.Description & vbCrLf & vbCrLf & " numéro d'erreur (" & Err.Number & " )", vbCritical, g_strNomApplication
  Print #g_FileDebug, Format(Now, "dd/MM/yyyy hh:mm:ss") & " Erreur n° : (" & Err.Number & ") " & Err.Description & " Procedure : fldstryeux_GotFocus  Module : frmGestPerso.frm"
  Resume fldstryeux_GotFocus_Exit
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

  If g_ModeDebug = vbUnchecked Then On Error GoTo Form_KeyDown_Error
  Select Case KeyCode
    Case vbKeyF8
      If btnEquipInserer(0).Enabled = True Then
        btnEquipInserer_Click (0)
      End If
    Case vbKeyF7
      If btnNivInserer.Enabled = True Then btnnivInserer_Click
    Case vbKeyF5
      If btnEnregistrer.Visible = True Then btnEnregistrer_Click
    Case vbKeyReturn
      SendKeys "{TAB}"
    Case vbKeyEscape
      Fermer
    Case vbKeyF9
      If btnAjouter.Visible = True Then btnAjouter_Click
    Case vbKeyF10
      If btnSupprimer.Visible = True Then btnSupprimer_Click
  End Select

Form_KeyDown_Exit:
  Exit Sub

Form_KeyDown_Error:
  MsgBox "L'erreur suivante s'est produite : " & Err.Description & vbCrLf & vbCrLf & " numéro d'erreur (" & Err.Number & " )", vbCritical, g_strNomApplication
  Print #g_FileDebug, Format(Now, "dd/MM/yyyy hh:mm:ss") & " Erreur n° : (" & Err.Number & ") " & Err.Description & " Procedure : Form_KeyDown  Module : frmGestPerso.frm"
  Resume Form_KeyDown_Exit
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

  If g_ModeDebug = vbUnchecked Then On Error GoTo Form_KeyPress_Error
  
  Select Case KeyAscii
    Case vbKeyEscape, vbKeyReturn
      KeyAscii = 0
  End Select

Form_KeyPress_Exit:
  Exit Sub

Form_KeyPress_Error:
  MsgBox "L'erreur suivante s'est produite : " & Err.Description & vbCrLf & vbCrLf & " numéro d'erreur (" & Err.Number & " )", vbCritical, g_strNomApplication
  Print #g_FileDebug, Format(Now, "dd/MM/yyyy hh:mm:ss") & " Erreur n° : (" & Err.Number & ") " & Err.Description & " Procedure : Form_KeyPress  Module : frmGestPerso.frm"
  Resume Form_KeyPress_Exit
End Sub

Private Sub Form_Load()
Dim rst As Recordset
Dim i As Integer, n As Integer
Dim bFind As Boolean
Dim strcombolist As String
Dim nb As Integer
Dim tabDomaine() As String

  If g_ModeDebug = vbUnchecked Then On Error GoTo Form_Load_Error
  
  bRefrech = True
  
  ReDim m_tabButton(7)
  m_tabButton(0).Nom = "BtnFermer"
  m_tabButton(1).Nom = "BtnAjouter"
  m_tabButton(2).Nom = "BtnEnregistrer"
  m_tabButton(3).Nom = "BtnSupprimer"
  m_tabButton(4).Nom = "BtnAnnuler"
  m_tabButton(5).Nom = "btnBackground"
  m_tabButton(6).Nom = "btnRenommer"

  Me.left = POS_FENETRE(Me.name, "POS_X", 0)
  Me.top = POS_FENETRE(Me.name, "POS_Y", 0)
 
  ChkDonsAPrendre = vbUnchecked
  ChkDonsAPrendre.Visible = g_bMarquerProgressionDon
  ChkDonsObtenus = vbUnchecked
  ChkDonsObtenus.Visible = g_bMarquerProgressionDon
  
  bChargePerso = True
  ''race
  CmbRace.Clear
  Set rst = g_bdRoles.OpenRecordset("select * from race order by race", dbOpenDynaset)
  Do While Not rst.EOF
    CmbRace.AddItem rst!RACE
    CmbRace.ItemData(CmbRace.ListCount - 1) = rst![n°]
    rst.MoveNext
  Loop
  rst.Close
  If CmbRace.ListCount > 0 Then CmbRace.ListIndex = 0
  
  ''Archétype
  
  Cmbarchetype.Clear
  Cmbarchetype.AddItem ""
  Set rst = g_bdRoles.OpenRecordset("select * from archetype order by archetype", dbOpenDynaset)
  Do While Not rst.EOF
    Cmbarchetype.AddItem rst!Archetype
    rst.MoveNext
  Loop
  rst.Close
  If Cmbarchetype.ListCount > 0 Then Cmbarchetype.ListIndex = 0
  
  ''classe
  strcombolist = ""
  Set rst = g_bdRoles.OpenRecordset("select * from classe order by classe", dbOpenDynaset)
  Do While Not rst.EOF
    If Not rst!affichage Then
      strcombolist = strcombolist & rst!Classe & "|"
    End If
    rst.MoveNext
  Loop
  rst.Close
  vsHistoPerso.ColComboList(L1_COL_CLASSE) = strcombolist
  
  vsEquip.ColHidden(L3_COL_EQUIP_PLACE) = True
  vsEquip.ColComboList(L3_COL_EQUIP_SUR_PERSO) = "Oui|Non"
  
  ''ecole de specialisation
  With CmbEcoleSpecialisation
    .Clear
    .AddItem ""
    .AddItem obj_GENERALISTE
    .AddItem obj_EVOCATION
    .AddItem obj_INVOCATION
    .AddItem obj_NECROMANCIE
    .AddItem obj_ILLUSION
    .AddItem obj_DIVINATION
    .AddItem obj_TRANSMUTATION
    .AddItem obj_ABJURATION
    .AddItem obj_ENCHANTEMENT
  End With
  
    ''ecole interdite
  For n = 0 To 3
    With CmbEcole_interdite(n)
      .Clear
      .AddItem ""
      .AddItem obj_EVOCATION
      .AddItem obj_INVOCATION
      .AddItem obj_NECROMANCIE
      .AddItem obj_ILLUSION
      .AddItem obj_TRANSMUTATION
      .AddItem obj_ABJURATION
      .AddItem obj_ENCHANTEMENT
    End With
  Next n
  ''domaine en fonction de la table sort
  cmbDomaine(0).Clear
  cmbDomaine(0).AddItem ""
  cmbDomaine(1).Clear
  cmbDomaine(1).AddItem ""
  cmbDomaine(2).Clear
  cmbDomaine(2).AddItem ""
  cmbDomaine(3).Clear
  cmbDomaine(3).AddItem ""
  '' artisanat
  For i = 0 To 2
    With CmbArtisanat(i)
      .Clear
      .AddItem ""
      Set rst = g_bdRoles.OpenRecordset("select * from artisanat order by nom", dbOpenDynaset)
      Do While Not rst.EOF
        .AddItem rst!Nom
        rst.MoveNext
      Loop
      rst.Close
    End With
  Next i
  '' profession
  
  For i = 0 To 3
    With CmbProfession(i)
      .Clear
      Set rst = g_bdRoles.OpenRecordset("select * from profession order by nom", dbOpenDynaset)
      Do While Not rst.EOF
        .AddItem rst!Nom
        rst.MoveNext
      Loop
      rst.Close
      If .ListCount > 0 Then .ListIndex = 0
    End With
  Next i
    
    '' Energie Elu
  For i = 0 To 2
    With CmbElu(i)
      .Clear
      Set rst = g_bdRoles.OpenRecordset("select * from energie order by nom", dbOpenDynaset)
      Do While Not rst.EOF
        .AddItem rst!Nom
        rst.MoveNext
      Loop
      rst.Close
    End With
  Next i
      '' Energie sorcier
  For i = 0 To 1
    With CmbSorcier(i)
      .Clear
      Set rst = g_bdRoles.OpenRecordset("select * from energie order by nom", dbOpenDynaset)
      Do While Not rst.EOF
        .AddItem rst!Nom
        rst.MoveNext
      Loop
      rst.Close
    End With
  Next i
  '' Element wujen
    With CmbWujen
      .Clear
      Set rst = g_bdRoles.OpenRecordset("select * from element order by nom", dbOpenDynaset)
      Do While Not rst.EOF
        .AddItem rst!Nom
        rst.MoveNext
      Loop
      rst.Close
    End With
  '' Element et ordre shugenja
    With CmbShugenja(1)
      .Clear
      Set rst = g_bdRoles.OpenRecordset("select Nom from Ordre order by Nom", dbOpenDynaset)
      Do While Not rst.EOF
        .AddItem rst!Nom
        rst.MoveNext
      Loop
      rst.Close
    End With
    With CmbShugenja(0)
      .Clear
      .AddItem "Terre"
      .AddItem "Air"
      .AddItem "Feu"
      .AddItem "Eau"
    End With
  '' Dragon Totem
    With CmbDragonTotem
      .Clear
      Set rst = g_bdRoles.OpenRecordset("select nom from dragontotem order by nom", dbOpenDynaset)
      Do While Not rst.EOF
        .AddItem rst!Nom
        rst.MoveNext
      Loop
      rst.Close
    End With
  ''alignement possible
  With CmbAlignement
    .Clear
    .AddItem ""
    Set rst = g_bdRoles.OpenRecordset("select * from alignements order by ordre", dbOpenDynaset)
    Do While Not rst.EOF
      .AddItem rst!Nom
      rst.MoveNext
    Loop
    rst.Close
  End With
  
  nb = ExtraitDomaineSort(tabDomaine, False, "", "", "", "")
  For i = 0 To nb - 1
    For n = 0 To cmbDomaine.Count - 1
      cmbDomaine(n).AddItem tabDomaine(i)
    Next n
  Next i
  BloqueCombo
  
  RempPerso
  
  bChargePerso = False
  
  If vsPerso.Rows > 1 Then
    bFind = False
    If frmMain.TrouveChild("frmAcceder") <> -1 Then
      If frmAcceder.vsnom.SelectedRows > 0 Then
        If frmAcceder.vsnom.SelectedRow(0) < vsPerso.Rows Then
          vsPerso.IsSelected(frmAcceder.vsnom.SelectedRow(0)) = True
          bFind = True
        End If
      End If
    End If
    
    If bFind = False Then
      vsPerso.IsSelected(1) = True
    End If
  End If
  

Form_Load_Exit:
  Exit Sub

Form_Load_Error:
  MsgBox "L'erreur suivante s'est produite : " & Err.Description & vbCrLf & vbCrLf & " numéro d'erreur (" & Err.Number & " )", vbCritical, g_strNomApplication
  Print #g_FileDebug, Format(Now, "dd/MM/yyyy hh:mm:ss") & " Erreur n° : (" & Err.Number & ") " & Err.Description & " Procedure : Form_Load  Module : frmGestPerso.frm"
  Resume Form_Load_Exit
End Sub

Private Sub Form_Unload(Cancel As Integer)

  If g_ModeDebug = vbUnchecked Then On Error GoTo Form_Unload_Error
  POS_FENETRE Me.name, "POS_X", Me.left
  POS_FENETRE Me.name, "POS_Y", Me.top

Form_Unload_Exit:
  Exit Sub

Form_Unload_Error:
  MsgBox "L'erreur suivante s'est produite : " & Err.Description & vbCrLf & vbCrLf & " numéro d'erreur (" & Err.Number & " )", vbCritical, g_strNomApplication
  Print #g_FileDebug, Format(Now, "dd/MM/yyyy hh:mm:ss") & " Erreur n° : (" & Err.Number & ") " & Err.Description & " Procedure : Form_Unload  Module : frmGestPerso.frm"
  Resume Form_Unload_Exit
End Sub


Private Sub Timer1_Timer()
  AjusteBouton Me, m_tabButton
End Sub


Private Sub vsEquip_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)

  If g_ModeDebug = vbUnchecked Then On Error GoTo vsEquip_BeforeEdit_Error
  
  If Row > 0 And Col = L3_COL_EQUIP_TYPE Then
    vsEquip.ComboList = "..."
  Else
    vsEquip.ComboList = ""
  End If

  If Row > 0 And (Col = L3_COL_EQUIP_CHARGE Or Col = L3_COL_EQUIP_MULTIPLE Or Col = L3_COL_EQUIP_SUR_PERSO) Then
    Modifperso
    vsEquip.SetFocus
  End If

vsEquip_BeforeEdit_Exit:
  Exit Sub

vsEquip_BeforeEdit_Error:
  MsgBox "L'erreur suivante s'est produite : " & Err.Description & vbCrLf & vbCrLf & " numéro d'erreur (" & Err.Number & " )", vbCritical, g_strNomApplication
  Print #g_FileDebug, Format(Now, "dd/MM/yyyy hh:mm:ss") & " Erreur n° : (" & Err.Number & ") " & Err.Description & " Procedure : vsEquip_BeforeEdit  Module : frmGestPerso.frm"
  Resume vsEquip_BeforeEdit_Exit
End Sub

Private Sub vsEquip_CellButtonClick(ByVal Row As Long, ByVal Col As Long)

  If g_ModeDebug = vbUnchecked Then On Error GoTo vsEquip_CellButtonClick_Error
  If Row > 0 And Col = L3_COL_EQUIP_TYPE Then
    FrmLstArticle.Show vbModal
  End If

vsEquip_CellButtonClick_Exit:
  Exit Sub

vsEquip_CellButtonClick_Error:
  MsgBox "L'erreur suivante s'est produite : " & Err.Description & vbCrLf & vbCrLf & " numéro d'erreur (" & Err.Number & " )", vbCritical, g_strNomApplication
  Print #g_FileDebug, Format(Now, "dd/MM/yyyy hh:mm:ss") & " Erreur n° : (" & Err.Number & ") " & Err.Description & " Procedure : vsEquip_CellButtonClick  Module : frmGestPerso.frm"
  Resume vsEquip_CellButtonClick_Exit
End Sub

Private Sub vsEquip_Click()

  If g_ModeDebug = vbUnchecked Then On Error GoTo vsEquip_Click_Error
  btnEquipEnlever.Enabled = True

vsEquip_Click_Exit:
  Exit Sub

vsEquip_Click_Error:
  MsgBox "L'erreur suivante s'est produite : " & Err.Description & vbCrLf & vbCrLf & " numéro d'erreur (" & Err.Number & " )", vbCritical, g_strNomApplication
  Print #g_FileDebug, Format(Now, "dd/MM/yyyy hh:mm:ss") & " Erreur n° : (" & Err.Number & ") " & Err.Description & " Procedure : vsEquip_Click  Module : frmGestPerso.frm"
  Resume vsEquip_Click_Exit
End Sub

Private Sub vsHistoPerso_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)

  If g_ModeDebug = vbUnchecked Then On Error GoTo vsHistoPerso_BeforeEdit_Error
  
  If Row > 0 And Col = L1_COL_DON_COMP Then
    vsHistoPerso.ComboList = "..."
  Else
    vsHistoPerso.ComboList = ""
    If Col = 2 Then BloqueCombo
  End If
  ChkDonsAPrendre_Click

vsHistoPerso_BeforeEdit_Exit:
  Exit Sub

vsHistoPerso_BeforeEdit_Error:
  MsgBox "L'erreur suivante s'est produite : " & Err.Description & vbCrLf & vbCrLf & " numéro d'erreur (" & Err.Number & " )", vbCritical, g_strNomApplication
  Print #g_FileDebug, Format(Now, "dd/MM/yyyy hh:mm:ss") & " Erreur n° : (" & Err.Number & ") " & Err.Description & " Procedure : vsHistoPerso_BeforeEdit  Module : frmGestPerso.frm"
  Resume vsHistoPerso_BeforeEdit_Exit
End Sub
Private Sub vsHistoPerso_CellButtonClick(ByVal Row As Long, ByVal Col As Long)

  If g_ModeDebug = vbUnchecked Then On Error GoTo vsHistoPerso_CellButtonClick_Error
  
  If Row > 0 And Col = L1_COL_DON_COMP Then
    frmDonsComp.Niveau = vsHistoPerso.Cell(flexcpValue, Row, L1_COL_NIVEAU)
    frmDonsComp.NiveauGlobal = Row
    frmDonsComp.bCreation = False
    frmDonsComp.Show vbModal
    ChkDonsAPrendre_Click
  End If

vsHistoPerso_CellButtonClick_Exit:
  Exit Sub

vsHistoPerso_CellButtonClick_Error:
  MsgBox "L'erreur suivante s'est produite : " & Err.Description & vbCrLf & vbCrLf & " numéro d'erreur (" & Err.Number & " )", vbCritical, g_strNomApplication
  Print #g_FileDebug, Format(Now, "dd/MM/yyyy hh:mm:ss") & " Erreur n° : (" & Err.Number & ") " & Err.Description & " Procedure : vsHistoPerso_CellButtonClick  Module : frmGestPerso.frm"
  Resume vsHistoPerso_CellButtonClick_Exit
End Sub

Private Sub vsHistoPerso_Click()
Dim i As Integer

  i = vsHistoPerso.Row
  BloqueCombo


End Sub

Private Sub vsHistoPerso_SelChange()
Dim i As Integer

  If g_ModeDebug = vbUnchecked Then On Error GoTo vsHistoPerso_SelChange_Error
  
  For i = 1 To vsHistoPerso.Rows - 1
    If vsHistoPerso.IsSelected(i) = True Then
      btnNivEnlever.Enabled = True
      Exit For
    End If
  Next i

vsHistoPerso_SelChange_Exit:
  Exit Sub

vsHistoPerso_SelChange_Error:
  MsgBox "L'erreur suivante s'est produite : " & Err.Description & vbCrLf & vbCrLf & " numéro d'erreur (" & Err.Number & " )", vbCritical, g_strNomApplication
  Print #g_FileDebug, Format(Now, "dd/MM/yyyy hh:mm:ss") & " Erreur n° : (" & Err.Number & ") " & Err.Description & " Procedure : vsHistoPerso_SelChange  Module : frmGestPerso.frm"
  Resume vsHistoPerso_SelChange_Exit
End Sub

Private Sub vsHistoPerso_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)

  If g_ModeDebug = vbUnchecked Then On Error GoTo vsHistoPerso_ValidateEdit_Error
  Modifperso

vsHistoPerso_ValidateEdit_Exit:
  Exit Sub

vsHistoPerso_ValidateEdit_Error:
  MsgBox "L'erreur suivante s'est produite : " & Err.Description & vbCrLf & vbCrLf & " numéro d'erreur (" & Err.Number & " )", vbCritical, g_strNomApplication
  Print #g_FileDebug, Format(Now, "dd/MM/yyyy hh:mm:ss") & " Erreur n° : (" & Err.Number & ") " & Err.Description & " Procedure : vsHistoPerso_ValidateEdit  Module : frmGestPerso.frm"
  Resume vsHistoPerso_ValidateEdit_Exit
End Sub

Private Sub vsPerso_Click()

  If g_ModeDebug = vbUnchecked Then On Error GoTo vsPerso_Click_Error
  vsPerso_SelChange

vsPerso_Click_Exit:
  Exit Sub

vsPerso_Click_Error:
  MsgBox "L'erreur suivante s'est produite : " & Err.Description & vbCrLf & vbCrLf & " numéro d'erreur (" & Err.Number & " )", vbCritical, g_strNomApplication
  Print #g_FileDebug, Format(Now, "dd/MM/yyyy hh:mm:ss") & " Erreur n° : (" & Err.Number & ") " & Err.Description & " Procedure : vsPerso_Click  Module : frmGestPerso.frm"
  Resume vsPerso_Click_Exit
End Sub

Public Sub vsPerso_SelChange()
Dim i As Integer, n As Integer
Dim bVisible As Boolean
Dim rst As Recordset

  If g_ModeDebug = vbUnchecked Then On Error GoTo vsPerso_SelChange_Error
  
  bVisible = False
  If bRefrech Then
    For i = 0 To vsPerso.Rows - 1
      If vsPerso.IsSelected(i) = True Then
        bChargePerso = True
        fldstrnom = vsPerso.Cell(flexcpText, i, L2_COL_NOM)
        RempHistoPerso (vsPerso.Cell(flexcpText, i, L2_COL_NOM))
        BloqueCombo
        CmbRace = vsPerso.Cell(flexcpText, i, L2_COL_RACE)
        SetValCombo CmbEcoleSpecialisation, vsPerso.Cell(flexcpText, i, L2_COL_TYPE_MAGICIEN)
        fldstrprofil = vsPerso.Cell(flexcpText, i, L2_COL_PROFIL)
        fldstrtitre = vsPerso.Cell(flexcpText, i, L2_COL_TITRE)
        fldstrdieu = vsPerso.Cell(flexcpText, i, L2_COL_DIEU)
        SetValCombo CmbAlignement, vsPerso.Cell(flexcpText, i, L2_COL_ALIGNEMENT)
        fldnBeaute = vsPerso.Cell(flexcpText, i, L2_COL_BEAUTE)
        fldstrcheveux = vsPerso.Cell(flexcpText, i, L2_COL_CHEVEUX)
        fldstryeux = vsPerso.Cell(flexcpText, i, L2_COL_YEUX)
        fldstrSexe = vsPerso.Cell(flexcpText, i, L2_COL_SEXE)
        fldnPoids = vsPerso.Cell(flexcpText, i, L2_COL_POIDS)
        fldnTaille = vsPerso.Cell(flexcpText, i, L2_COL_TAILLE)
        flddate = vsPerso.Cell(flexcpText, i, L2_COL_DATE)
        fldnAge = vsPerso.Cell(flexcpText, i, L2_COL_AGE)
        fldnMalusXP = vsPerso.Cell(flexcpText, i, L2_COL_MALUSXP)
        fldstrLangueConnue = vsPerso.Cell(flexcpText, i, L2_COL_LANGUE_CONNUE)
        fldstrAlphabet = vsPerso.Cell(flexcpText, i, L2_COL_ALPHABET)
        SetValCombo cmbDomaine(0), vsPerso.Cell(flexcpText, i, L2_COL_DOMAINE_1)
        SetValCombo cmbDomaine(1), vsPerso.Cell(flexcpText, i, L2_COL_DOMAINE_2)
        SetValCombo cmbDomaine(2), vsPerso.Cell(flexcpText, i, L2_COL_DOMAINE_3)
        SetValCombo cmbDomaine(3), vsPerso.Cell(flexcpText, i, L2_COL_DOMAINE_4)
        SetValCombo CmbEcole_interdite(0), vsPerso.Cell(flexcpText, i, L2_COL_ECOLE_INTERDITE_1)
        SetValCombo CmbEcole_interdite(1), vsPerso.Cell(flexcpText, i, L2_COL_ECOLE_INTERDITE_2)
        SetValCombo CmbEcole_interdite(2), vsPerso.Cell(flexcpText, i, L2_COL_ECOLE_INTERDITE_3)
        SetValCombo CmbEcole_interdite(3), vsPerso.Cell(flexcpText, i, L2_COL_ECOLE_INTERDITE_4)
        SetValCombo Cmbarchetype, vsPerso.Cell(flexcpText, i, L2_COL_ARCHETYPE)
        fldnTotalPo = vsPerso.Cell(flexcpText, i, L2_COL_TOTALPO)
        fldnTotalXP = vsPerso.Cell(flexcpText, i, L2_COL_TOTALXP)
        SetValCombo CmbProfession(0), vsPerso.Cell(flexcpText, i, L2_COL_PROFESSION_1)
        SetValCombo CmbProfession(1), vsPerso.Cell(flexcpText, i, L2_COL_PROFESSION_2)
        SetValCombo CmbProfession(2), vsPerso.Cell(flexcpText, i, L2_COL_PROFESSION_3)
        SetValCombo CmbProfession(3), vsPerso.Cell(flexcpText, i, L2_COL_PROFESSION_4)
        SetValCombo CmbArtisanat(0), vsPerso.Cell(flexcpText, i, L2_COL_ARTISANAT_1)
        SetValCombo CmbArtisanat(1), vsPerso.Cell(flexcpText, i, L2_COL_ARTISANAT_2)
        SetValCombo CmbArtisanat(2), vsPerso.Cell(flexcpText, i, L2_COL_ARTISANAT_3)
        SetValCombo CmbElu(0), vsPerso.Cell(flexcpText, i, L2_COL_ENERGIEELU_1)
        SetValCombo CmbElu(1), vsPerso.Cell(flexcpText, i, L2_COL_ENERGIEELU_2)
        SetValCombo CmbElu(2), vsPerso.Cell(flexcpText, i, L2_COL_ENERGIEELU_3)
        SetValCombo CmbSorcier(0), vsPerso.Cell(flexcpText, i, L2_COL_ENERGIESORCIER_1)
        SetValCombo CmbSorcier(1), vsPerso.Cell(flexcpText, i, L2_COL_ENERGIESORCIER_2)
        SetValCombo CmbWujen, vsPerso.Cell(flexcpText, i, L2_COL_ELEMENTWUJEN)
        SetValCombo CmbDragonTotem, vsPerso.Cell(flexcpText, i, L2_COL_DRAGONTOTEM)
        SetValCombo CmbShugenja(0), vsPerso.Cell(flexcpText, i, L2_COL_ELEMENTSHUGENJA)
        SetValCombo CmbShugenja(1), vsPerso.Cell(flexcpText, i, L2_COL_ORDRESHUGENJA)
  ' L2_COL_ENEMISJURES )
        n = 1
        vsEquip.Rows = 1
        Set rst = g_bdPerso.OpenRecordset("select * from Equipement where personnage='" & FormatStrSQL(fldstrnom) & "'", dbOpenDynaset)
        Do While Not rst.EOF
          vsEquip.Rows = n + 1
          vsEquip.Cell(flexcpText, n, L3_COL_EQUIP_NOM_OBJET) = rst!NomObjet
          vsEquip.Cell(flexcpText, n, L3_COL_EQUIP_TYPE) = rst!Type
          vsEquip.Cell(flexcpText, n, L3_COL_EQUIP_PLACE) = rst!Place
          vsEquip.Cell(flexcpText, n, L3_COL_EQUIP_CHARGE) = rst!CHARGE
          vsEquip.Cell(flexcpText, n, L3_COL_EQUIP_MULTIPLE) = IIf(IsNull(rst!Multiple), 1, rst!Multiple)
          vsEquip.Cell(flexcpText, n, L3_COL_EQUIP_SUR_PERSO) = IIf(rst!surpersonnage, "Oui", "Non")
          n = n + 1
          rst.MoveNext
        Loop
        rst.Close
        bVisible = True
      '  CmbRace.Enabled = vsHistoPerso.Rows < 1
        CmbRace.Enabled = True
        Exit For
      End If
    Next i
  bChargePerso = False
  bModif = False
  btnAjouter.Visible = True
  btnSupprimer.Visible = bVisible
  btnBackground.Visible = bVisible
  btnEnregistrer.Visible = False
  btnAnnuler.Visible = False
  BtnFermer.Visible = True
  btnRenommer.Visible = bVisible
  AjusteBouton Me, m_tabButton
  End If
vsPerso_SelChange_Exit:
  Exit Sub

vsPerso_SelChange_Error:
  MsgBox "L'erreur suivante s'est produite : " & Err.Description & vbCrLf & vbCrLf & " numéro d'erreur (" & Err.Number & " )", vbCritical, g_strNomApplication
  Print #g_FileDebug, Format(Now, "dd/MM/yyyy hh:mm:ss") & " Erreur n° : (" & Err.Number & ") " & Err.Description & " Procedure : vsPerso_SelChange  Module : frmGestPerso.frm"
  Resume vsPerso_SelChange_Exit
End Sub

