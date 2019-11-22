VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "Tabctl32.ocx"
Object = "{D76D7130-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7d.ocx"
Begin VB.Form frmGestobj 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Gestion des objets"
   ClientHeight    =   7860
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13410
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7860
   ScaleWidth      =   13410
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton btnRenommer 
      Caption         =   "Renommer"
      Height          =   385
      Left            =   8595
      TabIndex        =   451
      Top             =   45
      Width           =   1695
   End
   Begin VB.CommandButton btnAnnuler 
      Caption         =   "Annuler (ESC)"
      Height          =   385
      Left            =   6885
      TabIndex        =   450
      Top             =   45
      Width           =   1695
   End
   Begin VB.CommandButton btnSupprimer 
      Caption         =   "Supprimer (F10)"
      Height          =   385
      Left            =   5175
      TabIndex        =   449
      Top             =   45
      Width           =   1695
   End
   Begin VB.CommandButton btnEnregistrer 
      Caption         =   "Enregistrer (F5)"
      Height          =   385
      Left            =   3465
      TabIndex        =   448
      Top             =   45
      Width           =   1695
   End
   Begin VB.CommandButton btnAjouter 
      Caption         =   "Ajouter (F9)"
      Height          =   385
      Left            =   1755
      TabIndex        =   447
      Top             =   45
      Width           =   1695
   End
   Begin VB.CommandButton btnFermer 
      Caption         =   "Fermer (ESC)"
      Height          =   385
      Left            =   45
      TabIndex        =   446
      Top             =   45
      Width           =   1695
   End
   Begin VB.Timer Timer1 
      Interval        =   250
      Left            =   8640
      Top             =   90
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   9135
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   135
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.TextBox fldstrnomObjet 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   40
      TabIndex        =   8
      Top             =   7000
      Width           =   4140
   End
   Begin VSFlex7DAOCtl.VSFlexGrid vsObjets 
      Bindings        =   "frmGestobj.frx":0000
      Height          =   6525
      Left            =   15
      TabIndex        =   7
      Top             =   450
      Width           =   4140
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
      Cols            =   12
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   "Nom de l'objet"
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   1
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
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   9
      Top             =   7485
      Width           =   13410
      _ExtentX        =   23654
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   23601
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7065
      Left            =   4200
      TabIndex        =   10
      Top             =   450
      Width           =   9180
      _ExtentX        =   16193
      _ExtentY        =   12462
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Général"
      TabPicture(0)   =   "frmGestobj.frx":0014
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label18"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame4"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Propriétés"
      TabPicture(1)   =   "frmGestobj.frx":0030
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame5"
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame5 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6615
         Left            =   -74910
         TabIndex        =   442
         Top             =   330
         Width           =   8985
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
            Height          =   690
            Left            =   4080
            ScaleHeight     =   690
            ScaleWidth      =   585
            TabIndex        =   443
            Top             =   2520
            Width           =   585
            Begin VB.CommandButton btnAjout 
               Caption         =   "->"
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
               Left            =   120
               TabIndex        =   445
               Top             =   0
               Width           =   405
            End
            Begin VB.CommandButton btnEnlever 
               Caption         =   "<-"
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
               Left            =   120
               TabIndex        =   444
               Top             =   360
               Width           =   405
            End
         End
         Begin MSComctlLib.TreeView Tree1 
            Height          =   6225
            Left            =   45
            TabIndex        =   5
            Top             =   270
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   10980
            _Version        =   393217
            HideSelection   =   0   'False
            LabelEdit       =   1
            LineStyle       =   1
            Style           =   7
            Appearance      =   1
         End
         Begin VSFlex7DAOCtl.VSFlexGrid vsPropri 
            Height          =   6315
            Left            =   4680
            TabIndex        =   6
            Top             =   240
            Width           =   4200
            _cx             =   5080
            _cy             =   5080
            _ConvInfo       =   1
            Appearance      =   1
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MousePointer    =   0
            BackColor       =   16777215
            ForeColor       =   -2147483640
            BackColorFixed  =   -2147483633
            ForeColorFixed  =   -2147483630
            BackColorSel    =   -2147483635
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483643
            BackColorAlternate=   16777215
            GridColor       =   -2147483639
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
            GridLines       =   0
            GridLinesFixed  =   0
            GridLineWidth   =   0
            Rows            =   0
            Cols            =   5
            FixedRows       =   0
            FixedCols       =   0
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   "                                                             |Race                 "
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
      Begin VB.Frame Frame4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6680
         Left            =   135
         TabIndex        =   436
         Top             =   315
         Width           =   8985
         Begin VB.TextBox fldnCharge 
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
            Left            =   990
            TabIndex        =   476
            ToolTipText     =   "Définie le nombre de charge de l'objet dans la base de donnée"
            Top             =   2000
            Width           =   1275
         End
         Begin VB.ComboBox CmbTailles 
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
            Left            =   6750
            TabIndex        =   468
            Text            =   "Combo1"
            Top             =   225
            Width           =   1500
         End
         Begin VB.ComboBox CmbType 
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
            Left            =   6750
            TabIndex        =   467
            Text            =   "Combo1"
            Top             =   585
            Width           =   1500
         End
         Begin VB.ComboBox CmbDegatsDes 
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
            Left            =   6750
            TabIndex        =   466
            Text            =   "Combo1"
            Top             =   945
            Width           =   1500
         End
         Begin VB.ComboBox CmbZoneCrit 
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
            Left            =   6750
            TabIndex        =   465
            Text            =   "Combo1"
            Top             =   1305
            Width           =   1500
         End
         Begin VB.ComboBox CmbMultipCrit 
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
            Left            =   6750
            TabIndex        =   464
            Text            =   "Combo1"
            Top             =   1665
            Width           =   1500
         End
         Begin VB.TextBox fldnPortee 
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
            Left            =   6750
            MaxLength       =   5
            TabIndex        =   463
            Top             =   2025
            Width           =   1500
         End
         Begin VB.TextBox fldnresistance 
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
            Left            =   990
            TabIndex        =   460
            Top             =   1665
            Width           =   1275
         End
         Begin VB.TextBox fldnsolidite 
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
            Left            =   990
            TabIndex        =   459
            Top             =   1305
            Width           =   1275
         End
         Begin VB.ComboBox CmbGroupe 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1350
            Sorted          =   -1  'True
            TabIndex        =   455
            Text            =   "Combo1"
            Top             =   180
            Width           =   3885
         End
         Begin VB.TextBox fldnPoidsBase 
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
            Left            =   990
            TabIndex        =   454
            Top             =   945
            Width           =   1275
         End
         Begin VB.TextBox fldnCoutTotal 
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
            Left            =   990
            TabIndex        =   453
            Top             =   585
            Width           =   1275
         End
         Begin VB.TextBox fldstrdescription 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4230
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   452
            Top             =   2385
            Width           =   8760
         End
         Begin VB.ComboBox cmbClasseArmure 
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
            Left            =   3870
            TabIndex        =   0
            Text            =   "Combo1"
            Top             =   585
            Width           =   1365
         End
         Begin VB.ComboBox CmbTypeCa 
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
            Left            =   3870
            TabIndex        =   1
            Text            =   "Combo1"
            Top             =   945
            Width           =   1365
         End
         Begin VB.TextBox fldnEchecSorts 
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
            Left            =   3870
            MaxLength       =   2
            TabIndex        =   2
            Top             =   1305
            Width           =   1365
         End
         Begin VB.TextBox fldnPenaliteArmure 
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
            Left            =   3870
            MaxLength       =   2
            TabIndex        =   3
            Top             =   1665
            Width           =   1365
         End
         Begin VB.TextBox fldnBonusDexMax 
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
            Left            =   3870
            MaxLength       =   2
            TabIndex        =   4
            Top             =   2025
            Width           =   1365
         End
         Begin VB.Label Label7 
            Caption         =   "Charge Max"
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
            TabIndex        =   475
            Top             =   2025
            Width           =   990
         End
         Begin VB.Label Label2 
            Caption         =   "Taille"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   5355
            TabIndex        =   474
            Top             =   225
            Width           =   600
         End
         Begin VB.Label Label3 
            Caption         =   "Type"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   5355
            TabIndex        =   473
            Top             =   585
            Width           =   600
         End
         Begin VB.Label Label4 
            Caption         =   "Degats aux dès"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   5355
            TabIndex        =   472
            Top             =   945
            Width           =   1185
         End
         Begin VB.Label Label9 
            Caption         =   "Zone de critique"
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
            Left            =   5355
            TabIndex        =   471
            Top             =   1305
            Width           =   1275
         End
         Begin VB.Label Label10 
            Caption         =   "Multiplicateur crit"
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
            Left            =   5355
            TabIndex        =   470
            Top             =   1665
            Width           =   1275
         End
         Begin VB.Label Label1 
            Caption         =   "Portée"
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
            Left            =   5355
            TabIndex        =   469
            Top             =   2025
            Width           =   1005
         End
         Begin VB.Label Label6 
            Caption         =   "Résistance"
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
            TabIndex        =   462
            Top             =   1665
            Width           =   825
         End
         Begin VB.Label Label5 
            Caption         =   "Solidité"
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
            TabIndex        =   461
            Top             =   1305
            Width           =   825
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Groupe de l'objet"
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
            Left            =   45
            TabIndex        =   458
            Top             =   225
            Width           =   1245
         End
         Begin VB.Label Label16 
            Caption         =   "Poids Base"
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
            TabIndex        =   457
            Top             =   945
            Width           =   825
         End
         Begin VB.Label Label15 
            Caption         =   "Cout total"
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
            TabIndex        =   456
            Top             =   585
            Width           =   825
         End
         Begin VB.Label Label17 
            Caption         =   "Classe d'armure"
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
            Left            =   2385
            TabIndex        =   441
            Top             =   585
            Width           =   1230
         End
         Begin VB.Label Label19 
            Caption         =   "Type CA:"
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
            Left            =   2385
            TabIndex        =   440
            Top             =   945
            Width           =   1005
         End
         Begin VB.Label Label20 
            Caption         =   "Echec sorts profane"
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
            Left            =   2385
            TabIndex        =   439
            Top             =   1305
            Width           =   1455
         End
         Begin VB.Label Label21 
            Caption         =   "Pénalité armure"
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
            Left            =   2385
            TabIndex        =   438
            Top             =   1665
            Width           =   1455
         End
         Begin VB.Label Label22 
            Caption         =   "Bonus Dex max"
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
            Left            =   2400
            TabIndex        =   437
            Top             =   2025
            Width           =   1305
         End
      End
      Begin VSFlex7DAOCtl.VSFlexGrid vsEquip 
         Height          =   2385
         Left            =   -74100
         TabIndex        =   11
         Top             =   4725
         Width           =   7350
         _cx             =   5080
         _cy             =   5080
         _ConvInfo       =   1
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
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
         BackColorSel    =   -2147483639
         ForeColorSel    =   -2147483642
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
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   12
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   "Nom                  |Divers                     "
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
         Ellipsis        =   1
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
      Begin VSFlex7DAOCtl.VSFlexGrid vsDon 
         Height          =   7755
         Left            =   -74910
         TabIndex        =   12
         Top             =   510
         Width           =   8445
         _cx             =   5080
         _cy             =   5080
         _ConvInfo       =   1
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
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
         BackColorAlternate=   12648447
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
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   400
         RowHeightMax    =   400
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   "DONS                                  "
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
      Begin VB.Label Label18 
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
         Left            =   240
         TabIndex        =   435
         Top             =   1620
         Width           =   1005
      End
      Begin VB.Label LabCompetence 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   0
         Left            =   -74925
         TabIndex        =   434
         Top             =   720
         Width           =   1800
      End
      Begin VB.Label LabRace 
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
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   0
         Left            =   -73110
         TabIndex        =   433
         Top             =   720
         Width           =   300
      End
      Begin VB.Label Label77 
         Caption         =   "Cararctère    Niv objet  Total"
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
         Left            =   -72780
         TabIndex        =   432
         Top             =   405
         Width           =   2040
      End
      Begin VB.Label Label76 
         Caption         =   "Ra."
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
         Left            =   -73110
         TabIndex        =   431
         Top             =   405
         Width           =   270
      End
      Begin VB.Label Label75 
         Caption         =   "Competence"
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
         Left            =   -74925
         TabIndex        =   430
         Top             =   405
         Width           =   1725
      End
      Begin VB.Label labCar2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   0
         Left            =   -72510
         TabIndex        =   429
         Top             =   720
         Width           =   330
      End
      Begin VB.Label labCar1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   0
         Left            =   -72840
         TabIndex        =   428
         Top             =   720
         Width           =   330
      End
      Begin VB.Label labCar3 
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
         ForeColor       =   &H00404040&
         Height          =   285
         Index           =   0
         Left            =   -72210
         TabIndex        =   427
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
         ForeColor       =   &H000000FF&
         Height          =   285
         Index           =   0
         Left            =   -71040
         TabIndex        =   426
         Top             =   720
         Width           =   300
      End
      Begin VB.Label LabCompetence 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   -74925
         TabIndex        =   425
         Top             =   1035
         Width           =   1800
      End
      Begin VB.Label LabRace 
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
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   1
         Left            =   -73125
         TabIndex        =   424
         Top             =   1035
         Width           =   300
      End
      Begin VB.Label labCar2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   1
         Left            =   -72510
         TabIndex        =   423
         Top             =   1035
         Width           =   330
      End
      Begin VB.Label labCar1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   1
         Left            =   -72840
         TabIndex        =   422
         Top             =   1035
         Width           =   330
      End
      Begin VB.Label labCar3 
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
         ForeColor       =   &H00404040&
         Height          =   285
         Index           =   1
         Left            =   -72210
         TabIndex        =   421
         Top             =   1035
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
         ForeColor       =   &H000000FF&
         Height          =   285
         Index           =   1
         Left            =   -71040
         TabIndex        =   420
         Top             =   1035
         Width           =   300
      End
      Begin VB.Label LabCompetence 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   -74925
         TabIndex        =   419
         Top             =   1350
         Width           =   1800
      End
      Begin VB.Label LabRace 
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
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   2
         Left            =   -73125
         TabIndex        =   418
         Top             =   1350
         Width           =   300
      End
      Begin VB.Label labCar2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   2
         Left            =   -72510
         TabIndex        =   417
         Top             =   1350
         Width           =   330
      End
      Begin VB.Label labCar1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   2
         Left            =   -72840
         TabIndex        =   416
         Top             =   1350
         Width           =   330
      End
      Begin VB.Label labCar3 
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
         ForeColor       =   &H00404040&
         Height          =   285
         Index           =   2
         Left            =   -72210
         TabIndex        =   415
         Top             =   1350
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
         ForeColor       =   &H000000FF&
         Height          =   285
         Index           =   2
         Left            =   -71040
         TabIndex        =   414
         Top             =   1350
         Width           =   300
      End
      Begin VB.Label LabCompetence 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   -74925
         TabIndex        =   413
         Top             =   1665
         Width           =   1800
      End
      Begin VB.Label LabRace 
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
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   3
         Left            =   -73125
         TabIndex        =   412
         Top             =   1665
         Width           =   300
      End
      Begin VB.Label labCar2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   3
         Left            =   -72510
         TabIndex        =   411
         Top             =   1665
         Width           =   330
      End
      Begin VB.Label labCar1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   3
         Left            =   -72840
         TabIndex        =   410
         Top             =   1665
         Width           =   330
      End
      Begin VB.Label labCar3 
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
         ForeColor       =   &H00404040&
         Height          =   285
         Index           =   3
         Left            =   -72210
         TabIndex        =   409
         Top             =   1665
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
         ForeColor       =   &H000000FF&
         Height          =   285
         Index           =   3
         Left            =   -71040
         TabIndex        =   408
         Top             =   1665
         Width           =   300
      End
      Begin VB.Label LabCompetence 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   4
         Left            =   -74925
         TabIndex        =   407
         Top             =   1980
         Width           =   1800
      End
      Begin VB.Label LabRace 
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
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   4
         Left            =   -73125
         TabIndex        =   406
         Top             =   1980
         Width           =   300
      End
      Begin VB.Label labCar2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   4
         Left            =   -72510
         TabIndex        =   405
         Top             =   1980
         Width           =   330
      End
      Begin VB.Label labCar1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   4
         Left            =   -72840
         TabIndex        =   404
         Top             =   1980
         Width           =   330
      End
      Begin VB.Label labCar3 
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
         ForeColor       =   &H00404040&
         Height          =   285
         Index           =   4
         Left            =   -72210
         TabIndex        =   403
         Top             =   1980
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
         ForeColor       =   &H000000FF&
         Height          =   285
         Index           =   4
         Left            =   -71040
         TabIndex        =   402
         Top             =   1980
         Width           =   300
      End
      Begin VB.Label LabCompetence 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   5
         Left            =   -74925
         TabIndex        =   401
         Top             =   2295
         Width           =   1800
      End
      Begin VB.Label LabRace 
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
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   5
         Left            =   -73125
         TabIndex        =   400
         Top             =   2295
         Width           =   300
      End
      Begin VB.Label labCar2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   5
         Left            =   -72510
         TabIndex        =   399
         Top             =   2295
         Width           =   330
      End
      Begin VB.Label labCar1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   5
         Left            =   -72840
         TabIndex        =   398
         Top             =   2295
         Width           =   330
      End
      Begin VB.Label labCar3 
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
         ForeColor       =   &H00404040&
         Height          =   285
         Index           =   5
         Left            =   -72210
         TabIndex        =   397
         Top             =   2295
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
         ForeColor       =   &H000000FF&
         Height          =   285
         Index           =   5
         Left            =   -71040
         TabIndex        =   396
         Top             =   2295
         Width           =   300
      End
      Begin VB.Label LabCompetence 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   6
         Left            =   -74925
         TabIndex        =   395
         Top             =   2610
         Width           =   1800
      End
      Begin VB.Label LabRace 
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
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   6
         Left            =   -73125
         TabIndex        =   394
         Top             =   2610
         Width           =   300
      End
      Begin VB.Label labCar2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   6
         Left            =   -72510
         TabIndex        =   393
         Top             =   2610
         Width           =   330
      End
      Begin VB.Label labCar1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   6
         Left            =   -72840
         TabIndex        =   392
         Top             =   2610
         Width           =   330
      End
      Begin VB.Label labCar3 
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
         ForeColor       =   &H00404040&
         Height          =   285
         Index           =   6
         Left            =   -72210
         TabIndex        =   391
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
         ForeColor       =   &H000000FF&
         Height          =   285
         Index           =   6
         Left            =   -71040
         TabIndex        =   390
         Top             =   2610
         Width           =   300
      End
      Begin VB.Label LabCompetence 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   7
         Left            =   -74925
         TabIndex        =   389
         Top             =   2925
         Width           =   1800
      End
      Begin VB.Label LabRace 
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
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   7
         Left            =   -73125
         TabIndex        =   388
         Top             =   2925
         Width           =   300
      End
      Begin VB.Label labCar2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   7
         Left            =   -72510
         TabIndex        =   387
         Top             =   2925
         Width           =   330
      End
      Begin VB.Label labCar1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   7
         Left            =   -72840
         TabIndex        =   386
         Top             =   2925
         Width           =   330
      End
      Begin VB.Label labCar3 
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
         ForeColor       =   &H00404040&
         Height          =   285
         Index           =   7
         Left            =   -72210
         TabIndex        =   385
         Top             =   2925
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
         ForeColor       =   &H000000FF&
         Height          =   285
         Index           =   7
         Left            =   -71040
         TabIndex        =   384
         Top             =   2925
         Width           =   300
      End
      Begin VB.Label LabCompetence 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   8
         Left            =   -74925
         TabIndex        =   383
         Top             =   3240
         Width           =   1800
      End
      Begin VB.Label LabRace 
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
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   8
         Left            =   -73125
         TabIndex        =   382
         Top             =   3240
         Width           =   300
      End
      Begin VB.Label labCar2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   8
         Left            =   -72510
         TabIndex        =   381
         Top             =   3240
         Width           =   330
      End
      Begin VB.Label labCar1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   8
         Left            =   -72840
         TabIndex        =   380
         Top             =   3240
         Width           =   330
      End
      Begin VB.Label labCar3 
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
         ForeColor       =   &H00404040&
         Height          =   285
         Index           =   8
         Left            =   -72210
         TabIndex        =   379
         Top             =   3240
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
         ForeColor       =   &H000000FF&
         Height          =   285
         Index           =   8
         Left            =   -71040
         TabIndex        =   378
         Top             =   3240
         Width           =   300
      End
      Begin VB.Label LabCompetence 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   9
         Left            =   -74925
         TabIndex        =   377
         Top             =   3555
         Width           =   1800
      End
      Begin VB.Label LabRace 
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
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   9
         Left            =   -73125
         TabIndex        =   376
         Top             =   3555
         Width           =   300
      End
      Begin VB.Label labCar2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   9
         Left            =   -72510
         TabIndex        =   375
         Top             =   3555
         Width           =   330
      End
      Begin VB.Label labCar1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   9
         Left            =   -72840
         TabIndex        =   374
         Top             =   3555
         Width           =   330
      End
      Begin VB.Label labCar3 
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
         ForeColor       =   &H00404040&
         Height          =   285
         Index           =   9
         Left            =   -72210
         TabIndex        =   373
         Top             =   3555
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
         ForeColor       =   &H000000FF&
         Height          =   285
         Index           =   9
         Left            =   -71040
         TabIndex        =   372
         Top             =   3555
         Width           =   300
      End
      Begin VB.Label LabCompetence 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   10
         Left            =   -74925
         TabIndex        =   371
         Top             =   3870
         Width           =   1800
      End
      Begin VB.Label LabRace 
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
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   10
         Left            =   -73125
         TabIndex        =   370
         Top             =   3870
         Width           =   300
      End
      Begin VB.Label labCar2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   10
         Left            =   -72510
         TabIndex        =   369
         Top             =   3870
         Width           =   330
      End
      Begin VB.Label labCar1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   10
         Left            =   -72840
         TabIndex        =   368
         Top             =   3870
         Width           =   330
      End
      Begin VB.Label labCar3 
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
         ForeColor       =   &H00404040&
         Height          =   285
         Index           =   10
         Left            =   -72210
         TabIndex        =   367
         Top             =   3870
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
         ForeColor       =   &H000000FF&
         Height          =   285
         Index           =   10
         Left            =   -71040
         TabIndex        =   366
         Top             =   3870
         Width           =   300
      End
      Begin VB.Label LabCompetence 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   11
         Left            =   -74925
         TabIndex        =   365
         Top             =   4185
         Width           =   1800
      End
      Begin VB.Label LabRace 
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
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   11
         Left            =   -73125
         TabIndex        =   364
         Top             =   4185
         Width           =   300
      End
      Begin VB.Label labCar2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   11
         Left            =   -72510
         TabIndex        =   363
         Top             =   4185
         Width           =   330
      End
      Begin VB.Label labCar1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   11
         Left            =   -72840
         TabIndex        =   362
         Top             =   4185
         Width           =   330
      End
      Begin VB.Label labCar3 
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
         ForeColor       =   &H00404040&
         Height          =   285
         Index           =   11
         Left            =   -72210
         TabIndex        =   361
         Top             =   4185
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
         ForeColor       =   &H000000FF&
         Height          =   285
         Index           =   11
         Left            =   -71040
         TabIndex        =   360
         Top             =   4185
         Width           =   300
      End
      Begin VB.Label LabCompetence 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   12
         Left            =   -74925
         TabIndex        =   359
         Top             =   4500
         Width           =   1800
      End
      Begin VB.Label LabRace 
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
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   12
         Left            =   -73125
         TabIndex        =   358
         Top             =   4500
         Width           =   300
      End
      Begin VB.Label labCar2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   12
         Left            =   -72510
         TabIndex        =   357
         Top             =   4500
         Width           =   330
      End
      Begin VB.Label labCar1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   12
         Left            =   -72840
         TabIndex        =   356
         Top             =   4500
         Width           =   330
      End
      Begin VB.Label labCar3 
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
         ForeColor       =   &H00404040&
         Height          =   285
         Index           =   12
         Left            =   -72210
         TabIndex        =   355
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
         ForeColor       =   &H000000FF&
         Height          =   285
         Index           =   12
         Left            =   -71040
         TabIndex        =   354
         Top             =   4500
         Width           =   300
      End
      Begin VB.Label LabCompetence 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   13
         Left            =   -74925
         TabIndex        =   353
         Top             =   4815
         Width           =   1800
      End
      Begin VB.Label LabRace 
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
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   13
         Left            =   -73125
         TabIndex        =   352
         Top             =   4815
         Width           =   300
      End
      Begin VB.Label labCar2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   13
         Left            =   -72510
         TabIndex        =   351
         Top             =   4815
         Width           =   330
      End
      Begin VB.Label labCar1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   13
         Left            =   -72840
         TabIndex        =   350
         Top             =   4815
         Width           =   330
      End
      Begin VB.Label labCar3 
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
         ForeColor       =   &H00404040&
         Height          =   285
         Index           =   13
         Left            =   -72210
         TabIndex        =   349
         Top             =   4815
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
         ForeColor       =   &H000000FF&
         Height          =   285
         Index           =   13
         Left            =   -71040
         TabIndex        =   348
         Top             =   4815
         Width           =   300
      End
      Begin VB.Label LabCompetence 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   14
         Left            =   -74925
         TabIndex        =   347
         Top             =   5130
         Width           =   1800
      End
      Begin VB.Label LabRace 
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
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   14
         Left            =   -73125
         TabIndex        =   346
         Top             =   5130
         Width           =   300
      End
      Begin VB.Label labCar2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   14
         Left            =   -72510
         TabIndex        =   345
         Top             =   5130
         Width           =   330
      End
      Begin VB.Label labCar1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   14
         Left            =   -72840
         TabIndex        =   344
         Top             =   5130
         Width           =   330
      End
      Begin VB.Label labCar3 
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
         ForeColor       =   &H00404040&
         Height          =   285
         Index           =   14
         Left            =   -72210
         TabIndex        =   343
         Top             =   5130
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
         ForeColor       =   &H000000FF&
         Height          =   285
         Index           =   14
         Left            =   -71040
         TabIndex        =   342
         Top             =   5130
         Width           =   300
      End
      Begin VB.Label LabCompetence 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   15
         Left            =   -74925
         TabIndex        =   341
         Top             =   5445
         Width           =   1800
      End
      Begin VB.Label LabRace 
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
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   15
         Left            =   -73125
         TabIndex        =   340
         Top             =   5445
         Width           =   300
      End
      Begin VB.Label labCar2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   15
         Left            =   -72510
         TabIndex        =   339
         Top             =   5445
         Width           =   330
      End
      Begin VB.Label labCar1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   15
         Left            =   -72840
         TabIndex        =   338
         Top             =   5445
         Width           =   330
      End
      Begin VB.Label labCar3 
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
         ForeColor       =   &H00404040&
         Height          =   285
         Index           =   15
         Left            =   -72210
         TabIndex        =   337
         Top             =   5445
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
         ForeColor       =   &H000000FF&
         Height          =   285
         Index           =   15
         Left            =   -71040
         TabIndex        =   336
         Top             =   5445
         Width           =   300
      End
      Begin VB.Label LabCompetence 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   16
         Left            =   -74925
         TabIndex        =   335
         Top             =   5760
         Width           =   1800
      End
      Begin VB.Label LabRace 
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
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   16
         Left            =   -73125
         TabIndex        =   334
         Top             =   5760
         Width           =   300
      End
      Begin VB.Label labCar2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   16
         Left            =   -72510
         TabIndex        =   333
         Top             =   5760
         Width           =   330
      End
      Begin VB.Label labCar1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   16
         Left            =   -72840
         TabIndex        =   332
         Top             =   5760
         Width           =   330
      End
      Begin VB.Label labCar3 
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
         ForeColor       =   &H00404040&
         Height          =   285
         Index           =   16
         Left            =   -72210
         TabIndex        =   331
         Top             =   5760
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
         ForeColor       =   &H000000FF&
         Height          =   285
         Index           =   16
         Left            =   -71040
         TabIndex        =   330
         Top             =   5760
         Width           =   300
      End
      Begin VB.Label LabCompetence 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   17
         Left            =   -74925
         TabIndex        =   329
         Top             =   6075
         Width           =   1800
      End
      Begin VB.Label LabRace 
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
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   17
         Left            =   -73125
         TabIndex        =   328
         Top             =   6075
         Width           =   300
      End
      Begin VB.Label labCar2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   17
         Left            =   -72510
         TabIndex        =   327
         Top             =   6075
         Width           =   330
      End
      Begin VB.Label labCar1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   17
         Left            =   -72840
         TabIndex        =   326
         Top             =   6075
         Width           =   330
      End
      Begin VB.Label labCar3 
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
         ForeColor       =   &H00404040&
         Height          =   285
         Index           =   17
         Left            =   -72210
         TabIndex        =   325
         Top             =   6075
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
         ForeColor       =   &H000000FF&
         Height          =   285
         Index           =   17
         Left            =   -71040
         TabIndex        =   324
         Top             =   6075
         Width           =   300
      End
      Begin VB.Label LabCompetence 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   18
         Left            =   -74925
         TabIndex        =   323
         Top             =   6390
         Width           =   1800
      End
      Begin VB.Label LabRace 
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
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   18
         Left            =   -73125
         TabIndex        =   322
         Top             =   6390
         Width           =   300
      End
      Begin VB.Label labCar2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   18
         Left            =   -72510
         TabIndex        =   321
         Top             =   6390
         Width           =   330
      End
      Begin VB.Label labCar1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   18
         Left            =   -72840
         TabIndex        =   320
         Top             =   6390
         Width           =   330
      End
      Begin VB.Label labCar3 
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
         ForeColor       =   &H00404040&
         Height          =   285
         Index           =   18
         Left            =   -72210
         TabIndex        =   319
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
         ForeColor       =   &H000000FF&
         Height          =   285
         Index           =   18
         Left            =   -71040
         TabIndex        =   318
         Top             =   6390
         Width           =   300
      End
      Begin VB.Label LabCompetence 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   19
         Left            =   -74925
         TabIndex        =   317
         Top             =   6705
         Width           =   1800
      End
      Begin VB.Label LabRace 
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
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   19
         Left            =   -73125
         TabIndex        =   316
         Top             =   6705
         Width           =   300
      End
      Begin VB.Label labCar2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   19
         Left            =   -72510
         TabIndex        =   315
         Top             =   6705
         Width           =   330
      End
      Begin VB.Label labCar1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   19
         Left            =   -72840
         TabIndex        =   314
         Top             =   6705
         Width           =   330
      End
      Begin VB.Label labCar3 
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
         ForeColor       =   &H00404040&
         Height          =   285
         Index           =   19
         Left            =   -72210
         TabIndex        =   313
         Top             =   6705
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
         ForeColor       =   &H000000FF&
         Height          =   285
         Index           =   19
         Left            =   -71040
         TabIndex        =   312
         Top             =   6705
         Width           =   300
      End
      Begin VB.Label LabCompetence 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   20
         Left            =   -74925
         TabIndex        =   311
         Top             =   7020
         Width           =   1800
      End
      Begin VB.Label LabRace 
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
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   20
         Left            =   -73125
         TabIndex        =   310
         Top             =   7020
         Width           =   300
      End
      Begin VB.Label labCar2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   20
         Left            =   -72510
         TabIndex        =   309
         Top             =   7020
         Width           =   330
      End
      Begin VB.Label labCar1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   20
         Left            =   -72840
         TabIndex        =   308
         Top             =   7020
         Width           =   330
      End
      Begin VB.Label labCar3 
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
         ForeColor       =   &H00404040&
         Height          =   285
         Index           =   20
         Left            =   -72210
         TabIndex        =   307
         Top             =   7020
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
         ForeColor       =   &H000000FF&
         Height          =   285
         Index           =   20
         Left            =   -71040
         TabIndex        =   306
         Top             =   7020
         Width           =   300
      End
      Begin VB.Label LabCompetence 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   21
         Left            =   -74925
         TabIndex        =   305
         Top             =   7335
         Width           =   1800
      End
      Begin VB.Label LabRace 
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
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   21
         Left            =   -73125
         TabIndex        =   304
         Top             =   7335
         Width           =   300
      End
      Begin VB.Label labCar2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   21
         Left            =   -72510
         TabIndex        =   303
         Top             =   7335
         Width           =   330
      End
      Begin VB.Label labCar1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   21
         Left            =   -72840
         TabIndex        =   302
         Top             =   7335
         Width           =   330
      End
      Begin VB.Label labCar3 
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
         ForeColor       =   &H00404040&
         Height          =   285
         Index           =   21
         Left            =   -72210
         TabIndex        =   301
         Top             =   7335
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
         ForeColor       =   &H000000FF&
         Height          =   285
         Index           =   21
         Left            =   -71040
         TabIndex        =   300
         Top             =   7335
         Width           =   300
      End
      Begin VB.Label LabCompetence 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   22
         Left            =   -74925
         TabIndex        =   299
         Top             =   7650
         Width           =   1800
      End
      Begin VB.Label LabRace 
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
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   22
         Left            =   -73125
         TabIndex        =   298
         Top             =   7650
         Width           =   300
      End
      Begin VB.Label labCar2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   22
         Left            =   -72510
         TabIndex        =   297
         Top             =   7650
         Width           =   330
      End
      Begin VB.Label labCar1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   22
         Left            =   -72840
         TabIndex        =   296
         Top             =   7650
         Width           =   330
      End
      Begin VB.Label labCar3 
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
         ForeColor       =   &H00404040&
         Height          =   285
         Index           =   22
         Left            =   -72210
         TabIndex        =   295
         Top             =   7650
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
         ForeColor       =   &H000000FF&
         Height          =   285
         Index           =   22
         Left            =   -71040
         TabIndex        =   294
         Top             =   7650
         Width           =   300
      End
      Begin VB.Line Line4 
         X1              =   -70680
         X2              =   -70680
         Y1              =   480
         Y2              =   8040
      End
      Begin VB.Label labnNiv 
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
         ForeColor       =   &H00404040&
         Height          =   285
         Index           =   0
         Left            =   -71910
         TabIndex        =   293
         Top             =   720
         Width           =   300
      End
      Begin VB.Label labnNiv 
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
         ForeColor       =   &H00404040&
         Height          =   285
         Index           =   1
         Left            =   -71910
         TabIndex        =   292
         Top             =   1035
         Width           =   300
      End
      Begin VB.Label labnNiv 
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
         ForeColor       =   &H00404040&
         Height          =   285
         Index           =   2
         Left            =   -71910
         TabIndex        =   291
         Top             =   1350
         Width           =   300
      End
      Begin VB.Label labnNiv 
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
         ForeColor       =   &H00404040&
         Height          =   285
         Index           =   3
         Left            =   -71910
         TabIndex        =   290
         Top             =   1665
         Width           =   300
      End
      Begin VB.Label labnNiv 
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
         ForeColor       =   &H00404040&
         Height          =   285
         Index           =   4
         Left            =   -71910
         TabIndex        =   289
         Top             =   1980
         Width           =   300
      End
      Begin VB.Label labnNiv 
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
         ForeColor       =   &H00404040&
         Height          =   285
         Index           =   5
         Left            =   -71910
         TabIndex        =   288
         Top             =   2295
         Width           =   300
      End
      Begin VB.Label labnNiv 
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
         ForeColor       =   &H00404040&
         Height          =   285
         Index           =   6
         Left            =   -71910
         TabIndex        =   287
         Top             =   2610
         Width           =   300
      End
      Begin VB.Label labnNiv 
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
         ForeColor       =   &H00404040&
         Height          =   285
         Index           =   7
         Left            =   -71910
         TabIndex        =   286
         Top             =   2925
         Width           =   300
      End
      Begin VB.Label labnNiv 
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
         ForeColor       =   &H00404040&
         Height          =   285
         Index           =   8
         Left            =   -71910
         TabIndex        =   285
         Top             =   3240
         Width           =   300
      End
      Begin VB.Label labnNiv 
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
         ForeColor       =   &H00404040&
         Height          =   285
         Index           =   9
         Left            =   -71910
         TabIndex        =   284
         Top             =   3570
         Width           =   300
      End
      Begin VB.Label labnNiv 
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
         ForeColor       =   &H00404040&
         Height          =   285
         Index           =   10
         Left            =   -71910
         TabIndex        =   283
         Top             =   3870
         Width           =   300
      End
      Begin VB.Label labnNiv 
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
         ForeColor       =   &H00404040&
         Height          =   285
         Index           =   11
         Left            =   -71910
         TabIndex        =   282
         Top             =   4185
         Width           =   300
      End
      Begin VB.Label labnNiv 
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
         ForeColor       =   &H00404040&
         Height          =   285
         Index           =   12
         Left            =   -71910
         TabIndex        =   281
         Top             =   4500
         Width           =   300
      End
      Begin VB.Label labnNiv 
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
         ForeColor       =   &H00404040&
         Height          =   285
         Index           =   13
         Left            =   -71910
         TabIndex        =   280
         Top             =   4815
         Width           =   300
      End
      Begin VB.Label labnNiv 
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
         ForeColor       =   &H00404040&
         Height          =   285
         Index           =   14
         Left            =   -71910
         TabIndex        =   279
         Top             =   5130
         Width           =   300
      End
      Begin VB.Label labnNiv 
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
         ForeColor       =   &H00404040&
         Height          =   285
         Index           =   15
         Left            =   -71910
         TabIndex        =   278
         Top             =   5445
         Width           =   300
      End
      Begin VB.Label labnNiv 
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
         ForeColor       =   &H00404040&
         Height          =   285
         Index           =   16
         Left            =   -71910
         TabIndex        =   277
         Top             =   5760
         Width           =   300
      End
      Begin VB.Label labnNiv 
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
         ForeColor       =   &H00404040&
         Height          =   285
         Index           =   17
         Left            =   -71910
         TabIndex        =   276
         Top             =   6075
         Width           =   300
      End
      Begin VB.Label labnNiv 
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
         ForeColor       =   &H00404040&
         Height          =   285
         Index           =   18
         Left            =   -71910
         TabIndex        =   275
         Top             =   6390
         Width           =   300
      End
      Begin VB.Label labnNiv 
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
         ForeColor       =   &H00404040&
         Height          =   285
         Index           =   19
         Left            =   -71910
         TabIndex        =   274
         Top             =   6705
         Width           =   300
      End
      Begin VB.Label labnNiv 
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
         ForeColor       =   &H00404040&
         Height          =   285
         Index           =   20
         Left            =   -71910
         TabIndex        =   273
         Top             =   7020
         Width           =   300
      End
      Begin VB.Label labnNiv 
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
         ForeColor       =   &H00404040&
         Height          =   285
         Index           =   21
         Left            =   -71910
         TabIndex        =   272
         Top             =   7335
         Width           =   300
      End
      Begin VB.Label labnNiv 
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
         ForeColor       =   &H00404040&
         Height          =   285
         Index           =   22
         Left            =   -71910
         TabIndex        =   271
         Top             =   7650
         Width           =   300
      End
      Begin VB.Label Label74 
         Caption         =   "Competence"
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
         Left            =   -70500
         TabIndex        =   270
         Top             =   405
         Width           =   1050
      End
      Begin VB.Label Label57 
         Caption         =   "Ra."
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
         Left            =   -68760
         TabIndex        =   269
         Top             =   390
         Width           =   270
      End
      Begin VB.Label LabCompetence 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   45
         Left            =   -70620
         TabIndex        =   268
         Top             =   7650
         Width           =   1800
      End
      Begin VB.Label LabCompetence 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   44
         Left            =   -70620
         TabIndex        =   267
         Top             =   7335
         Width           =   1800
      End
      Begin VB.Label LabCompetence 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   43
         Left            =   -70620
         TabIndex        =   266
         Top             =   7020
         Width           =   1800
      End
      Begin VB.Label LabCompetence 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   42
         Left            =   -70620
         TabIndex        =   265
         Top             =   6705
         Width           =   1800
      End
      Begin VB.Label LabCompetence 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   41
         Left            =   -70620
         TabIndex        =   264
         Top             =   6390
         Width           =   1800
      End
      Begin VB.Label LabCompetence 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   40
         Left            =   -70620
         TabIndex        =   263
         Top             =   6075
         Width           =   1800
      End
      Begin VB.Label LabCompetence 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   39
         Left            =   -70620
         TabIndex        =   262
         Top             =   5760
         Width           =   1800
      End
      Begin VB.Label LabCompetence 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   38
         Left            =   -70620
         TabIndex        =   261
         Top             =   5445
         Width           =   1800
      End
      Begin VB.Label LabCompetence 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   37
         Left            =   -70620
         TabIndex        =   260
         Top             =   5130
         Width           =   1800
      End
      Begin VB.Label LabCompetence 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   36
         Left            =   -70620
         TabIndex        =   259
         Top             =   4815
         Width           =   1800
      End
      Begin VB.Label LabCompetence 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   35
         Left            =   -70620
         TabIndex        =   258
         Top             =   4500
         Width           =   1800
      End
      Begin VB.Label LabCompetence 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   34
         Left            =   -70620
         TabIndex        =   257
         Top             =   4185
         Width           =   1800
      End
      Begin VB.Label LabCompetence 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   33
         Left            =   -70620
         TabIndex        =   256
         Top             =   3870
         Width           =   1800
      End
      Begin VB.Label LabCompetence 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   32
         Left            =   -70620
         TabIndex        =   255
         Top             =   3555
         Width           =   1800
      End
      Begin VB.Label LabCompetence 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   31
         Left            =   -70620
         TabIndex        =   254
         Top             =   3240
         Width           =   1800
      End
      Begin VB.Label LabCompetence 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   30
         Left            =   -70620
         TabIndex        =   253
         Top             =   2925
         Width           =   1800
      End
      Begin VB.Label LabCompetence 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   29
         Left            =   -70620
         TabIndex        =   252
         Top             =   2610
         Width           =   1800
      End
      Begin VB.Label LabCompetence 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   28
         Left            =   -70620
         TabIndex        =   251
         Top             =   2295
         Width           =   1800
      End
      Begin VB.Label LabCompetence 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   27
         Left            =   -70620
         TabIndex        =   250
         Top             =   1980
         Width           =   1800
      End
      Begin VB.Label LabCompetence 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   26
         Left            =   -70620
         TabIndex        =   249
         Top             =   1665
         Width           =   1800
      End
      Begin VB.Label LabCompetence 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   25
         Left            =   -70620
         TabIndex        =   248
         Top             =   1350
         Width           =   1800
      End
      Begin VB.Label LabCompetence 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   24
         Left            =   -70620
         TabIndex        =   247
         Top             =   1035
         Width           =   1800
      End
      Begin VB.Label LabCompetence 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   23
         Left            =   -70620
         TabIndex        =   246
         Top             =   720
         Width           =   1800
      End
      Begin VB.Label LabRace 
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
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   23
         Left            =   -68805
         TabIndex        =   245
         Top             =   720
         Width           =   300
      End
      Begin VB.Label labCar2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   23
         Left            =   -68205
         TabIndex        =   244
         Top             =   720
         Width           =   330
      End
      Begin VB.Label labCar1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   23
         Left            =   -68535
         TabIndex        =   243
         Top             =   720
         Width           =   330
      End
      Begin VB.Label labCar3 
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
         ForeColor       =   &H00404040&
         Height          =   285
         Index           =   23
         Left            =   -67905
         TabIndex        =   242
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
         ForeColor       =   &H000000FF&
         Height          =   285
         Index           =   23
         Left            =   -66705
         TabIndex        =   241
         Top             =   720
         Width           =   300
      End
      Begin VB.Label LabRace 
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
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   24
         Left            =   -68820
         TabIndex        =   240
         Top             =   1035
         Width           =   300
      End
      Begin VB.Label labCar2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   24
         Left            =   -68205
         TabIndex        =   239
         Top             =   1035
         Width           =   330
      End
      Begin VB.Label labCar1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   24
         Left            =   -68535
         TabIndex        =   238
         Top             =   1035
         Width           =   330
      End
      Begin VB.Label labCar3 
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
         ForeColor       =   &H00404040&
         Height          =   285
         Index           =   24
         Left            =   -67905
         TabIndex        =   237
         Top             =   1035
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
         ForeColor       =   &H000000FF&
         Height          =   285
         Index           =   24
         Left            =   -66705
         TabIndex        =   236
         Top             =   1035
         Width           =   300
      End
      Begin VB.Label LabRace 
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
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   25
         Left            =   -68820
         TabIndex        =   235
         Top             =   1350
         Width           =   300
      End
      Begin VB.Label labCar2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   25
         Left            =   -68205
         TabIndex        =   234
         Top             =   1350
         Width           =   330
      End
      Begin VB.Label labCar1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   25
         Left            =   -68535
         TabIndex        =   233
         Top             =   1350
         Width           =   330
      End
      Begin VB.Label labCar3 
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
         ForeColor       =   &H00404040&
         Height          =   285
         Index           =   25
         Left            =   -67905
         TabIndex        =   232
         Top             =   1350
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
         ForeColor       =   &H000000FF&
         Height          =   285
         Index           =   25
         Left            =   -66705
         TabIndex        =   231
         Top             =   1350
         Width           =   300
      End
      Begin VB.Label LabRace 
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
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   26
         Left            =   -68820
         TabIndex        =   230
         Top             =   1665
         Width           =   300
      End
      Begin VB.Label labCar2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   26
         Left            =   -68205
         TabIndex        =   229
         Top             =   1665
         Width           =   330
      End
      Begin VB.Label labCar1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   26
         Left            =   -68535
         TabIndex        =   228
         Top             =   1665
         Width           =   330
      End
      Begin VB.Label labCar3 
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
         ForeColor       =   &H00404040&
         Height          =   285
         Index           =   26
         Left            =   -67905
         TabIndex        =   227
         Top             =   1665
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
         ForeColor       =   &H000000FF&
         Height          =   285
         Index           =   26
         Left            =   -66705
         TabIndex        =   226
         Top             =   1665
         Width           =   300
      End
      Begin VB.Label LabRace 
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
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   27
         Left            =   -68820
         TabIndex        =   225
         Top             =   1980
         Width           =   300
      End
      Begin VB.Label labCar2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   27
         Left            =   -68205
         TabIndex        =   224
         Top             =   1980
         Width           =   330
      End
      Begin VB.Label labCar1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   27
         Left            =   -68535
         TabIndex        =   223
         Top             =   1980
         Width           =   330
      End
      Begin VB.Label labCar3 
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
         ForeColor       =   &H00404040&
         Height          =   285
         Index           =   27
         Left            =   -67905
         TabIndex        =   222
         Top             =   1980
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
         ForeColor       =   &H000000FF&
         Height          =   285
         Index           =   27
         Left            =   -66705
         TabIndex        =   221
         Top             =   1980
         Width           =   300
      End
      Begin VB.Label LabRace 
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
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   28
         Left            =   -68820
         TabIndex        =   220
         Top             =   2295
         Width           =   300
      End
      Begin VB.Label labCar2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   28
         Left            =   -68205
         TabIndex        =   219
         Top             =   2295
         Width           =   330
      End
      Begin VB.Label labCar1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   28
         Left            =   -68535
         TabIndex        =   218
         Top             =   2295
         Width           =   330
      End
      Begin VB.Label labCar3 
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
         ForeColor       =   &H00404040&
         Height          =   285
         Index           =   28
         Left            =   -67905
         TabIndex        =   217
         Top             =   2295
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
         ForeColor       =   &H000000FF&
         Height          =   285
         Index           =   28
         Left            =   -66705
         TabIndex        =   216
         Top             =   2295
         Width           =   300
      End
      Begin VB.Label LabRace 
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
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   29
         Left            =   -68820
         TabIndex        =   215
         Top             =   2610
         Width           =   300
      End
      Begin VB.Label labCar2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   29
         Left            =   -68205
         TabIndex        =   214
         Top             =   2610
         Width           =   330
      End
      Begin VB.Label labCar1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   29
         Left            =   -68535
         TabIndex        =   213
         Top             =   2610
         Width           =   330
      End
      Begin VB.Label labCar3 
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
         ForeColor       =   &H00404040&
         Height          =   285
         Index           =   29
         Left            =   -67905
         TabIndex        =   212
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
         ForeColor       =   &H000000FF&
         Height          =   285
         Index           =   29
         Left            =   -66705
         TabIndex        =   211
         Top             =   2610
         Width           =   300
      End
      Begin VB.Label LabRace 
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
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   30
         Left            =   -68820
         TabIndex        =   210
         Top             =   2925
         Width           =   300
      End
      Begin VB.Label labCar2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   30
         Left            =   -68205
         TabIndex        =   209
         Top             =   2925
         Width           =   330
      End
      Begin VB.Label labCar1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   30
         Left            =   -68535
         TabIndex        =   208
         Top             =   2925
         Width           =   330
      End
      Begin VB.Label labCar3 
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
         ForeColor       =   &H00404040&
         Height          =   285
         Index           =   30
         Left            =   -67905
         TabIndex        =   207
         Top             =   2925
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
         ForeColor       =   &H000000FF&
         Height          =   285
         Index           =   30
         Left            =   -66705
         TabIndex        =   206
         Top             =   2925
         Width           =   300
      End
      Begin VB.Label LabRace 
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
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   31
         Left            =   -68820
         TabIndex        =   205
         Top             =   3240
         Width           =   300
      End
      Begin VB.Label labCar2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   31
         Left            =   -68205
         TabIndex        =   204
         Top             =   3240
         Width           =   330
      End
      Begin VB.Label labCar1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   31
         Left            =   -68535
         TabIndex        =   203
         Top             =   3240
         Width           =   330
      End
      Begin VB.Label labCar3 
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
         ForeColor       =   &H00404040&
         Height          =   285
         Index           =   31
         Left            =   -67905
         TabIndex        =   202
         Top             =   3240
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
         ForeColor       =   &H000000FF&
         Height          =   285
         Index           =   31
         Left            =   -66705
         TabIndex        =   201
         Top             =   3240
         Width           =   300
      End
      Begin VB.Label LabRace 
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
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   32
         Left            =   -68820
         TabIndex        =   200
         Top             =   3555
         Width           =   300
      End
      Begin VB.Label labCar2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   32
         Left            =   -68205
         TabIndex        =   199
         Top             =   3555
         Width           =   330
      End
      Begin VB.Label labCar1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   32
         Left            =   -68535
         TabIndex        =   198
         Top             =   3555
         Width           =   330
      End
      Begin VB.Label labCar3 
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
         ForeColor       =   &H00404040&
         Height          =   285
         Index           =   32
         Left            =   -67905
         TabIndex        =   197
         Top             =   3555
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
         ForeColor       =   &H000000FF&
         Height          =   285
         Index           =   32
         Left            =   -66705
         TabIndex        =   196
         Top             =   3555
         Width           =   300
      End
      Begin VB.Label LabRace 
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
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   33
         Left            =   -68820
         TabIndex        =   195
         Top             =   3870
         Width           =   300
      End
      Begin VB.Label labCar2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   33
         Left            =   -68205
         TabIndex        =   194
         Top             =   3870
         Width           =   330
      End
      Begin VB.Label labCar1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   33
         Left            =   -68535
         TabIndex        =   193
         Top             =   3870
         Width           =   330
      End
      Begin VB.Label labCar3 
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
         ForeColor       =   &H00404040&
         Height          =   285
         Index           =   33
         Left            =   -67905
         TabIndex        =   192
         Top             =   3870
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
         ForeColor       =   &H000000FF&
         Height          =   285
         Index           =   33
         Left            =   -66705
         TabIndex        =   191
         Top             =   3870
         Width           =   300
      End
      Begin VB.Label LabRace 
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
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   34
         Left            =   -68820
         TabIndex        =   190
         Top             =   4185
         Width           =   300
      End
      Begin VB.Label labCar2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   34
         Left            =   -68205
         TabIndex        =   189
         Top             =   4185
         Width           =   330
      End
      Begin VB.Label labCar1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   34
         Left            =   -68535
         TabIndex        =   188
         Top             =   4185
         Width           =   330
      End
      Begin VB.Label labCar3 
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
         ForeColor       =   &H00404040&
         Height          =   285
         Index           =   34
         Left            =   -67905
         TabIndex        =   187
         Top             =   4185
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
         ForeColor       =   &H000000FF&
         Height          =   285
         Index           =   34
         Left            =   -66705
         TabIndex        =   186
         Top             =   4185
         Width           =   300
      End
      Begin VB.Label LabRace 
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
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   35
         Left            =   -68820
         TabIndex        =   185
         Top             =   4500
         Width           =   300
      End
      Begin VB.Label labCar2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   35
         Left            =   -68205
         TabIndex        =   184
         Top             =   4500
         Width           =   330
      End
      Begin VB.Label labCar1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   35
         Left            =   -68535
         TabIndex        =   183
         Top             =   4500
         Width           =   330
      End
      Begin VB.Label labCar3 
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
         ForeColor       =   &H00404040&
         Height          =   285
         Index           =   35
         Left            =   -67905
         TabIndex        =   182
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
         ForeColor       =   &H000000FF&
         Height          =   285
         Index           =   35
         Left            =   -66705
         TabIndex        =   181
         Top             =   4500
         Width           =   300
      End
      Begin VB.Label LabRace 
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
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   36
         Left            =   -68820
         TabIndex        =   180
         Top             =   4815
         Width           =   300
      End
      Begin VB.Label labCar2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   36
         Left            =   -68205
         TabIndex        =   179
         Top             =   4815
         Width           =   330
      End
      Begin VB.Label labCar1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   36
         Left            =   -68535
         TabIndex        =   178
         Top             =   4815
         Width           =   330
      End
      Begin VB.Label labCar3 
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
         ForeColor       =   &H00404040&
         Height          =   285
         Index           =   36
         Left            =   -67905
         TabIndex        =   177
         Top             =   4815
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
         ForeColor       =   &H000000FF&
         Height          =   285
         Index           =   36
         Left            =   -66705
         TabIndex        =   176
         Top             =   4815
         Width           =   300
      End
      Begin VB.Label LabRace 
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
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   37
         Left            =   -68820
         TabIndex        =   175
         Top             =   5130
         Width           =   300
      End
      Begin VB.Label labCar2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   37
         Left            =   -68205
         TabIndex        =   174
         Top             =   5130
         Width           =   330
      End
      Begin VB.Label labCar1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   37
         Left            =   -68535
         TabIndex        =   173
         Top             =   5130
         Width           =   330
      End
      Begin VB.Label labCar3 
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
         ForeColor       =   &H00404040&
         Height          =   285
         Index           =   37
         Left            =   -67905
         TabIndex        =   172
         Top             =   5130
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
         ForeColor       =   &H000000FF&
         Height          =   285
         Index           =   37
         Left            =   -66705
         TabIndex        =   171
         Top             =   5130
         Width           =   300
      End
      Begin VB.Label LabRace 
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
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   38
         Left            =   -68820
         TabIndex        =   170
         Top             =   5445
         Width           =   300
      End
      Begin VB.Label labCar2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   38
         Left            =   -68205
         TabIndex        =   169
         Top             =   5445
         Width           =   330
      End
      Begin VB.Label labCar1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   38
         Left            =   -68535
         TabIndex        =   168
         Top             =   5445
         Width           =   330
      End
      Begin VB.Label labCar3 
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
         ForeColor       =   &H00404040&
         Height          =   285
         Index           =   38
         Left            =   -67905
         TabIndex        =   167
         Top             =   5445
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
         ForeColor       =   &H000000FF&
         Height          =   285
         Index           =   38
         Left            =   -66705
         TabIndex        =   166
         Top             =   5445
         Width           =   300
      End
      Begin VB.Label LabRace 
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
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   39
         Left            =   -68820
         TabIndex        =   165
         Top             =   5760
         Width           =   300
      End
      Begin VB.Label labCar2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   39
         Left            =   -68205
         TabIndex        =   164
         Top             =   5760
         Width           =   330
      End
      Begin VB.Label labCar1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   39
         Left            =   -68535
         TabIndex        =   163
         Top             =   5760
         Width           =   330
      End
      Begin VB.Label labCar3 
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
         ForeColor       =   &H00404040&
         Height          =   285
         Index           =   39
         Left            =   -67905
         TabIndex        =   162
         Top             =   5760
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
         ForeColor       =   &H000000FF&
         Height          =   285
         Index           =   39
         Left            =   -66705
         TabIndex        =   161
         Top             =   5760
         Width           =   300
      End
      Begin VB.Label LabRace 
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
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   40
         Left            =   -68820
         TabIndex        =   160
         Top             =   6075
         Width           =   300
      End
      Begin VB.Label labCar2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   40
         Left            =   -68205
         TabIndex        =   159
         Top             =   6075
         Width           =   330
      End
      Begin VB.Label labCar1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   40
         Left            =   -68535
         TabIndex        =   158
         Top             =   6075
         Width           =   330
      End
      Begin VB.Label labCar3 
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
         ForeColor       =   &H00404040&
         Height          =   285
         Index           =   40
         Left            =   -67905
         TabIndex        =   157
         Top             =   6075
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
         ForeColor       =   &H000000FF&
         Height          =   285
         Index           =   40
         Left            =   -66705
         TabIndex        =   156
         Top             =   6075
         Width           =   300
      End
      Begin VB.Label LabRace 
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
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   41
         Left            =   -68820
         TabIndex        =   155
         Top             =   6390
         Width           =   300
      End
      Begin VB.Label labCar2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   41
         Left            =   -68205
         TabIndex        =   154
         Top             =   6390
         Width           =   330
      End
      Begin VB.Label labCar1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   41
         Left            =   -68535
         TabIndex        =   153
         Top             =   6390
         Width           =   330
      End
      Begin VB.Label labCar3 
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
         ForeColor       =   &H00404040&
         Height          =   285
         Index           =   41
         Left            =   -67905
         TabIndex        =   152
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
         ForeColor       =   &H000000FF&
         Height          =   285
         Index           =   41
         Left            =   -66705
         TabIndex        =   151
         Top             =   6390
         Width           =   300
      End
      Begin VB.Label LabRace 
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
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   42
         Left            =   -68820
         TabIndex        =   150
         Top             =   6705
         Width           =   300
      End
      Begin VB.Label labCar2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   42
         Left            =   -68205
         TabIndex        =   149
         Top             =   6705
         Width           =   330
      End
      Begin VB.Label labCar1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   42
         Left            =   -68535
         TabIndex        =   148
         Top             =   6705
         Width           =   330
      End
      Begin VB.Label labCar3 
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
         ForeColor       =   &H00404040&
         Height          =   285
         Index           =   42
         Left            =   -67905
         TabIndex        =   147
         Top             =   6705
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
         ForeColor       =   &H000000FF&
         Height          =   285
         Index           =   42
         Left            =   -66705
         TabIndex        =   146
         Top             =   6705
         Width           =   300
      End
      Begin VB.Label LabRace 
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
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   43
         Left            =   -68820
         TabIndex        =   145
         Top             =   7020
         Width           =   300
      End
      Begin VB.Label labCar2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   43
         Left            =   -68205
         TabIndex        =   144
         Top             =   7020
         Width           =   330
      End
      Begin VB.Label labCar1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   43
         Left            =   -68535
         TabIndex        =   143
         Top             =   7020
         Width           =   330
      End
      Begin VB.Label labCar3 
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
         ForeColor       =   &H00404040&
         Height          =   285
         Index           =   43
         Left            =   -67905
         TabIndex        =   142
         Top             =   7020
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
         ForeColor       =   &H000000FF&
         Height          =   285
         Index           =   43
         Left            =   -66705
         TabIndex        =   141
         Top             =   7020
         Width           =   300
      End
      Begin VB.Label LabRace 
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
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   44
         Left            =   -68820
         TabIndex        =   140
         Top             =   7335
         Width           =   300
      End
      Begin VB.Label labCar2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   44
         Left            =   -68205
         TabIndex        =   139
         Top             =   7335
         Width           =   330
      End
      Begin VB.Label labCar1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   44
         Left            =   -68535
         TabIndex        =   138
         Top             =   7335
         Width           =   330
      End
      Begin VB.Label labCar3 
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
         ForeColor       =   &H00404040&
         Height          =   285
         Index           =   44
         Left            =   -67905
         TabIndex        =   137
         Top             =   7335
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
         ForeColor       =   &H000000FF&
         Height          =   285
         Index           =   44
         Left            =   -66705
         TabIndex        =   136
         Top             =   7335
         Width           =   300
      End
      Begin VB.Label LabRace 
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
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   45
         Left            =   -68820
         TabIndex        =   135
         Top             =   7650
         Width           =   300
      End
      Begin VB.Label labCar2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   45
         Left            =   -68205
         TabIndex        =   134
         Top             =   7650
         Width           =   330
      End
      Begin VB.Label labCar1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   45
         Left            =   -68535
         TabIndex        =   133
         Top             =   7650
         Width           =   330
      End
      Begin VB.Label labCar3 
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
         ForeColor       =   &H00404040&
         Height          =   285
         Index           =   45
         Left            =   -67905
         TabIndex        =   132
         Top             =   7650
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
         ForeColor       =   &H000000FF&
         Height          =   285
         Index           =   45
         Left            =   -66705
         TabIndex        =   131
         Top             =   7650
         Width           =   300
      End
      Begin VB.Label labnNiv 
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
         ForeColor       =   &H00404040&
         Height          =   285
         Index           =   23
         Left            =   -67605
         TabIndex        =   130
         Top             =   720
         Width           =   300
      End
      Begin VB.Label labnNiv 
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
         ForeColor       =   &H00404040&
         Height          =   285
         Index           =   24
         Left            =   -67605
         TabIndex        =   129
         Top             =   1035
         Width           =   300
      End
      Begin VB.Label labnNiv 
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
         ForeColor       =   &H00404040&
         Height          =   285
         Index           =   25
         Left            =   -67605
         TabIndex        =   128
         Top             =   1350
         Width           =   300
      End
      Begin VB.Label labnNiv 
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
         ForeColor       =   &H00404040&
         Height          =   285
         Index           =   26
         Left            =   -67605
         TabIndex        =   127
         Top             =   1665
         Width           =   300
      End
      Begin VB.Label labnNiv 
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
         ForeColor       =   &H00404040&
         Height          =   285
         Index           =   27
         Left            =   -67605
         TabIndex        =   126
         Top             =   1980
         Width           =   300
      End
      Begin VB.Label labnNiv 
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
         ForeColor       =   &H00404040&
         Height          =   285
         Index           =   28
         Left            =   -67605
         TabIndex        =   125
         Top             =   2295
         Width           =   300
      End
      Begin VB.Label labnNiv 
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
         ForeColor       =   &H00404040&
         Height          =   285
         Index           =   29
         Left            =   -67605
         TabIndex        =   124
         Top             =   2610
         Width           =   300
      End
      Begin VB.Label labnNiv 
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
         ForeColor       =   &H00404040&
         Height          =   285
         Index           =   30
         Left            =   -67605
         TabIndex        =   123
         Top             =   2925
         Width           =   300
      End
      Begin VB.Label labnNiv 
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
         ForeColor       =   &H00404040&
         Height          =   285
         Index           =   31
         Left            =   -67605
         TabIndex        =   122
         Top             =   3240
         Width           =   300
      End
      Begin VB.Label labnNiv 
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
         ForeColor       =   &H00404040&
         Height          =   285
         Index           =   32
         Left            =   -67605
         TabIndex        =   121
         Top             =   3570
         Width           =   300
      End
      Begin VB.Label labnNiv 
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
         ForeColor       =   &H00404040&
         Height          =   285
         Index           =   33
         Left            =   -67605
         TabIndex        =   120
         Top             =   3870
         Width           =   300
      End
      Begin VB.Label labnNiv 
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
         ForeColor       =   &H00404040&
         Height          =   285
         Index           =   34
         Left            =   -67605
         TabIndex        =   119
         Top             =   4185
         Width           =   300
      End
      Begin VB.Label labnNiv 
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
         ForeColor       =   &H00404040&
         Height          =   285
         Index           =   35
         Left            =   -67605
         TabIndex        =   118
         Top             =   4500
         Width           =   300
      End
      Begin VB.Label labnNiv 
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
         ForeColor       =   &H00404040&
         Height          =   285
         Index           =   36
         Left            =   -67605
         TabIndex        =   117
         Top             =   4815
         Width           =   300
      End
      Begin VB.Label labnNiv 
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
         ForeColor       =   &H00404040&
         Height          =   285
         Index           =   37
         Left            =   -67605
         TabIndex        =   116
         Top             =   5130
         Width           =   300
      End
      Begin VB.Label labnNiv 
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
         ForeColor       =   &H00404040&
         Height          =   285
         Index           =   38
         Left            =   -67605
         TabIndex        =   115
         Top             =   5445
         Width           =   300
      End
      Begin VB.Label labnNiv 
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
         ForeColor       =   &H00404040&
         Height          =   285
         Index           =   39
         Left            =   -67605
         TabIndex        =   114
         Top             =   5760
         Width           =   300
      End
      Begin VB.Label labnNiv 
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
         ForeColor       =   &H00404040&
         Height          =   285
         Index           =   40
         Left            =   -67605
         TabIndex        =   113
         Top             =   6075
         Width           =   300
      End
      Begin VB.Label labnNiv 
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
         ForeColor       =   &H00404040&
         Height          =   285
         Index           =   41
         Left            =   -67605
         TabIndex        =   112
         Top             =   6390
         Width           =   300
      End
      Begin VB.Label labnNiv 
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
         ForeColor       =   &H00404040&
         Height          =   285
         Index           =   42
         Left            =   -67605
         TabIndex        =   111
         Top             =   6705
         Width           =   300
      End
      Begin VB.Label labnNiv 
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
         ForeColor       =   &H00404040&
         Height          =   285
         Index           =   43
         Left            =   -67605
         TabIndex        =   110
         Top             =   7020
         Width           =   300
      End
      Begin VB.Label labnNiv 
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
         ForeColor       =   &H00404040&
         Height          =   285
         Index           =   44
         Left            =   -67605
         TabIndex        =   109
         Top             =   7335
         Width           =   300
      End
      Begin VB.Label labnNiv 
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
         ForeColor       =   &H00404040&
         Height          =   285
         Index           =   45
         Left            =   -67605
         TabIndex        =   108
         Top             =   7650
         Width           =   300
      End
      Begin VB.Label Label67 
         Caption         =   "projectiles"
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
         Left            =   -72075
         TabIndex        =   107
         Top             =   4095
         Width           =   795
      End
      Begin VB.Label Label86 
         Caption         =   "Attaque de base"
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
         Left            =   -74055
         TabIndex        =   106
         Top             =   4095
         Width           =   1545
      End
      Begin VB.Label labnobj 
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
         ForeColor       =   &H00FF8080&
         Height          =   285
         Index           =   0
         Left            =   -71640
         TabIndex        =   105
         Top             =   720
         Width           =   300
      End
      Begin VB.Label labnSyn 
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
         ForeColor       =   &H00FF8080&
         Height          =   285
         Index           =   0
         Left            =   -71340
         TabIndex        =   104
         Top             =   720
         Width           =   300
      End
      Begin VB.Label labnobj 
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
         ForeColor       =   &H00FF8080&
         Height          =   285
         Index           =   1
         Left            =   -71640
         TabIndex        =   103
         Top             =   1035
         Width           =   300
      End
      Begin VB.Label labnobj 
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
         ForeColor       =   &H00FF8080&
         Height          =   285
         Index           =   2
         Left            =   -71640
         TabIndex        =   102
         Top             =   1350
         Width           =   300
      End
      Begin VB.Label labnobj 
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
         ForeColor       =   &H00FF8080&
         Height          =   285
         Index           =   3
         Left            =   -71640
         TabIndex        =   101
         Top             =   1665
         Width           =   300
      End
      Begin VB.Label labnobj 
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
         ForeColor       =   &H00FF8080&
         Height          =   285
         Index           =   4
         Left            =   -71640
         TabIndex        =   100
         Top             =   1980
         Width           =   300
      End
      Begin VB.Label labnobj 
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
         ForeColor       =   &H00FF8080&
         Height          =   285
         Index           =   5
         Left            =   -71640
         TabIndex        =   99
         Top             =   2295
         Width           =   300
      End
      Begin VB.Label labnobj 
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
         ForeColor       =   &H00FF8080&
         Height          =   285
         Index           =   6
         Left            =   -71640
         TabIndex        =   98
         Top             =   2610
         Width           =   300
      End
      Begin VB.Label labnobj 
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
         ForeColor       =   &H00FF8080&
         Height          =   285
         Index           =   7
         Left            =   -71640
         TabIndex        =   97
         Top             =   2925
         Width           =   300
      End
      Begin VB.Label labnobj 
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
         ForeColor       =   &H00FF8080&
         Height          =   285
         Index           =   8
         Left            =   -71640
         TabIndex        =   96
         Top             =   3240
         Width           =   300
      End
      Begin VB.Label labnobj 
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
         ForeColor       =   &H00FF8080&
         Height          =   285
         Index           =   9
         Left            =   -71640
         TabIndex        =   95
         Top             =   3555
         Width           =   300
      End
      Begin VB.Label labnobj 
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
         ForeColor       =   &H00FF8080&
         Height          =   285
         Index           =   10
         Left            =   -71640
         TabIndex        =   94
         Top             =   3870
         Width           =   300
      End
      Begin VB.Label labnobj 
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
         ForeColor       =   &H00FF8080&
         Height          =   285
         Index           =   11
         Left            =   -71640
         TabIndex        =   93
         Top             =   4185
         Width           =   300
      End
      Begin VB.Label labnobj 
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
         ForeColor       =   &H00FF8080&
         Height          =   285
         Index           =   12
         Left            =   -71640
         TabIndex        =   92
         Top             =   4500
         Width           =   300
      End
      Begin VB.Label labnobj 
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
         ForeColor       =   &H00FF8080&
         Height          =   285
         Index           =   13
         Left            =   -71640
         TabIndex        =   91
         Top             =   4815
         Width           =   300
      End
      Begin VB.Label labnobj 
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
         ForeColor       =   &H00FF8080&
         Height          =   285
         Index           =   14
         Left            =   -71640
         TabIndex        =   90
         Top             =   5130
         Width           =   300
      End
      Begin VB.Label labnobj 
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
         ForeColor       =   &H00FF8080&
         Height          =   285
         Index           =   15
         Left            =   -71640
         TabIndex        =   89
         Top             =   5445
         Width           =   300
      End
      Begin VB.Label labnobj 
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
         ForeColor       =   &H00FF8080&
         Height          =   285
         Index           =   16
         Left            =   -71640
         TabIndex        =   88
         Top             =   5760
         Width           =   300
      End
      Begin VB.Label labnobj 
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
         ForeColor       =   &H00FF8080&
         Height          =   285
         Index           =   17
         Left            =   -71640
         TabIndex        =   87
         Top             =   6075
         Width           =   300
      End
      Begin VB.Label labnobj 
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
         ForeColor       =   &H00FF8080&
         Height          =   285
         Index           =   18
         Left            =   -71640
         TabIndex        =   86
         Top             =   6390
         Width           =   300
      End
      Begin VB.Label labnobj 
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
         ForeColor       =   &H00FF8080&
         Height          =   285
         Index           =   19
         Left            =   -71640
         TabIndex        =   85
         Top             =   6705
         Width           =   300
      End
      Begin VB.Label labnobj 
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
         ForeColor       =   &H00FF8080&
         Height          =   285
         Index           =   20
         Left            =   -71640
         TabIndex        =   84
         Top             =   7020
         Width           =   300
      End
      Begin VB.Label labnobj 
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
         ForeColor       =   &H00FF8080&
         Height          =   285
         Index           =   21
         Left            =   -71640
         TabIndex        =   83
         Top             =   7335
         Width           =   300
      End
      Begin VB.Label labnobj 
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
         ForeColor       =   &H00FF8080&
         Height          =   285
         Index           =   22
         Left            =   -71640
         TabIndex        =   82
         Top             =   7650
         Width           =   300
      End
      Begin VB.Label labnobj 
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
         ForeColor       =   &H00FF8080&
         Height          =   285
         Index           =   23
         Left            =   -67320
         TabIndex        =   81
         Top             =   720
         Width           =   300
      End
      Begin VB.Label labnobj 
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
         ForeColor       =   &H00FF8080&
         Height          =   285
         Index           =   24
         Left            =   -67320
         TabIndex        =   80
         Top             =   1035
         Width           =   300
      End
      Begin VB.Label labnobj 
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
         ForeColor       =   &H00FF8080&
         Height          =   285
         Index           =   25
         Left            =   -67320
         TabIndex        =   79
         Top             =   1350
         Width           =   300
      End
      Begin VB.Label labnobj 
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
         ForeColor       =   &H00FF8080&
         Height          =   285
         Index           =   26
         Left            =   -67320
         TabIndex        =   78
         Top             =   1665
         Width           =   300
      End
      Begin VB.Label labnobj 
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
         ForeColor       =   &H00FF8080&
         Height          =   285
         Index           =   27
         Left            =   -67320
         TabIndex        =   77
         Top             =   1980
         Width           =   300
      End
      Begin VB.Label labnobj 
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
         ForeColor       =   &H00FF8080&
         Height          =   285
         Index           =   28
         Left            =   -67320
         TabIndex        =   76
         Top             =   2295
         Width           =   300
      End
      Begin VB.Label labnobj 
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
         ForeColor       =   &H00FF8080&
         Height          =   285
         Index           =   29
         Left            =   -67320
         TabIndex        =   75
         Top             =   2610
         Width           =   300
      End
      Begin VB.Label labnobj 
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
         ForeColor       =   &H00FF8080&
         Height          =   285
         Index           =   30
         Left            =   -67320
         TabIndex        =   74
         Top             =   2925
         Width           =   300
      End
      Begin VB.Label labnobj 
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
         ForeColor       =   &H00FF8080&
         Height          =   285
         Index           =   31
         Left            =   -67320
         TabIndex        =   73
         Top             =   3240
         Width           =   300
      End
      Begin VB.Label labnobj 
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
         ForeColor       =   &H00FF8080&
         Height          =   285
         Index           =   32
         Left            =   -67320
         TabIndex        =   72
         Top             =   3555
         Width           =   300
      End
      Begin VB.Label labnobj 
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
         ForeColor       =   &H00FF8080&
         Height          =   285
         Index           =   33
         Left            =   -67320
         TabIndex        =   71
         Top             =   3870
         Width           =   300
      End
      Begin VB.Label labnobj 
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
         ForeColor       =   &H00FF8080&
         Height          =   285
         Index           =   34
         Left            =   -67320
         TabIndex        =   70
         Top             =   4185
         Width           =   300
      End
      Begin VB.Label labnobj 
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
         ForeColor       =   &H00FF8080&
         Height          =   285
         Index           =   35
         Left            =   -67320
         TabIndex        =   69
         Top             =   4500
         Width           =   300
      End
      Begin VB.Label labnobj 
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
         ForeColor       =   &H00FF8080&
         Height          =   285
         Index           =   36
         Left            =   -67320
         TabIndex        =   68
         Top             =   4815
         Width           =   300
      End
      Begin VB.Label labnobj 
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
         ForeColor       =   &H00FF8080&
         Height          =   285
         Index           =   37
         Left            =   -67320
         TabIndex        =   67
         Top             =   5130
         Width           =   300
      End
      Begin VB.Label labnobj 
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
         ForeColor       =   &H00FF8080&
         Height          =   285
         Index           =   38
         Left            =   -67320
         TabIndex        =   66
         Top             =   5445
         Width           =   300
      End
      Begin VB.Label labnobj 
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
         ForeColor       =   &H00FF8080&
         Height          =   285
         Index           =   39
         Left            =   -67320
         TabIndex        =   65
         Top             =   5760
         Width           =   300
      End
      Begin VB.Label labnobj 
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
         ForeColor       =   &H00FF8080&
         Height          =   285
         Index           =   40
         Left            =   -67320
         TabIndex        =   64
         Top             =   6075
         Width           =   300
      End
      Begin VB.Label labnobj 
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
         ForeColor       =   &H00FF8080&
         Height          =   285
         Index           =   41
         Left            =   -67320
         TabIndex        =   63
         Top             =   6390
         Width           =   300
      End
      Begin VB.Label labnobj 
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
         ForeColor       =   &H00FF8080&
         Height          =   285
         Index           =   42
         Left            =   -67320
         TabIndex        =   62
         Top             =   6705
         Width           =   300
      End
      Begin VB.Label labnobj 
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
         ForeColor       =   &H00FF8080&
         Height          =   285
         Index           =   43
         Left            =   -67320
         TabIndex        =   61
         Top             =   7020
         Width           =   300
      End
      Begin VB.Label labnobj 
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
         ForeColor       =   &H00FF8080&
         Height          =   285
         Index           =   44
         Left            =   -67320
         TabIndex        =   60
         Top             =   7335
         Width           =   300
      End
      Begin VB.Label labnobj 
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
         ForeColor       =   &H00FF8080&
         Height          =   285
         Index           =   45
         Left            =   -67320
         TabIndex        =   59
         Top             =   7650
         Width           =   300
      End
      Begin VB.Label labnSyn 
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
         ForeColor       =   &H00FF8080&
         Height          =   285
         Index           =   1
         Left            =   -71340
         TabIndex        =   58
         Top             =   1035
         Width           =   300
      End
      Begin VB.Label labnSyn 
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
         ForeColor       =   &H00FF8080&
         Height          =   285
         Index           =   2
         Left            =   -71340
         TabIndex        =   57
         Top             =   1350
         Width           =   300
      End
      Begin VB.Label labnSyn 
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
         ForeColor       =   &H00FF8080&
         Height          =   285
         Index           =   3
         Left            =   -71340
         TabIndex        =   56
         Top             =   1665
         Width           =   300
      End
      Begin VB.Label labnSyn 
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
         ForeColor       =   &H00FF8080&
         Height          =   285
         Index           =   4
         Left            =   -71340
         TabIndex        =   55
         Top             =   1980
         Width           =   300
      End
      Begin VB.Label labnSyn 
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
         ForeColor       =   &H00FF8080&
         Height          =   285
         Index           =   5
         Left            =   -71340
         TabIndex        =   54
         Top             =   2295
         Width           =   300
      End
      Begin VB.Label labnSyn 
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
         ForeColor       =   &H00FF8080&
         Height          =   285
         Index           =   6
         Left            =   -71340
         TabIndex        =   53
         Top             =   2610
         Width           =   300
      End
      Begin VB.Label labnSyn 
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
         ForeColor       =   &H00FF8080&
         Height          =   285
         Index           =   7
         Left            =   -71340
         TabIndex        =   52
         Top             =   2925
         Width           =   300
      End
      Begin VB.Label labnSyn 
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
         ForeColor       =   &H00FF8080&
         Height          =   285
         Index           =   8
         Left            =   -71340
         TabIndex        =   51
         Top             =   3240
         Width           =   300
      End
      Begin VB.Label labnSyn 
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
         ForeColor       =   &H00FF8080&
         Height          =   285
         Index           =   9
         Left            =   -71340
         TabIndex        =   50
         Top             =   3555
         Width           =   300
      End
      Begin VB.Label labnSyn 
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
         ForeColor       =   &H00FF8080&
         Height          =   285
         Index           =   10
         Left            =   -71340
         TabIndex        =   49
         Top             =   3870
         Width           =   300
      End
      Begin VB.Label labnSyn 
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
         ForeColor       =   &H00FF8080&
         Height          =   285
         Index           =   11
         Left            =   -71340
         TabIndex        =   48
         Top             =   4185
         Width           =   300
      End
      Begin VB.Label labnSyn 
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
         ForeColor       =   &H00FF8080&
         Height          =   285
         Index           =   12
         Left            =   -71340
         TabIndex        =   47
         Top             =   4500
         Width           =   300
      End
      Begin VB.Label labnSyn 
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
         ForeColor       =   &H00FF8080&
         Height          =   285
         Index           =   13
         Left            =   -71340
         TabIndex        =   46
         Top             =   4815
         Width           =   300
      End
      Begin VB.Label labnSyn 
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
         ForeColor       =   &H00FF8080&
         Height          =   285
         Index           =   14
         Left            =   -71340
         TabIndex        =   45
         Top             =   5130
         Width           =   300
      End
      Begin VB.Label labnSyn 
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
         ForeColor       =   &H00FF8080&
         Height          =   285
         Index           =   15
         Left            =   -71340
         TabIndex        =   44
         Top             =   5445
         Width           =   300
      End
      Begin VB.Label labnSyn 
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
         ForeColor       =   &H00FF8080&
         Height          =   285
         Index           =   16
         Left            =   -71340
         TabIndex        =   43
         Top             =   5760
         Width           =   300
      End
      Begin VB.Label labnSyn 
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
         ForeColor       =   &H00FF8080&
         Height          =   285
         Index           =   17
         Left            =   -71340
         TabIndex        =   42
         Top             =   6075
         Width           =   300
      End
      Begin VB.Label labnSyn 
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
         ForeColor       =   &H00FF8080&
         Height          =   285
         Index           =   18
         Left            =   -71340
         TabIndex        =   41
         Top             =   6390
         Width           =   300
      End
      Begin VB.Label labnSyn 
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
         ForeColor       =   &H00FF8080&
         Height          =   285
         Index           =   19
         Left            =   -71340
         TabIndex        =   40
         Top             =   6705
         Width           =   300
      End
      Begin VB.Label labnSyn 
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
         ForeColor       =   &H00FF8080&
         Height          =   285
         Index           =   20
         Left            =   -71340
         TabIndex        =   39
         Top             =   7020
         Width           =   300
      End
      Begin VB.Label labnSyn 
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
         ForeColor       =   &H00FF8080&
         Height          =   285
         Index           =   21
         Left            =   -71340
         TabIndex        =   38
         Top             =   7335
         Width           =   300
      End
      Begin VB.Label labnSyn 
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
         ForeColor       =   &H00FF8080&
         Height          =   285
         Index           =   22
         Left            =   -71340
         TabIndex        =   37
         Top             =   7650
         Width           =   300
      End
      Begin VB.Label labnSyn 
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
         ForeColor       =   &H00FF8080&
         Height          =   285
         Index           =   23
         Left            =   -67020
         TabIndex        =   36
         Top             =   720
         Width           =   300
      End
      Begin VB.Label labnSyn 
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
         ForeColor       =   &H00FF8080&
         Height          =   285
         Index           =   24
         Left            =   -67020
         TabIndex        =   35
         Top             =   1035
         Width           =   300
      End
      Begin VB.Label labnSyn 
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
         ForeColor       =   &H00FF8080&
         Height          =   285
         Index           =   25
         Left            =   -67020
         TabIndex        =   34
         Top             =   1350
         Width           =   300
      End
      Begin VB.Label labnSyn 
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
         ForeColor       =   &H00FF8080&
         Height          =   285
         Index           =   26
         Left            =   -67020
         TabIndex        =   33
         Top             =   1665
         Width           =   300
      End
      Begin VB.Label labnSyn 
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
         ForeColor       =   &H00FF8080&
         Height          =   285
         Index           =   27
         Left            =   -67020
         TabIndex        =   32
         Top             =   1980
         Width           =   300
      End
      Begin VB.Label labnSyn 
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
         ForeColor       =   &H00FF8080&
         Height          =   285
         Index           =   28
         Left            =   -67020
         TabIndex        =   31
         Top             =   2295
         Width           =   300
      End
      Begin VB.Label labnSyn 
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
         ForeColor       =   &H00FF8080&
         Height          =   285
         Index           =   29
         Left            =   -67020
         TabIndex        =   30
         Top             =   2610
         Width           =   300
      End
      Begin VB.Label labnSyn 
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
         ForeColor       =   &H00FF8080&
         Height          =   285
         Index           =   30
         Left            =   -67020
         TabIndex        =   29
         Top             =   2925
         Width           =   300
      End
      Begin VB.Label labnSyn 
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
         ForeColor       =   &H00FF8080&
         Height          =   285
         Index           =   31
         Left            =   -67020
         TabIndex        =   28
         Top             =   3240
         Width           =   300
      End
      Begin VB.Label labnSyn 
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
         ForeColor       =   &H00FF8080&
         Height          =   285
         Index           =   32
         Left            =   -67020
         TabIndex        =   27
         Top             =   3555
         Width           =   300
      End
      Begin VB.Label labnSyn 
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
         ForeColor       =   &H00FF8080&
         Height          =   285
         Index           =   33
         Left            =   -67020
         TabIndex        =   26
         Top             =   3870
         Width           =   300
      End
      Begin VB.Label labnSyn 
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
         ForeColor       =   &H00FF8080&
         Height          =   285
         Index           =   34
         Left            =   -67020
         TabIndex        =   25
         Top             =   4185
         Width           =   300
      End
      Begin VB.Label labnSyn 
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
         ForeColor       =   &H00FF8080&
         Height          =   285
         Index           =   35
         Left            =   -67020
         TabIndex        =   24
         Top             =   4500
         Width           =   300
      End
      Begin VB.Label labnSyn 
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
         ForeColor       =   &H00FF8080&
         Height          =   285
         Index           =   36
         Left            =   -67020
         TabIndex        =   23
         Top             =   4815
         Width           =   300
      End
      Begin VB.Label labnSyn 
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
         ForeColor       =   &H00FF8080&
         Height          =   285
         Index           =   37
         Left            =   -67020
         TabIndex        =   22
         Top             =   5130
         Width           =   300
      End
      Begin VB.Label labnSyn 
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
         ForeColor       =   &H00FF8080&
         Height          =   285
         Index           =   38
         Left            =   -67020
         TabIndex        =   21
         Top             =   5445
         Width           =   300
      End
      Begin VB.Label labnSyn 
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
         ForeColor       =   &H00FF8080&
         Height          =   285
         Index           =   39
         Left            =   -67020
         TabIndex        =   20
         Top             =   5760
         Width           =   300
      End
      Begin VB.Label labnSyn 
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
         ForeColor       =   &H00FF8080&
         Height          =   285
         Index           =   40
         Left            =   -67020
         TabIndex        =   19
         Top             =   6075
         Width           =   300
      End
      Begin VB.Label labnSyn 
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
         ForeColor       =   &H00FF8080&
         Height          =   285
         Index           =   41
         Left            =   -67020
         TabIndex        =   18
         Top             =   6390
         Width           =   300
      End
      Begin VB.Label labnSyn 
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
         ForeColor       =   &H00FF8080&
         Height          =   285
         Index           =   42
         Left            =   -67020
         TabIndex        =   17
         Top             =   6705
         Width           =   300
      End
      Begin VB.Label labnSyn 
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
         ForeColor       =   &H00FF8080&
         Height          =   285
         Index           =   43
         Left            =   -67020
         TabIndex        =   16
         Top             =   7020
         Width           =   300
      End
      Begin VB.Label labnSyn 
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
         ForeColor       =   &H00FF8080&
         Height          =   285
         Index           =   44
         Left            =   -67020
         TabIndex        =   15
         Top             =   7335
         Width           =   300
      End
      Begin VB.Label labnSyn 
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
         ForeColor       =   &H00FF8080&
         Height          =   285
         Index           =   45
         Left            =   -67020
         TabIndex        =   14
         Top             =   7650
         Width           =   300
      End
      Begin VB.Label Label56 
         Caption         =   "Cararctère    Niv objet  Total"
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
         Left            =   -68460
         TabIndex        =   13
         Top             =   405
         Width           =   2040
      End
   End
End
Attribute VB_Name = "frmGestobj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const COL_NOM_OBJET = 0
Const COL_COUT_TOTAL = 1
Const COL_POIDS_BASE = 2
Const COL_CLASSE_ARMURE = 3
Const COL_TYPE_CA = 4
Const COL_ECHEC_SORT_PROFANE = 5
Const COL_PENALITE_ARMURE = 6
Const COL_BONUS_DEX_MAX = 7
Const COL_PORTEE = 8
Const COL_GROUPE_OBJET = 9
Const COL_TAILLE = 10
Const COL_DEGAT_TYPE = 11
Const COL_DEGAT_DES = 12
Const COL_ZONE_CRITIQUE = 13
Const COL_MULT_CRIT = 14
Const COL_DESCRIPTION = 15
Const COL_SOLIDITE = 16
Const COL_RESISTANCE = 17
Const COL_CHARGE = 18

Dim m_bmodif As Boolean
Dim m_bCharge As Boolean
Dim m_tabButton() As S_TABVAL
Sub rempTree()
Dim key1 As String, key2 As String, key3 As String, key4 As String
Dim i As Integer, j As Integer, k As Integer
Dim rst As Recordset
Dim nSort As Integer
Dim sTab() As S_TABVAL
Dim tTab() As S_TABVAL

With Tree1
    .Nodes.Clear
    ''CARACTERISTIQUE
    key1 = "K1_" & obj_BONUS_CARACTERISTIQUE
    .Nodes.Add , , key1, obj_BONUS_CARACTERISTIQUE
    ''CA
    key1 = "K1_" & obj_BONUS_CLASSE_ARMURE
    .Nodes.Add , , key1, obj_BONUS_CLASSE_ARMURE
    ''BONUS JET ATTAQUE
    key1 = "K1_" & obj_BONUS_ATTAQUE
    .Nodes.Add , , key1, obj_BONUS_ATTAQUE
    ''BONUS_DON
    key1 = "K1_" & obj_BONUS_DON
    .Nodes.Add , , key1, obj_BONUS_DON
    ''BONUS AU DEGATS DOMMAGE
    key1 = "K1_" & obj_BONUS_DOMMAGE
    .Nodes.Add , , key1, obj_BONUS_DOMMAGE
    ''BONUS AU JET DE SAUVEGARDE
    key1 = "K1_" & obj_BONUS_JET_SAUVEGARDE
    .Nodes.Add , , key1, obj_BONUS_JET_SAUVEGARDE
    ''BONUS AU COMPETENCE
    key1 = "K1_" & obj_BONUS_COMPETENCE
    .Nodes.Add , , key1, obj_BONUS_COMPETENCE
    ''BONUS AU DEPLACEMENT
    key1 = "K1_" & obj_BONUS_DEPLACEMENT
    .Nodes.Add , , key1, obj_BONUS_DEPLACEMENT
    ''RESISTANCE A LA MAGIE
    key1 = "K1_" & obj_RESISTANCE_A_LA_MAGIE
    .Nodes.Add , , key1, obj_RESISTANCE_A_LA_MAGIE
    ''RESISTANCE AU FEU
    key1 = "K1_" & obj_RESISTANCE_AU_FEU
    .Nodes.Add , , key1, obj_RESISTANCE_AU_FEU
    ''RESISTANCE AU FROID
    key1 = "K1_" & obj_RESISTANCE_AU_FROID
    .Nodes.Add , , key1, obj_RESISTANCE_AU_FROID
     ''RESISTANCE A L ELECTRICITE
    key1 = "K1_" & obj_RESISTANCE_A_L_ELECTRICITE
    .Nodes.Add , , key1, obj_RESISTANCE_A_L_ELECTRICITE
    ''RESISTANCE A L ACIDE
    key1 = "K1_" & obj_RESISTANCE_A_L_ACIDE
    .Nodes.Add , , key1, obj_RESISTANCE_A_L_ACIDE
     ''RESISTANCE AU SON
    key1 = "K1_" & obj_RESISTANCE_AU_SON
    .Nodes.Add , , key1, obj_RESISTANCE_AU_SON
     ''RESISTANCE AUX DEGATS
    key1 = "K1_" & obj_RESISTANCE_AUX_DEGATS
    .Nodes.Add , , key1, obj_RESISTANCE_AUX_DEGATS
     ''ARC DE FORCE
    key1 = "K1_" & obj_ARC_DE_PUISSANCE
    .Nodes.Add , , key1, obj_ARC_DE_PUISSANCE
    ''BONUS ARC
    key1 = "K1_" & obj_BONUS_ARC
    .Nodes.Add , , key1, obj_BONUS_ARC
    ''BONUS JET ATTAQUE
    key1 = "K1_" & obj_BONUS_JET_ATTAQUE
    .Nodes.Add , , key1, obj_BONUS_JET_ATTAQUE
    ''BONUS ARMES NATURELLES
    key1 = "K1_" & obj_BONUS_ARME_NATURELLE
    .Nodes.Add , , key1, obj_BONUS_ARME_NATURELLE
    ''REJET DES MORTS VIVANTS
    key1 = "K1_" & obj_REJET_MORT_VIVANT_ACCRU
    .Nodes.Add , , key1, obj_REJET_MORT_VIVANT_ACCRU
    '' ARCANES
    key1 = "K1_" & obj_ARCANES
    .Nodes.Add , , key1, obj_ARCANES
    ''BONUS A L'INITIATIVE
    key1 = "K1_" & obj_BONUS_INITIATIVE
    .Nodes.Add , , key1, obj_BONUS_INITIATIVE
    ''SORT A CHARGES
    key1 = "K1_" & obj_LANCER_SORT
    .Nodes.Add , , key1, obj_LANCER_SORT
  End With
End Sub

Private Sub btnAjout_Click()
Dim i As Integer
Dim j As Long
Dim n As Integer
Dim ps1 As Integer, ps2 As Integer
Dim strItem As String
Dim prop1 As String, prop2 As String, prop3 As String, val1 As String
Dim bexist As Boolean

  j = Tree1.SelectedItem.Index
    If Tree1.Nodes.Item(j).Selected = True Then
      strItem = Tree1.Nodes.Item(j).Key
      If Right(strItem, 1) <> "_" Then strItem = strItem & "_"
      prop1 = ""
      prop2 = ""
      prop3 = ""
      val1 = ""
     
      ps1 = InStr(1, strItem, "K1_", vbBinaryCompare)
      If ps1 > 0 Then
        ps1 = ps1 + 3
        ps2 = InStr(ps1, strItem, "_", vbBinaryCompare)
        If ps2 > 0 Then
          prop1 = Mid(strItem, ps1, ps2 - ps1)
        End If
      End If
      
      ps1 = InStr(1, strItem, "K2_", vbBinaryCompare)
      If ps1 > 0 Then
        ps1 = ps1 + 3
        ps2 = InStr(ps1, strItem, "_", vbBinaryCompare)
        If ps2 > 0 Then
          prop2 = Mid(strItem, ps1, ps2 - ps1)
        End If
      End If
      
      ps1 = InStr(1, strItem, "K3_", vbBinaryCompare)
      If ps1 > 0 Then
        ps1 = ps1 + 3
        ps2 = InStr(ps1, strItem, "_", vbBinaryCompare)
        If ps2 > 0 Then
          prop3 = Mid(strItem, ps1, ps2 - ps1)
        End If
      End If

      ps1 = InStr(1, strItem, "K4_", vbBinaryCompare)
      If ps1 > 0 Then
        ps1 = ps1 + 3
        ps2 = InStr(ps1, strItem, "_", vbBinaryCompare)
        If ps2 > 0 Then
          val1 = Mid(strItem, ps1, ps2 - ps1)
          If Val(val1) < 0 Then
            val1 = val1
          ElseIf Val(val1) > 0 Then
            val1 = "+" & val1
          End If
        End If
      End If
      
      If prop1 <> "" Then
        If (objet_estarme(CmbGroupe) = False) And (prop1 = obj_BONUS_DOMMAGE Or prop1 = obj_BONUS_ATTAQUE) Then
          If Objet_Est_Projectile(CmbGroupe) = False Then
            Exit Sub
          End If
        End If
        
        vsPropri.AddItem prop1 & " " & prop2 & " " & prop3 & " " & val1 & vbTab & prop1 & vbTab & prop2 & vbTab & prop3 & vbTab & val1
        Modification
      End If
      Exit Sub
    End If
  
End Sub

Private Sub btnAjouter_Click()
  'ajout
  Efface Me
  vsPropri.Rows = 0
  Set_Focus fldstrnomObjet
End Sub

Private Sub btnAnnuler_Click()
  Fermer
End Sub

Private Sub btnEnlever_Click()
Dim i As Integer
  
  For i = 0 To vsPropri.Rows - 1
    If vsPropri.IsSelected(i) = True Then
      vsPropri.RemoveItem i
      Modification
      Exit For
    End If
  Next i
End Sub

Private Sub btnEnregistrer_Click()
  Enregistre
End Sub

Private Sub btnFermer_Click()
  Fermer
End Sub

Private Sub btnRenommer_Click()
Dim n As Integer
Dim rst As Recordset
  
  If vsObjets.SelectedRows < 1 Then
    Exit Sub
  End If
  
  n = vsObjets.SelectedRow(0)
  frmRenommer.fldstrNouveauNom = vsObjets.TextMatrix(n, COL_NOM_OBJET)
  frmRenommer.fldstrNouveauNom.ToolTipText = "Permet de renommer le nom del'objet, (retour en arriere possible)"
  frmRenommer.Show vbModal
  If frmRenommer.m_strnouveaunom <> "" Then
    If UCase(frmRenommer.m_strnouveaunom) <> UCase(vsObjets.TextMatrix(n, COL_NOM_OBJET)) Then
      Set rst = g_bdPerso.OpenRecordset("select * from objets where nomobjet='" & FormatStrSQL(frmRenommer.m_strnouveaunom) & "'", dbOpenDynaset)
      If Not rst.EOF Then
        MsgBox "Impossible de renommer vu que ce nouveau nom d'objet existe déjà", vbExclamation, g_strNomApplication
        rst.Close
        Exit Sub
      Else
        rst.Close
      End If
    End If
    
    g_bdPerso.Execute "update objets set nomobjet='" & FormatStrSQL(frmRenommer.m_strnouveaunom) & "' where nomobjet='" & FormatStrSQL(vsObjets.TextMatrix(n, COL_NOM_OBJET)) & "'", dbFailOnError
    g_bdPerso.Execute "update objetsPropriete set nomobjet='" & FormatStrSQL(frmRenommer.m_strnouveaunom) & "' where nomobjet='" & FormatStrSQL(vsObjets.TextMatrix(n, COL_NOM_OBJET)) & "'", dbFailOnError
    g_bdPerso.Execute "update equipement set nomobjet='" & FormatStrSQL(frmRenommer.m_strnouveaunom) & "' where nomobjet='" & FormatStrSQL(vsObjets.TextMatrix(n, COL_NOM_OBJET)) & "'", dbFailOnError
    vsObjets.TextMatrix(n, COL_NOM_OBJET) = frmRenommer.m_strnouveaunom
  End If
End Sub

Private Sub btnSupprimer_Click()
  Supprimer
End Sub

Private Sub cmbClasseArmure_Change()

  Modification

End Sub

Private Sub cmbClasseArmure_Click()

Modification

End Sub

Private Sub cmbClasseArmure_GotFocus()
  Selec
End Sub

Private Sub CmbDegatsDes_Change()

  Modification

End Sub

Private Sub CmbDegatsDes_Click()

  Modification

End Sub

Private Sub CmbDegatsDes_GotFocus()
  Selec
End Sub

Private Sub CmbGroupe_Click()
  Modification

End Sub

Private Sub CmbGroupe_GotFocus()
  Selec
End Sub

Private Sub CmbMultipCrit_Change()

  Modification

End Sub

Private Sub CmbMultipCrit_Click()

  Modification

End Sub

Private Sub CmbGroupe_Change()

  Modification
End Sub

Private Sub CmbMultipCrit_GotFocus()
  Selec
End Sub

Private Sub CmbTailles_Click()

  Modification
End Sub

Private Sub cmbTailles_GotFocus()

  Selec
End Sub
Private Sub cmbTailles_LostFocus()

  StatusBar1.Panels(1).Text = ""

End Sub
Sub Modification()

  
  If m_bCharge = False Then
    btnAjouter.Visible = False
    btnSupprimer.Visible = False
    btnEnregistrer.Visible = True
    btnAnnuler.Visible = True
    btnRenommer.Visible = False
    BtnFermer.Visible = False
  End If
  AjusteBouton Me, m_tabButton
  m_bmodif = True
End Sub
Private Function NonModifier() As Boolean
  
  NonModifier = False
  If m_bmodif = True Then
    Select Case MsgBox("Voulez vous enregistrer les modifications apportées à cet objet", vbYesNoCancel Or vbInformation, Me.Caption)
      Case vbYes
        If Enregistre() = False Then Exit Function
      Case vbCancel
        Call vsObjets_SelChange
        Exit Function
    End Select
    m_bmodif = False
  End If
  NonModifier = True
End Function
Function Enregistre()
Dim bexist As Boolean
Dim i As Long
Dim n As Long
Dim rstobj As Recordset
Dim cmb As ComboBox
Dim name As String
Dim bObligation As Boolean
Dim tblobjprop As Recordset

  Enregistre = False
  If Len(Trim(fldstrnomObjet)) < 1 Then
    Screen.MousePointer = vbNormal
    MsgBox "Saisie d'un nom d'objet obligatoire", vbExclamation, Me.Caption
    Call Set_Focus(fldstrnomObjet)
    Exit Function
  End If
  
  For i = 0 To Me.Count - 1
    name = Me(i).name
    If UCase(left(name, 3)) = "CMB" Then
      Set cmb = Me(i)
      bObligation = False
      If name = "CmbGroupe" Then
        bObligation = True
      End If
      If (Trim(cmb) <> "" Or bObligation = True) And cmb.Enabled = True Then
        bexist = False
        For n = 0 To cmb.ListCount - 1
          If UCase(cmb) = UCase(cmb.List(n)) Then
            bexist = True
            Exit For
          End If
        Next n
        If bexist = False Then
          Screen.MousePointer = vbNormal
          MsgBox "Saisie obligatoire d'un élément de la liste", vbExclamation, Me.Caption
          SSTab1.Tab = 0
          cmb.SetFocus
          Exit Function
        End If
      End If
    End If
  Next i
  
  If Trim(fldnCoutTotal) <> "" Then
    fldnCoutTotal = FrmtNum(fldnCoutTotal)
    If Not IsNumeric(fldnCoutTotal) Then
      Screen.MousePointer = vbNormal
      MsgBox "Saisie d'un cout total numérique", vbExclamation, Me.Caption
      SSTab1.Tab = 0
      Set_Focus fldnCoutTotal
      Exit Function
    End If
  End If
  
  If Trim(fldnPoidsBase) <> "" Then
    fldnPoidsBase = FrmtNum(fldnPoidsBase)
    If Not IsNumeric(fldnPoidsBase) Then
      Screen.MousePointer = vbNormal
      MsgBox "Saisie d'un poids numérique", vbExclamation, Me.Caption
      SSTab1.Tab = 0
      Set_Focus fldnPoidsBase
      Exit Function
    End If
  End If
  
  If Trim(fldnEchecSorts) <> "" Then
    fldnEchecSorts = FrmtNum(fldnEchecSorts)
    If Not IsNumeric(fldnEchecSorts) Then
      Screen.MousePointer = vbNormal
      MsgBox "Saisie d'un echec au sort numérique", vbExclamation, Me.Caption
      SSTab1.Tab = 0
      Set_Focus fldnEchecSorts
      Exit Function
    End If
  End If
  
  If Trim(fldnPenaliteArmure) <> "" Then
    fldnPenaliteArmure = FrmtNum(fldnPenaliteArmure)
    If Not IsNumeric(fldnPenaliteArmure) Then
      Screen.MousePointer = vbNormal
      MsgBox "Saisie d'une pénalité d'armure numérique", vbExclamation, Me.Caption
      SSTab1.Tab = 0
      Set_Focus fldnPenaliteArmure
      Exit Function
    End If
  End If

  If Trim(fldnBonusDexMax) <> "" Then
    fldnBonusDexMax = FrmtNum(fldnBonusDexMax)
    If Not IsNumeric(fldnBonusDexMax) Then
      Screen.MousePointer = vbNormal
      MsgBox "Saisie d'un bonus de dex max numérique", vbExclamation, Me.Caption
      SSTab1.Tab = 0
      Set_Focus fldnBonusDexMax
      Exit Function
    End If
  End If
  
  If Trim(fldnPortee) <> "" Then
    fldnPortee = FrmtNum(fldnPortee)
    If Not IsNumeric(fldnPortee) Then
      Screen.MousePointer = vbNormal
      MsgBox "Saisie d'une portée numérique", vbExclamation, Me.Caption
      SSTab1.Tab = 0
      Set_Focus fldnPortee
      Exit Function
    End If
  End If
  
  If Trim(fldnsolidite) <> "" Then
    fldnsolidite = FrmtNum(fldnsolidite)
    If Not IsNumeric(fldnsolidite) Then
      Screen.MousePointer = vbNormal
      MsgBox "Saisie d'une solidité numérique", vbExclamation, Me.Caption
      SSTab1.Tab = 0
      Set_Focus fldnsolidite
      Exit Function
    End If
  End If
  
  If Trim(fldnresistance) <> "" Then
    fldnresistance = FrmtNum(fldnresistance)
    If Not IsNumeric(fldnresistance) Then
      Screen.MousePointer = vbNormal
      MsgBox "Saisie d'une résistance numérique", vbExclamation, Me.Caption
      SSTab1.Tab = 0
      Set_Focus fldnresistance
      Exit Function
    End If
  End If
  
  If Trim(fldnCharge) <> "" Then
    fldnCharge = FrmtNum(fldnCharge)
    If Not IsNumeric(fldnCharge) Then
      Screen.MousePointer = vbNormal
      MsgBox "Saisie d'un nombre de charge numérique", vbExclamation, Me.Caption
      SSTab1.Tab = 0
      Set_Focus fldnCharge
      Exit Function
    End If
  End If
  
  If Trim(fldstrdescription) <> "" Then
    fldstrdescription = Trim(fldstrdescription)
  End If
  
  Screen.MousePointer = vbHourglass
  Set rstobj = g_bdPerso.OpenRecordset("objets", dbOpenTable)
  With rstobj
    .Index = "primarykey"
    .Seek "=", fldstrnomObjet
    If .NoMatch Then
      .AddNew
    Else
      .Edit
    End If
    !NomObjet = fldstrnomObjet
    !CoutTotal = IIf(IsNumeric(fldnCoutTotal), fldnCoutTotal, Null)
    !PoidsBase = IIf(IsNumeric(fldnPoidsBase), fldnPoidsBase, Null)
    !classeArmure = IIf(Trim(cmbClasseArmure) = "", Null, cmbClasseArmure)
    !typeca = IIf(Trim(CmbTypeCa) = "", Null, CmbTypeCa)
    !EchecSortProfane = IIf(IsNumeric(fldnEchecSorts), fldnEchecSorts, Null)
    !PenaliteArmure = IIf(IsNumeric(fldnPenaliteArmure), fldnPenaliteArmure, Null)
    !BonusDexMax = IIf(IsNumeric(fldnBonusDexMax), fldnBonusDexMax, Null)
    !Portee = IIf(IsNumeric(fldnPortee), fldnPortee, Null)
    !Solidite = IIf(IsNumeric(fldnsolidite), fldnsolidite, Null)
    !RESISTANCE = IIf(IsNumeric(fldnresistance), fldnresistance, Null)
    !GroupeObjet = CmbGroupe
    !taille = IIf(Trim(CmbTailles) = "", Null, CmbTailles)
    !Type = IIf(Trim(CmbType) = "", Null, CmbType)
    !DegatDes = IIf(IsNumeric(CmbDegatsDes), CmbDegatsDes, Null)
    !ZoneCritique = IIf(IsNumeric(CmbZoneCrit), CmbZoneCrit, Null)
    !MultCrit = IIf(IsNumeric(CmbMultipCrit), CmbMultipCrit, Null)
    !PoidsBase = IIf(IsNumeric(fldnPoidsBase), fldnPoidsBase, Null)
    !CHARGE = IIf(IsNumeric(fldnCharge), fldnCharge, Null)
    !Description = IIf(fldstrdescription = "", Null, fldstrdescription)
    .Update
    .Close
  End With
  g_bdPerso.Execute "delete from ObjetsPropriete where nomObjet='" & FormatStrSQL(fldstrnomObjet) & "'", dbFailOnError
  Set tblobjprop = g_bdPerso.OpenRecordset("objetspropriete", dbOpenTable)
  With tblobjprop
    For i = 0 To vsPropri.Rows - 1
      .AddNew
      !NomObjet = fldstrnomObjet
      !propriete_1 = IIf(vsPropri.Cell(flexcpText, i, 1) = "", Null, vsPropri.Cell(flexcpText, i, 1))
      !propriete_2 = IIf(vsPropri.Cell(flexcpText, i, 2) = "", Null, vsPropri.Cell(flexcpText, i, 2))
      !propriete_3 = IIf(vsPropri.Cell(flexcpText, i, 3) = "", Null, vsPropri.Cell(flexcpText, i, 3))
      !valeur = IIf(vsPropri.Cell(flexcpText, i, 4) = "", Null, vsPropri.Cell(flexcpText, i, 4))
      .Update
    Next i
    .Close
  End With
  
  vsObjets.Redraw = False
  RemplieLst
  RetrouvVSFLex vsObjets, fldstrnomObjet, COL_NOM_OBJET, 0
  vsObjets.Redraw = True
  m_bmodif = False
  m_bCharge = False
  Enregistre = True
End Function
Private Sub Supprimer()
Dim n As Integer
Dim rstTest As Recordset
Dim bexist As Boolean

 
  With vsObjets
    For n = 1 To .Rows - 1
      If .IsSelected(n) Then
        Screen.MousePointer = vbNormal
        If MsgBox("Confirmer la suppression de cet objet :" & .Cell(flexcpText, n, COL_NOM_OBJET), vbDefaultButton2 Or vbOKCancel Or vbQuestion, "Suppression") = vbOK Then
          Set rstTest = g_bdPerso.OpenRecordset("select * from Equipement where nomobjet='" & FormatStrSQL(.Cell(flexcpText, n, COL_NOM_OBJET)) & "'", dbOpenDynaset)
          bexist = Not rstTest.EOF
          rstTest.Close
          If bexist Then
            MsgBox "Certains personnages sont équipés de cette arme, elle ne peut être detruite", vbInformation, Me.Caption
            Exit Sub
          End If
          ''
          g_bdPerso.Execute "delete from objetsPropriete where [nomobjet]='" & FormatStrSQL(.Cell(flexcpText, n, COL_NOM_OBJET)) & "'", dbFailOnError
          g_bdPerso.Execute "delete from objets where [nomobjet]='" & FormatStrSQL(.Cell(flexcpText, n, COL_NOM_OBJET)) & "'", dbFailOnError
          Call RemplieLst
          If vsObjets.Rows > 1 Then
            vsObjets.IsSelected(1) = True
          Else
            Efface Me
            vsObjets_SelChange
          End If
        End If
        Exit For
      End If
    Next n
  End With
  

End Sub
Private Sub Fermer()


  If BtnFermer.Visible = True Then
    Unload Me
  Else
    Call vsObjets_SelChange
  End If


End Sub
Private Sub CmbType_Change()

  Modification

End Sub

Private Sub CmbType_Click()
 
  Modification
End Sub

Private Sub CmbType_GotFocus()
  Selec
End Sub

Private Sub CmbTypeCa_Change()
  Modification
End Sub

Private Sub CmbTypeCa_Click()
  Modification
End Sub

Private Sub CmbTypeCa_GotFocus()
  Selec
End Sub

Private Sub CmbZoneCrit_Change()

  Modification

End Sub

Private Sub CmbZoneCrit_Click()

Modification

End Sub



Private Sub CmbZoneCrit_GotFocus()
  Selec
End Sub

Private Sub fldnBonusDexMax_Change()
  Modification
End Sub

Private Sub fldnBonusDexMax_GotFocus()
  Selec
End Sub



Private Sub fldnCharge_Change()
  Modification
End Sub

Private Sub fldnCharge_GotFocus()
  Selec
  StatusBar1.Panels(1).Text = "Saisissez le nombre total de charge"
End Sub

Private Sub fldnCharge_LostFocus()
  StatusBar1.Panels(1).Text = ""
End Sub

Private Sub fldnCoutTotal_Change()
  Modification
End Sub

Private Sub fldnCoutTotal_GotFocus()
  Selec
End Sub

Private Sub fldnEchecSorts_Change()
  Modification
End Sub

Private Sub fldnEchecSorts_GotFocus()
  Selec
End Sub


Private Sub fldnPenaliteArmure_Change()
  Modification
End Sub

Private Sub fldnPenaliteArmure_GotFocus()
  Selec
End Sub

Private Sub fldnPoidsbase_Change()
  Modification
End Sub

Private Sub fldnPoidsbase_GotFocus()
  Selec
End Sub


Private Sub fldnPortee_Change()
  Modification
End Sub

Private Sub fldnPortee_GotFocus()
  Selec
  StatusBar1.Panels(1).Text = "Saisissez la portée"
End Sub

Private Sub fldnPortee_LostFocus()
  StatusBar1.Panels(1).Text = ""
End Sub
Private Sub fldnsolidite_Change()
  Modification
End Sub

Private Sub fldnsolidite_GotFocus()
  Selec
  StatusBar1.Panels(1).Text = "Saisissez la solidité"
End Sub

Private Sub fldnsolidite_LostFocus()
  StatusBar1.Panels(1).Text = ""
End Sub
Private Sub fldnresistance_Change()
  Modification
End Sub

Private Sub fldnresistance_GotFocus()
  Selec
  StatusBar1.Panels(1).Text = "Saisissez la résistance"
End Sub

Private Sub fldnresistance_LostFocus()
  StatusBar1.Panels(1).Text = ""
End Sub
Private Sub fldstrdescription_Change()
  Modification
End Sub
Private Sub fldstrdescription_gotfocus()
  Selec
End Sub
Private Sub fldstrnomobjet_Change()

  Modification
End Sub
Private Sub fldstrnomobjet_GotFocus()

  Selec
  StatusBar1.Panels(1).Text = "Saisissez l'identifiant de l'objet ex (Epée longue de ....)"
End Sub
Private Sub fldstrnomobjet_LostFocus()


  StatusBar1.Panels(1).Text = ""
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)


  Select Case KeyCode
    Case vbKeyF2
      If btnEnregistrer.Visible = True Then btnEnregistrer_Click
    Case vbKeyReturn
      SendKeys "{TAB}"
    Case vbKeyEscape
      Fermer
    Case vbKeyF9
      If btnAjouter.Visible = True Then btnAjouter_Click
    Case vbKeyF10
      If btnSupprimer.Visible = True Then Supprimer
  End Select
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)

  
  Select Case KeyAscii
    Case vbKeyEscape
      KeyAscii = 0
    Case vbKeyReturn
      KeyAscii = 0
  End Select

End Sub

Private Sub Form_Load()
Dim rst As Recordset
Dim i As Integer

  ReDim m_tabButton(6)
  m_tabButton(0).Nom = "BtnFermer"
  m_tabButton(1).Nom = "BtnAjouter"
  m_tabButton(2).Nom = "BtnEnregistrer"
  m_tabButton(3).Nom = "BtnSupprimer"
  m_tabButton(4).Nom = "BtnAnnuler"
  m_tabButton(5).Nom = "btnRenommer"
  
  
  vsPropri.ColHidden(1) = True
  vsPropri.ColHidden(2) = True
  vsPropri.ColHidden(3) = True
  vsPropri.ColHidden(4) = True
  vsPropri.ColWidth(0) = 2500
  
  SSTab1.Tab = 0
  
  rempTree
  
  
  CmbGroupe.Clear
  CmbGroupe.AddItem obj_HOOPAK
  CmbGroupe.AddItem obj_HOOPAK_DE_JET
  CmbGroupe.AddItem obj_EPIEU_DE_JET
  CmbGroupe.AddItem obj_GOURDIN_DE_JET
  CmbGroupe.AddItem obj_LANCE_DE_JET
  CmbGroupe.AddItem obj_JAVELINE_DE_JET
  CmbGroupe.AddItem obj_MARTEAU_LEGER_DE_JET
  CmbGroupe.AddItem obj_TRIDENT_DE_JET
  CmbGroupe.AddItem obj_GANTELET
  CmbGroupe.AddItem obj_DAGUES_DE_JET
  CmbGroupe.AddItem obj_FRELELAME_ELFE
  CmbGroupe.AddItem obj_FINELAME_ELFE
  CmbGroupe.AddItem obj_GRAND_PIC_DE_GUERRE
  CmbGroupe.AddItem obj_MAILLET
  CmbGroupe.AddItem obj_MASSE_DE_BRECHE
  CmbGroupe.AddItem obj_AGRIPPEUR
  CmbGroupe.AddItem obj_BARDICHE
  CmbGroupe.AddItem obj_GRAND_EPIEU
  CmbGroupe.AddItem obj_LAJATANG
  CmbGroupe.AddItem obj_MARTEAU_DOUBLE
  CmbGroupe.AddItem obj_BOOMERANG
  CmbGroupe.AddItem obj_GRAND_ARC
  CmbGroupe.AddItem obj_GRANDE_SARBACANE
  CmbGroupe.AddItem obj_MASSE_D_ARMES_LEGERE
  CmbGroupe.AddItem obj_MASSE_D_ARMES_LOURDE
  CmbGroupe.AddItem obj_SERPE
  CmbGroupe.AddItem obj_EPIEU
  CmbGroupe.AddItem obj_MORGENSTERN
  CmbGroupe.AddItem obj_BATON
  CmbGroupe.AddItem obj_LANCE
  CmbGroupe.AddItem obj_PIQUE
  CmbGroupe.AddItem obj_ARBALETE_LEGERE
  CmbGroupe.AddItem obj_ARBALETE_LOURDE
  CmbGroupe.AddItem obj_DARDS
  CmbGroupe.AddItem obj_JAVELINE
  CmbGroupe.AddItem obj_HACHE_DE_LANCER
  CmbGroupe.AddItem obj_ROCHER
  CmbGroupe.AddItem obj_HACHETTE
  CmbGroupe.AddItem obj_KUKRI
  CmbGroupe.AddItem obj_MARTEAU_LEGER
  CmbGroupe.AddItem obj_MATRAQUE
  CmbGroupe.AddItem obj_PIC_DE_GUERRE_LEGER
  CmbGroupe.AddItem obj_FLEAU_D_ARMES_LEGER
  CmbGroupe.AddItem obj_MARTEAU_DE_GUERRE
  CmbGroupe.AddItem obj_PIC_DE_GUERRE_LOURD
  CmbGroupe.AddItem obj_RAPIERE
  CmbGroupe.AddItem obj_TRIDENT
  CmbGroupe.AddItem obj_CIMETERRE_A_2_MAINS
  CmbGroupe.AddItem obj_CORSEQUE
  CmbGroupe.AddItem obj_COUTILLE
  CmbGroupe.AddItem obj_FAUX
  CmbGroupe.AddItem obj_FLEAU_D_ARMES_LOURD
  CmbGroupe.AddItem obj_GRANDE_HACHE
  CmbGroupe.AddItem obj_GUISARME
  CmbGroupe.AddItem obj_HALLEBARDE
  CmbGroupe.AddItem obj_LANCE_D_ARCON
  CmbGroupe.AddItem obj_MASSUE
  CmbGroupe.AddItem obj_KAMA
  CmbGroupe.AddItem obj_NUNCHAKU
  CmbGroupe.AddItem obj_SAI
  CmbGroupe.AddItem obj_SIANGHAM
  CmbGroupe.AddItem obj_EPEE_BATARDE
  CmbGroupe.AddItem obj_FOUET
  CmbGroupe.AddItem obj_CHAINE_CLOUTEE
  CmbGroupe.AddItem obj_DOUBLE_LAME
  CmbGroupe.AddItem obj_FLEAU_DOUBLE
  CmbGroupe.AddItem obj_HACHE_DOUBLE_ORQUE
  CmbGroupe.AddItem obj_MARTEAU_PIOLET_GNOME
  CmbGroupe.AddItem obj_URGROSH_NAIN
  CmbGroupe.AddItem obj_ARBALETE_DE_POING
  CmbGroupe.AddItem obj_BOLAS
  CmbGroupe.AddItem obj_FILET
  CmbGroupe.AddItem obj_SHURIKEN
  CmbGroupe.AddItem obj_VETEMENT
  CmbGroupe.AddItem obj_CIMETERE
  CmbGroupe.AddItem obj_DAGUE
  CmbGroupe.AddItem obj_EPEE_A_2_MAINS
  CmbGroupe.AddItem obj_EPEE_COURTE
  CmbGroupe.AddItem obj_EPEE_LONGUE
  CmbGroupe.AddItem obj_GOURDIN
  CmbGroupe.AddItem obj_HACHE_D_ARMES
  CmbGroupe.AddItem obj_HACHE_DOUBLE
  CmbGroupe.AddItem obj_HACHE_DE_NAIN
  CmbGroupe.AddItem obj_TARGE
  CmbGroupe.AddItem obj_RONDACHE
  CmbGroupe.AddItem obj_ECU
  CmbGroupe.AddItem obj_PAVOIS
  CmbGroupe.AddItem obj_ANNEAU
  CmbGroupe.AddItem obj_CEINTURE
  CmbGroupe.AddItem obj_AMULETTE
  CmbGroupe.AddItem obj_POTION
  CmbGroupe.AddItem obj_SCEPTRE
  CmbGroupe.AddItem obj_BAGUETTE
  CmbGroupe.AddItem obj_ARMURE
  CmbGroupe.AddItem obj_ARC_LONG
  CmbGroupe.AddItem obj_ARC_COURT
  CmbGroupe.AddItem obj_FRONDE
  CmbGroupe.AddItem obj_CASQUE
  CmbGroupe.AddItem obj_HEAUME
  CmbGroupe.AddItem obj_CAPE
  CmbGroupe.AddItem obj_GANTS
  CmbGroupe.AddItem obj_CARREAUX
  CmbGroupe.AddItem obj_FLECHES
  CmbGroupe.AddItem obj_BILLES
  CmbGroupe.AddItem obj_BRACELET
  CmbGroupe.AddItem obj_DIVERS
  CmbGroupe.AddItem obj_PERLES
  CmbGroupe.AddItem obj_MEDAILLON
  CmbGroupe.AddItem obj_COURONNE
  CmbGroupe.AddItem obj_BANDEAU
  CmbGroupe.AddItem obj_PHYLACTERE
  CmbGroupe.AddItem obj_LUNETTES
  CmbGroupe.AddItem obj_BROCHE
  CmbGroupe.AddItem obj_SCARABEE
  CmbGroupe.AddItem obj_COLLIER
  CmbGroupe.AddItem obj_CHARME
  CmbGroupe.AddItem obj_GILET
  CmbGroupe.AddItem obj_VESTE
  CmbGroupe.AddItem obj_CHEMISE
  CmbGroupe.AddItem obj_CHASUBLE
  CmbGroupe.AddItem obj_ROBE
  CmbGroupe.AddItem obj_MANTEAU
  CmbGroupe.AddItem obj_ECHARPE
  CmbGroupe.AddItem obj_BOTTES
  CmbGroupe.AddItem obj_CHAUSSONS
  CmbGroupe.AddItem obj_CEINTURON
  CmbGroupe.AddItem obj_MATERIEL_D_AVENTURIER
  CmbGroupe.AddItem obj_MATERIEL_DE_CLASSE
  CmbGroupe.AddItem obj_MONTURE_ET_HARNACHEMENT
  CmbGroupe.AddItem obj_MOYENS_DE_TRANSPORT
  CmbGroupe.AddItem obj_OBJETS_SPECIAUX
  CmbGroupe.AddItem obj_VETEMENTS
  CmbGroupe.AddItem obj_COUPE_VIF_GNOME
  CmbGroupe.AddItem obj_MARTEAU_DE_LANCER
  CmbGroupe.AddItem obj_CAPTE_EPEE_GNOME
  CmbGroupe.AddItem obj_GRAND_MARTEAU_GOLIATH
  CmbGroupe.AddItem obj_LANCE_DOUBLE_NAINE
  CmbGroupe.AddItem obj_PERTUISANE_NAINE
  CmbGroupe.AddItem obj_SORT_A_TOUCHER_DISTANT
  CmbGroupe.AddItem obj_SORT_AU_CONTACT
  CmbGroupe.AddItem obj_DECHARGE_FANTASTIQUE
  CmbGroupe.AddItem obj_RAYON
  CmbGroupe.AddItem obj_MAINS_NUES
  CmbGroupe.AddItem obj_SERRE
  CmbGroupe.AddItem obj_CORNE
  CmbGroupe.AddItem obj_MORSURE
  CmbGroupe.AddItem obj_TENTACULE
  CmbGroupe.AddItem obj_SABOT
  CmbGroupe.AddItem obj_COUP
  CmbGroupe.AddItem obj_DEFENCE
  CmbGroupe.AddItem obj_DARD
  CmbGroupe.AddItem obj_MANDIBULE
  CmbGroupe.AddItem obj_PINCE
  CmbGroupe.AddItem obj_BRANCHE
  CmbGroupe.AddItem obj_ANTENNE
  CmbGroupe.AddItem obj_PATTE
  CmbGroupe.AddItem obj_BEC
  CmbGroupe.AddItem obj_GRIFFE
  CmbGroupe.AddItem obj_ERGOT
  CmbGroupe.AddItem obj_PATTES_ARRIERE
  CmbGroupe.AddItem obj_AILLE
  CmbGroupe.AddItem obj_QUEUE
  CmbGroupe.AddItem obj_PIQUANT
  
  CmbTailles.Clear
  CmbTailles.AddItem obj_PETITE
  CmbTailles.AddItem obj_MOYENNE
  CmbTailles.AddItem obj_GRANDE
  
  CmbType.Clear
  CmbType.AddItem obj_PIQUANT
  CmbType.AddItem obj_TRANCHANT
  CmbType.AddItem obj_CONTONDANT
  
  CmbZoneCrit.Clear
  CmbZoneCrit.AddItem "20"
  CmbZoneCrit.AddItem "19"
  CmbZoneCrit.AddItem "18"
  
  CmbMultipCrit.Clear
  CmbMultipCrit.AddItem "2"
  CmbMultipCrit.AddItem "3"
  CmbMultipCrit.AddItem "4"
  
  CmbTypeCa.Clear
  CmbTypeCa.AddItem obj_CA_ARMURE
  CmbTypeCa.AddItem obj_CA_BOUCLIER
  
  CmbDegatsDes.Clear
  CmbDegatsDes.AddItem "1d2"
  CmbDegatsDes.AddItem "1d3"
  CmbDegatsDes.AddItem "1d4"
  CmbDegatsDes.AddItem "1d5"
  CmbDegatsDes.AddItem "1d6"
  CmbDegatsDes.AddItem "1d7"
  CmbDegatsDes.AddItem "1d8"
  CmbDegatsDes.AddItem "1d10"
  CmbDegatsDes.AddItem "1d12"
  CmbDegatsDes.AddItem "2d4"
  CmbDegatsDes.AddItem "2d6"
  CmbDegatsDes.AddItem "2d8"
  CmbDegatsDes.AddItem "2d10"
  CmbDegatsDes.AddItem "2d12"
  CmbDegatsDes.AddItem "3d6"
  
  cmbClasseArmure.Clear
  For i = 1 To 10: cmbClasseArmure.AddItem i: Next i
  
   ' Ouverture base de donnée
  Data1.DatabaseName = g_bdPerso.name
  RemplieLst
  If vsObjets.Rows > 1 Then
    vsObjets.IsSelected(1) = True
  Else
    vsObjets_SelChange
  End If
  
End Sub
Sub RemplieLst()
Dim i As Integer
Dim stSql As String

 
  Screen.MousePointer = vbHourglass
  stSql = "SELECT nomobjet,couttotal,poidsBase,ClasseArmure,TypeCa,EchecSortProfane,PenaliteArmure,BonusDexMax,portee,GroupeObjet,Taille,Type,DegatDes,ZoneCritique,MultCrit,description, solidite,resistance,charge from objets "
  Data1.RecordSource = stSql
  Data1.Refresh
  With vsObjets
    .Redraw = False
    .DataRefresh
    .FormatString = "Nom  "
    .ColAlignment(COL_NOM_OBJET) = flexAlignLeftCenter
    .Cols = 19
    For i = COL_COUT_TOTAL To COL_CHARGE
      .ColHidden(i) = True
    Next i
    .Redraw = True
  End With
  Screen.MousePointer = vbNormal

End Sub
Private Sub Form_Unload(Cancel As Integer)


  If NonModifier() = False Then Cancel = True: Exit Sub


End Sub

Private Sub Timer1_Timer()

  Timer1.Interval = 0
  If g_strNomEquipModif <> "" Then
    RetrouvVSFLex vsObjets, g_strNomEquipModif, COL_NOM_OBJET, 0
    fldstrnomObjet.Enabled = False
    vsObjets.Enabled = False
  Else
    fldstrnomObjet.Enabled = True
    vsObjets.Enabled = True
  End If
  AjusteBouton Me, m_tabButton
End Sub

Private Sub Tree1_NodeClick(ByVal Node As MSComctlLib.Node)

Dim i As Integer, j As Integer, k As Integer
Dim key1 As String, key2 As String, key3 As String, key4 As String, nSort As Integer
Dim tTab() As S_TABVAL
Dim sTab() As S_TABVAL
Dim rst As Recordset

With Tree1
  If .SelectedItem.Children = 0 Then
    Select Case .SelectedItem.FullPath
      Case obj_BONUS_CARACTERISTIQUE
        key1 = "K1_" & obj_BONUS_CARACTERISTIQUE
        ReDim sTab(6)
          sTab(0).Nom = obj_FORCE
          sTab(1).Nom = obj_DEXTERITE
          sTab(2).Nom = obj_CONSTITUTION
          sTab(3).Nom = obj_INTELLIGENCE
          sTab(4).Nom = obj_SAGESSE
          sTab(5).Nom = obj_CHARISME
        ReDim tTab(2)
          tTab(0).Nom = ALTERATION
          tTab(1).Nom = "Divin"
        For j = 0 To UBound(sTab) - 1
          key2 = key1 & "_" & "K2_" & sTab(j).Nom
          .Nodes.Add key1, tvwChild, key2, sTab(j).Nom
          For k = 0 To UBound(tTab) - 1
            key3 = key2 & "_" & "K3_" & tTab(k).Nom
            .Nodes.Add key2, tvwChild, key3, "+" & tTab(k).Nom
            For i = 1 To 10
              key4 = key3 & "_" & "K4_" & CStr(i)
              .Nodes.Add key3, tvwChild, key4, "+" & CStr(i)
            Next i
          Next k
        Next j
      Case obj_BONUS_CLASSE_ARMURE
        ReDim sTab(11)
        sTab(0).Nom = obj_CA_ALTERATION_ARMURE
        sTab(1).Nom = obj_CA_ALTERATION_BOUCLIER
        sTab(2).Nom = obj_CA_ARMURE_NATURELLE
        sTab(3).Nom = obj_CA_PARADE
        sTab(4).Nom = obj_CA_ESQUIVE
        sTab(5).Nom = obj_CA_FORCE_ARMURE
        sTab(6).Nom = obj_CA_FORCE_BOUCLIER
        sTab(7).Nom = obj_CA_INTUITION
        sTab(8).Nom = obj_CA_CHANCE
        sTab(9).Nom = obj_CA_ALTERATION_ARMURE_NATURELLE
        sTab(10).Nom = obj_CA_DIVERS
        key1 = "K1_" & obj_BONUS_CLASSE_ARMURE
        For j = 0 To UBound(sTab) - 1
          key2 = key1 & "_" & "K2_" & sTab(j).Nom
          .Nodes.Add key1, tvwChild, key2, sTab(j).Nom
          For i = 1 To 20
            key4 = key2 & "_" & "K4_" & CStr(i)
            .Nodes.Add key2, tvwChild, key4, "+" & CStr(i)
          Next i
        Next j
      Case obj_BONUS_ATTAQUE
        key1 = "K1_" & obj_BONUS_ATTAQUE
        For i = 1 To 30
          key4 = key1 & "_" & "K4_" & CStr(i)
          .Nodes.Add key1, tvwChild, key4, "+" & CStr(i)
        Next i
      Case obj_BONUS_DON
        ReDim sTab(8)
          sTab(0).Nom = obj_INVOCATION
          sTab(1).Nom = obj_NECROMANCIE
          sTab(2).Nom = obj_ILLUSION
          sTab(3).Nom = obj_DIVINATION
          sTab(4).Nom = obj_TRANSMUTATION
          sTab(5).Nom = obj_ABJURATION
          sTab(6).Nom = obj_ENCHANTEMENT
          sTab(7).Nom = obj_EVOCATION
        key1 = "K1_" & obj_BONUS_DON
        Set rst = g_bdRoles.OpenRecordset("select Nom from dons order by nom", dbOpenDynaset)
        Do While Not rst.EOF
          key2 = key1 & "_" & "K2_" & rst!Nom
          .Nodes.Add key1, tvwChild, key2, rst!Nom
          If rst!Nom = ECOLE_RENFORCEE Or rst!Nom = ECOLE_SUPERIEURE Or rst!Nom = TATOUAGE_FOCALISATEUR Or rst!Nom = ECOLE_INTERDITE_SUPPLEMENTAIRE Then
            For k = 0 To UBound(sTab) - 1
              key3 = key2 & "_" & "K3_" & sTab(k).Nom
              .Nodes.Add key2, tvwChild, key3, sTab(k).Nom
            Next k
          End If
          If UCase(rst!Nom) = UCase(EMPLACEMENT_DE_SORT_SUPPLEMENTAIRE) Then
            For k = 0 To 9
              key3 = key2 & "_" & "K3_" & k
              .Nodes.Add key2, tvwChild, key3, k
            Next k
          End If
          rst.MoveNext
        Loop
        rst.Close
      Case obj_BONUS_DOMMAGE
        key1 = "K1_" & obj_BONUS_DOMMAGE
        ReDim sTab(17)
          sTab(0).Nom = obj_PIQUANT
          sTab(1).Nom = obj_TRANCHANT
          sTab(2).Nom = obj_CONTONDANT
          sTab(3).Nom = obj_SONIQUE
          sTab(4).Nom = obj_FROID
          sTab(5).Nom = obj_FEU
          sTab(6).Nom = obj_ELECTRICITE
          sTab(7).Nom = obj_ACIDE
          sTab(8).Nom = obj_DIVIN
          sTab(9).Nom = obj_ANARCHIQUE
          sTab(10).Nom = obj_AXIOMATIQUE
          sTab(11).Nom = obj_IMPIE
          sTab(12).Nom = obj_SAINTE
          sTab(13).Nom = obj_TUEUSE
          sTab(14).Nom = obj_PRECISION_MORTELLE
          sTab(15).Nom = obj_SANGLANTE
          sTab(16).Nom = obj_POISON
        ReDim tTab(6)
          tTab(0).Nom = obj_FORCE
          tTab(1).Nom = obj_DEXTERITE
          tTab(2).Nom = obj_CONSTITUTION
          tTab(3).Nom = obj_INTELLIGENCE
          tTab(4).Nom = obj_SAGESSE
          tTab(5).Nom = obj_CHARISME
        For j = 0 To UBound(sTab) - 1
          key2 = key1 & "_" & "K2_" & sTab(j).Nom
          .Nodes.Add key1, tvwChild, key2, sTab(j).Nom
          Select Case sTab(j).Nom
            Case obj_PIQUANT, obj_TRANCHANT, obj_CONTONDANT
              For i = 1 To 30
                key4 = key2 & "_" & "K4_" & CStr(i)
                .Nodes.Add key2, tvwChild, key4, "+" & CStr(i)
              Next i
            Case obj_SANGLANTE
              For k = 0 To UBound(tTab) - 1
                key3 = key2 & "_" & "K3_" & tTab(k).Nom
                .Nodes.Add key2, tvwChild, key3, tTab(k).Nom
                For i = -1 To -5 Step -1
                  key4 = key3 & "_" & "K4_" & CStr(i)
                  .Nodes.Add key3, tvwChild, key4, CStr(i)
                Next i
                key4 = key3 & "_" & "K4_" & "-1d6"
                .Nodes.Add key3, tvwChild, key4, "-1d6"
                key4 = key3 & "_" & "K4_" & "-2d6"
                .Nodes.Add key3, tvwChild, key4, "-2d6"
              Next k
            Case Else
              ''Xd6
              For i = 1 To 4
                key4 = key2 & "_" & "K4_" & i & "d6"
                .Nodes.Add key2, tvwChild, key4, i & "d6"
              Next i
              ''2d10
              key4 = key2 & "_" & "K4_" & "1d10"
              .Nodes.Add key2, tvwChild, key4, "1d10"
          End Select
        Next j
      Case obj_BONUS_JET_SAUVEGARDE
        ReDim sTab(3)
          sTab(0).Nom = obj_VIGUEUR
          sTab(1).Nom = obj_REFLEXE
          sTab(2).Nom = obj_VOLONTE
        ReDim tTab(6)
          tTab(0).Nom = RESISTANCE
          tTab(1).Nom = FAMILlIER
          tTab(2).Nom = DIVIN
          tTab(3).Nom = APTITUDE
          tTab(4).Nom = CHANCE
          tTab(5).Nom = INTUITION
        key1 = "K1_" & obj_BONUS_JET_SAUVEGARDE
        For j = 0 To UBound(sTab) - 1
          key2 = key1 & "_" & "K2_" & sTab(j).Nom
          .Nodes.Add key1, tvwChild, key2, sTab(j).Nom
          For k = 0 To UBound(tTab) - 1
            key3 = key2 & "_" & "K3_" & tTab(k).Nom
            .Nodes.Add key2, tvwChild, key3, "+" & tTab(k).Nom
            For i = 1 To 10
              key4 = key3 & "_" & "K4_" & CStr(i)
              .Nodes.Add key3, tvwChild, key4, "+" & CStr(i)
            Next i
          Next k
        Next j
      Case obj_BONUS_COMPETENCE
        key1 = "K1_" & obj_BONUS_COMPETENCE
        Set rst = g_bdRoles.OpenRecordset("select nom from competence order by nom", dbOpenDynaset)
        Do While Not rst.EOF
          key2 = key1 & "_" & "K2_" & rst!Nom
          .Nodes.Add key1, tvwChild, key2, rst!Nom
          rst.MoveNext
        Loop
        rst.Close
      Case obj_BONUS_DEPLACEMENT
        ReDim sTab(3)
          sTab(0).Nom = obj_TERRE
          sTab(1).Nom = obj_AIR
          sTab(2).Nom = obj_EAU
        ReDim tTab(2)
          tTab(0).Nom = ALTERATION
          tTab(1).Nom = APTITUDE
        key1 = "K1_" & obj_BONUS_DEPLACEMENT
        For j = 0 To UBound(sTab) - 1
          key2 = key1 & "_" & "K2_" & sTab(j).Nom
          .Nodes.Add key1, tvwChild, key2, sTab(j).Nom
          For k = 0 To UBound(tTab) - 1
            key3 = key2 & "_" & "K3_" & tTab(k).Nom
            .Nodes.Add key2, tvwChild, key3, "+" & tTab(k).Nom
            For i = 1 To 20
              key4 = key3 & "_" & "K4_" & CStr(i)
              .Nodes.Add key3, tvwChild, key4, "+" & CStr(i)
            Next i
          Next k
        Next j
      Case obj_RESISTANCE_A_LA_MAGIE, obj_RESISTANCE_AU_FEU, obj_RESISTANCE_AU_FROID, obj_RESISTANCE_A_L_ELECTRICITE, obj_RESISTANCE_A_L_ACIDE, obj_RESISTANCE_AU_SON
          key1 = "K1_" & .SelectedItem.FullPath
        For i = 1 To 50
          key4 = key1 & "_" & "K4_" & CStr(i)
          .Nodes.Add key1, tvwChild, key4, "+" & CStr(i)
        Next i
      Case obj_RESISTANCE_AUX_DEGATS
        key1 = "K1_" & obj_RESISTANCE_AUX_DEGATS
        Set rst = g_bdRoles.OpenRecordset("select nom from Resistance order by nom", dbOpenDynaset)
        rst.MoveLast
        rst.MoveFirst
        ReDim tTab(rst.RecordCount)
        For i = 0 To rst.RecordCount - 1
          tTab(i).Nom = rst!Nom
          rst.MoveNext
        Next i
        rst.Close
        For k = 0 To UBound(tTab) - 1
          key2 = key1 & "_" & "K2_" & tTab(k).Nom
          .Nodes.Add key1, tvwChild, key2, "+" & tTab(k).Nom
          For i = 1 To 30
            key4 = key2 & "_" & "K4_" & CStr(i)
            .Nodes.Add key2, tvwChild, key4, "+" & CStr(i)
          Next i
        Next k
      Case obj_ARC_DE_PUISSANCE
        key1 = "K1_" & obj_ARC_DE_PUISSANCE
        For i = 1 To 20
          key4 = key1 & "_" & "K4_" & CStr(i)
          .Nodes.Add key1, tvwChild, key4, "+" & CStr(i)
        Next i
      Case obj_BONUS_ARC, obj_BONUS_JET_ATTAQUE
        ReDim sTab(2)
          sTab(0).Nom = obj_BONUS_ATTAQUE
          sTab(1).Nom = obj_BONUS_DOMMAGE
        ReDim tTab(1)
          tTab(0).Nom = APTITUDE
        key1 = "K1_" & .SelectedItem.FullPath
        For j = 0 To UBound(sTab) - 1
          key2 = key1 & "_" & "K2_" & sTab(j).Nom
          .Nodes.Add key1, tvwChild, key2, sTab(j).Nom
          For k = 0 To UBound(tTab) - 1
            key3 = key2 & "_" & "K3_" & tTab(k).Nom
            .Nodes.Add key2, tvwChild, key3, "+" & tTab(k).Nom
            For i = 1 To 3
              key4 = key3 & "_" & "K4_" & CStr(i)
              .Nodes.Add key3, tvwChild, key4, "+" & CStr(i)
            Next i
          Next k
        Next j
      Case obj_BONUS_ARME_NATURELLE
        ReDim sTab(2)
         sTab(0).Nom = obj_BONUS_ATTAQUE
          sTab(1).Nom = obj_BONUS_DOMMAGE
        ReDim tTab(1)
          tTab(0).Nom = ALTERATION
        key1 = "K1_" & obj_BONUS_ARME_NATURELLE
        For j = 0 To UBound(sTab) - 1
          key2 = key1 & "_" & "K2_" & sTab(j).Nom
          .Nodes.Add key1, tvwChild, key2, sTab(j).Nom
          For k = 0 To UBound(tTab) - 1
            key3 = key2 & "_" & "K3_" & tTab(k).Nom
            .Nodes.Add key2, tvwChild, key3, "+" & tTab(k).Nom
            For i = 1 To 7
              key4 = key3 & "_" & "K4_" & CStr(i)
              .Nodes.Add key3, tvwChild, key4, "+" & CStr(i)
            Next i
          Next k
        Next j
      Case obj_REJET_MORT_VIVANT_ACCRU
        key1 = "K1_" & obj_REJET_MORT_VIVANT_ACCRU
        For i = 1 To 10
          key4 = key1 & "_" & "K4_" & CStr(i)
          .Nodes.Add key1, tvwChild, key4, "+" & CStr(i)
        Next i
      Case obj_ARCANES
        ReDim sTab(9)
          sTab(0).Nom = obj_ARCANES_1
          sTab(1).Nom = obj_ARCANES_2
          sTab(2).Nom = obj_ARCANES_3
          sTab(3).Nom = obj_ARCANES_4
          sTab(4).Nom = obj_ARCANES_5
          sTab(5).Nom = obj_ARCANES_6
          sTab(6).Nom = obj_ARCANES_7
          sTab(7).Nom = obj_ARCANES_8
          sTab(8).Nom = obj_ARCANES_9
        key1 = "K1_" & obj_ARCANES
        For j = 0 To UBound(sTab) - 1
          key2 = key1 & "_" & "K2_" & sTab(j).Nom
          .Nodes.Add key1, tvwChild, key2, sTab(j).Nom
          For i = 2 To 3
            key4 = key2 & "_" & "K4_" & CStr(i)
            .Nodes.Add key2, tvwChild, key4, "x " & CStr(i)
          Next i
        Next j
      Case obj_BONUS_INITIATIVE
        key1 = "K1_" & obj_BONUS_INITIATIVE
        For i = 1 To 10
          key4 = key1 & "_" & "K4_" & CStr(i)
          .Nodes.Add key1, tvwChild, key4, "+" & CStr(i)
        Next i
      Case obj_LANCER_SORT
        If .SelectedItem.Children = 0 Then
          Set rst = g_bdRoles.OpenRecordset("select nom from sort order by nom", dbOpenDynaset)
          rst.MoveLast
          rst.MoveFirst
          ReDim sTab(rst.RecordCount)
          For i = 0 To rst.RecordCount - 1
            sTab(i).Nom = rst!Nom
            rst.MoveNext
          Next i
          rst.Close
          If UBound(sTab) > 0 Then
            key1 = "K1_" & obj_LANCER_SORT
            nSort = UBound(sTab) - 1
            For j = 0 To nSort
              key2 = key1 & "_" & "K2_" & sTab(j).Nom
              .Nodes.Add key1, tvwChild, key2, sTab(j).Nom
            Next j
          End If
        End If
      Case Else
        If InStr(1, .SelectedItem.FullPath, "\", vbTextCompare) <> 0 Then
          If .SelectedItem.Parent = obj_LANCER_SORT Then
          ReDim tTab(4)
            tTab(0).Nom = obj_LANCER_SORT_A_CHARGES
            tTab(1).Nom = obj_LANCER_SORT_PAR_JOUR
            tTab(2).Nom = obj_LANCER_SORT_PAR_SEMAINE
            tTab(3).Nom = obj_LANCER_SORT_PAR_MOIS
          key1 = "K1_" & obj_LANCER_SORT
          key2 = key1 & "_" & "K2_" & Node
            For k = 0 To UBound(tTab) - 1
              key3 = key2 & "_" & "K3_" & tTab(k).Nom
              .Nodes.Add key2, tvwChild, key3, tTab(k).Nom
              For i = 1 To 5
                key4 = key3 & "_" & "K4_" & CStr(i)
                .Nodes.Add key3, tvwChild, key4, CStr(i)
              Next i
            Next k
            key3 = key2 & "_" & "K3_" & "Illimité"
            .Nodes.Add key2, tvwChild, key3, "Illimité"
            key3 = key2 & "_" & "K3_" & "Permanent"
            .Nodes.Add key2, tvwChild, key3, "Permanent"
          End If
          If .SelectedItem.Parent = obj_BONUS_COMPETENCE Then
            ReDim tTab(4)
              tTab(0).Nom = APTITUDE
              tTab(1).Nom = FAMILlIER
              tTab(2).Nom = CHANCE
              tTab(3).Nom = CIRCONSTANCES
            key1 = "K1_" & obj_BONUS_COMPETENCE
            key2 = key1 & "_" & "K2_" & Node
            For k = 0 To UBound(tTab) - 1
              key3 = key2 & "_" & "K3_" & tTab(k).Nom
              .Nodes.Add key2, tvwChild, key3, "+" & tTab(k).Nom
              For i = 1 To 30
                key4 = key3 & "_" & "K4_" & CStr(i)
                .Nodes.Add key3, tvwChild, key4, "+" & CStr(i)
              Next i
            Next k
          End If
        End If
    End Select
  End If
End With
End Sub

Private Sub vsObjets_AfterEdit(ByVal Row As Long, ByVal Col As Long)

  Modification

End Sub

Private Sub vsObjets_Click()


  vsObjets_SelChange


End Sub
Private Sub vsObjets_SelChange()
Dim i As Integer, n As Integer
Dim bVisible As Boolean
Dim rst As Recordset
  

  bVisible = False
  With vsObjets
    For n = 1 To .Rows - 1
      If .IsSelected(n) = True Then
        m_bCharge = True
        fldstrnomObjet = .Cell(flexcpText, n, COL_NOM_OBJET)
        fldnCoutTotal = .Cell(flexcpText, n, COL_COUT_TOTAL)
        fldnPoidsBase = .Cell(flexcpText, n, COL_POIDS_BASE)
        cmbClasseArmure = .Cell(flexcpText, n, COL_CLASSE_ARMURE)
        CmbTypeCa = .Cell(flexcpText, n, COL_TYPE_CA)
        fldnEchecSorts = .Cell(flexcpText, n, COL_ECHEC_SORT_PROFANE)
        fldnPenaliteArmure = .Cell(flexcpText, n, COL_PENALITE_ARMURE)
        fldnBonusDexMax = .Cell(flexcpText, n, COL_BONUS_DEX_MAX)
        fldnPortee = .Cell(flexcpText, n, COL_PORTEE)
        fldnsolidite = .Cell(flexcpText, n, COL_SOLIDITE)
        fldnresistance = .Cell(flexcpText, n, COL_RESISTANCE)
        fldnCharge = .Cell(flexcpText, n, COL_CHARGE)
        fldstrdescription = .Cell(flexcpText, n, COL_DESCRIPTION)
        CmbGroupe = .Cell(flexcpText, n, COL_GROUPE_OBJET)
        CmbTailles = .Cell(flexcpText, n, COL_TAILLE)
        CmbType = .Cell(flexcpText, n, COL_DEGAT_TYPE)
        CmbDegatsDes = .Cell(flexcpText, n, COL_DEGAT_DES)
        CmbZoneCrit = .Cell(flexcpText, n, COL_ZONE_CRITIQUE)
        CmbMultipCrit = .Cell(flexcpText, n, COL_MULT_CRIT)
        Set rst = g_bdPerso.OpenRecordset("Select * from ObjetsPropriete where nomObjet='" & FormatStrSQL(fldstrnomObjet) & "'", dbOpenDynaset)
        vsPropri.Rows = 0
        Do While Not rst.EOF
          vsPropri.AddItem rst!propriete_1 & " " & rst!propriete_2 & " " & rst!propriete_3 & " " & rst!valeur & vbTab & rst!propriete_1 & vbTab & rst!propriete_2 & vbTab & rst!propriete_3 & vbTab & rst!valeur
          rst.MoveNext
        Loop
        rst.Close
        bVisible = True
        Exit For
      End If
    Next n
  End With
  m_bCharge = False
  m_bmodif = False
  btnAjouter.Visible = (g_strNomEquipModif = "")
  btnSupprimer.Visible = (bVisible And g_strNomEquipModif = "")
  btnRenommer.Visible = bVisible
  btnEnregistrer.Visible = False
  btnAnnuler.Visible = False
  BtnFermer.Visible = True
  AjusteBouton Me, m_tabButton
End Sub
