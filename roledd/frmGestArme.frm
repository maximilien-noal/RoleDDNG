VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "VSFLEX7D.ocx"
Object = "{0F987290-56EE-11D0-9C43-00A0C90F29FC}#1.0#0"; "ActBar.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmGestArme 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Armes"
   ClientHeight    =   5400
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6480
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5400
   ScaleWidth      =   6480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      Height          =   645
      Left            =   3060
      TabIndex        =   30
      Top             =   -60
      Width           =   3375
      Begin VB.TextBox fldnEchecSorts 
         Height          =   285
         Left            =   2340
         MaxLength       =   2
         TabIndex        =   34
         Top             =   240
         Width           =   825
      End
      Begin VB.TextBox fldnPoids 
         Height          =   285
         Left            =   600
         TabIndex        =   32
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label14 
         Caption         =   "Echec sorts"
         Height          =   285
         Left            =   1410
         TabIndex        =   33
         Top             =   270
         Width           =   855
      End
      Begin VB.Label Label13 
         Caption         =   "Poids"
         Height          =   285
         Left            =   90
         TabIndex        =   31
         Top             =   240
         Width           =   525
      End
   End
   Begin VB.CheckBox ChkDepourvu 
      Height          =   255
      Left            =   2730
      TabIndex        =   29
      Top             =   4230
      Width           =   225
   End
   Begin VB.ComboBox cmbClasseArmure 
      Enabled         =   0   'False
      Height          =   315
      Left            =   1350
      TabIndex        =   28
      Text            =   "Combo1"
      Top             =   4200
      Width           =   1305
   End
   Begin VB.ComboBox CmbNom 
      Height          =   315
      Left            =   870
      Sorted          =   -1  'True
      TabIndex        =   25
      Text            =   "Combo1"
      Top             =   2040
      Width           =   2100
   End
   Begin VB.ComboBox CmbMultipCrit 
      Height          =   315
      Left            =   1335
      TabIndex        =   22
      Text            =   "Combo1"
      Top             =   3840
      Width           =   1650
   End
   Begin VB.ComboBox CmbZoneCrit 
      Height          =   315
      Left            =   1335
      TabIndex        =   20
      Text            =   "Combo1"
      Top             =   3480
      Width           =   1650
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   45
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4590
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.ComboBox CmbDegatsDes 
      Height          =   315
      Left            =   1335
      TabIndex        =   13
      Text            =   "Combo1"
      Top             =   3120
      Width           =   1650
   End
   Begin VB.ComboBox CmbType 
      Height          =   315
      Left            =   870
      TabIndex        =   11
      Text            =   "Combo1"
      Top             =   2760
      Width           =   2115
   End
   Begin VB.ComboBox CmbTailles 
      Height          =   315
      Left            =   870
      TabIndex        =   9
      Text            =   "Combo1"
      Top             =   2400
      Width           =   2100
   End
   Begin VB.TextBox fldstrnomObjet 
      Height          =   285
      Left            =   60
      TabIndex        =   7
      Top             =   1710
      Width           =   2910
   End
   Begin VB.Frame Frame2 
      Caption         =   "Arme pour:"
      Height          =   2235
      Left            =   3060
      TabIndex        =   6
      Top             =   2340
      Width           =   3390
      Begin VSFLEX7DAOCtl.vsFlexGrid vsL1 
         Height          =   1605
         Left            =   90
         TabIndex        =   24
         Top             =   240
         Width           =   3210
         _ExtentX        =   5662
         _ExtentY        =   2831
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
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   "Race                         |Autorisé|Legere"
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
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
         Editable        =   -1  'True
         ShowComboButton =   -1  'True
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Additionnel"
      Height          =   1695
      Left            =   3060
      TabIndex        =   1
      Top             =   630
      Width           =   3375
      Begin VB.ComboBox CmbAddDegatAttaque 
         Height          =   315
         Left            =   1290
         TabIndex        =   18
         Text            =   "Combo1"
         Top             =   570
         Width           =   1995
      End
      Begin VB.ComboBox CmbAddJetAttaque 
         Height          =   315
         Left            =   1290
         TabIndex        =   17
         Text            =   "Combo1"
         Top             =   210
         Width           =   1995
      End
      Begin VB.ComboBox CmbAddTypeDegats 
         Height          =   315
         Left            =   1290
         TabIndex        =   16
         Text            =   "Combo1"
         Top             =   1290
         Width           =   1995
      End
      Begin VB.ComboBox CmbAddDegatsDes 
         Height          =   315
         Left            =   1290
         TabIndex        =   3
         Text            =   "Combo1"
         Top             =   930
         Width           =   1995
      End
      Begin VB.Label Label11 
         Caption         =   "jet attaque"
         Height          =   255
         Left            =   90
         TabIndex        =   23
         Top             =   270
         Width           =   1125
      End
      Begin VB.Label Label8 
         Caption         =   "Type de dégats"
         Height          =   285
         Left            =   90
         TabIndex        =   15
         Top             =   1290
         Width           =   1155
      End
      Begin VB.Label Label7 
         Caption         =   "Degat attaque"
         Height          =   255
         Left            =   90
         TabIndex        =   5
         Top             =   630
         Width           =   1125
      End
      Begin VB.Label Label6 
         Caption         =   "Jet attaque"
         Height          =   255
         Left            =   -180
         TabIndex        =   4
         Top             =   -420
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Dégats aux dès"
         Height          =   225
         Left            =   90
         TabIndex        =   2
         Top             =   990
         Width           =   1185
      End
   End
   Begin VSFLEX7DAOCtl.vsFlexGrid vsArmes 
      Bindings        =   "frmGestArme.frx":0000
      Height          =   1650
      Left            =   60
      TabIndex        =   0
      Top             =   30
      Width           =   2940
      _ExtentX        =   5186
      _ExtentY        =   2910
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
      FormatString    =   "Nom objet                                                                  |Race                 "
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
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
      Editable        =   0   'False
      ShowComboButton =   -1  'True
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   14
      Top             =   5025
      Width           =   6480
      _ExtentX        =   11430
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11377
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "CA,Depourvu"
      Height          =   255
      Left            =   90
      TabIndex        =   27
      Top             =   4200
      Width           =   1245
   End
   Begin VB.Label Label12 
      Caption         =   "Nom"
      Height          =   225
      Left            =   90
      TabIndex        =   26
      Top             =   2040
      Width           =   585
   End
   Begin VB.Label Label10 
      Caption         =   "Multiplicateur crit"
      Height          =   255
      Left            =   90
      TabIndex        =   21
      Top             =   3870
      Width           =   1245
   End
   Begin VB.Label Label9 
      Caption         =   "Zone critique"
      Height          =   255
      Left            =   90
      TabIndex        =   19
      Top             =   3480
      Width           =   1245
   End
   Begin ActiveBarLibraryCtl.ActiveBar ActiveBar1 
      Left            =   1335
      Top             =   4530
      _ExtentX        =   847
      _ExtentY        =   847
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Bands           =   "frmGestArme.frx":0014
   End
   Begin VB.Label Label4 
      Caption         =   "Degats aux dès"
      Height          =   225
      Left            =   90
      TabIndex        =   12
      Top             =   3120
      Width           =   1155
   End
   Begin VB.Label Label3 
      Caption         =   "Type"
      Height          =   225
      Left            =   90
      TabIndex        =   10
      Top             =   2760
      Width           =   585
   End
   Begin VB.Label Label2 
      Caption         =   "Taille"
      Height          =   225
      Left            =   90
      TabIndex        =   8
      Top             =   2400
      Width           =   585
   End
End
Attribute VB_Name = "frmGestArme"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const COL_NOM_OBJET = 0
Const COL_NOM = 1
Const COL_TAILLE = 2
Const COL_TYPE = 3
Const COL_DEGATS_DES = 4
Const COL_ZONE_CRITIQUE = 5
Const COL_MULT_CRITIQUE = 6
Const COL_ADD_JET_ATTAQUE = 7
Const COL_ADD_DEGATS_ATTAQUE = 8
Const COL_ADD_DEGATS_DES = 9
Const COL_ADD_TYPE_DEGATS = 10
Const COL_CA = 11
Const COL_CA_DEPOURVU = 12
Const COL_POIDS = 13
Const COL_ECHEC_SORT = 14

Dim m_bmodif As Boolean
Dim m_bCharge As Boolean


Private Sub ChkDepourvu_Click()

  If g_ModeDebug = vbUnchecked Then On Error GoTo ChkDepourvu_Click_Error
  Modification

ChkDepourvu_Click_Exit:
  Exit Sub

ChkDepourvu_Click_Error:
  MsgBox "L'erreur suivante s'est produite : " & Err.Description & vbCrLf & vbCrLf & " numéro d'erreur (" & Err.Number & " )", vbCritical, g_strNomApplication
  Print #g_FileDebug, Format(Now, "dd/MM/yyyy hh:mm:ss") & " Erreur n° : (" & Err.Number & ") " & Err.Description & " Procedure : ChkDepourvu_Click  Module : frmGestArme.frm"
  Resume ChkDepourvu_Click_Exit
End Sub

Private Sub CmbAddDegatAttaque_Change()

  If g_ModeDebug = vbUnchecked Then On Error GoTo CmbAddDegatAttaque_Change_Error
Modification

CmbAddDegatAttaque_Change_Exit:
  Exit Sub

CmbAddDegatAttaque_Change_Error:
  MsgBox "L'erreur suivante s'est produite : " & Err.Description & vbCrLf & vbCrLf & " numéro d'erreur (" & Err.Number & " )", vbCritical, g_strNomApplication
  Print #g_FileDebug, Format(Now, "dd/MM/yyyy hh:mm:ss") & " Erreur n° : (" & Err.Number & ") " & Err.Description & " Procedure : CmbAddDegatAttaque_Change  Module : frmGestArme.frm"
  Resume CmbAddDegatAttaque_Change_Exit
End Sub

Private Sub CmbAddDegatAttaque_Click()

  If g_ModeDebug = vbUnchecked Then On Error GoTo CmbAddDegatAttaque_Click_Error
Modification

CmbAddDegatAttaque_Click_Exit:
  Exit Sub

CmbAddDegatAttaque_Click_Error:
  MsgBox "L'erreur suivante s'est produite : " & Err.Description & vbCrLf & vbCrLf & " numéro d'erreur (" & Err.Number & " )", vbCritical, g_strNomApplication
  Print #g_FileDebug, Format(Now, "dd/MM/yyyy hh:mm:ss") & " Erreur n° : (" & Err.Number & ") " & Err.Description & " Procedure : CmbAddDegatAttaque_Click  Module : frmGestArme.frm"
  Resume CmbAddDegatAttaque_Click_Exit
End Sub

Private Sub CmbAddDegatsDes_Change()

  If g_ModeDebug = vbUnchecked Then On Error GoTo CmbAddDegatsDes_Change_Error
Modification

CmbAddDegatsDes_Change_Exit:
  Exit Sub

CmbAddDegatsDes_Change_Error:
  MsgBox "L'erreur suivante s'est produite : " & Err.Description & vbCrLf & vbCrLf & " numéro d'erreur (" & Err.Number & " )", vbCritical, g_strNomApplication
  Print #g_FileDebug, Format(Now, "dd/MM/yyyy hh:mm:ss") & " Erreur n° : (" & Err.Number & ") " & Err.Description & " Procedure : CmbAddDegatsDes_Change  Module : frmGestArme.frm"
  Resume CmbAddDegatsDes_Change_Exit
End Sub

Private Sub CmbAddDegatsDes_Click()

  If g_ModeDebug = vbUnchecked Then On Error GoTo CmbAddDegatsDes_Click_Error
Modification

CmbAddDegatsDes_Click_Exit:
  Exit Sub

CmbAddDegatsDes_Click_Error:
  MsgBox "L'erreur suivante s'est produite : " & Err.Description & vbCrLf & vbCrLf & " numéro d'erreur (" & Err.Number & " )", vbCritical, g_strNomApplication
  Print #g_FileDebug, Format(Now, "dd/MM/yyyy hh:mm:ss") & " Erreur n° : (" & Err.Number & ") " & Err.Description & " Procedure : CmbAddDegatsDes_Click  Module : frmGestArme.frm"
  Resume CmbAddDegatsDes_Click_Exit
End Sub

Private Sub CmbAddJetAttaque_Change()

  If g_ModeDebug = vbUnchecked Then On Error GoTo CmbAddJetAttaque_Change_Error
Modification

CmbAddJetAttaque_Change_Exit:
  Exit Sub

CmbAddJetAttaque_Change_Error:
  MsgBox "L'erreur suivante s'est produite : " & Err.Description & vbCrLf & vbCrLf & " numéro d'erreur (" & Err.Number & " )", vbCritical, g_strNomApplication
  Print #g_FileDebug, Format(Now, "dd/MM/yyyy hh:mm:ss") & " Erreur n° : (" & Err.Number & ") " & Err.Description & " Procedure : CmbAddJetAttaque_Change  Module : frmGestArme.frm"
  Resume CmbAddJetAttaque_Change_Exit
End Sub

Private Sub CmbAddJetAttaque_Click()

  If g_ModeDebug = vbUnchecked Then On Error GoTo CmbAddJetAttaque_Click_Error
Modification

CmbAddJetAttaque_Click_Exit:
  Exit Sub

CmbAddJetAttaque_Click_Error:
  MsgBox "L'erreur suivante s'est produite : " & Err.Description & vbCrLf & vbCrLf & " numéro d'erreur (" & Err.Number & " )", vbCritical, g_strNomApplication
  Print #g_FileDebug, Format(Now, "dd/MM/yyyy hh:mm:ss") & " Erreur n° : (" & Err.Number & ") " & Err.Description & " Procedure : CmbAddJetAttaque_Click  Module : frmGestArme.frm"
  Resume CmbAddJetAttaque_Click_Exit
End Sub

Private Sub CmbAddTypeDegats_Change()

  If g_ModeDebug = vbUnchecked Then On Error GoTo CmbAddTypeDegats_Change_Error
Modification

CmbAddTypeDegats_Change_Exit:
  Exit Sub

CmbAddTypeDegats_Change_Error:
  MsgBox "L'erreur suivante s'est produite : " & Err.Description & vbCrLf & vbCrLf & " numéro d'erreur (" & Err.Number & " )", vbCritical, g_strNomApplication
  Print #g_FileDebug, Format(Now, "dd/MM/yyyy hh:mm:ss") & " Erreur n° : (" & Err.Number & ") " & Err.Description & " Procedure : CmbAddTypeDegats_Change  Module : frmGestArme.frm"
  Resume CmbAddTypeDegats_Change_Exit
End Sub

Private Sub CmbAddTypeDegats_Click()

  If g_ModeDebug = vbUnchecked Then On Error GoTo CmbAddTypeDegats_Click_Error
Modification

CmbAddTypeDegats_Click_Exit:
  Exit Sub

CmbAddTypeDegats_Click_Error:
  MsgBox "L'erreur suivante s'est produite : " & Err.Description & vbCrLf & vbCrLf & " numéro d'erreur (" & Err.Number & " )", vbCritical, g_strNomApplication
  Print #g_FileDebug, Format(Now, "dd/MM/yyyy hh:mm:ss") & " Erreur n° : (" & Err.Number & ") " & Err.Description & " Procedure : CmbAddTypeDegats_Click  Module : frmGestArme.frm"
  Resume CmbAddTypeDegats_Click_Exit
End Sub

Private Sub cmbClasseArmure_Change()

  If g_ModeDebug = vbUnchecked Then On Error GoTo cmbClasseArmure_Change_Error
  Modification

cmbClasseArmure_Change_Exit:
  Exit Sub

cmbClasseArmure_Change_Error:
  MsgBox "L'erreur suivante s'est produite : " & Err.Description & vbCrLf & vbCrLf & " numéro d'erreur (" & Err.Number & " )", vbCritical, g_strNomApplication
  Print #g_FileDebug, Format(Now, "dd/MM/yyyy hh:mm:ss") & " Erreur n° : (" & Err.Number & ") " & Err.Description & " Procedure : cmbClasseArmure_Change  Module : frmGestArme.frm"
  Resume cmbClasseArmure_Change_Exit
End Sub

Private Sub cmbClasseArmure_Click()

  If g_ModeDebug = vbUnchecked Then On Error GoTo cmbClasseArmure_Click_Error
Modification

cmbClasseArmure_Click_Exit:
  Exit Sub

cmbClasseArmure_Click_Error:
  MsgBox "L'erreur suivante s'est produite : " & Err.Description & vbCrLf & vbCrLf & " numéro d'erreur (" & Err.Number & " )", vbCritical, g_strNomApplication
  Print #g_FileDebug, Format(Now, "dd/MM/yyyy hh:mm:ss") & " Erreur n° : (" & Err.Number & ") " & Err.Description & " Procedure : cmbClasseArmure_Click  Module : frmGestArme.frm"
  Resume cmbClasseArmure_Click_Exit
End Sub

Private Sub CmbDegatsDes_Change()

  If g_ModeDebug = vbUnchecked Then On Error GoTo CmbDegatsDes_Change_Error
Modification

CmbDegatsDes_Change_Exit:
  Exit Sub

CmbDegatsDes_Change_Error:
  MsgBox "L'erreur suivante s'est produite : " & Err.Description & vbCrLf & vbCrLf & " numéro d'erreur (" & Err.Number & " )", vbCritical, g_strNomApplication
  Print #g_FileDebug, Format(Now, "dd/MM/yyyy hh:mm:ss") & " Erreur n° : (" & Err.Number & ") " & Err.Description & " Procedure : CmbDegatsDes_Change  Module : frmGestArme.frm"
  Resume CmbDegatsDes_Change_Exit
End Sub

Private Sub CmbDegatsDes_Click()

  If g_ModeDebug = vbUnchecked Then On Error GoTo CmbDegatsDes_Click_Error
Modification

CmbDegatsDes_Click_Exit:
  Exit Sub

CmbDegatsDes_Click_Error:
  MsgBox "L'erreur suivante s'est produite : " & Err.Description & vbCrLf & vbCrLf & " numéro d'erreur (" & Err.Number & " )", vbCritical, g_strNomApplication
  Print #g_FileDebug, Format(Now, "dd/MM/yyyy hh:mm:ss") & " Erreur n° : (" & Err.Number & ") " & Err.Description & " Procedure : CmbDegatsDes_Click  Module : frmGestArme.frm"
  Resume CmbDegatsDes_Click_Exit
End Sub

Private Sub CmbMultipCrit_Change()

  If g_ModeDebug = vbUnchecked Then On Error GoTo CmbMultipCrit_Change_Error
Modification

CmbMultipCrit_Change_Exit:
  Exit Sub

CmbMultipCrit_Change_Error:
  MsgBox "L'erreur suivante s'est produite : " & Err.Description & vbCrLf & vbCrLf & " numéro d'erreur (" & Err.Number & " )", vbCritical, g_strNomApplication
  Print #g_FileDebug, Format(Now, "dd/MM/yyyy hh:mm:ss") & " Erreur n° : (" & Err.Number & ") " & Err.Description & " Procedure : CmbMultipCrit_Change  Module : frmGestArme.frm"
  Resume CmbMultipCrit_Change_Exit
End Sub

Private Sub CmbMultipCrit_Click()

  If g_ModeDebug = vbUnchecked Then On Error GoTo CmbMultipCrit_Click_Error
Modification

CmbMultipCrit_Click_Exit:
  Exit Sub

CmbMultipCrit_Click_Error:
  MsgBox "L'erreur suivante s'est produite : " & Err.Description & vbCrLf & vbCrLf & " numéro d'erreur (" & Err.Number & " )", vbCritical, g_strNomApplication
  Print #g_FileDebug, Format(Now, "dd/MM/yyyy hh:mm:ss") & " Erreur n° : (" & Err.Number & ") " & Err.Description & " Procedure : CmbMultipCrit_Click  Module : frmGestArme.frm"
  Resume CmbMultipCrit_Click_Exit
End Sub

Private Sub CmbNom_Change()

  If g_ModeDebug = vbUnchecked Then On Error GoTo CmbNom_Change_Error
  Modification

CmbNom_Change_Exit:
  Exit Sub

CmbNom_Change_Error:
  MsgBox "L'erreur suivante s'est produite : " & Err.Description & vbCrLf & vbCrLf & " numéro d'erreur (" & Err.Number & " )", vbCritical, g_strNomApplication
  Print #g_FileDebug, Format(Now, "dd/MM/yyyy hh:mm:ss") & " Erreur n° : (" & Err.Number & ") " & Err.Description & " Procedure : CmbNom_Change  Module : frmGestArme.frm"
  Resume CmbNom_Change_Exit
End Sub

Private Sub CmbNom_Click()
Dim strnom As String

  If g_ModeDebug = vbUnchecked Then On Error GoTo CmbNom_Click_Error
  
  Modification
  
  strnom = UCase(CmbNom)
  If strnom = "ECU" Or strnom = "PAVOIS" Or strnom = "TARGE" Or strnom = "RONDACHE" Then
    cmbClasseArmure.Enabled = True
  Else
    cmbClasseArmure.Enabled = False
  End If
  ChkDepourvu.Enabled = cmbClasseArmure.Enabled

CmbNom_Click_Exit:
  Exit Sub

CmbNom_Click_Error:
  MsgBox "L'erreur suivante s'est produite : " & Err.Description & vbCrLf & vbCrLf & " numéro d'erreur (" & Err.Number & " )", vbCritical, g_strNomApplication
  Print #g_FileDebug, Format(Now, "dd/MM/yyyy hh:mm:ss") & " Erreur n° : (" & Err.Number & ") " & Err.Description & " Procedure : CmbNom_Click  Module : frmGestArme.frm"
  Resume CmbNom_Click_Exit
End Sub

Private Sub cmbTailles_Change()

  If g_ModeDebug = vbUnchecked Then On Error GoTo cmbTailles_Change_Error

  CmbNom_Click


cmbTailles_Change_Exit:
  Exit Sub

cmbTailles_Change_Error:
  MsgBox "L'erreur suivante s'est produite : " & Err.Description & vbCrLf & vbCrLf & " numéro d'erreur (" & Err.Number & " )", vbCritical, g_strNomApplication
  Print #g_FileDebug, Format(Now, "dd/MM/yyyy hh:mm:ss") & " Erreur n° : (" & Err.Number & ") " & Err.Description & " Procedure : cmbTailles_Change  Module : frmGestArme.frm"
  Resume cmbTailles_Change_Exit
End Sub

Private Sub CmbTailles_Click()

  If g_ModeDebug = vbUnchecked Then On Error GoTo CmbTailles_Click_Error
  Modification

CmbTailles_Click_Exit:
  Exit Sub

CmbTailles_Click_Error:
  MsgBox "L'erreur suivante s'est produite : " & Err.Description & vbCrLf & vbCrLf & " numéro d'erreur (" & Err.Number & " )", vbCritical, g_strNomApplication
  Print #g_FileDebug, Format(Now, "dd/MM/yyyy hh:mm:ss") & " Erreur n° : (" & Err.Number & ") " & Err.Description & " Procedure : CmbTailles_Click  Module : frmGestArme.frm"
  Resume CmbTailles_Click_Exit
End Sub

Private Sub cmbTailles_GotFocus()

  If g_ModeDebug = vbUnchecked Then On Error GoTo cmbTailles_GotFocus_Error

  Selec
  StatusBar1.Panels(1).Text = "Saisissez l'enrichissement du contrat"


cmbTailles_GotFocus_Exit:
  Exit Sub

cmbTailles_GotFocus_Error:
  MsgBox "L'erreur suivante s'est produite : " & Err.Description & vbCrLf & vbCrLf & " numéro d'erreur (" & Err.Number & " )", vbCritical, g_strNomApplication
  Print #g_FileDebug, Format(Now, "dd/MM/yyyy hh:mm:ss") & " Erreur n° : (" & Err.Number & ") " & Err.Description & " Procedure : cmbTailles_GotFocus  Module : frmGestArme.frm"
  Resume cmbTailles_GotFocus_Exit
End Sub
Private Sub cmbTailles_LostFocus()

  If g_ModeDebug = vbUnchecked Then On Error GoTo cmbTailles_LostFocus_Error

  StatusBar1.Panels(1).Text = ""


cmbTailles_LostFocus_Exit:
  Exit Sub

cmbTailles_LostFocus_Error:
  MsgBox "L'erreur suivante s'est produite : " & Err.Description & vbCrLf & vbCrLf & " numéro d'erreur (" & Err.Number & " )", vbCritical, g_strNomApplication
  Print #g_FileDebug, Format(Now, "dd/MM/yyyy hh:mm:ss") & " Erreur n° : (" & Err.Number & ") " & Err.Description & " Procedure : cmbTailles_LostFocus  Module : frmGestArme.frm"
  Resume cmbTailles_LostFocus_Exit
End Sub

Sub Modification()

  If g_ModeDebug = vbUnchecked Then On Error GoTo Modification_Error

  
  With ActiveBar1.Bands(0)
    If m_bCharge = False Then
      .Tools("TBAjouter").Visible = False
      .Tools("TBSupprimer").Visible = False
      .Tools("TBEnregistrer").Visible = True
      .Tools("TBAnnuler").Visible = True
      .Tools("TBFermer").Visible = False
    End If
  End With
  ActiveBar1.RecalcLayout
  m_bmodif = True


Modification_Exit:
  Exit Sub

Modification_Error:
  MsgBox "L'erreur suivante s'est produite : " & Err.Description & vbCrLf & vbCrLf & " numéro d'erreur (" & Err.Number & " )", vbCritical, g_strNomApplication
  Print #g_FileDebug, Format(Now, "dd/MM/yyyy hh:mm:ss") & " Erreur n° : (" & Err.Number & ") " & Err.Description & " Procedure : Modification  Module : frmGestArme.frm"
  Resume Modification_Exit
End Sub
Private Function NonModifier() As Boolean
  
  NonModifier = False
  If m_bmodif = True Then
    Select Case MsgBox("Voulez vous enregistrer les modifications apportées à cette arme", vbYesNoCancel Or vbInformation, Me.Caption)
      Case vbYes
        If Enregistre() = False Then Exit Function
      Case vbCancel
        Call vsArmes_SelChange
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
Dim rstArmes As Recordset
Dim cmb As ComboBox
Dim name As String
Dim bObligation As Boolean

  Enregistre = False
  If Len(Trim(fldstrnomObjet)) < 1 Then
    MsgBox "Saisie d'un nom d'objet obligatoire", vbExclamation, Me.Caption
    Call Set_Focus(fldstrnomObjet)
    Exit Function
  End If
  
  For i = 0 To Me.Count - 1
    name = Me(i).name
    If UCase(Left(name, 3)) = "CMB" Then
      Set cmb = Me(i)
      bObligation = False
      If name = "CmbNom" Or _
          name = "CmbTailles" Or _
          name = "CmbType" Or _
          name = "CmbDegatsDes" Or _
          name = "CmbZoneCrit" Or _
          name = "CmbMultipCrit" Then
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
          MsgBox "Saisie obligatoire d'un élément de la liste", vbExclamation, Me.Caption
          cmb.SetFocus
          Exit Function
        End If
      End If
    End If
  Next i
  
  If Trim(fldnPoids) <> "" Then
    fldnPoids = FrmtNum(fldnPoids)
    If Not IsNumeric(fldnPoids) Then
      MsgBox "Saisie d'un poids numérique", vbExclamation, Me.Caption
      Set_Focus fldnPoids
      Exit Function
    End If
  End If
  
  If Trim(fldnEchecSorts) <> "" Then
    fldnEchecSorts = FrmtNum(fldnEchecSorts)
    If Not IsNumeric(fldnEchecSorts) Then
      MsgBox "Saisie d'un echec au sort numérique", vbExclamation, Me.Caption
      Set_Focus fldnEchecSorts
      Exit Function
    End If
  End If

      
  Screen.MousePointer = vbHourglass
  Set rstArmes = g_dbs.OpenRecordset("Arme", dbOpenTable)
  With rstArmes
    .Index = "primarykey"
    .Seek "=", fldstrnomObjet
    If .NoMatch Then
      .AddNew
    Else
      .Edit
    End If
    !nomobjet = fldstrnomObjet
    !nom = CmbNom
    !Taille = CmbTailles
    !Type = CmbType
    !Degats_des = CmbDegatsDes
    !zone_critique = CmbZoneCrit
    !mult_critique = CmbMultipCrit
    !Add_Jet_attaque = IIf(Trim(CmbAddJetAttaque) = "", Null, CmbAddJetAttaque)
    !Add_degats_attaque = IIf(Trim(CmbAddDegatAttaque) = "", Null, CmbAddDegatAttaque)
    !Add_Degats_des = IIf(Trim(CmbAddDegatsDes) = "", Null, CmbAddDegatsDes)
    !add_type_degats = IIf(Trim(CmbAddTypeDegats) = "", Null, CmbAddTypeDegats)
    !ca = IIf(Trim(cmbClasseArmure) = "" Or cmbClasseArmure.Enabled = False, Null, cmbClasseArmure)
    !ca_depourvu = IIf(Trim(cmbClasseArmure) = "" Or cmbClasseArmure.Enabled = False, False, ChkDepourvu)
    !Poids = IIf(IsNumeric(fldnPoids), fldnPoids, Null)
    !echec_sort = IIf(IsNumeric(fldnEchecSorts), fldnEchecSorts, Null)
    .Update
    .Close
  End With
  vsArmes.Redraw = False
  g_dbs.Execute "delete from ArmeRace where nomArme='" & FormatStrSQL(fldstrnomObjet) & "'", dbFailOnError
  For i = 1 To vsL1.Rows - 1
    If vsL1.Cell(flexcpValue, i, 1) = True Then
      g_dbs.Execute "insert into ArmeRace (nomArme,Autorise,race) values ('" & FormatStrSQL(fldstrnomObjet) & "','N','" & vsL1.Cell(flexcpText, i, 0) & "')", dbFailOnError
      If vsL1.Cell(flexcpValue, i, 2) = True Then
        g_dbs.Execute "insert into ArmeRace (nomArme,Autorise,race) values ('" & FormatStrSQL(fldstrnomObjet) & "','L','" & vsL1.Cell(flexcpText, i, 0) & "')", dbFailOnError
      End If
    End If
  Next i
  RemplieLst
  RetrouvVSFLex vsArmes, COL_NOM_OBJET, fldstrnomObjet, 0
  vsArmes.Redraw = True
  m_bmodif = False
  m_bCharge = False
  Enregistre = True
End Function
Private Sub Supprimer()
Dim n As Integer
Dim rstTest As Recordset
Dim bexist As Boolean

  If g_ModeDebug = vbUnchecked Then On Error GoTo Supprimer_Error
 
  With vsArmes
    For n = 1 To .Rows - 1
      If .IsSelected(n) Then
        If MsgBox("Confirmer la suppression de cette arme :" & .Cell(flexcpText, n, COL_NOM_OBJET), vbDefaultButton2 Or vbOKCancel Or vbQuestion, "Suppression") = vbOK Then
          Set rstTest = g_dbs.OpenRecordset("select * from Equipement where nomobjet='" & FormatStrSQL(.Cell(flexcpText, n, COL_NOM_OBJET)) & "'", dbOpenDynaset)
          bexist = Not rstTest.EOF
          rstTest.Close
          If bexist Then
            MsgBox "Certains personnages sont équipés de cet arme, elle ne peut être detruite", vbInformation, Me.Caption
            Exit Sub
          End If
          ''
          g_dbs.Execute "delete from arme where [nomobjet]='" & FormatStrSQL(.Cell(flexcpText, n, COL_NOM_OBJET)) & "'", dbFailOnError
          Call RemplieLst
          If vsArmes.Rows > 1 Then
            vsArmes.IsSelected(1) = True
          Else
            Efface Me
            Call vsArmes_SelChange
          End If
        End If
        Exit For
      End If
    Next n
  End With
  

Supprimer_Exit:
  Exit Sub

Supprimer_Error:
  MsgBox "L'erreur suivante s'est produite : " & Err.Description & vbCrLf & vbCrLf & " numéro d'erreur (" & Err.Number & " )", vbCritical, g_strNomApplication
  Print #g_FileDebug, Format(Now, "dd/MM/yyyy hh:mm:ss") & " Erreur n° : (" & Err.Number & ") " & Err.Description & " Procedure : Supprimer  Module : frmGestArme.frm"
  Resume Supprimer_Exit
End Sub
Private Sub Fermer()

  If g_ModeDebug = vbUnchecked Then On Error GoTo Fermer_Error

  If ActiveBar1.Bands(0).Tools("TBFermer").Visible = True Then
    Unload Me
  Else
    Call vsArmes_SelChange
  End If


Fermer_Exit:
  Exit Sub

Fermer_Error:
  MsgBox "L'erreur suivante s'est produite : " & Err.Description & vbCrLf & vbCrLf & " numéro d'erreur (" & Err.Number & " )", vbCritical, g_strNomApplication
  Print #g_FileDebug, Format(Now, "dd/MM/yyyy hh:mm:ss") & " Erreur n° : (" & Err.Number & ") " & Err.Description & " Procedure : Fermer  Module : frmGestArme.frm"
  Resume Fermer_Exit
End Sub
Private Sub ActiveBar1_Click(ByVal Tool As ActiveBarLibraryCtl.Tool)
Dim i As Integer

  If g_ModeDebug = vbUnchecked Then On Error GoTo ActiveBar1_Click_Error

  Select Case Tool.name
    Case "TBFermer"
      Fermer
    Case "TBEnregistrer"
      Enregistre
    Case "TBAnnuler"
      Fermer
    Case "TBAjouter"
      'ajout
      Efface Me
      For i = 1 To vsL1.Rows - 1
        vsL1.Cell(flexcpText, i, 1) = True
        vsL1.Cell(flexcpText, i, 2) = False
      Next i
      Set_Focus fldstrnomObjet
      
    Case "TBSupprimer"
      Supprimer
  End Select


ActiveBar1_Click_Exit:
  Exit Sub

ActiveBar1_Click_Error:
  MsgBox "L'erreur suivante s'est produite : " & Err.Description & vbCrLf & vbCrLf & " numéro d'erreur (" & Err.Number & " )", vbCritical, g_strNomApplication
  Print #g_FileDebug, Format(Now, "dd/MM/yyyy hh:mm:ss") & " Erreur n° : (" & Err.Number & ") " & Err.Description & " Procedure : ActiveBar1_Click  Module : frmGestArme.frm"
  Resume ActiveBar1_Click_Exit
End Sub

Private Sub CmbType_Change()

  If g_ModeDebug = vbUnchecked Then On Error GoTo CmbType_Change_Error
  Modification

CmbType_Change_Exit:
  Exit Sub

CmbType_Change_Error:
  MsgBox "L'erreur suivante s'est produite : " & Err.Description & vbCrLf & vbCrLf & " numéro d'erreur (" & Err.Number & " )", vbCritical, g_strNomApplication
  Print #g_FileDebug, Format(Now, "dd/MM/yyyy hh:mm:ss") & " Erreur n° : (" & Err.Number & ") " & Err.Description & " Procedure : CmbType_Change  Module : frmGestArme.frm"
  Resume CmbType_Change_Exit
End Sub

Private Sub CmbType_Click()

  If g_ModeDebug = vbUnchecked Then On Error GoTo CmbType_Click_Error
  
  Modification

CmbType_Click_Exit:
  Exit Sub

CmbType_Click_Error:
  MsgBox "L'erreur suivante s'est produite : " & Err.Description & vbCrLf & vbCrLf & " numéro d'erreur (" & Err.Number & " )", vbCritical, g_strNomApplication
  Print #g_FileDebug, Format(Now, "dd/MM/yyyy hh:mm:ss") & " Erreur n° : (" & Err.Number & ") " & Err.Description & " Procedure : CmbType_Click  Module : frmGestArme.frm"
  Resume CmbType_Click_Exit
End Sub

Private Sub CmbZoneCrit_Change()

  If g_ModeDebug = vbUnchecked Then On Error GoTo CmbZoneCrit_Change_Error
Modification

CmbZoneCrit_Change_Exit:
  Exit Sub

CmbZoneCrit_Change_Error:
  MsgBox "L'erreur suivante s'est produite : " & Err.Description & vbCrLf & vbCrLf & " numéro d'erreur (" & Err.Number & " )", vbCritical, g_strNomApplication
  Print #g_FileDebug, Format(Now, "dd/MM/yyyy hh:mm:ss") & " Erreur n° : (" & Err.Number & ") " & Err.Description & " Procedure : CmbZoneCrit_Change  Module : frmGestArme.frm"
  Resume CmbZoneCrit_Change_Exit
End Sub

Private Sub CmbZoneCrit_Click()

  If g_ModeDebug = vbUnchecked Then On Error GoTo CmbZoneCrit_Click_Error
Modification

CmbZoneCrit_Click_Exit:
  Exit Sub

CmbZoneCrit_Click_Error:
  MsgBox "L'erreur suivante s'est produite : " & Err.Description & vbCrLf & vbCrLf & " numéro d'erreur (" & Err.Number & " )", vbCritical, g_strNomApplication
  Print #g_FileDebug, Format(Now, "dd/MM/yyyy hh:mm:ss") & " Erreur n° : (" & Err.Number & ") " & Err.Description & " Procedure : CmbZoneCrit_Click  Module : frmGestArme.frm"
  Resume CmbZoneCrit_Click_Exit
End Sub

Private Sub fldnEchecSorts_Change()
  Modification
End Sub

Private Sub fldnEchecSorts_GotFocus()
  Selec
End Sub

Private Sub fldnPoids_Change()
  Modification
End Sub

Private Sub fldnPoids_GotFocus()
  Selec
End Sub

Private Sub fldstrnomobjet_Change()

  If g_ModeDebug = vbUnchecked Then On Error GoTo fldstrnomobjet_Change_Error

  Modification


fldstrnomobjet_Change_Exit:
  Exit Sub

fldstrnomobjet_Change_Error:
  MsgBox "L'erreur suivante s'est produite : " & Err.Description & vbCrLf & vbCrLf & " numéro d'erreur (" & Err.Number & " )", vbCritical, g_strNomApplication
  Print #g_FileDebug, Format(Now, "dd/MM/yyyy hh:mm:ss") & " Erreur n° : (" & Err.Number & ") " & Err.Description & " Procedure : fldstrnomobjet_Change  Module : frmGestArme.frm"
  Resume fldstrnomobjet_Change_Exit
End Sub
Private Sub fldstrnomobjet_GotFocus()

  If g_ModeDebug = vbUnchecked Then On Error GoTo fldstrnomobjet_GotFocus_Error

  Selec
  StatusBar1.Panels(1).Text = "Saisissez l'identifiant de l'arme ex (Epée longue de ....)"


fldstrnomobjet_GotFocus_Exit:
  Exit Sub

fldstrnomobjet_GotFocus_Error:
  MsgBox "L'erreur suivante s'est produite : " & Err.Description & vbCrLf & vbCrLf & " numéro d'erreur (" & Err.Number & " )", vbCritical, g_strNomApplication
  Print #g_FileDebug, Format(Now, "dd/MM/yyyy hh:mm:ss") & " Erreur n° : (" & Err.Number & ") " & Err.Description & " Procedure : fldstrnomobjet_GotFocus  Module : frmGestArme.frm"
  Resume fldstrnomobjet_GotFocus_Exit
End Sub
Private Sub fldstrnomobjet_LostFocus()

  If g_ModeDebug = vbUnchecked Then On Error GoTo fldstrnomobjet_LostFocus_Error

  StatusBar1.Panels(1).Text = ""


fldstrnomobjet_LostFocus_Exit:
  Exit Sub

fldstrnomobjet_LostFocus_Error:
  MsgBox "L'erreur suivante s'est produite : " & Err.Description & vbCrLf & vbCrLf & " numéro d'erreur (" & Err.Number & " )", vbCritical, g_strNomApplication
  Print #g_FileDebug, Format(Now, "dd/MM/yyyy hh:mm:ss") & " Erreur n° : (" & Err.Number & ") " & Err.Description & " Procedure : fldstrnomobjet_LostFocus  Module : frmGestArme.frm"
  Resume fldstrnomobjet_LostFocus_Exit
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

  If g_ModeDebug = vbUnchecked Then On Error GoTo Form_KeyDown_Error

  Select Case KeyCode
    Case vbKeyF2
      If ActiveBar1.Bands(0).Tools("TBEnregistrer").Visible = True Then Call ActiveBar1_Click(ActiveBar1.Bands(0).Tools("TBEnregistrer"))
    Case vbKeyReturn
      SendKeys "{TAB}"
    Case vbKeyEscape
      Fermer
    Case vbKeyF9
      If ActiveBar1.Bands(0).Tools("TBAjouter").Visible = True Then Call ActiveBar1_Click(ActiveBar1.Bands(0).Tools("TBAjouter"))
    Case vbKeyF10
      If ActiveBar1.Bands(0).Tools("TBSupprimer").Visible = True Then Supprimer
  End Select


Form_KeyDown_Exit:
  Exit Sub

Form_KeyDown_Error:
  MsgBox "L'erreur suivante s'est produite : " & Err.Description & vbCrLf & vbCrLf & " numéro d'erreur (" & Err.Number & " )", vbCritical, g_strNomApplication
  Print #g_FileDebug, Format(Now, "dd/MM/yyyy hh:mm:ss") & " Erreur n° : (" & Err.Number & ") " & Err.Description & " Procedure : Form_KeyDown  Module : frmGestArme.frm"
  Resume Form_KeyDown_Exit
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)

  If g_ModeDebug = vbUnchecked Then On Error GoTo Form_KeyPress_Error
  
  Select Case KeyAscii
    Case vbKeyEscape
      KeyAscii = 0
    Case vbKeyReturn
      KeyAscii = 0
  End Select

Form_KeyPress_Exit:
  Exit Sub

Form_KeyPress_Error:
  MsgBox "L'erreur suivante s'est produite : " & Err.Description & vbCrLf & vbCrLf & " numéro d'erreur (" & Err.Number & " )", vbCritical, g_strNomApplication
  Print #g_FileDebug, Format(Now, "dd/MM/yyyy hh:mm:ss") & " Erreur n° : (" & Err.Number & ") " & Err.Description & " Procedure : Form_KeyPress  Module : frmGestArme.frm"
  Resume Form_KeyPress_Exit
End Sub

Private Sub Form_Load()
Dim rst As Recordset
Dim i As Integer

  If g_ModeDebug = vbUnchecked Then On Error GoTo Form_Load_Error
  
  
  CmbNom.Clear
  CmbNom.AddItem "Cimetère"
  CmbNom.AddItem "Dague"
  CmbNom.AddItem "Epée courte"
  CmbNom.AddItem "Epée longue"
  CmbNom.AddItem "Gourdin"
  CmbNom.AddItem "Hache Double"
  CmbNom.AddItem "Hache de nain"
  CmbNom.AddItem "Targe"
  CmbNom.AddItem "Rondache"
  CmbNom.AddItem "Ecu"
  CmbNom.AddItem "Pavoi"
  
  CmbTailles.Clear
  CmbTailles.AddItem "Petite"
  CmbTailles.AddItem "Moyenne"
  CmbTailles.AddItem "Grande"
  
  CmbType.Clear
  CmbType.AddItem "Piquant"
  CmbType.AddItem "Tranchant"
  CmbType.AddItem "Contondant"
  
  CmbZoneCrit.Clear
  CmbZoneCrit.AddItem "20"
  CmbZoneCrit.AddItem "19"
  CmbZoneCrit.AddItem "18"
  
  CmbMultipCrit.Clear
  CmbMultipCrit.AddItem "2"
  CmbMultipCrit.AddItem "3"
  CmbMultipCrit.AddItem "4"
  
  CmbDegatsDes.Clear
  CmbDegatsDes.AddItem "1d4"
  CmbDegatsDes.AddItem "1d6"
  CmbDegatsDes.AddItem "1d8"
  CmbDegatsDes.AddItem "1d10"
  CmbDegatsDes.AddItem "1d12"
  CmbDegatsDes.AddItem "2d4"
  CmbDegatsDes.AddItem "2d6"
  CmbDegatsDes.AddItem "2d8"
  
  cmbClasseArmure.Clear
  For i = 1 To 10: cmbClasseArmure.AddItem i: Next i
  
  CmbAddDegatsDes.Clear
  CmbAddDegatsDes.AddItem "1d4"
  CmbAddDegatsDes.AddItem "1d6"
  CmbAddDegatsDes.AddItem "1d8"
  CmbAddDegatsDes.AddItem "1d10"
  CmbAddDegatsDes.AddItem "1d12"
  CmbAddDegatsDes.AddItem "2d4"
  CmbAddDegatsDes.AddItem "2d6"
  CmbAddDegatsDes.AddItem "2d8"
  
  CmbAddTypeDegats.Clear
  CmbAddTypeDegats.AddItem "Feu"
  CmbAddTypeDegats.AddItem "Froid"
  CmbAddTypeDegats.AddItem "Acide"
  CmbAddTypeDegats.AddItem "Son"

  CmbAddJetAttaque.Clear
  CmbAddJetAttaque.AddItem "1"
  CmbAddJetAttaque.AddItem "2"
  CmbAddJetAttaque.AddItem "3"
  CmbAddJetAttaque.AddItem "4"
  CmbAddJetAttaque.AddItem "5"
  
  CmbAddDegatAttaque.Clear
  CmbAddDegatAttaque.AddItem "1"
  CmbAddDegatAttaque.AddItem "2"
  CmbAddDegatAttaque.AddItem "3"
  CmbAddDegatAttaque.AddItem "4"
  CmbAddDegatAttaque.AddItem "5"

  vsL1.ColDataType(1) = flexDTBoolean
  vsL1.ColDataType(2) = flexDTBoolean

  
  
  vsL1.Rows = 1
  Set rst = g_dbs.OpenRecordset("select * from race order by race", dbOpenDynaset)
  Do While Not rst.EOF
    vsL1.AddItem rst!race
    rst.MoveNext
  Loop
  rst.Close
   ' Ouverture base de donnée
  Data1.DatabaseName = g_dbs.name
  Call RemplieLst
  If vsArmes.Rows > 1 Then
    vsArmes.IsSelected(1) = True
  Else
    Call vsArmes_SelChange
  End If

Form_Load_Exit:
  Exit Sub

Form_Load_Error:
  MsgBox "L'erreur suivante s'est produite : " & Err.Description & vbCrLf & vbCrLf & " numéro d'erreur (" & Err.Number & " )", vbCritical, g_strNomApplication
  Print #g_FileDebug, Format(Now, "dd/MM/yyyy hh:mm:ss") & " Erreur n° : (" & Err.Number & ") " & Err.Description & " Procedure : Form_Load  Module : frmGestArme.frm"
  Resume Form_Load_Exit
End Sub
Sub RemplieLst()
Dim i As Integer
Dim stSql As String

  If g_ModeDebug = vbUnchecked Then On Error GoTo RemplieLst_Error
 
  Screen.MousePointer = vbHourglass
  stSql = "SELECT nomobjet,nom,taille,type,Degats_des,zone_critique,mult_critique,Add_Jet_attaque,Add_degats_attaque,Add_Degats_des,Add_type_degats,ca,ca_depourvu,poids,echec_sort from arme "
  Data1.RecordSource = stSql
  Data1.Refresh
  With vsArmes
    .Redraw = False
    .DataRefresh
    .FormatString = "Nom  "
    .ColAlignment(COL_NOM_OBJET) = flexAlignLeftCenter
    .Cols = 15
    For i = COL_NOM To COL_ECHEC_SORT
      .ColHidden(i) = True
    Next i
    .Redraw = True
  End With
  Screen.MousePointer = vbNormal

RemplieLst_Exit:
  Exit Sub

RemplieLst_Error:
  MsgBox "L'erreur suivante s'est produite : " & Err.Description & vbCrLf & vbCrLf & " numéro d'erreur (" & Err.Number & " )", vbCritical, g_strNomApplication
  Print #g_FileDebug, Format(Now, "dd/MM/yyyy hh:mm:ss") & " Erreur n° : (" & Err.Number & ") " & Err.Description & " Procedure : RemplieLst  Module : frmGestArme.frm"
  Resume RemplieLst_Exit
End Sub
Private Sub Form_Unload(Cancel As Integer)

  If g_ModeDebug = vbUnchecked Then On Error GoTo Form_Unload_Error

  If NonModifier() = False Then Cancel = True: Exit Sub


Form_Unload_Exit:
  Exit Sub

Form_Unload_Error:
  MsgBox "L'erreur suivante s'est produite : " & Err.Description & vbCrLf & vbCrLf & " numéro d'erreur (" & Err.Number & " )", vbCritical, g_strNomApplication
  Print #g_FileDebug, Format(Now, "dd/MM/yyyy hh:mm:ss") & " Erreur n° : (" & Err.Number & ") " & Err.Description & " Procedure : Form_Unload  Module : frmGestArme.frm"
  Resume Form_Unload_Exit
End Sub

Private Sub vsArmes_AfterEdit(ByVal Row As Long, ByVal Col As Long)

  If g_ModeDebug = vbUnchecked Then On Error GoTo vsArmes_AfterEdit_Error
  Modification

vsArmes_AfterEdit_Exit:
  Exit Sub

vsArmes_AfterEdit_Error:
  MsgBox "L'erreur suivante s'est produite : " & Err.Description & vbCrLf & vbCrLf & " numéro d'erreur (" & Err.Number & " )", vbCritical, g_strNomApplication
  Print #g_FileDebug, Format(Now, "dd/MM/yyyy hh:mm:ss") & " Erreur n° : (" & Err.Number & ") " & Err.Description & " Procedure : vsArmes_AfterEdit  Module : frmGestArme.frm"
  Resume vsArmes_AfterEdit_Exit
End Sub

Private Sub vsArmes_Click()

  If g_ModeDebug = vbUnchecked Then On Error GoTo vsArmes_Click_Error

  vsArmes_SelChange


vsArmes_Click_Exit:
  Exit Sub

vsArmes_Click_Error:
  MsgBox "L'erreur suivante s'est produite : " & Err.Description & vbCrLf & vbCrLf & " numéro d'erreur (" & Err.Number & " )", vbCritical, g_strNomApplication
  Print #g_FileDebug, Format(Now, "dd/MM/yyyy hh:mm:ss") & " Erreur n° : (" & Err.Number & ") " & Err.Description & " Procedure : vsArmes_Click  Module : frmGestArme.frm"
  Resume vsArmes_Click_Exit
End Sub
Private Sub vsArmes_SelChange()
Dim i, n As Integer
Dim bVisible As Boolean
Dim tblArmeRace As Recordset

  If g_ModeDebug = vbUnchecked Then On Error GoTo vsArmes_SelChange_Error
  
  Set tblArmeRace = g_dbs.OpenRecordset("ArmeRace", dbOpenTable)
  tblArmeRace.Index = "index"

  bVisible = False
  With vsArmes
    For n = 1 To .Rows - 1
      If .IsSelected(n) = True Then
        m_bCharge = True
        fldstrnomObjet = .Cell(flexcpText, n, COL_NOM_OBJET)
        CmbNom = .Cell(flexcpText, n, COL_NOM)
        CmbTailles = .Cell(flexcpText, n, COL_TAILLE)
        CmbType = .Cell(flexcpText, n, COL_TYPE)
        CmbDegatsDes = .Cell(flexcpText, n, COL_DEGATS_DES)
        CmbZoneCrit = .Cell(flexcpText, n, COL_ZONE_CRITIQUE)
        CmbMultipCrit = .Cell(flexcpText, n, COL_MULT_CRITIQUE)
        CmbAddJetAttaque = .Cell(flexcpText, n, COL_ADD_JET_ATTAQUE)
        CmbAddDegatAttaque = .Cell(flexcpText, n, COL_ADD_DEGATS_ATTAQUE)
        CmbAddDegatsDes = .Cell(flexcpText, n, COL_ADD_DEGATS_DES)
        CmbAddTypeDegats = .Cell(flexcpText, n, COL_ADD_TYPE_DEGATS)
        cmbClasseArmure = .Cell(flexcpText, n, COL_CA)
        ChkDepourvu = IIf(.Cell(flexcpText, n, COL_CA_DEPOURVU), vbChecked, vbUnchecked)
        fldnPoids = .Cell(flexcpText, n, COL_POIDS)
        fldnEchecSorts = .Cell(flexcpText, n, COL_ECHEC_SORT)
        For i = 1 To vsL1.Rows - 1
          tblArmeRace.Seek "=", fldstrnomObjet, "L", vsL1.Cell(flexcpText, i, 0)
          vsL1.Cell(flexcpText, i, 2) = IIf(tblArmeRace.NoMatch, False, True)
          tblArmeRace.Seek "=", fldstrnomObjet, "N", vsL1.Cell(flexcpText, i, 0)
          vsL1.Cell(flexcpText, i, 1) = IIf(tblArmeRace.NoMatch, False, True)
        Next i
        bVisible = True
        Exit For
      End If
    Next n
  End With
  tblArmeRace.Close
  m_bCharge = False
  m_bmodif = False
  With ActiveBar1.Bands(0)
    .Tools("TBAjouter").Visible = True
    .Tools("TBSupprimer").Visible = bVisible
    .Tools("TBEnregistrer").Visible = False
    .Tools("TBAnnuler").Visible = False
    .Tools("TBFermer").Visible = True
  End With
  ActiveBar1.RecalcLayout


vsArmes_SelChange_Exit:
  Exit Sub

vsArmes_SelChange_Error:
  MsgBox "L'erreur suivante s'est produite : " & Err.Description & vbCrLf & vbCrLf & " numéro d'erreur (" & Err.Number & " )", vbCritical, g_strNomApplication
  Print #g_FileDebug, Format(Now, "dd/MM/yyyy hh:mm:ss") & " Erreur n° : (" & Err.Number & ") " & Err.Description & " Procedure : vsArmes_SelChange  Module : frmGestArme.frm"
  Resume vsArmes_SelChange_Exit
End Sub

Private Sub vsL1_AfterEdit(ByVal Row As Long, ByVal Col As Long)

  If g_ModeDebug = vbUnchecked Then On Error GoTo vsL1_AfterEdit_Error
  Modification

vsL1_AfterEdit_Exit:
  Exit Sub

vsL1_AfterEdit_Error:
  MsgBox "L'erreur suivante s'est produite : " & Err.Description & vbCrLf & vbCrLf & " numéro d'erreur (" & Err.Number & " )", vbCritical, g_strNomApplication
  Print #g_FileDebug, Format(Now, "dd/MM/yyyy hh:mm:ss") & " Erreur n° : (" & Err.Number & ") " & Err.Description & " Procedure : vsL1_AfterEdit  Module : frmGestArme.frm"
  Resume vsL1_AfterEdit_Exit
End Sub

