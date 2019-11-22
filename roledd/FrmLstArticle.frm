VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "VSFLEX7D.ocx"
Begin VB.Form FrmLstArticle 
   Caption         =   "Liste des objets"
   ClientHeight    =   5655
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   8760
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5655
   ScaleWidth      =   8760
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox fldstrdescription 
      Height          =   1500
      Left            =   45
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Top             =   3645
      Width           =   8700
   End
   Begin VB.TextBox fldnPrixUnitaire 
      Height          =   375
      Left            =   1800
      TabIndex        =   4
      Top             =   5205
      Width           =   975
   End
   Begin VB.TextBox fldnNombre 
      Height          =   375
      Left            =   5160
      TabIndex        =   3
      Text            =   "1"
      Top             =   5205
      Width           =   495
   End
   Begin VB.CommandButton BtnAnnuler 
      Cancel          =   -1  'True
      Caption         =   "&Annuler (ESC)"
      Height          =   350
      Left            =   6210
      TabIndex        =   2
      Top             =   5220
      Width           =   1335
   End
   Begin VB.CommandButton BtnOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   350
      Left            =   7695
      TabIndex        =   1
      Top             =   5220
      Width           =   855
   End
   Begin VSFLEX7DAOCtl.vsFlexGrid FGarticle 
      Align           =   1  'Align Top
      Bindings        =   "FrmLstArticle.frx":0000
      Height          =   3570
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8760
      _ExtentX        =   15452
      _ExtentY        =   6297
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
      AllowSelection  =   -1  'True
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   3
      GridLines       =   0
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FrmLstArticle.frx":001A
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   1
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   1
      ExplorerBar     =   1
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
   Begin VB.Label LblObjet 
      AutoSize        =   -1  'True
      Caption         =   "Nombre d'objet à acheter"
      Height          =   225
      Index           =   1
      Left            =   3120
      TabIndex        =   6
      Top             =   5280
      Width           =   1785
   End
   Begin VB.Label LblObjet 
      AutoSize        =   -1  'True
      Caption         =   "Prix unitaire de l'objet"
      Height          =   225
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   5280
      Width           =   1560
   End
End
Attribute VB_Name = "FrmLstArticle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public MultPrix As Integer

Private Sub btnAnnuler_Click()

  If g_ModeDebug = vbUnchecked Then On Error GoTo btnAnnuler_Click_Error

  Unload Me


btnAnnuler_Click_Exit:
  Exit Sub

btnAnnuler_Click_Error:
  MsgBox "L'erreur suivante s'est produite : " & Err.Description & vbCrLf & vbCrLf & " numéro d'erreur (" & Err.Number & " )", vbCritical, g_strNomApplication
  Print #g_FileDebug, Format(Now, "dd/MM/yyyy hh:mm:ss") & " Erreur n° : (" & Err.Number & ") " & Err.Description & " Procedure : btnAnnuler_Click  Module : FrmLstArticle.frm"
  Resume btnAnnuler_Click_Exit
End Sub
Private Sub btnOk_Click()

  If g_ModeDebug = vbUnchecked Then On Error GoTo BtnOK_Click_Error
  
  If fldnNombre <> "" Then
    If Not IsNumeric(fldnNombre) Then
      Screen.MousePointer = vbNormal
      MsgBox "Le nombre d'objet doit être entrée en tant que valeur numérique", vbExclamation, Me.name
      fldnNombre.SetFocus
      Exit Sub
    End If
  Else
    Screen.MousePointer = vbNormal
    MsgBox "Le nombre d'objet doit être entrée en tant que valeur numérique", vbExclamation, Me.name
    fldnNombre.SetFocus
    Exit Sub
  End If
  
  If fldnPrixUnitaire <> "" Then
    If Not IsNumeric(fldnPrixUnitaire) Then
      Screen.MousePointer = vbNormal
      MsgBox "Le prix unitaire de l'objet doit être entrée en tant que valeur numérique", vbExclamation, Me.name
      fldnPrixUnitaire.SetFocus
      Exit Sub
    End If
  Else
    fldnPrixUnitaire = 0
  End If
  
  If FGarticle.Row > 0 Then
    Err.Clear
    frmGestPerso.vsEquip.Cell(flexcpText, , L3_COL_EQUIP_NOM_OBJET) = FGarticle.Cell(flexcpText, FGarticle.Row, L13_COL_EQUIP_NOM_OBJET)
    frmGestPerso.vsEquip.Cell(flexcpText, , L3_COL_EQUIP_TYPE) = FGarticle.Cell(flexcpText, FGarticle.Row, L13_COL_EQUIP_TYPE)
    frmGestPerso.vsEquip.Cell(flexcpText, , L3_COL_EQUIP_CHARGE) = FGarticle.Cell(flexcpText, FGarticle.Row, L13_COL_EQUIP_CHARGE)
    frmGestPerso.vsEquip.Cell(flexcpText, , L3_COL_EQUIP_MULTIPLE) = Val(fldnNombre)
    frmGestPerso.vsEquip.Cell(flexcpText, , L3_COL_EQUIP_SUR_PERSO) = "Oui"
    frmGestPerso.fldnTotalPo = Val(frmGestPerso.fldnTotalPo) - MultPrix * Val(fldnNombre) * Val(fldnPrixUnitaire)
    g_nValid = FGarticle.Row
    Unload Me
  End If

BtnOK_Click_Exit:
  Exit Sub

BtnOK_Click_Error:
  MsgBox "L'erreur suivante s'est produite : " & Err.Description & vbCrLf & vbCrLf & " numéro d'erreur (" & Err.Number & " )", vbCritical, g_strNomApplication
  Print #g_FileDebug, Format(Now, "dd/MM/yyyy hh:mm:ss") & " Erreur n° : (" & Err.Number & ") " & Err.Description & " Procedure : BtnOK_Click  Module : FrmLstArticle.frm"
  Resume BtnOK_Click_Exit
End Sub

Private Sub FGarticle_AfterDataRefresh()

  If g_ModeDebug = vbUnchecked Then On Error GoTo FGarticle_AfterDataRefresh_Error

  FGarticle.AutoSize 0, FGarticle.Cols - 1, , 100


FGarticle_AfterDataRefresh_Exit:
  Exit Sub

FGarticle_AfterDataRefresh_Error:
  MsgBox "L'erreur suivante s'est produite : " & Err.Description & vbCrLf & vbCrLf & " numéro d'erreur (" & Err.Number & " )", vbCritical, g_strNomApplication
  Print #g_FileDebug, Format(Now, "dd/MM/yyyy hh:mm:ss") & " Erreur n° : (" & Err.Number & ") " & Err.Description & " Procedure : FGarticle_AfterDataRefresh  Module : FrmLstArticle.frm"
  Resume FGarticle_AfterDataRefresh_Exit
End Sub
Private Sub FGarticle_DblClick()

  If g_ModeDebug = vbUnchecked Then On Error GoTo FGarticle_DblClick_Error

  Call btnOk_Click

FGarticle_DblClick_Exit:
  Exit Sub

FGarticle_DblClick_Error:
  MsgBox "L'erreur suivante s'est produite : " & Err.Description & vbCrLf & vbCrLf & " numéro d'erreur (" & Err.Number & " )", vbCritical, g_strNomApplication
  Print #g_FileDebug, Format(Now, "dd/MM/yyyy hh:mm:ss") & " Erreur n° : (" & Err.Number & ") " & Err.Description & " Procedure : FGarticle_DblClick  Module : FrmLstArticle.frm"
  Resume FGarticle_DblClick_Exit
End Sub
Private Sub FGarticle_SelChange()

fldnPrixUnitaire = FGarticle.Cell(flexcpText, FGarticle.Row, L13_COL_EQUIP_PRIX)
fldstrdescription = FGarticle.Cell(flexcpText, FGarticle.Row, L13_COL_EQUIP_DESCRIPTION)
End Sub

Private Sub Form_Load()

  If g_ModeDebug = vbUnchecked Then On Error GoTo Form_Load_Error

 
  g_nValid = False
  If MultPrix = 0 Then
    fldnPrixUnitaire.Text = "0"
    fldnPrixUnitaire.Enabled = False
  Else
    fldnPrixUnitaire.Enabled = True
  End If
  FGarticle.ColHidden(L13_COL_EQUIP_DESCRIPTION) = True
  remp

Form_Load_Exit:
  Exit Sub

Form_Load_Error:
  MsgBox "L'erreur suivante s'est produite : " & Err.Description & vbCrLf & vbCrLf & " numéro d'erreur (" & Err.Number & " )", vbCritical, g_strNomApplication
  Print #g_FileDebug, Format(Now, "dd/MM/yyyy hh:mm:ss") & " Erreur n° : (" & Err.Number & ") " & Err.Description & " Procedure : Form_Load  Module : FrmLstArticle.frm"
  Resume Form_Load_Exit
End Sub
Sub remp()
Dim rst As Recordset

  If g_ModeDebug = vbUnchecked Then On Error GoTo remp_Error

  Screen.MousePointer = vbHourglass
  FGarticle.Rows = 1
  Set rst = g_bdPerso.OpenRecordset("select * from objets order by nomObjet", dbOpenDynaset)
  Do While Not rst.EOF
    FGarticle.AddItem rst!NomObjet & vbTab & rst!GroupeObjet & vbTab & rst!CoutTotal & vbTab & rst!CHARGE & vbTab & rst!Description
    rst.MoveNext
  Loop
  rst.Close
  Screen.MousePointer = vbNormal

remp_Exit:
  Exit Sub

remp_Error:
  MsgBox "L'erreur suivante s'est produite : " & Err.Description & vbCrLf & vbCrLf & " numéro d'erreur (" & Err.Number & " )", vbCritical, g_strNomApplication
  Print #g_FileDebug, Format(Now, "dd/MM/yyyy hh:mm:ss") & " Erreur n° : (" & Err.Number & ") " & Err.Description & " Procedure : remp  Module : FrmLstArticle.frm"
  Resume remp_Exit
End Sub

Private Sub Form_Resize()
fldstrdescription.Height = maxi_(1, FrmLstArticle.Height - 4700)
fldstrdescription.Width = maxi_(1, FrmLstArticle.Width - 300)
btnOk.top = maxi_(5220, FrmLstArticle.Height - 1000)
btnAnnuler.top = maxi_(5220, FrmLstArticle.Height - 1000)
fldnNombre.top = maxi_(5220, FrmLstArticle.Height - 1000)
fldnPrixUnitaire.top = maxi_(5220, FrmLstArticle.Height - 1000)
LblObjet(0).top = maxi_(5220, FrmLstArticle.Height - 1000)
LblObjet(1).top = maxi_(5220, FrmLstArticle.Height - 1000)

End Sub
