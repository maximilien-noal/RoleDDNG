VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "VSFLEX7D.ocx"
Begin VB.Form frmExperience 
   Caption         =   "Expérience des personnages"
   ClientHeight    =   7485
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11580
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7485
   ScaleWidth      =   11580
   Begin VB.TextBox fldnIdee 
      Height          =   375
      Left            =   9000
      TabIndex        =   7
      Text            =   "0"
      ToolTipText     =   "Saisir un pourcentage pour les XP idée."
      Top             =   3720
      Width           =   615
   End
   Begin VB.CommandButton btnAdd 
      Caption         =   "Ajouter"
      Height          =   375
      Left            =   9720
      TabIndex        =   5
      Top             =   2760
      Width           =   855
   End
   Begin VB.TextBox fldnFP 
      Height          =   375
      Left            =   9000
      TabIndex        =   4
      Top             =   2760
      Width           =   615
   End
   Begin VB.CommandButton btnFermer 
      Caption         =   "Fermer"
      Height          =   375
      Left            =   10080
      TabIndex        =   3
      Top             =   6960
      Width           =   1335
   End
   Begin VB.TextBox fldstrExperience 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2595
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   4800
      Width           =   8565
   End
   Begin VB.CommandButton btnEnregistrer 
      Caption         =   "Enregistrer"
      Height          =   375
      Left            =   8640
      TabIndex        =   1
      Top             =   6960
      Width           =   1335
   End
   Begin VSFLEX7DAOCtl.vsFlexGrid vspersonnage 
      Height          =   4740
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8580
      _ExtentX        =   15134
      _ExtentY        =   8361
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
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
      AllowUserResizing=   0
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
      FormatString    =   "Nom                                |Niv|NGE|Expérience|Niveau suivant|Gain d'XP      |Part"
      ScrollTrack     =   0   'False
      ScrollBars      =   2
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
      ExplorerBar     =   1
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
   Begin VB.Label fldstrTexteIdee 
      Caption         =   "Total XP idée"
      Height          =   255
      Index           =   1
      Left            =   9960
      TabIndex        =   10
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Label fldstrTexteIdee 
      Caption         =   "% d'XP idée"
      Height          =   255
      Index           =   0
      Left            =   8760
      TabIndex        =   9
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Label fldnTotalIdee 
      Height          =   375
      Left            =   10080
      TabIndex        =   8
      Top             =   3720
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   $"frmExperience.frx":0000
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   8760
      TabIndex        =   6
      Top             =   240
      Width           =   2655
   End
End
Attribute VB_Name = "frmExperience"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Sub RempPerso()
Dim rst As Recordset, rstRace As Recordset, rstArch As Recordset, rstDon As Recordset
Dim Niv_1 As Integer, Niv_2 As Integer, Niv_3 As Integer, Niv_4 As Integer, Niv_5 As Integer, Niv_6 As Integer, Niv_7 As Integer, Niv_8 As Integer, Niveau As Integer
Dim NiveauGE As Integer, MalusXP As Integer, AdjRace As Integer, AdjArch As Integer, i As Integer
Dim XP As Long

  If g_ModeDebug = vbUnchecked Then On Error GoTo RempPerso_Error

  vspersonnage.Rows = 1
  Set rst = g_bdPerso.OpenRecordset("select nom, race,niv_1,niv_2,niv_3,niv_4,niv_5,niv_6,niv_7,niv_8,malusxp,archetype,totalxp from personnage where exclu=false order by nom", dbOpenDynaset)
  
  
  Set rstRace = g_bdRoles.OpenRecordset("race", dbOpenTable)
  Set rstArch = g_bdRoles.OpenRecordset("archetype", dbOpenTable)
  rstRace.Index = "PrimaryKey"
  rstArch.Index = "PrimaryKey"
  i = 1
  Do While Not rst.EOF
    rstRace.Seek "=", rst!RACE
    rstArch.Seek "=", rst!Archetype
    Set rstDon = g_bdPerso.OpenRecordset("select dons from PersonnageDons where nom='" & FormatStrSQL(rst!Nom) & "'", dbOpenDynaset)
    Niv_1 = IIf(IsNull(rst!Niv_1), 0, rst!Niv_1)
    Niv_2 = IIf(IsNull(rst!Niv_2), 0, rst!Niv_2)
    Niv_3 = IIf(IsNull(rst!Niv_3), 0, rst!Niv_3)
    Niv_4 = IIf(IsNull(rst!Niv_4), 0, rst!Niv_4)
    Niv_5 = IIf(IsNull(rst!Niv_5), 0, rst!Niv_5)
    Niv_6 = IIf(IsNull(rst!Niv_6), 0, rst!Niv_6)
    Niv_7 = IIf(IsNull(rst!Niv_7), 0, rst!Niv_7)
    Niv_7 = IIf(IsNull(rst!Niv_8), 0, rst!Niv_8)
    MalusXP = IIf(IsNull(rst!MalusXP), 0, rst!MalusXP)
    
    If rstRace.NoMatch Then
      AdjRace = 0
    Else
      AdjRace = IIf(IsNull(rstRace!AdjNiv), 0, rstRace!AdjNiv)
    End If
  
    If rstArch.NoMatch Then
      AdjArch = 0
    Else
      AdjArch = IIf(IsNull(rstArch!AdjNiv), 0, rstArch!AdjNiv)
    End If
    Niveau = Niv_1 + Niv_2 + Niv_3 + Niv_4 + Niv_5 + Niv_6 + Niv_7 + Niv_8
    NiveauGE = Niveau + AdjRace + AdjArch
    Do While Not rstDon.EOF
      If rstDon!dons = REDUCTION_NIVEAU_1 Then
        NiveauGE = NiveauGE - 1
      End If
      rstDon.MoveNext
    Loop
    If rst!TotalXP > NiveauXP(Val(NiveauGE)) Then
      XP = rst!TotalXP
    Else
      XP = NiveauXP(Val(NiveauGE))
    End If
    vspersonnage.AddItem rst!Nom & vbTab & Niveau & vbTab & NiveauGE & vbTab & XP & vbTab & NiveauXP(NiveauGE + 1) & vbTab & vbTab & 1
    If vspersonnage.Cell(flexcpValue, i, L12_COL_EXPERIENCE) >= vspersonnage.Cell(flexcpValue, i, L12_COL_NIVEAU_SUIVANT) Then
      fldstrExperience = fldstrExperience & vspersonnage.Cell(flexcpText, i, L12_COL_NOM) & " devrait être au niveau " & vspersonnage.Cell(flexcpText, i, L12_COL_NIVEAU) + 1 & ", les calculs d'XP avec le FP des monstres seront faux. " & vbCrLf
    End If
    rst.MoveNext
    rstDon.Close
    i = i + 1
  Loop
  rst.Close
  rstRace.Close
  rstArch.Close
  
  vspersonnage.Row = 1
  vspersonnage.Col = 5
  Screen.MousePointer = vbNormal

RempPerso_Exit:
  Exit Sub

RempPerso_Error:
  MsgBox "L'erreur suivante s'est produite : " & Err.Description & vbCrLf & vbCrLf & " numéro d'erreur (" & Err.Number & " )", vbCritical, g_strNomApplication
  Print #g_FileDebug, Format(Now, "dd/MM/yyyy hh:mm:ss") & " Erreur n° : (" & Err.Number & ") " & Err.Description & " Procedure : RempPerso  Module : frmAcceder.frm"
  Resume RempPerso_Exit
End Sub

Private Sub btnAdd_Click()
Dim i As Integer, j As Integer
Dim NbPart As Double
Dim bLigneSelect As Boolean
Dim rstExperience As Recordset

If Val(fldnIdee) < 0 Or Val(fldnIdee) > 100 Then
  MsgBox "Le % des XP d'idée doit être un réél supérieur à 0 et inférieur à 100"
    fldnIdee.SetFocus
Else
  If Val(fldnFP) > 0 Then
    If vspersonnage.SelectedRows > 0 Then
      NbPart = 0
      For i = 1 To vspersonnage.Rows - 1
        For j = 0 To vspersonnage.Rows - 1
          If vspersonnage.SelectedRow(j) = i Then
            NbPart = NbPart + Val(vspersonnage.Cell(flexcpText, i, L12_COL_PART))
          End If
        Next j
      Next i
      For i = 1 To vspersonnage.Rows - 1
        bLigneSelect = False
        For j = 0 To vspersonnage.Rows - 1
          If vspersonnage.SelectedRow(j) = i Then
            bLigneSelect = True
          End If
        Next j
        If bLigneSelect Then
          fldnFP = mini_(Val(fldnFP), 125)
          Set rstExperience = g_bdRoles.OpenRecordset("select fp" & fldnFP & " from experience where niveau=" & maxi_(Val(vspersonnage.Cell(flexcpText, i, L12_COL_NGE)), 3), dbOpenDynaset)
  
          vspersonnage.Cell(flexcpText, i, L12_COL_GAIN_XP) = Int(Int(Val(vspersonnage.Cell(flexcpText, i, L12_COL_PART)) * rstExperience.Fields("fp" & fldnFP) / NbPart) * (1 - Val(fldnIdee) / 100))
          rstExperience.Close
        End If
      Next i
    Else
      MsgBox "Il faut au moins un personnage de sélectionné"
    End If
  Else
    MsgBox "Le FP doit être un réél strictement supérieur à 0"
    fldnFP.SetFocus
  End If
End If
End Sub

Private Sub btnEnregistrer_Click()
Dim rst As Recordset
Dim i As Integer

Set rst = g_bdPerso.OpenRecordset("personnage", dbOpenTable)
rst.Index = "Primarykey"

For i = 1 To vspersonnage.Rows - 1
  rst.Seek "=", vspersonnage.Cell(flexcpText, i, L12_COL_NOM)
  rst.Edit
  If vspersonnage.Cell(flexcpText, i, L12_COL_GAIN_XP) <> "" Then
    vspersonnage.Cell(flexcpText, i, L12_COL_EXPERIENCE) = CStr(vspersonnage.Cell(flexcpValue, i, L12_COL_EXPERIENCE) + Val(vspersonnage.Cell(flexcpText, i, L12_COL_GAIN_XP)))
    fldnTotalIdee = Val(fldnTotalIdee) + Int(Val(vspersonnage.Cell(flexcpText, i, L12_COL_GAIN_XP)) * Val(fldnIdee) / (100 - Val(fldnIdee)))
    fldstrExperience = fldstrExperience & vspersonnage.Cell(flexcpText, i, L12_COL_NOM) & " : " & Val(vspersonnage.Cell(flexcpText, i, L12_COL_GAIN_XP)) & " XP" & vbTab
    vspersonnage.Cell(flexcpText, i, L12_COL_GAIN_XP) = ""
  End If
  rst!TotalXP = vspersonnage.Cell(flexcpValue, i, L12_COL_EXPERIENCE)
  If vspersonnage.Cell(flexcpValue, i, L12_COL_EXPERIENCE) >= vspersonnage.Cell(flexcpValue, i, L12_COL_NIVEAU_SUIVANT) Then
    fldstrExperience = fldstrExperience & "Passage de " & vspersonnage.Cell(flexcpText, i, L12_COL_NOM) & " au niveau " & vspersonnage.Cell(flexcpText, i, 1) + 1 & vbTab
  End If
  rst.Update
Next i
rst.Close
End Sub

Private Sub btnFermer_Click()
  Unload Me
End Sub

Private Sub Form_Load()

  If g_ModeDebug = vbUnchecked Then On Error GoTo Form_Load_Error
  RempPerso

Form_Load_Exit:
  Exit Sub

Form_Load_Error:
  MsgBox "L'erreur suivante s'est produite : " & Err.Description & vbCrLf & vbCrLf & " numéro d'erreur (" & Err.Number & " )", vbCritical, g_strNomApplication
  Print #g_FileDebug, Format(Now, "dd/MM/yyyy hh:mm:ss") & " Erreur n° : (" & Err.Number & ") " & Err.Description & " Procedure : Form_Load  Module : frmexperience.frm"
  Resume Form_Load_Exit
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "shift " & Shift & vbTab & "x " & X & vbTab & "y " & Y
End Sub

Private Sub Form_Resize()
Dim b As Variant

If frmExperience.WindowState <> vbMaximized Then
  frmExperience.Width = 11700
  frmExperience.Height = mini_(frmExperience.Height, frmMain.Height)
End If

vspersonnage.Height = (frmExperience.Height - 800) * 0.6

fldstrExperience.Height = (frmExperience.Height - 800) * 0.4
fldstrExperience.top = vspersonnage.Height + 200
btnEnregistrer.top = frmExperience.Height - 1000
BtnFermer.top = btnEnregistrer.top
End Sub

Private Sub vspersonnage_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)

  If g_ModeDebug = vbUnchecked Then On Error GoTo vspersonnage_BeforeEdit_Error
  
  If Row > 0 And (Col = L12_COL_GAIN_XP Or Col = L12_COL_EXPERIENCE Or Col = L12_COL_PART) Then
    vspersonnage.ComboList = ""
  Else
    If Col <> 0 Then
      MsgBox "Vous ne pouvez pas changer cette céllule"
    End If
  End If

vspersonnage_BeforeEdit_Exit:
  Exit Sub

vspersonnage_BeforeEdit_Error:
  MsgBox "L'erreur suivante s'est produite : " & Err.Description & vbCrLf & vbCrLf & " numéro d'erreur (" & Err.Number & " )", vbCritical, g_strNomApplication
  Print #g_FileDebug, Format(Now, "dd/MM/yyyy hh:mm:ss") & " Erreur n° : (" & Err.Number & ") " & Err.Description & " Procedure : vspersonnage_BeforeEdit  Module : frmexperience.frm"
  Resume vspersonnage_BeforeEdit_Exit
End Sub


Private Sub vspersonnage_Click()
Dim i As Integer

  i = vspersonnage.Row

End Sub

Private Sub vspersonnage_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)

  If g_ModeDebug = vbUnchecked Then On Error GoTo vspersonnage_ValidateEdit_Error

vspersonnage_ValidateEdit_Exit:
  Exit Sub

vspersonnage_ValidateEdit_Error:
  MsgBox "L'erreur suivante s'est produite : " & Err.Description & vbCrLf & vbCrLf & " numéro d'erreur (" & Err.Number & " )", vbCritical, g_strNomApplication
  Print #g_FileDebug, Format(Now, "dd/MM/yyyy hh:mm:ss") & " Erreur n° : (" & Err.Number & ") " & Err.Description & " Procedure : vsHistoPerso_ValidateEdit  Module : frmGestPerso.frm"
  Resume vspersonnage_ValidateEdit_Exit
End Sub
