VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "VSFLEX7D.ocx"
Begin VB.Form frmImport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Importation de personnage d'une autre base de donnée"
   ClientHeight    =   7305
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10140
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7305
   ScaleWidth      =   10140
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton BtnFermer 
      Caption         =   "Fermer"
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
      Left            =   8520
      TabIndex        =   5
      Top             =   6840
      Width           =   1335
   End
   Begin VB.CommandButton BtnImport 
      Caption         =   "Importer"
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
      Left            =   7080
      TabIndex        =   4
      Top             =   6840
      Width           =   1335
   End
   Begin VB.CheckBox chkObjet 
      Caption         =   "Avec les objets"
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
      Left            =   120
      TabIndex        =   3
      Top             =   6840
      Width           =   1935
   End
   Begin VB.CommandButton btnTous 
      Caption         =   "Tous les personnages"
      Height          =   375
      Left            =   2280
      TabIndex        =   2
      Top             =   6840
      Width           =   1695
   End
   Begin VB.CommandButton btnAucun 
      Caption         =   "Aucun personnage"
      Height          =   375
      Left            =   4080
      TabIndex        =   1
      Top             =   6840
      Width           =   1695
   End
   Begin VSFLEX7DAOCtl.vsFlexGrid vsPersoImport 
      Height          =   6705
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10140
      _ExtentX        =   17886
      _ExtentY        =   11827
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
      FormatString    =   $"frmImport.frx":0000
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   -1  'True
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
      TabBehavior     =   1
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
Attribute VB_Name = "frmImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public m_bBase As Boolean ' si true alors on travaille sur la même base
Public BDI As Database

Private Sub btnAucun_Click()
  AffecteImport ("Non")
End Sub

Private Sub btnFermer_Click()
  Unload Me
End Sub
Private Sub BtnImport_Click()
  Dim rstImport As Recordset, rstperso As Recordset, rstSousImport As Recordset, rstSousperso As Recordset
  Dim indic As Boolean
  Dim i As Integer, j As Integer
    
  'les objets
  If chkObjet.Value And (Not m_bBase) Then
    Set rstImport = BDI.OpenRecordset("Objets", dbOpenTable)
    Set rstperso = g_bdPerso.OpenRecordset("Objets", dbOpenTable)
    With rstperso
      .Index = "PrimaryKey"
      Do While Not rstImport.EOF
        .Seek "=", rstImport!NomObjet
        If .NoMatch Then
          '' Objet inéxistant création de l'objet
          .AddNew
          !NomObjet = rstImport!NomObjet
        Else
          .Edit
          g_bdPerso.Execute "delete from objetsPropriete where [nomobjet]='" & FormatStrSQL(rstImport!NomObjet) & "'", dbFailOnError
          ''trouver l'objet modification des valeurs avec les nouvelles
        End If
        !PoidsBase = rstImport!PoidsBase
        !classeArmure = rstImport!classeArmure
        !typeca = rstImport!typeca
        !EchecSortProfane = rstImport!EchecSortProfane
        !PenaliteArmure = rstImport!PenaliteArmure
        !BonusDexMax = rstImport!BonusDexMax
        !GroupeObjet = rstImport!GroupeObjet
        !taille = rstImport!taille
        !Type = rstImport!Type
        !DegatDes = rstImport!DegatDes
        !ZoneCritique = rstImport!ZoneCritique
        !MultCrit = rstImport!MultCrit
        !Portee = rstImport!Portee
        !CoutTotal = rstImport!CoutTotal
        !Description = rstImport!Description
        !Solidite = rstImport!Solidite
        !RESISTANCE = rstImport!RESISTANCE
        !CHARGE = rstImport!CHARGE
        .Update
        rstImport.MoveNext
      Loop
    End With
    rstImport.Close
    rstperso.Close
    
    Set rstImport = BDI.OpenRecordset("Objetspropriete", dbOpenTable)
    Set rstperso = g_bdPerso.OpenRecordset("Select * from ObjetsPropriete", dbOpenDynaset)
    Do While Not rstImport.EOF
      With rstperso
        .AddNew
        !NomObjet = rstImport!NomObjet
        !propriete_1 = rstImport!propriete_1
        !propriete_2 = rstImport!propriete_2
        !propriete_3 = rstImport!propriete_3
        !valeur = rstImport!valeur
        .Update
      End With
      rstImport.MoveNext
    Loop
    rstperso.Close
    rstImport.Close
  End If
  
  'Les personnages
  Set rstImport = BDI.OpenRecordset("personnage", dbOpenTable)
  Set rstperso = g_bdPerso.OpenRecordset("personnage", dbOpenTable)
  
  For i = 1 To vsPersoImport.Rows - 1
    If vsPersoImport.Cell(flexcpText, i, L11_COL_IMPORT) = "Oui" Then
      rstperso.Index = "PrimaryKey"
      rstperso.Seek "=", vsPersoImport.Cell(flexcpText, i, L11_COL_NOM_IMPORT)
      rstImport.Index = "PrimaryKey"
      rstImport.Seek "=", vsPersoImport.Cell(flexcpText, i, L11_COL_NOM)
      If rstperso.NoMatch Then
        '' Le personnage n'existe pas
        rstperso.AddNew
        
        rstperso.Fields(0) = vsPersoImport.Cell(flexcpText, i, L11_COL_NOM_IMPORT)
        For j = 1 To 74
          If Not IsNull(rstImport.Fields(j)) Then
            rstperso.Fields(j) = rstImport.Fields(j)
          End If
        Next j
        rstperso.Update
        '' L'equipement du personnage
        If chkObjet.Value Then
          Set rstSousImport = BDI.OpenRecordset("select * from equipement where personnage='" & FormatStrSQL(vsPersoImport.Cell(flexcpText, i, L11_COL_NOM)) & "'", dbOpenDynaset)
          Set rstSousperso = g_bdPerso.OpenRecordset("Equipement", dbOpenTable)
          Do While Not rstSousImport.EOF
            rstSousperso.AddNew 'j'ajoute l'objet
            For j = 0 To 7
              If Not IsNull(rstSousImport.Fields(j)) Then
                If j = 1 Then
                  rstSousperso.Fields(j) = vsPersoImport.Cell(flexcpText, i, L11_COL_NOM_IMPORT)
                Else
                  rstSousperso.Fields(j) = rstSousImport.Fields(j)
                End If
              End If
            Next j
            rstSousperso.Update
            rstSousImport.MoveNext
          Loop
          rstSousperso.Close
          rstSousImport.Close
        End If
        '' Les dons
        Set rstSousImport = BDI.OpenRecordset("select * from PersonnageDons where nom='" & FormatStrSQL(vsPersoImport.Cell(flexcpText, i, L11_COL_NOM)) & "'", dbOpenDynaset)
        Set rstSousperso = g_bdPerso.OpenRecordset("PersonnageDons", dbOpenTable)
        Do While Not rstSousImport.EOF
          rstSousperso.AddNew 'j'ajoute le don
          
          rstSousperso.Fields(0) = vsPersoImport.Cell(flexcpText, i, L11_COL_NOM_IMPORT)
          For j = 1 To 5
            If Not IsNull(rstSousImport.Fields(j)) Then
              rstSousperso.Fields(j) = rstSousImport.Fields(j)
            End If
          Next j
          rstSousperso.Update
          rstSousImport.MoveNext
        Loop
        rstSousperso.Close
        rstSousImport.Close
        '' La progression
        Set rstSousImport = BDI.OpenRecordset("select * from PersonnageProgression where nom='" & FormatStrSQL(vsPersoImport.Cell(flexcpText, i, L11_COL_NOM)) & "'", dbOpenDynaset)
        Set rstSousperso = g_bdPerso.OpenRecordset("PersonnageProgression", dbOpenTable)
        Do While Not rstSousImport.EOF
          rstSousperso.AddNew 'j'ajoute le niveau
          rstSousperso.Fields(0) = vsPersoImport.Cell(flexcpText, i, L11_COL_NOM_IMPORT)
          For j = 1 To 12
            If Not IsNull(rstSousImport.Fields(j)) Then
              rstSousperso.Fields(j) = rstSousImport.Fields(j)
            End If
          Next j
          rstSousperso.Update
          rstSousImport.MoveNext
        Loop
        rstSousperso.Close
        rstSousImport.Close
        
        '' Les sorts
        Set rstSousImport = BDI.OpenRecordset("select * from SortPersonnage where NomPerso='" & FormatStrSQL(vsPersoImport.Cell(flexcpText, i, L11_COL_NOM)) & "'", dbOpenDynaset)
        Set rstSousperso = g_bdPerso.OpenRecordset("SortPersonnage", dbOpenTable)
        Do While Not rstSousImport.EOF
          rstSousperso.AddNew 'j'ajoute le sort
          rstSousperso.Fields(0) = vsPersoImport.Cell(flexcpText, i, L11_COL_NOM_IMPORT)
          For j = 1 To 3
            If Not IsNull(rstSousImport.Fields(j)) Then
              rstSousperso.Fields(j) = rstSousImport.Fields(j)
            End If
          Next j
          rstSousperso.Update
          rstSousImport.MoveNext
        Loop
        rstSousperso.Close
        rstSousImport.Close
        '' La liste de sort
        Set rstSousImport = BDI.OpenRecordset("select * from SortListe where NomPerso='" & FormatStrSQL(vsPersoImport.Cell(flexcpText, i, L11_COL_NOM)) & "'", dbOpenDynaset)
        Set rstSousperso = g_bdPerso.OpenRecordset("SortListe", dbOpenTable)
        Do While Not rstSousImport.EOF
          rstSousperso.AddNew 'j'ajoute le sort
          rstSousperso.Fields(0) = vsPersoImport.Cell(flexcpText, i, L11_COL_NOM_IMPORT)
          For j = 1 To 5
            If Not IsNull(rstSousImport.Fields(j)) Then
              rstSousperso.Fields(j) = rstSousImport.Fields(j)
            End If
          Next j
          rstSousperso.Update
          rstSousImport.MoveNext
        Loop
        rstSousperso.Close
        rstSousImport.Close
      Else
        MsgBox vsPersoImport.Cell(flexcpText, i, L11_COL_NOM_IMPORT) & " existe déjà dans cette base, vous devez le supprimer ou le renommer avant de l'importer."
      End If
    End If
  Next i
  rstperso.Close
  rstImport.Close
  Unload Me
  If frmMain.TrouveChild("frmAcceder") = -1 Then
    frmAcceder.Show
  Else
    frmAcceder.AccedePerso
  End If
  
End Sub

Private Sub btnTous_Click()
  AffecteImport ("Oui")
End Sub
Public Sub AffecteImport(Etat As String)
  
Dim i As Integer

  For i = 0 To vsPersoImport.Rows - 1
    vsPersoImport.Cell(flexcpText, i, L11_COL_IMPORT) = Etat
  Next i

End Sub

Private Sub chkObjet_Click()
  If chkObjet.Value Then
    chkObjet.Caption = "Avec les objets"
  Else
    chkObjet.Caption = "Sans les objets"
  End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()

  Dim rst As Recordset
  Dim ClasseMax As String
  Dim NivMax As Integer
  Dim Niv_1 As Integer, Niv_2 As Integer, Niv_3 As Integer, Niv_4 As Integer, Niv_5 As Integer, Niv_6 As Integer, Niv_ As Integer, Niv_8 As Integer

  If g_ModeDebug = vbUnchecked Then On Error Resume Next
  Unload frmAcceder
  ControleVersion BDI
  vsPersoImport.Rows = 1
  vsPersoImport.Col = L11_COL_NOM_IMPORT
  
  Set rst = BDI.OpenRecordset("select nom,race,niv_1,niv_2,niv_3,niv_4,niv_5,niv_6,niv_7,niv_8,classe_1,classe_2,classe_3,classe_4,classe_5,classe_6,classe_7,classe_8 from personnage where exclu=false order by nom", dbOpenDynaset)
  Do While Not rst.EOF
    Niv_1 = IIf(IsNull(rst!Niv_1), 0, rst!Niv_1)
    Niv_2 = IIf(IsNull(rst!Niv_2), 0, rst!Niv_2)
    Niv_3 = IIf(IsNull(rst!Niv_3), 0, rst!Niv_3)
    Niv_4 = IIf(IsNull(rst!Niv_4), 0, rst!Niv_4)
    Niv_5 = IIf(IsNull(rst!Niv_5), 0, rst!Niv_5)
    Niv_6 = IIf(IsNull(rst!Niv_6), 0, rst!Niv_6)
    Niv_7 = IIf(IsNull(rst!Niv_7), 0, rst!Niv_7)
    Niv_8 = IIf(IsNull(rst!Niv_8), 0, rst!Niv_8)
    
    ClasseMax = rst!Classe_1
    NivMax = rst!Niv_1
    
    If Niv_2 > NivMax Then
      ClasseMax = rst!Classe_2
      NivMax = Niv_2
    End If
    If Niv_3 > NivMax Then
      ClasseMax = rst!Classe_3
      NivMax = Niv_3
    End If
    If Niv_4 > NivMax Then
      ClasseMax = rst!Classe_4
      NivMax = Niv_4
    End If
    If Niv_5 > NivMax Then
      ClasseMax = rst!Classe_5
      NivMax = Niv_5
    End If
    If Niv_6 > NivMax Then
      ClasseMax = rst!Classe_6
      NivMax = Niv_6
    End If
    If Niv_7 > NivMax Then
      ClasseMax = rst!Classe_7
      NivMax = Niv_7
    End If
    If Niv_8 > NivMax Then
      ClasseMax = rst!Classe_8
      NivMax = Niv_8
    End If
    
    vsPersoImport.AddItem rst!Nom & vbTab & rst!Nom & vbTab & rst!RACE & vbTab & Niv_1 + Niv_2 + Niv_3 + Niv_4 + Niv_5 + Niv_6 & vbTab & ClasseMax & vbTab & "Non"
    rst.MoveNext
  Loop
  rst.Close
  vsPersoImport.ColComboList(L11_COL_IMPORT) = "Oui|Non"
  chkObjet.Value = 1
End Sub
Private Sub vspersoimport_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)

  If g_ModeDebug = vbUnchecked Then On Error GoTo vspersoimport_BeforeEdit_Error
  
  If Row > 0 Then
    If Col = L11_COL_NOM_IMPORT Or Col = L11_COL_IMPORT Then
      vsPersoImport.ComboList = ""
    Else
      MsgBox "Vous ne pouvez pas changer cette céllule"
    End If
  End If
vspersoimport_BeforeEdit_Exit:
  Exit Sub

vspersoimport_BeforeEdit_Error:
  MsgBox "L'erreur suivante s'est produite : " & Err.Description & vbCrLf & vbCrLf & " numéro d'erreur (" & Err.Number & " )", vbCritical, g_strNomApplication
  Print #g_FileDebug, Format(Now, "dd/MM/yyyy hh:mm:ss") & " Erreur n° : (" & Err.Number & ") " & Err.Description & " Procedure : vspersoimport_BeforeEdit  Module : frmimport.frm"
  Resume vspersoimport_BeforeEdit_Exit
End Sub
Private Sub Form_Unload(Cancel As Integer)
  BDI.Close
End Sub
