VERSION 5.00
Begin VB.Form frmInsertionDons 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Insertion de dons"
   ClientHeight    =   7710
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9180
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
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7710
   ScaleWidth      =   9180
   Begin VB.ComboBox cmbCaracteristique 
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
      Left            =   5355
      TabIndex        =   27
      Text            =   "Caractéristique"
      Top             =   5760
      Width           =   2055
   End
   Begin VB.ComboBox cmbGenre 
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
      Left            =   4590
      TabIndex        =   26
      Text            =   "Genre"
      Top             =   90
      Width           =   2775
   End
   Begin VB.ComboBox cmbSource 
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
      Left            =   1260
      TabIndex        =   0
      Text            =   "Livre"
      Top             =   465
      Width           =   2415
   End
   Begin VB.ComboBox cmbNom 
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
      Left            =   1260
      TabIndex        =   1
      Text            =   "Nom du don"
      Top             =   90
      Width           =   2415
   End
   Begin VB.CheckBox ChkPrive 
      Alignment       =   1  'Right Justify
      Caption         =   "Privé"
      Height          =   375
      Left            =   135
      TabIndex        =   25
      ToolTipText     =   "Un don accessible seulement à une classe ou une race ne peut être choisi en tant que don (ex : châtiment du mal)."
      Top             =   6285
      Width           =   2295
   End
   Begin VB.CheckBox ChkMultiple 
      Alignment       =   1  'Right Justify
      Caption         =   "Multiple"
      Height          =   375
      Left            =   3825
      TabIndex        =   24
      ToolTipText     =   "Un don que l'on peut prendre plusieurs fois (ex : robustesse)"
      Top             =   6285
      Width           =   2295
   End
   Begin VB.CheckBox ChkEpique 
      Alignment       =   1  'Right Justify
      Caption         =   "Epique"
      Height          =   375
      Left            =   6720
      TabIndex        =   23
      ToolTipText     =   "Un don que l'on peut seulement choisir à partir du niveau 21."
      Top             =   6285
      Width           =   2295
   End
   Begin VB.TextBox fldnPage 
      Height          =   285
      Left            =   4680
      TabIndex        =   22
      Top             =   495
      Width           =   615
   End
   Begin VB.TextBox fldnVersion 
      Height          =   285
      Left            =   8400
      TabIndex        =   21
      Text            =   "3.5"
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton btnNouveau 
      Caption         =   "&Nouveau"
      Height          =   375
      Left            =   5280
      TabIndex        =   20
      Top             =   7320
      Width           =   1215
   End
   Begin VB.CommandButton btnEnregistrer 
      Caption         =   "&Enregistrer"
      Height          =   375
      Left            =   6600
      TabIndex        =   19
      Top             =   7320
      Width           =   1215
   End
   Begin VB.CommandButton btnFermer 
      Caption         =   "&Fermer"
      Height          =   375
      Left            =   7920
      TabIndex        =   18
      Top             =   7320
      Width           =   1215
   End
   Begin VB.CommandButton btnSuivant 
      Caption         =   "&Suivant"
      Height          =   375
      Left            =   3960
      TabIndex        =   17
      Top             =   7320
      Width           =   1215
   End
   Begin VB.CommandButton btnPrecedent 
      Caption         =   "&Précédent"
      Height          =   375
      Left            =   2640
      TabIndex        =   16
      Top             =   7320
      Width           =   1215
   End
   Begin VB.TextBox fldstrResume 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   15
      Top             =   1200
      Width           =   8895
   End
   Begin VB.TextBox fldstrExplication 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   14
      Top             =   2640
      Width           =   8895
   End
   Begin VB.CheckBox ChkTare 
      Alignment       =   1  'Right Justify
      Caption         =   "Handicaps"
      Height          =   375
      Left            =   6720
      TabIndex        =   13
      ToolTipText     =   "Les andicaps sont le revers des dons (ex : Affligeant)"
      Top             =   6840
      Width           =   2295
   End
   Begin VB.CheckBox ChkTrait 
      Alignment       =   1  'Right Justify
      Caption         =   "Trait"
      Height          =   375
      Left            =   3800
      TabIndex        =   12
      ToolTipText     =   "Un trait représente un aspect de la personnalité du PJ (ex : Abrupt)"
      Top             =   6840
      Width           =   2295
   End
   Begin VB.CheckBox ChkLiearme 
      Alignment       =   1  'Right Justify
      Caption         =   "Lié a une arme"
      Height          =   375
      Left            =   120
      TabIndex        =   11
      ToolTipText     =   "Un don qui donne un éffet sur un type d'arme (ex : arme de prédilection)"
      Top             =   6840
      Width           =   2295
   End
   Begin VB.CheckBox ChkPouvoir 
      Alignment       =   1  'Right Justify
      Caption         =   "Pouvoir"
      Height          =   375
      Left            =   120
      TabIndex        =   10
      ToolTipText     =   "Un pouvoir qui se calcule sous la forme 10 + 1/2 dés de vie + bonus caratéristique"
      Top             =   5805
      Width           =   2295
   End
   Begin VB.Label LabelDon 
      Caption         =   "Caractéristique"
      Height          =   255
      Index           =   7
      Left            =   3780
      TabIndex        =   9
      ToolTipText     =   "La caractéristique associée au pouvoir"
      Top             =   5805
      Width           =   1455
   End
   Begin VB.Label LabelDon 
      Caption         =   "Explication"
      Height          =   255
      Index           =   6
      Left            =   3870
      TabIndex        =   8
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label LabelDon 
      Caption         =   "Résumé"
      Height          =   255
      Index           =   5
      Left            =   3840
      TabIndex        =   7
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label LabelDon 
      Caption         =   "Page"
      Height          =   255
      Index           =   4
      Left            =   3795
      TabIndex        =   6
      Top             =   510
      Width           =   1455
   End
   Begin VB.Label LabelDon 
      Caption         =   "Source"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   5
      Top             =   510
      Width           =   1455
   End
   Begin VB.Label LabelDon 
      Caption         =   "Genre"
      Height          =   255
      Index           =   2
      Left            =   3795
      TabIndex        =   4
      Top             =   135
      Width           =   1455
   End
   Begin VB.Label LabelDon 
      Caption         =   "Version"
      Height          =   255
      Index           =   1
      Left            =   7560
      TabIndex        =   3
      ToolTipText     =   "La version du livre dans lequel est tiré le don souvent 3.5 ou 3.0."
      Top             =   135
      Width           =   1455
   End
   Begin VB.Label LabelDon 
      Caption         =   "Nom du don"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   135
      Width           =   1455
   End
End
Attribute VB_Name = "frmInsertionDons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnEnregistrer_Click()
Dim rstRoles As Recordset, rstperso As Recordset, rstTest As Recordset
Dim bexist As Boolean
Dim NomDon As String

If IsNull(cmbNom) Or cmbNom = "" Then
 MsgBox "Il faut que le don ait un nom"
 cmbNom.SetFocus
 Exit Sub
End If
If IsNull(cmbGenre) Or cmbGenre = "" Then
 MsgBox "Il faut que le don ait un genre"
 cmbGenre.SetFocus
 Exit Sub
End If
If IsNull(cmbSource) Or cmbSource = "" Then
  MsgBox "Il faut que le don ait une source"
  cmbSource.SetFocus
  Exit Sub
End If
Set rstRoles = g_bdRoles.OpenRecordset("dons", dbOpenTable)
rstRoles.Index = "nom"
rstRoles.Seek "=", cmbNom
If Not rstRoles.NoMatch Then
  MsgBox "Attention ce don existe déja dans la base de donnée du programmme, il sera modifié si vous ne changer pas le nom."
End If
Set rstperso = g_bdPerso.OpenRecordset("donsperso", dbOpenTable)
With rstperso
NomDon = cmbNom
bexist = False
  Do While Not .EOF
    If !nom = cmbNom Then
      bexist = True
      Exit Do
    End If
    .MoveNext
  Loop
  If bexist Then
    .Edit
  Else
    .AddNew
    !nom = cmbNom
  End If
  !version = IIf(IsNull(fldnVersion), 0, Val(fldnVersion))
  
  !Genre = cmbGenre
  !Source = cmbSource & IIf(IsNull(fldnPage) Or fldnPage = "", "", " p" & fldnPage)
  If fldstrResume <> "" Then
    !résumé = fldstrResume
  End If
  If fldstrExplication <> "" Then
    !Explication = fldstrExplication
  End If
  If cmbCaracteristique <> "" Then
    !Caracteristique = cmbCaracteristique
  End If
  !prive = ChkPrive
  !liearme = ChkLiearme
  !Multiple = ChkMultiple
  !epique = ChkEpique
  !pouvoir = ChkPouvoir
  !tare = ChkTare
  !trait = ChkTrait
  .Update
End With

rstperso.Close
rstRoles.Close
With cmbNom
  .Clear
  Set rstperso = g_bdPerso.OpenRecordset("select * from donsperso order by nom", dbOpenDynaset)
  Do While Not rstperso.EOF
    .AddItem rstperso!nom
    rstperso.MoveNext
  Loop
  cmbNom.ListIndex = cmbNom.ListCount - 1
End With
rstperso.Close

End Sub
Private Sub btnFermer_Click()
MiseAJourDons g_bdPerso
Unload Me
End Sub

Private Sub btnNouveau_Click()
NouveauDon
End Sub

Private Sub btnPrecedent_Click()
If cmbNom.ListIndex = 0 Then
  AfficheDon (cmbNom.ListCount - 1)
Else
  AfficheDon (cmbNom.ListIndex - 1)
End If
End Sub
Private Sub btnSuivant_Click()
If cmbNom.ListCount = cmbNom.ListIndex + 1 Then
  AfficheDon (0)
Else
  AfficheDon (cmbNom.ListIndex + 1)
End If

End Sub

Private Sub ChkPouvoir_Click()
  If ChkPouvoir Then
    cmbCaracteristique.Enabled = True
  Else
    cmbCaracteristique.ListIndex = -1
    cmbCaracteristique.Enabled = False
  End If
End Sub

Private Sub cmbNom_Click()
AfficheDon (cmbNom.ListIndex)
End Sub

Private Sub Form_Load()
Dim rstperso As Recordset, rstRoles As Recordset

With cmbNom
  .Clear
  Set rstperso = g_bdPerso.OpenRecordset("select * from donsperso order by nom", dbOpenDynaset)
  Do While Not rstperso.EOF
    .AddItem rstperso!nom
    rstperso.MoveNext
  Loop
End With
rstperso.Close

With cmbGenre
  .Clear
  .AddItem ""
  Set rstRoles = g_bdRoles.OpenRecordset("select * from genre order by nom", dbOpenDynaset)
  Do While Not rstRoles.EOF
    .AddItem rstRoles!nom
    rstRoles.MoveNext
  Loop
End With
rstRoles.Close

With cmbSource
  .Clear
  .AddItem ""
  Set rstRoles = g_bdRoles.OpenRecordset("select * from source order by nom", dbOpenDynaset)
  Do While Not rstRoles.EOF
    .AddItem rstRoles!nom
    rstRoles.MoveNext
  Loop
End With
rstRoles.Close

With cmbCaracteristique
  .Clear
  .AddItem ""
  Set rstRoles = g_bdRoles.OpenRecordset("select * from caracteristiques order by nom", dbOpenDynaset)
  Do While Not rstRoles.EOF
    If rstRoles!nom <> "Modificateurs" Then
      .AddItem rstRoles!nom
    End If
    rstRoles.MoveNext
  Loop
    .Enabled = False
End With
rstRoles.Close
If cmbNom.ListCount > 0 Then
  AfficheDon (0)
Else
  NouveauDon
End If
End Sub
Sub AfficheDon(n As Integer)
Dim rstperso As Recordset

cmbNom.ListIndex = n
Set rstperso = g_bdPerso.OpenRecordset("select * from donsperso where nom='" & FormatStrSQL(cmbNom) & "'", dbOpenDynaset)
With rstperso
  If Not .EOF Then
    fldnVersion = !version
    cmbGenre = !Genre
    cmbSource = IIf(IsNull(!Source), "", !Source)
    fldstrResume = IIf(IsNull(!résumé), "", !résumé)
    fldstrExplication = IIf(IsNull(!Explication), "", !Explication)
    cmbCaracteristique = IIf(IsNull(!Caracteristique), "", !Caracteristique)
    ChkPrive = IIf(!prive, 1, 0)
    ChkLiearme = IIf(!liearme, 1, 0)
    ChkMultiple = IIf(!Multiple, 1, 0)
    ChkEpique = IIf(!epique, 1, 0)
    ChkPouvoir = IIf(!pouvoir, 1, 0)
    ChkTare = IIf(!tare, 1, 0)
    ChkTrait = IIf(!trait, 1, 0)
  End If
End With
rstperso.Close
End Sub
Sub NouveauDon()
cmbNom.AddItem ""
cmbNom.ListIndex = cmbNom.ListCount - 1
''cmbNom.SetFocus
cmbGenre.ListIndex = 0
cmbSource.ListIndex = 0
fldnVersion = "3.5"
fldnPage = ""
fldstrResume = ""
fldstrExplication = ""
cmbCaracteristique = ""
ChkPrive = 0
ChkLiearme = 0
ChkMultiple = 0
ChkEpique = 0
ChkPouvoir = 0
ChkTare = 0
ChkTrait = 0
End Sub
