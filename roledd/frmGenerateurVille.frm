VERSION 5.00
Begin VB.Form frmGenerateurVille 
   Caption         =   "Générateur de ville"
   ClientHeight    =   8895
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6825
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
   MDIChild        =   -1  'True
   ScaleHeight     =   8895
   ScaleWidth      =   6825
   Begin VB.CommandButton cmbPrint 
      Caption         =   "Imprimer"
      Height          =   375
      Left            =   4440
      TabIndex        =   12
      Top             =   1200
      Width           =   2295
   End
   Begin VB.CommandButton cmdGenerer 
      Caption         =   "Générer"
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   1200
      Width           =   2295
   End
   Begin VB.TextBox fldstrSortie 
      Height          =   7100
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   10
      Top             =   1680
      Width           =   6650
   End
   Begin VB.TextBox fldnDime 
      Height          =   375
      Left            =   6240
      TabIndex        =   9
      Text            =   "0"
      Top             =   600
      Width           =   495
   End
   Begin VB.TextBox fldnImpot 
      Height          =   375
      Left            =   4800
      TabIndex        =   8
      Text            =   "0"
      Top             =   600
      Width           =   495
   End
   Begin VB.TextBox fldstrnom 
      Height          =   375
      Left            =   1920
      TabIndex        =   5
      Text            =   "Nom"
      Top             =   120
      Width           =   1455
   End
   Begin VB.ComboBox CmbType 
      Height          =   345
      ItemData        =   "frmGenerateurVille.frx":0000
      Left            =   1920
      List            =   "frmGenerateurVille.frx":000D
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   615
      Width           =   1935
   End
   Begin VB.TextBox fldnPopulation 
      Height          =   375
      Left            =   5760
      TabIndex        =   0
      Text            =   "0"
      Top             =   120
      Width           =   975
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "% de dîme"
      Height          =   225
      Index           =   4
      Left            =   5400
      TabIndex        =   7
      Top             =   675
      Width           =   750
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "% d'impôt"
      Height          =   225
      Index           =   3
      Left            =   3960
      TabIndex        =   6
      Top             =   675
      Width           =   750
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "Nom de la communauté"
      Height          =   225
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   195
      Width           =   1695
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "Type de communauté"
      Height          =   225
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   675
      Width           =   1575
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "Population de la communauté"
      Height          =   225
      Index           =   1
      Left            =   3480
      TabIndex        =   1
      Top             =   195
      Width           =   2130
   End
End
Attribute VB_Name = "frmGenerateurVille"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmbPrint_Click()
    Printer.TrackDefault = True
    Printer.ScaleMode = 6
    Printer.Orientation = 1
    Printer.PaperSize = 9
    Printer.PrintQuality = -4
    Printer.ColorMode = 1
    'Printer.CurrentX = 20 - (Printer.Width - Printer.ScaleWidth) / 2
    'Printer.CurrentY = 15 - (Printer.Height - Printer.ScaleHeight) / 2
    Printer.Print fldstrSortie
    Printer.EndDoc
End Sub

Private Sub Form_Load()
CmbType.ListIndex = 0
End Sub
Private Sub cmdGenerer_Click()
'' test des entrées utilisateurs
If Trim(fldstrnom) = "" Then
  MsgBox "Saisie d'un nom obligatoire"
  fldstrnom.SetFocus
  Exit Sub
End If
If IsNumeric(fldnPopulation) = False Then
  MsgBox "Saisie d'une valeur numérique pour la population"
  fldnPopulation.SetFocus
  Exit Sub
End If
If IsNumeric(fldnImpot) = False Then
  MsgBox "Saisie d'une valeur numérique pour les impôts"
  fldnImpot.SetFocus
  Exit Sub
End If
If IsNumeric(fldnDime) = False Then
  MsgBox "Saisie d'une valeur numérique pour la dîme"
  fldnDime.SetFocus
  Exit Sub
End If
fldstrSortie = ""
fldstrSortie = fldstrnom & ", " & TypeVille(fldnPopulation) & " de " & fldnPopulation & " habitants." & vbCrLf
fldstrSortie = fldstrSortie & "Limite financière : " & LimiteFinanciere(fldnPopulation) & " Po" & vbCrLf
fldstrSortie = fldstrSortie & "Liquidite disponible : " & LiquiditeDisponible(fldnPopulation) & " Po" & vbCrLf
fldstrSortie = fldstrSortie & vbCrLf & "Les revenus" & vbCrLf
fldstrSortie = fldstrSortie & "Revenus en or : " & RevenuOr(fldnPopulation) & " Po/mois" & vbCrLf
fldstrSortie = fldstrSortie & "Revenus en matières premières : " & RevenuPremiere(fldnPopulation) & " Po/mois" & vbCrLf
fldstrSortie = fldstrSortie & "Revenus totaux : " & Revenu(fldnPopulation) & " Po/mois" & vbCrLf
fldstrSortie = fldstrSortie & vbCrLf & "Les impôts" & vbCrLf
fldstrSortie = fldstrSortie & "Impôts en or : " & RevenuOr(fldnPopulation) * fldnImpot / 100 & " Po/mois" & vbCrLf
fldstrSortie = fldstrSortie & "Impôts en matières premières : " & RevenuPremiere(fldnPopulation) * fldnImpot / 100 & " Po/mois" & vbCrLf
fldstrSortie = fldstrSortie & "Impôts totaux : " & Revenu(fldnPopulation) * fldnImpot / 100 & " Po/mois" & vbCrLf
fldstrSortie = fldstrSortie & vbCrLf & "La dîme" & vbCrLf
fldstrSortie = fldstrSortie & "Dîme en or : " & RevenuOr(fldnPopulation) * fldnDime / 100 & " Po/mois" & vbCrLf
fldstrSortie = fldstrSortie & "Dîme en matières premières : " & RevenuPremiere(fldnPopulation) * fldnDime / 100 & " Po/mois" & vbCrLf
fldstrSortie = fldstrSortie & "Dîme totaux : " & Revenu(fldnPopulation) * fldnDime / 100 & " Po/mois" & vbCrLf
fldstrSortie = fldstrSortie & vbCrLf & "Les instances" & vbCrLf
fldstrSortie = fldstrSortie & Instances(fldnPopulation) & vbCrLf
fldstrSortie = fldstrSortie & vbCrLf & "Les Pnjs" & vbCrLf
fldstrSortie = fldstrSortie & PNJ(fldnPopulation)
fldstrSortie = fldstrSortie & vbCrLf & "Mélange de races" & vbCrLf
fldstrSortie = fldstrSortie & RACE(CmbType, fldnPopulation) & vbCrLf

End Sub

Function TypeVille(Population As Long) As String
    If Population <= 80 Then
        TypeVille = "Lieu dit"
        Exit Function
    End If
    If Population <= 400 Then
        TypeVille = "Hameau"
        Exit Function
    End If
    If Population <= 900 Then
        TypeVille = "Village"
        Exit Function
    End If
    If Population <= 2000 Then
        TypeVille = "Bourg"
        Exit Function
    End If
    If Population <= 5000 Then
        TypeVille = "Ville"
        Exit Function
    End If
    If Population <= 12000 Then
        TypeVille = "Grande ville"
        Exit Function
    End If
    If Population <= 25000 Then
        TypeVille = "Cité"
        Exit Function
    End If
    TypeVille = "Métropole"
End Function
Function LimiteFinanciere(Population As Long) As Long
    If Population <= 0 Then
        LimiteFinanciere = 0
        Exit Function
    End If
    If Population <= 80 Then
        LimiteFinanciere = 40
        Exit Function
    End If
    If Population <= 400 Then
        LimiteFinanciere = 100
        Exit Function
    End If
    If Population <= 900 Then
        LimiteFinanciere = 200
        Exit Function
    End If
    If Population <= 2000 Then
        LimiteFinanciere = 800
        Exit Function
    End If
    If Population <= 5000 Then
        LimiteFinanciere = 3000
        Exit Function
    End If
    If Population <= 12000 Then
        LimiteFinanciere = 15000
        Exit Function
    End If
    If Population <= 25000 Then
        LimiteFinanciere = 40000
        Exit Function
    End If
    LimiteFinanciere = 100000
End Function
Public Function LiquiditeDisponible(Population As Long) As Long
    If Population <= 80 Then
        LiquiditeDisponible = 2 * Population
        Exit Function
    End If
    If Population <= 400 Then
        LiquiditeDisponible = 5 * Population
        Exit Function
    End If
    If Population <= 900 Then
        LiquiditeDisponible = 10 * Population
        Exit Function
    End If
    If Population <= 2000 Then
        LiquiditeDisponible = 40 * Population
        Exit Function
    End If
    If Population <= 5000 Then
        LiquiditeDisponible = 150 * Population
        Exit Function
    End If
    If Population <= 12000 Then
        LiquiditeDisponible = 750 * Population
        Exit Function
    End If
    If Population <= 25000 Then
        LiquiditeDisponible = 2000 * Population
        Exit Function
    End If
    LiquiditeDisponible = 5000 * Population
End Function
Public Function Revenu(Population As Long) As Long
    If Population <= 80 Then
        Revenu = 2.5 * Population
        Exit Function
    End If
    If Population <= 400 Then
        Revenu = 3 * Population
        Exit Function
    End If
    If Population <= 900 Then
        Revenu = 3.7 * Population
        Exit Function
    End If
    If Population <= 2000 Then
        Revenu = 4.5 * Population
        Exit Function
    End If
    If Population <= 5000 Then
        Revenu = 5.5 * Population
        Exit Function
    End If
    If Population <= 12000 Then
        Revenu = 6.7 * Population
        Exit Function
    End If
    If Population <= 25000 Then
        Revenu = 8.2 * Population
        Exit Function
    End If
    Revenu = 10 * Population
End Function
Public Function RevenuOr(Population As Long) As Long
    If Population <= 80 Then
        RevenuOr = 2.5 * Population * 0.05
        Exit Function
    End If
    If Population <= 400 Then
        RevenuOr = 3 * Population * 0.1
        Exit Function
    End If
    If Population <= 900 Then
        RevenuOr = 3.7 * Population * 0.2
        Exit Function
    End If
    If Population <= 2000 Then
        RevenuOr = 4.5 * Population * 0.3
        Exit Function
    End If
    If Population <= 5000 Then
        RevenuOr = 5.5 * Population * 0.4
        Exit Function
    End If
    If Population <= 12000 Then
        RevenuOr = 6.7 * Population * 0.5
        Exit Function
    End If
    If Population <= 25000 Then
        RevenuOr = 8.2 * Population * 0.65
        Exit Function
    End If
    RevenuOr = 10 * Population * 0.8
End Function
Public Function RevenuPremiere(Population As Long) As Long
    If Population <= 80 Then
        RevenuPremiere = 2.5 * Population * 0.95
        Exit Function
    End If
    If Population <= 400 Then
        RevenuPremiere = 3 * Population * 0.9
        Exit Function
    End If
    If Population <= 900 Then
        RevenuPremiere = 3.7 * Population * 0.8
        Exit Function
    End If
    If Population <= 2000 Then
        RevenuPremiere = 4.5 * Population * 0.7
        Exit Function
    End If
    If Population <= 5000 Then
        RevenuPremiere = 5.5 * Population * 0.6
        Exit Function
    End If
    If Population <= 12000 Then
        RevenuPremiere = 6.7 * Population * 0.5
        Exit Function
    End If
    If Population <= 25000 Then
        RevenuPremiere = 8.2 * Population * 0.35
        Exit Function
    End If
    RevenuPremiere = 10 * Population * 0.2
End Function
Public Function Instances(Population As Long) As String
Dim D20 As Integer
Dim Nombre As Integer
Dim i As Integer

Nombre = 1
    If Population > 5000 Then
        Nombre = Nombre + 1
    End If
    If Population > 12000 Then
        Nombre = Nombre + 1
    End If
    If Population > 25000 Then
        Nombre = Nombre + 1
    End If
For i = 1 To Nombre
    If Instances <> "" Then
      Instances = Instances & ", "
    End If
    If Population >= 0 Then
        D20 = des(1, 20, -1)
    End If
    If Population > 80 Then
        D20 = des(1, 20, 0)
    End If
    If Population > 400 Then
        D20 = des(1, 20, 1)
    End If
    If Population > 900 Then
        D20 = des(1, 20, 2)
    End If
    If Population > 2000 Then
        D20 = des(1, 20, 3)
    End If
    If Population > 5000 Then
        D20 = des(1, 20, 4)
    End If
    If Population > 12000 Then
        D20 = des(1, 20, 5)
    End If
    If Population > 25000 Then
        D20 = des(1, 20, 6)
    End If
    If D20 < 14 Then
      Instances = Instances & "Traditionnel"
      If des(1, 20, 0) = 20 Then
        Instances = Instances & " et influence monstrueuse"
      End If
    Else
      If D20 > 18 Then
        Instances = Instances & "Magique"
      Else
        Instances = Instances & "Inhabituel"
      End If
    End If
    Instances = Instances & " d'alignement " & Alignement()
Next i
End Function
Public Function des(Nombre As Integer, Face As Integer, Plus As Integer) As Integer
Dim i As Integer

des = 0
  For i = 1 To Nombre
    des = des + Int(Face * Rnd(1000)) + 1
  Next i
  des = des + Plus
End Function
Public Function Alignement() As String
Dim D100 As Integer

  D100 = des(1, 100, 0)
  
    If D100 > 0 Then
        Alignement = "Loyal bon"
    End If
    If D100 > 35 Then
        Alignement = "Neutre bon"
    End If
    If D100 > 39 Then
        Alignement = "Chaotique bon"
    End If
    If D100 > 41 Then
        Alignement = "Loyal neutre"
    End If
    If D100 > 61 Then
        Alignement = "Neutre"
    End If
    If D100 > 63 Then
        Alignement = "Chaotique neutre"
    End If
    If D100 > 64 Then
        Alignement = "Loyal mauvais"
    End If
    If D100 > 90 Then
        Alignement = "Neutre mauvais"
    End If
    If D100 > 98 Then
        Alignement = "Chaotique mauvais"
    End If
End Function
Public Function PNJ(Population As Long) As String
Dim Modificateur As Integer, BonusModificateur As Integer
Dim Nombre As Integer, Reste As Long
Dim i As Integer, j As Integer, TotalPnj As Integer, Multiplicateur As Integer, k As Integer
Dim Niveau As Integer
Dim TabPnj(16) As S_Tableau
Dim NiveauPnj() As Integer
Dim Commandant As Double
Dim NiveauCommandant As Integer

Commandant = Int(Rnd(1000) * 100) + 1
If Commandant < 61 Then
  Commandant = 8.1
  Else
    If Commandant < 81 Then
      Commandant = 7.2
    Else
      Commandant = 7.1
    End If
End If

TabPnj(0).Classe = "Adepte": TabPnj(0).Nombre = 1: TabPnj(0).Face = 6
TabPnj(1).Classe = "Barbare": TabPnj(1).Nombre = 1: TabPnj(1).Face = 4
TabPnj(2).Classe = "Barde": TabPnj(2).Nombre = 1: TabPnj(2).Face = 6
TabPnj(3).Classe = "Druide": TabPnj(3).Nombre = 1: TabPnj(3).Face = 6
TabPnj(4).Classe = "Ensorceleur": TabPnj(4).Nombre = 1: TabPnj(4).Face = 4
TabPnj(5).Classe = "Expert": TabPnj(5).Nombre = 3: TabPnj(5).Face = 4
TabPnj(6).Classe = "Gens du peuple": TabPnj(6).Nombre = 4: TabPnj(6).Face = 4
TabPnj(7).Classe = "Guerrier": TabPnj(7).Nombre = 1: TabPnj(7).Face = 8
TabPnj(8).Classe = "Homme d'armes": TabPnj(8).Nombre = 2: TabPnj(8).Face = 4
TabPnj(9).Classe = "Magicien": TabPnj(9).Nombre = 1: TabPnj(9).Face = 4
TabPnj(10).Classe = "Moine": TabPnj(10).Nombre = 1: TabPnj(10).Face = 4
TabPnj(11).Classe = "Noble": TabPnj(11).Nombre = 1: TabPnj(12).Face = 4
TabPnj(12).Classe = "Paladin": TabPnj(12).Nombre = 1: TabPnj(13).Face = 3
TabPnj(13).Classe = "Prêtre": TabPnj(13).Nombre = 1: TabPnj(14).Face = 6
TabPnj(14).Classe = "Rôdeur": TabPnj(14).Nombre = 1: TabPnj(15).Face = 3
TabPnj(15).Classe = "Roublard": TabPnj(15).Nombre = 1: TabPnj(16).Face = 8

PNJ = ""
  If Population >= 0 Then
      Modificateur = -3
  End If
  If Population > 80 Then
      Modificateur = -2
  End If
  If Population > 400 Then
      Modificateur = -1
  End If
  If Population > 900 Then
      Modificateur = 0
  End If
  If Population > 2000 Then
      Modificateur = 3
  End If
  If Population > 5000 Then
      Modificateur = 6
  End If
  If Population > 12000 Then
      Modificateur = 9
  End If
  If Population > 25000 Then
      Modificateur = 12
  End If
  Nombre = maxi_(1, Int(Modificateur / 3))
TotalPnj = 0

For j = 0 To 15
  BonusModificateur = 0
  If (j = 3 Or j = 14) And Modificateur < -1 Then
    If des(1, 20, 0) = 20 Then
      BonusModificateur = 10
    End If
  End If
  ReDim NiveauPnj(28)
  For i = 1 To Nombre
    Niveau = des(TabPnj(j).Nombre, TabPnj(j).Face, Modificateur + BonusModificateur)
    If Niveau > 0 Then
      NiveauPnj(Niveau) = NiveauPnj(Niveau) + 1
      TotalPnj = TotalPnj + 1
      Multiplicateur = 1
      Do While Niveau > 1
        Multiplicateur = Multiplicateur * 2
        Niveau = Int((Niveau + 1) / 2)
        If Not (Niveau = 1 And (TabPnj(j).Classe = "Adepte" Or TabPnj(j).Classe = "Expert" Or TabPnj(j).Classe = "Gens du peuple" Or TabPnj(j).Classe = "Homme d'armes" Or TabPnj(j).Classe = "Noble")) Then
          NiveauPnj(Niveau) = NiveauPnj(Niveau) + Multiplicateur
          TotalPnj = TotalPnj + Multiplicateur
        End If
      Loop
    End If
  Next i
  
  k = 0
  If Niveau > 0 Then
    PNJ = PNJ & TabPnj(j).Classe & " : "
    For i = 28 To 1 Step -1
      If NiveauPnj(i) <> 0 Then
        If k <> 0 Then
          PNJ = PNJ & ", "
        Else
          k = k + 1
        End If
        PNJ = PNJ & NiveauPnj(i) & " de niveau " & i
      End If
    Next i
  End If
  k = -1
  If j = Int(Commandant) Then
    For i = 28 To 1 Step -1
      If NiveauPnj(i) <> 0 Then
        If (Commandant - Int(Commandant)) * 10 < 1 Then
          NiveauCommandant = i
          Exit For
        Else
          If k <> -1 Then
            NiveauCommandant = i
            Exit For
          Else
            k = 0
          End If
        End If
      End If
    Next i
  End If
  If Niveau > 0 Then
    PNJ = PNJ & vbCrLf
  End If
Next j
Reste = fldnPopulation - TotalPnj
PNJ = PNJ & vbCrLf & "Les classes de Pnj de niveau 1" & vbCrLf
PNJ = PNJ & "Gens du peuple : " & Int(0.91 * Reste) & vbCrLf
PNJ = PNJ & "Homme d'armes : " & Int(0.05 * Reste) & vbCrLf
PNJ = PNJ & "Expert : " & Int(0.03 * Reste) & vbCrLf
PNJ = PNJ & "Adepte : " & Int(0.005 * Reste) & vbCrLf
PNJ = PNJ & "Noble : " & Int(0.005 * Reste) & vbCrLf

PNJ = PNJ & vbCrLf & "Commandant de la garde/prévôt" & vbCrLf
If NiveauCommandant > 0 Then
  Select Case Commandant
    Case 8.1
      PNJ = PNJ & "Homme d'armes du plus haut niveau " & NiveauCommandant & vbCrLf
    Case 7.2
      PNJ = PNJ & "Deuxième guerrier de la communauté " & NiveauCommandant & vbCrLf
    Case 7.1
      PNJ = PNJ & "Guerrier de plus haut niveau " & NiveauCommandant & vbCrLf
  End Select
Else
  PNJ = PNJ & "Il n'y a aucun guerrier ou homme d'armes dans cette communauté, choisir un autre PNJ" & vbCrLf
End If
End Function
Public Function RACE(Communaute As String, Population As Long) As String

RACE = ""
  Select Case Communaute
    Case "Isolée"
      RACE = RACE & "Humains : " & Int(Population * 0.96) & vbCrLf
      RACE = RACE & "Halfelins : " & Int(Population * 0.02) & vbCrLf
      RACE = RACE & "Elfes : " & Int(Population * 0.1) & vbCrLf
      RACE = RACE & "Autres races : " & Int(Population * 0.01) & vbCrLf
    Case "Ouverte"
      RACE = RACE & "Humains : " & Int(Population * 0.79) & vbCrLf
      RACE = RACE & "Halfelins : " & Int(Population * 0.09) & vbCrLf
      RACE = RACE & "Elfes : " & Int(Population * 0.05) & vbCrLf
      RACE = RACE & "Nains : " & Int(Population * 0.03) & vbCrLf
      RACE = RACE & "Gnomes : " & Int(Population * 0.02) & vbCrLf
      RACE = RACE & "Demi-elfes : " & Int(Population * 0.01) & vbCrLf
      RACE = RACE & "Demi-orques : " & Int(Population * 0.01) & vbCrLf
    Case "Intégrée"
      RACE = RACE & "Humains : " & Int(Population * 0.37) & vbCrLf
      RACE = RACE & "Halfelins : " & Int(Population * 0.2) & vbCrLf
      RACE = RACE & "Elfes : " & Int(Population * 0.18) & vbCrLf
      RACE = RACE & "Nains : " & Int(Population * 0.1) & vbCrLf
      RACE = RACE & "Gnomes : " & Int(Population * 0.07) & vbCrLf
      RACE = RACE & "Demi-elfes : " & Int(Population * 0.05) & vbCrLf
      RACE = RACE & "Demi-orques : " & Int(Population * 0.03) & vbCrLf
  End Select
End Function
Private Sub Form_Resize()

fldstrSortie.Height = maxi_(1, frmGenerateurVille.Height - 2300)
fldstrSortie.Width = maxi_(1, frmGenerateurVille.Width - 500)

End Sub
