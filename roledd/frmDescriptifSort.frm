VERSION 5.00
Begin VB.Form frmDescriptifSort 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Description du sort"
   ClientHeight    =   7515
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   13350
   ClipControls    =   0   'False
   ControlBox      =   0   'False
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
   MinButton       =   0   'False
   ScaleHeight     =   7515
   ScaleWidth      =   13350
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox CmbSort 
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
      Left            =   120
      TabIndex        =   38
      Text            =   "Combo1"
      Top             =   7200
      Width           =   3975
   End
   Begin VB.CommandButton btnPrecedent 
      Caption         =   "&Précédent"
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
      Left            =   6600
      TabIndex        =   37
      Top             =   7200
      Width           =   1275
   End
   Begin VB.CommandButton btnsuivant 
      Caption         =   "&Suivant"
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
      Left            =   8040
      TabIndex        =   36
      Top             =   7200
      Width           =   1275
   End
   Begin VB.TextBox fldstrClasseNiveau 
      BackColor       =   &H00C0C0C0&
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
      Left            =   10260
      MultiLine       =   -1  'True
      TabIndex        =   34
      Top             =   45
      Width           =   2955
   End
   Begin VB.TextBox fldstrSource 
      BackColor       =   &H00C0C0C0&
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
      Left            =   5160
      MultiLine       =   -1  'True
      TabIndex        =   32
      Top             =   6760
      Width           =   3960
   End
   Begin VB.CommandButton btnEnregistrer 
      Caption         =   "&Enregistrer"
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
      Left            =   10515
      TabIndex        =   25
      Top             =   7200
      Width           =   1275
   End
   Begin VB.Frame frm1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6780
      Left            =   45
      TabIndex        =   2
      Top             =   360
      Width           =   13260
      Begin VB.TextBox fldstrobjet 
         BackColor       =   &H00C0C0C0&
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
         Left            =   10200
         MultiLine       =   -1  'True
         TabIndex        =   35
         Top             =   6400
         Width           =   2760
      End
      Begin VB.TextBox fldnDD 
         BackColor       =   &H00C0C0C0&
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
         Left            =   4275
         MultiLine       =   -1  'True
         TabIndex        =   33
         Top             =   5800
         Width           =   615
      End
      Begin VB.TextBox fldstrzoneeffet 
         BackColor       =   &H00C0C0C0&
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
         Left            =   100
         MultiLine       =   -1  'True
         TabIndex        =   31
         Top             =   4600
         Width           =   4800
      End
      Begin VB.TextBox fldPourDevellopper 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1755
         Left            =   5100
         MultiLine       =   -1  'True
         TabIndex        =   28
         Top             =   1000
         Width           =   8100
      End
      Begin VB.TextBox fldDDDifficultee 
         BackColor       =   &H00C0C0C0&
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
         Left            =   5100
         MultiLine       =   -1  'True
         TabIndex        =   27
         Top             =   400
         Width           =   8100
      End
      Begin VB.TextBox fldstrDomaine 
         BackColor       =   &H00C0C0C0&
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
         Left            =   100
         MultiLine       =   -1  'True
         TabIndex        =   13
         Top             =   400
         Width           =   4800
      End
      Begin VB.TextBox fldstrNiveau 
         BackColor       =   &H00C0C0C0&
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
         Left            =   100
         MultiLine       =   -1  'True
         TabIndex        =   12
         Top             =   1000
         Width           =   4800
      End
      Begin VB.TextBox fldstrComposante 
         BackColor       =   &H00C0C0C0&
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
         Left            =   100
         MultiLine       =   -1  'True
         TabIndex        =   11
         Top             =   1600
         Width           =   4800
      End
      Begin VB.TextBox fldstrtempincantation 
         BackColor       =   &H00C0C0C0&
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
         Left            =   100
         MultiLine       =   -1  'True
         TabIndex        =   10
         Top             =   2200
         Width           =   4800
      End
      Begin VB.TextBox fldstrportee 
         BackColor       =   &H00C0C0C0&
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
         Left            =   100
         MultiLine       =   -1  'True
         TabIndex        =   9
         Top             =   2800
         Width           =   4800
      End
      Begin VB.TextBox fldstrcible 
         BackColor       =   &H00C0C0C0&
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
         Left            =   100
         MultiLine       =   -1  'True
         TabIndex        =   8
         Top             =   3400
         Width           =   4800
      End
      Begin VB.TextBox fldstrdurée 
         BackColor       =   &H00C0C0C0&
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
         Left            =   100
         MultiLine       =   -1  'True
         TabIndex        =   7
         Top             =   5200
         Width           =   4800
      End
      Begin VB.TextBox fldstrEffet 
         BackColor       =   &H00C0C0C0&
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
         Left            =   100
         MultiLine       =   -1  'True
         TabIndex        =   6
         Top             =   4000
         Width           =   4800
      End
      Begin VB.TextBox fldstrjetsauvegarde 
         BackColor       =   &H00C0C0C0&
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
         Left            =   100
         MultiLine       =   -1  'True
         TabIndex        =   5
         Top             =   5800
         Width           =   3765
      End
      Begin VB.TextBox fldstrresistancemagie 
         BackColor       =   &H00C0C0C0&
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
         Left            =   100
         MultiLine       =   -1  'True
         TabIndex        =   4
         Top             =   6400
         Width           =   4800
      End
      Begin VB.TextBox fldstrExplication 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3075
         Left            =   5100
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   3000
         Width           =   8100
      End
      Begin VB.Label Label8 
         Caption         =   "Zone d'effet"
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
         Left            =   120
         TabIndex        =   30
         Top             =   4320
         Width           =   1455
      End
      Begin VB.Label Label6 
         Caption         =   "Pour dévellopper"
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
         Left            =   5160
         TabIndex        =   29
         Top             =   720
         Width           =   1920
      End
      Begin VB.Label Label5 
         Caption         =   "DD pour lancer le sort épique"
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
         Left            =   5200
         TabIndex        =   26
         Top             =   150
         Width           =   2895
      End
      Begin VB.Label Label1 
         Caption         =   "Ecole [Branche] (Registre)"
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
         Left            =   120
         TabIndex        =   24
         Top             =   150
         Width           =   2115
      End
      Begin VB.Label Label2 
         Caption         =   "Niveau"
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
         Left            =   120
         TabIndex        =   23
         Top             =   720
         Width           =   690
      End
      Begin VB.Label Label3 
         Caption         =   "Composante"
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
         Left            =   120
         TabIndex        =   22
         Top             =   1320
         Width           =   960
      End
      Begin VB.Label Label4 
         Caption         =   "Temps d'incantation"
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
         Left            =   120
         TabIndex        =   21
         Top             =   1920
         Width           =   1695
      End
      Begin VB.Label gdfg 
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
         Height          =   240
         Left            =   120
         TabIndex        =   20
         Top             =   2520
         Width           =   1455
      End
      Begin VB.Label Cible 
         Caption         =   "Cible"
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
         Left            =   120
         TabIndex        =   19
         Top             =   3120
         Width           =   1455
      End
      Begin VB.Label Label7 
         Caption         =   "Durée"
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
         Left            =   120
         TabIndex        =   18
         Top             =   4920
         Width           =   555
      End
      Begin VB.Label sfdsf 
         Caption         =   "Effet"
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
         Left            =   120
         TabIndex        =   17
         Top             =   3720
         Width           =   1455
      End
      Begin VB.Label Label9 
         Caption         =   "Jet de sauvegarde"
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
         Left            =   120
         TabIndex        =   16
         Top             =   5520
         Width           =   1455
      End
      Begin VB.Label Label10 
         Caption         =   "Résistance à la magie"
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
         Left            =   120
         TabIndex        =   15
         Top             =   6120
         Width           =   1695
      End
      Begin VB.Label Label11 
         Caption         =   "Explication du sort"
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
         Left            =   5205
         TabIndex        =   14
         Top             =   2760
         Width           =   1455
      End
   End
   Begin VB.CommandButton btnFermer 
      Cancel          =   -1  'True
      Caption         =   "&Fermer"
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
      Left            =   11910
      TabIndex        =   1
      Top             =   7200
      Width           =   1275
   End
   Begin VB.Label labnom 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000006&
      Height          =   330
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   9870
   End
End
Attribute VB_Name = "frmDescriptifSort"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public m_classe As Integer
Public m_niveau As Integer
Public m_Objet As Integer
Public m_Row As Integer
Public m_Index As Integer
Public m_vsSort As String
Public m_Col_DD As Integer

Private Sub btnEnregistrer_Click()
Dim rst As Recordset

  Set rst = g_bdRoles.OpenRecordset("sort", dbOpenTable)
  rst.Index = "PrimaryKey"
  rst.Seek "=", labnom
  If Not rst.NoMatch Then
    rst.Edit
    rst!Ecole = fldstrDomaine
    rst![Classe/Niveau] = fldstrNiveau
    rst!composantes = fldstrComposante
    rst![Temps d’incantation] = fldstrtempincantation
    rst!Portée = fldstrportee
    rst!Cible = fldstrcible
    rst!durée = fldstrdurée
    rst!effet = fldstrEffet
    rst![Jet de sauvegarde] = fldstrjetsauvegarde
    rst![Résistance à la magie] = fldstrresistancemagie
    rst![Explication] = fldstrExplication
    rst![DDdeDifficultee] = IIf(Not IsNumeric(fldDDDifficultee), Null, fldDDDifficultee)
    rst![PourDevellopper] = IIf(Trim(fldPourDevellopper) = "", Null, fldPourDevellopper)
    rst.Update
  End If
  rst.Close
  Form_Load
End Sub

Private Sub btnFermer_Click()
  Unload Me
End Sub

Private Sub btnPrecedent_Click()
Dim rst As Recordset

If CmbSort.ListIndex - 1 = -1 Then
  CmbSort.ListIndex = CmbSort.ListCount - 1
Else
  CmbSort.ListIndex = CmbSort.ListIndex - 1
End If
If CmbSort <> "" Then
  Set rst = g_bdRoles.OpenRecordset("select nom from sort where nom='" & FormatStrSQL(CmbSort) & "'", dbOpenDynaset)
  If rst.EOF Then
    rst.Close
    btnPrecedent_Click
  End If
End If

End Sub

Private Sub btnSuivant_Click()
Dim rst As Recordset

If CmbSort.ListIndex + 1 = CmbSort.ListCount Then
  CmbSort.ListIndex = 0
Else
  CmbSort.ListIndex = CmbSort.ListIndex + 1
End If
If CmbSort <> "" Then
  Set rst = g_bdRoles.OpenRecordset("select nom from sort where nom='" & FormatStrSQL(CmbSort) & "'", dbOpenDynaset)
  If rst.EOF Then
    rst.Close
    btnSuivant_Click
  End If
End If
End Sub

Private Sub CmbSort_Click()
Affiche
End Sub

Private Sub Form_Load()
Dim rst As Recordset
Dim i As Integer

Select Case UCase(m_vsSort)
  Case "VSSORTCLASSE"
    For i = 0 To frmFichePerso.vsSortClasse(m_Index).Rows - 1
      CmbSort.AddItem frmFichePerso.vsSortClasse(m_Index).Cell(flexcpText, i, 0)
    Next i
    CmbSort.ListIndex = m_Row
  Case "VSSORTLISTE"
    For i = 0 To frmFichePerso.vsSortListe(m_Index).Rows - 1
      CmbSort.AddItem frmFichePerso.vsSortListe(m_Index).Cell(flexcpText, i, 0)
    Next i
    CmbSort.ListIndex = m_Row
  Case Else
    CmbSort.Clear
    Set rst = g_bdRoles.OpenRecordset("select Nom from sort order by nom", dbOpenDynaset)
    Do While Not rst.EOF
      CmbSort.AddItem rst!Nom
      rst.MoveNext
    Loop
    rst.Close
    CmbSort.ListIndex = 0
End Select

Affiche

End Sub
Sub Affiche()
Dim rst As Recordset

If CmbSort <> "" Then
    Set rst = g_bdRoles.OpenRecordset("select * from sort where nom='" & FormatStrSQL(CmbSort) & "'", dbOpenDynaset)
    If Not rst.EOF Then
      labnom = rst!Nom
      CmbSort.Text = labnom
      If Val(rst!version) <> 3.5 Then
        labnom = labnom & " (version 3.0)"
        labnom.ForeColor = &HFF&
      Else
        labnom.ForeColor = &H80000012
      End If
      fldstrSource = IIf(IsNull(rst!Source), "", rst!Source)
      fldstrDomaine = IIf(IsNull(rst!Ecole), "", rst!Ecole)
      fldstrNiveau = IIf(IsNull(rst![Classe/Niveau]), "", rst![Classe/Niveau])
      fldstrComposante = IIf(IsNull(rst!composantes), "", rst!composantes)
      fldstrtempincantation = IIf(IsNull(rst![Temps d’incantation]), "", rst![Temps d’incantation])
      fldstrportee = IIf(IsNull(rst!Portée), "", rst!Portée)
      fldstrcible = IIf(IsNull(rst!Cible), "", rst!Cible)
      fldstrEffet = IIf(IsNull(rst!effet), "", rst!effet)
      fldstrzoneeffet = IIf(IsNull(rst![Zone d'effet]), "", rst![Zone d'effet])
      fldstrdurée = IIf(IsNull(rst!durée), "", rst!durée)
      fldstrjetsauvegarde = IIf(IsNull(rst![Jet de sauvegarde]), "", rst![Jet de sauvegarde])
      fldstrresistancemagie = IIf(IsNull(rst![Résistance à la magie]), "", rst![Résistance à la magie])
      fldstrExplication = IIf(IsNull(rst![Explication]), "", rst![Explication])
      fldDDDifficultee = IIf(IsNull(rst![DDdeDifficultee]), "", rst![DDdeDifficultee])
      fldPourDevellopper = IIf(IsNull(rst![PourDevellopper]), "", rst![PourDevellopper])
      
      Select Case UCase(m_vsSort)
        Case "VSSORTCLASSE"
          If m_Index = 0 Then
            fldnDD = frmFichePerso.vsSortClasse(m_Index).Cell(flexcpText, CmbSort.ListIndex, m_Col_DD)
          End If
          fldstrClasseNiveau = frmFichePerso.vsSortClasse(m_Index).Cell(flexcpText, CmbSort.ListIndex, m_classe) & " " & frmFichePerso.vsSortClasse(m_Index).Cell(flexcpText, CmbSort.ListIndex, m_niveau)
          fldstrobjet = frmFichePerso.vsSortClasse(m_Index).Cell(flexcpText, CmbSort.ListIndex, m_Objet)
        Case "VSSORTLISTE"
          fldnDD = frmFichePerso.vsSortListe(m_Index).Cell(flexcpText, CmbSort.ListIndex, m_Col_DD)
          fldstrClasseNiveau = frmFichePerso.vsSortListe(1).Cell(flexcpText, 1, 0) & " " & frmFichePerso.vsSortListe(m_Index).Cell(flexcpText, CmbSort.ListIndex, m_niveau)
          fldstrobjet = ""
        Case Else
          fldnDD = ""
          fldstrClasseNiveau = ""
          fldstrobjet = ""
      End Select
    Else
      
    End If
    rst.Close
  End If
  
  If UCase(left(fldstrNiveau, 6)) = "EPIQUE" Then
    Label11.top = 2835
    fldstrExplication.top = 3200
    fldstrExplication.Height = 3015
    fldPourDevellopper.Visible = True
    Label6.Visible = True
    fldDDDifficultee.Visible = True
    Label5.Visible = True
  Else
    fldstrExplication.top = 400
    fldstrExplication.Height = 5800
    Label11.top = 150
    fldPourDevellopper.Visible = False
    Label6.Visible = False
    fldDDDifficultee.Visible = False
    Label5.Visible = False
  End If
End Sub
