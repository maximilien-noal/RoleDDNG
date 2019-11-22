VERSION 5.00
Begin VB.Form frmDescriptifDon 
   Caption         =   "Description du don"
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   8415
   ClipControls    =   0   'False
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
   LinkTopic       =   "Form1"
   ScaleHeight     =   8490
   ScaleWidth      =   8415
   Begin VB.CommandButton btnSuivant 
      Caption         =   "Suivant"
      Height          =   405
      Left            =   4280
      TabIndex        =   20
      Top             =   8500
      Width           =   1275
   End
   Begin VB.CommandButton btnPrecedent 
      Caption         =   "Précédent"
      Height          =   405
      Left            =   2880
      TabIndex        =   19
      Top             =   8500
      Width           =   1275
   End
   Begin VB.ComboBox CmbNomDon 
      Height          =   345
      Left            =   120
      TabIndex        =   18
      Text            =   "Combo1"
      Top             =   8500
      Width           =   2655
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7600
      Left            =   45
      TabIndex        =   3
      Top             =   840
      Width           =   8295
      Begin VB.TextBox fldstrResume 
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
         Height          =   660
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   13
         Top             =   500
         Width           =   7995
      End
      Begin VB.TextBox fldstrSpecial 
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
         Height          =   650
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   11
         Top             =   6800
         Width           =   7995
      End
      Begin VB.TextBox fldstrNormal 
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
         Height          =   650
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Top             =   5800
         Width           =   7995
      End
      Begin VB.TextBox fldstrCondition 
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
         Height          =   660
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Top             =   1500
         Width           =   7995
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
         Height          =   2970
         Left            =   90
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   2500
         Width           =   7995
      End
      Begin VB.Label Label1 
         Caption         =   "Résumé"
         Height          =   240
         Index           =   4
         Left            =   120
         TabIndex        =   12
         Top             =   200
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Spécial"
         Height          =   240
         Index           =   3
         Left            =   120
         TabIndex        =   10
         Top             =   6500
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Normal"
         Height          =   240
         Index           =   2
         Left            =   120
         TabIndex        =   8
         Top             =   5500
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Condition"
         Height          =   240
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Avantage"
         Height          =   240
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   2200
         Width           =   855
      End
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
      Height          =   405
      Left            =   5680
      TabIndex        =   2
      Top             =   8500
      Width           =   1275
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
      Height          =   405
      Left            =   7080
      TabIndex        =   1
      Top             =   8500
      Width           =   1275
   End
   Begin VB.Label LabType 
      Alignment       =   2  'Center
      BackColor       =   &H80000000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3240
      TabIndex        =   17
      Top             =   480
      Width           =   1695
   End
   Begin VB.Label LabSource 
      Alignment       =   2  'Center
      BackColor       =   &H80000000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   5160
      TabIndex        =   16
      Top             =   480
      Width           =   3015
   End
   Begin VB.Label LabCategorie 
      Alignment       =   2  'Center
      BackColor       =   &H80000000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   120
      TabIndex        =   15
      Top             =   480
      Width           =   3015
   End
   Begin VB.Label LabGenre 
      Alignment       =   2  'Center
      BackColor       =   &H80000000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   5160
      TabIndex        =   14
      Top             =   120
      Width           =   3015
   End
   Begin VB.Label LabDon 
      Alignment       =   2  'Center
      BackColor       =   &H80000000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4815
   End
End
Attribute VB_Name = "frmDescriptifDon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public m_Descriptif As Boolean
Option Explicit
Private Sub btnEnregistrer_Click()

ChampChaine fldstrExplication
fldstrExplication = FormatChaine(fldstrExplication)
fldstrResume = FormatChaine(fldstrResume)
fldstrCondition = FormatChaine(fldstrCondition)
fldstrNormal = FormatChaine(fldstrNormal)
fldstrSpecial = FormatChaine(fldstrSpecial)

  If LabDon <> "" And Trim(fldstrExplication) <> "" Then
    If Trim(fldstrResume) <> "" Then g_bdRoles.Execute "update dons set Résumé='" & FormatStrSQL(Trim(fldstrResume)) & "' where nom='" & FormatStrSQL(LabDon) & "'", dbFailOnError
    If Trim(fldstrCondition) <> "" Then g_bdRoles.Execute "update dons set Condition='" & FormatStrSQL(Trim(fldstrCondition)) & "' where nom='" & FormatStrSQL(LabDon) & "'", dbFailOnError
      g_bdRoles.Execute "update dons set explication='" & FormatStrSQL(Trim(fldstrExplication)) & "' where nom='" & FormatStrSQL(LabDon) & "'", dbFailOnError
    If Trim(fldstrNormal) <> "" Then g_bdRoles.Execute "update dons set NORMAL='" & FormatStrSQL(Trim(fldstrNormal)) & "' where nom='" & FormatStrSQL(LabDon) & "'", dbFailOnError
    If Trim(fldstrSpecial) <> "" Then g_bdRoles.Execute "update dons set Spécial='" & FormatStrSQL(Trim(fldstrSpecial)) & "' where nom='" & FormatStrSQL(LabDon) & "'", dbFailOnError
  End If
  Affiche
End Sub
Private Function ChampChaine(Chaine As String)
Dim strlig As String
Dim Pos As Long

Chaine = Replace(Chaine, "Condition.", "Conditions.")

If InStr(1, Chaine, "Avantage.", vbBinaryCompare) <> 0 Then
  strlig = Right(Chaine, Len(Chaine) - InStr(1, Chaine, "Avantage.", vbBinaryCompare) - 9)
  If InStr(1, Chaine, "Normal.", vbBinaryCompare) <> 0 Then
    Pos = InStr(1, strlig, "Normal.", vbBinaryCompare)
  Else
    If InStr(1, strlig, "Spécial.", vbBinaryCompare) <> 0 Then
      Pos = InStr(1, strlig, "Spécial.", vbBinaryCompare)
    Else
      Pos = Len(strlig)
    End If
  End If
  fldstrExplication = left(strlig, Pos - 1)
Else
  Exit Function
End If

If InStr(1, Chaine, "Conditions.", vbBinaryCompare) <> 0 Then
  fldstrResume = left(Chaine, InStr(1, Chaine, "Conditions.", vbBinaryCompare) - 1)
  If InStr(1, Chaine, "Avantage.", vbBinaryCompare) <> 0 Then
    strlig = left(Chaine, InStr(1, Chaine, "Avantage.", vbBinaryCompare) - 1)
    fldstrCondition = Right(strlig, Len(strlig) - InStr(1, strlig, "Conditions.", vbBinaryCompare) - 11)
  End If
Else
  If InStr(1, Chaine, "Avantage.", vbBinaryCompare) <> 0 Then
    fldstrResume = left(Chaine, InStr(1, Chaine, "Avantage.", vbBinaryCompare) - 1)
  End If
End If

If InStr(1, Chaine, "Normal.", vbBinaryCompare) <> 0 Then
  strlig = Right(Chaine, Len(Chaine) - InStr(1, Chaine, "Normal.", vbBinaryCompare) - 7)
  If InStr(1, Chaine, "Spécial.", vbBinaryCompare) <> 0 Then
    Pos = InStr(1, strlig, "Spécial.", vbBinaryCompare)
  Else
    Pos = Len(strlig)
  End If
  fldstrNormal = left(strlig, Pos - 1)
End If

If InStr(1, Chaine, "Spécial.", vbBinaryCompare) <> 0 Then
  strlig = Right(Chaine, Len(Chaine) - InStr(1, Chaine, "Spécial.", vbBinaryCompare) - 8)
  Pos = Len(strlig)
  fldstrSpecial = left(strlig, Pos - 1)
End If

End Function
Private Function FormatChaine(Chaine As String) As String
Dim Texte As String

Texte = Chaine
Texte = Trim(Texte)

Texte = Replace(Texte, "Conditions.", "")
Texte = Replace(Texte, "Avantage.", "")
Texte = Replace(Texte, "Normal.", "")
Texte = Replace(Texte, "Spécial.", "")
Texte = Replace(Texte, "." & vbCrLf, "^p")
Texte = Replace(Texte, ". " & vbCrLf, "^p")
Texte = Replace(Texte, ":" & vbCrLf, "^d")
Texte = Replace(Texte, ": " & vbCrLf, "^d")
Texte = Replace(Texte, vbCrLf, " ")
Texte = Replace(Texte, "  ", " ")
Texte = Replace(Texte, "^p", "." & vbCrLf)
Texte = Replace(Texte, "^d", ":" & vbCrLf)
If Len(Texte) > 1 Then
  If InStrRev(Texte, vbCrLf, -1, vbTextCompare) = Len(Texte) - 1 Then
    Texte = left(Texte, Len(Texte) - 2)
  End If
End If
FormatChaine = Texte
End Function
Private Sub btnPrecedent_Click()

If m_Descriptif Then
  If CmbNomDon.ListIndex - 1 = -1 Then
    CmbNomDon.ListIndex = CmbNomDon.ListCount - 1
  Else
    CmbNomDon.ListIndex = CmbNomDon.ListIndex - 1
  End If
End If
End Sub

Private Sub btnSuivant_Click()

If m_Descriptif Then
  If CmbNomDon.ListIndex + 1 = CmbNomDon.ListCount Then
    CmbNomDon.ListIndex = 0
  Else
    CmbNomDon.ListIndex = CmbNomDon.ListIndex + 1
  End If
End If
End Sub

Private Sub btnFermer_Click()
  Unload Me
End Sub

Private Sub CmbNomDon_Click()
  LabDon = CmbNomDon
  Affiche
End Sub

Private Sub Form_Activate()
Dim rst As Recordset
  
  If m_Descriptif Then
    btnSuivant.Enabled = True
    btnPrecedent.Enabled = True
    CmbNomDon.Clear
    Set rst = g_bdRoles.OpenRecordset("select Nom from dons order by nom", dbOpenDynaset)
    Do While Not rst.EOF
      CmbNomDon.AddItem rst!Nom
      rst.MoveNext
    Loop
    rst.Close
    CmbNomDon.ListIndex = 0
    LabDon = CmbNomDon
  Else
    btnSuivant.Enabled = False
    btnPrecedent.Enabled = False
    CmbNomDon.Enabled = False
  End If
  Affiche
End Sub
Public Sub Affiche()
Dim rst As Recordset

  If LabDon <> "" Then
    Set rst = g_bdRoles.OpenRecordset("select * from dons where nom='" & FormatStrSQL(LabDon) & "'", dbOpenDynaset)
    If Not rst.EOF Then
      LabGenre = IIf(IsNull(rst!Genre), "", Trim(rst!Genre))
      LabCategorie = IIf(IsNull(rst!Categorie), "", Trim(rst!Categorie))
      LabType = IIf(IsNull(rst!Type), "", Trim(rst!Type))
      LabSource = IIf(IsNull(rst!Source), "", Trim(rst!Source))
      fldstrResume = IIf(IsNull(rst!résumé), "", Trim(rst!résumé))
      fldstrCondition = IIf(IsNull(rst!Condition), "", Trim(rst!Condition))
      fldstrExplication = IIf(IsNull(rst!Explication), "", Trim(rst!Explication))
      fldstrNormal = IIf(IsNull(rst!NORMAL), "", Trim(rst!NORMAL))
      fldstrSpecial = IIf(IsNull(rst!Spécial), "", Trim(rst!Spécial))
    End If
    rst.Close
  End If
End Sub

Private Sub Form_Resize()
Dim W As Integer, H As Integer
Dim Agrandissement As Double

'' largeur
W = maxi_(1, frmDescriptifDon.Width - 540)
Frame1.Width = maxi_(1, frmDescriptifDon.Width - 240)
fldstrResume.Width = W
fldstrCondition.Width = W
fldstrExplication.Width = W
fldstrNormal.Width = W
fldstrSpecial.Width = W
LabDon.Width = W * 0.6
LabGenre.Width = W * 0.377
LabCategorie.Width = LabGenre.Width
LabType.Width = W * 0.212
LabSource.Width = LabGenre.Width

LabGenre.left = LabDon.Width + 345
LabType.left = LabCategorie.Width + 200
LabSource.left = LabType.left + LabType.Width + 245

''hauteur
If frmDescriptifDon.WindowState <> vbMaximized Then
  frmDescriptifDon.Height = mini_(frmMain.Height - 400, maxi_(frmDescriptifDon.Height, 9500 - 2805))
End If
H = frmDescriptifDon.Height - 9500
Agrandissement = (5610 + H) / 5610

Frame1.Height = 7600 + H
fldstrResume.Height = 660 * Agrandissement
fldstrCondition.Height = 660 * Agrandissement
fldstrExplication.Height = 2970 * Agrandissement
fldstrNormal.Height = 660 * Agrandissement
fldstrSpecial.Height = 660 * Agrandissement

fldstrCondition.top = 1500 + 660 * (Agrandissement - 1)
Label1(1).top = 1200 + 660 * (Agrandissement - 1)

fldstrExplication.top = 2500 + 2 * 660 * (Agrandissement - 1)
Label1(0).top = 2200 + 2 * 660 * (Agrandissement - 1)

fldstrNormal.top = 5800 + (2 * 660 + 2970) * (Agrandissement - 1)
Label1(2).top = 5500 + (2 * 660 + 2970) * (Agrandissement - 1)
fldstrSpecial.top = 6800 + (3 * 660 + 2970) * (Agrandissement - 1)
Label1(3).top = 6500 + (3 * 660 + 2970) * (Agrandissement - 1)

CmbNomDon.top = 8500 + (4 * 660 + 2970) * (Agrandissement - 1)
btnPrecedent.top = CmbNomDon.top
btnSuivant.top = CmbNomDon.top
btnEnregistrer.top = CmbNomDon.top
btnFermer.top = CmbNomDon.top

End Sub
