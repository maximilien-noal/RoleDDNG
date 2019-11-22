VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "RoleDD Gestion de personnages pour Donjons et Dragons 3.5"
   ClientHeight    =   13275
   ClientLeft      =   2250
   ClientTop       =   1290
   ClientWidth     =   18240
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   OLEDropMode     =   1  'Manual
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   12945
      Width           =   18240
      _ExtentX        =   32173
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   31644
         EndProperty
      EndProperty
   End
   Begin VB.Menu IDM_FICHIER 
      Caption         =   "&Fichier"
      Begin VB.Menu IDM_ASSISTANT_PERSO 
         Caption         =   "&Assistant création personnage"
         Shortcut        =   ^N
      End
      Begin VB.Menu IDM_GESTION_PERSO 
         Caption         =   "&Gestion des personnages"
         Shortcut        =   ^G
      End
      Begin VB.Menu IDM_ACCEDER 
         Caption         =   "&Accéder aux fiches des personnages"
      End
      Begin VB.Menu IDM_GEST_OBJETS 
         Caption         =   "&Gestion des objets"
         Shortcut        =   ^I
      End
      Begin VB.Menu IDM_SEP_3 
         Caption         =   "-"
      End
      Begin VB.Menu IMD_IMPORT 
         Caption         =   "&Importation de personnage"
         Shortcut        =   ^M
      End
      Begin VB.Menu IDM_SEP_1 
         Caption         =   "-"
      End
      Begin VB.Menu IDM_OUVRIR 
         Caption         =   "&Ouvrir une base de donnée de personnages"
         Shortcut        =   ^O
      End
      Begin VB.Menu IDM_SEP2 
         Caption         =   "-"
      End
      Begin VB.Menu IDM_QUITTER 
         Caption         =   "&Quitter"
      End
   End
   Begin VB.Menu IDM_OUTILS 
      Caption         =   "&Outils"
      Begin VB.Menu IDM_XP_PERSONNAGES 
         Caption         =   "&Expériences des personnages"
         Shortcut        =   ^X
      End
      Begin VB.Menu IDM_GENERATEUR_VILLE 
         Caption         =   "&Générateur de ville"
      End
      Begin VB.Menu IDM_LANCEUR_DES 
         Caption         =   "&Lanceur de dés"
         Shortcut        =   ^D
      End
   End
   Begin VB.Menu IDM_REGLES 
      Caption         =   "&Règles"
      Begin VB.Menu IDM_DESCRIPTION_SORT 
         Caption         =   "Description des &sorts"
         Shortcut        =   ^S
      End
      Begin VB.Menu IDM_DESCRIPTION_RACE 
         Caption         =   "Description des &races"
      End
      Begin VB.Menu IDM_DESCRIPTION_DON 
         Caption         =   "&Description des &dons"
      End
   End
   Begin VB.Menu IDM_PERSONNEL 
      Caption         =   "&Personnel"
      Begin VB.Menu IDM_INSERTION_DONS 
         Caption         =   "&Insertion de dons"
      End
      Begin VB.Menu IDM_INSERTION_RACES 
         Caption         =   "&Insertion de races"
      End
      Begin VB.Menu IDM_INSERTION_CLASSES 
         Caption         =   "&Insertion de classes"
      End
   End
   Begin VB.Menu IDM_OPTIONS 
      Caption         =   "&Options"
   End
   Begin VB.Menu IDM_AIDE 
      Caption         =   "&Aide"
      Begin VB.Menu IDM_ASTUCE 
         Caption         =   "&Astuce du jour"
      End
      Begin VB.Menu IDM_LIENSITE 
         Caption         =   "&Lien vers le site"
      End
      Begin VB.Menu IDM_MISES_JOUR 
         Caption         =   "&Vérifier les mises à jour"
      End
      Begin VB.Menu IDM_APROPOS 
         Caption         =   "&A propos de RoleDD"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Function TrouveChild(strname As String) As Integer
Dim i As Integer
  
  TrouveChild = -1
  For i = 0 To Forms.Count - 1
    If UCase(Forms(i).name) = UCase(strname) Then
      TrouveChild = i
      Exit Function
    End If
  Next i
  Exit Function
End Function
Private Sub IDM_ACCEDER_Click()

  If g_ModeDebug = vbUnchecked Then On Error GoTo IDM_ACCEDER_Click_Error
  
  If TrouveChild("frmAcceder") = -1 Then
    frmAcceder.Show
  Else
    frmAcceder.AccedePerso
  End If

IDM_ACCEDER_Click_Exit:
  Exit Sub

IDM_ACCEDER_Click_Error:
  MsgBox "L'erreur suivante s'est produite : " & Err.Description & vbCrLf & vbCrLf & " numéro d'erreur (" & Err.Number & " )", vbCritical, g_strNomApplication
  Print #g_FileDebug, Format(Now, "dd/MM/yyyy hh:mm:ss") & " Erreur n° : (" & Err.Number & ") " & Err.Description & " Procedure : IDM_ACCEDER_Click  Module : frmMain.frm"
  Resume IDM_ACCEDER_Click_Exit
End Sub

Private Sub IDM_APROPOS_Click()
  frmApropos.Show vbModal
End Sub

Private Sub IDM_ASSISTANT_PERSO_Click()
  frmGestPerso.Show
  frmAssistant.Show
End Sub

Private Sub IDM_ASTUCE_Click()
  frmAstuce.Show
End Sub

Private Sub IDM_DESCRIPTION_DON_Click()
    frmDescriptifDon.m_Descriptif = True
    frmDescriptifDon.Show
End Sub

Private Sub IDM_DESCRIPTION_RACE_Click()
  frmDescriptionRaces.Show
End Sub

Private Sub IDM_DESCRIPTION_SORT_Click()

frmDescriptifSort.m_vsSort = ""
frmDescriptifSort.Show vbModal
End Sub

Private Sub IDM_EXPORT_Click()
'cool
End Sub

Private Sub IDM_GENERATEUR_VILLE_Click()
  frmGenerateurVille.Show
End Sub


Private Sub IDM_GEST_OBJETS_Click()
frmGestobj.Show vbModal
End Sub

Private Sub IDM_GESTION_PERSO_Click()

  If g_ModeDebug = vbUnchecked Then On Error GoTo IDM_GESTION_PERSO_Click_Error
  frmGestPerso.Show

IDM_GESTION_PERSO_Click_Exit:
  Exit Sub

IDM_GESTION_PERSO_Click_Error:
  MsgBox "L'erreur suivante s'est produite : " & Err.Description & vbCrLf & vbCrLf & " numéro d'erreur (" & Err.Number & " )", vbCritical, g_strNomApplication
  Print #g_FileDebug, Format(Now, "dd/MM/yyyy hh:mm:ss") & " Erreur n° : (" & Err.Number & ") " & Err.Description & " Procedure : IDM_GESTION_PERSO_Click  Module : frmMain.frm"
  Resume IDM_GESTION_PERSO_Click_Exit
End Sub

Private Sub IDM_INSERTION_DONS_Click()
frmInsertionDons.Show
End Sub

Private Sub IDM_LANCEUR_DES_Click()
  frmLanceurDes.Show
End Sub

Private Sub IDM_LIENSITE_Click()
  ShellExecute Me.hwnd, "open", "http://bonnarien.dyndns.org", "", "", 0
End Sub

Private Sub IDM_MISES_JOUR_Click()
  ShellExecute Me.hwnd, "open", "http://bonnarien.dyndns.org/programme/test.php?version=" & App.Revision, "", "", 0
End Sub

Private Sub IDM_OUVRIR_Click()
  
  If g_ModeDebug = vbUnchecked Then On Error Resume Next
  

  BDD = FichierOuvrir(Me.hwnd, "Access97 (*.mdb)|*.mdb", 0, OFN_EXPLORER Or OFN_FILEMUSTEXIST Or OFN_PATHMUSTEXIST, "Ouvrir une base donnée", "", App.Path, "")
    
  OuvrirFichier (BDD)
End Sub

Public Sub IDM_QUITTER_Click()

  If g_ModeDebug = vbUnchecked Then On Error GoTo IDM_QUITTER_Click_Error
  Unload Me

IDM_QUITTER_Click_Exit:
  Exit Sub

IDM_QUITTER_Click_Error:
  MsgBox "L'erreur suivante s'est produite : " & Err.Description & vbCrLf & vbCrLf & " numéro d'erreur (" & Err.Number & " )", vbCritical, g_strNomApplication
  Print #g_FileDebug, Format(Now, "dd/MM/yyyy hh:mm:ss") & " Erreur n° : (" & Err.Number & ") " & Err.Description & " Procedure : IDM_QUITTER_Click  Module : frmMain.frm"
  Resume IDM_QUITTER_Click_Exit
End Sub

Private Sub IDM_XP_PERSONNAGES_Click()
frmExperience.Show
End Sub

Private Sub IMD_IMPORT_Click()
Dim StrBDI As String

  If g_ModeDebug = vbUnchecked Then On Error Resume Next
  
  StrBDI = FichierOuvrir(Me.hwnd, "Access97 (*.mdb)|*.mdb", 0, OFN_EXPLORER Or OFN_FILEMUSTEXIST Or OFN_PATHMUSTEXIST, "Ouvrir une base donnée pour l'import de personnage", "", App.Path, "")
  frmImport.m_bBase = False
  If StrBDI <> "" Then
    If UCase(App.Path & "\roles.mdb") = UCase(StrBDI) Then
      MsgBox "On ne peut pas ouvrir la base de donnée du programme !", vbExclamation, g_strNomApplication
      Exit Sub
    End If
    If StrBDI = BDD Then
      MsgBox "Vous ouvrez la base en cours d'utilisation." & vbCrLf & "Vous pourrez seulement dupliquer les personnages en les renommant.", vbExclamation, g_strNomApplication
      frmImport.m_bBase = True
    End If
    Set frmImport.BDI = OpenDatabase(StrBDI)
    frmImport.Show vbModal
  End If

End Sub

Private Sub MDIForm_Load()
Dim i As Integer
Dim n As Double
  
  If g_ModeDebug = vbUnchecked Then On Error GoTo MDIForm_Load_Error
  frmMain.backcolor = vbWindowBackground

  frmMain.Picture = LoadPicture(App.Path & "\back.jpg")
  frmMain.Caption = "RoleDD Gestion de personnages pour Donjons et Dragons 3.5 [" & BDD & "]"
  Me.left = POS_FENETRE(Me.name, "POS_X", 0)
  Me.top = POS_FENETRE(Me.name, "POS_Y", 0)
  frmAcceder.Show
    n = Now / 2
    n = n - Int(n)
    n = Int(n * 1000000)
    For i = 1 To (n - 256 * Int(n / 256))
      Rnd (1000)
    Next i
If Command <> "" Then
  OuvrirFichier Command         'lancement de la lecture du fichier
End If

MDIForm_Load_Exit:
  Exit Sub

MDIForm_Load_Error:
  MsgBox "L'erreur suivante s'est produite : " & Err.Description & vbCrLf & vbCrLf & " numéro d'erreur (" & Err.Number & " )", vbCritical, g_strNomApplication
  Print #g_FileDebug, Format(Now, "dd/MM/yyyy hh:mm:ss") & " Erreur n° : (" & Err.Number & ") " & Err.Description & " Procedure : MDIForm_Load  Module : frmMain.frm"
  Resume MDIForm_Load_Exit
End Sub

Private Sub MDIForm_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim Fichier As Variant
For Each Fichier In Data.Files
    ' Ajoute le chemin complet du fichier dans bdd
    BDD = Fichier
    Exit For
Next
OuvrirFichier BDD

End Sub

Private Sub MDIForm_Unload(Cancel As Integer)

  If g_ModeDebug = vbUnchecked Then On Error GoTo MDIForm_Unload_Error
  POS_FENETRE Me.name, "POS_X", Me.left
  POS_FENETRE Me.name, "POS_Y", Me.top
  g_bdPerso.Close
  g_bdRoles.Close
  End

MDIForm_Unload_Exit:
  Exit Sub

MDIForm_Unload_Error:
  MsgBox "L'erreur suivante s'est produite : " & Err.Description & vbCrLf & vbCrLf & " numéro d'erreur (" & Err.Number & " )", vbCritical, g_strNomApplication
  Print #g_FileDebug, Format(Now, "dd/MM/yyyy hh:mm:ss") & " Erreur n° : (" & Err.Number & ") " & Err.Description & " Procedure : MDIForm_Unload  Module : frmMain.frm"
  Resume MDIForm_Unload_Exit
End Sub
Private Sub OuvrirFichier(BDD As String)
  If BDD <> "" Then
    If UCase(App.Path & "\roles.mdb") = UCase(BDD) Then
      MsgBox "On ne peut pas ouvrir la base du programme", vbExclamation, g_strNomApplication
      Exit Sub
    End If
    
    If TrouveChild("frmAcceder") <> -1 Then
      Unload frmAcceder
    End If
      
    If TrouveChild("frmFichePerso") <> -1 Then
      Unload frmFichePerso
    End If
    
    If TrouveChild("frmGestPerso") <> -1 Then
      Unload frmGestPerso
    End If
    g_bdPerso.Close
    
    Set g_bdPerso = OpenDatabase(BDD)

    ControleVersion g_bdPerso
    frmMain.Caption = "RoleDD Gestion de personnages pour Donjons et Dragons 3.5 [" & BDD & "]"
    Unload frmAcceder
    frmAcceder.Show
    WritePrivateProfileString "FICHIER", "BDD", BDD, g_fichier_ini
  End If


End Sub
