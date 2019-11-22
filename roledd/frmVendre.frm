VERSION 5.00
Begin VB.Form frmVendre 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Vendre un objet"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4860
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
   ScaleHeight     =   3090
   ScaleWidth      =   4860
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton btnSupprimer 
      Caption         =   "Supprimer"
      Height          =   495
      Left            =   240
      TabIndex        =   8
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Frame frmObjet 
      Caption         =   "NonObjet"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   4695
      Begin VB.TextBox fldnPrixTotal 
         Height          =   375
         Left            =   2760
         TabIndex        =   10
         Text            =   "1"
         Top             =   1840
         Width           =   1000
      End
      Begin VB.TextBox fldnNombreObjet 
         Height          =   375
         Left            =   2760
         TabIndex        =   7
         Text            =   "1"
         Top             =   1340
         Width           =   1000
      End
      Begin VB.TextBox fldnPrixUnitaire 
         Height          =   375
         Left            =   2760
         TabIndex        =   5
         Text            =   "1"
         Top             =   840
         Width           =   1000
      End
      Begin VB.Label LblObjet 
         AutoSize        =   -1  'True
         Caption         =   "Prix de revente des objets"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   195
         TabIndex        =   9
         Top             =   1900
         Width           =   2280
      End
      Begin VB.Label LblObjet 
         AutoSize        =   -1  'True
         Caption         =   "Nombre d'objet à vendre"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   200
         TabIndex        =   6
         Top             =   1400
         Width           =   2175
      End
      Begin VB.Label LblObjet 
         AutoSize        =   -1  'True
         Caption         =   "Prix de revente d'un objet"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   200
         TabIndex        =   4
         Top             =   900
         Width           =   2250
      End
      Begin VB.Label LblObjet 
         AutoSize        =   -1  'True
         Caption         =   "Prix neuf de l'objet : "
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   200
         TabIndex        =   3
         Top             =   400
         Width           =   1815
      End
   End
   Begin VB.CommandButton btnAnnuler 
      Caption         =   "Annuler"
      Height          =   615
      Left            =   3720
      TabIndex        =   1
      Top             =   2400
      Width           =   1095
   End
   Begin VB.CommandButton btnVendre 
      Caption         =   "Vendre"
      Height          =   495
      Left            =   2400
      TabIndex        =   0
      Top             =   2520
      Width           =   1095
   End
End
Attribute VB_Name = "frmVendre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public m_PrixUnitaire As Double
Public m_NombreObjet As Integer
Public m_NomObjet As String

Option Explicit

Private Sub btnAnnuler_Click()

  If g_ModeDebug = vbUnchecked Then On Error GoTo btnAnnuler_Click_Error

  Unload Me


btnAnnuler_Click_Exit:
  Exit Sub

btnAnnuler_Click_Error:
  MsgBox "L'erreur suivante s'est produite : " & Err.Description & vbCrLf & vbCrLf & " numéro d'erreur (" & Err.Number & " )", vbCritical, g_strNomApplication
  Print #g_FileDebug, Format(Now, "dd/MM/yyyy hh:mm:ss") & " Erreur n° : (" & Err.Number & ") " & Err.Description & " Procedure : btnAnnuler_Click  Module : Frmvendre.frm"
  Resume btnAnnuler_Click_Exit
End Sub

Private Sub btnSupprimer_Click()
frmGestPerso.vsEquip.RemoveItem
Unload Me
End Sub

Private Sub btnVendre_Click()

  If g_ModeDebug = vbUnchecked Then On Error GoTo btnVendre_Click_Error
  
  If Val(fldnPrixTotal) = 0 Then
    frmGestPerso.vsEquip.RemoveItem
    Unload Me
  End If
  
  If Val(fldnNombreObjet) = m_NombreObjet Then
    frmGestPerso.fldnTotalPo = Val(frmGestPerso.fldnTotalPo) + Val(fldnPrixTotal)
    frmGestPerso.vsEquip.RemoveItem
    Unload Me
  Else
    frmGestPerso.vsEquip.Cell(flexcpText, , L3_COL_EQUIP_MULTIPLE) = Val(m_NombreObjet) - Val(fldnNombreObjet)
    frmGestPerso.fldnTotalPo = Val(frmGestPerso.fldnTotalPo) + Val(fldnPrixTotal)
    Unload Me
  End If

btnVendre_Click_Exit:
  Exit Sub

btnVendre_Click_Error:
  MsgBox "L'erreur suivante s'est produite : " & Err.Description & vbCrLf & vbCrLf & " numéro d'erreur (" & Err.Number & " )", vbCritical, g_strNomApplication
  Print #g_FileDebug, Format(Now, "dd/MM/yyyy hh:mm:ss") & " Erreur n° : (" & Err.Number & ") " & Err.Description & " Procedure : btnVendre_Click  Module : Frmvendre.frm"
  Resume btnVendre_Click_Exit
End Sub

Private Sub fldnNombreObjet_Change()

  If fldnNombreObjet <> "" Then
    If Not IsNumeric(fldnNombreObjet) Then
      Screen.MousePointer = vbNormal
      MsgBox "Le nombre d'objet doit être entrée en tant que valeur numérique", vbExclamation, Me.name
      fldnNombreObjet = m_NombreObjet
      fldnNombreObjet.SetFocus
      Exit Sub
    End If
  Else
    Screen.MousePointer = vbNormal
    MsgBox "Le nombre d'objet doit être entrée en tant que valeur numérique", vbExclamation, Me.name
    fldnNombreObjet = m_NombreObjet
    fldnNombreObjet.SetFocus
    Exit Sub
  End If
  
  If Val(fldnNombreObjet) > m_NombreObjet Then
    Screen.MousePointer = vbNormal
    MsgBox "Le nombre d'objet doit être plus petit que " & m_NombreObjet, vbExclamation, Me.name
    fldnNombreObjet = m_NombreObjet
    fldnNombreObjet.SetFocus
    Exit Sub
  End If
  
  If Val(fldnNombreObjet) < 1 Then
    Screen.MousePointer = vbNormal
    MsgBox "Le nombre d'objet doit être plus grand que 1", vbExclamation, Me.name
    fldnNombreObjet = m_NombreObjet
    fldnNombreObjet.SetFocus
    Exit Sub
  End If
  Calcule
End Sub

Private Sub fldnPrixTotal_Change()
  
  If fldnPrixTotal <> "" Then
    If Not IsNumeric(fldnPrixTotal) Then
      Screen.MousePointer = vbNormal
      MsgBox "Le prix total doit être entrée en tant que valeur numérique", vbExclamation, Me.name
      fldnPrixTotal = fldnPrixUnitaire * fldnNombreObjet
      fldnPrixTotal.SetFocus
      Exit Sub
    End If
  Else
    Screen.MousePointer = vbNormal
    MsgBox "Le prix total doit être entrée en tant que valeur numérique", vbExclamation, Me.name
    fldnPrixTotal = fldnPrixUnitaire * fldnNombreObjet
    fldnPrixTotal.SetFocus
    Exit Sub
  End If
  
  If Val(fldnPrixTotal) < 0 Then
    Screen.MousePointer = vbNormal
    MsgBox "Le prix total doit être plus grand que 0", vbExclamation, Me.name
    fldnPrixTotal = fldnPrixUnitaire * fldnNombreObjet
    fldnPrixTotal.SetFocus
    Exit Sub
  End If
  
  fldnPrixUnitaire = fldnPrixTotal / fldnNombreObjet
  
End Sub

Private Sub fldnPrixUnitaire_Change()

  If fldnPrixUnitaire <> "" Then
    If Not IsNumeric(fldnPrixUnitaire) Then
      Screen.MousePointer = vbNormal
      MsgBox "Le prix unitaire doit être entrée en tant que valeur numérique", vbExclamation, Me.name
      fldnPrixUnitaire = m_PrixUnitaire / 2
      fldnPrixUnitaire.SetFocus
      Exit Sub
    End If
  Else
    Screen.MousePointer = vbNormal
    MsgBox "Le prix unitaire doit être entrée en tant que valeur numérique", vbExclamation, Me.name
    fldnPrixUnitaire = m_PrixUnitaire / 2
    fldnPrixUnitaire.SetFocus
    Exit Sub
  End If
  
  If Val(fldnPrixUnitaire) < 0 Then
    Screen.MousePointer = vbNormal
    MsgBox "Le prix unitaire doit être plus grand que 0", vbExclamation, Me.name
    fldnNombreObjet = m_PrixUnitaire / 2
    fldnNombreObjet.SetFocus
    Exit Sub
  End If
  Calcule

End Sub

Private Sub Form_Load()

frmObjet.Caption = m_NomObjet
LblObjet(0).Caption = "Prix neuf de l'objet : " & m_PrixUnitaire & " Po"
fldnNombreObjet = m_NombreObjet
fldnPrixUnitaire = m_PrixUnitaire / 2
fldnPrixTotal = fldnPrixUnitaire * fldnNombreObjet
End Sub
Private Sub Calcule()
Dim NombreObjet As Integer

NombreObjet = Val(fldnNombreObjet)
fldnNombreObjet = NombreObjet
fldnPrixTotal = fldnPrixUnitaire * fldnNombreObjet
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyEscape Then Unload Me
End Sub

