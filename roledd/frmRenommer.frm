VERSION 5.00
Begin VB.Form frmRenommer 
   Caption         =   "Renommer"
   ClientHeight    =   945
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5445
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   945
   ScaleWidth      =   5445
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton btnAnnuler 
      Cancel          =   -1  'True
      Caption         =   "&Annuler"
      Height          =   330
      Left            =   4230
      TabIndex        =   2
      Top             =   540
      Width           =   1140
   End
   Begin VB.CommandButton btnOk 
      Caption         =   "&Ok"
      Enabled         =   0   'False
      Height          =   330
      Left            =   3015
      TabIndex        =   1
      Top             =   540
      Width           =   1140
   End
   Begin VB.TextBox fldstrNouveauNom 
      Height          =   285
      Left            =   1395
      MaxLength       =   50
      TabIndex        =   0
      Top             =   135
      Width           =   3930
   End
   Begin VB.Label Label1 
      Caption         =   "Nouveau nom:"
      Height          =   240
      Left            =   135
      TabIndex        =   3
      Top             =   180
      Width           =   1185
   End
End
Attribute VB_Name = "frmRenommer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public m_strnouveaunom As String

Private Sub btnAnnuler_Click()
  Unload Me
End Sub

Private Sub btnOk_Click()

  If Trim(fldstrNouveauNom) = "" Then
    fldstrNouveauNom.SetFocus
    Exit Sub
  End If
  m_strnouveaunom = fldstrNouveauNom
  Unload Me
End Sub

Private Sub fldstrNouveauNom_Change()
  btnOk.Enabled = Len(Trim(fldstrNouveauNom)) > 0
End Sub

Private Sub fldstrNouveauNom_GotFocus()
  Selec
End Sub

Private Sub Form_Load()
  m_strnouveaunom = ""
End Sub
