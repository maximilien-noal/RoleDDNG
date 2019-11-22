VERSION 5.00
Begin VB.Form frmAjouter 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Choix de la méthode"
   ClientHeight    =   1140
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   2535
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1140
   ScaleWidth      =   2535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton btnAnnuler 
      Cancel          =   -1  'True
      Caption         =   "Annuler"
      Height          =   330
      Left            =   1620
      TabIndex        =   3
      Top             =   720
      Width           =   855
   End
   Begin VB.CommandButton btnOk 
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   330
      Left            =   720
      TabIndex        =   2
      Top             =   720
      Width           =   855
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Assistant"
      Height          =   255
      Left            =   1320
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Manuelle"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "frmAjouter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public m_type As Integer
Private Sub btnAnnuler_Click()
  Unload Me
End Sub

Private Sub btnOk_Click()
  If Option1 = True Then
    m_type = 1
  ElseIf Option2 = True Then
    m_type = 2
  End If
  Unload Me
End Sub
Private Sub Form_Load()
  m_type = 0
End Sub

