VERSION 5.00
Begin VB.Form frmImpression 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Type d'impression"
   ClientHeight    =   3570
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3000
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3570
   ScaleWidth      =   3000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox ChkImprim 
      Caption         =   "Uniquement les sort quotidiens"
      Height          =   375
      Index           =   4
      Left            =   120
      TabIndex        =   10
      Top             =   2430
      Width           =   2820
   End
   Begin VB.CheckBox ChkImprim 
      Caption         =   "Sorts"
      Height          =   375
      Index           =   3
      Left            =   1560
      TabIndex        =   9
      Top             =   1900
      Width           =   1335
   End
   Begin VB.CheckBox ChkImprim 
      Caption         =   "Dons"
      Height          =   375
      Index           =   2
      Left            =   120
      TabIndex        =   8
      Top             =   1900
      Width           =   1335
   End
   Begin VB.CheckBox ChkImprim 
      Caption         =   "Histoire"
      Height          =   375
      Index           =   1
      Left            =   1560
      TabIndex        =   7
      Top             =   1400
      Width           =   1335
   End
   Begin VB.CheckBox ChkImprim 
      Caption         =   "background"
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   1400
      Width           =   1335
   End
   Begin VB.OptionButton Option4 
      Caption         =   "Txt"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   260
      Left            =   120
      TabIndex        =   5
      Top             =   720
      Width           =   1200
   End
   Begin VB.OptionButton Option3 
      Caption         =   "WordPerso"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   260
      Left            =   1300
      TabIndex        =   4
      Top             =   240
      Width           =   1200
   End
   Begin VB.CommandButton btnAnnuler 
      Cancel          =   -1  'True
      Caption         =   "Annuler"
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
      Left            =   1560
      TabIndex        =   3
      Top             =   3105
      Width           =   855
   End
   Begin VB.CommandButton btnOk 
      Caption         =   "Ok"
      Default         =   -1  'True
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
      Left            =   600
      TabIndex        =   2
      Top             =   3105
      Width           =   855
   End
   Begin VB.OptionButton Option2 
      Caption         =   "HTML"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   260
      Left            =   1300
      TabIndex        =   1
      Top             =   720
      Width           =   1200
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Word"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   260
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1200
   End
End
Attribute VB_Name = "frmImpression"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public m_bBackgroud As Boolean
Public m_bHistoire As Boolean
Public m_bDons As Boolean
Public m_bSorts As Boolean
Public m_type As Integer
Public m_bSortsQuotidiens As Boolean

Private Sub btnAnnuler_Click()
  Unload Me
End Sub

Private Sub btnOk_Click()
  If Option1 = True Then
    m_type = 1
  End If
  If Option2 = True Then
    m_type = 2
  End If
  If Option3 = True Then
    m_type = 3
  End If
  If Option4 = True Then
    m_type = 4
  End If
  
  m_bBackgroud = ChkImprim(0)
  m_bHistoire = ChkImprim(1)
  m_bDons = ChkImprim(2)
  m_bSorts = ChkImprim(3)
  m_bSortsQuotidiens = ChkImprim(4)
  Unload Me


End Sub

Private Sub ChkImprim_Click(Index As Integer)
If Index = 3 Then
  If ChkImprim(3) Then
    ChkImprim(4).Enabled = True
  Else
    ChkImprim(4).Enabled = False
  End If
End If
End Sub

Private Sub Form_Load()
  m_type = 0
  ChkImprim(4).Enabled = False
End Sub

