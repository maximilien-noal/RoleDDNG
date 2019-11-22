VERSION 5.00
Begin VB.Form frmDescriptifCompetence 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Description des compétences"
   ClientHeight    =   11865
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12945
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
   ScaleHeight     =   11865
   ScaleWidth      =   12945
   Begin VB.TextBox fldstrinne 
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
      Height          =   1065
      Left            =   1725
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   15
      Top             =   10425
      Width           =   11205
   End
   Begin VB.TextBox fldstrRestriction 
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
      Height          =   1065
      Left            =   1725
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   14
      Top             =   9300
      Width           =   11205
   End
   Begin VB.TextBox fldstrSynergie 
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
      Height          =   1065
      Left            =   1725
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   13
      Top             =   8220
      Width           =   11205
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
      Height          =   1065
      Left            =   1725
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   12
      Top             =   7140
      Width           =   11205
   End
   Begin VB.TextBox fldstrNouvelle 
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
      Height          =   1065
      Left            =   1725
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   11
      Top             =   6060
      Width           =   11205
   End
   Begin VB.TextBox fldstrAction 
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
      Height          =   1065
      Left            =   1725
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   10
      Top             =   4980
      Width           =   11205
   End
   Begin VB.CommandButton btnFermer 
      Cancel          =   -1  'True
      Caption         =   "&Fermer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   11880
      TabIndex        =   3
      Top             =   11520
      Width           =   1050
   End
   Begin VB.TextBox fldstrTest 
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
      Height          =   4560
      Left            =   1725
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   405
      Width           =   11205
   End
   Begin VB.Label Label7 
      Caption         =   "Test inné"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   80
      TabIndex        =   9
      Top             =   10755
      Width           =   1680
   End
   Begin VB.Label Label6 
      Caption         =   "Restrictions"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   80
      TabIndex        =   8
      Top             =   9675
      Width           =   1680
   End
   Begin VB.Label Label5 
      Caption         =   "Synergie"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   80
      TabIndex        =   7
      Top             =   8595
      Width           =   1680
   End
   Begin VB.Label Label4 
      Caption         =   "Spécial"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   80
      TabIndex        =   6
      Top             =   7515
      Width           =   1680
   End
   Begin VB.Label Label3 
      Caption         =   "Nouvelles tentatives :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   80
      TabIndex        =   5
      Top             =   6435
      Width           =   1680
   End
   Begin VB.Label Label1 
      Caption         =   "Action"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   80
      TabIndex        =   4
      Top             =   5355
      Width           =   1680
   End
   Begin VB.Label labnom 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   330
      Left            =   45
      TabIndex        =   2
      Top             =   45
      Width           =   12885
   End
   Begin VB.Label Label2 
      Caption         =   "Test de compétence"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   75
      TabIndex        =   1
      Top             =   2640
      Width           =   1680
   End
End
Attribute VB_Name = "frmDescriptifCompetence"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public m_competence As String
Private Sub btnFermer_Click()
  Unload Me
End Sub
Private Sub Form_Load()
Dim rst As Recordset

  labnom = m_competence
  Set rst = g_bdRoles.OpenRecordset("competence", dbOpenTable)
  rst.Index = "PrimaryKey"
  rst.Seek "=", m_competence
  If Not rst.NoMatch Then
    labnom = labnom & " (" & rst!carac1
    If rst!innee = "N" Then
      labnom = labnom & "; formation nécessaire"
    End If
    If rst!PenaliteArmure Then
      labnom = labnom & "; malus d'armure aux tests"
    End If
    labnom = labnom & ")"
    If Not IsNull(rst!Test) Then
      fldstrTest = rst!Test
    End If
    If Not IsNull(rst!Action) Then
      fldstrAction = rst!Action
    End If
    If Not IsNull(rst!Nouvelles) Then
      fldstrNouvelle = rst!Nouvelles
    End If
    If Not IsNull(rst!Spécial) Then
      fldstrSpecial = rst!Spécial
    End If
    If Not IsNull(rst!synergie) Then
      fldstrSynergie = rst!synergie
    End If
    If Not IsNull(rst!restrictions) Then
      fldstrRestriction = rst!restrictions
    End If
    If Not IsNull(rst!testinne) Then
      fldstrinne = rst!testinne
    End If
  End If
  rst.Close
End Sub

