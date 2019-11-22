VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmSousCarac 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Sous Caractèristiques"
   ClientHeight    =   2445
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3045
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2445
   ScaleWidth      =   3045
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton btnAnnuler 
      Cancel          =   -1  'True
      Caption         =   "&Annuler"
      Height          =   285
      Left            =   2025
      TabIndex        =   25
      Top             =   2115
      Width           =   960
   End
   Begin VB.CommandButton btnOk 
      Caption         =   "&Ok"
      Height          =   285
      Left            =   1080
      TabIndex        =   24
      Top             =   2115
      Width           =   870
   End
   Begin VB.TextBox fldnSousCarac 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      Height          =   285
      Index           =   11
      Left            =   2580
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   11
      Tag             =   "numeric"
      Top             =   1710
      Width           =   375
   End
   Begin VB.TextBox fldnSousCarac 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      Height          =   285
      Index           =   10
      Left            =   1815
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   10
      Tag             =   "numeric"
      Top             =   1710
      Width           =   375
   End
   Begin VB.TextBox fldnSousCarac 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      Height          =   285
      Index           =   9
      Left            =   2580
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   9
      Tag             =   "numeric"
      Top             =   1395
      Width           =   375
   End
   Begin VB.TextBox fldnSousCarac 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      Height          =   285
      Index           =   8
      Left            =   1815
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   8
      Tag             =   "numeric"
      Top             =   1395
      Width           =   375
   End
   Begin VB.TextBox fldnSousCarac 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      Height          =   285
      Index           =   7
      Left            =   2580
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   7
      Tag             =   "numeric"
      Top             =   1080
      Width           =   375
   End
   Begin VB.TextBox fldnSousCarac 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      Height          =   285
      Index           =   6
      Left            =   1815
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   6
      Tag             =   "numeric"
      Top             =   1080
      Width           =   375
   End
   Begin VB.TextBox fldnSousCarac 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      Height          =   285
      Index           =   5
      Left            =   2580
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   5
      Tag             =   "numeric"
      Top             =   750
      Width           =   375
   End
   Begin VB.TextBox fldnSousCarac 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      Height          =   285
      Index           =   4
      Left            =   1815
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   4
      Tag             =   "numeric"
      Top             =   750
      Width           =   375
   End
   Begin VB.TextBox fldnSousCarac 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      Height          =   285
      Index           =   3
      Left            =   2580
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   3
      Tag             =   "numeric"
      Top             =   435
      Width           =   375
   End
   Begin VB.TextBox fldnSousCarac 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      Height          =   285
      Index           =   2
      Left            =   1815
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   2
      Tag             =   "numeric"
      Top             =   435
      Width           =   375
   End
   Begin VB.TextBox fldnSousCarac 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      Height          =   285
      Index           =   1
      Left            =   2580
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   1
      Tag             =   "numeric"
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox fldnSousCarac 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      Height          =   285
      Index           =   0
      Left            =   1815
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   0
      Tag             =   "numeric"
      Top             =   120
      Width           =   375
   End
   Begin MSForms.SpinButton SpinComp 
      Height          =   255
      Index           =   5
      Left            =   2220
      TabIndex        =   23
      Top             =   1755
      Width           =   360
      Size            =   "635;459"
   End
   Begin MSForms.SpinButton SpinComp 
      Height          =   255
      Index           =   4
      Left            =   2220
      TabIndex        =   22
      Top             =   1440
      Width           =   360
      Size            =   "635;459"
   End
   Begin MSForms.SpinButton SpinComp 
      Height          =   255
      Index           =   3
      Left            =   2220
      TabIndex        =   21
      Top             =   1125
      Width           =   360
      Size            =   "635;459"
   End
   Begin MSForms.SpinButton SpinComp 
      Height          =   255
      Index           =   2
      Left            =   2220
      TabIndex        =   20
      Top             =   810
      Width           =   360
      Size            =   "635;459"
   End
   Begin MSForms.SpinButton SpinComp 
      Height          =   255
      Index           =   1
      Left            =   2220
      TabIndex        =   19
      Top             =   480
      Width           =   360
      Size            =   "635;459"
   End
   Begin MSForms.SpinButton SpinComp 
      Height          =   255
      Index           =   0
      Left            =   2220
      TabIndex        =   18
      Top             =   135
      Width           =   360
      Size            =   "635;459"
   End
   Begin VB.Label Label15 
      Caption         =   "Magnétisme/Charme:"
      Height          =   240
      Left            =   105
      TabIndex        =   17
      Top             =   1710
      Width           =   1680
   End
   Begin VB.Label Label14 
      Caption         =   "Intuition/volonté:"
      Height          =   240
      Left            =   105
      TabIndex        =   16
      Top             =   1395
      Width           =   1680
   End
   Begin VB.Label Label13 
      Caption         =   "Intellect/érudition:"
      Height          =   240
      Left            =   105
      TabIndex        =   15
      Top             =   1080
      Width           =   1680
   End
   Begin VB.Label Label11 
      Caption         =   "Resistance/Vitalité:"
      Height          =   240
      Left            =   105
      TabIndex        =   14
      Top             =   750
      Width           =   1680
   End
   Begin VB.Label Label10 
      Caption         =   "Precision/equilibre:"
      Height          =   240
      Left            =   105
      TabIndex        =   13
      Top             =   435
      Width           =   1680
   End
   Begin VB.Label Label32 
      Caption         =   "Endurance/Puissance:"
      Height          =   240
      Left            =   105
      TabIndex        =   12
      Top             =   120
      Width           =   1680
   End
End
Attribute VB_Name = "frmSousCarac"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const COL_NIVEAU = 0
Const COL_CLASSE = 1
Const COL_PV = 2
Const COL_FOR = 3
Const COL_DEX = 4
Const COL_CON = 5
Const COL_INT = 6
Const COL_SAG = 7
Const COL_CHA = 8
Const COL_DON_COMP = 9
Const COL_SOUS_CARAC = 10
Const COL_COMPETENCE = 11
Const COL_DON_1 = 12
Const COL_DON_2 = 13
Const COL_DON_3 = 14
Const COL_DON_4 = 15
Const COL_DON_5 = 16
Const COL_DON_6 = 17
Const COL_DON_7 = 18
Const COL_LIB_1 = 19
Const COL_LIB_2 = 20
Const COL_LIB_3 = 21
Const COL_LIB_4 = 22
Const COL_LIB_5 = 23
Const COL_LIB_6 = 24
Const COL_LIB_7 = 25
Const COL_ENDURANCE = 26
Const COL_PUISSANCE = 27
Const COL_PRECISION = 28
Const COL_EQUILIBRE = 29
Const COL_RESISTANCE = 30
Const COL_VITALITE = 31
Const COL_INTELLECT = 32
Const COL_ERUDITION = 33
Const COL_INTUITION = 34
Const COL_VOLONTE = 35
Const COL_MAGNETISME = 36
Const COL_CHARME = 37

Private Sub btnAnnuler_Click()
  Unload Me
End Sub

Private Sub SpinComp_SpinDown(Index As Integer)
  fldnSousCarac(Index * 2 + 0) = fldnSousCarac(Index * 2 + 0) - 1
  fldnSousCarac(Index * 2 + 1) = fldnSousCarac(Index * 2 + 1) + 1

End Sub

Private Sub SpinComp_spinup(Index As Integer)
  
  fldnSousCarac(Index * 2 + 0) = fldnSousCarac(Index * 2 + 0) + 1
  fldnSousCarac(Index * 2 + 1) = fldnSousCarac(Index * 2 + 1) - 1
End Sub
