VERSION 5.00
Begin VB.Form frmLanceurDes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Lanceur de dès"
   ClientHeight    =   2100
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4605
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2100
   ScaleWidth      =   4605
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox fldnResultat 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1215
      TabIndex        =   7
      Top             =   1650
      Width           =   690
   End
   Begin VB.TextBox fldstrDetailResultat 
      Height          =   375
      Left            =   135
      TabIndex        =   5
      Top             =   1170
      Width           =   4470
   End
   Begin VB.CommandButton btnGenerer 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   870
      Left            =   2295
      Picture         =   "frmLanceurDes.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   225
      Width           =   2265
   End
   Begin VB.ComboBox CmbTypeDes 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      ItemData        =   "frmLanceurDes.frx":61FA
      Left            =   1485
      List            =   "frmLanceurDes.frx":6213
      TabIndex        =   3
      Text            =   "6"
      Top             =   668
      Width           =   690
   End
   Begin VB.TextBox fldnNombreDes 
      Height          =   285
      Left            =   1485
      TabIndex        =   1
      Text            =   "1"
      Top             =   270
      Width           =   690
   End
   Begin VB.Label Label3 
      Caption         =   "Résultat"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   225
      TabIndex        =   6
      Top             =   1650
      Width           =   960
   End
   Begin VB.Label Label2 
      Caption         =   "Type de dès"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   180
      TabIndex        =   2
      Top             =   720
      Width           =   1140
   End
   Begin VB.Label Label1 
      Caption         =   "Nombre de dès"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   180
      TabIndex        =   0
      Top             =   292
      Width           =   1230
   End
End
Attribute VB_Name = "frmLanceurDes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnGenerer_Click()
Dim i As Integer, n As Integer
Dim total As Long

total = 0
fldstrDetailResultat = ""
For i = 1 To Val(fldnNombreDes)
  n = Int(Rnd(1000) * Val(CmbTypeDes)) + 1
  total = total + n
  fldstrDetailResultat = fldstrDetailResultat & n & " "
Next i
fldnResultat = total
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

  If KeyCode = vbKeyReturn Then
    btnGenerer_Click
  ElseIf KeyCode = vbKeyEscape Then
    Unload Me
  End If
End Sub

