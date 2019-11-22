VERSION 5.00
Begin VB.Form frmBackHist 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Background et Historique"
   ClientHeight    =   7455
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10185
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7455
   ScaleWidth      =   10185
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton btnFermer 
      Cancel          =   -1  'True
      Caption         =   "&Fermer"
      Height          =   285
      Left            =   8865
      TabIndex        =   5
      Top             =   7110
      Width           =   1275
   End
   Begin VB.CommandButton btnEnregistrer 
      Caption         =   "&Enregistrer"
      Height          =   285
      Left            =   7515
      TabIndex        =   4
      Top             =   7110
      Width           =   1275
   End
   Begin VB.TextBox fldstrhistorique 
      Height          =   3210
      Left            =   90
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   3825
      Width           =   10050
   End
   Begin VB.TextBox fldstrbackground 
      Height          =   3165
      Left            =   90
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   315
      Width           =   10050
   End
   Begin VB.Label Label2 
      Caption         =   "Historique du personnage"
      Height          =   330
      Left            =   240
      TabIndex        =   3
      Top             =   3510
      Width           =   2265
   End
   Begin VB.Label Label1 
      Caption         =   "Background du personnage"
      Height          =   285
      Left            =   225
      TabIndex        =   2
      Top             =   45
      Width           =   2400
   End
End
Attribute VB_Name = "frmBackHist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub btnEnregistrer_Click()
Dim rst As Recordset
  
  If frmGestPerso.fldstrnom <> "" Then
    Set rst = g_bdPerso.OpenRecordset("select * from personnage where nom='" & FormatStrSQL(frmGestPerso.fldstrnom) & "'", dbOpenDynaset)
    If Not rst.EOF Then
      rst.Edit
      rst!Background = IIf(Trim(fldstrbackground) = "", Null, fldstrbackground)
      rst!histoire = IIf(Trim(fldstrhistorique) = "", Null, fldstrhistorique)
      rst.Update
    End If
    rst.Close
  End If
  Unload Me
End Sub

Private Sub btnFermer_Click()
  Unload Me

End Sub

Private Sub Form_Load()
Dim rst As Recordset

  If frmGestPerso.fldstrnom <> "" Then
    Set rst = g_bdPerso.OpenRecordset("select background,histoire from personnage where nom='" & FormatStrSQL(frmGestPerso.fldstrnom) & "'", dbOpenDynaset)
    If Not rst.EOF Then
      fldstrbackground = IIf(IsNull(rst!Background), "", rst!Background)
      fldstrhistorique = IIf(IsNull(rst!histoire), "", rst!histoire)
    End If
    rst.Close
  End If
End Sub

