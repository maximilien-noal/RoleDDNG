VERSION 5.00
Begin VB.Form frmDescriptionRaces 
   Caption         =   "Description des races"
   ClientHeight    =   4230
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10695
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
   MDIChild        =   -1  'True
   ScaleHeight     =   4230
   ScaleWidth      =   10695
   Begin VB.CommandButton btnEnregistrer 
      Caption         =   "&Enregistrer"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5880
      TabIndex        =   2
      Top             =   120
      Width           =   1275
   End
   Begin VB.ComboBox CmbRace 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   90
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   90
      Width           =   3615
   End
   Begin VB.TextBox fldstrDescriptionRaces 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3660
      Left            =   45
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   585
      Width           =   10650
   End
End
Attribute VB_Name = "frmDescriptionRaces"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public m_race As String

Private Sub btnEnregistrer_Click()

If CmbRace <> "" And Trim(fldstrDescriptionRaces) <> "" Then
  g_bdRoles.Execute "update race set Description='" & FormatStrSQL(fldstrDescriptionRaces) & "' where race='" & FormatStrSQL(CmbRace) & "'", dbFailOnError
End If

End Sub

Private Sub CmbRace_click()
Dim rst As Recordset

Set rst = g_bdRoles.OpenRecordset("select * from race where race='" & FormatStrSQL(CmbRace) & "'", dbOpenDynaset)
fldstrDescriptionRaces = IIf(IsNull(rst!Description), "", rst!Description)
rst.Close
End Sub

Private Sub Form_Load()
Dim rst As Recordset

fldstrDescriptionRaces.Height = frmDescriptionRaces.Height - 1025
fldstrDescriptionRaces.Width = frmDescriptionRaces.Width - 165
  
  CmbRace.Clear
  Set rst = g_bdRoles.OpenRecordset("select * from race order by race", dbOpenDynaset)
  Do While Not rst.EOF
    CmbRace.AddItem rst!race
    CmbRace.ItemData(CmbRace.ListCount - 1) = rst![n°]
    rst.MoveNext
  Loop
  rst.Close
  CmbRace.ListIndex = 0
  If m_race <> "" Then
    CmbRace = m_race
    Set rst = g_bdRoles.OpenRecordset("select * from race where race='" & FormatStrSQL(CmbRace) & "'", dbOpenDynaset)
    fldstrDescriptionRaces = IIf(IsNull(rst!Description), "", rst!Description)
    rst.Close
  End If
End Sub

Private Sub Form_Resize()
fldstrDescriptionRaces.Height = maxi_(1, frmDescriptionRaces.Height - 1025)
fldstrDescriptionRaces.Width = maxi_(1, frmDescriptionRaces.Width - 165)
End Sub
