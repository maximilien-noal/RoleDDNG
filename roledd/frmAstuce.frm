VERSION 5.00
Begin VB.Form frmAstuce 
   Caption         =   "Astuce du jour"
   ClientHeight    =   4050
   ClientLeft      =   2370
   ClientTop       =   2400
   ClientWidth     =   6585
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
   ScaleHeight     =   4050
   ScaleWidth      =   6585
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.CheckBox chkLoadTipsAtStartup 
      Caption         =   "&Afficher les astuces au d�marrage"
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Top             =   3600
      Width           =   2895
   End
   Begin VB.CommandButton cmdNextTip 
      Caption         =   "Astuce &suivante"
      Height          =   375
      Left            =   5160
      TabIndex        =   2
      Top             =   600
      Width           =   1335
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   3435
      Left            =   120
      Picture         =   "frmAstuce.frx":0000
      ScaleHeight     =   3375
      ScaleWidth      =   4875
      TabIndex        =   1
      Top             =   120
      Width           =   4935
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Savez-vous..."
         Height          =   255
         Left            =   540
         TabIndex        =   5
         Top             =   180
         Width           =   2655
      End
      Begin VB.Label lblTipText 
         BackColor       =   &H00FFFFFF&
         Height          =   2475
         Left            =   120
         TabIndex        =   4
         Top             =   840
         Width           =   4575
      End
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   5160
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "frmAstuce"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Base de donn�es des astuces en m�moire.
Dim Tips As New Collection

' Nom du fichier des astuces.
Const TIP_FILE = "ASTUCE.TXT"

' Index dans la collection de l'astuce actuellement affich�s.
Dim CurrentTip As Long


Private Sub DoNextTip()

    ' S�lectionne une astuce de fa�on al�atoire.
   ' CurrentTip = Int((Tips.Count * Rnd) + 1)
    
    ' Ou, vous pouvez effectuer une op�ration cyclique en suivant l'ordre des astuces.

    CurrentTip = CurrentTip + 1
    If Tips.Count < CurrentTip Then
        CurrentTip = 1
    End If
    
    ' L'affiche.
    frmAstuce.DisplayCurrentTip
    
End Sub

Function LoadTips(sFile As String) As Boolean
    Dim NextTip As String   ' Chaque astuce est lu � partir du fichier.
    Dim InFile As Integer   ' Descripteur du fichier.
    
    ' Obtient le prochain descripteur de fichier libre.
    InFile = FreeFile
    
    ' S'assure qu'un fichier est sp�cifi�.
    If sFile = "" Then
        LoadTips = False
        Exit Function
    End If
    
    ' S'assure que le fichier existe avant d'essayer de l'ouvrir.
    If Dir(sFile) = "" Then
        LoadTips = False
        Exit Function
    End If
    
    ' Remplit la collection � partir d'un fichier texte.
    Open sFile For Input As InFile
    While Not EOF(InFile)
        Line Input #InFile, NextTip
        Tips.Add NextTip
    Wend
    Close InFile

    ' Affiche une astuce de fa�on al�atoire.
    DoNextTip
    
    LoadTips = True
    
End Function

Private Sub chkLoadTipsAtStartup_Click()
    ' Enregistre s'il faut ou non afficher cette feuille au d�marrage.
    SaveSetting App.EXEName, "Options", "Show Tips at Startup", chkLoadTipsAtStartup.Value
End Sub

Private Sub cmdNextTip_Click()
    DoNextTip
End Sub

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub Form_Load()
   
        
    ' D�finit la case � cocher, ce qui va forcer la r��criture de la valeur dans la base de registres.
    Me.chkLoadTipsAtStartup.Value = vbChecked
    
    ' Valeur de d�part pour nombre al�atoire.
    Randomize
    
    ' Lit le fichier des astuces et affiche une astuce de fa�on al�atoire.
    If LoadTips(App.Path & "\" & TIP_FILE) = False Then
        lblTipText.Caption = "Le fichier " & TIP_FILE & " n'a pas �t� trouv�? " & vbCrLf & vbCrLf & _
           "Cr�ez un fichier texte nomm� " & TIP_FILE & " en utilisant le Bloc-notes avec une astuce par ligne. " & _
           "Puis placez le dans le m�me dossier que l'application."
    End If

    
End Sub

Public Sub DisplayCurrentTip()
    If Tips.Count > 0 Then
        lblTipText.Caption = Tips.Item(CurrentTip)
    End If
End Sub
