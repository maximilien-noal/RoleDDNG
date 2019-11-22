VERSION 5.00
Begin VB.Form frmApropos 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "A propos de RoleDD"
   ClientHeight    =   3825
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6165
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3825
   ScaleWidth      =   6165
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      Height          =   975
      Left            =   4560
      Picture         =   "FrmApropos.frx":0000
      ScaleHeight     =   915
      ScaleWidth      =   1035
      TabIndex        =   5
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox fldStrVersion 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   975
      Left            =   1710
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   2205
      Width           =   3855
   End
   Begin VB.CommandButton btnOk 
      Cancel          =   -1  'True
      Caption         =   "&Ok"
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
      Left            =   4545
      TabIndex        =   1
      Top             =   3375
      Width           =   1005
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Versions :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   2
      Left            =   240
      TabIndex        =   4
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "http://bonnarien.dyndns.org"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   465
      Index           =   1
      Left            =   180
      TabIndex        =   3
      Top             =   1620
      Width           =   4380
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Gestion des feuilles de personnages pour Donjons et Dragons 3.5"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1230
      Index           =   0
      Left            =   135
      TabIndex        =   0
      Top             =   135
      Width           =   4380
   End
End
Attribute VB_Name = "frmApropos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub btnOk_Click()
  Unload Me
End Sub

Private Sub Form_Load()
  fldStrVersion = "Version du logiciel :        " & vbTab & App.Major & vbCrLf & "Paramètres roles.mdb : " & vbTab & App.Minor & vbCrLf & "Nombre de compilations :     " & vbTab & App.Revision
End Sub

Private Sub Label2_Click(Index As Integer)
  If Index = 1 Then
    Screen.MousePointer = vbHourglass
    ShellExecute Me.hwnd, "open", "http://bonnarien.dyndns.org", "", "", 0
    Screen.MousePointer = vbNormal
  End If
End Sub

