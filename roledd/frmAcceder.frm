VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "VSFLEX7D.ocx"
Begin VB.Form frmAcceder 
   Caption         =   "Accéder à un personnage"
   ClientHeight    =   3840
   ClientLeft      =   6420
   ClientTop       =   4725
   ClientWidth     =   6660
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
   MDIChild        =   -1  'True
   ScaleHeight     =   3840
   ScaleWidth      =   6660
   Begin VSFlex7DAOCtl.vsFlexGrid vsnom 
      Height          =   3825
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6360
      _ExtentX        =   11218
      _ExtentY        =   6747
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   3
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   4
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   "Nom                           |Race                                  |Niv|Classe       "
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   1
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0   'False
      ShowComboButton =   -1  'True
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
   End
End
Attribute VB_Name = "frmAcceder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Sub RempPerso()
Dim rst As Recordset
Dim ClasseMax As String
Dim NivMax As Integer
Dim Niv_1 As Integer, Niv_2 As Integer, Niv_3 As Integer, Niv_4 As Integer, Niv_5 As Integer, Niv_6 As Integer, Niv_7 As Integer, Niv_8 As Integer

  If g_ModeDebug = vbUnchecked Then On Error GoTo RempPerso_Error

  vsnom.Rows = 1
  Set rst = g_bdPerso.OpenRecordset("select nom,race,niv_1,niv_2,niv_3,niv_4,niv_5,niv_6,niv_7,niv_8,classe_1,classe_2,classe_3,classe_4,classe_5,classe_6,classe_7,classe_8 from personnage where exclu=false order by nom", dbOpenDynaset)
  Do While Not rst.EOF
    Niv_1 = IIf(IsNull(rst!Niv_1), 0, rst!Niv_1)
    Niv_2 = IIf(IsNull(rst!Niv_2), 0, rst!Niv_2)
    Niv_3 = IIf(IsNull(rst!Niv_3), 0, rst!Niv_3)
    Niv_4 = IIf(IsNull(rst!Niv_4), 0, rst!Niv_4)
    Niv_5 = IIf(IsNull(rst!Niv_5), 0, rst!Niv_5)
    Niv_6 = IIf(IsNull(rst!Niv_6), 0, rst!Niv_6)
    Niv_7 = IIf(IsNull(rst!Niv_7), 0, rst!Niv_7)
    Niv_8 = IIf(IsNull(rst!Niv_8), 0, rst!Niv_8)
    ClasseMax = rst!Classe_1
    NivMax = rst!Niv_1
    
    If Niv_2 > NivMax Then
      ClasseMax = rst!Classe_2
      NivMax = Niv_2
    End If
    If Niv_3 > NivMax Then
      ClasseMax = rst!Classe_3
      NivMax = Niv_3
    End If
    If Niv_4 > NivMax Then
      ClasseMax = rst!Classe_4
      NivMax = Niv_4
    End If
    If Niv_5 > NivMax Then
      ClasseMax = rst!Classe_5
      NivMax = Niv_5
    End If
    If Niv_6 > NivMax Then
      ClasseMax = rst!Classe_6
      NivMax = Niv_6
    End If
    If Niv_7 > NivMax Then
      ClasseMax = rst!Classe_7
      NivMax = Niv_7
    End If
    If Niv_8 > NivMax Then
      ClasseMax = rst!Classe_8
      NivMax = Niv_8
    End If
    
    vsnom.AddItem rst!Nom & vbTab & rst!RACE & vbTab & Niv_1 + Niv_2 + Niv_3 + Niv_4 + Niv_5 + Niv_6 + Niv_7 + Niv_8 & vbTab & ClasseMax
    rst.MoveNext
  Loop
  rst.Close

RempPerso_Exit:
  Exit Sub

RempPerso_Error:
  MsgBox "L'erreur suivante s'est produite : " & Err.Description & vbCrLf & vbCrLf & " numéro d'erreur (" & Err.Number & " )", vbCritical, g_strNomApplication
  Print #g_FileDebug, Format(Now, "dd/MM/yyyy hh:mm:ss") & " Erreur n° : (" & Err.Number & ") " & Err.Description & " Procedure : RempPerso  Module : frmAcceder.frm"
  Resume RempPerso_Exit
End Sub
Private Sub Form_Load()

  If g_ModeDebug = vbUnchecked Then On Error GoTo Form_Load_Error
  frmAcceder.left = frmMain.Width - 7000
  frmAcceder.top = frmMain.Height - 5600
  
  frmAcceder.Height = 4320
  frmAcceder.Width = 6600
  vsnom.Height = 3825
  vsnom.Width = 6360
  RempPerso

Form_Load_Exit:
  Exit Sub

Form_Load_Error:
  MsgBox "L'erreur suivante s'est produite : " & Err.Description & vbCrLf & vbCrLf & " numéro d'erreur (" & Err.Number & " )", vbCritical, g_strNomApplication
  Print #g_FileDebug, Format(Now, "dd/MM/yyyy hh:mm:ss") & " Erreur n° : (" & Err.Number & ") " & Err.Description & " Procedure : Form_Load  Module : frmAcceder.frm"
  Resume Form_Load_Exit
End Sub

Private Sub Form_Resize()
  vsnom.Height = maxi_(1, frmAcceder.Height - 500)
  vsnom.Width = frmAcceder.Width - 250
End Sub

Private Sub vsnom_Click()
  AccedePerso
End Sub
Public Sub AccedePerso()

  If g_ModeDebug = vbUnchecked Then On Error GoTo AccedePerso_Error
  
  If vsnom.SelectedRows > 0 Then
    frmFichePerso.m_strnomPerso = vsnom.Cell(flexcpText, vsnom.SelectedRow(0), 0)
    frmFichePerso.Show
    frmFichePerso.ZOrder 0
    frmFichePerso.RetrouveFiche
    frmFichePerso.ValideEquipementTous (1)
  End If

AccedePerso_Exit:
  Exit Sub

AccedePerso_Error:
  MsgBox "L'erreur suivante s'est produite : " & Err.Description & vbCrLf & vbCrLf & " numéro d'erreur (" & Err.Number & " )", vbCritical, g_strNomApplication
  Print #g_FileDebug, Format(Now, "dd/MM/yyyy hh:mm:ss") & " Erreur n° : (" & Err.Number & ") " & Err.Description & " Procedure : AccedePerso  Module : frmAcceder.frm"
  Resume AccedePerso_Exit
End Sub
