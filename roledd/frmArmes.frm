VERSION 5.00
Object = "{0F987290-56EE-11D0-9C43-00A0C90F29FC}#1.0#0"; "ActBar.ocx"
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "VSFLEX7D.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmContrat 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Contrat"
   ClientHeight    =   3780
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5550
   ControlBox      =   0   'False
   HelpContextID   =   16
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3780
   ScaleWidth      =   5550
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   3075
      Left            =   2880
      TabIndex        =   4
      Top             =   -45
      Width           =   2625
      Begin VB.TextBox fldStrIdentifiant 
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   1035
         MaxLength       =   15
         TabIndex        =   0
         Top             =   225
         Width           =   1500
      End
      Begin VB.TextBox fldnEnrichissement 
         Height          =   330
         Left            =   1845
         MaxLength       =   4
         TabIndex        =   1
         Top             =   630
         Width           =   690
      End
      Begin VB.Label Label4 
         Caption         =   "Identification"
         Height          =   240
         Left            =   90
         TabIndex        =   6
         Top             =   270
         Width           =   915
      End
      Begin VB.Label Label1 
         Caption         =   "Enrichissement (%)"
         Height          =   285
         Left            =   90
         TabIndex        =   5
         Top             =   675
         Width           =   1320
      End
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3375
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VSFLEX7DAOCtl.vsFlexGrid vsFlexcontrat 
      Bindings        =   "FRMCON~1.frx":0000
      Height          =   3030
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   2805
      _ExtentX        =   4948
      _ExtentY        =   5345
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
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
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   3
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   4
      Cols            =   4
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   "Identification        "
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
      ExplorerBar     =   0
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
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   3405
      Width           =   5550
      _ExtentX        =   9790
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   9737
         EndProperty
      EndProperty
   End
   Begin ActiveBarLibraryCtl.ActiveBar ActiveBar1 
      Left            =   90
      Top             =   3015
      _ExtentX        =   847
      _ExtentY        =   847
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Bands           =   "FRMCON~1.frx":0014
   End
End
Attribute VB_Name = "frmContrat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const COL_IDENTIFIANT = 0
Const COL_ENRICHISSEMENT = 1
Const COL_CLIENT = 2
Dim m_strCaption As String
Dim m_bmodif As Boolean
Dim m_bCharge As Boolean


Dim g_dbs As Database

Private Sub fldnEnrichissement_Change()

  If g_ModeDebug = vbUnchecked Then  On Error GoTo fldnEnrichissement_Change_Error

  Modification


fldnEnrichissement_Change_Exit:
  Exit Sub

fldnEnrichissement_Change_Error:
  MsgBox "L'erreur suivante s'est produite : " & Err.Description & vbCrLf & vbCrLf & " numéro d'erreur (" & Err.Number & " )", vbCritical, g_strNomApplication
  print #g_FileDebug,Format (Now,"dd/MM/yyyy hh:mm:ss") & " Erreur n° : (" & Err.Number & ") " & Err.Description & " Procedure : fldnEnrichissement_Change  Module : frmArmes.frm
  Resume fldnEnrichissement_Change_Exit
End Sub
Private Sub fldnEnrichissement_GotFocus()

  If g_ModeDebug = vbUnchecked Then  On Error GoTo fldnEnrichissement_GotFocus_Error

  Selec
  StatusBar1.Panels(1).Text = "Saisissez l'enrichissement du contrat"


fldnEnrichissement_GotFocus_Exit:
  Exit Sub

fldnEnrichissement_GotFocus_Error:
  MsgBox "L'erreur suivante s'est produite : " & Err.Description & vbCrLf & vbCrLf & " numéro d'erreur (" & Err.Number & " )", vbCritical, g_strNomApplication
  print #g_FileDebug,Format (Now,"dd/MM/yyyy hh:mm:ss") & " Erreur n° : (" & Err.Number & ") " & Err.Description & " Procedure : fldnEnrichissement_GotFocus  Module : frmArmes.frm
  Resume fldnEnrichissement_GotFocus_Exit
End Sub
Private Sub fldnEnrichissement_LostFocus()

  If g_ModeDebug = vbUnchecked Then  On Error GoTo fldnEnrichissement_LostFocus_Error

  StatusBar1.Panels(1).Text = ""


fldnEnrichissement_LostFocus_Exit:
  Exit Sub

fldnEnrichissement_LostFocus_Error:
  MsgBox "L'erreur suivante s'est produite : " & Err.Description & vbCrLf & vbCrLf & " numéro d'erreur (" & Err.Number & " )", vbCritical, g_strNomApplication
  print #g_FileDebug,Format (Now,"dd/MM/yyyy hh:mm:ss") & " Erreur n° : (" & Err.Number & ") " & Err.Description & " Procedure : fldnEnrichissement_LostFocus  Module : frmArmes.frm
  Resume fldnEnrichissement_LostFocus_Exit
End Sub

Sub Modification()

  If g_ModeDebug = vbUnchecked Then  On Error GoTo Modification_Error

  
  With ActiveBar1.bands(0)
    .tools("TBAjouter").visible = False
    .tools("TBSupprimer").visible = False
    .tools("TBEnregistrer").visible = True
    .tools("TBAnnuler").visible = True
    .tools("TBFermer").visible = False
  End With
  ActiveBar1.RecalcLayout
  m_bmodif = True


Modification_Exit:
  Exit Sub

Modification_Error:
  MsgBox "L'erreur suivante s'est produite : " & Err.Description & vbCrLf & vbCrLf & " numéro d'erreur (" & Err.Number & " )", vbCritical, g_strNomApplication
  print #g_FileDebug,Format (Now,"dd/MM/yyyy hh:mm:ss") & " Erreur n° : (" & Err.Number & ") " & Err.Description & " Procedure : Modification  Module : frmArmes.frm
  Resume Modification_Exit
End Sub
Private Function NonModifier() As Boolean
  
  NonModifier = False
  If m_bmodif = True Then
    Select Case MsgBox("Voulez vous enregistrer les modifications apportées à cette arme", vbYesNoCancel Or vbInformation, Me.Caption)
      Case vbYes
        If Enregistre() = False Then Exit Function
      Case vbCancel
        Call vsFlexcontrat_SelChange
        Exit Function
    End Select
    Me.Caption = m_strCaption
    m_bmodif = False
  End If
  NonModifier = True
End Function
Function Enregistre()
Dim bExist As Boolean
Dim i As Long
Dim rstTest As Recordset
Dim rstContrat As Recordset

  Enregistre = False
  If Len(Trim(fldStrIdentifiant)) < 4 Then
    MsgBox "Saisie d'un identifiant contrat de longueur minimum de quatres caractères obligatoire", vbExclamation, Me.Caption
    Call Set_Focus(fldStrIdentifiant)
    Exit Function
  End If
  
  If (Len(fldStrIdentifiant) = 4) And (Mid(fldStrIdentifiant, 1, 1) = "0" Or Mid(fldStrIdentifiant, 2, 1) = "0") Then
    If MsgBox("Confirmez vous la saisie de cet identifiant sous cette forme ?", vbDefaultButton2 Or vbOKCancel Or vbExclamation, Me.Caption) = vbCancel Then
      Call Set_Focus(fldStrIdentifiant)
      Exit Function
    End If
  End If
  
  fldnEnrichissement = FrmtNum(fldnEnrichissement)
  If Not IsNumeric(fldnEnrichissement) Then
    MsgBox "Veuillez saisir un enrichissement numérique", vbExclamation, Me.Caption
    Call Set_Focus(fldnEnrichissement)
    Exit Function
  Else
    If CDbl(fldnEnrichissement) <= 0 Or CDbl(fldnEnrichissement) >= 5 Then
      MsgBox "Veuillez saisir un enrichissement en % inclus entre 0 et 5 ", vbExclamation, Me.Caption
      Call Set_Focus(fldnEnrichissement)
      Exit Function
    End If
  End If
      
  Screen.MousePointer = vbHourglass
  Set rstContrat = g_dbs.OpenRecordset("contrat", dbOpenTable)
  rstContrat.Index = "primarykey"
  rstContrat.Seek "=", fldStrIdentifiant
  If rstContrat.NoMatch Then
    rstContrat.AddNew
    rstContrat!identifiant = fldStrIdentifiant
    rstContrat!DateHeure = Format(Now, "dd/MM/yyyy hh:mm:ss")
  Else
    rstContrat.Edit
  End If
  rstContrat!ENRICHISSEMENT = Format(fldnEnrichissement, "0.##")
  'rstContrat!CODE_OPE = g_OP
  rstContrat.Update
  rstContrat.Close
  vsFlexcontrat.Redraw = False
  i = vsFlexcontrat.TopRow
  Call RemplieLst
  Call RetrouvVSFLex(vsFlexcontrat, COL_IDENTIFIANT, fldStrIdentifiant, i)
  vsFlexcontrat.Redraw = True
  Enregistre = True
  m_bmodif = False
  m_bCharge = False
  
End Function
Private Sub Supprimer()
Dim n As Integer
Dim rstTest As Recordset
Dim bExist As Boolean

  If g_ModeDebug = vbUnchecked Then  On Error GoTo Supprimer_Error
 
  With vsFlexcontrat
    For n = 1 To .Rows - 1
      If .IsSelected(n) Then
        If MsgBox("Confirmer la suppression de ce contrat :" & .Cell(flexcpText, n, COL_IDENTIFIANT), vbDefaultButton2 Or vbOKCancel Or vbQuestion, "Suppression") = vbOK Then
          Set rstTest = g_dbs.OpenRecordset("select * from contratlot where numerocontrat='" & FormatStrSQL(.Cell(flexcpText, n, COL_IDENTIFIANT)) & "'", dbOpenDynaset)
          bExist = Not rstTest.EOF
          rstTest.Close
          If bExist Then
            If MsgBox("La table lot comprend des enregistrements connexes " & vbCrLf & "si vous supprimez ce contrat vous effacerez les lots correspondants " & vbCrLf & vbCrLf & "Appuyez Ok pour supprimer les lots...", vbExclamation Or vbOKCancel Or vbDefaultButton2, Me.Caption) = vbOK Then
              g_dbs.Execute "delete from MesurePastille where [numeroContrat]='" & .Cell(flexcpText, n, COL_IDENTIFIANT) & "'", dbFailOnError
              g_dbs.Execute "delete from Empilement where [numeroContrat]='" & .Cell(flexcpText, n, COL_IDENTIFIANT) & "'", dbFailOnError
              g_dbs.Execute "delete from FourGeometrique where [numeroContrat]='" & .Cell(flexcpText, n, COL_IDENTIFIANT) & "'", dbFailOnError
              g_dbs.Execute "delete from contratLot where [numeroContrat]='" & .Cell(flexcpText, n, COL_IDENTIFIANT) & "'", dbFailOnError
              g_dbs.Execute "delete from trainEmpilement where [numeroContrat]='" & .Cell(flexcpText, n, COL_IDENTIFIANT) & "'", dbFailOnError
              ''detruit les palettes sans correspondances
              g_dbs.Execute "Delete DISTINCTROW Palette.* FROM Palette LEFT JOIN Empilement ON (Palette.[N°Four] = Empilement.[N°Four]) AND (Palette.DatePalette = Empilement.DatePalette) AND (Palette.[N°Palette] = Empilement.[N°Palette]) WHERE (((Empilement.[N°Four]) Is Null) AND ((Empilement.[N°Palette]) Is Null) AND ((Empilement.DatePalette) Is Null))", dbFailOnError
            Else
              Exit Sub
            End If
          End If
          ''
          g_dbs.Execute "delete from contrat where [identifiant]='" & FormatStrSQL(.Cell(flexcpText, n, COL_IDENTIFIANT)) & "'", dbFailOnError
          Call RemplieLst
          If vsFlexcontrat.Rows > 1 Then
            vsFlexcontrat.IsSelected(1) = True
          Else
            efface Me
            Call vsFlexcontrat_SelChange
          End If
        End If
        Exit For
      End If
    Next n
  End With
  

Supprimer_Exit:
  Exit Sub

Supprimer_Error:
  MsgBox "L'erreur suivante s'est produite : " & Err.Description & vbCrLf & vbCrLf & " numéro d'erreur (" & Err.Number & " )", vbCritical, g_strNomApplication
  print #g_FileDebug,Format (Now,"dd/MM/yyyy hh:mm:ss") & " Erreur n° : (" & Err.Number & ") " & Err.Description & " Procedure : Supprimer  Module : frmArmes.frm
  Resume Supprimer_Exit
End Sub
Private Sub Fermer()

  If g_ModeDebug = vbUnchecked Then  On Error GoTo Fermer_Error

  If ActiveBar1.bands(0).tools("TBFermer").visible = True Then
    Unload Me
  Else
    Call vsFlexcontrat_SelChange
    Me.Caption = m_strCaption
  End If


Fermer_Exit:
  Exit Sub

Fermer_Error:
  MsgBox "L'erreur suivante s'est produite : " & Err.Description & vbCrLf & vbCrLf & " numéro d'erreur (" & Err.Number & " )", vbCritical, g_strNomApplication
  print #g_FileDebug,Format (Now,"dd/MM/yyyy hh:mm:ss") & " Erreur n° : (" & Err.Number & ") " & Err.Description & " Procedure : Fermer  Module : frmArmes.frm
  Resume Fermer_Exit
End Sub
Private Sub ActiveBar1_Click(ByVal Tool As ActiveBarLibraryCtl.Tool)

  If g_ModeDebug = vbUnchecked Then  On Error GoTo ActiveBar1_Click_Error

  Select Case Tool.Name
    Case "TBFermer"
      Fermer
    Case "TBEnregistrer"
      Enregistre
    Case "TBAnnuler"
      Fermer
    Case "TBAjouter"
      'ajout
      efface Me
      Call Set_Focus(fldStrIdentifiant)
    Case "TBSupprimer"
      Supprimer
  End Select


ActiveBar1_Click_Exit:
  Exit Sub

ActiveBar1_Click_Error:
  MsgBox "L'erreur suivante s'est produite : " & Err.Description & vbCrLf & vbCrLf & " numéro d'erreur (" & Err.Number & " )", vbCritical, g_strNomApplication
  print #g_FileDebug,Format (Now,"dd/MM/yyyy hh:mm:ss") & " Erreur n° : (" & Err.Number & ") " & Err.Description & " Procedure : ActiveBar1_Click  Module : frmArmes.frm
  Resume ActiveBar1_Click_Exit
End Sub
Private Sub fldStrIdentifiant_Change()

  If g_ModeDebug = vbUnchecked Then  On Error GoTo fldStrIdentifiant_Change_Error

  Modification


fldStrIdentifiant_Change_Exit:
  Exit Sub

fldStrIdentifiant_Change_Error:
  MsgBox "L'erreur suivante s'est produite : " & Err.Description & vbCrLf & vbCrLf & " numéro d'erreur (" & Err.Number & " )", vbCritical, g_strNomApplication
  print #g_FileDebug,Format (Now,"dd/MM/yyyy hh:mm:ss") & " Erreur n° : (" & Err.Number & ") " & Err.Description & " Procedure : fldStrIdentifiant_Change  Module : frmArmes.frm
  Resume fldStrIdentifiant_Change_Exit
End Sub
Private Sub fldStrIdentifiant_GotFocus()

  If g_ModeDebug = vbUnchecked Then  On Error GoTo fldStrIdentifiant_GotFocus_Error

  Selec
  StatusBar1.Panels(1).Text = "Saisissez l'identifiant du contrat"


fldStrIdentifiant_GotFocus_Exit:
  Exit Sub

fldStrIdentifiant_GotFocus_Error:
  MsgBox "L'erreur suivante s'est produite : " & Err.Description & vbCrLf & vbCrLf & " numéro d'erreur (" & Err.Number & " )", vbCritical, g_strNomApplication
  print #g_FileDebug,Format (Now,"dd/MM/yyyy hh:mm:ss") & " Erreur n° : (" & Err.Number & ") " & Err.Description & " Procedure : fldStrIdentifiant_GotFocus  Module : frmArmes.frm
  Resume fldStrIdentifiant_GotFocus_Exit
End Sub
Private Sub fldStrIdentifiant_KeyPress(KeyAscii As Integer)

  If g_ModeDebug = vbUnchecked Then  On Error GoTo fldStrIdentifiant_KeyPress_Error

  If KeyAscii > 96 And KeyAscii < 123 Then
    KeyAscii = KeyAscii - 32
  End If


fldStrIdentifiant_KeyPress_Exit:
  Exit Sub

fldStrIdentifiant_KeyPress_Error:
  MsgBox "L'erreur suivante s'est produite : " & Err.Description & vbCrLf & vbCrLf & " numéro d'erreur (" & Err.Number & " )", vbCritical, g_strNomApplication
  print #g_FileDebug,Format (Now,"dd/MM/yyyy hh:mm:ss") & " Erreur n° : (" & Err.Number & ") " & Err.Description & " Procedure : fldStrIdentifiant_KeyPress  Module : frmArmes.frm
  Resume fldStrIdentifiant_KeyPress_Exit
End Sub
Private Sub fldStrIdentifiant_LostFocus()

  If g_ModeDebug = vbUnchecked Then  On Error GoTo fldStrIdentifiant_LostFocus_Error

  StatusBar1.Panels(1).Text = ""


fldStrIdentifiant_LostFocus_Exit:
  Exit Sub

fldStrIdentifiant_LostFocus_Error:
  MsgBox "L'erreur suivante s'est produite : " & Err.Description & vbCrLf & vbCrLf & " numéro d'erreur (" & Err.Number & " )", vbCritical, g_strNomApplication
  print #g_FileDebug,Format (Now,"dd/MM/yyyy hh:mm:ss") & " Erreur n° : (" & Err.Number & ") " & Err.Description & " Procedure : fldStrIdentifiant_LostFocus  Module : frmArmes.frm
  Resume fldStrIdentifiant_LostFocus_Exit
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

  If g_ModeDebug = vbUnchecked Then  On Error GoTo Form_KeyDown_Error

  Select Case KeyCode
    Case vbKeyF2
      If ActiveBar1.bands(0).tools("TBEnregistrer").visible = True Then Call ActiveBar1_Click(ActiveBar1.bands(0).tools("TBEnregistrer"))
    Case vbKeyReturn
      SendKeys "{TAB}"
    Case vbKeyEscape
      Fermer
    Case vbKeyF9
      If ActiveBar1.bands(0).tools("TBAjouter").visible = True Then Call ActiveBar1_Click(ActiveBar1.bands(0).tools("TBAjouter"))
    Case vbKeyF10
      If ActiveBar1.bands(0).tools("TBSupprimer").visible = True Then Supprimer
  End Select


Form_KeyDown_Exit:
  Exit Sub

Form_KeyDown_Error:
  MsgBox "L'erreur suivante s'est produite : " & Err.Description & vbCrLf & vbCrLf & " numéro d'erreur (" & Err.Number & " )", vbCritical, g_strNomApplication
  print #g_FileDebug,Format (Now,"dd/MM/yyyy hh:mm:ss") & " Erreur n° : (" & Err.Number & ") " & Err.Description & " Procedure : Form_KeyDown  Module : frmArmes.frm
  Resume Form_KeyDown_Exit
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)

  If g_ModeDebug = vbUnchecked Then  On Error GoTo Form_KeyPress_Error
  
  Select Case KeyAscii
    Case vbKeyEscape
      KeyAscii = 0
    Case vbKeyReturn
      KeyAscii = 0
  End Select

Form_KeyPress_Exit:
  Exit Sub

Form_KeyPress_Error:
  MsgBox "L'erreur suivante s'est produite : " & Err.Description & vbCrLf & vbCrLf & " numéro d'erreur (" & Err.Number & " )", vbCritical, g_strNomApplication
  print #g_FileDebug,Format (Now,"dd/MM/yyyy hh:mm:ss") & " Erreur n° : (" & Err.Number & ") " & Err.Description & " Procedure : Form_KeyPress  Module : frmArmes.frm
  Resume Form_KeyPress_Exit
End Sub
Private Sub Form_Load()
Dim bExist As Boolean

  If g_ModeDebug = vbUnchecked Then  On Error GoTo Form_Load_Error
  
  
  Set g_dbs = OpenDatabase(App.Path & "\" & "bd1.mdb")
  
  
  m_strCaption = Me.Caption
   
  
   
   ' Ouverture base de donnée
  Data1.DatabaseName = g_dbs.Name
  Call RemplieLst
  If vsFlexcontrat.Rows > 1 Then
    vsFlexcontrat.IsSelected(1) = True
  Else
    Call vsFlexcontrat_SelChange
  End If

Form_Load_Exit:
  Exit Sub

Form_Load_Error:
  MsgBox "L'erreur suivante s'est produite : " & Err.Description & vbCrLf & vbCrLf & " numéro d'erreur (" & Err.Number & " )", vbCritical, g_strNomApplication
  print #g_FileDebug,Format (Now,"dd/MM/yyyy hh:mm:ss") & " Erreur n° : (" & Err.Number & ") " & Err.Description & " Procedure : Form_Load  Module : frmArmes.frm
  Resume Form_Load_Exit
End Sub
Sub RemplieLst()
Dim i As Integer
Dim stSql As String

  If g_ModeDebug = vbUnchecked Then  On Error GoTo RemplieLst_Error
 
  Screen.MousePointer = vbHourglass
  stSql = "SELECT Identifiant,Enrichissement from contrat"
  Data1.RecordSource = stSql
  Data1.Refresh
  With vsFlexcontrat
    .Cols = 9
    .Redraw = False
    .DataRefresh
    .FormatString = "Identification "
    .ColAlignment(COL_IDENTIFIANT) = flexAlignLeftCenter
    .ColHidden(COL_ENRICHISSEMENT) = True
    .Redraw = True
  End With
  Me.Caption = m_strCaption
  Screen.MousePointer = vbNormal

RemplieLst_Exit:
  Exit Sub

RemplieLst_Error:
  MsgBox "L'erreur suivante s'est produite : " & Err.Description & vbCrLf & vbCrLf & " numéro d'erreur (" & Err.Number & " )", vbCritical, g_strNomApplication
  print #g_FileDebug,Format (Now,"dd/MM/yyyy hh:mm:ss") & " Erreur n° : (" & Err.Number & ") " & Err.Description & " Procedure : RemplieLst  Module : frmArmes.frm
  Resume RemplieLst_Exit
End Sub
Private Sub Form_Unload(Cancel As Integer)

  If g_ModeDebug = vbUnchecked Then  On Error GoTo Form_Unload_Error

  If NonModifier() = False Then Cancel = True: Exit Sub


Form_Unload_Exit:
  Exit Sub

Form_Unload_Error:
  MsgBox "L'erreur suivante s'est produite : " & Err.Description & vbCrLf & vbCrLf & " numéro d'erreur (" & Err.Number & " )", vbCritical, g_strNomApplication
  print #g_FileDebug,Format (Now,"dd/MM/yyyy hh:mm:ss") & " Erreur n° : (" & Err.Number & ") " & Err.Description & " Procedure : Form_Unload  Module : frmArmes.frm
  Resume Form_Unload_Exit
End Sub

Private Sub vsFlexcontrat_Click()

  If g_ModeDebug = vbUnchecked Then  On Error GoTo vsFlexcontrat_Click_Error

  vsFlexcontrat_SelChange


vsFlexcontrat_Click_Exit:
  Exit Sub

vsFlexcontrat_Click_Error:
  MsgBox "L'erreur suivante s'est produite : " & Err.Description & vbCrLf & vbCrLf & " numéro d'erreur (" & Err.Number & " )", vbCritical, g_strNomApplication
  print #g_FileDebug,Format (Now,"dd/MM/yyyy hh:mm:ss") & " Erreur n° : (" & Err.Number & ") " & Err.Description & " Procedure : vsFlexcontrat_Click  Module : frmArmes.frm
  Resume vsFlexcontrat_Click_Exit
End Sub
Private Sub vsFlexcontrat_SelChange()
Dim i, n As Integer
Dim bVisible As Boolean

  If g_ModeDebug = vbUnchecked Then  On Error GoTo vsFlexcontrat_SelChange_Error

  bVisible = False
  With vsFlexcontrat
    For n = 1 To .Rows - 1
      If .IsSelected(n) = True Then
        m_bCharge = True
        fldStrIdentifiant = .Cell(flexcpText, n, COL_IDENTIFIANT)
        fldnEnrichissement = .Cell(flexcpText, n, COL_ENRICHISSEMENT)
        bVisible = True
        Exit For
      End If
    Next n
  End With
  m_bCharge = False
  m_bmodif = False
  Me.Caption = m_strCaption
  With ActiveBar1.bands(0)
    .tools("TBAjouter").visible = True
    .tools("TBSupprimer").visible = bVisible
    .tools("TBEnregistrer").visible = False
    .tools("TBAnnuler").visible = False
    .tools("TBFermer").visible = True
  End With
  ActiveBar1.RecalcLayout


vsFlexcontrat_SelChange_Exit:
  Exit Sub

vsFlexcontrat_SelChange_Error:
  MsgBox "L'erreur suivante s'est produite : " & Err.Description & vbCrLf & vbCrLf & " numéro d'erreur (" & Err.Number & " )", vbCritical, g_strNomApplication
  print #g_FileDebug,Format (Now,"dd/MM/yyyy hh:mm:ss") & " Erreur n° : (" & Err.Number & ") " & Err.Description & " Procedure : vsFlexcontrat_SelChange  Module : frmArmes.frm
  Resume vsFlexcontrat_SelChange_Exit
End Sub
