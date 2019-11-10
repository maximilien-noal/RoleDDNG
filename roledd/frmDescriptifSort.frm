VERSION 5.00

Begin VB.Form frmDescriptifSort
    Caption = "Description du sort"
    ScaleMode = 1
    AutoRedraw = 0              'False
    FontTransparent = -1              'True
    BorderStyle = 4
    LinkTopic = "Form1"
    MaxButton = 0              'False
    MinButton = 0              'False
    ControlBox = 0              'False
    ClipControls = 0              'False
    ClientLeft   = 45
    ClientTop    = 285
    ClientWidth  = 13350
    ClientHeight = 7515
    BeginProperty Font
        Name          = "Times New Roman"
        Size          = 9,75
        Charset       = 0
        Weight        = 700
        Underline     = 0              'False
        Italic        = 0              'False
        Strikethrough = 0              'False
    EndProperty
    ShowInTaskbar = 0              'False
    StartupPosition = 1
    Begin VB.ComboBox CmbSort
        Left   = 120
        Top    = 7200
        Width  = 3975
        Height = 345
        Text = "Combo1&&"
        TabIndex = 38
        BeginProperty Font
            Name          = "Times New Roman"
            Size          = 9
            Charset       = 0
            Weight        = 400
            Underline     = 0              'False
            Italic        = 0              'False
            Strikethrough = 0              'False
        EndProperty
    End
    Begin VB.CommandButton btnPrecedent
        Caption = "&Précédent"
        Left   = 6600
        Top    = 7200
        Width  = 1275
        Height = 285
        TabIndex = 37
        BeginProperty Font
            Name          = "Times New Roman"
            Size          = 9
            Charset       = 0
            Weight        = 400
            Underline     = 0              'False
            Italic        = 0              'False
            Strikethrough = 0              'False
        EndProperty
    End
    Begin VB.CommandButton btnsuivant
        Caption = "&Suivant"
        Left   = 8040
        Top    = 7200
        Width  = 1275
        Height = 285
        TabIndex = 36
        BeginProperty Font
            Name          = "Times New Roman"
            Size          = 9
            Charset       = 0
            Weight        = 400
            Underline     = 0              'False
            Italic        = 0              'False
            Strikethrough = 0              'False
        EndProperty
    End
    Begin VB.TextBox fldstrClasseNiveau
        BackColor = 12632256
        Left   = 10260
        Top    = 45
        Width  = 2955
        Height = 330
        TabIndex = 34
        MultiLine = -1              'True
        BeginProperty Font
            Name          = "Times New Roman"
            Size          = 9
            Charset       = 0
            Weight        = 400
            Underline     = 0              'False
            Italic        = 0              'False
            Strikethrough = 0              'False
        EndProperty
    End
    Begin VB.TextBox fldstrSource
        BackColor = 12632256
        Left   = 5160
        Top    = 6760
        Width  = 3960
        Height = 330
        TabIndex = 32
        MultiLine = -1              'True
        BeginProperty Font
            Name          = "Times New Roman"
            Size          = 9
            Charset       = 0
            Weight        = 400
            Underline     = 0              'False
            Italic        = 0              'False
            Strikethrough = 0              'False
        EndProperty
    End
    Begin VB.CommandButton btnEnregistrer
        Caption = "&Enregistrer"
        Left   = 10515
        Top    = 7200
        Width  = 1275
        Height = 285
        TabIndex = 25
        BeginProperty Font
            Name          = "Times New Roman"
            Size          = 9
            Charset       = 0
            Weight        = 400
            Underline     = 0              'False
            Italic        = 0              'False
            Strikethrough = 0              'False
        EndProperty
    End
    Begin VB.Frame frm1
        Left   = 45
        Top    = 360
        Width  = 13260
        Height = 6780
        TabIndex = 2
        BeginProperty Font
            Name          = "MS Sans Serif"
            Size          = 8,25
            Charset       = 0
            Weight        = 400
            Underline     = 0              'False
            Italic        = 0              'False
            Strikethrough = 0              'False
        EndProperty
        Begin VB.TextBox fldstrobjet
            BackColor = 12632256
            Left   = 10200
            Top    = 6400
            Width  = 2760
            Height = 330
            TabIndex = 35
            MultiLine = -1              'True
            BeginProperty Font
                Name          = "Times New Roman"
                Size          = 9
                Charset       = 0
                Weight        = 400
                Underline     = 0              'False
                Italic        = 0              'False
                Strikethrough = 0              'False
            EndProperty
        End
        Begin VB.TextBox fldnDD
            BackColor = 12632256
            Left   = 4275
            Top    = 5800
            Width  = 615
            Height = 330
            TabIndex = 33
            MultiLine = -1              'True
            BeginProperty Font
                Name          = "Times New Roman"
                Size          = 9
                Charset       = 0
                Weight        = 400
                Underline     = 0              'False
                Italic        = 0              'False
                Strikethrough = 0              'False
            EndProperty
        End
        Begin VB.TextBox fldstrzoneeffet
            BackColor = 12632256
            Left   = 100
            Top    = 4600
            Width  = 4800
            Height = 330
            TabIndex = 31
            MultiLine = -1              'True
            BeginProperty Font
                Name          = "Times New Roman"
                Size          = 9
                Charset       = 0
                Weight        = 400
                Underline     = 0              'False
                Italic        = 0              'False
                Strikethrough = 0              'False
            EndProperty
        End
        Begin VB.TextBox fldPourDevellopper
            BackColor = 12632256
            Left   = 5100
            Top    = 1000
            Width  = 8100
            Height = 1755
            TabIndex = 28
            MultiLine = -1              'True
            BeginProperty Font
                Name          = "Times New Roman"
                Size          = 9
                Charset       = 0
                Weight        = 400
                Underline     = 0              'False
                Italic        = 0              'False
                Strikethrough = 0              'False
            EndProperty
        End
        Begin VB.TextBox fldDDDifficultee
            BackColor = 12632256
            Left   = 5100
            Top    = 400
            Width  = 8100
            Height = 330
            TabIndex = 27
            MultiLine = -1              'True
            BeginProperty Font
                Name          = "Times New Roman"
                Size          = 9
                Charset       = 0
                Weight        = 400
                Underline     = 0              'False
                Italic        = 0              'False
                Strikethrough = 0              'False
            EndProperty
        End
        Begin VB.TextBox fldstrDomaine
            BackColor = 12632256
            Left   = 100
            Top    = 400
            Width  = 4800
            Height = 330
            TabIndex = 13
            MultiLine = -1              'True
            BeginProperty Font
                Name          = "Times New Roman"
                Size          = 9
                Charset       = 0
                Weight        = 400
                Underline     = 0              'False
                Italic        = 0              'False
                Strikethrough = 0              'False
            EndProperty
        End
        Begin VB.TextBox fldstrNiveau
            BackColor = 12632256
            Left   = 100
            Top    = 1000
            Width  = 4800
            Height = 330
            TabIndex = 12
            MultiLine = -1              'True
            BeginProperty Font
                Name          = "Times New Roman"
                Size          = 9
                Charset       = 0
                Weight        = 400
                Underline     = 0              'False
                Italic        = 0              'False
                Strikethrough = 0              'False
            EndProperty
        End
        Begin VB.TextBox fldstrComposante
            BackColor = 12632256
            Left   = 100
            Top    = 1600
            Width  = 4800
            Height = 330
            TabIndex = 11
            MultiLine = -1              'True
            BeginProperty Font
                Name          = "Times New Roman"
                Size          = 9
                Charset       = 0
                Weight        = 400
                Underline     = 0              'False
                Italic        = 0              'False
                Strikethrough = 0              'False
            EndProperty
        End
        Begin VB.TextBox fldstrtempincantation
            BackColor = 12632256
            Left   = 100
            Top    = 2200
            Width  = 4800
            Height = 330
            TabIndex = 10
            MultiLine = -1              'True
            BeginProperty Font
                Name          = "Times New Roman"
                Size          = 9
                Charset       = 0
                Weight        = 400
                Underline     = 0              'False
                Italic        = 0              'False
                Strikethrough = 0              'False
            EndProperty
        End
        Begin VB.TextBox fldstrportee
            BackColor = 12632256
            Left   = 100
            Top    = 2800
            Width  = 4800
            Height = 330
            TabIndex = 9
            MultiLine = -1              'True
            BeginProperty Font
                Name          = "Times New Roman"
                Size          = 9
                Charset       = 0
                Weight        = 400
                Underline     = 0              'False
                Italic        = 0              'False
                Strikethrough = 0              'False
            EndProperty
        End
        Begin VB.TextBox fldstrcible
            BackColor = 12632256
            Left   = 100
            Top    = 3400
            Width  = 4800
            Height = 330
            TabIndex = 8
            MultiLine = -1              'True
            BeginProperty Font
                Name          = "Times New Roman"
                Size          = 9
                Charset       = 0
                Weight        = 400
                Underline     = 0              'False
                Italic        = 0              'False
                Strikethrough = 0              'False
            EndProperty
        End
        Begin VB.TextBox fldstrdurée
            BackColor = 12632256
            Left   = 100
            Top    = 5200
            Width  = 4800
            Height = 330
            TabIndex = 7
            MultiLine = -1              'True
            BeginProperty Font
                Name          = "Times New Roman"
                Size          = 9
                Charset       = 0
                Weight        = 400
                Underline     = 0              'False
                Italic        = 0              'False
                Strikethrough = 0              'False
            EndProperty
        End
        Begin VB.TextBox fldstrEffet
            BackColor = 12632256
            Left   = 100
            Top    = 4000
            Width  = 4800
            Height = 330
            TabIndex = 6
            MultiLine = -1              'True
            BeginProperty Font
                Name          = "Times New Roman"
                Size          = 9
                Charset       = 0
                Weight        = 400
                Underline     = 0              'False
                Italic        = 0              'False
                Strikethrough = 0              'False
            EndProperty
        End
        Begin VB.TextBox fldstrjetsauvegarde
            BackColor = 12632256
            Left   = 100
            Top    = 5800
            Width  = 3765
            Height = 330
            TabIndex = 5
            MultiLine = -1              'True
            BeginProperty Font
                Name          = "Times New Roman"
                Size          = 9
                Charset       = 0
                Weight        = 400
                Underline     = 0              'False
                Italic        = 0              'False
                Strikethrough = 0              'False
            EndProperty
        End
        Begin VB.TextBox fldstrresistancemagie
            BackColor = 12632256
            Left   = 100
            Top    = 6400
            Width  = 4800
            Height = 330
            TabIndex = 4
            MultiLine = -1              'True
            BeginProperty Font
                Name          = "Times New Roman"
                Size          = 9
                Charset       = 0
                Weight        = 400
                Underline     = 0              'False
                Italic        = 0              'False
                Strikethrough = 0              'False
            EndProperty
        End
        Begin VB.TextBox fldstrExplication
            BackColor = 12632256
            Left   = 5100
            Top    = 3000
            Width  = 8100
            Height = 3075
            TabIndex = 3
            MultiLine = -1              'True
            ScrollBars = 2
            BeginProperty Font
                Name          = "Times New Roman"
                Size          = 9
                Charset       = 0
                Weight        = 400
                Underline     = 0              'False
                Italic        = 0              'False
                Strikethrough = 0              'False
            EndProperty
        End
        Begin VB.Label Label8
            Caption = "Zone d'effet"
            Left   = 120
            Top    = 4320
            Width  = 1455
            Height = 240
            TabIndex = 30
            BeginProperty Font
                Name          = "Times New Roman"
                Size          = 9
                Charset       = 0
                Weight        = 400
                Underline     = 0              'False
                Italic        = 0              'False
                Strikethrough = 0              'False
            EndProperty
        End
        Begin VB.Label Label6
            Caption = "Pour dévellopper"
            Left   = 5160
            Top    = 720
            Width  = 1920
            Height = 240
            TabIndex = 29
            BeginProperty Font
                Name          = "Times New Roman"
                Size          = 9
                Charset       = 0
                Weight        = 400
                Underline     = 0              'False
                Italic        = 0              'False
                Strikethrough = 0              'False
            EndProperty
        End
        Begin VB.Label Label5
            Caption = "DD pour lancer le sort épique"
            Left   = 5200
            Top    = 150
            Width  = 2895
            Height = 285
            TabIndex = 26
            BeginProperty Font
                Name          = "Times New Roman"
                Size          = 9
                Charset       = 0
                Weight        = 400
                Underline     = 0              'False
                Italic        = 0              'False
                Strikethrough = 0              'False
            EndProperty
        End
        Begin VB.Label Label1
            Caption = "Ecole [Branche] (Registre)"
            Left   = 120
            Top    = 150
            Width  = 2115
            Height = 240
            TabIndex = 24
            BeginProperty Font
                Name          = "Times New Roman"
                Size          = 9
                Charset       = 0
                Weight        = 400
                Underline     = 0              'False
                Italic        = 0              'False
                Strikethrough = 0              'False
            EndProperty
        End
        Begin VB.Label Label2
            Caption = "Niveau"
            Left   = 120
            Top    = 720
            Width  = 690
            Height = 240
            TabIndex = 23
            BeginProperty Font
                Name          = "Times New Roman"
                Size          = 9
                Charset       = 0
                Weight        = 400
                Underline     = 0              'False
                Italic        = 0              'False
                Strikethrough = 0              'False
            EndProperty
        End
        Begin VB.Label Label3
            Caption = "Composante"
            Left   = 120
            Top    = 1320
            Width  = 960
            Height = 240
            TabIndex = 22
            BeginProperty Font
                Name          = "Times New Roman"
                Size          = 9
                Charset       = 0
                Weight        = 400
                Underline     = 0              'False
                Italic        = 0              'False
                Strikethrough = 0              'False
            EndProperty
        End
        Begin VB.Label Label4
            Caption = "Temps d'incantation"
            Left   = 120
            Top    = 1920
            Width  = 1695
            Height = 240
            TabIndex = 21
            BeginProperty Font
                Name          = "Times New Roman"
                Size          = 9
                Charset       = 0
                Weight        = 400
                Underline     = 0              'False
                Italic        = 0              'False
                Strikethrough = 0              'False
            EndProperty
        End
        Begin VB.Label gdfg
            Caption = "Portée"
            Left   = 120
            Top    = 2520
            Width  = 1455
            Height = 240
            TabIndex = 20
            BeginProperty Font
                Name          = "Times New Roman"
                Size          = 9
                Charset       = 0
                Weight        = 400
                Underline     = 0              'False
                Italic        = 0              'False
                Strikethrough = 0              'False
            EndProperty
        End
        Begin VB.Label Cible
            Caption = "Cible"
            Left   = 120
            Top    = 3120
            Width  = 1455
            Height = 240
            TabIndex = 19
            BeginProperty Font
                Name          = "Times New Roman"
                Size          = 9
                Charset       = 0
                Weight        = 400
                Underline     = 0              'False
                Italic        = 0              'False
                Strikethrough = 0              'False
            EndProperty
        End
        Begin VB.Label Label7
            Caption = "Durée"
            Left   = 120
            Top    = 4920
            Width  = 555
            Height = 240
            TabIndex = 18
            BeginProperty Font
                Name          = "Times New Roman"
                Size          = 9
                Charset       = 0
                Weight        = 400
                Underline     = 0              'False
                Italic        = 0              'False
                Strikethrough = 0              'False
            EndProperty
        End
        Begin VB.Label sfdsf
            Caption = "Effet"
            Left   = 120
            Top    = 3720
            Width  = 1455
            Height = 240
            TabIndex = 17
            BeginProperty Font
                Name          = "Times New Roman"
                Size          = 9
                Charset       = 0
                Weight        = 400
                Underline     = 0              'False
                Italic        = 0              'False
                Strikethrough = 0              'False
            EndProperty
        End
        Begin VB.Label Label9
            Caption = "Jet de sauvegarde"
            Left   = 120
            Top    = 5520
            Width  = 1455
            Height = 240
            TabIndex = 16
            BeginProperty Font
                Name          = "Times New Roman"
                Size          = 9
                Charset       = 0
                Weight        = 400
                Underline     = 0              'False
                Italic        = 0              'False
                Strikethrough = 0              'False
            EndProperty
        End
        Begin VB.Label Label10
            Caption = "Résistance à la magie"
            Left   = 120
            Top    = 6120
            Width  = 1695
            Height = 240
            TabIndex = 15
            BeginProperty Font
                Name          = "Times New Roman"
                Size          = 9
                Charset       = 0
                Weight        = 400
                Underline     = 0              'False
                Italic        = 0              'False
                Strikethrough = 0              'False
            EndProperty
        End
        Begin VB.Label Label11
            Caption = "Explication du sort"
            Left   = 5205
            Top    = 2760
            Width  = 1455
            Height = 240
            TabIndex = 14
            BeginProperty Font
                Name          = "Times New Roman"
                Size          = 9
                Charset       = 0
                Weight        = 400
                Underline     = 0              'False
                Italic        = 0              'False
                Strikethrough = 0              'False
            EndProperty
        End
    End
    Begin VB.CommandButton btnFermer
        Caption = "&Fermer"
        Left   = 11910
        Top    = 7200
        Width  = 1275
        Height = 285
        TabIndex = 1
        Cancel = -1              'True
        BeginProperty Font
            Name          = "Times New Roman"
            Size          = 9
            Charset       = 0
            Weight        = 400
            Underline     = 0              'False
            Italic        = 0              'False
            Strikethrough = 0              'False
        EndProperty
    End
    Begin VB.Label labnom
        ForeColor = -2147483642
        Left   = 45
        Top    = 45
        Width  = 9870
        Height = 330
        TabIndex = 0
        BorderStyle = 1
        Alignment = 2
    End
End
Public Function Affiche(arg_0 As Unknow, arg_1 As Unknow, arg_2 As Unknow, arg_3 As Unknow, arg_4 As Unknow, arg_5 As Unknow, arg_6 As Unknow, arg_7 As Unknow, arg_8 As Unknow, arg_9 As Unknow, arg_A As Unknow, arg_B As Unknow, arg_C As Unknow, arg_D As Unknow, arg_E As Unknow, arg_F As Unknow, arg_10 As Unknow, arg_11 As Unknow, arg_12 As Unknow, arg_13 As Unknow, arg_14 As Unknow, arg_15 As Unknow, arg_16 As Unknow, arg_17 As Unknow, arg_18 As Unknow, arg_19 As Unknow, arg_1A As Unknow, arg_1B As Unknow, arg_1C As Unknow, arg_1D As Unknow, arg_1E As Unknow, arg_1F As Unknow, arg_20 As Unknow, arg_21 As Unknow, arg_22 As Unknow, arg_23 As Unknow, arg_24 As Unknow, arg_25 As Unknow, arg_26 As Unknow, arg_27 As Unknow, arg_28 As Unknow, arg_29 As Unknow, arg_2A As Unknow, arg_2B As Unknow, arg_2C As Unknow, arg_2D As Unknow, arg_2E As Unknow, arg_2F As Unknow, arg_30 As Unknow, arg_31 As Unknow, arg_32 As Unknow, arg_33 As Unknow, arg_34 As Unknow, arg_35 As Unknow, arg_36 As Unknow, arg_37 As Unknow, arg_38 As Unknow, arg_39 As Unknow, arg_3A As Unknow, arg_3B As Unknow, arg_3C As Unknow)
'006c8810    55                      push ebp
'006c8811    8bec                    mov ebp, esp
'006c8813    83ec0c                  sub esp, 0c

' *** Reference to "__vbaExceptHandler"
'006c8816    6866724000              push 00407266
'006c881b    64a100000000            mov ax, word ptr fs:[00000000]
'006c8821    50                      push eax
    'Reference to '__except_list'
'006c8822    64892500000000          mov dword ptr fs:[00000000], esp
'006c8829    81ec30020000            sub esp, 00000230
'006c882f    53                      push ebx
'006c8830    56                      push esi
'006c8831    57                      push edi
'006c8832    8965f4                  mov dword ptr [ebp-0c], esp
'006c8835    c745f898674000          mov dword ptr [ebp-08], 00406798
'006c883c    33ff                    xor edi, edi
'006c883e    897dfc                  mov dword ptr [ebp-04], edi
'006c8841    8b7508                  mov esi, dword ptr [ebp+08]
'006c8844    8b06                    mov eax, dword ptr [esi]
'006c8846    56                      push esi
'006c8847    ff5004                  call dword ptr [eax+04]
'006c884a    8b0e                    mov ecx, dword ptr [esi]
'006c884c    56                      push esi
'006c884d    897de8                  mov dword ptr [ebp-18], edi
'006c8850    897de4                  mov dword ptr [ebp-1c], edi
'006c8853    897de0                  mov dword ptr [ebp-20], edi
'006c8856    897ddc                  mov dword ptr [ebp-24], edi
'006c8859    897dd8                  mov dword ptr [ebp-28], edi
'006c885c    897dd4                  mov dword ptr [ebp-2c], edi
'006c885f    897dd0                  mov dword ptr [ebp-30], edi
'006c8862    897dcc                  mov dword ptr [ebp-34], edi
'006c8865    897dc8                  mov dword ptr [ebp-38], edi
'006c8868    897dc4                  mov dword ptr [ebp-3c], edi
'006c886b    897dc0                  mov dword ptr [ebp-40], edi
'006c886e    897dbc                  mov dword ptr [ebp-44], edi
'006c8871    897db8                  mov dword ptr [ebp-48], edi
'006c8874    897db4                  mov dword ptr [ebp-4c], edi
'006c8877    897da4                  mov dword ptr [ebp-5c], edi
'006c887a    897d94                  mov dword ptr [ebp-6c], edi
'006c887d    897d84                  mov dword ptr [ebp-7c], edi
'006c8880    89bd74ffffff            mov dword ptr [ebp+ffffff74], edi
'006c8886    89bd64ffffff            mov dword ptr [ebp+ffffff64], edi
'006c888c    89bd54ffffff            mov dword ptr [ebp+ffffff54], edi
'006c8892    89bd44ffffff            mov dword ptr [ebp+ffffff44], edi
'006c8898    89bd34ffffff            mov dword ptr [ebp+ffffff34], edi
'006c889e    89bd24ffffff            mov dword ptr [ebp+ffffff24], edi
'006c88a4    89bd04ffffff            mov dword ptr [ebp+ffffff04], edi
'006c88aa    89bdf4feffff            mov dword ptr [ebp+fffffef4], edi
'006c88b0    89bdd4feffff            mov dword ptr [ebp+fffffed4], edi
'006c88b6    89bdb4feffff            mov dword ptr [ebp+fffffeb4], edi
'006c88bc    89bda0feffff            mov dword ptr [ebp+fffffea0], edi
'006c88c2    89bd9cfeffff            mov dword ptr [ebp+fffffe9c], edi
'006c88c8    89bd64feffff            mov dword ptr [ebp+fffffe64], edi
'006c88ce    ff91fc020000            call dword ptr [ecx+000002fc]

' *** Reference to "__vbaObjSet"
'006c88d4    8b1db4104000            mov ebx, dword ptr [004010b4]
'006c88da    50                      push eax
'006c88db    8d55cc                  lea edx, dword ptr [ebp-34]
'006c88de    52                      push edx
'006c88df    ffd3                    call ebx
Set var_43 = Nothing
'006c88e1    8b08                    mov ecx, dword ptr [eax]
'006c88e3    8d55e4                  lea edx, dword ptr [ebp-1c]
'006c88e6    52                      push edx
'006c88e7    50                      push eax
'006c88e8    898598feffff            mov dword ptr [ebp+fffffe98], eax
'006c88ee    ff91a8000000            call dword ptr [ecx+000000a8]
'006c88f4    dbe2                    fnclex
'006c88f6    3bc7                    cmp eax, edi
'006c88f8    7d18                    jge 6c8912

If (var_43 < 0) Then
'006c88fa    8b8d98feffff            mov ecx, dword ptr [ebp+fffffe98]
'006c8900    68a8000000              push 000000a8
'006c8905    681cb94300              push 0043b91c
'006c890a    51                      push ecx
'006c890b    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c890c    ff1580104000            call dword ptr [00401080]
    
End If
'006c8912    8b55e4                  mov edx, dword ptr [ebp-1c]
'006c8915    52                      push edx
'006c8916    68cc134300              push 004313cc

' *** Reference to "__vbaStrCmp"
'006c891b    ff1530114000            call dword ptr [00401130]
'006c8921    8bf8                    mov edi, eax
'006c8923    f7df                    neg edi
'006c8925    1bff                    sbb edi, edi
'006c8927    f7df                    neg edi
'006c8929    8d4de4                  lea ecx, dword ptr [ebp-1c]
'006c892c    f7df                    neg edi

' *** Reference to "__vbaFreeStr"
'006c892e    ff150c134000            call dword ptr [0040130c]
'#Cleanup(var_40)
'006c8934    8d4dcc                  lea ecx, dword ptr [ebp-34]

' *** Reference to "__vbaFreeObj"
'006c8937    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_43)
'006c893d    6685ff                  test edi, edi
'006c8940    0f846a340000            je 6cbdb0

If (((vbNullString) <> (vbNullChar))) Then
'006c8946    8b06                    mov eax, dword ptr [esi]
'006c8948    56                      push esi
'006c8949    ff90fc020000            call dword ptr [eax+000002fc]
'006c894f    50                      push eax
'006c8950    8d4dcc                  lea ecx, dword ptr [ebp-34]
'006c8953    51                      push ecx
'006c8954    ffd3                    call ebx
    Set var_43 = Nothing
'006c8956    8bf8                    mov edi, eax
'006c8958    8b17                    mov edx, dword ptr [edi]
'006c895a    8d45e4                  lea eax, dword ptr [ebp-1c]
'006c895d    50                      push eax
'006c895e    57                      push edi
'006c895f    ff92a8000000            call dword ptr [edx+000000a8]
'006c8965    dbe2                    fnclex
'006c8967    85c0                    test ax, ax
'006c8969    7d12                    jge 6c897d
'006c896b    68a8000000              push 000000a8
'006c8970    681cb94300              push 0043b91c
'006c8975    57                      push edi
'006c8976    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c8977    ff1580104000            call dword ptr [00401080]
'006c897d    8b55e4                  mov edx, dword ptr [ebp-1c]

' *** Reference to "__vbaStrMove"
'006c8980    8b3dd0124000            mov edi, dword ptr [004012d0]
'006c8986    33db                    xor ebx, ebx
    var_num2 = Empty
'006c8988    8d4de0                  lea ecx, dword ptr [ebp-20]
'006c898b    895de4                  mov dword ptr [ebp-1c], ebx
'006c898e    ffd7                    call edi
'006c8990    8d4de0                  lea ecx, dword ptr [ebp-20]
'006c8993    51                      push ecx
'006c8994    e857b2e2ff              call 4f3bf0
    Call sub_4F3BF0()
'006c8999    8bd0                    mov edx, eax
'006c899b    8d4dd0                  lea ecx, dword ptr [ebp-30]
'006c899e    ffd7                    call edi
'006c89a0    8b55d0                  mov edx, dword ptr [ebp-30]
'006c89a3    899554feffff            mov dword ptr [ebp+fffffe54], edx
'006c89a9    8b154cb07200            mov edx, dword ptr [0072b04c]
'006c89af    b804000280              mov eax, 80020004
'006c89b4    89854cffffff            mov dword ptr [ebp+ffffff4c], eax
'006c89ba    b90a000000              mov ecx, 0000000a
'006c89bf    898d44ffffff            mov dword ptr [ebp+ffffff44], ecx
'006c89c5    c7856cffffff02000000    mov dword ptr [ebp+ffffff6c], 00000002
'006c89cf    c78564ffffff03000000    mov dword ptr [ebp+ffffff64], 00000003
'006c89d9    895dd0                  mov dword ptr [ebp-30], ebx
'006c89dc    8b1a                    mov ebx, dword ptr [edx]
'006c89de    898d54ffffff            mov dword ptr [ebp+ffffff54], ecx
'006c89e4    8d55c8                  lea edx, dword ptr [ebp-38]
'006c89e7    52                      push edx
'006c89e8    83ec10                  sub esp, 10
'006c89eb    8bd4                    mov edx, esp
'006c89ed    890a                    mov dword ptr [edx], ecx
'006c89ef    8b8d48ffffff            mov ecx, dword ptr [ebp+ffffff48]
'006c89f5    894a04                  mov dword ptr [edx+04], ecx
'006c89f8    894208                  mov dword ptr [edx+08], eax
'006c89fb    8bf8                    mov edi, eax
'006c89fd    8b8550ffffff            mov eax, dword ptr [ebp+ffffff50]
'006c8a03    89420c                  mov dword ptr [edx+0c], eax
'006c8a06    8b9554ffffff            mov edx, dword ptr [ebp+ffffff54]
'006c8a0c    8b8558ffffff            mov eax, dword ptr [ebp+ffffff58]
'006c8a12    83ec10                  sub esp, 10
'006c8a15    8bcc                    mov ecx, esp
'006c8a17    8911                    mov dword ptr [ecx], edx
'006c8a19    8b9560ffffff            mov edx, dword ptr [ebp+ffffff60]
'006c8a1f    894104                  mov dword ptr [ecx+04], eax
'006c8a22    897908                  mov dword ptr [ecx+08], edi

' *** Reference to "__vbaStrMove"
'006c8a25    8b3dd0124000            mov edi, dword ptr [004012d0]
'006c8a2b    89510c                  mov dword ptr [ecx+0c], edx
'006c8a2e    8b8d64ffffff            mov ecx, dword ptr [ebp+ffffff64]
'006c8a34    8b9568ffffff            mov edx, dword ptr [ebp+ffffff68]
'006c8a3a    83ec10                  sub esp, 10
'006c8a3d    8bc4                    mov eax, esp
'006c8a3f    8908                    mov dword ptr [eax], ecx
'006c8a41    8b8d6cffffff            mov ecx, dword ptr [ebp+ffffff6c]
'006c8a47    895004                  mov dword ptr [eax+04], edx
'006c8a4a    8b9570ffffff            mov edx, dword ptr [ebp+ffffff70]
'006c8a50    894808                  mov dword ptr [eax+08], ecx
'006c8a53    89500c                  mov dword ptr [eax+0c], edx
'006c8a56    8b9554feffff            mov edx, dword ptr [ebp+fffffe54]
'006c8a5c    6810cb4400              push 0044cb10
'006c8a61    8d4ddc                  lea ecx, dword ptr [ebp-24]
'006c8a64    ffd7                    call edi
'006c8a66    50                      push eax

' *** Reference to "__vbaStrCat"
'006c8a67    ff1570104000            call dword ptr [00401070]
    var_49 = ("select * from sort where nom='") & (vbNullString)
'006c8a6d    8bd0                    mov edx, eax
'006c8a6f    8d4dd8                  lea ecx, dword ptr [ebp-28]
'006c8a72    ffd7                    call edi
'006c8a74    50                      push eax
'006c8a75    6854a44300              push 0043a454

' *** Reference to "__vbaStrCat"
'006c8a7a    ff1570104000            call dword ptr [00401070]
    var_63 = (var_49) & ("'")
'006c8a80    8bd0                    mov edx, eax
'006c8a82    8d4dd4                  lea ecx, dword ptr [ebp-2c]
'006c8a85    ffd7                    call edi
'006c8a87    50                      push eax
'006c8a88    a14cb07200              mov ax, word ptr [0072b04c]
'006c8a8d    50                      push eax
'006c8a8e    ff93bc000000            call dword ptr [ebx+000000bc]
'006c8a94    dbe2                    fnclex
'006c8a96    85c0                    test ax, ax
'006c8a98    7d18                    jge 6c8ab2
'006c8a9a    8b0d4cb07200            mov ecx, dword ptr [0072b04c]
'006c8aa0    68bc000000              push 000000bc
'006c8aa5    68301f4300              push 00431f30
'006c8aaa    51                      push ecx
'006c8aab    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c8aac    ff1580104000            call dword ptr [00401080]
'006c8ab2    8b45c8                  mov eax, dword ptr [ebp-38]

' *** Reference to "__vbaObjSet"
'006c8ab5    8b1db4104000            mov ebx, dword ptr [004010b4]
'006c8abb    50                      push eax
'006c8abc    8d55e8                  lea edx, dword ptr [ebp-18]
'006c8abf    52                      push edx
'006c8ac0    c745c800000000          mov dword ptr [ebp-38], 00000000
'006c8ac7    ffd3                    call ebx
    Set var_41 = Nothing
'006c8ac9    8d45d0                  lea eax, dword ptr [ebp-30]
'006c8acc    50                      push eax
'006c8acd    8d4dd4                  lea ecx, dword ptr [ebp-2c]
'006c8ad0    51                      push ecx
'006c8ad1    8d55d8                  lea edx, dword ptr [ebp-28]
'006c8ad4    52                      push edx
'006c8ad5    8d45dc                  lea eax, dword ptr [ebp-24]
'006c8ad8    50                      push eax
'006c8ad9    8d4de0                  lea ecx, dword ptr [ebp-20]
'006c8adc    51                      push ecx
'006c8add    6a05                    push 05

' *** Reference to "__vbaFreeStrList"
'006c8adf    ff155c124000            call dword ptr [0040125c]
    '#Cleanup( 0, 0, -4500, -4504, __vbaObjSet)
'006c8ae5    83c418                  add esp, 18
'006c8ae8    8d4dcc                  lea ecx, dword ptr [ebp-34]

' *** Reference to "__vbaFreeObj"
'006c8aeb    ff1508134000            call dword ptr [00401308]
    '#Cleanup(var_43)
'006c8af1    8b45e8                  mov eax, dword ptr [ebp-18]
'006c8af4    8b10                    mov edx, dword ptr [eax]
'006c8af6    8d8da0feffff            lea ecx, dword ptr [ebp+fffffea0]
'006c8afc    51                      push ecx
'006c8afd    50                      push eax
'006c8afe    ff5234                  call dword ptr [edx+34]
'006c8b01    dbe2                    fnclex
'006c8b03    85c0                    test ax, ax
'006c8b05    7d12                    jge 6c8b19
'006c8b07    8b55e8                  mov edx, dword ptr [ebp-18]
'006c8b0a    6a34                    push 34
'006c8b0c    6830314300              push 00433130
'006c8b11    52                      push edx
'006c8b12    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c8b13    ff1580104000            call dword ptr [00401080]
'006c8b19    6683bda0feffff00        cmp word ptr [ebp+fffffea0], 00
'006c8b21    0f8562320000            jne 6cbd89
    
    If (    0 = 0) Then
'006c8b27    8b06                    mov eax, dword ptr [esi]
'006c8b29    56                      push esi
'006c8b2a    ff9094030000            call dword ptr [eax+00000394]
'006c8b30    50                      push eax
'006c8b31    8d4dc4                  lea ecx, dword ptr [ebp-3c]
'006c8b34    51                      push ecx
'006c8b35    ffd3                    call ebx
    Set var_9 = Nothing
'006c8b37    8d4dcc                  lea ecx, dword ptr [ebp-34]
'006c8b3a    898584feffff            mov dword ptr [ebp+fffffe84], eax
'006c8b40    8b45e8                  mov eax, dword ptr [ebp-18]
'006c8b43    8b10                    mov edx, dword ptr [eax]
'006c8b45    51                      push ecx
'006c8b46    50                      push eax
'006c8b47    ff92b4000000            call dword ptr [edx+000000b4]
'006c8b4d    dbe2                    fnclex
'006c8b4f    85c0                    test ax, ax
'006c8b51    7d15                    jge 6c8b68
'006c8b53    8b55e8                  mov edx, dword ptr [ebp-18]
'006c8b56    68b4000000              push 000000b4
'006c8b5b    6830314300              push 00433130
'006c8b60    52                      push edx
'006c8b61    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c8b62    ff1580104000            call dword ptr [00401080]
'006c8b68    8b45cc                  mov eax, dword ptr [ebp-34]
'006c8b6b    8d7dc8                  lea edi, dword ptr [ebp-38]
'006c8b6e    57                      push edi
'006c8b6f    83ec10                  sub esp, 10
'006c8b72    b908000000              mov ecx, 00000008
'006c8b77    898d64ffffff            mov dword ptr [ebp+ffffff64], ecx
'006c8b7d    8bfc                    mov edi, esp
'006c8b7f    890f                    mov dword ptr [edi], ecx
'006c8b81    8b8d68ffffff            mov ecx, dword ptr [ebp+ffffff68]
'006c8b87    894f04                  mov dword ptr [edi+04], ecx
'006c8b8a    c7856cffffff64b14300    mov dword ptr [ebp+ffffff6c], 0043b164
'006c8b94    8b8d6cffffff            mov ecx, dword ptr [ebp+ffffff6c]
'006c8b9a    8b10                    mov edx, dword ptr [eax]
'006c8b9c    894f08                  mov dword ptr [edi+08], ecx
'006c8b9f    8b8d70ffffff            mov ecx, dword ptr [ebp+ffffff70]
'006c8ba5    50                      push eax
'006c8ba6    898594feffff            mov dword ptr [ebp+fffffe94], eax
'006c8bac    894f0c                  mov dword ptr [edi+0c], ecx
'006c8baf    ff5230                  call dword ptr [edx+30]
'006c8bb2    dbe2                    fnclex
'006c8bb4    85c0                    test ax, ax
'006c8bb6    7d15                    jge 6c8bcd
'006c8bb8    8b9594feffff            mov edx, dword ptr [ebp+fffffe94]
'006c8bbe    6a30                    push 30
'006c8bc0    68d8304300              push 004330d8
'006c8bc5    52                      push edx
'006c8bc6    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c8bc7    ff1580104000            call dword ptr [00401080]
'006c8bcd    8b45c8                  mov eax, dword ptr [ebp-38]
'006c8bd0    8b08                    mov ecx, dword ptr [eax]
'006c8bd2    8d55a4                  lea edx, dword ptr [ebp-5c]
'006c8bd5    52                      push edx
'006c8bd6    50                      push eax
'006c8bd7    8bf8                    mov edi, eax
'006c8bd9    ff5144                  call dword ptr [ecx+44]
'006c8bdc    dbe2                    fnclex
'006c8bde    85c0                    test ax, ax
'006c8be0    7d0f                    jge 6c8bf1
'006c8be2    6a44                    push 44
'006c8be4    6880304300              push 00433080
'006c8be9    57                      push edi
'006c8bea    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c8beb    ff1580104000            call dword ptr [00401080]
'006c8bf1    8b8584feffff            mov eax, dword ptr [ebp+fffffe84]
'006c8bf7    8b38                    mov edi, dword ptr [eax]
'006c8bf9    8d4da4                  lea ecx, dword ptr [ebp-5c]
'006c8bfc    51                      push ecx

' *** Reference to "__vbaStrVarMove"
'006c8bfd    ff153c104000            call dword ptr [0040103c]
'006c8c03    8bd0                    mov edx, eax
'006c8c05    8d4de4                  lea ecx, dword ptr [ebp-1c]

' *** Reference to "__vbaStrMove"
'006c8c08    ff15d0124000            call dword ptr [004012d0]
'006c8c0e    8bd7                    mov edx, edi
'006c8c10    8bbd84feffff            mov edi, dword ptr [ebp+fffffe84]
'006c8c16    50                      push eax
'006c8c17    57                      push edi
'006c8c18    ff5254                  call dword ptr [edx+54]
'006c8c1b    dbe2                    fnclex
'006c8c1d    85c0                    test ax, ax
'006c8c1f    7d0f                    jge 6c8c30
'006c8c21    6a54                    push 54
'006c8c23    683ce04300              push 0043e03c
'006c8c28    57                      push edi
'006c8c29    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c8c2a    ff1580104000            call dword ptr [00401080]
'006c8c30    8d4de4                  lea ecx, dword ptr [ebp-1c]

' *** Reference to "__vbaFreeStr"
'006c8c33    ff150c134000            call dword ptr [0040130c]
    '#Cleanup(var_40)
'006c8c39    8d45c4                  lea eax, dword ptr [ebp-3c]
'006c8c3c    50                      push eax
'006c8c3d    8d4dc8                  lea ecx, dword ptr [ebp-38]
'006c8c40    51                      push ecx
'006c8c41    8d55cc                  lea edx, dword ptr [ebp-34]
'006c8c44    52                      push edx
'006c8c45    6a03                    push 03

' *** Reference to "__vbaFreeObjList"
'006c8c47    ff154c104000            call dword ptr [0040104c]
    '#Cleanup( var_43, var_46, var_9)
'006c8c4d    83c410                  add esp, 10
'006c8c50    8d4da4                  lea ecx, dword ptr [ebp-5c]

' *** Reference to "__vbaFreeVar"
'006c8c53    ff1528104000            call dword ptr [00401028]
    '#Cleanup(var_83)
'006c8c59    8b06                    mov eax, dword ptr [esi]
'006c8c5b    56                      push esi
'006c8c5c    ff90fc020000            call dword ptr [eax+000002fc]
'006c8c62    50                      push eax
'006c8c63    8d4dc8                  lea ecx, dword ptr [ebp-38]
'006c8c66    51                      push ecx
'006c8c67    ffd3                    call ebx
    Set var_46 = Nothing
'006c8c69    8b16                    mov edx, dword ptr [esi]
'006c8c6b    56                      push esi
'006c8c6c    898590feffff            mov dword ptr [ebp+fffffe90], eax
'006c8c72    ff9294030000            call dword ptr [edx+00000394]
'006c8c78    50                      push eax
'006c8c79    8d45cc                  lea eax, dword ptr [ebp-34]
'006c8c7c    50                      push eax
'006c8c7d    ffd3                    call ebx
    Set var_43 = Nothing
'006c8c7f    8d55e4                  lea edx, dword ptr [ebp-1c]
'006c8c82    8bf8                    mov edi, eax
'006c8c84    8b0f                    mov ecx, dword ptr [edi]
'006c8c86    52                      push edx
'006c8c87    57                      push edi
'006c8c88    ff5150                  call dword ptr [ecx+50]
'006c8c8b    dbe2                    fnclex
'006c8c8d    85c0                    test ax, ax
'006c8c8f    7d0f                    jge 6c8ca0
'006c8c91    6a50                    push 50
'006c8c93    683ce04300              push 0043e03c
'006c8c98    57                      push edi
'006c8c99    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c8c9a    ff1580104000            call dword ptr [00401080]
'006c8ca0    8b4de4                  mov ecx, dword ptr [ebp-1c]
'006c8ca3    8bbd90feffff            mov edi, dword ptr [ebp+fffffe90]
'006c8ca9    8b07                    mov eax, dword ptr [edi]
'006c8cab    51                      push ecx
'006c8cac    57                      push edi
'006c8cad    ff90ac000000            call dword ptr [eax+000000ac]
'006c8cb3    dbe2                    fnclex
'006c8cb5    85c0                    test ax, ax
'006c8cb7    7d12                    jge 6c8ccb
'006c8cb9    68ac000000              push 000000ac
'006c8cbe    681cb94300              push 0043b91c
'006c8cc3    57                      push edi
'006c8cc4    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c8cc5    ff1580104000            call dword ptr [00401080]
'006c8ccb    8d4de4                  lea ecx, dword ptr [ebp-1c]

' *** Reference to "__vbaFreeStr"
'006c8cce    ff150c134000            call dword ptr [0040130c]
    '#Cleanup(var_40)
'006c8cd4    8d55c8                  lea edx, dword ptr [ebp-38]
'006c8cd7    52                      push edx
'006c8cd8    8d45cc                  lea eax, dword ptr [ebp-34]
'006c8cdb    50                      push eax
'006c8cdc    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'006c8cde    ff154c104000            call dword ptr [0040104c]
    '#Cleanup( var_43, var_46)
'006c8ce4    8b45e8                  mov eax, dword ptr [ebp-18]
'006c8ce7    8b08                    mov ecx, dword ptr [eax]
'006c8ce9    83c40c                  add esp, 0c
'006c8cec    8d55cc                  lea edx, dword ptr [ebp-34]
'006c8cef    52                      push edx
'006c8cf0    50                      push eax
'006c8cf1    ff91b4000000            call dword ptr [ecx+000000b4]
'006c8cf7    dbe2                    fnclex
'006c8cf9    85c0                    test ax, ax
'006c8cfb    7d15                    jge 6c8d12
'006c8cfd    8b4de8                  mov ecx, dword ptr [ebp-18]
'006c8d00    68b4000000              push 000000b4
'006c8d05    6830314300              push 00433130
'006c8d0a    51                      push ecx
'006c8d0b    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c8d0c    ff1580104000            call dword ptr [00401080]
'006c8d12    8b45cc                  mov eax, dword ptr [ebp-34]
'006c8d15    8d7dc8                  lea edi, dword ptr [ebp-38]
'006c8d18    57                      push edi
'006c8d19    83ec10                  sub esp, 10
'006c8d1c    b908000000              mov ecx, 00000008
'006c8d21    898d64ffffff            mov dword ptr [ebp+ffffff64], ecx
'006c8d27    8bfc                    mov edi, esp
'006c8d29    890f                    mov dword ptr [edi], ecx
'006c8d2b    8b8d68ffffff            mov ecx, dword ptr [ebp+ffffff68]
'006c8d31    894f04                  mov dword ptr [edi+04], ecx
'006c8d34    c7856cffffff0ca94300    mov dword ptr [ebp+ffffff6c], 0043a90c
'006c8d3e    8b8d6cffffff            mov ecx, dword ptr [ebp+ffffff6c]
'006c8d44    8b10                    mov edx, dword ptr [eax]
'006c8d46    894f08                  mov dword ptr [edi+08], ecx
'006c8d49    8b8d70ffffff            mov ecx, dword ptr [ebp+ffffff70]
'006c8d4f    50                      push eax
'006c8d50    898594feffff            mov dword ptr [ebp+fffffe94], eax
'006c8d56    894f0c                  mov dword ptr [edi+0c], ecx
'006c8d59    ff5230                  call dword ptr [edx+30]
'006c8d5c    dbe2                    fnclex
'006c8d5e    85c0                    test ax, ax
'006c8d60    7d15                    jge 6c8d77
'006c8d62    8b9594feffff            mov edx, dword ptr [ebp+fffffe94]
'006c8d68    6a30                    push 30
'006c8d6a    68d8304300              push 004330d8
'006c8d6f    52                      push edx
'006c8d70    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c8d71    ff1580104000            call dword ptr [00401080]
'006c8d77    8b45c8                  mov eax, dword ptr [ebp-38]
'006c8d7a    8b08                    mov ecx, dword ptr [eax]
'006c8d7c    8d55a4                  lea edx, dword ptr [ebp-5c]
'006c8d7f    52                      push edx
'006c8d80    50                      push eax
'006c8d81    8bf8                    mov edi, eax
'006c8d83    ff5144                  call dword ptr [ecx+44]
'006c8d86    dbe2                    fnclex
'006c8d88    85c0                    test ax, ax
'006c8d8a    7d0f                    jge 6c8d9b
'006c8d8c    6a44                    push 44
'006c8d8e    6880304300              push 00433080
'006c8d93    57                      push edi
'006c8d94    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c8d95    ff1580104000            call dword ptr [00401080]
'006c8d9b    8d45a4                  lea eax, dword ptr [ebp-5c]
'006c8d9e    50                      push eax

' *** Reference to "__vbaStrVarMove"
'006c8d9f    ff153c104000            call dword ptr [0040103c]
'006c8da5    8bd0                    mov edx, eax
'006c8da7    8d4de4                  lea ecx, dword ptr [ebp-1c]

' *** Reference to "__vbaStrMove"
'006c8daa    ff15d0124000            call dword ptr [004012d0]
'006c8db0    50                      push eax

' *** Reference to "rtcR8ValFromBstr"
'006c8db1    ff1510134000            call dword ptr [00401310]

' *** Reference to "__vbaFpR8"
'006c8db7    ff15f8104000            call dword ptr [004010f8]
'006c8dbd    dc1d50514000            fcomp qword ptr [00405150]
'006c8dc3    dfe0                    fnstsw ax
'006c8dc5    f6c440                  test ah, 40
'006c8dc8    7507                    jne 6c8dd1
'006c8dca    bf01000000              mov edi, 00000001
'006c8dcf    eb02                    jmp 6c8dd3
'006c8dd1    33ff                    xor edi, edi
'006c8dd3    8d4de4                  lea ecx, dword ptr [ebp-1c]

' *** Reference to "__vbaFreeStr"
'006c8dd6    ff150c134000            call dword ptr [0040130c]
    '#Cleanup(var_40)
'006c8ddc    8d4dc8                  lea ecx, dword ptr [ebp-38]
'006c8ddf    51                      push ecx
'006c8de0    8d55cc                  lea edx, dword ptr [ebp-34]
'006c8de3    52                      push edx
'006c8de4    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'006c8de6    ff154c104000            call dword ptr [0040104c]
    '#Cleanup( var_43, var_46)
'006c8dec    83c40c                  add esp, 0c
'006c8def    8d4da4                  lea ecx, dword ptr [ebp-5c]

' *** Reference to "__vbaFreeVar"
'006c8df2    ff1528104000            call dword ptr [00401028]
    '#Cleanup(var_83)
'006c8df8    f7df                    neg edi
'006c8dfa    6685ff                  test edi, edi
'006c8dfd    0f84d3000000            je 6c8ed6
    
    If (    (3.5# <> Val(var_83))) Then
'006c8e03    8b06                    mov eax, dword ptr [esi]
'006c8e05    56                      push esi
'006c8e06    ff9094030000            call dword ptr [eax+00000394]
'006c8e0c    50                      push eax
'006c8e0d    8d4dc8                  lea ecx, dword ptr [ebp-38]
'006c8e10    51                      push ecx
'006c8e11    ffd3                    call ebx
    Set var_46 = Nothing
'006c8e13    8b16                    mov edx, dword ptr [esi]
'006c8e15    56                      push esi
'006c8e16    898590feffff            mov dword ptr [ebp+fffffe90], eax
'006c8e1c    ff9294030000            call dword ptr [edx+00000394]
'006c8e22    50                      push eax
'006c8e23    8d45cc                  lea eax, dword ptr [ebp-34]
'006c8e26    50                      push eax
'006c8e27    ffd3                    call ebx
    Set var_43 = Nothing
'006c8e29    8d55e4                  lea edx, dword ptr [ebp-1c]
'006c8e2c    8bf8                    mov edi, eax
'006c8e2e    8b0f                    mov ecx, dword ptr [edi]
'006c8e30    52                      push edx
'006c8e31    57                      push edi
'006c8e32    ff5150                  call dword ptr [ecx+50]
'006c8e35    dbe2                    fnclex
'006c8e37    85c0                    test ax, ax
'006c8e39    7d0f                    jge 6c8e4a
'006c8e3b    6a50                    push 50
'006c8e3d    683ce04300              push 0043e03c
'006c8e42    57                      push edi
'006c8e43    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c8e44    ff1580104000            call dword ptr [00401080]
'006c8e4a    8b4de4                  mov ecx, dword ptr [ebp-1c]
'006c8e4d    8b8590feffff            mov eax, dword ptr [ebp+fffffe90]
'006c8e53    8b38                    mov edi, dword ptr [eax]
'006c8e55    51                      push ecx
'006c8e56    688c394500              push 0045398c

' *** Reference to "__vbaStrCat"
'006c8e5b    ff1570104000            call dword ptr [00401070]
    var_49 = (vbNullString) & (" (version 3.0)")
'006c8e61    8bd0                    mov edx, eax
'006c8e63    8d4de0                  lea ecx, dword ptr [ebp-20]

' *** Reference to "__vbaStrMove"
'006c8e66    ff15d0124000            call dword ptr [004012d0]
'006c8e6c    8bd7                    mov edx, edi
'006c8e6e    8bbd90feffff            mov edi, dword ptr [ebp+fffffe90]
'006c8e74    50                      push eax
'006c8e75    57                      push edi
'006c8e76    ff5254                  call dword ptr [edx+54]
'006c8e79    dbe2                    fnclex
'006c8e7b    85c0                    test ax, ax
'006c8e7d    7d0f                    jge 6c8e8e
'006c8e7f    6a54                    push 54
'006c8e81    683ce04300              push 0043e03c
'006c8e86    57                      push edi
'006c8e87    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c8e88    ff1580104000            call dword ptr [00401080]
'006c8e8e    8d45e0                  lea eax, dword ptr [ebp-20]
'006c8e91    50                      push eax
'006c8e92    8d4de4                  lea ecx, dword ptr [ebp-1c]
'006c8e95    51                      push ecx
'006c8e96    6a02                    push 02

' *** Reference to "__vbaFreeStrList"
'006c8e98    ff155c124000            call dword ptr [0040125c]
    '#Cleanup( , -4516)
'006c8e9e    8d55c8                  lea edx, dword ptr [ebp-38]
'006c8ea1    52                      push edx
'006c8ea2    8d45cc                  lea eax, dword ptr [ebp-34]
'006c8ea5    50                      push eax
'006c8ea6    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'006c8ea8    ff154c104000            call dword ptr [0040104c]
    '#Cleanup( var_43, var_46)
'006c8eae    8b0e                    mov ecx, dword ptr [esi]
'006c8eb0    83c418                  add esp, 18
'006c8eb3    56                      push esi
'006c8eb4    ff9194030000            call dword ptr [ecx+00000394]
'006c8eba    50                      push eax
'006c8ebb    8d55cc                  lea edx, dword ptr [ebp-34]
'006c8ebe    52                      push edx
'006c8ebf    ffd3                    call ebx
    Set var_43 = 
'006c8ec1    8bf8                    mov edi, eax
'006c8ec3    8b07                    mov eax, dword ptr [edi]
'006c8ec5    68ff000000              push 000000ff
'006c8eca    57                      push edi
'006c8ecb    ff506c                  call dword ptr [eax+6c]
'006c8ece    dbe2                    fnclex
'006c8ed0    85c0                    test ax, ax
'006c8ed2    7d34                    jge 6c8f08
'006c8ed4    eb23                    jmp 6c8ef9
    
End If
'006c8ed6    8b0e                    mov ecx, dword ptr [esi]
'006c8ed8    56                      push esi
'006c8ed9    ff9194030000            call dword ptr [ecx+00000394]
'006c8edf    50                      push eax
'006c8ee0    8d55cc                  lea edx, dword ptr [ebp-34]
'006c8ee3    52                      push edx
'006c8ee4    ffd3                    call ebx
Set var_43 = (3.5# [:#] Val(var_83))
'006c8ee6    8bf8                    mov edi, eax
'006c8ee8    8b07                    mov eax, dword ptr [edi]
'006c8eea    6812000080              push 80000012
'006c8eef    57                      push edi
'006c8ef0    ff506c                  call dword ptr [eax+6c]
'006c8ef3    dbe2                    fnclex
'006c8ef5    85c0                    test ax, ax
'006c8ef7    7d0f                    jge 6c8f08
'006c8ef9    6a6c                    push 6c
'006c8efb    683ce04300              push 0043e03c
'006c8f00    57                      push edi
'006c8f01    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c8f02    ff1580104000            call dword ptr [00401080]
'006c8f08    8d4dcc                  lea ecx, dword ptr [ebp-34]

' *** Reference to "__vbaFreeObj"
'006c8f0b    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_43)
'006c8f11    8b45e8                  mov eax, dword ptr [ebp-18]
'006c8f14    8b08                    mov ecx, dword ptr [eax]
'006c8f16    8d55cc                  lea edx, dword ptr [ebp-34]
'006c8f19    52                      push edx
'006c8f1a    50                      push eax
'006c8f1b    ff91b4000000            call dword ptr [ecx+000000b4]
'006c8f21    dbe2                    fnclex
'006c8f23    85c0                    test ax, ax
'006c8f25    7d15                    jge 6c8f3c
'006c8f27    8b4de8                  mov ecx, dword ptr [ebp-18]
'006c8f2a    68b4000000              push 000000b4
'006c8f2f    6830314300              push 00433130
'006c8f34    51                      push ecx
'006c8f35    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c8f36    ff1580104000            call dword ptr [00401080]
'006c8f3c    8b45cc                  mov eax, dword ptr [ebp-34]
'006c8f3f    8d7dc8                  lea edi, dword ptr [ebp-38]
'006c8f42    57                      push edi
'006c8f43    83ec10                  sub esp, 10
'006c8f46    b908000000              mov ecx, 00000008
'006c8f4b    898d64ffffff            mov dword ptr [ebp+ffffff64], ecx
'006c8f51    8bfc                    mov edi, esp
'006c8f53    890f                    mov dword ptr [edi], ecx
'006c8f55    8b8d68ffffff            mov ecx, dword ptr [ebp+ffffff68]
'006c8f5b    894f04                  mov dword ptr [edi+04], ecx
'006c8f5e    c7856cffffffbcb24300    mov dword ptr [ebp+ffffff6c], 0043b2bc
'006c8f68    8b8d6cffffff            mov ecx, dword ptr [ebp+ffffff6c]
'006c8f6e    8b10                    mov edx, dword ptr [eax]
'006c8f70    894f08                  mov dword ptr [edi+08], ecx
'006c8f73    8b8d70ffffff            mov ecx, dword ptr [ebp+ffffff70]
'006c8f79    50                      push eax
'006c8f7a    898594feffff            mov dword ptr [ebp+fffffe94], eax
'006c8f80    894f0c                  mov dword ptr [edi+0c], ecx
'006c8f83    ff5230                  call dword ptr [edx+30]
'006c8f86    dbe2                    fnclex
'006c8f88    85c0                    test ax, ax
'006c8f8a    7d15                    jge 6c8fa1
'006c8f8c    8b9594feffff            mov edx, dword ptr [ebp+fffffe94]
'006c8f92    6a30                    push 30
'006c8f94    68d8304300              push 004330d8
'006c8f99    52                      push edx
'006c8f9a    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c8f9b    ff1580104000            call dword ptr [00401080]
'006c8fa1    8b45c8                  mov eax, dword ptr [ebp-38]
'006c8fa4    8945ac                  mov dword ptr [ebp-54], eax
'006c8fa7    8d45a4                  lea eax, dword ptr [ebp-5c]
'006c8faa    50                      push eax
'006c8fab    c745c800000000          mov dword ptr [ebp-38], 00000000
'006c8fb2    c745a409000000          mov dword ptr [ebp-5c], 00000009

' *** Reference to "rtcIsNull"
'006c8fb9    ff1540114000            call dword ptr [00401140]
'006c8fbf    8b0e                    mov ecx, dword ptr [esi]
'006c8fc1    56                      push esi
'006c8fc2    8985a0feffff            mov dword ptr [ebp+fffffea0], eax
'006c8fc8    ff910c030000            call dword ptr [ecx+0000030c]
'006c8fce    50                      push eax
'006c8fcf    8d55bc                  lea edx, dword ptr [ebp-44]
'006c8fd2    52                      push edx
'006c8fd3    ffd3                    call ebx
Set var_58 = IsNull(var_46)
'006c8fd5    8d55c4                  lea edx, dword ptr [ebp-3c]
'006c8fd8    898580feffff            mov dword ptr [ebp+fffffe80], eax
'006c8fde    8b45e8                  mov eax, dword ptr [ebp-18]
'006c8fe1    8b08                    mov ecx, dword ptr [eax]
'006c8fe3    52                      push edx
'006c8fe4    50                      push eax
'006c8fe5    ff91b4000000            call dword ptr [ecx+000000b4]
'006c8feb    dbe2                    fnclex
'006c8fed    85c0                    test ax, ax
'006c8fef    7d15                    jge 6c9006
'006c8ff1    8b4de8                  mov ecx, dword ptr [ebp-18]
'006c8ff4    68b4000000              push 000000b4
'006c8ff9    6830314300              push 00433130
'006c8ffe    51                      push ecx
'006c8fff    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c9000    ff1580104000            call dword ptr [00401080]
'006c9006    8b45c4                  mov eax, dword ptr [ebp-3c]
'006c9009    8b10                    mov edx, dword ptr [eax]
'006c900b    8d7dc0                  lea edi, dword ptr [ebp-40]
'006c900e    57                      push edi
'006c900f    83ec10                  sub esp, 10
'006c9012    8bfc                    mov edi, esp
'006c9014    b908000000              mov ecx, 00000008
'006c9019    890f                    mov dword ptr [edi], ecx
'006c901b    8b8d58ffffff            mov ecx, dword ptr [ebp+ffffff58]
'006c9021    894f04                  mov dword ptr [edi+04], ecx
'006c9024    c7855cffffffbcb24300    mov dword ptr [ebp+ffffff5c], 0043b2bc
'006c902e    8b8d5cffffff            mov ecx, dword ptr [ebp+ffffff5c]
'006c9034    894f08                  mov dword ptr [edi+08], ecx
'006c9037    8b8d60ffffff            mov ecx, dword ptr [ebp+ffffff60]
'006c903d    50                      push eax
'006c903e    898588feffff            mov dword ptr [ebp+fffffe88], eax
'006c9044    894f0c                  mov dword ptr [edi+0c], ecx
'006c9047    ff5230                  call dword ptr [edx+30]
'006c904a    dbe2                    fnclex
'006c904c    85c0                    test ax, ax
'006c904e    7d15                    jge 6c9065
'006c9050    8b9588feffff            mov edx, dword ptr [ebp+fffffe88]
'006c9056    6a30                    push 30
'006c9058    68d8304300              push 004330d8
'006c905d    52                      push edx
'006c905e    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c905f    ff1580104000            call dword ptr [00401080]
'006c9065    8b45c0                  mov eax, dword ptr [ebp-40]
'006c9068    8d9534ffffff            lea edx, dword ptr [ebp+ffffff34]
'006c906e    8d4d94                  lea ecx, dword ptr [ebp-6c]
'006c9071    c745c000000000          mov dword ptr [ebp-40], 00000000
'006c9078    89458c                  mov dword ptr [ebp-74], eax
'006c907b    c7458409000000          mov dword ptr [ebp-7c], 00000009
'006c9082    c7853cffffffcc134300    mov dword ptr [ebp+ffffff3c], 004313cc
'006c908c    c78534ffffff08000000    mov dword ptr [ebp+ffffff34], 00000008

' *** Reference to "__vbaVarDup"
'006c9096    ff1598124000            call dword ptr [00401298]
var_121 = (vbNullChar)
'006c909c    668b85a0feffff          mov ax, word ptr [ebp+fffffea0]
'006c90a3    8d4d84                  lea ecx, dword ptr [ebp-7c]
'006c90a6    51                      push ecx
'006c90a7    8d5594                  lea edx, dword ptr [ebp-6c]
'006c90aa    6689854cffffff          mov word ptr [ebp+ffffff4c], ax
'006c90b1    52                      push edx
'006c90b2    8d8544ffffff            lea eax, dword ptr [ebp+ffffff44]
'006c90b8    50                      push eax
'006c90b9    8d8d74ffffff            lea ecx, dword ptr [ebp+ffffff74]
'006c90bf    51                      push ecx
'006c90c0    c78544ffffff0b000000    mov dword ptr [ebp+ffffff44], 0000000b

' *** Reference to "rtcImmediateIf"
'006c90ca    ff154c124000            call dword ptr [0040124c]
'006c90d0    8b9580feffff            mov edx, dword ptr [ebp+fffffe80]
'006c90d6    8b3a                    mov edi, dword ptr [edx]
'006c90d8    8d8574ffffff            lea eax, dword ptr [ebp+ffffff74]
'006c90de    50                      push eax
'006c90df    8d4de4                  lea ecx, dword ptr [ebp-1c]
'006c90e2    51                      push ecx

' *** Reference to "__vbaStrVarVal"
'006c90e3    ff15fc114000            call dword ptr [004011fc]
'006c90e9    8bd7                    mov edx, edi
'006c90eb    8bbd80feffff            mov edi, dword ptr [ebp+fffffe80]
'006c90f1    50                      push eax
'006c90f2    57                      push edi
'006c90f3    ff92a4000000            call dword ptr [edx+000000a4]
'006c90f9    dbe2                    fnclex
'006c90fb    85c0                    test ax, ax
'006c90fd    7d12                    jge 6c9111
'006c90ff    68a4000000              push 000000a4
'006c9104    68d00d4300              push 00430dd0
'006c9109    57                      push edi
'006c910a    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c910b    ff1580104000            call dword ptr [00401080]
'006c9111    8d4de4                  lea ecx, dword ptr [ebp-1c]

' *** Reference to "__vbaFreeStr"
'006c9114    ff150c134000            call dword ptr [0040130c]
'#Cleanup(var_40)
'006c911a    8d45bc                  lea eax, dword ptr [ebp-44]
'006c911d    50                      push eax
'006c911e    8d4dc4                  lea ecx, dword ptr [ebp-3c]
'006c9121    51                      push ecx
'006c9122    8d55cc                  lea edx, dword ptr [ebp-34]
'006c9125    52                      push edx
'006c9126    6a03                    push 03

' *** Reference to "__vbaFreeObjList"
'006c9128    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_43, var_9, var_58)
'006c912e    8d8574ffffff            lea eax, dword ptr [ebp+ffffff74]
'006c9134    50                      push eax
'006c9135    8d4d84                  lea ecx, dword ptr [ebp-7c]
'006c9138    51                      push ecx
'006c9139    8d5594                  lea edx, dword ptr [ebp-6c]
'006c913c    52                      push edx
'006c913d    8d8544ffffff            lea eax, dword ptr [ebp+ffffff44]
'006c9143    50                      push eax
'006c9144    8d4da4                  lea ecx, dword ptr [ebp-5c]
'006c9147    51                      push ecx
'006c9148    6a05                    push 05

' *** Reference to "__vbaFreeVarList"
'006c914a    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_83, var_125, var_121, var_119, var_91)
'006c9150    8b45e8                  mov eax, dword ptr [ebp-18]
'006c9153    8b10                    mov edx, dword ptr [eax]
'006c9155    83c428                  add esp, 28
'006c9158    8d4dcc                  lea ecx, dword ptr [ebp-34]
'006c915b    51                      push ecx
'006c915c    50                      push eax
'006c915d    ff92b4000000            call dword ptr [edx+000000b4]
'006c9163    dbe2                    fnclex
'006c9165    85c0                    test ax, ax
'006c9167    7d15                    jge 6c917e
'006c9169    8b55e8                  mov edx, dword ptr [ebp-18]
'006c916c    68b4000000              push 000000b4
'006c9171    6830314300              push 00433130
'006c9176    52                      push edx
'006c9177    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c9178    ff1580104000            call dword ptr [00401080]
'006c917e    8b45cc                  mov eax, dword ptr [ebp-34]
'006c9181    8d7dc8                  lea edi, dword ptr [ebp-38]
'006c9184    57                      push edi
'006c9185    83ec10                  sub esp, 10
'006c9188    b908000000              mov ecx, 00000008
'006c918d    898d64ffffff            mov dword ptr [ebp+ffffff64], ecx
'006c9193    8bfc                    mov edi, esp
'006c9195    890f                    mov dword ptr [edi], ecx
'006c9197    8b8d68ffffff            mov ecx, dword ptr [ebp+ffffff68]
'006c919d    894f04                  mov dword ptr [edi+04], ecx
'006c91a0    c7856cffffff10f34300    mov dword ptr [ebp+ffffff6c], 0043f310
'006c91aa    8b8d6cffffff            mov ecx, dword ptr [ebp+ffffff6c]
'006c91b0    8b10                    mov edx, dword ptr [eax]
'006c91b2    894f08                  mov dword ptr [edi+08], ecx
'006c91b5    8b8d70ffffff            mov ecx, dword ptr [ebp+ffffff70]
'006c91bb    50                      push eax
'006c91bc    898594feffff            mov dword ptr [ebp+fffffe94], eax
'006c91c2    894f0c                  mov dword ptr [edi+0c], ecx
'006c91c5    ff5230                  call dword ptr [edx+30]
'006c91c8    dbe2                    fnclex
'006c91ca    85c0                    test ax, ax
'006c91cc    7d15                    jge 6c91e3
'006c91ce    8b9594feffff            mov edx, dword ptr [ebp+fffffe94]
'006c91d4    6a30                    push 30
'006c91d6    68d8304300              push 004330d8
'006c91db    52                      push edx
'006c91dc    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c91dd    ff1580104000            call dword ptr [00401080]
'006c91e3    8b45c8                  mov eax, dword ptr [ebp-38]
'006c91e6    8945ac                  mov dword ptr [ebp-54], eax
'006c91e9    8d45a4                  lea eax, dword ptr [ebp-5c]
'006c91ec    50                      push eax
'006c91ed    c745c800000000          mov dword ptr [ebp-38], 00000000
'006c91f4    c745a409000000          mov dword ptr [ebp-5c], 00000009

' *** Reference to "rtcIsNull"
'006c91fb    ff1540114000            call dword ptr [00401140]
'006c9201    8b0e                    mov ecx, dword ptr [esi]
'006c9203    56                      push esi
'006c9204    8985a0feffff            mov dword ptr [ebp+fffffea0], eax
'006c920a    ff912c030000            call dword ptr [ecx+0000032c]
'006c9210    50                      push eax
'006c9211    8d55bc                  lea edx, dword ptr [ebp-44]
'006c9214    52                      push edx
'006c9215    ffd3                    call ebx
Set var_58 = IsNull(var_46)
'006c9217    8d55c4                  lea edx, dword ptr [ebp-3c]
'006c921a    898580feffff            mov dword ptr [ebp+fffffe80], eax
'006c9220    8b45e8                  mov eax, dword ptr [ebp-18]
'006c9223    8b08                    mov ecx, dword ptr [eax]
'006c9225    52                      push edx
'006c9226    50                      push eax
'006c9227    ff91b4000000            call dword ptr [ecx+000000b4]
'006c922d    dbe2                    fnclex
'006c922f    85c0                    test ax, ax
'006c9231    7d15                    jge 6c9248
'006c9233    8b4de8                  mov ecx, dword ptr [ebp-18]
'006c9236    68b4000000              push 000000b4
'006c923b    6830314300              push 00433130
'006c9240    51                      push ecx
'006c9241    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c9242    ff1580104000            call dword ptr [00401080]
'006c9248    8b45c4                  mov eax, dword ptr [ebp-3c]
'006c924b    8b10                    mov edx, dword ptr [eax]
'006c924d    8d7dc0                  lea edi, dword ptr [ebp-40]
'006c9250    57                      push edi
'006c9251    83ec10                  sub esp, 10
'006c9254    8bfc                    mov edi, esp
'006c9256    b908000000              mov ecx, 00000008
'006c925b    890f                    mov dword ptr [edi], ecx
'006c925d    8b8d58ffffff            mov ecx, dword ptr [ebp+ffffff58]
'006c9263    894f04                  mov dword ptr [edi+04], ecx
'006c9266    c7855cffffff10f34300    mov dword ptr [ebp+ffffff5c], 0043f310
'006c9270    8b8d5cffffff            mov ecx, dword ptr [ebp+ffffff5c]
'006c9276    894f08                  mov dword ptr [edi+08], ecx
'006c9279    8b8d60ffffff            mov ecx, dword ptr [ebp+ffffff60]
'006c927f    50                      push eax
'006c9280    898588feffff            mov dword ptr [ebp+fffffe88], eax
'006c9286    894f0c                  mov dword ptr [edi+0c], ecx
'006c9289    ff5230                  call dword ptr [edx+30]
'006c928c    dbe2                    fnclex
'006c928e    85c0                    test ax, ax
'006c9290    7d15                    jge 6c92a7
'006c9292    8b9588feffff            mov edx, dword ptr [ebp+fffffe88]
'006c9298    6a30                    push 30
'006c929a    68d8304300              push 004330d8
'006c929f    52                      push edx
'006c92a0    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c92a1    ff1580104000            call dword ptr [00401080]
'006c92a7    8b45c0                  mov eax, dword ptr [ebp-40]
'006c92aa    8d9534ffffff            lea edx, dword ptr [ebp+ffffff34]
'006c92b0    8d4d94                  lea ecx, dword ptr [ebp-6c]
'006c92b3    c745c000000000          mov dword ptr [ebp-40], 00000000
'006c92ba    89458c                  mov dword ptr [ebp-74], eax
'006c92bd    c7458409000000          mov dword ptr [ebp-7c], 00000009
'006c92c4    c7853cffffffcc134300    mov dword ptr [ebp+ffffff3c], 004313cc
'006c92ce    c78534ffffff08000000    mov dword ptr [ebp+ffffff34], 00000008

' *** Reference to "__vbaVarDup"
'006c92d8    ff1598124000            call dword ptr [00401298]
var_121 = (vbNullChar)
'006c92de    668b85a0feffff          mov ax, word ptr [ebp+fffffea0]
'006c92e5    8d4d84                  lea ecx, dword ptr [ebp-7c]
'006c92e8    51                      push ecx
'006c92e9    8d5594                  lea edx, dword ptr [ebp-6c]
'006c92ec    6689854cffffff          mov word ptr [ebp+ffffff4c], ax
'006c92f3    52                      push edx
'006c92f4    8d8544ffffff            lea eax, dword ptr [ebp+ffffff44]
'006c92fa    50                      push eax
'006c92fb    8d8d74ffffff            lea ecx, dword ptr [ebp+ffffff74]
'006c9301    51                      push ecx
'006c9302    c78544ffffff0b000000    mov dword ptr [ebp+ffffff44], 0000000b

' *** Reference to "rtcImmediateIf"
'006c930c    ff154c124000            call dword ptr [0040124c]
'006c9312    8b9580feffff            mov edx, dword ptr [ebp+fffffe80]
'006c9318    8b3a                    mov edi, dword ptr [edx]
'006c931a    8d8574ffffff            lea eax, dword ptr [ebp+ffffff74]
'006c9320    50                      push eax
'006c9321    8d4de4                  lea ecx, dword ptr [ebp-1c]
'006c9324    51                      push ecx

' *** Reference to "__vbaStrVarVal"
'006c9325    ff15fc114000            call dword ptr [004011fc]
'006c932b    8bd7                    mov edx, edi
'006c932d    8bbd80feffff            mov edi, dword ptr [ebp+fffffe80]
'006c9333    50                      push eax
'006c9334    57                      push edi
'006c9335    ff92a4000000            call dword ptr [edx+000000a4]
'006c933b    dbe2                    fnclex
'006c933d    85c0                    test ax, ax
'006c933f    7d12                    jge 6c9353
'006c9341    68a4000000              push 000000a4
'006c9346    68d00d4300              push 00430dd0
'006c934b    57                      push edi
'006c934c    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c934d    ff1580104000            call dword ptr [00401080]
'006c9353    8d4de4                  lea ecx, dword ptr [ebp-1c]

' *** Reference to "__vbaFreeStr"
'006c9356    ff150c134000            call dword ptr [0040130c]
'#Cleanup(var_40)
'006c935c    8d45bc                  lea eax, dword ptr [ebp-44]
'006c935f    50                      push eax
'006c9360    8d4dc4                  lea ecx, dword ptr [ebp-3c]
'006c9363    51                      push ecx
'006c9364    8d55cc                  lea edx, dword ptr [ebp-34]
'006c9367    52                      push edx
'006c9368    6a03                    push 03

' *** Reference to "__vbaFreeObjList"
'006c936a    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_43, var_9, var_58)
'006c9370    8d8574ffffff            lea eax, dword ptr [ebp+ffffff74]
'006c9376    50                      push eax
'006c9377    8d4d84                  lea ecx, dword ptr [ebp-7c]
'006c937a    51                      push ecx
'006c937b    8d5594                  lea edx, dword ptr [ebp-6c]
'006c937e    52                      push edx
'006c937f    8d8544ffffff            lea eax, dword ptr [ebp+ffffff44]
'006c9385    50                      push eax
'006c9386    8d4da4                  lea ecx, dword ptr [ebp-5c]
'006c9389    51                      push ecx
'006c938a    6a05                    push 05

' *** Reference to "__vbaFreeVarList"
'006c938c    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_83, var_125, var_121, var_119, var_91)
'006c9392    8b45e8                  mov eax, dword ptr [ebp-18]
'006c9395    8b10                    mov edx, dword ptr [eax]
'006c9397    83c428                  add esp, 28
'006c939a    8d4dcc                  lea ecx, dword ptr [ebp-34]
'006c939d    51                      push ecx
'006c939e    50                      push eax
'006c939f    ff92b4000000            call dword ptr [edx+000000b4]
'006c93a5    dbe2                    fnclex
'006c93a7    85c0                    test ax, ax
'006c93a9    7d15                    jge 6c93c0
'006c93ab    8b55e8                  mov edx, dword ptr [ebp-18]
'006c93ae    68b4000000              push 000000b4
'006c93b3    6830314300              push 00433130
'006c93b8    52                      push edx
'006c93b9    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c93ba    ff1580104000            call dword ptr [00401080]
'006c93c0    8b45cc                  mov eax, dword ptr [ebp-34]
'006c93c3    8d7dc8                  lea edi, dword ptr [ebp-38]
'006c93c6    57                      push edi
'006c93c7    83ec10                  sub esp, 10
'006c93ca    b908000000              mov ecx, 00000008
'006c93cf    898d64ffffff            mov dword ptr [ebp+ffffff64], ecx
'006c93d5    8bfc                    mov edi, esp
'006c93d7    890f                    mov dword ptr [edi], ecx
'006c93d9    8b8d68ffffff            mov ecx, dword ptr [ebp+ffffff68]
'006c93df    894f04                  mov dword ptr [edi+04], ecx
'006c93e2    c7856cffffff20f34300    mov dword ptr [ebp+ffffff6c], 0043f320
'006c93ec    8b8d6cffffff            mov ecx, dword ptr [ebp+ffffff6c]
'006c93f2    8b10                    mov edx, dword ptr [eax]
'006c93f4    894f08                  mov dword ptr [edi+08], ecx
'006c93f7    8b8d70ffffff            mov ecx, dword ptr [ebp+ffffff70]
'006c93fd    50                      push eax
'006c93fe    898594feffff            mov dword ptr [ebp+fffffe94], eax
'006c9404    894f0c                  mov dword ptr [edi+0c], ecx
'006c9407    ff5230                  call dword ptr [edx+30]
'006c940a    dbe2                    fnclex
'006c940c    85c0                    test ax, ax
'006c940e    7d15                    jge 6c9425
'006c9410    8b9594feffff            mov edx, dword ptr [ebp+fffffe94]
'006c9416    6a30                    push 30
'006c9418    68d8304300              push 004330d8
'006c941d    52                      push edx
'006c941e    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c941f    ff1580104000            call dword ptr [00401080]
'006c9425    8b45c8                  mov eax, dword ptr [ebp-38]
'006c9428    8945ac                  mov dword ptr [ebp-54], eax
'006c942b    8d45a4                  lea eax, dword ptr [ebp-5c]
'006c942e    50                      push eax
'006c942f    c745c800000000          mov dword ptr [ebp-38], 00000000
'006c9436    c745a409000000          mov dword ptr [ebp-5c], 00000009

' *** Reference to "rtcIsNull"
'006c943d    ff1540114000            call dword ptr [00401140]
'006c9443    8b0e                    mov ecx, dword ptr [esi]
'006c9445    56                      push esi
'006c9446    8985a0feffff            mov dword ptr [ebp+fffffea0], eax
'006c944c    ff9130030000            call dword ptr [ecx+00000330]
'006c9452    50                      push eax
'006c9453    8d55bc                  lea edx, dword ptr [ebp-44]
'006c9456    52                      push edx
'006c9457    ffd3                    call ebx
Set var_58 = IsNull(var_46)
'006c9459    8d55c4                  lea edx, dword ptr [ebp-3c]
'006c945c    8bf8                    mov edi, eax
'006c945e    8b45e8                  mov eax, dword ptr [ebp-18]
'006c9461    8b08                    mov ecx, dword ptr [eax]
'006c9463    52                      push edx
'006c9464    50                      push eax
'006c9465    ff91b4000000            call dword ptr [ecx+000000b4]
'006c946b    dbe2                    fnclex
'006c946d    85c0                    test ax, ax
'006c946f    7d15                    jge 6c9486
'006c9471    8b4de8                  mov ecx, dword ptr [ebp-18]
'006c9474    68b4000000              push 000000b4
'006c9479    6830314300              push 00433130
'006c947e    51                      push ecx
'006c947f    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c9480    ff1580104000            call dword ptr [00401080]
'006c9486    8b45c4                  mov eax, dword ptr [ebp-3c]
'006c9489    8b10                    mov edx, dword ptr [eax]
'006c948b    8d5dc0                  lea ebx, dword ptr [ebp-40]
'006c948e    53                      push ebx
'006c948f    83ec10                  sub esp, 10
'006c9492    8bdc                    mov ebx, esp
'006c9494    b908000000              mov ecx, 00000008
'006c9499    890b                    mov dword ptr [ebx], ecx
'006c949b    8b8d58ffffff            mov ecx, dword ptr [ebp+ffffff58]
'006c94a1    894b04                  mov dword ptr [ebx+04], ecx
'006c94a4    c7855cffffff20f34300    mov dword ptr [ebp+ffffff5c], 0043f320
'006c94ae    8b8d5cffffff            mov ecx, dword ptr [ebp+ffffff5c]
'006c94b4    894b08                  mov dword ptr [ebx+08], ecx
'006c94b7    8b8d60ffffff            mov ecx, dword ptr [ebp+ffffff60]
'006c94bd    50                      push eax
'006c94be    898588feffff            mov dword ptr [ebp+fffffe88], eax
'006c94c4    894b0c                  mov dword ptr [ebx+0c], ecx
'006c94c7    ff5230                  call dword ptr [edx+30]
'006c94ca    dbe2                    fnclex
'006c94cc    85c0                    test ax, ax
'006c94ce    7d15                    jge 6c94e5
'006c94d0    8b9588feffff            mov edx, dword ptr [ebp+fffffe88]
'006c94d6    6a30                    push 30
'006c94d8    68d8304300              push 004330d8
'006c94dd    52                      push edx
'006c94de    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c94df    ff1580104000            call dword ptr [00401080]
'006c94e5    8b45c0                  mov eax, dword ptr [ebp-40]
'006c94e8    8d9534ffffff            lea edx, dword ptr [ebp+ffffff34]
'006c94ee    8d4d94                  lea ecx, dword ptr [ebp-6c]
'006c94f1    c745c000000000          mov dword ptr [ebp-40], 00000000
'006c94f8    89458c                  mov dword ptr [ebp-74], eax
'006c94fb    c7458409000000          mov dword ptr [ebp-7c], 00000009
'006c9502    c7853cffffffcc134300    mov dword ptr [ebp+ffffff3c], 004313cc
'006c950c    c78534ffffff08000000    mov dword ptr [ebp+ffffff34], 00000008

' *** Reference to "__vbaVarDup"
'006c9516    ff1598124000            call dword ptr [00401298]
var_121 = (vbNullChar)
'006c951c    668b85a0feffff          mov ax, word ptr [ebp+fffffea0]
'006c9523    8d4d84                  lea ecx, dword ptr [ebp-7c]
'006c9526    51                      push ecx
'006c9527    8d5594                  lea edx, dword ptr [ebp-6c]
'006c952a    6689854cffffff          mov word ptr [ebp+ffffff4c], ax
'006c9531    52                      push edx
'006c9532    8d8544ffffff            lea eax, dword ptr [ebp+ffffff44]
'006c9538    50                      push eax
'006c9539    8d8d74ffffff            lea ecx, dword ptr [ebp+ffffff74]
'006c953f    51                      push ecx
'006c9540    c78544ffffff0b000000    mov dword ptr [ebp+ffffff44], 0000000b

' *** Reference to "rtcImmediateIf"
'006c954a    ff154c124000            call dword ptr [0040124c]
'006c9550    8b1f                    mov ebx, dword ptr [edi]
'006c9552    8d9574ffffff            lea edx, dword ptr [ebp+ffffff74]
'006c9558    52                      push edx
'006c9559    8d45e4                  lea eax, dword ptr [ebp-1c]
'006c955c    50                      push eax

' *** Reference to "__vbaStrVarVal"
'006c955d    ff15fc114000            call dword ptr [004011fc]
'006c9563    50                      push eax
'006c9564    57                      push edi
'006c9565    ff93a4000000            call dword ptr [ebx+000000a4]
'006c956b    dbe2                    fnclex
'006c956d    85c0                    test ax, ax
'006c956f    7d12                    jge 6c9583
'006c9571    68a4000000              push 000000a4
'006c9576    68d00d4300              push 00430dd0
'006c957b    57                      push edi
'006c957c    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c957d    ff1580104000            call dword ptr [00401080]
'006c9583    8d4de4                  lea ecx, dword ptr [ebp-1c]

' *** Reference to "__vbaFreeStr"
'006c9586    ff150c134000            call dword ptr [0040130c]
'#Cleanup(var_40)
'006c958c    8d4dbc                  lea ecx, dword ptr [ebp-44]
'006c958f    51                      push ecx
'006c9590    8d55c4                  lea edx, dword ptr [ebp-3c]
'006c9593    52                      push edx
'006c9594    8d45cc                  lea eax, dword ptr [ebp-34]
'006c9597    50                      push eax
'006c9598    6a03                    push 03

' *** Reference to "__vbaFreeObjList"
'006c959a    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_43, var_9, var_58)
'006c95a0    8d8d74ffffff            lea ecx, dword ptr [ebp+ffffff74]
'006c95a6    51                      push ecx
'006c95a7    8d5584                  lea edx, dword ptr [ebp-7c]
'006c95aa    52                      push edx
'006c95ab    8d4594                  lea eax, dword ptr [ebp-6c]
'006c95ae    50                      push eax
'006c95af    8d8d44ffffff            lea ecx, dword ptr [ebp+ffffff44]
'006c95b5    51                      push ecx
'006c95b6    8d55a4                  lea edx, dword ptr [ebp-5c]
'006c95b9    52                      push edx
'006c95ba    6a05                    push 05

' *** Reference to "__vbaFreeVarList"
'006c95bc    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_83, var_125, var_121, var_119, var_91)
'006c95c2    8b45e8                  mov eax, dword ptr [ebp-18]
'006c95c5    8b08                    mov ecx, dword ptr [eax]
'006c95c7    83c428                  add esp, 28
'006c95ca    8d55cc                  lea edx, dword ptr [ebp-34]
'006c95cd    52                      push edx
'006c95ce    50                      push eax
'006c95cf    ff91b4000000            call dword ptr [ecx+000000b4]
'006c95d5    dbe2                    fnclex
'006c95d7    85c0                    test ax, ax
'006c95d9    7d15                    jge 6c95f0
'006c95db    8b4de8                  mov ecx, dword ptr [ebp-18]
'006c95de    68b4000000              push 000000b4
'006c95e3    6830314300              push 00433130
'006c95e8    51                      push ecx
'006c95e9    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c95ea    ff1580104000            call dword ptr [00401080]
'006c95f0    8b45cc                  mov eax, dword ptr [ebp-34]
'006c95f3    8d5dc8                  lea ebx, dword ptr [ebp-38]
'006c95f6    53                      push ebx
'006c95f7    83ec10                  sub esp, 10
'006c95fa    ba08000000              mov edx, 00000008
'006c95ff    8bdc                    mov ebx, esp
'006c9601    8913                    mov dword ptr [ebx], edx
'006c9603    899564ffffff            mov dword ptr [ebp+ffffff64], edx
'006c9609    8b9568ffffff            mov edx, dword ptr [ebp+ffffff68]
'006c960f    b9001b4400              mov ecx, 00441b00
'006c9614    895304                  mov dword ptr [ebx+04], edx
'006c9617    898d6cffffff            mov dword ptr [ebp+ffffff6c], ecx
'006c961d    8b38                    mov edi, dword ptr [eax]
'006c961f    894b08                  mov dword ptr [ebx+08], ecx
'006c9622    8b8d70ffffff            mov ecx, dword ptr [ebp+ffffff70]
'006c9628    50                      push eax
'006c9629    898594feffff            mov dword ptr [ebp+fffffe94], eax
'006c962f    894b0c                  mov dword ptr [ebx+0c], ecx
'006c9632    ff5730                  call dword ptr [edi+30]
'006c9635    dbe2                    fnclex
'006c9637    85c0                    test ax, ax
'006c9639    7d15                    jge 6c9650
'006c963b    8b9594feffff            mov edx, dword ptr [ebp+fffffe94]
'006c9641    6a30                    push 30
'006c9643    68d8304300              push 004330d8
'006c9648    52                      push edx
'006c9649    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c964a    ff1580104000            call dword ptr [00401080]
'006c9650    8b45c8                  mov eax, dword ptr [ebp-38]
'006c9653    8945ac                  mov dword ptr [ebp-54], eax
'006c9656    8d45a4                  lea eax, dword ptr [ebp-5c]
'006c9659    50                      push eax
'006c965a    c745c800000000          mov dword ptr [ebp-38], 00000000
'006c9661    c745a409000000          mov dword ptr [ebp-5c], 00000009

' *** Reference to "rtcIsNull"
'006c9668    ff1540114000            call dword ptr [00401140]
'006c966e    8b0e                    mov ecx, dword ptr [esi]
'006c9670    56                      push esi
'006c9671    8985a0feffff            mov dword ptr [ebp+fffffea0], eax
'006c9677    ff9134030000            call dword ptr [ecx+00000334]
'006c967d    50                      push eax
'006c967e    8d55bc                  lea edx, dword ptr [ebp-44]
'006c9681    52                      push edx

' *** Reference to "__vbaObjSet"
'006c9682    ff15b4104000            call dword ptr [004010b4]
Set var_58 = IsNull(var_46)
'006c9688    8d55c4                  lea edx, dword ptr [ebp-3c]
'006c968b    8bf8                    mov edi, eax
'006c968d    8b45e8                  mov eax, dword ptr [ebp-18]
'006c9690    8b08                    mov ecx, dword ptr [eax]
'006c9692    52                      push edx
'006c9693    50                      push eax
'006c9694    ff91b4000000            call dword ptr [ecx+000000b4]
'006c969a    dbe2                    fnclex
'006c969c    85c0                    test ax, ax
'006c969e    7d15                    jge 6c96b5
'006c96a0    8b4de8                  mov ecx, dword ptr [ebp-18]
'006c96a3    68b4000000              push 000000b4
'006c96a8    6830314300              push 00433130
'006c96ad    51                      push ecx
'006c96ae    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c96af    ff1580104000            call dword ptr [00401080]
'006c96b5    8b45c4                  mov eax, dword ptr [ebp-3c]
'006c96b8    8b10                    mov edx, dword ptr [eax]
'006c96ba    8d5dc0                  lea ebx, dword ptr [ebp-40]
'006c96bd    53                      push ebx
'006c96be    83ec10                  sub esp, 10
'006c96c1    8bdc                    mov ebx, esp
'006c96c3    b908000000              mov ecx, 00000008
'006c96c8    890b                    mov dword ptr [ebx], ecx
'006c96ca    8b8d58ffffff            mov ecx, dword ptr [ebp+ffffff58]
'006c96d0    894b04                  mov dword ptr [ebx+04], ecx
'006c96d3    c7855cffffff001b4400    mov dword ptr [ebp+ffffff5c], 00441b00
'006c96dd    8b8d5cffffff            mov ecx, dword ptr [ebp+ffffff5c]
'006c96e3    894b08                  mov dword ptr [ebx+08], ecx
'006c96e6    8b8d60ffffff            mov ecx, dword ptr [ebp+ffffff60]
'006c96ec    50                      push eax
'006c96ed    898588feffff            mov dword ptr [ebp+fffffe88], eax
'006c96f3    894b0c                  mov dword ptr [ebx+0c], ecx
'006c96f6    ff5230                  call dword ptr [edx+30]
'006c96f9    dbe2                    fnclex
'006c96fb    85c0                    test ax, ax
'006c96fd    7d15                    jge 6c9714
'006c96ff    8b9588feffff            mov edx, dword ptr [ebp+fffffe88]
'006c9705    6a30                    push 30
'006c9707    68d8304300              push 004330d8
'006c970c    52                      push edx
'006c970d    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c970e    ff1580104000            call dword ptr [00401080]
'006c9714    8b45c0                  mov eax, dword ptr [ebp-40]
'006c9717    8d9534ffffff            lea edx, dword ptr [ebp+ffffff34]
'006c971d    8d4d94                  lea ecx, dword ptr [ebp-6c]
'006c9720    c745c000000000          mov dword ptr [ebp-40], 00000000
'006c9727    89458c                  mov dword ptr [ebp-74], eax
'006c972a    c7458409000000          mov dword ptr [ebp-7c], 00000009
'006c9731    c7853cffffffcc134300    mov dword ptr [ebp+ffffff3c], 004313cc
'006c973b    c78534ffffff08000000    mov dword ptr [ebp+ffffff34], 00000008

' *** Reference to "__vbaVarDup"
'006c9745    ff1598124000            call dword ptr [00401298]
var_121 = (vbNullChar)
'006c974b    668b85a0feffff          mov ax, word ptr [ebp+fffffea0]
'006c9752    8d4d84                  lea ecx, dword ptr [ebp-7c]
'006c9755    51                      push ecx
'006c9756    8d5594                  lea edx, dword ptr [ebp-6c]
'006c9759    6689854cffffff          mov word ptr [ebp+ffffff4c], ax
'006c9760    52                      push edx
'006c9761    8d8544ffffff            lea eax, dword ptr [ebp+ffffff44]
'006c9767    50                      push eax
'006c9768    8d8d74ffffff            lea ecx, dword ptr [ebp+ffffff74]
'006c976e    51                      push ecx
'006c976f    c78544ffffff0b000000    mov dword ptr [ebp+ffffff44], 0000000b

' *** Reference to "rtcImmediateIf"
'006c9779    ff154c124000            call dword ptr [0040124c]
'006c977f    8b1f                    mov ebx, dword ptr [edi]
'006c9781    8d9574ffffff            lea edx, dword ptr [ebp+ffffff74]
'006c9787    52                      push edx
'006c9788    8d45e4                  lea eax, dword ptr [ebp-1c]
'006c978b    50                      push eax

' *** Reference to "__vbaStrVarVal"
'006c978c    ff15fc114000            call dword ptr [004011fc]
'006c9792    50                      push eax
'006c9793    57                      push edi
'006c9794    ff93a4000000            call dword ptr [ebx+000000a4]
'006c979a    dbe2                    fnclex
'006c979c    85c0                    test ax, ax
'006c979e    7d12                    jge 6c97b2
'006c97a0    68a4000000              push 000000a4
'006c97a5    68d00d4300              push 00430dd0
'006c97aa    57                      push edi
'006c97ab    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c97ac    ff1580104000            call dword ptr [00401080]
'006c97b2    8d4de4                  lea ecx, dword ptr [ebp-1c]

' *** Reference to "__vbaFreeStr"
'006c97b5    ff150c134000            call dword ptr [0040130c]
'#Cleanup(var_40)
'006c97bb    8d4dbc                  lea ecx, dword ptr [ebp-44]
'006c97be    51                      push ecx
'006c97bf    8d55c4                  lea edx, dword ptr [ebp-3c]
'006c97c2    52                      push edx
'006c97c3    8d45cc                  lea eax, dword ptr [ebp-34]
'006c97c6    50                      push eax
'006c97c7    6a03                    push 03

' *** Reference to "__vbaFreeObjList"
'006c97c9    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_43, var_9, var_58)
'006c97cf    8d8d74ffffff            lea ecx, dword ptr [ebp+ffffff74]
'006c97d5    51                      push ecx
'006c97d6    8d5584                  lea edx, dword ptr [ebp-7c]
'006c97d9    52                      push edx
'006c97da    8d4594                  lea eax, dword ptr [ebp-6c]
'006c97dd    50                      push eax
'006c97de    8d8d44ffffff            lea ecx, dword ptr [ebp+ffffff44]
'006c97e4    51                      push ecx
'006c97e5    8d55a4                  lea edx, dword ptr [ebp-5c]
'006c97e8    52                      push edx
'006c97e9    6a05                    push 05

' *** Reference to "__vbaFreeVarList"
'006c97eb    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_83, var_125, var_121, var_119, var_91)
'006c97f1    8b45e8                  mov eax, dword ptr [ebp-18]
'006c97f4    8b08                    mov ecx, dword ptr [eax]
'006c97f6    83c428                  add esp, 28
'006c97f9    8d55cc                  lea edx, dword ptr [ebp-34]
'006c97fc    52                      push edx
'006c97fd    50                      push eax
'006c97fe    ff91b4000000            call dword ptr [ecx+000000b4]
'006c9804    dbe2                    fnclex
'006c9806    85c0                    test ax, ax
'006c9808    7d15                    jge 6c981f
'006c980a    8b4de8                  mov ecx, dword ptr [ebp-18]
'006c980d    68b4000000              push 000000b4
'006c9812    6830314300              push 00433130
'006c9817    51                      push ecx
'006c9818    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c9819    ff1580104000            call dword ptr [00401080]
'006c981f    8b45cc                  mov eax, dword ptr [ebp-34]
'006c9822    8d5dc8                  lea ebx, dword ptr [ebp-38]
'006c9825    53                      push ebx
'006c9826    83ec10                  sub esp, 10
'006c9829    ba08000000              mov edx, 00000008
'006c982e    8bdc                    mov ebx, esp
'006c9830    8913                    mov dword ptr [ebx], edx
'006c9832    899564ffffff            mov dword ptr [ebp+ffffff64], edx
'006c9838    8b9568ffffff            mov edx, dword ptr [ebp+ffffff68]
'006c983e    b9401b4400              mov ecx, 00441b40
'006c9843    895304                  mov dword ptr [ebx+04], edx
'006c9846    898d6cffffff            mov dword ptr [ebp+ffffff6c], ecx
'006c984c    8b38                    mov edi, dword ptr [eax]
'006c984e    894b08                  mov dword ptr [ebx+08], ecx
'006c9851    8b8d70ffffff            mov ecx, dword ptr [ebp+ffffff70]
'006c9857    50                      push eax
'006c9858    898594feffff            mov dword ptr [ebp+fffffe94], eax
'006c985e    894b0c                  mov dword ptr [ebx+0c], ecx
'006c9861    ff5730                  call dword ptr [edi+30]
'006c9864    dbe2                    fnclex
'006c9866    85c0                    test ax, ax
'006c9868    7d15                    jge 6c987f
'006c986a    8b9594feffff            mov edx, dword ptr [ebp+fffffe94]
'006c9870    6a30                    push 30
'006c9872    68d8304300              push 004330d8
'006c9877    52                      push edx
'006c9878    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c9879    ff1580104000            call dword ptr [00401080]
'006c987f    8b45c8                  mov eax, dword ptr [ebp-38]
'006c9882    8945ac                  mov dword ptr [ebp-54], eax
'006c9885    8d45a4                  lea eax, dword ptr [ebp-5c]
'006c9888    50                      push eax
'006c9889    c745c800000000          mov dword ptr [ebp-38], 00000000
'006c9890    c745a409000000          mov dword ptr [ebp-5c], 00000009

' *** Reference to "rtcIsNull"
'006c9897    ff1540114000            call dword ptr [00401140]
'006c989d    8b0e                    mov ecx, dword ptr [esi]
'006c989f    56                      push esi
'006c98a0    8985a0feffff            mov dword ptr [ebp+fffffea0], eax
'006c98a6    ff9138030000            call dword ptr [ecx+00000338]
'006c98ac    50                      push eax
'006c98ad    8d55bc                  lea edx, dword ptr [ebp-44]
'006c98b0    52                      push edx

' *** Reference to "__vbaObjSet"
'006c98b1    ff15b4104000            call dword ptr [004010b4]
Set var_58 = IsNull(var_46)
'006c98b7    8d55c4                  lea edx, dword ptr [ebp-3c]
'006c98ba    8bf8                    mov edi, eax
'006c98bc    8b45e8                  mov eax, dword ptr [ebp-18]
'006c98bf    8b08                    mov ecx, dword ptr [eax]
'006c98c1    52                      push edx
'006c98c2    50                      push eax
'006c98c3    ff91b4000000            call dword ptr [ecx+000000b4]
'006c98c9    dbe2                    fnclex
'006c98cb    85c0                    test ax, ax
'006c98cd    7d15                    jge 6c98e4
'006c98cf    8b4de8                  mov ecx, dword ptr [ebp-18]
'006c98d2    68b4000000              push 000000b4
'006c98d7    6830314300              push 00433130
'006c98dc    51                      push ecx
'006c98dd    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c98de    ff1580104000            call dword ptr [00401080]
'006c98e4    8b45c4                  mov eax, dword ptr [ebp-3c]
'006c98e7    8b10                    mov edx, dword ptr [eax]
'006c98e9    8d5dc0                  lea ebx, dword ptr [ebp-40]
'006c98ec    53                      push ebx
'006c98ed    83ec10                  sub esp, 10
'006c98f0    8bdc                    mov ebx, esp
'006c98f2    b908000000              mov ecx, 00000008
'006c98f7    890b                    mov dword ptr [ebx], ecx
'006c98f9    8b8d58ffffff            mov ecx, dword ptr [ebp+ffffff58]
'006c98ff    894b04                  mov dword ptr [ebx+04], ecx
'006c9902    c7855cffffff401b4400    mov dword ptr [ebp+ffffff5c], 00441b40
'006c990c    8b8d5cffffff            mov ecx, dword ptr [ebp+ffffff5c]
'006c9912    894b08                  mov dword ptr [ebx+08], ecx
'006c9915    8b8d60ffffff            mov ecx, dword ptr [ebp+ffffff60]
'006c991b    50                      push eax
'006c991c    898588feffff            mov dword ptr [ebp+fffffe88], eax
'006c9922    894b0c                  mov dword ptr [ebx+0c], ecx
'006c9925    ff5230                  call dword ptr [edx+30]
'006c9928    dbe2                    fnclex
'006c992a    85c0                    test ax, ax
'006c992c    7d15                    jge 6c9943
'006c992e    8b9588feffff            mov edx, dword ptr [ebp+fffffe88]
'006c9934    6a30                    push 30
'006c9936    68d8304300              push 004330d8
'006c993b    52                      push edx
'006c993c    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c993d    ff1580104000            call dword ptr [00401080]
'006c9943    8b45c0                  mov eax, dword ptr [ebp-40]
'006c9946    8d9534ffffff            lea edx, dword ptr [ebp+ffffff34]
'006c994c    8d4d94                  lea ecx, dword ptr [ebp-6c]
'006c994f    c745c000000000          mov dword ptr [ebp-40], 00000000
'006c9956    89458c                  mov dword ptr [ebp-74], eax
'006c9959    c7458409000000          mov dword ptr [ebp-7c], 00000009
'006c9960    c7853cffffffcc134300    mov dword ptr [ebp+ffffff3c], 004313cc
'006c996a    c78534ffffff08000000    mov dword ptr [ebp+ffffff34], 00000008

' *** Reference to "__vbaVarDup"
'006c9974    ff1598124000            call dword ptr [00401298]
var_121 = (vbNullChar)
'006c997a    668b85a0feffff          mov ax, word ptr [ebp+fffffea0]
'006c9981    8d4d84                  lea ecx, dword ptr [ebp-7c]
'006c9984    51                      push ecx
'006c9985    8d5594                  lea edx, dword ptr [ebp-6c]
'006c9988    6689854cffffff          mov word ptr [ebp+ffffff4c], ax
'006c998f    52                      push edx
'006c9990    8d8544ffffff            lea eax, dword ptr [ebp+ffffff44]
'006c9996    50                      push eax
'006c9997    8d8d74ffffff            lea ecx, dword ptr [ebp+ffffff74]
'006c999d    51                      push ecx
'006c999e    c78544ffffff0b000000    mov dword ptr [ebp+ffffff44], 0000000b

' *** Reference to "rtcImmediateIf"
'006c99a8    ff154c124000            call dword ptr [0040124c]
'006c99ae    8b1f                    mov ebx, dword ptr [edi]
'006c99b0    8d9574ffffff            lea edx, dword ptr [ebp+ffffff74]
'006c99b6    52                      push edx
'006c99b7    8d45e4                  lea eax, dword ptr [ebp-1c]
'006c99ba    50                      push eax

' *** Reference to "__vbaStrVarVal"
'006c99bb    ff15fc114000            call dword ptr [004011fc]
'006c99c1    50                      push eax
'006c99c2    57                      push edi
'006c99c3    ff93a4000000            call dword ptr [ebx+000000a4]
'006c99c9    dbe2                    fnclex
'006c99cb    85c0                    test ax, ax
'006c99cd    7d12                    jge 6c99e1
'006c99cf    68a4000000              push 000000a4
'006c99d4    68d00d4300              push 00430dd0
'006c99d9    57                      push edi
'006c99da    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c99db    ff1580104000            call dword ptr [00401080]
'006c99e1    8d4de4                  lea ecx, dword ptr [ebp-1c]

' *** Reference to "__vbaFreeStr"
'006c99e4    ff150c134000            call dword ptr [0040130c]
'#Cleanup(var_40)
'006c99ea    8d4dbc                  lea ecx, dword ptr [ebp-44]
'006c99ed    51                      push ecx
'006c99ee    8d55c4                  lea edx, dword ptr [ebp-3c]
'006c99f1    52                      push edx
'006c99f2    8d45cc                  lea eax, dword ptr [ebp-34]
'006c99f5    50                      push eax
'006c99f6    6a03                    push 03

' *** Reference to "__vbaFreeObjList"
'006c99f8    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_43, var_9, var_58)
'006c99fe    8d8d74ffffff            lea ecx, dword ptr [ebp+ffffff74]
'006c9a04    51                      push ecx
'006c9a05    8d5584                  lea edx, dword ptr [ebp-7c]
'006c9a08    52                      push edx
'006c9a09    8d4594                  lea eax, dword ptr [ebp-6c]
'006c9a0c    50                      push eax
'006c9a0d    8d8d44ffffff            lea ecx, dword ptr [ebp+ffffff44]
'006c9a13    51                      push ecx
'006c9a14    8d55a4                  lea edx, dword ptr [ebp-5c]
'006c9a17    52                      push edx
'006c9a18    6a05                    push 05

' *** Reference to "__vbaFreeVarList"
'006c9a1a    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_83, var_125, var_121, var_119, var_91)
'006c9a20    8b45e8                  mov eax, dword ptr [ebp-18]
'006c9a23    8b08                    mov ecx, dword ptr [eax]
'006c9a25    83c428                  add esp, 28
'006c9a28    8d55cc                  lea edx, dword ptr [ebp-34]
'006c9a2b    52                      push edx
'006c9a2c    50                      push eax
'006c9a2d    ff91b4000000            call dword ptr [ecx+000000b4]
'006c9a33    dbe2                    fnclex
'006c9a35    85c0                    test ax, ax
'006c9a37    7d15                    jge 6c9a4e
'006c9a39    8b4de8                  mov ecx, dword ptr [ebp-18]
'006c9a3c    68b4000000              push 000000b4
'006c9a41    6830314300              push 00433130
'006c9a46    51                      push ecx
'006c9a47    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c9a48    ff1580104000            call dword ptr [00401080]
'006c9a4e    8b45cc                  mov eax, dword ptr [ebp-34]
'006c9a51    8d5dc8                  lea ebx, dword ptr [ebp-38]
'006c9a54    53                      push ebx
'006c9a55    83ec10                  sub esp, 10
'006c9a58    ba08000000              mov edx, 00000008
'006c9a5d    8bdc                    mov ebx, esp
'006c9a5f    8913                    mov dword ptr [ebx], edx
'006c9a61    899564ffffff            mov dword ptr [ebp+ffffff64], edx
'006c9a67    8b9568ffffff            mov edx, dword ptr [ebp+ffffff68]
'006c9a6d    b954f54300              mov ecx, 0043f554
'006c9a72    895304                  mov dword ptr [ebx+04], edx
'006c9a75    898d6cffffff            mov dword ptr [ebp+ffffff6c], ecx
'006c9a7b    8b38                    mov edi, dword ptr [eax]
'006c9a7d    894b08                  mov dword ptr [ebx+08], ecx
'006c9a80    8b8d70ffffff            mov ecx, dword ptr [ebp+ffffff70]
'006c9a86    50                      push eax
'006c9a87    898594feffff            mov dword ptr [ebp+fffffe94], eax
'006c9a8d    894b0c                  mov dword ptr [ebx+0c], ecx
'006c9a90    ff5730                  call dword ptr [edi+30]
'006c9a93    dbe2                    fnclex
'006c9a95    85c0                    test ax, ax
'006c9a97    7d15                    jge 6c9aae
'006c9a99    8b9594feffff            mov edx, dword ptr [ebp+fffffe94]
'006c9a9f    6a30                    push 30
'006c9aa1    68d8304300              push 004330d8
'006c9aa6    52                      push edx
'006c9aa7    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c9aa8    ff1580104000            call dword ptr [00401080]
'006c9aae    8b45c8                  mov eax, dword ptr [ebp-38]
'006c9ab1    8945ac                  mov dword ptr [ebp-54], eax
'006c9ab4    8d45a4                  lea eax, dword ptr [ebp-5c]
'006c9ab7    50                      push eax
'006c9ab8    c745c800000000          mov dword ptr [ebp-38], 00000000
'006c9abf    c745a409000000          mov dword ptr [ebp-5c], 00000009

' *** Reference to "rtcIsNull"
'006c9ac6    ff1540114000            call dword ptr [00401140]
'006c9acc    8b0e                    mov ecx, dword ptr [esi]
'006c9ace    56                      push esi
'006c9acf    8985a0feffff            mov dword ptr [ebp+fffffea0], eax
'006c9ad5    ff913c030000            call dword ptr [ecx+0000033c]
'006c9adb    50                      push eax
'006c9adc    8d55bc                  lea edx, dword ptr [ebp-44]
'006c9adf    52                      push edx

' *** Reference to "__vbaObjSet"
'006c9ae0    ff15b4104000            call dword ptr [004010b4]
Set var_58 = IsNull(var_46)
'006c9ae6    8d55c4                  lea edx, dword ptr [ebp-3c]
'006c9ae9    8bf8                    mov edi, eax
'006c9aeb    8b45e8                  mov eax, dword ptr [ebp-18]
'006c9aee    8b08                    mov ecx, dword ptr [eax]
'006c9af0    52                      push edx
'006c9af1    50                      push eax
'006c9af2    ff91b4000000            call dword ptr [ecx+000000b4]
'006c9af8    dbe2                    fnclex
'006c9afa    85c0                    test ax, ax
'006c9afc    7d15                    jge 6c9b13
'006c9afe    8b4de8                  mov ecx, dword ptr [ebp-18]
'006c9b01    68b4000000              push 000000b4
'006c9b06    6830314300              push 00433130
'006c9b0b    51                      push ecx
'006c9b0c    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c9b0d    ff1580104000            call dword ptr [00401080]
'006c9b13    8b45c4                  mov eax, dword ptr [ebp-3c]
'006c9b16    8b10                    mov edx, dword ptr [eax]
'006c9b18    8d5dc0                  lea ebx, dword ptr [ebp-40]
'006c9b1b    53                      push ebx
'006c9b1c    83ec10                  sub esp, 10
'006c9b1f    8bdc                    mov ebx, esp
'006c9b21    b908000000              mov ecx, 00000008
'006c9b26    890b                    mov dword ptr [ebx], ecx
'006c9b28    8b8d58ffffff            mov ecx, dword ptr [ebp+ffffff58]
'006c9b2e    894b04                  mov dword ptr [ebx+04], ecx
'006c9b31    c7855cffffff54f54300    mov dword ptr [ebp+ffffff5c], 0043f554
'006c9b3b    8b8d5cffffff            mov ecx, dword ptr [ebp+ffffff5c]
'006c9b41    894b08                  mov dword ptr [ebx+08], ecx
'006c9b44    8b8d60ffffff            mov ecx, dword ptr [ebp+ffffff60]
'006c9b4a    50                      push eax
'006c9b4b    898588feffff            mov dword ptr [ebp+fffffe88], eax
'006c9b51    894b0c                  mov dword ptr [ebx+0c], ecx
'006c9b54    ff5230                  call dword ptr [edx+30]
'006c9b57    dbe2                    fnclex
'006c9b59    85c0                    test ax, ax
'006c9b5b    7d15                    jge 6c9b72
'006c9b5d    8b9588feffff            mov edx, dword ptr [ebp+fffffe88]
'006c9b63    6a30                    push 30
'006c9b65    68d8304300              push 004330d8
'006c9b6a    52                      push edx
'006c9b6b    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c9b6c    ff1580104000            call dword ptr [00401080]
'006c9b72    8b45c0                  mov eax, dword ptr [ebp-40]
'006c9b75    8d9534ffffff            lea edx, dword ptr [ebp+ffffff34]
'006c9b7b    8d4d94                  lea ecx, dword ptr [ebp-6c]
'006c9b7e    c745c000000000          mov dword ptr [ebp-40], 00000000
'006c9b85    89458c                  mov dword ptr [ebp-74], eax
'006c9b88    c7458409000000          mov dword ptr [ebp-7c], 00000009
'006c9b8f    c7853cffffffcc134300    mov dword ptr [ebp+ffffff3c], 004313cc
'006c9b99    c78534ffffff08000000    mov dword ptr [ebp+ffffff34], 00000008

' *** Reference to "__vbaVarDup"
'006c9ba3    ff1598124000            call dword ptr [00401298]
var_121 = (vbNullChar)
'006c9ba9    668b85a0feffff          mov ax, word ptr [ebp+fffffea0]
'006c9bb0    8d4d84                  lea ecx, dword ptr [ebp-7c]
'006c9bb3    51                      push ecx
'006c9bb4    8d5594                  lea edx, dword ptr [ebp-6c]
'006c9bb7    6689854cffffff          mov word ptr [ebp+ffffff4c], ax
'006c9bbe    52                      push edx
'006c9bbf    8d8544ffffff            lea eax, dword ptr [ebp+ffffff44]
'006c9bc5    50                      push eax
'006c9bc6    8d8d74ffffff            lea ecx, dword ptr [ebp+ffffff74]
'006c9bcc    51                      push ecx
'006c9bcd    c78544ffffff0b000000    mov dword ptr [ebp+ffffff44], 0000000b

' *** Reference to "rtcImmediateIf"
'006c9bd7    ff154c124000            call dword ptr [0040124c]
'006c9bdd    8b1f                    mov ebx, dword ptr [edi]
'006c9bdf    8d9574ffffff            lea edx, dword ptr [ebp+ffffff74]
'006c9be5    52                      push edx
'006c9be6    8d45e4                  lea eax, dword ptr [ebp-1c]
'006c9be9    50                      push eax

' *** Reference to "__vbaStrVarVal"
'006c9bea    ff15fc114000            call dword ptr [004011fc]
'006c9bf0    50                      push eax
'006c9bf1    57                      push edi
'006c9bf2    ff93a4000000            call dword ptr [ebx+000000a4]
'006c9bf8    dbe2                    fnclex
'006c9bfa    85c0                    test ax, ax
'006c9bfc    7d12                    jge 6c9c10
'006c9bfe    68a4000000              push 000000a4
'006c9c03    68d00d4300              push 00430dd0
'006c9c08    57                      push edi
'006c9c09    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c9c0a    ff1580104000            call dword ptr [00401080]
'006c9c10    8d4de4                  lea ecx, dword ptr [ebp-1c]

' *** Reference to "__vbaFreeStr"
'006c9c13    ff150c134000            call dword ptr [0040130c]
'#Cleanup(var_40)
'006c9c19    8d4dbc                  lea ecx, dword ptr [ebp-44]
'006c9c1c    51                      push ecx
'006c9c1d    8d55c4                  lea edx, dword ptr [ebp-3c]
'006c9c20    52                      push edx
'006c9c21    8d45cc                  lea eax, dword ptr [ebp-34]
'006c9c24    50                      push eax
'006c9c25    6a03                    push 03

' *** Reference to "__vbaFreeObjList"
'006c9c27    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_43, var_9, var_58)
'006c9c2d    8d8d74ffffff            lea ecx, dword ptr [ebp+ffffff74]
'006c9c33    51                      push ecx
'006c9c34    8d5584                  lea edx, dword ptr [ebp-7c]
'006c9c37    52                      push edx
'006c9c38    8d4594                  lea eax, dword ptr [ebp-6c]
'006c9c3b    50                      push eax
'006c9c3c    8d8d44ffffff            lea ecx, dword ptr [ebp+ffffff44]
'006c9c42    51                      push ecx
'006c9c43    8d55a4                  lea edx, dword ptr [ebp-5c]
'006c9c46    52                      push edx
'006c9c47    6a05                    push 05

' *** Reference to "__vbaFreeVarList"
'006c9c49    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_83, var_125, var_121, var_119, var_91)
'006c9c4f    8b45e8                  mov eax, dword ptr [ebp-18]
'006c9c52    8b08                    mov ecx, dword ptr [eax]
'006c9c54    83c428                  add esp, 28
'006c9c57    8d55cc                  lea edx, dword ptr [ebp-34]
'006c9c5a    52                      push edx
'006c9c5b    50                      push eax
'006c9c5c    ff91b4000000            call dword ptr [ecx+000000b4]
'006c9c62    dbe2                    fnclex
'006c9c64    85c0                    test ax, ax
'006c9c66    7d15                    jge 6c9c7d
'006c9c68    8b4de8                  mov ecx, dword ptr [ebp-18]
'006c9c6b    68b4000000              push 000000b4
'006c9c70    6830314300              push 00433130
'006c9c75    51                      push ecx
'006c9c76    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c9c77    ff1580104000            call dword ptr [00401080]
'006c9c7d    8b45cc                  mov eax, dword ptr [ebp-34]
'006c9c80    8d5dc8                  lea ebx, dword ptr [ebp-38]
'006c9c83    53                      push ebx
'006c9c84    83ec10                  sub esp, 10
'006c9c87    ba08000000              mov edx, 00000008
'006c9c8c    8bdc                    mov ebx, esp
'006c9c8e    8913                    mov dword ptr [ebx], edx
'006c9c90    899564ffffff            mov dword ptr [ebp+ffffff64], edx
'006c9c96    8b9568ffffff            mov edx, dword ptr [ebp+ffffff68]
'006c9c9c    b9b41b4400              mov ecx, 00441bb4
'006c9ca1    895304                  mov dword ptr [ebx+04], edx
'006c9ca4    898d6cffffff            mov dword ptr [ebp+ffffff6c], ecx
'006c9caa    8b38                    mov edi, dword ptr [eax]
'006c9cac    894b08                  mov dword ptr [ebx+08], ecx
'006c9caf    8b8d70ffffff            mov ecx, dword ptr [ebp+ffffff70]
'006c9cb5    50                      push eax
'006c9cb6    898594feffff            mov dword ptr [ebp+fffffe94], eax
'006c9cbc    894b0c                  mov dword ptr [ebx+0c], ecx
'006c9cbf    ff5730                  call dword ptr [edi+30]
'006c9cc2    dbe2                    fnclex
'006c9cc4    85c0                    test ax, ax
'006c9cc6    7d15                    jge 6c9cdd
'006c9cc8    8b9594feffff            mov edx, dword ptr [ebp+fffffe94]
'006c9cce    6a30                    push 30
'006c9cd0    68d8304300              push 004330d8
'006c9cd5    52                      push edx
'006c9cd6    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c9cd7    ff1580104000            call dword ptr [00401080]
'006c9cdd    8b45c8                  mov eax, dword ptr [ebp-38]
'006c9ce0    8945ac                  mov dword ptr [ebp-54], eax
'006c9ce3    8d45a4                  lea eax, dword ptr [ebp-5c]
'006c9ce6    50                      push eax
'006c9ce7    c745c800000000          mov dword ptr [ebp-38], 00000000
'006c9cee    c745a409000000          mov dword ptr [ebp-5c], 00000009

' *** Reference to "rtcIsNull"
'006c9cf5    ff1540114000            call dword ptr [00401140]
'006c9cfb    8b0e                    mov ecx, dword ptr [esi]
'006c9cfd    56                      push esi
'006c9cfe    8985a0feffff            mov dword ptr [ebp+fffffea0], eax
'006c9d04    ff9140030000            call dword ptr [ecx+00000340]
'006c9d0a    50                      push eax
'006c9d0b    8d55bc                  lea edx, dword ptr [ebp-44]
'006c9d0e    52                      push edx

' *** Reference to "__vbaObjSet"
'006c9d0f    ff15b4104000            call dword ptr [004010b4]
Set var_58 = IsNull(var_46)
'006c9d15    8d55c4                  lea edx, dword ptr [ebp-3c]
'006c9d18    8bf8                    mov edi, eax
'006c9d1a    8b45e8                  mov eax, dword ptr [ebp-18]
'006c9d1d    8b08                    mov ecx, dword ptr [eax]
'006c9d1f    52                      push edx
'006c9d20    50                      push eax
'006c9d21    ff91b4000000            call dword ptr [ecx+000000b4]
'006c9d27    dbe2                    fnclex
'006c9d29    85c0                    test ax, ax
'006c9d2b    7d15                    jge 6c9d42
'006c9d2d    8b4de8                  mov ecx, dword ptr [ebp-18]
'006c9d30    68b4000000              push 000000b4
'006c9d35    6830314300              push 00433130
'006c9d3a    51                      push ecx
'006c9d3b    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c9d3c    ff1580104000            call dword ptr [00401080]
'006c9d42    8b45c4                  mov eax, dword ptr [ebp-3c]
'006c9d45    8b10                    mov edx, dword ptr [eax]
'006c9d47    8d5dc0                  lea ebx, dword ptr [ebp-40]
'006c9d4a    53                      push ebx
'006c9d4b    83ec10                  sub esp, 10
'006c9d4e    8bdc                    mov ebx, esp
'006c9d50    b908000000              mov ecx, 00000008
'006c9d55    890b                    mov dword ptr [ebx], ecx
'006c9d57    8b8d58ffffff            mov ecx, dword ptr [ebp+ffffff58]
'006c9d5d    894b04                  mov dword ptr [ebx+04], ecx
'006c9d60    c7855cffffffb41b4400    mov dword ptr [ebp+ffffff5c], 00441bb4
'006c9d6a    8b8d5cffffff            mov ecx, dword ptr [ebp+ffffff5c]
'006c9d70    894b08                  mov dword ptr [ebx+08], ecx
'006c9d73    8b8d60ffffff            mov ecx, dword ptr [ebp+ffffff60]
'006c9d79    50                      push eax
'006c9d7a    898588feffff            mov dword ptr [ebp+fffffe88], eax
'006c9d80    894b0c                  mov dword ptr [ebx+0c], ecx
'006c9d83    ff5230                  call dword ptr [edx+30]
'006c9d86    dbe2                    fnclex
'006c9d88    85c0                    test ax, ax
'006c9d8a    7d15                    jge 6c9da1
'006c9d8c    8b9588feffff            mov edx, dword ptr [ebp+fffffe88]
'006c9d92    6a30                    push 30
'006c9d94    68d8304300              push 004330d8
'006c9d99    52                      push edx
'006c9d9a    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c9d9b    ff1580104000            call dword ptr [00401080]
'006c9da1    8b45c0                  mov eax, dword ptr [ebp-40]
'006c9da4    8d9534ffffff            lea edx, dword ptr [ebp+ffffff34]
'006c9daa    8d4d94                  lea ecx, dword ptr [ebp-6c]
'006c9dad    c745c000000000          mov dword ptr [ebp-40], 00000000
'006c9db4    89458c                  mov dword ptr [ebp-74], eax
'006c9db7    c7458409000000          mov dword ptr [ebp-7c], 00000009
'006c9dbe    c7853cffffffcc134300    mov dword ptr [ebp+ffffff3c], 004313cc
'006c9dc8    c78534ffffff08000000    mov dword ptr [ebp+ffffff34], 00000008

' *** Reference to "__vbaVarDup"
'006c9dd2    ff1598124000            call dword ptr [00401298]
var_121 = (vbNullChar)
'006c9dd8    668b85a0feffff          mov ax, word ptr [ebp+fffffea0]
'006c9ddf    8d4d84                  lea ecx, dword ptr [ebp-7c]
'006c9de2    51                      push ecx
'006c9de3    8d5594                  lea edx, dword ptr [ebp-6c]
'006c9de6    6689854cffffff          mov word ptr [ebp+ffffff4c], ax
'006c9ded    52                      push edx
'006c9dee    8d8544ffffff            lea eax, dword ptr [ebp+ffffff44]
'006c9df4    50                      push eax
'006c9df5    8d8d74ffffff            lea ecx, dword ptr [ebp+ffffff74]
'006c9dfb    51                      push ecx
'006c9dfc    c78544ffffff0b000000    mov dword ptr [ebp+ffffff44], 0000000b

' *** Reference to "rtcImmediateIf"
'006c9e06    ff154c124000            call dword ptr [0040124c]
'006c9e0c    8b1f                    mov ebx, dword ptr [edi]
'006c9e0e    8d9574ffffff            lea edx, dword ptr [ebp+ffffff74]
'006c9e14    52                      push edx
'006c9e15    8d45e4                  lea eax, dword ptr [ebp-1c]
'006c9e18    50                      push eax

' *** Reference to "__vbaStrVarVal"
'006c9e19    ff15fc114000            call dword ptr [004011fc]
'006c9e1f    50                      push eax
'006c9e20    57                      push edi
'006c9e21    ff93a4000000            call dword ptr [ebx+000000a4]
'006c9e27    dbe2                    fnclex
'006c9e29    85c0                    test ax, ax
'006c9e2b    7d12                    jge 6c9e3f
'006c9e2d    68a4000000              push 000000a4
'006c9e32    68d00d4300              push 00430dd0
'006c9e37    57                      push edi
'006c9e38    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c9e39    ff1580104000            call dword ptr [00401080]
'006c9e3f    8d4de4                  lea ecx, dword ptr [ebp-1c]

' *** Reference to "__vbaFreeStr"
'006c9e42    ff150c134000            call dword ptr [0040130c]
'#Cleanup(var_40)
'006c9e48    8d4dbc                  lea ecx, dword ptr [ebp-44]
'006c9e4b    51                      push ecx
'006c9e4c    8d55c4                  lea edx, dword ptr [ebp-3c]
'006c9e4f    52                      push edx
'006c9e50    8d45cc                  lea eax, dword ptr [ebp-34]
'006c9e53    50                      push eax
'006c9e54    6a03                    push 03

' *** Reference to "__vbaFreeObjList"
'006c9e56    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_43, var_9, var_58)
'006c9e5c    8d8d74ffffff            lea ecx, dword ptr [ebp+ffffff74]
'006c9e62    51                      push ecx
'006c9e63    8d5584                  lea edx, dword ptr [ebp-7c]
'006c9e66    52                      push edx
'006c9e67    8d4594                  lea eax, dword ptr [ebp-6c]
'006c9e6a    50                      push eax
'006c9e6b    8d8d44ffffff            lea ecx, dword ptr [ebp+ffffff44]
'006c9e71    51                      push ecx
'006c9e72    8d55a4                  lea edx, dword ptr [ebp-5c]
'006c9e75    52                      push edx
'006c9e76    6a05                    push 05

' *** Reference to "__vbaFreeVarList"
'006c9e78    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_83, var_125, var_121, var_119, var_91)
'006c9e7e    8b45e8                  mov eax, dword ptr [ebp-18]
'006c9e81    8b08                    mov ecx, dword ptr [eax]
'006c9e83    83c428                  add esp, 28
'006c9e86    8d55cc                  lea edx, dword ptr [ebp-34]
'006c9e89    52                      push edx
'006c9e8a    50                      push eax
'006c9e8b    ff91b4000000            call dword ptr [ecx+000000b4]
'006c9e91    dbe2                    fnclex
'006c9e93    85c0                    test ax, ax
'006c9e95    7d15                    jge 6c9eac
'006c9e97    8b4de8                  mov ecx, dword ptr [ebp-18]
'006c9e9a    68b4000000              push 000000b4
'006c9e9f    6830314300              push 00433130
'006c9ea4    51                      push ecx
'006c9ea5    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c9ea6    ff1580104000            call dword ptr [00401080]
'006c9eac    8b45cc                  mov eax, dword ptr [ebp-34]
'006c9eaf    8d5dc8                  lea ebx, dword ptr [ebp-38]
'006c9eb2    53                      push ebx
'006c9eb3    83ec10                  sub esp, 10
'006c9eb6    ba08000000              mov edx, 00000008
'006c9ebb    8bdc                    mov ebx, esp
'006c9ebd    8913                    mov dword ptr [ebx], edx
'006c9ebf    899564ffffff            mov dword ptr [ebp+ffffff64], edx
'006c9ec5    8b9568ffffff            mov edx, dword ptr [ebp+ffffff68]
'006c9ecb    b9f81b4400              mov ecx, 00441bf8
'006c9ed0    895304                  mov dword ptr [ebx+04], edx
'006c9ed3    898d6cffffff            mov dword ptr [ebp+ffffff6c], ecx
'006c9ed9    8b38                    mov edi, dword ptr [eax]
'006c9edb    894b08                  mov dword ptr [ebx+08], ecx
'006c9ede    8b8d70ffffff            mov ecx, dword ptr [ebp+ffffff70]
'006c9ee4    50                      push eax
'006c9ee5    898594feffff            mov dword ptr [ebp+fffffe94], eax
'006c9eeb    894b0c                  mov dword ptr [ebx+0c], ecx
'006c9eee    ff5730                  call dword ptr [edi+30]
'006c9ef1    dbe2                    fnclex
'006c9ef3    85c0                    test ax, ax
'006c9ef5    7d15                    jge 6c9f0c
'006c9ef7    8b9594feffff            mov edx, dword ptr [ebp+fffffe94]
'006c9efd    6a30                    push 30
'006c9eff    68d8304300              push 004330d8
'006c9f04    52                      push edx
'006c9f05    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c9f06    ff1580104000            call dword ptr [00401080]
'006c9f0c    8b45c8                  mov eax, dword ptr [ebp-38]
'006c9f0f    8945ac                  mov dword ptr [ebp-54], eax
'006c9f12    8d45a4                  lea eax, dword ptr [ebp-5c]
'006c9f15    50                      push eax
'006c9f16    c745c800000000          mov dword ptr [ebp-38], 00000000
'006c9f1d    c745a409000000          mov dword ptr [ebp-5c], 00000009

' *** Reference to "rtcIsNull"
'006c9f24    ff1540114000            call dword ptr [00401140]
'006c9f2a    8b0e                    mov ecx, dword ptr [esi]
'006c9f2c    56                      push esi
'006c9f2d    8985a0feffff            mov dword ptr [ebp+fffffea0], eax
'006c9f33    ff9148030000            call dword ptr [ecx+00000348]
'006c9f39    50                      push eax
'006c9f3a    8d55bc                  lea edx, dword ptr [ebp-44]
'006c9f3d    52                      push edx

' *** Reference to "__vbaObjSet"
'006c9f3e    ff15b4104000            call dword ptr [004010b4]
Set var_58 = IsNull(var_46)
'006c9f44    8d55c4                  lea edx, dword ptr [ebp-3c]
'006c9f47    8bf8                    mov edi, eax
'006c9f49    8b45e8                  mov eax, dword ptr [ebp-18]
'006c9f4c    8b08                    mov ecx, dword ptr [eax]
'006c9f4e    52                      push edx
'006c9f4f    50                      push eax
'006c9f50    ff91b4000000            call dword ptr [ecx+000000b4]
'006c9f56    dbe2                    fnclex
'006c9f58    85c0                    test ax, ax
'006c9f5a    7d15                    jge 6c9f71
'006c9f5c    8b4de8                  mov ecx, dword ptr [ebp-18]
'006c9f5f    68b4000000              push 000000b4
'006c9f64    6830314300              push 00433130
'006c9f69    51                      push ecx
'006c9f6a    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c9f6b    ff1580104000            call dword ptr [00401080]
'006c9f71    8b45c4                  mov eax, dword ptr [ebp-3c]
'006c9f74    8b10                    mov edx, dword ptr [eax]
'006c9f76    8d5dc0                  lea ebx, dword ptr [ebp-40]
'006c9f79    53                      push ebx
'006c9f7a    83ec10                  sub esp, 10
'006c9f7d    8bdc                    mov ebx, esp
'006c9f7f    b908000000              mov ecx, 00000008
'006c9f84    890b                    mov dword ptr [ebx], ecx
'006c9f86    8b8d58ffffff            mov ecx, dword ptr [ebp+ffffff58]
'006c9f8c    894b04                  mov dword ptr [ebx+04], ecx
'006c9f8f    c7855cfffffff81b4400    mov dword ptr [ebp+ffffff5c], 00441bf8
'006c9f99    8b8d5cffffff            mov ecx, dword ptr [ebp+ffffff5c]
'006c9f9f    894b08                  mov dword ptr [ebx+08], ecx
'006c9fa2    8b8d60ffffff            mov ecx, dword ptr [ebp+ffffff60]
'006c9fa8    50                      push eax
'006c9fa9    898588feffff            mov dword ptr [ebp+fffffe88], eax
'006c9faf    894b0c                  mov dword ptr [ebx+0c], ecx
'006c9fb2    ff5230                  call dword ptr [edx+30]
'006c9fb5    dbe2                    fnclex
'006c9fb7    85c0                    test ax, ax
'006c9fb9    7d15                    jge 6c9fd0
'006c9fbb    8b9588feffff            mov edx, dword ptr [ebp+fffffe88]
'006c9fc1    6a30                    push 30
'006c9fc3    68d8304300              push 004330d8
'006c9fc8    52                      push edx
'006c9fc9    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c9fca    ff1580104000            call dword ptr [00401080]
'006c9fd0    8b45c0                  mov eax, dword ptr [ebp-40]
'006c9fd3    8d9534ffffff            lea edx, dword ptr [ebp+ffffff34]
'006c9fd9    8d4d94                  lea ecx, dword ptr [ebp-6c]
'006c9fdc    c745c000000000          mov dword ptr [ebp-40], 00000000
'006c9fe3    89458c                  mov dword ptr [ebp-74], eax
'006c9fe6    c7458409000000          mov dword ptr [ebp-7c], 00000009
'006c9fed    c7853cffffffcc134300    mov dword ptr [ebp+ffffff3c], 004313cc
'006c9ff7    c78534ffffff08000000    mov dword ptr [ebp+ffffff34], 00000008

' *** Reference to "__vbaVarDup"
'006ca001    ff1598124000            call dword ptr [00401298]
var_121 = (vbNullChar)
'006ca007    668b85a0feffff          mov ax, word ptr [ebp+fffffea0]
'006ca00e    8d4d84                  lea ecx, dword ptr [ebp-7c]
'006ca011    51                      push ecx
'006ca012    8d5594                  lea edx, dword ptr [ebp-6c]
'006ca015    6689854cffffff          mov word ptr [ebp+ffffff4c], ax
'006ca01c    52                      push edx
'006ca01d    8d8544ffffff            lea eax, dword ptr [ebp+ffffff44]
'006ca023    50                      push eax
'006ca024    8d8d74ffffff            lea ecx, dword ptr [ebp+ffffff74]
'006ca02a    51                      push ecx
'006ca02b    c78544ffffff0b000000    mov dword ptr [ebp+ffffff44], 0000000b

' *** Reference to "rtcImmediateIf"
'006ca035    ff154c124000            call dword ptr [0040124c]
'006ca03b    8b1f                    mov ebx, dword ptr [edi]
'006ca03d    8d9574ffffff            lea edx, dword ptr [ebp+ffffff74]
'006ca043    52                      push edx
'006ca044    8d45e4                  lea eax, dword ptr [ebp-1c]
'006ca047    50                      push eax

' *** Reference to "__vbaStrVarVal"
'006ca048    ff15fc114000            call dword ptr [004011fc]
'006ca04e    50                      push eax
'006ca04f    57                      push edi
'006ca050    ff93a4000000            call dword ptr [ebx+000000a4]
'006ca056    dbe2                    fnclex
'006ca058    85c0                    test ax, ax
'006ca05a    7d12                    jge 6ca06e
'006ca05c    68a4000000              push 000000a4
'006ca061    68d00d4300              push 00430dd0
'006ca066    57                      push edi
'006ca067    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006ca068    ff1580104000            call dword ptr [00401080]
'006ca06e    8d4de4                  lea ecx, dword ptr [ebp-1c]

' *** Reference to "__vbaFreeStr"
'006ca071    ff150c134000            call dword ptr [0040130c]
'#Cleanup(var_40)
'006ca077    8d4dbc                  lea ecx, dword ptr [ebp-44]
'006ca07a    51                      push ecx
'006ca07b    8d55c4                  lea edx, dword ptr [ebp-3c]
'006ca07e    52                      push edx
'006ca07f    8d45cc                  lea eax, dword ptr [ebp-34]
'006ca082    50                      push eax
'006ca083    6a03                    push 03

' *** Reference to "__vbaFreeObjList"
'006ca085    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_43, var_9, var_58)
'006ca08b    8d8d74ffffff            lea ecx, dword ptr [ebp+ffffff74]
'006ca091    51                      push ecx
'006ca092    8d5584                  lea edx, dword ptr [ebp-7c]
'006ca095    52                      push edx
'006ca096    8d4594                  lea eax, dword ptr [ebp-6c]
'006ca099    50                      push eax
'006ca09a    8d8d44ffffff            lea ecx, dword ptr [ebp+ffffff44]
'006ca0a0    51                      push ecx
'006ca0a1    8d55a4                  lea edx, dword ptr [ebp-5c]
'006ca0a4    52                      push edx
'006ca0a5    6a05                    push 05

' *** Reference to "__vbaFreeVarList"
'006ca0a7    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_83, var_125, var_121, var_119, var_91)
'006ca0ad    8b45e8                  mov eax, dword ptr [ebp-18]
'006ca0b0    8b08                    mov ecx, dword ptr [eax]
'006ca0b2    83c428                  add esp, 28
'006ca0b5    8d55cc                  lea edx, dword ptr [ebp-34]
'006ca0b8    52                      push edx
'006ca0b9    50                      push eax
'006ca0ba    ff91b4000000            call dword ptr [ecx+000000b4]
'006ca0c0    dbe2                    fnclex
'006ca0c2    85c0                    test ax, ax
'006ca0c4    7d15                    jge 6ca0db
'006ca0c6    8b4de8                  mov ecx, dword ptr [ebp-18]
'006ca0c9    68b4000000              push 000000b4
'006ca0ce    6830314300              push 00433130
'006ca0d3    51                      push ecx
'006ca0d4    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006ca0d5    ff1580104000            call dword ptr [00401080]
'006ca0db    8b45cc                  mov eax, dword ptr [ebp-34]
'006ca0de    8d5dc8                  lea ebx, dword ptr [ebp-38]
'006ca0e1    53                      push ebx
'006ca0e2    83ec10                  sub esp, 10
'006ca0e5    ba08000000              mov edx, 00000008
'006ca0ea    8bdc                    mov ebx, esp
'006ca0ec    8913                    mov dword ptr [ebx], edx
'006ca0ee    899564ffffff            mov dword ptr [ebp+ffffff64], edx
'006ca0f4    8b9568ffffff            mov edx, dword ptr [ebp+ffffff68]
'006ca0fa    b9201c4400              mov ecx, 00441c20
'006ca0ff    895304                  mov dword ptr [ebx+04], edx
'006ca102    898d6cffffff            mov dword ptr [ebp+ffffff6c], ecx
'006ca108    8b38                    mov edi, dword ptr [eax]
'006ca10a    894b08                  mov dword ptr [ebx+08], ecx
'006ca10d    8b8d70ffffff            mov ecx, dword ptr [ebp+ffffff70]
'006ca113    50                      push eax
'006ca114    898594feffff            mov dword ptr [ebp+fffffe94], eax
'006ca11a    894b0c                  mov dword ptr [ebx+0c], ecx
'006ca11d    ff5730                  call dword ptr [edi+30]
'006ca120    dbe2                    fnclex
'006ca122    85c0                    test ax, ax
'006ca124    7d15                    jge 6ca13b
'006ca126    8b9594feffff            mov edx, dword ptr [ebp+fffffe94]
'006ca12c    6a30                    push 30
'006ca12e    68d8304300              push 004330d8
'006ca133    52                      push edx
'006ca134    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006ca135    ff1580104000            call dword ptr [00401080]
'006ca13b    8b45c8                  mov eax, dword ptr [ebp-38]
'006ca13e    8945ac                  mov dword ptr [ebp-54], eax
'006ca141    8d45a4                  lea eax, dword ptr [ebp-5c]
'006ca144    50                      push eax
'006ca145    c745c800000000          mov dword ptr [ebp-38], 00000000
'006ca14c    c745a409000000          mov dword ptr [ebp-5c], 00000009

' *** Reference to "rtcIsNull"
'006ca153    ff1540114000            call dword ptr [00401140]
'006ca159    8b0e                    mov ecx, dword ptr [esi]
'006ca15b    56                      push esi
'006ca15c    8985a0feffff            mov dword ptr [ebp+fffffea0], eax
'006ca162    ff9120030000            call dword ptr [ecx+00000320]
'006ca168    50                      push eax
'006ca169    8d55bc                  lea edx, dword ptr [ebp-44]
'006ca16c    52                      push edx

' *** Reference to "__vbaObjSet"
'006ca16d    ff15b4104000            call dword ptr [004010b4]
Set var_58 = IsNull(var_46)
'006ca173    8d55c4                  lea edx, dword ptr [ebp-3c]
'006ca176    8bf8                    mov edi, eax
'006ca178    8b45e8                  mov eax, dword ptr [ebp-18]
'006ca17b    8b08                    mov ecx, dword ptr [eax]
'006ca17d    52                      push edx
'006ca17e    50                      push eax
'006ca17f    ff91b4000000            call dword ptr [ecx+000000b4]
'006ca185    dbe2                    fnclex
'006ca187    85c0                    test ax, ax
'006ca189    7d15                    jge 6ca1a0
'006ca18b    8b4de8                  mov ecx, dword ptr [ebp-18]
'006ca18e    68b4000000              push 000000b4
'006ca193    6830314300              push 00433130
'006ca198    51                      push ecx
'006ca199    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006ca19a    ff1580104000            call dword ptr [00401080]
'006ca1a0    8b45c4                  mov eax, dword ptr [ebp-3c]
'006ca1a3    8b10                    mov edx, dword ptr [eax]
'006ca1a5    8d5dc0                  lea ebx, dword ptr [ebp-40]
'006ca1a8    53                      push ebx
'006ca1a9    83ec10                  sub esp, 10
'006ca1ac    8bdc                    mov ebx, esp
'006ca1ae    b908000000              mov ecx, 00000008
'006ca1b3    890b                    mov dword ptr [ebx], ecx
'006ca1b5    8b8d58ffffff            mov ecx, dword ptr [ebp+ffffff58]
'006ca1bb    894b04                  mov dword ptr [ebx+04], ecx
'006ca1be    c7855cffffff201c4400    mov dword ptr [ebp+ffffff5c], 00441c20
'006ca1c8    8b8d5cffffff            mov ecx, dword ptr [ebp+ffffff5c]
'006ca1ce    894b08                  mov dword ptr [ebx+08], ecx
'006ca1d1    8b8d60ffffff            mov ecx, dword ptr [ebp+ffffff60]
'006ca1d7    50                      push eax
'006ca1d8    898588feffff            mov dword ptr [ebp+fffffe88], eax
'006ca1de    894b0c                  mov dword ptr [ebx+0c], ecx
'006ca1e1    ff5230                  call dword ptr [edx+30]
'006ca1e4    dbe2                    fnclex
'006ca1e6    85c0                    test ax, ax
'006ca1e8    7d15                    jge 6ca1ff
'006ca1ea    8b9588feffff            mov edx, dword ptr [ebp+fffffe88]
'006ca1f0    6a30                    push 30
'006ca1f2    68d8304300              push 004330d8
'006ca1f7    52                      push edx
'006ca1f8    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006ca1f9    ff1580104000            call dword ptr [00401080]
'006ca1ff    8b45c0                  mov eax, dword ptr [ebp-40]
'006ca202    8d9534ffffff            lea edx, dword ptr [ebp+ffffff34]
'006ca208    8d4d94                  lea ecx, dword ptr [ebp-6c]
'006ca20b    c745c000000000          mov dword ptr [ebp-40], 00000000
'006ca212    89458c                  mov dword ptr [ebp-74], eax
'006ca215    c7458409000000          mov dword ptr [ebp-7c], 00000009
'006ca21c    c7853cffffffcc134300    mov dword ptr [ebp+ffffff3c], 004313cc
'006ca226    c78534ffffff08000000    mov dword ptr [ebp+ffffff34], 00000008

' *** Reference to "__vbaVarDup"
'006ca230    ff1598124000            call dword ptr [00401298]
var_121 = (vbNullChar)
'006ca236    668b85a0feffff          mov ax, word ptr [ebp+fffffea0]
'006ca23d    8d4d84                  lea ecx, dword ptr [ebp-7c]
'006ca240    51                      push ecx
'006ca241    8d5594                  lea edx, dword ptr [ebp-6c]
'006ca244    6689854cffffff          mov word ptr [ebp+ffffff4c], ax
'006ca24b    52                      push edx
'006ca24c    8d8544ffffff            lea eax, dword ptr [ebp+ffffff44]
'006ca252    50                      push eax
'006ca253    8d8d74ffffff            lea ecx, dword ptr [ebp+ffffff74]
'006ca259    51                      push ecx
'006ca25a    c78544ffffff0b000000    mov dword ptr [ebp+ffffff44], 0000000b

' *** Reference to "rtcImmediateIf"
'006ca264    ff154c124000            call dword ptr [0040124c]
'006ca26a    8b1f                    mov ebx, dword ptr [edi]
'006ca26c    8d9574ffffff            lea edx, dword ptr [ebp+ffffff74]
'006ca272    52                      push edx
'006ca273    8d45e4                  lea eax, dword ptr [ebp-1c]
'006ca276    50                      push eax

' *** Reference to "__vbaStrVarVal"
'006ca277    ff15fc114000            call dword ptr [004011fc]
'006ca27d    50                      push eax
'006ca27e    57                      push edi
'006ca27f    ff93a4000000            call dword ptr [ebx+000000a4]
'006ca285    dbe2                    fnclex
'006ca287    85c0                    test ax, ax
'006ca289    7d12                    jge 6ca29d
'006ca28b    68a4000000              push 000000a4
'006ca290    68d00d4300              push 00430dd0
'006ca295    57                      push edi
'006ca296    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006ca297    ff1580104000            call dword ptr [00401080]
'006ca29d    8d4de4                  lea ecx, dword ptr [ebp-1c]

' *** Reference to "__vbaFreeStr"
'006ca2a0    ff150c134000            call dword ptr [0040130c]
'#Cleanup(var_40)
'006ca2a6    8d4dbc                  lea ecx, dword ptr [ebp-44]
'006ca2a9    51                      push ecx
'006ca2aa    8d55c4                  lea edx, dword ptr [ebp-3c]
'006ca2ad    52                      push edx
'006ca2ae    8d45cc                  lea eax, dword ptr [ebp-34]
'006ca2b1    50                      push eax
'006ca2b2    6a03                    push 03

' *** Reference to "__vbaFreeObjList"
'006ca2b4    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_43, var_9, var_58)
'006ca2ba    8d8d74ffffff            lea ecx, dword ptr [ebp+ffffff74]
'006ca2c0    51                      push ecx
'006ca2c1    8d5584                  lea edx, dword ptr [ebp-7c]
'006ca2c4    52                      push edx
'006ca2c5    8d4594                  lea eax, dword ptr [ebp-6c]
'006ca2c8    50                      push eax
'006ca2c9    8d8d44ffffff            lea ecx, dword ptr [ebp+ffffff44]
'006ca2cf    51                      push ecx
'006ca2d0    8d55a4                  lea edx, dword ptr [ebp-5c]
'006ca2d3    52                      push edx
'006ca2d4    6a05                    push 05

' *** Reference to "__vbaFreeVarList"
'006ca2d6    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_83, var_125, var_121, var_119, var_91)
'006ca2dc    8b45e8                  mov eax, dword ptr [ebp-18]
'006ca2df    8b08                    mov ecx, dword ptr [eax]
'006ca2e1    83c428                  add esp, 28
'006ca2e4    8d55cc                  lea edx, dword ptr [ebp-34]
'006ca2e7    52                      push edx
'006ca2e8    50                      push eax
'006ca2e9    ff91b4000000            call dword ptr [ecx+000000b4]
'006ca2ef    dbe2                    fnclex
'006ca2f1    85c0                    test ax, ax
'006ca2f3    7d15                    jge 6ca30a
'006ca2f5    8b4de8                  mov ecx, dword ptr [ebp-18]
'006ca2f8    68b4000000              push 000000b4
'006ca2fd    6830314300              push 00433130
'006ca302    51                      push ecx
'006ca303    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006ca304    ff1580104000            call dword ptr [00401080]
'006ca30a    8b45cc                  mov eax, dword ptr [ebp-34]
'006ca30d    8d5dc8                  lea ebx, dword ptr [ebp-38]
'006ca310    53                      push ebx
'006ca311    83ec10                  sub esp, 10
'006ca314    ba08000000              mov edx, 00000008
'006ca319    8bdc                    mov ebx, esp
'006ca31b    8913                    mov dword ptr [ebx], edx
'006ca31d    899564ffffff            mov dword ptr [ebp+ffffff64], edx
'006ca323    8b9568ffffff            mov edx, dword ptr [ebp+ffffff68]
'006ca329    b974f34300              mov ecx, 0043f374
'006ca32e    895304                  mov dword ptr [ebx+04], edx
'006ca331    898d6cffffff            mov dword ptr [ebp+ffffff6c], ecx
'006ca337    8b38                    mov edi, dword ptr [eax]
'006ca339    894b08                  mov dword ptr [ebx+08], ecx
'006ca33c    8b8d70ffffff            mov ecx, dword ptr [ebp+ffffff70]
'006ca342    50                      push eax
'006ca343    898594feffff            mov dword ptr [ebp+fffffe94], eax
'006ca349    894b0c                  mov dword ptr [ebx+0c], ecx
'006ca34c    ff5730                  call dword ptr [edi+30]
'006ca34f    dbe2                    fnclex
'006ca351    85c0                    test ax, ax
'006ca353    7d15                    jge 6ca36a
'006ca355    8b9594feffff            mov edx, dword ptr [ebp+fffffe94]
'006ca35b    6a30                    push 30
'006ca35d    68d8304300              push 004330d8
'006ca362    52                      push edx
'006ca363    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006ca364    ff1580104000            call dword ptr [00401080]
'006ca36a    8b45c8                  mov eax, dword ptr [ebp-38]
'006ca36d    8945ac                  mov dword ptr [ebp-54], eax
'006ca370    8d45a4                  lea eax, dword ptr [ebp-5c]
'006ca373    50                      push eax
'006ca374    c745c800000000          mov dword ptr [ebp-38], 00000000
'006ca37b    c745a409000000          mov dword ptr [ebp-5c], 00000009

' *** Reference to "rtcIsNull"
'006ca382    ff1540114000            call dword ptr [00401140]
'006ca388    8b0e                    mov ecx, dword ptr [esi]
'006ca38a    56                      push esi
'006ca38b    8985a0feffff            mov dword ptr [ebp+fffffea0], eax
'006ca391    ff9144030000            call dword ptr [ecx+00000344]
'006ca397    50                      push eax
'006ca398    8d55bc                  lea edx, dword ptr [ebp-44]
'006ca39b    52                      push edx

' *** Reference to "__vbaObjSet"
'006ca39c    ff15b4104000            call dword ptr [004010b4]
Set var_58 = IsNull(var_46)
'006ca3a2    8d55c4                  lea edx, dword ptr [ebp-3c]
'006ca3a5    8bf8                    mov edi, eax
'006ca3a7    8b45e8                  mov eax, dword ptr [ebp-18]
'006ca3aa    8b08                    mov ecx, dword ptr [eax]
'006ca3ac    52                      push edx
'006ca3ad    50                      push eax
'006ca3ae    ff91b4000000            call dword ptr [ecx+000000b4]
'006ca3b4    dbe2                    fnclex
'006ca3b6    85c0                    test ax, ax
'006ca3b8    7d15                    jge 6ca3cf
'006ca3ba    8b4de8                  mov ecx, dword ptr [ebp-18]
'006ca3bd    68b4000000              push 000000b4
'006ca3c2    6830314300              push 00433130
'006ca3c7    51                      push ecx
'006ca3c8    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006ca3c9    ff1580104000            call dword ptr [00401080]
'006ca3cf    8b45c4                  mov eax, dword ptr [ebp-3c]
'006ca3d2    8b10                    mov edx, dword ptr [eax]
'006ca3d4    8d5dc0                  lea ebx, dword ptr [ebp-40]
'006ca3d7    53                      push ebx
'006ca3d8    83ec10                  sub esp, 10
'006ca3db    8bdc                    mov ebx, esp
'006ca3dd    b908000000              mov ecx, 00000008
'006ca3e2    890b                    mov dword ptr [ebx], ecx
'006ca3e4    8b8d58ffffff            mov ecx, dword ptr [ebp+ffffff58]
'006ca3ea    894b04                  mov dword ptr [ebx+04], ecx
'006ca3ed    c7855cffffff74f34300    mov dword ptr [ebp+ffffff5c], 0043f374
'006ca3f7    8b8d5cffffff            mov ecx, dword ptr [ebp+ffffff5c]
'006ca3fd    894b08                  mov dword ptr [ebx+08], ecx
'006ca400    8b8d60ffffff            mov ecx, dword ptr [ebp+ffffff60]
'006ca406    50                      push eax
'006ca407    898588feffff            mov dword ptr [ebp+fffffe88], eax
'006ca40d    894b0c                  mov dword ptr [ebx+0c], ecx
'006ca410    ff5230                  call dword ptr [edx+30]
'006ca413    dbe2                    fnclex
'006ca415    85c0                    test ax, ax
'006ca417    7d15                    jge 6ca42e
'006ca419    8b9588feffff            mov edx, dword ptr [ebp+fffffe88]
'006ca41f    6a30                    push 30
'006ca421    68d8304300              push 004330d8
'006ca426    52                      push edx
'006ca427    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006ca428    ff1580104000            call dword ptr [00401080]
'006ca42e    8b45c0                  mov eax, dword ptr [ebp-40]
'006ca431    8d9534ffffff            lea edx, dword ptr [ebp+ffffff34]
'006ca437    8d4d94                  lea ecx, dword ptr [ebp-6c]
'006ca43a    c745c000000000          mov dword ptr [ebp-40], 00000000
'006ca441    89458c                  mov dword ptr [ebp-74], eax
'006ca444    c7458409000000          mov dword ptr [ebp-7c], 00000009
'006ca44b    c7853cffffffcc134300    mov dword ptr [ebp+ffffff3c], 004313cc
'006ca455    c78534ffffff08000000    mov dword ptr [ebp+ffffff34], 00000008

' *** Reference to "__vbaVarDup"
'006ca45f    ff1598124000            call dword ptr [00401298]
var_121 = (vbNullChar)
'006ca465    668b85a0feffff          mov ax, word ptr [ebp+fffffea0]
'006ca46c    8d4d84                  lea ecx, dword ptr [ebp-7c]
'006ca46f    51                      push ecx
'006ca470    8d5594                  lea edx, dword ptr [ebp-6c]
'006ca473    6689854cffffff          mov word ptr [ebp+ffffff4c], ax
'006ca47a    52                      push edx
'006ca47b    8d8544ffffff            lea eax, dword ptr [ebp+ffffff44]
'006ca481    50                      push eax
'006ca482    8d8d74ffffff            lea ecx, dword ptr [ebp+ffffff74]
'006ca488    51                      push ecx
'006ca489    c78544ffffff0b000000    mov dword ptr [ebp+ffffff44], 0000000b

' *** Reference to "rtcImmediateIf"
'006ca493    ff154c124000            call dword ptr [0040124c]
'006ca499    8b1f                    mov ebx, dword ptr [edi]
'006ca49b    8d9574ffffff            lea edx, dword ptr [ebp+ffffff74]
'006ca4a1    52                      push edx
'006ca4a2    8d45e4                  lea eax, dword ptr [ebp-1c]
'006ca4a5    50                      push eax

' *** Reference to "__vbaStrVarVal"
'006ca4a6    ff15fc114000            call dword ptr [004011fc]
'006ca4ac    50                      push eax
'006ca4ad    57                      push edi
'006ca4ae    ff93a4000000            call dword ptr [ebx+000000a4]
'006ca4b4    dbe2                    fnclex
'006ca4b6    85c0                    test ax, ax
'006ca4b8    7d12                    jge 6ca4cc
'006ca4ba    68a4000000              push 000000a4
'006ca4bf    68d00d4300              push 00430dd0
'006ca4c4    57                      push edi
'006ca4c5    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006ca4c6    ff1580104000            call dword ptr [00401080]
'006ca4cc    8d4de4                  lea ecx, dword ptr [ebp-1c]

' *** Reference to "__vbaFreeStr"
'006ca4cf    ff150c134000            call dword ptr [0040130c]
'#Cleanup(var_40)
'006ca4d5    8d4dbc                  lea ecx, dword ptr [ebp-44]
'006ca4d8    51                      push ecx
'006ca4d9    8d55c4                  lea edx, dword ptr [ebp-3c]
'006ca4dc    52                      push edx
'006ca4dd    8d45cc                  lea eax, dword ptr [ebp-34]
'006ca4e0    50                      push eax
'006ca4e1    6a03                    push 03

' *** Reference to "__vbaFreeObjList"
'006ca4e3    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_43, var_9, var_58)
'006ca4e9    8d8d74ffffff            lea ecx, dword ptr [ebp+ffffff74]
'006ca4ef    51                      push ecx
'006ca4f0    8d5584                  lea edx, dword ptr [ebp-7c]
'006ca4f3    52                      push edx
'006ca4f4    8d4594                  lea eax, dword ptr [ebp-6c]
'006ca4f7    50                      push eax
'006ca4f8    8d8d44ffffff            lea ecx, dword ptr [ebp+ffffff44]
'006ca4fe    51                      push ecx
'006ca4ff    8d55a4                  lea edx, dword ptr [ebp-5c]
'006ca502    52                      push edx
'006ca503    6a05                    push 05

' *** Reference to "__vbaFreeVarList"
'006ca505    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_83, var_125, var_121, var_119, var_91)
'006ca50b    8b45e8                  mov eax, dword ptr [ebp-18]
'006ca50e    8b08                    mov ecx, dword ptr [eax]
'006ca510    83c428                  add esp, 28
'006ca513    8d55cc                  lea edx, dword ptr [ebp-34]
'006ca516    52                      push edx
'006ca517    50                      push eax
'006ca518    ff91b4000000            call dword ptr [ecx+000000b4]
'006ca51e    dbe2                    fnclex
'006ca520    85c0                    test ax, ax
'006ca522    7d15                    jge 6ca539
'006ca524    8b4de8                  mov ecx, dword ptr [ebp-18]
'006ca527    68b4000000              push 000000b4
'006ca52c    6830314300              push 00433130
'006ca531    51                      push ecx
'006ca532    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006ca533    ff1580104000            call dword ptr [00401080]
'006ca539    8b45cc                  mov eax, dword ptr [ebp-34]
'006ca53c    8d5dc8                  lea ebx, dword ptr [ebp-38]
'006ca53f    53                      push ebx
'006ca540    83ec10                  sub esp, 10
'006ca543    ba08000000              mov edx, 00000008
'006ca548    8bdc                    mov ebx, esp
'006ca54a    8913                    mov dword ptr [ebx], edx
'006ca54c    899564ffffff            mov dword ptr [ebp+ffffff64], edx
'006ca552    8b9568ffffff            mov edx, dword ptr [ebp+ffffff68]
'006ca558    b9641c4400              mov ecx, 00441c64
'006ca55d    895304                  mov dword ptr [ebx+04], edx
'006ca560    898d6cffffff            mov dword ptr [ebp+ffffff6c], ecx
'006ca566    8b38                    mov edi, dword ptr [eax]
'006ca568    894b08                  mov dword ptr [ebx+08], ecx
'006ca56b    8b8d70ffffff            mov ecx, dword ptr [ebp+ffffff70]
'006ca571    50                      push eax
'006ca572    898594feffff            mov dword ptr [ebp+fffffe94], eax
'006ca578    894b0c                  mov dword ptr [ebx+0c], ecx
'006ca57b    ff5730                  call dword ptr [edi+30]
'006ca57e    dbe2                    fnclex
'006ca580    85c0                    test ax, ax
'006ca582    7d15                    jge 6ca599
'006ca584    8b9594feffff            mov edx, dword ptr [ebp+fffffe94]
'006ca58a    6a30                    push 30
'006ca58c    68d8304300              push 004330d8
'006ca591    52                      push edx
'006ca592    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006ca593    ff1580104000            call dword ptr [00401080]
'006ca599    8b45c8                  mov eax, dword ptr [ebp-38]
'006ca59c    8945ac                  mov dword ptr [ebp-54], eax
'006ca59f    8d45a4                  lea eax, dword ptr [ebp-5c]
'006ca5a2    50                      push eax
'006ca5a3    c745c800000000          mov dword ptr [ebp-38], 00000000
'006ca5aa    c745a409000000          mov dword ptr [ebp-5c], 00000009

' *** Reference to "rtcIsNull"
'006ca5b1    ff1540114000            call dword ptr [00401140]
'006ca5b7    8b0e                    mov ecx, dword ptr [esi]
'006ca5b9    56                      push esi
'006ca5ba    8985a0feffff            mov dword ptr [ebp+fffffea0], eax
'006ca5c0    ff914c030000            call dword ptr [ecx+0000034c]
'006ca5c6    50                      push eax
'006ca5c7    8d55bc                  lea edx, dword ptr [ebp-44]
'006ca5ca    52                      push edx

' *** Reference to "__vbaObjSet"
'006ca5cb    ff15b4104000            call dword ptr [004010b4]
Set var_58 = IsNull(var_46)
'006ca5d1    8d55c4                  lea edx, dword ptr [ebp-3c]
'006ca5d4    8bf8                    mov edi, eax
'006ca5d6    8b45e8                  mov eax, dword ptr [ebp-18]
'006ca5d9    8b08                    mov ecx, dword ptr [eax]
'006ca5db    52                      push edx
'006ca5dc    50                      push eax
'006ca5dd    ff91b4000000            call dword ptr [ecx+000000b4]
'006ca5e3    dbe2                    fnclex
'006ca5e5    85c0                    test ax, ax
'006ca5e7    7d15                    jge 6ca5fe
'006ca5e9    8b4de8                  mov ecx, dword ptr [ebp-18]
'006ca5ec    68b4000000              push 000000b4
'006ca5f1    6830314300              push 00433130
'006ca5f6    51                      push ecx
'006ca5f7    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006ca5f8    ff1580104000            call dword ptr [00401080]
'006ca5fe    8b45c4                  mov eax, dword ptr [ebp-3c]
'006ca601    8b10                    mov edx, dword ptr [eax]
'006ca603    8d5dc0                  lea ebx, dword ptr [ebp-40]
'006ca606    53                      push ebx
'006ca607    83ec10                  sub esp, 10
'006ca60a    8bdc                    mov ebx, esp
'006ca60c    b908000000              mov ecx, 00000008
'006ca611    890b                    mov dword ptr [ebx], ecx
'006ca613    8b8d58ffffff            mov ecx, dword ptr [ebp+ffffff58]
'006ca619    894b04                  mov dword ptr [ebx+04], ecx
'006ca61c    c7855cffffff641c4400    mov dword ptr [ebp+ffffff5c], 00441c64
'006ca626    8b8d5cffffff            mov ecx, dword ptr [ebp+ffffff5c]
'006ca62c    894b08                  mov dword ptr [ebx+08], ecx
'006ca62f    8b8d60ffffff            mov ecx, dword ptr [ebp+ffffff60]
'006ca635    50                      push eax
'006ca636    898588feffff            mov dword ptr [ebp+fffffe88], eax
'006ca63c    894b0c                  mov dword ptr [ebx+0c], ecx
'006ca63f    ff5230                  call dword ptr [edx+30]
'006ca642    dbe2                    fnclex
'006ca644    85c0                    test ax, ax
'006ca646    7d15                    jge 6ca65d
'006ca648    8b9588feffff            mov edx, dword ptr [ebp+fffffe88]
'006ca64e    6a30                    push 30
'006ca650    68d8304300              push 004330d8
'006ca655    52                      push edx
'006ca656    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006ca657    ff1580104000            call dword ptr [00401080]
'006ca65d    8b45c0                  mov eax, dword ptr [ebp-40]
'006ca660    8d9534ffffff            lea edx, dword ptr [ebp+ffffff34]
'006ca666    8d4d94                  lea ecx, dword ptr [ebp-6c]
'006ca669    c745c000000000          mov dword ptr [ebp-40], 00000000
'006ca670    89458c                  mov dword ptr [ebp-74], eax
'006ca673    c7458409000000          mov dword ptr [ebp-7c], 00000009
'006ca67a    c7853cffffffcc134300    mov dword ptr [ebp+ffffff3c], 004313cc
'006ca684    c78534ffffff08000000    mov dword ptr [ebp+ffffff34], 00000008

' *** Reference to "__vbaVarDup"
'006ca68e    ff1598124000            call dword ptr [00401298]
var_121 = (vbNullChar)
'006ca694    668b85a0feffff          mov ax, word ptr [ebp+fffffea0]
'006ca69b    8d4d84                  lea ecx, dword ptr [ebp-7c]
'006ca69e    51                      push ecx
'006ca69f    8d5594                  lea edx, dword ptr [ebp-6c]
'006ca6a2    6689854cffffff          mov word ptr [ebp+ffffff4c], ax
'006ca6a9    52                      push edx
'006ca6aa    8d8544ffffff            lea eax, dword ptr [ebp+ffffff44]
'006ca6b0    50                      push eax
'006ca6b1    8d8d74ffffff            lea ecx, dword ptr [ebp+ffffff74]
'006ca6b7    51                      push ecx
'006ca6b8    c78544ffffff0b000000    mov dword ptr [ebp+ffffff44], 0000000b

' *** Reference to "rtcImmediateIf"
'006ca6c2    ff154c124000            call dword ptr [0040124c]
'006ca6c8    8b1f                    mov ebx, dword ptr [edi]
'006ca6ca    8d9574ffffff            lea edx, dword ptr [ebp+ffffff74]
'006ca6d0    52                      push edx
'006ca6d1    8d45e4                  lea eax, dword ptr [ebp-1c]
'006ca6d4    50                      push eax

' *** Reference to "__vbaStrVarVal"
'006ca6d5    ff15fc114000            call dword ptr [004011fc]
'006ca6db    50                      push eax
'006ca6dc    57                      push edi
'006ca6dd    ff93a4000000            call dword ptr [ebx+000000a4]
'006ca6e3    dbe2                    fnclex
'006ca6e5    85c0                    test ax, ax
'006ca6e7    7d12                    jge 6ca6fb
'006ca6e9    68a4000000              push 000000a4
'006ca6ee    68d00d4300              push 00430dd0
'006ca6f3    57                      push edi
'006ca6f4    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006ca6f5    ff1580104000            call dword ptr [00401080]
'006ca6fb    8d4de4                  lea ecx, dword ptr [ebp-1c]

' *** Reference to "__vbaFreeStr"
'006ca6fe    ff150c134000            call dword ptr [0040130c]
'#Cleanup(var_40)
'006ca704    8d4dbc                  lea ecx, dword ptr [ebp-44]
'006ca707    51                      push ecx
'006ca708    8d55c4                  lea edx, dword ptr [ebp-3c]
'006ca70b    52                      push edx
'006ca70c    8d45cc                  lea eax, dword ptr [ebp-34]
'006ca70f    50                      push eax
'006ca710    6a03                    push 03

' *** Reference to "__vbaFreeObjList"
'006ca712    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_43, var_9, var_58)
'006ca718    8d8d74ffffff            lea ecx, dword ptr [ebp+ffffff74]
'006ca71e    51                      push ecx
'006ca71f    8d5584                  lea edx, dword ptr [ebp-7c]
'006ca722    52                      push edx
'006ca723    8d4594                  lea eax, dword ptr [ebp-6c]
'006ca726    50                      push eax
'006ca727    8d8d44ffffff            lea ecx, dword ptr [ebp+ffffff44]
'006ca72d    51                      push ecx
'006ca72e    8d55a4                  lea edx, dword ptr [ebp-5c]
'006ca731    52                      push edx
'006ca732    6a05                    push 05

' *** Reference to "__vbaFreeVarList"
'006ca734    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_83, var_125, var_121, var_119, var_91)
'006ca73a    8b45e8                  mov eax, dword ptr [ebp-18]
'006ca73d    8b08                    mov ecx, dword ptr [eax]
'006ca73f    83c428                  add esp, 28
'006ca742    8d55cc                  lea edx, dword ptr [ebp-34]
'006ca745    52                      push edx
'006ca746    50                      push eax
'006ca747    ff91b4000000            call dword ptr [ecx+000000b4]
'006ca74d    dbe2                    fnclex
'006ca74f    85c0                    test ax, ax
'006ca751    7d15                    jge 6ca768
'006ca753    8b4de8                  mov ecx, dword ptr [ebp-18]
'006ca756    68b4000000              push 000000b4
'006ca75b    6830314300              push 00433130
'006ca760    51                      push ecx
'006ca761    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006ca762    ff1580104000            call dword ptr [00401080]
'006ca768    8b45cc                  mov eax, dword ptr [ebp-34]
'006ca76b    8d5dc8                  lea ebx, dword ptr [ebp-38]
'006ca76e    53                      push ebx
'006ca76f    83ec10                  sub esp, 10
'006ca772    ba08000000              mov edx, 00000008
'006ca777    8bdc                    mov ebx, esp
'006ca779    8913                    mov dword ptr [ebx], edx
'006ca77b    899564ffffff            mov dword ptr [ebp+ffffff64], edx
'006ca781    8b9568ffffff            mov edx, dword ptr [ebp+ffffff68]
'006ca787    b918f44200              mov ecx, 0042f418
'006ca78c    895304                  mov dword ptr [ebx+04], edx
'006ca78f    898d6cffffff            mov dword ptr [ebp+ffffff6c], ecx
'006ca795    8b38                    mov edi, dword ptr [eax]
'006ca797    894b08                  mov dword ptr [ebx+08], ecx
'006ca79a    8b8d70ffffff            mov ecx, dword ptr [ebp+ffffff70]
'006ca7a0    50                      push eax
'006ca7a1    898594feffff            mov dword ptr [ebp+fffffe94], eax
'006ca7a7    894b0c                  mov dword ptr [ebx+0c], ecx
'006ca7aa    ff5730                  call dword ptr [edi+30]
'006ca7ad    dbe2                    fnclex
'006ca7af    85c0                    test ax, ax
'006ca7b1    7d15                    jge 6ca7c8
'006ca7b3    8b9594feffff            mov edx, dword ptr [ebp+fffffe94]
'006ca7b9    6a30                    push 30
'006ca7bb    68d8304300              push 004330d8
'006ca7c0    52                      push edx
'006ca7c1    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006ca7c2    ff1580104000            call dword ptr [00401080]
'006ca7c8    8b45c8                  mov eax, dword ptr [ebp-38]
'006ca7cb    8945ac                  mov dword ptr [ebp-54], eax
'006ca7ce    8d45a4                  lea eax, dword ptr [ebp-5c]
'006ca7d1    50                      push eax
'006ca7d2    c745c800000000          mov dword ptr [ebp-38], 00000000
'006ca7d9    c745a409000000          mov dword ptr [ebp-5c], 00000009

' *** Reference to "rtcIsNull"
'006ca7e0    ff1540114000            call dword ptr [00401140]
'006ca7e6    8b0e                    mov ecx, dword ptr [esi]
'006ca7e8    56                      push esi
'006ca7e9    8985a0feffff            mov dword ptr [ebp+fffffea0], eax
'006ca7ef    ff9150030000            call dword ptr [ecx+00000350]
'006ca7f5    50                      push eax
'006ca7f6    8d55bc                  lea edx, dword ptr [ebp-44]
'006ca7f9    52                      push edx

' *** Reference to "__vbaObjSet"
'006ca7fa    ff15b4104000            call dword ptr [004010b4]
Set var_58 = IsNull(var_46)
'006ca800    8d55c4                  lea edx, dword ptr [ebp-3c]
'006ca803    8bf8                    mov edi, eax
'006ca805    8b45e8                  mov eax, dword ptr [ebp-18]
'006ca808    8b08                    mov ecx, dword ptr [eax]
'006ca80a    52                      push edx
'006ca80b    50                      push eax
'006ca80c    ff91b4000000            call dword ptr [ecx+000000b4]
'006ca812    dbe2                    fnclex
'006ca814    85c0                    test ax, ax
'006ca816    7d15                    jge 6ca82d
'006ca818    8b4de8                  mov ecx, dword ptr [ebp-18]
'006ca81b    68b4000000              push 000000b4
'006ca820    6830314300              push 00433130
'006ca825    51                      push ecx
'006ca826    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006ca827    ff1580104000            call dword ptr [00401080]
'006ca82d    8b45c4                  mov eax, dword ptr [ebp-3c]
'006ca830    8b10                    mov edx, dword ptr [eax]
'006ca832    8d5dc0                  lea ebx, dword ptr [ebp-40]
'006ca835    53                      push ebx
'006ca836    83ec10                  sub esp, 10
'006ca839    8bdc                    mov ebx, esp
'006ca83b    b908000000              mov ecx, 00000008
'006ca840    890b                    mov dword ptr [ebx], ecx
'006ca842    8b8d58ffffff            mov ecx, dword ptr [ebp+ffffff58]
'006ca848    894b04                  mov dword ptr [ebx+04], ecx
'006ca84b    c7855cffffff18f44200    mov dword ptr [ebp+ffffff5c], 0042f418
'006ca855    8b8d5cffffff            mov ecx, dword ptr [ebp+ffffff5c]
'006ca85b    894b08                  mov dword ptr [ebx+08], ecx
'006ca85e    8b8d60ffffff            mov ecx, dword ptr [ebp+ffffff60]
'006ca864    50                      push eax
'006ca865    898588feffff            mov dword ptr [ebp+fffffe88], eax
'006ca86b    894b0c                  mov dword ptr [ebx+0c], ecx
'006ca86e    ff5230                  call dword ptr [edx+30]
'006ca871    dbe2                    fnclex
'006ca873    85c0                    test ax, ax
'006ca875    7d15                    jge 6ca88c
'006ca877    8b9588feffff            mov edx, dword ptr [ebp+fffffe88]
'006ca87d    6a30                    push 30
'006ca87f    68d8304300              push 004330d8
'006ca884    52                      push edx
'006ca885    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006ca886    ff1580104000            call dword ptr [00401080]
'006ca88c    8b45c0                  mov eax, dword ptr [ebp-40]
'006ca88f    8d9534ffffff            lea edx, dword ptr [ebp+ffffff34]
'006ca895    8d4d94                  lea ecx, dword ptr [ebp-6c]
'006ca898    c745c000000000          mov dword ptr [ebp-40], 00000000
'006ca89f    89458c                  mov dword ptr [ebp-74], eax
'006ca8a2    c7458409000000          mov dword ptr [ebp-7c], 00000009
'006ca8a9    c7853cffffffcc134300    mov dword ptr [ebp+ffffff3c], 004313cc
'006ca8b3    c78534ffffff08000000    mov dword ptr [ebp+ffffff34], 00000008

' *** Reference to "__vbaVarDup"
'006ca8bd    ff1598124000            call dword ptr [00401298]
var_121 = (vbNullChar)
'006ca8c3    668b85a0feffff          mov ax, word ptr [ebp+fffffea0]
'006ca8ca    8d4d84                  lea ecx, dword ptr [ebp-7c]
'006ca8cd    51                      push ecx
'006ca8ce    8d5594                  lea edx, dword ptr [ebp-6c]
'006ca8d1    6689854cffffff          mov word ptr [ebp+ffffff4c], ax
'006ca8d8    52                      push edx
'006ca8d9    8d8544ffffff            lea eax, dword ptr [ebp+ffffff44]
'006ca8df    50                      push eax
'006ca8e0    8d8d74ffffff            lea ecx, dword ptr [ebp+ffffff74]
'006ca8e6    51                      push ecx
'006ca8e7    c78544ffffff0b000000    mov dword ptr [ebp+ffffff44], 0000000b

' *** Reference to "rtcImmediateIf"
'006ca8f1    ff154c124000            call dword ptr [0040124c]
'006ca8f7    8b1f                    mov ebx, dword ptr [edi]
'006ca8f9    8d9574ffffff            lea edx, dword ptr [ebp+ffffff74]
'006ca8ff    52                      push edx
'006ca900    8d45e4                  lea eax, dword ptr [ebp-1c]
'006ca903    50                      push eax

' *** Reference to "__vbaStrVarVal"
'006ca904    ff15fc114000            call dword ptr [004011fc]
'006ca90a    50                      push eax
'006ca90b    57                      push edi
'006ca90c    ff93a4000000            call dword ptr [ebx+000000a4]
'006ca912    dbe2                    fnclex
'006ca914    85c0                    test ax, ax
'006ca916    7d12                    jge 6ca92a
'006ca918    68a4000000              push 000000a4
'006ca91d    68d00d4300              push 00430dd0
'006ca922    57                      push edi
'006ca923    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006ca924    ff1580104000            call dword ptr [00401080]
'006ca92a    8d4de4                  lea ecx, dword ptr [ebp-1c]

' *** Reference to "__vbaFreeStr"
'006ca92d    ff150c134000            call dword ptr [0040130c]
'#Cleanup(var_40)
'006ca933    8d4dbc                  lea ecx, dword ptr [ebp-44]
'006ca936    51                      push ecx
'006ca937    8d55c4                  lea edx, dword ptr [ebp-3c]
'006ca93a    52                      push edx
'006ca93b    8d45cc                  lea eax, dword ptr [ebp-34]
'006ca93e    50                      push eax
'006ca93f    6a03                    push 03

' *** Reference to "__vbaFreeObjList"
'006ca941    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_43, var_9, var_58)
'006ca947    8d8d74ffffff            lea ecx, dword ptr [ebp+ffffff74]
'006ca94d    51                      push ecx
'006ca94e    8d5584                  lea edx, dword ptr [ebp-7c]
'006ca951    52                      push edx
'006ca952    8d4594                  lea eax, dword ptr [ebp-6c]
'006ca955    50                      push eax
'006ca956    8d8d44ffffff            lea ecx, dword ptr [ebp+ffffff44]
'006ca95c    51                      push ecx
'006ca95d    8d55a4                  lea edx, dword ptr [ebp-5c]
'006ca960    52                      push edx
'006ca961    6a05                    push 05

' *** Reference to "__vbaFreeVarList"
'006ca963    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_83, var_125, var_121, var_119, var_91)
'006ca969    8b45e8                  mov eax, dword ptr [ebp-18]
'006ca96c    8b08                    mov ecx, dword ptr [eax]
'006ca96e    83c428                  add esp, 28
'006ca971    8d55cc                  lea edx, dword ptr [ebp-34]
'006ca974    52                      push edx
'006ca975    50                      push eax
'006ca976    ff91b4000000            call dword ptr [ecx+000000b4]
'006ca97c    dbe2                    fnclex
'006ca97e    85c0                    test ax, ax
'006ca980    7d15                    jge 6ca997
'006ca982    8b4de8                  mov ecx, dword ptr [ebp-18]
'006ca985    68b4000000              push 000000b4
'006ca98a    6830314300              push 00433130
'006ca98f    51                      push ecx
'006ca990    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006ca991    ff1580104000            call dword ptr [00401080]
'006ca997    8b45cc                  mov eax, dword ptr [ebp-34]
'006ca99a    8d5dc8                  lea ebx, dword ptr [ebp-38]
'006ca99d    53                      push ebx
'006ca99e    83ec10                  sub esp, 10
'006ca9a1    ba08000000              mov edx, 00000008
'006ca9a6    8bdc                    mov ebx, esp
'006ca9a8    8913                    mov dword ptr [ebx], edx
'006ca9aa    899564ffffff            mov dword ptr [ebp+ffffff64], edx
'006ca9b0    8b9568ffffff            mov edx, dword ptr [ebp+ffffff68]
'006ca9b6    b9e4b24300              mov ecx, 0043b2e4
'006ca9bb    895304                  mov dword ptr [ebx+04], edx
'006ca9be    898d6cffffff            mov dword ptr [ebp+ffffff6c], ecx
'006ca9c4    8b38                    mov edi, dword ptr [eax]
'006ca9c6    894b08                  mov dword ptr [ebx+08], ecx
'006ca9c9    8b8d70ffffff            mov ecx, dword ptr [ebp+ffffff70]
'006ca9cf    50                      push eax
'006ca9d0    898594feffff            mov dword ptr [ebp+fffffe94], eax
'006ca9d6    894b0c                  mov dword ptr [ebx+0c], ecx
'006ca9d9    ff5730                  call dword ptr [edi+30]
'006ca9dc    dbe2                    fnclex
'006ca9de    85c0                    test ax, ax
'006ca9e0    7d15                    jge 6ca9f7
'006ca9e2    8b9594feffff            mov edx, dword ptr [ebp+fffffe94]
'006ca9e8    6a30                    push 30
'006ca9ea    68d8304300              push 004330d8
'006ca9ef    52                      push edx
'006ca9f0    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006ca9f1    ff1580104000            call dword ptr [00401080]
'006ca9f7    8b45c8                  mov eax, dword ptr [ebp-38]
'006ca9fa    8945ac                  mov dword ptr [ebp-54], eax
'006ca9fd    8d45a4                  lea eax, dword ptr [ebp-5c]
'006caa00    50                      push eax
'006caa01    c745c800000000          mov dword ptr [ebp-38], 00000000
'006caa08    c745a409000000          mov dword ptr [ebp-5c], 00000009

' *** Reference to "rtcIsNull"
'006caa0f    ff1540114000            call dword ptr [00401140]
'006caa15    8b0e                    mov ecx, dword ptr [esi]
'006caa17    56                      push esi
'006caa18    8985a0feffff            mov dword ptr [ebp+fffffea0], eax
'006caa1e    ff9154030000            call dword ptr [ecx+00000354]
'006caa24    50                      push eax
'006caa25    8d55bc                  lea edx, dword ptr [ebp-44]
'006caa28    52                      push edx

' *** Reference to "__vbaObjSet"
'006caa29    ff15b4104000            call dword ptr [004010b4]
Set var_58 = IsNull(var_46)
'006caa2f    8d55c4                  lea edx, dword ptr [ebp-3c]
'006caa32    8bf8                    mov edi, eax
'006caa34    8b45e8                  mov eax, dword ptr [ebp-18]
'006caa37    8b08                    mov ecx, dword ptr [eax]
'006caa39    52                      push edx
'006caa3a    50                      push eax
'006caa3b    ff91b4000000            call dword ptr [ecx+000000b4]
'006caa41    dbe2                    fnclex
'006caa43    85c0                    test ax, ax
'006caa45    7d15                    jge 6caa5c
'006caa47    8b4de8                  mov ecx, dword ptr [ebp-18]
'006caa4a    68b4000000              push 000000b4
'006caa4f    6830314300              push 00433130
'006caa54    51                      push ecx
'006caa55    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006caa56    ff1580104000            call dword ptr [00401080]
'006caa5c    8b45c4                  mov eax, dword ptr [ebp-3c]
'006caa5f    8b10                    mov edx, dword ptr [eax]
'006caa61    8d5dc0                  lea ebx, dword ptr [ebp-40]
'006caa64    53                      push ebx
'006caa65    83ec10                  sub esp, 10
'006caa68    8bdc                    mov ebx, esp
'006caa6a    b908000000              mov ecx, 00000008
'006caa6f    890b                    mov dword ptr [ebx], ecx
'006caa71    8b8d58ffffff            mov ecx, dword ptr [ebp+ffffff58]
'006caa77    894b04                  mov dword ptr [ebx+04], ecx
'006caa7a    c7855cffffffe4b24300    mov dword ptr [ebp+ffffff5c], 0043b2e4
'006caa84    8b8d5cffffff            mov ecx, dword ptr [ebp+ffffff5c]
'006caa8a    894b08                  mov dword ptr [ebx+08], ecx
'006caa8d    8b8d60ffffff            mov ecx, dword ptr [ebp+ffffff60]
'006caa93    50                      push eax
'006caa94    898588feffff            mov dword ptr [ebp+fffffe88], eax
'006caa9a    894b0c                  mov dword ptr [ebx+0c], ecx
'006caa9d    ff5230                  call dword ptr [edx+30]
'006caaa0    dbe2                    fnclex
'006caaa2    85c0                    test ax, ax
'006caaa4    7d15                    jge 6caabb
'006caaa6    8b9588feffff            mov edx, dword ptr [ebp+fffffe88]
'006caaac    6a30                    push 30
'006caaae    68d8304300              push 004330d8
'006caab3    52                      push edx
'006caab4    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006caab5    ff1580104000            call dword ptr [00401080]
'006caabb    8b45c0                  mov eax, dword ptr [ebp-40]
'006caabe    8d9534ffffff            lea edx, dword ptr [ebp+ffffff34]
'006caac4    8d4d94                  lea ecx, dword ptr [ebp-6c]
'006caac7    c745c000000000          mov dword ptr [ebp-40], 00000000
'006caace    89458c                  mov dword ptr [ebp-74], eax
'006caad1    c7458409000000          mov dword ptr [ebp-7c], 00000009
'006caad8    c7853cffffffcc134300    mov dword ptr [ebp+ffffff3c], 004313cc
'006caae2    c78534ffffff08000000    mov dword ptr [ebp+ffffff34], 00000008

' *** Reference to "__vbaVarDup"
'006caaec    ff1598124000            call dword ptr [00401298]
var_121 = (vbNullChar)
'006caaf2    668b85a0feffff          mov ax, word ptr [ebp+fffffea0]
'006caaf9    8d4d84                  lea ecx, dword ptr [ebp-7c]
'006caafc    51                      push ecx
'006caafd    8d5594                  lea edx, dword ptr [ebp-6c]
'006cab00    6689854cffffff          mov word ptr [ebp+ffffff4c], ax
'006cab07    52                      push edx
'006cab08    8d8544ffffff            lea eax, dword ptr [ebp+ffffff44]
'006cab0e    50                      push eax
'006cab0f    8d8d74ffffff            lea ecx, dword ptr [ebp+ffffff74]
'006cab15    51                      push ecx
'006cab16    c78544ffffff0b000000    mov dword ptr [ebp+ffffff44], 0000000b

' *** Reference to "rtcImmediateIf"
'006cab20    ff154c124000            call dword ptr [0040124c]
'006cab26    8b1f                    mov ebx, dword ptr [edi]
'006cab28    8d9574ffffff            lea edx, dword ptr [ebp+ffffff74]
'006cab2e    52                      push edx
'006cab2f    8d45e4                  lea eax, dword ptr [ebp-1c]
'006cab32    50                      push eax

' *** Reference to "__vbaStrVarVal"
'006cab33    ff15fc114000            call dword ptr [004011fc]
'006cab39    50                      push eax
'006cab3a    57                      push edi
'006cab3b    ff93a4000000            call dword ptr [ebx+000000a4]
'006cab41    dbe2                    fnclex
'006cab43    85c0                    test ax, ax
'006cab45    7d12                    jge 6cab59
'006cab47    68a4000000              push 000000a4
'006cab4c    68d00d4300              push 00430dd0
'006cab51    57                      push edi
'006cab52    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006cab53    ff1580104000            call dword ptr [00401080]
'006cab59    8d4de4                  lea ecx, dword ptr [ebp-1c]

' *** Reference to "__vbaFreeStr"
'006cab5c    ff150c134000            call dword ptr [0040130c]
'#Cleanup(var_40)
'006cab62    8d4dbc                  lea ecx, dword ptr [ebp-44]
'006cab65    51                      push ecx
'006cab66    8d55c4                  lea edx, dword ptr [ebp-3c]
'006cab69    52                      push edx
'006cab6a    8d45cc                  lea eax, dword ptr [ebp-34]
'006cab6d    50                      push eax
'006cab6e    6a03                    push 03

' *** Reference to "__vbaFreeObjList"
'006cab70    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_43, var_9, var_58)
'006cab76    8d8d74ffffff            lea ecx, dword ptr [ebp+ffffff74]
'006cab7c    51                      push ecx
'006cab7d    8d5584                  lea edx, dword ptr [ebp-7c]
'006cab80    52                      push edx
'006cab81    8d4594                  lea eax, dword ptr [ebp-6c]
'006cab84    50                      push eax
'006cab85    8d8d44ffffff            lea ecx, dword ptr [ebp+ffffff44]
'006cab8b    51                      push ecx
'006cab8c    8d55a4                  lea edx, dword ptr [ebp-5c]
'006cab8f    52                      push edx
'006cab90    6a05                    push 05

' *** Reference to "__vbaFreeVarList"
'006cab92    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_83, var_125, var_121, var_119, var_91)
'006cab98    8b45e8                  mov eax, dword ptr [ebp-18]
'006cab9b    8b08                    mov ecx, dword ptr [eax]
'006cab9d    83c428                  add esp, 28
'006caba0    8d55cc                  lea edx, dword ptr [ebp-34]
'006caba3    52                      push edx
'006caba4    50                      push eax
'006caba5    ff91b4000000            call dword ptr [ecx+000000b4]
'006cabab    dbe2                    fnclex
'006cabad    85c0                    test ax, ax
'006cabaf    7d15                    jge 6cabc6
'006cabb1    8b4de8                  mov ecx, dword ptr [ebp-18]
'006cabb4    68b4000000              push 000000b4
'006cabb9    6830314300              push 00433130
'006cabbe    51                      push ecx
'006cabbf    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006cabc0    ff1580104000            call dword ptr [00401080]
'006cabc6    8b45cc                  mov eax, dword ptr [ebp-34]
'006cabc9    8d5dc8                  lea ebx, dword ptr [ebp-38]
'006cabcc    53                      push ebx
'006cabcd    83ec10                  sub esp, 10
'006cabd0    ba08000000              mov edx, 00000008
'006cabd5    8bdc                    mov ebx, esp
'006cabd7    8913                    mov dword ptr [ebx], edx
'006cabd9    899564ffffff            mov dword ptr [ebp+ffffff64], edx
'006cabdf    8b9568ffffff            mov edx, dword ptr [ebp+ffffff68]
'006cabe5    b958384500              mov ecx, 00453858
'006cabea    895304                  mov dword ptr [ebx+04], edx
'006cabed    898d6cffffff            mov dword ptr [ebp+ffffff6c], ecx
'006cabf3    8b38                    mov edi, dword ptr [eax]
'006cabf5    894b08                  mov dword ptr [ebx+08], ecx
'006cabf8    8b8d70ffffff            mov ecx, dword ptr [ebp+ffffff70]
'006cabfe    50                      push eax
'006cabff    898594feffff            mov dword ptr [ebp+fffffe94], eax
'006cac05    894b0c                  mov dword ptr [ebx+0c], ecx
'006cac08    ff5730                  call dword ptr [edi+30]
'006cac0b    dbe2                    fnclex
'006cac0d    85c0                    test ax, ax
'006cac0f    7d15                    jge 6cac26
'006cac11    8b9594feffff            mov edx, dword ptr [ebp+fffffe94]
'006cac17    6a30                    push 30
'006cac19    68d8304300              push 004330d8
'006cac1e    52                      push edx
'006cac1f    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006cac20    ff1580104000            call dword ptr [00401080]
'006cac26    8b45c8                  mov eax, dword ptr [ebp-38]
'006cac29    8945ac                  mov dword ptr [ebp-54], eax
'006cac2c    8d45a4                  lea eax, dword ptr [ebp-5c]
'006cac2f    50                      push eax
'006cac30    c745c800000000          mov dword ptr [ebp-38], 00000000
'006cac37    c745a409000000          mov dword ptr [ebp-5c], 00000009

' *** Reference to "rtcIsNull"
'006cac3e    ff1540114000            call dword ptr [00401140]
'006cac44    8b0e                    mov ecx, dword ptr [esi]
'006cac46    56                      push esi
'006cac47    8985a0feffff            mov dword ptr [ebp+fffffea0], eax
'006cac4d    ff9128030000            call dword ptr [ecx+00000328]
'006cac53    50                      push eax
'006cac54    8d55bc                  lea edx, dword ptr [ebp-44]
'006cac57    52                      push edx

' *** Reference to "__vbaObjSet"
'006cac58    ff15b4104000            call dword ptr [004010b4]
Set var_58 = IsNull(var_46)
'006cac5e    8d55c4                  lea edx, dword ptr [ebp-3c]
'006cac61    8bf8                    mov edi, eax
'006cac63    8b45e8                  mov eax, dword ptr [ebp-18]
'006cac66    8b08                    mov ecx, dword ptr [eax]
'006cac68    52                      push edx
'006cac69    50                      push eax
'006cac6a    ff91b4000000            call dword ptr [ecx+000000b4]
'006cac70    dbe2                    fnclex
'006cac72    85c0                    test ax, ax
'006cac74    7d15                    jge 6cac8b
'006cac76    8b4de8                  mov ecx, dword ptr [ebp-18]
'006cac79    68b4000000              push 000000b4
'006cac7e    6830314300              push 00433130
'006cac83    51                      push ecx
'006cac84    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006cac85    ff1580104000            call dword ptr [00401080]
'006cac8b    8b45c4                  mov eax, dword ptr [ebp-3c]
'006cac8e    8b10                    mov edx, dword ptr [eax]
'006cac90    8d5dc0                  lea ebx, dword ptr [ebp-40]
'006cac93    53                      push ebx
'006cac94    83ec10                  sub esp, 10
'006cac97    8bdc                    mov ebx, esp
'006cac99    b908000000              mov ecx, 00000008
'006cac9e    890b                    mov dword ptr [ebx], ecx
'006caca0    8b8d58ffffff            mov ecx, dword ptr [ebp+ffffff58]
'006caca6    894b04                  mov dword ptr [ebx+04], ecx
'006caca9    c7855cffffff58384500    mov dword ptr [ebp+ffffff5c], 00453858
'006cacb3    8b8d5cffffff            mov ecx, dword ptr [ebp+ffffff5c]
'006cacb9    894b08                  mov dword ptr [ebx+08], ecx
'006cacbc    8b8d60ffffff            mov ecx, dword ptr [ebp+ffffff60]
'006cacc2    50                      push eax
'006cacc3    898588feffff            mov dword ptr [ebp+fffffe88], eax
'006cacc9    894b0c                  mov dword ptr [ebx+0c], ecx
'006caccc    ff5230                  call dword ptr [edx+30]
'006caccf    dbe2                    fnclex
'006cacd1    85c0                    test ax, ax
'006cacd3    7d15                    jge 6cacea
'006cacd5    8b9588feffff            mov edx, dword ptr [ebp+fffffe88]
'006cacdb    6a30                    push 30
'006cacdd    68d8304300              push 004330d8
'006cace2    52                      push edx
'006cace3    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006cace4    ff1580104000            call dword ptr [00401080]
'006cacea    8b45c0                  mov eax, dword ptr [ebp-40]
'006caced    8d9534ffffff            lea edx, dword ptr [ebp+ffffff34]
'006cacf3    8d4d94                  lea ecx, dword ptr [ebp-6c]
'006cacf6    c745c000000000          mov dword ptr [ebp-40], 00000000
'006cacfd    89458c                  mov dword ptr [ebp-74], eax
'006cad00    c7458409000000          mov dword ptr [ebp-7c], 00000009
'006cad07    c7853cffffffcc134300    mov dword ptr [ebp+ffffff3c], 004313cc
'006cad11    c78534ffffff08000000    mov dword ptr [ebp+ffffff34], 00000008

' *** Reference to "__vbaVarDup"
'006cad1b    ff1598124000            call dword ptr [00401298]
var_121 = (vbNullChar)
'006cad21    668b85a0feffff          mov ax, word ptr [ebp+fffffea0]
'006cad28    8d4d84                  lea ecx, dword ptr [ebp-7c]
'006cad2b    51                      push ecx
'006cad2c    8d5594                  lea edx, dword ptr [ebp-6c]
'006cad2f    6689854cffffff          mov word ptr [ebp+ffffff4c], ax
'006cad36    52                      push edx
'006cad37    8d8544ffffff            lea eax, dword ptr [ebp+ffffff44]
'006cad3d    50                      push eax
'006cad3e    8d8d74ffffff            lea ecx, dword ptr [ebp+ffffff74]
'006cad44    51                      push ecx
'006cad45    c78544ffffff0b000000    mov dword ptr [ebp+ffffff44], 0000000b

' *** Reference to "rtcImmediateIf"
'006cad4f    ff154c124000            call dword ptr [0040124c]
'006cad55    8b1f                    mov ebx, dword ptr [edi]
'006cad57    8d9574ffffff            lea edx, dword ptr [ebp+ffffff74]
'006cad5d    52                      push edx
'006cad5e    8d45e4                  lea eax, dword ptr [ebp-1c]
'006cad61    50                      push eax

' *** Reference to "__vbaStrVarVal"
'006cad62    ff15fc114000            call dword ptr [004011fc]
'006cad68    50                      push eax
'006cad69    57                      push edi
'006cad6a    ff93a4000000            call dword ptr [ebx+000000a4]
'006cad70    dbe2                    fnclex
'006cad72    85c0                    test ax, ax
'006cad74    7d12                    jge 6cad88
'006cad76    68a4000000              push 000000a4
'006cad7b    68d00d4300              push 00430dd0
'006cad80    57                      push edi
'006cad81    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006cad82    ff1580104000            call dword ptr [00401080]
'006cad88    8d4de4                  lea ecx, dword ptr [ebp-1c]

' *** Reference to "__vbaFreeStr"
'006cad8b    ff150c134000            call dword ptr [0040130c]
'#Cleanup(var_40)
'006cad91    8d4dbc                  lea ecx, dword ptr [ebp-44]
'006cad94    51                      push ecx
'006cad95    8d55c4                  lea edx, dword ptr [ebp-3c]
'006cad98    52                      push edx
'006cad99    8d45cc                  lea eax, dword ptr [ebp-34]
'006cad9c    50                      push eax
'006cad9d    6a03                    push 03

' *** Reference to "__vbaFreeObjList"
'006cad9f    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_43, var_9, var_58)
'006cada5    8d8d74ffffff            lea ecx, dword ptr [ebp+ffffff74]
'006cadab    51                      push ecx
'006cadac    8d5584                  lea edx, dword ptr [ebp-7c]
'006cadaf    52                      push edx
'006cadb0    8d4594                  lea eax, dword ptr [ebp-6c]
'006cadb3    50                      push eax
'006cadb4    8d8d44ffffff            lea ecx, dword ptr [ebp+ffffff44]
'006cadba    51                      push ecx
'006cadbb    8d55a4                  lea edx, dword ptr [ebp-5c]
'006cadbe    52                      push edx
'006cadbf    6a05                    push 05

' *** Reference to "__vbaFreeVarList"
'006cadc1    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_83, var_125, var_121, var_119, var_91)
'006cadc7    8b45e8                  mov eax, dword ptr [ebp-18]
'006cadca    8b08                    mov ecx, dword ptr [eax]
'006cadcc    83c428                  add esp, 28
'006cadcf    8d55cc                  lea edx, dword ptr [ebp-34]
'006cadd2    52                      push edx
'006cadd3    50                      push eax
'006cadd4    ff91b4000000            call dword ptr [ecx+000000b4]
'006cadda    dbe2                    fnclex
'006caddc    85c0                    test ax, ax
'006cadde    7d15                    jge 6cadf5
'006cade0    8b4de8                  mov ecx, dword ptr [ebp-18]
'006cade3    68b4000000              push 000000b4
'006cade8    6830314300              push 00433130
'006caded    51                      push ecx
'006cadee    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006cadef    ff1580104000            call dword ptr [00401080]
'006cadf5    8b45cc                  mov eax, dword ptr [ebp-34]
'006cadf8    8d5dc8                  lea ebx, dword ptr [ebp-38]
'006cadfb    53                      push ebx
'006cadfc    83ec10                  sub esp, 10
'006cadff    ba08000000              mov edx, 00000008
'006cae04    8bdc                    mov ebx, esp
'006cae06    8913                    mov dword ptr [ebx], edx
'006cae08    899564ffffff            mov dword ptr [ebp+ffffff64], edx
'006cae0e    8b9568ffffff            mov edx, dword ptr [ebp+ffffff68]
'006cae14    b97c384500              mov ecx, 0045387c
'006cae19    895304                  mov dword ptr [ebx+04], edx
'006cae1c    898d6cffffff            mov dword ptr [ebp+ffffff6c], ecx
'006cae22    8b38                    mov edi, dword ptr [eax]
'006cae24    894b08                  mov dword ptr [ebx+08], ecx
'006cae27    8b8d70ffffff            mov ecx, dword ptr [ebp+ffffff70]
'006cae2d    50                      push eax
'006cae2e    898594feffff            mov dword ptr [ebp+fffffe94], eax
'006cae34    894b0c                  mov dword ptr [ebx+0c], ecx
'006cae37    ff5730                  call dword ptr [edi+30]
'006cae3a    dbe2                    fnclex
'006cae3c    85c0                    test ax, ax
'006cae3e    7d15                    jge 6cae55
'006cae40    8b9594feffff            mov edx, dword ptr [ebp+fffffe94]
'006cae46    6a30                    push 30
'006cae48    68d8304300              push 004330d8
'006cae4d    52                      push edx
'006cae4e    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006cae4f    ff1580104000            call dword ptr [00401080]
'006cae55    8b45c8                  mov eax, dword ptr [ebp-38]
'006cae58    8945ac                  mov dword ptr [ebp-54], eax
'006cae5b    8d45a4                  lea eax, dword ptr [ebp-5c]
'006cae5e    50                      push eax
'006cae5f    c745c800000000          mov dword ptr [ebp-38], 00000000
'006cae66    c745a409000000          mov dword ptr [ebp-5c], 00000009

' *** Reference to "rtcIsNull"
'006cae6d    ff1540114000            call dword ptr [00401140]
'006cae73    8b0e                    mov ecx, dword ptr [esi]
'006cae75    56                      push esi
'006cae76    8985a0feffff            mov dword ptr [ebp+fffffea0], eax
'006cae7c    ff9124030000            call dword ptr [ecx+00000324]
'006cae82    50                      push eax
'006cae83    8d55bc                  lea edx, dword ptr [ebp-44]
'006cae86    52                      push edx

' *** Reference to "__vbaObjSet"
'006cae87    ff15b4104000            call dword ptr [004010b4]
Set var_58 = IsNull(var_46)
'006cae8d    8d55c4                  lea edx, dword ptr [ebp-3c]
'006cae90    8bf8                    mov edi, eax
'006cae92    8b45e8                  mov eax, dword ptr [ebp-18]
'006cae95    8b08                    mov ecx, dword ptr [eax]
'006cae97    52                      push edx
'006cae98    50                      push eax
'006cae99    ff91b4000000            call dword ptr [ecx+000000b4]
'006cae9f    dbe2                    fnclex
'006caea1    85c0                    test ax, ax
'006caea3    7d15                    jge 6caeba
'006caea5    8b4de8                  mov ecx, dword ptr [ebp-18]
'006caea8    68b4000000              push 000000b4
'006caead    6830314300              push 00433130
'006caeb2    51                      push ecx
'006caeb3    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006caeb4    ff1580104000            call dword ptr [00401080]
'006caeba    8b45c4                  mov eax, dword ptr [ebp-3c]
'006caebd    8b10                    mov edx, dword ptr [eax]
'006caebf    8d5dc0                  lea ebx, dword ptr [ebp-40]
'006caec2    53                      push ebx
'006caec3    83ec10                  sub esp, 10
'006caec6    8bdc                    mov ebx, esp
'006caec8    b908000000              mov ecx, 00000008
'006caecd    890b                    mov dword ptr [ebx], ecx
'006caecf    8b8d58ffffff            mov ecx, dword ptr [ebp+ffffff58]
'006caed5    894b04                  mov dword ptr [ebx+04], ecx
'006caed8    c7855cffffff7c384500    mov dword ptr [ebp+ffffff5c], 0045387c
'006caee2    8b8d5cffffff            mov ecx, dword ptr [ebp+ffffff5c]
'006caee8    894b08                  mov dword ptr [ebx+08], ecx
'006caeeb    8b8d60ffffff            mov ecx, dword ptr [ebp+ffffff60]
'006caef1    50                      push eax
'006caef2    898588feffff            mov dword ptr [ebp+fffffe88], eax
'006caef8    894b0c                  mov dword ptr [ebx+0c], ecx
'006caefb    ff5230                  call dword ptr [edx+30]
'006caefe    dbe2                    fnclex
'006caf00    85c0                    test ax, ax
'006caf02    7d15                    jge 6caf19
'006caf04    8b9588feffff            mov edx, dword ptr [ebp+fffffe88]
'006caf0a    6a30                    push 30
'006caf0c    68d8304300              push 004330d8
'006caf11    52                      push edx
'006caf12    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006caf13    ff1580104000            call dword ptr [00401080]
'006caf19    8b45c0                  mov eax, dword ptr [ebp-40]
'006caf1c    8d9534ffffff            lea edx, dword ptr [ebp+ffffff34]
'006caf22    8d4d94                  lea ecx, dword ptr [ebp-6c]
'006caf25    c745c000000000          mov dword ptr [ebp-40], 00000000
'006caf2c    89458c                  mov dword ptr [ebp-74], eax
'006caf2f    c7458409000000          mov dword ptr [ebp-7c], 00000009
'006caf36    c7853cffffffcc134300    mov dword ptr [ebp+ffffff3c], 004313cc
'006caf40    c78534ffffff08000000    mov dword ptr [ebp+ffffff34], 00000008

' *** Reference to "__vbaVarDup"
'006caf4a    ff1598124000            call dword ptr [00401298]
var_121 = (vbNullChar)
'006caf50    668b85a0feffff          mov ax, word ptr [ebp+fffffea0]
'006caf57    8d4d84                  lea ecx, dword ptr [ebp-7c]
'006caf5a    51                      push ecx
'006caf5b    8d5594                  lea edx, dword ptr [ebp-6c]
'006caf5e    6689854cffffff          mov word ptr [ebp+ffffff4c], ax
'006caf65    52                      push edx
'006caf66    8d8544ffffff            lea eax, dword ptr [ebp+ffffff44]
'006caf6c    50                      push eax
'006caf6d    8d8d74ffffff            lea ecx, dword ptr [ebp+ffffff74]
'006caf73    51                      push ecx
'006caf74    c78544ffffff0b000000    mov dword ptr [ebp+ffffff44], 0000000b

' *** Reference to "rtcImmediateIf"
'006caf7e    ff154c124000            call dword ptr [0040124c]
'006caf84    8b1f                    mov ebx, dword ptr [edi]
'006caf86    8d9574ffffff            lea edx, dword ptr [ebp+ffffff74]
'006caf8c    52                      push edx
'006caf8d    8d45e4                  lea eax, dword ptr [ebp-1c]
'006caf90    50                      push eax

' *** Reference to "__vbaStrVarVal"
'006caf91    ff15fc114000            call dword ptr [004011fc]
'006caf97    50                      push eax
'006caf98    57                      push edi
'006caf99    ff93a4000000            call dword ptr [ebx+000000a4]
'006caf9f    dbe2                    fnclex
'006cafa1    85c0                    test ax, ax
'006cafa3    7d12                    jge 6cafb7
'006cafa5    68a4000000              push 000000a4
'006cafaa    68d00d4300              push 00430dd0
'006cafaf    57                      push edi
'006cafb0    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006cafb1    ff1580104000            call dword ptr [00401080]
'006cafb7    8d4de4                  lea ecx, dword ptr [ebp-1c]

' *** Reference to "__vbaFreeStr"
'006cafba    ff150c134000            call dword ptr [0040130c]
'#Cleanup(var_40)
'006cafc0    8d4dbc                  lea ecx, dword ptr [ebp-44]
'006cafc3    51                      push ecx
'006cafc4    8d55c4                  lea edx, dword ptr [ebp-3c]
'006cafc7    52                      push edx
'006cafc8    8d45cc                  lea eax, dword ptr [ebp-34]
'006cafcb    50                      push eax
'006cafcc    6a03                    push 03

' *** Reference to "__vbaFreeObjList"
'006cafce    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_43, var_9, var_58)
'006cafd4    8d8d74ffffff            lea ecx, dword ptr [ebp+ffffff74]
'006cafda    51                      push ecx
'006cafdb    8d5584                  lea edx, dword ptr [ebp-7c]
'006cafde    52                      push edx
'006cafdf    8d4594                  lea eax, dword ptr [ebp-6c]
'006cafe2    50                      push eax
'006cafe3    8d8d44ffffff            lea ecx, dword ptr [ebp+ffffff44]
'006cafe9    51                      push ecx
'006cafea    8d55a4                  lea edx, dword ptr [ebp-5c]
'006cafed    52                      push edx
'006cafee    6a05                    push 05

' *** Reference to "__vbaFreeVarList"
'006caff0    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_83, var_125, var_121, var_119, var_91)
'006caff6    83c428                  add esp, 28
'006caff9    8d8d64ffffff            lea ecx, dword ptr [ebp+ffffff64]
'006cafff    51                      push ecx
'006cb000    8d55a4                  lea edx, dword ptr [ebp-5c]
'006cb003    8d4640                  lea eax, dword ptr [esi+40]
'006cb006    52                      push edx
'006cb007    89856cffffff            mov dword ptr [ebp+ffffff6c], eax
'006cb00d    c78564ffffff08400000    mov dword ptr [ebp+ffffff64], 00004008

' *** Reference to "rtcUpperCaseVar"
'006cb017    ff152c114000            call dword ptr [0040112c]
'006cb01d    8d55a4                  lea edx, dword ptr [ebp-5c]
'006cb020    8d8d64feffff            lea ecx, dword ptr [ebp+fffffe64]

' *** Reference to "__vbaVarMove"
'006cb026    ff151c104000            call dword ptr [0040101c]
var_122 = (UCase(arg_7))

' *** Reference to "__vbaVarTstEq"
'006cb02c    8b3d3c114000            mov edi, dword ptr [0040113c]
'006cb032    8d8564feffff            lea eax, dword ptr [ebp+fffffe64]
'006cb038    50                      push eax
'006cb039    8d8d64ffffff            lea ecx, dword ptr [ebp+ffffff64]
'006cb03f    bb08800000              mov ebx, 00008008
'006cb044    51                      push ecx
'006cb045    c7856cffffff08394500    mov dword ptr [ebp+ffffff6c], 00453908
'006cb04f    899d64ffffff            mov dword ptr [ebp+ffffff64], ebx
'006cb055    ffd7                    call edi
'006cb057    6685c0                  test eax, eax
'006cb05a    0f8421070000            je 6cb781

If (((var_122) = ("VSSORTCLASSE"))) Then
'006cb060    66837e3c00              cmp word ptr [esi+3c], 00
'006cb065    0f85cc010000            jne 6cb237
    
    If (    arg_7 = 0) Then
'006cb06b    8b16                    mov edx, dword ptr [esi]
'006cb06d    56                      push esi
'006cb06e    ff921c030000            call dword ptr [edx+0000031c]

' *** Reference to "__vbaObjSet"
'006cb074    8b3db4104000            mov edi, dword ptr [004010b4]
'006cb07a    50                      push eax
'006cb07b    8d45c0                  lea eax, dword ptr [ebp-40]
'006cb07e    50                      push eax
'006cb07f    ffd7                    call edi
    Set var_5 = ((var_122) = ("VSSORTCLASSE"))
'006cb081    8b0e                    mov ecx, dword ptr [esi]
'006cb083    56                      push esi
'006cb084    8bd8                    mov ebx, eax
'006cb086    c7856cffffff00000000    mov dword ptr [ebp+ffffff6c], 00000000
'006cb090    c78564ffffff03000000    mov dword ptr [ebp+ffffff64], 00000003
'006cb09a    ff91fc020000            call dword ptr [ecx+000002fc]
'006cb0a0    50                      push eax
'006cb0a1    8d55cc                  lea edx, dword ptr [ebp-34]
'006cb0a4    52                      push edx
'006cb0a5    ffd7                    call edi
    Set var_43 = var_5
'006cb0a7    8d8da0feffff            lea ecx, dword ptr [ebp+fffffea0]
'006cb0ad    8bf8                    mov edi, eax
'006cb0af    8b07                    mov eax, dword ptr [edi]
'006cb0b1    51                      push ecx
'006cb0b2    57                      push edi
'006cb0b3    ff90f0000000            call dword ptr [eax+000000f0]
'006cb0b9    dbe2                    fnclex
'006cb0bb    85c0                    test ax, ax
'006cb0bd    7d12                    jge 6cb0d1
'006cb0bf    68f0000000              push 000000f0
'006cb0c4    681cb94300              push 0043b91c
'006cb0c9    57                      push edi
'006cb0ca    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006cb0cb    ff1580104000            call dword ptr [00401080]
'006cb0d1    668b4644                mov ax, word ptr [esi+44]
'006cb0d5    668b95a0feffff          mov dx, word ptr [ebp+fffffea0]
'006cb0dc    6689852cffffff          mov word ptr [ebp+ffffff2c], ax
'006cb0e3    a174b17200              mov ax, word ptr [0072b174]
'006cb0e8    85c0                    test ax, ax
'006cb0ea    6689954cffffff          mov word ptr [ebp+ffffff4c], dx
'006cb0f1    c78544ffffff02000000    mov dword ptr [ebp+ffffff44], 00000002
'006cb0fb    7515                    jne 6cb112
'006cb0fd    6874b17200              push 0072b174
'006cb102    6890c04100              push 0041c090

' *** Reference to "__vbaNew2"
'006cb107    ff152c124000            call dword ptr [0040122c]
    Dim var_77 As New frmFichePerso
'006cb10d    a174b17200              mov ax, word ptr [0072b174]
'006cb112    8b08                    mov ecx, dword ptr [eax]
'006cb114    50                      push eax
'006cb115    ff9120060000            call dword ptr [ecx+00000620]
'006cb11b    50                      push eax
'006cb11c    8d55c8                  lea edx, dword ptr [ebp-38]
'006cb11f    52                      push edx

' *** Reference to "__vbaObjSet"
'006cb120    ff15b4104000            call dword ptr [004010b4]
    Set var_46 = var_77
'006cb126    33d2                    xor edx, edx
'006cb128    668b563c                mov dx, word ptr [esi+3c]
'006cb12c    8d4dc4                  lea ecx, dword ptr [ebp-3c]
'006cb12f    51                      push ecx
'006cb130    8bf8                    mov edi, eax
'006cb132    8b07                    mov eax, dword ptr [edi]
'006cb134    52                      push edx
'006cb135    57                      push edi
'006cb136    ff5040                  call dword ptr [eax+40]
'006cb139    dbe2                    fnclex
'006cb13b    85c0                    test ax, ax
'006cb13d    7d0f                    jge 6cb14e
'006cb13f    6a40                    push 40
'006cb141    686c384300              push 0043386c
'006cb146    57                      push edi
'006cb147    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006cb148    ff1580104000            call dword ptr [00401080]
'006cb14e    8b8d64ffffff            mov ecx, dword ptr [ebp+ffffff64]
'006cb154    8b9568ffffff            mov edx, dword ptr [ebp+ffffff68]
'006cb15a    83ec10                  sub esp, 10
'006cb15d    8bc4                    mov eax, esp
'006cb15f    8908                    mov dword ptr [eax], ecx
'006cb161    8b8d6cffffff            mov ecx, dword ptr [ebp+ffffff6c]
'006cb167    895004                  mov dword ptr [eax+04], edx
'006cb16a    8b9570ffffff            mov edx, dword ptr [ebp+ffffff70]
'006cb170    894808                  mov dword ptr [eax+08], ecx
'006cb173    8b8d44ffffff            mov ecx, dword ptr [ebp+ffffff44]
'006cb179    89500c                  mov dword ptr [eax+0c], edx
'006cb17c    8b9548ffffff            mov edx, dword ptr [ebp+ffffff48]
'006cb182    8b3b                    mov edi, dword ptr [ebx]
'006cb184    83ec10                  sub esp, 10
'006cb187    8bc4                    mov eax, esp
'006cb189    8908                    mov dword ptr [eax], ecx
'006cb18b    8b8d4cffffff            mov ecx, dword ptr [ebp+ffffff4c]
'006cb191    895004                  mov dword ptr [eax+04], edx
'006cb194    8b9550ffffff            mov edx, dword ptr [ebp+ffffff50]
'006cb19a    894808                  mov dword ptr [eax+08], ecx
'006cb19d    89500c                  mov dword ptr [eax+0c], edx
'006cb1a0    8b9528ffffff            mov edx, dword ptr [ebp+ffffff28]
'006cb1a6    83ec10                  sub esp, 10
'006cb1a9    8bcc                    mov ecx, esp
'006cb1ab    b802000000              mov eax, 00000002
'006cb1b0    8901                    mov dword ptr [ecx], eax
'006cb1b2    8b852cffffff            mov eax, dword ptr [ebp+ffffff2c]
'006cb1b8    895104                  mov dword ptr [ecx+04], edx
'006cb1bb    8b9530ffffff            mov edx, dword ptr [ebp+ffffff30]
'006cb1c1    894108                  mov dword ptr [ecx+08], eax
'006cb1c4    8b45c4                  mov eax, dword ptr [ebp-3c]
'006cb1c7    6a03                    push 03
'006cb1c9    689e000000              push 0000009e
'006cb1ce    89510c                  mov dword ptr [ecx+0c], edx
'006cb1d1    50                      push eax
'006cb1d2    8d4da4                  lea ecx, dword ptr [ebp-5c]
'006cb1d5    51                      push ecx

' *** Reference to "__vbaLateIdCallLd"
'006cb1d6    ff157c114000            call dword ptr [0040117c]
    var_83 = var_9.UNK_-256 - 12_158
'006cb1dc    83c440                  add esp, 40
'006cb1df    50                      push eax
'006cb1e0    8d55e4                  lea edx, dword ptr [ebp-1c]
'006cb1e3    52                      push edx

' *** Reference to "__vbaStrVarVal"
'006cb1e4    ff15fc114000            call dword ptr [004011fc]
'006cb1ea    50                      push eax
'006cb1eb    53                      push ebx
'006cb1ec    ff97a4000000            call dword ptr [edi+000000a4]
'006cb1f2    dbe2                    fnclex
'006cb1f4    85c0                    test ax, ax
'006cb1f6    7d12                    jge 6cb20a
'006cb1f8    68a4000000              push 000000a4
'006cb1fd    68d00d4300              push 00430dd0
'006cb202    53                      push ebx
'006cb203    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006cb204    ff1580104000            call dword ptr [00401080]
'006cb20a    8d4de4                  lea ecx, dword ptr [ebp-1c]

' *** Reference to "__vbaFreeStr"
'006cb20d    ff150c134000            call dword ptr [0040130c]
    '#Cleanup(var_40)
'006cb213    8d45c0                  lea eax, dword ptr [ebp-40]
'006cb216    50                      push eax
'006cb217    8d4dc4                  lea ecx, dword ptr [ebp-3c]
'006cb21a    51                      push ecx
'006cb21b    8d55c8                  lea edx, dword ptr [ebp-38]
'006cb21e    52                      push edx
'006cb21f    8d45cc                  lea eax, dword ptr [ebp-34]
'006cb222    50                      push eax
'006cb223    6a04                    push 04

' *** Reference to "__vbaFreeObjList"
'006cb225    ff154c104000            call dword ptr [0040104c]
    '#Cleanup( var_43, var_46, var_9, var_5)
'006cb22b    83c414                  add esp, 14
'006cb22e    8d4da4                  lea ecx, dword ptr [ebp-5c]

' *** Reference to "__vbaFreeVar"
'006cb231    ff1528104000            call dword ptr [00401028]
    '#Cleanup(var_83)
End If
'006cb237    8b0e                    mov ecx, dword ptr [esi]

' *** Reference to "__vbaObjSet"
'006cb239    8b1db4104000            mov ebx, dword ptr [004010b4]
'006cb23f    56                      push esi
'006cb240    ff9108030000            call dword ptr [ecx+00000308]
'006cb246    50                      push eax
'006cb247    8d55b4                  lea edx, dword ptr [ebp-4c]
'006cb24a    52                      push edx
'006cb24b    ffd3                    call ebx
Set var_62 = ((var_122) = ("VSSORTCLASSE"))
'006cb24d    898578feffff            mov dword ptr [ebp+fffffe78], eax
'006cb253    8b06                    mov eax, dword ptr [esi]
'006cb255    56                      push esi
'006cb256    c7856cffffff00000000    mov dword ptr [ebp+ffffff6c], 00000000
'006cb260    c78564ffffff03000000    mov dword ptr [ebp+ffffff64], 00000003
'006cb26a    ff90fc020000            call dword ptr [eax+000002fc]
'006cb270    50                      push eax
'006cb271    8d4dcc                  lea ecx, dword ptr [ebp-34]
'006cb274    51                      push ecx
'006cb275    ffd3                    call ebx
Set var_43 = Nothing
'006cb277    8bf8                    mov edi, eax
'006cb279    8b17                    mov edx, dword ptr [edi]
'006cb27b    8d85a0feffff            lea eax, dword ptr [ebp+fffffea0]
'006cb281    50                      push eax
'006cb282    57                      push edi
'006cb283    ff92f0000000            call dword ptr [edx+000000f0]
'006cb289    dbe2                    fnclex
'006cb28b    85c0                    test ax, ax
'006cb28d    7d12                    jge 6cb2a1
'006cb28f    68f0000000              push 000000f0
'006cb294    681cb94300              push 0043b91c
'006cb299    57                      push edi
'006cb29a    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006cb29b    ff1580104000            call dword ptr [00401080]
'006cb2a1    668b8da0feffff          mov cx, word ptr [ebp+fffffea0]
'006cb2a8    668b5634                mov dx, word ptr [esi+34]
'006cb2ac    b802000000              mov eax, 00000002
'006cb2b1    898544ffffff            mov dword ptr [ebp+ffffff44], eax
'006cb2b7    898524ffffff            mov dword ptr [ebp+ffffff24], eax
'006cb2bd    a174b17200              mov ax, word ptr [0072b174]
'006cb2c2    85c0                    test ax, ax
'006cb2c4    66898d4cffffff          mov word ptr [ebp+ffffff4c], cx
'006cb2cb    6689952cffffff          mov word ptr [ebp+ffffff2c], dx
'006cb2d2    7515                    jne 6cb2e9
'006cb2d4    6874b17200              push 0072b174
'006cb2d9    6890c04100              push 0041c090

' *** Reference to "__vbaNew2"
'006cb2de    ff152c124000            call dword ptr [0040122c]
Set var_77 = New frmFichePerso
'006cb2e4    a174b17200              mov ax, word ptr [0072b174]
'006cb2e9    8b08                    mov ecx, dword ptr [eax]
'006cb2eb    50                      push eax
'006cb2ec    ff9120060000            call dword ptr [ecx+00000620]
'006cb2f2    50                      push eax
'006cb2f3    8d55c8                  lea edx, dword ptr [ebp-38]
'006cb2f6    52                      push edx
'006cb2f7    ffd3                    call ebx
Set var_46 = var_77
'006cb2f9    33d2                    xor edx, edx
'006cb2fb    668b563c                mov dx, word ptr [esi+3c]
'006cb2ff    8d4dc4                  lea ecx, dword ptr [ebp-3c]
'006cb302    51                      push ecx
'006cb303    8bf8                    mov edi, eax
'006cb305    8b07                    mov eax, dword ptr [edi]
'006cb307    52                      push edx
'006cb308    57                      push edi
'006cb309    ff5040                  call dword ptr [eax+40]
'006cb30c    dbe2                    fnclex
'006cb30e    85c0                    test ax, ax
'006cb310    7d0f                    jge 6cb321
'006cb312    6a40                    push 40
'006cb314    686c384300              push 0043386c
'006cb319    57                      push edi
'006cb31a    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006cb31b    ff1580104000            call dword ptr [00401080]
'006cb321    8b06                    mov eax, dword ptr [esi]
'006cb323    56                      push esi
'006cb324    c7850cffffffc48d4300    mov dword ptr [ebp+ffffff0c], 00438dc4
'006cb32e    c78504ffffff08000000    mov dword ptr [ebp+ffffff04], 00000008
'006cb338    ff90fc020000            call dword ptr [eax+000002fc]
'006cb33e    50                      push eax
'006cb33f    8d4dc0                  lea ecx, dword ptr [ebp-40]
'006cb342    51                      push ecx
'006cb343    ffd3                    call ebx
Set var_5 = Nothing
'006cb345    8bf8                    mov edi, eax
'006cb347    8b17                    mov edx, dword ptr [edi]
'006cb349    8d859cfeffff            lea eax, dword ptr [ebp+fffffe9c]
'006cb34f    50                      push eax
'006cb350    57                      push edi
'006cb351    ff92f0000000            call dword ptr [edx+000000f0]
'006cb357    dbe2                    fnclex
'006cb359    85c0                    test ax, ax
'006cb35b    7d12                    jge 6cb36f
'006cb35d    68f0000000              push 000000f0
'006cb362    681cb94300              push 0043b91c
'006cb367    57                      push edi
'006cb368    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006cb369    ff1580104000            call dword ptr [00401080]
'006cb36f    a174b17200              mov ax, word ptr [0072b174]
'006cb374    85c0                    test ax, ax
'006cb376    668b8d9cfeffff          mov cx, word ptr [ebp+fffffe9c]
'006cb37d    668b5636                mov dx, word ptr [esi+36]
'006cb381    66898ddcfeffff          mov word ptr [ebp+fffffedc], cx
'006cb388    668995bcfeffff          mov word ptr [ebp+fffffebc], dx
'006cb38f    7515                    jne 6cb3a6
'006cb391    6874b17200              push 0072b174
'006cb396    6890c04100              push 0041c090

' *** Reference to "__vbaNew2"
'006cb39b    ff152c124000            call dword ptr [0040122c]
Set var_77 = New frmFichePerso
'006cb3a1    a174b17200              mov ax, word ptr [0072b174]
'006cb3a6    8b08                    mov ecx, dword ptr [eax]
'006cb3a8    50                      push eax
'006cb3a9    ff9120060000            call dword ptr [ecx+00000620]
'006cb3af    50                      push eax
'006cb3b0    8d55bc                  lea edx, dword ptr [ebp-44]
'006cb3b3    52                      push edx
'006cb3b4    ffd3                    call ebx
Set var_58 = var_77
'006cb3b6    33d2                    xor edx, edx
'006cb3b8    668b563c                mov dx, word ptr [esi+3c]
'006cb3bc    8d4db8                  lea ecx, dword ptr [ebp-48]
'006cb3bf    51                      push ecx
'006cb3c0    8bf8                    mov edi, eax
'006cb3c2    8b07                    mov eax, dword ptr [edi]
'006cb3c4    52                      push edx
'006cb3c5    57                      push edi
'006cb3c6    ff5040                  call dword ptr [eax+40]
'006cb3c9    dbe2                    fnclex
'006cb3cb    85c0                    test ax, ax
'006cb3cd    7d0f                    jge 6cb3de
'006cb3cf    6a40                    push 40
'006cb3d1    686c384300              push 0043386c
'006cb3d6    57                      push edi
'006cb3d7    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006cb3d8    ff1580104000            call dword ptr [00401080]
'006cb3de    8b8578feffff            mov eax, dword ptr [ebp+fffffe78]
'006cb3e4    8b18                    mov ebx, dword ptr [eax]
'006cb3e6    8b9564ffffff            mov edx, dword ptr [ebp+ffffff64]
'006cb3ec    8b8568ffffff            mov eax, dword ptr [ebp+ffffff68]
'006cb3f2    83ec10                  sub esp, 10
'006cb3f5    8bcc                    mov ecx, esp
'006cb3f7    8911                    mov dword ptr [ecx], edx
'006cb3f9    8b956cffffff            mov edx, dword ptr [ebp+ffffff6c]
'006cb3ff    894104                  mov dword ptr [ecx+04], eax
'006cb402    8b8570ffffff            mov eax, dword ptr [ebp+ffffff70]
'006cb408    895108                  mov dword ptr [ecx+08], edx
'006cb40b    8b9544ffffff            mov edx, dword ptr [ebp+ffffff44]
'006cb411    89410c                  mov dword ptr [ecx+0c], eax
'006cb414    8b8548ffffff            mov eax, dword ptr [ebp+ffffff48]
'006cb41a    83ec10                  sub esp, 10
'006cb41d    8bcc                    mov ecx, esp
'006cb41f    8911                    mov dword ptr [ecx], edx
'006cb421    8b954cffffff            mov edx, dword ptr [ebp+ffffff4c]
'006cb427    894104                  mov dword ptr [ecx+04], eax
'006cb42a    8b8550ffffff            mov eax, dword ptr [ebp+ffffff50]
'006cb430    895108                  mov dword ptr [ecx+08], edx
'006cb433    8b9524ffffff            mov edx, dword ptr [ebp+ffffff24]
'006cb439    89410c                  mov dword ptr [ecx+0c], eax
'006cb43c    8b8528ffffff            mov eax, dword ptr [ebp+ffffff28]
'006cb442    83ec10                  sub esp, 10
'006cb445    8bcc                    mov ecx, esp
'006cb447    8911                    mov dword ptr [ecx], edx
'006cb449    8b952cffffff            mov edx, dword ptr [ebp+ffffff2c]
'006cb44f    894104                  mov dword ptr [ecx+04], eax
'006cb452    8b8530ffffff            mov eax, dword ptr [ebp+ffffff30]
'006cb458    895108                  mov dword ptr [ecx+08], edx
'006cb45b    6a03                    push 03
'006cb45d    89410c                  mov dword ptr [ecx+0c], eax
'006cb460    8b4dc4                  mov ecx, dword ptr [ebp-3c]
'006cb463    689e000000              push 0000009e
'006cb468    51                      push ecx
'006cb469    8d55a4                  lea edx, dword ptr [ebp-5c]
'006cb46c    52                      push edx

' *** Reference to "__vbaLateIdCallLd"
'006cb46d    ff157c114000            call dword ptr [0040117c]
var_83 = var_9.UNK_-256 - 12_158

' *** Reference to "__vbaVarCat"
'006cb473    8b3d08124000            mov edi, dword ptr [00401208]
'006cb479    83c440                  add esp, 40
'006cb47c    50                      push eax
'006cb47d    8d8504ffffff            lea eax, dword ptr [ebp+ffffff04]
'006cb483    50                      push eax
'006cb484    8d4d94                  lea ecx, dword ptr [ebp-6c]
'006cb487    51                      push ecx
'006cb488    ffd7                    call edi
'006cb48a    8b8d00ffffff            mov ecx, dword ptr [ebp+ffffff00]
'006cb490    50                      push eax
'006cb491    83ec10                  sub esp, 10
'006cb494    8bd4                    mov edx, esp
'006cb496    b803000000              mov eax, 00000003
'006cb49b    8902                    mov dword ptr [edx], eax
'006cb49d    8b85f8feffff            mov eax, dword ptr [ebp+fffffef8]
'006cb4a3    894204                  mov dword ptr [edx+04], eax
'006cb4a6    33c0                    xor eax, eax
'006cb4a8    894208                  mov dword ptr [edx+08], eax
'006cb4ab    894a0c                  mov dword ptr [edx+0c], ecx
'006cb4ae    8b8ddcfeffff            mov ecx, dword ptr [ebp+fffffedc]
'006cb4b4    83ec10                  sub esp, 10
'006cb4b7    8bd4                    mov edx, esp
'006cb4b9    b802000000              mov eax, 00000002
'006cb4be    8902                    mov dword ptr [edx], eax
'006cb4c0    8b85d8feffff            mov eax, dword ptr [ebp+fffffed8]
'006cb4c6    894204                  mov dword ptr [edx+04], eax
'006cb4c9    8b85e0feffff            mov eax, dword ptr [ebp+fffffee0]
'006cb4cf    894a08                  mov dword ptr [edx+08], ecx
'006cb4d2    89420c                  mov dword ptr [edx+0c], eax
'006cb4d5    8b95b8feffff            mov edx, dword ptr [ebp+fffffeb8]
'006cb4db    83ec10                  sub esp, 10
'006cb4de    8bcc                    mov ecx, esp
'006cb4e0    b802000000              mov eax, 00000002
'006cb4e5    8901                    mov dword ptr [ecx], eax
'006cb4e7    8b85bcfeffff            mov eax, dword ptr [ebp+fffffebc]
'006cb4ed    895104                  mov dword ptr [ecx+04], edx
'006cb4f0    8b95c0feffff            mov edx, dword ptr [ebp+fffffec0]
'006cb4f6    894108                  mov dword ptr [ecx+08], eax
'006cb4f9    8b45b8                  mov eax, dword ptr [ebp-48]
'006cb4fc    6a03                    push 03
'006cb4fe    89510c                  mov dword ptr [ecx+0c], edx
'006cb501    689e000000              push 0000009e
'006cb506    50                      push eax
'006cb507    8d4d84                  lea ecx, dword ptr [ebp-7c]
'006cb50a    51                      push ecx

' *** Reference to "__vbaLateIdCallLd"
'006cb50b    ff157c114000            call dword ptr [0040117c]
var_119 = 0.UNK_-256 - 12_158
'006cb511    83c440                  add esp, 40
'006cb514    50                      push eax
'006cb515    8d9574ffffff            lea edx, dword ptr [ebp+ffffff74]
'006cb51b    52                      push edx
'006cb51c    ffd7                    call edi
'006cb51e    50                      push eax
'006cb51f    8d45e4                  lea eax, dword ptr [ebp-1c]
'006cb522    50                      push eax

' *** Reference to "__vbaStrVarVal"
'006cb523    ff15fc114000            call dword ptr [004011fc]
'006cb529    8bbd78feffff            mov edi, dword ptr [ebp+fffffe78]
'006cb52f    50                      push eax
'006cb530    57                      push edi
'006cb531    ff93a4000000            call dword ptr [ebx+000000a4]
'006cb537    dbe2                    fnclex
'006cb539    85c0                    test ax, ax
'006cb53b    7d12                    jge 6cb54f
'006cb53d    68a4000000              push 000000a4
'006cb542    68d00d4300              push 00430dd0
'006cb547    57                      push edi
'006cb548    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006cb549    ff1580104000            call dword ptr [00401080]
'006cb54f    8d4de4                  lea ecx, dword ptr [ebp-1c]

' *** Reference to "__vbaFreeStr"
'006cb552    ff150c134000            call dword ptr [0040130c]
'#Cleanup(var_40)
'006cb558    8d4db4                  lea ecx, dword ptr [ebp-4c]
'006cb55b    51                      push ecx
'006cb55c    8d55b8                  lea edx, dword ptr [ebp-48]
'006cb55f    52                      push edx
'006cb560    8d45bc                  lea eax, dword ptr [ebp-44]
'006cb563    50                      push eax
'006cb564    8d4dc0                  lea ecx, dword ptr [ebp-40]
'006cb567    51                      push ecx
'006cb568    8d55c4                  lea edx, dword ptr [ebp-3c]
'006cb56b    52                      push edx
'006cb56c    8d45c8                  lea eax, dword ptr [ebp-38]
'006cb56f    50                      push eax
'006cb570    8d4dcc                  lea ecx, dword ptr [ebp-34]
'006cb573    51                      push ecx
'006cb574    6a07                    push 07

' *** Reference to "__vbaFreeObjList"
'006cb576    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_43, var_46, var_9, var_5, var_58, var_61, var_62)
'006cb57c    8d9574ffffff            lea edx, dword ptr [ebp+ffffff74]
'006cb582    52                      push edx
'006cb583    8d4584                  lea eax, dword ptr [ebp-7c]
'006cb586    50                      push eax
'006cb587    8d4d94                  lea ecx, dword ptr [ebp-6c]
'006cb58a    51                      push ecx
'006cb58b    8d55a4                  lea edx, dword ptr [ebp-5c]
'006cb58e    52                      push edx
'006cb58f    6a04                    push 04

' *** Reference to "__vbaFreeVarList"
'006cb591    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_83, var_121, var_119, var_91)
'006cb597    8b06                    mov eax, dword ptr [esi]
'006cb599    83c434                  add esp, 34
'006cb59c    56                      push esi
'006cb59d    ff9018030000            call dword ptr [eax+00000318]

' *** Reference to "__vbaObjSet"
'006cb5a3    8b1db4104000            mov ebx, dword ptr [004010b4]
'006cb5a9    50                      push eax
'006cb5aa    8d4dc0                  lea ecx, dword ptr [ebp-40]
'006cb5ad    51                      push ecx
'006cb5ae    ffd3                    call ebx
Set var_5 = Nothing
'006cb5b0    8b16                    mov edx, dword ptr [esi]
'006cb5b2    56                      push esi
'006cb5b3    898588feffff            mov dword ptr [ebp+fffffe88], eax
'006cb5b9    c7856cffffff00000000    mov dword ptr [ebp+ffffff6c], 00000000
'006cb5c3    c78564ffffff03000000    mov dword ptr [ebp+ffffff64], 00000003
'006cb5cd    ff92fc020000            call dword ptr [edx+000002fc]
'006cb5d3    50                      push eax
'006cb5d4    8d45cc                  lea eax, dword ptr [ebp-34]
'006cb5d7    50                      push eax
'006cb5d8    ffd3                    call ebx
Set var_43 = Nothing
'006cb5da    8d95a0feffff            lea edx, dword ptr [ebp+fffffea0]
'006cb5e0    8bf8                    mov edi, eax
'006cb5e2    8b0f                    mov ecx, dword ptr [edi]
'006cb5e4    52                      push edx
'006cb5e5    57                      push edi
'006cb5e6    ff91f0000000            call dword ptr [ecx+000000f0]
'006cb5ec    dbe2                    fnclex
'006cb5ee    85c0                    test ax, ax
'006cb5f0    7d12                    jge 6cb604
'006cb5f2    68f0000000              push 000000f0
'006cb5f7    681cb94300              push 0043b91c
'006cb5fc    57                      push edi
'006cb5fd    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006cb5fe    ff1580104000            call dword ptr [00401080]
'006cb604    668b85a0feffff          mov ax, word ptr [ebp+fffffea0]
'006cb60b    668b4e38                mov cx, word ptr [esi+38]
'006cb60f    6689854cffffff          mov word ptr [ebp+ffffff4c], ax
'006cb616    b802000000              mov eax, 00000002
'006cb61b    898544ffffff            mov dword ptr [ebp+ffffff44], eax
'006cb621    898524ffffff            mov dword ptr [ebp+ffffff24], eax
'006cb627    a174b17200              mov ax, word ptr [0072b174]
'006cb62c    85c0                    test ax, ax
'006cb62e    66898d2cffffff          mov word ptr [ebp+ffffff2c], cx
'006cb635    7515                    jne 6cb64c
'006cb637    6874b17200              push 0072b174
'006cb63c    6890c04100              push 0041c090

' *** Reference to "__vbaNew2"
'006cb641    ff152c124000            call dword ptr [0040122c]
Set var_77 = New frmFichePerso
'006cb647    a174b17200              mov ax, word ptr [0072b174]
'006cb64c    8b10                    mov edx, dword ptr [eax]
'006cb64e    50                      push eax
'006cb64f    ff9220060000            call dword ptr [edx+00000620]
'006cb655    50                      push eax
'006cb656    8d45c8                  lea eax, dword ptr [ebp-38]
'006cb659    50                      push eax
'006cb65a    ffd3                    call ebx
Set var_46 = var_77
'006cb65c    8bf8                    mov edi, eax
'006cb65e    8b0f                    mov ecx, dword ptr [edi]
'006cb660    33c0                    xor eax, eax
'006cb662    668b463c                mov ax, word ptr [esi+3c]
'006cb666    8d55c4                  lea edx, dword ptr [ebp-3c]
'006cb669    52                      push edx
'006cb66a    50                      push eax
'006cb66b    57                      push edi
'006cb66c    ff5140                  call dword ptr [ecx+40]
'006cb66f    dbe2                    fnclex
'006cb671    85c0                    test ax, ax
'006cb673    7d0f                    jge 6cb684
'006cb675    6a40                    push 40
'006cb677    686c384300              push 0043386c
'006cb67c    57                      push edi
'006cb67d    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006cb67e    ff1580104000            call dword ptr [00401080]
'006cb684    8b8d88feffff            mov ecx, dword ptr [ebp+fffffe88]
'006cb68a    8b39                    mov edi, dword ptr [ecx]
'006cb68c    8b8564ffffff            mov eax, dword ptr [ebp+ffffff64]
'006cb692    8b8d68ffffff            mov ecx, dword ptr [ebp+ffffff68]
'006cb698    83ec10                  sub esp, 10
'006cb69b    8bd4                    mov edx, esp
'006cb69d    8902                    mov dword ptr [edx], eax
'006cb69f    8b856cffffff            mov eax, dword ptr [ebp+ffffff6c]
'006cb6a5    894a04                  mov dword ptr [edx+04], ecx
'006cb6a8    8b8d70ffffff            mov ecx, dword ptr [ebp+ffffff70]
'006cb6ae    894208                  mov dword ptr [edx+08], eax
'006cb6b1    8b8544ffffff            mov eax, dword ptr [ebp+ffffff44]
'006cb6b7    894a0c                  mov dword ptr [edx+0c], ecx
'006cb6ba    8b8d48ffffff            mov ecx, dword ptr [ebp+ffffff48]
'006cb6c0    83ec10                  sub esp, 10
'006cb6c3    8bd4                    mov edx, esp
'006cb6c5    8902                    mov dword ptr [edx], eax
'006cb6c7    8b854cffffff            mov eax, dword ptr [ebp+ffffff4c]
'006cb6cd    894a04                  mov dword ptr [edx+04], ecx
'006cb6d0    8b8d50ffffff            mov ecx, dword ptr [ebp+ffffff50]
'006cb6d6    894208                  mov dword ptr [edx+08], eax
'006cb6d9    8b8524ffffff            mov eax, dword ptr [ebp+ffffff24]
'006cb6df    894a0c                  mov dword ptr [edx+0c], ecx
'006cb6e2    8b8d28ffffff            mov ecx, dword ptr [ebp+ffffff28]
'006cb6e8    83ec10                  sub esp, 10
'006cb6eb    8bd4                    mov edx, esp
'006cb6ed    8902                    mov dword ptr [edx], eax
'006cb6ef    8b852cffffff            mov eax, dword ptr [ebp+ffffff2c]
'006cb6f5    894a04                  mov dword ptr [edx+04], ecx
'006cb6f8    8b8d30ffffff            mov ecx, dword ptr [ebp+ffffff30]
'006cb6fe    894208                  mov dword ptr [edx+08], eax
'006cb701    6a03                    push 03
'006cb703    894a0c                  mov dword ptr [edx+0c], ecx
'006cb706    8b55c4                  mov edx, dword ptr [ebp-3c]
'006cb709    689e000000              push 0000009e
'006cb70e    52                      push edx
'006cb70f    8d45a4                  lea eax, dword ptr [ebp-5c]
'006cb712    50                      push eax

' *** Reference to "__vbaLateIdCallLd"
'006cb713    ff157c114000            call dword ptr [0040117c]
var_83 = var_9.UNK_-256 - 12_158
'006cb719    83c440                  add esp, 40
'006cb71c    50                      push eax
'006cb71d    8d4de4                  lea ecx, dword ptr [ebp-1c]
'006cb720    51                      push ecx

' *** Reference to "__vbaStrVarVal"
'006cb721    ff15fc114000            call dword ptr [004011fc]
'006cb727    8bd7                    mov edx, edi
'006cb729    8bbd88feffff            mov edi, dword ptr [ebp+fffffe88]
'006cb72f    50                      push eax
'006cb730    57                      push edi
'006cb731    ff92a4000000            call dword ptr [edx+000000a4]
'006cb737    dbe2                    fnclex
'006cb739    85c0                    test ax, ax
'006cb73b    7d12                    jge 6cb74f
'006cb73d    68a4000000              push 000000a4
'006cb742    68d00d4300              push 00430dd0
'006cb747    57                      push edi
'006cb748    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006cb749    ff1580104000            call dword ptr [00401080]
'006cb74f    8d4de4                  lea ecx, dword ptr [ebp-1c]

' *** Reference to "__vbaFreeStr"
'006cb752    ff150c134000            call dword ptr [0040130c]
'#Cleanup(var_40)
'006cb758    8d45c0                  lea eax, dword ptr [ebp-40]
'006cb75b    50                      push eax
'006cb75c    8d4dc4                  lea ecx, dword ptr [ebp-3c]
'006cb75f    51                      push ecx
'006cb760    8d55c8                  lea edx, dword ptr [ebp-38]
'006cb763    52                      push edx
'006cb764    8d45cc                  lea eax, dword ptr [ebp-34]
'006cb767    50                      push eax
'006cb768    6a04                    push 04

' *** Reference to "__vbaFreeObjList"
'006cb76a    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_43, var_46, var_9, var_5)
'006cb770    83c414                  add esp, 14
'006cb773    8d4da4                  lea ecx, dword ptr [ebp-5c]

' *** Reference to "__vbaFreeVar"
'006cb776    ff1528104000            call dword ptr [00401028]
'#Cleanup(var_83)
'006cb77c    e908060000              jmp 6cbd89

'ERROR: Two many next close:
End If
'006cb781    8d8d64feffff            lea ecx, dword ptr [ebp+fffffe64]
'006cb787    51                      push ecx
'006cb788    8d9564ffffff            lea edx, dword ptr [ebp+ffffff64]
'006cb78e    52                      push edx
'006cb78f    c7856cffffff28394500    mov dword ptr [ebp+ffffff6c], 00453928
'006cb799    899d64ffffff            mov dword ptr [ebp+ffffff64], ebx
'006cb79f    ffd7                    call edi
'006cb7a1    6685c0                  test eax, eax
'006cb7a4    8b06                    mov eax, dword ptr [esi]
'006cb7a6    56                      push esi
'006cb7a7    0f8416050000            je 6cbcc3

If (((var_122) = ("VSSORTLISTE"))) Then
'006cb7ad    ff901c030000            call dword ptr [eax+0000031c]

' *** Reference to "__vbaObjSet"
'006cb7b3    8b1db4104000            mov ebx, dword ptr [004010b4]
'006cb7b9    50                      push eax
'006cb7ba    8d4dc0                  lea ecx, dword ptr [ebp-40]
'006cb7bd    51                      push ecx
'006cb7be    ffd3                    call ebx
    Set var_5 = Nothing
'006cb7c0    8b16                    mov edx, dword ptr [esi]
'006cb7c2    56                      push esi
'006cb7c3    898588feffff            mov dword ptr [ebp+fffffe88], eax
'006cb7c9    c7856cffffff00000000    mov dword ptr [ebp+ffffff6c], 00000000
'006cb7d3    c78564ffffff03000000    mov dword ptr [ebp+ffffff64], 00000003
'006cb7dd    ff92fc020000            call dword ptr [edx+000002fc]
'006cb7e3    50                      push eax
'006cb7e4    8d45cc                  lea eax, dword ptr [ebp-34]
'006cb7e7    50                      push eax
'006cb7e8    ffd3                    call ebx
    Set var_43 = Nothing
'006cb7ea    8d95a0feffff            lea edx, dword ptr [ebp+fffffea0]
'006cb7f0    8bf8                    mov edi, eax
'006cb7f2    8b0f                    mov ecx, dword ptr [edi]
'006cb7f4    52                      push edx
'006cb7f5    57                      push edi
'006cb7f6    ff91f0000000            call dword ptr [ecx+000000f0]
'006cb7fc    dbe2                    fnclex
'006cb7fe    85c0                    test ax, ax
'006cb800    7d12                    jge 6cb814
'006cb802    68f0000000              push 000000f0
'006cb807    681cb94300              push 0043b91c
'006cb80c    57                      push edi
'006cb80d    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006cb80e    ff1580104000            call dword ptr [00401080]
'006cb814    668b85a0feffff          mov ax, word ptr [ebp+fffffea0]
'006cb81b    668b4e44                mov cx, word ptr [esi+44]
'006cb81f    6689854cffffff          mov word ptr [ebp+ffffff4c], ax
'006cb826    a174b17200              mov ax, word ptr [0072b174]
'006cb82b    85c0                    test ax, ax
'006cb82d    c78544ffffff02000000    mov dword ptr [ebp+ffffff44], 00000002
'006cb837    66898d2cffffff          mov word ptr [ebp+ffffff2c], cx
'006cb83e    7515                    jne 6cb855
'006cb840    6874b17200              push 0072b174
'006cb845    6890c04100              push 0041c090

' *** Reference to "__vbaNew2"
'006cb84a    ff152c124000            call dword ptr [0040122c]
    Set var_77 = New frmFichePerso
'006cb850    a174b17200              mov ax, word ptr [0072b174]
'006cb855    8b10                    mov edx, dword ptr [eax]
'006cb857    50                      push eax
'006cb858    ff9240060000            call dword ptr [edx+00000640]
'006cb85e    50                      push eax
'006cb85f    8d45c8                  lea eax, dword ptr [ebp-38]
'006cb862    50                      push eax
'006cb863    ffd3                    call ebx
    Set var_46 = var_77
'006cb865    8bf8                    mov edi, eax
'006cb867    8b0f                    mov ecx, dword ptr [edi]
'006cb869    33c0                    xor eax, eax
'006cb86b    668b463c                mov ax, word ptr [esi+3c]
'006cb86f    8d55c4                  lea edx, dword ptr [ebp-3c]
'006cb872    52                      push edx
'006cb873    50                      push eax
'006cb874    57                      push edi
'006cb875    ff5140                  call dword ptr [ecx+40]
'006cb878    dbe2                    fnclex
'006cb87a    85c0                    test ax, ax
'006cb87c    7d0f                    jge 6cb88d
'006cb87e    6a40                    push 40
'006cb880    686c384300              push 0043386c
'006cb885    57                      push edi
'006cb886    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006cb887    ff1580104000            call dword ptr [00401080]
'006cb88d    8b8d88feffff            mov ecx, dword ptr [ebp+fffffe88]
'006cb893    8b39                    mov edi, dword ptr [ecx]
'006cb895    8b8564ffffff            mov eax, dword ptr [ebp+ffffff64]
'006cb89b    8b8d68ffffff            mov ecx, dword ptr [ebp+ffffff68]
'006cb8a1    83ec10                  sub esp, 10
'006cb8a4    8bd4                    mov edx, esp
'006cb8a6    8902                    mov dword ptr [edx], eax
'006cb8a8    8b856cffffff            mov eax, dword ptr [ebp+ffffff6c]
'006cb8ae    894a04                  mov dword ptr [edx+04], ecx
'006cb8b1    8b8d70ffffff            mov ecx, dword ptr [ebp+ffffff70]
'006cb8b7    894208                  mov dword ptr [edx+08], eax
'006cb8ba    8b8544ffffff            mov eax, dword ptr [ebp+ffffff44]
'006cb8c0    894a0c                  mov dword ptr [edx+0c], ecx
'006cb8c3    8b8d48ffffff            mov ecx, dword ptr [ebp+ffffff48]
'006cb8c9    83ec10                  sub esp, 10
'006cb8cc    8bd4                    mov edx, esp
'006cb8ce    8902                    mov dword ptr [edx], eax
'006cb8d0    8b854cffffff            mov eax, dword ptr [ebp+ffffff4c]
'006cb8d6    894a04                  mov dword ptr [edx+04], ecx
'006cb8d9    8b8d50ffffff            mov ecx, dword ptr [ebp+ffffff50]
'006cb8df    894208                  mov dword ptr [edx+08], eax
'006cb8e2    894a0c                  mov dword ptr [edx+0c], ecx
'006cb8e5    8b8d2cffffff            mov ecx, dword ptr [ebp+ffffff2c]
'006cb8eb    83ec10                  sub esp, 10
'006cb8ee    8bd4                    mov edx, esp
'006cb8f0    b802000000              mov eax, 00000002
'006cb8f5    8902                    mov dword ptr [edx], eax
'006cb8f7    8b8528ffffff            mov eax, dword ptr [ebp+ffffff28]
'006cb8fd    894204                  mov dword ptr [edx+04], eax
'006cb900    8b8530ffffff            mov eax, dword ptr [ebp+ffffff30]
'006cb906    894a08                  mov dword ptr [edx+08], ecx
'006cb909    8b4dc4                  mov ecx, dword ptr [ebp-3c]
'006cb90c    6a03                    push 03
'006cb90e    689e000000              push 0000009e
'006cb913    89420c                  mov dword ptr [edx+0c], eax
'006cb916    51                      push ecx
'006cb917    8d55a4                  lea edx, dword ptr [ebp-5c]
'006cb91a    52                      push edx

' *** Reference to "__vbaLateIdCallLd"
'006cb91b    ff157c114000            call dword ptr [0040117c]
    var_83 = var_9.UNK_-256 - 12_158
'006cb921    83c440                  add esp, 40
'006cb924    50                      push eax
'006cb925    8d45e4                  lea eax, dword ptr [ebp-1c]
'006cb928    50                      push eax

' *** Reference to "__vbaStrVarVal"
'006cb929    ff15fc114000            call dword ptr [004011fc]
'006cb92f    8bcf                    mov ecx, edi
'006cb931    8bbd88feffff            mov edi, dword ptr [ebp+fffffe88]
'006cb937    50                      push eax
'006cb938    57                      push edi
'006cb939    ff91a4000000            call dword ptr [ecx+000000a4]
'006cb93f    dbe2                    fnclex
'006cb941    85c0                    test ax, ax
'006cb943    7d12                    jge 6cb957
'006cb945    68a4000000              push 000000a4
'006cb94a    68d00d4300              push 00430dd0
'006cb94f    57                      push edi
'006cb950    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006cb951    ff1580104000            call dword ptr [00401080]
'006cb957    8d4de4                  lea ecx, dword ptr [ebp-1c]

' *** Reference to "__vbaFreeStr"
'006cb95a    ff150c134000            call dword ptr [0040130c]
    '#Cleanup(var_40)
'006cb960    8d55c0                  lea edx, dword ptr [ebp-40]
'006cb963    52                      push edx
'006cb964    8d45c4                  lea eax, dword ptr [ebp-3c]
'006cb967    50                      push eax
'006cb968    8d4dc8                  lea ecx, dword ptr [ebp-38]
'006cb96b    51                      push ecx
'006cb96c    8d55cc                  lea edx, dword ptr [ebp-34]
'006cb96f    52                      push edx
'006cb970    6a04                    push 04

' *** Reference to "__vbaFreeObjList"
'006cb972    ff154c104000            call dword ptr [0040104c]
    '#Cleanup( var_43, var_46, var_9, var_5)
'006cb978    83c414                  add esp, 14
'006cb97b    8d4da4                  lea ecx, dword ptr [ebp-5c]

' *** Reference to "__vbaFreeVar"
'006cb97e    ff1528104000            call dword ptr [00401028]
    '#Cleanup(var_83)
'006cb984    8b06                    mov eax, dword ptr [esi]
'006cb986    56                      push esi
'006cb987    ff9008030000            call dword ptr [eax+00000308]
'006cb98d    50                      push eax
'006cb98e    8d4db8                  lea ecx, dword ptr [ebp-48]
'006cb991    51                      push ecx
'006cb992    ffd3                    call ebx
    Set var_61 = Nothing
'006cb994    898580feffff            mov dword ptr [ebp+fffffe80], eax
'006cb99a    a174b17200              mov ax, word ptr [0072b174]
'006cb99f    85c0                    test ax, ax
'006cb9a1    c7856cffffff00000000    mov dword ptr [ebp+ffffff6c], 00000000
'006cb9ab    c78564ffffff03000000    mov dword ptr [ebp+ffffff64], 00000003
'006cb9b5    c7854cffffff01000000    mov dword ptr [ebp+ffffff4c], 00000001
'006cb9bf    c78544ffffff02000000    mov dword ptr [ebp+ffffff44], 00000002
'006cb9c9    7515                    jne 6cb9e0
'006cb9cb    6874b17200              push 0072b174
'006cb9d0    6890c04100              push 0041c090

' *** Reference to "__vbaNew2"
'006cb9d5    ff152c124000            call dword ptr [0040122c]
    Set var_77 = New frmFichePerso
'006cb9db    a174b17200              mov ax, word ptr [0072b174]
'006cb9e0    8b10                    mov edx, dword ptr [eax]
'006cb9e2    50                      push eax
'006cb9e3    ff9240060000            call dword ptr [edx+00000640]
'006cb9e9    50                      push eax
'006cb9ea    8d45cc                  lea eax, dword ptr [ebp-34]
'006cb9ed    50                      push eax
'006cb9ee    ffd3                    call ebx
    Set var_43 = var_77
'006cb9f0    8d55c8                  lea edx, dword ptr [ebp-38]
'006cb9f3    52                      push edx
'006cb9f4    8bf8                    mov edi, eax
'006cb9f6    8b0f                    mov ecx, dword ptr [edi]
'006cb9f8    6a01                    push 01
'006cb9fa    57                      push edi
'006cb9fb    ff5140                  call dword ptr [ecx+40]
'006cb9fe    dbe2                    fnclex
'006cba00    85c0                    test ax, ax
'006cba02    7d0f                    jge 6cba13
'006cba04    6a40                    push 40
'006cba06    686c384300              push 0043386c
'006cba0b    57                      push edi
'006cba0c    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006cba0d    ff1580104000            call dword ptr [00401080]
'006cba13    8b06                    mov eax, dword ptr [esi]
'006cba15    56                      push esi
'006cba16    c7850cffffffc48d4300    mov dword ptr [ebp+ffffff0c], 00438dc4
'006cba20    c78504ffffff08000000    mov dword ptr [ebp+ffffff04], 00000008
'006cba2a    ff90fc020000            call dword ptr [eax+000002fc]
'006cba30    50                      push eax
'006cba31    8d4dc4                  lea ecx, dword ptr [ebp-3c]
'006cba34    51                      push ecx
'006cba35    ffd3                    call ebx
    Set var_9 = Nothing
'006cba37    8bf8                    mov edi, eax
'006cba39    8b17                    mov edx, dword ptr [edi]
'006cba3b    8d85a0feffff            lea eax, dword ptr [ebp+fffffea0]
'006cba41    50                      push eax
'006cba42    57                      push edi
'006cba43    ff92f0000000            call dword ptr [edx+000000f0]
'006cba49    dbe2                    fnclex
'006cba4b    85c0                    test ax, ax
'006cba4d    7d12                    jge 6cba61
'006cba4f    68f0000000              push 000000f0
'006cba54    681cb94300              push 0043b91c
'006cba59    57                      push edi
'006cba5a    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006cba5b    ff1580104000            call dword ptr [00401080]
'006cba61    a174b17200              mov ax, word ptr [0072b174]
'006cba66    85c0                    test ax, ax
'006cba68    668b8da0feffff          mov cx, word ptr [ebp+fffffea0]
'006cba6f    668b5636                mov dx, word ptr [esi+36]
'006cba73    66898ddcfeffff          mov word ptr [ebp+fffffedc], cx
'006cba7a    668995bcfeffff          mov word ptr [ebp+fffffebc], dx
'006cba81    c785b4feffff02000000    mov dword ptr [ebp+fffffeb4], 00000002
'006cba8b    7515                    jne 6cbaa2
'006cba8d    6874b17200              push 0072b174
'006cba92    6890c04100              push 0041c090

' *** Reference to "__vbaNew2"
'006cba97    ff152c124000            call dword ptr [0040122c]
    Set var_77 = New frmFichePerso
'006cba9d    a174b17200              mov ax, word ptr [0072b174]
'006cbaa2    8b08                    mov ecx, dword ptr [eax]
'006cbaa4    50                      push eax
'006cbaa5    ff9140060000            call dword ptr [ecx+00000640]
'006cbaab    50                      push eax
'006cbaac    8d55c0                  lea edx, dword ptr [ebp-40]
'006cbaaf    52                      push edx
'006cbab0    ffd3                    call ebx
    Set var_5 = var_77
'006cbab2    33d2                    xor edx, edx
'006cbab4    668b563c                mov dx, word ptr [esi+3c]
'006cbab8    8d4dbc                  lea ecx, dword ptr [ebp-44]
'006cbabb    51                      push ecx
'006cbabc    8bf8                    mov edi, eax
'006cbabe    8b07                    mov eax, dword ptr [edi]
'006cbac0    52                      push edx
'006cbac1    57                      push edi
'006cbac2    ff5040                  call dword ptr [eax+40]
'006cbac5    dbe2                    fnclex
'006cbac7    85c0                    test ax, ax
'006cbac9    7d0f                    jge 6cbada
'006cbacb    6a40                    push 40
'006cbacd    686c384300              push 0043386c
'006cbad2    57                      push edi
'006cbad3    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006cbad4    ff1580104000            call dword ptr [00401080]
'006cbada    8b8580feffff            mov eax, dword ptr [ebp+fffffe80]
'006cbae0    8b18                    mov ebx, dword ptr [eax]
'006cbae2    8b9564ffffff            mov edx, dword ptr [ebp+ffffff64]
'006cbae8    8b8568ffffff            mov eax, dword ptr [ebp+ffffff68]
'006cbaee    83ec10                  sub esp, 10
'006cbaf1    8bcc                    mov ecx, esp
'006cbaf3    8911                    mov dword ptr [ecx], edx
'006cbaf5    8b956cffffff            mov edx, dword ptr [ebp+ffffff6c]
'006cbafb    894104                  mov dword ptr [ecx+04], eax
'006cbafe    8b8570ffffff            mov eax, dword ptr [ebp+ffffff70]
'006cbb04    895108                  mov dword ptr [ecx+08], edx
'006cbb07    8b9544ffffff            mov edx, dword ptr [ebp+ffffff44]
'006cbb0d    89410c                  mov dword ptr [ecx+0c], eax
'006cbb10    8b8548ffffff            mov eax, dword ptr [ebp+ffffff48]
'006cbb16    83ec10                  sub esp, 10
'006cbb19    8bcc                    mov ecx, esp
'006cbb1b    8911                    mov dword ptr [ecx], edx
'006cbb1d    8b954cffffff            mov edx, dword ptr [ebp+ffffff4c]
'006cbb23    894104                  mov dword ptr [ecx+04], eax
'006cbb26    8b8550ffffff            mov eax, dword ptr [ebp+ffffff50]
'006cbb2c    895108                  mov dword ptr [ecx+08], edx
'006cbb2f    8b9528ffffff            mov edx, dword ptr [ebp+ffffff28]
'006cbb35    89410c                  mov dword ptr [ecx+0c], eax
'006cbb38    83ec10                  sub esp, 10
'006cbb3b    8bcc                    mov ecx, esp
'006cbb3d    b802000000              mov eax, 00000002
'006cbb42    8901                    mov dword ptr [ecx], eax
'006cbb44    895104                  mov dword ptr [ecx+04], edx
'006cbb47    33c0                    xor eax, eax
'006cbb49    894108                  mov dword ptr [ecx+08], eax
'006cbb4c    8b8530ffffff            mov eax, dword ptr [ebp+ffffff30]
'006cbb52    6a03                    push 03
'006cbb54    89410c                  mov dword ptr [ecx+0c], eax
'006cbb57    8b4dc8                  mov ecx, dword ptr [ebp-38]
'006cbb5a    689e000000              push 0000009e
'006cbb5f    51                      push ecx
'006cbb60    8d55a4                  lea edx, dword ptr [ebp-5c]
'006cbb63    52                      push edx

' *** Reference to "__vbaLateIdCallLd"
'006cbb64    ff157c114000            call dword ptr [0040117c]
    var_83 = var_46.UNK_frmFichePerso_158

' *** Reference to "__vbaVarCat"
'006cbb6a    8b3d08124000            mov edi, dword ptr [00401208]
'006cbb70    83c440                  add esp, 40
'006cbb73    50                      push eax
'006cbb74    8d8504ffffff            lea eax, dword ptr [ebp+ffffff04]
'006cbb7a    50                      push eax
'006cbb7b    8d4d94                  lea ecx, dword ptr [ebp-6c]
'006cbb7e    51                      push ecx
'006cbb7f    ffd7                    call edi
'006cbb81    8b8d00ffffff            mov ecx, dword ptr [ebp+ffffff00]
'006cbb87    50                      push eax
'006cbb88    83ec10                  sub esp, 10
'006cbb8b    8bd4                    mov edx, esp
'006cbb8d    b803000000              mov eax, 00000003
'006cbb92    8902                    mov dword ptr [edx], eax
'006cbb94    8b85f8feffff            mov eax, dword ptr [ebp+fffffef8]
'006cbb9a    894204                  mov dword ptr [edx+04], eax
'006cbb9d    33c0                    xor eax, eax
'006cbb9f    894208                  mov dword ptr [edx+08], eax
'006cbba2    894a0c                  mov dword ptr [edx+0c], ecx
'006cbba5    8b8ddcfeffff            mov ecx, dword ptr [ebp+fffffedc]
'006cbbab    83ec10                  sub esp, 10
'006cbbae    8bd4                    mov edx, esp
'006cbbb0    b802000000              mov eax, 00000002
'006cbbb5    8902                    mov dword ptr [edx], eax
'006cbbb7    8b85d8feffff            mov eax, dword ptr [ebp+fffffed8]
'006cbbbd    894204                  mov dword ptr [edx+04], eax
'006cbbc0    8b85e0feffff            mov eax, dword ptr [ebp+fffffee0]
'006cbbc6    894a08                  mov dword ptr [edx+08], ecx
'006cbbc9    89420c                  mov dword ptr [edx+0c], eax
'006cbbcc    8b95b4feffff            mov edx, dword ptr [ebp+fffffeb4]
'006cbbd2    8b85b8feffff            mov eax, dword ptr [ebp+fffffeb8]
'006cbbd8    83ec10                  sub esp, 10
'006cbbdb    8bcc                    mov ecx, esp
'006cbbdd    8911                    mov dword ptr [ecx], edx
'006cbbdf    8b95bcfeffff            mov edx, dword ptr [ebp+fffffebc]
'006cbbe5    894104                  mov dword ptr [ecx+04], eax
'006cbbe8    8b85c0feffff            mov eax, dword ptr [ebp+fffffec0]
'006cbbee    895108                  mov dword ptr [ecx+08], edx
'006cbbf1    6a03                    push 03
'006cbbf3    89410c                  mov dword ptr [ecx+0c], eax
'006cbbf6    8b4dbc                  mov ecx, dword ptr [ebp-44]
'006cbbf9    689e000000              push 0000009e
'006cbbfe    51                      push ecx
'006cbbff    8d5584                  lea edx, dword ptr [ebp-7c]
'006cbc02    52                      push edx

' *** Reference to "__vbaLateIdCallLd"
'006cbc03    ff157c114000            call dword ptr [0040117c]
    var_119 = var_58.UNK_frmFichePerso_158
'006cbc09    83c440                  add esp, 40
'006cbc0c    50                      push eax
'006cbc0d    8d8574ffffff            lea eax, dword ptr [ebp+ffffff74]
'006cbc13    50                      push eax
'006cbc14    ffd7                    call edi
'006cbc16    50                      push eax
'006cbc17    8d4de4                  lea ecx, dword ptr [ebp-1c]
'006cbc1a    51                      push ecx

' *** Reference to "__vbaStrVarVal"
'006cbc1b    ff15fc114000            call dword ptr [004011fc]
'006cbc21    8bbd80feffff            mov edi, dword ptr [ebp+fffffe80]
'006cbc27    50                      push eax
'006cbc28    57                      push edi
'006cbc29    ff93a4000000            call dword ptr [ebx+000000a4]
'006cbc2f    dbe2                    fnclex
'006cbc31    85c0                    test ax, ax
'006cbc33    7d12                    jge 6cbc47
'006cbc35    68a4000000              push 000000a4
'006cbc3a    68d00d4300              push 00430dd0
'006cbc3f    57                      push edi
'006cbc40    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006cbc41    ff1580104000            call dword ptr [00401080]
'006cbc47    8d4de4                  lea ecx, dword ptr [ebp-1c]

' *** Reference to "__vbaFreeStr"
'006cbc4a    ff150c134000            call dword ptr [0040130c]
    '#Cleanup(var_40)
'006cbc50    8d55b8                  lea edx, dword ptr [ebp-48]
'006cbc53    52                      push edx
'006cbc54    8d45bc                  lea eax, dword ptr [ebp-44]
'006cbc57    50                      push eax
'006cbc58    8d4dc0                  lea ecx, dword ptr [ebp-40]
'006cbc5b    51                      push ecx
'006cbc5c    8d55c4                  lea edx, dword ptr [ebp-3c]
'006cbc5f    52                      push edx
'006cbc60    8d45c8                  lea eax, dword ptr [ebp-38]
'006cbc63    50                      push eax
'006cbc64    8d4dcc                  lea ecx, dword ptr [ebp-34]
'006cbc67    51                      push ecx
'006cbc68    6a06                    push 06

' *** Reference to "__vbaFreeObjList"
'006cbc6a    ff154c104000            call dword ptr [0040104c]
    '#Cleanup( var_43, var_46, var_9, var_5, var_58, var_61)
'006cbc70    8d9574ffffff            lea edx, dword ptr [ebp+ffffff74]
'006cbc76    52                      push edx
'006cbc77    8d4584                  lea eax, dword ptr [ebp-7c]
'006cbc7a    50                      push eax
'006cbc7b    8d4d94                  lea ecx, dword ptr [ebp-6c]
'006cbc7e    51                      push ecx
'006cbc7f    8d55a4                  lea edx, dword ptr [ebp-5c]
'006cbc82    52                      push edx
'006cbc83    6a04                    push 04

' *** Reference to "__vbaFreeVarList"
'006cbc85    ff1540104000            call dword ptr [00401040]
    '#Cleanup( var_83, var_121, var_119, var_91)
'006cbc8b    8b06                    mov eax, dword ptr [esi]
'006cbc8d    83c430                  add esp, 30
'006cbc90    56                      push esi
'006cbc91    ff9018030000            call dword ptr [eax+00000318]

' *** Reference to "__vbaObjSet"
'006cbc97    8b1db4104000            mov ebx, dword ptr [004010b4]
'006cbc9d    50                      push eax
'006cbc9e    8d4dcc                  lea ecx, dword ptr [ebp-34]
'006cbca1    51                      push ecx
'006cbca2    ffd3                    call ebx
    Set var_43 = Nothing
'006cbca4    8bf8                    mov edi, eax
'006cbca6    8b17                    mov edx, dword ptr [edi]
'006cbca8    68cc134300              push 004313cc
'006cbcad    57                      push edi
'006cbcae    ff92a4000000            call dword ptr [edx+000000a4]
'006cbcb4    dbe2                    fnclex
'006cbcb6    85c0                    test ax, ax
'006cbcb8    0f8dc2000000            jge 6cbd80
'006cbcbe    e9ab000000              jmp 6cbd6e
    
End If
'006cbcc3    ff901c030000            call dword ptr [eax+0000031c]

' *** Reference to "__vbaObjSet"
'006cbcc9    8b1db4104000            mov ebx, dword ptr [004010b4]
'006cbccf    50                      push eax
'006cbcd0    8d4dcc                  lea ecx, dword ptr [ebp-34]
'006cbcd3    51                      push ecx
'006cbcd4    ffd3                    call ebx
Set var_43 = Nothing
'006cbcd6    8bf8                    mov edi, eax
'006cbcd8    8b17                    mov edx, dword ptr [edi]
'006cbcda    68cc134300              push 004313cc
'006cbcdf    57                      push edi
'006cbce0    ff92a4000000            call dword ptr [edx+000000a4]
'006cbce6    dbe2                    fnclex
'006cbce8    85c0                    test ax, ax
'006cbcea    7d12                    jge 6cbcfe
'006cbcec    68a4000000              push 000000a4
'006cbcf1    68d00d4300              push 00430dd0
'006cbcf6    57                      push edi
'006cbcf7    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006cbcf8    ff1580104000            call dword ptr [00401080]
'006cbcfe    8d4dcc                  lea ecx, dword ptr [ebp-34]

' *** Reference to "__vbaFreeObj"
'006cbd01    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_43)
'006cbd07    8b06                    mov eax, dword ptr [esi]
'006cbd09    56                      push esi
'006cbd0a    ff9008030000            call dword ptr [eax+00000308]
'006cbd10    50                      push eax
'006cbd11    8d4dcc                  lea ecx, dword ptr [ebp-34]
'006cbd14    51                      push ecx
'006cbd15    ffd3                    call ebx
Set var_43 = Nothing
'006cbd17    8bf8                    mov edi, eax
'006cbd19    8b17                    mov edx, dword ptr [edi]
'006cbd1b    68cc134300              push 004313cc
'006cbd20    57                      push edi
'006cbd21    ff92a4000000            call dword ptr [edx+000000a4]
'006cbd27    dbe2                    fnclex
'006cbd29    85c0                    test ax, ax
'006cbd2b    7d12                    jge 6cbd3f
'006cbd2d    68a4000000              push 000000a4
'006cbd32    68d00d4300              push 00430dd0
'006cbd37    57                      push edi
'006cbd38    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006cbd39    ff1580104000            call dword ptr [00401080]
'006cbd3f    8d4dcc                  lea ecx, dword ptr [ebp-34]

' *** Reference to "__vbaFreeObj"
'006cbd42    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_43)
'006cbd48    8b06                    mov eax, dword ptr [esi]
'006cbd4a    56                      push esi
'006cbd4b    ff9018030000            call dword ptr [eax+00000318]
'006cbd51    50                      push eax
'006cbd52    8d4dcc                  lea ecx, dword ptr [ebp-34]
'006cbd55    51                      push ecx
'006cbd56    ffd3                    call ebx
Set var_43 = Nothing
'006cbd58    8bf8                    mov edi, eax
'006cbd5a    8b17                    mov edx, dword ptr [edi]
'006cbd5c    68cc134300              push 004313cc
'006cbd61    57                      push edi
'006cbd62    ff92a4000000            call dword ptr [edx+000000a4]
'006cbd68    dbe2                    fnclex
'006cbd6a    85c0                    test ax, ax
'006cbd6c    7d12                    jge 6cbd80
'006cbd6e    68a4000000              push 000000a4
'006cbd73    68d00d4300              push 00430dd0
'006cbd78    57                      push edi
'006cbd79    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006cbd7a    ff1580104000            call dword ptr [00401080]
'006cbd80    8d4dcc                  lea ecx, dword ptr [ebp-34]

' *** Reference to "__vbaFreeObj"
'006cbd83    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_43)
'ERROR: Two many next close:
End If
'006cbd89    8b45e8                  mov eax, dword ptr [ebp-18]
'006cbd8c    8b08                    mov ecx, dword ptr [eax]
'006cbd8e    50                      push eax
'006cbd8f    ff91c4000000            call dword ptr [ecx+000000c4]
'006cbd95    dbe2                    fnclex
'006cbd97    85c0                    test ax, ax
'006cbd99    7d15                    jge 6cbdb0
'006cbd9b    8b55e8                  mov edx, dword ptr [ebp-18]
'006cbd9e    68c4000000              push 000000c4
'006cbda3    6830314300              push 00433130
'006cbda8    52                      push edx
'006cbda9    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006cbdaa    ff1580104000            call dword ptr [00401080]

'ERROR: Two many next close:
End If
'006cbdb0    8b06                    mov eax, dword ptr [esi]
'006cbdb2    56                      push esi
'006cbdb3    ff9030030000            call dword ptr [eax+00000330]
'006cbdb9    50                      push eax
'006cbdba    8d4dcc                  lea ecx, dword ptr [ebp-34]
'006cbdbd    51                      push ecx
'006cbdbe    ffd3                    call ebx
Set var_43 = Nothing
'006cbdc0    8b45cc                  mov eax, dword ptr [ebp-34]
'006cbdc3    6a06                    push 06
'006cbdc5    8d55a4                  lea edx, dword ptr [ebp-5c]
'006cbdc8    8945ac                  mov dword ptr [ebp-54], eax
'006cbdcb    52                      push edx
'006cbdcc    8d4594                  lea eax, dword ptr [ebp-6c]
'006cbdcf    50                      push eax
'006cbdd0    c745cc00000000          mov dword ptr [ebp-34], 00000000
'006cbdd7    c745a409000000          mov dword ptr [ebp-5c], 00000009

' *** Reference to "rtcLeftCharVar"
'006cbdde    ff15bc124000            call dword ptr [004012bc]
'006cbde4    8d4d94                  lea ecx, dword ptr [ebp-6c]
'006cbde7    51                      push ecx
'006cbde8    8d5584                  lea edx, dword ptr [ebp-7c]
'006cbdeb    52                      push edx

' *** Reference to "rtcUpperCaseVar"
'006cbdec    ff152c114000            call dword ptr [0040112c]
'006cbdf2    8d4584                  lea eax, dword ptr [ebp-7c]
'006cbdf5    50                      push eax
'006cbdf6    8d8d64ffffff            lea ecx, dword ptr [ebp+ffffff64]
'006cbdfc    51                      push ecx
'006cbdfd    c7856cffffffb0394500    mov dword ptr [ebp+ffffff6c], 004539b0
'006cbe07    c78564ffffff08800000    mov dword ptr [ebp+ffffff64], 00008008

' *** Reference to "__vbaVarTstEq"
'006cbe11    ff153c114000            call dword ptr [0040113c]

' *** Reference to "__vbaFreeObj"
'006cbe17    8b3d08134000            mov edi, dword ptr [00401308]
'006cbe1d    8d4dcc                  lea ecx, dword ptr [ebp-34]
'006cbe20    66898598feffff          mov word ptr [ebp+fffffe98], ax
'006cbe27    ffd7                    call edi
'#Cleanup(var_43)
'006cbe29    8d5584                  lea edx, dword ptr [ebp-7c]
'006cbe2c    52                      push edx
'006cbe2d    8d4594                  lea eax, dword ptr [ebp-6c]
'006cbe30    50                      push eax
'006cbe31    8d4da4                  lea ecx, dword ptr [ebp-5c]
'006cbe34    51                      push ecx
'006cbe35    6a03                    push 03

' *** Reference to "__vbaFreeVarList"
'006cbe37    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_83, var_121, var_119)
'006cbe3d    8b16                    mov edx, dword ptr [esi]
'006cbe3f    83c410                  add esp, 10
'006cbe42    6683bd98feffff00        cmp word ptr [ebp+fffffe98], 00
'006cbe4a    56                      push esi
'006cbe4b    0f84be010000            je 6cc00f

If (((UCase(Left(var_43, 6))) = ("EPIQUE"))) Then
'006cbe51    ff928c030000            call dword ptr [edx+0000038c]
'006cbe57    50                      push eax
'006cbe58    8d45cc                  lea eax, dword ptr [ebp-34]
'006cbe5b    50                      push eax
'006cbe5c    ffd3                    call ebx
    Set var_43 = 
'006cbe5e    8b08                    mov ecx, dword ptr [eax]
'006cbe60    6800303145              push 45313000
'006cbe65    50                      push eax
'006cbe66    898598feffff            mov dword ptr [ebp+fffffe98], eax
'006cbe6c    ff517c                  call dword ptr [ecx+7c]
'006cbe6f    dbe2                    fnclex
'006cbe71    85c0                    test ax, ax
'006cbe73    7d15                    jge 6cbe8a
    
    If (    var_43) Then
'006cbe75    8b9598feffff            mov edx, dword ptr [ebp+fffffe98]
'006cbe7b    6a7c                    push 7c
'006cbe7d    683ce04300              push 0043e03c
'006cbe82    52                      push edx
'006cbe83    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006cbe84    ff1580104000            call dword ptr [00401080]
    
End If
'006cbe8a    8d4dcc                  lea ecx, dword ptr [ebp-34]
'006cbe8d    ffd7                    call edi
'#Cleanup(var_43)
'006cbe8f    8b06                    mov eax, dword ptr [esi]
'006cbe91    56                      push esi
'006cbe92    ff9054030000            call dword ptr [eax+00000354]
'006cbe98    50                      push eax
'006cbe99    8d4dcc                  lea ecx, dword ptr [ebp-34]
'006cbe9c    51                      push ecx
'006cbe9d    ffd3                    call ebx
Set var_43 = Nothing
'006cbe9f    8b10                    mov edx, dword ptr [eax]
'006cbea1    6800004845              push 45480000
'006cbea6    50                      push eax
'006cbea7    898598feffff            mov dword ptr [ebp+fffffe98], eax
'006cbead    ff5274                  call dword ptr [edx+74]
'006cbeb0    dbe2                    fnclex
'006cbeb2    85c0                    test ax, ax
'006cbeb4    7d15                    jge 6cbecb
'006cbeb6    8b8d98feffff            mov ecx, dword ptr [ebp+fffffe98]
'006cbebc    6a74                    push 74
'006cbebe    68d00d4300              push 00430dd0
'006cbec3    51                      push ecx
'006cbec4    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006cbec5    ff1580104000            call dword ptr [00401080]
'006cbecb    8d4dcc                  lea ecx, dword ptr [ebp-34]
'006cbece    ffd7                    call edi
'#Cleanup(var_43)
'006cbed0    8b16                    mov edx, dword ptr [esi]
'006cbed2    56                      push esi
'006cbed3    ff9254030000            call dword ptr [edx+00000354]
'006cbed9    50                      push eax
'006cbeda    8d45cc                  lea eax, dword ptr [ebp-34]
'006cbedd    50                      push eax
'006cbede    ffd3                    call ebx
Set var_43 = Nothing
'006cbee0    8b08                    mov ecx, dword ptr [eax]
'006cbee2    6800703c45              push 453c7000
'006cbee7    50                      push eax
'006cbee8    898598feffff            mov dword ptr [ebp+fffffe98], eax
'006cbeee    ff9184000000            call dword ptr [ecx+00000084]
'006cbef4    dbe2                    fnclex
'006cbef6    85c0                    test ax, ax
'006cbef8    7d18                    jge 6cbf12
'006cbefa    8b9598feffff            mov edx, dword ptr [ebp+fffffe98]
'006cbf00    6884000000              push 00000084
'006cbf05    68d00d4300              push 00430dd0
'006cbf0a    52                      push edx
'006cbf0b    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006cbf0c    ff1580104000            call dword ptr [00401080]
'006cbf12    8d4dcc                  lea ecx, dword ptr [ebp-34]
'006cbf15    ffd7                    call edi
'#Cleanup(var_43)
'006cbf17    8b06                    mov eax, dword ptr [esi]
'006cbf19    56                      push esi
'006cbf1a    ff9024030000            call dword ptr [eax+00000324]
'006cbf20    50                      push eax
'006cbf21    8d4dcc                  lea ecx, dword ptr [ebp-34]
'006cbf24    51                      push ecx
'006cbf25    ffd3                    call ebx
Set var_43 = Nothing
'006cbf27    8b10                    mov edx, dword ptr [eax]
'006cbf29    6aff                    push ffffffff
'006cbf2b    50                      push eax
'006cbf2c    898598feffff            mov dword ptr [ebp+fffffe98], eax
'006cbf32    ff9294000000            call dword ptr [edx+00000094]
'006cbf38    dbe2                    fnclex
'006cbf3a    85c0                    test ax, ax
'006cbf3c    7d18                    jge 6cbf56
'006cbf3e    8b8d98feffff            mov ecx, dword ptr [ebp+fffffe98]
'006cbf44    6894000000              push 00000094
'006cbf49    68d00d4300              push 00430dd0
'006cbf4e    51                      push ecx
'006cbf4f    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006cbf50    ff1580104000            call dword ptr [00401080]
'006cbf56    8d4dcc                  lea ecx, dword ptr [ebp-34]
'006cbf59    ffd7                    call edi
'#Cleanup(var_43)
'006cbf5b    8b16                    mov edx, dword ptr [esi]
'006cbf5d    56                      push esi
'006cbf5e    ff925c030000            call dword ptr [edx+0000035c]
'006cbf64    50                      push eax
'006cbf65    8d45cc                  lea eax, dword ptr [ebp-34]
'006cbf68    50                      push eax
'006cbf69    ffd3                    call ebx
Set var_43 = Nothing
'006cbf6b    8b08                    mov ecx, dword ptr [eax]
'006cbf6d    6aff                    push ffffffff
'006cbf6f    50                      push eax
'006cbf70    898598feffff            mov dword ptr [ebp+fffffe98], eax
'006cbf76    ff919c000000            call dword ptr [ecx+0000009c]
'006cbf7c    dbe2                    fnclex
'006cbf7e    85c0                    test ax, ax
'006cbf80    7d18                    jge 6cbf9a
'006cbf82    8b9598feffff            mov edx, dword ptr [ebp+fffffe98]
'006cbf88    689c000000              push 0000009c
'006cbf8d    683ce04300              push 0043e03c
'006cbf92    52                      push edx
'006cbf93    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006cbf94    ff1580104000            call dword ptr [00401080]
'006cbf9a    8d4dcc                  lea ecx, dword ptr [ebp-34]
'006cbf9d    ffd7                    call edi
'#Cleanup(var_43)
'006cbf9f    8b06                    mov eax, dword ptr [esi]
'006cbfa1    56                      push esi
'006cbfa2    ff9028030000            call dword ptr [eax+00000328]
'006cbfa8    50                      push eax
'006cbfa9    8d4dcc                  lea ecx, dword ptr [ebp-34]
'006cbfac    51                      push ecx
'006cbfad    ffd3                    call ebx
Set var_43 = Nothing
'006cbfaf    8b10                    mov edx, dword ptr [eax]
'006cbfb1    6aff                    push ffffffff
'006cbfb3    50                      push eax
'006cbfb4    898598feffff            mov dword ptr [ebp+fffffe98], eax
'006cbfba    ff9294000000            call dword ptr [edx+00000094]
'006cbfc0    dbe2                    fnclex
'006cbfc2    85c0                    test ax, ax
'006cbfc4    7d18                    jge 6cbfde
'006cbfc6    8b8d98feffff            mov ecx, dword ptr [ebp+fffffe98]
'006cbfcc    6894000000              push 00000094
'006cbfd1    68d00d4300              push 00430dd0
'006cbfd6    51                      push ecx
'006cbfd7    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006cbfd8    ff1580104000            call dword ptr [00401080]
'006cbfde    8d4dcc                  lea ecx, dword ptr [ebp-34]
'006cbfe1    ffd7                    call edi
'#Cleanup(var_43)
'006cbfe3    8b16                    mov edx, dword ptr [esi]
'006cbfe5    56                      push esi
'006cbfe6    ff9260030000            call dword ptr [edx+00000360]
'006cbfec    50                      push eax
'006cbfed    8d45cc                  lea eax, dword ptr [ebp-34]
'006cbff0    50                      push eax
'006cbff1    ffd3                    call ebx
Set var_43 = Nothing
'006cbff3    8bf0                    mov esi, eax
'006cbff5    8b0e                    mov ecx, dword ptr [esi]
'006cbff7    6aff                    push ffffffff
'006cbff9    56                      push esi
'006cbffa    ff919c000000            call dword ptr [ecx+0000009c]
'006cc000    dbe2                    fnclex
'006cc002    85c0                    test ax, ax
'006cc004    0f8dcc010000            jge 6cc1d6
'006cc00a    e9b5010000              jmp 6cc1c4

'ERROR: Two many next close:
End If
'006cc00f    ff9254030000            call dword ptr [edx+00000354]
'006cc015    50                      push eax
'006cc016    8d45cc                  lea eax, dword ptr [ebp-34]
'006cc019    50                      push eax
'006cc01a    ffd3                    call ebx
Set var_43 = 
'006cc01c    8b08                    mov ecx, dword ptr [eax]
'006cc01e    680000c843              push 43c80000
'006cc023    50                      push eax
'006cc024    898598feffff            mov dword ptr [ebp+fffffe98], eax
'006cc02a    ff5174                  call dword ptr [ecx+74]
'006cc02d    dbe2                    fnclex
'006cc02f    85c0                    test ax, ax
'006cc031    7d15                    jge 6cc048
'006cc033    8b9598feffff            mov edx, dword ptr [ebp+fffffe98]
'006cc039    6a74                    push 74
'006cc03b    68d00d4300              push 00430dd0
'006cc040    52                      push edx
'006cc041    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006cc042    ff1580104000            call dword ptr [00401080]
'006cc048    8d4dcc                  lea ecx, dword ptr [ebp-34]
'006cc04b    ffd7                    call edi
'#Cleanup(var_43)
'006cc04d    8b06                    mov eax, dword ptr [esi]
'006cc04f    56                      push esi
'006cc050    ff9054030000            call dword ptr [eax+00000354]
'006cc056    50                      push eax
'006cc057    8d4dcc                  lea ecx, dword ptr [ebp-34]
'006cc05a    51                      push ecx
'006cc05b    ffd3                    call ebx
Set var_43 = Nothing
'006cc05d    8b10                    mov edx, dword ptr [eax]
'006cc05f    680040b545              push 45b54000
'006cc064    50                      push eax
'006cc065    898598feffff            mov dword ptr [ebp+fffffe98], eax
'006cc06b    ff9284000000            call dword ptr [edx+00000084]
'006cc071    dbe2                    fnclex
'006cc073    85c0                    test ax, ax
'006cc075    7d18                    jge 6cc08f
'006cc077    8b8d98feffff            mov ecx, dword ptr [ebp+fffffe98]
'006cc07d    6884000000              push 00000084
'006cc082    68d00d4300              push 00430dd0
'006cc087    51                      push ecx
'006cc088    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006cc089    ff1580104000            call dword ptr [00401080]
'006cc08f    8d4dcc                  lea ecx, dword ptr [ebp-34]
'006cc092    ffd7                    call edi
'#Cleanup(var_43)
'006cc094    8b16                    mov edx, dword ptr [esi]
'006cc096    56                      push esi
'006cc097    ff928c030000            call dword ptr [edx+0000038c]
'006cc09d    50                      push eax
'006cc09e    8d45cc                  lea eax, dword ptr [ebp-34]
'006cc0a1    50                      push eax
'006cc0a2    ffd3                    call ebx
Set var_43 = Nothing
'006cc0a4    8b08                    mov ecx, dword ptr [eax]
'006cc0a6    6800001643              push 43160000
'006cc0ab    50                      push eax
'006cc0ac    898598feffff            mov dword ptr [ebp+fffffe98], eax
'006cc0b2    ff517c                  call dword ptr [ecx+7c]
'006cc0b5    dbe2                    fnclex
'006cc0b7    85c0                    test ax, ax
'006cc0b9    7d15                    jge 6cc0d0
'006cc0bb    8b9598feffff            mov edx, dword ptr [ebp+fffffe98]
'006cc0c1    6a7c                    push 7c
'006cc0c3    683ce04300              push 0043e03c
'006cc0c8    52                      push edx
'006cc0c9    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006cc0ca    ff1580104000            call dword ptr [00401080]
'006cc0d0    8d4dcc                  lea ecx, dword ptr [ebp-34]
'006cc0d3    ffd7                    call edi
'#Cleanup(var_43)
'006cc0d5    8b06                    mov eax, dword ptr [esi]
'006cc0d7    56                      push esi
'006cc0d8    ff9024030000            call dword ptr [eax+00000324]
'006cc0de    50                      push eax
'006cc0df    8d4dcc                  lea ecx, dword ptr [ebp-34]
'006cc0e2    51                      push ecx
'006cc0e3    ffd3                    call ebx
Set var_43 = Nothing
'006cc0e5    8b10                    mov edx, dword ptr [eax]
'006cc0e7    6a00                    push 00
'006cc0e9    50                      push eax
'006cc0ea    898598feffff            mov dword ptr [ebp+fffffe98], eax
'006cc0f0    ff9294000000            call dword ptr [edx+00000094]
'006cc0f6    dbe2                    fnclex
'006cc0f8    85c0                    test ax, ax
'006cc0fa    7d18                    jge 6cc114
'006cc0fc    8b8d98feffff            mov ecx, dword ptr [ebp+fffffe98]
'006cc102    6894000000              push 00000094
'006cc107    68d00d4300              push 00430dd0
'006cc10c    51                      push ecx
'006cc10d    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006cc10e    ff1580104000            call dword ptr [00401080]
'006cc114    8d4dcc                  lea ecx, dword ptr [ebp-34]
'006cc117    ffd7                    call edi
'#Cleanup(var_43)
'006cc119    8b16                    mov edx, dword ptr [esi]
'006cc11b    56                      push esi
'006cc11c    ff925c030000            call dword ptr [edx+0000035c]
'006cc122    50                      push eax
'006cc123    8d45cc                  lea eax, dword ptr [ebp-34]
'006cc126    50                      push eax
'006cc127    ffd3                    call ebx
Set var_43 = Nothing
'006cc129    8b08                    mov ecx, dword ptr [eax]
'006cc12b    6a00                    push 00
'006cc12d    50                      push eax
'006cc12e    898598feffff            mov dword ptr [ebp+fffffe98], eax
'006cc134    ff919c000000            call dword ptr [ecx+0000009c]
'006cc13a    dbe2                    fnclex
'006cc13c    85c0                    test ax, ax
'006cc13e    7d18                    jge 6cc158
'006cc140    8b9598feffff            mov edx, dword ptr [ebp+fffffe98]
'006cc146    689c000000              push 0000009c
'006cc14b    683ce04300              push 0043e03c
'006cc150    52                      push edx
'006cc151    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006cc152    ff1580104000            call dword ptr [00401080]
'006cc158    8d4dcc                  lea ecx, dword ptr [ebp-34]
'006cc15b    ffd7                    call edi
'#Cleanup(var_43)
'006cc15d    8b06                    mov eax, dword ptr [esi]
'006cc15f    56                      push esi
'006cc160    ff9028030000            call dword ptr [eax+00000328]
'006cc166    50                      push eax
'006cc167    8d4dcc                  lea ecx, dword ptr [ebp-34]
'006cc16a    51                      push ecx
'006cc16b    ffd3                    call ebx
Set var_43 = Nothing
'006cc16d    8b10                    mov edx, dword ptr [eax]
'006cc16f    6a00                    push 00
'006cc171    50                      push eax
'006cc172    898598feffff            mov dword ptr [ebp+fffffe98], eax
'006cc178    ff9294000000            call dword ptr [edx+00000094]
'006cc17e    dbe2                    fnclex
'006cc180    85c0                    test ax, ax
'006cc182    7d18                    jge 6cc19c
'006cc184    8b8d98feffff            mov ecx, dword ptr [ebp+fffffe98]
'006cc18a    6894000000              push 00000094
'006cc18f    68d00d4300              push 00430dd0
'006cc194    51                      push ecx
'006cc195    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006cc196    ff1580104000            call dword ptr [00401080]
'006cc19c    8d4dcc                  lea ecx, dword ptr [ebp-34]
'006cc19f    ffd7                    call edi
'#Cleanup(var_43)
'006cc1a1    8b16                    mov edx, dword ptr [esi]
'006cc1a3    56                      push esi
'006cc1a4    ff9260030000            call dword ptr [edx+00000360]
'006cc1aa    50                      push eax
'006cc1ab    8d45cc                  lea eax, dword ptr [ebp-34]
'006cc1ae    50                      push eax
'006cc1af    ffd3                    call ebx
Set var_43 = Nothing
'006cc1b1    8bf0                    mov esi, eax
'006cc1b3    8b0e                    mov ecx, dword ptr [esi]
'006cc1b5    6a00                    push 00
'006cc1b7    56                      push esi
'006cc1b8    ff919c000000            call dword ptr [ecx+0000009c]
'006cc1be    dbe2                    fnclex
'006cc1c0    85c0                    test ax, ax
'006cc1c2    7d12                    jge 6cc1d6
'006cc1c4    689c000000              push 0000009c
'006cc1c9    683ce04300              push 0043e03c
'006cc1ce    56                      push esi
'006cc1cf    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006cc1d0    ff1580104000            call dword ptr [00401080]
'006cc1d6    8d4dcc                  lea ecx, dword ptr [ebp-34]
'006cc1d9    ffd7                    call edi
'#Cleanup(var_43)
'006cc1db    9b                      fwait
'006cc1dc    685cc26c00              push 006cc25c
'006cc1e1    eb63                    jmp 6cc246
'006cc1e3    8d55d0                  lea edx, dword ptr [ebp-30]
'006cc1e6    52                      push edx
'006cc1e7    8d45d4                  lea eax, dword ptr [ebp-2c]
'006cc1ea    50                      push eax
'006cc1eb    8d4dd8                  lea ecx, dword ptr [ebp-28]
'006cc1ee    51                      push ecx
'006cc1ef    8d55dc                  lea edx, dword ptr [ebp-24]
'006cc1f2    52                      push edx
'006cc1f3    8d45e0                  lea eax, dword ptr [ebp-20]
'006cc1f6    50                      push eax
'006cc1f7    8d4de4                  lea ecx, dword ptr [ebp-1c]
'006cc1fa    51                      push ecx
'006cc1fb    6a06                    push 06

' *** Reference to "__vbaFreeStrList"
'006cc1fd    ff155c124000            call dword ptr [0040125c]
'#Cleanup( , -4516, 0, -4500, -4504, __vbaObjSet)
'006cc203    8d55b4                  lea edx, dword ptr [ebp-4c]
'006cc206    52                      push edx
'006cc207    8d45b8                  lea eax, dword ptr [ebp-48]
'006cc20a    50                      push eax
'006cc20b    8d4dbc                  lea ecx, dword ptr [ebp-44]
'006cc20e    51                      push ecx
'006cc20f    8d55c0                  lea edx, dword ptr [ebp-40]
'006cc212    52                      push edx
'006cc213    8d45c4                  lea eax, dword ptr [ebp-3c]
'006cc216    50                      push eax
'006cc217    8d4dc8                  lea ecx, dword ptr [ebp-38]
'006cc21a    51                      push ecx
'006cc21b    8d55cc                  lea edx, dword ptr [ebp-34]
'006cc21e    52                      push edx
'006cc21f    6a07                    push 07

' *** Reference to "__vbaFreeObjList"
'006cc221    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_43, var_46, var_9, var_5, var_58, var_61, var_62)
'006cc227    8d8574ffffff            lea eax, dword ptr [ebp+ffffff74]
'006cc22d    50                      push eax
'006cc22e    8d4d84                  lea ecx, dword ptr [ebp-7c]
'006cc231    51                      push ecx
'006cc232    8d5594                  lea edx, dword ptr [ebp-6c]
'006cc235    52                      push edx
'006cc236    8d45a4                  lea eax, dword ptr [ebp-5c]
'006cc239    50                      push eax
'006cc23a    6a04                    push 04

' *** Reference to "__vbaFreeVarList"
'006cc23c    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_83, var_121, var_119, var_91)
'006cc242    83c450                  add esp, 50
'006cc245    c3                      ret
'006cc246    8d8d64feffff            lea ecx, dword ptr [ebp+fffffe64]

' *** Reference to "__vbaFreeVar"
'006cc24c    ff1528104000            call dword ptr [00401028]
'#Cleanup(var_122)
'006cc252    8d4de8                  lea ecx, dword ptr [ebp-18]

' *** Reference to "__vbaFreeObj"
'006cc255    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_41)
'006cc25b    c3                      ret
'006cc25c    8b4508                  mov eax, dword ptr [ebp+08]
'006cc25f    8b08                    mov ecx, dword ptr [eax]
'006cc261    50                      push eax
'006cc262    ff5108                  call dword ptr [ecx+08]
'006cc265    8b45fc                  mov eax, dword ptr [ebp-04]
'006cc268    8b4dec                  mov ecx, dword ptr [ebp-14]
'006cc26b    5f                      pop edi
'006cc26c    5e                      pop esi
    'Reference to '__except_list'
'006cc26d    64890d00000000          mov dword ptr fs:[00000000], ecx
'006cc274    5b                      pop ebx
'006cc275    8be5                    mov esp, ebp
'006cc277    5d                      pop ebp
'006cc278    c20400                  ret 0004


End Function


'Event for CmbSort
Private Sub CmbSort_Click()
'006c7dd0    55                      push ebp
'006c7dd1    8bec                    mov ebp, esp
'006c7dd3    83ec0c                  sub esp, 0c

' *** Reference to "__vbaExceptHandler"
'006c7dd6    6866724000              push 00407266
'006c7ddb    64a100000000            mov ax, word ptr fs:[00000000]
'006c7de1    50                      push eax
    'Reference to '__except_list'
'006c7de2    64892500000000          mov dword ptr fs:[00000000], esp
'006c7de9    83ec0c                  sub esp, 0c
'006c7dec    53                      push ebx
'006c7ded    56                      push esi
'006c7dee    57                      push edi
'006c7def    8965f4                  mov dword ptr [ebp-0c], esp
'006c7df2    c745f880674000          mov dword ptr [ebp-08], 00406780
'006c7df9    8b7508                  mov esi, dword ptr [ebp+08]
'006c7dfc    8bc6                    mov eax, esi
'006c7dfe    83e001                  and eax, 01
'006c7e01    8945fc                  mov dword ptr [ebp-04], eax
'006c7e04    83e6fe                  and esi, fffffffe
'006c7e07    8b0e                    mov ecx, dword ptr [esi]
'006c7e09    56                      push esi
'006c7e0a    897508                  mov dword ptr [ebp+08], esi
'006c7e0d    ff5104                  call dword ptr [ecx+04]
'006c7e10    8b16                    mov edx, dword ptr [esi]
'006c7e12    56                      push esi
'006c7e13    ff9230070000            call dword ptr [edx+00000730]
'006c7e19    85c0                    test ax, ax
'006c7e1b    7d12                    jge 6c7e2f
'006c7e1d    6830070000              push 00000730
'006c7e22    6800144300              push 00431400
'006c7e27    56                      push esi
'006c7e28    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c7e29    ff1580104000            call dword ptr [00401080]
'006c7e2f    c745fc00000000          mov dword ptr [ebp-04], 00000000
'006c7e36    8b4508                  mov eax, dword ptr [ebp+08]
'006c7e39    8b08                    mov ecx, dword ptr [eax]
'006c7e3b    50                      push eax
'006c7e3c    ff5108                  call dword ptr [ecx+08]
'006c7e3f    8b45fc                  mov eax, dword ptr [ebp-04]
'006c7e42    8b4dec                  mov ecx, dword ptr [ebp-14]
'006c7e45    5f                      pop edi
'006c7e46    5e                      pop esi
    'Reference to '__except_list'
'006c7e47    64890d00000000          mov dword ptr fs:[00000000], ecx
'006c7e4e    5b                      pop ebx
'006c7e4f    8be5                    mov esp, ebp
'006c7e51    5d                      pop ebp
'006c7e52    c20400                  ret 0004


End Sub


'Event for btnEnregistrer
Private Sub btnEnregistrer_Click()
'006c65f0    55                      push ebp
'006c65f1    8bec                    mov ebp, esp
'006c65f3    83ec0c                  sub esp, 0c

' *** Reference to "__vbaExceptHandler"
'006c65f6    6866724000              push 00407266
'006c65fb    64a100000000            mov ax, word ptr fs:[00000000]
'006c6601    50                      push eax
    'Reference to '__except_list'
'006c6602    64892500000000          mov dword ptr fs:[00000000], esp
'006c6609    81ec50010000            sub esp, 00000150
'006c660f    53                      push ebx
'006c6610    56                      push esi
'006c6611    57                      push edi
'006c6612    8965f4                  mov dword ptr [ebp-0c], esp
'006c6615    c745f840674000          mov dword ptr [ebp-08], 00406740
'006c661c    8b7508                  mov esi, dword ptr [ebp+08]
'006c661f    8bc6                    mov eax, esi
'006c6621    83e001                  and eax, 01
'006c6624    8945fc                  mov dword ptr [ebp-04], eax
'006c6627    83e6fe                  and esi, fffffffe
'006c662a    8b0e                    mov ecx, dword ptr [esi]
'006c662c    56                      push esi
'006c662d    897508                  mov dword ptr [ebp+08], esi
'006c6630    ff5104                  call dword ptr [ecx+04]
'006c6633    33ff                    xor edi, edi
'006c6635    8d5de4                  lea ebx, dword ptr [ebp-1c]
'006c6638    53                      push ebx
'006c6639    83ec10                  sub esp, 10
'006c663c    8bdc                    mov ebx, esp
'006c663e    b90a000000              mov ecx, 0000000a
'006c6643    890b                    mov dword ptr [ebx], ecx
'006c6645    89bd54ffffff            mov dword ptr [ebp+ffffff54], edi
'006c664b    898d54ffffff            mov dword ptr [ebp+ffffff54], ecx
'006c6651    8b8d48ffffff            mov ecx, dword ptr [ebp+ffffff48]
'006c6657    894b04                  mov dword ptr [ebx+04], ecx
'006c665a    8b154cb07200            mov edx, dword ptr [0072b04c]
'006c6660    b804000280              mov eax, 80020004
'006c6665    894308                  mov dword ptr [ebx+08], eax
'006c6668    89855cffffff            mov dword ptr [ebp+ffffff5c], eax
'006c666e    8b8550ffffff            mov eax, dword ptr [ebp+ffffff50]
'006c6674    89430c                  mov dword ptr [ebx+0c], eax
'006c6677    8b8554ffffff            mov eax, dword ptr [ebp+ffffff54]
'006c667d    83ec10                  sub esp, 10
'006c6680    8bcc                    mov ecx, esp
'006c6682    8901                    mov dword ptr [ecx], eax
'006c6684    8b8558ffffff            mov eax, dword ptr [ebp+ffffff58]
'006c668a    894104                  mov dword ptr [ecx+04], eax
'006c668d    8b855cffffff            mov eax, dword ptr [ebp+ffffff5c]
'006c6693    894108                  mov dword ptr [ecx+08], eax
'006c6696    8b8560ffffff            mov eax, dword ptr [ebp+ffffff60]
'006c669c    89410c                  mov dword ptr [ecx+0c], eax
'006c669f    83ec10                  sub esp, 10
'006c66a2    89bd64ffffff            mov dword ptr [ebp+ffffff64], edi
'006c66a8    c78564ffffff03000000    mov dword ptr [ebp+ffffff64], 00000003
'006c66b2    8b8564ffffff            mov eax, dword ptr [ebp+ffffff64]
'006c66b8    8bcc                    mov ecx, esp
'006c66ba    8901                    mov dword ptr [ecx], eax
'006c66bc    8b8568ffffff            mov eax, dword ptr [ebp+ffffff68]
'006c66c2    894104                  mov dword ptr [ecx+04], eax
'006c66c5    c7856cffffff01000000    mov dword ptr [ebp+ffffff6c], 00000001
'006c66cf    8b856cffffff            mov eax, dword ptr [ebp+ffffff6c]
'006c66d5    894108                  mov dword ptr [ecx+08], eax
'006c66d8    8b8570ffffff            mov eax, dword ptr [ebp+ffffff70]
'006c66de    89410c                  mov dword ptr [ecx+0c], eax
'006c66e1    8b0d4cb07200            mov ecx, dword ptr [0072b04c]
'006c66e7    6864354300              push 00433564
'006c66ec    897de8                  mov dword ptr [ebp-18], edi
'006c66ef    897de4                  mov dword ptr [ebp-1c], edi
'006c66f2    897dd4                  mov dword ptr [ebp-2c], edi
'006c66f5    897dc4                  mov dword ptr [ebp-3c], edi
'006c66f8    897db4                  mov dword ptr [ebp-4c], edi
'006c66fb    897da4                  mov dword ptr [ebp-5c], edi
'006c66fe    897d94                  mov dword ptr [ebp-6c], edi
'006c6701    897d84                  mov dword ptr [ebp-7c], edi
'006c6704    89bd74ffffff            mov dword ptr [ebp+ffffff74], edi
'006c670a    89bdb0feffff            mov dword ptr [ebp+fffffeb0], edi
'006c6710    8b12                    mov edx, dword ptr [edx]
'006c6712    51                      push ecx
'006c6713    ff92bc000000            call dword ptr [edx+000000bc]
'006c6719    dbe2                    fnclex
'006c671b    3bc7                    cmp eax, edi
'006c671d    7d18                    jge 6c6737

If (0 < 0) Then
'006c671f    8b154cb07200            mov edx, dword ptr [0072b04c]
'006c6725    68bc000000              push 000000bc
'006c672a    68301f4300              push 00431f30
'006c672f    52                      push edx
'006c6730    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c6731    ff1580104000            call dword ptr [00401080]
    
End If
'006c6737    8b45e4                  mov eax, dword ptr [ebp-1c]

' *** Reference to "__vbaObjSet"
'006c673a    8b1db4104000            mov ebx, dword ptr [004010b4]
'006c6740    50                      push eax
'006c6741    8d45e8                  lea eax, dword ptr [ebp-18]
'006c6744    50                      push eax
'006c6745    897de4                  mov dword ptr [ebp-1c], edi
'006c6748    ffd3                    call ebx
Set var_41 = Nothing
'006c674a    8b45e8                  mov eax, dword ptr [ebp-18]
'006c674d    8b08                    mov ecx, dword ptr [eax]
'006c674f    6830a84300              push 0043a830
'006c6754    50                      push eax
'006c6755    ff5144                  call dword ptr [ecx+44]
'006c6758    dbe2                    fnclex
'006c675a    3bc7                    cmp eax, edi
'006c675c    7d12                    jge 6c6770

If (var_41 < 0) Then
'006c675e    8b55e8                  mov edx, dword ptr [ebp-18]
'006c6761    6a44                    push 44
'006c6763    6830314300              push 00433130
'006c6768    52                      push edx
'006c6769    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c676a    ff1580104000            call dword ptr [00401080]
    
End If
'006c6770    8b06                    mov eax, dword ptr [esi]
'006c6772    56                      push esi
'006c6773    ff9094030000            call dword ptr [eax+00000394]
'006c6779    50                      push eax
'006c677a    8d4de4                  lea ecx, dword ptr [ebp-1c]
'006c677d    51                      push ecx
'006c677e    ffd3                    call ebx
Set var_40 = Nothing
'006c6780    8b5de4                  mov ebx, dword ptr [ebp-1c]
'006c6783    83ec10                  sub esp, 10
'006c6786    895ddc                  mov dword ptr [ebp-24], ebx
'006c6789    8bdc                    mov ebx, esp
'006c678b    b90a000000              mov ecx, 0000000a
'006c6790    890b                    mov dword ptr [ebx], ecx
'006c6792    898d54ffffff            mov dword ptr [ebp+ffffff54], ecx
'006c6798    898d64ffffff            mov dword ptr [ebp+ffffff64], ecx
'006c679e    8b8db8feffff            mov ecx, dword ptr [ebp+fffffeb8]
'006c67a4    894b04                  mov dword ptr [ebx+04], ecx
'006c67a7    b804000280              mov eax, 80020004
'006c67ac    894308                  mov dword ptr [ebx+08], eax
'006c67af    8bd0                    mov edx, eax
'006c67b1    89855cffffff            mov dword ptr [ebp+ffffff5c], eax
'006c67b7    89856cffffff            mov dword ptr [ebp+ffffff6c], eax
'006c67bd    8b85c0feffff            mov eax, dword ptr [ebp+fffffec0]
'006c67c3    89430c                  mov dword ptr [ebx+0c], eax
'006c67c6    83ec10                  sub esp, 10
'006c67c9    8bcc                    mov ecx, esp
'006c67cb    b80a000000              mov eax, 0000000a
'006c67d0    8901                    mov dword ptr [ecx], eax
'006c67d2    8b85c8feffff            mov eax, dword ptr [ebp+fffffec8]
'006c67d8    894104                  mov dword ptr [ecx+04], eax
'006c67db    895108                  mov dword ptr [ecx+08], edx
'006c67de    8b95d0feffff            mov edx, dword ptr [ebp+fffffed0]
'006c67e4    89510c                  mov dword ptr [ecx+0c], edx
'006c67e7    8b95d8feffff            mov edx, dword ptr [ebp+fffffed8]
'006c67ed    83ec10                  sub esp, 10
'006c67f0    8bcc                    mov ecx, esp
'006c67f2    b80a000000              mov eax, 0000000a
'006c67f7    8901                    mov dword ptr [ecx], eax
'006c67f9    895104                  mov dword ptr [ecx+04], edx
'006c67fc    8b95e8feffff            mov edx, dword ptr [ebp+fffffee8]
'006c6802    b804000280              mov eax, 80020004
'006c6807    894108                  mov dword ptr [ecx+08], eax
'006c680a    8b85e0feffff            mov eax, dword ptr [ebp+fffffee0]
'006c6810    89410c                  mov dword ptr [ecx+0c], eax
'006c6813    83ec10                  sub esp, 10
'006c6816    8bcc                    mov ecx, esp
'006c6818    b80a000000              mov eax, 0000000a
'006c681d    8901                    mov dword ptr [ecx], eax
'006c681f    895104                  mov dword ptr [ecx+04], edx
'006c6822    8b95f8feffff            mov edx, dword ptr [ebp+fffffef8]
'006c6828    b804000280              mov eax, 80020004
'006c682d    894108                  mov dword ptr [ecx+08], eax
'006c6830    8b85f0feffff            mov eax, dword ptr [ebp+fffffef0]
'006c6836    89410c                  mov dword ptr [ecx+0c], eax
'006c6839    83ec10                  sub esp, 10
'006c683c    8bcc                    mov ecx, esp
'006c683e    b80a000000              mov eax, 0000000a
'006c6843    8901                    mov dword ptr [ecx], eax
'006c6845    895104                  mov dword ptr [ecx+04], edx
'006c6848    8b9508ffffff            mov edx, dword ptr [ebp+ffffff08]
'006c684e    b804000280              mov eax, 80020004
'006c6853    894108                  mov dword ptr [ecx+08], eax
'006c6856    8b8500ffffff            mov eax, dword ptr [ebp+ffffff00]
'006c685c    89410c                  mov dword ptr [ecx+0c], eax
'006c685f    83ec10                  sub esp, 10
'006c6862    8bcc                    mov ecx, esp
'006c6864    b80a000000              mov eax, 0000000a
'006c6869    8901                    mov dword ptr [ecx], eax
'006c686b    895104                  mov dword ptr [ecx+04], edx
'006c686e    b804000280              mov eax, 80020004
'006c6873    894108                  mov dword ptr [ecx+08], eax
'006c6876    8b8510ffffff            mov eax, dword ptr [ebp+ffffff10]
'006c687c    89410c                  mov dword ptr [ecx+0c], eax
'006c687f    897de4                  mov dword ptr [ebp-1c], edi
'006c6882    8b7de8                  mov edi, dword ptr [ebp-18]
'006c6885    83ec10                  sub esp, 10
'006c6888    8bcc                    mov ecx, esp
'006c688a    b80a000000              mov eax, 0000000a
'006c688f    c745d409000000          mov dword ptr [ebp-2c], 00000009
'006c6896    8b3f                    mov edi, dword ptr [edi]
'006c6898    8901                    mov dword ptr [ecx], eax
'006c689a    8b9518ffffff            mov edx, dword ptr [ebp+ffffff18]
'006c68a0    895104                  mov dword ptr [ecx+04], edx
'006c68a3    8b9528ffffff            mov edx, dword ptr [ebp+ffffff28]
'006c68a9    83ec10                  sub esp, 10
'006c68ac    b804000280              mov eax, 80020004
'006c68b1    894108                  mov dword ptr [ecx+08], eax
'006c68b4    8b8520ffffff            mov eax, dword ptr [ebp+ffffff20]
'006c68ba    89410c                  mov dword ptr [ecx+0c], eax
'006c68bd    8bcc                    mov ecx, esp
'006c68bf    b80a000000              mov eax, 0000000a
'006c68c4    8901                    mov dword ptr [ecx], eax
'006c68c6    895104                  mov dword ptr [ecx+04], edx
'006c68c9    8b9538ffffff            mov edx, dword ptr [ebp+ffffff38]
'006c68cf    83ec10                  sub esp, 10
'006c68d2    b804000280              mov eax, 80020004
'006c68d7    894108                  mov dword ptr [ecx+08], eax
'006c68da    8b8530ffffff            mov eax, dword ptr [ebp+ffffff30]
'006c68e0    89410c                  mov dword ptr [ecx+0c], eax
'006c68e3    8bcc                    mov ecx, esp
'006c68e5    b80a000000              mov eax, 0000000a
'006c68ea    8901                    mov dword ptr [ecx], eax
'006c68ec    895104                  mov dword ptr [ecx+04], edx
'006c68ef    8b9548ffffff            mov edx, dword ptr [ebp+ffffff48]
'006c68f5    83ec10                  sub esp, 10
'006c68f8    b804000280              mov eax, 80020004
'006c68fd    894108                  mov dword ptr [ecx+08], eax
'006c6900    8b8540ffffff            mov eax, dword ptr [ebp+ffffff40]
'006c6906    89410c                  mov dword ptr [ecx+0c], eax
'006c6909    8bcc                    mov ecx, esp
'006c690b    b80a000000              mov eax, 0000000a
'006c6910    8901                    mov dword ptr [ecx], eax
'006c6912    895104                  mov dword ptr [ecx+04], edx
'006c6915    8b9554ffffff            mov edx, dword ptr [ebp+ffffff54]
'006c691b    b804000280              mov eax, 80020004
'006c6920    894108                  mov dword ptr [ecx+08], eax
'006c6923    8b8550ffffff            mov eax, dword ptr [ebp+ffffff50]
'006c6929    89410c                  mov dword ptr [ecx+0c], eax
'006c692c    8b8558ffffff            mov eax, dword ptr [ebp+ffffff58]
'006c6932    83ec10                  sub esp, 10
'006c6935    8bcc                    mov ecx, esp
'006c6937    8911                    mov dword ptr [ecx], edx
'006c6939    8b955cffffff            mov edx, dword ptr [ebp+ffffff5c]
'006c693f    894104                  mov dword ptr [ecx+04], eax
'006c6942    8b8560ffffff            mov eax, dword ptr [ebp+ffffff60]
'006c6948    895108                  mov dword ptr [ecx+08], edx
'006c694b    8b9564ffffff            mov edx, dword ptr [ebp+ffffff64]
'006c6951    89410c                  mov dword ptr [ecx+0c], eax
'006c6954    8b8568ffffff            mov eax, dword ptr [ebp+ffffff68]
'006c695a    83ec10                  sub esp, 10
'006c695d    8bcc                    mov ecx, esp
'006c695f    8911                    mov dword ptr [ecx], edx
'006c6961    8b956cffffff            mov edx, dword ptr [ebp+ffffff6c]
'006c6967    894104                  mov dword ptr [ecx+04], eax
'006c696a    8b8570ffffff            mov eax, dword ptr [ebp+ffffff70]
'006c6970    895108                  mov dword ptr [ecx+08], edx
'006c6973    8b55d4                  mov edx, dword ptr [ebp-2c]
'006c6976    89410c                  mov dword ptr [ecx+0c], eax
'006c6979    8b45d8                  mov eax, dword ptr [ebp-28]
'006c697c    83ec10                  sub esp, 10
'006c697f    8bcc                    mov ecx, esp
'006c6981    8911                    mov dword ptr [ecx], edx
'006c6983    8b55dc                  mov edx, dword ptr [ebp-24]
'006c6986    894104                  mov dword ptr [ecx+04], eax
'006c6989    8b45e0                  mov eax, dword ptr [ebp-20]
'006c698c    895108                  mov dword ptr [ecx+08], edx
'006c698f    89410c                  mov dword ptr [ecx+0c], eax
'006c6992    8b4de8                  mov ecx, dword ptr [ebp-18]
'006c6995    68209e4300              push 00439e20
'006c699a    51                      push ecx
'006c699b    ff97f4000000            call dword ptr [edi+000000f4]
'006c69a1    dbe2                    fnclex
'006c69a3    85c0                    test ax, ax
'006c69a5    7d19                    jge 6c69c0
'006c69a7    8b55e8                  mov edx, dword ptr [ebp-18]

' *** Reference to "__vbaHresultCheckObj"
'006c69aa    8b1d80104000            mov ebx, dword ptr [00401080]
'006c69b0    68f4000000              push 000000f4
'006c69b5    6830314300              push 00433130
'006c69ba    52                      push edx
'006c69bb    50                      push eax
'006c69bc    ffd3                    call ebx
'006c69be    eb06                    jmp 6c69c6

' *** Reference to "__vbaHresultCheckObj"
'006c69c0    8b1d80104000            mov ebx, dword ptr [00401080]
'006c69c6    8d4de4                  lea ecx, dword ptr [ebp-1c]

' *** Reference to "__vbaFreeObj"
'006c69c9    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_40)

' *** Reference to "__vbaFreeVar"
'006c69cf    8b3d28104000            mov edi, dword ptr [00401028]
'006c69d5    8d4dd4                  lea ecx, dword ptr [ebp-2c]
'006c69d8    ffd7                    call edi
'#Cleanup(var_44)
'006c69da    8b45e8                  mov eax, dword ptr [ebp-18]
'006c69dd    8b08                    mov ecx, dword ptr [eax]
'006c69df    8d95b0feffff            lea edx, dword ptr [ebp+fffffeb0]
'006c69e5    52                      push edx
'006c69e6    50                      push eax
'006c69e7    ff515c                  call dword ptr [ecx+5c]
'006c69ea    dbe2                    fnclex
'006c69ec    85c0                    test ax, ax
'006c69ee    7d0e                    jge 6c69fe
'006c69f0    8b4de8                  mov ecx, dword ptr [ebp-18]
'006c69f3    6a5c                    push 5c
'006c69f5    6830314300              push 00433130
'006c69fa    51                      push ecx
'006c69fb    50                      push eax
'006c69fc    ffd3                    call ebx
'006c69fe    6683bdb0feffff00        cmp word ptr [ebp+fffffeb0], 00
'006c6a06    0f85d5080000            jne 6c72e1

If (0 = 0) Then
'006c6a0c    8b45e8                  mov eax, dword ptr [ebp-18]
'006c6a0f    8b10                    mov edx, dword ptr [eax]
'006c6a11    50                      push eax
'006c6a12    ff92d0000000            call dword ptr [edx+000000d0]
'006c6a18    dbe2                    fnclex
'006c6a1a    85c0                    test ax, ax
'006c6a1c    7d11                    jge 6c6a2f
'006c6a1e    8b4de8                  mov ecx, dword ptr [ebp-18]
'006c6a21    68d0000000              push 000000d0
'006c6a26    6830314300              push 00433130
'006c6a2b    51                      push ecx
'006c6a2c    50                      push eax
'006c6a2d    ffd3                    call ebx
'006c6a2f    8b16                    mov edx, dword ptr [esi]
'006c6a31    56                      push esi
'006c6a32    ff922c030000            call dword ptr [edx+0000032c]
'006c6a38    8b55e8                  mov edx, dword ptr [ebp-18]
'006c6a3b    83ec10                  sub esp, 10
'006c6a3e    8bdc                    mov ebx, esp
'006c6a40    b909000000              mov ecx, 00000009
'006c6a45    890b                    mov dword ptr [ebx], ecx
'006c6a47    894dd4                  mov dword ptr [ebp-2c], ecx
'006c6a4a    8b4dd8                  mov ecx, dword ptr [ebp-28]
'006c6a4d    894b04                  mov dword ptr [ebx+04], ecx
'006c6a50    894308                  mov dword ptr [ebx+08], eax
'006c6a53    8945dc                  mov dword ptr [ebp-24], eax
'006c6a56    8b45e0                  mov eax, dword ptr [ebp-20]
'006c6a59    89430c                  mov dword ptr [ebx+0c], eax
'006c6a5c    83ec10                  sub esp, 10
'006c6a5f    c78564ffffff08000000    mov dword ptr [ebp+ffffff64], 00000008
'006c6a69    8b8564ffffff            mov eax, dword ptr [ebp+ffffff64]
'006c6a6f    8bcc                    mov ecx, esp
'006c6a71    8901                    mov dword ptr [ecx], eax
'006c6a73    8b8568ffffff            mov eax, dword ptr [ebp+ffffff68]
'006c6a79    894104                  mov dword ptr [ecx+04], eax
'006c6a7c    c7856cffffff10f34300    mov dword ptr [ebp+ffffff6c], 0043f310
'006c6a86    8b856cffffff            mov eax, dword ptr [ebp+ffffff6c]
'006c6a8c    8b12                    mov edx, dword ptr [edx]
'006c6a8e    894108                  mov dword ptr [ecx+08], eax
'006c6a91    8b8570ffffff            mov eax, dword ptr [ebp+ffffff70]
'006c6a97    89410c                  mov dword ptr [ecx+0c], eax
'006c6a9a    8b4de8                  mov ecx, dword ptr [ebp-18]
'006c6a9d    51                      push ecx
'006c6a9e    ff9228010000            call dword ptr [edx+00000128]
'006c6aa4    dbe2                    fnclex
'006c6aa6    85c0                    test ax, ax
'006c6aa8    7d15                    jge 6c6abf
'006c6aaa    8b55e8                  mov edx, dword ptr [ebp-18]
'006c6aad    6828010000              push 00000128
'006c6ab2    6830314300              push 00433130
'006c6ab7    52                      push edx
'006c6ab8    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c6ab9    ff1580104000            call dword ptr [00401080]
'006c6abf    8d4dd4                  lea ecx, dword ptr [ebp-2c]
'006c6ac2    ffd7                    call edi
    '#Cleanup(var_44)
'006c6ac4    8b06                    mov eax, dword ptr [esi]
'006c6ac6    56                      push esi
'006c6ac7    ff9030030000            call dword ptr [eax+00000330]
'006c6acd    8b55e8                  mov edx, dword ptr [ebp-18]
'006c6ad0    83ec10                  sub esp, 10
'006c6ad3    8bdc                    mov ebx, esp
'006c6ad5    b909000000              mov ecx, 00000009
'006c6ada    890b                    mov dword ptr [ebx], ecx
'006c6adc    894dd4                  mov dword ptr [ebp-2c], ecx
'006c6adf    8b4dd8                  mov ecx, dword ptr [ebp-28]
'006c6ae2    894b04                  mov dword ptr [ebx+04], ecx
'006c6ae5    894308                  mov dword ptr [ebx+08], eax
'006c6ae8    8945dc                  mov dword ptr [ebp-24], eax
'006c6aeb    8b45e0                  mov eax, dword ptr [ebp-20]
'006c6aee    89430c                  mov dword ptr [ebx+0c], eax
'006c6af1    83ec10                  sub esp, 10
'006c6af4    c78564ffffff08000000    mov dword ptr [ebp+ffffff64], 00000008
'006c6afe    8b8564ffffff            mov eax, dword ptr [ebp+ffffff64]
'006c6b04    8bcc                    mov ecx, esp
'006c6b06    8901                    mov dword ptr [ecx], eax
'006c6b08    8b8568ffffff            mov eax, dword ptr [ebp+ffffff68]
'006c6b0e    894104                  mov dword ptr [ecx+04], eax
'006c6b11    c7856cffffff20f34300    mov dword ptr [ebp+ffffff6c], 0043f320
'006c6b1b    8b856cffffff            mov eax, dword ptr [ebp+ffffff6c]
'006c6b21    8b12                    mov edx, dword ptr [edx]
'006c6b23    894108                  mov dword ptr [ecx+08], eax
'006c6b26    8b8570ffffff            mov eax, dword ptr [ebp+ffffff70]
'006c6b2c    89410c                  mov dword ptr [ecx+0c], eax
'006c6b2f    8b4de8                  mov ecx, dword ptr [ebp-18]
'006c6b32    51                      push ecx
'006c6b33    ff9228010000            call dword ptr [edx+00000128]
'006c6b39    dbe2                    fnclex
'006c6b3b    85c0                    test ax, ax
'006c6b3d    7d15                    jge 6c6b54
'006c6b3f    8b55e8                  mov edx, dword ptr [ebp-18]
'006c6b42    6828010000              push 00000128
'006c6b47    6830314300              push 00433130
'006c6b4c    52                      push edx
'006c6b4d    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c6b4e    ff1580104000            call dword ptr [00401080]
'006c6b54    8d4dd4                  lea ecx, dword ptr [ebp-2c]
'006c6b57    ffd7                    call edi
    '#Cleanup(var_44)
'006c6b59    8b06                    mov eax, dword ptr [esi]
'006c6b5b    56                      push esi
'006c6b5c    ff9034030000            call dword ptr [eax+00000334]
'006c6b62    8b55e8                  mov edx, dword ptr [ebp-18]
'006c6b65    83ec10                  sub esp, 10
'006c6b68    8bdc                    mov ebx, esp
'006c6b6a    b909000000              mov ecx, 00000009
'006c6b6f    890b                    mov dword ptr [ebx], ecx
'006c6b71    894dd4                  mov dword ptr [ebp-2c], ecx
'006c6b74    8b4dd8                  mov ecx, dword ptr [ebp-28]
'006c6b77    894b04                  mov dword ptr [ebx+04], ecx
'006c6b7a    894308                  mov dword ptr [ebx+08], eax
'006c6b7d    8945dc                  mov dword ptr [ebp-24], eax
'006c6b80    8b45e0                  mov eax, dword ptr [ebp-20]
'006c6b83    89430c                  mov dword ptr [ebx+0c], eax
'006c6b86    83ec10                  sub esp, 10
'006c6b89    c78564ffffff08000000    mov dword ptr [ebp+ffffff64], 00000008
'006c6b93    8b8564ffffff            mov eax, dword ptr [ebp+ffffff64]
'006c6b99    8bcc                    mov ecx, esp
'006c6b9b    8901                    mov dword ptr [ecx], eax
'006c6b9d    8b8568ffffff            mov eax, dword ptr [ebp+ffffff68]
'006c6ba3    894104                  mov dword ptr [ecx+04], eax
'006c6ba6    c7856cffffff001b4400    mov dword ptr [ebp+ffffff6c], 00441b00
'006c6bb0    8b856cffffff            mov eax, dword ptr [ebp+ffffff6c]
'006c6bb6    8b12                    mov edx, dword ptr [edx]
'006c6bb8    894108                  mov dword ptr [ecx+08], eax
'006c6bbb    8b8570ffffff            mov eax, dword ptr [ebp+ffffff70]
'006c6bc1    89410c                  mov dword ptr [ecx+0c], eax
'006c6bc4    8b4de8                  mov ecx, dword ptr [ebp-18]
'006c6bc7    51                      push ecx
'006c6bc8    ff9228010000            call dword ptr [edx+00000128]
'006c6bce    dbe2                    fnclex
'006c6bd0    85c0                    test ax, ax
'006c6bd2    7d15                    jge 6c6be9
'006c6bd4    8b55e8                  mov edx, dword ptr [ebp-18]
'006c6bd7    6828010000              push 00000128
'006c6bdc    6830314300              push 00433130
'006c6be1    52                      push edx
'006c6be2    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c6be3    ff1580104000            call dword ptr [00401080]
'006c6be9    8d4dd4                  lea ecx, dword ptr [ebp-2c]
'006c6bec    ffd7                    call edi
    '#Cleanup(var_44)
'006c6bee    8b06                    mov eax, dword ptr [esi]
'006c6bf0    56                      push esi
'006c6bf1    ff9038030000            call dword ptr [eax+00000338]
'006c6bf7    8b55e8                  mov edx, dword ptr [ebp-18]
'006c6bfa    83ec10                  sub esp, 10
'006c6bfd    8bdc                    mov ebx, esp
'006c6bff    b909000000              mov ecx, 00000009
'006c6c04    890b                    mov dword ptr [ebx], ecx
'006c6c06    894dd4                  mov dword ptr [ebp-2c], ecx
'006c6c09    8b4dd8                  mov ecx, dword ptr [ebp-28]
'006c6c0c    894b04                  mov dword ptr [ebx+04], ecx
'006c6c0f    894308                  mov dword ptr [ebx+08], eax
'006c6c12    8945dc                  mov dword ptr [ebp-24], eax
'006c6c15    8b45e0                  mov eax, dword ptr [ebp-20]
'006c6c18    89430c                  mov dword ptr [ebx+0c], eax
'006c6c1b    83ec10                  sub esp, 10
'006c6c1e    c78564ffffff08000000    mov dword ptr [ebp+ffffff64], 00000008
'006c6c28    8b8564ffffff            mov eax, dword ptr [ebp+ffffff64]
'006c6c2e    8bcc                    mov ecx, esp
'006c6c30    8901                    mov dword ptr [ecx], eax
'006c6c32    8b8568ffffff            mov eax, dword ptr [ebp+ffffff68]
'006c6c38    894104                  mov dword ptr [ecx+04], eax
'006c6c3b    c7856cffffff401b4400    mov dword ptr [ebp+ffffff6c], 00441b40
'006c6c45    8b856cffffff            mov eax, dword ptr [ebp+ffffff6c]
'006c6c4b    8b12                    mov edx, dword ptr [edx]
'006c6c4d    894108                  mov dword ptr [ecx+08], eax
'006c6c50    8b8570ffffff            mov eax, dword ptr [ebp+ffffff70]
'006c6c56    89410c                  mov dword ptr [ecx+0c], eax
'006c6c59    8b4de8                  mov ecx, dword ptr [ebp-18]
'006c6c5c    51                      push ecx
'006c6c5d    ff9228010000            call dword ptr [edx+00000128]
'006c6c63    dbe2                    fnclex
'006c6c65    85c0                    test ax, ax
'006c6c67    7d15                    jge 6c6c7e
'006c6c69    8b55e8                  mov edx, dword ptr [ebp-18]
'006c6c6c    6828010000              push 00000128
'006c6c71    6830314300              push 00433130
'006c6c76    52                      push edx
'006c6c77    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c6c78    ff1580104000            call dword ptr [00401080]
'006c6c7e    8d4dd4                  lea ecx, dword ptr [ebp-2c]
'006c6c81    ffd7                    call edi
    '#Cleanup(var_44)
'006c6c83    8b06                    mov eax, dword ptr [esi]
'006c6c85    56                      push esi
'006c6c86    ff903c030000            call dword ptr [eax+0000033c]
'006c6c8c    8b55e8                  mov edx, dword ptr [ebp-18]
'006c6c8f    83ec10                  sub esp, 10
'006c6c92    8bdc                    mov ebx, esp
'006c6c94    b909000000              mov ecx, 00000009
'006c6c99    890b                    mov dword ptr [ebx], ecx
'006c6c9b    894dd4                  mov dword ptr [ebp-2c], ecx
'006c6c9e    8b4dd8                  mov ecx, dword ptr [ebp-28]
'006c6ca1    894b04                  mov dword ptr [ebx+04], ecx
'006c6ca4    894308                  mov dword ptr [ebx+08], eax
'006c6ca7    8945dc                  mov dword ptr [ebp-24], eax
'006c6caa    8b45e0                  mov eax, dword ptr [ebp-20]
'006c6cad    89430c                  mov dword ptr [ebx+0c], eax
'006c6cb0    83ec10                  sub esp, 10
'006c6cb3    c78564ffffff08000000    mov dword ptr [ebp+ffffff64], 00000008
'006c6cbd    8b8564ffffff            mov eax, dword ptr [ebp+ffffff64]
'006c6cc3    8bcc                    mov ecx, esp
'006c6cc5    8901                    mov dword ptr [ecx], eax
'006c6cc7    8b8568ffffff            mov eax, dword ptr [ebp+ffffff68]
'006c6ccd    894104                  mov dword ptr [ecx+04], eax
'006c6cd0    c7856cffffff54f54300    mov dword ptr [ebp+ffffff6c], 0043f554
'006c6cda    8b856cffffff            mov eax, dword ptr [ebp+ffffff6c]
'006c6ce0    8b12                    mov edx, dword ptr [edx]
'006c6ce2    894108                  mov dword ptr [ecx+08], eax
'006c6ce5    8b8570ffffff            mov eax, dword ptr [ebp+ffffff70]
'006c6ceb    89410c                  mov dword ptr [ecx+0c], eax
'006c6cee    8b4de8                  mov ecx, dword ptr [ebp-18]
'006c6cf1    51                      push ecx
'006c6cf2    ff9228010000            call dword ptr [edx+00000128]
'006c6cf8    dbe2                    fnclex
'006c6cfa    85c0                    test ax, ax
'006c6cfc    7d15                    jge 6c6d13
'006c6cfe    8b55e8                  mov edx, dword ptr [ebp-18]
'006c6d01    6828010000              push 00000128
'006c6d06    6830314300              push 00433130
'006c6d0b    52                      push edx
'006c6d0c    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c6d0d    ff1580104000            call dword ptr [00401080]
'006c6d13    8d4dd4                  lea ecx, dword ptr [ebp-2c]
'006c6d16    ffd7                    call edi
    '#Cleanup(var_44)
'006c6d18    8b06                    mov eax, dword ptr [esi]
'006c6d1a    56                      push esi
'006c6d1b    ff9040030000            call dword ptr [eax+00000340]
'006c6d21    8b55e8                  mov edx, dword ptr [ebp-18]
'006c6d24    83ec10                  sub esp, 10
'006c6d27    8bdc                    mov ebx, esp
'006c6d29    b909000000              mov ecx, 00000009
'006c6d2e    890b                    mov dword ptr [ebx], ecx
'006c6d30    894dd4                  mov dword ptr [ebp-2c], ecx
'006c6d33    8b4dd8                  mov ecx, dword ptr [ebp-28]
'006c6d36    894b04                  mov dword ptr [ebx+04], ecx
'006c6d39    894308                  mov dword ptr [ebx+08], eax
'006c6d3c    8945dc                  mov dword ptr [ebp-24], eax
'006c6d3f    8b45e0                  mov eax, dword ptr [ebp-20]
'006c6d42    89430c                  mov dword ptr [ebx+0c], eax
'006c6d45    83ec10                  sub esp, 10
'006c6d48    c78564ffffff08000000    mov dword ptr [ebp+ffffff64], 00000008
'006c6d52    8b8564ffffff            mov eax, dword ptr [ebp+ffffff64]
'006c6d58    8bcc                    mov ecx, esp
'006c6d5a    8901                    mov dword ptr [ecx], eax
'006c6d5c    8b8568ffffff            mov eax, dword ptr [ebp+ffffff68]
'006c6d62    894104                  mov dword ptr [ecx+04], eax
'006c6d65    c7856cffffffb41b4400    mov dword ptr [ebp+ffffff6c], 00441bb4
'006c6d6f    8b856cffffff            mov eax, dword ptr [ebp+ffffff6c]
'006c6d75    8b12                    mov edx, dword ptr [edx]
'006c6d77    894108                  mov dword ptr [ecx+08], eax
'006c6d7a    8b8570ffffff            mov eax, dword ptr [ebp+ffffff70]
'006c6d80    89410c                  mov dword ptr [ecx+0c], eax
'006c6d83    8b4de8                  mov ecx, dword ptr [ebp-18]
'006c6d86    51                      push ecx
'006c6d87    ff9228010000            call dword ptr [edx+00000128]
'006c6d8d    dbe2                    fnclex
'006c6d8f    85c0                    test ax, ax
'006c6d91    7d15                    jge 6c6da8
'006c6d93    8b55e8                  mov edx, dword ptr [ebp-18]
'006c6d96    6828010000              push 00000128
'006c6d9b    6830314300              push 00433130
'006c6da0    52                      push edx
'006c6da1    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c6da2    ff1580104000            call dword ptr [00401080]
'006c6da8    8d4dd4                  lea ecx, dword ptr [ebp-2c]
'006c6dab    ffd7                    call edi
    '#Cleanup(var_44)
'006c6dad    8b06                    mov eax, dword ptr [esi]
'006c6daf    56                      push esi
'006c6db0    ff9044030000            call dword ptr [eax+00000344]
'006c6db6    8b55e8                  mov edx, dword ptr [ebp-18]
'006c6db9    83ec10                  sub esp, 10
'006c6dbc    8bdc                    mov ebx, esp
'006c6dbe    b909000000              mov ecx, 00000009
'006c6dc3    890b                    mov dword ptr [ebx], ecx
'006c6dc5    894dd4                  mov dword ptr [ebp-2c], ecx
'006c6dc8    8b4dd8                  mov ecx, dword ptr [ebp-28]
'006c6dcb    894b04                  mov dword ptr [ebx+04], ecx
'006c6dce    894308                  mov dword ptr [ebx+08], eax
'006c6dd1    8945dc                  mov dword ptr [ebp-24], eax
'006c6dd4    8b45e0                  mov eax, dword ptr [ebp-20]
'006c6dd7    89430c                  mov dword ptr [ebx+0c], eax
'006c6dda    83ec10                  sub esp, 10
'006c6ddd    c78564ffffff08000000    mov dword ptr [ebp+ffffff64], 00000008
'006c6de7    8b8564ffffff            mov eax, dword ptr [ebp+ffffff64]
'006c6ded    8bcc                    mov ecx, esp
'006c6def    8901                    mov dword ptr [ecx], eax
'006c6df1    8b8568ffffff            mov eax, dword ptr [ebp+ffffff68]
'006c6df7    894104                  mov dword ptr [ecx+04], eax
'006c6dfa    c7856cffffff74f34300    mov dword ptr [ebp+ffffff6c], 0043f374
'006c6e04    8b856cffffff            mov eax, dword ptr [ebp+ffffff6c]
'006c6e0a    8b12                    mov edx, dword ptr [edx]
'006c6e0c    894108                  mov dword ptr [ecx+08], eax
'006c6e0f    8b8570ffffff            mov eax, dword ptr [ebp+ffffff70]
'006c6e15    89410c                  mov dword ptr [ecx+0c], eax
'006c6e18    8b4de8                  mov ecx, dword ptr [ebp-18]
'006c6e1b    51                      push ecx
'006c6e1c    ff9228010000            call dword ptr [edx+00000128]
'006c6e22    dbe2                    fnclex
'006c6e24    85c0                    test ax, ax
'006c6e26    7d15                    jge 6c6e3d
'006c6e28    8b55e8                  mov edx, dword ptr [ebp-18]
'006c6e2b    6828010000              push 00000128
'006c6e30    6830314300              push 00433130
'006c6e35    52                      push edx
'006c6e36    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c6e37    ff1580104000            call dword ptr [00401080]
'006c6e3d    8d4dd4                  lea ecx, dword ptr [ebp-2c]
'006c6e40    ffd7                    call edi
    '#Cleanup(var_44)
'006c6e42    8b06                    mov eax, dword ptr [esi]
'006c6e44    56                      push esi
'006c6e45    ff9048030000            call dword ptr [eax+00000348]
'006c6e4b    8b55e8                  mov edx, dword ptr [ebp-18]
'006c6e4e    83ec10                  sub esp, 10
'006c6e51    8bdc                    mov ebx, esp
'006c6e53    b909000000              mov ecx, 00000009
'006c6e58    890b                    mov dword ptr [ebx], ecx
'006c6e5a    894dd4                  mov dword ptr [ebp-2c], ecx
'006c6e5d    8b4dd8                  mov ecx, dword ptr [ebp-28]
'006c6e60    894b04                  mov dword ptr [ebx+04], ecx
'006c6e63    894308                  mov dword ptr [ebx+08], eax
'006c6e66    8945dc                  mov dword ptr [ebp-24], eax
'006c6e69    8b45e0                  mov eax, dword ptr [ebp-20]
'006c6e6c    89430c                  mov dword ptr [ebx+0c], eax
'006c6e6f    83ec10                  sub esp, 10
'006c6e72    c78564ffffff08000000    mov dword ptr [ebp+ffffff64], 00000008
'006c6e7c    8b8564ffffff            mov eax, dword ptr [ebp+ffffff64]
'006c6e82    8bcc                    mov ecx, esp
'006c6e84    8901                    mov dword ptr [ecx], eax
'006c6e86    8b8568ffffff            mov eax, dword ptr [ebp+ffffff68]
'006c6e8c    894104                  mov dword ptr [ecx+04], eax
'006c6e8f    c7856cfffffff81b4400    mov dword ptr [ebp+ffffff6c], 00441bf8
'006c6e99    8b856cffffff            mov eax, dword ptr [ebp+ffffff6c]
'006c6e9f    8b12                    mov edx, dword ptr [edx]
'006c6ea1    894108                  mov dword ptr [ecx+08], eax
'006c6ea4    8b8570ffffff            mov eax, dword ptr [ebp+ffffff70]
'006c6eaa    89410c                  mov dword ptr [ecx+0c], eax
'006c6ead    8b4de8                  mov ecx, dword ptr [ebp-18]
'006c6eb0    51                      push ecx
'006c6eb1    ff9228010000            call dword ptr [edx+00000128]
'006c6eb7    dbe2                    fnclex
'006c6eb9    85c0                    test ax, ax
'006c6ebb    7d15                    jge 6c6ed2
'006c6ebd    8b55e8                  mov edx, dword ptr [ebp-18]
'006c6ec0    6828010000              push 00000128
'006c6ec5    6830314300              push 00433130
'006c6eca    52                      push edx
'006c6ecb    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c6ecc    ff1580104000            call dword ptr [00401080]
'006c6ed2    8d4dd4                  lea ecx, dword ptr [ebp-2c]
'006c6ed5    ffd7                    call edi
    '#Cleanup(var_44)
'006c6ed7    8b06                    mov eax, dword ptr [esi]
'006c6ed9    56                      push esi
'006c6eda    ff904c030000            call dword ptr [eax+0000034c]
'006c6ee0    8b55e8                  mov edx, dword ptr [ebp-18]
'006c6ee3    83ec10                  sub esp, 10
'006c6ee6    8bdc                    mov ebx, esp
'006c6ee8    b909000000              mov ecx, 00000009
'006c6eed    890b                    mov dword ptr [ebx], ecx
'006c6eef    894dd4                  mov dword ptr [ebp-2c], ecx
'006c6ef2    8b4dd8                  mov ecx, dword ptr [ebp-28]
'006c6ef5    894b04                  mov dword ptr [ebx+04], ecx
'006c6ef8    894308                  mov dword ptr [ebx+08], eax
'006c6efb    8945dc                  mov dword ptr [ebp-24], eax
'006c6efe    8b45e0                  mov eax, dword ptr [ebp-20]
'006c6f01    89430c                  mov dword ptr [ebx+0c], eax
'006c6f04    83ec10                  sub esp, 10
'006c6f07    c78564ffffff08000000    mov dword ptr [ebp+ffffff64], 00000008
'006c6f11    8b8564ffffff            mov eax, dword ptr [ebp+ffffff64]
'006c6f17    8bcc                    mov ecx, esp
'006c6f19    8901                    mov dword ptr [ecx], eax
'006c6f1b    8b8568ffffff            mov eax, dword ptr [ebp+ffffff68]
'006c6f21    894104                  mov dword ptr [ecx+04], eax
'006c6f24    c7856cffffff641c4400    mov dword ptr [ebp+ffffff6c], 00441c64
'006c6f2e    8b856cffffff            mov eax, dword ptr [ebp+ffffff6c]
'006c6f34    8b12                    mov edx, dword ptr [edx]
'006c6f36    894108                  mov dword ptr [ecx+08], eax
'006c6f39    8b8570ffffff            mov eax, dword ptr [ebp+ffffff70]
'006c6f3f    89410c                  mov dword ptr [ecx+0c], eax
'006c6f42    8b4de8                  mov ecx, dword ptr [ebp-18]
'006c6f45    51                      push ecx
'006c6f46    ff9228010000            call dword ptr [edx+00000128]
'006c6f4c    dbe2                    fnclex
'006c6f4e    85c0                    test ax, ax
'006c6f50    7d15                    jge 6c6f67
'006c6f52    8b55e8                  mov edx, dword ptr [ebp-18]
'006c6f55    6828010000              push 00000128
'006c6f5a    6830314300              push 00433130
'006c6f5f    52                      push edx
'006c6f60    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c6f61    ff1580104000            call dword ptr [00401080]
'006c6f67    8d4dd4                  lea ecx, dword ptr [ebp-2c]
'006c6f6a    ffd7                    call edi
    '#Cleanup(var_44)
'006c6f6c    8b06                    mov eax, dword ptr [esi]
'006c6f6e    56                      push esi
'006c6f6f    ff9050030000            call dword ptr [eax+00000350]
'006c6f75    8b55e8                  mov edx, dword ptr [ebp-18]
'006c6f78    83ec10                  sub esp, 10
'006c6f7b    8bdc                    mov ebx, esp
'006c6f7d    b909000000              mov ecx, 00000009
'006c6f82    890b                    mov dword ptr [ebx], ecx
'006c6f84    894dd4                  mov dword ptr [ebp-2c], ecx
'006c6f87    8b4dd8                  mov ecx, dword ptr [ebp-28]
'006c6f8a    894b04                  mov dword ptr [ebx+04], ecx
'006c6f8d    894308                  mov dword ptr [ebx+08], eax
'006c6f90    8945dc                  mov dword ptr [ebp-24], eax
'006c6f93    8b45e0                  mov eax, dword ptr [ebp-20]
'006c6f96    89430c                  mov dword ptr [ebx+0c], eax
'006c6f99    83ec10                  sub esp, 10
'006c6f9c    c78564ffffff08000000    mov dword ptr [ebp+ffffff64], 00000008
'006c6fa6    8b8564ffffff            mov eax, dword ptr [ebp+ffffff64]
'006c6fac    8bcc                    mov ecx, esp
'006c6fae    8901                    mov dword ptr [ecx], eax
'006c6fb0    8b8568ffffff            mov eax, dword ptr [ebp+ffffff68]
'006c6fb6    894104                  mov dword ptr [ecx+04], eax
'006c6fb9    c7856cffffff18f44200    mov dword ptr [ebp+ffffff6c], 0042f418
'006c6fc3    8b856cffffff            mov eax, dword ptr [ebp+ffffff6c]
'006c6fc9    8b12                    mov edx, dword ptr [edx]
'006c6fcb    894108                  mov dword ptr [ecx+08], eax
'006c6fce    8b8570ffffff            mov eax, dword ptr [ebp+ffffff70]
'006c6fd4    89410c                  mov dword ptr [ecx+0c], eax
'006c6fd7    8b4de8                  mov ecx, dword ptr [ebp-18]
'006c6fda    51                      push ecx
'006c6fdb    ff9228010000            call dword ptr [edx+00000128]
'006c6fe1    dbe2                    fnclex
'006c6fe3    85c0                    test ax, ax
'006c6fe5    7d15                    jge 6c6ffc
'006c6fe7    8b55e8                  mov edx, dword ptr [ebp-18]
'006c6fea    6828010000              push 00000128
'006c6fef    6830314300              push 00433130
'006c6ff4    52                      push edx
'006c6ff5    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c6ff6    ff1580104000            call dword ptr [00401080]
'006c6ffc    8d4dd4                  lea ecx, dword ptr [ebp-2c]
'006c6fff    ffd7                    call edi
    '#Cleanup(var_44)
'006c7001    8b06                    mov eax, dword ptr [esi]
'006c7003    56                      push esi
'006c7004    ff9054030000            call dword ptr [eax+00000354]
'006c700a    8b55e8                  mov edx, dword ptr [ebp-18]
'006c700d    83ec10                  sub esp, 10
'006c7010    8bdc                    mov ebx, esp
'006c7012    b909000000              mov ecx, 00000009
'006c7017    890b                    mov dword ptr [ebx], ecx
'006c7019    894dd4                  mov dword ptr [ebp-2c], ecx
'006c701c    8b4dd8                  mov ecx, dword ptr [ebp-28]
'006c701f    894b04                  mov dword ptr [ebx+04], ecx
'006c7022    894308                  mov dword ptr [ebx+08], eax
'006c7025    8945dc                  mov dword ptr [ebp-24], eax
'006c7028    8b45e0                  mov eax, dword ptr [ebp-20]
'006c702b    89430c                  mov dword ptr [ebx+0c], eax
'006c702e    83ec10                  sub esp, 10
'006c7031    c78564ffffff08000000    mov dword ptr [ebp+ffffff64], 00000008
'006c703b    8b8564ffffff            mov eax, dword ptr [ebp+ffffff64]
'006c7041    8bcc                    mov ecx, esp
'006c7043    8901                    mov dword ptr [ecx], eax
'006c7045    8b8568ffffff            mov eax, dword ptr [ebp+ffffff68]
'006c704b    894104                  mov dword ptr [ecx+04], eax
'006c704e    c7856cffffffe4b24300    mov dword ptr [ebp+ffffff6c], 0043b2e4
'006c7058    8b856cffffff            mov eax, dword ptr [ebp+ffffff6c]
'006c705e    8b12                    mov edx, dword ptr [edx]
'006c7060    894108                  mov dword ptr [ecx+08], eax
'006c7063    8b8570ffffff            mov eax, dword ptr [ebp+ffffff70]
'006c7069    89410c                  mov dword ptr [ecx+0c], eax
'006c706c    8b4de8                  mov ecx, dword ptr [ebp-18]
'006c706f    51                      push ecx
'006c7070    ff9228010000            call dword ptr [edx+00000128]
'006c7076    dbe2                    fnclex
'006c7078    85c0                    test ax, ax
'006c707a    7d15                    jge 6c7091
'006c707c    8b55e8                  mov edx, dword ptr [ebp-18]
'006c707f    6828010000              push 00000128
'006c7084    6830314300              push 00433130
'006c7089    52                      push edx
'006c708a    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c708b    ff1580104000            call dword ptr [00401080]
'006c7091    8d4dd4                  lea ecx, dword ptr [ebp-2c]
'006c7094    ffd7                    call edi
    '#Cleanup(var_44)
'006c7096    8b06                    mov eax, dword ptr [esi]
'006c7098    56                      push esi
'006c7099    ff9028030000            call dword ptr [eax+00000328]
'006c709f    8d4dd4                  lea ecx, dword ptr [ebp-2c]
'006c70a2    bf09000000              mov edi, 00000009
'006c70a7    51                      push ecx
'006c70a8    8945dc                  mov dword ptr [ebp-24], eax
'006c70ab    897dd4                  mov dword ptr [ebp-2c], edi

' *** Reference to "rtcIsNumeric"
'006c70ae    ff154c114000            call dword ptr [0040114c]
'006c70b4    8b16                    mov edx, dword ptr [esi]
'006c70b6    56                      push esi
'006c70b7    8985b0feffff            mov dword ptr [ebp+fffffeb0], eax
'006c70bd    ff9228030000            call dword ptr [edx+00000328]
'006c70c3    8945bc                  mov dword ptr [ebp-44], eax
'006c70c6    8b85b0feffff            mov eax, dword ptr [ebp+fffffeb0]
'006c70cc    f7d0                    not eax
'006c70ce    8d4db4                  lea ecx, dword ptr [ebp-4c]
'006c70d1    51                      push ecx
'006c70d2    8d55c4                  lea edx, dword ptr [ebp-3c]
'006c70d5    6689855cffffff          mov word ptr [ebp+ffffff5c], ax
'006c70dc    52                      push edx
'006c70dd    8d8554ffffff            lea eax, dword ptr [ebp+ffffff54]
'006c70e3    50                      push eax
'006c70e4    8d4da4                  lea ecx, dword ptr [ebp-5c]
'006c70e7    51                      push ecx
'006c70e8    897db4                  mov dword ptr [ebp-4c], edi
'006c70eb    c745c401000000          mov dword ptr [ebp-3c], 00000001
'006c70f2    c78554ffffff0b000000    mov dword ptr [ebp+ffffff54], 0000000b

' *** Reference to "rtcImmediateIf"
'006c70fc    ff154c124000            call dword ptr [0040124c]
'006c7102    8b5da4                  mov ebx, dword ptr [ebp-5c]
'006c7105    8b55e8                  mov edx, dword ptr [ebp-18]
'006c7108    8b12                    mov edx, dword ptr [edx]
'006c710a    83ec10                  sub esp, 10
'006c710d    8bfc                    mov edi, esp
'006c710f    891f                    mov dword ptr [edi], ebx
'006c7111    8b5da8                  mov ebx, dword ptr [ebp-58]
'006c7114    895f04                  mov dword ptr [edi+04], ebx
'006c7117    8b5dac                  mov ebx, dword ptr [ebp-54]
'006c711a    895f08                  mov dword ptr [edi+08], ebx
'006c711d    8b5db0                  mov ebx, dword ptr [ebp-50]
'006c7120    895f0c                  mov dword ptr [edi+0c], ebx
'006c7123    83ec10                  sub esp, 10
'006c7126    8bfc                    mov edi, esp
'006c7128    b908000000              mov ecx, 00000008
'006c712d    890f                    mov dword ptr [edi], ecx
'006c712f    8b8d48ffffff            mov ecx, dword ptr [ebp+ffffff48]
'006c7135    894f04                  mov dword ptr [edi+04], ecx
'006c7138    8b4de8                  mov ecx, dword ptr [ebp-18]
'006c713b    b858384500              mov eax, 00453858
'006c7140    894708                  mov dword ptr [edi+08], eax
'006c7143    8b8550ffffff            mov eax, dword ptr [ebp+ffffff50]
'006c7149    51                      push ecx
'006c714a    89470c                  mov dword ptr [edi+0c], eax
'006c714d    ff9228010000            call dword ptr [edx+00000128]
'006c7153    dbe2                    fnclex
'006c7155    85c0                    test ax, ax
'006c7157    7d15                    jge 6c716e
'006c7159    8b55e8                  mov edx, dword ptr [ebp-18]
'006c715c    6828010000              push 00000128
'006c7161    6830314300              push 00433130
'006c7166    52                      push edx
'006c7167    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c7168    ff1580104000            call dword ptr [00401080]
'006c716e    8d45a4                  lea eax, dword ptr [ebp-5c]
'006c7171    50                      push eax
'006c7172    8d4db4                  lea ecx, dword ptr [ebp-4c]
'006c7175    51                      push ecx
'006c7176    8d55c4                  lea edx, dword ptr [ebp-3c]
'006c7179    52                      push edx
'006c717a    8d8554ffffff            lea eax, dword ptr [ebp+ffffff54]
'006c7180    50                      push eax
'006c7181    8d4dd4                  lea ecx, dword ptr [ebp-2c]
'006c7184    51                      push ecx
'006c7185    6a05                    push 05

' *** Reference to "__vbaFreeVarList"
'006c7187    ff1540104000            call dword ptr [00401040]
    '#Cleanup( var_44, var_124, var_9, var_62, var_83)
'006c718d    8b16                    mov edx, dword ptr [esi]
'006c718f    83c418                  add esp, 18
'006c7192    56                      push esi
'006c7193    ff9224030000            call dword ptr [edx+00000324]
'006c7199    8945dc                  mov dword ptr [ebp-24], eax
'006c719c    8d45d4                  lea eax, dword ptr [ebp-2c]
'006c719f    50                      push eax
'006c71a0    8d4dc4                  lea ecx, dword ptr [ebp-3c]
'006c71a3    bf09000000              mov edi, 00000009
'006c71a8    51                      push ecx
'006c71a9    897dd4                  mov dword ptr [ebp-2c], edi

' *** Reference to "rtcTrimVar"
'006c71ac    ff15e4104000            call dword ptr [004010e4]
'006c71b2    8b16                    mov edx, dword ptr [esi]
'006c71b4    56                      push esi
'006c71b5    ff9224030000            call dword ptr [edx+00000324]
'006c71bb    89458c                  mov dword ptr [ebp-74], eax
'006c71be    8d45c4                  lea eax, dword ptr [ebp-3c]
'006c71c1    50                      push eax
'006c71c2    8d8d64ffffff            lea ecx, dword ptr [ebp+ffffff64]
'006c71c8    51                      push ecx
'006c71c9    8d55b4                  lea edx, dword ptr [ebp-4c]
'006c71cc    52                      push edx
'006c71cd    897d84                  mov dword ptr [ebp-7c], edi
'006c71d0    c7459401000000          mov dword ptr [ebp-6c], 00000001
'006c71d7    c7856cffffffcc134300    mov dword ptr [ebp+ffffff6c], 004313cc
'006c71e1    c78564ffffff08800000    mov dword ptr [ebp+ffffff64], 00008008

' *** Reference to "__vbaVarCmpEq"
'006c71eb    ff1580124000            call dword ptr [00401280]
'006c71f1    8bd0                    mov edx, eax
'006c71f3    8d4da4                  lea ecx, dword ptr [ebp-5c]

' *** Reference to "__vbaVarMove"
'006c71f6    ff151c104000            call dword ptr [0040101c]
    var_83 = (#NOT SUPPORTED#)
'006c71fc    8d4584                  lea eax, dword ptr [ebp-7c]
'006c71ff    50                      push eax
'006c7200    8d4d94                  lea ecx, dword ptr [ebp-6c]
'006c7203    51                      push ecx
'006c7204    8d55a4                  lea edx, dword ptr [ebp-5c]
'006c7207    52                      push edx
'006c7208    8d8574ffffff            lea eax, dword ptr [ebp+ffffff74]
'006c720e    50                      push eax

' *** Reference to "rtcImmediateIf"
'006c720f    ff154c124000            call dword ptr [0040124c]
'006c7215    8b9d74ffffff            mov ebx, dword ptr [ebp+ffffff74]
'006c721b    8b55e8                  mov edx, dword ptr [ebp-18]
'006c721e    8b12                    mov edx, dword ptr [edx]
'006c7220    83ec10                  sub esp, 10
'006c7223    8bfc                    mov edi, esp
'006c7225    891f                    mov dword ptr [edi], ebx
'006c7227    8b9d78ffffff            mov ebx, dword ptr [ebp+ffffff78]
'006c722d    895f04                  mov dword ptr [edi+04], ebx
'006c7230    8b9d7cffffff            mov ebx, dword ptr [ebp+ffffff7c]
'006c7236    895f08                  mov dword ptr [edi+08], ebx
'006c7239    8b5d80                  mov ebx, dword ptr [ebp-80]
'006c723c    895f0c                  mov dword ptr [edi+0c], ebx
'006c723f    83ec10                  sub esp, 10
'006c7242    8bfc                    mov edi, esp
'006c7244    b908000000              mov ecx, 00000008
'006c7249    890f                    mov dword ptr [edi], ecx
'006c724b    8b8d48ffffff            mov ecx, dword ptr [ebp+ffffff48]
'006c7251    894f04                  mov dword ptr [edi+04], ecx
'006c7254    8b4de8                  mov ecx, dword ptr [ebp-18]
'006c7257    b87c384500              mov eax, 0045387c
'006c725c    894708                  mov dword ptr [edi+08], eax
'006c725f    8b8550ffffff            mov eax, dword ptr [ebp+ffffff50]
'006c7265    51                      push ecx
'006c7266    89470c                  mov dword ptr [edi+0c], eax
'006c7269    ff9228010000            call dword ptr [edx+00000128]
'006c726f    dbe2                    fnclex
'006c7271    85c0                    test ax, ax
'006c7273    7d19                    jge 6c728e
'006c7275    8b55e8                  mov edx, dword ptr [ebp-18]

' *** Reference to "__vbaHresultCheckObj"
'006c7278    8b1d80104000            mov ebx, dword ptr [00401080]
'006c727e    6828010000              push 00000128
'006c7283    6830314300              push 00433130
'006c7288    52                      push edx
'006c7289    50                      push eax
'006c728a    ffd3                    call ebx
'006c728c    eb06                    jmp 6c7294

' *** Reference to "__vbaHresultCheckObj"
'006c728e    8b1d80104000            mov ebx, dword ptr [00401080]
'006c7294    8d8574ffffff            lea eax, dword ptr [ebp+ffffff74]
'006c729a    50                      push eax
'006c729b    8d4d84                  lea ecx, dword ptr [ebp-7c]
'006c729e    51                      push ecx
'006c729f    8d5594                  lea edx, dword ptr [ebp-6c]
'006c72a2    52                      push edx
'006c72a3    8d45a4                  lea eax, dword ptr [ebp-5c]
'006c72a6    50                      push eax
'006c72a7    8d4dc4                  lea ecx, dword ptr [ebp-3c]
'006c72aa    51                      push ecx
'006c72ab    8d55d4                  lea edx, dword ptr [ebp-2c]
'006c72ae    52                      push edx
'006c72af    6a06                    push 06

' *** Reference to "__vbaFreeVarList"
'006c72b1    ff1540104000            call dword ptr [00401040]
    '#Cleanup( var_44, var_9, var_83, var_121, var_119, var_91)
'006c72b7    8b45e8                  mov eax, dword ptr [ebp-18]
'006c72ba    8b08                    mov ecx, dword ptr [eax]
'006c72bc    83c41c                  add esp, 1c
'006c72bf    6a00                    push 00
'006c72c1    6a01                    push 01
'006c72c3    50                      push eax
'006c72c4    ff9164010000            call dword ptr [ecx+00000164]
'006c72ca    dbe2                    fnclex
'006c72cc    85c0                    test ax, ax
'006c72ce    7d11                    jge 6c72e1
'006c72d0    8b55e8                  mov edx, dword ptr [ebp-18]
'006c72d3    6864010000              push 00000164
'006c72d8    6830314300              push 00433130
'006c72dd    52                      push edx
'006c72de    50                      push eax
'006c72df    ffd3                    call ebx
    
End If
'006c72e1    8b45e8                  mov eax, dword ptr [ebp-18]
'006c72e4    8b08                    mov ecx, dword ptr [eax]
'006c72e6    50                      push eax
'006c72e7    ff91c4000000            call dword ptr [ecx+000000c4]
'006c72ed    dbe2                    fnclex
'006c72ef    85c0                    test ax, ax
'006c72f1    7d11                    jge 6c7304
'006c72f3    8b55e8                  mov edx, dword ptr [ebp-18]
'006c72f6    68c4000000              push 000000c4
'006c72fb    6830314300              push 00433130
'006c7300    52                      push edx
'006c7301    50                      push eax
'006c7302    ffd3                    call ebx
'006c7304    8b06                    mov eax, dword ptr [esi]
'006c7306    56                      push esi
'006c7307    ff9048070000            call dword ptr [eax+00000748]
'006c730d    85c0                    test ax, ax
'006c730f    7d0e                    jge 6c731f
'006c7311    6848070000              push 00000748
'006c7316    6800144300              push 00431400
'006c731b    56                      push esi
'006c731c    50                      push eax
'006c731d    ffd3                    call ebx
'006c731f    c745fc00000000          mov dword ptr [ebp-04], 00000000
'006c7326    686b736c00              push 006c736b
'006c732b    eb34                    jmp 6c7361
'006c732d    8d4de4                  lea ecx, dword ptr [ebp-1c]

' *** Reference to "__vbaFreeObj"
'006c7330    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_40)
'006c7336    8d8d74ffffff            lea ecx, dword ptr [ebp+ffffff74]
'006c733c    51                      push ecx
'006c733d    8d5584                  lea edx, dword ptr [ebp-7c]
'006c7340    52                      push edx
'006c7341    8d4594                  lea eax, dword ptr [ebp-6c]
'006c7344    50                      push eax
'006c7345    8d4da4                  lea ecx, dword ptr [ebp-5c]
'006c7348    51                      push ecx
'006c7349    8d55b4                  lea edx, dword ptr [ebp-4c]
'006c734c    52                      push edx
'006c734d    8d45c4                  lea eax, dword ptr [ebp-3c]
'006c7350    50                      push eax
'006c7351    8d4dd4                  lea ecx, dword ptr [ebp-2c]
'006c7354    51                      push ecx
'006c7355    6a07                    push 07

' *** Reference to "__vbaFreeVarList"
'006c7357    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_44, var_9, var_62, var_83, var_121, var_119, var_91)
'006c735d    83c420                  add esp, 20
'006c7360    c3                      ret
'006c7361    8d4de8                  lea ecx, dword ptr [ebp-18]

' *** Reference to "__vbaFreeObj"
'006c7364    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_41)
'006c736a    c3                      ret
'006c736b    8b4508                  mov eax, dword ptr [ebp+08]
'006c736e    8b10                    mov edx, dword ptr [eax]
'006c7370    50                      push eax
'006c7371    ff5208                  call dword ptr [edx+08]
'006c7374    8b45fc                  mov eax, dword ptr [ebp-04]
'006c7377    8b4dec                  mov ecx, dword ptr [ebp-14]
'006c737a    5f                      pop edi
'006c737b    5e                      pop esi
    'Reference to '__except_list'
'006c737c    64890d00000000          mov dword ptr fs:[00000000], ecx
'006c7383    5b                      pop ebx
'006c7384    8be5                    mov esp, ebp
'006c7386    5d                      pop ebp
'006c7387    c20400                  ret 0004


End Sub


'Event for btnSuivant
Private Sub btnSuivant_Click()
'006c7920    55                      push ebp
'006c7921    8bec                    mov ebp, esp
'006c7923    83ec0c                  sub esp, 0c

' *** Reference to "__vbaExceptHandler"
'006c7926    6866724000              push 00407266
'006c792b    64a100000000            mov ax, word ptr fs:[00000000]
'006c7931    50                      push eax
    'Reference to '__except_list'
'006c7932    64892500000000          mov dword ptr fs:[00000000], esp
'006c7939    81ec84000000            sub esp, 00000084
'006c793f    53                      push ebx
'006c7940    56                      push esi
'006c7941    57                      push edi
'006c7942    8965f4                  mov dword ptr [ebp-0c], esp
'006c7945    c745f870674000          mov dword ptr [ebp-08], 00406770
'006c794c    8b7d08                  mov edi, dword ptr [ebp+08]
'006c794f    8bc7                    mov eax, edi
'006c7951    83e001                  and eax, 01
'006c7954    8945fc                  mov dword ptr [ebp-04], eax
'006c7957    83e7fe                  and edi, fffffffe
'006c795a    8b0f                    mov ecx, dword ptr [edi]
'006c795c    57                      push edi
'006c795d    897d08                  mov dword ptr [ebp+08], edi
'006c7960    ff5104                  call dword ptr [ecx+04]
'006c7963    8b17                    mov edx, dword ptr [edi]
'006c7965    33db                    xor ebx, ebx
'006c7967    57                      push edi
'006c7968    895de8                  mov dword ptr [ebp-18], ebx
'006c796b    895de4                  mov dword ptr [ebp-1c], ebx
'006c796e    895de0                  mov dword ptr [ebp-20], ebx
'006c7971    895ddc                  mov dword ptr [ebp-24], ebx
'006c7974    895dd8                  mov dword ptr [ebp-28], ebx
'006c7977    895dd4                  mov dword ptr [ebp-2c], ebx
'006c797a    895dd0                  mov dword ptr [ebp-30], ebx
'006c797d    895dcc                  mov dword ptr [ebp-34], ebx
'006c7980    895dc8                  mov dword ptr [ebp-38], ebx
'006c7983    895db8                  mov dword ptr [ebp-48], ebx
'006c7986    895da8                  mov dword ptr [ebp-58], ebx
'006c7989    895d98                  mov dword ptr [ebp-68], ebx
'006c798c    895d94                  mov dword ptr [ebp-6c], ebx
'006c798f    895d90                  mov dword ptr [ebp-70], ebx
'006c7992    ff92fc020000            call dword ptr [edx+000002fc]
'006c7998    50                      push eax
'006c7999    8d45cc                  lea eax, dword ptr [ebp-34]
'006c799c    50                      push eax

' *** Reference to "__vbaObjSet"
'006c799d    ff15b4104000            call dword ptr [004010b4]
Set var_43 = Me
'006c79a3    8d5594                  lea edx, dword ptr [ebp-6c]
'006c79a6    8bf0                    mov esi, eax
'006c79a8    8b0e                    mov ecx, dword ptr [esi]
'006c79aa    52                      push edx
'006c79ab    56                      push esi
'006c79ac    ff91f0000000            call dword ptr [ecx+000000f0]
'006c79b2    dbe2                    fnclex
'006c79b4    3bc3                    cmp eax, ebx
'006c79b6    7d12                    jge 6c79ca

If (var_43 < 0) Then
'006c79b8    68f0000000              push 000000f0
'006c79bd    681cb94300              push 0043b91c
'006c79c2    56                      push esi
'006c79c3    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c79c4    ff1580104000            call dword ptr [00401080]
    
End If
'006c79ca    8b07                    mov eax, dword ptr [edi]
'006c79cc    57                      push edi
'006c79cd    ff90fc020000            call dword ptr [eax+000002fc]
'006c79d3    50                      push eax
'006c79d4    8d4dc8                  lea ecx, dword ptr [ebp-38]
'006c79d7    51                      push ecx

' *** Reference to "__vbaObjSet"
'006c79d8    ff15b4104000            call dword ptr [004010b4]
Set var_46 = Nothing
'006c79de    8bf0                    mov esi, eax
'006c79e0    8b16                    mov edx, dword ptr [esi]
'006c79e2    8d4590                  lea eax, dword ptr [ebp-70]
'006c79e5    50                      push eax
'006c79e6    56                      push esi
'006c79e7    ff92e8000000            call dword ptr [edx+000000e8]
'006c79ed    dbe2                    fnclex
'006c79ef    3bc3                    cmp eax, ebx
'006c79f1    7d12                    jge 6c7a05
'006c79f3    68e8000000              push 000000e8
'006c79f8    681cb94300              push 0043b91c
'006c79fd    56                      push esi
'006c79fe    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c79ff    ff1580104000            call dword ptr [00401080]
'006c7a05    668b7594                mov si, word ptr [ebp-6c]
'006c7a09    6683c601                add si, 01
var_num8 = 0 + 1
'006c7a0d    0f80b4030000            jo 6c7dc7
'006c7a13    662b7590                sub si, word ptr [ebp-70]
var_num8 = var_num8 - 0
'006c7a17    8d4dc8                  lea ecx, dword ptr [ebp-38]
'006c7a1a    66f7de                  neg si
'006c7a1d    51                      push ecx
'006c7a1e    8d55cc                  lea edx, dword ptr [ebp-34]
'006c7a21    52                      push edx
'006c7a22    6a02                    push 02
'006c7a24    1bf6                    sbb esi, esi
'006c7a26    46                      inc esi
'006c7a27    f7de                    neg esi

' *** Reference to "__vbaFreeObjList"
'006c7a29    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_43, var_46)
'006c7a2f    8b07                    mov eax, dword ptr [edi]
'006c7a31    83c40c                  add esp, 0c
'006c7a34    663bf3                  cmp si, bx
'006c7a37    57                      push edi
'006c7a38    7443                    je 6c7a7d

If (-(CBool((var_num8))) <> 0) Then
'006c7a3a    ff90fc020000            call dword ptr [eax+000002fc]
'006c7a40    50                      push eax
'006c7a41    8d4dcc                  lea ecx, dword ptr [ebp-34]
'006c7a44    51                      push ecx

' *** Reference to "__vbaObjSet"
'006c7a45    ff15b4104000            call dword ptr [004010b4]
    Set var_43 = Nothing
'006c7a4b    8bf0                    mov esi, eax
'006c7a4d    8b16                    mov edx, dword ptr [esi]
'006c7a4f    53                      push ebx
'006c7a50    56                      push esi
'006c7a51    ff92f4000000            call dword ptr [edx+000000f4]
'006c7a57    dbe2                    fnclex
'006c7a59    3bc3                    cmp eax, ebx
'006c7a5b    7d12                    jge 6c7a6f
    
    If (    var_43 < 0) Then
'006c7a5d    68f4000000              push 000000f4
'006c7a62    681cb94300              push 0043b91c
'006c7a67    56                      push esi
'006c7a68    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c7a69    ff1580104000            call dword ptr [00401080]
    
End If
'006c7a6f    8d4dcc                  lea ecx, dword ptr [ebp-34]

' *** Reference to "__vbaFreeObj"
'006c7a72    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_43)
'006c7a78    e993000000              jmp 6c7b10

'ERROR: Two many next close:
End If
'006c7a7d    ff90fc020000            call dword ptr [eax+000002fc]

' *** Reference to "__vbaObjSet"
'006c7a83    8b35b4104000            mov esi, dword ptr [004010b4]
'006c7a89    50                      push eax
'006c7a8a    8d4dc8                  lea ecx, dword ptr [ebp-38]
'006c7a8d    51                      push ecx
'006c7a8e    ffd6                    call esi
Set var_46 = Nothing
'006c7a90    8b17                    mov edx, dword ptr [edi]
'006c7a92    57                      push edi
'006c7a93    894584                  mov dword ptr [ebp-7c], eax
'006c7a96    ff92fc020000            call dword ptr [edx+000002fc]
'006c7a9c    50                      push eax
'006c7a9d    8d45cc                  lea eax, dword ptr [ebp-34]
'006c7aa0    50                      push eax
'006c7aa1    ffd6                    call esi
Set var_43 = Nothing
'006c7aa3    8d5594                  lea edx, dword ptr [ebp-6c]
'006c7aa6    8bf0                    mov esi, eax
'006c7aa8    8b0e                    mov ecx, dword ptr [esi]
'006c7aaa    52                      push edx
'006c7aab    56                      push esi
'006c7aac    ff91f0000000            call dword ptr [ecx+000000f0]
'006c7ab2    dbe2                    fnclex
'006c7ab4    3bc3                    cmp eax, ebx
'006c7ab6    7d12                    jge 6c7aca

If (var_43 < 0) Then
'006c7ab8    68f0000000              push 000000f0
'006c7abd    681cb94300              push 0043b91c
'006c7ac2    56                      push esi
'006c7ac3    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c7ac4    ff1580104000            call dword ptr [00401080]
    
End If
'006c7aca    668b4d94                mov cx, word ptr [ebp-6c]
'006c7ace    8b7584                  mov esi, dword ptr [ebp-7c]
'006c7ad1    8b06                    mov eax, dword ptr [esi]
'006c7ad3    6683c101                add cx, 01
var_num3 = 0 + 1
'006c7ad7    0f80ea020000            jo 6c7dc7
'006c7add    51                      push ecx
'006c7ade    56                      push esi
'006c7adf    ff90f4000000            call dword ptr [eax+000000f4]
'006c7ae5    dbe2                    fnclex
'006c7ae7    3bc3                    cmp eax, ebx
'006c7ae9    7d12                    jge 6c7afd

If (-256 - 12 < 0) Then
'006c7aeb    68f4000000              push 000000f4
'006c7af0    681cb94300              push 0043b91c
'006c7af5    56                      push esi
'006c7af6    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c7af7    ff1580104000            call dword ptr [00401080]
    
End If
'006c7afd    8d55c8                  lea edx, dword ptr [ebp-38]
'006c7b00    52                      push edx
'006c7b01    8d45cc                  lea eax, dword ptr [ebp-34]
'006c7b04    50                      push eax
'006c7b05    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'006c7b07    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_43, var_46)
'006c7b0d    83c40c                  add esp, 0c
'006c7b10    8b0f                    mov ecx, dword ptr [edi]
'006c7b12    57                      push edi
'006c7b13    ff91fc020000            call dword ptr [ecx+000002fc]
'006c7b19    50                      push eax
'006c7b1a    8d55cc                  lea edx, dword ptr [ebp-34]
'006c7b1d    52                      push edx

' *** Reference to "__vbaObjSet"
'006c7b1e    ff15b4104000            call dword ptr [004010b4]
Set var_43 = 
'006c7b24    8d4de4                  lea ecx, dword ptr [ebp-1c]
'006c7b27    8bf0                    mov esi, eax
'006c7b29    8b06                    mov eax, dword ptr [esi]
'006c7b2b    51                      push ecx
'006c7b2c    56                      push esi
'006c7b2d    ff90a8000000            call dword ptr [eax+000000a8]
'006c7b33    dbe2                    fnclex
'006c7b35    3bc3                    cmp eax, ebx
'006c7b37    7d12                    jge 6c7b4b

If (var_43 < 0) Then
'006c7b39    68a8000000              push 000000a8
'006c7b3e    681cb94300              push 0043b91c
'006c7b43    56                      push esi
'006c7b44    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c7b45    ff1580104000            call dword ptr [00401080]
    
End If
'006c7b4b    8b55e4                  mov edx, dword ptr [ebp-1c]
'006c7b4e    52                      push edx
'006c7b4f    68cc134300              push 004313cc

' *** Reference to "__vbaStrCmp"
'006c7b54    ff1530114000            call dword ptr [00401130]
'006c7b5a    8bf0                    mov esi, eax
'006c7b5c    f7de                    neg esi
'006c7b5e    1bf6                    sbb esi, esi
'006c7b60    f7de                    neg esi
'006c7b62    8d4de4                  lea ecx, dword ptr [ebp-1c]
'006c7b65    f7de                    neg esi

' *** Reference to "__vbaFreeStr"
'006c7b67    ff150c134000            call dword ptr [0040130c]
'#Cleanup(var_40)
'006c7b6d    8d4dcc                  lea ecx, dword ptr [ebp-34]

' *** Reference to "__vbaFreeObj"
'006c7b70    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_43)
'006c7b76    663bf3                  cmp si, bx
'006c7b79    0f84e1010000            je 6c7d60

If (((vbNullString) <> (vbNullChar))) Then
'006c7b7f    8b07                    mov eax, dword ptr [edi]
'006c7b81    57                      push edi
'006c7b82    ff90fc020000            call dword ptr [eax+000002fc]
'006c7b88    50                      push eax
'006c7b89    8d4dcc                  lea ecx, dword ptr [ebp-34]
'006c7b8c    51                      push ecx

' *** Reference to "__vbaObjSet"
'006c7b8d    ff15b4104000            call dword ptr [004010b4]
    Set var_43 = Nothing
'006c7b93    8bf0                    mov esi, eax
'006c7b95    8b16                    mov edx, dword ptr [esi]
'006c7b97    8d45e4                  lea eax, dword ptr [ebp-1c]
'006c7b9a    50                      push eax
'006c7b9b    56                      push esi
'006c7b9c    ff92a8000000            call dword ptr [edx+000000a8]
'006c7ba2    dbe2                    fnclex
'006c7ba4    3bc3                    cmp eax, ebx
'006c7ba6    7d12                    jge 6c7bba
'006c7ba8    68a8000000              push 000000a8
'006c7bad    681cb94300              push 0043b91c
'006c7bb2    56                      push esi
'006c7bb3    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c7bb4    ff1580104000            call dword ptr [00401080]
'006c7bba    8b55e4                  mov edx, dword ptr [ebp-1c]

' *** Reference to "__vbaStrMove"
'006c7bbd    8b35d0124000            mov esi, dword ptr [004012d0]
'006c7bc3    8d4de0                  lea ecx, dword ptr [ebp-20]
'006c7bc6    895de4                  mov dword ptr [ebp-1c], ebx
'006c7bc9    ffd6                    call esi
'006c7bcb    8d4de0                  lea ecx, dword ptr [ebp-20]
'006c7bce    51                      push ecx
'006c7bcf    e81cc0e2ff              call 4f3bf0
    Call sub_4F3BF0()
'006c7bd4    8bd0                    mov edx, eax
'006c7bd6    8d4dd0                  lea ecx, dword ptr [ebp-30]
'006c7bd9    ffd6                    call esi
'006c7bdb    8b55d0                  mov edx, dword ptr [ebp-30]
'006c7bde    89956cffffff            mov dword ptr [ebp+ffffff6c], edx
'006c7be4    8b154cb07200            mov edx, dword ptr [0072b04c]
'006c7bea    895dd0                  mov dword ptr [ebp-30], ebx
'006c7bed    8b1a                    mov ebx, dword ptr [edx]
'006c7bef    8d55c8                  lea edx, dword ptr [ebp-38]
'006c7bf2    52                      push edx
'006c7bf3    83ec10                  sub esp, 10
'006c7bf6    8bd4                    mov edx, esp
'006c7bf8    b90a000000              mov ecx, 0000000a
'006c7bfd    890a                    mov dword ptr [edx], ecx
'006c7bff    894da8                  mov dword ptr [ebp-58], ecx
'006c7c02    8b4d9c                  mov ecx, dword ptr [ebp-64]
'006c7c05    894a04                  mov dword ptr [edx+04], ecx
'006c7c08    b804000280              mov eax, 80020004
'006c7c0d    894208                  mov dword ptr [edx+08], eax
'006c7c10    8945b0                  mov dword ptr [ebp-50], eax
'006c7c13    8b45a4                  mov eax, dword ptr [ebp-5c]
'006c7c16    89420c                  mov dword ptr [edx+0c], eax
'006c7c19    8b55a8                  mov edx, dword ptr [ebp-58]
'006c7c1c    8b45ac                  mov eax, dword ptr [ebp-54]
'006c7c1f    83ec10                  sub esp, 10
'006c7c22    8bcc                    mov ecx, esp
'006c7c24    8911                    mov dword ptr [ecx], edx
'006c7c26    8b55b0                  mov edx, dword ptr [ebp-50]
'006c7c29    894104                  mov dword ptr [ecx+04], eax
'006c7c2c    8b45b4                  mov eax, dword ptr [ebp-4c]
'006c7c2f    895108                  mov dword ptr [ecx+08], edx
'006c7c32    8b55bc                  mov edx, dword ptr [ebp-44]
'006c7c35    89410c                  mov dword ptr [ecx+0c], eax
'006c7c38    83ec10                  sub esp, 10
'006c7c3b    8bcc                    mov ecx, esp
'006c7c3d    b803000000              mov eax, 00000003
'006c7c42    8901                    mov dword ptr [ecx], eax
'006c7c44    895104                  mov dword ptr [ecx+04], edx
'006c7c47    8b55c4                  mov edx, dword ptr [ebp-3c]
'006c7c4a    c745c002000000          mov dword ptr [ebp-40], 00000002
'006c7c51    8b45c0                  mov eax, dword ptr [ebp-40]
'006c7c54    894108                  mov dword ptr [ecx+08], eax
'006c7c57    89510c                  mov dword ptr [ecx+0c], edx
'006c7c5a    8b956cffffff            mov edx, dword ptr [ebp+ffffff6c]
'006c7c60    68c0384500              push 004538c0
'006c7c65    8d4ddc                  lea ecx, dword ptr [ebp-24]
'006c7c68    ffd6                    call esi
'006c7c6a    50                      push eax

' *** Reference to "__vbaStrCat"
'006c7c6b    ff1570104000            call dword ptr [00401070]
    var_49 = ("select nom from sort where nom='") & (vbNullString)
'006c7c71    8bd0                    mov edx, eax
'006c7c73    8d4dd8                  lea ecx, dword ptr [ebp-28]
'006c7c76    ffd6                    call esi
'006c7c78    50                      push eax
'006c7c79    6854a44300              push 0043a454

' *** Reference to "__vbaStrCat"
'006c7c7e    ff1570104000            call dword ptr [00401070]
    var_63 = (var_49) & ("'")
'006c7c84    8bd0                    mov edx, eax
'006c7c86    8d4dd4                  lea ecx, dword ptr [ebp-2c]
'006c7c89    ffd6                    call esi
'006c7c8b    50                      push eax
'006c7c8c    a14cb07200              mov ax, word ptr [0072b04c]
'006c7c91    50                      push eax
'006c7c92    ff93bc000000            call dword ptr [ebx+000000bc]
'006c7c98    dbe2                    fnclex
'006c7c9a    33db                    xor ebx, ebx
    var_num2 = Empty
'006c7c9c    3bc3                    cmp eax, ebx
'006c7c9e    7d1c                    jge 6c7cbc
    
    If (    0 < -256 - 12) Then
'006c7ca0    8b0d4cb07200            mov ecx, dword ptr [0072b04c]

' *** Reference to "__vbaHresultCheckObj"
'006c7ca6    8b3580104000            mov esi, dword ptr [00401080]
'006c7cac    68bc000000              push 000000bc
'006c7cb1    68301f4300              push 00431f30
'006c7cb6    51                      push ecx
'006c7cb7    50                      push eax
'006c7cb8    ffd6                    call esi
'006c7cba    eb06                    jmp 6c7cc2
    
End If

' *** Reference to "__vbaHresultCheckObj"
'006c7cbc    8b3580104000            mov esi, dword ptr [00401080]
'006c7cc2    8b45c8                  mov eax, dword ptr [ebp-38]
'006c7cc5    50                      push eax
'006c7cc6    8d55e8                  lea edx, dword ptr [ebp-18]
'006c7cc9    52                      push edx
'006c7cca    895dc8                  mov dword ptr [ebp-38], ebx

' *** Reference to "__vbaObjSet"
'006c7ccd    ff15b4104000            call dword ptr [004010b4]
Set var_41 = Nothing
'006c7cd3    8d45d0                  lea eax, dword ptr [ebp-30]
'006c7cd6    50                      push eax
'006c7cd7    8d4dd4                  lea ecx, dword ptr [ebp-2c]
'006c7cda    51                      push ecx
'006c7cdb    8d55d8                  lea edx, dword ptr [ebp-28]
'006c7cde    52                      push edx
'006c7cdf    8d45dc                  lea eax, dword ptr [ebp-24]
'006c7ce2    50                      push eax
'006c7ce3    8d4de0                  lea ecx, dword ptr [ebp-20]
'006c7ce6    51                      push ecx
'006c7ce7    6a05                    push 05

' *** Reference to "__vbaFreeStrList"
'006c7ce9    ff155c124000            call dword ptr [0040125c]
'#Cleanup( 0, 0, -4500, -4504, 0)
'006c7cef    83c418                  add esp, 18
'006c7cf2    8d4dcc                  lea ecx, dword ptr [ebp-34]

' *** Reference to "__vbaFreeObj"
'006c7cf5    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_43)
'006c7cfb    8b45e8                  mov eax, dword ptr [ebp-18]
'006c7cfe    8b10                    mov edx, dword ptr [eax]
'006c7d00    8d4d94                  lea ecx, dword ptr [ebp-6c]
'006c7d03    51                      push ecx
'006c7d04    50                      push eax
'006c7d05    ff5234                  call dword ptr [edx+34]
'006c7d08    dbe2                    fnclex
'006c7d0a    3bc3                    cmp eax, ebx
'006c7d0c    7d0e                    jge 6c7d1c

If (var_41 < -256 - 12) Then
'006c7d0e    8b55e8                  mov edx, dword ptr [ebp-18]
'006c7d11    6a34                    push 34
'006c7d13    6830314300              push 00433130
'006c7d18    52                      push edx
'006c7d19    50                      push eax
'006c7d1a    ffd6                    call esi
    
End If
'006c7d1c    66395d94                cmp word ptr [ebp-6c], bx
'006c7d20    743e                    je 6c7d60

If (0 <> -256 - 12) Then
'006c7d22    8b45e8                  mov eax, dword ptr [ebp-18]
'006c7d25    8b08                    mov ecx, dword ptr [eax]
'006c7d27    50                      push eax
'006c7d28    ff91c4000000            call dword ptr [ecx+000000c4]
'006c7d2e    dbe2                    fnclex
'006c7d30    3bc3                    cmp eax, ebx
'006c7d32    7d11                    jge 6c7d45
    
    If (    var_41 < -256 - 12) Then
'006c7d34    8b55e8                  mov edx, dword ptr [ebp-18]
'006c7d37    68c4000000              push 000000c4
'006c7d3c    6830314300              push 00433130
'006c7d41    52                      push edx
'006c7d42    50                      push eax
'006c7d43    ffd6                    call esi
    
End If
'006c7d45    8b07                    mov eax, dword ptr [edi]
'006c7d47    57                      push edi
'006c7d48    ff9040070000            call dword ptr [eax+00000740]
'006c7d4e    3bc3                    cmp eax, ebx
'006c7d50    7d0e                    jge 6c7d60

If (frmDescriptifSort < -256 - 12) Then
'006c7d52    6840070000              push 00000740
'006c7d57    6800144300              push 00431400
'006c7d5c    57                      push edi
'006c7d5d    50                      push eax
'006c7d5e    ffd6                    call esi
    
End If
'006c7d60    895dfc                  mov dword ptr [ebp-04], ebx
'006c7d63    68a87d6c00              push 006c7da8
'006c7d68    eb34                    jmp 6c7d9e
'006c7d6a    8d4dd0                  lea ecx, dword ptr [ebp-30]
'006c7d6d    51                      push ecx
'006c7d6e    8d55d4                  lea edx, dword ptr [ebp-2c]
'006c7d71    52                      push edx
'006c7d72    8d45d8                  lea eax, dword ptr [ebp-28]
'006c7d75    50                      push eax
'006c7d76    8d4ddc                  lea ecx, dword ptr [ebp-24]
'006c7d79    51                      push ecx
'006c7d7a    8d55e0                  lea edx, dword ptr [ebp-20]
'006c7d7d    52                      push edx
'006c7d7e    8d45e4                  lea eax, dword ptr [ebp-1c]
'006c7d81    50                      push eax
'006c7d82    6a06                    push 06

' *** Reference to "__vbaFreeStrList"
'006c7d84    ff155c124000            call dword ptr [0040125c]
'#Cleanup( 0, 0, 0, -4500, -4504, 0)
'006c7d8a    8d4dc8                  lea ecx, dword ptr [ebp-38]
'006c7d8d    51                      push ecx
'006c7d8e    8d55cc                  lea edx, dword ptr [ebp-34]
'006c7d91    52                      push edx
'006c7d92    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'006c7d94    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_43, var_46)
'006c7d9a    83c428                  add esp, 28
'006c7d9d    c3                      ret
'006c7d9e    8d4de8                  lea ecx, dword ptr [ebp-18]

' *** Reference to "__vbaFreeObj"
'006c7da1    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_41)
'006c7da7    c3                      ret
'006c7da8    8b4508                  mov eax, dword ptr [ebp+08]
'006c7dab    8b08                    mov ecx, dword ptr [eax]
'006c7dad    50                      push eax
'006c7dae    ff5108                  call dword ptr [ecx+08]
'006c7db1    8b45fc                  mov eax, dword ptr [ebp-04]
'006c7db4    8b4dec                  mov ecx, dword ptr [ebp-14]
'006c7db7    5f                      pop edi
'006c7db8    5e                      pop esi
    'Reference to '__except_list'
'006c7db9    64890d00000000          mov dword ptr fs:[00000000], ecx
'006c7dc0    5b                      pop ebx
'006c7dc1    8be5                    mov esp, ebp
'006c7dc3    5d                      pop ebp
'006c7dc4    c20400                  ret 0004


End Sub


'Event for btnPrecedent
Private Sub btnPrecedent_Click()
'006c7460    55                      push ebp
'006c7461    8bec                    mov ebp, esp
'006c7463    83ec0c                  sub esp, 0c

' *** Reference to "__vbaExceptHandler"
'006c7466    6866724000              push 00407266
'006c746b    64a100000000            mov ax, word ptr fs:[00000000]
'006c7471    50                      push eax
    'Reference to '__except_list'
'006c7472    64892500000000          mov dword ptr fs:[00000000], esp
'006c7479    83ec7c                  sub esp, 7c
'006c747c    53                      push ebx
'006c747d    56                      push esi
'006c747e    57                      push edi
'006c747f    8965f4                  mov dword ptr [ebp-0c], esp
'006c7482    c745f860674000          mov dword ptr [ebp-08], 00406760
'006c7489    8b7d08                  mov edi, dword ptr [ebp+08]
'006c748c    8bc7                    mov eax, edi
'006c748e    83e001                  and eax, 01
'006c7491    8945fc                  mov dword ptr [ebp-04], eax
'006c7494    83e7fe                  and edi, fffffffe
'006c7497    8b0f                    mov ecx, dword ptr [edi]
'006c7499    57                      push edi
'006c749a    897d08                  mov dword ptr [ebp+08], edi
'006c749d    ff5104                  call dword ptr [ecx+04]
'006c74a0    8b17                    mov edx, dword ptr [edi]
'006c74a2    33db                    xor ebx, ebx
'006c74a4    57                      push edi
'006c74a5    895de8                  mov dword ptr [ebp-18], ebx
'006c74a8    895de4                  mov dword ptr [ebp-1c], ebx
'006c74ab    895de0                  mov dword ptr [ebp-20], ebx
'006c74ae    895ddc                  mov dword ptr [ebp-24], ebx
'006c74b1    895dd8                  mov dword ptr [ebp-28], ebx
'006c74b4    895dd4                  mov dword ptr [ebp-2c], ebx
'006c74b7    895dd0                  mov dword ptr [ebp-30], ebx
'006c74ba    895dcc                  mov dword ptr [ebp-34], ebx
'006c74bd    895dc8                  mov dword ptr [ebp-38], ebx
'006c74c0    895db8                  mov dword ptr [ebp-48], ebx
'006c74c3    895da8                  mov dword ptr [ebp-58], ebx
'006c74c6    895d98                  mov dword ptr [ebp-68], ebx
'006c74c9    895d94                  mov dword ptr [ebp-6c], ebx
'006c74cc    ff92fc020000            call dword ptr [edx+000002fc]
'006c74d2    50                      push eax
'006c74d3    8d45cc                  lea eax, dword ptr [ebp-34]
'006c74d6    50                      push eax

' *** Reference to "__vbaObjSet"
'006c74d7    ff15b4104000            call dword ptr [004010b4]
Set var_43 = Me
'006c74dd    8d5594                  lea edx, dword ptr [ebp-6c]
'006c74e0    8bf0                    mov esi, eax
'006c74e2    8b0e                    mov ecx, dword ptr [esi]
'006c74e4    52                      push edx
'006c74e5    56                      push esi
'006c74e6    ff91f0000000            call dword ptr [ecx+000000f0]
'006c74ec    dbe2                    fnclex
'006c74ee    3bc3                    cmp eax, ebx
'006c74f0    7d12                    jge 6c7504

If (var_43 < 0) Then
'006c74f2    68f0000000              push 000000f0
'006c74f7    681cb94300              push 0043b91c
'006c74fc    56                      push esi
'006c74fd    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c74fe    ff1580104000            call dword ptr [00401080]
    
End If
'006c7504    668b7594                mov si, word ptr [ebp-6c]
'006c7508    6683ee01                sub si, 01
var_num8 = 0 - 1
'006c750c    0f80f9030000            jo 6c790b
'006c7512    6646                    inc si
'006c7514    66f7de                  neg si
'006c7517    8d4dcc                  lea ecx, dword ptr [ebp-34]
'006c751a    1bf6                    sbb esi, esi
'006c751c    46                      inc esi
'006c751d    f7de                    neg esi

' *** Reference to "__vbaFreeObj"
'006c751f    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_43)
'006c7525    663bf3                  cmp si, bx
'006c7528    0f8490000000            je 6c75be

If (-(CBool((var_num8))) <> 0) Then
'006c752e    8b07                    mov eax, dword ptr [edi]
'006c7530    57                      push edi
'006c7531    ff90fc020000            call dword ptr [eax+000002fc]

' *** Reference to "__vbaObjSet"
'006c7537    8b35b4104000            mov esi, dword ptr [004010b4]
'006c753d    50                      push eax
'006c753e    8d4dc8                  lea ecx, dword ptr [ebp-38]
'006c7541    51                      push ecx
'006c7542    ffd6                    call esi
    Set var_46 = Nothing
'006c7544    8b17                    mov edx, dword ptr [edi]
'006c7546    57                      push edi
'006c7547    894588                  mov dword ptr [ebp-78], eax
'006c754a    ff92fc020000            call dword ptr [edx+000002fc]
'006c7550    50                      push eax
'006c7551    8d45cc                  lea eax, dword ptr [ebp-34]
'006c7554    50                      push eax
'006c7555    ffd6                    call esi
    Set var_43 = Nothing
'006c7557    8d5594                  lea edx, dword ptr [ebp-6c]
'006c755a    8bf0                    mov esi, eax
'006c755c    8b0e                    mov ecx, dword ptr [esi]
'006c755e    52                      push edx
'006c755f    56                      push esi
'006c7560    ff91e8000000            call dword ptr [ecx+000000e8]
'006c7566    dbe2                    fnclex
'006c7568    3bc3                    cmp eax, ebx
'006c756a    7d12                    jge 6c757e
    
    If (    var_43 < 0) Then
'006c756c    68e8000000              push 000000e8
'006c7571    681cb94300              push 0043b91c
'006c7576    56                      push esi
'006c7577    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c7578    ff1580104000            call dword ptr [00401080]
    
End If
'006c757e    668b4d94                mov cx, word ptr [ebp-6c]
'006c7582    8b7588                  mov esi, dword ptr [ebp-78]
'006c7585    8b06                    mov eax, dword ptr [esi]
'006c7587    6683e901                sub cx, 01
var_num3 = 0 - 1
'006c758b    0f807a030000            jo 6c790b
'006c7591    51                      push ecx
'006c7592    56                      push esi
'006c7593    ff90f4000000            call dword ptr [eax+000000f4]
'006c7599    dbe2                    fnclex
'006c759b    3bc3                    cmp eax, ebx
'006c759d    7d12                    jge 6c75b1

If (-256 - 12 < 0) Then
'006c759f    68f4000000              push 000000f4
'006c75a4    681cb94300              push 0043b91c
'006c75a9    56                      push esi
'006c75aa    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c75ab    ff1580104000            call dword ptr [00401080]
    
End If
'006c75b1    8d55c8                  lea edx, dword ptr [ebp-38]
'006c75b4    52                      push edx
'006c75b5    8d45cc                  lea eax, dword ptr [ebp-34]
'006c75b8    50                      push eax
'006c75b9    e98b000000              jmp 6c7649

'ERROR: Two many next close:
End If
'006c75be    8b0f                    mov ecx, dword ptr [edi]
'006c75c0    57                      push edi
'006c75c1    ff91fc020000            call dword ptr [ecx+000002fc]

' *** Reference to "__vbaObjSet"
'006c75c7    8b35b4104000            mov esi, dword ptr [004010b4]
'006c75cd    50                      push eax
'006c75ce    8d55c8                  lea edx, dword ptr [ebp-38]
'006c75d1    52                      push edx
'006c75d2    ffd6                    call esi
Set var_46 = var_43
'006c75d4    894588                  mov dword ptr [ebp-78], eax
'006c75d7    8b07                    mov eax, dword ptr [edi]
'006c75d9    57                      push edi
'006c75da    ff90fc020000            call dword ptr [eax+000002fc]
'006c75e0    50                      push eax
'006c75e1    8d4dcc                  lea ecx, dword ptr [ebp-34]
'006c75e4    51                      push ecx
'006c75e5    ffd6                    call esi
Set var_43 = Nothing
'006c75e7    8bf0                    mov esi, eax
'006c75e9    8b16                    mov edx, dword ptr [esi]
'006c75eb    8d4594                  lea eax, dword ptr [ebp-6c]
'006c75ee    50                      push eax
'006c75ef    56                      push esi
'006c75f0    ff92f0000000            call dword ptr [edx+000000f0]
'006c75f6    dbe2                    fnclex
'006c75f8    3bc3                    cmp eax, ebx
'006c75fa    7d12                    jge 6c760e
'006c75fc    68f0000000              push 000000f0
'006c7601    681cb94300              push 0043b91c
'006c7606    56                      push esi
'006c7607    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c7608    ff1580104000            call dword ptr [00401080]
'006c760e    668b5594                mov dx, word ptr [ebp-6c]
'006c7612    8b7588                  mov esi, dword ptr [ebp-78]
'006c7615    8b0e                    mov ecx, dword ptr [esi]
'006c7617    6683ea01                sub dx, 01
var_num4 = 0 - 1
'006c761b    0f80ea020000            jo 6c790b
'006c7621    52                      push edx
'006c7622    56                      push esi
'006c7623    ff91f4000000            call dword ptr [ecx+000000f4]
'006c7629    dbe2                    fnclex
'006c762b    3bc3                    cmp eax, ebx
'006c762d    7d12                    jge 6c7641
'006c762f    68f4000000              push 000000f4
'006c7634    681cb94300              push 0043b91c
'006c7639    56                      push esi
'006c763a    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c763b    ff1580104000            call dword ptr [00401080]
'006c7641    8d45c8                  lea eax, dword ptr [ebp-38]
'006c7644    50                      push eax
'006c7645    8d4dcc                  lea ecx, dword ptr [ebp-34]
'006c7648    51                      push ecx
'006c7649    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'006c764b    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_43, var_46)
'006c7651    8b17                    mov edx, dword ptr [edi]
'006c7653    83c40c                  add esp, 0c
'006c7656    57                      push edi
'006c7657    ff92fc020000            call dword ptr [edx+000002fc]
'006c765d    50                      push eax
'006c765e    8d45cc                  lea eax, dword ptr [ebp-34]
'006c7661    50                      push eax

' *** Reference to "__vbaObjSet"
'006c7662    ff15b4104000            call dword ptr [004010b4]
Set var_43 = 
'006c7668    8d55e4                  lea edx, dword ptr [ebp-1c]
'006c766b    8bf0                    mov esi, eax
'006c766d    8b0e                    mov ecx, dword ptr [esi]
'006c766f    52                      push edx
'006c7670    56                      push esi
'006c7671    ff91a8000000            call dword ptr [ecx+000000a8]
'006c7677    dbe2                    fnclex
'006c7679    3bc3                    cmp eax, ebx
'006c767b    7d12                    jge 6c768f

If (var_43 < 0) Then
'006c767d    68a8000000              push 000000a8
'006c7682    681cb94300              push 0043b91c
'006c7687    56                      push esi
'006c7688    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c7689    ff1580104000            call dword ptr [00401080]
    
End If
'006c768f    8b45e4                  mov eax, dword ptr [ebp-1c]
'006c7692    50                      push eax
'006c7693    68cc134300              push 004313cc

' *** Reference to "__vbaStrCmp"
'006c7698    ff1530114000            call dword ptr [00401130]
'006c769e    8bf0                    mov esi, eax
'006c76a0    f7de                    neg esi
'006c76a2    1bf6                    sbb esi, esi
'006c76a4    f7de                    neg esi
'006c76a6    8d4de4                  lea ecx, dword ptr [ebp-1c]
'006c76a9    f7de                    neg esi

' *** Reference to "__vbaFreeStr"
'006c76ab    ff150c134000            call dword ptr [0040130c]
'#Cleanup(var_40)
'006c76b1    8d4dcc                  lea ecx, dword ptr [ebp-34]

' *** Reference to "__vbaFreeObj"
'006c76b4    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_43)
'006c76ba    663bf3                  cmp si, bx
'006c76bd    0f84e1010000            je 6c78a4

If (((vbNullString) <> (vbNullChar))) Then
'006c76c3    8b0f                    mov ecx, dword ptr [edi]
'006c76c5    57                      push edi
'006c76c6    ff91fc020000            call dword ptr [ecx+000002fc]
'006c76cc    50                      push eax
'006c76cd    8d55cc                  lea edx, dword ptr [ebp-34]
'006c76d0    52                      push edx

' *** Reference to "__vbaObjSet"
'006c76d1    ff15b4104000            call dword ptr [004010b4]
    Set var_43 = ((vbNullString) [##] (vbNullChar))
'006c76d7    8d4de4                  lea ecx, dword ptr [ebp-1c]
'006c76da    8bf0                    mov esi, eax
'006c76dc    8b06                    mov eax, dword ptr [esi]
'006c76de    51                      push ecx
'006c76df    56                      push esi
'006c76e0    ff90a8000000            call dword ptr [eax+000000a8]
'006c76e6    dbe2                    fnclex
'006c76e8    3bc3                    cmp eax, ebx
'006c76ea    7d12                    jge 6c76fe
    
    If (    0 < 0) Then
'006c76ec    68a8000000              push 000000a8
'006c76f1    681cb94300              push 0043b91c
'006c76f6    56                      push esi
'006c76f7    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c76f8    ff1580104000            call dword ptr [00401080]
    
End If
'006c76fe    8b55e4                  mov edx, dword ptr [ebp-1c]

' *** Reference to "__vbaStrMove"
'006c7701    8b35d0124000            mov esi, dword ptr [004012d0]
'006c7707    8d4de0                  lea ecx, dword ptr [ebp-20]
'006c770a    895de4                  mov dword ptr [ebp-1c], ebx
'006c770d    ffd6                    call esi
'006c770f    8d55e0                  lea edx, dword ptr [ebp-20]
'006c7712    52                      push edx
'006c7713    e8d8c4e2ff              call 4f3bf0
Call sub_4F3BF0()
'006c7718    8bd0                    mov edx, eax
'006c771a    8d4dd0                  lea ecx, dword ptr [ebp-30]
'006c771d    ffd6                    call esi
'006c771f    8b55d0                  mov edx, dword ptr [ebp-30]
'006c7722    899574ffffff            mov dword ptr [ebp+ffffff74], edx
'006c7728    8b154cb07200            mov edx, dword ptr [0072b04c]
'006c772e    895dd0                  mov dword ptr [ebp-30], ebx
'006c7731    8b1a                    mov ebx, dword ptr [edx]
'006c7733    8d55c8                  lea edx, dword ptr [ebp-38]
'006c7736    52                      push edx
'006c7737    83ec10                  sub esp, 10
'006c773a    8bd4                    mov edx, esp
'006c773c    b90a000000              mov ecx, 0000000a
'006c7741    890a                    mov dword ptr [edx], ecx
'006c7743    894da8                  mov dword ptr [ebp-58], ecx
'006c7746    8b4d9c                  mov ecx, dword ptr [ebp-64]
'006c7749    894a04                  mov dword ptr [edx+04], ecx
'006c774c    b804000280              mov eax, 80020004
'006c7751    894208                  mov dword ptr [edx+08], eax
'006c7754    8945b0                  mov dword ptr [ebp-50], eax
'006c7757    8b45a4                  mov eax, dword ptr [ebp-5c]
'006c775a    89420c                  mov dword ptr [edx+0c], eax
'006c775d    8b55a8                  mov edx, dword ptr [ebp-58]
'006c7760    8b45ac                  mov eax, dword ptr [ebp-54]
'006c7763    83ec10                  sub esp, 10
'006c7766    8bcc                    mov ecx, esp
'006c7768    8911                    mov dword ptr [ecx], edx
'006c776a    8b55b0                  mov edx, dword ptr [ebp-50]
'006c776d    894104                  mov dword ptr [ecx+04], eax
'006c7770    8b45b4                  mov eax, dword ptr [ebp-4c]
'006c7773    895108                  mov dword ptr [ecx+08], edx
'006c7776    8b55bc                  mov edx, dword ptr [ebp-44]
'006c7779    89410c                  mov dword ptr [ecx+0c], eax
'006c777c    83ec10                  sub esp, 10
'006c777f    8bcc                    mov ecx, esp
'006c7781    b803000000              mov eax, 00000003
'006c7786    8901                    mov dword ptr [ecx], eax
'006c7788    895104                  mov dword ptr [ecx+04], edx
'006c778b    8b55c4                  mov edx, dword ptr [ebp-3c]
'006c778e    c745c002000000          mov dword ptr [ebp-40], 00000002
'006c7795    8b45c0                  mov eax, dword ptr [ebp-40]
'006c7798    894108                  mov dword ptr [ecx+08], eax
'006c779b    89510c                  mov dword ptr [ecx+0c], edx
'006c779e    8b9574ffffff            mov edx, dword ptr [ebp+ffffff74]
'006c77a4    68c0384500              push 004538c0
'006c77a9    8d4ddc                  lea ecx, dword ptr [ebp-24]
'006c77ac    ffd6                    call esi
'006c77ae    50                      push eax

' *** Reference to "__vbaStrCat"
'006c77af    ff1570104000            call dword ptr [00401070]
var_49 = ("select nom from sort where nom='") & (vbNullString)
'006c77b5    8bd0                    mov edx, eax
'006c77b7    8d4dd8                  lea ecx, dword ptr [ebp-28]
'006c77ba    ffd6                    call esi
'006c77bc    50                      push eax
'006c77bd    6854a44300              push 0043a454

' *** Reference to "__vbaStrCat"
'006c77c2    ff1570104000            call dword ptr [00401070]
var_63 = (var_49) & ("'")
'006c77c8    8bd0                    mov edx, eax
'006c77ca    8d4dd4                  lea ecx, dword ptr [ebp-2c]
'006c77cd    ffd6                    call esi
'006c77cf    50                      push eax
'006c77d0    a14cb07200              mov ax, word ptr [0072b04c]
'006c77d5    50                      push eax
'006c77d6    ff93bc000000            call dword ptr [ebx+000000bc]
'006c77dc    dbe2                    fnclex
'006c77de    33db                    xor ebx, ebx
var_num2 = Empty
'006c77e0    3bc3                    cmp eax, ebx
'006c77e2    7d1c                    jge 6c7800

If (0 < -256 - 12) Then
'006c77e4    8b0d4cb07200            mov ecx, dword ptr [0072b04c]

' *** Reference to "__vbaHresultCheckObj"
'006c77ea    8b3580104000            mov esi, dword ptr [00401080]
'006c77f0    68bc000000              push 000000bc
'006c77f5    68301f4300              push 00431f30
'006c77fa    51                      push ecx
'006c77fb    50                      push eax
'006c77fc    ffd6                    call esi
'006c77fe    eb06                    jmp 6c7806
    
End If

' *** Reference to "__vbaHresultCheckObj"
'006c7800    8b3580104000            mov esi, dword ptr [00401080]
'006c7806    8b45c8                  mov eax, dword ptr [ebp-38]
'006c7809    50                      push eax
'006c780a    8d55e8                  lea edx, dword ptr [ebp-18]
'006c780d    52                      push edx
'006c780e    895dc8                  mov dword ptr [ebp-38], ebx

' *** Reference to "__vbaObjSet"
'006c7811    ff15b4104000            call dword ptr [004010b4]
Set var_41 = var_46
'006c7817    8d45d0                  lea eax, dword ptr [ebp-30]
'006c781a    50                      push eax
'006c781b    8d4dd4                  lea ecx, dword ptr [ebp-2c]
'006c781e    51                      push ecx
'006c781f    8d55d8                  lea edx, dword ptr [ebp-28]
'006c7822    52                      push edx
'006c7823    8d45dc                  lea eax, dword ptr [ebp-24]
'006c7826    50                      push eax
'006c7827    8d4de0                  lea ecx, dword ptr [ebp-20]
'006c782a    51                      push ecx
'006c782b    6a05                    push 05

' *** Reference to "__vbaFreeStrList"
'006c782d    ff155c124000            call dword ptr [0040125c]
'#Cleanup( 0, 0, -4500, -4504, 0)
'006c7833    83c418                  add esp, 18
'006c7836    8d4dcc                  lea ecx, dword ptr [ebp-34]

' *** Reference to "__vbaFreeObj"
'006c7839    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_43)
'006c783f    8b45e8                  mov eax, dword ptr [ebp-18]
'006c7842    8b10                    mov edx, dword ptr [eax]
'006c7844    8d4d94                  lea ecx, dword ptr [ebp-6c]
'006c7847    51                      push ecx
'006c7848    50                      push eax
'006c7849    ff5234                  call dword ptr [edx+34]
'006c784c    dbe2                    fnclex
'006c784e    3bc3                    cmp eax, ebx
'006c7850    7d0e                    jge 6c7860

If (var_41 < -256 - 12) Then
'006c7852    8b55e8                  mov edx, dword ptr [ebp-18]
'006c7855    6a34                    push 34
'006c7857    6830314300              push 00433130
'006c785c    52                      push edx
'006c785d    50                      push eax
'006c785e    ffd6                    call esi
    
End If
'006c7860    66395d94                cmp word ptr [ebp-6c], bx
'006c7864    743e                    je 6c78a4

If (0 <> -256 - 12) Then
'006c7866    8b45e8                  mov eax, dword ptr [ebp-18]
'006c7869    8b08                    mov ecx, dword ptr [eax]
'006c786b    50                      push eax
'006c786c    ff91c4000000            call dword ptr [ecx+000000c4]
'006c7872    dbe2                    fnclex
'006c7874    3bc3                    cmp eax, ebx
'006c7876    7d11                    jge 6c7889
    
    If (    var_41 < -256 - 12) Then
'006c7878    8b55e8                  mov edx, dword ptr [ebp-18]
'006c787b    68c4000000              push 000000c4
'006c7880    6830314300              push 00433130
'006c7885    52                      push edx
'006c7886    50                      push eax
'006c7887    ffd6                    call esi
    
End If
'006c7889    8b07                    mov eax, dword ptr [edi]
'006c788b    57                      push edi
'006c788c    ff903c070000            call dword ptr [eax+0000073c]
'006c7892    3bc3                    cmp eax, ebx
'006c7894    7d0e                    jge 6c78a4

If (frmDescriptifSort < -256 - 12) Then
'006c7896    683c070000              push 0000073c
'006c789b    6800144300              push 00431400
'006c78a0    57                      push edi
'006c78a1    50                      push eax
'006c78a2    ffd6                    call esi
    
End If
'006c78a4    895dfc                  mov dword ptr [ebp-04], ebx
'006c78a7    68ec786c00              push 006c78ec
'006c78ac    eb34                    jmp 6c78e2
'006c78ae    8d4dd0                  lea ecx, dword ptr [ebp-30]
'006c78b1    51                      push ecx
'006c78b2    8d55d4                  lea edx, dword ptr [ebp-2c]
'006c78b5    52                      push edx
'006c78b6    8d45d8                  lea eax, dword ptr [ebp-28]
'006c78b9    50                      push eax
'006c78ba    8d4ddc                  lea ecx, dword ptr [ebp-24]
'006c78bd    51                      push ecx
'006c78be    8d55e0                  lea edx, dword ptr [ebp-20]
'006c78c1    52                      push edx
'006c78c2    8d45e4                  lea eax, dword ptr [ebp-1c]
'006c78c5    50                      push eax
'006c78c6    6a06                    push 06

' *** Reference to "__vbaFreeStrList"
'006c78c8    ff155c124000            call dword ptr [0040125c]
'#Cleanup( 0, 0, 0, -4500, -4504, 0)
'006c78ce    8d4dc8                  lea ecx, dword ptr [ebp-38]
'006c78d1    51                      push ecx
'006c78d2    8d55cc                  lea edx, dword ptr [ebp-34]
'006c78d5    52                      push edx
'006c78d6    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'006c78d8    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_43, var_46)
'006c78de    83c428                  add esp, 28
'006c78e1    c3                      ret
'006c78e2    8d4de8                  lea ecx, dword ptr [ebp-18]

' *** Reference to "__vbaFreeObj"
'006c78e5    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_41)
'006c78eb    c3                      ret
'006c78ec    8b4508                  mov eax, dword ptr [ebp+08]
'006c78ef    8b08                    mov ecx, dword ptr [eax]
'006c78f1    50                      push eax
'006c78f2    ff5108                  call dword ptr [ecx+08]
'006c78f5    8b45fc                  mov eax, dword ptr [ebp-04]
'006c78f8    8b4dec                  mov ecx, dword ptr [ebp-14]
'006c78fb    5f                      pop edi
'006c78fc    5e                      pop esi
    'Reference to '__except_list'
'006c78fd    64890d00000000          mov dword ptr fs:[00000000], ecx
'006c7904    5b                      pop ebx
'006c7905    8be5                    mov esp, ebp
'006c7907    5d                      pop ebp
'006c7908    c20400                  ret 0004


End Sub


'Event for Form
Private Sub Form_Load()
'006c7e60    55                      push ebp
'006c7e61    8bec                    mov ebp, esp
'006c7e63    83ec0c                  sub esp, 0c

' *** Reference to "__vbaExceptHandler"
'006c7e66    6866724000              push 00407266
'006c7e6b    64a100000000            mov ax, word ptr fs:[00000000]
'006c7e71    50                      push eax
    'Reference to '__except_list'
'006c7e72    64892500000000          mov dword ptr fs:[00000000], esp
'006c7e79    81ece4000000            sub esp, 000000e4
'006c7e7f    53                      push ebx
'006c7e80    56                      push esi
'006c7e81    57                      push edi
'006c7e82    8965f4                  mov dword ptr [ebp-0c], esp
'006c7e85    c745f888674000          mov dword ptr [ebp-08], 00406788
'006c7e8c    8b7d08                  mov edi, dword ptr [ebp+08]
'006c7e8f    8bc7                    mov eax, edi
'006c7e91    83e001                  and eax, 01
'006c7e94    8945fc                  mov dword ptr [ebp-04], eax
'006c7e97    83e7fe                  and edi, fffffffe
'006c7e9a    8b0f                    mov ecx, dword ptr [edi]
'006c7e9c    57                      push edi
'006c7e9d    897d08                  mov dword ptr [ebp+08], edi
'006c7ea0    ff5104                  call dword ptr [ecx+04]
'006c7ea3    33db                    xor ebx, ebx
'006c7ea5    8d45b4                  lea eax, dword ptr [ebp-4c]
'006c7ea8    50                      push eax
'006c7ea9    8d4dc4                  lea ecx, dword ptr [ebp-3c]
'006c7eac    8d5740                  lea edx, dword ptr [edi+40]
'006c7eaf    895db4                  mov dword ptr [ebp-4c], ebx
'006c7eb2    51                      push ecx
'006c7eb3    895de4                  mov dword ptr [ebp-1c], ebx
'006c7eb6    895de0                  mov dword ptr [ebp-20], ebx
'006c7eb9    895ddc                  mov dword ptr [ebp-24], ebx
'006c7ebc    895dd8                  mov dword ptr [ebp-28], ebx
'006c7ebf    895dd4                  mov dword ptr [ebp-2c], ebx
'006c7ec2    895dc4                  mov dword ptr [ebp-3c], ebx
'006c7ec5    895da4                  mov dword ptr [ebp-5c], ebx
'006c7ec8    895d94                  mov dword ptr [ebp-6c], ebx
'006c7ecb    899d74ffffff            mov dword ptr [ebp+ffffff74], ebx
'006c7ed1    899d54ffffff            mov dword ptr [ebp+ffffff54], ebx
'006c7ed7    899d50ffffff            mov dword ptr [ebp+ffffff50], ebx
'006c7edd    899d24ffffff            mov dword ptr [ebp+ffffff24], ebx
'006c7ee3    8955bc                  mov dword ptr [ebp-44], edx
'006c7ee6    c745b408400000          mov dword ptr [ebp-4c], 00004008

' *** Reference to "rtcUpperCaseVar"
'006c7eed    ff152c114000            call dword ptr [0040112c]
'006c7ef3    8d55c4                  lea edx, dword ptr [ebp-3c]
'006c7ef6    8d8d24ffffff            lea ecx, dword ptr [ebp+ffffff24]

' *** Reference to "__vbaVarMove"
'006c7efc    ff151c104000            call dword ptr [0040101c]
var_432 = (UCase(arg_7))

' *** Reference to "__vbaVarTstEq"
'006c7f02    8b353c114000            mov esi, dword ptr [0040113c]
'006c7f08    8d9524ffffff            lea edx, dword ptr [ebp+ffffff24]
'006c7f0e    52                      push edx
'006c7f0f    8d45b4                  lea eax, dword ptr [ebp-4c]
'006c7f12    50                      push eax
'006c7f13    c745bc08394500          mov dword ptr [ebp-44], 00453908
'006c7f1a    c745b408800000          mov dword ptr [ebp-4c], 00008008
'006c7f21    ffd6                    call esi
'006c7f23    6685c0                  test eax, eax
'006c7f26    0f848a020000            je 6c81b6

If (((var_432) = ("VSSORTCLASSE"))) Then
'006c7f2c    a174b17200              mov ax, word ptr [0072b174]
'006c7f31    3bc3                    cmp eax, ebx
'006c7f33    7515                    jne 6c7f4a
    
    If (    0 = 0) Then
'006c7f35    6874b17200              push 0072b174
'006c7f3a    6890c04100              push 0041c090

' *** Reference to "__vbaNew2"
'006c7f3f    ff152c124000            call dword ptr [0040122c]
    Dim var_77 As New frmFichePerso
'006c7f45    a174b17200              mov ax, word ptr [0072b174]
    
End If
'006c7f4a    8b08                    mov ecx, dword ptr [eax]
'006c7f4c    50                      push eax
'006c7f4d    ff9120060000            call dword ptr [ecx+00000620]
'006c7f53    50                      push eax
'006c7f54    8d55dc                  lea edx, dword ptr [ebp-24]
'006c7f57    52                      push edx

' *** Reference to "__vbaObjSet"
'006c7f58    ff15b4104000            call dword ptr [004010b4]
Set var_39 = Nothing
'006c7f5e    33d2                    xor edx, edx
'006c7f60    668b573c                mov dx, word ptr [edi+3c]
'006c7f64    8d4dd8                  lea ecx, dword ptr [ebp-28]
'006c7f67    51                      push ecx
'006c7f68    8bf0                    mov esi, eax
'006c7f6a    8b06                    mov eax, dword ptr [esi]
'006c7f6c    52                      push edx
'006c7f6d    56                      push esi
'006c7f6e    ff5040                  call dword ptr [eax+40]
'006c7f71    dbe2                    fnclex
'006c7f73    3bc3                    cmp eax, ebx
'006c7f75    7d0f                    jge 6c7f86

If (-256 - 12 < 0) Then
'006c7f77    6a40                    push 40
'006c7f79    686c384300              push 0043386c
'006c7f7e    56                      push esi
'006c7f7f    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c7f80    ff1580104000            call dword ptr [00401080]
    
End If
'006c7f86    8b45d8                  mov eax, dword ptr [ebp-28]
'006c7f89    53                      push ebx
'006c7f8a    6a07                    push 07
'006c7f8c    50                      push eax
'006c7f8d    8d4dc4                  lea ecx, dword ptr [ebp-3c]
'006c7f90    51                      push ecx

' *** Reference to "__vbaLateIdCallLd"
'006c7f91    ff157c114000            call dword ptr [0040117c]
var_9 = 0.UNK_-256 - 12_7
'006c7f97    83c410                  add esp, 10
'006c7f9a    50                      push eax

' *** Reference to "__vbaI4Var"
'006c7f9b    ff157c124000            call dword ptr [0040127c]
'006c7fa1    8bc8                    mov ecx, eax
'006c7fa3    83e901                  sub ecx, 01
var_num3 = CLng(var_9) - 1
'006c7fa6    0f805b080000            jo 6c8807

' *** Reference to "__vbaI2I4"
'006c7fac    ff1550114000            call dword ptr [00401150]
'006c7fb2    8d55d8                  lea edx, dword ptr [ebp-28]
'006c7fb5    52                      push edx
'006c7fb6    33f6                    xor esi, esi
var_num8 = Empty
'006c7fb8    c78520ffffff01000000    mov dword ptr [ebp+ffffff20], 00000001
'006c7fc2    8975e8                  mov dword ptr [ebp-18], esi
'006c7fc5    89851cffffff            mov dword ptr [ebp+ffffff1c], eax
'006c7fcb    8d45dc                  lea eax, dword ptr [ebp-24]
'006c7fce    50                      push eax
'006c7fcf    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'006c7fd1    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_39, var_45)
'006c7fd7    83c40c                  add esp, 0c
'006c7fda    8d4dc4                  lea ecx, dword ptr [ebp-3c]

' *** Reference to "__vbaFreeVar"
'006c7fdd    ff1528104000            call dword ptr [00401028]
'#Cleanup(var_9)
'006c7fe3    663bb51cffffff          cmp si, word ptr [ebp+ffffff1c]
'006c7fea    0f8f91010000            jg 6c8181

Do While (var_39 <= WORD PTR [EBP+FFFFFF1C])
'006c7ff0    a174b17200              mov ax, word ptr [0072b174]
'006c7ff5    85c0                    test ax, ax
'006c7ff7    895dbc                  mov dword ptr [ebp-44], ebx
'006c7ffa    bb02000000              mov ebx, 00000002
'006c7fff    c745b403000000          mov dword ptr [ebp-4c], 00000003
'006c8006    6689759c                mov word ptr [ebp-64], si
'006c800a    899d74ffffff            mov dword ptr [ebp+ffffff74], ebx
'006c8010    7515                    jne 6c8027
'006c8012    6874b17200              push 0072b174
'006c8017    6890c04100              push 0041c090

' *** Reference to "__vbaNew2"
'006c801c    ff152c124000            call dword ptr [0040122c]
    Set var_77 = New frmFichePerso
'006c8022    a174b17200              mov ax, word ptr [0072b174]
'006c8027    8b08                    mov ecx, dword ptr [eax]
'006c8029    50                      push eax
'006c802a    ff9120060000            call dword ptr [ecx+00000620]
'006c8030    50                      push eax
'006c8031    8d55dc                  lea edx, dword ptr [ebp-24]
'006c8034    52                      push edx

' *** Reference to "__vbaObjSet"
'006c8035    ff15b4104000            call dword ptr [004010b4]
    Set var_39 = var_77
'006c803b    33d2                    xor edx, edx
'006c803d    668b573c                mov dx, word ptr [edi+3c]
'006c8041    8d4dd8                  lea ecx, dword ptr [ebp-28]
'006c8044    51                      push ecx
'006c8045    8bf0                    mov esi, eax
'006c8047    8b06                    mov eax, dword ptr [esi]
'006c8049    52                      push edx
'006c804a    56                      push esi
'006c804b    ff5040                  call dword ptr [eax+40]
'006c804e    dbe2                    fnclex
'006c8050    85c0                    test ax, ax
'006c8052    7d0f                    jge 6c8063
'006c8054    6a40                    push 40
'006c8056    686c384300              push 0043386c
'006c805b    56                      push esi
'006c805c    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c805d    ff1580104000            call dword ptr [00401080]
'006c8063    8b4db4                  mov ecx, dword ptr [ebp-4c]
'006c8066    8b55b8                  mov edx, dword ptr [ebp-48]
'006c8069    83ec10                  sub esp, 10
'006c806c    8bc4                    mov eax, esp
'006c806e    8908                    mov dword ptr [eax], ecx
'006c8070    8b4dbc                  mov ecx, dword ptr [ebp-44]
'006c8073    895004                  mov dword ptr [eax+04], edx
'006c8076    8b55c0                  mov edx, dword ptr [ebp-40]
'006c8079    894808                  mov dword ptr [eax+08], ecx
'006c807c    8b4d98                  mov ecx, dword ptr [ebp-68]
'006c807f    89500c                  mov dword ptr [eax+0c], edx
'006c8082    8b559c                  mov edx, dword ptr [ebp-64]
'006c8085    83ec10                  sub esp, 10
'006c8088    8bc4                    mov eax, esp
'006c808a    8918                    mov dword ptr [eax], ebx
'006c808c    894804                  mov dword ptr [eax+04], ecx
'006c808f    8b4da0                  mov ecx, dword ptr [ebp-60]
'006c8092    895008                  mov dword ptr [eax+08], edx
'006c8095    89480c                  mov dword ptr [eax+0c], ecx
'006c8098    8b8574ffffff            mov eax, dword ptr [ebp+ffffff74]
'006c809e    8b8d78ffffff            mov ecx, dword ptr [ebp+ffffff78]
'006c80a4    83ec10                  sub esp, 10
'006c80a7    8bd4                    mov edx, esp
'006c80a9    8902                    mov dword ptr [edx], eax
'006c80ab    894a04                  mov dword ptr [edx+04], ecx
'006c80ae    8b4dd8                  mov ecx, dword ptr [ebp-28]
'006c80b1    33c0                    xor eax, eax
    var_num1 = Empty
'006c80b3    894208                  mov dword ptr [edx+08], eax
'006c80b6    8b4580                  mov eax, dword ptr [ebp-80]
'006c80b9    6a03                    push 03
'006c80bb    689e000000              push 0000009e
'006c80c0    89420c                  mov dword ptr [edx+0c], eax
'006c80c3    51                      push ecx
'006c80c4    8d55c4                  lea edx, dword ptr [ebp-3c]
'006c80c7    52                      push edx

' *** Reference to "__vbaLateIdCallLd"
'006c80c8    ff157c114000            call dword ptr [0040117c]
    var_9 = 0.UNK_-256 - 12_158
'006c80ce    8b07                    mov eax, dword ptr [edi]
'006c80d0    83c440                  add esp, 40
'006c80d3    57                      push edi
'006c80d4    ff90fc020000            call dword ptr [eax+000002fc]
'006c80da    50                      push eax
'006c80db    8d4dd4                  lea ecx, dword ptr [ebp-2c]
'006c80de    51                      push ecx

' *** Reference to "__vbaObjSet"
'006c80df    ff15b4104000            call dword ptr [004010b4]
    Set var_44 = Nothing
'006c80e5    83ec10                  sub esp, 10
'006c80e8    8bd4                    mov edx, esp
'006c80ea    b90a000000              mov ecx, 0000000a
'006c80ef    890a                    mov dword ptr [edx], ecx
'006c80f1    8b8d58ffffff            mov ecx, dword ptr [ebp+ffffff58]
'006c80f7    8bf0                    mov esi, eax
'006c80f9    8b1e                    mov ebx, dword ptr [esi]
'006c80fb    894a04                  mov dword ptr [edx+04], ecx
'006c80fe    b804000280              mov eax, 80020004
'006c8103    894208                  mov dword ptr [edx+08], eax
'006c8106    8b8560ffffff            mov eax, dword ptr [ebp+ffffff60]
'006c810c    8d4dc4                  lea ecx, dword ptr [ebp-3c]
'006c810f    89420c                  mov dword ptr [edx+0c], eax
'006c8112    51                      push ecx
'006c8113    8d55e0                  lea edx, dword ptr [ebp-20]
'006c8116    52                      push edx

' *** Reference to "__vbaStrVarVal"
'006c8117    ff15fc114000            call dword ptr [004011fc]
'006c811d    50                      push eax
'006c811e    56                      push esi
'006c811f    ff93ec010000            call dword ptr [ebx+000001ec]
'006c8125    dbe2                    fnclex
'006c8127    85c0                    test ax, ax
'006c8129    7d12                    jge 6c813d
'006c812b    68ec010000              push 000001ec
'006c8130    681cb94300              push 0043b91c
'006c8135    56                      push esi
'006c8136    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c8137    ff1580104000            call dword ptr [00401080]
'006c813d    8d4de0                  lea ecx, dword ptr [ebp-20]

' *** Reference to "__vbaFreeStr"
'006c8140    ff150c134000            call dword ptr [0040130c]
    '#Cleanup(var_3)
'006c8146    8d45d4                  lea eax, dword ptr [ebp-2c]
'006c8149    50                      push eax
'006c814a    8d4dd8                  lea ecx, dword ptr [ebp-28]
'006c814d    51                      push ecx
'006c814e    8d55dc                  lea edx, dword ptr [ebp-24]
'006c8151    52                      push edx
'006c8152    6a03                    push 03

' *** Reference to "__vbaFreeObjList"
'006c8154    ff154c104000            call dword ptr [0040104c]
    '#Cleanup( var_39, var_45, var_44)
'006c815a    83c410                  add esp, 10
'006c815d    8d4dc4                  lea ecx, dword ptr [ebp-3c]

' *** Reference to "__vbaFreeVar"
'006c8160    ff1528104000            call dword ptr [00401028]
    '#Cleanup(var_9)
'006c8166    668bb520ffffff          mov si, word ptr [ebp+ffffff20]
'006c816d    660375e8                add si, word ptr [ebp-18]
    var_num8 = var_72 + var_39
'006c8171    0f8090060000            jo 6c8807
'006c8177    33db                    xor ebx, ebx
    var_num2 = Empty
'006c8179    8975e8                  mov dword ptr [ebp-18], esi
'006c817c    e962feffff              jmp 6c7fe3
    
Loop
'006c8181    8b07                    mov eax, dword ptr [edi]
'006c8183    57                      push edi
'006c8184    ff90fc020000            call dword ptr [eax+000002fc]
'006c818a    50                      push eax
'006c818b    8d4ddc                  lea ecx, dword ptr [ebp-24]
'006c818e    51                      push ecx

' *** Reference to "__vbaObjSet"
'006c818f    ff15b4104000            call dword ptr [004010b4]
Set var_39 = Nothing
'006c8195    8bf0                    mov esi, eax
'006c8197    8b16                    mov edx, dword ptr [esi]
'006c8199    33c0                    xor eax, eax
var_num1 = Empty
'006c819b    668b473a                mov ax, word ptr [edi+3a]
'006c819f    50                      push eax
'006c81a0    56                      push esi
'006c81a1    ff92f4000000            call dword ptr [edx+000000f4]
'006c81a7    dbe2                    fnclex
'006c81a9    3bc3                    cmp eax, ebx
'006c81ab    0f8dbc020000            jge 6c846d

If (0 < 0) Then
'006c81b1    e9a5020000              jmp 6c845b
    
End If
'006c81b6    8d8d24ffffff            lea ecx, dword ptr [ebp+ffffff24]
'006c81bc    51                      push ecx
'006c81bd    8d55b4                  lea edx, dword ptr [ebp-4c]
'006c81c0    52                      push edx
'006c81c1    c745bc28394500          mov dword ptr [ebp-44], 00453928
'006c81c8    c745b408800000          mov dword ptr [ebp-4c], 00008008
'006c81cf    ffd6                    call esi
'006c81d1    6685c0                  test eax, eax
'006c81d4    0f84a1020000            je 6c847b

If (((var_432) = ("VSSORTLISTE"))) Then
'006c81da    a174b17200              mov ax, word ptr [0072b174]
'006c81df    3bc3                    cmp eax, ebx
'006c81e1    7515                    jne 6c81f8
    
    If (    var_77 = 0) Then
'006c81e3    6874b17200              push 0072b174
'006c81e8    6890c04100              push 0041c090

' *** Reference to "__vbaNew2"
'006c81ed    ff152c124000            call dword ptr [0040122c]
    Set var_77 = New frmFichePerso
'006c81f3    a174b17200              mov ax, word ptr [0072b174]
    
End If
'006c81f8    8b08                    mov ecx, dword ptr [eax]
'006c81fa    50                      push eax
'006c81fb    ff9140060000            call dword ptr [ecx+00000640]
'006c8201    50                      push eax
'006c8202    8d55dc                  lea edx, dword ptr [ebp-24]
'006c8205    52                      push edx

' *** Reference to "__vbaObjSet"
'006c8206    ff15b4104000            call dword ptr [004010b4]
Set var_39 = var_77
'006c820c    33d2                    xor edx, edx
'006c820e    668b573c                mov dx, word ptr [edi+3c]
'006c8212    8d4dd8                  lea ecx, dword ptr [ebp-28]
'006c8215    51                      push ecx
'006c8216    8bf0                    mov esi, eax
'006c8218    8b06                    mov eax, dword ptr [esi]
'006c821a    52                      push edx
'006c821b    56                      push esi
'006c821c    ff5040                  call dword ptr [eax+40]
'006c821f    dbe2                    fnclex
'006c8221    3bc3                    cmp eax, ebx
'006c8223    7d0f                    jge 6c8234

If (frmFichePerso < 0) Then
'006c8225    6a40                    push 40
'006c8227    686c384300              push 0043386c
'006c822c    56                      push esi
'006c822d    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c822e    ff1580104000            call dword ptr [00401080]
    
End If
'006c8234    8b45d8                  mov eax, dword ptr [ebp-28]
'006c8237    53                      push ebx
'006c8238    6a07                    push 07
'006c823a    50                      push eax
'006c823b    8d4dc4                  lea ecx, dword ptr [ebp-3c]
'006c823e    51                      push ecx

' *** Reference to "__vbaLateIdCallLd"
'006c823f    ff157c114000            call dword ptr [0040117c]
var_9 = 0.UNK_-256 - 12_7
'006c8245    83c410                  add esp, 10
'006c8248    50                      push eax

' *** Reference to "__vbaI4Var"
'006c8249    ff157c124000            call dword ptr [0040127c]
'006c824f    8bc8                    mov ecx, eax
'006c8251    83e901                  sub ecx, 01
var_num3 = CLng(var_9) - 1
'006c8254    0f80ad050000            jo 6c8807

' *** Reference to "__vbaI2I4"
'006c825a    ff1550114000            call dword ptr [00401150]
'006c8260    8d55d8                  lea edx, dword ptr [ebp-28]
'006c8263    52                      push edx
'006c8264    33f6                    xor esi, esi
var_num8 = Empty
'006c8266    c78518ffffff01000000    mov dword ptr [ebp+ffffff18], 00000001
'006c8270    8975e8                  mov dword ptr [ebp-18], esi
'006c8273    898514ffffff            mov dword ptr [ebp+ffffff14], eax
'006c8279    8d45dc                  lea eax, dword ptr [ebp-24]
'006c827c    50                      push eax
'006c827d    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'006c827f    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_39, var_45)
'006c8285    83c40c                  add esp, 0c
'006c8288    8d4dc4                  lea ecx, dword ptr [ebp-3c]

' *** Reference to "__vbaFreeVar"
'006c828b    ff1528104000            call dword ptr [00401028]
'#Cleanup(var_9)
'006c8291    663bb514ffffff          cmp si, word ptr [ebp+ffffff14]
'006c8298    0f8f91010000            jg 6c842f

Do While (var_39 <= WORD PTR [EBP+FFFFFF14])
'006c829e    a174b17200              mov ax, word ptr [0072b174]
'006c82a3    85c0                    test ax, ax
'006c82a5    895dbc                  mov dword ptr [ebp-44], ebx
'006c82a8    bb02000000              mov ebx, 00000002
'006c82ad    c745b403000000          mov dword ptr [ebp-4c], 00000003
'006c82b4    6689759c                mov word ptr [ebp-64], si
'006c82b8    899d74ffffff            mov dword ptr [ebp+ffffff74], ebx
'006c82be    7515                    jne 6c82d5
'006c82c0    6874b17200              push 0072b174
'006c82c5    6890c04100              push 0041c090

' *** Reference to "__vbaNew2"
'006c82ca    ff152c124000            call dword ptr [0040122c]
    Set var_77 = New frmFichePerso
'006c82d0    a174b17200              mov ax, word ptr [0072b174]
'006c82d5    8b08                    mov ecx, dword ptr [eax]
'006c82d7    50                      push eax
'006c82d8    ff9140060000            call dword ptr [ecx+00000640]
'006c82de    50                      push eax
'006c82df    8d55dc                  lea edx, dword ptr [ebp-24]
'006c82e2    52                      push edx

' *** Reference to "__vbaObjSet"
'006c82e3    ff15b4104000            call dword ptr [004010b4]
    Set var_39 = var_77
'006c82e9    33d2                    xor edx, edx
'006c82eb    668b573c                mov dx, word ptr [edi+3c]
'006c82ef    8d4dd8                  lea ecx, dword ptr [ebp-28]
'006c82f2    51                      push ecx
'006c82f3    8bf0                    mov esi, eax
'006c82f5    8b06                    mov eax, dword ptr [esi]
'006c82f7    52                      push edx
'006c82f8    56                      push esi
'006c82f9    ff5040                  call dword ptr [eax+40]
'006c82fc    dbe2                    fnclex
'006c82fe    85c0                    test ax, ax
'006c8300    7d0f                    jge 6c8311
'006c8302    6a40                    push 40
'006c8304    686c384300              push 0043386c
'006c8309    56                      push esi
'006c830a    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c830b    ff1580104000            call dword ptr [00401080]
'006c8311    8b4db4                  mov ecx, dword ptr [ebp-4c]
'006c8314    8b55b8                  mov edx, dword ptr [ebp-48]
'006c8317    83ec10                  sub esp, 10
'006c831a    8bc4                    mov eax, esp
'006c831c    8908                    mov dword ptr [eax], ecx
'006c831e    8b4dbc                  mov ecx, dword ptr [ebp-44]
'006c8321    895004                  mov dword ptr [eax+04], edx
'006c8324    8b55c0                  mov edx, dword ptr [ebp-40]
'006c8327    894808                  mov dword ptr [eax+08], ecx
'006c832a    8b4d98                  mov ecx, dword ptr [ebp-68]
'006c832d    89500c                  mov dword ptr [eax+0c], edx
'006c8330    8b559c                  mov edx, dword ptr [ebp-64]
'006c8333    83ec10                  sub esp, 10
'006c8336    8bc4                    mov eax, esp
'006c8338    8918                    mov dword ptr [eax], ebx
'006c833a    894804                  mov dword ptr [eax+04], ecx
'006c833d    8b4da0                  mov ecx, dword ptr [ebp-60]
'006c8340    895008                  mov dword ptr [eax+08], edx
'006c8343    89480c                  mov dword ptr [eax+0c], ecx
'006c8346    8b8574ffffff            mov eax, dword ptr [ebp+ffffff74]
'006c834c    8b8d78ffffff            mov ecx, dword ptr [ebp+ffffff78]
'006c8352    83ec10                  sub esp, 10
'006c8355    8bd4                    mov edx, esp
'006c8357    8902                    mov dword ptr [edx], eax
'006c8359    894a04                  mov dword ptr [edx+04], ecx
'006c835c    8b4dd8                  mov ecx, dword ptr [ebp-28]
'006c835f    33c0                    xor eax, eax
    var_num1 = Empty
'006c8361    894208                  mov dword ptr [edx+08], eax
'006c8364    8b4580                  mov eax, dword ptr [ebp-80]
'006c8367    6a03                    push 03
'006c8369    689e000000              push 0000009e
'006c836e    89420c                  mov dword ptr [edx+0c], eax
'006c8371    51                      push ecx
'006c8372    8d55c4                  lea edx, dword ptr [ebp-3c]
'006c8375    52                      push edx

' *** Reference to "__vbaLateIdCallLd"
'006c8376    ff157c114000            call dword ptr [0040117c]
    var_9 = 0.UNK_-256 - 12_158
'006c837c    8b07                    mov eax, dword ptr [edi]
'006c837e    83c440                  add esp, 40
'006c8381    57                      push edi
'006c8382    ff90fc020000            call dword ptr [eax+000002fc]
'006c8388    50                      push eax
'006c8389    8d4dd4                  lea ecx, dword ptr [ebp-2c]
'006c838c    51                      push ecx

' *** Reference to "__vbaObjSet"
'006c838d    ff15b4104000            call dword ptr [004010b4]
    Set var_44 = Nothing
'006c8393    83ec10                  sub esp, 10
'006c8396    8bd4                    mov edx, esp
'006c8398    b90a000000              mov ecx, 0000000a
'006c839d    890a                    mov dword ptr [edx], ecx
'006c839f    8b8d58ffffff            mov ecx, dword ptr [ebp+ffffff58]
'006c83a5    8bf0                    mov esi, eax
'006c83a7    8b1e                    mov ebx, dword ptr [esi]
'006c83a9    894a04                  mov dword ptr [edx+04], ecx
'006c83ac    b804000280              mov eax, 80020004
'006c83b1    894208                  mov dword ptr [edx+08], eax
'006c83b4    8b8560ffffff            mov eax, dword ptr [ebp+ffffff60]
'006c83ba    8d4dc4                  lea ecx, dword ptr [ebp-3c]
'006c83bd    89420c                  mov dword ptr [edx+0c], eax
'006c83c0    51                      push ecx
'006c83c1    8d55e0                  lea edx, dword ptr [ebp-20]
'006c83c4    52                      push edx

' *** Reference to "__vbaStrVarVal"
'006c83c5    ff15fc114000            call dword ptr [004011fc]
'006c83cb    50                      push eax
'006c83cc    56                      push esi
'006c83cd    ff93ec010000            call dword ptr [ebx+000001ec]
'006c83d3    dbe2                    fnclex
'006c83d5    85c0                    test ax, ax
'006c83d7    7d12                    jge 6c83eb
'006c83d9    68ec010000              push 000001ec
'006c83de    681cb94300              push 0043b91c
'006c83e3    56                      push esi
'006c83e4    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c83e5    ff1580104000            call dword ptr [00401080]
'006c83eb    8d4de0                  lea ecx, dword ptr [ebp-20]

' *** Reference to "__vbaFreeStr"
'006c83ee    ff150c134000            call dword ptr [0040130c]
    '#Cleanup(var_3)
'006c83f4    8d45d4                  lea eax, dword ptr [ebp-2c]
'006c83f7    50                      push eax
'006c83f8    8d4dd8                  lea ecx, dword ptr [ebp-28]
'006c83fb    51                      push ecx
'006c83fc    8d55dc                  lea edx, dword ptr [ebp-24]
'006c83ff    52                      push edx
'006c8400    6a03                    push 03

' *** Reference to "__vbaFreeObjList"
'006c8402    ff154c104000            call dword ptr [0040104c]
    '#Cleanup( var_39, var_45, var_44)
'006c8408    83c410                  add esp, 10
'006c840b    8d4dc4                  lea ecx, dword ptr [ebp-3c]

' *** Reference to "__vbaFreeVar"
'006c840e    ff1528104000            call dword ptr [00401028]
    '#Cleanup(var_9)
'006c8414    668bb518ffffff          mov si, word ptr [ebp+ffffff18]
'006c841b    660375e8                add si, word ptr [ebp-18]
    var_num8 = var_849 + var_39
'006c841f    0f80e2030000            jo 6c8807
'006c8425    33db                    xor ebx, ebx
    var_num2 = Empty
'006c8427    8975e8                  mov dword ptr [ebp-18], esi
'006c842a    e962feffff              jmp 6c8291
    
Loop
'006c842f    8b07                    mov eax, dword ptr [edi]
'006c8431    57                      push edi
'006c8432    ff90fc020000            call dword ptr [eax+000002fc]
'006c8438    50                      push eax
'006c8439    8d4ddc                  lea ecx, dword ptr [ebp-24]
'006c843c    51                      push ecx

' *** Reference to "__vbaObjSet"
'006c843d    ff15b4104000            call dword ptr [004010b4]
Set var_39 = Nothing
'006c8443    8bf0                    mov esi, eax
'006c8445    8b16                    mov edx, dword ptr [esi]
'006c8447    33c0                    xor eax, eax
var_num1 = Empty
'006c8449    668b473a                mov ax, word ptr [edi+3a]
'006c844d    50                      push eax
'006c844e    56                      push esi
'006c844f    ff92f4000000            call dword ptr [edx+000000f4]
'006c8455    dbe2                    fnclex
'006c8457    3bc3                    cmp eax, ebx
'006c8459    7d12                    jge 6c846d

If (0 < 0) Then
'006c845b    68f4000000              push 000000f4
'006c8460    681cb94300              push 0043b91c
'006c8465    56                      push esi
'006c8466    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c8467    ff1580104000            call dword ptr [00401080]
    
End If
'006c846d    8d4ddc                  lea ecx, dword ptr [ebp-24]

' *** Reference to "__vbaFreeObj"
'006c8470    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_39)
'006c8476    e904030000              jmp 6c877f

'ERROR: Two many next close:
End If
'006c847b    8b0f                    mov ecx, dword ptr [edi]
'006c847d    57                      push edi
'006c847e    ff91fc020000            call dword ptr [ecx+000002fc]
'006c8484    50                      push eax
'006c8485    8d55dc                  lea edx, dword ptr [ebp-24]
'006c8488    52                      push edx

' *** Reference to "__vbaObjSet"
'006c8489    ff15b4104000            call dword ptr [004010b4]
Set var_39 = ((var_432) = ("VSSORTLISTE"))
'006c848f    8bf0                    mov esi, eax
'006c8491    8b06                    mov eax, dword ptr [esi]
'006c8493    56                      push esi
'006c8494    ff90e8010000            call dword ptr [eax+000001e8]
'006c849a    dbe2                    fnclex
'006c849c    3bc3                    cmp eax, ebx
'006c849e    7d12                    jge 6c84b2

If (0 < 0) Then
'006c84a0    68e8010000              push 000001e8
'006c84a5    681cb94300              push 0043b91c
'006c84aa    56                      push esi
'006c84ab    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c84ac    ff1580104000            call dword ptr [00401080]
    
End If
'006c84b2    8d4ddc                  lea ecx, dword ptr [ebp-24]

' *** Reference to "__vbaFreeObj"
'006c84b5    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_39)
'006c84bb    8d5ddc                  lea ebx, dword ptr [ebp-24]
'006c84be    53                      push ebx
'006c84bf    83ec10                  sub esp, 10
'006c84c2    8bdc                    mov ebx, esp
'006c84c4    b90a000000              mov ecx, 0000000a
'006c84c9    890b                    mov dword ptr [ebx], ecx
'006c84cb    894da4                  mov dword ptr [ebp-5c], ecx
'006c84ce    8b4d98                  mov ecx, dword ptr [ebp-68]
'006c84d1    894b04                  mov dword ptr [ebx+04], ecx
'006c84d4    8b354cb07200            mov esi, dword ptr [0072b04c]
'006c84da    b804000280              mov eax, 80020004
'006c84df    894308                  mov dword ptr [ebx+08], eax
'006c84e2    8bd0                    mov edx, eax
'006c84e4    8b45a0                  mov eax, dword ptr [ebp-60]
'006c84e7    89430c                  mov dword ptr [ebx+0c], eax
'006c84ea    8b45a4                  mov eax, dword ptr [ebp-5c]
'006c84ed    83ec10                  sub esp, 10
'006c84f0    8bcc                    mov ecx, esp
'006c84f2    8901                    mov dword ptr [ecx], eax
'006c84f4    8b45a8                  mov eax, dword ptr [ebp-58]
'006c84f7    894104                  mov dword ptr [ecx+04], eax
'006c84fa    895108                  mov dword ptr [ecx+08], edx
'006c84fd    8b55b0                  mov edx, dword ptr [ebp-50]
'006c8500    89510c                  mov dword ptr [ecx+0c], edx
'006c8503    8b55b8                  mov edx, dword ptr [ebp-48]
'006c8506    83ec10                  sub esp, 10
'006c8509    c745b403000000          mov dword ptr [ebp-4c], 00000003
'006c8510    8b4db4                  mov ecx, dword ptr [ebp-4c]
'006c8513    8bc4                    mov eax, esp
'006c8515    8908                    mov dword ptr [eax], ecx
'006c8517    895004                  mov dword ptr [eax+04], edx
'006c851a    8b55c0                  mov edx, dword ptr [ebp-40]
'006c851d    c745bc02000000          mov dword ptr [ebp-44], 00000002
'006c8524    8b4dbc                  mov ecx, dword ptr [ebp-44]
'006c8527    8b36                    mov esi, dword ptr [esi]
'006c8529    894808                  mov dword ptr [eax+08], ecx
'006c852c    89500c                  mov dword ptr [eax+0c], edx
'006c852f    a14cb07200              mov ax, word ptr [0072b04c]
'006c8534    6844394500              push 00453944
'006c8539    50                      push eax
'006c853a    ff96bc000000            call dword ptr [esi+000000bc]
'006c8540    dbe2                    fnclex
'006c8542    85c0                    test ax, ax
'006c8544    7d1c                    jge 6c8562
'006c8546    8b0d4cb07200            mov ecx, dword ptr [0072b04c]

' *** Reference to "__vbaHresultCheckObj"
'006c854c    8b1d80104000            mov ebx, dword ptr [00401080]
'006c8552    68bc000000              push 000000bc
'006c8557    68301f4300              push 00431f30
'006c855c    51                      push ecx
'006c855d    50                      push eax
'006c855e    ffd3                    call ebx
'006c8560    eb06                    jmp 6c8568

' *** Reference to "__vbaHresultCheckObj"
'006c8562    8b1d80104000            mov ebx, dword ptr [00401080]
'006c8568    8b45dc                  mov eax, dword ptr [ebp-24]
'006c856b    50                      push eax
'006c856c    8d55e4                  lea edx, dword ptr [ebp-1c]
'006c856f    52                      push edx
'006c8570    c745dc00000000          mov dword ptr [ebp-24], 00000000

' *** Reference to "__vbaObjSet"
'006c8577    ff15b4104000            call dword ptr [004010b4]
Set var_40 = var_39
'006c857d    8b45e4                  mov eax, dword ptr [ebp-1c]
'006c8580    8b08                    mov ecx, dword ptr [eax]
'006c8582    8d9550ffffff            lea edx, dword ptr [ebp+ffffff50]
'006c8588    52                      push edx
'006c8589    50                      push eax
'006c858a    ff5134                  call dword ptr [ecx+34]
'006c858d    dbe2                    fnclex
'006c858f    85c0                    test ax, ax
'006c8591    7d0e                    jge 6c85a1
'006c8593    8b4de4                  mov ecx, dword ptr [ebp-1c]
'006c8596    6a34                    push 34
'006c8598    6830314300              push 00433130
'006c859d    51                      push ecx
'006c859e    50                      push eax
'006c859f    ffd3                    call ebx
'006c85a1    6683bd50ffffff00        cmp word ptr [ebp+ffffff50], 00
'006c85a9    0f856d010000            jne 6c871c

Do While (0 = 0)
'006c85af    8b17                    mov edx, dword ptr [edi]
'006c85b1    57                      push edi
'006c85b2    ff92fc020000            call dword ptr [edx+000002fc]
'006c85b8    50                      push eax
'006c85b9    8d45d4                  lea eax, dword ptr [ebp-2c]
'006c85bc    50                      push eax

' *** Reference to "__vbaObjSet"
'006c85bd    ff15b4104000            call dword ptr [004010b4]
    Set var_44 = var_40
'006c85c3    8d55dc                  lea edx, dword ptr [ebp-24]
'006c85c6    8bf0                    mov esi, eax
'006c85c8    8b45e4                  mov eax, dword ptr [ebp-1c]
'006c85cb    8b08                    mov ecx, dword ptr [eax]
'006c85cd    52                      push edx
'006c85ce    50                      push eax
'006c85cf    ff91b4000000            call dword ptr [ecx+000000b4]
'006c85d5    dbe2                    fnclex
'006c85d7    85c0                    test ax, ax
'006c85d9    7d11                    jge 6c85ec
'006c85db    8b4de4                  mov ecx, dword ptr [ebp-1c]
'006c85de    68b4000000              push 000000b4
'006c85e3    6830314300              push 00433130
'006c85e8    51                      push ecx
'006c85e9    50                      push eax
'006c85ea    ffd3                    call ebx
'006c85ec    8b45dc                  mov eax, dword ptr [ebp-24]
'006c85ef    8d5dd8                  lea ebx, dword ptr [ebp-28]
'006c85f2    53                      push ebx
'006c85f3    83ec10                  sub esp, 10
'006c85f6    b908000000              mov ecx, 00000008
'006c85fb    894db4                  mov dword ptr [ebp-4c], ecx
'006c85fe    8bdc                    mov ebx, esp
'006c8600    890b                    mov dword ptr [ebx], ecx
'006c8602    8b4db8                  mov ecx, dword ptr [ebp-48]
'006c8605    894b04                  mov dword ptr [ebx+04], ecx
'006c8608    c745bc64b14300          mov dword ptr [ebp-44], 0043b164
'006c860f    8b4dbc                  mov ecx, dword ptr [ebp-44]
'006c8612    8b10                    mov edx, dword ptr [eax]
'006c8614    894b08                  mov dword ptr [ebx+08], ecx
'006c8617    8b4dc0                  mov ecx, dword ptr [ebp-40]
'006c861a    50                      push eax
'006c861b    898548ffffff            mov dword ptr [ebp+ffffff48], eax
'006c8621    894b0c                  mov dword ptr [ebx+0c], ecx
'006c8624    ff5230                  call dword ptr [edx+30]
'006c8627    dbe2                    fnclex
'006c8629    85c0                    test ax, ax
'006c862b    7d15                    jge 6c8642
'006c862d    8b9548ffffff            mov edx, dword ptr [ebp+ffffff48]
'006c8633    6a30                    push 30
'006c8635    68d8304300              push 004330d8
'006c863a    52                      push edx
'006c863b    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c863c    ff1580104000            call dword ptr [00401080]
'006c8642    8b45d8                  mov eax, dword ptr [ebp-28]
'006c8645    8b08                    mov ecx, dword ptr [eax]
'006c8647    8d55c4                  lea edx, dword ptr [ebp-3c]
'006c864a    52                      push edx
'006c864b    50                      push eax
'006c864c    8bd8                    mov ebx, eax
'006c864e    ff5144                  call dword ptr [ecx+44]
'006c8651    dbe2                    fnclex
'006c8653    85c0                    test ax, ax
'006c8655    7d0f                    jge 6c8666
'006c8657    6a44                    push 44
'006c8659    6880304300              push 00433080
'006c865e    53                      push ebx
'006c865f    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c8660    ff1580104000            call dword ptr [00401080]
'006c8666    8b55a8                  mov edx, dword ptr [ebp-58]
'006c8669    8b1e                    mov ebx, dword ptr [esi]
'006c866b    83ec10                  sub esp, 10
'006c866e    8bcc                    mov ecx, esp
'006c8670    b80a000000              mov eax, 0000000a
'006c8675    8901                    mov dword ptr [ecx], eax
'006c8677    895104                  mov dword ptr [ecx+04], edx
'006c867a    b804000280              mov eax, 80020004
'006c867f    894108                  mov dword ptr [ecx+08], eax
'006c8682    8b45b0                  mov eax, dword ptr [ebp-50]
'006c8685    89410c                  mov dword ptr [ecx+0c], eax
'006c8688    8d4dc4                  lea ecx, dword ptr [ebp-3c]
'006c868b    51                      push ecx

' *** Reference to "__vbaStrVarMove"
'006c868c    ff153c104000            call dword ptr [0040103c]
'006c8692    8bd0                    mov edx, eax
'006c8694    8d4de0                  lea ecx, dword ptr [ebp-20]

' *** Reference to "__vbaStrMove"
'006c8697    ff15d0124000            call dword ptr [004012d0]
'006c869d    50                      push eax
'006c869e    56                      push esi
'006c869f    ff93ec010000            call dword ptr [ebx+000001ec]
'006c86a5    dbe2                    fnclex
'006c86a7    85c0                    test ax, ax
'006c86a9    7d16                    jge 6c86c1

' *** Reference to "__vbaHresultCheckObj"
'006c86ab    8b1d80104000            mov ebx, dword ptr [00401080]
'006c86b1    68ec010000              push 000001ec
'006c86b6    681cb94300              push 0043b91c
'006c86bb    56                      push esi
'006c86bc    50                      push eax
'006c86bd    ffd3                    call ebx
'006c86bf    eb06                    jmp 6c86c7

' *** Reference to "__vbaHresultCheckObj"
'006c86c1    8b1d80104000            mov ebx, dword ptr [00401080]
'006c86c7    8d4de0                  lea ecx, dword ptr [ebp-20]

' *** Reference to "__vbaFreeStr"
'006c86ca    ff150c134000            call dword ptr [0040130c]
    '#Cleanup(var_3)
'006c86d0    8d55d4                  lea edx, dword ptr [ebp-2c]
'006c86d3    52                      push edx
'006c86d4    8d45d8                  lea eax, dword ptr [ebp-28]
'006c86d7    50                      push eax
'006c86d8    8d4ddc                  lea ecx, dword ptr [ebp-24]
'006c86db    51                      push ecx
'006c86dc    6a03                    push 03

' *** Reference to "__vbaFreeObjList"
'006c86de    ff154c104000            call dword ptr [0040104c]
    '#Cleanup( var_39, var_45, var_44)
'006c86e4    83c410                  add esp, 10
'006c86e7    8d4dc4                  lea ecx, dword ptr [ebp-3c]

' *** Reference to "__vbaFreeVar"
'006c86ea    ff1528104000            call dword ptr [00401028]
    '#Cleanup(var_9)
'006c86f0    8b45e4                  mov eax, dword ptr [ebp-1c]
'006c86f3    8b10                    mov edx, dword ptr [eax]
'006c86f5    50                      push eax
'006c86f6    ff92ec000000            call dword ptr [edx+000000ec]
'006c86fc    dbe2                    fnclex
'006c86fe    85c0                    test ax, ax
'006c8700    0f8d77feffff            jge 6c857d
'006c8706    8b4de4                  mov ecx, dword ptr [ebp-1c]
'006c8709    68ec000000              push 000000ec
'006c870e    6830314300              push 00433130
'006c8713    51                      push ecx
'006c8714    50                      push eax
'006c8715    ffd3                    call ebx
'006c8717    e961feffff              jmp 6c857d
    
Loop
'006c871c    8b45e4                  mov eax, dword ptr [ebp-1c]
'006c871f    8b10                    mov edx, dword ptr [eax]
'006c8721    50                      push eax
'006c8722    ff92c4000000            call dword ptr [edx+000000c4]
'006c8728    dbe2                    fnclex
'006c872a    85c0                    test ax, ax
'006c872c    7d11                    jge 6c873f
'006c872e    8b4de4                  mov ecx, dword ptr [ebp-1c]
'006c8731    68c4000000              push 000000c4
'006c8736    6830314300              push 00433130
'006c873b    51                      push ecx
'006c873c    50                      push eax
'006c873d    ffd3                    call ebx
'006c873f    8b17                    mov edx, dword ptr [edi]
'006c8741    57                      push edi
'006c8742    ff92fc020000            call dword ptr [edx+000002fc]
'006c8748    50                      push eax
'006c8749    8d45dc                  lea eax, dword ptr [ebp-24]
'006c874c    50                      push eax

' *** Reference to "__vbaObjSet"
'006c874d    ff15b4104000            call dword ptr [004010b4]
Set var_39 = var_40
'006c8753    8bf0                    mov esi, eax
'006c8755    8b0e                    mov ecx, dword ptr [esi]
'006c8757    6a00                    push 00
'006c8759    56                      push esi
'006c875a    ff91f4000000            call dword ptr [ecx+000000f4]
'006c8760    dbe2                    fnclex
'006c8762    85c0                    test ax, ax
'006c8764    7d0e                    jge 6c8774
'006c8766    68f4000000              push 000000f4
'006c876b    681cb94300              push 0043b91c
'006c8770    56                      push esi
'006c8771    50                      push eax
'006c8772    ffd3                    call ebx
'006c8774    8d4ddc                  lea ecx, dword ptr [ebp-24]

' *** Reference to "__vbaFreeObj"
'006c8777    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_39)
'006c877d    33db                    xor ebx, ebx
var_num2 = Empty
'006c877f    8b17                    mov edx, dword ptr [edi]
'006c8781    57                      push edi
'006c8782    ff9230070000            call dword ptr [edx+00000730]
'006c8788    3bc3                    cmp eax, ebx
'006c878a    7d12                    jge 6c879e

If (var_39 < __vbaHresultCheckObj) Then
'006c878c    6830070000              push 00000730
'006c8791    6800144300              push 00431400
'006c8796    57                      push edi
'006c8797    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c8798    ff1580104000            call dword ptr [00401080]
    
End If
'006c879e    895dfc                  mov dword ptr [ebp-04], ebx
'006c87a1    68e8876c00              push 006c87e8
'006c87a6    eb2a                    jmp 6c87d2
'006c87a8    8d4de0                  lea ecx, dword ptr [ebp-20]

' *** Reference to "__vbaFreeStr"
'006c87ab    ff150c134000            call dword ptr [0040130c]
'#Cleanup(var_3)
'006c87b1    8d45d4                  lea eax, dword ptr [ebp-2c]
'006c87b4    50                      push eax
'006c87b5    8d4dd8                  lea ecx, dword ptr [ebp-28]
'006c87b8    51                      push ecx
'006c87b9    8d55dc                  lea edx, dword ptr [ebp-24]
'006c87bc    52                      push edx
'006c87bd    6a03                    push 03

' *** Reference to "__vbaFreeObjList"
'006c87bf    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_39, var_45, var_44)
'006c87c5    83c410                  add esp, 10
'006c87c8    8d4dc4                  lea ecx, dword ptr [ebp-3c]

' *** Reference to "__vbaFreeVar"
'006c87cb    ff1528104000            call dword ptr [00401028]
'#Cleanup(var_9)
'006c87d1    c3                      ret
'006c87d2    8d8d24ffffff            lea ecx, dword ptr [ebp+ffffff24]

' *** Reference to "__vbaFreeVar"
'006c87d8    ff1528104000            call dword ptr [00401028]
'#Cleanup(var_432)
'006c87de    8d4de4                  lea ecx, dword ptr [ebp-1c]

' *** Reference to "__vbaFreeObj"
'006c87e1    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_40)
'006c87e7    c3                      ret
'006c87e8    8b4508                  mov eax, dword ptr [ebp+08]
'006c87eb    8b08                    mov ecx, dword ptr [eax]
'006c87ed    50                      push eax
'006c87ee    ff5108                  call dword ptr [ecx+08]
'006c87f1    8b45fc                  mov eax, dword ptr [ebp-04]
'006c87f4    8b4dec                  mov ecx, dword ptr [ebp-14]
'006c87f7    5f                      pop edi
'006c87f8    5e                      pop esi
    'Reference to '__except_list'
'006c87f9    64890d00000000          mov dword ptr fs:[00000000], ecx
'006c8800    5b                      pop ebx
'006c8801    8be5                    mov esp, ebp
'006c8803    5d                      pop ebp
'006c8804    c20400                  ret 0004


End Sub


'Event for BtnFermer
Private Sub BtnFermer_Click()
'006c7390    55                      push ebp
'006c7391    8bec                    mov ebp, esp
'006c7393    83ec0c                  sub esp, 0c

' *** Reference to "__vbaExceptHandler"
'006c7396    6866724000              push 00407266
'006c739b    64a100000000            mov ax, word ptr fs:[00000000]
'006c73a1    50                      push eax
    'Reference to '__except_list'
'006c73a2    64892500000000          mov dword ptr fs:[00000000], esp
'006c73a9    83ec18                  sub esp, 18
'006c73ac    53                      push ebx
'006c73ad    56                      push esi
'006c73ae    57                      push edi
'006c73af    8965f4                  mov dword ptr [ebp-0c], esp
'006c73b2    c745f850674000          mov dword ptr [ebp-08], 00406750
'006c73b9    8b7d08                  mov edi, dword ptr [ebp+08]
'006c73bc    8bc7                    mov eax, edi
'006c73be    83e001                  and eax, 01
'006c73c1    8945fc                  mov dword ptr [ebp-04], eax
'006c73c4    83e7fe                  and edi, fffffffe
'006c73c7    8b0f                    mov ecx, dword ptr [edi]
'006c73c9    57                      push edi
'006c73ca    897d08                  mov dword ptr [ebp+08], edi
'006c73cd    ff5104                  call dword ptr [ecx+04]
'006c73d0    a124be7200              mov ax, word ptr [0072be24]
'006c73d5    33db                    xor ebx, ebx
'006c73d7    3bc3                    cmp eax, ebx
'006c73d9    895de8                  mov dword ptr [ebp-18], ebx
'006c73dc    7510                    jne 6c73ee

If (0 = 0) Then
'006c73de    6824be7200              push 0072be24
'006c73e3    6870004300              push 00430070

' *** Reference to "__vbaNew2"
'006c73e8    ff152c124000            call dword ptr [0040122c]
    Dim var_2 As New Global
End If
'006c73ee    8b3524be7200            mov esi, dword ptr [0072be24]
'006c73f4    8b16                    mov edx, dword ptr [esi]
'006c73f6    57                      push edi
'006c73f7    8d45e8                  lea eax, dword ptr [ebp-18]
'006c73fa    50                      push eax
'006c73fb    8955d4                  mov dword ptr [ebp-2c], edx

' *** Reference to "__vbaObjSetAddref"
'006c73fe    ff15c8104000            call dword ptr [004010c8]
Set var_41 = Me
'006c7404    8b4dd4                  mov ecx, dword ptr [ebp-2c]
'006c7407    50                      push eax
'006c7408    56                      push esi
'006c7409    ff5110                  call dword ptr [ecx+10]
Call var_2.Unload(var_41)
'006c740c    dbe2                    fnclex
'006c740e    3bc3                    cmp eax, ebx
'006c7410    7d0f                    jge 6c7421
'006c7412    6a10                    push 10
'006c7414    6860004300              push 00430060
'006c7419    56                      push esi
'006c741a    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c741b    ff1580104000            call dword ptr [00401080]
'006c7421    8d4de8                  lea ecx, dword ptr [ebp-18]

' *** Reference to "__vbaFreeObj"
'006c7424    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_41)
'006c742a    895dfc                  mov dword ptr [ebp-04], ebx
'006c742d    683f746c00              push 006c743f
'006c7432    eb0a                    jmp 6c743e
'006c7434    8d4de8                  lea ecx, dword ptr [ebp-18]

' *** Reference to "__vbaFreeObj"
'006c7437    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_41)
'006c743d    c3                      ret
'006c743e    c3                      ret
'006c743f    8b4508                  mov eax, dword ptr [ebp+08]
'006c7442    8b10                    mov edx, dword ptr [eax]
'006c7444    50                      push eax
'006c7445    ff5208                  call dword ptr [edx+08]
'006c7448    8b45fc                  mov eax, dword ptr [ebp-04]
'006c744b    8b4dec                  mov ecx, dword ptr [ebp-14]
'006c744e    5f                      pop edi
'006c744f    5e                      pop esi
    'Reference to '__except_list'
'006c7450    64890d00000000          mov dword ptr fs:[00000000], ecx
'006c7457    5b                      pop ebx
'006c7458    8be5                    mov esp, ebp
'006c745a    5d                      pop ebp
'006c745b    c20400                  ret 0004


End Sub


