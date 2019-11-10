VERSION 5.00

Begin VB.Form frmInsertionDons
    Caption = "Insertion de dons"
    ScaleMode = 1
    AutoRedraw = 0              'False
    FontTransparent = -1              'True
    BorderStyle = 1
    LinkTopic = "Form1"
    MaxButton = 0              'False
    MinButton = 0              'False
    MDIChild = -1              'True
    ClientLeft   = 45
    ClientTop    = 435
    ClientWidth  = 9180
    ClientHeight = 7710
    BeginProperty Font
        Name          = "Times New Roman"
        Size          = 9,75
        Charset       = 0
        Weight        = 700
        Underline     = 0              'False
        Italic        = 0              'False
        Strikethrough = 0              'False
    EndProperty
    Begin VB.ComboBox cmbCaracteristique
        Left   = 5355
        Top    = 5760
        Width  = 2055
        Height = 345
        Text = "CaractÈristique&êê_"
        TabIndex = 27
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
    Begin VB.ComboBox cmbGenre
        Left   = 4590
        Top    = 90
        Width  = 2775
        Height = 345
        Text = "Genre&"
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
    Begin VB.ComboBox cmbSource
        Left   = 1260
        Top    = 465
        Width  = 2415
        Height = 345
        Text = "Livre&"
        TabIndex = 0
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
    Begin VB.ComboBox cmbNom
        Left   = 1260
        Top    = 90
        Width  = 2415
        Height = 345
        Text = "Nom du don&ê"
        TabIndex = 1
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
    Begin VB.CheckBox ChkPrive
        Caption = "PrivÈ"
        Left   = 135
        Top    = 6285
        Width  = 2295
        Height = 375
        TabIndex = 25
        Alignment = 1
        ToolTipText = "Un don accessible seulement ‡ une classe ou une race ne peut Ítre choisi en tant que don (ex : ch‚timent du mal)."
    End
    Begin VB.CheckBox ChkMultiple
        Caption = "Multiple"
        Left   = 3825
        Top    = 6285
        Width  = 2295
        Height = 375
        TabIndex = 24
        Alignment = 1
        ToolTipText = "Un don que l'on peut prendre plusieurs fois (ex : robustesse)"
    End
    Begin VB.CheckBox ChkEpique
        Caption = "Epique"
        Left   = 6720
        Top    = 6285
        Width  = 2295
        Height = 375
        TabIndex = 23
        Alignment = 1
        ToolTipText = "Un don que l'on peut seulement choisir ‡ partir du niveau 21."
    End
    Begin VB.TextBox fldnPage
        Left   = 4680
        Top    = 495
        Width  = 615
        Height = 285
        TabIndex = 22
    End
    Begin VB.TextBox fldnVersion
        Left   = 8400
        Top    = 120
        Width  = 615
        Height = 285
        Text = "3.5"
        TabIndex = 21
    End
    Begin VB.CommandButton btnNouveau
        Caption = "&Nouveau"
        Left   = 5280
        Top    = 7320
        Width  = 1215
        Height = 375
        TabIndex = 20
    End
    Begin VB.CommandButton btnEnregistrer
        Caption = "&Enregistrer"
        Left   = 6600
        Top    = 7320
        Width  = 1215
        Height = 375
        TabIndex = 19
    End
    Begin VB.CommandButton btnFermer
        Caption = "&Fermer"
        Left   = 7920
        Top    = 7320
        Width  = 1215
        Height = 375
        TabIndex = 18
    End
    Begin VB.CommandButton btnSuivant
        Caption = "&Suivant"
        Left   = 3960
        Top    = 7320
        Width  = 1215
        Height = 375
        TabIndex = 17
    End
    Begin VB.CommandButton btnPrecedent
        Caption = "&PrÈcÈdent"
        Left   = 2640
        Top    = 7320
        Width  = 1215
        Height = 375
        TabIndex = 16
    End
    Begin VB.TextBox fldstrResume
        Left   = 120
        Top    = 1200
        Width  = 8895
        Height = 975
        TabIndex = 15
        MultiLine = -1              'True
        ScrollBars = 2
        BeginProperty Font
            Name          = "Times New Roman"
            Size          = 9,75
            Charset       = 0
            Weight        = 400
            Underline     = 0              'False
            Italic        = 0              'False
            Strikethrough = 0              'False
        EndProperty
    End
    Begin VB.TextBox fldstrExplication
        Left   = 120
        Top    = 2640
        Width  = 8895
        Height = 3015
        TabIndex = 14
        MultiLine = -1              'True
        ScrollBars = 2
        BeginProperty Font
            Name          = "Times New Roman"
            Size          = 9,75
            Charset       = 0
            Weight        = 400
            Underline     = 0              'False
            Italic        = 0              'False
            Strikethrough = 0              'False
        EndProperty
    End
    Begin VB.CheckBox ChkTare
        Caption = "Handicaps"
        Left   = 6720
        Top    = 6840
        Width  = 2295
        Height = 375
        TabIndex = 13
        Alignment = 1
        ToolTipText = "Les andicaps sont le revers des dons (ex : Affligeant)"
    End
    Begin VB.CheckBox ChkTrait
        Caption = "Trait"
        Left   = 3800
        Top    = 6840
        Width  = 2295
        Height = 375
        TabIndex = 12
        Alignment = 1
        ToolTipText = "Un trait reprÈsente un aspect de la personnalitÈ du PJ (ex : Abrupt)"
    End
    Begin VB.CheckBox ChkLiearme
        Caption = "LiÈ a une arme"
        Left   = 120
        Top    = 6840
        Width  = 2295
        Height = 375
        TabIndex = 11
        Alignment = 1
        ToolTipText = "Un don qui donne un Èffet sur un type d'arme (ex : arme de prÈdilection)"
    End
    Begin VB.CheckBox ChkPouvoir
        Caption = "Pouvoir"
        Left   = 120
        Top    = 5805
        Width  = 2295
        Height = 375
        TabIndex = 10
        Alignment = 1
        ToolTipText = "Un pouvoir qui se calcule sous la forme 10 + 1/2 dÈs de vie + bonus caratÈristique"
    End
    Begin VB.Label LabelDon
        Caption = "CaractÈristique"
        Index = 7
        Left   = 3780
        Top    = 5805
        Width  = 1455
        Height = 255
        TabIndex = 9
        ToolTipText = "La caractÈristique associÈe au pouvoir"
    End
    Begin VB.Label LabelDon
        Caption = "Explication"
        Index = 6
        Left   = 3870
        Top    = 2280
        Width  = 1455
        Height = 255
        TabIndex = 8
    End
    Begin VB.Label LabelDon
        Caption = "RÈsumÈ"
        Index = 5
        Left   = 3840
        Top    = 840
        Width  = 1455
        Height = 255
        TabIndex = 7
    End
    Begin VB.Label LabelDon
        Caption = "Page"
        Index = 4
        Left   = 3795
        Top    = 510
        Width  = 1455
        Height = 255
        TabIndex = 6
    End
    Begin VB.Label LabelDon
        Caption = "Source"
        Index = 3
        Left   = 120
        Top    = 510
        Width  = 1455
        Height = 255
        TabIndex = 5
    End
    Begin VB.Label LabelDon
        Caption = "Genre"
        Index = 2
        Left   = 3795
        Top    = 135
        Width  = 1455
        Height = 255
        TabIndex = 4
    End
    Begin VB.Label LabelDon
        Caption = "Version"
        Index = 1
        Left   = 7560
        Top    = 135
        Width  = 1455
        Height = 255
        TabIndex = 3
        ToolTipText = "La version du livre dans lequel est tirÈ le don souvent 3.5 ou 3.0."
    End
    Begin VB.Label LabelDon
        Caption = "Nom du don"
        Index = 0
        Left   = 120
        Top    = 135
        Width  = 1455
        Height = 255
        TabIndex = 2
    End
End
Public Function AfficheDon(arg_0 As Unknow, arg_1 As Unknow, arg_2 As Unknow, arg_3 As Unknow, arg_4 As Unknow, arg_5 As Unknow, arg_6 As Unknow, arg_7 As Unknow, arg_8 As Unknow, arg_9 As Unknow, arg_A As Unknow, arg_B As Unknow, arg_C As Unknow, arg_D As Unknow, arg_E As Unknow, arg_F As Unknow, arg_10 As Unknow, arg_11 As Unknow, arg_12 As Unknow, arg_13 As Unknow, arg_14 As Unknow, arg_15 As Unknow, arg_16 As Unknow, arg_17 As Unknow, arg_18 As Unknow, arg_19 As Unknow, arg_1A As Unknow, arg_1B As Unknow, arg_1C As Unknow, arg_1D As Unknow, arg_1E As Unknow, arg_1F As Unknow, arg_20 As Unknow, arg_21 As Unknow, arg_22 As Unknow, arg_23 As Unknow, arg_24 As Unknow, arg_25 As Unknow, arg_26 As Unknow, arg_27 As Unknow, arg_28 As Unknow, arg_29 As Unknow, arg_2A As Unknow, arg_2B As Unknow, arg_2C As Unknow, arg_2D As Unknow, arg_2E As Unknow, arg_2F As Unknow, arg_30 As Unknow, arg_31 As Unknow, arg_32 As Unknow, arg_33 As Unknow, arg_34 As Unknow, arg_35 As Unknow, arg_36 As Unknow, arg_37 As Unknow, arg_38 As Unknow, arg_39 As Unknow, arg_3A As Unknow, arg_3B As Unknow, arg_3C As Unknow)
'00712e30    55                      push ebp
'00712e31    8bec                    mov ebp, esp
'00712e33    83ec0c                  sub esp, 0c

' *** Reference to "__vbaExceptHandler"
'00712e36    6866724000              push 00407266
'00712e3b    64a100000000            mov ax, word ptr fs:[00000000]
'00712e41    50                      push eax
    'Reference to '__except_list'
'00712e42    64892500000000          mov dword ptr fs:[00000000], esp
'00712e49    81ec2c010000            sub esp, 0000012c
'00712e4f    53                      push ebx
'00712e50    56                      push esi
'00712e51    57                      push edi
'00712e52    8965f4                  mov dword ptr [ebp-0c], esp
'00712e55    c745f8906d4000          mov dword ptr [ebp-08], 00406d90
'00712e5c    33f6                    xor esi, esi
'00712e5e    8975fc                  mov dword ptr [ebp-04], esi
'00712e61    8b7d08                  mov edi, dword ptr [ebp+08]
'00712e64    8b07                    mov eax, dword ptr [edi]
'00712e66    57                      push edi
'00712e67    ff5004                  call dword ptr [eax+04]
'00712e6a    8b0f                    mov ecx, dword ptr [edi]
'00712e6c    57                      push edi
'00712e6d    8975e8                  mov dword ptr [ebp-18], esi
'00712e70    8975e4                  mov dword ptr [ebp-1c], esi
'00712e73    8975e0                  mov dword ptr [ebp-20], esi
'00712e76    8975dc                  mov dword ptr [ebp-24], esi
'00712e79    8975d8                  mov dword ptr [ebp-28], esi
'00712e7c    8975d4                  mov dword ptr [ebp-2c], esi
'00712e7f    8975d0                  mov dword ptr [ebp-30], esi
'00712e82    8975cc                  mov dword ptr [ebp-34], esi
'00712e85    8975c8                  mov dword ptr [ebp-38], esi
'00712e88    8975c4                  mov dword ptr [ebp-3c], esi
'00712e8b    8975c0                  mov dword ptr [ebp-40], esi
'00712e8e    8975bc                  mov dword ptr [ebp-44], esi
'00712e91    8975ac                  mov dword ptr [ebp-54], esi
'00712e94    89759c                  mov dword ptr [ebp-64], esi
'00712e97    89758c                  mov dword ptr [ebp-74], esi
'00712e9a    89b57cffffff            mov dword ptr [ebp+ffffff7c], esi
'00712ea0    89b54cffffff            mov dword ptr [ebp+ffffff4c], esi
'00712ea6    89b53cffffff            mov dword ptr [ebp+ffffff3c], esi
'00712eac    89b538ffffff            mov dword ptr [ebp+ffffff38], esi
'00712eb2    89b514ffffff            mov dword ptr [ebp+ffffff14], esi
'00712eb8    ff9108030000            call dword ptr [ecx+00000308]
'00712ebe    50                      push eax
'00712ebf    8d55cc                  lea edx, dword ptr [ebp-34]
'00712ec2    52                      push edx

' *** Reference to "__vbaObjSet"
'00712ec3    ff15b4104000            call dword ptr [004010b4]
Set var_43 = Nothing
'00712ec9    8b4d0c                  mov ecx, dword ptr [ebp+0c]
'00712ecc    33d2                    xor edx, edx
'00712ece    668b11                  mov dx, word ptr [ecx]
'00712ed1    8bd8                    mov ebx, eax
'00712ed3    8b03                    mov eax, dword ptr [ebx]
'00712ed5    52                      push edx
'00712ed6    53                      push ebx
'00712ed7    ff90f4000000            call dword ptr [eax+000000f4]
'00712edd    dbe2                    fnclex
'00712edf    3bc6                    cmp eax, esi
'00712ee1    7d12                    jge 712ef5

If (-256 - 12 < 0) Then
'00712ee3    68f4000000              push 000000f4
'00712ee8    681cb94300              push 0043b91c
'00712eed    53                      push ebx
'00712eee    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00712eef    ff1580104000            call dword ptr [00401080]
    
End If
'00712ef5    8d4dcc                  lea ecx, dword ptr [ebp-34]

' *** Reference to "__vbaFreeObj"
'00712ef8    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_43)
'00712efe    8b07                    mov eax, dword ptr [edi]
'00712f00    57                      push edi
'00712f01    ff9008030000            call dword ptr [eax+00000308]
'00712f07    50                      push eax
'00712f08    8d4dcc                  lea ecx, dword ptr [ebp-34]
'00712f0b    51                      push ecx

' *** Reference to "__vbaObjSet"
'00712f0c    ff15b4104000            call dword ptr [004010b4]
Set var_43 = Nothing
'00712f12    8bf8                    mov edi, eax
'00712f14    8b17                    mov edx, dword ptr [edi]
'00712f16    8d45e4                  lea eax, dword ptr [ebp-1c]
'00712f19    50                      push eax
'00712f1a    57                      push edi
'00712f1b    ff92a8000000            call dword ptr [edx+000000a8]
'00712f21    dbe2                    fnclex
'00712f23    3bc6                    cmp eax, esi
'00712f25    7d12                    jge 712f39
'00712f27    68a8000000              push 000000a8
'00712f2c    681cb94300              push 0043b91c
'00712f31    57                      push edi
'00712f32    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00712f33    ff1580104000            call dword ptr [00401080]
'00712f39    8b55e4                  mov edx, dword ptr [ebp-1c]

' *** Reference to "__vbaStrMove"
'00712f3c    8b3dd0124000            mov edi, dword ptr [004012d0]
'00712f42    8d4de0                  lea ecx, dword ptr [ebp-20]
'00712f45    8975e4                  mov dword ptr [ebp-1c], esi
'00712f48    ffd7                    call edi
'00712f4a    8d4de0                  lea ecx, dword ptr [ebp-20]
'00712f4d    51                      push ecx
'00712f4e    e89d0cdeff              call 4f3bf0
Call sub_4F3BF0()
'00712f53    8bd0                    mov edx, eax
'00712f55    8d4dd0                  lea ecx, dword ptr [ebp-30]
'00712f58    ffd7                    call edi
'00712f5a    8b55d0                  mov edx, dword ptr [ebp-30]
'00712f5d    899504ffffff            mov dword ptr [ebp+ffffff04], edx
'00712f63    8b1548b07200            mov edx, dword ptr [0072b048]
'00712f69    b804000280              mov eax, 80020004
'00712f6e    898554ffffff            mov dword ptr [ebp+ffffff54], eax
'00712f74    b90a000000              mov ecx, 0000000a
'00712f79    898d4cffffff            mov dword ptr [ebp+ffffff4c], ecx
'00712f7f    898d5cffffff            mov dword ptr [ebp+ffffff5c], ecx
'00712f85    8975d0                  mov dword ptr [ebp-30], esi
'00712f88    8b1a                    mov ebx, dword ptr [edx]
'00712f8a    8d55c8                  lea edx, dword ptr [ebp-38]
'00712f8d    52                      push edx
'00712f8e    83ec10                  sub esp, 10
'00712f91    8bd4                    mov edx, esp
'00712f93    890a                    mov dword ptr [edx], ecx
'00712f95    8b8d50ffffff            mov ecx, dword ptr [ebp+ffffff50]
'00712f9b    894a04                  mov dword ptr [edx+04], ecx
'00712f9e    894208                  mov dword ptr [edx+08], eax
'00712fa1    898564ffffff            mov dword ptr [ebp+ffffff64], eax
'00712fa7    8b8558ffffff            mov eax, dword ptr [ebp+ffffff58]
'00712fad    89420c                  mov dword ptr [edx+0c], eax
'00712fb0    8b955cffffff            mov edx, dword ptr [ebp+ffffff5c]
'00712fb6    8b8560ffffff            mov eax, dword ptr [ebp+ffffff60]
'00712fbc    83ec10                  sub esp, 10
'00712fbf    8bcc                    mov ecx, esp
'00712fc1    8911                    mov dword ptr [ecx], edx
'00712fc3    8b9564ffffff            mov edx, dword ptr [ebp+ffffff64]
'00712fc9    894104                  mov dword ptr [ecx+04], eax
'00712fcc    8b8568ffffff            mov eax, dword ptr [ebp+ffffff68]
'00712fd2    895108                  mov dword ptr [ecx+08], edx
'00712fd5    8b9570ffffff            mov edx, dword ptr [ebp+ffffff70]
'00712fdb    89410c                  mov dword ptr [ecx+0c], eax
'00712fde    83ec10                  sub esp, 10
'00712fe1    8bcc                    mov ecx, esp
'00712fe3    b803000000              mov eax, 00000003
'00712fe8    8901                    mov dword ptr [ecx], eax
'00712fea    895104                  mov dword ptr [ecx+04], edx
'00712fed    8b9504ffffff            mov edx, dword ptr [ebp+ffffff04]
'00712ff3    b802000000              mov eax, 00000002
'00712ff8    894108                  mov dword ptr [ecx+08], eax
'00712ffb    8b8578ffffff            mov eax, dword ptr [ebp+ffffff78]
'00713001    89410c                  mov dword ptr [ecx+0c], eax
'00713004    68186b4500              push 00456b18
'00713009    8d4ddc                  lea ecx, dword ptr [ebp-24]
'0071300c    ffd7                    call edi
'0071300e    50                      push eax

' *** Reference to "__vbaStrCat"
'0071300f    ff1570104000            call dword ptr [00401070]
var_49 = ("select * from donsperso where nom='") & (vbNullString)
'00713015    8bd0                    mov edx, eax
'00713017    8d4dd8                  lea ecx, dword ptr [ebp-28]
'0071301a    ffd7                    call edi
'0071301c    50                      push eax
'0071301d    6854a44300              push 0043a454

' *** Reference to "__vbaStrCat"
'00713022    ff1570104000            call dword ptr [00401070]
var_84 = (var_49) & ("'")
'00713028    8bd0                    mov edx, eax
'0071302a    8d4dd4                  lea ecx, dword ptr [ebp-2c]
'0071302d    ffd7                    call edi
'0071302f    8b0d48b07200            mov ecx, dword ptr [0072b048]
'00713035    50                      push eax
'00713036    51                      push ecx
'00713037    ff93bc000000            call dword ptr [ebx+000000bc]
'0071303d    dbe2                    fnclex
'0071303f    3bc6                    cmp eax, esi
'00713041    7d18                    jge 71305b

If (-4500 < 0) Then
'00713043    8b1548b07200            mov edx, dword ptr [0072b048]
'00713049    68bc000000              push 000000bc
'0071304e    68301f4300              push 00431f30
'00713053    52                      push edx
'00713054    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00713055    ff1580104000            call dword ptr [00401080]
    
End If
'0071305b    8b45c8                  mov eax, dword ptr [ebp-38]

' *** Reference to "__vbaObjSet"
'0071305e    8b1db4104000            mov ebx, dword ptr [004010b4]
'00713064    50                      push eax
'00713065    8d45e8                  lea eax, dword ptr [ebp-18]
'00713068    50                      push eax
'00713069    8975c8                  mov dword ptr [ebp-38], esi
'0071306c    ffd3                    call ebx
Set var_41 = Nothing
'0071306e    8d4dd0                  lea ecx, dword ptr [ebp-30]
'00713071    51                      push ecx
'00713072    8d55d4                  lea edx, dword ptr [ebp-2c]
'00713075    52                      push edx
'00713076    8d45d8                  lea eax, dword ptr [ebp-28]
'00713079    50                      push eax
'0071307a    8d4ddc                  lea ecx, dword ptr [ebp-24]
'0071307d    51                      push ecx
'0071307e    8d55e0                  lea edx, dword ptr [ebp-20]
'00713081    52                      push edx
'00713082    6a05                    push 05

' *** Reference to "__vbaFreeStrList"
'00713084    ff155c124000            call dword ptr [0040125c]
'#Cleanup( 0, 0, -4496, -4500, 0)
'0071308a    83c418                  add esp, 18
'0071308d    8d4dcc                  lea ecx, dword ptr [ebp-34]

' *** Reference to "__vbaFreeObj"
'00713090    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_43)
'00713096    8b45e8                  mov eax, dword ptr [ebp-18]
'00713099    50                      push eax
'0071309a    8d8d14ffffff            lea ecx, dword ptr [ebp+ffffff14]
'007130a0    51                      push ecx

' *** Reference to "__vbaObjSetAddref"
'007130a1    ff15c8104000            call dword ptr [004010c8]
Set var_488 = Nothing
'007130a7    8b8514ffffff            mov eax, dword ptr [ebp+ffffff14]
'007130ad    8b10                    mov edx, dword ptr [eax]
'007130af    8d8d38ffffff            lea ecx, dword ptr [ebp+ffffff38]
'007130b5    51                      push ecx
'007130b6    50                      push eax
'007130b7    ff5234                  call dword ptr [edx+34]
'007130ba    dbe2                    fnclex
'007130bc    3bc6                    cmp eax, esi
'007130be    7d15                    jge 7130d5

If (var_488 < 0) Then
'007130c0    8b9514ffffff            mov edx, dword ptr [ebp+ffffff14]
'007130c6    6a34                    push 34
'007130c8    6830314300              push 00433130
'007130cd    52                      push edx
'007130ce    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'007130cf    ff1580104000            call dword ptr [00401080]
    
End If
'007130d5    6639b538ffffff          cmp word ptr [ebp+ffffff38], si
'007130dc    0f8591130000            jne 714473

If (0 = 0) Then
'007130e2    8b4508                  mov eax, dword ptr [ebp+08]
'007130e5    8b08                    mov ecx, dword ptr [eax]
'007130e7    50                      push eax
'007130e8    ff911c030000            call dword ptr [ecx+0000031c]
'007130ee    50                      push eax
'007130ef    8d55c4                  lea edx, dword ptr [ebp-3c]
'007130f2    52                      push edx
'007130f3    ffd3                    call ebx
    Set var_9 = Me
'007130f5    8d55cc                  lea edx, dword ptr [ebp-34]
'007130f8    898520ffffff            mov dword ptr [ebp+ffffff20], eax
'007130fe    8b8514ffffff            mov eax, dword ptr [ebp+ffffff14]
'00713104    8b08                    mov ecx, dword ptr [eax]
'00713106    52                      push edx
'00713107    50                      push eax
'00713108    ff91b4000000            call dword ptr [ecx+000000b4]
'0071310e    dbe2                    fnclex
'00713110    3bc6                    cmp eax, esi
'00713112    7d18                    jge 71312c
    
    If (    var_488 < 0) Then
'00713114    8b8d14ffffff            mov ecx, dword ptr [ebp+ffffff14]
'0071311a    68b4000000              push 000000b4
'0071311f    6830314300              push 00433130
'00713124    51                      push ecx
'00713125    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00713126    ff1580104000            call dword ptr [00401080]
    
End If
'0071312c    8b45cc                  mov eax, dword ptr [ebp-34]
'0071312f    8b10                    mov edx, dword ptr [eax]
'00713131    8d5dc8                  lea ebx, dword ptr [ebp-38]
'00713134    53                      push ebx
'00713135    83ec10                  sub esp, 10
'00713138    8bdc                    mov ebx, esp
'0071313a    b908000000              mov ecx, 00000008
'0071313f    890b                    mov dword ptr [ebx], ecx
'00713141    8b8d70ffffff            mov ecx, dword ptr [ebp+ffffff70]
'00713147    894b04                  mov dword ptr [ebx+04], ecx
'0071314a    b90ca94300              mov ecx, 0043a90c
'0071314f    894b08                  mov dword ptr [ebx+08], ecx
'00713152    8b8d78ffffff            mov ecx, dword ptr [ebp+ffffff78]
'00713158    50                      push eax
'00713159    898530ffffff            mov dword ptr [ebp+ffffff30], eax
'0071315f    894b0c                  mov dword ptr [ebx+0c], ecx
'00713162    ff5230                  call dword ptr [edx+30]
'00713165    dbe2                    fnclex
'00713167    3bc6                    cmp eax, esi
'00713169    7d15                    jge 713180

If (var_43 < 0) Then
'0071316b    8b9530ffffff            mov edx, dword ptr [ebp+ffffff30]
'00713171    6a30                    push 30
'00713173    68d8304300              push 004330d8
'00713178    52                      push edx
'00713179    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0071317a    ff1580104000            call dword ptr [00401080]
    
End If
'00713180    8b45c8                  mov eax, dword ptr [ebp-38]
'00713183    8b08                    mov ecx, dword ptr [eax]
'00713185    8d55ac                  lea edx, dword ptr [ebp-54]
'00713188    52                      push edx
'00713189    50                      push eax
'0071318a    8bd8                    mov ebx, eax
'0071318c    ff5144                  call dword ptr [ecx+44]
'0071318f    dbe2                    fnclex
'00713191    3bc6                    cmp eax, esi
'00713193    7d0f                    jge 7131a4

If (0 < 0) Then
'00713195    6a44                    push 44
'00713197    6880304300              push 00433080
'0071319c    53                      push ebx
'0071319d    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0071319e    ff1580104000            call dword ptr [00401080]
    
End If
'007131a4    8b8520ffffff            mov eax, dword ptr [ebp+ffffff20]
'007131aa    8b18                    mov ebx, dword ptr [eax]
'007131ac    8d4dac                  lea ecx, dword ptr [ebp-54]
'007131af    51                      push ecx

' *** Reference to "__vbaStrVarMove"
'007131b0    ff153c104000            call dword ptr [0040103c]
'007131b6    8bd0                    mov edx, eax
'007131b8    8d4de4                  lea ecx, dword ptr [ebp-1c]
'007131bb    ffd7                    call edi
'007131bd    8bd3                    mov edx, ebx
'007131bf    8b9d20ffffff            mov ebx, dword ptr [ebp+ffffff20]
'007131c5    50                      push eax
'007131c6    53                      push ebx
'007131c7    ff92a4000000            call dword ptr [edx+000000a4]
'007131cd    dbe2                    fnclex
'007131cf    3bc6                    cmp eax, esi
'007131d1    7d12                    jge 7131e5

If (-4504 < 0) Then
'007131d3    68a4000000              push 000000a4
'007131d8    68d00d4300              push 00430dd0
'007131dd    53                      push ebx
'007131de    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'007131df    ff1580104000            call dword ptr [00401080]
    
End If
'007131e5    8d4de4                  lea ecx, dword ptr [ebp-1c]

' *** Reference to "__vbaFreeStr"
'007131e8    ff150c134000            call dword ptr [0040130c]
'#Cleanup(var_40)
'007131ee    8d45c4                  lea eax, dword ptr [ebp-3c]
'007131f1    50                      push eax
'007131f2    8d4dc8                  lea ecx, dword ptr [ebp-38]
'007131f5    51                      push ecx
'007131f6    8d55cc                  lea edx, dword ptr [ebp-34]
'007131f9    52                      push edx
'007131fa    6a03                    push 03

' *** Reference to "__vbaFreeObjList"
'007131fc    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_43, var_46, var_9)
'00713202    83c410                  add esp, 10
'00713205    8d4dac                  lea ecx, dword ptr [ebp-54]

' *** Reference to "__vbaFreeVar"
'00713208    ff1528104000            call dword ptr [00401028]
'#Cleanup(var_50)
'0071320e    8b4508                  mov eax, dword ptr [ebp+08]
'00713211    8b08                    mov ecx, dword ptr [eax]
'00713213    50                      push eax
'00713214    ff9100030000            call dword ptr [ecx+00000300]
'0071321a    50                      push eax
'0071321b    8d55c4                  lea edx, dword ptr [ebp-3c]
'0071321e    52                      push edx

' *** Reference to "__vbaObjSet"
'0071321f    ff15b4104000            call dword ptr [004010b4]
Set var_9 = Me
'00713225    8d55cc                  lea edx, dword ptr [ebp-34]
'00713228    898520ffffff            mov dword ptr [ebp+ffffff20], eax
'0071322e    8b8514ffffff            mov eax, dword ptr [ebp+ffffff14]
'00713234    8b08                    mov ecx, dword ptr [eax]
'00713236    52                      push edx
'00713237    50                      push eax
'00713238    ff91b4000000            call dword ptr [ecx+000000b4]
'0071323e    dbe2                    fnclex
'00713240    3bc6                    cmp eax, esi
'00713242    7d18                    jge 71325c

If (var_488 < 0) Then
'00713244    8b8d14ffffff            mov ecx, dword ptr [ebp+ffffff14]
'0071324a    68b4000000              push 000000b4
'0071324f    6830314300              push 00433130
'00713254    51                      push ecx
'00713255    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00713256    ff1580104000            call dword ptr [00401080]
    
End If
'0071325c    8b45cc                  mov eax, dword ptr [ebp-34]
'0071325f    8b10                    mov edx, dword ptr [eax]
'00713261    8d5dc8                  lea ebx, dword ptr [ebp-38]
'00713264    53                      push ebx
'00713265    83ec10                  sub esp, 10
'00713268    8bdc                    mov ebx, esp
'0071326a    b908000000              mov ecx, 00000008
'0071326f    890b                    mov dword ptr [ebx], ecx
'00713271    8b8d70ffffff            mov ecx, dword ptr [ebp+ffffff70]
'00713277    894b04                  mov dword ptr [ebx+04], ecx
'0071327a    b9acb24300              mov ecx, 0043b2ac
'0071327f    894b08                  mov dword ptr [ebx+08], ecx
'00713282    8b8d78ffffff            mov ecx, dword ptr [ebp+ffffff78]
'00713288    50                      push eax
'00713289    898530ffffff            mov dword ptr [ebp+ffffff30], eax
'0071328f    894b0c                  mov dword ptr [ebx+0c], ecx
'00713292    ff5230                  call dword ptr [edx+30]
'00713295    dbe2                    fnclex
'00713297    3bc6                    cmp eax, esi
'00713299    7d15                    jge 7132b0

If (var_43 < 0) Then
'0071329b    8b9530ffffff            mov edx, dword ptr [ebp+ffffff30]
'007132a1    6a30                    push 30
'007132a3    68d8304300              push 004330d8
'007132a8    52                      push edx
'007132a9    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'007132aa    ff1580104000            call dword ptr [00401080]
    
End If
'007132b0    8b45c8                  mov eax, dword ptr [ebp-38]
'007132b3    8b08                    mov ecx, dword ptr [eax]
'007132b5    8d55ac                  lea edx, dword ptr [ebp-54]
'007132b8    52                      push edx
'007132b9    50                      push eax
'007132ba    8bd8                    mov ebx, eax
'007132bc    ff5144                  call dword ptr [ecx+44]
'007132bf    dbe2                    fnclex
'007132c1    3bc6                    cmp eax, esi
'007132c3    7d0f                    jge 7132d4

If (0 < 0) Then
'007132c5    6a44                    push 44
'007132c7    6880304300              push 00433080
'007132cc    53                      push ebx
'007132cd    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'007132ce    ff1580104000            call dword ptr [00401080]
    
End If
'007132d4    8b8520ffffff            mov eax, dword ptr [ebp+ffffff20]
'007132da    8b18                    mov ebx, dword ptr [eax]
'007132dc    8d4dac                  lea ecx, dword ptr [ebp-54]
'007132df    51                      push ecx

' *** Reference to "__vbaStrVarMove"
'007132e0    ff153c104000            call dword ptr [0040103c]
'007132e6    8bd0                    mov edx, eax
'007132e8    8d4de4                  lea ecx, dword ptr [ebp-1c]
'007132eb    ffd7                    call edi
'007132ed    8bbd20ffffff            mov edi, dword ptr [ebp+ffffff20]
'007132f3    50                      push eax
'007132f4    57                      push edi
'007132f5    ff93ac000000            call dword ptr [ebx+000000ac]
'007132fb    dbe2                    fnclex
'007132fd    3bc6                    cmp eax, esi
'007132ff    7d12                    jge 713313

If (-4508 < 0) Then
'00713301    68ac000000              push 000000ac
'00713306    681cb94300              push 0043b91c
'0071330b    57                      push edi
'0071330c    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0071330d    ff1580104000            call dword ptr [00401080]
    
End If
'00713313    8d4de4                  lea ecx, dword ptr [ebp-1c]

' *** Reference to "__vbaFreeStr"
'00713316    ff150c134000            call dword ptr [0040130c]
'#Cleanup(var_40)
'0071331c    8d55c4                  lea edx, dword ptr [ebp-3c]
'0071331f    52                      push edx
'00713320    8d45c8                  lea eax, dword ptr [ebp-38]
'00713323    50                      push eax
'00713324    8d4dcc                  lea ecx, dword ptr [ebp-34]
'00713327    51                      push ecx
'00713328    6a03                    push 03

' *** Reference to "__vbaFreeObjList"
'0071332a    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_43, var_46, var_9)
'00713330    83c410                  add esp, 10
'00713333    8d4dac                  lea ecx, dword ptr [ebp-54]

' *** Reference to "__vbaFreeVar"
'00713336    ff1528104000            call dword ptr [00401028]
'#Cleanup(var_50)
'0071333c    8b8514ffffff            mov eax, dword ptr [ebp+ffffff14]
'00713342    8b10                    mov edx, dword ptr [eax]
'00713344    8d4dcc                  lea ecx, dword ptr [ebp-34]
'00713347    51                      push ecx
'00713348    50                      push eax
'00713349    ff92b4000000            call dword ptr [edx+000000b4]
'0071334f    dbe2                    fnclex
'00713351    3bc6                    cmp eax, esi
'00713353    7d18                    jge 71336d

If (var_488 < 0) Then
'00713355    8b9514ffffff            mov edx, dword ptr [ebp+ffffff14]
'0071335b    68b4000000              push 000000b4
'00713360    6830314300              push 00433130
'00713365    52                      push edx
'00713366    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00713367    ff1580104000            call dword ptr [00401080]
    
End If
'0071336d    8b45cc                  mov eax, dword ptr [ebp-34]
'00713370    8b38                    mov edi, dword ptr [eax]
'00713372    8d5dc8                  lea ebx, dword ptr [ebp-38]
'00713375    53                      push ebx
'00713376    83ec10                  sub esp, 10
'00713379    8bdc                    mov ebx, esp
'0071337b    ba08000000              mov edx, 00000008
'00713380    8913                    mov dword ptr [ebx], edx
'00713382    8b9570ffffff            mov edx, dword ptr [ebp+ffffff70]
'00713388    895304                  mov dword ptr [ebx+04], edx
'0071338b    b9bcb24300              mov ecx, 0043b2bc
'00713390    894b08                  mov dword ptr [ebx+08], ecx
'00713393    8b8d78ffffff            mov ecx, dword ptr [ebp+ffffff78]
'00713399    50                      push eax
'0071339a    898530ffffff            mov dword ptr [ebp+ffffff30], eax
'007133a0    894b0c                  mov dword ptr [ebx+0c], ecx
'007133a3    ff5730                  call dword ptr [edi+30]
'007133a6    dbe2                    fnclex
'007133a8    3bc6                    cmp eax, esi
'007133aa    7d15                    jge 7133c1

If (var_43 < 0) Then
'007133ac    8b9530ffffff            mov edx, dword ptr [ebp+ffffff30]
'007133b2    6a30                    push 30
'007133b4    68d8304300              push 004330d8
'007133b9    52                      push edx
'007133ba    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'007133bb    ff1580104000            call dword ptr [00401080]
    
End If
'007133c1    8b45c8                  mov eax, dword ptr [ebp-38]
'007133c4    8945b4                  mov dword ptr [ebp-4c], eax
'007133c7    8d45ac                  lea eax, dword ptr [ebp-54]
'007133ca    50                      push eax
'007133cb    8975c8                  mov dword ptr [ebp-38], esi
'007133ce    c745ac09000000          mov dword ptr [ebp-54], 00000009

' *** Reference to "rtcIsNull"
'007133d5    ff1540114000            call dword ptr [00401140]
'007133db    898538ffffff            mov dword ptr [ebp+ffffff38], eax
'007133e1    8b4508                  mov eax, dword ptr [ebp+08]
'007133e4    8b08                    mov ecx, dword ptr [eax]
'007133e6    50                      push eax
'007133e7    ff9104030000            call dword ptr [ecx+00000304]
'007133ed    50                      push eax
'007133ee    8d55bc                  lea edx, dword ptr [ebp-44]
'007133f1    52                      push edx

' *** Reference to "__vbaObjSet"
'007133f2    ff15b4104000            call dword ptr [004010b4]
Set var_58 = Me
'007133f8    8d55c4                  lea edx, dword ptr [ebp-3c]
'007133fb    8bf8                    mov edi, eax
'007133fd    8b8514ffffff            mov eax, dword ptr [ebp+ffffff14]
'00713403    8b08                    mov ecx, dword ptr [eax]
'00713405    52                      push edx
'00713406    50                      push eax
'00713407    ff91b4000000            call dword ptr [ecx+000000b4]
'0071340d    dbe2                    fnclex
'0071340f    3bc6                    cmp eax, esi
'00713411    7d18                    jge 71342b

If (var_488 < 0) Then
'00713413    8b8d14ffffff            mov ecx, dword ptr [ebp+ffffff14]
'00713419    68b4000000              push 000000b4
'0071341e    6830314300              push 00433130
'00713423    51                      push ecx
'00713424    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00713425    ff1580104000            call dword ptr [00401080]
    
End If
'0071342b    8b45c4                  mov eax, dword ptr [ebp-3c]
'0071342e    8b10                    mov edx, dword ptr [eax]
'00713430    8d5dc0                  lea ebx, dword ptr [ebp-40]
'00713433    53                      push ebx
'00713434    83ec10                  sub esp, 10
'00713437    8bdc                    mov ebx, esp
'00713439    b908000000              mov ecx, 00000008
'0071343e    890b                    mov dword ptr [ebx], ecx
'00713440    8b8d60ffffff            mov ecx, dword ptr [ebp+ffffff60]
'00713446    894b04                  mov dword ptr [ebx+04], ecx
'00713449    c78564ffffffbcb24300    mov dword ptr [ebp+ffffff64], 0043b2bc
'00713453    8b8d64ffffff            mov ecx, dword ptr [ebp+ffffff64]
'00713459    894b08                  mov dword ptr [ebx+08], ecx
'0071345c    8b8d68ffffff            mov ecx, dword ptr [ebp+ffffff68]
'00713462    50                      push eax
'00713463    898524ffffff            mov dword ptr [ebp+ffffff24], eax
'00713469    894b0c                  mov dword ptr [ebx+0c], ecx
'0071346c    ff5230                  call dword ptr [edx+30]
'0071346f    dbe2                    fnclex
'00713471    3bc6                    cmp eax, esi
'00713473    7d15                    jge 71348a

If (var_9 < 0) Then
'00713475    8b9524ffffff            mov edx, dword ptr [ebp+ffffff24]
'0071347b    6a30                    push 30
'0071347d    68d8304300              push 004330d8
'00713482    52                      push edx
'00713483    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00713484    ff1580104000            call dword ptr [00401080]
    
End If
'0071348a    8b45c0                  mov eax, dword ptr [ebp-40]
'0071348d    8d953cffffff            lea edx, dword ptr [ebp+ffffff3c]
'00713493    8d4d9c                  lea ecx, dword ptr [ebp-64]
'00713496    8975c0                  mov dword ptr [ebp-40], esi
'00713499    894594                  mov dword ptr [ebp-6c], eax
'0071349c    c7458c09000000          mov dword ptr [ebp-74], 00000009
'007134a3    c78544ffffffcc134300    mov dword ptr [ebp+ffffff44], 004313cc
'007134ad    c7853cffffff08000000    mov dword ptr [ebp+ffffff3c], 00000008

' *** Reference to "__vbaVarDup"
'007134b7    ff1598124000            call dword ptr [00401298]
var_51 = (vbNullChar)
'007134bd    668b8538ffffff          mov ax, word ptr [ebp+ffffff38]
'007134c4    8d4d8c                  lea ecx, dword ptr [ebp-74]
'007134c7    51                      push ecx
'007134c8    8d559c                  lea edx, dword ptr [ebp-64]
'007134cb    66898554ffffff          mov word ptr [ebp+ffffff54], ax
'007134d2    52                      push edx
'007134d3    8d854cffffff            lea eax, dword ptr [ebp+ffffff4c]
'007134d9    50                      push eax
'007134da    8d8d7cffffff            lea ecx, dword ptr [ebp+ffffff7c]
'007134e0    51                      push ecx
'007134e1    c7854cffffff0b000000    mov dword ptr [ebp+ffffff4c], 0000000b

' *** Reference to "rtcImmediateIf"
'007134eb    ff154c124000            call dword ptr [0040124c]
'007134f1    8b1f                    mov ebx, dword ptr [edi]
'007134f3    8d957cffffff            lea edx, dword ptr [ebp+ffffff7c]
'007134f9    52                      push edx
'007134fa    8d45e4                  lea eax, dword ptr [ebp-1c]
'007134fd    50                      push eax

' *** Reference to "__vbaStrVarVal"
'007134fe    ff15fc114000            call dword ptr [004011fc]
'00713504    50                      push eax
'00713505    57                      push edi
'00713506    ff93ac000000            call dword ptr [ebx+000000ac]
'0071350c    dbe2                    fnclex
'0071350e    3bc6                    cmp eax, esi
'00713510    7d12                    jge 713524
'00713512    68ac000000              push 000000ac
'00713517    681cb94300              push 0043b91c
'0071351c    57                      push edi
'0071351d    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0071351e    ff1580104000            call dword ptr [00401080]
'00713524    8d4de4                  lea ecx, dword ptr [ebp-1c]

' *** Reference to "__vbaFreeStr"
'00713527    ff150c134000            call dword ptr [0040130c]
'#Cleanup(var_40)
'0071352d    8d4dbc                  lea ecx, dword ptr [ebp-44]
'00713530    51                      push ecx
'00713531    8d55c4                  lea edx, dword ptr [ebp-3c]
'00713534    52                      push edx
'00713535    8d45cc                  lea eax, dword ptr [ebp-34]
'00713538    50                      push eax
'00713539    6a03                    push 03

' *** Reference to "__vbaFreeObjList"
'0071353b    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_43, var_9, var_58)
'00713541    8d8d7cffffff            lea ecx, dword ptr [ebp+ffffff7c]
'00713547    51                      push ecx
'00713548    8d558c                  lea edx, dword ptr [ebp-74]
'0071354b    52                      push edx
'0071354c    8d459c                  lea eax, dword ptr [ebp-64]
'0071354f    50                      push eax
'00713550    8d8d4cffffff            lea ecx, dword ptr [ebp+ffffff4c]
'00713556    51                      push ecx
'00713557    8d55ac                  lea edx, dword ptr [ebp-54]
'0071355a    52                      push edx
'0071355b    6a05                    push 05

' *** Reference to "__vbaFreeVarList"
'0071355d    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_50, var_89, var_51, var_52, var_59)
'00713563    8b8514ffffff            mov eax, dword ptr [ebp+ffffff14]
'00713569    8b08                    mov ecx, dword ptr [eax]
'0071356b    83c428                  add esp, 28
'0071356e    8d55cc                  lea edx, dword ptr [ebp-34]
'00713571    52                      push edx
'00713572    50                      push eax
'00713573    ff91b4000000            call dword ptr [ecx+000000b4]
'00713579    dbe2                    fnclex
'0071357b    3bc6                    cmp eax, esi
'0071357d    7d18                    jge 713597

If (var_488 < 0) Then
'0071357f    8b8d14ffffff            mov ecx, dword ptr [ebp+ffffff14]
'00713585    68b4000000              push 000000b4
'0071358a    6830314300              push 00433130
'0071358f    51                      push ecx
'00713590    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00713591    ff1580104000            call dword ptr [00401080]
    
End If
'00713597    8b45cc                  mov eax, dword ptr [ebp-34]
'0071359a    8b38                    mov edi, dword ptr [eax]
'0071359c    8d5dc8                  lea ebx, dword ptr [ebp-38]
'0071359f    53                      push ebx
'007135a0    83ec10                  sub esp, 10
'007135a3    8bdc                    mov ebx, esp
'007135a5    ba08000000              mov edx, 00000008
'007135aa    8913                    mov dword ptr [ebx], edx
'007135ac    8b9570ffffff            mov edx, dword ptr [ebp+ffffff70]
'007135b2    895304                  mov dword ptr [ebx+04], edx
'007135b5    b9a0b84300              mov ecx, 0043b8a0
'007135ba    894b08                  mov dword ptr [ebx+08], ecx
'007135bd    8b8d78ffffff            mov ecx, dword ptr [ebp+ffffff78]
'007135c3    50                      push eax
'007135c4    898530ffffff            mov dword ptr [ebp+ffffff30], eax
'007135ca    894b0c                  mov dword ptr [ebx+0c], ecx
'007135cd    ff5730                  call dword ptr [edi+30]
'007135d0    dbe2                    fnclex
'007135d2    3bc6                    cmp eax, esi
'007135d4    7d15                    jge 7135eb

If (var_43 < 0) Then
'007135d6    8b9530ffffff            mov edx, dword ptr [ebp+ffffff30]
'007135dc    6a30                    push 30
'007135de    68d8304300              push 004330d8
'007135e3    52                      push edx
'007135e4    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'007135e5    ff1580104000            call dword ptr [00401080]
    
End If
'007135eb    8b45c8                  mov eax, dword ptr [ebp-38]
'007135ee    8945b4                  mov dword ptr [ebp-4c], eax
'007135f1    8d45ac                  lea eax, dword ptr [ebp-54]
'007135f4    50                      push eax
'007135f5    8975c8                  mov dword ptr [ebp-38], esi
'007135f8    c745ac09000000          mov dword ptr [ebp-54], 00000009

' *** Reference to "rtcIsNull"
'007135ff    ff1540114000            call dword ptr [00401140]
'00713605    898538ffffff            mov dword ptr [ebp+ffffff38], eax
'0071360b    8b4508                  mov eax, dword ptr [ebp+08]
'0071360e    8b08                    mov ecx, dword ptr [eax]
'00713610    50                      push eax
'00713611    ff9134030000            call dword ptr [ecx+00000334]
'00713617    50                      push eax
'00713618    8d55bc                  lea edx, dword ptr [ebp-44]
'0071361b    52                      push edx

' *** Reference to "__vbaObjSet"
'0071361c    ff15b4104000            call dword ptr [004010b4]
Set var_58 = Me
'00713622    8d55c4                  lea edx, dword ptr [ebp-3c]
'00713625    8bf8                    mov edi, eax
'00713627    8b8514ffffff            mov eax, dword ptr [ebp+ffffff14]
'0071362d    8b08                    mov ecx, dword ptr [eax]
'0071362f    52                      push edx
'00713630    50                      push eax
'00713631    ff91b4000000            call dword ptr [ecx+000000b4]
'00713637    dbe2                    fnclex
'00713639    3bc6                    cmp eax, esi
'0071363b    7d18                    jge 713655

If (var_488 < 0) Then
'0071363d    8b8d14ffffff            mov ecx, dword ptr [ebp+ffffff14]
'00713643    68b4000000              push 000000b4
'00713648    6830314300              push 00433130
'0071364d    51                      push ecx
'0071364e    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0071364f    ff1580104000            call dword ptr [00401080]
    
End If
'00713655    8b45c4                  mov eax, dword ptr [ebp-3c]
'00713658    8b10                    mov edx, dword ptr [eax]
'0071365a    8d5dc0                  lea ebx, dword ptr [ebp-40]
'0071365d    53                      push ebx
'0071365e    83ec10                  sub esp, 10
'00713661    8bdc                    mov ebx, esp
'00713663    b908000000              mov ecx, 00000008
'00713668    890b                    mov dword ptr [ebx], ecx
'0071366a    8b8d60ffffff            mov ecx, dword ptr [ebp+ffffff60]
'00713670    894b04                  mov dword ptr [ebx+04], ecx
'00713673    c78564ffffffa0b84300    mov dword ptr [ebp+ffffff64], 0043b8a0
'0071367d    8b8d64ffffff            mov ecx, dword ptr [ebp+ffffff64]
'00713683    894b08                  mov dword ptr [ebx+08], ecx
'00713686    8b8d68ffffff            mov ecx, dword ptr [ebp+ffffff68]
'0071368c    50                      push eax
'0071368d    898524ffffff            mov dword ptr [ebp+ffffff24], eax
'00713693    894b0c                  mov dword ptr [ebx+0c], ecx
'00713696    ff5230                  call dword ptr [edx+30]
'00713699    dbe2                    fnclex
'0071369b    3bc6                    cmp eax, esi
'0071369d    7d15                    jge 7136b4

If (var_9 < 0) Then
'0071369f    8b9524ffffff            mov edx, dword ptr [ebp+ffffff24]
'007136a5    6a30                    push 30
'007136a7    68d8304300              push 004330d8
'007136ac    52                      push edx
'007136ad    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'007136ae    ff1580104000            call dword ptr [00401080]
    
End If
'007136b4    8b45c0                  mov eax, dword ptr [ebp-40]
'007136b7    8d953cffffff            lea edx, dword ptr [ebp+ffffff3c]
'007136bd    8d4d9c                  lea ecx, dword ptr [ebp-64]
'007136c0    8975c0                  mov dword ptr [ebp-40], esi
'007136c3    894594                  mov dword ptr [ebp-6c], eax
'007136c6    c7458c09000000          mov dword ptr [ebp-74], 00000009
'007136cd    c78544ffffffcc134300    mov dword ptr [ebp+ffffff44], 004313cc
'007136d7    c7853cffffff08000000    mov dword ptr [ebp+ffffff3c], 00000008

' *** Reference to "__vbaVarDup"
'007136e1    ff1598124000            call dword ptr [00401298]
var_51 = (vbNullChar)
'007136e7    668b8538ffffff          mov ax, word ptr [ebp+ffffff38]
'007136ee    8d4d8c                  lea ecx, dword ptr [ebp-74]
'007136f1    51                      push ecx
'007136f2    8d559c                  lea edx, dword ptr [ebp-64]
'007136f5    66898554ffffff          mov word ptr [ebp+ffffff54], ax
'007136fc    52                      push edx
'007136fd    8d854cffffff            lea eax, dword ptr [ebp+ffffff4c]
'00713703    50                      push eax
'00713704    8d8d7cffffff            lea ecx, dword ptr [ebp+ffffff7c]
'0071370a    51                      push ecx
'0071370b    c7854cffffff0b000000    mov dword ptr [ebp+ffffff4c], 0000000b

' *** Reference to "rtcImmediateIf"
'00713715    ff154c124000            call dword ptr [0040124c]
'0071371b    8b1f                    mov ebx, dword ptr [edi]
'0071371d    8d957cffffff            lea edx, dword ptr [ebp+ffffff7c]
'00713723    52                      push edx
'00713724    8d45e4                  lea eax, dword ptr [ebp-1c]
'00713727    50                      push eax

' *** Reference to "__vbaStrVarVal"
'00713728    ff15fc114000            call dword ptr [004011fc]
'0071372e    50                      push eax
'0071372f    57                      push edi
'00713730    ff93a4000000            call dword ptr [ebx+000000a4]
'00713736    dbe2                    fnclex
'00713738    3bc6                    cmp eax, esi
'0071373a    7d12                    jge 71374e
'0071373c    68a4000000              push 000000a4
'00713741    68d00d4300              push 00430dd0
'00713746    57                      push edi
'00713747    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00713748    ff1580104000            call dword ptr [00401080]
'0071374e    8d4de4                  lea ecx, dword ptr [ebp-1c]

' *** Reference to "__vbaFreeStr"
'00713751    ff150c134000            call dword ptr [0040130c]
'#Cleanup(var_40)
'00713757    8d4dbc                  lea ecx, dword ptr [ebp-44]
'0071375a    51                      push ecx
'0071375b    8d55c4                  lea edx, dword ptr [ebp-3c]
'0071375e    52                      push edx
'0071375f    8d45cc                  lea eax, dword ptr [ebp-34]
'00713762    50                      push eax
'00713763    6a03                    push 03

' *** Reference to "__vbaFreeObjList"
'00713765    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_43, var_9, var_58)
'0071376b    8d8d7cffffff            lea ecx, dword ptr [ebp+ffffff7c]
'00713771    51                      push ecx
'00713772    8d558c                  lea edx, dword ptr [ebp-74]
'00713775    52                      push edx
'00713776    8d459c                  lea eax, dword ptr [ebp-64]
'00713779    50                      push eax
'0071377a    8d8d4cffffff            lea ecx, dword ptr [ebp+ffffff4c]
'00713780    51                      push ecx
'00713781    8d55ac                  lea edx, dword ptr [ebp-54]
'00713784    52                      push edx
'00713785    6a05                    push 05

' *** Reference to "__vbaFreeVarList"
'00713787    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_50, var_89, var_51, var_52, var_59)
'0071378d    8b8514ffffff            mov eax, dword ptr [ebp+ffffff14]
'00713793    8b08                    mov ecx, dword ptr [eax]
'00713795    83c428                  add esp, 28
'00713798    8d55cc                  lea edx, dword ptr [ebp-34]
'0071379b    52                      push edx
'0071379c    50                      push eax
'0071379d    ff91b4000000            call dword ptr [ecx+000000b4]
'007137a3    dbe2                    fnclex
'007137a5    3bc6                    cmp eax, esi
'007137a7    7d18                    jge 7137c1

If (var_488 < 0) Then
'007137a9    8b8d14ffffff            mov ecx, dword ptr [ebp+ffffff14]
'007137af    68b4000000              push 000000b4
'007137b4    6830314300              push 00433130
'007137b9    51                      push ecx
'007137ba    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'007137bb    ff1580104000            call dword ptr [00401080]
    
End If
'007137c1    8b45cc                  mov eax, dword ptr [ebp-34]
'007137c4    8b38                    mov edi, dword ptr [eax]
'007137c6    8d5dc8                  lea ebx, dword ptr [ebp-38]
'007137c9    53                      push ebx
'007137ca    83ec10                  sub esp, 10
'007137cd    8bdc                    mov ebx, esp
'007137cf    ba08000000              mov edx, 00000008
'007137d4    8913                    mov dword ptr [ebx], edx
'007137d6    8b9570ffffff            mov edx, dword ptr [ebp+ffffff70]
'007137dc    895304                  mov dword ptr [ebx+04], edx
'007137df    b9e4b24300              mov ecx, 0043b2e4
'007137e4    894b08                  mov dword ptr [ebx+08], ecx
'007137e7    8b8d78ffffff            mov ecx, dword ptr [ebp+ffffff78]
'007137ed    50                      push eax
'007137ee    898530ffffff            mov dword ptr [ebp+ffffff30], eax
'007137f4    894b0c                  mov dword ptr [ebx+0c], ecx
'007137f7    ff5730                  call dword ptr [edi+30]
'007137fa    dbe2                    fnclex
'007137fc    3bc6                    cmp eax, esi
'007137fe    7d15                    jge 713815

If (var_43 < 0) Then
'00713800    8b9530ffffff            mov edx, dword ptr [ebp+ffffff30]
'00713806    6a30                    push 30
'00713808    68d8304300              push 004330d8
'0071380d    52                      push edx
'0071380e    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0071380f    ff1580104000            call dword ptr [00401080]
    
End If
'00713815    8b45c8                  mov eax, dword ptr [ebp-38]
'00713818    8945b4                  mov dword ptr [ebp-4c], eax
'0071381b    8d45ac                  lea eax, dword ptr [ebp-54]
'0071381e    50                      push eax
'0071381f    8975c8                  mov dword ptr [ebp-38], esi
'00713822    c745ac09000000          mov dword ptr [ebp-54], 00000009

' *** Reference to "rtcIsNull"
'00713829    ff1540114000            call dword ptr [00401140]
'0071382f    898538ffffff            mov dword ptr [ebp+ffffff38], eax
'00713835    8b4508                  mov eax, dword ptr [ebp+08]
'00713838    8b08                    mov ecx, dword ptr [eax]
'0071383a    50                      push eax
'0071383b    ff9138030000            call dword ptr [ecx+00000338]
'00713841    50                      push eax
'00713842    8d55bc                  lea edx, dword ptr [ebp-44]
'00713845    52                      push edx

' *** Reference to "__vbaObjSet"
'00713846    ff15b4104000            call dword ptr [004010b4]
Set var_58 = Me
'0071384c    8d55c4                  lea edx, dword ptr [ebp-3c]
'0071384f    8bf8                    mov edi, eax
'00713851    8b8514ffffff            mov eax, dword ptr [ebp+ffffff14]
'00713857    8b08                    mov ecx, dword ptr [eax]
'00713859    52                      push edx
'0071385a    50                      push eax
'0071385b    ff91b4000000            call dword ptr [ecx+000000b4]
'00713861    dbe2                    fnclex
'00713863    3bc6                    cmp eax, esi
'00713865    7d18                    jge 71387f

If (var_488 < 0) Then
'00713867    8b8d14ffffff            mov ecx, dword ptr [ebp+ffffff14]
'0071386d    68b4000000              push 000000b4
'00713872    6830314300              push 00433130
'00713877    51                      push ecx
'00713878    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00713879    ff1580104000            call dword ptr [00401080]
    
End If
'0071387f    8b45c4                  mov eax, dword ptr [ebp-3c]
'00713882    8b10                    mov edx, dword ptr [eax]
'00713884    8d5dc0                  lea ebx, dword ptr [ebp-40]
'00713887    53                      push ebx
'00713888    83ec10                  sub esp, 10
'0071388b    8bdc                    mov ebx, esp
'0071388d    b908000000              mov ecx, 00000008
'00713892    890b                    mov dword ptr [ebx], ecx
'00713894    8b8d60ffffff            mov ecx, dword ptr [ebp+ffffff60]
'0071389a    894b04                  mov dword ptr [ebx+04], ecx
'0071389d    c78564ffffffe4b24300    mov dword ptr [ebp+ffffff64], 0043b2e4
'007138a7    8b8d64ffffff            mov ecx, dword ptr [ebp+ffffff64]
'007138ad    894b08                  mov dword ptr [ebx+08], ecx
'007138b0    8b8d68ffffff            mov ecx, dword ptr [ebp+ffffff68]
'007138b6    50                      push eax
'007138b7    898524ffffff            mov dword ptr [ebp+ffffff24], eax
'007138bd    894b0c                  mov dword ptr [ebx+0c], ecx
'007138c0    ff5230                  call dword ptr [edx+30]
'007138c3    dbe2                    fnclex
'007138c5    3bc6                    cmp eax, esi
'007138c7    7d15                    jge 7138de

If (var_9 < 0) Then
'007138c9    8b9524ffffff            mov edx, dword ptr [ebp+ffffff24]
'007138cf    6a30                    push 30
'007138d1    68d8304300              push 004330d8
'007138d6    52                      push edx
'007138d7    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'007138d8    ff1580104000            call dword ptr [00401080]
    
End If
'007138de    8b45c0                  mov eax, dword ptr [ebp-40]
'007138e1    8d953cffffff            lea edx, dword ptr [ebp+ffffff3c]
'007138e7    8d4d9c                  lea ecx, dword ptr [ebp-64]
'007138ea    8975c0                  mov dword ptr [ebp-40], esi
'007138ed    894594                  mov dword ptr [ebp-6c], eax
'007138f0    c7458c09000000          mov dword ptr [ebp-74], 00000009
'007138f7    c78544ffffffcc134300    mov dword ptr [ebp+ffffff44], 004313cc
'00713901    c7853cffffff08000000    mov dword ptr [ebp+ffffff3c], 00000008

' *** Reference to "__vbaVarDup"
'0071390b    ff1598124000            call dword ptr [00401298]
var_51 = (vbNullChar)
'00713911    668b8538ffffff          mov ax, word ptr [ebp+ffffff38]
'00713918    8d4d8c                  lea ecx, dword ptr [ebp-74]
'0071391b    51                      push ecx
'0071391c    8d559c                  lea edx, dword ptr [ebp-64]
'0071391f    66898554ffffff          mov word ptr [ebp+ffffff54], ax
'00713926    52                      push edx
'00713927    8d854cffffff            lea eax, dword ptr [ebp+ffffff4c]
'0071392d    50                      push eax
'0071392e    8d8d7cffffff            lea ecx, dword ptr [ebp+ffffff7c]
'00713934    51                      push ecx
'00713935    c7854cffffff0b000000    mov dword ptr [ebp+ffffff4c], 0000000b

' *** Reference to "rtcImmediateIf"
'0071393f    ff154c124000            call dword ptr [0040124c]
'00713945    8b1f                    mov ebx, dword ptr [edi]
'00713947    8d957cffffff            lea edx, dword ptr [ebp+ffffff7c]
'0071394d    52                      push edx
'0071394e    8d45e4                  lea eax, dword ptr [ebp-1c]
'00713951    50                      push eax

' *** Reference to "__vbaStrVarVal"
'00713952    ff15fc114000            call dword ptr [004011fc]
'00713958    50                      push eax
'00713959    57                      push edi
'0071395a    ff93a4000000            call dword ptr [ebx+000000a4]
'00713960    dbe2                    fnclex
'00713962    3bc6                    cmp eax, esi
'00713964    7d12                    jge 713978
'00713966    68a4000000              push 000000a4
'0071396b    68d00d4300              push 00430dd0
'00713970    57                      push edi
'00713971    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00713972    ff1580104000            call dword ptr [00401080]
'00713978    8d4de4                  lea ecx, dword ptr [ebp-1c]

' *** Reference to "__vbaFreeStr"
'0071397b    ff150c134000            call dword ptr [0040130c]
'#Cleanup(var_40)
'00713981    8d4dbc                  lea ecx, dword ptr [ebp-44]
'00713984    51                      push ecx
'00713985    8d55c4                  lea edx, dword ptr [ebp-3c]
'00713988    52                      push edx
'00713989    8d45cc                  lea eax, dword ptr [ebp-34]
'0071398c    50                      push eax
'0071398d    6a03                    push 03

' *** Reference to "__vbaFreeObjList"
'0071398f    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_43, var_9, var_58)
'00713995    8d8d7cffffff            lea ecx, dword ptr [ebp+ffffff7c]
'0071399b    51                      push ecx
'0071399c    8d558c                  lea edx, dword ptr [ebp-74]
'0071399f    52                      push edx
'007139a0    8d459c                  lea eax, dword ptr [ebp-64]
'007139a3    50                      push eax
'007139a4    8d8d4cffffff            lea ecx, dword ptr [ebp+ffffff4c]
'007139aa    51                      push ecx
'007139ab    8d55ac                  lea edx, dword ptr [ebp-54]
'007139ae    52                      push edx
'007139af    6a05                    push 05

' *** Reference to "__vbaFreeVarList"
'007139b1    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_50, var_89, var_51, var_52, var_59)
'007139b7    8b8514ffffff            mov eax, dword ptr [ebp+ffffff14]
'007139bd    8b08                    mov ecx, dword ptr [eax]
'007139bf    83c428                  add esp, 28
'007139c2    8d55cc                  lea edx, dword ptr [ebp-34]
'007139c5    52                      push edx
'007139c6    50                      push eax
'007139c7    ff91b4000000            call dword ptr [ecx+000000b4]
'007139cd    dbe2                    fnclex
'007139cf    3bc6                    cmp eax, esi
'007139d1    7d18                    jge 7139eb

If (var_488 < 0) Then
'007139d3    8b8d14ffffff            mov ecx, dword ptr [ebp+ffffff14]
'007139d9    68b4000000              push 000000b4
'007139de    6830314300              push 00433130
'007139e3    51                      push ecx
'007139e4    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'007139e5    ff1580104000            call dword ptr [00401080]
    
End If
'007139eb    8b45cc                  mov eax, dword ptr [ebp-34]
'007139ee    8b38                    mov edi, dword ptr [eax]
'007139f0    8d5dc8                  lea ebx, dword ptr [ebp-38]
'007139f3    53                      push ebx
'007139f4    83ec10                  sub esp, 10
'007139f7    8bdc                    mov ebx, esp
'007139f9    ba08000000              mov edx, 00000008
'007139fe    8913                    mov dword ptr [ebx], edx
'00713a00    8b9570ffffff            mov edx, dword ptr [ebp+ffffff70]
'00713a06    895304                  mov dword ptr [ebx+04], edx
'00713a09    b900b34300              mov ecx, 0043b300
'00713a0e    894b08                  mov dword ptr [ebx+08], ecx
'00713a11    8b8d78ffffff            mov ecx, dword ptr [ebp+ffffff78]
'00713a17    50                      push eax
'00713a18    898530ffffff            mov dword ptr [ebp+ffffff30], eax
'00713a1e    894b0c                  mov dword ptr [ebx+0c], ecx
'00713a21    ff5730                  call dword ptr [edi+30]
'00713a24    dbe2                    fnclex
'00713a26    3bc6                    cmp eax, esi
'00713a28    7d15                    jge 713a3f

If (var_43 < 0) Then
'00713a2a    8b9530ffffff            mov edx, dword ptr [ebp+ffffff30]
'00713a30    6a30                    push 30
'00713a32    68d8304300              push 004330d8
'00713a37    52                      push edx
'00713a38    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00713a39    ff1580104000            call dword ptr [00401080]
    
End If
'00713a3f    8b45c8                  mov eax, dword ptr [ebp-38]
'00713a42    8945b4                  mov dword ptr [ebp-4c], eax
'00713a45    8d45ac                  lea eax, dword ptr [ebp-54]
'00713a48    50                      push eax
'00713a49    8975c8                  mov dword ptr [ebp-38], esi
'00713a4c    c745ac09000000          mov dword ptr [ebp-54], 00000009

' *** Reference to "rtcIsNull"
'00713a53    ff1540114000            call dword ptr [00401140]
'00713a59    898538ffffff            mov dword ptr [ebp+ffffff38], eax
'00713a5f    8b4508                  mov eax, dword ptr [ebp+08]
'00713a62    8b08                    mov ecx, dword ptr [eax]
'00713a64    50                      push eax
'00713a65    ff91fc020000            call dword ptr [ecx+000002fc]
'00713a6b    50                      push eax
'00713a6c    8d55bc                  lea edx, dword ptr [ebp-44]
'00713a6f    52                      push edx

' *** Reference to "__vbaObjSet"
'00713a70    ff15b4104000            call dword ptr [004010b4]
Set var_58 = Me
'00713a76    8d55c4                  lea edx, dword ptr [ebp-3c]
'00713a79    8bf8                    mov edi, eax
'00713a7b    8b8514ffffff            mov eax, dword ptr [ebp+ffffff14]
'00713a81    8b08                    mov ecx, dword ptr [eax]
'00713a83    52                      push edx
'00713a84    50                      push eax
'00713a85    ff91b4000000            call dword ptr [ecx+000000b4]
'00713a8b    dbe2                    fnclex
'00713a8d    3bc6                    cmp eax, esi
'00713a8f    7d18                    jge 713aa9

If (var_488 < 0) Then
'00713a91    8b8d14ffffff            mov ecx, dword ptr [ebp+ffffff14]
'00713a97    68b4000000              push 000000b4
'00713a9c    6830314300              push 00433130
'00713aa1    51                      push ecx
'00713aa2    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00713aa3    ff1580104000            call dword ptr [00401080]
    
End If
'00713aa9    8b45c4                  mov eax, dword ptr [ebp-3c]
'00713aac    8b10                    mov edx, dword ptr [eax]
'00713aae    8d5dc0                  lea ebx, dword ptr [ebp-40]
'00713ab1    53                      push ebx
'00713ab2    83ec10                  sub esp, 10
'00713ab5    8bdc                    mov ebx, esp
'00713ab7    b908000000              mov ecx, 00000008
'00713abc    890b                    mov dword ptr [ebx], ecx
'00713abe    8b8d60ffffff            mov ecx, dword ptr [ebp+ffffff60]
'00713ac4    894b04                  mov dword ptr [ebx+04], ecx
'00713ac7    c78564ffffff00b34300    mov dword ptr [ebp+ffffff64], 0043b300
'00713ad1    8b8d64ffffff            mov ecx, dword ptr [ebp+ffffff64]
'00713ad7    894b08                  mov dword ptr [ebx+08], ecx
'00713ada    8b8d68ffffff            mov ecx, dword ptr [ebp+ffffff68]
'00713ae0    50                      push eax
'00713ae1    898524ffffff            mov dword ptr [ebp+ffffff24], eax
'00713ae7    894b0c                  mov dword ptr [ebx+0c], ecx
'00713aea    ff5230                  call dword ptr [edx+30]
'00713aed    dbe2                    fnclex
'00713aef    3bc6                    cmp eax, esi
'00713af1    7d15                    jge 713b08

If (var_9 < 0) Then
'00713af3    8b9524ffffff            mov edx, dword ptr [ebp+ffffff24]
'00713af9    6a30                    push 30
'00713afb    68d8304300              push 004330d8
'00713b00    52                      push edx
'00713b01    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00713b02    ff1580104000            call dword ptr [00401080]
    
End If
'00713b08    8b45c0                  mov eax, dword ptr [ebp-40]
'00713b0b    8d953cffffff            lea edx, dword ptr [ebp+ffffff3c]
'00713b11    8d4d9c                  lea ecx, dword ptr [ebp-64]
'00713b14    8975c0                  mov dword ptr [ebp-40], esi
'00713b17    894594                  mov dword ptr [ebp-6c], eax
'00713b1a    c7458c09000000          mov dword ptr [ebp-74], 00000009
'00713b21    c78544ffffffcc134300    mov dword ptr [ebp+ffffff44], 004313cc
'00713b2b    c7853cffffff08000000    mov dword ptr [ebp+ffffff3c], 00000008

' *** Reference to "__vbaVarDup"
'00713b35    ff1598124000            call dword ptr [00401298]
var_51 = (vbNullChar)
'00713b3b    668b8538ffffff          mov ax, word ptr [ebp+ffffff38]
'00713b42    8d4d8c                  lea ecx, dword ptr [ebp-74]
'00713b45    51                      push ecx
'00713b46    8d559c                  lea edx, dword ptr [ebp-64]
'00713b49    66898554ffffff          mov word ptr [ebp+ffffff54], ax
'00713b50    52                      push edx
'00713b51    8d854cffffff            lea eax, dword ptr [ebp+ffffff4c]
'00713b57    50                      push eax
'00713b58    8d8d7cffffff            lea ecx, dword ptr [ebp+ffffff7c]
'00713b5e    51                      push ecx
'00713b5f    c7854cffffff0b000000    mov dword ptr [ebp+ffffff4c], 0000000b

' *** Reference to "rtcImmediateIf"
'00713b69    ff154c124000            call dword ptr [0040124c]
'00713b6f    8b1f                    mov ebx, dword ptr [edi]
'00713b71    8d957cffffff            lea edx, dword ptr [ebp+ffffff7c]
'00713b77    52                      push edx
'00713b78    8d45e4                  lea eax, dword ptr [ebp-1c]
'00713b7b    50                      push eax

' *** Reference to "__vbaStrVarVal"
'00713b7c    ff15fc114000            call dword ptr [004011fc]
'00713b82    50                      push eax
'00713b83    57                      push edi
'00713b84    ff93ac000000            call dword ptr [ebx+000000ac]
'00713b8a    dbe2                    fnclex
'00713b8c    3bc6                    cmp eax, esi
'00713b8e    7d12                    jge 713ba2
'00713b90    68ac000000              push 000000ac
'00713b95    681cb94300              push 0043b91c
'00713b9a    57                      push edi
'00713b9b    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00713b9c    ff1580104000            call dword ptr [00401080]
'00713ba2    8d4de4                  lea ecx, dword ptr [ebp-1c]

' *** Reference to "__vbaFreeStr"
'00713ba5    ff150c134000            call dword ptr [0040130c]
'#Cleanup(var_40)
'00713bab    8d4dbc                  lea ecx, dword ptr [ebp-44]
'00713bae    51                      push ecx
'00713baf    8d55c4                  lea edx, dword ptr [ebp-3c]
'00713bb2    52                      push edx
'00713bb3    8d45cc                  lea eax, dword ptr [ebp-34]
'00713bb6    50                      push eax
'00713bb7    6a03                    push 03

' *** Reference to "__vbaFreeObjList"
'00713bb9    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_43, var_9, var_58)
'00713bbf    8d8d7cffffff            lea ecx, dword ptr [ebp+ffffff7c]
'00713bc5    51                      push ecx
'00713bc6    8d558c                  lea edx, dword ptr [ebp-74]
'00713bc9    52                      push edx
'00713bca    8d459c                  lea eax, dword ptr [ebp-64]
'00713bcd    50                      push eax
'00713bce    8d8d4cffffff            lea ecx, dword ptr [ebp+ffffff4c]
'00713bd4    51                      push ecx
'00713bd5    8d55ac                  lea edx, dword ptr [ebp-54]
'00713bd8    52                      push edx
'00713bd9    6a05                    push 05

' *** Reference to "__vbaFreeVarList"
'00713bdb    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_50, var_89, var_51, var_52, var_59)
'00713be1    8b8514ffffff            mov eax, dword ptr [ebp+ffffff14]
'00713be7    8b08                    mov ecx, dword ptr [eax]
'00713be9    83c428                  add esp, 28
'00713bec    8d55cc                  lea edx, dword ptr [ebp-34]
'00713bef    52                      push edx
'00713bf0    50                      push eax
'00713bf1    ff91b4000000            call dword ptr [ecx+000000b4]
'00713bf7    dbe2                    fnclex
'00713bf9    3bc6                    cmp eax, esi
'00713bfb    7d18                    jge 713c15

If (var_488 < 0) Then
'00713bfd    8b8d14ffffff            mov ecx, dword ptr [ebp+ffffff14]
'00713c03    68b4000000              push 000000b4
'00713c08    6830314300              push 00433130
'00713c0d    51                      push ecx
'00713c0e    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00713c0f    ff1580104000            call dword ptr [00401080]
    
End If
'00713c15    8b45cc                  mov eax, dword ptr [ebp-34]
'00713c18    8b38                    mov edi, dword ptr [eax]
'00713c1a    8d5dc8                  lea ebx, dword ptr [ebp-38]
'00713c1d    53                      push ebx
'00713c1e    83ec10                  sub esp, 10
'00713c21    8bdc                    mov ebx, esp
'00713c23    ba08000000              mov edx, 00000008
'00713c28    8913                    mov dword ptr [ebx], edx
'00713c2a    8b9570ffffff            mov edx, dword ptr [ebp+ffffff70]
'00713c30    895304                  mov dword ptr [ebx+04], edx
'00713c33    b9b4b84300              mov ecx, 0043b8b4
'00713c38    894b08                  mov dword ptr [ebx+08], ecx
'00713c3b    8b8d78ffffff            mov ecx, dword ptr [ebp+ffffff78]
'00713c41    50                      push eax
'00713c42    898530ffffff            mov dword ptr [ebp+ffffff30], eax
'00713c48    894b0c                  mov dword ptr [ebx+0c], ecx
'00713c4b    ff5730                  call dword ptr [edi+30]
'00713c4e    dbe2                    fnclex
'00713c50    3bc6                    cmp eax, esi
'00713c52    7d15                    jge 713c69

If (var_43 < 0) Then
'00713c54    8b9530ffffff            mov edx, dword ptr [ebp+ffffff30]
'00713c5a    6a30                    push 30
'00713c5c    68d8304300              push 004330d8
'00713c61    52                      push edx
'00713c62    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00713c63    ff1580104000            call dword ptr [00401080]
    
End If
'00713c69    8b4508                  mov eax, dword ptr [ebp+08]
'00713c6c    8b08                    mov ecx, dword ptr [eax]
'00713c6e    50                      push eax
'00713c6f    ff910c030000            call dword ptr [ecx+0000030c]
'00713c75    50                      push eax
'00713c76    8d55c4                  lea edx, dword ptr [ebp-3c]
'00713c79    52                      push edx

' *** Reference to "__vbaObjSet"
'00713c7a    ff15b4104000            call dword ptr [004010b4]
Set var_9 = Me
'00713c80    8bf8                    mov edi, eax
'00713c82    b802000000              mov eax, 00000002
'00713c87    89458c                  mov dword ptr [ebp-74], eax
'00713c8a    89459c                  mov dword ptr [ebp-64], eax
'00713c8d    8b45c8                  mov eax, dword ptr [ebp-38]
'00713c90    8945b4                  mov dword ptr [ebp-4c], eax
'00713c93    8d458c                  lea eax, dword ptr [ebp-74]
'00713c96    50                      push eax
'00713c97    8d4d9c                  lea ecx, dword ptr [ebp-64]
'00713c9a    51                      push ecx
'00713c9b    8d55ac                  lea edx, dword ptr [ebp-54]
'00713c9e    52                      push edx
'00713c9f    8d857cffffff            lea eax, dword ptr [ebp+ffffff7c]
'00713ca5    50                      push eax
'00713ca6    897594                  mov dword ptr [ebp-6c], esi
'00713ca9    c745a401000000          mov dword ptr [ebp-5c], 00000001
'00713cb0    8975c8                  mov dword ptr [ebp-38], esi
'00713cb3    c745ac09000000          mov dword ptr [ebp-54], 00000009

' *** Reference to "rtcImmediateIf"
'00713cba    ff154c124000            call dword ptr [0040124c]
'00713cc0    8b1f                    mov ebx, dword ptr [edi]
'00713cc2    8d8d7cffffff            lea ecx, dword ptr [ebp+ffffff7c]
'00713cc8    51                      push ecx

' *** Reference to "__vbaI2Var"
'00713cc9    ff150c124000            call dword ptr [0040120c]
'00713ccf    50                      push eax
'00713cd0    57                      push edi
'00713cd1    ff93e4000000            call dword ptr [ebx+000000e4]
'00713cd7    dbe2                    fnclex
'00713cd9    3bc6                    cmp eax, esi
'00713cdb    7d12                    jge 713cef

If (CInt(IIf(0, var_83, 0)) < 0) Then
'00713cdd    68e4000000              push 000000e4
'00713ce2    68dce24300              push 0043e2dc
'00713ce7    57                      push edi
'00713ce8    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00713ce9    ff1580104000            call dword ptr [00401080]
    
End If
'00713cef    8d55c4                  lea edx, dword ptr [ebp-3c]
'00713cf2    52                      push edx
'00713cf3    8d45cc                  lea eax, dword ptr [ebp-34]
'00713cf6    50                      push eax
'00713cf7    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'00713cf9    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_43, var_9)
'00713cff    8d8d7cffffff            lea ecx, dword ptr [ebp+ffffff7c]
'00713d05    51                      push ecx
'00713d06    8d558c                  lea edx, dword ptr [ebp-74]
'00713d09    52                      push edx
'00713d0a    8d459c                  lea eax, dword ptr [ebp-64]
'00713d0d    50                      push eax
'00713d0e    8d4dac                  lea ecx, dword ptr [ebp-54]
'00713d11    51                      push ecx
'00713d12    6a04                    push 04

' *** Reference to "__vbaFreeVarList"
'00713d14    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_50, var_51, var_52, var_59)
'00713d1a    8b8514ffffff            mov eax, dword ptr [ebp+ffffff14]
'00713d20    8b10                    mov edx, dword ptr [eax]
'00713d22    83c420                  add esp, 20
'00713d25    8d4dcc                  lea ecx, dword ptr [ebp-34]
'00713d28    51                      push ecx
'00713d29    50                      push eax
'00713d2a    ff92b4000000            call dword ptr [edx+000000b4]
'00713d30    dbe2                    fnclex
'00713d32    3bc6                    cmp eax, esi
'00713d34    7d18                    jge 713d4e

If (var_488 < 0) Then
'00713d36    8b9514ffffff            mov edx, dword ptr [ebp+ffffff14]
'00713d3c    68b4000000              push 000000b4
'00713d41    6830314300              push 00433130
'00713d46    52                      push edx
'00713d47    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00713d48    ff1580104000            call dword ptr [00401080]
    
End If
'00713d4e    8b45cc                  mov eax, dword ptr [ebp-34]
'00713d51    8b38                    mov edi, dword ptr [eax]
'00713d53    8d5dc8                  lea ebx, dword ptr [ebp-38]
'00713d56    53                      push ebx
'00713d57    83ec10                  sub esp, 10
'00713d5a    8bdc                    mov ebx, esp
'00713d5c    ba08000000              mov edx, 00000008
'00713d61    8913                    mov dword ptr [ebx], edx
'00713d63    8b9570ffffff            mov edx, dword ptr [ebp+ffffff70]
'00713d69    895304                  mov dword ptr [ebx+04], edx
'00713d6c    b9c4b84300              mov ecx, 0043b8c4
'00713d71    894b08                  mov dword ptr [ebx+08], ecx
'00713d74    8b8d78ffffff            mov ecx, dword ptr [ebp+ffffff78]
'00713d7a    50                      push eax
'00713d7b    898530ffffff            mov dword ptr [ebp+ffffff30], eax
'00713d81    894b0c                  mov dword ptr [ebx+0c], ecx
'00713d84    ff5730                  call dword ptr [edi+30]
'00713d87    dbe2                    fnclex
'00713d89    3bc6                    cmp eax, esi
'00713d8b    7d15                    jge 713da2

If (var_43 < 0) Then
'00713d8d    8b9530ffffff            mov edx, dword ptr [ebp+ffffff30]
'00713d93    6a30                    push 30
'00713d95    68d8304300              push 004330d8
'00713d9a    52                      push edx
'00713d9b    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00713d9c    ff1580104000            call dword ptr [00401080]
    
End If
'00713da2    8b4508                  mov eax, dword ptr [ebp+08]
'00713da5    8b08                    mov ecx, dword ptr [eax]
'00713da7    50                      push eax
'00713da8    ff9144030000            call dword ptr [ecx+00000344]
'00713dae    50                      push eax
'00713daf    8d55c4                  lea edx, dword ptr [ebp-3c]
'00713db2    52                      push edx

' *** Reference to "__vbaObjSet"
'00713db3    ff15b4104000            call dword ptr [004010b4]
Set var_9 = Me
'00713db9    8bf8                    mov edi, eax
'00713dbb    b802000000              mov eax, 00000002
'00713dc0    89458c                  mov dword ptr [ebp-74], eax
'00713dc3    89459c                  mov dword ptr [ebp-64], eax
'00713dc6    8b45c8                  mov eax, dword ptr [ebp-38]
'00713dc9    8945b4                  mov dword ptr [ebp-4c], eax
'00713dcc    8d458c                  lea eax, dword ptr [ebp-74]
'00713dcf    50                      push eax
'00713dd0    8d4d9c                  lea ecx, dword ptr [ebp-64]
'00713dd3    51                      push ecx
'00713dd4    8d55ac                  lea edx, dword ptr [ebp-54]
'00713dd7    52                      push edx
'00713dd8    8d857cffffff            lea eax, dword ptr [ebp+ffffff7c]
'00713dde    50                      push eax
'00713ddf    897594                  mov dword ptr [ebp-6c], esi
'00713de2    c745a401000000          mov dword ptr [ebp-5c], 00000001
'00713de9    8975c8                  mov dword ptr [ebp-38], esi
'00713dec    c745ac09000000          mov dword ptr [ebp-54], 00000009

' *** Reference to "rtcImmediateIf"
'00713df3    ff154c124000            call dword ptr [0040124c]
'00713df9    8b1f                    mov ebx, dword ptr [edi]
'00713dfb    8d8d7cffffff            lea ecx, dword ptr [ebp+ffffff7c]
'00713e01    51                      push ecx

' *** Reference to "__vbaI2Var"
'00713e02    ff150c124000            call dword ptr [0040120c]
'00713e08    50                      push eax
'00713e09    57                      push edi
'00713e0a    ff93e4000000            call dword ptr [ebx+000000e4]
'00713e10    dbe2                    fnclex
'00713e12    3bc6                    cmp eax, esi
'00713e14    7d12                    jge 713e28

If (CInt(IIf(0, var_83, 0)) < 0) Then
'00713e16    68e4000000              push 000000e4
'00713e1b    68dce24300              push 0043e2dc
'00713e20    57                      push edi
'00713e21    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00713e22    ff1580104000            call dword ptr [00401080]
    
End If
'00713e28    8d55c4                  lea edx, dword ptr [ebp-3c]
'00713e2b    52                      push edx
'00713e2c    8d45cc                  lea eax, dword ptr [ebp-34]
'00713e2f    50                      push eax
'00713e30    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'00713e32    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_43, var_9)
'00713e38    8d8d7cffffff            lea ecx, dword ptr [ebp+ffffff7c]
'00713e3e    51                      push ecx
'00713e3f    8d558c                  lea edx, dword ptr [ebp-74]
'00713e42    52                      push edx
'00713e43    8d459c                  lea eax, dword ptr [ebp-64]
'00713e46    50                      push eax
'00713e47    8d4dac                  lea ecx, dword ptr [ebp-54]
'00713e4a    51                      push ecx
'00713e4b    6a04                    push 04

' *** Reference to "__vbaFreeVarList"
'00713e4d    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_50, var_51, var_52, var_59)
'00713e53    8b8514ffffff            mov eax, dword ptr [ebp+ffffff14]
'00713e59    8b10                    mov edx, dword ptr [eax]
'00713e5b    83c420                  add esp, 20
'00713e5e    8d4dcc                  lea ecx, dword ptr [ebp-34]
'00713e61    51                      push ecx
'00713e62    50                      push eax
'00713e63    ff92b4000000            call dword ptr [edx+000000b4]
'00713e69    dbe2                    fnclex
'00713e6b    3bc6                    cmp eax, esi
'00713e6d    7d18                    jge 713e87

If (var_488 < 0) Then
'00713e6f    8b9514ffffff            mov edx, dword ptr [ebp+ffffff14]
'00713e75    68b4000000              push 000000b4
'00713e7a    6830314300              push 00433130
'00713e7f    52                      push edx
'00713e80    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00713e81    ff1580104000            call dword ptr [00401080]
    
End If
'00713e87    8b45cc                  mov eax, dword ptr [ebp-34]
'00713e8a    8b38                    mov edi, dword ptr [eax]
'00713e8c    8d5dc8                  lea ebx, dword ptr [ebp-38]
'00713e8f    53                      push ebx
'00713e90    83ec10                  sub esp, 10
'00713e93    8bdc                    mov ebx, esp
'00713e95    ba08000000              mov edx, 00000008
'00713e9a    8913                    mov dword ptr [ebx], edx
'00713e9c    8b9570ffffff            mov edx, dword ptr [ebp+ffffff70]
'00713ea2    895304                  mov dword ptr [ebx+04], edx
'00713ea5    b96ca44300              mov ecx, 0043a46c
'00713eaa    894b08                  mov dword ptr [ebx+08], ecx
'00713ead    8b8d78ffffff            mov ecx, dword ptr [ebp+ffffff78]
'00713eb3    50                      push eax
'00713eb4    898530ffffff            mov dword ptr [ebp+ffffff30], eax
'00713eba    894b0c                  mov dword ptr [ebx+0c], ecx
'00713ebd    ff5730                  call dword ptr [edi+30]
'00713ec0    dbe2                    fnclex
'00713ec2    3bc6                    cmp eax, esi
'00713ec4    7d15                    jge 713edb

If (var_43 < 0) Then
'00713ec6    8b9530ffffff            mov edx, dword ptr [ebp+ffffff30]
'00713ecc    6a30                    push 30
'00713ece    68d8304300              push 004330d8
'00713ed3    52                      push edx
'00713ed4    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00713ed5    ff1580104000            call dword ptr [00401080]
    
End If
'00713edb    8b4508                  mov eax, dword ptr [ebp+08]
'00713ede    8b08                    mov ecx, dword ptr [eax]
'00713ee0    50                      push eax
'00713ee1    ff9110030000            call dword ptr [ecx+00000310]
'00713ee7    50                      push eax
'00713ee8    8d55c4                  lea edx, dword ptr [ebp-3c]
'00713eeb    52                      push edx

' *** Reference to "__vbaObjSet"
'00713eec    ff15b4104000            call dword ptr [004010b4]
Set var_9 = Me
'00713ef2    8bf8                    mov edi, eax
'00713ef4    b802000000              mov eax, 00000002
'00713ef9    89458c                  mov dword ptr [ebp-74], eax
'00713efc    89459c                  mov dword ptr [ebp-64], eax
'00713eff    8b45c8                  mov eax, dword ptr [ebp-38]
'00713f02    8945b4                  mov dword ptr [ebp-4c], eax
'00713f05    8d458c                  lea eax, dword ptr [ebp-74]
'00713f08    50                      push eax
'00713f09    8d4d9c                  lea ecx, dword ptr [ebp-64]
'00713f0c    51                      push ecx
'00713f0d    8d55ac                  lea edx, dword ptr [ebp-54]
'00713f10    52                      push edx
'00713f11    8d857cffffff            lea eax, dword ptr [ebp+ffffff7c]
'00713f17    50                      push eax
'00713f18    897594                  mov dword ptr [ebp-6c], esi
'00713f1b    c745a401000000          mov dword ptr [ebp-5c], 00000001
'00713f22    8975c8                  mov dword ptr [ebp-38], esi
'00713f25    c745ac09000000          mov dword ptr [ebp-54], 00000009

' *** Reference to "rtcImmediateIf"
'00713f2c    ff154c124000            call dword ptr [0040124c]
'00713f32    8b1f                    mov ebx, dword ptr [edi]
'00713f34    8d8d7cffffff            lea ecx, dword ptr [ebp+ffffff7c]
'00713f3a    51                      push ecx

' *** Reference to "__vbaI2Var"
'00713f3b    ff150c124000            call dword ptr [0040120c]
'00713f41    50                      push eax
'00713f42    57                      push edi
'00713f43    ff93e4000000            call dword ptr [ebx+000000e4]
'00713f49    dbe2                    fnclex
'00713f4b    3bc6                    cmp eax, esi
'00713f4d    7d12                    jge 713f61

If (CInt(IIf(0, var_83, 0)) < 0) Then
'00713f4f    68e4000000              push 000000e4
'00713f54    68dce24300              push 0043e2dc
'00713f59    57                      push edi
'00713f5a    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00713f5b    ff1580104000            call dword ptr [00401080]
    
End If
'00713f61    8d55c4                  lea edx, dword ptr [ebp-3c]
'00713f64    52                      push edx
'00713f65    8d45cc                  lea eax, dword ptr [ebp-34]
'00713f68    50                      push eax
'00713f69    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'00713f6b    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_43, var_9)
'00713f71    8d8d7cffffff            lea ecx, dword ptr [ebp+ffffff7c]
'00713f77    51                      push ecx
'00713f78    8d558c                  lea edx, dword ptr [ebp-74]
'00713f7b    52                      push edx
'00713f7c    8d459c                  lea eax, dword ptr [ebp-64]
'00713f7f    50                      push eax
'00713f80    8d4dac                  lea ecx, dword ptr [ebp-54]
'00713f83    51                      push ecx
'00713f84    6a04                    push 04

' *** Reference to "__vbaFreeVarList"
'00713f86    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_50, var_51, var_52, var_59)
'00713f8c    8b8514ffffff            mov eax, dword ptr [ebp+ffffff14]
'00713f92    8b10                    mov edx, dword ptr [eax]
'00713f94    83c420                  add esp, 20
'00713f97    8d4dcc                  lea ecx, dword ptr [ebp-34]
'00713f9a    51                      push ecx
'00713f9b    50                      push eax
'00713f9c    ff92b4000000            call dword ptr [edx+000000b4]
'00713fa2    dbe2                    fnclex
'00713fa4    3bc6                    cmp eax, esi
'00713fa6    7d18                    jge 713fc0

If (var_488 < 0) Then
'00713fa8    8b9514ffffff            mov edx, dword ptr [ebp+ffffff14]
'00713fae    68b4000000              push 000000b4
'00713fb3    6830314300              push 00433130
'00713fb8    52                      push edx
'00713fb9    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00713fba    ff1580104000            call dword ptr [00401080]
    
End If
'00713fc0    8b45cc                  mov eax, dword ptr [ebp-34]
'00713fc3    8b38                    mov edi, dword ptr [eax]
'00713fc5    8d5dc8                  lea ebx, dword ptr [ebp-38]
'00713fc8    53                      push ebx
'00713fc9    83ec10                  sub esp, 10
'00713fcc    8bdc                    mov ebx, esp
'00713fce    ba08000000              mov edx, 00000008
'00713fd3    8913                    mov dword ptr [ebx], edx
'00713fd5    8b9570ffffff            mov edx, dword ptr [ebp+ffffff70]
'00713fdb    895304                  mov dword ptr [ebx+04], edx
'00713fde    b9d8b84300              mov ecx, 0043b8d8
'00713fe3    894b08                  mov dword ptr [ebx+08], ecx
'00713fe6    8b8d78ffffff            mov ecx, dword ptr [ebp+ffffff78]
'00713fec    50                      push eax
'00713fed    898530ffffff            mov dword ptr [ebp+ffffff30], eax
'00713ff3    894b0c                  mov dword ptr [ebx+0c], ecx
'00713ff6    ff5730                  call dword ptr [edi+30]
'00713ff9    dbe2                    fnclex
'00713ffb    3bc6                    cmp eax, esi
'00713ffd    7d15                    jge 714014

If (var_43 < 0) Then
'00713fff    8b9530ffffff            mov edx, dword ptr [ebp+ffffff30]
'00714005    6a30                    push 30
'00714007    68d8304300              push 004330d8
'0071400c    52                      push edx
'0071400d    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0071400e    ff1580104000            call dword ptr [00401080]
    
End If
'00714014    8b4508                  mov eax, dword ptr [ebp+08]
'00714017    8b08                    mov ecx, dword ptr [eax]
'00714019    50                      push eax
'0071401a    ff9114030000            call dword ptr [ecx+00000314]
'00714020    50                      push eax
'00714021    8d55c4                  lea edx, dword ptr [ebp-3c]
'00714024    52                      push edx

' *** Reference to "__vbaObjSet"
'00714025    ff15b4104000            call dword ptr [004010b4]
Set var_9 = Me
'0071402b    8bf8                    mov edi, eax
'0071402d    b802000000              mov eax, 00000002
'00714032    89458c                  mov dword ptr [ebp-74], eax
'00714035    89459c                  mov dword ptr [ebp-64], eax
'00714038    8b45c8                  mov eax, dword ptr [ebp-38]
'0071403b    8945b4                  mov dword ptr [ebp-4c], eax
'0071403e    8d458c                  lea eax, dword ptr [ebp-74]
'00714041    50                      push eax
'00714042    8d4d9c                  lea ecx, dword ptr [ebp-64]
'00714045    51                      push ecx
'00714046    8d55ac                  lea edx, dword ptr [ebp-54]
'00714049    52                      push edx
'0071404a    8d857cffffff            lea eax, dword ptr [ebp+ffffff7c]
'00714050    50                      push eax
'00714051    897594                  mov dword ptr [ebp-6c], esi
'00714054    c745a401000000          mov dword ptr [ebp-5c], 00000001
'0071405b    8975c8                  mov dword ptr [ebp-38], esi
'0071405e    c745ac09000000          mov dword ptr [ebp-54], 00000009

' *** Reference to "rtcImmediateIf"
'00714065    ff154c124000            call dword ptr [0040124c]
'0071406b    8b1f                    mov ebx, dword ptr [edi]
'0071406d    8d8d7cffffff            lea ecx, dword ptr [ebp+ffffff7c]
'00714073    51                      push ecx

' *** Reference to "__vbaI2Var"
'00714074    ff150c124000            call dword ptr [0040120c]
'0071407a    50                      push eax
'0071407b    57                      push edi
'0071407c    ff93e4000000            call dword ptr [ebx+000000e4]
'00714082    dbe2                    fnclex
'00714084    3bc6                    cmp eax, esi
'00714086    7d12                    jge 71409a

If (CInt(IIf(0, var_83, 0)) < 0) Then
'00714088    68e4000000              push 000000e4
'0071408d    68dce24300              push 0043e2dc
'00714092    57                      push edi
'00714093    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00714094    ff1580104000            call dword ptr [00401080]
    
End If
'0071409a    8d55c4                  lea edx, dword ptr [ebp-3c]
'0071409d    52                      push edx
'0071409e    8d45cc                  lea eax, dword ptr [ebp-34]
'007140a1    50                      push eax
'007140a2    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'007140a4    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_43, var_9)
'007140aa    8d8d7cffffff            lea ecx, dword ptr [ebp+ffffff7c]
'007140b0    51                      push ecx
'007140b1    8d558c                  lea edx, dword ptr [ebp-74]
'007140b4    52                      push edx
'007140b5    8d459c                  lea eax, dword ptr [ebp-64]
'007140b8    50                      push eax
'007140b9    8d4dac                  lea ecx, dword ptr [ebp-54]
'007140bc    51                      push ecx
'007140bd    6a04                    push 04

' *** Reference to "__vbaFreeVarList"
'007140bf    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_50, var_51, var_52, var_59)
'007140c5    8b8514ffffff            mov eax, dword ptr [ebp+ffffff14]
'007140cb    8b10                    mov edx, dword ptr [eax]
'007140cd    83c420                  add esp, 20
'007140d0    8d4dcc                  lea ecx, dword ptr [ebp-34]
'007140d3    51                      push ecx
'007140d4    50                      push eax
'007140d5    ff92b4000000            call dword ptr [edx+000000b4]
'007140db    dbe2                    fnclex
'007140dd    3bc6                    cmp eax, esi
'007140df    7d18                    jge 7140f9

If (var_488 < 0) Then
'007140e1    8b9514ffffff            mov edx, dword ptr [ebp+ffffff14]
'007140e7    68b4000000              push 000000b4
'007140ec    6830314300              push 00433130
'007140f1    52                      push edx
'007140f2    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'007140f3    ff1580104000            call dword ptr [00401080]
    
End If
'007140f9    8b45cc                  mov eax, dword ptr [ebp-34]
'007140fc    8b38                    mov edi, dword ptr [eax]
'007140fe    8d5dc8                  lea ebx, dword ptr [ebp-38]
'00714101    53                      push ebx
'00714102    83ec10                  sub esp, 10
'00714105    8bdc                    mov ebx, esp
'00714107    ba08000000              mov edx, 00000008
'0071410c    8913                    mov dword ptr [ebx], edx
'0071410e    8b9570ffffff            mov edx, dword ptr [ebp+ffffff70]
'00714114    895304                  mov dword ptr [ebx+04], edx
'00714117    b9ecb84300              mov ecx, 0043b8ec
'0071411c    894b08                  mov dword ptr [ebx+08], ecx
'0071411f    8b8d78ffffff            mov ecx, dword ptr [ebp+ffffff78]
'00714125    50                      push eax
'00714126    898530ffffff            mov dword ptr [ebp+ffffff30], eax
'0071412c    894b0c                  mov dword ptr [ebx+0c], ecx
'0071412f    ff5730                  call dword ptr [edi+30]
'00714132    dbe2                    fnclex
'00714134    3bc6                    cmp eax, esi
'00714136    7d15                    jge 71414d

If (var_43 < 0) Then
'00714138    8b9530ffffff            mov edx, dword ptr [ebp+ffffff30]
'0071413e    6a30                    push 30
'00714140    68d8304300              push 004330d8
'00714145    52                      push edx
'00714146    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00714147    ff1580104000            call dword ptr [00401080]
    
End If
'0071414d    8b4508                  mov eax, dword ptr [ebp+08]
'00714150    8b08                    mov ecx, dword ptr [eax]
'00714152    50                      push eax
'00714153    ff9148030000            call dword ptr [ecx+00000348]
'00714159    50                      push eax
'0071415a    8d55c4                  lea edx, dword ptr [ebp-3c]
'0071415d    52                      push edx

' *** Reference to "__vbaObjSet"
'0071415e    ff15b4104000            call dword ptr [004010b4]
Set var_9 = Me
'00714164    8bf8                    mov edi, eax
'00714166    b802000000              mov eax, 00000002
'0071416b    89458c                  mov dword ptr [ebp-74], eax
'0071416e    89459c                  mov dword ptr [ebp-64], eax
'00714171    8b45c8                  mov eax, dword ptr [ebp-38]
'00714174    8945b4                  mov dword ptr [ebp-4c], eax
'00714177    8d458c                  lea eax, dword ptr [ebp-74]
'0071417a    50                      push eax
'0071417b    8d4d9c                  lea ecx, dword ptr [ebp-64]
'0071417e    51                      push ecx
'0071417f    8d55ac                  lea edx, dword ptr [ebp-54]
'00714182    52                      push edx
'00714183    8d857cffffff            lea eax, dword ptr [ebp+ffffff7c]
'00714189    50                      push eax
'0071418a    897594                  mov dword ptr [ebp-6c], esi
'0071418d    c745a401000000          mov dword ptr [ebp-5c], 00000001
'00714194    8975c8                  mov dword ptr [ebp-38], esi
'00714197    c745ac09000000          mov dword ptr [ebp-54], 00000009

' *** Reference to "rtcImmediateIf"
'0071419e    ff154c124000            call dword ptr [0040124c]
'007141a4    8b1f                    mov ebx, dword ptr [edi]
'007141a6    8d8d7cffffff            lea ecx, dword ptr [ebp+ffffff7c]
'007141ac    51                      push ecx

' *** Reference to "__vbaI2Var"
'007141ad    ff150c124000            call dword ptr [0040120c]
'007141b3    50                      push eax
'007141b4    57                      push edi
'007141b5    ff93e4000000            call dword ptr [ebx+000000e4]
'007141bb    dbe2                    fnclex
'007141bd    3bc6                    cmp eax, esi
'007141bf    7d12                    jge 7141d3

If (CInt(IIf(0, var_83, 0)) < 0) Then
'007141c1    68e4000000              push 000000e4
'007141c6    68dce24300              push 0043e2dc
'007141cb    57                      push edi
'007141cc    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'007141cd    ff1580104000            call dword ptr [00401080]
    
End If
'007141d3    8d55c4                  lea edx, dword ptr [ebp-3c]
'007141d6    52                      push edx
'007141d7    8d45cc                  lea eax, dword ptr [ebp-34]
'007141da    50                      push eax
'007141db    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'007141dd    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_43, var_9)
'007141e3    8d8d7cffffff            lea ecx, dword ptr [ebp+ffffff7c]
'007141e9    51                      push ecx
'007141ea    8d558c                  lea edx, dword ptr [ebp-74]
'007141ed    52                      push edx
'007141ee    8d459c                  lea eax, dword ptr [ebp-64]
'007141f1    50                      push eax
'007141f2    8d4dac                  lea ecx, dword ptr [ebp-54]
'007141f5    51                      push ecx
'007141f6    6a04                    push 04

' *** Reference to "__vbaFreeVarList"
'007141f8    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_50, var_51, var_52, var_59)
'007141fe    8b8514ffffff            mov eax, dword ptr [ebp+ffffff14]
'00714204    8b10                    mov edx, dword ptr [eax]
'00714206    83c420                  add esp, 20
'00714209    8d4dcc                  lea ecx, dword ptr [ebp-34]
'0071420c    51                      push ecx
'0071420d    50                      push eax
'0071420e    ff92b4000000            call dword ptr [edx+000000b4]
'00714214    dbe2                    fnclex
'00714216    3bc6                    cmp eax, esi
'00714218    7d18                    jge 714232

If (var_488 < 0) Then
'0071421a    8b9514ffffff            mov edx, dword ptr [ebp+ffffff14]
'00714220    68b4000000              push 000000b4
'00714225    6830314300              push 00433130
'0071422a    52                      push edx
'0071422b    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0071422c    ff1580104000            call dword ptr [00401080]
    
End If
'00714232    8b45cc                  mov eax, dword ptr [ebp-34]
'00714235    8b38                    mov edi, dword ptr [eax]
'00714237    8d5dc8                  lea ebx, dword ptr [ebp-38]
'0071423a    53                      push ebx
'0071423b    83ec10                  sub esp, 10
'0071423e    8bdc                    mov ebx, esp
'00714240    ba08000000              mov edx, 00000008
'00714245    8913                    mov dword ptr [ebx], edx
'00714247    8b9570ffffff            mov edx, dword ptr [ebp+ffffff70]
'0071424d    895304                  mov dword ptr [ebx+04], edx
'00714250    b900b94300              mov ecx, 0043b900
'00714255    894b08                  mov dword ptr [ebx+08], ecx
'00714258    8b8d78ffffff            mov ecx, dword ptr [ebp+ffffff78]
'0071425e    50                      push eax
'0071425f    898530ffffff            mov dword ptr [ebp+ffffff30], eax
'00714265    894b0c                  mov dword ptr [ebx+0c], ecx
'00714268    ff5730                  call dword ptr [edi+30]
'0071426b    dbe2                    fnclex
'0071426d    3bc6                    cmp eax, esi
'0071426f    7d15                    jge 714286

If (var_43 < 0) Then
'00714271    8b9530ffffff            mov edx, dword ptr [ebp+ffffff30]
'00714277    6a30                    push 30
'00714279    68d8304300              push 004330d8
'0071427e    52                      push edx
'0071427f    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00714280    ff1580104000            call dword ptr [00401080]
    
End If
'00714286    8b4508                  mov eax, dword ptr [ebp+08]
'00714289    8b08                    mov ecx, dword ptr [eax]
'0071428b    50                      push eax
'0071428c    ff913c030000            call dword ptr [ecx+0000033c]
'00714292    50                      push eax
'00714293    8d55c4                  lea edx, dword ptr [ebp-3c]
'00714296    52                      push edx

' *** Reference to "__vbaObjSet"
'00714297    ff15b4104000            call dword ptr [004010b4]
Set var_9 = Me
'0071429d    8bf8                    mov edi, eax
'0071429f    b802000000              mov eax, 00000002
'007142a4    89458c                  mov dword ptr [ebp-74], eax
'007142a7    89459c                  mov dword ptr [ebp-64], eax
'007142aa    8b45c8                  mov eax, dword ptr [ebp-38]
'007142ad    8945b4                  mov dword ptr [ebp-4c], eax
'007142b0    8d458c                  lea eax, dword ptr [ebp-74]
'007142b3    50                      push eax
'007142b4    8d4d9c                  lea ecx, dword ptr [ebp-64]
'007142b7    51                      push ecx
'007142b8    8d55ac                  lea edx, dword ptr [ebp-54]
'007142bb    52                      push edx
'007142bc    8d857cffffff            lea eax, dword ptr [ebp+ffffff7c]
'007142c2    50                      push eax
'007142c3    897594                  mov dword ptr [ebp-6c], esi
'007142c6    c745a401000000          mov dword ptr [ebp-5c], 00000001
'007142cd    8975c8                  mov dword ptr [ebp-38], esi
'007142d0    c745ac09000000          mov dword ptr [ebp-54], 00000009

' *** Reference to "rtcImmediateIf"
'007142d7    ff154c124000            call dword ptr [0040124c]
'007142dd    8b1f                    mov ebx, dword ptr [edi]
'007142df    8d8d7cffffff            lea ecx, dword ptr [ebp+ffffff7c]
'007142e5    51                      push ecx

' *** Reference to "__vbaI2Var"
'007142e6    ff150c124000            call dword ptr [0040120c]
'007142ec    50                      push eax
'007142ed    57                      push edi
'007142ee    ff93e4000000            call dword ptr [ebx+000000e4]
'007142f4    dbe2                    fnclex
'007142f6    3bc6                    cmp eax, esi
'007142f8    7d12                    jge 71430c

If (CInt(IIf(0, var_83, 0)) < 0) Then
'007142fa    68e4000000              push 000000e4
'007142ff    68dce24300              push 0043e2dc
'00714304    57                      push edi
'00714305    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00714306    ff1580104000            call dword ptr [00401080]
    
End If
'0071430c    8d55c4                  lea edx, dword ptr [ebp-3c]
'0071430f    52                      push edx
'00714310    8d45cc                  lea eax, dword ptr [ebp-34]
'00714313    50                      push eax
'00714314    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'00714316    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_43, var_9)
'0071431c    8d8d7cffffff            lea ecx, dword ptr [ebp+ffffff7c]
'00714322    51                      push ecx
'00714323    8d558c                  lea edx, dword ptr [ebp-74]
'00714326    52                      push edx
'00714327    8d459c                  lea eax, dword ptr [ebp-64]
'0071432a    50                      push eax
'0071432b    8d4dac                  lea ecx, dword ptr [ebp-54]
'0071432e    51                      push ecx
'0071432f    6a04                    push 04

' *** Reference to "__vbaFreeVarList"
'00714331    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_50, var_51, var_52, var_59)
'00714337    8b8514ffffff            mov eax, dword ptr [ebp+ffffff14]
'0071433d    8b10                    mov edx, dword ptr [eax]
'0071433f    83c420                  add esp, 20
'00714342    8d4dcc                  lea ecx, dword ptr [ebp-34]
'00714345    51                      push ecx
'00714346    50                      push eax
'00714347    ff92b4000000            call dword ptr [edx+000000b4]
'0071434d    dbe2                    fnclex
'0071434f    3bc6                    cmp eax, esi
'00714351    7d18                    jge 71436b

If (var_488 < 0) Then
'00714353    8b9514ffffff            mov edx, dword ptr [ebp+ffffff14]
'00714359    68b4000000              push 000000b4
'0071435e    6830314300              push 00433130
'00714363    52                      push edx
'00714364    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00714365    ff1580104000            call dword ptr [00401080]
    
End If
'0071436b    8b45cc                  mov eax, dword ptr [ebp-34]
'0071436e    8b38                    mov edi, dword ptr [eax]
'00714370    8d5dc8                  lea ebx, dword ptr [ebp-38]
'00714373    53                      push ebx
'00714374    83ec10                  sub esp, 10
'00714377    8bdc                    mov ebx, esp
'00714379    ba08000000              mov edx, 00000008
'0071437e    8913                    mov dword ptr [ebx], edx
'00714380    8b9570ffffff            mov edx, dword ptr [ebp+ffffff70]
'00714386    895304                  mov dword ptr [ebx+04], edx
'00714389    b910b94300              mov ecx, 0043b910
'0071438e    894b08                  mov dword ptr [ebx+08], ecx
'00714391    8b8d78ffffff            mov ecx, dword ptr [ebp+ffffff78]
'00714397    50                      push eax
'00714398    898530ffffff            mov dword ptr [ebp+ffffff30], eax
'0071439e    894b0c                  mov dword ptr [ebx+0c], ecx
'007143a1    ff5730                  call dword ptr [edi+30]
'007143a4    dbe2                    fnclex
'007143a6    3bc6                    cmp eax, esi
'007143a8    7d15                    jge 7143bf

If (var_43 < 0) Then
'007143aa    8b9530ffffff            mov edx, dword ptr [ebp+ffffff30]
'007143b0    6a30                    push 30
'007143b2    68d8304300              push 004330d8
'007143b7    52                      push edx
'007143b8    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'007143b9    ff1580104000            call dword ptr [00401080]
    
End If
'007143bf    8b4508                  mov eax, dword ptr [ebp+08]
'007143c2    8b08                    mov ecx, dword ptr [eax]
'007143c4    50                      push eax
'007143c5    ff9140030000            call dword ptr [ecx+00000340]
'007143cb    50                      push eax
'007143cc    8d55c4                  lea edx, dword ptr [ebp-3c]
'007143cf    52                      push edx

' *** Reference to "__vbaObjSet"
'007143d0    ff15b4104000            call dword ptr [004010b4]
Set var_9 = Me
'007143d6    8bf8                    mov edi, eax
'007143d8    b802000000              mov eax, 00000002
'007143dd    89458c                  mov dword ptr [ebp-74], eax
'007143e0    89459c                  mov dword ptr [ebp-64], eax
'007143e3    8b45c8                  mov eax, dword ptr [ebp-38]
'007143e6    8945b4                  mov dword ptr [ebp-4c], eax
'007143e9    8d458c                  lea eax, dword ptr [ebp-74]
'007143ec    50                      push eax
'007143ed    8d4d9c                  lea ecx, dword ptr [ebp-64]
'007143f0    51                      push ecx
'007143f1    8d55ac                  lea edx, dword ptr [ebp-54]
'007143f4    52                      push edx
'007143f5    8d857cffffff            lea eax, dword ptr [ebp+ffffff7c]
'007143fb    50                      push eax
'007143fc    897594                  mov dword ptr [ebp-6c], esi
'007143ff    c745a401000000          mov dword ptr [ebp-5c], 00000001
'00714406    8975c8                  mov dword ptr [ebp-38], esi
'00714409    c745ac09000000          mov dword ptr [ebp-54], 00000009

' *** Reference to "rtcImmediateIf"
'00714410    ff154c124000            call dword ptr [0040124c]
'00714416    8b1f                    mov ebx, dword ptr [edi]
'00714418    8d8d7cffffff            lea ecx, dword ptr [ebp+ffffff7c]
'0071441e    51                      push ecx

' *** Reference to "__vbaI2Var"
'0071441f    ff150c124000            call dword ptr [0040120c]
'00714425    50                      push eax
'00714426    57                      push edi
'00714427    ff93e4000000            call dword ptr [ebx+000000e4]
'0071442d    dbe2                    fnclex
'0071442f    3bc6                    cmp eax, esi
'00714431    7d12                    jge 714445

If (CInt(IIf(0, var_83, 0)) < 0) Then
'00714433    68e4000000              push 000000e4
'00714438    68dce24300              push 0043e2dc
'0071443d    57                      push edi
'0071443e    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0071443f    ff1580104000            call dword ptr [00401080]
    
End If
'00714445    8d55c4                  lea edx, dword ptr [ebp-3c]
'00714448    52                      push edx
'00714449    8d45cc                  lea eax, dword ptr [ebp-34]
'0071444c    50                      push eax
'0071444d    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'0071444f    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_43, var_9)
'00714455    8d8d7cffffff            lea ecx, dword ptr [ebp+ffffff7c]
'0071445b    51                      push ecx
'0071445c    8d558c                  lea edx, dword ptr [ebp-74]
'0071445f    52                      push edx
'00714460    8d459c                  lea eax, dword ptr [ebp-64]
'00714463    50                      push eax
'00714464    8d4dac                  lea ecx, dword ptr [ebp-54]
'00714467    51                      push ecx
'00714468    6a04                    push 04

' *** Reference to "__vbaFreeVarList"
'0071446a    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_50, var_51, var_52, var_59)
'00714470    83c420                  add esp, 20

'ERROR: Two many next close:
End If
'00714473    56                      push esi
'00714474    8d9514ffffff            lea edx, dword ptr [ebp+ffffff14]
'0071447a    52                      push edx

' *** Reference to "__vbaObjSetAddref"
'0071447b    ff15c8104000            call dword ptr [004010c8]
Set var_488 = Nothing
'00714481    8b45e8                  mov eax, dword ptr [ebp-18]
'00714484    8b08                    mov ecx, dword ptr [eax]
'00714486    50                      push eax
'00714487    ff91c4000000            call dword ptr [ecx+000000c4]
'0071448d    dbe2                    fnclex
'0071448f    3bc6                    cmp eax, esi
'00714491    7d15                    jge 7144a8

If (var_41 < 0) Then
'00714493    8b55e8                  mov edx, dword ptr [ebp-18]
'00714496    68c4000000              push 000000c4
'0071449b    6830314300              push 00433130
'007144a0    52                      push edx
'007144a1    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'007144a2    ff1580104000            call dword ptr [00401080]
    
End If
'007144a8    681e457100              push 0071451e
'007144ad    eb5b                    jmp 71450a
'007144af    8d45d0                  lea eax, dword ptr [ebp-30]
'007144b2    50                      push eax
'007144b3    8d4dd4                  lea ecx, dword ptr [ebp-2c]
'007144b6    51                      push ecx
'007144b7    8d55d8                  lea edx, dword ptr [ebp-28]
'007144ba    52                      push edx
'007144bb    8d45dc                  lea eax, dword ptr [ebp-24]
'007144be    50                      push eax
'007144bf    8d4de0                  lea ecx, dword ptr [ebp-20]
'007144c2    51                      push ecx
'007144c3    8d55e4                  lea edx, dword ptr [ebp-1c]
'007144c6    52                      push edx
'007144c7    6a06                    push 06

' *** Reference to "__vbaFreeStrList"
'007144c9    ff155c124000            call dword ptr [0040125c]
'#Cleanup( , 0, 0, -4496, -4500, 0)
'007144cf    8d45bc                  lea eax, dword ptr [ebp-44]
'007144d2    50                      push eax
'007144d3    8d4dc0                  lea ecx, dword ptr [ebp-40]
'007144d6    51                      push ecx
'007144d7    8d55c4                  lea edx, dword ptr [ebp-3c]
'007144da    52                      push edx
'007144db    8d45c8                  lea eax, dword ptr [ebp-38]
'007144de    50                      push eax
'007144df    8d4dcc                  lea ecx, dword ptr [ebp-34]
'007144e2    51                      push ecx
'007144e3    6a05                    push 05

' *** Reference to "__vbaFreeObjList"
'007144e5    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_43, var_46, var_9, var_5, var_58)
'007144eb    8d957cffffff            lea edx, dword ptr [ebp+ffffff7c]
'007144f1    52                      push edx
'007144f2    8d458c                  lea eax, dword ptr [ebp-74]
'007144f5    50                      push eax
'007144f6    8d4d9c                  lea ecx, dword ptr [ebp-64]
'007144f9    51                      push ecx
'007144fa    8d55ac                  lea edx, dword ptr [ebp-54]
'007144fd    52                      push edx
'007144fe    6a04                    push 04

' *** Reference to "__vbaFreeVarList"
'00714500    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_50, var_51, var_52, var_59)
'00714506    83c448                  add esp, 48
'00714509    c3                      ret

' *** Reference to "__vbaFreeObj"
'0071450a    8b3508134000            mov esi, dword ptr [00401308]
'00714510    8d8d14ffffff            lea ecx, dword ptr [ebp+ffffff14]
'00714516    ffd6                    call esi
'#Cleanup(var_488)
'00714518    8d4de8                  lea ecx, dword ptr [ebp-18]
'0071451b    ffd6                    call esi
'#Cleanup(var_41)
'0071451d    c3                      ret
'0071451e    8b4508                  mov eax, dword ptr [ebp+08]
'00714521    8b08                    mov ecx, dword ptr [eax]
'00714523    50                      push eax
'00714524    ff5108                  call dword ptr [ecx+08]
'00714527    8b45fc                  mov eax, dword ptr [ebp-04]
'0071452a    8b4dec                  mov ecx, dword ptr [ebp-14]
'0071452d    5f                      pop edi
'0071452e    5e                      pop esi
    'Reference to '__except_list'
'0071452f    64890d00000000          mov dword ptr fs:[00000000], ecx
'00714536    5b                      pop ebx
'00714537    8be5                    mov esp, ebp
'00714539    5d                      pop ebp
'0071453a    c20800                  ret 0008


End Function


Public Function NouveauDon(arg_0 As Unknow, arg_1 As Unknow, arg_2 As Unknow, arg_3 As Unknow, arg_4 As Unknow, arg_5 As Unknow, arg_6 As Unknow, arg_7 As Unknow, arg_8 As Unknow, arg_9 As Unknow, arg_A As Unknow, arg_B As Unknow, arg_C As Unknow, arg_D As Unknow, arg_E As Unknow, arg_F As Unknow, arg_10 As Unknow, arg_11 As Unknow, arg_12 As Unknow, arg_13 As Unknow, arg_14 As Unknow, arg_15 As Unknow, arg_16 As Unknow, arg_17 As Unknow, arg_18 As Unknow, arg_19 As Unknow, arg_1A As Unknow, arg_1B As Unknow, arg_1C As Unknow, arg_1D As Unknow, arg_1E As Unknow, arg_1F As Unknow, arg_20 As Unknow, arg_21 As Unknow, arg_22 As Unknow, arg_23 As Unknow, arg_24 As Unknow, arg_25 As Unknow, arg_26 As Unknow, arg_27 As Unknow, arg_28 As Unknow, arg_29 As Unknow, arg_2A As Unknow, arg_2B As Unknow, arg_2C As Unknow, arg_2D As Unknow, arg_2E As Unknow, arg_2F As Unknow, arg_30 As Unknow, arg_31 As Unknow, arg_32 As Unknow, arg_33 As Unknow, arg_34 As Unknow, arg_35 As Unknow, arg_36 As Unknow, arg_37 As Unknow, arg_38 As Unknow, arg_39 As Unknow, arg_3A As Unknow, arg_3B As Unknow, arg_3C As Unknow)
'00714540    55                      push ebp
'00714541    8bec                    mov ebp, esp
'00714543    83ec0c                  sub esp, 0c

' *** Reference to "__vbaExceptHandler"
'00714546    6866724000              push 00407266
'0071454b    64a100000000            mov ax, word ptr fs:[00000000]
'00714551    50                      push eax
    'Reference to '__except_list'
'00714552    64892500000000          mov dword ptr fs:[00000000], esp
'00714559    83ec34                  sub esp, 34
'0071455c    53                      push ebx
'0071455d    56                      push esi
'0071455e    57                      push edi
'0071455f    8965f4                  mov dword ptr [ebp-0c], esp
'00714562    c745f8a06d4000          mov dword ptr [ebp-08], 00406da0
'00714569    33ff                    xor edi, edi
'0071456b    897dfc                  mov dword ptr [ebp-04], edi
'0071456e    8b7508                  mov esi, dword ptr [ebp+08]
'00714571    8b06                    mov eax, dword ptr [esi]
'00714573    56                      push esi
'00714574    ff5004                  call dword ptr [eax+04]
'00714577    8b0e                    mov ecx, dword ptr [esi]
'00714579    56                      push esi
'0071457a    897de8                  mov dword ptr [ebp-18], edi
'0071457d    897de4                  mov dword ptr [ebp-1c], edi
'00714580    897dd0                  mov dword ptr [ebp-30], edi
'00714583    ff9108030000            call dword ptr [ecx+00000308]

' *** Reference to "__vbaObjSet"
'00714589    8b3db4104000            mov edi, dword ptr [004010b4]
'0071458f    50                      push eax
'00714590    8d55e8                  lea edx, dword ptr [ebp-18]
'00714593    52                      push edx
'00714594    ffd7                    call edi
Set var_41 = Nothing
'00714596    8bd8                    mov ebx, eax
'00714598    8b0b                    mov ecx, dword ptr [ebx]
'0071459a    83ec10                  sub esp, 10
'0071459d    8bd4                    mov edx, esp
'0071459f    b80a000000              mov eax, 0000000a
'007145a4    8902                    mov dword ptr [edx], eax
'007145a6    8b45d8                  mov eax, dword ptr [ebp-28]
'007145a9    894204                  mov dword ptr [edx+04], eax
'007145ac    b804000280              mov eax, 80020004
'007145b1    894208                  mov dword ptr [edx+08], eax
'007145b4    8b45e0                  mov eax, dword ptr [ebp-20]
'007145b7    68cc134300              push 004313cc
'007145bc    53                      push ebx
'007145bd    89420c                  mov dword ptr [edx+0c], eax
'007145c0    ff91ec010000            call dword ptr [ecx+000001ec]
'007145c6    dbe2                    fnclex
'007145c8    85c0                    test ax, ax
'007145ca    7d12                    jge 7145de
'007145cc    68ec010000              push 000001ec
'007145d1    681cb94300              push 0043b91c
'007145d6    53                      push ebx
'007145d7    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'007145d8    ff1580104000            call dword ptr [00401080]

' *** Reference to "__vbaFreeObj"
'007145de    8b1d08134000            mov ebx, dword ptr [00401308]
'007145e4    8d4de8                  lea ecx, dword ptr [ebp-18]
'007145e7    ffd3                    call ebx
'#Cleanup(var_41)
'007145e9    8b0e                    mov ecx, dword ptr [esi]
'007145eb    56                      push esi
'007145ec    ff9108030000            call dword ptr [ecx+00000308]
'007145f2    50                      push eax
'007145f3    8d55e4                  lea edx, dword ptr [ebp-1c]
'007145f6    52                      push edx
'007145f7    ffd7                    call edi
Set var_40 = Nothing
'007145f9    8945c4                  mov dword ptr [ebp-3c], eax
'007145fc    8b06                    mov eax, dword ptr [esi]
'007145fe    56                      push esi
'007145ff    ff9008030000            call dword ptr [eax+00000308]
'00714605    50                      push eax
'00714606    8d4de8                  lea ecx, dword ptr [ebp-18]
'00714609    51                      push ecx
'0071460a    ffd7                    call edi
Set var_41 = Nothing
'0071460c    8b10                    mov edx, dword ptr [eax]
'0071460e    8d4dd0                  lea ecx, dword ptr [ebp-30]
'00714611    51                      push ecx
'00714612    50                      push eax
'00714613    8945cc                  mov dword ptr [ebp-34], eax
'00714616    ff92e8000000            call dword ptr [edx+000000e8]
'0071461c    dbe2                    fnclex
'0071461e    85c0                    test ax, ax
'00714620    7d15                    jge 714637
'00714622    8b55cc                  mov edx, dword ptr [ebp-34]
'00714625    68e8000000              push 000000e8
'0071462a    681cb94300              push 0043b91c
'0071462f    52                      push edx
'00714630    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00714631    ff1580104000            call dword ptr [00401080]
'00714637    668b55d0                mov dx, word ptr [ebp-30]
'0071463b    8b45c4                  mov eax, dword ptr [ebp-3c]
'0071463e    8b08                    mov ecx, dword ptr [eax]
'00714640    6683ea01                sub dx, 01
var_num4 = 0 - 1
'00714644    0f80e0030000            jo 714a2a
'0071464a    52                      push edx
'0071464b    50                      push eax
'0071464c    ff91f4000000            call dword ptr [ecx+000000f4]
'00714652    dbe2                    fnclex
'00714654    85c0                    test ax, ax
'00714656    7d15                    jge 71466d
'00714658    8b4dc4                  mov ecx, dword ptr [ebp-3c]
'0071465b    68f4000000              push 000000f4
'00714660    681cb94300              push 0043b91c
'00714665    51                      push ecx
'00714666    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00714667    ff1580104000            call dword ptr [00401080]
'0071466d    8d55e4                  lea edx, dword ptr [ebp-1c]
'00714670    52                      push edx
'00714671    8d45e8                  lea eax, dword ptr [ebp-18]
'00714674    50                      push eax
'00714675    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'00714677    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_41, var_40)
'0071467d    8b0e                    mov ecx, dword ptr [esi]
'0071467f    83c40c                  add esp, 0c
'00714682    56                      push esi
'00714683    ff9100030000            call dword ptr [ecx+00000300]
'00714689    50                      push eax
'0071468a    8d55e8                  lea edx, dword ptr [ebp-18]
'0071468d    52                      push edx
'0071468e    ffd7                    call edi
Set var_41 = 
'00714690    8b08                    mov ecx, dword ptr [eax]
'00714692    6a00                    push 00
'00714694    50                      push eax
'00714695    8945cc                  mov dword ptr [ebp-34], eax
'00714698    ff91f4000000            call dword ptr [ecx+000000f4]
'0071469e    dbe2                    fnclex
'007146a0    85c0                    test ax, ax
'007146a2    7d15                    jge 7146b9
'007146a4    8b55cc                  mov edx, dword ptr [ebp-34]
'007146a7    68f4000000              push 000000f4
'007146ac    681cb94300              push 0043b91c
'007146b1    52                      push edx
'007146b2    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'007146b3    ff1580104000            call dword ptr [00401080]
'007146b9    8d4de8                  lea ecx, dword ptr [ebp-18]
'007146bc    ffd3                    call ebx
'#Cleanup(var_41)
'007146be    8b06                    mov eax, dword ptr [esi]
'007146c0    56                      push esi
'007146c1    ff9004030000            call dword ptr [eax+00000304]
'007146c7    50                      push eax
'007146c8    8d4de8                  lea ecx, dword ptr [ebp-18]
'007146cb    51                      push ecx
'007146cc    ffd7                    call edi
Set var_41 = Nothing
'007146ce    8b10                    mov edx, dword ptr [eax]
'007146d0    6a00                    push 00
'007146d2    50                      push eax
'007146d3    8945cc                  mov dword ptr [ebp-34], eax
'007146d6    ff92f4000000            call dword ptr [edx+000000f4]
'007146dc    dbe2                    fnclex
'007146de    85c0                    test ax, ax
'007146e0    7d15                    jge 7146f7
'007146e2    8b4dcc                  mov ecx, dword ptr [ebp-34]
'007146e5    68f4000000              push 000000f4
'007146ea    681cb94300              push 0043b91c
'007146ef    51                      push ecx
'007146f0    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'007146f1    ff1580104000            call dword ptr [00401080]
'007146f7    8d4de8                  lea ecx, dword ptr [ebp-18]
'007146fa    ffd3                    call ebx
'#Cleanup(var_41)
'007146fc    8b16                    mov edx, dword ptr [esi]
'007146fe    56                      push esi
'007146ff    ff921c030000            call dword ptr [edx+0000031c]
'00714705    50                      push eax
'00714706    8d45e8                  lea eax, dword ptr [ebp-18]
'00714709    50                      push eax
'0071470a    ffd7                    call edi
Set var_41 = Nothing
'0071470c    8b08                    mov ecx, dword ptr [eax]
'0071470e    68646b4500              push 00456b64
'00714713    50                      push eax
'00714714    8945cc                  mov dword ptr [ebp-34], eax
'00714717    ff91a4000000            call dword ptr [ecx+000000a4]
'0071471d    dbe2                    fnclex
'0071471f    85c0                    test ax, ax
'00714721    7d15                    jge 714738
'00714723    8b55cc                  mov edx, dword ptr [ebp-34]
'00714726    68a4000000              push 000000a4
'0071472b    68d00d4300              push 00430dd0
'00714730    52                      push edx
'00714731    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00714732    ff1580104000            call dword ptr [00401080]
'00714738    8d4de8                  lea ecx, dword ptr [ebp-18]
'0071473b    ffd3                    call ebx
'#Cleanup(var_41)
'0071473d    8b06                    mov eax, dword ptr [esi]
'0071473f    56                      push esi
'00714740    ff9018030000            call dword ptr [eax+00000318]
'00714746    50                      push eax
'00714747    8d4de8                  lea ecx, dword ptr [ebp-18]
'0071474a    51                      push ecx
'0071474b    ffd7                    call edi
Set var_41 = Nothing
'0071474d    8b10                    mov edx, dword ptr [eax]
'0071474f    68cc134300              push 004313cc
'00714754    50                      push eax
'00714755    8945cc                  mov dword ptr [ebp-34], eax
'00714758    ff92a4000000            call dword ptr [edx+000000a4]
'0071475e    dbe2                    fnclex
'00714760    85c0                    test ax, ax
'00714762    7d15                    jge 714779
'00714764    8b4dcc                  mov ecx, dword ptr [ebp-34]
'00714767    68a4000000              push 000000a4
'0071476c    68d00d4300              push 00430dd0
'00714771    51                      push ecx
'00714772    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00714773    ff1580104000            call dword ptr [00401080]
'00714779    8d4de8                  lea ecx, dword ptr [ebp-18]
'0071477c    ffd3                    call ebx
'#Cleanup(var_41)
'0071477e    8b16                    mov edx, dword ptr [esi]
'00714780    56                      push esi
'00714781    ff9234030000            call dword ptr [edx+00000334]
'00714787    50                      push eax
'00714788    8d45e8                  lea eax, dword ptr [ebp-18]
'0071478b    50                      push eax
'0071478c    ffd7                    call edi
Set var_41 = Nothing
'0071478e    8b08                    mov ecx, dword ptr [eax]
'00714790    68cc134300              push 004313cc
'00714795    50                      push eax
'00714796    8945cc                  mov dword ptr [ebp-34], eax
'00714799    ff91a4000000            call dword ptr [ecx+000000a4]
'0071479f    dbe2                    fnclex
'007147a1    85c0                    test ax, ax
'007147a3    7d15                    jge 7147ba
'007147a5    8b55cc                  mov edx, dword ptr [ebp-34]
'007147a8    68a4000000              push 000000a4
'007147ad    68d00d4300              push 00430dd0
'007147b2    52                      push edx
'007147b3    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'007147b4    ff1580104000            call dword ptr [00401080]
'007147ba    8d4de8                  lea ecx, dword ptr [ebp-18]
'007147bd    ffd3                    call ebx
'#Cleanup(var_41)
'007147bf    8b06                    mov eax, dword ptr [esi]
'007147c1    56                      push esi
'007147c2    ff9038030000            call dword ptr [eax+00000338]
'007147c8    50                      push eax
'007147c9    8d4de8                  lea ecx, dword ptr [ebp-18]
'007147cc    51                      push ecx
'007147cd    ffd7                    call edi
Set var_41 = Nothing
'007147cf    8b10                    mov edx, dword ptr [eax]
'007147d1    68cc134300              push 004313cc
'007147d6    50                      push eax
'007147d7    8945cc                  mov dword ptr [ebp-34], eax
'007147da    ff92a4000000            call dword ptr [edx+000000a4]
'007147e0    dbe2                    fnclex
'007147e2    85c0                    test ax, ax
'007147e4    7d15                    jge 7147fb
'007147e6    8b4dcc                  mov ecx, dword ptr [ebp-34]
'007147e9    68a4000000              push 000000a4
'007147ee    68d00d4300              push 00430dd0
'007147f3    51                      push ecx
'007147f4    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'007147f5    ff1580104000            call dword ptr [00401080]
'007147fb    8d4de8                  lea ecx, dword ptr [ebp-18]
'007147fe    ffd3                    call ebx
'#Cleanup(var_41)
'00714800    8b16                    mov edx, dword ptr [esi]
'00714802    56                      push esi
'00714803    ff92fc020000            call dword ptr [edx+000002fc]
'00714809    50                      push eax
'0071480a    8d45e8                  lea eax, dword ptr [ebp-18]
'0071480d    50                      push eax
'0071480e    ffd7                    call edi
Set var_41 = Nothing
'00714810    8b08                    mov ecx, dword ptr [eax]
'00714812    68cc134300              push 004313cc
'00714817    50                      push eax
'00714818    8945cc                  mov dword ptr [ebp-34], eax
'0071481b    ff91ac000000            call dword ptr [ecx+000000ac]
'00714821    dbe2                    fnclex
'00714823    85c0                    test ax, ax
'00714825    7d15                    jge 71483c
'00714827    8b55cc                  mov edx, dword ptr [ebp-34]
'0071482a    68ac000000              push 000000ac
'0071482f    681cb94300              push 0043b91c
'00714834    52                      push edx
'00714835    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00714836    ff1580104000            call dword ptr [00401080]
'0071483c    8d4de8                  lea ecx, dword ptr [ebp-18]
'0071483f    ffd3                    call ebx
'#Cleanup(var_41)
'00714841    8b06                    mov eax, dword ptr [esi]
'00714843    56                      push esi
'00714844    ff900c030000            call dword ptr [eax+0000030c]
'0071484a    50                      push eax
'0071484b    8d4de8                  lea ecx, dword ptr [ebp-18]
'0071484e    51                      push ecx
'0071484f    ffd7                    call edi
Set var_41 = Nothing
'00714851    8b10                    mov edx, dword ptr [eax]
'00714853    6a00                    push 00
'00714855    50                      push eax
'00714856    8945cc                  mov dword ptr [ebp-34], eax
'00714859    ff92e4000000            call dword ptr [edx+000000e4]
'0071485f    dbe2                    fnclex
'00714861    85c0                    test ax, ax
'00714863    7d15                    jge 71487a
'00714865    8b4dcc                  mov ecx, dword ptr [ebp-34]
'00714868    68e4000000              push 000000e4
'0071486d    68dce24300              push 0043e2dc
'00714872    51                      push ecx
'00714873    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00714874    ff1580104000            call dword ptr [00401080]
'0071487a    8d4de8                  lea ecx, dword ptr [ebp-18]
'0071487d    ffd3                    call ebx
'#Cleanup(var_41)
'0071487f    8b16                    mov edx, dword ptr [esi]
'00714881    56                      push esi
'00714882    ff9244030000            call dword ptr [edx+00000344]
'00714888    50                      push eax
'00714889    8d45e8                  lea eax, dword ptr [ebp-18]
'0071488c    50                      push eax
'0071488d    ffd7                    call edi
Set var_41 = Nothing
'0071488f    8b08                    mov ecx, dword ptr [eax]
'00714891    6a00                    push 00
'00714893    50                      push eax
'00714894    8945cc                  mov dword ptr [ebp-34], eax
'00714897    ff91e4000000            call dword ptr [ecx+000000e4]
'0071489d    dbe2                    fnclex
'0071489f    85c0                    test ax, ax
'007148a1    7d15                    jge 7148b8
'007148a3    8b55cc                  mov edx, dword ptr [ebp-34]
'007148a6    68e4000000              push 000000e4
'007148ab    68dce24300              push 0043e2dc
'007148b0    52                      push edx
'007148b1    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'007148b2    ff1580104000            call dword ptr [00401080]
'007148b8    8d4de8                  lea ecx, dword ptr [ebp-18]
'007148bb    ffd3                    call ebx
'#Cleanup(var_41)
'007148bd    8b06                    mov eax, dword ptr [esi]
'007148bf    56                      push esi
'007148c0    ff9010030000            call dword ptr [eax+00000310]
'007148c6    50                      push eax
'007148c7    8d4de8                  lea ecx, dword ptr [ebp-18]
'007148ca    51                      push ecx
'007148cb    ffd7                    call edi
Set var_41 = Nothing
'007148cd    8b10                    mov edx, dword ptr [eax]
'007148cf    6a00                    push 00
'007148d1    50                      push eax
'007148d2    8945cc                  mov dword ptr [ebp-34], eax
'007148d5    ff92e4000000            call dword ptr [edx+000000e4]
'007148db    dbe2                    fnclex
'007148dd    85c0                    test ax, ax
'007148df    7d15                    jge 7148f6
'007148e1    8b4dcc                  mov ecx, dword ptr [ebp-34]
'007148e4    68e4000000              push 000000e4
'007148e9    68dce24300              push 0043e2dc
'007148ee    51                      push ecx
'007148ef    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'007148f0    ff1580104000            call dword ptr [00401080]
'007148f6    8d4de8                  lea ecx, dword ptr [ebp-18]
'007148f9    ffd3                    call ebx
'#Cleanup(var_41)
'007148fb    8b16                    mov edx, dword ptr [esi]
'007148fd    56                      push esi
'007148fe    ff9214030000            call dword ptr [edx+00000314]
'00714904    50                      push eax
'00714905    8d45e8                  lea eax, dword ptr [ebp-18]
'00714908    50                      push eax
'00714909    ffd7                    call edi
Set var_41 = Nothing
'0071490b    8b08                    mov ecx, dword ptr [eax]
'0071490d    6a00                    push 00
'0071490f    50                      push eax
'00714910    8945cc                  mov dword ptr [ebp-34], eax
'00714913    ff91e4000000            call dword ptr [ecx+000000e4]
'00714919    dbe2                    fnclex
'0071491b    85c0                    test ax, ax
'0071491d    7d15                    jge 714934
'0071491f    8b55cc                  mov edx, dword ptr [ebp-34]
'00714922    68e4000000              push 000000e4
'00714927    68dce24300              push 0043e2dc
'0071492c    52                      push edx
'0071492d    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0071492e    ff1580104000            call dword ptr [00401080]
'00714934    8d4de8                  lea ecx, dword ptr [ebp-18]
'00714937    ffd3                    call ebx
'#Cleanup(var_41)
'00714939    8b06                    mov eax, dword ptr [esi]
'0071493b    56                      push esi
'0071493c    ff9048030000            call dword ptr [eax+00000348]
'00714942    50                      push eax
'00714943    8d4de8                  lea ecx, dword ptr [ebp-18]
'00714946    51                      push ecx
'00714947    ffd7                    call edi
Set var_41 = Nothing
'00714949    8b10                    mov edx, dword ptr [eax]
'0071494b    6a00                    push 00
'0071494d    50                      push eax
'0071494e    8945cc                  mov dword ptr [ebp-34], eax
'00714951    ff92e4000000            call dword ptr [edx+000000e4]
'00714957    dbe2                    fnclex
'00714959    85c0                    test ax, ax
'0071495b    7d15                    jge 714972
'0071495d    8b4dcc                  mov ecx, dword ptr [ebp-34]
'00714960    68e4000000              push 000000e4
'00714965    68dce24300              push 0043e2dc
'0071496a    51                      push ecx
'0071496b    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0071496c    ff1580104000            call dword ptr [00401080]
'00714972    8d4de8                  lea ecx, dword ptr [ebp-18]
'00714975    ffd3                    call ebx
'#Cleanup(var_41)
'00714977    8b16                    mov edx, dword ptr [esi]
'00714979    56                      push esi
'0071497a    ff923c030000            call dword ptr [edx+0000033c]
'00714980    50                      push eax
'00714981    8d45e8                  lea eax, dword ptr [ebp-18]
'00714984    50                      push eax
'00714985    ffd7                    call edi
Set var_41 = Nothing
'00714987    8b08                    mov ecx, dword ptr [eax]
'00714989    6a00                    push 00
'0071498b    50                      push eax
'0071498c    8945cc                  mov dword ptr [ebp-34], eax
'0071498f    ff91e4000000            call dword ptr [ecx+000000e4]
'00714995    dbe2                    fnclex
'00714997    85c0                    test ax, ax
'00714999    7d15                    jge 7149b0
'0071499b    8b55cc                  mov edx, dword ptr [ebp-34]
'0071499e    68e4000000              push 000000e4
'007149a3    68dce24300              push 0043e2dc
'007149a8    52                      push edx
'007149a9    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'007149aa    ff1580104000            call dword ptr [00401080]
'007149b0    8d4de8                  lea ecx, dword ptr [ebp-18]
'007149b3    ffd3                    call ebx
'#Cleanup(var_41)
'007149b5    8b06                    mov eax, dword ptr [esi]
'007149b7    56                      push esi
'007149b8    ff9040030000            call dword ptr [eax+00000340]
'007149be    50                      push eax
'007149bf    8d4de8                  lea ecx, dword ptr [ebp-18]
'007149c2    51                      push ecx
'007149c3    ffd7                    call edi
Set var_41 = Nothing
'007149c5    8bf0                    mov esi, eax
'007149c7    8b16                    mov edx, dword ptr [esi]
'007149c9    6a00                    push 00
'007149cb    56                      push esi
'007149cc    ff92e4000000            call dword ptr [edx+000000e4]
'007149d2    dbe2                    fnclex
'007149d4    85c0                    test ax, ax
'007149d6    7d12                    jge 7149ea
'007149d8    68e4000000              push 000000e4
'007149dd    68dce24300              push 0043e2dc
'007149e2    56                      push esi
'007149e3    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'007149e4    ff1580104000            call dword ptr [00401080]
'007149ea    8d4de8                  lea ecx, dword ptr [ebp-18]
'007149ed    ffd3                    call ebx
'#Cleanup(var_41)
'007149ef    680b4a7100              push 00714a0b
'007149f4    eb14                    jmp 714a0a
'007149f6    8d45e4                  lea eax, dword ptr [ebp-1c]
'007149f9    50                      push eax
'007149fa    8d4de8                  lea ecx, dword ptr [ebp-18]
'007149fd    51                      push ecx
'007149fe    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'00714a00    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_41, var_40)
'00714a06    83c40c                  add esp, 0c
'00714a09    c3                      ret
'00714a0a    c3                      ret
'00714a0b    8b4508                  mov eax, dword ptr [ebp+08]
'00714a0e    8b10                    mov edx, dword ptr [eax]
'00714a10    50                      push eax
'00714a11    ff5208                  call dword ptr [edx+08]
'00714a14    8b45fc                  mov eax, dword ptr [ebp-04]
'00714a17    8b4dec                  mov ecx, dword ptr [ebp-14]
'00714a1a    5f                      pop edi
'00714a1b    5e                      pop esi
    'Reference to '__except_list'
'00714a1c    64890d00000000          mov dword ptr fs:[00000000], ecx
'00714a23    5b                      pop ebx
'00714a24    8be5                    mov esp, ebp
'00714a26    5d                      pop ebp
'00714a27    c20400                  ret 0004


End Function


'Event for BtnFermer
Private Sub BtnFermer_Click()
'00711870    55                      push ebp
'00711871    8bec                    mov ebp, esp
'00711873    83ec0c                  sub esp, 0c

' *** Reference to "__vbaExceptHandler"
'00711876    6866724000              push 00407266
'0071187b    64a100000000            mov ax, word ptr fs:[00000000]
'00711881    50                      push eax
    'Reference to '__except_list'
'00711882    64892500000000          mov dword ptr fs:[00000000], esp
'00711889    83ec28                  sub esp, 28
'0071188c    53                      push ebx
'0071188d    56                      push esi
'0071188e    57                      push edi
'0071188f    8965f4                  mov dword ptr [ebp-0c], esp
'00711892    c745f8286d4000          mov dword ptr [ebp-08], 00406d28
'00711899    8b7d08                  mov edi, dword ptr [ebp+08]
'0071189c    8bc7                    mov eax, edi
'0071189e    83e001                  and eax, 01
'007118a1    8945fc                  mov dword ptr [ebp-04], eax
'007118a4    83e7fe                  and edi, fffffffe
'007118a7    8b0f                    mov ecx, dword ptr [edi]
'007118a9    57                      push edi
'007118aa    897d08                  mov dword ptr [ebp+08], edi
'007118ad    ff5104                  call dword ptr [ecx+04]
'007118b0    6848b07200              push 0072b048
'007118b5    8d55d8                  lea edx, dword ptr [ebp-28]
'007118b8    33db                    xor ebx, ebx
'007118ba    52                      push edx
'007118bb    895de8                  mov dword ptr [ebp-18], ebx
'007118be    895dd8                  mov dword ptr [ebp-28], ebx
'007118c1    e88aa0ddff              call 4eb950
Call sub_4EB950()
'007118c6    8d4dd8                  lea ecx, dword ptr [ebp-28]

' *** Reference to "__vbaFreeVar"
'007118c9    ff1528104000            call dword ptr [00401028]
'#Cleanup(var_45)
'007118cf    391d24be7200            cmp dword ptr [0072be24], ebx
'007118d5    7510                    jne 7118e7
'007118d7    6824be7200              push 0072be24
'007118dc    6870004300              push 00430070

' *** Reference to "__vbaNew2"
'007118e1    ff152c124000            call dword ptr [0040122c]
Dim var_2 As New Global
'007118e7    8b3524be7200            mov esi, dword ptr [0072be24]
'007118ed    8b16                    mov edx, dword ptr [esi]
'007118ef    57                      push edi
'007118f0    8d45e8                  lea eax, dword ptr [ebp-18]
'007118f3    50                      push eax
'007118f4    8955c4                  mov dword ptr [ebp-3c], edx

' *** Reference to "__vbaObjSetAddref"
'007118f7    ff15c8104000            call dword ptr [004010c8]
Set var_41 = Me
'007118fd    8b4dc4                  mov ecx, dword ptr [ebp-3c]
'00711900    50                      push eax
'00711901    56                      push esi
'00711902    ff5110                  call dword ptr [ecx+10]
Call var_2.Unload(var_41)
'00711905    dbe2                    fnclex
'00711907    3bc3                    cmp eax, ebx
'00711909    7d0f                    jge 71191a
'0071190b    6a10                    push 10
'0071190d    6860004300              push 00430060
'00711912    56                      push esi
'00711913    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00711914    ff1580104000            call dword ptr [00401080]
'0071191a    8d4de8                  lea ecx, dword ptr [ebp-18]

' *** Reference to "__vbaFreeObj"
'0071191d    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_41)
'00711923    895dfc                  mov dword ptr [ebp-04], ebx
'00711926    6841197100              push 00711941
'0071192b    eb13                    jmp 711940
'0071192d    8d4de8                  lea ecx, dword ptr [ebp-18]

' *** Reference to "__vbaFreeObj"
'00711930    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_41)
'00711936    8d4dd8                  lea ecx, dword ptr [ebp-28]

' *** Reference to "__vbaFreeVar"
'00711939    ff1528104000            call dword ptr [00401028]
'#Cleanup(var_45)
'0071193f    c3                      ret
'00711940    c3                      ret
'00711941    8b4508                  mov eax, dword ptr [ebp+08]
'00711944    8b10                    mov edx, dword ptr [eax]
'00711946    50                      push eax
'00711947    ff5208                  call dword ptr [edx+08]
'0071194a    8b45fc                  mov eax, dword ptr [ebp-04]
'0071194d    8b4dec                  mov ecx, dword ptr [ebp-14]
'00711950    5f                      pop edi
'00711951    5e                      pop esi
    'Reference to '__except_list'
'00711952    64890d00000000          mov dword ptr fs:[00000000], ecx
'00711959    5b                      pop ebx
'0071195a    8be5                    mov esp, ebp
'0071195c    5d                      pop ebp
'0071195d    c20400                  ret 0004


End Sub


'Event for cmbNom
Private Sub cmbNom_Click()
'00711f10    55                      push ebp
'00711f11    8bec                    mov ebp, esp
'00711f13    83ec0c                  sub esp, 0c

' *** Reference to "__vbaExceptHandler"
'00711f16    6866724000              push 00407266
'00711f1b    64a100000000            mov ax, word ptr fs:[00000000]
'00711f21    50                      push eax
    'Reference to '__except_list'
'00711f22    64892500000000          mov dword ptr fs:[00000000], esp
'00711f29    83ec20                  sub esp, 20
'00711f2c    53                      push ebx
'00711f2d    56                      push esi
'00711f2e    57                      push edi
'00711f2f    8965f4                  mov dword ptr [ebp-0c], esp
'00711f32    c745f8706d4000          mov dword ptr [ebp-08], 00406d70
'00711f39    8b7508                  mov esi, dword ptr [ebp+08]
'00711f3c    8bc6                    mov eax, esi
'00711f3e    83e001                  and eax, 01
'00711f41    8945fc                  mov dword ptr [ebp-04], eax
'00711f44    83e6fe                  and esi, fffffffe
'00711f47    8b0e                    mov ecx, dword ptr [esi]
'00711f49    56                      push esi
'00711f4a    897508                  mov dword ptr [ebp+08], esi
'00711f4d    ff5104                  call dword ptr [ecx+04]
'00711f50    8b16                    mov edx, dword ptr [esi]
'00711f52    33db                    xor ebx, ebx
'00711f54    56                      push esi
'00711f55    895de8                  mov dword ptr [ebp-18], ebx
'00711f58    895de4                  mov dword ptr [ebp-1c], ebx
'00711f5b    895de0                  mov dword ptr [ebp-20], ebx
'00711f5e    ff9208030000            call dword ptr [edx+00000308]
'00711f64    50                      push eax
'00711f65    8d45e8                  lea eax, dword ptr [ebp-18]
'00711f68    50                      push eax

' *** Reference to "__vbaObjSet"
'00711f69    ff15b4104000            call dword ptr [004010b4]
Set var_41 = Me
'00711f6f    8d55e4                  lea edx, dword ptr [ebp-1c]
'00711f72    8bf8                    mov edi, eax
'00711f74    8b0f                    mov ecx, dword ptr [edi]
'00711f76    52                      push edx
'00711f77    57                      push edi
'00711f78    ff91f0000000            call dword ptr [ecx+000000f0]
'00711f7e    dbe2                    fnclex
'00711f80    3bc3                    cmp eax, ebx
'00711f82    7d12                    jge 711f96

If (var_41 < 0) Then
'00711f84    68f0000000              push 000000f0
'00711f89    681cb94300              push 0043b91c
'00711f8e    57                      push edi
'00711f8f    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00711f90    ff1580104000            call dword ptr [00401080]
    
End If
'00711f96    8b45e4                  mov eax, dword ptr [ebp-1c]
'00711f99    8b0e                    mov ecx, dword ptr [esi]
'00711f9b    8d55e0                  lea edx, dword ptr [ebp-20]
'00711f9e    52                      push edx
'00711f9f    56                      push esi
'00711fa0    8945e0                  mov dword ptr [ebp-20], eax
'00711fa3    ff91f8060000            call dword ptr [ecx+000006f8]
'00711fa9    3bc3                    cmp eax, ebx
'00711fab    7d12                    jge 711fbf

If (0 < 0) Then
'00711fad    68f8060000              push 000006f8
'00711fb2    68e4194300              push 004319e4
'00711fb7    56                      push esi
'00711fb8    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00711fb9    ff1580104000            call dword ptr [00401080]
    
End If
'00711fbf    8d4de8                  lea ecx, dword ptr [ebp-18]

' *** Reference to "__vbaFreeObj"
'00711fc2    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_41)
'00711fc8    895dfc                  mov dword ptr [ebp-04], ebx
'00711fcb    68dd1f7100              push 00711fdd
'00711fd0    eb0a                    jmp 711fdc
'00711fd2    8d4de8                  lea ecx, dword ptr [ebp-18]

' *** Reference to "__vbaFreeObj"
'00711fd5    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_41)
'00711fdb    c3                      ret
'00711fdc    c3                      ret
'00711fdd    8b4508                  mov eax, dword ptr [ebp+08]
'00711fe0    8b08                    mov ecx, dword ptr [eax]
'00711fe2    50                      push eax
'00711fe3    ff5108                  call dword ptr [ecx+08]
'00711fe6    8b45fc                  mov eax, dword ptr [ebp-04]
'00711fe9    8b4dec                  mov ecx, dword ptr [ebp-14]
'00711fec    5f                      pop edi
'00711fed    5e                      pop esi
    'Reference to '__except_list'
'00711fee    64890d00000000          mov dword ptr fs:[00000000], ecx
'00711ff5    5b                      pop ebx
'00711ff6    8be5                    mov esp, ebp
'00711ff8    5d                      pop ebp
'00711ff9    c20400                  ret 0004


End Sub


'Event for btnSuivant
Private Sub btnSuivant_Click()
'00711bb0    55                      push ebp
'00711bb1    8bec                    mov ebp, esp
'00711bb3    83ec0c                  sub esp, 0c

' *** Reference to "__vbaExceptHandler"
'00711bb6    6866724000              push 00407266
'00711bbb    64a100000000            mov ax, word ptr fs:[00000000]
'00711bc1    50                      push eax
    'Reference to '__except_list'
'00711bc2    64892500000000          mov dword ptr fs:[00000000], esp
'00711bc9    83ec2c                  sub esp, 2c
'00711bcc    53                      push ebx
'00711bcd    56                      push esi
'00711bce    57                      push edi
'00711bcf    8965f4                  mov dword ptr [ebp-0c], esp
'00711bd2    c745f8506d4000          mov dword ptr [ebp-08], 00406d50
'00711bd9    8b7508                  mov esi, dword ptr [ebp+08]
'00711bdc    8bc6                    mov eax, esi
'00711bde    83e001                  and eax, 01
'00711be1    8945fc                  mov dword ptr [ebp-04], eax
'00711be4    83e6fe                  and esi, fffffffe
'00711be7    8b0e                    mov ecx, dword ptr [esi]
'00711be9    56                      push esi
'00711bea    897508                  mov dword ptr [ebp+08], esi
'00711bed    ff5104                  call dword ptr [ecx+04]
'00711bf0    8b16                    mov edx, dword ptr [esi]
'00711bf2    33db                    xor ebx, ebx
'00711bf4    56                      push esi
'00711bf5    895de8                  mov dword ptr [ebp-18], ebx
'00711bf8    895de4                  mov dword ptr [ebp-1c], ebx
'00711bfb    895de0                  mov dword ptr [ebp-20], ebx
'00711bfe    895ddc                  mov dword ptr [ebp-24], ebx
'00711c01    ff9208030000            call dword ptr [edx+00000308]
'00711c07    50                      push eax
'00711c08    8d45e8                  lea eax, dword ptr [ebp-18]
'00711c0b    50                      push eax

' *** Reference to "__vbaObjSet"
'00711c0c    ff15b4104000            call dword ptr [004010b4]
Set var_41 = Me
'00711c12    8d55e0                  lea edx, dword ptr [ebp-20]
'00711c15    8bf8                    mov edi, eax
'00711c17    8b0f                    mov ecx, dword ptr [edi]
'00711c19    52                      push edx
'00711c1a    57                      push edi
'00711c1b    ff91e8000000            call dword ptr [ecx+000000e8]
'00711c21    dbe2                    fnclex
'00711c23    3bc3                    cmp eax, ebx
'00711c25    7d16                    jge 711c3d

If (var_41 < 0) Then

' *** Reference to "__vbaHresultCheckObj"
'00711c27    8b1d80104000            mov ebx, dword ptr [00401080]
'00711c2d    68e8000000              push 000000e8
'00711c32    681cb94300              push 0043b91c
'00711c37    57                      push edi
'00711c38    50                      push eax
'00711c39    ffd3                    call ebx
'00711c3b    eb06                    jmp 711c43
    
End If

' *** Reference to "__vbaHresultCheckObj"
'00711c3d    8b1d80104000            mov ebx, dword ptr [00401080]
'00711c43    8b06                    mov eax, dword ptr [esi]
'00711c45    56                      push esi
'00711c46    ff9008030000            call dword ptr [eax+00000308]
'00711c4c    50                      push eax
'00711c4d    8d4de4                  lea ecx, dword ptr [ebp-1c]
'00711c50    51                      push ecx

' *** Reference to "__vbaObjSet"
'00711c51    ff15b4104000            call dword ptr [004010b4]
Set var_40 = Nothing
'00711c57    8bf8                    mov edi, eax
'00711c59    8b17                    mov edx, dword ptr [edi]
'00711c5b    8d45dc                  lea eax, dword ptr [ebp-24]
'00711c5e    50                      push eax
'00711c5f    57                      push edi
'00711c60    ff92f0000000            call dword ptr [edx+000000f0]
'00711c66    dbe2                    fnclex
'00711c68    85c0                    test ax, ax
'00711c6a    7d0e                    jge 711c7a
'00711c6c    68f0000000              push 000000f0
'00711c71    681cb94300              push 0043b91c
'00711c76    57                      push edi
'00711c77    50                      push eax
'00711c78    ffd3                    call ebx
'00711c7a    668b4ddc                mov cx, word ptr [ebp-24]
'00711c7e    6683c101                add cx, 01
var_num3 = 0 + 1
'00711c82    0f80fb000000            jo 711d83
'00711c88    33d2                    xor edx, edx
var_num4 = Empty
'00711c8a    66394de0                cmp word ptr [ebp-20], cx
'00711c8e    8d45e4                  lea eax, dword ptr [ebp-1c]
'00711c91    0f94c2                  sete dl
'00711c94    50                      push eax
'00711c95    8d4de8                  lea ecx, dword ptr [ebp-18]
'00711c98    51                      push ecx
'00711c99    6a02                    push 02
'00711c9b    f7da                    neg edx
'00711c9d    8bfa                    mov edi, edx

' *** Reference to "__vbaFreeObjList"
'00711c9f    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_41, var_40)
'00711ca5    83c40c                  add esp, 0c
'00711ca8    6685ff                  test edi, edi
'00711cab    7428                    je 711cd5

If (0 = var_num3) Then
'00711cad    8b16                    mov edx, dword ptr [esi]
'00711caf    8d45e0                  lea eax, dword ptr [ebp-20]
'00711cb2    50                      push eax
'00711cb3    56                      push esi
'00711cb4    c745e000000000          mov dword ptr [ebp-20], 00000000
'00711cbb    ff92f8060000            call dword ptr [edx+000006f8]
'00711cc1    85c0                    test ax, ax
'00711cc3    7d7c                    jge 711d41
'00711cc5    68f8060000              push 000006f8
'00711cca    68e4194300              push 004319e4
'00711ccf    56                      push esi
'00711cd0    50                      push eax
'00711cd1    ffd3                    call ebx
'00711cd3    eb6c                    jmp 711d41
    
End If
'00711cd5    8b0e                    mov ecx, dword ptr [esi]
'00711cd7    56                      push esi
'00711cd8    ff9108030000            call dword ptr [ecx+00000308]
'00711cde    50                      push eax
'00711cdf    8d55e8                  lea edx, dword ptr [ebp-18]
'00711ce2    52                      push edx

' *** Reference to "__vbaObjSet"
'00711ce3    ff15b4104000            call dword ptr [004010b4]
Set var_41 = 
'00711ce9    8d4de0                  lea ecx, dword ptr [ebp-20]
'00711cec    8bf8                    mov edi, eax
'00711cee    8b07                    mov eax, dword ptr [edi]
'00711cf0    51                      push ecx
'00711cf1    57                      push edi
'00711cf2    ff90f0000000            call dword ptr [eax+000000f0]
'00711cf8    dbe2                    fnclex
'00711cfa    85c0                    test ax, ax
'00711cfc    7d0e                    jge 711d0c
'00711cfe    68f0000000              push 000000f0
'00711d03    681cb94300              push 0043b91c
'00711d08    57                      push edi
'00711d09    50                      push eax
'00711d0a    ffd3                    call ebx
'00711d0c    668b55e0                mov dx, word ptr [ebp-20]
'00711d10    8b06                    mov eax, dword ptr [esi]
'00711d12    6683c201                add dx, 01
var_num4 = var_3 + 1
'00711d16    706b                    jo 711d83
'00711d18    8d4ddc                  lea ecx, dword ptr [ebp-24]
'00711d1b    51                      push ecx
'00711d1c    56                      push esi
'00711d1d    8955dc                  mov dword ptr [ebp-24], edx
'00711d20    ff90f8060000            call dword ptr [eax+000006f8]
'00711d26    85c0                    test ax, ax
'00711d28    7d0e                    jge 711d38
'00711d2a    68f8060000              push 000006f8
'00711d2f    68e4194300              push 004319e4
'00711d34    56                      push esi
'00711d35    50                      push eax
'00711d36    ffd3                    call ebx
'00711d38    8d4de8                  lea ecx, dword ptr [ebp-18]

' *** Reference to "__vbaFreeObj"
'00711d3b    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_41)
'00711d41    c745fc00000000          mov dword ptr [ebp-04], 00000000
'00711d48    68641d7100              push 00711d64
'00711d4d    eb14                    jmp 711d63
'00711d4f    8d55e4                  lea edx, dword ptr [ebp-1c]
'00711d52    52                      push edx
'00711d53    8d45e8                  lea eax, dword ptr [ebp-18]
'00711d56    50                      push eax
'00711d57    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'00711d59    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_41, var_40)
'00711d5f    83c40c                  add esp, 0c
'00711d62    c3                      ret
'00711d63    c3                      ret
'00711d64    8b4508                  mov eax, dword ptr [ebp+08]
'00711d67    8b08                    mov ecx, dword ptr [eax]
'00711d69    50                      push eax
'00711d6a    ff5108                  call dword ptr [ecx+08]
'00711d6d    8b45fc                  mov eax, dword ptr [ebp-04]
'00711d70    8b4dec                  mov ecx, dword ptr [ebp-14]
'00711d73    5f                      pop edi
'00711d74    5e                      pop esi
    'Reference to '__except_list'
'00711d75    64890d00000000          mov dword ptr fs:[00000000], ecx
'00711d7c    5b                      pop ebx
'00711d7d    8be5                    mov esp, ebp
'00711d7f    5d                      pop ebp
'00711d80    c20400                  ret 0004


End Sub


'Event for btnPrecedent
Private Sub btnPrecedent_Click()
'007119f0    55                      push ebp
'007119f1    8bec                    mov ebp, esp
'007119f3    83ec0c                  sub esp, 0c

' *** Reference to "__vbaExceptHandler"
'007119f6    6866724000              push 00407266
'007119fb    64a100000000            mov ax, word ptr fs:[00000000]
'00711a01    50                      push eax
    'Reference to '__except_list'
'00711a02    64892500000000          mov dword ptr fs:[00000000], esp
'00711a09    83ec20                  sub esp, 20
'00711a0c    53                      push ebx
'00711a0d    56                      push esi
'00711a0e    57                      push edi
'00711a0f    8965f4                  mov dword ptr [ebp-0c], esp
'00711a12    c745f8406d4000          mov dword ptr [ebp-08], 00406d40
'00711a19    8b7508                  mov esi, dword ptr [ebp+08]
'00711a1c    8bc6                    mov eax, esi
'00711a1e    83e001                  and eax, 01
'00711a21    8945fc                  mov dword ptr [ebp-04], eax
'00711a24    83e6fe                  and esi, fffffffe
'00711a27    8b0e                    mov ecx, dword ptr [esi]
'00711a29    56                      push esi
'00711a2a    897508                  mov dword ptr [ebp+08], esi
'00711a2d    ff5104                  call dword ptr [ecx+04]
'00711a30    8b16                    mov edx, dword ptr [esi]
'00711a32    33db                    xor ebx, ebx
'00711a34    56                      push esi
'00711a35    895de8                  mov dword ptr [ebp-18], ebx
'00711a38    895de4                  mov dword ptr [ebp-1c], ebx
'00711a3b    895de0                  mov dword ptr [ebp-20], ebx
'00711a3e    ff9208030000            call dword ptr [edx+00000308]
'00711a44    50                      push eax
'00711a45    8d45e8                  lea eax, dword ptr [ebp-18]
'00711a48    50                      push eax

' *** Reference to "__vbaObjSet"
'00711a49    ff15b4104000            call dword ptr [004010b4]
Set var_41 = Me
'00711a4f    8d55e4                  lea edx, dword ptr [ebp-1c]
'00711a52    8bf8                    mov edi, eax
'00711a54    8b0f                    mov ecx, dword ptr [edi]
'00711a56    52                      push edx
'00711a57    57                      push edi
'00711a58    ff91f0000000            call dword ptr [ecx+000000f0]
'00711a5e    dbe2                    fnclex
'00711a60    3bc3                    cmp eax, ebx
'00711a62    7d16                    jge 711a7a

If (var_41 < 0) Then

' *** Reference to "__vbaHresultCheckObj"
'00711a64    8b1d80104000            mov ebx, dword ptr [00401080]
'00711a6a    68f0000000              push 000000f0
'00711a6f    681cb94300              push 0043b91c
'00711a74    57                      push edi
'00711a75    50                      push eax
'00711a76    ffd3                    call ebx
'00711a78    eb06                    jmp 711a80
    
End If

' *** Reference to "__vbaHresultCheckObj"
'00711a7a    8b1d80104000            mov ebx, dword ptr [00401080]
'00711a80    33c0                    xor eax, eax
var_num1 = Empty
'00711a82    663945e4                cmp word ptr [ebp-1c], ax
'00711a86    8d4de8                  lea ecx, dword ptr [ebp-18]
'00711a89    0f94c0                  sete al
'00711a8c    f7d8                    neg eax
'00711a8e    668bf8                  mov di, ax

' *** Reference to "__vbaFreeObj"
'00711a91    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_41)
'00711a97    6685ff                  test edi, edi
'00711a9a    745b                    je 711af7

If (0 = var_41) Then
'00711a9c    8b0e                    mov ecx, dword ptr [esi]
'00711a9e    56                      push esi
'00711a9f    ff9108030000            call dword ptr [ecx+00000308]
'00711aa5    50                      push eax
'00711aa6    8d55e8                  lea edx, dword ptr [ebp-18]
'00711aa9    52                      push edx

' *** Reference to "__vbaObjSet"
'00711aaa    ff15b4104000            call dword ptr [004010b4]
    Set var_41 = 0 = var_41
'00711ab0    8d4de4                  lea ecx, dword ptr [ebp-1c]
'00711ab3    8bf8                    mov edi, eax
'00711ab5    8b07                    mov eax, dword ptr [edi]
'00711ab7    51                      push ecx
'00711ab8    57                      push edi
'00711ab9    ff90e8000000            call dword ptr [eax+000000e8]
'00711abf    dbe2                    fnclex
'00711ac1    85c0                    test ax, ax
'00711ac3    7d0e                    jge 711ad3
'00711ac5    68e8000000              push 000000e8
'00711aca    681cb94300              push 0043b91c
'00711acf    57                      push edi
'00711ad0    50                      push eax
'00711ad1    ffd3                    call ebx
'00711ad3    668b55e4                mov dx, word ptr [ebp-1c]
'00711ad7    8b06                    mov eax, dword ptr [esi]
'00711ad9    6683ea01                sub dx, 01
    var_num4 = 0 - 1
'00711add    0f80b8000000            jo 711b9b
'00711ae3    8d4de0                  lea ecx, dword ptr [ebp-20]
'00711ae6    51                      push ecx
'00711ae7    56                      push esi
'00711ae8    8955e0                  mov dword ptr [ebp-20], edx
'00711aeb    ff90f8060000            call dword ptr [eax+000006f8]
'00711af1    85c0                    test ax, ax
'00711af3    7d65                    jge 711b5a
'00711af5    eb55                    jmp 711b4c
    
End If
'00711af7    8b16                    mov edx, dword ptr [esi]
'00711af9    56                      push esi
'00711afa    ff9208030000            call dword ptr [edx+00000308]
'00711b00    50                      push eax
'00711b01    8d45e8                  lea eax, dword ptr [ebp-18]
'00711b04    50                      push eax

' *** Reference to "__vbaObjSet"
'00711b05    ff15b4104000            call dword ptr [004010b4]
Set var_41 = 0 = var_41
'00711b0b    8d55e4                  lea edx, dword ptr [ebp-1c]
'00711b0e    8bf8                    mov edi, eax
'00711b10    8b0f                    mov ecx, dword ptr [edi]
'00711b12    52                      push edx
'00711b13    57                      push edi
'00711b14    ff91f0000000            call dword ptr [ecx+000000f0]
'00711b1a    dbe2                    fnclex
'00711b1c    85c0                    test ax, ax
'00711b1e    7d0e                    jge 711b2e
'00711b20    68f0000000              push 000000f0
'00711b25    681cb94300              push 0043b91c
'00711b2a    57                      push edi
'00711b2b    50                      push eax
'00711b2c    ffd3                    call ebx
'00711b2e    668b45e4                mov ax, word ptr [ebp-1c]
'00711b32    8b0e                    mov ecx, dword ptr [esi]
'00711b34    662d0100                sub ax, 0001
var_num1 = 0 - 1
'00711b38    7061                    jo 711b9b
'00711b3a    8d55e0                  lea edx, dword ptr [ebp-20]
'00711b3d    52                      push edx
'00711b3e    56                      push esi
'00711b3f    8945e0                  mov dword ptr [ebp-20], eax
'00711b42    ff91f8060000            call dword ptr [ecx+000006f8]
'00711b48    85c0                    test ax, ax
'00711b4a    7d0e                    jge 711b5a
'00711b4c    68f8060000              push 000006f8
'00711b51    68e4194300              push 004319e4
'00711b56    56                      push esi
'00711b57    50                      push eax
'00711b58    ffd3                    call ebx
'00711b5a    8d4de8                  lea ecx, dword ptr [ebp-18]

' *** Reference to "__vbaFreeObj"
'00711b5d    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_41)
'00711b63    c745fc00000000          mov dword ptr [ebp-04], 00000000
'00711b6a    687c1b7100              push 00711b7c
'00711b6f    eb0a                    jmp 711b7b
'00711b71    8d4de8                  lea ecx, dword ptr [ebp-18]

' *** Reference to "__vbaFreeObj"
'00711b74    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_41)
'00711b7a    c3                      ret
'00711b7b    c3                      ret
'00711b7c    8b4508                  mov eax, dword ptr [ebp+08]
'00711b7f    8b08                    mov ecx, dword ptr [eax]
'00711b81    50                      push eax
'00711b82    ff5108                  call dword ptr [ecx+08]
'00711b85    8b45fc                  mov eax, dword ptr [ebp-04]
'00711b88    8b4dec                  mov ecx, dword ptr [ebp-14]
'00711b8b    5f                      pop edi
'00711b8c    5e                      pop esi
    'Reference to '__except_list'
'00711b8d    64890d00000000          mov dword ptr fs:[00000000], ecx
'00711b94    5b                      pop ebx
'00711b95    8be5                    mov esp, ebp
'00711b97    5d                      pop ebp
'00711b98    c20400                  ret 0004


End Sub


'Event for ChkPouvoir
Private Sub ChkPouvoir_Click()
'00711d90    55                      push ebp
'00711d91    8bec                    mov ebp, esp
'00711d93    83ec0c                  sub esp, 0c

' *** Reference to "__vbaExceptHandler"
'00711d96    6866724000              push 00407266
'00711d9b    64a100000000            mov ax, word ptr fs:[00000000]
'00711da1    50                      push eax
    'Reference to '__except_list'
'00711da2    64892500000000          mov dword ptr fs:[00000000], esp
'00711da9    83ec24                  sub esp, 24
'00711dac    53                      push ebx
'00711dad    56                      push esi
'00711dae    57                      push edi
'00711daf    8965f4                  mov dword ptr [ebp-0c], esp
'00711db2    c745f8606d4000          mov dword ptr [ebp-08], 00406d60
'00711db9    8b7508                  mov esi, dword ptr [ebp+08]
'00711dbc    8bc6                    mov eax, esi
'00711dbe    83e001                  and eax, 01
'00711dc1    8945fc                  mov dword ptr [ebp-04], eax
'00711dc4    83e6fe                  and esi, fffffffe
'00711dc7    8b0e                    mov ecx, dword ptr [esi]
'00711dc9    56                      push esi
'00711dca    897508                  mov dword ptr [ebp+08], esi
'00711dcd    ff5104                  call dword ptr [ecx+04]
'00711dd0    8b16                    mov edx, dword ptr [esi]
'00711dd2    33db                    xor ebx, ebx
'00711dd4    56                      push esi
'00711dd5    895de8                  mov dword ptr [ebp-18], ebx
'00711dd8    895dd8                  mov dword ptr [ebp-28], ebx
'00711ddb    ff9248030000            call dword ptr [edx+00000348]
'00711de1    8945e0                  mov dword ptr [ebp-20], eax
'00711de4    8d45d8                  lea eax, dword ptr [ebp-28]
'00711de7    50                      push eax
'00711de8    c745d809000000          mov dword ptr [ebp-28], 00000009

' *** Reference to "__vbaBoolVarNull"
'00711def    ff15fc104000            call dword ptr [004010fc]
'00711df5    8d4dd8                  lea ecx, dword ptr [ebp-28]
'00711df8    668bf8                  mov di, ax

' *** Reference to "__vbaFreeVar"
'00711dfb    ff1528104000            call dword ptr [00401028]
'#Cleanup(var_45)
'00711e01    663bfb                  cmp di, bx
'00711e04    8b0e                    mov ecx, dword ptr [esi]
'00711e06    56                      push esi
'00711e07    7441                    je 711e4a

If (CBool(Me) <> 0) Then
'00711e09    ff91fc020000            call dword ptr [ecx+000002fc]
'00711e0f    50                      push eax
'00711e10    8d55e8                  lea edx, dword ptr [ebp-18]
'00711e13    52                      push edx

' *** Reference to "__vbaObjSet"
'00711e14    ff15b4104000            call dword ptr [004010b4]
    Set var_41 = CBool(Me)
'00711e1a    8bf0                    mov esi, eax
'00711e1c    8b06                    mov eax, dword ptr [esi]
'00711e1e    6aff                    push ffffffff
'00711e20    56                      push esi
'00711e21    ff9094000000            call dword ptr [eax+00000094]
'00711e27    dbe2                    fnclex
'00711e29    3bc3                    cmp eax, ebx
'00711e2b    7d12                    jge 711e3f
    
    If (    0 < 0) Then
'00711e2d    6894000000              push 00000094
'00711e32    681cb94300              push 0043b91c
'00711e37    56                      push esi
'00711e38    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00711e39    ff1580104000            call dword ptr [00401080]
    
End If
'00711e3f    8d4de8                  lea ecx, dword ptr [ebp-18]

' *** Reference to "__vbaFreeObj"
'00711e42    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_41)
'00711e48    eb7f                    jmp 711ec9

'ERROR: Two many next close:
End If
'00711e4a    ff91fc020000            call dword ptr [ecx+000002fc]

' *** Reference to "__vbaObjSet"
'00711e50    8b1db4104000            mov ebx, dword ptr [004010b4]
'00711e56    50                      push eax
'00711e57    8d55e8                  lea edx, dword ptr [ebp-18]
'00711e5a    52                      push edx
'00711e5b    ffd3                    call ebx
Set var_41 = CBool(Me)
'00711e5d    8bf8                    mov edi, eax
'00711e5f    8b07                    mov eax, dword ptr [edi]
'00711e61    6aff                    push ffffffff
'00711e63    57                      push edi
'00711e64    ff90f4000000            call dword ptr [eax+000000f4]
'00711e6a    dbe2                    fnclex
'00711e6c    85c0                    test ax, ax
'00711e6e    7d12                    jge 711e82
'00711e70    68f4000000              push 000000f4
'00711e75    681cb94300              push 0043b91c
'00711e7a    57                      push edi
'00711e7b    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00711e7c    ff1580104000            call dword ptr [00401080]

' *** Reference to "__vbaFreeObj"
'00711e82    8b3d08134000            mov edi, dword ptr [00401308]
'00711e88    8d4de8                  lea ecx, dword ptr [ebp-18]
'00711e8b    ffd7                    call edi
'#Cleanup(var_41)
'00711e8d    8b0e                    mov ecx, dword ptr [esi]
'00711e8f    56                      push esi
'00711e90    ff91fc020000            call dword ptr [ecx+000002fc]
'00711e96    50                      push eax
'00711e97    8d55e8                  lea edx, dword ptr [ebp-18]
'00711e9a    52                      push edx
'00711e9b    ffd3                    call ebx
Set var_41 = Nothing
'00711e9d    8bf0                    mov esi, eax
'00711e9f    8b06                    mov eax, dword ptr [esi]
'00711ea1    6a00                    push 00
'00711ea3    56                      push esi
'00711ea4    ff9094000000            call dword ptr [eax+00000094]
'00711eaa    dbe2                    fnclex
'00711eac    85c0                    test ax, ax
'00711eae    7d12                    jge 711ec2
'00711eb0    6894000000              push 00000094
'00711eb5    681cb94300              push 0043b91c
'00711eba    56                      push esi
'00711ebb    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00711ebc    ff1580104000            call dword ptr [00401080]
'00711ec2    8d4de8                  lea ecx, dword ptr [ebp-18]
'00711ec5    ffd7                    call edi
'#Cleanup(var_41)
'00711ec7    33db                    xor ebx, ebx
var_num2 = Empty
'00711ec9    895dfc                  mov dword ptr [ebp-04], ebx
'00711ecc    68e71e7100              push 00711ee7
'00711ed1    eb13                    jmp 711ee6
'00711ed3    8d4de8                  lea ecx, dword ptr [ebp-18]

' *** Reference to "__vbaFreeObj"
'00711ed6    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_41)
'00711edc    8d4dd8                  lea ecx, dword ptr [ebp-28]

' *** Reference to "__vbaFreeVar"
'00711edf    ff1528104000            call dword ptr [00401028]
'#Cleanup(var_45)
'00711ee5    c3                      ret
'00711ee6    c3                      ret
'00711ee7    8b4508                  mov eax, dword ptr [ebp+08]
'00711eea    8b08                    mov ecx, dword ptr [eax]
'00711eec    50                      push eax
'00711eed    ff5108                  call dword ptr [ecx+08]
'00711ef0    8b45fc                  mov eax, dword ptr [ebp-04]
'00711ef3    8b4dec                  mov ecx, dword ptr [ebp-14]
'00711ef6    5f                      pop edi
'00711ef7    5e                      pop esi
    'Reference to '__except_list'
'00711ef8    64890d00000000          mov dword ptr fs:[00000000], ecx
'00711eff    5b                      pop ebx
'00711f00    8be5                    mov esp, ebp
'00711f02    5d                      pop ebp
'00711f03    c20400                  ret 0004


End Sub


'Event for btnNouveau
Private Sub btnNouveau_Click()
'00711960    55                      push ebp
'00711961    8bec                    mov ebp, esp
'00711963    83ec0c                  sub esp, 0c

' *** Reference to "__vbaExceptHandler"
'00711966    6866724000              push 00407266
'0071196b    64a100000000            mov ax, word ptr fs:[00000000]
'00711971    50                      push eax
    'Reference to '__except_list'
'00711972    64892500000000          mov dword ptr fs:[00000000], esp
'00711979    83ec0c                  sub esp, 0c
'0071197c    53                      push ebx
'0071197d    56                      push esi
'0071197e    57                      push edi
'0071197f    8965f4                  mov dword ptr [ebp-0c], esp
'00711982    c745f8386d4000          mov dword ptr [ebp-08], 00406d38
'00711989    8b7508                  mov esi, dword ptr [ebp+08]
'0071198c    8bc6                    mov eax, esi
'0071198e    83e001                  and eax, 01
'00711991    8945fc                  mov dword ptr [ebp-04], eax
'00711994    83e6fe                  and esi, fffffffe
'00711997    8b0e                    mov ecx, dword ptr [esi]
'00711999    56                      push esi
'0071199a    897508                  mov dword ptr [ebp+08], esi
'0071199d    ff5104                  call dword ptr [ecx+04]
'007119a0    8b16                    mov edx, dword ptr [esi]
'007119a2    56                      push esi
'007119a3    ff92fc060000            call dword ptr [edx+000006fc]
'007119a9    85c0                    test ax, ax
'007119ab    7d12                    jge 7119bf
'007119ad    68fc060000              push 000006fc
'007119b2    68e4194300              push 004319e4
'007119b7    56                      push esi
'007119b8    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'007119b9    ff1580104000            call dword ptr [00401080]
'007119bf    c745fc00000000          mov dword ptr [ebp-04], 00000000
'007119c6    8b4508                  mov eax, dword ptr [ebp+08]
'007119c9    8b08                    mov ecx, dword ptr [eax]
'007119cb    50                      push eax
'007119cc    ff5108                  call dword ptr [ecx+08]
'007119cf    8b45fc                  mov eax, dword ptr [ebp-04]
'007119d2    8b4dec                  mov ecx, dword ptr [ebp-14]
'007119d5    5f                      pop edi
'007119d6    5e                      pop esi
    'Reference to '__except_list'
'007119d7    64890d00000000          mov dword ptr fs:[00000000], ecx
'007119de    5b                      pop ebx
'007119df    8be5                    mov esp, ebp
'007119e1    5d                      pop ebp
'007119e2    c20400                  ret 0004


End Sub


'Event for btnEnregistrer
Private Sub btnEnregistrer_Click()
'0070fb30    55                      push ebp
'0070fb31    8bec                    mov ebp, esp
'0070fb33    83ec0c                  sub esp, 0c

' *** Reference to "__vbaExceptHandler"
'0070fb36    6866724000              push 00407266
'0070fb3b    64a100000000            mov ax, word ptr fs:[00000000]
'0070fb41    50                      push eax
    'Reference to '__except_list'
'0070fb42    64892500000000          mov dword ptr fs:[00000000], esp
'0070fb49    81ec9c010000            sub esp, 0000019c
'0070fb4f    53                      push ebx
'0070fb50    56                      push esi
'0070fb51    57                      push edi
'0070fb52    8965f4                  mov dword ptr [ebp-0c], esp
'0070fb55    c745f8186d4000          mov dword ptr [ebp-08], 00406d18
'0070fb5c    8b7508                  mov esi, dword ptr [ebp+08]
'0070fb5f    8bc6                    mov eax, esi
'0070fb61    83e001                  and eax, 01
'0070fb64    8945fc                  mov dword ptr [ebp-04], eax
'0070fb67    83e6fe                  and esi, fffffffe
'0070fb6a    8b0e                    mov ecx, dword ptr [esi]
'0070fb6c    56                      push esi
'0070fb6d    897508                  mov dword ptr [ebp+08], esi
'0070fb70    ff5104                  call dword ptr [ecx+04]
'0070fb73    8b16                    mov edx, dword ptr [esi]
'0070fb75    33ff                    xor edi, edi
'0070fb77    56                      push esi
'0070fb78    897de8                  mov dword ptr [ebp-18], edi
'0070fb7b    897de4                  mov dword ptr [ebp-1c], edi
'0070fb7e    897de0                  mov dword ptr [ebp-20], edi
'0070fb81    897ddc                  mov dword ptr [ebp-24], edi
'0070fb84    897dd4                  mov dword ptr [ebp-2c], edi
'0070fb87    897dd0                  mov dword ptr [ebp-30], edi
'0070fb8a    897dcc                  mov dword ptr [ebp-34], edi
'0070fb8d    897dc8                  mov dword ptr [ebp-38], edi
'0070fb90    897dc4                  mov dword ptr [ebp-3c], edi
'0070fb93    897dc0                  mov dword ptr [ebp-40], edi
'0070fb96    897db0                  mov dword ptr [ebp-50], edi
'0070fb99    897da0                  mov dword ptr [ebp-60], edi
'0070fb9c    897d90                  mov dword ptr [ebp-70], edi
'0070fb9f    897d80                  mov dword ptr [ebp-80], edi
'0070fba2    89bd70ffffff            mov dword ptr [ebp+ffffff70], edi
'0070fba8    89bd60ffffff            mov dword ptr [ebp+ffffff60], edi
'0070fbae    89bd50ffffff            mov dword ptr [ebp+ffffff50], edi
'0070fbb4    89bd40ffffff            mov dword ptr [ebp+ffffff40], edi
'0070fbba    89bd30ffffff            mov dword ptr [ebp+ffffff30], edi
'0070fbc0    89bd20ffffff            mov dword ptr [ebp+ffffff20], edi
'0070fbc6    89bd10ffffff            mov dword ptr [ebp+ffffff10], edi
'0070fbcc    89bd00ffffff            mov dword ptr [ebp+ffffff00], edi
'0070fbd2    89bdf0feffff            mov dword ptr [ebp+fffffef0], edi
'0070fbd8    89bde0feffff            mov dword ptr [ebp+fffffee0], edi
'0070fbde    89bdd0feffff            mov dword ptr [ebp+fffffed0], edi
'0070fbe4    89bdc0feffff            mov dword ptr [ebp+fffffec0], edi
'0070fbea    89bdb0feffff            mov dword ptr [ebp+fffffeb0], edi
'0070fbf0    89bda0feffff            mov dword ptr [ebp+fffffea0], edi
'0070fbf6    89bd9cfeffff            mov dword ptr [ebp+fffffe9c], edi
'0070fbfc    89bd78feffff            mov dword ptr [ebp+fffffe78], edi
'0070fc02    89bd74feffff            mov dword ptr [ebp+fffffe74], edi
'0070fc08    ff9208030000            call dword ptr [edx+00000308]
'0070fc0e    8945b8                  mov dword ptr [ebp-48], eax
'0070fc11    8b06                    mov eax, dword ptr [esi]
'0070fc13    56                      push esi
'0070fc14    c745b009000000          mov dword ptr [ebp-50], 00000009
'0070fc1b    ff9008030000            call dword ptr [eax+00000308]

' *** Reference to "__vbaObjSet"
'0070fc21    8b1db4104000            mov ebx, dword ptr [004010b4]
'0070fc27    50                      push eax
'0070fc28    8d4dc8                  lea ecx, dword ptr [ebp-38]
'0070fc2b    51                      push ecx
'0070fc2c    ffd3                    call ebx
Set var_46 = Nothing
'0070fc2e    8b10                    mov edx, dword ptr [eax]
'0070fc30    8d4dd4                  lea ecx, dword ptr [ebp-2c]
'0070fc33    51                      push ecx
'0070fc34    50                      push eax
'0070fc35    898598feffff            mov dword ptr [ebp+fffffe98], eax
'0070fc3b    ff92a8000000            call dword ptr [edx+000000a8]
'0070fc41    dbe2                    fnclex
'0070fc43    3bc7                    cmp eax, edi
'0070fc45    7d18                    jge 70fc5f

If (var_46 < 0) Then
'0070fc47    8b9598feffff            mov edx, dword ptr [ebp+fffffe98]
'0070fc4d    68a8000000              push 000000a8
'0070fc52    681cb94300              push 0043b91c
'0070fc57    52                      push edx
'0070fc58    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0070fc59    ff1580104000            call dword ptr [00401080]
    
End If
'0070fc5f    8b45d4                  mov eax, dword ptr [ebp-2c]
'0070fc62    50                      push eax
'0070fc63    68cc134300              push 004313cc

' *** Reference to "__vbaStrCmp"
'0070fc68    ff1530114000            call dword ptr [00401130]
'0070fc6e    8bf8                    mov edi, eax
'0070fc70    f7df                    neg edi
'0070fc72    1bff                    sbb edi, edi
'0070fc74    8d4db0                  lea ecx, dword ptr [ebp-50]
'0070fc77    47                      inc edi
'0070fc78    51                      push ecx
'0070fc79    f7df                    neg edi

' *** Reference to "rtcIsNull"
'0070fc7b    ff1540114000            call dword ptr [00401140]
'0070fc81    8d4dd4                  lea ecx, dword ptr [ebp-2c]
'0070fc84    0bf8                    or edi, eax
var_num7 = ((vbNullString) = (vbNullChar)) Or IsNull(Me)

' *** Reference to "__vbaFreeStr"
'0070fc86    ff150c134000            call dword ptr [0040130c]
'#Cleanup(var_44)
'0070fc8c    8d4dc8                  lea ecx, dword ptr [ebp-38]

' *** Reference to "__vbaFreeObj"
'0070fc8f    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_46)
'0070fc95    8d4db0                  lea ecx, dword ptr [ebp-50]

' *** Reference to "__vbaFreeVar"
'0070fc98    ff1528104000            call dword ptr [00401028]
'#Cleanup(var_6)
'0070fc9e    6685ff                  test edi, edi
'0070fca1    0f849c000000            je 70fd43

If (var_num7) Then
'0070fca7    b904000280              mov ecx, 80020004
'0070fcac    b80a000000              mov eax, 0000000a
'0070fcb1    894d88                  mov dword ptr [ebp-78], ecx
'0070fcb4    894d98                  mov dword ptr [ebp-68], ecx
'0070fcb7    894da8                  mov dword ptr [ebp-58], ecx
'0070fcba    8d9550ffffff            lea edx, dword ptr [ebp+ffffff50]
'0070fcc0    8d4db0                  lea ecx, dword ptr [ebp-50]
'0070fcc3    894580                  mov dword ptr [ebp-80], eax
'0070fcc6    894590                  mov dword ptr [ebp-70], eax
'0070fcc9    8945a0                  mov dword ptr [ebp-60], eax
'0070fccc    c78558ffffff8c674500    mov dword ptr [ebp+ffffff58], 0045678c
'0070fcd6    c78550ffffff08000000    mov dword ptr [ebp+ffffff50], 00000008

' *** Reference to "__vbaVarDup"
'0070fce0    ff1598124000            call dword ptr [00401298]
    var_6 = ("Il faut que le don ait un nom")
'0070fce6    8d5580                  lea edx, dword ptr [ebp-80]
'0070fce9    52                      push edx
'0070fcea    8d4590                  lea eax, dword ptr [ebp-70]
'0070fced    50                      push eax
'0070fcee    8d4da0                  lea ecx, dword ptr [ebp-60]
'0070fcf1    51                      push ecx
'0070fcf2    6a00                    push 00
'0070fcf4    8d55b0                  lea edx, dword ptr [ebp-50]
'0070fcf7    52                      push edx

' *** Reference to "rtcMsgBox"
'0070fcf8    ff15b8104000            call dword ptr [004010b8]
    var_76 = MsgBox(var_6, 0)
'0070fcfe    8d4580                  lea eax, dword ptr [ebp-80]
'0070fd01    50                      push eax
'0070fd02    8d4d90                  lea ecx, dword ptr [ebp-70]
'0070fd05    51                      push ecx
'0070fd06    8d55a0                  lea edx, dword ptr [ebp-60]
'0070fd09    52                      push edx
'0070fd0a    8d45b0                  lea eax, dword ptr [ebp-50]
'0070fd0d    50                      push eax
'0070fd0e    6a04                    push 04

' *** Reference to "__vbaFreeVarList"
'0070fd10    ff1540104000            call dword ptr [00401040]
    '#Cleanup( var_6, var_7, var_8, var_18)
'0070fd16    8b0e                    mov ecx, dword ptr [esi]
'0070fd18    83c414                  add esp, 14
'0070fd1b    56                      push esi
'0070fd1c    ff9108030000            call dword ptr [ecx+00000308]
'0070fd22    50                      push eax
'0070fd23    8d55c8                  lea edx, dword ptr [ebp-38]
'0070fd26    52                      push edx
'0070fd27    ffd3                    call ebx
    Set var_46 = 
'0070fd29    8bf0                    mov esi, eax
'0070fd2b    8b06                    mov eax, dword ptr [esi]
'0070fd2d    56                      push esi
'0070fd2e    ff90f4010000            call dword ptr [eax+000001f4]
'0070fd34    dbe2                    fnclex
'0070fd36    85c0                    test ax, ax
'0070fd38    0f8d6a020000            jge 70ffa8
'0070fd3e    e953020000              jmp 70ff96
    
End If
'0070fd43    8b0e                    mov ecx, dword ptr [esi]
'0070fd45    56                      push esi
'0070fd46    ff9100030000            call dword ptr [ecx+00000300]
'0070fd4c    8b16                    mov edx, dword ptr [esi]
'0070fd4e    56                      push esi
'0070fd4f    8945b8                  mov dword ptr [ebp-48], eax
'0070fd52    c745b009000000          mov dword ptr [ebp-50], 00000009
'0070fd59    ff9200030000            call dword ptr [edx+00000300]
'0070fd5f    50                      push eax
'0070fd60    8d45c8                  lea eax, dword ptr [ebp-38]
'0070fd63    50                      push eax
'0070fd64    ffd3                    call ebx
Set var_46 = IsNull(Me)
'0070fd66    8d55d4                  lea edx, dword ptr [ebp-2c]
'0070fd69    8bf8                    mov edi, eax
'0070fd6b    8b0f                    mov ecx, dword ptr [edi]
'0070fd6d    52                      push edx
'0070fd6e    57                      push edi
'0070fd6f    ff91a8000000            call dword ptr [ecx+000000a8]
'0070fd75    dbe2                    fnclex
'0070fd77    85c0                    test ax, ax
'0070fd79    7d12                    jge 70fd8d
'0070fd7b    68a8000000              push 000000a8
'0070fd80    681cb94300              push 0043b91c
'0070fd85    57                      push edi
'0070fd86    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0070fd87    ff1580104000            call dword ptr [00401080]
'0070fd8d    8b45d4                  mov eax, dword ptr [ebp-2c]
'0070fd90    50                      push eax
'0070fd91    68cc134300              push 004313cc

' *** Reference to "__vbaStrCmp"
'0070fd96    ff1530114000            call dword ptr [00401130]
'0070fd9c    8bf8                    mov edi, eax
'0070fd9e    f7df                    neg edi
'0070fda0    1bff                    sbb edi, edi
'0070fda2    8d4db0                  lea ecx, dword ptr [ebp-50]
'0070fda5    47                      inc edi
'0070fda6    51                      push ecx
'0070fda7    f7df                    neg edi

' *** Reference to "rtcIsNull"
'0070fda9    ff1540114000            call dword ptr [00401140]
'0070fdaf    8d4dd4                  lea ecx, dword ptr [ebp-2c]
'0070fdb2    0bf8                    or edi, eax
var_num7 = ((vbNullString) = (vbNullChar)) Or IsNull(IsNull(Me))

' *** Reference to "__vbaFreeStr"
'0070fdb4    ff150c134000            call dword ptr [0040130c]
'#Cleanup(var_44)
'0070fdba    8d4dc8                  lea ecx, dword ptr [ebp-38]

' *** Reference to "__vbaFreeObj"
'0070fdbd    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_46)
'0070fdc3    8d4db0                  lea ecx, dword ptr [ebp-50]

' *** Reference to "__vbaFreeVar"
'0070fdc6    ff1528104000            call dword ptr [00401028]
'#Cleanup(var_6)
'0070fdcc    6685ff                  test edi, edi
'0070fdcf    0f849c000000            je 70fe71

If (var_num7) Then
'0070fdd5    b904000280              mov ecx, 80020004
'0070fdda    b80a000000              mov eax, 0000000a
'0070fddf    894d88                  mov dword ptr [ebp-78], ecx
'0070fde2    894d98                  mov dword ptr [ebp-68], ecx
'0070fde5    894da8                  mov dword ptr [ebp-58], ecx
'0070fde8    8d9550ffffff            lea edx, dword ptr [ebp+ffffff50]
'0070fdee    8d4db0                  lea ecx, dword ptr [ebp-50]
'0070fdf1    894580                  mov dword ptr [ebp-80], eax
'0070fdf4    894590                  mov dword ptr [ebp-70], eax
'0070fdf7    8945a0                  mov dword ptr [ebp-60], eax
'0070fdfa    c78558ffffffcc674500    mov dword ptr [ebp+ffffff58], 004567cc
'0070fe04    c78550ffffff08000000    mov dword ptr [ebp+ffffff50], 00000008

' *** Reference to "__vbaVarDup"
'0070fe0e    ff1598124000            call dword ptr [00401298]
    var_6 = ("Il faut que le don ait un genre")
'0070fe14    8d5580                  lea edx, dword ptr [ebp-80]
'0070fe17    52                      push edx
'0070fe18    8d4590                  lea eax, dword ptr [ebp-70]
'0070fe1b    50                      push eax
'0070fe1c    8d4da0                  lea ecx, dword ptr [ebp-60]
'0070fe1f    51                      push ecx
'0070fe20    6a00                    push 00
'0070fe22    8d55b0                  lea edx, dword ptr [ebp-50]
'0070fe25    52                      push edx

' *** Reference to "rtcMsgBox"
'0070fe26    ff15b8104000            call dword ptr [004010b8]
    var_15 = MsgBox(var_6, 0)
'0070fe2c    8d4580                  lea eax, dword ptr [ebp-80]
'0070fe2f    50                      push eax
'0070fe30    8d4d90                  lea ecx, dword ptr [ebp-70]
'0070fe33    51                      push ecx
'0070fe34    8d55a0                  lea edx, dword ptr [ebp-60]
'0070fe37    52                      push edx
'0070fe38    8d45b0                  lea eax, dword ptr [ebp-50]
'0070fe3b    50                      push eax
'0070fe3c    6a04                    push 04

' *** Reference to "__vbaFreeVarList"
'0070fe3e    ff1540104000            call dword ptr [00401040]
    '#Cleanup( var_6, var_7, var_8, var_18)
'0070fe44    8b0e                    mov ecx, dword ptr [esi]
'0070fe46    83c414                  add esp, 14
'0070fe49    56                      push esi
'0070fe4a    ff9100030000            call dword ptr [ecx+00000300]
'0070fe50    50                      push eax
'0070fe51    8d55c8                  lea edx, dword ptr [ebp-38]
'0070fe54    52                      push edx
'0070fe55    ffd3                    call ebx
    Set var_46 = 
'0070fe57    8bf0                    mov esi, eax
'0070fe59    8b06                    mov eax, dword ptr [esi]
'0070fe5b    56                      push esi
'0070fe5c    ff90f4010000            call dword ptr [eax+000001f4]
'0070fe62    dbe2                    fnclex
'0070fe64    85c0                    test ax, ax
'0070fe66    0f8d3c010000            jge 70ffa8
'0070fe6c    e925010000              jmp 70ff96
    
End If
'0070fe71    8b0e                    mov ecx, dword ptr [esi]
'0070fe73    56                      push esi
'0070fe74    ff9104030000            call dword ptr [ecx+00000304]
'0070fe7a    8b16                    mov edx, dword ptr [esi]
'0070fe7c    56                      push esi
'0070fe7d    8945b8                  mov dword ptr [ebp-48], eax
'0070fe80    c745b009000000          mov dword ptr [ebp-50], 00000009
'0070fe87    ff9204030000            call dword ptr [edx+00000304]
'0070fe8d    50                      push eax
'0070fe8e    8d45c8                  lea eax, dword ptr [ebp-38]
'0070fe91    50                      push eax
'0070fe92    ffd3                    call ebx
Set var_46 = IsNull(IsNull(Me))
'0070fe94    8d55d4                  lea edx, dword ptr [ebp-2c]
'0070fe97    8bf8                    mov edi, eax
'0070fe99    8b0f                    mov ecx, dword ptr [edi]
'0070fe9b    52                      push edx
'0070fe9c    57                      push edi
'0070fe9d    ff91a8000000            call dword ptr [ecx+000000a8]
'0070fea3    dbe2                    fnclex
'0070fea5    85c0                    test ax, ax
'0070fea7    7d12                    jge 70febb
'0070fea9    68a8000000              push 000000a8
'0070feae    681cb94300              push 0043b91c
'0070feb3    57                      push edi
'0070feb4    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0070feb5    ff1580104000            call dword ptr [00401080]
'0070febb    8b45d4                  mov eax, dword ptr [ebp-2c]
'0070febe    50                      push eax
'0070febf    68cc134300              push 004313cc

' *** Reference to "__vbaStrCmp"
'0070fec4    ff1530114000            call dword ptr [00401130]
'0070feca    8bf8                    mov edi, eax
'0070fecc    f7df                    neg edi
'0070fece    1bff                    sbb edi, edi
'0070fed0    8d4db0                  lea ecx, dword ptr [ebp-50]
'0070fed3    47                      inc edi
'0070fed4    51                      push ecx
'0070fed5    f7df                    neg edi

' *** Reference to "rtcIsNull"
'0070fed7    ff1540114000            call dword ptr [00401140]
'0070fedd    8d4dd4                  lea ecx, dword ptr [ebp-2c]
'0070fee0    0bf8                    or edi, eax
var_num7 = ((vbNullString) = (vbNullChar)) Or IsNull(IsNull(IsNull(Me)))

' *** Reference to "__vbaFreeStr"
'0070fee2    ff150c134000            call dword ptr [0040130c]
'#Cleanup(var_44)
'0070fee8    8d4dc8                  lea ecx, dword ptr [ebp-38]

' *** Reference to "__vbaFreeObj"
'0070feeb    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_46)
'0070fef1    8d4db0                  lea ecx, dword ptr [ebp-50]

' *** Reference to "__vbaFreeVar"
'0070fef4    ff1528104000            call dword ptr [00401028]
'#Cleanup(var_6)
'0070fefa    6685ff                  test edi, edi
'0070fefd    0f84b3000000            je 70ffb6

If (var_num7) Then
'0070ff03    b904000280              mov ecx, 80020004
'0070ff08    b80a000000              mov eax, 0000000a
'0070ff0d    894d88                  mov dword ptr [ebp-78], ecx
'0070ff10    894d98                  mov dword ptr [ebp-68], ecx
'0070ff13    894da8                  mov dword ptr [ebp-58], ecx
'0070ff16    8d9550ffffff            lea edx, dword ptr [ebp+ffffff50]
'0070ff1c    8d4db0                  lea ecx, dword ptr [ebp-50]
'0070ff1f    894580                  mov dword ptr [ebp-80], eax
'0070ff22    894590                  mov dword ptr [ebp-70], eax
'0070ff25    8945a0                  mov dword ptr [ebp-60], eax
'0070ff28    c78558ffffff48684500    mov dword ptr [ebp+ffffff58], 00456848
'0070ff32    c78550ffffff08000000    mov dword ptr [ebp+ffffff50], 00000008

' *** Reference to "__vbaVarDup"
'0070ff3c    ff1598124000            call dword ptr [00401298]
    var_6 = ("Il faut que le don ait une source")
'0070ff42    8d5580                  lea edx, dword ptr [ebp-80]
'0070ff45    52                      push edx
'0070ff46    8d4590                  lea eax, dword ptr [ebp-70]
'0070ff49    50                      push eax
'0070ff4a    8d4da0                  lea ecx, dword ptr [ebp-60]
'0070ff4d    51                      push ecx
'0070ff4e    6a00                    push 00
'0070ff50    8d55b0                  lea edx, dword ptr [ebp-50]
'0070ff53    52                      push edx

' *** Reference to "rtcMsgBox"
'0070ff54    ff15b8104000            call dword ptr [004010b8]
    var_53 = MsgBox(var_6, 0)
'0070ff5a    8d4580                  lea eax, dword ptr [ebp-80]
'0070ff5d    50                      push eax
'0070ff5e    8d4d90                  lea ecx, dword ptr [ebp-70]
'0070ff61    51                      push ecx
'0070ff62    8d55a0                  lea edx, dword ptr [ebp-60]
'0070ff65    52                      push edx
'0070ff66    8d45b0                  lea eax, dword ptr [ebp-50]
'0070ff69    50                      push eax
'0070ff6a    6a04                    push 04

' *** Reference to "__vbaFreeVarList"
'0070ff6c    ff1540104000            call dword ptr [00401040]
    '#Cleanup( var_6, var_7, var_8, var_18)
'0070ff72    8b0e                    mov ecx, dword ptr [esi]
'0070ff74    83c414                  add esp, 14
'0070ff77    56                      push esi
'0070ff78    ff9104030000            call dword ptr [ecx+00000304]
'0070ff7e    50                      push eax
'0070ff7f    8d55c8                  lea edx, dword ptr [ebp-38]
'0070ff82    52                      push edx
'0070ff83    ffd3                    call ebx
    Set var_46 = 
'0070ff85    8bf0                    mov esi, eax
'0070ff87    8b06                    mov eax, dword ptr [esi]
'0070ff89    56                      push esi
'0070ff8a    ff90f4010000            call dword ptr [eax+000001f4]
'0070ff90    dbe2                    fnclex
'0070ff92    85c0                    test ax, ax
'0070ff94    7d12                    jge 70ffa8
'0070ff96    68f4010000              push 000001f4
'0070ff9b    681cb94300              push 0043b91c
'0070ffa0    56                      push esi
'0070ffa1    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0070ffa2    ff1580104000            call dword ptr [00401080]
'0070ffa8    8d4dc8                  lea ecx, dword ptr [ebp-38]

' *** Reference to "__vbaFreeObj"
'0070ffab    ff1508134000            call dword ptr [00401308]
    '#Cleanup(var_46)
'0070ffb1    e9fc170000              jmp 7117b2
    
End If
'0070ffb6    8d7dc8                  lea edi, dword ptr [ebp-38]
'0070ffb9    57                      push edi
'0070ffba    83ec10                  sub esp, 10
'0070ffbd    8bfc                    mov edi, esp
'0070ffbf    b90a000000              mov ecx, 0000000a
'0070ffc4    890f                    mov dword ptr [edi], ecx
'0070ffc6    898d40ffffff            mov dword ptr [ebp+ffffff40], ecx
'0070ffcc    8b8d34ffffff            mov ecx, dword ptr [ebp+ffffff34]
'0070ffd2    894f04                  mov dword ptr [edi+04], ecx
'0070ffd5    8b154cb07200            mov edx, dword ptr [0072b04c]
'0070ffdb    b804000280              mov eax, 80020004
'0070ffe0    894708                  mov dword ptr [edi+08], eax
'0070ffe3    898548ffffff            mov dword ptr [ebp+ffffff48], eax
'0070ffe9    8b853cffffff            mov eax, dword ptr [ebp+ffffff3c]
'0070ffef    89470c                  mov dword ptr [edi+0c], eax
'0070fff2    8b8540ffffff            mov eax, dword ptr [ebp+ffffff40]
'0070fff8    83ec10                  sub esp, 10
'0070fffb    8bcc                    mov ecx, esp
'0070fffd    8901                    mov dword ptr [ecx], eax
'0070ffff    8b8544ffffff            mov eax, dword ptr [ebp+ffffff44]
'00710005    894104                  mov dword ptr [ecx+04], eax
'00710008    8b8548ffffff            mov eax, dword ptr [ebp+ffffff48]
'0071000e    894108                  mov dword ptr [ecx+08], eax
'00710011    8b854cffffff            mov eax, dword ptr [ebp+ffffff4c]
'00710017    89410c                  mov dword ptr [ecx+0c], eax
'0071001a    83ec10                  sub esp, 10
'0071001d    c78550ffffff03000000    mov dword ptr [ebp+ffffff50], 00000003
'00710027    8b8550ffffff            mov eax, dword ptr [ebp+ffffff50]
'0071002d    8bcc                    mov ecx, esp
'0071002f    8901                    mov dword ptr [ecx], eax
'00710031    8b8554ffffff            mov eax, dword ptr [ebp+ffffff54]
'00710037    894104                  mov dword ptr [ecx+04], eax
'0071003a    c78558ffffff01000000    mov dword ptr [ebp+ffffff58], 00000001
'00710044    8b8558ffffff            mov eax, dword ptr [ebp+ffffff58]
'0071004a    8b12                    mov edx, dword ptr [edx]
'0071004c    894108                  mov dword ptr [ecx+08], eax
'0071004f    8b855cffffff            mov eax, dword ptr [ebp+ffffff5c]
'00710055    89410c                  mov dword ptr [ecx+0c], eax
'00710058    8b0d4cb07200            mov ecx, dword ptr [0072b04c]
'0071005e    68049e4300              push 00439e04
'00710063    51                      push ecx
'00710064    ff92bc000000            call dword ptr [edx+000000bc]
'0071006a    dbe2                    fnclex
'0071006c    85c0                    test ax, ax
'0071006e    7d18                    jge 710088
'00710070    8b154cb07200            mov edx, dword ptr [0072b04c]
'00710076    68bc000000              push 000000bc
'0071007b    68301f4300              push 00431f30
'00710080    52                      push edx
'00710081    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00710082    ff1580104000            call dword ptr [00401080]
'00710088    8b45c8                  mov eax, dword ptr [ebp-38]
'0071008b    50                      push eax
'0071008c    8d45e0                  lea eax, dword ptr [ebp-20]
'0071008f    50                      push eax
'00710090    c745c800000000          mov dword ptr [ebp-38], 00000000
'00710097    ffd3                    call ebx
Set var_3 = var_46
'00710099    8b45e0                  mov eax, dword ptr [ebp-20]
'0071009c    8b08                    mov ecx, dword ptr [eax]
'0071009e    68149e4300              push 00439e14
'007100a3    50                      push eax
'007100a4    ff5144                  call dword ptr [ecx+44]
'007100a7    dbe2                    fnclex
'007100a9    85c0                    test ax, ax
'007100ab    7d12                    jge 7100bf
'007100ad    8b55e0                  mov edx, dword ptr [ebp-20]
'007100b0    6a44                    push 44
'007100b2    6830314300              push 00433130
'007100b7    52                      push edx
'007100b8    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'007100b9    ff1580104000            call dword ptr [00401080]
'007100bf    8b06                    mov eax, dword ptr [esi]
'007100c1    56                      push esi
'007100c2    ff9008030000            call dword ptr [eax+00000308]
'007100c8    50                      push eax
'007100c9    8d4dc8                  lea ecx, dword ptr [ebp-38]
'007100cc    51                      push ecx
'007100cd    ffd3                    call ebx
Set var_46 = Nothing
'007100cf    83ec10                  sub esp, 10
'007100d2    8bdc                    mov ebx, esp
'007100d4    b90a000000              mov ecx, 0000000a
'007100d9    890b                    mov dword ptr [ebx], ecx
'007100db    898d40ffffff            mov dword ptr [ebp+ffffff40], ecx
'007100e1    898d50ffffff            mov dword ptr [ebp+ffffff50], ecx
'007100e7    8b8da4feffff            mov ecx, dword ptr [ebp+fffffea4]
'007100ed    894b04                  mov dword ptr [ebx+04], ecx
'007100f0    b804000280              mov eax, 80020004
'007100f5    894308                  mov dword ptr [ebx+08], eax
'007100f8    8bd0                    mov edx, eax
'007100fa    898548ffffff            mov dword ptr [ebp+ffffff48], eax
'00710100    898558ffffff            mov dword ptr [ebp+ffffff58], eax
'00710106    8b85acfeffff            mov eax, dword ptr [ebp+fffffeac]
'0071010c    89430c                  mov dword ptr [ebx+0c], eax
'0071010f    83ec10                  sub esp, 10
'00710112    8bcc                    mov ecx, esp
'00710114    b80a000000              mov eax, 0000000a
'00710119    8901                    mov dword ptr [ecx], eax
'0071011b    8b85b4feffff            mov eax, dword ptr [ebp+fffffeb4]
'00710121    894104                  mov dword ptr [ecx+04], eax
'00710124    895108                  mov dword ptr [ecx+08], edx
'00710127    8b95bcfeffff            mov edx, dword ptr [ebp+fffffebc]
'0071012d    89510c                  mov dword ptr [ecx+0c], edx
'00710130    8b95c4feffff            mov edx, dword ptr [ebp+fffffec4]
'00710136    83ec10                  sub esp, 10
'00710139    8bcc                    mov ecx, esp
'0071013b    b80a000000              mov eax, 0000000a
'00710140    8901                    mov dword ptr [ecx], eax
'00710142    895104                  mov dword ptr [ecx+04], edx
'00710145    8b95d4feffff            mov edx, dword ptr [ebp+fffffed4]
'0071014b    8b7dc8                  mov edi, dword ptr [ebp-38]
'0071014e    b804000280              mov eax, 80020004
'00710153    894108                  mov dword ptr [ecx+08], eax
'00710156    8b85ccfeffff            mov eax, dword ptr [ebp+fffffecc]
'0071015c    89410c                  mov dword ptr [ecx+0c], eax
'0071015f    83ec10                  sub esp, 10
'00710162    8bcc                    mov ecx, esp
'00710164    b80a000000              mov eax, 0000000a
'00710169    8901                    mov dword ptr [ecx], eax
'0071016b    895104                  mov dword ptr [ecx+04], edx
'0071016e    8b95e4feffff            mov edx, dword ptr [ebp+fffffee4]
'00710174    b804000280              mov eax, 80020004
'00710179    894108                  mov dword ptr [ecx+08], eax
'0071017c    8b85dcfeffff            mov eax, dword ptr [ebp+fffffedc]
'00710182    89410c                  mov dword ptr [ecx+0c], eax
'00710185    83ec10                  sub esp, 10
'00710188    8bcc                    mov ecx, esp
'0071018a    b80a000000              mov eax, 0000000a
'0071018f    8901                    mov dword ptr [ecx], eax
'00710191    895104                  mov dword ptr [ecx+04], edx
'00710194    8b95f4feffff            mov edx, dword ptr [ebp+fffffef4]
'0071019a    b804000280              mov eax, 80020004
'0071019f    894108                  mov dword ptr [ecx+08], eax
'007101a2    8b85ecfeffff            mov eax, dword ptr [ebp+fffffeec]
'007101a8    89410c                  mov dword ptr [ecx+0c], eax
'007101ab    83ec10                  sub esp, 10
'007101ae    8bcc                    mov ecx, esp
'007101b0    b80a000000              mov eax, 0000000a
'007101b5    8901                    mov dword ptr [ecx], eax
'007101b7    895104                  mov dword ptr [ecx+04], edx
'007101ba    b804000280              mov eax, 80020004
'007101bf    894108                  mov dword ptr [ecx+08], eax
'007101c2    8b85fcfeffff            mov eax, dword ptr [ebp+fffffefc]
'007101c8    89410c                  mov dword ptr [ecx+0c], eax
'007101cb    897db8                  mov dword ptr [ebp-48], edi
'007101ce    8b7de0                  mov edi, dword ptr [ebp-20]
'007101d1    83ec10                  sub esp, 10
'007101d4    8bcc                    mov ecx, esp
'007101d6    b80a000000              mov eax, 0000000a
'007101db    c745c800000000          mov dword ptr [ebp-38], 00000000
'007101e2    c745b009000000          mov dword ptr [ebp-50], 00000009
'007101e9    8b3f                    mov edi, dword ptr [edi]
'007101eb    8901                    mov dword ptr [ecx], eax
'007101ed    8b9504ffffff            mov edx, dword ptr [ebp+ffffff04]
'007101f3    895104                  mov dword ptr [ecx+04], edx
'007101f6    8b9514ffffff            mov edx, dword ptr [ebp+ffffff14]
'007101fc    83ec10                  sub esp, 10
'007101ff    b804000280              mov eax, 80020004
'00710204    894108                  mov dword ptr [ecx+08], eax
'00710207    8b850cffffff            mov eax, dword ptr [ebp+ffffff0c]
'0071020d    89410c                  mov dword ptr [ecx+0c], eax
'00710210    8bcc                    mov ecx, esp
'00710212    b80a000000              mov eax, 0000000a
'00710217    8901                    mov dword ptr [ecx], eax
'00710219    895104                  mov dword ptr [ecx+04], edx
'0071021c    8b9524ffffff            mov edx, dword ptr [ebp+ffffff24]
'00710222    83ec10                  sub esp, 10
'00710225    b804000280              mov eax, 80020004
'0071022a    894108                  mov dword ptr [ecx+08], eax
'0071022d    8b851cffffff            mov eax, dword ptr [ebp+ffffff1c]
'00710233    89410c                  mov dword ptr [ecx+0c], eax
'00710236    8bcc                    mov ecx, esp
'00710238    b80a000000              mov eax, 0000000a
'0071023d    8901                    mov dword ptr [ecx], eax
'0071023f    895104                  mov dword ptr [ecx+04], edx
'00710242    8b9534ffffff            mov edx, dword ptr [ebp+ffffff34]
'00710248    83ec10                  sub esp, 10
'0071024b    b804000280              mov eax, 80020004
'00710250    894108                  mov dword ptr [ecx+08], eax
'00710253    8b852cffffff            mov eax, dword ptr [ebp+ffffff2c]
'00710259    89410c                  mov dword ptr [ecx+0c], eax
'0071025c    8bcc                    mov ecx, esp
'0071025e    b80a000000              mov eax, 0000000a
'00710263    8901                    mov dword ptr [ecx], eax
'00710265    895104                  mov dword ptr [ecx+04], edx
'00710268    8b9540ffffff            mov edx, dword ptr [ebp+ffffff40]
'0071026e    b804000280              mov eax, 80020004
'00710273    894108                  mov dword ptr [ecx+08], eax
'00710276    8b853cffffff            mov eax, dword ptr [ebp+ffffff3c]
'0071027c    89410c                  mov dword ptr [ecx+0c], eax
'0071027f    8b8544ffffff            mov eax, dword ptr [ebp+ffffff44]
'00710285    83ec10                  sub esp, 10
'00710288    8bcc                    mov ecx, esp
'0071028a    8911                    mov dword ptr [ecx], edx
'0071028c    8b9548ffffff            mov edx, dword ptr [ebp+ffffff48]
'00710292    894104                  mov dword ptr [ecx+04], eax
'00710295    8b854cffffff            mov eax, dword ptr [ebp+ffffff4c]
'0071029b    895108                  mov dword ptr [ecx+08], edx
'0071029e    8b9550ffffff            mov edx, dword ptr [ebp+ffffff50]
'007102a4    89410c                  mov dword ptr [ecx+0c], eax
'007102a7    8b8554ffffff            mov eax, dword ptr [ebp+ffffff54]
'007102ad    83ec10                  sub esp, 10
'007102b0    8bcc                    mov ecx, esp
'007102b2    8911                    mov dword ptr [ecx], edx
'007102b4    8b9558ffffff            mov edx, dword ptr [ebp+ffffff58]
'007102ba    894104                  mov dword ptr [ecx+04], eax
'007102bd    8b855cffffff            mov eax, dword ptr [ebp+ffffff5c]
'007102c3    895108                  mov dword ptr [ecx+08], edx
'007102c6    8b55b0                  mov edx, dword ptr [ebp-50]
'007102c9    89410c                  mov dword ptr [ecx+0c], eax
'007102cc    8b45b4                  mov eax, dword ptr [ebp-4c]
'007102cf    83ec10                  sub esp, 10
'007102d2    8bcc                    mov ecx, esp
'007102d4    8911                    mov dword ptr [ecx], edx
'007102d6    8b55b8                  mov edx, dword ptr [ebp-48]
'007102d9    894104                  mov dword ptr [ecx+04], eax
'007102dc    8b45bc                  mov eax, dword ptr [ebp-44]
'007102df    895108                  mov dword ptr [ecx+08], edx
'007102e2    89410c                  mov dword ptr [ecx+0c], eax
'007102e5    8b4de0                  mov ecx, dword ptr [ebp-20]
'007102e8    68209e4300              push 00439e20
'007102ed    51                      push ecx
'007102ee    ff97f4000000            call dword ptr [edi+000000f4]
'007102f4    dbe2                    fnclex
'007102f6    85c0                    test ax, ax
'007102f8    7d15                    jge 71030f
'007102fa    8b55e0                  mov edx, dword ptr [ebp-20]
'007102fd    68f4000000              push 000000f4
'00710302    6830314300              push 00433130
'00710307    52                      push edx
'00710308    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00710309    ff1580104000            call dword ptr [00401080]
'0071030f    8d4dc8                  lea ecx, dword ptr [ebp-38]

' *** Reference to "__vbaFreeObj"
'00710312    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_46)
'00710318    8d4db0                  lea ecx, dword ptr [ebp-50]

' *** Reference to "__vbaFreeVar"
'0071031b    ff1528104000            call dword ptr [00401028]
'#Cleanup(var_6)
'00710321    8b45e0                  mov eax, dword ptr [ebp-20]
'00710324    8b08                    mov ecx, dword ptr [eax]
'00710326    8d959cfeffff            lea edx, dword ptr [ebp+fffffe9c]
'0071032c    52                      push edx
'0071032d    50                      push eax
'0071032e    ff515c                  call dword ptr [ecx+5c]
'00710331    dbe2                    fnclex
'00710333    85c0                    test ax, ax
'00710335    7d12                    jge 710349
'00710337    8b4de0                  mov ecx, dword ptr [ebp-20]
'0071033a    6a5c                    push 5c
'0071033c    6830314300              push 00433130
'00710341    51                      push ecx
'00710342    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00710343    ff1580104000            call dword ptr [00401080]
'00710349    6683bd9cfeffff00        cmp word ptr [ebp+fffffe9c], 00
'00710351    bb04000280              mov ebx, 80020004
'00710356    756f                    jne 7103c7

If (0 = 0) Then
'00710358    bf0a000000              mov edi, 0000000a
'0071035d    8d9550ffffff            lea edx, dword ptr [ebp+ffffff50]
'00710363    8d4db0                  lea ecx, dword ptr [ebp-50]
'00710366    895d88                  mov dword ptr [ebp-78], ebx
'00710369    897d80                  mov dword ptr [ebp-80], edi
'0071036c    895d98                  mov dword ptr [ebp-68], ebx
'0071036f    897d90                  mov dword ptr [ebp-70], edi
'00710372    895da8                  mov dword ptr [ebp-58], ebx
'00710375    897da0                  mov dword ptr [ebp-60], edi
'00710378    c78558ffffff90684500    mov dword ptr [ebp+ffffff58], 00456890
'00710382    c78550ffffff08000000    mov dword ptr [ebp+ffffff50], 00000008

' *** Reference to "__vbaVarDup"
'0071038c    ff1598124000            call dword ptr [00401298]
    var_6 = ("Attention ce don existe dÈja dans la base de donnÈe du programmme, il sera modifiÈ si vous ne changer pas le nom.")
'00710392    8d5580                  lea edx, dword ptr [ebp-80]
'00710395    52                      push edx
'00710396    8d4590                  lea eax, dword ptr [ebp-70]
'00710399    50                      push eax
'0071039a    8d4da0                  lea ecx, dword ptr [ebp-60]
'0071039d    51                      push ecx
'0071039e    6a00                    push 00
'007103a0    8d55b0                  lea edx, dword ptr [ebp-50]
'007103a3    52                      push edx

' *** Reference to "rtcMsgBox"
'007103a4    ff15b8104000            call dword ptr [004010b8]
    var_55 = MsgBox(var_6, 0)
'007103aa    8d4580                  lea eax, dword ptr [ebp-80]
'007103ad    50                      push eax
'007103ae    8d4d90                  lea ecx, dword ptr [ebp-70]
'007103b1    51                      push ecx
'007103b2    8d55a0                  lea edx, dword ptr [ebp-60]
'007103b5    52                      push edx
'007103b6    8d45b0                  lea eax, dword ptr [ebp-50]
'007103b9    50                      push eax
'007103ba    6a04                    push 04

' *** Reference to "__vbaFreeVarList"
'007103bc    ff1540104000            call dword ptr [00401040]
    '#Cleanup( var_6, var_7, var_8, var_18)
'007103c2    83c414                  add esp, 14
'007103c5    eb05                    jmp 7103cc
    
End If
'007103c7    bf0a000000              mov edi, 0000000a
'007103cc    8bd3                    mov edx, ebx
'007103ce    8bc3                    mov eax, ebx
'007103d0    8d5dc8                  lea ebx, dword ptr [ebp-38]
'007103d3    53                      push ebx
'007103d4    83ec10                  sub esp, 10
'007103d7    8bdc                    mov ebx, esp
'007103d9    8bcf                    mov ecx, edi
'007103db    890b                    mov dword ptr [ebx], ecx
'007103dd    8b8d34ffffff            mov ecx, dword ptr [ebp+ffffff34]
'007103e3    894b04                  mov dword ptr [ebx+04], ecx
'007103e6    894308                  mov dword ptr [ebx+08], eax
'007103e9    8b853cffffff            mov eax, dword ptr [ebp+ffffff3c]
'007103ef    89430c                  mov dword ptr [ebx+0c], eax
'007103f2    83ec10                  sub esp, 10
'007103f5    89bd40ffffff            mov dword ptr [ebp+ffffff40], edi
'007103fb    8b8540ffffff            mov eax, dword ptr [ebp+ffffff40]
'00710401    8b3d48b07200            mov edi, dword ptr [0072b048]
'00710407    8bcc                    mov ecx, esp
'00710409    8901                    mov dword ptr [ecx], eax
'0071040b    8b8544ffffff            mov eax, dword ptr [ebp+ffffff44]
'00710411    894104                  mov dword ptr [ecx+04], eax
'00710414    895108                  mov dword ptr [ecx+08], edx
'00710417    899548ffffff            mov dword ptr [ebp+ffffff48], edx
'0071041d    8b954cffffff            mov edx, dword ptr [ebp+ffffff4c]
'00710423    89510c                  mov dword ptr [ecx+0c], edx
'00710426    8b9554ffffff            mov edx, dword ptr [ebp+ffffff54]
'0071042c    83ec10                  sub esp, 10
'0071042f    c78550ffffff03000000    mov dword ptr [ebp+ffffff50], 00000003
'00710439    8b8d50ffffff            mov ecx, dword ptr [ebp+ffffff50]
'0071043f    8bc4                    mov eax, esp
'00710441    8908                    mov dword ptr [eax], ecx
'00710443    895004                  mov dword ptr [eax+04], edx
'00710446    8b955cffffff            mov edx, dword ptr [ebp+ffffff5c]
'0071044c    c78558ffffff01000000    mov dword ptr [ebp+ffffff58], 00000001
'00710456    8b8d58ffffff            mov ecx, dword ptr [ebp+ffffff58]
'0071045c    8b3f                    mov edi, dword ptr [edi]
'0071045e    894808                  mov dword ptr [eax+08], ecx
'00710461    89500c                  mov dword ptr [eax+0c], edx
'00710464    a148b07200              mov ax, word ptr [0072b048]
'00710469    6878694500              push 00456978
'0071046e    50                      push eax
'0071046f    ff97bc000000            call dword ptr [edi+000000bc]
'00710475    dbe2                    fnclex
'00710477    85c0                    test ax, ax
'00710479    7d1c                    jge 710497
'0071047b    8b0d48b07200            mov ecx, dword ptr [0072b048]

' *** Reference to "__vbaHresultCheckObj"
'00710481    8b1d80104000            mov ebx, dword ptr [00401080]
'00710487    68bc000000              push 000000bc
'0071048c    68301f4300              push 00431f30
'00710491    51                      push ecx
'00710492    50                      push eax
'00710493    ffd3                    call ebx
'00710495    eb06                    jmp 71049d

' *** Reference to "__vbaHresultCheckObj"
'00710497    8b1d80104000            mov ebx, dword ptr [00401080]
'0071049d    8b45c8                  mov eax, dword ptr [ebp-38]

' *** Reference to "__vbaObjSet"
'007104a0    8b3db4104000            mov edi, dword ptr [004010b4]
'007104a6    50                      push eax
'007104a7    8d55dc                  lea edx, dword ptr [ebp-24]
'007104aa    52                      push edx
'007104ab    c745c800000000          mov dword ptr [ebp-38], 00000000
'007104b2    ffd7                    call edi
Set var_39 = Nothing
'007104b4    8b45dc                  mov eax, dword ptr [ebp-24]
'007104b7    50                      push eax
'007104b8    8d8d78feffff            lea ecx, dword ptr [ebp+fffffe78]
'007104be    51                      push ecx

' *** Reference to "__vbaObjSetAddref"
'007104bf    ff15c8104000            call dword ptr [004010c8]
Set var_73 = Nothing
'007104c5    8b16                    mov edx, dword ptr [esi]
'007104c7    56                      push esi
'007104c8    ff9208030000            call dword ptr [edx+00000308]
'007104ce    50                      push eax
'007104cf    8d45c8                  lea eax, dword ptr [ebp-38]
'007104d2    50                      push eax
'007104d3    ffd7                    call edi
Set var_46 = Nothing
'007104d5    8d55d4                  lea edx, dword ptr [ebp-2c]
'007104d8    8bf8                    mov edi, eax
'007104da    8b0f                    mov ecx, dword ptr [edi]
'007104dc    52                      push edx
'007104dd    57                      push edi
'007104de    ff91a8000000            call dword ptr [ecx+000000a8]
'007104e4    dbe2                    fnclex
'007104e6    85c0                    test ax, ax
'007104e8    7d0e                    jge 7104f8
'007104ea    68a8000000              push 000000a8
'007104ef    681cb94300              push 0043b91c
'007104f4    57                      push edi
'007104f5    50                      push eax
'007104f6    ffd3                    call ebx
'007104f8    8b55d4                  mov edx, dword ptr [ebp-2c]
'007104fb    33ff                    xor edi, edi
var_num7 = Empty
'007104fd    8d4de8                  lea ecx, dword ptr [ebp-18]
'00710500    897dd4                  mov dword ptr [ebp-2c], edi

' *** Reference to "__vbaStrMove"
'00710503    ff15d0124000            call dword ptr [004012d0]
'00710509    8d4dc8                  lea ecx, dword ptr [ebp-38]

' *** Reference to "__vbaFreeObj"
'0071050c    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_46)
'00710512    897dd8                  mov dword ptr [ebp-28], edi
'00710515    8b8578feffff            mov eax, dword ptr [ebp+fffffe78]
'0071051b    8b08                    mov ecx, dword ptr [eax]
'0071051d    8d959cfeffff            lea edx, dword ptr [ebp+fffffe9c]
'00710523    52                      push edx
'00710524    50                      push eax
'00710525    ff5134                  call dword ptr [ecx+34]
'00710528    dbe2                    fnclex
'0071052a    85c0                    test ax, ax
'0071052c    7d11                    jge 71053f
'0071052e    8b8d78feffff            mov ecx, dword ptr [ebp+fffffe78]
'00710534    6a34                    push 34
'00710536    6830314300              push 00433130
'0071053b    51                      push ecx
'0071053c    50                      push eax
'0071053d    ffd3                    call ebx
'0071053f    6683bd9cfeffff00        cmp word ptr [ebp+fffffe9c], 00
'00710547    0f8575010000            jne 7106c2

If (0 = 0) Then
'0071054d    8b8578feffff            mov eax, dword ptr [ebp+fffffe78]
'00710553    8b10                    mov edx, dword ptr [eax]
'00710555    8d4dc8                  lea ecx, dword ptr [ebp-38]
'00710558    51                      push ecx
'00710559    50                      push eax
'0071055a    ff92b4000000            call dword ptr [edx+000000b4]
'00710560    dbe2                    fnclex
'00710562    85c0                    test ax, ax
'00710564    7d14                    jge 71057a
'00710566    8b9578feffff            mov edx, dword ptr [ebp+fffffe78]
'0071056c    68b4000000              push 000000b4
'00710571    6830314300              push 00433130
'00710576    52                      push edx
'00710577    50                      push eax
'00710578    ffd3                    call ebx
'0071057a    8b45c8                  mov eax, dword ptr [ebp-38]
'0071057d    8d5dc4                  lea ebx, dword ptr [ebp-3c]
'00710580    53                      push ebx
'00710581    83ec10                  sub esp, 10
'00710584    ba08000000              mov edx, 00000008
'00710589    8bdc                    mov ebx, esp
'0071058b    8913                    mov dword ptr [ebx], edx
'0071058d    899550ffffff            mov dword ptr [ebp+ffffff50], edx
'00710593    8b9554ffffff            mov edx, dword ptr [ebp+ffffff54]
'00710599    b964b14300              mov ecx, 0043b164
'0071059e    895304                  mov dword ptr [ebx+04], edx
'007105a1    898d58ffffff            mov dword ptr [ebp+ffffff58], ecx
'007105a7    8b38                    mov edi, dword ptr [eax]
'007105a9    894b08                  mov dword ptr [ebx+08], ecx
'007105ac    8b8d5cffffff            mov ecx, dword ptr [ebp+ffffff5c]
'007105b2    50                      push eax
'007105b3    898594feffff            mov dword ptr [ebp+fffffe94], eax
'007105b9    894b0c                  mov dword ptr [ebx+0c], ecx
'007105bc    ff5730                  call dword ptr [edi+30]
'007105bf    dbe2                    fnclex
'007105c1    85c0                    test ax, ax
'007105c3    7d19                    jge 7105de
'007105c5    8b9594feffff            mov edx, dword ptr [ebp+fffffe94]

' *** Reference to "__vbaHresultCheckObj"
'007105cb    8b1d80104000            mov ebx, dword ptr [00401080]
'007105d1    6a30                    push 30
'007105d3    68d8304300              push 004330d8
'007105d8    52                      push edx
'007105d9    50                      push eax
'007105da    ffd3                    call ebx
'007105dc    eb06                    jmp 7105e4

' *** Reference to "__vbaHresultCheckObj"
'007105de    8b1d80104000            mov ebx, dword ptr [00401080]
'007105e4    8b45c4                  mov eax, dword ptr [ebp-3c]
'007105e7    8b08                    mov ecx, dword ptr [eax]
'007105e9    8d55b0                  lea edx, dword ptr [ebp-50]
'007105ec    52                      push edx
'007105ed    50                      push eax
'007105ee    8bf8                    mov edi, eax
'007105f0    ff5144                  call dword ptr [ecx+44]
'007105f3    dbe2                    fnclex
'007105f5    85c0                    test ax, ax
'007105f7    7d0b                    jge 710604
'007105f9    6a44                    push 44
'007105fb    6880304300              push 00433080
'00710600    57                      push edi
'00710601    50                      push eax
'00710602    ffd3                    call ebx
'00710604    8b06                    mov eax, dword ptr [esi]
'00710606    56                      push esi
'00710607    ff9008030000            call dword ptr [eax+00000308]
'0071060d    50                      push eax
'0071060e    8d4dc0                  lea ecx, dword ptr [ebp-40]
'00710611    51                      push ecx

' *** Reference to "__vbaObjSet"
'00710612    ff15b4104000            call dword ptr [004010b4]
    Set var_5 = Nothing
'00710618    8bf8                    mov edi, eax
'0071061a    8b17                    mov edx, dword ptr [edi]
'0071061c    8d45d4                  lea eax, dword ptr [ebp-2c]
'0071061f    50                      push eax
'00710620    57                      push edi
'00710621    ff92a8000000            call dword ptr [edx+000000a8]
'00710627    dbe2                    fnclex
'00710629    85c0                    test ax, ax
'0071062b    7d0e                    jge 71063b
'0071062d    68a8000000              push 000000a8
'00710632    681cb94300              push 0043b91c
'00710637    57                      push edi
'00710638    50                      push eax
'00710639    ffd3                    call ebx
'0071063b    8b45d4                  mov eax, dword ptr [ebp-2c]
'0071063e    8d4db0                  lea ecx, dword ptr [ebp-50]
'00710641    51                      push ecx
'00710642    8d55a0                  lea edx, dword ptr [ebp-60]
'00710645    52                      push edx
'00710646    c745d400000000          mov dword ptr [ebp-2c], 00000000
'0071064d    8945a8                  mov dword ptr [ebp-58], eax
'00710650    c745a008800000          mov dword ptr [ebp-60], 00008008

' *** Reference to "__vbaVarTstEq"
'00710657    ff153c114000            call dword ptr [0040113c]
'0071065d    8d4dc0                  lea ecx, dword ptr [ebp-40]
'00710660    8d55c8                  lea edx, dword ptr [ebp-38]
'00710663    8bf8                    mov edi, eax
'00710665    8d45c4                  lea eax, dword ptr [ebp-3c]
'00710668    50                      push eax
'00710669    51                      push ecx
'0071066a    52                      push edx
'0071066b    6a03                    push 03

' *** Reference to "__vbaFreeObjList"
'0071066d    ff154c104000            call dword ptr [0040104c]
    '#Cleanup( var_46, var_5, var_9)
'00710673    8d45a0                  lea eax, dword ptr [ebp-60]
'00710676    50                      push eax
'00710677    8d4db0                  lea ecx, dword ptr [ebp-50]
'0071067a    51                      push ecx
'0071067b    6a02                    push 02

' *** Reference to "__vbaFreeVarList"
'0071067d    ff1540104000            call dword ptr [00401040]
    '#Cleanup( var_6, var_7)
'00710683    83c41c                  add esp, 1c
'00710686    6685ff                  test edi, edi
'00710689    7532                    jne 7106bd
    
    Do While (    ((var_6) = (vbNullString)))
'0071068b    8b8578feffff            mov eax, dword ptr [ebp+fffffe78]
'00710691    8b10                    mov edx, dword ptr [eax]
'00710693    50                      push eax
'00710694    ff92ec000000            call dword ptr [edx+000000ec]
'0071069a    dbe2                    fnclex
'0071069c    85c0                    test ax, ax
'0071069e    0f8d71feffff            jge 710515
'007106a4    8b8d78feffff            mov ecx, dword ptr [ebp+fffffe78]
'007106aa    68ec000000              push 000000ec
'007106af    6830314300              push 00433130
'007106b4    51                      push ecx
'007106b5    50                      push eax
'007106b6    ffd3                    call ebx
'007106b8    e958feffff              jmp 710515
    
Loop
'007106bd    83c8ff                  or eax, ffffffff
'007106c0    eb03                    jmp 7106c5

'ERROR: Two many next close:
End If
'007106c2    8b45d8                  mov eax, dword ptr [ebp-28]
'007106c5    6685c0                  test eax, eax
'007106c8    8b8578feffff            mov eax, dword ptr [ebp+fffffe78]
'007106ce    8b10                    mov edx, dword ptr [eax]
'007106d0    50                      push eax
'007106d1    7429                    je 7106fc
'007106d3    ff92d0000000            call dword ptr [edx+000000d0]
'007106d9    dbe2                    fnclex
'007106db    85c0                    test ax, ax
'007106dd    0f8de0000000            jge 7107c3
'007106e3    8b8d78feffff            mov ecx, dword ptr [ebp+fffffe78]
'007106e9    68d0000000              push 000000d0
'007106ee    6830314300              push 00433130
'007106f3    51                      push ecx
'007106f4    50                      push eax
'007106f5    ffd3                    call ebx
'007106f7    e9c7000000              jmp 7107c3
'007106fc    ff92c0000000            call dword ptr [edx+000000c0]
'00710702    dbe2                    fnclex
'00710704    85c0                    test ax, ax
'00710706    7d14                    jge 71071c
'00710708    8b8d78feffff            mov ecx, dword ptr [ebp+fffffe78]
'0071070e    68c0000000              push 000000c0
'00710713    6830314300              push 00433130
'00710718    51                      push ecx
'00710719    50                      push eax
'0071071a    ffd3                    call ebx
'0071071c    8b16                    mov edx, dword ptr [esi]
'0071071e    56                      push esi
'0071071f    ff9208030000            call dword ptr [edx+00000308]
'00710725    8bbd78feffff            mov edi, dword ptr [ebp+fffffe78]
'0071072b    83ec10                  sub esp, 10
'0071072e    8bdc                    mov ebx, esp
'00710730    b909000000              mov ecx, 00000009
'00710735    890b                    mov dword ptr [ebx], ecx
'00710737    894db0                  mov dword ptr [ebp-50], ecx
'0071073a    8b4db4                  mov ecx, dword ptr [ebp-4c]
'0071073d    894b04                  mov dword ptr [ebx+04], ecx
'00710740    894308                  mov dword ptr [ebx+08], eax
'00710743    8945b8                  mov dword ptr [ebp-48], eax
'00710746    8b45bc                  mov eax, dword ptr [ebp-44]
'00710749    89430c                  mov dword ptr [ebx+0c], eax
'0071074c    83ec10                  sub esp, 10
'0071074f    c78550ffffff08000000    mov dword ptr [ebp+ffffff50], 00000008
'00710759    8b8550ffffff            mov eax, dword ptr [ebp+ffffff50]
'0071075f    8bcc                    mov ecx, esp
'00710761    8901                    mov dword ptr [ecx], eax
'00710763    8b8554ffffff            mov eax, dword ptr [ebp+ffffff54]
'00710769    894104                  mov dword ptr [ecx+04], eax
'0071076c    8b8578feffff            mov eax, dword ptr [ebp+fffffe78]
'00710772    ba64b14300              mov edx, 0043b164
'00710777    899558ffffff            mov dword ptr [ebp+ffffff58], edx
'0071077d    8b3f                    mov edi, dword ptr [edi]
'0071077f    895108                  mov dword ptr [ecx+08], edx
'00710782    8b955cffffff            mov edx, dword ptr [ebp+ffffff5c]
'00710788    50                      push eax
'00710789    89510c                  mov dword ptr [ecx+0c], edx
'0071078c    ff9728010000            call dword ptr [edi+00000128]
'00710792    dbe2                    fnclex
'00710794    85c0                    test ax, ax
'00710796    7d1c                    jge 7107b4
'00710798    8b8d78feffff            mov ecx, dword ptr [ebp+fffffe78]

' *** Reference to "__vbaHresultCheckObj"
'0071079e    8b1d80104000            mov ebx, dword ptr [00401080]
'007107a4    6828010000              push 00000128
'007107a9    6830314300              push 00433130
'007107ae    51                      push ecx
'007107af    50                      push eax
'007107b0    ffd3                    call ebx
'007107b2    eb06                    jmp 7107ba

' *** Reference to "__vbaHresultCheckObj"
'007107b4    8b1d80104000            mov ebx, dword ptr [00401080]
'007107ba    8d4db0                  lea ecx, dword ptr [ebp-50]

' *** Reference to "__vbaFreeVar"
'007107bd    ff1528104000            call dword ptr [00401028]
'#Cleanup(var_6)
'007107c3    8b16                    mov edx, dword ptr [esi]
'007107c5    56                      push esi
'007107c6    ff921c030000            call dword ptr [edx+0000031c]
'007107cc    8945b8                  mov dword ptr [ebp-48], eax
'007107cf    8d45b0                  lea eax, dword ptr [ebp-50]
'007107d2    50                      push eax
'007107d3    c745b009000000          mov dword ptr [ebp-50], 00000009

' *** Reference to "rtcIsNull"
'007107da    ff1540114000            call dword ptr [00401140]
'007107e0    8b0e                    mov ecx, dword ptr [esi]
'007107e2    56                      push esi
'007107e3    89859cfeffff            mov dword ptr [ebp+fffffe9c], eax
'007107e9    ff911c030000            call dword ptr [ecx+0000031c]
'007107ef    50                      push eax
'007107f0    8d55c8                  lea edx, dword ptr [ebp-38]
'007107f3    52                      push edx

' *** Reference to "__vbaObjSet"
'007107f4    ff15b4104000            call dword ptr [004010b4]
Set var_46 = IsNull(var_73)
'007107fa    8d4dd4                  lea ecx, dword ptr [ebp-2c]
'007107fd    8bf8                    mov edi, eax
'007107ff    8b07                    mov eax, dword ptr [edi]
'00710801    51                      push ecx
'00710802    57                      push edi
'00710803    ff90a0000000            call dword ptr [eax+000000a0]
'00710809    dbe2                    fnclex
'0071080b    85c0                    test ax, ax
'0071080d    7d0e                    jge 71081d
'0071080f    68a0000000              push 000000a0
'00710814    68d00d4300              push 00430dd0
'00710819    57                      push edi
'0071081a    50                      push eax
'0071081b    ffd3                    call ebx
'0071081d    8b55d4                  mov edx, dword ptr [ebp-2c]
'00710820    52                      push edx

' *** Reference to "rtcR8ValFromBstr"
'00710821    ff1510134000            call dword ptr [00401310]
'00710827    dd5d98                  fstp qword ptr [ebp-68]
'var_130 = (00)
'0071082a    668b859cfeffff          mov ax, word ptr [ebp+fffffe9c]
'00710831    8d4d90                  lea ecx, dword ptr [ebp-70]
'00710834    51                      push ecx
'00710835    8d55a0                  lea edx, dword ptr [ebp-60]
'00710838    66898558ffffff          mov word ptr [ebp+ffffff58], ax
'0071083f    52                      push edx
'00710840    8d8550ffffff            lea eax, dword ptr [ebp+ffffff50]
'00710846    50                      push eax
'00710847    8d4d80                  lea ecx, dword ptr [ebp-80]
'0071084a    51                      push ecx
'0071084b    c7459005000000          mov dword ptr [ebp-70], 00000005
'00710852    c745a800000000          mov dword ptr [ebp-58], 00000000
'00710859    c745a002000000          mov dword ptr [ebp-60], 00000002
'00710860    c78550ffffff0b000000    mov dword ptr [ebp+ffffff50], 0000000b

' *** Reference to "rtcImmediateIf"
'0071086a    ff154c124000            call dword ptr [0040124c]
'00710870    8b5d80                  mov ebx, dword ptr [ebp-80]
'00710873    8b9578feffff            mov edx, dword ptr [ebp+fffffe78]
'00710879    8b12                    mov edx, dword ptr [edx]
'0071087b    83ec10                  sub esp, 10
'0071087e    8bfc                    mov edi, esp
'00710880    891f                    mov dword ptr [edi], ebx
'00710882    8b5d84                  mov ebx, dword ptr [ebp-7c]
'00710885    895f04                  mov dword ptr [edi+04], ebx
'00710888    8b5d88                  mov ebx, dword ptr [ebp-78]
'0071088b    895f08                  mov dword ptr [edi+08], ebx
'0071088e    8b5d8c                  mov ebx, dword ptr [ebp-74]
'00710891    895f0c                  mov dword ptr [edi+0c], ebx
'00710894    83ec10                  sub esp, 10
'00710897    8bfc                    mov edi, esp
'00710899    b908000000              mov ecx, 00000008
'0071089e    890f                    mov dword ptr [edi], ecx
'007108a0    8b8d24ffffff            mov ecx, dword ptr [ebp+ffffff24]
'007108a6    894f04                  mov dword ptr [edi+04], ecx
'007108a9    8b8d78feffff            mov ecx, dword ptr [ebp+fffffe78]
'007108af    b80ca94300              mov eax, 0043a90c
'007108b4    894708                  mov dword ptr [edi+08], eax
'007108b7    8b852cffffff            mov eax, dword ptr [ebp+ffffff2c]
'007108bd    51                      push ecx
'007108be    89470c                  mov dword ptr [edi+0c], eax
'007108c1    ff9228010000            call dword ptr [edx+00000128]
'007108c7    dbe2                    fnclex
'007108c9    85c0                    test ax, ax
'007108cb    7d18                    jge 7108e5
'007108cd    8b9578feffff            mov edx, dword ptr [ebp+fffffe78]
'007108d3    6828010000              push 00000128
'007108d8    6830314300              push 00433130
'007108dd    52                      push edx
'007108de    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'007108df    ff1580104000            call dword ptr [00401080]
'007108e5    8d4dd4                  lea ecx, dword ptr [ebp-2c]

' *** Reference to "__vbaFreeStr"
'007108e8    ff150c134000            call dword ptr [0040130c]
'#Cleanup(var_44)
'007108ee    8d4dc8                  lea ecx, dword ptr [ebp-38]

' *** Reference to "__vbaFreeObj"
'007108f1    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_46)
'007108f7    8d4580                  lea eax, dword ptr [ebp-80]
'007108fa    50                      push eax
'007108fb    8d4d90                  lea ecx, dword ptr [ebp-70]
'007108fe    51                      push ecx
'007108ff    8d55a0                  lea edx, dword ptr [ebp-60]
'00710902    52                      push edx
'00710903    8d8550ffffff            lea eax, dword ptr [ebp+ffffff50]
'00710909    50                      push eax
'0071090a    8d4db0                  lea ecx, dword ptr [ebp-50]
'0071090d    51                      push ecx
'0071090e    6a05                    push 05

' *** Reference to "__vbaFreeVarList"
'00710910    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_6, var_21, var_7, var_8, var_18)
'00710916    8b16                    mov edx, dword ptr [esi]
'00710918    83c418                  add esp, 18
'0071091b    56                      push esi
'0071091c    ff9200030000            call dword ptr [edx+00000300]
'00710922    8bbd78feffff            mov edi, dword ptr [ebp+fffffe78]
'00710928    83ec10                  sub esp, 10
'0071092b    8bdc                    mov ebx, esp
'0071092d    b909000000              mov ecx, 00000009
'00710932    890b                    mov dword ptr [ebx], ecx
'00710934    894db0                  mov dword ptr [ebp-50], ecx
'00710937    8b4db4                  mov ecx, dword ptr [ebp-4c]
'0071093a    894b04                  mov dword ptr [ebx+04], ecx
'0071093d    894308                  mov dword ptr [ebx+08], eax
'00710940    8945b8                  mov dword ptr [ebp-48], eax
'00710943    8b45bc                  mov eax, dword ptr [ebp-44]
'00710946    89430c                  mov dword ptr [ebx+0c], eax
'00710949    83ec10                  sub esp, 10
'0071094c    c78550ffffff08000000    mov dword ptr [ebp+ffffff50], 00000008
'00710956    8b8550ffffff            mov eax, dword ptr [ebp+ffffff50]
'0071095c    8bcc                    mov ecx, esp
'0071095e    8901                    mov dword ptr [ecx], eax
'00710960    8b8554ffffff            mov eax, dword ptr [ebp+ffffff54]
'00710966    894104                  mov dword ptr [ecx+04], eax
'00710969    8b8578feffff            mov eax, dword ptr [ebp+fffffe78]
'0071096f    baacb24300              mov edx, 0043b2ac
'00710974    899558ffffff            mov dword ptr [ebp+ffffff58], edx
'0071097a    8b3f                    mov edi, dword ptr [edi]
'0071097c    895108                  mov dword ptr [ecx+08], edx
'0071097f    8b955cffffff            mov edx, dword ptr [ebp+ffffff5c]
'00710985    50                      push eax
'00710986    89510c                  mov dword ptr [ecx+0c], edx
'00710989    ff9728010000            call dword ptr [edi+00000128]
'0071098f    dbe2                    fnclex
'00710991    85c0                    test ax, ax
'00710993    7d18                    jge 7109ad
'00710995    8b8d78feffff            mov ecx, dword ptr [ebp+fffffe78]
'0071099b    6828010000              push 00000128
'007109a0    6830314300              push 00433130
'007109a5    51                      push ecx
'007109a6    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'007109a7    ff1580104000            call dword ptr [00401080]
'007109ad    8d4db0                  lea ecx, dword ptr [ebp-50]

' *** Reference to "__vbaFreeVar"
'007109b0    ff1528104000            call dword ptr [00401028]
'#Cleanup(var_6)
'007109b6    8b16                    mov edx, dword ptr [esi]
'007109b8    56                      push esi
'007109b9    ff9218030000            call dword ptr [edx+00000318]
'007109bf    8945b8                  mov dword ptr [ebp-48], eax
'007109c2    8d45b0                  lea eax, dword ptr [ebp-50]
'007109c5    50                      push eax
'007109c6    c745b009000000          mov dword ptr [ebp-50], 00000009

' *** Reference to "rtcIsNull"
'007109cd    ff1540114000            call dword ptr [00401140]
'007109d3    8b0e                    mov ecx, dword ptr [esi]
'007109d5    56                      push esi
'007109d6    89859cfeffff            mov dword ptr [ebp+fffffe9c], eax
'007109dc    ff9118030000            call dword ptr [ecx+00000318]

' *** Reference to "__vbaObjSet"
'007109e2    8b1db4104000            mov ebx, dword ptr [004010b4]
'007109e8    50                      push eax
'007109e9    8d55c8                  lea edx, dword ptr [ebp-38]
'007109ec    52                      push edx
'007109ed    ffd3                    call ebx
Set var_46 = IsNull(var_73)
'007109ef    8d4dd4                  lea ecx, dword ptr [ebp-2c]
'007109f2    8bf8                    mov edi, eax
'007109f4    8b07                    mov eax, dword ptr [edi]
'007109f6    51                      push ecx
'007109f7    57                      push edi
'007109f8    ff90a0000000            call dword ptr [eax+000000a0]
'007109fe    dbe2                    fnclex
'00710a00    85c0                    test ax, ax
'00710a02    7d12                    jge 710a16
'00710a04    68a0000000              push 000000a0
'00710a09    68d00d4300              push 00430dd0
'00710a0e    57                      push edi
'00710a0f    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00710a10    ff1580104000            call dword ptr [00401080]
'00710a16    8b16                    mov edx, dword ptr [esi]
'00710a18    56                      push esi
'00710a19    ff9204030000            call dword ptr [edx+00000304]
'00710a1f    50                      push eax
'00710a20    8d45c0                  lea eax, dword ptr [ebp-40]
'00710a23    50                      push eax
'00710a24    ffd3                    call ebx
Set var_5 = Nothing
'00710a26    8d55cc                  lea edx, dword ptr [ebp-34]
'00710a29    8bf8                    mov edi, eax
'00710a2b    8b0f                    mov ecx, dword ptr [edi]
'00710a2d    52                      push edx
'00710a2e    57                      push edi
'00710a2f    ff91a8000000            call dword ptr [ecx+000000a8]
'00710a35    dbe2                    fnclex
'00710a37    85c0                    test ax, ax
'00710a39    7d12                    jge 710a4d
'00710a3b    68a8000000              push 000000a8
'00710a40    681cb94300              push 0043b91c
'00710a45    57                      push edi
'00710a46    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00710a47    ff1580104000            call dword ptr [00401080]
'00710a4d    8b45cc                  mov eax, dword ptr [ebp-34]
'00710a50    898578ffffff            mov dword ptr [ebp+ffffff78], eax
'00710a56    8b06                    mov eax, dword ptr [esi]
'00710a58    56                      push esi
'00710a59    c745cc00000000          mov dword ptr [ebp-34], 00000000
'00710a60    c78570ffffff08000000    mov dword ptr [ebp+ffffff70], 00000008
'00710a6a    ff9018030000            call dword ptr [eax+00000318]
'00710a70    50                      push eax
'00710a71    8d4dc4                  lea ecx, dword ptr [ebp-3c]
'00710a74    51                      push ecx
'00710a75    ffd3                    call ebx
Set var_9 = Nothing
'00710a77    8bf8                    mov edi, eax
'00710a79    8b17                    mov edx, dword ptr [edi]
'00710a7b    8d45d0                  lea eax, dword ptr [ebp-30]
'00710a7e    50                      push eax
'00710a7f    57                      push edi
'00710a80    ff92a0000000            call dword ptr [edx+000000a0]
'00710a86    dbe2                    fnclex
'00710a88    85c0                    test ax, ax
'00710a8a    7d12                    jge 710a9e
'00710a8c    68a0000000              push 000000a0
'00710a91    68d00d4300              push 00430dd0
'00710a96    57                      push edi
'00710a97    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00710a98    ff1580104000            call dword ptr [00401080]
'00710a9e    8b4dd0                  mov ecx, dword ptr [ebp-30]
'00710aa1    6890694500              push 00456990
'00710aa6    51                      push ecx

' *** Reference to "__vbaStrCat"
'00710aa7    ff1570104000            call dword ptr [00401070]
var_4 = (" p") & (vbNullString)
'00710aad    894598                  mov dword ptr [ebp-68], eax
'00710ab0    b808000000              mov eax, 00000008
'00710ab5    8d9540ffffff            lea edx, dword ptr [ebp+ffffff40]
'00710abb    8d4da0                  lea ecx, dword ptr [ebp-60]
'00710abe    894590                  mov dword ptr [ebp-70], eax
'00710ac1    c78548ffffffcc134300    mov dword ptr [ebp+ffffff48], 004313cc
'00710acb    898540ffffff            mov dword ptr [ebp+ffffff40], eax

' *** Reference to "__vbaVarDup"
'00710ad1    ff1598124000            call dword ptr [00401298]
var_7 = (vbNullChar)
'00710ad7    8b55d4                  mov edx, dword ptr [ebp-2c]
'00710ada    52                      push edx
'00710adb    68cc134300              push 004313cc

' *** Reference to "__vbaStrCmp"
'00710ae0    ff1530114000            call dword ptr [00401130]
'00710ae6    8b959cfeffff            mov edx, dword ptr [ebp+fffffe9c]
'00710aec    f7d8                    neg eax
'00710aee    1bc0                    sbb eax, eax
'00710af0    40                      inc eax
'00710af1    f7d8                    neg eax
'00710af3    0bc2                    or eax, edx
var_num1 = ((vbNullString) = (vbNullChar)) Or IsNull(var_73)
'00710af5    66898558ffffff          mov word ptr [ebp+ffffff58], ax
'00710afc    8d4590                  lea eax, dword ptr [ebp-70]
'00710aff    50                      push eax
'00710b00    8d4da0                  lea ecx, dword ptr [ebp-60]
'00710b03    51                      push ecx
'00710b04    8d9550ffffff            lea edx, dword ptr [ebp+ffffff50]
'00710b0a    52                      push edx
'00710b0b    8d4580                  lea eax, dword ptr [ebp-80]
'00710b0e    50                      push eax
'00710b0f    c78550ffffff0b000000    mov dword ptr [ebp+ffffff50], 0000000b

' *** Reference to "rtcImmediateIf"
'00710b19    ff154c124000            call dword ptr [0040124c]
'00710b1f    8b8d78feffff            mov ecx, dword ptr [ebp+fffffe78]
'00710b25    8b19                    mov ebx, dword ptr [ecx]
'00710b27    8d9570ffffff            lea edx, dword ptr [ebp+ffffff70]
'00710b2d    52                      push edx
'00710b2e    8d4580                  lea eax, dword ptr [ebp-80]
'00710b31    50                      push eax
'00710b32    8d8d60ffffff            lea ecx, dword ptr [ebp+ffffff60]
'00710b38    51                      push ecx
'00710b39    bfbcb24300              mov edi, 0043b2bc

' *** Reference to "__vbaVarCat"
'00710b3e    ff1508124000            call dword ptr [00401208]
'00710b44    8b08                    mov ecx, dword ptr [eax]
'00710b46    83ec10                  sub esp, 10
'00710b49    8bd4                    mov edx, esp
'00710b4b    890a                    mov dword ptr [edx], ecx
'00710b4d    8b4804                  mov ecx, dword ptr [eax+04]
'00710b50    894a04                  mov dword ptr [edx+04], ecx
'00710b53    8b4808                  mov ecx, dword ptr [eax+08]
'00710b56    8b400c                  mov eax, dword ptr [eax+0c]
'00710b59    894a08                  mov dword ptr [edx+08], ecx
'00710b5c    89420c                  mov dword ptr [edx+0c], eax
'00710b5f    8b9534ffffff            mov edx, dword ptr [ebp+ffffff34]
'00710b65    83ec10                  sub esp, 10
'00710b68    8bcc                    mov ecx, esp
'00710b6a    b808000000              mov eax, 00000008
'00710b6f    8901                    mov dword ptr [ecx], eax
'00710b71    8b853cffffff            mov eax, dword ptr [ebp+ffffff3c]
'00710b77    895104                  mov dword ptr [ecx+04], edx
'00710b7a    897908                  mov dword ptr [ecx+08], edi
'00710b7d    89410c                  mov dword ptr [ecx+0c], eax
'00710b80    8b8d78feffff            mov ecx, dword ptr [ebp+fffffe78]
'00710b86    51                      push ecx
'00710b87    ff9328010000            call dword ptr [ebx+00000128]
'00710b8d    dbe2                    fnclex
'00710b8f    85c0                    test ax, ax
'00710b91    7d18                    jge 710bab
'00710b93    8b9578feffff            mov edx, dword ptr [ebp+fffffe78]
'00710b99    6828010000              push 00000128
'00710b9e    6830314300              push 00433130
'00710ba3    52                      push edx
'00710ba4    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00710ba5    ff1580104000            call dword ptr [00401080]
'00710bab    8d45d0                  lea eax, dword ptr [ebp-30]
'00710bae    50                      push eax
'00710baf    8d4dd4                  lea ecx, dword ptr [ebp-2c]
'00710bb2    51                      push ecx
'00710bb3    6a02                    push 02

' *** Reference to "__vbaFreeStrList"
'00710bb5    ff155c124000            call dword ptr [0040125c]
'#Cleanup( , 0)
'00710bbb    8d55c0                  lea edx, dword ptr [ebp-40]
'00710bbe    52                      push edx
'00710bbf    8d45c4                  lea eax, dword ptr [ebp-3c]
'00710bc2    50                      push eax
'00710bc3    8d4dc8                  lea ecx, dword ptr [ebp-38]
'00710bc6    51                      push ecx
'00710bc7    6a03                    push 03

' *** Reference to "__vbaFreeObjList"
'00710bc9    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_46, var_9, var_5)
'00710bcf    8d9560ffffff            lea edx, dword ptr [ebp+ffffff60]
'00710bd5    52                      push edx
'00710bd6    8d4580                  lea eax, dword ptr [ebp-80]
'00710bd9    50                      push eax
'00710bda    8d8d70ffffff            lea ecx, dword ptr [ebp+ffffff70]
'00710be0    51                      push ecx
'00710be1    8d5590                  lea edx, dword ptr [ebp-70]
'00710be4    52                      push edx
'00710be5    8d45a0                  lea eax, dword ptr [ebp-60]
'00710be8    50                      push eax
'00710be9    8d8d50ffffff            lea ecx, dword ptr [ebp+ffffff50]
'00710bef    51                      push ecx
'00710bf0    8d55b0                  lea edx, dword ptr [ebp-50]
'00710bf3    52                      push edx
'00710bf4    6a07                    push 07

' *** Reference to "__vbaFreeVarList"
'00710bf6    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_6, var_21, var_7, var_8, var_19, var_18, var_20)
'00710bfc    8b06                    mov eax, dword ptr [esi]
'00710bfe    83c43c                  add esp, 3c
'00710c01    56                      push esi
'00710c02    ff9034030000            call dword ptr [eax+00000334]
'00710c08    50                      push eax
'00710c09    8d4dc8                  lea ecx, dword ptr [ebp-38]
'00710c0c    51                      push ecx

' *** Reference to "__vbaObjSet"
'00710c0d    ff15b4104000            call dword ptr [004010b4]
Set var_46 = Nothing
'00710c13    8bf8                    mov edi, eax
'00710c15    8b17                    mov edx, dword ptr [edi]
'00710c17    8d45d4                  lea eax, dword ptr [ebp-2c]
'00710c1a    50                      push eax
'00710c1b    57                      push edi
'00710c1c    ff92a0000000            call dword ptr [edx+000000a0]
'00710c22    dbe2                    fnclex
'00710c24    85c0                    test ax, ax
'00710c26    7d12                    jge 710c3a
'00710c28    68a0000000              push 000000a0
'00710c2d    68d00d4300              push 00430dd0
'00710c32    57                      push edi
'00710c33    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00710c34    ff1580104000            call dword ptr [00401080]
'00710c3a    8b4dd4                  mov ecx, dword ptr [ebp-2c]
'00710c3d    51                      push ecx
'00710c3e    68cc134300              push 004313cc

' *** Reference to "__vbaStrCmp"
'00710c43    ff1530114000            call dword ptr [00401130]
'00710c49    8bf8                    mov edi, eax
'00710c4b    f7df                    neg edi
'00710c4d    1bff                    sbb edi, edi
'00710c4f    f7df                    neg edi
'00710c51    8d4dd4                  lea ecx, dword ptr [ebp-2c]
'00710c54    f7df                    neg edi

' *** Reference to "__vbaFreeStr"
'00710c56    ff150c134000            call dword ptr [0040130c]
'#Cleanup(var_44)
'00710c5c    8d4dc8                  lea ecx, dword ptr [ebp-38]

' *** Reference to "__vbaFreeObj"
'00710c5f    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_46)
'00710c65    6685ff                  test edi, edi
'00710c68    0f84a1000000            je 710d0f

If (((vbNullString) <> (vbNullChar))) Then
'00710c6e    8b16                    mov edx, dword ptr [esi]
'00710c70    56                      push esi
'00710c71    ff9234030000            call dword ptr [edx+00000334]
'00710c77    8bbd78feffff            mov edi, dword ptr [ebp+fffffe78]
'00710c7d    83ec10                  sub esp, 10
'00710c80    8bdc                    mov ebx, esp
'00710c82    b909000000              mov ecx, 00000009
'00710c87    890b                    mov dword ptr [ebx], ecx
'00710c89    894db0                  mov dword ptr [ebp-50], ecx
'00710c8c    8b4db4                  mov ecx, dword ptr [ebp-4c]
'00710c8f    894b04                  mov dword ptr [ebx+04], ecx
'00710c92    894308                  mov dword ptr [ebx+08], eax
'00710c95    8945b8                  mov dword ptr [ebp-48], eax
'00710c98    8b45bc                  mov eax, dword ptr [ebp-44]
'00710c9b    89430c                  mov dword ptr [ebx+0c], eax
'00710c9e    83ec10                  sub esp, 10
'00710ca1    c78550ffffff08000000    mov dword ptr [ebp+ffffff50], 00000008
'00710cab    8b8550ffffff            mov eax, dword ptr [ebp+ffffff50]
'00710cb1    8bcc                    mov ecx, esp
'00710cb3    8901                    mov dword ptr [ecx], eax
'00710cb5    8b8554ffffff            mov eax, dword ptr [ebp+ffffff54]
'00710cbb    894104                  mov dword ptr [ecx+04], eax
'00710cbe    8b8578feffff            mov eax, dword ptr [ebp+fffffe78]
'00710cc4    baa0b84300              mov edx, 0043b8a0
'00710cc9    899558ffffff            mov dword ptr [ebp+ffffff58], edx
'00710ccf    8b3f                    mov edi, dword ptr [edi]
'00710cd1    895108                  mov dword ptr [ecx+08], edx
'00710cd4    8b955cffffff            mov edx, dword ptr [ebp+ffffff5c]
'00710cda    50                      push eax
'00710cdb    89510c                  mov dword ptr [ecx+0c], edx
'00710cde    ff9728010000            call dword ptr [edi+00000128]
'00710ce4    dbe2                    fnclex
'00710ce6    85c0                    test ax, ax
'00710ce8    7d18                    jge 710d02
'00710cea    8b8d78feffff            mov ecx, dword ptr [ebp+fffffe78]
'00710cf0    6828010000              push 00000128
'00710cf5    6830314300              push 00433130
'00710cfa    51                      push ecx
'00710cfb    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00710cfc    ff1580104000            call dword ptr [00401080]

' *** Reference to "__vbaFreeVar"
'00710d02    8b1d28104000            mov ebx, dword ptr [00401028]
'00710d08    8d4db0                  lea ecx, dword ptr [ebp-50]
'00710d0b    ffd3                    call ebx
    '#Cleanup(var_6)
'00710d0d    eb06                    jmp 710d15
    
End If

' *** Reference to "__vbaFreeVar"
'00710d0f    8b1d28104000            mov ebx, dword ptr [00401028]
'00710d15    8b16                    mov edx, dword ptr [esi]
'00710d17    56                      push esi
'00710d18    ff9238030000            call dword ptr [edx+00000338]
'00710d1e    50                      push eax
'00710d1f    8d45c8                  lea eax, dword ptr [ebp-38]
'00710d22    50                      push eax

' *** Reference to "__vbaObjSet"
'00710d23    ff15b4104000            call dword ptr [004010b4]
Set var_46 = ((vbNullString) [##] (vbNullChar))
'00710d29    8d55d4                  lea edx, dword ptr [ebp-2c]
'00710d2c    8bf8                    mov edi, eax
'00710d2e    8b0f                    mov ecx, dword ptr [edi]
'00710d30    52                      push edx
'00710d31    57                      push edi
'00710d32    ff91a0000000            call dword ptr [ecx+000000a0]
'00710d38    dbe2                    fnclex
'00710d3a    85c0                    test ax, ax
'00710d3c    7d12                    jge 710d50
'00710d3e    68a0000000              push 000000a0
'00710d43    68d00d4300              push 00430dd0
'00710d48    57                      push edi
'00710d49    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00710d4a    ff1580104000            call dword ptr [00401080]
'00710d50    8b45d4                  mov eax, dword ptr [ebp-2c]
'00710d53    50                      push eax
'00710d54    68cc134300              push 004313cc

' *** Reference to "__vbaStrCmp"
'00710d59    ff1530114000            call dword ptr [00401130]
'00710d5f    8bf8                    mov edi, eax
'00710d61    f7df                    neg edi
'00710d63    1bff                    sbb edi, edi
'00710d65    f7df                    neg edi
'00710d67    8d4dd4                  lea ecx, dword ptr [ebp-2c]
'00710d6a    f7df                    neg edi

' *** Reference to "__vbaFreeStr"
'00710d6c    ff150c134000            call dword ptr [0040130c]
'#Cleanup(var_44)
'00710d72    8d4dc8                  lea ecx, dword ptr [ebp-38]

' *** Reference to "__vbaFreeObj"
'00710d75    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_46)
'00710d7b    6685ff                  test edi, edi
'00710d7e    0f849e000000            je 710e22

If (((vbNullString) <> (vbNullChar))) Then
'00710d84    8b0e                    mov ecx, dword ptr [esi]
'00710d86    56                      push esi
'00710d87    ff9138030000            call dword ptr [ecx+00000338]
'00710d8d    8b9578feffff            mov edx, dword ptr [ebp+fffffe78]
'00710d93    83ec10                  sub esp, 10
'00710d96    8bfc                    mov edi, esp
'00710d98    b909000000              mov ecx, 00000009
'00710d9d    890f                    mov dword ptr [edi], ecx
'00710d9f    894db0                  mov dword ptr [ebp-50], ecx
'00710da2    8b4db4                  mov ecx, dword ptr [ebp-4c]
'00710da5    894f04                  mov dword ptr [edi+04], ecx
'00710da8    894708                  mov dword ptr [edi+08], eax
'00710dab    8945b8                  mov dword ptr [ebp-48], eax
'00710dae    8b45bc                  mov eax, dword ptr [ebp-44]
'00710db1    89470c                  mov dword ptr [edi+0c], eax
'00710db4    83ec10                  sub esp, 10
'00710db7    c78550ffffff08000000    mov dword ptr [ebp+ffffff50], 00000008
'00710dc1    8b8550ffffff            mov eax, dword ptr [ebp+ffffff50]
'00710dc7    8bcc                    mov ecx, esp
'00710dc9    8901                    mov dword ptr [ecx], eax
'00710dcb    8b8554ffffff            mov eax, dword ptr [ebp+ffffff54]
'00710dd1    894104                  mov dword ptr [ecx+04], eax
'00710dd4    c78558ffffffe4b24300    mov dword ptr [ebp+ffffff58], 0043b2e4
'00710dde    8b8558ffffff            mov eax, dword ptr [ebp+ffffff58]
'00710de4    8b12                    mov edx, dword ptr [edx]
'00710de6    894108                  mov dword ptr [ecx+08], eax
'00710de9    8b855cffffff            mov eax, dword ptr [ebp+ffffff5c]
'00710def    89410c                  mov dword ptr [ecx+0c], eax
'00710df2    8b8d78feffff            mov ecx, dword ptr [ebp+fffffe78]
'00710df8    51                      push ecx
'00710df9    ff9228010000            call dword ptr [edx+00000128]
'00710dff    dbe2                    fnclex
'00710e01    85c0                    test ax, ax
'00710e03    7d18                    jge 710e1d
'00710e05    8b9578feffff            mov edx, dword ptr [ebp+fffffe78]
'00710e0b    6828010000              push 00000128
'00710e10    6830314300              push 00433130
'00710e15    52                      push edx
'00710e16    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00710e17    ff1580104000            call dword ptr [00401080]
'00710e1d    8d4db0                  lea ecx, dword ptr [ebp-50]
'00710e20    ffd3                    call ebx
    '#Cleanup(var_6)
End If
'00710e22    8b06                    mov eax, dword ptr [esi]
'00710e24    56                      push esi
'00710e25    ff90fc020000            call dword ptr [eax+000002fc]
'00710e2b    50                      push eax
'00710e2c    8d4dc8                  lea ecx, dword ptr [ebp-38]
'00710e2f    51                      push ecx

' *** Reference to "__vbaObjSet"
'00710e30    ff15b4104000            call dword ptr [004010b4]
Set var_46 = Nothing
'00710e36    8bf8                    mov edi, eax
'00710e38    8b17                    mov edx, dword ptr [edi]
'00710e3a    8d45d4                  lea eax, dword ptr [ebp-2c]
'00710e3d    50                      push eax
'00710e3e    57                      push edi
'00710e3f    ff92a8000000            call dword ptr [edx+000000a8]
'00710e45    dbe2                    fnclex
'00710e47    85c0                    test ax, ax
'00710e49    7d12                    jge 710e5d
'00710e4b    68a8000000              push 000000a8
'00710e50    681cb94300              push 0043b91c
'00710e55    57                      push edi
'00710e56    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00710e57    ff1580104000            call dword ptr [00401080]
'00710e5d    8b4dd4                  mov ecx, dword ptr [ebp-2c]
'00710e60    51                      push ecx
'00710e61    68cc134300              push 004313cc

' *** Reference to "__vbaStrCmp"
'00710e66    ff1530114000            call dword ptr [00401130]
'00710e6c    8bf8                    mov edi, eax
'00710e6e    f7df                    neg edi
'00710e70    1bff                    sbb edi, edi
'00710e72    f7df                    neg edi
'00710e74    8d4dd4                  lea ecx, dword ptr [ebp-2c]
'00710e77    f7df                    neg edi

' *** Reference to "__vbaFreeStr"
'00710e79    ff150c134000            call dword ptr [0040130c]
'#Cleanup(var_44)
'00710e7f    8d4dc8                  lea ecx, dword ptr [ebp-38]

' *** Reference to "__vbaFreeObj"
'00710e82    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_46)
'00710e88    6685ff                  test edi, edi
'00710e8b    0f849e000000            je 710f2f

If (((vbNullString) <> (vbNullChar))) Then
'00710e91    8b16                    mov edx, dword ptr [esi]
'00710e93    56                      push esi
'00710e94    ff92fc020000            call dword ptr [edx+000002fc]
'00710e9a    8b9578feffff            mov edx, dword ptr [ebp+fffffe78]
'00710ea0    83ec10                  sub esp, 10
'00710ea3    8bfc                    mov edi, esp
'00710ea5    b909000000              mov ecx, 00000009
'00710eaa    890f                    mov dword ptr [edi], ecx
'00710eac    894db0                  mov dword ptr [ebp-50], ecx
'00710eaf    8b4db4                  mov ecx, dword ptr [ebp-4c]
'00710eb2    894f04                  mov dword ptr [edi+04], ecx
'00710eb5    894708                  mov dword ptr [edi+08], eax
'00710eb8    8945b8                  mov dword ptr [ebp-48], eax
'00710ebb    8b45bc                  mov eax, dword ptr [ebp-44]
'00710ebe    89470c                  mov dword ptr [edi+0c], eax
'00710ec1    83ec10                  sub esp, 10
'00710ec4    c78550ffffff08000000    mov dword ptr [ebp+ffffff50], 00000008
'00710ece    8b8550ffffff            mov eax, dword ptr [ebp+ffffff50]
'00710ed4    8bcc                    mov ecx, esp
'00710ed6    8901                    mov dword ptr [ecx], eax
'00710ed8    8b8554ffffff            mov eax, dword ptr [ebp+ffffff54]
'00710ede    894104                  mov dword ptr [ecx+04], eax
'00710ee1    c78558ffffff00b34300    mov dword ptr [ebp+ffffff58], 0043b300
'00710eeb    8b8558ffffff            mov eax, dword ptr [ebp+ffffff58]
'00710ef1    8b12                    mov edx, dword ptr [edx]
'00710ef3    894108                  mov dword ptr [ecx+08], eax
'00710ef6    8b855cffffff            mov eax, dword ptr [ebp+ffffff5c]
'00710efc    89410c                  mov dword ptr [ecx+0c], eax
'00710eff    8b8d78feffff            mov ecx, dword ptr [ebp+fffffe78]
'00710f05    51                      push ecx
'00710f06    ff9228010000            call dword ptr [edx+00000128]
'00710f0c    dbe2                    fnclex
'00710f0e    85c0                    test ax, ax
'00710f10    7d18                    jge 710f2a
'00710f12    8b9578feffff            mov edx, dword ptr [ebp+fffffe78]
'00710f18    6828010000              push 00000128
'00710f1d    6830314300              push 00433130
'00710f22    52                      push edx
'00710f23    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00710f24    ff1580104000            call dword ptr [00401080]
'00710f2a    8d4db0                  lea ecx, dword ptr [ebp-50]
'00710f2d    ffd3                    call ebx
    '#Cleanup(var_6)
End If
'00710f2f    8b06                    mov eax, dword ptr [esi]
'00710f31    56                      push esi
'00710f32    ff900c030000            call dword ptr [eax+0000030c]
'00710f38    8b9578feffff            mov edx, dword ptr [ebp+fffffe78]
'00710f3e    83ec10                  sub esp, 10
'00710f41    8bfc                    mov edi, esp
'00710f43    b909000000              mov ecx, 00000009
'00710f48    890f                    mov dword ptr [edi], ecx
'00710f4a    894db0                  mov dword ptr [ebp-50], ecx
'00710f4d    8b4db4                  mov ecx, dword ptr [ebp-4c]
'00710f50    894f04                  mov dword ptr [edi+04], ecx
'00710f53    894708                  mov dword ptr [edi+08], eax
'00710f56    8945b8                  mov dword ptr [ebp-48], eax
'00710f59    8b45bc                  mov eax, dword ptr [ebp-44]
'00710f5c    89470c                  mov dword ptr [edi+0c], eax
'00710f5f    83ec10                  sub esp, 10
'00710f62    c78550ffffff08000000    mov dword ptr [ebp+ffffff50], 00000008
'00710f6c    8b8550ffffff            mov eax, dword ptr [ebp+ffffff50]
'00710f72    8bcc                    mov ecx, esp
'00710f74    8901                    mov dword ptr [ecx], eax
'00710f76    8b8554ffffff            mov eax, dword ptr [ebp+ffffff54]
'00710f7c    894104                  mov dword ptr [ecx+04], eax
'00710f7f    c78558ffffffb4b84300    mov dword ptr [ebp+ffffff58], 0043b8b4
'00710f89    8b8558ffffff            mov eax, dword ptr [ebp+ffffff58]
'00710f8f    8b12                    mov edx, dword ptr [edx]
'00710f91    894108                  mov dword ptr [ecx+08], eax
'00710f94    8b855cffffff            mov eax, dword ptr [ebp+ffffff5c]
'00710f9a    89410c                  mov dword ptr [ecx+0c], eax
'00710f9d    8b8d78feffff            mov ecx, dword ptr [ebp+fffffe78]
'00710fa3    51                      push ecx
'00710fa4    ff9228010000            call dword ptr [edx+00000128]
'00710faa    dbe2                    fnclex
'00710fac    85c0                    test ax, ax
'00710fae    7d18                    jge 710fc8
'00710fb0    8b9578feffff            mov edx, dword ptr [ebp+fffffe78]
'00710fb6    6828010000              push 00000128
'00710fbb    6830314300              push 00433130
'00710fc0    52                      push edx
'00710fc1    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00710fc2    ff1580104000            call dword ptr [00401080]
'00710fc8    8d4db0                  lea ecx, dword ptr [ebp-50]
'00710fcb    ffd3                    call ebx
'#Cleanup(var_6)
'00710fcd    8b06                    mov eax, dword ptr [esi]
'00710fcf    56                      push esi
'00710fd0    ff9044030000            call dword ptr [eax+00000344]
'00710fd6    8b9578feffff            mov edx, dword ptr [ebp+fffffe78]
'00710fdc    83ec10                  sub esp, 10
'00710fdf    8bfc                    mov edi, esp
'00710fe1    b909000000              mov ecx, 00000009
'00710fe6    890f                    mov dword ptr [edi], ecx
'00710fe8    894db0                  mov dword ptr [ebp-50], ecx
'00710feb    8b4db4                  mov ecx, dword ptr [ebp-4c]
'00710fee    894f04                  mov dword ptr [edi+04], ecx
'00710ff1    894708                  mov dword ptr [edi+08], eax
'00710ff4    8945b8                  mov dword ptr [ebp-48], eax
'00710ff7    8b45bc                  mov eax, dword ptr [ebp-44]
'00710ffa    89470c                  mov dword ptr [edi+0c], eax
'00710ffd    83ec10                  sub esp, 10
'00711000    c78550ffffff08000000    mov dword ptr [ebp+ffffff50], 00000008
'0071100a    8b8550ffffff            mov eax, dword ptr [ebp+ffffff50]
'00711010    8bcc                    mov ecx, esp
'00711012    8901                    mov dword ptr [ecx], eax
'00711014    8b8554ffffff            mov eax, dword ptr [ebp+ffffff54]
'0071101a    894104                  mov dword ptr [ecx+04], eax
'0071101d    c78558ffffffc4b84300    mov dword ptr [ebp+ffffff58], 0043b8c4
'00711027    8b8558ffffff            mov eax, dword ptr [ebp+ffffff58]
'0071102d    8b12                    mov edx, dword ptr [edx]
'0071102f    894108                  mov dword ptr [ecx+08], eax
'00711032    8b855cffffff            mov eax, dword ptr [ebp+ffffff5c]
'00711038    89410c                  mov dword ptr [ecx+0c], eax
'0071103b    8b8d78feffff            mov ecx, dword ptr [ebp+fffffe78]
'00711041    51                      push ecx
'00711042    ff9228010000            call dword ptr [edx+00000128]
'00711048    dbe2                    fnclex
'0071104a    85c0                    test ax, ax
'0071104c    7d18                    jge 711066
'0071104e    8b9578feffff            mov edx, dword ptr [ebp+fffffe78]
'00711054    6828010000              push 00000128
'00711059    6830314300              push 00433130
'0071105e    52                      push edx
'0071105f    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00711060    ff1580104000            call dword ptr [00401080]
'00711066    8d4db0                  lea ecx, dword ptr [ebp-50]
'00711069    ffd3                    call ebx
'#Cleanup(var_6)
'0071106b    8b06                    mov eax, dword ptr [esi]
'0071106d    56                      push esi
'0071106e    ff9010030000            call dword ptr [eax+00000310]
'00711074    8b9578feffff            mov edx, dword ptr [ebp+fffffe78]
'0071107a    83ec10                  sub esp, 10
'0071107d    8bfc                    mov edi, esp
'0071107f    b909000000              mov ecx, 00000009
'00711084    890f                    mov dword ptr [edi], ecx
'00711086    894db0                  mov dword ptr [ebp-50], ecx
'00711089    8b4db4                  mov ecx, dword ptr [ebp-4c]
'0071108c    894f04                  mov dword ptr [edi+04], ecx
'0071108f    894708                  mov dword ptr [edi+08], eax
'00711092    8945b8                  mov dword ptr [ebp-48], eax
'00711095    8b45bc                  mov eax, dword ptr [ebp-44]
'00711098    89470c                  mov dword ptr [edi+0c], eax
'0071109b    83ec10                  sub esp, 10
'0071109e    c78550ffffff08000000    mov dword ptr [ebp+ffffff50], 00000008
'007110a8    8b8550ffffff            mov eax, dword ptr [ebp+ffffff50]
'007110ae    8bcc                    mov ecx, esp
'007110b0    8901                    mov dword ptr [ecx], eax
'007110b2    8b8554ffffff            mov eax, dword ptr [ebp+ffffff54]
'007110b8    894104                  mov dword ptr [ecx+04], eax
'007110bb    c78558ffffff6ca44300    mov dword ptr [ebp+ffffff58], 0043a46c
'007110c5    8b8558ffffff            mov eax, dword ptr [ebp+ffffff58]
'007110cb    8b12                    mov edx, dword ptr [edx]
'007110cd    894108                  mov dword ptr [ecx+08], eax
'007110d0    8b855cffffff            mov eax, dword ptr [ebp+ffffff5c]
'007110d6    89410c                  mov dword ptr [ecx+0c], eax
'007110d9    8b8d78feffff            mov ecx, dword ptr [ebp+fffffe78]
'007110df    51                      push ecx
'007110e0    ff9228010000            call dword ptr [edx+00000128]
'007110e6    dbe2                    fnclex
'007110e8    85c0                    test ax, ax
'007110ea    7d18                    jge 711104
'007110ec    8b9578feffff            mov edx, dword ptr [ebp+fffffe78]
'007110f2    6828010000              push 00000128
'007110f7    6830314300              push 00433130
'007110fc    52                      push edx
'007110fd    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'007110fe    ff1580104000            call dword ptr [00401080]
'00711104    8d4db0                  lea ecx, dword ptr [ebp-50]
'00711107    ffd3                    call ebx
'#Cleanup(var_6)
'00711109    8b06                    mov eax, dword ptr [esi]
'0071110b    56                      push esi
'0071110c    ff9014030000            call dword ptr [eax+00000314]
'00711112    8b9578feffff            mov edx, dword ptr [ebp+fffffe78]
'00711118    83ec10                  sub esp, 10
'0071111b    8bfc                    mov edi, esp
'0071111d    b909000000              mov ecx, 00000009
'00711122    890f                    mov dword ptr [edi], ecx
'00711124    894db0                  mov dword ptr [ebp-50], ecx
'00711127    8b4db4                  mov ecx, dword ptr [ebp-4c]
'0071112a    894f04                  mov dword ptr [edi+04], ecx
'0071112d    894708                  mov dword ptr [edi+08], eax
'00711130    8945b8                  mov dword ptr [ebp-48], eax
'00711133    8b45bc                  mov eax, dword ptr [ebp-44]
'00711136    89470c                  mov dword ptr [edi+0c], eax
'00711139    83ec10                  sub esp, 10
'0071113c    c78550ffffff08000000    mov dword ptr [ebp+ffffff50], 00000008
'00711146    8b8550ffffff            mov eax, dword ptr [ebp+ffffff50]
'0071114c    8bcc                    mov ecx, esp
'0071114e    8901                    mov dword ptr [ecx], eax
'00711150    8b8554ffffff            mov eax, dword ptr [ebp+ffffff54]
'00711156    894104                  mov dword ptr [ecx+04], eax
'00711159    c78558ffffffd8b84300    mov dword ptr [ebp+ffffff58], 0043b8d8
'00711163    8b8558ffffff            mov eax, dword ptr [ebp+ffffff58]
'00711169    8b12                    mov edx, dword ptr [edx]
'0071116b    894108                  mov dword ptr [ecx+08], eax
'0071116e    8b855cffffff            mov eax, dword ptr [ebp+ffffff5c]
'00711174    89410c                  mov dword ptr [ecx+0c], eax
'00711177    8b8d78feffff            mov ecx, dword ptr [ebp+fffffe78]
'0071117d    51                      push ecx
'0071117e    ff9228010000            call dword ptr [edx+00000128]
'00711184    dbe2                    fnclex
'00711186    85c0                    test ax, ax
'00711188    7d18                    jge 7111a2
'0071118a    8b9578feffff            mov edx, dword ptr [ebp+fffffe78]
'00711190    6828010000              push 00000128
'00711195    6830314300              push 00433130
'0071119a    52                      push edx
'0071119b    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0071119c    ff1580104000            call dword ptr [00401080]
'007111a2    8d4db0                  lea ecx, dword ptr [ebp-50]
'007111a5    ffd3                    call ebx
'#Cleanup(var_6)
'007111a7    8b06                    mov eax, dword ptr [esi]
'007111a9    56                      push esi
'007111aa    ff9048030000            call dword ptr [eax+00000348]
'007111b0    8b9578feffff            mov edx, dword ptr [ebp+fffffe78]
'007111b6    83ec10                  sub esp, 10
'007111b9    8bfc                    mov edi, esp
'007111bb    b909000000              mov ecx, 00000009
'007111c0    890f                    mov dword ptr [edi], ecx
'007111c2    894db0                  mov dword ptr [ebp-50], ecx
'007111c5    8b4db4                  mov ecx, dword ptr [ebp-4c]
'007111c8    894f04                  mov dword ptr [edi+04], ecx
'007111cb    894708                  mov dword ptr [edi+08], eax
'007111ce    8945b8                  mov dword ptr [ebp-48], eax
'007111d1    8b45bc                  mov eax, dword ptr [ebp-44]
'007111d4    89470c                  mov dword ptr [edi+0c], eax
'007111d7    83ec10                  sub esp, 10
'007111da    c78550ffffff08000000    mov dword ptr [ebp+ffffff50], 00000008
'007111e4    8b8550ffffff            mov eax, dword ptr [ebp+ffffff50]
'007111ea    8bcc                    mov ecx, esp
'007111ec    8901                    mov dword ptr [ecx], eax
'007111ee    8b8554ffffff            mov eax, dword ptr [ebp+ffffff54]
'007111f4    894104                  mov dword ptr [ecx+04], eax
'007111f7    c78558ffffffecb84300    mov dword ptr [ebp+ffffff58], 0043b8ec
'00711201    8b8558ffffff            mov eax, dword ptr [ebp+ffffff58]
'00711207    8b12                    mov edx, dword ptr [edx]
'00711209    894108                  mov dword ptr [ecx+08], eax
'0071120c    8b855cffffff            mov eax, dword ptr [ebp+ffffff5c]
'00711212    89410c                  mov dword ptr [ecx+0c], eax
'00711215    8b8d78feffff            mov ecx, dword ptr [ebp+fffffe78]
'0071121b    51                      push ecx
'0071121c    ff9228010000            call dword ptr [edx+00000128]
'00711222    dbe2                    fnclex
'00711224    85c0                    test ax, ax
'00711226    7d18                    jge 711240
'00711228    8b9578feffff            mov edx, dword ptr [ebp+fffffe78]
'0071122e    6828010000              push 00000128
'00711233    6830314300              push 00433130
'00711238    52                      push edx
'00711239    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0071123a    ff1580104000            call dword ptr [00401080]
'00711240    8d4db0                  lea ecx, dword ptr [ebp-50]
'00711243    ffd3                    call ebx
'#Cleanup(var_6)
'00711245    8b06                    mov eax, dword ptr [esi]
'00711247    56                      push esi
'00711248    ff903c030000            call dword ptr [eax+0000033c]
'0071124e    8b9578feffff            mov edx, dword ptr [ebp+fffffe78]
'00711254    83ec10                  sub esp, 10
'00711257    8bfc                    mov edi, esp
'00711259    b909000000              mov ecx, 00000009
'0071125e    890f                    mov dword ptr [edi], ecx
'00711260    894db0                  mov dword ptr [ebp-50], ecx
'00711263    8b4db4                  mov ecx, dword ptr [ebp-4c]
'00711266    894f04                  mov dword ptr [edi+04], ecx
'00711269    894708                  mov dword ptr [edi+08], eax
'0071126c    8945b8                  mov dword ptr [ebp-48], eax
'0071126f    8b45bc                  mov eax, dword ptr [ebp-44]
'00711272    89470c                  mov dword ptr [edi+0c], eax
'00711275    83ec10                  sub esp, 10
'00711278    c78550ffffff08000000    mov dword ptr [ebp+ffffff50], 00000008
'00711282    8b8550ffffff            mov eax, dword ptr [ebp+ffffff50]
'00711288    8bcc                    mov ecx, esp
'0071128a    8901                    mov dword ptr [ecx], eax
'0071128c    8b8554ffffff            mov eax, dword ptr [ebp+ffffff54]
'00711292    894104                  mov dword ptr [ecx+04], eax
'00711295    c78558ffffff00b94300    mov dword ptr [ebp+ffffff58], 0043b900
'0071129f    8b8558ffffff            mov eax, dword ptr [ebp+ffffff58]
'007112a5    8b12                    mov edx, dword ptr [edx]
'007112a7    894108                  mov dword ptr [ecx+08], eax
'007112aa    8b855cffffff            mov eax, dword ptr [ebp+ffffff5c]
'007112b0    89410c                  mov dword ptr [ecx+0c], eax
'007112b3    8b8d78feffff            mov ecx, dword ptr [ebp+fffffe78]
'007112b9    51                      push ecx
'007112ba    ff9228010000            call dword ptr [edx+00000128]
'007112c0    dbe2                    fnclex
'007112c2    85c0                    test ax, ax
'007112c4    7d18                    jge 7112de
'007112c6    8b9578feffff            mov edx, dword ptr [ebp+fffffe78]
'007112cc    6828010000              push 00000128
'007112d1    6830314300              push 00433130
'007112d6    52                      push edx
'007112d7    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'007112d8    ff1580104000            call dword ptr [00401080]
'007112de    8d4db0                  lea ecx, dword ptr [ebp-50]
'007112e1    ffd3                    call ebx
'#Cleanup(var_6)
'007112e3    8b06                    mov eax, dword ptr [esi]
'007112e5    56                      push esi
'007112e6    ff9040030000            call dword ptr [eax+00000340]
'007112ec    8b9578feffff            mov edx, dword ptr [ebp+fffffe78]
'007112f2    83ec10                  sub esp, 10
'007112f5    8bfc                    mov edi, esp
'007112f7    b909000000              mov ecx, 00000009
'007112fc    890f                    mov dword ptr [edi], ecx
'007112fe    894db0                  mov dword ptr [ebp-50], ecx
'00711301    8b4db4                  mov ecx, dword ptr [ebp-4c]
'00711304    894f04                  mov dword ptr [edi+04], ecx
'00711307    894708                  mov dword ptr [edi+08], eax
'0071130a    8945b8                  mov dword ptr [ebp-48], eax
'0071130d    8b45bc                  mov eax, dword ptr [ebp-44]
'00711310    89470c                  mov dword ptr [edi+0c], eax
'00711313    83ec10                  sub esp, 10
'00711316    c78550ffffff08000000    mov dword ptr [ebp+ffffff50], 00000008
'00711320    8b8550ffffff            mov eax, dword ptr [ebp+ffffff50]
'00711326    8bcc                    mov ecx, esp
'00711328    8901                    mov dword ptr [ecx], eax
'0071132a    8b8554ffffff            mov eax, dword ptr [ebp+ffffff54]
'00711330    894104                  mov dword ptr [ecx+04], eax
'00711333    c78558ffffff10b94300    mov dword ptr [ebp+ffffff58], 0043b910
'0071133d    8b8558ffffff            mov eax, dword ptr [ebp+ffffff58]
'00711343    8b12                    mov edx, dword ptr [edx]
'00711345    894108                  mov dword ptr [ecx+08], eax
'00711348    8b855cffffff            mov eax, dword ptr [ebp+ffffff5c]
'0071134e    89410c                  mov dword ptr [ecx+0c], eax
'00711351    8b8d78feffff            mov ecx, dword ptr [ebp+fffffe78]
'00711357    51                      push ecx
'00711358    ff9228010000            call dword ptr [edx+00000128]
'0071135e    dbe2                    fnclex
'00711360    85c0                    test ax, ax
'00711362    7d1c                    jge 711380
'00711364    8b9578feffff            mov edx, dword ptr [ebp+fffffe78]

' *** Reference to "__vbaHresultCheckObj"
'0071136a    8b3d80104000            mov edi, dword ptr [00401080]
'00711370    6828010000              push 00000128
'00711375    6830314300              push 00433130
'0071137a    52                      push edx
'0071137b    50                      push eax
'0071137c    ffd7                    call edi
'0071137e    eb06                    jmp 711386

' *** Reference to "__vbaHresultCheckObj"
'00711380    8b3d80104000            mov edi, dword ptr [00401080]
'00711386    8d4db0                  lea ecx, dword ptr [ebp-50]
'00711389    ffd3                    call ebx
'#Cleanup(var_6)
'0071138b    8b8578feffff            mov eax, dword ptr [ebp+fffffe78]
'00711391    8b08                    mov ecx, dword ptr [eax]
'00711393    6a00                    push 00
'00711395    6a01                    push 01
'00711397    50                      push eax
'00711398    ff9164010000            call dword ptr [ecx+00000164]
'0071139e    dbe2                    fnclex
'007113a0    85c0                    test ax, ax
'007113a2    7d14                    jge 7113b8
'007113a4    8b9578feffff            mov edx, dword ptr [ebp+fffffe78]
'007113aa    6864010000              push 00000164
'007113af    6830314300              push 00433130
'007113b4    52                      push edx
'007113b5    50                      push eax
'007113b6    ffd7                    call edi
'007113b8    6a00                    push 00
'007113ba    8d8578feffff            lea eax, dword ptr [ebp+fffffe78]
'007113c0    50                      push eax

' *** Reference to "__vbaObjSetAddref"
'007113c1    ff15c8104000            call dword ptr [004010c8]
Set var_73 = Nothing
'007113c7    8b45dc                  mov eax, dword ptr [ebp-24]
'007113ca    8b08                    mov ecx, dword ptr [eax]
'007113cc    50                      push eax
'007113cd    ff91c4000000            call dword ptr [ecx+000000c4]
'007113d3    dbe2                    fnclex
'007113d5    85c0                    test ax, ax
'007113d7    7d11                    jge 7113ea
'007113d9    8b55dc                  mov edx, dword ptr [ebp-24]
'007113dc    68c4000000              push 000000c4
'007113e1    6830314300              push 00433130
'007113e6    52                      push edx
'007113e7    50                      push eax
'007113e8    ffd7                    call edi
'007113ea    8b45e0                  mov eax, dword ptr [ebp-20]
'007113ed    8b08                    mov ecx, dword ptr [eax]
'007113ef    50                      push eax
'007113f0    ff91c4000000            call dword ptr [ecx+000000c4]
'007113f6    dbe2                    fnclex
'007113f8    85c0                    test ax, ax
'007113fa    7d11                    jge 71140d
'007113fc    8b55e0                  mov edx, dword ptr [ebp-20]
'007113ff    68c4000000              push 000000c4
'00711404    6830314300              push 00433130
'00711409    52                      push edx
'0071140a    50                      push eax
'0071140b    ffd7                    call edi
'0071140d    8b06                    mov eax, dword ptr [esi]
'0071140f    56                      push esi
'00711410    ff9008030000            call dword ptr [eax+00000308]
'00711416    50                      push eax
'00711417    8d8d74feffff            lea ecx, dword ptr [ebp+fffffe74]
'0071141d    51                      push ecx

' *** Reference to "__vbaObjSet"
'0071141e    ff15b4104000            call dword ptr [004010b4]
Set var_304 = Nothing
'00711424    8b8574feffff            mov eax, dword ptr [ebp+fffffe74]
'0071142a    8b10                    mov edx, dword ptr [eax]
'0071142c    50                      push eax
'0071142d    ff92e8010000            call dword ptr [edx+000001e8]
'00711433    dbe2                    fnclex
'00711435    85c0                    test ax, ax
'00711437    7d14                    jge 71144d
'00711439    8b8d74feffff            mov ecx, dword ptr [ebp+fffffe74]
'0071143f    68e8010000              push 000001e8
'00711444    681cb94300              push 0043b91c
'00711449    51                      push ecx
'0071144a    50                      push eax
'0071144b    ffd7                    call edi
'0071144d    8d5dc8                  lea ebx, dword ptr [ebp-38]
'00711450    53                      push ebx
'00711451    83ec10                  sub esp, 10
'00711454    8bdc                    mov ebx, esp
'00711456    b90a000000              mov ecx, 0000000a
'0071145b    890b                    mov dword ptr [ebx], ecx
'0071145d    898d40ffffff            mov dword ptr [ebp+ffffff40], ecx
'00711463    8b8d34ffffff            mov ecx, dword ptr [ebp+ffffff34]
'00711469    894b04                  mov dword ptr [ebx+04], ecx
'0071146c    8b3d48b07200            mov edi, dword ptr [0072b048]
'00711472    b804000280              mov eax, 80020004
'00711477    894308                  mov dword ptr [ebx+08], eax
'0071147a    8bd0                    mov edx, eax
'0071147c    8b853cffffff            mov eax, dword ptr [ebp+ffffff3c]
'00711482    89430c                  mov dword ptr [ebx+0c], eax
'00711485    8b8540ffffff            mov eax, dword ptr [ebp+ffffff40]
'0071148b    83ec10                  sub esp, 10
'0071148e    8bcc                    mov ecx, esp
'00711490    8901                    mov dword ptr [ecx], eax
'00711492    8b8544ffffff            mov eax, dword ptr [ebp+ffffff44]
'00711498    894104                  mov dword ptr [ecx+04], eax
'0071149b    895108                  mov dword ptr [ecx+08], edx
'0071149e    899548ffffff            mov dword ptr [ebp+ffffff48], edx
'007114a4    8b954cffffff            mov edx, dword ptr [ebp+ffffff4c]
'007114aa    89510c                  mov dword ptr [ecx+0c], edx
'007114ad    8b9554ffffff            mov edx, dword ptr [ebp+ffffff54]
'007114b3    83ec10                  sub esp, 10
'007114b6    c78550ffffff03000000    mov dword ptr [ebp+ffffff50], 00000003
'007114c0    8b8d50ffffff            mov ecx, dword ptr [ebp+ffffff50]
'007114c6    8bc4                    mov eax, esp
'007114c8    8908                    mov dword ptr [eax], ecx
'007114ca    895004                  mov dword ptr [eax+04], edx
'007114cd    8b955cffffff            mov edx, dword ptr [ebp+ffffff5c]
'007114d3    c78558ffffff02000000    mov dword ptr [ebp+ffffff58], 00000002
'007114dd    8b8d58ffffff            mov ecx, dword ptr [ebp+ffffff58]
'007114e3    8b3f                    mov edi, dword ptr [edi]
'007114e5    894808                  mov dword ptr [eax+08], ecx
'007114e8    89500c                  mov dword ptr [eax+0c], edx
'007114eb    a148b07200              mov ax, word ptr [0072b048]
'007114f0    689c694500              push 0045699c
'007114f5    50                      push eax
'007114f6    ff97bc000000            call dword ptr [edi+000000bc]
'007114fc    dbe2                    fnclex
'007114fe    85c0                    test ax, ax
'00711500    7d1c                    jge 71151e
'00711502    8b0d48b07200            mov ecx, dword ptr [0072b048]

' *** Reference to "__vbaHresultCheckObj"
'00711508    8b3d80104000            mov edi, dword ptr [00401080]
'0071150e    68bc000000              push 000000bc
'00711513    68301f4300              push 00431f30
'00711518    51                      push ecx
'00711519    50                      push eax
'0071151a    ffd7                    call edi
'0071151c    eb06                    jmp 711524

' *** Reference to "__vbaHresultCheckObj"
'0071151e    8b3d80104000            mov edi, dword ptr [00401080]
'00711524    8b45c8                  mov eax, dword ptr [ebp-38]
'00711527    50                      push eax
'00711528    8d55dc                  lea edx, dword ptr [ebp-24]
'0071152b    52                      push edx
'0071152c    c745c800000000          mov dword ptr [ebp-38], 00000000

' *** Reference to "__vbaObjSet"
'00711533    ff15b4104000            call dword ptr [004010b4]
Set var_39 = Nothing
'00711539    8b45dc                  mov eax, dword ptr [ebp-24]
'0071153c    8b08                    mov ecx, dword ptr [eax]
'0071153e    8d959cfeffff            lea edx, dword ptr [ebp+fffffe9c]
'00711544    52                      push edx
'00711545    50                      push eax
'00711546    ff5134                  call dword ptr [ecx+34]
'00711549    dbe2                    fnclex
'0071154b    85c0                    test ax, ax
'0071154d    7d0e                    jge 71155d
'0071154f    8b4ddc                  mov ecx, dword ptr [ebp-24]
'00711552    6a34                    push 34
'00711554    6830314300              push 00433130
'00711559    51                      push ecx
'0071155a    50                      push eax
'0071155b    ffd7                    call edi
'0071155d    6683bd9cfeffff00        cmp word ptr [ebp+fffffe9c], 00
'00711565    0f8583010000            jne 7116ee

Do While (IsNull(var_73) = 0)
'0071156b    8b45dc                  mov eax, dword ptr [ebp-24]
'0071156e    8d4dc8                  lea ecx, dword ptr [ebp-38]
'00711571    51                      push ecx
'00711572    c78548ffffff04000280    mov dword ptr [ebp+ffffff48], 80020004
'0071157c    c78540ffffff0a000000    mov dword ptr [ebp+ffffff40], 0000000a
'00711586    8b10                    mov edx, dword ptr [eax]
'00711588    50                      push eax
'00711589    ff92b4000000            call dword ptr [edx+000000b4]
'0071158f    dbe2                    fnclex
'00711591    85c0                    test ax, ax
'00711593    7d11                    jge 7115a6
'00711595    8b55dc                  mov edx, dword ptr [ebp-24]
'00711598    68b4000000              push 000000b4
'0071159d    6830314300              push 00433130
'007115a2    52                      push edx
'007115a3    50                      push eax
'007115a4    ffd7                    call edi
'007115a6    8b45c8                  mov eax, dword ptr [ebp-38]
'007115a9    8d5dc4                  lea ebx, dword ptr [ebp-3c]
'007115ac    53                      push ebx
'007115ad    83ec10                  sub esp, 10
'007115b0    ba08000000              mov edx, 00000008
'007115b5    8bdc                    mov ebx, esp
'007115b7    8913                    mov dword ptr [ebx], edx
'007115b9    899550ffffff            mov dword ptr [ebp+ffffff50], edx
'007115bf    8b9554ffffff            mov edx, dword ptr [ebp+ffffff54]
'007115c5    b964b14300              mov ecx, 0043b164
'007115ca    895304                  mov dword ptr [ebx+04], edx
'007115cd    898d58ffffff            mov dword ptr [ebp+ffffff58], ecx
'007115d3    8b38                    mov edi, dword ptr [eax]
'007115d5    894b08                  mov dword ptr [ebx+08], ecx
'007115d8    8b8d5cffffff            mov ecx, dword ptr [ebp+ffffff5c]
'007115de    50                      push eax
'007115df    898594feffff            mov dword ptr [ebp+fffffe94], eax
'007115e5    894b0c                  mov dword ptr [ebx+0c], ecx
'007115e8    ff5730                  call dword ptr [edi+30]
'007115eb    dbe2                    fnclex
'007115ed    85c0                    test ax, ax
'007115ef    7d19                    jge 71160a
'007115f1    8b9594feffff            mov edx, dword ptr [ebp+fffffe94]

' *** Reference to "__vbaHresultCheckObj"
'007115f7    8b3d80104000            mov edi, dword ptr [00401080]
'007115fd    6a30                    push 30
'007115ff    68d8304300              push 004330d8
'00711604    52                      push edx
'00711605    50                      push eax
'00711606    ffd7                    call edi
'00711608    eb06                    jmp 711610

' *** Reference to "__vbaHresultCheckObj"
'0071160a    8b3d80104000            mov edi, dword ptr [00401080]
'00711610    8b45c4                  mov eax, dword ptr [ebp-3c]
'00711613    8b08                    mov ecx, dword ptr [eax]
'00711615    8d55b0                  lea edx, dword ptr [ebp-50]
'00711618    52                      push edx
'00711619    50                      push eax
'0071161a    8bd8                    mov ebx, eax
'0071161c    ff5144                  call dword ptr [ecx+44]
'0071161f    dbe2                    fnclex
'00711621    85c0                    test ax, ax
'00711623    7d0b                    jge 711630
'00711625    6a44                    push 44
'00711627    6880304300              push 00433080
'0071162c    53                      push ebx
'0071162d    50                      push eax
'0071162e    ffd7                    call edi
'00711630    8b8574feffff            mov eax, dword ptr [ebp+fffffe74]
'00711636    8b9540ffffff            mov edx, dword ptr [ebp+ffffff40]
'0071163c    8b18                    mov ebx, dword ptr [eax]
'0071163e    8b8544ffffff            mov eax, dword ptr [ebp+ffffff44]
'00711644    83ec10                  sub esp, 10
'00711647    8bcc                    mov ecx, esp
'00711649    8911                    mov dword ptr [ecx], edx
'0071164b    8b9548ffffff            mov edx, dword ptr [ebp+ffffff48]
'00711651    894104                  mov dword ptr [ecx+04], eax
'00711654    8b854cffffff            mov eax, dword ptr [ebp+ffffff4c]
'0071165a    895108                  mov dword ptr [ecx+08], edx
'0071165d    89410c                  mov dword ptr [ecx+0c], eax
'00711660    8d4db0                  lea ecx, dword ptr [ebp-50]
'00711663    51                      push ecx

' *** Reference to "__vbaStrVarMove"
'00711664    ff153c104000            call dword ptr [0040103c]
'0071166a    8bd0                    mov edx, eax
'0071166c    8d4dd4                  lea ecx, dword ptr [ebp-2c]

' *** Reference to "__vbaStrMove"
'0071166f    ff15d0124000            call dword ptr [004012d0]
'00711675    8b9574feffff            mov edx, dword ptr [ebp+fffffe74]
'0071167b    50                      push eax
'0071167c    52                      push edx
'0071167d    ff93ec010000            call dword ptr [ebx+000001ec]
'00711683    dbe2                    fnclex
'00711685    85c0                    test ax, ax
'00711687    7d14                    jge 71169d
'00711689    8b8d74feffff            mov ecx, dword ptr [ebp+fffffe74]
'0071168f    68ec010000              push 000001ec
'00711694    681cb94300              push 0043b91c
'00711699    51                      push ecx
'0071169a    50                      push eax
'0071169b    ffd7                    call edi
'0071169d    8d4dd4                  lea ecx, dword ptr [ebp-2c]

' *** Reference to "__vbaFreeStr"
'007116a0    ff150c134000            call dword ptr [0040130c]
    '#Cleanup(var_44)
'007116a6    8d55c4                  lea edx, dword ptr [ebp-3c]
'007116a9    52                      push edx
'007116aa    8d45c8                  lea eax, dword ptr [ebp-38]
'007116ad    50                      push eax
'007116ae    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'007116b0    ff154c104000            call dword ptr [0040104c]
    '#Cleanup( var_46, var_9)
'007116b6    83c40c                  add esp, 0c
'007116b9    8d4db0                  lea ecx, dword ptr [ebp-50]

' *** Reference to "__vbaFreeVar"
'007116bc    ff1528104000            call dword ptr [00401028]
    '#Cleanup(var_6)
'007116c2    8b45dc                  mov eax, dword ptr [ebp-24]
'007116c5    8b08                    mov ecx, dword ptr [eax]
'007116c7    50                      push eax
'007116c8    ff91ec000000            call dword ptr [ecx+000000ec]
'007116ce    dbe2                    fnclex
'007116d0    85c0                    test ax, ax
'007116d2    0f8d61feffff            jge 711539
'007116d8    8b55dc                  mov edx, dword ptr [ebp-24]
'007116db    68ec000000              push 000000ec
'007116e0    6830314300              push 00433130
'007116e5    52                      push edx
'007116e6    50                      push eax
'007116e7    ffd7                    call edi
'007116e9    e94bfeffff              jmp 711539
    
Loop
'007116ee    8b06                    mov eax, dword ptr [esi]
'007116f0    56                      push esi
'007116f1    ff9008030000            call dword ptr [eax+00000308]
'007116f7    50                      push eax
'007116f8    8d4dc4                  lea ecx, dword ptr [ebp-3c]
'007116fb    51                      push ecx

' *** Reference to "__vbaObjSet"
'007116fc    ff15b4104000            call dword ptr [004010b4]
Set var_9 = Nothing
'00711702    8b16                    mov edx, dword ptr [esi]
'00711704    56                      push esi
'00711705    8bd8                    mov ebx, eax
'00711707    ff9208030000            call dword ptr [edx+00000308]
'0071170d    50                      push eax
'0071170e    8d45c8                  lea eax, dword ptr [ebp-38]
'00711711    50                      push eax

' *** Reference to "__vbaObjSet"
'00711712    ff15b4104000            call dword ptr [004010b4]
Set var_46 = Nothing
'00711718    8d959cfeffff            lea edx, dword ptr [ebp+fffffe9c]
'0071171e    8bf0                    mov esi, eax
'00711720    8b0e                    mov ecx, dword ptr [esi]
'00711722    52                      push edx
'00711723    56                      push esi
'00711724    ff91e8000000            call dword ptr [ecx+000000e8]
'0071172a    dbe2                    fnclex
'0071172c    85c0                    test ax, ax
'0071172e    7d0e                    jge 71173e
'00711730    68e8000000              push 000000e8
'00711735    681cb94300              push 0043b91c
'0071173a    56                      push esi
'0071173b    50                      push eax
'0071173c    ffd7                    call edi
'0071173e    668b8d9cfeffff          mov cx, word ptr [ebp+fffffe9c]
'00711745    8b03                    mov eax, dword ptr [ebx]
'00711747    6683e901                sub cx, 01
var_num3 = IsNull(var_73) - 1
'0071174b    0f8019010000            jo 71186a
'00711751    51                      push ecx
'00711752    53                      push ebx
'00711753    ff90f4000000            call dword ptr [eax+000000f4]
'00711759    dbe2                    fnclex
'0071175b    85c0                    test ax, ax
'0071175d    7d0e                    jge 71176d
'0071175f    68f4000000              push 000000f4
'00711764    681cb94300              push 0043b91c
'00711769    53                      push ebx
'0071176a    50                      push eax
'0071176b    ffd7                    call edi
'0071176d    8d55c4                  lea edx, dword ptr [ebp-3c]
'00711770    52                      push edx
'00711771    8d45c8                  lea eax, dword ptr [ebp-38]
'00711774    50                      push eax
'00711775    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'00711777    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_46, var_9)
'0071177d    83c40c                  add esp, 0c
'00711780    6a00                    push 00
'00711782    8d8d74feffff            lea ecx, dword ptr [ebp+fffffe74]
'00711788    51                      push ecx

' *** Reference to "__vbaObjSetAddref"
'00711789    ff15c8104000            call dword ptr [004010c8]
Set var_304 = Nothing
'0071178f    8b45dc                  mov eax, dword ptr [ebp-24]
'00711792    8b10                    mov edx, dword ptr [eax]
'00711794    50                      push eax
'00711795    ff92c4000000            call dword ptr [edx+000000c4]
'0071179b    dbe2                    fnclex
'0071179d    85c0                    test ax, ax
'0071179f    7d11                    jge 7117b2
'007117a1    8b4ddc                  mov ecx, dword ptr [ebp-24]
'007117a4    68c4000000              push 000000c4
'007117a9    6830314300              push 00433130
'007117ae    51                      push ecx
'007117af    50                      push eax
'007117b0    ffd7                    call edi
'007117b2    c745fc00000000          mov dword ptr [ebp-04], 00000000
'007117b9    9b                      fwait
'007117ba    684b187100              push 0071184b
'007117bf    eb52                    jmp 711813
'007117c1    8d55cc                  lea edx, dword ptr [ebp-34]
'007117c4    52                      push edx
'007117c5    8d45d0                  lea eax, dword ptr [ebp-30]
'007117c8    50                      push eax
'007117c9    8d4dd4                  lea ecx, dword ptr [ebp-2c]
'007117cc    51                      push ecx
'007117cd    6a03                    push 03

' *** Reference to "__vbaFreeStrList"
'007117cf    ff155c124000            call dword ptr [0040125c]
'#Cleanup( , 0, var_43)
'007117d5    8d55c0                  lea edx, dword ptr [ebp-40]
'007117d8    52                      push edx
'007117d9    8d45c4                  lea eax, dword ptr [ebp-3c]
'007117dc    50                      push eax
'007117dd    8d4dc8                  lea ecx, dword ptr [ebp-38]
'007117e0    51                      push ecx
'007117e1    6a03                    push 03

' *** Reference to "__vbaFreeObjList"
'007117e3    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_46, var_9, var_5)
'007117e9    8d9560ffffff            lea edx, dword ptr [ebp+ffffff60]
'007117ef    52                      push edx
'007117f0    8d8570ffffff            lea eax, dword ptr [ebp+ffffff70]
'007117f6    50                      push eax
'007117f7    8d4d80                  lea ecx, dword ptr [ebp-80]
'007117fa    51                      push ecx
'007117fb    8d5590                  lea edx, dword ptr [ebp-70]
'007117fe    52                      push edx
'007117ff    8d45a0                  lea eax, dword ptr [ebp-60]
'00711802    50                      push eax
'00711803    8d4db0                  lea ecx, dword ptr [ebp-50]
'00711806    51                      push ecx
'00711807    6a06                    push 06

' *** Reference to "__vbaFreeVarList"
'00711809    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_6, var_7, var_8, var_18, var_19, var_20)
'0071180f    83c43c                  add esp, 3c
'00711812    c3                      ret
'00711813    8d9574feffff            lea edx, dword ptr [ebp+fffffe74]
'00711819    52                      push edx
'0071181a    8d8578feffff            lea eax, dword ptr [ebp+fffffe78]
'00711820    50                      push eax
'00711821    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'00711823    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_73, var_304)
'00711829    83c40c                  add esp, 0c
'0071182c    8d4de8                  lea ecx, dword ptr [ebp-18]

' *** Reference to "__vbaFreeStr"
'0071182f    ff150c134000            call dword ptr [0040130c]
'#Cleanup(var_41)

' *** Reference to "__vbaFreeObj"
'00711835    8b3508134000            mov esi, dword ptr [00401308]
'0071183b    8d4de4                  lea ecx, dword ptr [ebp-1c]
'0071183e    ffd6                    call esi
'#Cleanup(var_40)
'00711840    8d4de0                  lea ecx, dword ptr [ebp-20]
'00711843    ffd6                    call esi
'#Cleanup(var_3)
'00711845    8d4ddc                  lea ecx, dword ptr [ebp-24]
'00711848    ffd6                    call esi
'#Cleanup(var_39)
'0071184a    c3                      ret
'0071184b    8b4508                  mov eax, dword ptr [ebp+08]
'0071184e    8b08                    mov ecx, dword ptr [eax]
'00711850    50                      push eax
'00711851    ff5108                  call dword ptr [ecx+08]
'00711854    8b45fc                  mov eax, dword ptr [ebp-04]
'00711857    8b4dec                  mov ecx, dword ptr [ebp-14]
'0071185a    5f                      pop edi
'0071185b    5e                      pop esi
    'Reference to '__except_list'
'0071185c    64890d00000000          mov dword ptr fs:[00000000], ecx
'00711863    5b                      pop ebx
'00711864    8be5                    mov esp, ebp
'00711866    5d                      pop ebp
'00711867    c20400                  ret 0004


End Sub


'Event for Form
Private Sub Form_Load()
'00712000    55                      push ebp
'00712001    8bec                    mov ebp, esp
'00712003    83ec0c                  sub esp, 0c

' *** Reference to "__vbaExceptHandler"
'00712006    6866724000              push 00407266
'0071200b    64a100000000            mov ax, word ptr fs:[00000000]
'00712011    50                      push eax
    'Reference to '__except_list'
'00712012    64892500000000          mov dword ptr fs:[00000000], esp
'00712019    81eca8000000            sub esp, 000000a8
'0071201f    53                      push ebx
'00712020    56                      push esi
'00712021    57                      push edi
'00712022    8965f4                  mov dword ptr [ebp-0c], esp
'00712025    c745f8806d4000          mov dword ptr [ebp-08], 00406d80
'0071202c    8b7d08                  mov edi, dword ptr [ebp+08]
'0071202f    8bc7                    mov eax, edi
'00712031    83e001                  and eax, 01
'00712034    8945fc                  mov dword ptr [ebp-04], eax
'00712037    83e7fe                  and edi, fffffffe
'0071203a    8b0f                    mov ecx, dword ptr [edi]
'0071203c    57                      push edi
'0071203d    897d08                  mov dword ptr [ebp+08], edi
'00712040    ff5104                  call dword ptr [ecx+04]
'00712043    8b17                    mov edx, dword ptr [edi]
'00712045    33f6                    xor esi, esi
'00712047    57                      push edi
'00712048    8975e8                  mov dword ptr [ebp-18], esi
'0071204b    8975e4                  mov dword ptr [ebp-1c], esi
'0071204e    8975e0                  mov dword ptr [ebp-20], esi
'00712051    8975dc                  mov dword ptr [ebp-24], esi
'00712054    8975d8                  mov dword ptr [ebp-28], esi
'00712057    8975c8                  mov dword ptr [ebp-38], esi
'0071205a    8975b8                  mov dword ptr [ebp-48], esi
'0071205d    897598                  mov dword ptr [ebp-68], esi
'00712060    897584                  mov dword ptr [ebp-7c], esi
'00712063    89b568ffffff            mov dword ptr [ebp+ffffff68], esi
'00712069    89b564ffffff            mov dword ptr [ebp+ffffff64], esi
'0071206f    89b560ffffff            mov dword ptr [ebp+ffffff60], esi
'00712075    89b55cffffff            mov dword ptr [ebp+ffffff5c], esi
'0071207b    ff9208030000            call dword ptr [edx+00000308]
'00712081    50                      push eax
'00712082    8d8568ffffff            lea eax, dword ptr [ebp+ffffff68]
'00712088    50                      push eax

' *** Reference to "__vbaObjSet"
'00712089    ff15b4104000            call dword ptr [004010b4]
Set var_132 = Me
'0071208f    8b8568ffffff            mov eax, dword ptr [ebp+ffffff68]
'00712095    8b08                    mov ecx, dword ptr [eax]
'00712097    50                      push eax
'00712098    ff91e8010000            call dword ptr [ecx+000001e8]
'0071209e    dbe2                    fnclex
'007120a0    3bc6                    cmp eax, esi
'007120a2    7d1c                    jge 7120c0

If (var_132 < 0) Then
'007120a4    8b9568ffffff            mov edx, dword ptr [ebp+ffffff68]

' *** Reference to "__vbaHresultCheckObj"
'007120aa    8b3580104000            mov esi, dword ptr [00401080]
'007120b0    68e8010000              push 000001e8
'007120b5    681cb94300              push 0043b91c
'007120ba    52                      push edx
'007120bb    50                      push eax
'007120bc    ffd6                    call esi
'007120be    eb06                    jmp 7120c6
    
End If

' *** Reference to "__vbaHresultCheckObj"
'007120c0    8b3580104000            mov esi, dword ptr [00401080]
'007120c6    8d5ddc                  lea ebx, dword ptr [ebp-24]
'007120c9    53                      push ebx
'007120ca    83ec10                  sub esp, 10
'007120cd    8bdc                    mov ebx, esp
'007120cf    b90a000000              mov ecx, 0000000a
'007120d4    890b                    mov dword ptr [ebx], ecx
'007120d6    894d98                  mov dword ptr [ebp-68], ecx
'007120d9    8b4d8c                  mov ecx, dword ptr [ebp-74]
'007120dc    894b04                  mov dword ptr [ebx+04], ecx
'007120df    8b3d48b07200            mov edi, dword ptr [0072b048]
'007120e5    b804000280              mov eax, 80020004
'007120ea    894308                  mov dword ptr [ebx+08], eax
'007120ed    8bd0                    mov edx, eax
'007120ef    8b4594                  mov eax, dword ptr [ebp-6c]
'007120f2    89430c                  mov dword ptr [ebx+0c], eax
'007120f5    8b4598                  mov eax, dword ptr [ebp-68]
'007120f8    83ec10                  sub esp, 10
'007120fb    8bcc                    mov ecx, esp
'007120fd    8901                    mov dword ptr [ecx], eax
'007120ff    8b459c                  mov eax, dword ptr [ebp-64]
'00712102    894104                  mov dword ptr [ecx+04], eax
'00712105    895108                  mov dword ptr [ecx+08], edx
'00712108    8955a0                  mov dword ptr [ebp-60], edx
'0071210b    8b55a4                  mov edx, dword ptr [ebp-5c]
'0071210e    8b3f                    mov edi, dword ptr [edi]
'00712110    89510c                  mov dword ptr [ecx+0c], edx
'00712113    8b55ac                  mov edx, dword ptr [ebp-54]
'00712116    83ec10                  sub esp, 10
'00712119    8bcc                    mov ecx, esp
'0071211b    b803000000              mov eax, 00000003
'00712120    8901                    mov dword ptr [ecx], eax
'00712122    895104                  mov dword ptr [ecx+04], edx
'00712125    b802000000              mov eax, 00000002
'0071212a    894108                  mov dword ptr [ecx+08], eax
'0071212d    8b45b4                  mov eax, dword ptr [ebp-4c]
'00712130    89410c                  mov dword ptr [ecx+0c], eax
'00712133    8b0d48b07200            mov ecx, dword ptr [0072b048]
'00712139    689c694500              push 0045699c
'0071213e    51                      push ecx
'0071213f    ff97bc000000            call dword ptr [edi+000000bc]
'00712145    dbe2                    fnclex
'00712147    85c0                    test ax, ax
'00712149    7d14                    jge 71215f
'0071214b    8b1548b07200            mov edx, dword ptr [0072b048]
'00712151    68bc000000              push 000000bc
'00712156    68301f4300              push 00431f30
'0071215b    52                      push edx
'0071215c    50                      push eax
'0071215d    ffd6                    call esi
'0071215f    8b45dc                  mov eax, dword ptr [ebp-24]
'00712162    50                      push eax
'00712163    8d45e4                  lea eax, dword ptr [ebp-1c]
'00712166    50                      push eax
'00712167    c745dc00000000          mov dword ptr [ebp-24], 00000000

' *** Reference to "__vbaObjSet"
'0071216e    ff15b4104000            call dword ptr [004010b4]
Set var_40 = Nothing
'00712174    8b45e4                  mov eax, dword ptr [ebp-1c]
'00712177    8b08                    mov ecx, dword ptr [eax]
'00712179    8d5584                  lea edx, dword ptr [ebp-7c]
'0071217c    52                      push edx
'0071217d    50                      push eax
'0071217e    ff5134                  call dword ptr [ecx+34]
'00712181    dbe2                    fnclex
'00712183    85c0                    test ax, ax
'00712185    7d0e                    jge 712195
'00712187    8b4de4                  mov ecx, dword ptr [ebp-1c]
'0071218a    6a34                    push 34
'0071218c    6830314300              push 00433130
'00712191    51                      push ecx
'00712192    50                      push eax
'00712193    ffd6                    call esi
'00712195    66837d8400              cmp word ptr [ebp-7c], 00
'0071219a    0f8551010000            jne 7122f1

Do While (0 = 0)
'007121a0    8b45e4                  mov eax, dword ptr [ebp-1c]
'007121a3    8d4ddc                  lea ecx, dword ptr [ebp-24]
'007121a6    51                      push ecx
'007121a7    c745a004000280          mov dword ptr [ebp-60], 80020004
'007121ae    c745980a000000          mov dword ptr [ebp-68], 0000000a
'007121b5    8b10                    mov edx, dword ptr [eax]
'007121b7    50                      push eax
'007121b8    ff92b4000000            call dword ptr [edx+000000b4]
'007121be    dbe2                    fnclex
'007121c0    85c0                    test ax, ax
'007121c2    7d11                    jge 7121d5
'007121c4    8b55e4                  mov edx, dword ptr [ebp-1c]
'007121c7    68b4000000              push 000000b4
'007121cc    6830314300              push 00433130
'007121d1    52                      push edx
'007121d2    50                      push eax
'007121d3    ffd6                    call esi
'007121d5    8b45dc                  mov eax, dword ptr [ebp-24]
'007121d8    8b38                    mov edi, dword ptr [eax]
'007121da    8d5dd8                  lea ebx, dword ptr [ebp-28]
'007121dd    53                      push ebx
'007121de    83ec10                  sub esp, 10
'007121e1    8bdc                    mov ebx, esp
'007121e3    ba08000000              mov edx, 00000008
'007121e8    8913                    mov dword ptr [ebx], edx
'007121ea    8b55ac                  mov edx, dword ptr [ebp-54]
'007121ed    895304                  mov dword ptr [ebx+04], edx
'007121f0    b964b14300              mov ecx, 0043b164
'007121f5    894b08                  mov dword ptr [ebx+08], ecx
'007121f8    8b4db4                  mov ecx, dword ptr [ebp-4c]
'007121fb    50                      push eax
'007121fc    89857cffffff            mov dword ptr [ebp+ffffff7c], eax
'00712202    894b0c                  mov dword ptr [ebx+0c], ecx
'00712205    ff5730                  call dword ptr [edi+30]
'00712208    dbe2                    fnclex
'0071220a    85c0                    test ax, ax
'0071220c    7d11                    jge 71221f
'0071220e    8b957cffffff            mov edx, dword ptr [ebp+ffffff7c]
'00712214    6a30                    push 30
'00712216    68d8304300              push 004330d8
'0071221b    52                      push edx
'0071221c    50                      push eax
'0071221d    ffd6                    call esi
'0071221f    8b45d8                  mov eax, dword ptr [ebp-28]
'00712222    8b08                    mov ecx, dword ptr [eax]
'00712224    8d55c8                  lea edx, dword ptr [ebp-38]
'00712227    52                      push edx
'00712228    50                      push eax
'00712229    8bf8                    mov edi, eax
'0071222b    ff5144                  call dword ptr [ecx+44]
'0071222e    dbe2                    fnclex
'00712230    85c0                    test ax, ax
'00712232    7d0b                    jge 71223f
'00712234    6a44                    push 44
'00712236    6880304300              push 00433080
'0071223b    57                      push edi
'0071223c    50                      push eax
'0071223d    ffd6                    call esi
'0071223f    8b8568ffffff            mov eax, dword ptr [ebp+ffffff68]
'00712245    8b5598                  mov edx, dword ptr [ebp-68]
'00712248    8b38                    mov edi, dword ptr [eax]
'0071224a    8b459c                  mov eax, dword ptr [ebp-64]
'0071224d    83ec10                  sub esp, 10
'00712250    8bcc                    mov ecx, esp
'00712252    8911                    mov dword ptr [ecx], edx
'00712254    8b55a0                  mov edx, dword ptr [ebp-60]
'00712257    894104                  mov dword ptr [ecx+04], eax
'0071225a    8b45a4                  mov eax, dword ptr [ebp-5c]
'0071225d    895108                  mov dword ptr [ecx+08], edx
'00712260    89410c                  mov dword ptr [ecx+0c], eax
'00712263    8d4dc8                  lea ecx, dword ptr [ebp-38]
'00712266    51                      push ecx

' *** Reference to "__vbaStrVarMove"
'00712267    ff153c104000            call dword ptr [0040103c]
'0071226d    8bd0                    mov edx, eax
'0071226f    8d4de0                  lea ecx, dword ptr [ebp-20]

' *** Reference to "__vbaStrMove"
'00712272    ff15d0124000            call dword ptr [004012d0]
'00712278    8b9568ffffff            mov edx, dword ptr [ebp+ffffff68]
'0071227e    50                      push eax
'0071227f    52                      push edx
'00712280    ff97ec010000            call dword ptr [edi+000001ec]
'00712286    dbe2                    fnclex
'00712288    85c0                    test ax, ax
'0071228a    7d14                    jge 7122a0
'0071228c    8b8d68ffffff            mov ecx, dword ptr [ebp+ffffff68]
'00712292    68ec010000              push 000001ec
'00712297    681cb94300              push 0043b91c
'0071229c    51                      push ecx
'0071229d    50                      push eax
'0071229e    ffd6                    call esi
'007122a0    8d4de0                  lea ecx, dword ptr [ebp-20]

' *** Reference to "__vbaFreeStr"
'007122a3    ff150c134000            call dword ptr [0040130c]
    '#Cleanup(var_3)
'007122a9    8d55d8                  lea edx, dword ptr [ebp-28]
'007122ac    52                      push edx
'007122ad    8d45dc                  lea eax, dword ptr [ebp-24]
'007122b0    50                      push eax
'007122b1    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'007122b3    ff154c104000            call dword ptr [0040104c]
    '#Cleanup( var_39, var_45)
'007122b9    83c40c                  add esp, 0c
'007122bc    8d4dc8                  lea ecx, dword ptr [ebp-38]

' *** Reference to "__vbaFreeVar"
'007122bf    ff1528104000            call dword ptr [00401028]
    '#Cleanup(var_46)
'007122c5    8b45e4                  mov eax, dword ptr [ebp-1c]
'007122c8    8b08                    mov ecx, dword ptr [eax]
'007122ca    50                      push eax
'007122cb    ff91ec000000            call dword ptr [ecx+000000ec]
'007122d1    dbe2                    fnclex
'007122d3    85c0                    test ax, ax
'007122d5    0f8d99feffff            jge 712174
'007122db    8b55e4                  mov edx, dword ptr [ebp-1c]
'007122de    68ec000000              push 000000ec
'007122e3    6830314300              push 00433130
'007122e8    52                      push edx
'007122e9    50                      push eax
'007122ea    ffd6                    call esi
'007122ec    e983feffff              jmp 712174
    
Loop
'007122f1    6a00                    push 00
'007122f3    8d8568ffffff            lea eax, dword ptr [ebp+ffffff68]
'007122f9    50                      push eax

' *** Reference to "__vbaObjSetAddref"
'007122fa    ff15c8104000            call dword ptr [004010c8]
Set var_132 = Nothing
'00712300    8b45e4                  mov eax, dword ptr [ebp-1c]
'00712303    8b08                    mov ecx, dword ptr [eax]
'00712305    50                      push eax
'00712306    ff91c4000000            call dword ptr [ecx+000000c4]
'0071230c    dbe2                    fnclex
'0071230e    85c0                    test ax, ax
'00712310    7d11                    jge 712323
'00712312    8b55e4                  mov edx, dword ptr [ebp-1c]
'00712315    68c4000000              push 000000c4
'0071231a    6830314300              push 00433130
'0071231f    52                      push edx
'00712320    50                      push eax
'00712321    ffd6                    call esi
'00712323    8b4508                  mov eax, dword ptr [ebp+08]
'00712326    8b08                    mov ecx, dword ptr [eax]
'00712328    50                      push eax
'00712329    ff9100030000            call dword ptr [ecx+00000300]
'0071232f    50                      push eax
'00712330    8d9564ffffff            lea edx, dword ptr [ebp+ffffff64]
'00712336    52                      push edx

' *** Reference to "__vbaObjSet"
'00712337    ff15b4104000            call dword ptr [004010b4]
Set var_123 = Me
'0071233d    8b8564ffffff            mov eax, dword ptr [ebp+ffffff64]
'00712343    8b08                    mov ecx, dword ptr [eax]
'00712345    50                      push eax
'00712346    ff91e8010000            call dword ptr [ecx+000001e8]
'0071234c    dbe2                    fnclex
'0071234e    85c0                    test ax, ax
'00712350    7d14                    jge 712366
'00712352    8b9564ffffff            mov edx, dword ptr [ebp+ffffff64]
'00712358    68e8010000              push 000001e8
'0071235d    681cb94300              push 0043b91c
'00712362    52                      push edx
'00712363    50                      push eax
'00712364    ffd6                    call esi
'00712366    8b9564ffffff            mov edx, dword ptr [ebp+ffffff64]
'0071236c    8b3a                    mov edi, dword ptr [edx]
'0071236e    83ec10                  sub esp, 10
'00712371    8bdc                    mov ebx, esp
'00712373    b90a000000              mov ecx, 0000000a
'00712378    890b                    mov dword ptr [ebx], ecx
'0071237a    8b4dac                  mov ecx, dword ptr [ebp-54]
'0071237d    894b04                  mov dword ptr [ebx+04], ecx
'00712380    b804000280              mov eax, 80020004
'00712385    894308                  mov dword ptr [ebx+08], eax
'00712388    8b45b4                  mov eax, dword ptr [ebp-4c]
'0071238b    68cc134300              push 004313cc
'00712390    52                      push edx
'00712391    89430c                  mov dword ptr [ebx+0c], eax
'00712394    ff97ec010000            call dword ptr [edi+000001ec]
'0071239a    dbe2                    fnclex
'0071239c    85c0                    test ax, ax
'0071239e    7d14                    jge 7123b4
'007123a0    8b8d64ffffff            mov ecx, dword ptr [ebp+ffffff64]
'007123a6    68ec010000              push 000001ec
'007123ab    681cb94300              push 0043b91c
'007123b0    51                      push ecx
'007123b1    50                      push eax
'007123b2    ffd6                    call esi
'007123b4    8d5ddc                  lea ebx, dword ptr [ebp-24]
'007123b7    53                      push ebx
'007123b8    83ec10                  sub esp, 10
'007123bb    8bdc                    mov ebx, esp
'007123bd    b90a000000              mov ecx, 0000000a
'007123c2    890b                    mov dword ptr [ebx], ecx
'007123c4    894d98                  mov dword ptr [ebp-68], ecx
'007123c7    8b4d8c                  mov ecx, dword ptr [ebp-74]
'007123ca    894b04                  mov dword ptr [ebx+04], ecx
'007123cd    8b3d4cb07200            mov edi, dword ptr [0072b04c]
'007123d3    b804000280              mov eax, 80020004
'007123d8    894308                  mov dword ptr [ebx+08], eax
'007123db    8bd0                    mov edx, eax
'007123dd    8b4594                  mov eax, dword ptr [ebp-6c]
'007123e0    89430c                  mov dword ptr [ebx+0c], eax
'007123e3    8b4598                  mov eax, dword ptr [ebp-68]
'007123e6    83ec10                  sub esp, 10
'007123e9    8bcc                    mov ecx, esp
'007123eb    8901                    mov dword ptr [ecx], eax
'007123ed    8b459c                  mov eax, dword ptr [ebp-64]
'007123f0    894104                  mov dword ptr [ecx+04], eax
'007123f3    895108                  mov dword ptr [ecx+08], edx
'007123f6    8955a0                  mov dword ptr [ebp-60], edx
'007123f9    8b55a4                  mov edx, dword ptr [ebp-5c]
'007123fc    8b3f                    mov edi, dword ptr [edi]
'007123fe    89510c                  mov dword ptr [ecx+0c], edx
'00712401    8b55ac                  mov edx, dword ptr [ebp-54]
'00712404    83ec10                  sub esp, 10
'00712407    8bc4                    mov eax, esp
'00712409    c745a803000000          mov dword ptr [ebp-58], 00000003
'00712410    8b4da8                  mov ecx, dword ptr [ebp-58]
'00712413    8908                    mov dword ptr [eax], ecx
'00712415    895004                  mov dword ptr [eax+04], edx
'00712418    8b55b4                  mov edx, dword ptr [ebp-4c]
'0071241b    c745b002000000          mov dword ptr [ebp-50], 00000002
'00712422    8b4db0                  mov ecx, dword ptr [ebp-50]
'00712425    894808                  mov dword ptr [eax+08], ecx
'00712428    89500c                  mov dword ptr [eax+0c], edx
'0071242b    a14cb07200              mov ax, word ptr [0072b04c]
'00712430    682c6a4500              push 00456a2c
'00712435    50                      push eax
'00712436    ff97bc000000            call dword ptr [edi+000000bc]
'0071243c    dbe2                    fnclex
'0071243e    85c0                    test ax, ax
'00712440    7d14                    jge 712456
'00712442    8b0d4cb07200            mov ecx, dword ptr [0072b04c]
'00712448    68bc000000              push 000000bc
'0071244d    68301f4300              push 00431f30
'00712452    51                      push ecx
'00712453    50                      push eax
'00712454    ffd6                    call esi
'00712456    8b45dc                  mov eax, dword ptr [ebp-24]
'00712459    50                      push eax
'0071245a    8d55e8                  lea edx, dword ptr [ebp-18]
'0071245d    52                      push edx
'0071245e    c745dc00000000          mov dword ptr [ebp-24], 00000000

' *** Reference to "__vbaObjSet"
'00712465    ff15b4104000            call dword ptr [004010b4]
Set var_41 = Nothing
'0071246b    8b45e8                  mov eax, dword ptr [ebp-18]
'0071246e    8b08                    mov ecx, dword ptr [eax]
'00712470    8d5584                  lea edx, dword ptr [ebp-7c]
'00712473    52                      push edx
'00712474    50                      push eax
'00712475    ff5134                  call dword ptr [ecx+34]
'00712478    dbe2                    fnclex
'0071247a    85c0                    test ax, ax
'0071247c    7d0e                    jge 71248c
'0071247e    8b4de8                  mov ecx, dword ptr [ebp-18]
'00712481    6a34                    push 34
'00712483    6830314300              push 00433130
'00712488    51                      push ecx
'00712489    50                      push eax
'0071248a    ffd6                    call esi
'0071248c    66837d8400              cmp word ptr [ebp-7c], 00
'00712491    0f8551010000            jne 7125e8

Do While (0 = 0)
'00712497    8b45e8                  mov eax, dword ptr [ebp-18]
'0071249a    8d4ddc                  lea ecx, dword ptr [ebp-24]
'0071249d    51                      push ecx
'0071249e    c745a004000280          mov dword ptr [ebp-60], 80020004
'007124a5    c745980a000000          mov dword ptr [ebp-68], 0000000a
'007124ac    8b10                    mov edx, dword ptr [eax]
'007124ae    50                      push eax
'007124af    ff92b4000000            call dword ptr [edx+000000b4]
'007124b5    dbe2                    fnclex
'007124b7    85c0                    test ax, ax
'007124b9    7d11                    jge 7124cc
'007124bb    8b55e8                  mov edx, dword ptr [ebp-18]
'007124be    68b4000000              push 000000b4
'007124c3    6830314300              push 00433130
'007124c8    52                      push edx
'007124c9    50                      push eax
'007124ca    ffd6                    call esi
'007124cc    8b45dc                  mov eax, dword ptr [ebp-24]
'007124cf    8b38                    mov edi, dword ptr [eax]
'007124d1    8d5dd8                  lea ebx, dword ptr [ebp-28]
'007124d4    53                      push ebx
'007124d5    83ec10                  sub esp, 10
'007124d8    8bdc                    mov ebx, esp
'007124da    ba08000000              mov edx, 00000008
'007124df    8913                    mov dword ptr [ebx], edx
'007124e1    8b55ac                  mov edx, dword ptr [ebp-54]
'007124e4    895304                  mov dword ptr [ebx+04], edx
'007124e7    b964b14300              mov ecx, 0043b164
'007124ec    894b08                  mov dword ptr [ebx+08], ecx
'007124ef    8b4db4                  mov ecx, dword ptr [ebp-4c]
'007124f2    50                      push eax
'007124f3    89857cffffff            mov dword ptr [ebp+ffffff7c], eax
'007124f9    894b0c                  mov dword ptr [ebx+0c], ecx
'007124fc    ff5730                  call dword ptr [edi+30]
'007124ff    dbe2                    fnclex
'00712501    85c0                    test ax, ax
'00712503    7d11                    jge 712516
'00712505    8b957cffffff            mov edx, dword ptr [ebp+ffffff7c]
'0071250b    6a30                    push 30
'0071250d    68d8304300              push 004330d8
'00712512    52                      push edx
'00712513    50                      push eax
'00712514    ffd6                    call esi
'00712516    8b45d8                  mov eax, dword ptr [ebp-28]
'00712519    8b08                    mov ecx, dword ptr [eax]
'0071251b    8d55c8                  lea edx, dword ptr [ebp-38]
'0071251e    52                      push edx
'0071251f    50                      push eax
'00712520    8bf8                    mov edi, eax
'00712522    ff5144                  call dword ptr [ecx+44]
'00712525    dbe2                    fnclex
'00712527    85c0                    test ax, ax
'00712529    7d0b                    jge 712536
'0071252b    6a44                    push 44
'0071252d    6880304300              push 00433080
'00712532    57                      push edi
'00712533    50                      push eax
'00712534    ffd6                    call esi
'00712536    8b8564ffffff            mov eax, dword ptr [ebp+ffffff64]
'0071253c    8b5598                  mov edx, dword ptr [ebp-68]
'0071253f    8b38                    mov edi, dword ptr [eax]
'00712541    8b459c                  mov eax, dword ptr [ebp-64]
'00712544    83ec10                  sub esp, 10
'00712547    8bcc                    mov ecx, esp
'00712549    8911                    mov dword ptr [ecx], edx
'0071254b    8b55a0                  mov edx, dword ptr [ebp-60]
'0071254e    894104                  mov dword ptr [ecx+04], eax
'00712551    8b45a4                  mov eax, dword ptr [ebp-5c]
'00712554    895108                  mov dword ptr [ecx+08], edx
'00712557    89410c                  mov dword ptr [ecx+0c], eax
'0071255a    8d4dc8                  lea ecx, dword ptr [ebp-38]
'0071255d    51                      push ecx

' *** Reference to "__vbaStrVarMove"
'0071255e    ff153c104000            call dword ptr [0040103c]
'00712564    8bd0                    mov edx, eax
'00712566    8d4de0                  lea ecx, dword ptr [ebp-20]

' *** Reference to "__vbaStrMove"
'00712569    ff15d0124000            call dword ptr [004012d0]
'0071256f    8b9564ffffff            mov edx, dword ptr [ebp+ffffff64]
'00712575    50                      push eax
'00712576    52                      push edx
'00712577    ff97ec010000            call dword ptr [edi+000001ec]
'0071257d    dbe2                    fnclex
'0071257f    85c0                    test ax, ax
'00712581    7d14                    jge 712597
'00712583    8b8d64ffffff            mov ecx, dword ptr [ebp+ffffff64]
'00712589    68ec010000              push 000001ec
'0071258e    681cb94300              push 0043b91c
'00712593    51                      push ecx
'00712594    50                      push eax
'00712595    ffd6                    call esi
'00712597    8d4de0                  lea ecx, dword ptr [ebp-20]

' *** Reference to "__vbaFreeStr"
'0071259a    ff150c134000            call dword ptr [0040130c]
    '#Cleanup(var_3)
'007125a0    8d55d8                  lea edx, dword ptr [ebp-28]
'007125a3    52                      push edx
'007125a4    8d45dc                  lea eax, dword ptr [ebp-24]
'007125a7    50                      push eax
'007125a8    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'007125aa    ff154c104000            call dword ptr [0040104c]
    '#Cleanup( var_39, var_45)
'007125b0    83c40c                  add esp, 0c
'007125b3    8d4dc8                  lea ecx, dword ptr [ebp-38]

' *** Reference to "__vbaFreeVar"
'007125b6    ff1528104000            call dword ptr [00401028]
    '#Cleanup(var_46)
'007125bc    8b45e8                  mov eax, dword ptr [ebp-18]
'007125bf    8b08                    mov ecx, dword ptr [eax]
'007125c1    50                      push eax
'007125c2    ff91ec000000            call dword ptr [ecx+000000ec]
'007125c8    dbe2                    fnclex
'007125ca    85c0                    test ax, ax
'007125cc    0f8d99feffff            jge 71246b
'007125d2    8b55e8                  mov edx, dword ptr [ebp-18]
'007125d5    68ec000000              push 000000ec
'007125da    6830314300              push 00433130
'007125df    52                      push edx
'007125e0    50                      push eax
'007125e1    ffd6                    call esi
'007125e3    e983feffff              jmp 71246b
    
Loop
'007125e8    6a00                    push 00
'007125ea    8d8564ffffff            lea eax, dword ptr [ebp+ffffff64]
'007125f0    50                      push eax

' *** Reference to "__vbaObjSetAddref"
'007125f1    ff15c8104000            call dword ptr [004010c8]
Set var_123 = Nothing
'007125f7    8b45e8                  mov eax, dword ptr [ebp-18]
'007125fa    8b08                    mov ecx, dword ptr [eax]
'007125fc    50                      push eax
'007125fd    ff91c4000000            call dword ptr [ecx+000000c4]
'00712603    dbe2                    fnclex
'00712605    85c0                    test ax, ax
'00712607    7d11                    jge 71261a
'00712609    8b55e8                  mov edx, dword ptr [ebp-18]
'0071260c    68c4000000              push 000000c4
'00712611    6830314300              push 00433130
'00712616    52                      push edx
'00712617    50                      push eax
'00712618    ffd6                    call esi
'0071261a    8b4508                  mov eax, dword ptr [ebp+08]
'0071261d    8b08                    mov ecx, dword ptr [eax]
'0071261f    50                      push eax
'00712620    ff9104030000            call dword ptr [ecx+00000304]
'00712626    50                      push eax
'00712627    8d9560ffffff            lea edx, dword ptr [ebp+ffffff60]
'0071262d    52                      push edx

' *** Reference to "__vbaObjSet"
'0071262e    ff15b4104000            call dword ptr [004010b4]
Set var_20 = Me
'00712634    8b8560ffffff            mov eax, dword ptr [ebp+ffffff60]
'0071263a    8b08                    mov ecx, dword ptr [eax]
'0071263c    50                      push eax
'0071263d    ff91e8010000            call dword ptr [ecx+000001e8]
'00712643    dbe2                    fnclex
'00712645    85c0                    test ax, ax
'00712647    7d14                    jge 71265d
'00712649    8b9560ffffff            mov edx, dword ptr [ebp+ffffff60]
'0071264f    68e8010000              push 000001e8
'00712654    681cb94300              push 0043b91c
'00712659    52                      push edx
'0071265a    50                      push eax
'0071265b    ffd6                    call esi
'0071265d    8b9560ffffff            mov edx, dword ptr [ebp+ffffff60]
'00712663    8b3a                    mov edi, dword ptr [edx]
'00712665    83ec10                  sub esp, 10
'00712668    8bdc                    mov ebx, esp
'0071266a    b90a000000              mov ecx, 0000000a
'0071266f    890b                    mov dword ptr [ebx], ecx
'00712671    8b4dac                  mov ecx, dword ptr [ebp-54]
'00712674    894b04                  mov dword ptr [ebx+04], ecx
'00712677    b804000280              mov eax, 80020004
'0071267c    894308                  mov dword ptr [ebx+08], eax
'0071267f    8b45b4                  mov eax, dword ptr [ebp-4c]
'00712682    68cc134300              push 004313cc
'00712687    52                      push edx
'00712688    89430c                  mov dword ptr [ebx+0c], eax
'0071268b    ff97ec010000            call dword ptr [edi+000001ec]
'00712691    dbe2                    fnclex
'00712693    85c0                    test ax, ax
'00712695    7d14                    jge 7126ab
'00712697    8b8d60ffffff            mov ecx, dword ptr [ebp+ffffff60]
'0071269d    68ec010000              push 000001ec
'007126a2    681cb94300              push 0043b91c
'007126a7    51                      push ecx
'007126a8    50                      push eax
'007126a9    ffd6                    call esi
'007126ab    8d5ddc                  lea ebx, dword ptr [ebp-24]
'007126ae    53                      push ebx
'007126af    83ec10                  sub esp, 10
'007126b2    8bdc                    mov ebx, esp
'007126b4    b90a000000              mov ecx, 0000000a
'007126b9    890b                    mov dword ptr [ebx], ecx
'007126bb    894d98                  mov dword ptr [ebp-68], ecx
'007126be    8b4d8c                  mov ecx, dword ptr [ebp-74]
'007126c1    894b04                  mov dword ptr [ebx+04], ecx
'007126c4    8b3d4cb07200            mov edi, dword ptr [0072b04c]
'007126ca    b804000280              mov eax, 80020004
'007126cf    894308                  mov dword ptr [ebx+08], eax
'007126d2    8bd0                    mov edx, eax
'007126d4    8b4594                  mov eax, dword ptr [ebp-6c]
'007126d7    89430c                  mov dword ptr [ebx+0c], eax
'007126da    8b4598                  mov eax, dword ptr [ebp-68]
'007126dd    83ec10                  sub esp, 10
'007126e0    8bcc                    mov ecx, esp
'007126e2    8901                    mov dword ptr [ecx], eax
'007126e4    8b459c                  mov eax, dword ptr [ebp-64]
'007126e7    894104                  mov dword ptr [ecx+04], eax
'007126ea    895108                  mov dword ptr [ecx+08], edx
'007126ed    8955a0                  mov dword ptr [ebp-60], edx
'007126f0    8b55a4                  mov edx, dword ptr [ebp-5c]
'007126f3    8b3f                    mov edi, dword ptr [edi]
'007126f5    89510c                  mov dword ptr [ecx+0c], edx
'007126f8    8b55ac                  mov edx, dword ptr [ebp-54]
'007126fb    83ec10                  sub esp, 10
'007126fe    8bcc                    mov ecx, esp
'00712700    b803000000              mov eax, 00000003
'00712705    8901                    mov dword ptr [ecx], eax
'00712707    895104                  mov dword ptr [ecx+04], edx
'0071270a    b802000000              mov eax, 00000002
'0071270f    894108                  mov dword ptr [ecx+08], eax
'00712712    8b45b4                  mov eax, dword ptr [ebp-4c]
'00712715    89410c                  mov dword ptr [ecx+0c], eax
'00712718    8b0d4cb07200            mov ecx, dword ptr [0072b04c]
'0071271e    68746a4500              push 00456a74
'00712723    51                      push ecx
'00712724    ff97bc000000            call dword ptr [edi+000000bc]
'0071272a    dbe2                    fnclex
'0071272c    85c0                    test ax, ax
'0071272e    7d14                    jge 712744
'00712730    8b154cb07200            mov edx, dword ptr [0072b04c]
'00712736    68bc000000              push 000000bc
'0071273b    68301f4300              push 00431f30
'00712740    52                      push edx
'00712741    50                      push eax
'00712742    ffd6                    call esi
'00712744    8b45dc                  mov eax, dword ptr [ebp-24]
'00712747    50                      push eax
'00712748    8d45e8                  lea eax, dword ptr [ebp-18]
'0071274b    50                      push eax
'0071274c    c745dc00000000          mov dword ptr [ebp-24], 00000000

' *** Reference to "__vbaObjSet"
'00712753    ff15b4104000            call dword ptr [004010b4]
Set var_41 = Nothing
'00712759    8b45e8                  mov eax, dword ptr [ebp-18]
'0071275c    8b08                    mov ecx, dword ptr [eax]
'0071275e    8d5584                  lea edx, dword ptr [ebp-7c]
'00712761    52                      push edx
'00712762    50                      push eax
'00712763    ff5134                  call dword ptr [ecx+34]
'00712766    dbe2                    fnclex
'00712768    85c0                    test ax, ax
'0071276a    7d0e                    jge 71277a
'0071276c    8b4de8                  mov ecx, dword ptr [ebp-18]
'0071276f    6a34                    push 34
'00712771    6830314300              push 00433130
'00712776    51                      push ecx
'00712777    50                      push eax
'00712778    ffd6                    call esi
'0071277a    66837d8400              cmp word ptr [ebp-7c], 00
'0071277f    0f8551010000            jne 7128d6

Do While (0 = 0)
'00712785    8b45e8                  mov eax, dword ptr [ebp-18]
'00712788    8d4ddc                  lea ecx, dword ptr [ebp-24]
'0071278b    51                      push ecx
'0071278c    c745a004000280          mov dword ptr [ebp-60], 80020004
'00712793    c745980a000000          mov dword ptr [ebp-68], 0000000a
'0071279a    8b10                    mov edx, dword ptr [eax]
'0071279c    50                      push eax
'0071279d    ff92b4000000            call dword ptr [edx+000000b4]
'007127a3    dbe2                    fnclex
'007127a5    85c0                    test ax, ax
'007127a7    7d11                    jge 7127ba
'007127a9    8b55e8                  mov edx, dword ptr [ebp-18]
'007127ac    68b4000000              push 000000b4
'007127b1    6830314300              push 00433130
'007127b6    52                      push edx
'007127b7    50                      push eax
'007127b8    ffd6                    call esi
'007127ba    8b45dc                  mov eax, dword ptr [ebp-24]
'007127bd    8b38                    mov edi, dword ptr [eax]
'007127bf    8d5dd8                  lea ebx, dword ptr [ebp-28]
'007127c2    53                      push ebx
'007127c3    83ec10                  sub esp, 10
'007127c6    8bdc                    mov ebx, esp
'007127c8    ba08000000              mov edx, 00000008
'007127cd    8913                    mov dword ptr [ebx], edx
'007127cf    8b55ac                  mov edx, dword ptr [ebp-54]
'007127d2    895304                  mov dword ptr [ebx+04], edx
'007127d5    b964b14300              mov ecx, 0043b164
'007127da    894b08                  mov dword ptr [ebx+08], ecx
'007127dd    8b4db4                  mov ecx, dword ptr [ebp-4c]
'007127e0    50                      push eax
'007127e1    89857cffffff            mov dword ptr [ebp+ffffff7c], eax
'007127e7    894b0c                  mov dword ptr [ebx+0c], ecx
'007127ea    ff5730                  call dword ptr [edi+30]
'007127ed    dbe2                    fnclex
'007127ef    85c0                    test ax, ax
'007127f1    7d11                    jge 712804
'007127f3    8b957cffffff            mov edx, dword ptr [ebp+ffffff7c]
'007127f9    6a30                    push 30
'007127fb    68d8304300              push 004330d8
'00712800    52                      push edx
'00712801    50                      push eax
'00712802    ffd6                    call esi
'00712804    8b45d8                  mov eax, dword ptr [ebp-28]
'00712807    8b08                    mov ecx, dword ptr [eax]
'00712809    8d55c8                  lea edx, dword ptr [ebp-38]
'0071280c    52                      push edx
'0071280d    50                      push eax
'0071280e    8bf8                    mov edi, eax
'00712810    ff5144                  call dword ptr [ecx+44]
'00712813    dbe2                    fnclex
'00712815    85c0                    test ax, ax
'00712817    7d0b                    jge 712824
'00712819    6a44                    push 44
'0071281b    6880304300              push 00433080
'00712820    57                      push edi
'00712821    50                      push eax
'00712822    ffd6                    call esi
'00712824    8b8560ffffff            mov eax, dword ptr [ebp+ffffff60]
'0071282a    8b5598                  mov edx, dword ptr [ebp-68]
'0071282d    8b38                    mov edi, dword ptr [eax]
'0071282f    8b459c                  mov eax, dword ptr [ebp-64]
'00712832    83ec10                  sub esp, 10
'00712835    8bcc                    mov ecx, esp
'00712837    8911                    mov dword ptr [ecx], edx
'00712839    8b55a0                  mov edx, dword ptr [ebp-60]
'0071283c    894104                  mov dword ptr [ecx+04], eax
'0071283f    8b45a4                  mov eax, dword ptr [ebp-5c]
'00712842    895108                  mov dword ptr [ecx+08], edx
'00712845    89410c                  mov dword ptr [ecx+0c], eax
'00712848    8d4dc8                  lea ecx, dword ptr [ebp-38]
'0071284b    51                      push ecx

' *** Reference to "__vbaStrVarMove"
'0071284c    ff153c104000            call dword ptr [0040103c]
'00712852    8bd0                    mov edx, eax
'00712854    8d4de0                  lea ecx, dword ptr [ebp-20]

' *** Reference to "__vbaStrMove"
'00712857    ff15d0124000            call dword ptr [004012d0]
'0071285d    8b9560ffffff            mov edx, dword ptr [ebp+ffffff60]
'00712863    50                      push eax
'00712864    52                      push edx
'00712865    ff97ec010000            call dword ptr [edi+000001ec]
'0071286b    dbe2                    fnclex
'0071286d    85c0                    test ax, ax
'0071286f    7d14                    jge 712885
'00712871    8b8d60ffffff            mov ecx, dword ptr [ebp+ffffff60]
'00712877    68ec010000              push 000001ec
'0071287c    681cb94300              push 0043b91c
'00712881    51                      push ecx
'00712882    50                      push eax
'00712883    ffd6                    call esi
'00712885    8d4de0                  lea ecx, dword ptr [ebp-20]

' *** Reference to "__vbaFreeStr"
'00712888    ff150c134000            call dword ptr [0040130c]
    '#Cleanup(var_3)
'0071288e    8d55d8                  lea edx, dword ptr [ebp-28]
'00712891    52                      push edx
'00712892    8d45dc                  lea eax, dword ptr [ebp-24]
'00712895    50                      push eax
'00712896    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'00712898    ff154c104000            call dword ptr [0040104c]
    '#Cleanup( var_39, var_45)
'0071289e    83c40c                  add esp, 0c
'007128a1    8d4dc8                  lea ecx, dword ptr [ebp-38]

' *** Reference to "__vbaFreeVar"
'007128a4    ff1528104000            call dword ptr [00401028]
    '#Cleanup(var_46)
'007128aa    8b45e8                  mov eax, dword ptr [ebp-18]
'007128ad    8b08                    mov ecx, dword ptr [eax]
'007128af    50                      push eax
'007128b0    ff91ec000000            call dword ptr [ecx+000000ec]
'007128b6    dbe2                    fnclex
'007128b8    85c0                    test ax, ax
'007128ba    0f8d99feffff            jge 712759
'007128c0    8b55e8                  mov edx, dword ptr [ebp-18]
'007128c3    68ec000000              push 000000ec
'007128c8    6830314300              push 00433130
'007128cd    52                      push edx
'007128ce    50                      push eax
'007128cf    ffd6                    call esi
'007128d1    e983feffff              jmp 712759
    
Loop
'007128d6    6a00                    push 00
'007128d8    8d8560ffffff            lea eax, dword ptr [ebp+ffffff60]
'007128de    50                      push eax

' *** Reference to "__vbaObjSetAddref"
'007128df    ff15c8104000            call dword ptr [004010c8]
Set var_20 = Nothing
'007128e5    8b45e8                  mov eax, dword ptr [ebp-18]
'007128e8    8b08                    mov ecx, dword ptr [eax]
'007128ea    50                      push eax
'007128eb    ff91c4000000            call dword ptr [ecx+000000c4]
'007128f1    dbe2                    fnclex
'007128f3    85c0                    test ax, ax
'007128f5    7d11                    jge 712908
'007128f7    8b55e8                  mov edx, dword ptr [ebp-18]
'007128fa    68c4000000              push 000000c4
'007128ff    6830314300              push 00433130
'00712904    52                      push edx
'00712905    50                      push eax
'00712906    ffd6                    call esi
'00712908    8b4508                  mov eax, dword ptr [ebp+08]
'0071290b    8b08                    mov ecx, dword ptr [eax]
'0071290d    50                      push eax
'0071290e    ff91fc020000            call dword ptr [ecx+000002fc]
'00712914    50                      push eax
'00712915    8d955cffffff            lea edx, dword ptr [ebp+ffffff5c]
'0071291b    52                      push edx

' *** Reference to "__vbaObjSet"
'0071291c    ff15b4104000            call dword ptr [004010b4]
Set var_88 = Me
'00712922    8b855cffffff            mov eax, dword ptr [ebp+ffffff5c]
'00712928    8b08                    mov ecx, dword ptr [eax]
'0071292a    50                      push eax
'0071292b    ff91e8010000            call dword ptr [ecx+000001e8]
'00712931    dbe2                    fnclex
'00712933    85c0                    test ax, ax
'00712935    7d14                    jge 71294b
'00712937    8b955cffffff            mov edx, dword ptr [ebp+ffffff5c]
'0071293d    68e8010000              push 000001e8
'00712942    681cb94300              push 0043b91c
'00712947    52                      push edx
'00712948    50                      push eax
'00712949    ffd6                    call esi
'0071294b    8b955cffffff            mov edx, dword ptr [ebp+ffffff5c]
'00712951    8b3a                    mov edi, dword ptr [edx]
'00712953    83ec10                  sub esp, 10
'00712956    8bdc                    mov ebx, esp
'00712958    b90a000000              mov ecx, 0000000a
'0071295d    890b                    mov dword ptr [ebx], ecx
'0071295f    8b4dac                  mov ecx, dword ptr [ebp-54]
'00712962    894b04                  mov dword ptr [ebx+04], ecx
'00712965    b804000280              mov eax, 80020004
'0071296a    894308                  mov dword ptr [ebx+08], eax
'0071296d    8b45b4                  mov eax, dword ptr [ebp-4c]
'00712970    68cc134300              push 004313cc
'00712975    52                      push edx
'00712976    89430c                  mov dword ptr [ebx+0c], eax
'00712979    ff97ec010000            call dword ptr [edi+000001ec]
'0071297f    dbe2                    fnclex
'00712981    85c0                    test ax, ax
'00712983    7d14                    jge 712999
'00712985    8b8d5cffffff            mov ecx, dword ptr [ebp+ffffff5c]
'0071298b    68ec010000              push 000001ec
'00712990    681cb94300              push 0043b91c
'00712995    51                      push ecx
'00712996    50                      push eax
'00712997    ffd6                    call esi
'00712999    8d5ddc                  lea ebx, dword ptr [ebp-24]
'0071299c    53                      push ebx
'0071299d    83ec10                  sub esp, 10
'007129a0    8bdc                    mov ebx, esp
'007129a2    b90a000000              mov ecx, 0000000a
'007129a7    890b                    mov dword ptr [ebx], ecx
'007129a9    894d98                  mov dword ptr [ebp-68], ecx
'007129ac    8b4d8c                  mov ecx, dword ptr [ebp-74]
'007129af    894b04                  mov dword ptr [ebx+04], ecx
'007129b2    8b3d4cb07200            mov edi, dword ptr [0072b04c]
'007129b8    b804000280              mov eax, 80020004
'007129bd    894308                  mov dword ptr [ebx+08], eax
'007129c0    8bd0                    mov edx, eax
'007129c2    8b4594                  mov eax, dword ptr [ebp-6c]
'007129c5    89430c                  mov dword ptr [ebx+0c], eax
'007129c8    8b4598                  mov eax, dword ptr [ebp-68]
'007129cb    83ec10                  sub esp, 10
'007129ce    8bcc                    mov ecx, esp
'007129d0    8901                    mov dword ptr [ecx], eax
'007129d2    8b459c                  mov eax, dword ptr [ebp-64]
'007129d5    894104                  mov dword ptr [ecx+04], eax
'007129d8    895108                  mov dword ptr [ecx+08], edx
'007129db    8955a0                  mov dword ptr [ebp-60], edx
'007129de    8b55a4                  mov edx, dword ptr [ebp-5c]
'007129e1    8b3f                    mov edi, dword ptr [edi]
'007129e3    89510c                  mov dword ptr [ecx+0c], edx
'007129e6    8b55ac                  mov edx, dword ptr [ebp-54]
'007129e9    83ec10                  sub esp, 10
'007129ec    8bc4                    mov eax, esp
'007129ee    c745a803000000          mov dword ptr [ebp-58], 00000003
'007129f5    8b4da8                  mov ecx, dword ptr [ebp-58]
'007129f8    8908                    mov dword ptr [eax], ecx
'007129fa    895004                  mov dword ptr [eax+04], edx
'007129fd    8b55b4                  mov edx, dword ptr [ebp-4c]
'00712a00    c745b002000000          mov dword ptr [ebp-50], 00000002
'00712a07    8b4db0                  mov ecx, dword ptr [ebp-50]
'00712a0a    894808                  mov dword ptr [eax+08], ecx
'00712a0d    89500c                  mov dword ptr [eax+0c], edx
'00712a10    a14cb07200              mov ax, word ptr [0072b04c]
'00712a15    68bc6a4500              push 00456abc
'00712a1a    50                      push eax
'00712a1b    ff97bc000000            call dword ptr [edi+000000bc]
'00712a21    dbe2                    fnclex
'00712a23    85c0                    test ax, ax
'00712a25    7d14                    jge 712a3b
'00712a27    8b0d4cb07200            mov ecx, dword ptr [0072b04c]
'00712a2d    68bc000000              push 000000bc
'00712a32    68301f4300              push 00431f30
'00712a37    51                      push ecx
'00712a38    50                      push eax
'00712a39    ffd6                    call esi
'00712a3b    8b45dc                  mov eax, dword ptr [ebp-24]
'00712a3e    50                      push eax
'00712a3f    8d55e8                  lea edx, dword ptr [ebp-18]
'00712a42    52                      push edx
'00712a43    c745dc00000000          mov dword ptr [ebp-24], 00000000

' *** Reference to "__vbaObjSet"
'00712a4a    ff15b4104000            call dword ptr [004010b4]
Set var_41 = Nothing
'00712a50    8b5d08                  mov ebx, dword ptr [ebp+08]
'00712a53    8b45e8                  mov eax, dword ptr [ebp-18]
'00712a56    8b08                    mov ecx, dword ptr [eax]
'00712a58    8d5584                  lea edx, dword ptr [ebp-7c]
'00712a5b    52                      push edx
'00712a5c    50                      push eax
'00712a5d    ff5134                  call dword ptr [ecx+34]
'00712a60    dbe2                    fnclex
'00712a62    85c0                    test ax, ax
'00712a64    7d0e                    jge 712a74
'00712a66    8b4de8                  mov ecx, dword ptr [ebp-18]
'00712a69    6a34                    push 34
'00712a6b    6830314300              push 00433130
'00712a70    51                      push ecx
'00712a71    50                      push eax
'00712a72    ffd6                    call esi
'00712a74    66837d8400              cmp word ptr [ebp-7c], 00
'00712a79    0f8529020000            jne 712ca8

Do While (0 = 0)
'00712a7f    8b45e8                  mov eax, dword ptr [ebp-18]
'00712a82    8b10                    mov edx, dword ptr [eax]
'00712a84    8d4ddc                  lea ecx, dword ptr [ebp-24]
'00712a87    51                      push ecx
'00712a88    50                      push eax
'00712a89    ff92b4000000            call dword ptr [edx+000000b4]
'00712a8f    dbe2                    fnclex
'00712a91    85c0                    test ax, ax
'00712a93    7d11                    jge 712aa6
'00712a95    8b55e8                  mov edx, dword ptr [ebp-18]
'00712a98    68b4000000              push 000000b4
'00712a9d    6830314300              push 00433130
'00712aa2    52                      push edx
'00712aa3    50                      push eax
'00712aa4    ffd6                    call esi
'00712aa6    8b45dc                  mov eax, dword ptr [ebp-24]
'00712aa9    8b38                    mov edi, dword ptr [eax]
'00712aab    8d5dd8                  lea ebx, dword ptr [ebp-28]
'00712aae    53                      push ebx
'00712aaf    83ec10                  sub esp, 10
'00712ab2    8bdc                    mov ebx, esp
'00712ab4    ba08000000              mov edx, 00000008
'00712ab9    8913                    mov dword ptr [ebx], edx
'00712abb    8b55ac                  mov edx, dword ptr [ebp-54]
'00712abe    895304                  mov dword ptr [ebx+04], edx
'00712ac1    b964b14300              mov ecx, 0043b164
'00712ac6    894b08                  mov dword ptr [ebx+08], ecx
'00712ac9    8b4db4                  mov ecx, dword ptr [ebp-4c]
'00712acc    50                      push eax
'00712acd    89857cffffff            mov dword ptr [ebp+ffffff7c], eax
'00712ad3    894b0c                  mov dword ptr [ebx+0c], ecx
'00712ad6    ff5730                  call dword ptr [edi+30]
'00712ad9    dbe2                    fnclex
'00712adb    85c0                    test ax, ax
'00712add    7d11                    jge 712af0
'00712adf    8b957cffffff            mov edx, dword ptr [ebp+ffffff7c]
'00712ae5    6a30                    push 30
'00712ae7    68d8304300              push 004330d8
'00712aec    52                      push edx
'00712aed    50                      push eax
'00712aee    ffd6                    call esi
'00712af0    8b45d8                  mov eax, dword ptr [ebp-28]
'00712af3    8b08                    mov ecx, dword ptr [eax]
'00712af5    8d55c8                  lea edx, dword ptr [ebp-38]
'00712af8    52                      push edx
'00712af9    50                      push eax
'00712afa    8bf8                    mov edi, eax
'00712afc    ff5144                  call dword ptr [ecx+44]
'00712aff    dbe2                    fnclex
'00712b01    85c0                    test ax, ax
'00712b03    7d0b                    jge 712b10
'00712b05    6a44                    push 44
'00712b07    6880304300              push 00433080
'00712b0c    57                      push edi
'00712b0d    50                      push eax
'00712b0e    ffd6                    call esi
'00712b10    8d45c8                  lea eax, dword ptr [ebp-38]
'00712b13    50                      push eax
'00712b14    8d4d98                  lea ecx, dword ptr [ebp-68]
'00712b17    51                      push ecx
'00712b18    c745a0c84a4500          mov dword ptr [ebp-60], 00454ac8
'00712b1f    c7459808800000          mov dword ptr [ebp-68], 00008008

' *** Reference to "__vbaVarTstNe"
'00712b26    ff1574124000            call dword ptr [00401274]
'00712b2c    8d55d8                  lea edx, dword ptr [ebp-28]
'00712b2f    668bf8                  mov di, ax
'00712b32    52                      push edx
'00712b33    8d45dc                  lea eax, dword ptr [ebp-24]
'00712b36    50                      push eax
'00712b37    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'00712b39    ff154c104000            call dword ptr [0040104c]
    '#Cleanup( var_39, var_45)
'00712b3f    83c40c                  add esp, 0c
'00712b42    8d4dc8                  lea ecx, dword ptr [ebp-38]

' *** Reference to "__vbaFreeVar"
'00712b45    ff1528104000            call dword ptr [00401028]
    '#Cleanup(var_46)
'00712b4b    6685ff                  test edi, edi
'00712b4e    0f8425010000            je 712c79
    
    If (    ((var_46) <> ("Modificateurs"))) Then
'00712b54    8b45e8                  mov eax, dword ptr [ebp-18]
'00712b57    8d55dc                  lea edx, dword ptr [ebp-24]
'00712b5a    52                      push edx
'00712b5b    c745a004000280          mov dword ptr [ebp-60], 80020004
'00712b62    c745980a000000          mov dword ptr [ebp-68], 0000000a
'00712b69    8b08                    mov ecx, dword ptr [eax]
'00712b6b    50                      push eax
'00712b6c    ff91b4000000            call dword ptr [ecx+000000b4]
'00712b72    dbe2                    fnclex
'00712b74    85c0                    test ax, ax
'00712b76    7d11                    jge 712b89
'00712b78    8b4de8                  mov ecx, dword ptr [ebp-18]
'00712b7b    68b4000000              push 000000b4
'00712b80    6830314300              push 00433130
'00712b85    51                      push ecx
'00712b86    50                      push eax
'00712b87    ffd6                    call esi
'00712b89    8b45dc                  mov eax, dword ptr [ebp-24]
'00712b8c    8b38                    mov edi, dword ptr [eax]
'00712b8e    8d5dd8                  lea ebx, dword ptr [ebp-28]
'00712b91    53                      push ebx
'00712b92    83ec10                  sub esp, 10
'00712b95    8bdc                    mov ebx, esp
'00712b97    ba08000000              mov edx, 00000008
'00712b9c    8913                    mov dword ptr [ebx], edx
'00712b9e    8b55ac                  mov edx, dword ptr [ebp-54]
'00712ba1    895304                  mov dword ptr [ebx+04], edx
'00712ba4    b964b14300              mov ecx, 0043b164
'00712ba9    894b08                  mov dword ptr [ebx+08], ecx
'00712bac    8b4db4                  mov ecx, dword ptr [ebp-4c]
'00712baf    50                      push eax
'00712bb0    89857cffffff            mov dword ptr [ebp+ffffff7c], eax
'00712bb6    894b0c                  mov dword ptr [ebx+0c], ecx
'00712bb9    ff5730                  call dword ptr [edi+30]
'00712bbc    dbe2                    fnclex
'00712bbe    85c0                    test ax, ax
'00712bc0    7d11                    jge 712bd3
'00712bc2    8b957cffffff            mov edx, dword ptr [ebp+ffffff7c]
'00712bc8    6a30                    push 30
'00712bca    68d8304300              push 004330d8
'00712bcf    52                      push edx
'00712bd0    50                      push eax
'00712bd1    ffd6                    call esi
'00712bd3    8b45d8                  mov eax, dword ptr [ebp-28]
'00712bd6    8b08                    mov ecx, dword ptr [eax]
'00712bd8    8d55c8                  lea edx, dword ptr [ebp-38]
'00712bdb    52                      push edx
'00712bdc    50                      push eax
'00712bdd    8bf8                    mov edi, eax
'00712bdf    ff5144                  call dword ptr [ecx+44]
'00712be2    dbe2                    fnclex
'00712be4    85c0                    test ax, ax
'00712be6    7d0b                    jge 712bf3
    
    If (    0) Then
'00712be8    6a44                    push 44
'00712bea    6880304300              push 00433080
'00712bef    57                      push edi
'00712bf0    50                      push eax
'00712bf1    ffd6                    call esi
    
End If
'00712bf3    8b855cffffff            mov eax, dword ptr [ebp+ffffff5c]
'00712bf9    8b5598                  mov edx, dword ptr [ebp-68]
'00712bfc    8b38                    mov edi, dword ptr [eax]
'00712bfe    8b459c                  mov eax, dword ptr [ebp-64]
'00712c01    83ec10                  sub esp, 10
'00712c04    8bcc                    mov ecx, esp
'00712c06    8911                    mov dword ptr [ecx], edx
'00712c08    8b55a0                  mov edx, dword ptr [ebp-60]
'00712c0b    894104                  mov dword ptr [ecx+04], eax
'00712c0e    8b45a4                  mov eax, dword ptr [ebp-5c]
'00712c11    895108                  mov dword ptr [ecx+08], edx
'00712c14    89410c                  mov dword ptr [ecx+0c], eax
'00712c17    8d4dc8                  lea ecx, dword ptr [ebp-38]
'00712c1a    51                      push ecx

' *** Reference to "__vbaStrVarMove"
'00712c1b    ff153c104000            call dword ptr [0040103c]
'00712c21    8bd0                    mov edx, eax
'00712c23    8d4de0                  lea ecx, dword ptr [ebp-20]

' *** Reference to "__vbaStrMove"
'00712c26    ff15d0124000            call dword ptr [004012d0]
'00712c2c    8b955cffffff            mov edx, dword ptr [ebp+ffffff5c]
'00712c32    50                      push eax
'00712c33    52                      push edx
'00712c34    ff97ec010000            call dword ptr [edi+000001ec]
'00712c3a    dbe2                    fnclex
'00712c3c    85c0                    test ax, ax
'00712c3e    7d14                    jge 712c54
'00712c40    8b8d5cffffff            mov ecx, dword ptr [ebp+ffffff5c]
'00712c46    68ec010000              push 000001ec
'00712c4b    681cb94300              push 0043b91c
'00712c50    51                      push ecx
'00712c51    50                      push eax
'00712c52    ffd6                    call esi
'00712c54    8d4de0                  lea ecx, dword ptr [ebp-20]

' *** Reference to "__vbaFreeStr"
'00712c57    ff150c134000            call dword ptr [0040130c]
'#Cleanup(var_3)
'00712c5d    8d55d8                  lea edx, dword ptr [ebp-28]
'00712c60    52                      push edx
'00712c61    8d45dc                  lea eax, dword ptr [ebp-24]
'00712c64    50                      push eax
'00712c65    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'00712c67    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_39, var_45)
'00712c6d    83c40c                  add esp, 0c
'00712c70    8d4dc8                  lea ecx, dword ptr [ebp-38]

' *** Reference to "__vbaFreeVar"
'00712c73    ff1528104000            call dword ptr [00401028]
'#Cleanup(var_46)
'ERROR: Two many next close:
End If
'00712c79    8b45e8                  mov eax, dword ptr [ebp-18]
'00712c7c    8b08                    mov ecx, dword ptr [eax]
'00712c7e    8b5d08                  mov ebx, dword ptr [ebp+08]
'00712c81    50                      push eax
'00712c82    ff91ec000000            call dword ptr [ecx+000000ec]
'00712c88    dbe2                    fnclex
'00712c8a    85c0                    test ax, ax
'00712c8c    0f8dc1fdffff            jge 712a53
'00712c92    8b55e8                  mov edx, dword ptr [ebp-18]
'00712c95    68ec000000              push 000000ec
'00712c9a    6830314300              push 00433130
'00712c9f    52                      push edx
'00712ca0    50                      push eax
'00712ca1    ffd6                    call esi
'00712ca3    e9abfdffff              jmp 712a53

'ERROR: Two many next close:
Loop
'00712ca8    8b855cffffff            mov eax, dword ptr [ebp+ffffff5c]
'00712cae    8b08                    mov ecx, dword ptr [eax]
'00712cb0    6a00                    push 00
'00712cb2    50                      push eax
'00712cb3    ff9194000000            call dword ptr [ecx+00000094]
'00712cb9    dbe2                    fnclex
'00712cbb    85c0                    test ax, ax
'00712cbd    7d14                    jge 712cd3
'00712cbf    8b955cffffff            mov edx, dword ptr [ebp+ffffff5c]
'00712cc5    6894000000              push 00000094
'00712cca    681cb94300              push 0043b91c
'00712ccf    52                      push edx
'00712cd0    50                      push eax
'00712cd1    ffd6                    call esi
'00712cd3    6a00                    push 00
'00712cd5    8d855cffffff            lea eax, dword ptr [ebp+ffffff5c]
'00712cdb    50                      push eax

' *** Reference to "__vbaObjSetAddref"
'00712cdc    ff15c8104000            call dword ptr [004010c8]
Set var_88 = Nothing
'00712ce2    8b45e8                  mov eax, dword ptr [ebp-18]
'00712ce5    8b08                    mov ecx, dword ptr [eax]
'00712ce7    50                      push eax
'00712ce8    ff91c4000000            call dword ptr [ecx+000000c4]
'00712cee    dbe2                    fnclex
'00712cf0    85c0                    test ax, ax
'00712cf2    7d11                    jge 712d05
'00712cf4    8b55e8                  mov edx, dword ptr [ebp-18]
'00712cf7    68c4000000              push 000000c4
'00712cfc    6830314300              push 00433130
'00712d01    52                      push edx
'00712d02    50                      push eax
'00712d03    ffd6                    call esi
'00712d05    8b03                    mov eax, dword ptr [ebx]
'00712d07    53                      push ebx
'00712d08    ff9008030000            call dword ptr [eax+00000308]
'00712d0e    50                      push eax
'00712d0f    8d4ddc                  lea ecx, dword ptr [ebp-24]
'00712d12    51                      push ecx

' *** Reference to "__vbaObjSet"
'00712d13    ff15b4104000            call dword ptr [004010b4]
Set var_39 = Nothing
'00712d19    8bf8                    mov edi, eax
'00712d1b    8b17                    mov edx, dword ptr [edi]
'00712d1d    8d4584                  lea eax, dword ptr [ebp-7c]
'00712d20    50                      push eax
'00712d21    57                      push edi
'00712d22    ff92e8000000            call dword ptr [edx+000000e8]
'00712d28    dbe2                    fnclex
'00712d2a    85c0                    test ax, ax
'00712d2c    7d0e                    jge 712d3c
'00712d2e    68e8000000              push 000000e8
'00712d33    681cb94300              push 0043b91c
'00712d38    57                      push edi
'00712d39    50                      push eax
'00712d3a    ffd6                    call esi
'00712d3c    33c9                    xor ecx, ecx
'00712d3e    66394d84                cmp word ptr [ebp-7c], cx
'00712d42    0f9fc1                  setg cl
'00712d45    f7d9                    neg ecx
'00712d47    8bf9                    mov edi, ecx
'00712d49    8d4ddc                  lea ecx, dword ptr [ebp-24]

' *** Reference to "__vbaFreeObj"
'00712d4c    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_39)
'00712d52    6685ff                  test edi, edi
'00712d55    741f                    je 712d76

If (0 > 0) Then
'00712d57    8b13                    mov edx, dword ptr [ebx]
'00712d59    8d4584                  lea eax, dword ptr [ebp-7c]
'00712d5c    50                      push eax
'00712d5d    53                      push ebx
'00712d5e    c7458400000000          mov dword ptr [ebp-7c], 00000000
'00712d65    ff92f8060000            call dword ptr [edx+000006f8]
'00712d6b    85c0                    test ax, ax
'00712d6d    7d22                    jge 712d91
'00712d6f    68f8060000              push 000006f8
'00712d74    eb12                    jmp 712d88
    
End If
'00712d76    8b0b                    mov ecx, dword ptr [ebx]
'00712d78    53                      push ebx
'00712d79    ff91fc060000            call dword ptr [ecx+000006fc]
'00712d7f    85c0                    test ax, ax
'00712d81    7d0e                    jge 712d91
'00712d83    68fc060000              push 000006fc
'00712d88    68e4194300              push 004319e4
'00712d8d    53                      push ebx
'00712d8e    50                      push eax
'00712d8f    ffd6                    call esi
'00712d91    c745fc00000000          mov dword ptr [ebp-04], 00000000
'00712d98    68042e7100              push 00712e04
'00712d9d    eb2d                    jmp 712dcc
'00712d9f    8d4de0                  lea ecx, dword ptr [ebp-20]

' *** Reference to "__vbaFreeStr"
'00712da2    ff150c134000            call dword ptr [0040130c]
'#Cleanup(var_3)
'00712da8    8d55d8                  lea edx, dword ptr [ebp-28]
'00712dab    52                      push edx
'00712dac    8d45dc                  lea eax, dword ptr [ebp-24]
'00712daf    50                      push eax
'00712db0    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'00712db2    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_39, var_45)
'00712db8    8d4db8                  lea ecx, dword ptr [ebp-48]
'00712dbb    51                      push ecx
'00712dbc    8d55c8                  lea edx, dword ptr [ebp-38]
'00712dbf    52                      push edx
'00712dc0    6a02                    push 02

' *** Reference to "__vbaFreeVarList"
'00712dc2    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_46, var_61)
'00712dc8    83c418                  add esp, 18
'00712dcb    c3                      ret
'00712dcc    8d855cffffff            lea eax, dword ptr [ebp+ffffff5c]
'00712dd2    50                      push eax
'00712dd3    8d8d60ffffff            lea ecx, dword ptr [ebp+ffffff60]
'00712dd9    51                      push ecx
'00712dda    8d9564ffffff            lea edx, dword ptr [ebp+ffffff64]
'00712de0    52                      push edx
'00712de1    8d8568ffffff            lea eax, dword ptr [ebp+ffffff68]
'00712de7    50                      push eax
'00712de8    6a04                    push 04

' *** Reference to "__vbaFreeObjList"
'00712dea    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_132, var_123, var_20, var_88)

' *** Reference to "__vbaFreeObj"
'00712df0    8b3508134000            mov esi, dword ptr [00401308]
'00712df6    83c414                  add esp, 14
'00712df9    8d4de8                  lea ecx, dword ptr [ebp-18]
'00712dfc    ffd6                    call esi
'#Cleanup(var_41)
'00712dfe    8d4de4                  lea ecx, dword ptr [ebp-1c]
'00712e01    ffd6                    call esi
'#Cleanup(var_40)
'00712e03    c3                      ret
'00712e04    8b4508                  mov eax, dword ptr [ebp+08]
'00712e07    8b08                    mov ecx, dword ptr [eax]
'00712e09    50                      push eax
'00712e0a    ff5108                  call dword ptr [ecx+08]
'00712e0d    8b45fc                  mov eax, dword ptr [ebp-04]
'00712e10    8b4dec                  mov ecx, dword ptr [ebp-14]
'00712e13    5f                      pop edi
'00712e14    5e                      pop esi
    'Reference to '__except_list'
'00712e15    64890d00000000          mov dword ptr fs:[00000000], ecx
'00712e1c    5b                      pop ebx
'00712e1d    8be5                    mov esp, ebp
'00712e1f    5d                      pop ebp
'00712e20    c20400                  ret 0004


End Sub


