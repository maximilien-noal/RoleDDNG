VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "C:\Program Files (x86)\VBReFormer\Lib\MSCOMCTL.OCX"

Begin VB.MDIForm frmMain
    Caption = "RoleDD Gestion de personnages pour Donjons et Dragons 3.5"
    BackColor = -2147483636
    Icon = frmMain.frx:0000
    LinkTopic = "MDIForm1"
    ClientLeft   = 2250
    ClientTop    = -1155
    ClientWidth  = 18240
    ClientHeight = 10830
    OLEDropMode = 1
    Begin MSComctlLib.StatusBar StatusBar1
        Left   = 0
        Top    = 10500
        Width  = 18240
        Height = 330
        TabIndex = 0
        OleObjectBlob = frmMain.frx:0EDE
        Align = 2
    End
    Begin VB.Menu IDM_FICHIER
        Caption = "&Fichier"
        Begin VB.Menu IDM_ASSISTANT_PERSO
            Caption = "&Assistant création personnage"
            ShortCut = 14
        End
        Begin VB.Menu IDM_GESTION_PERSO
            Caption = "&Gestion des personnages"
            ShortCut = 7
        End
        Begin VB.Menu IDM_ACCEDER
            Caption = "&Accéder aux fiches des personnages"
        End
        Begin VB.Menu IDM_GEST_OBJETS
            Caption = "&Gestion des objets"
            ShortCut = 9
        End
        Begin VB.Menu IDM_SEP_3
            Caption = "-"
        End
        Begin VB.Menu IMD_IMPORT
            Caption = "&Importation de personnage"
            ShortCut = 13
        End
        Begin VB.Menu IDM_SEP_1
            Caption = "-"
        End
        Begin VB.Menu IDM_OUVRIR
            Caption = "&Ouvrir une base de donnée de personnages"
            ShortCut = 15
        End
        Begin VB.Menu IDM_SEP2
            Caption = "-"
        End
        Begin VB.Menu IDM_QUITTER
            Caption = "&Quitter"
        End
    End
    Begin VB.Menu IDM_OUTILS
        Caption = "&Outils"
        Begin VB.Menu IDM_XP_PERSONNAGES
            Caption = "&Expériences des personnages"
            ShortCut = 24
        End
        Begin VB.Menu IDM_GENERATEUR_VILLE
            Caption = "&Générateur de ville"
        End
        Begin VB.Menu IDM_LANCEUR_DES
            Caption = "&Lanceur de dés"
            ShortCut = 4
        End
    End
    Begin VB.Menu IDM_REGLES
        Caption = "&Règles"
        Begin VB.Menu IDM_DESCRIPTION_SORT
            Caption = "Description des &sorts"
            ShortCut = 19
        End
        Begin VB.Menu IDM_DESCRIPTION_RACE
            Caption = "Description des &races"
        End
        Begin VB.Menu IDM_DESCRIPTION_DON
            Caption = "&Description des &dons"
        End
    End
    Begin VB.Menu IDM_PERSONNEL
        Caption = "&Personnel"
        Begin VB.Menu IDM_INSERTION_DONS
            Caption = "&Insertion de dons"
        End
        Begin VB.Menu IDM_INSERTION_RACES
            Caption = "&Insertion de races"
        End
        Begin VB.Menu IDM_INSERTION_CLASSES
            Caption = "&Insertion de classes"
        End
    End
    Begin VB.Menu IDM_OPTIONS
        Caption = "&Options"
    End
    Begin VB.Menu IDM_AIDE
        Caption = "&Aide"
        Begin VB.Menu IDM_ASTUCE
            Caption = "&Astuce du jour"
        End
        Begin VB.Menu IDM_LIENSITE
            Caption = "&Lien vers le site"
        End
        Begin VB.Menu IDM_MISES_JOUR
            Caption = "&Vérifier les mises à jour"
        End
        Begin VB.Menu IDM_APROPOS
            Caption = "&A propos de RoleDD"
        End
    End
End
Public Function TrouveChild(arg_0 As Unknow, arg_1 As Unknow, arg_2 As Unknow, arg_3 As Unknow, arg_4 As Unknow, arg_5 As Unknow, arg_6 As Unknow, arg_7 As Unknow, arg_8 As Unknow, arg_9 As Unknow, arg_A As Unknow, arg_B As Unknow, arg_C As Unknow, arg_D As Unknow, arg_E As Unknow, arg_F As Unknow, arg_10 As Unknow, arg_11 As Unknow, arg_12 As Unknow, arg_13 As Unknow, arg_14 As Unknow, arg_15 As Unknow, arg_16 As Unknow, arg_17 As Unknow, arg_18 As Unknow, arg_19 As Unknow, arg_1A As Unknow, arg_1B As Unknow, arg_1C As Unknow, arg_1D As Unknow, arg_1E As Unknow, arg_1F As Unknow, arg_20 As Unknow, arg_21 As Unknow, arg_22 As Unknow, arg_23 As Unknow, arg_24 As Unknow, arg_25 As Unknow, arg_26 As Unknow, arg_27 As Unknow, arg_28 As Unknow, arg_29 As Unknow, arg_2A As Unknow, arg_2B As Unknow, arg_2C As Unknow, arg_2D As Unknow, arg_2E As Unknow, arg_2F As Unknow, arg_30 As Unknow, arg_31 As Unknow, arg_32 As Unknow, arg_33 As Unknow, arg_34 As Unknow, arg_35 As Unknow, arg_36 As Unknow, arg_37 As Unknow, arg_38 As Unknow, arg_39 As Unknow, arg_3A As Unknow, arg_3B As Unknow, arg_3C As Unknow)
'004916f0    55                      push ebp
'004916f1    8bec                    mov ebp, esp
'004916f3    83ec0c                  sub esp, 0c

' *** Reference to "__vbaExceptHandler"
'004916f6    6866724000              push 00407266
'004916fb    64a100000000            mov ax, word ptr fs:[00000000]
'00491701    50                      push eax
    'Reference to '__except_list'
'00491702    64892500000000          mov dword ptr fs:[00000000], esp
'00491709    81ec98000000            sub esp, 00000098
'0049170f    53                      push ebx
'00491710    56                      push esi
'00491711    57                      push edi
'00491712    8965f4                  mov dword ptr [ebp-0c], esp
'00491715    c745f818134000          mov dword ptr [ebp-08], 00401318
'0049171c    33ff                    xor edi, edi
'0049171e    897dfc                  mov dword ptr [ebp-04], edi
'00491721    8b4508                  mov eax, dword ptr [ebp+08]
'00491724    8b08                    mov ecx, dword ptr [eax]
'00491726    50                      push eax
'00491727    ff5104                  call dword ptr [ecx+04]
'0049172a    393d24be7200            cmp dword ptr [0072be24], edi
'00491730    897de4                  mov dword ptr [ebp-1c], edi
'00491733    897de8                  mov dword ptr [ebp-18], edi
'00491736    897de0                  mov dword ptr [ebp-20], edi
'00491739    897dd0                  mov dword ptr [ebp-30], edi
'0049173c    897dc0                  mov dword ptr [ebp-40], edi
'0049173f    897db0                  mov dword ptr [ebp-50], edi
'00491742    897da0                  mov dword ptr [ebp-60], edi
'00491745    897d90                  mov dword ptr [ebp-70], edi
'00491748    897d80                  mov dword ptr [ebp-80], edi
'0049174b    89bd70ffffff            mov dword ptr [ebp+ffffff70], edi
'00491751    c745e4ffffffff          mov dword ptr [ebp-1c], ffffffff
'00491758    7510                    jne 49176a
'0049175a    6824be7200              push 0072be24
'0049175f    6870004300              push 00430070

' *** Reference to "__vbaNew2"
'00491764    ff152c124000            call dword ptr [0040122c]
Dim var_2 As New Global
'0049176a    8b3524be7200            mov esi, dword ptr [0072be24]
'00491770    8b16                    mov edx, dword ptr [esi]
'00491772    8d45e0                  lea eax, dword ptr [ebp-20]
'00491775    50                      push eax
'00491776    56                      push esi
'00491777    ff5228                  call dword ptr [edx+28]
Set var_3 = var_2.Forms
'0049177a    dbe2                    fnclex
'0049177c    3bc7                    cmp eax, edi
'0049177e    7d0f                    jge 49178f
'00491780    6a28                    push 28
'00491782    6860004300              push 00430060
'00491787    56                      push esi
'00491788    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00491789    ff1580104000            call dword ptr [00401080]
'0049178f    8b4de0                  mov ecx, dword ptr [ebp-20]
'00491792    57                      push edi
'00491793    6880004300              push 00430080
'00491798    51                      push ecx
'00491799    8d55d0                  lea edx, dword ptr [ebp-30]
'0049179c    52                      push edx
'0049179d    c7458801000000          mov dword ptr [ebp-78], 00000001
'004917a4    c7458002000000          mov dword ptr [ebp-80], 00000002

' *** Reference to "__vbaLateMemCallLd"
'004917ab    ff15c4124000            call dword ptr [004012c4]
var_4 = var_3.Count()
'004917b1    83c410                  add esp, 10
'004917b4    50                      push eax
'004917b5    8d4580                  lea eax, dword ptr [ebp-80]
'004917b8    50                      push eax
'004917b9    8d4dc0                  lea ecx, dword ptr [ebp-40]
'004917bc    51                      push ecx

' *** Reference to "__vbaVarSub"
'004917bd    ff1508104000            call dword ptr [00401008]
'004917c3    50                      push eax

' *** Reference to "__vbaI2Var"
'004917c4    ff150c124000            call dword ptr [0040120c]
'004917ca    8d4de0                  lea ecx, dword ptr [ebp-20]
'004917cd    897de8                  mov dword ptr [ebp-18], edi
'004917d0    89855cffffff            mov dword ptr [ebp+ffffff5c], eax

' *** Reference to "__vbaFreeObj"
'004917d6    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_3)
'004917dc    8d4dd0                  lea ecx, dword ptr [ebp-30]

' *** Reference to "__vbaFreeVar"
'004917df    ff1528104000            call dword ptr [00401028]
'#Cleanup(var_4)
'004917e5    8b5d0c                  mov ebx, dword ptr [ebp+0c]

' *** Reference to "rtcUpperCaseVar"
'004917e8    8b3d2c114000            mov edi, dword ptr [0040112c]
'004917ee    8b45e8                  mov eax, dword ptr [ebp-18]
'004917f1    663b855cffffff          cmp ax, word ptr [ebp+ffffff5c]
'004917f8    0f8f07010000            jg 491905

If (0 <= WORD PTR [EBP+FFFFFF5C]) Then
'004917fe    a124be7200              mov ax, word ptr [0072be24]
'00491803    85c0                    test ax, ax
'00491805    8d55e8                  lea edx, dword ptr [ebp-18]
'00491808    895588                  mov dword ptr [ebp-78], edx
'0049180b    c7458002400000          mov dword ptr [ebp-80], 00004002
'00491812    7510                    jne 491824
'00491814    6824be7200              push 0072be24
'00491819    6870004300              push 00430070

' *** Reference to "__vbaNew2"
'0049181e    ff152c124000            call dword ptr [0040122c]
    Set var_2 = New Global
'00491824    8b3524be7200            mov esi, dword ptr [0072be24]
'0049182a    8b06                    mov eax, dword ptr [esi]
'0049182c    8d4de0                  lea ecx, dword ptr [ebp-20]
'0049182f    51                      push ecx
'00491830    56                      push esi
'00491831    ff5028                  call dword ptr [eax+28]
    Set var_3 = var_2.Forms
'00491834    dbe2                    fnclex
'00491836    85c0                    test ax, ax
'00491838    7d0f                    jge 491849
'0049183a    6a28                    push 28
'0049183c    6860004300              push 00430060
'00491841    56                      push esi
'00491842    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00491843    ff1580104000            call dword ptr [00401080]
'00491849    8b4580                  mov eax, dword ptr [ebp-80]
'0049184c    8b4d84                  mov ecx, dword ptr [ebp-7c]
'0049184f    6a00                    push 00
'00491851    688c004300              push 0043008c
'00491856    83ec10                  sub esp, 10
'00491859    8bd4                    mov edx, esp
'0049185b    8902                    mov dword ptr [edx], eax
'0049185d    8b4588                  mov eax, dword ptr [ebp-78]
'00491860    894a04                  mov dword ptr [edx+04], ecx
'00491863    8b4d8c                  mov ecx, dword ptr [ebp-74]
'00491866    894208                  mov dword ptr [edx+08], eax
'00491869    6a01                    push 01
'0049186b    894a0c                  mov dword ptr [edx+0c], ecx
'0049186e    8b55e0                  mov edx, dword ptr [ebp-20]
'00491871    6a00                    push 00
'00491873    52                      push edx
'00491874    8d45d0                  lea eax, dword ptr [ebp-30]
'00491877    50                      push eax

' *** Reference to "__vbaLateIdCallLd"
'00491878    ff157c114000            call dword ptr [0040117c]
    var_4 = var_3.UNK__0
'0049187e    83c420                  add esp, 20
'00491881    50                      push eax
'00491882    8d4dc0                  lea ecx, dword ptr [ebp-40]
'00491885    51                      push ecx

' *** Reference to "__vbaVarLateMemCallLdRf"
'00491886    ff1530124000            call dword ptr [00401230]
    var_5 = var_4.()
'0049188c    83c410                  add esp, 10
'0049188f    50                      push eax
'00491890    8d55b0                  lea edx, dword ptr [ebp-50]
'00491893    52                      push edx
'00491894    ffd7                    call edi
'00491896    8d8570ffffff            lea eax, dword ptr [ebp+ffffff70]
'0049189c    50                      push eax
'0049189d    8d4da0                  lea ecx, dword ptr [ebp-60]
'004918a0    51                      push ecx
'004918a1    899d78ffffff            mov dword ptr [ebp+ffffff78], ebx
'004918a7    c78570ffffff08400000    mov dword ptr [ebp+ffffff70], 00004008
'004918b1    ffd7                    call edi
'004918b3    8d55b0                  lea edx, dword ptr [ebp-50]
'004918b6    52                      push edx
'004918b7    8d45a0                  lea eax, dword ptr [ebp-60]
'004918ba    50                      push eax

' *** Reference to "__vbaVarTstEq"
'004918bb    ff153c114000            call dword ptr [0040113c]
'004918c1    8d4de0                  lea ecx, dword ptr [ebp-20]
'004918c4    8bf0                    mov esi, eax

' *** Reference to "__vbaFreeObj"
'004918c6    ff1508134000            call dword ptr [00401308]
    '#Cleanup(var_3)
'004918cc    8d4da0                  lea ecx, dword ptr [ebp-60]
'004918cf    51                      push ecx
'004918d0    8d55b0                  lea edx, dword ptr [ebp-50]
'004918d3    52                      push edx
'004918d4    8d45c0                  lea eax, dword ptr [ebp-40]
'004918d7    50                      push eax
'004918d8    8d4dd0                  lea ecx, dword ptr [ebp-30]
'004918db    51                      push ecx
'004918dc    6a04                    push 04

' *** Reference to "__vbaFreeVarList"
'004918de    ff1540104000            call dword ptr [00401040]
    '#Cleanup( var_4, var_5, var_6, var_7)
'004918e4    83c414                  add esp, 14
'004918e7    6685f6                  test esi, esi
'004918ea    7513                    jne 4918ff
    
    Do While (    ((UCase(var_5)) = (UCase(arg_0))))
'004918ec    b801000000              mov eax, 00000001
'004918f1    660345e8                add ax, word ptr [ebp-18]
    var_num1 = 1 + 0
'004918f5    7068                    jo 49195f
'004918f7    8945e8                  mov dword ptr [ebp-18], eax
'004918fa    e9f2feffff              jmp 4917f1
    
Loop
'004918ff    8b55e8                  mov edx, dword ptr [ebp-18]
'00491902    8955e4                  mov dword ptr [ebp-1c], edx

'ERROR: Two many next close:
End If
'00491905    6836194900              push 00491936
'0049190a    eb29                    jmp 491935
'0049190c    8d4de0                  lea ecx, dword ptr [ebp-20]

' *** Reference to "__vbaFreeObj"
'0049190f    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_3)
'00491915    8d4590                  lea eax, dword ptr [ebp-70]
'00491918    50                      push eax
'00491919    8d4da0                  lea ecx, dword ptr [ebp-60]
'0049191c    51                      push ecx
'0049191d    8d55b0                  lea edx, dword ptr [ebp-50]
'00491920    52                      push edx
'00491921    8d45c0                  lea eax, dword ptr [ebp-40]
'00491924    50                      push eax
'00491925    8d4dd0                  lea ecx, dword ptr [ebp-30]
'00491928    51                      push ecx
'00491929    6a05                    push 05

' *** Reference to "__vbaFreeVarList"
'0049192b    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_4, var_5, var_6, var_7, var_8)
'00491931    83c418                  add esp, 18
'00491934    c3                      ret
'00491935    c3                      ret
'00491936    8b4508                  mov eax, dword ptr [ebp+08]
'00491939    8b10                    mov edx, dword ptr [eax]
'0049193b    50                      push eax
'0049193c    ff5208                  call dword ptr [edx+08]
'0049193f    8b4510                  mov eax, dword ptr [ebp+10]
'00491942    668b4de4                mov cx, word ptr [ebp-1c]
'00491946    668908                  mov word ptr [eax], cx
'00491949    8b45fc                  mov eax, dword ptr [ebp-04]
'0049194c    8b4dec                  mov ecx, dword ptr [ebp-14]
'0049194f    5f                      pop edi
'00491950    5e                      pop esi
    'Reference to '__except_list'
'00491951    64890d00000000          mov dword ptr fs:[00000000], ecx
'00491958    5b                      pop ebx
'00491959    8be5                    mov esp, ebp
'0049195b    5d                      pop ebp
'0049195c    c20c00                  ret 000c


End Function


Private Function sub_4925F0(arg_0 As Unknow, arg_1 As Unknow, arg_2 As Unknow, arg_3 As Unknow, arg_4 As Unknow, arg_5 As Unknow, arg_6 As Unknow, arg_7 As Unknow, arg_8 As Unknow, arg_9 As Unknow, arg_A As Unknow, arg_B As Unknow, arg_C As Unknow, arg_D As Unknow, arg_E As Unknow, arg_F As Unknow, arg_10 As Unknow, arg_11 As Unknow, arg_12 As Unknow, arg_13 As Unknow, arg_14 As Unknow, arg_15 As Unknow, arg_16 As Unknow, arg_17 As Unknow, arg_18 As Unknow, arg_19 As Unknow, arg_1A As Unknow, arg_1B As Unknow, arg_1C As Unknow, arg_1D As Unknow, arg_1E As Unknow, arg_1F As Unknow, arg_20 As Unknow, arg_21 As Unknow, arg_22 As Unknow, arg_23 As Unknow, arg_24 As Unknow, arg_25 As Unknow, arg_26 As Unknow, arg_27 As Unknow, arg_28 As Unknow, arg_29 As Unknow, arg_2A As Unknow, arg_2B As Unknow, arg_2C As Unknow, arg_2D As Unknow, arg_2E As Unknow, arg_2F As Unknow, arg_30 As Unknow, arg_31 As Unknow, arg_32 As Unknow, arg_33 As Unknow, arg_34 As Unknow, arg_35 As Unknow, arg_36 As Unknow, arg_37 As Unknow, arg_38 As Unknow, arg_39 As Unknow, arg_3A As Unknow, arg_3B As Unknow, arg_3C As Unknow)
'004925f0    33c0                    xor eax, eax
'004925f2    c20400                  ret 0004


End Function


Private Function sub_4957E0(arg_0 As Unknow, arg_1 As Unknow, arg_2 As Unknow, arg_3 As Unknow, arg_4 As Unknow, arg_5 As Unknow, arg_6 As Unknow, arg_7 As Unknow, arg_8 As Unknow, arg_9 As Unknow, arg_A As Unknow, arg_B As Unknow, arg_C As Unknow, arg_D As Unknow, arg_E As Unknow, arg_F As Unknow, arg_10 As Unknow, arg_11 As Unknow, arg_12 As Unknow, arg_13 As Unknow, arg_14 As Unknow, arg_15 As Unknow, arg_16 As Unknow, arg_17 As Unknow, arg_18 As Unknow, arg_19 As Unknow, arg_1A As Unknow, arg_1B As Unknow, arg_1C As Unknow, arg_1D As Unknow, arg_1E As Unknow, arg_1F As Unknow, arg_20 As Unknow, arg_21 As Unknow, arg_22 As Unknow, arg_23 As Unknow, arg_24 As Unknow, arg_25 As Unknow, arg_26 As Unknow, arg_27 As Unknow, arg_28 As Unknow, arg_29 As Unknow, arg_2A As Unknow, arg_2B As Unknow, arg_2C As Unknow, arg_2D As Unknow, arg_2E As Unknow, arg_2F As Unknow, arg_30 As Unknow, arg_31 As Unknow, arg_32 As Unknow, arg_33 As Unknow, arg_34 As Unknow, arg_35 As Unknow, arg_36 As Unknow, arg_37 As Unknow, arg_38 As Unknow, arg_39 As Unknow, arg_3A As Unknow, arg_3B As Unknow, arg_3C As Unknow)
'004957e0    55                      push ebp
'004957e1    8bec                    mov ebp, esp
'004957e3    83ec08                  sub esp, 08

' *** Reference to "__vbaExceptHandler"
'004957e6    6866724000              push 00407266
'004957eb    64a100000000            mov ax, word ptr fs:[00000000]
'004957f1    50                      push eax
    'Reference to '__except_list'
'004957f2    64892500000000          mov dword ptr fs:[00000000], esp
'004957f9    81ecb4000000            sub esp, 000000b4
'004957ff    53                      push ebx
'00495800    56                      push esi
'00495801    57                      push edi
'00495802    8965f8                  mov dword ptr [ebp-08], esp
'00495805    c745fc38154000          mov dword ptr [ebp-04], 00401538
'0049580c    8b5d0c                  mov ebx, dword ptr [ebp+0c]
'0049580f    8b03                    mov eax, dword ptr [ebx]
'00495811    33ff                    xor edi, edi
'00495813    50                      push eax
'00495814    68cc134300              push 004313cc
'00495819    897dec                  mov dword ptr [ebp-14], edi
'0049581c    897de8                  mov dword ptr [ebp-18], edi
'0049581f    897de4                  mov dword ptr [ebp-1c], edi
'00495822    897de0                  mov dword ptr [ebp-20], edi
'00495825    897ddc                  mov dword ptr [ebp-24], edi
'00495828    897dcc                  mov dword ptr [ebp-34], edi
'0049582b    897dbc                  mov dword ptr [ebp-44], edi
'0049582e    897dac                  mov dword ptr [ebp-54], edi
'00495831    897d9c                  mov dword ptr [ebp-64], edi
'00495834    897d8c                  mov dword ptr [ebp-74], edi
'00495837    89bd7cffffff            mov dword ptr [ebp+ffffff7c], edi
'0049583d    89bd6cffffff            mov dword ptr [ebp+ffffff6c], edi
'00495843    89bd58ffffff            mov dword ptr [ebp+ffffff58], edi

' *** Reference to "__vbaStrCmp"
'00495849    ff1530114000            call dword ptr [00401130]
'0049584f    85c0                    test ax, ax
'00495851    0f8499060000            je 495ef0

If (((arg_0) <> (vbNullChar))) Then
'00495857    393d24be7200            cmp dword ptr [0072be24], edi
'0049585d    7510                    jne 49586f
'0049585f    6824be7200              push 0072be24
'00495864    6870004300              push 00430070

' *** Reference to "__vbaNew2"
'00495869    ff152c124000            call dword ptr [0040122c]
    Dim var_2 As New Global
'0049586f    8b3524be7200            mov esi, dword ptr [0072be24]
'00495875    8b0e                    mov ecx, dword ptr [esi]
'00495877    8d55dc                  lea edx, dword ptr [ebp-24]
'0049587a    52                      push edx
'0049587b    56                      push esi
'0049587c    ff5114                  call dword ptr [ecx+14]
    Set var_39 = var_2.App
'0049587f    dbe2                    fnclex
'00495881    3bc7                    cmp eax, edi
'00495883    7d0f                    jge 495894
    
    If (    var_2.App < 0) Then
'00495885    6a14                    push 14
'00495887    6860004300              push 00430060
'0049588c    56                      push esi
'0049588d    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0049588e    ff1580104000            call dword ptr [00401080]
    
End If
'00495894    8b45dc                  mov eax, dword ptr [ebp-24]
'00495897    8b08                    mov ecx, dword ptr [eax]
'00495899    8d55ec                  lea edx, dword ptr [ebp-14]
'0049589c    52                      push edx
'0049589d    50                      push eax
'0049589e    8bf0                    mov esi, eax
'004958a0    ff5150                  call dword ptr [ecx+50]
var_75 = var_39.Path
'004958a3    dbe2                    fnclex
'004958a5    3bc7                    cmp eax, edi
'004958a7    7d0f                    jge 4958b8
'004958a9    6a50                    push 50
'004958ab    682c1c4300              push 00431c2c
'004958b0    56                      push esi
'004958b1    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'004958b2    ff1580104000            call dword ptr [00401080]
'004958b8    8b45ec                  mov eax, dword ptr [ebp-14]
'004958bb    50                      push eax
'004958bc    68641f4300              push 00431f64

' *** Reference to "__vbaStrCat"
'004958c1    ff1570104000            call dword ptr [00401070]
var_76 = (var_75) & ("\roles.mdb")

' *** Reference to "rtcUpperCaseVar"
'004958c7    8b352c114000            mov esi, dword ptr [0040112c]
'004958cd    8d4dcc                  lea ecx, dword ptr [ebp-34]
'004958d0    51                      push ecx
'004958d1    8d55bc                  lea edx, dword ptr [ebp-44]
'004958d4    52                      push edx
'004958d5    8945d4                  mov dword ptr [ebp-2c], eax
'004958d8    c745cc08000000          mov dword ptr [ebp-34], 00000008
'004958df    ffd6                    call esi
'004958e1    8d458c                  lea eax, dword ptr [ebp-74]
'004958e4    895d94                  mov dword ptr [ebp-6c], ebx
'004958e7    50                      push eax
'004958e8    8d4dac                  lea ecx, dword ptr [ebp-54]
'004958eb    bb08400000              mov ebx, 00004008
'004958f0    51                      push ecx
'004958f1    895d8c                  mov dword ptr [ebp-74], ebx
'004958f4    ffd6                    call esi
'004958f6    8d55bc                  lea edx, dword ptr [ebp-44]
'004958f9    52                      push edx
'004958fa    8d45ac                  lea eax, dword ptr [ebp-54]
'004958fd    50                      push eax

' *** Reference to "__vbaVarTstEq"
'004958fe    ff153c114000            call dword ptr [0040113c]
'00495904    8d4dec                  lea ecx, dword ptr [ebp-14]
'00495907    8bf0                    mov esi, eax

' *** Reference to "__vbaFreeStr"
'00495909    ff150c134000            call dword ptr [0040130c]
'#Cleanup(var_75)
'0049590f    8d4ddc                  lea ecx, dword ptr [ebp-24]

' *** Reference to "__vbaFreeObj"
'00495912    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_39)

' *** Reference to "__vbaFreeVarList"
'00495918    8b3d40104000            mov edi, dword ptr [00401040]
'0049591e    8d4dac                  lea ecx, dword ptr [ebp-54]
'00495921    51                      push ecx
'00495922    8d55bc                  lea edx, dword ptr [ebp-44]
'00495925    52                      push edx
'00495926    8d45cc                  lea eax, dword ptr [ebp-34]
'00495929    50                      push eax
'0049592a    6a03                    push 03
'0049592c    ffd7                    call edi
'#Cleanup( var_43, var_58, var_50)
'0049592e    83c410                  add esp, 10
'00495931    6685f6                  test esi, esi
'00495934    7475                    je 4959ab

If (((UCase(var_76)) = (UCase(arg_0)))) Then
'00495936    b904000280              mov ecx, 80020004
'0049593b    b80a000000              mov eax, 0000000a
'00495940    894db4                  mov dword ptr [ebp-4c], ecx
'00495943    894dc4                  mov dword ptr [ebp-3c], ecx
'00495946    8d558c                  lea edx, dword ptr [ebp-74]
'00495949    8d4dcc                  lea ecx, dword ptr [ebp-34]
'0049594c    8945ac                  mov dword ptr [ebp-54], eax
'0049594f    8945bc                  mov dword ptr [ebp-44], eax
'00495952    c7458424b07200          mov dword ptr [ebp-7c], 0072b024
'00495959    899d7cffffff            mov dword ptr [ebp+ffffff7c], ebx
'0049595f    c7459438234300          mov dword ptr [ebp-6c], 00432338
'00495966    c7458c08000000          mov dword ptr [ebp-74], 00000008

' *** Reference to "__vbaVarDup"
'0049596d    ff1598124000            call dword ptr [00401298]
    var_43 = ("On ne peut pas ouvrir la base du programme")
'00495973    8d4dac                  lea ecx, dword ptr [ebp-54]
'00495976    51                      push ecx
'00495977    8d55bc                  lea edx, dword ptr [ebp-44]
'0049597a    52                      push edx
'0049597b    8d857cffffff            lea eax, dword ptr [ebp+ffffff7c]
'00495981    50                      push eax
'00495982    6a30                    push 30
'00495984    8d4dcc                  lea ecx, dword ptr [ebp-34]
'00495987    51                      push ecx

' *** Reference to "rtcMsgBox"
'00495988    ff15b8104000            call dword ptr [004010b8]
    var_16 = MsgBox(var_43, 48, vbNullString)
'0049598e    8d55ac                  lea edx, dword ptr [ebp-54]
'00495991    52                      push edx
'00495992    8d45bc                  lea eax, dword ptr [ebp-44]
'00495995    50                      push eax
'00495996    8d4dcc                  lea ecx, dword ptr [ebp-34]
'00495999    51                      push ecx
'0049599a    6a03                    push 03
'0049599c    ffd7                    call edi
    '#Cleanup( var_43, var_58, var_50)
'0049599e    83c410                  add esp, 10
'004959a1    68385f4900              push 00495f38
'004959a6    e98c050000              jmp 495f37
    
End If
'004959ab    baf8064300              mov edx, 004306f8
'004959b0    8d4dec                  lea ecx, dword ptr [ebp-14]

' *** Reference to "__vbaStrCopy"
'004959b3    ff1548124000            call dword ptr [00401248]
var_75 = ("frmAcceder")
'004959b9    8b7d08                  mov edi, dword ptr [ebp+08]
'004959bc    8b17                    mov edx, dword ptr [edi]
'004959be    8d8558ffffff            lea eax, dword ptr [ebp+ffffff58]
'004959c4    50                      push eax
'004959c5    8d4dec                  lea ecx, dword ptr [ebp-14]
'004959c8    51                      push ecx
'004959c9    57                      push edi
'004959ca    ff92f8060000            call dword ptr [edx+000006f8]
'004959d0    85c0                    test ax, ax
'004959d2    7d16                    jge 4959ea

' *** Reference to "__vbaHresultCheckObj"
'004959d4    8b1d80104000            mov ebx, dword ptr [00401080]
'004959da    68f8060000              push 000006f8
'004959df    687cfd4200              push 0042fd7c
'004959e4    57                      push edi
'004959e5    50                      push eax
'004959e6    ffd3                    call ebx
'004959e8    eb06                    jmp 4959f0

' *** Reference to "__vbaHresultCheckObj"
'004959ea    8b1d80104000            mov ebx, dword ptr [00401080]
'004959f0    33d2                    xor edx, edx
var_num4 = Empty
'004959f2    6683bd58ffffffff        cmp word ptr [ebp+ffffff58], ffffffff
'004959fa    8d4dec                  lea ecx, dword ptr [ebp-14]
'004959fd    0f95c2                  setne dl
'00495a00    f7da                    neg edx
'00495a02    668bf2                  mov si, dx

' *** Reference to "__vbaFreeStr"
'00495a05    ff150c134000            call dword ptr [0040130c]
'#Cleanup(var_75)
'00495a0b    6685f6                  test esi, esi
'00495a0e    7477                    je 495a87

If (0 <> -1) Then
'00495a10    a124be7200              mov ax, word ptr [0072be24]
'00495a15    85c0                    test ax, ax
'00495a17    7510                    jne 495a29
'00495a19    6824be7200              push 0072be24
'00495a1e    6870004300              push 00430070

' *** Reference to "__vbaNew2"
'00495a23    ff152c124000            call dword ptr [0040122c]
    Set var_2 = New Global
'00495a29    a15cb07200              mov ax, word ptr [0072b05c]
'00495a2e    85c0                    test ax, ax
'00495a30    8b3524be7200            mov esi, dword ptr [0072be24]
'00495a36    7510                    jne 495a48
'00495a38    685cb07200              push 0072b05c
'00495a3d    68347b4000              push 00407b34

' *** Reference to "__vbaNew2"
'00495a42    ff152c124000            call dword ptr [0040122c]
    Dim var_24 As New frmAcceder
'00495a48    a15cb07200              mov ax, word ptr [0072b05c]
'00495a4d    8b1e                    mov ebx, dword ptr [esi]
'00495a4f    50                      push eax
'00495a50    8d4ddc                  lea ecx, dword ptr [ebp-24]
'00495a53    51                      push ecx

' *** Reference to "__vbaObjSetAddref"
'00495a54    ff15c8104000            call dword ptr [004010c8]
    Set var_39 = var_24
'00495a5a    50                      push eax
'00495a5b    56                      push esi
'00495a5c    ff5310                  call dword ptr [ebx+10]
    Call var_2.Unload(var_39)
'00495a5f    dbe2                    fnclex
'00495a61    85c0                    test ax, ax
'00495a63    7d13                    jge 495a78

' *** Reference to "__vbaHresultCheckObj"
'00495a65    8b1d80104000            mov ebx, dword ptr [00401080]
'00495a6b    6a10                    push 10
'00495a6d    6860004300              push 00430060
'00495a72    56                      push esi
'00495a73    50                      push eax
'00495a74    ffd3                    call ebx
'00495a76    eb06                    jmp 495a7e

' *** Reference to "__vbaHresultCheckObj"
'00495a78    8b1d80104000            mov ebx, dword ptr [00401080]
'00495a7e    8d4ddc                  lea ecx, dword ptr [ebp-24]

' *** Reference to "__vbaFreeObj"
'00495a81    ff1508134000            call dword ptr [00401308]
    '#Cleanup(var_39)
End If
'00495a87    ba94234300              mov edx, 00432394
'00495a8c    8d4dec                  lea ecx, dword ptr [ebp-14]

' *** Reference to "__vbaStrCopy"
'00495a8f    ff1548124000            call dword ptr [00401248]
var_75 = ("frmFichePerso")
'00495a95    8b17                    mov edx, dword ptr [edi]
'00495a97    8d8558ffffff            lea eax, dword ptr [ebp+ffffff58]
'00495a9d    50                      push eax
'00495a9e    8d4dec                  lea ecx, dword ptr [ebp-14]
'00495aa1    51                      push ecx
'00495aa2    57                      push edi
'00495aa3    ff92f8060000            call dword ptr [edx+000006f8]
'00495aa9    85c0                    test ax, ax
'00495aab    7d0e                    jge 495abb
'00495aad    68f8060000              push 000006f8
'00495ab2    687cfd4200              push 0042fd7c
'00495ab7    57                      push edi
'00495ab8    50                      push eax
'00495ab9    ffd3                    call ebx
'00495abb    33d2                    xor edx, edx
var_num4 = Empty
'00495abd    6683bd58ffffffff        cmp word ptr [ebp+ffffff58], ffffffff
'00495ac5    8d4dec                  lea ecx, dword ptr [ebp-14]
'00495ac8    0f95c2                  setne dl
'00495acb    f7da                    neg edx
'00495acd    668bf2                  mov si, dx

' *** Reference to "__vbaFreeStr"
'00495ad0    ff150c134000            call dword ptr [0040130c]
'#Cleanup(var_75)
'00495ad6    6685f6                  test esi, esi
'00495ad9    7477                    je 495b52

If (0 <> -1) Then
'00495adb    a124be7200              mov ax, word ptr [0072be24]
'00495ae0    85c0                    test ax, ax
'00495ae2    7510                    jne 495af4
'00495ae4    6824be7200              push 0072be24
'00495ae9    6870004300              push 00430070

' *** Reference to "__vbaNew2"
'00495aee    ff152c124000            call dword ptr [0040122c]
    Set var_2 = New Global
'00495af4    a174b17200              mov ax, word ptr [0072b174]
'00495af9    85c0                    test ax, ax
'00495afb    8b3524be7200            mov esi, dword ptr [0072be24]
'00495b01    7510                    jne 495b13
'00495b03    6874b17200              push 0072b174
'00495b08    6890c04100              push 0041c090

' *** Reference to "__vbaNew2"
'00495b0d    ff152c124000            call dword ptr [0040122c]
    Dim var_77 As New frmFichePerso
'00495b13    a174b17200              mov ax, word ptr [0072b174]
'00495b18    8b1e                    mov ebx, dword ptr [esi]
'00495b1a    50                      push eax
'00495b1b    8d4ddc                  lea ecx, dword ptr [ebp-24]
'00495b1e    51                      push ecx

' *** Reference to "__vbaObjSetAddref"
'00495b1f    ff15c8104000            call dword ptr [004010c8]
    Set var_39 = var_77
'00495b25    50                      push eax
'00495b26    56                      push esi
'00495b27    ff5310                  call dword ptr [ebx+10]
    Call var_2.Unload(var_39)
'00495b2a    dbe2                    fnclex
'00495b2c    85c0                    test ax, ax
'00495b2e    7d13                    jge 495b43

' *** Reference to "__vbaHresultCheckObj"
'00495b30    8b1d80104000            mov ebx, dword ptr [00401080]
'00495b36    6a10                    push 10
'00495b38    6860004300              push 00430060
'00495b3d    56                      push esi
'00495b3e    50                      push eax
'00495b3f    ffd3                    call ebx
'00495b41    eb06                    jmp 495b49

' *** Reference to "__vbaHresultCheckObj"
'00495b43    8b1d80104000            mov ebx, dword ptr [00401080]
'00495b49    8d4ddc                  lea ecx, dword ptr [ebp-24]

' *** Reference to "__vbaFreeObj"
'00495b4c    ff1508134000            call dword ptr [00401308]
    '#Cleanup(var_39)
End If
'00495b52    baac2e4300              mov edx, 00432eac
'00495b57    8d4dec                  lea ecx, dword ptr [ebp-14]

' *** Reference to "__vbaStrCopy"
'00495b5a    ff1548124000            call dword ptr [00401248]
var_75 = ("frmGestPerso")
'00495b60    8b17                    mov edx, dword ptr [edi]
'00495b62    8d8558ffffff            lea eax, dword ptr [ebp+ffffff58]
'00495b68    50                      push eax
'00495b69    8d4dec                  lea ecx, dword ptr [ebp-14]
'00495b6c    51                      push ecx
'00495b6d    57                      push edi
'00495b6e    ff92f8060000            call dword ptr [edx+000006f8]
'00495b74    85c0                    test ax, ax
'00495b76    7d0e                    jge 495b86
'00495b78    68f8060000              push 000006f8
'00495b7d    687cfd4200              push 0042fd7c
'00495b82    57                      push edi
'00495b83    50                      push eax
'00495b84    ffd3                    call ebx
'00495b86    33d2                    xor edx, edx
var_num4 = Empty
'00495b88    6683bd58ffffffff        cmp word ptr [ebp+ffffff58], ffffffff
'00495b90    8d4dec                  lea ecx, dword ptr [ebp-14]
'00495b93    0f95c2                  setne dl
'00495b96    f7da                    neg edx
'00495b98    668bf2                  mov si, dx

' *** Reference to "__vbaFreeStr"
'00495b9b    ff150c134000            call dword ptr [0040130c]
'#Cleanup(var_75)
'00495ba1    6685f6                  test esi, esi
'00495ba4    7469                    je 495c0f

If (0 <> -1) Then
'00495ba6    a124be7200              mov ax, word ptr [0072be24]
'00495bab    85c0                    test ax, ax
'00495bad    7510                    jne 495bbf
'00495baf    6824be7200              push 0072be24
'00495bb4    6870004300              push 00430070

' *** Reference to "__vbaNew2"
'00495bb9    ff152c124000            call dword ptr [0040122c]
    Set var_2 = New Global
'00495bbf    a184b07200              mov ax, word ptr [0072b084]
'00495bc4    85c0                    test ax, ax
'00495bc6    8b3524be7200            mov esi, dword ptr [0072be24]
'00495bcc    7510                    jne 495bde
'00495bce    6884b07200              push 0072b084
'00495bd3    68f8894100              push 004189f8

' *** Reference to "__vbaNew2"
'00495bd8    ff152c124000            call dword ptr [0040122c]
    Dim var_28 As New frmGestPerso
'00495bde    a184b07200              mov ax, word ptr [0072b084]
'00495be3    8b3e                    mov edi, dword ptr [esi]
'00495be5    50                      push eax
'00495be6    8d4ddc                  lea ecx, dword ptr [ebp-24]
'00495be9    51                      push ecx

' *** Reference to "__vbaObjSetAddref"
'00495bea    ff15c8104000            call dword ptr [004010c8]
    Set var_39 = var_28
'00495bf0    50                      push eax
'00495bf1    56                      push esi
'00495bf2    ff5710                  call dword ptr [edi+10]
    Call var_2.Unload(var_39)
'00495bf5    dbe2                    fnclex
'00495bf7    85c0                    test ax, ax
'00495bf9    7d0b                    jge 495c06
'00495bfb    6a10                    push 10
'00495bfd    6860004300              push 00430060
'00495c02    56                      push esi
'00495c03    50                      push eax
'00495c04    ffd3                    call ebx
'00495c06    8d4ddc                  lea ecx, dword ptr [ebp-24]

' *** Reference to "__vbaFreeObj"
'00495c09    ff1508134000            call dword ptr [00401308]
    '#Cleanup(var_39)
End If
'00495c0f    a148b07200              mov ax, word ptr [0072b048]
'00495c14    8b10                    mov edx, dword ptr [eax]
'00495c16    50                      push eax
'00495c17    ff5258                  call dword ptr [edx+58]
'00495c1a    dbe2                    fnclex
'00495c1c    85c0                    test ax, ax
'00495c1e    7d11                    jge 495c31
'00495c20    8b0d48b07200            mov ecx, dword ptr [0072b048]
'00495c26    6a58                    push 58
'00495c28    68301f4300              push 00431f30
'00495c2d    51                      push ecx
'00495c2e    50                      push eax
'00495c2f    ffd3                    call ebx
'00495c31    a128be7200              mov ax, word ptr [0072be28]
'00495c36    85c0                    test ax, ax
'00495c38    7510                    jne 495c4a
'00495c3a    6828be7200              push 0072be28
'00495c3f    6830214300              push 00432130

' *** Reference to "__vbaNew2"
'00495c44    ff152c124000            call dword ptr [0040122c]
Dim var_57 As New DBEngine
'00495c4a    8d7ddc                  lea edi, dword ptr [ebp-24]
'00495c4d    57                      push edi
'00495c4e    83ec10                  sub esp, 10
'00495c51    8bfc                    mov edi, esp
'00495c53    b90a000000              mov ecx, 0000000a
'00495c58    890f                    mov dword ptr [edi], ecx
'00495c5a    898d7cffffff            mov dword ptr [ebp+ffffff7c], ecx
'00495c60    894d8c                  mov dword ptr [ebp-74], ecx
'00495c63    8b8d70ffffff            mov ecx, dword ptr [ebp+ffffff70]
'00495c69    894f04                  mov dword ptr [edi+04], ecx
'00495c6c    8b3528be7200            mov esi, dword ptr [0072be28]
'00495c72    ba04000280              mov edx, 80020004
'00495c77    8bc2                    mov eax, edx
'00495c79    894708                  mov dword ptr [edi+08], eax
'00495c7c    8b8578ffffff            mov eax, dword ptr [ebp+ffffff78]
'00495c82    89470c                  mov dword ptr [edi+0c], eax
'00495c85    8b857cffffff            mov eax, dword ptr [ebp+ffffff7c]
'00495c8b    8b7d0c                  mov edi, dword ptr [ebp+0c]
'00495c8e    83ec10                  sub esp, 10
'00495c91    8bcc                    mov ecx, esp
'00495c93    8901                    mov dword ptr [ecx], eax
'00495c95    8b4580                  mov eax, dword ptr [ebp-80]
'00495c98    894104                  mov dword ptr [ecx+04], eax
'00495c9b    895108                  mov dword ptr [ecx+08], edx
'00495c9e    895594                  mov dword ptr [ebp-6c], edx
'00495ca1    895584                  mov dword ptr [ebp-7c], edx
'00495ca4    8b5588                  mov edx, dword ptr [ebp-78]
'00495ca7    8b1e                    mov ebx, dword ptr [esi]
'00495ca9    89510c                  mov dword ptr [ecx+0c], edx
'00495cac    8b4d8c                  mov ecx, dword ptr [ebp-74]
'00495caf    8b5590                  mov edx, dword ptr [ebp-70]
'00495cb2    83ec10                  sub esp, 10
'00495cb5    8bc4                    mov eax, esp
'00495cb7    8908                    mov dword ptr [eax], ecx
'00495cb9    8b4d94                  mov ecx, dword ptr [ebp-6c]
'00495cbc    895004                  mov dword ptr [eax+04], edx
'00495cbf    8b5598                  mov edx, dword ptr [ebp-68]
'00495cc2    894808                  mov dword ptr [eax+08], ecx
'00495cc5    89500c                  mov dword ptr [eax+0c], edx
'00495cc8    8b07                    mov eax, dword ptr [edi]
'00495cca    50                      push eax
'00495ccb    56                      push esi
'00495ccc    ff5358                  call dword ptr [ebx+58]
Set var_39 = var_57.OpenDatabase(arg_0)
'00495ccf    dbe2                    fnclex
'00495cd1    85c0                    test ax, ax
'00495cd3    7d0f                    jge 495ce4
'00495cd5    6a58                    push 58
'00495cd7    6820214300              push 00432120
'00495cdc    56                      push esi
'00495cdd    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00495cde    ff1580104000            call dword ptr [00401080]
'00495ce4    8b45dc                  mov eax, dword ptr [ebp-24]
'00495ce7    50                      push eax
'00495ce8    6848b07200              push 0072b048
'00495ced    c745dc00000000          mov dword ptr [ebp-24], 00000000

' *** Reference to "__vbaObjSet"
'00495cf4    ff15b4104000            call dword ptr [004010b4]
Set var_78 = var_39
'00495cfa    6848b07200              push 0072b048
'00495cff    e8cc5f0400              call 4dbcd0
Call sub_4DBCD0()
'00495d04    a110b07200              mov ax, word ptr [0072b010]
'00495d09    85c0                    test ax, ax
'00495d0b    7510                    jne 495d1d
'00495d0d    6810b07200              push 0072b010
'00495d12    68b0e54000              push 0040e5b0

' *** Reference to "__vbaNew2"
'00495d17    ff152c124000            call dword ptr [0040122c]
Dim var_60 As New frmMain
'00495d1d    8b0f                    mov ecx, dword ptr [edi]
'00495d1f    8b3510b07200            mov esi, dword ptr [0072b010]
'00495d25    8b1e                    mov ebx, dword ptr [esi]
'00495d27    68a8214300              push 004321a8
'00495d2c    51                      push ecx

' *** Reference to "__vbaStrCat"
'00495d2d    ff1570104000            call dword ptr [00401070]
var_49 = ("RoleDD Gestion de personnages pour Donjons et Dragons 3.5 [") & (arg_0)

' *** Reference to "__vbaStrMove"
'00495d33    8b3dd0124000            mov edi, dword ptr [004012d0]
'00495d39    8bd0                    mov edx, eax
'00495d3b    8d4dec                  lea ecx, dword ptr [ebp-14]
'00495d3e    ffd7                    call edi
'00495d40    50                      push eax
'00495d41    6824224300              push 00432224

' *** Reference to "__vbaStrCat"
'00495d46    ff1570104000            call dword ptr [00401070]
var_79 = (var_49) & ("]")
'00495d4c    8bd0                    mov edx, eax
'00495d4e    8d4de8                  lea ecx, dword ptr [ebp-18]
'00495d51    ffd7                    call edi
'00495d53    50                      push eax
'00495d54    56                      push esi
'00495d55    ff5354                  call dword ptr [ebx+54]
'00495d58    dbe2                    fnclex
'00495d5a    85c0                    test ax, ax
'00495d5c    7d0f                    jge 495d6d
'00495d5e    6a54                    push 54
'00495d60    684cfd4200              push 0042fd4c
'00495d65    56                      push esi
'00495d66    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00495d67    ff1580104000            call dword ptr [00401080]
'00495d6d    8d55e8                  lea edx, dword ptr [ebp-18]
'00495d70    52                      push edx
'00495d71    8d45ec                  lea eax, dword ptr [ebp-14]
'00495d74    50                      push eax
'00495d75    6a02                    push 02

' *** Reference to "__vbaFreeStrList"
'00495d77    ff155c124000            call dword ptr [0040125c]
'#Cleanup( -4624, -4628)
'00495d7d    a124be7200              mov ax, word ptr [0072be24]
'00495d82    83c40c                  add esp, 0c
'00495d85    85c0                    test ax, ax
'00495d87    7514                    jne 495d9d

' *** Reference to "__vbaNew2"
'00495d89    8b3d2c124000            mov edi, dword ptr [0040122c]
'00495d8f    6824be7200              push 0072be24
'00495d94    6870004300              push 00430070
'00495d99    ffd7                    call edi
Set var_2 = New Global
'00495d9b    eb06                    jmp 495da3

' *** Reference to "__vbaNew2"
'00495d9d    8b3d2c124000            mov edi, dword ptr [0040122c]
'00495da3    a15cb07200              mov ax, word ptr [0072b05c]
'00495da8    85c0                    test ax, ax
'00495daa    8b3524be7200            mov esi, dword ptr [0072be24]
'00495db0    750c                    jne 495dbe
'00495db2    685cb07200              push 0072b05c
'00495db7    68347b4000              push 00407b34
'00495dbc    ffd7                    call edi
Set var_24 = New frmAcceder
'00495dbe    8b0d5cb07200            mov ecx, dword ptr [0072b05c]
'00495dc4    8b1e                    mov ebx, dword ptr [esi]
'00495dc6    51                      push ecx
'00495dc7    8d55dc                  lea edx, dword ptr [ebp-24]
'00495dca    52                      push edx

' *** Reference to "__vbaObjSetAddref"
'00495dcb    ff15c8104000            call dword ptr [004010c8]
Set var_39 = var_24
'00495dd1    50                      push eax
'00495dd2    56                      push esi
'00495dd3    ff5310                  call dword ptr [ebx+10]
Call var_2.Unload(var_39)
'00495dd6    dbe2                    fnclex
'00495dd8    85c0                    test ax, ax
'00495dda    7d0f                    jge 495deb
'00495ddc    6a10                    push 10
'00495dde    6860004300              push 00430060
'00495de3    56                      push esi
'00495de4    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00495de5    ff1580104000            call dword ptr [00401080]
'00495deb    8d4ddc                  lea ecx, dword ptr [ebp-24]

' *** Reference to "__vbaFreeObj"
'00495dee    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_39)
'00495df4    a15cb07200              mov ax, word ptr [0072b05c]
'00495df9    85c0                    test ax, ax
'00495dfb    750c                    jne 495e09
'00495dfd    685cb07200              push 0072b05c
'00495e02    68347b4000              push 00407b34
'00495e07    ffd7                    call edi
Set var_24 = New frmAcceder
'00495e09    8b355cb07200            mov esi, dword ptr [0072b05c]
'00495e0f    83ec10                  sub esp, 10
'00495e12    b90a000000              mov ecx, 0000000a
'00495e17    894d8c                  mov dword ptr [ebp-74], ecx
'00495e1a    898d7cffffff            mov dword ptr [ebp+ffffff7c], ecx
'00495e20    8bdc                    mov ebx, esp
'00495e22    890b                    mov dword ptr [ebx], ecx
'00495e24    8b4d80                  mov ecx, dword ptr [ebp-80]
'00495e27    894b04                  mov dword ptr [ebx+04], ecx
'00495e2a    b804000280              mov eax, 80020004
'00495e2f    894308                  mov dword ptr [ebx+08], eax
'00495e32    8bd0                    mov edx, eax
'00495e34    894584                  mov dword ptr [ebp-7c], eax
'00495e37    8b4588                  mov eax, dword ptr [ebp-78]
'00495e3a    89430c                  mov dword ptr [ebx+0c], eax
'00495e3d    8b458c                  mov eax, dword ptr [ebp-74]
'00495e40    83ec10                  sub esp, 10
'00495e43    8bcc                    mov ecx, esp
'00495e45    8901                    mov dword ptr [ecx], eax
'00495e47    8b4590                  mov eax, dword ptr [ebp-70]
'00495e4a    894104                  mov dword ptr [ecx+04], eax
'00495e4d    895594                  mov dword ptr [ebp-6c], edx
'00495e50    8b3e                    mov edi, dword ptr [esi]
'00495e52    895108                  mov dword ptr [ecx+08], edx
'00495e55    8b5598                  mov edx, dword ptr [ebp-68]
'00495e58    56                      push esi
'00495e59    89510c                  mov dword ptr [ecx+0c], edx
'00495e5c    ff97b0020000            call dword ptr [edi+000002b0]
'00495e62    dbe2                    fnclex
'00495e64    85c0                    test ax, ax
'00495e66    7d12                    jge 495e7a
'00495e68    68b0020000              push 000002b0
'00495e6d    6810074300              push 00430710
'00495e72    56                      push esi
'00495e73    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00495e74    ff1580104000            call dword ptr [00401080]
'00495e7a    a134b07200              mov ax, word ptr [0072b034]

' *** Reference to "__vbaStrToAnsi"
'00495e7f    8b359c124000            mov esi, dword ptr [0040129c]
'00495e85    50                      push eax
'00495e86    8d4de0                  lea ecx, dword ptr [ebp-20]
'00495e89    51                      push ecx
'00495e8a    ffd6                    call esi
var_3 = (vbNullString)
'00495e8c    8b7d0c                  mov edi, dword ptr [ebp+0c]
'00495e8f    8b17                    mov edx, dword ptr [edi]
'00495e91    50                      push eax
'00495e92    52                      push edx
'00495e93    8d45e4                  lea eax, dword ptr [ebp-1c]
'00495e96    50                      push eax
'00495e97    ffd6                    call esi
var_40 = (arg_0)
'00495e99    50                      push eax
'00495e9a    68e02e4300              push 00432ee0
'00495e9f    8d4de8                  lea ecx, dword ptr [ebp-18]
'00495ea2    51                      push ecx
'00495ea3    ffd6                    call esi
var_41 = ("BDD")
'00495ea5    50                      push eax
'00495ea6    68cc2e4300              push 00432ecc
'00495eab    8d55ec                  lea edx, dword ptr [ebp-14]
'00495eae    52                      push edx
'00495eaf    ffd6                    call esi
var_75 = ("FICHIER")
'00495eb1    50                      push eax
'00495eb2    e8f9a4f9ff              call 4303b0
var_80 = WritePrivateProfileStringA (var_75, var_41, var_40, var_3)  '{Function}

' *** Reference to "__vbaSetSystemError"
'00495eb7    ff157c104000            call dword ptr [0040107c]
'#SetAPISystemError()
'00495ebd    8b45e4                  mov eax, dword ptr [ebp-1c]

' *** Reference to "__vbaStrToUnicode"
'00495ec0    8b35bc114000            mov esi, dword ptr [004011bc]
'00495ec6    50                      push eax
'00495ec7    57                      push edi
'00495ec8    ffd6                    call esi
var_81 = (arg_0)
'00495eca    8b4de0                  mov ecx, dword ptr [ebp-20]
'00495ecd    51                      push ecx
'00495ece    6834b07200              push 0072b034
'00495ed3    ffd6                    call esi
var_82 = (vbNullString)
'00495ed5    8d55e0                  lea edx, dword ptr [ebp-20]
'00495ed8    52                      push edx
'00495ed9    8d45e4                  lea eax, dword ptr [ebp-1c]
'00495edc    50                      push eax
'00495edd    8d4de8                  lea ecx, dword ptr [ebp-18]
'00495ee0    51                      push ecx
'00495ee1    8d55ec                  lea edx, dword ptr [ebp-14]
'00495ee4    52                      push edx
'00495ee5    6a04                    push 04

' *** Reference to "__vbaFreeStrList"
'00495ee7    ff155c124000            call dword ptr [0040125c]
'#Cleanup( var_75, var_41, var_40, var_3)
'00495eed    83c414                  add esp, 14

'ERROR: Two many next close:
End If
'00495ef0    68385f4900              push 00495f38
'00495ef5    eb40                    jmp 495f37
'00495ef7    8d45e0                  lea eax, dword ptr [ebp-20]
'00495efa    50                      push eax
'00495efb    8d4de4                  lea ecx, dword ptr [ebp-1c]
'00495efe    51                      push ecx
'00495eff    8d55e8                  lea edx, dword ptr [ebp-18]
'00495f02    52                      push edx
'00495f03    8d45ec                  lea eax, dword ptr [ebp-14]
'00495f06    50                      push eax
'00495f07    6a04                    push 04

' *** Reference to "__vbaFreeStrList"
'00495f09    ff155c124000            call dword ptr [0040125c]
'#Cleanup( var_75, var_41, var_40, var_3)
'00495f0f    83c414                  add esp, 14
'00495f12    8d4ddc                  lea ecx, dword ptr [ebp-24]

' *** Reference to "__vbaFreeObj"
'00495f15    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_39)
'00495f1b    8d4d9c                  lea ecx, dword ptr [ebp-64]
'00495f1e    51                      push ecx
'00495f1f    8d55ac                  lea edx, dword ptr [ebp-54]
'00495f22    52                      push edx
'00495f23    8d45bc                  lea eax, dword ptr [ebp-44]
'00495f26    50                      push eax
'00495f27    8d4dcc                  lea ecx, dword ptr [ebp-34]
'00495f2a    51                      push ecx
'00495f2b    6a04                    push 04

' *** Reference to "__vbaFreeVarList"
'00495f2d    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_43, var_58, var_50, var_51)
'00495f33    83c414                  add esp, 14
'00495f36    c3                      ret
'00495f37    c3                      ret
'00495f38    8b4df0                  mov ecx, dword ptr [ebp-10]
'00495f3b    5f                      pop edi
'00495f3c    5e                      pop esi
'00495f3d    33c0                    xor eax, eax
    'Reference to '__except_list'
'00495f3f    64890d00000000          mov dword ptr fs:[00000000], ecx
'00495f46    5b                      pop ebx
'00495f47    8be5                    mov esp, ebp
'00495f49    5d                      pop ebp
'00495f4a    c20800                  ret 0008


End Function


'Event for IDM_MISES_JOUR
Private Sub IDM_MISES_JOUR_Click()
'00493020    55                      push ebp
'00493021    8bec                    mov ebp, esp
'00493023    83ec0c                  sub esp, 0c

' *** Reference to "__vbaExceptHandler"
'00493026    6866724000              push 00407266
'0049302b    64a100000000            mov ax, word ptr fs:[00000000]
'00493031    50                      push eax
    'Reference to '__except_list'
'00493032    64892500000000          mov dword ptr fs:[00000000], esp
'00493039    83ec40                  sub esp, 40
'0049303c    53                      push ebx
'0049303d    56                      push esi
'0049303e    57                      push edi
'0049303f    8965f4                  mov dword ptr [ebp-0c], esp
'00493042    c745f8d8134000          mov dword ptr [ebp-08], 004013d8
'00493049    8b7508                  mov esi, dword ptr [ebp+08]
'0049304c    8bc6                    mov eax, esi
'0049304e    83e001                  and eax, 01
'00493051    8945fc                  mov dword ptr [ebp-04], eax
'00493054    83e6fe                  and esi, fffffffe
'00493057    8b0e                    mov ecx, dword ptr [esi]
'00493059    56                      push esi
'0049305a    897508                  mov dword ptr [ebp+08], esi
'0049305d    ff5104                  call dword ptr [ecx+04]
'00493060    8b16                    mov edx, dword ptr [esi]
'00493062    33ff                    xor edi, edi
'00493064    8d45c8                  lea eax, dword ptr [ebp-38]
'00493067    50                      push eax
'00493068    56                      push esi
'00493069    897de8                  mov dword ptr [ebp-18], edi
'0049306c    897de4                  mov dword ptr [ebp-1c], edi
'0049306f    897de0                  mov dword ptr [ebp-20], edi
'00493072    897ddc                  mov dword ptr [ebp-24], edi
'00493075    897dd8                  mov dword ptr [ebp-28], edi
'00493078    897dd4                  mov dword ptr [ebp-2c], edi
'0049307b    897dd0                  mov dword ptr [ebp-30], edi
'0049307e    897dcc                  mov dword ptr [ebp-34], edi
'00493081    897dc8                  mov dword ptr [ebp-38], edi
'00493084    ff5258                  call dword ptr [edx+58]
'00493087    dbe2                    fnclex
'00493089    3bc7                    cmp eax, edi
'0049308b    7d13                    jge 4930a0

' *** Reference to "__vbaHresultCheckObj"
'0049308d    8b1d80104000            mov ebx, dword ptr [00401080]
'00493093    6a58                    push 58
'00493095    684cfd4200              push 0042fd4c
'0049309a    56                      push esi
'0049309b    50                      push eax
'0049309c    ffd3                    call ebx
'0049309e    eb06                    jmp 4930a6

' *** Reference to "__vbaHresultCheckObj"
'004930a0    8b1d80104000            mov ebx, dword ptr [00401080]
'004930a6    393d24be7200            cmp dword ptr [0072be24], edi
'004930ac    7510                    jne 4930be
'004930ae    6824be7200              push 0072be24
'004930b3    6870004300              push 00430070

' *** Reference to "__vbaNew2"
'004930b8    ff152c124000            call dword ptr [0040122c]
Dim var_2 As New Global
'004930be    8b3524be7200            mov esi, dword ptr [0072be24]
'004930c4    8b0e                    mov ecx, dword ptr [esi]
'004930c6    8d55d0                  lea edx, dword ptr [ebp-30]
'004930c9    52                      push edx
'004930ca    56                      push esi
'004930cb    ff5114                  call dword ptr [ecx+14]
Set var_4 = var_2.App
'004930ce    dbe2                    fnclex
'004930d0    3bc7                    cmp eax, edi
'004930d2    7d0b                    jge 4930df

If (var_2.App < 0) Then
'004930d4    6a14                    push 14
'004930d6    6860004300              push 00430060
'004930db    56                      push esi
'004930dc    50                      push eax
'004930dd    ffd3                    call ebx
    
End If
'004930df    8b45d0                  mov eax, dword ptr [ebp-30]
'004930e2    8b08                    mov ecx, dword ptr [eax]
'004930e4    8d55cc                  lea edx, dword ptr [ebp-34]
'004930e7    52                      push edx
'004930e8    50                      push eax
'004930e9    8bf0                    mov esi, eax
'004930eb    ff91c8000000            call dword ptr [ecx+000000c8]
var_43 = var_4.Revision
'004930f1    dbe2                    fnclex
'004930f3    3bc7                    cmp eax, edi
'004930f5    7d0e                    jge 493105
'004930f7    68c8000000              push 000000c8
'004930fc    682c1c4300              push 00431c2c
'00493101    56                      push esi
'00493102    50                      push eax
'00493103    ffd3                    call ebx

' *** Reference to "__vbaStrToAnsi"
'00493105    8b359c124000            mov esi, dword ptr [0040129c]
'0049310b    57                      push edi
'0049310c    68cc134300              push 004313cc
'00493111    8d45d4                  lea eax, dword ptr [ebp-2c]
'00493114    50                      push eax
'00493115    ffd6                    call esi
var_44 = (vbNullChar)
'00493117    50                      push eax
'00493118    68cc134300              push 004313cc
'0049311d    8d4dd8                  lea ecx, dword ptr [ebp-28]
'00493120    51                      push ecx
'00493121    ffd6                    call esi
var_45 = (vbNullChar)
'00493123    8b55cc                  mov edx, dword ptr [ebp-34]
'00493126    50                      push eax
'00493127    68bc1b4300              push 00431bbc
'0049312c    52                      push edx

' *** Reference to "__vbaStrI2"
'0049312d    ff1510104000            call dword ptr [00401010]

' *** Reference to "__vbaStrMove"
'00493133    8b1dd0124000            mov ebx, dword ptr [004012d0]
'00493139    8bd0                    mov edx, eax
'0049313b    8d4de8                  lea ecx, dword ptr [ebp-18]
'0049313e    ffd3                    call ebx
'00493140    50                      push eax

' *** Reference to "__vbaStrCat"
'00493141    ff1570104000            call dword ptr [00401070]
var_12 = ("http://bonnarien.dyndns.org/programme/test.php?version=") & (CStr(var_43))
'00493147    8bd0                    mov edx, eax
'00493149    8d4de0                  lea ecx, dword ptr [ebp-20]
'0049314c    ffd3                    call ebx
'0049314e    50                      push eax
'0049314f    8d45dc                  lea eax, dword ptr [ebp-24]
'00493152    50                      push eax
'00493153    ffd6                    call esi
var_39 = (var_12)
'00493155    50                      push eax
'00493156    68701b4300              push 00431b70
'0049315b    8d4de4                  lea ecx, dword ptr [ebp-1c]
'0049315e    51                      push ecx
'0049315f    ffd6                    call esi
var_40 = ("open")
'00493161    8b55c8                  mov edx, dword ptr [ebp-38]
'00493164    50                      push eax
'00493165    52                      push edx
'00493166    e80dd4f9ff              call 430578
var_42 = ShellExecuteA (0, var_40, var_39, var_45, var_44, 0)  '{Function}

' *** Reference to "__vbaSetSystemError"
'0049316b    ff157c104000            call dword ptr [0040107c]
'#SetAPISystemError()
'00493171    8d45d4                  lea eax, dword ptr [ebp-2c]
'00493174    50                      push eax
'00493175    8d4dd8                  lea ecx, dword ptr [ebp-28]
'00493178    51                      push ecx
'00493179    8d55dc                  lea edx, dword ptr [ebp-24]
'0049317c    52                      push edx
'0049317d    8d45e0                  lea eax, dword ptr [ebp-20]
'00493180    50                      push eax
'00493181    8d4de4                  lea ecx, dword ptr [ebp-1c]
'00493184    51                      push ecx
'00493185    8d55e8                  lea edx, dword ptr [ebp-18]
'00493188    52                      push edx
'00493189    6a06                    push 06

' *** Reference to "__vbaFreeStrList"
'0049318b    ff155c124000            call dword ptr [0040125c]
'#Cleanup( -4512, var_40, -4516, var_39, var_45, var_44)
'00493191    83c41c                  add esp, 1c
'00493194    8d4dd0                  lea ecx, dword ptr [ebp-30]

' *** Reference to "__vbaFreeObj"
'00493197    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_4)
'0049319d    897dfc                  mov dword ptr [ebp-04], edi
'004931a0    68d5314900              push 004931d5
'004931a5    eb2d                    jmp 4931d4
'004931a7    8d45d4                  lea eax, dword ptr [ebp-2c]
'004931aa    50                      push eax
'004931ab    8d4dd8                  lea ecx, dword ptr [ebp-28]
'004931ae    51                      push ecx
'004931af    8d55dc                  lea edx, dword ptr [ebp-24]
'004931b2    52                      push edx
'004931b3    8d45e0                  lea eax, dword ptr [ebp-20]
'004931b6    50                      push eax
'004931b7    8d4de4                  lea ecx, dword ptr [ebp-1c]
'004931ba    51                      push ecx
'004931bb    8d55e8                  lea edx, dword ptr [ebp-18]
'004931be    52                      push edx
'004931bf    6a06                    push 06

' *** Reference to "__vbaFreeStrList"
'004931c1    ff155c124000            call dword ptr [0040125c]
'#Cleanup( -4512, var_40, -4516, var_39, var_45, var_44)
'004931c7    83c41c                  add esp, 1c
'004931ca    8d4dd0                  lea ecx, dword ptr [ebp-30]

' *** Reference to "__vbaFreeObj"
'004931cd    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_4)
'004931d3    c3                      ret
'004931d4    c3                      ret
'004931d5    8b4508                  mov eax, dword ptr [ebp+08]
'004931d8    8b08                    mov ecx, dword ptr [eax]
'004931da    50                      push eax
'004931db    ff5108                  call dword ptr [ecx+08]
'004931de    8b45fc                  mov eax, dword ptr [ebp-04]
'004931e1    8b4dec                  mov ecx, dword ptr [ebp-14]
'004931e4    5f                      pop edi
'004931e5    5e                      pop esi
    'Reference to '__except_list'
'004931e6    64890d00000000          mov dword ptr fs:[00000000], ecx
'004931ed    5b                      pop ebx
'004931ee    8be5                    mov esp, ebp
'004931f0    5d                      pop ebp
'004931f1    c20400                  ret 0004


End Sub


'Event for IMD_IMPORT
Private Sub IMD_IMPORT_Click()
'00493ac0    55                      push ebp
'00493ac1    8bec                    mov ebp, esp
'00493ac3    83ec18                  sub esp, 18

' *** Reference to "__vbaExceptHandler"
'00493ac6    6866724000              push 00407266
'00493acb    64a100000000            mov ax, word ptr fs:[00000000]
'00493ad1    50                      push eax
    'Reference to '__except_list'
'00493ad2    64892500000000          mov dword ptr fs:[00000000], esp
'00493ad9    b828010000              mov eax, 00000128

' *** Reference to "__vbaChkstk"
'00493ade    e87d37f7ff              call 407260
'00493ae3    53                      push ebx
'00493ae4    56                      push esi
'00493ae5    57                      push edi
'00493ae6    8965e8                  mov dword ptr [ebp-18], esp
'00493ae9    c745ec58144000          mov dword ptr [ebp-14], 00401458
'00493af0    8b4508                  mov eax, dword ptr [ebp+08]
'00493af3    83e001                  and eax, 01
'00493af6    8945f0                  mov dword ptr [ebp-10], eax
'00493af9    8b4d08                  mov ecx, dword ptr [ebp+08]
'00493afc    83e1fe                  and ecx, fffffffe
'00493aff    894d08                  mov dword ptr [ebp+08], ecx
'00493b02    c745f400000000          mov dword ptr [ebp-0c], 00000000
'00493b09    8b5508                  mov edx, dword ptr [ebp+08]
'00493b0c    8b02                    mov eax, dword ptr [edx]
'00493b0e    8b4d08                  mov ecx, dword ptr [ebp+08]
'00493b11    51                      push ecx
'00493b12    ff5004                  call dword ptr [eax+04]
'00493b15    c745fc01000000          mov dword ptr [ebp-04], 00000001
'00493b1c    c745fc02000000          mov dword ptr [ebp-04], 00000002
'00493b23    0fbf1528b07200          movsx edx, word ptr [0072b028]
'00493b2a    85d2                    test dx, dx
'00493b2c    750f                    jne 493b3d
'00493b2e    c745fc03000000          mov dword ptr [ebp-04], 00000003
'00493b35    6aff                    push ffffffff

' *** Reference to "__vbaOnError"
'00493b37    ff15b0104000            call dword ptr [004010b0]
On Error Resume Next
'00493b3d    c745fc05000000          mov dword ptr [ebp-04], 00000005
'00493b44    8d8538ffffff            lea eax, dword ptr [ebp+ffffff38]
'00493b4a    50                      push eax
'00493b4b    8b4d08                  mov ecx, dword ptr [ebp+08]
'00493b4e    8b11                    mov edx, dword ptr [ecx]
'00493b50    8b4508                  mov eax, dword ptr [ebp+08]
'00493b53    50                      push eax
'00493b54    ff5258                  call dword ptr [edx+58]
'00493b57    dbe2                    fnclex
'00493b59    898528ffffff            mov dword ptr [ebp+ffffff28], eax
'00493b5f    83bd28ffffff00          cmp dword ptr [ebp+ffffff28], 00
'00493b66    7d20                    jge 493b88

If (Me < 0) Then
'00493b68    6a58                    push 58
'00493b6a    684cfd4200              push 0042fd4c
'00493b6f    8b4d08                  mov ecx, dword ptr [ebp+08]
'00493b72    51                      push ecx
'00493b73    8b9528ffffff            mov edx, dword ptr [ebp+ffffff28]
'00493b79    52                      push edx

' *** Reference to "__vbaHresultCheckObj"
'00493b7a    ff1580104000            call dword ptr [00401080]
'00493b80    8985f8feffff            mov dword ptr [ebp+fffffef8], eax
'00493b86    eb0a                    jmp 493b92
    
End If
'00493b88    c785f8feffff00000000    mov dword ptr [ebp+fffffef8], 00000000
'00493b92    833d24be720000          cmp dword ptr [0072be24], 00
'00493b99    751c                    jne 493bb7
'00493b9b    6824be7200              push 0072be24
'00493ba0    6870004300              push 00430070

' *** Reference to "__vbaNew2"
'00493ba5    ff152c124000            call dword ptr [0040122c]
Dim var_2 As New Global
'00493bab    c785f4feffff24be7200    mov dword ptr [ebp+fffffef4], 0072be24
'00493bb5    eb0a                    jmp 493bc1
'00493bb7    c785f4feffff24be7200    mov dword ptr [ebp+fffffef4], 0072be24
'00493bc1    8b85f4feffff            mov eax, dword ptr [ebp+fffffef4]
'00493bc7    8b08                    mov ecx, dword ptr [eax]
'00493bc9    898d24ffffff            mov dword ptr [ebp+ffffff24], ecx
'00493bcf    8d55c0                  lea edx, dword ptr [ebp-40]
'00493bd2    52                      push edx
'00493bd3    8b8524ffffff            mov eax, dword ptr [ebp+ffffff24]
'00493bd9    8b08                    mov ecx, dword ptr [eax]
'00493bdb    8b9524ffffff            mov edx, dword ptr [ebp+ffffff24]
'00493be1    52                      push edx
'00493be2    ff5114                  call dword ptr [ecx+14]
Set var_5 = var_2.App
'00493be5    dbe2                    fnclex
'00493be7    898520ffffff            mov dword ptr [ebp+ffffff20], eax
'00493bed    83bd20ffffff00          cmp dword ptr [ebp+ffffff20], 00
'00493bf4    7d23                    jge 493c19

If (var_2.App < 0) Then
'00493bf6    6a14                    push 14
'00493bf8    6860004300              push 00430060
'00493bfd    8b8524ffffff            mov eax, dword ptr [ebp+ffffff24]
'00493c03    50                      push eax
'00493c04    8b8d20ffffff            mov ecx, dword ptr [ebp+ffffff20]
'00493c0a    51                      push ecx

' *** Reference to "__vbaHresultCheckObj"
'00493c0b    ff1580104000            call dword ptr [00401080]
'00493c11    8985f0feffff            mov dword ptr [ebp+fffffef0], eax
'00493c17    eb0a                    jmp 493c23
    
End If
'00493c19    c785f0feffff00000000    mov dword ptr [ebp+fffffef0], 00000000
'00493c23    8b55c0                  mov edx, dword ptr [ebp-40]
'00493c26    89951cffffff            mov dword ptr [ebp+ffffff1c], edx
'00493c2c    8d45d8                  lea eax, dword ptr [ebp-28]
'00493c2f    50                      push eax
'00493c30    8b8d1cffffff            mov ecx, dword ptr [ebp+ffffff1c]
'00493c36    8b11                    mov edx, dword ptr [ecx]
'00493c38    8b851cffffff            mov eax, dword ptr [ebp+ffffff1c]
'00493c3e    50                      push eax
'00493c3f    ff5250                  call dword ptr [edx+50]
var_45 = var_5.Path
'00493c42    dbe2                    fnclex
'00493c44    898518ffffff            mov dword ptr [ebp+ffffff18], eax
'00493c4a    83bd18ffffff00          cmp dword ptr [ebp+ffffff18], 00
'00493c51    7d23                    jge 493c76

If (0 < 0) Then
'00493c53    6a50                    push 50
'00493c55    682c1c4300              push 00431c2c
'00493c5a    8b8d1cffffff            mov ecx, dword ptr [ebp+ffffff1c]
'00493c60    51                      push ecx
'00493c61    8b9518ffffff            mov edx, dword ptr [ebp+ffffff18]
'00493c67    52                      push edx

' *** Reference to "__vbaHresultCheckObj"
'00493c68    ff1580104000            call dword ptr [00401080]
'00493c6e    8985ecfeffff            mov dword ptr [ebp+fffffeec], eax
'00493c74    eb0a                    jmp 493c80
    
End If
'00493c76    c785ecfeffff00000000    mov dword ptr [ebp+fffffeec], 00000000
'00493c80    bacc134300              mov edx, 004313cc
'00493c85    8d4dc4                  lea ecx, dword ptr [ebp-3c]

' *** Reference to "__vbaStrCopy"
'00493c88    ff1548124000            call dword ptr [00401248]
var_9 = (vbNullChar)
'00493c8e    8b45d8                  mov eax, dword ptr [ebp-28]
'00493c91    898500ffffff            mov dword ptr [ebp+ffffff00], eax
'00493c97    c745d800000000          mov dword ptr [ebp-28], 00000000
'00493c9e    8b9500ffffff            mov edx, dword ptr [ebp+ffffff00]
'00493ca4    8d4dc8                  lea ecx, dword ptr [ebp-38]

' *** Reference to "__vbaStrMove"
'00493ca7    ff15d0124000            call dword ptr [004012d0]
'00493cad    bacc134300              mov edx, 004313cc
'00493cb2    8d4dcc                  lea ecx, dword ptr [ebp-34]

' *** Reference to "__vbaStrCopy"
'00493cb5    ff1548124000            call dword ptr [00401248]
var_43 = (vbNullChar)
'00493cbb    bac81d4300              mov edx, 00431dc8
'00493cc0    8d4dd0                  lea ecx, dword ptr [ebp-30]

' *** Reference to "__vbaStrCopy"
'00493cc3    ff1548124000            call dword ptr [00401248]
var_4 = ("Ouvrir une base donnée pour l'import de personnage")
'00493cc9    c7852cffffff00180800    mov dword ptr [ebp+ffffff2c], 00081800
'00493cd3    c78530ffffff00000000    mov dword ptr [ebp+ffffff30], 00000000
'00493cdd    ba401c4300              mov edx, 00431c40
'00493ce2    8d4dd4                  lea ecx, dword ptr [ebp-2c]

' *** Reference to "__vbaStrCopy"
'00493ce5    ff1548124000            call dword ptr [00401248]
var_44 = ("Access97 (*.mdb)|*.mdb")
'00493ceb    8b8d38ffffff            mov ecx, dword ptr [ebp+ffffff38]
'00493cf1    898d34ffffff            mov dword ptr [ebp+ffffff34], ecx
'00493cf7    8d55c4                  lea edx, dword ptr [ebp-3c]
'00493cfa    52                      push edx
'00493cfb    8d45c8                  lea eax, dword ptr [ebp-38]
'00493cfe    50                      push eax
'00493cff    8d4dcc                  lea ecx, dword ptr [ebp-34]
'00493d02    51                      push ecx
'00493d03    8d55d0                  lea edx, dword ptr [ebp-30]
'00493d06    52                      push edx
'00493d07    8d852cffffff            lea eax, dword ptr [ebp+ffffff2c]
'00493d0d    50                      push eax
'00493d0e    8d8d30ffffff            lea ecx, dword ptr [ebp+ffffff30]
'00493d14    51                      push ecx
'00493d15    8d55d4                  lea edx, dword ptr [ebp-2c]
'00493d18    52                      push edx
'00493d19    8d8534ffffff            lea eax, dword ptr [ebp+ffffff34]
'00493d1f    50                      push eax
'00493d20    e82b220000              call 495f50
Call sub_495F50()
'00493d25    8bd0                    mov edx, eax
'00493d27    8d4ddc                  lea ecx, dword ptr [ebp-24]

' *** Reference to "__vbaStrMove"
'00493d2a    ff15d0124000            call dword ptr [004012d0]
'00493d30    8d4dc4                  lea ecx, dword ptr [ebp-3c]
'00493d33    51                      push ecx
'00493d34    8d55c8                  lea edx, dword ptr [ebp-38]
'00493d37    52                      push edx
'00493d38    8d45cc                  lea eax, dword ptr [ebp-34]
'00493d3b    50                      push eax
'00493d3c    8d4dd0                  lea ecx, dword ptr [ebp-30]
'00493d3f    51                      push ecx
'00493d40    8d55d4                  lea edx, dword ptr [ebp-2c]
'00493d43    52                      push edx
'00493d44    6a05                    push 05

' *** Reference to "__vbaFreeStrList"
'00493d46    ff155c124000            call dword ptr [0040125c]
'#Cleanup( var_44, var_4, var_43, -4504, var_9)
'00493d4c    83c418                  add esp, 18
'00493d4f    8d4dc0                  lea ecx, dword ptr [ebp-40]

' *** Reference to "__vbaFreeObj"
'00493d52    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_5)
'00493d58    c745fc06000000          mov dword ptr [ebp-04], 00000006
'00493d5f    833d60b1720000          cmp dword ptr [0072b160], 00
'00493d66    751c                    jne 493d84
'00493d68    6860b17200              push 0072b160
'00493d6d    68e0a74000              push 0040a7e0

' *** Reference to "__vbaNew2"
'00493d72    ff152c124000            call dword ptr [0040122c]
Dim var_48 As New frmImport
'00493d78    c785e8feffff60b17200    mov dword ptr [ebp+fffffee8], 0072b160
'00493d82    eb0a                    jmp 493d8e
'00493d84    c785e8feffff60b17200    mov dword ptr [ebp+fffffee8], 0072b160
'00493d8e    8b85e8feffff            mov eax, dword ptr [ebp+fffffee8]
'00493d94    8b08                    mov ecx, dword ptr [eax]
'00493d96    898d28ffffff            mov dword ptr [ebp+ffffff28], ecx
'00493d9c    6a00                    push 00
'00493d9e    8b9528ffffff            mov edx, dword ptr [ebp+ffffff28]
'00493da4    8b02                    mov eax, dword ptr [edx]
'00493da6    8b8d28ffffff            mov ecx, dword ptr [ebp+ffffff28]
'00493dac    51                      push ecx
'00493dad    ff90fc060000            call dword ptr [eax+000006fc]
'00493db3    dbe2                    fnclex
'00493db5    898524ffffff            mov dword ptr [ebp+ffffff24], eax
'00493dbb    83bd24ffffff00          cmp dword ptr [ebp+ffffff24], 00
'00493dc2    7d26                    jge 493dea

If (frmImport < 0) Then
'00493dc4    68fc060000              push 000006fc
'00493dc9    68601e4300              push 00431e60
'00493dce    8b9528ffffff            mov edx, dword ptr [ebp+ffffff28]
'00493dd4    52                      push edx
'00493dd5    8b8524ffffff            mov eax, dword ptr [ebp+ffffff24]
'00493ddb    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00493ddc    ff1580104000            call dword ptr [00401080]
'00493de2    8985e4feffff            mov dword ptr [ebp+fffffee4], eax
'00493de8    eb0a                    jmp 493df4
    
End If
'00493dea    c785e4feffff00000000    mov dword ptr [ebp+fffffee4], 00000000
'00493df4    c745fc07000000          mov dword ptr [ebp-04], 00000007
'00493dfb    8b4ddc                  mov ecx, dword ptr [ebp-24]
'00493dfe    51                      push ecx
'00493dff    68cc134300              push 004313cc

' *** Reference to "__vbaStrCmp"
'00493e04    ff1530114000            call dword ptr [00401130]
'00493e0a    85c0                    test ax, ax
'00493e0c    0f84c0060000            je 4944d2

If (((0) <> (vbNullChar))) Then
'00493e12    c745fc08000000          mov dword ptr [ebp-04], 00000008
'00493e19    833d24be720000          cmp dword ptr [0072be24], 00
'00493e20    751c                    jne 493e3e
    
    If (    var_2 = 0) Then
'00493e22    6824be7200              push 0072be24
'00493e27    6870004300              push 00430070

' *** Reference to "__vbaNew2"
'00493e2c    ff152c124000            call dword ptr [0040122c]
    Set var_2 = New Global
'00493e32    c785e0feffff24be7200    mov dword ptr [ebp+fffffee0], 0072be24
'00493e3c    eb0a                    jmp 493e48
    
End If
'00493e3e    c785e0feffff24be7200    mov dword ptr [ebp+fffffee0], 0072be24
'00493e48    8b95e0feffff            mov edx, dword ptr [ebp+fffffee0]
'00493e4e    8b02                    mov eax, dword ptr [edx]
'00493e50    898528ffffff            mov dword ptr [ebp+ffffff28], eax
'00493e56    8d4dc0                  lea ecx, dword ptr [ebp-40]
'00493e59    51                      push ecx
'00493e5a    8b9528ffffff            mov edx, dword ptr [ebp+ffffff28]
'00493e60    8b02                    mov eax, dword ptr [edx]
'00493e62    8b8d28ffffff            mov ecx, dword ptr [ebp+ffffff28]
'00493e68    51                      push ecx
'00493e69    ff5014                  call dword ptr [eax+14]
Set var_5 = var_2.App
'00493e6c    dbe2                    fnclex
'00493e6e    898524ffffff            mov dword ptr [ebp+ffffff24], eax
'00493e74    83bd24ffffff00          cmp dword ptr [ebp+ffffff24], 00
'00493e7b    7d23                    jge 493ea0

If (var_2.App < 0) Then
'00493e7d    6a14                    push 14
'00493e7f    6860004300              push 00430060
'00493e84    8b9528ffffff            mov edx, dword ptr [ebp+ffffff28]
'00493e8a    52                      push edx
'00493e8b    8b8524ffffff            mov eax, dword ptr [ebp+ffffff24]
'00493e91    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00493e92    ff1580104000            call dword ptr [00401080]
'00493e98    8985dcfeffff            mov dword ptr [ebp+fffffedc], eax
'00493e9e    eb0a                    jmp 493eaa
    
End If
'00493ea0    c785dcfeffff00000000    mov dword ptr [ebp+fffffedc], 00000000
'00493eaa    8b4dc0                  mov ecx, dword ptr [ebp-40]
'00493ead    898d20ffffff            mov dword ptr [ebp+ffffff20], ecx
'00493eb3    8d55d8                  lea edx, dword ptr [ebp-28]
'00493eb6    52                      push edx
'00493eb7    8b8520ffffff            mov eax, dword ptr [ebp+ffffff20]
'00493ebd    8b08                    mov ecx, dword ptr [eax]
'00493ebf    8b9520ffffff            mov edx, dword ptr [ebp+ffffff20]
'00493ec5    52                      push edx
'00493ec6    ff5150                  call dword ptr [ecx+50]
var_45 = var_5.Path
'00493ec9    dbe2                    fnclex
'00493ecb    89851cffffff            mov dword ptr [ebp+ffffff1c], eax
'00493ed1    83bd1cffffff00          cmp dword ptr [ebp+ffffff1c], 00
'00493ed8    7d23                    jge 493efd

If (0 < 0) Then
'00493eda    6a50                    push 50
'00493edc    682c1c4300              push 00431c2c
'00493ee1    8b8520ffffff            mov eax, dword ptr [ebp+ffffff20]
'00493ee7    50                      push eax
'00493ee8    8b8d1cffffff            mov ecx, dword ptr [ebp+ffffff1c]
'00493eee    51                      push ecx

' *** Reference to "__vbaHresultCheckObj"
'00493eef    ff1580104000            call dword ptr [00401080]
'00493ef5    8985d8feffff            mov dword ptr [ebp+fffffed8], eax
'00493efb    eb0a                    jmp 493f07
    
End If
'00493efd    c785d8feffff00000000    mov dword ptr [ebp+fffffed8], 00000000
'00493f07    8b55d8                  mov edx, dword ptr [ebp-28]
'00493f0a    52                      push edx
'00493f0b    68641f4300              push 00431f64

' *** Reference to "__vbaStrCat"
'00493f10    ff1570104000            call dword ptr [00401070]
var_49 = (var_45) & ("\roles.mdb")
'00493f16    8945b4                  mov dword ptr [ebp-4c], eax
'00493f19    c745ac08000000          mov dword ptr [ebp-54], 00000008
'00493f20    8d45ac                  lea eax, dword ptr [ebp-54]
'00493f23    50                      push eax
'00493f24    8d4d9c                  lea ecx, dword ptr [ebp-64]
'00493f27    51                      push ecx

' *** Reference to "rtcUpperCaseVar"
'00493f28    ff152c114000            call dword ptr [0040112c]
'00493f2e    8d55dc                  lea edx, dword ptr [ebp-24]
'00493f31    899574ffffff            mov dword ptr [ebp+ffffff74], edx
'00493f37    c7856cffffff08400000    mov dword ptr [ebp+ffffff6c], 00004008
'00493f41    8d856cffffff            lea eax, dword ptr [ebp+ffffff6c]
'00493f47    50                      push eax
'00493f48    8d4d8c                  lea ecx, dword ptr [ebp-74]
'00493f4b    51                      push ecx

' *** Reference to "rtcUpperCaseVar"
'00493f4c    ff152c114000            call dword ptr [0040112c]
'00493f52    8d559c                  lea edx, dword ptr [ebp-64]
'00493f55    52                      push edx
'00493f56    8d458c                  lea eax, dword ptr [ebp-74]
'00493f59    50                      push eax

' *** Reference to "__vbaVarTstEq"
'00493f5a    ff153c114000            call dword ptr [0040113c]
'00493f60    66898518ffffff          mov word ptr [ebp+ffffff18], ax
'00493f67    8d4dd8                  lea ecx, dword ptr [ebp-28]

' *** Reference to "__vbaFreeStr"
'00493f6a    ff150c134000            call dword ptr [0040130c]
'#Cleanup(var_45)
'00493f70    8d4dc0                  lea ecx, dword ptr [ebp-40]

' *** Reference to "__vbaFreeObj"
'00493f73    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_5)
'00493f79    8d4d8c                  lea ecx, dword ptr [ebp-74]
'00493f7c    51                      push ecx
'00493f7d    8d559c                  lea edx, dword ptr [ebp-64]
'00493f80    52                      push edx
'00493f81    8d45ac                  lea eax, dword ptr [ebp-54]
'00493f84    50                      push eax
'00493f85    6a03                    push 03

' *** Reference to "__vbaFreeVarList"
'00493f87    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_50, var_51, var_52)
'00493f8d    83c410                  add esp, 10
'00493f90    0fbf8d18ffffff          movsx ecx, word ptr [ebp+ffffff18]
'00493f97    85c9                    test cx, cx
'00493f99    0f8491000000            je 494030

If (((UCase(var_49)) = (UCase(0)))) Then
'00493f9f    c745fc09000000          mov dword ptr [ebp-04], 00000009
'00493fa6    c7459404000280          mov dword ptr [ebp-6c], 80020004
'00493fad    c7458c0a000000          mov dword ptr [ebp-74], 0000000a
'00493fb4    c745a404000280          mov dword ptr [ebp-5c], 80020004
'00493fbb    c7459c0a000000          mov dword ptr [ebp-64], 0000000a
'00493fc2    c78564ffffff24b07200    mov dword ptr [ebp+ffffff64], 0072b024
'00493fcc    c7855cffffff08400000    mov dword ptr [ebp+ffffff5c], 00004008
'00493fd6    c78574ffffff801f4300    mov dword ptr [ebp+ffffff74], 00431f80
'00493fe0    c7856cffffff08000000    mov dword ptr [ebp+ffffff6c], 00000008
'00493fea    8d956cffffff            lea edx, dword ptr [ebp+ffffff6c]
'00493ff0    8d4dac                  lea ecx, dword ptr [ebp-54]

' *** Reference to "__vbaVarDup"
'00493ff3    ff1598124000            call dword ptr [00401298]
    var_50 = ("On ne peut pas ouvrir la base de donnée du programme !")
'00493ff9    8d558c                  lea edx, dword ptr [ebp-74]
'00493ffc    52                      push edx
'00493ffd    8d459c                  lea eax, dword ptr [ebp-64]
'00494000    50                      push eax
'00494001    8d8d5cffffff            lea ecx, dword ptr [ebp+ffffff5c]
'00494007    51                      push ecx
'00494008    6a30                    push 30
'0049400a    8d55ac                  lea edx, dword ptr [ebp-54]
'0049400d    52                      push edx

' *** Reference to "rtcMsgBox"
'0049400e    ff15b8104000            call dword ptr [004010b8]
    var_53 = MsgBox(var_50, 48, vbNullString)
'00494014    8d458c                  lea eax, dword ptr [ebp-74]
'00494017    50                      push eax
'00494018    8d4d9c                  lea ecx, dword ptr [ebp-64]
'0049401b    51                      push ecx
'0049401c    8d55ac                  lea edx, dword ptr [ebp-54]
'0049401f    52                      push edx
'00494020    6a03                    push 03

' *** Reference to "__vbaFreeVarList"
'00494022    ff1540104000            call dword ptr [00401040]
    '#Cleanup( var_50, var_51, var_52)
'00494028    83c410                  add esp, 10
'0049402b    e9a2040000              jmp 4944d2
    
End If
'00494030    c745fc0c000000          mov dword ptr [ebp-04], 0000000c
'00494037    8b45dc                  mov eax, dword ptr [ebp-24]
'0049403a    50                      push eax
'0049403b    8b0d3cb07200            mov ecx, dword ptr [0072b03c]
'00494041    51                      push ecx

' *** Reference to "__vbaStrCmp"
'00494042    ff1530114000            call dword ptr [00401130]
'00494048    85c0                    test ax, ax
'0049404a    0f853f010000            jne 49418f

If (((0) = (vbNullString))) Then
'00494050    c745fc0d000000          mov dword ptr [ebp-04], 0000000d
'00494057    c7459404000280          mov dword ptr [ebp-6c], 80020004
'0049405e    c7458c0a000000          mov dword ptr [ebp-74], 0000000a
'00494065    c745a404000280          mov dword ptr [ebp-5c], 80020004
'0049406c    c7459c0a000000          mov dword ptr [ebp-64], 0000000a
'00494073    c78574ffffff24b07200    mov dword ptr [ebp+ffffff74], 0072b024
'0049407d    c7856cffffff08400000    mov dword ptr [ebp+ffffff6c], 00004008
'00494087    68f41f4300              push 00431ff4
'0049408c    6870084300              push 00430870

' *** Reference to "__vbaStrCat"
'00494091    ff1570104000            call dword ptr [00401070]
    var_54 = ("Vous ouvrez la base en cours d'utilisation.") & (vbCrLf)
'00494097    8bd0                    mov edx, eax
'00494099    8d4dd8                  lea ecx, dword ptr [ebp-28]

' *** Reference to "__vbaStrMove"
'0049409c    ff15d0124000            call dword ptr [004012d0]
'004940a2    50                      push eax
'004940a3    6888204300              push 00432088

' *** Reference to "__vbaStrCat"
'004940a8    ff1570104000            call dword ptr [00401070]
    var_55 = (var_54) & ("Vous pourrez seulement dupliquer les personnages en les renommant.")
'004940ae    8945b4                  mov dword ptr [ebp-4c], eax
'004940b1    c745ac08000000          mov dword ptr [ebp-54], 00000008
'004940b8    8d558c                  lea edx, dword ptr [ebp-74]
'004940bb    52                      push edx
'004940bc    8d459c                  lea eax, dword ptr [ebp-64]
'004940bf    50                      push eax
'004940c0    8d8d6cffffff            lea ecx, dword ptr [ebp+ffffff6c]
'004940c6    51                      push ecx
'004940c7    6a30                    push 30
'004940c9    8d55ac                  lea edx, dword ptr [ebp-54]
'004940cc    52                      push edx

' *** Reference to "rtcMsgBox"
'004940cd    ff15b8104000            call dword ptr [004010b8]
    var_56 = MsgBox(var_55, 48, vbNullString)
'004940d3    8d4dd8                  lea ecx, dword ptr [ebp-28]

' *** Reference to "__vbaFreeStr"
'004940d6    ff150c134000            call dword ptr [0040130c]
    '#Cleanup(var_45)
'004940dc    8d458c                  lea eax, dword ptr [ebp-74]
'004940df    50                      push eax
'004940e0    8d4d9c                  lea ecx, dword ptr [ebp-64]
'004940e3    51                      push ecx
'004940e4    8d55ac                  lea edx, dword ptr [ebp-54]
'004940e7    52                      push edx
'004940e8    6a03                    push 03

' *** Reference to "__vbaFreeVarList"
'004940ea    ff1540104000            call dword ptr [00401040]
    '#Cleanup( var_50, var_51, var_52)
'004940f0    83c410                  add esp, 10
'004940f3    c745fc0e000000          mov dword ptr [ebp-04], 0000000e
'004940fa    833d60b1720000          cmp dword ptr [0072b160], 00
'00494101    751c                    jne 49411f
    
    If (    var_48 = 0) Then
'00494103    6860b17200              push 0072b160
'00494108    68e0a74000              push 0040a7e0

' *** Reference to "__vbaNew2"
'0049410d    ff152c124000            call dword ptr [0040122c]
    Set var_48 = New frmImport
'00494113    c785d4feffff60b17200    mov dword ptr [ebp+fffffed4], 0072b160
'0049411d    eb0a                    jmp 494129
    
End If
'0049411f    c785d4feffff60b17200    mov dword ptr [ebp+fffffed4], 0072b160
'00494129    8b85d4feffff            mov eax, dword ptr [ebp+fffffed4]
'0049412f    8b08                    mov ecx, dword ptr [eax]
'00494131    898d28ffffff            mov dword ptr [ebp+ffffff28], ecx
'00494137    6aff                    push ffffffff
'00494139    8b9528ffffff            mov edx, dword ptr [ebp+ffffff28]
'0049413f    8b02                    mov eax, dword ptr [edx]
'00494141    8b8d28ffffff            mov ecx, dword ptr [ebp+ffffff28]
'00494147    51                      push ecx
'00494148    ff90fc060000            call dword ptr [eax+000006fc]
'0049414e    dbe2                    fnclex
'00494150    898524ffffff            mov dword ptr [ebp+ffffff24], eax
'00494156    83bd24ffffff00          cmp dword ptr [ebp+ffffff24], 00
'0049415d    7d26                    jge 494185

If (frmImport < 0) Then
'0049415f    68fc060000              push 000006fc
'00494164    68601e4300              push 00431e60
'00494169    8b9528ffffff            mov edx, dword ptr [ebp+ffffff28]
'0049416f    52                      push edx
'00494170    8b8524ffffff            mov eax, dword ptr [ebp+ffffff24]
'00494176    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00494177    ff1580104000            call dword ptr [00401080]
'0049417d    8985d0feffff            mov dword ptr [ebp+fffffed0], eax
'00494183    eb0a                    jmp 49418f
    
End If
'00494185    c785d0feffff00000000    mov dword ptr [ebp+fffffed0], 00000000

'ERROR: Two many next close:
End If
'0049418f    c745fc10000000          mov dword ptr [ebp-04], 00000010
'00494196    833d60b1720000          cmp dword ptr [0072b160], 00
'0049419d    751c                    jne 4941bb

If (var_48 = 0) Then
'0049419f    6860b17200              push 0072b160
'004941a4    68e0a74000              push 0040a7e0

' *** Reference to "__vbaNew2"
'004941a9    ff152c124000            call dword ptr [0040122c]
    Set var_48 = New frmImport
'004941af    c785ccfeffff60b17200    mov dword ptr [ebp+fffffecc], 0072b160
'004941b9    eb0a                    jmp 4941c5
    
End If
'004941bb    c785ccfeffff60b17200    mov dword ptr [ebp+fffffecc], 0072b160
'004941c5    8b8dccfeffff            mov ecx, dword ptr [ebp+fffffecc]
'004941cb    8b11                    mov edx, dword ptr [ecx]
'004941cd    899520ffffff            mov dword ptr [ebp+ffffff20], edx
'004941d3    833d28be720000          cmp dword ptr [0072be28], 00
'004941da    751c                    jne 4941f8
'004941dc    6828be7200              push 0072be28
'004941e1    6830214300              push 00432130

' *** Reference to "__vbaNew2"
'004941e6    ff152c124000            call dword ptr [0040122c]
Dim var_57 As New DBEngine
'004941ec    c785c8feffff28be7200    mov dword ptr [ebp+fffffec8], 0072be28
'004941f6    eb0a                    jmp 494202
'004941f8    c785c8feffff28be7200    mov dword ptr [ebp+fffffec8], 0072be28
'00494202    8b85c8feffff            mov eax, dword ptr [ebp+fffffec8]
'00494208    8b08                    mov ecx, dword ptr [eax]
'0049420a    898d28ffffff            mov dword ptr [ebp+ffffff28], ecx
'00494210    c78554ffffff04000280    mov dword ptr [ebp+ffffff54], 80020004
'0049421a    c7854cffffff0a000000    mov dword ptr [ebp+ffffff4c], 0000000a
'00494224    c78564ffffff04000280    mov dword ptr [ebp+ffffff64], 80020004
'0049422e    c7855cffffff0a000000    mov dword ptr [ebp+ffffff5c], 0000000a
'00494238    c78574ffffff04000280    mov dword ptr [ebp+ffffff74], 80020004
'00494242    c7856cffffff0a000000    mov dword ptr [ebp+ffffff6c], 0000000a
'0049424c    8d55c0                  lea edx, dword ptr [ebp-40]
'0049424f    52                      push edx
'00494250    b810000000              mov eax, 00000010

' *** Reference to "__vbaChkstk"
'00494255    e80630f7ff              call 407260
'0049425a    8bc4                    mov eax, esp
'0049425c    8b8d4cffffff            mov ecx, dword ptr [ebp+ffffff4c]
'00494262    8908                    mov dword ptr [eax], ecx
'00494264    8b9550ffffff            mov edx, dword ptr [ebp+ffffff50]
'0049426a    895004                  mov dword ptr [eax+04], edx
'0049426d    8b8d54ffffff            mov ecx, dword ptr [ebp+ffffff54]
'00494273    894808                  mov dword ptr [eax+08], ecx
'00494276    8b9558ffffff            mov edx, dword ptr [ebp+ffffff58]
'0049427c    89500c                  mov dword ptr [eax+0c], edx
'0049427f    b810000000              mov eax, 00000010

' *** Reference to "__vbaChkstk"
'00494284    e8d72ff7ff              call 407260
'00494289    8bc4                    mov eax, esp
'0049428b    8b8d5cffffff            mov ecx, dword ptr [ebp+ffffff5c]
'00494291    8908                    mov dword ptr [eax], ecx
'00494293    8b9560ffffff            mov edx, dword ptr [ebp+ffffff60]
'00494299    895004                  mov dword ptr [eax+04], edx
'0049429c    8b8d64ffffff            mov ecx, dword ptr [ebp+ffffff64]
'004942a2    894808                  mov dword ptr [eax+08], ecx
'004942a5    8b9568ffffff            mov edx, dword ptr [ebp+ffffff68]
'004942ab    89500c                  mov dword ptr [eax+0c], edx
'004942ae    b810000000              mov eax, 00000010

' *** Reference to "__vbaChkstk"
'004942b3    e8a82ff7ff              call 407260
'004942b8    8bc4                    mov eax, esp
'004942ba    8b8d6cffffff            mov ecx, dword ptr [ebp+ffffff6c]
'004942c0    8908                    mov dword ptr [eax], ecx
'004942c2    8b9570ffffff            mov edx, dword ptr [ebp+ffffff70]
'004942c8    895004                  mov dword ptr [eax+04], edx
'004942cb    8b8d74ffffff            mov ecx, dword ptr [ebp+ffffff74]
'004942d1    894808                  mov dword ptr [eax+08], ecx
'004942d4    8b9578ffffff            mov edx, dword ptr [ebp+ffffff78]
'004942da    89500c                  mov dword ptr [eax+0c], edx
'004942dd    8b45dc                  mov eax, dword ptr [ebp-24]
'004942e0    50                      push eax
'004942e1    8b8d28ffffff            mov ecx, dword ptr [ebp+ffffff28]
'004942e7    8b11                    mov edx, dword ptr [ecx]
'004942e9    8b8528ffffff            mov eax, dword ptr [ebp+ffffff28]
'004942ef    50                      push eax
'004942f0    ff5258                  call dword ptr [edx+58]
Set var_5 = var_57.OpenDatabase(0)
'004942f3    dbe2                    fnclex
'004942f5    898524ffffff            mov dword ptr [ebp+ffffff24], eax
'004942fb    83bd24ffffff00          cmp dword ptr [ebp+ffffff24], 00
'00494302    7d23                    jge 494327

If (0 < 0) Then
'00494304    6a58                    push 58
'00494306    6820214300              push 00432120
'0049430b    8b8d28ffffff            mov ecx, dword ptr [ebp+ffffff28]
'00494311    51                      push ecx
'00494312    8b9524ffffff            mov edx, dword ptr [ebp+ffffff24]
'00494318    52                      push edx

' *** Reference to "__vbaHresultCheckObj"
'00494319    ff1580104000            call dword ptr [00401080]
'0049431f    8985c4feffff            mov dword ptr [ebp+fffffec4], eax
'00494325    eb0a                    jmp 494331
    
End If
'00494327    c785c4feffff00000000    mov dword ptr [ebp+fffffec4], 00000000
'00494331    8b45c0                  mov eax, dword ptr [ebp-40]
'00494334    8985fcfeffff            mov dword ptr [ebp+fffffefc], eax
'0049433a    c745c000000000          mov dword ptr [ebp-40], 00000000
'00494341    8b8dfcfeffff            mov ecx, dword ptr [ebp+fffffefc]
'00494347    51                      push ecx
'00494348    8d55bc                  lea edx, dword ptr [ebp-44]
'0049434b    52                      push edx

' *** Reference to "__vbaObjSet"
'0049434c    ff15b4104000            call dword ptr [004010b4]
Set var_58 = var_5
'00494352    50                      push eax
'00494353    8b8520ffffff            mov eax, dword ptr [ebp+ffffff20]
'00494359    8b08                    mov ecx, dword ptr [eax]
'0049435b    8b9520ffffff            mov edx, dword ptr [ebp+ffffff20]
'00494361    52                      push edx
'00494362    ff9108070000            call dword ptr [ecx+00000708]
'00494368    dbe2                    fnclex
'0049436a    89851cffffff            mov dword ptr [ebp+ffffff1c], eax
'00494370    83bd1cffffff00          cmp dword ptr [ebp+ffffff1c], 00
'00494377    7d26                    jge 49439f

If (var_48 < 0) Then
'00494379    6808070000              push 00000708
'0049437e    68601e4300              push 00431e60
'00494383    8b8520ffffff            mov eax, dword ptr [ebp+ffffff20]
'00494389    50                      push eax
'0049438a    8b8d1cffffff            mov ecx, dword ptr [ebp+ffffff1c]
'00494390    51                      push ecx

' *** Reference to "__vbaHresultCheckObj"
'00494391    ff1580104000            call dword ptr [00401080]
'00494397    8985c0feffff            mov dword ptr [ebp+fffffec0], eax
'0049439d    eb0a                    jmp 4943a9
    
End If
'0049439f    c785c0feffff00000000    mov dword ptr [ebp+fffffec0], 00000000
'004943a9    8d4dbc                  lea ecx, dword ptr [ebp-44]

' *** Reference to "__vbaFreeObj"
'004943ac    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_58)
'004943b2    c745fc11000000          mov dword ptr [ebp-04], 00000011
'004943b9    833d60b1720000          cmp dword ptr [0072b160], 00
'004943c0    751c                    jne 4943de

If (var_48 = 0) Then
'004943c2    6860b17200              push 0072b160
'004943c7    68e0a74000              push 0040a7e0

' *** Reference to "__vbaNew2"
'004943cc    ff152c124000            call dword ptr [0040122c]
    Set var_48 = New frmImport
'004943d2    c785bcfeffff60b17200    mov dword ptr [ebp+fffffebc], 0072b160
'004943dc    eb0a                    jmp 4943e8
    
End If
'004943de    c785bcfeffff60b17200    mov dword ptr [ebp+fffffebc], 0072b160
'004943e8    8b95bcfeffff            mov edx, dword ptr [ebp+fffffebc]
'004943ee    8b02                    mov eax, dword ptr [edx]
'004943f0    898528ffffff            mov dword ptr [ebp+ffffff28], eax
'004943f6    c78564ffffff04000280    mov dword ptr [ebp+ffffff64], 80020004
'00494400    c7855cffffff0a000000    mov dword ptr [ebp+ffffff5c], 0000000a
'0049440a    c78574ffffff01000000    mov dword ptr [ebp+ffffff74], 00000001
'00494414    c7856cffffff03000000    mov dword ptr [ebp+ffffff6c], 00000003
'0049441e    b810000000              mov eax, 00000010

' *** Reference to "__vbaChkstk"
'00494423    e8382ef7ff              call 407260
'00494428    8bcc                    mov ecx, esp
'0049442a    8b955cffffff            mov edx, dword ptr [ebp+ffffff5c]
'00494430    8911                    mov dword ptr [ecx], edx
'00494432    8b8560ffffff            mov eax, dword ptr [ebp+ffffff60]
'00494438    894104                  mov dword ptr [ecx+04], eax
'0049443b    8b9564ffffff            mov edx, dword ptr [ebp+ffffff64]
'00494441    895108                  mov dword ptr [ecx+08], edx
'00494444    8b8568ffffff            mov eax, dword ptr [ebp+ffffff68]
'0049444a    89410c                  mov dword ptr [ecx+0c], eax
'0049444d    b810000000              mov eax, 00000010

' *** Reference to "__vbaChkstk"
'00494452    e8092ef7ff              call 407260
'00494457    8bcc                    mov ecx, esp
'00494459    8b956cffffff            mov edx, dword ptr [ebp+ffffff6c]
'0049445f    8911                    mov dword ptr [ecx], edx
'00494461    8b8570ffffff            mov eax, dword ptr [ebp+ffffff70]
'00494467    894104                  mov dword ptr [ecx+04], eax
'0049446a    8b9574ffffff            mov edx, dword ptr [ebp+ffffff74]
'00494470    895108                  mov dword ptr [ecx+08], edx
'00494473    8b8578ffffff            mov eax, dword ptr [ebp+ffffff78]
'00494479    89410c                  mov dword ptr [ecx+0c], eax
'0049447c    8b8d28ffffff            mov ecx, dword ptr [ebp+ffffff28]
'00494482    8b11                    mov edx, dword ptr [ecx]
'00494484    8b8528ffffff            mov eax, dword ptr [ebp+ffffff28]
'0049448a    50                      push eax
'0049448b    ff92b0020000            call dword ptr [edx+000002b0]
'00494491    dbe2                    fnclex
'00494493    898524ffffff            mov dword ptr [ebp+ffffff24], eax
'00494499    83bd24ffffff00          cmp dword ptr [ebp+ffffff24], 00
'004944a0    7d26                    jge 4944c8

If (var_48 < 0) Then
'004944a2    68b0020000              push 000002b0
'004944a7    68301e4300              push 00431e30
'004944ac    8b8d28ffffff            mov ecx, dword ptr [ebp+ffffff28]
'004944b2    51                      push ecx
'004944b3    8b9524ffffff            mov edx, dword ptr [ebp+ffffff24]
'004944b9    52                      push edx

' *** Reference to "__vbaHresultCheckObj"
'004944ba    ff1580104000            call dword ptr [00401080]
'004944c0    8985b8feffff            mov dword ptr [ebp+fffffeb8], eax
'004944c6    eb0a                    jmp 4944d2
    
End If
'004944c8    c785b8feffff00000000    mov dword ptr [ebp+fffffeb8], 00000000

'ERROR: Two many next close:
End If
'004944d2    c745f000000000          mov dword ptr [ebp-10], 00000000
'004944d9    683f454900              push 0049453f
'004944de    eb55                    jmp 494535
'004944e0    8d45c4                  lea eax, dword ptr [ebp-3c]
'004944e3    50                      push eax
'004944e4    8d4dc8                  lea ecx, dword ptr [ebp-38]
'004944e7    51                      push ecx
'004944e8    8d55cc                  lea edx, dword ptr [ebp-34]
'004944eb    52                      push edx
'004944ec    8d45d0                  lea eax, dword ptr [ebp-30]
'004944ef    50                      push eax
'004944f0    8d4dd4                  lea ecx, dword ptr [ebp-2c]
'004944f3    51                      push ecx
'004944f4    8d55d8                  lea edx, dword ptr [ebp-28]
'004944f7    52                      push edx
'004944f8    6a06                    push 06

' *** Reference to "__vbaFreeStrList"
'004944fa    ff155c124000            call dword ptr [0040125c]
'#Cleanup( , var_44, var_4, var_43, -4504, var_9)
'00494500    83c41c                  add esp, 1c
'00494503    8d45bc                  lea eax, dword ptr [ebp-44]
'00494506    50                      push eax
'00494507    8d4dc0                  lea ecx, dword ptr [ebp-40]
'0049450a    51                      push ecx
'0049450b    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'0049450d    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_5, var_58)
'00494513    83c40c                  add esp, 0c
'00494516    8d957cffffff            lea edx, dword ptr [ebp+ffffff7c]
'0049451c    52                      push edx
'0049451d    8d458c                  lea eax, dword ptr [ebp-74]
'00494520    50                      push eax
'00494521    8d4d9c                  lea ecx, dword ptr [ebp-64]
'00494524    51                      push ecx
'00494525    8d55ac                  lea edx, dword ptr [ebp-54]
'00494528    52                      push edx
'00494529    6a04                    push 04

' *** Reference to "__vbaFreeVarList"
'0049452b    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_50, var_51, var_52, var_59)
'00494531    83c414                  add esp, 14
'00494534    c3                      ret
'00494535    8d4ddc                  lea ecx, dword ptr [ebp-24]

' *** Reference to "__vbaFreeStr"
'00494538    ff150c134000            call dword ptr [0040130c]
'#Cleanup(var_39)
'0049453e    c3                      ret
'0049453f    8b4508                  mov eax, dword ptr [ebp+08]
'00494542    8b08                    mov ecx, dword ptr [eax]
'00494544    8b5508                  mov edx, dword ptr [ebp+08]
'00494547    52                      push edx
'00494548    ff5108                  call dword ptr [ecx+08]
'0049454b    8b45f0                  mov eax, dword ptr [ebp-10]
'0049454e    8b4de0                  mov ecx, dword ptr [ebp-20]
    'Reference to '__except_list'
'00494551    64890d00000000          mov dword ptr fs:[00000000], ecx
'00494558    5f                      pop edi
'00494559    5e                      pop esi
'0049455a    5b                      pop ebx
'0049455b    8be5                    mov esp, ebp
'0049455d    5d                      pop ebp
'0049455e    c20400                  ret 0004


End Sub


'Event for IDM_DESCRIPTION_SORT
Private Sub IDM_DESCRIPTION_SORT_Click()
'004924c0    55                      push ebp
'004924c1    8bec                    mov ebp, esp
'004924c3    83ec0c                  sub esp, 0c

' *** Reference to "__vbaExceptHandler"
'004924c6    6866724000              push 00407266
'004924cb    64a100000000            mov ax, word ptr fs:[00000000]
'004924d1    50                      push eax
    'Reference to '__except_list'
'004924d2    64892500000000          mov dword ptr fs:[00000000], esp
'004924d9    83ec30                  sub esp, 30
'004924dc    53                      push ebx
'004924dd    56                      push esi
'004924de    57                      push edi
'004924df    8965f4                  mov dword ptr [ebp-0c], esp
'004924e2    c745f878134000          mov dword ptr [ebp-08], 00401378
'004924e9    8b4508                  mov eax, dword ptr [ebp+08]
'004924ec    8bc8                    mov ecx, eax
'004924ee    83e101                  and ecx, 01
'004924f1    894dfc                  mov dword ptr [ebp-04], ecx
'004924f4    83e0fe                  and eax, fffffffe
'004924f7    8b10                    mov edx, dword ptr [eax]
'004924f9    50                      push eax
'004924fa    894508                  mov dword ptr [ebp+08], eax
'004924fd    ff5204                  call dword ptr [edx+04]
'00492500    a1e8b07200              mov ax, word ptr [0072b0e8]
'00492505    85c0                    test ax, ax
'00492507    7510                    jne 492519
'00492509    68e8b07200              push 0072b0e8
'0049250e    68d01c4100              push 00411cd0

' *** Reference to "__vbaNew2"
'00492513    ff152c124000            call dword ptr [0040122c]
Dim var_33 As New frmDescriptifSort
'00492519    8b35e8b07200            mov esi, dword ptr [0072b0e8]
'0049251f    8b06                    mov eax, dword ptr [esi]
'00492521    68cc134300              push 004313cc
'00492526    56                      push esi
'00492527    ff9024070000            call dword ptr [eax+00000724]
'0049252d    dbe2                    fnclex
'0049252f    85c0                    test ax, ax
'00492531    7d12                    jge 492545
'00492533    6824070000              push 00000724
'00492538    6800144300              push 00431400
'0049253d    56                      push esi
'0049253e    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0049253f    ff1580104000            call dword ptr [00401080]
'00492545    a1e8b07200              mov ax, word ptr [0072b0e8]
'0049254a    85c0                    test ax, ax
'0049254c    7510                    jne 49255e
'0049254e    68e8b07200              push 0072b0e8
'00492553    68d01c4100              push 00411cd0

' *** Reference to "__vbaNew2"
'00492558    ff152c124000            call dword ptr [0040122c]
Set var_33 = New frmDescriptifSort
'0049255e    8b35e8b07200            mov esi, dword ptr [0072b0e8]
'00492564    8b3e                    mov edi, dword ptr [esi]
'00492566    83ec10                  sub esp, 10
'00492569    8bdc                    mov ebx, esp
'0049256b    b90a000000              mov ecx, 0000000a
'00492570    890b                    mov dword ptr [ebx], ecx
'00492572    8b4dd0                  mov ecx, dword ptr [ebp-30]
'00492575    894b04                  mov dword ptr [ebx+04], ecx
'00492578    b804000280              mov eax, 80020004
'0049257d    894308                  mov dword ptr [ebx+08], eax
'00492580    8b45d8                  mov eax, dword ptr [ebp-28]
'00492583    89430c                  mov dword ptr [ebx+0c], eax
'00492586    83ec10                  sub esp, 10
'00492589    8bcc                    mov ecx, esp
'0049258b    b803000000              mov eax, 00000003
'00492590    8901                    mov dword ptr [ecx], eax
'00492592    8b45e0                  mov eax, dword ptr [ebp-20]
'00492595    894104                  mov dword ptr [ecx+04], eax
'00492598    ba01000000              mov edx, 00000001
'0049259d    895108                  mov dword ptr [ecx+08], edx
'004925a0    8b55e8                  mov edx, dword ptr [ebp-18]
'004925a3    56                      push esi
'004925a4    89510c                  mov dword ptr [ecx+0c], edx
'004925a7    ff97b0020000            call dword ptr [edi+000002b0]
'004925ad    dbe2                    fnclex
'004925af    85c0                    test ax, ax
'004925b1    7d12                    jge 4925c5
'004925b3    68b0020000              push 000002b0
'004925b8    68d0134300              push 004313d0
'004925bd    56                      push esi
'004925be    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'004925bf    ff1580104000            call dword ptr [00401080]
'004925c5    c745fc00000000          mov dword ptr [ebp-04], 00000000
'004925cc    8b4508                  mov eax, dword ptr [ebp+08]
'004925cf    8b08                    mov ecx, dword ptr [eax]
'004925d1    50                      push eax
'004925d2    ff5108                  call dword ptr [ecx+08]
'004925d5    8b45fc                  mov eax, dword ptr [ebp-04]
'004925d8    8b4dec                  mov ecx, dword ptr [ebp-14]
'004925db    5f                      pop edi
'004925dc    5e                      pop esi
    'Reference to '__except_list'
'004925dd    64890d00000000          mov dword ptr fs:[00000000], ecx
'004925e4    5b                      pop ebx
'004925e5    8be5                    mov esp, ebp
'004925e7    5d                      pop ebp
'004925e8    c20400                  ret 0004


End Sub


'Event for IDM_XP_PERSONNAGES
Private Sub IDM_XP_PERSONNAGES_Click()
'004939d0    55                      push ebp
'004939d1    8bec                    mov ebp, esp
'004939d3    83ec0c                  sub esp, 0c

' *** Reference to "__vbaExceptHandler"
'004939d6    6866724000              push 00407266
'004939db    64a100000000            mov ax, word ptr fs:[00000000]
'004939e1    50                      push eax
    'Reference to '__except_list'
'004939e2    64892500000000          mov dword ptr fs:[00000000], esp
'004939e9    83ec30                  sub esp, 30
'004939ec    53                      push ebx
'004939ed    56                      push esi
'004939ee    57                      push edi
'004939ef    8965f4                  mov dword ptr [ebp-0c], esp
'004939f2    c745f850144000          mov dword ptr [ebp-08], 00401450
'004939f9    8b4508                  mov eax, dword ptr [ebp+08]
'004939fc    8bc8                    mov ecx, eax
'004939fe    83e101                  and ecx, 01
'00493a01    894dfc                  mov dword ptr [ebp-04], ecx
'00493a04    83e0fe                  and eax, fffffffe
'00493a07    8b10                    mov edx, dword ptr [eax]
'00493a09    50                      push eax
'00493a0a    894508                  mov dword ptr [ebp+08], eax
'00493a0d    ff5204                  call dword ptr [edx+04]
'00493a10    a14cb17200              mov ax, word ptr [0072b14c]
'00493a15    85c0                    test ax, ax
'00493a17    7510                    jne 493a29
'00493a19    684cb17200              push 0072b14c
'00493a1e    68c4d24000              push 0040d2c4

' *** Reference to "__vbaNew2"
'00493a23    ff152c124000            call dword ptr [0040122c]
Dim var_47 As New frmExperience
'00493a29    8b354cb17200            mov esi, dword ptr [0072b14c]
'00493a2f    8b3e                    mov edi, dword ptr [esi]
'00493a31    83ec10                  sub esp, 10
'00493a34    8bdc                    mov ebx, esp
'00493a36    b90a000000              mov ecx, 0000000a
'00493a3b    890b                    mov dword ptr [ebx], ecx
'00493a3d    894ddc                  mov dword ptr [ebp-24], ecx
'00493a40    8b4dd0                  mov ecx, dword ptr [ebp-30]
'00493a43    894b04                  mov dword ptr [ebx+04], ecx
'00493a46    b804000280              mov eax, 80020004
'00493a4b    894308                  mov dword ptr [ebx+08], eax
'00493a4e    8bd0                    mov edx, eax
'00493a50    8b45d8                  mov eax, dword ptr [ebp-28]
'00493a53    89430c                  mov dword ptr [ebx+0c], eax
'00493a56    8b45dc                  mov eax, dword ptr [ebp-24]
'00493a59    83ec10                  sub esp, 10
'00493a5c    8bcc                    mov ecx, esp
'00493a5e    8901                    mov dword ptr [ecx], eax
'00493a60    8b45e0                  mov eax, dword ptr [ebp-20]
'00493a63    894104                  mov dword ptr [ecx+04], eax
'00493a66    895108                  mov dword ptr [ecx+08], edx
'00493a69    8b55e8                  mov edx, dword ptr [ebp-18]
'00493a6c    56                      push esi
'00493a6d    89510c                  mov dword ptr [ecx+0c], edx
'00493a70    ff97b0020000            call dword ptr [edi+000002b0]
'00493a76    dbe2                    fnclex
'00493a78    85c0                    test ax, ax
'00493a7a    7d12                    jge 493a8e
'00493a7c    68b0020000              push 000002b0
'00493a81    68141d4300              push 00431d14
'00493a86    56                      push esi
'00493a87    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00493a88    ff1580104000            call dword ptr [00401080]
'00493a8e    c745fc00000000          mov dword ptr [ebp-04], 00000000
'00493a95    8b4508                  mov eax, dword ptr [ebp+08]
'00493a98    8b08                    mov ecx, dword ptr [eax]
'00493a9a    50                      push eax
'00493a9b    ff5108                  call dword ptr [ecx+08]
'00493a9e    8b45fc                  mov eax, dword ptr [ebp-04]
'00493aa1    8b4dec                  mov ecx, dword ptr [ebp-14]
'00493aa4    5f                      pop edi
'00493aa5    5e                      pop esi
    'Reference to '__except_list'
'00493aa6    64890d00000000          mov dword ptr fs:[00000000], ecx
'00493aad    5b                      pop ebx
'00493aae    8be5                    mov esp, ebp
'00493ab0    5d                      pop ebp
'00493ab1    c20400                  ret 0004


End Sub


'Event for IDM_LANCEUR_DES
Private Sub IDM_LANCEUR_DES_Click()
'00492e10    55                      push ebp
'00492e11    8bec                    mov ebp, esp
'00492e13    83ec0c                  sub esp, 0c

' *** Reference to "__vbaExceptHandler"
'00492e16    6866724000              push 00407266
'00492e1b    64a100000000            mov ax, word ptr fs:[00000000]
'00492e21    50                      push eax
    'Reference to '__except_list'
'00492e22    64892500000000          mov dword ptr fs:[00000000], esp
'00492e29    83ec30                  sub esp, 30
'00492e2c    53                      push ebx
'00492e2d    56                      push esi
'00492e2e    57                      push edi
'00492e2f    8965f4                  mov dword ptr [ebp-0c], esp
'00492e32    c745f8c0134000          mov dword ptr [ebp-08], 004013c0
'00492e39    8b4508                  mov eax, dword ptr [ebp+08]
'00492e3c    8bc8                    mov ecx, eax
'00492e3e    83e101                  and ecx, 01
'00492e41    894dfc                  mov dword ptr [ebp-04], ecx
'00492e44    83e0fe                  and eax, fffffffe
'00492e47    8b10                    mov edx, dword ptr [eax]
'00492e49    50                      push eax
'00492e4a    894508                  mov dword ptr [ebp+08], eax
'00492e4d    ff5204                  call dword ptr [edx+04]
'00492e50    a138b17200              mov ax, word ptr [0072b138]
'00492e55    85c0                    test ax, ax
'00492e57    7510                    jne 492e69
'00492e59    6838b17200              push 0072b138
'00492e5e    68d49b4000              push 00409bd4

' *** Reference to "__vbaNew2"
'00492e63    ff152c124000            call dword ptr [0040122c]
Dim var_38 As New frmLanceurDes
'00492e69    8b3538b17200            mov esi, dword ptr [0072b138]
'00492e6f    8b3e                    mov edi, dword ptr [esi]
'00492e71    83ec10                  sub esp, 10
'00492e74    8bdc                    mov ebx, esp
'00492e76    b90a000000              mov ecx, 0000000a
'00492e7b    890b                    mov dword ptr [ebx], ecx
'00492e7d    894ddc                  mov dword ptr [ebp-24], ecx
'00492e80    8b4dd0                  mov ecx, dword ptr [ebp-30]
'00492e83    894b04                  mov dword ptr [ebx+04], ecx
'00492e86    b804000280              mov eax, 80020004
'00492e8b    894308                  mov dword ptr [ebx+08], eax
'00492e8e    8bd0                    mov edx, eax
'00492e90    8b45d8                  mov eax, dword ptr [ebp-28]
'00492e93    89430c                  mov dword ptr [ebx+0c], eax
'00492e96    8b45dc                  mov eax, dword ptr [ebp-24]
'00492e99    83ec10                  sub esp, 10
'00492e9c    8bcc                    mov ecx, esp
'00492e9e    8901                    mov dword ptr [ecx], eax
'00492ea0    8b45e0                  mov eax, dword ptr [ebp-20]
'00492ea3    894104                  mov dword ptr [ecx+04], eax
'00492ea6    895108                  mov dword ptr [ecx+08], edx
'00492ea9    8b55e8                  mov edx, dword ptr [ebp-18]
'00492eac    56                      push esi
'00492ead    89510c                  mov dword ptr [ecx+0c], edx
'00492eb0    ff97b0020000            call dword ptr [edi+000002b0]
'00492eb6    dbe2                    fnclex
'00492eb8    85c0                    test ax, ax
'00492eba    7d12                    jge 492ece
'00492ebc    68b0020000              push 000002b0
'00492ec1    68cc1a4300              push 00431acc
'00492ec6    56                      push esi
'00492ec7    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00492ec8    ff1580104000            call dword ptr [00401080]
'00492ece    c745fc00000000          mov dword ptr [ebp-04], 00000000
'00492ed5    8b4508                  mov eax, dword ptr [ebp+08]
'00492ed8    8b08                    mov ecx, dword ptr [eax]
'00492eda    50                      push eax
'00492edb    ff5108                  call dword ptr [ecx+08]
'00492ede    8b45fc                  mov eax, dword ptr [ebp-04]
'00492ee1    8b4dec                  mov ecx, dword ptr [ebp-14]
'00492ee4    5f                      pop edi
'00492ee5    5e                      pop esi
    'Reference to '__except_list'
'00492ee6    64890d00000000          mov dword ptr fs:[00000000], ecx
'00492eed    5b                      pop ebx
'00492eee    8be5                    mov esp, ebp
'00492ef0    5d                      pop ebp
'00492ef1    c20400                  ret 0004


End Sub


'Event for IDM_LIENSITE
Private Sub IDM_LIENSITE_Click()
'00492f00    55                      push ebp
'00492f01    8bec                    mov ebp, esp
'00492f03    83ec0c                  sub esp, 0c

' *** Reference to "__vbaExceptHandler"
'00492f06    6866724000              push 00407266
'00492f0b    64a100000000            mov ax, word ptr fs:[00000000]
'00492f11    50                      push eax
    'Reference to '__except_list'
'00492f12    64892500000000          mov dword ptr fs:[00000000], esp
'00492f19    83ec20                  sub esp, 20
'00492f1c    53                      push ebx
'00492f1d    56                      push esi
'00492f1e    57                      push edi
'00492f1f    8965f4                  mov dword ptr [ebp-0c], esp
'00492f22    c745f8c8134000          mov dword ptr [ebp-08], 004013c8
'00492f29    8b7508                  mov esi, dword ptr [ebp+08]
'00492f2c    8bc6                    mov eax, esi
'00492f2e    83e001                  and eax, 01
'00492f31    8945fc                  mov dword ptr [ebp-04], eax
'00492f34    83e6fe                  and esi, fffffffe
'00492f37    8b0e                    mov ecx, dword ptr [esi]
'00492f39    56                      push esi
'00492f3a    897508                  mov dword ptr [ebp+08], esi
'00492f3d    ff5104                  call dword ptr [ecx+04]
'00492f40    8b16                    mov edx, dword ptr [esi]
'00492f42    33ff                    xor edi, edi
'00492f44    8d45d8                  lea eax, dword ptr [ebp-28]
'00492f47    50                      push eax
'00492f48    56                      push esi
'00492f49    897de8                  mov dword ptr [ebp-18], edi
'00492f4c    897de4                  mov dword ptr [ebp-1c], edi
'00492f4f    897de0                  mov dword ptr [ebp-20], edi
'00492f52    897ddc                  mov dword ptr [ebp-24], edi
'00492f55    897dd8                  mov dword ptr [ebp-28], edi
'00492f58    ff5258                  call dword ptr [edx+58]
'00492f5b    dbe2                    fnclex
'00492f5d    3bc7                    cmp eax, edi
'00492f5f    7d0f                    jge 492f70
'00492f61    6a58                    push 58
'00492f63    684cfd4200              push 0042fd4c
'00492f68    56                      push esi
'00492f69    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00492f6a    ff1580104000            call dword ptr [00401080]

' *** Reference to "__vbaStrToAnsi"
'00492f70    8b359c124000            mov esi, dword ptr [0040129c]
'00492f76    57                      push edi
'00492f77    68cc134300              push 004313cc
'00492f7c    8d4ddc                  lea ecx, dword ptr [ebp-24]
'00492f7f    51                      push ecx
'00492f80    ffd6                    call esi
var_39 = (vbNullChar)
'00492f82    50                      push eax
'00492f83    68cc134300              push 004313cc
'00492f88    8d55e0                  lea edx, dword ptr [ebp-20]
'00492f8b    52                      push edx
'00492f8c    ffd6                    call esi
var_3 = (vbNullChar)
'00492f8e    50                      push eax
'00492f8f    68801b4300              push 00431b80
'00492f94    8d45e4                  lea eax, dword ptr [ebp-1c]
'00492f97    50                      push eax
'00492f98    ffd6                    call esi
var_40 = ("http://bonnarien.dyndns.org")
'00492f9a    50                      push eax
'00492f9b    68701b4300              push 00431b70
'00492fa0    8d4de8                  lea ecx, dword ptr [ebp-18]
'00492fa3    51                      push ecx
'00492fa4    ffd6                    call esi
var_41 = ("open")
'00492fa6    8b55d8                  mov edx, dword ptr [ebp-28]
'00492fa9    50                      push eax
'00492faa    52                      push edx
'00492fab    e8c8d5f9ff              call 430578
var_42 = ShellExecuteA (0, var_41, var_40, var_3, var_39, 0)  '{Function}

' *** Reference to "__vbaSetSystemError"
'00492fb0    ff157c104000            call dword ptr [0040107c]
'#SetAPISystemError()
'00492fb6    8d45dc                  lea eax, dword ptr [ebp-24]
'00492fb9    50                      push eax
'00492fba    8d4de0                  lea ecx, dword ptr [ebp-20]
'00492fbd    51                      push ecx
'00492fbe    8d55e4                  lea edx, dword ptr [ebp-1c]
'00492fc1    52                      push edx
'00492fc2    8d45e8                  lea eax, dword ptr [ebp-18]
'00492fc5    50                      push eax
'00492fc6    6a04                    push 04

' *** Reference to "__vbaFreeStrList"
'00492fc8    ff155c124000            call dword ptr [0040125c]
'#Cleanup( var_41, var_40, var_3, var_39)
'00492fce    83c414                  add esp, 14
'00492fd1    897dfc                  mov dword ptr [ebp-04], edi
'00492fd4    68f82f4900              push 00492ff8
'00492fd9    eb1c                    jmp 492ff7
'00492fdb    8d4ddc                  lea ecx, dword ptr [ebp-24]
'00492fde    51                      push ecx
'00492fdf    8d55e0                  lea edx, dword ptr [ebp-20]
'00492fe2    52                      push edx
'00492fe3    8d45e4                  lea eax, dword ptr [ebp-1c]
'00492fe6    50                      push eax
'00492fe7    8d4de8                  lea ecx, dword ptr [ebp-18]
'00492fea    51                      push ecx
'00492feb    6a04                    push 04

' *** Reference to "__vbaFreeStrList"
'00492fed    ff155c124000            call dword ptr [0040125c]
'#Cleanup( var_41, var_40, var_3, var_39)
'00492ff3    83c414                  add esp, 14
'00492ff6    c3                      ret
'00492ff7    c3                      ret
'00492ff8    8b4508                  mov eax, dword ptr [ebp+08]
'00492ffb    8b10                    mov edx, dword ptr [eax]
'00492ffd    50                      push eax
'00492ffe    ff5208                  call dword ptr [edx+08]
'00493001    8b45fc                  mov eax, dword ptr [ebp-04]
'00493004    8b4dec                  mov ecx, dword ptr [ebp-14]
'00493007    5f                      pop edi
'00493008    5e                      pop esi
    'Reference to '__except_list'
'00493009    64890d00000000          mov dword ptr fs:[00000000], ecx
'00493010    5b                      pop ebx
'00493011    8be5                    mov esp, ebp
'00493013    5d                      pop ebp
'00493014    c20400                  ret 0004


End Sub


'Event for IDM_APROPOS
Private Sub IDM_APROPOS_Click()
'00491f50    55                      push ebp
'00491f51    8bec                    mov ebp, esp
'00491f53    83ec0c                  sub esp, 0c

' *** Reference to "__vbaExceptHandler"
'00491f56    6866724000              push 00407266
'00491f5b    64a100000000            mov ax, word ptr fs:[00000000]
'00491f61    50                      push eax
    'Reference to '__except_list'
'00491f62    64892500000000          mov dword ptr fs:[00000000], esp
'00491f69    83ec30                  sub esp, 30
'00491f6c    53                      push ebx
'00491f6d    56                      push esi
'00491f6e    57                      push edi
'00491f6f    8965f4                  mov dword ptr [ebp-0c], esp
'00491f72    c745f850134000          mov dword ptr [ebp-08], 00401350
'00491f79    8b4508                  mov eax, dword ptr [ebp+08]
'00491f7c    8bc8                    mov ecx, eax
'00491f7e    83e101                  and ecx, 01
'00491f81    894dfc                  mov dword ptr [ebp-04], ecx
'00491f84    83e0fe                  and eax, fffffffe
'00491f87    8b10                    mov edx, dword ptr [eax]
'00491f89    50                      push eax
'00491f8a    894508                  mov dword ptr [ebp+08], eax
'00491f8d    ff5204                  call dword ptr [edx+04]
'00491f90    a170b07200              mov ax, word ptr [0072b070]
'00491f95    85c0                    test ax, ax
'00491f97    7510                    jne 491fa9
'00491f99    6870b07200              push 0072b070
'00491f9e    6830854000              push 00408530

' *** Reference to "__vbaNew2"
'00491fa3    ff152c124000            call dword ptr [0040122c]
Dim var_27 As New frmApropos
'00491fa9    8b3570b07200            mov esi, dword ptr [0072b070]
'00491faf    8b3e                    mov edi, dword ptr [esi]
'00491fb1    83ec10                  sub esp, 10
'00491fb4    8bdc                    mov ebx, esp
'00491fb6    b90a000000              mov ecx, 0000000a
'00491fbb    890b                    mov dword ptr [ebx], ecx
'00491fbd    8b4dd0                  mov ecx, dword ptr [ebp-30]
'00491fc0    894b04                  mov dword ptr [ebx+04], ecx
'00491fc3    b804000280              mov eax, 80020004
'00491fc8    894308                  mov dword ptr [ebx+08], eax
'00491fcb    8b45d8                  mov eax, dword ptr [ebp-28]
'00491fce    89430c                  mov dword ptr [ebx+0c], eax
'00491fd1    83ec10                  sub esp, 10
'00491fd4    8bcc                    mov ecx, esp
'00491fd6    b803000000              mov eax, 00000003
'00491fdb    8901                    mov dword ptr [ecx], eax
'00491fdd    8b45e0                  mov eax, dword ptr [ebp-20]
'00491fe0    894104                  mov dword ptr [ecx+04], eax
'00491fe3    ba01000000              mov edx, 00000001
'00491fe8    895108                  mov dword ptr [ecx+08], edx
'00491feb    8b55e8                  mov edx, dword ptr [ebp-18]
'00491fee    56                      push esi
'00491fef    89510c                  mov dword ptr [ecx+0c], edx
'00491ff2    ff97b0020000            call dword ptr [edi+000002b0]
'00491ff8    dbe2                    fnclex
'00491ffa    85c0                    test ax, ax
'00491ffc    7d12                    jge 492010
'00491ffe    68b0020000              push 000002b0
'00492003    6888094300              push 00430988
'00492008    56                      push esi
'00492009    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0049200a    ff1580104000            call dword ptr [00401080]
'00492010    c745fc00000000          mov dword ptr [ebp-04], 00000000
'00492017    8b4508                  mov eax, dword ptr [ebp+08]
'0049201a    8b08                    mov ecx, dword ptr [eax]
'0049201c    50                      push eax
'0049201d    ff5108                  call dword ptr [ecx+08]
'00492020    8b45fc                  mov eax, dword ptr [ebp-04]
'00492023    8b4dec                  mov ecx, dword ptr [ebp-14]
'00492026    5f                      pop edi
'00492027    5e                      pop esi
    'Reference to '__except_list'
'00492028    64890d00000000          mov dword ptr fs:[00000000], ecx
'0049202f    5b                      pop ebx
'00492030    8be5                    mov esp, ebp
'00492032    5d                      pop ebp
'00492033    c20400                  ret 0004


End Sub


'Event for IDM_DESCRIPTION_DON
Private Sub IDM_DESCRIPTION_DON_Click()
'004922a0    55                      push ebp
'004922a1    8bec                    mov ebp, esp
'004922a3    83ec0c                  sub esp, 0c

' *** Reference to "__vbaExceptHandler"
'004922a6    6866724000              push 00407266
'004922ab    64a100000000            mov ax, word ptr fs:[00000000]
'004922b1    50                      push eax
    'Reference to '__except_list'
'004922b2    64892500000000          mov dword ptr fs:[00000000], esp
'004922b9    83ec30                  sub esp, 30
'004922bc    53                      push ebx
'004922bd    56                      push esi
'004922be    57                      push edi
'004922bf    8965f4                  mov dword ptr [ebp-0c], esp
'004922c2    c745f868134000          mov dword ptr [ebp-08], 00401368
'004922c9    8b4508                  mov eax, dword ptr [ebp+08]
'004922cc    8bc8                    mov ecx, eax
'004922ce    83e101                  and ecx, 01
'004922d1    894dfc                  mov dword ptr [ebp-04], ecx
'004922d4    83e0fe                  and eax, fffffffe
'004922d7    8b10                    mov edx, dword ptr [eax]
'004922d9    50                      push eax
'004922da    894508                  mov dword ptr [ebp+08], eax
'004922dd    ff5204                  call dword ptr [edx+04]
'004922e0    a1c0b07200              mov ax, word ptr [0072b0c0]
'004922e5    85c0                    test ax, ax
'004922e7    7510                    jne 4922f9
'004922e9    68c0b07200              push 0072b0c0
'004922ee    68d0f04000              push 0040f0d0

' *** Reference to "__vbaNew2"
'004922f3    ff152c124000            call dword ptr [0040122c]
Dim var_31 As New frmDescriptifDon
'004922f9    8b35c0b07200            mov esi, dword ptr [0072b0c0]
'004922ff    8b06                    mov eax, dword ptr [esi]
'00492301    6aff                    push ffffffff
'00492303    56                      push esi
'00492304    ff90fc060000            call dword ptr [eax+000006fc]
'0049230a    dbe2                    fnclex
'0049230c    85c0                    test ax, ax
'0049230e    7d12                    jge 492322
'00492310    68fc060000              push 000006fc
'00492315    6888124300              push 00431288
'0049231a    56                      push esi
'0049231b    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0049231c    ff1580104000            call dword ptr [00401080]
'00492322    a1c0b07200              mov ax, word ptr [0072b0c0]
'00492327    85c0                    test ax, ax
'00492329    7510                    jne 49233b
'0049232b    68c0b07200              push 0072b0c0
'00492330    68d0f04000              push 0040f0d0

' *** Reference to "__vbaNew2"
'00492335    ff152c124000            call dword ptr [0040122c]
Set var_31 = New frmDescriptifDon
'0049233b    8b35c0b07200            mov esi, dword ptr [0072b0c0]
'00492341    8b3e                    mov edi, dword ptr [esi]
'00492343    83ec10                  sub esp, 10
'00492346    8bdc                    mov ebx, esp
'00492348    b90a000000              mov ecx, 0000000a
'0049234d    890b                    mov dword ptr [ebx], ecx
'0049234f    894ddc                  mov dword ptr [ebp-24], ecx
'00492352    8b4dd0                  mov ecx, dword ptr [ebp-30]
'00492355    894b04                  mov dword ptr [ebx+04], ecx
'00492358    b804000280              mov eax, 80020004
'0049235d    894308                  mov dword ptr [ebx+08], eax
'00492360    8bd0                    mov edx, eax
'00492362    8b45d8                  mov eax, dword ptr [ebp-28]
'00492365    89430c                  mov dword ptr [ebx+0c], eax
'00492368    8b45dc                  mov eax, dword ptr [ebp-24]
'0049236b    83ec10                  sub esp, 10
'0049236e    8bcc                    mov ecx, esp
'00492370    8901                    mov dword ptr [ecx], eax
'00492372    8b45e0                  mov eax, dword ptr [ebp-20]
'00492375    894104                  mov dword ptr [ecx+04], eax
'00492378    895108                  mov dword ptr [ecx+08], edx
'0049237b    8b55e8                  mov edx, dword ptr [ebp-18]
'0049237e    56                      push esi
'0049237f    89510c                  mov dword ptr [ecx+0c], edx
'00492382    ff97b0020000            call dword ptr [edi+000002b0]
'00492388    dbe2                    fnclex
'0049238a    85c0                    test ax, ax
'0049238c    7d12                    jge 4923a0
'0049238e    68b0020000              push 000002b0
'00492393    6858124300              push 00431258
'00492398    56                      push esi
'00492399    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0049239a    ff1580104000            call dword ptr [00401080]
'004923a0    c745fc00000000          mov dword ptr [ebp-04], 00000000
'004923a7    8b4508                  mov eax, dword ptr [ebp+08]
'004923aa    8b08                    mov ecx, dword ptr [eax]
'004923ac    50                      push eax
'004923ad    ff5108                  call dword ptr [ecx+08]
'004923b0    8b45fc                  mov eax, dword ptr [ebp-04]
'004923b3    8b4dec                  mov ecx, dword ptr [ebp-14]
'004923b6    5f                      pop edi
'004923b7    5e                      pop esi
    'Reference to '__except_list'
'004923b8    64890d00000000          mov dword ptr fs:[00000000], ecx
'004923bf    5b                      pop ebx
'004923c0    8be5                    mov esp, ebp
'004923c2    5d                      pop ebp
'004923c3    c20400                  ret 0004


End Sub


'Event for IDM_ASTUCE
Private Sub IDM_ASTUCE_Click()
'004921b0    55                      push ebp
'004921b1    8bec                    mov ebp, esp
'004921b3    83ec0c                  sub esp, 0c

' *** Reference to "__vbaExceptHandler"
'004921b6    6866724000              push 00407266
'004921bb    64a100000000            mov ax, word ptr fs:[00000000]
'004921c1    50                      push eax
    'Reference to '__except_list'
'004921c2    64892500000000          mov dword ptr fs:[00000000], esp
'004921c9    83ec30                  sub esp, 30
'004921cc    53                      push ebx
'004921cd    56                      push esi
'004921ce    57                      push edi
'004921cf    8965f4                  mov dword ptr [ebp-0c], esp
'004921d2    c745f860134000          mov dword ptr [ebp-08], 00401360
'004921d9    8b4508                  mov eax, dword ptr [ebp+08]
'004921dc    8bc8                    mov ecx, eax
'004921de    83e101                  and ecx, 01
'004921e1    894dfc                  mov dword ptr [ebp-04], ecx
'004921e4    83e0fe                  and eax, fffffffe
'004921e7    8b10                    mov edx, dword ptr [eax]
'004921e9    50                      push eax
'004921ea    894508                  mov dword ptr [ebp+08], eax
'004921ed    ff5204                  call dword ptr [edx+04]
'004921f0    a1acb07200              mov ax, word ptr [0072b0ac]
'004921f5    85c0                    test ax, ax
'004921f7    7510                    jne 492209
'004921f9    68acb07200              push 0072b0ac
'004921fe    68a4914000              push 004091a4

' *** Reference to "__vbaNew2"
'00492203    ff152c124000            call dword ptr [0040122c]
Dim var_30 As New frmAstuce
'00492209    8b35acb07200            mov esi, dword ptr [0072b0ac]
'0049220f    8b3e                    mov edi, dword ptr [esi]
'00492211    83ec10                  sub esp, 10
'00492214    8bdc                    mov ebx, esp
'00492216    b90a000000              mov ecx, 0000000a
'0049221b    890b                    mov dword ptr [ebx], ecx
'0049221d    894ddc                  mov dword ptr [ebp-24], ecx
'00492220    8b4dd0                  mov ecx, dword ptr [ebp-30]
'00492223    894b04                  mov dword ptr [ebx+04], ecx
'00492226    b804000280              mov eax, 80020004
'0049222b    894308                  mov dword ptr [ebx+08], eax
'0049222e    8bd0                    mov edx, eax
'00492230    8b45d8                  mov eax, dword ptr [ebp-28]
'00492233    89430c                  mov dword ptr [ebx+0c], eax
'00492236    8b45dc                  mov eax, dword ptr [ebp-24]
'00492239    83ec10                  sub esp, 10
'0049223c    8bcc                    mov ecx, esp
'0049223e    8901                    mov dword ptr [ecx], eax
'00492240    8b45e0                  mov eax, dword ptr [ebp-20]
'00492243    894104                  mov dword ptr [ecx+04], eax
'00492246    895108                  mov dword ptr [ecx+08], edx
'00492249    8b55e8                  mov edx, dword ptr [ebp-18]
'0049224c    56                      push esi
'0049224d    89510c                  mov dword ptr [ecx+0c], edx
'00492250    ff97b0020000            call dword ptr [edi+000002b0]
'00492256    dbe2                    fnclex
'00492258    85c0                    test ax, ax
'0049225a    7d12                    jge 49226e
'0049225c    68b0020000              push 000002b0
'00492261    68c0114300              push 004311c0
'00492266    56                      push esi
'00492267    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00492268    ff1580104000            call dword ptr [00401080]
'0049226e    c745fc00000000          mov dword ptr [ebp-04], 00000000
'00492275    8b4508                  mov eax, dword ptr [ebp+08]
'00492278    8b08                    mov ecx, dword ptr [eax]
'0049227a    50                      push eax
'0049227b    ff5108                  call dword ptr [ecx+08]
'0049227e    8b45fc                  mov eax, dword ptr [ebp-04]
'00492281    8b4dec                  mov ecx, dword ptr [ebp-14]
'00492284    5f                      pop edi
'00492285    5e                      pop esi
    'Reference to '__except_list'
'00492286    64890d00000000          mov dword ptr fs:[00000000], ecx
'0049228d    5b                      pop ebx
'0049228e    8be5                    mov esp, ebp
'00492290    5d                      pop ebp
'00492291    c20400                  ret 0004


End Sub


'Event for IDM_GESTION_PERSO
Private Sub IDM_GESTION_PERSO_Click()
'004927e0    55                      push ebp
'004927e1    8bec                    mov ebp, esp
'004927e3    83ec14                  sub esp, 14

' *** Reference to "__vbaExceptHandler"
'004927e6    6866724000              push 00407266
'004927eb    64a100000000            mov ax, word ptr fs:[00000000]
'004927f1    50                      push eax
    'Reference to '__except_list'
'004927f2    64892500000000          mov dword ptr fs:[00000000], esp
'004927f9    81ec2c010000            sub esp, 0000012c
'004927ff    53                      push ebx
'00492800    56                      push esi
'00492801    57                      push edi
'00492802    8965ec                  mov dword ptr [ebp-14], esp
'00492805    c745f090134000          mov dword ptr [ebp-10], 00401390
'0049280c    8b4508                  mov eax, dword ptr [ebp+08]
'0049280f    8bc8                    mov ecx, eax
'00492811    83e101                  and ecx, 01
'00492814    894df4                  mov dword ptr [ebp-0c], ecx
'00492817    83e0fe                  and eax, fffffffe
'0049281a    894508                  mov dword ptr [ebp+08], eax
'0049281d    33ff                    xor edi, edi
'0049281f    897df8                  mov dword ptr [ebp-08], edi
'00492822    8b10                    mov edx, dword ptr [eax]
'00492824    50                      push eax
'00492825    ff5204                  call dword ptr [edx+04]
'00492828    897de0                  mov dword ptr [ebp-20], edi
'0049282b    897ddc                  mov dword ptr [ebp-24], edi
'0049282e    897dd8                  mov dword ptr [ebp-28], edi
'00492831    897dd4                  mov dword ptr [ebp-2c], edi
'00492834    897dd0                  mov dword ptr [ebp-30], edi
'00492837    897dcc                  mov dword ptr [ebp-34], edi
'0049283a    897dc8                  mov dword ptr [ebp-38], edi
'0049283d    897dc4                  mov dword ptr [ebp-3c], edi
'00492840    897dc0                  mov dword ptr [ebp-40], edi
'00492843    897db0                  mov dword ptr [ebp-50], edi
'00492846    897da0                  mov dword ptr [ebp-60], edi
'00492849    897d90                  mov dword ptr [ebp-70], edi
'0049284c    897d80                  mov dword ptr [ebp-80], edi
'0049284f    89bd70ffffff            mov dword ptr [ebp+ffffff70], edi
'00492855    89bd60ffffff            mov dword ptr [ebp+ffffff60], edi
'0049285b    89bd50ffffff            mov dword ptr [ebp+ffffff50], edi
'00492861    89bd40ffffff            mov dword ptr [ebp+ffffff40], edi
'00492867    89bd30ffffff            mov dword ptr [ebp+ffffff30], edi
'0049286d    89bd20ffffff            mov dword ptr [ebp+ffffff20], edi
'00492873    89bd10ffffff            mov dword ptr [ebp+ffffff10], edi
'00492879    89bd00ffffff            mov dword ptr [ebp+ffffff00], edi
'0049287f    89bdf0feffff            mov dword ptr [ebp+fffffef0], edi
'00492885    89bde0feffff            mov dword ptr [ebp+fffffee0], edi
'0049288b    89bddcfeffff            mov dword ptr [ebp+fffffedc], edi
'00492891    66393d28b07200          cmp word ptr [0072b028], di
'00492898    7508                    jne 4928a2
'0049289a    6a01                    push 01

' *** Reference to "__vbaOnError"
'0049289c    ff15b0104000            call dword ptr [004010b0]
On Error Goto handler_0
'004928a2    393d84b07200            cmp dword ptr [0072b084], edi
'004928a8    7510                    jne 4928ba
'004928aa    6884b07200              push 0072b084
'004928af    68f8894100              push 004189f8

' *** Reference to "__vbaNew2"
'004928b4    ff152c124000            call dword ptr [0040122c]
Dim var_28 As New frmGestPerso
'004928ba    8b3584b07200            mov esi, dword ptr [0072b084]
'004928c0    b804000280              mov eax, 80020004
'004928c5    898518ffffff            mov dword ptr [ebp+ffffff18], eax
'004928cb    b90a000000              mov ecx, 0000000a
'004928d0    898d10ffffff            mov dword ptr [ebp+ffffff10], ecx
'004928d6    898528ffffff            mov dword ptr [ebp+ffffff28], eax
'004928dc    898d20ffffff            mov dword ptr [ebp+ffffff20], ecx
'004928e2    8b16                    mov edx, dword ptr [esi]
'004928e4    83ec10                  sub esp, 10
'004928e7    8bdc                    mov ebx, esp
'004928e9    890b                    mov dword ptr [ebx], ecx
'004928eb    8b8d14ffffff            mov ecx, dword ptr [ebp+ffffff14]
'004928f1    894b04                  mov dword ptr [ebx+04], ecx
'004928f4    894308                  mov dword ptr [ebx+08], eax
'004928f7    8b851cffffff            mov eax, dword ptr [ebp+ffffff1c]
'004928fd    89430c                  mov dword ptr [ebx+0c], eax
'00492900    83ec10                  sub esp, 10
'00492903    8bcc                    mov ecx, esp
'00492905    8b8520ffffff            mov eax, dword ptr [ebp+ffffff20]
'0049290b    8901                    mov dword ptr [ecx], eax
'0049290d    8b8524ffffff            mov eax, dword ptr [ebp+ffffff24]
'00492913    894104                  mov dword ptr [ecx+04], eax
'00492916    8b8528ffffff            mov eax, dword ptr [ebp+ffffff28]
'0049291c    894108                  mov dword ptr [ecx+08], eax
'0049291f    8b852cffffff            mov eax, dword ptr [ebp+ffffff2c]
'00492925    89410c                  mov dword ptr [ecx+0c], eax
'00492928    56                      push esi
'00492929    ff92b0020000            call dword ptr [edx+000002b0]
'0049292f    dbe2                    fnclex
'00492931    3bc7                    cmp eax, edi
'00492933    7d12                    jge 492947

If (0 < 0) Then
'00492935    68b0020000              push 000002b0
'0049293a    6850014300              push 00430150
'0049293f    56                      push esi
'00492940    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00492941    ff1580104000            call dword ptr [00401080]
    
End If

' *** Reference to "__vbaExitProc"
'00492947    ff15a0104000            call dword ptr [004010a0]
'0049294d    68012d4900              push 00492d01
'00492952    e9a9030000              jmp 492d00

' *** Reference to "rtcErrObj"
'00492957    8b1d6c124000            mov ebx, dword ptr [0040126c]
'0049295d    ffd3                    call ebx
'0049295f    50                      push eax
'00492960    8d45c4                  lea eax, dword ptr [ebp-3c]
'00492963    50                      push eax

' *** Reference to "__vbaObjSet"
'00492964    8b3db4104000            mov edi, dword ptr [004010b4]
'0049296a    ffd7                    call edi
Set var_9 = Err
'0049296c    8bf0                    mov esi, eax
'0049296e    8b0e                    mov ecx, dword ptr [esi]
'00492970    8d55e0                  lea edx, dword ptr [ebp-20]
'00492973    52                      push edx
'00492974    56                      push esi
'00492975    ff512c                  call dword ptr [ecx+2c]
var_3 = var_9.Description
'00492978    dbe2                    fnclex
'0049297a    85c0                    test ax, ax
'0049297c    7d0f                    jge 49298d
'0049297e    6a2c                    push 2c
'00492980    685c084300              push 0043085c
'00492985    56                      push esi
'00492986    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00492987    ff1580104000            call dword ptr [00401080]
'0049298d    ffd3                    call ebx
'0049298f    50                      push eax
'00492990    8d45c0                  lea eax, dword ptr [ebp-40]
'00492993    50                      push eax
'00492994    ffd7                    call edi
Set var_5 = Err
'00492996    8bf0                    mov esi, eax
'00492998    8b0e                    mov ecx, dword ptr [esi]
'0049299a    8d95dcfeffff            lea edx, dword ptr [ebp+fffffedc]
'004929a0    52                      push edx
'004929a1    56                      push esi
'004929a2    ff511c                  call dword ptr [ecx+1c]
var_10 = var_5.Number
'004929a5    dbe2                    fnclex
'004929a7    85c0                    test ax, ax
'004929a9    7d0f                    jge 4929ba
'004929ab    6a1c                    push 1c
'004929ad    685c084300              push 0043085c
'004929b2    56                      push esi
'004929b3    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'004929b4    ff1580104000            call dword ptr [00401080]
'004929ba    b804000280              mov eax, 80020004
'004929bf    894598                  mov dword ptr [ebp-68], eax
'004929c2    b90a000000              mov ecx, 0000000a
'004929c7    894d90                  mov dword ptr [ebp-70], ecx
'004929ca    8945a8                  mov dword ptr [ebp-58], eax
'004929cd    894da0                  mov dword ptr [ebp-60], ecx
'004929d0    c78528ffffff24b07200    mov dword ptr [ebp+ffffff28], 0072b024
'004929da    c78520ffffff08400000    mov dword ptr [ebp+ffffff20], 00004008
'004929e4    6814084300              push 00430814
'004929e9    8b45e0                  mov eax, dword ptr [ebp-20]
'004929ec    50                      push eax

' *** Reference to "__vbaStrCat"
'004929ed    8b3570104000            mov esi, dword ptr [00401070]
'004929f3    ffd6                    call esi
var_36 = ("L'erreur suivante s'est produite : ") & (var_3)
'004929f5    8bd0                    mov edx, eax
'004929f7    8d4ddc                  lea ecx, dword ptr [ebp-24]

' *** Reference to "__vbaStrMove"
'004929fa    8b3dd0124000            mov edi, dword ptr [004012d0]
'00492a00    ffd7                    call edi
'00492a02    50                      push eax
'00492a03    6870084300              push 00430870
'00492a08    ffd6                    call esi
var_12 = (var_36) & (vbCrLf)
'00492a0a    8bd0                    mov edx, eax
'00492a0c    8d4dd8                  lea ecx, dword ptr [ebp-28]
'00492a0f    ffd7                    call edi
'00492a11    50                      push eax
'00492a12    6870084300              push 00430870
'00492a17    ffd6                    call esi
var_13 = (var_12) & (vbCrLf)
'00492a19    8bd0                    mov edx, eax
'00492a1b    8d4dd4                  lea ecx, dword ptr [ebp-2c]
'00492a1e    ffd7                    call edi
'00492a20    50                      push eax
'00492a21    6880084300              push 00430880
'00492a26    ffd6                    call esi
var_14 = (var_13) & (" numéro d'erreur (")
'00492a28    8bd0                    mov edx, eax
'00492a2a    8d4dd0                  lea ecx, dword ptr [ebp-30]
'00492a2d    ffd7                    call edi
'00492a2f    50                      push eax
'00492a30    8b8ddcfeffff            mov ecx, dword ptr [ebp+fffffedc]
'00492a36    51                      push ecx

' *** Reference to "__vbaStrI4"
'00492a37    ff1520104000            call dword ptr [00401020]
'00492a3d    8bd0                    mov edx, eax
'00492a3f    8d4dcc                  lea ecx, dword ptr [ebp-34]
'00492a42    ffd7                    call edi
'00492a44    50                      push eax
'00492a45    ffd6                    call esi
var_15 = (var_14) & (CStr(var_10))
'00492a47    8bd0                    mov edx, eax
'00492a49    8d4dc8                  lea ecx, dword ptr [ebp-38]
'00492a4c    ffd7                    call edi
'00492a4e    50                      push eax
'00492a4f    68ac084300              push 004308ac
'00492a54    ffd6                    call esi
var_16 = (var_15) & (" )")
'00492a56    8945b8                  mov dword ptr [ebp-48], eax
'00492a59    bf08000000              mov edi, 00000008
'00492a5e    897db0                  mov dword ptr [ebp-50], edi
'00492a61    8d5590                  lea edx, dword ptr [ebp-70]
'00492a64    52                      push edx
'00492a65    8d45a0                  lea eax, dword ptr [ebp-60]
'00492a68    50                      push eax
'00492a69    8d8d20ffffff            lea ecx, dword ptr [ebp+ffffff20]
'00492a6f    51                      push ecx
'00492a70    6a10                    push 10
'00492a72    8d55b0                  lea edx, dword ptr [ebp-50]
'00492a75    52                      push edx

' *** Reference to "rtcMsgBox"
'00492a76    ff15b8104000            call dword ptr [004010b8]
var_17 = MsgBox(var_16, 16, vbNullString)
'00492a7c    8d45c8                  lea eax, dword ptr [ebp-38]
'00492a7f    50                      push eax
'00492a80    8d4dcc                  lea ecx, dword ptr [ebp-34]
'00492a83    51                      push ecx
'00492a84    8d55d0                  lea edx, dword ptr [ebp-30]
'00492a87    52                      push edx
'00492a88    8d45d4                  lea eax, dword ptr [ebp-2c]
'00492a8b    50                      push eax
'00492a8c    8d4dd8                  lea ecx, dword ptr [ebp-28]
'00492a8f    51                      push ecx
'00492a90    8d55dc                  lea edx, dword ptr [ebp-24]
'00492a93    52                      push edx
'00492a94    8d45e0                  lea eax, dword ptr [ebp-20]
'00492a97    50                      push eax
'00492a98    6a07                    push 07

' *** Reference to "__vbaFreeStrList"
'00492a9a    ff155c124000            call dword ptr [0040125c]
'#Cleanup( 0, -4512, -4516, -4520, -4524, -4528, -4532)
'00492aa0    8d4dc0                  lea ecx, dword ptr [ebp-40]
'00492aa3    51                      push ecx
'00492aa4    8d55c4                  lea edx, dword ptr [ebp-3c]
'00492aa7    52                      push edx
'00492aa8    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'00492aaa    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_9, var_5)
'00492ab0    8d4590                  lea eax, dword ptr [ebp-70]
'00492ab3    50                      push eax
'00492ab4    8d4da0                  lea ecx, dword ptr [ebp-60]
'00492ab7    51                      push ecx
'00492ab8    8d55b0                  lea edx, dword ptr [ebp-50]
'00492abb    52                      push edx
'00492abc    6a03                    push 03

' *** Reference to "__vbaFreeVarList"
'00492abe    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_6, var_7, var_8)
'00492ac4    83c43c                  add esp, 3c
'00492ac7    8d45b0                  lea eax, dword ptr [ebp-50]
'00492aca    50                      push eax

' *** Reference to "rtcGetPresentDate"
'00492acb    ff15f4124000            call dword ptr [004012f4]
var_16 = Now()
'00492ad1    c78528ffffffb8084300    mov dword ptr [ebp+ffffff28], 004308b8
'00492adb    89bd20ffffff            mov dword ptr [ebp+ffffff20], edi
'00492ae1    8d9520ffffff            lea edx, dword ptr [ebp+ffffff20]
'00492ae7    8d4da0                  lea ecx, dword ptr [ebp-60]

' *** Reference to "__vbaVarDup"
'00492aea    ff1598124000            call dword ptr [00401298]
var_7 = ("dd/MM/yyyy hh:mm:ss")
'00492af0    6a01                    push 01
'00492af2    6a01                    push 01
'00492af4    8d4da0                  lea ecx, dword ptr [ebp-60]
'00492af7    51                      push ecx
'00492af8    8d55b0                  lea edx, dword ptr [ebp-50]
'00492afb    52                      push edx
'00492afc    8d4590                  lea eax, dword ptr [ebp-70]
'00492aff    50                      push eax

' *** Reference to "rtcVarFromFormatVar"
'00492b00    ff1574104000            call dword ptr [00401074]
'00492b06    ffd3                    call ebx
'00492b08    50                      push eax
'00492b09    8d4dc4                  lea ecx, dword ptr [ebp-3c]
'00492b0c    51                      push ecx

' *** Reference to "__vbaObjSet"
'00492b0d    ff15b4104000            call dword ptr [004010b4]
Set var_9 = Err
'00492b13    8bf0                    mov esi, eax
'00492b15    8b16                    mov edx, dword ptr [esi]
'00492b17    8d85dcfeffff            lea eax, dword ptr [ebp+fffffedc]
'00492b1d    50                      push eax
'00492b1e    56                      push esi
'00492b1f    ff521c                  call dword ptr [edx+1c]
var_10 = var_9.Number
'00492b22    dbe2                    fnclex
'00492b24    85c0                    test ax, ax
'00492b26    7d0f                    jge 492b37
'00492b28    6a1c                    push 1c
'00492b2a    685c084300              push 0043085c
'00492b2f    56                      push esi
'00492b30    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00492b31    ff1580104000            call dword ptr [00401080]
'00492b37    ffd3                    call ebx
'00492b39    50                      push eax
'00492b3a    8d4dc0                  lea ecx, dword ptr [ebp-40]
'00492b3d    51                      push ecx

' *** Reference to "__vbaObjSet"
'00492b3e    ff15b4104000            call dword ptr [004010b4]
Set var_5 = Err
'00492b44    8bf0                    mov esi, eax
'00492b46    8b16                    mov edx, dword ptr [esi]
'00492b48    8d45e0                  lea eax, dword ptr [ebp-20]
'00492b4b    50                      push eax
'00492b4c    56                      push esi
'00492b4d    ff522c                  call dword ptr [edx+2c]
var_3 = var_5.Description
'00492b50    dbe2                    fnclex
'00492b52    85c0                    test ax, ax
'00492b54    7d0f                    jge 492b65
'00492b56    6a2c                    push 2c
'00492b58    685c084300              push 0043085c
'00492b5d    56                      push esi
'00492b5e    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00492b5f    ff1580104000            call dword ptr [00401080]
'00492b65    c78518ffffffe4084300    mov dword ptr [ebp+ffffff18], 004308e4
'00492b6f    89bd10ffffff            mov dword ptr [ebp+ffffff10], edi
'00492b75    8b8ddcfeffff            mov ecx, dword ptr [ebp+fffffedc]
'00492b7b    898d08ffffff            mov dword ptr [ebp+ffffff08], ecx
'00492b81    c78500ffffff03000000    mov dword ptr [ebp+ffffff00], 00000003
'00492b8b    c785f8feffff08094300    mov dword ptr [ebp+fffffef8], 00430908
'00492b95    89bdf0feffff            mov dword ptr [ebp+fffffef0], edi
'00492b9b    8b45e0                  mov eax, dword ptr [ebp-20]
'00492b9e    c745e000000000          mov dword ptr [ebp-20], 00000000
'00492ba5    898558ffffff            mov dword ptr [ebp+ffffff58], eax
'00492bab    89bd50ffffff            mov dword ptr [ebp+ffffff50], edi
'00492bb1    c785e8feffff38194300    mov dword ptr [ebp+fffffee8], 00431938
'00492bbb    89bde0feffff            mov dword ptr [ebp+fffffee0], edi
'00492bc1    8d5590                  lea edx, dword ptr [ebp-70]
'00492bc4    52                      push edx
'00492bc5    8d8510ffffff            lea eax, dword ptr [ebp+ffffff10]
'00492bcb    50                      push eax
'00492bcc    8d4d80                  lea ecx, dword ptr [ebp-80]
'00492bcf    51                      push ecx

' *** Reference to "__vbaVarCat"
'00492bd0    8b3508124000            mov esi, dword ptr [00401208]
'00492bd6    ffd6                    call esi
'00492bd8    50                      push eax
'00492bd9    8d9500ffffff            lea edx, dword ptr [ebp+ffffff00]
'00492bdf    52                      push edx
'00492be0    8d8570ffffff            lea eax, dword ptr [ebp+ffffff70]
'00492be6    50                      push eax
'00492be7    ffd6                    call esi
'00492be9    50                      push eax
'00492bea    8d8df0feffff            lea ecx, dword ptr [ebp+fffffef0]
'00492bf0    51                      push ecx
'00492bf1    8d9560ffffff            lea edx, dword ptr [ebp+ffffff60]
'00492bf7    52                      push edx
'00492bf8    ffd6                    call esi
'00492bfa    50                      push eax
'00492bfb    8d8550ffffff            lea eax, dword ptr [ebp+ffffff50]
'00492c01    50                      push eax
'00492c02    8d8d40ffffff            lea ecx, dword ptr [ebp+ffffff40]
'00492c08    51                      push ecx
'00492c09    ffd6                    call esi
'00492c0b    50                      push eax
'00492c0c    8d95e0feffff            lea edx, dword ptr [ebp+fffffee0]
'00492c12    52                      push edx
'00492c13    8d8530ffffff            lea eax, dword ptr [ebp+ffffff30]
'00492c19    50                      push eax
'00492c1a    ffd6                    call esi
'00492c1c    50                      push eax
'00492c1d    33c9                    xor ecx, ecx
'00492c1f    668b0d2ab07200          mov cx, word ptr [0072b02a]
'00492c26    51                      push ecx
'00492c27    6884094300              push 00430984

' *** Reference to "__vbaPrintFile"
'00492c2c    ff15b8114000            call dword ptr [004011b8]
Print #0, 
'00492c32    8d55c0                  lea edx, dword ptr [ebp-40]
'00492c35    52                      push edx
'00492c36    8d45c4                  lea eax, dword ptr [ebp-3c]
'00492c39    50                      push eax
'00492c3a    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'00492c3c    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_9, var_5)
'00492c42    8d8d30ffffff            lea ecx, dword ptr [ebp+ffffff30]
'00492c48    51                      push ecx
'00492c49    8d9540ffffff            lea edx, dword ptr [ebp+ffffff40]
'00492c4f    52                      push edx
'00492c50    8d8550ffffff            lea eax, dword ptr [ebp+ffffff50]
'00492c56    50                      push eax
'00492c57    8d8d60ffffff            lea ecx, dword ptr [ebp+ffffff60]
'00492c5d    51                      push ecx
'00492c5e    8d9570ffffff            lea edx, dword ptr [ebp+ffffff70]
'00492c64    52                      push edx
'00492c65    8d4580                  lea eax, dword ptr [ebp-80]
'00492c68    50                      push eax
'00492c69    8d4d90                  lea ecx, dword ptr [ebp-70]
'00492c6c    51                      push ecx
'00492c6d    8d55a0                  lea edx, dword ptr [ebp-60]
'00492c70    52                      push edx
'00492c71    8d45b0                  lea eax, dword ptr [ebp-50]
'00492c74    50                      push eax
'00492c75    6a09                    push 09

' *** Reference to "__vbaFreeVarList"
'00492c77    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_6, var_7, var_8, var_18, var_19, var_20, var_21, var_22, var_23)
'00492c7d    83c440                  add esp, 40
'00492c80    6a00                    push 00

' *** Reference to "__vbaResume"
'00492c82    ff1568104000            call dword ptr [00401068]
'00492c88    e9bafcffff              jmp 492947
Resume handler_492947
'00492c8d    8d4dc8                  lea ecx, dword ptr [ebp-38]
'00492c90    51                      push ecx
'00492c91    8d55cc                  lea edx, dword ptr [ebp-34]
'00492c94    52                      push edx
'00492c95    8d45d0                  lea eax, dword ptr [ebp-30]
'00492c98    50                      push eax
'00492c99    8d4dd4                  lea ecx, dword ptr [ebp-2c]
'00492c9c    51                      push ecx
'00492c9d    8d55d8                  lea edx, dword ptr [ebp-28]
'00492ca0    52                      push edx
'00492ca1    8d45dc                  lea eax, dword ptr [ebp-24]
'00492ca4    50                      push eax
'00492ca5    8d4de0                  lea ecx, dword ptr [ebp-20]
'00492ca8    51                      push ecx
'00492ca9    6a07                    push 07

' *** Reference to "__vbaFreeStrList"
'00492cab    ff155c124000            call dword ptr [0040125c]
'#Cleanup( var_3, -4512, -4516, -4520, -4524, -4528, -4532)
'00492cb1    8d55c0                  lea edx, dword ptr [ebp-40]
'00492cb4    52                      push edx
'00492cb5    8d45c4                  lea eax, dword ptr [ebp-3c]
'00492cb8    50                      push eax
'00492cb9    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'00492cbb    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_9, var_5)
'00492cc1    8d8d30ffffff            lea ecx, dword ptr [ebp+ffffff30]
'00492cc7    51                      push ecx
'00492cc8    8d9540ffffff            lea edx, dword ptr [ebp+ffffff40]
'00492cce    52                      push edx
'00492ccf    8d8550ffffff            lea eax, dword ptr [ebp+ffffff50]
'00492cd5    50                      push eax
'00492cd6    8d8d60ffffff            lea ecx, dword ptr [ebp+ffffff60]
'00492cdc    51                      push ecx
'00492cdd    8d9570ffffff            lea edx, dword ptr [ebp+ffffff70]
'00492ce3    52                      push edx
'00492ce4    8d4580                  lea eax, dword ptr [ebp-80]
'00492ce7    50                      push eax
'00492ce8    8d4d90                  lea ecx, dword ptr [ebp-70]
'00492ceb    51                      push ecx
'00492cec    8d55a0                  lea edx, dword ptr [ebp-60]
'00492cef    52                      push edx
'00492cf0    8d45b0                  lea eax, dword ptr [ebp-50]
'00492cf3    50                      push eax
'00492cf4    6a09                    push 09

' *** Reference to "__vbaFreeVarList"
'00492cf6    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_6, var_7, var_8, var_18, var_19, var_20, var_21, var_22, var_23)
'00492cfc    83c454                  add esp, 54
'00492cff    c3                      ret
'00492d00    c3                      ret
'00492d01    8b4508                  mov eax, dword ptr [ebp+08]
'00492d04    8b08                    mov ecx, dword ptr [eax]
'00492d06    50                      push eax
'00492d07    ff5108                  call dword ptr [ecx+08]
'00492d0a    8b45f4                  mov eax, dword ptr [ebp-0c]
'00492d0d    8b4de4                  mov ecx, dword ptr [ebp-1c]
    'Reference to '__except_list'
'00492d10    64890d00000000          mov dword ptr fs:[00000000], ecx
'00492d17    5f                      pop edi
'00492d18    5e                      pop esi
'00492d19    5b                      pop ebx
'00492d1a    8be5                    mov esp, ebp
'00492d1c    5d                      pop ebp
'00492d1d    c20400                  ret 0004


End Sub


'Event for MDIForm
Private Sub MDIForm_Load()
'00494570    55                      push ebp
'00494571    8bec                    mov ebp, esp
'00494573    83ec14                  sub esp, 14

' *** Reference to "__vbaExceptHandler"
'00494576    6866724000              push 00407266
'0049457b    64a100000000            mov ax, word ptr fs:[00000000]
'00494581    50                      push eax
    'Reference to '__except_list'
'00494582    64892500000000          mov dword ptr fs:[00000000], esp
'00494589    81ec74010000            sub esp, 00000174
'0049458f    53                      push ebx
'00494590    56                      push esi
'00494591    57                      push edi
'00494592    8965ec                  mov dword ptr [ebp-14], esp
'00494595    c745f0c8144000          mov dword ptr [ebp-10], 004014c8
'0049459c    8b5d08                  mov ebx, dword ptr [ebp+08]
'0049459f    8bc3                    mov eax, ebx
'004945a1    83e001                  and eax, 01
'004945a4    8945f4                  mov dword ptr [ebp-0c], eax
'004945a7    83e3fe                  and ebx, fffffffe
'004945aa    895d08                  mov dword ptr [ebp+08], ebx
'004945ad    33ff                    xor edi, edi
'004945af    897df8                  mov dword ptr [ebp-08], edi
'004945b2    8b0b                    mov ecx, dword ptr [ebx]
'004945b4    53                      push ebx
'004945b5    ff5104                  call dword ptr [ecx+04]
'004945b8    897dd4                  mov dword ptr [ebp-2c], edi
'004945bb    897dd0                  mov dword ptr [ebp-30], edi
'004945be    897dcc                  mov dword ptr [ebp-34], edi
'004945c1    897dc8                  mov dword ptr [ebp-38], edi
'004945c4    897dc4                  mov dword ptr [ebp-3c], edi
'004945c7    897dc0                  mov dword ptr [ebp-40], edi
'004945ca    897dbc                  mov dword ptr [ebp-44], edi
'004945cd    897db8                  mov dword ptr [ebp-48], edi
'004945d0    897db4                  mov dword ptr [ebp-4c], edi
'004945d3    897db0                  mov dword ptr [ebp-50], edi
'004945d6    897da0                  mov dword ptr [ebp-60], edi
'004945d9    897d90                  mov dword ptr [ebp-70], edi
'004945dc    897d80                  mov dword ptr [ebp-80], edi
'004945df    89bd70ffffff            mov dword ptr [ebp+ffffff70], edi
'004945e5    89bd60ffffff            mov dword ptr [ebp+ffffff60], edi
'004945eb    89bd50ffffff            mov dword ptr [ebp+ffffff50], edi
'004945f1    89bd40ffffff            mov dword ptr [ebp+ffffff40], edi
'004945f7    89bd30ffffff            mov dword ptr [ebp+ffffff30], edi
'004945fd    89bd20ffffff            mov dword ptr [ebp+ffffff20], edi
'00494603    89bd10ffffff            mov dword ptr [ebp+ffffff10], edi
'00494609    89bd00ffffff            mov dword ptr [ebp+ffffff00], edi
'0049460f    89bdf0feffff            mov dword ptr [ebp+fffffef0], edi
'00494615    89bde0feffff            mov dword ptr [ebp+fffffee0], edi
'0049461b    89bdd0feffff            mov dword ptr [ebp+fffffed0], edi
'00494621    89bdccfeffff            mov dword ptr [ebp+fffffecc], edi
'00494627    89bdc8feffff            mov dword ptr [ebp+fffffec8], edi
'0049462d    66393d28b07200          cmp word ptr [0072b028], di
'00494634    7508                    jne 49463e
'00494636    6a01                    push 01

' *** Reference to "__vbaOnError"
'00494638    ff15b0104000            call dword ptr [004010b0]
On Error Goto handler_0
'0049463e    393d10b07200            cmp dword ptr [0072b010], edi
'00494644    7510                    jne 494656
'00494646    6810b07200              push 0072b010
'0049464b    68b0e54000              push 0040e5b0

' *** Reference to "__vbaNew2"
'00494650    ff152c124000            call dword ptr [0040122c]
Dim var_60 As New frmMain
'00494656    8b3510b07200            mov esi, dword ptr [0072b010]
'0049465c    8b06                    mov eax, dword ptr [esi]
'0049465e    6805000080              push 80000005
'00494663    56                      push esi
'00494664    ff5064                  call dword ptr [eax+64]
'00494667    dbe2                    fnclex
'00494669    3bc7                    cmp eax, edi
'0049466b    7d0f                    jge 49467c

If (frmMain < 0) Then
'0049466d    6a64                    push 64
'0049466f    684cfd4200              push 0042fd4c
'00494674    56                      push esi
'00494675    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00494676    ff1580104000            call dword ptr [00401080]
    
End If
'0049467c    393d10b07200            cmp dword ptr [0072b010], edi
'00494682    7510                    jne 494694

If (var_60 = 0) Then
'00494684    6810b07200              push 0072b010
'00494689    68b0e54000              push 0040e5b0

' *** Reference to "__vbaNew2"
'0049468e    ff152c124000            call dword ptr [0040122c]
    Set var_60 = New frmMain
End If
'00494694    8b0d10b07200            mov ecx, dword ptr [0072b010]
'0049469a    898dacfeffff            mov dword ptr [ebp+fffffeac], ecx
'004946a0    393d24be7200            cmp dword ptr [0072be24], edi
'004946a6    7510                    jne 4946b8
'004946a8    6824be7200              push 0072be24
'004946ad    6870004300              push 00430070

' *** Reference to "__vbaNew2"
'004946b2    ff152c124000            call dword ptr [0040122c]
Dim var_2 As New Global
'004946b8    8b3d24be7200            mov edi, dword ptr [0072be24]
'004946be    b904000280              mov ecx, 80020004
'004946c3    898de8feffff            mov dword ptr [ebp+fffffee8], ecx
'004946c9    b80a000000              mov eax, 0000000a
'004946ce    8985e0feffff            mov dword ptr [ebp+fffffee0], eax
'004946d4    898df8feffff            mov dword ptr [ebp+fffffef8], ecx
'004946da    8985f0feffff            mov dword ptr [ebp+fffffef0], eax
'004946e0    898d08ffffff            mov dword ptr [ebp+ffffff08], ecx
'004946e6    898500ffffff            mov dword ptr [ebp+ffffff00], eax
'004946ec    898d18ffffff            mov dword ptr [ebp+ffffff18], ecx
'004946f2    898510ffffff            mov dword ptr [ebp+ffffff10], eax
'004946f8    a124be7200              mov ax, word ptr [0072be24]
'004946fd    85c0                    test ax, ax
'004946ff    7510                    jne 494711
'00494701    6824be7200              push 0072be24
'00494706    6870004300              push 00430070

' *** Reference to "__vbaNew2"
'0049470b    ff152c124000            call dword ptr [0040122c]
Set var_2 = New Global
'00494711    8b3524be7200            mov esi, dword ptr [0072be24]
'00494717    8b16                    mov edx, dword ptr [esi]
'00494719    8d45b8                  lea eax, dword ptr [ebp-48]
'0049471c    50                      push eax
'0049471d    56                      push esi
'0049471e    ff5214                  call dword ptr [edx+14]
Set var_61 = var_2.App
'00494721    dbe2                    fnclex
'00494723    85c0                    test ax, ax
'00494725    7d0f                    jge 494736
'00494727    6a14                    push 14
'00494729    6860004300              push 00430060
'0049472e    56                      push esi
'0049472f    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00494730    ff1580104000            call dword ptr [00401080]
'00494736    8b45b8                  mov eax, dword ptr [ebp-48]
'00494739    8bf0                    mov esi, eax
'0049473b    8b08                    mov ecx, dword ptr [eax]
'0049473d    8d55d4                  lea edx, dword ptr [ebp-2c]
'00494740    52                      push edx
'00494741    50                      push eax
'00494742    ff5150                  call dword ptr [ecx+50]
var_44 = var_61.Path
'00494745    dbe2                    fnclex
'00494747    85c0                    test ax, ax
'00494749    7d0f                    jge 49475a
'0049474b    6a50                    push 50
'0049474d    682c1c4300              push 00431c2c
'00494752    56                      push esi
'00494753    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00494754    ff1580104000            call dword ptr [00401080]
'0049475a    8b45d4                  mov eax, dword ptr [ebp-2c]
'0049475d    50                      push eax
'0049475e    6890214300              push 00432190

' *** Reference to "__vbaStrCat"
'00494763    ff1570104000            call dword ptr [00401070]
var_13 = (var_44) & ("\back.jpg")
'00494769    8945a8                  mov dword ptr [ebp-58], eax
'0049476c    c745a008000000          mov dword ptr [ebp-60], 00000008
'00494773    8b0f                    mov ecx, dword ptr [edi]
'00494775    8d55b4                  lea edx, dword ptr [ebp-4c]
'00494778    52                      push edx
'00494779    83ec10                  sub esp, 10
'0049477c    8bd4                    mov edx, esp
'0049477e    8bb5e0feffff            mov esi, dword ptr [ebp+fffffee0]
'00494784    8932                    mov dword ptr [edx], esi
'00494786    8bb5e4feffff            mov esi, dword ptr [ebp+fffffee4]
'0049478c    897204                  mov dword ptr [edx+04], esi
'0049478f    8bb5e8feffff            mov esi, dword ptr [ebp+fffffee8]
'00494795    897208                  mov dword ptr [edx+08], esi
'00494798    8bb5ecfeffff            mov esi, dword ptr [ebp+fffffeec]
'0049479e    89720c                  mov dword ptr [edx+0c], esi
'004947a1    83ec10                  sub esp, 10
'004947a4    8bd4                    mov edx, esp
'004947a6    8bb5f0feffff            mov esi, dword ptr [ebp+fffffef0]
'004947ac    8932                    mov dword ptr [edx], esi
'004947ae    8bb5f4feffff            mov esi, dword ptr [ebp+fffffef4]
'004947b4    897204                  mov dword ptr [edx+04], esi
'004947b7    8bb5f8feffff            mov esi, dword ptr [ebp+fffffef8]
'004947bd    897208                  mov dword ptr [edx+08], esi
'004947c0    8bb5fcfeffff            mov esi, dword ptr [ebp+fffffefc]
'004947c6    89720c                  mov dword ptr [edx+0c], esi
'004947c9    83ec10                  sub esp, 10
'004947cc    8bd4                    mov edx, esp
'004947ce    8bb500ffffff            mov esi, dword ptr [ebp+ffffff00]
'004947d4    8932                    mov dword ptr [edx], esi
'004947d6    8bb504ffffff            mov esi, dword ptr [ebp+ffffff04]
'004947dc    897204                  mov dword ptr [edx+04], esi
'004947df    8bb508ffffff            mov esi, dword ptr [ebp+ffffff08]
'004947e5    897208                  mov dword ptr [edx+08], esi
'004947e8    8bb50cffffff            mov esi, dword ptr [ebp+ffffff0c]
'004947ee    89720c                  mov dword ptr [edx+0c], esi
'004947f1    83ec10                  sub esp, 10
'004947f4    8bd4                    mov edx, esp
'004947f6    8bb510ffffff            mov esi, dword ptr [ebp+ffffff10]
'004947fc    8932                    mov dword ptr [edx], esi
'004947fe    8bb514ffffff            mov esi, dword ptr [ebp+ffffff14]
'00494804    897204                  mov dword ptr [edx+04], esi
'00494807    8bb518ffffff            mov esi, dword ptr [ebp+ffffff18]
'0049480d    897208                  mov dword ptr [edx+08], esi
'00494810    8bb51cffffff            mov esi, dword ptr [ebp+ffffff1c]
'00494816    89720c                  mov dword ptr [edx+0c], esi
'00494819    83ec10                  sub esp, 10
'0049481c    8bd4                    mov edx, esp
'0049481e    8b75a0                  mov esi, dword ptr [ebp-60]
'00494821    8932                    mov dword ptr [edx], esi
'00494823    8b75a4                  mov esi, dword ptr [ebp-5c]
'00494826    897204                  mov dword ptr [edx+04], esi
'00494829    894208                  mov dword ptr [edx+08], eax
'0049482c    8b45ac                  mov eax, dword ptr [ebp-54]
'0049482f    89420c                  mov dword ptr [edx+0c], eax
'00494832    57                      push edi
'00494833    ff5144                  call dword ptr [ecx+44]
Set var_62 = var_2.LoadPicture(var_13)
'00494836    dbe2                    fnclex
'00494838    85c0                    test ax, ax
'0049483a    7d0f                    jge 49484b
'0049483c    6a44                    push 44
'0049483e    6860004300              push 00430060
'00494843    57                      push edi
'00494844    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00494845    ff1580104000            call dword ptr [00401080]
'0049484b    8b45b4                  mov eax, dword ptr [ebp-4c]
'0049484e    c745b400000000          mov dword ptr [ebp-4c], 00000000
'00494855    8bb5acfeffff            mov esi, dword ptr [ebp+fffffeac]
'0049485b    8b3e                    mov edi, dword ptr [esi]
'0049485d    50                      push eax
'0049485e    8d4db0                  lea ecx, dword ptr [ebp-50]
'00494861    51                      push ecx

' *** Reference to "__vbaObjSet"
'00494862    ff15b4104000            call dword ptr [004010b4]
Set var_6 = var_62
'00494868    50                      push eax
'00494869    56                      push esi
'0049486a    ff9754010000            call dword ptr [edi+00000154]
'00494870    dbe2                    fnclex
'00494872    85c0                    test ax, ax
'00494874    7d12                    jge 494888
'00494876    6854010000              push 00000154
'0049487b    684cfd4200              push 0042fd4c
'00494880    56                      push esi
'00494881    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00494882    ff1580104000            call dword ptr [00401080]
'00494888    8d4dd4                  lea ecx, dword ptr [ebp-2c]

' *** Reference to "__vbaFreeStr"
'0049488b    ff150c134000            call dword ptr [0040130c]
'#Cleanup(var_44)
'00494891    8d55b0                  lea edx, dword ptr [ebp-50]
'00494894    52                      push edx
'00494895    8d45b8                  lea eax, dword ptr [ebp-48]
'00494898    50                      push eax
'00494899    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'0049489b    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_61, var_6)
'004948a1    83c40c                  add esp, 0c
'004948a4    8d4da0                  lea ecx, dword ptr [ebp-60]

' *** Reference to "__vbaFreeVar"
'004948a7    ff1528104000            call dword ptr [00401028]
'#Cleanup(var_7)
'004948ad    a110b07200              mov ax, word ptr [0072b010]
'004948b2    85c0                    test ax, ax
'004948b4    7510                    jne 4948c6
'004948b6    6810b07200              push 0072b010
'004948bb    68b0e54000              push 0040e5b0

' *** Reference to "__vbaNew2"
'004948c0    ff152c124000            call dword ptr [0040122c]
Set var_60 = New frmMain
'004948c6    8b3510b07200            mov esi, dword ptr [0072b010]
'004948cc    8b3e                    mov edi, dword ptr [esi]
'004948ce    68a8214300              push 004321a8
'004948d3    8b0d3cb07200            mov ecx, dword ptr [0072b03c]
'004948d9    51                      push ecx

' *** Reference to "__vbaStrCat"
'004948da    ff1570104000            call dword ptr [00401070]
var_63 = ("RoleDD Gestion de personnages pour Donjons et Dragons 3.5 [") & (vbNullString)
'004948e0    8bd0                    mov edx, eax
'004948e2    8d4dd4                  lea ecx, dword ptr [ebp-2c]

' *** Reference to "__vbaStrMove"
'004948e5    ff15d0124000            call dword ptr [004012d0]
'004948eb    50                      push eax
'004948ec    6824224300              push 00432224

' *** Reference to "__vbaStrCat"
'004948f1    ff1570104000            call dword ptr [00401070]
var_64 = (var_63) & ("]")
'004948f7    8bd0                    mov edx, eax
'004948f9    8d4dd0                  lea ecx, dword ptr [ebp-30]

' *** Reference to "__vbaStrMove"
'004948fc    ff15d0124000            call dword ptr [004012d0]
'00494902    50                      push eax
'00494903    56                      push esi
'00494904    ff5754                  call dword ptr [edi+54]
'00494907    dbe2                    fnclex
'00494909    33ff                    xor edi, edi
var_num7 = Empty
'0049490b    3bc7                    cmp eax, edi
'0049490d    0f8d51030000            jge 494c64

If (-4616 < frmMain) Then
'00494913    6a54                    push 54
'00494915    684cfd4200              push 0042fd4c
'0049491a    56                      push esi
'0049491b    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0049491c    8b3580104000            mov esi, dword ptr [00401080]
'00494922    ffd6                    call esi
'00494924    e941030000              jmp 494c6a

' *** Reference to "rtcErrObj"
'00494929    8b1d6c124000            mov ebx, dword ptr [0040126c]
'0049492f    ffd3                    call ebx
'00494931    50                      push eax
'00494932    8d55b8                  lea edx, dword ptr [ebp-48]
'00494935    52                      push edx

' *** Reference to "__vbaObjSet"
'00494936    8b3db4104000            mov edi, dword ptr [004010b4]
'0049493c    ffd7                    call edi
    Set var_61 = Err
'0049493e    8bf0                    mov esi, eax
'00494940    8b06                    mov eax, dword ptr [esi]
'00494942    8d4dd4                  lea ecx, dword ptr [ebp-2c]
'00494945    51                      push ecx
'00494946    56                      push esi
'00494947    ff502c                  call dword ptr [eax+2c]
    var_44 = var_61.Description
'0049494a    dbe2                    fnclex
'0049494c    85c0                    test ax, ax
'0049494e    7d0f                    jge 49495f
'00494950    6a2c                    push 2c
'00494952    685c084300              push 0043085c
'00494957    56                      push esi
'00494958    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00494959    ff1580104000            call dword ptr [00401080]
'0049495f    ffd3                    call ebx
'00494961    50                      push eax
'00494962    8d55b4                  lea edx, dword ptr [ebp-4c]
'00494965    52                      push edx
'00494966    ffd7                    call edi
    Set var_62 = Err
'00494968    8bf0                    mov esi, eax
'0049496a    8b06                    mov eax, dword ptr [esi]
'0049496c    8d8dc8feffff            lea ecx, dword ptr [ebp+fffffec8]
'00494972    51                      push ecx
'00494973    56                      push esi
'00494974    ff501c                  call dword ptr [eax+1c]
    var_65 = var_62.Number
'00494977    dbe2                    fnclex
'00494979    85c0                    test ax, ax
'0049497b    7d0f                    jge 49498c
'0049497d    6a1c                    push 1c
'0049497f    685c084300              push 0043085c
'00494984    56                      push esi
'00494985    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00494986    ff1580104000            call dword ptr [00401080]
'0049498c    b804000280              mov eax, 80020004
'00494991    894588                  mov dword ptr [ebp-78], eax
'00494994    b90a000000              mov ecx, 0000000a
'00494999    894d80                  mov dword ptr [ebp-80], ecx
'0049499c    894598                  mov dword ptr [ebp-68], eax
'0049499f    894d90                  mov dword ptr [ebp-70], ecx
'004949a2    c78518ffffff24b07200    mov dword ptr [ebp+ffffff18], 0072b024
'004949ac    c78510ffffff08400000    mov dword ptr [ebp+ffffff10], 00004008
'004949b6    6814084300              push 00430814
'004949bb    8b55d4                  mov edx, dword ptr [ebp-2c]
'004949be    52                      push edx

' *** Reference to "__vbaStrCat"
'004949bf    8b3570104000            mov esi, dword ptr [00401070]
'004949c5    ffd6                    call esi
    var_26 = ("L'erreur suivante s'est produite : ") & (var_44)
'004949c7    8bd0                    mov edx, eax
'004949c9    8d4dd0                  lea ecx, dword ptr [ebp-30]

' *** Reference to "__vbaStrMove"
'004949cc    8b3dd0124000            mov edi, dword ptr [004012d0]
'004949d2    ffd7                    call edi
'004949d4    50                      push eax
'004949d5    6870084300              push 00430870
'004949da    ffd6                    call esi
    var_66 = (var_26) & (vbCrLf)
'004949dc    8bd0                    mov edx, eax
'004949de    8d4dcc                  lea ecx, dword ptr [ebp-34]
'004949e1    ffd7                    call edi
'004949e3    50                      push eax
'004949e4    6870084300              push 00430870
'004949e9    ffd6                    call esi
    var_67 = (var_66) & (vbCrLf)
'004949eb    8bd0                    mov edx, eax
'004949ed    8d4dc8                  lea ecx, dword ptr [ebp-38]
'004949f0    ffd7                    call edi
'004949f2    50                      push eax
'004949f3    6880084300              push 00430880
'004949f8    ffd6                    call esi
    var_68 = (var_67) & (" numéro d'erreur (")
'004949fa    8bd0                    mov edx, eax
'004949fc    8d4dc4                  lea ecx, dword ptr [ebp-3c]
'004949ff    ffd7                    call edi
'00494a01    50                      push eax
'00494a02    8b85c8feffff            mov eax, dword ptr [ebp+fffffec8]
'00494a08    50                      push eax

' *** Reference to "__vbaStrI4"
'00494a09    ff1520104000            call dword ptr [00401020]
'00494a0f    8bd0                    mov edx, eax
'00494a11    8d4dc0                  lea ecx, dword ptr [ebp-40]
'00494a14    ffd7                    call edi
'00494a16    50                      push eax
'00494a17    ffd6                    call esi
    var_69 = (var_68) & (CStr(var_65))
'00494a19    8bd0                    mov edx, eax
'00494a1b    8d4dbc                  lea ecx, dword ptr [ebp-44]
'00494a1e    ffd7                    call edi
'00494a20    50                      push eax
'00494a21    68ac084300              push 004308ac
'00494a26    ffd6                    call esi
    var_70 = (var_69) & (" )")
'00494a28    8945a8                  mov dword ptr [ebp-58], eax
'00494a2b    bf08000000              mov edi, 00000008
'00494a30    897da0                  mov dword ptr [ebp-60], edi
'00494a33    8d4d80                  lea ecx, dword ptr [ebp-80]
'00494a36    51                      push ecx
'00494a37    8d5590                  lea edx, dword ptr [ebp-70]
'00494a3a    52                      push edx
'00494a3b    8d8510ffffff            lea eax, dword ptr [ebp+ffffff10]
'00494a41    50                      push eax
'00494a42    6a10                    push 10
'00494a44    8d4da0                  lea ecx, dword ptr [ebp-60]
'00494a47    51                      push ecx

' *** Reference to "rtcMsgBox"
'00494a48    ff15b8104000            call dword ptr [004010b8]
    var_71 = MsgBox(var_70, 16, vbNullString)
'00494a4e    8d55bc                  lea edx, dword ptr [ebp-44]
'00494a51    52                      push edx
'00494a52    8d45c0                  lea eax, dword ptr [ebp-40]
'00494a55    50                      push eax
'00494a56    8d4dc4                  lea ecx, dword ptr [ebp-3c]
'00494a59    51                      push ecx
'00494a5a    8d55c8                  lea edx, dword ptr [ebp-38]
'00494a5d    52                      push edx
'00494a5e    8d45cc                  lea eax, dword ptr [ebp-34]
'00494a61    50                      push eax
'00494a62    8d4dd0                  lea ecx, dword ptr [ebp-30]
'00494a65    51                      push ecx
'00494a66    8d55d4                  lea edx, dword ptr [ebp-2c]
'00494a69    52                      push edx
'00494a6a    6a07                    push 07

' *** Reference to "__vbaFreeStrList"
'00494a6c    ff155c124000            call dword ptr [0040125c]
    '#Cleanup( -4612, -4632, -4636, -4640, -4644, -4648, -4652)
'00494a72    8d45b4                  lea eax, dword ptr [ebp-4c]
'00494a75    50                      push eax
'00494a76    8d4db8                  lea ecx, dword ptr [ebp-48]
'00494a79    51                      push ecx
'00494a7a    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'00494a7c    ff154c104000            call dword ptr [0040104c]
    '#Cleanup( var_61, var_62)
'00494a82    8d5580                  lea edx, dword ptr [ebp-80]
'00494a85    52                      push edx
'00494a86    8d4590                  lea eax, dword ptr [ebp-70]
'00494a89    50                      push eax
'00494a8a    8d4da0                  lea ecx, dword ptr [ebp-60]
'00494a8d    51                      push ecx
'00494a8e    6a03                    push 03

' *** Reference to "__vbaFreeVarList"
'00494a90    ff1540104000            call dword ptr [00401040]
    '#Cleanup( var_7, var_8, var_18)
'00494a96    83c43c                  add esp, 3c
'00494a99    8d55a0                  lea edx, dword ptr [ebp-60]
'00494a9c    52                      push edx

' *** Reference to "rtcGetPresentDate"
'00494a9d    ff15f4124000            call dword ptr [004012f4]
    var_70 = Now()
'00494aa3    c78518ffffffb8084300    mov dword ptr [ebp+ffffff18], 004308b8
'00494aad    89bd10ffffff            mov dword ptr [ebp+ffffff10], edi
'00494ab3    8d9510ffffff            lea edx, dword ptr [ebp+ffffff10]
'00494ab9    8d4d90                  lea ecx, dword ptr [ebp-70]

' *** Reference to "__vbaVarDup"
'00494abc    ff1598124000            call dword ptr [00401298]
    var_8 = ("dd/MM/yyyy hh:mm:ss")
'00494ac2    6a01                    push 01
'00494ac4    6a01                    push 01
'00494ac6    8d4590                  lea eax, dword ptr [ebp-70]
'00494ac9    50                      push eax
'00494aca    8d4da0                  lea ecx, dword ptr [ebp-60]
'00494acd    51                      push ecx
'00494ace    8d5580                  lea edx, dword ptr [ebp-80]
'00494ad1    52                      push edx

' *** Reference to "rtcVarFromFormatVar"
'00494ad2    ff1574104000            call dword ptr [00401074]
'00494ad8    ffd3                    call ebx
'00494ada    50                      push eax
'00494adb    8d45b8                  lea eax, dword ptr [ebp-48]
'00494ade    50                      push eax

' *** Reference to "__vbaObjSet"
'00494adf    ff15b4104000            call dword ptr [004010b4]
    Set var_61 = Err
'00494ae5    8bf0                    mov esi, eax
'00494ae7    8b0e                    mov ecx, dword ptr [esi]
'00494ae9    8d95c8feffff            lea edx, dword ptr [ebp+fffffec8]
'00494aef    52                      push edx
'00494af0    56                      push esi
'00494af1    ff511c                  call dword ptr [ecx+1c]
    var_65 = var_61.Number
'00494af4    dbe2                    fnclex
'00494af6    85c0                    test ax, ax
'00494af8    7d0f                    jge 494b09
'00494afa    6a1c                    push 1c
'00494afc    685c084300              push 0043085c
'00494b01    56                      push esi
'00494b02    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00494b03    ff1580104000            call dword ptr [00401080]
'00494b09    ffd3                    call ebx
'00494b0b    50                      push eax
'00494b0c    8d45b4                  lea eax, dword ptr [ebp-4c]
'00494b0f    50                      push eax

' *** Reference to "__vbaObjSet"
'00494b10    ff15b4104000            call dword ptr [004010b4]
    Set var_62 = Err
'00494b16    8bf0                    mov esi, eax
'00494b18    8b0e                    mov ecx, dword ptr [esi]
'00494b1a    8d55d4                  lea edx, dword ptr [ebp-2c]
'00494b1d    52                      push edx
'00494b1e    56                      push esi
'00494b1f    ff512c                  call dword ptr [ecx+2c]
    var_44 = var_62.Description
'00494b22    dbe2                    fnclex
'00494b24    85c0                    test ax, ax
'00494b26    7d0f                    jge 494b37
'00494b28    6a2c                    push 2c
'00494b2a    685c084300              push 0043085c
'00494b2f    56                      push esi
'00494b30    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00494b31    ff1580104000            call dword ptr [00401080]
'00494b37    c78508ffffffe4084300    mov dword ptr [ebp+ffffff08], 004308e4
'00494b41    89bd00ffffff            mov dword ptr [ebp+ffffff00], edi
'00494b47    8b85c8feffff            mov eax, dword ptr [ebp+fffffec8]
'00494b4d    8985f8feffff            mov dword ptr [ebp+fffffef8], eax
'00494b53    c785f0feffff03000000    mov dword ptr [ebp+fffffef0], 00000003
'00494b5d    c785e8feffff08094300    mov dword ptr [ebp+fffffee8], 00430908
'00494b67    89bde0feffff            mov dword ptr [ebp+fffffee0], edi
'00494b6d    8b45d4                  mov eax, dword ptr [ebp-2c]
'00494b70    c745d400000000          mov dword ptr [ebp-2c], 00000000
'00494b77    898548ffffff            mov dword ptr [ebp+ffffff48], eax
'00494b7d    89bd40ffffff            mov dword ptr [ebp+ffffff40], edi
'00494b83    c785d8feffff5c224300    mov dword ptr [ebp+fffffed8], 0043225c
'00494b8d    89bdd0feffff            mov dword ptr [ebp+fffffed0], edi
'00494b93    8d4d80                  lea ecx, dword ptr [ebp-80]
'00494b96    51                      push ecx
'00494b97    8d9500ffffff            lea edx, dword ptr [ebp+ffffff00]
'00494b9d    52                      push edx
'00494b9e    8d8570ffffff            lea eax, dword ptr [ebp+ffffff70]
'00494ba4    50                      push eax

' *** Reference to "__vbaVarCat"
'00494ba5    8b3508124000            mov esi, dword ptr [00401208]
'00494bab    ffd6                    call esi
'00494bad    50                      push eax
'00494bae    8d8df0feffff            lea ecx, dword ptr [ebp+fffffef0]
'00494bb4    51                      push ecx
'00494bb5    8d9560ffffff            lea edx, dword ptr [ebp+ffffff60]
'00494bbb    52                      push edx
'00494bbc    ffd6                    call esi
'00494bbe    50                      push eax
'00494bbf    8d85e0feffff            lea eax, dword ptr [ebp+fffffee0]
'00494bc5    50                      push eax
'00494bc6    8d8d50ffffff            lea ecx, dword ptr [ebp+ffffff50]
'00494bcc    51                      push ecx
'00494bcd    ffd6                    call esi
'00494bcf    50                      push eax
'00494bd0    8d9540ffffff            lea edx, dword ptr [ebp+ffffff40]
'00494bd6    52                      push edx
'00494bd7    8d8530ffffff            lea eax, dword ptr [ebp+ffffff30]
'00494bdd    50                      push eax
'00494bde    ffd6                    call esi
'00494be0    50                      push eax
'00494be1    8d8dd0feffff            lea ecx, dword ptr [ebp+fffffed0]
'00494be7    51                      push ecx
'00494be8    8d9520ffffff            lea edx, dword ptr [ebp+ffffff20]
'00494bee    52                      push edx
'00494bef    ffd6                    call esi
'00494bf1    50                      push eax
'00494bf2    33c0                    xor eax, eax
'00494bf4    66a12ab07200            mov eax, dword ptr [0072b02a]
'00494bfa    50                      push eax
'00494bfb    6884094300              push 00430984

' *** Reference to "__vbaPrintFile"
'00494c00    ff15b8114000            call dword ptr [004011b8]
    Print #0, 
'00494c06    8d4db4                  lea ecx, dword ptr [ebp-4c]
'00494c09    51                      push ecx
'00494c0a    8d55b8                  lea edx, dword ptr [ebp-48]
'00494c0d    52                      push edx
'00494c0e    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'00494c10    ff154c104000            call dword ptr [0040104c]
    '#Cleanup( var_61, var_62)
'00494c16    8d8520ffffff            lea eax, dword ptr [ebp+ffffff20]
'00494c1c    50                      push eax
'00494c1d    8d8d30ffffff            lea ecx, dword ptr [ebp+ffffff30]
'00494c23    51                      push ecx
'00494c24    8d9540ffffff            lea edx, dword ptr [ebp+ffffff40]
'00494c2a    52                      push edx
'00494c2b    8d8550ffffff            lea eax, dword ptr [ebp+ffffff50]
'00494c31    50                      push eax
'00494c32    8d8d60ffffff            lea ecx, dword ptr [ebp+ffffff60]
'00494c38    51                      push ecx
'00494c39    8d9570ffffff            lea edx, dword ptr [ebp+ffffff70]
'00494c3f    52                      push edx
'00494c40    8d4580                  lea eax, dword ptr [ebp-80]
'00494c43    50                      push eax
'00494c44    8d4d90                  lea ecx, dword ptr [ebp-70]
'00494c47    51                      push ecx
'00494c48    8d55a0                  lea edx, dword ptr [ebp-60]
'00494c4b    52                      push edx
'00494c4c    6a09                    push 09

' *** Reference to "__vbaFreeVarList"
'00494c4e    ff1540104000            call dword ptr [00401040]
    '#Cleanup( var_7, var_8, var_18, var_19, var_20, var_21, var_22, var_23, var_72)
'00494c54    83c440                  add esp, 40
'00494c57    6a00                    push 00

' *** Reference to "__vbaResume"
'00494c59    ff1568104000            call dword ptr [00401068]
'00494c5f    e975030000              jmp 494fd9
    Resume handler_494FD9
End If

' *** Reference to "__vbaHresultCheckObj"
'00494c64    8b3580104000            mov esi, dword ptr [00401080]
'00494c6a    8d55d0                  lea edx, dword ptr [ebp-30]
'00494c6d    52                      push edx
'00494c6e    8d45d4                  lea eax, dword ptr [ebp-2c]
'00494c71    50                      push eax
'00494c72    6a02                    push 02

' *** Reference to "__vbaFreeStrList"
'00494c74    ff155c124000            call dword ptr [0040125c]
'#Cleanup( var_44, -4632)
'00494c7a    83c40c                  add esp, 0c
'00494c7d    8b0b                    mov ecx, dword ptr [ebx]
'00494c7f    8d55d4                  lea edx, dword ptr [ebp-2c]
'00494c82    52                      push edx
'00494c83    53                      push ebx
'00494c84    ff5148                  call dword ptr [ecx+48]
'00494c87    dbe2                    fnclex
'00494c89    3bc7                    cmp eax, edi
'00494c8b    7d0b                    jge 494c98
'00494c8d    6a48                    push 48
'00494c8f    684cfd4200              push 0042fd4c
'00494c94    53                      push ebx
'00494c95    50                      push eax
'00494c96    ffd6                    call esi
'00494c98    89bdccfeffff            mov dword ptr [ebp+fffffecc], edi
'00494c9e    ba2c224300              mov edx, 0043222c
'00494ca3    8d4dcc                  lea ecx, dword ptr [ebp-34]

' *** Reference to "__vbaStrCopy"
'00494ca6    ff1548124000            call dword ptr [00401248]
var_43 = ("POS_X")
'00494cac    8b55d4                  mov edx, dword ptr [ebp-2c]
'00494caf    897dd4                  mov dword ptr [ebp-2c], edi
'00494cb2    8d4dd0                  lea ecx, dword ptr [ebp-30]

' *** Reference to "__vbaStrMove"
'00494cb5    ff15d0124000            call dword ptr [004012d0]
'00494cbb    8b3b                    mov edi, dword ptr [ebx]
'00494cbd    8d85ccfeffff            lea eax, dword ptr [ebp+fffffecc]
'00494cc3    50                      push eax
'00494cc4    8d4dcc                  lea ecx, dword ptr [ebp-34]
'00494cc7    51                      push ecx
'00494cc8    8d55d0                  lea edx, dword ptr [ebp-30]
'00494ccb    52                      push edx
'00494ccc    e89ffb0500              call 4f4870
Call sub_4F4870()
'00494cd1    0fbfc0                  movsx eax, ax
'00494cd4    89857cfeffff            mov dword ptr [ebp+fffffe7c], eax
'00494cda    db857cfeffff            fild dword ptr [ebp+fffffe7c]
'00494ce0    d99d78feffff            fstp dword ptr [ebp+fffffe78]
'var_73 = (00)
'00494ce6    8b8d78feffff            mov ecx, dword ptr [ebp+fffffe78]
'00494cec    51                      push ecx
'00494ced    53                      push ebx
'00494cee    ff5774                  call dword ptr [edi+74]
'00494cf1    dbe2                    fnclex
'00494cf3    33ff                    xor edi, edi
var_num7 = Empty
'00494cf5    3bc7                    cmp eax, edi
'00494cf7    7d0b                    jge 494d04

If (-564 < frmMain) Then
'00494cf9    6a74                    push 74
'00494cfb    684cfd4200              push 0042fd4c
'00494d00    53                      push ebx
'00494d01    50                      push eax
'00494d02    ffd6                    call esi
    
End If
'00494d04    8d55cc                  lea edx, dword ptr [ebp-34]
'00494d07    52                      push edx
'00494d08    8d45d0                  lea eax, dword ptr [ebp-30]
'00494d0b    50                      push eax
'00494d0c    6a02                    push 02

' *** Reference to "__vbaFreeStrList"
'00494d0e    ff155c124000            call dword ptr [0040125c]
'#Cleanup( var_44, var_43)
'00494d14    83c40c                  add esp, 0c
'00494d17    8b0b                    mov ecx, dword ptr [ebx]
'00494d19    8d55d4                  lea edx, dword ptr [ebp-2c]
'00494d1c    52                      push edx
'00494d1d    53                      push ebx
'00494d1e    ff5148                  call dword ptr [ecx+48]
'00494d21    dbe2                    fnclex
'00494d23    3bc7                    cmp eax, edi
'00494d25    7d0b                    jge 494d32
'00494d27    6a48                    push 48
'00494d29    684cfd4200              push 0042fd4c
'00494d2e    53                      push ebx
'00494d2f    50                      push eax
'00494d30    ffd6                    call esi
'00494d32    89bdccfeffff            mov dword ptr [ebp+fffffecc], edi
'00494d38    ba3c224300              mov edx, 0043223c
'00494d3d    8d4dcc                  lea ecx, dword ptr [ebp-34]

' *** Reference to "__vbaStrCopy"
'00494d40    ff1548124000            call dword ptr [00401248]
var_43 = ("POS_Y")
'00494d46    8b55d4                  mov edx, dword ptr [ebp-2c]
'00494d49    897dd4                  mov dword ptr [ebp-2c], edi
'00494d4c    8d4dd0                  lea ecx, dword ptr [ebp-30]

' *** Reference to "__vbaStrMove"
'00494d4f    ff15d0124000            call dword ptr [004012d0]
'00494d55    8b3b                    mov edi, dword ptr [ebx]
'00494d57    8d85ccfeffff            lea eax, dword ptr [ebp+fffffecc]
'00494d5d    50                      push eax
'00494d5e    8d4dcc                  lea ecx, dword ptr [ebp-34]
'00494d61    51                      push ecx
'00494d62    8d55d0                  lea edx, dword ptr [ebp-30]
'00494d65    52                      push edx
'00494d66    e805fb0500              call 4f4870
Call sub_4F4870()
'00494d6b    0fbfc0                  movsx eax, ax
'00494d6e    898574feffff            mov dword ptr [ebp+fffffe74], eax
'00494d74    db8574feffff            fild dword ptr [ebp+fffffe74]
'00494d7a    d99d70feffff            fstp dword ptr [ebp+fffffe70]
'var_74 = (00)
'00494d80    8b8d70feffff            mov ecx, dword ptr [ebp+fffffe70]
'00494d86    51                      push ecx
'00494d87    53                      push ebx
'00494d88    ff577c                  call dword ptr [edi+7c]
'00494d8b    dbe2                    fnclex
'00494d8d    85c0                    test ax, ax
'00494d8f    7d0b                    jge 494d9c
'00494d91    6a7c                    push 7c
'00494d93    684cfd4200              push 0042fd4c
'00494d98    53                      push ebx
'00494d99    50                      push eax
'00494d9a    ffd6                    call esi
'00494d9c    8d55cc                  lea edx, dword ptr [ebp-34]
'00494d9f    52                      push edx
'00494da0    8d45d0                  lea eax, dword ptr [ebp-30]
'00494da3    50                      push eax
'00494da4    6a02                    push 02

' *** Reference to "__vbaFreeStrList"
'00494da6    ff155c124000            call dword ptr [0040125c]
'#Cleanup( frmMain, var_43)
'00494dac    83c40c                  add esp, 0c
'00494daf    a15cb07200              mov ax, word ptr [0072b05c]
'00494db4    85c0                    test ax, ax
'00494db6    7510                    jne 494dc8
'00494db8    685cb07200              push 0072b05c
'00494dbd    68347b4000              push 00407b34

' *** Reference to "__vbaNew2"
'00494dc2    ff152c124000            call dword ptr [0040122c]
Dim var_24 As New frmAcceder
'00494dc8    8b355cb07200            mov esi, dword ptr [0072b05c]
'00494dce    b804000280              mov eax, 80020004
'00494dd3    898508ffffff            mov dword ptr [ebp+ffffff08], eax
'00494dd9    b90a000000              mov ecx, 0000000a
'00494dde    898d00ffffff            mov dword ptr [ebp+ffffff00], ecx
'00494de4    8bd0                    mov edx, eax
'00494de6    899518ffffff            mov dword ptr [ebp+ffffff18], edx
'00494dec    898d10ffffff            mov dword ptr [ebp+ffffff10], ecx
'00494df2    8b3e                    mov edi, dword ptr [esi]
'00494df4    83ec10                  sub esp, 10
'00494df7    8bdc                    mov ebx, esp
'00494df9    890b                    mov dword ptr [ebx], ecx
'00494dfb    8b8d04ffffff            mov ecx, dword ptr [ebp+ffffff04]
'00494e01    894b04                  mov dword ptr [ebx+04], ecx
'00494e04    894308                  mov dword ptr [ebx+08], eax
'00494e07    8b850cffffff            mov eax, dword ptr [ebp+ffffff0c]
'00494e0d    89430c                  mov dword ptr [ebx+0c], eax
'00494e10    83ec10                  sub esp, 10
'00494e13    8bcc                    mov ecx, esp
'00494e15    8b8510ffffff            mov eax, dword ptr [ebp+ffffff10]
'00494e1b    8901                    mov dword ptr [ecx], eax
'00494e1d    8b8514ffffff            mov eax, dword ptr [ebp+ffffff14]
'00494e23    894104                  mov dword ptr [ecx+04], eax
'00494e26    895108                  mov dword ptr [ecx+08], edx
'00494e29    8b951cffffff            mov edx, dword ptr [ebp+ffffff1c]
'00494e2f    89510c                  mov dword ptr [ecx+0c], edx
'00494e32    56                      push esi
'00494e33    ff97b0020000            call dword ptr [edi+000002b0]
'00494e39    dbe2                    fnclex
'00494e3b    85c0                    test ax, ax
'00494e3d    7d12                    jge 494e51
'00494e3f    68b0020000              push 000002b0
'00494e44    6810074300              push 00430710
'00494e49    56                      push esi
'00494e4a    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00494e4b    ff1580104000            call dword ptr [00401080]
'00494e51    8d45a0                  lea eax, dword ptr [ebp-60]
'00494e54    50                      push eax

' *** Reference to "rtcGetPresentDate"
'00494e55    ff15f4124000            call dword ptr [004012f4]
var_70 = Now()
'00494e5b    b802000000              mov eax, 00000002
'00494e60    898518ffffff            mov dword ptr [ebp+ffffff18], eax
'00494e66    898510ffffff            mov dword ptr [ebp+ffffff10], eax
'00494e6c    8d4da0                  lea ecx, dword ptr [ebp-60]
'00494e6f    51                      push ecx
'00494e70    8d9510ffffff            lea edx, dword ptr [ebp+ffffff10]
'00494e76    52                      push edx
'00494e77    8d4590                  lea eax, dword ptr [ebp-70]
'00494e7a    50                      push eax

' *** Reference to "__vbaVarDiv"
'00494e7b    ff15d8114000            call dword ptr [004011d8]
'00494e81    50                      push eax

' *** Reference to "__vbaR8Var"
'00494e82    ff1564124000            call dword ptr [00401264]
'00494e88    dd5dd8                  fstp qword ptr [ebp-28]
'var_45 = (00)
'00494e8b    8d4da0                  lea ecx, dword ptr [ebp-60]

' *** Reference to "__vbaFreeVar"
'00494e8e    8b3d28104000            mov edi, dword ptr [00401028]
'00494e94    ffd7                    call edi
'#Cleanup(var_7)
'00494e96    dd45d8                  fld qword ptr [ebp-28]

' *** Reference to "__vbaFPInt"
'00494e99    ff15fc124000            call dword ptr [004012fc]
'00494e9f    dc6dd8                  fsubr qword ptr [ebp-28]
'00494ea2    dd5dd8                  fstp qword ptr [ebp-28]
'var_45 = (00)
'00494ea5    dfe0                    fnstsw ax
'00494ea7    a80d                    test al, 0d
'00494ea9    0f85d2010000            jne 495081
'00494eaf    dd45d8                  fld qword ptr [ebp-28]
'00494eb2    dc0df8144000            fmul qword ptr [004014f8]
'00494eb8    dfe0                    fnstsw ax
'00494eba    a80d                    test al, 0d
'00494ebc    0f85bf010000            jne 495081

' *** Reference to "__vbaFPInt"
'00494ec2    ff15fc124000            call dword ptr [004012fc]
'00494ec8    dd5dd8                  fstp qword ptr [ebp-28]
'var_45 = (00)
'00494ecb    dd45d8                  fld qword ptr [ebp-28]
'00494ece    833d00b0720000          cmp dword ptr [0072b000], 00
'00494ed5    7508                    jne 494edf

If (vbNullChar = 0) Then
'00494ed7    dc35f0144000            fdiv qword ptr [004014f0]
'00494edd    eb11                    jmp 494ef0
    
End If
'00494edf    ff35f4144000            push dword ptr [004014f4]
'00494ee5    ff35f0144000            push dword ptr [004014f0]

' *** Reference to "_adj_fdiv_m64"
'00494eeb    e89423f7ff              call 407284
'00494ef0    dfe0                    fnstsw ax
'00494ef2    a80d                    test al, 0d
'00494ef4    0f8587010000            jne 495081

' *** Reference to "__vbaFPInt"
'00494efa    ff15fc124000            call dword ptr [004012fc]
'00494f00    dc0df0144000            fmul qword ptr [004014f0]
'00494f06    dc6dd8                  fsubr qword ptr [ebp-28]
'00494f09    dfe0                    fnstsw ax
'00494f0b    a80d                    test al, 0d
'00494f0d    0f856e010000            jne 495081

' *** Reference to "__vbaFpI2"
'00494f13    ff15a0124000            call dword ptr [004012a0]
var_num1 = CLng(((Int(((Int((((Int((CDbl(#NOT SUPPORTED#))) - CDbl(#NOT SUPPORTED#))) * 1000000#))) / 256#)) * 256#) - Int((((Int((CDbl(#NOT SUPPORTED#))) - CDbl(#NOT SUPPORTED#))) * 1000000#))))
'00494f19    8985a0feffff            mov dword ptr [ebp+fffffea0], eax
'00494f1f    bb01000000              mov ebx, 00000001
'00494f24    8bf3                    mov esi, ebx
'00494f26    663bb5a0feffff          cmp si, word ptr [ebp+fffffea0]
'00494f2d    7f33                    jg 494f62

Do While (1 <= WORD PTR [EBP+FFFFFEA0])
'00494f2f    c745a8e8030000          mov dword ptr [ebp-58], 000003e8
'00494f36    c745a002000000          mov dword ptr [ebp-60], 00000002
'00494f3d    8d4da0                  lea ecx, dword ptr [ebp-60]
'00494f40    51                      push ecx

' *** Reference to "rtcRandomNext"
'00494f41    ff15a4104000            call dword ptr [004010a4]
'00494f47    d99dc8feffff            fstp dword ptr [ebp+fffffec8]
    'var_65 = (00)
'00494f4d    8d4da0                  lea ecx, dword ptr [ebp-60]
'00494f50    ffd7                    call edi
    '#Cleanup(var_7)
'00494f52    668bd3                  mov dx, bx
'00494f55    6603d6                  add dx, si
    var_num4 = 1 + 1
'00494f58    0f8028010000            jo 495086
'00494f5e    8bf2                    mov esi, edx
'00494f60    ebc4                    jmp 494f26
    
Loop
'00494f62    8d45a0                  lea eax, dword ptr [ebp-60]
'00494f65    50                      push eax

' *** Reference to "rtcCommandVar"
'00494f66    8b3564114000            mov esi, dword ptr [00401164]
'00494f6c    ffd6                    call esi
'00494f6e    c78518ffffffcc134300    mov dword ptr [ebp+ffffff18], 004313cc
'00494f78    c78510ffffff08800000    mov dword ptr [ebp+ffffff10], 00008008
'00494f82    8d4da0                  lea ecx, dword ptr [ebp-60]
'00494f85    51                      push ecx
'00494f86    8d9510ffffff            lea edx, dword ptr [ebp+ffffff10]
'00494f8c    52                      push edx

' *** Reference to "__vbaVarTstNe"
'00494f8d    ff1574124000            call dword ptr [00401274]
'00494f93    668bd8                  mov bx, ax
'00494f96    8d4da0                  lea ecx, dword ptr [ebp-60]
'00494f99    ffd7                    call edi
'#Cleanup(var_7)
'00494f9b    6685db                  test ebx, ebx
'00494f9e    7439                    je 494fd9

If (((Command()) <> (vbNullChar))) Then
'00494fa0    8d45a0                  lea eax, dword ptr [ebp-60]
'00494fa3    50                      push eax
'00494fa4    ffd6                    call esi
'00494fa6    8d4da0                  lea ecx, dword ptr [ebp-60]
'00494fa9    51                      push ecx

' *** Reference to "__vbaStrVarMove"
'00494faa    ff153c104000            call dword ptr [0040103c]
'00494fb0    8bd0                    mov edx, eax
'00494fb2    8d4dd4                  lea ecx, dword ptr [ebp-2c]

' *** Reference to "__vbaStrMove"
'00494fb5    ff15d0124000            call dword ptr [004012d0]
'00494fbb    8b4508                  mov eax, dword ptr [ebp+08]
'00494fbe    8b10                    mov edx, dword ptr [eax]
'00494fc0    8d4dd4                  lea ecx, dword ptr [ebp-2c]
'00494fc3    51                      push ecx
'00494fc4    50                      push eax
'00494fc5    ff9254070000            call dword ptr [edx+00000754]
'00494fcb    8d4dd4                  lea ecx, dword ptr [ebp-2c]

' *** Reference to "__vbaFreeStr"
'00494fce    ff150c134000            call dword ptr [0040130c]
    '#Cleanup(var_44)
'00494fd4    8d4da0                  lea ecx, dword ptr [ebp-60]
'00494fd7    ffd7                    call edi
    '#Cleanup(var_7)
End If

' *** Reference to "__vbaExitProc"
'00494fd9    ff15a0104000            call dword ptr [004010a0]
'00494fdf    9b                      fwait
'00494fe0    6862504900              push 00495062
'00494fe5    eb7a                    jmp 495061
'00494fe7    8d55bc                  lea edx, dword ptr [ebp-44]
'00494fea    52                      push edx
'00494feb    8d45c0                  lea eax, dword ptr [ebp-40]
'00494fee    50                      push eax
'00494fef    8d4dc4                  lea ecx, dword ptr [ebp-3c]
'00494ff2    51                      push ecx
'00494ff3    8d55c8                  lea edx, dword ptr [ebp-38]
'00494ff6    52                      push edx
'00494ff7    8d45cc                  lea eax, dword ptr [ebp-34]
'00494ffa    50                      push eax
'00494ffb    8d4dd0                  lea ecx, dword ptr [ebp-30]
'00494ffe    51                      push ecx
'00494fff    8d55d4                  lea edx, dword ptr [ebp-2c]
'00495002    52                      push edx
'00495003    6a07                    push 07

' *** Reference to "__vbaFreeStrList"
'00495005    ff155c124000            call dword ptr [0040125c]
'#Cleanup( , frmMain, var_43, -4640, -4644, -4648, -4652)
'0049500b    8d45b0                  lea eax, dword ptr [ebp-50]
'0049500e    50                      push eax
'0049500f    8d4db4                  lea ecx, dword ptr [ebp-4c]
'00495012    51                      push ecx
'00495013    8d55b8                  lea edx, dword ptr [ebp-48]
'00495016    52                      push edx
'00495017    6a03                    push 03

' *** Reference to "__vbaFreeObjList"
'00495019    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_61, var_62, var_6)
'0049501f    8d8520ffffff            lea eax, dword ptr [ebp+ffffff20]
'00495025    50                      push eax
'00495026    8d8d30ffffff            lea ecx, dword ptr [ebp+ffffff30]
'0049502c    51                      push ecx
'0049502d    8d9540ffffff            lea edx, dword ptr [ebp+ffffff40]
'00495033    52                      push edx
'00495034    8d8550ffffff            lea eax, dword ptr [ebp+ffffff50]
'0049503a    50                      push eax
'0049503b    8d8d60ffffff            lea ecx, dword ptr [ebp+ffffff60]
'00495041    51                      push ecx
'00495042    8d9570ffffff            lea edx, dword ptr [ebp+ffffff70]
'00495048    52                      push edx
'00495049    8d4580                  lea eax, dword ptr [ebp-80]
'0049504c    50                      push eax
'0049504d    8d4d90                  lea ecx, dword ptr [ebp-70]
'00495050    51                      push ecx
'00495051    8d55a0                  lea edx, dword ptr [ebp-60]
'00495054    52                      push edx
'00495055    6a09                    push 09

' *** Reference to "__vbaFreeVarList"
'00495057    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_7, var_8, var_18, var_19, var_20, var_21, var_22, var_23, var_72)
'0049505d    83c458                  add esp, 58
'00495060    c3                      ret
'00495061    c3                      ret
'00495062    8b4508                  mov eax, dword ptr [ebp+08]
'00495065    8b08                    mov ecx, dword ptr [eax]
'00495067    50                      push eax
'00495068    ff5108                  call dword ptr [ecx+08]
'0049506b    8b45f4                  mov eax, dword ptr [ebp-0c]
'0049506e    8b4de4                  mov ecx, dword ptr [ebp-1c]
    'Reference to '__except_list'
'00495071    64890d00000000          mov dword ptr fs:[00000000], ecx
'00495078    5f                      pop edi
'00495079    5e                      pop esi
'0049507a    5b                      pop ebx
'0049507b    8be5                    mov esp, ebp
'0049507d    5d                      pop ebp
'0049507e    c20400                  ret 0004


End Sub


Private Sub MDIForm_Unload(Cancel as Integer)
'004951b0    55                      push ebp
'004951b1    8bec                    mov ebp, esp
'004951b3    83ec14                  sub esp, 14

' *** Reference to "__vbaExceptHandler"
'004951b6    6866724000              push 00407266
'004951bb    64a100000000            mov ax, word ptr fs:[00000000]
'004951c1    50                      push eax
    'Reference to '__except_list'
'004951c2    64892500000000          mov dword ptr fs:[00000000], esp
'004951c9    81ec38010000            sub esp, 00000138
'004951cf    53                      push ebx
'004951d0    56                      push esi
'004951d1    57                      push edi
'004951d2    8965ec                  mov dword ptr [ebp-14], esp
'004951d5    c745f010154000          mov dword ptr [ebp-10], 00401510
'004951dc    8b7508                  mov esi, dword ptr [ebp+08]
'004951df    8bc6                    mov eax, esi
'004951e1    83e001                  and eax, 01
'004951e4    8945f4                  mov dword ptr [ebp-0c], eax
'004951e7    83e6fe                  and esi, fffffffe
'004951ea    897508                  mov dword ptr [ebp+08], esi
'004951ed    33ff                    xor edi, edi
'004951ef    897df8                  mov dword ptr [ebp-08], edi
'004951f2    8b0e                    mov ecx, dword ptr [esi]
'004951f4    56                      push esi
'004951f5    ff5104                  call dword ptr [ecx+04]
'004951f8    897de0                  mov dword ptr [ebp-20], edi
'004951fb    897ddc                  mov dword ptr [ebp-24], edi
'004951fe    897dd8                  mov dword ptr [ebp-28], edi
'00495201    897dd4                  mov dword ptr [ebp-2c], edi
'00495204    897dd0                  mov dword ptr [ebp-30], edi
'00495207    897dcc                  mov dword ptr [ebp-34], edi
'0049520a    897dc8                  mov dword ptr [ebp-38], edi
'0049520d    897dc4                  mov dword ptr [ebp-3c], edi
'00495210    897dc0                  mov dword ptr [ebp-40], edi
'00495213    897db0                  mov dword ptr [ebp-50], edi
'00495216    897da0                  mov dword ptr [ebp-60], edi
'00495219    897d90                  mov dword ptr [ebp-70], edi
'0049521c    897d80                  mov dword ptr [ebp-80], edi
'0049521f    89bd70ffffff            mov dword ptr [ebp+ffffff70], edi
'00495225    89bd60ffffff            mov dword ptr [ebp+ffffff60], edi
'0049522b    89bd50ffffff            mov dword ptr [ebp+ffffff50], edi
'00495231    89bd40ffffff            mov dword ptr [ebp+ffffff40], edi
'00495237    89bd30ffffff            mov dword ptr [ebp+ffffff30], edi
'0049523d    89bd20ffffff            mov dword ptr [ebp+ffffff20], edi
'00495243    89bd10ffffff            mov dword ptr [ebp+ffffff10], edi
'00495249    89bd00ffffff            mov dword ptr [ebp+ffffff00], edi
'0049524f    89bdf0feffff            mov dword ptr [ebp+fffffef0], edi
'00495255    89bde0feffff            mov dword ptr [ebp+fffffee0], edi
'0049525b    89bddcfeffff            mov dword ptr [ebp+fffffedc], edi
'00495261    89bdd8feffff            mov dword ptr [ebp+fffffed8], edi
'00495267    66393d28b07200          cmp word ptr [0072b028], di
'0049526e    7508                    jne 495278
'00495270    6a01                    push 01

' *** Reference to "__vbaOnError"
'00495272    ff15b0104000            call dword ptr [004010b0]
    On Error Goto handler_0
'00495278    8b06                    mov eax, dword ptr [esi]
'0049527a    8d4de0                  lea ecx, dword ptr [ebp-20]
'0049527d    51                      push ecx
'0049527e    56                      push esi
'0049527f    ff5048                  call dword ptr [eax+48]
'00495282    dbe2                    fnclex
'00495284    3bc7                    cmp eax, edi
'00495286    7d0f                    jge 495297
    
    If (    frmMain < 0) Then
'00495288    6a48                    push 48
'0049528a    684cfd4200              push 0042fd4c
'0049528f    56                      push esi
'00495290    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00495291    ff1580104000            call dword ptr [00401080]
    
End If
'00495297    8b16                    mov edx, dword ptr [esi]
'00495299    8d85d8feffff            lea eax, dword ptr [ebp+fffffed8]
'0049529f    50                      push eax
'004952a0    56                      push esi
'004952a1    ff5270                  call dword ptr [edx+70]
'004952a4    dbe2                    fnclex
'004952a6    3bc7                    cmp eax, edi
'004952a8    7d0f                    jge 4952b9
'004952aa    6a70                    push 70
'004952ac    684cfd4200              push 0042fd4c
'004952b1    56                      push esi
'004952b2    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'004952b3    ff1580104000            call dword ptr [00401080]
'004952b9    d985d8feffff            fld dword ptr [ebp+fffffed8]

' *** Reference to "__vbaFpI2"
'004952bf    8b1da0124000            mov ebx, dword ptr [004012a0]
'004952c5    ffd3                    call ebx
var_num1 = CLng((0))
'004952c7    8985dcfeffff            mov dword ptr [ebp+fffffedc], eax
'004952cd    ba2c224300              mov edx, 0043222c
'004952d2    8d4dd8                  lea ecx, dword ptr [ebp-28]

' *** Reference to "__vbaStrCopy"
'004952d5    ff1548124000            call dword ptr [00401248]
var_45 = ("POS_X")
'004952db    8b55e0                  mov edx, dword ptr [ebp-20]
'004952de    897de0                  mov dword ptr [ebp-20], edi
'004952e1    8d4ddc                  lea ecx, dword ptr [ebp-24]

' *** Reference to "__vbaStrMove"
'004952e4    ff15d0124000            call dword ptr [004012d0]
'004952ea    8d8ddcfeffff            lea ecx, dword ptr [ebp+fffffedc]
'004952f0    51                      push ecx
'004952f1    8d55d8                  lea edx, dword ptr [ebp-28]
'004952f4    52                      push edx
'004952f5    8d45dc                  lea eax, dword ptr [ebp-24]
'004952f8    50                      push eax
'004952f9    e872f50500              call 4f4870
Call sub_4F4870()
'004952fe    8d4dd8                  lea ecx, dword ptr [ebp-28]
'00495301    51                      push ecx
'00495302    8d55dc                  lea edx, dword ptr [ebp-24]
'00495305    52                      push edx
'00495306    6a02                    push 02

' *** Reference to "__vbaFreeStrList"
'00495308    ff155c124000            call dword ptr [0040125c]
'#Cleanup( 0, var_45)
'0049530e    83c40c                  add esp, 0c
'00495311    8b06                    mov eax, dword ptr [esi]
'00495313    8d4de0                  lea ecx, dword ptr [ebp-20]
'00495316    51                      push ecx
'00495317    56                      push esi
'00495318    ff5048                  call dword ptr [eax+48]
'0049531b    dbe2                    fnclex
'0049531d    3bc7                    cmp eax, edi
'0049531f    7d0f                    jge 495330

If (frmMain < 0) Then
'00495321    6a48                    push 48
'00495323    684cfd4200              push 0042fd4c
'00495328    56                      push esi
'00495329    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0049532a    ff1580104000            call dword ptr [00401080]
    
End If
'00495330    8b16                    mov edx, dword ptr [esi]
'00495332    8d85d8feffff            lea eax, dword ptr [ebp+fffffed8]
'00495338    50                      push eax
'00495339    56                      push esi
'0049533a    ff5278                  call dword ptr [edx+78]
'0049533d    dbe2                    fnclex
'0049533f    3bc7                    cmp eax, edi
'00495341    0f8d4b030000            jge 495692
'00495347    6a78                    push 78
'00495349    684cfd4200              push 0042fd4c
'0049534e    56                      push esi
'0049534f    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00495350    8b3580104000            mov esi, dword ptr [00401080]
'00495356    ffd6                    call esi
'00495358    e93b030000              jmp 495698

' *** Reference to "rtcErrObj"
'0049535d    8b1d6c124000            mov ebx, dword ptr [0040126c]
'00495363    ffd3                    call ebx
'00495365    50                      push eax
'00495366    8d55c4                  lea edx, dword ptr [ebp-3c]
'00495369    52                      push edx

' *** Reference to "__vbaObjSet"
'0049536a    8b3db4104000            mov edi, dword ptr [004010b4]
'00495370    ffd7                    call edi
Set var_9 = Err
'00495372    8bf0                    mov esi, eax
'00495374    8b06                    mov eax, dword ptr [esi]
'00495376    8d4de0                  lea ecx, dword ptr [ebp-20]
'00495379    51                      push ecx
'0049537a    56                      push esi
'0049537b    ff502c                  call dword ptr [eax+2c]
var_3 = var_9.Description
'0049537e    dbe2                    fnclex
'00495380    85c0                    test ax, ax
'00495382    7d0f                    jge 495393
'00495384    6a2c                    push 2c
'00495386    685c084300              push 0043085c
'0049538b    56                      push esi
'0049538c    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0049538d    ff1580104000            call dword ptr [00401080]
'00495393    ffd3                    call ebx
'00495395    50                      push eax
'00495396    8d55c0                  lea edx, dword ptr [ebp-40]
'00495399    52                      push edx
'0049539a    ffd7                    call edi
Set var_5 = Err
'0049539c    8bf0                    mov esi, eax
'0049539e    8b06                    mov eax, dword ptr [esi]
'004953a0    8d8dd8feffff            lea ecx, dword ptr [ebp+fffffed8]
'004953a6    51                      push ecx
'004953a7    56                      push esi
'004953a8    ff501c                  call dword ptr [eax+1c]
var_25 = var_5.Number
'004953ab    dbe2                    fnclex
'004953ad    85c0                    test ax, ax
'004953af    7d0f                    jge 4953c0
'004953b1    6a1c                    push 1c
'004953b3    685c084300              push 0043085c
'004953b8    56                      push esi
'004953b9    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'004953ba    ff1580104000            call dword ptr [00401080]
'004953c0    b904000280              mov ecx, 80020004
'004953c5    894d98                  mov dword ptr [ebp-68], ecx
'004953c8    b80a000000              mov eax, 0000000a
'004953cd    894590                  mov dword ptr [ebp-70], eax
'004953d0    894da8                  mov dword ptr [ebp-58], ecx
'004953d3    8945a0                  mov dword ptr [ebp-60], eax
'004953d6    c78528ffffff24b07200    mov dword ptr [ebp+ffffff28], 0072b024
'004953e0    c78520ffffff08400000    mov dword ptr [ebp+ffffff20], 00004008
'004953ea    6814084300              push 00430814
'004953ef    8b55e0                  mov edx, dword ptr [ebp-20]
'004953f2    52                      push edx

' *** Reference to "__vbaStrCat"
'004953f3    8b3570104000            mov esi, dword ptr [00401070]
'004953f9    ffd6                    call esi
var_11 = ("L'erreur suivante s'est produite : ") & (var_3)
'004953fb    8bd0                    mov edx, eax
'004953fd    8d4ddc                  lea ecx, dword ptr [ebp-24]

' *** Reference to "__vbaStrMove"
'00495400    8b3dd0124000            mov edi, dword ptr [004012d0]
'00495406    ffd7                    call edi
'00495408    50                      push eax
'00495409    6870084300              push 00430870
'0049540e    ffd6                    call esi
var_12 = (var_11) & (vbCrLf)
'00495410    8bd0                    mov edx, eax
'00495412    8d4dd8                  lea ecx, dword ptr [ebp-28]
'00495415    ffd7                    call edi
'00495417    50                      push eax
'00495418    6870084300              push 00430870
'0049541d    ffd6                    call esi
var_13 = (var_12) & (vbCrLf)
'0049541f    8bd0                    mov edx, eax
'00495421    8d4dd4                  lea ecx, dword ptr [ebp-2c]
'00495424    ffd7                    call edi
'00495426    50                      push eax
'00495427    6880084300              push 00430880
'0049542c    ffd6                    call esi
var_14 = (var_13) & (" numéro d'erreur (")
'0049542e    8bd0                    mov edx, eax
'00495430    8d4dd0                  lea ecx, dword ptr [ebp-30]
'00495433    ffd7                    call edi
'00495435    50                      push eax
'00495436    8b85d8feffff            mov eax, dword ptr [ebp+fffffed8]
'0049543c    50                      push eax

' *** Reference to "__vbaStrI4"
'0049543d    ff1520104000            call dword ptr [00401020]
'00495443    8bd0                    mov edx, eax
'00495445    8d4dcc                  lea ecx, dword ptr [ebp-34]
'00495448    ffd7                    call edi
'0049544a    50                      push eax
'0049544b    ffd6                    call esi
var_15 = (var_14) & (CStr(var_25))
'0049544d    8bd0                    mov edx, eax
'0049544f    8d4dc8                  lea ecx, dword ptr [ebp-38]
'00495452    ffd7                    call edi
'00495454    50                      push eax
'00495455    68ac084300              push 004308ac
'0049545a    ffd6                    call esi
var_16 = (var_15) & (" )")
'0049545c    8945b8                  mov dword ptr [ebp-48], eax
'0049545f    bf08000000              mov edi, 00000008
'00495464    897db0                  mov dword ptr [ebp-50], edi
'00495467    8d4d90                  lea ecx, dword ptr [ebp-70]
'0049546a    51                      push ecx
'0049546b    8d55a0                  lea edx, dword ptr [ebp-60]
'0049546e    52                      push edx
'0049546f    8d8520ffffff            lea eax, dword ptr [ebp+ffffff20]
'00495475    50                      push eax
'00495476    6a10                    push 10
'00495478    8d4db0                  lea ecx, dword ptr [ebp-50]
'0049547b    51                      push ecx

' *** Reference to "rtcMsgBox"
'0049547c    ff15b8104000            call dword ptr [004010b8]
var_17 = MsgBox(var_16, 16, vbNullString)
'00495482    8d55c8                  lea edx, dword ptr [ebp-38]
'00495485    52                      push edx
'00495486    8d45cc                  lea eax, dword ptr [ebp-34]
'00495489    50                      push eax
'0049548a    8d4dd0                  lea ecx, dword ptr [ebp-30]
'0049548d    51                      push ecx
'0049548e    8d55d4                  lea edx, dword ptr [ebp-2c]
'00495491    52                      push edx
'00495492    8d45d8                  lea eax, dword ptr [ebp-28]
'00495495    50                      push eax
'00495496    8d4ddc                  lea ecx, dword ptr [ebp-24]
'00495499    51                      push ecx
'0049549a    8d55e0                  lea edx, dword ptr [ebp-20]
'0049549d    52                      push edx
'0049549e    6a07                    push 07

' *** Reference to "__vbaFreeStrList"
'004954a0    ff155c124000            call dword ptr [0040125c]
'#Cleanup( 0, -4512, -4516, -4520, -4524, -4528, -4532)
'004954a6    8d45c0                  lea eax, dword ptr [ebp-40]
'004954a9    50                      push eax
'004954aa    8d4dc4                  lea ecx, dword ptr [ebp-3c]
'004954ad    51                      push ecx
'004954ae    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'004954b0    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_9, var_5)
'004954b6    8d5590                  lea edx, dword ptr [ebp-70]
'004954b9    52                      push edx
'004954ba    8d45a0                  lea eax, dword ptr [ebp-60]
'004954bd    50                      push eax
'004954be    8d4db0                  lea ecx, dword ptr [ebp-50]
'004954c1    51                      push ecx
'004954c2    6a03                    push 03

' *** Reference to "__vbaFreeVarList"
'004954c4    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_6, var_7, var_8)
'004954ca    83c43c                  add esp, 3c
'004954cd    8d55b0                  lea edx, dword ptr [ebp-50]
'004954d0    52                      push edx

' *** Reference to "rtcGetPresentDate"
'004954d1    ff15f4124000            call dword ptr [004012f4]
var_16 = Now()
'004954d7    c78528ffffffb8084300    mov dword ptr [ebp+ffffff28], 004308b8
'004954e1    89bd20ffffff            mov dword ptr [ebp+ffffff20], edi
'004954e7    8d9520ffffff            lea edx, dword ptr [ebp+ffffff20]
'004954ed    8d4da0                  lea ecx, dword ptr [ebp-60]

' *** Reference to "__vbaVarDup"
'004954f0    ff1598124000            call dword ptr [00401298]
var_7 = ("dd/MM/yyyy hh:mm:ss")
'004954f6    6a01                    push 01
'004954f8    6a01                    push 01
'004954fa    8d45a0                  lea eax, dword ptr [ebp-60]
'004954fd    50                      push eax
'004954fe    8d4db0                  lea ecx, dword ptr [ebp-50]
'00495501    51                      push ecx
'00495502    8d5590                  lea edx, dword ptr [ebp-70]
'00495505    52                      push edx

' *** Reference to "rtcVarFromFormatVar"
'00495506    ff1574104000            call dword ptr [00401074]
'0049550c    ffd3                    call ebx
'0049550e    50                      push eax
'0049550f    8d45c4                  lea eax, dword ptr [ebp-3c]
'00495512    50                      push eax

' *** Reference to "__vbaObjSet"
'00495513    ff15b4104000            call dword ptr [004010b4]
Set var_9 = Err
'00495519    8bf0                    mov esi, eax
'0049551b    8b0e                    mov ecx, dword ptr [esi]
'0049551d    8d95d8feffff            lea edx, dword ptr [ebp+fffffed8]
'00495523    52                      push edx
'00495524    56                      push esi
'00495525    ff511c                  call dword ptr [ecx+1c]
var_25 = var_9.Number
'00495528    dbe2                    fnclex
'0049552a    85c0                    test ax, ax
'0049552c    7d0f                    jge 49553d
'0049552e    6a1c                    push 1c
'00495530    685c084300              push 0043085c
'00495535    56                      push esi
'00495536    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00495537    ff1580104000            call dword ptr [00401080]
'0049553d    ffd3                    call ebx
'0049553f    50                      push eax
'00495540    8d45c0                  lea eax, dword ptr [ebp-40]
'00495543    50                      push eax

' *** Reference to "__vbaObjSet"
'00495544    ff15b4104000            call dword ptr [004010b4]
Set var_5 = Err
'0049554a    8bf0                    mov esi, eax
'0049554c    8b0e                    mov ecx, dword ptr [esi]
'0049554e    8d55e0                  lea edx, dword ptr [ebp-20]
'00495551    52                      push edx
'00495552    56                      push esi
'00495553    ff512c                  call dword ptr [ecx+2c]
var_3 = var_5.Description
'00495556    dbe2                    fnclex
'00495558    85c0                    test ax, ax
'0049555a    7d0f                    jge 49556b
'0049555c    6a2c                    push 2c
'0049555e    685c084300              push 0043085c
'00495563    56                      push esi
'00495564    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00495565    ff1580104000            call dword ptr [00401080]
'0049556b    c78518ffffffe4084300    mov dword ptr [ebp+ffffff18], 004308e4
'00495575    89bd10ffffff            mov dword ptr [ebp+ffffff10], edi
'0049557b    8b85d8feffff            mov eax, dword ptr [ebp+fffffed8]
'00495581    898508ffffff            mov dword ptr [ebp+ffffff08], eax
'00495587    c78500ffffff03000000    mov dword ptr [ebp+ffffff00], 00000003
'00495591    c785f8feffff08094300    mov dword ptr [ebp+fffffef8], 00430908
'0049559b    89bdf0feffff            mov dword ptr [ebp+fffffef0], edi
'004955a1    8b45e0                  mov eax, dword ptr [ebp-20]
'004955a4    c745e000000000          mov dword ptr [ebp-20], 00000000
'004955ab    898558ffffff            mov dword ptr [ebp+ffffff58], eax
'004955b1    89bd50ffffff            mov dword ptr [ebp+ffffff50], edi
'004955b7    c785e8feffffd0224300    mov dword ptr [ebp+fffffee8], 004322d0
'004955c1    89bde0feffff            mov dword ptr [ebp+fffffee0], edi
'004955c7    8d4d90                  lea ecx, dword ptr [ebp-70]
'004955ca    51                      push ecx
'004955cb    8d9510ffffff            lea edx, dword ptr [ebp+ffffff10]
'004955d1    52                      push edx
'004955d2    8d4580                  lea eax, dword ptr [ebp-80]
'004955d5    50                      push eax

' *** Reference to "__vbaVarCat"
'004955d6    8b3508124000            mov esi, dword ptr [00401208]
'004955dc    ffd6                    call esi
'004955de    50                      push eax
'004955df    8d8d00ffffff            lea ecx, dword ptr [ebp+ffffff00]
'004955e5    51                      push ecx
'004955e6    8d9570ffffff            lea edx, dword ptr [ebp+ffffff70]
'004955ec    52                      push edx
'004955ed    ffd6                    call esi
'004955ef    50                      push eax
'004955f0    8d85f0feffff            lea eax, dword ptr [ebp+fffffef0]
'004955f6    50                      push eax
'004955f7    8d8d60ffffff            lea ecx, dword ptr [ebp+ffffff60]
'004955fd    51                      push ecx
'004955fe    ffd6                    call esi
'00495600    50                      push eax
'00495601    8d9550ffffff            lea edx, dword ptr [ebp+ffffff50]
'00495607    52                      push edx
'00495608    8d8540ffffff            lea eax, dword ptr [ebp+ffffff40]
'0049560e    50                      push eax
'0049560f    ffd6                    call esi
'00495611    50                      push eax
'00495612    8d8de0feffff            lea ecx, dword ptr [ebp+fffffee0]
'00495618    51                      push ecx
'00495619    8d9530ffffff            lea edx, dword ptr [ebp+ffffff30]
'0049561f    52                      push edx
'00495620    ffd6                    call esi
'00495622    50                      push eax
'00495623    33c0                    xor eax, eax
'00495625    66a12ab07200            mov eax, dword ptr [0072b02a]
'0049562b    50                      push eax
'0049562c    6884094300              push 00430984

' *** Reference to "__vbaPrintFile"
'00495631    ff15b8114000            call dword ptr [004011b8]
Print #0, 
'00495637    8d4dc0                  lea ecx, dword ptr [ebp-40]
'0049563a    51                      push ecx
'0049563b    8d55c4                  lea edx, dword ptr [ebp-3c]
'0049563e    52                      push edx
'0049563f    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'00495641    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_9, var_5)
'00495647    8d8530ffffff            lea eax, dword ptr [ebp+ffffff30]
'0049564d    50                      push eax
'0049564e    8d8d40ffffff            lea ecx, dword ptr [ebp+ffffff40]
'00495654    51                      push ecx
'00495655    8d9550ffffff            lea edx, dword ptr [ebp+ffffff50]
'0049565b    52                      push edx
'0049565c    8d8560ffffff            lea eax, dword ptr [ebp+ffffff60]
'00495662    50                      push eax
'00495663    8d8d70ffffff            lea ecx, dword ptr [ebp+ffffff70]
'00495669    51                      push ecx
'0049566a    8d5580                  lea edx, dword ptr [ebp-80]
'0049566d    52                      push edx
'0049566e    8d4590                  lea eax, dword ptr [ebp-70]
'00495671    50                      push eax
'00495672    8d4da0                  lea ecx, dword ptr [ebp-60]
'00495675    51                      push ecx
'00495676    8d55b0                  lea edx, dword ptr [ebp-50]
'00495679    52                      push edx
'0049567a    6a09                    push 09

' *** Reference to "__vbaFreeVarList"
'0049567c    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_6, var_7, var_8, var_18, var_19, var_20, var_21, var_22, var_23)
'00495682    83c440                  add esp, 40
'00495685    6a00                    push 00

' *** Reference to "__vbaResume"
'00495687    ff1568104000            call dword ptr [00401068]
'0049568d    e9a2000000              jmp 495734
Resume handler_495734

' *** Reference to "__vbaHresultCheckObj"
'00495692    8b3580104000            mov esi, dword ptr [00401080]
'00495698    d985d8feffff            fld dword ptr [ebp+fffffed8]
'0049569e    ffd3                    call ebx
'004956a0    8985dcfeffff            mov dword ptr [ebp+fffffedc], eax
'004956a6    ba3c224300              mov edx, 0043223c
'004956ab    8d4dd8                  lea ecx, dword ptr [ebp-28]

' *** Reference to "__vbaStrCopy"
'004956ae    ff1548124000            call dword ptr [00401248]
var_45 = ("POS_Y")
'004956b4    8b55e0                  mov edx, dword ptr [ebp-20]
'004956b7    897de0                  mov dword ptr [ebp-20], edi
'004956ba    8d4ddc                  lea ecx, dword ptr [ebp-24]

' *** Reference to "__vbaStrMove"
'004956bd    ff15d0124000            call dword ptr [004012d0]
'004956c3    8d8ddcfeffff            lea ecx, dword ptr [ebp+fffffedc]
'004956c9    51                      push ecx
'004956ca    8d55d8                  lea edx, dword ptr [ebp-28]
'004956cd    52                      push edx
'004956ce    8d45dc                  lea eax, dword ptr [ebp-24]
'004956d1    50                      push eax
'004956d2    e899f10500              call 4f4870
Call sub_4F4870()
'004956d7    8d4dd8                  lea ecx, dword ptr [ebp-28]
'004956da    51                      push ecx
'004956db    8d55dc                  lea edx, dword ptr [ebp-24]
'004956de    52                      push edx
'004956df    6a02                    push 02

' *** Reference to "__vbaFreeStrList"
'004956e1    ff155c124000            call dword ptr [0040125c]
'#Cleanup( var_3, var_45)
'004956e7    83c40c                  add esp, 0c
'004956ea    a148b07200              mov ax, word ptr [0072b048]
'004956ef    8b08                    mov ecx, dword ptr [eax]
'004956f1    50                      push eax
'004956f2    ff5158                  call dword ptr [ecx+58]
'004956f5    dbe2                    fnclex
'004956f7    3bc7                    cmp eax, edi
'004956f9    7d11                    jge 49570c

If (0 < 8) Then
'004956fb    6a58                    push 58
'004956fd    68301f4300              push 00431f30
'00495702    8b1548b07200            mov edx, dword ptr [0072b048]
'00495708    52                      push edx
'00495709    50                      push eax
'0049570a    ffd6                    call esi
    
End If
'0049570c    a14cb07200              mov ax, word ptr [0072b04c]
'00495711    8b08                    mov ecx, dword ptr [eax]
'00495713    50                      push eax
'00495714    ff5158                  call dword ptr [ecx+58]
'00495717    dbe2                    fnclex
'00495719    3bc7                    cmp eax, edi
'0049571b    7d11                    jge 49572e

If (0 < 8) Then
'0049571d    6a58                    push 58
'0049571f    68301f4300              push 00431f30
'00495724    8b154cb07200            mov edx, dword ptr [0072b04c]
'0049572a    52                      push edx
'0049572b    50                      push eax
'0049572c    ffd6                    call esi
    
End If

' *** Reference to "__vbaEnd"
'0049572e    ff1544104000            call dword ptr [00401044]
End

' *** Reference to "__vbaExitProc"
'00495734    ff15a0104000            call dword ptr [004010a0]
'0049573a    9b                      fwait
'0049573b    68b6574900              push 004957b6
'00495740    eb73                    jmp 4957b5
'00495742    8d45c8                  lea eax, dword ptr [ebp-38]
'00495745    50                      push eax
'00495746    8d4dcc                  lea ecx, dword ptr [ebp-34]
'00495749    51                      push ecx
'0049574a    8d55d0                  lea edx, dword ptr [ebp-30]
'0049574d    52                      push edx
'0049574e    8d45d4                  lea eax, dword ptr [ebp-2c]
'00495751    50                      push eax
'00495752    8d4dd8                  lea ecx, dword ptr [ebp-28]
'00495755    51                      push ecx
'00495756    8d55dc                  lea edx, dword ptr [ebp-24]
'00495759    52                      push edx
'0049575a    8d45e0                  lea eax, dword ptr [ebp-20]
'0049575d    50                      push eax
'0049575e    6a07                    push 07

' *** Reference to "__vbaFreeStrList"
'00495760    ff155c124000            call dword ptr [0040125c]
'#Cleanup( 8, var_3, var_45, -4520, -4524, -4528, -4532)
'00495766    8d4dc0                  lea ecx, dword ptr [ebp-40]
'00495769    51                      push ecx
'0049576a    8d55c4                  lea edx, dword ptr [ebp-3c]
'0049576d    52                      push edx
'0049576e    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'00495770    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_9, var_5)
'00495776    8d8530ffffff            lea eax, dword ptr [ebp+ffffff30]
'0049577c    50                      push eax
'0049577d    8d8d40ffffff            lea ecx, dword ptr [ebp+ffffff40]
'00495783    51                      push ecx
'00495784    8d9550ffffff            lea edx, dword ptr [ebp+ffffff50]
'0049578a    52                      push edx
'0049578b    8d8560ffffff            lea eax, dword ptr [ebp+ffffff60]
'00495791    50                      push eax
'00495792    8d8d70ffffff            lea ecx, dword ptr [ebp+ffffff70]
'00495798    51                      push ecx
'00495799    8d5580                  lea edx, dword ptr [ebp-80]
'0049579c    52                      push edx
'0049579d    8d4590                  lea eax, dword ptr [ebp-70]
'004957a0    50                      push eax
'004957a1    8d4da0                  lea ecx, dword ptr [ebp-60]
'004957a4    51                      push ecx
'004957a5    8d55b0                  lea edx, dword ptr [ebp-50]
'004957a8    52                      push edx
'004957a9    6a09                    push 09

' *** Reference to "__vbaFreeVarList"
'004957ab    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_6, var_7, var_8, var_18, var_19, var_20, var_21, var_22, var_23)
'004957b1    83c454                  add esp, 54
'004957b4    c3                      ret
'004957b5    c3                      ret
'004957b6    8b4508                  mov eax, dword ptr [ebp+08]
'004957b9    8b08                    mov ecx, dword ptr [eax]
'004957bb    50                      push eax
'004957bc    ff5108                  call dword ptr [ecx+08]
'004957bf    8b45f4                  mov eax, dword ptr [ebp-0c]
'004957c2    8b4de4                  mov ecx, dword ptr [ebp-1c]
    'Reference to '__except_list'
'004957c5    64890d00000000          mov dword ptr fs:[00000000], ecx
'004957cc    5f                      pop edi
'004957cd    5e                      pop esi
'004957ce    5b                      pop ebx
'004957cf    8be5                    mov esp, ebp
'004957d1    5d                      pop ebp
'004957d2    c20800                  ret 0008


End Sub


Private Sub MDIForm_OLEDragDrop(Data as DataObject, Effect as Long, Button as Integer, Shift as Integer, X as Single, Y as Single)
'00495090    55                      push ebp
'00495091    8bec                    mov ebp, esp
'00495093    83ec0c                  sub esp, 0c

' *** Reference to "__vbaExceptHandler"
'00495096    6866724000              push 00407266
'0049509b    64a100000000            mov ax, word ptr fs:[00000000]
'004950a1    50                      push eax
    'Reference to '__except_list'
'004950a2    64892500000000          mov dword ptr fs:[00000000], esp
'004950a9    83ec34                  sub esp, 34
'004950ac    53                      push ebx
'004950ad    56                      push esi
'004950ae    57                      push edi
'004950af    8965f4                  mov dword ptr [ebp-0c], esp
'004950b2    c745f800154000          mov dword ptr [ebp-08], 00401500
'004950b9    8b7d08                  mov edi, dword ptr [ebp+08]
'004950bc    8bc7                    mov eax, edi
'004950be    83e001                  and eax, 01
'004950c1    8945fc                  mov dword ptr [ebp-04], eax
'004950c4    83e7fe                  and edi, fffffffe
'004950c7    8b0f                    mov ecx, dword ptr [edi]
'004950c9    57                      push edi
'004950ca    897d08                  mov dword ptr [ebp+08], edi
'004950cd    ff5104                  call dword ptr [ecx+04]
'004950d0    8b550c                  mov edx, dword ptr [ebp+0c]
'004950d3    8b32                    mov esi, dword ptr [edx]
'004950d5    33db                    xor ebx, ebx
'004950d7    8d4dd8                  lea ecx, dword ptr [ebp-28]
'004950da    51                      push ecx
'004950db    895ddc                  mov dword ptr [ebp-24], ebx
'004950de    895dd8                  mov dword ptr [ebp-28], ebx
'004950e1    895dcc                  mov dword ptr [ebp-34], ebx
'004950e4    895dc8                  mov dword ptr [ebp-38], ebx
'004950e7    8b06                    mov eax, dword ptr [esi]
'004950e9    56                      push esi
'004950ea    ff502c                  call dword ptr [eax+2c]
'004950ed    dbe2                    fnclex
'004950ef    3bc3                    cmp eax, ebx
'004950f1    7d0f                    jge 495102

If (arg_0 < 0) Then
'004950f3    6a2c                    push 2c
'004950f5    68bc224300              push 004322bc
'004950fa    56                      push esi
'004950fb    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'004950fc    ff1580104000            call dword ptr [00401080]
    
End If
'00495102    8b45d8                  mov eax, dword ptr [ebp-28]
'00495105    50                      push eax
'00495106    8d55cc                  lea edx, dword ptr [ebp-34]
'00495109    52                      push edx
'0049510a    895dd8                  mov dword ptr [ebp-28], ebx

' *** Reference to "__vbaObjSet"
'0049510d    ff15b4104000            call dword ptr [004010b4]
Set var_43 = Nothing
'00495113    50                      push eax
'00495114    8d45dc                  lea eax, dword ptr [ebp-24]
'00495117    50                      push eax
'00495118    8d4dc8                  lea ecx, dword ptr [ebp-38]
'0049511b    51                      push ecx

' *** Reference to "__vbaForEachCollVar"
'0049511c    ff15e8104000            call dword ptr [004010e8]
For Each var_39 In var_43
'00495122    3bc3                    cmp eax, ebx
'00495124    7421                    je 495147
'00495126    8d55dc                  lea edx, dword ptr [ebp-24]
'00495129    52                      push edx

' *** Reference to "__vbaStrVarCopy"
'0049512a    ff15e0124000            call dword ptr [004012e0]
'00495130    8bd0                    mov edx, eax
'00495132    b93cb07200              mov ecx, 0072b03c

' *** Reference to "__vbaStrMove"
'00495137    ff15d0124000            call dword ptr [004012d0]
'0049513d    8d45c8                  lea eax, dword ptr [ebp-38]
'00495140    50                      push eax

' *** Reference to "__vbaExitEachColl"
'00495141    ff1534114000            call dword ptr [00401134]
    Exit For
'00495147    8b0f                    mov ecx, dword ptr [edi]
'00495149    683cb07200              push 0072b03c
'0049514e    57                      push edi
'0049514f    ff9154070000            call dword ptr [ecx+00000754]
'00495155    895dfc                  mov dword ptr [ebp-04], ebx
'00495158    6886514900              push 00495186
'0049515d    eb0a                    jmp 495169
'0049515f    8d4dd8                  lea ecx, dword ptr [ebp-28]

' *** Reference to "__vbaFreeObj"
'00495162    ff1508134000            call dword ptr [00401308]
    '#Cleanup(var_45)
'00495168    c3                      ret
'00495169    8d55c8                  lea edx, dword ptr [ebp-38]
'0049516c    52                      push edx
'0049516d    8d45cc                  lea eax, dword ptr [ebp-34]
'00495170    50                      push eax
'00495171    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'00495173    ff154c104000            call dword ptr [0040104c]
    '#Cleanup( var_43, var_46)
'00495179    83c40c                  add esp, 0c
'0049517c    8d4ddc                  lea ecx, dword ptr [ebp-24]

' *** Reference to "__vbaFreeVar"
'0049517f    ff1528104000            call dword ptr [00401028]
    '#Cleanup(var_39)
'00495185    c3                      ret
'00495186    8b4508                  mov eax, dword ptr [ebp+08]
'00495189    8b08                    mov ecx, dword ptr [eax]
'0049518b    50                      push eax
'0049518c    ff5108                  call dword ptr [ecx+08]
'0049518f    8b45fc                  mov eax, dword ptr [ebp-04]
'00495192    8b4dec                  mov ecx, dword ptr [ebp-14]
'00495195    5f                      pop edi
'00495196    5e                      pop esi
    'Reference to '__except_list'
'00495197    64890d00000000          mov dword ptr fs:[00000000], ecx
'0049519e    5b                      pop ebx
'0049519f    8be5                    mov esp, ebp
'004951a1    5d                      pop ebp
'004951a2    c21c00                  ret 001c


End Sub


'Event for IDM_GEST_OBJETS
Private Sub IDM_GEST_OBJETS_Click()
'004926f0    55                      push ebp
'004926f1    8bec                    mov ebp, esp
'004926f3    83ec0c                  sub esp, 0c

' *** Reference to "__vbaExceptHandler"
'004926f6    6866724000              push 00407266
'004926fb    64a100000000            mov ax, word ptr fs:[00000000]
'00492701    50                      push eax
    'Reference to '__except_list'
'00492702    64892500000000          mov dword ptr fs:[00000000], esp
'00492709    83ec30                  sub esp, 30
'0049270c    53                      push ebx
'0049270d    56                      push esi
'0049270e    57                      push edi
'0049270f    8965f4                  mov dword ptr [ebp-0c], esp
'00492712    c745f888134000          mov dword ptr [ebp-08], 00401388
'00492719    8b4508                  mov eax, dword ptr [ebp+08]
'0049271c    8bc8                    mov ecx, eax
'0049271e    83e101                  and ecx, 01
'00492721    894dfc                  mov dword ptr [ebp-04], ecx
'00492724    83e0fe                  and eax, fffffffe
'00492727    8b10                    mov edx, dword ptr [eax]
'00492729    50                      push eax
'0049272a    894508                  mov dword ptr [ebp+08], eax
'0049272d    ff5204                  call dword ptr [edx+04]
'00492730    a110b17200              mov ax, word ptr [0072b110]
'00492735    85c0                    test ax, ax
'00492737    7510                    jne 492749
'00492739    6810b17200              push 0072b110
'0049273e    6830584100              push 00415830

' *** Reference to "__vbaNew2"
'00492743    ff152c124000            call dword ptr [0040122c]
Dim var_35 As New frmGestobj
'00492749    8b3510b17200            mov esi, dword ptr [0072b110]
'0049274f    8b3e                    mov edi, dword ptr [esi]
'00492751    83ec10                  sub esp, 10
'00492754    8bdc                    mov ebx, esp
'00492756    b90a000000              mov ecx, 0000000a
'0049275b    890b                    mov dword ptr [ebx], ecx
'0049275d    8b4dd0                  mov ecx, dword ptr [ebp-30]
'00492760    894b04                  mov dword ptr [ebx+04], ecx
'00492763    b804000280              mov eax, 80020004
'00492768    894308                  mov dword ptr [ebx+08], eax
'0049276b    8b45d8                  mov eax, dword ptr [ebp-28]
'0049276e    89430c                  mov dword ptr [ebx+0c], eax
'00492771    83ec10                  sub esp, 10
'00492774    8bcc                    mov ecx, esp
'00492776    b803000000              mov eax, 00000003
'0049277b    8901                    mov dword ptr [ecx], eax
'0049277d    8b45e0                  mov eax, dword ptr [ebp-20]
'00492780    894104                  mov dword ptr [ecx+04], eax
'00492783    ba01000000              mov edx, 00000001
'00492788    895108                  mov dword ptr [ecx+08], edx
'0049278b    8b55e8                  mov edx, dword ptr [ebp-18]
'0049278e    56                      push esi
'0049278f    89510c                  mov dword ptr [ecx+0c], edx
'00492792    ff97b0020000            call dword ptr [edi+000002b0]
'00492798    dbe2                    fnclex
'0049279a    85c0                    test ax, ax
'0049279c    7d12                    jge 4927b0
'0049279e    68b0020000              push 000002b0
'004927a3    68a0164300              push 004316a0
'004927a8    56                      push esi
'004927a9    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'004927aa    ff1580104000            call dword ptr [00401080]
'004927b0    c745fc00000000          mov dword ptr [ebp-04], 00000000
'004927b7    8b4508                  mov eax, dword ptr [ebp+08]
'004927ba    8b08                    mov ecx, dword ptr [eax]
'004927bc    50                      push eax
'004927bd    ff5108                  call dword ptr [ecx+08]
'004927c0    8b45fc                  mov eax, dword ptr [ebp-04]
'004927c3    8b4dec                  mov ecx, dword ptr [ebp-14]
'004927c6    5f                      pop edi
'004927c7    5e                      pop esi
    'Reference to '__except_list'
'004927c8    64890d00000000          mov dword ptr fs:[00000000], ecx
'004927cf    5b                      pop ebx
'004927d0    8be5                    mov esp, ebp
'004927d2    5d                      pop ebp
'004927d3    c20400                  ret 0004


End Sub


'Event for IDM_GENERATEUR_VILLE
Private Sub IDM_GENERATEUR_VILLE_Click()
'00492600    55                      push ebp
'00492601    8bec                    mov ebp, esp
'00492603    83ec0c                  sub esp, 0c

' *** Reference to "__vbaExceptHandler"
'00492606    6866724000              push 00407266
'0049260b    64a100000000            mov ax, word ptr fs:[00000000]
'00492611    50                      push eax
    'Reference to '__except_list'
'00492612    64892500000000          mov dword ptr fs:[00000000], esp
'00492619    83ec30                  sub esp, 30
'0049261c    53                      push ebx
'0049261d    56                      push esi
'0049261e    57                      push edi
'0049261f    8965f4                  mov dword ptr [ebp-0c], esp
'00492622    c745f880134000          mov dword ptr [ebp-08], 00401380
'00492629    8b4508                  mov eax, dword ptr [ebp+08]
'0049262c    8bc8                    mov ecx, eax
'0049262e    83e101                  and ecx, 01
'00492631    894dfc                  mov dword ptr [ebp-04], ecx
'00492634    83e0fe                  and eax, fffffffe
'00492637    8b10                    mov edx, dword ptr [eax]
'00492639    50                      push eax
'0049263a    894508                  mov dword ptr [ebp+08], eax
'0049263d    ff5204                  call dword ptr [edx+04]
'00492640    a1fcb07200              mov ax, word ptr [0072b0fc]
'00492645    85c0                    test ax, ax
'00492647    7510                    jne 492659
'00492649    68fcb07200              push 0072b0fc
'0049264e    687cbb4000              push 0040bb7c

' *** Reference to "__vbaNew2"
'00492653    ff152c124000            call dword ptr [0040122c]
Dim var_34 As New frmGenerateurVille
'00492659    8b35fcb07200            mov esi, dword ptr [0072b0fc]
'0049265f    8b3e                    mov edi, dword ptr [esi]
'00492661    83ec10                  sub esp, 10
'00492664    8bdc                    mov ebx, esp
'00492666    b90a000000              mov ecx, 0000000a
'0049266b    890b                    mov dword ptr [ebx], ecx
'0049266d    894ddc                  mov dword ptr [ebp-24], ecx
'00492670    8b4dd0                  mov ecx, dword ptr [ebp-30]
'00492673    894b04                  mov dword ptr [ebx+04], ecx
'00492676    b804000280              mov eax, 80020004
'0049267b    894308                  mov dword ptr [ebx+08], eax
'0049267e    8bd0                    mov edx, eax
'00492680    8b45d8                  mov eax, dword ptr [ebp-28]
'00492683    89430c                  mov dword ptr [ebx+0c], eax
'00492686    8b45dc                  mov eax, dword ptr [ebp-24]
'00492689    83ec10                  sub esp, 10
'0049268c    8bcc                    mov ecx, esp
'0049268e    8901                    mov dword ptr [ecx], eax
'00492690    8b45e0                  mov eax, dword ptr [ebp-20]
'00492693    894104                  mov dword ptr [ecx+04], eax
'00492696    895108                  mov dword ptr [ecx+08], edx
'00492699    8b55e8                  mov edx, dword ptr [ebp-18]
'0049269c    56                      push esi
'0049269d    89510c                  mov dword ptr [ecx+0c], edx
'004926a0    ff97b0020000            call dword ptr [edi+000002b0]
'004926a6    dbe2                    fnclex
'004926a8    85c0                    test ax, ax
'004926aa    7d12                    jge 4926be
'004926ac    68b0020000              push 000002b0
'004926b1    6880154300              push 00431580
'004926b6    56                      push esi
'004926b7    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'004926b8    ff1580104000            call dword ptr [00401080]
'004926be    c745fc00000000          mov dword ptr [ebp-04], 00000000
'004926c5    8b4508                  mov eax, dword ptr [ebp+08]
'004926c8    8b08                    mov ecx, dword ptr [eax]
'004926ca    50                      push eax
'004926cb    ff5108                  call dword ptr [ecx+08]
'004926ce    8b45fc                  mov eax, dword ptr [ebp-04]
'004926d1    8b4dec                  mov ecx, dword ptr [ebp-14]
'004926d4    5f                      pop edi
'004926d5    5e                      pop esi
    'Reference to '__except_list'
'004926d6    64890d00000000          mov dword ptr fs:[00000000], ecx
'004926dd    5b                      pop ebx
'004926de    8be5                    mov esp, ebp
'004926e0    5d                      pop ebp
'004926e1    c20400                  ret 0004


End Sub


'Event for IDM_OUVRIR
Private Sub IDM_OUVRIR_Click()
'00493200    55                      push ebp
'00493201    8bec                    mov ebp, esp
'00493203    83ec18                  sub esp, 18

' *** Reference to "__vbaExceptHandler"
'00493206    6866724000              push 00407266
'0049320b    64a100000000            mov ax, word ptr fs:[00000000]
'00493211    50                      push eax
    'Reference to '__except_list'
'00493212    64892500000000          mov dword ptr fs:[00000000], esp
'00493219    b868000000              mov eax, 00000068

' *** Reference to "__vbaChkstk"
'0049321e    e83d40f7ff              call 407260
'00493223    53                      push ebx
'00493224    56                      push esi
'00493225    57                      push edi
'00493226    8965e8                  mov dword ptr [ebp-18], esp
'00493229    c745ece8134000          mov dword ptr [ebp-14], 004013e8
'00493230    8b4508                  mov eax, dword ptr [ebp+08]
'00493233    83e001                  and eax, 01
'00493236    8945f0                  mov dword ptr [ebp-10], eax
'00493239    8b4d08                  mov ecx, dword ptr [ebp+08]
'0049323c    83e1fe                  and ecx, fffffffe
'0049323f    894d08                  mov dword ptr [ebp+08], ecx
'00493242    c745f400000000          mov dword ptr [ebp-0c], 00000000
'00493249    8b5508                  mov edx, dword ptr [ebp+08]
'0049324c    8b02                    mov eax, dword ptr [edx]
'0049324e    8b4d08                  mov ecx, dword ptr [ebp+08]
'00493251    51                      push ecx
'00493252    ff5004                  call dword ptr [eax+04]
'00493255    c745fc01000000          mov dword ptr [ebp-04], 00000001
'0049325c    c745fc02000000          mov dword ptr [ebp-04], 00000002
'00493263    0fbf1528b07200          movsx edx, word ptr [0072b028]
'0049326a    85d2                    test dx, dx
'0049326c    750f                    jne 49327d
'0049326e    c745fc03000000          mov dword ptr [ebp-04], 00000003
'00493275    6aff                    push ffffffff

' *** Reference to "__vbaOnError"
'00493277    ff15b0104000            call dword ptr [004010b0]
On Error Resume Next
'0049327d    c745fc05000000          mov dword ptr [ebp-04], 00000005
'00493284    8d45c0                  lea eax, dword ptr [ebp-40]
'00493287    50                      push eax
'00493288    8b4d08                  mov ecx, dword ptr [ebp+08]
'0049328b    8b11                    mov edx, dword ptr [ecx]
'0049328d    8b4508                  mov eax, dword ptr [ebp+08]
'00493290    50                      push eax
'00493291    ff5258                  call dword ptr [edx+58]
'00493294    dbe2                    fnclex
'00493296    8945b0                  mov dword ptr [ebp-50], eax
'00493299    837db000                cmp dword ptr [ebp-50], 00
'0049329d    7d1a                    jge 4932b9

If (Me < 0) Then
'0049329f    6a58                    push 58
'004932a1    684cfd4200              push 0042fd4c
'004932a6    8b4d08                  mov ecx, dword ptr [ebp+08]
'004932a9    51                      push ecx
'004932aa    8b55b0                  mov edx, dword ptr [ebp-50]
'004932ad    52                      push edx

' *** Reference to "__vbaHresultCheckObj"
'004932ae    ff1580104000            call dword ptr [00401080]
'004932b4    894584                  mov dword ptr [ebp-7c], eax
'004932b7    eb07                    jmp 4932c0
    
End If
'004932b9    c7458400000000          mov dword ptr [ebp-7c], 00000000
'004932c0    833d24be720000          cmp dword ptr [0072be24], 00
'004932c7    7519                    jne 4932e2
'004932c9    6824be7200              push 0072be24
'004932ce    6870004300              push 00430070

' *** Reference to "__vbaNew2"
'004932d3    ff152c124000            call dword ptr [0040122c]
Dim var_2 As New Global
'004932d9    c7458024be7200          mov dword ptr [ebp-80], 0072be24
'004932e0    eb07                    jmp 4932e9
'004932e2    c7458024be7200          mov dword ptr [ebp-80], 0072be24
'004932e9    8b4580                  mov eax, dword ptr [ebp-80]
'004932ec    8b08                    mov ecx, dword ptr [eax]
'004932ee    894dac                  mov dword ptr [ebp-54], ecx
'004932f1    8d55c4                  lea edx, dword ptr [ebp-3c]
'004932f4    52                      push edx
'004932f5    8b45ac                  mov eax, dword ptr [ebp-54]
'004932f8    8b08                    mov ecx, dword ptr [eax]
'004932fa    8b55ac                  mov edx, dword ptr [ebp-54]
'004932fd    52                      push edx
'004932fe    ff5114                  call dword ptr [ecx+14]
Set var_9 = var_2.App
'00493301    dbe2                    fnclex
'00493303    8945a8                  mov dword ptr [ebp-58], eax
'00493306    837da800                cmp dword ptr [ebp-58], 00
'0049330a    7d1d                    jge 493329

If (var_2.App < 0) Then
'0049330c    6a14                    push 14
'0049330e    6860004300              push 00430060
'00493313    8b45ac                  mov eax, dword ptr [ebp-54]
'00493316    50                      push eax
'00493317    8b4da8                  mov ecx, dword ptr [ebp-58]
'0049331a    51                      push ecx

' *** Reference to "__vbaHresultCheckObj"
'0049331b    ff1580104000            call dword ptr [00401080]
'00493321    89857cffffff            mov dword ptr [ebp+ffffff7c], eax
'00493327    eb0a                    jmp 493333
    
End If
'00493329    c7857cffffff00000000    mov dword ptr [ebp+ffffff7c], 00000000
'00493333    8b55c4                  mov edx, dword ptr [ebp-3c]
'00493336    8955a4                  mov dword ptr [ebp-5c], edx
'00493339    8d45dc                  lea eax, dword ptr [ebp-24]
'0049333c    50                      push eax
'0049333d    8b4da4                  mov ecx, dword ptr [ebp-5c]
'00493340    8b11                    mov edx, dword ptr [ecx]
'00493342    8b45a4                  mov eax, dword ptr [ebp-5c]
'00493345    50                      push eax
'00493346    ff5250                  call dword ptr [edx+50]
var_39 = var_9.Path
'00493349    dbe2                    fnclex
'0049334b    8945a0                  mov dword ptr [ebp-60], eax
'0049334e    837da000                cmp dword ptr [ebp-60], 00
'00493352    7d1d                    jge 493371

If (0 < 0) Then
'00493354    6a50                    push 50
'00493356    682c1c4300              push 00431c2c
'0049335b    8b4da4                  mov ecx, dword ptr [ebp-5c]
'0049335e    51                      push ecx
'0049335f    8b55a0                  mov edx, dword ptr [ebp-60]
'00493362    52                      push edx

' *** Reference to "__vbaHresultCheckObj"
'00493363    ff1580104000            call dword ptr [00401080]
'00493369    898578ffffff            mov dword ptr [ebp+ffffff78], eax
'0049336f    eb0a                    jmp 49337b
    
End If
'00493371    c78578ffffff00000000    mov dword ptr [ebp+ffffff78], 00000000
'0049337b    bacc134300              mov edx, 004313cc
'00493380    8d4dc8                  lea ecx, dword ptr [ebp-38]

' *** Reference to "__vbaStrCopy"
'00493383    ff1548124000            call dword ptr [00401248]
var_46 = (vbNullChar)
'00493389    8b45dc                  mov eax, dword ptr [ebp-24]
'0049338c    894588                  mov dword ptr [ebp-78], eax
'0049338f    c745dc00000000          mov dword ptr [ebp-24], 00000000
'00493396    8b5588                  mov edx, dword ptr [ebp-78]
'00493399    8d4dcc                  lea ecx, dword ptr [ebp-34]

' *** Reference to "__vbaStrMove"
'0049339c    ff15d0124000            call dword ptr [004012d0]
'004933a2    bacc134300              mov edx, 004313cc
'004933a7    8d4dd0                  lea ecx, dword ptr [ebp-30]

' *** Reference to "__vbaStrCopy"
'004933aa    ff1548124000            call dword ptr [00401248]
var_4 = (vbNullChar)
'004933b0    ba741c4300              mov edx, 00431c74
'004933b5    8d4dd4                  lea ecx, dword ptr [ebp-2c]

' *** Reference to "__vbaStrCopy"
'004933b8    ff1548124000            call dword ptr [00401248]
var_44 = ("Ouvrir une base donnée")
'004933be    c745b400180800          mov dword ptr [ebp-4c], 00081800
'004933c5    c745b800000000          mov dword ptr [ebp-48], 00000000
'004933cc    ba401c4300              mov edx, 00431c40
'004933d1    8d4dd8                  lea ecx, dword ptr [ebp-28]

' *** Reference to "__vbaStrCopy"
'004933d4    ff1548124000            call dword ptr [00401248]
var_45 = ("Access97 (*.mdb)|*.mdb")
'004933da    8b4dc0                  mov ecx, dword ptr [ebp-40]
'004933dd    894dbc                  mov dword ptr [ebp-44], ecx
'004933e0    8d55c8                  lea edx, dword ptr [ebp-38]
'004933e3    52                      push edx
'004933e4    8d45cc                  lea eax, dword ptr [ebp-34]
'004933e7    50                      push eax
'004933e8    8d4dd0                  lea ecx, dword ptr [ebp-30]
'004933eb    51                      push ecx
'004933ec    8d55d4                  lea edx, dword ptr [ebp-2c]
'004933ef    52                      push edx
'004933f0    8d45b4                  lea eax, dword ptr [ebp-4c]
'004933f3    50                      push eax
'004933f4    8d4db8                  lea ecx, dword ptr [ebp-48]
'004933f7    51                      push ecx
'004933f8    8d55d8                  lea edx, dword ptr [ebp-28]
'004933fb    52                      push edx
'004933fc    8d45bc                  lea eax, dword ptr [ebp-44]
'004933ff    50                      push eax
'00493400    e84b2b0000              call 495f50
Call sub_495F50()
'00493405    8bd0                    mov edx, eax
'00493407    b93cb07200              mov ecx, 0072b03c

' *** Reference to "__vbaStrMove"
'0049340c    ff15d0124000            call dword ptr [004012d0]
'00493412    8d4dc8                  lea ecx, dword ptr [ebp-38]
'00493415    51                      push ecx
'00493416    8d55cc                  lea edx, dword ptr [ebp-34]
'00493419    52                      push edx
'0049341a    8d45d0                  lea eax, dword ptr [ebp-30]
'0049341d    50                      push eax
'0049341e    8d4dd4                  lea ecx, dword ptr [ebp-2c]
'00493421    51                      push ecx
'00493422    8d55d8                  lea edx, dword ptr [ebp-28]
'00493425    52                      push edx
'00493426    6a05                    push 05

' *** Reference to "__vbaFreeStrList"
'00493428    ff155c124000            call dword ptr [0040125c]
'#Cleanup( var_45, var_44, var_4, -4504, var_46)
'0049342e    83c418                  add esp, 18
'00493431    8d4dc4                  lea ecx, dword ptr [ebp-3c]

' *** Reference to "__vbaFreeObj"
'00493434    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_9)
'0049343a    c745fc06000000          mov dword ptr [ebp-04], 00000006
'00493441    8b153cb07200            mov edx, dword ptr [0072b03c]
'00493447    8d4ddc                  lea ecx, dword ptr [ebp-24]

' *** Reference to "__vbaStrCopy"
'0049344a    ff1548124000            call dword ptr [00401248]
var_39 = (0)
'00493450    8d45dc                  lea eax, dword ptr [ebp-24]
'00493453    50                      push eax
'00493454    8b4d08                  mov ecx, dword ptr [ebp+08]
'00493457    8b11                    mov edx, dword ptr [ecx]
'00493459    8b4508                  mov eax, dword ptr [ebp+08]
'0049345c    50                      push eax
'0049345d    ff9254070000            call dword ptr [edx+00000754]
'00493463    8d4ddc                  lea ecx, dword ptr [ebp-24]

' *** Reference to "__vbaFreeStr"
'00493466    ff150c134000            call dword ptr [0040130c]
'#Cleanup(var_39)
'0049346c    c745f000000000          mov dword ptr [ebp-10], 00000000
'00493473    68a8344900              push 004934a8
'00493478    eb2d                    jmp 4934a7
'0049347a    8d4dc8                  lea ecx, dword ptr [ebp-38]
'0049347d    51                      push ecx
'0049347e    8d55cc                  lea edx, dword ptr [ebp-34]
'00493481    52                      push edx
'00493482    8d45d0                  lea eax, dword ptr [ebp-30]
'00493485    50                      push eax
'00493486    8d4dd4                  lea ecx, dword ptr [ebp-2c]
'00493489    51                      push ecx
'0049348a    8d55d8                  lea edx, dword ptr [ebp-28]
'0049348d    52                      push edx
'0049348e    8d45dc                  lea eax, dword ptr [ebp-24]
'00493491    50                      push eax
'00493492    6a06                    push 06

' *** Reference to "__vbaFreeStrList"
'00493494    ff155c124000            call dword ptr [0040125c]
'#Cleanup( , var_45, var_44, var_4, -4504, var_46)
'0049349a    83c41c                  add esp, 1c
'0049349d    8d4dc4                  lea ecx, dword ptr [ebp-3c]

' *** Reference to "__vbaFreeObj"
'004934a0    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_9)
'004934a6    c3                      ret
'004934a7    c3                      ret
'004934a8    8b4d08                  mov ecx, dword ptr [ebp+08]
'004934ab    8b11                    mov edx, dword ptr [ecx]
'004934ad    8b4508                  mov eax, dword ptr [ebp+08]
'004934b0    50                      push eax
'004934b1    ff5208                  call dword ptr [edx+08]
'004934b4    8b45f0                  mov eax, dword ptr [ebp-10]
'004934b7    8b4de0                  mov ecx, dword ptr [ebp-20]
    'Reference to '__except_list'
'004934ba    64890d00000000          mov dword ptr fs:[00000000], ecx
'004934c1    5f                      pop edi
'004934c2    5e                      pop esi
'004934c3    5b                      pop ebx
'004934c4    8be5                    mov esp, ebp
'004934c6    5d                      pop ebp
'004934c7    c20400                  ret 0004


End Sub


'Event for IDM_DESCRIPTION_RACE
Private Sub IDM_DESCRIPTION_RACE_Click()
'004923d0    55                      push ebp
'004923d1    8bec                    mov ebp, esp
'004923d3    83ec0c                  sub esp, 0c

' *** Reference to "__vbaExceptHandler"
'004923d6    6866724000              push 00407266
'004923db    64a100000000            mov ax, word ptr fs:[00000000]
'004923e1    50                      push eax
    'Reference to '__except_list'
'004923e2    64892500000000          mov dword ptr fs:[00000000], esp
'004923e9    83ec30                  sub esp, 30
'004923ec    53                      push ebx
'004923ed    56                      push esi
'004923ee    57                      push edi
'004923ef    8965f4                  mov dword ptr [ebp-0c], esp
'004923f2    c745f870134000          mov dword ptr [ebp-08], 00401370
'004923f9    8b4508                  mov eax, dword ptr [ebp+08]
'004923fc    8bc8                    mov ecx, eax
'004923fe    83e101                  and ecx, 01
'00492401    894dfc                  mov dword ptr [ebp-04], ecx
'00492404    83e0fe                  and eax, fffffffe
'00492407    8b10                    mov edx, dword ptr [eax]
'00492409    50                      push eax
'0049240a    894508                  mov dword ptr [ebp+08], eax
'0049240d    ff5204                  call dword ptr [edx+04]
'00492410    a1d4b07200              mov ax, word ptr [0072b0d4]
'00492415    85c0                    test ax, ax
'00492417    7510                    jne 492429
'00492419    68d4b07200              push 0072b0d4
'0049241e    68207e4000              push 00407e20

' *** Reference to "__vbaNew2"
'00492423    ff152c124000            call dword ptr [0040122c]
Dim var_32 As New frmDescriptionRaces
'00492429    8b35d4b07200            mov esi, dword ptr [0072b0d4]
'0049242f    8b3e                    mov edi, dword ptr [esi]
'00492431    83ec10                  sub esp, 10
'00492434    8bdc                    mov ebx, esp
'00492436    b90a000000              mov ecx, 0000000a
'0049243b    890b                    mov dword ptr [ebx], ecx
'0049243d    894ddc                  mov dword ptr [ebp-24], ecx
'00492440    8b4dd0                  mov ecx, dword ptr [ebp-30]
'00492443    894b04                  mov dword ptr [ebx+04], ecx
'00492446    b804000280              mov eax, 80020004
'0049244b    894308                  mov dword ptr [ebx+08], eax
'0049244e    8bd0                    mov edx, eax
'00492450    8b45d8                  mov eax, dword ptr [ebp-28]
'00492453    89430c                  mov dword ptr [ebx+0c], eax
'00492456    8b45dc                  mov eax, dword ptr [ebp-24]
'00492459    83ec10                  sub esp, 10
'0049245c    8bcc                    mov ecx, esp
'0049245e    8901                    mov dword ptr [ecx], eax
'00492460    8b45e0                  mov eax, dword ptr [ebp-20]
'00492463    894104                  mov dword ptr [ecx+04], eax
'00492466    895108                  mov dword ptr [ecx+08], edx
'00492469    8b55e8                  mov edx, dword ptr [ebp-18]
'0049246c    56                      push esi
'0049246d    89510c                  mov dword ptr [ecx+0c], edx
'00492470    ff97b0020000            call dword ptr [edi+000002b0]
'00492476    dbe2                    fnclex
'00492478    85c0                    test ax, ax
'0049247a    7d12                    jge 49248e
'0049247c    68b0020000              push 000002b0
'00492481    6854134300              push 00431354
'00492486    56                      push esi
'00492487    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00492488    ff1580104000            call dword ptr [00401080]
'0049248e    c745fc00000000          mov dword ptr [ebp-04], 00000000
'00492495    8b4508                  mov eax, dword ptr [ebp+08]
'00492498    8b08                    mov ecx, dword ptr [eax]
'0049249a    50                      push eax
'0049249b    ff5108                  call dword ptr [ecx+08]
'0049249e    8b45fc                  mov eax, dword ptr [ebp-04]
'004924a1    8b4dec                  mov ecx, dword ptr [ebp-14]
'004924a4    5f                      pop edi
'004924a5    5e                      pop esi
    'Reference to '__except_list'
'004924a6    64890d00000000          mov dword ptr fs:[00000000], ecx
'004924ad    5b                      pop ebx
'004924ae    8be5                    mov esp, ebp
'004924b0    5d                      pop ebp
'004924b1    c20400                  ret 0004


End Sub


'Event for IDM_INSERTION_DONS
Private Sub IDM_INSERTION_DONS_Click()
'00492d20    55                      push ebp
'00492d21    8bec                    mov ebp, esp
'00492d23    83ec0c                  sub esp, 0c

' *** Reference to "__vbaExceptHandler"
'00492d26    6866724000              push 00407266
'00492d2b    64a100000000            mov ax, word ptr fs:[00000000]
'00492d31    50                      push eax
    'Reference to '__except_list'
'00492d32    64892500000000          mov dword ptr fs:[00000000], esp
'00492d39    83ec30                  sub esp, 30
'00492d3c    53                      push ebx
'00492d3d    56                      push esi
'00492d3e    57                      push edi
'00492d3f    8965f4                  mov dword ptr [ebp-0c], esp
'00492d42    c745f8b8134000          mov dword ptr [ebp-08], 004013b8
'00492d49    8b4508                  mov eax, dword ptr [ebp+08]
'00492d4c    8bc8                    mov ecx, eax
'00492d4e    83e101                  and ecx, 01
'00492d51    894dfc                  mov dword ptr [ebp-04], ecx
'00492d54    83e0fe                  and eax, fffffffe
'00492d57    8b10                    mov edx, dword ptr [eax]
'00492d59    50                      push eax
'00492d5a    894508                  mov dword ptr [ebp+08], eax
'00492d5d    ff5204                  call dword ptr [edx+04]
'00492d60    a124b17200              mov ax, word ptr [0072b124]
'00492d65    85c0                    test ax, ax
'00492d67    7510                    jne 492d79
'00492d69    6824b17200              push 0072b124
'00492d6e    683cfc4000              push 0040fc3c

' *** Reference to "__vbaNew2"
'00492d73    ff152c124000            call dword ptr [0040122c]
Dim var_37 As New frmInsertionDons
'00492d79    8b3524b17200            mov esi, dword ptr [0072b124]
'00492d7f    8b3e                    mov edi, dword ptr [esi]
'00492d81    83ec10                  sub esp, 10
'00492d84    8bdc                    mov ebx, esp
'00492d86    b90a000000              mov ecx, 0000000a
'00492d8b    890b                    mov dword ptr [ebx], ecx
'00492d8d    894ddc                  mov dword ptr [ebp-24], ecx
'00492d90    8b4dd0                  mov ecx, dword ptr [ebp-30]
'00492d93    894b04                  mov dword ptr [ebx+04], ecx
'00492d96    b804000280              mov eax, 80020004
'00492d9b    894308                  mov dword ptr [ebx+08], eax
'00492d9e    8bd0                    mov edx, eax
'00492da0    8b45d8                  mov eax, dword ptr [ebp-28]
'00492da3    89430c                  mov dword ptr [ebx+0c], eax
'00492da6    8b45dc                  mov eax, dword ptr [ebp-24]
'00492da9    83ec10                  sub esp, 10
'00492dac    8bcc                    mov ecx, esp
'00492dae    8901                    mov dword ptr [ecx], eax
'00492db0    8b45e0                  mov eax, dword ptr [ebp-20]
'00492db3    894104                  mov dword ptr [ecx+04], eax
'00492db6    895108                  mov dword ptr [ecx+08], edx
'00492db9    8b55e8                  mov edx, dword ptr [ebp-18]
'00492dbc    56                      push esi
'00492dbd    89510c                  mov dword ptr [ecx+0c], edx
'00492dc0    ff97b0020000            call dword ptr [edi+000002b0]
'00492dc6    dbe2                    fnclex
'00492dc8    85c0                    test ax, ax
'00492dca    7d12                    jge 492dde
'00492dcc    68b0020000              push 000002b0
'00492dd1    68b0194300              push 004319b0
'00492dd6    56                      push esi
'00492dd7    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00492dd8    ff1580104000            call dword ptr [00401080]
'00492dde    c745fc00000000          mov dword ptr [ebp-04], 00000000
'00492de5    8b4508                  mov eax, dword ptr [ebp+08]
'00492de8    8b08                    mov ecx, dword ptr [eax]
'00492dea    50                      push eax
'00492deb    ff5108                  call dword ptr [ecx+08]
'00492dee    8b45fc                  mov eax, dword ptr [ebp-04]
'00492df1    8b4dec                  mov ecx, dword ptr [ebp-14]
'00492df4    5f                      pop edi
'00492df5    5e                      pop esi
    'Reference to '__except_list'
'00492df6    64890d00000000          mov dword ptr fs:[00000000], ecx
'00492dfd    5b                      pop ebx
'00492dfe    8be5                    mov esp, ebp
'00492e00    5d                      pop ebp
'00492e01    c20400                  ret 0004


End Sub


'Event for IDM_ASSISTANT_PERSO
Private Sub IDM_ASSISTANT_PERSO_Click()
'00492040    55                      push ebp
'00492041    8bec                    mov ebp, esp
'00492043    83ec0c                  sub esp, 0c

' *** Reference to "__vbaExceptHandler"
'00492046    6866724000              push 00407266
'0049204b    64a100000000            mov ax, word ptr fs:[00000000]
'00492051    50                      push eax
    'Reference to '__except_list'
'00492052    64892500000000          mov dword ptr fs:[00000000], esp
'00492059    83ec30                  sub esp, 30
'0049205c    53                      push ebx
'0049205d    56                      push esi
'0049205e    57                      push edi
'0049205f    8965f4                  mov dword ptr [ebp-0c], esp
'00492062    c745f858134000          mov dword ptr [ebp-08], 00401358
'00492069    8b4508                  mov eax, dword ptr [ebp+08]
'0049206c    8bc8                    mov ecx, eax
'0049206e    83e101                  and ecx, 01
'00492071    894dfc                  mov dword ptr [ebp-04], ecx
'00492074    83e0fe                  and eax, fffffffe
'00492077    8b10                    mov edx, dword ptr [eax]
'00492079    50                      push eax
'0049207a    894508                  mov dword ptr [ebp+08], eax
'0049207d    ff5204                  call dword ptr [edx+04]
'00492080    a184b07200              mov ax, word ptr [0072b084]
'00492085    85c0                    test ax, ax
'00492087    7510                    jne 492099
'00492089    6884b07200              push 0072b084
'0049208e    68f8894100              push 004189f8

' *** Reference to "__vbaNew2"
'00492093    ff152c124000            call dword ptr [0040122c]
Dim var_28 As New frmGestPerso
'00492099    8b3584b07200            mov esi, dword ptr [0072b084]
'0049209f    8b3e                    mov edi, dword ptr [esi]
'004920a1    83ec10                  sub esp, 10
'004920a4    8bdc                    mov ebx, esp
'004920a6    b90a000000              mov ecx, 0000000a
'004920ab    890b                    mov dword ptr [ebx], ecx
'004920ad    8b4dd0                  mov ecx, dword ptr [ebp-30]
'004920b0    894b04                  mov dword ptr [ebx+04], ecx
'004920b3    b804000280              mov eax, 80020004
'004920b8    894308                  mov dword ptr [ebx+08], eax
'004920bb    8bd0                    mov edx, eax
'004920bd    8b45d8                  mov eax, dword ptr [ebp-28]
'004920c0    89430c                  mov dword ptr [ebx+0c], eax
'004920c3    83ec10                  sub esp, 10
'004920c6    8bcc                    mov ecx, esp
'004920c8    b80a000000              mov eax, 0000000a
'004920cd    8901                    mov dword ptr [ecx], eax
'004920cf    8b45e0                  mov eax, dword ptr [ebp-20]
'004920d2    894104                  mov dword ptr [ecx+04], eax
'004920d5    895108                  mov dword ptr [ecx+08], edx
'004920d8    8b55e8                  mov edx, dword ptr [ebp-18]
'004920db    56                      push esi
'004920dc    89510c                  mov dword ptr [ecx+0c], edx
'004920df    ff97b0020000            call dword ptr [edi+000002b0]
'004920e5    dbe2                    fnclex
'004920e7    85c0                    test ax, ax
'004920e9    7d12                    jge 4920fd
'004920eb    68b0020000              push 000002b0
'004920f0    6850014300              push 00430150
'004920f5    56                      push esi
'004920f6    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'004920f7    ff1580104000            call dword ptr [00401080]
'004920fd    a198b07200              mov ax, word ptr [0072b098]
'00492102    85c0                    test ax, ax
'00492104    7510                    jne 492116
'00492106    6898b07200              push 0072b098
'0049210b    68d8354100              push 004135d8

' *** Reference to "__vbaNew2"
'00492110    ff152c124000            call dword ptr [0040122c]
Dim var_29 As New frmAssistant
'00492116    8b3598b07200            mov esi, dword ptr [0072b098]
'0049211c    8b3e                    mov edi, dword ptr [esi]
'0049211e    83ec10                  sub esp, 10
'00492121    8bdc                    mov ebx, esp
'00492123    b90a000000              mov ecx, 0000000a
'00492128    890b                    mov dword ptr [ebx], ecx
'0049212a    894ddc                  mov dword ptr [ebp-24], ecx
'0049212d    8b4dd0                  mov ecx, dword ptr [ebp-30]
'00492130    894b04                  mov dword ptr [ebx+04], ecx
'00492133    b804000280              mov eax, 80020004
'00492138    894308                  mov dword ptr [ebx+08], eax
'0049213b    8bd0                    mov edx, eax
'0049213d    8b45d8                  mov eax, dword ptr [ebp-28]
'00492140    89430c                  mov dword ptr [ebx+0c], eax
'00492143    8b45dc                  mov eax, dword ptr [ebp-24]
'00492146    83ec10                  sub esp, 10
'00492149    8bcc                    mov ecx, esp
'0049214b    8901                    mov dword ptr [ecx], eax
'0049214d    8b45e0                  mov eax, dword ptr [ebp-20]
'00492150    894104                  mov dword ptr [ecx+04], eax
'00492153    895108                  mov dword ptr [ecx+08], edx
'00492156    8b55e8                  mov edx, dword ptr [ebp-18]
'00492159    56                      push esi
'0049215a    89510c                  mov dword ptr [ecx+0c], edx
'0049215d    ff97b0020000            call dword ptr [edi+000002b0]
'00492163    dbe2                    fnclex
'00492165    85c0                    test ax, ax
'00492167    7d12                    jge 49217b
'00492169    68b0020000              push 000002b0
'0049216e    682c0e4300              push 00430e2c
'00492173    56                      push esi
'00492174    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00492175    ff1580104000            call dword ptr [00401080]
'0049217b    c745fc00000000          mov dword ptr [ebp-04], 00000000
'00492182    8b4508                  mov eax, dword ptr [ebp+08]
'00492185    8b08                    mov ecx, dword ptr [eax]
'00492187    50                      push eax
'00492188    ff5108                  call dword ptr [ecx+08]
'0049218b    8b45fc                  mov eax, dword ptr [ebp-04]
'0049218e    8b4dec                  mov ecx, dword ptr [ebp-14]
'00492191    5f                      pop edi
'00492192    5e                      pop esi
    'Reference to '__except_list'
'00492193    64890d00000000          mov dword ptr fs:[00000000], ecx
'0049219a    5b                      pop ebx
'0049219b    8be5                    mov esp, ebp
'0049219d    5d                      pop ebp
'0049219e    c20400                  ret 0004


End Sub


'Event for IDM_QUITTER
Private Sub IDM_QUITTER_Click()
'004934d0    55                      push ebp
'004934d1    8bec                    mov ebp, esp
'004934d3    83ec14                  sub esp, 14

' *** Reference to "__vbaExceptHandler"
'004934d6    6866724000              push 00407266
'004934db    64a100000000            mov ax, word ptr fs:[00000000]
'004934e1    50                      push eax
    'Reference to '__except_list'
'004934e2    64892500000000          mov dword ptr fs:[00000000], esp
'004934e9    81ec30010000            sub esp, 00000130
'004934ef    53                      push ebx
'004934f0    56                      push esi
'004934f1    57                      push edi
'004934f2    8965ec                  mov dword ptr [ebp-14], esp
'004934f5    c745f028144000          mov dword ptr [ebp-10], 00401428
'004934fc    8b5d08                  mov ebx, dword ptr [ebp+08]
'004934ff    8bc3                    mov eax, ebx
'00493501    83e001                  and eax, 01
'00493504    8945f4                  mov dword ptr [ebp-0c], eax
'00493507    83e3fe                  and ebx, fffffffe
'0049350a    895d08                  mov dword ptr [ebp+08], ebx
'0049350d    33f6                    xor esi, esi
'0049350f    8975f8                  mov dword ptr [ebp-08], esi
'00493512    8b0b                    mov ecx, dword ptr [ebx]
'00493514    53                      push ebx
'00493515    ff5104                  call dword ptr [ecx+04]
'00493518    8975e0                  mov dword ptr [ebp-20], esi
'0049351b    8975dc                  mov dword ptr [ebp-24], esi
'0049351e    8975d8                  mov dword ptr [ebp-28], esi
'00493521    8975d4                  mov dword ptr [ebp-2c], esi
'00493524    8975d0                  mov dword ptr [ebp-30], esi
'00493527    8975cc                  mov dword ptr [ebp-34], esi
'0049352a    8975c8                  mov dword ptr [ebp-38], esi
'0049352d    8975c4                  mov dword ptr [ebp-3c], esi
'00493530    8975c0                  mov dword ptr [ebp-40], esi
'00493533    8975b0                  mov dword ptr [ebp-50], esi
'00493536    8975a0                  mov dword ptr [ebp-60], esi
'00493539    897590                  mov dword ptr [ebp-70], esi
'0049353c    897580                  mov dword ptr [ebp-80], esi
'0049353f    89b570ffffff            mov dword ptr [ebp+ffffff70], esi
'00493545    89b560ffffff            mov dword ptr [ebp+ffffff60], esi
'0049354b    89b550ffffff            mov dword ptr [ebp+ffffff50], esi
'00493551    89b540ffffff            mov dword ptr [ebp+ffffff40], esi
'00493557    89b530ffffff            mov dword ptr [ebp+ffffff30], esi
'0049355d    89b520ffffff            mov dword ptr [ebp+ffffff20], esi
'00493563    89b510ffffff            mov dword ptr [ebp+ffffff10], esi
'00493569    89b500ffffff            mov dword ptr [ebp+ffffff00], esi
'0049356f    89b5f0feffff            mov dword ptr [ebp+fffffef0], esi
'00493575    89b5e0feffff            mov dword ptr [ebp+fffffee0], esi
'0049357b    89b5dcfeffff            mov dword ptr [ebp+fffffedc], esi
'00493581    66393528b07200          cmp word ptr [0072b028], si
'00493588    7508                    jne 493592
'0049358a    6a01                    push 01

' *** Reference to "__vbaOnError"
'0049358c    ff15b0104000            call dword ptr [004010b0]
On Error Goto handler_0
'00493592    393524be7200            cmp dword ptr [0072be24], esi
'00493598    7510                    jne 4935aa
'0049359a    6824be7200              push 0072be24
'0049359f    6870004300              push 00430070

' *** Reference to "__vbaNew2"
'004935a4    ff152c124000            call dword ptr [0040122c]
Dim var_2 As New Global
'004935aa    8b3d24be7200            mov edi, dword ptr [0072be24]
'004935b0    8b17                    mov edx, dword ptr [edi]
'004935b2    53                      push ebx
'004935b3    8d45c4                  lea eax, dword ptr [ebp-3c]
'004935b6    50                      push eax
'004935b7    8995b4feffff            mov dword ptr [ebp+fffffeb4], edx

' *** Reference to "__vbaObjSetAddref"
'004935bd    ff15c8104000            call dword ptr [004010c8]
Set var_9 = Me
'004935c3    50                      push eax
'004935c4    57                      push edi
'004935c5    8b8db4feffff            mov ecx, dword ptr [ebp+fffffeb4]
'004935cb    ff5110                  call dword ptr [ecx+10]
Call var_2.Unload(var_9)
'004935ce    dbe2                    fnclex
'004935d0    3bc6                    cmp eax, esi
'004935d2    7d0f                    jge 4935e3
'004935d4    6a10                    push 10
'004935d6    6860004300              push 00430060
'004935db    57                      push edi
'004935dc    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'004935dd    ff1580104000            call dword ptr [00401080]
'004935e3    8d4dc4                  lea ecx, dword ptr [ebp-3c]

' *** Reference to "__vbaFreeObj"
'004935e6    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_9)

' *** Reference to "__vbaExitProc"
'004935ec    ff15a0104000            call dword ptr [004010a0]
'004935f2    68a5394900              push 004939a5
'004935f7    e9a8030000              jmp 4939a4

' *** Reference to "rtcErrObj"
'004935fc    8b1d6c124000            mov ebx, dword ptr [0040126c]
'00493602    ffd3                    call ebx
'00493604    50                      push eax
'00493605    8d55c4                  lea edx, dword ptr [ebp-3c]
'00493608    52                      push edx

' *** Reference to "__vbaObjSet"
'00493609    8b3db4104000            mov edi, dword ptr [004010b4]
'0049360f    ffd7                    call edi
Set var_9 = Err
'00493611    8bf0                    mov esi, eax
'00493613    8b06                    mov eax, dword ptr [esi]
'00493615    8d4de0                  lea ecx, dword ptr [ebp-20]
'00493618    51                      push ecx
'00493619    56                      push esi
'0049361a    ff502c                  call dword ptr [eax+2c]
var_3 = var_9.Description
'0049361d    dbe2                    fnclex
'0049361f    85c0                    test ax, ax
'00493621    7d0f                    jge 493632
'00493623    6a2c                    push 2c
'00493625    685c084300              push 0043085c
'0049362a    56                      push esi
'0049362b    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0049362c    ff1580104000            call dword ptr [00401080]
'00493632    ffd3                    call ebx
'00493634    50                      push eax
'00493635    8d55c0                  lea edx, dword ptr [ebp-40]
'00493638    52                      push edx
'00493639    ffd7                    call edi
Set var_5 = Err
'0049363b    8bf0                    mov esi, eax
'0049363d    8b06                    mov eax, dword ptr [esi]
'0049363f    8d8ddcfeffff            lea ecx, dword ptr [ebp+fffffedc]
'00493645    51                      push ecx
'00493646    56                      push esi
'00493647    ff501c                  call dword ptr [eax+1c]
var_10 = var_5.Number
'0049364a    dbe2                    fnclex
'0049364c    85c0                    test ax, ax
'0049364e    7d0f                    jge 49365f
'00493650    6a1c                    push 1c
'00493652    685c084300              push 0043085c
'00493657    56                      push esi
'00493658    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00493659    ff1580104000            call dword ptr [00401080]
'0049365f    b904000280              mov ecx, 80020004
'00493664    894d98                  mov dword ptr [ebp-68], ecx
'00493667    b80a000000              mov eax, 0000000a
'0049366c    894590                  mov dword ptr [ebp-70], eax
'0049366f    894da8                  mov dword ptr [ebp-58], ecx
'00493672    8945a0                  mov dword ptr [ebp-60], eax
'00493675    c78528ffffff24b07200    mov dword ptr [ebp+ffffff28], 0072b024
'0049367f    c78520ffffff08400000    mov dword ptr [ebp+ffffff20], 00004008
'00493689    6814084300              push 00430814
'0049368e    8b55e0                  mov edx, dword ptr [ebp-20]
'00493691    52                      push edx

' *** Reference to "__vbaStrCat"
'00493692    8b3570104000            mov esi, dword ptr [00401070]
'00493698    ffd6                    call esi
var_11 = ("L'erreur suivante s'est produite : ") & (var_3)
'0049369a    8bd0                    mov edx, eax
'0049369c    8d4ddc                  lea ecx, dword ptr [ebp-24]

' *** Reference to "__vbaStrMove"
'0049369f    8b3dd0124000            mov edi, dword ptr [004012d0]
'004936a5    ffd7                    call edi
'004936a7    50                      push eax
'004936a8    6870084300              push 00430870
'004936ad    ffd6                    call esi
var_12 = (var_11) & (vbCrLf)
'004936af    8bd0                    mov edx, eax
'004936b1    8d4dd8                  lea ecx, dword ptr [ebp-28]
'004936b4    ffd7                    call edi
'004936b6    50                      push eax
'004936b7    6870084300              push 00430870
'004936bc    ffd6                    call esi
var_13 = (var_12) & (vbCrLf)
'004936be    8bd0                    mov edx, eax
'004936c0    8d4dd4                  lea ecx, dword ptr [ebp-2c]
'004936c3    ffd7                    call edi
'004936c5    50                      push eax
'004936c6    6880084300              push 00430880
'004936cb    ffd6                    call esi
var_14 = (var_13) & (" numéro d'erreur (")
'004936cd    8bd0                    mov edx, eax
'004936cf    8d4dd0                  lea ecx, dword ptr [ebp-30]
'004936d2    ffd7                    call edi
'004936d4    50                      push eax
'004936d5    8b85dcfeffff            mov eax, dword ptr [ebp+fffffedc]
'004936db    50                      push eax

' *** Reference to "__vbaStrI4"
'004936dc    ff1520104000            call dword ptr [00401020]
'004936e2    8bd0                    mov edx, eax
'004936e4    8d4dcc                  lea ecx, dword ptr [ebp-34]
'004936e7    ffd7                    call edi
'004936e9    50                      push eax
'004936ea    ffd6                    call esi
var_15 = (var_14) & (CStr(var_10))
'004936ec    8bd0                    mov edx, eax
'004936ee    8d4dc8                  lea ecx, dword ptr [ebp-38]
'004936f1    ffd7                    call edi
'004936f3    50                      push eax
'004936f4    68ac084300              push 004308ac
'004936f9    ffd6                    call esi
var_16 = (var_15) & (" )")
'004936fb    8945b8                  mov dword ptr [ebp-48], eax
'004936fe    bf08000000              mov edi, 00000008
'00493703    897db0                  mov dword ptr [ebp-50], edi
'00493706    8d4d90                  lea ecx, dword ptr [ebp-70]
'00493709    51                      push ecx
'0049370a    8d55a0                  lea edx, dword ptr [ebp-60]
'0049370d    52                      push edx
'0049370e    8d8520ffffff            lea eax, dword ptr [ebp+ffffff20]
'00493714    50                      push eax
'00493715    6a10                    push 10
'00493717    8d4db0                  lea ecx, dword ptr [ebp-50]
'0049371a    51                      push ecx

' *** Reference to "rtcMsgBox"
'0049371b    ff15b8104000            call dword ptr [004010b8]
var_17 = MsgBox(var_16, 16, vbNullString)
'00493721    8d55c8                  lea edx, dword ptr [ebp-38]
'00493724    52                      push edx
'00493725    8d45cc                  lea eax, dword ptr [ebp-34]
'00493728    50                      push eax
'00493729    8d4dd0                  lea ecx, dword ptr [ebp-30]
'0049372c    51                      push ecx
'0049372d    8d55d4                  lea edx, dword ptr [ebp-2c]
'00493730    52                      push edx
'00493731    8d45d8                  lea eax, dword ptr [ebp-28]
'00493734    50                      push eax
'00493735    8d4ddc                  lea ecx, dword ptr [ebp-24]
'00493738    51                      push ecx
'00493739    8d55e0                  lea edx, dword ptr [ebp-20]
'0049373c    52                      push edx
'0049373d    6a07                    push 07

' *** Reference to "__vbaFreeStrList"
'0049373f    ff155c124000            call dword ptr [0040125c]
'#Cleanup( 0, -4512, -4516, -4520, -4524, -4528, -4532)
'00493745    8d45c0                  lea eax, dword ptr [ebp-40]
'00493748    50                      push eax
'00493749    8d4dc4                  lea ecx, dword ptr [ebp-3c]
'0049374c    51                      push ecx
'0049374d    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'0049374f    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_9, var_5)
'00493755    8d5590                  lea edx, dword ptr [ebp-70]
'00493758    52                      push edx
'00493759    8d45a0                  lea eax, dword ptr [ebp-60]
'0049375c    50                      push eax
'0049375d    8d4db0                  lea ecx, dword ptr [ebp-50]
'00493760    51                      push ecx
'00493761    6a03                    push 03

' *** Reference to "__vbaFreeVarList"
'00493763    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_6, var_7, var_8)
'00493769    83c43c                  add esp, 3c
'0049376c    8d55b0                  lea edx, dword ptr [ebp-50]
'0049376f    52                      push edx

' *** Reference to "rtcGetPresentDate"
'00493770    ff15f4124000            call dword ptr [004012f4]
var_16 = Now()
'00493776    c78528ffffffb8084300    mov dword ptr [ebp+ffffff28], 004308b8
'00493780    89bd20ffffff            mov dword ptr [ebp+ffffff20], edi
'00493786    8d9520ffffff            lea edx, dword ptr [ebp+ffffff20]
'0049378c    8d4da0                  lea ecx, dword ptr [ebp-60]

' *** Reference to "__vbaVarDup"
'0049378f    ff1598124000            call dword ptr [00401298]
var_7 = ("dd/MM/yyyy hh:mm:ss")
'00493795    6a01                    push 01
'00493797    6a01                    push 01
'00493799    8d45a0                  lea eax, dword ptr [ebp-60]
'0049379c    50                      push eax
'0049379d    8d4db0                  lea ecx, dword ptr [ebp-50]
'004937a0    51                      push ecx
'004937a1    8d5590                  lea edx, dword ptr [ebp-70]
'004937a4    52                      push edx

' *** Reference to "rtcVarFromFormatVar"
'004937a5    ff1574104000            call dword ptr [00401074]
'004937ab    ffd3                    call ebx
'004937ad    50                      push eax
'004937ae    8d45c4                  lea eax, dword ptr [ebp-3c]
'004937b1    50                      push eax

' *** Reference to "__vbaObjSet"
'004937b2    ff15b4104000            call dword ptr [004010b4]
Set var_9 = Err
'004937b8    8bf0                    mov esi, eax
'004937ba    8b0e                    mov ecx, dword ptr [esi]
'004937bc    8d95dcfeffff            lea edx, dword ptr [ebp+fffffedc]
'004937c2    52                      push edx
'004937c3    56                      push esi
'004937c4    ff511c                  call dword ptr [ecx+1c]
var_10 = var_9.Number
'004937c7    dbe2                    fnclex
'004937c9    85c0                    test ax, ax
'004937cb    7d0f                    jge 4937dc
'004937cd    6a1c                    push 1c
'004937cf    685c084300              push 0043085c
'004937d4    56                      push esi
'004937d5    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'004937d6    ff1580104000            call dword ptr [00401080]
'004937dc    ffd3                    call ebx
'004937de    50                      push eax
'004937df    8d45c0                  lea eax, dword ptr [ebp-40]
'004937e2    50                      push eax

' *** Reference to "__vbaObjSet"
'004937e3    ff15b4104000            call dword ptr [004010b4]
Set var_5 = Err
'004937e9    8bf0                    mov esi, eax
'004937eb    8b0e                    mov ecx, dword ptr [esi]
'004937ed    8d55e0                  lea edx, dword ptr [ebp-20]
'004937f0    52                      push edx
'004937f1    56                      push esi
'004937f2    ff512c                  call dword ptr [ecx+2c]
var_3 = var_5.Description
'004937f5    dbe2                    fnclex
'004937f7    85c0                    test ax, ax
'004937f9    7d0f                    jge 49380a
'004937fb    6a2c                    push 2c
'004937fd    685c084300              push 0043085c
'00493802    56                      push esi
'00493803    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00493804    ff1580104000            call dword ptr [00401080]
'0049380a    c78518ffffffe4084300    mov dword ptr [ebp+ffffff18], 004308e4
'00493814    89bd10ffffff            mov dword ptr [ebp+ffffff10], edi
'0049381a    8b85dcfeffff            mov eax, dword ptr [ebp+fffffedc]
'00493820    898508ffffff            mov dword ptr [ebp+ffffff08], eax
'00493826    c78500ffffff03000000    mov dword ptr [ebp+ffffff00], 00000003
'00493830    c785f8feffff08094300    mov dword ptr [ebp+fffffef8], 00430908
'0049383a    89bdf0feffff            mov dword ptr [ebp+fffffef0], edi
'00493840    8b45e0                  mov eax, dword ptr [ebp-20]
'00493843    c745e000000000          mov dword ptr [ebp-20], 00000000
'0049384a    898558ffffff            mov dword ptr [ebp+ffffff58], eax
'00493850    89bd50ffffff            mov dword ptr [ebp+ffffff50], edi
'00493856    c785e8feffffa81c4300    mov dword ptr [ebp+fffffee8], 00431ca8
'00493860    89bde0feffff            mov dword ptr [ebp+fffffee0], edi
'00493866    8d4d90                  lea ecx, dword ptr [ebp-70]
'00493869    51                      push ecx
'0049386a    8d9510ffffff            lea edx, dword ptr [ebp+ffffff10]
'00493870    52                      push edx
'00493871    8d4580                  lea eax, dword ptr [ebp-80]
'00493874    50                      push eax

' *** Reference to "__vbaVarCat"
'00493875    8b3508124000            mov esi, dword ptr [00401208]
'0049387b    ffd6                    call esi
'0049387d    50                      push eax
'0049387e    8d8d00ffffff            lea ecx, dword ptr [ebp+ffffff00]
'00493884    51                      push ecx
'00493885    8d9570ffffff            lea edx, dword ptr [ebp+ffffff70]
'0049388b    52                      push edx
'0049388c    ffd6                    call esi
'0049388e    50                      push eax
'0049388f    8d85f0feffff            lea eax, dword ptr [ebp+fffffef0]
'00493895    50                      push eax
'00493896    8d8d60ffffff            lea ecx, dword ptr [ebp+ffffff60]
'0049389c    51                      push ecx
'0049389d    ffd6                    call esi
'0049389f    50                      push eax
'004938a0    8d9550ffffff            lea edx, dword ptr [ebp+ffffff50]
'004938a6    52                      push edx
'004938a7    8d8540ffffff            lea eax, dword ptr [ebp+ffffff40]
'004938ad    50                      push eax
'004938ae    ffd6                    call esi
'004938b0    50                      push eax
'004938b1    8d8de0feffff            lea ecx, dword ptr [ebp+fffffee0]
'004938b7    51                      push ecx
'004938b8    8d9530ffffff            lea edx, dword ptr [ebp+ffffff30]
'004938be    52                      push edx
'004938bf    ffd6                    call esi
'004938c1    50                      push eax
'004938c2    33c0                    xor eax, eax
'004938c4    66a12ab07200            mov eax, dword ptr [0072b02a]
'004938ca    50                      push eax
'004938cb    6884094300              push 00430984

' *** Reference to "__vbaPrintFile"
'004938d0    ff15b8114000            call dword ptr [004011b8]
Print #0, 
'004938d6    8d4dc0                  lea ecx, dword ptr [ebp-40]
'004938d9    51                      push ecx
'004938da    8d55c4                  lea edx, dword ptr [ebp-3c]
'004938dd    52                      push edx
'004938de    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'004938e0    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_9, var_5)
'004938e6    8d8530ffffff            lea eax, dword ptr [ebp+ffffff30]
'004938ec    50                      push eax
'004938ed    8d8d40ffffff            lea ecx, dword ptr [ebp+ffffff40]
'004938f3    51                      push ecx
'004938f4    8d9550ffffff            lea edx, dword ptr [ebp+ffffff50]
'004938fa    52                      push edx
'004938fb    8d8560ffffff            lea eax, dword ptr [ebp+ffffff60]
'00493901    50                      push eax
'00493902    8d8d70ffffff            lea ecx, dword ptr [ebp+ffffff70]
'00493908    51                      push ecx
'00493909    8d5580                  lea edx, dword ptr [ebp-80]
'0049390c    52                      push edx
'0049390d    8d4590                  lea eax, dword ptr [ebp-70]
'00493910    50                      push eax
'00493911    8d4da0                  lea ecx, dword ptr [ebp-60]
'00493914    51                      push ecx
'00493915    8d55b0                  lea edx, dword ptr [ebp-50]
'00493918    52                      push edx
'00493919    6a09                    push 09

' *** Reference to "__vbaFreeVarList"
'0049391b    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_6, var_7, var_8, var_18, var_19, var_20, var_21, var_22, var_23)
'00493921    83c440                  add esp, 40
'00493924    6a00                    push 00

' *** Reference to "__vbaResume"
'00493926    ff1568104000            call dword ptr [00401068]
'0049392c    e9bbfcffff              jmp 4935ec
Resume handler_4935EC
'00493931    8d55c8                  lea edx, dword ptr [ebp-38]
'00493934    52                      push edx
'00493935    8d45cc                  lea eax, dword ptr [ebp-34]
'00493938    50                      push eax
'00493939    8d4dd0                  lea ecx, dword ptr [ebp-30]
'0049393c    51                      push ecx
'0049393d    8d55d4                  lea edx, dword ptr [ebp-2c]
'00493940    52                      push edx
'00493941    8d45d8                  lea eax, dword ptr [ebp-28]
'00493944    50                      push eax
'00493945    8d4ddc                  lea ecx, dword ptr [ebp-24]
'00493948    51                      push ecx
'00493949    8d55e0                  lea edx, dword ptr [ebp-20]
'0049394c    52                      push edx
'0049394d    6a07                    push 07

' *** Reference to "__vbaFreeStrList"
'0049394f    ff155c124000            call dword ptr [0040125c]
'#Cleanup( var_3, -4512, -4516, -4520, -4524, -4528, -4532)
'00493955    8d45c0                  lea eax, dword ptr [ebp-40]
'00493958    50                      push eax
'00493959    8d4dc4                  lea ecx, dword ptr [ebp-3c]
'0049395c    51                      push ecx
'0049395d    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'0049395f    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_9, var_5)
'00493965    8d9530ffffff            lea edx, dword ptr [ebp+ffffff30]
'0049396b    52                      push edx
'0049396c    8d8540ffffff            lea eax, dword ptr [ebp+ffffff40]
'00493972    50                      push eax
'00493973    8d8d50ffffff            lea ecx, dword ptr [ebp+ffffff50]
'00493979    51                      push ecx
'0049397a    8d9560ffffff            lea edx, dword ptr [ebp+ffffff60]
'00493980    52                      push edx
'00493981    8d8570ffffff            lea eax, dword ptr [ebp+ffffff70]
'00493987    50                      push eax
'00493988    8d4d80                  lea ecx, dword ptr [ebp-80]
'0049398b    51                      push ecx
'0049398c    8d5590                  lea edx, dword ptr [ebp-70]
'0049398f    52                      push edx
'00493990    8d45a0                  lea eax, dword ptr [ebp-60]
'00493993    50                      push eax
'00493994    8d4db0                  lea ecx, dword ptr [ebp-50]
'00493997    51                      push ecx
'00493998    6a09                    push 09

' *** Reference to "__vbaFreeVarList"
'0049399a    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_6, var_7, var_8, var_18, var_19, var_20, var_21, var_22, var_23)
'004939a0    83c454                  add esp, 54
'004939a3    c3                      ret
'004939a4    c3                      ret
'004939a5    8b4508                  mov eax, dword ptr [ebp+08]
'004939a8    8b10                    mov edx, dword ptr [eax]
'004939aa    50                      push eax
'004939ab    ff5208                  call dword ptr [edx+08]
'004939ae    8b45f4                  mov eax, dword ptr [ebp-0c]
'004939b1    8b4de4                  mov ecx, dword ptr [ebp-1c]
    'Reference to '__except_list'
'004939b4    64890d00000000          mov dword ptr fs:[00000000], ecx
'004939bb    5f                      pop edi
'004939bc    5e                      pop esi
'004939bd    5b                      pop ebx
'004939be    8be5                    mov esp, ebp
'004939c0    5d                      pop ebp
'004939c1    c20400                  ret 0004


End Sub


'Event for IDM_ACCEDER
Private Sub IDM_ACCEDER_Click()
'00491970    55                      push ebp
'00491971    8bec                    mov ebp, esp
'00491973    83ec14                  sub esp, 14

' *** Reference to "__vbaExceptHandler"
'00491976    6866724000              push 00407266
'0049197b    64a100000000            mov ax, word ptr fs:[00000000]
'00491981    50                      push eax
    'Reference to '__except_list'
'00491982    64892500000000          mov dword ptr fs:[00000000], esp
'00491989    81ec30010000            sub esp, 00000130
'0049198f    53                      push ebx
'00491990    56                      push esi
'00491991    57                      push edi
'00491992    8965ec                  mov dword ptr [ebp-14], esp
'00491995    c745f028134000          mov dword ptr [ebp-10], 00401328
'0049199c    8b7508                  mov esi, dword ptr [ebp+08]
'0049199f    8bc6                    mov eax, esi
'004919a1    83e001                  and eax, 01
'004919a4    8945f4                  mov dword ptr [ebp-0c], eax
'004919a7    83e6fe                  and esi, fffffffe
'004919aa    897508                  mov dword ptr [ebp+08], esi
'004919ad    33ff                    xor edi, edi
'004919af    897df8                  mov dword ptr [ebp-08], edi
'004919b2    8b0e                    mov ecx, dword ptr [esi]
'004919b4    56                      push esi
'004919b5    ff5104                  call dword ptr [ecx+04]
'004919b8    897de0                  mov dword ptr [ebp-20], edi
'004919bb    897ddc                  mov dword ptr [ebp-24], edi
'004919be    897dd8                  mov dword ptr [ebp-28], edi
'004919c1    897dd4                  mov dword ptr [ebp-2c], edi
'004919c4    897dd0                  mov dword ptr [ebp-30], edi
'004919c7    897dcc                  mov dword ptr [ebp-34], edi
'004919ca    897dc8                  mov dword ptr [ebp-38], edi
'004919cd    897dc4                  mov dword ptr [ebp-3c], edi
'004919d0    897dc0                  mov dword ptr [ebp-40], edi
'004919d3    897db0                  mov dword ptr [ebp-50], edi
'004919d6    897da0                  mov dword ptr [ebp-60], edi
'004919d9    897d90                  mov dword ptr [ebp-70], edi
'004919dc    897d80                  mov dword ptr [ebp-80], edi
'004919df    89bd70ffffff            mov dword ptr [ebp+ffffff70], edi
'004919e5    89bd60ffffff            mov dword ptr [ebp+ffffff60], edi
'004919eb    89bd50ffffff            mov dword ptr [ebp+ffffff50], edi
'004919f1    89bd40ffffff            mov dword ptr [ebp+ffffff40], edi
'004919f7    89bd30ffffff            mov dword ptr [ebp+ffffff30], edi
'004919fd    89bd20ffffff            mov dword ptr [ebp+ffffff20], edi
'00491a03    89bd10ffffff            mov dword ptr [ebp+ffffff10], edi
'00491a09    89bd00ffffff            mov dword ptr [ebp+ffffff00], edi
'00491a0f    89bdf0feffff            mov dword ptr [ebp+fffffef0], edi
'00491a15    89bde0feffff            mov dword ptr [ebp+fffffee0], edi
'00491a1b    89bddcfeffff            mov dword ptr [ebp+fffffedc], edi
'00491a21    89bdd8feffff            mov dword ptr [ebp+fffffed8], edi
'00491a27    66393d28b07200          cmp word ptr [0072b028], di
'00491a2e    7508                    jne 491a38
'00491a30    6a01                    push 01

' *** Reference to "__vbaOnError"
'00491a32    ff15b0104000            call dword ptr [004010b0]
On Error Goto handler_0
'00491a38    baf8064300              mov edx, 004306f8
'00491a3d    8d4de0                  lea ecx, dword ptr [ebp-20]

' *** Reference to "__vbaStrCopy"
'00491a40    ff1548124000            call dword ptr [00401248]
var_3 = ("frmAcceder")
'00491a46    8b06                    mov eax, dword ptr [esi]
'00491a48    8d8ddcfeffff            lea ecx, dword ptr [ebp+fffffedc]
'00491a4e    51                      push ecx
'00491a4f    8d55e0                  lea edx, dword ptr [ebp-20]
'00491a52    52                      push edx
'00491a53    56                      push esi
'00491a54    ff90f8060000            call dword ptr [eax+000006f8]
'00491a5a    3bc7                    cmp eax, edi
'00491a5c    7d12                    jge 491a70

If (frmMain < 0) Then
'00491a5e    68f8060000              push 000006f8
'00491a63    687cfd4200              push 0042fd7c
'00491a68    56                      push esi
'00491a69    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00491a6a    ff1580104000            call dword ptr [00401080]
    
End If
'00491a70    33c0                    xor eax, eax
var_num1 = Empty
'00491a72    6683bddcfeffffff        cmp word ptr [ebp+fffffedc], ffffffff
'00491a7a    0f94c0                  sete al
'00491a7d    f7d8                    neg eax
'00491a7f    668bf0                  mov si, ax
'00491a82    8d4de0                  lea ecx, dword ptr [ebp-20]

' *** Reference to "__vbaFreeStr"
'00491a85    ff150c134000            call dword ptr [0040130c]
'#Cleanup(var_3)
'00491a8b    663bf7                  cmp si, di
'00491a8e    a15cb07200              mov ax, word ptr [0072b05c]
'00491a93    0f84d4030000            je 491e6d

If (0 = -1) Then
'00491a99    3bc7                    cmp eax, edi
'00491a9b    7510                    jne 491aad
    
    If (    0 = 0) Then
'00491a9d    685cb07200              push 0072b05c
'00491aa2    68347b4000              push 00407b34

' *** Reference to "__vbaNew2"
'00491aa7    ff152c124000            call dword ptr [0040122c]
    Dim var_24 As New frmAcceder
End If
'00491aad    8b355cb07200            mov esi, dword ptr [0072b05c]
'00491ab3    b804000280              mov eax, 80020004
'00491ab8    898518ffffff            mov dword ptr [ebp+ffffff18], eax
'00491abe    b90a000000              mov ecx, 0000000a
'00491ac3    898d10ffffff            mov dword ptr [ebp+ffffff10], ecx
'00491ac9    898528ffffff            mov dword ptr [ebp+ffffff28], eax
'00491acf    898d20ffffff            mov dword ptr [ebp+ffffff20], ecx
'00491ad5    8b16                    mov edx, dword ptr [esi]
'00491ad7    83ec10                  sub esp, 10
'00491ada    8bdc                    mov ebx, esp
'00491adc    890b                    mov dword ptr [ebx], ecx
'00491ade    8b8d14ffffff            mov ecx, dword ptr [ebp+ffffff14]
'00491ae4    894b04                  mov dword ptr [ebx+04], ecx
'00491ae7    894308                  mov dword ptr [ebx+08], eax
'00491aea    8b851cffffff            mov eax, dword ptr [ebp+ffffff1c]
'00491af0    89430c                  mov dword ptr [ebx+0c], eax
'00491af3    83ec10                  sub esp, 10
'00491af6    8bcc                    mov ecx, esp
'00491af8    8b8520ffffff            mov eax, dword ptr [ebp+ffffff20]
'00491afe    8901                    mov dword ptr [ecx], eax
'00491b00    8b8524ffffff            mov eax, dword ptr [ebp+ffffff24]
'00491b06    894104                  mov dword ptr [ecx+04], eax
'00491b09    8b8528ffffff            mov eax, dword ptr [ebp+ffffff28]
'00491b0f    894108                  mov dword ptr [ecx+08], eax
'00491b12    8b852cffffff            mov eax, dword ptr [ebp+ffffff2c]
'00491b18    89410c                  mov dword ptr [ecx+0c], eax
'00491b1b    56                      push esi
'00491b1c    ff92b0020000            call dword ptr [edx+000002b0]
'00491b22    dbe2                    fnclex
'00491b24    3bc7                    cmp eax, edi
'00491b26    0f8d7c030000            jge 491ea8

If (0 < 0) Then
'00491b2c    68b0020000              push 000002b0
'00491b31    6810074300              push 00430710
'00491b36    e965030000              jmp 491ea0

' *** Reference to "rtcErrObj"
'00491b3b    8b1d6c124000            mov ebx, dword ptr [0040126c]
'00491b41    ffd3                    call ebx
'00491b43    50                      push eax
'00491b44    8d55c4                  lea edx, dword ptr [ebp-3c]
'00491b47    52                      push edx

' *** Reference to "__vbaObjSet"
'00491b48    8b3db4104000            mov edi, dword ptr [004010b4]
'00491b4e    ffd7                    call edi
    Set var_9 = Err
'00491b50    8bf0                    mov esi, eax
'00491b52    8b06                    mov eax, dword ptr [esi]
'00491b54    8d4de0                  lea ecx, dword ptr [ebp-20]
'00491b57    51                      push ecx
'00491b58    56                      push esi
'00491b59    ff502c                  call dword ptr [eax+2c]
    var_3 = var_9.Description
'00491b5c    dbe2                    fnclex
'00491b5e    85c0                    test ax, ax
'00491b60    7d0f                    jge 491b71
'00491b62    6a2c                    push 2c
'00491b64    685c084300              push 0043085c
'00491b69    56                      push esi
'00491b6a    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00491b6b    ff1580104000            call dword ptr [00401080]
'00491b71    ffd3                    call ebx
'00491b73    50                      push eax
'00491b74    8d55c0                  lea edx, dword ptr [ebp-40]
'00491b77    52                      push edx
'00491b78    ffd7                    call edi
    Set var_5 = Err
'00491b7a    8bf0                    mov esi, eax
'00491b7c    8b06                    mov eax, dword ptr [esi]
'00491b7e    8d8dd8feffff            lea ecx, dword ptr [ebp+fffffed8]
'00491b84    51                      push ecx
'00491b85    56                      push esi
'00491b86    ff501c                  call dword ptr [eax+1c]
    var_25 = var_5.Number
'00491b89    dbe2                    fnclex
'00491b8b    85c0                    test ax, ax
'00491b8d    7d0f                    jge 491b9e
'00491b8f    6a1c                    push 1c
'00491b91    685c084300              push 0043085c
'00491b96    56                      push esi
'00491b97    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00491b98    ff1580104000            call dword ptr [00401080]
'00491b9e    b804000280              mov eax, 80020004
'00491ba3    894598                  mov dword ptr [ebp-68], eax
'00491ba6    b90a000000              mov ecx, 0000000a
'00491bab    894d90                  mov dword ptr [ebp-70], ecx
'00491bae    8945a8                  mov dword ptr [ebp-58], eax
'00491bb1    894da0                  mov dword ptr [ebp-60], ecx
'00491bb4    c78528ffffff24b07200    mov dword ptr [ebp+ffffff28], 0072b024
'00491bbe    c78520ffffff08400000    mov dword ptr [ebp+ffffff20], 00004008
'00491bc8    6814084300              push 00430814
'00491bcd    8b55e0                  mov edx, dword ptr [ebp-20]
'00491bd0    52                      push edx

' *** Reference to "__vbaStrCat"
'00491bd1    8b3570104000            mov esi, dword ptr [00401070]
'00491bd7    ffd6                    call esi
    var_26 = ("L'erreur suivante s'est produite : ") & (var_3)
'00491bd9    8bd0                    mov edx, eax
'00491bdb    8d4ddc                  lea ecx, dword ptr [ebp-24]

' *** Reference to "__vbaStrMove"
'00491bde    8b3dd0124000            mov edi, dword ptr [004012d0]
'00491be4    ffd7                    call edi
'00491be6    50                      push eax
'00491be7    6870084300              push 00430870
'00491bec    ffd6                    call esi
    var_12 = (var_26) & (vbCrLf)
'00491bee    8bd0                    mov edx, eax
'00491bf0    8d4dd8                  lea ecx, dword ptr [ebp-28]
'00491bf3    ffd7                    call edi
'00491bf5    50                      push eax
'00491bf6    6870084300              push 00430870
'00491bfb    ffd6                    call esi
    var_13 = (var_12) & (vbCrLf)
'00491bfd    8bd0                    mov edx, eax
'00491bff    8d4dd4                  lea ecx, dword ptr [ebp-2c]
'00491c02    ffd7                    call edi
'00491c04    50                      push eax
'00491c05    6880084300              push 00430880
'00491c0a    ffd6                    call esi
    var_14 = (var_13) & (" numéro d'erreur (")
'00491c0c    8bd0                    mov edx, eax
'00491c0e    8d4dd0                  lea ecx, dword ptr [ebp-30]
'00491c11    ffd7                    call edi
'00491c13    50                      push eax
'00491c14    8b85d8feffff            mov eax, dword ptr [ebp+fffffed8]
'00491c1a    50                      push eax

' *** Reference to "__vbaStrI4"
'00491c1b    ff1520104000            call dword ptr [00401020]
'00491c21    8bd0                    mov edx, eax
'00491c23    8d4dcc                  lea ecx, dword ptr [ebp-34]
'00491c26    ffd7                    call edi
'00491c28    50                      push eax
'00491c29    ffd6                    call esi
    var_15 = (var_14) & (CStr(var_25))
'00491c2b    8bd0                    mov edx, eax
'00491c2d    8d4dc8                  lea ecx, dword ptr [ebp-38]
'00491c30    ffd7                    call edi
'00491c32    50                      push eax
'00491c33    68ac084300              push 004308ac
'00491c38    ffd6                    call esi
    var_16 = (var_15) & (" )")
'00491c3a    8945b8                  mov dword ptr [ebp-48], eax
'00491c3d    bf08000000              mov edi, 00000008
'00491c42    897db0                  mov dword ptr [ebp-50], edi
'00491c45    8d4d90                  lea ecx, dword ptr [ebp-70]
'00491c48    51                      push ecx
'00491c49    8d55a0                  lea edx, dword ptr [ebp-60]
'00491c4c    52                      push edx
'00491c4d    8d8520ffffff            lea eax, dword ptr [ebp+ffffff20]
'00491c53    50                      push eax
'00491c54    6a10                    push 10
'00491c56    8d4db0                  lea ecx, dword ptr [ebp-50]
'00491c59    51                      push ecx

' *** Reference to "rtcMsgBox"
'00491c5a    ff15b8104000            call dword ptr [004010b8]
    var_17 = MsgBox(var_16, 16, vbNullString)
'00491c60    8d55c8                  lea edx, dword ptr [ebp-38]
'00491c63    52                      push edx
'00491c64    8d45cc                  lea eax, dword ptr [ebp-34]
'00491c67    50                      push eax
'00491c68    8d4dd0                  lea ecx, dword ptr [ebp-30]
'00491c6b    51                      push ecx
'00491c6c    8d55d4                  lea edx, dword ptr [ebp-2c]
'00491c6f    52                      push edx
'00491c70    8d45d8                  lea eax, dword ptr [ebp-28]
'00491c73    50                      push eax
'00491c74    8d4ddc                  lea ecx, dword ptr [ebp-24]
'00491c77    51                      push ecx
'00491c78    8d55e0                  lea edx, dword ptr [ebp-20]
'00491c7b    52                      push edx
'00491c7c    6a07                    push 07

' *** Reference to "__vbaFreeStrList"
'00491c7e    ff155c124000            call dword ptr [0040125c]
    '#Cleanup( , -4512, -4516, -4520, -4524, -4528, -4532)
'00491c84    8d45c0                  lea eax, dword ptr [ebp-40]
'00491c87    50                      push eax
'00491c88    8d4dc4                  lea ecx, dword ptr [ebp-3c]
'00491c8b    51                      push ecx
'00491c8c    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'00491c8e    ff154c104000            call dword ptr [0040104c]
    '#Cleanup( var_9, var_5)
'00491c94    8d5590                  lea edx, dword ptr [ebp-70]
'00491c97    52                      push edx
'00491c98    8d45a0                  lea eax, dword ptr [ebp-60]
'00491c9b    50                      push eax
'00491c9c    8d4db0                  lea ecx, dword ptr [ebp-50]
'00491c9f    51                      push ecx
'00491ca0    6a03                    push 03

' *** Reference to "__vbaFreeVarList"
'00491ca2    ff1540104000            call dword ptr [00401040]
    '#Cleanup( var_6, var_7, var_8)
'00491ca8    83c43c                  add esp, 3c
'00491cab    8d55b0                  lea edx, dword ptr [ebp-50]
'00491cae    52                      push edx

' *** Reference to "rtcGetPresentDate"
'00491caf    ff15f4124000            call dword ptr [004012f4]
    var_16 = Now()
'00491cb5    c78528ffffffb8084300    mov dword ptr [ebp+ffffff28], 004308b8
'00491cbf    89bd20ffffff            mov dword ptr [ebp+ffffff20], edi
'00491cc5    8d9520ffffff            lea edx, dword ptr [ebp+ffffff20]
'00491ccb    8d4da0                  lea ecx, dword ptr [ebp-60]

' *** Reference to "__vbaVarDup"
'00491cce    ff1598124000            call dword ptr [00401298]
    var_7 = ("dd/MM/yyyy hh:mm:ss")
'00491cd4    6a01                    push 01
'00491cd6    6a01                    push 01
'00491cd8    8d45a0                  lea eax, dword ptr [ebp-60]
'00491cdb    50                      push eax
'00491cdc    8d4db0                  lea ecx, dword ptr [ebp-50]
'00491cdf    51                      push ecx
'00491ce0    8d5590                  lea edx, dword ptr [ebp-70]
'00491ce3    52                      push edx

' *** Reference to "rtcVarFromFormatVar"
'00491ce4    ff1574104000            call dword ptr [00401074]
'00491cea    ffd3                    call ebx
'00491cec    50                      push eax
'00491ced    8d45c4                  lea eax, dword ptr [ebp-3c]
'00491cf0    50                      push eax

' *** Reference to "__vbaObjSet"
'00491cf1    ff15b4104000            call dword ptr [004010b4]
    Set var_9 = Err
'00491cf7    8bf0                    mov esi, eax
'00491cf9    8b0e                    mov ecx, dword ptr [esi]
'00491cfb    8d95d8feffff            lea edx, dword ptr [ebp+fffffed8]
'00491d01    52                      push edx
'00491d02    56                      push esi
'00491d03    ff511c                  call dword ptr [ecx+1c]
    var_25 = var_9.Number
'00491d06    dbe2                    fnclex
'00491d08    85c0                    test ax, ax
'00491d0a    7d0f                    jge 491d1b
'00491d0c    6a1c                    push 1c
'00491d0e    685c084300              push 0043085c
'00491d13    56                      push esi
'00491d14    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00491d15    ff1580104000            call dword ptr [00401080]
'00491d1b    ffd3                    call ebx
'00491d1d    50                      push eax
'00491d1e    8d45c0                  lea eax, dword ptr [ebp-40]
'00491d21    50                      push eax

' *** Reference to "__vbaObjSet"
'00491d22    ff15b4104000            call dword ptr [004010b4]
    Set var_5 = Err
'00491d28    8bf0                    mov esi, eax
'00491d2a    8b0e                    mov ecx, dword ptr [esi]
'00491d2c    8d55e0                  lea edx, dword ptr [ebp-20]
'00491d2f    52                      push edx
'00491d30    56                      push esi
'00491d31    ff512c                  call dword ptr [ecx+2c]
    var_3 = var_5.Description
'00491d34    dbe2                    fnclex
'00491d36    85c0                    test ax, ax
'00491d38    7d0f                    jge 491d49
'00491d3a    6a2c                    push 2c
'00491d3c    685c084300              push 0043085c
'00491d41    56                      push esi
'00491d42    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00491d43    ff1580104000            call dword ptr [00401080]
'00491d49    c78518ffffffe4084300    mov dword ptr [ebp+ffffff18], 004308e4
'00491d53    89bd10ffffff            mov dword ptr [ebp+ffffff10], edi
'00491d59    8b85d8feffff            mov eax, dword ptr [ebp+fffffed8]
'00491d5f    898508ffffff            mov dword ptr [ebp+ffffff08], eax
'00491d65    c78500ffffff03000000    mov dword ptr [ebp+ffffff00], 00000003
'00491d6f    c785f8feffff08094300    mov dword ptr [ebp+fffffef8], 00430908
'00491d79    89bdf0feffff            mov dword ptr [ebp+fffffef0], edi
'00491d7f    8b45e0                  mov eax, dword ptr [ebp-20]
'00491d82    c745e000000000          mov dword ptr [ebp-20], 00000000
'00491d89    898558ffffff            mov dword ptr [ebp+ffffff58], eax
'00491d8f    89bd50ffffff            mov dword ptr [ebp+ffffff50], edi
'00491d95    c785e8feffff14094300    mov dword ptr [ebp+fffffee8], 00430914
'00491d9f    89bde0feffff            mov dword ptr [ebp+fffffee0], edi
'00491da5    8d4d90                  lea ecx, dword ptr [ebp-70]
'00491da8    51                      push ecx
'00491da9    8d9510ffffff            lea edx, dword ptr [ebp+ffffff10]
'00491daf    52                      push edx
'00491db0    8d4580                  lea eax, dword ptr [ebp-80]
'00491db3    50                      push eax

' *** Reference to "__vbaVarCat"
'00491db4    8b3508124000            mov esi, dword ptr [00401208]
'00491dba    ffd6                    call esi
'00491dbc    50                      push eax
'00491dbd    8d8d00ffffff            lea ecx, dword ptr [ebp+ffffff00]
'00491dc3    51                      push ecx
'00491dc4    8d9570ffffff            lea edx, dword ptr [ebp+ffffff70]
'00491dca    52                      push edx
'00491dcb    ffd6                    call esi
'00491dcd    50                      push eax
'00491dce    8d85f0feffff            lea eax, dword ptr [ebp+fffffef0]
'00491dd4    50                      push eax
'00491dd5    8d8d60ffffff            lea ecx, dword ptr [ebp+ffffff60]
'00491ddb    51                      push ecx
'00491ddc    ffd6                    call esi
'00491dde    50                      push eax
'00491ddf    8d9550ffffff            lea edx, dword ptr [ebp+ffffff50]
'00491de5    52                      push edx
'00491de6    8d8540ffffff            lea eax, dword ptr [ebp+ffffff40]
'00491dec    50                      push eax
'00491ded    ffd6                    call esi
'00491def    50                      push eax
'00491df0    8d8de0feffff            lea ecx, dword ptr [ebp+fffffee0]
'00491df6    51                      push ecx
'00491df7    8d9530ffffff            lea edx, dword ptr [ebp+ffffff30]
'00491dfd    52                      push edx
'00491dfe    ffd6                    call esi
'00491e00    50                      push eax
'00491e01    33c0                    xor eax, eax
'00491e03    66a12ab07200            mov eax, dword ptr [0072b02a]
'00491e09    50                      push eax
'00491e0a    6884094300              push 00430984

' *** Reference to "__vbaPrintFile"
'00491e0f    ff15b8114000            call dword ptr [004011b8]
    Print #0, 
'00491e15    8d4dc0                  lea ecx, dword ptr [ebp-40]
'00491e18    51                      push ecx
'00491e19    8d55c4                  lea edx, dword ptr [ebp-3c]
'00491e1c    52                      push edx
'00491e1d    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'00491e1f    ff154c104000            call dword ptr [0040104c]
    '#Cleanup( var_9, var_5)
'00491e25    8d8530ffffff            lea eax, dword ptr [ebp+ffffff30]
'00491e2b    50                      push eax
'00491e2c    8d8d40ffffff            lea ecx, dword ptr [ebp+ffffff40]
'00491e32    51                      push ecx
'00491e33    8d9550ffffff            lea edx, dword ptr [ebp+ffffff50]
'00491e39    52                      push edx
'00491e3a    8d8560ffffff            lea eax, dword ptr [ebp+ffffff60]
'00491e40    50                      push eax
'00491e41    8d8d70ffffff            lea ecx, dword ptr [ebp+ffffff70]
'00491e47    51                      push ecx
'00491e48    8d5580                  lea edx, dword ptr [ebp-80]
'00491e4b    52                      push edx
'00491e4c    8d4590                  lea eax, dword ptr [ebp-70]
'00491e4f    50                      push eax
'00491e50    8d4da0                  lea ecx, dword ptr [ebp-60]
'00491e53    51                      push ecx
'00491e54    8d55b0                  lea edx, dword ptr [ebp-50]
'00491e57    52                      push edx
'00491e58    6a09                    push 09

' *** Reference to "__vbaFreeVarList"
'00491e5a    ff1540104000            call dword ptr [00401040]
    '#Cleanup( var_6, var_7, var_8, var_18, var_19, var_20, var_21, var_22, var_23)
'00491e60    83c440                  add esp, 40
'00491e63    6a00                    push 00

' *** Reference to "__vbaResume"
'00491e65    ff1568104000            call dword ptr [00401068]
'00491e6b    eb3b                    jmp 491ea8
    Resume handler_491EA8
End If
'00491e6d    3bc7                    cmp eax, edi
'00491e6f    7510                    jne 491e81

If (0 = 0) Then
'00491e71    685cb07200              push 0072b05c
'00491e76    68347b4000              push 00407b34

' *** Reference to "__vbaNew2"
'00491e7b    ff152c124000            call dword ptr [0040122c]
    Set var_24 = New frmAcceder
End If
'00491e81    8b355cb07200            mov esi, dword ptr [0072b05c]
'00491e87    8b0e                    mov ecx, dword ptr [esi]
'00491e89    56                      push esi
'00491e8a    ff91fc060000            call dword ptr [ecx+000006fc]
'00491e90    dbe2                    fnclex
'00491e92    3bc7                    cmp eax, edi
'00491e94    7d12                    jge 491ea8

If (0 < 0) Then
'00491e96    68fc060000              push 000006fc
'00491e9b    6840074300              push 00430740
'00491ea0    56                      push esi
'00491ea1    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00491ea2    ff1580104000            call dword ptr [00401080]
    
End If

' *** Reference to "__vbaExitProc"
'00491ea8    ff15a0104000            call dword ptr [004010a0]
'00491eae    68291f4900              push 00491f29
'00491eb3    eb73                    jmp 491f28
'00491eb5    8d55c8                  lea edx, dword ptr [ebp-38]
'00491eb8    52                      push edx
'00491eb9    8d45cc                  lea eax, dword ptr [ebp-34]
'00491ebc    50                      push eax
'00491ebd    8d4dd0                  lea ecx, dword ptr [ebp-30]
'00491ec0    51                      push ecx
'00491ec1    8d55d4                  lea edx, dword ptr [ebp-2c]
'00491ec4    52                      push edx
'00491ec5    8d45d8                  lea eax, dword ptr [ebp-28]
'00491ec8    50                      push eax
'00491ec9    8d4ddc                  lea ecx, dword ptr [ebp-24]
'00491ecc    51                      push ecx
'00491ecd    8d55e0                  lea edx, dword ptr [ebp-20]
'00491ed0    52                      push edx
'00491ed1    6a07                    push 07

' *** Reference to "__vbaFreeStrList"
'00491ed3    ff155c124000            call dword ptr [0040125c]
'#Cleanup( var_3, -4512, -4516, -4520, -4524, -4528, -4532)
'00491ed9    8d45c0                  lea eax, dword ptr [ebp-40]
'00491edc    50                      push eax
'00491edd    8d4dc4                  lea ecx, dword ptr [ebp-3c]
'00491ee0    51                      push ecx
'00491ee1    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'00491ee3    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_9, var_5)
'00491ee9    8d9530ffffff            lea edx, dword ptr [ebp+ffffff30]
'00491eef    52                      push edx
'00491ef0    8d8540ffffff            lea eax, dword ptr [ebp+ffffff40]
'00491ef6    50                      push eax
'00491ef7    8d8d50ffffff            lea ecx, dword ptr [ebp+ffffff50]
'00491efd    51                      push ecx
'00491efe    8d9560ffffff            lea edx, dword ptr [ebp+ffffff60]
'00491f04    52                      push edx
'00491f05    8d8570ffffff            lea eax, dword ptr [ebp+ffffff70]
'00491f0b    50                      push eax
'00491f0c    8d4d80                  lea ecx, dword ptr [ebp-80]
'00491f0f    51                      push ecx
'00491f10    8d5590                  lea edx, dword ptr [ebp-70]
'00491f13    52                      push edx
'00491f14    8d45a0                  lea eax, dword ptr [ebp-60]
'00491f17    50                      push eax
'00491f18    8d4db0                  lea ecx, dword ptr [ebp-50]
'00491f1b    51                      push ecx
'00491f1c    6a09                    push 09

' *** Reference to "__vbaFreeVarList"
'00491f1e    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_6, var_7, var_8, var_18, var_19, var_20, var_21, var_22, var_23)
'00491f24    83c454                  add esp, 54
'00491f27    c3                      ret
'00491f28    c3                      ret
'00491f29    8b4508                  mov eax, dword ptr [ebp+08]
'00491f2c    8b10                    mov edx, dword ptr [eax]
'00491f2e    50                      push eax
'00491f2f    ff5208                  call dword ptr [edx+08]
'00491f32    8b45f4                  mov eax, dword ptr [ebp-0c]
'00491f35    8b4de4                  mov ecx, dword ptr [ebp-1c]
    'Reference to '__except_list'
'00491f38    64890d00000000          mov dword ptr fs:[00000000], ecx
'00491f3f    5f                      pop edi
'00491f40    5e                      pop esi
'00491f41    5b                      pop ebx
'00491f42    8be5                    mov esp, ebp
'00491f44    5d                      pop ebp
'00491f45    c20400                  ret 0004


End Sub


