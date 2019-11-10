VERSION 5.00

Begin VB.Form frmBackHist
    Caption = "Background et Historique"
    ScaleMode = 1
    AutoRedraw = 0              'False
    FontTransparent = -1              'True
    BorderStyle = 1
    LinkTopic = "Form1"
    MaxButton = 0              'False
    MinButton = 0              'False
    ControlBox = 0              'False
    ClientLeft   = 45
    ClientTop    = 435
    ClientWidth  = 10185
    ClientHeight = 7455
    BeginProperty Font
        Name          = "Times New Roman"
        Size          = 9
        Charset       = 0
        Weight        = 400
        Underline     = 0              'False
        Italic        = 0              'False
        Strikethrough = 0              'False
    EndProperty
    StartupPosition = 1
    Begin VB.CommandButton btnFermer
        Caption = "&Fermer"
        Left   = 8865
        Top    = 7110
        Width  = 1275
        Height = 285
        TabIndex = 5
        Cancel = -1              'True
    End
    Begin VB.CommandButton btnEnregistrer
        Caption = "&Enregistrer"
        Left   = 7515
        Top    = 7110
        Width  = 1275
        Height = 285
        TabIndex = 4
    End
    Begin VB.TextBox fldstrhistorique
        Left   = 90
        Top    = 3825
        Width  = 10050
        Height = 3210
        TabIndex = 1
        MultiLine = -1              'True
        ScrollBars = 2
    End
    Begin VB.TextBox fldstrbackground
        Left   = 90
        Top    = 315
        Width  = 10050
        Height = 3165
        TabIndex = 0
        MultiLine = -1              'True
        ScrollBars = 2
    End
    Begin VB.Label Label2
        Caption = "Historique du personnage"
        Left   = 240
        Top    = 3510
        Width  = 2265
        Height = 330
        TabIndex = 3
    End
    Begin VB.Label Label1
        Caption = "Background du personnage"
        Left   = 225
        Top    = 45
        Width  = 2400
        Height = 285
        TabIndex = 2
    End
End
'Event for Form
Private Sub Form_Load()
'006cd0c0    55                      push ebp
'006cd0c1    8bec                    mov ebp, esp
'006cd0c3    83ec0c                  sub esp, 0c

' *** Reference to "__vbaExceptHandler"
'006cd0c6    6866724000              push 00407266
'006cd0cb    64a100000000            mov ax, word ptr fs:[00000000]
'006cd0d1    50                      push eax
    'Reference to '__except_list'
'006cd0d2    64892500000000          mov dword ptr fs:[00000000], esp
'006cd0d9    81ecf8000000            sub esp, 000000f8
'006cd0df    53                      push ebx
'006cd0e0    56                      push esi
'006cd0e1    57                      push edi
'006cd0e2    8965f4                  mov dword ptr [ebp-0c], esp
'006cd0e5    c745f8f8674000          mov dword ptr [ebp-08], 004067f8
'006cd0ec    8b4508                  mov eax, dword ptr [ebp+08]
'006cd0ef    8bc8                    mov ecx, eax
'006cd0f1    83e101                  and ecx, 01
'006cd0f4    894dfc                  mov dword ptr [ebp-04], ecx
'006cd0f7    83e0fe                  and eax, fffffffe
'006cd0fa    8b10                    mov edx, dword ptr [eax]
'006cd0fc    50                      push eax
'006cd0fd    894508                  mov dword ptr [ebp+08], eax
'006cd100    ff5204                  call dword ptr [edx+04]
'006cd103    a184b07200              mov ax, word ptr [0072b084]
'006cd108    33db                    xor ebx, ebx
'006cd10a    3bc3                    cmp eax, ebx
'006cd10c    895de8                  mov dword ptr [ebp-18], ebx
'006cd10f    895de4                  mov dword ptr [ebp-1c], ebx
'006cd112    895de0                  mov dword ptr [ebp-20], ebx
'006cd115    895ddc                  mov dword ptr [ebp-24], ebx
'006cd118    895dd8                  mov dword ptr [ebp-28], ebx
'006cd11b    895dd4                  mov dword ptr [ebp-2c], ebx
'006cd11e    895dd0                  mov dword ptr [ebp-30], ebx
'006cd121    895dcc                  mov dword ptr [ebp-34], ebx
'006cd124    895dc8                  mov dword ptr [ebp-38], ebx
'006cd127    895dc4                  mov dword ptr [ebp-3c], ebx
'006cd12a    895dc0                  mov dword ptr [ebp-40], ebx
'006cd12d    895dbc                  mov dword ptr [ebp-44], ebx
'006cd130    895dac                  mov dword ptr [ebp-54], ebx
'006cd133    895d9c                  mov dword ptr [ebp-64], ebx
'006cd136    895d8c                  mov dword ptr [ebp-74], ebx
'006cd139    899d7cffffff            mov dword ptr [ebp+ffffff7c], ebx
'006cd13f    899d6cffffff            mov dword ptr [ebp+ffffff6c], ebx
'006cd145    899d5cffffff            mov dword ptr [ebp+ffffff5c], ebx
'006cd14b    899d4cffffff            mov dword ptr [ebp+ffffff4c], ebx
'006cd151    899d3cffffff            mov dword ptr [ebp+ffffff3c], ebx
'006cd157    899d38ffffff            mov dword ptr [ebp+ffffff38], ebx
'006cd15d    7515                    jne 6cd174

If (0 = 0) Then
'006cd15f    6884b07200              push 0072b084
'006cd164    68f8894100              push 004189f8

' *** Reference to "__vbaNew2"
'006cd169    ff152c124000            call dword ptr [0040122c]
    Dim var_28 As New frmGestPerso
'006cd16f    a184b07200              mov ax, word ptr [0072b084]
    
End If
'006cd174    8b08                    mov ecx, dword ptr [eax]
'006cd176    50                      push eax
'006cd177    ff91ac030000            call dword ptr [ecx+000003ac]

' *** Reference to "__vbaObjSet"
'006cd17d    8b3db4104000            mov edi, dword ptr [004010b4]
'006cd183    50                      push eax
'006cd184    8d55cc                  lea edx, dword ptr [ebp-34]
'006cd187    52                      push edx
'006cd188    ffd7                    call edi
Set var_43 = Nothing
'006cd18a    8d4de4                  lea ecx, dword ptr [ebp-1c]
'006cd18d    8bf0                    mov esi, eax
'006cd18f    8b06                    mov eax, dword ptr [esi]
'006cd191    51                      push ecx
'006cd192    56                      push esi
'006cd193    ff90a0000000            call dword ptr [eax+000000a0]
'006cd199    dbe2                    fnclex
'006cd19b    3bc3                    cmp eax, ebx
'006cd19d    7d12                    jge 6cd1b1

If (-256 - 12 < 0) Then
'006cd19f    68a0000000              push 000000a0
'006cd1a4    68d00d4300              push 00430dd0
'006cd1a9    56                      push esi
'006cd1aa    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006cd1ab    ff1580104000            call dword ptr [00401080]
    
End If
'006cd1b1    8b55e4                  mov edx, dword ptr [ebp-1c]
'006cd1b4    52                      push edx
'006cd1b5    68cc134300              push 004313cc

' *** Reference to "__vbaStrCmp"
'006cd1ba    ff1530114000            call dword ptr [00401130]
'006cd1c0    8bf0                    mov esi, eax
'006cd1c2    f7de                    neg esi
'006cd1c4    1bf6                    sbb esi, esi
'006cd1c6    f7de                    neg esi
'006cd1c8    8d4de4                  lea ecx, dword ptr [ebp-1c]
'006cd1cb    f7de                    neg esi

' *** Reference to "__vbaFreeStr"
'006cd1cd    ff150c134000            call dword ptr [0040130c]
'#Cleanup(var_40)
'006cd1d3    8d4dcc                  lea ecx, dword ptr [ebp-34]

' *** Reference to "__vbaFreeObj"
'006cd1d6    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_43)
'006cd1dc    663bf3                  cmp si, bx
'006cd1df    0f8446060000            je 6cd82b

If (((vbNullString) <> (vbNullChar))) Then
'006cd1e5    a184b07200              mov ax, word ptr [0072b084]
'006cd1ea    3bc3                    cmp eax, ebx
'006cd1ec    7515                    jne 6cd203
    
    If (    var_28 = 0) Then
'006cd1ee    6884b07200              push 0072b084
'006cd1f3    68f8894100              push 004189f8

' *** Reference to "__vbaNew2"
'006cd1f8    ff152c124000            call dword ptr [0040122c]
    Set var_28 = New frmGestPerso
'006cd1fe    a184b07200              mov ax, word ptr [0072b084]
    
End If
'006cd203    8b08                    mov ecx, dword ptr [eax]
'006cd205    50                      push eax
'006cd206    ff91ac030000            call dword ptr [ecx+000003ac]
'006cd20c    50                      push eax
'006cd20d    8d55cc                  lea edx, dword ptr [ebp-34]
'006cd210    52                      push edx
'006cd211    ffd7                    call edi
Set var_43 = var_28
'006cd213    8d4de4                  lea ecx, dword ptr [ebp-1c]
'006cd216    8bf0                    mov esi, eax
'006cd218    8b06                    mov eax, dword ptr [esi]
'006cd21a    51                      push ecx
'006cd21b    56                      push esi
'006cd21c    ff90a0000000            call dword ptr [eax+000000a0]
'006cd222    dbe2                    fnclex
'006cd224    3bc3                    cmp eax, ebx
'006cd226    7d12                    jge 6cd23a

If (frmGestPerso < 0) Then
'006cd228    68a0000000              push 000000a0
'006cd22d    68d00d4300              push 00430dd0
'006cd232    56                      push esi
'006cd233    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006cd234    ff1580104000            call dword ptr [00401080]
    
End If
'006cd23a    8b55e4                  mov edx, dword ptr [ebp-1c]

' *** Reference to "__vbaStrMove"
'006cd23d    8b35d0124000            mov esi, dword ptr [004012d0]
'006cd243    8d4de0                  lea ecx, dword ptr [ebp-20]
'006cd246    895de4                  mov dword ptr [ebp-1c], ebx
'006cd249    ffd6                    call esi
'006cd24b    8d55e0                  lea edx, dword ptr [ebp-20]
'006cd24e    52                      push edx
'006cd24f    e89c69e2ff              call 4f3bf0
Call sub_4F3BF0()
'006cd254    8bd0                    mov edx, eax
'006cd256    8d4dd0                  lea ecx, dword ptr [ebp-30]
'006cd259    ffd6                    call esi
'006cd25b    8b55d0                  mov edx, dword ptr [ebp-30]
'006cd25e    899508ffffff            mov dword ptr [ebp+ffffff08], edx
'006cd264    8b1548b07200            mov edx, dword ptr [0072b048]
'006cd26a    b804000280              mov eax, 80020004
'006cd26f    898554ffffff            mov dword ptr [ebp+ffffff54], eax
'006cd275    b90a000000              mov ecx, 0000000a
'006cd27a    898d4cffffff            mov dword ptr [ebp+ffffff4c], ecx
'006cd280    898d5cffffff            mov dword ptr [ebp+ffffff5c], ecx
'006cd286    895dd0                  mov dword ptr [ebp-30], ebx
'006cd289    8b1a                    mov ebx, dword ptr [edx]
'006cd28b    8d55c8                  lea edx, dword ptr [ebp-38]
'006cd28e    52                      push edx
'006cd28f    83ec10                  sub esp, 10
'006cd292    8bd4                    mov edx, esp
'006cd294    890a                    mov dword ptr [edx], ecx
'006cd296    8b8d50ffffff            mov ecx, dword ptr [ebp+ffffff50]
'006cd29c    894a04                  mov dword ptr [edx+04], ecx
'006cd29f    894208                  mov dword ptr [edx+08], eax
'006cd2a2    8bf8                    mov edi, eax
'006cd2a4    8b8558ffffff            mov eax, dword ptr [ebp+ffffff58]
'006cd2aa    89420c                  mov dword ptr [edx+0c], eax
'006cd2ad    8b955cffffff            mov edx, dword ptr [ebp+ffffff5c]
'006cd2b3    8b8560ffffff            mov eax, dword ptr [ebp+ffffff60]
'006cd2b9    83ec10                  sub esp, 10
'006cd2bc    8bcc                    mov ecx, esp
'006cd2be    8911                    mov dword ptr [ecx], edx
'006cd2c0    8b9568ffffff            mov edx, dword ptr [ebp+ffffff68]
'006cd2c6    894104                  mov dword ptr [ecx+04], eax
'006cd2c9    897908                  mov dword ptr [ecx+08], edi
'006cd2cc    89510c                  mov dword ptr [ecx+0c], edx
'006cd2cf    8b9570ffffff            mov edx, dword ptr [ebp+ffffff70]
'006cd2d5    83ec10                  sub esp, 10
'006cd2d8    8bcc                    mov ecx, esp
'006cd2da    b803000000              mov eax, 00000003
'006cd2df    8901                    mov dword ptr [ecx], eax
'006cd2e1    895104                  mov dword ptr [ecx+04], edx
'006cd2e4    8b9508ffffff            mov edx, dword ptr [ebp+ffffff08]
'006cd2ea    b802000000              mov eax, 00000002
'006cd2ef    894108                  mov dword ptr [ecx+08], eax
'006cd2f2    8b8578ffffff            mov eax, dword ptr [ebp+ffffff78]
'006cd2f8    89410c                  mov dword ptr [ecx+0c], eax
'006cd2fb    68883a4500              push 00453a88
'006cd300    8d4ddc                  lea ecx, dword ptr [ebp-24]
'006cd303    ffd6                    call esi

' *** Reference to "__vbaStrCat"
'006cd305    8b3d70104000            mov edi, dword ptr [00401070]
'006cd30b    50                      push eax
'006cd30c    ffd7                    call edi
var_49 = ("select background,histoire from personnage where nom='") & (vbNullString)
'006cd30e    8bd0                    mov edx, eax
'006cd310    8d4dd8                  lea ecx, dword ptr [ebp-28]
'006cd313    ffd6                    call esi
'006cd315    50                      push eax
'006cd316    6854a44300              push 0043a454
'006cd31b    ffd7                    call edi
var_76 = (var_49) & ("'")
'006cd31d    8bd0                    mov edx, eax
'006cd31f    8d4dd4                  lea ecx, dword ptr [ebp-2c]
'006cd322    ffd6                    call esi
'006cd324    8b0d48b07200            mov ecx, dword ptr [0072b048]
'006cd32a    50                      push eax
'006cd32b    51                      push ecx
'006cd32c    ff93bc000000            call dword ptr [ebx+000000bc]
'006cd332    dbe2                    fnclex
'006cd334    33db                    xor ebx, ebx
var_num2 = Empty
'006cd336    3bc3                    cmp eax, ebx
'006cd338    7d1c                    jge 6cd356

If (-4512 < -256 - 12) Then
'006cd33a    8b1548b07200            mov edx, dword ptr [0072b048]

' *** Reference to "__vbaHresultCheckObj"
'006cd340    8b3580104000            mov esi, dword ptr [00401080]
'006cd346    68bc000000              push 000000bc
'006cd34b    68301f4300              push 00431f30
'006cd350    52                      push edx
'006cd351    50                      push eax
'006cd352    ffd6                    call esi
'006cd354    eb06                    jmp 6cd35c
    
End If

' *** Reference to "__vbaHresultCheckObj"
'006cd356    8b3580104000            mov esi, dword ptr [00401080]
'006cd35c    8b45c8                  mov eax, dword ptr [ebp-38]
'006cd35f    50                      push eax
'006cd360    8d45e8                  lea eax, dword ptr [ebp-18]
'006cd363    50                      push eax
'006cd364    895dc8                  mov dword ptr [ebp-38], ebx

' *** Reference to "__vbaObjSet"
'006cd367    ff15b4104000            call dword ptr [004010b4]
Set var_41 = Nothing
'006cd36d    8d4dd0                  lea ecx, dword ptr [ebp-30]
'006cd370    51                      push ecx
'006cd371    8d55d4                  lea edx, dword ptr [ebp-2c]
'006cd374    52                      push edx
'006cd375    8d45d8                  lea eax, dword ptr [ebp-28]
'006cd378    50                      push eax
'006cd379    8d4ddc                  lea ecx, dword ptr [ebp-24]
'006cd37c    51                      push ecx
'006cd37d    8d55e0                  lea edx, dword ptr [ebp-20]
'006cd380    52                      push edx
'006cd381    6a05                    push 05

' *** Reference to "__vbaFreeStrList"
'006cd383    ff155c124000            call dword ptr [0040125c]
'#Cleanup( 0, 0, -4508, -4512, 0)
'006cd389    83c418                  add esp, 18
'006cd38c    8d4dcc                  lea ecx, dword ptr [ebp-34]

' *** Reference to "__vbaFreeObj"
'006cd38f    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_43)
'006cd395    8b45e8                  mov eax, dword ptr [ebp-18]
'006cd398    8b08                    mov ecx, dword ptr [eax]
'006cd39a    8d9538ffffff            lea edx, dword ptr [ebp+ffffff38]
'006cd3a0    52                      push edx
'006cd3a1    50                      push eax
'006cd3a2    ff5134                  call dword ptr [ecx+34]
'006cd3a5    dbe2                    fnclex
'006cd3a7    3bc3                    cmp eax, ebx
'006cd3a9    7d0e                    jge 6cd3b9

If (var_41 < -256 - 12) Then
'006cd3ab    8b4de8                  mov ecx, dword ptr [ebp-18]
'006cd3ae    6a34                    push 34
'006cd3b0    6830314300              push 00433130
'006cd3b5    51                      push ecx
'006cd3b6    50                      push eax
'006cd3b7    ffd6                    call esi
    
End If
'006cd3b9    66399d38ffffff          cmp word ptr [ebp+ffffff38], bx
'006cd3c0    0f8542040000            jne 6cd808

If (0 = -256 - 12) Then
'006cd3c6    8b45e8                  mov eax, dword ptr [ebp-18]
'006cd3c9    8b10                    mov edx, dword ptr [eax]
'006cd3cb    8d4dcc                  lea ecx, dword ptr [ebp-34]
'006cd3ce    51                      push ecx
'006cd3cf    50                      push eax
'006cd3d0    ff92b4000000            call dword ptr [edx+000000b4]
'006cd3d6    dbe2                    fnclex
'006cd3d8    3bc3                    cmp eax, ebx
'006cd3da    7d11                    jge 6cd3ed
    
    If (    var_41 < -256 - 12) Then
'006cd3dc    8b55e8                  mov edx, dword ptr [ebp-18]
'006cd3df    68b4000000              push 000000b4
'006cd3e4    6830314300              push 00433130
'006cd3e9    52                      push edx
'006cd3ea    50                      push eax
'006cd3eb    ffd6                    call esi
    
End If
'006cd3ed    8b45cc                  mov eax, dword ptr [ebp-34]
'006cd3f0    8b30                    mov esi, dword ptr [eax]
'006cd3f2    8d7dc8                  lea edi, dword ptr [ebp-38]
'006cd3f5    57                      push edi
'006cd3f6    83ec10                  sub esp, 10
'006cd3f9    8bfc                    mov edi, esp
'006cd3fb    ba08000000              mov edx, 00000008
'006cd400    8917                    mov dword ptr [edi], edx
'006cd402    8b9570ffffff            mov edx, dword ptr [ebp+ffffff70]
'006cd408    895704                  mov dword ptr [edi+04], edx
'006cd40b    b9ac174400              mov ecx, 004417ac
'006cd410    894f08                  mov dword ptr [edi+08], ecx
'006cd413    8b8d78ffffff            mov ecx, dword ptr [ebp+ffffff78]
'006cd419    50                      push eax
'006cd41a    898530ffffff            mov dword ptr [ebp+ffffff30], eax
'006cd420    894f0c                  mov dword ptr [edi+0c], ecx
'006cd423    ff5630                  call dword ptr [esi+30]
'006cd426    dbe2                    fnclex
'006cd428    3bc3                    cmp eax, ebx
'006cd42a    7d15                    jge 6cd441

If (var_43 < -256 - 12) Then
'006cd42c    8b9530ffffff            mov edx, dword ptr [ebp+ffffff30]
'006cd432    6a30                    push 30
'006cd434    68d8304300              push 004330d8
'006cd439    52                      push edx
'006cd43a    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006cd43b    ff1580104000            call dword ptr [00401080]
    
End If
'006cd441    8b45c8                  mov eax, dword ptr [ebp-38]
'006cd444    8945b4                  mov dword ptr [ebp-4c], eax
'006cd447    8d45ac                  lea eax, dword ptr [ebp-54]
'006cd44a    50                      push eax
'006cd44b    895dc8                  mov dword ptr [ebp-38], ebx
'006cd44e    c745ac09000000          mov dword ptr [ebp-54], 00000009

' *** Reference to "rtcIsNull"
'006cd455    ff1540114000            call dword ptr [00401140]
'006cd45b    898538ffffff            mov dword ptr [ebp+ffffff38], eax
'006cd461    8b4508                  mov eax, dword ptr [ebp+08]
'006cd464    8b08                    mov ecx, dword ptr [eax]
'006cd466    50                      push eax
'006cd467    ff9108030000            call dword ptr [ecx+00000308]
'006cd46d    50                      push eax
'006cd46e    8d55bc                  lea edx, dword ptr [ebp-44]
'006cd471    52                      push edx

' *** Reference to "__vbaObjSet"
'006cd472    ff15b4104000            call dword ptr [004010b4]
Set var_58 = Me
'006cd478    8d55c4                  lea edx, dword ptr [ebp-3c]
'006cd47b    8bf0                    mov esi, eax
'006cd47d    8b45e8                  mov eax, dword ptr [ebp-18]
'006cd480    8b08                    mov ecx, dword ptr [eax]
'006cd482    52                      push edx
'006cd483    50                      push eax
'006cd484    ff91b4000000            call dword ptr [ecx+000000b4]
'006cd48a    dbe2                    fnclex
'006cd48c    3bc3                    cmp eax, ebx
'006cd48e    7d15                    jge 6cd4a5

If (var_41 < -256 - 12) Then
'006cd490    8b4de8                  mov ecx, dword ptr [ebp-18]
'006cd493    68b4000000              push 000000b4
'006cd498    6830314300              push 00433130
'006cd49d    51                      push ecx
'006cd49e    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006cd49f    ff1580104000            call dword ptr [00401080]
    
End If
'006cd4a5    8b45c4                  mov eax, dword ptr [ebp-3c]
'006cd4a8    8b10                    mov edx, dword ptr [eax]
'006cd4aa    8d7dc0                  lea edi, dword ptr [ebp-40]
'006cd4ad    57                      push edi
'006cd4ae    83ec10                  sub esp, 10
'006cd4b1    8bfc                    mov edi, esp
'006cd4b3    b908000000              mov ecx, 00000008
'006cd4b8    890f                    mov dword ptr [edi], ecx
'006cd4ba    8b8d60ffffff            mov ecx, dword ptr [ebp+ffffff60]
'006cd4c0    894f04                  mov dword ptr [edi+04], ecx
'006cd4c3    c78564ffffffac174400    mov dword ptr [ebp+ffffff64], 004417ac
'006cd4cd    8b8d64ffffff            mov ecx, dword ptr [ebp+ffffff64]
'006cd4d3    894f08                  mov dword ptr [edi+08], ecx
'006cd4d6    8b8d68ffffff            mov ecx, dword ptr [ebp+ffffff68]
'006cd4dc    50                      push eax
'006cd4dd    898524ffffff            mov dword ptr [ebp+ffffff24], eax
'006cd4e3    894f0c                  mov dword ptr [edi+0c], ecx
'006cd4e6    ff5230                  call dword ptr [edx+30]
'006cd4e9    dbe2                    fnclex
'006cd4eb    3bc3                    cmp eax, ebx
'006cd4ed    7d15                    jge 6cd504

If (0 < -256 - 12) Then
'006cd4ef    8b9524ffffff            mov edx, dword ptr [ebp+ffffff24]
'006cd4f5    6a30                    push 30
'006cd4f7    68d8304300              push 004330d8
'006cd4fc    52                      push edx
'006cd4fd    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006cd4fe    ff1580104000            call dword ptr [00401080]
    
End If
'006cd504    8b45c0                  mov eax, dword ptr [ebp-40]
'006cd507    8d953cffffff            lea edx, dword ptr [ebp+ffffff3c]
'006cd50d    8d4d9c                  lea ecx, dword ptr [ebp-64]
'006cd510    895dc0                  mov dword ptr [ebp-40], ebx
'006cd513    894594                  mov dword ptr [ebp-6c], eax
'006cd516    c7458c09000000          mov dword ptr [ebp-74], 00000009
'006cd51d    c78544ffffffcc134300    mov dword ptr [ebp+ffffff44], 004313cc
'006cd527    c7853cffffff08000000    mov dword ptr [ebp+ffffff3c], 00000008

' *** Reference to "__vbaVarDup"
'006cd531    ff1598124000            call dword ptr [00401298]
var_51 = (vbNullChar)
'006cd537    668b8538ffffff          mov ax, word ptr [ebp+ffffff38]
'006cd53e    8d4d8c                  lea ecx, dword ptr [ebp-74]
'006cd541    51                      push ecx
'006cd542    8d559c                  lea edx, dword ptr [ebp-64]
'006cd545    66898554ffffff          mov word ptr [ebp+ffffff54], ax
'006cd54c    52                      push edx
'006cd54d    8d854cffffff            lea eax, dword ptr [ebp+ffffff4c]
'006cd553    50                      push eax
'006cd554    8d8d7cffffff            lea ecx, dword ptr [ebp+ffffff7c]
'006cd55a    51                      push ecx
'006cd55b    c7854cffffff0b000000    mov dword ptr [ebp+ffffff4c], 0000000b

' *** Reference to "rtcImmediateIf"
'006cd565    ff154c124000            call dword ptr [0040124c]
'006cd56b    8b3e                    mov edi, dword ptr [esi]
'006cd56d    8d957cffffff            lea edx, dword ptr [ebp+ffffff7c]
'006cd573    52                      push edx
'006cd574    8d45e4                  lea eax, dword ptr [ebp-1c]
'006cd577    50                      push eax

' *** Reference to "__vbaStrVarVal"
'006cd578    ff15fc114000            call dword ptr [004011fc]
'006cd57e    50                      push eax
'006cd57f    56                      push esi
'006cd580    ff97a4000000            call dword ptr [edi+000000a4]
'006cd586    dbe2                    fnclex
'006cd588    3bc3                    cmp eax, ebx
'006cd58a    7d12                    jge 6cd59e
'006cd58c    68a4000000              push 000000a4
'006cd591    68d00d4300              push 00430dd0
'006cd596    56                      push esi
'006cd597    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006cd598    ff1580104000            call dword ptr [00401080]
'006cd59e    8d4de4                  lea ecx, dword ptr [ebp-1c]

' *** Reference to "__vbaFreeStr"
'006cd5a1    ff150c134000            call dword ptr [0040130c]
'#Cleanup(var_40)
'006cd5a7    8d4dbc                  lea ecx, dword ptr [ebp-44]
'006cd5aa    51                      push ecx
'006cd5ab    8d55c4                  lea edx, dword ptr [ebp-3c]
'006cd5ae    52                      push edx
'006cd5af    8d45cc                  lea eax, dword ptr [ebp-34]
'006cd5b2    50                      push eax
'006cd5b3    6a03                    push 03

' *** Reference to "__vbaFreeObjList"
'006cd5b5    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_43, var_9, var_58)
'006cd5bb    8d8d7cffffff            lea ecx, dword ptr [ebp+ffffff7c]
'006cd5c1    51                      push ecx
'006cd5c2    8d558c                  lea edx, dword ptr [ebp-74]
'006cd5c5    52                      push edx
'006cd5c6    8d459c                  lea eax, dword ptr [ebp-64]
'006cd5c9    50                      push eax
'006cd5ca    8d8d4cffffff            lea ecx, dword ptr [ebp+ffffff4c]
'006cd5d0    51                      push ecx
'006cd5d1    8d55ac                  lea edx, dword ptr [ebp-54]
'006cd5d4    52                      push edx
'006cd5d5    6a05                    push 05

' *** Reference to "__vbaFreeVarList"
'006cd5d7    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_50, var_89, var_51, var_52, var_59)
'006cd5dd    8b45e8                  mov eax, dword ptr [ebp-18]
'006cd5e0    8b08                    mov ecx, dword ptr [eax]
'006cd5e2    83c428                  add esp, 28
'006cd5e5    8d55cc                  lea edx, dword ptr [ebp-34]
'006cd5e8    52                      push edx
'006cd5e9    50                      push eax
'006cd5ea    ff91b4000000            call dword ptr [ecx+000000b4]
'006cd5f0    dbe2                    fnclex
'006cd5f2    3bc3                    cmp eax, ebx
'006cd5f4    7d15                    jge 6cd60b

If (var_41 < -256 - 12) Then
'006cd5f6    8b4de8                  mov ecx, dword ptr [ebp-18]
'006cd5f9    68b4000000              push 000000b4
'006cd5fe    6830314300              push 00433130
'006cd603    51                      push ecx
'006cd604    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006cd605    ff1580104000            call dword ptr [00401080]
    
End If
'006cd60b    8b45cc                  mov eax, dword ptr [ebp-34]
'006cd60e    8b30                    mov esi, dword ptr [eax]
'006cd610    8d7dc8                  lea edi, dword ptr [ebp-38]
'006cd613    57                      push edi
'006cd614    83ec10                  sub esp, 10
'006cd617    8bfc                    mov edi, esp
'006cd619    ba08000000              mov edx, 00000008
'006cd61e    8917                    mov dword ptr [edi], edx
'006cd620    8b9570ffffff            mov edx, dword ptr [ebp+ffffff70]
'006cd626    895704                  mov dword ptr [edi+04], edx
'006cd629    b9c8174400              mov ecx, 004417c8
'006cd62e    894f08                  mov dword ptr [edi+08], ecx
'006cd631    8b8d78ffffff            mov ecx, dword ptr [ebp+ffffff78]
'006cd637    50                      push eax
'006cd638    898530ffffff            mov dword ptr [ebp+ffffff30], eax
'006cd63e    894f0c                  mov dword ptr [edi+0c], ecx
'006cd641    ff5630                  call dword ptr [esi+30]
'006cd644    dbe2                    fnclex
'006cd646    3bc3                    cmp eax, ebx
'006cd648    7d15                    jge 6cd65f

If (var_43 < -256 - 12) Then
'006cd64a    8b9530ffffff            mov edx, dword ptr [ebp+ffffff30]
'006cd650    6a30                    push 30
'006cd652    68d8304300              push 004330d8
'006cd657    52                      push edx
'006cd658    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006cd659    ff1580104000            call dword ptr [00401080]
    
End If
'006cd65f    8b45c8                  mov eax, dword ptr [ebp-38]
'006cd662    8945b4                  mov dword ptr [ebp-4c], eax
'006cd665    8d45ac                  lea eax, dword ptr [ebp-54]
'006cd668    50                      push eax
'006cd669    895dc8                  mov dword ptr [ebp-38], ebx
'006cd66c    c745ac09000000          mov dword ptr [ebp-54], 00000009

' *** Reference to "rtcIsNull"
'006cd673    ff1540114000            call dword ptr [00401140]
'006cd679    898538ffffff            mov dword ptr [ebp+ffffff38], eax
'006cd67f    8b4508                  mov eax, dword ptr [ebp+08]
'006cd682    8b08                    mov ecx, dword ptr [eax]
'006cd684    50                      push eax
'006cd685    ff9104030000            call dword ptr [ecx+00000304]
'006cd68b    50                      push eax
'006cd68c    8d55bc                  lea edx, dword ptr [ebp-44]
'006cd68f    52                      push edx

' *** Reference to "__vbaObjSet"
'006cd690    ff15b4104000            call dword ptr [004010b4]
Set var_58 = Me
'006cd696    8d55c4                  lea edx, dword ptr [ebp-3c]
'006cd699    8bf0                    mov esi, eax
'006cd69b    8b45e8                  mov eax, dword ptr [ebp-18]
'006cd69e    8b08                    mov ecx, dword ptr [eax]
'006cd6a0    52                      push edx
'006cd6a1    50                      push eax
'006cd6a2    ff91b4000000            call dword ptr [ecx+000000b4]
'006cd6a8    dbe2                    fnclex
'006cd6aa    3bc3                    cmp eax, ebx
'006cd6ac    7d15                    jge 6cd6c3

If (var_41 < -256 - 12) Then
'006cd6ae    8b4de8                  mov ecx, dword ptr [ebp-18]
'006cd6b1    68b4000000              push 000000b4
'006cd6b6    6830314300              push 00433130
'006cd6bb    51                      push ecx
'006cd6bc    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006cd6bd    ff1580104000            call dword ptr [00401080]
    
End If
'006cd6c3    8b45c4                  mov eax, dword ptr [ebp-3c]
'006cd6c6    8b10                    mov edx, dword ptr [eax]
'006cd6c8    8d7dc0                  lea edi, dword ptr [ebp-40]
'006cd6cb    57                      push edi
'006cd6cc    83ec10                  sub esp, 10
'006cd6cf    8bfc                    mov edi, esp
'006cd6d1    b908000000              mov ecx, 00000008
'006cd6d6    890f                    mov dword ptr [edi], ecx
'006cd6d8    8b8d60ffffff            mov ecx, dword ptr [ebp+ffffff60]
'006cd6de    894f04                  mov dword ptr [edi+04], ecx
'006cd6e1    c78564ffffffc8174400    mov dword ptr [ebp+ffffff64], 004417c8
'006cd6eb    8b8d64ffffff            mov ecx, dword ptr [ebp+ffffff64]
'006cd6f1    894f08                  mov dword ptr [edi+08], ecx
'006cd6f4    8b8d68ffffff            mov ecx, dword ptr [ebp+ffffff68]
'006cd6fa    50                      push eax
'006cd6fb    898524ffffff            mov dword ptr [ebp+ffffff24], eax
'006cd701    894f0c                  mov dword ptr [edi+0c], ecx
'006cd704    ff5230                  call dword ptr [edx+30]
'006cd707    dbe2                    fnclex
'006cd709    3bc3                    cmp eax, ebx
'006cd70b    7d15                    jge 6cd722

If (0 < -256 - 12) Then
'006cd70d    8b9524ffffff            mov edx, dword ptr [ebp+ffffff24]
'006cd713    6a30                    push 30
'006cd715    68d8304300              push 004330d8
'006cd71a    52                      push edx
'006cd71b    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006cd71c    ff1580104000            call dword ptr [00401080]
    
End If
'006cd722    8b45c0                  mov eax, dword ptr [ebp-40]
'006cd725    8d953cffffff            lea edx, dword ptr [ebp+ffffff3c]
'006cd72b    8d4d9c                  lea ecx, dword ptr [ebp-64]
'006cd72e    895dc0                  mov dword ptr [ebp-40], ebx
'006cd731    894594                  mov dword ptr [ebp-6c], eax
'006cd734    c7458c09000000          mov dword ptr [ebp-74], 00000009
'006cd73b    c78544ffffffcc134300    mov dword ptr [ebp+ffffff44], 004313cc
'006cd745    c7853cffffff08000000    mov dword ptr [ebp+ffffff3c], 00000008

' *** Reference to "__vbaVarDup"
'006cd74f    ff1598124000            call dword ptr [00401298]
var_51 = (vbNullChar)
'006cd755    668b8538ffffff          mov ax, word ptr [ebp+ffffff38]
'006cd75c    8d4d8c                  lea ecx, dword ptr [ebp-74]
'006cd75f    51                      push ecx
'006cd760    8d559c                  lea edx, dword ptr [ebp-64]
'006cd763    66898554ffffff          mov word ptr [ebp+ffffff54], ax
'006cd76a    52                      push edx
'006cd76b    8d854cffffff            lea eax, dword ptr [ebp+ffffff4c]
'006cd771    50                      push eax
'006cd772    8d8d7cffffff            lea ecx, dword ptr [ebp+ffffff7c]
'006cd778    51                      push ecx
'006cd779    c7854cffffff0b000000    mov dword ptr [ebp+ffffff4c], 0000000b

' *** Reference to "rtcImmediateIf"
'006cd783    ff154c124000            call dword ptr [0040124c]
'006cd789    8b3e                    mov edi, dword ptr [esi]
'006cd78b    8d957cffffff            lea edx, dword ptr [ebp+ffffff7c]
'006cd791    52                      push edx
'006cd792    8d45e4                  lea eax, dword ptr [ebp-1c]
'006cd795    50                      push eax

' *** Reference to "__vbaStrVarVal"
'006cd796    ff15fc114000            call dword ptr [004011fc]
'006cd79c    50                      push eax
'006cd79d    56                      push esi
'006cd79e    ff97a4000000            call dword ptr [edi+000000a4]
'006cd7a4    dbe2                    fnclex
'006cd7a6    3bc3                    cmp eax, ebx
'006cd7a8    7d16                    jge 6cd7c0
'006cd7aa    68a4000000              push 000000a4
'006cd7af    68d00d4300              push 00430dd0
'006cd7b4    56                      push esi

' *** Reference to "__vbaHresultCheckObj"
'006cd7b5    8b3580104000            mov esi, dword ptr [00401080]
'006cd7bb    50                      push eax
'006cd7bc    ffd6                    call esi
'006cd7be    eb06                    jmp 6cd7c6

' *** Reference to "__vbaHresultCheckObj"
'006cd7c0    8b3580104000            mov esi, dword ptr [00401080]
'006cd7c6    8d4de4                  lea ecx, dword ptr [ebp-1c]

' *** Reference to "__vbaFreeStr"
'006cd7c9    ff150c134000            call dword ptr [0040130c]
'#Cleanup(var_40)
'006cd7cf    8d4dbc                  lea ecx, dword ptr [ebp-44]
'006cd7d2    51                      push ecx
'006cd7d3    8d55c4                  lea edx, dword ptr [ebp-3c]
'006cd7d6    52                      push edx
'006cd7d7    8d45cc                  lea eax, dword ptr [ebp-34]
'006cd7da    50                      push eax
'006cd7db    6a03                    push 03

' *** Reference to "__vbaFreeObjList"
'006cd7dd    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_43, var_9, var_58)
'006cd7e3    8d8d7cffffff            lea ecx, dword ptr [ebp+ffffff7c]
'006cd7e9    51                      push ecx
'006cd7ea    8d558c                  lea edx, dword ptr [ebp-74]
'006cd7ed    52                      push edx
'006cd7ee    8d459c                  lea eax, dword ptr [ebp-64]
'006cd7f1    50                      push eax
'006cd7f2    8d8d4cffffff            lea ecx, dword ptr [ebp+ffffff4c]
'006cd7f8    51                      push ecx
'006cd7f9    8d55ac                  lea edx, dword ptr [ebp-54]
'006cd7fc    52                      push edx
'006cd7fd    6a05                    push 05

' *** Reference to "__vbaFreeVarList"
'006cd7ff    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_50, var_89, var_51, var_52, var_59)
'006cd805    83c428                  add esp, 28

'ERROR: Two many next close:
End If
'006cd808    8b45e8                  mov eax, dword ptr [ebp-18]
'006cd80b    8b08                    mov ecx, dword ptr [eax]
'006cd80d    50                      push eax
'006cd80e    ff91c4000000            call dword ptr [ecx+000000c4]
'006cd814    dbe2                    fnclex
'006cd816    3bc3                    cmp eax, ebx
'006cd818    7d11                    jge 6cd82b

If (var_41 < -256 - 12) Then
'006cd81a    8b55e8                  mov edx, dword ptr [ebp-18]
'006cd81d    68c4000000              push 000000c4
'006cd822    6830314300              push 00433130
'006cd827    52                      push edx
'006cd828    50                      push eax
'006cd829    ffd6                    call esi
    
End If
'006cd82b    895dfc                  mov dword ptr [ebp-04], ebx
'006cd82e    689ad86c00              push 006cd89a
'006cd833    eb5b                    jmp 6cd890
'006cd835    8d45d0                  lea eax, dword ptr [ebp-30]
'006cd838    50                      push eax
'006cd839    8d4dd4                  lea ecx, dword ptr [ebp-2c]
'006cd83c    51                      push ecx
'006cd83d    8d55d8                  lea edx, dword ptr [ebp-28]
'006cd840    52                      push edx
'006cd841    8d45dc                  lea eax, dword ptr [ebp-24]
'006cd844    50                      push eax
'006cd845    8d4de0                  lea ecx, dword ptr [ebp-20]
'006cd848    51                      push ecx
'006cd849    8d55e4                  lea edx, dword ptr [ebp-1c]
'006cd84c    52                      push edx
'006cd84d    6a06                    push 06

' *** Reference to "__vbaFreeStrList"
'006cd84f    ff155c124000            call dword ptr [0040125c]
'#Cleanup( , 0, 0, -4508, -4512, 0)
'006cd855    8d45bc                  lea eax, dword ptr [ebp-44]
'006cd858    50                      push eax
'006cd859    8d4dc0                  lea ecx, dword ptr [ebp-40]
'006cd85c    51                      push ecx
'006cd85d    8d55c4                  lea edx, dword ptr [ebp-3c]
'006cd860    52                      push edx
'006cd861    8d45c8                  lea eax, dword ptr [ebp-38]
'006cd864    50                      push eax
'006cd865    8d4dcc                  lea ecx, dword ptr [ebp-34]
'006cd868    51                      push ecx
'006cd869    6a05                    push 05

' *** Reference to "__vbaFreeObjList"
'006cd86b    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_43, var_46, var_9, var_5, var_58)
'006cd871    8d957cffffff            lea edx, dword ptr [ebp+ffffff7c]
'006cd877    52                      push edx
'006cd878    8d458c                  lea eax, dword ptr [ebp-74]
'006cd87b    50                      push eax
'006cd87c    8d4d9c                  lea ecx, dword ptr [ebp-64]
'006cd87f    51                      push ecx
'006cd880    8d55ac                  lea edx, dword ptr [ebp-54]
'006cd883    52                      push edx
'006cd884    6a04                    push 04

' *** Reference to "__vbaFreeVarList"
'006cd886    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_50, var_51, var_52, var_59)
'006cd88c    83c448                  add esp, 48
'006cd88f    c3                      ret
'006cd890    8d4de8                  lea ecx, dword ptr [ebp-18]

' *** Reference to "__vbaFreeObj"
'006cd893    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_41)
'006cd899    c3                      ret
'006cd89a    8b4508                  mov eax, dword ptr [ebp+08]
'006cd89d    8b08                    mov ecx, dword ptr [eax]
'006cd89f    50                      push eax
'006cd8a0    ff5108                  call dword ptr [ecx+08]
'006cd8a3    8b45fc                  mov eax, dword ptr [ebp-04]
'006cd8a6    8b4dec                  mov ecx, dword ptr [ebp-14]
'006cd8a9    5f                      pop edi
'006cd8aa    5e                      pop esi
    'Reference to '__except_list'
'006cd8ab    64890d00000000          mov dword ptr fs:[00000000], ecx
'006cd8b2    5b                      pop ebx
'006cd8b3    8be5                    mov esp, ebp
'006cd8b5    5d                      pop ebp
'006cd8b6    c20400                  ret 0004


End Sub


'Event for btnEnregistrer
Private Sub btnEnregistrer_Click()
'006cc900    55                      push ebp
'006cc901    8bec                    mov ebp, esp
'006cc903    83ec0c                  sub esp, 0c

' *** Reference to "__vbaExceptHandler"
'006cc906    6866724000              push 00407266
'006cc90b    64a100000000            mov ax, word ptr fs:[00000000]
'006cc911    50                      push eax
    'Reference to '__except_list'
'006cc912    64892500000000          mov dword ptr fs:[00000000], esp
'006cc919    81ecec000000            sub esp, 000000ec
'006cc91f    53                      push ebx
'006cc920    56                      push esi
'006cc921    57                      push edi
'006cc922    8965f4                  mov dword ptr [ebp-0c], esp
'006cc925    c745f8d8674000          mov dword ptr [ebp-08], 004067d8
'006cc92c    8b7508                  mov esi, dword ptr [ebp+08]
'006cc92f    8bc6                    mov eax, esi
'006cc931    83e001                  and eax, 01
'006cc934    8945fc                  mov dword ptr [ebp-04], eax
'006cc937    83e6fe                  and esi, fffffffe
'006cc93a    8b0e                    mov ecx, dword ptr [esi]
'006cc93c    56                      push esi
'006cc93d    897508                  mov dword ptr [ebp+08], esi
'006cc940    ff5104                  call dword ptr [ecx+04]
'006cc943    a184b07200              mov ax, word ptr [0072b084]
'006cc948    33db                    xor ebx, ebx
'006cc94a    3bc3                    cmp eax, ebx
'006cc94c    895de8                  mov dword ptr [ebp-18], ebx
'006cc94f    895de4                  mov dword ptr [ebp-1c], ebx
'006cc952    895de0                  mov dword ptr [ebp-20], ebx
'006cc955    895ddc                  mov dword ptr [ebp-24], ebx
'006cc958    895dd8                  mov dword ptr [ebp-28], ebx
'006cc95b    895dd4                  mov dword ptr [ebp-2c], ebx
'006cc95e    895dd0                  mov dword ptr [ebp-30], ebx
'006cc961    895dcc                  mov dword ptr [ebp-34], ebx
'006cc964    895dc8                  mov dword ptr [ebp-38], ebx
'006cc967    895db8                  mov dword ptr [ebp-48], ebx
'006cc96a    895da8                  mov dword ptr [ebp-58], ebx
'006cc96d    895d98                  mov dword ptr [ebp-68], ebx
'006cc970    895d88                  mov dword ptr [ebp-78], ebx
'006cc973    899d78ffffff            mov dword ptr [ebp+ffffff78], ebx
'006cc979    899d68ffffff            mov dword ptr [ebp+ffffff68], ebx
'006cc97f    899d58ffffff            mov dword ptr [ebp+ffffff58], ebx
'006cc985    899d48ffffff            mov dword ptr [ebp+ffffff48], ebx
'006cc98b    899d38ffffff            mov dword ptr [ebp+ffffff38], ebx
'006cc991    899d28ffffff            mov dword ptr [ebp+ffffff28], ebx
'006cc997    899d24ffffff            mov dword ptr [ebp+ffffff24], ebx
'006cc99d    7515                    jne 6cc9b4

If (0 = 0) Then
'006cc99f    6884b07200              push 0072b084
'006cc9a4    68f8894100              push 004189f8

' *** Reference to "__vbaNew2"
'006cc9a9    ff152c124000            call dword ptr [0040122c]
    Dim var_28 As New frmGestPerso
'006cc9af    a184b07200              mov ax, word ptr [0072b084]
    
End If
'006cc9b4    8b10                    mov edx, dword ptr [eax]
'006cc9b6    50                      push eax
'006cc9b7    ff92ac030000            call dword ptr [edx+000003ac]
'006cc9bd    50                      push eax
'006cc9be    8d45cc                  lea eax, dword ptr [ebp-34]
'006cc9c1    50                      push eax

' *** Reference to "__vbaObjSet"
'006cc9c2    ff15b4104000            call dword ptr [004010b4]
Set var_43 = Nothing
'006cc9c8    8d55e4                  lea edx, dword ptr [ebp-1c]
'006cc9cb    8bf8                    mov edi, eax
'006cc9cd    8b0f                    mov ecx, dword ptr [edi]
'006cc9cf    52                      push edx
'006cc9d0    57                      push edi
'006cc9d1    ff91a0000000            call dword ptr [ecx+000000a0]
'006cc9d7    dbe2                    fnclex
'006cc9d9    3bc3                    cmp eax, ebx
'006cc9db    7d12                    jge 6cc9ef

If (var_43 < 0) Then
'006cc9dd    68a0000000              push 000000a0
'006cc9e2    68d00d4300              push 00430dd0
'006cc9e7    57                      push edi
'006cc9e8    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006cc9e9    ff1580104000            call dword ptr [00401080]
    
End If
'006cc9ef    8b45e4                  mov eax, dword ptr [ebp-1c]
'006cc9f2    50                      push eax
'006cc9f3    68cc134300              push 004313cc

' *** Reference to "__vbaStrCmp"
'006cc9f8    ff1530114000            call dword ptr [00401130]
'006cc9fe    8bf8                    mov edi, eax
'006cca00    f7df                    neg edi
'006cca02    1bff                    sbb edi, edi
'006cca04    f7df                    neg edi
'006cca06    8d4de4                  lea ecx, dword ptr [ebp-1c]
'006cca09    f7df                    neg edi

' *** Reference to "__vbaFreeStr"
'006cca0b    ff150c134000            call dword ptr [0040130c]
'#Cleanup(var_40)
'006cca11    8d4dcc                  lea ecx, dword ptr [ebp-34]

' *** Reference to "__vbaFreeObj"
'006cca14    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_43)
'006cca1a    663bfb                  cmp di, bx
'006cca1d    0f84df040000            je 6ccf02

If (((vbNullString) <> (vbNullChar))) Then
'006cca23    a184b07200              mov ax, word ptr [0072b084]
'006cca28    3bc3                    cmp eax, ebx
'006cca2a    7515                    jne 6cca41
    
    If (    var_28 = 0) Then
'006cca2c    6884b07200              push 0072b084
'006cca31    68f8894100              push 004189f8

' *** Reference to "__vbaNew2"
'006cca36    ff152c124000            call dword ptr [0040122c]
    Set var_28 = New frmGestPerso
'006cca3c    a184b07200              mov ax, word ptr [0072b084]
    
End If
'006cca41    8b08                    mov ecx, dword ptr [eax]
'006cca43    50                      push eax
'006cca44    ff91ac030000            call dword ptr [ecx+000003ac]
'006cca4a    50                      push eax
'006cca4b    8d55cc                  lea edx, dword ptr [ebp-34]
'006cca4e    52                      push edx

' *** Reference to "__vbaObjSet"
'006cca4f    ff15b4104000            call dword ptr [004010b4]
Set var_43 = var_28
'006cca55    8d4de4                  lea ecx, dword ptr [ebp-1c]
'006cca58    8bf8                    mov edi, eax
'006cca5a    8b07                    mov eax, dword ptr [edi]
'006cca5c    51                      push ecx
'006cca5d    57                      push edi
'006cca5e    ff90a0000000            call dword ptr [eax+000000a0]
'006cca64    dbe2                    fnclex
'006cca66    3bc3                    cmp eax, ebx
'006cca68    7d12                    jge 6cca7c

If (frmGestPerso < 0) Then
'006cca6a    68a0000000              push 000000a0
'006cca6f    68d00d4300              push 00430dd0
'006cca74    57                      push edi
'006cca75    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006cca76    ff1580104000            call dword ptr [00401080]
    
End If
'006cca7c    8b55e4                  mov edx, dword ptr [ebp-1c]

' *** Reference to "__vbaStrMove"
'006cca7f    8b3dd0124000            mov edi, dword ptr [004012d0]
'006cca85    8d4de0                  lea ecx, dword ptr [ebp-20]
'006cca88    895de4                  mov dword ptr [ebp-1c], ebx
'006cca8b    ffd7                    call edi
'006cca8d    8d55e0                  lea edx, dword ptr [ebp-20]
'006cca90    52                      push edx
'006cca91    e85a71e2ff              call 4f3bf0
Call sub_4F3BF0()
'006cca96    8bd0                    mov edx, eax
'006cca98    8d4dd0                  lea ecx, dword ptr [ebp-30]
'006cca9b    ffd7                    call edi
'006cca9d    8b55d0                  mov edx, dword ptr [ebp-30]
'006ccaa0    899508ffffff            mov dword ptr [ebp+ffffff08], edx
'006ccaa6    8b1548b07200            mov edx, dword ptr [0072b048]
'006ccaac    c78550ffffff02000000    mov dword ptr [ebp+ffffff50], 00000002
'006ccab6    c78548ffffff03000000    mov dword ptr [ebp+ffffff48], 00000003
'006ccac0    895dd0                  mov dword ptr [ebp-30], ebx
'006ccac3    8b1a                    mov ebx, dword ptr [edx]
'006ccac5    8d55c8                  lea edx, dword ptr [ebp-38]
'006ccac8    52                      push edx
'006ccac9    83ec10                  sub esp, 10
'006ccacc    8bd4                    mov edx, esp
'006ccace    b90a000000              mov ecx, 0000000a
'006ccad3    890a                    mov dword ptr [edx], ecx
'006ccad5    898d38ffffff            mov dword ptr [ebp+ffffff38], ecx
'006ccadb    8b8d2cffffff            mov ecx, dword ptr [ebp+ffffff2c]
'006ccae1    894a04                  mov dword ptr [edx+04], ecx
'006ccae4    b804000280              mov eax, 80020004
'006ccae9    894208                  mov dword ptr [edx+08], eax
'006ccaec    898540ffffff            mov dword ptr [ebp+ffffff40], eax
'006ccaf2    8b8534ffffff            mov eax, dword ptr [ebp+ffffff34]
'006ccaf8    89420c                  mov dword ptr [edx+0c], eax
'006ccafb    8b9538ffffff            mov edx, dword ptr [ebp+ffffff38]
'006ccb01    8b853cffffff            mov eax, dword ptr [ebp+ffffff3c]
'006ccb07    83ec10                  sub esp, 10
'006ccb0a    8bcc                    mov ecx, esp
'006ccb0c    8911                    mov dword ptr [ecx], edx
'006ccb0e    8b9540ffffff            mov edx, dword ptr [ebp+ffffff40]
'006ccb14    894104                  mov dword ptr [ecx+04], eax
'006ccb17    8b8544ffffff            mov eax, dword ptr [ebp+ffffff44]
'006ccb1d    895108                  mov dword ptr [ecx+08], edx
'006ccb20    8b9548ffffff            mov edx, dword ptr [ebp+ffffff48]
'006ccb26    89410c                  mov dword ptr [ecx+0c], eax
'006ccb29    8b854cffffff            mov eax, dword ptr [ebp+ffffff4c]
'006ccb2f    83ec10                  sub esp, 10
'006ccb32    8bcc                    mov ecx, esp
'006ccb34    8911                    mov dword ptr [ecx], edx
'006ccb36    8b9550ffffff            mov edx, dword ptr [ebp+ffffff50]
'006ccb3c    894104                  mov dword ptr [ecx+04], eax
'006ccb3f    8b8554ffffff            mov eax, dword ptr [ebp+ffffff54]
'006ccb45    895108                  mov dword ptr [ecx+08], edx
'006ccb48    8b9508ffffff            mov edx, dword ptr [ebp+ffffff08]
'006ccb4e    89410c                  mov dword ptr [ecx+0c], eax
'006ccb51    6884fa4400              push 0044fa84
'006ccb56    8d4ddc                  lea ecx, dword ptr [ebp-24]
'006ccb59    ffd7                    call edi
'006ccb5b    50                      push eax

' *** Reference to "__vbaStrCat"
'006ccb5c    ff1570104000            call dword ptr [00401070]
var_49 = ("select * from personnage where nom='") & (vbNullString)
'006ccb62    8bd0                    mov edx, eax
'006ccb64    8d4dd8                  lea ecx, dword ptr [ebp-28]
'006ccb67    ffd7                    call edi
'006ccb69    50                      push eax
'006ccb6a    6854a44300              push 0043a454

' *** Reference to "__vbaStrCat"
'006ccb6f    ff1570104000            call dword ptr [00401070]
var_76 = (var_49) & ("'")
'006ccb75    8bd0                    mov edx, eax
'006ccb77    8d4dd4                  lea ecx, dword ptr [ebp-2c]
'006ccb7a    ffd7                    call edi
'006ccb7c    8b0d48b07200            mov ecx, dword ptr [0072b048]
'006ccb82    50                      push eax
'006ccb83    51                      push ecx
'006ccb84    ff93bc000000            call dword ptr [ebx+000000bc]
'006ccb8a    dbe2                    fnclex
'006ccb8c    33db                    xor ebx, ebx
var_num2 = Empty
'006ccb8e    3bc3                    cmp eax, ebx
'006ccb90    7d18                    jge 6ccbaa

If (-4512 < -256 - 12) Then
'006ccb92    8b1548b07200            mov edx, dword ptr [0072b048]
'006ccb98    68bc000000              push 000000bc
'006ccb9d    68301f4300              push 00431f30
'006ccba2    52                      push edx
'006ccba3    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006ccba4    ff1580104000            call dword ptr [00401080]
    
End If
'006ccbaa    8b45c8                  mov eax, dword ptr [ebp-38]
'006ccbad    50                      push eax
'006ccbae    8d45e8                  lea eax, dword ptr [ebp-18]
'006ccbb1    50                      push eax
'006ccbb2    895dc8                  mov dword ptr [ebp-38], ebx

' *** Reference to "__vbaObjSet"
'006ccbb5    ff15b4104000            call dword ptr [004010b4]
Set var_41 = Nothing
'006ccbbb    8d4dd0                  lea ecx, dword ptr [ebp-30]
'006ccbbe    51                      push ecx
'006ccbbf    8d55d4                  lea edx, dword ptr [ebp-2c]
'006ccbc2    52                      push edx
'006ccbc3    8d45d8                  lea eax, dword ptr [ebp-28]
'006ccbc6    50                      push eax
'006ccbc7    8d4ddc                  lea ecx, dword ptr [ebp-24]
'006ccbca    51                      push ecx
'006ccbcb    8d55e0                  lea edx, dword ptr [ebp-20]
'006ccbce    52                      push edx
'006ccbcf    6a05                    push 05

' *** Reference to "__vbaFreeStrList"
'006ccbd1    ff155c124000            call dword ptr [0040125c]
'#Cleanup( 0, 0, -4508, -4512, 0)
'006ccbd7    83c418                  add esp, 18
'006ccbda    8d4dcc                  lea ecx, dword ptr [ebp-34]

' *** Reference to "__vbaFreeObj"
'006ccbdd    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_43)
'006ccbe3    8b45e8                  mov eax, dword ptr [ebp-18]
'006ccbe6    8b08                    mov ecx, dword ptr [eax]
'006ccbe8    8d9524ffffff            lea edx, dword ptr [ebp+ffffff24]
'006ccbee    52                      push edx
'006ccbef    50                      push eax
'006ccbf0    ff5134                  call dword ptr [ecx+34]
'006ccbf3    dbe2                    fnclex
'006ccbf5    3bc3                    cmp eax, ebx
'006ccbf7    7d12                    jge 6ccc0b

If (var_41 < -256 - 12) Then
'006ccbf9    8b4de8                  mov ecx, dword ptr [ebp-18]
'006ccbfc    6a34                    push 34
'006ccbfe    6830314300              push 00433130
'006ccc03    51                      push ecx
'006ccc04    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006ccc05    ff1580104000            call dword ptr [00401080]
    
End If
'006ccc0b    66399d24ffffff          cmp word ptr [ebp+ffffff24], bx
'006ccc12    0f85c3020000            jne 6ccedb

If (0 = -256 - 12) Then
'006ccc18    8b45e8                  mov eax, dword ptr [ebp-18]
'006ccc1b    8b10                    mov edx, dword ptr [eax]
'006ccc1d    50                      push eax
'006ccc1e    ff92d0000000            call dword ptr [edx+000000d0]
'006ccc24    dbe2                    fnclex
'006ccc26    3bc3                    cmp eax, ebx
'006ccc28    7d15                    jge 6ccc3f
    
    If (    var_41 < -256 - 12) Then
'006ccc2a    8b4de8                  mov ecx, dword ptr [ebp-18]
'006ccc2d    68d0000000              push 000000d0
'006ccc32    6830314300              push 00433130
'006ccc37    51                      push ecx
'006ccc38    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006ccc39    ff1580104000            call dword ptr [00401080]
    
End If
'006ccc3f    8b16                    mov edx, dword ptr [esi]
'006ccc41    56                      push esi
'006ccc42    ff9208030000            call dword ptr [edx+00000308]
'006ccc48    8945c0                  mov dword ptr [ebp-40], eax
'006ccc4b    8d45b8                  lea eax, dword ptr [ebp-48]
'006ccc4e    50                      push eax
'006ccc4f    8d4da8                  lea ecx, dword ptr [ebp-58]
'006ccc52    bf09000000              mov edi, 00000009
'006ccc57    51                      push ecx
'006ccc58    897db8                  mov dword ptr [ebp-48], edi

' *** Reference to "rtcTrimVar"
'006ccc5b    ff15e4104000            call dword ptr [004010e4]
'006ccc61    8b16                    mov edx, dword ptr [esi]
'006ccc63    56                      push esi
'006ccc64    ff9208030000            call dword ptr [edx+00000308]
'006ccc6a    898570ffffff            mov dword ptr [ebp+ffffff70], eax
'006ccc70    8d45a8                  lea eax, dword ptr [ebp-58]
'006ccc73    50                      push eax
'006ccc74    8d8d48ffffff            lea ecx, dword ptr [ebp+ffffff48]
'006ccc7a    51                      push ecx
'006ccc7b    8d5598                  lea edx, dword ptr [ebp-68]
'006ccc7e    52                      push edx
'006ccc7f    89bd68ffffff            mov dword ptr [ebp+ffffff68], edi
'006ccc85    c78578ffffff01000000    mov dword ptr [ebp+ffffff78], 00000001
'006ccc8f    c78550ffffffcc134300    mov dword ptr [ebp+ffffff50], 004313cc
'006ccc99    c78548ffffff08800000    mov dword ptr [ebp+ffffff48], 00008008

' *** Reference to "__vbaVarCmpEq"
'006ccca3    ff1580124000            call dword ptr [00401280]
'006ccca9    8bd0                    mov edx, eax
'006cccab    8d4d88                  lea ecx, dword ptr [ebp-78]

' *** Reference to "__vbaVarMove"
'006cccae    ff151c104000            call dword ptr [0040101c]
var_131 = (#NOT SUPPORTED#)
'006cccb4    8d8568ffffff            lea eax, dword ptr [ebp+ffffff68]
'006cccba    50                      push eax
'006cccbb    8d8d78ffffff            lea ecx, dword ptr [ebp+ffffff78]
'006cccc1    51                      push ecx
'006cccc2    8d5588                  lea edx, dword ptr [ebp-78]
'006cccc5    52                      push edx
'006cccc6    8d8558ffffff            lea eax, dword ptr [ebp+ffffff58]
'006ccccc    50                      push eax

' *** Reference to "rtcImmediateIf"
'006ccccd    ff154c124000            call dword ptr [0040124c]
'006cccd3    8b9d58ffffff            mov ebx, dword ptr [ebp+ffffff58]
'006cccd9    8b55e8                  mov edx, dword ptr [ebp-18]
'006cccdc    8b12                    mov edx, dword ptr [edx]
'006cccde    83ec10                  sub esp, 10
'006ccce1    8bfc                    mov edi, esp
'006ccce3    891f                    mov dword ptr [edi], ebx
'006ccce5    8b9d5cffffff            mov ebx, dword ptr [ebp+ffffff5c]
'006ccceb    895f04                  mov dword ptr [edi+04], ebx
'006cccee    8b9d60ffffff            mov ebx, dword ptr [ebp+ffffff60]
'006cccf4    895f08                  mov dword ptr [edi+08], ebx
'006cccf7    8b9d64ffffff            mov ebx, dword ptr [ebp+ffffff64]
'006cccfd    895f0c                  mov dword ptr [edi+0c], ebx
'006ccd00    83ec10                  sub esp, 10
'006ccd03    8bfc                    mov edi, esp
'006ccd05    b908000000              mov ecx, 00000008
'006ccd0a    890f                    mov dword ptr [edi], ecx
'006ccd0c    8b8d2cffffff            mov ecx, dword ptr [ebp+ffffff2c]
'006ccd12    894f04                  mov dword ptr [edi+04], ecx
'006ccd15    8b4de8                  mov ecx, dword ptr [ebp-18]
'006ccd18    b8ac174400              mov eax, 004417ac
'006ccd1d    894708                  mov dword ptr [edi+08], eax
'006ccd20    8b8534ffffff            mov eax, dword ptr [ebp+ffffff34]
'006ccd26    51                      push ecx
'006ccd27    89470c                  mov dword ptr [edi+0c], eax
'006ccd2a    ff9228010000            call dword ptr [edx+00000128]
'006ccd30    dbe2                    fnclex
'006ccd32    85c0                    test ax, ax
'006ccd34    7d15                    jge 6ccd4b
'006ccd36    8b55e8                  mov edx, dword ptr [ebp-18]
'006ccd39    6828010000              push 00000128
'006ccd3e    6830314300              push 00433130
'006ccd43    52                      push edx
'006ccd44    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006ccd45    ff1580104000            call dword ptr [00401080]
'006ccd4b    8d8558ffffff            lea eax, dword ptr [ebp+ffffff58]
'006ccd51    50                      push eax
'006ccd52    8d8d68ffffff            lea ecx, dword ptr [ebp+ffffff68]
'006ccd58    51                      push ecx
'006ccd59    8d9578ffffff            lea edx, dword ptr [ebp+ffffff78]
'006ccd5f    52                      push edx
'006ccd60    8d4588                  lea eax, dword ptr [ebp-78]
'006ccd63    50                      push eax
'006ccd64    8d4da8                  lea ecx, dword ptr [ebp-58]
'006ccd67    51                      push ecx
'006ccd68    8d55b8                  lea edx, dword ptr [ebp-48]
'006ccd6b    52                      push edx
'006ccd6c    6a06                    push 06

' *** Reference to "__vbaFreeVarList"
'006ccd6e    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_61, var_86, var_131, var_87, var_132, var_133)
'006ccd74    8b06                    mov eax, dword ptr [esi]
'006ccd76    83c41c                  add esp, 1c
'006ccd79    56                      push esi
'006ccd7a    ff9004030000            call dword ptr [eax+00000304]
'006ccd80    8d4db8                  lea ecx, dword ptr [ebp-48]
'006ccd83    51                      push ecx
'006ccd84    8d55a8                  lea edx, dword ptr [ebp-58]
'006ccd87    bf09000000              mov edi, 00000009
'006ccd8c    52                      push edx
'006ccd8d    8945c0                  mov dword ptr [ebp-40], eax
'006ccd90    897db8                  mov dword ptr [ebp-48], edi

' *** Reference to "rtcTrimVar"
'006ccd93    ff15e4104000            call dword ptr [004010e4]
'006ccd99    8b06                    mov eax, dword ptr [esi]
'006ccd9b    56                      push esi
'006ccd9c    ff9004030000            call dword ptr [eax+00000304]
'006ccda2    8d4da8                  lea ecx, dword ptr [ebp-58]
'006ccda5    51                      push ecx
'006ccda6    8d9548ffffff            lea edx, dword ptr [ebp+ffffff48]
'006ccdac    898570ffffff            mov dword ptr [ebp+ffffff70], eax
'006ccdb2    52                      push edx
'006ccdb3    8d4598                  lea eax, dword ptr [ebp-68]
'006ccdb6    50                      push eax
'006ccdb7    89bd68ffffff            mov dword ptr [ebp+ffffff68], edi
'006ccdbd    c78578ffffff01000000    mov dword ptr [ebp+ffffff78], 00000001
'006ccdc7    c78550ffffffcc134300    mov dword ptr [ebp+ffffff50], 004313cc
'006ccdd1    c78548ffffff08800000    mov dword ptr [ebp+ffffff48], 00008008

' *** Reference to "__vbaVarCmpEq"
'006ccddb    ff1580124000            call dword ptr [00401280]
'006ccde1    8bd0                    mov edx, eax
'006ccde3    8d4d88                  lea ecx, dword ptr [ebp-78]

' *** Reference to "__vbaVarMove"
'006ccde6    ff151c104000            call dword ptr [0040101c]
var_131 = (#NOT SUPPORTED#)
'006ccdec    8d8d68ffffff            lea ecx, dword ptr [ebp+ffffff68]
'006ccdf2    51                      push ecx
'006ccdf3    8d9578ffffff            lea edx, dword ptr [ebp+ffffff78]
'006ccdf9    52                      push edx
'006ccdfa    8d4588                  lea eax, dword ptr [ebp-78]
'006ccdfd    50                      push eax
'006ccdfe    8d8d58ffffff            lea ecx, dword ptr [ebp+ffffff58]
'006cce04    51                      push ecx

' *** Reference to "rtcImmediateIf"
'006cce05    ff154c124000            call dword ptr [0040124c]
'006cce0b    8b9d58ffffff            mov ebx, dword ptr [ebp+ffffff58]
'006cce11    8b55e8                  mov edx, dword ptr [ebp-18]
'006cce14    8b12                    mov edx, dword ptr [edx]
'006cce16    83ec10                  sub esp, 10
'006cce19    8bfc                    mov edi, esp
'006cce1b    891f                    mov dword ptr [edi], ebx
'006cce1d    8b9d5cffffff            mov ebx, dword ptr [ebp+ffffff5c]
'006cce23    895f04                  mov dword ptr [edi+04], ebx
'006cce26    8b9d60ffffff            mov ebx, dword ptr [ebp+ffffff60]
'006cce2c    895f08                  mov dword ptr [edi+08], ebx
'006cce2f    8b9d64ffffff            mov ebx, dword ptr [ebp+ffffff64]
'006cce35    895f0c                  mov dword ptr [edi+0c], ebx
'006cce38    83ec10                  sub esp, 10
'006cce3b    8bfc                    mov edi, esp
'006cce3d    b908000000              mov ecx, 00000008
'006cce42    890f                    mov dword ptr [edi], ecx
'006cce44    8b8d2cffffff            mov ecx, dword ptr [ebp+ffffff2c]
'006cce4a    894f04                  mov dword ptr [edi+04], ecx
'006cce4d    8b4de8                  mov ecx, dword ptr [ebp-18]
'006cce50    b8c8174400              mov eax, 004417c8
'006cce55    894708                  mov dword ptr [edi+08], eax
'006cce58    8b8534ffffff            mov eax, dword ptr [ebp+ffffff34]
'006cce5e    51                      push ecx
'006cce5f    89470c                  mov dword ptr [edi+0c], eax
'006cce62    ff9228010000            call dword ptr [edx+00000128]
'006cce68    dbe2                    fnclex
'006cce6a    33db                    xor ebx, ebx
var_num2 = Empty
'006cce6c    3bc3                    cmp eax, ebx
'006cce6e    7d15                    jge 6cce85

If (0 < 0) Then
'006cce70    8b55e8                  mov edx, dword ptr [ebp-18]
'006cce73    6828010000              push 00000128
'006cce78    6830314300              push 00433130
'006cce7d    52                      push edx
'006cce7e    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006cce7f    ff1580104000            call dword ptr [00401080]
    
End If
'006cce85    8d8558ffffff            lea eax, dword ptr [ebp+ffffff58]
'006cce8b    50                      push eax
'006cce8c    8d8d68ffffff            lea ecx, dword ptr [ebp+ffffff68]
'006cce92    51                      push ecx
'006cce93    8d9578ffffff            lea edx, dword ptr [ebp+ffffff78]
'006cce99    52                      push edx
'006cce9a    8d4588                  lea eax, dword ptr [ebp-78]
'006cce9d    50                      push eax
'006cce9e    8d4da8                  lea ecx, dword ptr [ebp-58]
'006ccea1    51                      push ecx
'006ccea2    8d55b8                  lea edx, dword ptr [ebp-48]
'006ccea5    52                      push edx
'006ccea6    6a06                    push 06

' *** Reference to "__vbaFreeVarList"
'006ccea8    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_61, var_86, var_131, var_87, var_132, var_133)
'006cceae    8b45e8                  mov eax, dword ptr [ebp-18]
'006cceb1    8b08                    mov ecx, dword ptr [eax]
'006cceb3    83c41c                  add esp, 1c
'006cceb6    53                      push ebx
'006cceb7    6a01                    push 01
'006cceb9    50                      push eax
'006cceba    ff9164010000            call dword ptr [ecx+00000164]
'006ccec0    dbe2                    fnclex
'006ccec2    3bc3                    cmp eax, ebx
'006ccec4    7d15                    jge 6ccedb

If (var_41 < 0) Then
'006ccec6    8b55e8                  mov edx, dword ptr [ebp-18]
'006ccec9    6864010000              push 00000164
'006ccece    6830314300              push 00433130
'006cced3    52                      push edx
'006cced4    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006cced5    ff1580104000            call dword ptr [00401080]
    
End If
'006ccedb    8b45e8                  mov eax, dword ptr [ebp-18]
'006ccede    8b08                    mov ecx, dword ptr [eax]
'006ccee0    50                      push eax
'006ccee1    ff91c4000000            call dword ptr [ecx+000000c4]
'006ccee7    dbe2                    fnclex
'006ccee9    3bc3                    cmp eax, ebx
'006cceeb    7d15                    jge 6ccf02

If (var_41 < -256 - 12) Then
'006cceed    8b55e8                  mov edx, dword ptr [ebp-18]
'006ccef0    68c4000000              push 000000c4
'006ccef5    6830314300              push 00433130
'006ccefa    52                      push edx
'006ccefb    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006ccefc    ff1580104000            call dword ptr [00401080]
    
End If
'006ccf02    391d24be7200            cmp dword ptr [0072be24], ebx
'006ccf08    7510                    jne 6ccf1a
'006ccf0a    6824be7200              push 0072be24
'006ccf0f    6870004300              push 00430070

' *** Reference to "__vbaNew2"
'006ccf14    ff152c124000            call dword ptr [0040122c]
Dim var_2 As New Global
'006ccf1a    8b3d24be7200            mov edi, dword ptr [0072be24]
'006ccf20    8b17                    mov edx, dword ptr [edi]
'006ccf22    56                      push esi
'006ccf23    8d45cc                  lea eax, dword ptr [ebp-34]
'006ccf26    50                      push eax
'006ccf27    899500ffffff            mov dword ptr [ebp+ffffff00], edx

' *** Reference to "__vbaObjSetAddref"
'006ccf2d    ff15c8104000            call dword ptr [004010c8]
Set var_43 = Me
'006ccf33    8b8d00ffffff            mov ecx, dword ptr [ebp+ffffff00]
'006ccf39    50                      push eax
'006ccf3a    57                      push edi
'006ccf3b    ff5110                  call dword ptr [ecx+10]
Call var_2.Unload(var_43)
'006ccf3e    dbe2                    fnclex
'006ccf40    3bc3                    cmp eax, ebx
'006ccf42    7d0f                    jge 6ccf53
'006ccf44    6a10                    push 10
'006ccf46    6860004300              push 00430060
'006ccf4b    57                      push edi
'006ccf4c    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006ccf4d    ff1580104000            call dword ptr [00401080]
'006ccf53    8d4dcc                  lea ecx, dword ptr [ebp-34]

' *** Reference to "__vbaFreeObj"
'006ccf56    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_43)
'006ccf5c    895dfc                  mov dword ptr [ebp-04], ebx
'006ccf5f    68d1cf6c00              push 006ccfd1
'006ccf64    eb61                    jmp 6ccfc7
'006ccf66    8d55d0                  lea edx, dword ptr [ebp-30]
'006ccf69    52                      push edx
'006ccf6a    8d45d4                  lea eax, dword ptr [ebp-2c]
'006ccf6d    50                      push eax
'006ccf6e    8d4dd8                  lea ecx, dword ptr [ebp-28]
'006ccf71    51                      push ecx
'006ccf72    8d55dc                  lea edx, dword ptr [ebp-24]
'006ccf75    52                      push edx
'006ccf76    8d45e0                  lea eax, dword ptr [ebp-20]
'006ccf79    50                      push eax
'006ccf7a    8d4de4                  lea ecx, dword ptr [ebp-1c]
'006ccf7d    51                      push ecx
'006ccf7e    6a06                    push 06

' *** Reference to "__vbaFreeStrList"
'006ccf80    ff155c124000            call dword ptr [0040125c]
'#Cleanup( 0, 0, 0, -4508, -4512, 0)
'006ccf86    8d55c8                  lea edx, dword ptr [ebp-38]
'006ccf89    52                      push edx
'006ccf8a    8d45cc                  lea eax, dword ptr [ebp-34]
'006ccf8d    50                      push eax
'006ccf8e    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'006ccf90    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_43, var_46)
'006ccf96    8d8d58ffffff            lea ecx, dword ptr [ebp+ffffff58]
'006ccf9c    51                      push ecx
'006ccf9d    8d9568ffffff            lea edx, dword ptr [ebp+ffffff68]
'006ccfa3    52                      push edx
'006ccfa4    8d8578ffffff            lea eax, dword ptr [ebp+ffffff78]
'006ccfaa    50                      push eax
'006ccfab    8d4d88                  lea ecx, dword ptr [ebp-78]
'006ccfae    51                      push ecx
'006ccfaf    8d5598                  lea edx, dword ptr [ebp-68]
'006ccfb2    52                      push edx
'006ccfb3    8d45a8                  lea eax, dword ptr [ebp-58]
'006ccfb6    50                      push eax
'006ccfb7    8d4db8                  lea ecx, dword ptr [ebp-48]
'006ccfba    51                      push ecx
'006ccfbb    6a07                    push 07

' *** Reference to "__vbaFreeVarList"
'006ccfbd    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_61, var_86, var_130, var_131, var_87, var_132, var_133)
'006ccfc3    83c448                  add esp, 48
'006ccfc6    c3                      ret
'006ccfc7    8d4de8                  lea ecx, dword ptr [ebp-18]

' *** Reference to "__vbaFreeObj"
'006ccfca    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_41)
'006ccfd0    c3                      ret
'006ccfd1    8b4508                  mov eax, dword ptr [ebp+08]
'006ccfd4    8b10                    mov edx, dword ptr [eax]
'006ccfd6    50                      push eax
'006ccfd7    ff5208                  call dword ptr [edx+08]
'006ccfda    8b45fc                  mov eax, dword ptr [ebp-04]
'006ccfdd    8b4dec                  mov ecx, dword ptr [ebp-14]
'006ccfe0    5f                      pop edi
'006ccfe1    5e                      pop esi
    'Reference to '__except_list'
'006ccfe2    64890d00000000          mov dword ptr fs:[00000000], ecx
'006ccfe9    5b                      pop ebx
'006ccfea    8be5                    mov esp, ebp
'006ccfec    5d                      pop ebp
'006ccfed    c20400                  ret 0004


End Sub


'Event for BtnFermer
Private Sub BtnFermer_Click()
'006ccff0    55                      push ebp
'006ccff1    8bec                    mov ebp, esp
'006ccff3    83ec0c                  sub esp, 0c

' *** Reference to "__vbaExceptHandler"
'006ccff6    6866724000              push 00407266
'006ccffb    64a100000000            mov ax, word ptr fs:[00000000]
'006cd001    50                      push eax
    'Reference to '__except_list'
'006cd002    64892500000000          mov dword ptr fs:[00000000], esp
'006cd009    83ec18                  sub esp, 18
'006cd00c    53                      push ebx
'006cd00d    56                      push esi
'006cd00e    57                      push edi
'006cd00f    8965f4                  mov dword ptr [ebp-0c], esp
'006cd012    c745f8e8674000          mov dword ptr [ebp-08], 004067e8
'006cd019    8b7d08                  mov edi, dword ptr [ebp+08]
'006cd01c    8bc7                    mov eax, edi
'006cd01e    83e001                  and eax, 01
'006cd021    8945fc                  mov dword ptr [ebp-04], eax
'006cd024    83e7fe                  and edi, fffffffe
'006cd027    8b0f                    mov ecx, dword ptr [edi]
'006cd029    57                      push edi
'006cd02a    897d08                  mov dword ptr [ebp+08], edi
'006cd02d    ff5104                  call dword ptr [ecx+04]
'006cd030    a124be7200              mov ax, word ptr [0072be24]
'006cd035    33db                    xor ebx, ebx
'006cd037    3bc3                    cmp eax, ebx
'006cd039    895de8                  mov dword ptr [ebp-18], ebx
'006cd03c    7510                    jne 6cd04e

If (0 = 0) Then
'006cd03e    6824be7200              push 0072be24
'006cd043    6870004300              push 00430070

' *** Reference to "__vbaNew2"
'006cd048    ff152c124000            call dword ptr [0040122c]
    Dim var_2 As New Global
End If
'006cd04e    8b3524be7200            mov esi, dword ptr [0072be24]
'006cd054    8b16                    mov edx, dword ptr [esi]
'006cd056    57                      push edi
'006cd057    8d45e8                  lea eax, dword ptr [ebp-18]
'006cd05a    50                      push eax
'006cd05b    8955d4                  mov dword ptr [ebp-2c], edx

' *** Reference to "__vbaObjSetAddref"
'006cd05e    ff15c8104000            call dword ptr [004010c8]
Set var_41 = Me
'006cd064    8b4dd4                  mov ecx, dword ptr [ebp-2c]
'006cd067    50                      push eax
'006cd068    56                      push esi
'006cd069    ff5110                  call dword ptr [ecx+10]
Call var_2.Unload(var_41)
'006cd06c    dbe2                    fnclex
'006cd06e    3bc3                    cmp eax, ebx
'006cd070    7d0f                    jge 6cd081
'006cd072    6a10                    push 10
'006cd074    6860004300              push 00430060
'006cd079    56                      push esi
'006cd07a    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006cd07b    ff1580104000            call dword ptr [00401080]
'006cd081    8d4de8                  lea ecx, dword ptr [ebp-18]

' *** Reference to "__vbaFreeObj"
'006cd084    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_41)
'006cd08a    895dfc                  mov dword ptr [ebp-04], ebx
'006cd08d    689fd06c00              push 006cd09f
'006cd092    eb0a                    jmp 6cd09e
'006cd094    8d4de8                  lea ecx, dword ptr [ebp-18]

' *** Reference to "__vbaFreeObj"
'006cd097    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_41)
'006cd09d    c3                      ret
'006cd09e    c3                      ret
'006cd09f    8b4508                  mov eax, dword ptr [ebp+08]
'006cd0a2    8b10                    mov edx, dword ptr [eax]
'006cd0a4    50                      push eax
'006cd0a5    ff5208                  call dword ptr [edx+08]
'006cd0a8    8b45fc                  mov eax, dword ptr [ebp-04]
'006cd0ab    8b4dec                  mov ecx, dword ptr [ebp-14]
'006cd0ae    5f                      pop edi
'006cd0af    5e                      pop esi
    'Reference to '__except_list'
'006cd0b0    64890d00000000          mov dword ptr fs:[00000000], ecx
'006cd0b7    5b                      pop ebx
'006cd0b8    8be5                    mov esp, ebp
'006cd0ba    5d                      pop ebp
'006cd0bb    c20400                  ret 0004


End Sub


