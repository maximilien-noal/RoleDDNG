VERSION 5.00

Begin VB.Form frmVendre
    Caption = "Vendre un objet"
    ScaleMode = 1
    AutoRedraw = 0              'False
    FontTransparent = -1              'True
    BorderStyle = 1
    LinkTopic = "Form1"
    MaxButton = 0              'False
    MinButton = 0              'False
    ClientLeft   = 45
    ClientTop    = 435
    ClientWidth  = 4860
    ClientHeight = 3090
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
    Begin VB.CommandButton btnSupprimer
        Caption = "Supprimer"
        Left   = 240
        Top    = 2520
        Width  = 1095
        Height = 495
        TabIndex = 8
    End
    Begin VB.Frame frmObjet
        Caption = "NonObjet"
        Left   = 120
        Top    = 0
        Width  = 4695
        Height = 2415
        TabIndex = 2
        BeginProperty Font
            Name          = "Times New Roman"
            Size          = 11,25
            Charset       = 0
            Weight        = 400
            Underline     = 0              'False
            Italic        = 0              'False
            Strikethrough = 0              'False
        EndProperty
        Begin VB.TextBox fldnPrixTotal
            Left   = 2760
            Top    = 1840
            Width  = 1000
            Height = 375
            Text = "1"
            TabIndex = 10
        End
        Begin VB.TextBox fldnNombreObjet
            Left   = 2760
            Top    = 1340
            Width  = 1000
            Height = 375
            Text = "1"
            TabIndex = 7
        End
        Begin VB.TextBox fldnPrixUnitaire
            Left   = 2760
            Top    = 840
            Width  = 1000
            Height = 375
            Text = "1"
            TabIndex = 5
        End
        Begin VB.Label LblObjet
            Caption = "Prix de revente des objets"
            Index = 3
            Left   = 195
            Top    = 1900
            Width  = 2280
            Height = 255
            TabIndex = 9
            AutoSize = -1              'True
            BeginProperty Font
                Name          = "Times New Roman"
                Size          = 11,25
                Charset       = 0
                Weight        = 400
                Underline     = 0              'False
                Italic        = 0              'False
                Strikethrough = 0              'False
            EndProperty
        End
        Begin VB.Label LblObjet
            Caption = "Nombre d'objet à vendre"
            Index = 2
            Left   = 200
            Top    = 1400
            Width  = 2175
            Height = 255
            TabIndex = 6
            AutoSize = -1              'True
            BeginProperty Font
                Name          = "Times New Roman"
                Size          = 11,25
                Charset       = 0
                Weight        = 400
                Underline     = 0              'False
                Italic        = 0              'False
                Strikethrough = 0              'False
            EndProperty
        End
        Begin VB.Label LblObjet
            Caption = "Prix de revente d'un objet"
            Index = 1
            Left   = 200
            Top    = 900
            Width  = 2250
            Height = 255
            TabIndex = 4
            AutoSize = -1              'True
            BeginProperty Font
                Name          = "Times New Roman"
                Size          = 11,25
                Charset       = 0
                Weight        = 400
                Underline     = 0              'False
                Italic        = 0              'False
                Strikethrough = 0              'False
            EndProperty
        End
        Begin VB.Label LblObjet
            Caption = "Prix neuf de l'objet : "
            Index = 0
            Left   = 200
            Top    = 400
            Width  = 1815
            Height = 255
            TabIndex = 3
            AutoSize = -1              'True
            BeginProperty Font
                Name          = "Times New Roman"
                Size          = 11,25
                Charset       = 0
                Weight        = 400
                Underline     = 0              'False
                Italic        = 0              'False
                Strikethrough = 0              'False
            EndProperty
        End
    End
    Begin VB.CommandButton btnAnnuler
        Caption = "Annuler"
        Left   = 3720
        Top    = 2400
        Width  = 1095
        Height = 615
        TabIndex = 1
    End
    Begin VB.CommandButton btnVendre
        Caption = "Vendre"
        Left   = 2400
        Top    = 2520
        Width  = 1095
        Height = 495
        TabIndex = 0
    End
End
Private Function sub_727760(arg_0 As Unknow, arg_1 As Unknow, arg_2 As Unknow, arg_3 As Unknow, arg_4 As Unknow, arg_5 As Unknow, arg_6 As Unknow, arg_7 As Unknow, arg_8 As Unknow, arg_9 As Unknow, arg_A As Unknow, arg_B As Unknow, arg_C As Unknow, arg_D As Unknow, arg_E As Unknow, arg_F As Unknow, arg_10 As Unknow, arg_11 As Unknow, arg_12 As Unknow, arg_13 As Unknow, arg_14 As Unknow, arg_15 As Unknow, arg_16 As Unknow, arg_17 As Unknow, arg_18 As Unknow, arg_19 As Unknow, arg_1A As Unknow, arg_1B As Unknow, arg_1C As Unknow, arg_1D As Unknow, arg_1E As Unknow, arg_1F As Unknow, arg_20 As Unknow, arg_21 As Unknow, arg_22 As Unknow, arg_23 As Unknow, arg_24 As Unknow, arg_25 As Unknow, arg_26 As Unknow, arg_27 As Unknow, arg_28 As Unknow, arg_29 As Unknow, arg_2A As Unknow, arg_2B As Unknow, arg_2C As Unknow, arg_2D As Unknow, arg_2E As Unknow, arg_2F As Unknow, arg_30 As Unknow, arg_31 As Unknow, arg_32 As Unknow, arg_33 As Unknow, arg_34 As Unknow, arg_35 As Unknow, arg_36 As Unknow, arg_37 As Unknow, arg_38 As Unknow, arg_39 As Unknow, arg_3A As Unknow, arg_3B As Unknow, arg_3C As Unknow)
'00727760    55                      push ebp
'00727761    8bec                    mov ebp, esp
'00727763    83ec08                  sub esp, 08

' *** Reference to "__vbaExceptHandler"
'00727766    6866724000              push 00407266
'0072776b    64a100000000            mov ax, word ptr fs:[00000000]
'00727771    50                      push eax
    'Reference to '__except_list'
'00727772    64892500000000          mov dword ptr fs:[00000000], esp
'00727779    83ec44                  sub esp, 44
'0072777c    53                      push ebx
'0072777d    56                      push esi
'0072777e    57                      push edi
'0072777f    8965f8                  mov dword ptr [ebp-08], esp
'00727782    c745fc08724000          mov dword ptr [ebp-04], 00407208
'00727789    8b7508                  mov esi, dword ptr [ebp+08]
'0072778c    8b06                    mov eax, dword ptr [esi]
'0072778e    33db                    xor ebx, ebx
'00727790    56                      push esi
'00727791    895de8                  mov dword ptr [ebp-18], ebx
'00727794    895de4                  mov dword ptr [ebp-1c], ebx
'00727797    895de0                  mov dword ptr [ebp-20], ebx
'0072779a    895ddc                  mov dword ptr [ebp-24], ebx
'0072779d    895dd8                  mov dword ptr [ebp-28], ebx
'007277a0    895dd4                  mov dword ptr [ebp-2c], ebx
'007277a3    ff9008030000            call dword ptr [eax+00000308]
'007277a9    50                      push eax
'007277aa    8d4ddc                  lea ecx, dword ptr [ebp-24]
'007277ad    51                      push ecx

' *** Reference to "__vbaObjSet"
'007277ae    ff15b4104000            call dword ptr [004010b4]
Set var_39 = Nothing
'007277b4    8bf8                    mov edi, eax
'007277b6    8b17                    mov edx, dword ptr [edi]
'007277b8    8d45e8                  lea eax, dword ptr [ebp-18]
'007277bb    50                      push eax
'007277bc    57                      push edi
'007277bd    ff92a0000000            call dword ptr [edx+000000a0]
'007277c3    dbe2                    fnclex
'007277c5    3bc3                    cmp eax, ebx
'007277c7    7d12                    jge 7277db
'007277c9    68a0000000              push 000000a0
'007277ce    68d00d4300              push 00430dd0
'007277d3    57                      push edi
'007277d4    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'007277d5    ff1580104000            call dword ptr [00401080]
'007277db    8b4de8                  mov ecx, dword ptr [ebp-18]
'007277de    51                      push ecx

' *** Reference to "rtcR8ValFromBstr"
'007277df    ff1510134000            call dword ptr [00401310]

' *** Reference to "__vbaFpI2"
'007277e5    ff15a0124000            call dword ptr [004012a0]
var_num1 = CLng(Val(vbNullString))
'007277eb    8d4de8                  lea ecx, dword ptr [ebp-18]
'007277ee    8bd8                    mov ebx, eax

' *** Reference to "__vbaFreeStr"
'007277f0    ff150c134000            call dword ptr [0040130c]
'#Cleanup(var_41)
'007277f6    8d4ddc                  lea ecx, dword ptr [ebp-24]

' *** Reference to "__vbaFreeObj"
'007277f9    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_39)
'007277ff    8b16                    mov edx, dword ptr [esi]
'00727801    56                      push esi
'00727802    ff9208030000            call dword ptr [edx+00000308]
'00727808    50                      push eax
'00727809    8d45dc                  lea eax, dword ptr [ebp-24]
'0072780c    50                      push eax

' *** Reference to "__vbaObjSet"
'0072780d    ff15b4104000            call dword ptr [004010b4]
Set var_39 = var_num1
'00727813    8bf8                    mov edi, eax
'00727815    8b17                    mov edx, dword ptr [edi]
'00727817    53                      push ebx
'00727818    8955b4                  mov dword ptr [ebp-4c], edx

' *** Reference to "__vbaStrI2"
'0072781b    ff1510104000            call dword ptr [00401010]
'00727821    8bd0                    mov edx, eax
'00727823    8d4de8                  lea ecx, dword ptr [ebp-18]

' *** Reference to "__vbaStrMove"
'00727826    ff15d0124000            call dword ptr [004012d0]
'0072782c    8b4db4                  mov ecx, dword ptr [ebp-4c]
'0072782f    50                      push eax
'00727830    57                      push edi
'00727831    ff91a4000000            call dword ptr [ecx+000000a4]
'00727837    dbe2                    fnclex
'00727839    85c0                    test ax, ax
'0072783b    7d12                    jge 72784f
'0072783d    68a4000000              push 000000a4
'00727842    68d00d4300              push 00430dd0
'00727847    57                      push edi
'00727848    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00727849    ff1580104000            call dword ptr [00401080]
'0072784f    8d4de8                  lea ecx, dword ptr [ebp-18]

' *** Reference to "__vbaFreeStr"
'00727852    ff150c134000            call dword ptr [0040130c]
'#Cleanup(var_41)
'00727858    8d4ddc                  lea ecx, dword ptr [ebp-24]

' *** Reference to "__vbaFreeObj"
'0072785b    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_39)
'00727861    8b16                    mov edx, dword ptr [esi]
'00727863    56                      push esi
'00727864    ff9208030000            call dword ptr [edx+00000308]

' *** Reference to "__vbaObjSet"
'0072786a    8b1db4104000            mov ebx, dword ptr [004010b4]
'00727870    50                      push eax
'00727871    8d45dc                  lea eax, dword ptr [ebp-24]
'00727874    50                      push eax
'00727875    ffd3                    call ebx
Set var_39 = -4500
'00727877    8d55e8                  lea edx, dword ptr [ebp-18]
'0072787a    8bf8                    mov edi, eax
'0072787c    8b0f                    mov ecx, dword ptr [edi]
'0072787e    52                      push edx
'0072787f    57                      push edi
'00727880    ff91a0000000            call dword ptr [ecx+000000a0]
'00727886    dbe2                    fnclex
'00727888    85c0                    test ax, ax
'0072788a    7d12                    jge 72789e
'0072788c    68a0000000              push 000000a0
'00727891    68d00d4300              push 00430dd0
'00727896    57                      push edi
'00727897    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00727898    ff1580104000            call dword ptr [00401080]
'0072789e    8b06                    mov eax, dword ptr [esi]
'007278a0    56                      push esi
'007278a1    ff9004030000            call dword ptr [eax+00000304]
'007278a7    50                      push eax
'007278a8    8d4dd4                  lea ecx, dword ptr [ebp-2c]
'007278ab    51                      push ecx
'007278ac    ffd3                    call ebx
Set var_44 = Nothing
'007278ae    8b16                    mov edx, dword ptr [esi]
'007278b0    56                      push esi
'007278b1    8bf8                    mov edi, eax
'007278b3    ff920c030000            call dword ptr [edx+0000030c]
'007278b9    50                      push eax
'007278ba    8d45d8                  lea eax, dword ptr [ebp-28]
'007278bd    50                      push eax
'007278be    ffd3                    call ebx
Set var_45 = Nothing
'007278c0    8d55e4                  lea edx, dword ptr [ebp-1c]
'007278c3    8bf0                    mov esi, eax
'007278c5    8b0e                    mov ecx, dword ptr [esi]
'007278c7    52                      push edx
'007278c8    56                      push esi
'007278c9    ff91a0000000            call dword ptr [ecx+000000a0]
'007278cf    dbe2                    fnclex
'007278d1    85c0                    test ax, ax
'007278d3    7d12                    jge 7278e7
'007278d5    68a0000000              push 000000a0
'007278da    68d00d4300              push 00430dd0
'007278df    56                      push esi
'007278e0    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'007278e1    ff1580104000            call dword ptr [00401080]
'007278e7    8b45e4                  mov eax, dword ptr [ebp-1c]
'007278ea    8b37                    mov esi, dword ptr [edi]
'007278ec    50                      push eax

' *** Reference to "__vbaR8Str"
'007278ed    ff1524124000            call dword ptr [00401224]
'007278f3    dd5dac                  fstp qword ptr [ebp-54]
'var_50 = (00)
'007278f6    8b4de8                  mov ecx, dword ptr [ebp-18]
'007278f9    51                      push ecx

' *** Reference to "__vbaR8Str"
'007278fa    ff1524124000            call dword ptr [00401224]
'00727900    dc4dac                  fmul qword ptr [ebp-54]
'00727903    83ec08                  sub esp, 08
'00727906    dfe0                    fnstsw ax
'00727908    a80d                    test al, 0d
'0072790a    0f85a9000000            jne 7279b9
'00727910    dd1c24                  fstp qword ptr [esp]
'var_131 = (00)

' *** Reference to "__vbaStrR8"
'00727913    ff1580114000            call dword ptr [00401180]
'00727919    8bd0                    mov edx, eax
'0072791b    8d4de0                  lea ecx, dword ptr [ebp-20]

' *** Reference to "__vbaStrMove"
'0072791e    ff15d0124000            call dword ptr [004012d0]
'00727924    50                      push eax
'00727925    57                      push edi
'00727926    ff96a4000000            call dword ptr [esi+000000a4]
'0072792c    dbe2                    fnclex
'0072792e    85c0                    test ax, ax
'00727930    7d12                    jge 727944
'00727932    68a4000000              push 000000a4
'00727937    68d00d4300              push 00430dd0
'0072793c    57                      push edi
'0072793d    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0072793e    ff1580104000            call dword ptr [00401080]
'00727944    8d55e0                  lea edx, dword ptr [ebp-20]
'00727947    52                      push edx
'00727948    8d45e8                  lea eax, dword ptr [ebp-18]
'0072794b    50                      push eax
'0072794c    8d4de4                  lea ecx, dword ptr [ebp-1c]
'0072794f    51                      push ecx
'00727950    6a03                    push 03

' *** Reference to "__vbaFreeStrList"
'00727952    ff155c124000            call dword ptr [0040125c]
'#Cleanup( 0, , -4504)
'00727958    8d55d4                  lea edx, dword ptr [ebp-2c]
'0072795b    52                      push edx
'0072795c    8d45d8                  lea eax, dword ptr [ebp-28]
'0072795f    50                      push eax
'00727960    8d4ddc                  lea ecx, dword ptr [ebp-24]
'00727963    51                      push ecx
'00727964    6a03                    push 03

' *** Reference to "__vbaFreeObjList"
'00727966    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_39, var_45, var_44)
'0072796c    9b                      fwait
'0072796d    83c420                  add esp, 20
'00727970    68a4797200              push 007279a4
'00727975    eb2c                    jmp 7279a3
'00727977    8d55e0                  lea edx, dword ptr [ebp-20]
'0072797a    52                      push edx
'0072797b    8d45e4                  lea eax, dword ptr [ebp-1c]
'0072797e    50                      push eax
'0072797f    8d4de8                  lea ecx, dword ptr [ebp-18]
'00727982    51                      push ecx
'00727983    6a03                    push 03

' *** Reference to "__vbaFreeStrList"
'00727985    ff155c124000            call dword ptr [0040125c]
'#Cleanup( , 0, -4504)
'0072798b    8d55d4                  lea edx, dword ptr [ebp-2c]
'0072798e    52                      push edx
'0072798f    8d45d8                  lea eax, dword ptr [ebp-28]
'00727992    50                      push eax
'00727993    8d4ddc                  lea ecx, dword ptr [ebp-24]
'00727996    51                      push ecx
'00727997    6a03                    push 03

' *** Reference to "__vbaFreeObjList"
'00727999    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_39, var_45, var_44)
'0072799f    83c420                  add esp, 20
'007279a2    c3                      ret
'007279a3    c3                      ret
'007279a4    8b4df0                  mov ecx, dword ptr [ebp-10]
'007279a7    5f                      pop edi
'007279a8    5e                      pop esi
'007279a9    33c0                    xor eax, eax
    'Reference to '__except_list'
'007279ab    64890d00000000          mov dword ptr fs:[00000000], ecx
'007279b2    5b                      pop ebx
'007279b3    8be5                    mov esp, ebp
'007279b5    5d                      pop ebp
'007279b6    c20400                  ret 0004


End Function


'Event for btnSupprimer
Private Sub btnSupprimer_Click()
'00724bf0    55                      push ebp
'00724bf1    8bec                    mov ebp, esp
'00724bf3    83ec0c                  sub esp, 0c

' *** Reference to "__vbaExceptHandler"
'00724bf6    6866724000              push 00407266
'00724bfb    64a100000000            mov ax, word ptr fs:[00000000]
'00724c01    50                      push eax
    'Reference to '__except_list'
'00724c02    64892500000000          mov dword ptr fs:[00000000], esp
'00724c09    83ec18                  sub esp, 18
'00724c0c    53                      push ebx
'00724c0d    56                      push esi
'00724c0e    57                      push edi
'00724c0f    8965f4                  mov dword ptr [ebp-0c], esp
'00724c12    c745f890714000          mov dword ptr [ebp-08], 00407190
'00724c19    8b7d08                  mov edi, dword ptr [ebp+08]
'00724c1c    8bc7                    mov eax, edi
'00724c1e    83e001                  and eax, 01
'00724c21    8945fc                  mov dword ptr [ebp-04], eax
'00724c24    83e7fe                  and edi, fffffffe
'00724c27    8b0f                    mov ecx, dword ptr [edi]
'00724c29    57                      push edi
'00724c2a    897d08                  mov dword ptr [ebp+08], edi
'00724c2d    ff5104                  call dword ptr [ecx+04]
'00724c30    a184b07200              mov ax, word ptr [0072b084]
'00724c35    85c0                    test ax, ax
'00724c37    c745e800000000          mov dword ptr [ebp-18], 00000000
'00724c3e    7515                    jne 724c55
'00724c40    6884b07200              push 0072b084
'00724c45    68f8894100              push 004189f8

' *** Reference to "__vbaNew2"
'00724c4a    ff152c124000            call dword ptr [0040122c]
Dim var_28 As New frmGestPerso
'00724c50    a184b07200              mov ax, word ptr [0072b084]
'00724c55    8b10                    mov edx, dword ptr [eax]
'00724c57    6a00                    push 00
'00724c59    6881000000              push 00000081
'00724c5e    50                      push eax
'00724c5f    ff9240040000            call dword ptr [edx+00000440]
'00724c65    50                      push eax
'00724c66    8d45e8                  lea eax, dword ptr [ebp-18]
'00724c69    50                      push eax

' *** Reference to "__vbaObjSet"
'00724c6a    ff15b4104000            call dword ptr [004010b4]
Set var_41 = var_28
'00724c70    50                      push eax

' *** Reference to "__vbaLateIdCall"
'00724c71    ff1538104000            call dword ptr [00401038]
Call frmGestPerso.Member_0x81()

' *** Reference to "__vbaFreeObj"
'00724c77    8b1d08134000            mov ebx, dword ptr [00401308]
'00724c7d    83c40c                  add esp, 0c
'00724c80    8d4de8                  lea ecx, dword ptr [ebp-18]
'00724c83    ffd3                    call ebx
'#Cleanup(var_41)
'00724c85    a124be7200              mov ax, word ptr [0072be24]
'00724c8a    85c0                    test ax, ax
'00724c8c    7510                    jne 724c9e
'00724c8e    6824be7200              push 0072be24
'00724c93    6870004300              push 00430070

' *** Reference to "__vbaNew2"
'00724c98    ff152c124000            call dword ptr [0040122c]
Dim var_2 As New Global
'00724c9e    8b3524be7200            mov esi, dword ptr [0072be24]
'00724ca4    8b16                    mov edx, dword ptr [esi]
'00724ca6    57                      push edi
'00724ca7    8d4de8                  lea ecx, dword ptr [ebp-18]
'00724caa    51                      push ecx
'00724cab    8955d4                  mov dword ptr [ebp-2c], edx

' *** Reference to "__vbaObjSetAddref"
'00724cae    ff15c8104000            call dword ptr [004010c8]
Set var_41 = Me
'00724cb4    8b55d4                  mov edx, dword ptr [ebp-2c]
'00724cb7    50                      push eax
'00724cb8    56                      push esi
'00724cb9    ff5210                  call dword ptr [edx+10]
Call var_2.Unload(var_41)
'00724cbc    dbe2                    fnclex
'00724cbe    85c0                    test ax, ax
'00724cc0    7d0f                    jge 724cd1
'00724cc2    6a10                    push 10
'00724cc4    6860004300              push 00430060
'00724cc9    56                      push esi
'00724cca    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00724ccb    ff1580104000            call dword ptr [00401080]
'00724cd1    8d4de8                  lea ecx, dword ptr [ebp-18]
'00724cd4    ffd3                    call ebx
'#Cleanup(var_41)
'00724cd6    c745fc00000000          mov dword ptr [ebp-04], 00000000
'00724cdd    68ef4c7200              push 00724cef
'00724ce2    eb0a                    jmp 724cee
'00724ce4    8d4de8                  lea ecx, dword ptr [ebp-18]

' *** Reference to "__vbaFreeObj"
'00724ce7    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_41)
'00724ced    c3                      ret
'00724cee    c3                      ret
'00724cef    8b4508                  mov eax, dword ptr [ebp+08]
'00724cf2    8b08                    mov ecx, dword ptr [eax]
'00724cf4    50                      push eax
'00724cf5    ff5108                  call dword ptr [ecx+08]
'00724cf8    8b45fc                  mov eax, dword ptr [ebp-04]
'00724cfb    8b4dec                  mov ecx, dword ptr [ebp-14]
'00724cfe    5f                      pop edi
'00724cff    5e                      pop esi
    'Reference to '__except_list'
'00724d00    64890d00000000          mov dword ptr fs:[00000000], ecx
'00724d07    5b                      pop ebx
'00724d08    8be5                    mov esp, ebp
'00724d0a    5d                      pop ebp
'00724d0b    c20400                  ret 0004


End Sub


'Event for btnAnnuler
Private Sub btnAnnuler_Click()
'007246f0    55                      push ebp
'007246f1    8bec                    mov ebp, esp
'007246f3    83ec14                  sub esp, 14

' *** Reference to "__vbaExceptHandler"
'007246f6    6866724000              push 00407266
'007246fb    64a100000000            mov ax, word ptr fs:[00000000]
'00724701    50                      push eax
    'Reference to '__except_list'
'00724702    64892500000000          mov dword ptr fs:[00000000], esp
'00724709    81ec30010000            sub esp, 00000130
'0072470f    53                      push ebx
'00724710    56                      push esi
'00724711    57                      push edi
'00724712    8965ec                  mov dword ptr [ebp-14], esp
'00724715    c745f068714000          mov dword ptr [ebp-10], 00407168
'0072471c    8b5d08                  mov ebx, dword ptr [ebp+08]
'0072471f    8bc3                    mov eax, ebx
'00724721    83e001                  and eax, 01
'00724724    8945f4                  mov dword ptr [ebp-0c], eax
'00724727    83e3fe                  and ebx, fffffffe
'0072472a    895d08                  mov dword ptr [ebp+08], ebx
'0072472d    33f6                    xor esi, esi
'0072472f    8975f8                  mov dword ptr [ebp-08], esi
'00724732    8b0b                    mov ecx, dword ptr [ebx]
'00724734    53                      push ebx
'00724735    ff5104                  call dword ptr [ecx+04]
'00724738    8975e0                  mov dword ptr [ebp-20], esi
'0072473b    8975dc                  mov dword ptr [ebp-24], esi
'0072473e    8975d8                  mov dword ptr [ebp-28], esi
'00724741    8975d4                  mov dword ptr [ebp-2c], esi
'00724744    8975d0                  mov dword ptr [ebp-30], esi
'00724747    8975cc                  mov dword ptr [ebp-34], esi
'0072474a    8975c8                  mov dword ptr [ebp-38], esi
'0072474d    8975c4                  mov dword ptr [ebp-3c], esi
'00724750    8975c0                  mov dword ptr [ebp-40], esi
'00724753    8975b0                  mov dword ptr [ebp-50], esi
'00724756    8975a0                  mov dword ptr [ebp-60], esi
'00724759    897590                  mov dword ptr [ebp-70], esi
'0072475c    897580                  mov dword ptr [ebp-80], esi
'0072475f    89b570ffffff            mov dword ptr [ebp+ffffff70], esi
'00724765    89b560ffffff            mov dword ptr [ebp+ffffff60], esi
'0072476b    89b550ffffff            mov dword ptr [ebp+ffffff50], esi
'00724771    89b540ffffff            mov dword ptr [ebp+ffffff40], esi
'00724777    89b530ffffff            mov dword ptr [ebp+ffffff30], esi
'0072477d    89b520ffffff            mov dword ptr [ebp+ffffff20], esi
'00724783    89b510ffffff            mov dword ptr [ebp+ffffff10], esi
'00724789    89b500ffffff            mov dword ptr [ebp+ffffff00], esi
'0072478f    89b5f0feffff            mov dword ptr [ebp+fffffef0], esi
'00724795    89b5e0feffff            mov dword ptr [ebp+fffffee0], esi
'0072479b    89b5dcfeffff            mov dword ptr [ebp+fffffedc], esi
'007247a1    66393528b07200          cmp word ptr [0072b028], si
'007247a8    7508                    jne 7247b2
'007247aa    6a01                    push 01

' *** Reference to "__vbaOnError"
'007247ac    ff15b0104000            call dword ptr [004010b0]
On Error Goto handler_0
'007247b2    393524be7200            cmp dword ptr [0072be24], esi
'007247b8    7510                    jne 7247ca
'007247ba    6824be7200              push 0072be24
'007247bf    6870004300              push 00430070

' *** Reference to "__vbaNew2"
'007247c4    ff152c124000            call dword ptr [0040122c]
Dim var_2 As New Global
'007247ca    8b3d24be7200            mov edi, dword ptr [0072be24]
'007247d0    8b17                    mov edx, dword ptr [edi]
'007247d2    53                      push ebx
'007247d3    8d45c4                  lea eax, dword ptr [ebp-3c]
'007247d6    50                      push eax
'007247d7    8995b4feffff            mov dword ptr [ebp+fffffeb4], edx

' *** Reference to "__vbaObjSetAddref"
'007247dd    ff15c8104000            call dword ptr [004010c8]
Set var_9 = Me
'007247e3    50                      push eax
'007247e4    57                      push edi
'007247e5    8b8db4feffff            mov ecx, dword ptr [ebp+fffffeb4]
'007247eb    ff5110                  call dword ptr [ecx+10]
Call var_2.Unload(var_9)
'007247ee    dbe2                    fnclex
'007247f0    3bc6                    cmp eax, esi
'007247f2    7d0f                    jge 724803
'007247f4    6a10                    push 10
'007247f6    6860004300              push 00430060
'007247fb    57                      push edi
'007247fc    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'007247fd    ff1580104000            call dword ptr [00401080]
'00724803    8d4dc4                  lea ecx, dword ptr [ebp-3c]

' *** Reference to "__vbaFreeObj"
'00724806    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_9)

' *** Reference to "__vbaExitProc"
'0072480c    ff15a0104000            call dword ptr [004010a0]
'00724812    68c54b7200              push 00724bc5
'00724817    e9a8030000              jmp 724bc4

' *** Reference to "rtcErrObj"
'0072481c    8b1d6c124000            mov ebx, dword ptr [0040126c]
'00724822    ffd3                    call ebx
'00724824    50                      push eax
'00724825    8d55c4                  lea edx, dword ptr [ebp-3c]
'00724828    52                      push edx

' *** Reference to "__vbaObjSet"
'00724829    8b3db4104000            mov edi, dword ptr [004010b4]
'0072482f    ffd7                    call edi
Set var_9 = Err
'00724831    8bf0                    mov esi, eax
'00724833    8b06                    mov eax, dword ptr [esi]
'00724835    8d4de0                  lea ecx, dword ptr [ebp-20]
'00724838    51                      push ecx
'00724839    56                      push esi
'0072483a    ff502c                  call dword ptr [eax+2c]
var_3 = var_9.Description
'0072483d    dbe2                    fnclex
'0072483f    85c0                    test ax, ax
'00724841    7d0f                    jge 724852
'00724843    6a2c                    push 2c
'00724845    685c084300              push 0043085c
'0072484a    56                      push esi
'0072484b    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0072484c    ff1580104000            call dword ptr [00401080]
'00724852    ffd3                    call ebx
'00724854    50                      push eax
'00724855    8d55c0                  lea edx, dword ptr [ebp-40]
'00724858    52                      push edx
'00724859    ffd7                    call edi
Set var_5 = Err
'0072485b    8bf0                    mov esi, eax
'0072485d    8b06                    mov eax, dword ptr [esi]
'0072485f    8d8ddcfeffff            lea ecx, dword ptr [ebp+fffffedc]
'00724865    51                      push ecx
'00724866    56                      push esi
'00724867    ff501c                  call dword ptr [eax+1c]
var_10 = var_5.Number
'0072486a    dbe2                    fnclex
'0072486c    85c0                    test ax, ax
'0072486e    7d0f                    jge 72487f
'00724870    6a1c                    push 1c
'00724872    685c084300              push 0043085c
'00724877    56                      push esi
'00724878    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00724879    ff1580104000            call dword ptr [00401080]
'0072487f    b904000280              mov ecx, 80020004
'00724884    894d98                  mov dword ptr [ebp-68], ecx
'00724887    b80a000000              mov eax, 0000000a
'0072488c    894590                  mov dword ptr [ebp-70], eax
'0072488f    894da8                  mov dword ptr [ebp-58], ecx
'00724892    8945a0                  mov dword ptr [ebp-60], eax
'00724895    c78528ffffff24b07200    mov dword ptr [ebp+ffffff28], 0072b024
'0072489f    c78520ffffff08400000    mov dword ptr [ebp+ffffff20], 00004008
'007248a9    6814084300              push 00430814
'007248ae    8b55e0                  mov edx, dword ptr [ebp-20]
'007248b1    52                      push edx

' *** Reference to "__vbaStrCat"
'007248b2    8b3570104000            mov esi, dword ptr [00401070]
'007248b8    ffd6                    call esi
var_11 = ("L'erreur suivante s'est produite : ") & (var_3)
'007248ba    8bd0                    mov edx, eax
'007248bc    8d4ddc                  lea ecx, dword ptr [ebp-24]

' *** Reference to "__vbaStrMove"
'007248bf    8b3dd0124000            mov edi, dword ptr [004012d0]
'007248c5    ffd7                    call edi
'007248c7    50                      push eax
'007248c8    6870084300              push 00430870
'007248cd    ffd6                    call esi
var_12 = (var_11) & (vbCrLf)
'007248cf    8bd0                    mov edx, eax
'007248d1    8d4dd8                  lea ecx, dword ptr [ebp-28]
'007248d4    ffd7                    call edi
'007248d6    50                      push eax
'007248d7    6870084300              push 00430870
'007248dc    ffd6                    call esi
var_13 = (var_12) & (vbCrLf)
'007248de    8bd0                    mov edx, eax
'007248e0    8d4dd4                  lea ecx, dword ptr [ebp-2c]
'007248e3    ffd7                    call edi
'007248e5    50                      push eax
'007248e6    6880084300              push 00430880
'007248eb    ffd6                    call esi
var_14 = (var_13) & (" numéro d'erreur (")
'007248ed    8bd0                    mov edx, eax
'007248ef    8d4dd0                  lea ecx, dword ptr [ebp-30]
'007248f2    ffd7                    call edi
'007248f4    50                      push eax
'007248f5    8b85dcfeffff            mov eax, dword ptr [ebp+fffffedc]
'007248fb    50                      push eax

' *** Reference to "__vbaStrI4"
'007248fc    ff1520104000            call dword ptr [00401020]
'00724902    8bd0                    mov edx, eax
'00724904    8d4dcc                  lea ecx, dword ptr [ebp-34]
'00724907    ffd7                    call edi
'00724909    50                      push eax
'0072490a    ffd6                    call esi
var_15 = (var_14) & (CStr(var_10))
'0072490c    8bd0                    mov edx, eax
'0072490e    8d4dc8                  lea ecx, dword ptr [ebp-38]
'00724911    ffd7                    call edi
'00724913    50                      push eax
'00724914    68ac084300              push 004308ac
'00724919    ffd6                    call esi
var_16 = (var_15) & (" )")
'0072491b    8945b8                  mov dword ptr [ebp-48], eax
'0072491e    bf08000000              mov edi, 00000008
'00724923    897db0                  mov dword ptr [ebp-50], edi
'00724926    8d4d90                  lea ecx, dword ptr [ebp-70]
'00724929    51                      push ecx
'0072492a    8d55a0                  lea edx, dword ptr [ebp-60]
'0072492d    52                      push edx
'0072492e    8d8520ffffff            lea eax, dword ptr [ebp+ffffff20]
'00724934    50                      push eax
'00724935    6a10                    push 10
'00724937    8d4db0                  lea ecx, dword ptr [ebp-50]
'0072493a    51                      push ecx

' *** Reference to "rtcMsgBox"
'0072493b    ff15b8104000            call dword ptr [004010b8]
var_17 = MsgBox(var_16, 16, vbNullString)
'00724941    8d55c8                  lea edx, dword ptr [ebp-38]
'00724944    52                      push edx
'00724945    8d45cc                  lea eax, dword ptr [ebp-34]
'00724948    50                      push eax
'00724949    8d4dd0                  lea ecx, dword ptr [ebp-30]
'0072494c    51                      push ecx
'0072494d    8d55d4                  lea edx, dword ptr [ebp-2c]
'00724950    52                      push edx
'00724951    8d45d8                  lea eax, dword ptr [ebp-28]
'00724954    50                      push eax
'00724955    8d4ddc                  lea ecx, dword ptr [ebp-24]
'00724958    51                      push ecx
'00724959    8d55e0                  lea edx, dword ptr [ebp-20]
'0072495c    52                      push edx
'0072495d    6a07                    push 07

' *** Reference to "__vbaFreeStrList"
'0072495f    ff155c124000            call dword ptr [0040125c]
'#Cleanup( 0, -4512, -4516, -4520, -4524, -4528, -4532)
'00724965    8d45c0                  lea eax, dword ptr [ebp-40]
'00724968    50                      push eax
'00724969    8d4dc4                  lea ecx, dword ptr [ebp-3c]
'0072496c    51                      push ecx
'0072496d    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'0072496f    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_9, var_5)
'00724975    8d5590                  lea edx, dword ptr [ebp-70]
'00724978    52                      push edx
'00724979    8d45a0                  lea eax, dword ptr [ebp-60]
'0072497c    50                      push eax
'0072497d    8d4db0                  lea ecx, dword ptr [ebp-50]
'00724980    51                      push ecx
'00724981    6a03                    push 03

' *** Reference to "__vbaFreeVarList"
'00724983    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_6, var_7, var_8)
'00724989    83c43c                  add esp, 3c
'0072498c    8d55b0                  lea edx, dword ptr [ebp-50]
'0072498f    52                      push edx

' *** Reference to "rtcGetPresentDate"
'00724990    ff15f4124000            call dword ptr [004012f4]
var_16 = Now()
'00724996    c78528ffffffb8084300    mov dword ptr [ebp+ffffff28], 004308b8
'007249a0    89bd20ffffff            mov dword ptr [ebp+ffffff20], edi
'007249a6    8d9520ffffff            lea edx, dword ptr [ebp+ffffff20]
'007249ac    8d4da0                  lea ecx, dword ptr [ebp-60]

' *** Reference to "__vbaVarDup"
'007249af    ff1598124000            call dword ptr [00401298]
var_7 = ("dd/MM/yyyy hh:mm:ss")
'007249b5    6a01                    push 01
'007249b7    6a01                    push 01
'007249b9    8d45a0                  lea eax, dword ptr [ebp-60]
'007249bc    50                      push eax
'007249bd    8d4db0                  lea ecx, dword ptr [ebp-50]
'007249c0    51                      push ecx
'007249c1    8d5590                  lea edx, dword ptr [ebp-70]
'007249c4    52                      push edx

' *** Reference to "rtcVarFromFormatVar"
'007249c5    ff1574104000            call dword ptr [00401074]
'007249cb    ffd3                    call ebx
'007249cd    50                      push eax
'007249ce    8d45c4                  lea eax, dword ptr [ebp-3c]
'007249d1    50                      push eax

' *** Reference to "__vbaObjSet"
'007249d2    ff15b4104000            call dword ptr [004010b4]
Set var_9 = Err
'007249d8    8bf0                    mov esi, eax
'007249da    8b0e                    mov ecx, dword ptr [esi]
'007249dc    8d95dcfeffff            lea edx, dword ptr [ebp+fffffedc]
'007249e2    52                      push edx
'007249e3    56                      push esi
'007249e4    ff511c                  call dword ptr [ecx+1c]
var_10 = var_9.Number
'007249e7    dbe2                    fnclex
'007249e9    85c0                    test ax, ax
'007249eb    7d0f                    jge 7249fc
'007249ed    6a1c                    push 1c
'007249ef    685c084300              push 0043085c
'007249f4    56                      push esi
'007249f5    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'007249f6    ff1580104000            call dword ptr [00401080]
'007249fc    ffd3                    call ebx
'007249fe    50                      push eax
'007249ff    8d45c0                  lea eax, dword ptr [ebp-40]
'00724a02    50                      push eax

' *** Reference to "__vbaObjSet"
'00724a03    ff15b4104000            call dword ptr [004010b4]
Set var_5 = Err
'00724a09    8bf0                    mov esi, eax
'00724a0b    8b0e                    mov ecx, dword ptr [esi]
'00724a0d    8d55e0                  lea edx, dword ptr [ebp-20]
'00724a10    52                      push edx
'00724a11    56                      push esi
'00724a12    ff512c                  call dword ptr [ecx+2c]
var_3 = var_5.Description
'00724a15    dbe2                    fnclex
'00724a17    85c0                    test ax, ax
'00724a19    7d0f                    jge 724a2a
'00724a1b    6a2c                    push 2c
'00724a1d    685c084300              push 0043085c
'00724a22    56                      push esi
'00724a23    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00724a24    ff1580104000            call dword ptr [00401080]
'00724a2a    c78518ffffffe4084300    mov dword ptr [ebp+ffffff18], 004308e4
'00724a34    89bd10ffffff            mov dword ptr [ebp+ffffff10], edi
'00724a3a    8b85dcfeffff            mov eax, dword ptr [ebp+fffffedc]
'00724a40    898508ffffff            mov dword ptr [ebp+ffffff08], eax
'00724a46    c78500ffffff03000000    mov dword ptr [ebp+ffffff00], 00000003
'00724a50    c785f8feffff08094300    mov dword ptr [ebp+fffffef8], 00430908
'00724a5a    89bdf0feffff            mov dword ptr [ebp+fffffef0], edi
'00724a60    8b45e0                  mov eax, dword ptr [ebp-20]
'00724a63    c745e000000000          mov dword ptr [ebp-20], 00000000
'00724a6a    898558ffffff            mov dword ptr [ebp+ffffff58], eax
'00724a70    89bd50ffffff            mov dword ptr [ebp+ffffff50], edi
'00724a76    c785e8feffffd4794500    mov dword ptr [ebp+fffffee8], 004579d4
'00724a80    89bde0feffff            mov dword ptr [ebp+fffffee0], edi
'00724a86    8d4d90                  lea ecx, dword ptr [ebp-70]
'00724a89    51                      push ecx
'00724a8a    8d9510ffffff            lea edx, dword ptr [ebp+ffffff10]
'00724a90    52                      push edx
'00724a91    8d4580                  lea eax, dword ptr [ebp-80]
'00724a94    50                      push eax

' *** Reference to "__vbaVarCat"
'00724a95    8b3508124000            mov esi, dword ptr [00401208]
'00724a9b    ffd6                    call esi
'00724a9d    50                      push eax
'00724a9e    8d8d00ffffff            lea ecx, dword ptr [ebp+ffffff00]
'00724aa4    51                      push ecx
'00724aa5    8d9570ffffff            lea edx, dword ptr [ebp+ffffff70]
'00724aab    52                      push edx
'00724aac    ffd6                    call esi
'00724aae    50                      push eax
'00724aaf    8d85f0feffff            lea eax, dword ptr [ebp+fffffef0]
'00724ab5    50                      push eax
'00724ab6    8d8d60ffffff            lea ecx, dword ptr [ebp+ffffff60]
'00724abc    51                      push ecx
'00724abd    ffd6                    call esi
'00724abf    50                      push eax
'00724ac0    8d9550ffffff            lea edx, dword ptr [ebp+ffffff50]
'00724ac6    52                      push edx
'00724ac7    8d8540ffffff            lea eax, dword ptr [ebp+ffffff40]
'00724acd    50                      push eax
'00724ace    ffd6                    call esi
'00724ad0    50                      push eax
'00724ad1    8d8de0feffff            lea ecx, dword ptr [ebp+fffffee0]
'00724ad7    51                      push ecx
'00724ad8    8d9530ffffff            lea edx, dword ptr [ebp+ffffff30]
'00724ade    52                      push edx
'00724adf    ffd6                    call esi
'00724ae1    50                      push eax
'00724ae2    33c0                    xor eax, eax
'00724ae4    66a12ab07200            mov eax, dword ptr [0072b02a]
'00724aea    50                      push eax
'00724aeb    6884094300              push 00430984

' *** Reference to "__vbaPrintFile"
'00724af0    ff15b8114000            call dword ptr [004011b8]
Print #0, 
'00724af6    8d4dc0                  lea ecx, dword ptr [ebp-40]
'00724af9    51                      push ecx
'00724afa    8d55c4                  lea edx, dword ptr [ebp-3c]
'00724afd    52                      push edx
'00724afe    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'00724b00    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_9, var_5)
'00724b06    8d8530ffffff            lea eax, dword ptr [ebp+ffffff30]
'00724b0c    50                      push eax
'00724b0d    8d8d40ffffff            lea ecx, dword ptr [ebp+ffffff40]
'00724b13    51                      push ecx
'00724b14    8d9550ffffff            lea edx, dword ptr [ebp+ffffff50]
'00724b1a    52                      push edx
'00724b1b    8d8560ffffff            lea eax, dword ptr [ebp+ffffff60]
'00724b21    50                      push eax
'00724b22    8d8d70ffffff            lea ecx, dword ptr [ebp+ffffff70]
'00724b28    51                      push ecx
'00724b29    8d5580                  lea edx, dword ptr [ebp-80]
'00724b2c    52                      push edx
'00724b2d    8d4590                  lea eax, dword ptr [ebp-70]
'00724b30    50                      push eax
'00724b31    8d4da0                  lea ecx, dword ptr [ebp-60]
'00724b34    51                      push ecx
'00724b35    8d55b0                  lea edx, dword ptr [ebp-50]
'00724b38    52                      push edx
'00724b39    6a09                    push 09

' *** Reference to "__vbaFreeVarList"
'00724b3b    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_6, var_7, var_8, var_18, var_19, var_20, var_21, var_22, var_23)
'00724b41    83c440                  add esp, 40
'00724b44    6a00                    push 00

' *** Reference to "__vbaResume"
'00724b46    ff1568104000            call dword ptr [00401068]
'00724b4c    e9bbfcffff              jmp 72480c
Resume handler_72480C
'00724b51    8d55c8                  lea edx, dword ptr [ebp-38]
'00724b54    52                      push edx
'00724b55    8d45cc                  lea eax, dword ptr [ebp-34]
'00724b58    50                      push eax
'00724b59    8d4dd0                  lea ecx, dword ptr [ebp-30]
'00724b5c    51                      push ecx
'00724b5d    8d55d4                  lea edx, dword ptr [ebp-2c]
'00724b60    52                      push edx
'00724b61    8d45d8                  lea eax, dword ptr [ebp-28]
'00724b64    50                      push eax
'00724b65    8d4ddc                  lea ecx, dword ptr [ebp-24]
'00724b68    51                      push ecx
'00724b69    8d55e0                  lea edx, dword ptr [ebp-20]
'00724b6c    52                      push edx
'00724b6d    6a07                    push 07

' *** Reference to "__vbaFreeStrList"
'00724b6f    ff155c124000            call dword ptr [0040125c]
'#Cleanup( var_3, -4512, -4516, -4520, -4524, -4528, -4532)
'00724b75    8d45c0                  lea eax, dword ptr [ebp-40]
'00724b78    50                      push eax
'00724b79    8d4dc4                  lea ecx, dword ptr [ebp-3c]
'00724b7c    51                      push ecx
'00724b7d    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'00724b7f    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_9, var_5)
'00724b85    8d9530ffffff            lea edx, dword ptr [ebp+ffffff30]
'00724b8b    52                      push edx
'00724b8c    8d8540ffffff            lea eax, dword ptr [ebp+ffffff40]
'00724b92    50                      push eax
'00724b93    8d8d50ffffff            lea ecx, dword ptr [ebp+ffffff50]
'00724b99    51                      push ecx
'00724b9a    8d9560ffffff            lea edx, dword ptr [ebp+ffffff60]
'00724ba0    52                      push edx
'00724ba1    8d8570ffffff            lea eax, dword ptr [ebp+ffffff70]
'00724ba7    50                      push eax
'00724ba8    8d4d80                  lea ecx, dword ptr [ebp-80]
'00724bab    51                      push ecx
'00724bac    8d5590                  lea edx, dword ptr [ebp-70]
'00724baf    52                      push edx
'00724bb0    8d45a0                  lea eax, dword ptr [ebp-60]
'00724bb3    50                      push eax
'00724bb4    8d4db0                  lea ecx, dword ptr [ebp-50]
'00724bb7    51                      push ecx
'00724bb8    6a09                    push 09

' *** Reference to "__vbaFreeVarList"
'00724bba    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_6, var_7, var_8, var_18, var_19, var_20, var_21, var_22, var_23)
'00724bc0    83c454                  add esp, 54
'00724bc3    c3                      ret
'00724bc4    c3                      ret
'00724bc5    8b4508                  mov eax, dword ptr [ebp+08]
'00724bc8    8b10                    mov edx, dword ptr [eax]
'00724bca    50                      push eax
'00724bcb    ff5208                  call dword ptr [edx+08]
'00724bce    8b45f4                  mov eax, dword ptr [ebp-0c]
'00724bd1    8b4de4                  mov ecx, dword ptr [ebp-1c]
    'Reference to '__except_list'
'00724bd4    64890d00000000          mov dword ptr fs:[00000000], ecx
'00724bdb    5f                      pop edi
'00724bdc    5e                      pop esi
'00724bdd    5b                      pop ebx
'00724bde    8be5                    mov esp, ebp
'00724be0    5d                      pop ebp
'00724be1    c20400                  ret 0004


End Sub


'Event for Form
Private Sub Form_Load()
'007273a0    55                      push ebp
'007273a1    8bec                    mov ebp, esp
'007273a3    83ec0c                  sub esp, 0c

' *** Reference to "__vbaExceptHandler"
'007273a6    6866724000              push 00407266
'007273ab    64a100000000            mov ax, word ptr fs:[00000000]
'007273b1    50                      push eax
    'Reference to '__except_list'
'007273b2    64892500000000          mov dword ptr fs:[00000000], esp
'007273b9    83ec40                  sub esp, 40
'007273bc    53                      push ebx
'007273bd    56                      push esi
'007273be    57                      push edi
'007273bf    8965f4                  mov dword ptr [ebp-0c], esp
'007273c2    c745f8f8714000          mov dword ptr [ebp-08], 004071f8
'007273c9    8b7508                  mov esi, dword ptr [ebp+08]
'007273cc    8bc6                    mov eax, esi
'007273ce    83e001                  and eax, 01
'007273d1    8945fc                  mov dword ptr [ebp-04], eax
'007273d4    83e6fe                  and esi, fffffffe
'007273d7    8b0e                    mov ecx, dword ptr [esi]
'007273d9    56                      push esi
'007273da    897508                  mov dword ptr [ebp+08], esi
'007273dd    ff5104                  call dword ptr [ecx+04]
'007273e0    8b16                    mov edx, dword ptr [esi]
'007273e2    33db                    xor ebx, ebx
'007273e4    56                      push esi
'007273e5    895de8                  mov dword ptr [ebp-18], ebx
'007273e8    895de4                  mov dword ptr [ebp-1c], ebx
'007273eb    895de0                  mov dword ptr [ebp-20], ebx
'007273ee    895ddc                  mov dword ptr [ebp-24], ebx
'007273f1    895dd8                  mov dword ptr [ebp-28], ebx
'007273f4    895dd4                  mov dword ptr [ebp-2c], ebx
'007273f7    ff9200030000            call dword ptr [edx+00000300]
'007273fd    50                      push eax
'007273fe    8d45dc                  lea eax, dword ptr [ebp-24]
'00727401    50                      push eax

' *** Reference to "__vbaObjSet"
'00727402    ff15b4104000            call dword ptr [004010b4]
Set var_39 = Me
'00727408    8b5640                  mov edx, dword ptr [esi+40]
'0072740b    8bf8                    mov edi, eax
'0072740d    8b0f                    mov ecx, dword ptr [edi]
'0072740f    52                      push edx
'00727410    57                      push edi
'00727411    ff5154                  call dword ptr [ecx+54]
'00727414    dbe2                    fnclex
'00727416    3bc3                    cmp eax, ebx
'00727418    7d0f                    jge 727429

If (var_39 < 0) Then
'0072741a    6a54                    push 54
'0072741c    6874f14300              push 0043f174
'00727421    57                      push edi
'00727422    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00727423    ff1580104000            call dword ptr [00401080]
    
End If
'00727429    8d4ddc                  lea ecx, dword ptr [ebp-24]

' *** Reference to "__vbaFreeObj"
'0072742c    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_39)
'00727432    8b06                    mov eax, dword ptr [esi]
'00727434    56                      push esi
'00727435    ff9010030000            call dword ptr [eax+00000310]
'0072743b    50                      push eax
'0072743c    8d4ddc                  lea ecx, dword ptr [ebp-24]
'0072743f    51                      push ecx

' *** Reference to "__vbaObjSet"
'00727440    ff15b4104000            call dword ptr [004010b4]
Set var_39 = Nothing
'00727446    8bf8                    mov edi, eax
'00727448    8b17                    mov edx, dword ptr [edi]
'0072744a    8d45d8                  lea eax, dword ptr [ebp-28]
'0072744d    50                      push eax
'0072744e    53                      push ebx
'0072744f    57                      push edi
'00727450    ff5240                  call dword ptr [edx+40]
'00727453    dbe2                    fnclex
'00727455    3bc3                    cmp eax, ebx
'00727457    7d0f                    jge 727468
'00727459    6a40                    push 40
'0072745b    686c384300              push 0043386c
'00727460    57                      push edi
'00727461    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00727462    ff1580104000            call dword ptr [00401080]
'00727468    8b4e38                  mov ecx, dword ptr [esi+38]
'0072746b    8b45d8                  mov eax, dword ptr [ebp-28]
'0072746e    8b5634                  mov edx, dword ptr [esi+34]
'00727471    8b18                    mov ebx, dword ptr [eax]
'00727473    68247d4500              push 00457d24
'00727478    51                      push ecx
'00727479    52                      push edx
'0072747a    8945c8                  mov dword ptr [ebp-38], eax

' *** Reference to "__vbaStrR8"
'0072747d    ff1580114000            call dword ptr [00401180]

' *** Reference to "__vbaStrMove"
'00727483    8b3dd0124000            mov edi, dword ptr [004012d0]
'00727489    8bd0                    mov edx, eax
'0072748b    8d4de8                  lea ecx, dword ptr [ebp-18]
'0072748e    ffd7                    call edi
'00727490    50                      push eax

' *** Reference to "__vbaStrCat"
'00727491    ff1570104000            call dword ptr [00401070]
var_84 = ("Prix neuf de l'objet : ") & (CStr(arg_6))
'00727497    8bd0                    mov edx, eax
'00727499    8d4de4                  lea ecx, dword ptr [ebp-1c]
'0072749c    ffd7                    call edi
'0072749e    50                      push eax
'0072749f    683c724500              push 0045723c

' *** Reference to "__vbaStrCat"
'007274a4    ff1570104000            call dword ptr [00401070]
var_63 = (var_84) & (" Po")
'007274aa    8bd0                    mov edx, eax
'007274ac    8d4de0                  lea ecx, dword ptr [ebp-20]
'007274af    ffd7                    call edi
'007274b1    8b7dc8                  mov edi, dword ptr [ebp-38]
'007274b4    50                      push eax
'007274b5    57                      push edi
'007274b6    ff5354                  call dword ptr [ebx+54]
'007274b9    dbe2                    fnclex
'007274bb    85c0                    test ax, ax
'007274bd    7d0f                    jge 7274ce
'007274bf    6a54                    push 54
'007274c1    683ce04300              push 0043e03c
'007274c6    57                      push edi
'007274c7    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'007274c8    ff1580104000            call dword ptr [00401080]
'007274ce    8d45e0                  lea eax, dword ptr [ebp-20]
'007274d1    50                      push eax
'007274d2    8d4de4                  lea ecx, dword ptr [ebp-1c]
'007274d5    51                      push ecx
'007274d6    8d55e8                  lea edx, dword ptr [ebp-18]
'007274d9    52                      push edx
'007274da    6a03                    push 03

' *** Reference to "__vbaFreeStrList"
'007274dc    ff155c124000            call dword ptr [0040125c]
'#Cleanup( -4496, -4500, -4504)
'007274e2    8d45d8                  lea eax, dword ptr [ebp-28]
'007274e5    50                      push eax
'007274e6    8d4ddc                  lea ecx, dword ptr [ebp-24]
'007274e9    51                      push ecx
'007274ea    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'007274ec    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_39, var_45)
'007274f2    8b16                    mov edx, dword ptr [esi]
'007274f4    83c41c                  add esp, 1c
'007274f7    56                      push esi
'007274f8    ff9208030000            call dword ptr [edx+00000308]
'007274fe    50                      push eax
'007274ff    8d45dc                  lea eax, dword ptr [ebp-24]
'00727502    50                      push eax

' *** Reference to "__vbaObjSet"
'00727503    ff15b4104000            call dword ptr [004010b4]
Set var_39 = 
'00727509    33c9                    xor ecx, ecx
'0072750b    668b4e3c                mov cx, word ptr [esi+3c]
'0072750f    8bf8                    mov edi, eax
'00727511    8b1f                    mov ebx, dword ptr [edi]
'00727513    51                      push ecx

' *** Reference to "__vbaStrI2"
'00727514    ff1510104000            call dword ptr [00401010]
'0072751a    8bd0                    mov edx, eax
'0072751c    8d4de8                  lea ecx, dword ptr [ebp-18]

' *** Reference to "__vbaStrMove"
'0072751f    ff15d0124000            call dword ptr [004012d0]
'00727525    50                      push eax
'00727526    57                      push edi
'00727527    ff93a4000000            call dword ptr [ebx+000000a4]
'0072752d    dbe2                    fnclex
'0072752f    85c0                    test ax, ax
'00727531    7d12                    jge 727545
'00727533    68a4000000              push 000000a4
'00727538    68d00d4300              push 00430dd0
'0072753d    57                      push edi
'0072753e    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0072753f    ff1580104000            call dword ptr [00401080]
'00727545    8d4de8                  lea ecx, dword ptr [ebp-18]

' *** Reference to "__vbaFreeStr"
'00727548    ff150c134000            call dword ptr [0040130c]
'#Cleanup(var_41)
'0072754e    8d4ddc                  lea ecx, dword ptr [ebp-24]

' *** Reference to "__vbaFreeObj"
'00727551    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_39)
'00727557    8b16                    mov edx, dword ptr [esi]
'00727559    56                      push esi
'0072755a    ff920c030000            call dword ptr [edx+0000030c]
'00727560    50                      push eax
'00727561    8d45dc                  lea eax, dword ptr [ebp-24]
'00727564    50                      push eax

' *** Reference to "__vbaObjSet"
'00727565    ff15b4104000            call dword ptr [004010b4]
Set var_39 = -4508
'0072756b    dd4634                  fld qword ptr [esi+34]
'0072756e    833d00b0720000          cmp dword ptr [0072b000], 00
'00727575    7508                    jne 72757f

If (vbNullChar = 0) Then
'00727577    dc35d8154000            fdiv qword ptr [004015d8]
'0072757d    eb11                    jmp 727590
    
End If
'0072757f    ff35dc154000            push dword ptr [004015dc]
'00727585    ff35d8154000            push dword ptr [004015d8]

' *** Reference to "_adj_fdiv_m64"
'0072758b    e8f4fccdff              call 407284
'00727590    8bf8                    mov edi, eax
'00727592    8b1f                    mov ebx, dword ptr [edi]
'00727594    83ec08                  sub esp, 08
'00727597    dfe0                    fnstsw ax
'00727599    a80d                    test al, 0d
'0072759b    0f85af010000            jne 727750
'007275a1    dd1c24                  fstp qword ptr [esp]
'var_7 = (00)

' *** Reference to "__vbaStrR8"
'007275a4    ff1580114000            call dword ptr [00401180]
'007275aa    8bd0                    mov edx, eax
'007275ac    8d4de8                  lea ecx, dword ptr [ebp-18]

' *** Reference to "__vbaStrMove"
'007275af    ff15d0124000            call dword ptr [004012d0]
'007275b5    50                      push eax
'007275b6    57                      push edi
'007275b7    ff93a4000000            call dword ptr [ebx+000000a4]
'007275bd    dbe2                    fnclex
'007275bf    85c0                    test ax, ax
'007275c1    7d12                    jge 7275d5
'007275c3    68a4000000              push 000000a4
'007275c8    68d00d4300              push 00430dd0
'007275cd    57                      push edi
'007275ce    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'007275cf    ff1580104000            call dword ptr [00401080]
'007275d5    8d4de8                  lea ecx, dword ptr [ebp-18]

' *** Reference to "__vbaFreeStr"
'007275d8    ff150c134000            call dword ptr [0040130c]
'#Cleanup(var_41)
'007275de    8d4ddc                  lea ecx, dword ptr [ebp-24]

' *** Reference to "__vbaFreeObj"
'007275e1    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_39)
'007275e7    8b0e                    mov ecx, dword ptr [esi]
'007275e9    56                      push esi
'007275ea    ff9108030000            call dword ptr [ecx+00000308]

' *** Reference to "__vbaObjSet"
'007275f0    8b1db4104000            mov ebx, dword ptr [004010b4]
'007275f6    50                      push eax
'007275f7    8d55dc                  lea edx, dword ptr [ebp-24]
'007275fa    52                      push edx
'007275fb    ffd3                    call ebx
Set var_39 = -4512
'007275fd    8d4de8                  lea ecx, dword ptr [ebp-18]
'00727600    8bf8                    mov edi, eax
'00727602    8b07                    mov eax, dword ptr [edi]
'00727604    51                      push ecx
'00727605    57                      push edi
'00727606    ff90a0000000            call dword ptr [eax+000000a0]
'0072760c    dbe2                    fnclex
'0072760e    85c0                    test ax, ax
'00727610    7d12                    jge 727624
'00727612    68a0000000              push 000000a0
'00727617    68d00d4300              push 00430dd0
'0072761c    57                      push edi
'0072761d    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0072761e    ff1580104000            call dword ptr [00401080]
'00727624    8b16                    mov edx, dword ptr [esi]
'00727626    56                      push esi
'00727627    ff9204030000            call dword ptr [edx+00000304]
'0072762d    50                      push eax
'0072762e    8d45d4                  lea eax, dword ptr [ebp-2c]
'00727631    50                      push eax
'00727632    ffd3                    call ebx
Set var_44 = CStr(((arg_6) / 2#))
'00727634    8b0e                    mov ecx, dword ptr [esi]
'00727636    56                      push esi
'00727637    8bf8                    mov edi, eax
'00727639    ff910c030000            call dword ptr [ecx+0000030c]
'0072763f    50                      push eax
'00727640    8d55d8                  lea edx, dword ptr [ebp-28]
'00727643    52                      push edx
'00727644    ffd3                    call ebx
Set var_45 = var_44
'00727646    8d4de4                  lea ecx, dword ptr [ebp-1c]
'00727649    8bf0                    mov esi, eax
'0072764b    8b06                    mov eax, dword ptr [esi]
'0072764d    51                      push ecx
'0072764e    56                      push esi
'0072764f    ff90a0000000            call dword ptr [eax+000000a0]
'00727655    dbe2                    fnclex
'00727657    85c0                    test ax, ax
'00727659    7d12                    jge 72766d
'0072765b    68a0000000              push 000000a0
'00727660    68d00d4300              push 00430dd0
'00727665    56                      push esi
'00727666    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00727667    ff1580104000            call dword ptr [00401080]
'0072766d    8b55e4                  mov edx, dword ptr [ebp-1c]
'00727670    8b37                    mov esi, dword ptr [edi]
'00727672    52                      push edx

' *** Reference to "__vbaR8Str"
'00727673    ff1524124000            call dword ptr [00401224]
'00727679    dd5dac                  fstp qword ptr [ebp-54]
'var_50 = (00)
'0072767c    8b45e8                  mov eax, dword ptr [ebp-18]
'0072767f    50                      push eax

' *** Reference to "__vbaR8Str"
'00727680    ff1524124000            call dword ptr [00401224]
'00727686    dc4dac                  fmul qword ptr [ebp-54]
'00727689    83ec08                  sub esp, 08
'0072768c    dfe0                    fnstsw ax
'0072768e    a80d                    test al, 0d
'00727690    0f85ba000000            jne 727750
'00727696    dd1c24                  fstp qword ptr [esp]
'var_121 = (00)

' *** Reference to "__vbaStrR8"
'00727699    ff1580114000            call dword ptr [00401180]
'0072769f    8bd0                    mov edx, eax
'007276a1    8d4de0                  lea ecx, dword ptr [ebp-20]

' *** Reference to "__vbaStrMove"
'007276a4    ff15d0124000            call dword ptr [004012d0]
'007276aa    50                      push eax
'007276ab    57                      push edi
'007276ac    ff96a4000000            call dword ptr [esi+000000a4]
'007276b2    dbe2                    fnclex
'007276b4    85c0                    test ax, ax
'007276b6    7d12                    jge 7276ca
'007276b8    68a4000000              push 000000a4
'007276bd    68d00d4300              push 00430dd0
'007276c2    57                      push edi
'007276c3    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'007276c4    ff1580104000            call dword ptr [00401080]
'007276ca    8d4de0                  lea ecx, dword ptr [ebp-20]
'007276cd    51                      push ecx
'007276ce    8d55e8                  lea edx, dword ptr [ebp-18]
'007276d1    52                      push edx
'007276d2    8d45e4                  lea eax, dword ptr [ebp-1c]
'007276d5    50                      push eax
'007276d6    6a03                    push 03

' *** Reference to "__vbaFreeStrList"
'007276d8    ff155c124000            call dword ptr [0040125c]
'#Cleanup( -4500, , -4516)
'007276de    8d4dd4                  lea ecx, dword ptr [ebp-2c]
'007276e1    51                      push ecx
'007276e2    8d55d8                  lea edx, dword ptr [ebp-28]
'007276e5    52                      push edx
'007276e6    8d45dc                  lea eax, dword ptr [ebp-24]
'007276e9    50                      push eax
'007276ea    6a03                    push 03

' *** Reference to "__vbaFreeObjList"
'007276ec    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_39, var_45, var_44)
'007276f2    83c420                  add esp, 20
'007276f5    c745fc00000000          mov dword ptr [ebp-04], 00000000
'007276fc    9b                      fwait
'007276fd    6831777200              push 00727731
'00727702    eb2c                    jmp 727730
'00727704    8d4de0                  lea ecx, dword ptr [ebp-20]
'00727707    51                      push ecx
'00727708    8d55e4                  lea edx, dword ptr [ebp-1c]
'0072770b    52                      push edx
'0072770c    8d45e8                  lea eax, dword ptr [ebp-18]
'0072770f    50                      push eax
'00727710    6a03                    push 03

' *** Reference to "__vbaFreeStrList"
'00727712    ff155c124000            call dword ptr [0040125c]
'#Cleanup( , -4500, -4516)
'00727718    8d4dd4                  lea ecx, dword ptr [ebp-2c]
'0072771b    51                      push ecx
'0072771c    8d55d8                  lea edx, dword ptr [ebp-28]
'0072771f    52                      push edx
'00727720    8d45dc                  lea eax, dword ptr [ebp-24]
'00727723    50                      push eax
'00727724    6a03                    push 03

' *** Reference to "__vbaFreeObjList"
'00727726    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_39, var_45, var_44)
'0072772c    83c420                  add esp, 20
'0072772f    c3                      ret
'00727730    c3                      ret
'00727731    8b4508                  mov eax, dword ptr [ebp+08]
'00727734    8b08                    mov ecx, dword ptr [eax]
'00727736    50                      push eax
'00727737    ff5108                  call dword ptr [ecx+08]
'0072773a    8b45fc                  mov eax, dword ptr [ebp-04]
'0072773d    8b4dec                  mov ecx, dword ptr [ebp-14]
'00727740    5f                      pop edi
'00727741    5e                      pop esi
    'Reference to '__except_list'
'00727742    64890d00000000          mov dword ptr fs:[00000000], ecx
'00727749    5b                      pop ebx
'0072774a    8be5                    mov esp, ebp
'0072774c    5d                      pop ebp
'0072774d    c20400                  ret 0004


End Sub


Private Sub Form_KeyDown(KeyCode as Integer, Shift as Integer)
'007279c0    55                      push ebp
'007279c1    8bec                    mov ebp, esp
'007279c3    83ec0c                  sub esp, 0c

' *** Reference to "__vbaExceptHandler"
'007279c6    6866724000              push 00407266
'007279cb    64a100000000            mov ax, word ptr fs:[00000000]
'007279d1    50                      push eax
    'Reference to '__except_list'
'007279d2    64892500000000          mov dword ptr fs:[00000000], esp
'007279d9    83ec18                  sub esp, 18
'007279dc    53                      push ebx
'007279dd    56                      push esi
'007279de    57                      push edi
'007279df    8965f4                  mov dword ptr [ebp-0c], esp
'007279e2    c745f818724000          mov dword ptr [ebp-08], 00407218
'007279e9    8b7d08                  mov edi, dword ptr [ebp+08]
'007279ec    8bc7                    mov eax, edi
'007279ee    83e001                  and eax, 01
'007279f1    8945fc                  mov dword ptr [ebp-04], eax
'007279f4    83e7fe                  and edi, fffffffe
'007279f7    8b0f                    mov ecx, dword ptr [edi]
'007279f9    57                      push edi
'007279fa    897d08                  mov dword ptr [ebp+08], edi
'007279fd    ff5104                  call dword ptr [ecx+04]
'00727a00    8b550c                  mov edx, dword ptr [ebp+0c]
'00727a03    33db                    xor ebx, ebx
'00727a05    66833a1b                cmp word ptr [edx], 1b
'00727a09    895de8                  mov dword ptr [ebp-18], ebx
'00727a0c    7554                    jne 727a62

If (arg_0 = 27) Then
'00727a0e    391d24be7200            cmp dword ptr [0072be24], ebx
'00727a14    7510                    jne 727a26
'00727a16    6824be7200              push 0072be24
'00727a1b    6870004300              push 00430070

' *** Reference to "__vbaNew2"
'00727a20    ff152c124000            call dword ptr [0040122c]
    Dim var_2 As New Global
'00727a26    8b3524be7200            mov esi, dword ptr [0072be24]
'00727a2c    8b16                    mov edx, dword ptr [esi]
'00727a2e    57                      push edi
'00727a2f    8d45e8                  lea eax, dword ptr [ebp-18]
'00727a32    50                      push eax
'00727a33    8955d4                  mov dword ptr [ebp-2c], edx

' *** Reference to "__vbaObjSetAddref"
'00727a36    ff15c8104000            call dword ptr [004010c8]
    Set var_41 = Me
'00727a3c    8b4dd4                  mov ecx, dword ptr [ebp-2c]
'00727a3f    50                      push eax
'00727a40    56                      push esi
'00727a41    ff5110                  call dword ptr [ecx+10]
    Call var_2.Unload(var_41)
'00727a44    dbe2                    fnclex
'00727a46    3bc3                    cmp eax, ebx
'00727a48    7d0f                    jge 727a59
'00727a4a    6a10                    push 10
'00727a4c    6860004300              push 00430060
'00727a51    56                      push esi
'00727a52    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00727a53    ff1580104000            call dword ptr [00401080]
'00727a59    8d4de8                  lea ecx, dword ptr [ebp-18]

' *** Reference to "__vbaFreeObj"
'00727a5c    ff1508134000            call dword ptr [00401308]
    '#Cleanup(var_41)
End If
'00727a62    895dfc                  mov dword ptr [ebp-04], ebx
'00727a65    68777a7200              push 00727a77
'00727a6a    eb0a                    jmp 727a76
'00727a6c    8d4de8                  lea ecx, dword ptr [ebp-18]

' *** Reference to "__vbaFreeObj"
'00727a6f    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_41)
'00727a75    c3                      ret
'00727a76    c3                      ret
'00727a77    8b4508                  mov eax, dword ptr [ebp+08]
'00727a7a    8b10                    mov edx, dword ptr [eax]
'00727a7c    50                      push eax
'00727a7d    ff5208                  call dword ptr [edx+08]
'00727a80    8b45fc                  mov eax, dword ptr [ebp-04]
'00727a83    8b4dec                  mov ecx, dword ptr [ebp-14]
'00727a86    5f                      pop edi
'00727a87    5e                      pop esi
    'Reference to '__except_list'
'00727a88    64890d00000000          mov dword ptr fs:[00000000], ecx
'00727a8f    5b                      pop ebx
'00727a90    8be5                    mov esp, ebp
'00727a92    5d                      pop ebp
'00727a93    c20c00                  ret 000c


End Sub


'Event for fldnPrixUnitaire
Private Sub fldnPrixUnitaire_Change()
'00726c60    55                      push ebp
'00726c61    8bec                    mov ebp, esp
'00726c63    83ec0c                  sub esp, 0c

' *** Reference to "__vbaExceptHandler"
'00726c66    6866724000              push 00407266
'00726c6b    64a100000000            mov ax, word ptr fs:[00000000]
'00726c71    50                      push eax
    'Reference to '__except_list'
'00726c72    64892500000000          mov dword ptr fs:[00000000], esp
'00726c79    81ec9c000000            sub esp, 0000009c
'00726c7f    53                      push ebx
'00726c80    56                      push esi
'00726c81    57                      push edi
'00726c82    8965f4                  mov dword ptr [ebp-0c], esp
'00726c85    c745f8e8714000          mov dword ptr [ebp-08], 004071e8
'00726c8c    8b7508                  mov esi, dword ptr [ebp+08]
'00726c8f    8bc6                    mov eax, esi
'00726c91    83e001                  and eax, 01
'00726c94    8945fc                  mov dword ptr [ebp-04], eax
'00726c97    83e6fe                  and esi, fffffffe
'00726c9a    8b0e                    mov ecx, dword ptr [esi]
'00726c9c    56                      push esi
'00726c9d    897508                  mov dword ptr [ebp+08], esi
'00726ca0    ff5104                  call dword ptr [ecx+04]
'00726ca3    8b16                    mov edx, dword ptr [esi]
'00726ca5    33c0                    xor eax, eax
var_num1 = Empty
'00726ca7    56                      push esi
'00726ca8    8945e8                  mov dword ptr [ebp-18], eax
'00726cab    8945e4                  mov dword ptr [ebp-1c], eax
'00726cae    8945d4                  mov dword ptr [ebp-2c], eax
'00726cb1    8945c4                  mov dword ptr [ebp-3c], eax
'00726cb4    8945b4                  mov dword ptr [ebp-4c], eax
'00726cb7    8945a4                  mov dword ptr [ebp-5c], eax
'00726cba    894594                  mov dword ptr [ebp-6c], eax
'00726cbd    ff920c030000            call dword ptr [edx+0000030c]

' *** Reference to "__vbaObjSet"
'00726cc3    8b1db4104000            mov ebx, dword ptr [004010b4]
'00726cc9    50                      push eax
'00726cca    8d45e4                  lea eax, dword ptr [ebp-1c]
'00726ccd    50                      push eax
'00726cce    ffd3                    call ebx
Set var_40 = Me
'00726cd0    8d55e8                  lea edx, dword ptr [ebp-18]
'00726cd3    8bf8                    mov edi, eax
'00726cd5    8b0f                    mov ecx, dword ptr [edi]
'00726cd7    52                      push edx
'00726cd8    57                      push edi
'00726cd9    ff91a0000000            call dword ptr [ecx+000000a0]
'00726cdf    dbe2                    fnclex
'00726ce1    85c0                    test ax, ax
'00726ce3    7d12                    jge 726cf7
'00726ce5    68a0000000              push 000000a0
'00726cea    68d00d4300              push 00430dd0
'00726cef    57                      push edi
'00726cf0    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00726cf1    ff1580104000            call dword ptr [00401080]
'00726cf7    8b45e8                  mov eax, dword ptr [ebp-18]
'00726cfa    50                      push eax
'00726cfb    68cc134300              push 004313cc

' *** Reference to "__vbaStrCmp"
'00726d00    ff1530114000            call dword ptr [00401130]
'00726d06    8bf8                    mov edi, eax
'00726d08    f7df                    neg edi
'00726d0a    1bff                    sbb edi, edi
'00726d0c    f7df                    neg edi
'00726d0e    8d4de8                  lea ecx, dword ptr [ebp-18]
'00726d11    f7df                    neg edi

' *** Reference to "__vbaFreeStr"
'00726d13    ff150c134000            call dword ptr [0040130c]
'#Cleanup(var_41)
'00726d19    8d4de4                  lea ecx, dword ptr [ebp-1c]

' *** Reference to "__vbaFreeObj"
'00726d1c    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_40)
'00726d22    6685ff                  test edi, edi
'00726d25    0f8443040000            je 72716e

If (((frmVendre) <> (vbNullChar))) Then
'00726d2b    8b0e                    mov ecx, dword ptr [esi]
'00726d2d    56                      push esi
'00726d2e    ff910c030000            call dword ptr [ecx+0000030c]
'00726d34    8d55d4                  lea edx, dword ptr [ebp-2c]
'00726d37    52                      push edx
'00726d38    8945dc                  mov dword ptr [ebp-24], eax
'00726d3b    c745d409000000          mov dword ptr [ebp-2c], 00000009

' *** Reference to "rtcIsNumeric"
'00726d42    ff154c114000            call dword ptr [0040114c]
'00726d48    33ff                    xor edi, edi
    var_num7 = Empty
'00726d4a    668bf8                  mov di, ax
'00726d4d    8d4dd4                  lea ecx, dword ptr [ebp-2c]
'00726d50    f7d7                    not edi

' *** Reference to "__vbaFreeVar"
'00726d52    ff1528104000            call dword ptr [00401028]
    '#Cleanup(var_44)
'00726d58    6685ff                  test edi, edi
'00726d5b    0f84c1010000            je 726f22
    
    If (    IsNumeric(((frmVendre) [#!#] (vbNullChar)))) Then
'00726d61    a124be7200              mov ax, word ptr [0072be24]
'00726d66    85c0                    test ax, ax
'00726d68    7510                    jne 726d7a
'00726d6a    6824be7200              push 0072be24
'00726d6f    6870004300              push 00430070

' *** Reference to "__vbaNew2"
'00726d74    ff152c124000            call dword ptr [0040122c]
    Dim var_2 As New Global
'00726d7a    8b3d24be7200            mov edi, dword ptr [0072be24]
'00726d80    8b07                    mov eax, dword ptr [edi]
'00726d82    8d4de4                  lea ecx, dword ptr [ebp-1c]
'00726d85    51                      push ecx
'00726d86    57                      push edi
'00726d87    ff5018                  call dword ptr [eax+18]
    Set var_40 = var_2.Screen
'00726d8a    dbe2                    fnclex
'00726d8c    85c0                    test ax, ax
'00726d8e    7d0f                    jge 726d9f
'00726d90    6a18                    push 18
'00726d92    6860004300              push 00430060
'00726d97    57                      push edi
'00726d98    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00726d99    ff1580104000            call dword ptr [00401080]
'00726d9f    8b7de4                  mov edi, dword ptr [ebp-1c]
'00726da2    8b1f                    mov ebx, dword ptr [edi]
'00726da4    33c9                    xor ecx, ecx

' *** Reference to "__vbaI2I4"
'00726da6    ff1550114000            call dword ptr [00401150]
'00726dac    50                      push eax
'00726dad    57                      push edi
'00726dae    ff537c                  call dword ptr [ebx+7c]
    var_40.MousePointer = CInt(0)
'00726db1    dbe2                    fnclex
'00726db3    85c0                    test ax, ax
'00726db5    7d0f                    jge 726dc6
'00726db7    6a7c                    push 7c
'00726db9    6810be4300              push 0043be10
'00726dbe    57                      push edi
'00726dbf    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00726dc0    ff1580104000            call dword ptr [00401080]
'00726dc6    8d4de4                  lea ecx, dword ptr [ebp-1c]

' *** Reference to "__vbaFreeObj"
'00726dc9    ff1508134000            call dword ptr [00401308]
    '#Cleanup(var_40)
'00726dcf    8b16                    mov edx, dword ptr [esi]
'00726dd1    8d45e8                  lea eax, dword ptr [ebp-18]
'00726dd4    50                      push eax
'00726dd5    56                      push esi
'00726dd6    ff5248                  call dword ptr [edx+48]
'00726dd9    dbe2                    fnclex
'00726ddb    85c0                    test ax, ax
'00726ddd    7d0f                    jge 726dee
'00726ddf    6a48                    push 48
'00726de1    6810f54400              push 0044f510
'00726de6    56                      push esi
'00726de7    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00726de8    ff1580104000            call dword ptr [00401080]
'00726dee    b80a000000              mov eax, 0000000a
'00726df3    8945a4                  mov dword ptr [ebp-5c], eax
'00726df6    8945b4                  mov dword ptr [ebp-4c], eax
'00726df9    8b45e8                  mov eax, dword ptr [ebp-18]
'00726dfc    b904000280              mov ecx, 80020004
'00726e01    8945cc                  mov dword ptr [ebp-34], eax
'00726e04    b808000000              mov eax, 00000008
'00726e09    894dac                  mov dword ptr [ebp-54], ecx
'00726e0c    894dbc                  mov dword ptr [ebp-44], ecx
'00726e0f    8d5594                  lea edx, dword ptr [ebp-6c]
'00726e12    8d4dd4                  lea ecx, dword ptr [ebp-2c]
'00726e15    c745e800000000          mov dword ptr [ebp-18], 00000000
'00726e1c    8945c4                  mov dword ptr [ebp-3c], eax
'00726e1f    c7459c447c4500          mov dword ptr [ebp-64], 00457c44
'00726e26    894594                  mov dword ptr [ebp-6c], eax

' *** Reference to "__vbaVarDup"
'00726e29    ff1598124000            call dword ptr [00401298]
    var_44 = ("Le prix unitaire doit être entrée en tant que valeur numérique")
'00726e2f    8d4da4                  lea ecx, dword ptr [ebp-5c]
'00726e32    51                      push ecx
'00726e33    8d55b4                  lea edx, dword ptr [ebp-4c]
'00726e36    52                      push edx
'00726e37    8d45c4                  lea eax, dword ptr [ebp-3c]
'00726e3a    50                      push eax
'00726e3b    6a30                    push 30
'00726e3d    8d4dd4                  lea ecx, dword ptr [ebp-2c]
'00726e40    51                      push ecx

' *** Reference to "rtcMsgBox"
'00726e41    ff15b8104000            call dword ptr [004010b8]
    var_14 = MsgBox(var_44, 48, vbNullString)
'00726e47    8d55a4                  lea edx, dword ptr [ebp-5c]
'00726e4a    52                      push edx
'00726e4b    8d45b4                  lea eax, dword ptr [ebp-4c]
'00726e4e    50                      push eax
'00726e4f    8d4dc4                  lea ecx, dword ptr [ebp-3c]
'00726e52    51                      push ecx
'00726e53    8d55d4                  lea edx, dword ptr [ebp-2c]
'00726e56    52                      push edx
'00726e57    6a04                    push 04

' *** Reference to "__vbaFreeVarList"
'00726e59    ff1540104000            call dword ptr [00401040]
    '#Cleanup( var_44, var_9, var_62, var_83)
'00726e5f    8b06                    mov eax, dword ptr [esi]
'00726e61    83c414                  add esp, 14
'00726e64    56                      push esi
'00726e65    ff900c030000            call dword ptr [eax+0000030c]
'00726e6b    50                      push eax
'00726e6c    8d4de4                  lea ecx, dword ptr [ebp-1c]
'00726e6f    51                      push ecx

' *** Reference to "__vbaObjSet"
'00726e70    ff15b4104000            call dword ptr [004010b4]
    Set var_40 = Nothing
'00726e76    dd4634                  fld qword ptr [esi+34]
'00726e79    833d00b0720000          cmp dword ptr [0072b000], 00
'00726e80    7508                    jne 726e8a
    
    If (    vbNullChar = 0) Then
'00726e82    dc35d8154000            fdiv qword ptr [004015d8]
'00726e88    eb11                    jmp 726e9b
    
End If
'00726e8a    ff35dc154000            push dword ptr [004015dc]
'00726e90    ff35d8154000            push dword ptr [004015d8]

' *** Reference to "_adj_fdiv_m64"
'00726e96    e8e903ceff              call 407284
'00726e9b    8bf8                    mov edi, eax
'00726e9d    8b1f                    mov ebx, dword ptr [edi]
'00726e9f    83ec08                  sub esp, 08
'00726ea2    dfe0                    fnstsw ax
'00726ea4    a80d                    test al, 0d
'00726ea6    0f85ee040000            jne 72739a
'00726eac    dd1c24                  fstp qword ptr [esp]
'var_134 = (00)

' *** Reference to "__vbaStrR8"
'00726eaf    ff1580114000            call dword ptr [00401180]
'00726eb5    8bd0                    mov edx, eax
'00726eb7    8d4de8                  lea ecx, dword ptr [ebp-18]

' *** Reference to "__vbaStrMove"
'00726eba    ff15d0124000            call dword ptr [004012d0]
'00726ec0    50                      push eax
'00726ec1    57                      push edi
'00726ec2    ff93a4000000            call dword ptr [ebx+000000a4]
'00726ec8    dbe2                    fnclex
'00726eca    85c0                    test ax, ax
'00726ecc    7d12                    jge 726ee0
'00726ece    68a4000000              push 000000a4
'00726ed3    68d00d4300              push 00430dd0
'00726ed8    57                      push edi
'00726ed9    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00726eda    ff1580104000            call dword ptr [00401080]
'00726ee0    8d4de8                  lea ecx, dword ptr [ebp-18]

' *** Reference to "__vbaFreeStr"
'00726ee3    ff150c134000            call dword ptr [0040130c]
'#Cleanup(var_41)

' *** Reference to "__vbaFreeObj"
'00726ee9    8b3d08134000            mov edi, dword ptr [00401308]
'00726eef    8d4de4                  lea ecx, dword ptr [ebp-1c]
'00726ef2    ffd7                    call edi
'#Cleanup(var_40)
'00726ef4    8b16                    mov edx, dword ptr [esi]
'00726ef6    56                      push esi
'00726ef7    ff920c030000            call dword ptr [edx+0000030c]
'00726efd    50                      push eax
'00726efe    8d45e4                  lea eax, dword ptr [ebp-1c]
'00726f01    50                      push eax

' *** Reference to "__vbaObjSet"
'00726f02    ff15b4104000            call dword ptr [004010b4]
Set var_40 = -4528
'00726f08    8bf0                    mov esi, eax
'00726f0a    8b0e                    mov ecx, dword ptr [esi]
'00726f0c    56                      push esi
'00726f0d    ff9104020000            call dword ptr [ecx+00000204]
'00726f13    dbe2                    fnclex
'00726f15    85c0                    test ax, ax
'00726f17    0f8d1b040000            jge 727338
'00726f1d    e904040000              jmp 727326

'ERROR: Two many next close:
End If
'00726f22    8b16                    mov edx, dword ptr [esi]
'00726f24    56                      push esi
'00726f25    ff920c030000            call dword ptr [edx+0000030c]
'00726f2b    50                      push eax
'00726f2c    8d45e4                  lea eax, dword ptr [ebp-1c]
'00726f2f    50                      push eax
'00726f30    ffd3                    call ebx
Set var_40 = IsNumeric(((frmVendre) [##] (vbNullChar)))
'00726f32    8d55e8                  lea edx, dword ptr [ebp-18]
'00726f35    8bf8                    mov edi, eax
'00726f37    8b0f                    mov ecx, dword ptr [edi]
'00726f39    52                      push edx
'00726f3a    57                      push edi
'00726f3b    ff91a0000000            call dword ptr [ecx+000000a0]
'00726f41    dbe2                    fnclex
'00726f43    85c0                    test ax, ax
'00726f45    7d12                    jge 726f59
'00726f47    68a0000000              push 000000a0
'00726f4c    68d00d4300              push 00430dd0
'00726f51    57                      push edi
'00726f52    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00726f53    ff1580104000            call dword ptr [00401080]
'00726f59    8b45e8                  mov eax, dword ptr [ebp-18]
'00726f5c    50                      push eax

' *** Reference to "rtcR8ValFromBstr"
'00726f5d    ff1510134000            call dword ptr [00401310]

' *** Reference to "__vbaFpR8"
'00726f63    ff15f8104000            call dword ptr [004010f8]
'00726f69    dc1d68154000            fcomp qword ptr [00401568]
'00726f6f    dfe0                    fnstsw ax
'00726f71    f6c401                  test ah, 01
'00726f74    7407                    je 726f7d
'00726f76    b801000000              mov eax, 00000001
'00726f7b    eb02                    jmp 726f7f
'00726f7d    33c0                    xor eax, eax
'00726f7f    f7d8                    neg eax
'00726f81    8d4de8                  lea ecx, dword ptr [ebp-18]
'00726f84    668bf8                  mov di, ax

' *** Reference to "__vbaFreeStr"
'00726f87    ff150c134000            call dword ptr [0040130c]
'#Cleanup(var_41)
'00726f8d    8d4de4                  lea ecx, dword ptr [ebp-1c]

' *** Reference to "__vbaFreeObj"
'00726f90    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_40)
'00726f96    6685ff                  test edi, edi
'00726f99    0f84c1010000            je 727160

If ((0# > Val(vbNullString))) Then
'00726f9f    a124be7200              mov ax, word ptr [0072be24]
'00726fa4    85c0                    test ax, ax
'00726fa6    7510                    jne 726fb8
'00726fa8    6824be7200              push 0072be24
'00726fad    6870004300              push 00430070

' *** Reference to "__vbaNew2"
'00726fb2    ff152c124000            call dword ptr [0040122c]
    Set var_2 = New Global
'00726fb8    8b3d24be7200            mov edi, dword ptr [0072be24]
'00726fbe    8b0f                    mov ecx, dword ptr [edi]
'00726fc0    8d55e4                  lea edx, dword ptr [ebp-1c]
'00726fc3    52                      push edx
'00726fc4    57                      push edi
'00726fc5    ff5118                  call dword ptr [ecx+18]
    Set var_40 = var_2.Screen
'00726fc8    dbe2                    fnclex
'00726fca    85c0                    test ax, ax
'00726fcc    7d0f                    jge 726fdd
'00726fce    6a18                    push 18
'00726fd0    6860004300              push 00430060
'00726fd5    57                      push edi
'00726fd6    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00726fd7    ff1580104000            call dword ptr [00401080]
'00726fdd    8b7de4                  mov edi, dword ptr [ebp-1c]
'00726fe0    8b1f                    mov ebx, dword ptr [edi]
'00726fe2    33c9                    xor ecx, ecx

' *** Reference to "__vbaI2I4"
'00726fe4    ff1550114000            call dword ptr [00401150]
'00726fea    50                      push eax
'00726feb    57                      push edi
'00726fec    ff537c                  call dword ptr [ebx+7c]
    var_40.MousePointer = CInt((0# [#>] Val(vbNullString)))
'00726fef    dbe2                    fnclex
'00726ff1    85c0                    test ax, ax
'00726ff3    7d0f                    jge 727004
'00726ff5    6a7c                    push 7c
'00726ff7    6810be4300              push 0043be10
'00726ffc    57                      push edi
'00726ffd    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00726ffe    ff1580104000            call dword ptr [00401080]
'00727004    8d4de4                  lea ecx, dword ptr [ebp-1c]

' *** Reference to "__vbaFreeObj"
'00727007    ff1508134000            call dword ptr [00401308]
    '#Cleanup(var_40)
'0072700d    8b06                    mov eax, dword ptr [esi]
'0072700f    8d4de8                  lea ecx, dword ptr [ebp-18]
'00727012    51                      push ecx
'00727013    56                      push esi
'00727014    ff5048                  call dword ptr [eax+48]
'00727017    dbe2                    fnclex
'00727019    85c0                    test ax, ax
'0072701b    7d0f                    jge 72702c
'0072701d    6a48                    push 48
'0072701f    6810f54400              push 0044f510
'00727024    56                      push esi
'00727025    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00727026    ff1580104000            call dword ptr [00401080]
'0072702c    b80a000000              mov eax, 0000000a
'00727031    8945a4                  mov dword ptr [ebp-5c], eax
'00727034    8945b4                  mov dword ptr [ebp-4c], eax
'00727037    8b45e8                  mov eax, dword ptr [ebp-18]
'0072703a    b904000280              mov ecx, 80020004
'0072703f    8945cc                  mov dword ptr [ebp-34], eax
'00727042    b808000000              mov eax, 00000008
'00727047    894dac                  mov dword ptr [ebp-54], ecx
'0072704a    894dbc                  mov dword ptr [ebp-44], ecx
'0072704d    8d5594                  lea edx, dword ptr [ebp-6c]
'00727050    8d4dd4                  lea ecx, dword ptr [ebp-2c]
'00727053    c745e800000000          mov dword ptr [ebp-18], 00000000
'0072705a    8945c4                  mov dword ptr [ebp-3c], eax
'0072705d    c7459cc87c4500          mov dword ptr [ebp-64], 00457cc8
'00727064    894594                  mov dword ptr [ebp-6c], eax

' *** Reference to "__vbaVarDup"
'00727067    ff1598124000            call dword ptr [00401298]
    var_44 = ("Le prix unitaire doit être plus grand que 0")
'0072706d    8d55a4                  lea edx, dword ptr [ebp-5c]
'00727070    52                      push edx
'00727071    8d45b4                  lea eax, dword ptr [ebp-4c]
'00727074    50                      push eax
'00727075    8d4dc4                  lea ecx, dword ptr [ebp-3c]
'00727078    51                      push ecx
'00727079    6a30                    push 30
'0072707b    8d55d4                  lea edx, dword ptr [ebp-2c]
'0072707e    52                      push edx

' *** Reference to "rtcMsgBox"
'0072707f    ff15b8104000            call dword ptr [004010b8]
    var_53 = MsgBox(var_44, 48, vbNullString)
'00727085    8d45a4                  lea eax, dword ptr [ebp-5c]
'00727088    50                      push eax
'00727089    8d4db4                  lea ecx, dword ptr [ebp-4c]
'0072708c    51                      push ecx
'0072708d    8d55c4                  lea edx, dword ptr [ebp-3c]
'00727090    52                      push edx
'00727091    8d45d4                  lea eax, dword ptr [ebp-2c]
'00727094    50                      push eax
'00727095    6a04                    push 04

' *** Reference to "__vbaFreeVarList"
'00727097    ff1540104000            call dword ptr [00401040]
    '#Cleanup( var_44, var_9, var_62, var_83)
'0072709d    8b0e                    mov ecx, dword ptr [esi]
'0072709f    83c414                  add esp, 14
'007270a2    56                      push esi
'007270a3    ff9108030000            call dword ptr [ecx+00000308]
'007270a9    50                      push eax
'007270aa    8d55e4                  lea edx, dword ptr [ebp-1c]
'007270ad    52                      push edx

' *** Reference to "__vbaObjSet"
'007270ae    ff15b4104000            call dword ptr [004010b4]
    Set var_40 = 
'007270b4    dd4634                  fld qword ptr [esi+34]
'007270b7    833d00b0720000          cmp dword ptr [0072b000], 00
'007270be    7508                    jne 7270c8
    
    If (    vbNullChar = 0) Then
'007270c0    dc35d8154000            fdiv qword ptr [004015d8]
'007270c6    eb11                    jmp 7270d9
    
End If
'007270c8    ff35dc154000            push dword ptr [004015dc]
'007270ce    ff35d8154000            push dword ptr [004015d8]

' *** Reference to "_adj_fdiv_m64"
'007270d4    e8ab01ceff              call 407284
'007270d9    8bf8                    mov edi, eax
'007270db    8b1f                    mov ebx, dword ptr [edi]
'007270dd    83ec08                  sub esp, 08
'007270e0    dfe0                    fnstsw ax
'007270e2    a80d                    test al, 0d
'007270e4    0f85b0020000            jne 72739a
'007270ea    dd1c24                  fstp qword ptr [esp]
'var_125 = (00)

' *** Reference to "__vbaStrR8"
'007270ed    ff1580114000            call dword ptr [00401180]
'007270f3    8bd0                    mov edx, eax
'007270f5    8d4de8                  lea ecx, dword ptr [ebp-18]

' *** Reference to "__vbaStrMove"
'007270f8    ff15d0124000            call dword ptr [004012d0]
'007270fe    50                      push eax
'007270ff    57                      push edi
'00727100    ff93a4000000            call dword ptr [ebx+000000a4]
'00727106    dbe2                    fnclex
'00727108    85c0                    test ax, ax
'0072710a    7d12                    jge 72711e
'0072710c    68a4000000              push 000000a4
'00727111    68d00d4300              push 00430dd0
'00727116    57                      push edi
'00727117    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00727118    ff1580104000            call dword ptr [00401080]
'0072711e    8d4de8                  lea ecx, dword ptr [ebp-18]

' *** Reference to "__vbaFreeStr"
'00727121    ff150c134000            call dword ptr [0040130c]
'#Cleanup(var_41)

' *** Reference to "__vbaFreeObj"
'00727127    8b3d08134000            mov edi, dword ptr [00401308]
'0072712d    8d4de4                  lea ecx, dword ptr [ebp-1c]
'00727130    ffd7                    call edi
'#Cleanup(var_40)
'00727132    8b06                    mov eax, dword ptr [esi]
'00727134    56                      push esi
'00727135    ff9008030000            call dword ptr [eax+00000308]
'0072713b    50                      push eax
'0072713c    8d4de4                  lea ecx, dword ptr [ebp-1c]
'0072713f    51                      push ecx

' *** Reference to "__vbaObjSet"
'00727140    ff15b4104000            call dword ptr [004010b4]
Set var_40 = Nothing
'00727146    8bf0                    mov esi, eax
'00727148    8b16                    mov edx, dword ptr [esi]
'0072714a    56                      push esi
'0072714b    ff9204020000            call dword ptr [edx+00000204]
'00727151    dbe2                    fnclex
'00727153    85c0                    test ax, ax
'00727155    0f8ddd010000            jge 727338
'0072715b    e9c6010000              jmp 727326

'ERROR: Two many next close:
End If
'00727160    8b06                    mov eax, dword ptr [esi]
'00727162    56                      push esi
'00727163    ff902c070000            call dword ptr [eax+0000072c]
'00727169    e9cf010000              jmp 72733d

'ERROR: Two many next close:
End If
'0072716e    a124be7200              mov ax, word ptr [0072be24]
'00727173    85c0                    test ax, ax
'00727175    7510                    jne 727187
'00727177    6824be7200              push 0072be24
'0072717c    6870004300              push 00430070

' *** Reference to "__vbaNew2"
'00727181    ff152c124000            call dword ptr [0040122c]
Set var_2 = New Global
'00727187    8b3d24be7200            mov edi, dword ptr [0072be24]
'0072718d    8b0f                    mov ecx, dword ptr [edi]
'0072718f    8d55e4                  lea edx, dword ptr [ebp-1c]
'00727192    52                      push edx
'00727193    57                      push edi
'00727194    ff5118                  call dword ptr [ecx+18]
Set var_40 = var_2.Screen
'00727197    dbe2                    fnclex
'00727199    85c0                    test ax, ax
'0072719b    7d0f                    jge 7271ac
'0072719d    6a18                    push 18
'0072719f    6860004300              push 00430060
'007271a4    57                      push edi
'007271a5    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'007271a6    ff1580104000            call dword ptr [00401080]
'007271ac    8b7de4                  mov edi, dword ptr [ebp-1c]
'007271af    8b1f                    mov ebx, dword ptr [edi]
'007271b1    33c9                    xor ecx, ecx
var_num3 = Empty

' *** Reference to "__vbaI2I4"
'007271b3    ff1550114000            call dword ptr [00401150]
'007271b9    50                      push eax
'007271ba    57                      push edi
'007271bb    ff537c                  call dword ptr [ebx+7c]
var_40.MousePointer = CInt(Global)
'007271be    dbe2                    fnclex
'007271c0    85c0                    test ax, ax
'007271c2    7d0f                    jge 7271d3
'007271c4    6a7c                    push 7c
'007271c6    6810be4300              push 0043be10
'007271cb    57                      push edi
'007271cc    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'007271cd    ff1580104000            call dword ptr [00401080]
'007271d3    8d4de4                  lea ecx, dword ptr [ebp-1c]

' *** Reference to "__vbaFreeObj"
'007271d6    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_40)
'007271dc    8b06                    mov eax, dword ptr [esi]
'007271de    8d4de8                  lea ecx, dword ptr [ebp-18]
'007271e1    51                      push ecx
'007271e2    56                      push esi
'007271e3    ff5048                  call dword ptr [eax+48]
'007271e6    dbe2                    fnclex
'007271e8    85c0                    test ax, ax
'007271ea    7d0f                    jge 7271fb
'007271ec    6a48                    push 48
'007271ee    6810f54400              push 0044f510
'007271f3    56                      push esi
'007271f4    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'007271f5    ff1580104000            call dword ptr [00401080]
'007271fb    b80a000000              mov eax, 0000000a
'00727200    8945a4                  mov dword ptr [ebp-5c], eax
'00727203    8945b4                  mov dword ptr [ebp-4c], eax
'00727206    8b45e8                  mov eax, dword ptr [ebp-18]
'00727209    b904000280              mov ecx, 80020004
'0072720e    8945cc                  mov dword ptr [ebp-34], eax
'00727211    b808000000              mov eax, 00000008
'00727216    894dac                  mov dword ptr [ebp-54], ecx
'00727219    894dbc                  mov dword ptr [ebp-44], ecx
'0072721c    8d5594                  lea edx, dword ptr [ebp-6c]
'0072721f    8d4dd4                  lea ecx, dword ptr [ebp-2c]
'00727222    c745e800000000          mov dword ptr [ebp-18], 00000000
'00727229    8945c4                  mov dword ptr [ebp-3c], eax
'0072722c    c7459c447c4500          mov dword ptr [ebp-64], 00457c44
'00727233    894594                  mov dword ptr [ebp-6c], eax

' *** Reference to "__vbaVarDup"
'00727236    ff1598124000            call dword ptr [00401298]
var_44 = ("Le prix unitaire doit être entrée en tant que valeur numérique")
'0072723c    8d55a4                  lea edx, dword ptr [ebp-5c]
'0072723f    52                      push edx
'00727240    8d45b4                  lea eax, dword ptr [ebp-4c]
'00727243    50                      push eax
'00727244    8d4dc4                  lea ecx, dword ptr [ebp-3c]
'00727247    51                      push ecx
'00727248    6a30                    push 30
'0072724a    8d55d4                  lea edx, dword ptr [ebp-2c]
'0072724d    52                      push edx

' *** Reference to "rtcMsgBox"
'0072724e    ff15b8104000            call dword ptr [004010b8]
var_286 = MsgBox(var_44, 48, vbNullString)
'00727254    8d45a4                  lea eax, dword ptr [ebp-5c]
'00727257    50                      push eax
'00727258    8d4db4                  lea ecx, dword ptr [ebp-4c]
'0072725b    51                      push ecx
'0072725c    8d55c4                  lea edx, dword ptr [ebp-3c]
'0072725f    52                      push edx
'00727260    8d45d4                  lea eax, dword ptr [ebp-2c]
'00727263    50                      push eax
'00727264    6a04                    push 04

' *** Reference to "__vbaFreeVarList"
'00727266    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_44, var_9, var_62, var_83)
'0072726c    8b0e                    mov ecx, dword ptr [esi]
'0072726e    83c414                  add esp, 14
'00727271    56                      push esi
'00727272    ff910c030000            call dword ptr [ecx+0000030c]
'00727278    50                      push eax
'00727279    8d55e4                  lea edx, dword ptr [ebp-1c]
'0072727c    52                      push edx

' *** Reference to "__vbaObjSet"
'0072727d    ff15b4104000            call dword ptr [004010b4]
Set var_40 = 
'00727283    dd4634                  fld qword ptr [esi+34]
'00727286    833d00b0720000          cmp dword ptr [0072b000], 00
'0072728d    7508                    jne 727297

If (vbNullChar = 0) Then
'0072728f    dc35d8154000            fdiv qword ptr [004015d8]
'00727295    eb11                    jmp 7272a8
    
End If
'00727297    ff35dc154000            push dword ptr [004015dc]
'0072729d    ff35d8154000            push dword ptr [004015d8]

' *** Reference to "_adj_fdiv_m64"
'007272a3    e8dcffcdff              call 407284
'007272a8    8bf8                    mov edi, eax
'007272aa    8b1f                    mov ebx, dword ptr [edi]
'007272ac    83ec08                  sub esp, 08
'007272af    dfe0                    fnstsw ax
'007272b1    a80d                    test al, 0d
'007272b3    0f85e1000000            jne 72739a
'007272b9    dd1c24                  fstp qword ptr [esp]
'var_134 = (00)

' *** Reference to "__vbaStrR8"
'007272bc    ff1580114000            call dword ptr [00401180]
'007272c2    8bd0                    mov edx, eax
'007272c4    8d4de8                  lea ecx, dword ptr [ebp-18]

' *** Reference to "__vbaStrMove"
'007272c7    ff15d0124000            call dword ptr [004012d0]
'007272cd    50                      push eax
'007272ce    57                      push edi
'007272cf    ff93a4000000            call dword ptr [ebx+000000a4]
'007272d5    dbe2                    fnclex
'007272d7    85c0                    test ax, ax
'007272d9    7d12                    jge 7272ed
'007272db    68a4000000              push 000000a4
'007272e0    68d00d4300              push 00430dd0
'007272e5    57                      push edi
'007272e6    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'007272e7    ff1580104000            call dword ptr [00401080]
'007272ed    8d4de8                  lea ecx, dword ptr [ebp-18]

' *** Reference to "__vbaFreeStr"
'007272f0    ff150c134000            call dword ptr [0040130c]
'#Cleanup(var_41)

' *** Reference to "__vbaFreeObj"
'007272f6    8b3d08134000            mov edi, dword ptr [00401308]
'007272fc    8d4de4                  lea ecx, dword ptr [ebp-1c]
'007272ff    ffd7                    call edi
'#Cleanup(var_40)
'00727301    8b06                    mov eax, dword ptr [esi]
'00727303    56                      push esi
'00727304    ff900c030000            call dword ptr [eax+0000030c]
'0072730a    50                      push eax
'0072730b    8d4de4                  lea ecx, dword ptr [ebp-1c]
'0072730e    51                      push ecx

' *** Reference to "__vbaObjSet"
'0072730f    ff15b4104000            call dword ptr [004010b4]
Set var_40 = Nothing
'00727315    8bf0                    mov esi, eax
'00727317    8b16                    mov edx, dword ptr [esi]
'00727319    56                      push esi
'0072731a    ff9204020000            call dword ptr [edx+00000204]
'00727320    dbe2                    fnclex
'00727322    85c0                    test ax, ax
'00727324    7d12                    jge 727338
'00727326    6804020000              push 00000204
'0072732b    68d00d4300              push 00430dd0
'00727330    56                      push esi
'00727331    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00727332    ff1580104000            call dword ptr [00401080]
'00727338    8d4de4                  lea ecx, dword ptr [ebp-1c]
'0072733b    ffd7                    call edi
'#Cleanup(var_40)
'0072733d    c745fc00000000          mov dword ptr [ebp-04], 00000000
'00727344    9b                      fwait
'00727345    687b737200              push 0072737b
'0072734a    eb2e                    jmp 72737a
'0072734c    8d4de8                  lea ecx, dword ptr [ebp-18]

' *** Reference to "__vbaFreeStr"
'0072734f    ff150c134000            call dword ptr [0040130c]
'#Cleanup(var_41)
'00727355    8d4de4                  lea ecx, dword ptr [ebp-1c]

' *** Reference to "__vbaFreeObj"
'00727358    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_40)
'0072735e    8d45a4                  lea eax, dword ptr [ebp-5c]
'00727361    50                      push eax
'00727362    8d4db4                  lea ecx, dword ptr [ebp-4c]
'00727365    51                      push ecx
'00727366    8d55c4                  lea edx, dword ptr [ebp-3c]
'00727369    52                      push edx
'0072736a    8d45d4                  lea eax, dword ptr [ebp-2c]
'0072736d    50                      push eax
'0072736e    6a04                    push 04

' *** Reference to "__vbaFreeVarList"
'00727370    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_44, var_9, var_62, var_83)
'00727376    83c414                  add esp, 14
'00727379    c3                      ret
'0072737a    c3                      ret
'0072737b    8b4508                  mov eax, dword ptr [ebp+08]
'0072737e    8b08                    mov ecx, dword ptr [eax]
'00727380    50                      push eax
'00727381    ff5108                  call dword ptr [ecx+08]
'00727384    8b45fc                  mov eax, dword ptr [ebp-04]
'00727387    8b4dec                  mov ecx, dword ptr [ebp-14]
'0072738a    5f                      pop edi
'0072738b    5e                      pop esi
    'Reference to '__except_list'
'0072738c    64890d00000000          mov dword ptr fs:[00000000], ecx
'00727393    5b                      pop ebx
'00727394    8be5                    mov esp, ebp
'00727396    5d                      pop ebp
'00727397    c20400                  ret 0004


End Sub


'Event for btnVendre
Private Sub btnVendre_Click()
'00724d10    55                      push ebp
'00724d11    8bec                    mov ebp, esp
'00724d13    83ec14                  sub esp, 14

' *** Reference to "__vbaExceptHandler"
'00724d16    6866724000              push 00407266
'00724d1b    64a100000000            mov ax, word ptr fs:[00000000]
'00724d21    50                      push eax
    'Reference to '__except_list'
'00724d22    64892500000000          mov dword ptr fs:[00000000], esp
'00724d29    81ec88010000            sub esp, 00000188
'00724d2f    53                      push ebx
'00724d30    56                      push esi
'00724d31    57                      push edi
'00724d32    8965ec                  mov dword ptr [ebp-14], esp
'00724d35    c745f0a0714000          mov dword ptr [ebp-10], 004071a0
'00724d3c    8b7d08                  mov edi, dword ptr [ebp+08]
'00724d3f    8bc7                    mov eax, edi
'00724d41    83e001                  and eax, 01
'00724d44    8945f4                  mov dword ptr [ebp-0c], eax
'00724d47    83e7fe                  and edi, fffffffe
'00724d4a    897d08                  mov dword ptr [ebp+08], edi
'00724d4d    33f6                    xor esi, esi
'00724d4f    8975f8                  mov dword ptr [ebp-08], esi
'00724d52    8b0f                    mov ecx, dword ptr [edi]
'00724d54    57                      push edi
'00724d55    ff5104                  call dword ptr [ecx+04]
'00724d58    8975e0                  mov dword ptr [ebp-20], esi
'00724d5b    8975dc                  mov dword ptr [ebp-24], esi
'00724d5e    8975d8                  mov dword ptr [ebp-28], esi
'00724d61    8975d4                  mov dword ptr [ebp-2c], esi
'00724d64    8975d0                  mov dword ptr [ebp-30], esi
'00724d67    8975cc                  mov dword ptr [ebp-34], esi
'00724d6a    8975c8                  mov dword ptr [ebp-38], esi
'00724d6d    8975c4                  mov dword ptr [ebp-3c], esi
'00724d70    8975c0                  mov dword ptr [ebp-40], esi
'00724d73    8975bc                  mov dword ptr [ebp-44], esi
'00724d76    8975ac                  mov dword ptr [ebp-54], esi
'00724d79    89759c                  mov dword ptr [ebp-64], esi
'00724d7c    89758c                  mov dword ptr [ebp-74], esi
'00724d7f    89b57cffffff            mov dword ptr [ebp+ffffff7c], esi
'00724d85    89b56cffffff            mov dword ptr [ebp+ffffff6c], esi
'00724d8b    89b55cffffff            mov dword ptr [ebp+ffffff5c], esi
'00724d91    89b54cffffff            mov dword ptr [ebp+ffffff4c], esi
'00724d97    89b53cffffff            mov dword ptr [ebp+ffffff3c], esi
'00724d9d    89b52cffffff            mov dword ptr [ebp+ffffff2c], esi
'00724da3    89b51cffffff            mov dword ptr [ebp+ffffff1c], esi
'00724da9    89b50cffffff            mov dword ptr [ebp+ffffff0c], esi
'00724daf    89b5fcfeffff            mov dword ptr [ebp+fffffefc], esi
'00724db5    89b5ecfeffff            mov dword ptr [ebp+fffffeec], esi
'00724dbb    89b5dcfeffff            mov dword ptr [ebp+fffffedc], esi
'00724dc1    89b5bcfeffff            mov dword ptr [ebp+fffffebc], esi
'00724dc7    89b5a8feffff            mov dword ptr [ebp+fffffea8], esi
'00724dcd    66393528b07200          cmp word ptr [0072b028], si
'00724dd4    7508                    jne 724dde
'00724dd6    6a01                    push 01

' *** Reference to "__vbaOnError"
'00724dd8    ff15b0104000            call dword ptr [004010b0]
On Error Goto handler_0
'00724dde    8b07                    mov eax, dword ptr [edi]
'00724de0    57                      push edi
'00724de1    ff9004030000            call dword ptr [eax+00000304]
'00724de7    50                      push eax
'00724de8    8d4dc4                  lea ecx, dword ptr [ebp-3c]
'00724deb    51                      push ecx

' *** Reference to "__vbaObjSet"
'00724dec    8b1db4104000            mov ebx, dword ptr [004010b4]
'00724df2    ffd3                    call ebx
Set var_9 = Nothing
'00724df4    8bf0                    mov esi, eax
'00724df6    8b16                    mov edx, dword ptr [esi]
'00724df8    8d45e0                  lea eax, dword ptr [ebp-20]
'00724dfb    50                      push eax
'00724dfc    56                      push esi
'00724dfd    ff92a0000000            call dword ptr [edx+000000a0]
'00724e03    dbe2                    fnclex
'00724e05    85c0                    test ax, ax
'00724e07    7d12                    jge 724e1b
'00724e09    68a0000000              push 000000a0
'00724e0e    68d00d4300              push 00430dd0
'00724e13    56                      push esi
'00724e14    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00724e15    ff1580104000            call dword ptr [00401080]
'00724e1b    8b4de0                  mov ecx, dword ptr [ebp-20]
'00724e1e    51                      push ecx

' *** Reference to "rtcR8ValFromBstr"
'00724e1f    ff1510134000            call dword ptr [00401310]

' *** Reference to "__vbaFpR8"
'00724e25    ff15f8104000            call dword ptr [004010f8]
'00724e2b    dc1d68154000            fcomp qword ptr [00401568]
'00724e31    dfe0                    fnstsw ax
'00724e33    f6c440                  test ah, 40
'00724e36    0f8445030000            je 725181
'00724e3c    be01000000              mov esi, 00000001
'00724e41    e93d030000              jmp 725183

' *** Reference to "rtcErrObj"
'00724e46    8b1d6c124000            mov ebx, dword ptr [0040126c]
'00724e4c    ffd3                    call ebx
'00724e4e    50                      push eax
'00724e4f    8d55c4                  lea edx, dword ptr [ebp-3c]
'00724e52    52                      push edx

' *** Reference to "__vbaObjSet"
'00724e53    8b3db4104000            mov edi, dword ptr [004010b4]
'00724e59    ffd7                    call edi
Set var_9 = Err
'00724e5b    8bf0                    mov esi, eax
'00724e5d    8b06                    mov eax, dword ptr [esi]
'00724e5f    8d4de0                  lea ecx, dword ptr [ebp-20]
'00724e62    51                      push ecx
'00724e63    56                      push esi
'00724e64    ff502c                  call dword ptr [eax+2c]
var_3 = var_9.Description
'00724e67    dbe2                    fnclex
'00724e69    85c0                    test ax, ax
'00724e6b    7d0f                    jge 724e7c
'00724e6d    6a2c                    push 2c
'00724e6f    685c084300              push 0043085c
'00724e74    56                      push esi
'00724e75    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00724e76    ff1580104000            call dword ptr [00401080]
'00724e7c    ffd3                    call ebx
'00724e7e    50                      push eax
'00724e7f    8d55c0                  lea edx, dword ptr [ebp-40]
'00724e82    52                      push edx
'00724e83    ffd7                    call edi
Set var_5 = Err
'00724e85    8bf0                    mov esi, eax
'00724e87    8b06                    mov eax, dword ptr [esi]
'00724e89    8d8da8feffff            lea ecx, dword ptr [ebp+fffffea8]
'00724e8f    51                      push ecx
'00724e90    56                      push esi
'00724e91    ff501c                  call dword ptr [eax+1c]
var_519 = var_5.Number
'00724e94    dbe2                    fnclex
'00724e96    85c0                    test ax, ax
'00724e98    7d0f                    jge 724ea9
'00724e9a    6a1c                    push 1c
'00724e9c    685c084300              push 0043085c
'00724ea1    56                      push esi
'00724ea2    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00724ea3    ff1580104000            call dword ptr [00401080]
'00724ea9    b904000280              mov ecx, 80020004
'00724eae    894d94                  mov dword ptr [ebp-6c], ecx
'00724eb1    b80a000000              mov eax, 0000000a
'00724eb6    89458c                  mov dword ptr [ebp-74], eax
'00724eb9    894da4                  mov dword ptr [ebp-5c], ecx
'00724ebc    89459c                  mov dword ptr [ebp-64], eax
'00724ebf    c78524ffffff24b07200    mov dword ptr [ebp+ffffff24], 0072b024
'00724ec9    c7851cffffff08400000    mov dword ptr [ebp+ffffff1c], 00004008
'00724ed3    6814084300              push 00430814
'00724ed8    8b55e0                  mov edx, dword ptr [ebp-20]
'00724edb    52                      push edx

' *** Reference to "__vbaStrCat"
'00724edc    8b3570104000            mov esi, dword ptr [00401070]
'00724ee2    ffd6                    call esi
var_11 = ("L'erreur suivante s'est produite : ") & (var_3)
'00724ee4    8bd0                    mov edx, eax
'00724ee6    8d4ddc                  lea ecx, dword ptr [ebp-24]

' *** Reference to "__vbaStrMove"
'00724ee9    8b3dd0124000            mov edi, dword ptr [004012d0]
'00724eef    ffd7                    call edi
'00724ef1    50                      push eax
'00724ef2    6870084300              push 00430870
'00724ef7    ffd6                    call esi
var_76 = (var_11) & (vbCrLf)
'00724ef9    8bd0                    mov edx, eax
'00724efb    8d4dd8                  lea ecx, dword ptr [ebp-28]
'00724efe    ffd7                    call edi
'00724f00    50                      push eax
'00724f01    6870084300              push 00430870
'00724f06    ffd6                    call esi
var_12 = (var_76) & (vbCrLf)
'00724f08    8bd0                    mov edx, eax
'00724f0a    8d4dd4                  lea ecx, dword ptr [ebp-2c]
'00724f0d    ffd7                    call edi
'00724f0f    50                      push eax
'00724f10    6880084300              push 00430880
'00724f15    ffd6                    call esi
var_13 = (var_12) & (" numéro d'erreur (")
'00724f17    8bd0                    mov edx, eax
'00724f19    8d4dd0                  lea ecx, dword ptr [ebp-30]
'00724f1c    ffd7                    call edi
'00724f1e    50                      push eax
'00724f1f    8b85a8feffff            mov eax, dword ptr [ebp+fffffea8]
'00724f25    50                      push eax

' *** Reference to "__vbaStrI4"
'00724f26    ff1520104000            call dword ptr [00401020]
'00724f2c    8bd0                    mov edx, eax
'00724f2e    8d4dcc                  lea ecx, dword ptr [ebp-34]
'00724f31    ffd7                    call edi
'00724f33    50                      push eax
'00724f34    ffd6                    call esi
var_127 = (var_13) & (CStr(var_519))
'00724f36    8bd0                    mov edx, eax
'00724f38    8d4dc8                  lea ecx, dword ptr [ebp-38]
'00724f3b    ffd7                    call edi
'00724f3d    50                      push eax
'00724f3e    68ac084300              push 004308ac
'00724f43    ffd6                    call esi
var_15 = (var_127) & (" )")
'00724f45    8945b4                  mov dword ptr [ebp-4c], eax
'00724f48    bf08000000              mov edi, 00000008
'00724f4d    897dac                  mov dword ptr [ebp-54], edi
'00724f50    8d4d8c                  lea ecx, dword ptr [ebp-74]
'00724f53    51                      push ecx
'00724f54    8d559c                  lea edx, dword ptr [ebp-64]
'00724f57    52                      push edx
'00724f58    8d851cffffff            lea eax, dword ptr [ebp+ffffff1c]
'00724f5e    50                      push eax
'00724f5f    6a10                    push 10
'00724f61    8d4dac                  lea ecx, dword ptr [ebp-54]
'00724f64    51                      push ecx

' *** Reference to "rtcMsgBox"
'00724f65    ff15b8104000            call dword ptr [004010b8]
var_128 = MsgBox(var_15, 16, vbNullString)
'00724f6b    8d55c8                  lea edx, dword ptr [ebp-38]
'00724f6e    52                      push edx
'00724f6f    8d45cc                  lea eax, dword ptr [ebp-34]
'00724f72    50                      push eax
'00724f73    8d4dd0                  lea ecx, dword ptr [ebp-30]
'00724f76    51                      push ecx
'00724f77    8d55d4                  lea edx, dword ptr [ebp-2c]
'00724f7a    52                      push edx
'00724f7b    8d45d8                  lea eax, dword ptr [ebp-28]
'00724f7e    50                      push eax
'00724f7f    8d4ddc                  lea ecx, dword ptr [ebp-24]
'00724f82    51                      push ecx
'00724f83    8d55e0                  lea edx, dword ptr [ebp-20]
'00724f86    52                      push edx
'00724f87    6a07                    push 07

' *** Reference to "__vbaFreeStrList"
'00724f89    ff155c124000            call dword ptr [0040125c]
'#Cleanup( 0, -4508, -4512, -4516, -4520, -4524, -4528)
'00724f8f    8d45c0                  lea eax, dword ptr [ebp-40]
'00724f92    50                      push eax
'00724f93    8d4dc4                  lea ecx, dword ptr [ebp-3c]
'00724f96    51                      push ecx
'00724f97    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'00724f99    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_9, var_5)
'00724f9f    8d558c                  lea edx, dword ptr [ebp-74]
'00724fa2    52                      push edx
'00724fa3    8d459c                  lea eax, dword ptr [ebp-64]
'00724fa6    50                      push eax
'00724fa7    8d4dac                  lea ecx, dword ptr [ebp-54]
'00724faa    51                      push ecx
'00724fab    6a03                    push 03

' *** Reference to "__vbaFreeVarList"
'00724fad    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_50, var_51, var_52)
'00724fb3    83c43c                  add esp, 3c
'00724fb6    8d55ac                  lea edx, dword ptr [ebp-54]
'00724fb9    52                      push edx

' *** Reference to "rtcGetPresentDate"
'00724fba    ff15f4124000            call dword ptr [004012f4]
var_15 = Now()
'00724fc0    c78524ffffffb8084300    mov dword ptr [ebp+ffffff24], 004308b8
'00724fca    89bd1cffffff            mov dword ptr [ebp+ffffff1c], edi
'00724fd0    8d951cffffff            lea edx, dword ptr [ebp+ffffff1c]
'00724fd6    8d4d9c                  lea ecx, dword ptr [ebp-64]

' *** Reference to "__vbaVarDup"
'00724fd9    ff1598124000            call dword ptr [00401298]
var_51 = ("dd/MM/yyyy hh:mm:ss")
'00724fdf    6a01                    push 01
'00724fe1    6a01                    push 01
'00724fe3    8d459c                  lea eax, dword ptr [ebp-64]
'00724fe6    50                      push eax
'00724fe7    8d4dac                  lea ecx, dword ptr [ebp-54]
'00724fea    51                      push ecx
'00724feb    8d558c                  lea edx, dword ptr [ebp-74]
'00724fee    52                      push edx

' *** Reference to "rtcVarFromFormatVar"
'00724fef    ff1574104000            call dword ptr [00401074]
'00724ff5    ffd3                    call ebx
'00724ff7    50                      push eax
'00724ff8    8d45c4                  lea eax, dword ptr [ebp-3c]
'00724ffb    50                      push eax

' *** Reference to "__vbaObjSet"
'00724ffc    ff15b4104000            call dword ptr [004010b4]
Set var_9 = Err
'00725002    8bf0                    mov esi, eax
'00725004    8b0e                    mov ecx, dword ptr [esi]
'00725006    8d95a8feffff            lea edx, dword ptr [ebp+fffffea8]
'0072500c    52                      push edx
'0072500d    56                      push esi
'0072500e    ff511c                  call dword ptr [ecx+1c]
var_519 = var_9.Number
'00725011    dbe2                    fnclex
'00725013    85c0                    test ax, ax
'00725015    7d0f                    jge 725026
'00725017    6a1c                    push 1c
'00725019    685c084300              push 0043085c
'0072501e    56                      push esi
'0072501f    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00725020    ff1580104000            call dword ptr [00401080]
'00725026    ffd3                    call ebx
'00725028    50                      push eax
'00725029    8d45c0                  lea eax, dword ptr [ebp-40]
'0072502c    50                      push eax

' *** Reference to "__vbaObjSet"
'0072502d    ff15b4104000            call dword ptr [004010b4]
Set var_5 = Err
'00725033    8bf0                    mov esi, eax
'00725035    8b0e                    mov ecx, dword ptr [esi]
'00725037    8d55e0                  lea edx, dword ptr [ebp-20]
'0072503a    52                      push edx
'0072503b    56                      push esi
'0072503c    ff512c                  call dword ptr [ecx+2c]
var_3 = var_5.Description
'0072503f    dbe2                    fnclex
'00725041    85c0                    test ax, ax
'00725043    7d0f                    jge 725054
'00725045    6a2c                    push 2c
'00725047    685c084300              push 0043085c
'0072504c    56                      push esi
'0072504d    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0072504e    ff1580104000            call dword ptr [00401080]
'00725054    c78514ffffffe4084300    mov dword ptr [ebp+ffffff14], 004308e4
'0072505e    89bd0cffffff            mov dword ptr [ebp+ffffff0c], edi
'00725064    8b85a8feffff            mov eax, dword ptr [ebp+fffffea8]
'0072506a    898504ffffff            mov dword ptr [ebp+ffffff04], eax
'00725070    c785fcfeffff03000000    mov dword ptr [ebp+fffffefc], 00000003
'0072507a    c785f4feffff08094300    mov dword ptr [ebp+fffffef4], 00430908
'00725084    89bdecfeffff            mov dword ptr [ebp+fffffeec], edi
'0072508a    8b45e0                  mov eax, dword ptr [ebp-20]
'0072508d    c745e000000000          mov dword ptr [ebp-20], 00000000
'00725094    898554ffffff            mov dword ptr [ebp+ffffff54], eax
'0072509a    89bd4cffffff            mov dword ptr [ebp+ffffff4c], edi
'007250a0    c785e4feffff447a4500    mov dword ptr [ebp+fffffee4], 00457a44
'007250aa    89bddcfeffff            mov dword ptr [ebp+fffffedc], edi
'007250b0    8d4d8c                  lea ecx, dword ptr [ebp-74]
'007250b3    51                      push ecx
'007250b4    8d950cffffff            lea edx, dword ptr [ebp+ffffff0c]
'007250ba    52                      push edx
'007250bb    8d857cffffff            lea eax, dword ptr [ebp+ffffff7c]
'007250c1    50                      push eax

' *** Reference to "__vbaVarCat"
'007250c2    8b3508124000            mov esi, dword ptr [00401208]
'007250c8    ffd6                    call esi
'007250ca    50                      push eax
'007250cb    8d8dfcfeffff            lea ecx, dword ptr [ebp+fffffefc]
'007250d1    51                      push ecx
'007250d2    8d956cffffff            lea edx, dword ptr [ebp+ffffff6c]
'007250d8    52                      push edx
'007250d9    ffd6                    call esi
'007250db    50                      push eax
'007250dc    8d85ecfeffff            lea eax, dword ptr [ebp+fffffeec]
'007250e2    50                      push eax
'007250e3    8d8d5cffffff            lea ecx, dword ptr [ebp+ffffff5c]
'007250e9    51                      push ecx
'007250ea    ffd6                    call esi
'007250ec    50                      push eax
'007250ed    8d954cffffff            lea edx, dword ptr [ebp+ffffff4c]
'007250f3    52                      push edx
'007250f4    8d853cffffff            lea eax, dword ptr [ebp+ffffff3c]
'007250fa    50                      push eax
'007250fb    ffd6                    call esi
'007250fd    50                      push eax
'007250fe    8d8ddcfeffff            lea ecx, dword ptr [ebp+fffffedc]
'00725104    51                      push ecx
'00725105    8d952cffffff            lea edx, dword ptr [ebp+ffffff2c]
'0072510b    52                      push edx
'0072510c    ffd6                    call esi
'0072510e    50                      push eax
'0072510f    33c0                    xor eax, eax
'00725111    66a12ab07200            mov eax, dword ptr [0072b02a]
'00725117    50                      push eax
'00725118    6884094300              push 00430984

' *** Reference to "__vbaPrintFile"
'0072511d    ff15b8114000            call dword ptr [004011b8]
Print #0, 
'00725123    8d4dc0                  lea ecx, dword ptr [ebp-40]
'00725126    51                      push ecx
'00725127    8d55c4                  lea edx, dword ptr [ebp-3c]
'0072512a    52                      push edx
'0072512b    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'0072512d    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_9, var_5)
'00725133    8d852cffffff            lea eax, dword ptr [ebp+ffffff2c]
'00725139    50                      push eax
'0072513a    8d8d3cffffff            lea ecx, dword ptr [ebp+ffffff3c]
'00725140    51                      push ecx
'00725141    8d954cffffff            lea edx, dword ptr [ebp+ffffff4c]
'00725147    52                      push edx
'00725148    8d855cffffff            lea eax, dword ptr [ebp+ffffff5c]
'0072514e    50                      push eax
'0072514f    8d8d6cffffff            lea ecx, dword ptr [ebp+ffffff6c]
'00725155    51                      push ecx
'00725156    8d957cffffff            lea edx, dword ptr [ebp+ffffff7c]
'0072515c    52                      push edx
'0072515d    8d458c                  lea eax, dword ptr [ebp-74]
'00725160    50                      push eax
'00725161    8d4d9c                  lea ecx, dword ptr [ebp-64]
'00725164    51                      push ecx
'00725165    8d55ac                  lea edx, dword ptr [ebp-54]
'00725168    52                      push edx
'00725169    6a09                    push 09

' *** Reference to "__vbaFreeVarList"
'0072516b    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_50, var_51, var_52, var_59, var_94, var_88, var_89, var_90, var_93)
'00725171    83c440                  add esp, 40
'00725174    6a00                    push 00

' *** Reference to "__vbaResume"
'00725176    ff1568104000            call dword ptr [00401068]
'0072517c    e9a7060000              jmp 725828
Resume handler_725828
'00725181    33f6                    xor esi, esi
'00725183    8d4de0                  lea ecx, dword ptr [ebp-20]

' *** Reference to "__vbaFreeStr"
'00725186    ff150c134000            call dword ptr [0040130c]
'#Cleanup(var_3)
'0072518c    8d4dc4                  lea ecx, dword ptr [ebp-3c]

' *** Reference to "__vbaFreeObj"
'0072518f    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_9)
'00725195    f7de                    neg esi
'00725197    6685f6                  test esi, esi
'0072519a    0f84a6000000            je 725246

If ((0# = Val(vbNullString))) Then
'007251a0    a184b07200              mov ax, word ptr [0072b084]
'007251a5    85c0                    test ax, ax
'007251a7    7515                    jne 7251be
'007251a9    6884b07200              push 0072b084
'007251ae    68f8894100              push 004189f8

' *** Reference to "__vbaNew2"
'007251b3    ff152c124000            call dword ptr [0040122c]
    Dim var_28 As New frmGestPerso
'007251b9    a184b07200              mov ax, word ptr [0072b084]
'007251be    6a00                    push 00
'007251c0    6881000000              push 00000081
'007251c5    8b10                    mov edx, dword ptr [eax]
'007251c7    50                      push eax
'007251c8    ff9240040000            call dword ptr [edx+00000440]
'007251ce    50                      push eax
'007251cf    8d45c4                  lea eax, dword ptr [ebp-3c]
'007251d2    50                      push eax
'007251d3    ffd3                    call ebx
'007251d5    50                      push eax

' *** Reference to "__vbaLateIdCall"
'007251d6    ff1538104000            call dword ptr [00401038]
    Call ErrObject.Member_0xFFFFFEC4()
'007251dc    83c40c                  add esp, 0c
'007251df    8d4dc4                  lea ecx, dword ptr [ebp-3c]

' *** Reference to "__vbaFreeObj"
'007251e2    ff1508134000            call dword ptr [00401308]
    '#Cleanup(var_9)
'007251e8    a124be7200              mov ax, word ptr [0072be24]
'007251ed    85c0                    test ax, ax
'007251ef    7510                    jne 725201
'007251f1    6824be7200              push 0072be24
'007251f6    6870004300              push 00430070

' *** Reference to "__vbaNew2"
'007251fb    ff152c124000            call dword ptr [0040122c]
    Dim var_2 As New Global
'00725201    8b3524be7200            mov esi, dword ptr [0072be24]
'00725207    8b16                    mov edx, dword ptr [esi]
'00725209    57                      push edi
'0072520a    8d4dc4                  lea ecx, dword ptr [ebp-3c]
'0072520d    51                      push ecx
'0072520e    899570feffff            mov dword ptr [ebp+fffffe70], edx

' *** Reference to "__vbaObjSetAddref"
'00725214    ff15c8104000            call dword ptr [004010c8]
    Set var_9 = 8
'0072521a    50                      push eax
'0072521b    56                      push esi
'0072521c    8b9570feffff            mov edx, dword ptr [ebp+fffffe70]
'00725222    ff5210                  call dword ptr [edx+10]
    Call var_2.Unload(var_9)
'00725225    dbe2                    fnclex
'00725227    85c0                    test ax, ax
'00725229    7d0f                    jge 72523a
'0072522b    6a10                    push 10
'0072522d    6860004300              push 00430060
'00725232    56                      push esi
'00725233    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00725234    ff1580104000            call dword ptr [00401080]
'0072523a    8d4dc4                  lea ecx, dword ptr [ebp-3c]

' *** Reference to "__vbaFreeObj"
'0072523d    ff1508134000            call dword ptr [00401308]
    '#Cleanup(var_9)
'00725243    8b7d08                  mov edi, dword ptr [ebp+08]
    
End If
'00725246    8b07                    mov eax, dword ptr [edi]
'00725248    57                      push edi
'00725249    ff9008030000            call dword ptr [eax+00000308]
'0072524f    50                      push eax
'00725250    8d4dc4                  lea ecx, dword ptr [ebp-3c]
'00725253    51                      push ecx
'00725254    ffd3                    call ebx
'00725256    8bf0                    mov esi, eax
'00725258    8b16                    mov edx, dword ptr [esi]
'0072525a    8d45e0                  lea eax, dword ptr [ebp-20]
'0072525d    50                      push eax
'0072525e    56                      push esi
'0072525f    ff92a0000000            call dword ptr [edx+000000a0]
'00725265    dbe2                    fnclex
'00725267    85c0                    test ax, ax
'00725269    7d12                    jge 72527d
'0072526b    68a0000000              push 000000a0
'00725270    68d00d4300              push 00430dd0
'00725275    56                      push esi
'00725276    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00725277    ff1580104000            call dword ptr [00401080]
'0072527d    8b4de0                  mov ecx, dword ptr [ebp-20]
'00725280    51                      push ecx

' *** Reference to "rtcR8ValFromBstr"
'00725281    ff1510134000            call dword ptr [00401310]

' *** Reference to "__vbaFpR8"
'00725287    ff15f8104000            call dword ptr [004010f8]
'0072528d    0fbf573c                movsx edx, word ptr [edi+3c]
'00725291    89956cfeffff            mov dword ptr [ebp+fffffe6c], edx
'00725297    db856cfeffff            fild dword ptr [ebp+fffffe6c]
'0072529d    dd9d64feffff            fstp qword ptr [ebp+fffffe64]
'var_122 = (00)
'007252a3    dc9d64feffff            fcomp qword ptr [ebp+fffffe64]
'007252a9    dfe0                    fnstsw ax
'007252ab    f6c440                  test ah, 40
'007252ae    7407                    je 7252b7
'007252b0    be01000000              mov esi, 00000001
'007252b5    eb02                    jmp 7252b9
'007252b7    33f6                    xor esi, esi
'007252b9    8d4de0                  lea ecx, dword ptr [ebp-20]

' *** Reference to "__vbaFreeStr"
'007252bc    ff150c134000            call dword ptr [0040130c]
'#Cleanup(var_3)
'007252c2    8d4dc4                  lea ecx, dword ptr [ebp-3c]

' *** Reference to "__vbaFreeObj"
'007252c5    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_9)
'007252cb    f7de                    neg esi
'007252cd    6685f6                  test esi, esi
'007252d0    8b07                    mov eax, dword ptr [edi]
'007252d2    57                      push edi
'007252d3    0f84ff010000            je 7254d8

If (((0) = Val(vbNullString))) Then
'007252d9    ff9004030000            call dword ptr [eax+00000304]
'007252df    50                      push eax
'007252e0    8d4dc0                  lea ecx, dword ptr [ebp-40]
'007252e3    51                      push ecx
'007252e4    ffd3                    call ebx
'007252e6    8bf0                    mov esi, eax
'007252e8    8b16                    mov edx, dword ptr [esi]
'007252ea    8d45dc                  lea eax, dword ptr [ebp-24]
'007252ed    50                      push eax
'007252ee    56                      push esi
'007252ef    ff92a0000000            call dword ptr [edx+000000a0]
'007252f5    dbe2                    fnclex
'007252f7    85c0                    test ax, ax
'007252f9    7d12                    jge 72530d
'007252fb    68a0000000              push 000000a0
'00725300    68d00d4300              push 00430dd0
'00725305    56                      push esi
'00725306    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00725307    ff1580104000            call dword ptr [00401080]
'0072530d    8b4ddc                  mov ecx, dword ptr [ebp-24]
'00725310    51                      push ecx

' *** Reference to "rtcR8ValFromBstr"
'00725311    ff1510134000            call dword ptr [00401310]
'00725317    dd9da0feffff            fstp qword ptr [ebp+fffffea0]
    'var_527 = (00)
'0072531d    a184b07200              mov ax, word ptr [0072b084]
'00725322    85c0                    test ax, ax
'00725324    7515                    jne 72533b
'00725326    6884b07200              push 0072b084
'0072532b    68f8894100              push 004189f8

' *** Reference to "__vbaNew2"
'00725330    ff152c124000            call dword ptr [0040122c]
    Set var_28 = New frmGestPerso
'00725336    a184b07200              mov ax, word ptr [0072b084]
'0072533b    8b10                    mov edx, dword ptr [eax]
'0072533d    50                      push eax
'0072533e    ff9228030000            call dword ptr [edx+00000328]
'00725344    50                      push eax
'00725345    8d45bc                  lea eax, dword ptr [ebp-44]
'00725348    50                      push eax
'00725349    ffd3                    call ebx
'0072534b    89858cfeffff            mov dword ptr [ebp+fffffe8c], eax
'00725351    a184b07200              mov ax, word ptr [0072b084]
'00725356    85c0                    test ax, ax
'00725358    7515                    jne 72536f
'0072535a    6884b07200              push 0072b084
'0072535f    68f8894100              push 004189f8

' *** Reference to "__vbaNew2"
'00725364    ff152c124000            call dword ptr [0040122c]
    Set var_28 = New frmGestPerso
'0072536a    a184b07200              mov ax, word ptr [0072b084]
'0072536f    8b08                    mov ecx, dword ptr [eax]
'00725371    50                      push eax
'00725372    ff9128030000            call dword ptr [ecx+00000328]
'00725378    50                      push eax
'00725379    8d55c4                  lea edx, dword ptr [ebp-3c]
'0072537c    52                      push edx
'0072537d    ffd3                    call ebx
'0072537f    8bf0                    mov esi, eax
'00725381    8b06                    mov eax, dword ptr [esi]
'00725383    8d4de0                  lea ecx, dword ptr [ebp-20]
'00725386    51                      push ecx
'00725387    56                      push esi
'00725388    ff90a0000000            call dword ptr [eax+000000a0]
'0072538e    dbe2                    fnclex
'00725390    85c0                    test ax, ax
'00725392    7d12                    jge 7253a6
'00725394    68a0000000              push 000000a0
'00725399    68d00d4300              push 00430dd0
'0072539e    56                      push esi
'0072539f    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'007253a0    ff1580104000            call dword ptr [00401080]
'007253a6    8b958cfeffff            mov edx, dword ptr [ebp+fffffe8c]
'007253ac    8b32                    mov esi, dword ptr [edx]
'007253ae    8b45e0                  mov eax, dword ptr [ebp-20]
'007253b1    50                      push eax

' *** Reference to "rtcR8ValFromBstr"
'007253b2    ff1510134000            call dword ptr [00401310]
'007253b8    dc85a0feffff            fadd qword ptr [ebp+fffffea0]
'007253be    dfe0                    fnstsw ax
'007253c0    a80d                    test al, 0d
'007253c2    0f8508050000            jne 7258d0
'007253c8    83ec08                  sub esp, 08
'007253cb    dd1c24                  fstp qword ptr [esp]
    'var_538 = (00)

' *** Reference to "__vbaStrR8"
'007253ce    ff1580114000            call dword ptr [00401180]
'007253d4    8bd0                    mov edx, eax
'007253d6    8d4dd8                  lea ecx, dword ptr [ebp-28]

' *** Reference to "__vbaStrMove"
'007253d9    ff15d0124000            call dword ptr [004012d0]
'007253df    50                      push eax
'007253e0    8bce                    mov ecx, esi
'007253e2    8bb58cfeffff            mov esi, dword ptr [ebp+fffffe8c]
'007253e8    56                      push esi
'007253e9    ff91a4000000            call dword ptr [ecx+000000a4]
'007253ef    dbe2                    fnclex
'007253f1    85c0                    test ax, ax
'007253f3    7d12                    jge 725407
'007253f5    68a4000000              push 000000a4
'007253fa    68d00d4300              push 00430dd0
'007253ff    56                      push esi
'00725400    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00725401    ff1580104000            call dword ptr [00401080]
'00725407    8d55d8                  lea edx, dword ptr [ebp-28]
'0072540a    52                      push edx
'0072540b    8d45dc                  lea eax, dword ptr [ebp-24]
'0072540e    50                      push eax
'0072540f    8d4de0                  lea ecx, dword ptr [ebp-20]
'00725412    51                      push ecx
'00725413    6a03                    push 03

' *** Reference to "__vbaFreeStrList"
'00725415    ff155c124000            call dword ptr [0040125c]
    '#Cleanup( , -4508, -4680)
'0072541b    8d55bc                  lea edx, dword ptr [ebp-44]
'0072541e    52                      push edx
'0072541f    8d45c0                  lea eax, dword ptr [ebp-40]
'00725422    50                      push eax
'00725423    8d4dc4                  lea ecx, dword ptr [ebp-3c]
'00725426    51                      push ecx
'00725427    6a03                    push 03

' *** Reference to "__vbaFreeObjList"
'00725429    ff154c104000            call dword ptr [0040104c]
    '#Cleanup( var_9, var_5, var_58)
'0072542f    83c420                  add esp, 20
'00725432    a184b07200              mov ax, word ptr [0072b084]
'00725437    85c0                    test ax, ax
'00725439    7515                    jne 725450
'0072543b    6884b07200              push 0072b084
'00725440    68f8894100              push 004189f8

' *** Reference to "__vbaNew2"
'00725445    ff152c124000            call dword ptr [0040122c]
    Set var_28 = New frmGestPerso
'0072544b    a184b07200              mov ax, word ptr [0072b084]
'00725450    6a00                    push 00
'00725452    6881000000              push 00000081
'00725457    8b10                    mov edx, dword ptr [eax]
'00725459    50                      push eax
'0072545a    ff9240040000            call dword ptr [edx+00000440]
'00725460    50                      push eax
'00725461    8d45c4                  lea eax, dword ptr [ebp-3c]
'00725464    50                      push eax
'00725465    ffd3                    call ebx
'00725467    50                      push eax

' *** Reference to "__vbaLateIdCall"
'00725468    ff1538104000            call dword ptr [00401038]
    Call ErrObject.Member_0xFFFFFEC4()
'0072546e    83c40c                  add esp, 0c
'00725471    8d4dc4                  lea ecx, dword ptr [ebp-3c]

' *** Reference to "__vbaFreeObj"
'00725474    8b1d08134000            mov ebx, dword ptr [00401308]
'0072547a    ffd3                    call ebx
    '#Cleanup(var_9)
'0072547c    a124be7200              mov ax, word ptr [0072be24]
'00725481    85c0                    test ax, ax
'00725483    7510                    jne 725495
'00725485    6824be7200              push 0072be24
'0072548a    6870004300              push 00430070

' *** Reference to "__vbaNew2"
'0072548f    ff152c124000            call dword ptr [0040122c]
    Set var_2 = New Global
'00725495    8b3524be7200            mov esi, dword ptr [0072be24]
'0072549b    8b16                    mov edx, dword ptr [esi]
'0072549d    57                      push edi
'0072549e    8d4dc4                  lea ecx, dword ptr [ebp-3c]
'007254a1    51                      push ecx
'007254a2    89955cfeffff            mov dword ptr [ebp+fffffe5c], edx

' *** Reference to "__vbaObjSetAddref"
'007254a8    ff15c8104000            call dword ptr [004010c8]
    Set var_9 = 8
'007254ae    50                      push eax
'007254af    56                      push esi
'007254b0    8b955cfeffff            mov edx, dword ptr [ebp+fffffe5c]
'007254b6    ff5210                  call dword ptr [edx+10]
    Call var_2.Unload(var_9)
'007254b9    dbe2                    fnclex
'007254bb    85c0                    test ax, ax
'007254bd    7d0f                    jge 7254ce
'007254bf    6a10                    push 10
'007254c1    6860004300              push 00430060
'007254c6    56                      push esi
'007254c7    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'007254c8    ff1580104000            call dword ptr [00401080]
'007254ce    8d4dc4                  lea ecx, dword ptr [ebp-3c]
'007254d1    ffd3                    call ebx
    '#Cleanup(var_9)
'007254d3    e950030000              jmp 725828
    
End If
'007254d8    ff9008030000            call dword ptr [eax+00000308]
'007254de    50                      push eax
'007254df    8d4dc4                  lea ecx, dword ptr [ebp-3c]
'007254e2    51                      push ecx
'007254e3    ffd3                    call ebx
'007254e5    8bf0                    mov esi, eax
'007254e7    8b16                    mov edx, dword ptr [esi]
'007254e9    8d45dc                  lea eax, dword ptr [ebp-24]
'007254ec    50                      push eax
'007254ed    56                      push esi
'007254ee    ff92a0000000            call dword ptr [edx+000000a0]
'007254f4    dbe2                    fnclex
'007254f6    85c0                    test ax, ax
'007254f8    7d12                    jge 72550c
'007254fa    68a0000000              push 000000a0
'007254ff    68d00d4300              push 00430dd0
'00725504    56                      push esi
'00725505    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00725506    ff1580104000            call dword ptr [00401080]
'0072550c    8b4ddc                  mov ecx, dword ptr [ebp-24]
'0072550f    51                      push ecx

' *** Reference to "rtcR8ValFromBstr"
'00725510    ff1510134000            call dword ptr [00401310]
'00725516    dd9da0feffff            fstp qword ptr [ebp+fffffea0]
'var_527 = (00)
'0072551c    c78524ffffff00000000    mov dword ptr [ebp+ffffff24], 00000000
'00725526    c7851cffffff03000000    mov dword ptr [ebp+ffffff1c], 00000003
'00725530    c78504ffffff04000280    mov dword ptr [ebp+ffffff04], 80020004
'0072553a    c785fcfeffff0a000000    mov dword ptr [ebp+fffffefc], 0000000a
'00725544    be05000000              mov esi, 00000005
'00725549    89b5e4feffff            mov dword ptr [ebp+fffffee4], esi
'0072554f    c785dcfeffff02000000    mov dword ptr [ebp+fffffedc], 00000002
'00725559    33d2                    xor edx, edx
var_num4 = Empty
'0072555b    668b573c                mov dx, word ptr [edi+3c]
'0072555f    52                      push edx

' *** Reference to "__vbaStrI2"
'00725560    ff1510104000            call dword ptr [00401010]
'00725566    8bd0                    mov edx, eax
'00725568    8d4de0                  lea ecx, dword ptr [ebp-20]

' *** Reference to "__vbaStrMove"
'0072556b    ff15d0124000            call dword ptr [004012d0]
'00725571    50                      push eax

' *** Reference to "rtcR8ValFromBstr"
'00725572    ff1510134000            call dword ptr [00401310]
'00725578    dca5a0feffff            fsub qword ptr [ebp+fffffea0]
'0072557e    dd9dc4feffff            fstp qword ptr [ebp+fffffec4]
'var_344 = (00)
'00725584    dfe0                    fnstsw ax
'00725586    a80d                    test al, 0d
'00725588    0f8542030000            jne 7258d0
'0072558e    a184b07200              mov ax, word ptr [0072b084]
'00725593    85c0                    test ax, ax
'00725595    7515                    jne 7255ac
'00725597    6884b07200              push 0072b084
'0072559c    68f8894100              push 004189f8

' *** Reference to "__vbaNew2"
'007255a1    ff152c124000            call dword ptr [0040122c]
Set var_28 = New frmGestPerso
'007255a7    a184b07200              mov ax, word ptr [0072b084]
'007255ac    83ec10                  sub esp, 10
'007255af    8bcc                    mov ecx, esp
'007255b1    8b951cffffff            mov edx, dword ptr [ebp+ffffff1c]
'007255b7    8911                    mov dword ptr [ecx], edx
'007255b9    8b9520ffffff            mov edx, dword ptr [ebp+ffffff20]
'007255bf    895104                  mov dword ptr [ecx+04], edx
'007255c2    8b9524ffffff            mov edx, dword ptr [ebp+ffffff24]
'007255c8    895108                  mov dword ptr [ecx+08], edx
'007255cb    8b9528ffffff            mov edx, dword ptr [ebp+ffffff28]
'007255d1    89510c                  mov dword ptr [ecx+0c], edx
'007255d4    83ec10                  sub esp, 10
'007255d7    8bcc                    mov ecx, esp
'007255d9    8b95fcfeffff            mov edx, dword ptr [ebp+fffffefc]
'007255df    8911                    mov dword ptr [ecx], edx
'007255e1    8b9500ffffff            mov edx, dword ptr [ebp+ffffff00]
'007255e7    895104                  mov dword ptr [ecx+04], edx
'007255ea    8b9504ffffff            mov edx, dword ptr [ebp+ffffff04]
'007255f0    895108                  mov dword ptr [ecx+08], edx
'007255f3    8b9508ffffff            mov edx, dword ptr [ebp+ffffff08]
'007255f9    89510c                  mov dword ptr [ecx+0c], edx
'007255fc    83ec10                  sub esp, 10
'007255ff    8bcc                    mov ecx, esp
'00725601    8b95dcfeffff            mov edx, dword ptr [ebp+fffffedc]
'00725607    8911                    mov dword ptr [ecx], edx
'00725609    8b95e0feffff            mov edx, dword ptr [ebp+fffffee0]
'0072560f    895104                  mov dword ptr [ecx+04], edx
'00725612    8b95e4feffff            mov edx, dword ptr [ebp+fffffee4]
'00725618    895108                  mov dword ptr [ecx+08], edx
'0072561b    8b95e8feffff            mov edx, dword ptr [ebp+fffffee8]
'00725621    89510c                  mov dword ptr [ecx+0c], edx
'00725624    83ec10                  sub esp, 10
'00725627    8bcc                    mov ecx, esp
'00725629    8931                    mov dword ptr [ecx], esi
'0072562b    8b95c0feffff            mov edx, dword ptr [ebp+fffffec0]
'00725631    895104                  mov dword ptr [ecx+04], edx
'00725634    8b95c4feffff            mov edx, dword ptr [ebp+fffffec4]
'0072563a    895108                  mov dword ptr [ecx+08], edx
'0072563d    8b95c8feffff            mov edx, dword ptr [ebp+fffffec8]
'00725643    89510c                  mov dword ptr [ecx+0c], edx
'00725646    6a03                    push 03
'00725648    689e000000              push 0000009e
'0072564d    8b08                    mov ecx, dword ptr [eax]
'0072564f    50                      push eax
'00725650    ff9140040000            call dword ptr [ecx+00000440]
'00725656    50                      push eax
'00725657    8d55c0                  lea edx, dword ptr [ebp-40]
'0072565a    52                      push edx
'0072565b    ffd3                    call ebx
'0072565d    50                      push eax

' *** Reference to "__vbaLateIdCallSt"
'0072565e    ff159c114000            call dword ptr [0040119c]
'00725664    8d45dc                  lea eax, dword ptr [ebp-24]
'00725667    50                      push eax
'00725668    8d4de0                  lea ecx, dword ptr [ebp-20]
'0072566b    51                      push ecx
'0072566c    6a02                    push 02

' *** Reference to "__vbaFreeStrList"
'0072566e    ff155c124000            call dword ptr [0040125c]
'#Cleanup( -4700, -4508)
'00725674    83c458                  add esp, 58
'00725677    8d55c0                  lea edx, dword ptr [ebp-40]
'0072567a    52                      push edx
'0072567b    8d45c4                  lea eax, dword ptr [ebp-3c]
'0072567e    50                      push eax
'0072567f    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'00725681    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_9, var_5)
'00725687    83c40c                  add esp, 0c
'0072568a    8b0f                    mov ecx, dword ptr [edi]
'0072568c    57                      push edi
'0072568d    ff9104030000            call dword ptr [ecx+00000304]
'00725693    50                      push eax
'00725694    8d55c0                  lea edx, dword ptr [ebp-40]
'00725697    52                      push edx
'00725698    ffd3                    call ebx
'0072569a    8bf0                    mov esi, eax
'0072569c    8b06                    mov eax, dword ptr [esi]
'0072569e    8d4ddc                  lea ecx, dword ptr [ebp-24]
'007256a1    51                      push ecx
'007256a2    56                      push esi
'007256a3    ff90a0000000            call dword ptr [eax+000000a0]
'007256a9    dbe2                    fnclex
'007256ab    85c0                    test ax, ax
'007256ad    7d12                    jge 7256c1
'007256af    68a0000000              push 000000a0
'007256b4    68d00d4300              push 00430dd0
'007256b9    56                      push esi
'007256ba    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'007256bb    ff1580104000            call dword ptr [00401080]
'007256c1    8b55dc                  mov edx, dword ptr [ebp-24]
'007256c4    52                      push edx

' *** Reference to "rtcR8ValFromBstr"
'007256c5    ff1510134000            call dword ptr [00401310]
'007256cb    dd9da0feffff            fstp qword ptr [ebp+fffffea0]
'var_527 = (00)
'007256d1    a184b07200              mov ax, word ptr [0072b084]
'007256d6    85c0                    test ax, ax
'007256d8    7515                    jne 7256ef
'007256da    6884b07200              push 0072b084
'007256df    68f8894100              push 004189f8

' *** Reference to "__vbaNew2"
'007256e4    ff152c124000            call dword ptr [0040122c]
Set var_28 = New frmGestPerso
'007256ea    a184b07200              mov ax, word ptr [0072b084]
'007256ef    8b08                    mov ecx, dword ptr [eax]
'007256f1    50                      push eax
'007256f2    ff9128030000            call dword ptr [ecx+00000328]
'007256f8    50                      push eax
'007256f9    8d55bc                  lea edx, dword ptr [ebp-44]
'007256fc    52                      push edx
'007256fd    ffd3                    call ebx
'007256ff    8bf8                    mov edi, eax
'00725701    a184b07200              mov ax, word ptr [0072b084]
'00725706    85c0                    test ax, ax
'00725708    7515                    jne 72571f
'0072570a    6884b07200              push 0072b084
'0072570f    68f8894100              push 004189f8

' *** Reference to "__vbaNew2"
'00725714    ff152c124000            call dword ptr [0040122c]
Set var_28 = New frmGestPerso
'0072571a    a184b07200              mov ax, word ptr [0072b084]
'0072571f    8b08                    mov ecx, dword ptr [eax]
'00725721    50                      push eax
'00725722    ff9128030000            call dword ptr [ecx+00000328]
'00725728    50                      push eax
'00725729    8d55c4                  lea edx, dword ptr [ebp-3c]
'0072572c    52                      push edx
'0072572d    ffd3                    call ebx
'0072572f    8bf0                    mov esi, eax
'00725731    8b06                    mov eax, dword ptr [esi]
'00725733    8d4de0                  lea ecx, dword ptr [ebp-20]
'00725736    51                      push ecx
'00725737    56                      push esi
'00725738    ff90a0000000            call dword ptr [eax+000000a0]
'0072573e    dbe2                    fnclex
'00725740    85c0                    test ax, ax
'00725742    7d16                    jge 72575a
'00725744    68a0000000              push 000000a0
'00725749    68d00d4300              push 00430dd0
'0072574e    56                      push esi
'0072574f    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00725750    8b1d80104000            mov ebx, dword ptr [00401080]
'00725756    ffd3                    call ebx
'00725758    eb06                    jmp 725760

' *** Reference to "__vbaHresultCheckObj"
'0072575a    8b1d80104000            mov ebx, dword ptr [00401080]
'00725760    8b37                    mov esi, dword ptr [edi]
'00725762    8b55e0                  mov edx, dword ptr [ebp-20]
'00725765    52                      push edx

' *** Reference to "rtcR8ValFromBstr"
'00725766    ff1510134000            call dword ptr [00401310]
'0072576c    dc85a0feffff            fadd qword ptr [ebp+fffffea0]
'00725772    dfe0                    fnstsw ax
'00725774    a80d                    test al, 0d
'00725776    0f8554010000            jne 7258d0
'0072577c    83ec08                  sub esp, 08
'0072577f    dd1c24                  fstp qword ptr [esp]
'var_537 = (00)

' *** Reference to "__vbaStrR8"
'00725782    ff1580114000            call dword ptr [00401180]
'00725788    8bd0                    mov edx, eax
'0072578a    8d4dd8                  lea ecx, dword ptr [ebp-28]

' *** Reference to "__vbaStrMove"
'0072578d    ff15d0124000            call dword ptr [004012d0]
'00725793    50                      push eax
'00725794    57                      push edi
'00725795    ff96a4000000            call dword ptr [esi+000000a4]
'0072579b    dbe2                    fnclex
'0072579d    85c0                    test ax, ax
'0072579f    7d0e                    jge 7257af
'007257a1    68a4000000              push 000000a4
'007257a6    68d00d4300              push 00430dd0
'007257ab    57                      push edi
'007257ac    50                      push eax
'007257ad    ffd3                    call ebx
'007257af    8d45d8                  lea eax, dword ptr [ebp-28]
'007257b2    50                      push eax
'007257b3    8d4ddc                  lea ecx, dword ptr [ebp-24]
'007257b6    51                      push ecx
'007257b7    8d55e0                  lea edx, dword ptr [ebp-20]
'007257ba    52                      push edx
'007257bb    6a03                    push 03

' *** Reference to "__vbaFreeStrList"
'007257bd    ff155c124000            call dword ptr [0040125c]
'#Cleanup( -4700, -4508, -4732)
'007257c3    8d45bc                  lea eax, dword ptr [ebp-44]
'007257c6    50                      push eax
'007257c7    8d4dc0                  lea ecx, dword ptr [ebp-40]
'007257ca    51                      push ecx
'007257cb    8d55c4                  lea edx, dword ptr [ebp-3c]
'007257ce    52                      push edx
'007257cf    6a03                    push 03

' *** Reference to "__vbaFreeObjList"
'007257d1    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_9, var_5, var_58)
'007257d7    83c420                  add esp, 20
'007257da    a124be7200              mov ax, word ptr [0072be24]
'007257df    85c0                    test ax, ax
'007257e1    7510                    jne 7257f3
'007257e3    6824be7200              push 0072be24
'007257e8    6870004300              push 00430070

' *** Reference to "__vbaNew2"
'007257ed    ff152c124000            call dword ptr [0040122c]
Set var_2 = New Global
'007257f3    8b3524be7200            mov esi, dword ptr [0072be24]
'007257f9    8b3e                    mov edi, dword ptr [esi]
'007257fb    8b4508                  mov eax, dword ptr [ebp+08]
'007257fe    50                      push eax
'007257ff    8d4dc4                  lea ecx, dword ptr [ebp-3c]
'00725802    51                      push ecx

' *** Reference to "__vbaObjSetAddref"
'00725803    ff15c8104000            call dword ptr [004010c8]
Set var_9 = Me
'00725809    50                      push eax
'0072580a    56                      push esi
'0072580b    ff5710                  call dword ptr [edi+10]
Call var_2.Unload(var_9)
'0072580e    dbe2                    fnclex
'00725810    85c0                    test ax, ax
'00725812    7d0b                    jge 72581f
'00725814    6a10                    push 10
'00725816    6860004300              push 00430060
'0072581b    56                      push esi
'0072581c    50                      push eax
'0072581d    ffd3                    call ebx
'0072581f    8d4dc4                  lea ecx, dword ptr [ebp-3c]

' *** Reference to "__vbaFreeObj"
'00725822    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_9)

' *** Reference to "__vbaExitProc"
'00725828    ff15a0104000            call dword ptr [004010a0]
'0072582e    9b                      fwait
'0072582f    68b1587200              push 007258b1
'00725834    eb7a                    jmp 7258b0
'00725836    8d55c8                  lea edx, dword ptr [ebp-38]
'00725839    52                      push edx
'0072583a    8d45cc                  lea eax, dword ptr [ebp-34]
'0072583d    50                      push eax
'0072583e    8d4dd0                  lea ecx, dword ptr [ebp-30]
'00725841    51                      push ecx
'00725842    8d55d4                  lea edx, dword ptr [ebp-2c]
'00725845    52                      push edx
'00725846    8d45d8                  lea eax, dword ptr [ebp-28]
'00725849    50                      push eax
'0072584a    8d4ddc                  lea ecx, dword ptr [ebp-24]
'0072584d    51                      push ecx
'0072584e    8d55e0                  lea edx, dword ptr [ebp-20]
'00725851    52                      push edx
'00725852    6a07                    push 07

' *** Reference to "__vbaFreeStrList"
'00725854    ff155c124000            call dword ptr [0040125c]
'#Cleanup( -4700, -4508, -4732, -4516, -4520, -4524, -4528)
'0072585a    8d45bc                  lea eax, dword ptr [ebp-44]
'0072585d    50                      push eax
'0072585e    8d4dc0                  lea ecx, dword ptr [ebp-40]
'00725861    51                      push ecx
'00725862    8d55c4                  lea edx, dword ptr [ebp-3c]
'00725865    52                      push edx
'00725866    6a03                    push 03

' *** Reference to "__vbaFreeObjList"
'00725868    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_9, var_5, var_58)
'0072586e    8d852cffffff            lea eax, dword ptr [ebp+ffffff2c]
'00725874    50                      push eax
'00725875    8d8d3cffffff            lea ecx, dword ptr [ebp+ffffff3c]
'0072587b    51                      push ecx
'0072587c    8d954cffffff            lea edx, dword ptr [ebp+ffffff4c]
'00725882    52                      push edx
'00725883    8d855cffffff            lea eax, dword ptr [ebp+ffffff5c]
'00725889    50                      push eax
'0072588a    8d8d6cffffff            lea ecx, dword ptr [ebp+ffffff6c]
'00725890    51                      push ecx
'00725891    8d957cffffff            lea edx, dword ptr [ebp+ffffff7c]
'00725897    52                      push edx
'00725898    8d458c                  lea eax, dword ptr [ebp-74]
'0072589b    50                      push eax
'0072589c    8d4d9c                  lea ecx, dword ptr [ebp-64]
'0072589f    51                      push ecx
'007258a0    8d55ac                  lea edx, dword ptr [ebp-54]
'007258a3    52                      push edx
'007258a4    6a09                    push 09

' *** Reference to "__vbaFreeVarList"
'007258a6    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_50, var_51, var_52, var_59, var_94, var_88, var_89, var_90, var_93)
'007258ac    83c458                  add esp, 58
'007258af    c3                      ret
'007258b0    c3                      ret
'007258b1    8b4508                  mov eax, dword ptr [ebp+08]
'007258b4    8b08                    mov ecx, dword ptr [eax]
'007258b6    50                      push eax
'007258b7    ff5108                  call dword ptr [ecx+08]
'007258ba    8b45f4                  mov eax, dword ptr [ebp-0c]
'007258bd    8b4de4                  mov ecx, dword ptr [ebp-1c]
    'Reference to '__except_list'
'007258c0    64890d00000000          mov dword ptr fs:[00000000], ecx
'007258c7    5f                      pop edi
'007258c8    5e                      pop esi
'007258c9    5b                      pop ebx
'007258ca    8be5                    mov esp, ebp
'007258cc    5d                      pop ebp
'007258cd    c20400                  ret 0004


End Sub


'Event for fldnPrixTotal
Private Sub fldnPrixTotal_Change()
'007261f0    55                      push ebp
'007261f1    8bec                    mov ebp, esp
'007261f3    83ec0c                  sub esp, 0c

' *** Reference to "__vbaExceptHandler"
'007261f6    6866724000              push 00407266
'007261fb    64a100000000            mov ax, word ptr fs:[00000000]
'00726201    50                      push eax
    'Reference to '__except_list'
'00726202    64892500000000          mov dword ptr fs:[00000000], esp
'00726209    81ecec000000            sub esp, 000000ec
'0072620f    53                      push ebx
'00726210    56                      push esi
'00726211    57                      push edi
'00726212    8965f4                  mov dword ptr [ebp-0c], esp
'00726215    c745f8d8714000          mov dword ptr [ebp-08], 004071d8
'0072621c    8b7508                  mov esi, dword ptr [ebp+08]
'0072621f    8bc6                    mov eax, esi
'00726221    83e001                  and eax, 01
'00726224    8945fc                  mov dword ptr [ebp-04], eax
'00726227    83e6fe                  and esi, fffffffe
'0072622a    8b0e                    mov ecx, dword ptr [esi]
'0072622c    56                      push esi
'0072622d    897508                  mov dword ptr [ebp+08], esi
'00726230    ff5104                  call dword ptr [ecx+04]
'00726233    8b16                    mov edx, dword ptr [esi]
'00726235    33ff                    xor edi, edi
'00726237    56                      push esi
'00726238    897de8                  mov dword ptr [ebp-18], edi
'0072623b    897de4                  mov dword ptr [ebp-1c], edi
'0072623e    897de0                  mov dword ptr [ebp-20], edi
'00726241    897ddc                  mov dword ptr [ebp-24], edi
'00726244    897dd8                  mov dword ptr [ebp-28], edi
'00726247    897dd4                  mov dword ptr [ebp-2c], edi
'0072624a    897dc4                  mov dword ptr [ebp-3c], edi
'0072624d    897db4                  mov dword ptr [ebp-4c], edi
'00726250    897da4                  mov dword ptr [ebp-5c], edi
'00726253    897d94                  mov dword ptr [ebp-6c], edi
'00726256    897d84                  mov dword ptr [ebp-7c], edi
'00726259    ff9204030000            call dword ptr [edx+00000304]

' *** Reference to "__vbaObjSet"
'0072625f    8b1db4104000            mov ebx, dword ptr [004010b4]
'00726265    50                      push eax
'00726266    8d45dc                  lea eax, dword ptr [ebp-24]
'00726269    50                      push eax
'0072626a    ffd3                    call ebx
Set var_39 = Me
'0072626c    8b08                    mov ecx, dword ptr [eax]
'0072626e    8d55e8                  lea edx, dword ptr [ebp-18]
'00726271    52                      push edx
'00726272    50                      push eax
'00726273    898560ffffff            mov dword ptr [ebp+ffffff60], eax
'00726279    ff91a0000000            call dword ptr [ecx+000000a0]
'0072627f    dbe2                    fnclex
'00726281    3bc7                    cmp eax, edi
'00726283    7d18                    jge 72629d

If (var_39 < 0) Then
'00726285    8b8d60ffffff            mov ecx, dword ptr [ebp+ffffff60]
'0072628b    68a0000000              push 000000a0
'00726290    68d00d4300              push 00430dd0
'00726295    51                      push ecx
'00726296    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00726297    ff1580104000            call dword ptr [00401080]
    
End If
'0072629d    8b55e8                  mov edx, dword ptr [ebp-18]
'007262a0    52                      push edx
'007262a1    68cc134300              push 004313cc

' *** Reference to "__vbaStrCmp"
'007262a6    ff1530114000            call dword ptr [00401130]
'007262ac    8bf8                    mov edi, eax
'007262ae    f7df                    neg edi
'007262b0    1bff                    sbb edi, edi
'007262b2    f7df                    neg edi
'007262b4    8d4de8                  lea ecx, dword ptr [ebp-18]
'007262b7    f7df                    neg edi

' *** Reference to "__vbaFreeStr"
'007262b9    ff150c134000            call dword ptr [0040130c]
'#Cleanup(var_41)
'007262bf    8d4ddc                  lea ecx, dword ptr [ebp-24]

' *** Reference to "__vbaFreeObj"
'007262c2    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_39)
'007262c8    6685ff                  test edi, edi
'007262cb    0f849b060000            je 72696c

If (((vbNullString) <> (vbNullChar))) Then
'007262d1    8b06                    mov eax, dword ptr [esi]
'007262d3    56                      push esi
'007262d4    ff9004030000            call dword ptr [eax+00000304]
'007262da    8d4dc4                  lea ecx, dword ptr [ebp-3c]
'007262dd    51                      push ecx
'007262de    8945cc                  mov dword ptr [ebp-34], eax
'007262e1    c745c409000000          mov dword ptr [ebp-3c], 00000009

' *** Reference to "rtcIsNumeric"
'007262e8    ff154c114000            call dword ptr [0040114c]
'007262ee    33ff                    xor edi, edi
    var_num7 = Empty
'007262f0    668bf8                  mov di, ax
'007262f3    8d4dc4                  lea ecx, dword ptr [ebp-3c]
'007262f6    f7d7                    not edi

' *** Reference to "__vbaFreeVar"
'007262f8    ff1528104000            call dword ptr [00401028]
    '#Cleanup(var_9)
'007262fe    6685ff                  test edi, edi
'00726301    0f8459020000            je 726560
    
    If (    Not (IsNumeric(frmVendre))) Then
'00726307    a124be7200              mov ax, word ptr [0072be24]
'0072630c    85c0                    test ax, ax
'0072630e    7510                    jne 726320
'00726310    6824be7200              push 0072be24
'00726315    6870004300              push 00430070

' *** Reference to "__vbaNew2"
'0072631a    ff152c124000            call dword ptr [0040122c]
    Dim var_2 As New Global
'00726320    8b3d24be7200            mov edi, dword ptr [0072be24]
'00726326    8b17                    mov edx, dword ptr [edi]
'00726328    8d45dc                  lea eax, dword ptr [ebp-24]
'0072632b    50                      push eax
'0072632c    57                      push edi
'0072632d    ff5218                  call dword ptr [edx+18]
    Set var_39 = var_2.Screen
'00726330    dbe2                    fnclex
'00726332    85c0                    test ax, ax
'00726334    7d0f                    jge 726345
'00726336    6a18                    push 18
'00726338    6860004300              push 00430060
'0072633d    57                      push edi
'0072633e    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0072633f    ff1580104000            call dword ptr [00401080]
'00726345    8b45dc                  mov eax, dword ptr [ebp-24]
'00726348    8b38                    mov edi, dword ptr [eax]
'0072634a    33c9                    xor ecx, ecx
'0072634c    898558ffffff            mov dword ptr [ebp+ffffff58], eax

' *** Reference to "__vbaI2I4"
'00726352    ff1550114000            call dword ptr [00401150]
'00726358    8bcf                    mov ecx, edi
'0072635a    8bbd58ffffff            mov edi, dword ptr [ebp+ffffff58]
'00726360    50                      push eax
'00726361    57                      push edi
'00726362    ff517c                  call dword ptr [ecx+7c]
    var_39.MousePointer = CInt(0)
'00726365    dbe2                    fnclex
'00726367    85c0                    test ax, ax
'00726369    7d0f                    jge 72637a
'0072636b    6a7c                    push 7c
'0072636d    6810be4300              push 0043be10
'00726372    57                      push edi
'00726373    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00726374    ff1580104000            call dword ptr [00401080]
'0072637a    8d4ddc                  lea ecx, dword ptr [ebp-24]

' *** Reference to "__vbaFreeObj"
'0072637d    ff1508134000            call dword ptr [00401308]
    '#Cleanup(var_39)
'00726383    8b16                    mov edx, dword ptr [esi]
'00726385    8d45e8                  lea eax, dword ptr [ebp-18]
'00726388    50                      push eax
'00726389    56                      push esi
'0072638a    ff5248                  call dword ptr [edx+48]
'0072638d    dbe2                    fnclex
'0072638f    85c0                    test ax, ax
'00726391    7d0f                    jge 7263a2
'00726393    6a48                    push 48
'00726395    6810f54400              push 0044f510
'0072639a    56                      push esi
'0072639b    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0072639c    ff1580104000            call dword ptr [00401080]
'007263a2    b80a000000              mov eax, 0000000a
'007263a7    894594                  mov dword ptr [ebp-6c], eax
'007263aa    8945a4                  mov dword ptr [ebp-5c], eax
'007263ad    8b45e8                  mov eax, dword ptr [ebp-18]
'007263b0    b904000280              mov ecx, 80020004
'007263b5    8945bc                  mov dword ptr [ebp-44], eax
'007263b8    b808000000              mov eax, 00000008
'007263bd    894d9c                  mov dword ptr [ebp-64], ecx
'007263c0    894dac                  mov dword ptr [ebp-54], ecx
'007263c3    8d5584                  lea edx, dword ptr [ebp-7c]
'007263c6    8d4dc4                  lea ecx, dword ptr [ebp-3c]
'007263c9    c745e800000000          mov dword ptr [ebp-18], 00000000
'007263d0    8945b4                  mov dword ptr [ebp-4c], eax
'007263d3    c7458c707b4500          mov dword ptr [ebp-74], 00457b70
'007263da    894584                  mov dword ptr [ebp-7c], eax

' *** Reference to "__vbaVarDup"
'007263dd    ff1598124000            call dword ptr [00401298]
    var_9 = ("Le prix total doit être entrée en tant que valeur numérique")
'007263e3    8d4d94                  lea ecx, dword ptr [ebp-6c]
'007263e6    51                      push ecx
'007263e7    8d55a4                  lea edx, dword ptr [ebp-5c]
'007263ea    52                      push edx
'007263eb    8d45b4                  lea eax, dword ptr [ebp-4c]
'007263ee    50                      push eax
'007263ef    6a30                    push 30
'007263f1    8d4dc4                  lea ecx, dword ptr [ebp-3c]
'007263f4    51                      push ecx

' *** Reference to "rtcMsgBox"
'007263f5    ff15b8104000            call dword ptr [004010b8]
    var_14 = MsgBox(var_9, 48, vbNullString)
'007263fb    8d5594                  lea edx, dword ptr [ebp-6c]
'007263fe    52                      push edx
'007263ff    8d45a4                  lea eax, dword ptr [ebp-5c]
'00726402    50                      push eax
'00726403    8d4db4                  lea ecx, dword ptr [ebp-4c]
'00726406    51                      push ecx
'00726407    8d55c4                  lea edx, dword ptr [ebp-3c]
'0072640a    52                      push edx
'0072640b    6a04                    push 04

' *** Reference to "__vbaFreeVarList"
'0072640d    ff1540104000            call dword ptr [00401040]
    '#Cleanup( var_9, var_62, var_83, var_121)
'00726413    8b06                    mov eax, dword ptr [esi]
'00726415    83c414                  add esp, 14
'00726418    56                      push esi
'00726419    ff9008030000            call dword ptr [eax+00000308]
'0072641f    50                      push eax
'00726420    8d4ddc                  lea ecx, dword ptr [ebp-24]
'00726423    51                      push ecx
'00726424    ffd3                    call ebx
    Set var_39 = Nothing
'00726426    8bf8                    mov edi, eax
'00726428    8b17                    mov edx, dword ptr [edi]
'0072642a    8d45e8                  lea eax, dword ptr [ebp-18]
'0072642d    50                      push eax
'0072642e    57                      push edi
'0072642f    ff92a0000000            call dword ptr [edx+000000a0]
'00726435    dbe2                    fnclex
'00726437    85c0                    test ax, ax
'00726439    7d12                    jge 72644d
'0072643b    68a0000000              push 000000a0
'00726440    68d00d4300              push 00430dd0
'00726445    57                      push edi
'00726446    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00726447    ff1580104000            call dword ptr [00401080]
'0072644d    8b0e                    mov ecx, dword ptr [esi]
'0072644f    56                      push esi
'00726450    ff9104030000            call dword ptr [ecx+00000304]
'00726456    50                      push eax
'00726457    8d55d4                  lea edx, dword ptr [ebp-2c]
'0072645a    52                      push edx
'0072645b    ffd3                    call ebx
    Set var_44 = 
'0072645d    898550ffffff            mov dword ptr [ebp+ffffff50], eax
'00726463    8b06                    mov eax, dword ptr [esi]
'00726465    56                      push esi
'00726466    ff900c030000            call dword ptr [eax+0000030c]
'0072646c    50                      push eax
'0072646d    8d4dd8                  lea ecx, dword ptr [ebp-28]
'00726470    51                      push ecx
'00726471    ffd3                    call ebx
    Set var_45 = Nothing
'00726473    8bf8                    mov edi, eax
'00726475    8b17                    mov edx, dword ptr [edi]
'00726477    8d45e4                  lea eax, dword ptr [ebp-1c]
'0072647a    50                      push eax
'0072647b    57                      push edi
'0072647c    ff92a0000000            call dword ptr [edx+000000a0]
'00726482    dbe2                    fnclex
'00726484    85c0                    test ax, ax
'00726486    7d12                    jge 72649a
'00726488    68a0000000              push 000000a0
'0072648d    68d00d4300              push 00430dd0
'00726492    57                      push edi
'00726493    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00726494    ff1580104000            call dword ptr [00401080]
'0072649a    8b55e4                  mov edx, dword ptr [ebp-1c]
'0072649d    8b8d50ffffff            mov ecx, dword ptr [ebp+ffffff50]
'007264a3    8b39                    mov edi, dword ptr [ecx]
'007264a5    52                      push edx

' *** Reference to "__vbaR8Str"
'007264a6    ff1524124000            call dword ptr [00401224]
'007264ac    dd9d2cffffff            fstp qword ptr [ebp+ffffff2c]
    'var_93 = (00)
'007264b2    8b45e8                  mov eax, dword ptr [ebp-18]
'007264b5    50                      push eax

' *** Reference to "__vbaR8Str"
'007264b6    ff1524124000            call dword ptr [00401224]
'007264bc    dc8d2cffffff            fmul qword ptr [ebp+ffffff2c]
'007264c2    83ec08                  sub esp, 08
'007264c5    dfe0                    fnstsw ax
'007264c7    a80d                    test al, 0d
'007264c9    0f8585070000            jne 726c54
'007264cf    dd1c24                  fstp qword ptr [esp]
    'var_748 = (00)

' *** Reference to "__vbaStrR8"
'007264d2    ff1580114000            call dword ptr [00401180]
'007264d8    8bd0                    mov edx, eax
'007264da    8d4de0                  lea ecx, dword ptr [ebp-20]

' *** Reference to "__vbaStrMove"
'007264dd    ff15d0124000            call dword ptr [004012d0]
'007264e3    8bcf                    mov ecx, edi
'007264e5    8bbd50ffffff            mov edi, dword ptr [ebp+ffffff50]
'007264eb    50                      push eax
'007264ec    57                      push edi
'007264ed    ff91a4000000            call dword ptr [ecx+000000a4]
'007264f3    dbe2                    fnclex
'007264f5    85c0                    test ax, ax
'007264f7    7d12                    jge 72650b
'007264f9    68a4000000              push 000000a4
'007264fe    68d00d4300              push 00430dd0
'00726503    57                      push edi
'00726504    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00726505    ff1580104000            call dword ptr [00401080]
'0072650b    8d55e0                  lea edx, dword ptr [ebp-20]
'0072650e    52                      push edx
'0072650f    8d45e8                  lea eax, dword ptr [ebp-18]
'00726512    50                      push eax
'00726513    8d4de4                  lea ecx, dword ptr [ebp-1c]
'00726516    51                      push ecx
'00726517    6a03                    push 03

' *** Reference to "__vbaFreeStrList"
'00726519    ff155c124000            call dword ptr [0040125c]
    '#Cleanup( 0, var_41, -4528)
'0072651f    8d55d4                  lea edx, dword ptr [ebp-2c]
'00726522    52                      push edx
'00726523    8d45d8                  lea eax, dword ptr [ebp-28]
'00726526    50                      push eax
'00726527    8d4ddc                  lea ecx, dword ptr [ebp-24]
'0072652a    51                      push ecx
'0072652b    6a03                    push 03

' *** Reference to "__vbaFreeObjList"
'0072652d    ff154c104000            call dword ptr [0040104c]
    '#Cleanup( var_39, var_45, var_44)
'00726533    8b16                    mov edx, dword ptr [esi]
'00726535    83c420                  add esp, 20
'00726538    56                      push esi
'00726539    ff9204030000            call dword ptr [edx+00000304]
'0072653f    50                      push eax
'00726540    8d45dc                  lea eax, dword ptr [ebp-24]
'00726543    50                      push eax
'00726544    ffd3                    call ebx
    Set var_39 = 
'00726546    8bf0                    mov esi, eax
'00726548    8b0e                    mov ecx, dword ptr [esi]
'0072654a    56                      push esi
'0072654b    ff9104020000            call dword ptr [ecx+00000204]
'00726551    dbe2                    fnclex
'00726553    85c0                    test ax, ax
'00726555    0f8d7d060000            jge 726bd8
'0072655b    e966060000              jmp 726bc6
    
End If
'00726560    8b16                    mov edx, dword ptr [esi]
'00726562    56                      push esi
'00726563    ff9204030000            call dword ptr [edx+00000304]
'00726569    50                      push eax
'0072656a    8d45dc                  lea eax, dword ptr [ebp-24]
'0072656d    50                      push eax
'0072656e    ffd3                    call ebx
Set var_39 = IsNumeric(frmVendre)
'00726570    8d55e8                  lea edx, dword ptr [ebp-18]
'00726573    8bf8                    mov edi, eax
'00726575    8b0f                    mov ecx, dword ptr [edi]
'00726577    52                      push edx
'00726578    57                      push edi
'00726579    ff91a0000000            call dword ptr [ecx+000000a0]
'0072657f    dbe2                    fnclex
'00726581    85c0                    test ax, ax
'00726583    7d12                    jge 726597
'00726585    68a0000000              push 000000a0
'0072658a    68d00d4300              push 00430dd0
'0072658f    57                      push edi
'00726590    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00726591    ff1580104000            call dword ptr [00401080]
'00726597    8b45e8                  mov eax, dword ptr [ebp-18]
'0072659a    50                      push eax

' *** Reference to "rtcR8ValFromBstr"
'0072659b    ff1510134000            call dword ptr [00401310]

' *** Reference to "__vbaFpR8"
'007265a1    ff15f8104000            call dword ptr [004010f8]
'007265a7    dc1d68154000            fcomp qword ptr [00401568]
'007265ad    dfe0                    fnstsw ax
'007265af    f6c401                  test ah, 01
'007265b2    7407                    je 7265bb
'007265b4    bf01000000              mov edi, 00000001
'007265b9    eb02                    jmp 7265bd
'007265bb    33ff                    xor edi, edi
'007265bd    8d4de8                  lea ecx, dword ptr [ebp-18]

' *** Reference to "__vbaFreeStr"
'007265c0    ff150c134000            call dword ptr [0040130c]
'#Cleanup(var_41)
'007265c6    8d4ddc                  lea ecx, dword ptr [ebp-24]

' *** Reference to "__vbaFreeObj"
'007265c9    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_39)
'007265cf    f7df                    neg edi
'007265d1    6685ff                  test edi, edi
'007265d4    0f8463020000            je 72683d

If ((0# > Val(vbNullString))) Then
'007265da    a124be7200              mov ax, word ptr [0072be24]
'007265df    85c0                    test ax, ax
'007265e1    7510                    jne 7265f3
'007265e3    6824be7200              push 0072be24
'007265e8    6870004300              push 00430070

' *** Reference to "__vbaNew2"
'007265ed    ff152c124000            call dword ptr [0040122c]
    Set var_2 = New Global
'007265f3    8b3d24be7200            mov edi, dword ptr [0072be24]
'007265f9    8b0f                    mov ecx, dword ptr [edi]
'007265fb    8d55dc                  lea edx, dword ptr [ebp-24]
'007265fe    52                      push edx
'007265ff    57                      push edi
'00726600    ff5118                  call dword ptr [ecx+18]
    Set var_39 = var_2.Screen
'00726603    dbe2                    fnclex
'00726605    85c0                    test ax, ax
'00726607    7d0f                    jge 726618
'00726609    6a18                    push 18
'0072660b    6860004300              push 00430060
'00726610    57                      push edi
'00726611    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00726612    ff1580104000            call dword ptr [00401080]
'00726618    8b45dc                  mov eax, dword ptr [ebp-24]
'0072661b    8b38                    mov edi, dword ptr [eax]
'0072661d    33c9                    xor ecx, ecx
'0072661f    898558ffffff            mov dword ptr [ebp+ffffff58], eax

' *** Reference to "__vbaI2I4"
'00726625    ff1550114000            call dword ptr [00401150]
'0072662b    89bd24ffffff            mov dword ptr [ebp+ffffff24], edi
'00726631    8bbd58ffffff            mov edi, dword ptr [ebp+ffffff58]
'00726637    50                      push eax
'00726638    8b8524ffffff            mov eax, dword ptr [ebp+ffffff24]
'0072663e    57                      push edi
'0072663f    ff507c                  call dword ptr [eax+7c]
    var_39.MousePointer = CInt((0# [#>] Val(vbNullString)))
'00726642    dbe2                    fnclex
'00726644    85c0                    test ax, ax
'00726646    7d0f                    jge 726657
'00726648    6a7c                    push 7c
'0072664a    6810be4300              push 0043be10
'0072664f    57                      push edi
'00726650    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00726651    ff1580104000            call dword ptr [00401080]
'00726657    8d4ddc                  lea ecx, dword ptr [ebp-24]

' *** Reference to "__vbaFreeObj"
'0072665a    ff1508134000            call dword ptr [00401308]
    '#Cleanup(var_39)
'00726660    8b0e                    mov ecx, dword ptr [esi]
'00726662    8d55e8                  lea edx, dword ptr [ebp-18]
'00726665    52                      push edx
'00726666    56                      push esi
'00726667    ff5148                  call dword ptr [ecx+48]
'0072666a    dbe2                    fnclex
'0072666c    85c0                    test ax, ax
'0072666e    7d0f                    jge 72667f
'00726670    6a48                    push 48
'00726672    6810f54400              push 0044f510
'00726677    56                      push esi
'00726678    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00726679    ff1580104000            call dword ptr [00401080]
'0072667f    b80a000000              mov eax, 0000000a
'00726684    894594                  mov dword ptr [ebp-6c], eax
'00726687    8945a4                  mov dword ptr [ebp-5c], eax
'0072668a    8b45e8                  mov eax, dword ptr [ebp-18]
'0072668d    b904000280              mov ecx, 80020004
'00726692    8945bc                  mov dword ptr [ebp-44], eax
'00726695    b808000000              mov eax, 00000008
'0072669a    894d9c                  mov dword ptr [ebp-64], ecx
'0072669d    894dac                  mov dword ptr [ebp-54], ecx
'007266a0    8d5584                  lea edx, dword ptr [ebp-7c]
'007266a3    8d4dc4                  lea ecx, dword ptr [ebp-3c]
'007266a6    c745e800000000          mov dword ptr [ebp-18], 00000000
'007266ad    8945b4                  mov dword ptr [ebp-4c], eax
'007266b0    c7458cec7b4500          mov dword ptr [ebp-74], 00457bec
'007266b7    894584                  mov dword ptr [ebp-7c], eax

' *** Reference to "__vbaVarDup"
'007266ba    ff1598124000            call dword ptr [00401298]
    var_9 = ("Le prix total doit être plus grand que 0")
'007266c0    8d4594                  lea eax, dword ptr [ebp-6c]
'007266c3    50                      push eax
'007266c4    8d4da4                  lea ecx, dword ptr [ebp-5c]
'007266c7    51                      push ecx
'007266c8    8d55b4                  lea edx, dword ptr [ebp-4c]
'007266cb    52                      push edx
'007266cc    6a30                    push 30
'007266ce    8d45c4                  lea eax, dword ptr [ebp-3c]
'007266d1    50                      push eax

' *** Reference to "rtcMsgBox"
'007266d2    ff15b8104000            call dword ptr [004010b8]
    var_53 = MsgBox(var_9, 48, vbNullString)
'007266d8    8d4d94                  lea ecx, dword ptr [ebp-6c]
'007266db    51                      push ecx
'007266dc    8d55a4                  lea edx, dword ptr [ebp-5c]
'007266df    52                      push edx
'007266e0    8d45b4                  lea eax, dword ptr [ebp-4c]
'007266e3    50                      push eax
'007266e4    8d4dc4                  lea ecx, dword ptr [ebp-3c]
'007266e7    51                      push ecx
'007266e8    6a04                    push 04

' *** Reference to "__vbaFreeVarList"
'007266ea    ff1540104000            call dword ptr [00401040]
    '#Cleanup( var_9, var_62, var_83, var_121)
'007266f0    8b16                    mov edx, dword ptr [esi]
'007266f2    83c414                  add esp, 14
'007266f5    56                      push esi
'007266f6    ff9208030000            call dword ptr [edx+00000308]
'007266fc    50                      push eax
'007266fd    8d45dc                  lea eax, dword ptr [ebp-24]
'00726700    50                      push eax
'00726701    ffd3                    call ebx
    Set var_39 = 
'00726703    8d55e8                  lea edx, dword ptr [ebp-18]
'00726706    8bf8                    mov edi, eax
'00726708    8b0f                    mov ecx, dword ptr [edi]
'0072670a    52                      push edx
'0072670b    57                      push edi
'0072670c    ff91a0000000            call dword ptr [ecx+000000a0]
'00726712    dbe2                    fnclex
'00726714    85c0                    test ax, ax
'00726716    7d12                    jge 72672a
'00726718    68a0000000              push 000000a0
'0072671d    68d00d4300              push 00430dd0
'00726722    57                      push edi
'00726723    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00726724    ff1580104000            call dword ptr [00401080]
'0072672a    8b06                    mov eax, dword ptr [esi]
'0072672c    56                      push esi
'0072672d    ff9004030000            call dword ptr [eax+00000304]
'00726733    50                      push eax
'00726734    8d4dd4                  lea ecx, dword ptr [ebp-2c]
'00726737    51                      push ecx
'00726738    ffd3                    call ebx
    Set var_44 = Nothing
'0072673a    8b16                    mov edx, dword ptr [esi]
'0072673c    56                      push esi
'0072673d    898550ffffff            mov dword ptr [ebp+ffffff50], eax
'00726743    ff920c030000            call dword ptr [edx+0000030c]
'00726749    50                      push eax
'0072674a    8d45d8                  lea eax, dword ptr [ebp-28]
'0072674d    50                      push eax
'0072674e    ffd3                    call ebx
    Set var_45 = Nothing
'00726750    8d55e4                  lea edx, dword ptr [ebp-1c]
'00726753    8bf8                    mov edi, eax
'00726755    8b0f                    mov ecx, dword ptr [edi]
'00726757    52                      push edx
'00726758    57                      push edi
'00726759    ff91a0000000            call dword ptr [ecx+000000a0]
'0072675f    dbe2                    fnclex
'00726761    85c0                    test ax, ax
'00726763    7d12                    jge 726777
'00726765    68a0000000              push 000000a0
'0072676a    68d00d4300              push 00430dd0
'0072676f    57                      push edi
'00726770    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00726771    ff1580104000            call dword ptr [00401080]
'00726777    8b4de4                  mov ecx, dword ptr [ebp-1c]
'0072677a    8b8550ffffff            mov eax, dword ptr [ebp+ffffff50]
'00726780    8b38                    mov edi, dword ptr [eax]
'00726782    51                      push ecx

' *** Reference to "__vbaR8Str"
'00726783    ff1524124000            call dword ptr [00401224]
'00726789    dd9d1cffffff            fstp qword ptr [ebp+ffffff1c]
    'var_95 = (00)
'0072678f    8b55e8                  mov edx, dword ptr [ebp-18]
'00726792    52                      push edx

' *** Reference to "__vbaR8Str"
'00726793    ff1524124000            call dword ptr [00401224]
'00726799    dc8d1cffffff            fmul qword ptr [ebp+ffffff1c]
'0072679f    83ec08                  sub esp, 08
'007267a2    dfe0                    fnstsw ax
'007267a4    a80d                    test al, 0d
'007267a6    0f85a8040000            jne 726c54
'007267ac    dd1c24                  fstp qword ptr [esp]
    'var_118 = (00)

' *** Reference to "__vbaStrR8"
'007267af    ff1580114000            call dword ptr [00401180]
'007267b5    8bd0                    mov edx, eax
'007267b7    8d4de0                  lea ecx, dword ptr [ebp-20]

' *** Reference to "__vbaStrMove"
'007267ba    ff15d0124000            call dword ptr [004012d0]
'007267c0    50                      push eax
'007267c1    8bc7                    mov eax, edi
'007267c3    8bbd50ffffff            mov edi, dword ptr [ebp+ffffff50]
'007267c9    57                      push edi
'007267ca    ff90a4000000            call dword ptr [eax+000000a4]
'007267d0    dbe2                    fnclex
'007267d2    85c0                    test ax, ax
'007267d4    7d12                    jge 7267e8
'007267d6    68a4000000              push 000000a4
'007267db    68d00d4300              push 00430dd0
'007267e0    57                      push edi
'007267e1    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'007267e2    ff1580104000            call dword ptr [00401080]
'007267e8    8d4de0                  lea ecx, dword ptr [ebp-20]
'007267eb    51                      push ecx
'007267ec    8d55e8                  lea edx, dword ptr [ebp-18]
'007267ef    52                      push edx
'007267f0    8d45e4                  lea eax, dword ptr [ebp-1c]
'007267f3    50                      push eax
'007267f4    6a03                    push 03

' *** Reference to "__vbaFreeStrList"
'007267f6    ff155c124000            call dword ptr [0040125c]
    '#Cleanup( 0, var_41, -4556)
'007267fc    8d4dd4                  lea ecx, dword ptr [ebp-2c]
'007267ff    51                      push ecx
'00726800    8d55d8                  lea edx, dword ptr [ebp-28]
'00726803    52                      push edx
'00726804    8d45dc                  lea eax, dword ptr [ebp-24]
'00726807    50                      push eax
'00726808    6a03                    push 03

' *** Reference to "__vbaFreeObjList"
'0072680a    ff154c104000            call dword ptr [0040104c]
    '#Cleanup( var_39, var_45, var_44)
'00726810    8b0e                    mov ecx, dword ptr [esi]
'00726812    83c420                  add esp, 20
'00726815    56                      push esi
'00726816    ff9104030000            call dword ptr [ecx+00000304]
'0072681c    50                      push eax
'0072681d    8d55dc                  lea edx, dword ptr [ebp-24]
'00726820    52                      push edx
'00726821    ffd3                    call ebx
    Set var_39 = 
'00726823    8bf0                    mov esi, eax
'00726825    8b06                    mov eax, dword ptr [esi]
'00726827    56                      push esi
'00726828    ff9004020000            call dword ptr [eax+00000204]
'0072682e    dbe2                    fnclex
'00726830    85c0                    test ax, ax
'00726832    0f8da0030000            jge 726bd8
'00726838    e989030000              jmp 726bc6
    
End If
'0072683d    8b0e                    mov ecx, dword ptr [esi]
'0072683f    56                      push esi
'00726840    ff9108030000            call dword ptr [ecx+00000308]
'00726846    50                      push eax
'00726847    8d55dc                  lea edx, dword ptr [ebp-24]
'0072684a    52                      push edx
'0072684b    ffd3                    call ebx
Set var_39 = Nothing
'0072684d    8d4de8                  lea ecx, dword ptr [ebp-18]
'00726850    8bf8                    mov edi, eax
'00726852    8b07                    mov eax, dword ptr [edi]
'00726854    51                      push ecx
'00726855    57                      push edi
'00726856    ff90a0000000            call dword ptr [eax+000000a0]
'0072685c    dbe2                    fnclex
'0072685e    85c0                    test ax, ax
'00726860    7d12                    jge 726874
'00726862    68a0000000              push 000000a0
'00726867    68d00d4300              push 00430dd0
'0072686c    57                      push edi
'0072686d    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0072686e    ff1580104000            call dword ptr [00401080]
'00726874    8b16                    mov edx, dword ptr [esi]
'00726876    56                      push esi
'00726877    ff920c030000            call dword ptr [edx+0000030c]
'0072687d    50                      push eax
'0072687e    8d45d4                  lea eax, dword ptr [ebp-2c]
'00726881    50                      push eax
'00726882    ffd3                    call ebx
Set var_44 = -256 - 12
'00726884    8b0e                    mov ecx, dword ptr [esi]
'00726886    56                      push esi
'00726887    8bf8                    mov edi, eax
'00726889    ff9104030000            call dword ptr [ecx+00000304]
'0072688f    50                      push eax
'00726890    8d55d8                  lea edx, dword ptr [ebp-28]
'00726893    52                      push edx
'00726894    ffd3                    call ebx
Set var_45 = var_44
'00726896    8d4de4                  lea ecx, dword ptr [ebp-1c]
'00726899    8bf0                    mov esi, eax
'0072689b    8b06                    mov eax, dword ptr [esi]
'0072689d    51                      push ecx
'0072689e    56                      push esi
'0072689f    ff90a0000000            call dword ptr [eax+000000a0]
'007268a5    dbe2                    fnclex
'007268a7    85c0                    test ax, ax
'007268a9    7d12                    jge 7268bd
'007268ab    68a0000000              push 000000a0
'007268b0    68d00d4300              push 00430dd0
'007268b5    56                      push esi
'007268b6    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'007268b7    ff1580104000            call dword ptr [00401080]
'007268bd    8b55e4                  mov edx, dword ptr [ebp-1c]
'007268c0    8b37                    mov esi, dword ptr [edi]
'007268c2    52                      push edx

' *** Reference to "__vbaR8Str"
'007268c3    ff1524124000            call dword ptr [00401224]
'007268c9    dd9d10ffffff            fstp qword ptr [ebp+ffffff10]
'var_1494 = (00)
'007268cf    8b45e8                  mov eax, dword ptr [ebp-18]
'007268d2    50                      push eax

' *** Reference to "__vbaR8Str"
'007268d3    ff1524124000            call dword ptr [00401224]
'007268d9    833d00b0720000          cmp dword ptr [0072b000], 00
'007268e0    7508                    jne 7268ea

If (vbNullChar = 0) Then
'007268e2    dcbd10ffffff            fdivr qword ptr [ebp+ffffff10]
'007268e8    eb11                    jmp 7268fb
    
End If
'007268ea    ffb514ffffff            push dword ptr [ebp+ffffff14]
'007268f0    ffb510ffffff            push dword ptr [ebp+ffffff10]

' *** Reference to "_adj_fdivr_m64"
'007268f6    e8a709ceff              call 4072a2
'007268fb    83ec08                  sub esp, 08
'007268fe    dfe0                    fnstsw ax
'00726900    a80d                    test al, 0d
'00726902    0f854c030000            jne 726c54
'00726908    dd1c24                  fstp qword ptr [esp]
'var_825 = (00)

' *** Reference to "__vbaStrR8"
'0072690b    ff1580114000            call dword ptr [00401180]
'00726911    8bd0                    mov edx, eax
'00726913    8d4de0                  lea ecx, dword ptr [ebp-20]

' *** Reference to "__vbaStrMove"
'00726916    ff15d0124000            call dword ptr [004012d0]
'0072691c    50                      push eax
'0072691d    57                      push edi
'0072691e    ff96a4000000            call dword ptr [esi+000000a4]
'00726924    dbe2                    fnclex
'00726926    85c0                    test ax, ax
'00726928    7d12                    jge 72693c
'0072692a    68a4000000              push 000000a4
'0072692f    68d00d4300              push 00430dd0
'00726934    57                      push edi
'00726935    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00726936    ff1580104000            call dword ptr [00401080]
'0072693c    8d4de0                  lea ecx, dword ptr [ebp-20]
'0072693f    51                      push ecx
'00726940    8d55e8                  lea edx, dword ptr [ebp-18]
'00726943    52                      push edx
'00726944    8d45e4                  lea eax, dword ptr [ebp-1c]
'00726947    50                      push eax
'00726948    6a03                    push 03

' *** Reference to "__vbaFreeStrList"
'0072694a    ff155c124000            call dword ptr [0040125c]
'#Cleanup( 0, var_41, -4560)
'00726950    8d4dd4                  lea ecx, dword ptr [ebp-2c]
'00726953    51                      push ecx
'00726954    8d55d8                  lea edx, dword ptr [ebp-28]
'00726957    52                      push edx
'00726958    8d45dc                  lea eax, dword ptr [ebp-24]
'0072695b    50                      push eax
'0072695c    6a03                    push 03

' *** Reference to "__vbaFreeObjList"
'0072695e    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_39, var_45, var_44)
'00726964    83c420                  add esp, 20
'00726967    e975020000              jmp 726be1

'ERROR: Two many next close:
End If
'0072696c    a124be7200              mov ax, word ptr [0072be24]
'00726971    85c0                    test ax, ax
'00726973    7510                    jne 726985
'00726975    6824be7200              push 0072be24
'0072697a    6870004300              push 00430070

' *** Reference to "__vbaNew2"
'0072697f    ff152c124000            call dword ptr [0040122c]
Set var_2 = New Global
'00726985    8b3d24be7200            mov edi, dword ptr [0072be24]
'0072698b    8b0f                    mov ecx, dword ptr [edi]
'0072698d    8d55dc                  lea edx, dword ptr [ebp-24]
'00726990    52                      push edx
'00726991    57                      push edi
'00726992    ff5118                  call dword ptr [ecx+18]
Set var_39 = var_2.Screen
'00726995    dbe2                    fnclex
'00726997    85c0                    test ax, ax
'00726999    7d0f                    jge 7269aa
'0072699b    6a18                    push 18
'0072699d    6860004300              push 00430060
'007269a2    57                      push edi
'007269a3    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'007269a4    ff1580104000            call dword ptr [00401080]
'007269aa    8b45dc                  mov eax, dword ptr [ebp-24]
'007269ad    8b38                    mov edi, dword ptr [eax]
'007269af    33c9                    xor ecx, ecx
var_num3 = Empty
'007269b1    898558ffffff            mov dword ptr [ebp+ffffff58], eax

' *** Reference to "__vbaI2I4"
'007269b7    ff1550114000            call dword ptr [00401150]
'007269bd    89bd0cffffff            mov dword ptr [ebp+ffffff0c], edi
'007269c3    8bbd58ffffff            mov edi, dword ptr [ebp+ffffff58]
'007269c9    50                      push eax
'007269ca    8b850cffffff            mov eax, dword ptr [ebp+ffffff0c]
'007269d0    57                      push edi
'007269d1    ff507c                  call dword ptr [eax+7c]
var_39.MousePointer = CInt(Global)
'007269d4    dbe2                    fnclex
'007269d6    85c0                    test ax, ax
'007269d8    7d0f                    jge 7269e9
'007269da    6a7c                    push 7c
'007269dc    6810be4300              push 0043be10
'007269e1    57                      push edi
'007269e2    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'007269e3    ff1580104000            call dword ptr [00401080]
'007269e9    8d4ddc                  lea ecx, dword ptr [ebp-24]

' *** Reference to "__vbaFreeObj"
'007269ec    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_39)
'007269f2    8b0e                    mov ecx, dword ptr [esi]
'007269f4    8d55e8                  lea edx, dword ptr [ebp-18]
'007269f7    52                      push edx
'007269f8    56                      push esi
'007269f9    ff5148                  call dword ptr [ecx+48]
'007269fc    dbe2                    fnclex
'007269fe    85c0                    test ax, ax
'00726a00    7d0f                    jge 726a11
'00726a02    6a48                    push 48
'00726a04    6810f54400              push 0044f510
'00726a09    56                      push esi
'00726a0a    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00726a0b    ff1580104000            call dword ptr [00401080]
'00726a11    b80a000000              mov eax, 0000000a
'00726a16    894594                  mov dword ptr [ebp-6c], eax
'00726a19    8945a4                  mov dword ptr [ebp-5c], eax
'00726a1c    8b45e8                  mov eax, dword ptr [ebp-18]
'00726a1f    b904000280              mov ecx, 80020004
'00726a24    8945bc                  mov dword ptr [ebp-44], eax
'00726a27    b808000000              mov eax, 00000008
'00726a2c    894d9c                  mov dword ptr [ebp-64], ecx
'00726a2f    894dac                  mov dword ptr [ebp-54], ecx
'00726a32    8d5584                  lea edx, dword ptr [ebp-7c]
'00726a35    8d4dc4                  lea ecx, dword ptr [ebp-3c]
'00726a38    c745e800000000          mov dword ptr [ebp-18], 00000000
'00726a3f    8945b4                  mov dword ptr [ebp-4c], eax
'00726a42    c7458c707b4500          mov dword ptr [ebp-74], 00457b70
'00726a49    894584                  mov dword ptr [ebp-7c], eax

' *** Reference to "__vbaVarDup"
'00726a4c    ff1598124000            call dword ptr [00401298]
var_9 = ("Le prix total doit être entrée en tant que valeur numérique")
'00726a52    8d4594                  lea eax, dword ptr [ebp-6c]
'00726a55    50                      push eax
'00726a56    8d4da4                  lea ecx, dword ptr [ebp-5c]
'00726a59    51                      push ecx
'00726a5a    8d55b4                  lea edx, dword ptr [ebp-4c]
'00726a5d    52                      push edx
'00726a5e    6a30                    push 30
'00726a60    8d45c4                  lea eax, dword ptr [ebp-3c]
'00726a63    50                      push eax

' *** Reference to "rtcMsgBox"
'00726a64    ff15b8104000            call dword ptr [004010b8]
var_1231 = MsgBox(var_9, 48, vbNullString)
'00726a6a    8d4d94                  lea ecx, dword ptr [ebp-6c]
'00726a6d    51                      push ecx
'00726a6e    8d55a4                  lea edx, dword ptr [ebp-5c]
'00726a71    52                      push edx
'00726a72    8d45b4                  lea eax, dword ptr [ebp-4c]
'00726a75    50                      push eax
'00726a76    8d4dc4                  lea ecx, dword ptr [ebp-3c]
'00726a79    51                      push ecx
'00726a7a    6a04                    push 04

' *** Reference to "__vbaFreeVarList"
'00726a7c    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_9, var_62, var_83, var_121)
'00726a82    8b16                    mov edx, dword ptr [esi]
'00726a84    83c414                  add esp, 14
'00726a87    56                      push esi
'00726a88    ff9208030000            call dword ptr [edx+00000308]
'00726a8e    50                      push eax
'00726a8f    8d45dc                  lea eax, dword ptr [ebp-24]
'00726a92    50                      push eax
'00726a93    ffd3                    call ebx
Set var_39 = 
'00726a95    8d55e8                  lea edx, dword ptr [ebp-18]
'00726a98    8bf8                    mov edi, eax
'00726a9a    8b0f                    mov ecx, dword ptr [edi]
'00726a9c    52                      push edx
'00726a9d    57                      push edi
'00726a9e    ff91a0000000            call dword ptr [ecx+000000a0]
'00726aa4    dbe2                    fnclex
'00726aa6    85c0                    test ax, ax
'00726aa8    7d12                    jge 726abc
'00726aaa    68a0000000              push 000000a0
'00726aaf    68d00d4300              push 00430dd0
'00726ab4    57                      push edi
'00726ab5    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00726ab6    ff1580104000            call dword ptr [00401080]
'00726abc    8b06                    mov eax, dword ptr [esi]
'00726abe    56                      push esi
'00726abf    ff9004030000            call dword ptr [eax+00000304]
'00726ac5    50                      push eax
'00726ac6    8d4dd4                  lea ecx, dword ptr [ebp-2c]
'00726ac9    51                      push ecx
'00726aca    ffd3                    call ebx
Set var_44 = Nothing
'00726acc    8b16                    mov edx, dword ptr [esi]
'00726ace    56                      push esi
'00726acf    898550ffffff            mov dword ptr [ebp+ffffff50], eax
'00726ad5    ff920c030000            call dword ptr [edx+0000030c]
'00726adb    50                      push eax
'00726adc    8d45d8                  lea eax, dword ptr [ebp-28]
'00726adf    50                      push eax
'00726ae0    ffd3                    call ebx
Set var_45 = Nothing
'00726ae2    8d55e4                  lea edx, dword ptr [ebp-1c]
'00726ae5    8bf8                    mov edi, eax
'00726ae7    8b0f                    mov ecx, dword ptr [edi]
'00726ae9    52                      push edx
'00726aea    57                      push edi
'00726aeb    ff91a0000000            call dword ptr [ecx+000000a0]
'00726af1    dbe2                    fnclex
'00726af3    85c0                    test ax, ax
'00726af5    7d12                    jge 726b09
'00726af7    68a0000000              push 000000a0
'00726afc    68d00d4300              push 00430dd0
'00726b01    57                      push edi
'00726b02    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00726b03    ff1580104000            call dword ptr [00401080]
'00726b09    8b4de4                  mov ecx, dword ptr [ebp-1c]
'00726b0c    8b8550ffffff            mov eax, dword ptr [ebp+ffffff50]
'00726b12    8b38                    mov edi, dword ptr [eax]
'00726b14    51                      push ecx

' *** Reference to "__vbaR8Str"
'00726b15    ff1524124000            call dword ptr [00401224]
'00726b1b    dd9d04ffffff            fstp qword ptr [ebp+ffffff04]
'var_458 = (00)
'00726b21    8b55e8                  mov edx, dword ptr [ebp-18]
'00726b24    52                      push edx

' *** Reference to "__vbaR8Str"
'00726b25    ff1524124000            call dword ptr [00401224]
'00726b2b    dc8d04ffffff            fmul qword ptr [ebp+ffffff04]
'00726b31    83ec08                  sub esp, 08
'00726b34    dfe0                    fnstsw ax
'00726b36    a80d                    test al, 0d
'00726b38    0f8516010000            jne 726c54
'00726b3e    dd1c24                  fstp qword ptr [esp]
'var_748 = (00)

' *** Reference to "__vbaStrR8"
'00726b41    ff1580114000            call dword ptr [00401180]
'00726b47    8bd0                    mov edx, eax
'00726b49    8d4de0                  lea ecx, dword ptr [ebp-20]

' *** Reference to "__vbaStrMove"
'00726b4c    ff15d0124000            call dword ptr [004012d0]
'00726b52    50                      push eax
'00726b53    8bc7                    mov eax, edi
'00726b55    8bbd50ffffff            mov edi, dword ptr [ebp+ffffff50]
'00726b5b    57                      push edi
'00726b5c    ff90a4000000            call dword ptr [eax+000000a4]
'00726b62    dbe2                    fnclex
'00726b64    85c0                    test ax, ax
'00726b66    7d12                    jge 726b7a
'00726b68    68a4000000              push 000000a4
'00726b6d    68d00d4300              push 00430dd0
'00726b72    57                      push edi
'00726b73    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00726b74    ff1580104000            call dword ptr [00401080]
'00726b7a    8d4de0                  lea ecx, dword ptr [ebp-20]
'00726b7d    51                      push ecx
'00726b7e    8d55e8                  lea edx, dword ptr [ebp-18]
'00726b81    52                      push edx
'00726b82    8d45e4                  lea eax, dword ptr [ebp-1c]
'00726b85    50                      push eax
'00726b86    6a03                    push 03

' *** Reference to "__vbaFreeStrList"
'00726b88    ff155c124000            call dword ptr [0040125c]
'#Cleanup( 0, var_41, -4588)
'00726b8e    8d4dd4                  lea ecx, dword ptr [ebp-2c]
'00726b91    51                      push ecx
'00726b92    8d55d8                  lea edx, dword ptr [ebp-28]
'00726b95    52                      push edx
'00726b96    8d45dc                  lea eax, dword ptr [ebp-24]
'00726b99    50                      push eax
'00726b9a    6a03                    push 03

' *** Reference to "__vbaFreeObjList"
'00726b9c    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_39, var_45, var_44)
'00726ba2    8b0e                    mov ecx, dword ptr [esi]
'00726ba4    83c420                  add esp, 20
'00726ba7    56                      push esi
'00726ba8    ff9104030000            call dword ptr [ecx+00000304]
'00726bae    50                      push eax
'00726baf    8d55dc                  lea edx, dword ptr [ebp-24]
'00726bb2    52                      push edx
'00726bb3    ffd3                    call ebx
Set var_39 = 
'00726bb5    8bf0                    mov esi, eax
'00726bb7    8b06                    mov eax, dword ptr [esi]
'00726bb9    56                      push esi
'00726bba    ff9004020000            call dword ptr [eax+00000204]
'00726bc0    dbe2                    fnclex
'00726bc2    85c0                    test ax, ax
'00726bc4    7d12                    jge 726bd8
'00726bc6    6804020000              push 00000204
'00726bcb    68d00d4300              push 00430dd0
'00726bd0    56                      push esi
'00726bd1    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00726bd2    ff1580104000            call dword ptr [00401080]
'00726bd8    8d4ddc                  lea ecx, dword ptr [ebp-24]

' *** Reference to "__vbaFreeObj"
'00726bdb    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_39)
'00726be1    c745fc00000000          mov dword ptr [ebp-04], 00000000
'00726be8    9b                      fwait
'00726be9    68356c7200              push 00726c35
'00726bee    eb44                    jmp 726c34
'00726bf0    8d4de0                  lea ecx, dword ptr [ebp-20]
'00726bf3    51                      push ecx
'00726bf4    8d55e4                  lea edx, dword ptr [ebp-1c]
'00726bf7    52                      push edx
'00726bf8    8d45e8                  lea eax, dword ptr [ebp-18]
'00726bfb    50                      push eax
'00726bfc    6a03                    push 03

' *** Reference to "__vbaFreeStrList"
'00726bfe    ff155c124000            call dword ptr [0040125c]
'#Cleanup( var_41, 0, -4588)
'00726c04    8d4dd4                  lea ecx, dword ptr [ebp-2c]
'00726c07    51                      push ecx
'00726c08    8d55d8                  lea edx, dword ptr [ebp-28]
'00726c0b    52                      push edx
'00726c0c    8d45dc                  lea eax, dword ptr [ebp-24]
'00726c0f    50                      push eax
'00726c10    6a03                    push 03

' *** Reference to "__vbaFreeObjList"
'00726c12    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_39, var_45, var_44)
'00726c18    8d4d94                  lea ecx, dword ptr [ebp-6c]
'00726c1b    51                      push ecx
'00726c1c    8d55a4                  lea edx, dword ptr [ebp-5c]
'00726c1f    52                      push edx
'00726c20    8d45b4                  lea eax, dword ptr [ebp-4c]
'00726c23    50                      push eax
'00726c24    8d4dc4                  lea ecx, dword ptr [ebp-3c]
'00726c27    51                      push ecx
'00726c28    6a04                    push 04

' *** Reference to "__vbaFreeVarList"
'00726c2a    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_9, var_62, var_83, var_121)
'00726c30    83c434                  add esp, 34
'00726c33    c3                      ret
'00726c34    c3                      ret
'00726c35    8b4508                  mov eax, dword ptr [ebp+08]
'00726c38    8b10                    mov edx, dword ptr [eax]
'00726c3a    50                      push eax
'00726c3b    ff5208                  call dword ptr [edx+08]
'00726c3e    8b45fc                  mov eax, dword ptr [ebp-04]
'00726c41    8b4dec                  mov ecx, dword ptr [ebp-14]
'00726c44    5f                      pop edi
'00726c45    5e                      pop esi
    'Reference to '__except_list'
'00726c46    64890d00000000          mov dword ptr fs:[00000000], ecx
'00726c4d    5b                      pop ebx
'00726c4e    8be5                    mov esp, ebp
'00726c50    5d                      pop ebp
'00726c51    c20400                  ret 0004


End Sub


'Event for fldnNombreObjet
Private Sub fldnNombreObjet_Change()
'007258e0    55                      push ebp
'007258e1    8bec                    mov ebp, esp
'007258e3    83ec0c                  sub esp, 0c

' *** Reference to "__vbaExceptHandler"
'007258e6    6866724000              push 00407266
'007258eb    64a100000000            mov ax, word ptr fs:[00000000]
'007258f1    50                      push eax
    'Reference to '__except_list'
'007258f2    64892500000000          mov dword ptr fs:[00000000], esp
'007258f9    81ecb4000000            sub esp, 000000b4
'007258ff    53                      push ebx
'00725900    56                      push esi
'00725901    57                      push edi
'00725902    8965f4                  mov dword ptr [ebp-0c], esp
'00725905    c745f8c8714000          mov dword ptr [ebp-08], 004071c8
'0072590c    8b7508                  mov esi, dword ptr [ebp+08]
'0072590f    8bc6                    mov eax, esi
'00725911    83e001                  and eax, 01
'00725914    8945fc                  mov dword ptr [ebp-04], eax
'00725917    83e6fe                  and esi, fffffffe
'0072591a    8b0e                    mov ecx, dword ptr [esi]
'0072591c    56                      push esi
'0072591d    897508                  mov dword ptr [ebp+08], esi
'00725920    ff5104                  call dword ptr [ecx+04]
'00725923    8b16                    mov edx, dword ptr [esi]
'00725925    33c0                    xor eax, eax
var_num1 = Empty
'00725927    56                      push esi
'00725928    8945e8                  mov dword ptr [ebp-18], eax
'0072592b    8945e4                  mov dword ptr [ebp-1c], eax
'0072592e    8945e0                  mov dword ptr [ebp-20], eax
'00725931    8945d0                  mov dword ptr [ebp-30], eax
'00725934    8945c0                  mov dword ptr [ebp-40], eax
'00725937    8945b0                  mov dword ptr [ebp-50], eax
'0072593a    8945a0                  mov dword ptr [ebp-60], eax
'0072593d    894590                  mov dword ptr [ebp-70], eax
'00725940    ff9208030000            call dword ptr [edx+00000308]

' *** Reference to "__vbaObjSet"
'00725946    8b1db4104000            mov ebx, dword ptr [004010b4]
'0072594c    50                      push eax
'0072594d    8d45e0                  lea eax, dword ptr [ebp-20]
'00725950    50                      push eax
'00725951    ffd3                    call ebx
Set var_3 = Me
'00725953    8d55e8                  lea edx, dword ptr [ebp-18]
'00725956    8bf8                    mov edi, eax
'00725958    8b0f                    mov ecx, dword ptr [edi]
'0072595a    52                      push edx
'0072595b    57                      push edi
'0072595c    ff91a0000000            call dword ptr [ecx+000000a0]
'00725962    dbe2                    fnclex
'00725964    85c0                    test ax, ax
'00725966    7d12                    jge 72597a
'00725968    68a0000000              push 000000a0
'0072596d    68d00d4300              push 00430dd0
'00725972    57                      push edi
'00725973    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00725974    ff1580104000            call dword ptr [00401080]
'0072597a    8b45e8                  mov eax, dword ptr [ebp-18]
'0072597d    50                      push eax
'0072597e    68cc134300              push 004313cc

' *** Reference to "__vbaStrCmp"
'00725983    ff1530114000            call dword ptr [00401130]
'00725989    8bf8                    mov edi, eax
'0072598b    f7df                    neg edi
'0072598d    1bff                    sbb edi, edi
'0072598f    f7df                    neg edi
'00725991    8d4de8                  lea ecx, dword ptr [ebp-18]
'00725994    f7df                    neg edi

' *** Reference to "__vbaFreeStr"
'00725996    ff150c134000            call dword ptr [0040130c]
'#Cleanup(var_41)
'0072599c    8d4de0                  lea ecx, dword ptr [ebp-20]

' *** Reference to "__vbaFreeObj"
'0072599f    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_3)
'007259a5    6685ff                  test edi, edi
'007259a8    0f8431060000            je 725fdf

If (((frmVendre) <> (vbNullChar))) Then
'007259ae    8b0e                    mov ecx, dword ptr [esi]
'007259b0    56                      push esi
'007259b1    ff9108030000            call dword ptr [ecx+00000308]
'007259b7    8d55d0                  lea edx, dword ptr [ebp-30]
'007259ba    52                      push edx
'007259bb    8945d8                  mov dword ptr [ebp-28], eax
'007259be    c745d009000000          mov dword ptr [ebp-30], 00000009

' *** Reference to "rtcIsNumeric"
'007259c5    ff154c114000            call dword ptr [0040114c]
'007259cb    33ff                    xor edi, edi
    var_num7 = Empty
'007259cd    668bf8                  mov di, ax
'007259d0    8d4dd0                  lea ecx, dword ptr [ebp-30]
'007259d3    f7d7                    not edi

' *** Reference to "__vbaFreeVar"
'007259d5    ff1528104000            call dword ptr [00401028]
    '#Cleanup(var_4)
'007259db    6685ff                  test edi, edi
'007259de    0f8493010000            je 725b77
    
    If (    IsNumeric(((frmVendre) [#!#] (vbNullChar)))) Then
'007259e4    a124be7200              mov ax, word ptr [0072be24]
'007259e9    85c0                    test ax, ax
'007259eb    7510                    jne 7259fd
'007259ed    6824be7200              push 0072be24
'007259f2    6870004300              push 00430070

' *** Reference to "__vbaNew2"
'007259f7    ff152c124000            call dword ptr [0040122c]
    Dim var_2 As New Global
'007259fd    8b3d24be7200            mov edi, dword ptr [0072be24]
'00725a03    8b07                    mov eax, dword ptr [edi]
'00725a05    8d4de0                  lea ecx, dword ptr [ebp-20]
'00725a08    51                      push ecx
'00725a09    57                      push edi
'00725a0a    ff5018                  call dword ptr [eax+18]
    Set var_3 = var_2.Screen
'00725a0d    dbe2                    fnclex
'00725a0f    85c0                    test ax, ax
'00725a11    7d0f                    jge 725a22
'00725a13    6a18                    push 18
'00725a15    6860004300              push 00430060
'00725a1a    57                      push edi
'00725a1b    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00725a1c    ff1580104000            call dword ptr [00401080]
'00725a22    8b7de0                  mov edi, dword ptr [ebp-20]
'00725a25    8b1f                    mov ebx, dword ptr [edi]
'00725a27    33c9                    xor ecx, ecx

' *** Reference to "__vbaI2I4"
'00725a29    ff1550114000            call dword ptr [00401150]
'00725a2f    50                      push eax
'00725a30    57                      push edi
'00725a31    ff537c                  call dword ptr [ebx+7c]
    var_3.MousePointer = CInt(0)
'00725a34    dbe2                    fnclex
'00725a36    85c0                    test ax, ax
'00725a38    7d0f                    jge 725a49
'00725a3a    6a7c                    push 7c
'00725a3c    6810be4300              push 0043be10
'00725a41    57                      push edi
'00725a42    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00725a43    ff1580104000            call dword ptr [00401080]
'00725a49    8d4de0                  lea ecx, dword ptr [ebp-20]

' *** Reference to "__vbaFreeObj"
'00725a4c    ff1508134000            call dword ptr [00401308]
    '#Cleanup(var_3)
'00725a52    8b16                    mov edx, dword ptr [esi]
'00725a54    8d45e8                  lea eax, dword ptr [ebp-18]
'00725a57    50                      push eax
'00725a58    56                      push esi
'00725a59    ff5248                  call dword ptr [edx+48]
'00725a5c    dbe2                    fnclex
'00725a5e    85c0                    test ax, ax
'00725a60    7d0f                    jge 725a71
'00725a62    6a48                    push 48
'00725a64    6810f54400              push 0044f510
'00725a69    56                      push esi
'00725a6a    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00725a6b    ff1580104000            call dword ptr [00401080]
'00725a71    b80a000000              mov eax, 0000000a
'00725a76    b904000280              mov ecx, 80020004
'00725a7b    8945a0                  mov dword ptr [ebp-60], eax
'00725a7e    8945b0                  mov dword ptr [ebp-50], eax
'00725a81    8b45e8                  mov eax, dword ptr [ebp-18]
'00725a84    894da8                  mov dword ptr [ebp-58], ecx
'00725a87    894db8                  mov dword ptr [ebp-48], ecx
'00725a8a    bf08000000              mov edi, 00000008
'00725a8f    8d5590                  lea edx, dword ptr [ebp-70]
'00725a92    8d4dd0                  lea ecx, dword ptr [ebp-30]
'00725a95    c745e800000000          mov dword ptr [ebp-18], 00000000
'00725a9c    8945c8                  mov dword ptr [ebp-38], eax
'00725a9f    897dc0                  mov dword ptr [ebp-40], edi
'00725aa2    c745987c244500          mov dword ptr [ebp-68], 0045247c
'00725aa9    897d90                  mov dword ptr [ebp-70], edi

' *** Reference to "__vbaVarDup"
'00725aac    ff1598124000            call dword ptr [00401298]
    var_4 = ("Le nombre d'objet doit être entrée en tant que valeur numérique")
'00725ab2    8d4da0                  lea ecx, dword ptr [ebp-60]
'00725ab5    51                      push ecx
'00725ab6    8d55b0                  lea edx, dword ptr [ebp-50]
'00725ab9    52                      push edx
'00725aba    8d45c0                  lea eax, dword ptr [ebp-40]
'00725abd    50                      push eax
'00725abe    6a30                    push 30
'00725ac0    8d4dd0                  lea ecx, dword ptr [ebp-30]
'00725ac3    51                      push ecx

' *** Reference to "rtcMsgBox"
'00725ac4    ff15b8104000            call dword ptr [004010b8]
    var_14 = MsgBox(var_4, 48, vbNullString)
'00725aca    8d55a0                  lea edx, dword ptr [ebp-60]
'00725acd    52                      push edx
'00725ace    8d45b0                  lea eax, dword ptr [ebp-50]
'00725ad1    50                      push eax
'00725ad2    8d4dc0                  lea ecx, dword ptr [ebp-40]
'00725ad5    51                      push ecx
'00725ad6    8d55d0                  lea edx, dword ptr [ebp-30]
'00725ad9    52                      push edx
'00725ada    6a04                    push 04

' *** Reference to "__vbaFreeVarList"
'00725adc    ff1540104000            call dword ptr [00401040]
    '#Cleanup( var_4, var_5, var_6, var_7)
'00725ae2    8b06                    mov eax, dword ptr [esi]
'00725ae4    83c414                  add esp, 14
'00725ae7    56                      push esi
'00725ae8    ff9008030000            call dword ptr [eax+00000308]
'00725aee    50                      push eax
'00725aef    8d4de0                  lea ecx, dword ptr [ebp-20]
'00725af2    51                      push ecx

' *** Reference to "__vbaObjSet"
'00725af3    ff15b4104000            call dword ptr [004010b4]
    Set var_3 = Nothing
'00725af9    33d2                    xor edx, edx
'00725afb    668b563c                mov dx, word ptr [esi+3c]
'00725aff    8bf8                    mov edi, eax
'00725b01    8b1f                    mov ebx, dword ptr [edi]
'00725b03    52                      push edx

' *** Reference to "__vbaStrI2"
'00725b04    ff1510104000            call dword ptr [00401010]
'00725b0a    8bd0                    mov edx, eax
'00725b0c    8d4de8                  lea ecx, dword ptr [ebp-18]

' *** Reference to "__vbaStrMove"
'00725b0f    ff15d0124000            call dword ptr [004012d0]
'00725b15    50                      push eax
'00725b16    57                      push edi
'00725b17    ff93a4000000            call dword ptr [ebx+000000a4]
'00725b1d    dbe2                    fnclex
'00725b1f    85c0                    test ax, ax
'00725b21    7d12                    jge 725b35
'00725b23    68a4000000              push 000000a4
'00725b28    68d00d4300              push 00430dd0
'00725b2d    57                      push edi
'00725b2e    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00725b2f    ff1580104000            call dword ptr [00401080]
'00725b35    8d4de8                  lea ecx, dword ptr [ebp-18]

' *** Reference to "__vbaFreeStr"
'00725b38    ff150c134000            call dword ptr [0040130c]
    '#Cleanup(var_41)

' *** Reference to "__vbaFreeObj"
'00725b3e    8b3d08134000            mov edi, dword ptr [00401308]
'00725b44    8d4de0                  lea ecx, dword ptr [ebp-20]
'00725b47    ffd7                    call edi
    '#Cleanup(var_3)
'00725b49    8b06                    mov eax, dword ptr [esi]
'00725b4b    56                      push esi
'00725b4c    ff9008030000            call dword ptr [eax+00000308]
'00725b52    50                      push eax
'00725b53    8d4de0                  lea ecx, dword ptr [ebp-20]
'00725b56    51                      push ecx

' *** Reference to "__vbaObjSet"
'00725b57    ff15b4104000            call dword ptr [004010b4]
    Set var_3 = Nothing
'00725b5d    8bf0                    mov esi, eax
'00725b5f    8b16                    mov edx, dword ptr [esi]
'00725b61    56                      push esi
'00725b62    ff9204020000            call dword ptr [edx+00000204]
'00725b68    dbe2                    fnclex
'00725b6a    85c0                    test ax, ax
'00725b6c    0f8d09060000            jge 72617b
'00725b72    e9f2050000              jmp 726169
    
End If
'00725b77    8b06                    mov eax, dword ptr [esi]
'00725b79    56                      push esi
'00725b7a    ff9008030000            call dword ptr [eax+00000308]
'00725b80    50                      push eax
'00725b81    8d4de0                  lea ecx, dword ptr [ebp-20]
'00725b84    51                      push ecx
'00725b85    ffd3                    call ebx
Set var_3 = Nothing
'00725b87    8bf8                    mov edi, eax
'00725b89    8b17                    mov edx, dword ptr [edi]
'00725b8b    8d45e8                  lea eax, dword ptr [ebp-18]
'00725b8e    50                      push eax
'00725b8f    57                      push edi
'00725b90    ff92a0000000            call dword ptr [edx+000000a0]
'00725b96    dbe2                    fnclex
'00725b98    85c0                    test ax, ax
'00725b9a    7d12                    jge 725bae
'00725b9c    68a0000000              push 000000a0
'00725ba1    68d00d4300              push 00430dd0
'00725ba6    57                      push edi
'00725ba7    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00725ba8    ff1580104000            call dword ptr [00401080]
'00725bae    8b4de8                  mov ecx, dword ptr [ebp-18]
'00725bb1    51                      push ecx

' *** Reference to "rtcR8ValFromBstr"
'00725bb2    ff1510134000            call dword ptr [00401310]

' *** Reference to "__vbaFpR8"
'00725bb8    ff15f8104000            call dword ptr [004010f8]
'00725bbe    0fbf563c                movsx edx, word ptr [esi+3c]
'00725bc2    899544ffffff            mov dword ptr [ebp+ffffff44], edx
'00725bc8    db8544ffffff            fild dword ptr [ebp+ffffff44]
'00725bce    dd9d3cffffff            fstp qword ptr [ebp+ffffff3c]
'var_90 = (00)
'00725bd4    dc9d3cffffff            fcomp qword ptr [ebp+ffffff3c]
'00725bda    dfe0                    fnstsw ax
'00725bdc    f6c441                  test ah, 41
'00725bdf    7507                    jne 725be8
'00725be1    bf01000000              mov edi, 00000001
'00725be6    eb02                    jmp 725bea
'00725be8    33ff                    xor edi, edi
'00725bea    8d4de8                  lea ecx, dword ptr [ebp-18]

' *** Reference to "__vbaFreeStr"
'00725bed    ff150c134000            call dword ptr [0040130c]
'#Cleanup(var_41)
'00725bf3    8d4de0                  lea ecx, dword ptr [ebp-20]

' *** Reference to "__vbaFreeObj"
'00725bf6    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_3)
'00725bfc    f7df                    neg edi
'00725bfe    6685ff                  test edi, edi
'00725c01    0f84ba010000            je 725dc1

If (((arg_7) < Val(vbNullString))) Then
'00725c07    a124be7200              mov ax, word ptr [0072be24]
'00725c0c    85c0                    test ax, ax
'00725c0e    7510                    jne 725c20
'00725c10    6824be7200              push 0072be24
'00725c15    6870004300              push 00430070

' *** Reference to "__vbaNew2"
'00725c1a    ff152c124000            call dword ptr [0040122c]
    Set var_2 = New Global
'00725c20    8b3d24be7200            mov edi, dword ptr [0072be24]
'00725c26    8b07                    mov eax, dword ptr [edi]
'00725c28    8d4de0                  lea ecx, dword ptr [ebp-20]
'00725c2b    51                      push ecx
'00725c2c    57                      push edi
'00725c2d    ff5018                  call dword ptr [eax+18]
    Set var_3 = var_2.Screen
'00725c30    dbe2                    fnclex
'00725c32    85c0                    test ax, ax
'00725c34    7d0f                    jge 725c45
'00725c36    6a18                    push 18
'00725c38    6860004300              push 00430060
'00725c3d    57                      push edi
'00725c3e    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00725c3f    ff1580104000            call dword ptr [00401080]
'00725c45    8b7de0                  mov edi, dword ptr [ebp-20]
'00725c48    8b1f                    mov ebx, dword ptr [edi]
'00725c4a    33c9                    xor ecx, ecx

' *** Reference to "__vbaI2I4"
'00725c4c    ff1550114000            call dword ptr [00401150]
'00725c52    50                      push eax
'00725c53    57                      push edi
'00725c54    ff537c                  call dword ptr [ebx+7c]
    var_3.MousePointer = CInt(((arg_7) [#<] Val(vbNullString)))
'00725c57    dbe2                    fnclex
'00725c59    85c0                    test ax, ax
'00725c5b    7d0f                    jge 725c6c
'00725c5d    6a7c                    push 7c
'00725c5f    6810be4300              push 0043be10
'00725c64    57                      push edi
'00725c65    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00725c66    ff1580104000            call dword ptr [00401080]
'00725c6c    8d4de0                  lea ecx, dword ptr [ebp-20]

' *** Reference to "__vbaFreeObj"
'00725c6f    ff1508134000            call dword ptr [00401308]
    '#Cleanup(var_3)
'00725c75    8b16                    mov edx, dword ptr [esi]
'00725c77    8d45e4                  lea eax, dword ptr [ebp-1c]
'00725c7a    50                      push eax
'00725c7b    56                      push esi
'00725c7c    ff5248                  call dword ptr [edx+48]
'00725c7f    dbe2                    fnclex
'00725c81    85c0                    test ax, ax
'00725c83    7d0f                    jge 725c94
'00725c85    6a48                    push 48
'00725c87    6810f54400              push 0044f510
'00725c8c    56                      push esi
'00725c8d    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00725c8e    ff1580104000            call dword ptr [00401080]

' *** Reference to "__vbaStrI2"
'00725c94    8b1d10104000            mov ebx, dword ptr [00401010]
'00725c9a    b904000280              mov ecx, 80020004
'00725c9f    894da8                  mov dword ptr [ebp-58], ecx
'00725ca2    894db8                  mov dword ptr [ebp-48], ecx
'00725ca5    33c9                    xor ecx, ecx
'00725ca7    668b4e3c                mov cx, word ptr [esi+3c]
'00725cab    b80a000000              mov eax, 0000000a
'00725cb0    8945a0                  mov dword ptr [ebp-60], eax
'00725cb3    8945b0                  mov dword ptr [ebp-50], eax
'00725cb6    8b45e4                  mov eax, dword ptr [ebp-1c]
'00725cb9    68b47a4500              push 00457ab4
'00725cbe    bf08000000              mov edi, 00000008
'00725cc3    51                      push ecx
'00725cc4    c745e400000000          mov dword ptr [ebp-1c], 00000000
'00725ccb    8945c8                  mov dword ptr [ebp-38], eax
'00725cce    897dc0                  mov dword ptr [ebp-40], edi
'00725cd1    ffd3                    call ebx
'00725cd3    8bd0                    mov edx, eax
'00725cd5    8d4de8                  lea ecx, dword ptr [ebp-18]

' *** Reference to "__vbaStrMove"
'00725cd8    ff15d0124000            call dword ptr [004012d0]
'00725cde    50                      push eax

' *** Reference to "__vbaStrCat"
'00725cdf    ff1570104000            call dword ptr [00401070]
    var_129 = ("Le nombre d'objet doit être plus petit que ") & (CStr(arg_7))
'00725ce5    8d55a0                  lea edx, dword ptr [ebp-60]
'00725ce8    52                      push edx
'00725ce9    8945d8                  mov dword ptr [ebp-28], eax
'00725cec    8d45b0                  lea eax, dword ptr [ebp-50]
'00725cef    50                      push eax
'00725cf0    8d4dc0                  lea ecx, dword ptr [ebp-40]
'00725cf3    51                      push ecx
'00725cf4    6a30                    push 30
'00725cf6    8d55d0                  lea edx, dword ptr [ebp-30]
'00725cf9    52                      push edx
'00725cfa    897dd0                  mov dword ptr [ebp-30], edi

' *** Reference to "rtcMsgBox"
'00725cfd    ff15b8104000            call dword ptr [004010b8]
    var_285 = MsgBox(var_129, 48, frmVendre)
'00725d03    8d4de8                  lea ecx, dword ptr [ebp-18]

' *** Reference to "__vbaFreeStr"
'00725d06    ff150c134000            call dword ptr [0040130c]
    '#Cleanup(var_41)
'00725d0c    8d45a0                  lea eax, dword ptr [ebp-60]
'00725d0f    50                      push eax
'00725d10    8d4db0                  lea ecx, dword ptr [ebp-50]
'00725d13    51                      push ecx
'00725d14    8d55c0                  lea edx, dword ptr [ebp-40]
'00725d17    52                      push edx
'00725d18    8d45d0                  lea eax, dword ptr [ebp-30]
'00725d1b    50                      push eax
'00725d1c    6a04                    push 04

' *** Reference to "__vbaFreeVarList"
'00725d1e    ff1540104000            call dword ptr [00401040]
    '#Cleanup( var_4, var_5, var_6, var_7)
'00725d24    8b0e                    mov ecx, dword ptr [esi]
'00725d26    83c414                  add esp, 14
'00725d29    56                      push esi
'00725d2a    ff9108030000            call dword ptr [ecx+00000308]
'00725d30    50                      push eax
'00725d31    8d55e0                  lea edx, dword ptr [ebp-20]
'00725d34    52                      push edx

' *** Reference to "__vbaObjSet"
'00725d35    ff15b4104000            call dword ptr [004010b4]
    Set var_3 = 
'00725d3b    8bf8                    mov edi, eax
'00725d3d    8b17                    mov edx, dword ptr [edi]
'00725d3f    33c0                    xor eax, eax
'00725d41    668b463c                mov ax, word ptr [esi+3c]
'00725d45    899538ffffff            mov dword ptr [ebp+ffffff38], edx
'00725d4b    50                      push eax
'00725d4c    ffd3                    call ebx
'00725d4e    8bd0                    mov edx, eax
'00725d50    8d4de8                  lea ecx, dword ptr [ebp-18]

' *** Reference to "__vbaStrMove"
'00725d53    ff15d0124000            call dword ptr [004012d0]
'00725d59    8b8d38ffffff            mov ecx, dword ptr [ebp+ffffff38]
'00725d5f    50                      push eax
'00725d60    57                      push edi
'00725d61    ff91a4000000            call dword ptr [ecx+000000a4]
'00725d67    dbe2                    fnclex
'00725d69    85c0                    test ax, ax
'00725d6b    7d12                    jge 725d7f
'00725d6d    68a4000000              push 000000a4
'00725d72    68d00d4300              push 00430dd0
'00725d77    57                      push edi
'00725d78    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00725d79    ff1580104000            call dword ptr [00401080]
'00725d7f    8d4de8                  lea ecx, dword ptr [ebp-18]

' *** Reference to "__vbaFreeStr"
'00725d82    ff150c134000            call dword ptr [0040130c]
    '#Cleanup(var_41)

' *** Reference to "__vbaFreeObj"
'00725d88    8b3d08134000            mov edi, dword ptr [00401308]
'00725d8e    8d4de0                  lea ecx, dword ptr [ebp-20]
'00725d91    ffd7                    call edi
    '#Cleanup(var_3)
'00725d93    8b16                    mov edx, dword ptr [esi]
'00725d95    56                      push esi
'00725d96    ff9208030000            call dword ptr [edx+00000308]
'00725d9c    50                      push eax
'00725d9d    8d45e0                  lea eax, dword ptr [ebp-20]
'00725da0    50                      push eax

' *** Reference to "__vbaObjSet"
'00725da1    ff15b4104000            call dword ptr [004010b4]
    Set var_3 = -4560
'00725da7    8bf0                    mov esi, eax
'00725da9    8b0e                    mov ecx, dword ptr [esi]
'00725dab    56                      push esi
'00725dac    ff9104020000            call dword ptr [ecx+00000204]
'00725db2    dbe2                    fnclex
'00725db4    85c0                    test ax, ax
'00725db6    0f8dbf030000            jge 72617b
'00725dbc    e9a8030000              jmp 726169
    
End If
'00725dc1    8b16                    mov edx, dword ptr [esi]
'00725dc3    56                      push esi
'00725dc4    ff9208030000            call dword ptr [edx+00000308]
'00725dca    50                      push eax
'00725dcb    8d45e0                  lea eax, dword ptr [ebp-20]
'00725dce    50                      push eax
'00725dcf    ffd3                    call ebx
Set var_3 = ((arg_7) [:#] Val(vbNullString))
'00725dd1    8d55e8                  lea edx, dword ptr [ebp-18]
'00725dd4    8bf8                    mov edi, eax
'00725dd6    8b0f                    mov ecx, dword ptr [edi]
'00725dd8    52                      push edx
'00725dd9    57                      push edi
'00725dda    ff91a0000000            call dword ptr [ecx+000000a0]
'00725de0    dbe2                    fnclex
'00725de2    85c0                    test ax, ax
'00725de4    7d12                    jge 725df8
'00725de6    68a0000000              push 000000a0
'00725deb    68d00d4300              push 00430dd0
'00725df0    57                      push edi
'00725df1    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00725df2    ff1580104000            call dword ptr [00401080]
'00725df8    8b45e8                  mov eax, dword ptr [ebp-18]
'00725dfb    50                      push eax

' *** Reference to "rtcR8ValFromBstr"
'00725dfc    ff1510134000            call dword ptr [00401310]

' *** Reference to "__vbaFpR8"
'00725e02    ff15f8104000            call dword ptr [004010f8]
'00725e08    dc1d58164000            fcomp qword ptr [00401658]
'00725e0e    dfe0                    fnstsw ax
'00725e10    f6c401                  test ah, 01
'00725e13    7407                    je 725e1c
'00725e15    b801000000              mov eax, 00000001
'00725e1a    eb02                    jmp 725e1e
'00725e1c    33c0                    xor eax, eax
'00725e1e    f7d8                    neg eax
'00725e20    8d4de8                  lea ecx, dword ptr [ebp-18]
'00725e23    668bf8                  mov di, ax

' *** Reference to "__vbaFreeStr"
'00725e26    ff150c134000            call dword ptr [0040130c]
'#Cleanup(var_41)
'00725e2c    8d4de0                  lea ecx, dword ptr [ebp-20]

' *** Reference to "__vbaFreeObj"
'00725e2f    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_3)
'00725e35    6685ff                  test edi, edi
'00725e38    0f8493010000            je 725fd1

If ((1# > Val(vbNullString))) Then
'00725e3e    a124be7200              mov ax, word ptr [0072be24]
'00725e43    85c0                    test ax, ax
'00725e45    7510                    jne 725e57
'00725e47    6824be7200              push 0072be24
'00725e4c    6870004300              push 00430070

' *** Reference to "__vbaNew2"
'00725e51    ff152c124000            call dword ptr [0040122c]
    Set var_2 = New Global
'00725e57    8b3d24be7200            mov edi, dword ptr [0072be24]
'00725e5d    8b0f                    mov ecx, dword ptr [edi]
'00725e5f    8d55e0                  lea edx, dword ptr [ebp-20]
'00725e62    52                      push edx
'00725e63    57                      push edi
'00725e64    ff5118                  call dword ptr [ecx+18]
    Set var_3 = var_2.Screen
'00725e67    dbe2                    fnclex
'00725e69    85c0                    test ax, ax
'00725e6b    7d0f                    jge 725e7c
'00725e6d    6a18                    push 18
'00725e6f    6860004300              push 00430060
'00725e74    57                      push edi
'00725e75    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00725e76    ff1580104000            call dword ptr [00401080]
'00725e7c    8b7de0                  mov edi, dword ptr [ebp-20]
'00725e7f    8b1f                    mov ebx, dword ptr [edi]
'00725e81    33c9                    xor ecx, ecx

' *** Reference to "__vbaI2I4"
'00725e83    ff1550114000            call dword ptr [00401150]
'00725e89    50                      push eax
'00725e8a    57                      push edi
'00725e8b    ff537c                  call dword ptr [ebx+7c]
    var_3.MousePointer = CInt((1# [#>] Val(vbNullString)))
'00725e8e    dbe2                    fnclex
'00725e90    85c0                    test ax, ax
'00725e92    7d0f                    jge 725ea3
'00725e94    6a7c                    push 7c
'00725e96    6810be4300              push 0043be10
'00725e9b    57                      push edi
'00725e9c    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00725e9d    ff1580104000            call dword ptr [00401080]
'00725ea3    8d4de0                  lea ecx, dword ptr [ebp-20]

' *** Reference to "__vbaFreeObj"
'00725ea6    ff1508134000            call dword ptr [00401308]
    '#Cleanup(var_3)
'00725eac    8b06                    mov eax, dword ptr [esi]
'00725eae    8d4de8                  lea ecx, dword ptr [ebp-18]
'00725eb1    51                      push ecx
'00725eb2    56                      push esi
'00725eb3    ff5048                  call dword ptr [eax+48]
'00725eb6    dbe2                    fnclex
'00725eb8    85c0                    test ax, ax
'00725eba    7d0f                    jge 725ecb
'00725ebc    6a48                    push 48
'00725ebe    6810f54400              push 0044f510
'00725ec3    56                      push esi
'00725ec4    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00725ec5    ff1580104000            call dword ptr [00401080]
'00725ecb    b80a000000              mov eax, 0000000a
'00725ed0    b904000280              mov ecx, 80020004
'00725ed5    8945a0                  mov dword ptr [ebp-60], eax
'00725ed8    8945b0                  mov dword ptr [ebp-50], eax
'00725edb    8b45e8                  mov eax, dword ptr [ebp-18]
'00725ede    894da8                  mov dword ptr [ebp-58], ecx
'00725ee1    894db8                  mov dword ptr [ebp-48], ecx
'00725ee4    bf08000000              mov edi, 00000008
'00725ee9    8d5590                  lea edx, dword ptr [ebp-70]
'00725eec    8d4dd0                  lea ecx, dword ptr [ebp-30]
'00725eef    c745e800000000          mov dword ptr [ebp-18], 00000000
'00725ef6    8945c8                  mov dword ptr [ebp-38], eax
'00725ef9    897dc0                  mov dword ptr [ebp-40], edi
'00725efc    c74598107b4500          mov dword ptr [ebp-68], 00457b10
'00725f03    897d90                  mov dword ptr [ebp-70], edi

' *** Reference to "__vbaVarDup"
'00725f06    ff1598124000            call dword ptr [00401298]
    var_4 = ("Le nombre d'objet doit être plus grand que 1")
'00725f0c    8d55a0                  lea edx, dword ptr [ebp-60]
'00725f0f    52                      push edx
'00725f10    8d45b0                  lea eax, dword ptr [ebp-50]
'00725f13    50                      push eax
'00725f14    8d4dc0                  lea ecx, dword ptr [ebp-40]
'00725f17    51                      push ecx
'00725f18    6a30                    push 30
'00725f1a    8d55d0                  lea edx, dword ptr [ebp-30]
'00725f1d    52                      push edx

' *** Reference to "rtcMsgBox"
'00725f1e    ff15b8104000            call dword ptr [004010b8]
    var_1231 = MsgBox(var_4, 48, vbNullString)
'00725f24    8d45a0                  lea eax, dword ptr [ebp-60]
'00725f27    50                      push eax
'00725f28    8d4db0                  lea ecx, dword ptr [ebp-50]
'00725f2b    51                      push ecx
'00725f2c    8d55c0                  lea edx, dword ptr [ebp-40]
'00725f2f    52                      push edx
'00725f30    8d45d0                  lea eax, dword ptr [ebp-30]
'00725f33    50                      push eax
'00725f34    6a04                    push 04

' *** Reference to "__vbaFreeVarList"
'00725f36    ff1540104000            call dword ptr [00401040]
    '#Cleanup( var_4, var_5, var_6, var_7)
'00725f3c    8b0e                    mov ecx, dword ptr [esi]
'00725f3e    83c414                  add esp, 14
'00725f41    56                      push esi
'00725f42    ff9108030000            call dword ptr [ecx+00000308]
'00725f48    50                      push eax
'00725f49    8d55e0                  lea edx, dword ptr [ebp-20]
'00725f4c    52                      push edx

' *** Reference to "__vbaObjSet"
'00725f4d    ff15b4104000            call dword ptr [004010b4]
    Set var_3 = 
'00725f53    8bf8                    mov edi, eax
'00725f55    8b1f                    mov ebx, dword ptr [edi]
'00725f57    33c0                    xor eax, eax
'00725f59    668b463c                mov ax, word ptr [esi+3c]
'00725f5d    50                      push eax

' *** Reference to "__vbaStrI2"
'00725f5e    ff1510104000            call dword ptr [00401010]
'00725f64    8bd0                    mov edx, eax
'00725f66    8d4de8                  lea ecx, dword ptr [ebp-18]

' *** Reference to "__vbaStrMove"
'00725f69    ff15d0124000            call dword ptr [004012d0]
'00725f6f    50                      push eax
'00725f70    57                      push edi
'00725f71    ff93a4000000            call dword ptr [ebx+000000a4]
'00725f77    dbe2                    fnclex
'00725f79    85c0                    test ax, ax
'00725f7b    7d12                    jge 725f8f
'00725f7d    68a4000000              push 000000a4
'00725f82    68d00d4300              push 00430dd0
'00725f87    57                      push edi
'00725f88    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00725f89    ff1580104000            call dword ptr [00401080]
'00725f8f    8d4de8                  lea ecx, dword ptr [ebp-18]

' *** Reference to "__vbaFreeStr"
'00725f92    ff150c134000            call dword ptr [0040130c]
    '#Cleanup(var_41)

' *** Reference to "__vbaFreeObj"
'00725f98    8b3d08134000            mov edi, dword ptr [00401308]
'00725f9e    8d4de0                  lea ecx, dword ptr [ebp-20]
'00725fa1    ffd7                    call edi
    '#Cleanup(var_3)
'00725fa3    8b0e                    mov ecx, dword ptr [esi]
'00725fa5    56                      push esi
'00725fa6    ff9108030000            call dword ptr [ecx+00000308]
'00725fac    50                      push eax
'00725fad    8d55e0                  lea edx, dword ptr [ebp-20]
'00725fb0    52                      push edx

' *** Reference to "__vbaObjSet"
'00725fb1    ff15b4104000            call dword ptr [004010b4]
    Set var_3 = -4588
'00725fb7    8bf0                    mov esi, eax
'00725fb9    8b06                    mov eax, dword ptr [esi]
'00725fbb    56                      push esi
'00725fbc    ff9004020000            call dword ptr [eax+00000204]
'00725fc2    dbe2                    fnclex
'00725fc4    85c0                    test ax, ax
'00725fc6    0f8daf010000            jge 72617b
'00725fcc    e998010000              jmp 726169
    
End If
'00725fd1    8b0e                    mov ecx, dword ptr [esi]
'00725fd3    56                      push esi
'00725fd4    ff912c070000            call dword ptr [ecx+0000072c]
'00725fda    e9a1010000              jmp 726180

'ERROR: Two many next close:
End If
'00725fdf    a124be7200              mov ax, word ptr [0072be24]
'00725fe4    85c0                    test ax, ax
'00725fe6    7510                    jne 725ff8
'00725fe8    6824be7200              push 0072be24
'00725fed    6870004300              push 00430070

' *** Reference to "__vbaNew2"
'00725ff2    ff152c124000            call dword ptr [0040122c]
Set var_2 = New Global
'00725ff8    8b3d24be7200            mov edi, dword ptr [0072be24]
'00725ffe    8b17                    mov edx, dword ptr [edi]
'00726000    8d45e0                  lea eax, dword ptr [ebp-20]
'00726003    50                      push eax
'00726004    57                      push edi
'00726005    ff5218                  call dword ptr [edx+18]
Set var_3 = var_2.Screen
'00726008    dbe2                    fnclex
'0072600a    85c0                    test ax, ax
'0072600c    7d0f                    jge 72601d
'0072600e    6a18                    push 18
'00726010    6860004300              push 00430060
'00726015    57                      push edi
'00726016    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00726017    ff1580104000            call dword ptr [00401080]
'0072601d    8b7de0                  mov edi, dword ptr [ebp-20]
'00726020    8b1f                    mov ebx, dword ptr [edi]
'00726022    33c9                    xor ecx, ecx

' *** Reference to "__vbaI2I4"
'00726024    ff1550114000            call dword ptr [00401150]
'0072602a    50                      push eax
'0072602b    57                      push edi
'0072602c    ff537c                  call dword ptr [ebx+7c]
var_3.MousePointer = CInt((1# [#>] Val(vbNullString)))
'0072602f    dbe2                    fnclex
'00726031    85c0                    test ax, ax
'00726033    7d0f                    jge 726044
'00726035    6a7c                    push 7c
'00726037    6810be4300              push 0043be10
'0072603c    57                      push edi
'0072603d    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0072603e    ff1580104000            call dword ptr [00401080]
'00726044    8d4de0                  lea ecx, dword ptr [ebp-20]

' *** Reference to "__vbaFreeObj"
'00726047    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_3)
'0072604d    8b0e                    mov ecx, dword ptr [esi]
'0072604f    8d55e8                  lea edx, dword ptr [ebp-18]
'00726052    52                      push edx
'00726053    56                      push esi
'00726054    ff5148                  call dword ptr [ecx+48]
'00726057    dbe2                    fnclex
'00726059    85c0                    test ax, ax
'0072605b    7d0f                    jge 72606c
'0072605d    6a48                    push 48
'0072605f    6810f54400              push 0044f510
'00726064    56                      push esi
'00726065    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00726066    ff1580104000            call dword ptr [00401080]
'0072606c    b80a000000              mov eax, 0000000a
'00726071    b904000280              mov ecx, 80020004
'00726076    8945a0                  mov dword ptr [ebp-60], eax
'00726079    8945b0                  mov dword ptr [ebp-50], eax
'0072607c    8b45e8                  mov eax, dword ptr [ebp-18]
'0072607f    894da8                  mov dword ptr [ebp-58], ecx
'00726082    894db8                  mov dword ptr [ebp-48], ecx
'00726085    bf08000000              mov edi, 00000008
'0072608a    8d5590                  lea edx, dword ptr [ebp-70]
'0072608d    8d4dd0                  lea ecx, dword ptr [ebp-30]
'00726090    c745e800000000          mov dword ptr [ebp-18], 00000000
'00726097    8945c8                  mov dword ptr [ebp-38], eax
'0072609a    897dc0                  mov dword ptr [ebp-40], edi
'0072609d    c745987c244500          mov dword ptr [ebp-68], 0045247c
'007260a4    897d90                  mov dword ptr [ebp-70], edi

' *** Reference to "__vbaVarDup"
'007260a7    ff1598124000            call dword ptr [00401298]
var_4 = ("Le nombre d'objet doit être entrée en tant que valeur numérique")
'007260ad    8d45a0                  lea eax, dword ptr [ebp-60]
'007260b0    50                      push eax
'007260b1    8d4db0                  lea ecx, dword ptr [ebp-50]
'007260b4    51                      push ecx
'007260b5    8d55c0                  lea edx, dword ptr [ebp-40]
'007260b8    52                      push edx
'007260b9    6a30                    push 30
'007260bb    8d45d0                  lea eax, dword ptr [ebp-30]
'007260be    50                      push eax

' *** Reference to "rtcMsgBox"
'007260bf    ff15b8104000            call dword ptr [004010b8]
var_870 = MsgBox(var_4, 48, vbNullString)
'007260c5    8d4da0                  lea ecx, dword ptr [ebp-60]
'007260c8    51                      push ecx
'007260c9    8d55b0                  lea edx, dword ptr [ebp-50]
'007260cc    52                      push edx
'007260cd    8d45c0                  lea eax, dword ptr [ebp-40]
'007260d0    50                      push eax
'007260d1    8d4dd0                  lea ecx, dword ptr [ebp-30]
'007260d4    51                      push ecx
'007260d5    6a04                    push 04

' *** Reference to "__vbaFreeVarList"
'007260d7    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_4, var_5, var_6, var_7)
'007260dd    8b16                    mov edx, dword ptr [esi]
'007260df    83c414                  add esp, 14
'007260e2    56                      push esi
'007260e3    ff9208030000            call dword ptr [edx+00000308]
'007260e9    50                      push eax
'007260ea    8d45e0                  lea eax, dword ptr [ebp-20]
'007260ed    50                      push eax

' *** Reference to "__vbaObjSet"
'007260ee    ff15b4104000            call dword ptr [004010b4]
Set var_3 = 
'007260f4    33c9                    xor ecx, ecx
'007260f6    668b4e3c                mov cx, word ptr [esi+3c]
'007260fa    8bf8                    mov edi, eax
'007260fc    8b1f                    mov ebx, dword ptr [edi]
'007260fe    51                      push ecx

' *** Reference to "__vbaStrI2"
'007260ff    ff1510104000            call dword ptr [00401010]
'00726105    8bd0                    mov edx, eax
'00726107    8d4de8                  lea ecx, dword ptr [ebp-18]

' *** Reference to "__vbaStrMove"
'0072610a    ff15d0124000            call dword ptr [004012d0]
'00726110    50                      push eax
'00726111    57                      push edi
'00726112    ff93a4000000            call dword ptr [ebx+000000a4]
'00726118    dbe2                    fnclex
'0072611a    85c0                    test ax, ax
'0072611c    7d12                    jge 726130
'0072611e    68a4000000              push 000000a4
'00726123    68d00d4300              push 00430dd0
'00726128    57                      push edi
'00726129    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0072612a    ff1580104000            call dword ptr [00401080]
'00726130    8d4de8                  lea ecx, dword ptr [ebp-18]

' *** Reference to "__vbaFreeStr"
'00726133    ff150c134000            call dword ptr [0040130c]
'#Cleanup(var_41)

' *** Reference to "__vbaFreeObj"
'00726139    8b3d08134000            mov edi, dword ptr [00401308]
'0072613f    8d4de0                  lea ecx, dword ptr [ebp-20]
'00726142    ffd7                    call edi
'#Cleanup(var_3)
'00726144    8b16                    mov edx, dword ptr [esi]
'00726146    56                      push esi
'00726147    ff9208030000            call dword ptr [edx+00000308]
'0072614d    50                      push eax
'0072614e    8d45e0                  lea eax, dword ptr [ebp-20]
'00726151    50                      push eax

' *** Reference to "__vbaObjSet"
'00726152    ff15b4104000            call dword ptr [004010b4]
Set var_3 = -4616
'00726158    8bf0                    mov esi, eax
'0072615a    8b0e                    mov ecx, dword ptr [esi]
'0072615c    56                      push esi
'0072615d    ff9104020000            call dword ptr [ecx+00000204]
'00726163    dbe2                    fnclex
'00726165    85c0                    test ax, ax
'00726167    7d12                    jge 72617b
'00726169    6804020000              push 00000204
'0072616e    68d00d4300              push 00430dd0
'00726173    56                      push esi
'00726174    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00726175    ff1580104000            call dword ptr [00401080]
'0072617b    8d4de0                  lea ecx, dword ptr [ebp-20]
'0072617e    ffd7                    call edi
'#Cleanup(var_3)
'00726180    c745fc00000000          mov dword ptr [ebp-04], 00000000
'00726187    9b                      fwait
'00726188    68c8617200              push 007261c8
'0072618d    eb38                    jmp 7261c7
'0072618f    8d55e4                  lea edx, dword ptr [ebp-1c]
'00726192    52                      push edx
'00726193    8d45e8                  lea eax, dword ptr [ebp-18]
'00726196    50                      push eax
'00726197    6a02                    push 02

' *** Reference to "__vbaFreeStrList"
'00726199    ff155c124000            call dword ptr [0040125c]
'#Cleanup( , var_40)
'0072619f    83c40c                  add esp, 0c
'007261a2    8d4de0                  lea ecx, dword ptr [ebp-20]

' *** Reference to "__vbaFreeObj"
'007261a5    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_3)
'007261ab    8d4da0                  lea ecx, dword ptr [ebp-60]
'007261ae    51                      push ecx
'007261af    8d55b0                  lea edx, dword ptr [ebp-50]
'007261b2    52                      push edx
'007261b3    8d45c0                  lea eax, dword ptr [ebp-40]
'007261b6    50                      push eax
'007261b7    8d4dd0                  lea ecx, dword ptr [ebp-30]
'007261ba    51                      push ecx
'007261bb    6a04                    push 04

' *** Reference to "__vbaFreeVarList"
'007261bd    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_4, var_5, var_6, var_7)
'007261c3    83c414                  add esp, 14
'007261c6    c3                      ret
'007261c7    c3                      ret
'007261c8    8b4508                  mov eax, dword ptr [ebp+08]
'007261cb    8b10                    mov edx, dword ptr [eax]
'007261cd    50                      push eax
'007261ce    ff5208                  call dword ptr [edx+08]
'007261d1    8b45fc                  mov eax, dword ptr [ebp-04]
'007261d4    8b4dec                  mov ecx, dword ptr [ebp-14]
'007261d7    5f                      pop edi
'007261d8    5e                      pop esi
    'Reference to '__except_list'
'007261d9    64890d00000000          mov dword ptr fs:[00000000], ecx
'007261e0    5b                      pop ebx
'007261e1    8be5                    mov esp, ebp
'007261e3    5d                      pop ebp
'007261e4    c20400                  ret 0004


End Sub


