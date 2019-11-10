VERSION 5.00

Begin VB.Form frmAstuce
    Caption = "Astuce du jour"
    ScaleMode = 1
    AutoRedraw = 0              'False
    FontTransparent = -1              'True
    LinkTopic = "Form1"
    MaxButton = 0              'False
    MinButton = 0              'False
    ClientLeft   = 2370
    ClientTop    = 2400
    ClientWidth  = 6585
    ClientHeight = 4050
    BeginProperty Font
        Name          = "Times New Roman"
        Size          = 9
        Charset       = 0
        Weight        = 400
        Underline     = 0              'False
        Italic        = 0              'False
        Strikethrough = 0              'False
    EndProperty
    WhatsThisButton = 255
    WhatsThisHelp = 255
    Begin VB.CheckBox chkLoadTipsAtStartup
        Caption = "&Afficher les astuces au démarrage"
        Left   = 120
        Top    = 3600
        Width  = 2895
        Height = 315
        TabIndex = 3
    End
    Begin VB.CommandButton cmdNextTip
        Caption = "Astuce &suivante"
        Left   = 5160
        Top    = 600
        Width  = 1335
        Height = 375
        TabIndex = 2
    End
    Begin VB.PictureBox Picture1
        BackColor = 16777215
        Picture = frmAstuce.frx:0000
        Left   = 120
        Top    = 120
        Width  = 4935
        Height = 3435
        TabIndex = 1
        ScaleMode = 1
        AutoRedraw = 0              'False
        FontTransparent = -1              'True
        Begin VB.Label Label1
            Caption = "Savez-vous..."
            BackColor = 16777215
            Left   = 540
            Top    = 180
            Width  = 2655
            Height = 255
            TabIndex = 5
        End
        Begin VB.Label lblTipText
            BackColor = 16777215
            Left   = 120
            Top    = 840
            Width  = 4575
            Height = 2475
            TabIndex = 4
        End
    End
    Begin VB.CommandButton cmdOK
        Caption = "OK"
        Left   = 5160
        Top    = 120
        Width  = 1335
        Height = 375
        TabIndex = 0
        Default = -1              'True
        Cancel = -1              'True
    End
End
Private Function sub_71E720(arg_0 As Unknow, arg_1 As Unknow, arg_2 As Unknow, arg_3 As Unknow, arg_4 As Unknow, arg_5 As Unknow, arg_6 As Unknow, arg_7 As Unknow, arg_8 As Unknow, arg_9 As Unknow, arg_A As Unknow, arg_B As Unknow, arg_C As Unknow, arg_D As Unknow, arg_E As Unknow, arg_F As Unknow, arg_10 As Unknow, arg_11 As Unknow, arg_12 As Unknow, arg_13 As Unknow, arg_14 As Unknow, arg_15 As Unknow, arg_16 As Unknow, arg_17 As Unknow, arg_18 As Unknow, arg_19 As Unknow, arg_1A As Unknow, arg_1B As Unknow, arg_1C As Unknow, arg_1D As Unknow, arg_1E As Unknow, arg_1F As Unknow, arg_20 As Unknow, arg_21 As Unknow, arg_22 As Unknow, arg_23 As Unknow, arg_24 As Unknow, arg_25 As Unknow, arg_26 As Unknow, arg_27 As Unknow, arg_28 As Unknow, arg_29 As Unknow, arg_2A As Unknow, arg_2B As Unknow, arg_2C As Unknow, arg_2D As Unknow, arg_2E As Unknow, arg_2F As Unknow, arg_30 As Unknow, arg_31 As Unknow, arg_32 As Unknow, arg_33 As Unknow, arg_34 As Unknow, arg_35 As Unknow, arg_36 As Unknow, arg_37 As Unknow, arg_38 As Unknow, arg_39 As Unknow, arg_3A As Unknow, arg_3B As Unknow, arg_3C As Unknow)
'0071e720    51                      push ecx
'0071e721    53                      push ebx
'0071e722    55                      push ebp

' *** Reference to "__vbaNew2"
'0071e723    8b2d2c124000            mov ebp, dword ptr [0040122c]
'0071e729    56                      push esi
'0071e72a    57                      push edi
'0071e72b    8b7c2418                mov edi, dword ptr [esp+18]
'0071e72f    8b4738                  mov eax, dword ptr [edi+38]
'0071e732    83c001                  add eax, 01
var_num1 = arg_6 + 1
'0071e735    0f8091000000            jo 71e7cc
'0071e73b    8d7734                  lea esi, dword ptr [edi+34]
'0071e73e    894738                  mov dword ptr [edi+38], eax
'0071e741    833e00                  cmp dword ptr [esi], 00
'0071e744    c744241000000000        mov dword ptr [esp+10], 00000000
'0071e74c    7508                    jne 71e756

If (arg_6 = 0) Then
'0071e74e    56                      push esi
'0071e74f    688c6e4500              push 00456e8c
'0071e754    ffd5                    call ebp
    Dim var_138 As New Collection
End If
'0071e756    8b36                    mov esi, dword ptr [esi]
'0071e758    8b0e                    mov ecx, dword ptr [esi]
'0071e75a    8d542410                lea edx, dword ptr [esp+10]
'0071e75e    52                      push edx
'0071e75f    56                      push esi
'0071e760    ff5124                  call dword ptr [ecx+24]
var_2584 = var_138.Count()
'0071e763    dbe2                    fnclex
'0071e765    85c0                    test ax, ax

' *** Reference to "__vbaHresultCheckObj"
'0071e767    8b1d80104000            mov ebx, dword ptr [00401080]
'0071e76d    7d0b                    jge 71e77a
'0071e76f    6a24                    push 24
'0071e771    687c6e4500              push 00456e7c
'0071e776    56                      push esi
'0071e777    50                      push eax
'0071e778    ffd3                    call ebx
'0071e77a    8b442410                mov eax, dword ptr [esp+10]
'0071e77e    3b4738                  cmp eax, dword ptr [edi+38]
'0071e781    7d07                    jge 71e78a

If (var_2584 < DWORD PTR [EDI+38]) Then
'0071e783    c7473801000000          mov dword ptr [edi+38], 00000001
    
End If
'0071e78a    a1acb07200              mov ax, word ptr [0072b0ac]
'0071e78f    85c0                    test ax, ax
'0071e791    750c                    jne 71e79f
'0071e793    68acb07200              push 0072b0ac
'0071e798    68a4914000              push 004091a4
'0071e79d    ffd5                    call ebp
Dim var_30 As New frmAstuce
'0071e79f    8b35acb07200            mov esi, dword ptr [0072b0ac]
'0071e7a5    8b0e                    mov ecx, dword ptr [esi]
'0071e7a7    56                      push esi
'0071e7a8    ff91fc060000            call dword ptr [ecx+000006fc]
'0071e7ae    dbe2                    fnclex
'0071e7b0    85c0                    test ax, ax
'0071e7b2    7d0e                    jge 71e7c2
'0071e7b4    68fc060000              push 000006fc
'0071e7b9    6854104300              push 00431054
'0071e7be    56                      push esi
'0071e7bf    50                      push eax
'0071e7c0    ffd3                    call ebx
'0071e7c2    5f                      pop edi
'0071e7c3    5e                      pop esi
'0071e7c4    5d                      pop ebp
'0071e7c5    33c0                    xor eax, eax
var_num1 = Empty
'0071e7c7    5b                      pop ebx
'0071e7c8    59                      pop ecx
'0071e7c9    c20400                  ret 0004


End Function


Public Function LoadTips(arg_0 As Unknow, arg_1 As Unknow, arg_2 As Unknow, arg_3 As Unknow, arg_4 As Unknow, arg_5 As Unknow, arg_6 As Unknow, arg_7 As Unknow, arg_8 As Unknow, arg_9 As Unknow, arg_A As Unknow, arg_B As Unknow, arg_C As Unknow, arg_D As Unknow, arg_E As Unknow, arg_F As Unknow, arg_10 As Unknow, arg_11 As Unknow, arg_12 As Unknow, arg_13 As Unknow, arg_14 As Unknow, arg_15 As Unknow, arg_16 As Unknow, arg_17 As Unknow, arg_18 As Unknow, arg_19 As Unknow, arg_1A As Unknow, arg_1B As Unknow, arg_1C As Unknow, arg_1D As Unknow, arg_1E As Unknow, arg_1F As Unknow, arg_20 As Unknow, arg_21 As Unknow, arg_22 As Unknow, arg_23 As Unknow, arg_24 As Unknow, arg_25 As Unknow, arg_26 As Unknow, arg_27 As Unknow, arg_28 As Unknow, arg_29 As Unknow, arg_2A As Unknow, arg_2B As Unknow, arg_2C As Unknow, arg_2D As Unknow, arg_2E As Unknow, arg_2F As Unknow, arg_30 As Unknow, arg_31 As Unknow, arg_32 As Unknow, arg_33 As Unknow, arg_34 As Unknow, arg_35 As Unknow, arg_36 As Unknow, arg_37 As Unknow, arg_38 As Unknow, arg_39 As Unknow, arg_3A As Unknow, arg_3B As Unknow, arg_3C As Unknow)
'0071e7e0    55                      push ebp
'0071e7e1    8bec                    mov ebp, esp
'0071e7e3    83ec0c                  sub esp, 0c

' *** Reference to "__vbaExceptHandler"
'0071e7e6    6866724000              push 00407266
'0071e7eb    64a100000000            mov ax, word ptr fs:[00000000]
'0071e7f1    50                      push eax
    'Reference to '__except_list'
'0071e7f2    64892500000000          mov dword ptr fs:[00000000], esp
'0071e7f9    81ec90000000            sub esp, 00000090
'0071e7ff    53                      push ebx
'0071e800    56                      push esi
'0071e801    57                      push edi
'0071e802    8965f4                  mov dword ptr [ebp-0c], esp
'0071e805    c745f8506f4000          mov dword ptr [ebp-08], 00406f50
'0071e80c    33db                    xor ebx, ebx
'0071e80e    895dfc                  mov dword ptr [ebp-04], ebx
'0071e811    8b4508                  mov eax, dword ptr [ebp+08]
'0071e814    8b08                    mov ecx, dword ptr [eax]
'0071e816    50                      push eax
'0071e817    ff5104                  call dword ptr [ecx+04]
'0071e81a    8d55cc                  lea edx, dword ptr [ebp-34]
'0071e81d    895dcc                  mov dword ptr [ebp-34], ebx
'0071e820    52                      push edx
'0071e821    895de8                  mov dword ptr [ebp-18], ebx
'0071e824    895de4                  mov dword ptr [ebp-1c], ebx
'0071e827    895ddc                  mov dword ptr [ebp-24], ebx
'0071e82a    895dbc                  mov dword ptr [ebp-44], ebx
'0071e82d    895dac                  mov dword ptr [ebp-54], ebx
'0071e830    895d9c                  mov dword ptr [ebp-64], ebx
'0071e833    c745d404000280          mov dword ptr [ebp-2c], 80020004
'0071e83a    c745cc0a000000          mov dword ptr [ebp-34], 0000000a

' *** Reference to "rtcFreeFile"
'0071e841    ff1528124000            call dword ptr [00401228]
var_num1 = FreeFile()
'0071e847    8d4dcc                  lea ecx, dword ptr [ebp-34]
'0071e84a    8945e0                  mov dword ptr [ebp-20], eax

' *** Reference to "__vbaFreeVar"
'0071e84d    ff1528104000            call dword ptr [00401028]
'#Cleanup(var_43)
'0071e853    8b7d0c                  mov edi, dword ptr [ebp+0c]
'0071e856    8b07                    mov eax, dword ptr [edi]

' *** Reference to "__vbaStrCmp"
'0071e858    8b3530114000            mov esi, dword ptr [00401130]
'0071e85e    50                      push eax
'0071e85f    68cc134300              push 004313cc
'0071e864    ffd6                    call esi
'0071e866    85c0                    test ax, ax
'0071e868    7508                    jne 71e872

If (((arg_0) = (vbNullChar))) Then
'0071e86a    895de4                  mov dword ptr [ebp-1c], ebx
'0071e86d    e917010000              jmp 71e989
    
End If
'0071e872    53                      push ebx
'0071e873    8d4d9c                  lea ecx, dword ptr [ebp-64]
'0071e876    51                      push ecx
'0071e877    897da4                  mov dword ptr [ebp-5c], edi
'0071e87a    c7459c08400000          mov dword ptr [ebp-64], 00004008

' *** Reference to "rtcDir"
'0071e881    ff1510124000            call dword ptr [00401210]
'0071e887    8bd0                    mov edx, eax
'0071e889    8d4ddc                  lea ecx, dword ptr [ebp-24]

' *** Reference to "__vbaStrMove"
'0071e88c    ff15d0124000            call dword ptr [004012d0]
'0071e892    50                      push eax
'0071e893    68cc134300              push 004313cc
'0071e898    ffd6                    call esi
'0071e89a    8bf0                    mov esi, eax
'0071e89c    f7de                    neg esi
'0071e89e    1bf6                    sbb esi, esi
'0071e8a0    46                      inc esi
'0071e8a1    8d4ddc                  lea ecx, dword ptr [ebp-24]
'0071e8a4    f7de                    neg esi

' *** Reference to "__vbaFreeStr"
'0071e8a6    ff150c134000            call dword ptr [0040130c]
'#Cleanup(var_39)
'0071e8ac    663bf3                  cmp si, bx
'0071e8af    7408                    je 71e8b9

If (((Dir(arg_0, 0)) = (vbNullChar))) Then
'0071e8b1    895de4                  mov dword ptr [ebp-1c], ebx
'0071e8b4    e9d0000000              jmp 71e989
    
End If
'0071e8b9    8b17                    mov edx, dword ptr [edi]
'0071e8bb    8b7de0                  mov edi, dword ptr [ebp-20]
'0071e8be    52                      push edx
'0071e8bf    57                      push edi
'0071e8c0    6aff                    push ffffffff
'0071e8c2    6a01                    push 01

' *** Reference to "__vbaFileOpen"
'0071e8c4    ff151c124000            call dword ptr [0040121c]
Open arg_0 For Input As var_num1 Len = -1
'0071e8ca    57                      push edi

' *** Reference to "rtcEndOfFile"
'0071e8cb    ff1534124000            call dword ptr [00401234]
'0071e8d1    6685c0                  test eax, eax
'0071e8d4    57                      push edi
'0071e8d5    0f8595000000            jne 71e970

Do While (EOF(var_num1))
'0071e8db    8d45e8                  lea eax, dword ptr [ebp-18]
'0071e8de    50                      push eax

' *** Reference to "__vbaLineInputStr"
'0071e8df    ff152c104000            call dword ptr [0040102c]
'0071e8e5    8b4d08                  mov ecx, dword ptr [ebp+08]
'0071e8e8    8b4134                  mov eax, dword ptr [ecx+34]
'0071e8eb    3bc3                    cmp eax, ebx
'0071e8ed    8d7134                  lea esi, dword ptr [ecx+34]
'0071e8f0    750c                    jne 71e8fe
    
    If (    arg_6 = 0) Then
'0071e8f2    56                      push esi
'0071e8f3    688c6e4500              push 00456e8c

' *** Reference to "__vbaNew2"
'0071e8f8    ff152c124000            call dword ptr [0040122c]
    Dim var_138 As New Collection
End If
'0071e8fe    8b36                    mov esi, dword ptr [esi]
'0071e900    b90a000000              mov ecx, 0000000a
'0071e905    894dac                  mov dword ptr [ebp-54], ecx
'0071e908    894dbc                  mov dword ptr [ebp-44], ecx
'0071e90b    894dcc                  mov dword ptr [ebp-34], ecx
'0071e90e    8d55e8                  lea edx, dword ptr [ebp-18]
'0071e911    8955a4                  mov dword ptr [ebp-5c], edx
'0071e914    8d4dac                  lea ecx, dword ptr [ebp-54]
'0071e917    51                      push ecx
'0071e918    8d55bc                  lea edx, dword ptr [ebp-44]
'0071e91b    52                      push edx
'0071e91c    8d4dcc                  lea ecx, dword ptr [ebp-34]
'0071e91f    b804000280              mov eax, 80020004
'0071e924    51                      push ecx
'0071e925    8d559c                  lea edx, dword ptr [ebp-64]
'0071e928    52                      push edx
'0071e929    8945b4                  mov dword ptr [ebp-4c], eax
'0071e92c    8945c4                  mov dword ptr [ebp-3c], eax
'0071e92f    8945d4                  mov dword ptr [ebp-2c], eax
'0071e932    c7459c08400000          mov dword ptr [ebp-64], 00004008
'0071e939    8b06                    mov eax, dword ptr [esi]
'0071e93b    56                      push esi
'0071e93c    ff5020                  call dword ptr [eax+20]
Call var_138.Add(vbNullString)
'0071e93f    dbe2                    fnclex
'0071e941    3bc3                    cmp eax, ebx
'0071e943    7d0f                    jge 71e954
'0071e945    6a20                    push 20
'0071e947    687c6e4500              push 00456e7c
'0071e94c    56                      push esi
'0071e94d    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0071e94e    ff1580104000            call dword ptr [00401080]
'0071e954    8d45ac                  lea eax, dword ptr [ebp-54]
'0071e957    50                      push eax
'0071e958    8d4dbc                  lea ecx, dword ptr [ebp-44]
'0071e95b    51                      push ecx
'0071e95c    8d55cc                  lea edx, dword ptr [ebp-34]
'0071e95f    52                      push edx
'0071e960    6a03                    push 03

' *** Reference to "__vbaFreeVarList"
'0071e962    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_43, var_58, var_50)
'0071e968    83c410                  add esp, 10
'0071e96b    e95affffff              jmp 71e8ca

'ERROR: Two many next close:
Loop

' *** Reference to "__vbaFileClose"
'0071e970    ff1520114000            call dword ptr [00401120]
Close var_num1
'0071e976    8b4508                  mov eax, dword ptr [ebp+08]
'0071e979    8b08                    mov ecx, dword ptr [eax]
'0071e97b    50                      push eax
'0071e97c    ff9100070000            call dword ptr [ecx+00000700]
'0071e982    c745e4ffffffff          mov dword ptr [ebp-1c], ffffffff
'0071e989    68bbe97100              push 0071e9bb
'0071e98e    eb21                    jmp 71e9b1
'0071e990    8d4ddc                  lea ecx, dword ptr [ebp-24]

' *** Reference to "__vbaFreeStr"
'0071e993    ff150c134000            call dword ptr [0040130c]
'#Cleanup(var_39)
'0071e999    8d55ac                  lea edx, dword ptr [ebp-54]
'0071e99c    52                      push edx
'0071e99d    8d45bc                  lea eax, dword ptr [ebp-44]
'0071e9a0    50                      push eax
'0071e9a1    8d4dcc                  lea ecx, dword ptr [ebp-34]
'0071e9a4    51                      push ecx
'0071e9a5    6a03                    push 03

' *** Reference to "__vbaFreeVarList"
'0071e9a7    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_43, var_58, var_50)
'0071e9ad    83c410                  add esp, 10
'0071e9b0    c3                      ret
'0071e9b1    8d4de8                  lea ecx, dword ptr [ebp-18]

' *** Reference to "__vbaFreeStr"
'0071e9b4    ff150c134000            call dword ptr [0040130c]
'#Cleanup(var_41)
'0071e9ba    c3                      ret
'0071e9bb    8b4508                  mov eax, dword ptr [ebp+08]
'0071e9be    8b10                    mov edx, dword ptr [eax]
'0071e9c0    50                      push eax
'0071e9c1    ff5208                  call dword ptr [edx+08]
'0071e9c4    8b4510                  mov eax, dword ptr [ebp+10]
'0071e9c7    668b4de4                mov cx, word ptr [ebp-1c]
'0071e9cb    668908                  mov word ptr [eax], cx
'0071e9ce    8b45fc                  mov eax, dword ptr [ebp-04]
'0071e9d1    8b4dec                  mov ecx, dword ptr [ebp-14]
'0071e9d4    5f                      pop edi
'0071e9d5    5e                      pop esi
    'Reference to '__except_list'
'0071e9d6    64890d00000000          mov dword ptr fs:[00000000], ecx
'0071e9dd    5b                      pop ebx
'0071e9de    8be5                    mov esp, ebp
'0071e9e0    5d                      pop ebp
'0071e9e1    c20c00                  ret 000c


End Function


Public Function DisplayCurrentTip(arg_0 As Unknow, arg_1 As Unknow, arg_2 As Unknow, arg_3 As Unknow, arg_4 As Unknow, arg_5 As Unknow, arg_6 As Unknow, arg_7 As Unknow, arg_8 As Unknow, arg_9 As Unknow, arg_A As Unknow, arg_B As Unknow, arg_C As Unknow, arg_D As Unknow, arg_E As Unknow, arg_F As Unknow, arg_10 As Unknow, arg_11 As Unknow, arg_12 As Unknow, arg_13 As Unknow, arg_14 As Unknow, arg_15 As Unknow, arg_16 As Unknow, arg_17 As Unknow, arg_18 As Unknow, arg_19 As Unknow, arg_1A As Unknow, arg_1B As Unknow, arg_1C As Unknow, arg_1D As Unknow, arg_1E As Unknow, arg_1F As Unknow, arg_20 As Unknow, arg_21 As Unknow, arg_22 As Unknow, arg_23 As Unknow, arg_24 As Unknow, arg_25 As Unknow, arg_26 As Unknow, arg_27 As Unknow, arg_28 As Unknow, arg_29 As Unknow, arg_2A As Unknow, arg_2B As Unknow, arg_2C As Unknow, arg_2D As Unknow, arg_2E As Unknow, arg_2F As Unknow, arg_30 As Unknow, arg_31 As Unknow, arg_32 As Unknow, arg_33 As Unknow, arg_34 As Unknow, arg_35 As Unknow, arg_36 As Unknow, arg_37 As Unknow, arg_38 As Unknow, arg_39 As Unknow, arg_3A As Unknow, arg_3B As Unknow, arg_3C As Unknow)
'0071efe0    55                      push ebp
'0071efe1    8bec                    mov ebp, esp
'0071efe3    83ec0c                  sub esp, 0c

' *** Reference to "__vbaExceptHandler"
'0071efe6    6866724000              push 00407266
'0071efeb    64a100000000            mov ax, word ptr fs:[00000000]
'0071eff1    50                      push eax
    'Reference to '__except_list'
'0071eff2    64892500000000          mov dword ptr fs:[00000000], esp
'0071eff9    83ec44                  sub esp, 44
'0071effc    53                      push ebx
'0071effd    56                      push esi
'0071effe    57                      push edi
'0071efff    8965f4                  mov dword ptr [ebp-0c], esp
'0071f002    c745f8986f4000          mov dword ptr [ebp-08], 00406f98
'0071f009    33ff                    xor edi, edi
'0071f00b    897dfc                  mov dword ptr [ebp-04], edi
'0071f00e    8b5d08                  mov ebx, dword ptr [ebp+08]
'0071f011    8b03                    mov eax, dword ptr [ebx]
'0071f013    53                      push ebx
'0071f014    ff5004                  call dword ptr [eax+04]
'0071f017    8b4334                  mov eax, dword ptr [ebx+34]
'0071f01a    3bc7                    cmp eax, edi
'0071f01c    8d7334                  lea esi, dword ptr [ebx+34]
'0071f01f    897de8                  mov dword ptr [ebp-18], edi
'0071f022    897de4                  mov dword ptr [ebp-1c], edi
'0071f025    897dd4                  mov dword ptr [ebp-2c], edi
'0071f028    897dc4                  mov dword ptr [ebp-3c], edi
'0071f02b    897dc0                  mov dword ptr [ebp-40], edi
'0071f02e    750c                    jne 71f03c

If (arg_6 = 0) Then
'0071f030    56                      push esi
'0071f031    688c6e4500              push 00456e8c

' *** Reference to "__vbaNew2"
'0071f036    ff152c124000            call dword ptr [0040122c]
    Dim var_138 As New Collection
End If
'0071f03c    8b3e                    mov edi, dword ptr [esi]
'0071f03e    8b0f                    mov ecx, dword ptr [edi]
'0071f040    8d55c0                  lea edx, dword ptr [ebp-40]
'0071f043    52                      push edx
'0071f044    57                      push edi
'0071f045    ff5124                  call dword ptr [ecx+24]
var_5 = var_138.Count()
'0071f048    dbe2                    fnclex
'0071f04a    85c0                    test ax, ax
'0071f04c    7d0f                    jge 71f05d
'0071f04e    6a24                    push 24
'0071f050    687c6e4500              push 00456e7c
'0071f055    57                      push edi
'0071f056    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0071f057    ff1580104000            call dword ptr [00401080]
'0071f05d    8b45c0                  mov eax, dword ptr [ebp-40]
'0071f060    85c0                    test ax, ax
'0071f062    0f8e9e000000            jle 71f106
'0071f068    8b03                    mov eax, dword ptr [ebx]
'0071f06a    53                      push ebx
'0071f06b    ff900c030000            call dword ptr [eax+0000030c]
'0071f071    50                      push eax
'0071f072    8d4de4                  lea ecx, dword ptr [ebp-1c]
'0071f075    51                      push ecx

' *** Reference to "__vbaObjSet"
'0071f076    ff15b4104000            call dword ptr [004010b4]
Set var_40 = Nothing
'0071f07c    8bf8                    mov edi, eax
'0071f07e    833e00                  cmp dword ptr [esi], 00
'0071f081    750c                    jne 71f08f

If (var_138 = 0) Then
'0071f083    56                      push esi
'0071f084    688c6e4500              push 00456e8c

' *** Reference to "__vbaNew2"
'0071f089    ff152c124000            call dword ptr [0040122c]
    Set var_138 = New Collection
End If
'0071f08f    8b36                    mov esi, dword ptr [esi]
'0071f091    8d45d4                  lea eax, dword ptr [ebp-2c]
'0071f094    50                      push eax
'0071f095    8d4dc4                  lea ecx, dword ptr [ebp-3c]
'0071f098    83c338                  add ebx, 38
var_num2 = Me + 56
'0071f09b    51                      push ecx
'0071f09c    895dcc                  mov dword ptr [ebp-34], ebx
'0071f09f    c745c403400000          mov dword ptr [ebp-3c], 00004003
'0071f0a6    8b16                    mov edx, dword ptr [esi]
'0071f0a8    56                      push esi
'0071f0a9    ff521c                  call dword ptr [edx+1c]
[VAR_Unknown] = var_138.Item(arg_6)
'0071f0ac    dbe2                    fnclex
'0071f0ae    85c0                    test ax, ax
'0071f0b0    7d0f                    jge 71f0c1
'0071f0b2    6a1c                    push 1c
'0071f0b4    687c6e4500              push 00456e7c
'0071f0b9    56                      push esi
'0071f0ba    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0071f0bb    ff1580104000            call dword ptr [00401080]
'0071f0c1    8b37                    mov esi, dword ptr [edi]
'0071f0c3    8d55d4                  lea edx, dword ptr [ebp-2c]
'0071f0c6    52                      push edx
'0071f0c7    8d45e8                  lea eax, dword ptr [ebp-18]
'0071f0ca    50                      push eax

' *** Reference to "__vbaStrVarVal"
'0071f0cb    ff15fc114000            call dword ptr [004011fc]
'0071f0d1    50                      push eax
'0071f0d2    57                      push edi
'0071f0d3    ff5654                  call dword ptr [esi+54]
'0071f0d6    dbe2                    fnclex
'0071f0d8    85c0                    test ax, ax
'0071f0da    7d0f                    jge 71f0eb
'0071f0dc    6a54                    push 54
'0071f0de    683ce04300              push 0043e03c
'0071f0e3    57                      push edi
'0071f0e4    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0071f0e5    ff1580104000            call dword ptr [00401080]
'0071f0eb    8d4de8                  lea ecx, dword ptr [ebp-18]

' *** Reference to "__vbaFreeStr"
'0071f0ee    ff150c134000            call dword ptr [0040130c]
'#Cleanup(var_41)
'0071f0f4    8d4de4                  lea ecx, dword ptr [ebp-1c]

' *** Reference to "__vbaFreeObj"
'0071f0f7    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_40)
'0071f0fd    8d4dd4                  lea ecx, dword ptr [ebp-2c]

' *** Reference to "__vbaFreeVar"
'0071f100    ff1528104000            call dword ptr [00401028]
'#Cleanup(var_44)
'0071f106    682af17100              push 0071f12a
'0071f10b    eb1c                    jmp 71f129
'0071f10d    8d4de8                  lea ecx, dword ptr [ebp-18]

' *** Reference to "__vbaFreeStr"
'0071f110    ff150c134000            call dword ptr [0040130c]
'#Cleanup(var_41)
'0071f116    8d4de4                  lea ecx, dword ptr [ebp-1c]

' *** Reference to "__vbaFreeObj"
'0071f119    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_40)
'0071f11f    8d4dd4                  lea ecx, dword ptr [ebp-2c]

' *** Reference to "__vbaFreeVar"
'0071f122    ff1528104000            call dword ptr [00401028]
'#Cleanup(var_44)
'0071f128    c3                      ret
'0071f129    c3                      ret
'0071f12a    8b4508                  mov eax, dword ptr [ebp+08]
'0071f12d    8b08                    mov ecx, dword ptr [eax]
'0071f12f    50                      push eax
'0071f130    ff5108                  call dword ptr [ecx+08]
'0071f133    8b45fc                  mov eax, dword ptr [ebp-04]
'0071f136    8b4dec                  mov ecx, dword ptr [ebp-14]
'0071f139    5f                      pop edi
'0071f13a    5e                      pop esi
    'Reference to '__except_list'
'0071f13b    64890d00000000          mov dword ptr fs:[00000000], ecx
'0071f142    5b                      pop ebx
'0071f143    8be5                    mov esp, ebp
'0071f145    5d                      pop ebp
'0071f146    c20400                  ret 0004


End Function


'Event for chkLoadTipsAtStartup
Private Sub chkLoadTipsAtStartup_Click()
'0071e9f0    55                      push ebp
'0071e9f1    8bec                    mov ebp, esp
'0071e9f3    83ec0c                  sub esp, 0c

' *** Reference to "__vbaExceptHandler"
'0071e9f6    6866724000              push 00407266
'0071e9fb    64a100000000            mov ax, word ptr fs:[00000000]
'0071ea01    50                      push eax
    'Reference to '__except_list'
'0071ea02    64892500000000          mov dword ptr fs:[00000000], esp
'0071ea09    83ec34                  sub esp, 34
'0071ea0c    53                      push ebx
'0071ea0d    56                      push esi
'0071ea0e    57                      push edi
'0071ea0f    8965f4                  mov dword ptr [ebp-0c], esp
'0071ea12    c745f8606f4000          mov dword ptr [ebp-08], 00406f60
'0071ea19    8b7d08                  mov edi, dword ptr [ebp+08]
'0071ea1c    8bc7                    mov eax, edi
'0071ea1e    83e001                  and eax, 01
'0071ea21    8945fc                  mov dword ptr [ebp-04], eax
'0071ea24    83e7fe                  and edi, fffffffe
'0071ea27    8b0f                    mov ecx, dword ptr [edi]
'0071ea29    57                      push edi
'0071ea2a    897d08                  mov dword ptr [ebp+08], edi
'0071ea2d    ff5104                  call dword ptr [ecx+04]
'0071ea30    a124be7200              mov ax, word ptr [0072be24]
'0071ea35    33db                    xor ebx, ebx
'0071ea37    3bc3                    cmp eax, ebx
'0071ea39    895de8                  mov dword ptr [ebp-18], ebx
'0071ea3c    895de4                  mov dword ptr [ebp-1c], ebx
'0071ea3f    895de0                  mov dword ptr [ebp-20], ebx
'0071ea42    895ddc                  mov dword ptr [ebp-24], ebx
'0071ea45    895dd8                  mov dword ptr [ebp-28], ebx
'0071ea48    7510                    jne 71ea5a

If (0 = 0) Then
'0071ea4a    6824be7200              push 0072be24
'0071ea4f    6870004300              push 00430070

' *** Reference to "__vbaNew2"
'0071ea54    ff152c124000            call dword ptr [0040122c]
    Dim var_2 As New Global
End If
'0071ea5a    8b3524be7200            mov esi, dword ptr [0072be24]
'0071ea60    8b16                    mov edx, dword ptr [esi]
'0071ea62    8d45e0                  lea eax, dword ptr [ebp-20]
'0071ea65    50                      push eax
'0071ea66    56                      push esi
'0071ea67    ff5214                  call dword ptr [edx+14]
Set var_3 = var_2.App
'0071ea6a    dbe2                    fnclex
'0071ea6c    3bc3                    cmp eax, ebx
'0071ea6e    7d13                    jge 71ea83

If (var_2.App < 0) Then

' *** Reference to "__vbaHresultCheckObj"
'0071ea70    8b1d80104000            mov ebx, dword ptr [00401080]
'0071ea76    6a14                    push 14
'0071ea78    6860004300              push 00430060
'0071ea7d    56                      push esi
'0071ea7e    50                      push eax
'0071ea7f    ffd3                    call ebx
'0071ea81    eb06                    jmp 71ea89
    
End If

' *** Reference to "__vbaHresultCheckObj"
'0071ea83    8b1d80104000            mov ebx, dword ptr [00401080]
'0071ea89    8b45e0                  mov eax, dword ptr [ebp-20]
'0071ea8c    8b08                    mov ecx, dword ptr [eax]
'0071ea8e    8d55e8                  lea edx, dword ptr [ebp-18]
'0071ea91    52                      push edx
'0071ea92    50                      push eax
'0071ea93    8bf0                    mov esi, eax
'0071ea95    ff5158                  call dword ptr [ecx+58]
var_41 = var_3.EXEName
'0071ea98    dbe2                    fnclex
'0071ea9a    85c0                    test ax, ax
'0071ea9c    7d0b                    jge 71eaa9
'0071ea9e    6a58                    push 58
'0071eaa0    682c1c4300              push 00431c2c
'0071eaa5    56                      push esi
'0071eaa6    50                      push eax
'0071eaa7    ffd3                    call ebx
'0071eaa9    8b07                    mov eax, dword ptr [edi]
'0071eaab    57                      push edi
'0071eaac    ff90fc020000            call dword ptr [eax+000002fc]
'0071eab2    50                      push eax
'0071eab3    8d4ddc                  lea ecx, dword ptr [ebp-24]
'0071eab6    51                      push ecx

' *** Reference to "__vbaObjSet"
'0071eab7    ff15b4104000            call dword ptr [004010b4]
Set var_39 = Nothing
'0071eabd    8bf0                    mov esi, eax
'0071eabf    8b16                    mov edx, dword ptr [esi]
'0071eac1    8d45d8                  lea eax, dword ptr [ebp-28]
'0071eac4    50                      push eax
'0071eac5    56                      push esi
'0071eac6    ff92e0000000            call dword ptr [edx+000000e0]
'0071eacc    dbe2                    fnclex
'0071eace    85c0                    test ax, ax
'0071ead0    7d0e                    jge 71eae0
'0071ead2    68e0000000              push 000000e0
'0071ead7    68dce24300              push 0043e2dc
'0071eadc    56                      push esi
'0071eadd    50                      push eax
'0071eade    ffd3                    call ebx
'0071eae0    8b4dd8                  mov ecx, dword ptr [ebp-28]
'0071eae3    51                      push ecx

' *** Reference to "__vbaStrI2"
'0071eae4    ff1510104000            call dword ptr [00401010]
'0071eaea    8bd0                    mov edx, eax
'0071eaec    8d4de4                  lea ecx, dword ptr [ebp-1c]

' *** Reference to "__vbaStrMove"
'0071eaef    ff15d0124000            call dword ptr [004012d0]
'0071eaf5    8b55e8                  mov edx, dword ptr [ebp-18]
'0071eaf8    50                      push eax
'0071eaf9    68a4d84300              push 0043d8a4
'0071eafe    6890d84300              push 0043d890
'0071eb03    52                      push edx

' *** Reference to "rtcSaveSetting"
'0071eb04    ff150c104000            call dword ptr [0040100c]
SaveSetting var_41, "Options", "Show Tips at Startup", CStr(0)
'0071eb0a    8d45e4                  lea eax, dword ptr [ebp-1c]
'0071eb0d    50                      push eax
'0071eb0e    8d4de8                  lea ecx, dword ptr [ebp-18]
'0071eb11    51                      push ecx
'0071eb12    6a02                    push 02

' *** Reference to "__vbaFreeStrList"
'0071eb14    ff155c124000            call dword ptr [0040125c]
'#Cleanup( 0, -4508)
'0071eb1a    8d55dc                  lea edx, dword ptr [ebp-24]
'0071eb1d    52                      push edx
'0071eb1e    8d45e0                  lea eax, dword ptr [ebp-20]
'0071eb21    50                      push eax
'0071eb22    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'0071eb24    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_3, var_39)
'0071eb2a    83c418                  add esp, 18
'0071eb2d    c745fc00000000          mov dword ptr [ebp-04], 00000000
'0071eb34    6860eb7100              push 0071eb60
'0071eb39    eb24                    jmp 71eb5f
'0071eb3b    8d4de4                  lea ecx, dword ptr [ebp-1c]
'0071eb3e    51                      push ecx
'0071eb3f    8d55e8                  lea edx, dword ptr [ebp-18]
'0071eb42    52                      push edx
'0071eb43    6a02                    push 02

' *** Reference to "__vbaFreeStrList"
'0071eb45    ff155c124000            call dword ptr [0040125c]
'#Cleanup( 0, -4508)
'0071eb4b    8d45dc                  lea eax, dword ptr [ebp-24]
'0071eb4e    50                      push eax
'0071eb4f    8d4de0                  lea ecx, dword ptr [ebp-20]
'0071eb52    51                      push ecx
'0071eb53    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'0071eb55    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_3, var_39)
'0071eb5b    83c418                  add esp, 18
'0071eb5e    c3                      ret
'0071eb5f    c3                      ret
'0071eb60    8b4508                  mov eax, dword ptr [ebp+08]
'0071eb63    8b10                    mov edx, dword ptr [eax]
'0071eb65    50                      push eax
'0071eb66    ff5208                  call dword ptr [edx+08]
'0071eb69    8b45fc                  mov eax, dword ptr [ebp-04]
'0071eb6c    8b4dec                  mov ecx, dword ptr [ebp-14]
'0071eb6f    5f                      pop edi
'0071eb70    5e                      pop esi
    'Reference to '__except_list'
'0071eb71    64890d00000000          mov dword ptr fs:[00000000], ecx
'0071eb78    5b                      pop ebx
'0071eb79    8be5                    mov esp, ebp
'0071eb7b    5d                      pop ebp
'0071eb7c    c20400                  ret 0004


End Sub


'Event for Form
Private Sub Form_Load()
'0071ecc0    55                      push ebp
'0071ecc1    8bec                    mov ebp, esp
'0071ecc3    83ec0c                  sub esp, 0c

' *** Reference to "__vbaExceptHandler"
'0071ecc6    6866724000              push 00407266
'0071eccb    64a100000000            mov ax, word ptr fs:[00000000]
'0071ecd1    50                      push eax
    'Reference to '__except_list'
'0071ecd2    64892500000000          mov dword ptr fs:[00000000], esp
'0071ecd9    83ec6c                  sub esp, 6c
'0071ecdc    53                      push ebx
'0071ecdd    56                      push esi
'0071ecde    57                      push edi
'0071ecdf    8965f4                  mov dword ptr [ebp-0c], esp
'0071ece2    c745f8886f4000          mov dword ptr [ebp-08], 00406f88
'0071ece9    8b5d08                  mov ebx, dword ptr [ebp+08]
'0071ecec    8bc3                    mov eax, ebx
'0071ecee    83e001                  and eax, 01
'0071ecf1    8945fc                  mov dword ptr [ebp-04], eax
'0071ecf4    83e3fe                  and ebx, fffffffe
'0071ecf7    8b0b                    mov ecx, dword ptr [ebx]
'0071ecf9    53                      push ebx
'0071ecfa    895d08                  mov dword ptr [ebp+08], ebx
'0071ecfd    ff5104                  call dword ptr [ecx+04]
'0071ed00    8b13                    mov edx, dword ptr [ebx]
'0071ed02    33f6                    xor esi, esi
'0071ed04    53                      push ebx
'0071ed05    8975e8                  mov dword ptr [ebp-18], esi
'0071ed08    8975e4                  mov dword ptr [ebp-1c], esi
'0071ed0b    8975e0                  mov dword ptr [ebp-20], esi
'0071ed0e    8975dc                  mov dword ptr [ebp-24], esi
'0071ed11    8975d8                  mov dword ptr [ebp-28], esi
'0071ed14    8975d4                  mov dword ptr [ebp-2c], esi
'0071ed17    8975d0                  mov dword ptr [ebp-30], esi
'0071ed1a    8975cc                  mov dword ptr [ebp-34], esi
'0071ed1d    8975c8                  mov dword ptr [ebp-38], esi
'0071ed20    8975b8                  mov dword ptr [ebp-48], esi
'0071ed23    8975a4                  mov dword ptr [ebp-5c], esi
'0071ed26    ff92fc020000            call dword ptr [edx+000002fc]
'0071ed2c    50                      push eax
'0071ed2d    8d45c8                  lea eax, dword ptr [ebp-38]
'0071ed30    50                      push eax

' *** Reference to "__vbaObjSet"
'0071ed31    ff15b4104000            call dword ptr [004010b4]
Set var_46 = Me
'0071ed37    8b38                    mov edi, dword ptr [eax]
'0071ed39    b901000000              mov ecx, 00000001
'0071ed3e    8945a0                  mov dword ptr [ebp-60], eax

' *** Reference to "__vbaI2I4"
'0071ed41    ff1550114000            call dword ptr [00401150]
'0071ed47    8bcf                    mov ecx, edi
'0071ed49    8b7da0                  mov edi, dword ptr [ebp-60]
'0071ed4c    50                      push eax
'0071ed4d    57                      push edi
'0071ed4e    ff91e4000000            call dword ptr [ecx+000000e4]
'0071ed54    dbe2                    fnclex
'0071ed56    3bc6                    cmp eax, esi
'0071ed58    7d12                    jge 71ed6c

If (CInt(1) < 0) Then
'0071ed5a    68e4000000              push 000000e4
'0071ed5f    68dce24300              push 0043e2dc
'0071ed64    57                      push edi
'0071ed65    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0071ed66    ff1580104000            call dword ptr [00401080]
    
End If
'0071ed6c    8d4dc8                  lea ecx, dword ptr [ebp-38]

' *** Reference to "__vbaFreeObj"
'0071ed6f    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_46)
'0071ed75    8d55b8                  lea edx, dword ptr [ebp-48]
'0071ed78    52                      push edx
'0071ed79    c745c004000280          mov dword ptr [ebp-40], 80020004
'0071ed80    c745b80a000000          mov dword ptr [ebp-48], 0000000a

' *** Reference to "rtcRandomize"
'0071ed87    ff15ac104000            call dword ptr [004010ac]
Randomize 
'0071ed8d    8d4db8                  lea ecx, dword ptr [ebp-48]

' *** Reference to "__vbaFreeVar"
'0071ed90    ff1528104000            call dword ptr [00401028]
'#Cleanup(var_61)
'0071ed96    393524be7200            cmp dword ptr [0072be24], esi
'0071ed9c    7510                    jne 71edae
'0071ed9e    6824be7200              push 0072be24
'0071eda3    6870004300              push 00430070

' *** Reference to "__vbaNew2"
'0071eda8    ff152c124000            call dword ptr [0040122c]
Dim var_2 As New Global
'0071edae    8b3d24be7200            mov edi, dword ptr [0072be24]
'0071edb4    8b07                    mov eax, dword ptr [edi]
'0071edb6    8d4dc8                  lea ecx, dword ptr [ebp-38]
'0071edb9    51                      push ecx
'0071edba    57                      push edi
'0071edbb    ff5014                  call dword ptr [eax+14]
Set var_46 = var_2.App
'0071edbe    dbe2                    fnclex
'0071edc0    3bc6                    cmp eax, esi
'0071edc2    7d0f                    jge 71edd3

If (var_2.App < 0) Then
'0071edc4    6a14                    push 14
'0071edc6    6860004300              push 00430060
'0071edcb    57                      push edi
'0071edcc    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0071edcd    ff1580104000            call dword ptr [00401080]
    
End If
'0071edd3    8b45c8                  mov eax, dword ptr [ebp-38]
'0071edd6    8b10                    mov edx, dword ptr [eax]
'0071edd8    8d4de8                  lea ecx, dword ptr [ebp-18]
'0071eddb    51                      push ecx
'0071eddc    50                      push eax
'0071eddd    8bf8                    mov edi, eax
'0071eddf    ff5250                  call dword ptr [edx+50]
var_41 = var_46.Path
'0071ede2    dbe2                    fnclex
'0071ede4    3bc6                    cmp eax, esi
'0071ede6    7d0f                    jge 71edf7
'0071ede8    6a50                    push 50
'0071edea    682c1c4300              push 00431c2c
'0071edef    57                      push edi
'0071edf0    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0071edf1    ff1580104000            call dword ptr [00401080]
'0071edf7    8b55e8                  mov edx, dword ptr [ebp-18]

' *** Reference to "__vbaStrCat"
'0071edfa    8b3570104000            mov esi, dword ptr [00401070]
'0071ee00    52                      push edx
'0071ee01    68b0d64300              push 0043d6b0
'0071ee06    ffd6                    call esi
var_49 = (var_41) & ("\")

' *** Reference to "__vbaStrMove"
'0071ee08    8b3dd0124000            mov edi, dword ptr [004012d0]
'0071ee0e    8bd0                    mov edx, eax
'0071ee10    8d4de4                  lea ecx, dword ptr [ebp-1c]
'0071ee13    ffd7                    call edi
'0071ee15    50                      push eax
'0071ee16    6834fd4200              push 0042fd34
'0071ee1b    ffd6                    call esi
var_12 = (var_49) & ("ASTUCE.TXT")
'0071ee1d    8bd0                    mov edx, eax
'0071ee1f    8d4de0                  lea ecx, dword ptr [ebp-20]
'0071ee22    ffd7                    call edi
'0071ee24    8b03                    mov eax, dword ptr [ebx]
'0071ee26    8d4da4                  lea ecx, dword ptr [ebp-5c]
'0071ee29    51                      push ecx
'0071ee2a    8d55e0                  lea edx, dword ptr [ebp-20]
'0071ee2d    52                      push edx
'0071ee2e    53                      push ebx
'0071ee2f    ff90f8060000            call dword ptr [eax+000006f8]
'0071ee35    85c0                    test ax, ax
'0071ee37    7d12                    jge 71ee4b
'0071ee39    68f8060000              push 000006f8
'0071ee3e    6854104300              push 00431054
'0071ee43    53                      push ebx
'0071ee44    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0071ee45    ff1580104000            call dword ptr [00401080]
'0071ee4b    33c0                    xor eax, eax
var_num1 = Empty
'0071ee4d    663945a4                cmp word ptr [ebp-5c], ax
'0071ee51    8d4de0                  lea ecx, dword ptr [ebp-20]
'0071ee54    0f94c0                  sete al
'0071ee57    51                      push ecx
'0071ee58    8d55e4                  lea edx, dword ptr [ebp-1c]
'0071ee5b    52                      push edx
'0071ee5c    f7d8                    neg eax
'0071ee5e    89458c                  mov dword ptr [ebp-74], eax
'0071ee61    8d45e8                  lea eax, dword ptr [ebp-18]
'0071ee64    50                      push eax
'0071ee65    6a03                    push 03

' *** Reference to "__vbaFreeStrList"
'0071ee67    ff155c124000            call dword ptr [0040125c]
'#Cleanup( 0, -4512, -4516)
'0071ee6d    83c410                  add esp, 10
'0071ee70    8d4dc8                  lea ecx, dword ptr [ebp-38]

' *** Reference to "__vbaFreeObj"
'0071ee73    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_46)
'0071ee79    66837d8c00              cmp word ptr [ebp-74], 00
'0071ee7e    0f84e6000000            je 71ef6a

If (0 = frmAstuce) Then
'0071ee84    8b0b                    mov ecx, dword ptr [ebx]
'0071ee86    53                      push ebx
'0071ee87    ff910c030000            call dword ptr [ecx+0000030c]
'0071ee8d    50                      push eax
'0071ee8e    8d55c8                  lea edx, dword ptr [ebp-38]
'0071ee91    52                      push edx

' *** Reference to "__vbaObjSet"
'0071ee92    ff15b4104000            call dword ptr [004010b4]
    Set var_46 = 
'0071ee98    8b18                    mov ebx, dword ptr [eax]
'0071ee9a    68a06e4500              push 00456ea0
'0071ee9f    6834fd4200              push 0042fd34
'0071eea4    8945a0                  mov dword ptr [ebp-60], eax
'0071eea7    ffd6                    call esi
    var_41 = ("Le fichier ") & ("ASTUCE.TXT")
'0071eea9    8bd0                    mov edx, eax
'0071eeab    8d4de8                  lea ecx, dword ptr [ebp-18]
'0071eeae    ffd7                    call edi
'0071eeb0    50                      push eax
'0071eeb1    68bc6e4500              push 00456ebc
'0071eeb6    ffd6                    call esi
    var_14 = (var_41) & (" n'a pas été trouvé? ")
'0071eeb8    8bd0                    mov edx, eax
'0071eeba    8d4de4                  lea ecx, dword ptr [ebp-1c]
'0071eebd    ffd7                    call edi
'0071eebf    50                      push eax
'0071eec0    6870084300              push 00430870
'0071eec5    ffd6                    call esi
    var_127 = (var_14) & (vbCrLf)
'0071eec7    8bd0                    mov edx, eax
'0071eec9    8d4de0                  lea ecx, dword ptr [ebp-20]
'0071eecc    ffd7                    call edi
'0071eece    50                      push eax
'0071eecf    6870084300              push 00430870
'0071eed4    ffd6                    call esi
    var_15 = (var_127) & (vbCrLf)
'0071eed6    8bd0                    mov edx, eax
'0071eed8    8d4ddc                  lea ecx, dword ptr [ebp-24]
'0071eedb    ffd7                    call edi
'0071eedd    50                      push eax
'0071eede    68ec6e4500              push 00456eec
'0071eee3    ffd6                    call esi
    var_16 = (var_15) & ("Créez un fichier texte nommé ")
'0071eee5    8bd0                    mov edx, eax
'0071eee7    8d4dd8                  lea ecx, dword ptr [ebp-28]
'0071eeea    ffd7                    call edi
'0071eeec    50                      push eax
'0071eeed    6834fd4200              push 0042fd34
'0071eef2    ffd6                    call esi
    var_128 = (var_16) & ("ASTUCE.TXT")
'0071eef4    8bd0                    mov edx, eax
'0071eef6    8d4dd4                  lea ecx, dword ptr [ebp-2c]
'0071eef9    ffd7                    call edi
'0071eefb    50                      push eax
'0071eefc    682c6f4500              push 00456f2c
'0071ef01    ffd6                    call esi
    var_17 = (var_128) & (" en utilisant le Bloc-notes avec une astuce par ligne. ")
'0071ef03    8bd0                    mov edx, eax
'0071ef05    8d4dd0                  lea ecx, dword ptr [ebp-30]
'0071ef08    ffd7                    call edi
'0071ef0a    50                      push eax
'0071ef0b    68c86f4500              push 00456fc8
'0071ef10    ffd6                    call esi
    var_129 = (var_17) & ("Puis placez le dans le même dossier que l'application.")
'0071ef12    8bd0                    mov edx, eax
'0071ef14    8d4dcc                  lea ecx, dword ptr [ebp-34]
'0071ef17    ffd7                    call edi
'0071ef19    8b75a0                  mov esi, dword ptr [ebp-60]
'0071ef1c    50                      push eax
'0071ef1d    56                      push esi
'0071ef1e    ff5354                  call dword ptr [ebx+54]
'0071ef21    dbe2                    fnclex
'0071ef23    85c0                    test ax, ax
'0071ef25    7d0f                    jge 71ef36
'0071ef27    6a54                    push 54
'0071ef29    683ce04300              push 0043e03c
'0071ef2e    56                      push esi
'0071ef2f    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0071ef30    ff1580104000            call dword ptr [00401080]
'0071ef36    8d45cc                  lea eax, dword ptr [ebp-34]
'0071ef39    50                      push eax
'0071ef3a    8d4dd0                  lea ecx, dword ptr [ebp-30]
'0071ef3d    51                      push ecx
'0071ef3e    8d55d4                  lea edx, dword ptr [ebp-2c]
'0071ef41    52                      push edx
'0071ef42    8d45d8                  lea eax, dword ptr [ebp-28]
'0071ef45    50                      push eax
'0071ef46    8d4ddc                  lea ecx, dword ptr [ebp-24]
'0071ef49    51                      push ecx
'0071ef4a    8d55e0                  lea edx, dword ptr [ebp-20]
'0071ef4d    52                      push edx
'0071ef4e    8d45e4                  lea eax, dword ptr [ebp-1c]
'0071ef51    50                      push eax
'0071ef52    8d4de8                  lea ecx, dword ptr [ebp-18]
'0071ef55    51                      push ecx
'0071ef56    6a08                    push 08

' *** Reference to "__vbaFreeStrList"
'0071ef58    ff155c124000            call dword ptr [0040125c]
    '#Cleanup( -4520, -4524, -4528, -4532, -4536, -4540, -4544, -4548)
'0071ef5e    83c424                  add esp, 24
'0071ef61    8d4dc8                  lea ecx, dword ptr [ebp-38]

' *** Reference to "__vbaFreeObj"
'0071ef64    ff1508134000            call dword ptr [00401308]
    '#Cleanup(var_46)
End If
'0071ef6a    c745fc00000000          mov dword ptr [ebp-04], 00000000
'0071ef71    68b7ef7100              push 0071efb7
'0071ef76    eb3e                    jmp 71efb6
'0071ef78    8d55cc                  lea edx, dword ptr [ebp-34]
'0071ef7b    52                      push edx
'0071ef7c    8d45d0                  lea eax, dword ptr [ebp-30]
'0071ef7f    50                      push eax
'0071ef80    8d4dd4                  lea ecx, dword ptr [ebp-2c]
'0071ef83    51                      push ecx
'0071ef84    8d55d8                  lea edx, dword ptr [ebp-28]
'0071ef87    52                      push edx
'0071ef88    8d45dc                  lea eax, dword ptr [ebp-24]
'0071ef8b    50                      push eax
'0071ef8c    8d4de0                  lea ecx, dword ptr [ebp-20]
'0071ef8f    51                      push ecx
'0071ef90    8d55e4                  lea edx, dword ptr [ebp-1c]
'0071ef93    52                      push edx
'0071ef94    8d45e8                  lea eax, dword ptr [ebp-18]
'0071ef97    50                      push eax
'0071ef98    6a08                    push 08

' *** Reference to "__vbaFreeStrList"
'0071ef9a    ff155c124000            call dword ptr [0040125c]
'#Cleanup( -4520, -4524, -4528, -4532, -4536, -4540, -4544, -4548)
'0071efa0    83c424                  add esp, 24
'0071efa3    8d4dc8                  lea ecx, dword ptr [ebp-38]

' *** Reference to "__vbaFreeObj"
'0071efa6    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_46)
'0071efac    8d4db8                  lea ecx, dword ptr [ebp-48]

' *** Reference to "__vbaFreeVar"
'0071efaf    ff1528104000            call dword ptr [00401028]
'#Cleanup(var_61)
'0071efb5    c3                      ret
'0071efb6    c3                      ret
'0071efb7    8b4508                  mov eax, dword ptr [ebp+08]
'0071efba    8b08                    mov ecx, dword ptr [eax]
'0071efbc    50                      push eax
'0071efbd    ff5108                  call dword ptr [ecx+08]
'0071efc0    8b45fc                  mov eax, dword ptr [ebp-04]
'0071efc3    8b4dec                  mov ecx, dword ptr [ebp-14]
'0071efc6    5f                      pop edi
'0071efc7    5e                      pop esi
    'Reference to '__except_list'
'0071efc8    64890d00000000          mov dword ptr fs:[00000000], ecx
'0071efcf    5b                      pop ebx
'0071efd0    8be5                    mov esp, ebp
'0071efd2    5d                      pop ebp
'0071efd3    c20400                  ret 0004


End Sub


'Event for cmdOK
Private Sub cmdOK_Click()
'0071ebf0    55                      push ebp
'0071ebf1    8bec                    mov ebp, esp
'0071ebf3    83ec0c                  sub esp, 0c

' *** Reference to "__vbaExceptHandler"
'0071ebf6    6866724000              push 00407266
'0071ebfb    64a100000000            mov ax, word ptr fs:[00000000]
'0071ec01    50                      push eax
    'Reference to '__except_list'
'0071ec02    64892500000000          mov dword ptr fs:[00000000], esp
'0071ec09    83ec18                  sub esp, 18
'0071ec0c    53                      push ebx
'0071ec0d    56                      push esi
'0071ec0e    57                      push edi
'0071ec0f    8965f4                  mov dword ptr [ebp-0c], esp
'0071ec12    c745f8786f4000          mov dword ptr [ebp-08], 00406f78
'0071ec19    8b7d08                  mov edi, dword ptr [ebp+08]
'0071ec1c    8bc7                    mov eax, edi
'0071ec1e    83e001                  and eax, 01
'0071ec21    8945fc                  mov dword ptr [ebp-04], eax
'0071ec24    83e7fe                  and edi, fffffffe
'0071ec27    8b0f                    mov ecx, dword ptr [edi]
'0071ec29    57                      push edi
'0071ec2a    897d08                  mov dword ptr [ebp+08], edi
'0071ec2d    ff5104                  call dword ptr [ecx+04]
'0071ec30    a124be7200              mov ax, word ptr [0072be24]
'0071ec35    33db                    xor ebx, ebx
'0071ec37    3bc3                    cmp eax, ebx
'0071ec39    895de8                  mov dword ptr [ebp-18], ebx
'0071ec3c    7510                    jne 71ec4e

If (0 = 0) Then
'0071ec3e    6824be7200              push 0072be24
'0071ec43    6870004300              push 00430070

' *** Reference to "__vbaNew2"
'0071ec48    ff152c124000            call dword ptr [0040122c]
    Dim var_2 As New Global
End If
'0071ec4e    8b3524be7200            mov esi, dword ptr [0072be24]
'0071ec54    8b16                    mov edx, dword ptr [esi]
'0071ec56    57                      push edi
'0071ec57    8d45e8                  lea eax, dword ptr [ebp-18]
'0071ec5a    50                      push eax
'0071ec5b    8955d4                  mov dword ptr [ebp-2c], edx

' *** Reference to "__vbaObjSetAddref"
'0071ec5e    ff15c8104000            call dword ptr [004010c8]
Set var_41 = Me
'0071ec64    8b4dd4                  mov ecx, dword ptr [ebp-2c]
'0071ec67    50                      push eax
'0071ec68    56                      push esi
'0071ec69    ff5110                  call dword ptr [ecx+10]
Call var_2.Unload(var_41)
'0071ec6c    dbe2                    fnclex
'0071ec6e    3bc3                    cmp eax, ebx
'0071ec70    7d0f                    jge 71ec81
'0071ec72    6a10                    push 10
'0071ec74    6860004300              push 00430060
'0071ec79    56                      push esi
'0071ec7a    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0071ec7b    ff1580104000            call dword ptr [00401080]
'0071ec81    8d4de8                  lea ecx, dword ptr [ebp-18]

' *** Reference to "__vbaFreeObj"
'0071ec84    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_41)
'0071ec8a    895dfc                  mov dword ptr [ebp-04], ebx
'0071ec8d    689fec7100              push 0071ec9f
'0071ec92    eb0a                    jmp 71ec9e
'0071ec94    8d4de8                  lea ecx, dword ptr [ebp-18]

' *** Reference to "__vbaFreeObj"
'0071ec97    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_41)
'0071ec9d    c3                      ret
'0071ec9e    c3                      ret
'0071ec9f    8b4508                  mov eax, dword ptr [ebp+08]
'0071eca2    8b10                    mov edx, dword ptr [eax]
'0071eca4    50                      push eax
'0071eca5    ff5208                  call dword ptr [edx+08]
'0071eca8    8b45fc                  mov eax, dword ptr [ebp-04]
'0071ecab    8b4dec                  mov ecx, dword ptr [ebp-14]
'0071ecae    5f                      pop edi
'0071ecaf    5e                      pop esi
    'Reference to '__except_list'
'0071ecb0    64890d00000000          mov dword ptr fs:[00000000], ecx
'0071ecb7    5b                      pop ebx
'0071ecb8    8be5                    mov esp, ebp
'0071ecba    5d                      pop ebp
'0071ecbb    c20400                  ret 0004


End Sub


'Event for cmdNextTip
Private Sub cmdNextTip_Click()
'0071eb80    55                      push ebp
'0071eb81    8bec                    mov ebp, esp
'0071eb83    83ec0c                  sub esp, 0c

' *** Reference to "__vbaExceptHandler"
'0071eb86    6866724000              push 00407266
'0071eb8b    64a100000000            mov ax, word ptr fs:[00000000]
'0071eb91    50                      push eax
    'Reference to '__except_list'
'0071eb92    64892500000000          mov dword ptr fs:[00000000], esp
'0071eb99    83ec08                  sub esp, 08
'0071eb9c    53                      push ebx
'0071eb9d    56                      push esi
'0071eb9e    57                      push edi
'0071eb9f    8965f4                  mov dword ptr [ebp-0c], esp
'0071eba2    c745f8706f4000          mov dword ptr [ebp-08], 00406f70
'0071eba9    8b7508                  mov esi, dword ptr [ebp+08]
'0071ebac    8bc6                    mov eax, esi
'0071ebae    83e001                  and eax, 01
'0071ebb1    8945fc                  mov dword ptr [ebp-04], eax
'0071ebb4    83e6fe                  and esi, fffffffe
'0071ebb7    8b0e                    mov ecx, dword ptr [esi]
'0071ebb9    56                      push esi
'0071ebba    897508                  mov dword ptr [ebp+08], esi
'0071ebbd    ff5104                  call dword ptr [ecx+04]
'0071ebc0    8b16                    mov edx, dword ptr [esi]
'0071ebc2    56                      push esi
'0071ebc3    ff9200070000            call dword ptr [edx+00000700]
'0071ebc9    c745fc00000000          mov dword ptr [ebp-04], 00000000
'0071ebd0    8b4508                  mov eax, dword ptr [ebp+08]
'0071ebd3    8b08                    mov ecx, dword ptr [eax]
'0071ebd5    50                      push eax
'0071ebd6    ff5108                  call dword ptr [ecx+08]
'0071ebd9    8b45fc                  mov eax, dword ptr [ebp-04]
'0071ebdc    8b4dec                  mov ecx, dword ptr [ebp-14]
'0071ebdf    5f                      pop edi
'0071ebe0    5e                      pop esi
    'Reference to '__except_list'
'0071ebe1    64890d00000000          mov dword ptr fs:[00000000], ecx
'0071ebe8    5b                      pop ebx
'0071ebe9    8be5                    mov esp, ebp
'0071ebeb    5d                      pop ebp
'0071ebec    c20400                  ret 0004


End Sub


