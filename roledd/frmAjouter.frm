VERSION 5.00

Begin VB.Form frmAjouter
    Caption = "Choix de la méthode"
    ScaleMode = 1
    AutoRedraw = 0              'False
    FontTransparent = -1              'True
    BorderStyle = 4
    LinkTopic = "Form1"
    MaxButton = 0              'False
    MinButton = 0              'False
    ControlBox = 0              'False
    ClientLeft   = 45
    ClientTop    = 285
    ClientWidth  = 2535
    ClientHeight = 1140
    ShowInTaskbar = 0              'False
    StartupPosition = 1
    Begin VB.CommandButton btnAnnuler
        Caption = "Annuler"
        Left   = 1620
        Top    = 720
        Width  = 855
        Height = 330
        TabIndex = 3
        Cancel = -1              'True
    End
    Begin VB.CommandButton btnOk
        Caption = "Ok"
        Left   = 720
        Top    = 720
        Width  = 855
        Height = 330
        TabIndex = 2
        Default = -1              'True
    End
    Begin VB.OptionButton Option2
        Caption = "Assistant"
        Left   = 1320
        Top    = 240
        Width  = 1215
        Height = 255
        TabIndex = 1
    End
    Begin VB.OptionButton Option1
        Caption = "Manuelle"
        Left   = 120
        Top    = 240
        Width  = 1215
        Height = 255
        TabIndex = 0
    End
End
'Event for btnOk
Private Sub btnOk_Click()
'00708070    55                      push ebp
'00708071    8bec                    mov ebp, esp
'00708073    83ec0c                  sub esp, 0c

' *** Reference to "__vbaExceptHandler"
'00708076    6866724000              push 00407266
'0070807b    64a100000000            mov ax, word ptr fs:[00000000]
'00708081    50                      push eax
    'Reference to '__except_list'
'00708082    64892500000000          mov dword ptr fs:[00000000], esp
'00708089    83ec1c                  sub esp, 1c
'0070808c    53                      push ebx
'0070808d    56                      push esi
'0070808e    57                      push edi
'0070808f    8965f4                  mov dword ptr [ebp-0c], esp
'00708092    c745f8e86b4000          mov dword ptr [ebp-08], 00406be8
'00708099    8b7508                  mov esi, dword ptr [ebp+08]
'0070809c    8bc6                    mov eax, esi
'0070809e    83e001                  and eax, 01
'007080a1    8945fc                  mov dword ptr [ebp-04], eax
'007080a4    83e6fe                  and esi, fffffffe
'007080a7    8b0e                    mov ecx, dword ptr [esi]
'007080a9    56                      push esi
'007080aa    897508                  mov dword ptr [ebp+08], esi
'007080ad    ff5104                  call dword ptr [ecx+04]
'007080b0    8b16                    mov edx, dword ptr [esi]
'007080b2    33c0                    xor eax, eax
var_num1 = Empty
'007080b4    56                      push esi
'007080b5    8945e8                  mov dword ptr [ebp-18], eax
'007080b8    8945e4                  mov dword ptr [ebp-1c], eax
'007080bb    ff9208030000            call dword ptr [edx+00000308]

' *** Reference to "__vbaObjSet"
'007080c1    8b1db4104000            mov ebx, dword ptr [004010b4]
'007080c7    50                      push eax
'007080c8    8d45e8                  lea eax, dword ptr [ebp-18]
'007080cb    50                      push eax
'007080cc    ffd3                    call ebx
Set var_41 = Me
'007080ce    8d55e4                  lea edx, dword ptr [ebp-1c]
'007080d1    8bf8                    mov edi, eax
'007080d3    8b0f                    mov ecx, dword ptr [edi]
'007080d5    52                      push edx
'007080d6    57                      push edi
'007080d7    ff91e0000000            call dword ptr [ecx+000000e0]
'007080dd    dbe2                    fnclex
'007080df    85c0                    test ax, ax
'007080e1    7d12                    jge 7080f5
'007080e3    68e0000000              push 000000e0
'007080e8    68f43b4500              push 00453bf4
'007080ed    57                      push edi
'007080ee    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'007080ef    ff1580104000            call dword ptr [00401080]
'007080f5    33c0                    xor eax, eax
var_num1 = Empty
'007080f7    66837de4ff              cmp word ptr [ebp-1c], ffffffff
'007080fc    8d4de8                  lea ecx, dword ptr [ebp-18]
'007080ff    0f94c0                  sete al
'00708102    f7d8                    neg eax
'00708104    8bf8                    mov edi, eax

' *** Reference to "__vbaFreeObj"
'00708106    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_41)
'0070810c    6685ff                  test edi, edi
'0070810f    7408                    je 708119

If (Me = -1) Then
'00708111    66c746340100            mov word ptr [esi+34], 0001
'00708117    eb59                    jmp 708172
    
End If
'00708119    8b0e                    mov ecx, dword ptr [esi]
'0070811b    56                      push esi
'0070811c    ff9104030000            call dword ptr [ecx+00000304]
'00708122    50                      push eax
'00708123    8d55e8                  lea edx, dword ptr [ebp-18]
'00708126    52                      push edx
'00708127    ffd3                    call ebx
Set var_41 = Me = -1
'00708129    8d4de4                  lea ecx, dword ptr [ebp-1c]
'0070812c    8bf8                    mov edi, eax
'0070812e    8b07                    mov eax, dword ptr [edi]
'00708130    51                      push ecx
'00708131    57                      push edi
'00708132    ff90e0000000            call dword ptr [eax+000000e0]
'00708138    dbe2                    fnclex
'0070813a    85c0                    test ax, ax
'0070813c    7d12                    jge 708150
'0070813e    68e0000000              push 000000e0
'00708143    68f43b4500              push 00453bf4
'00708148    57                      push edi
'00708149    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0070814a    ff1580104000            call dword ptr [00401080]
'00708150    33d2                    xor edx, edx
'00708152    66837de4ff              cmp word ptr [ebp-1c], ffffffff
'00708157    8d4de8                  lea ecx, dword ptr [ebp-18]
'0070815a    0f94c2                  sete dl
'0070815d    f7da                    neg edx
'0070815f    8bfa                    mov edi, edx

' *** Reference to "__vbaFreeObj"
'00708161    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_41)
'00708167    6685ff                  test edi, edi
'0070816a    7406                    je 708172

If (Me = -1) Then
'0070816c    66c746340200            mov word ptr [esi+34], 0002
    
End If
'00708172    a124be7200              mov ax, word ptr [0072be24]
'00708177    85c0                    test ax, ax
'00708179    7510                    jne 70818b
'0070817b    6824be7200              push 0072be24
'00708180    6870004300              push 00430070

' *** Reference to "__vbaNew2"
'00708185    ff152c124000            call dword ptr [0040122c]
Dim var_2 As New Global
'0070818b    8b3d24be7200            mov edi, dword ptr [0072be24]
'00708191    8b1f                    mov ebx, dword ptr [edi]
'00708193    56                      push esi
'00708194    8d45e8                  lea eax, dword ptr [ebp-18]
'00708197    50                      push eax

' *** Reference to "__vbaObjSetAddref"
'00708198    ff15c8104000            call dword ptr [004010c8]
Set var_41 = Me
'0070819e    50                      push eax
'0070819f    57                      push edi
'007081a0    ff5310                  call dword ptr [ebx+10]
Call var_2.Unload(var_41)
'007081a3    dbe2                    fnclex
'007081a5    85c0                    test ax, ax
'007081a7    7d0f                    jge 7081b8
'007081a9    6a10                    push 10
'007081ab    6860004300              push 00430060
'007081b0    57                      push edi
'007081b1    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'007081b2    ff1580104000            call dword ptr [00401080]
'007081b8    8d4de8                  lea ecx, dword ptr [ebp-18]

' *** Reference to "__vbaFreeObj"
'007081bb    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_41)
'007081c1    c745fc00000000          mov dword ptr [ebp-04], 00000000
'007081c8    68da817000              push 007081da
'007081cd    eb0a                    jmp 7081d9
'007081cf    8d4de8                  lea ecx, dword ptr [ebp-18]

' *** Reference to "__vbaFreeObj"
'007081d2    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_41)
'007081d8    c3                      ret
'007081d9    c3                      ret
'007081da    8b4508                  mov eax, dword ptr [ebp+08]
'007081dd    8b08                    mov ecx, dword ptr [eax]
'007081df    50                      push eax
'007081e0    ff5108                  call dword ptr [ecx+08]
'007081e3    8b45fc                  mov eax, dword ptr [ebp-04]
'007081e6    8b4dec                  mov ecx, dword ptr [ebp-14]
'007081e9    5f                      pop edi
'007081ea    5e                      pop esi
    'Reference to '__except_list'
'007081eb    64890d00000000          mov dword ptr fs:[00000000], ecx
'007081f2    5b                      pop ebx
'007081f3    8be5                    mov esp, ebp
'007081f5    5d                      pop ebp
'007081f6    c20400                  ret 0004


End Sub


'Event for Form
Private Sub Form_Load()
'00708200    55                      push ebp
'00708201    8bec                    mov ebp, esp
'00708203    83ec0c                  sub esp, 0c

' *** Reference to "__vbaExceptHandler"
'00708206    6866724000              push 00407266
'0070820b    64a100000000            mov ax, word ptr fs:[00000000]
'00708211    50                      push eax
    'Reference to '__except_list'
'00708212    64892500000000          mov dword ptr fs:[00000000], esp
'00708219    83ec08                  sub esp, 08
'0070821c    53                      push ebx
'0070821d    56                      push esi
'0070821e    57                      push edi
'0070821f    8965f4                  mov dword ptr [ebp-0c], esp
'00708222    c745f8f86b4000          mov dword ptr [ebp-08], 00406bf8
'00708229    8b7508                  mov esi, dword ptr [ebp+08]
'0070822c    8bc6                    mov eax, esi
'0070822e    83e001                  and eax, 01
'00708231    8945fc                  mov dword ptr [ebp-04], eax
'00708234    83e6fe                  and esi, fffffffe
'00708237    8b0e                    mov ecx, dword ptr [esi]
'00708239    56                      push esi
'0070823a    897508                  mov dword ptr [ebp+08], esi
'0070823d    ff5104                  call dword ptr [ecx+04]
'00708240    33c0                    xor eax, eax
var_num1 = Empty
'00708242    66894634                mov word ptr [esi+34], ax
'00708246    8945fc                  mov dword ptr [ebp-04], eax
'00708249    8b4508                  mov eax, dword ptr [ebp+08]
'0070824c    8b10                    mov edx, dword ptr [eax]
'0070824e    50                      push eax
'0070824f    ff5208                  call dword ptr [edx+08]
'00708252    8b45fc                  mov eax, dword ptr [ebp-04]
'00708255    8b4dec                  mov ecx, dword ptr [ebp-14]
'00708258    5f                      pop edi
'00708259    5e                      pop esi
    'Reference to '__except_list'
'0070825a    64890d00000000          mov dword ptr fs:[00000000], ecx
'00708261    5b                      pop ebx
'00708262    8be5                    mov esp, ebp
'00708264    5d                      pop ebp
'00708265    c20400                  ret 0004


End Sub


'Event for btnAnnuler
Private Sub btnAnnuler_Click()
'00707fa0    55                      push ebp
'00707fa1    8bec                    mov ebp, esp
'00707fa3    83ec0c                  sub esp, 0c

' *** Reference to "__vbaExceptHandler"
'00707fa6    6866724000              push 00407266
'00707fab    64a100000000            mov ax, word ptr fs:[00000000]
'00707fb1    50                      push eax
    'Reference to '__except_list'
'00707fb2    64892500000000          mov dword ptr fs:[00000000], esp
'00707fb9    83ec18                  sub esp, 18
'00707fbc    53                      push ebx
'00707fbd    56                      push esi
'00707fbe    57                      push edi
'00707fbf    8965f4                  mov dword ptr [ebp-0c], esp
'00707fc2    c745f8d86b4000          mov dword ptr [ebp-08], 00406bd8
'00707fc9    8b7d08                  mov edi, dword ptr [ebp+08]
'00707fcc    8bc7                    mov eax, edi
'00707fce    83e001                  and eax, 01
'00707fd1    8945fc                  mov dword ptr [ebp-04], eax
'00707fd4    83e7fe                  and edi, fffffffe
'00707fd7    8b0f                    mov ecx, dword ptr [edi]
'00707fd9    57                      push edi
'00707fda    897d08                  mov dword ptr [ebp+08], edi
'00707fdd    ff5104                  call dword ptr [ecx+04]
'00707fe0    a124be7200              mov ax, word ptr [0072be24]
'00707fe5    33db                    xor ebx, ebx
'00707fe7    3bc3                    cmp eax, ebx
'00707fe9    895de8                  mov dword ptr [ebp-18], ebx
'00707fec    7510                    jne 707ffe

If (0 = 0) Then
'00707fee    6824be7200              push 0072be24
'00707ff3    6870004300              push 00430070

' *** Reference to "__vbaNew2"
'00707ff8    ff152c124000            call dword ptr [0040122c]
    Dim var_2 As New Global
End If
'00707ffe    8b3524be7200            mov esi, dword ptr [0072be24]
'00708004    8b16                    mov edx, dword ptr [esi]
'00708006    57                      push edi
'00708007    8d45e8                  lea eax, dword ptr [ebp-18]
'0070800a    50                      push eax
'0070800b    8955d4                  mov dword ptr [ebp-2c], edx

' *** Reference to "__vbaObjSetAddref"
'0070800e    ff15c8104000            call dword ptr [004010c8]
Set var_41 = Me
'00708014    8b4dd4                  mov ecx, dword ptr [ebp-2c]
'00708017    50                      push eax
'00708018    56                      push esi
'00708019    ff5110                  call dword ptr [ecx+10]
Call var_2.Unload(var_41)
'0070801c    dbe2                    fnclex
'0070801e    3bc3                    cmp eax, ebx
'00708020    7d0f                    jge 708031
'00708022    6a10                    push 10
'00708024    6860004300              push 00430060
'00708029    56                      push esi
'0070802a    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0070802b    ff1580104000            call dword ptr [00401080]
'00708031    8d4de8                  lea ecx, dword ptr [ebp-18]

' *** Reference to "__vbaFreeObj"
'00708034    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_41)
'0070803a    895dfc                  mov dword ptr [ebp-04], ebx
'0070803d    684f807000              push 0070804f
'00708042    eb0a                    jmp 70804e
'00708044    8d4de8                  lea ecx, dword ptr [ebp-18]

' *** Reference to "__vbaFreeObj"
'00708047    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_41)
'0070804d    c3                      ret
'0070804e    c3                      ret
'0070804f    8b4508                  mov eax, dword ptr [ebp+08]
'00708052    8b10                    mov edx, dword ptr [eax]
'00708054    50                      push eax
'00708055    ff5208                  call dword ptr [edx+08]
'00708058    8b45fc                  mov eax, dword ptr [ebp-04]
'0070805b    8b4dec                  mov ecx, dword ptr [ebp-14]
'0070805e    5f                      pop edi
'0070805f    5e                      pop esi
    'Reference to '__except_list'
'00708060    64890d00000000          mov dword ptr fs:[00000000], ecx
'00708067    5b                      pop ebx
'00708068    8be5                    mov esp, ebp
'0070806a    5d                      pop ebp
'0070806b    c20400                  ret 0004


End Sub


