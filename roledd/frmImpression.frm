VERSION 5.00

Begin VB.Form frmImpression
    Caption = "Type d'impression"
    ScaleMode = 1
    AutoRedraw = 0              'False
    FontTransparent = -1              'True
    BorderStyle = 3
    LinkTopic = "Form1"
    MaxButton = 0              'False
    MinButton = 0              'False
    ControlBox = 0              'False
    ClientLeft   = 45
    ClientTop    = 330
    ClientWidth  = 3000
    ClientHeight = 3570
    BeginProperty Font
        Name          = "Times New Roman"
        Size          = 9,75
        Charset       = 0
        Weight        = 400
        Underline     = 0              'False
        Italic        = 0              'False
        Strikethrough = 0              'False
    EndProperty
    ShowInTaskbar = 0              'False
    StartupPosition = 1
    Begin VB.CheckBox ChkImprim
        Caption = "Uniquement les sort quotidiens"
        Index = 4
        Left   = 120
        Top    = 2430
        Width  = 2820
        Height = 375
        TabIndex = 10
    End
    Begin VB.CheckBox ChkImprim
        Caption = "Sorts"
        Index = 3
        Left   = 1560
        Top    = 1900
        Width  = 1335
        Height = 375
        TabIndex = 9
    End
    Begin VB.CheckBox ChkImprim
        Caption = "Dons"
        Index = 2
        Left   = 120
        Top    = 1900
        Width  = 1335
        Height = 375
        TabIndex = 8
    End
    Begin VB.CheckBox ChkImprim
        Caption = "Histoire"
        Index = 1
        Left   = 1560
        Top    = 1400
        Width  = 1335
        Height = 375
        TabIndex = 7
    End
    Begin VB.CheckBox ChkImprim
        Caption = "background"
        Index = 0
        Left   = 120
        Top    = 1400
        Width  = 1335
        Height = 375
        TabIndex = 6
    End
    Begin VB.OptionButton Option4
        Caption = "Txt"
        Left   = 120
        Top    = 720
        Width  = 1200
        Height = 260
        TabIndex = 5
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
    Begin VB.OptionButton Option3
        Caption = "WordPerso"
        Left   = 1300
        Top    = 240
        Width  = 1200
        Height = 260
        TabIndex = 4
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
    Begin VB.CommandButton btnAnnuler
        Caption = "Annuler"
        Left   = 1560
        Top    = 3105
        Width  = 855
        Height = 330
        TabIndex = 3
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
    Begin VB.CommandButton btnOk
        Caption = "Ok"
        Left   = 600
        Top    = 3105
        Width  = 855
        Height = 330
        TabIndex = 2
        Default = -1              'True
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
    Begin VB.OptionButton Option2
        Caption = "HTML"
        Left   = 1300
        Top    = 720
        Width  = 1200
        Height = 260
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
    Begin VB.OptionButton Option1
        Caption = "Word"
        Left   = 120
        Top    = 240
        Width  = 1200
        Height = 260
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
End
'Event for btnOk
Private Sub btnOk_Click()
'006cf500    55                      push ebp
'006cf501    8bec                    mov ebp, esp
'006cf503    83ec0c                  sub esp, 0c

' *** Reference to "__vbaExceptHandler"
'006cf506    6866724000              push 00407266
'006cf50b    64a100000000            mov ax, word ptr fs:[00000000]
'006cf511    50                      push eax
    'Reference to '__except_list'
'006cf512    64892500000000          mov dword ptr fs:[00000000], esp
'006cf519    83ec28                  sub esp, 28
'006cf51c    53                      push ebx
'006cf51d    56                      push esi
'006cf51e    57                      push edi
'006cf51f    8965f4                  mov dword ptr [ebp-0c], esp
'006cf522    c745f878684000          mov dword ptr [ebp-08], 00406878
'006cf529    8b7508                  mov esi, dword ptr [ebp+08]
'006cf52c    8bc6                    mov eax, esi
'006cf52e    83e001                  and eax, 01
'006cf531    8945fc                  mov dword ptr [ebp-04], eax
'006cf534    83e6fe                  and esi, fffffffe
'006cf537    8b0e                    mov ecx, dword ptr [esi]
'006cf539    56                      push esi
'006cf53a    897508                  mov dword ptr [ebp+08], esi
'006cf53d    ff5104                  call dword ptr [ecx+04]
'006cf540    8b16                    mov edx, dword ptr [esi]
'006cf542    33db                    xor ebx, ebx
'006cf544    56                      push esi
'006cf545    895de8                  mov dword ptr [ebp-18], ebx
'006cf548    895de4                  mov dword ptr [ebp-1c], ebx
'006cf54b    895de0                  mov dword ptr [ebp-20], ebx
'006cf54e    ff9214030000            call dword ptr [edx+00000314]
'006cf554    50                      push eax
'006cf555    8d45e8                  lea eax, dword ptr [ebp-18]
'006cf558    50                      push eax

' *** Reference to "__vbaObjSet"
'006cf559    ff15b4104000            call dword ptr [004010b4]
Set var_41 = Me
'006cf55f    8d55e0                  lea edx, dword ptr [ebp-20]
'006cf562    8bf8                    mov edi, eax
'006cf564    8b0f                    mov ecx, dword ptr [edi]
'006cf566    52                      push edx
'006cf567    57                      push edi
'006cf568    ff91e0000000            call dword ptr [ecx+000000e0]
'006cf56e    dbe2                    fnclex
'006cf570    3bc3                    cmp eax, ebx
'006cf572    7d16                    jge 6cf58a

If (var_41 < 0) Then

' *** Reference to "__vbaHresultCheckObj"
'006cf574    8b1d80104000            mov ebx, dword ptr [00401080]
'006cf57a    68e0000000              push 000000e0
'006cf57f    68f43b4500              push 00453bf4
'006cf584    57                      push edi
'006cf585    50                      push eax
'006cf586    ffd3                    call ebx
'006cf588    eb06                    jmp 6cf590
    
End If

' *** Reference to "__vbaHresultCheckObj"
'006cf58a    8b1d80104000            mov ebx, dword ptr [00401080]
'006cf590    33c0                    xor eax, eax
var_num1 = Empty
'006cf592    66837de0ff              cmp word ptr [ebp-20], ffffffff
'006cf597    8d4de8                  lea ecx, dword ptr [ebp-18]
'006cf59a    0f94c0                  sete al
'006cf59d    f7d8                    neg eax
'006cf59f    668bf8                  mov di, ax

' *** Reference to "__vbaFreeObj"
'006cf5a2    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_41)
'006cf5a8    6685ff                  test edi, edi
'006cf5ab    7406                    je 6cf5b3

If (0 = -1) Then
'006cf5ad    66c7463c0100            mov word ptr [esi+3c], 0001
    
End If
'006cf5b3    8b0e                    mov ecx, dword ptr [esi]
'006cf5b5    56                      push esi
'006cf5b6    ff9110030000            call dword ptr [ecx+00000310]
'006cf5bc    50                      push eax
'006cf5bd    8d55e8                  lea edx, dword ptr [ebp-18]
'006cf5c0    52                      push edx

' *** Reference to "__vbaObjSet"
'006cf5c1    ff15b4104000            call dword ptr [004010b4]
Set var_41 = 0 = -1
'006cf5c7    8d4de0                  lea ecx, dword ptr [ebp-20]
'006cf5ca    8bf8                    mov edi, eax
'006cf5cc    8b07                    mov eax, dword ptr [edi]
'006cf5ce    51                      push ecx
'006cf5cf    57                      push edi
'006cf5d0    ff90e0000000            call dword ptr [eax+000000e0]
'006cf5d6    dbe2                    fnclex
'006cf5d8    85c0                    test ax, ax
'006cf5da    7d0e                    jge 6cf5ea
'006cf5dc    68e0000000              push 000000e0
'006cf5e1    68f43b4500              push 00453bf4
'006cf5e6    57                      push edi
'006cf5e7    50                      push eax
'006cf5e8    ffd3                    call ebx
'006cf5ea    33d2                    xor edx, edx
'006cf5ec    66837de0ff              cmp word ptr [ebp-20], ffffffff
'006cf5f1    8d4de8                  lea ecx, dword ptr [ebp-18]
'006cf5f4    0f94c2                  sete dl
'006cf5f7    f7da                    neg edx
'006cf5f9    668bfa                  mov di, dx

' *** Reference to "__vbaFreeObj"
'006cf5fc    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_41)
'006cf602    6685ff                  test edi, edi
'006cf605    7406                    je 6cf60d

If (0 = -1) Then
'006cf607    66c7463c0200            mov word ptr [esi+3c], 0002
    
End If
'006cf60d    8b06                    mov eax, dword ptr [esi]
'006cf60f    56                      push esi
'006cf610    ff9004030000            call dword ptr [eax+00000304]
'006cf616    50                      push eax
'006cf617    8d4de8                  lea ecx, dword ptr [ebp-18]
'006cf61a    51                      push ecx

' *** Reference to "__vbaObjSet"
'006cf61b    ff15b4104000            call dword ptr [004010b4]
Set var_41 = Nothing
'006cf621    8bf8                    mov edi, eax
'006cf623    8b17                    mov edx, dword ptr [edi]
'006cf625    8d45e0                  lea eax, dword ptr [ebp-20]
'006cf628    50                      push eax
'006cf629    57                      push edi
'006cf62a    ff92e0000000            call dword ptr [edx+000000e0]
'006cf630    dbe2                    fnclex
'006cf632    85c0                    test ax, ax
'006cf634    7d0e                    jge 6cf644
'006cf636    68e0000000              push 000000e0
'006cf63b    68f43b4500              push 00453bf4
'006cf640    57                      push edi
'006cf641    50                      push eax
'006cf642    ffd3                    call ebx
'006cf644    33c9                    xor ecx, ecx
'006cf646    66837de0ff              cmp word ptr [ebp-20], ffffffff
'006cf64b    0f94c1                  sete cl
'006cf64e    f7d9                    neg ecx
'006cf650    668bf9                  mov di, cx
'006cf653    8d4de8                  lea ecx, dword ptr [ebp-18]

' *** Reference to "__vbaFreeObj"
'006cf656    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_41)
'006cf65c    6685ff                  test edi, edi
'006cf65f    7406                    je 6cf667

If (0 = -1) Then
'006cf661    66c7463c0300            mov word ptr [esi+3c], 0003
    
End If
'006cf667    8b16                    mov edx, dword ptr [esi]
'006cf669    56                      push esi
'006cf66a    ff9200030000            call dword ptr [edx+00000300]
'006cf670    50                      push eax
'006cf671    8d45e8                  lea eax, dword ptr [ebp-18]
'006cf674    50                      push eax

' *** Reference to "__vbaObjSet"
'006cf675    ff15b4104000            call dword ptr [004010b4]
Set var_41 = 
'006cf67b    8d55e0                  lea edx, dword ptr [ebp-20]
'006cf67e    8bf8                    mov edi, eax
'006cf680    8b0f                    mov ecx, dword ptr [edi]
'006cf682    52                      push edx
'006cf683    57                      push edi
'006cf684    ff91e0000000            call dword ptr [ecx+000000e0]
'006cf68a    dbe2                    fnclex
'006cf68c    85c0                    test ax, ax
'006cf68e    7d0e                    jge 6cf69e
'006cf690    68e0000000              push 000000e0
'006cf695    68f43b4500              push 00453bf4
'006cf69a    57                      push edi
'006cf69b    50                      push eax
'006cf69c    ffd3                    call ebx
'006cf69e    33c0                    xor eax, eax
var_num1 = Empty
'006cf6a0    66837de0ff              cmp word ptr [ebp-20], ffffffff
'006cf6a5    8d4de8                  lea ecx, dword ptr [ebp-18]
'006cf6a8    0f94c0                  sete al
'006cf6ab    f7d8                    neg eax
'006cf6ad    668bf8                  mov di, ax

' *** Reference to "__vbaFreeObj"
'006cf6b0    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_41)
'006cf6b6    6685ff                  test edi, edi
'006cf6b9    7406                    je 6cf6c1

If (0 = -1) Then
'006cf6bb    66c7463c0400            mov word ptr [esi+3c], 0004
    
End If
'006cf6c1    8b0e                    mov ecx, dword ptr [esi]
'006cf6c3    56                      push esi
'006cf6c4    ff91fc020000            call dword ptr [ecx+000002fc]
'006cf6ca    50                      push eax
'006cf6cb    8d55e8                  lea edx, dword ptr [ebp-18]
'006cf6ce    52                      push edx

' *** Reference to "__vbaObjSet"
'006cf6cf    ff15b4104000            call dword ptr [004010b4]
Set var_41 = 0 = -1
'006cf6d5    8d4de4                  lea ecx, dword ptr [ebp-1c]
'006cf6d8    51                      push ecx
'006cf6d9    8bf8                    mov edi, eax
'006cf6db    8b07                    mov eax, dword ptr [edi]
'006cf6dd    6a00                    push 00
'006cf6df    57                      push edi
'006cf6e0    ff5040                  call dword ptr [eax+40]
'006cf6e3    dbe2                    fnclex
'006cf6e5    85c0                    test ax, ax
'006cf6e7    7d0b                    jge 6cf6f4
'006cf6e9    6a40                    push 40
'006cf6eb    686c384300              push 0043386c
'006cf6f0    57                      push edi
'006cf6f1    50                      push eax
'006cf6f2    ffd3                    call ebx
'006cf6f4    8b45e4                  mov eax, dword ptr [ebp-1c]
'006cf6f7    8b10                    mov edx, dword ptr [eax]
'006cf6f9    8d4de0                  lea ecx, dword ptr [ebp-20]
'006cf6fc    51                      push ecx
'006cf6fd    50                      push eax
'006cf6fe    8bf8                    mov edi, eax
'006cf700    ff92e0000000            call dword ptr [edx+000000e0]
'006cf706    dbe2                    fnclex
'006cf708    85c0                    test ax, ax
'006cf70a    7d0e                    jge 6cf71a
'006cf70c    68e0000000              push 000000e0
'006cf711    68dce24300              push 0043e2dc
'006cf716    57                      push edi
'006cf717    50                      push eax
'006cf718    ffd3                    call ebx
'006cf71a    33d2                    xor edx, edx
var_num4 = Empty
'006cf71c    663955e0                cmp word ptr [ebp-20], dx
'006cf720    8d45e4                  lea eax, dword ptr [ebp-1c]
'006cf723    0f95c2                  setne dl
'006cf726    50                      push eax
'006cf727    8d4de8                  lea ecx, dword ptr [ebp-18]
'006cf72a    51                      push ecx
'006cf72b    6a02                    push 02
'006cf72d    f7da                    neg edx
'006cf72f    66895634                mov word ptr [esi+34], dx

' *** Reference to "__vbaFreeObjList"
'006cf733    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_41, var_40)
'006cf739    8b16                    mov edx, dword ptr [esi]
'006cf73b    83c40c                  add esp, 0c
'006cf73e    56                      push esi
'006cf73f    ff92fc020000            call dword ptr [edx+000002fc]
'006cf745    50                      push eax
'006cf746    8d45e8                  lea eax, dword ptr [ebp-18]
'006cf749    50                      push eax

' *** Reference to "__vbaObjSet"
'006cf74a    ff15b4104000            call dword ptr [004010b4]
Set var_41 = 
'006cf750    8d55e4                  lea edx, dword ptr [ebp-1c]
'006cf753    52                      push edx
'006cf754    8bf8                    mov edi, eax
'006cf756    8b0f                    mov ecx, dword ptr [edi]
'006cf758    6a01                    push 01
'006cf75a    57                      push edi
'006cf75b    ff5140                  call dword ptr [ecx+40]
'006cf75e    dbe2                    fnclex
'006cf760    85c0                    test ax, ax
'006cf762    7d0b                    jge 6cf76f
'006cf764    6a40                    push 40
'006cf766    686c384300              push 0043386c
'006cf76b    57                      push edi
'006cf76c    50                      push eax
'006cf76d    ffd3                    call ebx
'006cf76f    8b45e4                  mov eax, dword ptr [ebp-1c]
'006cf772    8b08                    mov ecx, dword ptr [eax]
'006cf774    8d55e0                  lea edx, dword ptr [ebp-20]
'006cf777    52                      push edx
'006cf778    50                      push eax
'006cf779    8bf8                    mov edi, eax
'006cf77b    ff91e0000000            call dword ptr [ecx+000000e0]
'006cf781    dbe2                    fnclex
'006cf783    85c0                    test ax, ax
'006cf785    7d0e                    jge 6cf795
'006cf787    68e0000000              push 000000e0
'006cf78c    68dce24300              push 0043e2dc
'006cf791    57                      push edi
'006cf792    50                      push eax
'006cf793    ffd3                    call ebx
'006cf795    33c0                    xor eax, eax
var_num1 = Empty
'006cf797    663945e0                cmp word ptr [ebp-20], ax
'006cf79b    8d4de4                  lea ecx, dword ptr [ebp-1c]
'006cf79e    0f95c0                  setne al
'006cf7a1    51                      push ecx
'006cf7a2    8d55e8                  lea edx, dword ptr [ebp-18]
'006cf7a5    52                      push edx
'006cf7a6    6a02                    push 02
'006cf7a8    f7d8                    neg eax
'006cf7aa    66894636                mov word ptr [esi+36], ax

' *** Reference to "__vbaFreeObjList"
'006cf7ae    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_41, var_40)
'006cf7b4    8b06                    mov eax, dword ptr [esi]
'006cf7b6    83c40c                  add esp, 0c
'006cf7b9    56                      push esi
'006cf7ba    ff90fc020000            call dword ptr [eax+000002fc]
'006cf7c0    50                      push eax
'006cf7c1    8d4de8                  lea ecx, dword ptr [ebp-18]
'006cf7c4    51                      push ecx

' *** Reference to "__vbaObjSet"
'006cf7c5    ff15b4104000            call dword ptr [004010b4]
Set var_41 = Nothing
'006cf7cb    8bf8                    mov edi, eax
'006cf7cd    8b17                    mov edx, dword ptr [edi]
'006cf7cf    8d45e4                  lea eax, dword ptr [ebp-1c]
'006cf7d2    50                      push eax
'006cf7d3    6a02                    push 02
'006cf7d5    57                      push edi
'006cf7d6    ff5240                  call dword ptr [edx+40]
'006cf7d9    dbe2                    fnclex
'006cf7db    85c0                    test ax, ax
'006cf7dd    7d0b                    jge 6cf7ea
'006cf7df    6a40                    push 40
'006cf7e1    686c384300              push 0043386c
'006cf7e6    57                      push edi
'006cf7e7    50                      push eax
'006cf7e8    ffd3                    call ebx
'006cf7ea    8b45e4                  mov eax, dword ptr [ebp-1c]
'006cf7ed    8b08                    mov ecx, dword ptr [eax]
'006cf7ef    8d55e0                  lea edx, dword ptr [ebp-20]
'006cf7f2    52                      push edx
'006cf7f3    50                      push eax
'006cf7f4    8bf8                    mov edi, eax
'006cf7f6    ff91e0000000            call dword ptr [ecx+000000e0]
'006cf7fc    dbe2                    fnclex
'006cf7fe    85c0                    test ax, ax
'006cf800    7d0e                    jge 6cf810
'006cf802    68e0000000              push 000000e0
'006cf807    68dce24300              push 0043e2dc
'006cf80c    57                      push edi
'006cf80d    50                      push eax
'006cf80e    ffd3                    call ebx
'006cf810    33c0                    xor eax, eax
var_num1 = Empty
'006cf812    663945e0                cmp word ptr [ebp-20], ax
'006cf816    8d4de4                  lea ecx, dword ptr [ebp-1c]
'006cf819    0f95c0                  setne al
'006cf81c    51                      push ecx
'006cf81d    8d55e8                  lea edx, dword ptr [ebp-18]
'006cf820    52                      push edx
'006cf821    6a02                    push 02
'006cf823    f7d8                    neg eax
'006cf825    66894638                mov word ptr [esi+38], ax

' *** Reference to "__vbaFreeObjList"
'006cf829    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_41, var_40)
'006cf82f    8b06                    mov eax, dword ptr [esi]
'006cf831    83c40c                  add esp, 0c
'006cf834    56                      push esi
'006cf835    ff90fc020000            call dword ptr [eax+000002fc]
'006cf83b    50                      push eax
'006cf83c    8d4de8                  lea ecx, dword ptr [ebp-18]
'006cf83f    51                      push ecx

' *** Reference to "__vbaObjSet"
'006cf840    ff15b4104000            call dword ptr [004010b4]
Set var_41 = Nothing
'006cf846    8bf8                    mov edi, eax
'006cf848    8b17                    mov edx, dword ptr [edi]
'006cf84a    8d45e4                  lea eax, dword ptr [ebp-1c]
'006cf84d    50                      push eax
'006cf84e    6a03                    push 03
'006cf850    57                      push edi
'006cf851    ff5240                  call dword ptr [edx+40]
'006cf854    dbe2                    fnclex
'006cf856    85c0                    test ax, ax
'006cf858    7d0b                    jge 6cf865
'006cf85a    6a40                    push 40
'006cf85c    686c384300              push 0043386c
'006cf861    57                      push edi
'006cf862    50                      push eax
'006cf863    ffd3                    call ebx
'006cf865    8b45e4                  mov eax, dword ptr [ebp-1c]
'006cf868    8b08                    mov ecx, dword ptr [eax]
'006cf86a    8d55e0                  lea edx, dword ptr [ebp-20]
'006cf86d    52                      push edx
'006cf86e    50                      push eax
'006cf86f    8bf8                    mov edi, eax
'006cf871    ff91e0000000            call dword ptr [ecx+000000e0]
'006cf877    dbe2                    fnclex
'006cf879    85c0                    test ax, ax
'006cf87b    7d0e                    jge 6cf88b
'006cf87d    68e0000000              push 000000e0
'006cf882    68dce24300              push 0043e2dc
'006cf887    57                      push edi
'006cf888    50                      push eax
'006cf889    ffd3                    call ebx
'006cf88b    33c0                    xor eax, eax
var_num1 = Empty
'006cf88d    663945e0                cmp word ptr [ebp-20], ax
'006cf891    8d4de4                  lea ecx, dword ptr [ebp-1c]
'006cf894    0f95c0                  setne al
'006cf897    51                      push ecx
'006cf898    8d55e8                  lea edx, dword ptr [ebp-18]
'006cf89b    52                      push edx
'006cf89c    6a02                    push 02
'006cf89e    f7d8                    neg eax
'006cf8a0    6689463a                mov word ptr [esi+3a], ax

' *** Reference to "__vbaFreeObjList"
'006cf8a4    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_41, var_40)
'006cf8aa    8b06                    mov eax, dword ptr [esi]
'006cf8ac    83c40c                  add esp, 0c
'006cf8af    56                      push esi
'006cf8b0    ff90fc020000            call dword ptr [eax+000002fc]
'006cf8b6    50                      push eax
'006cf8b7    8d4de8                  lea ecx, dword ptr [ebp-18]
'006cf8ba    51                      push ecx

' *** Reference to "__vbaObjSet"
'006cf8bb    ff15b4104000            call dword ptr [004010b4]
Set var_41 = Nothing
'006cf8c1    8bf8                    mov edi, eax
'006cf8c3    8b17                    mov edx, dword ptr [edi]
'006cf8c5    8d45e4                  lea eax, dword ptr [ebp-1c]
'006cf8c8    50                      push eax
'006cf8c9    6a04                    push 04
'006cf8cb    57                      push edi
'006cf8cc    ff5240                  call dword ptr [edx+40]
'006cf8cf    dbe2                    fnclex
'006cf8d1    85c0                    test ax, ax
'006cf8d3    7d0b                    jge 6cf8e0
'006cf8d5    6a40                    push 40
'006cf8d7    686c384300              push 0043386c
'006cf8dc    57                      push edi
'006cf8dd    50                      push eax
'006cf8de    ffd3                    call ebx
'006cf8e0    8b45e4                  mov eax, dword ptr [ebp-1c]
'006cf8e3    8b08                    mov ecx, dword ptr [eax]
'006cf8e5    8d55e0                  lea edx, dword ptr [ebp-20]
'006cf8e8    52                      push edx
'006cf8e9    50                      push eax
'006cf8ea    8bf8                    mov edi, eax
'006cf8ec    ff91e0000000            call dword ptr [ecx+000000e0]
'006cf8f2    dbe2                    fnclex
'006cf8f4    85c0                    test ax, ax
'006cf8f6    7d0e                    jge 6cf906
'006cf8f8    68e0000000              push 000000e0
'006cf8fd    68dce24300              push 0043e2dc
'006cf902    57                      push edi
'006cf903    50                      push eax
'006cf904    ffd3                    call ebx
'006cf906    33c0                    xor eax, eax
var_num1 = Empty
'006cf908    663945e0                cmp word ptr [ebp-20], ax
'006cf90c    8d4de4                  lea ecx, dword ptr [ebp-1c]
'006cf90f    0f95c0                  setne al
'006cf912    51                      push ecx
'006cf913    8d55e8                  lea edx, dword ptr [ebp-18]
'006cf916    52                      push edx
'006cf917    6a02                    push 02
'006cf919    f7d8                    neg eax
'006cf91b    6689463e                mov word ptr [esi+3e], ax

' *** Reference to "__vbaFreeObjList"
'006cf91f    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_41, var_40)
'006cf925    a124be7200              mov ax, word ptr [0072be24]
'006cf92a    83c40c                  add esp, 0c
'006cf92d    85c0                    test ax, ax
'006cf92f    7510                    jne 6cf941
'006cf931    6824be7200              push 0072be24
'006cf936    6870004300              push 00430070

' *** Reference to "__vbaNew2"
'006cf93b    ff152c124000            call dword ptr [0040122c]
Dim var_2 As New Global
'006cf941    8b3d24be7200            mov edi, dword ptr [0072be24]
'006cf947    8b17                    mov edx, dword ptr [edi]
'006cf949    56                      push esi
'006cf94a    8d45e8                  lea eax, dword ptr [ebp-18]
'006cf94d    50                      push eax
'006cf94e    8955c4                  mov dword ptr [ebp-3c], edx

' *** Reference to "__vbaObjSetAddref"
'006cf951    ff15c8104000            call dword ptr [004010c8]
Set var_41 = Me
'006cf957    8b4dc4                  mov ecx, dword ptr [ebp-3c]
'006cf95a    50                      push eax
'006cf95b    57                      push edi
'006cf95c    ff5110                  call dword ptr [ecx+10]
Call var_2.Unload(var_41)
'006cf95f    dbe2                    fnclex
'006cf961    85c0                    test ax, ax
'006cf963    7d0b                    jge 6cf970
'006cf965    6a10                    push 10
'006cf967    6860004300              push 00430060
'006cf96c    57                      push edi
'006cf96d    50                      push eax
'006cf96e    ffd3                    call ebx
'006cf970    8d4de8                  lea ecx, dword ptr [ebp-18]

' *** Reference to "__vbaFreeObj"
'006cf973    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_41)
'006cf979    c745fc00000000          mov dword ptr [ebp-04], 00000000
'006cf980    689cf96c00              push 006cf99c
'006cf985    eb14                    jmp 6cf99b
'006cf987    8d55e4                  lea edx, dword ptr [ebp-1c]
'006cf98a    52                      push edx
'006cf98b    8d45e8                  lea eax, dword ptr [ebp-18]
'006cf98e    50                      push eax
'006cf98f    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'006cf991    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_41, var_40)
'006cf997    83c40c                  add esp, 0c
'006cf99a    c3                      ret
'006cf99b    c3                      ret
'006cf99c    8b4508                  mov eax, dword ptr [ebp+08]
'006cf99f    8b08                    mov ecx, dword ptr [eax]
'006cf9a1    50                      push eax
'006cf9a2    ff5108                  call dword ptr [ecx+08]
'006cf9a5    8b45fc                  mov eax, dword ptr [ebp-04]
'006cf9a8    8b4dec                  mov ecx, dword ptr [ebp-14]
'006cf9ab    5f                      pop edi
'006cf9ac    5e                      pop esi
    'Reference to '__except_list'
'006cf9ad    64890d00000000          mov dword ptr fs:[00000000], ecx
'006cf9b4    5b                      pop ebx
'006cf9b5    8be5                    mov esp, ebp
'006cf9b7    5d                      pop ebp
'006cf9b8    c20400                  ret 0004


End Sub


'Event for ChkImprim
Private Sub ChkImprim_Click()
'006cf9c0    55                      push ebp
'006cf9c1    8bec                    mov ebp, esp
'006cf9c3    83ec0c                  sub esp, 0c

' *** Reference to "__vbaExceptHandler"
'006cf9c6    6866724000              push 00407266
'006cf9cb    64a100000000            mov ax, word ptr fs:[00000000]
'006cf9d1    50                      push eax
    'Reference to '__except_list'
'006cf9d2    64892500000000          mov dword ptr fs:[00000000], esp
'006cf9d9    83ec34                  sub esp, 34
'006cf9dc    53                      push ebx
'006cf9dd    56                      push esi
'006cf9de    57                      push edi
'006cf9df    8965f4                  mov dword ptr [ebp-0c], esp
'006cf9e2    c745f888684000          mov dword ptr [ebp-08], 00406888
'006cf9e9    8b7508                  mov esi, dword ptr [ebp+08]
'006cf9ec    8bc6                    mov eax, esi
'006cf9ee    83e001                  and eax, 01
'006cf9f1    8945fc                  mov dword ptr [ebp-04], eax
'006cf9f4    83e6fe                  and esi, fffffffe
'006cf9f7    8b0e                    mov ecx, dword ptr [esi]
'006cf9f9    56                      push esi
'006cf9fa    897508                  mov dword ptr [ebp+08], esi
'006cf9fd    ff5104                  call dword ptr [ecx+04]
'006cfa00    8b550c                  mov edx, dword ptr [ebp+0c]
'006cfa03    33c0                    xor eax, eax
var_num1 = Empty
'006cfa05    66833a03                cmp word ptr [edx], 03
'006cfa09    8945e8                  mov dword ptr [ebp-18], eax
'006cfa0c    8945e4                  mov dword ptr [ebp-1c], eax
'006cfa0f    8945d4                  mov dword ptr [ebp-2c], eax
'006cfa12    0f8543010000            jne 6cfb5b

If (arg_0 = 3) Then
'006cfa18    8b06                    mov eax, dword ptr [esi]
'006cfa1a    56                      push esi
'006cfa1b    ff90fc020000            call dword ptr [eax+000002fc]

' *** Reference to "__vbaObjSet"
'006cfa21    8b1db4104000            mov ebx, dword ptr [004010b4]
'006cfa27    50                      push eax
'006cfa28    8d4de8                  lea ecx, dword ptr [ebp-18]
'006cfa2b    51                      push ecx
'006cfa2c    ffd3                    call ebx
    Set var_41 = Nothing
'006cfa2e    8bf8                    mov edi, eax
'006cfa30    8b17                    mov edx, dword ptr [edi]
'006cfa32    8d45e4                  lea eax, dword ptr [ebp-1c]
'006cfa35    50                      push eax
'006cfa36    6a03                    push 03
'006cfa38    57                      push edi
'006cfa39    ff5240                  call dword ptr [edx+40]
'006cfa3c    dbe2                    fnclex
'006cfa3e    85c0                    test ax, ax
'006cfa40    7d13                    jge 6cfa55
'006cfa42    6a40                    push 40
'006cfa44    686c384300              push 0043386c
'006cfa49    57                      push edi

' *** Reference to "__vbaHresultCheckObj"
'006cfa4a    8b3d80104000            mov edi, dword ptr [00401080]
'006cfa50    50                      push eax
'006cfa51    ffd7                    call edi
'006cfa53    eb06                    jmp 6cfa5b

' *** Reference to "__vbaHresultCheckObj"
'006cfa55    8b3d80104000            mov edi, dword ptr [00401080]
'006cfa5b    8b45e4                  mov eax, dword ptr [ebp-1c]
'006cfa5e    8d4dd4                  lea ecx, dword ptr [ebp-2c]
'006cfa61    51                      push ecx
'006cfa62    c745e400000000          mov dword ptr [ebp-1c], 00000000
'006cfa69    8945dc                  mov dword ptr [ebp-24], eax
'006cfa6c    c745d409000000          mov dword ptr [ebp-2c], 00000009

' *** Reference to "__vbaBoolVarNull"
'006cfa73    ff15fc104000            call dword ptr [004010fc]
'006cfa79    8d4de8                  lea ecx, dword ptr [ebp-18]
'006cfa7c    668945c8                mov word ptr [ebp-38], ax

' *** Reference to "__vbaFreeObj"
'006cfa80    ff1508134000            call dword ptr [00401308]
    '#Cleanup(var_41)
'006cfa86    8d4dd4                  lea ecx, dword ptr [ebp-2c]

' *** Reference to "__vbaFreeVar"
'006cfa89    ff1528104000            call dword ptr [00401028]
    '#Cleanup(var_44)
'006cfa8f    66837dc800              cmp word ptr [ebp-38], 00
'006cfa94    745d                    je 6cfaf3
    
    If (    CBool(Me) <> 0) Then
'006cfa96    8b16                    mov edx, dword ptr [esi]
'006cfa98    56                      push esi
'006cfa99    ff92fc020000            call dword ptr [edx+000002fc]
'006cfa9f    50                      push eax
'006cfaa0    8d45e8                  lea eax, dword ptr [ebp-18]
'006cfaa3    50                      push eax
'006cfaa4    ffd3                    call ebx
    Set var_41 = CBool(Me)
'006cfaa6    8d55e4                  lea edx, dword ptr [ebp-1c]
'006cfaa9    52                      push edx
'006cfaaa    8bf0                    mov esi, eax
'006cfaac    8b0e                    mov ecx, dword ptr [esi]
'006cfaae    6a04                    push 04
'006cfab0    56                      push esi
'006cfab1    ff5140                  call dword ptr [ecx+40]
'006cfab4    dbe2                    fnclex
'006cfab6    85c0                    test ax, ax
'006cfab8    7d0b                    jge 6cfac5
    
    If (    var_41) Then
'006cfaba    6a40                    push 40
'006cfabc    686c384300              push 0043386c
'006cfac1    56                      push esi
'006cfac2    50                      push eax
'006cfac3    ffd7                    call edi
    
End If
'006cfac5    8b45e4                  mov eax, dword ptr [ebp-1c]
'006cfac8    8b08                    mov ecx, dword ptr [eax]
'006cfaca    6aff                    push ffffffff
'006cfacc    50                      push eax
'006cfacd    8bf0                    mov esi, eax
'006cfacf    ff9194000000            call dword ptr [ecx+00000094]
'006cfad5    dbe2                    fnclex
'006cfad7    85c0                    test ax, ax
'006cfad9    7d0e                    jge 6cfae9
'006cfadb    6894000000              push 00000094
'006cfae0    68dce24300              push 0043e2dc
'006cfae5    56                      push esi
'006cfae6    50                      push eax
'006cfae7    ffd7                    call edi
'006cfae9    8d55e4                  lea edx, dword ptr [ebp-1c]
'006cfaec    52                      push edx
'006cfaed    8d45e8                  lea eax, dword ptr [ebp-18]
'006cfaf0    50                      push eax
'006cfaf1    eb5b                    jmp 6cfb4e

'ERROR: Two many next close:
End If
'006cfaf3    8b0e                    mov ecx, dword ptr [esi]
'006cfaf5    56                      push esi
'006cfaf6    ff91fc020000            call dword ptr [ecx+000002fc]
'006cfafc    50                      push eax
'006cfafd    8d55e8                  lea edx, dword ptr [ebp-18]
'006cfb00    52                      push edx
'006cfb01    ffd3                    call ebx
Set var_41 = CBool(Me)
'006cfb03    8d4de4                  lea ecx, dword ptr [ebp-1c]
'006cfb06    51                      push ecx
'006cfb07    8bf0                    mov esi, eax
'006cfb09    8b06                    mov eax, dword ptr [esi]
'006cfb0b    6a04                    push 04
'006cfb0d    56                      push esi
'006cfb0e    ff5040                  call dword ptr [eax+40]
'006cfb11    dbe2                    fnclex
'006cfb13    85c0                    test ax, ax
'006cfb15    7d0b                    jge 6cfb22
'006cfb17    6a40                    push 40
'006cfb19    686c384300              push 0043386c
'006cfb1e    56                      push esi
'006cfb1f    50                      push eax
'006cfb20    ffd7                    call edi
'006cfb22    8b45e4                  mov eax, dword ptr [ebp-1c]
'006cfb25    8b10                    mov edx, dword ptr [eax]
'006cfb27    6a00                    push 00
'006cfb29    50                      push eax
'006cfb2a    8bf0                    mov esi, eax
'006cfb2c    ff9294000000            call dword ptr [edx+00000094]
'006cfb32    dbe2                    fnclex
'006cfb34    85c0                    test ax, ax
'006cfb36    7d0e                    jge 6cfb46
'006cfb38    6894000000              push 00000094
'006cfb3d    68dce24300              push 0043e2dc
'006cfb42    56                      push esi
'006cfb43    50                      push eax
'006cfb44    ffd7                    call edi
'006cfb46    8d45e4                  lea eax, dword ptr [ebp-1c]
'006cfb49    50                      push eax
'006cfb4a    8d4de8                  lea ecx, dword ptr [ebp-18]
'006cfb4d    51                      push ecx
'006cfb4e    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'006cfb50    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_41, var_40)
'006cfb56    83c40c                  add esp, 0c
'006cfb59    33c0                    xor eax, eax

'ERROR: Two many next close:
End If
'006cfb5b    8945fc                  mov dword ptr [ebp-04], eax
'006cfb5e    6883fb6c00              push 006cfb83
'006cfb63    eb1d                    jmp 6cfb82
'006cfb65    8d55e4                  lea edx, dword ptr [ebp-1c]
'006cfb68    52                      push edx
'006cfb69    8d45e8                  lea eax, dword ptr [ebp-18]
'006cfb6c    50                      push eax
'006cfb6d    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'006cfb6f    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_41, var_40)
'006cfb75    83c40c                  add esp, 0c
'006cfb78    8d4dd4                  lea ecx, dword ptr [ebp-2c]

' *** Reference to "__vbaFreeVar"
'006cfb7b    ff1528104000            call dword ptr [00401028]
'#Cleanup(var_44)
'006cfb81    c3                      ret
'006cfb82    c3                      ret
'006cfb83    8b4508                  mov eax, dword ptr [ebp+08]
'006cfb86    8b08                    mov ecx, dword ptr [eax]
'006cfb88    50                      push eax
'006cfb89    ff5108                  call dword ptr [ecx+08]
'006cfb8c    8b45fc                  mov eax, dword ptr [ebp-04]
'006cfb8f    8b4dec                  mov ecx, dword ptr [ebp-14]
'006cfb92    5f                      pop edi
'006cfb93    5e                      pop esi
    'Reference to '__except_list'
'006cfb94    64890d00000000          mov dword ptr fs:[00000000], ecx
'006cfb9b    5b                      pop ebx
'006cfb9c    8be5                    mov esp, ebp
'006cfb9e    5d                      pop ebp
'006cfb9f    c20800                  ret 0008


End Sub


'Event for Form
Private Sub Form_Load()
'006cfbb0    55                      push ebp
'006cfbb1    8bec                    mov ebp, esp
'006cfbb3    83ec0c                  sub esp, 0c

' *** Reference to "__vbaExceptHandler"
'006cfbb6    6866724000              push 00407266
'006cfbbb    64a100000000            mov ax, word ptr fs:[00000000]
'006cfbc1    50                      push eax
    'Reference to '__except_list'
'006cfbc2    64892500000000          mov dword ptr fs:[00000000], esp
'006cfbc9    83ec20                  sub esp, 20
'006cfbcc    53                      push ebx
'006cfbcd    56                      push esi
'006cfbce    57                      push edi
'006cfbcf    8965f4                  mov dword ptr [ebp-0c], esp
'006cfbd2    c745f898684000          mov dword ptr [ebp-08], 00406898
'006cfbd9    8b7508                  mov esi, dword ptr [ebp+08]
'006cfbdc    8bc6                    mov eax, esi
'006cfbde    83e001                  and eax, 01
'006cfbe1    8945fc                  mov dword ptr [ebp-04], eax
'006cfbe4    83e6fe                  and esi, fffffffe
'006cfbe7    8b0e                    mov ecx, dword ptr [esi]
'006cfbe9    56                      push esi
'006cfbea    897508                  mov dword ptr [ebp+08], esi
'006cfbed    ff5104                  call dword ptr [ecx+04]
'006cfbf0    8b16                    mov edx, dword ptr [esi]
'006cfbf2    33ff                    xor edi, edi
'006cfbf4    56                      push esi
'006cfbf5    897de8                  mov dword ptr [ebp-18], edi
'006cfbf8    897de4                  mov dword ptr [ebp-1c], edi
'006cfbfb    66897e3c                mov word ptr [esi+3c], di
'006cfbff    ff92fc020000            call dword ptr [edx+000002fc]
'006cfc05    50                      push eax
'006cfc06    8d45e8                  lea eax, dword ptr [ebp-18]
'006cfc09    50                      push eax

' *** Reference to "__vbaObjSet"
'006cfc0a    ff15b4104000            call dword ptr [004010b4]
Set var_41 = Me
'006cfc10    8d55e4                  lea edx, dword ptr [ebp-1c]
'006cfc13    52                      push edx
'006cfc14    8bf0                    mov esi, eax
'006cfc16    8b0e                    mov ecx, dword ptr [esi]
'006cfc18    6a04                    push 04
'006cfc1a    56                      push esi
'006cfc1b    ff5140                  call dword ptr [ecx+40]
'006cfc1e    dbe2                    fnclex
'006cfc20    3bc7                    cmp eax, edi
'006cfc22    7d0f                    jge 6cfc33

If (var_41 < 0) Then
'006cfc24    6a40                    push 40
'006cfc26    686c384300              push 0043386c
'006cfc2b    56                      push esi
'006cfc2c    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006cfc2d    ff1580104000            call dword ptr [00401080]
    
End If
'006cfc33    8b45e4                  mov eax, dword ptr [ebp-1c]
'006cfc36    8b08                    mov ecx, dword ptr [eax]
'006cfc38    57                      push edi
'006cfc39    50                      push eax
'006cfc3a    8bf0                    mov esi, eax
'006cfc3c    ff9194000000            call dword ptr [ecx+00000094]
'006cfc42    dbe2                    fnclex
'006cfc44    3bc7                    cmp eax, edi
'006cfc46    7d12                    jge 6cfc5a

If (0 < 0) Then
'006cfc48    6894000000              push 00000094
'006cfc4d    68dce24300              push 0043e2dc
'006cfc52    56                      push esi
'006cfc53    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006cfc54    ff1580104000            call dword ptr [00401080]
    
End If
'006cfc5a    8d55e4                  lea edx, dword ptr [ebp-1c]
'006cfc5d    52                      push edx
'006cfc5e    8d45e8                  lea eax, dword ptr [ebp-18]
'006cfc61    50                      push eax
'006cfc62    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'006cfc64    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_41, var_40)
'006cfc6a    83c40c                  add esp, 0c
'006cfc6d    897dfc                  mov dword ptr [ebp-04], edi
'006cfc70    688cfc6c00              push 006cfc8c
'006cfc75    eb14                    jmp 6cfc8b
'006cfc77    8d4de4                  lea ecx, dword ptr [ebp-1c]
'006cfc7a    51                      push ecx
'006cfc7b    8d55e8                  lea edx, dword ptr [ebp-18]
'006cfc7e    52                      push edx
'006cfc7f    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'006cfc81    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_41, var_40)
'006cfc87    83c40c                  add esp, 0c
'006cfc8a    c3                      ret
'006cfc8b    c3                      ret
'006cfc8c    8b4508                  mov eax, dword ptr [ebp+08]
'006cfc8f    8b08                    mov ecx, dword ptr [eax]
'006cfc91    50                      push eax
'006cfc92    ff5108                  call dword ptr [ecx+08]
'006cfc95    8b45fc                  mov eax, dword ptr [ebp-04]
'006cfc98    8b4dec                  mov ecx, dword ptr [ebp-14]
'006cfc9b    5f                      pop edi
'006cfc9c    5e                      pop esi
    'Reference to '__except_list'
'006cfc9d    64890d00000000          mov dword ptr fs:[00000000], ecx
'006cfca4    5b                      pop ebx
'006cfca5    8be5                    mov esp, ebp
'006cfca7    5d                      pop ebp
'006cfca8    c20400                  ret 0004


End Sub


'Event for btnAnnuler
Private Sub btnAnnuler_Click()
'006cf430    55                      push ebp
'006cf431    8bec                    mov ebp, esp
'006cf433    83ec0c                  sub esp, 0c

' *** Reference to "__vbaExceptHandler"
'006cf436    6866724000              push 00407266
'006cf43b    64a100000000            mov ax, word ptr fs:[00000000]
'006cf441    50                      push eax
    'Reference to '__except_list'
'006cf442    64892500000000          mov dword ptr fs:[00000000], esp
'006cf449    83ec18                  sub esp, 18
'006cf44c    53                      push ebx
'006cf44d    56                      push esi
'006cf44e    57                      push edi
'006cf44f    8965f4                  mov dword ptr [ebp-0c], esp
'006cf452    c745f868684000          mov dword ptr [ebp-08], 00406868
'006cf459    8b7d08                  mov edi, dword ptr [ebp+08]
'006cf45c    8bc7                    mov eax, edi
'006cf45e    83e001                  and eax, 01
'006cf461    8945fc                  mov dword ptr [ebp-04], eax
'006cf464    83e7fe                  and edi, fffffffe
'006cf467    8b0f                    mov ecx, dword ptr [edi]
'006cf469    57                      push edi
'006cf46a    897d08                  mov dword ptr [ebp+08], edi
'006cf46d    ff5104                  call dword ptr [ecx+04]
'006cf470    a124be7200              mov ax, word ptr [0072be24]
'006cf475    33db                    xor ebx, ebx
'006cf477    3bc3                    cmp eax, ebx
'006cf479    895de8                  mov dword ptr [ebp-18], ebx
'006cf47c    7510                    jne 6cf48e

If (0 = 0) Then
'006cf47e    6824be7200              push 0072be24
'006cf483    6870004300              push 00430070

' *** Reference to "__vbaNew2"
'006cf488    ff152c124000            call dword ptr [0040122c]
    Dim var_2 As New Global
End If
'006cf48e    8b3524be7200            mov esi, dword ptr [0072be24]
'006cf494    8b16                    mov edx, dword ptr [esi]
'006cf496    57                      push edi
'006cf497    8d45e8                  lea eax, dword ptr [ebp-18]
'006cf49a    50                      push eax
'006cf49b    8955d4                  mov dword ptr [ebp-2c], edx

' *** Reference to "__vbaObjSetAddref"
'006cf49e    ff15c8104000            call dword ptr [004010c8]
Set var_41 = Me
'006cf4a4    8b4dd4                  mov ecx, dword ptr [ebp-2c]
'006cf4a7    50                      push eax
'006cf4a8    56                      push esi
'006cf4a9    ff5110                  call dword ptr [ecx+10]
Call var_2.Unload(var_41)
'006cf4ac    dbe2                    fnclex
'006cf4ae    3bc3                    cmp eax, ebx
'006cf4b0    7d0f                    jge 6cf4c1
'006cf4b2    6a10                    push 10
'006cf4b4    6860004300              push 00430060
'006cf4b9    56                      push esi
'006cf4ba    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006cf4bb    ff1580104000            call dword ptr [00401080]
'006cf4c1    8d4de8                  lea ecx, dword ptr [ebp-18]

' *** Reference to "__vbaFreeObj"
'006cf4c4    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_41)
'006cf4ca    895dfc                  mov dword ptr [ebp-04], ebx
'006cf4cd    68dff46c00              push 006cf4df
'006cf4d2    eb0a                    jmp 6cf4de
'006cf4d4    8d4de8                  lea ecx, dword ptr [ebp-18]

' *** Reference to "__vbaFreeObj"
'006cf4d7    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_41)
'006cf4dd    c3                      ret
'006cf4de    c3                      ret
'006cf4df    8b4508                  mov eax, dword ptr [ebp+08]
'006cf4e2    8b10                    mov edx, dword ptr [eax]
'006cf4e4    50                      push eax
'006cf4e5    ff5208                  call dword ptr [edx+08]
'006cf4e8    8b45fc                  mov eax, dword ptr [ebp-04]
'006cf4eb    8b4dec                  mov ecx, dword ptr [ebp-14]
'006cf4ee    5f                      pop edi
'006cf4ef    5e                      pop esi
    'Reference to '__except_list'
'006cf4f0    64890d00000000          mov dword ptr fs:[00000000], ecx
'006cf4f7    5b                      pop ebx
'006cf4f8    8be5                    mov esp, ebp
'006cf4fa    5d                      pop ebp
'006cf4fb    c20400                  ret 0004


End Sub


