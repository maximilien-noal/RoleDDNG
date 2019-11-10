VERSION 5.00

Begin VB.Form frmApropos
    Caption = "A propos de RoleDD"
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
    ClientWidth  = 6165
    ClientHeight = 3825
    BeginProperty Font
        Name          = "Times New Roman"
        Size          = 12
        Charset       = 0
        Weight        = 400
        Underline     = 0              'False
        Italic        = 0              'False
        Strikethrough = 0              'False
    EndProperty
    ShowInTaskbar = 0              'False
    StartupPosition = 1
    Begin VB.PictureBox Picture1
        Picture = frmApropos.frx:0000
        Left   = 4560
        Top    = 120
        Width  = 1095
        Height = 975
        TabIndex = 5
        ScaleMode = 1
        AutoRedraw = 0              'False
        FontTransparent = -1              'True
    End
    Begin VB.TextBox fldStrVersion
        ForeColor = 0
        Left   = 1710
        Top    = 2205
        Width  = 3855
        Height = 975
        TabIndex = 2
        MultiLine = -1              'True
        Locked = -1              'True
        BeginProperty Font
            Name          = "Times New Roman"
            Size          = 12
            Charset       = 0
            Weight        = 700
            Underline     = 0              'False
            Italic        = 0              'False
            Strikethrough = 0              'False
        EndProperty
    End
    Begin VB.CommandButton btnOk
        Caption = "&Ok"
        Left   = 4545
        Top    = 3375
        Width  = 1005
        Height = 330
        TabIndex = 1
        Cancel = -1              'True
        BeginProperty Font
            Name          = "MS Sans Serif"
            Size          = 8,25
            Charset       = 0
            Weight        = 400
            Underline     = 0              'False
            Italic        = 0              'False
            Strikethrough = 0              'False
        EndProperty
    End
    Begin VB.Label Label2
        Caption = "Versions :"
        Index = 2
        ForeColor = 0
        Left   = 240
        Top    = 2520
        Width  = 1335
        Height = 375
        TabIndex = 4
        Alignment = 2
        BackStyle = 0
        BeginProperty Font
            Name          = "Times New Roman"
            Size          = 15,75
            Charset       = 0
            Weight        = 400
            Underline     = 0              'False
            Italic        = 0              'False
            Strikethrough = 0              'False
        EndProperty
    End
    Begin VB.Label Label2
        Caption = "http://bonnarien.dyndns.org"
        Index = 1
        ForeColor = 12582912
        Left   = 180
        Top    = 1620
        Width  = 4380
        Height = 465
        TabIndex = 3
        Alignment = 2
        BackStyle = 0
        BeginProperty Font
            Name          = "Times New Roman"
            Size          = 15,75
            Charset       = 0
            Weight        = 400
            Underline     = -1              'True
            Italic        = 0              'False
            Strikethrough = 0              'False
        EndProperty
    End
    Begin VB.Label Label2
        Caption = "Gestion des feuilles de personnages pour Donjons et Dragons 3.5"
        Index = 0
        ForeColor = 0
        Left   = 135
        Top    = 135
        Width  = 4380
        Height = 1230
        TabIndex = 0
        Alignment = 2
        BackStyle = 0
        BeginProperty Font
            Name          = "Times New Roman"
            Size          = 15,75
            Charset       = 0
            Weight        = 400
            Underline     = 0              'False
            Italic        = 0              'False
            Strikethrough = 0              'False
        EndProperty
    End
End
'Event for Label2
Private Sub Label2_Click()
'006cc6f0    55                      push ebp
'006cc6f1    8bec                    mov ebp, esp
'006cc6f3    83ec0c                  sub esp, 0c

' *** Reference to "__vbaExceptHandler"
'006cc6f6    6866724000              push 00407266
'006cc6fb    64a100000000            mov ax, word ptr fs:[00000000]
'006cc701    50                      push eax
    'Reference to '__except_list'
'006cc702    64892500000000          mov dword ptr fs:[00000000], esp
'006cc709    83ec30                  sub esp, 30
'006cc70c    53                      push ebx
'006cc70d    56                      push esi
'006cc70e    57                      push edi
'006cc70f    8965f4                  mov dword ptr [ebp-0c], esp
'006cc712    c745f8c8674000          mov dword ptr [ebp-08], 004067c8
'006cc719    8b4508                  mov eax, dword ptr [ebp+08]
'006cc71c    8bc8                    mov ecx, eax
'006cc71e    83e101                  and ecx, 01
'006cc721    894dfc                  mov dword ptr [ebp-04], ecx
'006cc724    83e0fe                  and eax, fffffffe
'006cc727    8b10                    mov edx, dword ptr [eax]
'006cc729    50                      push eax
'006cc72a    894508                  mov dword ptr [ebp+08], eax
'006cc72d    ff5204                  call dword ptr [edx+04]
'006cc730    8b450c                  mov eax, dword ptr [ebp+0c]
'006cc733    33db                    xor ebx, ebx
'006cc735    66833801                cmp word ptr [eax], 01
'006cc739    895de8                  mov dword ptr [ebp-18], ebx
'006cc73c    895de4                  mov dword ptr [ebp-1c], ebx
'006cc73f    895de0                  mov dword ptr [ebp-20], ebx
'006cc742    895ddc                  mov dword ptr [ebp-24], ebx
'006cc745    895dd8                  mov dword ptr [ebp-28], ebx
'006cc748    895dd4                  mov dword ptr [ebp-2c], ebx
'006cc74b    0f855e010000            jne 6cc8af

If (arg_0 = 1) Then
'006cc751    391d24be7200            cmp dword ptr [0072be24], ebx
'006cc757    7510                    jne 6cc769
'006cc759    6824be7200              push 0072be24
'006cc75e    6870004300              push 00430070

' *** Reference to "__vbaNew2"
'006cc763    ff152c124000            call dword ptr [0040122c]
    Dim var_2 As New Global
'006cc769    8b3524be7200            mov esi, dword ptr [0072be24]
'006cc76f    8b0e                    mov ecx, dword ptr [esi]
'006cc771    8d55d8                  lea edx, dword ptr [ebp-28]
'006cc774    52                      push edx
'006cc775    56                      push esi
'006cc776    ff5118                  call dword ptr [ecx+18]
    Set var_45 = var_2.Screen
'006cc779    dbe2                    fnclex
'006cc77b    3bc3                    cmp eax, ebx
'006cc77d    7d13                    jge 6cc792
    
    If (    var_2.Screen < 0) Then

' *** Reference to "__vbaHresultCheckObj"
'006cc77f    8b3d80104000            mov edi, dword ptr [00401080]
'006cc785    6a18                    push 18
'006cc787    6860004300              push 00430060
'006cc78c    56                      push esi
'006cc78d    50                      push eax
'006cc78e    ffd7                    call edi
'006cc790    eb06                    jmp 6cc798
    
End If

' *** Reference to "__vbaHresultCheckObj"
'006cc792    8b3d80104000            mov edi, dword ptr [00401080]
'006cc798    8b75d8                  mov esi, dword ptr [ebp-28]
'006cc79b    8b1e                    mov ebx, dword ptr [esi]
'006cc79d    b90b000000              mov ecx, 0000000b

' *** Reference to "__vbaI2I4"
'006cc7a2    ff1550114000            call dword ptr [00401150]
'006cc7a8    50                      push eax
'006cc7a9    56                      push esi
'006cc7aa    ff537c                  call dword ptr [ebx+7c]
var_45.MousePointer = CInt(11)
'006cc7ad    dbe2                    fnclex
'006cc7af    85c0                    test ax, ax
'006cc7b1    7d0b                    jge 6cc7be
'006cc7b3    6a7c                    push 7c
'006cc7b5    6810be4300              push 0043be10
'006cc7ba    56                      push esi
'006cc7bb    50                      push eax
'006cc7bc    ffd7                    call edi
'006cc7be    8d4dd8                  lea ecx, dword ptr [ebp-28]

' *** Reference to "__vbaFreeObj"
'006cc7c1    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_45)
'006cc7c7    8b7508                  mov esi, dword ptr [ebp+08]
'006cc7ca    8b06                    mov eax, dword ptr [esi]
'006cc7cc    8d4dd4                  lea ecx, dword ptr [ebp-2c]
'006cc7cf    51                      push ecx
'006cc7d0    56                      push esi
'006cc7d1    ff5058                  call dword ptr [eax+58]
'006cc7d4    dbe2                    fnclex
'006cc7d6    85c0                    test ax, ax
'006cc7d8    7d0b                    jge 6cc7e5
'006cc7da    6a58                    push 58
'006cc7dc    6888094300              push 00430988
'006cc7e1    56                      push esi
'006cc7e2    50                      push eax
'006cc7e3    ffd7                    call edi

' *** Reference to "__vbaStrToAnsi"
'006cc7e5    8b359c124000            mov esi, dword ptr [0040129c]
'006cc7eb    6a00                    push 00
'006cc7ed    68cc134300              push 004313cc
'006cc7f2    8d55dc                  lea edx, dword ptr [ebp-24]
'006cc7f5    52                      push edx
'006cc7f6    ffd6                    call esi
var_39 = (vbNullChar)
'006cc7f8    50                      push eax
'006cc7f9    68cc134300              push 004313cc
'006cc7fe    8d45e0                  lea eax, dword ptr [ebp-20]
'006cc801    50                      push eax
'006cc802    ffd6                    call esi
var_3 = (vbNullChar)
'006cc804    50                      push eax
'006cc805    68801b4300              push 00431b80
'006cc80a    8d4de4                  lea ecx, dword ptr [ebp-1c]
'006cc80d    51                      push ecx
'006cc80e    ffd6                    call esi
var_40 = ("http://bonnarien.dyndns.org")
'006cc810    50                      push eax
'006cc811    68701b4300              push 00431b70
'006cc816    8d55e8                  lea edx, dword ptr [ebp-18]
'006cc819    52                      push edx
'006cc81a    ffd6                    call esi
var_41 = ("open")
'006cc81c    50                      push eax
'006cc81d    8b45d4                  mov eax, dword ptr [ebp-2c]
'006cc820    50                      push eax
'006cc821    e8523dd6ff              call 430578
var_42 = ShellExecuteA (0, var_41, var_40, var_3, var_39, 0)  '{Function}

' *** Reference to "__vbaSetSystemError"
'006cc826    ff157c104000            call dword ptr [0040107c]
'#SetAPISystemError()
'006cc82c    8d4ddc                  lea ecx, dword ptr [ebp-24]
'006cc82f    51                      push ecx
'006cc830    8d55e0                  lea edx, dword ptr [ebp-20]
'006cc833    52                      push edx
'006cc834    8d45e4                  lea eax, dword ptr [ebp-1c]
'006cc837    50                      push eax
'006cc838    8d4de8                  lea ecx, dword ptr [ebp-18]
'006cc83b    51                      push ecx
'006cc83c    6a04                    push 04

' *** Reference to "__vbaFreeStrList"
'006cc83e    ff155c124000            call dword ptr [0040125c]
'#Cleanup( var_41, var_40, var_3, var_39)
'006cc844    a124be7200              mov ax, word ptr [0072be24]
'006cc849    83c414                  add esp, 14
'006cc84c    85c0                    test ax, ax
'006cc84e    7510                    jne 6cc860
'006cc850    6824be7200              push 0072be24
'006cc855    6870004300              push 00430070

' *** Reference to "__vbaNew2"
'006cc85a    ff152c124000            call dword ptr [0040122c]
Set var_2 = New Global
'006cc860    8b3524be7200            mov esi, dword ptr [0072be24]
'006cc866    8b16                    mov edx, dword ptr [esi]
'006cc868    8d45d8                  lea eax, dword ptr [ebp-28]
'006cc86b    50                      push eax
'006cc86c    56                      push esi
'006cc86d    ff5218                  call dword ptr [edx+18]
Set var_45 = var_2.Screen
'006cc870    dbe2                    fnclex
'006cc872    85c0                    test ax, ax
'006cc874    7d0b                    jge 6cc881
'006cc876    6a18                    push 18
'006cc878    6860004300              push 00430060
'006cc87d    56                      push esi
'006cc87e    50                      push eax
'006cc87f    ffd7                    call edi
'006cc881    8b75d8                  mov esi, dword ptr [ebp-28]
'006cc884    8b1e                    mov ebx, dword ptr [esi]
'006cc886    33c9                    xor ecx, ecx

' *** Reference to "__vbaI2I4"
'006cc888    ff1550114000            call dword ptr [00401150]
'006cc88e    50                      push eax
'006cc88f    56                      push esi
'006cc890    ff537c                  call dword ptr [ebx+7c]
var_45.MousePointer = CInt(0)
'006cc893    dbe2                    fnclex
'006cc895    33db                    xor ebx, ebx
var_num2 = Empty
'006cc897    3bc3                    cmp eax, ebx
'006cc899    7d0b                    jge 6cc8a6
'006cc89b    6a7c                    push 7c
'006cc89d    6810be4300              push 0043be10
'006cc8a2    56                      push esi
'006cc8a3    50                      push eax
'006cc8a4    ffd7                    call edi
'006cc8a6    8d4dd8                  lea ecx, dword ptr [ebp-28]

' *** Reference to "__vbaFreeObj"
'006cc8a9    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_45)
'ERROR: Two many next close:
End If
'006cc8af    895dfc                  mov dword ptr [ebp-04], ebx
'006cc8b2    68dfc86c00              push 006cc8df
'006cc8b7    eb25                    jmp 6cc8de
'006cc8b9    8d4ddc                  lea ecx, dword ptr [ebp-24]
'006cc8bc    51                      push ecx
'006cc8bd    8d55e0                  lea edx, dword ptr [ebp-20]
'006cc8c0    52                      push edx
'006cc8c1    8d45e4                  lea eax, dword ptr [ebp-1c]
'006cc8c4    50                      push eax
'006cc8c5    8d4de8                  lea ecx, dword ptr [ebp-18]
'006cc8c8    51                      push ecx
'006cc8c9    6a04                    push 04

' *** Reference to "__vbaFreeStrList"
'006cc8cb    ff155c124000            call dword ptr [0040125c]
'#Cleanup( var_41, var_40, var_3, var_39)
'006cc8d1    83c414                  add esp, 14
'006cc8d4    8d4dd8                  lea ecx, dword ptr [ebp-28]

' *** Reference to "__vbaFreeObj"
'006cc8d7    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_45)
'006cc8dd    c3                      ret
'006cc8de    c3                      ret
'006cc8df    8b4508                  mov eax, dword ptr [ebp+08]
'006cc8e2    8b10                    mov edx, dword ptr [eax]
'006cc8e4    50                      push eax
'006cc8e5    ff5208                  call dword ptr [edx+08]
'006cc8e8    8b45fc                  mov eax, dword ptr [ebp-04]
'006cc8eb    8b4dec                  mov ecx, dword ptr [ebp-14]
'006cc8ee    5f                      pop edi
'006cc8ef    5e                      pop esi
    'Reference to '__except_list'
'006cc8f0    64890d00000000          mov dword ptr fs:[00000000], ecx
'006cc8f7    5b                      pop ebx
'006cc8f8    8be5                    mov esp, ebp
'006cc8fa    5d                      pop ebp
'006cc8fb    c20800                  ret 0008


End Sub


'Event for btnOk
Private Sub btnOk_Click()
'006cc280    55                      push ebp
'006cc281    8bec                    mov ebp, esp
'006cc283    83ec0c                  sub esp, 0c

' *** Reference to "__vbaExceptHandler"
'006cc286    6866724000              push 00407266
'006cc28b    64a100000000            mov ax, word ptr fs:[00000000]
'006cc291    50                      push eax
    'Reference to '__except_list'
'006cc292    64892500000000          mov dword ptr fs:[00000000], esp
'006cc299    83ec18                  sub esp, 18
'006cc29c    53                      push ebx
'006cc29d    56                      push esi
'006cc29e    57                      push edi
'006cc29f    8965f4                  mov dword ptr [ebp-0c], esp
'006cc2a2    c745f8a8674000          mov dword ptr [ebp-08], 004067a8
'006cc2a9    8b7d08                  mov edi, dword ptr [ebp+08]
'006cc2ac    8bc7                    mov eax, edi
'006cc2ae    83e001                  and eax, 01
'006cc2b1    8945fc                  mov dword ptr [ebp-04], eax
'006cc2b4    83e7fe                  and edi, fffffffe
'006cc2b7    8b0f                    mov ecx, dword ptr [edi]
'006cc2b9    57                      push edi
'006cc2ba    897d08                  mov dword ptr [ebp+08], edi
'006cc2bd    ff5104                  call dword ptr [ecx+04]
'006cc2c0    a124be7200              mov ax, word ptr [0072be24]
'006cc2c5    33db                    xor ebx, ebx
'006cc2c7    3bc3                    cmp eax, ebx
'006cc2c9    895de8                  mov dword ptr [ebp-18], ebx
'006cc2cc    7510                    jne 6cc2de

If (0 = 0) Then
'006cc2ce    6824be7200              push 0072be24
'006cc2d3    6870004300              push 00430070

' *** Reference to "__vbaNew2"
'006cc2d8    ff152c124000            call dword ptr [0040122c]
    Dim var_2 As New Global
End If
'006cc2de    8b3524be7200            mov esi, dword ptr [0072be24]
'006cc2e4    8b16                    mov edx, dword ptr [esi]
'006cc2e6    57                      push edi
'006cc2e7    8d45e8                  lea eax, dword ptr [ebp-18]
'006cc2ea    50                      push eax
'006cc2eb    8955d4                  mov dword ptr [ebp-2c], edx

' *** Reference to "__vbaObjSetAddref"
'006cc2ee    ff15c8104000            call dword ptr [004010c8]
Set var_41 = Me
'006cc2f4    8b4dd4                  mov ecx, dword ptr [ebp-2c]
'006cc2f7    50                      push eax
'006cc2f8    56                      push esi
'006cc2f9    ff5110                  call dword ptr [ecx+10]
Call var_2.Unload(var_41)
'006cc2fc    dbe2                    fnclex
'006cc2fe    3bc3                    cmp eax, ebx
'006cc300    7d0f                    jge 6cc311
'006cc302    6a10                    push 10
'006cc304    6860004300              push 00430060
'006cc309    56                      push esi
'006cc30a    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006cc30b    ff1580104000            call dword ptr [00401080]
'006cc311    8d4de8                  lea ecx, dword ptr [ebp-18]

' *** Reference to "__vbaFreeObj"
'006cc314    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_41)
'006cc31a    895dfc                  mov dword ptr [ebp-04], ebx
'006cc31d    682fc36c00              push 006cc32f
'006cc322    eb0a                    jmp 6cc32e
'006cc324    8d4de8                  lea ecx, dword ptr [ebp-18]

' *** Reference to "__vbaFreeObj"
'006cc327    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_41)
'006cc32d    c3                      ret
'006cc32e    c3                      ret
'006cc32f    8b4508                  mov eax, dword ptr [ebp+08]
'006cc332    8b10                    mov edx, dword ptr [eax]
'006cc334    50                      push eax
'006cc335    ff5208                  call dword ptr [edx+08]
'006cc338    8b45fc                  mov eax, dword ptr [ebp-04]
'006cc33b    8b4dec                  mov ecx, dword ptr [ebp-14]
'006cc33e    5f                      pop edi
'006cc33f    5e                      pop esi
    'Reference to '__except_list'
'006cc340    64890d00000000          mov dword ptr fs:[00000000], ecx
'006cc347    5b                      pop ebx
'006cc348    8be5                    mov esp, ebp
'006cc34a    5d                      pop ebp
'006cc34b    c20400                  ret 0004


End Sub


'Event for Form
Private Sub Form_Load()
'006cc350    55                      push ebp
'006cc351    8bec                    mov ebp, esp
'006cc353    83ec0c                  sub esp, 0c

' *** Reference to "__vbaExceptHandler"
'006cc356    6866724000              push 00407266
'006cc35b    64a100000000            mov ax, word ptr fs:[00000000]
'006cc361    50                      push eax
    'Reference to '__except_list'
'006cc362    64892500000000          mov dword ptr fs:[00000000], esp
'006cc369    81ec90000000            sub esp, 00000090
'006cc36f    53                      push ebx
'006cc370    56                      push esi
'006cc371    57                      push edi
'006cc372    8965f4                  mov dword ptr [ebp-0c], esp
'006cc375    c745f8b8674000          mov dword ptr [ebp-08], 004067b8
'006cc37c    8b7d08                  mov edi, dword ptr [ebp+08]
'006cc37f    8bc7                    mov eax, edi
'006cc381    83e001                  and eax, 01
'006cc384    8945fc                  mov dword ptr [ebp-04], eax
'006cc387    83e7fe                  and edi, fffffffe
'006cc38a    8b0f                    mov ecx, dword ptr [edi]
'006cc38c    57                      push edi
'006cc38d    897d08                  mov dword ptr [ebp+08], edi
'006cc390    ff5104                  call dword ptr [ecx+04]
'006cc393    8b17                    mov edx, dword ptr [edi]
'006cc395    33f6                    xor esi, esi
'006cc397    57                      push edi
'006cc398    8975e8                  mov dword ptr [ebp-18], esi
'006cc39b    8975e4                  mov dword ptr [ebp-1c], esi
'006cc39e    8975e0                  mov dword ptr [ebp-20], esi
'006cc3a1    8975dc                  mov dword ptr [ebp-24], esi
'006cc3a4    8975d8                  mov dword ptr [ebp-28], esi
'006cc3a7    8975d4                  mov dword ptr [ebp-2c], esi
'006cc3aa    8975d0                  mov dword ptr [ebp-30], esi
'006cc3ad    8975cc                  mov dword ptr [ebp-34], esi
'006cc3b0    8975c8                  mov dword ptr [ebp-38], esi
'006cc3b3    8975c4                  mov dword ptr [ebp-3c], esi
'006cc3b6    8975c0                  mov dword ptr [ebp-40], esi
'006cc3b9    8975bc                  mov dword ptr [ebp-44], esi
'006cc3bc    8975b8                  mov dword ptr [ebp-48], esi
'006cc3bf    8975b4                  mov dword ptr [ebp-4c], esi
'006cc3c2    8975b0                  mov dword ptr [ebp-50], esi
'006cc3c5    8975ac                  mov dword ptr [ebp-54], esi
'006cc3c8    8975a8                  mov dword ptr [ebp-58], esi
'006cc3cb    8975a4                  mov dword ptr [ebp-5c], esi
'006cc3ce    8975a0                  mov dword ptr [ebp-60], esi
'006cc3d1    89759c                  mov dword ptr [ebp-64], esi
'006cc3d4    ff9200030000            call dword ptr [edx+00000300]
'006cc3da    50                      push eax
'006cc3db    8d45a8                  lea eax, dword ptr [ebp-58]
'006cc3de    50                      push eax

' *** Reference to "__vbaObjSet"
'006cc3df    ff15b4104000            call dword ptr [004010b4]
Set var_86 = Me
'006cc3e5    898568ffffff            mov dword ptr [ebp+ffffff68], eax
'006cc3eb    393524be7200            cmp dword ptr [0072be24], esi
'006cc3f1    7510                    jne 6cc403
'006cc3f3    6824be7200              push 0072be24
'006cc3f8    6870004300              push 00430070

' *** Reference to "__vbaNew2"
'006cc3fd    ff152c124000            call dword ptr [0040122c]
Dim var_2 As New Global
'006cc403    8b3d24be7200            mov edi, dword ptr [0072be24]
'006cc409    8b0f                    mov ecx, dword ptr [edi]
'006cc40b    8d55b4                  lea edx, dword ptr [ebp-4c]
'006cc40e    52                      push edx
'006cc40f    57                      push edi
'006cc410    ff5114                  call dword ptr [ecx+14]
Set var_62 = var_2.App
'006cc413    dbe2                    fnclex
'006cc415    3bc6                    cmp eax, esi
'006cc417    7d13                    jge 6cc42c

If (var_2.App < 0) Then

' *** Reference to "__vbaHresultCheckObj"
'006cc419    8b1d80104000            mov ebx, dword ptr [00401080]
'006cc41f    6a14                    push 14
'006cc421    6860004300              push 00430060
'006cc426    57                      push edi
'006cc427    50                      push eax
'006cc428    ffd3                    call ebx
'006cc42a    eb06                    jmp 6cc432
    
End If

' *** Reference to "__vbaHresultCheckObj"
'006cc42c    8b1d80104000            mov ebx, dword ptr [00401080]
'006cc432    8b45b4                  mov eax, dword ptr [ebp-4c]
'006cc435    8b08                    mov ecx, dword ptr [eax]
'006cc437    8d55a4                  lea edx, dword ptr [ebp-5c]
'006cc43a    52                      push edx
'006cc43b    50                      push eax
'006cc43c    8bf8                    mov edi, eax
'006cc43e    ff91b8000000            call dword ptr [ecx+000000b8]
var_83 = var_62.Major
'006cc444    dbe2                    fnclex
'006cc446    3bc6                    cmp eax, esi
'006cc448    7d0e                    jge 6cc458
'006cc44a    68b8000000              push 000000b8
'006cc44f    682c1c4300              push 00431c2c
'006cc454    57                      push edi
'006cc455    50                      push eax
'006cc456    ffd3                    call ebx
'006cc458    393524be7200            cmp dword ptr [0072be24], esi
'006cc45e    7510                    jne 6cc470

If (var_2 = 0) Then
'006cc460    6824be7200              push 0072be24
'006cc465    6870004300              push 00430070

' *** Reference to "__vbaNew2"
'006cc46a    ff152c124000            call dword ptr [0040122c]
    Set var_2 = New Global
End If
'006cc470    8b3d24be7200            mov edi, dword ptr [0072be24]
'006cc476    8b07                    mov eax, dword ptr [edi]
'006cc478    8d4db0                  lea ecx, dword ptr [ebp-50]
'006cc47b    51                      push ecx
'006cc47c    57                      push edi
'006cc47d    ff5014                  call dword ptr [eax+14]
Set var_6 = var_2.App
'006cc480    dbe2                    fnclex
'006cc482    3bc6                    cmp eax, esi
'006cc484    7d0b                    jge 6cc491

If (var_2.App < 0) Then
'006cc486    6a14                    push 14
'006cc488    6860004300              push 00430060
'006cc48d    57                      push edi
'006cc48e    50                      push eax
'006cc48f    ffd3                    call ebx
    
End If
'006cc491    8b45b0                  mov eax, dword ptr [ebp-50]
'006cc494    8b10                    mov edx, dword ptr [eax]
'006cc496    8d4da0                  lea ecx, dword ptr [ebp-60]
'006cc499    51                      push ecx
'006cc49a    50                      push eax
'006cc49b    8bf8                    mov edi, eax
'006cc49d    ff92c0000000            call dword ptr [edx+000000c0]
var_7 = var_6.Minor
'006cc4a3    dbe2                    fnclex
'006cc4a5    3bc6                    cmp eax, esi
'006cc4a7    7d0e                    jge 6cc4b7
'006cc4a9    68c0000000              push 000000c0
'006cc4ae    682c1c4300              push 00431c2c
'006cc4b3    57                      push edi
'006cc4b4    50                      push eax
'006cc4b5    ffd3                    call ebx
'006cc4b7    393524be7200            cmp dword ptr [0072be24], esi
'006cc4bd    7510                    jne 6cc4cf

If (var_2 = 0) Then
'006cc4bf    6824be7200              push 0072be24
'006cc4c4    6870004300              push 00430070

' *** Reference to "__vbaNew2"
'006cc4c9    ff152c124000            call dword ptr [0040122c]
    Set var_2 = New Global
End If
'006cc4cf    8b3d24be7200            mov edi, dword ptr [0072be24]
'006cc4d5    8b17                    mov edx, dword ptr [edi]
'006cc4d7    8d45ac                  lea eax, dword ptr [ebp-54]
'006cc4da    50                      push eax
'006cc4db    57                      push edi
'006cc4dc    ff5214                  call dword ptr [edx+14]
Set var_50 = var_2.App
'006cc4df    dbe2                    fnclex
'006cc4e1    3bc6                    cmp eax, esi
'006cc4e3    7d0b                    jge 6cc4f0

If (var_2.App < 0) Then
'006cc4e5    6a14                    push 14
'006cc4e7    6860004300              push 00430060
'006cc4ec    57                      push edi
'006cc4ed    50                      push eax
'006cc4ee    ffd3                    call ebx
    
End If
'006cc4f0    8b45ac                  mov eax, dword ptr [ebp-54]
'006cc4f3    8b08                    mov ecx, dword ptr [eax]
'006cc4f5    8d559c                  lea edx, dword ptr [ebp-64]
'006cc4f8    52                      push edx
'006cc4f9    50                      push eax
'006cc4fa    8bf8                    mov edi, eax
'006cc4fc    ff91c8000000            call dword ptr [ecx+000000c8]
var_51 = var_50.Revision
'006cc502    dbe2                    fnclex
'006cc504    3bc6                    cmp eax, esi
'006cc506    7d0e                    jge 6cc516
'006cc508    68c8000000              push 000000c8
'006cc50d    682c1c4300              push 00431c2c
'006cc512    57                      push edi
'006cc513    50                      push eax
'006cc514    ffd3                    call ebx
'006cc516    8b8568ffffff            mov eax, dword ptr [ebp+ffffff68]

' *** Reference to "__vbaStrCat"
'006cc51c    8b3d70104000            mov edi, dword ptr [00401070]
'006cc522    8b18                    mov ebx, dword ptr [eax]
'006cc524    68c4394500              push 004539c4
'006cc529    6844ed4300              push 0043ed44
'006cc52e    ffd7                    call edi
var_493 = ("Version du logiciel :        ") & (vbTab)

' *** Reference to "__vbaStrMove"
'006cc530    8b35d0124000            mov esi, dword ptr [004012d0]
'006cc536    8bd0                    mov edx, eax
'006cc538    8d4de8                  lea ecx, dword ptr [ebp-18]
'006cc53b    ffd6                    call esi
'006cc53d    8b4da4                  mov ecx, dword ptr [ebp-5c]
'006cc540    50                      push eax
'006cc541    51                      push ecx

' *** Reference to "__vbaStrI2"
'006cc542    ff1510104000            call dword ptr [00401010]
'006cc548    8bd0                    mov edx, eax
'006cc54a    8d4de4                  lea ecx, dword ptr [ebp-1c]
'006cc54d    ffd6                    call esi
'006cc54f    50                      push eax
'006cc550    ffd7                    call edi
var_127 = (var_493) & (CStr(var_83))
'006cc552    8bd0                    mov edx, eax
'006cc554    8d4de0                  lea ecx, dword ptr [ebp-20]
'006cc557    ffd6                    call esi
'006cc559    50                      push eax
'006cc55a    6870084300              push 00430870
'006cc55f    ffd7                    call edi
var_15 = (var_127) & (vbCrLf)
'006cc561    8bd0                    mov edx, eax
'006cc563    8d4ddc                  lea ecx, dword ptr [ebp-24]
'006cc566    ffd6                    call esi
'006cc568    50                      push eax
'006cc569    68043a4500              push 00453a04
'006cc56e    ffd7                    call edi
var_16 = (var_15) & ("Paramètres roles.mdb : ")
'006cc570    8bd0                    mov edx, eax
'006cc572    8d4dd8                  lea ecx, dword ptr [ebp-28]
'006cc575    ffd6                    call esi
'006cc577    50                      push eax
'006cc578    6844ed4300              push 0043ed44
'006cc57d    ffd7                    call edi
var_128 = (var_16) & (vbTab)
'006cc57f    8bd0                    mov edx, eax
'006cc581    8d4dd4                  lea ecx, dword ptr [ebp-2c]
'006cc584    ffd6                    call esi
'006cc586    8b55a0                  mov edx, dword ptr [ebp-60]
'006cc589    50                      push eax
'006cc58a    52                      push edx

' *** Reference to "__vbaStrI2"
'006cc58b    ff1510104000            call dword ptr [00401010]
'006cc591    8bd0                    mov edx, eax
'006cc593    8d4dd0                  lea ecx, dword ptr [ebp-30]
'006cc596    ffd6                    call esi
'006cc598    50                      push eax
'006cc599    ffd7                    call edi
var_129 = (var_128) & (CStr(var_7))
'006cc59b    8bd0                    mov edx, eax
'006cc59d    8d4dcc                  lea ecx, dword ptr [ebp-34]
'006cc5a0    ffd6                    call esi
'006cc5a2    50                      push eax
'006cc5a3    6870084300              push 00430870
'006cc5a8    ffd7                    call edi
var_53 = (var_129) & (vbCrLf)
'006cc5aa    8bd0                    mov edx, eax
'006cc5ac    8d4dc8                  lea ecx, dword ptr [ebp-38]
'006cc5af    ffd6                    call esi
'006cc5b1    50                      push eax
'006cc5b2    68383a4500              push 00453a38
'006cc5b7    ffd7                    call edi
var_285 = (var_53) & ("Nombre de compilations :     ")
'006cc5b9    8bd0                    mov edx, eax
'006cc5bb    8d4dc4                  lea ecx, dword ptr [ebp-3c]
'006cc5be    ffd6                    call esi
'006cc5c0    50                      push eax
'006cc5c1    6844ed4300              push 0043ed44
'006cc5c6    ffd7                    call edi
var_54 = (var_285) & (vbTab)
'006cc5c8    8bd0                    mov edx, eax
'006cc5ca    8d4dc0                  lea ecx, dword ptr [ebp-40]
'006cc5cd    ffd6                    call esi
'006cc5cf    50                      push eax
'006cc5d0    8b459c                  mov eax, dword ptr [ebp-64]
'006cc5d3    50                      push eax

' *** Reference to "__vbaStrI2"
'006cc5d4    ff1510104000            call dword ptr [00401010]
'006cc5da    8bd0                    mov edx, eax
'006cc5dc    8d4dbc                  lea ecx, dword ptr [ebp-44]
'006cc5df    ffd6                    call esi
'006cc5e1    50                      push eax
'006cc5e2    ffd7                    call edi
var_139 = (var_54) & (CStr(var_51))
'006cc5e4    8bd0                    mov edx, eax
'006cc5e6    8d4db8                  lea ecx, dword ptr [ebp-48]
'006cc5e9    ffd6                    call esi
'006cc5eb    8bb568ffffff            mov esi, dword ptr [ebp+ffffff68]
'006cc5f1    50                      push eax
'006cc5f2    56                      push esi
'006cc5f3    ff93a4000000            call dword ptr [ebx+000000a4]
'006cc5f9    dbe2                    fnclex
'006cc5fb    85c0                    test ax, ax
'006cc5fd    7d12                    jge 6cc611
'006cc5ff    68a4000000              push 000000a4
'006cc604    68d00d4300              push 00430dd0
'006cc609    56                      push esi
'006cc60a    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006cc60b    ff1580104000            call dword ptr [00401080]
'006cc611    8d4db8                  lea ecx, dword ptr [ebp-48]
'006cc614    51                      push ecx
'006cc615    8d55bc                  lea edx, dword ptr [ebp-44]
'006cc618    52                      push edx
'006cc619    8d45c0                  lea eax, dword ptr [ebp-40]
'006cc61c    50                      push eax
'006cc61d    8d4dc4                  lea ecx, dword ptr [ebp-3c]
'006cc620    51                      push ecx
'006cc621    8d55c8                  lea edx, dword ptr [ebp-38]
'006cc624    52                      push edx
'006cc625    8d45cc                  lea eax, dword ptr [ebp-34]
'006cc628    50                      push eax
'006cc629    8d4dd0                  lea ecx, dword ptr [ebp-30]
'006cc62c    51                      push ecx
'006cc62d    8d55d4                  lea edx, dword ptr [ebp-2c]
'006cc630    52                      push edx
'006cc631    8d45d8                  lea eax, dword ptr [ebp-28]
'006cc634    50                      push eax
'006cc635    8d4ddc                  lea ecx, dword ptr [ebp-24]
'006cc638    51                      push ecx
'006cc639    8d55e0                  lea edx, dword ptr [ebp-20]
'006cc63c    52                      push edx
'006cc63d    8d45e4                  lea eax, dword ptr [ebp-1c]
'006cc640    50                      push eax
'006cc641    8d4de8                  lea ecx, dword ptr [ebp-18]
'006cc644    51                      push ecx
'006cc645    6a0d                    push 0d

' *** Reference to "__vbaFreeStrList"
'006cc647    ff155c124000            call dword ptr [0040125c]
'#Cleanup( -4520, -4524, -4528, -4532, -4536, -4540, -4544, -4548, -4552, -4556, -4560, -4564, -4568)
'006cc64d    8d55a8                  lea edx, dword ptr [ebp-58]
'006cc650    52                      push edx
'006cc651    8d45ac                  lea eax, dword ptr [ebp-54]
'006cc654    50                      push eax
'006cc655    8d4db0                  lea ecx, dword ptr [ebp-50]
'006cc658    51                      push ecx
'006cc659    8d55b4                  lea edx, dword ptr [ebp-4c]
'006cc65c    52                      push edx
'006cc65d    6a04                    push 04

' *** Reference to "__vbaFreeObjList"
'006cc65f    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_62, var_6, var_50, var_86)
'006cc665    83c44c                  add esp, 4c
'006cc668    c745fc00000000          mov dword ptr [ebp-04], 00000000
'006cc66f    68cfc66c00              push 006cc6cf
'006cc674    eb58                    jmp 6cc6ce
'006cc676    8d45b8                  lea eax, dword ptr [ebp-48]
'006cc679    50                      push eax
'006cc67a    8d4dbc                  lea ecx, dword ptr [ebp-44]
'006cc67d    51                      push ecx
'006cc67e    8d55c0                  lea edx, dword ptr [ebp-40]
'006cc681    52                      push edx
'006cc682    8d45c4                  lea eax, dword ptr [ebp-3c]
'006cc685    50                      push eax
'006cc686    8d4dc8                  lea ecx, dword ptr [ebp-38]
'006cc689    51                      push ecx
'006cc68a    8d55cc                  lea edx, dword ptr [ebp-34]
'006cc68d    52                      push edx
'006cc68e    8d45d0                  lea eax, dword ptr [ebp-30]
'006cc691    50                      push eax
'006cc692    8d4dd4                  lea ecx, dword ptr [ebp-2c]
'006cc695    51                      push ecx
'006cc696    8d55d8                  lea edx, dword ptr [ebp-28]
'006cc699    52                      push edx
'006cc69a    8d45dc                  lea eax, dword ptr [ebp-24]
'006cc69d    50                      push eax
'006cc69e    8d4de0                  lea ecx, dword ptr [ebp-20]
'006cc6a1    51                      push ecx
'006cc6a2    8d55e4                  lea edx, dword ptr [ebp-1c]
'006cc6a5    52                      push edx
'006cc6a6    8d45e8                  lea eax, dword ptr [ebp-18]
'006cc6a9    50                      push eax
'006cc6aa    6a0d                    push 0d

' *** Reference to "__vbaFreeStrList"
'006cc6ac    ff155c124000            call dword ptr [0040125c]
'#Cleanup( -4520, -4524, -4528, -4532, -4536, -4540, -4544, -4548, -4552, -4556, -4560, -4564, -4568)
'006cc6b2    8d4da8                  lea ecx, dword ptr [ebp-58]
'006cc6b5    51                      push ecx
'006cc6b6    8d55ac                  lea edx, dword ptr [ebp-54]
'006cc6b9    52                      push edx
'006cc6ba    8d45b0                  lea eax, dword ptr [ebp-50]
'006cc6bd    50                      push eax
'006cc6be    8d4db4                  lea ecx, dword ptr [ebp-4c]
'006cc6c1    51                      push ecx
'006cc6c2    6a04                    push 04

' *** Reference to "__vbaFreeObjList"
'006cc6c4    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_62, var_6, var_50, var_86)
'006cc6ca    83c44c                  add esp, 4c
'006cc6cd    c3                      ret
'006cc6ce    c3                      ret
'006cc6cf    8b4508                  mov eax, dword ptr [ebp+08]
'006cc6d2    8b10                    mov edx, dword ptr [eax]
'006cc6d4    50                      push eax
'006cc6d5    ff5208                  call dword ptr [edx+08]
'006cc6d8    8b45fc                  mov eax, dword ptr [ebp-04]
'006cc6db    8b4dec                  mov ecx, dword ptr [ebp-14]
'006cc6de    5f                      pop edi
'006cc6df    5e                      pop esi
    'Reference to '__except_list'
'006cc6e0    64890d00000000          mov dword ptr fs:[00000000], ecx
'006cc6e7    5b                      pop ebx
'006cc6e8    8be5                    mov esp, ebp
'006cc6ea    5d                      pop ebp
'006cc6eb    c20400                  ret 0004


End Sub


