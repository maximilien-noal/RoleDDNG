VERSION 5.00

Begin VB.Form frmDescriptifDon
    Caption = "Description du don"
    ScaleMode = 1
    AutoRedraw = 0              'False
    FontTransparent = -1              'True
    LinkTopic = "Form1"
    ControlBox = 0              'False
    ClipControls = 0              'False
    ClientLeft   = 60
    ClientTop    = 420
    ClientWidth  = 8415
    ClientHeight = 8490
    BeginProperty Font
        Name          = "Times New Roman"
        Size          = 9,75
        Charset       = 0
        Weight        = 400
        Underline     = 0              'False
        Italic        = 0              'False
        Strikethrough = 0              'False
    EndProperty
    Begin VB.CommandButton btnSuivant
        Caption = "Suivant"
        Left   = 4280
        Top    = 8500
        Width  = 1275
        Height = 405
        TabIndex = 20
    End
    Begin VB.CommandButton btnPrecedent
        Caption = "Précédent"
        Left   = 2880
        Top    = 8500
        Width  = 1275
        Height = 405
        TabIndex = 19
    End
    Begin VB.ComboBox CmbNomDon
        Left   = 120
        Top    = 8500
        Width  = 2655
        Height = 345
        Text = "Combo1ÿ"
        TabIndex = 18
    End
    Begin VB.Frame Frame1
        Left   = 45
        Top    = 840
        Width  = 8295
        Height = 7600
        TabIndex = 3
        BeginProperty Font
            Name          = "Times New Roman"
            Size          = 9
            Charset       = 0
            Weight        = 400
            Underline     = 0              'False
            Italic        = 0              'False
            Strikethrough = 0              'False
        EndProperty
        Begin VB.TextBox fldstrResume
            BackColor = 12632256
            Left   = 120
            Top    = 500
            Width  = 7995
            Height = 660
            TabIndex = 13
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
        Begin VB.TextBox fldstrSpecial
            BackColor = 12632256
            Left   = 120
            Top    = 6800
            Width  = 7995
            Height = 650
            TabIndex = 11
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
        Begin VB.TextBox fldstrNormal
            BackColor = 12632256
            Left   = 120
            Top    = 5800
            Width  = 7995
            Height = 650
            TabIndex = 9
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
        Begin VB.TextBox fldstrCondition
            BackColor = 12632256
            Left   = 120
            Top    = 1500
            Width  = 7995
            Height = 660
            TabIndex = 6
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
        Begin VB.TextBox fldstrExplication
            BackColor = 12632256
            Left   = 90
            Top    = 2500
            Width  = 7995
            Height = 2970
            TabIndex = 4
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
        Begin VB.Label Label1
            Caption = "Résumé"
            Index = 4
            Left   = 120
            Top    = 200
            Width  = 975
            Height = 240
            TabIndex = 12
        End
        Begin VB.Label Label1
            Caption = "Spécial"
            Index = 3
            Left   = 120
            Top    = 6500
            Width  = 735
            Height = 240
            TabIndex = 10
        End
        Begin VB.Label Label1
            Caption = "Normal"
            Index = 2
            Left   = 120
            Top    = 5500
            Width  = 735
            Height = 240
            TabIndex = 8
        End
        Begin VB.Label Label1
            Caption = "Condition"
            Index = 1
            Left   = 120
            Top    = 1200
            Width  = 975
            Height = 240
            TabIndex = 7
        End
        Begin VB.Label Label1
            Caption = "Avantage"
            Index = 0
            Left   = 120
            Top    = 2200
            Width  = 855
            Height = 240
            TabIndex = 5
        End
    End
    Begin VB.CommandButton btnEnregistrer
        Caption = "&Enregistrer"
        Left   = 5680
        Top    = 8500
        Width  = 1275
        Height = 405
        TabIndex = 2
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
    Begin VB.CommandButton btnFermer
        Caption = "&Fermer"
        Left   = 7080
        Top    = 8500
        Width  = 1275
        Height = 405
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
    Begin VB.Label LabType
        BackColor = -2147483648
        Left   = 3240
        Top    = 480
        Width  = 1695
        Height = 330
        TabIndex = 17
        BorderStyle = 1
        Alignment = 2
        BeginProperty Font
            Name          = "Times New Roman"
            Size          = 9
            Charset       = 0
            Weight        = 700
            Underline     = 0              'False
            Italic        = 0              'False
            Strikethrough = 0              'False
        EndProperty
    End
    Begin VB.Label LabSource
        BackColor = -2147483648
        Left   = 5160
        Top    = 480
        Width  = 3015
        Height = 330
        TabIndex = 16
        BorderStyle = 1
        Alignment = 2
        BeginProperty Font
            Name          = "Times New Roman"
            Size          = 9
            Charset       = 0
            Weight        = 700
            Underline     = 0              'False
            Italic        = 0              'False
            Strikethrough = 0              'False
        EndProperty
    End
    Begin VB.Label LabCategorie
        BackColor = -2147483648
        Left   = 120
        Top    = 480
        Width  = 3015
        Height = 330
        TabIndex = 15
        BorderStyle = 1
        Alignment = 2
        BeginProperty Font
            Name          = "Times New Roman"
            Size          = 9
            Charset       = 0
            Weight        = 700
            Underline     = 0              'False
            Italic        = 0              'False
            Strikethrough = 0              'False
        EndProperty
    End
    Begin VB.Label LabGenre
        BackColor = -2147483648
        Left   = 5160
        Top    = 120
        Width  = 3015
        Height = 330
        TabIndex = 14
        BorderStyle = 1
        Alignment = 2
        BeginProperty Font
            Name          = "Times New Roman"
            Size          = 9
            Charset       = 0
            Weight        = 700
            Underline     = 0              'False
            Italic        = 0              'False
            Strikethrough = 0              'False
        EndProperty
    End
    Begin VB.Label LabDon
        BackColor = -2147483648
        Left   = 120
        Top    = 120
        Width  = 4815
        Height = 330
        TabIndex = 0
        BorderStyle = 1
        Alignment = 2
        BeginProperty Font
            Name          = "Times New Roman"
            Size          = 9
            Charset       = 0
            Weight        = 700
            Underline     = 0              'False
            Italic        = 0              'False
            Strikethrough = 0              'False
        EndProperty
    End
End
Private Function sub_6C23A0(arg_0 As Unknow, arg_1 As Long, arg_2 As Unknow, arg_3 As Unknow, arg_4 As Unknow, arg_5 As Unknow, arg_6 As Unknow, arg_7 As Unknow, arg_8 As Unknow, arg_9 As Unknow, arg_A As Unknow, arg_B As Unknow, arg_C As Unknow, arg_D As Unknow, arg_E As Unknow, arg_F As Unknow, arg_10 As Unknow, arg_11 As Unknow, arg_12 As Unknow, arg_13 As Unknow, arg_14 As Unknow, arg_15 As Unknow, arg_16 As Unknow, arg_17 As Unknow, arg_18 As Unknow, arg_19 As Unknow, arg_1A As Unknow, arg_1B As Unknow, arg_1C As Unknow, arg_1D As Unknow, arg_1E As Unknow, arg_1F As Unknow, arg_20 As Unknow, arg_21 As Unknow, arg_22 As Unknow, arg_23 As Unknow, arg_24 As Unknow, arg_25 As Unknow, arg_26 As Unknow, arg_27 As Unknow, arg_28 As Unknow, arg_29 As Unknow, arg_2A As Unknow, arg_2B As Unknow, arg_2C As Unknow, arg_2D As Unknow, arg_2E As Unknow, arg_2F As Unknow, arg_30 As Unknow, arg_31 As Unknow, arg_32 As Unknow, arg_33 As Unknow, arg_34 As Unknow, arg_35 As Unknow, arg_36 As Unknow, arg_37 As Unknow, arg_38 As Unknow, arg_39 As Unknow, arg_3A As Unknow, arg_3B As Unknow, arg_3C As Unknow)
'006c23a0    55                      push ebp
'006c23a1    8bec                    mov ebp, esp
'006c23a3    83ec0c                  sub esp, 0c

' *** Reference to "__vbaExceptHandler"
'006c23a6    6866724000              push 00407266
'006c23ab    64a100000000            mov ax, word ptr fs:[00000000]
'006c23b1    50                      push eax
    'Reference to '__except_list'
'006c23b2    64892500000000          mov dword ptr fs:[00000000], esp
'006c23b9    83ec64                  sub esp, 64
'006c23bc    53                      push ebx
'006c23bd    56                      push esi
'006c23be    57                      push edi
'006c23bf    8965f4                  mov dword ptr [ebp-0c], esp
'006c23c2    c745f800664000          mov dword ptr [ebp-08], 00406600
'006c23c9    8b4510                  mov eax, dword ptr [ebp+10]
'006c23cc    8b750c                  mov esi, dword ptr [ebp+0c]
'006c23cf    33db                    xor ebx, ebx
'006c23d1    53                      push ebx
'006c23d2    6aff                    push ffffffff
'006c23d4    6a01                    push 01
'006c23d6    6854374500              push 00453754
'006c23db    8918                    mov dword ptr [eax], ebx
'006c23dd    8b0e                    mov ecx, dword ptr [esi]
'006c23df    6838374500              push 00453738
'006c23e4    51                      push ecx
'006c23e5    895dd8                  mov dword ptr [ebp-28], ebx
'006c23e8    895dd4                  mov dword ptr [ebp-2c], ebx
'006c23eb    895dd0                  mov dword ptr [ebp-30], ebx
'006c23ee    895dcc                  mov dword ptr [ebp-34], ebx
'006c23f1    895dbc                  mov dword ptr [ebp-44], ebx
'006c23f4    895dac                  mov dword ptr [ebp-54], ebx

' *** Reference to "rtcReplace"
'006c23f7    ff15b4114000            call dword ptr [004011b4]
'006c23fd    8bd0                    mov edx, eax
'006c23ff    8bce                    mov ecx, esi

' *** Reference to "__vbaStrMove"
'006c2401    ff15d0124000            call dword ptr [004012d0]
'006c2407    8b16                    mov edx, dword ptr [esi]

' *** Reference to "__vbaInStr"
'006c2409    8b3d20124000            mov edi, dword ptr [00401220]
'006c240f    6a01                    push 01
'006c2411    52                      push edx
'006c2412    6870374500              push 00453770
'006c2417    53                      push ebx
'006c2418    ffd7                    call edi
'006c241a    85c0                    test ax, ax
'006c241c    0f84da050000            je 6c29fc
'006c2422    8b06                    mov eax, dword ptr [esi]
'006c2424    6a01                    push 01
'006c2426    50                      push eax
'006c2427    6870374500              push 00453770
'006c242c    53                      push ebx
'006c242d    8975b4                  mov dword ptr [ebp-4c], esi
'006c2430    c745ac08400000          mov dword ptr [ebp-54], 00004008
'006c2437    ffd7                    call edi
'006c2439    8b0e                    mov ecx, dword ptr [esi]
'006c243b    51                      push ecx
'006c243c    8bd8                    mov ebx, eax

' *** Reference to "__vbaLenBstr"
'006c243e    ff1534104000            call dword ptr [00401034]
'006c2444    2bc3                    sub eax, ebx
var_num1 = Len(Replace(arg_0, "Condition.", "Conditions.", 1, -1, 0)) - InStr(1, Replace(arg_0, "Condition.", "Conditions.", 1, -1, 0), "Avantage.", 0)
'006c2446    0f801b060000            jo 6c2a67
'006c244c    83e809                  sub eax, 09
var_num1 = var_num1 - 9
'006c244f    0f8012060000            jo 6c2a67
'006c2455    50                      push eax
'006c2456    8d55ac                  lea edx, dword ptr [ebp-54]
'006c2459    52                      push edx
'006c245a    8d45bc                  lea eax, dword ptr [ebp-44]
'006c245d    50                      push eax

' *** Reference to "rtcRightCharVar"
'006c245e    ff15d8124000            call dword ptr [004012d8]
'006c2464    8d4dbc                  lea ecx, dword ptr [ebp-44]
'006c2467    51                      push ecx

' *** Reference to "__vbaStrVarMove"
'006c2468    ff153c104000            call dword ptr [0040103c]
'006c246e    8bd0                    mov edx, eax
'006c2470    8d4dd4                  lea ecx, dword ptr [ebp-2c]

' *** Reference to "__vbaStrMove"
'006c2473    ff15d0124000            call dword ptr [004012d0]
'006c2479    8d4dbc                  lea ecx, dword ptr [ebp-44]

' *** Reference to "__vbaFreeVar"
'006c247c    ff1528104000            call dword ptr [00401028]
'#Cleanup(var_58)
'006c2482    8b16                    mov edx, dword ptr [esi]
'006c2484    6a01                    push 01
'006c2486    52                      push edx
'006c2487    6888374500              push 00453788
'006c248c    6a00                    push 00
'006c248e    ffd7                    call edi
'006c2490    85c0                    test ax, ax
'006c2492    6a01                    push 01
'006c2494    740f                    je 6c24a5
'006c2496    8b45d4                  mov eax, dword ptr [ebp-2c]
'006c2499    50                      push eax
'006c249a    6888374500              push 00453788
'006c249f    6a00                    push 00
'006c24a1    ffd7                    call edi
'006c24a3    eb2c                    jmp 6c24d1
'006c24a5    8b4dd4                  mov ecx, dword ptr [ebp-2c]
'006c24a8    51                      push ecx
'006c24a9    689c374500              push 0045379c
'006c24ae    6a00                    push 00
'006c24b0    ffd7                    call edi
'006c24b2    85c0                    test ax, ax
'006c24b4    7411                    je 6c24c7
'006c24b6    8b55d4                  mov edx, dword ptr [ebp-2c]
'006c24b9    6a01                    push 01
'006c24bb    52                      push edx
'006c24bc    689c374500              push 0045379c
'006c24c1    6a00                    push 00
'006c24c3    ffd7                    call edi
'006c24c5    eb0a                    jmp 6c24d1
'006c24c7    8b45d4                  mov eax, dword ptr [ebp-2c]
'006c24ca    50                      push eax

' *** Reference to "__vbaLenBstr"
'006c24cb    ff1534104000            call dword ptr [00401034]
'006c24d1    8bd8                    mov ebx, eax
'006c24d3    8b4508                  mov eax, dword ptr [ebp+08]
'006c24d6    8b08                    mov ecx, dword ptr [eax]
'006c24d8    50                      push eax
'006c24d9    ff911c030000            call dword ptr [ecx+0000031c]
'006c24df    50                      push eax
'006c24e0    8d55cc                  lea edx, dword ptr [ebp-34]
'006c24e3    52                      push edx

' *** Reference to "__vbaObjSet"
'006c24e4    ff15b4104000            call dword ptr [004010b4]
Set var_43 = Me
'006c24ea    83eb01                  sub ebx, 01
var_num2 = Len(Right(Replace(arg_0, "Condition.", "Conditions.", 1, -1, 0), var_num1)) - 1
'006c24ed    0f8074050000            jo 6c2a67
'006c24f3    53                      push ebx
'006c24f4    8d4dac                  lea ecx, dword ptr [ebp-54]
'006c24f7    8945a8                  mov dword ptr [ebp-58], eax
'006c24fa    51                      push ecx
'006c24fb    8d55bc                  lea edx, dword ptr [ebp-44]
'006c24fe    8d45d4                  lea eax, dword ptr [ebp-2c]
'006c2501    52                      push edx
'006c2502    8945b4                  mov dword ptr [ebp-4c], eax
'006c2505    c745ac08400000          mov dword ptr [ebp-54], 00004008

' *** Reference to "rtcLeftCharVar"
'006c250c    ff15bc124000            call dword ptr [004012bc]
'006c2512    8b45a8                  mov eax, dword ptr [ebp-58]
'006c2515    8b18                    mov ebx, dword ptr [eax]
'006c2517    8d4dbc                  lea ecx, dword ptr [ebp-44]
'006c251a    51                      push ecx
'006c251b    8d55d0                  lea edx, dword ptr [ebp-30]
'006c251e    52                      push edx

' *** Reference to "__vbaStrVarVal"
'006c251f    ff15fc114000            call dword ptr [004011fc]
'006c2525    50                      push eax
'006c2526    8bc3                    mov eax, ebx
'006c2528    8b5da8                  mov ebx, dword ptr [ebp-58]
'006c252b    53                      push ebx
'006c252c    ff90a4000000            call dword ptr [eax+000000a4]
'006c2532    dbe2                    fnclex
'006c2534    85c0                    test ax, ax
'006c2536    7d12                    jge 6c254a
'006c2538    68a4000000              push 000000a4
'006c253d    68d00d4300              push 00430dd0
'006c2542    53                      push ebx
'006c2543    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c2544    ff1580104000            call dword ptr [00401080]
'006c254a    8d4dd0                  lea ecx, dword ptr [ebp-30]

' *** Reference to "__vbaFreeStr"
'006c254d    ff150c134000            call dword ptr [0040130c]
'#Cleanup(var_4)
'006c2553    8d4dcc                  lea ecx, dword ptr [ebp-34]

' *** Reference to "__vbaFreeObj"
'006c2556    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_43)

' *** Reference to "__vbaFreeVar"
'006c255c    8b1d28104000            mov ebx, dword ptr [00401028]
'006c2562    8d4dbc                  lea ecx, dword ptr [ebp-44]
'006c2565    ffd3                    call ebx
'#Cleanup(var_58)
'006c2567    8b0e                    mov ecx, dword ptr [esi]
'006c2569    6a01                    push 01
'006c256b    51                      push ecx
'006c256c    6854374500              push 00453754
'006c2571    6a00                    push 00
'006c2573    ffd7                    call edi
'006c2575    85c0                    test ax, ax
'006c2577    0f848f010000            je 6c270c
'006c257d    8b4508                  mov eax, dword ptr [ebp+08]
'006c2580    8b10                    mov edx, dword ptr [eax]
'006c2582    50                      push eax
'006c2583    ff920c030000            call dword ptr [edx+0000030c]
'006c2589    50                      push eax
'006c258a    8d45cc                  lea eax, dword ptr [ebp-34]
'006c258d    50                      push eax

' *** Reference to "__vbaObjSet"
'006c258e    ff15b4104000            call dword ptr [004010b4]
Set var_43 = Me
'006c2594    8b0e                    mov ecx, dword ptr [esi]
'006c2596    6a01                    push 01
'006c2598    51                      push ecx
'006c2599    6854374500              push 00453754
'006c259e    8bd8                    mov ebx, eax
'006c25a0    6a00                    push 00
'006c25a2    895da8                  mov dword ptr [ebp-58], ebx
'006c25a5    8975b4                  mov dword ptr [ebp-4c], esi
'006c25a8    c745ac08400000          mov dword ptr [ebp-54], 00004008
'006c25af    ffd7                    call edi
'006c25b1    83e801                  sub eax, 01
var_num1 = InStr(1, Replace(arg_0, "Condition.", "Conditions.", 1, -1, 0), "Conditions.", 0) - 1
'006c25b4    0f80ad040000            jo 6c2a67
'006c25ba    50                      push eax
'006c25bb    8d55ac                  lea edx, dword ptr [ebp-54]
'006c25be    52                      push edx
'006c25bf    8d45bc                  lea eax, dword ptr [ebp-44]
'006c25c2    50                      push eax

' *** Reference to "rtcLeftCharVar"
'006c25c3    ff15bc124000            call dword ptr [004012bc]
'006c25c9    8b1b                    mov ebx, dword ptr [ebx]
'006c25cb    8d4dbc                  lea ecx, dword ptr [ebp-44]
'006c25ce    51                      push ecx
'006c25cf    8d55d0                  lea edx, dword ptr [ebp-30]
'006c25d2    52                      push edx

' *** Reference to "__vbaStrVarVal"
'006c25d3    ff15fc114000            call dword ptr [004011fc]
'006c25d9    50                      push eax
'006c25da    8bc3                    mov eax, ebx
'006c25dc    8b5da8                  mov ebx, dword ptr [ebp-58]
'006c25df    53                      push ebx
'006c25e0    ff90a4000000            call dword ptr [eax+000000a4]
'006c25e6    dbe2                    fnclex
'006c25e8    85c0                    test ax, ax
'006c25ea    7d12                    jge 6c25fe
'006c25ec    68a4000000              push 000000a4
'006c25f1    68d00d4300              push 00430dd0
'006c25f6    53                      push ebx
'006c25f7    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c25f8    ff1580104000            call dword ptr [00401080]
'006c25fe    8d4dd0                  lea ecx, dword ptr [ebp-30]

' *** Reference to "__vbaFreeStr"
'006c2601    ff150c134000            call dword ptr [0040130c]
'#Cleanup(var_4)
'006c2607    8d4dcc                  lea ecx, dword ptr [ebp-34]

' *** Reference to "__vbaFreeObj"
'006c260a    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_43)

' *** Reference to "__vbaFreeVar"
'006c2610    8b1d28104000            mov ebx, dword ptr [00401028]
'006c2616    8d4dbc                  lea ecx, dword ptr [ebp-44]
'006c2619    ffd3                    call ebx
'#Cleanup(var_58)
'006c261b    8b0e                    mov ecx, dword ptr [esi]
'006c261d    6a01                    push 01
'006c261f    51                      push ecx
'006c2620    6870374500              push 00453770
'006c2625    6a00                    push 00
'006c2627    ffd7                    call edi
'006c2629    85c0                    test ax, ax
'006c262b    0f848f010000            je 6c27c0
'006c2631    8b16                    mov edx, dword ptr [esi]
'006c2633    6a01                    push 01
'006c2635    52                      push edx
'006c2636    6870374500              push 00453770
'006c263b    6a00                    push 00
'006c263d    8975b4                  mov dword ptr [ebp-4c], esi
'006c2640    c745ac08400000          mov dword ptr [ebp-54], 00004008
'006c2647    ffd7                    call edi
'006c2649    83e801                  sub eax, 01
var_num1 = InStr(1, Replace(arg_0, "Condition.", "Conditions.", 1, -1, 0), "Avantage.", 0) - 1
'006c264c    0f8015040000            jo 6c2a67
'006c2652    50                      push eax
'006c2653    8d45ac                  lea eax, dword ptr [ebp-54]
'006c2656    50                      push eax
'006c2657    8d4dbc                  lea ecx, dword ptr [ebp-44]
'006c265a    51                      push ecx

' *** Reference to "rtcLeftCharVar"
'006c265b    ff15bc124000            call dword ptr [004012bc]
'006c2661    8d55bc                  lea edx, dword ptr [ebp-44]
'006c2664    52                      push edx

' *** Reference to "__vbaStrVarMove"
'006c2665    ff153c104000            call dword ptr [0040103c]
'006c266b    8bd0                    mov edx, eax
'006c266d    8d4dd4                  lea ecx, dword ptr [ebp-2c]

' *** Reference to "__vbaStrMove"
'006c2670    ff15d0124000            call dword ptr [004012d0]
'006c2676    8d4dbc                  lea ecx, dword ptr [ebp-44]
'006c2679    ffd3                    call ebx
'#Cleanup(var_58)
'006c267b    8b4508                  mov eax, dword ptr [ebp+08]
'006c267e    8b08                    mov ecx, dword ptr [eax]
'006c2680    50                      push eax
'006c2681    ff9118030000            call dword ptr [ecx+00000318]
'006c2687    50                      push eax
'006c2688    8d55cc                  lea edx, dword ptr [ebp-34]
'006c268b    52                      push edx

' *** Reference to "__vbaObjSet"
'006c268c    ff15b4104000            call dword ptr [004010b4]
Set var_43 = Me
'006c2692    8b4dd4                  mov ecx, dword ptr [ebp-2c]
'006c2695    6a01                    push 01
'006c2697    51                      push ecx
'006c2698    8945a8                  mov dword ptr [ebp-58], eax
'006c269b    6854374500              push 00453754
'006c26a0    8d45d4                  lea eax, dword ptr [ebp-2c]
'006c26a3    6a00                    push 00
'006c26a5    8945b4                  mov dword ptr [ebp-4c], eax
'006c26a8    c745ac08400000          mov dword ptr [ebp-54], 00004008
'006c26af    ffd7                    call edi
'006c26b1    8b55d4                  mov edx, dword ptr [ebp-2c]
'006c26b4    52                      push edx
'006c26b5    8bd8                    mov ebx, eax

' *** Reference to "__vbaLenBstr"
'006c26b7    ff1534104000            call dword ptr [00401034]
'006c26bd    2bc3                    sub eax, ebx
var_num1 = Len(Left(Replace(arg_0, "Condition.", "Conditions.", 1, -1, 0), var_num1)) - InStr(1, Left(Replace(arg_0, "Condition.", "Conditions.", 1, -1, 0), var_num1), "Conditions.", 0)
'006c26bf    0f80a2030000            jo 6c2a67
'006c26c5    83e80b                  sub eax, 0b
var_num1 = var_num1 - 11
'006c26c8    0f8099030000            jo 6c2a67
'006c26ce    50                      push eax
'006c26cf    8d45ac                  lea eax, dword ptr [ebp-54]
'006c26d2    50                      push eax
'006c26d3    8d4dbc                  lea ecx, dword ptr [ebp-44]
'006c26d6    51                      push ecx

' *** Reference to "rtcRightCharVar"
'006c26d7    ff15d8124000            call dword ptr [004012d8]
'006c26dd    8b55a8                  mov edx, dword ptr [ebp-58]
'006c26e0    8b1a                    mov ebx, dword ptr [edx]
'006c26e2    8d45bc                  lea eax, dword ptr [ebp-44]
'006c26e5    50                      push eax
'006c26e6    8d4dd0                  lea ecx, dword ptr [ebp-30]
'006c26e9    51                      push ecx

' *** Reference to "__vbaStrVarVal"
'006c26ea    ff15fc114000            call dword ptr [004011fc]
'006c26f0    8bd3                    mov edx, ebx
'006c26f2    8b5da8                  mov ebx, dword ptr [ebp-58]
'006c26f5    50                      push eax
'006c26f6    53                      push ebx
'006c26f7    ff92a4000000            call dword ptr [edx+000000a4]
'006c26fd    dbe2                    fnclex
'006c26ff    85c0                    test ax, ax
'006c2701    0f8d9c000000            jge 6c27a3
'006c2707    e985000000              jmp 6c2791
'006c270c    8b06                    mov eax, dword ptr [esi]
'006c270e    6a01                    push 01
'006c2710    50                      push eax
'006c2711    6870374500              push 00453770
'006c2716    6a00                    push 00
'006c2718    ffd7                    call edi
'006c271a    85c0                    test ax, ax
'006c271c    0f849e000000            je 6c27c0
'006c2722    8b4508                  mov eax, dword ptr [ebp+08]
'006c2725    8b08                    mov ecx, dword ptr [eax]
'006c2727    50                      push eax
'006c2728    ff910c030000            call dword ptr [ecx+0000030c]
'006c272e    50                      push eax
'006c272f    8d55cc                  lea edx, dword ptr [ebp-34]
'006c2732    52                      push edx

' *** Reference to "__vbaObjSet"
'006c2733    ff15b4104000            call dword ptr [004010b4]
Set var_43 = Me
'006c2739    8bd8                    mov ebx, eax
'006c273b    8b06                    mov eax, dword ptr [esi]
'006c273d    6a01                    push 01
'006c273f    50                      push eax
'006c2740    6870374500              push 00453770
'006c2745    6a00                    push 00
'006c2747    895da8                  mov dword ptr [ebp-58], ebx
'006c274a    8975b4                  mov dword ptr [ebp-4c], esi
'006c274d    c745ac08400000          mov dword ptr [ebp-54], 00004008
'006c2754    ffd7                    call edi
'006c2756    83e801                  sub eax, 01
var_num1 = InStr(1, Replace(arg_0, "Condition.", "Conditions.", 1, -1, 0), "Avantage.", 0) - 1
'006c2759    0f8008030000            jo 6c2a67
'006c275f    50                      push eax
'006c2760    8d4dac                  lea ecx, dword ptr [ebp-54]
'006c2763    51                      push ecx
'006c2764    8d55bc                  lea edx, dword ptr [ebp-44]
'006c2767    52                      push edx

' *** Reference to "rtcLeftCharVar"
'006c2768    ff15bc124000            call dword ptr [004012bc]
'006c276e    8b1b                    mov ebx, dword ptr [ebx]
'006c2770    8d45bc                  lea eax, dword ptr [ebp-44]
'006c2773    50                      push eax
'006c2774    8d4dd0                  lea ecx, dword ptr [ebp-30]
'006c2777    51                      push ecx

' *** Reference to "__vbaStrVarVal"
'006c2778    ff15fc114000            call dword ptr [004011fc]
'006c277e    8bd3                    mov edx, ebx
'006c2780    8b5da8                  mov ebx, dword ptr [ebp-58]
'006c2783    50                      push eax
'006c2784    53                      push ebx
'006c2785    ff92a4000000            call dword ptr [edx+000000a4]
'006c278b    dbe2                    fnclex
'006c278d    85c0                    test ax, ax
'006c278f    7d12                    jge 6c27a3
'006c2791    68a4000000              push 000000a4
'006c2796    68d00d4300              push 00430dd0
'006c279b    53                      push ebx
'006c279c    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c279d    ff1580104000            call dword ptr [00401080]
'006c27a3    8d4dd0                  lea ecx, dword ptr [ebp-30]

' *** Reference to "__vbaFreeStr"
'006c27a6    ff150c134000            call dword ptr [0040130c]
'#Cleanup(var_4)
'006c27ac    8d4dcc                  lea ecx, dword ptr [ebp-34]

' *** Reference to "__vbaFreeObj"
'006c27af    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_43)

' *** Reference to "__vbaFreeVar"
'006c27b5    8b1d28104000            mov ebx, dword ptr [00401028]
'006c27bb    8d4dbc                  lea ecx, dword ptr [ebp-44]
'006c27be    ffd3                    call ebx
'#Cleanup(var_58)
'006c27c0    8b06                    mov eax, dword ptr [esi]
'006c27c2    6a01                    push 01
'006c27c4    50                      push eax
'006c27c5    6888374500              push 00453788
'006c27ca    6a00                    push 00
'006c27cc    ffd7                    call edi
'006c27ce    85c0                    test ax, ax
'006c27d0    0f8424010000            je 6c28fa
'006c27d6    8b0e                    mov ecx, dword ptr [esi]
'006c27d8    6a01                    push 01
'006c27da    51                      push ecx
'006c27db    6888374500              push 00453788
'006c27e0    6a00                    push 00
'006c27e2    8975b4                  mov dword ptr [ebp-4c], esi
'006c27e5    c745ac08400000          mov dword ptr [ebp-54], 00004008
'006c27ec    ffd7                    call edi
'006c27ee    8b16                    mov edx, dword ptr [esi]
'006c27f0    52                      push edx
'006c27f1    8bd8                    mov ebx, eax

' *** Reference to "__vbaLenBstr"
'006c27f3    ff1534104000            call dword ptr [00401034]
'006c27f9    2bc3                    sub eax, ebx
var_num1 = Len(Replace(arg_0, "Condition.", "Conditions.", 1, -1, 0)) - InStr(1, Replace(arg_0, "Condition.", "Conditions.", 1, -1, 0), "Normal.", 0)
'006c27fb    0f8066020000            jo 6c2a67
'006c2801    83e807                  sub eax, 07
var_num1 = var_num1 - 7
'006c2804    0f805d020000            jo 6c2a67
'006c280a    50                      push eax
'006c280b    8d45ac                  lea eax, dword ptr [ebp-54]
'006c280e    50                      push eax
'006c280f    8d4dbc                  lea ecx, dword ptr [ebp-44]
'006c2812    51                      push ecx

' *** Reference to "rtcRightCharVar"
'006c2813    ff15d8124000            call dword ptr [004012d8]
'006c2819    8d55bc                  lea edx, dword ptr [ebp-44]
'006c281c    52                      push edx

' *** Reference to "__vbaStrVarMove"
'006c281d    ff153c104000            call dword ptr [0040103c]
'006c2823    8bd0                    mov edx, eax
'006c2825    8d4dd4                  lea ecx, dword ptr [ebp-2c]

' *** Reference to "__vbaStrMove"
'006c2828    ff15d0124000            call dword ptr [004012d0]
'006c282e    8d4dbc                  lea ecx, dword ptr [ebp-44]

' *** Reference to "__vbaFreeVar"
'006c2831    ff1528104000            call dword ptr [00401028]
'#Cleanup(var_58)
'006c2837    8b06                    mov eax, dword ptr [esi]
'006c2839    6a01                    push 01
'006c283b    50                      push eax
'006c283c    689c374500              push 0045379c
'006c2841    6a00                    push 00
'006c2843    ffd7                    call edi
'006c2845    85c0                    test ax, ax
'006c2847    7411                    je 6c285a
'006c2849    8b4dd4                  mov ecx, dword ptr [ebp-2c]
'006c284c    6a01                    push 01
'006c284e    51                      push ecx
'006c284f    689c374500              push 0045379c
'006c2854    6a00                    push 00
'006c2856    ffd7                    call edi
'006c2858    eb0a                    jmp 6c2864
'006c285a    8b55d4                  mov edx, dword ptr [ebp-2c]
'006c285d    52                      push edx

' *** Reference to "__vbaLenBstr"
'006c285e    ff1534104000            call dword ptr [00401034]
'006c2864    8bd8                    mov ebx, eax
'006c2866    8b4508                  mov eax, dword ptr [ebp+08]
'006c2869    8b08                    mov ecx, dword ptr [eax]
'006c286b    50                      push eax
'006c286c    ff9114030000            call dword ptr [ecx+00000314]
'006c2872    50                      push eax
'006c2873    8d55cc                  lea edx, dword ptr [ebp-34]
'006c2876    52                      push edx

' *** Reference to "__vbaObjSet"
'006c2877    ff15b4104000            call dword ptr [004010b4]
Set var_43 = Me
'006c287d    83eb01                  sub ebx, 01
var_num2 = Len(Right(Replace(arg_0, "Condition.", "Conditions.", 1, -1, 0), var_num1)) - 1
'006c2880    0f80e1010000            jo 6c2a67
'006c2886    53                      push ebx
'006c2887    8d4dac                  lea ecx, dword ptr [ebp-54]
'006c288a    8945a8                  mov dword ptr [ebp-58], eax
'006c288d    51                      push ecx
'006c288e    8d55bc                  lea edx, dword ptr [ebp-44]
'006c2891    8d45d4                  lea eax, dword ptr [ebp-2c]
'006c2894    52                      push edx
'006c2895    8945b4                  mov dword ptr [ebp-4c], eax
'006c2898    c745ac08400000          mov dword ptr [ebp-54], 00004008

' *** Reference to "rtcLeftCharVar"
'006c289f    ff15bc124000            call dword ptr [004012bc]
'006c28a5    8b45a8                  mov eax, dword ptr [ebp-58]
'006c28a8    8b18                    mov ebx, dword ptr [eax]
'006c28aa    8d4dbc                  lea ecx, dword ptr [ebp-44]
'006c28ad    51                      push ecx
'006c28ae    8d55d0                  lea edx, dword ptr [ebp-30]
'006c28b1    52                      push edx

' *** Reference to "__vbaStrVarVal"
'006c28b2    ff15fc114000            call dword ptr [004011fc]
'006c28b8    50                      push eax
'006c28b9    8bc3                    mov eax, ebx
'006c28bb    8b5da8                  mov ebx, dword ptr [ebp-58]
'006c28be    53                      push ebx
'006c28bf    ff90a4000000            call dword ptr [eax+000000a4]
'006c28c5    dbe2                    fnclex
'006c28c7    85c0                    test ax, ax
'006c28c9    7d12                    jge 6c28dd
'006c28cb    68a4000000              push 000000a4
'006c28d0    68d00d4300              push 00430dd0
'006c28d5    53                      push ebx
'006c28d6    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c28d7    ff1580104000            call dword ptr [00401080]
'006c28dd    8d4dd0                  lea ecx, dword ptr [ebp-30]

' *** Reference to "__vbaFreeStr"
'006c28e0    ff150c134000            call dword ptr [0040130c]
'#Cleanup(var_4)
'006c28e6    8d4dcc                  lea ecx, dword ptr [ebp-34]

' *** Reference to "__vbaFreeObj"
'006c28e9    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_43)

' *** Reference to "__vbaFreeVar"
'006c28ef    8b1d28104000            mov ebx, dword ptr [00401028]
'006c28f5    8d4dbc                  lea ecx, dword ptr [ebp-44]
'006c28f8    ffd3                    call ebx
'#Cleanup(var_58)
'006c28fa    8b0e                    mov ecx, dword ptr [esi]
'006c28fc    6a01                    push 01
'006c28fe    51                      push ecx
'006c28ff    689c374500              push 0045379c
'006c2904    6a00                    push 00
'006c2906    ffd7                    call edi
'006c2908    85c0                    test ax, ax
'006c290a    0f84ec000000            je 6c29fc
'006c2910    8b16                    mov edx, dword ptr [esi]
'006c2912    6a01                    push 01
'006c2914    52                      push edx
'006c2915    689c374500              push 0045379c
'006c291a    6a00                    push 00
'006c291c    8975b4                  mov dword ptr [ebp-4c], esi
'006c291f    c745ac08400000          mov dword ptr [ebp-54], 00004008
'006c2926    ffd7                    call edi
'006c2928    8bf8                    mov edi, eax
'006c292a    8b06                    mov eax, dword ptr [esi]

' *** Reference to "__vbaLenBstr"
'006c292c    8b3534104000            mov esi, dword ptr [00401034]
'006c2932    50                      push eax
'006c2933    ffd6                    call esi
'006c2935    2bc7                    sub eax, edi
var_num1 = Len(Replace(arg_0, "Condition.", "Conditions.", 1, -1, 0)) - InStr(1, Replace(arg_0, "Condition.", "Conditions.", 1, -1, 0), "Spécial.", 0)
'006c2937    0f802a010000            jo 6c2a67
'006c293d    83e808                  sub eax, 08
var_num1 = var_num1 - 8
'006c2940    0f8021010000            jo 6c2a67
'006c2946    50                      push eax
'006c2947    8d4dac                  lea ecx, dword ptr [ebp-54]
'006c294a    51                      push ecx
'006c294b    8d55bc                  lea edx, dword ptr [ebp-44]
'006c294e    52                      push edx

' *** Reference to "rtcRightCharVar"
'006c294f    ff15d8124000            call dword ptr [004012d8]
'006c2955    8d45bc                  lea eax, dword ptr [ebp-44]
'006c2958    50                      push eax

' *** Reference to "__vbaStrVarMove"
'006c2959    ff153c104000            call dword ptr [0040103c]
'006c295f    8bd0                    mov edx, eax
'006c2961    8d4dd4                  lea ecx, dword ptr [ebp-2c]

' *** Reference to "__vbaStrMove"
'006c2964    ff15d0124000            call dword ptr [004012d0]
'006c296a    8d4dbc                  lea ecx, dword ptr [ebp-44]
'006c296d    ffd3                    call ebx
'#Cleanup(var_58)
'006c296f    8b4dd4                  mov ecx, dword ptr [ebp-2c]
'006c2972    51                      push ecx
'006c2973    ffd6                    call esi
'006c2975    8bf0                    mov esi, eax
'006c2977    8b4508                  mov eax, dword ptr [ebp+08]
'006c297a    8b10                    mov edx, dword ptr [eax]
'006c297c    50                      push eax
'006c297d    ff9210030000            call dword ptr [edx+00000310]
'006c2983    50                      push eax
'006c2984    8d45cc                  lea eax, dword ptr [ebp-34]
'006c2987    50                      push eax

' *** Reference to "__vbaObjSet"
'006c2988    ff15b4104000            call dword ptr [004010b4]
Set var_43 = Me
'006c298e    83ee01                  sub esi, 01
var_num8 = Len(Right(Replace(arg_0, "Condition.", "Conditions.", 1, -1, 0), var_num1)) - 1
'006c2991    0f80d0000000            jo 6c2a67
'006c2997    56                      push esi
'006c2998    8d55ac                  lea edx, dword ptr [ebp-54]
'006c299b    8bf8                    mov edi, eax
'006c299d    52                      push edx
'006c299e    8d45bc                  lea eax, dword ptr [ebp-44]
'006c29a1    8d4dd4                  lea ecx, dword ptr [ebp-2c]
'006c29a4    50                      push eax
'006c29a5    894db4                  mov dword ptr [ebp-4c], ecx
'006c29a8    c745ac08400000          mov dword ptr [ebp-54], 00004008

' *** Reference to "rtcLeftCharVar"
'006c29af    ff15bc124000            call dword ptr [004012bc]
'006c29b5    8b37                    mov esi, dword ptr [edi]
'006c29b7    8d4dbc                  lea ecx, dword ptr [ebp-44]
'006c29ba    51                      push ecx
'006c29bb    8d55d0                  lea edx, dword ptr [ebp-30]
'006c29be    52                      push edx

' *** Reference to "__vbaStrVarVal"
'006c29bf    ff15fc114000            call dword ptr [004011fc]
'006c29c5    50                      push eax
'006c29c6    57                      push edi
'006c29c7    ff96a4000000            call dword ptr [esi+000000a4]
'006c29cd    dbe2                    fnclex
'006c29cf    85c0                    test ax, ax
'006c29d1    7d12                    jge 6c29e5
'006c29d3    68a4000000              push 000000a4
'006c29d8    68d00d4300              push 00430dd0
'006c29dd    57                      push edi
'006c29de    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c29df    ff1580104000            call dword ptr [00401080]
'006c29e5    8d4dd0                  lea ecx, dword ptr [ebp-30]

' *** Reference to "__vbaFreeStr"
'006c29e8    ff150c134000            call dword ptr [0040130c]
'#Cleanup(var_4)
'006c29ee    8d4dcc                  lea ecx, dword ptr [ebp-34]

' *** Reference to "__vbaFreeObj"
'006c29f1    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_43)
'006c29f7    8d4dbc                  lea ecx, dword ptr [ebp-44]
'006c29fa    ffd3                    call ebx
'#Cleanup(var_58)
'006c29fc    68382a6c00              push 006c2a38
'006c2a01    eb2b                    jmp 6c2a2e
'006c2a03    f645fc04                test byte ptr [ebp-04], 04
'006c2a07    7409                    je 6c2a12
'006c2a09    8d4dd8                  lea ecx, dword ptr [ebp-28]

' *** Reference to "__vbaFreeVar"
'006c2a0c    ff1528104000            call dword ptr [00401028]
'#Cleanup(var_45)
'006c2a12    8d4dd0                  lea ecx, dword ptr [ebp-30]

' *** Reference to "__vbaFreeStr"
'006c2a15    ff150c134000            call dword ptr [0040130c]
'#Cleanup(var_4)
'006c2a1b    8d4dcc                  lea ecx, dword ptr [ebp-34]

' *** Reference to "__vbaFreeObj"
'006c2a1e    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_43)
'006c2a24    8d4dbc                  lea ecx, dword ptr [ebp-44]

' *** Reference to "__vbaFreeVar"
'006c2a27    ff1528104000            call dword ptr [00401028]
'#Cleanup(var_58)
'006c2a2d    c3                      ret
'006c2a2e    8d4dd4                  lea ecx, dword ptr [ebp-2c]

' *** Reference to "__vbaFreeStr"
'006c2a31    ff150c134000            call dword ptr [0040130c]
'#Cleanup(var_44)
'006c2a37    c3                      ret
'006c2a38    8b4dd8                  mov ecx, dword ptr [ebp-28]
'006c2a3b    8b4510                  mov eax, dword ptr [ebp+10]
'006c2a3e    8b55dc                  mov edx, dword ptr [ebp-24]
'006c2a41    8908                    mov dword ptr [eax], ecx
'006c2a43    8b4de0                  mov ecx, dword ptr [ebp-20]
'006c2a46    895004                  mov dword ptr [eax+04], edx
'006c2a49    8b55e4                  mov edx, dword ptr [ebp-1c]
'006c2a4c    894808                  mov dword ptr [eax+08], ecx
'006c2a4f    8b4dec                  mov ecx, dword ptr [ebp-14]
'006c2a52    5f                      pop edi
'006c2a53    89500c                  mov dword ptr [eax+0c], edx
'006c2a56    5e                      pop esi
'006c2a57    33c0                    xor eax, eax
var_num1 = Empty
    'Reference to '__except_list'
'006c2a59    64890d00000000          mov dword ptr fs:[00000000], ecx
'006c2a60    5b                      pop ebx
'006c2a61    8be5                    mov esp, ebp
'006c2a63    5d                      pop ebp
'006c2a64    c20c00                  ret 000c


End Function


Private Function sub_6C2A70(arg_0 As Unknow, arg_1 As Unknow, arg_2 As Unknow, arg_3 As Unknow, arg_4 As Unknow, arg_5 As Unknow, arg_6 As Unknow, arg_7 As Unknow, arg_8 As Unknow, arg_9 As Unknow, arg_A As Unknow, arg_B As Unknow, arg_C As Unknow, arg_D As Unknow, arg_E As Unknow, arg_F As Unknow, arg_10 As Unknow, arg_11 As Unknow, arg_12 As Unknow, arg_13 As Unknow, arg_14 As Unknow, arg_15 As Unknow, arg_16 As Unknow, arg_17 As Unknow, arg_18 As Unknow, arg_19 As Unknow, arg_1A As Unknow, arg_1B As Unknow, arg_1C As Unknow, arg_1D As Unknow, arg_1E As Unknow, arg_1F As Unknow, arg_20 As Unknow, arg_21 As Unknow, arg_22 As Unknow, arg_23 As Unknow, arg_24 As Unknow, arg_25 As Unknow, arg_26 As Unknow, arg_27 As Unknow, arg_28 As Unknow, arg_29 As Unknow, arg_2A As Unknow, arg_2B As Unknow, arg_2C As Unknow, arg_2D As Unknow, arg_2E As Unknow, arg_2F As Unknow, arg_30 As Unknow, arg_31 As Unknow, arg_32 As Unknow, arg_33 As Unknow, arg_34 As Unknow, arg_35 As Unknow, arg_36 As Unknow, arg_37 As Unknow, arg_38 As Unknow, arg_39 As Unknow, arg_3A As Unknow, arg_3B As Unknow, arg_3C As Unknow)
'006c2a70    55                      push ebp
'006c2a71    8bec                    mov ebp, esp
'006c2a73    83ec0c                  sub esp, 0c

' *** Reference to "__vbaExceptHandler"
'006c2a76    6866724000              push 00407266
'006c2a7b    64a100000000            mov ax, word ptr fs:[00000000]
'006c2a81    50                      push eax
    'Reference to '__except_list'
'006c2a82    64892500000000          mov dword ptr fs:[00000000], esp
'006c2a89    83ec34                  sub esp, 34
'006c2a8c    53                      push ebx
'006c2a8d    56                      push esi
'006c2a8e    57                      push edi
'006c2a8f    8965f4                  mov dword ptr [ebp-0c], esp
'006c2a92    c745f810664000          mov dword ptr [ebp-08], 00406610
'006c2a99    8b4510                  mov eax, dword ptr [ebp+10]
'006c2a9c    8b4d0c                  mov ecx, dword ptr [ebp+0c]
'006c2a9f    33db                    xor ebx, ebx
'006c2aa1    8918                    mov dword ptr [eax], ebx
'006c2aa3    8b11                    mov edx, dword ptr [ecx]
'006c2aa5    8d4de4                  lea ecx, dword ptr [ebp-1c]
'006c2aa8    895de8                  mov dword ptr [ebp-18], ebx
'006c2aab    895de4                  mov dword ptr [ebp-1c], ebx
'006c2aae    895de0                  mov dword ptr [ebp-20], ebx
'006c2ab1    895dd0                  mov dword ptr [ebp-30], ebx
'006c2ab4    895dc0                  mov dword ptr [ebp-40], ebx

' *** Reference to "__vbaStrCopy"
'006c2ab7    ff1548124000            call dword ptr [00401248]
var_40 = (arg_0)
'006c2abd    8d45c0                  lea eax, dword ptr [ebp-40]
'006c2ac0    50                      push eax
'006c2ac1    8d4dd0                  lea ecx, dword ptr [ebp-30]
'006c2ac4    8d55e4                  lea edx, dword ptr [ebp-1c]
'006c2ac7    51                      push ecx
'006c2ac8    8955c8                  mov dword ptr [ebp-38], edx
'006c2acb    c745c008400000          mov dword ptr [ebp-40], 00004008

' *** Reference to "rtcTrimVar"
'006c2ad2    ff15e4104000            call dword ptr [004010e4]
'006c2ad8    8d55d0                  lea edx, dword ptr [ebp-30]
'006c2adb    52                      push edx

' *** Reference to "__vbaStrVarMove"
'006c2adc    ff153c104000            call dword ptr [0040103c]

' *** Reference to "__vbaStrMove"
'006c2ae2    8b35d0124000            mov esi, dword ptr [004012d0]
'006c2ae8    8bd0                    mov edx, eax
'006c2aea    8d4de4                  lea ecx, dword ptr [ebp-1c]
'006c2aed    ffd6                    call esi
'006c2aef    8d4dd0                  lea ecx, dword ptr [ebp-30]

' *** Reference to "__vbaFreeVar"
'006c2af2    ff1528104000            call dword ptr [00401028]
'#Cleanup(var_4)
'006c2af8    8b45e4                  mov eax, dword ptr [ebp-1c]

' *** Reference to "rtcReplace"
'006c2afb    8b3db4114000            mov edi, dword ptr [004011b4]
'006c2b01    53                      push ebx
'006c2b02    6aff                    push ffffffff
'006c2b04    6a01                    push 01
'006c2b06    68cc134300              push 004313cc
'006c2b0b    6854374500              push 00453754
'006c2b10    50                      push eax
'006c2b11    ffd7                    call edi
'006c2b13    8bd0                    mov edx, eax
'006c2b15    8d4de4                  lea ecx, dword ptr [ebp-1c]
'006c2b18    ffd6                    call esi
'006c2b1a    8b4de4                  mov ecx, dword ptr [ebp-1c]
'006c2b1d    53                      push ebx
'006c2b1e    6aff                    push ffffffff
'006c2b20    6a01                    push 01
'006c2b22    68cc134300              push 004313cc
'006c2b27    6870374500              push 00453770
'006c2b2c    51                      push ecx
'006c2b2d    ffd7                    call edi
'006c2b2f    8bd0                    mov edx, eax
'006c2b31    8d4de4                  lea ecx, dword ptr [ebp-1c]
'006c2b34    ffd6                    call esi
'006c2b36    8b55e4                  mov edx, dword ptr [ebp-1c]
'006c2b39    53                      push ebx
'006c2b3a    6aff                    push ffffffff
'006c2b3c    6a01                    push 01
'006c2b3e    68cc134300              push 004313cc
'006c2b43    6888374500              push 00453788
'006c2b48    52                      push edx
'006c2b49    ffd7                    call edi
'006c2b4b    8bd0                    mov edx, eax
'006c2b4d    8d4de4                  lea ecx, dword ptr [ebp-1c]
'006c2b50    ffd6                    call esi
'006c2b52    8b45e4                  mov eax, dword ptr [ebp-1c]
'006c2b55    53                      push ebx
'006c2b56    6aff                    push ffffffff
'006c2b58    6a01                    push 01
'006c2b5a    68cc134300              push 004313cc
'006c2b5f    689c374500              push 0045379c
'006c2b64    50                      push eax
'006c2b65    ffd7                    call edi
'006c2b67    8bd0                    mov edx, eax
'006c2b69    8d4de4                  lea ecx, dword ptr [ebp-1c]
'006c2b6c    ffd6                    call esi
'006c2b6e    53                      push ebx
'006c2b6f    6aff                    push ffffffff

' *** Reference to "__vbaStrCat"
'006c2b71    8b1d70104000            mov ebx, dword ptr [00401070]
'006c2b77    6a01                    push 01
'006c2b79    68b4374500              push 004537b4
'006c2b7e    6830b94300              push 0043b930
'006c2b83    6870084300              push 00430870
'006c2b88    ffd3                    call ebx
var_13 = (".") & (vbCrLf)
'006c2b8a    8bd0                    mov edx, eax
'006c2b8c    8d4de0                  lea ecx, dword ptr [ebp-20]
'006c2b8f    ffd6                    call esi
'006c2b91    8b4de4                  mov ecx, dword ptr [ebp-1c]
'006c2b94    50                      push eax
'006c2b95    51                      push ecx
'006c2b96    ffd7                    call edi
'006c2b98    8bd0                    mov edx, eax
'006c2b9a    8d4de4                  lea ecx, dword ptr [ebp-1c]
'006c2b9d    ffd6                    call esi
'006c2b9f    8d4de0                  lea ecx, dword ptr [ebp-20]

' *** Reference to "__vbaFreeStr"
'006c2ba2    ff150c134000            call dword ptr [0040130c]
'#Cleanup(var_3)
'006c2ba8    6a00                    push 00
'006c2baa    6aff                    push ffffffff
'006c2bac    6a01                    push 01
'006c2bae    68b4374500              push 004537b4
'006c2bb3    689c384400              push 0044389c
'006c2bb8    6870084300              push 00430870
'006c2bbd    ffd3                    call ebx
var_127 = (". ") & (vbCrLf)
'006c2bbf    8bd0                    mov edx, eax
'006c2bc1    8d4de0                  lea ecx, dword ptr [ebp-20]
'006c2bc4    ffd6                    call esi
'006c2bc6    8b55e4                  mov edx, dword ptr [ebp-1c]
'006c2bc9    50                      push eax
'006c2bca    52                      push edx
'006c2bcb    ffd7                    call edi
'006c2bcd    8bd0                    mov edx, eax
'006c2bcf    8d4de4                  lea ecx, dword ptr [ebp-1c]
'006c2bd2    ffd6                    call esi
'006c2bd4    8d4de0                  lea ecx, dword ptr [ebp-20]

' *** Reference to "__vbaFreeStr"
'006c2bd7    ff150c134000            call dword ptr [0040130c]
'#Cleanup(var_3)
'006c2bdd    6a00                    push 00
'006c2bdf    6aff                    push ffffffff
'006c2be1    6a01                    push 01
'006c2be3    68c8374500              push 004537c8
'006c2be8    68c0374500              push 004537c0
'006c2bed    6870084300              push 00430870
'006c2bf2    ffd3                    call ebx
var_16 = (":") & (vbCrLf)
'006c2bf4    8bd0                    mov edx, eax
'006c2bf6    8d4de0                  lea ecx, dword ptr [ebp-20]
'006c2bf9    ffd6                    call esi
'006c2bfb    50                      push eax
'006c2bfc    8b45e4                  mov eax, dword ptr [ebp-1c]
'006c2bff    50                      push eax
'006c2c00    ffd7                    call edi
'006c2c02    8bd0                    mov edx, eax
'006c2c04    8d4de4                  lea ecx, dword ptr [ebp-1c]
'006c2c07    ffd6                    call esi
'006c2c09    8d4de0                  lea ecx, dword ptr [ebp-20]

' *** Reference to "__vbaFreeStr"
'006c2c0c    ff150c134000            call dword ptr [0040130c]
'#Cleanup(var_3)
'006c2c12    6a00                    push 00
'006c2c14    6aff                    push ffffffff
'006c2c16    6a01                    push 01
'006c2c18    68c8374500              push 004537c8
'006c2c1d    6814544400              push 00445414
'006c2c22    6870084300              push 00430870
'006c2c27    ffd3                    call ebx
var_17 = (": ") & (vbCrLf)
'006c2c29    8bd0                    mov edx, eax
'006c2c2b    8d4de0                  lea ecx, dword ptr [ebp-20]
'006c2c2e    ffd6                    call esi
'006c2c30    8b4de4                  mov ecx, dword ptr [ebp-1c]
'006c2c33    50                      push eax
'006c2c34    51                      push ecx
'006c2c35    ffd7                    call edi
'006c2c37    8bd0                    mov edx, eax
'006c2c39    8d4de4                  lea ecx, dword ptr [ebp-1c]
'006c2c3c    ffd6                    call esi
'006c2c3e    8d4de0                  lea ecx, dword ptr [ebp-20]

' *** Reference to "__vbaFreeStr"
'006c2c41    ff150c134000            call dword ptr [0040130c]
'#Cleanup(var_3)
'006c2c47    8b55e4                  mov edx, dword ptr [ebp-1c]
'006c2c4a    6a00                    push 00
'006c2c4c    6aff                    push ffffffff
'006c2c4e    6a01                    push 01
'006c2c50    68c48d4300              push 00438dc4
'006c2c55    6870084300              push 00430870
'006c2c5a    52                      push edx
'006c2c5b    ffd7                    call edi
'006c2c5d    8bd0                    mov edx, eax
'006c2c5f    8d4de4                  lea ecx, dword ptr [ebp-1c]
'006c2c62    ffd6                    call esi
'006c2c64    8b45e4                  mov eax, dword ptr [ebp-1c]
'006c2c67    6a00                    push 00
'006c2c69    6aff                    push ffffffff
'006c2c6b    6a01                    push 01
'006c2c6d    68c48d4300              push 00438dc4
'006c2c72    6864164400              push 00441664
'006c2c77    50                      push eax
'006c2c78    ffd7                    call edi
'006c2c7a    8bd0                    mov edx, eax
'006c2c7c    8d4de4                  lea ecx, dword ptr [ebp-1c]
'006c2c7f    ffd6                    call esi
'006c2c81    6a00                    push 00
'006c2c83    6aff                    push ffffffff
'006c2c85    6a01                    push 01
'006c2c87    6830b94300              push 0043b930
'006c2c8c    6870084300              push 00430870
'006c2c91    ffd3                    call ebx
var_54 = (".") & (vbCrLf)
'006c2c93    8bd0                    mov edx, eax
'006c2c95    8d4de0                  lea ecx, dword ptr [ebp-20]
'006c2c98    ffd6                    call esi
'006c2c9a    8b4de4                  mov ecx, dword ptr [ebp-1c]
'006c2c9d    50                      push eax
'006c2c9e    68b4374500              push 004537b4
'006c2ca3    51                      push ecx
'006c2ca4    ffd7                    call edi
'006c2ca6    8bd0                    mov edx, eax
'006c2ca8    8d4de4                  lea ecx, dword ptr [ebp-1c]
'006c2cab    ffd6                    call esi
'006c2cad    8d4de0                  lea ecx, dword ptr [ebp-20]

' *** Reference to "__vbaFreeStr"
'006c2cb0    ff150c134000            call dword ptr [0040130c]
'#Cleanup(var_3)
'006c2cb6    6a00                    push 00
'006c2cb8    6aff                    push ffffffff
'006c2cba    6a01                    push 01
'006c2cbc    68c0374500              push 004537c0
'006c2cc1    6870084300              push 00430870
'006c2cc6    ffd3                    call ebx
var_139 = (":") & (vbCrLf)
'006c2cc8    8bd0                    mov edx, eax
'006c2cca    8d4de0                  lea ecx, dword ptr [ebp-20]
'006c2ccd    ffd6                    call esi
'006c2ccf    8b55e4                  mov edx, dword ptr [ebp-1c]
'006c2cd2    50                      push eax
'006c2cd3    68c8374500              push 004537c8
'006c2cd8    52                      push edx
'006c2cd9    ffd7                    call edi
'006c2cdb    8bd0                    mov edx, eax
'006c2cdd    8d4de4                  lea ecx, dword ptr [ebp-1c]
'006c2ce0    ffd6                    call esi
'006c2ce2    8d4de0                  lea ecx, dword ptr [ebp-20]

' *** Reference to "__vbaFreeStr"
'006c2ce5    ff150c134000            call dword ptr [0040130c]
'#Cleanup(var_3)
'006c2ceb    8b45e4                  mov eax, dword ptr [ebp-1c]

' *** Reference to "__vbaLenBstr"
'006c2cee    8b3d34104000            mov edi, dword ptr [00401034]
'006c2cf4    50                      push eax
'006c2cf5    ffd7                    call edi
'006c2cf7    83f801                  cmp eax, 01
'006c2cfa    7e6d                    jle 6c2d69

If (Len(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Trim(arg_0), "Conditions.", vbNullChar, 1, -1, 0), "Avantage.", vbNullChar, 1, -1, 0), "Normal.", vbNullChar, 1, -1, 0), "Spécial.", vbNullChar, 1, -1, 0), var_13, "^p", 1, -1, 0), var_127, "^p", 1, -1, 0), var_16, "^d", 1, -1, 0), var_17, "^d", 1, -1, 0), vbCrLf, " ", 1, -1, 0), "  ", " ", 1, -1, 0), "^p", var_54, 1, -1, 0), "^d", var_139, 1, -1, 0)) > 1) Then
'006c2cfc    8b4de4                  mov ecx, dword ptr [ebp-1c]
'006c2cff    6a01                    push 01
'006c2d01    6aff                    push ffffffff
'006c2d03    6870084300              push 00430870
'006c2d08    51                      push ecx

' *** Reference to "rtcInStrRev"
'006c2d09    ff1508114000            call dword ptr [00401108]
'006c2d0f    8b55e4                  mov edx, dword ptr [ebp-1c]
'006c2d12    52                      push edx
'006c2d13    8bd8                    mov ebx, eax
'006c2d15    ffd7                    call edi
'006c2d17    83e801                  sub eax, 01
    var_num1 = Len(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Trim(arg_0), "Conditions.", vbNullChar, 1, -1, 0), "Avantage.", vbNullChar, 1, -1, 0), "Normal.", vbNullChar, 1, -1, 0), "Spécial.", vbNullChar, 1, -1, 0), var_13, "^p", 1, -1, 0), var_127, "^p", 1, -1, 0), var_16, "^d", 1, -1, 0), var_17, "^d", 1, -1, 0), vbCrLf, " ", 1, -1, 0), "  ", " ", 1, -1, 0), "^p", var_54, 1, -1, 0), "^d", var_139, 1, -1, 0)) - 1
'006c2d1a    0f80a5000000            jo 6c2dc5
'006c2d20    3bd8                    cmp ebx, eax
'006c2d22    7545                    jne 6c2d69
    
    If (    InstrRev(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Trim(arg_0), "Conditions.", vbNullChar, 1, -1, 0), "Avantage.", vbNullChar, 1, -1, 0), "Normal.", vbNullChar, 1, -1, 0), "Spécial.", vbNullChar, 1, -1, 0), var_13, "^p", 1, -1, 0), var_127, "^p", 1, -1, 0), var_16, "^d", 1, -1, 0), var_17, "^d", 1, -1, 0), vbCrLf, " ", 1, -1, 0), "  ", " ", 1, -1, 0), "^p", var_54, 1, -1, 0), "^d", var_139, 1, -1, 0), vbCrLf, -1, 1) = var_num1) Then
'006c2d24    8b4de4                  mov ecx, dword ptr [ebp-1c]
'006c2d27    8d45e4                  lea eax, dword ptr [ebp-1c]
'006c2d2a    51                      push ecx
'006c2d2b    8945c8                  mov dword ptr [ebp-38], eax
'006c2d2e    c745c008400000          mov dword ptr [ebp-40], 00004008
'006c2d35    ffd7                    call edi
'006c2d37    83e802                  sub eax, 02
    var_num1 = Len(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Trim(arg_0), "Conditions.", vbNullChar, 1, -1, 0), "Avantage.", vbNullChar, 1, -1, 0), "Normal.", vbNullChar, 1, -1, 0), "Spécial.", vbNullChar, 1, -1, 0), var_13, "^p", 1, -1, 0), var_127, "^p", 1, -1, 0), var_16, "^d", 1, -1, 0), var_17, "^d", 1, -1, 0), vbCrLf, " ", 1, -1, 0), "  ", " ", 1, -1, 0), "^p", var_54, 1, -1, 0), "^d", var_139, 1, -1, 0)) - 2
'006c2d3a    0f8085000000            jo 6c2dc5
'006c2d40    50                      push eax
'006c2d41    8d55c0                  lea edx, dword ptr [ebp-40]
'006c2d44    52                      push edx
'006c2d45    8d45d0                  lea eax, dword ptr [ebp-30]
'006c2d48    50                      push eax

' *** Reference to "rtcLeftCharVar"
'006c2d49    ff15bc124000            call dword ptr [004012bc]
'006c2d4f    8d4dd0                  lea ecx, dword ptr [ebp-30]
'006c2d52    51                      push ecx

' *** Reference to "__vbaStrVarMove"
'006c2d53    ff153c104000            call dword ptr [0040103c]
'006c2d59    8bd0                    mov edx, eax
'006c2d5b    8d4de4                  lea ecx, dword ptr [ebp-1c]
'006c2d5e    ffd6                    call esi
'006c2d60    8d4dd0                  lea ecx, dword ptr [ebp-30]

' *** Reference to "__vbaFreeVar"
'006c2d63    ff1528104000            call dword ptr [00401028]
    '#Cleanup(var_4)
End If
'006c2d69    8b55e4                  mov edx, dword ptr [ebp-1c]
'006c2d6c    8d4de8                  lea ecx, dword ptr [ebp-18]

' *** Reference to "__vbaStrCopy"
'006c2d6f    ff1548124000            call dword ptr [00401248]
var_41 = (Left(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Trim(arg_0), "Conditions.", vbNullChar, 1, -1, 0), "Avantage.", vbNullChar, 1, -1, 0), "Normal.", vbNullChar, 1, -1, 0), "Spécial.", vbNullChar, 1, -1, 0), var_13, "^p", 1, -1, 0), var_127, "^p", 1, -1, 0), var_16, "^d", 1, -1, 0), var_17, "^d", 1, -1, 0), vbCrLf, " ", 1, -1, 0), "  ", " ", 1, -1, 0), "^p", var_54, 1, -1, 0), "^d", var_139, 1, -1, 0), var_num1))
'006c2d75    68a82d6c00              push 006c2da8
'006c2d7a    eb22                    jmp 6c2d9e
'006c2d7c    f645fc04                test byte ptr [ebp-04], 04
'006c2d80    7409                    je 6c2d8b
'006c2d82    8d4de8                  lea ecx, dword ptr [ebp-18]

' *** Reference to "__vbaFreeStr"
'006c2d85    ff150c134000            call dword ptr [0040130c]
'#Cleanup(var_41)
'006c2d8b    8d4de0                  lea ecx, dword ptr [ebp-20]

' *** Reference to "__vbaFreeStr"
'006c2d8e    ff150c134000            call dword ptr [0040130c]
'#Cleanup(var_3)
'006c2d94    8d4dd0                  lea ecx, dword ptr [ebp-30]

' *** Reference to "__vbaFreeVar"
'006c2d97    ff1528104000            call dword ptr [00401028]
'#Cleanup(var_4)
'006c2d9d    c3                      ret
'006c2d9e    8d4de4                  lea ecx, dword ptr [ebp-1c]

' *** Reference to "__vbaFreeStr"
'006c2da1    ff150c134000            call dword ptr [0040130c]
'#Cleanup(var_40)
'006c2da7    c3                      ret
'006c2da8    8b45e8                  mov eax, dword ptr [ebp-18]
'006c2dab    8b5510                  mov edx, dword ptr [ebp+10]
'006c2dae    8b4dec                  mov ecx, dword ptr [ebp-14]
'006c2db1    5f                      pop edi
'006c2db2    8902                    mov dword ptr [edx], eax
'006c2db4    5e                      pop esi
'006c2db5    33c0                    xor eax, eax
var_num1 = Empty
    'Reference to '__except_list'
'006c2db7    64890d00000000          mov dword ptr fs:[00000000], ecx
'006c2dbe    5b                      pop ebx
'006c2dbf    8be5                    mov esp, ebp
'006c2dc1    5d                      pop ebp
'006c2dc2    c20c00                  ret 000c


End Function


Public Function Affiche(arg_0 As Unknow, arg_1 As Unknow, arg_2 As Unknow, arg_3 As Unknow, arg_4 As Unknow, arg_5 As Unknow, arg_6 As Unknow, arg_7 As Unknow, arg_8 As Unknow, arg_9 As Unknow, arg_A As Unknow, arg_B As Unknow, arg_C As Unknow, arg_D As Unknow, arg_E As Unknow, arg_F As Unknow, arg_10 As Unknow, arg_11 As Unknow, arg_12 As Unknow, arg_13 As Unknow, arg_14 As Unknow, arg_15 As Unknow, arg_16 As Unknow, arg_17 As Unknow, arg_18 As Unknow, arg_19 As Unknow, arg_1A As Unknow, arg_1B As Unknow, arg_1C As Unknow, arg_1D As Unknow, arg_1E As Unknow, arg_1F As Unknow, arg_20 As Unknow, arg_21 As Unknow, arg_22 As Unknow, arg_23 As Unknow, arg_24 As Unknow, arg_25 As Unknow, arg_26 As Unknow, arg_27 As Unknow, arg_28 As Unknow, arg_29 As Unknow, arg_2A As Unknow, arg_2B As Unknow, arg_2C As Unknow, arg_2D As Unknow, arg_2E As Unknow, arg_2F As Unknow, arg_30 As Unknow, arg_31 As Unknow, arg_32 As Unknow, arg_33 As Unknow, arg_34 As Unknow, arg_35 As Unknow, arg_36 As Unknow, arg_37 As Unknow, arg_38 As Unknow, arg_39 As Unknow, arg_3A As Unknow, arg_3B As Unknow, arg_3C As Unknow)
'006c39c0    55                      push ebp
'006c39c1    8bec                    mov ebp, esp
'006c39c3    83ec0c                  sub esp, 0c

' *** Reference to "__vbaExceptHandler"
'006c39c6    6866724000              push 00407266
'006c39cb    64a100000000            mov ax, word ptr fs:[00000000]
'006c39d1    50                      push eax
    'Reference to '__except_list'
'006c39d2    64892500000000          mov dword ptr fs:[00000000], esp
'006c39d9    81ec40010000            sub esp, 00000140
'006c39df    53                      push ebx
'006c39e0    56                      push esi
'006c39e1    57                      push edi
'006c39e2    8965f4                  mov dword ptr [ebp-0c], esp
'006c39e5    c745f870664000          mov dword ptr [ebp-08], 00406670
'006c39ec    33ff                    xor edi, edi
'006c39ee    897dfc                  mov dword ptr [ebp-04], edi
'006c39f1    8b5d08                  mov ebx, dword ptr [ebp+08]
'006c39f4    8b03                    mov eax, dword ptr [ebx]
'006c39f6    53                      push ebx
'006c39f7    ff5004                  call dword ptr [eax+04]
'006c39fa    8b0b                    mov ecx, dword ptr [ebx]
'006c39fc    53                      push ebx
'006c39fd    897de8                  mov dword ptr [ebp-18], edi
'006c3a00    897de4                  mov dword ptr [ebp-1c], edi
'006c3a03    897de0                  mov dword ptr [ebp-20], edi
'006c3a06    897ddc                  mov dword ptr [ebp-24], edi
'006c3a09    897dd8                  mov dword ptr [ebp-28], edi
'006c3a0c    897dd4                  mov dword ptr [ebp-2c], edi
'006c3a0f    897dd0                  mov dword ptr [ebp-30], edi
'006c3a12    897dcc                  mov dword ptr [ebp-34], edi
'006c3a15    897dc8                  mov dword ptr [ebp-38], edi
'006c3a18    897dc4                  mov dword ptr [ebp-3c], edi
'006c3a1b    897dc0                  mov dword ptr [ebp-40], edi
'006c3a1e    897dbc                  mov dword ptr [ebp-44], edi
'006c3a21    897dac                  mov dword ptr [ebp-54], edi
'006c3a24    897d9c                  mov dword ptr [ebp-64], edi
'006c3a27    897d8c                  mov dword ptr [ebp-74], edi
'006c3a2a    89bd7cffffff            mov dword ptr [ebp+ffffff7c], edi
'006c3a30    89bd6cffffff            mov dword ptr [ebp+ffffff6c], edi
'006c3a36    89bd5cffffff            mov dword ptr [ebp+ffffff5c], edi
'006c3a3c    89bd4cffffff            mov dword ptr [ebp+ffffff4c], edi
'006c3a42    89bd3cffffff            mov dword ptr [ebp+ffffff3c], edi
'006c3a48    89bd2cffffff            mov dword ptr [ebp+ffffff2c], edi
'006c3a4e    89bd28ffffff            mov dword ptr [ebp+ffffff28], edi
'006c3a54    ff913c030000            call dword ptr [ecx+0000033c]
'006c3a5a    50                      push eax
'006c3a5b    8d55cc                  lea edx, dword ptr [ebp-34]
'006c3a5e    52                      push edx

' *** Reference to "__vbaObjSet"
'006c3a5f    ff15b4104000            call dword ptr [004010b4]
Set var_43 = Nothing
'006c3a65    8d4de4                  lea ecx, dword ptr [ebp-1c]
'006c3a68    8bf0                    mov esi, eax
'006c3a6a    8b06                    mov eax, dword ptr [esi]
'006c3a6c    51                      push ecx
'006c3a6d    56                      push esi
'006c3a6e    ff5050                  call dword ptr [eax+50]
'006c3a71    dbe2                    fnclex
'006c3a73    3bc7                    cmp eax, edi
'006c3a75    7d0f                    jge 6c3a86

If (-256 - 12 < 0) Then
'006c3a77    6a50                    push 50
'006c3a79    683ce04300              push 0043e03c
'006c3a7e    56                      push esi
'006c3a7f    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c3a80    ff1580104000            call dword ptr [00401080]
    
End If
'006c3a86    8b55e4                  mov edx, dword ptr [ebp-1c]
'006c3a89    52                      push edx
'006c3a8a    68cc134300              push 004313cc

' *** Reference to "__vbaStrCmp"
'006c3a8f    ff1530114000            call dword ptr [00401130]
'006c3a95    8bf0                    mov esi, eax
'006c3a97    f7de                    neg esi
'006c3a99    1bf6                    sbb esi, esi
'006c3a9b    f7de                    neg esi
'006c3a9d    8d4de4                  lea ecx, dword ptr [ebp-1c]
'006c3aa0    f7de                    neg esi

' *** Reference to "__vbaFreeStr"
'006c3aa2    ff150c134000            call dword ptr [0040130c]
'#Cleanup(var_40)
'006c3aa8    8d4dcc                  lea ecx, dword ptr [ebp-34]

' *** Reference to "__vbaFreeObj"
'006c3aab    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_43)
'006c3ab1    663bf7                  cmp si, di
'006c3ab4    0f84dd150000            je 6c5097

If (((vbNullString) <> (vbNullChar))) Then
'006c3aba    8b03                    mov eax, dword ptr [ebx]
'006c3abc    53                      push ebx
'006c3abd    ff903c030000            call dword ptr [eax+0000033c]
'006c3ac3    50                      push eax
'006c3ac4    8d4dcc                  lea ecx, dword ptr [ebp-34]
'006c3ac7    51                      push ecx

' *** Reference to "__vbaObjSet"
'006c3ac8    ff15b4104000            call dword ptr [004010b4]
    Set var_43 = Nothing
'006c3ace    8bf0                    mov esi, eax
'006c3ad0    8b16                    mov edx, dword ptr [esi]
'006c3ad2    8d45e4                  lea eax, dword ptr [ebp-1c]
'006c3ad5    50                      push eax
'006c3ad6    56                      push esi
'006c3ad7    ff5250                  call dword ptr [edx+50]
'006c3ada    dbe2                    fnclex
'006c3adc    3bc7                    cmp eax, edi
'006c3ade    7d0f                    jge 6c3aef
'006c3ae0    6a50                    push 50
'006c3ae2    683ce04300              push 0043e03c
'006c3ae7    56                      push esi
'006c3ae8    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c3ae9    ff1580104000            call dword ptr [00401080]
'006c3aef    8b55e4                  mov edx, dword ptr [ebp-1c]

' *** Reference to "__vbaStrMove"
'006c3af2    8b35d0124000            mov esi, dword ptr [004012d0]
'006c3af8    8d4de0                  lea ecx, dword ptr [ebp-20]
'006c3afb    897de4                  mov dword ptr [ebp-1c], edi
'006c3afe    ffd6                    call esi
'006c3b00    8d4de0                  lea ecx, dword ptr [ebp-20]
'006c3b03    51                      push ecx
'006c3b04    e8e700e3ff              call 4f3bf0
    Call sub_4F3BF0()
'006c3b09    8bd0                    mov edx, eax
'006c3b0b    8d4dd0                  lea ecx, dword ptr [ebp-30]
'006c3b0e    ffd6                    call esi
'006c3b10    8b55d0                  mov edx, dword ptr [ebp-30]
'006c3b13    8995f8feffff            mov dword ptr [ebp+fffffef8], edx
'006c3b19    8b154cb07200            mov edx, dword ptr [0072b04c]
'006c3b1f    b804000280              mov eax, 80020004
'006c3b24    898544ffffff            mov dword ptr [ebp+ffffff44], eax
'006c3b2a    b90a000000              mov ecx, 0000000a
'006c3b2f    898d3cffffff            mov dword ptr [ebp+ffffff3c], ecx
'006c3b35    898d4cffffff            mov dword ptr [ebp+ffffff4c], ecx
'006c3b3b    897dd0                  mov dword ptr [ebp-30], edi
'006c3b3e    8b1a                    mov ebx, dword ptr [edx]
'006c3b40    8d55c8                  lea edx, dword ptr [ebp-38]
'006c3b43    52                      push edx
'006c3b44    83ec10                  sub esp, 10
'006c3b47    8bd4                    mov edx, esp
'006c3b49    890a                    mov dword ptr [edx], ecx
'006c3b4b    8b8d40ffffff            mov ecx, dword ptr [ebp+ffffff40]
'006c3b51    894a04                  mov dword ptr [edx+04], ecx
'006c3b54    894208                  mov dword ptr [edx+08], eax
'006c3b57    898554ffffff            mov dword ptr [ebp+ffffff54], eax
'006c3b5d    8b8548ffffff            mov eax, dword ptr [ebp+ffffff48]
'006c3b63    89420c                  mov dword ptr [edx+0c], eax
'006c3b66    8b954cffffff            mov edx, dword ptr [ebp+ffffff4c]
'006c3b6c    8b8550ffffff            mov eax, dword ptr [ebp+ffffff50]
'006c3b72    83ec10                  sub esp, 10
'006c3b75    8bcc                    mov ecx, esp
'006c3b77    8911                    mov dword ptr [ecx], edx
'006c3b79    8b9554ffffff            mov edx, dword ptr [ebp+ffffff54]
'006c3b7f    894104                  mov dword ptr [ecx+04], eax
'006c3b82    8b8558ffffff            mov eax, dword ptr [ebp+ffffff58]
'006c3b88    895108                  mov dword ptr [ecx+08], edx
'006c3b8b    8b9560ffffff            mov edx, dword ptr [ebp+ffffff60]
'006c3b91    89410c                  mov dword ptr [ecx+0c], eax
'006c3b94    83ec10                  sub esp, 10
'006c3b97    8bcc                    mov ecx, esp
'006c3b99    b803000000              mov eax, 00000003
'006c3b9e    8901                    mov dword ptr [ecx], eax
'006c3ba0    895104                  mov dword ptr [ecx+04], edx
'006c3ba3    8b95f8feffff            mov edx, dword ptr [ebp+fffffef8]
'006c3ba9    b802000000              mov eax, 00000002
'006c3bae    894108                  mov dword ptr [ecx+08], eax
'006c3bb1    8b8568ffffff            mov eax, dword ptr [ebp+ffffff68]
'006c3bb7    89410c                  mov dword ptr [ecx+0c], eax
'006c3bba    68d4374500              push 004537d4
'006c3bbf    8d4ddc                  lea ecx, dword ptr [ebp-24]
'006c3bc2    ffd6                    call esi
'006c3bc4    50                      push eax

' *** Reference to "__vbaStrCat"
'006c3bc5    ff1570104000            call dword ptr [00401070]
    var_49 = ("select * from dons where nom='") & (vbNullString)
'006c3bcb    8bd0                    mov edx, eax
'006c3bcd    8d4dd8                  lea ecx, dword ptr [ebp-28]
'006c3bd0    ffd6                    call esi
'006c3bd2    50                      push eax
'006c3bd3    6854a44300              push 0043a454

' *** Reference to "__vbaStrCat"
'006c3bd8    ff1570104000            call dword ptr [00401070]
    var_63 = (var_49) & ("'")
'006c3bde    8bd0                    mov edx, eax
'006c3be0    8d4dd4                  lea ecx, dword ptr [ebp-2c]
'006c3be3    ffd6                    call esi
'006c3be5    8b0d4cb07200            mov ecx, dword ptr [0072b04c]
'006c3beb    50                      push eax
'006c3bec    51                      push ecx
'006c3bed    ff93bc000000            call dword ptr [ebx+000000bc]
'006c3bf3    dbe2                    fnclex
'006c3bf5    3bc7                    cmp eax, edi
'006c3bf7    7d1c                    jge 6c3c15
    
    If (    -4504 < 0) Then
'006c3bf9    8b154cb07200            mov edx, dword ptr [0072b04c]

' *** Reference to "__vbaHresultCheckObj"
'006c3bff    8b3580104000            mov esi, dword ptr [00401080]
'006c3c05    68bc000000              push 000000bc
'006c3c0a    68301f4300              push 00431f30
'006c3c0f    52                      push edx
'006c3c10    50                      push eax
'006c3c11    ffd6                    call esi
'006c3c13    eb06                    jmp 6c3c1b
    
End If

' *** Reference to "__vbaHresultCheckObj"
'006c3c15    8b3580104000            mov esi, dword ptr [00401080]
'006c3c1b    8b45c8                  mov eax, dword ptr [ebp-38]
'006c3c1e    50                      push eax
'006c3c1f    8d45e8                  lea eax, dword ptr [ebp-18]
'006c3c22    50                      push eax
'006c3c23    897dc8                  mov dword ptr [ebp-38], edi

' *** Reference to "__vbaObjSet"
'006c3c26    ff15b4104000            call dword ptr [004010b4]
Set var_41 = Nothing
'006c3c2c    8d4dd0                  lea ecx, dword ptr [ebp-30]
'006c3c2f    51                      push ecx
'006c3c30    8d55d4                  lea edx, dword ptr [ebp-2c]
'006c3c33    52                      push edx
'006c3c34    8d45d8                  lea eax, dword ptr [ebp-28]
'006c3c37    50                      push eax
'006c3c38    8d4ddc                  lea ecx, dword ptr [ebp-24]
'006c3c3b    51                      push ecx
'006c3c3c    8d55e0                  lea edx, dword ptr [ebp-20]
'006c3c3f    52                      push edx
'006c3c40    6a05                    push 05

' *** Reference to "__vbaFreeStrList"
'006c3c42    ff155c124000            call dword ptr [0040125c]
'#Cleanup( 0, 0, -4500, -4504, 0)
'006c3c48    83c418                  add esp, 18
'006c3c4b    8d4dcc                  lea ecx, dword ptr [ebp-34]

' *** Reference to "__vbaFreeObj"
'006c3c4e    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_43)
'006c3c54    8b45e8                  mov eax, dword ptr [ebp-18]
'006c3c57    8b08                    mov ecx, dword ptr [eax]
'006c3c59    8d9528ffffff            lea edx, dword ptr [ebp+ffffff28]
'006c3c5f    52                      push edx
'006c3c60    50                      push eax
'006c3c61    ff5134                  call dword ptr [ecx+34]
'006c3c64    dbe2                    fnclex
'006c3c66    3bc7                    cmp eax, edi
'006c3c68    7d0e                    jge 6c3c78

If (var_41 < 0) Then
'006c3c6a    8b4de8                  mov ecx, dword ptr [ebp-18]
'006c3c6d    6a34                    push 34
'006c3c6f    6830314300              push 00433130
'006c3c74    51                      push ecx
'006c3c75    50                      push eax
'006c3c76    ffd6                    call esi
    
End If
'006c3c78    6639bd28ffffff          cmp word ptr [ebp+ffffff28], di
'006c3c7f    0f85ef130000            jne 6c5074

If (0 = 0) Then
'006c3c85    8b45e8                  mov eax, dword ptr [ebp-18]
'006c3c88    8b10                    mov edx, dword ptr [eax]
'006c3c8a    8d4dcc                  lea ecx, dword ptr [ebp-34]
'006c3c8d    51                      push ecx
'006c3c8e    50                      push eax
'006c3c8f    ff92b4000000            call dword ptr [edx+000000b4]
'006c3c95    dbe2                    fnclex
'006c3c97    3bc7                    cmp eax, edi
'006c3c99    7d11                    jge 6c3cac
    
    If (    var_41 < 0) Then
'006c3c9b    8b55e8                  mov edx, dword ptr [ebp-18]
'006c3c9e    68b4000000              push 000000b4
'006c3ca3    6830314300              push 00433130
'006c3ca8    52                      push edx
'006c3ca9    50                      push eax
'006c3caa    ffd6                    call esi
    
End If
'006c3cac    8b45cc                  mov eax, dword ptr [ebp-34]
'006c3caf    8b30                    mov esi, dword ptr [eax]
'006c3cb1    8d5dc8                  lea ebx, dword ptr [ebp-38]
'006c3cb4    53                      push ebx
'006c3cb5    83ec10                  sub esp, 10
'006c3cb8    8bdc                    mov ebx, esp
'006c3cba    ba08000000              mov edx, 00000008
'006c3cbf    8913                    mov dword ptr [ebx], edx
'006c3cc1    8b9560ffffff            mov edx, dword ptr [ebp+ffffff60]
'006c3cc7    895304                  mov dword ptr [ebx+04], edx
'006c3cca    b9acb24300              mov ecx, 0043b2ac
'006c3ccf    894b08                  mov dword ptr [ebx+08], ecx
'006c3cd2    8b8d68ffffff            mov ecx, dword ptr [ebp+ffffff68]
'006c3cd8    50                      push eax
'006c3cd9    898520ffffff            mov dword ptr [ebp+ffffff20], eax
'006c3cdf    894b0c                  mov dword ptr [ebx+0c], ecx
'006c3ce2    ff5630                  call dword ptr [esi+30]
'006c3ce5    dbe2                    fnclex
'006c3ce7    3bc7                    cmp eax, edi
'006c3ce9    7d15                    jge 6c3d00

If (var_43 < 0) Then
'006c3ceb    8b9520ffffff            mov edx, dword ptr [ebp+ffffff20]
'006c3cf1    6a30                    push 30
'006c3cf3    68d8304300              push 004330d8
'006c3cf8    52                      push edx
'006c3cf9    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c3cfa    ff1580104000            call dword ptr [00401080]
    
End If
'006c3d00    8b45c8                  mov eax, dword ptr [ebp-38]
'006c3d03    8945b4                  mov dword ptr [ebp-4c], eax
'006c3d06    8d45ac                  lea eax, dword ptr [ebp-54]
'006c3d09    50                      push eax
'006c3d0a    897dc8                  mov dword ptr [ebp-38], edi
'006c3d0d    c745ac09000000          mov dword ptr [ebp-54], 00000009

' *** Reference to "rtcIsNull"
'006c3d14    ff1540114000            call dword ptr [00401140]
'006c3d1a    898528ffffff            mov dword ptr [ebp+ffffff28], eax
'006c3d20    8b4508                  mov eax, dword ptr [ebp+08]
'006c3d23    8b08                    mov ecx, dword ptr [eax]
'006c3d25    50                      push eax
'006c3d26    ff9138030000            call dword ptr [ecx+00000338]
'006c3d2c    50                      push eax
'006c3d2d    8d55bc                  lea edx, dword ptr [ebp-44]
'006c3d30    52                      push edx

' *** Reference to "__vbaObjSet"
'006c3d31    ff15b4104000            call dword ptr [004010b4]
Set var_58 = Me
'006c3d37    8d55c4                  lea edx, dword ptr [ebp-3c]
'006c3d3a    8bf0                    mov esi, eax
'006c3d3c    8b45e8                  mov eax, dword ptr [ebp-18]
'006c3d3f    8b08                    mov ecx, dword ptr [eax]
'006c3d41    52                      push edx
'006c3d42    50                      push eax
'006c3d43    ff91b4000000            call dword ptr [ecx+000000b4]
'006c3d49    dbe2                    fnclex
'006c3d4b    3bc7                    cmp eax, edi
'006c3d4d    7d15                    jge 6c3d64

If (var_41 < 0) Then
'006c3d4f    8b4de8                  mov ecx, dword ptr [ebp-18]
'006c3d52    68b4000000              push 000000b4
'006c3d57    6830314300              push 00433130
'006c3d5c    51                      push ecx
'006c3d5d    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c3d5e    ff1580104000            call dword ptr [00401080]
    
End If
'006c3d64    8b45c4                  mov eax, dword ptr [ebp-3c]
'006c3d67    8b10                    mov edx, dword ptr [eax]
'006c3d69    8d5dc0                  lea ebx, dword ptr [ebp-40]
'006c3d6c    53                      push ebx
'006c3d6d    83ec10                  sub esp, 10
'006c3d70    8bdc                    mov ebx, esp
'006c3d72    b908000000              mov ecx, 00000008
'006c3d77    890b                    mov dword ptr [ebx], ecx
'006c3d79    8b8d50ffffff            mov ecx, dword ptr [ebp+ffffff50]
'006c3d7f    894b04                  mov dword ptr [ebx+04], ecx
'006c3d82    c78554ffffffacb24300    mov dword ptr [ebp+ffffff54], 0043b2ac
'006c3d8c    8b8d54ffffff            mov ecx, dword ptr [ebp+ffffff54]
'006c3d92    894b08                  mov dword ptr [ebx+08], ecx
'006c3d95    8b8d58ffffff            mov ecx, dword ptr [ebp+ffffff58]
'006c3d9b    50                      push eax
'006c3d9c    898514ffffff            mov dword ptr [ebp+ffffff14], eax
'006c3da2    894b0c                  mov dword ptr [ebx+0c], ecx
'006c3da5    ff5230                  call dword ptr [edx+30]
'006c3da8    dbe2                    fnclex
'006c3daa    3bc7                    cmp eax, edi
'006c3dac    7d15                    jge 6c3dc3

If (0 < 0) Then
'006c3dae    8b9514ffffff            mov edx, dword ptr [ebp+ffffff14]
'006c3db4    6a30                    push 30
'006c3db6    68d8304300              push 004330d8
'006c3dbb    52                      push edx
'006c3dbc    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c3dbd    ff1580104000            call dword ptr [00401080]
    
End If
'006c3dc3    8b45c0                  mov eax, dword ptr [ebp-40]
'006c3dc6    8945a4                  mov dword ptr [ebp-5c], eax
'006c3dc9    8d459c                  lea eax, dword ptr [ebp-64]
'006c3dcc    50                      push eax
'006c3dcd    8d4d8c                  lea ecx, dword ptr [ebp-74]
'006c3dd0    51                      push ecx
'006c3dd1    897dc0                  mov dword ptr [ebp-40], edi
'006c3dd4    c7459c09000000          mov dword ptr [ebp-64], 00000009

' *** Reference to "rtcTrimVar"
'006c3ddb    ff15e4104000            call dword ptr [004010e4]
'006c3de1    8d952cffffff            lea edx, dword ptr [ebp+ffffff2c]
'006c3de7    8d8d7cffffff            lea ecx, dword ptr [ebp+ffffff7c]
'006c3ded    c78534ffffffcc134300    mov dword ptr [ebp+ffffff34], 004313cc
'006c3df7    c7852cffffff08000000    mov dword ptr [ebp+ffffff2c], 00000008

' *** Reference to "__vbaVarDup"
'006c3e01    ff1598124000            call dword ptr [00401298]
var_59 = (vbNullChar)
'006c3e07    668b9528ffffff          mov dx, word ptr [ebp+ffffff28]
'006c3e0e    8d458c                  lea eax, dword ptr [ebp-74]
'006c3e11    50                      push eax
'006c3e12    8d8d7cffffff            lea ecx, dword ptr [ebp+ffffff7c]
'006c3e18    66899544ffffff          mov word ptr [ebp+ffffff44], dx
'006c3e1f    51                      push ecx
'006c3e20    8d953cffffff            lea edx, dword ptr [ebp+ffffff3c]
'006c3e26    52                      push edx
'006c3e27    8d856cffffff            lea eax, dword ptr [ebp+ffffff6c]
'006c3e2d    50                      push eax
'006c3e2e    c7853cffffff0b000000    mov dword ptr [ebp+ffffff3c], 0000000b

' *** Reference to "rtcImmediateIf"
'006c3e38    ff154c124000            call dword ptr [0040124c]
'006c3e3e    8b1e                    mov ebx, dword ptr [esi]
'006c3e40    8d8d6cffffff            lea ecx, dword ptr [ebp+ffffff6c]
'006c3e46    51                      push ecx
'006c3e47    8d55e4                  lea edx, dword ptr [ebp-1c]
'006c3e4a    52                      push edx

' *** Reference to "__vbaStrVarVal"
'006c3e4b    ff15fc114000            call dword ptr [004011fc]
'006c3e51    50                      push eax
'006c3e52    56                      push esi
'006c3e53    ff5354                  call dword ptr [ebx+54]
'006c3e56    dbe2                    fnclex
'006c3e58    3bc7                    cmp eax, edi
'006c3e5a    7d0f                    jge 6c3e6b
'006c3e5c    6a54                    push 54
'006c3e5e    683ce04300              push 0043e03c
'006c3e63    56                      push esi
'006c3e64    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c3e65    ff1580104000            call dword ptr [00401080]
'006c3e6b    8d4de4                  lea ecx, dword ptr [ebp-1c]

' *** Reference to "__vbaFreeStr"
'006c3e6e    ff150c134000            call dword ptr [0040130c]
'#Cleanup(var_40)
'006c3e74    8d45bc                  lea eax, dword ptr [ebp-44]
'006c3e77    50                      push eax
'006c3e78    8d4dc4                  lea ecx, dword ptr [ebp-3c]
'006c3e7b    51                      push ecx
'006c3e7c    8d55cc                  lea edx, dword ptr [ebp-34]
'006c3e7f    52                      push edx
'006c3e80    6a03                    push 03

' *** Reference to "__vbaFreeObjList"
'006c3e82    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_43, var_9, var_58)
'006c3e88    8d856cffffff            lea eax, dword ptr [ebp+ffffff6c]
'006c3e8e    50                      push eax
'006c3e8f    8d4d8c                  lea ecx, dword ptr [ebp-74]
'006c3e92    51                      push ecx
'006c3e93    8d957cffffff            lea edx, dword ptr [ebp+ffffff7c]
'006c3e99    52                      push edx
'006c3e9a    8d853cffffff            lea eax, dword ptr [ebp+ffffff3c]
'006c3ea0    50                      push eax
'006c3ea1    8d4d9c                  lea ecx, dword ptr [ebp-64]
'006c3ea4    51                      push ecx
'006c3ea5    8d55ac                  lea edx, dword ptr [ebp-54]
'006c3ea8    52                      push edx
'006c3ea9    6a06                    push 06

' *** Reference to "__vbaFreeVarList"
'006c3eab    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_50, var_51, var_90, var_59, var_52, var_94)
'006c3eb1    8b45e8                  mov eax, dword ptr [ebp-18]
'006c3eb4    8b08                    mov ecx, dword ptr [eax]
'006c3eb6    83c42c                  add esp, 2c
'006c3eb9    8d55cc                  lea edx, dword ptr [ebp-34]
'006c3ebc    52                      push edx
'006c3ebd    50                      push eax
'006c3ebe    ff91b4000000            call dword ptr [ecx+000000b4]
'006c3ec4    dbe2                    fnclex
'006c3ec6    3bc7                    cmp eax, edi
'006c3ec8    7d15                    jge 6c3edf

If (var_41 < 0) Then
'006c3eca    8b4de8                  mov ecx, dword ptr [ebp-18]
'006c3ecd    68b4000000              push 000000b4
'006c3ed2    6830314300              push 00433130
'006c3ed7    51                      push ecx
'006c3ed8    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c3ed9    ff1580104000            call dword ptr [00401080]
    
End If
'006c3edf    8b45cc                  mov eax, dword ptr [ebp-34]
'006c3ee2    8b30                    mov esi, dword ptr [eax]
'006c3ee4    8d5dc8                  lea ebx, dword ptr [ebp-38]
'006c3ee7    53                      push ebx
'006c3ee8    83ec10                  sub esp, 10
'006c3eeb    8bdc                    mov ebx, esp
'006c3eed    ba08000000              mov edx, 00000008
'006c3ef2    8913                    mov dword ptr [ebx], edx
'006c3ef4    8b9560ffffff            mov edx, dword ptr [ebp+ffffff60]
'006c3efa    895304                  mov dword ptr [ebx+04], edx
'006c3efd    b928354400              mov ecx, 00443528
'006c3f02    894b08                  mov dword ptr [ebx+08], ecx
'006c3f05    8b8d68ffffff            mov ecx, dword ptr [ebp+ffffff68]
'006c3f0b    50                      push eax
'006c3f0c    898520ffffff            mov dword ptr [ebp+ffffff20], eax
'006c3f12    894b0c                  mov dword ptr [ebx+0c], ecx
'006c3f15    ff5630                  call dword ptr [esi+30]
'006c3f18    dbe2                    fnclex
'006c3f1a    3bc7                    cmp eax, edi
'006c3f1c    7d15                    jge 6c3f33

If (var_43 < 0) Then
'006c3f1e    8b9520ffffff            mov edx, dword ptr [ebp+ffffff20]
'006c3f24    6a30                    push 30
'006c3f26    68d8304300              push 004330d8
'006c3f2b    52                      push edx
'006c3f2c    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c3f2d    ff1580104000            call dword ptr [00401080]
    
End If
'006c3f33    8b45c8                  mov eax, dword ptr [ebp-38]
'006c3f36    8945b4                  mov dword ptr [ebp-4c], eax
'006c3f39    8d45ac                  lea eax, dword ptr [ebp-54]
'006c3f3c    50                      push eax
'006c3f3d    897dc8                  mov dword ptr [ebp-38], edi
'006c3f40    c745ac09000000          mov dword ptr [ebp-54], 00000009

' *** Reference to "rtcIsNull"
'006c3f47    ff1540114000            call dword ptr [00401140]
'006c3f4d    898528ffffff            mov dword ptr [ebp+ffffff28], eax
'006c3f53    8b4508                  mov eax, dword ptr [ebp+08]
'006c3f56    8b08                    mov ecx, dword ptr [eax]
'006c3f58    50                      push eax
'006c3f59    ff9134030000            call dword ptr [ecx+00000334]
'006c3f5f    50                      push eax
'006c3f60    8d55bc                  lea edx, dword ptr [ebp-44]
'006c3f63    52                      push edx

' *** Reference to "__vbaObjSet"
'006c3f64    ff15b4104000            call dword ptr [004010b4]
Set var_58 = Me
'006c3f6a    8d55c4                  lea edx, dword ptr [ebp-3c]
'006c3f6d    8bf0                    mov esi, eax
'006c3f6f    8b45e8                  mov eax, dword ptr [ebp-18]
'006c3f72    8b08                    mov ecx, dword ptr [eax]
'006c3f74    52                      push edx
'006c3f75    50                      push eax
'006c3f76    ff91b4000000            call dword ptr [ecx+000000b4]
'006c3f7c    dbe2                    fnclex
'006c3f7e    3bc7                    cmp eax, edi
'006c3f80    7d15                    jge 6c3f97

If (var_41 < 0) Then
'006c3f82    8b4de8                  mov ecx, dword ptr [ebp-18]
'006c3f85    68b4000000              push 000000b4
'006c3f8a    6830314300              push 00433130
'006c3f8f    51                      push ecx
'006c3f90    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c3f91    ff1580104000            call dword ptr [00401080]
    
End If
'006c3f97    8b45c4                  mov eax, dword ptr [ebp-3c]
'006c3f9a    8b10                    mov edx, dword ptr [eax]
'006c3f9c    8d5dc0                  lea ebx, dword ptr [ebp-40]
'006c3f9f    53                      push ebx
'006c3fa0    83ec10                  sub esp, 10
'006c3fa3    8bdc                    mov ebx, esp
'006c3fa5    b908000000              mov ecx, 00000008
'006c3faa    890b                    mov dword ptr [ebx], ecx
'006c3fac    8b8d50ffffff            mov ecx, dword ptr [ebp+ffffff50]
'006c3fb2    894b04                  mov dword ptr [ebx+04], ecx
'006c3fb5    c78554ffffff28354400    mov dword ptr [ebp+ffffff54], 00443528
'006c3fbf    8b8d54ffffff            mov ecx, dword ptr [ebp+ffffff54]
'006c3fc5    894b08                  mov dword ptr [ebx+08], ecx
'006c3fc8    8b8d58ffffff            mov ecx, dword ptr [ebp+ffffff58]
'006c3fce    50                      push eax
'006c3fcf    898514ffffff            mov dword ptr [ebp+ffffff14], eax
'006c3fd5    894b0c                  mov dword ptr [ebx+0c], ecx
'006c3fd8    ff5230                  call dword ptr [edx+30]
'006c3fdb    dbe2                    fnclex
'006c3fdd    3bc7                    cmp eax, edi
'006c3fdf    7d15                    jge 6c3ff6

If (0 < 0) Then
'006c3fe1    8b9514ffffff            mov edx, dword ptr [ebp+ffffff14]
'006c3fe7    6a30                    push 30
'006c3fe9    68d8304300              push 004330d8
'006c3fee    52                      push edx
'006c3fef    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c3ff0    ff1580104000            call dword ptr [00401080]
    
End If
'006c3ff6    8b45c0                  mov eax, dword ptr [ebp-40]
'006c3ff9    8945a4                  mov dword ptr [ebp-5c], eax
'006c3ffc    8d459c                  lea eax, dword ptr [ebp-64]
'006c3fff    50                      push eax
'006c4000    8d4d8c                  lea ecx, dword ptr [ebp-74]
'006c4003    51                      push ecx
'006c4004    897dc0                  mov dword ptr [ebp-40], edi
'006c4007    c7459c09000000          mov dword ptr [ebp-64], 00000009

' *** Reference to "rtcTrimVar"
'006c400e    ff15e4104000            call dword ptr [004010e4]
'006c4014    8d952cffffff            lea edx, dword ptr [ebp+ffffff2c]
'006c401a    8d8d7cffffff            lea ecx, dword ptr [ebp+ffffff7c]
'006c4020    c78534ffffffcc134300    mov dword ptr [ebp+ffffff34], 004313cc
'006c402a    c7852cffffff08000000    mov dword ptr [ebp+ffffff2c], 00000008

' *** Reference to "__vbaVarDup"
'006c4034    ff1598124000            call dword ptr [00401298]
var_59 = (vbNullChar)
'006c403a    668b9528ffffff          mov dx, word ptr [ebp+ffffff28]
'006c4041    8d458c                  lea eax, dword ptr [ebp-74]
'006c4044    50                      push eax
'006c4045    8d8d7cffffff            lea ecx, dword ptr [ebp+ffffff7c]
'006c404b    66899544ffffff          mov word ptr [ebp+ffffff44], dx
'006c4052    51                      push ecx
'006c4053    8d953cffffff            lea edx, dword ptr [ebp+ffffff3c]
'006c4059    52                      push edx
'006c405a    8d856cffffff            lea eax, dword ptr [ebp+ffffff6c]
'006c4060    50                      push eax
'006c4061    c7853cffffff0b000000    mov dword ptr [ebp+ffffff3c], 0000000b

' *** Reference to "rtcImmediateIf"
'006c406b    ff154c124000            call dword ptr [0040124c]
'006c4071    8b1e                    mov ebx, dword ptr [esi]
'006c4073    8d8d6cffffff            lea ecx, dword ptr [ebp+ffffff6c]
'006c4079    51                      push ecx
'006c407a    8d55e4                  lea edx, dword ptr [ebp-1c]
'006c407d    52                      push edx

' *** Reference to "__vbaStrVarVal"
'006c407e    ff15fc114000            call dword ptr [004011fc]
'006c4084    50                      push eax
'006c4085    56                      push esi
'006c4086    ff5354                  call dword ptr [ebx+54]
'006c4089    dbe2                    fnclex
'006c408b    3bc7                    cmp eax, edi
'006c408d    7d0f                    jge 6c409e
'006c408f    6a54                    push 54
'006c4091    683ce04300              push 0043e03c
'006c4096    56                      push esi
'006c4097    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c4098    ff1580104000            call dword ptr [00401080]
'006c409e    8d4de4                  lea ecx, dword ptr [ebp-1c]

' *** Reference to "__vbaFreeStr"
'006c40a1    ff150c134000            call dword ptr [0040130c]
'#Cleanup(var_40)
'006c40a7    8d45bc                  lea eax, dword ptr [ebp-44]
'006c40aa    50                      push eax
'006c40ab    8d4dc4                  lea ecx, dword ptr [ebp-3c]
'006c40ae    51                      push ecx
'006c40af    8d55cc                  lea edx, dword ptr [ebp-34]
'006c40b2    52                      push edx
'006c40b3    6a03                    push 03

' *** Reference to "__vbaFreeObjList"
'006c40b5    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_43, var_9, var_58)
'006c40bb    8d856cffffff            lea eax, dword ptr [ebp+ffffff6c]
'006c40c1    50                      push eax
'006c40c2    8d4d8c                  lea ecx, dword ptr [ebp-74]
'006c40c5    51                      push ecx
'006c40c6    8d957cffffff            lea edx, dword ptr [ebp+ffffff7c]
'006c40cc    52                      push edx
'006c40cd    8d853cffffff            lea eax, dword ptr [ebp+ffffff3c]
'006c40d3    50                      push eax
'006c40d4    8d4d9c                  lea ecx, dword ptr [ebp-64]
'006c40d7    51                      push ecx
'006c40d8    8d55ac                  lea edx, dword ptr [ebp-54]
'006c40db    52                      push edx
'006c40dc    6a06                    push 06

' *** Reference to "__vbaFreeVarList"
'006c40de    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_50, var_51, var_90, var_59, var_52, var_94)
'006c40e4    8b45e8                  mov eax, dword ptr [ebp-18]
'006c40e7    8b08                    mov ecx, dword ptr [eax]
'006c40e9    83c42c                  add esp, 2c
'006c40ec    8d55cc                  lea edx, dword ptr [ebp-34]
'006c40ef    52                      push edx
'006c40f0    50                      push eax
'006c40f1    ff91b4000000            call dword ptr [ecx+000000b4]
'006c40f7    dbe2                    fnclex
'006c40f9    3bc7                    cmp eax, edi
'006c40fb    7d15                    jge 6c4112

If (var_41 < 0) Then
'006c40fd    8b4de8                  mov ecx, dword ptr [ebp-18]
'006c4100    68b4000000              push 000000b4
'006c4105    6830314300              push 00433130
'006c410a    51                      push ecx
'006c410b    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c410c    ff1580104000            call dword ptr [00401080]
    
End If
'006c4112    8b45cc                  mov eax, dword ptr [ebp-34]
'006c4115    8b30                    mov esi, dword ptr [eax]
'006c4117    8d5dc8                  lea ebx, dword ptr [ebp-38]
'006c411a    53                      push ebx
'006c411b    83ec10                  sub esp, 10
'006c411e    8bdc                    mov ebx, esp
'006c4120    ba08000000              mov edx, 00000008
'006c4125    8913                    mov dword ptr [ebx], edx
'006c4127    8b9560ffffff            mov edx, dword ptr [ebp+ffffff60]
'006c412d    895304                  mov dword ptr [ebx+04], edx
'006c4130    b974a64300              mov ecx, 0043a674
'006c4135    894b08                  mov dword ptr [ebx+08], ecx
'006c4138    8b8d68ffffff            mov ecx, dword ptr [ebp+ffffff68]
'006c413e    50                      push eax
'006c413f    898520ffffff            mov dword ptr [ebp+ffffff20], eax
'006c4145    894b0c                  mov dword ptr [ebx+0c], ecx
'006c4148    ff5630                  call dword ptr [esi+30]
'006c414b    dbe2                    fnclex
'006c414d    3bc7                    cmp eax, edi
'006c414f    7d15                    jge 6c4166

If (var_43 < 0) Then
'006c4151    8b9520ffffff            mov edx, dword ptr [ebp+ffffff20]
'006c4157    6a30                    push 30
'006c4159    68d8304300              push 004330d8
'006c415e    52                      push edx
'006c415f    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c4160    ff1580104000            call dword ptr [00401080]
    
End If
'006c4166    8b45c8                  mov eax, dword ptr [ebp-38]
'006c4169    8945b4                  mov dword ptr [ebp-4c], eax
'006c416c    8d45ac                  lea eax, dword ptr [ebp-54]
'006c416f    50                      push eax
'006c4170    897dc8                  mov dword ptr [ebp-38], edi
'006c4173    c745ac09000000          mov dword ptr [ebp-54], 00000009

' *** Reference to "rtcIsNull"
'006c417a    ff1540114000            call dword ptr [00401140]
'006c4180    898528ffffff            mov dword ptr [ebp+ffffff28], eax
'006c4186    8b4508                  mov eax, dword ptr [ebp+08]
'006c4189    8b08                    mov ecx, dword ptr [eax]
'006c418b    50                      push eax
'006c418c    ff912c030000            call dword ptr [ecx+0000032c]
'006c4192    50                      push eax
'006c4193    8d55bc                  lea edx, dword ptr [ebp-44]
'006c4196    52                      push edx

' *** Reference to "__vbaObjSet"
'006c4197    ff15b4104000            call dword ptr [004010b4]
Set var_58 = Me
'006c419d    8d55c4                  lea edx, dword ptr [ebp-3c]
'006c41a0    8bf0                    mov esi, eax
'006c41a2    8b45e8                  mov eax, dword ptr [ebp-18]
'006c41a5    8b08                    mov ecx, dword ptr [eax]
'006c41a7    52                      push edx
'006c41a8    50                      push eax
'006c41a9    ff91b4000000            call dword ptr [ecx+000000b4]
'006c41af    dbe2                    fnclex
'006c41b1    3bc7                    cmp eax, edi
'006c41b3    7d15                    jge 6c41ca

If (var_41 < 0) Then
'006c41b5    8b4de8                  mov ecx, dword ptr [ebp-18]
'006c41b8    68b4000000              push 000000b4
'006c41bd    6830314300              push 00433130
'006c41c2    51                      push ecx
'006c41c3    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c41c4    ff1580104000            call dword ptr [00401080]
    
End If
'006c41ca    8b45c4                  mov eax, dword ptr [ebp-3c]
'006c41cd    8b10                    mov edx, dword ptr [eax]
'006c41cf    8d5dc0                  lea ebx, dword ptr [ebp-40]
'006c41d2    53                      push ebx
'006c41d3    83ec10                  sub esp, 10
'006c41d6    8bdc                    mov ebx, esp
'006c41d8    b908000000              mov ecx, 00000008
'006c41dd    890b                    mov dword ptr [ebx], ecx
'006c41df    8b8d50ffffff            mov ecx, dword ptr [ebp+ffffff50]
'006c41e5    894b04                  mov dword ptr [ebx+04], ecx
'006c41e8    c78554ffffff74a64300    mov dword ptr [ebp+ffffff54], 0043a674
'006c41f2    8b8d54ffffff            mov ecx, dword ptr [ebp+ffffff54]
'006c41f8    894b08                  mov dword ptr [ebx+08], ecx
'006c41fb    8b8d58ffffff            mov ecx, dword ptr [ebp+ffffff58]
'006c4201    50                      push eax
'006c4202    898514ffffff            mov dword ptr [ebp+ffffff14], eax
'006c4208    894b0c                  mov dword ptr [ebx+0c], ecx
'006c420b    ff5230                  call dword ptr [edx+30]
'006c420e    dbe2                    fnclex
'006c4210    3bc7                    cmp eax, edi
'006c4212    7d15                    jge 6c4229

If (0 < 0) Then
'006c4214    8b9514ffffff            mov edx, dword ptr [ebp+ffffff14]
'006c421a    6a30                    push 30
'006c421c    68d8304300              push 004330d8
'006c4221    52                      push edx
'006c4222    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c4223    ff1580104000            call dword ptr [00401080]
    
End If
'006c4229    8b45c0                  mov eax, dword ptr [ebp-40]
'006c422c    8945a4                  mov dword ptr [ebp-5c], eax
'006c422f    8d459c                  lea eax, dword ptr [ebp-64]
'006c4232    50                      push eax
'006c4233    8d4d8c                  lea ecx, dword ptr [ebp-74]
'006c4236    51                      push ecx
'006c4237    897dc0                  mov dword ptr [ebp-40], edi
'006c423a    c7459c09000000          mov dword ptr [ebp-64], 00000009

' *** Reference to "rtcTrimVar"
'006c4241    ff15e4104000            call dword ptr [004010e4]
'006c4247    8d952cffffff            lea edx, dword ptr [ebp+ffffff2c]
'006c424d    8d8d7cffffff            lea ecx, dword ptr [ebp+ffffff7c]
'006c4253    c78534ffffffcc134300    mov dword ptr [ebp+ffffff34], 004313cc
'006c425d    c7852cffffff08000000    mov dword ptr [ebp+ffffff2c], 00000008

' *** Reference to "__vbaVarDup"
'006c4267    ff1598124000            call dword ptr [00401298]
var_59 = (vbNullChar)
'006c426d    668b9528ffffff          mov dx, word ptr [ebp+ffffff28]
'006c4274    8d458c                  lea eax, dword ptr [ebp-74]
'006c4277    50                      push eax
'006c4278    8d8d7cffffff            lea ecx, dword ptr [ebp+ffffff7c]
'006c427e    66899544ffffff          mov word ptr [ebp+ffffff44], dx
'006c4285    51                      push ecx
'006c4286    8d953cffffff            lea edx, dword ptr [ebp+ffffff3c]
'006c428c    52                      push edx
'006c428d    8d856cffffff            lea eax, dword ptr [ebp+ffffff6c]
'006c4293    50                      push eax
'006c4294    c7853cffffff0b000000    mov dword ptr [ebp+ffffff3c], 0000000b

' *** Reference to "rtcImmediateIf"
'006c429e    ff154c124000            call dword ptr [0040124c]
'006c42a4    8b1e                    mov ebx, dword ptr [esi]
'006c42a6    8d8d6cffffff            lea ecx, dword ptr [ebp+ffffff6c]
'006c42ac    51                      push ecx
'006c42ad    8d55e4                  lea edx, dword ptr [ebp-1c]
'006c42b0    52                      push edx

' *** Reference to "__vbaStrVarVal"
'006c42b1    ff15fc114000            call dword ptr [004011fc]
'006c42b7    50                      push eax
'006c42b8    56                      push esi
'006c42b9    ff5354                  call dword ptr [ebx+54]
'006c42bc    dbe2                    fnclex
'006c42be    3bc7                    cmp eax, edi
'006c42c0    7d0f                    jge 6c42d1
'006c42c2    6a54                    push 54
'006c42c4    683ce04300              push 0043e03c
'006c42c9    56                      push esi
'006c42ca    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c42cb    ff1580104000            call dword ptr [00401080]
'006c42d1    8d4de4                  lea ecx, dword ptr [ebp-1c]

' *** Reference to "__vbaFreeStr"
'006c42d4    ff150c134000            call dword ptr [0040130c]
'#Cleanup(var_40)
'006c42da    8d45bc                  lea eax, dword ptr [ebp-44]
'006c42dd    50                      push eax
'006c42de    8d4dc4                  lea ecx, dword ptr [ebp-3c]
'006c42e1    51                      push ecx
'006c42e2    8d55cc                  lea edx, dword ptr [ebp-34]
'006c42e5    52                      push edx
'006c42e6    6a03                    push 03

' *** Reference to "__vbaFreeObjList"
'006c42e8    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_43, var_9, var_58)
'006c42ee    8d856cffffff            lea eax, dword ptr [ebp+ffffff6c]
'006c42f4    50                      push eax
'006c42f5    8d4d8c                  lea ecx, dword ptr [ebp-74]
'006c42f8    51                      push ecx
'006c42f9    8d957cffffff            lea edx, dword ptr [ebp+ffffff7c]
'006c42ff    52                      push edx
'006c4300    8d853cffffff            lea eax, dword ptr [ebp+ffffff3c]
'006c4306    50                      push eax
'006c4307    8d4d9c                  lea ecx, dword ptr [ebp-64]
'006c430a    51                      push ecx
'006c430b    8d55ac                  lea edx, dword ptr [ebp-54]
'006c430e    52                      push edx
'006c430f    6a06                    push 06

' *** Reference to "__vbaFreeVarList"
'006c4311    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_50, var_51, var_90, var_59, var_52, var_94)
'006c4317    8b45e8                  mov eax, dword ptr [ebp-18]
'006c431a    8b08                    mov ecx, dword ptr [eax]
'006c431c    83c42c                  add esp, 2c
'006c431f    8d55cc                  lea edx, dword ptr [ebp-34]
'006c4322    52                      push edx
'006c4323    50                      push eax
'006c4324    ff91b4000000            call dword ptr [ecx+000000b4]
'006c432a    dbe2                    fnclex
'006c432c    3bc7                    cmp eax, edi
'006c432e    7d15                    jge 6c4345

If (var_41 < 0) Then
'006c4330    8b4de8                  mov ecx, dword ptr [ebp-18]
'006c4333    68b4000000              push 000000b4
'006c4338    6830314300              push 00433130
'006c433d    51                      push ecx
'006c433e    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c433f    ff1580104000            call dword ptr [00401080]
    
End If
'006c4345    8b45cc                  mov eax, dword ptr [ebp-34]
'006c4348    8b30                    mov esi, dword ptr [eax]
'006c434a    8d5dc8                  lea ebx, dword ptr [ebp-38]
'006c434d    53                      push ebx
'006c434e    83ec10                  sub esp, 10
'006c4351    8bdc                    mov ebx, esp
'006c4353    ba08000000              mov edx, 00000008
'006c4358    8913                    mov dword ptr [ebx], edx
'006c435a    8b9560ffffff            mov edx, dword ptr [ebp+ffffff60]
'006c4360    895304                  mov dword ptr [ebx+04], edx
'006c4363    b9bcb24300              mov ecx, 0043b2bc
'006c4368    894b08                  mov dword ptr [ebx+08], ecx
'006c436b    8b8d68ffffff            mov ecx, dword ptr [ebp+ffffff68]
'006c4371    50                      push eax
'006c4372    898520ffffff            mov dword ptr [ebp+ffffff20], eax
'006c4378    894b0c                  mov dword ptr [ebx+0c], ecx
'006c437b    ff5630                  call dword ptr [esi+30]
'006c437e    dbe2                    fnclex
'006c4380    3bc7                    cmp eax, edi
'006c4382    7d15                    jge 6c4399

If (var_43 < 0) Then
'006c4384    8b9520ffffff            mov edx, dword ptr [ebp+ffffff20]
'006c438a    6a30                    push 30
'006c438c    68d8304300              push 004330d8
'006c4391    52                      push edx
'006c4392    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c4393    ff1580104000            call dword ptr [00401080]
    
End If
'006c4399    8b45c8                  mov eax, dword ptr [ebp-38]
'006c439c    8945b4                  mov dword ptr [ebp-4c], eax
'006c439f    8d45ac                  lea eax, dword ptr [ebp-54]
'006c43a2    50                      push eax
'006c43a3    897dc8                  mov dword ptr [ebp-38], edi
'006c43a6    c745ac09000000          mov dword ptr [ebp-54], 00000009

' *** Reference to "rtcIsNull"
'006c43ad    ff1540114000            call dword ptr [00401140]
'006c43b3    898528ffffff            mov dword ptr [ebp+ffffff28], eax
'006c43b9    8b4508                  mov eax, dword ptr [ebp+08]
'006c43bc    8b08                    mov ecx, dword ptr [eax]
'006c43be    50                      push eax
'006c43bf    ff9130030000            call dword ptr [ecx+00000330]
'006c43c5    50                      push eax
'006c43c6    8d55bc                  lea edx, dword ptr [ebp-44]
'006c43c9    52                      push edx

' *** Reference to "__vbaObjSet"
'006c43ca    ff15b4104000            call dword ptr [004010b4]
Set var_58 = Me
'006c43d0    8d55c4                  lea edx, dword ptr [ebp-3c]
'006c43d3    8bf0                    mov esi, eax
'006c43d5    8b45e8                  mov eax, dword ptr [ebp-18]
'006c43d8    8b08                    mov ecx, dword ptr [eax]
'006c43da    52                      push edx
'006c43db    50                      push eax
'006c43dc    ff91b4000000            call dword ptr [ecx+000000b4]
'006c43e2    dbe2                    fnclex
'006c43e4    3bc7                    cmp eax, edi
'006c43e6    7d15                    jge 6c43fd

If (var_41 < 0) Then
'006c43e8    8b4de8                  mov ecx, dword ptr [ebp-18]
'006c43eb    68b4000000              push 000000b4
'006c43f0    6830314300              push 00433130
'006c43f5    51                      push ecx
'006c43f6    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c43f7    ff1580104000            call dword ptr [00401080]
    
End If
'006c43fd    8b45c4                  mov eax, dword ptr [ebp-3c]
'006c4400    8b10                    mov edx, dword ptr [eax]
'006c4402    8d5dc0                  lea ebx, dword ptr [ebp-40]
'006c4405    53                      push ebx
'006c4406    83ec10                  sub esp, 10
'006c4409    8bdc                    mov ebx, esp
'006c440b    b908000000              mov ecx, 00000008
'006c4410    890b                    mov dword ptr [ebx], ecx
'006c4412    8b8d50ffffff            mov ecx, dword ptr [ebp+ffffff50]
'006c4418    894b04                  mov dword ptr [ebx+04], ecx
'006c441b    c78554ffffffbcb24300    mov dword ptr [ebp+ffffff54], 0043b2bc
'006c4425    8b8d54ffffff            mov ecx, dword ptr [ebp+ffffff54]
'006c442b    894b08                  mov dword ptr [ebx+08], ecx
'006c442e    8b8d58ffffff            mov ecx, dword ptr [ebp+ffffff58]
'006c4434    50                      push eax
'006c4435    898514ffffff            mov dword ptr [ebp+ffffff14], eax
'006c443b    894b0c                  mov dword ptr [ebx+0c], ecx
'006c443e    ff5230                  call dword ptr [edx+30]
'006c4441    dbe2                    fnclex
'006c4443    3bc7                    cmp eax, edi
'006c4445    7d15                    jge 6c445c

If (0 < 0) Then
'006c4447    8b9514ffffff            mov edx, dword ptr [ebp+ffffff14]
'006c444d    6a30                    push 30
'006c444f    68d8304300              push 004330d8
'006c4454    52                      push edx
'006c4455    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c4456    ff1580104000            call dword ptr [00401080]
    
End If
'006c445c    8b45c0                  mov eax, dword ptr [ebp-40]
'006c445f    8945a4                  mov dword ptr [ebp-5c], eax
'006c4462    8d459c                  lea eax, dword ptr [ebp-64]
'006c4465    50                      push eax
'006c4466    8d4d8c                  lea ecx, dword ptr [ebp-74]
'006c4469    51                      push ecx
'006c446a    897dc0                  mov dword ptr [ebp-40], edi
'006c446d    c7459c09000000          mov dword ptr [ebp-64], 00000009

' *** Reference to "rtcTrimVar"
'006c4474    ff15e4104000            call dword ptr [004010e4]
'006c447a    8d952cffffff            lea edx, dword ptr [ebp+ffffff2c]
'006c4480    8d8d7cffffff            lea ecx, dword ptr [ebp+ffffff7c]
'006c4486    c78534ffffffcc134300    mov dword ptr [ebp+ffffff34], 004313cc
'006c4490    c7852cffffff08000000    mov dword ptr [ebp+ffffff2c], 00000008

' *** Reference to "__vbaVarDup"
'006c449a    ff1598124000            call dword ptr [00401298]
var_59 = (vbNullChar)
'006c44a0    668b9528ffffff          mov dx, word ptr [ebp+ffffff28]
'006c44a7    8d458c                  lea eax, dword ptr [ebp-74]
'006c44aa    50                      push eax
'006c44ab    8d8d7cffffff            lea ecx, dword ptr [ebp+ffffff7c]
'006c44b1    66899544ffffff          mov word ptr [ebp+ffffff44], dx
'006c44b8    51                      push ecx
'006c44b9    8d953cffffff            lea edx, dword ptr [ebp+ffffff3c]
'006c44bf    52                      push edx
'006c44c0    8d856cffffff            lea eax, dword ptr [ebp+ffffff6c]
'006c44c6    50                      push eax
'006c44c7    c7853cffffff0b000000    mov dword ptr [ebp+ffffff3c], 0000000b

' *** Reference to "rtcImmediateIf"
'006c44d1    ff154c124000            call dword ptr [0040124c]
'006c44d7    8b1e                    mov ebx, dword ptr [esi]
'006c44d9    8d8d6cffffff            lea ecx, dword ptr [ebp+ffffff6c]
'006c44df    51                      push ecx
'006c44e0    8d55e4                  lea edx, dword ptr [ebp-1c]
'006c44e3    52                      push edx

' *** Reference to "__vbaStrVarVal"
'006c44e4    ff15fc114000            call dword ptr [004011fc]
'006c44ea    50                      push eax
'006c44eb    56                      push esi
'006c44ec    ff5354                  call dword ptr [ebx+54]
'006c44ef    dbe2                    fnclex
'006c44f1    3bc7                    cmp eax, edi
'006c44f3    7d0f                    jge 6c4504
'006c44f5    6a54                    push 54
'006c44f7    683ce04300              push 0043e03c
'006c44fc    56                      push esi
'006c44fd    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c44fe    ff1580104000            call dword ptr [00401080]
'006c4504    8d4de4                  lea ecx, dword ptr [ebp-1c]

' *** Reference to "__vbaFreeStr"
'006c4507    ff150c134000            call dword ptr [0040130c]
'#Cleanup(var_40)
'006c450d    8d45bc                  lea eax, dword ptr [ebp-44]
'006c4510    50                      push eax
'006c4511    8d4dc4                  lea ecx, dword ptr [ebp-3c]
'006c4514    51                      push ecx
'006c4515    8d55cc                  lea edx, dword ptr [ebp-34]
'006c4518    52                      push edx
'006c4519    6a03                    push 03

' *** Reference to "__vbaFreeObjList"
'006c451b    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_43, var_9, var_58)
'006c4521    8d856cffffff            lea eax, dword ptr [ebp+ffffff6c]
'006c4527    50                      push eax
'006c4528    8d4d8c                  lea ecx, dword ptr [ebp-74]
'006c452b    51                      push ecx
'006c452c    8d957cffffff            lea edx, dword ptr [ebp+ffffff7c]
'006c4532    52                      push edx
'006c4533    8d853cffffff            lea eax, dword ptr [ebp+ffffff3c]
'006c4539    50                      push eax
'006c453a    8d4d9c                  lea ecx, dword ptr [ebp-64]
'006c453d    51                      push ecx
'006c453e    8d55ac                  lea edx, dword ptr [ebp-54]
'006c4541    52                      push edx
'006c4542    6a06                    push 06

' *** Reference to "__vbaFreeVarList"
'006c4544    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_50, var_51, var_90, var_59, var_52, var_94)
'006c454a    8b45e8                  mov eax, dword ptr [ebp-18]
'006c454d    8b08                    mov ecx, dword ptr [eax]
'006c454f    83c42c                  add esp, 2c
'006c4552    8d55cc                  lea edx, dword ptr [ebp-34]
'006c4555    52                      push edx
'006c4556    50                      push eax
'006c4557    ff91b4000000            call dword ptr [ecx+000000b4]
'006c455d    dbe2                    fnclex
'006c455f    3bc7                    cmp eax, edi
'006c4561    7d15                    jge 6c4578

If (var_41 < 0) Then
'006c4563    8b4de8                  mov ecx, dword ptr [ebp-18]
'006c4566    68b4000000              push 000000b4
'006c456b    6830314300              push 00433130
'006c4570    51                      push ecx
'006c4571    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c4572    ff1580104000            call dword ptr [00401080]
    
End If
'006c4578    8b45cc                  mov eax, dword ptr [ebp-34]
'006c457b    8b30                    mov esi, dword ptr [eax]
'006c457d    8d5dc8                  lea ebx, dword ptr [ebp-38]
'006c4580    53                      push ebx
'006c4581    83ec10                  sub esp, 10
'006c4584    8bdc                    mov ebx, esp
'006c4586    ba08000000              mov edx, 00000008
'006c458b    8913                    mov dword ptr [ebx], edx
'006c458d    8b9560ffffff            mov edx, dword ptr [ebp+ffffff60]
'006c4593    895304                  mov dword ptr [ebx+04], edx
'006c4596    b9a0b84300              mov ecx, 0043b8a0
'006c459b    894b08                  mov dword ptr [ebx+08], ecx
'006c459e    8b8d68ffffff            mov ecx, dword ptr [ebp+ffffff68]
'006c45a4    50                      push eax
'006c45a5    898520ffffff            mov dword ptr [ebp+ffffff20], eax
'006c45ab    894b0c                  mov dword ptr [ebx+0c], ecx
'006c45ae    ff5630                  call dword ptr [esi+30]
'006c45b1    dbe2                    fnclex
'006c45b3    3bc7                    cmp eax, edi
'006c45b5    7d15                    jge 6c45cc

If (var_43 < 0) Then
'006c45b7    8b9520ffffff            mov edx, dword ptr [ebp+ffffff20]
'006c45bd    6a30                    push 30
'006c45bf    68d8304300              push 004330d8
'006c45c4    52                      push edx
'006c45c5    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c45c6    ff1580104000            call dword ptr [00401080]
    
End If
'006c45cc    8b45c8                  mov eax, dword ptr [ebp-38]
'006c45cf    8945b4                  mov dword ptr [ebp-4c], eax
'006c45d2    8d45ac                  lea eax, dword ptr [ebp-54]
'006c45d5    50                      push eax
'006c45d6    897dc8                  mov dword ptr [ebp-38], edi
'006c45d9    c745ac09000000          mov dword ptr [ebp-54], 00000009

' *** Reference to "rtcIsNull"
'006c45e0    ff1540114000            call dword ptr [00401140]
'006c45e6    898528ffffff            mov dword ptr [ebp+ffffff28], eax
'006c45ec    8b4508                  mov eax, dword ptr [ebp+08]
'006c45ef    8b08                    mov ecx, dword ptr [eax]
'006c45f1    50                      push eax
'006c45f2    ff910c030000            call dword ptr [ecx+0000030c]
'006c45f8    50                      push eax
'006c45f9    8d55bc                  lea edx, dword ptr [ebp-44]
'006c45fc    52                      push edx

' *** Reference to "__vbaObjSet"
'006c45fd    ff15b4104000            call dword ptr [004010b4]
Set var_58 = Me
'006c4603    8d55c4                  lea edx, dword ptr [ebp-3c]
'006c4606    8bf0                    mov esi, eax
'006c4608    8b45e8                  mov eax, dword ptr [ebp-18]
'006c460b    8b08                    mov ecx, dword ptr [eax]
'006c460d    52                      push edx
'006c460e    50                      push eax
'006c460f    ff91b4000000            call dword ptr [ecx+000000b4]
'006c4615    dbe2                    fnclex
'006c4617    3bc7                    cmp eax, edi
'006c4619    7d15                    jge 6c4630

If (var_41 < 0) Then
'006c461b    8b4de8                  mov ecx, dword ptr [ebp-18]
'006c461e    68b4000000              push 000000b4
'006c4623    6830314300              push 00433130
'006c4628    51                      push ecx
'006c4629    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c462a    ff1580104000            call dword ptr [00401080]
    
End If
'006c4630    8b45c4                  mov eax, dword ptr [ebp-3c]
'006c4633    8b10                    mov edx, dword ptr [eax]
'006c4635    8d5dc0                  lea ebx, dword ptr [ebp-40]
'006c4638    53                      push ebx
'006c4639    83ec10                  sub esp, 10
'006c463c    8bdc                    mov ebx, esp
'006c463e    b908000000              mov ecx, 00000008
'006c4643    890b                    mov dword ptr [ebx], ecx
'006c4645    8b8d50ffffff            mov ecx, dword ptr [ebp+ffffff50]
'006c464b    894b04                  mov dword ptr [ebx+04], ecx
'006c464e    c78554ffffffa0b84300    mov dword ptr [ebp+ffffff54], 0043b8a0
'006c4658    8b8d54ffffff            mov ecx, dword ptr [ebp+ffffff54]
'006c465e    894b08                  mov dword ptr [ebx+08], ecx
'006c4661    8b8d58ffffff            mov ecx, dword ptr [ebp+ffffff58]
'006c4667    50                      push eax
'006c4668    898514ffffff            mov dword ptr [ebp+ffffff14], eax
'006c466e    894b0c                  mov dword ptr [ebx+0c], ecx
'006c4671    ff5230                  call dword ptr [edx+30]
'006c4674    dbe2                    fnclex
'006c4676    3bc7                    cmp eax, edi
'006c4678    7d15                    jge 6c468f

If (0 < 0) Then
'006c467a    8b9514ffffff            mov edx, dword ptr [ebp+ffffff14]
'006c4680    6a30                    push 30
'006c4682    68d8304300              push 004330d8
'006c4687    52                      push edx
'006c4688    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c4689    ff1580104000            call dword ptr [00401080]
    
End If
'006c468f    8b45c0                  mov eax, dword ptr [ebp-40]
'006c4692    8945a4                  mov dword ptr [ebp-5c], eax
'006c4695    8d459c                  lea eax, dword ptr [ebp-64]
'006c4698    50                      push eax
'006c4699    8d4d8c                  lea ecx, dword ptr [ebp-74]
'006c469c    51                      push ecx
'006c469d    897dc0                  mov dword ptr [ebp-40], edi
'006c46a0    c7459c09000000          mov dword ptr [ebp-64], 00000009

' *** Reference to "rtcTrimVar"
'006c46a7    ff15e4104000            call dword ptr [004010e4]
'006c46ad    8d952cffffff            lea edx, dword ptr [ebp+ffffff2c]
'006c46b3    8d8d7cffffff            lea ecx, dword ptr [ebp+ffffff7c]
'006c46b9    c78534ffffffcc134300    mov dword ptr [ebp+ffffff34], 004313cc
'006c46c3    c7852cffffff08000000    mov dword ptr [ebp+ffffff2c], 00000008

' *** Reference to "__vbaVarDup"
'006c46cd    ff1598124000            call dword ptr [00401298]
var_59 = (vbNullChar)
'006c46d3    668b9528ffffff          mov dx, word ptr [ebp+ffffff28]
'006c46da    8d458c                  lea eax, dword ptr [ebp-74]
'006c46dd    50                      push eax
'006c46de    8d8d7cffffff            lea ecx, dword ptr [ebp+ffffff7c]
'006c46e4    66899544ffffff          mov word ptr [ebp+ffffff44], dx
'006c46eb    51                      push ecx
'006c46ec    8d953cffffff            lea edx, dword ptr [ebp+ffffff3c]
'006c46f2    52                      push edx
'006c46f3    8d856cffffff            lea eax, dword ptr [ebp+ffffff6c]
'006c46f9    50                      push eax
'006c46fa    c7853cffffff0b000000    mov dword ptr [ebp+ffffff3c], 0000000b

' *** Reference to "rtcImmediateIf"
'006c4704    ff154c124000            call dword ptr [0040124c]
'006c470a    8b1e                    mov ebx, dword ptr [esi]
'006c470c    8d8d6cffffff            lea ecx, dword ptr [ebp+ffffff6c]
'006c4712    51                      push ecx
'006c4713    8d55e4                  lea edx, dword ptr [ebp-1c]
'006c4716    52                      push edx

' *** Reference to "__vbaStrVarVal"
'006c4717    ff15fc114000            call dword ptr [004011fc]
'006c471d    50                      push eax
'006c471e    56                      push esi
'006c471f    ff93a4000000            call dword ptr [ebx+000000a4]
'006c4725    dbe2                    fnclex
'006c4727    3bc7                    cmp eax, edi
'006c4729    7d12                    jge 6c473d
'006c472b    68a4000000              push 000000a4
'006c4730    68d00d4300              push 00430dd0
'006c4735    56                      push esi
'006c4736    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c4737    ff1580104000            call dword ptr [00401080]
'006c473d    8d4de4                  lea ecx, dword ptr [ebp-1c]

' *** Reference to "__vbaFreeStr"
'006c4740    ff150c134000            call dword ptr [0040130c]
'#Cleanup(var_40)
'006c4746    8d45bc                  lea eax, dword ptr [ebp-44]
'006c4749    50                      push eax
'006c474a    8d4dc4                  lea ecx, dword ptr [ebp-3c]
'006c474d    51                      push ecx
'006c474e    8d55cc                  lea edx, dword ptr [ebp-34]
'006c4751    52                      push edx
'006c4752    6a03                    push 03

' *** Reference to "__vbaFreeObjList"
'006c4754    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_43, var_9, var_58)
'006c475a    8d856cffffff            lea eax, dword ptr [ebp+ffffff6c]
'006c4760    50                      push eax
'006c4761    8d4d8c                  lea ecx, dword ptr [ebp-74]
'006c4764    51                      push ecx
'006c4765    8d957cffffff            lea edx, dword ptr [ebp+ffffff7c]
'006c476b    52                      push edx
'006c476c    8d853cffffff            lea eax, dword ptr [ebp+ffffff3c]
'006c4772    50                      push eax
'006c4773    8d4d9c                  lea ecx, dword ptr [ebp-64]
'006c4776    51                      push ecx
'006c4777    8d55ac                  lea edx, dword ptr [ebp-54]
'006c477a    52                      push edx
'006c477b    6a06                    push 06

' *** Reference to "__vbaFreeVarList"
'006c477d    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_50, var_51, var_90, var_59, var_52, var_94)
'006c4783    8b45e8                  mov eax, dword ptr [ebp-18]
'006c4786    8b08                    mov ecx, dword ptr [eax]
'006c4788    83c42c                  add esp, 2c
'006c478b    8d55cc                  lea edx, dword ptr [ebp-34]
'006c478e    52                      push edx
'006c478f    50                      push eax
'006c4790    ff91b4000000            call dword ptr [ecx+000000b4]
'006c4796    dbe2                    fnclex
'006c4798    3bc7                    cmp eax, edi
'006c479a    7d15                    jge 6c47b1

If (var_41 < 0) Then
'006c479c    8b4de8                  mov ecx, dword ptr [ebp-18]
'006c479f    68b4000000              push 000000b4
'006c47a4    6830314300              push 00433130
'006c47a9    51                      push ecx
'006c47aa    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c47ab    ff1580104000            call dword ptr [00401080]
    
End If
'006c47b1    8b45cc                  mov eax, dword ptr [ebp-34]
'006c47b4    8b30                    mov esi, dword ptr [eax]
'006c47b6    8d5dc8                  lea ebx, dword ptr [ebp-38]
'006c47b9    53                      push ebx
'006c47ba    83ec10                  sub esp, 10
'006c47bd    8bdc                    mov ebx, esp
'006c47bf    ba08000000              mov edx, 00000008
'006c47c4    8913                    mov dword ptr [ebx], edx
'006c47c6    8b9560ffffff            mov edx, dword ptr [ebp+ffffff60]
'006c47cc    895304                  mov dword ptr [ebx+04], edx
'006c47cf    b918384500              mov ecx, 00453818
'006c47d4    894b08                  mov dword ptr [ebx+08], ecx
'006c47d7    8b8d68ffffff            mov ecx, dword ptr [ebp+ffffff68]
'006c47dd    50                      push eax
'006c47de    898520ffffff            mov dword ptr [ebp+ffffff20], eax
'006c47e4    894b0c                  mov dword ptr [ebx+0c], ecx
'006c47e7    ff5630                  call dword ptr [esi+30]
'006c47ea    dbe2                    fnclex
'006c47ec    3bc7                    cmp eax, edi
'006c47ee    7d15                    jge 6c4805

If (var_43 < 0) Then
'006c47f0    8b9520ffffff            mov edx, dword ptr [ebp+ffffff20]
'006c47f6    6a30                    push 30
'006c47f8    68d8304300              push 004330d8
'006c47fd    52                      push edx
'006c47fe    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c47ff    ff1580104000            call dword ptr [00401080]
    
End If
'006c4805    8b45c8                  mov eax, dword ptr [ebp-38]
'006c4808    8945b4                  mov dword ptr [ebp-4c], eax
'006c480b    8d45ac                  lea eax, dword ptr [ebp-54]
'006c480e    50                      push eax
'006c480f    897dc8                  mov dword ptr [ebp-38], edi
'006c4812    c745ac09000000          mov dword ptr [ebp-54], 00000009

' *** Reference to "rtcIsNull"
'006c4819    ff1540114000            call dword ptr [00401140]
'006c481f    898528ffffff            mov dword ptr [ebp+ffffff28], eax
'006c4825    8b4508                  mov eax, dword ptr [ebp+08]
'006c4828    8b08                    mov ecx, dword ptr [eax]
'006c482a    50                      push eax
'006c482b    ff9118030000            call dword ptr [ecx+00000318]
'006c4831    50                      push eax
'006c4832    8d55bc                  lea edx, dword ptr [ebp-44]
'006c4835    52                      push edx

' *** Reference to "__vbaObjSet"
'006c4836    ff15b4104000            call dword ptr [004010b4]
Set var_58 = Me
'006c483c    8d55c4                  lea edx, dword ptr [ebp-3c]
'006c483f    8bf0                    mov esi, eax
'006c4841    8b45e8                  mov eax, dword ptr [ebp-18]
'006c4844    8b08                    mov ecx, dword ptr [eax]
'006c4846    52                      push edx
'006c4847    50                      push eax
'006c4848    ff91b4000000            call dword ptr [ecx+000000b4]
'006c484e    dbe2                    fnclex
'006c4850    3bc7                    cmp eax, edi
'006c4852    7d15                    jge 6c4869

If (var_41 < 0) Then
'006c4854    8b4de8                  mov ecx, dword ptr [ebp-18]
'006c4857    68b4000000              push 000000b4
'006c485c    6830314300              push 00433130
'006c4861    51                      push ecx
'006c4862    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c4863    ff1580104000            call dword ptr [00401080]
    
End If
'006c4869    8b45c4                  mov eax, dword ptr [ebp-3c]
'006c486c    8b10                    mov edx, dword ptr [eax]
'006c486e    8d5dc0                  lea ebx, dword ptr [ebp-40]
'006c4871    53                      push ebx
'006c4872    83ec10                  sub esp, 10
'006c4875    8bdc                    mov ebx, esp
'006c4877    b908000000              mov ecx, 00000008
'006c487c    890b                    mov dword ptr [ebx], ecx
'006c487e    8b8d50ffffff            mov ecx, dword ptr [ebp+ffffff50]
'006c4884    894b04                  mov dword ptr [ebx+04], ecx
'006c4887    c78554ffffff18384500    mov dword ptr [ebp+ffffff54], 00453818
'006c4891    8b8d54ffffff            mov ecx, dword ptr [ebp+ffffff54]
'006c4897    894b08                  mov dword ptr [ebx+08], ecx
'006c489a    8b8d58ffffff            mov ecx, dword ptr [ebp+ffffff58]
'006c48a0    50                      push eax
'006c48a1    898514ffffff            mov dword ptr [ebp+ffffff14], eax
'006c48a7    894b0c                  mov dword ptr [ebx+0c], ecx
'006c48aa    ff5230                  call dword ptr [edx+30]
'006c48ad    dbe2                    fnclex
'006c48af    3bc7                    cmp eax, edi
'006c48b1    7d15                    jge 6c48c8

If (0 < 0) Then
'006c48b3    8b9514ffffff            mov edx, dword ptr [ebp+ffffff14]
'006c48b9    6a30                    push 30
'006c48bb    68d8304300              push 004330d8
'006c48c0    52                      push edx
'006c48c1    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c48c2    ff1580104000            call dword ptr [00401080]
    
End If
'006c48c8    8b45c0                  mov eax, dword ptr [ebp-40]
'006c48cb    8945a4                  mov dword ptr [ebp-5c], eax
'006c48ce    8d459c                  lea eax, dword ptr [ebp-64]
'006c48d1    50                      push eax
'006c48d2    8d4d8c                  lea ecx, dword ptr [ebp-74]
'006c48d5    51                      push ecx
'006c48d6    897dc0                  mov dword ptr [ebp-40], edi
'006c48d9    c7459c09000000          mov dword ptr [ebp-64], 00000009

' *** Reference to "rtcTrimVar"
'006c48e0    ff15e4104000            call dword ptr [004010e4]
'006c48e6    8d952cffffff            lea edx, dword ptr [ebp+ffffff2c]
'006c48ec    8d8d7cffffff            lea ecx, dword ptr [ebp+ffffff7c]
'006c48f2    c78534ffffffcc134300    mov dword ptr [ebp+ffffff34], 004313cc
'006c48fc    c7852cffffff08000000    mov dword ptr [ebp+ffffff2c], 00000008

' *** Reference to "__vbaVarDup"
'006c4906    ff1598124000            call dword ptr [00401298]
var_59 = (vbNullChar)
'006c490c    668b9528ffffff          mov dx, word ptr [ebp+ffffff28]
'006c4913    8d458c                  lea eax, dword ptr [ebp-74]
'006c4916    50                      push eax
'006c4917    8d8d7cffffff            lea ecx, dword ptr [ebp+ffffff7c]
'006c491d    66899544ffffff          mov word ptr [ebp+ffffff44], dx
'006c4924    51                      push ecx
'006c4925    8d953cffffff            lea edx, dword ptr [ebp+ffffff3c]
'006c492b    52                      push edx
'006c492c    8d856cffffff            lea eax, dword ptr [ebp+ffffff6c]
'006c4932    50                      push eax
'006c4933    c7853cffffff0b000000    mov dword ptr [ebp+ffffff3c], 0000000b

' *** Reference to "rtcImmediateIf"
'006c493d    ff154c124000            call dword ptr [0040124c]
'006c4943    8b1e                    mov ebx, dword ptr [esi]
'006c4945    8d8d6cffffff            lea ecx, dword ptr [ebp+ffffff6c]
'006c494b    51                      push ecx
'006c494c    8d55e4                  lea edx, dword ptr [ebp-1c]
'006c494f    52                      push edx

' *** Reference to "__vbaStrVarVal"
'006c4950    ff15fc114000            call dword ptr [004011fc]
'006c4956    50                      push eax
'006c4957    56                      push esi
'006c4958    ff93a4000000            call dword ptr [ebx+000000a4]
'006c495e    dbe2                    fnclex
'006c4960    3bc7                    cmp eax, edi
'006c4962    7d12                    jge 6c4976
'006c4964    68a4000000              push 000000a4
'006c4969    68d00d4300              push 00430dd0
'006c496e    56                      push esi
'006c496f    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c4970    ff1580104000            call dword ptr [00401080]
'006c4976    8d4de4                  lea ecx, dword ptr [ebp-1c]

' *** Reference to "__vbaFreeStr"
'006c4979    ff150c134000            call dword ptr [0040130c]
'#Cleanup(var_40)
'006c497f    8d45bc                  lea eax, dword ptr [ebp-44]
'006c4982    50                      push eax
'006c4983    8d4dc4                  lea ecx, dword ptr [ebp-3c]
'006c4986    51                      push ecx
'006c4987    8d55cc                  lea edx, dword ptr [ebp-34]
'006c498a    52                      push edx
'006c498b    6a03                    push 03

' *** Reference to "__vbaFreeObjList"
'006c498d    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_43, var_9, var_58)
'006c4993    8d856cffffff            lea eax, dword ptr [ebp+ffffff6c]
'006c4999    50                      push eax
'006c499a    8d4d8c                  lea ecx, dword ptr [ebp-74]
'006c499d    51                      push ecx
'006c499e    8d957cffffff            lea edx, dword ptr [ebp+ffffff7c]
'006c49a4    52                      push edx
'006c49a5    8d853cffffff            lea eax, dword ptr [ebp+ffffff3c]
'006c49ab    50                      push eax
'006c49ac    8d4d9c                  lea ecx, dword ptr [ebp-64]
'006c49af    51                      push ecx
'006c49b0    8d55ac                  lea edx, dword ptr [ebp-54]
'006c49b3    52                      push edx
'006c49b4    6a06                    push 06

' *** Reference to "__vbaFreeVarList"
'006c49b6    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_50, var_51, var_90, var_59, var_52, var_94)
'006c49bc    8b45e8                  mov eax, dword ptr [ebp-18]
'006c49bf    8b08                    mov ecx, dword ptr [eax]
'006c49c1    83c42c                  add esp, 2c
'006c49c4    8d55cc                  lea edx, dword ptr [ebp-34]
'006c49c7    52                      push edx
'006c49c8    50                      push eax
'006c49c9    ff91b4000000            call dword ptr [ecx+000000b4]
'006c49cf    dbe2                    fnclex
'006c49d1    3bc7                    cmp eax, edi
'006c49d3    7d15                    jge 6c49ea

If (var_41 < 0) Then
'006c49d5    8b4de8                  mov ecx, dword ptr [ebp-18]
'006c49d8    68b4000000              push 000000b4
'006c49dd    6830314300              push 00433130
'006c49e2    51                      push ecx
'006c49e3    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c49e4    ff1580104000            call dword ptr [00401080]
    
End If
'006c49ea    8b45cc                  mov eax, dword ptr [ebp-34]
'006c49ed    8b30                    mov esi, dword ptr [eax]
'006c49ef    8d5dc8                  lea ebx, dword ptr [ebp-38]
'006c49f2    53                      push ebx
'006c49f3    83ec10                  sub esp, 10
'006c49f6    8bdc                    mov ebx, esp
'006c49f8    ba08000000              mov edx, 00000008
'006c49fd    8913                    mov dword ptr [ebx], edx
'006c49ff    8b9560ffffff            mov edx, dword ptr [ebp+ffffff60]
'006c4a05    895304                  mov dword ptr [ebx+04], edx
'006c4a08    b9e4b24300              mov ecx, 0043b2e4
'006c4a0d    894b08                  mov dword ptr [ebx+08], ecx
'006c4a10    8b8d68ffffff            mov ecx, dword ptr [ebp+ffffff68]
'006c4a16    50                      push eax
'006c4a17    898520ffffff            mov dword ptr [ebp+ffffff20], eax
'006c4a1d    894b0c                  mov dword ptr [ebx+0c], ecx
'006c4a20    ff5630                  call dword ptr [esi+30]
'006c4a23    dbe2                    fnclex
'006c4a25    3bc7                    cmp eax, edi
'006c4a27    7d15                    jge 6c4a3e

If (var_43 < 0) Then
'006c4a29    8b9520ffffff            mov edx, dword ptr [ebp+ffffff20]
'006c4a2f    6a30                    push 30
'006c4a31    68d8304300              push 004330d8
'006c4a36    52                      push edx
'006c4a37    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c4a38    ff1580104000            call dword ptr [00401080]
    
End If
'006c4a3e    8b45c8                  mov eax, dword ptr [ebp-38]
'006c4a41    8945b4                  mov dword ptr [ebp-4c], eax
'006c4a44    8d45ac                  lea eax, dword ptr [ebp-54]
'006c4a47    50                      push eax
'006c4a48    897dc8                  mov dword ptr [ebp-38], edi
'006c4a4b    c745ac09000000          mov dword ptr [ebp-54], 00000009

' *** Reference to "rtcIsNull"
'006c4a52    ff1540114000            call dword ptr [00401140]
'006c4a58    898528ffffff            mov dword ptr [ebp+ffffff28], eax
'006c4a5e    8b4508                  mov eax, dword ptr [ebp+08]
'006c4a61    8b08                    mov ecx, dword ptr [eax]
'006c4a63    50                      push eax
'006c4a64    ff911c030000            call dword ptr [ecx+0000031c]
'006c4a6a    50                      push eax
'006c4a6b    8d55bc                  lea edx, dword ptr [ebp-44]
'006c4a6e    52                      push edx

' *** Reference to "__vbaObjSet"
'006c4a6f    ff15b4104000            call dword ptr [004010b4]
Set var_58 = Me
'006c4a75    8d55c4                  lea edx, dword ptr [ebp-3c]
'006c4a78    8bf0                    mov esi, eax
'006c4a7a    8b45e8                  mov eax, dword ptr [ebp-18]
'006c4a7d    8b08                    mov ecx, dword ptr [eax]
'006c4a7f    52                      push edx
'006c4a80    50                      push eax
'006c4a81    ff91b4000000            call dword ptr [ecx+000000b4]
'006c4a87    dbe2                    fnclex
'006c4a89    3bc7                    cmp eax, edi
'006c4a8b    7d15                    jge 6c4aa2

If (var_41 < 0) Then
'006c4a8d    8b4de8                  mov ecx, dword ptr [ebp-18]
'006c4a90    68b4000000              push 000000b4
'006c4a95    6830314300              push 00433130
'006c4a9a    51                      push ecx
'006c4a9b    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c4a9c    ff1580104000            call dword ptr [00401080]
    
End If
'006c4aa2    8b45c4                  mov eax, dword ptr [ebp-3c]
'006c4aa5    8b10                    mov edx, dword ptr [eax]
'006c4aa7    8d5dc0                  lea ebx, dword ptr [ebp-40]
'006c4aaa    53                      push ebx
'006c4aab    83ec10                  sub esp, 10
'006c4aae    8bdc                    mov ebx, esp
'006c4ab0    b908000000              mov ecx, 00000008
'006c4ab5    890b                    mov dword ptr [ebx], ecx
'006c4ab7    8b8d50ffffff            mov ecx, dword ptr [ebp+ffffff50]
'006c4abd    894b04                  mov dword ptr [ebx+04], ecx
'006c4ac0    c78554ffffffe4b24300    mov dword ptr [ebp+ffffff54], 0043b2e4
'006c4aca    8b8d54ffffff            mov ecx, dword ptr [ebp+ffffff54]
'006c4ad0    894b08                  mov dword ptr [ebx+08], ecx
'006c4ad3    8b8d58ffffff            mov ecx, dword ptr [ebp+ffffff58]
'006c4ad9    50                      push eax
'006c4ada    898514ffffff            mov dword ptr [ebp+ffffff14], eax
'006c4ae0    894b0c                  mov dword ptr [ebx+0c], ecx
'006c4ae3    ff5230                  call dword ptr [edx+30]
'006c4ae6    dbe2                    fnclex
'006c4ae8    3bc7                    cmp eax, edi
'006c4aea    7d15                    jge 6c4b01

If (0 < 0) Then
'006c4aec    8b9514ffffff            mov edx, dword ptr [ebp+ffffff14]
'006c4af2    6a30                    push 30
'006c4af4    68d8304300              push 004330d8
'006c4af9    52                      push edx
'006c4afa    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c4afb    ff1580104000            call dword ptr [00401080]
    
End If
'006c4b01    8b45c0                  mov eax, dword ptr [ebp-40]
'006c4b04    8945a4                  mov dword ptr [ebp-5c], eax
'006c4b07    8d459c                  lea eax, dword ptr [ebp-64]
'006c4b0a    50                      push eax
'006c4b0b    8d4d8c                  lea ecx, dword ptr [ebp-74]
'006c4b0e    51                      push ecx
'006c4b0f    897dc0                  mov dword ptr [ebp-40], edi
'006c4b12    c7459c09000000          mov dword ptr [ebp-64], 00000009

' *** Reference to "rtcTrimVar"
'006c4b19    ff15e4104000            call dword ptr [004010e4]
'006c4b1f    8d952cffffff            lea edx, dword ptr [ebp+ffffff2c]
'006c4b25    8d8d7cffffff            lea ecx, dword ptr [ebp+ffffff7c]
'006c4b2b    c78534ffffffcc134300    mov dword ptr [ebp+ffffff34], 004313cc
'006c4b35    c7852cffffff08000000    mov dword ptr [ebp+ffffff2c], 00000008

' *** Reference to "__vbaVarDup"
'006c4b3f    ff1598124000            call dword ptr [00401298]
var_59 = (vbNullChar)
'006c4b45    668b9528ffffff          mov dx, word ptr [ebp+ffffff28]
'006c4b4c    8d458c                  lea eax, dword ptr [ebp-74]
'006c4b4f    50                      push eax
'006c4b50    8d8d7cffffff            lea ecx, dword ptr [ebp+ffffff7c]
'006c4b56    66899544ffffff          mov word ptr [ebp+ffffff44], dx
'006c4b5d    51                      push ecx
'006c4b5e    8d953cffffff            lea edx, dword ptr [ebp+ffffff3c]
'006c4b64    52                      push edx
'006c4b65    8d856cffffff            lea eax, dword ptr [ebp+ffffff6c]
'006c4b6b    50                      push eax
'006c4b6c    c7853cffffff0b000000    mov dword ptr [ebp+ffffff3c], 0000000b

' *** Reference to "rtcImmediateIf"
'006c4b76    ff154c124000            call dword ptr [0040124c]
'006c4b7c    8b1e                    mov ebx, dword ptr [esi]
'006c4b7e    8d8d6cffffff            lea ecx, dword ptr [ebp+ffffff6c]
'006c4b84    51                      push ecx
'006c4b85    8d55e4                  lea edx, dword ptr [ebp-1c]
'006c4b88    52                      push edx

' *** Reference to "__vbaStrVarVal"
'006c4b89    ff15fc114000            call dword ptr [004011fc]
'006c4b8f    50                      push eax
'006c4b90    56                      push esi
'006c4b91    ff93a4000000            call dword ptr [ebx+000000a4]
'006c4b97    dbe2                    fnclex
'006c4b99    3bc7                    cmp eax, edi
'006c4b9b    7d12                    jge 6c4baf
'006c4b9d    68a4000000              push 000000a4
'006c4ba2    68d00d4300              push 00430dd0
'006c4ba7    56                      push esi
'006c4ba8    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c4ba9    ff1580104000            call dword ptr [00401080]
'006c4baf    8d4de4                  lea ecx, dword ptr [ebp-1c]

' *** Reference to "__vbaFreeStr"
'006c4bb2    ff150c134000            call dword ptr [0040130c]
'#Cleanup(var_40)
'006c4bb8    8d45bc                  lea eax, dword ptr [ebp-44]
'006c4bbb    50                      push eax
'006c4bbc    8d4dc4                  lea ecx, dword ptr [ebp-3c]
'006c4bbf    51                      push ecx
'006c4bc0    8d55cc                  lea edx, dword ptr [ebp-34]
'006c4bc3    52                      push edx
'006c4bc4    6a03                    push 03

' *** Reference to "__vbaFreeObjList"
'006c4bc6    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_43, var_9, var_58)
'006c4bcc    8d856cffffff            lea eax, dword ptr [ebp+ffffff6c]
'006c4bd2    50                      push eax
'006c4bd3    8d4d8c                  lea ecx, dword ptr [ebp-74]
'006c4bd6    51                      push ecx
'006c4bd7    8d957cffffff            lea edx, dword ptr [ebp+ffffff7c]
'006c4bdd    52                      push edx
'006c4bde    8d853cffffff            lea eax, dword ptr [ebp+ffffff3c]
'006c4be4    50                      push eax
'006c4be5    8d4d9c                  lea ecx, dword ptr [ebp-64]
'006c4be8    51                      push ecx
'006c4be9    8d55ac                  lea edx, dword ptr [ebp-54]
'006c4bec    52                      push edx
'006c4bed    6a06                    push 06

' *** Reference to "__vbaFreeVarList"
'006c4bef    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_50, var_51, var_90, var_59, var_52, var_94)
'006c4bf5    8b45e8                  mov eax, dword ptr [ebp-18]
'006c4bf8    8b08                    mov ecx, dword ptr [eax]
'006c4bfa    83c42c                  add esp, 2c
'006c4bfd    8d55cc                  lea edx, dword ptr [ebp-34]
'006c4c00    52                      push edx
'006c4c01    50                      push eax
'006c4c02    ff91b4000000            call dword ptr [ecx+000000b4]
'006c4c08    dbe2                    fnclex
'006c4c0a    3bc7                    cmp eax, edi
'006c4c0c    7d15                    jge 6c4c23

If (var_41 < 0) Then
'006c4c0e    8b4de8                  mov ecx, dword ptr [ebp-18]
'006c4c11    68b4000000              push 000000b4
'006c4c16    6830314300              push 00433130
'006c4c1b    51                      push ecx
'006c4c1c    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c4c1d    ff1580104000            call dword ptr [00401080]
    
End If
'006c4c23    8b45cc                  mov eax, dword ptr [ebp-34]
'006c4c26    8b30                    mov esi, dword ptr [eax]
'006c4c28    8d5dc8                  lea ebx, dword ptr [ebp-38]
'006c4c2b    53                      push ebx
'006c4c2c    83ec10                  sub esp, 10
'006c4c2f    8bdc                    mov ebx, esp
'006c4c31    ba08000000              mov edx, 00000008
'006c4c36    8913                    mov dword ptr [ebx], edx
'006c4c38    8b9560ffffff            mov edx, dword ptr [ebp+ffffff60]
'006c4c3e    895304                  mov dword ptr [ebx+04], edx
'006c4c41    b930384500              mov ecx, 00453830
'006c4c46    894b08                  mov dword ptr [ebx+08], ecx
'006c4c49    8b8d68ffffff            mov ecx, dword ptr [ebp+ffffff68]
'006c4c4f    50                      push eax
'006c4c50    898520ffffff            mov dword ptr [ebp+ffffff20], eax
'006c4c56    894b0c                  mov dword ptr [ebx+0c], ecx
'006c4c59    ff5630                  call dword ptr [esi+30]
'006c4c5c    dbe2                    fnclex
'006c4c5e    3bc7                    cmp eax, edi
'006c4c60    7d15                    jge 6c4c77

If (var_43 < 0) Then
'006c4c62    8b9520ffffff            mov edx, dword ptr [ebp+ffffff20]
'006c4c68    6a30                    push 30
'006c4c6a    68d8304300              push 004330d8
'006c4c6f    52                      push edx
'006c4c70    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c4c71    ff1580104000            call dword ptr [00401080]
    
End If
'006c4c77    8b45c8                  mov eax, dword ptr [ebp-38]
'006c4c7a    8945b4                  mov dword ptr [ebp-4c], eax
'006c4c7d    8d45ac                  lea eax, dword ptr [ebp-54]
'006c4c80    50                      push eax
'006c4c81    897dc8                  mov dword ptr [ebp-38], edi
'006c4c84    c745ac09000000          mov dword ptr [ebp-54], 00000009

' *** Reference to "rtcIsNull"
'006c4c8b    ff1540114000            call dword ptr [00401140]
'006c4c91    898528ffffff            mov dword ptr [ebp+ffffff28], eax
'006c4c97    8b4508                  mov eax, dword ptr [ebp+08]
'006c4c9a    8b08                    mov ecx, dword ptr [eax]
'006c4c9c    50                      push eax
'006c4c9d    ff9114030000            call dword ptr [ecx+00000314]
'006c4ca3    50                      push eax
'006c4ca4    8d55bc                  lea edx, dword ptr [ebp-44]
'006c4ca7    52                      push edx

' *** Reference to "__vbaObjSet"
'006c4ca8    ff15b4104000            call dword ptr [004010b4]
Set var_58 = Me
'006c4cae    8d55c4                  lea edx, dword ptr [ebp-3c]
'006c4cb1    8bf0                    mov esi, eax
'006c4cb3    8b45e8                  mov eax, dword ptr [ebp-18]
'006c4cb6    8b08                    mov ecx, dword ptr [eax]
'006c4cb8    52                      push edx
'006c4cb9    50                      push eax
'006c4cba    ff91b4000000            call dword ptr [ecx+000000b4]
'006c4cc0    dbe2                    fnclex
'006c4cc2    3bc7                    cmp eax, edi
'006c4cc4    7d15                    jge 6c4cdb

If (var_41 < 0) Then
'006c4cc6    8b4de8                  mov ecx, dword ptr [ebp-18]
'006c4cc9    68b4000000              push 000000b4
'006c4cce    6830314300              push 00433130
'006c4cd3    51                      push ecx
'006c4cd4    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c4cd5    ff1580104000            call dword ptr [00401080]
    
End If
'006c4cdb    8b45c4                  mov eax, dword ptr [ebp-3c]
'006c4cde    8b10                    mov edx, dword ptr [eax]
'006c4ce0    8d5dc0                  lea ebx, dword ptr [ebp-40]
'006c4ce3    53                      push ebx
'006c4ce4    83ec10                  sub esp, 10
'006c4ce7    8bdc                    mov ebx, esp
'006c4ce9    b908000000              mov ecx, 00000008
'006c4cee    890b                    mov dword ptr [ebx], ecx
'006c4cf0    8b8d50ffffff            mov ecx, dword ptr [ebp+ffffff50]
'006c4cf6    894b04                  mov dword ptr [ebx+04], ecx
'006c4cf9    c78554ffffff30384500    mov dword ptr [ebp+ffffff54], 00453830
'006c4d03    8b8d54ffffff            mov ecx, dword ptr [ebp+ffffff54]
'006c4d09    894b08                  mov dword ptr [ebx+08], ecx
'006c4d0c    8b8d58ffffff            mov ecx, dword ptr [ebp+ffffff58]
'006c4d12    50                      push eax
'006c4d13    898514ffffff            mov dword ptr [ebp+ffffff14], eax
'006c4d19    894b0c                  mov dword ptr [ebx+0c], ecx
'006c4d1c    ff5230                  call dword ptr [edx+30]
'006c4d1f    dbe2                    fnclex
'006c4d21    3bc7                    cmp eax, edi
'006c4d23    7d15                    jge 6c4d3a

If (0 < 0) Then
'006c4d25    8b9514ffffff            mov edx, dword ptr [ebp+ffffff14]
'006c4d2b    6a30                    push 30
'006c4d2d    68d8304300              push 004330d8
'006c4d32    52                      push edx
'006c4d33    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c4d34    ff1580104000            call dword ptr [00401080]
    
End If
'006c4d3a    8b45c0                  mov eax, dword ptr [ebp-40]
'006c4d3d    8945a4                  mov dword ptr [ebp-5c], eax
'006c4d40    8d459c                  lea eax, dword ptr [ebp-64]
'006c4d43    50                      push eax
'006c4d44    8d4d8c                  lea ecx, dword ptr [ebp-74]
'006c4d47    51                      push ecx
'006c4d48    897dc0                  mov dword ptr [ebp-40], edi
'006c4d4b    c7459c09000000          mov dword ptr [ebp-64], 00000009

' *** Reference to "rtcTrimVar"
'006c4d52    ff15e4104000            call dword ptr [004010e4]
'006c4d58    8d952cffffff            lea edx, dword ptr [ebp+ffffff2c]
'006c4d5e    8d8d7cffffff            lea ecx, dword ptr [ebp+ffffff7c]
'006c4d64    c78534ffffffcc134300    mov dword ptr [ebp+ffffff34], 004313cc
'006c4d6e    c7852cffffff08000000    mov dword ptr [ebp+ffffff2c], 00000008

' *** Reference to "__vbaVarDup"
'006c4d78    ff1598124000            call dword ptr [00401298]
var_59 = (vbNullChar)
'006c4d7e    668b9528ffffff          mov dx, word ptr [ebp+ffffff28]
'006c4d85    8d458c                  lea eax, dword ptr [ebp-74]
'006c4d88    50                      push eax
'006c4d89    8d8d7cffffff            lea ecx, dword ptr [ebp+ffffff7c]
'006c4d8f    66899544ffffff          mov word ptr [ebp+ffffff44], dx
'006c4d96    51                      push ecx
'006c4d97    8d953cffffff            lea edx, dword ptr [ebp+ffffff3c]
'006c4d9d    52                      push edx
'006c4d9e    8d856cffffff            lea eax, dword ptr [ebp+ffffff6c]
'006c4da4    50                      push eax
'006c4da5    c7853cffffff0b000000    mov dword ptr [ebp+ffffff3c], 0000000b

' *** Reference to "rtcImmediateIf"
'006c4daf    ff154c124000            call dword ptr [0040124c]
'006c4db5    8b1e                    mov ebx, dword ptr [esi]
'006c4db7    8d8d6cffffff            lea ecx, dword ptr [ebp+ffffff6c]
'006c4dbd    51                      push ecx
'006c4dbe    8d55e4                  lea edx, dword ptr [ebp-1c]
'006c4dc1    52                      push edx

' *** Reference to "__vbaStrVarVal"
'006c4dc2    ff15fc114000            call dword ptr [004011fc]
'006c4dc8    50                      push eax
'006c4dc9    56                      push esi
'006c4dca    ff93a4000000            call dword ptr [ebx+000000a4]
'006c4dd0    dbe2                    fnclex
'006c4dd2    3bc7                    cmp eax, edi
'006c4dd4    7d12                    jge 6c4de8
'006c4dd6    68a4000000              push 000000a4
'006c4ddb    68d00d4300              push 00430dd0
'006c4de0    56                      push esi
'006c4de1    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c4de2    ff1580104000            call dword ptr [00401080]
'006c4de8    8d4de4                  lea ecx, dword ptr [ebp-1c]

' *** Reference to "__vbaFreeStr"
'006c4deb    ff150c134000            call dword ptr [0040130c]
'#Cleanup(var_40)
'006c4df1    8d45bc                  lea eax, dword ptr [ebp-44]
'006c4df4    50                      push eax
'006c4df5    8d4dc4                  lea ecx, dword ptr [ebp-3c]
'006c4df8    51                      push ecx
'006c4df9    8d55cc                  lea edx, dword ptr [ebp-34]
'006c4dfc    52                      push edx
'006c4dfd    6a03                    push 03

' *** Reference to "__vbaFreeObjList"
'006c4dff    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_43, var_9, var_58)
'006c4e05    8d856cffffff            lea eax, dword ptr [ebp+ffffff6c]
'006c4e0b    50                      push eax
'006c4e0c    8d4d8c                  lea ecx, dword ptr [ebp-74]
'006c4e0f    51                      push ecx
'006c4e10    8d957cffffff            lea edx, dword ptr [ebp+ffffff7c]
'006c4e16    52                      push edx
'006c4e17    8d853cffffff            lea eax, dword ptr [ebp+ffffff3c]
'006c4e1d    50                      push eax
'006c4e1e    8d4d9c                  lea ecx, dword ptr [ebp-64]
'006c4e21    51                      push ecx
'006c4e22    8d55ac                  lea edx, dword ptr [ebp-54]
'006c4e25    52                      push edx
'006c4e26    6a06                    push 06

' *** Reference to "__vbaFreeVarList"
'006c4e28    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_50, var_51, var_90, var_59, var_52, var_94)
'006c4e2e    8b45e8                  mov eax, dword ptr [ebp-18]
'006c4e31    8b08                    mov ecx, dword ptr [eax]
'006c4e33    83c42c                  add esp, 2c
'006c4e36    8d55cc                  lea edx, dword ptr [ebp-34]
'006c4e39    52                      push edx
'006c4e3a    50                      push eax
'006c4e3b    ff91b4000000            call dword ptr [ecx+000000b4]
'006c4e41    dbe2                    fnclex
'006c4e43    3bc7                    cmp eax, edi
'006c4e45    7d15                    jge 6c4e5c

If (var_41 < 0) Then
'006c4e47    8b4de8                  mov ecx, dword ptr [ebp-18]
'006c4e4a    68b4000000              push 000000b4
'006c4e4f    6830314300              push 00433130
'006c4e54    51                      push ecx
'006c4e55    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c4e56    ff1580104000            call dword ptr [00401080]
    
End If
'006c4e5c    8b45cc                  mov eax, dword ptr [ebp-34]
'006c4e5f    8b30                    mov esi, dword ptr [eax]
'006c4e61    8d5dc8                  lea ebx, dword ptr [ebp-38]
'006c4e64    53                      push ebx
'006c4e65    83ec10                  sub esp, 10
'006c4e68    8bdc                    mov ebx, esp
'006c4e6a    ba08000000              mov edx, 00000008
'006c4e6f    8913                    mov dword ptr [ebx], edx
'006c4e71    8b9560ffffff            mov edx, dword ptr [ebp+ffffff60]
'006c4e77    895304                  mov dword ptr [ebx+04], edx
'006c4e7a    b944384500              mov ecx, 00453844
'006c4e7f    894b08                  mov dword ptr [ebx+08], ecx
'006c4e82    8b8d68ffffff            mov ecx, dword ptr [ebp+ffffff68]
'006c4e88    50                      push eax
'006c4e89    898520ffffff            mov dword ptr [ebp+ffffff20], eax
'006c4e8f    894b0c                  mov dword ptr [ebx+0c], ecx
'006c4e92    ff5630                  call dword ptr [esi+30]
'006c4e95    dbe2                    fnclex
'006c4e97    3bc7                    cmp eax, edi
'006c4e99    7d15                    jge 6c4eb0

If (var_43 < 0) Then
'006c4e9b    8b9520ffffff            mov edx, dword ptr [ebp+ffffff20]
'006c4ea1    6a30                    push 30
'006c4ea3    68d8304300              push 004330d8
'006c4ea8    52                      push edx
'006c4ea9    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c4eaa    ff1580104000            call dword ptr [00401080]
    
End If
'006c4eb0    8b45c8                  mov eax, dword ptr [ebp-38]
'006c4eb3    8945b4                  mov dword ptr [ebp-4c], eax
'006c4eb6    8d45ac                  lea eax, dword ptr [ebp-54]
'006c4eb9    50                      push eax
'006c4eba    897dc8                  mov dword ptr [ebp-38], edi
'006c4ebd    c745ac09000000          mov dword ptr [ebp-54], 00000009

' *** Reference to "rtcIsNull"
'006c4ec4    ff1540114000            call dword ptr [00401140]
'006c4eca    898528ffffff            mov dword ptr [ebp+ffffff28], eax
'006c4ed0    8b4508                  mov eax, dword ptr [ebp+08]
'006c4ed3    8b08                    mov ecx, dword ptr [eax]
'006c4ed5    50                      push eax
'006c4ed6    ff9110030000            call dword ptr [ecx+00000310]
'006c4edc    50                      push eax
'006c4edd    8d55bc                  lea edx, dword ptr [ebp-44]
'006c4ee0    52                      push edx

' *** Reference to "__vbaObjSet"
'006c4ee1    ff15b4104000            call dword ptr [004010b4]
Set var_58 = Me
'006c4ee7    8d55c4                  lea edx, dword ptr [ebp-3c]
'006c4eea    8bf0                    mov esi, eax
'006c4eec    8b45e8                  mov eax, dword ptr [ebp-18]
'006c4eef    8b08                    mov ecx, dword ptr [eax]
'006c4ef1    52                      push edx
'006c4ef2    50                      push eax
'006c4ef3    ff91b4000000            call dword ptr [ecx+000000b4]
'006c4ef9    dbe2                    fnclex
'006c4efb    3bc7                    cmp eax, edi
'006c4efd    7d15                    jge 6c4f14

If (var_41 < 0) Then
'006c4eff    8b4de8                  mov ecx, dword ptr [ebp-18]
'006c4f02    68b4000000              push 000000b4
'006c4f07    6830314300              push 00433130
'006c4f0c    51                      push ecx
'006c4f0d    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c4f0e    ff1580104000            call dword ptr [00401080]
    
End If
'006c4f14    8b45c4                  mov eax, dword ptr [ebp-3c]
'006c4f17    8b10                    mov edx, dword ptr [eax]
'006c4f19    8d5dc0                  lea ebx, dword ptr [ebp-40]
'006c4f1c    53                      push ebx
'006c4f1d    83ec10                  sub esp, 10
'006c4f20    8bdc                    mov ebx, esp
'006c4f22    b908000000              mov ecx, 00000008
'006c4f27    890b                    mov dword ptr [ebx], ecx
'006c4f29    8b8d50ffffff            mov ecx, dword ptr [ebp+ffffff50]
'006c4f2f    894b04                  mov dword ptr [ebx+04], ecx
'006c4f32    c78554ffffff44384500    mov dword ptr [ebp+ffffff54], 00453844
'006c4f3c    8b8d54ffffff            mov ecx, dword ptr [ebp+ffffff54]
'006c4f42    894b08                  mov dword ptr [ebx+08], ecx
'006c4f45    8b8d58ffffff            mov ecx, dword ptr [ebp+ffffff58]
'006c4f4b    50                      push eax
'006c4f4c    898514ffffff            mov dword ptr [ebp+ffffff14], eax
'006c4f52    894b0c                  mov dword ptr [ebx+0c], ecx
'006c4f55    ff5230                  call dword ptr [edx+30]
'006c4f58    dbe2                    fnclex
'006c4f5a    3bc7                    cmp eax, edi
'006c4f5c    7d15                    jge 6c4f73

If (0 < 0) Then
'006c4f5e    8b9514ffffff            mov edx, dword ptr [ebp+ffffff14]
'006c4f64    6a30                    push 30
'006c4f66    68d8304300              push 004330d8
'006c4f6b    52                      push edx
'006c4f6c    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c4f6d    ff1580104000            call dword ptr [00401080]
    
End If
'006c4f73    8b45c0                  mov eax, dword ptr [ebp-40]
'006c4f76    8945a4                  mov dword ptr [ebp-5c], eax
'006c4f79    8d459c                  lea eax, dword ptr [ebp-64]
'006c4f7c    50                      push eax
'006c4f7d    8d4d8c                  lea ecx, dword ptr [ebp-74]
'006c4f80    51                      push ecx
'006c4f81    897dc0                  mov dword ptr [ebp-40], edi
'006c4f84    c7459c09000000          mov dword ptr [ebp-64], 00000009

' *** Reference to "rtcTrimVar"
'006c4f8b    ff15e4104000            call dword ptr [004010e4]
'006c4f91    8d952cffffff            lea edx, dword ptr [ebp+ffffff2c]
'006c4f97    8d8d7cffffff            lea ecx, dword ptr [ebp+ffffff7c]
'006c4f9d    c78534ffffffcc134300    mov dword ptr [ebp+ffffff34], 004313cc
'006c4fa7    c7852cffffff08000000    mov dword ptr [ebp+ffffff2c], 00000008

' *** Reference to "__vbaVarDup"
'006c4fb1    ff1598124000            call dword ptr [00401298]
var_59 = (vbNullChar)
'006c4fb7    668b9528ffffff          mov dx, word ptr [ebp+ffffff28]
'006c4fbe    8d458c                  lea eax, dword ptr [ebp-74]
'006c4fc1    50                      push eax
'006c4fc2    8d8d7cffffff            lea ecx, dword ptr [ebp+ffffff7c]
'006c4fc8    66899544ffffff          mov word ptr [ebp+ffffff44], dx
'006c4fcf    51                      push ecx
'006c4fd0    8d953cffffff            lea edx, dword ptr [ebp+ffffff3c]
'006c4fd6    52                      push edx
'006c4fd7    8d856cffffff            lea eax, dword ptr [ebp+ffffff6c]
'006c4fdd    50                      push eax
'006c4fde    c7853cffffff0b000000    mov dword ptr [ebp+ffffff3c], 0000000b

' *** Reference to "rtcImmediateIf"
'006c4fe8    ff154c124000            call dword ptr [0040124c]
'006c4fee    8b1e                    mov ebx, dword ptr [esi]
'006c4ff0    8d8d6cffffff            lea ecx, dword ptr [ebp+ffffff6c]
'006c4ff6    51                      push ecx
'006c4ff7    8d55e4                  lea edx, dword ptr [ebp-1c]
'006c4ffa    52                      push edx

' *** Reference to "__vbaStrVarVal"
'006c4ffb    ff15fc114000            call dword ptr [004011fc]
'006c5001    50                      push eax
'006c5002    56                      push esi
'006c5003    ff93a4000000            call dword ptr [ebx+000000a4]
'006c5009    dbe2                    fnclex
'006c500b    3bc7                    cmp eax, edi
'006c500d    7d16                    jge 6c5025
'006c500f    68a4000000              push 000000a4
'006c5014    68d00d4300              push 00430dd0
'006c5019    56                      push esi

' *** Reference to "__vbaHresultCheckObj"
'006c501a    8b3580104000            mov esi, dword ptr [00401080]
'006c5020    50                      push eax
'006c5021    ffd6                    call esi
'006c5023    eb06                    jmp 6c502b

' *** Reference to "__vbaHresultCheckObj"
'006c5025    8b3580104000            mov esi, dword ptr [00401080]
'006c502b    8d4de4                  lea ecx, dword ptr [ebp-1c]

' *** Reference to "__vbaFreeStr"
'006c502e    ff150c134000            call dword ptr [0040130c]
'#Cleanup(var_40)
'006c5034    8d45bc                  lea eax, dword ptr [ebp-44]
'006c5037    50                      push eax
'006c5038    8d4dc4                  lea ecx, dword ptr [ebp-3c]
'006c503b    51                      push ecx
'006c503c    8d55cc                  lea edx, dword ptr [ebp-34]
'006c503f    52                      push edx
'006c5040    6a03                    push 03

' *** Reference to "__vbaFreeObjList"
'006c5042    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_43, var_9, var_58)
'006c5048    8d856cffffff            lea eax, dword ptr [ebp+ffffff6c]
'006c504e    50                      push eax
'006c504f    8d4d8c                  lea ecx, dword ptr [ebp-74]
'006c5052    51                      push ecx
'006c5053    8d957cffffff            lea edx, dword ptr [ebp+ffffff7c]
'006c5059    52                      push edx
'006c505a    8d853cffffff            lea eax, dword ptr [ebp+ffffff3c]
'006c5060    50                      push eax
'006c5061    8d4d9c                  lea ecx, dword ptr [ebp-64]
'006c5064    51                      push ecx
'006c5065    8d55ac                  lea edx, dword ptr [ebp-54]
'006c5068    52                      push edx
'006c5069    6a06                    push 06

' *** Reference to "__vbaFreeVarList"
'006c506b    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_50, var_51, var_90, var_59, var_52, var_94)
'006c5071    83c42c                  add esp, 2c

'ERROR: Two many next close:
End If
'006c5074    8b45e8                  mov eax, dword ptr [ebp-18]
'006c5077    8b08                    mov ecx, dword ptr [eax]
'006c5079    50                      push eax
'006c507a    ff91c4000000            call dword ptr [ecx+000000c4]
'006c5080    dbe2                    fnclex
'006c5082    3bc7                    cmp eax, edi
'006c5084    7d11                    jge 6c5097

If (var_41 < 0) Then
'006c5086    8b55e8                  mov edx, dword ptr [ebp-18]
'006c5089    68c4000000              push 000000c4
'006c508e    6830314300              push 00433130
'006c5093    52                      push edx
'006c5094    50                      push eax
'006c5095    ffd6                    call esi
    
End If
'006c5097    680a516c00              push 006c510a
'006c509c    eb62                    jmp 6c5100
'006c509e    8d45d0                  lea eax, dword ptr [ebp-30]
'006c50a1    50                      push eax
'006c50a2    8d4dd4                  lea ecx, dword ptr [ebp-2c]
'006c50a5    51                      push ecx
'006c50a6    8d55d8                  lea edx, dword ptr [ebp-28]
'006c50a9    52                      push edx
'006c50aa    8d45dc                  lea eax, dword ptr [ebp-24]
'006c50ad    50                      push eax
'006c50ae    8d4de0                  lea ecx, dword ptr [ebp-20]
'006c50b1    51                      push ecx
'006c50b2    8d55e4                  lea edx, dword ptr [ebp-1c]
'006c50b5    52                      push edx
'006c50b6    6a06                    push 06

' *** Reference to "__vbaFreeStrList"
'006c50b8    ff155c124000            call dword ptr [0040125c]
'#Cleanup( , 0, 0, -4500, -4504, 0)
'006c50be    8d45bc                  lea eax, dword ptr [ebp-44]
'006c50c1    50                      push eax
'006c50c2    8d4dc0                  lea ecx, dword ptr [ebp-40]
'006c50c5    51                      push ecx
'006c50c6    8d55c4                  lea edx, dword ptr [ebp-3c]
'006c50c9    52                      push edx
'006c50ca    8d45c8                  lea eax, dword ptr [ebp-38]
'006c50cd    50                      push eax
'006c50ce    8d4dcc                  lea ecx, dword ptr [ebp-34]
'006c50d1    51                      push ecx
'006c50d2    6a05                    push 05

' *** Reference to "__vbaFreeObjList"
'006c50d4    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_43, var_46, var_9, var_5, var_58)
'006c50da    8d956cffffff            lea edx, dword ptr [ebp+ffffff6c]
'006c50e0    52                      push edx
'006c50e1    8d857cffffff            lea eax, dword ptr [ebp+ffffff7c]
'006c50e7    50                      push eax
'006c50e8    8d4d8c                  lea ecx, dword ptr [ebp-74]
'006c50eb    51                      push ecx
'006c50ec    8d559c                  lea edx, dword ptr [ebp-64]
'006c50ef    52                      push edx
'006c50f0    8d45ac                  lea eax, dword ptr [ebp-54]
'006c50f3    50                      push eax
'006c50f4    6a05                    push 05

' *** Reference to "__vbaFreeVarList"
'006c50f6    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_50, var_51, var_52, var_59, var_94)
'006c50fc    83c44c                  add esp, 4c
'006c50ff    c3                      ret
'006c5100    8d4de8                  lea ecx, dword ptr [ebp-18]

' *** Reference to "__vbaFreeObj"
'006c5103    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_41)
'006c5109    c3                      ret
'006c510a    8b4508                  mov eax, dword ptr [ebp+08]
'006c510d    8b08                    mov ecx, dword ptr [eax]
'006c510f    50                      push eax
'006c5110    ff5108                  call dword ptr [ecx+08]
'006c5113    8b45fc                  mov eax, dword ptr [ebp-04]
'006c5116    8b4dec                  mov ecx, dword ptr [ebp-14]
'006c5119    5f                      pop edi
'006c511a    5e                      pop esi
    'Reference to '__except_list'
'006c511b    64890d00000000          mov dword ptr fs:[00000000], ecx
'006c5122    5b                      pop ebx
'006c5123    8be5                    mov esp, ebp
'006c5125    5d                      pop ebp
'006c5126    c20400                  ret 0004


End Function


'Event for BtnFermer
Private Sub BtnFermer_Click()
'006c31f0    55                      push ebp
'006c31f1    8bec                    mov ebp, esp
'006c31f3    83ec0c                  sub esp, 0c

' *** Reference to "__vbaExceptHandler"
'006c31f6    6866724000              push 00407266
'006c31fb    64a100000000            mov ax, word ptr fs:[00000000]
'006c3201    50                      push eax
    'Reference to '__except_list'
'006c3202    64892500000000          mov dword ptr fs:[00000000], esp
'006c3209    83ec18                  sub esp, 18
'006c320c    53                      push ebx
'006c320d    56                      push esi
'006c320e    57                      push edi
'006c320f    8965f4                  mov dword ptr [ebp-0c], esp
'006c3212    c745f840664000          mov dword ptr [ebp-08], 00406640
'006c3219    8b7d08                  mov edi, dword ptr [ebp+08]
'006c321c    8bc7                    mov eax, edi
'006c321e    83e001                  and eax, 01
'006c3221    8945fc                  mov dword ptr [ebp-04], eax
'006c3224    83e7fe                  and edi, fffffffe
'006c3227    8b0f                    mov ecx, dword ptr [edi]
'006c3229    57                      push edi
'006c322a    897d08                  mov dword ptr [ebp+08], edi
'006c322d    ff5104                  call dword ptr [ecx+04]
'006c3230    a124be7200              mov ax, word ptr [0072be24]
'006c3235    33db                    xor ebx, ebx
'006c3237    3bc3                    cmp eax, ebx
'006c3239    895de8                  mov dword ptr [ebp-18], ebx
'006c323c    7510                    jne 6c324e

If (0 = 0) Then
'006c323e    6824be7200              push 0072be24
'006c3243    6870004300              push 00430070

' *** Reference to "__vbaNew2"
'006c3248    ff152c124000            call dword ptr [0040122c]
    Dim var_2 As New Global
End If
'006c324e    8b3524be7200            mov esi, dword ptr [0072be24]
'006c3254    8b16                    mov edx, dword ptr [esi]
'006c3256    57                      push edi
'006c3257    8d45e8                  lea eax, dword ptr [ebp-18]
'006c325a    50                      push eax
'006c325b    8955d4                  mov dword ptr [ebp-2c], edx

' *** Reference to "__vbaObjSetAddref"
'006c325e    ff15c8104000            call dword ptr [004010c8]
Set var_41 = Me
'006c3264    8b4dd4                  mov ecx, dword ptr [ebp-2c]
'006c3267    50                      push eax
'006c3268    56                      push esi
'006c3269    ff5110                  call dword ptr [ecx+10]
Call var_2.Unload(var_41)
'006c326c    dbe2                    fnclex
'006c326e    3bc3                    cmp eax, ebx
'006c3270    7d0f                    jge 6c3281
'006c3272    6a10                    push 10
'006c3274    6860004300              push 00430060
'006c3279    56                      push esi
'006c327a    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c327b    ff1580104000            call dword ptr [00401080]
'006c3281    8d4de8                  lea ecx, dword ptr [ebp-18]

' *** Reference to "__vbaFreeObj"
'006c3284    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_41)
'006c328a    895dfc                  mov dword ptr [ebp-04], ebx
'006c328d    689f326c00              push 006c329f
'006c3292    eb0a                    jmp 6c329e
'006c3294    8d4de8                  lea ecx, dword ptr [ebp-18]

' *** Reference to "__vbaFreeObj"
'006c3297    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_41)
'006c329d    c3                      ret
'006c329e    c3                      ret
'006c329f    8b4508                  mov eax, dword ptr [ebp+08]
'006c32a2    8b10                    mov edx, dword ptr [eax]
'006c32a4    50                      push eax
'006c32a5    ff5208                  call dword ptr [edx+08]
'006c32a8    8b45fc                  mov eax, dword ptr [ebp-04]
'006c32ab    8b4dec                  mov ecx, dword ptr [ebp-14]
'006c32ae    5f                      pop edi
'006c32af    5e                      pop esi
    'Reference to '__except_list'
'006c32b0    64890d00000000          mov dword ptr fs:[00000000], ecx
'006c32b7    5b                      pop ebx
'006c32b8    8be5                    mov esp, ebp
'006c32ba    5d                      pop ebp
'006c32bb    c20400                  ret 0004


End Sub


'Event for Form
Private Sub Form_Resize()
'006c5130    55                      push ebp
'006c5131    8bec                    mov ebp, esp
'006c5133    83ec0c                  sub esp, 0c

' *** Reference to "__vbaExceptHandler"
'006c5136    6866724000              push 00407266
'006c513b    64a100000000            mov ax, word ptr fs:[00000000]
'006c5141    50                      push eax
    'Reference to '__except_list'
'006c5142    64892500000000          mov dword ptr fs:[00000000], esp
'006c5149    81ec68010000            sub esp, 00000168
'006c514f    53                      push ebx
'006c5150    56                      push esi
'006c5151    57                      push edi
'006c5152    8965f4                  mov dword ptr [ebp-0c], esp
'006c5155    c745f830674000          mov dword ptr [ebp-08], 00406730
'006c515c    8b7508                  mov esi, dword ptr [ebp+08]
'006c515f    8bc6                    mov eax, esi
'006c5161    83e001                  and eax, 01
'006c5164    8945fc                  mov dword ptr [ebp-04], eax
'006c5167    83e6fe                  and esi, fffffffe
'006c516a    8b0e                    mov ecx, dword ptr [esi]
'006c516c    56                      push esi
'006c516d    897508                  mov dword ptr [ebp+08], esi
'006c5170    ff5104                  call dword ptr [ecx+04]
'006c5173    a1c0b07200              mov ax, word ptr [0072b0c0]
'006c5178    33db                    xor ebx, ebx
'006c517a    3bc3                    cmp eax, ebx
'006c517c    895dd8                  mov dword ptr [ebp-28], ebx
'006c517f    895dd4                  mov dword ptr [ebp-2c], ebx
'006c5182    895dd0                  mov dword ptr [ebp-30], ebx
'006c5185    895dc0                  mov dword ptr [ebp-40], ebx
'006c5188    895db0                  mov dword ptr [ebp-50], ebx
'006c518b    895da0                  mov dword ptr [ebp-60], ebx
'006c518e    895d90                  mov dword ptr [ebp-70], ebx
'006c5191    899d4cffffff            mov dword ptr [ebp+ffffff4c], ebx
'006c5197    899d48ffffff            mov dword ptr [ebp+ffffff48], ebx
'006c519d    899d44ffffff            mov dword ptr [ebp+ffffff44], ebx
'006c51a3    7510                    jne 6c51b5

If (0 = 0) Then
'006c51a5    68c0b07200              push 0072b0c0
'006c51aa    68d0f04000              push 0040f0d0

' *** Reference to "__vbaNew2"
'006c51af    ff152c124000            call dword ptr [0040122c]
    Dim var_31 As New frmDescriptifDon
End If
'006c51b5    8b3dc0b07200            mov edi, dword ptr [0072b0c0]
'006c51bb    8b17                    mov edx, dword ptr [edi]
'006c51bd    8d8548ffffff            lea eax, dword ptr [ebp+ffffff48]
'006c51c3    50                      push eax
'006c51c4    57                      push edi
'006c51c5    ff9280000000            call dword ptr [edx+00000080]
'006c51cb    dbe2                    fnclex
'006c51cd    3bc3                    cmp eax, ebx
'006c51cf    7d12                    jge 6c51e3
'006c51d1    6880000000              push 00000080
'006c51d6    6858124300              push 00431258
'006c51db    57                      push edi
'006c51dc    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c51dd    ff1580104000            call dword ptr [00401080]
'006c51e3    d98548ffffff            fld dword ptr [ebp+ffffff48]
'006c51e9    8d4db0                  lea ecx, dword ptr [ebp-50]
'006c51ec    d8252c674000            fsub dword ptr [0040672c]
'006c51f2    51                      push ecx
'006c51f3    8d55c0                  lea edx, dword ptr [ebp-40]
'006c51f6    52                      push edx
'006c51f7    d95db8                  fstp dword ptr [ebp-48]
'var_61 = (00)
'006c51fa    dfe0                    fnstsw ax
'006c51fc    a80d                    test al, 0d
'006c51fe    0f85df130000            jne 6c65e3
'006c5204    c745b004000000          mov dword ptr [ebp-50], 00000004
'006c520b    c745c801000000          mov dword ptr [ebp-38], 00000001
'006c5212    c745c002000000          mov dword ptr [ebp-40], 00000002
'006c5219    e8d2d4e2ff              call 4f26f0
Call sub_4F26F0()
'006c521e    8d4dc0                  lea ecx, dword ptr [ebp-40]
'006c5221    8945e0                  mov dword ptr [ebp-20], eax
'006c5224    8d45b0                  lea eax, dword ptr [ebp-50]
'006c5227    50                      push eax
'006c5228    51                      push ecx
'006c5229    6a02                    push 02

' *** Reference to "__vbaFreeVarList"
'006c522b    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_5, var_6)
'006c5231    8b16                    mov edx, dword ptr [esi]
'006c5233    83c40c                  add esp, 0c
'006c5236    56                      push esi
'006c5237    ff9208030000            call dword ptr [edx+00000308]

' *** Reference to "__vbaObjSet"
'006c523d    8b3db4104000            mov edi, dword ptr [004010b4]
'006c5243    50                      push eax
'006c5244    8d45d8                  lea eax, dword ptr [ebp-28]
'006c5247    50                      push eax
'006c5248    ffd7                    call edi
Set var_45 = 
'006c524a    898538ffffff            mov dword ptr [ebp+ffffff38], eax
'006c5250    391dc0b07200            cmp dword ptr [0072b0c0], ebx
'006c5256    7510                    jne 6c5268

If (var_31 = 0) Then
'006c5258    68c0b07200              push 0072b0c0
'006c525d    68d0f04000              push 0040f0d0

' *** Reference to "__vbaNew2"
'006c5262    ff152c124000            call dword ptr [0040122c]
    Set var_31 = New frmDescriptifDon
End If
'006c5268    8b1dc0b07200            mov ebx, dword ptr [0072b0c0]
'006c526e    8b0b                    mov ecx, dword ptr [ebx]
'006c5270    8d9548ffffff            lea edx, dword ptr [ebp+ffffff48]
'006c5276    52                      push edx
'006c5277    53                      push ebx
'006c5278    ff9180000000            call dword ptr [ecx+00000080]
'006c527e    dbe2                    fnclex
'006c5280    85c0                    test ax, ax
'006c5282    7d12                    jge 6c5296
'006c5284    6880000000              push 00000080
'006c5289    6858124300              push 00431258
'006c528e    53                      push ebx
'006c528f    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c5290    ff1580104000            call dword ptr [00401080]
'006c5296    d98548ffffff            fld dword ptr [ebp+ffffff48]
'006c529c    8d4db0                  lea ecx, dword ptr [ebp-50]
'006c529f    d82528674000            fsub dword ptr [00406728]
'006c52a5    51                      push ecx
'006c52a6    8d55c0                  lea edx, dword ptr [ebp-40]
'006c52a9    c745b004000000          mov dword ptr [ebp-50], 00000004
'006c52b0    d95db8                  fstp dword ptr [ebp-48]
'var_61 = (00)
'006c52b3    dfe0                    fnstsw ax
'006c52b5    a80d                    test al, 0d
'006c52b7    0f8526130000            jne 6c65e3
'006c52bd    8b8538ffffff            mov eax, dword ptr [ebp+ffffff38]
'006c52c3    c745c801000000          mov dword ptr [ebp-38], 00000001
'006c52ca    c745c002000000          mov dword ptr [ebp-40], 00000002
'006c52d1    8b18                    mov ebx, dword ptr [eax]
'006c52d3    52                      push edx
'006c52d4    e817d4e2ff              call 4f26f0
Call sub_4F26F0()
'006c52d9    0fbfc0                  movsx eax, ax
'006c52dc    898520ffffff            mov dword ptr [ebp+ffffff20], eax
'006c52e2    8bd3                    mov edx, ebx
'006c52e4    8b9d38ffffff            mov ebx, dword ptr [ebp+ffffff38]
'006c52ea    db8520ffffff            fild dword ptr [ebp+ffffff20]
'006c52f0    d99d1cffffff            fstp dword ptr [ebp+ffffff1c]
'var_95 = (00)
'006c52f6    8b8d1cffffff            mov ecx, dword ptr [ebp+ffffff1c]
'006c52fc    51                      push ecx
'006c52fd    53                      push ebx
'006c52fe    ff9284000000            call dword ptr [edx+00000084]
'006c5304    dbe2                    fnclex
'006c5306    85c0                    test ax, ax
'006c5308    7d16                    jge 6c5320
'006c530a    6884000000              push 00000084
'006c530f    6874f14300              push 0043f174
'006c5314    53                      push ebx

' *** Reference to "__vbaHresultCheckObj"
'006c5315    8b1d80104000            mov ebx, dword ptr [00401080]
'006c531b    50                      push eax
'006c531c    ffd3                    call ebx
'006c531e    eb06                    jmp 6c5326

' *** Reference to "__vbaHresultCheckObj"
'006c5320    8b1d80104000            mov ebx, dword ptr [00401080]
'006c5326    8d4dd8                  lea ecx, dword ptr [ebp-28]

' *** Reference to "__vbaFreeObj"
'006c5329    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_45)
'006c532f    8d45b0                  lea eax, dword ptr [ebp-50]
'006c5332    50                      push eax
'006c5333    8d4dc0                  lea ecx, dword ptr [ebp-40]
'006c5336    51                      push ecx
'006c5337    6a02                    push 02

' *** Reference to "__vbaFreeVarList"
'006c5339    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_5, var_6)
'006c533f    8b16                    mov edx, dword ptr [esi]
'006c5341    83c40c                  add esp, 0c
'006c5344    56                      push esi
'006c5345    ff920c030000            call dword ptr [edx+0000030c]
'006c534b    50                      push eax
'006c534c    8d45d8                  lea eax, dword ptr [ebp-28]
'006c534f    50                      push eax
'006c5350    ffd7                    call edi
Set var_45 = 
'006c5352    0fbf4de0                movsx ecx, word ptr [ebp-20]
'006c5356    8b10                    mov edx, dword ptr [eax]
'006c5358    898d14ffffff            mov dword ptr [ebp+ffffff14], ecx
'006c535e    898540ffffff            mov dword ptr [ebp+ffffff40], eax
'006c5364    db8514ffffff            fild dword ptr [ebp+ffffff14]
'006c536a    d99d10ffffff            fstp dword ptr [ebp+ffffff10]
'var_1494 = (00)
'006c5370    8b8d10ffffff            mov ecx, dword ptr [ebp+ffffff10]
'006c5376    51                      push ecx
'006c5377    50                      push eax
'006c5378    ff527c                  call dword ptr [edx+7c]
'006c537b    dbe2                    fnclex
'006c537d    85c0                    test ax, ax
'006c537f    7d11                    jge 6c5392
'006c5381    8b9540ffffff            mov edx, dword ptr [ebp+ffffff40]
'006c5387    6a7c                    push 7c
'006c5389    68d00d4300              push 00430dd0
'006c538e    52                      push edx
'006c538f    50                      push eax
'006c5390    ffd3                    call ebx
'006c5392    8d4dd8                  lea ecx, dword ptr [ebp-28]

' *** Reference to "__vbaFreeObj"
'006c5395    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_45)
'006c539b    8b06                    mov eax, dword ptr [esi]
'006c539d    56                      push esi
'006c539e    ff9018030000            call dword ptr [eax+00000318]
'006c53a4    50                      push eax
'006c53a5    8d4dd8                  lea ecx, dword ptr [ebp-28]
'006c53a8    51                      push ecx
'006c53a9    ffd7                    call edi
Set var_45 = Nothing
'006c53ab    db8514ffffff            fild dword ptr [ebp+ffffff14]
'006c53b1    8b10                    mov edx, dword ptr [eax]
'006c53b3    898540ffffff            mov dword ptr [ebp+ffffff40], eax
'006c53b9    d99d0cffffff            fstp dword ptr [ebp+ffffff0c]
'var_116 = (00)
'006c53bf    8b8d0cffffff            mov ecx, dword ptr [ebp+ffffff0c]
'006c53c5    51                      push ecx
'006c53c6    50                      push eax
'006c53c7    ff527c                  call dword ptr [edx+7c]
'006c53ca    dbe2                    fnclex
'006c53cc    85c0                    test ax, ax
'006c53ce    7d11                    jge 6c53e1
'006c53d0    8b9540ffffff            mov edx, dword ptr [ebp+ffffff40]
'006c53d6    6a7c                    push 7c
'006c53d8    68d00d4300              push 00430dd0
'006c53dd    52                      push edx
'006c53de    50                      push eax
'006c53df    ffd3                    call ebx
'006c53e1    8d4dd8                  lea ecx, dword ptr [ebp-28]

' *** Reference to "__vbaFreeObj"
'006c53e4    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_45)
'006c53ea    8b06                    mov eax, dword ptr [esi]
'006c53ec    56                      push esi
'006c53ed    ff901c030000            call dword ptr [eax+0000031c]
'006c53f3    50                      push eax
'006c53f4    8d4dd8                  lea ecx, dword ptr [ebp-28]
'006c53f7    51                      push ecx
'006c53f8    ffd7                    call edi
Set var_45 = Nothing
'006c53fa    db8514ffffff            fild dword ptr [ebp+ffffff14]
'006c5400    8b10                    mov edx, dword ptr [eax]
'006c5402    898540ffffff            mov dword ptr [ebp+ffffff40], eax
'006c5408    d99d08ffffff            fstp dword ptr [ebp+ffffff08]
'var_310 = (00)
'006c540e    8b8d08ffffff            mov ecx, dword ptr [ebp+ffffff08]
'006c5414    51                      push ecx
'006c5415    50                      push eax
'006c5416    ff527c                  call dword ptr [edx+7c]
'006c5419    dbe2                    fnclex
'006c541b    85c0                    test ax, ax
'006c541d    7d11                    jge 6c5430
'006c541f    8b9540ffffff            mov edx, dword ptr [ebp+ffffff40]
'006c5425    6a7c                    push 7c
'006c5427    68d00d4300              push 00430dd0
'006c542c    52                      push edx
'006c542d    50                      push eax
'006c542e    ffd3                    call ebx
'006c5430    8d4dd8                  lea ecx, dword ptr [ebp-28]

' *** Reference to "__vbaFreeObj"
'006c5433    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_45)
'006c5439    8b06                    mov eax, dword ptr [esi]
'006c543b    56                      push esi
'006c543c    ff9014030000            call dword ptr [eax+00000314]
'006c5442    50                      push eax
'006c5443    8d4dd8                  lea ecx, dword ptr [ebp-28]
'006c5446    51                      push ecx
'006c5447    ffd7                    call edi
Set var_45 = Nothing
'006c5449    db8514ffffff            fild dword ptr [ebp+ffffff14]
'006c544f    8b10                    mov edx, dword ptr [eax]
'006c5451    898540ffffff            mov dword ptr [ebp+ffffff40], eax
'006c5457    d99d04ffffff            fstp dword ptr [ebp+ffffff04]
'var_458 = (00)
'006c545d    8b8d04ffffff            mov ecx, dword ptr [ebp+ffffff04]
'006c5463    51                      push ecx
'006c5464    50                      push eax
'006c5465    ff527c                  call dword ptr [edx+7c]
'006c5468    dbe2                    fnclex
'006c546a    85c0                    test ax, ax
'006c546c    7d11                    jge 6c547f
'006c546e    8b9540ffffff            mov edx, dword ptr [ebp+ffffff40]
'006c5474    6a7c                    push 7c
'006c5476    68d00d4300              push 00430dd0
'006c547b    52                      push edx
'006c547c    50                      push eax
'006c547d    ffd3                    call ebx
'006c547f    8d4dd8                  lea ecx, dword ptr [ebp-28]

' *** Reference to "__vbaFreeObj"
'006c5482    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_45)
'006c5488    8b06                    mov eax, dword ptr [esi]
'006c548a    56                      push esi
'006c548b    ff9010030000            call dword ptr [eax+00000310]
'006c5491    50                      push eax
'006c5492    8d4dd8                  lea ecx, dword ptr [ebp-28]
'006c5495    51                      push ecx
'006c5496    ffd7                    call edi
Set var_45 = Nothing
'006c5498    db8514ffffff            fild dword ptr [ebp+ffffff14]
'006c549e    8b10                    mov edx, dword ptr [eax]
'006c54a0    898540ffffff            mov dword ptr [ebp+ffffff40], eax
'006c54a6    d99d00ffffff            fstp dword ptr [ebp+ffffff00]
'var_444 = (00)
'006c54ac    8b8d00ffffff            mov ecx, dword ptr [ebp+ffffff00]
'006c54b2    51                      push ecx
'006c54b3    50                      push eax
'006c54b4    ff527c                  call dword ptr [edx+7c]
'006c54b7    dbe2                    fnclex
'006c54b9    85c0                    test ax, ax
'006c54bb    7d11                    jge 6c54ce
'006c54bd    8b9540ffffff            mov edx, dword ptr [ebp+ffffff40]
'006c54c3    6a7c                    push 7c
'006c54c5    68d00d4300              push 00430dd0
'006c54ca    52                      push edx
'006c54cb    50                      push eax
'006c54cc    ffd3                    call ebx
'006c54ce    8d4dd8                  lea ecx, dword ptr [ebp-28]

' *** Reference to "__vbaFreeObj"
'006c54d1    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_45)
'006c54d7    8b06                    mov eax, dword ptr [esi]
'006c54d9    56                      push esi
'006c54da    ff903c030000            call dword ptr [eax+0000033c]
'006c54e0    50                      push eax
'006c54e1    8d4dd8                  lea ecx, dword ptr [ebp-28]
'006c54e4    51                      push ecx
'006c54e5    ffd7                    call edi
Set var_45 = Nothing
'006c54e7    db8514ffffff            fild dword ptr [ebp+ffffff14]
'006c54ed    8bc8                    mov ecx, eax
'006c54ef    8b11                    mov edx, dword ptr [ecx]
'006c54f1    dd9df8feffff            fstp qword ptr [ebp+fffffef8]
'var_298 = (00)
'006c54f7    51                      push ecx
'006c54f8    dd85f8feffff            fld qword ptr [ebp+fffffef8]
'006c54fe    898d40ffffff            mov dword ptr [ebp+ffffff40], ecx
'006c5504    dc0d20674000            fmul qword ptr [00406720]
'006c550a    dfe0                    fnstsw ax
'006c550c    a80d                    test al, 0d
'006c550e    0f85cf100000            jne 6c65e3
'006c5514    d99df4feffff            fstp dword ptr [ebp+fffffef4]
'var_824 = (00)
'006c551a    d985f4feffff            fld dword ptr [ebp+fffffef4]
'006c5520    d91c24                  fstp dword ptr [esp]
'var_483 = (00)
'006c5523    51                      push ecx
'006c5524    ff9284000000            call dword ptr [edx+00000084]
'006c552a    dbe2                    fnclex
'006c552c    85c0                    test ax, ax
'006c552e    7d14                    jge 6c5544
'006c5530    8b8d40ffffff            mov ecx, dword ptr [ebp+ffffff40]
'006c5536    6884000000              push 00000084
'006c553b    683ce04300              push 0043e03c
'006c5540    51                      push ecx
'006c5541    50                      push eax
'006c5542    ffd3                    call ebx
'006c5544    8d4dd8                  lea ecx, dword ptr [ebp-28]

' *** Reference to "__vbaFreeObj"
'006c5547    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_45)
'006c554d    8b16                    mov edx, dword ptr [esi]
'006c554f    56                      push esi
'006c5550    ff9238030000            call dword ptr [edx+00000338]
'006c5556    50                      push eax
'006c5557    8d45d8                  lea eax, dword ptr [ebp-28]
'006c555a    50                      push eax
'006c555b    ffd7                    call edi
Set var_45 = Nothing
'006c555d    db8514ffffff            fild dword ptr [ebp+ffffff14]
'006c5563    8bc8                    mov ecx, eax
'006c5565    8b11                    mov edx, dword ptr [ecx]
'006c5567    dd9decfeffff            fstp qword ptr [ebp+fffffeec]
'var_118 = (00)
'006c556d    51                      push ecx
'006c556e    dd85ecfeffff            fld qword ptr [ebp+fffffeec]
'006c5574    898d40ffffff            mov dword ptr [ebp+ffffff40], ecx
'006c557a    dc0d18674000            fmul qword ptr [00406718]
'006c5580    dfe0                    fnstsw ax
'006c5582    a80d                    test al, 0d
'006c5584    0f8559100000            jne 6c65e3
'006c558a    d99de8feffff            fstp dword ptr [ebp+fffffee8]
'var_539 = (00)
'006c5590    d985e8feffff            fld dword ptr [ebp+fffffee8]
'006c5596    d91c24                  fstp dword ptr [esp]
'var_518 = (00)
'006c5599    51                      push ecx
'006c559a    ff9284000000            call dword ptr [edx+00000084]
'006c55a0    dbe2                    fnclex
'006c55a2    85c0                    test ax, ax
'006c55a4    7d14                    jge 6c55ba
'006c55a6    8b8d40ffffff            mov ecx, dword ptr [ebp+ffffff40]
'006c55ac    6884000000              push 00000084
'006c55b1    683ce04300              push 0043e03c
'006c55b6    51                      push ecx
'006c55b7    50                      push eax
'006c55b8    ffd3                    call ebx
'006c55ba    8d4dd8                  lea ecx, dword ptr [ebp-28]

' *** Reference to "__vbaFreeObj"
'006c55bd    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_45)
'006c55c3    8b16                    mov edx, dword ptr [esi]
'006c55c5    56                      push esi
'006c55c6    ff9234030000            call dword ptr [edx+00000334]
'006c55cc    50                      push eax
'006c55cd    8d45d4                  lea eax, dword ptr [ebp-2c]
'006c55d0    50                      push eax
'006c55d1    ffd7                    call edi
Set var_44 = Nothing
'006c55d3    8b0e                    mov ecx, dword ptr [esi]
'006c55d5    56                      push esi
'006c55d6    898538ffffff            mov dword ptr [ebp+ffffff38], eax
'006c55dc    ff9138030000            call dword ptr [ecx+00000338]
'006c55e2    50                      push eax
'006c55e3    8d55d8                  lea edx, dword ptr [ebp-28]
'006c55e6    52                      push edx
'006c55e7    ffd7                    call edi
Set var_45 = Nothing
'006c55e9    8b08                    mov ecx, dword ptr [eax]
'006c55eb    8d9548ffffff            lea edx, dword ptr [ebp+ffffff48]
'006c55f1    52                      push edx
'006c55f2    50                      push eax
'006c55f3    898540ffffff            mov dword ptr [ebp+ffffff40], eax
'006c55f9    ff9180000000            call dword ptr [ecx+00000080]
'006c55ff    dbe2                    fnclex
'006c5601    85c0                    test ax, ax
'006c5603    7d14                    jge 6c5619
'006c5605    8b8d40ffffff            mov ecx, dword ptr [ebp+ffffff40]
'006c560b    6880000000              push 00000080
'006c5610    683ce04300              push 0043e03c
'006c5615    51                      push ecx
'006c5616    50                      push eax
'006c5617    ffd3                    call ebx
'006c5619    8b8d48ffffff            mov ecx, dword ptr [ebp+ffffff48]
'006c561f    8b8538ffffff            mov eax, dword ptr [ebp+ffffff38]
'006c5625    8b10                    mov edx, dword ptr [eax]
'006c5627    51                      push ecx
'006c5628    50                      push eax
'006c5629    ff9284000000            call dword ptr [edx+00000084]
'006c562f    dbe2                    fnclex
'006c5631    85c0                    test ax, ax
'006c5633    7d14                    jge 6c5649
'006c5635    8b9538ffffff            mov edx, dword ptr [ebp+ffffff38]
'006c563b    6884000000              push 00000084
'006c5640    683ce04300              push 0043e03c
'006c5645    52                      push edx
'006c5646    50                      push eax
'006c5647    ffd3                    call ebx
'006c5649    8d45d4                  lea eax, dword ptr [ebp-2c]
'006c564c    50                      push eax
'006c564d    8d4dd8                  lea ecx, dword ptr [ebp-28]
'006c5650    51                      push ecx
'006c5651    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'006c5653    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_45, var_44)
'006c5659    8b16                    mov edx, dword ptr [esi]
'006c565b    83c40c                  add esp, 0c
'006c565e    56                      push esi
'006c565f    ff922c030000            call dword ptr [edx+0000032c]
'006c5665    50                      push eax
'006c5666    8d45d8                  lea eax, dword ptr [ebp-28]
'006c5669    50                      push eax
'006c566a    ffd7                    call edi
Set var_45 = 
'006c566c    db8514ffffff            fild dword ptr [ebp+ffffff14]
'006c5672    8bc8                    mov ecx, eax
'006c5674    8b11                    mov edx, dword ptr [ecx]
'006c5676    dd9de0feffff            fstp qword ptr [ebp+fffffee0]
'var_427 = (00)
'006c567c    51                      push ecx
'006c567d    dd85e0feffff            fld qword ptr [ebp+fffffee0]
'006c5683    898d40ffffff            mov dword ptr [ebp+ffffff40], ecx
'006c5689    dc0d10674000            fmul qword ptr [00406710]
'006c568f    dfe0                    fnstsw ax
'006c5691    a80d                    test al, 0d
'006c5693    0f854a0f0000            jne 6c65e3
'006c5699    d99ddcfeffff            fstp dword ptr [ebp+fffffedc]
'var_10 = (00)
'006c569f    d985dcfeffff            fld dword ptr [ebp+fffffedc]
'006c56a5    d91c24                  fstp dword ptr [esp]
'var_518 = (00)
'006c56a8    51                      push ecx
'006c56a9    ff9284000000            call dword ptr [edx+00000084]
'006c56af    dbe2                    fnclex
'006c56b1    85c0                    test ax, ax
'006c56b3    7d14                    jge 6c56c9
'006c56b5    8b8d40ffffff            mov ecx, dword ptr [ebp+ffffff40]
'006c56bb    6884000000              push 00000084
'006c56c0    683ce04300              push 0043e03c
'006c56c5    51                      push ecx
'006c56c6    50                      push eax
'006c56c7    ffd3                    call ebx
'006c56c9    8d4dd8                  lea ecx, dword ptr [ebp-28]

' *** Reference to "__vbaFreeObj"
'006c56cc    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_45)
'006c56d2    8b16                    mov edx, dword ptr [esi]
'006c56d4    56                      push esi
'006c56d5    ff9230030000            call dword ptr [edx+00000330]
'006c56db    50                      push eax
'006c56dc    8d45d4                  lea eax, dword ptr [ebp-2c]
'006c56df    50                      push eax
'006c56e0    ffd7                    call edi
Set var_44 = 
'006c56e2    8b0e                    mov ecx, dword ptr [esi]
'006c56e4    56                      push esi
'006c56e5    898538ffffff            mov dword ptr [ebp+ffffff38], eax
'006c56eb    ff9138030000            call dword ptr [ecx+00000338]
'006c56f1    50                      push eax
'006c56f2    8d55d8                  lea edx, dword ptr [ebp-28]
'006c56f5    52                      push edx
'006c56f6    ffd7                    call edi
Set var_45 = var_44
'006c56f8    8b08                    mov ecx, dword ptr [eax]
'006c56fa    8d9548ffffff            lea edx, dword ptr [ebp+ffffff48]
'006c5700    52                      push edx
'006c5701    50                      push eax
'006c5702    898540ffffff            mov dword ptr [ebp+ffffff40], eax
'006c5708    ff9180000000            call dword ptr [ecx+00000080]
'006c570e    dbe2                    fnclex
'006c5710    85c0                    test ax, ax
'006c5712    7d14                    jge 6c5728
'006c5714    8b8d40ffffff            mov ecx, dword ptr [ebp+ffffff40]
'006c571a    6880000000              push 00000080
'006c571f    683ce04300              push 0043e03c
'006c5724    51                      push ecx
'006c5725    50                      push eax
'006c5726    ffd3                    call ebx
'006c5728    8b8d48ffffff            mov ecx, dword ptr [ebp+ffffff48]
'006c572e    8b8538ffffff            mov eax, dword ptr [ebp+ffffff38]
'006c5734    8b10                    mov edx, dword ptr [eax]
'006c5736    51                      push ecx
'006c5737    50                      push eax
'006c5738    ff9284000000            call dword ptr [edx+00000084]
'006c573e    dbe2                    fnclex
'006c5740    85c0                    test ax, ax
'006c5742    7d14                    jge 6c5758
'006c5744    8b9538ffffff            mov edx, dword ptr [ebp+ffffff38]
'006c574a    6884000000              push 00000084
'006c574f    683ce04300              push 0043e03c
'006c5754    52                      push edx
'006c5755    50                      push eax
'006c5756    ffd3                    call ebx
'006c5758    8d45d4                  lea eax, dword ptr [ebp-2c]
'006c575b    50                      push eax
'006c575c    8d4dd8                  lea ecx, dword ptr [ebp-28]
'006c575f    51                      push ecx
'006c5760    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'006c5762    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_45, var_44)
'006c5768    8b16                    mov edx, dword ptr [esi]
'006c576a    83c40c                  add esp, 0c
'006c576d    56                      push esi
'006c576e    ff9238030000            call dword ptr [edx+00000338]
'006c5774    50                      push eax
'006c5775    8d45d4                  lea eax, dword ptr [ebp-2c]
'006c5778    50                      push eax
'006c5779    ffd7                    call edi
Set var_44 = 
'006c577b    8b0e                    mov ecx, dword ptr [esi]
'006c577d    56                      push esi
'006c577e    898538ffffff            mov dword ptr [ebp+ffffff38], eax
'006c5784    ff913c030000            call dword ptr [ecx+0000033c]
'006c578a    50                      push eax
'006c578b    8d55d8                  lea edx, dword ptr [ebp-28]
'006c578e    52                      push edx
'006c578f    ffd7                    call edi
Set var_45 = var_44
'006c5791    8b08                    mov ecx, dword ptr [eax]
'006c5793    8d9548ffffff            lea edx, dword ptr [ebp+ffffff48]
'006c5799    52                      push edx
'006c579a    50                      push eax
'006c579b    898540ffffff            mov dword ptr [ebp+ffffff40], eax
'006c57a1    ff9180000000            call dword ptr [ecx+00000080]
'006c57a7    dbe2                    fnclex
'006c57a9    85c0                    test ax, ax
'006c57ab    7d14                    jge 6c57c1
'006c57ad    8b8d40ffffff            mov ecx, dword ptr [ebp+ffffff40]
'006c57b3    6880000000              push 00000080
'006c57b8    683ce04300              push 0043e03c
'006c57bd    51                      push ecx
'006c57be    50                      push eax
'006c57bf    ffd3                    call ebx
'006c57c1    d98548ffffff            fld dword ptr [ebp+ffffff48]
'006c57c7    8b8d38ffffff            mov ecx, dword ptr [ebp+ffffff38]
'006c57cd    d80508674000            fadd dword ptr [00406708]
'006c57d3    8b11                    mov edx, dword ptr [ecx]
'006c57d5    51                      push ecx
'006c57d6    dfe0                    fnstsw ax
'006c57d8    a80d                    test al, 0d
'006c57da    0f85030e0000            jne 6c65e3
'006c57e0    d91c24                  fstp dword ptr [esp]
'var_376 = (00)
'006c57e3    51                      push ecx
'006c57e4    ff5274                  call dword ptr [edx+74]
'006c57e7    dbe2                    fnclex
'006c57e9    85c0                    test ax, ax
'006c57eb    7d11                    jge 6c57fe
'006c57ed    8b8d38ffffff            mov ecx, dword ptr [ebp+ffffff38]
'006c57f3    6a74                    push 74
'006c57f5    683ce04300              push 0043e03c
'006c57fa    51                      push ecx
'006c57fb    50                      push eax
'006c57fc    ffd3                    call ebx
'006c57fe    8d55d4                  lea edx, dword ptr [ebp-2c]
'006c5801    52                      push edx
'006c5802    8d45d8                  lea eax, dword ptr [ebp-28]
'006c5805    50                      push eax
'006c5806    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'006c5808    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_45, var_44)
'006c580e    8b0e                    mov ecx, dword ptr [esi]
'006c5810    83c40c                  add esp, 0c
'006c5813    56                      push esi
'006c5814    ff912c030000            call dword ptr [ecx+0000032c]
'006c581a    50                      push eax
'006c581b    8d55d4                  lea edx, dword ptr [ebp-2c]
'006c581e    52                      push edx
'006c581f    ffd7                    call edi
Set var_44 = 
'006c5821    898538ffffff            mov dword ptr [ebp+ffffff38], eax
'006c5827    8b06                    mov eax, dword ptr [esi]
'006c5829    56                      push esi
'006c582a    ff9034030000            call dword ptr [eax+00000334]
'006c5830    50                      push eax
'006c5831    8d4dd8                  lea ecx, dword ptr [ebp-28]
'006c5834    51                      push ecx
'006c5835    ffd7                    call edi
Set var_45 = Nothing
'006c5837    8b10                    mov edx, dword ptr [eax]
'006c5839    8d8d48ffffff            lea ecx, dword ptr [ebp+ffffff48]
'006c583f    51                      push ecx
'006c5840    50                      push eax
'006c5841    898540ffffff            mov dword ptr [ebp+ffffff40], eax
'006c5847    ff9280000000            call dword ptr [edx+00000080]
'006c584d    dbe2                    fnclex
'006c584f    85c0                    test ax, ax
'006c5851    7d14                    jge 6c5867
'006c5853    8b9540ffffff            mov edx, dword ptr [ebp+ffffff40]
'006c5859    6880000000              push 00000080
'006c585e    683ce04300              push 0043e03c
'006c5863    52                      push edx
'006c5864    50                      push eax
'006c5865    ffd3                    call ebx
'006c5867    d98548ffffff            fld dword ptr [ebp+ffffff48]
'006c586d    8b8d38ffffff            mov ecx, dword ptr [ebp+ffffff38]
'006c5873    d80504674000            fadd dword ptr [00406704]
'006c5879    8b11                    mov edx, dword ptr [ecx]
'006c587b    51                      push ecx
'006c587c    dfe0                    fnstsw ax
'006c587e    a80d                    test al, 0d
'006c5880    0f855d0d0000            jne 6c65e3
'006c5886    d91c24                  fstp dword ptr [esp]
'var_518 = (00)
'006c5889    51                      push ecx
'006c588a    ff5274                  call dword ptr [edx+74]
'006c588d    dbe2                    fnclex
'006c588f    85c0                    test ax, ax
'006c5891    7d11                    jge 6c58a4
'006c5893    8b8d38ffffff            mov ecx, dword ptr [ebp+ffffff38]
'006c5899    6a74                    push 74
'006c589b    683ce04300              push 0043e03c
'006c58a0    51                      push ecx
'006c58a1    50                      push eax
'006c58a2    ffd3                    call ebx
'006c58a4    8d55d4                  lea edx, dword ptr [ebp-2c]
'006c58a7    52                      push edx
'006c58a8    8d45d8                  lea eax, dword ptr [ebp-28]
'006c58ab    50                      push eax
'006c58ac    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'006c58ae    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_45, var_44)
'006c58b4    8b0e                    mov ecx, dword ptr [esi]
'006c58b6    83c40c                  add esp, 0c
'006c58b9    56                      push esi
'006c58ba    ff912c030000            call dword ptr [ecx+0000032c]
'006c58c0    50                      push eax
'006c58c1    8d55d4                  lea edx, dword ptr [ebp-2c]
'006c58c4    52                      push edx
'006c58c5    ffd7                    call edi
Set var_44 = 
'006c58c7    8b08                    mov ecx, dword ptr [eax]
'006c58c9    8d9544ffffff            lea edx, dword ptr [ebp+ffffff44]
'006c58cf    52                      push edx
'006c58d0    50                      push eax
'006c58d1    898540ffffff            mov dword ptr [ebp+ffffff40], eax
'006c58d7    ff9180000000            call dword ptr [ecx+00000080]
'006c58dd    dbe2                    fnclex
'006c58df    85c0                    test ax, ax
'006c58e1    7d14                    jge 6c58f7
'006c58e3    8b8d40ffffff            mov ecx, dword ptr [ebp+ffffff40]
'006c58e9    6880000000              push 00000080
'006c58ee    683ce04300              push 0043e03c
'006c58f3    51                      push ecx
'006c58f4    50                      push eax
'006c58f5    ffd3                    call ebx
'006c58f7    8b16                    mov edx, dword ptr [esi]
'006c58f9    56                      push esi
'006c58fa    ff9230030000            call dword ptr [edx+00000330]
'006c5900    50                      push eax
'006c5901    8d45d0                  lea eax, dword ptr [ebp-30]
'006c5904    50                      push eax
'006c5905    ffd7                    call edi
Set var_4 = var_44
'006c5907    8b0e                    mov ecx, dword ptr [esi]
'006c5909    56                      push esi
'006c590a    898530ffffff            mov dword ptr [ebp+ffffff30], eax
'006c5910    ff912c030000            call dword ptr [ecx+0000032c]
'006c5916    50                      push eax
'006c5917    8d55d8                  lea edx, dword ptr [ebp-28]
'006c591a    52                      push edx
'006c591b    ffd7                    call edi
Set var_45 = var_4
'006c591d    8b08                    mov ecx, dword ptr [eax]
'006c591f    8d9548ffffff            lea edx, dword ptr [ebp+ffffff48]
'006c5925    52                      push edx
'006c5926    50                      push eax
'006c5927    898538ffffff            mov dword ptr [ebp+ffffff38], eax
'006c592d    ff5170                  call dword ptr [ecx+70]
'006c5930    dbe2                    fnclex
'006c5932    85c0                    test ax, ax
'006c5934    7d11                    jge 6c5947
'006c5936    8b8d38ffffff            mov ecx, dword ptr [ebp+ffffff38]
'006c593c    6a70                    push 70
'006c593e    683ce04300              push 0043e03c
'006c5943    51                      push ecx
'006c5944    50                      push eax
'006c5945    ffd3                    call ebx
'006c5947    d98544ffffff            fld dword ptr [ebp+ffffff44]
'006c594d    8b8d30ffffff            mov ecx, dword ptr [ebp+ffffff30]
'006c5953    d88548ffffff            fadd dword ptr [ebp+ffffff48]
'006c5959    8b11                    mov edx, dword ptr [ecx]
'006c595b    51                      push ecx
'006c595c    d80500674000            fadd dword ptr [00406700]
'006c5962    dfe0                    fnstsw ax
'006c5964    a80d                    test al, 0d
'006c5966    0f85770c0000            jne 6c65e3
'006c596c    d91c24                  fstp dword ptr [esp]
'var_518 = (00)
'006c596f    51                      push ecx
'006c5970    ff5274                  call dword ptr [edx+74]
'006c5973    dbe2                    fnclex
'006c5975    85c0                    test ax, ax
'006c5977    7d11                    jge 6c598a
'006c5979    8b8d30ffffff            mov ecx, dword ptr [ebp+ffffff30]
'006c597f    6a74                    push 74
'006c5981    683ce04300              push 0043e03c
'006c5986    51                      push ecx
'006c5987    50                      push eax
'006c5988    ffd3                    call ebx
'006c598a    8d55d0                  lea edx, dword ptr [ebp-30]
'006c598d    52                      push edx
'006c598e    8d45d4                  lea eax, dword ptr [ebp-2c]
'006c5991    50                      push eax
'006c5992    8d4dd8                  lea ecx, dword ptr [ebp-28]
'006c5995    51                      push ecx
'006c5996    6a03                    push 03

' *** Reference to "__vbaFreeObjList"
'006c5998    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_45, var_44, var_4)
'006c599e    a1c0b07200              mov ax, word ptr [0072b0c0]
'006c59a3    83c410                  add esp, 10
'006c59a6    85c0                    test ax, ax
'006c59a8    7510                    jne 6c59ba
'006c59aa    68c0b07200              push 0072b0c0
'006c59af    68d0f04000              push 0040f0d0

' *** Reference to "__vbaNew2"
'006c59b4    ff152c124000            call dword ptr [0040122c]
Set var_31 = New frmDescriptifDon
'006c59ba    a1c0b07200              mov ax, word ptr [0072b0c0]
'006c59bf    8b10                    mov edx, dword ptr [eax]
'006c59c1    8d8d4cffffff            lea ecx, dword ptr [ebp+ffffff4c]
'006c59c7    51                      push ecx
'006c59c8    50                      push eax
'006c59c9    898540ffffff            mov dword ptr [ebp+ffffff40], eax
'006c59cf    ff9298000000            call dword ptr [edx+00000098]
'006c59d5    dbe2                    fnclex
'006c59d7    85c0                    test ax, ax
'006c59d9    7d14                    jge 6c59ef
'006c59db    8b9540ffffff            mov edx, dword ptr [ebp+ffffff40]
'006c59e1    6898000000              push 00000098
'006c59e6    6858124300              push 00431258
'006c59eb    52                      push edx
'006c59ec    50                      push eax
'006c59ed    ffd3                    call ebx
'006c59ef    6683bd4cffffff02        cmp word ptr [ebp+ffffff4c], 02
'006c59f7    0f848e010000            je 6c5b8b

If (0 <> 2) Then
'006c59fd    a110b07200              mov ax, word ptr [0072b010]
'006c5a02    85c0                    test ax, ax
'006c5a04    7510                    jne 6c5a16
'006c5a06    6810b07200              push 0072b010
'006c5a0b    68b0e54000              push 0040e5b0

' *** Reference to "__vbaNew2"
'006c5a10    ff152c124000            call dword ptr [0040122c]
    Dim var_60 As New frmMain
'006c5a16    a110b07200              mov ax, word ptr [0072b010]
'006c5a1b    8b08                    mov ecx, dword ptr [eax]
'006c5a1d    8d9548ffffff            lea edx, dword ptr [ebp+ffffff48]
'006c5a23    52                      push edx
'006c5a24    50                      push eax
'006c5a25    898540ffffff            mov dword ptr [ebp+ffffff40], eax
'006c5a2b    ff9188000000            call dword ptr [ecx+00000088]
'006c5a31    dbe2                    fnclex
'006c5a33    85c0                    test ax, ax
'006c5a35    7d14                    jge 6c5a4b
'006c5a37    8b8d40ffffff            mov ecx, dword ptr [ebp+ffffff40]
'006c5a3d    6888000000              push 00000088
'006c5a42    684cfd4200              push 0042fd4c
'006c5a47    51                      push ecx
'006c5a48    50                      push eax
'006c5a49    ffd3                    call ebx
'006c5a4b    a1c0b07200              mov ax, word ptr [0072b0c0]
'006c5a50    85c0                    test ax, ax
'006c5a52    7510                    jne 6c5a64
'006c5a54    68c0b07200              push 0072b0c0
'006c5a59    68d0f04000              push 0040f0d0

' *** Reference to "__vbaNew2"
'006c5a5e    ff152c124000            call dword ptr [0040122c]
    Set var_31 = New frmDescriptifDon
'006c5a64    a1c0b07200              mov ax, word ptr [0072b0c0]
'006c5a69    8b10                    mov edx, dword ptr [eax]
'006c5a6b    8d8d44ffffff            lea ecx, dword ptr [ebp+ffffff44]
'006c5a71    51                      push ecx
'006c5a72    50                      push eax
'006c5a73    898538ffffff            mov dword ptr [ebp+ffffff38], eax
'006c5a79    ff9288000000            call dword ptr [edx+00000088]
'006c5a7f    dbe2                    fnclex
'006c5a81    85c0                    test ax, ax
'006c5a83    7d14                    jge 6c5a99
'006c5a85    8b9538ffffff            mov edx, dword ptr [ebp+ffffff38]
'006c5a8b    6888000000              push 00000088
'006c5a90    6858124300              push 00431258
'006c5a95    52                      push edx
'006c5a96    50                      push eax
'006c5a97    ffd3                    call ebx
'006c5a99    a1c0b07200              mov ax, word ptr [0072b0c0]
'006c5a9e    85c0                    test ax, ax
'006c5aa0    7510                    jne 6c5ab2
'006c5aa2    68c0b07200              push 0072b0c0
'006c5aa7    68d0f04000              push 0040f0d0

' *** Reference to "__vbaNew2"
'006c5aac    ff152c124000            call dword ptr [0040122c]
    Set var_31 = New frmDescriptifDon
'006c5ab2    8b1dc0b07200            mov ebx, dword ptr [0072b0c0]
'006c5ab8    8b8544ffffff            mov eax, dword ptr [ebp+ffffff44]
'006c5abe    8d4db0                  lea ecx, dword ptr [ebp-50]
'006c5ac1    51                      push ecx
'006c5ac2    8d55c0                  lea edx, dword ptr [ebp-40]
'006c5ac5    52                      push edx
'006c5ac6    899d30ffffff            mov dword ptr [ebp+ffffff30], ebx
'006c5acc    c745b8271a0000          mov dword ptr [ebp-48], 00001a27
'006c5ad3    c745b002000000          mov dword ptr [ebp-50], 00000002
'006c5ada    8945c8                  mov dword ptr [ebp-38], eax
'006c5add    c745c004000000          mov dword ptr [ebp-40], 00000004
'006c5ae4    e807cce2ff              call 4f26f0
    Call sub_4F26F0()
'006c5ae9    d98548ffffff            fld dword ptr [ebp+ffffff48]
'006c5aef    d825fc664000            fsub dword ptr [004066fc]
'006c5af5    66894598                mov word ptr [ebp-68], ax
'006c5af9    8d4da0                  lea ecx, dword ptr [ebp-60]
'006c5afc    c7459002000000          mov dword ptr [ebp-70], 00000002
'006c5b03    d95da8                  fstp dword ptr [ebp-58]
    'var_86 = (00)
'006c5b06    dfe0                    fnstsw ax
'006c5b08    a80d                    test al, 0d
'006c5b0a    0f85d30a0000            jne 6c65e3
'006c5b10    8d4590                  lea eax, dword ptr [ebp-70]
'006c5b13    50                      push eax
'006c5b14    c745a004000000          mov dword ptr [ebp-60], 00000004
'006c5b1b    8b1b                    mov ebx, dword ptr [ebx]
'006c5b1d    51                      push ecx
'006c5b1e    e86dcce2ff              call 4f2790
    Call sub_4F2790()
'006c5b23    0fbfd0                  movsx edx, ax
'006c5b26    8995d8feffff            mov dword ptr [ebp+fffffed8], edx
'006c5b2c    8bcb                    mov ecx, ebx
'006c5b2e    8b9d30ffffff            mov ebx, dword ptr [ebp+ffffff30]
'006c5b34    db85d8feffff            fild dword ptr [ebp+fffffed8]
'006c5b3a    d99dd4feffff            fstp dword ptr [ebp+fffffed4]
    'var_825 = (00)
'006c5b40    8b85d4feffff            mov eax, dword ptr [ebp+fffffed4]
'006c5b46    50                      push eax
'006c5b47    53                      push ebx
'006c5b48    ff918c000000            call dword ptr [ecx+0000008c]
'006c5b4e    dbe2                    fnclex
'006c5b50    85c0                    test ax, ax
'006c5b52    7d16                    jge 6c5b6a
'006c5b54    688c000000              push 0000008c
'006c5b59    6858124300              push 00431258
'006c5b5e    53                      push ebx

' *** Reference to "__vbaHresultCheckObj"
'006c5b5f    8b1d80104000            mov ebx, dword ptr [00401080]
'006c5b65    50                      push eax
'006c5b66    ffd3                    call ebx
'006c5b68    eb06                    jmp 6c5b70

' *** Reference to "__vbaHresultCheckObj"
'006c5b6a    8b1d80104000            mov ebx, dword ptr [00401080]
'006c5b70    8d5590                  lea edx, dword ptr [ebp-70]
'006c5b73    52                      push edx
'006c5b74    8d45a0                  lea eax, dword ptr [ebp-60]
'006c5b77    50                      push eax
'006c5b78    8d4db0                  lea ecx, dword ptr [ebp-50]
'006c5b7b    51                      push ecx
'006c5b7c    8d55c0                  lea edx, dword ptr [ebp-40]
'006c5b7f    52                      push edx
'006c5b80    6a04                    push 04

' *** Reference to "__vbaFreeVarList"
'006c5b82    ff1540104000            call dword ptr [00401040]
    '#Cleanup( var_5, var_6, var_7, var_8)
'006c5b88    83c414                  add esp, 14
    
End If
'006c5b8b    a1c0b07200              mov ax, word ptr [0072b0c0]
'006c5b90    85c0                    test ax, ax
'006c5b92    7510                    jne 6c5ba4
'006c5b94    68c0b07200              push 0072b0c0
'006c5b99    68d0f04000              push 0040f0d0

' *** Reference to "__vbaNew2"
'006c5b9e    ff152c124000            call dword ptr [0040122c]
Set var_31 = New frmDescriptifDon
'006c5ba4    a1c0b07200              mov ax, word ptr [0072b0c0]
'006c5ba9    8b08                    mov ecx, dword ptr [eax]
'006c5bab    8d9548ffffff            lea edx, dword ptr [ebp+ffffff48]
'006c5bb1    52                      push edx
'006c5bb2    50                      push eax
'006c5bb3    898540ffffff            mov dword ptr [ebp+ffffff40], eax
'006c5bb9    ff9188000000            call dword ptr [ecx+00000088]
'006c5bbf    dbe2                    fnclex
'006c5bc1    85c0                    test ax, ax
'006c5bc3    7d14                    jge 6c5bd9
'006c5bc5    8b8d40ffffff            mov ecx, dword ptr [ebp+ffffff40]
'006c5bcb    6888000000              push 00000088
'006c5bd0    6858124300              push 00431258
'006c5bd5    51                      push ecx
'006c5bd6    50                      push eax
'006c5bd7    ffd3                    call ebx
'006c5bd9    d98548ffffff            fld dword ptr [ebp+ffffff48]
'006c5bdf    d825f8664000            fsub dword ptr [004066f8]
'006c5be5    dfe0                    fnstsw ax
'006c5be7    a80d                    test al, 0d
'006c5be9    0f85f4090000            jne 6c65e3

' *** Reference to "__vbaFpI2"
'006c5bef    ff15a0124000            call dword ptr [004012a0]
var_num1 = CLng(((0) - 0!))
'006c5bf5    56                      push esi
'006c5bf6    8945dc                  mov dword ptr [ebp-24], eax
'006c5bf9    6605ea15                add ax, 15ea
var_num1 = var_num1 + 5610
'006c5bfd    0f80e5090000            jo 6c65e8
'006c5c03    0fbfd0                  movsx edx, ax
'006c5c06    8995ccfeffff            mov dword ptr [ebp+fffffecc], edx
'006c5c0c    db85ccfeffff            fild dword ptr [ebp+fffffecc]
'006c5c12    dd9dc4feffff            fstp qword ptr [ebp+fffffec4]
'var_344 = (00)
'006c5c18    dd85c4feffff            fld qword ptr [ebp+fffffec4]
'006c5c1e    833d00b0720000          cmp dword ptr [0072b000], 00
'006c5c25    7508                    jne 6c5c2f

If (vbNullChar = 0) Then
'006c5c27    dc35f0664000            fdiv qword ptr [004066f0]
'006c5c2d    eb11                    jmp 6c5c40
    
End If
'006c5c2f    ff35f4664000            push dword ptr [004066f4]
'006c5c35    ff35f0664000            push dword ptr [004066f0]

' *** Reference to "_adj_fdiv_m64"
'006c5c3b    e84416d4ff              call 407284
'006c5c40    dd5de4                  fstp qword ptr [ebp-1c]
'var_40 = (00)
'006c5c43    dfe0                    fnstsw ax
'006c5c45    a80d                    test al, 0d
'006c5c47    0f8596090000            jne 6c65e3
'006c5c4d    8b06                    mov eax, dword ptr [esi]
'006c5c4f    ff9008030000            call dword ptr [eax+00000308]
'006c5c55    50                      push eax
'006c5c56    8d4dd8                  lea ecx, dword ptr [ebp-28]
'006c5c59    51                      push ecx
'006c5c5a    ffd7                    call edi
Set var_45 = Nothing
'006c5c5c    8b4ddc                  mov ecx, dword ptr [ebp-24]
'006c5c5f    8b10                    mov edx, dword ptr [eax]
'006c5c61    6681c1b01d              add cx, 1db0
var_num3 = var_num1 + 7600
'006c5c66    0f807c090000            jo 6c65e8
'006c5c6c    0fbfc9                  movsx ecx, cx
'006c5c6f    898dc0feffff            mov dword ptr [ebp+fffffec0], ecx
'006c5c75    898540ffffff            mov dword ptr [ebp+ffffff40], eax
'006c5c7b    db85c0feffff            fild dword ptr [ebp+fffffec0]
'006c5c81    d99dbcfeffff            fstp dword ptr [ebp+fffffebc]
'var_482 = (00)
'006c5c87    8b8dbcfeffff            mov ecx, dword ptr [ebp+fffffebc]
'006c5c8d    51                      push ecx
'006c5c8e    50                      push eax
'006c5c8f    ff928c000000            call dword ptr [edx+0000008c]
'006c5c95    dbe2                    fnclex
'006c5c97    85c0                    test ax, ax
'006c5c99    7d14                    jge 6c5caf
'006c5c9b    8b9540ffffff            mov edx, dword ptr [ebp+ffffff40]
'006c5ca1    688c000000              push 0000008c
'006c5ca6    6874f14300              push 0043f174
'006c5cab    52                      push edx
'006c5cac    50                      push eax
'006c5cad    ffd3                    call ebx
'006c5caf    8d4dd8                  lea ecx, dword ptr [ebp-28]

' *** Reference to "__vbaFreeObj"
'006c5cb2    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_45)
'006c5cb8    8b06                    mov eax, dword ptr [esi]
'006c5cba    56                      push esi
'006c5cbb    ff900c030000            call dword ptr [eax+0000030c]
'006c5cc1    50                      push eax
'006c5cc2    8d4dd8                  lea ecx, dword ptr [ebp-28]
'006c5cc5    51                      push ecx
'006c5cc6    ffd7                    call edi
Set var_45 = Nothing
'006c5cc8    dd45e4                  fld qword ptr [ebp-1c]
'006c5ccb    dc0de8664000            fmul qword ptr [004066e8]
'006c5cd1    8bc8                    mov ecx, eax
'006c5cd3    8b11                    mov edx, dword ptr [ecx]
'006c5cd5    51                      push ecx
'006c5cd6    dfe0                    fnstsw ax
'006c5cd8    a80d                    test al, 0d
'006c5cda    0f8503090000            jne 6c65e3
'006c5ce0    d99db8feffff            fstp dword ptr [ebp+fffffeb8]
'var_767 = (00)
'006c5ce6    d985b8feffff            fld dword ptr [ebp+fffffeb8]
'006c5cec    898d40ffffff            mov dword ptr [ebp+ffffff40], ecx
'006c5cf2    d91c24                  fstp dword ptr [esp]
'var_518 = (00)
'006c5cf5    51                      push ecx
'006c5cf6    ff9284000000            call dword ptr [edx+00000084]
'006c5cfc    dbe2                    fnclex
'006c5cfe    85c0                    test ax, ax
'006c5d00    7d14                    jge 6c5d16
'006c5d02    8b8d40ffffff            mov ecx, dword ptr [ebp+ffffff40]
'006c5d08    6884000000              push 00000084
'006c5d0d    68d00d4300              push 00430dd0
'006c5d12    51                      push ecx
'006c5d13    50                      push eax
'006c5d14    ffd3                    call ebx
'006c5d16    8d4dd8                  lea ecx, dword ptr [ebp-28]

' *** Reference to "__vbaFreeObj"
'006c5d19    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_45)
'006c5d1f    8b16                    mov edx, dword ptr [esi]
'006c5d21    56                      push esi
'006c5d22    ff9218030000            call dword ptr [edx+00000318]
'006c5d28    50                      push eax
'006c5d29    8d45d8                  lea eax, dword ptr [ebp-28]
'006c5d2c    50                      push eax
'006c5d2d    ffd7                    call edi
Set var_45 = Nothing
'006c5d2f    dd45e4                  fld qword ptr [ebp-1c]
'006c5d32    dc0de8664000            fmul qword ptr [004066e8]
'006c5d38    8bc8                    mov ecx, eax
'006c5d3a    8b11                    mov edx, dword ptr [ecx]
'006c5d3c    51                      push ecx
'006c5d3d    dfe0                    fnstsw ax
'006c5d3f    a80d                    test al, 0d
'006c5d41    0f859c080000            jne 6c65e3
'006c5d47    d99db4feffff            fstp dword ptr [ebp+fffffeb4]
'var_492 = (00)
'006c5d4d    d985b4feffff            fld dword ptr [ebp+fffffeb4]
'006c5d53    898d40ffffff            mov dword ptr [ebp+ffffff40], ecx
'006c5d59    d91c24                  fstp dword ptr [esp]
'var_376 = (00)
'006c5d5c    51                      push ecx
'006c5d5d    ff9284000000            call dword ptr [edx+00000084]
'006c5d63    dbe2                    fnclex
'006c5d65    85c0                    test ax, ax
'006c5d67    7d14                    jge 6c5d7d
'006c5d69    8b8d40ffffff            mov ecx, dword ptr [ebp+ffffff40]
'006c5d6f    6884000000              push 00000084
'006c5d74    68d00d4300              push 00430dd0
'006c5d79    51                      push ecx
'006c5d7a    50                      push eax
'006c5d7b    ffd3                    call ebx
'006c5d7d    8d4dd8                  lea ecx, dword ptr [ebp-28]

' *** Reference to "__vbaFreeObj"
'006c5d80    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_45)
'006c5d86    8b16                    mov edx, dword ptr [esi]
'006c5d88    56                      push esi
'006c5d89    ff921c030000            call dword ptr [edx+0000031c]
'006c5d8f    50                      push eax
'006c5d90    8d45d8                  lea eax, dword ptr [ebp-28]
'006c5d93    50                      push eax
'006c5d94    ffd7                    call edi
Set var_45 = Nothing
'006c5d96    dd45e4                  fld qword ptr [ebp-1c]
'006c5d99    dc0de0664000            fmul qword ptr [004066e0]
'006c5d9f    8bc8                    mov ecx, eax
'006c5da1    8b11                    mov edx, dword ptr [ecx]
'006c5da3    51                      push ecx
'006c5da4    dfe0                    fnstsw ax
'006c5da6    a80d                    test al, 0d
'006c5da8    0f8535080000            jne 6c65e3
'006c5dae    d99db0feffff            fstp dword ptr [ebp+fffffeb0]
'var_521 = (00)
'006c5db4    d985b0feffff            fld dword ptr [ebp+fffffeb0]
'006c5dba    898d40ffffff            mov dword ptr [ebp+ffffff40], ecx
'006c5dc0    d91c24                  fstp dword ptr [esp]
'var_280 = (00)
'006c5dc3    51                      push ecx
'006c5dc4    ff9284000000            call dword ptr [edx+00000084]
'006c5dca    dbe2                    fnclex
'006c5dcc    85c0                    test ax, ax
'006c5dce    7d14                    jge 6c5de4
'006c5dd0    8b8d40ffffff            mov ecx, dword ptr [ebp+ffffff40]
'006c5dd6    6884000000              push 00000084
'006c5ddb    68d00d4300              push 00430dd0
'006c5de0    51                      push ecx
'006c5de1    50                      push eax
'006c5de2    ffd3                    call ebx
'006c5de4    8d4dd8                  lea ecx, dword ptr [ebp-28]

' *** Reference to "__vbaFreeObj"
'006c5de7    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_45)
'006c5ded    8b16                    mov edx, dword ptr [esi]
'006c5def    56                      push esi
'006c5df0    ff9214030000            call dword ptr [edx+00000314]
'006c5df6    50                      push eax
'006c5df7    8d45d8                  lea eax, dword ptr [ebp-28]
'006c5dfa    50                      push eax
'006c5dfb    ffd7                    call edi
Set var_45 = Nothing
'006c5dfd    dd45e4                  fld qword ptr [ebp-1c]
'006c5e00    dc0de8664000            fmul qword ptr [004066e8]
'006c5e06    8bc8                    mov ecx, eax
'006c5e08    8b11                    mov edx, dword ptr [ecx]
'006c5e0a    51                      push ecx
'006c5e0b    dfe0                    fnstsw ax
'006c5e0d    a80d                    test al, 0d
'006c5e0f    0f85ce070000            jne 6c65e3
'006c5e15    d99dacfeffff            fstp dword ptr [ebp+fffffeac]
'var_520 = (00)
'006c5e1b    d985acfeffff            fld dword ptr [ebp+fffffeac]
'006c5e21    898d40ffffff            mov dword ptr [ebp+ffffff40], ecx
'006c5e27    d91c24                  fstp dword ptr [esp]
'var_287 = (00)
'006c5e2a    51                      push ecx
'006c5e2b    ff9284000000            call dword ptr [edx+00000084]
'006c5e31    dbe2                    fnclex
'006c5e33    85c0                    test ax, ax
'006c5e35    7d14                    jge 6c5e4b
'006c5e37    8b8d40ffffff            mov ecx, dword ptr [ebp+ffffff40]
'006c5e3d    6884000000              push 00000084
'006c5e42    68d00d4300              push 00430dd0
'006c5e47    51                      push ecx
'006c5e48    50                      push eax
'006c5e49    ffd3                    call ebx
'006c5e4b    8d4dd8                  lea ecx, dword ptr [ebp-28]

' *** Reference to "__vbaFreeObj"
'006c5e4e    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_45)
'006c5e54    8b16                    mov edx, dword ptr [esi]
'006c5e56    56                      push esi
'006c5e57    ff9210030000            call dword ptr [edx+00000310]
'006c5e5d    50                      push eax
'006c5e5e    8d45d8                  lea eax, dword ptr [ebp-28]
'006c5e61    50                      push eax
'006c5e62    ffd7                    call edi
Set var_45 = Nothing
'006c5e64    dd45e4                  fld qword ptr [ebp-1c]
'006c5e67    dc0de8664000            fmul qword ptr [004066e8]
'006c5e6d    8bc8                    mov ecx, eax
'006c5e6f    8b11                    mov edx, dword ptr [ecx]
'006c5e71    51                      push ecx
'006c5e72    dfe0                    fnstsw ax
'006c5e74    a80d                    test al, 0d
'006c5e76    0f8567070000            jne 6c65e3
'006c5e7c    d99da8feffff            fstp dword ptr [ebp+fffffea8]
'var_519 = (00)
'006c5e82    d985a8feffff            fld dword ptr [ebp+fffffea8]
'006c5e88    898d40ffffff            mov dword ptr [ebp+ffffff40], ecx
'006c5e8e    d91c24                  fstp dword ptr [esp]
'var_603 = (00)
'006c5e91    51                      push ecx
'006c5e92    ff9284000000            call dword ptr [edx+00000084]
'006c5e98    dbe2                    fnclex
'006c5e9a    85c0                    test ax, ax
'006c5e9c    7d14                    jge 6c5eb2
'006c5e9e    8b8d40ffffff            mov ecx, dword ptr [ebp+ffffff40]
'006c5ea4    6884000000              push 00000084
'006c5ea9    68d00d4300              push 00430dd0
'006c5eae    51                      push ecx
'006c5eaf    50                      push eax
'006c5eb0    ffd3                    call ebx
'006c5eb2    8d4dd8                  lea ecx, dword ptr [ebp-28]

' *** Reference to "__vbaFreeObj"
'006c5eb5    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_45)
'006c5ebb    8b16                    mov edx, dword ptr [esi]
'006c5ebd    56                      push esi
'006c5ebe    ff9218030000            call dword ptr [edx+00000318]
'006c5ec4    50                      push eax
'006c5ec5    8d45d8                  lea eax, dword ptr [ebp-28]
'006c5ec8    50                      push eax
'006c5ec9    ffd7                    call edi
Set var_45 = Nothing
'006c5ecb    dd45e4                  fld qword ptr [ebp-1c]
'006c5ece    dc2558164000            fsub qword ptr [00401658]
'006c5ed4    8bc8                    mov ecx, eax
'006c5ed6    8b11                    mov edx, dword ptr [ecx]
'006c5ed8    51                      push ecx
'006c5ed9    dc0de8664000            fmul qword ptr [004066e8]
'006c5edf    898d40ffffff            mov dword ptr [ebp+ffffff40], ecx
'006c5ee5    dc05d8664000            fadd qword ptr [004066d8]
'006c5eeb    dfe0                    fnstsw ax
'006c5eed    a80d                    test al, 0d
'006c5eef    0f85ee060000            jne 6c65e3
'006c5ef5    d99da4feffff            fstp dword ptr [ebp+fffffea4]
'var_534 = (00)
'006c5efb    d985a4feffff            fld dword ptr [ebp+fffffea4]
'006c5f01    d91c24                  fstp dword ptr [esp]
'var_325 = (00)
'006c5f04    51                      push ecx
'006c5f05    ff5274                  call dword ptr [edx+74]
'006c5f08    dbe2                    fnclex
'006c5f0a    85c0                    test ax, ax
'006c5f0c    7d11                    jge 6c5f1f
'006c5f0e    8b8d40ffffff            mov ecx, dword ptr [ebp+ffffff40]
'006c5f14    6a74                    push 74
'006c5f16    68d00d4300              push 00430dd0
'006c5f1b    51                      push ecx
'006c5f1c    50                      push eax
'006c5f1d    ffd3                    call ebx
'006c5f1f    8d4dd8                  lea ecx, dword ptr [ebp-28]

' *** Reference to "__vbaFreeObj"
'006c5f22    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_45)
'006c5f28    8b16                    mov edx, dword ptr [esi]
'006c5f2a    56                      push esi
'006c5f2b    ff9220030000            call dword ptr [edx+00000320]
'006c5f31    50                      push eax
'006c5f32    8d45d8                  lea eax, dword ptr [ebp-28]
'006c5f35    50                      push eax
'006c5f36    ffd7                    call edi
Set var_45 = Nothing
'006c5f38    8b08                    mov ecx, dword ptr [eax]
'006c5f3a    8d55d4                  lea edx, dword ptr [ebp-2c]
'006c5f3d    52                      push edx
'006c5f3e    6a01                    push 01
'006c5f40    50                      push eax
'006c5f41    898540ffffff            mov dword ptr [ebp+ffffff40], eax
'006c5f47    ff5140                  call dword ptr [ecx+40]
'006c5f4a    dbe2                    fnclex
'006c5f4c    85c0                    test ax, ax
'006c5f4e    7d11                    jge 6c5f61
'006c5f50    8b8d40ffffff            mov ecx, dword ptr [ebp+ffffff40]
'006c5f56    6a40                    push 40
'006c5f58    686c384300              push 0043386c
'006c5f5d    51                      push ecx
'006c5f5e    50                      push eax
'006c5f5f    ffd3                    call ebx
'006c5f61    dd45e4                  fld qword ptr [ebp-1c]
'006c5f64    8b4dd4                  mov ecx, dword ptr [ebp-2c]
'006c5f67    dc2558164000            fsub qword ptr [00401658]
'006c5f6d    8b11                    mov edx, dword ptr [ecx]
'006c5f6f    51                      push ecx
'006c5f70    898d38ffffff            mov dword ptr [ebp+ffffff38], ecx
'006c5f76    dc0de8664000            fmul qword ptr [004066e8]
'006c5f7c    dc05d0664000            fadd qword ptr [004066d0]
'006c5f82    dfe0                    fnstsw ax
'006c5f84    a80d                    test al, 0d
'006c5f86    0f8557060000            jne 6c65e3
'006c5f8c    d99da0feffff            fstp dword ptr [ebp+fffffea0]
'var_527 = (00)
'006c5f92    d985a0feffff            fld dword ptr [ebp+fffffea0]
'006c5f98    d91c24                  fstp dword ptr [esp]
'var_690 = (00)
'006c5f9b    51                      push ecx
'006c5f9c    ff527c                  call dword ptr [edx+7c]
'006c5f9f    dbe2                    fnclex
'006c5fa1    85c0                    test ax, ax
'006c5fa3    7d11                    jge 6c5fb6
'006c5fa5    8b8d38ffffff            mov ecx, dword ptr [ebp+ffffff38]
'006c5fab    6a7c                    push 7c
'006c5fad    683ce04300              push 0043e03c
'006c5fb2    51                      push ecx
'006c5fb3    50                      push eax
'006c5fb4    ffd3                    call ebx
'006c5fb6    8d55d4                  lea edx, dword ptr [ebp-2c]
'006c5fb9    52                      push edx
'006c5fba    8d45d8                  lea eax, dword ptr [ebp-28]
'006c5fbd    50                      push eax
'006c5fbe    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'006c5fc0    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_45, var_44)
'006c5fc6    8b0e                    mov ecx, dword ptr [esi]
'006c5fc8    83c40c                  add esp, 0c
'006c5fcb    56                      push esi
'006c5fcc    ff911c030000            call dword ptr [ecx+0000031c]
'006c5fd2    50                      push eax
'006c5fd3    8d55d8                  lea edx, dword ptr [ebp-28]
'006c5fd6    52                      push edx
'006c5fd7    ffd7                    call edi
Set var_45 = 
'006c5fd9    dd45e4                  fld qword ptr [ebp-1c]
'006c5fdc    dc2558164000            fsub qword ptr [00401658]
'006c5fe2    8bc8                    mov ecx, eax
'006c5fe4    8b11                    mov edx, dword ptr [ecx]
'006c5fe6    51                      push ecx
'006c5fe7    dc0dc8664000            fmul qword ptr [004066c8]
'006c5fed    898d40ffffff            mov dword ptr [ebp+ffffff40], ecx
'006c5ff3    dc05c0664000            fadd qword ptr [004066c0]
'006c5ff9    dfe0                    fnstsw ax
'006c5ffb    a80d                    test al, 0d
'006c5ffd    0f85e0050000            jne 6c65e3
'006c6003    d99d9cfeffff            fstp dword ptr [ebp+fffffe9c]
'var_822 = (00)
'006c6009    d9859cfeffff            fld dword ptr [ebp+fffffe9c]
'006c600f    d91c24                  fstp dword ptr [esp]
'var_689 = (00)
'006c6012    51                      push ecx
'006c6013    ff5274                  call dword ptr [edx+74]
'006c6016    dbe2                    fnclex
'006c6018    85c0                    test ax, ax
'006c601a    7d11                    jge 6c602d
'006c601c    8b8d40ffffff            mov ecx, dword ptr [ebp+ffffff40]
'006c6022    6a74                    push 74
'006c6024    68d00d4300              push 00430dd0
'006c6029    51                      push ecx
'006c602a    50                      push eax
'006c602b    ffd3                    call ebx
'006c602d    8d4dd8                  lea ecx, dword ptr [ebp-28]

' *** Reference to "__vbaFreeObj"
'006c6030    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_45)
'006c6036    8b16                    mov edx, dword ptr [esi]
'006c6038    56                      push esi
'006c6039    ff9220030000            call dword ptr [edx+00000320]
'006c603f    50                      push eax
'006c6040    8d45d8                  lea eax, dword ptr [ebp-28]
'006c6043    50                      push eax
'006c6044    ffd7                    call edi
Set var_45 = 
'006c6046    8b08                    mov ecx, dword ptr [eax]
'006c6048    8d55d4                  lea edx, dword ptr [ebp-2c]
'006c604b    52                      push edx
'006c604c    6a00                    push 00
'006c604e    50                      push eax
'006c604f    898540ffffff            mov dword ptr [ebp+ffffff40], eax
'006c6055    ff5140                  call dword ptr [ecx+40]
'006c6058    dbe2                    fnclex
'006c605a    85c0                    test ax, ax
'006c605c    7d11                    jge 6c606f
'006c605e    8b8d40ffffff            mov ecx, dword ptr [ebp+ffffff40]
'006c6064    6a40                    push 40
'006c6066    686c384300              push 0043386c
'006c606b    51                      push ecx
'006c606c    50                      push eax
'006c606d    ffd3                    call ebx
'006c606f    dd45e4                  fld qword ptr [ebp-1c]
'006c6072    8b4dd4                  mov ecx, dword ptr [ebp-2c]
'006c6075    dc2558164000            fsub qword ptr [00401658]
'006c607b    8b11                    mov edx, dword ptr [ecx]
'006c607d    51                      push ecx
'006c607e    898d38ffffff            mov dword ptr [ebp+ffffff38], ecx
'006c6084    dc0dc8664000            fmul qword ptr [004066c8]
'006c608a    dc05b8664000            fadd qword ptr [004066b8]
'006c6090    dfe0                    fnstsw ax
'006c6092    a80d                    test al, 0d
'006c6094    0f8549050000            jne 6c65e3
'006c609a    d99d98feffff            fstp dword ptr [ebp+fffffe98]
'var_537 = (00)
'006c60a0    d98598feffff            fld dword ptr [ebp+fffffe98]
'006c60a6    d91c24                  fstp dword ptr [esp]
'var_677 = (00)
'006c60a9    51                      push ecx
'006c60aa    ff527c                  call dword ptr [edx+7c]
'006c60ad    dbe2                    fnclex
'006c60af    85c0                    test ax, ax
'006c60b1    7d11                    jge 6c60c4
'006c60b3    8b8d38ffffff            mov ecx, dword ptr [ebp+ffffff38]
'006c60b9    6a7c                    push 7c
'006c60bb    683ce04300              push 0043e03c
'006c60c0    51                      push ecx
'006c60c1    50                      push eax
'006c60c2    ffd3                    call ebx
'006c60c4    8d55d4                  lea edx, dword ptr [ebp-2c]
'006c60c7    52                      push edx
'006c60c8    8d45d8                  lea eax, dword ptr [ebp-28]
'006c60cb    50                      push eax
'006c60cc    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'006c60ce    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_45, var_44)
'006c60d4    8b0e                    mov ecx, dword ptr [esi]
'006c60d6    83c40c                  add esp, 0c
'006c60d9    56                      push esi
'006c60da    ff9114030000            call dword ptr [ecx+00000314]
'006c60e0    50                      push eax
'006c60e1    8d55d8                  lea edx, dword ptr [ebp-28]
'006c60e4    52                      push edx
'006c60e5    ffd7                    call edi
Set var_45 = 
'006c60e7    dd45e4                  fld qword ptr [ebp-1c]
'006c60ea    dc2558164000            fsub qword ptr [00401658]
'006c60f0    8bc8                    mov ecx, eax
'006c60f2    8b11                    mov edx, dword ptr [ecx]
'006c60f4    51                      push ecx
'006c60f5    dc0db0664000            fmul qword ptr [004066b0]
'006c60fb    898d40ffffff            mov dword ptr [ebp+ffffff40], ecx
'006c6101    dc05a8664000            fadd qword ptr [004066a8]
'006c6107    dfe0                    fnstsw ax
'006c6109    a80d                    test al, 0d
'006c610b    0f85d2040000            jne 6c65e3
'006c6111    d99d94feffff            fstp dword ptr [ebp+fffffe94]
'var_538 = (00)
'006c6117    d98594feffff            fld dword ptr [ebp+fffffe94]
'006c611d    d91c24                  fstp dword ptr [esp]
'var_288 = (00)
'006c6120    51                      push ecx
'006c6121    ff5274                  call dword ptr [edx+74]
'006c6124    dbe2                    fnclex
'006c6126    85c0                    test ax, ax
'006c6128    7d11                    jge 6c613b
'006c612a    8b8d40ffffff            mov ecx, dword ptr [ebp+ffffff40]
'006c6130    6a74                    push 74
'006c6132    68d00d4300              push 00430dd0
'006c6137    51                      push ecx
'006c6138    50                      push eax
'006c6139    ffd3                    call ebx
'006c613b    8d4dd8                  lea ecx, dword ptr [ebp-28]

' *** Reference to "__vbaFreeObj"
'006c613e    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_45)
'006c6144    8b16                    mov edx, dword ptr [esi]
'006c6146    56                      push esi
'006c6147    ff9220030000            call dword ptr [edx+00000320]
'006c614d    50                      push eax
'006c614e    8d45d8                  lea eax, dword ptr [ebp-28]
'006c6151    50                      push eax
'006c6152    ffd7                    call edi
Set var_45 = 
'006c6154    8b08                    mov ecx, dword ptr [eax]
'006c6156    8d55d4                  lea edx, dword ptr [ebp-2c]
'006c6159    52                      push edx
'006c615a    6a02                    push 02
'006c615c    50                      push eax
'006c615d    898540ffffff            mov dword ptr [ebp+ffffff40], eax
'006c6163    ff5140                  call dword ptr [ecx+40]
'006c6166    dbe2                    fnclex
'006c6168    85c0                    test ax, ax
'006c616a    7d11                    jge 6c617d
'006c616c    8b8d40ffffff            mov ecx, dword ptr [ebp+ffffff40]
'006c6172    6a40                    push 40
'006c6174    686c384300              push 0043386c
'006c6179    51                      push ecx
'006c617a    50                      push eax
'006c617b    ffd3                    call ebx
'006c617d    dd45e4                  fld qword ptr [ebp-1c]
'006c6180    8b4dd4                  mov ecx, dword ptr [ebp-2c]
'006c6183    dc2558164000            fsub qword ptr [00401658]
'006c6189    8b11                    mov edx, dword ptr [ecx]
'006c618b    51                      push ecx
'006c618c    898d38ffffff            mov dword ptr [ebp+ffffff38], ecx
'006c6192    dc0db0664000            fmul qword ptr [004066b0]
'006c6198    dc05a0664000            fadd qword ptr [004066a0]
'006c619e    dfe0                    fnstsw ax
'006c61a0    a80d                    test al, 0d
'006c61a2    0f853b040000            jne 6c65e3
'006c61a8    d99d90feffff            fstp dword ptr [ebp+fffffe90]
'var_771 = (00)
'006c61ae    d98590feffff            fld dword ptr [ebp+fffffe90]
'006c61b4    d91c24                  fstp dword ptr [esp]
'var_691 = (00)
'006c61b7    51                      push ecx
'006c61b8    ff527c                  call dword ptr [edx+7c]
'006c61bb    dbe2                    fnclex
'006c61bd    85c0                    test ax, ax
'006c61bf    7d11                    jge 6c61d2
'006c61c1    8b8d38ffffff            mov ecx, dword ptr [ebp+ffffff38]
'006c61c7    6a7c                    push 7c
'006c61c9    683ce04300              push 0043e03c
'006c61ce    51                      push ecx
'006c61cf    50                      push eax
'006c61d0    ffd3                    call ebx
'006c61d2    8d55d4                  lea edx, dword ptr [ebp-2c]
'006c61d5    52                      push edx
'006c61d6    8d45d8                  lea eax, dword ptr [ebp-28]
'006c61d9    50                      push eax
'006c61da    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'006c61dc    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_45, var_44)
'006c61e2    8b0e                    mov ecx, dword ptr [esi]
'006c61e4    83c40c                  add esp, 0c
'006c61e7    56                      push esi
'006c61e8    ff9110030000            call dword ptr [ecx+00000310]
'006c61ee    50                      push eax
'006c61ef    8d55d8                  lea edx, dword ptr [ebp-28]
'006c61f2    52                      push edx
'006c61f3    ffd7                    call edi
Set var_45 = 
'006c61f5    dd45e4                  fld qword ptr [ebp-1c]
'006c61f8    dc2558164000            fsub qword ptr [00401658]
'006c61fe    8bc8                    mov ecx, eax
'006c6200    8b11                    mov edx, dword ptr [ecx]
'006c6202    51                      push ecx
'006c6203    dc0d98664000            fmul qword ptr [00406698]
'006c6209    898d40ffffff            mov dword ptr [ebp+ffffff40], ecx
'006c620f    dc0590664000            fadd qword ptr [00406690]
'006c6215    dfe0                    fnstsw ax
'006c6217    a80d                    test al, 0d
'006c6219    0f85c4030000            jne 6c65e3
'006c621f    d99d8cfeffff            fstp dword ptr [ebp+fffffe8c]
'var_772 = (00)
'006c6225    d9858cfeffff            fld dword ptr [ebp+fffffe8c]
'006c622b    d91c24                  fstp dword ptr [esp]
'var_690 = (00)
'006c622e    51                      push ecx
'006c622f    ff5274                  call dword ptr [edx+74]
'006c6232    dbe2                    fnclex
'006c6234    85c0                    test ax, ax
'006c6236    7d11                    jge 6c6249
'006c6238    8b8d40ffffff            mov ecx, dword ptr [ebp+ffffff40]
'006c623e    6a74                    push 74
'006c6240    68d00d4300              push 00430dd0
'006c6245    51                      push ecx
'006c6246    50                      push eax
'006c6247    ffd3                    call ebx
'006c6249    8d4dd8                  lea ecx, dword ptr [ebp-28]

' *** Reference to "__vbaFreeObj"
'006c624c    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_45)
'006c6252    8b16                    mov edx, dword ptr [esi]
'006c6254    56                      push esi
'006c6255    ff9220030000            call dword ptr [edx+00000320]
'006c625b    50                      push eax
'006c625c    8d45d8                  lea eax, dword ptr [ebp-28]
'006c625f    50                      push eax
'006c6260    ffd7                    call edi
Set var_45 = 
'006c6262    8b08                    mov ecx, dword ptr [eax]
'006c6264    8d55d4                  lea edx, dword ptr [ebp-2c]
'006c6267    52                      push edx
'006c6268    6a03                    push 03
'006c626a    50                      push eax
'006c626b    898540ffffff            mov dword ptr [ebp+ffffff40], eax
'006c6271    ff5140                  call dword ptr [ecx+40]
'006c6274    dbe2                    fnclex
'006c6276    85c0                    test ax, ax
'006c6278    7d11                    jge 6c628b
'006c627a    8b8d40ffffff            mov ecx, dword ptr [ebp+ffffff40]
'006c6280    6a40                    push 40
'006c6282    686c384300              push 0043386c
'006c6287    51                      push ecx
'006c6288    50                      push eax
'006c6289    ffd3                    call ebx
'006c628b    dd45e4                  fld qword ptr [ebp-1c]
'006c628e    8b4dd4                  mov ecx, dword ptr [ebp-2c]
'006c6291    dc2558164000            fsub qword ptr [00401658]
'006c6297    8b11                    mov edx, dword ptr [ecx]
'006c6299    51                      push ecx
'006c629a    898d38ffffff            mov dword ptr [ebp+ffffff38], ecx
'006c62a0    dc0d98664000            fmul qword ptr [00406698]
'006c62a6    dc0588664000            fadd qword ptr [00406688]
'006c62ac    dfe0                    fnstsw ax
'006c62ae    a80d                    test al, 0d
'006c62b0    0f852d030000            jne 6c65e3
'006c62b6    d99d88feffff            fstp dword ptr [ebp+fffffe88]
'var_281 = (00)
'006c62bc    d98588feffff            fld dword ptr [ebp+fffffe88]
'006c62c2    d91c24                  fstp dword ptr [esp]
'var_289 = (00)
'006c62c5    51                      push ecx
'006c62c6    ff527c                  call dword ptr [edx+7c]
'006c62c9    dbe2                    fnclex
'006c62cb    85c0                    test ax, ax
'006c62cd    7d11                    jge 6c62e0
'006c62cf    8b8d38ffffff            mov ecx, dword ptr [ebp+ffffff38]
'006c62d5    6a7c                    push 7c
'006c62d7    683ce04300              push 0043e03c
'006c62dc    51                      push ecx
'006c62dd    50                      push eax
'006c62de    ffd3                    call ebx
'006c62e0    8d55d4                  lea edx, dword ptr [ebp-2c]
'006c62e3    52                      push edx
'006c62e4    8d45d8                  lea eax, dword ptr [ebp-28]
'006c62e7    50                      push eax
'006c62e8    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'006c62ea    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_45, var_44)
'006c62f0    8b0e                    mov ecx, dword ptr [esi]
'006c62f2    83c40c                  add esp, 0c
'006c62f5    56                      push esi
'006c62f6    ff9104030000            call dword ptr [ecx+00000304]
'006c62fc    50                      push eax
'006c62fd    8d55d8                  lea edx, dword ptr [ebp-28]
'006c6300    52                      push edx
'006c6301    ffd7                    call edi
Set var_45 = 
'006c6303    dd45e4                  fld qword ptr [ebp-1c]
'006c6306    dc2558164000            fsub qword ptr [00401658]
'006c630c    8bc8                    mov ecx, eax
'006c630e    8b11                    mov edx, dword ptr [ecx]
'006c6310    51                      push ecx
'006c6311    dc0df0664000            fmul qword ptr [004066f0]
'006c6317    898d40ffffff            mov dword ptr [ebp+ffffff40], ecx
'006c631d    dc0580664000            fadd qword ptr [00406680]
'006c6323    dfe0                    fnstsw ax
'006c6325    a80d                    test al, 0d
'006c6327    0f85b6020000            jne 6c65e3
'006c632d    d99d84feffff            fstp dword ptr [ebp+fffffe84]
'var_282 = (00)
'006c6333    d98584feffff            fld dword ptr [ebp+fffffe84]
'006c6339    d91c24                  fstp dword ptr [esp]
'var_677 = (00)
'006c633c    51                      push ecx
'006c633d    ff527c                  call dword ptr [edx+7c]
'006c6340    dbe2                    fnclex
'006c6342    85c0                    test ax, ax
'006c6344    7d11                    jge 6c6357
'006c6346    8b8d40ffffff            mov ecx, dword ptr [ebp+ffffff40]
'006c634c    6a7c                    push 7c
'006c634e    681cb94300              push 0043b91c
'006c6353    51                      push ecx
'006c6354    50                      push eax
'006c6355    ffd3                    call ebx
'006c6357    8d4dd8                  lea ecx, dword ptr [ebp-28]

' *** Reference to "__vbaFreeObj"
'006c635a    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_45)
'006c6360    8b16                    mov edx, dword ptr [esi]
'006c6362    56                      push esi
'006c6363    ff9200030000            call dword ptr [edx+00000300]
'006c6369    50                      push eax
'006c636a    8d45d4                  lea eax, dword ptr [ebp-2c]
'006c636d    50                      push eax
'006c636e    ffd7                    call edi
Set var_44 = 
'006c6370    8b0e                    mov ecx, dword ptr [esi]
'006c6372    56                      push esi
'006c6373    898538ffffff            mov dword ptr [ebp+ffffff38], eax
'006c6379    ff9104030000            call dword ptr [ecx+00000304]
'006c637f    50                      push eax
'006c6380    8d55d8                  lea edx, dword ptr [ebp-28]
'006c6383    52                      push edx
'006c6384    ffd7                    call edi
Set var_45 = var_44
'006c6386    8b08                    mov ecx, dword ptr [eax]
'006c6388    8d9548ffffff            lea edx, dword ptr [ebp+ffffff48]
'006c638e    52                      push edx
'006c638f    50                      push eax
'006c6390    898540ffffff            mov dword ptr [ebp+ffffff40], eax
'006c6396    ff5178                  call dword ptr [ecx+78]
'006c6399    dbe2                    fnclex
'006c639b    85c0                    test ax, ax
'006c639d    7d11                    jge 6c63b0
'006c639f    8b8d40ffffff            mov ecx, dword ptr [ebp+ffffff40]
'006c63a5    6a78                    push 78
'006c63a7    681cb94300              push 0043b91c
'006c63ac    51                      push ecx
'006c63ad    50                      push eax
'006c63ae    ffd3                    call ebx
'006c63b0    8b8d48ffffff            mov ecx, dword ptr [ebp+ffffff48]
'006c63b6    8b8538ffffff            mov eax, dword ptr [ebp+ffffff38]
'006c63bc    8b10                    mov edx, dword ptr [eax]
'006c63be    51                      push ecx
'006c63bf    50                      push eax
'006c63c0    ff5274                  call dword ptr [edx+74]
'006c63c3    dbe2                    fnclex
'006c63c5    85c0                    test ax, ax
'006c63c7    7d11                    jge 6c63da
'006c63c9    8b9538ffffff            mov edx, dword ptr [ebp+ffffff38]
'006c63cf    6a74                    push 74
'006c63d1    6838eb4300              push 0043eb38
'006c63d6    52                      push edx
'006c63d7    50                      push eax
'006c63d8    ffd3                    call ebx
'006c63da    8d45d4                  lea eax, dword ptr [ebp-2c]
'006c63dd    50                      push eax
'006c63de    8d4dd8                  lea ecx, dword ptr [ebp-28]
'006c63e1    51                      push ecx
'006c63e2    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'006c63e4    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_45, var_44)
'006c63ea    8b16                    mov edx, dword ptr [esi]
'006c63ec    83c40c                  add esp, 0c
'006c63ef    56                      push esi
'006c63f0    ff92fc020000            call dword ptr [edx+000002fc]
'006c63f6    50                      push eax
'006c63f7    8d45d4                  lea eax, dword ptr [ebp-2c]
'006c63fa    50                      push eax
'006c63fb    ffd7                    call edi
Set var_44 = 
'006c63fd    8b0e                    mov ecx, dword ptr [esi]
'006c63ff    56                      push esi
'006c6400    898538ffffff            mov dword ptr [ebp+ffffff38], eax
'006c6406    ff9104030000            call dword ptr [ecx+00000304]
'006c640c    50                      push eax
'006c640d    8d55d8                  lea edx, dword ptr [ebp-28]
'006c6410    52                      push edx
'006c6411    ffd7                    call edi
Set var_45 = var_44
'006c6413    8b08                    mov ecx, dword ptr [eax]
'006c6415    8d9548ffffff            lea edx, dword ptr [ebp+ffffff48]
'006c641b    52                      push edx
'006c641c    50                      push eax
'006c641d    898540ffffff            mov dword ptr [ebp+ffffff40], eax
'006c6423    ff5178                  call dword ptr [ecx+78]
'006c6426    dbe2                    fnclex
'006c6428    85c0                    test ax, ax
'006c642a    7d11                    jge 6c643d
'006c642c    8b8d40ffffff            mov ecx, dword ptr [ebp+ffffff40]
'006c6432    6a78                    push 78
'006c6434    681cb94300              push 0043b91c
'006c6439    51                      push ecx
'006c643a    50                      push eax
'006c643b    ffd3                    call ebx
'006c643d    8b8d48ffffff            mov ecx, dword ptr [ebp+ffffff48]
'006c6443    8b8538ffffff            mov eax, dword ptr [ebp+ffffff38]
'006c6449    8b10                    mov edx, dword ptr [eax]
'006c644b    51                      push ecx
'006c644c    50                      push eax
'006c644d    ff5274                  call dword ptr [edx+74]
'006c6450    dbe2                    fnclex
'006c6452    85c0                    test ax, ax
'006c6454    7d11                    jge 6c6467
'006c6456    8b9538ffffff            mov edx, dword ptr [ebp+ffffff38]
'006c645c    6a74                    push 74
'006c645e    6838eb4300              push 0043eb38
'006c6463    52                      push edx
'006c6464    50                      push eax
'006c6465    ffd3                    call ebx
'006c6467    8d45d4                  lea eax, dword ptr [ebp-2c]
'006c646a    50                      push eax
'006c646b    8d4dd8                  lea ecx, dword ptr [ebp-28]
'006c646e    51                      push ecx
'006c646f    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'006c6471    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_45, var_44)
'006c6477    8b16                    mov edx, dword ptr [esi]
'006c6479    83c40c                  add esp, 0c
'006c647c    56                      push esi
'006c647d    ff9224030000            call dword ptr [edx+00000324]
'006c6483    50                      push eax
'006c6484    8d45d4                  lea eax, dword ptr [ebp-2c]
'006c6487    50                      push eax
'006c6488    ffd7                    call edi
Set var_44 = 
'006c648a    8b0e                    mov ecx, dword ptr [esi]
'006c648c    56                      push esi
'006c648d    898538ffffff            mov dword ptr [ebp+ffffff38], eax
'006c6493    ff9104030000            call dword ptr [ecx+00000304]
'006c6499    50                      push eax
'006c649a    8d55d8                  lea edx, dword ptr [ebp-28]
'006c649d    52                      push edx
'006c649e    ffd7                    call edi
Set var_45 = var_44
'006c64a0    8b08                    mov ecx, dword ptr [eax]
'006c64a2    8d9548ffffff            lea edx, dword ptr [ebp+ffffff48]
'006c64a8    52                      push edx
'006c64a9    50                      push eax
'006c64aa    898540ffffff            mov dword ptr [ebp+ffffff40], eax
'006c64b0    ff5178                  call dword ptr [ecx+78]
'006c64b3    dbe2                    fnclex
'006c64b5    85c0                    test ax, ax
'006c64b7    7d11                    jge 6c64ca
'006c64b9    8b8d40ffffff            mov ecx, dword ptr [ebp+ffffff40]
'006c64bf    6a78                    push 78
'006c64c1    681cb94300              push 0043b91c
'006c64c6    51                      push ecx
'006c64c7    50                      push eax
'006c64c8    ffd3                    call ebx
'006c64ca    8b8d48ffffff            mov ecx, dword ptr [ebp+ffffff48]
'006c64d0    8b8538ffffff            mov eax, dword ptr [ebp+ffffff38]
'006c64d6    8b10                    mov edx, dword ptr [eax]
'006c64d8    51                      push ecx
'006c64d9    50                      push eax
'006c64da    ff5274                  call dword ptr [edx+74]
'006c64dd    dbe2                    fnclex
'006c64df    85c0                    test ax, ax
'006c64e1    7d11                    jge 6c64f4
'006c64e3    8b9538ffffff            mov edx, dword ptr [ebp+ffffff38]
'006c64e9    6a74                    push 74
'006c64eb    6838eb4300              push 0043eb38
'006c64f0    52                      push edx
'006c64f1    50                      push eax
'006c64f2    ffd3                    call ebx
'006c64f4    8d45d4                  lea eax, dword ptr [ebp-2c]
'006c64f7    50                      push eax
'006c64f8    8d4dd8                  lea ecx, dword ptr [ebp-28]
'006c64fb    51                      push ecx
'006c64fc    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'006c64fe    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_45, var_44)
'006c6504    8b16                    mov edx, dword ptr [esi]
'006c6506    83c40c                  add esp, 0c
'006c6509    56                      push esi
'006c650a    ff9228030000            call dword ptr [edx+00000328]
'006c6510    50                      push eax
'006c6511    8d45d4                  lea eax, dword ptr [ebp-2c]
'006c6514    50                      push eax
'006c6515    ffd7                    call edi
Set var_44 = 
'006c6517    8b0e                    mov ecx, dword ptr [esi]
'006c6519    56                      push esi
'006c651a    898538ffffff            mov dword ptr [ebp+ffffff38], eax
'006c6520    ff9104030000            call dword ptr [ecx+00000304]
'006c6526    50                      push eax
'006c6527    8d55d8                  lea edx, dword ptr [ebp-28]
'006c652a    52                      push edx
'006c652b    ffd7                    call edi
Set var_45 = var_44
'006c652d    8d8d48ffffff            lea ecx, dword ptr [ebp+ffffff48]
'006c6533    8bf0                    mov esi, eax
'006c6535    8b06                    mov eax, dword ptr [esi]
'006c6537    51                      push ecx
'006c6538    56                      push esi
'006c6539    ff5078                  call dword ptr [eax+78]
'006c653c    dbe2                    fnclex
'006c653e    85c0                    test ax, ax
'006c6540    7d0b                    jge 6c654d
'006c6542    6a78                    push 78
'006c6544    681cb94300              push 0043b91c
'006c6549    56                      push esi
'006c654a    50                      push eax
'006c654b    ffd3                    call ebx
'006c654d    8b8548ffffff            mov eax, dword ptr [ebp+ffffff48]
'006c6553    8bb538ffffff            mov esi, dword ptr [ebp+ffffff38]
'006c6559    8b16                    mov edx, dword ptr [esi]
'006c655b    50                      push eax
'006c655c    56                      push esi
'006c655d    ff5274                  call dword ptr [edx+74]
'006c6560    dbe2                    fnclex
'006c6562    85c0                    test ax, ax
'006c6564    7d0b                    jge 6c6571
'006c6566    6a74                    push 74
'006c6568    6838eb4300              push 0043eb38
'006c656d    56                      push esi
'006c656e    50                      push eax
'006c656f    ffd3                    call ebx
'006c6571    8d4dd4                  lea ecx, dword ptr [ebp-2c]
'006c6574    51                      push ecx
'006c6575    8d55d8                  lea edx, dword ptr [ebp-28]
'006c6578    52                      push edx
'006c6579    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'006c657b    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_45, var_44)
'006c6581    83c40c                  add esp, 0c
'006c6584    c745fc00000000          mov dword ptr [ebp-04], 00000000
'006c658b    9b                      fwait
'006c658c    68c4656c00              push 006c65c4
'006c6591    eb30                    jmp 6c65c3
'006c6593    8d45d0                  lea eax, dword ptr [ebp-30]
'006c6596    50                      push eax
'006c6597    8d4dd4                  lea ecx, dword ptr [ebp-2c]
'006c659a    51                      push ecx
'006c659b    8d55d8                  lea edx, dword ptr [ebp-28]
'006c659e    52                      push edx
'006c659f    6a03                    push 03

' *** Reference to "__vbaFreeObjList"
'006c65a1    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_45, var_44, var_4)
'006c65a7    8d4590                  lea eax, dword ptr [ebp-70]
'006c65aa    50                      push eax
'006c65ab    8d4da0                  lea ecx, dword ptr [ebp-60]
'006c65ae    51                      push ecx
'006c65af    8d55b0                  lea edx, dword ptr [ebp-50]
'006c65b2    52                      push edx
'006c65b3    8d45c0                  lea eax, dword ptr [ebp-40]
'006c65b6    50                      push eax
'006c65b7    6a04                    push 04

' *** Reference to "__vbaFreeVarList"
'006c65b9    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_5, var_6, var_7, var_8)
'006c65bf    83c424                  add esp, 24
'006c65c2    c3                      ret
'006c65c3    c3                      ret
'006c65c4    8b4508                  mov eax, dword ptr [ebp+08]
'006c65c7    8b08                    mov ecx, dword ptr [eax]
'006c65c9    50                      push eax
'006c65ca    ff5108                  call dword ptr [ecx+08]
'006c65cd    8b45fc                  mov eax, dword ptr [ebp-04]
'006c65d0    8b4dec                  mov ecx, dword ptr [ebp-14]
'006c65d3    5f                      pop edi
'006c65d4    5e                      pop esi
    'Reference to '__except_list'
'006c65d5    64890d00000000          mov dword ptr fs:[00000000], ecx
'006c65dc    5b                      pop ebx
'006c65dd    8be5                    mov esp, ebp
'006c65df    5d                      pop ebp
'006c65e0    c20400                  ret 0004


End Sub


Private Sub Form_Activate()
'006c3410    55                      push ebp
'006c3411    8bec                    mov ebp, esp
'006c3413    83ec0c                  sub esp, 0c

' *** Reference to "__vbaExceptHandler"
'006c3416    6866724000              push 00407266
'006c341b    64a100000000            mov ax, word ptr fs:[00000000]
'006c3421    50                      push eax
    'Reference to '__except_list'
'006c3422    64892500000000          mov dword ptr fs:[00000000], esp
'006c3429    81ec80000000            sub esp, 00000080
'006c342f    53                      push ebx
'006c3430    56                      push esi
'006c3431    57                      push edi
'006c3432    8965f4                  mov dword ptr [ebp-0c], esp
'006c3435    c745f860664000          mov dword ptr [ebp-08], 00406660
'006c343c    8b7508                  mov esi, dword ptr [ebp+08]
'006c343f    8bc6                    mov eax, esi
'006c3441    83e001                  and eax, 01
'006c3444    8945fc                  mov dword ptr [ebp-04], eax
'006c3447    83e6fe                  and esi, fffffffe
'006c344a    8b0e                    mov ecx, dword ptr [esi]
'006c344c    56                      push esi
'006c344d    897508                  mov dword ptr [ebp+08], esi
'006c3450    ff5104                  call dword ptr [ecx+04]
'006c3453    33db                    xor ebx, ebx
'006c3455    66395e34                cmp word ptr [esi+34], bx
'006c3459    895de8                  mov dword ptr [ebp-18], ebx
'006c345c    895de4                  mov dword ptr [ebp-1c], ebx
'006c345f    895de0                  mov dword ptr [ebp-20], ebx
'006c3462    895ddc                  mov dword ptr [ebp-24], ebx
'006c3465    895dd8                  mov dword ptr [ebp-28], ebx
'006c3468    895dc8                  mov dword ptr [ebp-38], ebx
'006c346b    895db8                  mov dword ptr [ebp-48], ebx
'006c346e    895da8                  mov dword ptr [ebp-58], ebx
'006c3471    895d98                  mov dword ptr [ebp-68], ebx
'006c3474    895d94                  mov dword ptr [ebp-6c], ebx
'006c3477    0f84f8030000            je 6c3875

If (arg_6 <> 0) Then
'006c347d    8b16                    mov edx, dword ptr [esi]
'006c347f    56                      push esi
'006c3480    ff92fc020000            call dword ptr [edx+000002fc]
'006c3486    50                      push eax
'006c3487    8d45e0                  lea eax, dword ptr [ebp-20]
'006c348a    50                      push eax

' *** Reference to "__vbaObjSet"
'006c348b    ff15b4104000            call dword ptr [004010b4]
    Set var_3 = Me
'006c3491    8bf8                    mov edi, eax
'006c3493    8b0f                    mov ecx, dword ptr [edi]
'006c3495    6aff                    push ffffffff
'006c3497    57                      push edi
'006c3498    ff918c000000            call dword ptr [ecx+0000008c]
'006c349e    dbe2                    fnclex
'006c34a0    3bc3                    cmp eax, ebx
'006c34a2    7d12                    jge 6c34b6
    
    If (    var_3 < 0) Then
'006c34a4    688c000000              push 0000008c
'006c34a9    6838eb4300              push 0043eb38
'006c34ae    57                      push edi
'006c34af    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c34b0    ff1580104000            call dword ptr [00401080]
    
End If

' *** Reference to "__vbaFreeObj"
'006c34b6    8b1d08134000            mov ebx, dword ptr [00401308]
'006c34bc    8d4de0                  lea ecx, dword ptr [ebp-20]
'006c34bf    ffd3                    call ebx
'#Cleanup(var_3)
'006c34c1    8b16                    mov edx, dword ptr [esi]
'006c34c3    56                      push esi
'006c34c4    ff9200030000            call dword ptr [edx+00000300]
'006c34ca    50                      push eax
'006c34cb    8d45e0                  lea eax, dword ptr [ebp-20]
'006c34ce    50                      push eax

' *** Reference to "__vbaObjSet"
'006c34cf    ff15b4104000            call dword ptr [004010b4]
Set var_3 = var_3
'006c34d5    8bf8                    mov edi, eax
'006c34d7    8b0f                    mov ecx, dword ptr [edi]
'006c34d9    6aff                    push ffffffff
'006c34db    57                      push edi
'006c34dc    ff918c000000            call dword ptr [ecx+0000008c]
'006c34e2    dbe2                    fnclex
'006c34e4    85c0                    test ax, ax
'006c34e6    7d12                    jge 6c34fa
'006c34e8    688c000000              push 0000008c
'006c34ed    6838eb4300              push 0043eb38
'006c34f2    57                      push edi
'006c34f3    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c34f4    ff1580104000            call dword ptr [00401080]
'006c34fa    8d4de0                  lea ecx, dword ptr [ebp-20]
'006c34fd    ffd3                    call ebx
'#Cleanup(var_3)
'006c34ff    8b16                    mov edx, dword ptr [esi]
'006c3501    56                      push esi
'006c3502    ff9204030000            call dword ptr [edx+00000304]
'006c3508    50                      push eax
'006c3509    8d45e0                  lea eax, dword ptr [ebp-20]
'006c350c    50                      push eax

' *** Reference to "__vbaObjSet"
'006c350d    ff15b4104000            call dword ptr [004010b4]
Set var_3 = var_3
'006c3513    8bf8                    mov edi, eax
'006c3515    8b0f                    mov ecx, dword ptr [edi]
'006c3517    57                      push edi
'006c3518    ff91e8010000            call dword ptr [ecx+000001e8]
'006c351e    dbe2                    fnclex
'006c3520    85c0                    test ax, ax
'006c3522    7d12                    jge 6c3536
'006c3524    68e8010000              push 000001e8
'006c3529    681cb94300              push 0043b91c
'006c352e    57                      push edi
'006c352f    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c3530    ff1580104000            call dword ptr [00401080]
'006c3536    8d4de0                  lea ecx, dword ptr [ebp-20]
'006c3539    ffd3                    call ebx
'#Cleanup(var_3)
'006c353b    8d5de0                  lea ebx, dword ptr [ebp-20]
'006c353e    53                      push ebx
'006c353f    83ec10                  sub esp, 10
'006c3542    8bdc                    mov ebx, esp
'006c3544    b90a000000              mov ecx, 0000000a
'006c3549    890b                    mov dword ptr [ebx], ecx
'006c354b    894da8                  mov dword ptr [ebp-58], ecx
'006c354e    8b4d9c                  mov ecx, dword ptr [ebp-64]
'006c3551    894b04                  mov dword ptr [ebx+04], ecx
'006c3554    8b3d4cb07200            mov edi, dword ptr [0072b04c]
'006c355a    8b3f                    mov edi, dword ptr [edi]
'006c355c    b804000280              mov eax, 80020004
'006c3561    894308                  mov dword ptr [ebx+08], eax
'006c3564    8bd0                    mov edx, eax
'006c3566    8b45a4                  mov eax, dword ptr [ebp-5c]
'006c3569    89430c                  mov dword ptr [ebx+0c], eax
'006c356c    8b45a8                  mov eax, dword ptr [ebp-58]
'006c356f    83ec10                  sub esp, 10
'006c3572    8bcc                    mov ecx, esp
'006c3574    8901                    mov dword ptr [ecx], eax
'006c3576    8b45ac                  mov eax, dword ptr [ebp-54]
'006c3579    894104                  mov dword ptr [ecx+04], eax
'006c357c    895108                  mov dword ptr [ecx+08], edx
'006c357f    8b55b4                  mov edx, dword ptr [ebp-4c]
'006c3582    89510c                  mov dword ptr [ecx+0c], edx
'006c3585    8b55bc                  mov edx, dword ptr [ebp-44]
'006c3588    83ec10                  sub esp, 10
'006c358b    8bcc                    mov ecx, esp
'006c358d    b803000000              mov eax, 00000003
'006c3592    8901                    mov dword ptr [ecx], eax
'006c3594    895104                  mov dword ptr [ecx+04], edx
'006c3597    b802000000              mov eax, 00000002
'006c359c    894108                  mov dword ptr [ecx+08], eax
'006c359f    8b45c4                  mov eax, dword ptr [ebp-3c]
'006c35a2    89410c                  mov dword ptr [ecx+0c], eax
'006c35a5    8b0d4cb07200            mov ecx, dword ptr [0072b04c]
'006c35ab    685c324500              push 0045325c
'006c35b0    51                      push ecx
'006c35b1    ff97bc000000            call dword ptr [edi+000000bc]
'006c35b7    dbe2                    fnclex
'006c35b9    85c0                    test ax, ax
'006c35bb    7d1c                    jge 6c35d9
'006c35bd    8b154cb07200            mov edx, dword ptr [0072b04c]

' *** Reference to "__vbaHresultCheckObj"
'006c35c3    8b1d80104000            mov ebx, dword ptr [00401080]
'006c35c9    68bc000000              push 000000bc
'006c35ce    68301f4300              push 00431f30
'006c35d3    52                      push edx
'006c35d4    50                      push eax
'006c35d5    ffd3                    call ebx
'006c35d7    eb06                    jmp 6c35df

' *** Reference to "__vbaHresultCheckObj"
'006c35d9    8b1d80104000            mov ebx, dword ptr [00401080]
'006c35df    8b45e0                  mov eax, dword ptr [ebp-20]
'006c35e2    50                      push eax
'006c35e3    8d45e8                  lea eax, dword ptr [ebp-18]
'006c35e6    50                      push eax
'006c35e7    c745e000000000          mov dword ptr [ebp-20], 00000000

' *** Reference to "__vbaObjSet"
'006c35ee    ff15b4104000            call dword ptr [004010b4]
Set var_41 = var_3
'006c35f4    8b45e8                  mov eax, dword ptr [ebp-18]
'006c35f7    8b08                    mov ecx, dword ptr [eax]
'006c35f9    8d5594                  lea edx, dword ptr [ebp-6c]
'006c35fc    52                      push edx
'006c35fd    50                      push eax
'006c35fe    ff5134                  call dword ptr [ecx+34]
'006c3601    dbe2                    fnclex
'006c3603    85c0                    test ax, ax
'006c3605    7d0e                    jge 6c3615
'006c3607    8b4de8                  mov ecx, dword ptr [ebp-18]
'006c360a    6a34                    push 34
'006c360c    6830314300              push 00433130
'006c3611    51                      push ecx
'006c3612    50                      push eax
'006c3613    ffd3                    call ebx
'006c3615    66837d9400              cmp word ptr [ebp-6c], 00
'006c361a    0f8569010000            jne 6c3789

Do While (0 = 0)
'006c3620    8b16                    mov edx, dword ptr [esi]
'006c3622    56                      push esi
'006c3623    ff9204030000            call dword ptr [edx+00000304]
'006c3629    50                      push eax
'006c362a    8d45d8                  lea eax, dword ptr [ebp-28]
'006c362d    50                      push eax

' *** Reference to "__vbaObjSet"
'006c362e    ff15b4104000            call dword ptr [004010b4]
    Set var_45 = var_41
'006c3634    8d55e0                  lea edx, dword ptr [ebp-20]
'006c3637    89857cffffff            mov dword ptr [ebp+ffffff7c], eax
'006c363d    8b45e8                  mov eax, dword ptr [ebp-18]
'006c3640    8b08                    mov ecx, dword ptr [eax]
'006c3642    52                      push edx
'006c3643    50                      push eax
'006c3644    ff91b4000000            call dword ptr [ecx+000000b4]
'006c364a    dbe2                    fnclex
'006c364c    85c0                    test ax, ax
'006c364e    7d11                    jge 6c3661
'006c3650    8b4de8                  mov ecx, dword ptr [ebp-18]
'006c3653    68b4000000              push 000000b4
'006c3658    6830314300              push 00433130
'006c365d    51                      push ecx
'006c365e    50                      push eax
'006c365f    ffd3                    call ebx
'006c3661    8b45e0                  mov eax, dword ptr [ebp-20]
'006c3664    8b38                    mov edi, dword ptr [eax]
'006c3666    8d5ddc                  lea ebx, dword ptr [ebp-24]
'006c3669    53                      push ebx
'006c366a    83ec10                  sub esp, 10
'006c366d    8bdc                    mov ebx, esp
'006c366f    ba08000000              mov edx, 00000008
'006c3674    8913                    mov dword ptr [ebx], edx
'006c3676    8b55bc                  mov edx, dword ptr [ebp-44]
'006c3679    895304                  mov dword ptr [ebx+04], edx
'006c367c    b964b14300              mov ecx, 0043b164
'006c3681    894b08                  mov dword ptr [ebx+08], ecx
'006c3684    8b4dc4                  mov ecx, dword ptr [ebp-3c]
'006c3687    50                      push eax
'006c3688    89458c                  mov dword ptr [ebp-74], eax
'006c368b    894b0c                  mov dword ptr [ebx+0c], ecx
'006c368e    ff5730                  call dword ptr [edi+30]
'006c3691    dbe2                    fnclex
'006c3693    85c0                    test ax, ax
'006c3695    7d12                    jge 6c36a9
'006c3697    8b558c                  mov edx, dword ptr [ebp-74]
'006c369a    6a30                    push 30
'006c369c    68d8304300              push 004330d8
'006c36a1    52                      push edx
'006c36a2    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c36a3    ff1580104000            call dword ptr [00401080]
'006c36a9    8b45dc                  mov eax, dword ptr [ebp-24]
'006c36ac    8b08                    mov ecx, dword ptr [eax]
'006c36ae    8d55c8                  lea edx, dword ptr [ebp-38]
'006c36b1    52                      push edx
'006c36b2    50                      push eax
'006c36b3    8bf8                    mov edi, eax
'006c36b5    ff5144                  call dword ptr [ecx+44]
'006c36b8    dbe2                    fnclex
'006c36ba    85c0                    test ax, ax
'006c36bc    7d0f                    jge 6c36cd
'006c36be    6a44                    push 44
'006c36c0    6880304300              push 00433080
'006c36c5    57                      push edi
'006c36c6    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c36c7    ff1580104000            call dword ptr [00401080]
'006c36cd    8b55ac                  mov edx, dword ptr [ebp-54]
'006c36d0    8bbd7cffffff            mov edi, dword ptr [ebp+ffffff7c]
'006c36d6    8b1f                    mov ebx, dword ptr [edi]
'006c36d8    83ec10                  sub esp, 10
'006c36db    8bcc                    mov ecx, esp
'006c36dd    b80a000000              mov eax, 0000000a
'006c36e2    8901                    mov dword ptr [ecx], eax
'006c36e4    895104                  mov dword ptr [ecx+04], edx
'006c36e7    b804000280              mov eax, 80020004
'006c36ec    894108                  mov dword ptr [ecx+08], eax
'006c36ef    8b45b4                  mov eax, dword ptr [ebp-4c]
'006c36f2    89410c                  mov dword ptr [ecx+0c], eax
'006c36f5    8d4dc8                  lea ecx, dword ptr [ebp-38]
'006c36f8    51                      push ecx

' *** Reference to "__vbaStrVarMove"
'006c36f9    ff153c104000            call dword ptr [0040103c]
'006c36ff    8bd0                    mov edx, eax
'006c3701    8d4de4                  lea ecx, dword ptr [ebp-1c]

' *** Reference to "__vbaStrMove"
'006c3704    ff15d0124000            call dword ptr [004012d0]
'006c370a    50                      push eax
'006c370b    57                      push edi
'006c370c    ff93ec010000            call dword ptr [ebx+000001ec]
'006c3712    dbe2                    fnclex
'006c3714    85c0                    test ax, ax
'006c3716    7d16                    jge 6c372e

' *** Reference to "__vbaHresultCheckObj"
'006c3718    8b1d80104000            mov ebx, dword ptr [00401080]
'006c371e    68ec010000              push 000001ec
'006c3723    681cb94300              push 0043b91c
'006c3728    57                      push edi
'006c3729    50                      push eax
'006c372a    ffd3                    call ebx
'006c372c    eb06                    jmp 6c3734

' *** Reference to "__vbaHresultCheckObj"
'006c372e    8b1d80104000            mov ebx, dword ptr [00401080]
'006c3734    8d4de4                  lea ecx, dword ptr [ebp-1c]

' *** Reference to "__vbaFreeStr"
'006c3737    ff150c134000            call dword ptr [0040130c]
    '#Cleanup(var_40)
'006c373d    8d55d8                  lea edx, dword ptr [ebp-28]
'006c3740    52                      push edx
'006c3741    8d45dc                  lea eax, dword ptr [ebp-24]
'006c3744    50                      push eax
'006c3745    8d4de0                  lea ecx, dword ptr [ebp-20]
'006c3748    51                      push ecx
'006c3749    6a03                    push 03

' *** Reference to "__vbaFreeObjList"
'006c374b    ff154c104000            call dword ptr [0040104c]
    '#Cleanup( var_3, var_39, var_45)
'006c3751    83c410                  add esp, 10
'006c3754    8d4dc8                  lea ecx, dword ptr [ebp-38]

' *** Reference to "__vbaFreeVar"
'006c3757    ff1528104000            call dword ptr [00401028]
    '#Cleanup(var_46)
'006c375d    8b45e8                  mov eax, dword ptr [ebp-18]
'006c3760    8b10                    mov edx, dword ptr [eax]
'006c3762    50                      push eax
'006c3763    ff92ec000000            call dword ptr [edx+000000ec]
'006c3769    dbe2                    fnclex
'006c376b    85c0                    test ax, ax
'006c376d    0f8d81feffff            jge 6c35f4
'006c3773    8b4de8                  mov ecx, dword ptr [ebp-18]
'006c3776    68ec000000              push 000000ec
'006c377b    6830314300              push 00433130
'006c3780    51                      push ecx
'006c3781    50                      push eax
'006c3782    ffd3                    call ebx
'006c3784    e96bfeffff              jmp 6c35f4
    
Loop
'006c3789    8b45e8                  mov eax, dword ptr [ebp-18]
'006c378c    8b10                    mov edx, dword ptr [eax]
'006c378e    50                      push eax
'006c378f    ff92c4000000            call dword ptr [edx+000000c4]
'006c3795    dbe2                    fnclex
'006c3797    85c0                    test ax, ax
'006c3799    7d11                    jge 6c37ac
'006c379b    8b4de8                  mov ecx, dword ptr [ebp-18]
'006c379e    68c4000000              push 000000c4
'006c37a3    6830314300              push 00433130
'006c37a8    51                      push ecx
'006c37a9    50                      push eax
'006c37aa    ffd3                    call ebx
'006c37ac    8b16                    mov edx, dword ptr [esi]
'006c37ae    56                      push esi
'006c37af    ff9204030000            call dword ptr [edx+00000304]
'006c37b5    50                      push eax
'006c37b6    8d45e0                  lea eax, dword ptr [ebp-20]
'006c37b9    50                      push eax

' *** Reference to "__vbaObjSet"
'006c37ba    ff15b4104000            call dword ptr [004010b4]
Set var_3 = var_41
'006c37c0    8bf8                    mov edi, eax
'006c37c2    8b0f                    mov ecx, dword ptr [edi]
'006c37c4    6a00                    push 00
'006c37c6    57                      push edi
'006c37c7    ff91f4000000            call dword ptr [ecx+000000f4]
'006c37cd    dbe2                    fnclex
'006c37cf    85c0                    test ax, ax
'006c37d1    7d0e                    jge 6c37e1
'006c37d3    68f4000000              push 000000f4
'006c37d8    681cb94300              push 0043b91c
'006c37dd    57                      push edi
'006c37de    50                      push eax
'006c37df    ffd3                    call ebx
'006c37e1    8d4de0                  lea ecx, dword ptr [ebp-20]

' *** Reference to "__vbaFreeObj"
'006c37e4    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_3)
'006c37ea    8b16                    mov edx, dword ptr [esi]
'006c37ec    56                      push esi
'006c37ed    ff923c030000            call dword ptr [edx+0000033c]

' *** Reference to "__vbaObjSet"
'006c37f3    8b3db4104000            mov edi, dword ptr [004010b4]
'006c37f9    50                      push eax
'006c37fa    8d45dc                  lea eax, dword ptr [ebp-24]
'006c37fd    50                      push eax
'006c37fe    ffd7                    call edi
Set var_39 = var_3
'006c3800    8b0e                    mov ecx, dword ptr [esi]
'006c3802    56                      push esi
'006c3803    894588                  mov dword ptr [ebp-78], eax
'006c3806    ff9104030000            call dword ptr [ecx+00000304]
'006c380c    50                      push eax
'006c380d    8d55e0                  lea edx, dword ptr [ebp-20]
'006c3810    52                      push edx
'006c3811    ffd7                    call edi
Set var_3 = var_39
'006c3813    8d4de4                  lea ecx, dword ptr [ebp-1c]
'006c3816    8bf8                    mov edi, eax
'006c3818    8b07                    mov eax, dword ptr [edi]
'006c381a    51                      push ecx
'006c381b    57                      push edi
'006c381c    ff90a8000000            call dword ptr [eax+000000a8]
'006c3822    dbe2                    fnclex
'006c3824    85c0                    test ax, ax
'006c3826    7d0e                    jge 6c3836
'006c3828    68a8000000              push 000000a8
'006c382d    681cb94300              push 0043b91c
'006c3832    57                      push edi
'006c3833    50                      push eax
'006c3834    ffd3                    call ebx
'006c3836    8b45e4                  mov eax, dword ptr [ebp-1c]
'006c3839    8b7d88                  mov edi, dword ptr [ebp-78]
'006c383c    8b17                    mov edx, dword ptr [edi]
'006c383e    50                      push eax
'006c383f    57                      push edi
'006c3840    ff5254                  call dword ptr [edx+54]
'006c3843    dbe2                    fnclex
'006c3845    85c0                    test ax, ax
'006c3847    7d0b                    jge 6c3854
'006c3849    6a54                    push 54
'006c384b    683ce04300              push 0043e03c
'006c3850    57                      push edi
'006c3851    50                      push eax
'006c3852    ffd3                    call ebx
'006c3854    8d4de4                  lea ecx, dword ptr [ebp-1c]

' *** Reference to "__vbaFreeStr"
'006c3857    ff150c134000            call dword ptr [0040130c]
'#Cleanup(var_40)
'006c385d    8d4ddc                  lea ecx, dword ptr [ebp-24]
'006c3860    51                      push ecx
'006c3861    8d55e0                  lea edx, dword ptr [ebp-20]
'006c3864    52                      push edx
'006c3865    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'006c3867    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_3, var_39)
'006c386d    83c40c                  add esp, 0c
'006c3870    e9c7000000              jmp 6c393c

'ERROR: Two many next close:
End If
'006c3875    8b06                    mov eax, dword ptr [esi]
'006c3877    56                      push esi
'006c3878    ff90fc020000            call dword ptr [eax+000002fc]
'006c387e    50                      push eax
'006c387f    8d4de0                  lea ecx, dword ptr [ebp-20]
'006c3882    51                      push ecx

' *** Reference to "__vbaObjSet"
'006c3883    ff15b4104000            call dword ptr [004010b4]
Set var_3 = Nothing
'006c3889    8bf8                    mov edi, eax
'006c388b    8b17                    mov edx, dword ptr [edi]
'006c388d    53                      push ebx
'006c388e    57                      push edi
'006c388f    ff928c000000            call dword ptr [edx+0000008c]
'006c3895    dbe2                    fnclex
'006c3897    3bc3                    cmp eax, ebx
'006c3899    7d16                    jge 6c38b1

If (var_3 < 0) Then

' *** Reference to "__vbaHresultCheckObj"
'006c389b    8b1d80104000            mov ebx, dword ptr [00401080]
'006c38a1    688c000000              push 0000008c
'006c38a6    6838eb4300              push 0043eb38
'006c38ab    57                      push edi
'006c38ac    50                      push eax
'006c38ad    ffd3                    call ebx
'006c38af    eb06                    jmp 6c38b7
    
End If

' *** Reference to "__vbaHresultCheckObj"
'006c38b1    8b1d80104000            mov ebx, dword ptr [00401080]
'006c38b7    8d4de0                  lea ecx, dword ptr [ebp-20]

' *** Reference to "__vbaFreeObj"
'006c38ba    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_3)
'006c38c0    8b06                    mov eax, dword ptr [esi]
'006c38c2    56                      push esi
'006c38c3    ff9000030000            call dword ptr [eax+00000300]
'006c38c9    50                      push eax
'006c38ca    8d4de0                  lea ecx, dword ptr [ebp-20]
'006c38cd    51                      push ecx

' *** Reference to "__vbaObjSet"
'006c38ce    ff15b4104000            call dword ptr [004010b4]
Set var_3 = Nothing
'006c38d4    8bf8                    mov edi, eax
'006c38d6    8b17                    mov edx, dword ptr [edi]
'006c38d8    6a00                    push 00
'006c38da    57                      push edi
'006c38db    ff928c000000            call dword ptr [edx+0000008c]
'006c38e1    dbe2                    fnclex
'006c38e3    85c0                    test ax, ax
'006c38e5    7d0e                    jge 6c38f5
'006c38e7    688c000000              push 0000008c
'006c38ec    6838eb4300              push 0043eb38
'006c38f1    57                      push edi
'006c38f2    50                      push eax
'006c38f3    ffd3                    call ebx
'006c38f5    8d4de0                  lea ecx, dword ptr [ebp-20]

' *** Reference to "__vbaFreeObj"
'006c38f8    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_3)
'006c38fe    8b06                    mov eax, dword ptr [esi]
'006c3900    56                      push esi
'006c3901    ff9004030000            call dword ptr [eax+00000304]
'006c3907    50                      push eax
'006c3908    8d4de0                  lea ecx, dword ptr [ebp-20]
'006c390b    51                      push ecx

' *** Reference to "__vbaObjSet"
'006c390c    ff15b4104000            call dword ptr [004010b4]
Set var_3 = Nothing
'006c3912    8bf8                    mov edi, eax
'006c3914    8b17                    mov edx, dword ptr [edi]
'006c3916    6a00                    push 00
'006c3918    57                      push edi
'006c3919    ff9294000000            call dword ptr [edx+00000094]
'006c391f    dbe2                    fnclex
'006c3921    85c0                    test ax, ax
'006c3923    7d0e                    jge 6c3933
'006c3925    6894000000              push 00000094
'006c392a    681cb94300              push 0043b91c
'006c392f    57                      push edi
'006c3930    50                      push eax
'006c3931    ffd3                    call ebx
'006c3933    8d4de0                  lea ecx, dword ptr [ebp-20]

' *** Reference to "__vbaFreeObj"
'006c3936    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_3)
'006c393c    8b06                    mov eax, dword ptr [esi]
'006c393e    56                      push esi
'006c393f    ff9000070000            call dword ptr [eax+00000700]
'006c3945    85c0                    test ax, ax
'006c3947    7d0e                    jge 6c3957
'006c3949    6800070000              push 00000700
'006c394e    6888124300              push 00431288
'006c3953    56                      push esi
'006c3954    50                      push eax
'006c3955    ffd3                    call ebx
'006c3957    c745fc00000000          mov dword ptr [ebp-04], 00000000
'006c395e    6899396c00              push 006c3999
'006c3963    eb2a                    jmp 6c398f
'006c3965    8d4de4                  lea ecx, dword ptr [ebp-1c]

' *** Reference to "__vbaFreeStr"
'006c3968    ff150c134000            call dword ptr [0040130c]
'#Cleanup(var_40)
'006c396e    8d4dd8                  lea ecx, dword ptr [ebp-28]
'006c3971    51                      push ecx
'006c3972    8d55dc                  lea edx, dword ptr [ebp-24]
'006c3975    52                      push edx
'006c3976    8d45e0                  lea eax, dword ptr [ebp-20]
'006c3979    50                      push eax
'006c397a    6a03                    push 03

' *** Reference to "__vbaFreeObjList"
'006c397c    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_3, var_39, var_45)
'006c3982    83c410                  add esp, 10
'006c3985    8d4dc8                  lea ecx, dword ptr [ebp-38]

' *** Reference to "__vbaFreeVar"
'006c3988    ff1528104000            call dword ptr [00401028]
'#Cleanup(var_46)
'006c398e    c3                      ret
'006c398f    8d4de8                  lea ecx, dword ptr [ebp-18]

' *** Reference to "__vbaFreeObj"
'006c3992    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_41)
'006c3998    c3                      ret
'006c3999    8b4508                  mov eax, dword ptr [ebp+08]
'006c399c    8b08                    mov ecx, dword ptr [eax]
'006c399e    50                      push eax
'006c399f    ff5108                  call dword ptr [ecx+08]
'006c39a2    8b45fc                  mov eax, dword ptr [ebp-04]
'006c39a5    8b4dec                  mov ecx, dword ptr [ebp-14]
'006c39a8    5f                      pop edi
'006c39a9    5e                      pop esi
    'Reference to '__except_list'
'006c39aa    64890d00000000          mov dword ptr fs:[00000000], ecx
'006c39b1    5b                      pop ebx
'006c39b2    8be5                    mov esp, ebp
'006c39b4    5d                      pop ebp
'006c39b5    c20400                  ret 0004


End Sub


'Event for btnEnregistrer
Private Sub btnEnregistrer_Click()
'006c1270    55                      push ebp
'006c1271    8bec                    mov ebp, esp
'006c1273    83ec0c                  sub esp, 0c

' *** Reference to "__vbaExceptHandler"
'006c1276    6866724000              push 00407266
'006c127b    64a100000000            mov ax, word ptr fs:[00000000]
'006c1281    50                      push eax
    'Reference to '__except_list'
'006c1282    64892500000000          mov dword ptr fs:[00000000], esp
'006c1289    81ec00010000            sub esp, 00000100
'006c128f    53                      push ebx
'006c1290    56                      push esi
'006c1291    57                      push edi
'006c1292    8965f4                  mov dword ptr [ebp-0c], esp
'006c1295    c745f8f0654000          mov dword ptr [ebp-08], 004065f0
'006c129c    8b7d08                  mov edi, dword ptr [ebp+08]
'006c129f    8bc7                    mov eax, edi
'006c12a1    83e001                  and eax, 01
'006c12a4    8945fc                  mov dword ptr [ebp-04], eax
'006c12a7    83e7fe                  and edi, fffffffe
'006c12aa    8b0f                    mov ecx, dword ptr [edi]
'006c12ac    57                      push edi
'006c12ad    897d08                  mov dword ptr [ebp+08], edi
'006c12b0    ff5104                  call dword ptr [ecx+04]
'006c12b3    8b17                    mov edx, dword ptr [edi]
'006c12b5    33f6                    xor esi, esi
'006c12b7    57                      push edi
'006c12b8    8975e8                  mov dword ptr [ebp-18], esi
'006c12bb    8975e4                  mov dword ptr [ebp-1c], esi
'006c12be    8975e0                  mov dword ptr [ebp-20], esi
'006c12c1    8975dc                  mov dword ptr [ebp-24], esi
'006c12c4    8975d8                  mov dword ptr [ebp-28], esi
'006c12c7    8975d4                  mov dword ptr [ebp-2c], esi
'006c12ca    8975d0                  mov dword ptr [ebp-30], esi
'006c12cd    8975cc                  mov dword ptr [ebp-34], esi
'006c12d0    8975c8                  mov dword ptr [ebp-38], esi
'006c12d3    8975c4                  mov dword ptr [ebp-3c], esi
'006c12d6    8975c0                  mov dword ptr [ebp-40], esi
'006c12d9    8975bc                  mov dword ptr [ebp-44], esi
'006c12dc    8975b8                  mov dword ptr [ebp-48], esi
'006c12df    8975a8                  mov dword ptr [ebp-58], esi
'006c12e2    897598                  mov dword ptr [ebp-68], esi
'006c12e5    897588                  mov dword ptr [ebp-78], esi
'006c12e8    89b578ffffff            mov dword ptr [ebp+ffffff78], esi
'006c12ee    89b568ffffff            mov dword ptr [ebp+ffffff68], esi
'006c12f4    89b558ffffff            mov dword ptr [ebp+ffffff58], esi
'006c12fa    ff921c030000            call dword ptr [edx+0000031c]

' *** Reference to "__vbaObjSet"
'006c1300    8b1db4104000            mov ebx, dword ptr [004010b4]
'006c1306    50                      push eax
'006c1307    8d45bc                  lea eax, dword ptr [ebp-44]
'006c130a    50                      push eax
'006c130b    ffd3                    call ebx
Set var_58 = Me
'006c130d    8b08                    mov ecx, dword ptr [eax]
'006c130f    8d55e8                  lea edx, dword ptr [ebp-18]
'006c1312    52                      push edx
'006c1313    50                      push eax
'006c1314    898554ffffff            mov dword ptr [ebp+ffffff54], eax
'006c131a    ff91a0000000            call dword ptr [ecx+000000a0]
'006c1320    dbe2                    fnclex
'006c1322    3bc6                    cmp eax, esi
'006c1324    7d18                    jge 6c133e

If (var_58 < 0) Then
'006c1326    8b8d54ffffff            mov ecx, dword ptr [ebp+ffffff54]
'006c132c    68a0000000              push 000000a0
'006c1331    68d00d4300              push 00430dd0
'006c1336    51                      push ecx
'006c1337    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c1338    ff1580104000            call dword ptr [00401080]
    
End If
'006c133e    8b55e8                  mov edx, dword ptr [ebp-18]
'006c1341    8975e8                  mov dword ptr [ebp-18], esi

' *** Reference to "__vbaStrMove"
'006c1344    8b35d0124000            mov esi, dword ptr [004012d0]
'006c134a    8d4de4                  lea ecx, dword ptr [ebp-1c]
'006c134d    ffd6                    call esi
'006c134f    8b17                    mov edx, dword ptr [edi]
'006c1351    8d45a8                  lea eax, dword ptr [ebp-58]
'006c1354    50                      push eax
'006c1355    8d4de4                  lea ecx, dword ptr [ebp-1c]
'006c1358    51                      push ecx
'006c1359    57                      push edi
'006c135a    ff9208070000            call dword ptr [edx+00000708]
'006c1360    8d4de4                  lea ecx, dword ptr [ebp-1c]

' *** Reference to "__vbaFreeStr"
'006c1363    ff150c134000            call dword ptr [0040130c]
'#Cleanup(var_40)
'006c1369    8d4dbc                  lea ecx, dword ptr [ebp-44]

' *** Reference to "__vbaFreeObj"
'006c136c    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_58)
'006c1372    8d4da8                  lea ecx, dword ptr [ebp-58]

' *** Reference to "__vbaFreeVar"
'006c1375    ff1528104000            call dword ptr [00401028]
'#Cleanup(var_86)
'006c137b    8b17                    mov edx, dword ptr [edi]
'006c137d    57                      push edi
'006c137e    ff921c030000            call dword ptr [edx+0000031c]
'006c1384    50                      push eax
'006c1385    8d45b8                  lea eax, dword ptr [ebp-48]
'006c1388    50                      push eax
'006c1389    ffd3                    call ebx
Set var_61 = 
'006c138b    8b0f                    mov ecx, dword ptr [edi]
'006c138d    57                      push edi
'006c138e    89854cffffff            mov dword ptr [ebp+ffffff4c], eax
'006c1394    ff911c030000            call dword ptr [ecx+0000031c]
'006c139a    50                      push eax
'006c139b    8d55bc                  lea edx, dword ptr [ebp-44]
'006c139e    52                      push edx
'006c139f    ffd3                    call ebx
Set var_58 = var_61
'006c13a1    8b08                    mov ecx, dword ptr [eax]
'006c13a3    8d55e8                  lea edx, dword ptr [ebp-18]
'006c13a6    52                      push edx
'006c13a7    50                      push eax
'006c13a8    898554ffffff            mov dword ptr [ebp+ffffff54], eax
'006c13ae    ff91a0000000            call dword ptr [ecx+000000a0]
'006c13b4    dbe2                    fnclex
'006c13b6    85c0                    test ax, ax
'006c13b8    7d18                    jge 6c13d2
'006c13ba    8b8d54ffffff            mov ecx, dword ptr [ebp+ffffff54]
'006c13c0    68a0000000              push 000000a0
'006c13c5    68d00d4300              push 00430dd0
'006c13ca    51                      push ecx
'006c13cb    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c13cc    ff1580104000            call dword ptr [00401080]
'006c13d2    8b55e8                  mov edx, dword ptr [ebp-18]
'006c13d5    8d4de4                  lea ecx, dword ptr [ebp-1c]
'006c13d8    c745e800000000          mov dword ptr [ebp-18], 00000000
'006c13df    ffd6                    call esi
'006c13e1    8b17                    mov edx, dword ptr [edi]
'006c13e3    8d45e0                  lea eax, dword ptr [ebp-20]
'006c13e6    50                      push eax
'006c13e7    8d4de4                  lea ecx, dword ptr [ebp-1c]
'006c13ea    51                      push ecx
'006c13eb    57                      push edi
'006c13ec    ff920c070000            call dword ptr [edx+0000070c]
'006c13f2    8b4de0                  mov ecx, dword ptr [ebp-20]
'006c13f5    8b854cffffff            mov eax, dword ptr [ebp+ffffff4c]
'006c13fb    8b10                    mov edx, dword ptr [eax]
'006c13fd    51                      push ecx
'006c13fe    50                      push eax
'006c13ff    ff92a4000000            call dword ptr [edx+000000a4]
'006c1405    dbe2                    fnclex
'006c1407    85c0                    test ax, ax
'006c1409    7d18                    jge 6c1423
'006c140b    8b954cffffff            mov edx, dword ptr [ebp+ffffff4c]
'006c1411    68a4000000              push 000000a4
'006c1416    68d00d4300              push 00430dd0
'006c141b    52                      push edx
'006c141c    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c141d    ff1580104000            call dword ptr [00401080]
'006c1423    8d45e0                  lea eax, dword ptr [ebp-20]
'006c1426    50                      push eax
'006c1427    8d4de4                  lea ecx, dword ptr [ebp-1c]
'006c142a    51                      push ecx
'006c142b    6a02                    push 02

' *** Reference to "__vbaFreeStrList"
'006c142d    ff155c124000            call dword ptr [0040125c]
'#Cleanup( 0, 0)
'006c1433    8d55b8                  lea edx, dword ptr [ebp-48]
'006c1436    52                      push edx
'006c1437    8d45bc                  lea eax, dword ptr [ebp-44]
'006c143a    50                      push eax
'006c143b    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'006c143d    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_58, var_61)
'006c1443    8b0f                    mov ecx, dword ptr [edi]
'006c1445    83c418                  add esp, 18
'006c1448    57                      push edi
'006c1449    ff910c030000            call dword ptr [ecx+0000030c]
'006c144f    50                      push eax
'006c1450    8d55b8                  lea edx, dword ptr [ebp-48]
'006c1453    52                      push edx
'006c1454    ffd3                    call ebx
Set var_61 = 
'006c1456    89854cffffff            mov dword ptr [ebp+ffffff4c], eax
'006c145c    8b07                    mov eax, dword ptr [edi]
'006c145e    57                      push edi
'006c145f    ff900c030000            call dword ptr [eax+0000030c]
'006c1465    50                      push eax
'006c1466    8d4dbc                  lea ecx, dword ptr [ebp-44]
'006c1469    51                      push ecx
'006c146a    ffd3                    call ebx
Set var_58 = Nothing
'006c146c    8b10                    mov edx, dword ptr [eax]
'006c146e    8d4de8                  lea ecx, dword ptr [ebp-18]
'006c1471    51                      push ecx
'006c1472    50                      push eax
'006c1473    898554ffffff            mov dword ptr [ebp+ffffff54], eax
'006c1479    ff92a0000000            call dword ptr [edx+000000a0]
'006c147f    dbe2                    fnclex
'006c1481    85c0                    test ax, ax
'006c1483    7d18                    jge 6c149d
'006c1485    8b9554ffffff            mov edx, dword ptr [ebp+ffffff54]
'006c148b    68a0000000              push 000000a0
'006c1490    68d00d4300              push 00430dd0
'006c1495    52                      push edx
'006c1496    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c1497    ff1580104000            call dword ptr [00401080]
'006c149d    8b55e8                  mov edx, dword ptr [ebp-18]
'006c14a0    8d4de4                  lea ecx, dword ptr [ebp-1c]
'006c14a3    c745e800000000          mov dword ptr [ebp-18], 00000000
'006c14aa    ffd6                    call esi
'006c14ac    8b07                    mov eax, dword ptr [edi]
'006c14ae    8d4de0                  lea ecx, dword ptr [ebp-20]
'006c14b1    51                      push ecx
'006c14b2    8d55e4                  lea edx, dword ptr [ebp-1c]
'006c14b5    52                      push edx
'006c14b6    57                      push edi
'006c14b7    ff900c070000            call dword ptr [eax+0000070c]
'006c14bd    8b55e0                  mov edx, dword ptr [ebp-20]
'006c14c0    8b854cffffff            mov eax, dword ptr [ebp+ffffff4c]
'006c14c6    8b08                    mov ecx, dword ptr [eax]
'006c14c8    52                      push edx
'006c14c9    50                      push eax
'006c14ca    ff91a4000000            call dword ptr [ecx+000000a4]
'006c14d0    dbe2                    fnclex
'006c14d2    85c0                    test ax, ax
'006c14d4    7d18                    jge 6c14ee
'006c14d6    8b8d4cffffff            mov ecx, dword ptr [ebp+ffffff4c]
'006c14dc    68a4000000              push 000000a4
'006c14e1    68d00d4300              push 00430dd0
'006c14e6    51                      push ecx
'006c14e7    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c14e8    ff1580104000            call dword ptr [00401080]
'006c14ee    8d55e0                  lea edx, dword ptr [ebp-20]
'006c14f1    52                      push edx
'006c14f2    8d45e4                  lea eax, dword ptr [ebp-1c]
'006c14f5    50                      push eax
'006c14f6    6a02                    push 02

' *** Reference to "__vbaFreeStrList"
'006c14f8    ff155c124000            call dword ptr [0040125c]
'#Cleanup( var_41, 0)
'006c14fe    8d4db8                  lea ecx, dword ptr [ebp-48]
'006c1501    51                      push ecx
'006c1502    8d55bc                  lea edx, dword ptr [ebp-44]
'006c1505    52                      push edx
'006c1506    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'006c1508    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_58, var_61)
'006c150e    8b07                    mov eax, dword ptr [edi]
'006c1510    83c418                  add esp, 18
'006c1513    57                      push edi
'006c1514    ff9018030000            call dword ptr [eax+00000318]
'006c151a    50                      push eax
'006c151b    8d4db8                  lea ecx, dword ptr [ebp-48]
'006c151e    51                      push ecx
'006c151f    ffd3                    call ebx
Set var_61 = Nothing
'006c1521    8b17                    mov edx, dword ptr [edi]
'006c1523    57                      push edi
'006c1524    89854cffffff            mov dword ptr [ebp+ffffff4c], eax
'006c152a    ff9218030000            call dword ptr [edx+00000318]
'006c1530    50                      push eax
'006c1531    8d45bc                  lea eax, dword ptr [ebp-44]
'006c1534    50                      push eax
'006c1535    ffd3                    call ebx
Set var_58 = Nothing
'006c1537    8b08                    mov ecx, dword ptr [eax]
'006c1539    8d55e8                  lea edx, dword ptr [ebp-18]
'006c153c    52                      push edx
'006c153d    50                      push eax
'006c153e    898554ffffff            mov dword ptr [ebp+ffffff54], eax
'006c1544    ff91a0000000            call dword ptr [ecx+000000a0]
'006c154a    dbe2                    fnclex
'006c154c    85c0                    test ax, ax
'006c154e    7d18                    jge 6c1568
'006c1550    8b8d54ffffff            mov ecx, dword ptr [ebp+ffffff54]
'006c1556    68a0000000              push 000000a0
'006c155b    68d00d4300              push 00430dd0
'006c1560    51                      push ecx
'006c1561    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c1562    ff1580104000            call dword ptr [00401080]
'006c1568    8b55e8                  mov edx, dword ptr [ebp-18]
'006c156b    8d4de4                  lea ecx, dword ptr [ebp-1c]
'006c156e    c745e800000000          mov dword ptr [ebp-18], 00000000
'006c1575    ffd6                    call esi
'006c1577    8b17                    mov edx, dword ptr [edi]
'006c1579    8d45e0                  lea eax, dword ptr [ebp-20]
'006c157c    50                      push eax
'006c157d    8d4de4                  lea ecx, dword ptr [ebp-1c]
'006c1580    51                      push ecx
'006c1581    57                      push edi
'006c1582    ff920c070000            call dword ptr [edx+0000070c]
'006c1588    8b4de0                  mov ecx, dword ptr [ebp-20]
'006c158b    8b854cffffff            mov eax, dword ptr [ebp+ffffff4c]
'006c1591    8b10                    mov edx, dword ptr [eax]
'006c1593    51                      push ecx
'006c1594    50                      push eax
'006c1595    ff92a4000000            call dword ptr [edx+000000a4]
'006c159b    dbe2                    fnclex
'006c159d    85c0                    test ax, ax
'006c159f    7d18                    jge 6c15b9
'006c15a1    8b954cffffff            mov edx, dword ptr [ebp+ffffff4c]
'006c15a7    68a4000000              push 000000a4
'006c15ac    68d00d4300              push 00430dd0
'006c15b1    52                      push edx
'006c15b2    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c15b3    ff1580104000            call dword ptr [00401080]
'006c15b9    8d45e0                  lea eax, dword ptr [ebp-20]
'006c15bc    50                      push eax
'006c15bd    8d4de4                  lea ecx, dword ptr [ebp-1c]
'006c15c0    51                      push ecx
'006c15c1    6a02                    push 02

' *** Reference to "__vbaFreeStrList"
'006c15c3    ff155c124000            call dword ptr [0040125c]
'#Cleanup( var_41, 0)
'006c15c9    8d55b8                  lea edx, dword ptr [ebp-48]
'006c15cc    52                      push edx
'006c15cd    8d45bc                  lea eax, dword ptr [ebp-44]
'006c15d0    50                      push eax
'006c15d1    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'006c15d3    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_58, var_61)
'006c15d9    8b0f                    mov ecx, dword ptr [edi]
'006c15db    83c418                  add esp, 18
'006c15de    57                      push edi
'006c15df    ff9114030000            call dword ptr [ecx+00000314]
'006c15e5    50                      push eax
'006c15e6    8d55b8                  lea edx, dword ptr [ebp-48]
'006c15e9    52                      push edx
'006c15ea    ffd3                    call ebx
Set var_61 = 
'006c15ec    89854cffffff            mov dword ptr [ebp+ffffff4c], eax
'006c15f2    8b07                    mov eax, dword ptr [edi]
'006c15f4    57                      push edi
'006c15f5    ff9014030000            call dword ptr [eax+00000314]
'006c15fb    50                      push eax
'006c15fc    8d4dbc                  lea ecx, dword ptr [ebp-44]
'006c15ff    51                      push ecx
'006c1600    ffd3                    call ebx
Set var_58 = Nothing
'006c1602    8b10                    mov edx, dword ptr [eax]
'006c1604    8d4de8                  lea ecx, dword ptr [ebp-18]
'006c1607    51                      push ecx
'006c1608    50                      push eax
'006c1609    898554ffffff            mov dword ptr [ebp+ffffff54], eax
'006c160f    ff92a0000000            call dword ptr [edx+000000a0]
'006c1615    dbe2                    fnclex
'006c1617    85c0                    test ax, ax
'006c1619    7d18                    jge 6c1633
'006c161b    8b9554ffffff            mov edx, dword ptr [ebp+ffffff54]
'006c1621    68a0000000              push 000000a0
'006c1626    68d00d4300              push 00430dd0
'006c162b    52                      push edx
'006c162c    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c162d    ff1580104000            call dword ptr [00401080]
'006c1633    8b55e8                  mov edx, dword ptr [ebp-18]
'006c1636    8d4de4                  lea ecx, dword ptr [ebp-1c]
'006c1639    c745e800000000          mov dword ptr [ebp-18], 00000000
'006c1640    ffd6                    call esi
'006c1642    8b07                    mov eax, dword ptr [edi]
'006c1644    8d4de0                  lea ecx, dword ptr [ebp-20]
'006c1647    51                      push ecx
'006c1648    8d55e4                  lea edx, dword ptr [ebp-1c]
'006c164b    52                      push edx
'006c164c    57                      push edi
'006c164d    ff900c070000            call dword ptr [eax+0000070c]
'006c1653    8b55e0                  mov edx, dword ptr [ebp-20]
'006c1656    8b854cffffff            mov eax, dword ptr [ebp+ffffff4c]
'006c165c    8b08                    mov ecx, dword ptr [eax]
'006c165e    52                      push edx
'006c165f    50                      push eax
'006c1660    ff91a4000000            call dword ptr [ecx+000000a4]
'006c1666    dbe2                    fnclex
'006c1668    85c0                    test ax, ax
'006c166a    7d18                    jge 6c1684
'006c166c    8b8d4cffffff            mov ecx, dword ptr [ebp+ffffff4c]
'006c1672    68a4000000              push 000000a4
'006c1677    68d00d4300              push 00430dd0
'006c167c    51                      push ecx
'006c167d    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c167e    ff1580104000            call dword ptr [00401080]
'006c1684    8d55e0                  lea edx, dword ptr [ebp-20]
'006c1687    52                      push edx
'006c1688    8d45e4                  lea eax, dword ptr [ebp-1c]
'006c168b    50                      push eax
'006c168c    6a02                    push 02

' *** Reference to "__vbaFreeStrList"
'006c168e    ff155c124000            call dword ptr [0040125c]
'#Cleanup( var_41, 0)
'006c1694    8d4db8                  lea ecx, dword ptr [ebp-48]
'006c1697    51                      push ecx
'006c1698    8d55bc                  lea edx, dword ptr [ebp-44]
'006c169b    52                      push edx
'006c169c    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'006c169e    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_58, var_61)
'006c16a4    8b07                    mov eax, dword ptr [edi]
'006c16a6    83c418                  add esp, 18
'006c16a9    57                      push edi
'006c16aa    ff9010030000            call dword ptr [eax+00000310]
'006c16b0    50                      push eax
'006c16b1    8d4db8                  lea ecx, dword ptr [ebp-48]
'006c16b4    51                      push ecx
'006c16b5    ffd3                    call ebx
Set var_61 = Nothing
'006c16b7    8b17                    mov edx, dword ptr [edi]
'006c16b9    57                      push edi
'006c16ba    89854cffffff            mov dword ptr [ebp+ffffff4c], eax
'006c16c0    ff9210030000            call dword ptr [edx+00000310]
'006c16c6    50                      push eax
'006c16c7    8d45bc                  lea eax, dword ptr [ebp-44]
'006c16ca    50                      push eax
'006c16cb    ffd3                    call ebx
Set var_58 = Nothing
'006c16cd    8b08                    mov ecx, dword ptr [eax]
'006c16cf    8d55e8                  lea edx, dword ptr [ebp-18]
'006c16d2    52                      push edx
'006c16d3    50                      push eax
'006c16d4    898554ffffff            mov dword ptr [ebp+ffffff54], eax
'006c16da    ff91a0000000            call dword ptr [ecx+000000a0]
'006c16e0    dbe2                    fnclex
'006c16e2    85c0                    test ax, ax
'006c16e4    7d18                    jge 6c16fe
'006c16e6    8b8d54ffffff            mov ecx, dword ptr [ebp+ffffff54]
'006c16ec    68a0000000              push 000000a0
'006c16f1    68d00d4300              push 00430dd0
'006c16f6    51                      push ecx
'006c16f7    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c16f8    ff1580104000            call dword ptr [00401080]
'006c16fe    8b55e8                  mov edx, dword ptr [ebp-18]
'006c1701    8d4de4                  lea ecx, dword ptr [ebp-1c]
'006c1704    c745e800000000          mov dword ptr [ebp-18], 00000000
'006c170b    ffd6                    call esi
'006c170d    8b17                    mov edx, dword ptr [edi]
'006c170f    8d45e0                  lea eax, dword ptr [ebp-20]
'006c1712    50                      push eax
'006c1713    8d4de4                  lea ecx, dword ptr [ebp-1c]
'006c1716    51                      push ecx
'006c1717    57                      push edi
'006c1718    ff920c070000            call dword ptr [edx+0000070c]
'006c171e    8b4de0                  mov ecx, dword ptr [ebp-20]
'006c1721    8b854cffffff            mov eax, dword ptr [ebp+ffffff4c]
'006c1727    8b10                    mov edx, dword ptr [eax]
'006c1729    51                      push ecx
'006c172a    50                      push eax
'006c172b    ff92a4000000            call dword ptr [edx+000000a4]
'006c1731    dbe2                    fnclex
'006c1733    85c0                    test ax, ax
'006c1735    7d18                    jge 6c174f
'006c1737    8b954cffffff            mov edx, dword ptr [ebp+ffffff4c]
'006c173d    68a4000000              push 000000a4
'006c1742    68d00d4300              push 00430dd0
'006c1747    52                      push edx
'006c1748    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c1749    ff1580104000            call dword ptr [00401080]
'006c174f    8d45e0                  lea eax, dword ptr [ebp-20]
'006c1752    50                      push eax
'006c1753    8d4de4                  lea ecx, dword ptr [ebp-1c]
'006c1756    51                      push ecx
'006c1757    6a02                    push 02

' *** Reference to "__vbaFreeStrList"
'006c1759    ff155c124000            call dword ptr [0040125c]
'#Cleanup( var_41, 0)
'006c175f    8d55b8                  lea edx, dword ptr [ebp-48]
'006c1762    52                      push edx
'006c1763    8d45bc                  lea eax, dword ptr [ebp-44]
'006c1766    50                      push eax
'006c1767    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'006c1769    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_58, var_61)
'006c176f    8b0f                    mov ecx, dword ptr [edi]
'006c1771    83c418                  add esp, 18
'006c1774    57                      push edi
'006c1775    ff913c030000            call dword ptr [ecx+0000033c]
'006c177b    50                      push eax
'006c177c    8d55bc                  lea edx, dword ptr [ebp-44]
'006c177f    52                      push edx
'006c1780    ffd3                    call ebx
Set var_58 = 
'006c1782    8b08                    mov ecx, dword ptr [eax]
'006c1784    8d55e8                  lea edx, dword ptr [ebp-18]
'006c1787    52                      push edx
'006c1788    50                      push eax
'006c1789    898554ffffff            mov dword ptr [ebp+ffffff54], eax
'006c178f    ff5150                  call dword ptr [ecx+50]
'006c1792    dbe2                    fnclex
'006c1794    85c0                    test ax, ax
'006c1796    7d15                    jge 6c17ad
'006c1798    8b8d54ffffff            mov ecx, dword ptr [ebp+ffffff54]
'006c179e    6a50                    push 50
'006c17a0    683ce04300              push 0043e03c
'006c17a5    51                      push ecx
'006c17a6    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c17a7    ff1580104000            call dword ptr [00401080]
'006c17ad    8b55e8                  mov edx, dword ptr [ebp-18]
'006c17b0    52                      push edx
'006c17b1    68cc134300              push 004313cc

' *** Reference to "__vbaStrCmp"
'006c17b6    ff1530114000            call dword ptr [00401130]
'006c17bc    f7d8                    neg eax
'006c17be    1bc0                    sbb eax, eax
'006c17c0    f7d8                    neg eax
'006c17c2    f7d8                    neg eax
'006c17c4    66898560ffffff          mov word ptr [ebp+ffffff60], ax
'006c17cb    8b07                    mov eax, dword ptr [edi]
'006c17cd    57                      push edi
'006c17ce    c78558ffffff0b000000    mov dword ptr [ebp+ffffff58], 0000000b
'006c17d8    ff901c030000            call dword ptr [eax+0000031c]
'006c17de    8d4da8                  lea ecx, dword ptr [ebp-58]
'006c17e1    51                      push ecx
'006c17e2    8d5598                  lea edx, dword ptr [ebp-68]
'006c17e5    52                      push edx
'006c17e6    8945b0                  mov dword ptr [ebp-50], eax
'006c17e9    c745a809000000          mov dword ptr [ebp-58], 00000009

' *** Reference to "rtcTrimVar"
'006c17f0    ff15e4104000            call dword ptr [004010e4]
'006c17f6    8d8558ffffff            lea eax, dword ptr [ebp+ffffff58]
'006c17fc    50                      push eax
'006c17fd    8d4d98                  lea ecx, dword ptr [ebp-68]
'006c1800    51                      push ecx
'006c1801    8d9568ffffff            lea edx, dword ptr [ebp+ffffff68]
'006c1807    52                      push edx
'006c1808    8d4588                  lea eax, dword ptr [ebp-78]
'006c180b    50                      push eax
'006c180c    c78570ffffffcc134300    mov dword ptr [ebp+ffffff70], 004313cc
'006c1816    c78568ffffff08800000    mov dword ptr [ebp+ffffff68], 00008008

' *** Reference to "__vbaVarCmpNe"
'006c1820    ff156c104000            call dword ptr [0040106c]
'006c1826    50                      push eax
'006c1827    8d8d78ffffff            lea ecx, dword ptr [ebp+ffffff78]
'006c182d    51                      push ecx

' *** Reference to "__vbaVarAnd"
'006c182e    ff1598114000            call dword ptr [00401198]
'006c1834    50                      push eax

' *** Reference to "__vbaBoolVarNull"
'006c1835    ff15fc104000            call dword ptr [004010fc]
'006c183b    8d4de8                  lea ecx, dword ptr [ebp-18]
'006c183e    6689854cffffff          mov word ptr [ebp+ffffff4c], ax

' *** Reference to "__vbaFreeStr"
'006c1845    ff150c134000            call dword ptr [0040130c]
'#Cleanup(var_41)
'006c184b    8d4dbc                  lea ecx, dword ptr [ebp-44]

' *** Reference to "__vbaFreeObj"
'006c184e    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_58)
'006c1854    8d9558ffffff            lea edx, dword ptr [ebp+ffffff58]
'006c185a    52                      push edx
'006c185b    8d4598                  lea eax, dword ptr [ebp-68]
'006c185e    50                      push eax
'006c185f    8d4da8                  lea ecx, dword ptr [ebp-58]
'006c1862    51                      push ecx
'006c1863    6a03                    push 03

' *** Reference to "__vbaFreeVarList"
'006c1865    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_86, var_130, var_133)
'006c186b    83c410                  add esp, 10
'006c186e    6683bd4cffffff00        cmp word ptr [ebp+ffffff4c], 00
'006c1876    0f84650a0000            je 6c22e1

If (CBool(#NOT SUPPORTED#) <> 0) Then
'006c187c    8b17                    mov edx, dword ptr [edi]
'006c187e    57                      push edi
'006c187f    ff920c030000            call dword ptr [edx+0000030c]
'006c1885    8945b0                  mov dword ptr [ebp-50], eax
'006c1888    8d45a8                  lea eax, dword ptr [ebp-58]
'006c188b    50                      push eax
'006c188c    8d4d98                  lea ecx, dword ptr [ebp-68]
'006c188f    51                      push ecx
'006c1890    c745a809000000          mov dword ptr [ebp-58], 00000009

' *** Reference to "rtcTrimVar"
'006c1897    ff15e4104000            call dword ptr [004010e4]
'006c189d    8d5598                  lea edx, dword ptr [ebp-68]
'006c18a0    52                      push edx
'006c18a1    8d8568ffffff            lea eax, dword ptr [ebp+ffffff68]
'006c18a7    50                      push eax
'006c18a8    c78570ffffffcc134300    mov dword ptr [ebp+ffffff70], 004313cc
'006c18b2    c78568ffffff08800000    mov dword ptr [ebp+ffffff68], 00008008

' *** Reference to "__vbaVarTstNe"
'006c18bc    ff1574124000            call dword ptr [00401274]
'006c18c2    8d4d98                  lea ecx, dword ptr [ebp-68]
'006c18c5    51                      push ecx
'006c18c6    8d55a8                  lea edx, dword ptr [ebp-58]
'006c18c9    52                      push edx
'006c18ca    6a02                    push 02
'006c18cc    66898554ffffff          mov word ptr [ebp+ffffff54], ax

' *** Reference to "__vbaFreeVarList"
'006c18d3    ff1540104000            call dword ptr [00401040]
    '#Cleanup( var_86, var_130)
'006c18d9    83c40c                  add esp, 0c
'006c18dc    6683bd54ffffff00        cmp word ptr [ebp+ffffff54], 00
'006c18e4    0f84bf010000            je 6c1aa9
    
    If (    ((Trim(-360)) <> (vbNullChar))) Then
'006c18ea    8b07                    mov eax, dword ptr [edi]
'006c18ec    57                      push edi
'006c18ed    ff900c030000            call dword ptr [eax+0000030c]
'006c18f3    8d4da8                  lea ecx, dword ptr [ebp-58]
'006c18f6    51                      push ecx
'006c18f7    8d5598                  lea edx, dword ptr [ebp-68]
'006c18fa    52                      push edx
'006c18fb    8945b0                  mov dword ptr [ebp-50], eax
'006c18fe    c745a809000000          mov dword ptr [ebp-58], 00000009

' *** Reference to "rtcTrimVar"
'006c1905    ff15e4104000            call dword ptr [004010e4]
'006c190b    8d4598                  lea eax, dword ptr [ebp-68]
'006c190e    50                      push eax

' *** Reference to "__vbaStrVarMove"
'006c190f    ff153c104000            call dword ptr [0040103c]
'006c1915    8bd0                    mov edx, eax
'006c1917    8d4de8                  lea ecx, dword ptr [ebp-18]
'006c191a    ffd6                    call esi
'006c191c    8d4de8                  lea ecx, dword ptr [ebp-18]
'006c191f    51                      push ecx
'006c1920    e8cb22e3ff              call 4f3bf0
    Call sub_4F3BF0()
'006c1925    8bd0                    mov edx, eax
'006c1927    8d4dc4                  lea ecx, dword ptr [ebp-3c]
'006c192a    ffd6                    call esi
'006c192c    8b17                    mov edx, dword ptr [edi]
'006c192e    57                      push edi
'006c192f    ff923c030000            call dword ptr [edx+0000033c]
'006c1935    50                      push eax
'006c1936    8d45bc                  lea eax, dword ptr [ebp-44]
'006c1939    50                      push eax
'006c193a    ffd3                    call ebx
    Set var_58 = -4552
'006c193c    8d55dc                  lea edx, dword ptr [ebp-24]
'006c193f    8bd8                    mov ebx, eax
'006c1941    8b0b                    mov ecx, dword ptr [ebx]
'006c1943    52                      push edx
'006c1944    53                      push ebx
'006c1945    ff5150                  call dword ptr [ecx+50]
'006c1948    dbe2                    fnclex
'006c194a    85c0                    test ax, ax
'006c194c    7d0f                    jge 6c195d
    
    If (    var_58) Then
'006c194e    6a50                    push 50
'006c1950    683ce04300              push 0043e03c
'006c1955    53                      push ebx
'006c1956    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c1957    ff1580104000            call dword ptr [00401080]
    
End If
'006c195d    8b55dc                  mov edx, dword ptr [ebp-24]
'006c1960    8d4dd8                  lea ecx, dword ptr [ebp-28]
'006c1963    c745dc00000000          mov dword ptr [ebp-24], 00000000
'006c196a    ffd6                    call esi
'006c196c    8d45d8                  lea eax, dword ptr [ebp-28]
'006c196f    50                      push eax
'006c1970    e87b22e3ff              call 4f3bf0
Call sub_4F3BF0()
'006c1975    8bd0                    mov edx, eax
'006c1977    8d4dc0                  lea ecx, dword ptr [ebp-40]
'006c197a    ffd6                    call esi
'006c197c    8b55c4                  mov edx, dword ptr [ebp-3c]
'006c197f    8b5dc0                  mov ebx, dword ptr [ebp-40]
'006c1982    899520ffffff            mov dword ptr [ebp+ffffff20], edx
'006c1988    33d2                    xor edx, edx
var_num4 = Empty
'006c198a    8955c4                  mov dword ptr [ebp-3c], edx
'006c198d    8955c0                  mov dword ptr [ebp-40], edx
'006c1990    8b154cb07200            mov edx, dword ptr [0072b04c]
'006c1996    b880000000              mov eax, 00000080
'006c199b    b903000000              mov ecx, 00000003
'006c19a0    898d68ffffff            mov dword ptr [ebp+ffffff68], ecx
'006c19a6    898570ffffff            mov dword ptr [ebp+ffffff70], eax
'006c19ac    83ec10                  sub esp, 10
'006c19af    899d1cffffff            mov dword ptr [ebp+ffffff1c], ebx
'006c19b5    8b1a                    mov ebx, dword ptr [edx]
'006c19b7    8bd4                    mov edx, esp
'006c19b9    890a                    mov dword ptr [edx], ecx
'006c19bb    8b8d6cffffff            mov ecx, dword ptr [ebp+ffffff6c]
'006c19c1    894a04                  mov dword ptr [edx+04], ecx
'006c19c4    894208                  mov dword ptr [edx+08], eax
'006c19c7    8b8574ffffff            mov eax, dword ptr [ebp+ffffff74]
'006c19cd    89420c                  mov dword ptr [edx+0c], eax
'006c19d0    8b9520ffffff            mov edx, dword ptr [ebp+ffffff20]
'006c19d6    684c364500              push 0045364c
'006c19db    8d4de4                  lea ecx, dword ptr [ebp-1c]
'006c19de    ffd6                    call esi
'006c19e0    50                      push eax

' *** Reference to "__vbaStrCat"
'006c19e1    ff1570104000            call dword ptr [00401070]
var_285 = ("update dons set Résumé='") & (Trim(frmDescriptifDon))
'006c19e7    8bd0                    mov edx, eax
'006c19e9    8d4de0                  lea ecx, dword ptr [ebp-20]
'006c19ec    ffd6                    call esi
'006c19ee    50                      push eax
'006c19ef    68a4fb4400              push 0044fba4

' *** Reference to "__vbaStrCat"
'006c19f4    ff1570104000            call dword ptr [00401070]
var_54 = (var_285) & ("' where nom='")
'006c19fa    8bd0                    mov edx, eax
'006c19fc    8d4dd4                  lea ecx, dword ptr [ebp-2c]
'006c19ff    ffd6                    call esi
'006c1a01    8b951cffffff            mov edx, dword ptr [ebp+ffffff1c]
'006c1a07    50                      push eax
'006c1a08    8d4dd0                  lea ecx, dword ptr [ebp-30]
'006c1a0b    ffd6                    call esi
'006c1a0d    50                      push eax

' *** Reference to "__vbaStrCat"
'006c1a0e    ff1570104000            call dword ptr [00401070]
var_45 = (var_54) & (0)
'006c1a14    8bd0                    mov edx, eax
'006c1a16    8d4dcc                  lea ecx, dword ptr [ebp-34]
'006c1a19    ffd6                    call esi
'006c1a1b    50                      push eax
'006c1a1c    6854a44300              push 0043a454

' *** Reference to "__vbaStrCat"
'006c1a21    ff1570104000            call dword ptr [00401070]
var_139 = (var_45) & ("'")
'006c1a27    8bd0                    mov edx, eax
'006c1a29    8d4dc8                  lea ecx, dword ptr [ebp-38]
'006c1a2c    ffd6                    call esi
'006c1a2e    8b0d4cb07200            mov ecx, dword ptr [0072b04c]
'006c1a34    50                      push eax
'006c1a35    51                      push ecx
'006c1a36    ff535c                  call dword ptr [ebx+5c]
'006c1a39    dbe2                    fnclex
'006c1a3b    85c0                    test ax, ax
'006c1a3d    7d15                    jge 6c1a54
'006c1a3f    8b154cb07200            mov edx, dword ptr [0072b04c]
'006c1a45    6a5c                    push 5c
'006c1a47    68301f4300              push 00431f30
'006c1a4c    52                      push edx
'006c1a4d    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c1a4e    ff1580104000            call dword ptr [00401080]
'006c1a54    8d45c0                  lea eax, dword ptr [ebp-40]
'006c1a57    50                      push eax
'006c1a58    8d4dc4                  lea ecx, dword ptr [ebp-3c]
'006c1a5b    51                      push ecx
'006c1a5c    8d55c8                  lea edx, dword ptr [ebp-38]
'006c1a5f    52                      push edx
'006c1a60    8d45cc                  lea eax, dword ptr [ebp-34]
'006c1a63    50                      push eax
'006c1a64    8d4dd0                  lea ecx, dword ptr [ebp-30]
'006c1a67    51                      push ecx
'006c1a68    8d55d4                  lea edx, dword ptr [ebp-2c]
'006c1a6b    52                      push edx
'006c1a6c    8d45d8                  lea eax, dword ptr [ebp-28]
'006c1a6f    50                      push eax
'006c1a70    8d4de0                  lea ecx, dword ptr [ebp-20]
'006c1a73    51                      push ecx
'006c1a74    8d55e4                  lea edx, dword ptr [ebp-1c]
'006c1a77    52                      push edx
'006c1a78    8d45e8                  lea eax, dword ptr [ebp-18]
'006c1a7b    50                      push eax
'006c1a7c    6a0a                    push 0a

' *** Reference to "__vbaFreeStrList"
'006c1a7e    ff155c124000            call dword ptr [0040125c]
'#Cleanup( -4552, -4552, -4556, 0, -4560, -296, -4564, -4568, -4552, -4552)
'006c1a84    83c42c                  add esp, 2c
'006c1a87    8d4dbc                  lea ecx, dword ptr [ebp-44]

' *** Reference to "__vbaFreeObj"
'006c1a8a    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_58)
'006c1a90    8d4d98                  lea ecx, dword ptr [ebp-68]
'006c1a93    51                      push ecx
'006c1a94    8d55a8                  lea edx, dword ptr [ebp-58]
'006c1a97    52                      push edx
'006c1a98    6a02                    push 02

' *** Reference to "__vbaFreeVarList"
'006c1a9a    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_86, var_130)

' *** Reference to "__vbaObjSet"
'006c1aa0    8b1db4104000            mov ebx, dword ptr [004010b4]
'006c1aa6    83c40c                  add esp, 0c

'ERROR: Two many next close:
End If
'006c1aa9    8b07                    mov eax, dword ptr [edi]
'006c1aab    57                      push edi
'006c1aac    ff9018030000            call dword ptr [eax+00000318]
'006c1ab2    8d4da8                  lea ecx, dword ptr [ebp-58]
'006c1ab5    51                      push ecx
'006c1ab6    8d5598                  lea edx, dword ptr [ebp-68]
'006c1ab9    52                      push edx
'006c1aba    8945b0                  mov dword ptr [ebp-50], eax
'006c1abd    c745a809000000          mov dword ptr [ebp-58], 00000009

' *** Reference to "rtcTrimVar"
'006c1ac4    ff15e4104000            call dword ptr [004010e4]
'006c1aca    8d4598                  lea eax, dword ptr [ebp-68]
'006c1acd    50                      push eax
'006c1ace    8d8d68ffffff            lea ecx, dword ptr [ebp+ffffff68]
'006c1ad4    51                      push ecx
'006c1ad5    c78570ffffffcc134300    mov dword ptr [ebp+ffffff70], 004313cc
'006c1adf    c78568ffffff08800000    mov dword ptr [ebp+ffffff68], 00008008

' *** Reference to "__vbaVarTstNe"
'006c1ae9    ff1574124000            call dword ptr [00401274]
'006c1aef    8d5598                  lea edx, dword ptr [ebp-68]
'006c1af2    66898554ffffff          mov word ptr [ebp+ffffff54], ax
'006c1af9    52                      push edx
'006c1afa    8d45a8                  lea eax, dword ptr [ebp-58]
'006c1afd    50                      push eax
'006c1afe    6a02                    push 02

' *** Reference to "__vbaFreeVarList"
'006c1b00    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_86, var_130)
'006c1b06    83c40c                  add esp, 0c
'006c1b09    6683bd54ffffff00        cmp word ptr [ebp+ffffff54], 00
'006c1b11    0f84bf010000            je 6c1cd6

If (((Trim(frmDescriptifDon)) <> (vbNullChar))) Then
'006c1b17    8b0f                    mov ecx, dword ptr [edi]
'006c1b19    57                      push edi
'006c1b1a    ff9118030000            call dword ptr [ecx+00000318]
'006c1b20    8d55a8                  lea edx, dword ptr [ebp-58]
'006c1b23    8945b0                  mov dword ptr [ebp-50], eax
'006c1b26    52                      push edx
'006c1b27    8d4598                  lea eax, dword ptr [ebp-68]
'006c1b2a    50                      push eax
'006c1b2b    c745a809000000          mov dword ptr [ebp-58], 00000009

' *** Reference to "rtcTrimVar"
'006c1b32    ff15e4104000            call dword ptr [004010e4]
'006c1b38    8d4d98                  lea ecx, dword ptr [ebp-68]
'006c1b3b    51                      push ecx

' *** Reference to "__vbaStrVarMove"
'006c1b3c    ff153c104000            call dword ptr [0040103c]
'006c1b42    8bd0                    mov edx, eax
'006c1b44    8d4de8                  lea ecx, dword ptr [ebp-18]
'006c1b47    ffd6                    call esi
'006c1b49    8d55e8                  lea edx, dword ptr [ebp-18]
'006c1b4c    52                      push edx
'006c1b4d    e89e20e3ff              call 4f3bf0
    Call sub_4F3BF0()
'006c1b52    8bd0                    mov edx, eax
'006c1b54    8d4dc4                  lea ecx, dword ptr [ebp-3c]
'006c1b57    ffd6                    call esi
'006c1b59    8b07                    mov eax, dword ptr [edi]
'006c1b5b    57                      push edi
'006c1b5c    ff903c030000            call dword ptr [eax+0000033c]
'006c1b62    50                      push eax
'006c1b63    8d4dbc                  lea ecx, dword ptr [ebp-44]
'006c1b66    51                      push ecx
'006c1b67    ffd3                    call ebx
    Set var_58 = Nothing
'006c1b69    8bd8                    mov ebx, eax
'006c1b6b    8b13                    mov edx, dword ptr [ebx]
'006c1b6d    8d45dc                  lea eax, dword ptr [ebp-24]
'006c1b70    50                      push eax
'006c1b71    53                      push ebx
'006c1b72    ff5250                  call dword ptr [edx+50]
'006c1b75    dbe2                    fnclex
'006c1b77    85c0                    test ax, ax
'006c1b79    7d0f                    jge 6c1b8a
'006c1b7b    6a50                    push 50
'006c1b7d    683ce04300              push 0043e03c
'006c1b82    53                      push ebx
'006c1b83    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c1b84    ff1580104000            call dword ptr [00401080]
'006c1b8a    8b55dc                  mov edx, dword ptr [ebp-24]
'006c1b8d    8d4dd8                  lea ecx, dword ptr [ebp-28]
'006c1b90    c745dc00000000          mov dword ptr [ebp-24], 00000000
'006c1b97    ffd6                    call esi
'006c1b99    8d4dd8                  lea ecx, dword ptr [ebp-28]
'006c1b9c    51                      push ecx
'006c1b9d    e84e20e3ff              call 4f3bf0
    Call sub_4F3BF0()
'006c1ba2    8bd0                    mov edx, eax
'006c1ba4    8d4dc0                  lea ecx, dword ptr [ebp-40]
'006c1ba7    ffd6                    call esi
'006c1ba9    8b55c4                  mov edx, dword ptr [ebp-3c]
'006c1bac    8b5dc0                  mov ebx, dword ptr [ebp-40]
'006c1baf    899514ffffff            mov dword ptr [ebp+ffffff14], edx
'006c1bb5    33d2                    xor edx, edx
    var_num4 = Empty
'006c1bb7    8955c4                  mov dword ptr [ebp-3c], edx
'006c1bba    8955c0                  mov dword ptr [ebp-40], edx
'006c1bbd    8b154cb07200            mov edx, dword ptr [0072b04c]
'006c1bc3    b880000000              mov eax, 00000080
'006c1bc8    b903000000              mov ecx, 00000003
'006c1bcd    898d68ffffff            mov dword ptr [ebp+ffffff68], ecx
'006c1bd3    898570ffffff            mov dword ptr [ebp+ffffff70], eax
'006c1bd9    83ec10                  sub esp, 10
'006c1bdc    899d10ffffff            mov dword ptr [ebp+ffffff10], ebx
'006c1be2    8b1a                    mov ebx, dword ptr [edx]
'006c1be4    8bd4                    mov edx, esp
'006c1be6    890a                    mov dword ptr [edx], ecx
'006c1be8    8b8d6cffffff            mov ecx, dword ptr [ebp+ffffff6c]
'006c1bee    894a04                  mov dword ptr [edx+04], ecx
'006c1bf1    894208                  mov dword ptr [edx+08], eax
'006c1bf4    8b8574ffffff            mov eax, dword ptr [ebp+ffffff74]
'006c1bfa    89420c                  mov dword ptr [edx+0c], eax
'006c1bfd    8b9514ffffff            mov edx, dword ptr [ebp+ffffff14]
'006c1c03    6884364500              push 00453684
'006c1c08    8d4de4                  lea ecx, dword ptr [ebp-1c]
'006c1c0b    ffd6                    call esi
'006c1c0d    50                      push eax

' *** Reference to "__vbaStrCat"
'006c1c0e    ff1570104000            call dword ptr [00401070]
    var_2068 = ("update dons set Condition='") & (Trim(-344))
'006c1c14    8bd0                    mov edx, eax
'006c1c16    8d4de0                  lea ecx, dword ptr [ebp-20]
'006c1c19    ffd6                    call esi
'006c1c1b    50                      push eax
'006c1c1c    68a4fb4400              push 0044fba4

' *** Reference to "__vbaStrCat"
'006c1c21    ff1570104000            call dword ptr [00401070]
    var_2091 = (var_2068) & ("' where nom='")
'006c1c27    8bd0                    mov edx, eax
'006c1c29    8d4dd4                  lea ecx, dword ptr [ebp-2c]
'006c1c2c    ffd6                    call esi
'006c1c2e    8b9510ffffff            mov edx, dword ptr [ebp+ffffff10]
'006c1c34    50                      push eax
'006c1c35    8d4dd0                  lea ecx, dword ptr [ebp-30]
'006c1c38    ffd6                    call esi
'006c1c3a    50                      push eax

' *** Reference to "__vbaStrCat"
'006c1c3b    ff1570104000            call dword ptr [00401070]
    var_49 = (var_2091) & (vbNullString)
'006c1c41    8bd0                    mov edx, eax
'006c1c43    8d4dcc                  lea ecx, dword ptr [ebp-34]
'006c1c46    ffd6                    call esi
'006c1c48    50                      push eax
'006c1c49    6854a44300              push 0043a454

' *** Reference to "__vbaStrCat"
'006c1c4e    ff1570104000            call dword ptr [00401070]
    var_2092 = (var_49) & ("'")
'006c1c54    8bd0                    mov edx, eax
'006c1c56    8d4dc8                  lea ecx, dword ptr [ebp-38]
'006c1c59    ffd6                    call esi
'006c1c5b    8b0d4cb07200            mov ecx, dword ptr [0072b04c]
'006c1c61    50                      push eax
'006c1c62    51                      push ecx
'006c1c63    ff535c                  call dword ptr [ebx+5c]
'006c1c66    dbe2                    fnclex
'006c1c68    85c0                    test ax, ax
'006c1c6a    7d15                    jge 6c1c81
'006c1c6c    8b154cb07200            mov edx, dword ptr [0072b04c]
'006c1c72    6a5c                    push 5c
'006c1c74    68301f4300              push 00431f30
'006c1c79    52                      push edx
'006c1c7a    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c1c7b    ff1580104000            call dword ptr [00401080]
'006c1c81    8d45c0                  lea eax, dword ptr [ebp-40]
'006c1c84    50                      push eax
'006c1c85    8d4dc4                  lea ecx, dword ptr [ebp-3c]
'006c1c88    51                      push ecx
'006c1c89    8d55c8                  lea edx, dword ptr [ebp-38]
'006c1c8c    52                      push edx
'006c1c8d    8d45cc                  lea eax, dword ptr [ebp-34]
'006c1c90    50                      push eax
'006c1c91    8d4dd0                  lea ecx, dword ptr [ebp-30]
'006c1c94    51                      push ecx
'006c1c95    8d55d4                  lea edx, dword ptr [ebp-2c]
'006c1c98    52                      push edx
'006c1c99    8d45d8                  lea eax, dword ptr [ebp-28]
'006c1c9c    50                      push eax
'006c1c9d    8d4de0                  lea ecx, dword ptr [ebp-20]
'006c1ca0    51                      push ecx
'006c1ca1    8d55e4                  lea edx, dword ptr [ebp-1c]
'006c1ca4    52                      push edx
'006c1ca5    8d45e8                  lea eax, dword ptr [ebp-18]
'006c1ca8    50                      push eax
'006c1ca9    6a0a                    push 0a

' *** Reference to "__vbaFreeStrList"
'006c1cab    ff155c124000            call dword ptr [0040125c]
    '#Cleanup( -4584, -4584, -4588, var_39, -4592, 0, -4596, -4600, -4584, -4584)
'006c1cb1    83c42c                  add esp, 2c
'006c1cb4    8d4dbc                  lea ecx, dword ptr [ebp-44]

' *** Reference to "__vbaFreeObj"
'006c1cb7    ff1508134000            call dword ptr [00401308]
    '#Cleanup(var_58)
'006c1cbd    8d4d98                  lea ecx, dword ptr [ebp-68]
'006c1cc0    51                      push ecx
'006c1cc1    8d55a8                  lea edx, dword ptr [ebp-58]
'006c1cc4    52                      push edx
'006c1cc5    6a02                    push 02

' *** Reference to "__vbaFreeVarList"
'006c1cc7    ff1540104000            call dword ptr [00401040]
    '#Cleanup( var_86, var_130)

' *** Reference to "__vbaObjSet"
'006c1ccd    8b1db4104000            mov ebx, dword ptr [004010b4]
'006c1cd3    83c40c                  add esp, 0c
    
End If
'006c1cd6    8b07                    mov eax, dword ptr [edi]
'006c1cd8    57                      push edi
'006c1cd9    ff901c030000            call dword ptr [eax+0000031c]
'006c1cdf    8d4da8                  lea ecx, dword ptr [ebp-58]
'006c1ce2    51                      push ecx
'006c1ce3    8d5598                  lea edx, dword ptr [ebp-68]
'006c1ce6    52                      push edx
'006c1ce7    8945b0                  mov dword ptr [ebp-50], eax
'006c1cea    c745a809000000          mov dword ptr [ebp-58], 00000009

' *** Reference to "rtcTrimVar"
'006c1cf1    ff15e4104000            call dword ptr [004010e4]
'006c1cf7    8d4598                  lea eax, dword ptr [ebp-68]
'006c1cfa    50                      push eax

' *** Reference to "__vbaStrVarMove"
'006c1cfb    ff153c104000            call dword ptr [0040103c]
'006c1d01    8bd0                    mov edx, eax
'006c1d03    8d4de8                  lea ecx, dword ptr [ebp-18]
'006c1d06    ffd6                    call esi
'006c1d08    8d4de8                  lea ecx, dword ptr [ebp-18]
'006c1d0b    51                      push ecx
'006c1d0c    e8df1ee3ff              call 4f3bf0
Call sub_4F3BF0()
'006c1d11    8bd0                    mov edx, eax
'006c1d13    8d4dc4                  lea ecx, dword ptr [ebp-3c]
'006c1d16    ffd6                    call esi
'006c1d18    8b17                    mov edx, dword ptr [edi]
'006c1d1a    57                      push edi
'006c1d1b    ff923c030000            call dword ptr [edx+0000033c]
'006c1d21    50                      push eax
'006c1d22    8d45bc                  lea eax, dword ptr [ebp-44]
'006c1d25    50                      push eax
'006c1d26    ffd3                    call ebx
Set var_58 = -4608
'006c1d28    8d55dc                  lea edx, dword ptr [ebp-24]
'006c1d2b    8bd8                    mov ebx, eax
'006c1d2d    8b0b                    mov ecx, dword ptr [ebx]
'006c1d2f    52                      push edx
'006c1d30    53                      push ebx
'006c1d31    ff5150                  call dword ptr [ecx+50]
'006c1d34    dbe2                    fnclex
'006c1d36    85c0                    test ax, ax
'006c1d38    7d0f                    jge 6c1d49

If (var_58) Then
'006c1d3a    6a50                    push 50
'006c1d3c    683ce04300              push 0043e03c
'006c1d41    53                      push ebx
'006c1d42    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c1d43    ff1580104000            call dword ptr [00401080]
    
End If
'006c1d49    8b55dc                  mov edx, dword ptr [ebp-24]
'006c1d4c    8d4dd8                  lea ecx, dword ptr [ebp-28]
'006c1d4f    c745dc00000000          mov dword ptr [ebp-24], 00000000
'006c1d56    ffd6                    call esi
'006c1d58    8d45d8                  lea eax, dword ptr [ebp-28]
'006c1d5b    50                      push eax
'006c1d5c    e88f1ee3ff              call 4f3bf0
Call sub_4F3BF0()
'006c1d61    8bd0                    mov edx, eax
'006c1d63    8d4dc0                  lea ecx, dword ptr [ebp-40]
'006c1d66    ffd6                    call esi
'006c1d68    8b55c4                  mov edx, dword ptr [ebp-3c]
'006c1d6b    8b5dc0                  mov ebx, dword ptr [ebp-40]
'006c1d6e    899508ffffff            mov dword ptr [ebp+ffffff08], edx
'006c1d74    33d2                    xor edx, edx
var_num4 = Empty
'006c1d76    8955c4                  mov dword ptr [ebp-3c], edx
'006c1d79    8955c0                  mov dword ptr [ebp-40], edx
'006c1d7c    8b154cb07200            mov edx, dword ptr [0072b04c]
'006c1d82    b880000000              mov eax, 00000080
'006c1d87    b903000000              mov ecx, 00000003
'006c1d8c    898d68ffffff            mov dword ptr [ebp+ffffff68], ecx
'006c1d92    898570ffffff            mov dword ptr [ebp+ffffff70], eax
'006c1d98    83ec10                  sub esp, 10
'006c1d9b    899d04ffffff            mov dword ptr [ebp+ffffff04], ebx
'006c1da1    8b1a                    mov ebx, dword ptr [edx]
'006c1da3    8bd4                    mov edx, esp
'006c1da5    890a                    mov dword ptr [edx], ecx
'006c1da7    8b8d6cffffff            mov ecx, dword ptr [ebp+ffffff6c]
'006c1dad    894a04                  mov dword ptr [edx+04], ecx
'006c1db0    894208                  mov dword ptr [edx+08], eax
'006c1db3    8b8574ffffff            mov eax, dword ptr [ebp+ffffff74]
'006c1db9    89420c                  mov dword ptr [edx+0c], eax
'006c1dbc    8b9508ffffff            mov edx, dword ptr [ebp+ffffff08]
'006c1dc2    68c0364500              push 004536c0
'006c1dc7    8d4de4                  lea ecx, dword ptr [ebp-1c]
'006c1dca    ffd6                    call esi
'006c1dcc    50                      push eax

' *** Reference to "__vbaStrCat"
'006c1dcd    ff1570104000            call dword ptr [00401070]
var_49 = ("update dons set explication='") & (vbNullString)
'006c1dd3    8bd0                    mov edx, eax
'006c1dd5    8d4de0                  lea ecx, dword ptr [ebp-20]
'006c1dd8    ffd6                    call esi
'006c1dda    50                      push eax
'006c1ddb    68a4fb4400              push 0044fba4

' *** Reference to "__vbaStrCat"
'006c1de0    ff1570104000            call dword ptr [00401070]
var_64 = (var_49) & ("' where nom='")
'006c1de6    8bd0                    mov edx, eax
'006c1de8    8d4dd4                  lea ecx, dword ptr [ebp-2c]
'006c1deb    ffd6                    call esi
'006c1ded    8b9504ffffff            mov edx, dword ptr [ebp+ffffff04]
'006c1df3    50                      push eax
'006c1df4    8d4dd0                  lea ecx, dword ptr [ebp-30]
'006c1df7    ffd6                    call esi
'006c1df9    50                      push eax

' *** Reference to "__vbaStrCat"
'006c1dfa    ff1570104000            call dword ptr [00401070]
var_97 = (var_64) & ()
'006c1e00    8bd0                    mov edx, eax
'006c1e02    8d4dcc                  lea ecx, dword ptr [ebp-34]
'006c1e05    ffd6                    call esi
'006c1e07    50                      push eax
'006c1e08    6854a44300              push 0043a454

' *** Reference to "__vbaStrCat"
'006c1e0d    ff1570104000            call dword ptr [00401070]
var_871 = (var_97) & ("'")
'006c1e13    8bd0                    mov edx, eax
'006c1e15    8d4dc8                  lea ecx, dword ptr [ebp-38]
'006c1e18    ffd6                    call esi
'006c1e1a    8b0d4cb07200            mov ecx, dword ptr [0072b04c]
'006c1e20    50                      push eax
'006c1e21    51                      push ecx
'006c1e22    ff535c                  call dword ptr [ebx+5c]
'006c1e25    dbe2                    fnclex
'006c1e27    85c0                    test ax, ax
'006c1e29    7d15                    jge 6c1e40
'006c1e2b    8b154cb07200            mov edx, dword ptr [0072b04c]
'006c1e31    6a5c                    push 5c
'006c1e33    68301f4300              push 00431f30
'006c1e38    52                      push edx
'006c1e39    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c1e3a    ff1580104000            call dword ptr [00401080]
'006c1e40    8d45c0                  lea eax, dword ptr [ebp-40]
'006c1e43    50                      push eax
'006c1e44    8d4dc4                  lea ecx, dword ptr [ebp-3c]
'006c1e47    51                      push ecx
'006c1e48    8d55c8                  lea edx, dword ptr [ebp-38]
'006c1e4b    52                      push edx
'006c1e4c    8d45cc                  lea eax, dword ptr [ebp-34]
'006c1e4f    50                      push eax
'006c1e50    8d4dd0                  lea ecx, dword ptr [ebp-30]
'006c1e53    51                      push ecx
'006c1e54    8d55d4                  lea edx, dword ptr [ebp-2c]
'006c1e57    52                      push edx
'006c1e58    8d45d8                  lea eax, dword ptr [ebp-28]
'006c1e5b    50                      push eax
'006c1e5c    8d4de0                  lea ecx, dword ptr [ebp-20]
'006c1e5f    51                      push ecx
'006c1e60    8d55e4                  lea edx, dword ptr [ebp-1c]
'006c1e63    52                      push edx
'006c1e64    8d45e8                  lea eax, dword ptr [ebp-18]
'006c1e67    50                      push eax
'006c1e68    6a0a                    push 0a

' *** Reference to "__vbaFreeStrList"
'006c1e6a    ff155c124000            call dword ptr [0040125c]
'#Cleanup( -4608, 0, -4612, var_39, -4616, 3, -4620, -4624, -4608, -4608)
'006c1e70    83c42c                  add esp, 2c
'006c1e73    8d4dbc                  lea ecx, dword ptr [ebp-44]

' *** Reference to "__vbaFreeObj"
'006c1e76    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_58)

' *** Reference to "__vbaFreeVarList"
'006c1e7c    8b1d40104000            mov ebx, dword ptr [00401040]
'006c1e82    8d4d98                  lea ecx, dword ptr [ebp-68]
'006c1e85    51                      push ecx
'006c1e86    8d55a8                  lea edx, dword ptr [ebp-58]
'006c1e89    52                      push edx
'006c1e8a    6a02                    push 02
'006c1e8c    ffd3                    call ebx
'#Cleanup( var_86, var_130)
'006c1e8e    8b07                    mov eax, dword ptr [edi]
'006c1e90    83c40c                  add esp, 0c
'006c1e93    57                      push edi
'006c1e94    ff9014030000            call dword ptr [eax+00000314]
'006c1e9a    8d4da8                  lea ecx, dword ptr [ebp-58]
'006c1e9d    51                      push ecx
'006c1e9e    8d5598                  lea edx, dword ptr [ebp-68]
'006c1ea1    52                      push edx
'006c1ea2    8945b0                  mov dword ptr [ebp-50], eax
'006c1ea5    c745a809000000          mov dword ptr [ebp-58], 00000009

' *** Reference to "rtcTrimVar"
'006c1eac    ff15e4104000            call dword ptr [004010e4]
'006c1eb2    8d4598                  lea eax, dword ptr [ebp-68]
'006c1eb5    50                      push eax
'006c1eb6    8d8d68ffffff            lea ecx, dword ptr [ebp+ffffff68]
'006c1ebc    51                      push ecx
'006c1ebd    c78570ffffffcc134300    mov dword ptr [ebp+ffffff70], 004313cc
'006c1ec7    c78568ffffff08800000    mov dword ptr [ebp+ffffff68], 00008008

' *** Reference to "__vbaVarTstNe"
'006c1ed1    ff1574124000            call dword ptr [00401274]
'006c1ed7    8d5598                  lea edx, dword ptr [ebp-68]
'006c1eda    66898554ffffff          mov word ptr [ebp+ffffff54], ax
'006c1ee1    52                      push edx
'006c1ee2    8d45a8                  lea eax, dword ptr [ebp-58]
'006c1ee5    50                      push eax
'006c1ee6    6a02                    push 02
'006c1ee8    ffd3                    call ebx
'#Cleanup( var_86, var_130)
'006c1eea    83c40c                  add esp, 0c
'006c1eed    6683bd54ffffff00        cmp word ptr [ebp+ffffff54], 00
'006c1ef5    0f84bf010000            je 6c20ba

If (((Trim(frmDescriptifDon)) <> (vbNullChar))) Then
'006c1efb    8b0f                    mov ecx, dword ptr [edi]
'006c1efd    57                      push edi
'006c1efe    ff9114030000            call dword ptr [ecx+00000314]
'006c1f04    8d55a8                  lea edx, dword ptr [ebp-58]
'006c1f07    8945b0                  mov dword ptr [ebp-50], eax
'006c1f0a    52                      push edx
'006c1f0b    8d4598                  lea eax, dword ptr [ebp-68]
'006c1f0e    50                      push eax
'006c1f0f    c745a809000000          mov dword ptr [ebp-58], 00000009

' *** Reference to "rtcTrimVar"
'006c1f16    ff15e4104000            call dword ptr [004010e4]
'006c1f1c    8d4d98                  lea ecx, dword ptr [ebp-68]
'006c1f1f    51                      push ecx

' *** Reference to "__vbaStrVarMove"
'006c1f20    ff153c104000            call dword ptr [0040103c]
'006c1f26    8bd0                    mov edx, eax
'006c1f28    8d4de8                  lea ecx, dword ptr [ebp-18]
'006c1f2b    ffd6                    call esi
'006c1f2d    8d55e8                  lea edx, dword ptr [ebp-18]
'006c1f30    52                      push edx
'006c1f31    e8ba1ce3ff              call 4f3bf0
    Call sub_4F3BF0()
'006c1f36    8bd0                    mov edx, eax
'006c1f38    8d4dc4                  lea ecx, dword ptr [ebp-3c]
'006c1f3b    ffd6                    call esi
'006c1f3d    8b07                    mov eax, dword ptr [edi]
'006c1f3f    57                      push edi
'006c1f40    ff903c030000            call dword ptr [eax+0000033c]
'006c1f46    50                      push eax
'006c1f47    8d4dbc                  lea ecx, dword ptr [ebp-44]
'006c1f4a    51                      push ecx

' *** Reference to "__vbaObjSet"
'006c1f4b    ff15b4104000            call dword ptr [004010b4]
    Set var_58 = Nothing
'006c1f51    8bd8                    mov ebx, eax
'006c1f53    8b13                    mov edx, dword ptr [ebx]
'006c1f55    8d45dc                  lea eax, dword ptr [ebp-24]
'006c1f58    50                      push eax
'006c1f59    53                      push ebx
'006c1f5a    ff5250                  call dword ptr [edx+50]
'006c1f5d    dbe2                    fnclex
'006c1f5f    85c0                    test ax, ax
'006c1f61    7d0f                    jge 6c1f72
'006c1f63    6a50                    push 50
'006c1f65    683ce04300              push 0043e03c
'006c1f6a    53                      push ebx
'006c1f6b    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c1f6c    ff1580104000            call dword ptr [00401080]
'006c1f72    8b55dc                  mov edx, dword ptr [ebp-24]
'006c1f75    8d4dd8                  lea ecx, dword ptr [ebp-28]
'006c1f78    c745dc00000000          mov dword ptr [ebp-24], 00000000
'006c1f7f    ffd6                    call esi
'006c1f81    8d4dd8                  lea ecx, dword ptr [ebp-28]
'006c1f84    51                      push ecx
'006c1f85    e8661ce3ff              call 4f3bf0
    Call sub_4F3BF0()
'006c1f8a    8bd0                    mov edx, eax
'006c1f8c    8d4dc0                  lea ecx, dword ptr [ebp-40]
'006c1f8f    ffd6                    call esi
'006c1f91    8b55c4                  mov edx, dword ptr [ebp-3c]
'006c1f94    8b5dc0                  mov ebx, dword ptr [ebp-40]
'006c1f97    8995fcfeffff            mov dword ptr [ebp+fffffefc], edx
'006c1f9d    33d2                    xor edx, edx
    var_num4 = Empty
'006c1f9f    8955c4                  mov dword ptr [ebp-3c], edx
'006c1fa2    8955c0                  mov dword ptr [ebp-40], edx
'006c1fa5    8b154cb07200            mov edx, dword ptr [0072b04c]
'006c1fab    b880000000              mov eax, 00000080
'006c1fb0    b903000000              mov ecx, 00000003
'006c1fb5    898d68ffffff            mov dword ptr [ebp+ffffff68], ecx
'006c1fbb    898570ffffff            mov dword ptr [ebp+ffffff70], eax
'006c1fc1    83ec10                  sub esp, 10
'006c1fc4    899df8feffff            mov dword ptr [ebp+fffffef8], ebx
'006c1fca    8b1a                    mov ebx, dword ptr [edx]
'006c1fcc    8bd4                    mov edx, esp
'006c1fce    890a                    mov dword ptr [edx], ecx
'006c1fd0    8b8d6cffffff            mov ecx, dword ptr [ebp+ffffff6c]
'006c1fd6    894a04                  mov dword ptr [edx+04], ecx
'006c1fd9    894208                  mov dword ptr [edx+08], eax
'006c1fdc    8b8574ffffff            mov eax, dword ptr [ebp+ffffff74]
'006c1fe2    89420c                  mov dword ptr [edx+0c], eax
'006c1fe5    8b95fcfeffff            mov edx, dword ptr [ebp+fffffefc]
'006c1feb    68c8f94400              push 0044f9c8
'006c1ff0    8d4de4                  lea ecx, dword ptr [ebp-1c]
'006c1ff3    ffd6                    call esi
'006c1ff5    50                      push eax

' *** Reference to "__vbaStrCat"
'006c1ff6    ff1570104000            call dword ptr [00401070]
    var_68 = ("update dons set NORMAL='") & (Trim(-344))
'006c1ffc    8bd0                    mov edx, eax
'006c1ffe    8d4de0                  lea ecx, dword ptr [ebp-20]
'006c2001    ffd6                    call esi
'006c2003    50                      push eax
'006c2004    68a4fb4400              push 0044fba4

' *** Reference to "__vbaStrCat"
'006c2009    ff1570104000            call dword ptr [00401070]
    var_2062 = (var_68) & ("' where nom='")
'006c200f    8bd0                    mov edx, eax
'006c2011    8d4dd4                  lea ecx, dword ptr [ebp-2c]
'006c2014    ffd6                    call esi
'006c2016    8b95f8feffff            mov edx, dword ptr [ebp+fffffef8]
'006c201c    50                      push eax
'006c201d    8d4dd0                  lea ecx, dword ptr [ebp-30]
'006c2020    ffd6                    call esi
'006c2022    50                      push eax

' *** Reference to "__vbaStrCat"
'006c2023    ff1570104000            call dword ptr [00401070]
    var_49 = (var_2062) & (vbNullString)
'006c2029    8bd0                    mov edx, eax
'006c202b    8d4dcc                  lea ecx, dword ptr [ebp-34]
'006c202e    ffd6                    call esi
'006c2030    50                      push eax
'006c2031    6854a44300              push 0043a454

' *** Reference to "__vbaStrCat"
'006c2036    ff1570104000            call dword ptr [00401070]
    var_70 = (var_49) & ("'")
'006c203c    8bd0                    mov edx, eax
'006c203e    8d4dc8                  lea ecx, dword ptr [ebp-38]
'006c2041    ffd6                    call esi
'006c2043    8b0d4cb07200            mov ecx, dword ptr [0072b04c]
'006c2049    50                      push eax
'006c204a    51                      push ecx
'006c204b    ff535c                  call dword ptr [ebx+5c]
'006c204e    dbe2                    fnclex
'006c2050    85c0                    test ax, ax
'006c2052    7d15                    jge 6c2069
'006c2054    8b154cb07200            mov edx, dword ptr [0072b04c]
'006c205a    6a5c                    push 5c
'006c205c    68301f4300              push 00431f30
'006c2061    52                      push edx
'006c2062    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c2063    ff1580104000            call dword ptr [00401080]
'006c2069    8d45c0                  lea eax, dword ptr [ebp-40]
'006c206c    50                      push eax
'006c206d    8d4dc4                  lea ecx, dword ptr [ebp-3c]
'006c2070    51                      push ecx
'006c2071    8d55c8                  lea edx, dword ptr [ebp-38]
'006c2074    52                      push edx
'006c2075    8d45cc                  lea eax, dword ptr [ebp-34]
'006c2078    50                      push eax
'006c2079    8d4dd0                  lea ecx, dword ptr [ebp-30]
'006c207c    51                      push ecx
'006c207d    8d55d4                  lea edx, dword ptr [ebp-2c]
'006c2080    52                      push edx
'006c2081    8d45d8                  lea eax, dword ptr [ebp-28]
'006c2084    50                      push eax
'006c2085    8d4de0                  lea ecx, dword ptr [ebp-20]
'006c2088    51                      push ecx
'006c2089    8d55e4                  lea edx, dword ptr [ebp-1c]
'006c208c    52                      push edx
'006c208d    8d45e8                  lea eax, dword ptr [ebp-18]
'006c2090    50                      push eax
'006c2091    6a0a                    push 0a

' *** Reference to "__vbaFreeStrList"
'006c2093    ff155c124000            call dword ptr [0040125c]
    '#Cleanup( -4640, -4640, -4644, var_39, -4648, var_39, -4652, -4656, -4640, -4640)
'006c2099    83c42c                  add esp, 2c
'006c209c    8d4dbc                  lea ecx, dword ptr [ebp-44]

' *** Reference to "__vbaFreeObj"
'006c209f    ff1508134000            call dword ptr [00401308]
    '#Cleanup(var_58)

' *** Reference to "__vbaFreeVarList"
'006c20a5    8b1d40104000            mov ebx, dword ptr [00401040]
'006c20ab    8d4d98                  lea ecx, dword ptr [ebp-68]
'006c20ae    51                      push ecx
'006c20af    8d55a8                  lea edx, dword ptr [ebp-58]
'006c20b2    52                      push edx
'006c20b3    6a02                    push 02
'006c20b5    ffd3                    call ebx
    '#Cleanup( var_86, var_130)
'006c20b7    83c40c                  add esp, 0c
    
End If
'006c20ba    8b07                    mov eax, dword ptr [edi]
'006c20bc    57                      push edi
'006c20bd    ff9010030000            call dword ptr [eax+00000310]
'006c20c3    8d4da8                  lea ecx, dword ptr [ebp-58]
'006c20c6    51                      push ecx
'006c20c7    8d5598                  lea edx, dword ptr [ebp-68]
'006c20ca    52                      push edx
'006c20cb    8945b0                  mov dword ptr [ebp-50], eax
'006c20ce    c745a809000000          mov dword ptr [ebp-58], 00000009

' *** Reference to "rtcTrimVar"
'006c20d5    ff15e4104000            call dword ptr [004010e4]
'006c20db    8d4598                  lea eax, dword ptr [ebp-68]
'006c20de    50                      push eax
'006c20df    8d8d68ffffff            lea ecx, dword ptr [ebp+ffffff68]
'006c20e5    51                      push ecx
'006c20e6    c78570ffffffcc134300    mov dword ptr [ebp+ffffff70], 004313cc
'006c20f0    c78568ffffff08800000    mov dword ptr [ebp+ffffff68], 00008008

' *** Reference to "__vbaVarTstNe"
'006c20fa    ff1574124000            call dword ptr [00401274]
'006c2100    8d5598                  lea edx, dword ptr [ebp-68]
'006c2103    66898554ffffff          mov word ptr [ebp+ffffff54], ax
'006c210a    52                      push edx
'006c210b    8d45a8                  lea eax, dword ptr [ebp-58]
'006c210e    50                      push eax
'006c210f    6a02                    push 02
'006c2111    ffd3                    call ebx
'#Cleanup( var_86, var_130)
'006c2113    83c40c                  add esp, 0c
'006c2116    6683bd54ffffff00        cmp word ptr [ebp+ffffff54], 00
'006c211e    0f84bd010000            je 6c22e1

If (((Trim(frmDescriptifDon)) <> (vbNullChar))) Then
'006c2124    8b0f                    mov ecx, dword ptr [edi]
'006c2126    57                      push edi
'006c2127    ff9110030000            call dword ptr [ecx+00000310]
'006c212d    8d55a8                  lea edx, dword ptr [ebp-58]
'006c2130    8945b0                  mov dword ptr [ebp-50], eax
'006c2133    52                      push edx
'006c2134    8d4598                  lea eax, dword ptr [ebp-68]
'006c2137    50                      push eax
'006c2138    c745a809000000          mov dword ptr [ebp-58], 00000009

' *** Reference to "rtcTrimVar"
'006c213f    ff15e4104000            call dword ptr [004010e4]
'006c2145    8d4d98                  lea ecx, dword ptr [ebp-68]
'006c2148    51                      push ecx

' *** Reference to "__vbaStrVarMove"
'006c2149    ff153c104000            call dword ptr [0040103c]
'006c214f    8bd0                    mov edx, eax
'006c2151    8d4de8                  lea ecx, dword ptr [ebp-18]
'006c2154    ffd6                    call esi
'006c2156    8d55e8                  lea edx, dword ptr [ebp-18]
'006c2159    52                      push edx
'006c215a    e8911ae3ff              call 4f3bf0
    Call sub_4F3BF0()
'006c215f    8bd0                    mov edx, eax
'006c2161    8d4dc4                  lea ecx, dword ptr [ebp-3c]
'006c2164    ffd6                    call esi
'006c2166    8b07                    mov eax, dword ptr [edi]
'006c2168    57                      push edi
'006c2169    ff903c030000            call dword ptr [eax+0000033c]
'006c216f    50                      push eax
'006c2170    8d4dbc                  lea ecx, dword ptr [ebp-44]
'006c2173    51                      push ecx

' *** Reference to "__vbaObjSet"
'006c2174    ff15b4104000            call dword ptr [004010b4]
    Set var_58 = Nothing
'006c217a    8bd8                    mov ebx, eax
'006c217c    8b13                    mov edx, dword ptr [ebx]
'006c217e    8d45dc                  lea eax, dword ptr [ebp-24]
'006c2181    50                      push eax
'006c2182    53                      push ebx
'006c2183    ff5250                  call dword ptr [edx+50]
'006c2186    dbe2                    fnclex
'006c2188    85c0                    test ax, ax
'006c218a    7d0f                    jge 6c219b
'006c218c    6a50                    push 50
'006c218e    683ce04300              push 0043e03c
'006c2193    53                      push ebx
'006c2194    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c2195    ff1580104000            call dword ptr [00401080]
'006c219b    8b55dc                  mov edx, dword ptr [ebp-24]
'006c219e    8d4dd8                  lea ecx, dword ptr [ebp-28]
'006c21a1    c745dc00000000          mov dword ptr [ebp-24], 00000000
'006c21a8    ffd6                    call esi
'006c21aa    8d4dd8                  lea ecx, dword ptr [ebp-28]
'006c21ad    51                      push ecx
'006c21ae    e83d1ae3ff              call 4f3bf0
    Call sub_4F3BF0()
'006c21b3    8bd0                    mov edx, eax
'006c21b5    8d4dc0                  lea ecx, dword ptr [ebp-40]
'006c21b8    ffd6                    call esi
'006c21ba    8b55c4                  mov edx, dword ptr [ebp-3c]
'006c21bd    8b5dc0                  mov ebx, dword ptr [ebp-40]
'006c21c0    8995f0feffff            mov dword ptr [ebp+fffffef0], edx
'006c21c6    33d2                    xor edx, edx
    var_num4 = Empty
'006c21c8    8955c4                  mov dword ptr [ebp-3c], edx
'006c21cb    8955c0                  mov dword ptr [ebp-40], edx
'006c21ce    8b154cb07200            mov edx, dword ptr [0072b04c]
'006c21d4    b880000000              mov eax, 00000080
'006c21d9    b903000000              mov ecx, 00000003
'006c21de    898d68ffffff            mov dword ptr [ebp+ffffff68], ecx
'006c21e4    898570ffffff            mov dword ptr [ebp+ffffff70], eax
'006c21ea    83ec10                  sub esp, 10
'006c21ed    899decfeffff            mov dword ptr [ebp+fffffeec], ebx
'006c21f3    8b1a                    mov ebx, dword ptr [edx]
'006c21f5    8bd4                    mov edx, esp
'006c21f7    890a                    mov dword ptr [edx], ecx
'006c21f9    8b8d6cffffff            mov ecx, dword ptr [ebp+ffffff6c]
'006c21ff    894a04                  mov dword ptr [edx+04], ecx
'006c2202    894208                  mov dword ptr [edx+08], eax
'006c2205    8b8574ffffff            mov eax, dword ptr [ebp+ffffff74]
'006c220b    89420c                  mov dword ptr [edx+0c], eax
'006c220e    8b95f0feffff            mov edx, dword ptr [ebp+fffffef0]
'006c2214    6800374500              push 00453700
'006c2219    8d4de4                  lea ecx, dword ptr [ebp-1c]
'006c221c    ffd6                    call esi
'006c221e    50                      push eax

' *** Reference to "__vbaStrCat"
'006c221f    ff1570104000            call dword ptr [00401070]
    var_875 = ("update dons set Spécial='") & (Trim(-344))
'006c2225    8bd0                    mov edx, eax
'006c2227    8d4de0                  lea ecx, dword ptr [ebp-20]
'006c222a    ffd6                    call esi
'006c222c    50                      push eax
'006c222d    68a4fb4400              push 0044fba4

' *** Reference to "__vbaStrCat"
'006c2232    ff1570104000            call dword ptr [00401070]
    var_876 = (var_875) & ("' where nom='")
'006c2238    8bd0                    mov edx, eax
'006c223a    8d4dd4                  lea ecx, dword ptr [ebp-2c]
'006c223d    ffd6                    call esi
'006c223f    8b95ecfeffff            mov edx, dword ptr [ebp+fffffeec]
'006c2245    50                      push eax
'006c2246    8d4dd0                  lea ecx, dword ptr [ebp-30]
'006c2249    ffd6                    call esi
'006c224b    50                      push eax

' *** Reference to "__vbaStrCat"
'006c224c    ff1570104000            call dword ptr [00401070]
    var_49 = (var_876) & (vbNullString)
'006c2252    8bd0                    mov edx, eax
'006c2254    8d4dcc                  lea ecx, dword ptr [ebp-34]
'006c2257    ffd6                    call esi
'006c2259    50                      push eax
'006c225a    6854a44300              push 0043a454

' *** Reference to "__vbaStrCat"
'006c225f    ff1570104000            call dword ptr [00401070]
    var_878 = (var_49) & ("'")
'006c2265    8bd0                    mov edx, eax
'006c2267    8d4dc8                  lea ecx, dword ptr [ebp-38]
'006c226a    ffd6                    call esi
'006c226c    8b0d4cb07200            mov ecx, dword ptr [0072b04c]
'006c2272    50                      push eax
'006c2273    51                      push ecx
'006c2274    ff535c                  call dword ptr [ebx+5c]
'006c2277    dbe2                    fnclex
'006c2279    85c0                    test ax, ax
'006c227b    7d15                    jge 6c2292
'006c227d    8b154cb07200            mov edx, dword ptr [0072b04c]
'006c2283    6a5c                    push 5c
'006c2285    68301f4300              push 00431f30
'006c228a    52                      push edx
'006c228b    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c228c    ff1580104000            call dword ptr [00401080]
'006c2292    8d45c0                  lea eax, dword ptr [ebp-40]
'006c2295    50                      push eax
'006c2296    8d4dc4                  lea ecx, dword ptr [ebp-3c]
'006c2299    51                      push ecx
'006c229a    8d55c8                  lea edx, dword ptr [ebp-38]
'006c229d    52                      push edx
'006c229e    8d45cc                  lea eax, dword ptr [ebp-34]
'006c22a1    50                      push eax
'006c22a2    8d4dd0                  lea ecx, dword ptr [ebp-30]
'006c22a5    51                      push ecx
'006c22a6    8d55d4                  lea edx, dword ptr [ebp-2c]
'006c22a9    52                      push edx
'006c22aa    8d45d8                  lea eax, dword ptr [ebp-28]
'006c22ad    50                      push eax
'006c22ae    8d4de0                  lea ecx, dword ptr [ebp-20]
'006c22b1    51                      push ecx
'006c22b2    8d55e4                  lea edx, dword ptr [ebp-1c]
'006c22b5    52                      push edx
'006c22b6    8d45e8                  lea eax, dword ptr [ebp-18]
'006c22b9    50                      push eax
'006c22ba    6a0a                    push 0a

' *** Reference to "__vbaFreeStrList"
'006c22bc    ff155c124000            call dword ptr [0040125c]
    '#Cleanup( -4672, -4672, -4676, var_39, -4680, var_39, -4684, -4688, -4672, -4672)
'006c22c2    83c42c                  add esp, 2c
'006c22c5    8d4dbc                  lea ecx, dword ptr [ebp-44]

' *** Reference to "__vbaFreeObj"
'006c22c8    ff1508134000            call dword ptr [00401308]
    '#Cleanup(var_58)
'006c22ce    8d4d98                  lea ecx, dword ptr [ebp-68]
'006c22d1    51                      push ecx
'006c22d2    8d55a8                  lea edx, dword ptr [ebp-58]
'006c22d5    52                      push edx
'006c22d6    6a02                    push 02

' *** Reference to "__vbaFreeVarList"
'006c22d8    ff1540104000            call dword ptr [00401040]
    '#Cleanup( var_86, var_130)
'006c22de    83c40c                  add esp, 0c
    
End If
'006c22e1    8b07                    mov eax, dword ptr [edi]
'006c22e3    57                      push edi
'006c22e4    ff9000070000            call dword ptr [eax+00000700]
'006c22ea    85c0                    test ax, ax
'006c22ec    7d12                    jge 6c2300

If (frmDescriptifDon) Then
'006c22ee    6800070000              push 00000700
'006c22f3    6888124300              push 00431288
'006c22f8    57                      push edi
'006c22f9    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c22fa    ff1580104000            call dword ptr [00401080]
    
End If
'006c2300    c745fc00000000          mov dword ptr [ebp-04], 00000000
'006c2307    6872236c00              push 006c2372
'006c230c    eb63                    jmp 6c2371
'006c230e    8d4dc0                  lea ecx, dword ptr [ebp-40]
'006c2311    51                      push ecx
'006c2312    8d55c4                  lea edx, dword ptr [ebp-3c]
'006c2315    52                      push edx
'006c2316    8d45c8                  lea eax, dword ptr [ebp-38]
'006c2319    50                      push eax
'006c231a    8d4dcc                  lea ecx, dword ptr [ebp-34]
'006c231d    51                      push ecx
'006c231e    8d55d0                  lea edx, dword ptr [ebp-30]
'006c2321    52                      push edx
'006c2322    8d45d4                  lea eax, dword ptr [ebp-2c]
'006c2325    50                      push eax
'006c2326    8d4dd8                  lea ecx, dword ptr [ebp-28]
'006c2329    51                      push ecx
'006c232a    8d55dc                  lea edx, dword ptr [ebp-24]
'006c232d    52                      push edx
'006c232e    8d45e0                  lea eax, dword ptr [ebp-20]
'006c2331    50                      push eax
'006c2332    8d4de4                  lea ecx, dword ptr [ebp-1c]
'006c2335    51                      push ecx
'006c2336    8d55e8                  lea edx, dword ptr [ebp-18]
'006c2339    52                      push edx
'006c233a    6a0b                    push 0b

' *** Reference to "__vbaFreeStrList"
'006c233c    ff155c124000            call dword ptr [0040125c]
'#Cleanup( -4672, -4672, -4676, var_39, var_39, -4680, var_39, -4684, -4688, -4672, -4672)
'006c2342    8d45b8                  lea eax, dword ptr [ebp-48]
'006c2345    50                      push eax
'006c2346    8d4dbc                  lea ecx, dword ptr [ebp-44]
'006c2349    51                      push ecx
'006c234a    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'006c234c    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_58, var_61)
'006c2352    8d9578ffffff            lea edx, dword ptr [ebp+ffffff78]
'006c2358    52                      push edx
'006c2359    8d4588                  lea eax, dword ptr [ebp-78]
'006c235c    50                      push eax
'006c235d    8d4d98                  lea ecx, dword ptr [ebp-68]
'006c2360    51                      push ecx
'006c2361    8d55a8                  lea edx, dword ptr [ebp-58]
'006c2364    52                      push edx
'006c2365    6a04                    push 04

' *** Reference to "__vbaFreeVarList"
'006c2367    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_86, var_130, var_131, var_87)
'006c236d    83c450                  add esp, 50
'006c2370    c3                      ret
'006c2371    c3                      ret
'006c2372    8b4508                  mov eax, dword ptr [ebp+08]
'006c2375    8b08                    mov ecx, dword ptr [eax]
'006c2377    50                      push eax
'006c2378    ff5108                  call dword ptr [ecx+08]
'006c237b    8b45fc                  mov eax, dword ptr [ebp-04]
'006c237e    8b4dec                  mov ecx, dword ptr [ebp-14]
'006c2381    5f                      pop edi
'006c2382    5e                      pop esi
    'Reference to '__except_list'
'006c2383    64890d00000000          mov dword ptr fs:[00000000], ecx
'006c238a    5b                      pop ebx
'006c238b    8be5                    mov esp, ebp
'006c238d    5d                      pop ebp
'006c238e    c20400                  ret 0004


End Sub


'Event for CmbNomDon
Private Sub CmbNomDon_Click()
'006c32c0    55                      push ebp
'006c32c1    8bec                    mov ebp, esp
'006c32c3    83ec0c                  sub esp, 0c

' *** Reference to "__vbaExceptHandler"
'006c32c6    6866724000              push 00407266
'006c32cb    64a100000000            mov ax, word ptr fs:[00000000]
'006c32d1    50                      push eax
    'Reference to '__except_list'
'006c32d2    64892500000000          mov dword ptr fs:[00000000], esp
'006c32d9    83ec24                  sub esp, 24
'006c32dc    53                      push ebx
'006c32dd    56                      push esi
'006c32de    57                      push edi
'006c32df    8965f4                  mov dword ptr [ebp-0c], esp
'006c32e2    c745f850664000          mov dword ptr [ebp-08], 00406650
'006c32e9    8b7508                  mov esi, dword ptr [ebp+08]
'006c32ec    8bc6                    mov eax, esi
'006c32ee    83e001                  and eax, 01
'006c32f1    8945fc                  mov dword ptr [ebp-04], eax
'006c32f4    83e6fe                  and esi, fffffffe
'006c32f7    8b0e                    mov ecx, dword ptr [esi]
'006c32f9    56                      push esi
'006c32fa    897508                  mov dword ptr [ebp+08], esi
'006c32fd    ff5104                  call dword ptr [ecx+04]
'006c3300    8b16                    mov edx, dword ptr [esi]
'006c3302    33c0                    xor eax, eax
var_num1 = Empty
'006c3304    56                      push esi
'006c3305    8945e8                  mov dword ptr [ebp-18], eax
'006c3308    8945e4                  mov dword ptr [ebp-1c], eax
'006c330b    8945e0                  mov dword ptr [ebp-20], eax
'006c330e    ff923c030000            call dword ptr [edx+0000033c]

' *** Reference to "__vbaObjSet"
'006c3314    8b3db4104000            mov edi, dword ptr [004010b4]
'006c331a    50                      push eax
'006c331b    8d45e0                  lea eax, dword ptr [ebp-20]
'006c331e    50                      push eax
'006c331f    ffd7                    call edi
Set var_3 = Me
'006c3321    8b0e                    mov ecx, dword ptr [esi]
'006c3323    56                      push esi
'006c3324    8bd8                    mov ebx, eax
'006c3326    ff9104030000            call dword ptr [ecx+00000304]
'006c332c    50                      push eax
'006c332d    8d55e4                  lea edx, dword ptr [ebp-1c]
'006c3330    52                      push edx
'006c3331    ffd7                    call edi
Set var_40 = var_3
'006c3333    8d4de8                  lea ecx, dword ptr [ebp-18]
'006c3336    8bf8                    mov edi, eax
'006c3338    8b07                    mov eax, dword ptr [edi]
'006c333a    51                      push ecx
'006c333b    57                      push edi
'006c333c    ff90a8000000            call dword ptr [eax+000000a8]
'006c3342    dbe2                    fnclex
'006c3344    85c0                    test ax, ax
'006c3346    7d16                    jge 6c335e
'006c3348    68a8000000              push 000000a8
'006c334d    681cb94300              push 0043b91c
'006c3352    57                      push edi

' *** Reference to "__vbaHresultCheckObj"
'006c3353    8b3d80104000            mov edi, dword ptr [00401080]
'006c3359    50                      push eax
'006c335a    ffd7                    call edi
'006c335c    eb06                    jmp 6c3364

' *** Reference to "__vbaHresultCheckObj"
'006c335e    8b3d80104000            mov edi, dword ptr [00401080]
'006c3364    8b45e8                  mov eax, dword ptr [ebp-18]
'006c3367    8b13                    mov edx, dword ptr [ebx]
'006c3369    50                      push eax
'006c336a    53                      push ebx
'006c336b    ff5254                  call dword ptr [edx+54]
'006c336e    dbe2                    fnclex
'006c3370    85c0                    test ax, ax
'006c3372    7d0b                    jge 6c337f
'006c3374    6a54                    push 54
'006c3376    683ce04300              push 0043e03c
'006c337b    53                      push ebx
'006c337c    50                      push eax
'006c337d    ffd7                    call edi
'006c337f    8d4de8                  lea ecx, dword ptr [ebp-18]

' *** Reference to "__vbaFreeStr"
'006c3382    ff150c134000            call dword ptr [0040130c]
'#Cleanup(var_41)
'006c3388    8d4de0                  lea ecx, dword ptr [ebp-20]
'006c338b    51                      push ecx
'006c338c    8d55e4                  lea edx, dword ptr [ebp-1c]
'006c338f    52                      push edx
'006c3390    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'006c3392    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_40, var_3)
'006c3398    8b06                    mov eax, dword ptr [esi]
'006c339a    83c40c                  add esp, 0c
'006c339d    56                      push esi
'006c339e    ff9000070000            call dword ptr [eax+00000700]
'006c33a4    85c0                    test ax, ax
'006c33a6    7d0e                    jge 6c33b6
'006c33a8    6800070000              push 00000700
'006c33ad    6888124300              push 00431288
'006c33b2    56                      push esi
'006c33b3    50                      push eax
'006c33b4    ffd7                    call edi
'006c33b6    c745fc00000000          mov dword ptr [ebp-04], 00000000
'006c33bd    68e2336c00              push 006c33e2
'006c33c2    eb1d                    jmp 6c33e1
'006c33c4    8d4de8                  lea ecx, dword ptr [ebp-18]

' *** Reference to "__vbaFreeStr"
'006c33c7    ff150c134000            call dword ptr [0040130c]
'#Cleanup(var_41)
'006c33cd    8d4de0                  lea ecx, dword ptr [ebp-20]
'006c33d0    51                      push ecx
'006c33d1    8d55e4                  lea edx, dword ptr [ebp-1c]
'006c33d4    52                      push edx
'006c33d5    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'006c33d7    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_40, var_3)
'006c33dd    83c40c                  add esp, 0c
'006c33e0    c3                      ret
'006c33e1    c3                      ret
'006c33e2    8b4508                  mov eax, dword ptr [ebp+08]
'006c33e5    8b08                    mov ecx, dword ptr [eax]
'006c33e7    50                      push eax
'006c33e8    ff5108                  call dword ptr [ecx+08]
'006c33eb    8b45fc                  mov eax, dword ptr [ebp-04]
'006c33ee    8b4dec                  mov ecx, dword ptr [ebp-14]
'006c33f1    5f                      pop edi
'006c33f2    5e                      pop esi
    'Reference to '__except_list'
'006c33f3    64890d00000000          mov dword ptr fs:[00000000], ecx
'006c33fa    5b                      pop ebx
'006c33fb    8be5                    mov esp, ebp
'006c33fd    5d                      pop ebp
'006c33fe    c20400                  ret 0004


End Sub


'Event for btnSuivant
Private Sub btnSuivant_Click()
'006c2fe0    55                      push ebp
'006c2fe1    8bec                    mov ebp, esp
'006c2fe3    83ec0c                  sub esp, 0c

' *** Reference to "__vbaExceptHandler"
'006c2fe6    6866724000              push 00407266
'006c2feb    64a100000000            mov ax, word ptr fs:[00000000]
'006c2ff1    50                      push eax
    'Reference to '__except_list'
'006c2ff2    64892500000000          mov dword ptr fs:[00000000], esp
'006c2ff9    83ec2c                  sub esp, 2c
'006c2ffc    53                      push ebx
'006c2ffd    56                      push esi
'006c2ffe    57                      push edi
'006c2fff    8965f4                  mov dword ptr [ebp-0c], esp
'006c3002    c745f830664000          mov dword ptr [ebp-08], 00406630
'006c3009    8b7d08                  mov edi, dword ptr [ebp+08]
'006c300c    8bc7                    mov eax, edi
'006c300e    83e001                  and eax, 01
'006c3011    8945fc                  mov dword ptr [ebp-04], eax
'006c3014    83e7fe                  and edi, fffffffe
'006c3017    8b0f                    mov ecx, dword ptr [edi]
'006c3019    57                      push edi
'006c301a    897d08                  mov dword ptr [ebp+08], edi
'006c301d    ff5104                  call dword ptr [ecx+04]
'006c3020    33c0                    xor eax, eax
var_num1 = Empty
'006c3022    66394734                cmp word ptr [edi+34], ax
'006c3026    8945e8                  mov dword ptr [ebp-18], eax
'006c3029    8945e4                  mov dword ptr [ebp-1c], eax
'006c302c    8945e0                  mov dword ptr [ebp-20], eax
'006c302f    8945dc                  mov dword ptr [ebp-24], eax
'006c3032    0f8474010000            je 6c31ac

If (arg_6 <> Me) Then
'006c3038    8b17                    mov edx, dword ptr [edi]
'006c303a    57                      push edi
'006c303b    ff9204030000            call dword ptr [edx+00000304]

' *** Reference to "__vbaObjSet"
'006c3041    8b1db4104000            mov ebx, dword ptr [004010b4]
'006c3047    50                      push eax
'006c3048    8d45e8                  lea eax, dword ptr [ebp-18]
'006c304b    50                      push eax
'006c304c    ffd3                    call ebx
    Set var_41 = Me
'006c304e    8d55e0                  lea edx, dword ptr [ebp-20]
'006c3051    8bf0                    mov esi, eax
'006c3053    8b0e                    mov ecx, dword ptr [esi]
'006c3055    52                      push edx
'006c3056    56                      push esi
'006c3057    ff91f0000000            call dword ptr [ecx+000000f0]
'006c305d    dbe2                    fnclex
'006c305f    85c0                    test ax, ax
'006c3061    7d12                    jge 6c3075
'006c3063    68f0000000              push 000000f0
'006c3068    681cb94300              push 0043b91c
'006c306d    56                      push esi
'006c306e    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c306f    ff1580104000            call dword ptr [00401080]
'006c3075    8b07                    mov eax, dword ptr [edi]
'006c3077    57                      push edi
'006c3078    ff9004030000            call dword ptr [eax+00000304]
'006c307e    50                      push eax
'006c307f    8d4de4                  lea ecx, dword ptr [ebp-1c]
'006c3082    51                      push ecx
'006c3083    ffd3                    call ebx
    Set var_40 = Nothing
'006c3085    8bf0                    mov esi, eax
'006c3087    8b16                    mov edx, dword ptr [esi]
'006c3089    8d45dc                  lea eax, dword ptr [ebp-24]
'006c308c    50                      push eax
'006c308d    56                      push esi
'006c308e    ff92e8000000            call dword ptr [edx+000000e8]
'006c3094    dbe2                    fnclex
'006c3096    85c0                    test ax, ax
'006c3098    7d12                    jge 6c30ac
'006c309a    68e8000000              push 000000e8
'006c309f    681cb94300              push 0043b91c
'006c30a4    56                      push esi
'006c30a5    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c30a6    ff1580104000            call dword ptr [00401080]
'006c30ac    668b75e0                mov si, word ptr [ebp-20]
'006c30b0    6683c601                add si, 01
    var_num8 = Me + 1
'006c30b4    0f8030010000            jo 6c31ea
'006c30ba    662b75dc                sub si, word ptr [ebp-24]
    var_num8 = var_num8 - Me
'006c30be    8d4de4                  lea ecx, dword ptr [ebp-1c]
'006c30c1    66f7de                  neg si
'006c30c4    51                      push ecx
'006c30c5    8d55e8                  lea edx, dword ptr [ebp-18]
'006c30c8    52                      push edx
'006c30c9    6a02                    push 02
'006c30cb    1bf6                    sbb esi, esi
'006c30cd    46                      inc esi
'006c30ce    f7de                    neg esi

' *** Reference to "__vbaFreeObjList"
'006c30d0    ff154c104000            call dword ptr [0040104c]
    '#Cleanup( var_41, var_40)
'006c30d6    8b07                    mov eax, dword ptr [edi]
'006c30d8    83c40c                  add esp, 0c
'006c30db    6685f6                  test esi, esi
'006c30de    57                      push edi
'006c30df    7440                    je 6c3121
    
    If (    -(CBool((var_num8)))) Then
'006c30e1    ff9004030000            call dword ptr [eax+00000304]
'006c30e7    50                      push eax
'006c30e8    8d4de8                  lea ecx, dword ptr [ebp-18]
'006c30eb    51                      push ecx
'006c30ec    ffd3                    call ebx
    Set var_41 = Nothing
'006c30ee    8bf0                    mov esi, eax
'006c30f0    8b16                    mov edx, dword ptr [esi]
'006c30f2    6a00                    push 00
'006c30f4    56                      push esi
'006c30f5    ff92f4000000            call dword ptr [edx+000000f4]
'006c30fb    dbe2                    fnclex
'006c30fd    85c0                    test ax, ax
'006c30ff    7d12                    jge 6c3113
'006c3101    68f4000000              push 000000f4
'006c3106    681cb94300              push 0043b91c
'006c310b    56                      push esi
'006c310c    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c310d    ff1580104000            call dword ptr [00401080]
'006c3113    8d4de8                  lea ecx, dword ptr [ebp-18]

' *** Reference to "__vbaFreeObj"
'006c3116    ff1508134000            call dword ptr [00401308]
    '#Cleanup(var_41)
'006c311c    e989000000              jmp 6c31aa
    
End If
'006c3121    ff9004030000            call dword ptr [eax+00000304]
'006c3127    50                      push eax
'006c3128    8d4de4                  lea ecx, dword ptr [ebp-1c]
'006c312b    51                      push ecx
'006c312c    ffd3                    call ebx
Set var_40 = Nothing
'006c312e    8b17                    mov edx, dword ptr [edi]
'006c3130    57                      push edi
'006c3131    8945d0                  mov dword ptr [ebp-30], eax
'006c3134    ff9204030000            call dword ptr [edx+00000304]
'006c313a    50                      push eax
'006c313b    8d45e8                  lea eax, dword ptr [ebp-18]
'006c313e    50                      push eax
'006c313f    ffd3                    call ebx
Set var_41 = Nothing
'006c3141    8d55e0                  lea edx, dword ptr [ebp-20]
'006c3144    8bf0                    mov esi, eax
'006c3146    8b0e                    mov ecx, dword ptr [esi]
'006c3148    52                      push edx
'006c3149    56                      push esi
'006c314a    ff91f0000000            call dword ptr [ecx+000000f0]
'006c3150    dbe2                    fnclex
'006c3152    85c0                    test ax, ax
'006c3154    7d12                    jge 6c3168
'006c3156    68f0000000              push 000000f0
'006c315b    681cb94300              push 0043b91c
'006c3160    56                      push esi
'006c3161    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c3162    ff1580104000            call dword ptr [00401080]
'006c3168    668b4de0                mov cx, word ptr [ebp-20]
'006c316c    8b75d0                  mov esi, dword ptr [ebp-30]
'006c316f    8b06                    mov eax, dword ptr [esi]
'006c3171    6683c101                add cx, 01
var_num3 = Me + 1
'006c3175    7073                    jo 6c31ea
'006c3177    51                      push ecx
'006c3178    56                      push esi
'006c3179    ff90f4000000            call dword ptr [eax+000000f4]
'006c317f    dbe2                    fnclex
'006c3181    85c0                    test ax, ax
'006c3183    7d12                    jge 6c3197
'006c3185    68f4000000              push 000000f4
'006c318a    681cb94300              push 0043b91c
'006c318f    56                      push esi
'006c3190    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c3191    ff1580104000            call dword ptr [00401080]
'006c3197    8d55e4                  lea edx, dword ptr [ebp-1c]
'006c319a    52                      push edx
'006c319b    8d45e8                  lea eax, dword ptr [ebp-18]
'006c319e    50                      push eax
'006c319f    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'006c31a1    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_41, var_40)
'006c31a7    83c40c                  add esp, 0c
'006c31aa    33c0                    xor eax, eax

'ERROR: Two many next close:
End If
'006c31ac    8945fc                  mov dword ptr [ebp-04], eax
'006c31af    68cb316c00              push 006c31cb
'006c31b4    eb14                    jmp 6c31ca
'006c31b6    8d4de4                  lea ecx, dword ptr [ebp-1c]
'006c31b9    51                      push ecx
'006c31ba    8d55e8                  lea edx, dword ptr [ebp-18]
'006c31bd    52                      push edx
'006c31be    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'006c31c0    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_41, var_40)
'006c31c6    83c40c                  add esp, 0c
'006c31c9    c3                      ret
'006c31ca    c3                      ret
'006c31cb    8b4508                  mov eax, dword ptr [ebp+08]
'006c31ce    8b08                    mov ecx, dword ptr [eax]
'006c31d0    50                      push eax
'006c31d1    ff5108                  call dword ptr [ecx+08]
'006c31d4    8b45fc                  mov eax, dword ptr [ebp-04]
'006c31d7    8b4dec                  mov ecx, dword ptr [ebp-14]
'006c31da    5f                      pop edi
'006c31db    5e                      pop esi
    'Reference to '__except_list'
'006c31dc    64890d00000000          mov dword ptr fs:[00000000], ecx
'006c31e3    5b                      pop ebx
'006c31e4    8be5                    mov esp, ebp
'006c31e6    5d                      pop ebp
'006c31e7    c20400                  ret 0004


End Sub


'Event for btnPrecedent
Private Sub btnPrecedent_Click()
'006c2dd0    55                      push ebp
'006c2dd1    8bec                    mov ebp, esp
'006c2dd3    83ec0c                  sub esp, 0c

' *** Reference to "__vbaExceptHandler"
'006c2dd6    6866724000              push 00407266
'006c2ddb    64a100000000            mov ax, word ptr fs:[00000000]
'006c2de1    50                      push eax
    'Reference to '__except_list'
'006c2de2    64892500000000          mov dword ptr fs:[00000000], esp
'006c2de9    83ec24                  sub esp, 24
'006c2dec    53                      push ebx
'006c2ded    56                      push esi
'006c2dee    57                      push edi
'006c2def    8965f4                  mov dword ptr [ebp-0c], esp
'006c2df2    c745f820664000          mov dword ptr [ebp-08], 00406620
'006c2df9    8b7508                  mov esi, dword ptr [ebp+08]
'006c2dfc    8bc6                    mov eax, esi
'006c2dfe    83e001                  and eax, 01
'006c2e01    8945fc                  mov dword ptr [ebp-04], eax
'006c2e04    83e6fe                  and esi, fffffffe
'006c2e07    8b0e                    mov ecx, dword ptr [esi]
'006c2e09    56                      push esi
'006c2e0a    897508                  mov dword ptr [ebp+08], esi
'006c2e0d    ff5104                  call dword ptr [ecx+04]
'006c2e10    33c0                    xor eax, eax
var_num1 = Empty
'006c2e12    66394634                cmp word ptr [esi+34], ax
'006c2e16    8945e8                  mov dword ptr [ebp-18], eax
'006c2e19    8945e4                  mov dword ptr [ebp-1c], eax
'006c2e1c    8945e0                  mov dword ptr [ebp-20], eax
'006c2e1f    0f8474010000            je 6c2f99

If (arg_6 <> Me) Then
'006c2e25    8b16                    mov edx, dword ptr [esi]
'006c2e27    56                      push esi
'006c2e28    ff9204030000            call dword ptr [edx+00000304]

' *** Reference to "__vbaObjSet"
'006c2e2e    8b1db4104000            mov ebx, dword ptr [004010b4]
'006c2e34    50                      push eax
'006c2e35    8d45e8                  lea eax, dword ptr [ebp-18]
'006c2e38    50                      push eax
'006c2e39    ffd3                    call ebx
    Set var_41 = Me
'006c2e3b    8d55e0                  lea edx, dword ptr [ebp-20]
'006c2e3e    8bf8                    mov edi, eax
'006c2e40    8b0f                    mov ecx, dword ptr [edi]
'006c2e42    52                      push edx
'006c2e43    57                      push edi
'006c2e44    ff91f0000000            call dword ptr [ecx+000000f0]
'006c2e4a    dbe2                    fnclex
'006c2e4c    85c0                    test ax, ax
'006c2e4e    7d12                    jge 6c2e62
'006c2e50    68f0000000              push 000000f0
'006c2e55    681cb94300              push 0043b91c
'006c2e5a    57                      push edi
'006c2e5b    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c2e5c    ff1580104000            call dword ptr [00401080]
'006c2e62    668b7de0                mov di, word ptr [ebp-20]
'006c2e66    6683ef01                sub di, 01
    var_num7 = Me - 1
'006c2e6a    0f8067010000            jo 6c2fd7
'006c2e70    6647                    inc di
'006c2e72    66f7df                  neg di
'006c2e75    8d4de8                  lea ecx, dword ptr [ebp-18]
'006c2e78    1bff                    sbb edi, edi
'006c2e7a    47                      inc edi
'006c2e7b    f7df                    neg edi

' *** Reference to "__vbaFreeObj"
'006c2e7d    ff1508134000            call dword ptr [00401308]
    '#Cleanup(var_41)
'006c2e83    6685ff                  test edi, edi
'006c2e86    0f8483000000            je 6c2f0f
    
    If (    -(CBool((var_num7)))) Then
'006c2e8c    8b06                    mov eax, dword ptr [esi]
'006c2e8e    56                      push esi
'006c2e8f    ff9004030000            call dword ptr [eax+00000304]
'006c2e95    50                      push eax
'006c2e96    8d4de4                  lea ecx, dword ptr [ebp-1c]
'006c2e99    51                      push ecx
'006c2e9a    ffd3                    call ebx
    Set var_40 = Nothing
'006c2e9c    8b16                    mov edx, dword ptr [esi]
'006c2e9e    56                      push esi
'006c2e9f    8bf8                    mov edi, eax
'006c2ea1    ff9204030000            call dword ptr [edx+00000304]
'006c2ea7    50                      push eax
'006c2ea8    8d45e8                  lea eax, dword ptr [ebp-18]
'006c2eab    50                      push eax
'006c2eac    ffd3                    call ebx
    Set var_41 = Nothing
'006c2eae    8d55e0                  lea edx, dword ptr [ebp-20]
'006c2eb1    8bf0                    mov esi, eax
'006c2eb3    8b0e                    mov ecx, dword ptr [esi]
'006c2eb5    52                      push edx
'006c2eb6    56                      push esi
'006c2eb7    ff91e8000000            call dword ptr [ecx+000000e8]
'006c2ebd    dbe2                    fnclex
'006c2ebf    85c0                    test ax, ax
'006c2ec1    7d12                    jge 6c2ed5
'006c2ec3    68e8000000              push 000000e8
'006c2ec8    681cb94300              push 0043b91c
'006c2ecd    56                      push esi
'006c2ece    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c2ecf    ff1580104000            call dword ptr [00401080]
'006c2ed5    668b4de0                mov cx, word ptr [ebp-20]
'006c2ed9    8b07                    mov eax, dword ptr [edi]
'006c2edb    6683e901                sub cx, 01
    var_num3 = Me - 1
'006c2edf    0f80f2000000            jo 6c2fd7
'006c2ee5    51                      push ecx
'006c2ee6    57                      push edi
'006c2ee7    ff90f4000000            call dword ptr [eax+000000f4]
'006c2eed    dbe2                    fnclex
'006c2eef    85c0                    test ax, ax
'006c2ef1    7d12                    jge 6c2f05
'006c2ef3    68f4000000              push 000000f4
'006c2ef8    681cb94300              push 0043b91c
'006c2efd    57                      push edi
'006c2efe    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c2eff    ff1580104000            call dword ptr [00401080]
'006c2f05    8d55e4                  lea edx, dword ptr [ebp-1c]
'006c2f08    52                      push edx
'006c2f09    8d45e8                  lea eax, dword ptr [ebp-18]
'006c2f0c    50                      push eax
'006c2f0d    eb7d                    jmp 6c2f8c
    
End If
'006c2f0f    8b0e                    mov ecx, dword ptr [esi]
'006c2f11    56                      push esi
'006c2f12    ff9104030000            call dword ptr [ecx+00000304]
'006c2f18    50                      push eax
'006c2f19    8d55e4                  lea edx, dword ptr [ebp-1c]
'006c2f1c    52                      push edx
'006c2f1d    ffd3                    call ebx
Set var_40 = var_41
'006c2f1f    8bf8                    mov edi, eax
'006c2f21    8b06                    mov eax, dword ptr [esi]
'006c2f23    56                      push esi
'006c2f24    ff9004030000            call dword ptr [eax+00000304]
'006c2f2a    50                      push eax
'006c2f2b    8d4de8                  lea ecx, dword ptr [ebp-18]
'006c2f2e    51                      push ecx
'006c2f2f    ffd3                    call ebx
Set var_41 = Nothing
'006c2f31    8bf0                    mov esi, eax
'006c2f33    8b16                    mov edx, dword ptr [esi]
'006c2f35    8d45e0                  lea eax, dword ptr [ebp-20]
'006c2f38    50                      push eax
'006c2f39    56                      push esi
'006c2f3a    ff92f0000000            call dword ptr [edx+000000f0]
'006c2f40    dbe2                    fnclex
'006c2f42    85c0                    test ax, ax
'006c2f44    7d12                    jge 6c2f58
'006c2f46    68f0000000              push 000000f0
'006c2f4b    681cb94300              push 0043b91c
'006c2f50    56                      push esi
'006c2f51    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c2f52    ff1580104000            call dword ptr [00401080]
'006c2f58    668b55e0                mov dx, word ptr [ebp-20]
'006c2f5c    8b0f                    mov ecx, dword ptr [edi]
'006c2f5e    6683ea01                sub dx, 01
var_num4 = Me - 1
'006c2f62    7073                    jo 6c2fd7
'006c2f64    52                      push edx
'006c2f65    57                      push edi
'006c2f66    ff91f4000000            call dword ptr [ecx+000000f4]
'006c2f6c    dbe2                    fnclex
'006c2f6e    85c0                    test ax, ax
'006c2f70    7d12                    jge 6c2f84
'006c2f72    68f4000000              push 000000f4
'006c2f77    681cb94300              push 0043b91c
'006c2f7c    57                      push edi
'006c2f7d    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006c2f7e    ff1580104000            call dword ptr [00401080]
'006c2f84    8d45e4                  lea eax, dword ptr [ebp-1c]
'006c2f87    50                      push eax
'006c2f88    8d4de8                  lea ecx, dword ptr [ebp-18]
'006c2f8b    51                      push ecx
'006c2f8c    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'006c2f8e    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_41, var_40)
'006c2f94    83c40c                  add esp, 0c
'006c2f97    33c0                    xor eax, eax

'ERROR: Two many next close:
End If
'006c2f99    8945fc                  mov dword ptr [ebp-04], eax
'006c2f9c    68b82f6c00              push 006c2fb8
'006c2fa1    eb14                    jmp 6c2fb7
'006c2fa3    8d55e4                  lea edx, dword ptr [ebp-1c]
'006c2fa6    52                      push edx
'006c2fa7    8d45e8                  lea eax, dword ptr [ebp-18]
'006c2faa    50                      push eax
'006c2fab    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'006c2fad    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_41, var_40)
'006c2fb3    83c40c                  add esp, 0c
'006c2fb6    c3                      ret
'006c2fb7    c3                      ret
'006c2fb8    8b4508                  mov eax, dword ptr [ebp+08]
'006c2fbb    8b08                    mov ecx, dword ptr [eax]
'006c2fbd    50                      push eax
'006c2fbe    ff5108                  call dword ptr [ecx+08]
'006c2fc1    8b45fc                  mov eax, dword ptr [ebp-04]
'006c2fc4    8b4dec                  mov ecx, dword ptr [ebp-14]
'006c2fc7    5f                      pop edi
'006c2fc8    5e                      pop esi
    'Reference to '__except_list'
'006c2fc9    64890d00000000          mov dword ptr fs:[00000000], ecx
'006c2fd0    5b                      pop ebx
'006c2fd1    8be5                    mov esp, ebp
'006c2fd3    5d                      pop ebp
'006c2fd4    c20400                  ret 0004


End Sub


