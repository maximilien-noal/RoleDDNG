VERSION 5.00
Object = "{D76D7130-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "C:\WINDOWS\SysWow64\Vsflex7d.ocx"

Begin VB.Form frmImport
    Caption = "Importation de personnage d'une autre base de donnée"
    ScaleMode = 1
    AutoRedraw = 0              'False
    FontTransparent = -1              'True
    BorderStyle = 1
    LinkTopic = "Form1"
    MaxButton = 0              'False
    MinButton = 0              'False
    ClientLeft   = 45
    ClientTop    = 435
    ClientWidth  = 10140
    ClientHeight = 7305
    BeginProperty Font
        Name          = "Times New Roman"
        Size          = 9
        Charset       = 0
        Weight        = 400
        Underline     = 0              'False
        Italic        = 0              'False
        Strikethrough = 0              'False
    EndProperty
    StartupPosition = 3
    Begin VB.CommandButton BtnFermer
        Caption = "Fermer"
        Left   = 8520
        Top    = 6840
        Width  = 1335
        Height = 375
        TabIndex = 5
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
    Begin VB.CommandButton BtnImport
        Caption = "Importer"
        Left   = 7080
        Top    = 6840
        Width  = 1335
        Height = 375
        TabIndex = 4
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
    Begin VB.CheckBox chkObjet
        Caption = "Avec les objets"
        Left   = 120
        Top    = 6840
        Width  = 1935
        Height = 375
        TabIndex = 3
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
    Begin VB.CommandButton btnTous
        Caption = "Tous les personnages"
        Left   = 2280
        Top    = 6840
        Width  = 1695
        Height = 375
        TabIndex = 2
    End
    Begin VB.CommandButton btnAucun
        Caption = "Aucun personnage"
        Left   = 4080
        Top    = 6840
        Width  = 1695
        Height = 375
        TabIndex = 1
    End
    Begin VSFlex7DAOCtl.VSFlexGrid vsPersoImport
        Left   = 0
        Top    = 0
        Width  = 10140
        Height = 6705
        TabIndex = 0
        OleObjectBlob = frmImport.frx:0000
    End
End
Public Function AffecteImport(arg_0 As Unknow, arg_1 As Unknow, arg_2 As Unknow, arg_3 As Unknow, arg_4 As Unknow, arg_5 As Unknow, arg_6 As Unknow, arg_7 As Unknow, arg_8 As Unknow, arg_9 As Unknow, arg_A As Unknow, arg_B As Unknow, arg_C As Unknow, arg_D As Unknow, arg_E As Unknow, arg_F As Unknow, arg_10 As Unknow, arg_11 As Unknow, arg_12 As Unknow, arg_13 As Unknow, arg_14 As Unknow, arg_15 As Unknow, arg_16 As Unknow, arg_17 As Unknow, arg_18 As Unknow, arg_19 As Unknow, arg_1A As Unknow, arg_1B As Unknow, arg_1C As Unknow, arg_1D As Unknow, arg_1E As Unknow, arg_1F As Unknow, arg_20 As Unknow, arg_21 As Unknow, arg_22 As Unknow, arg_23 As Unknow, arg_24 As Unknow, arg_25 As Unknow, arg_26 As Unknow, arg_27 As Unknow, arg_28 As Unknow, arg_29 As Unknow, arg_2A As Unknow, arg_2B As Unknow, arg_2C As Unknow, arg_2D As Unknow, arg_2E As Unknow, arg_2F As Unknow, arg_30 As Unknow, arg_31 As Unknow, arg_32 As Unknow, arg_33 As Unknow, arg_34 As Unknow, arg_35 As Unknow, arg_36 As Unknow, arg_37 As Unknow, arg_38 As Unknow, arg_39 As Unknow, arg_3A As Unknow, arg_3B As Unknow, arg_3C As Unknow)
'0071a940    55                      push ebp
'0071a941    8bec                    mov ebp, esp
'0071a943    83ec0c                  sub esp, 0c

' *** Reference to "__vbaExceptHandler"
'0071a946    6866724000              push 00407266
'0071a94b    64a100000000            mov ax, word ptr fs:[00000000]
'0071a951    50                      push eax
    'Reference to '__except_list'
'0071a952    64892500000000          mov dword ptr fs:[00000000], esp
'0071a959    81eca8000000            sub esp, 000000a8
'0071a95f    53                      push ebx
'0071a960    56                      push esi
'0071a961    57                      push edi
'0071a962    8965f4                  mov dword ptr [ebp-0c], esp
'0071a965    c745f8f06d4000          mov dword ptr [ebp-08], 00406df0
'0071a96c    33db                    xor ebx, ebx
'0071a96e    895dfc                  mov dword ptr [ebp-04], ebx
'0071a971    8b7d08                  mov edi, dword ptr [ebp+08]
'0071a974    8b07                    mov eax, dword ptr [edi]
'0071a976    57                      push edi
'0071a977    ff5004                  call dword ptr [eax+04]
'0071a97a    8b0f                    mov ecx, dword ptr [edi]
'0071a97c    53                      push ebx
'0071a97d    6a07                    push 07
'0071a97f    57                      push edi
'0071a980    895de4                  mov dword ptr [ebp-1c], ebx
'0071a983    895dd4                  mov dword ptr [ebp-2c], ebx
'0071a986    895dc4                  mov dword ptr [ebp-3c], ebx
'0071a989    895da4                  mov dword ptr [ebp-5c], ebx
'0071a98c    895d84                  mov dword ptr [ebp-7c], ebx
'0071a98f    899d64ffffff            mov dword ptr [ebp+ffffff64], ebx
'0071a995    ff9110030000            call dword ptr [ecx+00000310]
'0071a99b    50                      push eax
'0071a99c    8d55e4                  lea edx, dword ptr [ebp-1c]
'0071a99f    52                      push edx

' *** Reference to "__vbaObjSet"
'0071a9a0    ff15b4104000            call dword ptr [004010b4]
Set var_40 = Nothing
'0071a9a6    50                      push eax
'0071a9a7    8d45d4                  lea eax, dword ptr [ebp-2c]
'0071a9aa    50                      push eax

' *** Reference to "__vbaLateIdCallLd"
'0071a9ab    ff157c114000            call dword ptr [0040117c]
var_44 = var_40.UNK_-256 - 12_7
'0071a9b1    83c410                  add esp, 10
'0071a9b4    50                      push eax

' *** Reference to "__vbaI4Var"
'0071a9b5    ff157c124000            call dword ptr [0040127c]
'0071a9bb    8bc8                    mov ecx, eax
'0071a9bd    83e901                  sub ecx, 01
var_num3 = CLng(var_44) - 1
'0071a9c0    0f8039010000            jo 71aaff

' *** Reference to "__vbaI2I4"
'0071a9c6    ff1550114000            call dword ptr [00401150]
'0071a9cc    8d4de4                  lea ecx, dword ptr [ebp-1c]
'0071a9cf    c78550ffffff01000000    mov dword ptr [ebp+ffffff50], 00000001
'0071a9d9    33f6                    xor esi, esi
'0071a9db    89854cffffff            mov dword ptr [ebp+ffffff4c], eax

' *** Reference to "__vbaFreeObj"
'0071a9e1    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_40)
'0071a9e7    8d4dd4                  lea ecx, dword ptr [ebp-2c]

' *** Reference to "__vbaFreeVar"
'0071a9ea    ff1528104000            call dword ptr [00401028]
'#Cleanup(var_44)
'0071a9f0    663bb54cffffff          cmp si, word ptr [ebp+ffffff4c]
'0071a9f7    0f8fc8000000            jg 71aac5
'0071a9fd    83ec10                  sub esp, 10
'0071aa00    8bdc                    mov ebx, esp
'0071aa02    33c0                    xor eax, eax
var_num1 = Empty
'0071aa04    b903000000              mov ecx, 00000003
'0071aa09    890b                    mov dword ptr [ebx], ecx
'0071aa0b    8b4dc8                  mov ecx, dword ptr [ebp-38]
'0071aa0e    894b04                  mov dword ptr [ebx+04], ecx
'0071aa11    894308                  mov dword ptr [ebx+08], eax
'0071aa14    8b45d0                  mov eax, dword ptr [ebp-30]
'0071aa17    89430c                  mov dword ptr [ebx+0c], eax
'0071aa1a    83ec10                  sub esp, 10
'0071aa1d    8bcc                    mov ecx, esp
'0071aa1f    ba02000000              mov edx, 00000002
'0071aa24    8911                    mov dword ptr [ecx], edx
'0071aa26    8b55a8                  mov edx, dword ptr [ebp-58]
'0071aa29    895104                  mov dword ptr [ecx+04], edx
'0071aa2c    8b55b0                  mov edx, dword ptr [ebp-50]
'0071aa2f    83ec10                  sub esp, 10
'0071aa32    668975ac                mov word ptr [ebp-54], si
'0071aa36    8b45ac                  mov eax, dword ptr [ebp-54]
'0071aa39    894108                  mov dword ptr [ecx+08], eax
'0071aa3c    89510c                  mov dword ptr [ecx+0c], edx
'0071aa3f    8b5588                  mov edx, dword ptr [ebp-78]
'0071aa42    8bcc                    mov ecx, esp
'0071aa44    b802000000              mov eax, 00000002
'0071aa49    8901                    mov dword ptr [ecx], eax
'0071aa4b    895104                  mov dword ptr [ecx+04], edx
'0071aa4e    8b9568ffffff            mov edx, dword ptr [ebp+ffffff68]
'0071aa54    b805000000              mov eax, 00000005
'0071aa59    894108                  mov dword ptr [ecx+08], eax
'0071aa5c    8b4590                  mov eax, dword ptr [ebp-70]
'0071aa5f    89410c                  mov dword ptr [ecx+0c], eax
'0071aa62    83ec10                  sub esp, 10
'0071aa65    8bcc                    mov ecx, esp
'0071aa67    b808000000              mov eax, 00000008
'0071aa6c    8901                    mov dword ptr [ecx], eax
'0071aa6e    8b450c                  mov eax, dword ptr [ebp+0c]
'0071aa71    8b00                    mov eax, dword ptr [eax]
'0071aa73    895104                  mov dword ptr [ecx+04], edx
'0071aa76    8b9570ffffff            mov edx, dword ptr [ebp+ffffff70]
'0071aa7c    6a03                    push 03
'0071aa7e    894108                  mov dword ptr [ecx+08], eax
'0071aa81    8b07                    mov eax, dword ptr [edi]
'0071aa83    689e000000              push 0000009e
'0071aa88    57                      push edi
'0071aa89    89510c                  mov dword ptr [ecx+0c], edx
'0071aa8c    ff9010030000            call dword ptr [eax+00000310]
'0071aa92    50                      push eax
'0071aa93    8d4de4                  lea ecx, dword ptr [ebp-1c]
'0071aa96    51                      push ecx

' *** Reference to "__vbaObjSet"
'0071aa97    ff15b4104000            call dword ptr [004010b4]
Set var_40 = Nothing
'0071aa9d    50                      push eax

' *** Reference to "__vbaLateIdCallSt"
'0071aa9e    ff159c114000            call dword ptr [0040119c]
'0071aaa4    83c44c                  add esp, 4c
'0071aaa7    8d4de4                  lea ecx, dword ptr [ebp-1c]

' *** Reference to "__vbaFreeObj"
'0071aaaa    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_40)
'0071aab0    668b9550ffffff          mov dx, word ptr [ebp+ffffff50]
'0071aab7    6603d6                  add dx, si
var_num4 = var_21 + 0
'0071aaba    7043                    jo 71aaff
'0071aabc    33db                    xor ebx, ebx
var_num2 = Empty
'0071aabe    8bf2                    mov esi, edx
'0071aac0    e92bffffff              jmp 71a9f0
'0071aac5    68e0aa7100              push 0071aae0
'0071aaca    eb13                    jmp 71aadf
'0071aacc    8d4de4                  lea ecx, dword ptr [ebp-1c]

' *** Reference to "__vbaFreeObj"
'0071aacf    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_40)
'0071aad5    8d4dd4                  lea ecx, dword ptr [ebp-2c]

' *** Reference to "__vbaFreeVar"
'0071aad8    ff1528104000            call dword ptr [00401028]
'#Cleanup(var_44)
'0071aade    c3                      ret
'0071aadf    c3                      ret
'0071aae0    8b4508                  mov eax, dword ptr [ebp+08]
'0071aae3    8b08                    mov ecx, dword ptr [eax]
'0071aae5    50                      push eax
'0071aae6    ff5108                  call dword ptr [ecx+08]
'0071aae9    8b45fc                  mov eax, dword ptr [ebp-04]
'0071aaec    8b4dec                  mov ecx, dword ptr [ebp-14]
'0071aaef    5f                      pop edi
'0071aaf0    5e                      pop esi
    'Reference to '__except_list'
'0071aaf1    64890d00000000          mov dword ptr fs:[00000000], ecx
'0071aaf8    5b                      pop ebx
'0071aaf9    8be5                    mov esp, ebp
'0071aafb    5d                      pop ebp
'0071aafc    c20800                  ret 0008


End Function


'Event for chkObjet
Private Sub chkObjet_Click()
'0071ab10    55                      push ebp
'0071ab11    8bec                    mov ebp, esp
'0071ab13    83ec0c                  sub esp, 0c

' *** Reference to "__vbaExceptHandler"
'0071ab16    6866724000              push 00407266
'0071ab1b    64a100000000            mov ax, word ptr fs:[00000000]
'0071ab21    50                      push eax
    'Reference to '__except_list'
'0071ab22    64892500000000          mov dword ptr fs:[00000000], esp
'0071ab29    83ec1c                  sub esp, 1c
'0071ab2c    53                      push ebx
'0071ab2d    56                      push esi
'0071ab2e    57                      push edi
'0071ab2f    8965f4                  mov dword ptr [ebp-0c], esp
'0071ab32    c745f8006e4000          mov dword ptr [ebp-08], 00406e00
'0071ab39    8b7508                  mov esi, dword ptr [ebp+08]
'0071ab3c    8bc6                    mov eax, esi
'0071ab3e    83e001                  and eax, 01
'0071ab41    8945fc                  mov dword ptr [ebp-04], eax
'0071ab44    83e6fe                  and esi, fffffffe
'0071ab47    8b0e                    mov ecx, dword ptr [esi]
'0071ab49    56                      push esi
'0071ab4a    897508                  mov dword ptr [ebp+08], esi
'0071ab4d    ff5104                  call dword ptr [ecx+04]
'0071ab50    8b16                    mov edx, dword ptr [esi]
'0071ab52    33c0                    xor eax, eax
var_num1 = Empty
'0071ab54    56                      push esi
'0071ab55    8945e8                  mov dword ptr [ebp-18], eax
'0071ab58    8945e4                  mov dword ptr [ebp-1c], eax
'0071ab5b    ff9204030000            call dword ptr [edx+00000304]

' *** Reference to "__vbaObjSet"
'0071ab61    8b1db4104000            mov ebx, dword ptr [004010b4]
'0071ab67    50                      push eax
'0071ab68    8d45e8                  lea eax, dword ptr [ebp-18]
'0071ab6b    50                      push eax
'0071ab6c    ffd3                    call ebx
Set var_41 = Me
'0071ab6e    8d55e4                  lea edx, dword ptr [ebp-1c]
'0071ab71    8bf8                    mov edi, eax
'0071ab73    8b0f                    mov ecx, dword ptr [edi]
'0071ab75    52                      push edx
'0071ab76    57                      push edi
'0071ab77    ff91e0000000            call dword ptr [ecx+000000e0]
'0071ab7d    dbe2                    fnclex
'0071ab7f    85c0                    test ax, ax
'0071ab81    7d12                    jge 71ab95
'0071ab83    68e0000000              push 000000e0
'0071ab88    68dce24300              push 0043e2dc
'0071ab8d    57                      push edi
'0071ab8e    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0071ab8f    ff1580104000            call dword ptr [00401080]
'0071ab95    8b7de4                  mov edi, dword ptr [ebp-1c]
'0071ab98    8d4de8                  lea ecx, dword ptr [ebp-18]

' *** Reference to "__vbaFreeObj"
'0071ab9b    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_41)
'0071aba1    6685ff                  test edi, edi
'0071aba4    8b06                    mov eax, dword ptr [esi]
'0071aba6    56                      push esi
'0071aba7    7422                    je 71abcb
'0071aba9    ff9004030000            call dword ptr [eax+00000304]
'0071abaf    50                      push eax
'0071abb0    8d4de8                  lea ecx, dword ptr [ebp-18]
'0071abb3    51                      push ecx
'0071abb4    ffd3                    call ebx
Set var_41 = Nothing
'0071abb6    8bf0                    mov esi, eax
'0071abb8    8b16                    mov edx, dword ptr [esi]
'0071abba    68cc6d4500              push 00456dcc
'0071abbf    56                      push esi
'0071abc0    ff5254                  call dword ptr [edx+54]
'0071abc3    dbe2                    fnclex
'0071abc5    85c0                    test ax, ax
'0071abc7    7d31                    jge 71abfa
'0071abc9    eb20                    jmp 71abeb
'0071abcb    ff9004030000            call dword ptr [eax+00000304]
'0071abd1    50                      push eax
'0071abd2    8d4de8                  lea ecx, dword ptr [ebp-18]
'0071abd5    51                      push ecx
'0071abd6    ffd3                    call ebx
Set var_41 = Nothing
'0071abd8    8bf0                    mov esi, eax
'0071abda    8b16                    mov edx, dword ptr [esi]
'0071abdc    68cc6b4500              push 00456bcc
'0071abe1    56                      push esi
'0071abe2    ff5254                  call dword ptr [edx+54]
'0071abe5    dbe2                    fnclex
'0071abe7    85c0                    test ax, ax
'0071abe9    7d0f                    jge 71abfa
'0071abeb    6a54                    push 54
'0071abed    68dce24300              push 0043e2dc
'0071abf2    56                      push esi
'0071abf3    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0071abf4    ff1580104000            call dword ptr [00401080]
'0071abfa    8d4de8                  lea ecx, dword ptr [ebp-18]

' *** Reference to "__vbaFreeObj"
'0071abfd    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_41)
'0071ac03    c745fc00000000          mov dword ptr [ebp-04], 00000000
'0071ac0a    681cac7100              push 0071ac1c
'0071ac0f    eb0a                    jmp 71ac1b
'0071ac11    8d4de8                  lea ecx, dword ptr [ebp-18]

' *** Reference to "__vbaFreeObj"
'0071ac14    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_41)
'0071ac1a    c3                      ret
'0071ac1b    c3                      ret
'0071ac1c    8b4508                  mov eax, dword ptr [ebp+08]
'0071ac1f    8b08                    mov ecx, dword ptr [eax]
'0071ac21    50                      push eax
'0071ac22    ff5108                  call dword ptr [ecx+08]
'0071ac25    8b45fc                  mov eax, dword ptr [ebp-04]
'0071ac28    8b4dec                  mov ecx, dword ptr [ebp-14]
'0071ac2b    5f                      pop edi
'0071ac2c    5e                      pop esi
    'Reference to '__except_list'
'0071ac2d    64890d00000000          mov dword ptr fs:[00000000], ecx
'0071ac34    5b                      pop ebx
'0071ac35    8be5                    mov esp, ebp
'0071ac37    5d                      pop ebp
'0071ac38    c20400                  ret 0004


End Sub


'Event for Form
Private Sub Form_Load()
'0071ad20    55                      push ebp
'0071ad21    8bec                    mov ebp, esp
'0071ad23    83ec18                  sub esp, 18

' *** Reference to "__vbaExceptHandler"
'0071ad26    6866724000              push 00407266
'0071ad2b    64a100000000            mov ax, word ptr fs:[00000000]
'0071ad31    50                      push eax
    'Reference to '__except_list'
'0071ad32    64892500000000          mov dword ptr fs:[00000000], esp
'0071ad39    b8c0030000              mov eax, 000003c0

' *** Reference to "__vbaChkstk"
'0071ad3e    e81dc5ceff              call 407260
'0071ad43    53                      push ebx
'0071ad44    56                      push esi
'0071ad45    57                      push edi
'0071ad46    8965e8                  mov dword ptr [ebp-18], esp
'0071ad49    c745ec206e4000          mov dword ptr [ebp-14], 00406e20
'0071ad50    8b4508                  mov eax, dword ptr [ebp+08]
'0071ad53    83e001                  and eax, 01
'0071ad56    8945f0                  mov dword ptr [ebp-10], eax
'0071ad59    8b4d08                  mov ecx, dword ptr [ebp+08]
'0071ad5c    83e1fe                  and ecx, fffffffe
'0071ad5f    894d08                  mov dword ptr [ebp+08], ecx
'0071ad62    c745f400000000          mov dword ptr [ebp-0c], 00000000
'0071ad69    8b5508                  mov edx, dword ptr [ebp+08]
'0071ad6c    8b02                    mov eax, dword ptr [edx]
'0071ad6e    8b4d08                  mov ecx, dword ptr [ebp+08]
'0071ad71    51                      push ecx
'0071ad72    ff5004                  call dword ptr [eax+04]
'0071ad75    c745fc01000000          mov dword ptr [ebp-04], 00000001
'0071ad7c    c745fc02000000          mov dword ptr [ebp-04], 00000002
'0071ad83    0fbf1528b07200          movsx edx, word ptr [0072b028]
'0071ad8a    85d2                    test dx, dx
'0071ad8c    750f                    jne 71ad9d
'0071ad8e    c745fc03000000          mov dword ptr [ebp-04], 00000003
'0071ad95    6aff                    push ffffffff

' *** Reference to "__vbaOnError"
'0071ad97    ff15b0104000            call dword ptr [004010b0]
On Error Resume Next
'0071ad9d    c745fc05000000          mov dword ptr [ebp-04], 00000005
'0071ada4    833d24be720000          cmp dword ptr [0072be24], 00
'0071adab    751c                    jne 71adc9
'0071adad    6824be7200              push 0072be24
'0071adb2    6870004300              push 00430070

' *** Reference to "__vbaNew2"
'0071adb7    ff152c124000            call dword ptr [0040122c]
Dim var_2 As New Global
'0071adbd    c7854cfdffff24be7200    mov dword ptr [ebp+fffffd4c], 0072be24
'0071adc7    eb0a                    jmp 71add3
'0071adc9    c7854cfdffff24be7200    mov dword ptr [ebp+fffffd4c], 0072be24
'0071add3    8b854cfdffff            mov eax, dword ptr [ebp+fffffd4c]
'0071add9    8b08                    mov ecx, dword ptr [eax]
'0071addb    898de0fdffff            mov dword ptr [ebp+fffffde0], ecx
'0071ade1    833d5cb0720000          cmp dword ptr [0072b05c], 00
'0071ade8    751c                    jne 71ae06
'0071adea    685cb07200              push 0072b05c
'0071adef    68347b4000              push 00407b34

' *** Reference to "__vbaNew2"
'0071adf4    ff152c124000            call dword ptr [0040122c]
Dim var_24 As New frmAcceder
'0071adfa    c78548fdffff5cb07200    mov dword ptr [ebp+fffffd48], 0072b05c
'0071ae04    eb0a                    jmp 71ae10
'0071ae06    c78548fdffff5cb07200    mov dword ptr [ebp+fffffd48], 0072b05c
'0071ae10    8b9548fdffff            mov edx, dword ptr [ebp+fffffd48]
'0071ae16    8b02                    mov eax, dword ptr [edx]
'0071ae18    50                      push eax
'0071ae19    8d4da0                  lea ecx, dword ptr [ebp-60]
'0071ae1c    51                      push ecx

' *** Reference to "__vbaObjSetAddref"
'0071ae1d    ff15c8104000            call dword ptr [004010c8]
Set var_7 = var_24
'0071ae23    50                      push eax
'0071ae24    8b95e0fdffff            mov edx, dword ptr [ebp+fffffde0]
'0071ae2a    8b02                    mov eax, dword ptr [edx]
'0071ae2c    8b8de0fdffff            mov ecx, dword ptr [ebp+fffffde0]
'0071ae32    51                      push ecx
'0071ae33    ff5010                  call dword ptr [eax+10]
Call var_2.Unload(var_7)
'0071ae36    dbe2                    fnclex
'0071ae38    8985dcfdffff            mov dword ptr [ebp+fffffddc], eax
'0071ae3e    83bddcfdffff00          cmp dword ptr [ebp+fffffddc], 00
'0071ae45    7d23                    jge 71ae6a

If (0 < 0) Then
'0071ae47    6a10                    push 10
'0071ae49    6860004300              push 00430060
'0071ae4e    8b95e0fdffff            mov edx, dword ptr [ebp+fffffde0]
'0071ae54    52                      push edx
'0071ae55    8b85dcfdffff            mov eax, dword ptr [ebp+fffffddc]
'0071ae5b    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0071ae5c    ff1580104000            call dword ptr [00401080]
'0071ae62    898544fdffff            mov dword ptr [ebp+fffffd44], eax
'0071ae68    eb0a                    jmp 71ae74
    
End If
'0071ae6a    c78544fdffff00000000    mov dword ptr [ebp+fffffd44], 00000000
'0071ae74    8d4da0                  lea ecx, dword ptr [ebp-60]

' *** Reference to "__vbaFreeObj"
'0071ae77    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_7)
'0071ae7d    c745fc06000000          mov dword ptr [ebp-04], 00000006
'0071ae84    8b4d08                  mov ecx, dword ptr [ebp+08]
'0071ae87    83c138                  add ecx, 38
var_num3 = Me + 56
'0071ae8a    51                      push ecx
'0071ae8b    e8400edcff              call 4dbcd0
Call sub_4DBCD0()
'0071ae90    c745fc07000000          mov dword ptr [ebp-04], 00000007
'0071ae97    c785a0feffff01000000    mov dword ptr [ebp+fffffea0], 00000001
'0071aea1    c78598feffff03000000    mov dword ptr [ebp+fffffe98], 00000003
'0071aeab    b810000000              mov eax, 00000010

' *** Reference to "__vbaChkstk"
'0071aeb0    e8abc3ceff              call 407260
'0071aeb5    8bd4                    mov edx, esp
'0071aeb7    8b8598feffff            mov eax, dword ptr [ebp+fffffe98]
'0071aebd    8902                    mov dword ptr [edx], eax
'0071aebf    8b8d9cfeffff            mov ecx, dword ptr [ebp+fffffe9c]
'0071aec5    894a04                  mov dword ptr [edx+04], ecx
'0071aec8    8b85a0feffff            mov eax, dword ptr [ebp+fffffea0]
'0071aece    894208                  mov dword ptr [edx+08], eax
'0071aed1    8b8da4feffff            mov ecx, dword ptr [ebp+fffffea4]
'0071aed7    894a0c                  mov dword ptr [edx+0c], ecx
'0071aeda    6a07                    push 07
'0071aedc    8b5508                  mov edx, dword ptr [ebp+08]
'0071aedf    8b02                    mov eax, dword ptr [edx]
'0071aee1    8b4d08                  mov ecx, dword ptr [ebp+08]
'0071aee4    51                      push ecx
'0071aee5    ff9010030000            call dword ptr [eax+00000310]
'0071aeeb    50                      push eax
'0071aeec    8d55a0                  lea edx, dword ptr [ebp-60]
'0071aeef    52                      push edx

' *** Reference to "__vbaObjSet"
'0071aef0    ff15b4104000            call dword ptr [004010b4]
Set var_7 = Nothing
'0071aef6    50                      push eax

' *** Reference to "__vbaLateIdSt"
'0071aef7    ff15ec124000            call dword ptr [004012ec]
var_7.[UNMANAGED] = var_527
'0071aefd    8d4da0                  lea ecx, dword ptr [ebp-60]

' *** Reference to "__vbaFreeObj"
'0071af00    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_7)
'0071af06    c745fc08000000          mov dword ptr [ebp-04], 00000008
'0071af0d    c785a0feffff01000000    mov dword ptr [ebp+fffffea0], 00000001
'0071af17    c78598feffff03000000    mov dword ptr [ebp+fffffe98], 00000003
'0071af21    b810000000              mov eax, 00000010

' *** Reference to "__vbaChkstk"
'0071af26    e835c3ceff              call 407260
'0071af2b    8bc4                    mov eax, esp
'0071af2d    8b8d98feffff            mov ecx, dword ptr [ebp+fffffe98]
'0071af33    8908                    mov dword ptr [eax], ecx
'0071af35    8b959cfeffff            mov edx, dword ptr [ebp+fffffe9c]
'0071af3b    895004                  mov dword ptr [eax+04], edx
'0071af3e    8b8da0feffff            mov ecx, dword ptr [ebp+fffffea0]
'0071af44    894808                  mov dword ptr [eax+08], ecx
'0071af47    8b95a4feffff            mov edx, dword ptr [ebp+fffffea4]
'0071af4d    89500c                  mov dword ptr [eax+0c], edx
'0071af50    6a12                    push 12
'0071af52    8b4508                  mov eax, dword ptr [ebp+08]
'0071af55    8b08                    mov ecx, dword ptr [eax]
'0071af57    8b5508                  mov edx, dword ptr [ebp+08]
'0071af5a    52                      push edx
'0071af5b    ff9110030000            call dword ptr [ecx+00000310]
'0071af61    50                      push eax
'0071af62    8d45a0                  lea eax, dword ptr [ebp-60]
'0071af65    50                      push eax

' *** Reference to "__vbaObjSet"
'0071af66    ff15b4104000            call dword ptr [004010b4]
Set var_7 = Me
'0071af6c    50                      push eax

' *** Reference to "__vbaLateIdSt"
'0071af6d    ff15ec124000            call dword ptr [004012ec]
var_7.[UNMANAGED] = var_527
'0071af73    8d4da0                  lea ecx, dword ptr [ebp-60]

' *** Reference to "__vbaFreeObj"
'0071af76    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_7)
'0071af7c    c745fc09000000          mov dword ptr [ebp-04], 00000009
'0071af83    c78580feffff04000280    mov dword ptr [ebp+fffffe80], 80020004
'0071af8d    c78578feffff0a000000    mov dword ptr [ebp+fffffe78], 0000000a
'0071af97    c78590feffff04000280    mov dword ptr [ebp+fffffe90], 80020004
'0071afa1    c78588feffff0a000000    mov dword ptr [ebp+fffffe88], 0000000a
'0071afab    c785a0feffff02000000    mov dword ptr [ebp+fffffea0], 00000002
'0071afb5    c78598feffff03000000    mov dword ptr [ebp+fffffe98], 00000003
'0071afbf    8d4da0                  lea ecx, dword ptr [ebp-60]
'0071afc2    51                      push ecx
'0071afc3    b810000000              mov eax, 00000010

' *** Reference to "__vbaChkstk"
'0071afc8    e893c2ceff              call 407260
'0071afcd    8bd4                    mov edx, esp
'0071afcf    8b8578feffff            mov eax, dword ptr [ebp+fffffe78]
'0071afd5    8902                    mov dword ptr [edx], eax
'0071afd7    8b8d7cfeffff            mov ecx, dword ptr [ebp+fffffe7c]
'0071afdd    894a04                  mov dword ptr [edx+04], ecx
'0071afe0    8b8580feffff            mov eax, dword ptr [ebp+fffffe80]
'0071afe6    894208                  mov dword ptr [edx+08], eax
'0071afe9    8b8d84feffff            mov ecx, dword ptr [ebp+fffffe84]
'0071afef    894a0c                  mov dword ptr [edx+0c], ecx
'0071aff2    b810000000              mov eax, 00000010

' *** Reference to "__vbaChkstk"
'0071aff7    e864c2ceff              call 407260
'0071affc    8bd4                    mov edx, esp
'0071affe    8b8588feffff            mov eax, dword ptr [ebp+fffffe88]
'0071b004    8902                    mov dword ptr [edx], eax
'0071b006    8b8d8cfeffff            mov ecx, dword ptr [ebp+fffffe8c]
'0071b00c    894a04                  mov dword ptr [edx+04], ecx
'0071b00f    8b8590feffff            mov eax, dword ptr [ebp+fffffe90]
'0071b015    894208                  mov dword ptr [edx+08], eax
'0071b018    8b8d94feffff            mov ecx, dword ptr [ebp+fffffe94]
'0071b01e    894a0c                  mov dword ptr [edx+0c], ecx
'0071b021    b810000000              mov eax, 00000010

' *** Reference to "__vbaChkstk"
'0071b026    e835c2ceff              call 407260
'0071b02b    8bd4                    mov edx, esp
'0071b02d    8b8598feffff            mov eax, dword ptr [ebp+fffffe98]
'0071b033    8902                    mov dword ptr [edx], eax
'0071b035    8b8d9cfeffff            mov ecx, dword ptr [ebp+fffffe9c]
'0071b03b    894a04                  mov dword ptr [edx+04], ecx
'0071b03e    8b85a0feffff            mov eax, dword ptr [ebp+fffffea0]
'0071b044    894208                  mov dword ptr [edx+08], eax
'0071b047    8b8da4feffff            mov ecx, dword ptr [ebp+fffffea4]
'0071b04d    894a0c                  mov dword ptr [edx+0c], ecx
'0071b050    68d8db4400              push 0044dbd8
'0071b055    8b5508                  mov edx, dword ptr [ebp+08]
'0071b058    8b4238                  mov eax, dword ptr [edx+38]
'0071b05b    8b4d08                  mov ecx, dword ptr [ebp+08]
'0071b05e    8b5138                  mov edx, dword ptr [ecx+38]
'0071b061    8b0a                    mov ecx, dword ptr [edx]
'0071b063    50                      push eax
'0071b064    ff91bc000000            call dword ptr [ecx+000000bc]
'0071b06a    dbe2                    fnclex
'0071b06c    8985e0fdffff            mov dword ptr [ebp+fffffde0], eax
'0071b072    83bde0fdffff00          cmp dword ptr [ebp+fffffde0], 00
'0071b079    7d26                    jge 71b0a1

If (arg_6 < 0) Then
'0071b07b    68bc000000              push 000000bc
'0071b080    68301f4300              push 00431f30
'0071b085    8b5508                  mov edx, dword ptr [ebp+08]
'0071b088    8b4238                  mov eax, dword ptr [edx+38]
'0071b08b    50                      push eax
'0071b08c    8b8de0fdffff            mov ecx, dword ptr [ebp+fffffde0]
'0071b092    51                      push ecx

' *** Reference to "__vbaHresultCheckObj"
'0071b093    ff1580104000            call dword ptr [00401080]
'0071b099    898540fdffff            mov dword ptr [ebp+fffffd40], eax
'0071b09f    eb0a                    jmp 71b0ab
    
End If
'0071b0a1    c78540fdffff00000000    mov dword ptr [ebp+fffffd40], 00000000
'0071b0ab    8b55a0                  mov edx, dword ptr [ebp-60]
'0071b0ae    899590fdffff            mov dword ptr [ebp+fffffd90], edx
'0071b0b4    c745a000000000          mov dword ptr [ebp-60], 00000000
'0071b0bb    8b8590fdffff            mov eax, dword ptr [ebp+fffffd90]
'0071b0c1    50                      push eax
'0071b0c2    8d4dac                  lea ecx, dword ptr [ebp-54]
'0071b0c5    51                      push ecx

' *** Reference to "__vbaObjSet"
'0071b0c6    ff15b4104000            call dword ptr [004010b4]
Set var_50 = var_7
'0071b0cc    c745fc0a000000          mov dword ptr [ebp-04], 0000000a
'0071b0d3    8d95e4fdffff            lea edx, dword ptr [ebp+fffffde4]
'0071b0d9    52                      push edx
'0071b0da    8b45ac                  mov eax, dword ptr [ebp-54]
'0071b0dd    8b08                    mov ecx, dword ptr [eax]
'0071b0df    8b55ac                  mov edx, dword ptr [ebp-54]
'0071b0e2    52                      push edx
'0071b0e3    ff5134                  call dword ptr [ecx+34]
'0071b0e6    dbe2                    fnclex
'0071b0e8    8985e0fdffff            mov dword ptr [ebp+fffffde0], eax
'0071b0ee    83bde0fdffff00          cmp dword ptr [ebp+fffffde0], 00
'0071b0f5    7d20                    jge 71b117

If (var_50 < 0) Then
'0071b0f7    6a34                    push 34
'0071b0f9    6830314300              push 00433130
'0071b0fe    8b45ac                  mov eax, dword ptr [ebp-54]
'0071b101    50                      push eax
'0071b102    8b8de0fdffff            mov ecx, dword ptr [ebp+fffffde0]
'0071b108    51                      push ecx

' *** Reference to "__vbaHresultCheckObj"
'0071b109    ff1580104000            call dword ptr [00401080]
'0071b10f    89853cfdffff            mov dword ptr [ebp+fffffd3c], eax
'0071b115    eb0a                    jmp 71b121
    
End If
'0071b117    c7853cfdffff00000000    mov dword ptr [ebp+fffffd3c], 00000000
'0071b121    0fbf95e4fdffff          movsx edx, word ptr [ebp+fffffde4]
'0071b128    85d2                    test dx, dx
'0071b12a    0f854e2d0000            jne 71de7e
'0071b130    c745fc0b000000          mov dword ptr [ebp-04], 0000000b
'0071b137    8d45a0                  lea eax, dword ptr [ebp-60]
'0071b13a    50                      push eax
'0071b13b    8b4dac                  mov ecx, dword ptr [ebp-54]
'0071b13e    8b11                    mov edx, dword ptr [ecx]
'0071b140    8b45ac                  mov eax, dword ptr [ebp-54]
'0071b143    50                      push eax
'0071b144    ff92b4000000            call dword ptr [edx+000000b4]
'0071b14a    dbe2                    fnclex
'0071b14c    8985e0fdffff            mov dword ptr [ebp+fffffde0], eax
'0071b152    83bde0fdffff00          cmp dword ptr [ebp+fffffde0], 00
'0071b159    7d23                    jge 71b17e

If (var_50 < 0) Then
'0071b15b    68b4000000              push 000000b4
'0071b160    6830314300              push 00433130
'0071b165    8b4dac                  mov ecx, dword ptr [ebp-54]
'0071b168    51                      push ecx
'0071b169    8b95e0fdffff            mov edx, dword ptr [ebp+fffffde0]
'0071b16f    52                      push edx

' *** Reference to "__vbaHresultCheckObj"
'0071b170    ff1580104000            call dword ptr [00401080]
'0071b176    898538fdffff            mov dword ptr [ebp+fffffd38], eax
'0071b17c    eb0a                    jmp 71b188
    
End If
'0071b17e    c78538fdffff00000000    mov dword ptr [ebp+fffffd38], 00000000
'0071b188    8b45a0                  mov eax, dword ptr [ebp-60]
'0071b18b    8985dcfdffff            mov dword ptr [ebp+fffffddc], eax
'0071b191    c785a0feffff34d44300    mov dword ptr [ebp+fffffea0], 0043d434
'0071b19b    c78598feffff08000000    mov dword ptr [ebp+fffffe98], 00000008
'0071b1a5    8d4d9c                  lea ecx, dword ptr [ebp-64]
'0071b1a8    51                      push ecx
'0071b1a9    b810000000              mov eax, 00000010

' *** Reference to "__vbaChkstk"
'0071b1ae    e8adc0ceff              call 407260
'0071b1b3    8bd4                    mov edx, esp
'0071b1b5    8b8598feffff            mov eax, dword ptr [ebp+fffffe98]
'0071b1bb    8902                    mov dword ptr [edx], eax
'0071b1bd    8b8d9cfeffff            mov ecx, dword ptr [ebp+fffffe9c]
'0071b1c3    894a04                  mov dword ptr [edx+04], ecx
'0071b1c6    8b85a0feffff            mov eax, dword ptr [ebp+fffffea0]
'0071b1cc    894208                  mov dword ptr [edx+08], eax
'0071b1cf    8b8da4feffff            mov ecx, dword ptr [ebp+fffffea4]
'0071b1d5    894a0c                  mov dword ptr [edx+0c], ecx
'0071b1d8    8b95dcfdffff            mov edx, dword ptr [ebp+fffffddc]
'0071b1de    8b02                    mov eax, dword ptr [edx]
'0071b1e0    8b8ddcfdffff            mov ecx, dword ptr [ebp+fffffddc]
'0071b1e6    51                      push ecx
'0071b1e7    ff5030                  call dword ptr [eax+30]
'0071b1ea    dbe2                    fnclex
'0071b1ec    8985d8fdffff            mov dword ptr [ebp+fffffdd8], eax
'0071b1f2    83bdd8fdffff00          cmp dword ptr [ebp+fffffdd8], 00
'0071b1f9    7d23                    jge 71b21e

If (-256 - 24 < 0) Then
'0071b1fb    6a30                    push 30
'0071b1fd    68d8304300              push 004330d8
'0071b202    8b95dcfdffff            mov edx, dword ptr [ebp+fffffddc]
'0071b208    52                      push edx
'0071b209    8b85d8fdffff            mov eax, dword ptr [ebp+fffffdd8]
'0071b20f    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0071b210    ff1580104000            call dword ptr [00401080]
'0071b216    898534fdffff            mov dword ptr [ebp+fffffd34], eax
'0071b21c    eb0a                    jmp 71b228
    
End If
'0071b21e    c78534fdffff00000000    mov dword ptr [ebp+fffffd34], 00000000
'0071b228    8b4d9c                  mov ecx, dword ptr [ebp-64]
'0071b22b    898d8cfdffff            mov dword ptr [ebp+fffffd8c], ecx
'0071b231    c7459c00000000          mov dword ptr [ebp-64], 00000000
'0071b238    8b958cfdffff            mov edx, dword ptr [ebp+fffffd8c]
'0071b23e    895580                  mov dword ptr [ebp-80], edx
'0071b241    c78578ffffff09000000    mov dword ptr [ebp+ffffff78], 00000009
'0071b24b    8d8578ffffff            lea eax, dword ptr [ebp+ffffff78]
'0071b251    50                      push eax

' *** Reference to "rtcIsNull"
'0071b252    ff1540114000            call dword ptr [00401140]
'0071b258    668985e4fdffff          mov word ptr [ebp+fffffde4], ax
'0071b25f    8d4d98                  lea ecx, dword ptr [ebp-68]
'0071b262    51                      push ecx
'0071b263    8b55ac                  mov edx, dword ptr [ebp-54]
'0071b266    8b02                    mov eax, dword ptr [edx]
'0071b268    8b4dac                  mov ecx, dword ptr [ebp-54]
'0071b26b    51                      push ecx
'0071b26c    ff90b4000000            call dword ptr [eax+000000b4]
'0071b272    dbe2                    fnclex
'0071b274    8985d4fdffff            mov dword ptr [ebp+fffffdd4], eax
'0071b27a    83bdd4fdffff00          cmp dword ptr [ebp+fffffdd4], 00
'0071b281    7d23                    jge 71b2a6

If (frmImport < 0) Then
'0071b283    68b4000000              push 000000b4
'0071b288    6830314300              push 00433130
'0071b28d    8b55ac                  mov edx, dword ptr [ebp-54]
'0071b290    52                      push edx
'0071b291    8b85d4fdffff            mov eax, dword ptr [ebp+fffffdd4]
'0071b297    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0071b298    ff1580104000            call dword ptr [00401080]
'0071b29e    898530fdffff            mov dword ptr [ebp+fffffd30], eax
'0071b2a4    eb0a                    jmp 71b2b0
    
End If
'0071b2a6    c78530fdffff00000000    mov dword ptr [ebp+fffffd30], 00000000
'0071b2b0    8b4d98                  mov ecx, dword ptr [ebp-68]
'0071b2b3    898dd0fdffff            mov dword ptr [ebp+fffffdd0], ecx
'0071b2b9    c78590feffff34d44300    mov dword ptr [ebp+fffffe90], 0043d434
'0071b2c3    c78588feffff08000000    mov dword ptr [ebp+fffffe88], 00000008
'0071b2cd    8d5594                  lea edx, dword ptr [ebp-6c]
'0071b2d0    52                      push edx
'0071b2d1    b810000000              mov eax, 00000010

' *** Reference to "__vbaChkstk"
'0071b2d6    e885bfceff              call 407260
'0071b2db    8bc4                    mov eax, esp
'0071b2dd    8b8d88feffff            mov ecx, dword ptr [ebp+fffffe88]
'0071b2e3    8908                    mov dword ptr [eax], ecx
'0071b2e5    8b958cfeffff            mov edx, dword ptr [ebp+fffffe8c]
'0071b2eb    895004                  mov dword ptr [eax+04], edx
'0071b2ee    8b8d90feffff            mov ecx, dword ptr [ebp+fffffe90]
'0071b2f4    894808                  mov dword ptr [eax+08], ecx
'0071b2f7    8b9594feffff            mov edx, dword ptr [ebp+fffffe94]
'0071b2fd    89500c                  mov dword ptr [eax+0c], edx
'0071b300    8b85d0fdffff            mov eax, dword ptr [ebp+fffffdd0]
'0071b306    8b08                    mov ecx, dword ptr [eax]
'0071b308    8b95d0fdffff            mov edx, dword ptr [ebp+fffffdd0]
'0071b30e    52                      push edx
'0071b30f    ff5130                  call dword ptr [ecx+30]
'0071b312    dbe2                    fnclex
'0071b314    8985ccfdffff            mov dword ptr [ebp+fffffdcc], eax
'0071b31a    83bdccfdffff00          cmp dword ptr [ebp+fffffdcc], 00
'0071b321    7d23                    jge 71b346

If (0 < 0) Then
'0071b323    6a30                    push 30
'0071b325    68d8304300              push 004330d8
'0071b32a    8b85d0fdffff            mov eax, dword ptr [ebp+fffffdd0]
'0071b330    50                      push eax
'0071b331    8b8dccfdffff            mov ecx, dword ptr [ebp+fffffdcc]
'0071b337    51                      push ecx

' *** Reference to "__vbaHresultCheckObj"
'0071b338    ff1580104000            call dword ptr [00401080]
'0071b33e    89852cfdffff            mov dword ptr [ebp+fffffd2c], eax
'0071b344    eb0a                    jmp 71b350
    
End If
'0071b346    c7852cfdffff00000000    mov dword ptr [ebp+fffffd2c], 00000000
'0071b350    8b5594                  mov edx, dword ptr [ebp-6c]
'0071b353    899588fdffff            mov dword ptr [ebp+fffffd88], edx
'0071b359    c7459400000000          mov dword ptr [ebp-6c], 00000000
'0071b360    8b8588fdffff            mov eax, dword ptr [ebp+fffffd88]
'0071b366    898560ffffff            mov dword ptr [ebp+ffffff60], eax
'0071b36c    c78558ffffff09000000    mov dword ptr [ebp+ffffff58], 00000009
'0071b376    c78570ffffff00000000    mov dword ptr [ebp+ffffff70], 00000000
'0071b380    c78568ffffff02000000    mov dword ptr [ebp+ffffff68], 00000002
'0071b38a    668b8de4fdffff          mov cx, word ptr [ebp+fffffde4]
'0071b391    66898d80feffff          mov word ptr [ebp+fffffe80], cx
'0071b398    c78578feffff0b000000    mov dword ptr [ebp+fffffe78], 0000000b
'0071b3a2    8d9558ffffff            lea edx, dword ptr [ebp+ffffff58]
'0071b3a8    52                      push edx
'0071b3a9    8d8568ffffff            lea eax, dword ptr [ebp+ffffff68]
'0071b3af    50                      push eax
'0071b3b0    8d8d78feffff            lea ecx, dword ptr [ebp+fffffe78]
'0071b3b6    51                      push ecx
'0071b3b7    8d9548ffffff            lea edx, dword ptr [ebp+ffffff48]
'0071b3bd    52                      push edx

' *** Reference to "rtcImmediateIf"
'0071b3be    ff154c124000            call dword ptr [0040124c]
'0071b3c4    8d8548ffffff            lea eax, dword ptr [ebp+ffffff48]
'0071b3ca    50                      push eax

' *** Reference to "__vbaI2Var"
'0071b3cb    ff150c124000            call dword ptr [0040120c]
'0071b3d1    668945a8                mov word ptr [ebp-58], ax
'0071b3d5    8d4d98                  lea ecx, dword ptr [ebp-68]
'0071b3d8    51                      push ecx
'0071b3d9    8d55a0                  lea edx, dword ptr [ebp-60]
'0071b3dc    52                      push edx
'0071b3dd    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'0071b3df    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_7, var_130)
'0071b3e5    83c40c                  add esp, 0c
'0071b3e8    8d8548ffffff            lea eax, dword ptr [ebp+ffffff48]
'0071b3ee    50                      push eax
'0071b3ef    8d8d58ffffff            lea ecx, dword ptr [ebp+ffffff58]
'0071b3f5    51                      push ecx
'0071b3f6    8d9568ffffff            lea edx, dword ptr [ebp+ffffff68]
'0071b3fc    52                      push edx
'0071b3fd    8d8578feffff            lea eax, dword ptr [ebp+fffffe78]
'0071b403    50                      push eax
'0071b404    8d8d78ffffff            lea ecx, dword ptr [ebp+ffffff78]
'0071b40a    51                      push ecx
'0071b40b    6a05                    push 05

' *** Reference to "__vbaFreeVarList"
'0071b40d    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_87, var_73, var_132, var_133, var_134)
'0071b413    83c418                  add esp, 18
'0071b416    c745fc0c000000          mov dword ptr [ebp-04], 0000000c
'0071b41d    8d55a0                  lea edx, dword ptr [ebp-60]
'0071b420    52                      push edx
'0071b421    8b45ac                  mov eax, dword ptr [ebp-54]
'0071b424    8b08                    mov ecx, dword ptr [eax]
'0071b426    8b55ac                  mov edx, dword ptr [ebp-54]
'0071b429    52                      push edx
'0071b42a    ff91b4000000            call dword ptr [ecx+000000b4]
'0071b430    dbe2                    fnclex
'0071b432    8985e0fdffff            mov dword ptr [ebp+fffffde0], eax
'0071b438    83bde0fdffff00          cmp dword ptr [ebp+fffffde0], 00
'0071b43f    7d23                    jge 71b464

If (var_50 < 0) Then
'0071b441    68b4000000              push 000000b4
'0071b446    6830314300              push 00433130
'0071b44b    8b45ac                  mov eax, dword ptr [ebp-54]
'0071b44e    50                      push eax
'0071b44f    8b8de0fdffff            mov ecx, dword ptr [ebp+fffffde0]
'0071b455    51                      push ecx

' *** Reference to "__vbaHresultCheckObj"
'0071b456    ff1580104000            call dword ptr [00401080]
'0071b45c    898528fdffff            mov dword ptr [ebp+fffffd28], eax
'0071b462    eb0a                    jmp 71b46e
    
End If
'0071b464    c78528fdffff00000000    mov dword ptr [ebp+fffffd28], 00000000
'0071b46e    8b55a0                  mov edx, dword ptr [ebp-60]
'0071b471    8995dcfdffff            mov dword ptr [ebp+fffffddc], edx
'0071b477    c785a0feffff5cd44300    mov dword ptr [ebp+fffffea0], 0043d45c
'0071b481    c78598feffff08000000    mov dword ptr [ebp+fffffe98], 00000008
'0071b48b    8d459c                  lea eax, dword ptr [ebp-64]
'0071b48e    50                      push eax
'0071b48f    b810000000              mov eax, 00000010

' *** Reference to "__vbaChkstk"
'0071b494    e8c7bdceff              call 407260
'0071b499    8bcc                    mov ecx, esp
'0071b49b    8b9598feffff            mov edx, dword ptr [ebp+fffffe98]
'0071b4a1    8911                    mov dword ptr [ecx], edx
'0071b4a3    8b859cfeffff            mov eax, dword ptr [ebp+fffffe9c]
'0071b4a9    894104                  mov dword ptr [ecx+04], eax
'0071b4ac    8b95a0feffff            mov edx, dword ptr [ebp+fffffea0]
'0071b4b2    895108                  mov dword ptr [ecx+08], edx
'0071b4b5    8b85a4feffff            mov eax, dword ptr [ebp+fffffea4]
'0071b4bb    89410c                  mov dword ptr [ecx+0c], eax
'0071b4be    8b8ddcfdffff            mov ecx, dword ptr [ebp+fffffddc]
'0071b4c4    8b11                    mov edx, dword ptr [ecx]
'0071b4c6    8b85dcfdffff            mov eax, dword ptr [ebp+fffffddc]
'0071b4cc    50                      push eax
'0071b4cd    ff5230                  call dword ptr [edx+30]
'0071b4d0    dbe2                    fnclex
'0071b4d2    8985d8fdffff            mov dword ptr [ebp+fffffdd8], eax
'0071b4d8    83bdd8fdffff00          cmp dword ptr [ebp+fffffdd8], 00
'0071b4df    7d23                    jge 71b504

If (var_7 < 0) Then
'0071b4e1    6a30                    push 30
'0071b4e3    68d8304300              push 004330d8
'0071b4e8    8b8ddcfdffff            mov ecx, dword ptr [ebp+fffffddc]
'0071b4ee    51                      push ecx
'0071b4ef    8b95d8fdffff            mov edx, dword ptr [ebp+fffffdd8]
'0071b4f5    52                      push edx

' *** Reference to "__vbaHresultCheckObj"
'0071b4f6    ff1580104000            call dword ptr [00401080]
'0071b4fc    898524fdffff            mov dword ptr [ebp+fffffd24], eax
'0071b502    eb0a                    jmp 71b50e
    
End If
'0071b504    c78524fdffff00000000    mov dword ptr [ebp+fffffd24], 00000000
'0071b50e    8b459c                  mov eax, dword ptr [ebp-64]
'0071b511    898584fdffff            mov dword ptr [ebp+fffffd84], eax
'0071b517    c7459c00000000          mov dword ptr [ebp-64], 00000000
'0071b51e    8b8d84fdffff            mov ecx, dword ptr [ebp+fffffd84]
'0071b524    894d80                  mov dword ptr [ebp-80], ecx
'0071b527    c78578ffffff09000000    mov dword ptr [ebp+ffffff78], 00000009
'0071b531    8d9578ffffff            lea edx, dword ptr [ebp+ffffff78]
'0071b537    52                      push edx

' *** Reference to "rtcIsNull"
'0071b538    ff1540114000            call dword ptr [00401140]
'0071b53e    668985e4fdffff          mov word ptr [ebp+fffffde4], ax
'0071b545    8d4598                  lea eax, dword ptr [ebp-68]
'0071b548    50                      push eax
'0071b549    8b4dac                  mov ecx, dword ptr [ebp-54]
'0071b54c    8b11                    mov edx, dword ptr [ecx]
'0071b54e    8b45ac                  mov eax, dword ptr [ebp-54]
'0071b551    50                      push eax
'0071b552    ff92b4000000            call dword ptr [edx+000000b4]
'0071b558    dbe2                    fnclex
'0071b55a    8985d4fdffff            mov dword ptr [ebp+fffffdd4], eax
'0071b560    83bdd4fdffff00          cmp dword ptr [ebp+fffffdd4], 00
'0071b567    7d23                    jge 71b58c

If (var_50 < 0) Then
'0071b569    68b4000000              push 000000b4
'0071b56e    6830314300              push 00433130
'0071b573    8b4dac                  mov ecx, dword ptr [ebp-54]
'0071b576    51                      push ecx
'0071b577    8b95d4fdffff            mov edx, dword ptr [ebp+fffffdd4]
'0071b57d    52                      push edx

' *** Reference to "__vbaHresultCheckObj"
'0071b57e    ff1580104000            call dword ptr [00401080]
'0071b584    898520fdffff            mov dword ptr [ebp+fffffd20], eax
'0071b58a    eb0a                    jmp 71b596
    
End If
'0071b58c    c78520fdffff00000000    mov dword ptr [ebp+fffffd20], 00000000
'0071b596    8b4598                  mov eax, dword ptr [ebp-68]
'0071b599    8985d0fdffff            mov dword ptr [ebp+fffffdd0], eax
'0071b59f    c78590feffff5cd44300    mov dword ptr [ebp+fffffe90], 0043d45c
'0071b5a9    c78588feffff08000000    mov dword ptr [ebp+fffffe88], 00000008
'0071b5b3    8d4d94                  lea ecx, dword ptr [ebp-6c]
'0071b5b6    51                      push ecx
'0071b5b7    b810000000              mov eax, 00000010

' *** Reference to "__vbaChkstk"
'0071b5bc    e89fbcceff              call 407260
'0071b5c1    8bd4                    mov edx, esp
'0071b5c3    8b8588feffff            mov eax, dword ptr [ebp+fffffe88]
'0071b5c9    8902                    mov dword ptr [edx], eax
'0071b5cb    8b8d8cfeffff            mov ecx, dword ptr [ebp+fffffe8c]
'0071b5d1    894a04                  mov dword ptr [edx+04], ecx
'0071b5d4    8b8590feffff            mov eax, dword ptr [ebp+fffffe90]
'0071b5da    894208                  mov dword ptr [edx+08], eax
'0071b5dd    8b8d94feffff            mov ecx, dword ptr [ebp+fffffe94]
'0071b5e3    894a0c                  mov dword ptr [edx+0c], ecx
'0071b5e6    8b95d0fdffff            mov edx, dword ptr [ebp+fffffdd0]
'0071b5ec    8b02                    mov eax, dword ptr [edx]
'0071b5ee    8b8dd0fdffff            mov ecx, dword ptr [ebp+fffffdd0]
'0071b5f4    51                      push ecx
'0071b5f5    ff5030                  call dword ptr [eax+30]
'0071b5f8    dbe2                    fnclex
'0071b5fa    8985ccfdffff            mov dword ptr [ebp+fffffdcc], eax
'0071b600    83bdccfdffff00          cmp dword ptr [ebp+fffffdcc], 00
'0071b607    7d23                    jge 71b62c

If (-256 - 24 < 0) Then
'0071b609    6a30                    push 30
'0071b60b    68d8304300              push 004330d8
'0071b610    8b95d0fdffff            mov edx, dword ptr [ebp+fffffdd0]
'0071b616    52                      push edx
'0071b617    8b85ccfdffff            mov eax, dword ptr [ebp+fffffdcc]
'0071b61d    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0071b61e    ff1580104000            call dword ptr [00401080]
'0071b624    89851cfdffff            mov dword ptr [ebp+fffffd1c], eax
'0071b62a    eb0a                    jmp 71b636
    
End If
'0071b62c    c7851cfdffff00000000    mov dword ptr [ebp+fffffd1c], 00000000
'0071b636    8b4d94                  mov ecx, dword ptr [ebp-6c]
'0071b639    898d80fdffff            mov dword ptr [ebp+fffffd80], ecx
'0071b63f    c7459400000000          mov dword ptr [ebp-6c], 00000000
'0071b646    8b9580fdffff            mov edx, dword ptr [ebp+fffffd80]
'0071b64c    899560ffffff            mov dword ptr [ebp+ffffff60], edx
'0071b652    c78558ffffff09000000    mov dword ptr [ebp+ffffff58], 00000009
'0071b65c    c78570ffffff00000000    mov dword ptr [ebp+ffffff70], 00000000
'0071b666    c78568ffffff02000000    mov dword ptr [ebp+ffffff68], 00000002
'0071b670    668b85e4fdffff          mov ax, word ptr [ebp+fffffde4]
'0071b677    66898580feffff          mov word ptr [ebp+fffffe80], ax
'0071b67e    c78578feffff0b000000    mov dword ptr [ebp+fffffe78], 0000000b
'0071b688    8d8d58ffffff            lea ecx, dword ptr [ebp+ffffff58]
'0071b68e    51                      push ecx
'0071b68f    8d9568ffffff            lea edx, dword ptr [ebp+ffffff68]
'0071b695    52                      push edx
'0071b696    8d8578feffff            lea eax, dword ptr [ebp+fffffe78]
'0071b69c    50                      push eax
'0071b69d    8d8d48ffffff            lea ecx, dword ptr [ebp+ffffff48]
'0071b6a3    51                      push ecx

' *** Reference to "rtcImmediateIf"
'0071b6a4    ff154c124000            call dword ptr [0040124c]
'0071b6aa    8d9548ffffff            lea edx, dword ptr [ebp+ffffff48]
'0071b6b0    52                      push edx

' *** Reference to "__vbaI2Var"
'0071b6b1    ff150c124000            call dword ptr [0040120c]
'0071b6b7    668945a4                mov word ptr [ebp-5c], ax
'0071b6bb    8d4598                  lea eax, dword ptr [ebp-68]
'0071b6be    50                      push eax
'0071b6bf    8d4da0                  lea ecx, dword ptr [ebp-60]
'0071b6c2    51                      push ecx
'0071b6c3    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'0071b6c5    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_7, var_130)
'0071b6cb    83c40c                  add esp, 0c
'0071b6ce    8d9548ffffff            lea edx, dword ptr [ebp+ffffff48]
'0071b6d4    52                      push edx
'0071b6d5    8d8558ffffff            lea eax, dword ptr [ebp+ffffff58]
'0071b6db    50                      push eax
'0071b6dc    8d8d68ffffff            lea ecx, dword ptr [ebp+ffffff68]
'0071b6e2    51                      push ecx
'0071b6e3    8d9578feffff            lea edx, dword ptr [ebp+fffffe78]
'0071b6e9    52                      push edx
'0071b6ea    8d8578ffffff            lea eax, dword ptr [ebp+ffffff78]
'0071b6f0    50                      push eax
'0071b6f1    6a05                    push 05

' *** Reference to "__vbaFreeVarList"
'0071b6f3    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_87, var_73, var_132, var_133, var_134)
'0071b6f9    83c418                  add esp, 18
'0071b6fc    c745fc0d000000          mov dword ptr [ebp-04], 0000000d
'0071b703    8d4da0                  lea ecx, dword ptr [ebp-60]
'0071b706    51                      push ecx
'0071b707    8b55ac                  mov edx, dword ptr [ebp-54]
'0071b70a    8b02                    mov eax, dword ptr [edx]
'0071b70c    8b4dac                  mov ecx, dword ptr [ebp-54]
'0071b70f    51                      push ecx
'0071b710    ff90b4000000            call dword ptr [eax+000000b4]
'0071b716    dbe2                    fnclex
'0071b718    8985e0fdffff            mov dword ptr [ebp+fffffde0], eax
'0071b71e    83bde0fdffff00          cmp dword ptr [ebp+fffffde0], 00
'0071b725    7d23                    jge 71b74a

If (frmImport < 0) Then
'0071b727    68b4000000              push 000000b4
'0071b72c    6830314300              push 00433130
'0071b731    8b55ac                  mov edx, dword ptr [ebp-54]
'0071b734    52                      push edx
'0071b735    8b85e0fdffff            mov eax, dword ptr [ebp+fffffde0]
'0071b73b    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0071b73c    ff1580104000            call dword ptr [00401080]
'0071b742    898518fdffff            mov dword ptr [ebp+fffffd18], eax
'0071b748    eb0a                    jmp 71b754
    
End If
'0071b74a    c78518fdffff00000000    mov dword ptr [ebp+fffffd18], 00000000
'0071b754    8b4da0                  mov ecx, dword ptr [ebp-60]
'0071b757    898ddcfdffff            mov dword ptr [ebp+fffffddc], ecx
'0071b75d    c785a0feffff84d44300    mov dword ptr [ebp+fffffea0], 0043d484
'0071b767    c78598feffff08000000    mov dword ptr [ebp+fffffe98], 00000008
'0071b771    8d559c                  lea edx, dword ptr [ebp-64]
'0071b774    52                      push edx
'0071b775    b810000000              mov eax, 00000010

' *** Reference to "__vbaChkstk"
'0071b77a    e8e1baceff              call 407260
'0071b77f    8bc4                    mov eax, esp
'0071b781    8b8d98feffff            mov ecx, dword ptr [ebp+fffffe98]
'0071b787    8908                    mov dword ptr [eax], ecx
'0071b789    8b959cfeffff            mov edx, dword ptr [ebp+fffffe9c]
'0071b78f    895004                  mov dword ptr [eax+04], edx
'0071b792    8b8da0feffff            mov ecx, dword ptr [ebp+fffffea0]
'0071b798    894808                  mov dword ptr [eax+08], ecx
'0071b79b    8b95a4feffff            mov edx, dword ptr [ebp+fffffea4]
'0071b7a1    89500c                  mov dword ptr [eax+0c], edx
'0071b7a4    8b85dcfdffff            mov eax, dword ptr [ebp+fffffddc]
'0071b7aa    8b08                    mov ecx, dword ptr [eax]
'0071b7ac    8b95dcfdffff            mov edx, dword ptr [ebp+fffffddc]
'0071b7b2    52                      push edx
'0071b7b3    ff5130                  call dword ptr [ecx+30]
'0071b7b6    dbe2                    fnclex
'0071b7b8    8985d8fdffff            mov dword ptr [ebp+fffffdd8], eax
'0071b7be    83bdd8fdffff00          cmp dword ptr [ebp+fffffdd8], 00
'0071b7c5    7d23                    jge 71b7ea

If (var_7 < 0) Then
'0071b7c7    6a30                    push 30
'0071b7c9    68d8304300              push 004330d8
'0071b7ce    8b85dcfdffff            mov eax, dword ptr [ebp+fffffddc]
'0071b7d4    50                      push eax
'0071b7d5    8b8dd8fdffff            mov ecx, dword ptr [ebp+fffffdd8]
'0071b7db    51                      push ecx

' *** Reference to "__vbaHresultCheckObj"
'0071b7dc    ff1580104000            call dword ptr [00401080]
'0071b7e2    898514fdffff            mov dword ptr [ebp+fffffd14], eax
'0071b7e8    eb0a                    jmp 71b7f4
    
End If
'0071b7ea    c78514fdffff00000000    mov dword ptr [ebp+fffffd14], 00000000
'0071b7f4    8b559c                  mov edx, dword ptr [ebp-64]
'0071b7f7    89957cfdffff            mov dword ptr [ebp+fffffd7c], edx
'0071b7fd    c7459c00000000          mov dword ptr [ebp-64], 00000000
'0071b804    8b857cfdffff            mov eax, dword ptr [ebp+fffffd7c]
'0071b80a    894580                  mov dword ptr [ebp-80], eax
'0071b80d    c78578ffffff09000000    mov dword ptr [ebp+ffffff78], 00000009
'0071b817    8d8d78ffffff            lea ecx, dword ptr [ebp+ffffff78]
'0071b81d    51                      push ecx

' *** Reference to "rtcIsNull"
'0071b81e    ff1540114000            call dword ptr [00401140]
'0071b824    668985e4fdffff          mov word ptr [ebp+fffffde4], ax
'0071b82b    8d5598                  lea edx, dword ptr [ebp-68]
'0071b82e    52                      push edx
'0071b82f    8b45ac                  mov eax, dword ptr [ebp-54]
'0071b832    8b08                    mov ecx, dword ptr [eax]
'0071b834    8b55ac                  mov edx, dword ptr [ebp-54]
'0071b837    52                      push edx
'0071b838    ff91b4000000            call dword ptr [ecx+000000b4]
'0071b83e    dbe2                    fnclex
'0071b840    8985d4fdffff            mov dword ptr [ebp+fffffdd4], eax
'0071b846    83bdd4fdffff00          cmp dword ptr [ebp+fffffdd4], 00
'0071b84d    7d23                    jge 71b872

If (var_50 < 0) Then
'0071b84f    68b4000000              push 000000b4
'0071b854    6830314300              push 00433130
'0071b859    8b45ac                  mov eax, dword ptr [ebp-54]
'0071b85c    50                      push eax
'0071b85d    8b8dd4fdffff            mov ecx, dword ptr [ebp+fffffdd4]
'0071b863    51                      push ecx

' *** Reference to "__vbaHresultCheckObj"
'0071b864    ff1580104000            call dword ptr [00401080]
'0071b86a    898510fdffff            mov dword ptr [ebp+fffffd10], eax
'0071b870    eb0a                    jmp 71b87c
    
End If
'0071b872    c78510fdffff00000000    mov dword ptr [ebp+fffffd10], 00000000
'0071b87c    8b5598                  mov edx, dword ptr [ebp-68]
'0071b87f    8995d0fdffff            mov dword ptr [ebp+fffffdd0], edx
'0071b885    c78590feffff84d44300    mov dword ptr [ebp+fffffe90], 0043d484
'0071b88f    c78588feffff08000000    mov dword ptr [ebp+fffffe88], 00000008
'0071b899    8d4594                  lea eax, dword ptr [ebp-6c]
'0071b89c    50                      push eax
'0071b89d    b810000000              mov eax, 00000010

' *** Reference to "__vbaChkstk"
'0071b8a2    e8b9b9ceff              call 407260
'0071b8a7    8bcc                    mov ecx, esp
'0071b8a9    8b9588feffff            mov edx, dword ptr [ebp+fffffe88]
'0071b8af    8911                    mov dword ptr [ecx], edx
'0071b8b1    8b858cfeffff            mov eax, dword ptr [ebp+fffffe8c]
'0071b8b7    894104                  mov dword ptr [ecx+04], eax
'0071b8ba    8b9590feffff            mov edx, dword ptr [ebp+fffffe90]
'0071b8c0    895108                  mov dword ptr [ecx+08], edx
'0071b8c3    8b8594feffff            mov eax, dword ptr [ebp+fffffe94]
'0071b8c9    89410c                  mov dword ptr [ecx+0c], eax
'0071b8cc    8b8dd0fdffff            mov ecx, dword ptr [ebp+fffffdd0]
'0071b8d2    8b11                    mov edx, dword ptr [ecx]
'0071b8d4    8b85d0fdffff            mov eax, dword ptr [ebp+fffffdd0]
'0071b8da    50                      push eax
'0071b8db    ff5230                  call dword ptr [edx+30]
'0071b8de    dbe2                    fnclex
'0071b8e0    8985ccfdffff            mov dword ptr [ebp+fffffdcc], eax
'0071b8e6    83bdccfdffff00          cmp dword ptr [ebp+fffffdcc], 00
'0071b8ed    7d23                    jge 71b912

If (0 < 0) Then
'0071b8ef    6a30                    push 30
'0071b8f1    68d8304300              push 004330d8
'0071b8f6    8b8dd0fdffff            mov ecx, dword ptr [ebp+fffffdd0]
'0071b8fc    51                      push ecx
'0071b8fd    8b95ccfdffff            mov edx, dword ptr [ebp+fffffdcc]
'0071b903    52                      push edx

' *** Reference to "__vbaHresultCheckObj"
'0071b904    ff1580104000            call dword ptr [00401080]
'0071b90a    89850cfdffff            mov dword ptr [ebp+fffffd0c], eax
'0071b910    eb0a                    jmp 71b91c
    
End If
'0071b912    c7850cfdffff00000000    mov dword ptr [ebp+fffffd0c], 00000000
'0071b91c    8b4594                  mov eax, dword ptr [ebp-6c]
'0071b91f    898578fdffff            mov dword ptr [ebp+fffffd78], eax
'0071b925    c7459400000000          mov dword ptr [ebp-6c], 00000000
'0071b92c    8b8d78fdffff            mov ecx, dword ptr [ebp+fffffd78]
'0071b932    898d60ffffff            mov dword ptr [ebp+ffffff60], ecx
'0071b938    c78558ffffff09000000    mov dword ptr [ebp+ffffff58], 00000009
'0071b942    c78570ffffff00000000    mov dword ptr [ebp+ffffff70], 00000000
'0071b94c    c78568ffffff02000000    mov dword ptr [ebp+ffffff68], 00000002
'0071b956    668b95e4fdffff          mov dx, word ptr [ebp+fffffde4]
'0071b95d    66899580feffff          mov word ptr [ebp+fffffe80], dx
'0071b964    c78578feffff0b000000    mov dword ptr [ebp+fffffe78], 0000000b
'0071b96e    8d8558ffffff            lea eax, dword ptr [ebp+ffffff58]
'0071b974    50                      push eax
'0071b975    8d8d68ffffff            lea ecx, dword ptr [ebp+ffffff68]
'0071b97b    51                      push ecx
'0071b97c    8d9578feffff            lea edx, dword ptr [ebp+fffffe78]
'0071b982    52                      push edx
'0071b983    8d8548ffffff            lea eax, dword ptr [ebp+ffffff48]
'0071b989    50                      push eax

' *** Reference to "rtcImmediateIf"
'0071b98a    ff154c124000            call dword ptr [0040124c]
'0071b990    8d8d48ffffff            lea ecx, dword ptr [ebp+ffffff48]
'0071b996    51                      push ecx

' *** Reference to "__vbaI2Var"
'0071b997    ff150c124000            call dword ptr [0040120c]
'0071b99d    668945d8                mov word ptr [ebp-28], ax
'0071b9a1    8d5598                  lea edx, dword ptr [ebp-68]
'0071b9a4    52                      push edx
'0071b9a5    8d45a0                  lea eax, dword ptr [ebp-60]
'0071b9a8    50                      push eax
'0071b9a9    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'0071b9ab    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_7, var_130)
'0071b9b1    83c40c                  add esp, 0c
'0071b9b4    8d8d48ffffff            lea ecx, dword ptr [ebp+ffffff48]
'0071b9ba    51                      push ecx
'0071b9bb    8d9558ffffff            lea edx, dword ptr [ebp+ffffff58]
'0071b9c1    52                      push edx
'0071b9c2    8d8568ffffff            lea eax, dword ptr [ebp+ffffff68]
'0071b9c8    50                      push eax
'0071b9c9    8d8d78feffff            lea ecx, dword ptr [ebp+fffffe78]
'0071b9cf    51                      push ecx
'0071b9d0    8d9578ffffff            lea edx, dword ptr [ebp+ffffff78]
'0071b9d6    52                      push edx
'0071b9d7    6a05                    push 05

' *** Reference to "__vbaFreeVarList"
'0071b9d9    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_87, var_73, var_132, var_133, var_134)
'0071b9df    83c418                  add esp, 18
'0071b9e2    c745fc0e000000          mov dword ptr [ebp-04], 0000000e
'0071b9e9    8d45a0                  lea eax, dword ptr [ebp-60]
'0071b9ec    50                      push eax
'0071b9ed    8b4dac                  mov ecx, dword ptr [ebp-54]
'0071b9f0    8b11                    mov edx, dword ptr [ecx]
'0071b9f2    8b45ac                  mov eax, dword ptr [ebp-54]
'0071b9f5    50                      push eax
'0071b9f6    ff92b4000000            call dword ptr [edx+000000b4]
'0071b9fc    dbe2                    fnclex
'0071b9fe    8985e0fdffff            mov dword ptr [ebp+fffffde0], eax
'0071ba04    83bde0fdffff00          cmp dword ptr [ebp+fffffde0], 00
'0071ba0b    7d23                    jge 71ba30

If (var_50 < 0) Then
'0071ba0d    68b4000000              push 000000b4
'0071ba12    6830314300              push 00433130
'0071ba17    8b4dac                  mov ecx, dword ptr [ebp-54]
'0071ba1a    51                      push ecx
'0071ba1b    8b95e0fdffff            mov edx, dword ptr [ebp+fffffde0]
'0071ba21    52                      push edx

' *** Reference to "__vbaHresultCheckObj"
'0071ba22    ff1580104000            call dword ptr [00401080]
'0071ba28    898508fdffff            mov dword ptr [ebp+fffffd08], eax
'0071ba2e    eb0a                    jmp 71ba3a
    
End If
'0071ba30    c78508fdffff00000000    mov dword ptr [ebp+fffffd08], 00000000
'0071ba3a    8b45a0                  mov eax, dword ptr [ebp-60]
'0071ba3d    8985dcfdffff            mov dword ptr [ebp+fffffddc], eax
'0071ba43    c785a0feffff08ab4300    mov dword ptr [ebp+fffffea0], 0043ab08
'0071ba4d    c78598feffff08000000    mov dword ptr [ebp+fffffe98], 00000008
'0071ba57    8d4d9c                  lea ecx, dword ptr [ebp-64]
'0071ba5a    51                      push ecx
'0071ba5b    b810000000              mov eax, 00000010

' *** Reference to "__vbaChkstk"
'0071ba60    e8fbb7ceff              call 407260
'0071ba65    8bd4                    mov edx, esp
'0071ba67    8b8598feffff            mov eax, dword ptr [ebp+fffffe98]
'0071ba6d    8902                    mov dword ptr [edx], eax
'0071ba6f    8b8d9cfeffff            mov ecx, dword ptr [ebp+fffffe9c]
'0071ba75    894a04                  mov dword ptr [edx+04], ecx
'0071ba78    8b85a0feffff            mov eax, dword ptr [ebp+fffffea0]
'0071ba7e    894208                  mov dword ptr [edx+08], eax
'0071ba81    8b8da4feffff            mov ecx, dword ptr [ebp+fffffea4]
'0071ba87    894a0c                  mov dword ptr [edx+0c], ecx
'0071ba8a    8b95dcfdffff            mov edx, dword ptr [ebp+fffffddc]
'0071ba90    8b02                    mov eax, dword ptr [edx]
'0071ba92    8b8ddcfdffff            mov ecx, dword ptr [ebp+fffffddc]
'0071ba98    51                      push ecx
'0071ba99    ff5030                  call dword ptr [eax+30]
'0071ba9c    dbe2                    fnclex
'0071ba9e    8985d8fdffff            mov dword ptr [ebp+fffffdd8], eax
'0071baa4    83bdd8fdffff00          cmp dword ptr [ebp+fffffdd8], 00
'0071baab    7d23                    jge 71bad0

If (-256 - 24 < 0) Then
'0071baad    6a30                    push 30
'0071baaf    68d8304300              push 004330d8
'0071bab4    8b95dcfdffff            mov edx, dword ptr [ebp+fffffddc]
'0071baba    52                      push edx
'0071babb    8b85d8fdffff            mov eax, dword ptr [ebp+fffffdd8]
'0071bac1    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0071bac2    ff1580104000            call dword ptr [00401080]
'0071bac8    898504fdffff            mov dword ptr [ebp+fffffd04], eax
'0071bace    eb0a                    jmp 71bada
    
End If
'0071bad0    c78504fdffff00000000    mov dword ptr [ebp+fffffd04], 00000000
'0071bada    8b4d9c                  mov ecx, dword ptr [ebp-64]
'0071badd    898d74fdffff            mov dword ptr [ebp+fffffd74], ecx
'0071bae3    c7459c00000000          mov dword ptr [ebp-64], 00000000
'0071baea    8b9574fdffff            mov edx, dword ptr [ebp+fffffd74]
'0071baf0    895580                  mov dword ptr [ebp-80], edx
'0071baf3    c78578ffffff09000000    mov dword ptr [ebp+ffffff78], 00000009
'0071bafd    8d8578ffffff            lea eax, dword ptr [ebp+ffffff78]
'0071bb03    50                      push eax

' *** Reference to "rtcIsNull"
'0071bb04    ff1540114000            call dword ptr [00401140]
'0071bb0a    668985e4fdffff          mov word ptr [ebp+fffffde4], ax
'0071bb11    8d4d98                  lea ecx, dword ptr [ebp-68]
'0071bb14    51                      push ecx
'0071bb15    8b55ac                  mov edx, dword ptr [ebp-54]
'0071bb18    8b02                    mov eax, dword ptr [edx]
'0071bb1a    8b4dac                  mov ecx, dword ptr [ebp-54]
'0071bb1d    51                      push ecx
'0071bb1e    ff90b4000000            call dword ptr [eax+000000b4]
'0071bb24    dbe2                    fnclex
'0071bb26    8985d4fdffff            mov dword ptr [ebp+fffffdd4], eax
'0071bb2c    83bdd4fdffff00          cmp dword ptr [ebp+fffffdd4], 00
'0071bb33    7d23                    jge 71bb58

If (frmImport < 0) Then
'0071bb35    68b4000000              push 000000b4
'0071bb3a    6830314300              push 00433130
'0071bb3f    8b55ac                  mov edx, dword ptr [ebp-54]
'0071bb42    52                      push edx
'0071bb43    8b85d4fdffff            mov eax, dword ptr [ebp+fffffdd4]
'0071bb49    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0071bb4a    ff1580104000            call dword ptr [00401080]
'0071bb50    898500fdffff            mov dword ptr [ebp+fffffd00], eax
'0071bb56    eb0a                    jmp 71bb62
    
End If
'0071bb58    c78500fdffff00000000    mov dword ptr [ebp+fffffd00], 00000000
'0071bb62    8b4d98                  mov ecx, dword ptr [ebp-68]
'0071bb65    898dd0fdffff            mov dword ptr [ebp+fffffdd0], ecx
'0071bb6b    c78590feffff08ab4300    mov dword ptr [ebp+fffffe90], 0043ab08
'0071bb75    c78588feffff08000000    mov dword ptr [ebp+fffffe88], 00000008
'0071bb7f    8d5594                  lea edx, dword ptr [ebp-6c]
'0071bb82    52                      push edx
'0071bb83    b810000000              mov eax, 00000010

' *** Reference to "__vbaChkstk"
'0071bb88    e8d3b6ceff              call 407260
'0071bb8d    8bc4                    mov eax, esp
'0071bb8f    8b8d88feffff            mov ecx, dword ptr [ebp+fffffe88]
'0071bb95    8908                    mov dword ptr [eax], ecx
'0071bb97    8b958cfeffff            mov edx, dword ptr [ebp+fffffe8c]
'0071bb9d    895004                  mov dword ptr [eax+04], edx
'0071bba0    8b8d90feffff            mov ecx, dword ptr [ebp+fffffe90]
'0071bba6    894808                  mov dword ptr [eax+08], ecx
'0071bba9    8b9594feffff            mov edx, dword ptr [ebp+fffffe94]
'0071bbaf    89500c                  mov dword ptr [eax+0c], edx
'0071bbb2    8b85d0fdffff            mov eax, dword ptr [ebp+fffffdd0]
'0071bbb8    8b08                    mov ecx, dword ptr [eax]
'0071bbba    8b95d0fdffff            mov edx, dword ptr [ebp+fffffdd0]
'0071bbc0    52                      push edx
'0071bbc1    ff5130                  call dword ptr [ecx+30]
'0071bbc4    dbe2                    fnclex
'0071bbc6    8985ccfdffff            mov dword ptr [ebp+fffffdcc], eax
'0071bbcc    83bdccfdffff00          cmp dword ptr [ebp+fffffdcc], 00
'0071bbd3    7d23                    jge 71bbf8

If (0 < 0) Then
'0071bbd5    6a30                    push 30
'0071bbd7    68d8304300              push 004330d8
'0071bbdc    8b85d0fdffff            mov eax, dword ptr [ebp+fffffdd0]
'0071bbe2    50                      push eax
'0071bbe3    8b8dccfdffff            mov ecx, dword ptr [ebp+fffffdcc]
'0071bbe9    51                      push ecx

' *** Reference to "__vbaHresultCheckObj"
'0071bbea    ff1580104000            call dword ptr [00401080]
'0071bbf0    8985fcfcffff            mov dword ptr [ebp+fffffcfc], eax
'0071bbf6    eb0a                    jmp 71bc02
    
End If
'0071bbf8    c785fcfcffff00000000    mov dword ptr [ebp+fffffcfc], 00000000
'0071bc02    8b5594                  mov edx, dword ptr [ebp-6c]
'0071bc05    899570fdffff            mov dword ptr [ebp+fffffd70], edx
'0071bc0b    c7459400000000          mov dword ptr [ebp-6c], 00000000
'0071bc12    8b8570fdffff            mov eax, dword ptr [ebp+fffffd70]
'0071bc18    898560ffffff            mov dword ptr [ebp+ffffff60], eax
'0071bc1e    c78558ffffff09000000    mov dword ptr [ebp+ffffff58], 00000009
'0071bc28    c78570ffffff00000000    mov dword ptr [ebp+ffffff70], 00000000
'0071bc32    c78568ffffff02000000    mov dword ptr [ebp+ffffff68], 00000002
'0071bc3c    668b8de4fdffff          mov cx, word ptr [ebp+fffffde4]
'0071bc43    66898d80feffff          mov word ptr [ebp+fffffe80], cx
'0071bc4a    c78578feffff0b000000    mov dword ptr [ebp+fffffe78], 0000000b
'0071bc54    8d9558ffffff            lea edx, dword ptr [ebp+ffffff58]
'0071bc5a    52                      push edx
'0071bc5b    8d8568ffffff            lea eax, dword ptr [ebp+ffffff68]
'0071bc61    50                      push eax
'0071bc62    8d8d78feffff            lea ecx, dword ptr [ebp+fffffe78]
'0071bc68    51                      push ecx
'0071bc69    8d9548ffffff            lea edx, dword ptr [ebp+ffffff48]
'0071bc6f    52                      push edx

' *** Reference to "rtcImmediateIf"
'0071bc70    ff154c124000            call dword ptr [0040124c]
'0071bc76    8d8548ffffff            lea eax, dword ptr [ebp+ffffff48]
'0071bc7c    50                      push eax

' *** Reference to "__vbaI2Var"
'0071bc7d    ff150c124000            call dword ptr [0040120c]
'0071bc83    668945d0                mov word ptr [ebp-30], ax
'0071bc87    8d4d98                  lea ecx, dword ptr [ebp-68]
'0071bc8a    51                      push ecx
'0071bc8b    8d55a0                  lea edx, dword ptr [ebp-60]
'0071bc8e    52                      push edx
'0071bc8f    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'0071bc91    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_7, var_130)
'0071bc97    83c40c                  add esp, 0c
'0071bc9a    8d8548ffffff            lea eax, dword ptr [ebp+ffffff48]
'0071bca0    50                      push eax
'0071bca1    8d8d58ffffff            lea ecx, dword ptr [ebp+ffffff58]
'0071bca7    51                      push ecx
'0071bca8    8d9568ffffff            lea edx, dword ptr [ebp+ffffff68]
'0071bcae    52                      push edx
'0071bcaf    8d8578feffff            lea eax, dword ptr [ebp+fffffe78]
'0071bcb5    50                      push eax
'0071bcb6    8d8d78ffffff            lea ecx, dword ptr [ebp+ffffff78]
'0071bcbc    51                      push ecx
'0071bcbd    6a05                    push 05

' *** Reference to "__vbaFreeVarList"
'0071bcbf    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_87, var_73, var_132, var_133, var_134)
'0071bcc5    83c418                  add esp, 18
'0071bcc8    c745fc0f000000          mov dword ptr [ebp-04], 0000000f
'0071bccf    8d55a0                  lea edx, dword ptr [ebp-60]
'0071bcd2    52                      push edx
'0071bcd3    8b45ac                  mov eax, dword ptr [ebp-54]
'0071bcd6    8b08                    mov ecx, dword ptr [eax]
'0071bcd8    8b55ac                  mov edx, dword ptr [ebp-54]
'0071bcdb    52                      push edx
'0071bcdc    ff91b4000000            call dword ptr [ecx+000000b4]
'0071bce2    dbe2                    fnclex
'0071bce4    8985e0fdffff            mov dword ptr [ebp+fffffde0], eax
'0071bcea    83bde0fdffff00          cmp dword ptr [ebp+fffffde0], 00
'0071bcf1    7d23                    jge 71bd16

If (var_50 < 0) Then
'0071bcf3    68b4000000              push 000000b4
'0071bcf8    6830314300              push 00433130
'0071bcfd    8b45ac                  mov eax, dword ptr [ebp-54]
'0071bd00    50                      push eax
'0071bd01    8b8de0fdffff            mov ecx, dword ptr [ebp+fffffde0]
'0071bd07    51                      push ecx

' *** Reference to "__vbaHresultCheckObj"
'0071bd08    ff1580104000            call dword ptr [00401080]
'0071bd0e    8985f8fcffff            mov dword ptr [ebp+fffffcf8], eax
'0071bd14    eb0a                    jmp 71bd20
    
End If
'0071bd16    c785f8fcffff00000000    mov dword ptr [ebp+fffffcf8], 00000000
'0071bd20    8b55a0                  mov edx, dword ptr [ebp-60]
'0071bd23    8995dcfdffff            mov dword ptr [ebp+fffffddc], edx
'0071bd29    c785a0feffff18ab4300    mov dword ptr [ebp+fffffea0], 0043ab18
'0071bd33    c78598feffff08000000    mov dword ptr [ebp+fffffe98], 00000008
'0071bd3d    8d459c                  lea eax, dword ptr [ebp-64]
'0071bd40    50                      push eax
'0071bd41    b810000000              mov eax, 00000010

' *** Reference to "__vbaChkstk"
'0071bd46    e815b5ceff              call 407260
'0071bd4b    8bcc                    mov ecx, esp
'0071bd4d    8b9598feffff            mov edx, dword ptr [ebp+fffffe98]
'0071bd53    8911                    mov dword ptr [ecx], edx
'0071bd55    8b859cfeffff            mov eax, dword ptr [ebp+fffffe9c]
'0071bd5b    894104                  mov dword ptr [ecx+04], eax
'0071bd5e    8b95a0feffff            mov edx, dword ptr [ebp+fffffea0]
'0071bd64    895108                  mov dword ptr [ecx+08], edx
'0071bd67    8b85a4feffff            mov eax, dword ptr [ebp+fffffea4]
'0071bd6d    89410c                  mov dword ptr [ecx+0c], eax
'0071bd70    8b8ddcfdffff            mov ecx, dword ptr [ebp+fffffddc]
'0071bd76    8b11                    mov edx, dword ptr [ecx]
'0071bd78    8b85dcfdffff            mov eax, dword ptr [ebp+fffffddc]
'0071bd7e    50                      push eax
'0071bd7f    ff5230                  call dword ptr [edx+30]
'0071bd82    dbe2                    fnclex
'0071bd84    8985d8fdffff            mov dword ptr [ebp+fffffdd8], eax
'0071bd8a    83bdd8fdffff00          cmp dword ptr [ebp+fffffdd8], 00
'0071bd91    7d23                    jge 71bdb6

If (var_7 < 0) Then
'0071bd93    6a30                    push 30
'0071bd95    68d8304300              push 004330d8
'0071bd9a    8b8ddcfdffff            mov ecx, dword ptr [ebp+fffffddc]
'0071bda0    51                      push ecx
'0071bda1    8b95d8fdffff            mov edx, dword ptr [ebp+fffffdd8]
'0071bda7    52                      push edx

' *** Reference to "__vbaHresultCheckObj"
'0071bda8    ff1580104000            call dword ptr [00401080]
'0071bdae    8985f4fcffff            mov dword ptr [ebp+fffffcf4], eax
'0071bdb4    eb0a                    jmp 71bdc0
    
End If
'0071bdb6    c785f4fcffff00000000    mov dword ptr [ebp+fffffcf4], 00000000
'0071bdc0    8b459c                  mov eax, dword ptr [ebp-64]
'0071bdc3    89856cfdffff            mov dword ptr [ebp+fffffd6c], eax
'0071bdc9    c7459c00000000          mov dword ptr [ebp-64], 00000000
'0071bdd0    8b8d6cfdffff            mov ecx, dword ptr [ebp+fffffd6c]
'0071bdd6    894d80                  mov dword ptr [ebp-80], ecx
'0071bdd9    c78578ffffff09000000    mov dword ptr [ebp+ffffff78], 00000009
'0071bde3    8d9578ffffff            lea edx, dword ptr [ebp+ffffff78]
'0071bde9    52                      push edx

' *** Reference to "rtcIsNull"
'0071bdea    ff1540114000            call dword ptr [00401140]
'0071bdf0    668985e4fdffff          mov word ptr [ebp+fffffde4], ax
'0071bdf7    8d4598                  lea eax, dword ptr [ebp-68]
'0071bdfa    50                      push eax
'0071bdfb    8b4dac                  mov ecx, dword ptr [ebp-54]
'0071bdfe    8b11                    mov edx, dword ptr [ecx]
'0071be00    8b45ac                  mov eax, dword ptr [ebp-54]
'0071be03    50                      push eax
'0071be04    ff92b4000000            call dword ptr [edx+000000b4]
'0071be0a    dbe2                    fnclex
'0071be0c    8985d4fdffff            mov dword ptr [ebp+fffffdd4], eax
'0071be12    83bdd4fdffff00          cmp dword ptr [ebp+fffffdd4], 00
'0071be19    7d23                    jge 71be3e

If (var_50 < 0) Then
'0071be1b    68b4000000              push 000000b4
'0071be20    6830314300              push 00433130
'0071be25    8b4dac                  mov ecx, dword ptr [ebp-54]
'0071be28    51                      push ecx
'0071be29    8b95d4fdffff            mov edx, dword ptr [ebp+fffffdd4]
'0071be2f    52                      push edx

' *** Reference to "__vbaHresultCheckObj"
'0071be30    ff1580104000            call dword ptr [00401080]
'0071be36    8985f0fcffff            mov dword ptr [ebp+fffffcf0], eax
'0071be3c    eb0a                    jmp 71be48
    
End If
'0071be3e    c785f0fcffff00000000    mov dword ptr [ebp+fffffcf0], 00000000
'0071be48    8b4598                  mov eax, dword ptr [ebp-68]
'0071be4b    8985d0fdffff            mov dword ptr [ebp+fffffdd0], eax
'0071be51    c78590feffff18ab4300    mov dword ptr [ebp+fffffe90], 0043ab18
'0071be5b    c78588feffff08000000    mov dword ptr [ebp+fffffe88], 00000008
'0071be65    8d4d94                  lea ecx, dword ptr [ebp-6c]
'0071be68    51                      push ecx
'0071be69    b810000000              mov eax, 00000010

' *** Reference to "__vbaChkstk"
'0071be6e    e8edb3ceff              call 407260
'0071be73    8bd4                    mov edx, esp
'0071be75    8b8588feffff            mov eax, dword ptr [ebp+fffffe88]
'0071be7b    8902                    mov dword ptr [edx], eax
'0071be7d    8b8d8cfeffff            mov ecx, dword ptr [ebp+fffffe8c]
'0071be83    894a04                  mov dword ptr [edx+04], ecx
'0071be86    8b8590feffff            mov eax, dword ptr [ebp+fffffe90]
'0071be8c    894208                  mov dword ptr [edx+08], eax
'0071be8f    8b8d94feffff            mov ecx, dword ptr [ebp+fffffe94]
'0071be95    894a0c                  mov dword ptr [edx+0c], ecx
'0071be98    8b95d0fdffff            mov edx, dword ptr [ebp+fffffdd0]
'0071be9e    8b02                    mov eax, dword ptr [edx]
'0071bea0    8b8dd0fdffff            mov ecx, dword ptr [ebp+fffffdd0]
'0071bea6    51                      push ecx
'0071bea7    ff5030                  call dword ptr [eax+30]
'0071beaa    dbe2                    fnclex
'0071beac    8985ccfdffff            mov dword ptr [ebp+fffffdcc], eax
'0071beb2    83bdccfdffff00          cmp dword ptr [ebp+fffffdcc], 00
'0071beb9    7d23                    jge 71bede

If (-256 - 24 < 0) Then
'0071bebb    6a30                    push 30
'0071bebd    68d8304300              push 004330d8
'0071bec2    8b95d0fdffff            mov edx, dword ptr [ebp+fffffdd0]
'0071bec8    52                      push edx
'0071bec9    8b85ccfdffff            mov eax, dword ptr [ebp+fffffdcc]
'0071becf    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0071bed0    ff1580104000            call dword ptr [00401080]
'0071bed6    8985ecfcffff            mov dword ptr [ebp+fffffcec], eax
'0071bedc    eb0a                    jmp 71bee8
    
End If
'0071bede    c785ecfcffff00000000    mov dword ptr [ebp+fffffcec], 00000000
'0071bee8    8b4d94                  mov ecx, dword ptr [ebp-6c]
'0071beeb    898d68fdffff            mov dword ptr [ebp+fffffd68], ecx
'0071bef1    c7459400000000          mov dword ptr [ebp-6c], 00000000
'0071bef8    8b9568fdffff            mov edx, dword ptr [ebp+fffffd68]
'0071befe    899560ffffff            mov dword ptr [ebp+ffffff60], edx
'0071bf04    c78558ffffff09000000    mov dword ptr [ebp+ffffff58], 00000009
'0071bf0e    c78570ffffff00000000    mov dword ptr [ebp+ffffff70], 00000000
'0071bf18    c78568ffffff02000000    mov dword ptr [ebp+ffffff68], 00000002
'0071bf22    668b85e4fdffff          mov ax, word ptr [ebp+fffffde4]
'0071bf29    66898580feffff          mov word ptr [ebp+fffffe80], ax
'0071bf30    c78578feffff0b000000    mov dword ptr [ebp+fffffe78], 0000000b
'0071bf3a    8d8d58ffffff            lea ecx, dword ptr [ebp+ffffff58]
'0071bf40    51                      push ecx
'0071bf41    8d9568ffffff            lea edx, dword ptr [ebp+ffffff68]
'0071bf47    52                      push edx
'0071bf48    8d8578feffff            lea eax, dword ptr [ebp+fffffe78]
'0071bf4e    50                      push eax
'0071bf4f    8d8d48ffffff            lea ecx, dword ptr [ebp+ffffff48]
'0071bf55    51                      push ecx

' *** Reference to "rtcImmediateIf"
'0071bf56    ff154c124000            call dword ptr [0040124c]
'0071bf5c    8d9548ffffff            lea edx, dword ptr [ebp+ffffff48]
'0071bf62    52                      push edx

' *** Reference to "__vbaI2Var"
'0071bf63    ff150c124000            call dword ptr [0040120c]
'0071bf69    668945c8                mov word ptr [ebp-38], ax
'0071bf6d    8d4598                  lea eax, dword ptr [ebp-68]
'0071bf70    50                      push eax
'0071bf71    8d4da0                  lea ecx, dword ptr [ebp-60]
'0071bf74    51                      push ecx
'0071bf75    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'0071bf77    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_7, var_130)
'0071bf7d    83c40c                  add esp, 0c
'0071bf80    8d9548ffffff            lea edx, dword ptr [ebp+ffffff48]
'0071bf86    52                      push edx
'0071bf87    8d8558ffffff            lea eax, dword ptr [ebp+ffffff58]
'0071bf8d    50                      push eax
'0071bf8e    8d8d68ffffff            lea ecx, dword ptr [ebp+ffffff68]
'0071bf94    51                      push ecx
'0071bf95    8d9578feffff            lea edx, dword ptr [ebp+fffffe78]
'0071bf9b    52                      push edx
'0071bf9c    8d8578ffffff            lea eax, dword ptr [ebp+ffffff78]
'0071bfa2    50                      push eax
'0071bfa3    6a05                    push 05

' *** Reference to "__vbaFreeVarList"
'0071bfa5    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_87, var_73, var_132, var_133, var_134)
'0071bfab    83c418                  add esp, 18
'0071bfae    c745fc10000000          mov dword ptr [ebp-04], 00000010
'0071bfb5    8d4da0                  lea ecx, dword ptr [ebp-60]
'0071bfb8    51                      push ecx
'0071bfb9    8b55ac                  mov edx, dword ptr [ebp-54]
'0071bfbc    8b02                    mov eax, dword ptr [edx]
'0071bfbe    8b4dac                  mov ecx, dword ptr [ebp-54]
'0071bfc1    51                      push ecx
'0071bfc2    ff90b4000000            call dword ptr [eax+000000b4]
'0071bfc8    dbe2                    fnclex
'0071bfca    8985e0fdffff            mov dword ptr [ebp+fffffde0], eax
'0071bfd0    83bde0fdffff00          cmp dword ptr [ebp+fffffde0], 00
'0071bfd7    7d23                    jge 71bffc

If (frmImport < 0) Then
'0071bfd9    68b4000000              push 000000b4
'0071bfde    6830314300              push 00433130
'0071bfe3    8b55ac                  mov edx, dword ptr [ebp-54]
'0071bfe6    52                      push edx
'0071bfe7    8b85e0fdffff            mov eax, dword ptr [ebp+fffffde0]
'0071bfed    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0071bfee    ff1580104000            call dword ptr [00401080]
'0071bff4    8985e8fcffff            mov dword ptr [ebp+fffffce8], eax
'0071bffa    eb0a                    jmp 71c006
    
End If
'0071bffc    c785e8fcffff00000000    mov dword ptr [ebp+fffffce8], 00000000
'0071c006    8b4da0                  mov ecx, dword ptr [ebp-60]
'0071c009    898ddcfdffff            mov dword ptr [ebp+fffffddc], ecx
'0071c00f    c785a0feffff28ab4300    mov dword ptr [ebp+fffffea0], 0043ab28
'0071c019    c78598feffff08000000    mov dword ptr [ebp+fffffe98], 00000008
'0071c023    8d559c                  lea edx, dword ptr [ebp-64]
'0071c026    52                      push edx
'0071c027    b810000000              mov eax, 00000010

' *** Reference to "__vbaChkstk"
'0071c02c    e82fb2ceff              call 407260
'0071c031    8bc4                    mov eax, esp
'0071c033    8b8d98feffff            mov ecx, dword ptr [ebp+fffffe98]
'0071c039    8908                    mov dword ptr [eax], ecx
'0071c03b    8b959cfeffff            mov edx, dword ptr [ebp+fffffe9c]
'0071c041    895004                  mov dword ptr [eax+04], edx
'0071c044    8b8da0feffff            mov ecx, dword ptr [ebp+fffffea0]
'0071c04a    894808                  mov dword ptr [eax+08], ecx
'0071c04d    8b95a4feffff            mov edx, dword ptr [ebp+fffffea4]
'0071c053    89500c                  mov dword ptr [eax+0c], edx
'0071c056    8b85dcfdffff            mov eax, dword ptr [ebp+fffffddc]
'0071c05c    8b08                    mov ecx, dword ptr [eax]
'0071c05e    8b95dcfdffff            mov edx, dword ptr [ebp+fffffddc]
'0071c064    52                      push edx
'0071c065    ff5130                  call dword ptr [ecx+30]
'0071c068    dbe2                    fnclex
'0071c06a    8985d8fdffff            mov dword ptr [ebp+fffffdd8], eax
'0071c070    83bdd8fdffff00          cmp dword ptr [ebp+fffffdd8], 00
'0071c077    7d23                    jge 71c09c

If (var_7 < 0) Then
'0071c079    6a30                    push 30
'0071c07b    68d8304300              push 004330d8
'0071c080    8b85dcfdffff            mov eax, dword ptr [ebp+fffffddc]
'0071c086    50                      push eax
'0071c087    8b8dd8fdffff            mov ecx, dword ptr [ebp+fffffdd8]
'0071c08d    51                      push ecx

' *** Reference to "__vbaHresultCheckObj"
'0071c08e    ff1580104000            call dword ptr [00401080]
'0071c094    8985e4fcffff            mov dword ptr [ebp+fffffce4], eax
'0071c09a    eb0a                    jmp 71c0a6
    
End If
'0071c09c    c785e4fcffff00000000    mov dword ptr [ebp+fffffce4], 00000000
'0071c0a6    8b559c                  mov edx, dword ptr [ebp-64]
'0071c0a9    899564fdffff            mov dword ptr [ebp+fffffd64], edx
'0071c0af    c7459c00000000          mov dword ptr [ebp-64], 00000000
'0071c0b6    8b8564fdffff            mov eax, dword ptr [ebp+fffffd64]
'0071c0bc    894580                  mov dword ptr [ebp-80], eax
'0071c0bf    c78578ffffff09000000    mov dword ptr [ebp+ffffff78], 00000009
'0071c0c9    8d8d78ffffff            lea ecx, dword ptr [ebp+ffffff78]
'0071c0cf    51                      push ecx

' *** Reference to "rtcIsNull"
'0071c0d0    ff1540114000            call dword ptr [00401140]
'0071c0d6    668985e4fdffff          mov word ptr [ebp+fffffde4], ax
'0071c0dd    8d5598                  lea edx, dword ptr [ebp-68]
'0071c0e0    52                      push edx
'0071c0e1    8b45ac                  mov eax, dword ptr [ebp-54]
'0071c0e4    8b08                    mov ecx, dword ptr [eax]
'0071c0e6    8b55ac                  mov edx, dword ptr [ebp-54]
'0071c0e9    52                      push edx
'0071c0ea    ff91b4000000            call dword ptr [ecx+000000b4]
'0071c0f0    dbe2                    fnclex
'0071c0f2    8985d4fdffff            mov dword ptr [ebp+fffffdd4], eax
'0071c0f8    83bdd4fdffff00          cmp dword ptr [ebp+fffffdd4], 00
'0071c0ff    7d23                    jge 71c124

If (var_50 < 0) Then
'0071c101    68b4000000              push 000000b4
'0071c106    6830314300              push 00433130
'0071c10b    8b45ac                  mov eax, dword ptr [ebp-54]
'0071c10e    50                      push eax
'0071c10f    8b8dd4fdffff            mov ecx, dword ptr [ebp+fffffdd4]
'0071c115    51                      push ecx

' *** Reference to "__vbaHresultCheckObj"
'0071c116    ff1580104000            call dword ptr [00401080]
'0071c11c    8985e0fcffff            mov dword ptr [ebp+fffffce0], eax
'0071c122    eb0a                    jmp 71c12e
    
End If
'0071c124    c785e0fcffff00000000    mov dword ptr [ebp+fffffce0], 00000000
'0071c12e    8b5598                  mov edx, dword ptr [ebp-68]
'0071c131    8995d0fdffff            mov dword ptr [ebp+fffffdd0], edx
'0071c137    c78590feffff28ab4300    mov dword ptr [ebp+fffffe90], 0043ab28
'0071c141    c78588feffff08000000    mov dword ptr [ebp+fffffe88], 00000008
'0071c14b    8d4594                  lea eax, dword ptr [ebp-6c]
'0071c14e    50                      push eax
'0071c14f    b810000000              mov eax, 00000010

' *** Reference to "__vbaChkstk"
'0071c154    e807b1ceff              call 407260
'0071c159    8bcc                    mov ecx, esp
'0071c15b    8b9588feffff            mov edx, dword ptr [ebp+fffffe88]
'0071c161    8911                    mov dword ptr [ecx], edx
'0071c163    8b858cfeffff            mov eax, dword ptr [ebp+fffffe8c]
'0071c169    894104                  mov dword ptr [ecx+04], eax
'0071c16c    8b9590feffff            mov edx, dword ptr [ebp+fffffe90]
'0071c172    895108                  mov dword ptr [ecx+08], edx
'0071c175    8b8594feffff            mov eax, dword ptr [ebp+fffffe94]
'0071c17b    89410c                  mov dword ptr [ecx+0c], eax
'0071c17e    8b8dd0fdffff            mov ecx, dword ptr [ebp+fffffdd0]
'0071c184    8b11                    mov edx, dword ptr [ecx]
'0071c186    8b85d0fdffff            mov eax, dword ptr [ebp+fffffdd0]
'0071c18c    50                      push eax
'0071c18d    ff5230                  call dword ptr [edx+30]
'0071c190    dbe2                    fnclex
'0071c192    8985ccfdffff            mov dword ptr [ebp+fffffdcc], eax
'0071c198    83bdccfdffff00          cmp dword ptr [ebp+fffffdcc], 00
'0071c19f    7d23                    jge 71c1c4

If (0 < 0) Then
'0071c1a1    6a30                    push 30
'0071c1a3    68d8304300              push 004330d8
'0071c1a8    8b8dd0fdffff            mov ecx, dword ptr [ebp+fffffdd0]
'0071c1ae    51                      push ecx
'0071c1af    8b95ccfdffff            mov edx, dword ptr [ebp+fffffdcc]
'0071c1b5    52                      push edx

' *** Reference to "__vbaHresultCheckObj"
'0071c1b6    ff1580104000            call dword ptr [00401080]
'0071c1bc    8985dcfcffff            mov dword ptr [ebp+fffffcdc], eax
'0071c1c2    eb0a                    jmp 71c1ce
    
End If
'0071c1c4    c785dcfcffff00000000    mov dword ptr [ebp+fffffcdc], 00000000
'0071c1ce    8b4594                  mov eax, dword ptr [ebp-6c]
'0071c1d1    898560fdffff            mov dword ptr [ebp+fffffd60], eax
'0071c1d7    c7459400000000          mov dword ptr [ebp-6c], 00000000
'0071c1de    8b8d60fdffff            mov ecx, dword ptr [ebp+fffffd60]
'0071c1e4    898d60ffffff            mov dword ptr [ebp+ffffff60], ecx
'0071c1ea    c78558ffffff09000000    mov dword ptr [ebp+ffffff58], 00000009
'0071c1f4    c78570ffffff00000000    mov dword ptr [ebp+ffffff70], 00000000
'0071c1fe    c78568ffffff02000000    mov dword ptr [ebp+ffffff68], 00000002
'0071c208    668b95e4fdffff          mov dx, word ptr [ebp+fffffde4]
'0071c20f    66899580feffff          mov word ptr [ebp+fffffe80], dx
'0071c216    c78578feffff0b000000    mov dword ptr [ebp+fffffe78], 0000000b
'0071c220    8d8558ffffff            lea eax, dword ptr [ebp+ffffff58]
'0071c226    50                      push eax
'0071c227    8d8d68ffffff            lea ecx, dword ptr [ebp+ffffff68]
'0071c22d    51                      push ecx
'0071c22e    8d9578feffff            lea edx, dword ptr [ebp+fffffe78]
'0071c234    52                      push edx
'0071c235    8d8548ffffff            lea eax, dword ptr [ebp+ffffff48]
'0071c23b    50                      push eax

' *** Reference to "rtcImmediateIf"
'0071c23c    ff154c124000            call dword ptr [0040124c]
'0071c242    8d8d48ffffff            lea ecx, dword ptr [ebp+ffffff48]
'0071c248    51                      push ecx

' *** Reference to "__vbaI2Var"
'0071c249    ff150c124000            call dword ptr [0040120c]
'0071c24f    668945c4                mov word ptr [ebp-3c], ax
'0071c253    8d5598                  lea edx, dword ptr [ebp-68]
'0071c256    52                      push edx
'0071c257    8d45a0                  lea eax, dword ptr [ebp-60]
'0071c25a    50                      push eax
'0071c25b    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'0071c25d    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_7, var_130)
'0071c263    83c40c                  add esp, 0c
'0071c266    8d8d48ffffff            lea ecx, dword ptr [ebp+ffffff48]
'0071c26c    51                      push ecx
'0071c26d    8d9558ffffff            lea edx, dword ptr [ebp+ffffff58]
'0071c273    52                      push edx
'0071c274    8d8568ffffff            lea eax, dword ptr [ebp+ffffff68]
'0071c27a    50                      push eax
'0071c27b    8d8d78feffff            lea ecx, dword ptr [ebp+fffffe78]
'0071c281    51                      push ecx
'0071c282    8d9578ffffff            lea edx, dword ptr [ebp+ffffff78]
'0071c288    52                      push edx
'0071c289    6a05                    push 05

' *** Reference to "__vbaFreeVarList"
'0071c28b    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_87, var_73, var_132, var_133, var_134)
'0071c291    83c418                  add esp, 18
'0071c294    c745fc11000000          mov dword ptr [ebp-04], 00000011
'0071c29b    8d45a0                  lea eax, dword ptr [ebp-60]
'0071c29e    50                      push eax
'0071c29f    8b4dac                  mov ecx, dword ptr [ebp-54]
'0071c2a2    8b11                    mov edx, dword ptr [ecx]
'0071c2a4    8b45ac                  mov eax, dword ptr [ebp-54]
'0071c2a7    50                      push eax
'0071c2a8    ff92b4000000            call dword ptr [edx+000000b4]
'0071c2ae    dbe2                    fnclex
'0071c2b0    8985e0fdffff            mov dword ptr [ebp+fffffde0], eax
'0071c2b6    83bde0fdffff00          cmp dword ptr [ebp+fffffde0], 00
'0071c2bd    7d23                    jge 71c2e2

If (var_50 < 0) Then
'0071c2bf    68b4000000              push 000000b4
'0071c2c4    6830314300              push 00433130
'0071c2c9    8b4dac                  mov ecx, dword ptr [ebp-54]
'0071c2cc    51                      push ecx
'0071c2cd    8b95e0fdffff            mov edx, dword ptr [ebp+fffffde0]
'0071c2d3    52                      push edx

' *** Reference to "__vbaHresultCheckObj"
'0071c2d4    ff1580104000            call dword ptr [00401080]
'0071c2da    8985d8fcffff            mov dword ptr [ebp+fffffcd8], eax
'0071c2e0    eb0a                    jmp 71c2ec
    
End If
'0071c2e2    c785d8fcffff00000000    mov dword ptr [ebp+fffffcd8], 00000000
'0071c2ec    8b45a0                  mov eax, dword ptr [ebp-60]
'0071c2ef    8985dcfdffff            mov dword ptr [ebp+fffffddc], eax
'0071c2f5    c785a0feffffac874300    mov dword ptr [ebp+fffffea0], 004387ac
'0071c2ff    c78598feffff08000000    mov dword ptr [ebp+fffffe98], 00000008
'0071c309    8d4d9c                  lea ecx, dword ptr [ebp-64]
'0071c30c    51                      push ecx
'0071c30d    b810000000              mov eax, 00000010

' *** Reference to "__vbaChkstk"
'0071c312    e849afceff              call 407260
'0071c317    8bd4                    mov edx, esp
'0071c319    8b8598feffff            mov eax, dword ptr [ebp+fffffe98]
'0071c31f    8902                    mov dword ptr [edx], eax
'0071c321    8b8d9cfeffff            mov ecx, dword ptr [ebp+fffffe9c]
'0071c327    894a04                  mov dword ptr [edx+04], ecx
'0071c32a    8b85a0feffff            mov eax, dword ptr [ebp+fffffea0]
'0071c330    894208                  mov dword ptr [edx+08], eax
'0071c333    8b8da4feffff            mov ecx, dword ptr [ebp+fffffea4]
'0071c339    894a0c                  mov dword ptr [edx+0c], ecx
'0071c33c    8b95dcfdffff            mov edx, dword ptr [ebp+fffffddc]
'0071c342    8b02                    mov eax, dword ptr [edx]
'0071c344    8b8ddcfdffff            mov ecx, dword ptr [ebp+fffffddc]
'0071c34a    51                      push ecx
'0071c34b    ff5030                  call dword ptr [eax+30]
'0071c34e    dbe2                    fnclex
'0071c350    8985d8fdffff            mov dword ptr [ebp+fffffdd8], eax
'0071c356    83bdd8fdffff00          cmp dword ptr [ebp+fffffdd8], 00
'0071c35d    7d23                    jge 71c382

If (-256 - 24 < 0) Then
'0071c35f    6a30                    push 30
'0071c361    68d8304300              push 004330d8
'0071c366    8b95dcfdffff            mov edx, dword ptr [ebp+fffffddc]
'0071c36c    52                      push edx
'0071c36d    8b85d8fdffff            mov eax, dword ptr [ebp+fffffdd8]
'0071c373    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0071c374    ff1580104000            call dword ptr [00401080]
'0071c37a    8985d4fcffff            mov dword ptr [ebp+fffffcd4], eax
'0071c380    eb0a                    jmp 71c38c
    
End If
'0071c382    c785d4fcffff00000000    mov dword ptr [ebp+fffffcd4], 00000000
'0071c38c    8b4d9c                  mov ecx, dword ptr [ebp-64]
'0071c38f    898d5cfdffff            mov dword ptr [ebp+fffffd5c], ecx
'0071c395    c7459c00000000          mov dword ptr [ebp-64], 00000000
'0071c39c    8b955cfdffff            mov edx, dword ptr [ebp+fffffd5c]
'0071c3a2    895580                  mov dword ptr [ebp-80], edx
'0071c3a5    c78578ffffff09000000    mov dword ptr [ebp+ffffff78], 00000009
'0071c3af    8d8578ffffff            lea eax, dword ptr [ebp+ffffff78]
'0071c3b5    50                      push eax

' *** Reference to "rtcIsNull"
'0071c3b6    ff1540114000            call dword ptr [00401140]
'0071c3bc    668985e4fdffff          mov word ptr [ebp+fffffde4], ax
'0071c3c3    8d4d98                  lea ecx, dword ptr [ebp-68]
'0071c3c6    51                      push ecx
'0071c3c7    8b55ac                  mov edx, dword ptr [ebp-54]
'0071c3ca    8b02                    mov eax, dword ptr [edx]
'0071c3cc    8b4dac                  mov ecx, dword ptr [ebp-54]
'0071c3cf    51                      push ecx
'0071c3d0    ff90b4000000            call dword ptr [eax+000000b4]
'0071c3d6    dbe2                    fnclex
'0071c3d8    8985d4fdffff            mov dword ptr [ebp+fffffdd4], eax
'0071c3de    83bdd4fdffff00          cmp dword ptr [ebp+fffffdd4], 00
'0071c3e5    7d23                    jge 71c40a

If (frmImport < 0) Then
'0071c3e7    68b4000000              push 000000b4
'0071c3ec    6830314300              push 00433130
'0071c3f1    8b55ac                  mov edx, dword ptr [ebp-54]
'0071c3f4    52                      push edx
'0071c3f5    8b85d4fdffff            mov eax, dword ptr [ebp+fffffdd4]
'0071c3fb    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0071c3fc    ff1580104000            call dword ptr [00401080]
'0071c402    8985d0fcffff            mov dword ptr [ebp+fffffcd0], eax
'0071c408    eb0a                    jmp 71c414
    
End If
'0071c40a    c785d0fcffff00000000    mov dword ptr [ebp+fffffcd0], 00000000
'0071c414    8b4d98                  mov ecx, dword ptr [ebp-68]
'0071c417    898dd0fdffff            mov dword ptr [ebp+fffffdd0], ecx
'0071c41d    c78590feffffac874300    mov dword ptr [ebp+fffffe90], 004387ac
'0071c427    c78588feffff08000000    mov dword ptr [ebp+fffffe88], 00000008
'0071c431    8d5594                  lea edx, dword ptr [ebp-6c]
'0071c434    52                      push edx
'0071c435    b810000000              mov eax, 00000010

' *** Reference to "__vbaChkstk"
'0071c43a    e821aeceff              call 407260
'0071c43f    8bc4                    mov eax, esp
'0071c441    8b8d88feffff            mov ecx, dword ptr [ebp+fffffe88]
'0071c447    8908                    mov dword ptr [eax], ecx
'0071c449    8b958cfeffff            mov edx, dword ptr [ebp+fffffe8c]
'0071c44f    895004                  mov dword ptr [eax+04], edx
'0071c452    8b8d90feffff            mov ecx, dword ptr [ebp+fffffe90]
'0071c458    894808                  mov dword ptr [eax+08], ecx
'0071c45b    8b9594feffff            mov edx, dword ptr [ebp+fffffe94]
'0071c461    89500c                  mov dword ptr [eax+0c], edx
'0071c464    8b85d0fdffff            mov eax, dword ptr [ebp+fffffdd0]
'0071c46a    8b08                    mov ecx, dword ptr [eax]
'0071c46c    8b95d0fdffff            mov edx, dword ptr [ebp+fffffdd0]
'0071c472    52                      push edx
'0071c473    ff5130                  call dword ptr [ecx+30]
'0071c476    dbe2                    fnclex
'0071c478    8985ccfdffff            mov dword ptr [ebp+fffffdcc], eax
'0071c47e    83bdccfdffff00          cmp dword ptr [ebp+fffffdcc], 00
'0071c485    7d23                    jge 71c4aa

If (0 < 0) Then
'0071c487    6a30                    push 30
'0071c489    68d8304300              push 004330d8
'0071c48e    8b85d0fdffff            mov eax, dword ptr [ebp+fffffdd0]
'0071c494    50                      push eax
'0071c495    8b8dccfdffff            mov ecx, dword ptr [ebp+fffffdcc]
'0071c49b    51                      push ecx

' *** Reference to "__vbaHresultCheckObj"
'0071c49c    ff1580104000            call dword ptr [00401080]
'0071c4a2    8985ccfcffff            mov dword ptr [ebp+fffffccc], eax
'0071c4a8    eb0a                    jmp 71c4b4
    
End If
'0071c4aa    c785ccfcffff00000000    mov dword ptr [ebp+fffffccc], 00000000
'0071c4b4    8b5594                  mov edx, dword ptr [ebp-6c]
'0071c4b7    899558fdffff            mov dword ptr [ebp+fffffd58], edx
'0071c4bd    c7459400000000          mov dword ptr [ebp-6c], 00000000
'0071c4c4    8b8558fdffff            mov eax, dword ptr [ebp+fffffd58]
'0071c4ca    898560ffffff            mov dword ptr [ebp+ffffff60], eax
'0071c4d0    c78558ffffff09000000    mov dword ptr [ebp+ffffff58], 00000009
'0071c4da    c78570ffffff00000000    mov dword ptr [ebp+ffffff70], 00000000
'0071c4e4    c78568ffffff02000000    mov dword ptr [ebp+ffffff68], 00000002
'0071c4ee    668b8de4fdffff          mov cx, word ptr [ebp+fffffde4]
'0071c4f5    66898d80feffff          mov word ptr [ebp+fffffe80], cx
'0071c4fc    c78578feffff0b000000    mov dword ptr [ebp+fffffe78], 0000000b
'0071c506    8d9558ffffff            lea edx, dword ptr [ebp+ffffff58]
'0071c50c    52                      push edx
'0071c50d    8d8568ffffff            lea eax, dword ptr [ebp+ffffff68]
'0071c513    50                      push eax
'0071c514    8d8d78feffff            lea ecx, dword ptr [ebp+fffffe78]
'0071c51a    51                      push ecx
'0071c51b    8d9548ffffff            lea edx, dword ptr [ebp+ffffff48]
'0071c521    52                      push edx

' *** Reference to "rtcImmediateIf"
'0071c522    ff154c124000            call dword ptr [0040124c]
'0071c528    8d9548ffffff            lea edx, dword ptr [ebp+ffffff48]
'0071c52e    8d4db4                  lea ecx, dword ptr [ebp-4c]

' *** Reference to "__vbaVarMove"
'0071c531    ff151c104000            call dword ptr [0040101c]
var_62 = (IIf(IsNull(var_51), var_19, var_121))
'0071c537    8d4598                  lea eax, dword ptr [ebp-68]
'0071c53a    50                      push eax
'0071c53b    8d4da0                  lea ecx, dword ptr [ebp-60]
'0071c53e    51                      push ecx
'0071c53f    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'0071c541    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_7, var_130)
'0071c547    83c40c                  add esp, 0c
'0071c54a    8d9558ffffff            lea edx, dword ptr [ebp+ffffff58]
'0071c550    52                      push edx
'0071c551    8d8568ffffff            lea eax, dword ptr [ebp+ffffff68]
'0071c557    50                      push eax
'0071c558    8d8d78feffff            lea ecx, dword ptr [ebp+fffffe78]
'0071c55e    51                      push ecx
'0071c55f    8d9578ffffff            lea edx, dword ptr [ebp+ffffff78]
'0071c565    52                      push edx
'0071c566    6a04                    push 04

' *** Reference to "__vbaFreeVarList"
'0071c568    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_87, var_73, var_132, var_133)
'0071c56e    83c414                  add esp, 14
'0071c571    c745fc12000000          mov dword ptr [ebp-04], 00000012
'0071c578    8d45a0                  lea eax, dword ptr [ebp-60]
'0071c57b    50                      push eax
'0071c57c    8b4dac                  mov ecx, dword ptr [ebp-54]
'0071c57f    8b11                    mov edx, dword ptr [ecx]
'0071c581    8b45ac                  mov eax, dword ptr [ebp-54]
'0071c584    50                      push eax
'0071c585    ff92b4000000            call dword ptr [edx+000000b4]
'0071c58b    dbe2                    fnclex
'0071c58d    8985e0fdffff            mov dword ptr [ebp+fffffde0], eax
'0071c593    83bde0fdffff00          cmp dword ptr [ebp+fffffde0], 00
'0071c59a    7d23                    jge 71c5bf

If (var_50 < 0) Then
'0071c59c    68b4000000              push 000000b4
'0071c5a1    6830314300              push 00433130
'0071c5a6    8b4dac                  mov ecx, dword ptr [ebp-54]
'0071c5a9    51                      push ecx
'0071c5aa    8b95e0fdffff            mov edx, dword ptr [ebp+fffffde0]
'0071c5b0    52                      push edx

' *** Reference to "__vbaHresultCheckObj"
'0071c5b1    ff1580104000            call dword ptr [00401080]
'0071c5b7    8985c8fcffff            mov dword ptr [ebp+fffffcc8], eax
'0071c5bd    eb0a                    jmp 71c5c9
    
End If
'0071c5bf    c785c8fcffff00000000    mov dword ptr [ebp+fffffcc8], 00000000
'0071c5c9    8b45a0                  mov eax, dword ptr [ebp-60]
'0071c5cc    8985dcfdffff            mov dword ptr [ebp+fffffddc], eax
'0071c5d2    c785a0feffff38784300    mov dword ptr [ebp+fffffea0], 00437838
'0071c5dc    c78598feffff08000000    mov dword ptr [ebp+fffffe98], 00000008
'0071c5e6    8d4d9c                  lea ecx, dword ptr [ebp-64]
'0071c5e9    51                      push ecx
'0071c5ea    b810000000              mov eax, 00000010

' *** Reference to "__vbaChkstk"
'0071c5ef    e86cacceff              call 407260
'0071c5f4    8bd4                    mov edx, esp
'0071c5f6    8b8598feffff            mov eax, dword ptr [ebp+fffffe98]
'0071c5fc    8902                    mov dword ptr [edx], eax
'0071c5fe    8b8d9cfeffff            mov ecx, dword ptr [ebp+fffffe9c]
'0071c604    894a04                  mov dword ptr [edx+04], ecx
'0071c607    8b85a0feffff            mov eax, dword ptr [ebp+fffffea0]
'0071c60d    894208                  mov dword ptr [edx+08], eax
'0071c610    8b8da4feffff            mov ecx, dword ptr [ebp+fffffea4]
'0071c616    894a0c                  mov dword ptr [edx+0c], ecx
'0071c619    8b95dcfdffff            mov edx, dword ptr [ebp+fffffddc]
'0071c61f    8b02                    mov eax, dword ptr [edx]
'0071c621    8b8ddcfdffff            mov ecx, dword ptr [ebp+fffffddc]
'0071c627    51                      push ecx
'0071c628    ff5030                  call dword ptr [eax+30]
'0071c62b    dbe2                    fnclex
'0071c62d    8985d8fdffff            mov dword ptr [ebp+fffffdd8], eax
'0071c633    83bdd8fdffff00          cmp dword ptr [ebp+fffffdd8], 00
'0071c63a    7d23                    jge 71c65f

If (-256 - 24 < 0) Then
'0071c63c    6a30                    push 30
'0071c63e    68d8304300              push 004330d8
'0071c643    8b95dcfdffff            mov edx, dword ptr [ebp+fffffddc]
'0071c649    52                      push edx
'0071c64a    8b85d8fdffff            mov eax, dword ptr [ebp+fffffdd8]
'0071c650    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0071c651    ff1580104000            call dword ptr [00401080]
'0071c657    8985c4fcffff            mov dword ptr [ebp+fffffcc4], eax
'0071c65d    eb0a                    jmp 71c669
    
End If
'0071c65f    c785c4fcffff00000000    mov dword ptr [ebp+fffffcc4], 00000000
'0071c669    8b4d9c                  mov ecx, dword ptr [ebp-64]
'0071c66c    898d54fdffff            mov dword ptr [ebp+fffffd54], ecx
'0071c672    c7459c00000000          mov dword ptr [ebp-64], 00000000
'0071c679    8b9554fdffff            mov edx, dword ptr [ebp+fffffd54]
'0071c67f    895580                  mov dword ptr [ebp-80], edx
'0071c682    c78578ffffff09000000    mov dword ptr [ebp+ffffff78], 00000009
'0071c68c    8d8578ffffff            lea eax, dword ptr [ebp+ffffff78]
'0071c692    50                      push eax

' *** Reference to "rtcIsNull"
'0071c693    ff1540114000            call dword ptr [00401140]
'0071c699    668985e4fdffff          mov word ptr [ebp+fffffde4], ax
'0071c6a0    8d4d98                  lea ecx, dword ptr [ebp-68]
'0071c6a3    51                      push ecx
'0071c6a4    8b55ac                  mov edx, dword ptr [ebp-54]
'0071c6a7    8b02                    mov eax, dword ptr [edx]
'0071c6a9    8b4dac                  mov ecx, dword ptr [ebp-54]
'0071c6ac    51                      push ecx
'0071c6ad    ff90b4000000            call dword ptr [eax+000000b4]
'0071c6b3    dbe2                    fnclex
'0071c6b5    8985d4fdffff            mov dword ptr [ebp+fffffdd4], eax
'0071c6bb    83bdd4fdffff00          cmp dword ptr [ebp+fffffdd4], 00
'0071c6c2    7d23                    jge 71c6e7

If (frmImport < 0) Then
'0071c6c4    68b4000000              push 000000b4
'0071c6c9    6830314300              push 00433130
'0071c6ce    8b55ac                  mov edx, dword ptr [ebp-54]
'0071c6d1    52                      push edx
'0071c6d2    8b85d4fdffff            mov eax, dword ptr [ebp+fffffdd4]
'0071c6d8    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0071c6d9    ff1580104000            call dword ptr [00401080]
'0071c6df    8985c0fcffff            mov dword ptr [ebp+fffffcc0], eax
'0071c6e5    eb0a                    jmp 71c6f1
    
End If
'0071c6e7    c785c0fcffff00000000    mov dword ptr [ebp+fffffcc0], 00000000
'0071c6f1    8b4d98                  mov ecx, dword ptr [ebp-68]
'0071c6f4    898dd0fdffff            mov dword ptr [ebp+fffffdd0], ecx
'0071c6fa    c78590feffff38784300    mov dword ptr [ebp+fffffe90], 00437838
'0071c704    c78588feffff08000000    mov dword ptr [ebp+fffffe88], 00000008
'0071c70e    8d5594                  lea edx, dword ptr [ebp-6c]
'0071c711    52                      push edx
'0071c712    b810000000              mov eax, 00000010

' *** Reference to "__vbaChkstk"
'0071c717    e844abceff              call 407260
'0071c71c    8bc4                    mov eax, esp
'0071c71e    8b8d88feffff            mov ecx, dword ptr [ebp+fffffe88]
'0071c724    8908                    mov dword ptr [eax], ecx
'0071c726    8b958cfeffff            mov edx, dword ptr [ebp+fffffe8c]
'0071c72c    895004                  mov dword ptr [eax+04], edx
'0071c72f    8b8d90feffff            mov ecx, dword ptr [ebp+fffffe90]
'0071c735    894808                  mov dword ptr [eax+08], ecx
'0071c738    8b9594feffff            mov edx, dword ptr [ebp+fffffe94]
'0071c73e    89500c                  mov dword ptr [eax+0c], edx
'0071c741    8b85d0fdffff            mov eax, dword ptr [ebp+fffffdd0]
'0071c747    8b08                    mov ecx, dword ptr [eax]
'0071c749    8b95d0fdffff            mov edx, dword ptr [ebp+fffffdd0]
'0071c74f    52                      push edx
'0071c750    ff5130                  call dword ptr [ecx+30]
'0071c753    dbe2                    fnclex
'0071c755    8985ccfdffff            mov dword ptr [ebp+fffffdcc], eax
'0071c75b    83bdccfdffff00          cmp dword ptr [ebp+fffffdcc], 00
'0071c762    7d23                    jge 71c787

If (0 < 0) Then
'0071c764    6a30                    push 30
'0071c766    68d8304300              push 004330d8
'0071c76b    8b85d0fdffff            mov eax, dword ptr [ebp+fffffdd0]
'0071c771    50                      push eax
'0071c772    8b8dccfdffff            mov ecx, dword ptr [ebp+fffffdcc]
'0071c778    51                      push ecx

' *** Reference to "__vbaHresultCheckObj"
'0071c779    ff1580104000            call dword ptr [00401080]
'0071c77f    8985bcfcffff            mov dword ptr [ebp+fffffcbc], eax
'0071c785    eb0a                    jmp 71c791
    
End If
'0071c787    c785bcfcffff00000000    mov dword ptr [ebp+fffffcbc], 00000000
'0071c791    8b5594                  mov edx, dword ptr [ebp-6c]
'0071c794    899550fdffff            mov dword ptr [ebp+fffffd50], edx
'0071c79a    c7459400000000          mov dword ptr [ebp-6c], 00000000
'0071c7a1    8b8550fdffff            mov eax, dword ptr [ebp+fffffd50]
'0071c7a7    898560ffffff            mov dword ptr [ebp+ffffff60], eax
'0071c7ad    c78558ffffff09000000    mov dword ptr [ebp+ffffff58], 00000009
'0071c7b7    c78570ffffff00000000    mov dword ptr [ebp+ffffff70], 00000000
'0071c7c1    c78568ffffff02000000    mov dword ptr [ebp+ffffff68], 00000002
'0071c7cb    668b8de4fdffff          mov cx, word ptr [ebp+fffffde4]
'0071c7d2    66898d80feffff          mov word ptr [ebp+fffffe80], cx
'0071c7d9    c78578feffff0b000000    mov dword ptr [ebp+fffffe78], 0000000b
'0071c7e3    8d9558ffffff            lea edx, dword ptr [ebp+ffffff58]
'0071c7e9    52                      push edx
'0071c7ea    8d8568ffffff            lea eax, dword ptr [ebp+ffffff68]
'0071c7f0    50                      push eax
'0071c7f1    8d8d78feffff            lea ecx, dword ptr [ebp+fffffe78]
'0071c7f7    51                      push ecx
'0071c7f8    8d9548ffffff            lea edx, dword ptr [ebp+ffffff48]
'0071c7fe    52                      push edx

' *** Reference to "rtcImmediateIf"
'0071c7ff    ff154c124000            call dword ptr [0040124c]
'0071c805    8d8548ffffff            lea eax, dword ptr [ebp+ffffff48]
'0071c80b    50                      push eax

' *** Reference to "__vbaI2Var"
'0071c80c    ff150c124000            call dword ptr [0040120c]
'0071c812    668945b0                mov word ptr [ebp-50], ax
'0071c816    8d4d98                  lea ecx, dword ptr [ebp-68]
'0071c819    51                      push ecx
'0071c81a    8d55a0                  lea edx, dword ptr [ebp-60]
'0071c81d    52                      push edx
'0071c81e    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'0071c820    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_7, var_130)
'0071c826    83c40c                  add esp, 0c
'0071c829    8d8548ffffff            lea eax, dword ptr [ebp+ffffff48]
'0071c82f    50                      push eax
'0071c830    8d8d58ffffff            lea ecx, dword ptr [ebp+ffffff58]
'0071c836    51                      push ecx
'0071c837    8d9568ffffff            lea edx, dword ptr [ebp+ffffff68]
'0071c83d    52                      push edx
'0071c83e    8d8578feffff            lea eax, dword ptr [ebp+fffffe78]
'0071c844    50                      push eax
'0071c845    8d8d78ffffff            lea ecx, dword ptr [ebp+ffffff78]
'0071c84b    51                      push ecx
'0071c84c    6a05                    push 05

' *** Reference to "__vbaFreeVarList"
'0071c84e    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_87, var_73, var_132, var_133, var_134)
'0071c854    83c418                  add esp, 18
'0071c857    c745fc13000000          mov dword ptr [ebp-04], 00000013
'0071c85e    8d55a0                  lea edx, dword ptr [ebp-60]
'0071c861    52                      push edx
'0071c862    8b45ac                  mov eax, dword ptr [ebp-54]
'0071c865    8b08                    mov ecx, dword ptr [eax]
'0071c867    8b55ac                  mov edx, dword ptr [ebp-54]
'0071c86a    52                      push edx
'0071c86b    ff91b4000000            call dword ptr [ecx+000000b4]
'0071c871    dbe2                    fnclex
'0071c873    8985e0fdffff            mov dword ptr [ebp+fffffde0], eax
'0071c879    83bde0fdffff00          cmp dword ptr [ebp+fffffde0], 00
'0071c880    7d23                    jge 71c8a5

If (var_50 < 0) Then
'0071c882    68b4000000              push 000000b4
'0071c887    6830314300              push 00433130
'0071c88c    8b45ac                  mov eax, dword ptr [ebp-54]
'0071c88f    50                      push eax
'0071c890    8b8de0fdffff            mov ecx, dword ptr [ebp+fffffde0]
'0071c896    51                      push ecx

' *** Reference to "__vbaHresultCheckObj"
'0071c897    ff1580104000            call dword ptr [00401080]
'0071c89d    8985b8fcffff            mov dword ptr [ebp+fffffcb8], eax
'0071c8a3    eb0a                    jmp 71c8af
    
End If
'0071c8a5    c785b8fcffff00000000    mov dword ptr [ebp+fffffcb8], 00000000
'0071c8af    8b55a0                  mov edx, dword ptr [ebp-60]
'0071c8b2    8995dcfdffff            mov dword ptr [ebp+fffffddc], edx
'0071c8b8    c785a0feffff1cd44300    mov dword ptr [ebp+fffffea0], 0043d41c
'0071c8c2    c78598feffff08000000    mov dword ptr [ebp+fffffe98], 00000008
'0071c8cc    8d459c                  lea eax, dword ptr [ebp-64]
'0071c8cf    50                      push eax
'0071c8d0    b810000000              mov eax, 00000010

' *** Reference to "__vbaChkstk"
'0071c8d5    e886a9ceff              call 407260
'0071c8da    8bcc                    mov ecx, esp
'0071c8dc    8b9598feffff            mov edx, dword ptr [ebp+fffffe98]
'0071c8e2    8911                    mov dword ptr [ecx], edx
'0071c8e4    8b859cfeffff            mov eax, dword ptr [ebp+fffffe9c]
'0071c8ea    894104                  mov dword ptr [ecx+04], eax
'0071c8ed    8b95a0feffff            mov edx, dword ptr [ebp+fffffea0]
'0071c8f3    895108                  mov dword ptr [ecx+08], edx
'0071c8f6    8b85a4feffff            mov eax, dword ptr [ebp+fffffea4]
'0071c8fc    89410c                  mov dword ptr [ecx+0c], eax
'0071c8ff    8b8ddcfdffff            mov ecx, dword ptr [ebp+fffffddc]
'0071c905    8b11                    mov edx, dword ptr [ecx]
'0071c907    8b85dcfdffff            mov eax, dword ptr [ebp+fffffddc]
'0071c90d    50                      push eax
'0071c90e    ff5230                  call dword ptr [edx+30]
'0071c911    dbe2                    fnclex
'0071c913    8985d8fdffff            mov dword ptr [ebp+fffffdd8], eax
'0071c919    83bdd8fdffff00          cmp dword ptr [ebp+fffffdd8], 00
'0071c920    7d23                    jge 71c945

If (var_7 < 0) Then
'0071c922    6a30                    push 30
'0071c924    68d8304300              push 004330d8
'0071c929    8b8ddcfdffff            mov ecx, dword ptr [ebp+fffffddc]
'0071c92f    51                      push ecx
'0071c930    8b95d8fdffff            mov edx, dword ptr [ebp+fffffdd8]
'0071c936    52                      push edx

' *** Reference to "__vbaHresultCheckObj"
'0071c937    ff1580104000            call dword ptr [00401080]
'0071c93d    8985b4fcffff            mov dword ptr [ebp+fffffcb4], eax
'0071c943    eb0a                    jmp 71c94f
    
End If
'0071c945    c785b4fcffff00000000    mov dword ptr [ebp+fffffcb4], 00000000
'0071c94f    8b459c                  mov eax, dword ptr [ebp-64]
'0071c952    8985d4fdffff            mov dword ptr [ebp+fffffdd4], eax
'0071c958    8d8d78ffffff            lea ecx, dword ptr [ebp+ffffff78]
'0071c95e    51                      push ecx
'0071c95f    8b95d4fdffff            mov edx, dword ptr [ebp+fffffdd4]
'0071c965    8b02                    mov eax, dword ptr [edx]
'0071c967    8b8dd4fdffff            mov ecx, dword ptr [ebp+fffffdd4]
'0071c96d    51                      push ecx
'0071c96e    ff5044                  call dword ptr [eax+44]
'0071c971    dbe2                    fnclex
'0071c973    8985d0fdffff            mov dword ptr [ebp+fffffdd0], eax
'0071c979    83bdd0fdffff00          cmp dword ptr [ebp+fffffdd0], 00
'0071c980    7d23                    jge 71c9a5

If (-256 - 24 < 0) Then
'0071c982    6a44                    push 44
'0071c984    6880304300              push 00433080
'0071c989    8b95d4fdffff            mov edx, dword ptr [ebp+fffffdd4]
'0071c98f    52                      push edx
'0071c990    8b85d0fdffff            mov eax, dword ptr [ebp+fffffdd0]
'0071c996    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0071c997    ff1580104000            call dword ptr [00401080]
'0071c99d    8985b0fcffff            mov dword ptr [ebp+fffffcb0], eax
'0071c9a3    eb0a                    jmp 71c9af
    
End If
'0071c9a5    c785b0fcffff00000000    mov dword ptr [ebp+fffffcb0], 00000000
'0071c9af    8d8d78ffffff            lea ecx, dword ptr [ebp+ffffff78]
'0071c9b5    51                      push ecx

' *** Reference to "__vbaStrVarMove"
'0071c9b6    ff153c104000            call dword ptr [0040103c]
'0071c9bc    8bd0                    mov edx, eax
'0071c9be    8d4dd4                  lea ecx, dword ptr [ebp-2c]

' *** Reference to "__vbaStrMove"
'0071c9c1    ff15d0124000            call dword ptr [004012d0]
'0071c9c7    8d559c                  lea edx, dword ptr [ebp-64]
'0071c9ca    52                      push edx
'0071c9cb    8d45a0                  lea eax, dword ptr [ebp-60]
'0071c9ce    50                      push eax
'0071c9cf    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'0071c9d1    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_7, var_51)
'0071c9d7    83c40c                  add esp, 0c
'0071c9da    8d8d78ffffff            lea ecx, dword ptr [ebp+ffffff78]

' *** Reference to "__vbaFreeVar"
'0071c9e0    ff1528104000            call dword ptr [00401028]
'#Cleanup(var_87)
'0071c9e6    c745fc14000000          mov dword ptr [ebp-04], 00000014
'0071c9ed    8d4da0                  lea ecx, dword ptr [ebp-60]
'0071c9f0    51                      push ecx
'0071c9f1    8b55ac                  mov edx, dword ptr [ebp-54]
'0071c9f4    8b02                    mov eax, dword ptr [edx]
'0071c9f6    8b4dac                  mov ecx, dword ptr [ebp-54]
'0071c9f9    51                      push ecx
'0071c9fa    ff90b4000000            call dword ptr [eax+000000b4]
'0071ca00    dbe2                    fnclex
'0071ca02    8985e0fdffff            mov dword ptr [ebp+fffffde0], eax
'0071ca08    83bde0fdffff00          cmp dword ptr [ebp+fffffde0], 00
'0071ca0f    7d23                    jge 71ca34

If (frmImport < 0) Then
'0071ca11    68b4000000              push 000000b4
'0071ca16    6830314300              push 00433130
'0071ca1b    8b55ac                  mov edx, dword ptr [ebp-54]
'0071ca1e    52                      push edx
'0071ca1f    8b85e0fdffff            mov eax, dword ptr [ebp+fffffde0]
'0071ca25    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0071ca26    ff1580104000            call dword ptr [00401080]
'0071ca2c    8985acfcffff            mov dword ptr [ebp+fffffcac], eax
'0071ca32    eb0a                    jmp 71ca3e
    
End If
'0071ca34    c785acfcffff00000000    mov dword ptr [ebp+fffffcac], 00000000
'0071ca3e    8b4da0                  mov ecx, dword ptr [ebp-60]
'0071ca41    898ddcfdffff            mov dword ptr [ebp+fffffddc], ecx
'0071ca47    c785a0feffff34d44300    mov dword ptr [ebp+fffffea0], 0043d434
'0071ca51    c78598feffff08000000    mov dword ptr [ebp+fffffe98], 00000008
'0071ca5b    8d559c                  lea edx, dword ptr [ebp-64]
'0071ca5e    52                      push edx
'0071ca5f    b810000000              mov eax, 00000010

' *** Reference to "__vbaChkstk"
'0071ca64    e8f7a7ceff              call 407260
'0071ca69    8bc4                    mov eax, esp
'0071ca6b    8b8d98feffff            mov ecx, dword ptr [ebp+fffffe98]
'0071ca71    8908                    mov dword ptr [eax], ecx
'0071ca73    8b959cfeffff            mov edx, dword ptr [ebp+fffffe9c]
'0071ca79    895004                  mov dword ptr [eax+04], edx
'0071ca7c    8b8da0feffff            mov ecx, dword ptr [ebp+fffffea0]
'0071ca82    894808                  mov dword ptr [eax+08], ecx
'0071ca85    8b95a4feffff            mov edx, dword ptr [ebp+fffffea4]
'0071ca8b    89500c                  mov dword ptr [eax+0c], edx
'0071ca8e    8b85dcfdffff            mov eax, dword ptr [ebp+fffffddc]
'0071ca94    8b08                    mov ecx, dword ptr [eax]
'0071ca96    8b95dcfdffff            mov edx, dword ptr [ebp+fffffddc]
'0071ca9c    52                      push edx
'0071ca9d    ff5130                  call dword ptr [ecx+30]
'0071caa0    dbe2                    fnclex
'0071caa2    8985d8fdffff            mov dword ptr [ebp+fffffdd8], eax
'0071caa8    83bdd8fdffff00          cmp dword ptr [ebp+fffffdd8], 00
'0071caaf    7d23                    jge 71cad4

If (var_7 < 0) Then
'0071cab1    6a30                    push 30
'0071cab3    68d8304300              push 004330d8
'0071cab8    8b85dcfdffff            mov eax, dword ptr [ebp+fffffddc]
'0071cabe    50                      push eax
'0071cabf    8b8dd8fdffff            mov ecx, dword ptr [ebp+fffffdd8]
'0071cac5    51                      push ecx

' *** Reference to "__vbaHresultCheckObj"
'0071cac6    ff1580104000            call dword ptr [00401080]
'0071cacc    8985a8fcffff            mov dword ptr [ebp+fffffca8], eax
'0071cad2    eb0a                    jmp 71cade
    
End If
'0071cad4    c785a8fcffff00000000    mov dword ptr [ebp+fffffca8], 00000000
'0071cade    8b559c                  mov edx, dword ptr [ebp-64]
'0071cae1    8995d4fdffff            mov dword ptr [ebp+fffffdd4], edx
'0071cae7    8d8578ffffff            lea eax, dword ptr [ebp+ffffff78]
'0071caed    50                      push eax
'0071caee    8b8dd4fdffff            mov ecx, dword ptr [ebp+fffffdd4]
'0071caf4    8b11                    mov edx, dword ptr [ecx]
'0071caf6    8b85d4fdffff            mov eax, dword ptr [ebp+fffffdd4]
'0071cafc    50                      push eax
'0071cafd    ff5244                  call dword ptr [edx+44]
'0071cb00    dbe2                    fnclex
'0071cb02    8985d0fdffff            mov dword ptr [ebp+fffffdd0], eax
'0071cb08    83bdd0fdffff00          cmp dword ptr [ebp+fffffdd0], 00
'0071cb0f    7d23                    jge 71cb34

If (var_51 < 0) Then
'0071cb11    6a44                    push 44
'0071cb13    6880304300              push 00433080
'0071cb18    8b8dd4fdffff            mov ecx, dword ptr [ebp+fffffdd4]
'0071cb1e    51                      push ecx
'0071cb1f    8b95d0fdffff            mov edx, dword ptr [ebp+fffffdd0]
'0071cb25    52                      push edx

' *** Reference to "__vbaHresultCheckObj"
'0071cb26    ff1580104000            call dword ptr [00401080]
'0071cb2c    8985a4fcffff            mov dword ptr [ebp+fffffca4], eax
'0071cb32    eb0a                    jmp 71cb3e
    
End If
'0071cb34    c785a4fcffff00000000    mov dword ptr [ebp+fffffca4], 00000000
'0071cb3e    8d8578ffffff            lea eax, dword ptr [ebp+ffffff78]
'0071cb44    50                      push eax

' *** Reference to "__vbaI2Var"
'0071cb45    ff150c124000            call dword ptr [0040120c]
'0071cb4b    668945cc                mov word ptr [ebp-34], ax
'0071cb4f    8d4d9c                  lea ecx, dword ptr [ebp-64]
'0071cb52    51                      push ecx
'0071cb53    8d55a0                  lea edx, dword ptr [ebp-60]
'0071cb56    52                      push edx
'0071cb57    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'0071cb59    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_7, var_51)
'0071cb5f    83c40c                  add esp, 0c
'0071cb62    8d8d78ffffff            lea ecx, dword ptr [ebp+ffffff78]

' *** Reference to "__vbaFreeVar"
'0071cb68    ff1528104000            call dword ptr [00401028]
'#Cleanup(var_87)
'0071cb6e    c745fc15000000          mov dword ptr [ebp-04], 00000015
'0071cb75    668b45a4                mov ax, word ptr [ebp-5c]
'0071cb79    663b45cc                cmp ax, word ptr [ebp-34]
'0071cb7d    0f8e9e010000            jle 71cd21

If (CInt(IIf(IsNull(var_51), var_19, var_121)) > WORD PTR [EBP-34]) Then
'0071cb83    c745fc16000000          mov dword ptr [ebp-04], 00000016
'0071cb8a    8d4da0                  lea ecx, dword ptr [ebp-60]
'0071cb8d    51                      push ecx
'0071cb8e    8b55ac                  mov edx, dword ptr [ebp-54]
'0071cb91    8b02                    mov eax, dword ptr [edx]
'0071cb93    8b4dac                  mov ecx, dword ptr [ebp-54]
'0071cb96    51                      push ecx
'0071cb97    ff90b4000000            call dword ptr [eax+000000b4]
'0071cb9d    dbe2                    fnclex
'0071cb9f    8985e0fdffff            mov dword ptr [ebp+fffffde0], eax
'0071cba5    83bde0fdffff00          cmp dword ptr [ebp+fffffde0], 00
'0071cbac    7d23                    jge 71cbd1
    
    If (    frmImport < 0) Then
'0071cbae    68b4000000              push 000000b4
'0071cbb3    6830314300              push 00433130
'0071cbb8    8b55ac                  mov edx, dword ptr [ebp-54]
'0071cbbb    52                      push edx
'0071cbbc    8b85e0fdffff            mov eax, dword ptr [ebp+fffffde0]
'0071cbc2    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0071cbc3    ff1580104000            call dword ptr [00401080]
'0071cbc9    8985a0fcffff            mov dword ptr [ebp+fffffca0], eax
'0071cbcf    eb0a                    jmp 71cbdb
    
End If
'0071cbd1    c785a0fcffff00000000    mov dword ptr [ebp+fffffca0], 00000000
'0071cbdb    8b4da0                  mov ecx, dword ptr [ebp-60]
'0071cbde    898ddcfdffff            mov dword ptr [ebp+fffffddc], ecx
'0071cbe4    c785a0feffff44d44300    mov dword ptr [ebp+fffffea0], 0043d444
'0071cbee    c78598feffff08000000    mov dword ptr [ebp+fffffe98], 00000008
'0071cbf8    8d559c                  lea edx, dword ptr [ebp-64]
'0071cbfb    52                      push edx
'0071cbfc    b810000000              mov eax, 00000010

' *** Reference to "__vbaChkstk"
'0071cc01    e85aa6ceff              call 407260
'0071cc06    8bc4                    mov eax, esp
'0071cc08    8b8d98feffff            mov ecx, dword ptr [ebp+fffffe98]
'0071cc0e    8908                    mov dword ptr [eax], ecx
'0071cc10    8b959cfeffff            mov edx, dword ptr [ebp+fffffe9c]
'0071cc16    895004                  mov dword ptr [eax+04], edx
'0071cc19    8b8da0feffff            mov ecx, dword ptr [ebp+fffffea0]
'0071cc1f    894808                  mov dword ptr [eax+08], ecx
'0071cc22    8b95a4feffff            mov edx, dword ptr [ebp+fffffea4]
'0071cc28    89500c                  mov dword ptr [eax+0c], edx
'0071cc2b    8b85dcfdffff            mov eax, dword ptr [ebp+fffffddc]
'0071cc31    8b08                    mov ecx, dword ptr [eax]
'0071cc33    8b95dcfdffff            mov edx, dword ptr [ebp+fffffddc]
'0071cc39    52                      push edx
'0071cc3a    ff5130                  call dword ptr [ecx+30]
'0071cc3d    dbe2                    fnclex
'0071cc3f    8985d8fdffff            mov dword ptr [ebp+fffffdd8], eax
'0071cc45    83bdd8fdffff00          cmp dword ptr [ebp+fffffdd8], 00
'0071cc4c    7d23                    jge 71cc71

If (var_7 < 0) Then
'0071cc4e    6a30                    push 30
'0071cc50    68d8304300              push 004330d8
'0071cc55    8b85dcfdffff            mov eax, dword ptr [ebp+fffffddc]
'0071cc5b    50                      push eax
'0071cc5c    8b8dd8fdffff            mov ecx, dword ptr [ebp+fffffdd8]
'0071cc62    51                      push ecx

' *** Reference to "__vbaHresultCheckObj"
'0071cc63    ff1580104000            call dword ptr [00401080]
'0071cc69    89859cfcffff            mov dword ptr [ebp+fffffc9c], eax
'0071cc6f    eb0a                    jmp 71cc7b
    
End If
'0071cc71    c7859cfcffff00000000    mov dword ptr [ebp+fffffc9c], 00000000
'0071cc7b    8b559c                  mov edx, dword ptr [ebp-64]
'0071cc7e    8995d4fdffff            mov dword ptr [ebp+fffffdd4], edx
'0071cc84    8d8578ffffff            lea eax, dword ptr [ebp+ffffff78]
'0071cc8a    50                      push eax
'0071cc8b    8b8dd4fdffff            mov ecx, dword ptr [ebp+fffffdd4]
'0071cc91    8b11                    mov edx, dword ptr [ecx]
'0071cc93    8b85d4fdffff            mov eax, dword ptr [ebp+fffffdd4]
'0071cc99    50                      push eax
'0071cc9a    ff5244                  call dword ptr [edx+44]
'0071cc9d    dbe2                    fnclex
'0071cc9f    8985d0fdffff            mov dword ptr [ebp+fffffdd0], eax
'0071cca5    83bdd0fdffff00          cmp dword ptr [ebp+fffffdd0], 00
'0071ccac    7d23                    jge 71ccd1

If (var_51 < 0) Then
'0071ccae    6a44                    push 44
'0071ccb0    6880304300              push 00433080
'0071ccb5    8b8dd4fdffff            mov ecx, dword ptr [ebp+fffffdd4]
'0071ccbb    51                      push ecx
'0071ccbc    8b95d0fdffff            mov edx, dword ptr [ebp+fffffdd0]
'0071ccc2    52                      push edx

' *** Reference to "__vbaHresultCheckObj"
'0071ccc3    ff1580104000            call dword ptr [00401080]
'0071ccc9    898598fcffff            mov dword ptr [ebp+fffffc98], eax
'0071cccf    eb0a                    jmp 71ccdb
    
End If
'0071ccd1    c78598fcffff00000000    mov dword ptr [ebp+fffffc98], 00000000
'0071ccdb    8d8578ffffff            lea eax, dword ptr [ebp+ffffff78]
'0071cce1    50                      push eax

' *** Reference to "__vbaStrVarMove"
'0071cce2    ff153c104000            call dword ptr [0040103c]
'0071cce8    8bd0                    mov edx, eax
'0071ccea    8d4dd4                  lea ecx, dword ptr [ebp-2c]

' *** Reference to "__vbaStrMove"
'0071cced    ff15d0124000            call dword ptr [004012d0]
'0071ccf3    8d4d9c                  lea ecx, dword ptr [ebp-64]
'0071ccf6    51                      push ecx
'0071ccf7    8d55a0                  lea edx, dword ptr [ebp-60]
'0071ccfa    52                      push edx
'0071ccfb    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'0071ccfd    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_7, var_51)
'0071cd03    83c40c                  add esp, 0c
'0071cd06    8d8d78ffffff            lea ecx, dword ptr [ebp+ffffff78]

' *** Reference to "__vbaFreeVar"
'0071cd0c    ff1528104000            call dword ptr [00401028]
'#Cleanup(var_87)
'0071cd12    c745fc17000000          mov dword ptr [ebp-04], 00000017
'0071cd19    668b45a4                mov ax, word ptr [ebp-5c]
'0071cd1d    668945cc                mov word ptr [ebp-34], ax

'ERROR: Two many next close:
End If
'0071cd21    c745fc19000000          mov dword ptr [ebp-04], 00000019
'0071cd28    668b4dd8                mov cx, word ptr [ebp-28]
'0071cd2c    663b4dcc                cmp cx, word ptr [ebp-34]
'0071cd30    0f8e9e010000            jle 71ced4

If (CInt(IIf(IsNull(var_51), var_19, var_121)) > WORD PTR [EBP-34]) Then
'0071cd36    c745fc1a000000          mov dword ptr [ebp-04], 0000001a
'0071cd3d    8d55a0                  lea edx, dword ptr [ebp-60]
'0071cd40    52                      push edx
'0071cd41    8b45ac                  mov eax, dword ptr [ebp-54]
'0071cd44    8b08                    mov ecx, dword ptr [eax]
'0071cd46    8b55ac                  mov edx, dword ptr [ebp-54]
'0071cd49    52                      push edx
'0071cd4a    ff91b4000000            call dword ptr [ecx+000000b4]
'0071cd50    dbe2                    fnclex
'0071cd52    8985e0fdffff            mov dword ptr [ebp+fffffde0], eax
'0071cd58    83bde0fdffff00          cmp dword ptr [ebp+fffffde0], 00
'0071cd5f    7d23                    jge 71cd84
    
    If (    var_50 < 0) Then
'0071cd61    68b4000000              push 000000b4
'0071cd66    6830314300              push 00433130
'0071cd6b    8b45ac                  mov eax, dword ptr [ebp-54]
'0071cd6e    50                      push eax
'0071cd6f    8b8de0fdffff            mov ecx, dword ptr [ebp+fffffde0]
'0071cd75    51                      push ecx

' *** Reference to "__vbaHresultCheckObj"
'0071cd76    ff1580104000            call dword ptr [00401080]
'0071cd7c    898594fcffff            mov dword ptr [ebp+fffffc94], eax
'0071cd82    eb0a                    jmp 71cd8e
    
End If
'0071cd84    c78594fcffff00000000    mov dword ptr [ebp+fffffc94], 00000000
'0071cd8e    8b55a0                  mov edx, dword ptr [ebp-60]
'0071cd91    8995dcfdffff            mov dword ptr [ebp+fffffddc], edx
'0071cd97    c785a0feffff6cd44300    mov dword ptr [ebp+fffffea0], 0043d46c
'0071cda1    c78598feffff08000000    mov dword ptr [ebp+fffffe98], 00000008
'0071cdab    8d459c                  lea eax, dword ptr [ebp-64]
'0071cdae    50                      push eax
'0071cdaf    b810000000              mov eax, 00000010

' *** Reference to "__vbaChkstk"
'0071cdb4    e8a7a4ceff              call 407260
'0071cdb9    8bcc                    mov ecx, esp
'0071cdbb    8b9598feffff            mov edx, dword ptr [ebp+fffffe98]
'0071cdc1    8911                    mov dword ptr [ecx], edx
'0071cdc3    8b859cfeffff            mov eax, dword ptr [ebp+fffffe9c]
'0071cdc9    894104                  mov dword ptr [ecx+04], eax
'0071cdcc    8b95a0feffff            mov edx, dword ptr [ebp+fffffea0]
'0071cdd2    895108                  mov dword ptr [ecx+08], edx
'0071cdd5    8b85a4feffff            mov eax, dword ptr [ebp+fffffea4]
'0071cddb    89410c                  mov dword ptr [ecx+0c], eax
'0071cdde    8b8ddcfdffff            mov ecx, dword ptr [ebp+fffffddc]
'0071cde4    8b11                    mov edx, dword ptr [ecx]
'0071cde6    8b85dcfdffff            mov eax, dword ptr [ebp+fffffddc]
'0071cdec    50                      push eax
'0071cded    ff5230                  call dword ptr [edx+30]
'0071cdf0    dbe2                    fnclex
'0071cdf2    8985d8fdffff            mov dword ptr [ebp+fffffdd8], eax
'0071cdf8    83bdd8fdffff00          cmp dword ptr [ebp+fffffdd8], 00
'0071cdff    7d23                    jge 71ce24

If (var_7 < 0) Then
'0071ce01    6a30                    push 30
'0071ce03    68d8304300              push 004330d8
'0071ce08    8b8ddcfdffff            mov ecx, dword ptr [ebp+fffffddc]
'0071ce0e    51                      push ecx
'0071ce0f    8b95d8fdffff            mov edx, dword ptr [ebp+fffffdd8]
'0071ce15    52                      push edx

' *** Reference to "__vbaHresultCheckObj"
'0071ce16    ff1580104000            call dword ptr [00401080]
'0071ce1c    898590fcffff            mov dword ptr [ebp+fffffc90], eax
'0071ce22    eb0a                    jmp 71ce2e
    
End If
'0071ce24    c78590fcffff00000000    mov dword ptr [ebp+fffffc90], 00000000
'0071ce2e    8b459c                  mov eax, dword ptr [ebp-64]
'0071ce31    8985d4fdffff            mov dword ptr [ebp+fffffdd4], eax
'0071ce37    8d8d78ffffff            lea ecx, dword ptr [ebp+ffffff78]
'0071ce3d    51                      push ecx
'0071ce3e    8b95d4fdffff            mov edx, dword ptr [ebp+fffffdd4]
'0071ce44    8b02                    mov eax, dword ptr [edx]
'0071ce46    8b8dd4fdffff            mov ecx, dword ptr [ebp+fffffdd4]
'0071ce4c    51                      push ecx
'0071ce4d    ff5044                  call dword ptr [eax+44]
'0071ce50    dbe2                    fnclex
'0071ce52    8985d0fdffff            mov dword ptr [ebp+fffffdd0], eax
'0071ce58    83bdd0fdffff00          cmp dword ptr [ebp+fffffdd0], 00
'0071ce5f    7d23                    jge 71ce84

If (-256 - 24 < 0) Then
'0071ce61    6a44                    push 44
'0071ce63    6880304300              push 00433080
'0071ce68    8b95d4fdffff            mov edx, dword ptr [ebp+fffffdd4]
'0071ce6e    52                      push edx
'0071ce6f    8b85d0fdffff            mov eax, dword ptr [ebp+fffffdd0]
'0071ce75    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0071ce76    ff1580104000            call dword ptr [00401080]
'0071ce7c    89858cfcffff            mov dword ptr [ebp+fffffc8c], eax
'0071ce82    eb0a                    jmp 71ce8e
    
End If
'0071ce84    c7858cfcffff00000000    mov dword ptr [ebp+fffffc8c], 00000000
'0071ce8e    8d8d78ffffff            lea ecx, dword ptr [ebp+ffffff78]
'0071ce94    51                      push ecx

' *** Reference to "__vbaStrVarMove"
'0071ce95    ff153c104000            call dword ptr [0040103c]
'0071ce9b    8bd0                    mov edx, eax
'0071ce9d    8d4dd4                  lea ecx, dword ptr [ebp-2c]

' *** Reference to "__vbaStrMove"
'0071cea0    ff15d0124000            call dword ptr [004012d0]
'0071cea6    8d559c                  lea edx, dword ptr [ebp-64]
'0071cea9    52                      push edx
'0071ceaa    8d45a0                  lea eax, dword ptr [ebp-60]
'0071cead    50                      push eax
'0071ceae    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'0071ceb0    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_7, var_51)
'0071ceb6    83c40c                  add esp, 0c
'0071ceb9    8d8d78ffffff            lea ecx, dword ptr [ebp+ffffff78]

' *** Reference to "__vbaFreeVar"
'0071cebf    ff1528104000            call dword ptr [00401028]
'#Cleanup(var_87)
'0071cec5    c745fc1b000000          mov dword ptr [ebp-04], 0000001b
'0071cecc    668b4dd8                mov cx, word ptr [ebp-28]
'0071ced0    66894dcc                mov word ptr [ebp-34], cx

'ERROR: Two many next close:
End If
'0071ced4    c745fc1d000000          mov dword ptr [ebp-04], 0000001d
'0071cedb    668b55d0                mov dx, word ptr [ebp-30]
'0071cedf    663b55cc                cmp dx, word ptr [ebp-34]
'0071cee3    0f8e9e010000            jle 71d087

If (CInt(IIf(IsNull(var_51), var_19, var_121)) > WORD PTR [EBP-34]) Then
'0071cee9    c745fc1e000000          mov dword ptr [ebp-04], 0000001e
'0071cef0    8d45a0                  lea eax, dword ptr [ebp-60]
'0071cef3    50                      push eax
'0071cef4    8b4dac                  mov ecx, dword ptr [ebp-54]
'0071cef7    8b11                    mov edx, dword ptr [ecx]
'0071cef9    8b45ac                  mov eax, dword ptr [ebp-54]
'0071cefc    50                      push eax
'0071cefd    ff92b4000000            call dword ptr [edx+000000b4]
'0071cf03    dbe2                    fnclex
'0071cf05    8985e0fdffff            mov dword ptr [ebp+fffffde0], eax
'0071cf0b    83bde0fdffff00          cmp dword ptr [ebp+fffffde0], 00
'0071cf12    7d23                    jge 71cf37
    
    If (    var_50 < 0) Then
'0071cf14    68b4000000              push 000000b4
'0071cf19    6830314300              push 00433130
'0071cf1e    8b4dac                  mov ecx, dword ptr [ebp-54]
'0071cf21    51                      push ecx
'0071cf22    8b95e0fdffff            mov edx, dword ptr [ebp+fffffde0]
'0071cf28    52                      push edx

' *** Reference to "__vbaHresultCheckObj"
'0071cf29    ff1580104000            call dword ptr [00401080]
'0071cf2f    898588fcffff            mov dword ptr [ebp+fffffc88], eax
'0071cf35    eb0a                    jmp 71cf41
    
End If
'0071cf37    c78588fcffff00000000    mov dword ptr [ebp+fffffc88], 00000000
'0071cf41    8b45a0                  mov eax, dword ptr [ebp-60]
'0071cf44    8985dcfdffff            mov dword ptr [ebp+fffffddc], eax
'0071cf4a    c785a0feffff0cad4300    mov dword ptr [ebp+fffffea0], 0043ad0c
'0071cf54    c78598feffff08000000    mov dword ptr [ebp+fffffe98], 00000008
'0071cf5e    8d4d9c                  lea ecx, dword ptr [ebp-64]
'0071cf61    51                      push ecx
'0071cf62    b810000000              mov eax, 00000010

' *** Reference to "__vbaChkstk"
'0071cf67    e8f4a2ceff              call 407260
'0071cf6c    8bd4                    mov edx, esp
'0071cf6e    8b8598feffff            mov eax, dword ptr [ebp+fffffe98]
'0071cf74    8902                    mov dword ptr [edx], eax
'0071cf76    8b8d9cfeffff            mov ecx, dword ptr [ebp+fffffe9c]
'0071cf7c    894a04                  mov dword ptr [edx+04], ecx
'0071cf7f    8b85a0feffff            mov eax, dword ptr [ebp+fffffea0]
'0071cf85    894208                  mov dword ptr [edx+08], eax
'0071cf88    8b8da4feffff            mov ecx, dword ptr [ebp+fffffea4]
'0071cf8e    894a0c                  mov dword ptr [edx+0c], ecx
'0071cf91    8b95dcfdffff            mov edx, dword ptr [ebp+fffffddc]
'0071cf97    8b02                    mov eax, dword ptr [edx]
'0071cf99    8b8ddcfdffff            mov ecx, dword ptr [ebp+fffffddc]
'0071cf9f    51                      push ecx
'0071cfa0    ff5030                  call dword ptr [eax+30]
'0071cfa3    dbe2                    fnclex
'0071cfa5    8985d8fdffff            mov dword ptr [ebp+fffffdd8], eax
'0071cfab    83bdd8fdffff00          cmp dword ptr [ebp+fffffdd8], 00
'0071cfb2    7d23                    jge 71cfd7

If (-256 - 24 < 0) Then
'0071cfb4    6a30                    push 30
'0071cfb6    68d8304300              push 004330d8
'0071cfbb    8b95dcfdffff            mov edx, dword ptr [ebp+fffffddc]
'0071cfc1    52                      push edx
'0071cfc2    8b85d8fdffff            mov eax, dword ptr [ebp+fffffdd8]
'0071cfc8    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0071cfc9    ff1580104000            call dword ptr [00401080]
'0071cfcf    898584fcffff            mov dword ptr [ebp+fffffc84], eax
'0071cfd5    eb0a                    jmp 71cfe1
    
End If
'0071cfd7    c78584fcffff00000000    mov dword ptr [ebp+fffffc84], 00000000
'0071cfe1    8b4d9c                  mov ecx, dword ptr [ebp-64]
'0071cfe4    898dd4fdffff            mov dword ptr [ebp+fffffdd4], ecx
'0071cfea    8d9578ffffff            lea edx, dword ptr [ebp+ffffff78]
'0071cff0    52                      push edx
'0071cff1    8b85d4fdffff            mov eax, dword ptr [ebp+fffffdd4]
'0071cff7    8b08                    mov ecx, dword ptr [eax]
'0071cff9    8b95d4fdffff            mov edx, dword ptr [ebp+fffffdd4]
'0071cfff    52                      push edx
'0071d000    ff5144                  call dword ptr [ecx+44]
'0071d003    dbe2                    fnclex
'0071d005    8985d0fdffff            mov dword ptr [ebp+fffffdd0], eax
'0071d00b    83bdd0fdffff00          cmp dword ptr [ebp+fffffdd0], 00
'0071d012    7d23                    jge 71d037

If (var_51 < 0) Then
'0071d014    6a44                    push 44
'0071d016    6880304300              push 00433080
'0071d01b    8b85d4fdffff            mov eax, dword ptr [ebp+fffffdd4]
'0071d021    50                      push eax
'0071d022    8b8dd0fdffff            mov ecx, dword ptr [ebp+fffffdd0]
'0071d028    51                      push ecx

' *** Reference to "__vbaHresultCheckObj"
'0071d029    ff1580104000            call dword ptr [00401080]
'0071d02f    898580fcffff            mov dword ptr [ebp+fffffc80], eax
'0071d035    eb0a                    jmp 71d041
    
End If
'0071d037    c78580fcffff00000000    mov dword ptr [ebp+fffffc80], 00000000
'0071d041    8d9578ffffff            lea edx, dword ptr [ebp+ffffff78]
'0071d047    52                      push edx

' *** Reference to "__vbaStrVarMove"
'0071d048    ff153c104000            call dword ptr [0040103c]
'0071d04e    8bd0                    mov edx, eax
'0071d050    8d4dd4                  lea ecx, dword ptr [ebp-2c]

' *** Reference to "__vbaStrMove"
'0071d053    ff15d0124000            call dword ptr [004012d0]
'0071d059    8d459c                  lea eax, dword ptr [ebp-64]
'0071d05c    50                      push eax
'0071d05d    8d4da0                  lea ecx, dword ptr [ebp-60]
'0071d060    51                      push ecx
'0071d061    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'0071d063    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_7, var_51)
'0071d069    83c40c                  add esp, 0c
'0071d06c    8d8d78ffffff            lea ecx, dword ptr [ebp+ffffff78]

' *** Reference to "__vbaFreeVar"
'0071d072    ff1528104000            call dword ptr [00401028]
'#Cleanup(var_87)
'0071d078    c745fc1f000000          mov dword ptr [ebp-04], 0000001f
'0071d07f    668b55d0                mov dx, word ptr [ebp-30]
'0071d083    668955cc                mov word ptr [ebp-34], dx

'ERROR: Two many next close:
End If
'0071d087    c745fc21000000          mov dword ptr [ebp-04], 00000021
'0071d08e    668b45c8                mov ax, word ptr [ebp-38]
'0071d092    663b45cc                cmp ax, word ptr [ebp-34]
'0071d096    0f8e9e010000            jle 71d23a

If (CInt(IIf(IsNull(var_51), var_19, var_121)) > WORD PTR [EBP-34]) Then
'0071d09c    c745fc22000000          mov dword ptr [ebp-04], 00000022
'0071d0a3    8d4da0                  lea ecx, dword ptr [ebp-60]
'0071d0a6    51                      push ecx
'0071d0a7    8b55ac                  mov edx, dword ptr [ebp-54]
'0071d0aa    8b02                    mov eax, dword ptr [edx]
'0071d0ac    8b4dac                  mov ecx, dword ptr [ebp-54]
'0071d0af    51                      push ecx
'0071d0b0    ff90b4000000            call dword ptr [eax+000000b4]
'0071d0b6    dbe2                    fnclex
'0071d0b8    8985e0fdffff            mov dword ptr [ebp+fffffde0], eax
'0071d0be    83bde0fdffff00          cmp dword ptr [ebp+fffffde0], 00
'0071d0c5    7d23                    jge 71d0ea
    
    If (    frmImport < 0) Then
'0071d0c7    68b4000000              push 000000b4
'0071d0cc    6830314300              push 00433130
'0071d0d1    8b55ac                  mov edx, dword ptr [ebp-54]
'0071d0d4    52                      push edx
'0071d0d5    8b85e0fdffff            mov eax, dword ptr [ebp+fffffde0]
'0071d0db    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0071d0dc    ff1580104000            call dword ptr [00401080]
'0071d0e2    89857cfcffff            mov dword ptr [ebp+fffffc7c], eax
'0071d0e8    eb0a                    jmp 71d0f4
    
End If
'0071d0ea    c7857cfcffff00000000    mov dword ptr [ebp+fffffc7c], 00000000
'0071d0f4    8b4da0                  mov ecx, dword ptr [ebp-60]
'0071d0f7    898ddcfdffff            mov dword ptr [ebp+fffffddc], ecx
'0071d0fd    c785a0feffff24ad4300    mov dword ptr [ebp+fffffea0], 0043ad24
'0071d107    c78598feffff08000000    mov dword ptr [ebp+fffffe98], 00000008
'0071d111    8d559c                  lea edx, dword ptr [ebp-64]
'0071d114    52                      push edx
'0071d115    b810000000              mov eax, 00000010

' *** Reference to "__vbaChkstk"
'0071d11a    e841a1ceff              call 407260
'0071d11f    8bc4                    mov eax, esp
'0071d121    8b8d98feffff            mov ecx, dword ptr [ebp+fffffe98]
'0071d127    8908                    mov dword ptr [eax], ecx
'0071d129    8b959cfeffff            mov edx, dword ptr [ebp+fffffe9c]
'0071d12f    895004                  mov dword ptr [eax+04], edx
'0071d132    8b8da0feffff            mov ecx, dword ptr [ebp+fffffea0]
'0071d138    894808                  mov dword ptr [eax+08], ecx
'0071d13b    8b95a4feffff            mov edx, dword ptr [ebp+fffffea4]
'0071d141    89500c                  mov dword ptr [eax+0c], edx
'0071d144    8b85dcfdffff            mov eax, dword ptr [ebp+fffffddc]
'0071d14a    8b08                    mov ecx, dword ptr [eax]
'0071d14c    8b95dcfdffff            mov edx, dword ptr [ebp+fffffddc]
'0071d152    52                      push edx
'0071d153    ff5130                  call dword ptr [ecx+30]
'0071d156    dbe2                    fnclex
'0071d158    8985d8fdffff            mov dword ptr [ebp+fffffdd8], eax
'0071d15e    83bdd8fdffff00          cmp dword ptr [ebp+fffffdd8], 00
'0071d165    7d23                    jge 71d18a

If (var_7 < 0) Then
'0071d167    6a30                    push 30
'0071d169    68d8304300              push 004330d8
'0071d16e    8b85dcfdffff            mov eax, dword ptr [ebp+fffffddc]
'0071d174    50                      push eax
'0071d175    8b8dd8fdffff            mov ecx, dword ptr [ebp+fffffdd8]
'0071d17b    51                      push ecx

' *** Reference to "__vbaHresultCheckObj"
'0071d17c    ff1580104000            call dword ptr [00401080]
'0071d182    898578fcffff            mov dword ptr [ebp+fffffc78], eax
'0071d188    eb0a                    jmp 71d194
    
End If
'0071d18a    c78578fcffff00000000    mov dword ptr [ebp+fffffc78], 00000000
'0071d194    8b559c                  mov edx, dword ptr [ebp-64]
'0071d197    8995d4fdffff            mov dword ptr [ebp+fffffdd4], edx
'0071d19d    8d8578ffffff            lea eax, dword ptr [ebp+ffffff78]
'0071d1a3    50                      push eax
'0071d1a4    8b8dd4fdffff            mov ecx, dword ptr [ebp+fffffdd4]
'0071d1aa    8b11                    mov edx, dword ptr [ecx]
'0071d1ac    8b85d4fdffff            mov eax, dword ptr [ebp+fffffdd4]
'0071d1b2    50                      push eax
'0071d1b3    ff5244                  call dword ptr [edx+44]
'0071d1b6    dbe2                    fnclex
'0071d1b8    8985d0fdffff            mov dword ptr [ebp+fffffdd0], eax
'0071d1be    83bdd0fdffff00          cmp dword ptr [ebp+fffffdd0], 00
'0071d1c5    7d23                    jge 71d1ea

If (var_51 < 0) Then
'0071d1c7    6a44                    push 44
'0071d1c9    6880304300              push 00433080
'0071d1ce    8b8dd4fdffff            mov ecx, dword ptr [ebp+fffffdd4]
'0071d1d4    51                      push ecx
'0071d1d5    8b95d0fdffff            mov edx, dword ptr [ebp+fffffdd0]
'0071d1db    52                      push edx

' *** Reference to "__vbaHresultCheckObj"
'0071d1dc    ff1580104000            call dword ptr [00401080]
'0071d1e2    898574fcffff            mov dword ptr [ebp+fffffc74], eax
'0071d1e8    eb0a                    jmp 71d1f4
    
End If
'0071d1ea    c78574fcffff00000000    mov dword ptr [ebp+fffffc74], 00000000
'0071d1f4    8d8578ffffff            lea eax, dword ptr [ebp+ffffff78]
'0071d1fa    50                      push eax

' *** Reference to "__vbaStrVarMove"
'0071d1fb    ff153c104000            call dword ptr [0040103c]
'0071d201    8bd0                    mov edx, eax
'0071d203    8d4dd4                  lea ecx, dword ptr [ebp-2c]

' *** Reference to "__vbaStrMove"
'0071d206    ff15d0124000            call dword ptr [004012d0]
'0071d20c    8d4d9c                  lea ecx, dword ptr [ebp-64]
'0071d20f    51                      push ecx
'0071d210    8d55a0                  lea edx, dword ptr [ebp-60]
'0071d213    52                      push edx
'0071d214    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'0071d216    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_7, var_51)
'0071d21c    83c40c                  add esp, 0c
'0071d21f    8d8d78ffffff            lea ecx, dword ptr [ebp+ffffff78]

' *** Reference to "__vbaFreeVar"
'0071d225    ff1528104000            call dword ptr [00401028]
'#Cleanup(var_87)
'0071d22b    c745fc23000000          mov dword ptr [ebp-04], 00000023
'0071d232    668b45c8                mov ax, word ptr [ebp-38]
'0071d236    668945cc                mov word ptr [ebp-34], ax

'ERROR: Two many next close:
End If
'0071d23a    c745fc25000000          mov dword ptr [ebp-04], 00000025
'0071d241    668b4dc4                mov cx, word ptr [ebp-3c]
'0071d245    663b4dcc                cmp cx, word ptr [ebp-34]
'0071d249    0f8e9e010000            jle 71d3ed

If (CInt(IIf(IsNull(var_51), var_19, var_121)) > WORD PTR [EBP-34]) Then
'0071d24f    c745fc26000000          mov dword ptr [ebp-04], 00000026
'0071d256    8d55a0                  lea edx, dword ptr [ebp-60]
'0071d259    52                      push edx
'0071d25a    8b45ac                  mov eax, dword ptr [ebp-54]
'0071d25d    8b08                    mov ecx, dword ptr [eax]
'0071d25f    8b55ac                  mov edx, dword ptr [ebp-54]
'0071d262    52                      push edx
'0071d263    ff91b4000000            call dword ptr [ecx+000000b4]
'0071d269    dbe2                    fnclex
'0071d26b    8985e0fdffff            mov dword ptr [ebp+fffffde0], eax
'0071d271    83bde0fdffff00          cmp dword ptr [ebp+fffffde0], 00
'0071d278    7d23                    jge 71d29d
    
    If (    var_50 < 0) Then
'0071d27a    68b4000000              push 000000b4
'0071d27f    6830314300              push 00433130
'0071d284    8b45ac                  mov eax, dword ptr [ebp-54]
'0071d287    50                      push eax
'0071d288    8b8de0fdffff            mov ecx, dword ptr [ebp+fffffde0]
'0071d28e    51                      push ecx

' *** Reference to "__vbaHresultCheckObj"
'0071d28f    ff1580104000            call dword ptr [00401080]
'0071d295    898570fcffff            mov dword ptr [ebp+fffffc70], eax
'0071d29b    eb0a                    jmp 71d2a7
    
End If
'0071d29d    c78570fcffff00000000    mov dword ptr [ebp+fffffc70], 00000000
'0071d2a7    8b55a0                  mov edx, dword ptr [ebp-60]
'0071d2aa    8995dcfdffff            mov dword ptr [ebp+fffffddc], edx
'0071d2b0    c785a0feffff3cad4300    mov dword ptr [ebp+fffffea0], 0043ad3c
'0071d2ba    c78598feffff08000000    mov dword ptr [ebp+fffffe98], 00000008
'0071d2c4    8d459c                  lea eax, dword ptr [ebp-64]
'0071d2c7    50                      push eax
'0071d2c8    b810000000              mov eax, 00000010

' *** Reference to "__vbaChkstk"
'0071d2cd    e88e9fceff              call 407260
'0071d2d2    8bcc                    mov ecx, esp
'0071d2d4    8b9598feffff            mov edx, dword ptr [ebp+fffffe98]
'0071d2da    8911                    mov dword ptr [ecx], edx
'0071d2dc    8b859cfeffff            mov eax, dword ptr [ebp+fffffe9c]
'0071d2e2    894104                  mov dword ptr [ecx+04], eax
'0071d2e5    8b95a0feffff            mov edx, dword ptr [ebp+fffffea0]
'0071d2eb    895108                  mov dword ptr [ecx+08], edx
'0071d2ee    8b85a4feffff            mov eax, dword ptr [ebp+fffffea4]
'0071d2f4    89410c                  mov dword ptr [ecx+0c], eax
'0071d2f7    8b8ddcfdffff            mov ecx, dword ptr [ebp+fffffddc]
'0071d2fd    8b11                    mov edx, dword ptr [ecx]
'0071d2ff    8b85dcfdffff            mov eax, dword ptr [ebp+fffffddc]
'0071d305    50                      push eax
'0071d306    ff5230                  call dword ptr [edx+30]
'0071d309    dbe2                    fnclex
'0071d30b    8985d8fdffff            mov dword ptr [ebp+fffffdd8], eax
'0071d311    83bdd8fdffff00          cmp dword ptr [ebp+fffffdd8], 00
'0071d318    7d23                    jge 71d33d

If (var_7 < 0) Then
'0071d31a    6a30                    push 30
'0071d31c    68d8304300              push 004330d8
'0071d321    8b8ddcfdffff            mov ecx, dword ptr [ebp+fffffddc]
'0071d327    51                      push ecx
'0071d328    8b95d8fdffff            mov edx, dword ptr [ebp+fffffdd8]
'0071d32e    52                      push edx

' *** Reference to "__vbaHresultCheckObj"
'0071d32f    ff1580104000            call dword ptr [00401080]
'0071d335    89856cfcffff            mov dword ptr [ebp+fffffc6c], eax
'0071d33b    eb0a                    jmp 71d347
    
End If
'0071d33d    c7856cfcffff00000000    mov dword ptr [ebp+fffffc6c], 00000000
'0071d347    8b459c                  mov eax, dword ptr [ebp-64]
'0071d34a    8985d4fdffff            mov dword ptr [ebp+fffffdd4], eax
'0071d350    8d8d78ffffff            lea ecx, dword ptr [ebp+ffffff78]
'0071d356    51                      push ecx
'0071d357    8b95d4fdffff            mov edx, dword ptr [ebp+fffffdd4]
'0071d35d    8b02                    mov eax, dword ptr [edx]
'0071d35f    8b8dd4fdffff            mov ecx, dword ptr [ebp+fffffdd4]
'0071d365    51                      push ecx
'0071d366    ff5044                  call dword ptr [eax+44]
'0071d369    dbe2                    fnclex
'0071d36b    8985d0fdffff            mov dword ptr [ebp+fffffdd0], eax
'0071d371    83bdd0fdffff00          cmp dword ptr [ebp+fffffdd0], 00
'0071d378    7d23                    jge 71d39d

If (-256 - 24 < 0) Then
'0071d37a    6a44                    push 44
'0071d37c    6880304300              push 00433080
'0071d381    8b95d4fdffff            mov edx, dword ptr [ebp+fffffdd4]
'0071d387    52                      push edx
'0071d388    8b85d0fdffff            mov eax, dword ptr [ebp+fffffdd0]
'0071d38e    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0071d38f    ff1580104000            call dword ptr [00401080]
'0071d395    898568fcffff            mov dword ptr [ebp+fffffc68], eax
'0071d39b    eb0a                    jmp 71d3a7
    
End If
'0071d39d    c78568fcffff00000000    mov dword ptr [ebp+fffffc68], 00000000
'0071d3a7    8d8d78ffffff            lea ecx, dword ptr [ebp+ffffff78]
'0071d3ad    51                      push ecx

' *** Reference to "__vbaStrVarMove"
'0071d3ae    ff153c104000            call dword ptr [0040103c]
'0071d3b4    8bd0                    mov edx, eax
'0071d3b6    8d4dd4                  lea ecx, dword ptr [ebp-2c]

' *** Reference to "__vbaStrMove"
'0071d3b9    ff15d0124000            call dword ptr [004012d0]
'0071d3bf    8d559c                  lea edx, dword ptr [ebp-64]
'0071d3c2    52                      push edx
'0071d3c3    8d45a0                  lea eax, dword ptr [ebp-60]
'0071d3c6    50                      push eax
'0071d3c7    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'0071d3c9    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_7, var_51)
'0071d3cf    83c40c                  add esp, 0c
'0071d3d2    8d8d78ffffff            lea ecx, dword ptr [ebp+ffffff78]

' *** Reference to "__vbaFreeVar"
'0071d3d8    ff1528104000            call dword ptr [00401028]
'#Cleanup(var_87)
'0071d3de    c745fc27000000          mov dword ptr [ebp-04], 00000027
'0071d3e5    668b4dc4                mov cx, word ptr [ebp-3c]
'0071d3e9    66894dcc                mov word ptr [ebp-34], cx

'ERROR: Two many next close:
End If
'0071d3ed    c745fc29000000          mov dword ptr [ebp-04], 00000029
'0071d3f4    668b55cc                mov dx, word ptr [ebp-34]
'0071d3f8    668995a0feffff          mov word ptr [ebp+fffffea0], dx
'0071d3ff    c78598feffff02800000    mov dword ptr [ebp+fffffe98], 00008002
'0071d409    8d45b4                  lea eax, dword ptr [ebp-4c]
'0071d40c    50                      push eax
'0071d40d    8d8d98feffff            lea ecx, dword ptr [ebp+fffffe98]
'0071d413    51                      push ecx

' *** Reference to "__vbaVarTstGt"
'0071d414    ff1504104000            call dword ptr [00401004]
'0071d41a    0fbfd0                  movsx edx, ax
'0071d41d    85d2                    test dx, dx
'0071d41f    0f84a4010000            je 71d5c9

If (((var_62) > (CInt(IIf(IsNull(var_51), var_19, var_121))))) Then
'0071d425    c745fc2a000000          mov dword ptr [ebp-04], 0000002a
'0071d42c    8d45a0                  lea eax, dword ptr [ebp-60]
'0071d42f    50                      push eax
'0071d430    8b4dac                  mov ecx, dword ptr [ebp-54]
'0071d433    8b11                    mov edx, dword ptr [ecx]
'0071d435    8b45ac                  mov eax, dword ptr [ebp-54]
'0071d438    50                      push eax
'0071d439    ff92b4000000            call dword ptr [edx+000000b4]
'0071d43f    dbe2                    fnclex
'0071d441    8985e0fdffff            mov dword ptr [ebp+fffffde0], eax
'0071d447    83bde0fdffff00          cmp dword ptr [ebp+fffffde0], 00
'0071d44e    7d23                    jge 71d473
    
    If (    var_50 < 0) Then
'0071d450    68b4000000              push 000000b4
'0071d455    6830314300              push 00433130
'0071d45a    8b4dac                  mov ecx, dword ptr [ebp-54]
'0071d45d    51                      push ecx
'0071d45e    8b95e0fdffff            mov edx, dword ptr [ebp+fffffde0]
'0071d464    52                      push edx

' *** Reference to "__vbaHresultCheckObj"
'0071d465    ff1580104000            call dword ptr [00401080]
'0071d46b    898564fcffff            mov dword ptr [ebp+fffffc64], eax
'0071d471    eb0a                    jmp 71d47d
    
End If
'0071d473    c78564fcffff00000000    mov dword ptr [ebp+fffffc64], 00000000
'0071d47d    8b45a0                  mov eax, dword ptr [ebp-60]
'0071d480    8985dcfdffff            mov dword ptr [ebp+fffffddc], eax
'0071d486    c785a0feffffb4b54300    mov dword ptr [ebp+fffffea0], 0043b5b4
'0071d490    c78598feffff08000000    mov dword ptr [ebp+fffffe98], 00000008
'0071d49a    8d4d9c                  lea ecx, dword ptr [ebp-64]
'0071d49d    51                      push ecx
'0071d49e    b810000000              mov eax, 00000010

' *** Reference to "__vbaChkstk"
'0071d4a3    e8b89dceff              call 407260
'0071d4a8    8bd4                    mov edx, esp
'0071d4aa    8b8598feffff            mov eax, dword ptr [ebp+fffffe98]
'0071d4b0    8902                    mov dword ptr [edx], eax
'0071d4b2    8b8d9cfeffff            mov ecx, dword ptr [ebp+fffffe9c]
'0071d4b8    894a04                  mov dword ptr [edx+04], ecx
'0071d4bb    8b85a0feffff            mov eax, dword ptr [ebp+fffffea0]
'0071d4c1    894208                  mov dword ptr [edx+08], eax
'0071d4c4    8b8da4feffff            mov ecx, dword ptr [ebp+fffffea4]
'0071d4ca    894a0c                  mov dword ptr [edx+0c], ecx
'0071d4cd    8b95dcfdffff            mov edx, dword ptr [ebp+fffffddc]
'0071d4d3    8b02                    mov eax, dword ptr [edx]
'0071d4d5    8b8ddcfdffff            mov ecx, dword ptr [ebp+fffffddc]
'0071d4db    51                      push ecx
'0071d4dc    ff5030                  call dword ptr [eax+30]
'0071d4df    dbe2                    fnclex
'0071d4e1    8985d8fdffff            mov dword ptr [ebp+fffffdd8], eax
'0071d4e7    83bdd8fdffff00          cmp dword ptr [ebp+fffffdd8], 00
'0071d4ee    7d23                    jge 71d513

If (-256 - 24 < 0) Then
'0071d4f0    6a30                    push 30
'0071d4f2    68d8304300              push 004330d8
'0071d4f7    8b95dcfdffff            mov edx, dword ptr [ebp+fffffddc]
'0071d4fd    52                      push edx
'0071d4fe    8b85d8fdffff            mov eax, dword ptr [ebp+fffffdd8]
'0071d504    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0071d505    ff1580104000            call dword ptr [00401080]
'0071d50b    898560fcffff            mov dword ptr [ebp+fffffc60], eax
'0071d511    eb0a                    jmp 71d51d
    
End If
'0071d513    c78560fcffff00000000    mov dword ptr [ebp+fffffc60], 00000000
'0071d51d    8b4d9c                  mov ecx, dword ptr [ebp-64]
'0071d520    898dd4fdffff            mov dword ptr [ebp+fffffdd4], ecx
'0071d526    8d9578ffffff            lea edx, dword ptr [ebp+ffffff78]
'0071d52c    52                      push edx
'0071d52d    8b85d4fdffff            mov eax, dword ptr [ebp+fffffdd4]
'0071d533    8b08                    mov ecx, dword ptr [eax]
'0071d535    8b95d4fdffff            mov edx, dword ptr [ebp+fffffdd4]
'0071d53b    52                      push edx
'0071d53c    ff5144                  call dword ptr [ecx+44]
'0071d53f    dbe2                    fnclex
'0071d541    8985d0fdffff            mov dword ptr [ebp+fffffdd0], eax
'0071d547    83bdd0fdffff00          cmp dword ptr [ebp+fffffdd0], 00
'0071d54e    7d23                    jge 71d573

If (var_51 < 0) Then
'0071d550    6a44                    push 44
'0071d552    6880304300              push 00433080
'0071d557    8b85d4fdffff            mov eax, dword ptr [ebp+fffffdd4]
'0071d55d    50                      push eax
'0071d55e    8b8dd0fdffff            mov ecx, dword ptr [ebp+fffffdd0]
'0071d564    51                      push ecx

' *** Reference to "__vbaHresultCheckObj"
'0071d565    ff1580104000            call dword ptr [00401080]
'0071d56b    89855cfcffff            mov dword ptr [ebp+fffffc5c], eax
'0071d571    eb0a                    jmp 71d57d
    
End If
'0071d573    c7855cfcffff00000000    mov dword ptr [ebp+fffffc5c], 00000000
'0071d57d    8d9578ffffff            lea edx, dword ptr [ebp+ffffff78]
'0071d583    52                      push edx

' *** Reference to "__vbaStrVarMove"
'0071d584    ff153c104000            call dword ptr [0040103c]
'0071d58a    8bd0                    mov edx, eax
'0071d58c    8d4dd4                  lea ecx, dword ptr [ebp-2c]

' *** Reference to "__vbaStrMove"
'0071d58f    ff15d0124000            call dword ptr [004012d0]
'0071d595    8d459c                  lea eax, dword ptr [ebp-64]
'0071d598    50                      push eax
'0071d599    8d4da0                  lea ecx, dword ptr [ebp-60]
'0071d59c    51                      push ecx
'0071d59d    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'0071d59f    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_7, var_51)
'0071d5a5    83c40c                  add esp, 0c
'0071d5a8    8d8d78ffffff            lea ecx, dword ptr [ebp+ffffff78]

' *** Reference to "__vbaFreeVar"
'0071d5ae    ff1528104000            call dword ptr [00401028]
'#Cleanup(var_87)
'0071d5b4    c745fc2b000000          mov dword ptr [ebp-04], 0000002b
'0071d5bb    8d55b4                  lea edx, dword ptr [ebp-4c]
'0071d5be    52                      push edx

' *** Reference to "__vbaI2Var"
'0071d5bf    ff150c124000            call dword ptr [0040120c]
'0071d5c5    668945cc                mov word ptr [ebp-34], ax

'ERROR: Two many next close:
End If
'0071d5c9    c745fc2d000000          mov dword ptr [ebp-04], 0000002d
'0071d5d0    668b45b0                mov ax, word ptr [ebp-50]
'0071d5d4    663b45cc                cmp ax, word ptr [ebp-34]
'0071d5d8    0f8e9e010000            jle 71d77c

If (CInt(IIf(IsNull(var_51), var_19, var_121)) > WORD PTR [EBP-34]) Then
'0071d5de    c745fc2e000000          mov dword ptr [ebp-04], 0000002e
'0071d5e5    8d4da0                  lea ecx, dword ptr [ebp-60]
'0071d5e8    51                      push ecx
'0071d5e9    8b55ac                  mov edx, dword ptr [ebp-54]
'0071d5ec    8b02                    mov eax, dword ptr [edx]
'0071d5ee    8b4dac                  mov ecx, dword ptr [ebp-54]
'0071d5f1    51                      push ecx
'0071d5f2    ff90b4000000            call dword ptr [eax+000000b4]
'0071d5f8    dbe2                    fnclex
'0071d5fa    8985e0fdffff            mov dword ptr [ebp+fffffde0], eax
'0071d600    83bde0fdffff00          cmp dword ptr [ebp+fffffde0], 00
'0071d607    7d23                    jge 71d62c
    
    If (    frmImport < 0) Then
'0071d609    68b4000000              push 000000b4
'0071d60e    6830314300              push 00433130
'0071d613    8b55ac                  mov edx, dword ptr [ebp-54]
'0071d616    52                      push edx
'0071d617    8b85e0fdffff            mov eax, dword ptr [ebp+fffffde0]
'0071d61d    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0071d61e    ff1580104000            call dword ptr [00401080]
'0071d624    898558fcffff            mov dword ptr [ebp+fffffc58], eax
'0071d62a    eb0a                    jmp 71d636
    
End If
'0071d62c    c78558fcffff00000000    mov dword ptr [ebp+fffffc58], 00000000
'0071d636    8b4da0                  mov ecx, dword ptr [ebp-60]
'0071d639    898ddcfdffff            mov dword ptr [ebp+fffffddc], ecx
'0071d63f    c785a0feffffccb54300    mov dword ptr [ebp+fffffea0], 0043b5cc
'0071d649    c78598feffff08000000    mov dword ptr [ebp+fffffe98], 00000008
'0071d653    8d559c                  lea edx, dword ptr [ebp-64]
'0071d656    52                      push edx
'0071d657    b810000000              mov eax, 00000010

' *** Reference to "__vbaChkstk"
'0071d65c    e8ff9bceff              call 407260
'0071d661    8bc4                    mov eax, esp
'0071d663    8b8d98feffff            mov ecx, dword ptr [ebp+fffffe98]
'0071d669    8908                    mov dword ptr [eax], ecx
'0071d66b    8b959cfeffff            mov edx, dword ptr [ebp+fffffe9c]
'0071d671    895004                  mov dword ptr [eax+04], edx
'0071d674    8b8da0feffff            mov ecx, dword ptr [ebp+fffffea0]
'0071d67a    894808                  mov dword ptr [eax+08], ecx
'0071d67d    8b95a4feffff            mov edx, dword ptr [ebp+fffffea4]
'0071d683    89500c                  mov dword ptr [eax+0c], edx
'0071d686    8b85dcfdffff            mov eax, dword ptr [ebp+fffffddc]
'0071d68c    8b08                    mov ecx, dword ptr [eax]
'0071d68e    8b95dcfdffff            mov edx, dword ptr [ebp+fffffddc]
'0071d694    52                      push edx
'0071d695    ff5130                  call dword ptr [ecx+30]
'0071d698    dbe2                    fnclex
'0071d69a    8985d8fdffff            mov dword ptr [ebp+fffffdd8], eax
'0071d6a0    83bdd8fdffff00          cmp dword ptr [ebp+fffffdd8], 00
'0071d6a7    7d23                    jge 71d6cc

If (var_7 < 0) Then
'0071d6a9    6a30                    push 30
'0071d6ab    68d8304300              push 004330d8
'0071d6b0    8b85dcfdffff            mov eax, dword ptr [ebp+fffffddc]
'0071d6b6    50                      push eax
'0071d6b7    8b8dd8fdffff            mov ecx, dword ptr [ebp+fffffdd8]
'0071d6bd    51                      push ecx

' *** Reference to "__vbaHresultCheckObj"
'0071d6be    ff1580104000            call dword ptr [00401080]
'0071d6c4    898554fcffff            mov dword ptr [ebp+fffffc54], eax
'0071d6ca    eb0a                    jmp 71d6d6
    
End If
'0071d6cc    c78554fcffff00000000    mov dword ptr [ebp+fffffc54], 00000000
'0071d6d6    8b559c                  mov edx, dword ptr [ebp-64]
'0071d6d9    8995d4fdffff            mov dword ptr [ebp+fffffdd4], edx
'0071d6df    8d8578ffffff            lea eax, dword ptr [ebp+ffffff78]
'0071d6e5    50                      push eax
'0071d6e6    8b8dd4fdffff            mov ecx, dword ptr [ebp+fffffdd4]
'0071d6ec    8b11                    mov edx, dword ptr [ecx]
'0071d6ee    8b85d4fdffff            mov eax, dword ptr [ebp+fffffdd4]
'0071d6f4    50                      push eax
'0071d6f5    ff5244                  call dword ptr [edx+44]
'0071d6f8    dbe2                    fnclex
'0071d6fa    8985d0fdffff            mov dword ptr [ebp+fffffdd0], eax
'0071d700    83bdd0fdffff00          cmp dword ptr [ebp+fffffdd0], 00
'0071d707    7d23                    jge 71d72c

If (var_51 < 0) Then
'0071d709    6a44                    push 44
'0071d70b    6880304300              push 00433080
'0071d710    8b8dd4fdffff            mov ecx, dword ptr [ebp+fffffdd4]
'0071d716    51                      push ecx
'0071d717    8b95d0fdffff            mov edx, dword ptr [ebp+fffffdd0]
'0071d71d    52                      push edx

' *** Reference to "__vbaHresultCheckObj"
'0071d71e    ff1580104000            call dword ptr [00401080]
'0071d724    898550fcffff            mov dword ptr [ebp+fffffc50], eax
'0071d72a    eb0a                    jmp 71d736
    
End If
'0071d72c    c78550fcffff00000000    mov dword ptr [ebp+fffffc50], 00000000
'0071d736    8d8578ffffff            lea eax, dword ptr [ebp+ffffff78]
'0071d73c    50                      push eax

' *** Reference to "__vbaStrVarMove"
'0071d73d    ff153c104000            call dword ptr [0040103c]
'0071d743    8bd0                    mov edx, eax
'0071d745    8d4dd4                  lea ecx, dword ptr [ebp-2c]

' *** Reference to "__vbaStrMove"
'0071d748    ff15d0124000            call dword ptr [004012d0]
'0071d74e    8d4d9c                  lea ecx, dword ptr [ebp-64]
'0071d751    51                      push ecx
'0071d752    8d55a0                  lea edx, dword ptr [ebp-60]
'0071d755    52                      push edx
'0071d756    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'0071d758    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_7, var_51)
'0071d75e    83c40c                  add esp, 0c
'0071d761    8d8d78ffffff            lea ecx, dword ptr [ebp+ffffff78]

' *** Reference to "__vbaFreeVar"
'0071d767    ff1528104000            call dword ptr [00401028]
'#Cleanup(var_87)
'0071d76d    c745fc2f000000          mov dword ptr [ebp-04], 0000002f
'0071d774    668b45b0                mov ax, word ptr [ebp-50]
'0071d778    668945cc                mov word ptr [ebp-34], ax

'ERROR: Two many next close:
End If
'0071d77c    c745fc31000000          mov dword ptr [ebp-04], 00000031
'0071d783    8d4da0                  lea ecx, dword ptr [ebp-60]
'0071d786    51                      push ecx
'0071d787    8b55ac                  mov edx, dword ptr [ebp-54]
'0071d78a    8b02                    mov eax, dword ptr [edx]
'0071d78c    8b4dac                  mov ecx, dword ptr [ebp-54]
'0071d78f    51                      push ecx
'0071d790    ff90b4000000            call dword ptr [eax+000000b4]
'0071d796    dbe2                    fnclex
'0071d798    8985e0fdffff            mov dword ptr [ebp+fffffde0], eax
'0071d79e    83bde0fdffff00          cmp dword ptr [ebp+fffffde0], 00
'0071d7a5    7d23                    jge 71d7ca

If (frmImport < 0) Then
'0071d7a7    68b4000000              push 000000b4
'0071d7ac    6830314300              push 00433130
'0071d7b1    8b55ac                  mov edx, dword ptr [ebp-54]
'0071d7b4    52                      push edx
'0071d7b5    8b85e0fdffff            mov eax, dword ptr [ebp+fffffde0]
'0071d7bb    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0071d7bc    ff1580104000            call dword ptr [00401080]
'0071d7c2    89854cfcffff            mov dword ptr [ebp+fffffc4c], eax
'0071d7c8    eb0a                    jmp 71d7d4
    
End If
'0071d7ca    c7854cfcffff00000000    mov dword ptr [ebp+fffffc4c], 00000000
'0071d7d4    8b4da0                  mov ecx, dword ptr [ebp-60]
'0071d7d7    898ddcfdffff            mov dword ptr [ebp+fffffddc], ecx
'0071d7dd    c785a0feffff64b14300    mov dword ptr [ebp+fffffea0], 0043b164
'0071d7e7    c78598feffff08000000    mov dword ptr [ebp+fffffe98], 00000008
'0071d7f1    8d559c                  lea edx, dword ptr [ebp-64]
'0071d7f4    52                      push edx
'0071d7f5    b810000000              mov eax, 00000010

' *** Reference to "__vbaChkstk"
'0071d7fa    e8619aceff              call 407260
'0071d7ff    8bc4                    mov eax, esp
'0071d801    8b8d98feffff            mov ecx, dword ptr [ebp+fffffe98]
'0071d807    8908                    mov dword ptr [eax], ecx
'0071d809    8b959cfeffff            mov edx, dword ptr [ebp+fffffe9c]
'0071d80f    895004                  mov dword ptr [eax+04], edx
'0071d812    8b8da0feffff            mov ecx, dword ptr [ebp+fffffea0]
'0071d818    894808                  mov dword ptr [eax+08], ecx
'0071d81b    8b95a4feffff            mov edx, dword ptr [ebp+fffffea4]
'0071d821    89500c                  mov dword ptr [eax+0c], edx
'0071d824    8b85dcfdffff            mov eax, dword ptr [ebp+fffffddc]
'0071d82a    8b08                    mov ecx, dword ptr [eax]
'0071d82c    8b95dcfdffff            mov edx, dword ptr [ebp+fffffddc]
'0071d832    52                      push edx
'0071d833    ff5130                  call dword ptr [ecx+30]
'0071d836    dbe2                    fnclex
'0071d838    8985d8fdffff            mov dword ptr [ebp+fffffdd8], eax
'0071d83e    83bdd8fdffff00          cmp dword ptr [ebp+fffffdd8], 00
'0071d845    7d23                    jge 71d86a

If (var_7 < 0) Then
'0071d847    6a30                    push 30
'0071d849    68d8304300              push 004330d8
'0071d84e    8b85dcfdffff            mov eax, dword ptr [ebp+fffffddc]
'0071d854    50                      push eax
'0071d855    8b8dd8fdffff            mov ecx, dword ptr [ebp+fffffdd8]
'0071d85b    51                      push ecx

' *** Reference to "__vbaHresultCheckObj"
'0071d85c    ff1580104000            call dword ptr [00401080]
'0071d862    898548fcffff            mov dword ptr [ebp+fffffc48], eax
'0071d868    eb0a                    jmp 71d874
    
End If
'0071d86a    c78548fcffff00000000    mov dword ptr [ebp+fffffc48], 00000000
'0071d874    8b559c                  mov edx, dword ptr [ebp-64]
'0071d877    8995d4fdffff            mov dword ptr [ebp+fffffdd4], edx
'0071d87d    8d8578ffffff            lea eax, dword ptr [ebp+ffffff78]
'0071d883    50                      push eax
'0071d884    8b8dd4fdffff            mov ecx, dword ptr [ebp+fffffdd4]
'0071d88a    8b11                    mov edx, dword ptr [ecx]
'0071d88c    8b85d4fdffff            mov eax, dword ptr [ebp+fffffdd4]
'0071d892    50                      push eax
'0071d893    ff5244                  call dword ptr [edx+44]
'0071d896    dbe2                    fnclex
'0071d898    8985d0fdffff            mov dword ptr [ebp+fffffdd0], eax
'0071d89e    83bdd0fdffff00          cmp dword ptr [ebp+fffffdd0], 00
'0071d8a5    7d23                    jge 71d8ca

If (var_51 < 0) Then
'0071d8a7    6a44                    push 44
'0071d8a9    6880304300              push 00433080
'0071d8ae    8b8dd4fdffff            mov ecx, dword ptr [ebp+fffffdd4]
'0071d8b4    51                      push ecx
'0071d8b5    8b95d0fdffff            mov edx, dword ptr [ebp+fffffdd0]
'0071d8bb    52                      push edx

' *** Reference to "__vbaHresultCheckObj"
'0071d8bc    ff1580104000            call dword ptr [00401080]
'0071d8c2    898544fcffff            mov dword ptr [ebp+fffffc44], eax
'0071d8c8    eb0a                    jmp 71d8d4
    
End If
'0071d8ca    c78544fcffff00000000    mov dword ptr [ebp+fffffc44], 00000000
'0071d8d4    c78590feffff44ed4300    mov dword ptr [ebp+fffffe90], 0043ed44
'0071d8de    c78588feffff08000000    mov dword ptr [ebp+fffffe88], 00000008
'0071d8e8    8d4598                  lea eax, dword ptr [ebp-68]
'0071d8eb    50                      push eax
'0071d8ec    8b4dac                  mov ecx, dword ptr [ebp-54]
'0071d8ef    8b11                    mov edx, dword ptr [ecx]
'0071d8f1    8b45ac                  mov eax, dword ptr [ebp-54]
'0071d8f4    50                      push eax
'0071d8f5    ff92b4000000            call dword ptr [edx+000000b4]
'0071d8fb    dbe2                    fnclex
'0071d8fd    8985ccfdffff            mov dword ptr [ebp+fffffdcc], eax
'0071d903    83bdccfdffff00          cmp dword ptr [ebp+fffffdcc], 00
'0071d90a    7d23                    jge 71d92f

If (var_50 < 0) Then
'0071d90c    68b4000000              push 000000b4
'0071d911    6830314300              push 00433130
'0071d916    8b4dac                  mov ecx, dword ptr [ebp-54]
'0071d919    51                      push ecx
'0071d91a    8b95ccfdffff            mov edx, dword ptr [ebp+fffffdcc]
'0071d920    52                      push edx

' *** Reference to "__vbaHresultCheckObj"
'0071d921    ff1580104000            call dword ptr [00401080]
'0071d927    898540fcffff            mov dword ptr [ebp+fffffc40], eax
'0071d92d    eb0a                    jmp 71d939
    
End If
'0071d92f    c78540fcffff00000000    mov dword ptr [ebp+fffffc40], 00000000
'0071d939    8b4598                  mov eax, dword ptr [ebp-68]
'0071d93c    8985c8fdffff            mov dword ptr [ebp+fffffdc8], eax
'0071d942    c78580feffff64b14300    mov dword ptr [ebp+fffffe80], 0043b164
'0071d94c    c78578feffff08000000    mov dword ptr [ebp+fffffe78], 00000008
'0071d956    8d4d94                  lea ecx, dword ptr [ebp-6c]
'0071d959    51                      push ecx
'0071d95a    b810000000              mov eax, 00000010

' *** Reference to "__vbaChkstk"
'0071d95f    e8fc98ceff              call 407260
'0071d964    8bd4                    mov edx, esp
'0071d966    8b8578feffff            mov eax, dword ptr [ebp+fffffe78]
'0071d96c    8902                    mov dword ptr [edx], eax
'0071d96e    8b8d7cfeffff            mov ecx, dword ptr [ebp+fffffe7c]
'0071d974    894a04                  mov dword ptr [edx+04], ecx
'0071d977    8b8580feffff            mov eax, dword ptr [ebp+fffffe80]
'0071d97d    894208                  mov dword ptr [edx+08], eax
'0071d980    8b8d84feffff            mov ecx, dword ptr [ebp+fffffe84]
'0071d986    894a0c                  mov dword ptr [edx+0c], ecx
'0071d989    8b95c8fdffff            mov edx, dword ptr [ebp+fffffdc8]
'0071d98f    8b02                    mov eax, dword ptr [edx]
'0071d991    8b8dc8fdffff            mov ecx, dword ptr [ebp+fffffdc8]
'0071d997    51                      push ecx
'0071d998    ff5030                  call dword ptr [eax+30]
'0071d99b    dbe2                    fnclex
'0071d99d    8985c4fdffff            mov dword ptr [ebp+fffffdc4], eax
'0071d9a3    83bdc4fdffff00          cmp dword ptr [ebp+fffffdc4], 00
'0071d9aa    7d23                    jge 71d9cf

If (-256 - 24 < 0) Then
'0071d9ac    6a30                    push 30
'0071d9ae    68d8304300              push 004330d8
'0071d9b3    8b95c8fdffff            mov edx, dword ptr [ebp+fffffdc8]
'0071d9b9    52                      push edx
'0071d9ba    8b85c4fdffff            mov eax, dword ptr [ebp+fffffdc4]
'0071d9c0    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0071d9c1    ff1580104000            call dword ptr [00401080]
'0071d9c7    89853cfcffff            mov dword ptr [ebp+fffffc3c], eax
'0071d9cd    eb0a                    jmp 71d9d9
    
End If
'0071d9cf    c7853cfcffff00000000    mov dword ptr [ebp+fffffc3c], 00000000
'0071d9d9    8b4d94                  mov ecx, dword ptr [ebp-6c]
'0071d9dc    898dc0fdffff            mov dword ptr [ebp+fffffdc0], ecx
'0071d9e2    8d9558ffffff            lea edx, dword ptr [ebp+ffffff58]
'0071d9e8    52                      push edx
'0071d9e9    8b85c0fdffff            mov eax, dword ptr [ebp+fffffdc0]
'0071d9ef    8b08                    mov ecx, dword ptr [eax]
'0071d9f1    8b95c0fdffff            mov edx, dword ptr [ebp+fffffdc0]
'0071d9f7    52                      push edx
'0071d9f8    ff5144                  call dword ptr [ecx+44]
'0071d9fb    dbe2                    fnclex
'0071d9fd    8985bcfdffff            mov dword ptr [ebp+fffffdbc], eax
'0071da03    83bdbcfdffff00          cmp dword ptr [ebp+fffffdbc], 00
'0071da0a    7d23                    jge 71da2f

If (var_121 < 0) Then
'0071da0c    6a44                    push 44
'0071da0e    6880304300              push 00433080
'0071da13    8b85c0fdffff            mov eax, dword ptr [ebp+fffffdc0]
'0071da19    50                      push eax
'0071da1a    8b8dbcfdffff            mov ecx, dword ptr [ebp+fffffdbc]
'0071da20    51                      push ecx

' *** Reference to "__vbaHresultCheckObj"
'0071da21    ff1580104000            call dword ptr [00401080]
'0071da27    898538fcffff            mov dword ptr [ebp+fffffc38], eax
'0071da2d    eb0a                    jmp 71da39
    
End If
'0071da2f    c78538fcffff00000000    mov dword ptr [ebp+fffffc38], 00000000
'0071da39    c78570feffff44ed4300    mov dword ptr [ebp+fffffe70], 0043ed44
'0071da43    c78568feffff08000000    mov dword ptr [ebp+fffffe68], 00000008
'0071da4d    8d5590                  lea edx, dword ptr [ebp-70]
'0071da50    52                      push edx
'0071da51    8b45ac                  mov eax, dword ptr [ebp-54]
'0071da54    8b08                    mov ecx, dword ptr [eax]
'0071da56    8b55ac                  mov edx, dword ptr [ebp-54]
'0071da59    52                      push edx
'0071da5a    ff91b4000000            call dword ptr [ecx+000000b4]
'0071da60    dbe2                    fnclex
'0071da62    8985b8fdffff            mov dword ptr [ebp+fffffdb8], eax
'0071da68    83bdb8fdffff00          cmp dword ptr [ebp+fffffdb8], 00
'0071da6f    7d23                    jge 71da94

If (var_50 < 0) Then
'0071da71    68b4000000              push 000000b4
'0071da76    6830314300              push 00433130
'0071da7b    8b45ac                  mov eax, dword ptr [ebp-54]
'0071da7e    50                      push eax
'0071da7f    8b8db8fdffff            mov ecx, dword ptr [ebp+fffffdb8]
'0071da85    51                      push ecx

' *** Reference to "__vbaHresultCheckObj"
'0071da86    ff1580104000            call dword ptr [00401080]
'0071da8c    898534fcffff            mov dword ptr [ebp+fffffc34], eax
'0071da92    eb0a                    jmp 71da9e
    
End If
'0071da94    c78534fcffff00000000    mov dword ptr [ebp+fffffc34], 00000000
'0071da9e    8b5590                  mov edx, dword ptr [ebp-70]
'0071daa1    8995b4fdffff            mov dword ptr [ebp+fffffdb4], edx
'0071daa7    c78560feffffe4b14300    mov dword ptr [ebp+fffffe60], 0043b1e4
'0071dab1    c78558feffff08000000    mov dword ptr [ebp+fffffe58], 00000008
'0071dabb    8d458c                  lea eax, dword ptr [ebp-74]
'0071dabe    50                      push eax
'0071dabf    b810000000              mov eax, 00000010

' *** Reference to "__vbaChkstk"
'0071dac4    e89797ceff              call 407260
'0071dac9    8bcc                    mov ecx, esp
'0071dacb    8b9558feffff            mov edx, dword ptr [ebp+fffffe58]
'0071dad1    8911                    mov dword ptr [ecx], edx
'0071dad3    8b855cfeffff            mov eax, dword ptr [ebp+fffffe5c]
'0071dad9    894104                  mov dword ptr [ecx+04], eax
'0071dadc    8b9560feffff            mov edx, dword ptr [ebp+fffffe60]
'0071dae2    895108                  mov dword ptr [ecx+08], edx
'0071dae5    8b8564feffff            mov eax, dword ptr [ebp+fffffe64]
'0071daeb    89410c                  mov dword ptr [ecx+0c], eax
'0071daee    8b8db4fdffff            mov ecx, dword ptr [ebp+fffffdb4]
'0071daf4    8b11                    mov edx, dword ptr [ecx]
'0071daf6    8b85b4fdffff            mov eax, dword ptr [ebp+fffffdb4]
'0071dafc    50                      push eax
'0071dafd    ff5230                  call dword ptr [edx+30]
'0071db00    dbe2                    fnclex
'0071db02    8985b0fdffff            mov dword ptr [ebp+fffffdb0], eax
'0071db08    83bdb0fdffff00          cmp dword ptr [ebp+fffffdb0], 00
'0071db0f    7d23                    jge 71db34

If (0 < 0) Then
'0071db11    6a30                    push 30
'0071db13    68d8304300              push 004330d8
'0071db18    8b8db4fdffff            mov ecx, dword ptr [ebp+fffffdb4]
'0071db1e    51                      push ecx
'0071db1f    8b95b0fdffff            mov edx, dword ptr [ebp+fffffdb0]
'0071db25    52                      push edx

' *** Reference to "__vbaHresultCheckObj"
'0071db26    ff1580104000            call dword ptr [00401080]
'0071db2c    898530fcffff            mov dword ptr [ebp+fffffc30], eax
'0071db32    eb0a                    jmp 71db3e
    
End If
'0071db34    c78530fcffff00000000    mov dword ptr [ebp+fffffc30], 00000000
'0071db3e    8b458c                  mov eax, dword ptr [ebp-74]
'0071db41    8985acfdffff            mov dword ptr [ebp+fffffdac], eax
'0071db47    8d8d28ffffff            lea ecx, dword ptr [ebp+ffffff28]
'0071db4d    51                      push ecx
'0071db4e    8b95acfdffff            mov edx, dword ptr [ebp+fffffdac]
'0071db54    8b02                    mov eax, dword ptr [edx]
'0071db56    8b8dacfdffff            mov ecx, dword ptr [ebp+fffffdac]
'0071db5c    51                      push ecx
'0071db5d    ff5044                  call dword ptr [eax+44]
'0071db60    dbe2                    fnclex
'0071db62    8985a8fdffff            mov dword ptr [ebp+fffffda8], eax
'0071db68    83bda8fdffff00          cmp dword ptr [ebp+fffffda8], 00
'0071db6f    7d23                    jge 71db94

If (-256 - 24 < 0) Then
'0071db71    6a44                    push 44
'0071db73    6880304300              push 00433080
'0071db78    8b95acfdffff            mov edx, dword ptr [ebp+fffffdac]
'0071db7e    52                      push edx
'0071db7f    8b85a8fdffff            mov eax, dword ptr [ebp+fffffda8]
'0071db85    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0071db86    ff1580104000            call dword ptr [00401080]
'0071db8c    89852cfcffff            mov dword ptr [ebp+fffffc2c], eax
'0071db92    eb0a                    jmp 71db9e
    
End If
'0071db94    c7852cfcffff00000000    mov dword ptr [ebp+fffffc2c], 00000000
'0071db9e    c78550feffff44ed4300    mov dword ptr [ebp+fffffe50], 0043ed44
'0071dba8    c78548feffff08000000    mov dword ptr [ebp+fffffe48], 00000008
'0071dbb2    668b4da8                mov cx, word ptr [ebp-58]
'0071dbb6    66034da4                add cx, word ptr [ebp-5c]
var_num3 = CInt(IIf(IsNull(0), var_19, 0)) + CInt(IIf(IsNull(var_51), var_19, var_121))
'0071dbba    0f803f050000            jo 71e0ff
'0071dbc0    66034dd8                add cx, word ptr [ebp-28]
var_num3 = var_num3 + CInt(IIf(IsNull(var_51), var_19, var_121))
'0071dbc4    0f8035050000            jo 71e0ff
'0071dbca    66034dd0                add cx, word ptr [ebp-30]
var_num3 = var_num3 + CInt(IIf(IsNull(var_51), var_19, var_121))
'0071dbce    0f802b050000            jo 71e0ff
'0071dbd4    66034dc8                add cx, word ptr [ebp-38]
var_num3 = var_num3 + CInt(IIf(IsNull(var_51), var_19, var_121))
'0071dbd8    0f8021050000            jo 71e0ff
'0071dbde    66034dc4                add cx, word ptr [ebp-3c]
var_num3 = var_num3 + CInt(IIf(IsNull(var_51), var_19, var_121))
'0071dbe2    0f8017050000            jo 71e0ff
'0071dbe8    66898d40feffff          mov word ptr [ebp+fffffe40], cx
'0071dbef    c78538feffff02000000    mov dword ptr [ebp+fffffe38], 00000002
'0071dbf9    c78530feffff44ed4300    mov dword ptr [ebp+fffffe30], 0043ed44
'0071dc03    c78528feffff08000000    mov dword ptr [ebp+fffffe28], 00000008
'0071dc0d    8b55d4                  mov edx, dword ptr [ebp-2c]
'0071dc10    899520feffff            mov dword ptr [ebp+fffffe20], edx
'0071dc16    c78518feffff08000000    mov dword ptr [ebp+fffffe18], 00000008
'0071dc20    c78510feffff44ed4300    mov dword ptr [ebp+fffffe10], 0043ed44
'0071dc2a    c78508feffff08000000    mov dword ptr [ebp+fffffe08], 00000008
'0071dc34    c78500feffff70164500    mov dword ptr [ebp+fffffe00], 00451670
'0071dc3e    c785f8fdffff08000000    mov dword ptr [ebp+fffffdf8], 00000008
'0071dc48    8d8578ffffff            lea eax, dword ptr [ebp+ffffff78]
'0071dc4e    50                      push eax
'0071dc4f    8d8d88feffff            lea ecx, dword ptr [ebp+fffffe88]
'0071dc55    51                      push ecx
'0071dc56    8d9568ffffff            lea edx, dword ptr [ebp+ffffff68]
'0071dc5c    52                      push edx

' *** Reference to "__vbaVarCat"
'0071dc5d    ff1508124000            call dword ptr [00401208]
'0071dc63    50                      push eax
'0071dc64    8d8558ffffff            lea eax, dword ptr [ebp+ffffff58]
'0071dc6a    50                      push eax
'0071dc6b    8d8d48ffffff            lea ecx, dword ptr [ebp+ffffff48]
'0071dc71    51                      push ecx

' *** Reference to "__vbaVarCat"
'0071dc72    ff1508124000            call dword ptr [00401208]
'0071dc78    50                      push eax
'0071dc79    8d9568feffff            lea edx, dword ptr [ebp+fffffe68]
'0071dc7f    52                      push edx
'0071dc80    8d8538ffffff            lea eax, dword ptr [ebp+ffffff38]
'0071dc86    50                      push eax

' *** Reference to "__vbaVarCat"
'0071dc87    ff1508124000            call dword ptr [00401208]
'0071dc8d    50                      push eax
'0071dc8e    8d8d28ffffff            lea ecx, dword ptr [ebp+ffffff28]
'0071dc94    51                      push ecx
'0071dc95    8d9518ffffff            lea edx, dword ptr [ebp+ffffff18]
'0071dc9b    52                      push edx

' *** Reference to "__vbaVarCat"
'0071dc9c    ff1508124000            call dword ptr [00401208]
'0071dca2    50                      push eax
'0071dca3    8d8548feffff            lea eax, dword ptr [ebp+fffffe48]
'0071dca9    50                      push eax
'0071dcaa    8d8d08ffffff            lea ecx, dword ptr [ebp+ffffff08]
'0071dcb0    51                      push ecx

' *** Reference to "__vbaVarCat"
'0071dcb1    ff1508124000            call dword ptr [00401208]
'0071dcb7    50                      push eax
'0071dcb8    8d9538feffff            lea edx, dword ptr [ebp+fffffe38]
'0071dcbe    52                      push edx
'0071dcbf    8d85f8feffff            lea eax, dword ptr [ebp+fffffef8]
'0071dcc5    50                      push eax

' *** Reference to "__vbaVarCat"
'0071dcc6    ff1508124000            call dword ptr [00401208]
'0071dccc    50                      push eax
'0071dccd    8d8d28feffff            lea ecx, dword ptr [ebp+fffffe28]
'0071dcd3    51                      push ecx
'0071dcd4    8d95e8feffff            lea edx, dword ptr [ebp+fffffee8]
'0071dcda    52                      push edx

' *** Reference to "__vbaVarCat"
'0071dcdb    ff1508124000            call dword ptr [00401208]
'0071dce1    50                      push eax
'0071dce2    8d8518feffff            lea eax, dword ptr [ebp+fffffe18]
'0071dce8    50                      push eax
'0071dce9    8d8dd8feffff            lea ecx, dword ptr [ebp+fffffed8]
'0071dcef    51                      push ecx

' *** Reference to "__vbaVarCat"
'0071dcf0    ff1508124000            call dword ptr [00401208]
'0071dcf6    50                      push eax
'0071dcf7    8d9508feffff            lea edx, dword ptr [ebp+fffffe08]
'0071dcfd    52                      push edx
'0071dcfe    8d85c8feffff            lea eax, dword ptr [ebp+fffffec8]
'0071dd04    50                      push eax

' *** Reference to "__vbaVarCat"
'0071dd05    ff1508124000            call dword ptr [00401208]
'0071dd0b    50                      push eax
'0071dd0c    8d8df8fdffff            lea ecx, dword ptr [ebp+fffffdf8]
'0071dd12    51                      push ecx
'0071dd13    8d95b8feffff            lea edx, dword ptr [ebp+fffffeb8]
'0071dd19    52                      push edx

' *** Reference to "__vbaVarCat"
'0071dd1a    ff1508124000            call dword ptr [00401208]
'0071dd20    50                      push eax

' *** Reference to "__vbaStrVarMove"
'0071dd21    ff153c104000            call dword ptr [0040103c]
'0071dd27    8985b0feffff            mov dword ptr [ebp+fffffeb0], eax
'0071dd2d    c785a8feffff08000000    mov dword ptr [ebp+fffffea8], 00000008
'0071dd37    b810000000              mov eax, 00000010

' *** Reference to "__vbaChkstk"
'0071dd3c    e81f95ceff              call 407260
'0071dd41    8bc4                    mov eax, esp
'0071dd43    8b8da8feffff            mov ecx, dword ptr [ebp+fffffea8]
'0071dd49    8908                    mov dword ptr [eax], ecx
'0071dd4b    8b95acfeffff            mov edx, dword ptr [ebp+fffffeac]
'0071dd51    895004                  mov dword ptr [eax+04], edx
'0071dd54    8b8db0feffff            mov ecx, dword ptr [ebp+fffffeb0]
'0071dd5a    894808                  mov dword ptr [eax+08], ecx
'0071dd5d    8b95b4feffff            mov edx, dword ptr [ebp+fffffeb4]
'0071dd63    89500c                  mov dword ptr [eax+0c], edx
'0071dd66    6a01                    push 01
'0071dd68    6880000000              push 00000080
'0071dd6d    8b4508                  mov eax, dword ptr [ebp+08]
'0071dd70    8b08                    mov ecx, dword ptr [eax]
'0071dd72    8b5508                  mov edx, dword ptr [ebp+08]
'0071dd75    52                      push edx
'0071dd76    ff9110030000            call dword ptr [ecx+00000310]
'0071dd7c    50                      push eax
'0071dd7d    8d4588                  lea eax, dword ptr [ebp-78]
'0071dd80    50                      push eax

' *** Reference to "__vbaObjSet"
'0071dd81    ff15b4104000            call dword ptr [004010b4]
Set var_131 = Me
'0071dd87    50                      push eax

' *** Reference to "__vbaLateIdCall"
'0071dd88    ff1538104000            call dword ptr [00401038]
Call frmImport.Member_0x80(#NOT SUPPORTED#)
'0071dd8e    83c41c                  add esp, 1c
'0071dd91    8d4d88                  lea ecx, dword ptr [ebp-78]
'0071dd94    51                      push ecx
'0071dd95    8d558c                  lea edx, dword ptr [ebp-74]
'0071dd98    52                      push edx
'0071dd99    8d4590                  lea eax, dword ptr [ebp-70]
'0071dd9c    50                      push eax
'0071dd9d    8d4d94                  lea ecx, dword ptr [ebp-6c]
'0071dda0    51                      push ecx
'0071dda1    8d5598                  lea edx, dword ptr [ebp-68]
'0071dda4    52                      push edx
'0071dda5    8d459c                  lea eax, dword ptr [ebp-64]
'0071dda8    50                      push eax
'0071dda9    8d4da0                  lea ecx, dword ptr [ebp-60]
'0071ddac    51                      push ecx
'0071ddad    6a07                    push 07

' *** Reference to "__vbaFreeObjList"
'0071ddaf    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_7, var_51, var_130, var_121, var_8, var_52, var_131)
'0071ddb5    83c420                  add esp, 20
'0071ddb8    8d95a8feffff            lea edx, dword ptr [ebp+fffffea8]
'0071ddbe    52                      push edx
'0071ddbf    8d85b8feffff            lea eax, dword ptr [ebp+fffffeb8]
'0071ddc5    50                      push eax
'0071ddc6    8d8dc8feffff            lea ecx, dword ptr [ebp+fffffec8]
'0071ddcc    51                      push ecx
'0071ddcd    8d95d8feffff            lea edx, dword ptr [ebp+fffffed8]
'0071ddd3    52                      push edx
'0071ddd4    8d85e8feffff            lea eax, dword ptr [ebp+fffffee8]
'0071ddda    50                      push eax
'0071dddb    8d8df8feffff            lea ecx, dword ptr [ebp+fffffef8]
'0071dde1    51                      push ecx
'0071dde2    8d9508ffffff            lea edx, dword ptr [ebp+ffffff08]
'0071dde8    52                      push edx
'0071dde9    8d8518ffffff            lea eax, dword ptr [ebp+ffffff18]
'0071ddef    50                      push eax
'0071ddf0    8d8d28ffffff            lea ecx, dword ptr [ebp+ffffff28]
'0071ddf6    51                      push ecx
'0071ddf7    8d9538ffffff            lea edx, dword ptr [ebp+ffffff38]
'0071ddfd    52                      push edx
'0071ddfe    8d8548ffffff            lea eax, dword ptr [ebp+ffffff48]
'0071de04    50                      push eax
'0071de05    8d8d58ffffff            lea ecx, dword ptr [ebp+ffffff58]
'0071de0b    51                      push ecx
'0071de0c    8d9568ffffff            lea edx, dword ptr [ebp+ffffff68]
'0071de12    52                      push edx
'0071de13    8d8578ffffff            lea eax, dword ptr [ebp+ffffff78]
'0071de19    50                      push eax
'0071de1a    6a0e                    push 0e

' *** Reference to "__vbaFreeVarList"
'0071de1c    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_87, var_132, var_133, var_134, var_135, var_136, var_849, var_310, var_298, var_539, var_25, var_65, var_767, var_519)
'0071de22    83c43c                  add esp, 3c
'0071de25    c745fc32000000          mov dword ptr [ebp-04], 00000032
'0071de2c    8b4dac                  mov ecx, dword ptr [ebp-54]
'0071de2f    8b11                    mov edx, dword ptr [ecx]
'0071de31    8b45ac                  mov eax, dword ptr [ebp-54]
'0071de34    50                      push eax
'0071de35    ff92ec000000            call dword ptr [edx+000000ec]
'0071de3b    dbe2                    fnclex
'0071de3d    8985e0fdffff            mov dword ptr [ebp+fffffde0], eax
'0071de43    83bde0fdffff00          cmp dword ptr [ebp+fffffde0], 00
'0071de4a    7d23                    jge 71de6f

If (var_50 < 0) Then
'0071de4c    68ec000000              push 000000ec
'0071de51    6830314300              push 00433130
'0071de56    8b4dac                  mov ecx, dword ptr [ebp-54]
'0071de59    51                      push ecx
'0071de5a    8b95e0fdffff            mov edx, dword ptr [ebp+fffffde0]
'0071de60    52                      push edx

' *** Reference to "__vbaHresultCheckObj"
'0071de61    ff1580104000            call dword ptr [00401080]
'0071de67    898528fcffff            mov dword ptr [ebp+fffffc28], eax
'0071de6d    eb0a                    jmp 71de79
    
End If
'0071de6f    c78528fcffff00000000    mov dword ptr [ebp+fffffc28], 00000000
'0071de79    e94ed2ffff              jmp 71b0cc
'0071de7e    c745fc34000000          mov dword ptr [ebp-04], 00000034
'0071de85    8b45ac                  mov eax, dword ptr [ebp-54]
'0071de88    8b08                    mov ecx, dword ptr [eax]
'0071de8a    8b55ac                  mov edx, dword ptr [ebp-54]
'0071de8d    52                      push edx
'0071de8e    ff91c4000000            call dword ptr [ecx+000000c4]
'0071de94    dbe2                    fnclex
'0071de96    8985e0fdffff            mov dword ptr [ebp+fffffde0], eax
'0071de9c    83bde0fdffff00          cmp dword ptr [ebp+fffffde0], 00
'0071dea3    7d23                    jge 71dec8

If (var_50 < 0) Then
'0071dea5    68c4000000              push 000000c4
'0071deaa    6830314300              push 00433130
'0071deaf    8b45ac                  mov eax, dword ptr [ebp-54]
'0071deb2    50                      push eax
'0071deb3    8b8de0fdffff            mov ecx, dword ptr [ebp+fffffde0]
'0071deb9    51                      push ecx

' *** Reference to "__vbaHresultCheckObj"
'0071deba    ff1580104000            call dword ptr [00401080]
'0071dec0    898524fcffff            mov dword ptr [ebp+fffffc24], eax
'0071dec6    eb0a                    jmp 71ded2
    
End If
'0071dec8    c78524fcffff00000000    mov dword ptr [ebp+fffffc24], 00000000
'0071ded2    c745fc35000000          mov dword ptr [ebp-04], 00000035
'0071ded9    c785a0feffff05000000    mov dword ptr [ebp+fffffea0], 00000005
'0071dee3    c78598feffff03000000    mov dword ptr [ebp+fffffe98], 00000003
'0071deed    c78580feffff340d4500    mov dword ptr [ebp+fffffe80], 00450d34
'0071def7    c78578feffff08000000    mov dword ptr [ebp+fffffe78], 00000008
'0071df01    b810000000              mov eax, 00000010

' *** Reference to "__vbaChkstk"
'0071df06    e85593ceff              call 407260
'0071df0b    8bd4                    mov edx, esp
'0071df0d    8b8598feffff            mov eax, dword ptr [ebp+fffffe98]
'0071df13    8902                    mov dword ptr [edx], eax
'0071df15    8b8d9cfeffff            mov ecx, dword ptr [ebp+fffffe9c]
'0071df1b    894a04                  mov dword ptr [edx+04], ecx
'0071df1e    8b85a0feffff            mov eax, dword ptr [ebp+fffffea0]
'0071df24    894208                  mov dword ptr [edx+08], eax
'0071df27    8b8da4feffff            mov ecx, dword ptr [ebp+fffffea4]
'0071df2d    894a0c                  mov dword ptr [edx+0c], ecx
'0071df30    b810000000              mov eax, 00000010

' *** Reference to "__vbaChkstk"
'0071df35    e82693ceff              call 407260
'0071df3a    8bd4                    mov edx, esp
'0071df3c    8b8578feffff            mov eax, dword ptr [ebp+fffffe78]
'0071df42    8902                    mov dword ptr [edx], eax
'0071df44    8b8d7cfeffff            mov ecx, dword ptr [ebp+fffffe7c]
'0071df4a    894a04                  mov dword ptr [edx+04], ecx
'0071df4d    8b8580feffff            mov eax, dword ptr [ebp+fffffe80]
'0071df53    894208                  mov dword ptr [edx+08], eax
'0071df56    8b8d84feffff            mov ecx, dword ptr [ebp+fffffe84]
'0071df5c    894a0c                  mov dword ptr [edx+0c], ecx
'0071df5f    6a01                    push 01
'0071df61    68a5000000              push 000000a5
'0071df66    8b5508                  mov edx, dword ptr [ebp+08]
'0071df69    8b02                    mov eax, dword ptr [edx]
'0071df6b    8b4d08                  mov ecx, dword ptr [ebp+08]
'0071df6e    51                      push ecx
'0071df6f    ff9010030000            call dword ptr [eax+00000310]
'0071df75    50                      push eax
'0071df76    8d55a0                  lea edx, dword ptr [ebp-60]
'0071df79    52                      push edx

' *** Reference to "__vbaObjSet"
'0071df7a    ff15b4104000            call dword ptr [004010b4]
Set var_7 = Nothing
'0071df80    50                      push eax

' *** Reference to "__vbaLateIdCallSt"
'0071df81    ff159c114000            call dword ptr [0040119c]
'0071df87    83c42c                  add esp, 2c
'0071df8a    8d4da0                  lea ecx, dword ptr [ebp-60]

' *** Reference to "__vbaFreeObj"
'0071df8d    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_7)
'0071df93    c745fc36000000          mov dword ptr [ebp-04], 00000036
'0071df9a    8b4508                  mov eax, dword ptr [ebp+08]
'0071df9d    8b08                    mov ecx, dword ptr [eax]
'0071df9f    8b5508                  mov edx, dword ptr [ebp+08]
'0071dfa2    52                      push edx
'0071dfa3    ff9104030000            call dword ptr [ecx+00000304]
'0071dfa9    50                      push eax
'0071dfaa    8d45a0                  lea eax, dword ptr [ebp-60]
'0071dfad    50                      push eax

' *** Reference to "__vbaObjSet"
'0071dfae    ff15b4104000            call dword ptr [004010b4]
Set var_7 = Me
'0071dfb4    8985e0fdffff            mov dword ptr [ebp+fffffde0], eax
'0071dfba    6a01                    push 01
'0071dfbc    8b8de0fdffff            mov ecx, dword ptr [ebp+fffffde0]
'0071dfc2    8b11                    mov edx, dword ptr [ecx]
'0071dfc4    8b85e0fdffff            mov eax, dword ptr [ebp+fffffde0]
'0071dfca    50                      push eax
'0071dfcb    ff92e4000000            call dword ptr [edx+000000e4]
'0071dfd1    dbe2                    fnclex
'0071dfd3    8985dcfdffff            mov dword ptr [ebp+fffffddc], eax
'0071dfd9    83bddcfdffff00          cmp dword ptr [ebp+fffffddc], 00
'0071dfe0    7d26                    jge 71e008

If (var_7 < 0) Then
'0071dfe2    68e4000000              push 000000e4
'0071dfe7    68dce24300              push 0043e2dc
'0071dfec    8b8de0fdffff            mov ecx, dword ptr [ebp+fffffde0]
'0071dff2    51                      push ecx
'0071dff3    8b95dcfdffff            mov edx, dword ptr [ebp+fffffddc]
'0071dff9    52                      push edx

' *** Reference to "__vbaHresultCheckObj"
'0071dffa    ff1580104000            call dword ptr [00401080]
'0071e000    898520fcffff            mov dword ptr [ebp+fffffc20], eax
'0071e006    eb0a                    jmp 71e012
    
End If
'0071e008    c78520fcffff00000000    mov dword ptr [ebp+fffffc20], 00000000
'0071e012    8d4da0                  lea ecx, dword ptr [ebp-60]

' *** Reference to "__vbaFreeObj"
'0071e015    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_7)
'0071e01b    c745f000000000          mov dword ptr [ebp-10], 00000000
'0071e022    68dde07100              push 0071e0dd
'0071e027    e995000000              jmp 71e0c1
'0071e02c    8d4588                  lea eax, dword ptr [ebp-78]
'0071e02f    50                      push eax
'0071e030    8d4d8c                  lea ecx, dword ptr [ebp-74]
'0071e033    51                      push ecx
'0071e034    8d5590                  lea edx, dword ptr [ebp-70]
'0071e037    52                      push edx
'0071e038    8d4594                  lea eax, dword ptr [ebp-6c]
'0071e03b    50                      push eax
'0071e03c    8d4d98                  lea ecx, dword ptr [ebp-68]
'0071e03f    51                      push ecx
'0071e040    8d559c                  lea edx, dword ptr [ebp-64]
'0071e043    52                      push edx
'0071e044    8d45a0                  lea eax, dword ptr [ebp-60]
'0071e047    50                      push eax
'0071e048    6a07                    push 07

' *** Reference to "__vbaFreeObjList"
'0071e04a    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_7, var_51, var_130, var_121, var_8, var_52, var_131)
'0071e050    83c420                  add esp, 20
'0071e053    8d8da8feffff            lea ecx, dword ptr [ebp+fffffea8]
'0071e059    51                      push ecx
'0071e05a    8d95b8feffff            lea edx, dword ptr [ebp+fffffeb8]
'0071e060    52                      push edx
'0071e061    8d85c8feffff            lea eax, dword ptr [ebp+fffffec8]
'0071e067    50                      push eax
'0071e068    8d8dd8feffff            lea ecx, dword ptr [ebp+fffffed8]
'0071e06e    51                      push ecx
'0071e06f    8d95e8feffff            lea edx, dword ptr [ebp+fffffee8]
'0071e075    52                      push edx
'0071e076    8d85f8feffff            lea eax, dword ptr [ebp+fffffef8]
'0071e07c    50                      push eax
'0071e07d    8d8d08ffffff            lea ecx, dword ptr [ebp+ffffff08]
'0071e083    51                      push ecx
'0071e084    8d9518ffffff            lea edx, dword ptr [ebp+ffffff18]
'0071e08a    52                      push edx
'0071e08b    8d8528ffffff            lea eax, dword ptr [ebp+ffffff28]
'0071e091    50                      push eax
'0071e092    8d8d38ffffff            lea ecx, dword ptr [ebp+ffffff38]
'0071e098    51                      push ecx
'0071e099    8d9548ffffff            lea edx, dword ptr [ebp+ffffff48]
'0071e09f    52                      push edx
'0071e0a0    8d8558ffffff            lea eax, dword ptr [ebp+ffffff58]
'0071e0a6    50                      push eax
'0071e0a7    8d8d68ffffff            lea ecx, dword ptr [ebp+ffffff68]
'0071e0ad    51                      push ecx
'0071e0ae    8d9578ffffff            lea edx, dword ptr [ebp+ffffff78]
'0071e0b4    52                      push edx
'0071e0b5    6a0e                    push 0e

' *** Reference to "__vbaFreeVarList"
'0071e0b7    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_87, var_132, var_133, var_134, var_135, var_136, var_849, var_310, var_298, var_539, var_25, var_65, var_767, var_519)
'0071e0bd    83c43c                  add esp, 3c
'0071e0c0    c3                      ret
'0071e0c1    8d4dd4                  lea ecx, dword ptr [ebp-2c]

' *** Reference to "__vbaFreeStr"
'0071e0c4    ff150c134000            call dword ptr [0040130c]
'#Cleanup(var_44)
'0071e0ca    8d4db4                  lea ecx, dword ptr [ebp-4c]

' *** Reference to "__vbaFreeVar"
'0071e0cd    ff1528104000            call dword ptr [00401028]
'#Cleanup(var_62)
'0071e0d3    8d4dac                  lea ecx, dword ptr [ebp-54]

' *** Reference to "__vbaFreeObj"
'0071e0d6    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_50)
'0071e0dc    c3                      ret
'0071e0dd    8b4508                  mov eax, dword ptr [ebp+08]
'0071e0e0    8b08                    mov ecx, dword ptr [eax]
'0071e0e2    8b5508                  mov edx, dword ptr [ebp+08]
'0071e0e5    52                      push edx
'0071e0e6    ff5108                  call dword ptr [ecx+08]
'0071e0e9    8b45f0                  mov eax, dword ptr [ebp-10]
'0071e0ec    8b4de0                  mov ecx, dword ptr [ebp-20]
    'Reference to '__except_list'
'0071e0ef    64890d00000000          mov dword ptr fs:[00000000], ecx
'0071e0f6    5f                      pop edi
'0071e0f7    5e                      pop esi
'0071e0f8    5b                      pop ebx
'0071e0f9    8be5                    mov esp, ebp
'0071e0fb    5d                      pop ebp
'0071e0fc    c20400                  ret 0004


End Sub


Private Sub Form_Unload(Cancel as Integer)
'0071e690    55                      push ebp
'0071e691    8bec                    mov ebp, esp
'0071e693    83ec0c                  sub esp, 0c

' *** Reference to "__vbaExceptHandler"
'0071e696    6866724000              push 00407266
'0071e69b    64a100000000            mov ax, word ptr fs:[00000000]
'0071e6a1    50                      push eax
    'Reference to '__except_list'
'0071e6a2    64892500000000          mov dword ptr fs:[00000000], esp
'0071e6a9    83ec0c                  sub esp, 0c
'0071e6ac    53                      push ebx
'0071e6ad    56                      push esi
'0071e6ae    57                      push edi
'0071e6af    8965f4                  mov dword ptr [ebp-0c], esp
'0071e6b2    c745f8486f4000          mov dword ptr [ebp-08], 00406f48
'0071e6b9    8b7508                  mov esi, dword ptr [ebp+08]
'0071e6bc    8bc6                    mov eax, esi
'0071e6be    83e001                  and eax, 01
'0071e6c1    8945fc                  mov dword ptr [ebp-04], eax
'0071e6c4    83e6fe                  and esi, fffffffe
'0071e6c7    8b0e                    mov ecx, dword ptr [esi]
'0071e6c9    56                      push esi
'0071e6ca    897508                  mov dword ptr [ebp+08], esi
'0071e6cd    ff5104                  call dword ptr [ecx+04]
'0071e6d0    8b4638                  mov eax, dword ptr [esi+38]
'0071e6d3    8b10                    mov edx, dword ptr [eax]
'0071e6d5    50                      push eax
'0071e6d6    ff5258                  call dword ptr [edx+58]
'0071e6d9    dbe2                    fnclex
'0071e6db    85c0                    test ax, ax
'0071e6dd    7d12                    jge 71e6f1
'0071e6df    8b4e38                  mov ecx, dword ptr [esi+38]
'0071e6e2    6a58                    push 58
'0071e6e4    68301f4300              push 00431f30
'0071e6e9    51                      push ecx
'0071e6ea    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0071e6eb    ff1580104000            call dword ptr [00401080]
'0071e6f1    c745fc00000000          mov dword ptr [ebp-04], 00000000
'0071e6f8    8b4508                  mov eax, dword ptr [ebp+08]
'0071e6fb    8b10                    mov edx, dword ptr [eax]
'0071e6fd    50                      push eax
'0071e6fe    ff5208                  call dword ptr [edx+08]
'0071e701    8b45fc                  mov eax, dword ptr [ebp-04]
'0071e704    8b4dec                  mov ecx, dword ptr [ebp-14]
'0071e707    5f                      pop edi
'0071e708    5e                      pop esi
    'Reference to '__except_list'
'0071e709    64890d00000000          mov dword ptr fs:[00000000], ecx
'0071e710    5b                      pop ebx
'0071e711    8be5                    mov esp, ebp
'0071e713    5d                      pop ebp
'0071e714    c20800                  ret 0008


End Sub


Private Sub Form_KeyDown(KeyCode as Integer, Shift as Integer)
'0071ac40    55                      push ebp
'0071ac41    8bec                    mov ebp, esp
'0071ac43    83ec0c                  sub esp, 0c

' *** Reference to "__vbaExceptHandler"
'0071ac46    6866724000              push 00407266
'0071ac4b    64a100000000            mov ax, word ptr fs:[00000000]
'0071ac51    50                      push eax
    'Reference to '__except_list'
'0071ac52    64892500000000          mov dword ptr fs:[00000000], esp
'0071ac59    83ec18                  sub esp, 18
'0071ac5c    53                      push ebx
'0071ac5d    56                      push esi
'0071ac5e    57                      push edi
'0071ac5f    8965f4                  mov dword ptr [ebp-0c], esp
'0071ac62    c745f8106e4000          mov dword ptr [ebp-08], 00406e10
'0071ac69    8b7d08                  mov edi, dword ptr [ebp+08]
'0071ac6c    8bc7                    mov eax, edi
'0071ac6e    83e001                  and eax, 01
'0071ac71    8945fc                  mov dword ptr [ebp-04], eax
'0071ac74    83e7fe                  and edi, fffffffe
'0071ac77    8b0f                    mov ecx, dword ptr [edi]
'0071ac79    57                      push edi
'0071ac7a    897d08                  mov dword ptr [ebp+08], edi
'0071ac7d    ff5104                  call dword ptr [ecx+04]
'0071ac80    8b550c                  mov edx, dword ptr [ebp+0c]
'0071ac83    33db                    xor ebx, ebx
'0071ac85    66833a1b                cmp word ptr [edx], 1b
'0071ac89    895de8                  mov dword ptr [ebp-18], ebx
'0071ac8c    7554                    jne 71ace2

If (arg_0 = 27) Then
'0071ac8e    391d24be7200            cmp dword ptr [0072be24], ebx
'0071ac94    7510                    jne 71aca6
'0071ac96    6824be7200              push 0072be24
'0071ac9b    6870004300              push 00430070

' *** Reference to "__vbaNew2"
'0071aca0    ff152c124000            call dword ptr [0040122c]
    Dim var_2 As New Global
'0071aca6    8b3524be7200            mov esi, dword ptr [0072be24]
'0071acac    8b16                    mov edx, dword ptr [esi]
'0071acae    57                      push edi
'0071acaf    8d45e8                  lea eax, dword ptr [ebp-18]
'0071acb2    50                      push eax
'0071acb3    8955d4                  mov dword ptr [ebp-2c], edx

' *** Reference to "__vbaObjSetAddref"
'0071acb6    ff15c8104000            call dword ptr [004010c8]
    Set var_41 = Me
'0071acbc    8b4dd4                  mov ecx, dword ptr [ebp-2c]
'0071acbf    50                      push eax
'0071acc0    56                      push esi
'0071acc1    ff5110                  call dword ptr [ecx+10]
    Call var_2.Unload(var_41)
'0071acc4    dbe2                    fnclex
'0071acc6    3bc3                    cmp eax, ebx
'0071acc8    7d0f                    jge 71acd9
'0071acca    6a10                    push 10
'0071accc    6860004300              push 00430060
'0071acd1    56                      push esi
'0071acd2    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0071acd3    ff1580104000            call dword ptr [00401080]
'0071acd9    8d4de8                  lea ecx, dword ptr [ebp-18]

' *** Reference to "__vbaFreeObj"
'0071acdc    ff1508134000            call dword ptr [00401308]
    '#Cleanup(var_41)
End If
'0071ace2    895dfc                  mov dword ptr [ebp-04], ebx
'0071ace5    68f7ac7100              push 0071acf7
'0071acea    eb0a                    jmp 71acf6
'0071acec    8d4de8                  lea ecx, dword ptr [ebp-18]

' *** Reference to "__vbaFreeObj"
'0071acef    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_41)
'0071acf5    c3                      ret
'0071acf6    c3                      ret
'0071acf7    8b4508                  mov eax, dword ptr [ebp+08]
'0071acfa    8b10                    mov edx, dword ptr [eax]
'0071acfc    50                      push eax
'0071acfd    ff5208                  call dword ptr [edx+08]
'0071ad00    8b45fc                  mov eax, dword ptr [ebp-04]
'0071ad03    8b4dec                  mov ecx, dword ptr [ebp-14]
'0071ad06    5f                      pop edi
'0071ad07    5e                      pop esi
    'Reference to '__except_list'
'0071ad08    64890d00000000          mov dword ptr fs:[00000000], ecx
'0071ad0f    5b                      pop ebx
'0071ad10    8be5                    mov esp, ebp
'0071ad12    5d                      pop ebp
'0071ad13    c20c00                  ret 000c


End Sub


'Event for btnTous
Private Sub btnTous_Click()
'0071a880    55                      push ebp
'0071a881    8bec                    mov ebp, esp
'0071a883    83ec0c                  sub esp, 0c

' *** Reference to "__vbaExceptHandler"
'0071a886    6866724000              push 00407266
'0071a88b    64a100000000            mov ax, word ptr fs:[00000000]
'0071a891    50                      push eax
    'Reference to '__except_list'
'0071a892    64892500000000          mov dword ptr fs:[00000000], esp
'0071a899    83ec10                  sub esp, 10
'0071a89c    53                      push ebx
'0071a89d    56                      push esi
'0071a89e    57                      push edi
'0071a89f    8965f4                  mov dword ptr [ebp-0c], esp
'0071a8a2    c745f8e06d4000          mov dword ptr [ebp-08], 00406de0
'0071a8a9    8b7508                  mov esi, dword ptr [ebp+08]
'0071a8ac    8bc6                    mov eax, esi
'0071a8ae    83e001                  and eax, 01
'0071a8b1    8945fc                  mov dword ptr [ebp-04], eax
'0071a8b4    83e6fe                  and esi, fffffffe
'0071a8b7    8b0e                    mov ecx, dword ptr [esi]
'0071a8b9    56                      push esi
'0071a8ba    897508                  mov dword ptr [ebp+08], esi
'0071a8bd    ff5104                  call dword ptr [ecx+04]
'0071a8c0    33ff                    xor edi, edi
'0071a8c2    bafce94400              mov edx, 0044e9fc
'0071a8c7    8d4de8                  lea ecx, dword ptr [ebp-18]
'0071a8ca    897de8                  mov dword ptr [ebp-18], edi

' *** Reference to "__vbaStrCopy"
'0071a8cd    ff1548124000            call dword ptr [00401248]
var_41 = ("Oui")
'0071a8d3    8b16                    mov edx, dword ptr [esi]
'0071a8d5    8d45e8                  lea eax, dword ptr [ebp-18]
'0071a8d8    50                      push eax
'0071a8d9    56                      push esi
'0071a8da    ff920c070000            call dword ptr [edx+0000070c]
'0071a8e0    3bc7                    cmp eax, edi
'0071a8e2    7d12                    jge 71a8f6
'0071a8e4    680c070000              push 0000070c
'0071a8e9    68601e4300              push 00431e60
'0071a8ee    56                      push esi
'0071a8ef    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0071a8f0    ff1580104000            call dword ptr [00401080]
'0071a8f6    8d4de8                  lea ecx, dword ptr [ebp-18]

' *** Reference to "__vbaFreeStr"
'0071a8f9    ff150c134000            call dword ptr [0040130c]
'#Cleanup(var_41)
'0071a8ff    897dfc                  mov dword ptr [ebp-04], edi
'0071a902    6814a97100              push 0071a914
'0071a907    eb0a                    jmp 71a913
'0071a909    8d4de8                  lea ecx, dword ptr [ebp-18]

' *** Reference to "__vbaFreeStr"
'0071a90c    ff150c134000            call dword ptr [0040130c]
'#Cleanup(var_41)
'0071a912    c3                      ret
'0071a913    c3                      ret
'0071a914    8b4508                  mov eax, dword ptr [ebp+08]
'0071a917    8b08                    mov ecx, dword ptr [eax]
'0071a919    50                      push eax
'0071a91a    ff5108                  call dword ptr [ecx+08]
'0071a91d    8b45fc                  mov eax, dword ptr [ebp-04]
'0071a920    8b4dec                  mov ecx, dword ptr [ebp-14]
'0071a923    5f                      pop edi
'0071a924    5e                      pop esi
    'Reference to '__except_list'
'0071a925    64890d00000000          mov dword ptr fs:[00000000], ecx
'0071a92c    5b                      pop ebx
'0071a92d    8be5                    mov esp, ebp
'0071a92f    5d                      pop ebp
'0071a930    c20400                  ret 0004


End Sub


'Event for BtnImport
Private Sub BtnImport_Click()
'00714bc0    55                      push ebp
'00714bc1    8bec                    mov ebp, esp
'00714bc3    83ec0c                  sub esp, 0c

' *** Reference to "__vbaExceptHandler"
'00714bc6    6866724000              push 00407266
'00714bcb    64a100000000            mov ax, word ptr fs:[00000000]
'00714bd1    50                      push eax
    'Reference to '__except_list'
'00714bd2    64892500000000          mov dword ptr fs:[00000000], esp
'00714bd9    81ec0c030000            sub esp, 0000030c
'00714bdf    53                      push ebx
'00714be0    56                      push esi
'00714be1    57                      push edi
'00714be2    8965f4                  mov dword ptr [ebp-0c], esp
'00714be5    c745f8d06d4000          mov dword ptr [ebp-08], 00406dd0
'00714bec    8b7d08                  mov edi, dword ptr [ebp+08]
'00714bef    8bc7                    mov eax, edi
'00714bf1    83e001                  and eax, 01
'00714bf4    8945fc                  mov dword ptr [ebp-04], eax
'00714bf7    83e7fe                  and edi, fffffffe
'00714bfa    8b0f                    mov ecx, dword ptr [edi]
'00714bfc    57                      push edi
'00714bfd    897d08                  mov dword ptr [ebp+08], edi
'00714c00    ff5104                  call dword ptr [ecx+04]
'00714c03    33f6                    xor esi, esi
'00714c05    8b17                    mov edx, dword ptr [edi]
'00714c07    57                      push edi
'00714c08    8975dc                  mov dword ptr [ebp-24], esi
'00714c0b    8975d8                  mov dword ptr [ebp-28], esi
'00714c0e    8975d4                  mov dword ptr [ebp-2c], esi
'00714c11    8975d0                  mov dword ptr [ebp-30], esi
'00714c14    8975cc                  mov dword ptr [ebp-34], esi
'00714c17    8975c8                  mov dword ptr [ebp-38], esi
'00714c1a    8975c4                  mov dword ptr [ebp-3c], esi
'00714c1d    8975c0                  mov dword ptr [ebp-40], esi
'00714c20    8975bc                  mov dword ptr [ebp-44], esi
'00714c23    8975b8                  mov dword ptr [ebp-48], esi
'00714c26    8975b4                  mov dword ptr [ebp-4c], esi
'00714c29    8975b0                  mov dword ptr [ebp-50], esi
'00714c2c    8975ac                  mov dword ptr [ebp-54], esi
'00714c2f    89759c                  mov dword ptr [ebp-64], esi
'00714c32    89758c                  mov dword ptr [ebp-74], esi
'00714c35    89b57cffffff            mov dword ptr [ebp+ffffff7c], esi
'00714c3b    89b56cffffff            mov dword ptr [ebp+ffffff6c], esi
'00714c41    89b55cffffff            mov dword ptr [ebp+ffffff5c], esi
'00714c47    89b51cffffff            mov dword ptr [ebp+ffffff1c], esi
'00714c4d    89b50cffffff            mov dword ptr [ebp+ffffff0c], esi
'00714c53    89b5fcfeffff            mov dword ptr [ebp+fffffefc], esi
'00714c59    89b5ecfeffff            mov dword ptr [ebp+fffffeec], esi
'00714c5f    89b5dcfeffff            mov dword ptr [ebp+fffffedc], esi
'00714c65    89b5ccfeffff            mov dword ptr [ebp+fffffecc], esi
'00714c6b    89b5bcfeffff            mov dword ptr [ebp+fffffebc], esi
'00714c71    89b5acfeffff            mov dword ptr [ebp+fffffeac], esi
'00714c77    89b59cfeffff            mov dword ptr [ebp+fffffe9c], esi
'00714c7d    89b58cfeffff            mov dword ptr [ebp+fffffe8c], esi
'00714c83    89b57cfeffff            mov dword ptr [ebp+fffffe7c], esi
'00714c89    89b56cfeffff            mov dword ptr [ebp+fffffe6c], esi
'00714c8f    89b55cfeffff            mov dword ptr [ebp+fffffe5c], esi
'00714c95    89b54cfeffff            mov dword ptr [ebp+fffffe4c], esi
'00714c9b    89b53cfeffff            mov dword ptr [ebp+fffffe3c], esi
'00714ca1    89b538feffff            mov dword ptr [ebp+fffffe38], esi
'00714ca7    89b514feffff            mov dword ptr [ebp+fffffe14], esi
'00714cad    89b510feffff            mov dword ptr [ebp+fffffe10], esi
'00714cb3    ff9204030000            call dword ptr [edx+00000304]
'00714cb9    50                      push eax
'00714cba    8d45b8                  lea eax, dword ptr [ebp-48]
'00714cbd    50                      push eax

' *** Reference to "__vbaObjSet"
'00714cbe    ff15b4104000            call dword ptr [004010b4]
Set var_61 = Me
'00714cc4    8d9538feffff            lea edx, dword ptr [ebp+fffffe38]
'00714cca    8bd8                    mov ebx, eax
'00714ccc    8b0b                    mov ecx, dword ptr [ebx]
'00714cce    52                      push edx
'00714ccf    53                      push ebx
'00714cd0    ff91e0000000            call dword ptr [ecx+000000e0]
'00714cd6    dbe2                    fnclex
'00714cd8    3bc6                    cmp eax, esi
'00714cda    7d12                    jge 714cee

If (var_61 < 0) Then
'00714cdc    68e0000000              push 000000e0
'00714ce1    68dce24300              push 0043e2dc
'00714ce6    53                      push ebx
'00714ce7    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00714ce8    ff1580104000            call dword ptr [00401080]
    
End If
'00714cee    668b5f34                mov bx, word ptr [edi+34]
'00714cf2    8b8d38feffff            mov ecx, dword ptr [ebp+fffffe38]
'00714cf8    66f7d3                  not bx
'00714cfb    23d9                    and ebx, ecx
var_num2 = Not (arg_6) And 0
'00714cfd    8d4db8                  lea ecx, dword ptr [ebp-48]

' *** Reference to "__vbaFreeObj"
'00714d00    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_61)
'00714d06    663bde                  cmp bx, si
'00714d09    0f84fc220000            je 71700b

If (var_num2 <> 0) Then
'00714d0f    8d5db8                  lea ebx, dword ptr [ebp-48]
'00714d12    53                      push ebx
'00714d13    83ec10                  sub esp, 10
'00714d16    8bdc                    mov ebx, esp
'00714d18    8b4738                  mov eax, dword ptr [edi+38]
'00714d1b    b90a000000              mov ecx, 0000000a
'00714d20    890b                    mov dword ptr [ebx], ecx
'00714d22    8b8d30ffffff            mov ecx, dword ptr [ebp+ffffff30]
'00714d28    894b04                  mov dword ptr [ebx+04], ecx
'00714d2b    8b10                    mov edx, dword ptr [eax]
'00714d2d    b904000280              mov ecx, 80020004
'00714d32    894b08                  mov dword ptr [ebx+08], ecx
'00714d35    8b8d38ffffff            mov ecx, dword ptr [ebp+ffffff38]
'00714d3b    894b0c                  mov dword ptr [ebx+0c], ecx
'00714d3e    83ec10                  sub esp, 10
'00714d41    8bdc                    mov ebx, esp
'00714d43    b90a000000              mov ecx, 0000000a
'00714d48    890b                    mov dword ptr [ebx], ecx
'00714d4a    8b8d40ffffff            mov ecx, dword ptr [ebp+ffffff40]
'00714d50    894b04                  mov dword ptr [ebx+04], ecx
'00714d53    b904000280              mov ecx, 80020004
'00714d58    894b08                  mov dword ptr [ebx+08], ecx
'00714d5b    8b8d48ffffff            mov ecx, dword ptr [ebp+ffffff48]
'00714d61    894b0c                  mov dword ptr [ebx+0c], ecx
'00714d64    83ec10                  sub esp, 10
'00714d67    8bdc                    mov ebx, esp
'00714d69    b903000000              mov ecx, 00000003
'00714d6e    890b                    mov dword ptr [ebx], ecx
'00714d70    8b8d50ffffff            mov ecx, dword ptr [ebp+ffffff50]
'00714d76    894b04                  mov dword ptr [ebx+04], ecx
'00714d79    b901000000              mov ecx, 00000001
'00714d7e    894b08                  mov dword ptr [ebx+08], ecx
'00714d81    8b8d58ffffff            mov ecx, dword ptr [ebp+ffffff58]
'00714d87    689c484300              push 0043489c
'00714d8c    50                      push eax
'00714d8d    894b0c                  mov dword ptr [ebx+0c], ecx
'00714d90    ff92bc000000            call dword ptr [edx+000000bc]
'00714d96    dbe2                    fnclex
'00714d98    3bc6                    cmp eax, esi
'00714d9a    7d15                    jge 714db1
    
    If (    arg_6 < 0) Then
'00714d9c    8b5738                  mov edx, dword ptr [edi+38]
'00714d9f    68bc000000              push 000000bc
'00714da4    68301f4300              push 00431f30
'00714da9    52                      push edx
'00714daa    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00714dab    ff1580104000            call dword ptr [00401080]
    
End If
'00714db1    8b45b8                  mov eax, dword ptr [ebp-48]
'00714db4    50                      push eax
'00714db5    8d45d0                  lea eax, dword ptr [ebp-30]
'00714db8    50                      push eax
'00714db9    8975b8                  mov dword ptr [ebp-48], esi

' *** Reference to "__vbaObjSet"
'00714dbc    ff15b4104000            call dword ptr [004010b4]
Set var_4 = var_61
'00714dc2    8d5db8                  lea ebx, dword ptr [ebp-48]
'00714dc5    53                      push ebx
'00714dc6    83ec10                  sub esp, 10
'00714dc9    8bdc                    mov ebx, esp
'00714dcb    b90a000000              mov ecx, 0000000a
'00714dd0    890b                    mov dword ptr [ebx], ecx
'00714dd2    8b8d30ffffff            mov ecx, dword ptr [ebp+ffffff30]
'00714dd8    894b04                  mov dword ptr [ebx+04], ecx
'00714ddb    8b3d48b07200            mov edi, dword ptr [0072b048]
'00714de1    8b3f                    mov edi, dword ptr [edi]
'00714de3    b804000280              mov eax, 80020004
'00714de8    894308                  mov dword ptr [ebx+08], eax
'00714deb    8bd0                    mov edx, eax
'00714ded    8b8538ffffff            mov eax, dword ptr [ebp+ffffff38]
'00714df3    89430c                  mov dword ptr [ebx+0c], eax
'00714df6    83ec10                  sub esp, 10
'00714df9    8bcc                    mov ecx, esp
'00714dfb    b80a000000              mov eax, 0000000a
'00714e00    8901                    mov dword ptr [ecx], eax
'00714e02    8b8540ffffff            mov eax, dword ptr [ebp+ffffff40]
'00714e08    894104                  mov dword ptr [ecx+04], eax
'00714e0b    895108                  mov dword ptr [ecx+08], edx
'00714e0e    8b9548ffffff            mov edx, dword ptr [ebp+ffffff48]
'00714e14    89510c                  mov dword ptr [ecx+0c], edx
'00714e17    8b9550ffffff            mov edx, dword ptr [ebp+ffffff50]
'00714e1d    83ec10                  sub esp, 10
'00714e20    8bcc                    mov ecx, esp
'00714e22    b803000000              mov eax, 00000003
'00714e27    8901                    mov dword ptr [ecx], eax
'00714e29    895104                  mov dword ptr [ecx+04], edx
'00714e2c    b801000000              mov eax, 00000001
'00714e31    894108                  mov dword ptr [ecx+08], eax
'00714e34    8b8558ffffff            mov eax, dword ptr [ebp+ffffff58]
'00714e3a    89410c                  mov dword ptr [ecx+0c], eax
'00714e3d    8b0d48b07200            mov ecx, dword ptr [0072b048]
'00714e43    689c484300              push 0043489c
'00714e48    51                      push ecx
'00714e49    ff97bc000000            call dword ptr [edi+000000bc]
'00714e4f    dbe2                    fnclex
'00714e51    3bc6                    cmp eax, esi
'00714e53    7d18                    jge 714e6d

If (0 < 0) Then
'00714e55    8b1548b07200            mov edx, dword ptr [0072b048]
'00714e5b    68bc000000              push 000000bc
'00714e60    68301f4300              push 00431f30
'00714e65    52                      push edx
'00714e66    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00714e67    ff1580104000            call dword ptr [00401080]
    
End If
'00714e6d    8b45b8                  mov eax, dword ptr [ebp-48]
'00714e70    50                      push eax
'00714e71    8d45d4                  lea eax, dword ptr [ebp-2c]
'00714e74    50                      push eax
'00714e75    8975b8                  mov dword ptr [ebp-48], esi

' *** Reference to "__vbaObjSet"
'00714e78    ff15b4104000            call dword ptr [004010b4]
Set var_44 = Nothing
'00714e7e    8b4dd4                  mov ecx, dword ptr [ebp-2c]
'00714e81    51                      push ecx
'00714e82    8d9514feffff            lea edx, dword ptr [ebp+fffffe14]
'00714e88    52                      push edx

' *** Reference to "__vbaObjSetAddref"
'00714e89    ff15c8104000            call dword ptr [004010c8]
Set var_695 = Nothing
'00714e8f    8b8514feffff            mov eax, dword ptr [ebp+fffffe14]
'00714e95    8b08                    mov ecx, dword ptr [eax]
'00714e97    6830a84300              push 0043a830
'00714e9c    50                      push eax
'00714e9d    ff5144                  call dword ptr [ecx+44]
'00714ea0    dbe2                    fnclex
'00714ea2    3bc6                    cmp eax, esi
'00714ea4    7d15                    jge 714ebb

If (var_695 < 0) Then
'00714ea6    8b9514feffff            mov edx, dword ptr [ebp+fffffe14]
'00714eac    6a44                    push 44
'00714eae    6830314300              push 00433130
'00714eb3    52                      push edx
'00714eb4    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00714eb5    ff1580104000            call dword ptr [00401080]
    
End If

' *** Reference to "__vbaHresultCheckObj"
'00714ebb    8b3580104000            mov esi, dword ptr [00401080]
'00714ec1    8b45d0                  mov eax, dword ptr [ebp-30]
'00714ec4    8b08                    mov ecx, dword ptr [eax]
'00714ec6    8d9538feffff            lea edx, dword ptr [ebp+fffffe38]
'00714ecc    52                      push edx
'00714ecd    50                      push eax
'00714ece    ff5134                  call dword ptr [ecx+34]
'00714ed1    dbe2                    fnclex
'00714ed3    85c0                    test ax, ax
'00714ed5    7d0e                    jge 714ee5
'00714ed7    8b4dd0                  mov ecx, dword ptr [ebp-30]
'00714eda    6a34                    push 34
'00714edc    6830314300              push 00433130
'00714ee1    51                      push ecx
'00714ee2    50                      push eax
'00714ee3    ffd6                    call esi
'00714ee5    6683bd38feffff00        cmp word ptr [ebp+fffffe38], 00
'00714eed    0f85d4180000            jne 7167c7

Do While (0 = 0)
'00714ef3    8b45d0                  mov eax, dword ptr [ebp-30]
'00714ef6    8b10                    mov edx, dword ptr [eax]
'00714ef8    8d4db8                  lea ecx, dword ptr [ebp-48]
'00714efb    51                      push ecx
'00714efc    50                      push eax
'00714efd    ff92b4000000            call dword ptr [edx+000000b4]
'00714f03    dbe2                    fnclex
'00714f05    85c0                    test ax, ax
'00714f07    7d11                    jge 714f1a
'00714f09    8b55d0                  mov edx, dword ptr [ebp-30]
'00714f0c    68b4000000              push 000000b4
'00714f11    6830314300              push 00433130
'00714f16    52                      push edx
'00714f17    50                      push eax
'00714f18    ffd6                    call esi
'00714f1a    8b45b8                  mov eax, dword ptr [ebp-48]
'00714f1d    8b38                    mov edi, dword ptr [eax]
'00714f1f    8d5db4                  lea ebx, dword ptr [ebp-4c]
'00714f22    53                      push ebx
'00714f23    83ec10                  sub esp, 10
'00714f26    8bdc                    mov ebx, esp
'00714f28    ba08000000              mov edx, 00000008
'00714f2d    8913                    mov dword ptr [ebx], edx
'00714f2f    8b9550ffffff            mov edx, dword ptr [ebp+ffffff50]
'00714f35    895304                  mov dword ptr [ebx+04], edx
'00714f38    b9c4a54300              mov ecx, 0043a5c4
'00714f3d    894b08                  mov dword ptr [ebx+08], ecx
'00714f40    8b8d58ffffff            mov ecx, dword ptr [ebp+ffffff58]
'00714f46    50                      push eax
'00714f47    898530feffff            mov dword ptr [ebp+fffffe30], eax
'00714f4d    894b0c                  mov dword ptr [ebx+0c], ecx
'00714f50    ff5730                  call dword ptr [edi+30]
'00714f53    dbe2                    fnclex
'00714f55    85c0                    test ax, ax
'00714f57    7d11                    jge 714f6a
'00714f59    8b9530feffff            mov edx, dword ptr [ebp+fffffe30]
'00714f5f    6a30                    push 30
'00714f61    68d8304300              push 004330d8
'00714f66    52                      push edx
'00714f67    50                      push eax
'00714f68    ffd6                    call esi
'00714f6a    be0a000000              mov esi, 0000000a
'00714f6f    83ec10                  sub esp, 10
'00714f72    8bdc                    mov ebx, esp
'00714f74    8bce                    mov ecx, esi
'00714f76    890b                    mov dword ptr [ebx], ecx
'00714f78    8b8d90feffff            mov ecx, dword ptr [ebp+fffffe90]
'00714f7e    894b04                  mov dword ptr [ebx+04], ecx
'00714f81    ba04000280              mov edx, 80020004
'00714f86    8bc2                    mov eax, edx
'00714f88    894308                  mov dword ptr [ebx+08], eax
'00714f8b    8b8598feffff            mov eax, dword ptr [ebp+fffffe98]
'00714f91    89430c                  mov dword ptr [ebx+0c], eax
'00714f94    8b85a0feffff            mov eax, dword ptr [ebp+fffffea0]
'00714f9a    83ec10                  sub esp, 10
'00714f9d    8bcc                    mov ecx, esp
'00714f9f    8931                    mov dword ptr [ecx], esi
'00714fa1    894104                  mov dword ptr [ecx+04], eax
'00714fa4    895108                  mov dword ptr [ecx+08], edx
'00714fa7    8995f4feffff            mov dword ptr [ebp+fffffef4], edx
'00714fad    8b95a8feffff            mov edx, dword ptr [ebp+fffffea8]
'00714fb3    89510c                  mov dword ptr [ecx+0c], edx
'00714fb6    8b95b0feffff            mov edx, dword ptr [ebp+fffffeb0]
'00714fbc    83ec10                  sub esp, 10
'00714fbf    8bcc                    mov ecx, esp
'00714fc1    8bc6                    mov eax, esi
'00714fc3    8901                    mov dword ptr [ecx], eax
'00714fc5    895104                  mov dword ptr [ecx+04], edx
'00714fc8    8b95c0feffff            mov edx, dword ptr [ebp+fffffec0]
'00714fce    b804000280              mov eax, 80020004
'00714fd3    894108                  mov dword ptr [ecx+08], eax
'00714fd6    8b85b8feffff            mov eax, dword ptr [ebp+fffffeb8]
'00714fdc    89410c                  mov dword ptr [ecx+0c], eax
'00714fdf    83ec10                  sub esp, 10
'00714fe2    8bcc                    mov ecx, esp
'00714fe4    8bc6                    mov eax, esi
'00714fe6    8901                    mov dword ptr [ecx], eax
'00714fe8    895104                  mov dword ptr [ecx+04], edx
'00714feb    8b95d0feffff            mov edx, dword ptr [ebp+fffffed0]
'00714ff1    8b7db4                  mov edi, dword ptr [ebp-4c]
'00714ff4    b804000280              mov eax, 80020004
'00714ff9    894108                  mov dword ptr [ecx+08], eax
'00714ffc    8b85c8feffff            mov eax, dword ptr [ebp+fffffec8]
'00715002    89410c                  mov dword ptr [ecx+0c], eax
'00715005    83ec10                  sub esp, 10
'00715008    8bcc                    mov ecx, esp
'0071500a    8bc6                    mov eax, esi
'0071500c    8901                    mov dword ptr [ecx], eax
'0071500e    895104                  mov dword ptr [ecx+04], edx
'00715011    8b95e0feffff            mov edx, dword ptr [ebp+fffffee0]
'00715017    b804000280              mov eax, 80020004
'0071501c    894108                  mov dword ptr [ecx+08], eax
'0071501f    8b85d8feffff            mov eax, dword ptr [ebp+fffffed8]
'00715025    89410c                  mov dword ptr [ecx+0c], eax
'00715028    83ec10                  sub esp, 10
'0071502b    8bcc                    mov ecx, esp
'0071502d    8bc6                    mov eax, esi
'0071502f    8901                    mov dword ptr [ecx], eax
'00715031    895104                  mov dword ptr [ecx+04], edx
'00715034    b804000280              mov eax, 80020004
'00715039    894108                  mov dword ptr [ecx+08], eax
'0071503c    8b85e8feffff            mov eax, dword ptr [ebp+fffffee8]
'00715042    89410c                  mov dword ptr [ecx+0c], eax
'00715045    8b85f0feffff            mov eax, dword ptr [ebp+fffffef0]
'0071504b    83ec10                  sub esp, 10
'0071504e    8bcc                    mov ecx, esp
'00715050    8bd6                    mov edx, esi
'00715052    8911                    mov dword ptr [ecx], edx
'00715054    8b95f4feffff            mov edx, dword ptr [ebp+fffffef4]
'0071505a    894104                  mov dword ptr [ecx+04], eax
'0071505d    8b85f8feffff            mov eax, dword ptr [ebp+fffffef8]
'00715063    897da4                  mov dword ptr [ebp-5c], edi
'00715066    8bbd14feffff            mov edi, dword ptr [ebp+fffffe14]
'0071506c    895108                  mov dword ptr [ecx+08], edx
'0071506f    89410c                  mov dword ptr [ecx+0c], eax
'00715072    83ec10                  sub esp, 10
'00715075    89b5ecfeffff            mov dword ptr [ebp+fffffeec], esi
'0071507b    c745b400000000          mov dword ptr [ebp-4c], 00000000
'00715082    c7459c09000000          mov dword ptr [ebp-64], 00000009
'00715089    8b3f                    mov edi, dword ptr [edi]
'0071508b    8bcc                    mov ecx, esp
'0071508d    8bc6                    mov eax, esi
'0071508f    8901                    mov dword ptr [ecx], eax
'00715091    8b9500ffffff            mov edx, dword ptr [ebp+ffffff00]
'00715097    895104                  mov dword ptr [ecx+04], edx
'0071509a    8b9510ffffff            mov edx, dword ptr [ebp+ffffff10]
'007150a0    83ec10                  sub esp, 10
'007150a3    b804000280              mov eax, 80020004
'007150a8    894108                  mov dword ptr [ecx+08], eax
'007150ab    8b8508ffffff            mov eax, dword ptr [ebp+ffffff08]
'007150b1    89410c                  mov dword ptr [ecx+0c], eax
'007150b4    8bcc                    mov ecx, esp
'007150b6    8bc6                    mov eax, esi
'007150b8    8901                    mov dword ptr [ecx], eax
'007150ba    895104                  mov dword ptr [ecx+04], edx
'007150bd    8b9520ffffff            mov edx, dword ptr [ebp+ffffff20]
'007150c3    83ec10                  sub esp, 10
'007150c6    b804000280              mov eax, 80020004
'007150cb    894108                  mov dword ptr [ecx+08], eax
'007150ce    8b8518ffffff            mov eax, dword ptr [ebp+ffffff18]
'007150d4    89410c                  mov dword ptr [ecx+0c], eax
'007150d7    8bcc                    mov ecx, esp
'007150d9    8bc6                    mov eax, esi
'007150db    8901                    mov dword ptr [ecx], eax
'007150dd    895104                  mov dword ptr [ecx+04], edx
'007150e0    8b9530ffffff            mov edx, dword ptr [ebp+ffffff30]
'007150e6    b804000280              mov eax, 80020004
'007150eb    894108                  mov dword ptr [ecx+08], eax
'007150ee    8b8528ffffff            mov eax, dword ptr [ebp+ffffff28]
'007150f4    89410c                  mov dword ptr [ecx+0c], eax
'007150f7    83ec10                  sub esp, 10
'007150fa    8bcc                    mov ecx, esp
'007150fc    8bc6                    mov eax, esi
'007150fe    8901                    mov dword ptr [ecx], eax
'00715100    895104                  mov dword ptr [ecx+04], edx
'00715103    8b9540ffffff            mov edx, dword ptr [ebp+ffffff40]
'00715109    b804000280              mov eax, 80020004
'0071510e    894108                  mov dword ptr [ecx+08], eax
'00715111    8b8538ffffff            mov eax, dword ptr [ebp+ffffff38]
'00715117    89410c                  mov dword ptr [ecx+0c], eax
'0071511a    83ec10                  sub esp, 10
'0071511d    8bcc                    mov ecx, esp
'0071511f    8bc6                    mov eax, esi
'00715121    8901                    mov dword ptr [ecx], eax
'00715123    895104                  mov dword ptr [ecx+04], edx
'00715126    8b559c                  mov edx, dword ptr [ebp-64]
'00715129    b804000280              mov eax, 80020004
'0071512e    894108                  mov dword ptr [ecx+08], eax
'00715131    8b8548ffffff            mov eax, dword ptr [ebp+ffffff48]
'00715137    89410c                  mov dword ptr [ecx+0c], eax
'0071513a    8b45a0                  mov eax, dword ptr [ebp-60]
'0071513d    83ec10                  sub esp, 10
'00715140    8bcc                    mov ecx, esp
'00715142    8911                    mov dword ptr [ecx], edx
'00715144    8b55a4                  mov edx, dword ptr [ebp-5c]
'00715147    894104                  mov dword ptr [ecx+04], eax
'0071514a    8b45a8                  mov eax, dword ptr [ebp-58]
'0071514d    895108                  mov dword ptr [ecx+08], edx
'00715150    89410c                  mov dword ptr [ecx+0c], eax
'00715153    8b8d14feffff            mov ecx, dword ptr [ebp+fffffe14]
'00715159    68209e4300              push 00439e20
'0071515e    51                      push ecx
'0071515f    ff97f4000000            call dword ptr [edi+000000f4]
'00715165    dbe2                    fnclex
'00715167    85c0                    test ax, ax
'00715169    7d1c                    jge 715187
'0071516b    8b9514feffff            mov edx, dword ptr [ebp+fffffe14]

' *** Reference to "__vbaHresultCheckObj"
'00715171    8b3580104000            mov esi, dword ptr [00401080]
'00715177    68f4000000              push 000000f4
'0071517c    6830314300              push 00433130
'00715181    52                      push edx
'00715182    50                      push eax
'00715183    ffd6                    call esi
'00715185    eb06                    jmp 71518d

' *** Reference to "__vbaHresultCheckObj"
'00715187    8b3580104000            mov esi, dword ptr [00401080]
'0071518d    8d4db8                  lea ecx, dword ptr [ebp-48]

' *** Reference to "__vbaFreeObj"
'00715190    ff1508134000            call dword ptr [00401308]
    '#Cleanup(var_61)
'00715196    8d4d9c                  lea ecx, dword ptr [ebp-64]

' *** Reference to "__vbaFreeVar"
'00715199    ff1528104000            call dword ptr [00401028]
    '#Cleanup(var_51)
'0071519f    8b8514feffff            mov eax, dword ptr [ebp+fffffe14]
'007151a5    8b08                    mov ecx, dword ptr [eax]
'007151a7    8d9538feffff            lea edx, dword ptr [ebp+fffffe38]
'007151ad    52                      push edx
'007151ae    50                      push eax
'007151af    ff515c                  call dword ptr [ecx+5c]
'007151b2    dbe2                    fnclex
'007151b4    85c0                    test ax, ax
'007151b6    7d11                    jge 7151c9
'007151b8    8b8d14feffff            mov ecx, dword ptr [ebp+fffffe14]
'007151be    6a5c                    push 5c
'007151c0    6830314300              push 00433130
'007151c5    51                      push ecx
'007151c6    50                      push eax
'007151c7    ffd6                    call esi
'007151c9    6683bd38feffff00        cmp word ptr [ebp+fffffe38], 00
'007151d1    8b8514feffff            mov eax, dword ptr [ebp+fffffe14]
'007151d7    8b10                    mov edx, dword ptr [eax]
'007151d9    50                      push eax
'007151da    0f8423010000            je 715303
    
    If (    0 <> 0) Then
'007151e0    ff92c0000000            call dword ptr [edx+000000c0]
'007151e6    dbe2                    fnclex
'007151e8    85c0                    test ax, ax
'007151ea    7d14                    jge 715200
'007151ec    8b8d14feffff            mov ecx, dword ptr [ebp+fffffe14]
'007151f2    68c0000000              push 000000c0
'007151f7    6830314300              push 00433130
'007151fc    51                      push ecx
'007151fd    50                      push eax
'007151fe    ffd6                    call esi
'00715200    8b45d0                  mov eax, dword ptr [ebp-30]
'00715203    8b10                    mov edx, dword ptr [eax]
'00715205    8d4db8                  lea ecx, dword ptr [ebp-48]
'00715208    51                      push ecx
'00715209    50                      push eax
'0071520a    ff92b4000000            call dword ptr [edx+000000b4]
'00715210    dbe2                    fnclex
'00715212    85c0                    test ax, ax
'00715214    7d11                    jge 715227
'00715216    8b55d0                  mov edx, dword ptr [ebp-30]
'00715219    68b4000000              push 000000b4
'0071521e    6830314300              push 00433130
'00715223    52                      push edx
'00715224    50                      push eax
'00715225    ffd6                    call esi
'00715227    8b45b8                  mov eax, dword ptr [ebp-48]
'0071522a    8b38                    mov edi, dword ptr [eax]
'0071522c    8d5db4                  lea ebx, dword ptr [ebp-4c]
'0071522f    53                      push ebx
'00715230    83ec10                  sub esp, 10
'00715233    8bdc                    mov ebx, esp
'00715235    ba08000000              mov edx, 00000008
'0071523a    8913                    mov dword ptr [ebx], edx
'0071523c    8b9550ffffff            mov edx, dword ptr [ebp+ffffff50]
'00715242    895304                  mov dword ptr [ebx+04], edx
'00715245    b9c4a54300              mov ecx, 0043a5c4
'0071524a    894b08                  mov dword ptr [ebx+08], ecx
'0071524d    8b8d58ffffff            mov ecx, dword ptr [ebp+ffffff58]
'00715253    50                      push eax
'00715254    8bf0                    mov esi, eax
'00715256    894b0c                  mov dword ptr [ebx+0c], ecx
'00715259    ff5730                  call dword ptr [edi+30]
'0071525c    dbe2                    fnclex
'0071525e    85c0                    test ax, ax
'00715260    7d0f                    jge 715271
'00715262    6a30                    push 30
'00715264    68d8304300              push 004330d8
'00715269    56                      push esi
'0071526a    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0071526b    ff1580104000            call dword ptr [00401080]
'00715271    8b45b4                  mov eax, dword ptr [ebp-4c]
'00715274    8bbd14feffff            mov edi, dword ptr [ebp+fffffe14]
'0071527a    83ec10                  sub esp, 10
'0071527d    8bdc                    mov ebx, esp
'0071527f    b909000000              mov ecx, 00000009
'00715284    890b                    mov dword ptr [ebx], ecx
'00715286    894d9c                  mov dword ptr [ebp-64], ecx
'00715289    8b4da0                  mov ecx, dword ptr [ebp-60]
'0071528c    894b04                  mov dword ptr [ebx+04], ecx
'0071528f    894308                  mov dword ptr [ebx+08], eax
'00715292    8945a4                  mov dword ptr [ebp-5c], eax
'00715295    8b45a8                  mov eax, dword ptr [ebp-58]
'00715298    89430c                  mov dword ptr [ebx+0c], eax
'0071529b    8b8540ffffff            mov eax, dword ptr [ebp+ffffff40]
'007152a1    83ec10                  sub esp, 10
'007152a4    8bcc                    mov ecx, esp
'007152a6    be08000000              mov esi, 00000008
'007152ab    8931                    mov dword ptr [ecx], esi
'007152ad    894104                  mov dword ptr [ecx+04], eax
'007152b0    8b8514feffff            mov eax, dword ptr [ebp+fffffe14]
'007152b6    bac4a54300              mov edx, 0043a5c4
'007152bb    895108                  mov dword ptr [ecx+08], edx
'007152be    8b9548ffffff            mov edx, dword ptr [ebp+ffffff48]
'007152c4    c745b400000000          mov dword ptr [ebp-4c], 00000000
'007152cb    8b3f                    mov edi, dword ptr [edi]
'007152cd    50                      push eax
'007152ce    89510c                  mov dword ptr [ecx+0c], edx
'007152d1    ff9728010000            call dword ptr [edi+00000128]
'007152d7    dbe2                    fnclex
'007152d9    85c0                    test ax, ax
'007152db    7d18                    jge 7152f5
'007152dd    8b8d14feffff            mov ecx, dword ptr [ebp+fffffe14]
'007152e3    6828010000              push 00000128
'007152e8    6830314300              push 00433130
'007152ed    51                      push ecx
'007152ee    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'007152ef    ff1580104000            call dword ptr [00401080]
'007152f5    8d4db8                  lea ecx, dword ptr [ebp-48]

' *** Reference to "__vbaFreeObj"
'007152f8    ff1508134000            call dword ptr [00401308]
    '#Cleanup(var_61)
'007152fe    e992010000              jmp 715495
    
End If
'00715303    ff92d0000000            call dword ptr [edx+000000d0]
'00715309    dbe2                    fnclex
'0071530b    85c0                    test ax, ax
'0071530d    7d14                    jge 715323
'0071530f    8b8d14feffff            mov ecx, dword ptr [ebp+fffffe14]
'00715315    68d0000000              push 000000d0
'0071531a    6830314300              push 00433130
'0071531f    51                      push ecx
'00715320    50                      push eax
'00715321    ffd6                    call esi
'00715323    8b45d0                  mov eax, dword ptr [ebp-30]
'00715326    8b10                    mov edx, dword ptr [eax]
'00715328    8d4db8                  lea ecx, dword ptr [ebp-48]
'0071532b    51                      push ecx
'0071532c    50                      push eax
'0071532d    ff92b4000000            call dword ptr [edx+000000b4]
'00715333    dbe2                    fnclex
'00715335    85c0                    test ax, ax
'00715337    7d11                    jge 71534a
'00715339    8b55d0                  mov edx, dword ptr [ebp-30]
'0071533c    68b4000000              push 000000b4
'00715341    6830314300              push 00433130
'00715346    52                      push edx
'00715347    50                      push eax
'00715348    ffd6                    call esi
'0071534a    8b45b8                  mov eax, dword ptr [ebp-48]
'0071534d    8b38                    mov edi, dword ptr [eax]
'0071534f    8d5db4                  lea ebx, dword ptr [ebp-4c]
'00715352    53                      push ebx
'00715353    83ec10                  sub esp, 10
'00715356    8bdc                    mov ebx, esp
'00715358    ba08000000              mov edx, 00000008
'0071535d    8913                    mov dword ptr [ebx], edx
'0071535f    8b9550ffffff            mov edx, dword ptr [ebp+ffffff50]
'00715365    895304                  mov dword ptr [ebx+04], edx
'00715368    b9c4a54300              mov ecx, 0043a5c4
'0071536d    894b08                  mov dword ptr [ebx+08], ecx
'00715370    8b8d58ffffff            mov ecx, dword ptr [ebp+ffffff58]
'00715376    50                      push eax
'00715377    8bf0                    mov esi, eax
'00715379    894b0c                  mov dword ptr [ebx+0c], ecx
'0071537c    ff5730                  call dword ptr [edi+30]
'0071537f    dbe2                    fnclex
'00715381    85c0                    test ax, ax
'00715383    7d0f                    jge 715394
'00715385    6a30                    push 30
'00715387    68d8304300              push 004330d8
'0071538c    56                      push esi
'0071538d    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0071538e    ff1580104000            call dword ptr [00401080]
'00715394    8b45b4                  mov eax, dword ptr [ebp-4c]
'00715397    8b10                    mov edx, dword ptr [eax]
'00715399    8d4d9c                  lea ecx, dword ptr [ebp-64]
'0071539c    51                      push ecx
'0071539d    50                      push eax
'0071539e    8bf0                    mov esi, eax
'007153a0    ff5244                  call dword ptr [edx+44]
'007153a3    dbe2                    fnclex
'007153a5    85c0                    test ax, ax
'007153a7    7d0f                    jge 7153b8
'007153a9    6a44                    push 44
'007153ab    6880304300              push 00433080
'007153b0    56                      push esi
'007153b1    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'007153b2    ff1580104000            call dword ptr [00401080]
'007153b8    8d559c                  lea edx, dword ptr [ebp-64]
'007153bb    52                      push edx

' *** Reference to "__vbaStrVarMove"
'007153bc    ff153c104000            call dword ptr [0040103c]

' *** Reference to "__vbaStrMove"
'007153c2    8b35d0124000            mov esi, dword ptr [004012d0]
'007153c8    8bd0                    mov edx, eax
'007153ca    8d4dcc                  lea ecx, dword ptr [ebp-34]
'007153cd    ffd6                    call esi
'007153cf    8d45cc                  lea eax, dword ptr [ebp-34]
'007153d2    50                      push eax
'007153d3    e818e8ddff              call 4f3bf0
Call sub_4F3BF0()
'007153d8    8bd0                    mov edx, eax
'007153da    8d4dbc                  lea ecx, dword ptr [ebp-44]
'007153dd    ffd6                    call esi
'007153df    8b3d48b07200            mov edi, dword ptr [0072b048]
'007153e5    8b55bc                  mov edx, dword ptr [ebp-44]
'007153e8    83ec10                  sub esp, 10
'007153eb    c745bc00000000          mov dword ptr [ebp-44], 00000000
'007153f2    8b1f                    mov ebx, dword ptr [edi]
'007153f4    8bfc                    mov edi, esp
'007153f6    b903000000              mov ecx, 00000003
'007153fb    890f                    mov dword ptr [edi], ecx
'007153fd    8b8d40ffffff            mov ecx, dword ptr [ebp+ffffff40]
'00715403    894f04                  mov dword ptr [edi+04], ecx
'00715406    b880000000              mov eax, 00000080
'0071540b    894708                  mov dword ptr [edi+08], eax
'0071540e    8b8548ffffff            mov eax, dword ptr [ebp+ffffff48]
'00715414    68c8b74300              push 0043b7c8
'00715419    8d4dc8                  lea ecx, dword ptr [ebp-38]
'0071541c    89470c                  mov dword ptr [edi+0c], eax
'0071541f    ffd6                    call esi

' *** Reference to "__vbaStrCat"
'00715421    8b3d70104000            mov edi, dword ptr [00401070]
'00715427    50                      push eax
'00715428    ffd7                    call edi
var_43 = ("delete from objetsPropriete where [nomobjet]='") & (-4496)
'0071542a    8bd0                    mov edx, eax
'0071542c    8d4dc4                  lea ecx, dword ptr [ebp-3c]
'0071542f    ffd6                    call esi
'00715431    50                      push eax
'00715432    6854a44300              push 0043a454
'00715437    ffd7                    call edi
var_63 = (var_43) & ("'")
'00715439    8bd0                    mov edx, eax
'0071543b    8d4dc0                  lea ecx, dword ptr [ebp-40]
'0071543e    ffd6                    call esi
'00715440    8b0d48b07200            mov ecx, dword ptr [0072b048]
'00715446    50                      push eax
'00715447    51                      push ecx
'00715448    ff535c                  call dword ptr [ebx+5c]
'0071544b    dbe2                    fnclex
'0071544d    85c0                    test ax, ax
'0071544f    7d15                    jge 715466
'00715451    8b1548b07200            mov edx, dword ptr [0072b048]
'00715457    6a5c                    push 5c
'00715459    68301f4300              push 00431f30
'0071545e    52                      push edx
'0071545f    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00715460    ff1580104000            call dword ptr [00401080]
'00715466    8d45bc                  lea eax, dword ptr [ebp-44]
'00715469    50                      push eax
'0071546a    8d4dc0                  lea ecx, dword ptr [ebp-40]
'0071546d    51                      push ecx
'0071546e    8d55c4                  lea edx, dword ptr [ebp-3c]
'00715471    52                      push edx
'00715472    8d45c8                  lea eax, dword ptr [ebp-38]
'00715475    50                      push eax
'00715476    8d4dcc                  lea ecx, dword ptr [ebp-34]
'00715479    51                      push ecx
'0071547a    6a05                    push 05

' *** Reference to "__vbaFreeStrList"
'0071547c    ff155c124000            call dword ptr [0040125c]
'#Cleanup( -4496, -308, -4500, -4504, var_58)
'00715482    8d55b4                  lea edx, dword ptr [ebp-4c]
'00715485    52                      push edx
'00715486    8d45b8                  lea eax, dword ptr [ebp-48]
'00715489    50                      push eax
'0071548a    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'0071548c    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_61, var_62)
'00715492    83c424                  add esp, 24
'00715495    8d4d9c                  lea ecx, dword ptr [ebp-64]

' *** Reference to "__vbaFreeVar"
'00715498    ff1528104000            call dword ptr [00401028]
'#Cleanup(var_51)
'0071549e    8b45d0                  mov eax, dword ptr [ebp-30]
'007154a1    8b08                    mov ecx, dword ptr [eax]
'007154a3    8d55b8                  lea edx, dword ptr [ebp-48]
'007154a6    52                      push edx
'007154a7    50                      push eax
'007154a8    ff91b4000000            call dword ptr [ecx+000000b4]
'007154ae    dbe2                    fnclex
'007154b0    85c0                    test ax, ax
'007154b2    7d15                    jge 7154c9
'007154b4    8b4dd0                  mov ecx, dword ptr [ebp-30]
'007154b7    68b4000000              push 000000b4
'007154bc    6830314300              push 00433130
'007154c1    51                      push ecx
'007154c2    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'007154c3    ff1580104000            call dword ptr [00401080]
'007154c9    8b45b8                  mov eax, dword ptr [ebp-48]
'007154cc    8b38                    mov edi, dword ptr [eax]
'007154ce    8d5db4                  lea ebx, dword ptr [ebp-4c]
'007154d1    53                      push ebx
'007154d2    83ec10                  sub esp, 10
'007154d5    8bdc                    mov ebx, esp
'007154d7    ba08000000              mov edx, 00000008
'007154dc    8913                    mov dword ptr [ebx], edx
'007154de    8b9550ffffff            mov edx, dword ptr [ebp+ffffff50]
'007154e4    895304                  mov dword ptr [ebx+04], edx
'007154e7    b9f4a54300              mov ecx, 0043a5f4
'007154ec    894b08                  mov dword ptr [ebx+08], ecx
'007154ef    8b8d58ffffff            mov ecx, dword ptr [ebp+ffffff58]
'007154f5    50                      push eax
'007154f6    8bf0                    mov esi, eax
'007154f8    894b0c                  mov dword ptr [ebx+0c], ecx
'007154fb    ff5730                  call dword ptr [edi+30]
'007154fe    dbe2                    fnclex
'00715500    85c0                    test ax, ax
'00715502    7d0f                    jge 715513
'00715504    6a30                    push 30
'00715506    68d8304300              push 004330d8
'0071550b    56                      push esi
'0071550c    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0071550d    ff1580104000            call dword ptr [00401080]
'00715513    8b45b4                  mov eax, dword ptr [ebp-4c]
'00715516    8bbd14feffff            mov edi, dword ptr [ebp+fffffe14]
'0071551c    83ec10                  sub esp, 10
'0071551f    8bdc                    mov ebx, esp
'00715521    b909000000              mov ecx, 00000009
'00715526    890b                    mov dword ptr [ebx], ecx
'00715528    894d9c                  mov dword ptr [ebp-64], ecx
'0071552b    8b4da0                  mov ecx, dword ptr [ebp-60]
'0071552e    894b04                  mov dword ptr [ebx+04], ecx
'00715531    894308                  mov dword ptr [ebx+08], eax
'00715534    8945a4                  mov dword ptr [ebp-5c], eax
'00715537    8b45a8                  mov eax, dword ptr [ebp-58]
'0071553a    89430c                  mov dword ptr [ebx+0c], eax
'0071553d    8b8540ffffff            mov eax, dword ptr [ebp+ffffff40]
'00715543    83ec10                  sub esp, 10
'00715546    8bcc                    mov ecx, esp
'00715548    be08000000              mov esi, 00000008
'0071554d    8931                    mov dword ptr [ecx], esi
'0071554f    894104                  mov dword ptr [ecx+04], eax
'00715552    8b8514feffff            mov eax, dword ptr [ebp+fffffe14]
'00715558    baf4a54300              mov edx, 0043a5f4
'0071555d    895108                  mov dword ptr [ecx+08], edx
'00715560    8b9548ffffff            mov edx, dword ptr [ebp+ffffff48]
'00715566    c745b400000000          mov dword ptr [ebp-4c], 00000000
'0071556d    8b3f                    mov edi, dword ptr [edi]
'0071556f    50                      push eax
'00715570    89510c                  mov dword ptr [ecx+0c], edx
'00715573    ff9728010000            call dword ptr [edi+00000128]
'00715579    dbe2                    fnclex
'0071557b    85c0                    test ax, ax
'0071557d    7d18                    jge 715597
'0071557f    8b8d14feffff            mov ecx, dword ptr [ebp+fffffe14]
'00715585    6828010000              push 00000128
'0071558a    6830314300              push 00433130
'0071558f    51                      push ecx
'00715590    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00715591    ff1580104000            call dword ptr [00401080]
'00715597    8d4db8                  lea ecx, dword ptr [ebp-48]

' *** Reference to "__vbaFreeObj"
'0071559a    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_61)
'007155a0    8d4d9c                  lea ecx, dword ptr [ebp-64]

' *** Reference to "__vbaFreeVar"
'007155a3    ff1528104000            call dword ptr [00401028]
'#Cleanup(var_51)
'007155a9    8b45d0                  mov eax, dword ptr [ebp-30]
'007155ac    8b10                    mov edx, dword ptr [eax]
'007155ae    8d4db8                  lea ecx, dword ptr [ebp-48]
'007155b1    51                      push ecx
'007155b2    50                      push eax
'007155b3    ff92b4000000            call dword ptr [edx+000000b4]
'007155b9    dbe2                    fnclex
'007155bb    85c0                    test ax, ax
'007155bd    7d15                    jge 7155d4
'007155bf    8b55d0                  mov edx, dword ptr [ebp-30]
'007155c2    68b4000000              push 000000b4
'007155c7    6830314300              push 00433130
'007155cc    52                      push edx
'007155cd    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'007155ce    ff1580104000            call dword ptr [00401080]
'007155d4    8b45b8                  mov eax, dword ptr [ebp-48]
'007155d7    8b38                    mov edi, dword ptr [eax]
'007155d9    8d5db4                  lea ebx, dword ptr [ebp-4c]
'007155dc    53                      push ebx
'007155dd    83ec10                  sub esp, 10
'007155e0    8bdc                    mov ebx, esp
'007155e2    ba08000000              mov edx, 00000008
'007155e7    8913                    mov dword ptr [ebx], edx
'007155e9    8b9550ffffff            mov edx, dword ptr [ebp+ffffff50]
'007155ef    895304                  mov dword ptr [ebx+04], edx
'007155f2    b964a74300              mov ecx, 0043a764
'007155f7    894b08                  mov dword ptr [ebx+08], ecx
'007155fa    8b8d58ffffff            mov ecx, dword ptr [ebp+ffffff58]
'00715600    50                      push eax
'00715601    8bf0                    mov esi, eax
'00715603    894b0c                  mov dword ptr [ebx+0c], ecx
'00715606    ff5730                  call dword ptr [edi+30]
'00715609    dbe2                    fnclex
'0071560b    85c0                    test ax, ax
'0071560d    7d0f                    jge 71561e
'0071560f    6a30                    push 30
'00715611    68d8304300              push 004330d8
'00715616    56                      push esi
'00715617    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00715618    ff1580104000            call dword ptr [00401080]
'0071561e    8b45b4                  mov eax, dword ptr [ebp-4c]
'00715621    8bbd14feffff            mov edi, dword ptr [ebp+fffffe14]
'00715627    83ec10                  sub esp, 10
'0071562a    8bdc                    mov ebx, esp
'0071562c    b909000000              mov ecx, 00000009
'00715631    890b                    mov dword ptr [ebx], ecx
'00715633    894d9c                  mov dword ptr [ebp-64], ecx
'00715636    8b4da0                  mov ecx, dword ptr [ebp-60]
'00715639    894b04                  mov dword ptr [ebx+04], ecx
'0071563c    894308                  mov dword ptr [ebx+08], eax
'0071563f    8945a4                  mov dword ptr [ebp-5c], eax
'00715642    8b45a8                  mov eax, dword ptr [ebp-58]
'00715645    89430c                  mov dword ptr [ebx+0c], eax
'00715648    8b8540ffffff            mov eax, dword ptr [ebp+ffffff40]
'0071564e    83ec10                  sub esp, 10
'00715651    8bcc                    mov ecx, esp
'00715653    be08000000              mov esi, 00000008
'00715658    8931                    mov dword ptr [ecx], esi
'0071565a    894104                  mov dword ptr [ecx+04], eax
'0071565d    8b8514feffff            mov eax, dword ptr [ebp+fffffe14]
'00715663    ba64a74300              mov edx, 0043a764
'00715668    895108                  mov dword ptr [ecx+08], edx
'0071566b    8b9548ffffff            mov edx, dword ptr [ebp+ffffff48]
'00715671    c745b400000000          mov dword ptr [ebp-4c], 00000000
'00715678    8b3f                    mov edi, dword ptr [edi]
'0071567a    50                      push eax
'0071567b    89510c                  mov dword ptr [ecx+0c], edx
'0071567e    ff9728010000            call dword ptr [edi+00000128]
'00715684    dbe2                    fnclex
'00715686    85c0                    test ax, ax
'00715688    7d18                    jge 7156a2
'0071568a    8b8d14feffff            mov ecx, dword ptr [ebp+fffffe14]
'00715690    6828010000              push 00000128
'00715695    6830314300              push 00433130
'0071569a    51                      push ecx
'0071569b    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0071569c    ff1580104000            call dword ptr [00401080]
'007156a2    8d4db8                  lea ecx, dword ptr [ebp-48]

' *** Reference to "__vbaFreeObj"
'007156a5    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_61)
'007156ab    8d4d9c                  lea ecx, dword ptr [ebp-64]

' *** Reference to "__vbaFreeVar"
'007156ae    ff1528104000            call dword ptr [00401028]
'#Cleanup(var_51)
'007156b4    8b45d0                  mov eax, dword ptr [ebp-30]
'007156b7    8b10                    mov edx, dword ptr [eax]
'007156b9    8d4db8                  lea ecx, dword ptr [ebp-48]
'007156bc    51                      push ecx
'007156bd    50                      push eax
'007156be    ff92b4000000            call dword ptr [edx+000000b4]
'007156c4    dbe2                    fnclex
'007156c6    85c0                    test ax, ax
'007156c8    7d15                    jge 7156df
'007156ca    8b55d0                  mov edx, dword ptr [ebp-30]
'007156cd    68b4000000              push 000000b4
'007156d2    6830314300              push 00433130
'007156d7    52                      push edx
'007156d8    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'007156d9    ff1580104000            call dword ptr [00401080]
'007156df    8b45b8                  mov eax, dword ptr [ebp-48]
'007156e2    8b38                    mov edi, dword ptr [eax]
'007156e4    8d5db4                  lea ebx, dword ptr [ebp-4c]
'007156e7    53                      push ebx
'007156e8    83ec10                  sub esp, 10
'007156eb    8bdc                    mov ebx, esp
'007156ed    ba08000000              mov edx, 00000008
'007156f2    8913                    mov dword ptr [ebx], edx
'007156f4    8b9550ffffff            mov edx, dword ptr [ebp+ffffff50]
'007156fa    895304                  mov dword ptr [ebx+04], edx
'007156fd    b950a74300              mov ecx, 0043a750
'00715702    894b08                  mov dword ptr [ebx+08], ecx
'00715705    8b8d58ffffff            mov ecx, dword ptr [ebp+ffffff58]
'0071570b    50                      push eax
'0071570c    8bf0                    mov esi, eax
'0071570e    894b0c                  mov dword ptr [ebx+0c], ecx
'00715711    ff5730                  call dword ptr [edi+30]
'00715714    dbe2                    fnclex
'00715716    85c0                    test ax, ax
'00715718    7d0f                    jge 715729
'0071571a    6a30                    push 30
'0071571c    68d8304300              push 004330d8
'00715721    56                      push esi
'00715722    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00715723    ff1580104000            call dword ptr [00401080]
'00715729    8b45b4                  mov eax, dword ptr [ebp-4c]
'0071572c    8bbd14feffff            mov edi, dword ptr [ebp+fffffe14]
'00715732    83ec10                  sub esp, 10
'00715735    8bdc                    mov ebx, esp
'00715737    b909000000              mov ecx, 00000009
'0071573c    890b                    mov dword ptr [ebx], ecx
'0071573e    894d9c                  mov dword ptr [ebp-64], ecx
'00715741    8b4da0                  mov ecx, dword ptr [ebp-60]
'00715744    894b04                  mov dword ptr [ebx+04], ecx
'00715747    894308                  mov dword ptr [ebx+08], eax
'0071574a    8945a4                  mov dword ptr [ebp-5c], eax
'0071574d    8b45a8                  mov eax, dword ptr [ebp-58]
'00715750    89430c                  mov dword ptr [ebx+0c], eax
'00715753    8b8540ffffff            mov eax, dword ptr [ebp+ffffff40]
'00715759    83ec10                  sub esp, 10
'0071575c    8bcc                    mov ecx, esp
'0071575e    be08000000              mov esi, 00000008
'00715763    8931                    mov dword ptr [ecx], esi
'00715765    894104                  mov dword ptr [ecx+04], eax
'00715768    8b8514feffff            mov eax, dword ptr [ebp+fffffe14]
'0071576e    ba50a74300              mov edx, 0043a750
'00715773    895108                  mov dword ptr [ecx+08], edx
'00715776    8b9548ffffff            mov edx, dword ptr [ebp+ffffff48]
'0071577c    c745b400000000          mov dword ptr [ebp-4c], 00000000
'00715783    8b3f                    mov edi, dword ptr [edi]
'00715785    50                      push eax
'00715786    89510c                  mov dword ptr [ecx+0c], edx
'00715789    ff9728010000            call dword ptr [edi+00000128]
'0071578f    dbe2                    fnclex
'00715791    85c0                    test ax, ax
'00715793    7d18                    jge 7157ad
'00715795    8b8d14feffff            mov ecx, dword ptr [ebp+fffffe14]
'0071579b    6828010000              push 00000128
'007157a0    6830314300              push 00433130
'007157a5    51                      push ecx
'007157a6    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'007157a7    ff1580104000            call dword ptr [00401080]
'007157ad    8d4db8                  lea ecx, dword ptr [ebp-48]

' *** Reference to "__vbaFreeObj"
'007157b0    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_61)
'007157b6    8d4d9c                  lea ecx, dword ptr [ebp-64]

' *** Reference to "__vbaFreeVar"
'007157b9    ff1528104000            call dword ptr [00401028]
'#Cleanup(var_51)
'007157bf    8b45d0                  mov eax, dword ptr [ebp-30]
'007157c2    8b10                    mov edx, dword ptr [eax]
'007157c4    8d4db8                  lea ecx, dword ptr [ebp-48]
'007157c7    51                      push ecx
'007157c8    50                      push eax
'007157c9    ff92b4000000            call dword ptr [edx+000000b4]
'007157cf    dbe2                    fnclex
'007157d1    85c0                    test ax, ax
'007157d3    7d15                    jge 7157ea
'007157d5    8b55d0                  mov edx, dword ptr [ebp-30]
'007157d8    68b4000000              push 000000b4
'007157dd    6830314300              push 00433130
'007157e2    52                      push edx
'007157e3    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'007157e4    ff1580104000            call dword ptr [00401080]
'007157ea    8b45b8                  mov eax, dword ptr [ebp-48]
'007157ed    8b38                    mov edi, dword ptr [eax]
'007157ef    8d5db4                  lea ebx, dword ptr [ebp-4c]
'007157f2    53                      push ebx
'007157f3    83ec10                  sub esp, 10
'007157f6    8bdc                    mov ebx, esp
'007157f8    ba08000000              mov edx, 00000008
'007157fd    8913                    mov dword ptr [ebx], edx
'007157ff    8b9550ffffff            mov edx, dword ptr [ebp+ffffff50]
'00715805    895304                  mov dword ptr [ebx+04], edx
'00715808    b90ca64300              mov ecx, 0043a60c
'0071580d    894b08                  mov dword ptr [ebx+08], ecx
'00715810    8b8d58ffffff            mov ecx, dword ptr [ebp+ffffff58]
'00715816    50                      push eax
'00715817    8bf0                    mov esi, eax
'00715819    894b0c                  mov dword ptr [ebx+0c], ecx
'0071581c    ff5730                  call dword ptr [edi+30]
'0071581f    dbe2                    fnclex
'00715821    85c0                    test ax, ax
'00715823    7d0f                    jge 715834
'00715825    6a30                    push 30
'00715827    68d8304300              push 004330d8
'0071582c    56                      push esi
'0071582d    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0071582e    ff1580104000            call dword ptr [00401080]
'00715834    8b45b4                  mov eax, dword ptr [ebp-4c]
'00715837    8bbd14feffff            mov edi, dword ptr [ebp+fffffe14]
'0071583d    83ec10                  sub esp, 10
'00715840    8bdc                    mov ebx, esp
'00715842    b909000000              mov ecx, 00000009
'00715847    890b                    mov dword ptr [ebx], ecx
'00715849    894d9c                  mov dword ptr [ebp-64], ecx
'0071584c    8b4da0                  mov ecx, dword ptr [ebp-60]
'0071584f    894b04                  mov dword ptr [ebx+04], ecx
'00715852    894308                  mov dword ptr [ebx+08], eax
'00715855    8945a4                  mov dword ptr [ebp-5c], eax
'00715858    8b45a8                  mov eax, dword ptr [ebp-58]
'0071585b    89430c                  mov dword ptr [ebx+0c], eax
'0071585e    8b8540ffffff            mov eax, dword ptr [ebp+ffffff40]
'00715864    83ec10                  sub esp, 10
'00715867    8bcc                    mov ecx, esp
'00715869    be08000000              mov esi, 00000008
'0071586e    8931                    mov dword ptr [ecx], esi
'00715870    894104                  mov dword ptr [ecx+04], eax
'00715873    8b8514feffff            mov eax, dword ptr [ebp+fffffe14]
'00715879    ba0ca64300              mov edx, 0043a60c
'0071587e    895108                  mov dword ptr [ecx+08], edx
'00715881    8b9548ffffff            mov edx, dword ptr [ebp+ffffff48]
'00715887    c745b400000000          mov dword ptr [ebp-4c], 00000000
'0071588e    8b3f                    mov edi, dword ptr [edi]
'00715890    50                      push eax
'00715891    89510c                  mov dword ptr [ecx+0c], edx
'00715894    ff9728010000            call dword ptr [edi+00000128]
'0071589a    dbe2                    fnclex
'0071589c    85c0                    test ax, ax
'0071589e    7d18                    jge 7158b8
'007158a0    8b8d14feffff            mov ecx, dword ptr [ebp+fffffe14]
'007158a6    6828010000              push 00000128
'007158ab    6830314300              push 00433130
'007158b0    51                      push ecx
'007158b1    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'007158b2    ff1580104000            call dword ptr [00401080]
'007158b8    8d4db8                  lea ecx, dword ptr [ebp-48]

' *** Reference to "__vbaFreeObj"
'007158bb    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_61)
'007158c1    8d4d9c                  lea ecx, dword ptr [ebp-64]

' *** Reference to "__vbaFreeVar"
'007158c4    ff1528104000            call dword ptr [00401028]
'#Cleanup(var_51)
'007158ca    8b45d0                  mov eax, dword ptr [ebp-30]
'007158cd    8b10                    mov edx, dword ptr [eax]
'007158cf    8d4db8                  lea ecx, dword ptr [ebp-48]
'007158d2    51                      push ecx
'007158d3    50                      push eax
'007158d4    ff92b4000000            call dword ptr [edx+000000b4]
'007158da    dbe2                    fnclex
'007158dc    85c0                    test ax, ax
'007158de    7d15                    jge 7158f5
'007158e0    8b55d0                  mov edx, dword ptr [ebp-30]
'007158e3    68b4000000              push 000000b4
'007158e8    6830314300              push 00433130
'007158ed    52                      push edx
'007158ee    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'007158ef    ff1580104000            call dword ptr [00401080]
'007158f5    8b45b8                  mov eax, dword ptr [ebp-48]
'007158f8    8b38                    mov edi, dword ptr [eax]
'007158fa    8d5db4                  lea ebx, dword ptr [ebp-4c]
'007158fd    53                      push ebx
'007158fe    83ec10                  sub esp, 10
'00715901    8bdc                    mov ebx, esp
'00715903    ba08000000              mov edx, 00000008
'00715908    8913                    mov dword ptr [ebx], edx
'0071590a    8b9550ffffff            mov edx, dword ptr [ebp+ffffff50]
'00715910    895304                  mov dword ptr [ebx+04], edx
'00715913    b934a64300              mov ecx, 0043a634
'00715918    894b08                  mov dword ptr [ebx+08], ecx
'0071591b    8b8d58ffffff            mov ecx, dword ptr [ebp+ffffff58]
'00715921    50                      push eax
'00715922    8bf0                    mov esi, eax
'00715924    894b0c                  mov dword ptr [ebx+0c], ecx
'00715927    ff5730                  call dword ptr [edi+30]
'0071592a    dbe2                    fnclex
'0071592c    85c0                    test ax, ax
'0071592e    7d0f                    jge 71593f
'00715930    6a30                    push 30
'00715932    68d8304300              push 004330d8
'00715937    56                      push esi
'00715938    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00715939    ff1580104000            call dword ptr [00401080]
'0071593f    8b45b4                  mov eax, dword ptr [ebp-4c]
'00715942    8bbd14feffff            mov edi, dword ptr [ebp+fffffe14]
'00715948    83ec10                  sub esp, 10
'0071594b    8bdc                    mov ebx, esp
'0071594d    b909000000              mov ecx, 00000009
'00715952    890b                    mov dword ptr [ebx], ecx
'00715954    894d9c                  mov dword ptr [ebp-64], ecx
'00715957    8b4da0                  mov ecx, dword ptr [ebp-60]
'0071595a    894b04                  mov dword ptr [ebx+04], ecx
'0071595d    894308                  mov dword ptr [ebx+08], eax
'00715960    8945a4                  mov dword ptr [ebp-5c], eax
'00715963    8b45a8                  mov eax, dword ptr [ebp-58]
'00715966    89430c                  mov dword ptr [ebx+0c], eax
'00715969    8b8540ffffff            mov eax, dword ptr [ebp+ffffff40]
'0071596f    83ec10                  sub esp, 10
'00715972    8bcc                    mov ecx, esp
'00715974    be08000000              mov esi, 00000008
'00715979    8931                    mov dword ptr [ecx], esi
'0071597b    894104                  mov dword ptr [ecx+04], eax
'0071597e    8b8514feffff            mov eax, dword ptr [ebp+fffffe14]
'00715984    ba34a64300              mov edx, 0043a634
'00715989    895108                  mov dword ptr [ecx+08], edx
'0071598c    8b9548ffffff            mov edx, dword ptr [ebp+ffffff48]
'00715992    c745b400000000          mov dword ptr [ebp-4c], 00000000
'00715999    8b3f                    mov edi, dword ptr [edi]
'0071599b    50                      push eax
'0071599c    89510c                  mov dword ptr [ecx+0c], edx
'0071599f    ff9728010000            call dword ptr [edi+00000128]
'007159a5    dbe2                    fnclex
'007159a7    85c0                    test ax, ax
'007159a9    7d18                    jge 7159c3
'007159ab    8b8d14feffff            mov ecx, dword ptr [ebp+fffffe14]
'007159b1    6828010000              push 00000128
'007159b6    6830314300              push 00433130
'007159bb    51                      push ecx
'007159bc    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'007159bd    ff1580104000            call dword ptr [00401080]
'007159c3    8d4db8                  lea ecx, dword ptr [ebp-48]

' *** Reference to "__vbaFreeObj"
'007159c6    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_61)
'007159cc    8d4d9c                  lea ecx, dword ptr [ebp-64]

' *** Reference to "__vbaFreeVar"
'007159cf    ff1528104000            call dword ptr [00401028]
'#Cleanup(var_51)
'007159d5    8b45d0                  mov eax, dword ptr [ebp-30]
'007159d8    8b10                    mov edx, dword ptr [eax]
'007159da    8d4db8                  lea ecx, dword ptr [ebp-48]
'007159dd    51                      push ecx
'007159de    50                      push eax
'007159df    ff92b4000000            call dword ptr [edx+000000b4]
'007159e5    dbe2                    fnclex
'007159e7    85c0                    test ax, ax
'007159e9    7d15                    jge 715a00
'007159eb    8b55d0                  mov edx, dword ptr [ebp-30]
'007159ee    68b4000000              push 000000b4
'007159f3    6830314300              push 00433130
'007159f8    52                      push edx
'007159f9    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'007159fa    ff1580104000            call dword ptr [00401080]
'00715a00    8b45b8                  mov eax, dword ptr [ebp-48]
'00715a03    8b38                    mov edi, dword ptr [eax]
'00715a05    8d5db4                  lea ebx, dword ptr [ebp-4c]
'00715a08    53                      push ebx
'00715a09    83ec10                  sub esp, 10
'00715a0c    8bdc                    mov ebx, esp
'00715a0e    ba08000000              mov edx, 00000008
'00715a13    8913                    mov dword ptr [ebx], edx
'00715a15    8b9550ffffff            mov edx, dword ptr [ebp+ffffff50]
'00715a1b    895304                  mov dword ptr [ebx+04], edx
'00715a1e    b9a8a24300              mov ecx, 0043a2a8
'00715a23    894b08                  mov dword ptr [ebx+08], ecx
'00715a26    8b8d58ffffff            mov ecx, dword ptr [ebp+ffffff58]
'00715a2c    50                      push eax
'00715a2d    8bf0                    mov esi, eax
'00715a2f    894b0c                  mov dword ptr [ebx+0c], ecx
'00715a32    ff5730                  call dword ptr [edi+30]
'00715a35    dbe2                    fnclex
'00715a37    85c0                    test ax, ax
'00715a39    7d0f                    jge 715a4a
'00715a3b    6a30                    push 30
'00715a3d    68d8304300              push 004330d8
'00715a42    56                      push esi
'00715a43    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00715a44    ff1580104000            call dword ptr [00401080]
'00715a4a    8b45b4                  mov eax, dword ptr [ebp-4c]
'00715a4d    8bbd14feffff            mov edi, dword ptr [ebp+fffffe14]
'00715a53    83ec10                  sub esp, 10
'00715a56    8bdc                    mov ebx, esp
'00715a58    b909000000              mov ecx, 00000009
'00715a5d    890b                    mov dword ptr [ebx], ecx
'00715a5f    894d9c                  mov dword ptr [ebp-64], ecx
'00715a62    8b4da0                  mov ecx, dword ptr [ebp-60]
'00715a65    894b04                  mov dword ptr [ebx+04], ecx
'00715a68    894308                  mov dword ptr [ebx+08], eax
'00715a6b    8945a4                  mov dword ptr [ebp-5c], eax
'00715a6e    8b45a8                  mov eax, dword ptr [ebp-58]
'00715a71    89430c                  mov dword ptr [ebx+0c], eax
'00715a74    8b8540ffffff            mov eax, dword ptr [ebp+ffffff40]
'00715a7a    83ec10                  sub esp, 10
'00715a7d    8bcc                    mov ecx, esp
'00715a7f    be08000000              mov esi, 00000008
'00715a84    8931                    mov dword ptr [ecx], esi
'00715a86    894104                  mov dword ptr [ecx+04], eax
'00715a89    8b8514feffff            mov eax, dword ptr [ebp+fffffe14]
'00715a8f    baa8a24300              mov edx, 0043a2a8
'00715a94    895108                  mov dword ptr [ecx+08], edx
'00715a97    8b9548ffffff            mov edx, dword ptr [ebp+ffffff48]
'00715a9d    c745b400000000          mov dword ptr [ebp-4c], 00000000
'00715aa4    8b3f                    mov edi, dword ptr [edi]
'00715aa6    50                      push eax
'00715aa7    89510c                  mov dword ptr [ecx+0c], edx
'00715aaa    ff9728010000            call dword ptr [edi+00000128]
'00715ab0    dbe2                    fnclex
'00715ab2    85c0                    test ax, ax
'00715ab4    7d18                    jge 715ace
'00715ab6    8b8d14feffff            mov ecx, dword ptr [ebp+fffffe14]
'00715abc    6828010000              push 00000128
'00715ac1    6830314300              push 00433130
'00715ac6    51                      push ecx
'00715ac7    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00715ac8    ff1580104000            call dword ptr [00401080]
'00715ace    8d4db8                  lea ecx, dword ptr [ebp-48]

' *** Reference to "__vbaFreeObj"
'00715ad1    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_61)
'00715ad7    8d4d9c                  lea ecx, dword ptr [ebp-64]

' *** Reference to "__vbaFreeVar"
'00715ada    ff1528104000            call dword ptr [00401028]
'#Cleanup(var_51)
'00715ae0    8b45d0                  mov eax, dword ptr [ebp-30]
'00715ae3    8b10                    mov edx, dword ptr [eax]
'00715ae5    8d4db8                  lea ecx, dword ptr [ebp-48]
'00715ae8    51                      push ecx
'00715ae9    50                      push eax
'00715aea    ff92b4000000            call dword ptr [edx+000000b4]
'00715af0    dbe2                    fnclex
'00715af2    85c0                    test ax, ax
'00715af4    7d15                    jge 715b0b
'00715af6    8b55d0                  mov edx, dword ptr [ebp-30]
'00715af9    68b4000000              push 000000b4
'00715afe    6830314300              push 00433130
'00715b03    52                      push edx
'00715b04    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00715b05    ff1580104000            call dword ptr [00401080]
'00715b0b    8b45b8                  mov eax, dword ptr [ebp-48]
'00715b0e    8b38                    mov edi, dword ptr [eax]
'00715b10    8d5db4                  lea ebx, dword ptr [ebp-4c]
'00715b13    53                      push ebx
'00715b14    83ec10                  sub esp, 10
'00715b17    8bdc                    mov ebx, esp
'00715b19    ba08000000              mov edx, 00000008
'00715b1e    8913                    mov dword ptr [ebx], edx
'00715b20    8b9550ffffff            mov edx, dword ptr [ebp+ffffff50]
'00715b26    895304                  mov dword ptr [ebx+04], edx
'00715b29    b920a74300              mov ecx, 0043a720
'00715b2e    894b08                  mov dword ptr [ebx+08], ecx
'00715b31    8b8d58ffffff            mov ecx, dword ptr [ebp+ffffff58]
'00715b37    50                      push eax
'00715b38    8bf0                    mov esi, eax
'00715b3a    894b0c                  mov dword ptr [ebx+0c], ecx
'00715b3d    ff5730                  call dword ptr [edi+30]
'00715b40    dbe2                    fnclex
'00715b42    85c0                    test ax, ax
'00715b44    7d0f                    jge 715b55
'00715b46    6a30                    push 30
'00715b48    68d8304300              push 004330d8
'00715b4d    56                      push esi
'00715b4e    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00715b4f    ff1580104000            call dword ptr [00401080]
'00715b55    8b45b4                  mov eax, dword ptr [ebp-4c]
'00715b58    8bbd14feffff            mov edi, dword ptr [ebp+fffffe14]
'00715b5e    83ec10                  sub esp, 10
'00715b61    8bdc                    mov ebx, esp
'00715b63    b909000000              mov ecx, 00000009
'00715b68    890b                    mov dword ptr [ebx], ecx
'00715b6a    894d9c                  mov dword ptr [ebp-64], ecx
'00715b6d    8b4da0                  mov ecx, dword ptr [ebp-60]
'00715b70    894b04                  mov dword ptr [ebx+04], ecx
'00715b73    894308                  mov dword ptr [ebx+08], eax
'00715b76    8945a4                  mov dword ptr [ebp-5c], eax
'00715b79    8b45a8                  mov eax, dword ptr [ebp-58]
'00715b7c    89430c                  mov dword ptr [ebx+0c], eax
'00715b7f    8b8540ffffff            mov eax, dword ptr [ebp+ffffff40]
'00715b85    83ec10                  sub esp, 10
'00715b88    8bcc                    mov ecx, esp
'00715b8a    be08000000              mov esi, 00000008
'00715b8f    8931                    mov dword ptr [ecx], esi
'00715b91    894104                  mov dword ptr [ecx+04], eax
'00715b94    8b8514feffff            mov eax, dword ptr [ebp+fffffe14]
'00715b9a    ba20a74300              mov edx, 0043a720
'00715b9f    895108                  mov dword ptr [ecx+08], edx
'00715ba2    8b9548ffffff            mov edx, dword ptr [ebp+ffffff48]
'00715ba8    c745b400000000          mov dword ptr [ebp-4c], 00000000
'00715baf    8b3f                    mov edi, dword ptr [edi]
'00715bb1    50                      push eax
'00715bb2    89510c                  mov dword ptr [ecx+0c], edx
'00715bb5    ff9728010000            call dword ptr [edi+00000128]
'00715bbb    dbe2                    fnclex
'00715bbd    85c0                    test ax, ax
'00715bbf    7d18                    jge 715bd9
'00715bc1    8b8d14feffff            mov ecx, dword ptr [ebp+fffffe14]
'00715bc7    6828010000              push 00000128
'00715bcc    6830314300              push 00433130
'00715bd1    51                      push ecx
'00715bd2    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00715bd3    ff1580104000            call dword ptr [00401080]
'00715bd9    8d4db8                  lea ecx, dword ptr [ebp-48]

' *** Reference to "__vbaFreeObj"
'00715bdc    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_61)
'00715be2    8d4d9c                  lea ecx, dword ptr [ebp-64]

' *** Reference to "__vbaFreeVar"
'00715be5    ff1528104000            call dword ptr [00401028]
'#Cleanup(var_51)
'00715beb    8b45d0                  mov eax, dword ptr [ebp-30]
'00715bee    8b10                    mov edx, dword ptr [eax]
'00715bf0    8d4db8                  lea ecx, dword ptr [ebp-48]
'00715bf3    51                      push ecx
'00715bf4    50                      push eax
'00715bf5    ff92b4000000            call dword ptr [edx+000000b4]
'00715bfb    dbe2                    fnclex
'00715bfd    85c0                    test ax, ax
'00715bff    7d15                    jge 715c16
'00715c01    8b55d0                  mov edx, dword ptr [ebp-30]
'00715c04    68b4000000              push 000000b4
'00715c09    6830314300              push 00433130
'00715c0e    52                      push edx
'00715c0f    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00715c10    ff1580104000            call dword ptr [00401080]
'00715c16    8b45b8                  mov eax, dword ptr [ebp-48]
'00715c19    8b38                    mov edi, dword ptr [eax]
'00715c1b    8d5db4                  lea ebx, dword ptr [ebp-4c]
'00715c1e    53                      push ebx
'00715c1f    83ec10                  sub esp, 10
'00715c22    8bdc                    mov ebx, esp
'00715c24    ba08000000              mov edx, 00000008
'00715c29    8913                    mov dword ptr [ebx], edx
'00715c2b    8b9550ffffff            mov edx, dword ptr [ebp+ffffff50]
'00715c31    895304                  mov dword ptr [ebx+04], edx
'00715c34    b93ca74300              mov ecx, 0043a73c
'00715c39    894b08                  mov dword ptr [ebx+08], ecx
'00715c3c    8b8d58ffffff            mov ecx, dword ptr [ebp+ffffff58]
'00715c42    50                      push eax
'00715c43    8bf0                    mov esi, eax
'00715c45    894b0c                  mov dword ptr [ebx+0c], ecx
'00715c48    ff5730                  call dword ptr [edi+30]
'00715c4b    dbe2                    fnclex
'00715c4d    85c0                    test ax, ax
'00715c4f    7d0f                    jge 715c60
'00715c51    6a30                    push 30
'00715c53    68d8304300              push 004330d8
'00715c58    56                      push esi
'00715c59    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00715c5a    ff1580104000            call dword ptr [00401080]
'00715c60    8b45b4                  mov eax, dword ptr [ebp-4c]
'00715c63    8bbd14feffff            mov edi, dword ptr [ebp+fffffe14]
'00715c69    83ec10                  sub esp, 10
'00715c6c    8bdc                    mov ebx, esp
'00715c6e    b909000000              mov ecx, 00000009
'00715c73    890b                    mov dword ptr [ebx], ecx
'00715c75    894d9c                  mov dword ptr [ebp-64], ecx
'00715c78    8b4da0                  mov ecx, dword ptr [ebp-60]
'00715c7b    894b04                  mov dword ptr [ebx+04], ecx
'00715c7e    894308                  mov dword ptr [ebx+08], eax
'00715c81    8945a4                  mov dword ptr [ebp-5c], eax
'00715c84    8b45a8                  mov eax, dword ptr [ebp-58]
'00715c87    89430c                  mov dword ptr [ebx+0c], eax
'00715c8a    8b8540ffffff            mov eax, dword ptr [ebp+ffffff40]
'00715c90    83ec10                  sub esp, 10
'00715c93    8bcc                    mov ecx, esp
'00715c95    be08000000              mov esi, 00000008
'00715c9a    8931                    mov dword ptr [ecx], esi
'00715c9c    894104                  mov dword ptr [ecx+04], eax
'00715c9f    8b8514feffff            mov eax, dword ptr [ebp+fffffe14]
'00715ca5    ba3ca74300              mov edx, 0043a73c
'00715caa    895108                  mov dword ptr [ecx+08], edx
'00715cad    8b9548ffffff            mov edx, dword ptr [ebp+ffffff48]
'00715cb3    c745b400000000          mov dword ptr [ebp-4c], 00000000
'00715cba    8b3f                    mov edi, dword ptr [edi]
'00715cbc    50                      push eax
'00715cbd    89510c                  mov dword ptr [ecx+0c], edx
'00715cc0    ff9728010000            call dword ptr [edi+00000128]
'00715cc6    dbe2                    fnclex
'00715cc8    85c0                    test ax, ax
'00715cca    7d18                    jge 715ce4
'00715ccc    8b8d14feffff            mov ecx, dword ptr [ebp+fffffe14]
'00715cd2    6828010000              push 00000128
'00715cd7    6830314300              push 00433130
'00715cdc    51                      push ecx
'00715cdd    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00715cde    ff1580104000            call dword ptr [00401080]
'00715ce4    8d4db8                  lea ecx, dword ptr [ebp-48]

' *** Reference to "__vbaFreeObj"
'00715ce7    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_61)
'00715ced    8d4d9c                  lea ecx, dword ptr [ebp-64]

' *** Reference to "__vbaFreeVar"
'00715cf0    ff1528104000            call dword ptr [00401028]
'#Cleanup(var_51)
'00715cf6    8b45d0                  mov eax, dword ptr [ebp-30]
'00715cf9    8b10                    mov edx, dword ptr [eax]
'00715cfb    8d4db8                  lea ecx, dword ptr [ebp-48]
'00715cfe    51                      push ecx
'00715cff    50                      push eax
'00715d00    ff92b4000000            call dword ptr [edx+000000b4]
'00715d06    dbe2                    fnclex
'00715d08    85c0                    test ax, ax
'00715d0a    7d15                    jge 715d21
'00715d0c    8b55d0                  mov edx, dword ptr [ebp-30]
'00715d0f    68b4000000              push 000000b4
'00715d14    6830314300              push 00433130
'00715d19    52                      push edx
'00715d1a    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00715d1b    ff1580104000            call dword ptr [00401080]
'00715d21    8b45b8                  mov eax, dword ptr [ebp-48]
'00715d24    8b38                    mov edi, dword ptr [eax]
'00715d26    8d5db4                  lea ebx, dword ptr [ebp-4c]
'00715d29    53                      push ebx
'00715d2a    83ec10                  sub esp, 10
'00715d2d    8bdc                    mov ebx, esp
'00715d2f    ba08000000              mov edx, 00000008
'00715d34    8913                    mov dword ptr [ebx], edx
'00715d36    8b9550ffffff            mov edx, dword ptr [ebp+ffffff50]
'00715d3c    895304                  mov dword ptr [ebx+04], edx
'00715d3f    b974a64300              mov ecx, 0043a674
'00715d44    894b08                  mov dword ptr [ebx+08], ecx
'00715d47    8b8d58ffffff            mov ecx, dword ptr [ebp+ffffff58]
'00715d4d    50                      push eax
'00715d4e    8bf0                    mov esi, eax
'00715d50    894b0c                  mov dword ptr [ebx+0c], ecx
'00715d53    ff5730                  call dword ptr [edi+30]
'00715d56    dbe2                    fnclex
'00715d58    85c0                    test ax, ax
'00715d5a    7d0f                    jge 715d6b
'00715d5c    6a30                    push 30
'00715d5e    68d8304300              push 004330d8
'00715d63    56                      push esi
'00715d64    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00715d65    ff1580104000            call dword ptr [00401080]
'00715d6b    8b45b4                  mov eax, dword ptr [ebp-4c]
'00715d6e    8bbd14feffff            mov edi, dword ptr [ebp+fffffe14]
'00715d74    83ec10                  sub esp, 10
'00715d77    8bdc                    mov ebx, esp
'00715d79    b909000000              mov ecx, 00000009
'00715d7e    890b                    mov dword ptr [ebx], ecx
'00715d80    894d9c                  mov dword ptr [ebp-64], ecx
'00715d83    8b4da0                  mov ecx, dword ptr [ebp-60]
'00715d86    894b04                  mov dword ptr [ebx+04], ecx
'00715d89    894308                  mov dword ptr [ebx+08], eax
'00715d8c    8945a4                  mov dword ptr [ebp-5c], eax
'00715d8f    8b45a8                  mov eax, dword ptr [ebp-58]
'00715d92    89430c                  mov dword ptr [ebx+0c], eax
'00715d95    8b8540ffffff            mov eax, dword ptr [ebp+ffffff40]
'00715d9b    83ec10                  sub esp, 10
'00715d9e    8bcc                    mov ecx, esp
'00715da0    be08000000              mov esi, 00000008
'00715da5    8931                    mov dword ptr [ecx], esi
'00715da7    894104                  mov dword ptr [ecx+04], eax
'00715daa    8b8514feffff            mov eax, dword ptr [ebp+fffffe14]
'00715db0    ba74a64300              mov edx, 0043a674
'00715db5    895108                  mov dword ptr [ecx+08], edx
'00715db8    8b9548ffffff            mov edx, dword ptr [ebp+ffffff48]
'00715dbe    c745b400000000          mov dword ptr [ebp-4c], 00000000
'00715dc5    8b3f                    mov edi, dword ptr [edi]
'00715dc7    50                      push eax
'00715dc8    89510c                  mov dword ptr [ecx+0c], edx
'00715dcb    ff9728010000            call dword ptr [edi+00000128]
'00715dd1    dbe2                    fnclex
'00715dd3    85c0                    test ax, ax
'00715dd5    7d18                    jge 715def
'00715dd7    8b8d14feffff            mov ecx, dword ptr [ebp+fffffe14]
'00715ddd    6828010000              push 00000128
'00715de2    6830314300              push 00433130
'00715de7    51                      push ecx
'00715de8    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00715de9    ff1580104000            call dword ptr [00401080]
'00715def    8d4db8                  lea ecx, dword ptr [ebp-48]

' *** Reference to "__vbaFreeObj"
'00715df2    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_61)
'00715df8    8d4d9c                  lea ecx, dword ptr [ebp-64]

' *** Reference to "__vbaFreeVar"
'00715dfb    ff1528104000            call dword ptr [00401028]
'#Cleanup(var_51)
'00715e01    8b45d0                  mov eax, dword ptr [ebp-30]
'00715e04    8b10                    mov edx, dword ptr [eax]
'00715e06    8d4db8                  lea ecx, dword ptr [ebp-48]
'00715e09    51                      push ecx
'00715e0a    50                      push eax
'00715e0b    ff92b4000000            call dword ptr [edx+000000b4]
'00715e11    dbe2                    fnclex
'00715e13    85c0                    test ax, ax
'00715e15    7d15                    jge 715e2c
'00715e17    8b55d0                  mov edx, dword ptr [ebp-30]
'00715e1a    68b4000000              push 000000b4
'00715e1f    6830314300              push 00433130
'00715e24    52                      push edx
'00715e25    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00715e26    ff1580104000            call dword ptr [00401080]
'00715e2c    8b45b8                  mov eax, dword ptr [ebp-48]
'00715e2f    8b38                    mov edi, dword ptr [eax]
'00715e31    8d5db4                  lea ebx, dword ptr [ebp-4c]
'00715e34    53                      push ebx
'00715e35    83ec10                  sub esp, 10
'00715e38    8bdc                    mov ebx, esp
'00715e3a    ba08000000              mov edx, 00000008
'00715e3f    8913                    mov dword ptr [ebx], edx
'00715e41    8b9550ffffff            mov edx, dword ptr [ebp+ffffff50]
'00715e47    895304                  mov dword ptr [ebx+04], edx
'00715e4a    b95ca64300              mov ecx, 0043a65c
'00715e4f    894b08                  mov dword ptr [ebx+08], ecx
'00715e52    8b8d58ffffff            mov ecx, dword ptr [ebp+ffffff58]
'00715e58    50                      push eax
'00715e59    8bf0                    mov esi, eax
'00715e5b    894b0c                  mov dword ptr [ebx+0c], ecx
'00715e5e    ff5730                  call dword ptr [edi+30]
'00715e61    dbe2                    fnclex
'00715e63    85c0                    test ax, ax
'00715e65    7d0f                    jge 715e76
'00715e67    6a30                    push 30
'00715e69    68d8304300              push 004330d8
'00715e6e    56                      push esi
'00715e6f    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00715e70    ff1580104000            call dword ptr [00401080]
'00715e76    8b45b4                  mov eax, dword ptr [ebp-4c]
'00715e79    8bbd14feffff            mov edi, dword ptr [ebp+fffffe14]
'00715e7f    83ec10                  sub esp, 10
'00715e82    8bdc                    mov ebx, esp
'00715e84    b909000000              mov ecx, 00000009
'00715e89    890b                    mov dword ptr [ebx], ecx
'00715e8b    894d9c                  mov dword ptr [ebp-64], ecx
'00715e8e    8b4da0                  mov ecx, dword ptr [ebp-60]
'00715e91    894b04                  mov dword ptr [ebx+04], ecx
'00715e94    894308                  mov dword ptr [ebx+08], eax
'00715e97    8945a4                  mov dword ptr [ebp-5c], eax
'00715e9a    8b45a8                  mov eax, dword ptr [ebp-58]
'00715e9d    89430c                  mov dword ptr [ebx+0c], eax
'00715ea0    8b8540ffffff            mov eax, dword ptr [ebp+ffffff40]
'00715ea6    83ec10                  sub esp, 10
'00715ea9    8bcc                    mov ecx, esp
'00715eab    be08000000              mov esi, 00000008
'00715eb0    8931                    mov dword ptr [ecx], esi
'00715eb2    894104                  mov dword ptr [ecx+04], eax
'00715eb5    8b8514feffff            mov eax, dword ptr [ebp+fffffe14]
'00715ebb    ba5ca64300              mov edx, 0043a65c
'00715ec0    895108                  mov dword ptr [ecx+08], edx
'00715ec3    8b9548ffffff            mov edx, dword ptr [ebp+ffffff48]
'00715ec9    c745b400000000          mov dword ptr [ebp-4c], 00000000
'00715ed0    8b3f                    mov edi, dword ptr [edi]
'00715ed2    50                      push eax
'00715ed3    89510c                  mov dword ptr [ecx+0c], edx
'00715ed6    ff9728010000            call dword ptr [edi+00000128]
'00715edc    dbe2                    fnclex
'00715ede    85c0                    test ax, ax
'00715ee0    7d18                    jge 715efa
'00715ee2    8b8d14feffff            mov ecx, dword ptr [ebp+fffffe14]
'00715ee8    6828010000              push 00000128
'00715eed    6830314300              push 00433130
'00715ef2    51                      push ecx
'00715ef3    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00715ef4    ff1580104000            call dword ptr [00401080]
'00715efa    8d4db8                  lea ecx, dword ptr [ebp-48]

' *** Reference to "__vbaFreeObj"
'00715efd    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_61)
'00715f03    8d4d9c                  lea ecx, dword ptr [ebp-64]

' *** Reference to "__vbaFreeVar"
'00715f06    ff1528104000            call dword ptr [00401028]
'#Cleanup(var_51)
'00715f0c    8b45d0                  mov eax, dword ptr [ebp-30]
'00715f0f    8b10                    mov edx, dword ptr [eax]
'00715f11    8d4db8                  lea ecx, dword ptr [ebp-48]
'00715f14    51                      push ecx
'00715f15    50                      push eax
'00715f16    ff92b4000000            call dword ptr [edx+000000b4]
'00715f1c    dbe2                    fnclex
'00715f1e    85c0                    test ax, ax
'00715f20    7d15                    jge 715f37
'00715f22    8b55d0                  mov edx, dword ptr [ebp-30]
'00715f25    68b4000000              push 000000b4
'00715f2a    6830314300              push 00433130
'00715f2f    52                      push edx
'00715f30    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00715f31    ff1580104000            call dword ptr [00401080]
'00715f37    8b45b8                  mov eax, dword ptr [ebp-48]
'00715f3a    8b38                    mov edi, dword ptr [eax]
'00715f3c    8d5db4                  lea ebx, dword ptr [ebp-4c]
'00715f3f    53                      push ebx
'00715f40    83ec10                  sub esp, 10
'00715f43    8bdc                    mov ebx, esp
'00715f45    ba08000000              mov edx, 00000008
'00715f4a    8913                    mov dword ptr [ebx], edx
'00715f4c    8b9550ffffff            mov edx, dword ptr [ebp+ffffff50]
'00715f52    895304                  mov dword ptr [ebx+04], edx
'00715f55    b9e8a64300              mov ecx, 0043a6e8
'00715f5a    894b08                  mov dword ptr [ebx+08], ecx
'00715f5d    8b8d58ffffff            mov ecx, dword ptr [ebp+ffffff58]
'00715f63    50                      push eax
'00715f64    8bf0                    mov esi, eax
'00715f66    894b0c                  mov dword ptr [ebx+0c], ecx
'00715f69    ff5730                  call dword ptr [edi+30]
'00715f6c    dbe2                    fnclex
'00715f6e    85c0                    test ax, ax
'00715f70    7d0f                    jge 715f81
'00715f72    6a30                    push 30
'00715f74    68d8304300              push 004330d8
'00715f79    56                      push esi
'00715f7a    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00715f7b    ff1580104000            call dword ptr [00401080]
'00715f81    8b45b4                  mov eax, dword ptr [ebp-4c]
'00715f84    8bbd14feffff            mov edi, dword ptr [ebp+fffffe14]
'00715f8a    83ec10                  sub esp, 10
'00715f8d    8bdc                    mov ebx, esp
'00715f8f    b909000000              mov ecx, 00000009
'00715f94    890b                    mov dword ptr [ebx], ecx
'00715f96    894d9c                  mov dword ptr [ebp-64], ecx
'00715f99    8b4da0                  mov ecx, dword ptr [ebp-60]
'00715f9c    894b04                  mov dword ptr [ebx+04], ecx
'00715f9f    894308                  mov dword ptr [ebx+08], eax
'00715fa2    8945a4                  mov dword ptr [ebp-5c], eax
'00715fa5    8b45a8                  mov eax, dword ptr [ebp-58]
'00715fa8    89430c                  mov dword ptr [ebx+0c], eax
'00715fab    8b8540ffffff            mov eax, dword ptr [ebp+ffffff40]
'00715fb1    83ec10                  sub esp, 10
'00715fb4    8bcc                    mov ecx, esp
'00715fb6    be08000000              mov esi, 00000008
'00715fbb    8931                    mov dword ptr [ecx], esi
'00715fbd    894104                  mov dword ptr [ecx+04], eax
'00715fc0    8b8514feffff            mov eax, dword ptr [ebp+fffffe14]
'00715fc6    bae8a64300              mov edx, 0043a6e8
'00715fcb    895108                  mov dword ptr [ecx+08], edx
'00715fce    8b9548ffffff            mov edx, dword ptr [ebp+ffffff48]
'00715fd4    c745b400000000          mov dword ptr [ebp-4c], 00000000
'00715fdb    8b3f                    mov edi, dword ptr [edi]
'00715fdd    50                      push eax
'00715fde    89510c                  mov dword ptr [ecx+0c], edx
'00715fe1    ff9728010000            call dword ptr [edi+00000128]
'00715fe7    dbe2                    fnclex
'00715fe9    85c0                    test ax, ax
'00715feb    7d18                    jge 716005
'00715fed    8b8d14feffff            mov ecx, dword ptr [ebp+fffffe14]
'00715ff3    6828010000              push 00000128
'00715ff8    6830314300              push 00433130
'00715ffd    51                      push ecx
'00715ffe    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00715fff    ff1580104000            call dword ptr [00401080]
'00716005    8d4db8                  lea ecx, dword ptr [ebp-48]

' *** Reference to "__vbaFreeObj"
'00716008    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_61)
'0071600e    8d4d9c                  lea ecx, dword ptr [ebp-64]

' *** Reference to "__vbaFreeVar"
'00716011    ff1528104000            call dword ptr [00401028]
'#Cleanup(var_51)
'00716017    8b45d0                  mov eax, dword ptr [ebp-30]
'0071601a    8b10                    mov edx, dword ptr [eax]
'0071601c    8d4db8                  lea ecx, dword ptr [ebp-48]
'0071601f    51                      push ecx
'00716020    50                      push eax
'00716021    ff92b4000000            call dword ptr [edx+000000b4]
'00716027    dbe2                    fnclex
'00716029    85c0                    test ax, ax
'0071602b    7d15                    jge 716042
'0071602d    8b55d0                  mov edx, dword ptr [ebp-30]
'00716030    68b4000000              push 000000b4
'00716035    6830314300              push 00433130
'0071603a    52                      push edx
'0071603b    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0071603c    ff1580104000            call dword ptr [00401080]
'00716042    8b45b8                  mov eax, dword ptr [ebp-48]
'00716045    8b38                    mov edi, dword ptr [eax]
'00716047    8d5db4                  lea ebx, dword ptr [ebp-4c]
'0071604a    53                      push ebx
'0071604b    83ec10                  sub esp, 10
'0071604e    8bdc                    mov ebx, esp
'00716050    ba08000000              mov edx, 00000008
'00716055    8913                    mov dword ptr [ebx], edx
'00716057    8b9550ffffff            mov edx, dword ptr [ebp+ffffff50]
'0071605d    895304                  mov dword ptr [ebx+04], edx
'00716060    b908a74300              mov ecx, 0043a708
'00716065    894b08                  mov dword ptr [ebx+08], ecx
'00716068    8b8d58ffffff            mov ecx, dword ptr [ebp+ffffff58]
'0071606e    50                      push eax
'0071606f    8bf0                    mov esi, eax
'00716071    894b0c                  mov dword ptr [ebx+0c], ecx
'00716074    ff5730                  call dword ptr [edi+30]
'00716077    dbe2                    fnclex
'00716079    85c0                    test ax, ax
'0071607b    7d0f                    jge 71608c
'0071607d    6a30                    push 30
'0071607f    68d8304300              push 004330d8
'00716084    56                      push esi
'00716085    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00716086    ff1580104000            call dword ptr [00401080]
'0071608c    8b45b4                  mov eax, dword ptr [ebp-4c]
'0071608f    8bbd14feffff            mov edi, dword ptr [ebp+fffffe14]
'00716095    83ec10                  sub esp, 10
'00716098    8bdc                    mov ebx, esp
'0071609a    b909000000              mov ecx, 00000009
'0071609f    890b                    mov dword ptr [ebx], ecx
'007160a1    894d9c                  mov dword ptr [ebp-64], ecx
'007160a4    8b4da0                  mov ecx, dword ptr [ebp-60]
'007160a7    894b04                  mov dword ptr [ebx+04], ecx
'007160aa    894308                  mov dword ptr [ebx+08], eax
'007160ad    8945a4                  mov dword ptr [ebp-5c], eax
'007160b0    8b45a8                  mov eax, dword ptr [ebp-58]
'007160b3    89430c                  mov dword ptr [ebx+0c], eax
'007160b6    8b8540ffffff            mov eax, dword ptr [ebp+ffffff40]
'007160bc    83ec10                  sub esp, 10
'007160bf    8bcc                    mov ecx, esp
'007160c1    be08000000              mov esi, 00000008
'007160c6    8931                    mov dword ptr [ecx], esi
'007160c8    894104                  mov dword ptr [ecx+04], eax
'007160cb    8b8514feffff            mov eax, dword ptr [ebp+fffffe14]
'007160d1    ba08a74300              mov edx, 0043a708
'007160d6    895108                  mov dword ptr [ecx+08], edx
'007160d9    8b9548ffffff            mov edx, dword ptr [ebp+ffffff48]
'007160df    c745b400000000          mov dword ptr [ebp-4c], 00000000
'007160e6    8b3f                    mov edi, dword ptr [edi]
'007160e8    50                      push eax
'007160e9    89510c                  mov dword ptr [ecx+0c], edx
'007160ec    ff9728010000            call dword ptr [edi+00000128]
'007160f2    dbe2                    fnclex
'007160f4    85c0                    test ax, ax
'007160f6    7d18                    jge 716110
'007160f8    8b8d14feffff            mov ecx, dword ptr [ebp+fffffe14]
'007160fe    6828010000              push 00000128
'00716103    6830314300              push 00433130
'00716108    51                      push ecx
'00716109    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0071610a    ff1580104000            call dword ptr [00401080]
'00716110    8d4db8                  lea ecx, dword ptr [ebp-48]

' *** Reference to "__vbaFreeObj"
'00716113    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_61)
'00716119    8d4d9c                  lea ecx, dword ptr [ebp-64]

' *** Reference to "__vbaFreeVar"
'0071611c    ff1528104000            call dword ptr [00401028]
'#Cleanup(var_51)
'00716122    8b45d0                  mov eax, dword ptr [ebp-30]
'00716125    8b10                    mov edx, dword ptr [eax]
'00716127    8d4db8                  lea ecx, dword ptr [ebp-48]
'0071612a    51                      push ecx
'0071612b    50                      push eax
'0071612c    ff92b4000000            call dword ptr [edx+000000b4]
'00716132    dbe2                    fnclex
'00716134    85c0                    test ax, ax
'00716136    7d15                    jge 71614d
'00716138    8b55d0                  mov edx, dword ptr [ebp-30]
'0071613b    68b4000000              push 000000b4
'00716140    6830314300              push 00433130
'00716145    52                      push edx
'00716146    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00716147    ff1580104000            call dword ptr [00401080]
'0071614d    8b45b8                  mov eax, dword ptr [ebp-48]
'00716150    8b38                    mov edi, dword ptr [eax]
'00716152    8d5db4                  lea ebx, dword ptr [ebp-4c]
'00716155    53                      push ebx
'00716156    83ec10                  sub esp, 10
'00716159    8bdc                    mov ebx, esp
'0071615b    ba08000000              mov edx, 00000008
'00716160    8913                    mov dword ptr [ebx], edx
'00716162    8b9550ffffff            mov edx, dword ptr [ebp+ffffff50]
'00716168    895304                  mov dword ptr [ebx+04], edx
'0071616b    b984a64300              mov ecx, 0043a684
'00716170    894b08                  mov dword ptr [ebx+08], ecx
'00716173    8b8d58ffffff            mov ecx, dword ptr [ebp+ffffff58]
'00716179    50                      push eax
'0071617a    8bf0                    mov esi, eax
'0071617c    894b0c                  mov dword ptr [ebx+0c], ecx
'0071617f    ff5730                  call dword ptr [edi+30]
'00716182    dbe2                    fnclex
'00716184    85c0                    test ax, ax
'00716186    7d0f                    jge 716197
'00716188    6a30                    push 30
'0071618a    68d8304300              push 004330d8
'0071618f    56                      push esi
'00716190    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00716191    ff1580104000            call dword ptr [00401080]
'00716197    8b45b4                  mov eax, dword ptr [ebp-4c]
'0071619a    8bbd14feffff            mov edi, dword ptr [ebp+fffffe14]
'007161a0    83ec10                  sub esp, 10
'007161a3    8bdc                    mov ebx, esp
'007161a5    b909000000              mov ecx, 00000009
'007161aa    890b                    mov dword ptr [ebx], ecx
'007161ac    894d9c                  mov dword ptr [ebp-64], ecx
'007161af    8b4da0                  mov ecx, dword ptr [ebp-60]
'007161b2    894b04                  mov dword ptr [ebx+04], ecx
'007161b5    894308                  mov dword ptr [ebx+08], eax
'007161b8    8945a4                  mov dword ptr [ebp-5c], eax
'007161bb    8b45a8                  mov eax, dword ptr [ebp-58]
'007161be    89430c                  mov dword ptr [ebx+0c], eax
'007161c1    8b8540ffffff            mov eax, dword ptr [ebp+ffffff40]
'007161c7    83ec10                  sub esp, 10
'007161ca    8bcc                    mov ecx, esp
'007161cc    be08000000              mov esi, 00000008
'007161d1    8931                    mov dword ptr [ecx], esi
'007161d3    894104                  mov dword ptr [ecx+04], eax
'007161d6    8b8514feffff            mov eax, dword ptr [ebp+fffffe14]
'007161dc    ba84a64300              mov edx, 0043a684
'007161e1    895108                  mov dword ptr [ecx+08], edx
'007161e4    8b9548ffffff            mov edx, dword ptr [ebp+ffffff48]
'007161ea    c745b400000000          mov dword ptr [ebp-4c], 00000000
'007161f1    8b3f                    mov edi, dword ptr [edi]
'007161f3    50                      push eax
'007161f4    89510c                  mov dword ptr [ecx+0c], edx
'007161f7    ff9728010000            call dword ptr [edi+00000128]
'007161fd    dbe2                    fnclex
'007161ff    85c0                    test ax, ax
'00716201    7d18                    jge 71621b
'00716203    8b8d14feffff            mov ecx, dword ptr [ebp+fffffe14]
'00716209    6828010000              push 00000128
'0071620e    6830314300              push 00433130
'00716213    51                      push ecx
'00716214    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00716215    ff1580104000            call dword ptr [00401080]
'0071621b    8d4db8                  lea ecx, dword ptr [ebp-48]

' *** Reference to "__vbaFreeObj"
'0071621e    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_61)
'00716224    8d4d9c                  lea ecx, dword ptr [ebp-64]

' *** Reference to "__vbaFreeVar"
'00716227    ff1528104000            call dword ptr [00401028]
'#Cleanup(var_51)
'0071622d    8b45d0                  mov eax, dword ptr [ebp-30]
'00716230    8b10                    mov edx, dword ptr [eax]
'00716232    8d4db8                  lea ecx, dword ptr [ebp-48]
'00716235    51                      push ecx
'00716236    50                      push eax
'00716237    ff92b4000000            call dword ptr [edx+000000b4]
'0071623d    dbe2                    fnclex
'0071623f    85c0                    test ax, ax
'00716241    7d15                    jge 716258
'00716243    8b55d0                  mov edx, dword ptr [ebp-30]
'00716246    68b4000000              push 000000b4
'0071624b    6830314300              push 00433130
'00716250    52                      push edx
'00716251    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00716252    ff1580104000            call dword ptr [00401080]
'00716258    8b45b8                  mov eax, dword ptr [ebp-48]
'0071625b    8b38                    mov edi, dword ptr [eax]
'0071625d    8d5db4                  lea ebx, dword ptr [ebp-4c]
'00716260    53                      push ebx
'00716261    83ec10                  sub esp, 10
'00716264    8bdc                    mov ebx, esp
'00716266    ba08000000              mov edx, 00000008
'0071626b    8913                    mov dword ptr [ebx], edx
'0071626d    8b9550ffffff            mov edx, dword ptr [ebp+ffffff50]
'00716273    895304                  mov dword ptr [ebx+04], edx
'00716276    b9dca54300              mov ecx, 0043a5dc
'0071627b    894b08                  mov dword ptr [ebx+08], ecx
'0071627e    8b8d58ffffff            mov ecx, dword ptr [ebp+ffffff58]
'00716284    50                      push eax
'00716285    8bf0                    mov esi, eax
'00716287    894b0c                  mov dword ptr [ebx+0c], ecx
'0071628a    ff5730                  call dword ptr [edi+30]
'0071628d    dbe2                    fnclex
'0071628f    85c0                    test ax, ax
'00716291    7d0f                    jge 7162a2
'00716293    6a30                    push 30
'00716295    68d8304300              push 004330d8
'0071629a    56                      push esi
'0071629b    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0071629c    ff1580104000            call dword ptr [00401080]
'007162a2    8b45b4                  mov eax, dword ptr [ebp-4c]
'007162a5    8bbd14feffff            mov edi, dword ptr [ebp+fffffe14]
'007162ab    83ec10                  sub esp, 10
'007162ae    8bdc                    mov ebx, esp
'007162b0    b909000000              mov ecx, 00000009
'007162b5    890b                    mov dword ptr [ebx], ecx
'007162b7    894d9c                  mov dword ptr [ebp-64], ecx
'007162ba    8b4da0                  mov ecx, dword ptr [ebp-60]
'007162bd    894b04                  mov dword ptr [ebx+04], ecx
'007162c0    894308                  mov dword ptr [ebx+08], eax
'007162c3    8945a4                  mov dword ptr [ebp-5c], eax
'007162c6    8b45a8                  mov eax, dword ptr [ebp-58]
'007162c9    89430c                  mov dword ptr [ebx+0c], eax
'007162cc    8b8540ffffff            mov eax, dword ptr [ebp+ffffff40]
'007162d2    83ec10                  sub esp, 10
'007162d5    8bcc                    mov ecx, esp
'007162d7    be08000000              mov esi, 00000008
'007162dc    8931                    mov dword ptr [ecx], esi
'007162de    894104                  mov dword ptr [ecx+04], eax
'007162e1    8b8514feffff            mov eax, dword ptr [ebp+fffffe14]
'007162e7    badca54300              mov edx, 0043a5dc
'007162ec    895108                  mov dword ptr [ecx+08], edx
'007162ef    8b9548ffffff            mov edx, dword ptr [ebp+ffffff48]
'007162f5    c745b400000000          mov dword ptr [ebp-4c], 00000000
'007162fc    8b3f                    mov edi, dword ptr [edi]
'007162fe    50                      push eax
'007162ff    89510c                  mov dword ptr [ecx+0c], edx
'00716302    ff9728010000            call dword ptr [edi+00000128]
'00716308    dbe2                    fnclex
'0071630a    85c0                    test ax, ax
'0071630c    7d18                    jge 716326
'0071630e    8b8d14feffff            mov ecx, dword ptr [ebp+fffffe14]
'00716314    6828010000              push 00000128
'00716319    6830314300              push 00433130
'0071631e    51                      push ecx
'0071631f    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00716320    ff1580104000            call dword ptr [00401080]
'00716326    8d4db8                  lea ecx, dword ptr [ebp-48]

' *** Reference to "__vbaFreeObj"
'00716329    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_61)
'0071632f    8d4d9c                  lea ecx, dword ptr [ebp-64]

' *** Reference to "__vbaFreeVar"
'00716332    ff1528104000            call dword ptr [00401028]
'#Cleanup(var_51)
'00716338    8b45d0                  mov eax, dword ptr [ebp-30]
'0071633b    8b10                    mov edx, dword ptr [eax]
'0071633d    8d4db8                  lea ecx, dword ptr [ebp-48]
'00716340    51                      push ecx
'00716341    50                      push eax
'00716342    ff92b4000000            call dword ptr [edx+000000b4]
'00716348    dbe2                    fnclex
'0071634a    85c0                    test ax, ax
'0071634c    7d15                    jge 716363
'0071634e    8b55d0                  mov edx, dword ptr [ebp-30]
'00716351    68b4000000              push 000000b4
'00716356    6830314300              push 00433130
'0071635b    52                      push edx
'0071635c    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0071635d    ff1580104000            call dword ptr [00401080]
'00716363    8b45b8                  mov eax, dword ptr [ebp-48]
'00716366    8b38                    mov edi, dword ptr [eax]
'00716368    8d5db4                  lea ebx, dword ptr [ebp-4c]
'0071636b    53                      push ebx
'0071636c    83ec10                  sub esp, 10
'0071636f    8bdc                    mov ebx, esp
'00716371    ba08000000              mov edx, 00000008
'00716376    8913                    mov dword ptr [ebx], edx
'00716378    8b9550ffffff            mov edx, dword ptr [ebp+ffffff50]
'0071637e    895304                  mov dword ptr [ebx+04], edx
'00716381    b9cca64300              mov ecx, 0043a6cc
'00716386    894b08                  mov dword ptr [ebx+08], ecx
'00716389    8b8d58ffffff            mov ecx, dword ptr [ebp+ffffff58]
'0071638f    50                      push eax
'00716390    8bf0                    mov esi, eax
'00716392    894b0c                  mov dword ptr [ebx+0c], ecx
'00716395    ff5730                  call dword ptr [edi+30]
'00716398    dbe2                    fnclex
'0071639a    85c0                    test ax, ax
'0071639c    7d0f                    jge 7163ad
'0071639e    6a30                    push 30
'007163a0    68d8304300              push 004330d8
'007163a5    56                      push esi
'007163a6    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'007163a7    ff1580104000            call dword ptr [00401080]
'007163ad    8b45b4                  mov eax, dword ptr [ebp-4c]
'007163b0    8bbd14feffff            mov edi, dword ptr [ebp+fffffe14]
'007163b6    83ec10                  sub esp, 10
'007163b9    8bdc                    mov ebx, esp
'007163bb    b909000000              mov ecx, 00000009
'007163c0    890b                    mov dword ptr [ebx], ecx
'007163c2    894d9c                  mov dword ptr [ebp-64], ecx
'007163c5    8b4da0                  mov ecx, dword ptr [ebp-60]
'007163c8    894b04                  mov dword ptr [ebx+04], ecx
'007163cb    894308                  mov dword ptr [ebx+08], eax
'007163ce    8945a4                  mov dword ptr [ebp-5c], eax
'007163d1    8b45a8                  mov eax, dword ptr [ebp-58]
'007163d4    89430c                  mov dword ptr [ebx+0c], eax
'007163d7    8b8540ffffff            mov eax, dword ptr [ebp+ffffff40]
'007163dd    83ec10                  sub esp, 10
'007163e0    8bcc                    mov ecx, esp
'007163e2    be08000000              mov esi, 00000008
'007163e7    8931                    mov dword ptr [ecx], esi
'007163e9    894104                  mov dword ptr [ecx+04], eax
'007163ec    8b8514feffff            mov eax, dword ptr [ebp+fffffe14]
'007163f2    bacca64300              mov edx, 0043a6cc
'007163f7    895108                  mov dword ptr [ecx+08], edx
'007163fa    8b9548ffffff            mov edx, dword ptr [ebp+ffffff48]
'00716400    c745b400000000          mov dword ptr [ebp-4c], 00000000
'00716407    8b3f                    mov edi, dword ptr [edi]
'00716409    50                      push eax
'0071640a    89510c                  mov dword ptr [ecx+0c], edx
'0071640d    ff9728010000            call dword ptr [edi+00000128]
'00716413    dbe2                    fnclex
'00716415    85c0                    test ax, ax
'00716417    7d18                    jge 716431
'00716419    8b8d14feffff            mov ecx, dword ptr [ebp+fffffe14]
'0071641f    6828010000              push 00000128
'00716424    6830314300              push 00433130
'00716429    51                      push ecx
'0071642a    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0071642b    ff1580104000            call dword ptr [00401080]
'00716431    8d4db8                  lea ecx, dword ptr [ebp-48]

' *** Reference to "__vbaFreeObj"
'00716434    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_61)
'0071643a    8d4d9c                  lea ecx, dword ptr [ebp-64]

' *** Reference to "__vbaFreeVar"
'0071643d    ff1528104000            call dword ptr [00401028]
'#Cleanup(var_51)
'00716443    8b45d0                  mov eax, dword ptr [ebp-30]
'00716446    8b10                    mov edx, dword ptr [eax]
'00716448    8d4db8                  lea ecx, dword ptr [ebp-48]
'0071644b    51                      push ecx
'0071644c    50                      push eax
'0071644d    ff92b4000000            call dword ptr [edx+000000b4]
'00716453    dbe2                    fnclex
'00716455    85c0                    test ax, ax
'00716457    7d15                    jge 71646e
'00716459    8b55d0                  mov edx, dword ptr [ebp-30]
'0071645c    68b4000000              push 000000b4
'00716461    6830314300              push 00433130
'00716466    52                      push edx
'00716467    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00716468    ff1580104000            call dword ptr [00401080]
'0071646e    8b45b8                  mov eax, dword ptr [ebp-48]
'00716471    8b38                    mov edi, dword ptr [eax]
'00716473    8d5db4                  lea ebx, dword ptr [ebp-4c]
'00716476    53                      push ebx
'00716477    83ec10                  sub esp, 10
'0071647a    8bdc                    mov ebx, esp
'0071647c    ba08000000              mov edx, 00000008
'00716481    8913                    mov dword ptr [ebx], edx
'00716483    8b9550ffffff            mov edx, dword ptr [ebp+ffffff50]
'00716489    895304                  mov dword ptr [ebx+04], edx
'0071648c    b998a64300              mov ecx, 0043a698
'00716491    894b08                  mov dword ptr [ebx+08], ecx
'00716494    8b8d58ffffff            mov ecx, dword ptr [ebp+ffffff58]
'0071649a    50                      push eax
'0071649b    8bf0                    mov esi, eax
'0071649d    894b0c                  mov dword ptr [ebx+0c], ecx
'007164a0    ff5730                  call dword ptr [edi+30]
'007164a3    dbe2                    fnclex
'007164a5    85c0                    test ax, ax
'007164a7    7d0f                    jge 7164b8
'007164a9    6a30                    push 30
'007164ab    68d8304300              push 004330d8
'007164b0    56                      push esi
'007164b1    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'007164b2    ff1580104000            call dword ptr [00401080]
'007164b8    8b45b4                  mov eax, dword ptr [ebp-4c]
'007164bb    8bbd14feffff            mov edi, dword ptr [ebp+fffffe14]
'007164c1    83ec10                  sub esp, 10
'007164c4    8bdc                    mov ebx, esp
'007164c6    b909000000              mov ecx, 00000009
'007164cb    890b                    mov dword ptr [ebx], ecx
'007164cd    894d9c                  mov dword ptr [ebp-64], ecx
'007164d0    8b4da0                  mov ecx, dword ptr [ebp-60]
'007164d3    894b04                  mov dword ptr [ebx+04], ecx
'007164d6    894308                  mov dword ptr [ebx+08], eax
'007164d9    8945a4                  mov dword ptr [ebp-5c], eax
'007164dc    8b45a8                  mov eax, dword ptr [ebp-58]
'007164df    89430c                  mov dword ptr [ebx+0c], eax
'007164e2    8b8540ffffff            mov eax, dword ptr [ebp+ffffff40]
'007164e8    83ec10                  sub esp, 10
'007164eb    8bcc                    mov ecx, esp
'007164ed    be08000000              mov esi, 00000008
'007164f2    8931                    mov dword ptr [ecx], esi
'007164f4    894104                  mov dword ptr [ecx+04], eax
'007164f7    8b8514feffff            mov eax, dword ptr [ebp+fffffe14]
'007164fd    ba98a64300              mov edx, 0043a698
'00716502    895108                  mov dword ptr [ecx+08], edx
'00716505    8b9548ffffff            mov edx, dword ptr [ebp+ffffff48]
'0071650b    c745b400000000          mov dword ptr [ebp-4c], 00000000
'00716512    8b3f                    mov edi, dword ptr [edi]
'00716514    50                      push eax
'00716515    89510c                  mov dword ptr [ecx+0c], edx
'00716518    ff9728010000            call dword ptr [edi+00000128]
'0071651e    dbe2                    fnclex
'00716520    85c0                    test ax, ax
'00716522    7d18                    jge 71653c
'00716524    8b8d14feffff            mov ecx, dword ptr [ebp+fffffe14]
'0071652a    6828010000              push 00000128
'0071652f    6830314300              push 00433130
'00716534    51                      push ecx
'00716535    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00716536    ff1580104000            call dword ptr [00401080]
'0071653c    8d4db8                  lea ecx, dword ptr [ebp-48]

' *** Reference to "__vbaFreeObj"
'0071653f    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_61)
'00716545    8d4d9c                  lea ecx, dword ptr [ebp-64]

' *** Reference to "__vbaFreeVar"
'00716548    ff1528104000            call dword ptr [00401028]
'#Cleanup(var_51)
'0071654e    8b45d0                  mov eax, dword ptr [ebp-30]
'00716551    8b10                    mov edx, dword ptr [eax]
'00716553    8d4db8                  lea ecx, dword ptr [ebp-48]
'00716556    51                      push ecx
'00716557    50                      push eax
'00716558    ff92b4000000            call dword ptr [edx+000000b4]
'0071655e    dbe2                    fnclex
'00716560    85c0                    test ax, ax
'00716562    7d15                    jge 716579
'00716564    8b55d0                  mov edx, dword ptr [ebp-30]
'00716567    68b4000000              push 000000b4
'0071656c    6830314300              push 00433130
'00716571    52                      push edx
'00716572    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00716573    ff1580104000            call dword ptr [00401080]
'00716579    8b45b8                  mov eax, dword ptr [ebp-48]
'0071657c    8b38                    mov edi, dword ptr [eax]
'0071657e    8d5db4                  lea ebx, dword ptr [ebp-4c]
'00716581    53                      push ebx
'00716582    83ec10                  sub esp, 10
'00716585    8bdc                    mov ebx, esp
'00716587    ba08000000              mov edx, 00000008
'0071658c    8913                    mov dword ptr [ebx], edx
'0071658e    8b9550ffffff            mov edx, dword ptr [ebp+ffffff50]
'00716594    895304                  mov dword ptr [ebx+04], edx
'00716597    b9b0a64300              mov ecx, 0043a6b0
'0071659c    894b08                  mov dword ptr [ebx+08], ecx
'0071659f    8b8d58ffffff            mov ecx, dword ptr [ebp+ffffff58]
'007165a5    50                      push eax
'007165a6    8bf0                    mov esi, eax
'007165a8    894b0c                  mov dword ptr [ebx+0c], ecx
'007165ab    ff5730                  call dword ptr [edi+30]
'007165ae    dbe2                    fnclex
'007165b0    85c0                    test ax, ax
'007165b2    7d0f                    jge 7165c3
'007165b4    6a30                    push 30
'007165b6    68d8304300              push 004330d8
'007165bb    56                      push esi
'007165bc    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'007165bd    ff1580104000            call dword ptr [00401080]
'007165c3    8b45b4                  mov eax, dword ptr [ebp-4c]
'007165c6    8bbd14feffff            mov edi, dword ptr [ebp+fffffe14]
'007165cc    83ec10                  sub esp, 10
'007165cf    8bdc                    mov ebx, esp
'007165d1    b909000000              mov ecx, 00000009
'007165d6    890b                    mov dword ptr [ebx], ecx
'007165d8    894d9c                  mov dword ptr [ebp-64], ecx
'007165db    8b4da0                  mov ecx, dword ptr [ebp-60]
'007165de    894b04                  mov dword ptr [ebx+04], ecx
'007165e1    894308                  mov dword ptr [ebx+08], eax
'007165e4    8945a4                  mov dword ptr [ebp-5c], eax
'007165e7    8b45a8                  mov eax, dword ptr [ebp-58]
'007165ea    89430c                  mov dword ptr [ebx+0c], eax
'007165ed    8b8540ffffff            mov eax, dword ptr [ebp+ffffff40]
'007165f3    83ec10                  sub esp, 10
'007165f6    8bcc                    mov ecx, esp
'007165f8    be08000000              mov esi, 00000008
'007165fd    8931                    mov dword ptr [ecx], esi
'007165ff    894104                  mov dword ptr [ecx+04], eax
'00716602    8b8514feffff            mov eax, dword ptr [ebp+fffffe14]
'00716608    bab0a64300              mov edx, 0043a6b0
'0071660d    895108                  mov dword ptr [ecx+08], edx
'00716610    8b9548ffffff            mov edx, dword ptr [ebp+ffffff48]
'00716616    c745b400000000          mov dword ptr [ebp-4c], 00000000
'0071661d    8b3f                    mov edi, dword ptr [edi]
'0071661f    50                      push eax
'00716620    89510c                  mov dword ptr [ecx+0c], edx
'00716623    ff9728010000            call dword ptr [edi+00000128]
'00716629    dbe2                    fnclex
'0071662b    85c0                    test ax, ax
'0071662d    7d18                    jge 716647
'0071662f    8b8d14feffff            mov ecx, dword ptr [ebp+fffffe14]
'00716635    6828010000              push 00000128
'0071663a    6830314300              push 00433130
'0071663f    51                      push ecx
'00716640    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00716641    ff1580104000            call dword ptr [00401080]
'00716647    8d4db8                  lea ecx, dword ptr [ebp-48]

' *** Reference to "__vbaFreeObj"
'0071664a    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_61)
'00716650    8d4d9c                  lea ecx, dword ptr [ebp-64]

' *** Reference to "__vbaFreeVar"
'00716653    ff1528104000            call dword ptr [00401028]
'#Cleanup(var_51)
'00716659    8b45d0                  mov eax, dword ptr [ebp-30]
'0071665c    8b10                    mov edx, dword ptr [eax]
'0071665e    8d4db8                  lea ecx, dword ptr [ebp-48]
'00716661    51                      push ecx
'00716662    50                      push eax
'00716663    ff92b4000000            call dword ptr [edx+000000b4]
'00716669    dbe2                    fnclex
'0071666b    85c0                    test ax, ax
'0071666d    7d15                    jge 716684
'0071666f    8b55d0                  mov edx, dword ptr [ebp-30]
'00716672    68b4000000              push 000000b4
'00716677    6830314300              push 00433130
'0071667c    52                      push edx
'0071667d    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0071667e    ff1580104000            call dword ptr [00401080]
'00716684    8b45b8                  mov eax, dword ptr [ebp-48]
'00716687    8b38                    mov edi, dword ptr [eax]
'00716689    8d5db4                  lea ebx, dword ptr [ebp-4c]
'0071668c    53                      push ebx
'0071668d    83ec10                  sub esp, 10
'00716690    8bdc                    mov ebx, esp
'00716692    ba08000000              mov edx, 00000008
'00716697    8913                    mov dword ptr [ebx], edx
'00716699    8b9550ffffff            mov edx, dword ptr [ebp+ffffff50]
'0071669f    895304                  mov dword ptr [ebx+04], edx
'007166a2    b94ca84300              mov ecx, 0043a84c
'007166a7    894b08                  mov dword ptr [ebx+08], ecx
'007166aa    8b8d58ffffff            mov ecx, dword ptr [ebp+ffffff58]
'007166b0    50                      push eax
'007166b1    8bf0                    mov esi, eax
'007166b3    894b0c                  mov dword ptr [ebx+0c], ecx
'007166b6    ff5730                  call dword ptr [edi+30]
'007166b9    dbe2                    fnclex
'007166bb    85c0                    test ax, ax
'007166bd    7d0f                    jge 7166ce
'007166bf    6a30                    push 30
'007166c1    68d8304300              push 004330d8
'007166c6    56                      push esi
'007166c7    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'007166c8    ff1580104000            call dword ptr [00401080]
'007166ce    8b45b4                  mov eax, dword ptr [ebp-4c]
'007166d1    8bbd14feffff            mov edi, dword ptr [ebp+fffffe14]
'007166d7    83ec10                  sub esp, 10
'007166da    8bdc                    mov ebx, esp
'007166dc    b909000000              mov ecx, 00000009
'007166e1    890b                    mov dword ptr [ebx], ecx
'007166e3    894d9c                  mov dword ptr [ebp-64], ecx
'007166e6    8b4da0                  mov ecx, dword ptr [ebp-60]
'007166e9    894b04                  mov dword ptr [ebx+04], ecx
'007166ec    894308                  mov dword ptr [ebx+08], eax
'007166ef    8945a4                  mov dword ptr [ebp-5c], eax
'007166f2    8b45a8                  mov eax, dword ptr [ebp-58]
'007166f5    89430c                  mov dword ptr [ebx+0c], eax
'007166f8    8b8540ffffff            mov eax, dword ptr [ebp+ffffff40]
'007166fe    83ec10                  sub esp, 10
'00716701    8bcc                    mov ecx, esp
'00716703    be08000000              mov esi, 00000008
'00716708    8931                    mov dword ptr [ecx], esi
'0071670a    894104                  mov dword ptr [ecx+04], eax
'0071670d    8b8514feffff            mov eax, dword ptr [ebp+fffffe14]
'00716713    ba4ca84300              mov edx, 0043a84c
'00716718    895108                  mov dword ptr [ecx+08], edx
'0071671b    8b9548ffffff            mov edx, dword ptr [ebp+ffffff48]
'00716721    c745b400000000          mov dword ptr [ebp-4c], 00000000
'00716728    8b3f                    mov edi, dword ptr [edi]
'0071672a    50                      push eax
'0071672b    89510c                  mov dword ptr [ecx+0c], edx
'0071672e    ff9728010000            call dword ptr [edi+00000128]
'00716734    dbe2                    fnclex
'00716736    85c0                    test ax, ax
'00716738    7d1c                    jge 716756
'0071673a    8b8d14feffff            mov ecx, dword ptr [ebp+fffffe14]

' *** Reference to "__vbaHresultCheckObj"
'00716740    8b3580104000            mov esi, dword ptr [00401080]
'00716746    6828010000              push 00000128
'0071674b    6830314300              push 00433130
'00716750    51                      push ecx
'00716751    50                      push eax
'00716752    ffd6                    call esi
'00716754    eb06                    jmp 71675c

' *** Reference to "__vbaHresultCheckObj"
'00716756    8b3580104000            mov esi, dword ptr [00401080]
'0071675c    8d4db8                  lea ecx, dword ptr [ebp-48]

' *** Reference to "__vbaFreeObj"
'0071675f    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_61)
'00716765    8d4d9c                  lea ecx, dword ptr [ebp-64]

' *** Reference to "__vbaFreeVar"
'00716768    ff1528104000            call dword ptr [00401028]
'#Cleanup(var_51)
'0071676e    8b8514feffff            mov eax, dword ptr [ebp+fffffe14]
'00716774    8b10                    mov edx, dword ptr [eax]
'00716776    6a00                    push 00
'00716778    6a01                    push 01
'0071677a    50                      push eax
'0071677b    ff9264010000            call dword ptr [edx+00000164]
'00716781    dbe2                    fnclex
'00716783    85c0                    test ax, ax
'00716785    7d14                    jge 71679b
'00716787    8b8d14feffff            mov ecx, dword ptr [ebp+fffffe14]
'0071678d    6864010000              push 00000164
'00716792    6830314300              push 00433130
'00716797    51                      push ecx
'00716798    50                      push eax
'00716799    ffd6                    call esi
'0071679b    8b45d0                  mov eax, dword ptr [ebp-30]
'0071679e    8b10                    mov edx, dword ptr [eax]
'007167a0    50                      push eax
'007167a1    ff92ec000000            call dword ptr [edx+000000ec]
'007167a7    dbe2                    fnclex
'007167a9    85c0                    test ax, ax
'007167ab    0f8d10e7ffff            jge 714ec1
'007167b1    8b4dd0                  mov ecx, dword ptr [ebp-30]
'007167b4    68ec000000              push 000000ec
'007167b9    6830314300              push 00433130
'007167be    51                      push ecx
'007167bf    50                      push eax
'007167c0    ffd6                    call esi
'007167c2    e9fae6ffff              jmp 714ec1

'ERROR: Two many next close:
Loop
'007167c7    6a00                    push 00
'007167c9    8d9514feffff            lea edx, dword ptr [ebp+fffffe14]
'007167cf    52                      push edx

' *** Reference to "__vbaObjSetAddref"
'007167d0    ff15c8104000            call dword ptr [004010c8]
Set var_695 = Nothing
'007167d6    8b45d0                  mov eax, dword ptr [ebp-30]
'007167d9    8b08                    mov ecx, dword ptr [eax]
'007167db    50                      push eax
'007167dc    ff91c4000000            call dword ptr [ecx+000000c4]
'007167e2    dbe2                    fnclex
'007167e4    85c0                    test ax, ax
'007167e6    7d11                    jge 7167f9
'007167e8    8b55d0                  mov edx, dword ptr [ebp-30]
'007167eb    68c4000000              push 000000c4
'007167f0    6830314300              push 00433130
'007167f5    52                      push edx
'007167f6    50                      push eax
'007167f7    ffd6                    call esi
'007167f9    8b45d4                  mov eax, dword ptr [ebp-2c]
'007167fc    8b08                    mov ecx, dword ptr [eax]
'007167fe    50                      push eax
'007167ff    ff91c4000000            call dword ptr [ecx+000000c4]
'00716805    dbe2                    fnclex
'00716807    85c0                    test ax, ax
'00716809    7d11                    jge 71681c
'0071680b    8b55d4                  mov edx, dword ptr [ebp-2c]
'0071680e    68c4000000              push 000000c4
'00716813    6830314300              push 00433130
'00716818    52                      push edx
'00716819    50                      push eax
'0071681a    ffd6                    call esi
'0071681c    8d5db8                  lea ebx, dword ptr [ebp-48]
'0071681f    53                      push ebx
'00716820    83ec10                  sub esp, 10
'00716823    8bdc                    mov ebx, esp
'00716825    8b4508                  mov eax, dword ptr [ebp+08]
'00716828    8b4038                  mov eax, dword ptr [eax+38]
'0071682b    ba0a000000              mov edx, 0000000a
'00716830    8913                    mov dword ptr [ebx], edx
'00716832    8b9530ffffff            mov edx, dword ptr [ebp+ffffff30]
'00716838    895304                  mov dword ptr [ebx+04], edx
'0071683b    8b38                    mov edi, dword ptr [eax]
'0071683d    b904000280              mov ecx, 80020004
'00716842    894b08                  mov dword ptr [ebx+08], ecx
'00716845    8b8d38ffffff            mov ecx, dword ptr [ebp+ffffff38]
'0071684b    894b0c                  mov dword ptr [ebx+0c], ecx
'0071684e    83ec10                  sub esp, 10
'00716851    8bd4                    mov edx, esp
'00716853    b90a000000              mov ecx, 0000000a
'00716858    890a                    mov dword ptr [edx], ecx
'0071685a    8b8d40ffffff            mov ecx, dword ptr [ebp+ffffff40]
'00716860    894a04                  mov dword ptr [edx+04], ecx
'00716863    b904000280              mov ecx, 80020004
'00716868    894a08                  mov dword ptr [edx+08], ecx
'0071686b    8b8d48ffffff            mov ecx, dword ptr [ebp+ffffff48]
'00716871    894a0c                  mov dword ptr [edx+0c], ecx
'00716874    83ec10                  sub esp, 10
'00716877    8bd4                    mov edx, esp
'00716879    b903000000              mov ecx, 00000003
'0071687e    890a                    mov dword ptr [edx], ecx
'00716880    8b8d50ffffff            mov ecx, dword ptr [ebp+ffffff50]
'00716886    894a04                  mov dword ptr [edx+04], ecx
'00716889    c78554ffffff01000000    mov dword ptr [ebp+ffffff54], 00000001
'00716893    8b8d54ffffff            mov ecx, dword ptr [ebp+ffffff54]
'00716899    894a08                  mov dword ptr [edx+08], ecx
'0071689c    8b8d58ffffff            mov ecx, dword ptr [ebp+ffffff58]
'007168a2    682cb84300              push 0043b82c
'007168a7    50                      push eax
'007168a8    894a0c                  mov dword ptr [edx+0c], ecx
'007168ab    ff97bc000000            call dword ptr [edi+000000bc]
'007168b1    dbe2                    fnclex
'007168b3    85c0                    test ax, ax
'007168b5    7d14                    jge 7168cb
'007168b7    8b5508                  mov edx, dword ptr [ebp+08]
'007168ba    8b4a38                  mov ecx, dword ptr [edx+38]
'007168bd    68bc000000              push 000000bc
'007168c2    68301f4300              push 00431f30
'007168c7    51                      push ecx
'007168c8    50                      push eax
'007168c9    ffd6                    call esi
'007168cb    8b45b8                  mov eax, dword ptr [ebp-48]
'007168ce    50                      push eax
'007168cf    8d55d0                  lea edx, dword ptr [ebp-30]
'007168d2    52                      push edx
'007168d3    c745b800000000          mov dword ptr [ebp-48], 00000000

' *** Reference to "__vbaObjSet"
'007168da    ff15b4104000            call dword ptr [004010b4]
Set var_4 = Nothing
'007168e0    8d5db8                  lea ebx, dword ptr [ebp-48]
'007168e3    53                      push ebx
'007168e4    83ec10                  sub esp, 10
'007168e7    8bdc                    mov ebx, esp
'007168e9    8b3d48b07200            mov edi, dword ptr [0072b048]
'007168ef    b90a000000              mov ecx, 0000000a
'007168f4    890b                    mov dword ptr [ebx], ecx
'007168f6    8b8d30ffffff            mov ecx, dword ptr [ebp+ffffff30]
'007168fc    894b04                  mov dword ptr [ebx+04], ecx
'007168ff    8b3f                    mov edi, dword ptr [edi]
'00716901    b804000280              mov eax, 80020004
'00716906    894308                  mov dword ptr [ebx+08], eax
'00716909    8bd0                    mov edx, eax
'0071690b    8b8538ffffff            mov eax, dword ptr [ebp+ffffff38]
'00716911    89430c                  mov dword ptr [ebx+0c], eax
'00716914    83ec10                  sub esp, 10
'00716917    8bcc                    mov ecx, esp
'00716919    b80a000000              mov eax, 0000000a
'0071691e    8901                    mov dword ptr [ecx], eax
'00716920    8b8540ffffff            mov eax, dword ptr [ebp+ffffff40]
'00716926    894104                  mov dword ptr [ecx+04], eax
'00716929    8b8550ffffff            mov eax, dword ptr [ebp+ffffff50]
'0071692f    895108                  mov dword ptr [ecx+08], edx
'00716932    8b9548ffffff            mov edx, dword ptr [ebp+ffffff48]
'00716938    89510c                  mov dword ptr [ecx+0c], edx
'0071693b    83ec10                  sub esp, 10
'0071693e    8bcc                    mov ecx, esp
'00716940    c7854cffffff03000000    mov dword ptr [ebp+ffffff4c], 00000003
'0071694a    8b954cffffff            mov edx, dword ptr [ebp+ffffff4c]
'00716950    8911                    mov dword ptr [ecx], edx
'00716952    8b9558ffffff            mov edx, dword ptr [ebp+ffffff58]
'00716958    894104                  mov dword ptr [ecx+04], eax
'0071695b    b802000000              mov eax, 00000002
'00716960    894108                  mov dword ptr [ecx+08], eax
'00716963    a148b07200              mov ax, word ptr [0072b048]
'00716968    6850b84300              push 0043b850
'0071696d    50                      push eax
'0071696e    89510c                  mov dword ptr [ecx+0c], edx
'00716971    ff97bc000000            call dword ptr [edi+000000bc]
'00716977    dbe2                    fnclex
'00716979    85c0                    test ax, ax
'0071697b    7d14                    jge 716991
'0071697d    8b0d48b07200            mov ecx, dword ptr [0072b048]
'00716983    68bc000000              push 000000bc
'00716988    68301f4300              push 00431f30
'0071698d    51                      push ecx
'0071698e    50                      push eax
'0071698f    ffd6                    call esi
'00716991    8b45b8                  mov eax, dword ptr [ebp-48]
'00716994    50                      push eax
'00716995    8d55d4                  lea edx, dword ptr [ebp-2c]
'00716998    52                      push edx
'00716999    c745b800000000          mov dword ptr [ebp-48], 00000000

' *** Reference to "__vbaObjSet"
'007169a0    ff15b4104000            call dword ptr [004010b4]
Set var_44 = Nothing

' *** Reference to "__vbaObjSetAddref"
'007169a6    8b3dc8104000            mov edi, dword ptr [004010c8]
'007169ac    8b45d0                  mov eax, dword ptr [ebp-30]
'007169af    8b08                    mov ecx, dword ptr [eax]
'007169b1    8d9538feffff            lea edx, dword ptr [ebp+fffffe38]
'007169b7    52                      push edx
'007169b8    50                      push eax
'007169b9    ff5134                  call dword ptr [ecx+34]
'007169bc    dbe2                    fnclex
'007169be    85c0                    test ax, ax
'007169c0    7d0e                    jge 7169d0
'007169c2    8b4dd0                  mov ecx, dword ptr [ebp-30]
'007169c5    6a34                    push 34
'007169c7    6830314300              push 00433130
'007169cc    51                      push ecx
'007169cd    50                      push eax
'007169ce    ffd6                    call esi
'007169d0    6683bd38feffff00        cmp word ptr [ebp+fffffe38], 00
'007169d8    0f85e3050000            jne 716fc1

Do While (0 = 0)
'007169de    8b55d4                  mov edx, dword ptr [ebp-2c]
'007169e1    52                      push edx
'007169e2    8d8510feffff            lea eax, dword ptr [ebp+fffffe10]
'007169e8    50                      push eax
'007169e9    ffd7                    call edi
    Set var_915 = Nothing
'007169eb    8b8510feffff            mov eax, dword ptr [ebp+fffffe10]
'007169f1    8b08                    mov ecx, dword ptr [eax]
'007169f3    50                      push eax
'007169f4    ff91c0000000            call dword ptr [ecx+000000c0]
'007169fa    dbe2                    fnclex
'007169fc    85c0                    test ax, ax
'007169fe    7d14                    jge 716a14
'00716a00    8b9510feffff            mov edx, dword ptr [ebp+fffffe10]
'00716a06    68c0000000              push 000000c0
'00716a0b    6830314300              push 00433130
'00716a10    52                      push edx
'00716a11    50                      push eax
'00716a12    ffd6                    call esi
'00716a14    8b45d0                  mov eax, dword ptr [ebp-30]
'00716a17    8b08                    mov ecx, dword ptr [eax]
'00716a19    8d55b8                  lea edx, dword ptr [ebp-48]
'00716a1c    52                      push edx
'00716a1d    50                      push eax
'00716a1e    ff91b4000000            call dword ptr [ecx+000000b4]
'00716a24    dbe2                    fnclex
'00716a26    85c0                    test ax, ax
'00716a28    7d11                    jge 716a3b
'00716a2a    8b4dd0                  mov ecx, dword ptr [ebp-30]
'00716a2d    68b4000000              push 000000b4
'00716a32    6830314300              push 00433130
'00716a37    51                      push ecx
'00716a38    50                      push eax
'00716a39    ffd6                    call esi
'00716a3b    8b45b8                  mov eax, dword ptr [ebp-48]
'00716a3e    8b38                    mov edi, dword ptr [eax]
'00716a40    8d5db4                  lea ebx, dword ptr [ebp-4c]
'00716a43    53                      push ebx
'00716a44    83ec10                  sub esp, 10
'00716a47    8bdc                    mov ebx, esp
'00716a49    ba08000000              mov edx, 00000008
'00716a4e    8913                    mov dword ptr [ebx], edx
'00716a50    8b9550ffffff            mov edx, dword ptr [ebp+ffffff50]
'00716a56    895304                  mov dword ptr [ebx+04], edx
'00716a59    b9c4a54300              mov ecx, 0043a5c4
'00716a5e    894b08                  mov dword ptr [ebx+08], ecx
'00716a61    8b8d58ffffff            mov ecx, dword ptr [ebp+ffffff58]
'00716a67    50                      push eax
'00716a68    898530feffff            mov dword ptr [ebp+fffffe30], eax
'00716a6e    894b0c                  mov dword ptr [ebx+0c], ecx
'00716a71    ff5730                  call dword ptr [edi+30]
'00716a74    dbe2                    fnclex
'00716a76    85c0                    test ax, ax
'00716a78    7d11                    jge 716a8b
'00716a7a    8b9530feffff            mov edx, dword ptr [ebp+fffffe30]
'00716a80    6a30                    push 30
'00716a82    68d8304300              push 004330d8
'00716a87    52                      push edx
'00716a88    50                      push eax
'00716a89    ffd6                    call esi
'00716a8b    8b45b4                  mov eax, dword ptr [ebp-4c]
'00716a8e    8bbd10feffff            mov edi, dword ptr [ebp+fffffe10]
'00716a94    83ec10                  sub esp, 10
'00716a97    8bdc                    mov ebx, esp
'00716a99    b909000000              mov ecx, 00000009
'00716a9e    890b                    mov dword ptr [ebx], ecx
'00716aa0    894d9c                  mov dword ptr [ebp-64], ecx
'00716aa3    8b4da0                  mov ecx, dword ptr [ebp-60]
'00716aa6    894b04                  mov dword ptr [ebx+04], ecx
'00716aa9    894308                  mov dword ptr [ebx+08], eax
'00716aac    8945a4                  mov dword ptr [ebp-5c], eax
'00716aaf    8b45a8                  mov eax, dword ptr [ebp-58]
'00716ab2    89430c                  mov dword ptr [ebx+0c], eax
'00716ab5    8b8540ffffff            mov eax, dword ptr [ebp+ffffff40]
'00716abb    83ec10                  sub esp, 10
'00716abe    8bcc                    mov ecx, esp
'00716ac0    be08000000              mov esi, 00000008
'00716ac5    8931                    mov dword ptr [ecx], esi
'00716ac7    894104                  mov dword ptr [ecx+04], eax
'00716aca    8b8510feffff            mov eax, dword ptr [ebp+fffffe10]
'00716ad0    bac4a54300              mov edx, 0043a5c4
'00716ad5    895108                  mov dword ptr [ecx+08], edx
'00716ad8    8b9548ffffff            mov edx, dword ptr [ebp+ffffff48]
'00716ade    c745b400000000          mov dword ptr [ebp-4c], 00000000
'00716ae5    8b3f                    mov edi, dword ptr [edi]
'00716ae7    50                      push eax
'00716ae8    89510c                  mov dword ptr [ecx+0c], edx
'00716aeb    ff9728010000            call dword ptr [edi+00000128]
'00716af1    dbe2                    fnclex
'00716af3    85c0                    test ax, ax
'00716af5    7d18                    jge 716b0f
'00716af7    8b8d10feffff            mov ecx, dword ptr [ebp+fffffe10]
'00716afd    6828010000              push 00000128
'00716b02    6830314300              push 00433130
'00716b07    51                      push ecx
'00716b08    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00716b09    ff1580104000            call dword ptr [00401080]
'00716b0f    8d4db8                  lea ecx, dword ptr [ebp-48]

' *** Reference to "__vbaFreeObj"
'00716b12    ff1508134000            call dword ptr [00401308]
    '#Cleanup(var_61)
'00716b18    8d4d9c                  lea ecx, dword ptr [ebp-64]

' *** Reference to "__vbaFreeVar"
'00716b1b    ff1528104000            call dword ptr [00401028]
    '#Cleanup(var_51)
'00716b21    8b45d0                  mov eax, dword ptr [ebp-30]
'00716b24    8b10                    mov edx, dword ptr [eax]
'00716b26    8d4db8                  lea ecx, dword ptr [ebp-48]
'00716b29    51                      push ecx
'00716b2a    50                      push eax
'00716b2b    ff92b4000000            call dword ptr [edx+000000b4]
'00716b31    dbe2                    fnclex
'00716b33    85c0                    test ax, ax
'00716b35    7d15                    jge 716b4c
'00716b37    8b55d0                  mov edx, dword ptr [ebp-30]
'00716b3a    68b4000000              push 000000b4
'00716b3f    6830314300              push 00433130
'00716b44    52                      push edx
'00716b45    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00716b46    ff1580104000            call dword ptr [00401080]
'00716b4c    8b45b8                  mov eax, dword ptr [ebp-48]
'00716b4f    8b38                    mov edi, dword ptr [eax]
'00716b51    8d5db4                  lea ebx, dword ptr [ebp-4c]
'00716b54    53                      push ebx
'00716b55    83ec10                  sub esp, 10
'00716b58    8bdc                    mov ebx, esp
'00716b5a    ba08000000              mov edx, 00000008
'00716b5f    8913                    mov dword ptr [ebx], edx
'00716b61    8b9550ffffff            mov edx, dword ptr [ebp+ffffff50]
'00716b67    895304                  mov dword ptr [ebx+04], edx
'00716b6a    b9a4a74300              mov ecx, 0043a7a4
'00716b6f    894b08                  mov dword ptr [ebx+08], ecx
'00716b72    8b8d58ffffff            mov ecx, dword ptr [ebp+ffffff58]
'00716b78    50                      push eax
'00716b79    8bf0                    mov esi, eax
'00716b7b    894b0c                  mov dword ptr [ebx+0c], ecx
'00716b7e    ff5730                  call dword ptr [edi+30]
'00716b81    dbe2                    fnclex
'00716b83    85c0                    test ax, ax
'00716b85    7d0f                    jge 716b96
'00716b87    6a30                    push 30
'00716b89    68d8304300              push 004330d8
'00716b8e    56                      push esi
'00716b8f    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00716b90    ff1580104000            call dword ptr [00401080]
'00716b96    8b45b4                  mov eax, dword ptr [ebp-4c]
'00716b99    8bbd10feffff            mov edi, dword ptr [ebp+fffffe10]
'00716b9f    83ec10                  sub esp, 10
'00716ba2    8bdc                    mov ebx, esp
'00716ba4    b909000000              mov ecx, 00000009
'00716ba9    890b                    mov dword ptr [ebx], ecx
'00716bab    894d9c                  mov dword ptr [ebp-64], ecx
'00716bae    8b4da0                  mov ecx, dword ptr [ebp-60]
'00716bb1    894b04                  mov dword ptr [ebx+04], ecx
'00716bb4    894308                  mov dword ptr [ebx+08], eax
'00716bb7    8945a4                  mov dword ptr [ebp-5c], eax
'00716bba    8b45a8                  mov eax, dword ptr [ebp-58]
'00716bbd    89430c                  mov dword ptr [ebx+0c], eax
'00716bc0    8b8540ffffff            mov eax, dword ptr [ebp+ffffff40]
'00716bc6    83ec10                  sub esp, 10
'00716bc9    8bcc                    mov ecx, esp
'00716bcb    be08000000              mov esi, 00000008
'00716bd0    8931                    mov dword ptr [ecx], esi
'00716bd2    894104                  mov dword ptr [ecx+04], eax
'00716bd5    8b8510feffff            mov eax, dword ptr [ebp+fffffe10]
'00716bdb    baa4a74300              mov edx, 0043a7a4
'00716be0    895108                  mov dword ptr [ecx+08], edx
'00716be3    8b9548ffffff            mov edx, dword ptr [ebp+ffffff48]
'00716be9    c745b400000000          mov dword ptr [ebp-4c], 00000000
'00716bf0    8b3f                    mov edi, dword ptr [edi]
'00716bf2    50                      push eax
'00716bf3    89510c                  mov dword ptr [ecx+0c], edx
'00716bf6    ff9728010000            call dword ptr [edi+00000128]
'00716bfc    dbe2                    fnclex
'00716bfe    85c0                    test ax, ax
'00716c00    7d18                    jge 716c1a
'00716c02    8b8d10feffff            mov ecx, dword ptr [ebp+fffffe10]
'00716c08    6828010000              push 00000128
'00716c0d    6830314300              push 00433130
'00716c12    51                      push ecx
'00716c13    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00716c14    ff1580104000            call dword ptr [00401080]
'00716c1a    8d4db8                  lea ecx, dword ptr [ebp-48]

' *** Reference to "__vbaFreeObj"
'00716c1d    ff1508134000            call dword ptr [00401308]
    '#Cleanup(var_61)
'00716c23    8d4d9c                  lea ecx, dword ptr [ebp-64]

' *** Reference to "__vbaFreeVar"
'00716c26    ff1528104000            call dword ptr [00401028]
    '#Cleanup(var_51)
'00716c2c    8b45d0                  mov eax, dword ptr [ebp-30]
'00716c2f    8b10                    mov edx, dword ptr [eax]
'00716c31    8d4db8                  lea ecx, dword ptr [ebp-48]
'00716c34    51                      push ecx
'00716c35    50                      push eax
'00716c36    ff92b4000000            call dword ptr [edx+000000b4]
'00716c3c    dbe2                    fnclex
'00716c3e    85c0                    test ax, ax
'00716c40    7d15                    jge 716c57
'00716c42    8b55d0                  mov edx, dword ptr [ebp-30]
'00716c45    68b4000000              push 000000b4
'00716c4a    6830314300              push 00433130
'00716c4f    52                      push edx
'00716c50    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00716c51    ff1580104000            call dword ptr [00401080]
'00716c57    8b45b8                  mov eax, dword ptr [ebp-48]
'00716c5a    8b38                    mov edi, dword ptr [eax]
'00716c5c    8d5db4                  lea ebx, dword ptr [ebp-4c]
'00716c5f    53                      push ebx
'00716c60    83ec10                  sub esp, 10
'00716c63    8bdc                    mov ebx, esp
'00716c65    ba08000000              mov edx, 00000008
'00716c6a    8913                    mov dword ptr [ebx], edx
'00716c6c    8b9550ffffff            mov edx, dword ptr [ebp+ffffff50]
'00716c72    895304                  mov dword ptr [ebx+04], edx
'00716c75    b9c0a74300              mov ecx, 0043a7c0
'00716c7a    894b08                  mov dword ptr [ebx+08], ecx
'00716c7d    8b8d58ffffff            mov ecx, dword ptr [ebp+ffffff58]
'00716c83    50                      push eax
'00716c84    8bf0                    mov esi, eax
'00716c86    894b0c                  mov dword ptr [ebx+0c], ecx
'00716c89    ff5730                  call dword ptr [edi+30]
'00716c8c    dbe2                    fnclex
'00716c8e    85c0                    test ax, ax
'00716c90    7d0f                    jge 716ca1
'00716c92    6a30                    push 30
'00716c94    68d8304300              push 004330d8
'00716c99    56                      push esi
'00716c9a    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00716c9b    ff1580104000            call dword ptr [00401080]
'00716ca1    8b45b4                  mov eax, dword ptr [ebp-4c]
'00716ca4    8bbd10feffff            mov edi, dword ptr [ebp+fffffe10]
'00716caa    83ec10                  sub esp, 10
'00716cad    8bdc                    mov ebx, esp
'00716caf    b909000000              mov ecx, 00000009
'00716cb4    890b                    mov dword ptr [ebx], ecx
'00716cb6    894d9c                  mov dword ptr [ebp-64], ecx
'00716cb9    8b4da0                  mov ecx, dword ptr [ebp-60]
'00716cbc    894b04                  mov dword ptr [ebx+04], ecx
'00716cbf    894308                  mov dword ptr [ebx+08], eax
'00716cc2    8945a4                  mov dword ptr [ebp-5c], eax
'00716cc5    8b45a8                  mov eax, dword ptr [ebp-58]
'00716cc8    89430c                  mov dword ptr [ebx+0c], eax
'00716ccb    8b8540ffffff            mov eax, dword ptr [ebp+ffffff40]
'00716cd1    83ec10                  sub esp, 10
'00716cd4    8bcc                    mov ecx, esp
'00716cd6    be08000000              mov esi, 00000008
'00716cdb    8931                    mov dword ptr [ecx], esi
'00716cdd    894104                  mov dword ptr [ecx+04], eax
'00716ce0    8b8510feffff            mov eax, dword ptr [ebp+fffffe10]
'00716ce6    bac0a74300              mov edx, 0043a7c0
'00716ceb    895108                  mov dword ptr [ecx+08], edx
'00716cee    8b9548ffffff            mov edx, dword ptr [ebp+ffffff48]
'00716cf4    c745b400000000          mov dword ptr [ebp-4c], 00000000
'00716cfb    8b3f                    mov edi, dword ptr [edi]
'00716cfd    50                      push eax
'00716cfe    89510c                  mov dword ptr [ecx+0c], edx
'00716d01    ff9728010000            call dword ptr [edi+00000128]
'00716d07    dbe2                    fnclex
'00716d09    85c0                    test ax, ax
'00716d0b    7d18                    jge 716d25
'00716d0d    8b8d10feffff            mov ecx, dword ptr [ebp+fffffe10]
'00716d13    6828010000              push 00000128
'00716d18    6830314300              push 00433130
'00716d1d    51                      push ecx
'00716d1e    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00716d1f    ff1580104000            call dword ptr [00401080]
'00716d25    8d4db8                  lea ecx, dword ptr [ebp-48]

' *** Reference to "__vbaFreeObj"
'00716d28    ff1508134000            call dword ptr [00401308]
    '#Cleanup(var_61)
'00716d2e    8d4d9c                  lea ecx, dword ptr [ebp-64]

' *** Reference to "__vbaFreeVar"
'00716d31    ff1528104000            call dword ptr [00401028]
    '#Cleanup(var_51)
'00716d37    8b45d0                  mov eax, dword ptr [ebp-30]
'00716d3a    8b10                    mov edx, dword ptr [eax]
'00716d3c    8d4db8                  lea ecx, dword ptr [ebp-48]
'00716d3f    51                      push ecx
'00716d40    50                      push eax
'00716d41    ff92b4000000            call dword ptr [edx+000000b4]
'00716d47    dbe2                    fnclex
'00716d49    85c0                    test ax, ax
'00716d4b    7d15                    jge 716d62
'00716d4d    8b55d0                  mov edx, dword ptr [ebp-30]
'00716d50    68b4000000              push 000000b4
'00716d55    6830314300              push 00433130
'00716d5a    52                      push edx
'00716d5b    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00716d5c    ff1580104000            call dword ptr [00401080]
'00716d62    8b45b8                  mov eax, dword ptr [ebp-48]
'00716d65    8b38                    mov edi, dword ptr [eax]
'00716d67    8d5db4                  lea ebx, dword ptr [ebp-4c]
'00716d6a    53                      push ebx
'00716d6b    83ec10                  sub esp, 10
'00716d6e    8bdc                    mov ebx, esp
'00716d70    ba08000000              mov edx, 00000008
'00716d75    8913                    mov dword ptr [ebx], edx
'00716d77    8b9550ffffff            mov edx, dword ptr [ebp+ffffff50]
'00716d7d    895304                  mov dword ptr [ebx+04], edx
'00716d80    b9f0a74300              mov ecx, 0043a7f0
'00716d85    894b08                  mov dword ptr [ebx+08], ecx
'00716d88    8b8d58ffffff            mov ecx, dword ptr [ebp+ffffff58]
'00716d8e    50                      push eax
'00716d8f    8bf0                    mov esi, eax
'00716d91    894b0c                  mov dword ptr [ebx+0c], ecx
'00716d94    ff5730                  call dword ptr [edi+30]
'00716d97    dbe2                    fnclex
'00716d99    85c0                    test ax, ax
'00716d9b    7d0f                    jge 716dac
'00716d9d    6a30                    push 30
'00716d9f    68d8304300              push 004330d8
'00716da4    56                      push esi
'00716da5    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00716da6    ff1580104000            call dword ptr [00401080]
'00716dac    8b45b4                  mov eax, dword ptr [ebp-4c]
'00716daf    8bbd10feffff            mov edi, dword ptr [ebp+fffffe10]
'00716db5    83ec10                  sub esp, 10
'00716db8    8bdc                    mov ebx, esp
'00716dba    b909000000              mov ecx, 00000009
'00716dbf    890b                    mov dword ptr [ebx], ecx
'00716dc1    894d9c                  mov dword ptr [ebp-64], ecx
'00716dc4    8b4da0                  mov ecx, dword ptr [ebp-60]
'00716dc7    894b04                  mov dword ptr [ebx+04], ecx
'00716dca    894308                  mov dword ptr [ebx+08], eax
'00716dcd    8945a4                  mov dword ptr [ebp-5c], eax
'00716dd0    8b45a8                  mov eax, dword ptr [ebp-58]
'00716dd3    89430c                  mov dword ptr [ebx+0c], eax
'00716dd6    8b8540ffffff            mov eax, dword ptr [ebp+ffffff40]
'00716ddc    83ec10                  sub esp, 10
'00716ddf    8bcc                    mov ecx, esp
'00716de1    be08000000              mov esi, 00000008
'00716de6    8931                    mov dword ptr [ecx], esi
'00716de8    894104                  mov dword ptr [ecx+04], eax
'00716deb    8b8510feffff            mov eax, dword ptr [ebp+fffffe10]
'00716df1    baf0a74300              mov edx, 0043a7f0
'00716df6    895108                  mov dword ptr [ecx+08], edx
'00716df9    8b9548ffffff            mov edx, dword ptr [ebp+ffffff48]
'00716dff    c745b400000000          mov dword ptr [ebp-4c], 00000000
'00716e06    8b3f                    mov edi, dword ptr [edi]
'00716e08    50                      push eax
'00716e09    89510c                  mov dword ptr [ecx+0c], edx
'00716e0c    ff9728010000            call dword ptr [edi+00000128]
'00716e12    dbe2                    fnclex
'00716e14    85c0                    test ax, ax
'00716e16    7d18                    jge 716e30
'00716e18    8b8d10feffff            mov ecx, dword ptr [ebp+fffffe10]
'00716e1e    6828010000              push 00000128
'00716e23    6830314300              push 00433130
'00716e28    51                      push ecx
'00716e29    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00716e2a    ff1580104000            call dword ptr [00401080]
'00716e30    8d4db8                  lea ecx, dword ptr [ebp-48]

' *** Reference to "__vbaFreeObj"
'00716e33    ff1508134000            call dword ptr [00401308]
    '#Cleanup(var_61)
'00716e39    8d4d9c                  lea ecx, dword ptr [ebp-64]

' *** Reference to "__vbaFreeVar"
'00716e3c    ff1528104000            call dword ptr [00401028]
    '#Cleanup(var_51)
'00716e42    8b45d0                  mov eax, dword ptr [ebp-30]
'00716e45    8b10                    mov edx, dword ptr [eax]
'00716e47    8d4db8                  lea ecx, dword ptr [ebp-48]
'00716e4a    51                      push ecx
'00716e4b    50                      push eax
'00716e4c    ff92b4000000            call dword ptr [edx+000000b4]
'00716e52    dbe2                    fnclex
'00716e54    85c0                    test ax, ax
'00716e56    7d15                    jge 716e6d
'00716e58    8b55d0                  mov edx, dword ptr [ebp-30]
'00716e5b    68b4000000              push 000000b4
'00716e60    6830314300              push 00433130
'00716e65    52                      push edx
'00716e66    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00716e67    ff1580104000            call dword ptr [00401080]
'00716e6d    8b45b8                  mov eax, dword ptr [ebp-48]
'00716e70    8b38                    mov edi, dword ptr [eax]
'00716e72    8d5db4                  lea ebx, dword ptr [ebp-4c]
'00716e75    53                      push ebx
'00716e76    83ec10                  sub esp, 10
'00716e79    8bdc                    mov ebx, esp
'00716e7b    ba08000000              mov edx, 00000008
'00716e80    8913                    mov dword ptr [ebx], edx
'00716e82    8b9550ffffff            mov edx, dword ptr [ebp+ffffff50]
'00716e88    895304                  mov dword ptr [ebx+04], edx
'00716e8b    b9dca74300              mov ecx, 0043a7dc
'00716e90    894b08                  mov dword ptr [ebx+08], ecx
'00716e93    8b8d58ffffff            mov ecx, dword ptr [ebp+ffffff58]
'00716e99    50                      push eax
'00716e9a    8bf0                    mov esi, eax
'00716e9c    894b0c                  mov dword ptr [ebx+0c], ecx
'00716e9f    ff5730                  call dword ptr [edi+30]
'00716ea2    dbe2                    fnclex
'00716ea4    85c0                    test ax, ax
'00716ea6    7d0f                    jge 716eb7
'00716ea8    6a30                    push 30
'00716eaa    68d8304300              push 004330d8
'00716eaf    56                      push esi
'00716eb0    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00716eb1    ff1580104000            call dword ptr [00401080]
'00716eb7    8b45b4                  mov eax, dword ptr [ebp-4c]
'00716eba    8bbd10feffff            mov edi, dword ptr [ebp+fffffe10]
'00716ec0    83ec10                  sub esp, 10
'00716ec3    8bdc                    mov ebx, esp
'00716ec5    b909000000              mov ecx, 00000009
'00716eca    890b                    mov dword ptr [ebx], ecx
'00716ecc    894d9c                  mov dword ptr [ebp-64], ecx
'00716ecf    8b4da0                  mov ecx, dword ptr [ebp-60]
'00716ed2    894b04                  mov dword ptr [ebx+04], ecx
'00716ed5    894308                  mov dword ptr [ebx+08], eax
'00716ed8    8945a4                  mov dword ptr [ebp-5c], eax
'00716edb    8b45a8                  mov eax, dword ptr [ebp-58]
'00716ede    89430c                  mov dword ptr [ebx+0c], eax
'00716ee1    8b8540ffffff            mov eax, dword ptr [ebp+ffffff40]
'00716ee7    83ec10                  sub esp, 10
'00716eea    8bcc                    mov ecx, esp
'00716eec    be08000000              mov esi, 00000008
'00716ef1    8931                    mov dword ptr [ecx], esi
'00716ef3    894104                  mov dword ptr [ecx+04], eax
'00716ef6    8b8510feffff            mov eax, dword ptr [ebp+fffffe10]
'00716efc    badca74300              mov edx, 0043a7dc
'00716f01    895108                  mov dword ptr [ecx+08], edx
'00716f04    8b9548ffffff            mov edx, dword ptr [ebp+ffffff48]
'00716f0a    c745b400000000          mov dword ptr [ebp-4c], 00000000
'00716f11    8b3f                    mov edi, dword ptr [edi]
'00716f13    50                      push eax
'00716f14    89510c                  mov dword ptr [ecx+0c], edx
'00716f17    ff9728010000            call dword ptr [edi+00000128]
'00716f1d    dbe2                    fnclex
'00716f1f    85c0                    test ax, ax
'00716f21    7d1c                    jge 716f3f
'00716f23    8b8d10feffff            mov ecx, dword ptr [ebp+fffffe10]

' *** Reference to "__vbaHresultCheckObj"
'00716f29    8b3580104000            mov esi, dword ptr [00401080]
'00716f2f    6828010000              push 00000128
'00716f34    6830314300              push 00433130
'00716f39    51                      push ecx
'00716f3a    50                      push eax
'00716f3b    ffd6                    call esi
'00716f3d    eb06                    jmp 716f45

' *** Reference to "__vbaHresultCheckObj"
'00716f3f    8b3580104000            mov esi, dword ptr [00401080]
'00716f45    8d4db8                  lea ecx, dword ptr [ebp-48]

' *** Reference to "__vbaFreeObj"
'00716f48    ff1508134000            call dword ptr [00401308]
    '#Cleanup(var_61)
'00716f4e    8d4d9c                  lea ecx, dword ptr [ebp-64]

' *** Reference to "__vbaFreeVar"
'00716f51    ff1528104000            call dword ptr [00401028]
    '#Cleanup(var_51)
'00716f57    8b8510feffff            mov eax, dword ptr [ebp+fffffe10]
'00716f5d    8b10                    mov edx, dword ptr [eax]
'00716f5f    6a00                    push 00
'00716f61    6a01                    push 01
'00716f63    50                      push eax
'00716f64    ff9264010000            call dword ptr [edx+00000164]
'00716f6a    dbe2                    fnclex
'00716f6c    85c0                    test ax, ax
'00716f6e    7d14                    jge 716f84
'00716f70    8b8d10feffff            mov ecx, dword ptr [ebp+fffffe10]
'00716f76    6864010000              push 00000164
'00716f7b    6830314300              push 00433130
'00716f80    51                      push ecx
'00716f81    50                      push eax
'00716f82    ffd6                    call esi

' *** Reference to "__vbaObjSetAddref"
'00716f84    8b3dc8104000            mov edi, dword ptr [004010c8]
'00716f8a    6a00                    push 00
'00716f8c    8d9510feffff            lea edx, dword ptr [ebp+fffffe10]
'00716f92    52                      push edx
'00716f93    ffd7                    call edi
    Set var_915 = Nothing
'00716f95    8b45d0                  mov eax, dword ptr [ebp-30]
'00716f98    8b08                    mov ecx, dword ptr [eax]
'00716f9a    50                      push eax
'00716f9b    ff91ec000000            call dword ptr [ecx+000000ec]
'00716fa1    dbe2                    fnclex
'00716fa3    85c0                    test ax, ax
'00716fa5    0f8d01faffff            jge 7169ac
'00716fab    8b55d0                  mov edx, dword ptr [ebp-30]
'00716fae    68ec000000              push 000000ec
'00716fb3    6830314300              push 00433130
'00716fb8    52                      push edx
'00716fb9    50                      push eax
'00716fba    ffd6                    call esi
'00716fbc    e9ebf9ffff              jmp 7169ac
    
Loop
'00716fc1    8b45d4                  mov eax, dword ptr [ebp-2c]
'00716fc4    8b08                    mov ecx, dword ptr [eax]
'00716fc6    50                      push eax
'00716fc7    ff91c4000000            call dword ptr [ecx+000000c4]
'00716fcd    dbe2                    fnclex
'00716fcf    85c0                    test ax, ax
'00716fd1    7d11                    jge 716fe4
'00716fd3    8b55d4                  mov edx, dword ptr [ebp-2c]
'00716fd6    68c4000000              push 000000c4
'00716fdb    6830314300              push 00433130
'00716fe0    52                      push edx
'00716fe1    50                      push eax
'00716fe2    ffd6                    call esi
'00716fe4    8b45d0                  mov eax, dword ptr [ebp-30]
'00716fe7    8b08                    mov ecx, dword ptr [eax]
'00716fe9    50                      push eax
'00716fea    ff91c4000000            call dword ptr [ecx+000000c4]
'00716ff0    dbe2                    fnclex
'00716ff2    85c0                    test ax, ax
'00716ff4    7d15                    jge 71700b
'00716ff6    8b55d0                  mov edx, dword ptr [ebp-30]
'00716ff9    68c4000000              push 000000c4
'00716ffe    6830314300              push 00433130
'00717003    52                      push edx
'00717004    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00717005    ff1580104000            call dword ptr [00401080]

'ERROR: Two many next close:
End If
'0071700b    8d5db8                  lea ebx, dword ptr [ebp-48]
'0071700e    53                      push ebx
'0071700f    83ec10                  sub esp, 10
'00717012    8bdc                    mov ebx, esp
'00717014    8b7508                  mov esi, dword ptr [ebp+08]
'00717017    8b4638                  mov eax, dword ptr [esi+38]
'0071701a    ba0a000000              mov edx, 0000000a
'0071701f    8913                    mov dword ptr [ebx], edx
'00717021    8b9530ffffff            mov edx, dword ptr [ebp+ffffff30]
'00717027    895304                  mov dword ptr [ebx+04], edx
'0071702a    8b38                    mov edi, dword ptr [eax]
'0071702c    b904000280              mov ecx, 80020004
'00717031    894b08                  mov dword ptr [ebx+08], ecx
'00717034    8b8d38ffffff            mov ecx, dword ptr [ebp+ffffff38]
'0071703a    894b0c                  mov dword ptr [ebx+0c], ecx
'0071703d    83ec10                  sub esp, 10
'00717040    8bd4                    mov edx, esp
'00717042    b90a000000              mov ecx, 0000000a
'00717047    890a                    mov dword ptr [edx], ecx
'00717049    8b8d40ffffff            mov ecx, dword ptr [ebp+ffffff40]
'0071704f    894a04                  mov dword ptr [edx+04], ecx
'00717052    b904000280              mov ecx, 80020004
'00717057    894a08                  mov dword ptr [edx+08], ecx
'0071705a    8b8d48ffffff            mov ecx, dword ptr [ebp+ffffff48]
'00717060    894a0c                  mov dword ptr [edx+0c], ecx
'00717063    83ec10                  sub esp, 10
'00717066    8bd4                    mov edx, esp
'00717068    b903000000              mov ecx, 00000003
'0071706d    890a                    mov dword ptr [edx], ecx
'0071706f    8b8d50ffffff            mov ecx, dword ptr [ebp+ffffff50]
'00717075    894a04                  mov dword ptr [edx+04], ecx
'00717078    b901000000              mov ecx, 00000001
'0071707d    894a08                  mov dword ptr [edx+08], ecx
'00717080    8b8d58ffffff            mov ecx, dword ptr [ebp+ffffff58]
'00717086    68c8b14300              push 0043b1c8
'0071708b    50                      push eax
'0071708c    894a0c                  mov dword ptr [edx+0c], ecx
'0071708f    ff97bc000000            call dword ptr [edi+000000bc]
'00717095    dbe2                    fnclex
'00717097    85c0                    test ax, ax
'00717099    7d15                    jge 7170b0
'0071709b    8b5638                  mov edx, dword ptr [esi+38]
'0071709e    68bc000000              push 000000bc
'007170a3    68301f4300              push 00431f30
'007170a8    52                      push edx
'007170a9    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'007170aa    ff1580104000            call dword ptr [00401080]
'007170b0    8b45b8                  mov eax, dword ptr [ebp-48]
'007170b3    50                      push eax
'007170b4    8d45d0                  lea eax, dword ptr [ebp-30]
'007170b7    50                      push eax
'007170b8    c745b800000000          mov dword ptr [ebp-48], 00000000

' *** Reference to "__vbaObjSet"
'007170bf    ff15b4104000            call dword ptr [004010b4]
Set var_4 = Nothing
'007170c5    8d5db8                  lea ebx, dword ptr [ebp-48]
'007170c8    53                      push ebx
'007170c9    83ec10                  sub esp, 10
'007170cc    8bdc                    mov ebx, esp
'007170ce    b90a000000              mov ecx, 0000000a
'007170d3    890b                    mov dword ptr [ebx], ecx
'007170d5    8b8d30ffffff            mov ecx, dword ptr [ebp+ffffff30]
'007170db    894b04                  mov dword ptr [ebx+04], ecx
'007170de    8b3d48b07200            mov edi, dword ptr [0072b048]
'007170e4    b804000280              mov eax, 80020004
'007170e9    894308                  mov dword ptr [ebx+08], eax
'007170ec    8bd0                    mov edx, eax
'007170ee    898534ffffff            mov dword ptr [ebp+ffffff34], eax
'007170f4    8b8538ffffff            mov eax, dword ptr [ebp+ffffff38]
'007170fa    89430c                  mov dword ptr [ebx+0c], eax
'007170fd    8b3f                    mov edi, dword ptr [edi]
'007170ff    83ec10                  sub esp, 10
'00717102    8bcc                    mov ecx, esp
'00717104    b80a000000              mov eax, 0000000a
'00717109    8901                    mov dword ptr [ecx], eax
'0071710b    8b8540ffffff            mov eax, dword ptr [ebp+ffffff40]
'00717111    894104                  mov dword ptr [ecx+04], eax
'00717114    895108                  mov dword ptr [ecx+08], edx
'00717117    899544ffffff            mov dword ptr [ebp+ffffff44], edx
'0071711d    8b9548ffffff            mov edx, dword ptr [ebp+ffffff48]
'00717123    89510c                  mov dword ptr [ecx+0c], edx
'00717126    8b9550ffffff            mov edx, dword ptr [ebp+ffffff50]
'0071712c    83ec10                  sub esp, 10
'0071712f    8bcc                    mov ecx, esp
'00717131    b803000000              mov eax, 00000003
'00717136    8901                    mov dword ptr [ecx], eax
'00717138    895104                  mov dword ptr [ecx+04], edx
'0071713b    b801000000              mov eax, 00000001
'00717140    894108                  mov dword ptr [ecx+08], eax
'00717143    8b8558ffffff            mov eax, dword ptr [ebp+ffffff58]
'00717149    89410c                  mov dword ptr [ecx+0c], eax
'0071714c    8b0d48b07200            mov ecx, dword ptr [0072b048]
'00717152    68c8b14300              push 0043b1c8
'00717157    51                      push ecx
'00717158    ff97bc000000            call dword ptr [edi+000000bc]
'0071715e    dbe2                    fnclex
'00717160    85c0                    test ax, ax
'00717162    7d18                    jge 71717c
'00717164    8b1548b07200            mov edx, dword ptr [0072b048]
'0071716a    68bc000000              push 000000bc
'0071716f    68301f4300              push 00431f30
'00717174    52                      push edx
'00717175    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00717176    ff1580104000            call dword ptr [00401080]
'0071717c    8b45b8                  mov eax, dword ptr [ebp-48]

' *** Reference to "__vbaObjSet"
'0071717f    8b3db4104000            mov edi, dword ptr [004010b4]
'00717185    50                      push eax
'00717186    8d45d4                  lea eax, dword ptr [ebp-2c]
'00717189    50                      push eax
'0071718a    c745b800000000          mov dword ptr [ebp-48], 00000000
'00717191    ffd7                    call edi
Set var_44 = Nothing
'00717193    8b0e                    mov ecx, dword ptr [esi]
'00717195    6a00                    push 00
'00717197    6a07                    push 07
'00717199    56                      push esi
'0071719a    ff9110030000            call dword ptr [ecx+00000310]
'007171a0    50                      push eax
'007171a1    8d55b8                  lea edx, dword ptr [ebp-48]
'007171a4    52                      push edx
'007171a5    ffd7                    call edi
Set var_61 = Nothing
'007171a7    50                      push eax
'007171a8    8d459c                  lea eax, dword ptr [ebp-64]
'007171ab    50                      push eax

' *** Reference to "__vbaLateIdCallLd"
'007171ac    ff157c114000            call dword ptr [0040117c]
var_51 = var_61.UNK_-256 - 12_7
'007171b2    83c410                  add esp, 10
'007171b5    50                      push eax

' *** Reference to "__vbaI4Var"
'007171b6    ff157c124000            call dword ptr [0040127c]
'007171bc    8bc8                    mov ecx, eax
'007171be    83e901                  sub ecx, 01
var_num3 = CLng(var_51) - 1
'007171c1    0f80ac360000            jo 71a873

' *** Reference to "__vbaI2I4"
'007171c7    ff1550114000            call dword ptr [00401150]
'007171cd    bf01000000              mov edi, 00000001
'007171d2    8d4db8                  lea ecx, dword ptr [ebp-48]
'007171d5    897de8                  mov dword ptr [ebp-18], edi
'007171d8    898508feffff            mov dword ptr [ebp+fffffe08], eax

' *** Reference to "__vbaFreeObj"
'007171de    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_61)
'007171e4    8d4d9c                  lea ecx, dword ptr [ebp-64]

' *** Reference to "__vbaFreeVar"
'007171e7    ff1528104000            call dword ptr [00401028]
'#Cleanup(var_51)
'007171ed    663bbd08feffff          cmp di, word ptr [ebp+fffffe08]
'007171f4    0f8fd3330000            jg 71a5cd

Do While (1 <= WORD PTR [EBP+FFFFFE08])
'007171fa    83ec10                  sub esp, 10
'007171fd    8bdc                    mov ebx, esp
'007171ff    33c0                    xor eax, eax
    var_num1 = Empty
'00717201    b903000000              mov ecx, 00000003
'00717206    890b                    mov dword ptr [ebx], ecx
'00717208    8b8d50ffffff            mov ecx, dword ptr [ebp+ffffff50]
'0071720e    894b04                  mov dword ptr [ebx+04], ecx
'00717211    894308                  mov dword ptr [ebx+08], eax
'00717214    8b8558ffffff            mov eax, dword ptr [ebp+ffffff58]
'0071721a    89430c                  mov dword ptr [ebx+0c], eax
'0071721d    83ec10                  sub esp, 10
'00717220    8bcc                    mov ecx, esp
'00717222    ba02000000              mov edx, 00000002
'00717227    8911                    mov dword ptr [ecx], edx
'00717229    8b9530ffffff            mov edx, dword ptr [ebp+ffffff30]
'0071722f    895104                  mov dword ptr [ecx+04], edx
'00717232    8b9538ffffff            mov edx, dword ptr [ebp+ffffff38]
'00717238    6689bd34ffffff          mov word ptr [ebp+ffffff34], di
'0071723f    8b8534ffffff            mov eax, dword ptr [ebp+ffffff34]
'00717245    894108                  mov dword ptr [ecx+08], eax
'00717248    89510c                  mov dword ptr [ecx+0c], edx
'0071724b    8b9510ffffff            mov edx, dword ptr [ebp+ffffff10]
'00717251    83ec10                  sub esp, 10
'00717254    8bcc                    mov ecx, esp
'00717256    b802000000              mov eax, 00000002
'0071725b    8901                    mov dword ptr [ecx], eax
'0071725d    895104                  mov dword ptr [ecx+04], edx
'00717260    b805000000              mov eax, 00000005
'00717265    894108                  mov dword ptr [ecx+08], eax
'00717268    8b8518ffffff            mov eax, dword ptr [ebp+ffffff18]
'0071726e    6a03                    push 03
'00717270    689e000000              push 0000009e
'00717275    89410c                  mov dword ptr [ecx+0c], eax
'00717278    8b0e                    mov ecx, dword ptr [esi]
'0071727a    56                      push esi
'0071727b    c785f4fefffffce94400    mov dword ptr [ebp+fffffef4], 0044e9fc
'00717285    c785ecfeffff08800000    mov dword ptr [ebp+fffffeec], 00008008
'0071728f    ff9110030000            call dword ptr [ecx+00000310]
'00717295    50                      push eax
'00717296    8d55b8                  lea edx, dword ptr [ebp-48]
'00717299    52                      push edx

' *** Reference to "__vbaObjSet"
'0071729a    ff15b4104000            call dword ptr [004010b4]
    Set var_61 = Nothing
'007172a0    50                      push eax
'007172a1    8d459c                  lea eax, dword ptr [ebp-64]
'007172a4    50                      push eax

' *** Reference to "__vbaLateIdCallLd"
'007172a5    ff157c114000            call dword ptr [0040117c]
    var_51 = var_61.UNK_-256 - 12_158
'007172ab    83c440                  add esp, 40
'007172ae    50                      push eax
'007172af    8d8decfeffff            lea ecx, dword ptr [ebp+fffffeec]
'007172b5    51                      push ecx

' *** Reference to "__vbaVarTstEq"
'007172b6    ff153c114000            call dword ptr [0040113c]
'007172bc    8d4db8                  lea ecx, dword ptr [ebp-48]
'007172bf    668bf0                  mov si, ax

' *** Reference to "__vbaFreeObj"
'007172c2    ff1508134000            call dword ptr [00401308]
    '#Cleanup(var_61)
'007172c8    8d4d9c                  lea ecx, dword ptr [ebp-64]

' *** Reference to "__vbaFreeVar"
'007172cb    ff1528104000            call dword ptr [00401028]
    '#Cleanup(var_51)
'007172d1    6685f6                  test esi, esi
'007172d4    0f84d8320000            je 71a5b2
    
    If (    ((var_51) = ("Oui"))) Then
'007172da    8b45d4                  mov eax, dword ptr [ebp-2c]
'007172dd    8b10                    mov edx, dword ptr [eax]
'007172df    6830a84300              push 0043a830
'007172e4    50                      push eax
'007172e5    ff5244                  call dword ptr [edx+44]
'007172e8    dbe2                    fnclex
'007172ea    85c0                    test ax, ax
'007172ec    7d12                    jge 717300
'007172ee    8b4dd4                  mov ecx, dword ptr [ebp-2c]
'007172f1    6a44                    push 44
'007172f3    6830314300              push 00433130
'007172f8    51                      push ecx
'007172f9    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'007172fa    ff1580104000            call dword ptr [00401080]
'00717300    83ec10                  sub esp, 10
'00717303    8bdc                    mov ebx, esp
'00717305    33c0                    xor eax, eax
    var_num1 = Empty
'00717307    b903000000              mov ecx, 00000003
'0071730c    890b                    mov dword ptr [ebx], ecx
'0071730e    8b8d50ffffff            mov ecx, dword ptr [ebp+ffffff50]
'00717314    894b04                  mov dword ptr [ebx+04], ecx
'00717317    894308                  mov dword ptr [ebx+08], eax
'0071731a    8b8558ffffff            mov eax, dword ptr [ebp+ffffff58]
'00717320    89430c                  mov dword ptr [ebx+0c], eax
'00717323    83ec10                  sub esp, 10
'00717326    8bcc                    mov ecx, esp
'00717328    ba02000000              mov edx, 00000002
'0071732d    8911                    mov dword ptr [ecx], edx
'0071732f    6689bd34ffffff          mov word ptr [ebp+ffffff34], di
'00717336    8b8534ffffff            mov eax, dword ptr [ebp+ffffff34]
'0071733c    8bfa                    mov edi, edx
'0071733e    8b9530ffffff            mov edx, dword ptr [ebp+ffffff30]
'00717344    895104                  mov dword ptr [ecx+04], edx
'00717347    8b9538ffffff            mov edx, dword ptr [ebp+ffffff38]
'0071734d    894108                  mov dword ptr [ecx+08], eax
'00717350    89510c                  mov dword ptr [ecx+0c], edx
'00717353    8b8d10ffffff            mov ecx, dword ptr [ebp+ffffff10]
'00717359    8b9518ffffff            mov edx, dword ptr [ebp+ffffff18]
'0071735f    83ec10                  sub esp, 10
'00717362    8bc4                    mov eax, esp
'00717364    8938                    mov dword ptr [eax], edi
'00717366    894804                  mov dword ptr [eax+04], ecx
'00717369    be01000000              mov esi, 00000001
'0071736e    897008                  mov dword ptr [eax+08], esi
'00717371    6a03                    push 03
'00717373    89500c                  mov dword ptr [eax+0c], edx
'00717376    8b4508                  mov eax, dword ptr [ebp+08]
'00717379    8b08                    mov ecx, dword ptr [eax]
'0071737b    689e000000              push 0000009e
'00717380    50                      push eax
'00717381    ff9110030000            call dword ptr [ecx+00000310]
'00717387    50                      push eax
'00717388    8d55b8                  lea edx, dword ptr [ebp-48]
'0071738b    52                      push edx

' *** Reference to "__vbaObjSet"
'0071738c    ff15b4104000            call dword ptr [004010b4]
    Set var_61 = Me
'00717392    50                      push eax
'00717393    8d459c                  lea eax, dword ptr [ebp-64]
'00717396    50                      push eax

' *** Reference to "__vbaLateIdCallLd"
'00717397    ff157c114000            call dword ptr [0040117c]
    var_51 = var_61.UNK_frmImport_158
'0071739d    be0a000000              mov esi, 0000000a
'007173a2    8b7dd4                  mov edi, dword ptr [ebp-2c]
'007173a5    83c430                  add esp, 30
'007173a8    8bdc                    mov ebx, esp
'007173aa    8bce                    mov ecx, esi
'007173ac    890b                    mov dword ptr [ebx], ecx
'007173ae    8b8d40feffff            mov ecx, dword ptr [ebp+fffffe40]
'007173b4    894b04                  mov dword ptr [ebx+04], ecx
'007173b7    ba04000280              mov edx, 80020004
'007173bc    8bc2                    mov eax, edx
'007173be    894308                  mov dword ptr [ebx+08], eax
'007173c1    8b8548feffff            mov eax, dword ptr [ebp+fffffe48]
'007173c7    89430c                  mov dword ptr [ebx+0c], eax
'007173ca    8b8550feffff            mov eax, dword ptr [ebp+fffffe50]
'007173d0    83ec10                  sub esp, 10
'007173d3    8bcc                    mov ecx, esp
'007173d5    8931                    mov dword ptr [ecx], esi
'007173d7    894104                  mov dword ptr [ecx+04], eax
'007173da    895108                  mov dword ptr [ecx+08], edx
'007173dd    8995f4feffff            mov dword ptr [ebp+fffffef4], edx
'007173e3    8b9558feffff            mov edx, dword ptr [ebp+fffffe58]
'007173e9    89510c                  mov dword ptr [ecx+0c], edx
'007173ec    8b9560feffff            mov edx, dword ptr [ebp+fffffe60]
'007173f2    83ec10                  sub esp, 10
'007173f5    8bcc                    mov ecx, esp
'007173f7    8bc6                    mov eax, esi
'007173f9    8901                    mov dword ptr [ecx], eax
'007173fb    895104                  mov dword ptr [ecx+04], edx
'007173fe    b804000280              mov eax, 80020004
'00717403    894108                  mov dword ptr [ecx+08], eax
'00717406    8b8568feffff            mov eax, dword ptr [ebp+fffffe68]
'0071740c    83ec10                  sub esp, 10
'0071740f    89b5ecfeffff            mov dword ptr [ebp+fffffeec], esi
'00717415    8b3f                    mov edi, dword ptr [edi]
'00717417    89410c                  mov dword ptr [ecx+0c], eax
'0071741a    8bcc                    mov ecx, esp
'0071741c    8b9570feffff            mov edx, dword ptr [ebp+fffffe70]
'00717422    8bc6                    mov eax, esi
'00717424    8901                    mov dword ptr [ecx], eax
'00717426    895104                  mov dword ptr [ecx+04], edx
'00717429    8b9580feffff            mov edx, dword ptr [ebp+fffffe80]
'0071742f    83ec10                  sub esp, 10
'00717432    b804000280              mov eax, 80020004
'00717437    894108                  mov dword ptr [ecx+08], eax
'0071743a    8b8578feffff            mov eax, dword ptr [ebp+fffffe78]
'00717440    89410c                  mov dword ptr [ecx+0c], eax
'00717443    8bcc                    mov ecx, esp
'00717445    8bc6                    mov eax, esi
'00717447    8901                    mov dword ptr [ecx], eax
'00717449    895104                  mov dword ptr [ecx+04], edx
'0071744c    8b9590feffff            mov edx, dword ptr [ebp+fffffe90]
'00717452    83ec10                  sub esp, 10
'00717455    b804000280              mov eax, 80020004
'0071745a    894108                  mov dword ptr [ecx+08], eax
'0071745d    8b8588feffff            mov eax, dword ptr [ebp+fffffe88]
'00717463    89410c                  mov dword ptr [ecx+0c], eax
'00717466    8bcc                    mov ecx, esp
'00717468    8bc6                    mov eax, esi
'0071746a    8901                    mov dword ptr [ecx], eax
'0071746c    895104                  mov dword ptr [ecx+04], edx
'0071746f    8b95a0feffff            mov edx, dword ptr [ebp+fffffea0]
'00717475    b804000280              mov eax, 80020004
'0071747a    894108                  mov dword ptr [ecx+08], eax
'0071747d    8b8598feffff            mov eax, dword ptr [ebp+fffffe98]
'00717483    89410c                  mov dword ptr [ecx+0c], eax
'00717486    83ec10                  sub esp, 10
'00717489    8bcc                    mov ecx, esp
'0071748b    8bc6                    mov eax, esi
'0071748d    8901                    mov dword ptr [ecx], eax
'0071748f    895104                  mov dword ptr [ecx+04], edx
'00717492    8b95b0feffff            mov edx, dword ptr [ebp+fffffeb0]
'00717498    b804000280              mov eax, 80020004
'0071749d    894108                  mov dword ptr [ecx+08], eax
'007174a0    8b85a8feffff            mov eax, dword ptr [ebp+fffffea8]
'007174a6    89410c                  mov dword ptr [ecx+0c], eax
'007174a9    83ec10                  sub esp, 10
'007174ac    8bcc                    mov ecx, esp
'007174ae    8bc6                    mov eax, esi
'007174b0    8901                    mov dword ptr [ecx], eax
'007174b2    895104                  mov dword ptr [ecx+04], edx
'007174b5    8b95c0feffff            mov edx, dword ptr [ebp+fffffec0]
'007174bb    b804000280              mov eax, 80020004
'007174c0    894108                  mov dword ptr [ecx+08], eax
'007174c3    8b85b8feffff            mov eax, dword ptr [ebp+fffffeb8]
'007174c9    89410c                  mov dword ptr [ecx+0c], eax
'007174cc    83ec10                  sub esp, 10
'007174cf    8bcc                    mov ecx, esp
'007174d1    8bc6                    mov eax, esi
'007174d3    8901                    mov dword ptr [ecx], eax
'007174d5    895104                  mov dword ptr [ecx+04], edx
'007174d8    8b95d0feffff            mov edx, dword ptr [ebp+fffffed0]
'007174de    b804000280              mov eax, 80020004
'007174e3    894108                  mov dword ptr [ecx+08], eax
'007174e6    8b85c8feffff            mov eax, dword ptr [ebp+fffffec8]
'007174ec    89410c                  mov dword ptr [ecx+0c], eax
'007174ef    83ec10                  sub esp, 10
'007174f2    8bcc                    mov ecx, esp
'007174f4    8bc6                    mov eax, esi
'007174f6    8901                    mov dword ptr [ecx], eax
'007174f8    895104                  mov dword ptr [ecx+04], edx
'007174fb    8b95e0feffff            mov edx, dword ptr [ebp+fffffee0]
'00717501    b804000280              mov eax, 80020004
'00717506    894108                  mov dword ptr [ecx+08], eax
'00717509    8b85d8feffff            mov eax, dword ptr [ebp+fffffed8]
'0071750f    89410c                  mov dword ptr [ecx+0c], eax
'00717512    83ec10                  sub esp, 10
'00717515    8bcc                    mov ecx, esp
'00717517    8bc6                    mov eax, esi
'00717519    8901                    mov dword ptr [ecx], eax
'0071751b    895104                  mov dword ptr [ecx+04], edx
'0071751e    b804000280              mov eax, 80020004
'00717523    894108                  mov dword ptr [ecx+08], eax
'00717526    8b85e8feffff            mov eax, dword ptr [ebp+fffffee8]
'0071752c    83ec10                  sub esp, 10
'0071752f    89410c                  mov dword ptr [ecx+0c], eax
'00717532    8bcc                    mov ecx, esp
'00717534    8bd6                    mov edx, esi
'00717536    8b85f0feffff            mov eax, dword ptr [ebp+fffffef0]
'0071753c    8911                    mov dword ptr [ecx], edx
'0071753e    8b95f4feffff            mov edx, dword ptr [ebp+fffffef4]
'00717544    894104                  mov dword ptr [ecx+04], eax
'00717547    8b85f8feffff            mov eax, dword ptr [ebp+fffffef8]
'0071754d    895108                  mov dword ptr [ecx+08], edx
'00717550    8b559c                  mov edx, dword ptr [ebp-64]
'00717553    89410c                  mov dword ptr [ecx+0c], eax
'00717556    8b45a0                  mov eax, dword ptr [ebp-60]
'00717559    83ec10                  sub esp, 10
'0071755c    8bcc                    mov ecx, esp
'0071755e    8911                    mov dword ptr [ecx], edx
'00717560    8b55a4                  mov edx, dword ptr [ebp-5c]
'00717563    894104                  mov dword ptr [ecx+04], eax
'00717566    8b45a8                  mov eax, dword ptr [ebp-58]
'00717569    895108                  mov dword ptr [ecx+08], edx
'0071756c    89410c                  mov dword ptr [ecx+0c], eax
'0071756f    8b4dd4                  mov ecx, dword ptr [ebp-2c]
'00717572    68209e4300              push 00439e20
'00717577    51                      push ecx
'00717578    ff97f4000000            call dword ptr [edi+000000f4]
'0071757e    dbe2                    fnclex
'00717580    85c0                    test ax, ax
'00717582    7d15                    jge 717599
'00717584    8b55d4                  mov edx, dword ptr [ebp-2c]
'00717587    68f4000000              push 000000f4
'0071758c    6830314300              push 00433130
'00717591    52                      push edx
'00717592    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00717593    ff1580104000            call dword ptr [00401080]
'00717599    8d4db8                  lea ecx, dword ptr [ebp-48]

' *** Reference to "__vbaFreeObj"
'0071759c    ff1508134000            call dword ptr [00401308]
    '#Cleanup(var_61)
'007175a2    8d4d9c                  lea ecx, dword ptr [ebp-64]

' *** Reference to "__vbaFreeVar"
'007175a5    ff1528104000            call dword ptr [00401028]
    '#Cleanup(var_51)
'007175ab    8b45d0                  mov eax, dword ptr [ebp-30]
'007175ae    8b08                    mov ecx, dword ptr [eax]
'007175b0    6830a84300              push 0043a830
'007175b5    50                      push eax
'007175b6    ff5144                  call dword ptr [ecx+44]
'007175b9    dbe2                    fnclex
'007175bb    85c0                    test ax, ax
'007175bd    7d12                    jge 7175d1
'007175bf    8b55d0                  mov edx, dword ptr [ebp-30]
'007175c2    6a44                    push 44
'007175c4    6830314300              push 00433130
'007175c9    52                      push edx
'007175ca    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'007175cb    ff1580104000            call dword ptr [00401080]
'007175d1    668b55e8                mov dx, word ptr [ebp-18]
'007175d5    83ec10                  sub esp, 10
'007175d8    8bdc                    mov ebx, esp
'007175da    33c0                    xor eax, eax
    var_num1 = Empty
'007175dc    b903000000              mov ecx, 00000003
'007175e1    890b                    mov dword ptr [ebx], ecx
'007175e3    8b8d50ffffff            mov ecx, dword ptr [ebp+ffffff50]
'007175e9    894b04                  mov dword ptr [ebx+04], ecx
'007175ec    894308                  mov dword ptr [ebx+08], eax
'007175ef    8b8558ffffff            mov eax, dword ptr [ebp+ffffff58]
'007175f5    89430c                  mov dword ptr [ebx+0c], eax
'007175f8    83ec10                  sub esp, 10
'007175fb    66899534ffffff          mov word ptr [ebp+ffffff34], dx
'00717602    8b8534ffffff            mov eax, dword ptr [ebp+ffffff34]
'00717608    8bcc                    mov ecx, esp
'0071760a    ba02000000              mov edx, 00000002
'0071760f    8911                    mov dword ptr [ecx], edx
'00717611    8bfa                    mov edi, edx
'00717613    8b9530ffffff            mov edx, dword ptr [ebp+ffffff30]
'00717619    895104                  mov dword ptr [ecx+04], edx
'0071761c    8b9538ffffff            mov edx, dword ptr [ebp+ffffff38]
'00717622    894108                  mov dword ptr [ecx+08], eax
'00717625    89510c                  mov dword ptr [ecx+0c], edx
'00717628    8b8d10ffffff            mov ecx, dword ptr [ebp+ffffff10]
'0071762e    8b9518ffffff            mov edx, dword ptr [ebp+ffffff18]
'00717634    83ec10                  sub esp, 10
'00717637    8bc4                    mov eax, esp
'00717639    8938                    mov dword ptr [eax], edi
'0071763b    894804                  mov dword ptr [eax+04], ecx
'0071763e    33f6                    xor esi, esi
    var_num8 = Empty
'00717640    897008                  mov dword ptr [eax+08], esi
'00717643    6a03                    push 03
'00717645    89500c                  mov dword ptr [eax+0c], edx
'00717648    8b4508                  mov eax, dword ptr [ebp+08]
'0071764b    8b08                    mov ecx, dword ptr [eax]
'0071764d    689e000000              push 0000009e
'00717652    50                      push eax
'00717653    ff9110030000            call dword ptr [ecx+00000310]
'00717659    50                      push eax
'0071765a    8d55b8                  lea edx, dword ptr [ebp-48]
'0071765d    52                      push edx

' *** Reference to "__vbaObjSet"
'0071765e    ff15b4104000            call dword ptr [004010b4]
    Set var_61 = Me
'00717664    50                      push eax
'00717665    8d459c                  lea eax, dword ptr [ebp-64]
'00717668    50                      push eax

' *** Reference to "__vbaLateIdCallLd"
'00717669    ff157c114000            call dword ptr [0040117c]
    var_51 = var_61.UNK_frmImport_158
'0071766f    8b7dd0                  mov edi, dword ptr [ebp-30]
'00717672    be0a000000              mov esi, 0000000a
'00717677    83c430                  add esp, 30
'0071767a    8bdc                    mov ebx, esp
'0071767c    8bce                    mov ecx, esi
'0071767e    890b                    mov dword ptr [ebx], ecx
'00717680    8b8d40feffff            mov ecx, dword ptr [ebp+fffffe40]
'00717686    894b04                  mov dword ptr [ebx+04], ecx
'00717689    ba04000280              mov edx, 80020004
'0071768e    8bc2                    mov eax, edx
'00717690    894308                  mov dword ptr [ebx+08], eax
'00717693    8b8548feffff            mov eax, dword ptr [ebp+fffffe48]
'00717699    89430c                  mov dword ptr [ebx+0c], eax
'0071769c    8b8550feffff            mov eax, dword ptr [ebp+fffffe50]
'007176a2    83ec10                  sub esp, 10
'007176a5    8bcc                    mov ecx, esp
'007176a7    8931                    mov dword ptr [ecx], esi
'007176a9    894104                  mov dword ptr [ecx+04], eax
'007176ac    895108                  mov dword ptr [ecx+08], edx
'007176af    8995f4feffff            mov dword ptr [ebp+fffffef4], edx
'007176b5    8b9558feffff            mov edx, dword ptr [ebp+fffffe58]
'007176bb    89510c                  mov dword ptr [ecx+0c], edx
'007176be    8b9560feffff            mov edx, dword ptr [ebp+fffffe60]
'007176c4    83ec10                  sub esp, 10
'007176c7    8bcc                    mov ecx, esp
'007176c9    8bc6                    mov eax, esi
'007176cb    8901                    mov dword ptr [ecx], eax
'007176cd    895104                  mov dword ptr [ecx+04], edx
'007176d0    b804000280              mov eax, 80020004
'007176d5    894108                  mov dword ptr [ecx+08], eax
'007176d8    8b8568feffff            mov eax, dword ptr [ebp+fffffe68]
'007176de    89b5ecfeffff            mov dword ptr [ebp+fffffeec], esi
'007176e4    8b3f                    mov edi, dword ptr [edi]
'007176e6    89410c                  mov dword ptr [ecx+0c], eax
'007176e9    83ec10                  sub esp, 10
'007176ec    8b9570feffff            mov edx, dword ptr [ebp+fffffe70]
'007176f2    8bcc                    mov ecx, esp
'007176f4    8bc6                    mov eax, esi
'007176f6    8901                    mov dword ptr [ecx], eax
'007176f8    895104                  mov dword ptr [ecx+04], edx
'007176fb    8b9580feffff            mov edx, dword ptr [ebp+fffffe80]
'00717701    83ec10                  sub esp, 10
'00717704    b804000280              mov eax, 80020004
'00717709    894108                  mov dword ptr [ecx+08], eax
'0071770c    8b8578feffff            mov eax, dword ptr [ebp+fffffe78]
'00717712    89410c                  mov dword ptr [ecx+0c], eax
'00717715    8bcc                    mov ecx, esp
'00717717    8bc6                    mov eax, esi
'00717719    8901                    mov dword ptr [ecx], eax
'0071771b    895104                  mov dword ptr [ecx+04], edx
'0071771e    8b9590feffff            mov edx, dword ptr [ebp+fffffe90]
'00717724    83ec10                  sub esp, 10
'00717727    b804000280              mov eax, 80020004
'0071772c    894108                  mov dword ptr [ecx+08], eax
'0071772f    8b8588feffff            mov eax, dword ptr [ebp+fffffe88]
'00717735    89410c                  mov dword ptr [ecx+0c], eax
'00717738    8bcc                    mov ecx, esp
'0071773a    8bc6                    mov eax, esi
'0071773c    8901                    mov dword ptr [ecx], eax
'0071773e    895104                  mov dword ptr [ecx+04], edx
'00717741    8b95a0feffff            mov edx, dword ptr [ebp+fffffea0]
'00717747    b804000280              mov eax, 80020004
'0071774c    894108                  mov dword ptr [ecx+08], eax
'0071774f    8b8598feffff            mov eax, dword ptr [ebp+fffffe98]
'00717755    89410c                  mov dword ptr [ecx+0c], eax
'00717758    83ec10                  sub esp, 10
'0071775b    8bcc                    mov ecx, esp
'0071775d    8bc6                    mov eax, esi
'0071775f    8901                    mov dword ptr [ecx], eax
'00717761    895104                  mov dword ptr [ecx+04], edx
'00717764    8b95b0feffff            mov edx, dword ptr [ebp+fffffeb0]
'0071776a    b804000280              mov eax, 80020004
'0071776f    894108                  mov dword ptr [ecx+08], eax
'00717772    8b85a8feffff            mov eax, dword ptr [ebp+fffffea8]
'00717778    89410c                  mov dword ptr [ecx+0c], eax
'0071777b    83ec10                  sub esp, 10
'0071777e    8bcc                    mov ecx, esp
'00717780    8bc6                    mov eax, esi
'00717782    8901                    mov dword ptr [ecx], eax
'00717784    895104                  mov dword ptr [ecx+04], edx
'00717787    8b95c0feffff            mov edx, dword ptr [ebp+fffffec0]
'0071778d    b804000280              mov eax, 80020004
'00717792    894108                  mov dword ptr [ecx+08], eax
'00717795    8b85b8feffff            mov eax, dword ptr [ebp+fffffeb8]
'0071779b    89410c                  mov dword ptr [ecx+0c], eax
'0071779e    83ec10                  sub esp, 10
'007177a1    8bcc                    mov ecx, esp
'007177a3    8bc6                    mov eax, esi
'007177a5    8901                    mov dword ptr [ecx], eax
'007177a7    895104                  mov dword ptr [ecx+04], edx
'007177aa    8b95d0feffff            mov edx, dword ptr [ebp+fffffed0]
'007177b0    b804000280              mov eax, 80020004
'007177b5    894108                  mov dword ptr [ecx+08], eax
'007177b8    8b85c8feffff            mov eax, dword ptr [ebp+fffffec8]
'007177be    89410c                  mov dword ptr [ecx+0c], eax
'007177c1    83ec10                  sub esp, 10
'007177c4    8bcc                    mov ecx, esp
'007177c6    8bc6                    mov eax, esi
'007177c8    8901                    mov dword ptr [ecx], eax
'007177ca    895104                  mov dword ptr [ecx+04], edx
'007177cd    8b95e0feffff            mov edx, dword ptr [ebp+fffffee0]
'007177d3    b804000280              mov eax, 80020004
'007177d8    894108                  mov dword ptr [ecx+08], eax
'007177db    8b85d8feffff            mov eax, dword ptr [ebp+fffffed8]
'007177e1    89410c                  mov dword ptr [ecx+0c], eax
'007177e4    83ec10                  sub esp, 10
'007177e7    8bcc                    mov ecx, esp
'007177e9    8bc6                    mov eax, esi
'007177eb    8901                    mov dword ptr [ecx], eax
'007177ed    895104                  mov dword ptr [ecx+04], edx
'007177f0    b804000280              mov eax, 80020004
'007177f5    894108                  mov dword ptr [ecx+08], eax
'007177f8    8b85e8feffff            mov eax, dword ptr [ebp+fffffee8]
'007177fe    83ec10                  sub esp, 10
'00717801    89410c                  mov dword ptr [ecx+0c], eax
'00717804    8bcc                    mov ecx, esp
'00717806    8b85f0feffff            mov eax, dword ptr [ebp+fffffef0]
'0071780c    8bd6                    mov edx, esi
'0071780e    8911                    mov dword ptr [ecx], edx
'00717810    8b95f4feffff            mov edx, dword ptr [ebp+fffffef4]
'00717816    894104                  mov dword ptr [ecx+04], eax
'00717819    8b85f8feffff            mov eax, dword ptr [ebp+fffffef8]
'0071781f    895108                  mov dword ptr [ecx+08], edx
'00717822    8b559c                  mov edx, dword ptr [ebp-64]
'00717825    89410c                  mov dword ptr [ecx+0c], eax
'00717828    8b45a0                  mov eax, dword ptr [ebp-60]
'0071782b    83ec10                  sub esp, 10
'0071782e    8bcc                    mov ecx, esp
'00717830    8911                    mov dword ptr [ecx], edx
'00717832    8b55a4                  mov edx, dword ptr [ebp-5c]
'00717835    894104                  mov dword ptr [ecx+04], eax
'00717838    8b45a8                  mov eax, dword ptr [ebp-58]
'0071783b    895108                  mov dword ptr [ecx+08], edx
'0071783e    89410c                  mov dword ptr [ecx+0c], eax
'00717841    8b4dd0                  mov ecx, dword ptr [ebp-30]
'00717844    68209e4300              push 00439e20
'00717849    51                      push ecx
'0071784a    ff97f4000000            call dword ptr [edi+000000f4]
'00717850    dbe2                    fnclex
'00717852    85c0                    test ax, ax
'00717854    7d19                    jge 71786f
'00717856    8b55d0                  mov edx, dword ptr [ebp-30]

' *** Reference to "__vbaHresultCheckObj"
'00717859    8b3580104000            mov esi, dword ptr [00401080]
'0071785f    68f4000000              push 000000f4
'00717864    6830314300              push 00433130
'00717869    52                      push edx
'0071786a    50                      push eax
'0071786b    ffd6                    call esi
'0071786d    eb06                    jmp 717875

' *** Reference to "__vbaHresultCheckObj"
'0071786f    8b3580104000            mov esi, dword ptr [00401080]
'00717875    8d4db8                  lea ecx, dword ptr [ebp-48]

' *** Reference to "__vbaFreeObj"
'00717878    ff1508134000            call dword ptr [00401308]
    '#Cleanup(var_61)
'0071787e    8d4d9c                  lea ecx, dword ptr [ebp-64]

' *** Reference to "__vbaFreeVar"
'00717881    ff1528104000            call dword ptr [00401028]
    '#Cleanup(var_51)
'00717887    8b45d4                  mov eax, dword ptr [ebp-2c]
'0071788a    8b08                    mov ecx, dword ptr [eax]
'0071788c    8d9538feffff            lea edx, dword ptr [ebp+fffffe38]
'00717892    52                      push edx
'00717893    50                      push eax
'00717894    ff515c                  call dword ptr [ecx+5c]
'00717897    dbe2                    fnclex
'00717899    85c0                    test ax, ax
'0071789b    7d0e                    jge 7178ab
'0071789d    8b4dd4                  mov ecx, dword ptr [ebp-2c]
'007178a0    6a5c                    push 5c
'007178a2    6830314300              push 00433130
'007178a7    51                      push ecx
'007178a8    50                      push eax
'007178a9    ffd6                    call esi
'007178ab    6683bd38feffff00        cmp word ptr [ebp+fffffe38], 00
'007178b3    0f84b02b0000            je 71a469
    
    If (    0 <> 0) Then
'007178b9    8b45d4                  mov eax, dword ptr [ebp-2c]
'007178bc    8b10                    mov edx, dword ptr [eax]
'007178be    50                      push eax
'007178bf    ff92c0000000            call dword ptr [edx+000000c0]
'007178c5    dbe2                    fnclex
'007178c7    85c0                    test ax, ax
'007178c9    7d11                    jge 7178dc
'007178cb    8b4dd4                  mov ecx, dword ptr [ebp-2c]
'007178ce    68c0000000              push 000000c0
'007178d3    6830314300              push 00433130
'007178d8    51                      push ecx
'007178d9    50                      push eax
'007178da    ffd6                    call esi
'007178dc    8b45d4                  mov eax, dword ptr [ebp-2c]
'007178df    8b10                    mov edx, dword ptr [eax]
'007178e1    8d4db4                  lea ecx, dword ptr [ebp-4c]
'007178e4    51                      push ecx
'007178e5    50                      push eax
'007178e6    ff92b4000000            call dword ptr [edx+000000b4]
'007178ec    dbe2                    fnclex
'007178ee    85c0                    test ax, ax
'007178f0    7d11                    jge 717903
'007178f2    8b55d4                  mov edx, dword ptr [ebp-2c]
'007178f5    68b4000000              push 000000b4
'007178fa    6830314300              push 00433130
'007178ff    52                      push edx
'00717900    50                      push eax
'00717901    ffd6                    call esi
'00717903    8b45b4                  mov eax, dword ptr [ebp-4c]
'00717906    8d5db0                  lea ebx, dword ptr [ebp-50]
'00717909    53                      push ebx
'0071790a    83ec10                  sub esp, 10
'0071790d    ba02000000              mov edx, 00000002
'00717912    8bdc                    mov ebx, esp
'00717914    8913                    mov dword ptr [ebx], edx
'00717916    8995ecfeffff            mov dword ptr [ebp+fffffeec], edx
'0071791c    8b95f0feffff            mov edx, dword ptr [ebp+fffffef0]
'00717922    33c9                    xor ecx, ecx
'00717924    895304                  mov dword ptr [ebx+04], edx
'00717927    898df4feffff            mov dword ptr [ebp+fffffef4], ecx
'0071792d    8b38                    mov edi, dword ptr [eax]
'0071792f    894b08                  mov dword ptr [ebx+08], ecx
'00717932    8b8df8feffff            mov ecx, dword ptr [ebp+fffffef8]
'00717938    50                      push eax
'00717939    8bf0                    mov esi, eax
'0071793b    894b0c                  mov dword ptr [ebx+0c], ecx
'0071793e    ff5730                  call dword ptr [edi+30]
'00717941    dbe2                    fnclex
'00717943    85c0                    test ax, ax
'00717945    7d0f                    jge 717956
'00717947    6a30                    push 30
'00717949    68d8304300              push 004330d8
'0071794e    56                      push esi
'0071794f    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00717950    ff1580104000            call dword ptr [00401080]
'00717956    668b55e8                mov dx, word ptr [ebp-18]
'0071795a    83ec10                  sub esp, 10
'0071795d    8b75b0                  mov esi, dword ptr [ebp-50]
'00717960    8bdc                    mov ebx, esp
'00717962    8b3e                    mov edi, dword ptr [esi]
'00717964    33c0                    xor eax, eax
    var_num1 = Empty
'00717966    b903000000              mov ecx, 00000003
'0071796b    890b                    mov dword ptr [ebx], ecx
'0071796d    8b8d50ffffff            mov ecx, dword ptr [ebp+ffffff50]
'00717973    894b04                  mov dword ptr [ebx+04], ecx
'00717976    894308                  mov dword ptr [ebx+08], eax
'00717979    898554ffffff            mov dword ptr [ebp+ffffff54], eax
'0071797f    8b8558ffffff            mov eax, dword ptr [ebp+ffffff58]
'00717985    89430c                  mov dword ptr [ebx+0c], eax
'00717988    8b5d08                  mov ebx, dword ptr [ebp+08]
'0071798b    66899534ffffff          mov word ptr [ebp+ffffff34], dx
'00717992    8b8534ffffff            mov eax, dword ptr [ebp+ffffff34]
'00717998    83ec10                  sub esp, 10
'0071799b    8bcc                    mov ecx, esp
'0071799d    ba02000000              mov edx, 00000002
'007179a2    8911                    mov dword ptr [ecx], edx
'007179a4    89950cffffff            mov dword ptr [ebp+ffffff0c], edx
'007179aa    8b9530ffffff            mov edx, dword ptr [ebp+ffffff30]
'007179b0    895104                  mov dword ptr [ecx+04], edx
'007179b3    8b9538ffffff            mov edx, dword ptr [ebp+ffffff38]
'007179b9    894108                  mov dword ptr [ecx+08], eax
'007179bc    8b8510ffffff            mov eax, dword ptr [ebp+ffffff10]
'007179c2    89510c                  mov dword ptr [ecx+0c], edx
'007179c5    8b950cffffff            mov edx, dword ptr [ebp+ffffff0c]
'007179cb    83ec10                  sub esp, 10
'007179ce    8bcc                    mov ecx, esp
'007179d0    8911                    mov dword ptr [ecx], edx
'007179d2    8b9518ffffff            mov edx, dword ptr [ebp+ffffff18]
'007179d8    894104                  mov dword ptr [ecx+04], eax
'007179db    6a03                    push 03
'007179dd    b801000000              mov eax, 00000001
'007179e2    894108                  mov dword ptr [ecx+08], eax
'007179e5    8b03                    mov eax, dword ptr [ebx]
'007179e7    689e000000              push 0000009e
'007179ec    53                      push ebx
'007179ed    89510c                  mov dword ptr [ecx+0c], edx
'007179f0    ff9010030000            call dword ptr [eax+00000310]
'007179f6    50                      push eax
'007179f7    8d4db8                  lea ecx, dword ptr [ebp-48]
'007179fa    51                      push ecx

' *** Reference to "__vbaObjSet"
'007179fb    ff15b4104000            call dword ptr [004010b4]
    Set var_61 = Nothing
'00717a01    50                      push eax
'00717a02    8d559c                  lea edx, dword ptr [ebp-64]
'00717a05    52                      push edx

' *** Reference to "__vbaLateIdCallLd"
'00717a06    ff157c114000            call dword ptr [0040117c]
    var_51 = var_61.UNK_-256 - 12_158
'00717a0c    8b10                    mov edx, dword ptr [eax]
'00717a0e    83c430                  add esp, 30
'00717a11    8bcc                    mov ecx, esp
'00717a13    8911                    mov dword ptr [ecx], edx
'00717a15    8b5004                  mov edx, dword ptr [eax+04]
'00717a18    895104                  mov dword ptr [ecx+04], edx
'00717a1b    8b5008                  mov edx, dword ptr [eax+08]
'00717a1e    8b400c                  mov eax, dword ptr [eax+0c]
'00717a21    895108                  mov dword ptr [ecx+08], edx
'00717a24    56                      push esi
'00717a25    89410c                  mov dword ptr [ecx+0c], eax
'00717a28    ff5748                  call dword ptr [edi+48]
'00717a2b    dbe2                    fnclex
'00717a2d    85c0                    test ax, ax
'00717a2f    7d0f                    jge 717a40
'00717a31    6a48                    push 48
'00717a33    6880304300              push 00433080
'00717a38    56                      push esi
'00717a39    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00717a3a    ff1580104000            call dword ptr [00401080]
'00717a40    8d4db0                  lea ecx, dword ptr [ebp-50]
'00717a43    51                      push ecx
'00717a44    8d55b4                  lea edx, dword ptr [ebp-4c]
'00717a47    52                      push edx
'00717a48    8d45b8                  lea eax, dword ptr [ebp-48]
'00717a4b    50                      push eax
'00717a4c    6a03                    push 03

' *** Reference to "__vbaFreeObjList"
'00717a4e    ff154c104000            call dword ptr [0040104c]
    '#Cleanup( var_61, var_62, var_6)
'00717a54    83c410                  add esp, 10
'00717a57    8d4d9c                  lea ecx, dword ptr [ebp-64]

' *** Reference to "__vbaFreeVar"
'00717a5a    ff1528104000            call dword ptr [00401028]
    '#Cleanup(var_51)
'00717a60    bf01000000              mov edi, 00000001
'00717a65    b84a000000              mov eax, 0000004a
'00717a6a    663bf8                  cmp di, ax
'00717a6d    897de0                  mov dword ptr [ebp-20], edi
'00717a70    0f8f4a020000            jg 717cc0
    
    Do While (    1 <= 74)
'00717a76    8b45d0                  mov eax, dword ptr [ebp-30]
'00717a79    8b08                    mov ecx, dword ptr [eax]
'00717a7b    8d55b8                  lea edx, dword ptr [ebp-48]
'00717a7e    52                      push edx
'00717a7f    50                      push eax
'00717a80    ff91b4000000            call dword ptr [ecx+000000b4]
'00717a86    dbe2                    fnclex
'00717a88    85c0                    test ax, ax
'00717a8a    7d15                    jge 717aa1
'00717a8c    8b4dd0                  mov ecx, dword ptr [ebp-30]
'00717a8f    68b4000000              push 000000b4
'00717a94    6830314300              push 00433130
'00717a99    51                      push ecx
'00717a9a    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00717a9b    ff1580104000            call dword ptr [00401080]
'00717aa1    8b45b8                  mov eax, dword ptr [ebp-48]
'00717aa4    8b10                    mov edx, dword ptr [eax]
'00717aa6    6689bd54ffffff          mov word ptr [ebp+ffffff54], di
'00717aad    8d7db4                  lea edi, dword ptr [ebp-4c]
'00717ab0    57                      push edi
'00717ab1    83ec10                  sub esp, 10
'00717ab4    8bfc                    mov edi, esp
'00717ab6    b902000000              mov ecx, 00000002
'00717abb    890f                    mov dword ptr [edi], ecx
'00717abd    8b8d50ffffff            mov ecx, dword ptr [ebp+ffffff50]
'00717ac3    894f04                  mov dword ptr [edi+04], ecx
'00717ac6    8b8d54ffffff            mov ecx, dword ptr [ebp+ffffff54]
'00717acc    894f08                  mov dword ptr [edi+08], ecx
'00717acf    8b8d58ffffff            mov ecx, dword ptr [ebp+ffffff58]
'00717ad5    50                      push eax
'00717ad6    8bf0                    mov esi, eax
'00717ad8    894f0c                  mov dword ptr [edi+0c], ecx
'00717adb    ff5230                  call dword ptr [edx+30]
'00717ade    dbe2                    fnclex
'00717ae0    85c0                    test ax, ax
'00717ae2    7d0f                    jge 717af3
'00717ae4    6a30                    push 30
'00717ae6    68d8304300              push 004330d8
'00717aeb    56                      push esi
'00717aec    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00717aed    ff1580104000            call dword ptr [00401080]
'00717af3    8b45b4                  mov eax, dword ptr [ebp-4c]
'00717af6    8d559c                  lea edx, dword ptr [ebp-64]
'00717af9    52                      push edx
'00717afa    c745b400000000          mov dword ptr [ebp-4c], 00000000
'00717b01    8945a4                  mov dword ptr [ebp-5c], eax
'00717b04    c7459c09000000          mov dword ptr [ebp-64], 00000009

' *** Reference to "rtcIsNull"
'00717b0b    ff1540114000            call dword ptr [00401140]
'00717b11    33f6                    xor esi, esi
    var_num8 = Empty
'00717b13    668bf0                  mov si, ax
'00717b16    8d4db8                  lea ecx, dword ptr [ebp-48]
'00717b19    f7d6                    not esi

' *** Reference to "__vbaFreeObj"
'00717b1b    ff1508134000            call dword ptr [00401308]
    '#Cleanup(var_61)
'00717b21    8d4d9c                  lea ecx, dword ptr [ebp-64]

' *** Reference to "__vbaFreeVar"
'00717b24    ff1528104000            call dword ptr [00401028]
    '#Cleanup(var_51)
'00717b2a    6685f6                  test esi, esi
'00717b2d    0f8477010000            je 717caa
    
    If (    Not (IsNull(var_62))) Then
'00717b33    8b45d4                  mov eax, dword ptr [ebp-2c]
'00717b36    8b08                    mov ecx, dword ptr [eax]
'00717b38    8d55b0                  lea edx, dword ptr [ebp-50]
'00717b3b    52                      push edx
'00717b3c    50                      push eax
'00717b3d    ff91b4000000            call dword ptr [ecx+000000b4]
'00717b43    dbe2                    fnclex
'00717b45    85c0                    test ax, ax
'00717b47    7d15                    jge 717b5e
'00717b49    8b4dd4                  mov ecx, dword ptr [ebp-2c]
'00717b4c    68b4000000              push 000000b4
'00717b51    6830314300              push 00433130
'00717b56    51                      push ecx
'00717b57    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00717b58    ff1580104000            call dword ptr [00401080]
'00717b5e    668b55e0                mov dx, word ptr [ebp-20]
'00717b62    8b45b0                  mov eax, dword ptr [ebp-50]
'00717b65    8d7dac                  lea edi, dword ptr [ebp-54]
'00717b68    57                      push edi
'00717b69    83ec10                  sub esp, 10
'00717b6c    8bfc                    mov edi, esp
'00717b6e    b902000000              mov ecx, 00000002
'00717b73    890f                    mov dword ptr [edi], ecx
'00717b75    8b8d40ffffff            mov ecx, dword ptr [ebp+ffffff40]
'00717b7b    894f04                  mov dword ptr [edi+04], ecx
'00717b7e    66899544ffffff          mov word ptr [ebp+ffffff44], dx
'00717b85    8b8d44ffffff            mov ecx, dword ptr [ebp+ffffff44]
'00717b8b    8b10                    mov edx, dword ptr [eax]
'00717b8d    894f08                  mov dword ptr [edi+08], ecx
'00717b90    8b8d48ffffff            mov ecx, dword ptr [ebp+ffffff48]
'00717b96    50                      push eax
'00717b97    8bf0                    mov esi, eax
'00717b99    894f0c                  mov dword ptr [edi+0c], ecx
'00717b9c    ff5230                  call dword ptr [edx+30]
'00717b9f    dbe2                    fnclex
'00717ba1    85c0                    test ax, ax
'00717ba3    7d0f                    jge 717bb4
'00717ba5    6a30                    push 30
'00717ba7    68d8304300              push 004330d8
'00717bac    56                      push esi
'00717bad    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00717bae    ff1580104000            call dword ptr [00401080]
'00717bb4    8b55ac                  mov edx, dword ptr [ebp-54]
'00717bb7    8b45d0                  mov eax, dword ptr [ebp-30]
'00717bba    8b08                    mov ecx, dword ptr [eax]
'00717bbc    89951cfeffff            mov dword ptr [ebp+fffffe1c], edx
'00717bc2    8d55b8                  lea edx, dword ptr [ebp-48]
'00717bc5    52                      push edx
'00717bc6    50                      push eax
'00717bc7    ff91b4000000            call dword ptr [ecx+000000b4]
'00717bcd    dbe2                    fnclex
'00717bcf    85c0                    test ax, ax
'00717bd1    7d15                    jge 717be8
'00717bd3    8b4dd0                  mov ecx, dword ptr [ebp-30]
'00717bd6    68b4000000              push 000000b4
'00717bdb    6830314300              push 00433130
'00717be0    51                      push ecx
'00717be1    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00717be2    ff1580104000            call dword ptr [00401080]
'00717be8    668b55e0                mov dx, word ptr [ebp-20]
'00717bec    8b45b8                  mov eax, dword ptr [ebp-48]
'00717bef    8d7db4                  lea edi, dword ptr [ebp-4c]
'00717bf2    57                      push edi
'00717bf3    83ec10                  sub esp, 10
'00717bf6    8bfc                    mov edi, esp
'00717bf8    b902000000              mov ecx, 00000002
'00717bfd    890f                    mov dword ptr [edi], ecx
'00717bff    8b8d50ffffff            mov ecx, dword ptr [ebp+ffffff50]
'00717c05    894f04                  mov dword ptr [edi+04], ecx
'00717c08    66899554ffffff          mov word ptr [ebp+ffffff54], dx
'00717c0f    8b8d54ffffff            mov ecx, dword ptr [ebp+ffffff54]
'00717c15    8b10                    mov edx, dword ptr [eax]
'00717c17    894f08                  mov dword ptr [edi+08], ecx
'00717c1a    8b8d58ffffff            mov ecx, dword ptr [ebp+ffffff58]
'00717c20    50                      push eax
'00717c21    8bf0                    mov esi, eax
'00717c23    894f0c                  mov dword ptr [edi+0c], ecx
'00717c26    ff5230                  call dword ptr [edx+30]
'00717c29    dbe2                    fnclex
'00717c2b    85c0                    test ax, ax
'00717c2d    7d0f                    jge 717c3e
'00717c2f    6a30                    push 30
'00717c31    68d8304300              push 004330d8
'00717c36    56                      push esi
'00717c37    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00717c38    ff1580104000            call dword ptr [00401080]
'00717c3e    8b45b4                  mov eax, dword ptr [ebp-4c]
'00717c41    8bb51cfeffff            mov esi, dword ptr [ebp+fffffe1c]
'00717c47    83ec10                  sub esp, 10
'00717c4a    b909000000              mov ecx, 00000009
'00717c4f    8bfc                    mov edi, esp
'00717c51    890f                    mov dword ptr [edi], ecx
'00717c53    894d9c                  mov dword ptr [ebp-64], ecx
'00717c56    8b4da0                  mov ecx, dword ptr [ebp-60]
'00717c59    894f04                  mov dword ptr [edi+04], ecx
'00717c5c    8945a4                  mov dword ptr [ebp-5c], eax
'00717c5f    894708                  mov dword ptr [edi+08], eax
'00717c62    8b45a8                  mov eax, dword ptr [ebp-58]
'00717c65    c745b400000000          mov dword ptr [ebp-4c], 00000000
'00717c6c    8b16                    mov edx, dword ptr [esi]
'00717c6e    56                      push esi
'00717c6f    89470c                  mov dword ptr [edi+0c], eax
'00717c72    ff5248                  call dword ptr [edx+48]
'00717c75    dbe2                    fnclex
'00717c77    85c0                    test ax, ax
'00717c79    7d0f                    jge 717c8a
'00717c7b    6a48                    push 48
'00717c7d    6880304300              push 00433080
'00717c82    56                      push esi
'00717c83    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00717c84    ff1580104000            call dword ptr [00401080]
'00717c8a    8d4dac                  lea ecx, dword ptr [ebp-54]
'00717c8d    51                      push ecx
'00717c8e    8d55b0                  lea edx, dword ptr [ebp-50]
'00717c91    52                      push edx
'00717c92    8d45b8                  lea eax, dword ptr [ebp-48]
'00717c95    50                      push eax
'00717c96    6a03                    push 03

' *** Reference to "__vbaFreeObjList"
'00717c98    ff154c104000            call dword ptr [0040104c]
    '#Cleanup( var_61, var_6, var_50)
'00717c9e    83c410                  add esp, 10
'00717ca1    8d4d9c                  lea ecx, dword ptr [ebp-64]

' *** Reference to "__vbaFreeVar"
'00717ca4    ff1528104000            call dword ptr [00401028]
    '#Cleanup(var_51)
End If
'00717caa    b801000000              mov eax, 00000001
'00717caf    660345e0                add ax, word ptr [ebp-20]
var_num1 = 1 + 1
'00717cb3    0f80ba2b0000            jo 71a873
'00717cb9    8bf8                    mov edi, eax
'00717cbb    e9a5fdffff              jmp 717a65

'ERROR: Two many next close:
Loop
'00717cc0    8b45d4                  mov eax, dword ptr [ebp-2c]
'00717cc3    8b08                    mov ecx, dword ptr [eax]
'00717cc5    6a00                    push 00
'00717cc7    6a01                    push 01
'00717cc9    50                      push eax
'00717cca    ff9164010000            call dword ptr [ecx+00000164]
'00717cd0    dbe2                    fnclex
'00717cd2    85c0                    test ax, ax
'00717cd4    7d15                    jge 717ceb
'00717cd6    8b55d4                  mov edx, dword ptr [ebp-2c]
'00717cd9    6864010000              push 00000164
'00717cde    6830314300              push 00433130
'00717ce3    52                      push edx
'00717ce4    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00717ce5    ff1580104000            call dword ptr [00401080]
'00717ceb    8b03                    mov eax, dword ptr [ebx]
'00717ced    53                      push ebx
'00717cee    ff9004030000            call dword ptr [eax+00000304]
'00717cf4    50                      push eax
'00717cf5    8d4db8                  lea ecx, dword ptr [ebp-48]
'00717cf8    51                      push ecx

' *** Reference to "__vbaObjSet"
'00717cf9    ff15b4104000            call dword ptr [004010b4]
Set var_61 = Nothing
'00717cff    8bf0                    mov esi, eax
'00717d01    8b16                    mov edx, dword ptr [esi]
'00717d03    8d8538feffff            lea eax, dword ptr [ebp+fffffe38]
'00717d09    50                      push eax
'00717d0a    56                      push esi
'00717d0b    ff92e0000000            call dword ptr [edx+000000e0]
'00717d11    dbe2                    fnclex
'00717d13    85c0                    test ax, ax
'00717d15    7d12                    jge 717d29
'00717d17    68e0000000              push 000000e0
'00717d1c    68dce24300              push 0043e2dc
'00717d21    56                      push esi
'00717d22    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00717d23    ff1580104000            call dword ptr [00401080]
'00717d29    8bb538feffff            mov esi, dword ptr [ebp+fffffe38]
'00717d2f    8d4db8                  lea ecx, dword ptr [ebp-48]

' *** Reference to "__vbaFreeObj"
'00717d32    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_61)
'00717d38    6685f6                  test esi, esi
'00717d3b    0f840a080000            je 71854b
'00717d41    668b55e8                mov dx, word ptr [ebp-18]
'00717d45    83ec10                  sub esp, 10
'00717d48    8bfc                    mov edi, esp
'00717d4a    b903000000              mov ecx, 00000003
'00717d4f    890f                    mov dword ptr [edi], ecx
'00717d51    8b8d50ffffff            mov ecx, dword ptr [ebp+ffffff50]
'00717d57    894f04                  mov dword ptr [edi+04], ecx
'00717d5a    33c0                    xor eax, eax
'00717d5c    894708                  mov dword ptr [edi+08], eax
'00717d5f    8b8558ffffff            mov eax, dword ptr [ebp+ffffff58]
'00717d65    89470c                  mov dword ptr [edi+0c], eax
'00717d68    83ec10                  sub esp, 10
'00717d6b    66899534ffffff          mov word ptr [ebp+ffffff34], dx
'00717d72    8b8534ffffff            mov eax, dword ptr [ebp+ffffff34]
'00717d78    8bcc                    mov ecx, esp
'00717d7a    ba02000000              mov edx, 00000002
'00717d7f    8911                    mov dword ptr [ecx], edx
'00717d81    8b9530ffffff            mov edx, dword ptr [ebp+ffffff30]
'00717d87    895104                  mov dword ptr [ecx+04], edx
'00717d8a    8b9538ffffff            mov edx, dword ptr [ebp+ffffff38]
'00717d90    894108                  mov dword ptr [ecx+08], eax
'00717d93    89510c                  mov dword ptr [ecx+0c], edx
'00717d96    8b9510ffffff            mov edx, dword ptr [ebp+ffffff10]
'00717d9c    83ec10                  sub esp, 10
'00717d9f    8bcc                    mov ecx, esp
'00717da1    b802000000              mov eax, 00000002
'00717da6    8901                    mov dword ptr [ecx], eax
'00717da8    8b8518ffffff            mov eax, dword ptr [ebp+ffffff18]
'00717dae    895104                  mov dword ptr [ecx+04], edx
'00717db1    6a03                    push 03
'00717db3    33f6                    xor esi, esi
var_num8 = Empty
'00717db5    897108                  mov dword ptr [ecx+08], esi
'00717db8    689e000000              push 0000009e
'00717dbd    89410c                  mov dword ptr [ecx+0c], eax
'00717dc0    8b0b                    mov ecx, dword ptr [ebx]
'00717dc2    53                      push ebx
'00717dc3    ff9110030000            call dword ptr [ecx+00000310]
'00717dc9    50                      push eax
'00717dca    8d55b8                  lea edx, dword ptr [ebp-48]
'00717dcd    52                      push edx

' *** Reference to "__vbaObjSet"
'00717dce    ff15b4104000            call dword ptr [004010b4]
Set var_61 = Nothing
'00717dd4    50                      push eax
'00717dd5    8d459c                  lea eax, dword ptr [ebp-64]
'00717dd8    50                      push eax

' *** Reference to "__vbaLateIdCallLd"
'00717dd9    ff157c114000            call dword ptr [0040117c]
var_51 = var_61.UNK_-256 - 12_158
'00717ddf    83c440                  add esp, 40
'00717de2    50                      push eax

' *** Reference to "__vbaStrVarMove"
'00717de3    ff153c104000            call dword ptr [0040103c]

' *** Reference to "__vbaStrMove"
'00717de9    8b3dd0124000            mov edi, dword ptr [004012d0]
'00717def    8bd0                    mov edx, eax
'00717df1    8d4dcc                  lea ecx, dword ptr [ebp-34]
'00717df4    ffd7                    call edi
'00717df6    8d4dcc                  lea ecx, dword ptr [ebp-34]
'00717df9    51                      push ecx
'00717dfa    e8f1bdddff              call 4f3bf0
Call sub_4F3BF0()
'00717dff    8bd0                    mov edx, eax
'00717e01    8d4dbc                  lea ecx, dword ptr [ebp-44]
'00717e04    ffd7                    call edi
'00717e06    8b55bc                  mov edx, dword ptr [ebp-44]
'00717e09    899544fdffff            mov dword ptr [ebp+fffffd44], edx
'00717e0f    8b5338                  mov edx, dword ptr [ebx+38]
'00717e12    c785f4feffff02000000    mov dword ptr [ebp+fffffef4], 00000002
'00717e1c    c785ecfeffff03000000    mov dword ptr [ebp+fffffeec], 00000003
'00717e26    8975bc                  mov dword ptr [ebp-44], esi
'00717e29    8b32                    mov esi, dword ptr [edx]
'00717e2b    8d55b4                  lea edx, dword ptr [ebp-4c]
'00717e2e    52                      push edx
'00717e2f    83ec10                  sub esp, 10
'00717e32    b90a000000              mov ecx, 0000000a
'00717e37    8bd4                    mov edx, esp
'00717e39    890a                    mov dword ptr [edx], ecx
'00717e3b    898ddcfeffff            mov dword ptr [ebp+fffffedc], ecx
'00717e41    8b8dd0feffff            mov ecx, dword ptr [ebp+fffffed0]
'00717e47    b804000280              mov eax, 80020004
'00717e4c    894a04                  mov dword ptr [edx+04], ecx
'00717e4f    894208                  mov dword ptr [edx+08], eax
'00717e52    8985e4feffff            mov dword ptr [ebp+fffffee4], eax
'00717e58    8b85d8feffff            mov eax, dword ptr [ebp+fffffed8]
'00717e5e    83ec10                  sub esp, 10
'00717e61    89420c                  mov dword ptr [edx+0c], eax
'00717e64    8bcc                    mov ecx, esp
'00717e66    8b95dcfeffff            mov edx, dword ptr [ebp+fffffedc]
'00717e6c    8b85e0feffff            mov eax, dword ptr [ebp+fffffee0]
'00717e72    8911                    mov dword ptr [ecx], edx
'00717e74    8b95e4feffff            mov edx, dword ptr [ebp+fffffee4]
'00717e7a    894104                  mov dword ptr [ecx+04], eax
'00717e7d    8b85e8feffff            mov eax, dword ptr [ebp+fffffee8]
'00717e83    895108                  mov dword ptr [ecx+08], edx
'00717e86    8b95ecfeffff            mov edx, dword ptr [ebp+fffffeec]
'00717e8c    89410c                  mov dword ptr [ecx+0c], eax
'00717e8f    8b85f0feffff            mov eax, dword ptr [ebp+fffffef0]
'00717e95    83ec10                  sub esp, 10
'00717e98    8bcc                    mov ecx, esp
'00717e9a    8911                    mov dword ptr [ecx], edx
'00717e9c    8b95f4feffff            mov edx, dword ptr [ebp+fffffef4]
'00717ea2    894104                  mov dword ptr [ecx+04], eax
'00717ea5    8b85f8feffff            mov eax, dword ptr [ebp+fffffef8]
'00717eab    895108                  mov dword ptr [ecx+08], edx
'00717eae    8b9544fdffff            mov edx, dword ptr [ebp+fffffd44]
'00717eb4    89410c                  mov dword ptr [ecx+0c], eax
'00717eb7    68706b4500              push 00456b70
'00717ebc    8d4dc8                  lea ecx, dword ptr [ebp-38]
'00717ebf    ffd7                    call edi
'00717ec1    50                      push eax

' *** Reference to "__vbaStrCat"
'00717ec2    ff1570104000            call dword ptr [00401070]
var_2583 = ("select * from equipement where personnage='") & (var_51)
'00717ec8    8bd0                    mov edx, eax
'00717eca    8d4dc4                  lea ecx, dword ptr [ebp-3c]
'00717ecd    ffd7                    call edi
'00717ecf    50                      push eax
'00717ed0    6854a44300              push 0043a454

' *** Reference to "__vbaStrCat"
'00717ed5    ff1570104000            call dword ptr [00401070]
var_875 = (var_2583) & ("'")
'00717edb    8bd0                    mov edx, eax
'00717edd    8d4dc0                  lea ecx, dword ptr [ebp-40]
'00717ee0    ffd7                    call edi
'00717ee2    8b4b38                  mov ecx, dword ptr [ebx+38]
'00717ee5    50                      push eax
'00717ee6    51                      push ecx
'00717ee7    ff96bc000000            call dword ptr [esi+000000bc]
'00717eed    dbe2                    fnclex
'00717eef    85c0                    test ax, ax
'00717ef1    7d15                    jge 717f08
'00717ef3    8b5338                  mov edx, dword ptr [ebx+38]
'00717ef6    68bc000000              push 000000bc
'00717efb    68301f4300              push 00431f30
'00717f00    52                      push edx
'00717f01    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00717f02    ff1580104000            call dword ptr [00401080]
'00717f08    8b45b4                  mov eax, dword ptr [ebp-4c]
'00717f0b    50                      push eax
'00717f0c    8d45d8                  lea eax, dword ptr [ebp-28]
'00717f0f    50                      push eax
'00717f10    c745b400000000          mov dword ptr [ebp-4c], 00000000

' *** Reference to "__vbaObjSet"
'00717f17    ff15b4104000            call dword ptr [004010b4]
Set var_45 = Nothing
'00717f1d    8d4dbc                  lea ecx, dword ptr [ebp-44]
'00717f20    51                      push ecx
'00717f21    8d55c0                  lea edx, dword ptr [ebp-40]
'00717f24    52                      push edx
'00717f25    8d45c4                  lea eax, dword ptr [ebp-3c]
'00717f28    50                      push eax
'00717f29    8d4dc8                  lea ecx, dword ptr [ebp-38]
'00717f2c    51                      push ecx
'00717f2d    8d55cc                  lea edx, dword ptr [ebp-34]
'00717f30    52                      push edx
'00717f31    6a05                    push 05

' *** Reference to "__vbaFreeStrList"
'00717f33    ff155c124000            call dword ptr [0040125c]
'#Cleanup( -4668, -4668, -4672, -4676, 0)
'00717f39    83c418                  add esp, 18
'00717f3c    8d4db8                  lea ecx, dword ptr [ebp-48]

' *** Reference to "__vbaFreeObj"
'00717f3f    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_61)
'00717f45    8d4d9c                  lea ecx, dword ptr [ebp-64]

' *** Reference to "__vbaFreeVar"
'00717f48    ff1528104000            call dword ptr [00401028]
'#Cleanup(var_51)
'00717f4e    8d75b8                  lea esi, dword ptr [ebp-48]
'00717f51    56                      push esi
'00717f52    83ec10                  sub esp, 10
'00717f55    8bf4                    mov esi, esp
'00717f57    b90a000000              mov ecx, 0000000a
'00717f5c    890e                    mov dword ptr [esi], ecx
'00717f5e    898d3cffffff            mov dword ptr [ebp+ffffff3c], ecx
'00717f64    8b8d30ffffff            mov ecx, dword ptr [ebp+ffffff30]
'00717f6a    894e04                  mov dword ptr [esi+04], ecx
'00717f6d    b804000280              mov eax, 80020004
'00717f72    894608                  mov dword ptr [esi+08], eax
'00717f75    898534ffffff            mov dword ptr [ebp+ffffff34], eax
'00717f7b    898544ffffff            mov dword ptr [ebp+ffffff44], eax
'00717f81    8b8538ffffff            mov eax, dword ptr [ebp+ffffff38]
'00717f87    89460c                  mov dword ptr [esi+0c], eax
'00717f8a    8b853cffffff            mov eax, dword ptr [ebp+ffffff3c]
'00717f90    8b1548b07200            mov edx, dword ptr [0072b048]
'00717f96    8b12                    mov edx, dword ptr [edx]
'00717f98    83ec10                  sub esp, 10
'00717f9b    8bcc                    mov ecx, esp
'00717f9d    8901                    mov dword ptr [ecx], eax
'00717f9f    8b8540ffffff            mov eax, dword ptr [ebp+ffffff40]
'00717fa5    894104                  mov dword ptr [ecx+04], eax
'00717fa8    8b8544ffffff            mov eax, dword ptr [ebp+ffffff44]
'00717fae    894108                  mov dword ptr [ecx+08], eax
'00717fb1    8b8548ffffff            mov eax, dword ptr [ebp+ffffff48]
'00717fb7    89410c                  mov dword ptr [ecx+0c], eax
'00717fba    83ec10                  sub esp, 10
'00717fbd    8bcc                    mov ecx, esp
'00717fbf    b803000000              mov eax, 00000003
'00717fc4    8901                    mov dword ptr [ecx], eax
'00717fc6    8b8550ffffff            mov eax, dword ptr [ebp+ffffff50]
'00717fcc    894104                  mov dword ptr [ecx+04], eax
'00717fcf    c78554ffffff01000000    mov dword ptr [ebp+ffffff54], 00000001
'00717fd9    8b8554ffffff            mov eax, dword ptr [ebp+ffffff54]
'00717fdf    894108                  mov dword ptr [ecx+08], eax
'00717fe2    8b8558ffffff            mov eax, dword ptr [ebp+ffffff58]
'00717fe8    89410c                  mov dword ptr [ecx+0c], eax
'00717feb    8b0d48b07200            mov ecx, dword ptr [0072b048]
'00717ff1    6894aa4300              push 0043aa94
'00717ff6    51                      push ecx
'00717ff7    ff92bc000000            call dword ptr [edx+000000bc]
'00717ffd    dbe2                    fnclex
'00717fff    85c0                    test ax, ax
'00718001    7d18                    jge 71801b
'00718003    8b1548b07200            mov edx, dword ptr [0072b048]
'00718009    68bc000000              push 000000bc
'0071800e    68301f4300              push 00431f30
'00718013    52                      push edx
'00718014    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00718015    ff1580104000            call dword ptr [00401080]
'0071801b    8b45b8                  mov eax, dword ptr [ebp-48]
'0071801e    50                      push eax
'0071801f    8d45dc                  lea eax, dword ptr [ebp-24]
'00718022    50                      push eax
'00718023    c745b800000000          mov dword ptr [ebp-48], 00000000

' *** Reference to "__vbaObjSet"
'0071802a    ff15b4104000            call dword ptr [004010b4]
Set var_39 = Nothing
'00718030    8b45d8                  mov eax, dword ptr [ebp-28]
'00718033    8b08                    mov ecx, dword ptr [eax]
'00718035    8d9538feffff            lea edx, dword ptr [ebp+fffffe38]
'0071803b    52                      push edx
'0071803c    50                      push eax
'0071803d    ff5134                  call dword ptr [ecx+34]
'00718040    dbe2                    fnclex
'00718042    85c0                    test ax, ax
'00718044    7d16                    jge 71805c
'00718046    8b4dd8                  mov ecx, dword ptr [ebp-28]

' *** Reference to "__vbaHresultCheckObj"
'00718049    8b3580104000            mov esi, dword ptr [00401080]
'0071804f    6a34                    push 34
'00718051    6830314300              push 00433130
'00718056    51                      push ecx
'00718057    50                      push eax
'00718058    ffd6                    call esi
'0071805a    eb06                    jmp 718062

' *** Reference to "__vbaHresultCheckObj"
'0071805c    8b3580104000            mov esi, dword ptr [00401080]
'00718062    6683bd38feffff00        cmp word ptr [ebp+fffffe38], 00
'0071806a    8b45dc                  mov eax, dword ptr [ebp-24]
'0071806d    0f8593040000            jne 718506

Do While (0 = 0)
'00718073    8b10                    mov edx, dword ptr [eax]
'00718075    50                      push eax
'00718076    ff92c0000000            call dword ptr [edx+000000c0]
'0071807c    dbe2                    fnclex
'0071807e    85c0                    test ax, ax
'00718080    7d11                    jge 718093
'00718082    8b4ddc                  mov ecx, dword ptr [ebp-24]
'00718085    68c0000000              push 000000c0
'0071808a    6830314300              push 00433130
'0071808f    51                      push ecx
'00718090    50                      push eax
'00718091    ffd6                    call esi
'00718093    33f6                    xor esi, esi
    var_num8 = Empty
'00718095    8975e0                  mov dword ptr [ebp-20], esi
'00718098    b807000000              mov eax, 00000007
'0071809d    663bf0                  cmp si, ax
'007180a0    0f8f05040000            jg 7184ab
    
    Do While (    __vbaHresultCheckObj <= 7)
'007180a6    8b45d8                  mov eax, dword ptr [ebp-28]
'007180a9    8b10                    mov edx, dword ptr [eax]
'007180ab    8d4db8                  lea ecx, dword ptr [ebp-48]
'007180ae    51                      push ecx
'007180af    50                      push eax
'007180b0    ff92b4000000            call dword ptr [edx+000000b4]
'007180b6    dbe2                    fnclex
'007180b8    85c0                    test ax, ax
'007180ba    7d15                    jge 7180d1
'007180bc    8b55d8                  mov edx, dword ptr [ebp-28]
'007180bf    68b4000000              push 000000b4
'007180c4    6830314300              push 00433130
'007180c9    52                      push edx
'007180ca    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'007180cb    ff1580104000            call dword ptr [00401080]
'007180d1    8b45b8                  mov eax, dword ptr [ebp-48]
'007180d4    8b10                    mov edx, dword ptr [eax]
'007180d6    6689b554ffffff          mov word ptr [ebp+ffffff54], si
'007180dd    8d75b4                  lea esi, dword ptr [ebp-4c]
'007180e0    56                      push esi
'007180e1    83ec10                  sub esp, 10
'007180e4    8bf4                    mov esi, esp
'007180e6    b902000000              mov ecx, 00000002
'007180eb    890e                    mov dword ptr [esi], ecx
'007180ed    8b8d50ffffff            mov ecx, dword ptr [ebp+ffffff50]
'007180f3    894e04                  mov dword ptr [esi+04], ecx
'007180f6    8b8d54ffffff            mov ecx, dword ptr [ebp+ffffff54]
'007180fc    894e08                  mov dword ptr [esi+08], ecx
'007180ff    8b8d58ffffff            mov ecx, dword ptr [ebp+ffffff58]
'00718105    50                      push eax
'00718106    898530feffff            mov dword ptr [ebp+fffffe30], eax
'0071810c    894e0c                  mov dword ptr [esi+0c], ecx
'0071810f    ff5230                  call dword ptr [edx+30]
'00718112    dbe2                    fnclex
'00718114    85c0                    test ax, ax
'00718116    7d15                    jge 71812d
'00718118    8b9530feffff            mov edx, dword ptr [ebp+fffffe30]
'0071811e    6a30                    push 30
'00718120    68d8304300              push 004330d8
'00718125    52                      push edx
'00718126    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00718127    ff1580104000            call dword ptr [00401080]
'0071812d    8b45b4                  mov eax, dword ptr [ebp-4c]
'00718130    8945a4                  mov dword ptr [ebp-5c], eax
'00718133    8d459c                  lea eax, dword ptr [ebp-64]
'00718136    50                      push eax
'00718137    c745b400000000          mov dword ptr [ebp-4c], 00000000
'0071813e    c7459c09000000          mov dword ptr [ebp-64], 00000009

' *** Reference to "rtcIsNull"
'00718145    ff1540114000            call dword ptr [00401140]
'0071814b    33f6                    xor esi, esi
    var_num8 = Empty
'0071814d    668bf0                  mov si, ax
'00718150    8d4db8                  lea ecx, dword ptr [ebp-48]
'00718153    f7d6                    not esi

' *** Reference to "__vbaFreeObj"
'00718155    ff1508134000            call dword ptr [00401308]
    '#Cleanup(var_61)
'0071815b    8d4d9c                  lea ecx, dword ptr [ebp-64]

' *** Reference to "__vbaFreeVar"
'0071815e    ff1528104000            call dword ptr [00401028]
    '#Cleanup(var_51)
'00718164    6685f6                  test esi, esi
'00718167    0f8425030000            je 718492
    
    If (    Not (IsNull(var_62))) Then
'0071816d    8b75e0                  mov esi, dword ptr [ebp-20]
'00718170    6683fe01                cmp si, 01
'00718174    8b45dc                  mov eax, dword ptr [ebp-24]
'00718177    8b08                    mov ecx, dword ptr [eax]
'00718179    0f858e010000            jne 71830d
    
    If (    __vbaHresultCheckObj = 1) Then
'0071817f    8d55b4                  lea edx, dword ptr [ebp-4c]
'00718182    52                      push edx
'00718183    50                      push eax
'00718184    ff91b4000000            call dword ptr [ecx+000000b4]
'0071818a    dbe2                    fnclex
'0071818c    85c0                    test ax, ax
'0071818e    7d15                    jge 7181a5
'00718190    8b4ddc                  mov ecx, dword ptr [ebp-24]
'00718193    68b4000000              push 000000b4
'00718198    6830314300              push 00433130
'0071819d    51                      push ecx
'0071819e    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0071819f    ff1580104000            call dword ptr [00401080]
'007181a5    8b45b4                  mov eax, dword ptr [ebp-4c]
'007181a8    8d7db0                  lea edi, dword ptr [ebp-50]
'007181ab    57                      push edi
'007181ac    83ec10                  sub esp, 10
'007181af    b902000000              mov ecx, 00000002
'007181b4    898decfeffff            mov dword ptr [ebp+fffffeec], ecx
'007181ba    8bfc                    mov edi, esp
'007181bc    890f                    mov dword ptr [edi], ecx
'007181be    8b8df0feffff            mov ecx, dword ptr [ebp+fffffef0]
'007181c4    894f04                  mov dword ptr [edi+04], ecx
'007181c7    66c785f4feffff0100      mov word ptr [ebp+fffffef4], 0001
'007181d0    8b8df4feffff            mov ecx, dword ptr [ebp+fffffef4]
'007181d6    8b10                    mov edx, dword ptr [eax]
'007181d8    894f08                  mov dword ptr [edi+08], ecx
'007181db    8b8df8feffff            mov ecx, dword ptr [ebp+fffffef8]
'007181e1    50                      push eax
'007181e2    8bf0                    mov esi, eax
'007181e4    894f0c                  mov dword ptr [edi+0c], ecx
'007181e7    ff5230                  call dword ptr [edx+30]
'007181ea    dbe2                    fnclex
'007181ec    85c0                    test ax, ax
'007181ee    7d0f                    jge 7181ff
'007181f0    6a30                    push 30
'007181f2    68d8304300              push 004330d8
'007181f7    56                      push esi
'007181f8    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'007181f9    ff1580104000            call dword ptr [00401080]
'007181ff    668b55e8                mov dx, word ptr [ebp-18]
'00718203    83ec10                  sub esp, 10
'00718206    8bdc                    mov ebx, esp
'00718208    8b75b0                  mov esi, dword ptr [ebp-50]
'0071820b    33c0                    xor eax, eax
    var_num1 = Empty
'0071820d    8b3e                    mov edi, dword ptr [esi]
'0071820f    b903000000              mov ecx, 00000003
'00718214    890b                    mov dword ptr [ebx], ecx
'00718216    8b8d50ffffff            mov ecx, dword ptr [ebp+ffffff50]
'0071821c    894b04                  mov dword ptr [ebx+04], ecx
'0071821f    894308                  mov dword ptr [ebx+08], eax
'00718222    898554ffffff            mov dword ptr [ebp+ffffff54], eax
'00718228    8b8558ffffff            mov eax, dword ptr [ebp+ffffff58]
'0071822e    89430c                  mov dword ptr [ebx+0c], eax
'00718231    8b5d08                  mov ebx, dword ptr [ebp+08]
'00718234    83ec10                  sub esp, 10
'00718237    66899534ffffff          mov word ptr [ebp+ffffff34], dx
'0071823e    8b8534ffffff            mov eax, dword ptr [ebp+ffffff34]
'00718244    8bcc                    mov ecx, esp
'00718246    ba02000000              mov edx, 00000002
'0071824b    8911                    mov dword ptr [ecx], edx
'0071824d    8b9530ffffff            mov edx, dword ptr [ebp+ffffff30]
'00718253    895104                  mov dword ptr [ecx+04], edx
'00718256    8b9538ffffff            mov edx, dword ptr [ebp+ffffff38]
'0071825c    894108                  mov dword ptr [ecx+08], eax
'0071825f    89510c                  mov dword ptr [ecx+0c], edx
'00718262    8b9510ffffff            mov edx, dword ptr [ebp+ffffff10]
'00718268    83ec10                  sub esp, 10
'0071826b    8bcc                    mov ecx, esp
'0071826d    b802000000              mov eax, 00000002
'00718272    8901                    mov dword ptr [ecx], eax
'00718274    895104                  mov dword ptr [ecx+04], edx
'00718277    b801000000              mov eax, 00000001
'0071827c    894108                  mov dword ptr [ecx+08], eax
'0071827f    8b8518ffffff            mov eax, dword ptr [ebp+ffffff18]
'00718285    6a03                    push 03
'00718287    689e000000              push 0000009e
'0071828c    89410c                  mov dword ptr [ecx+0c], eax
'0071828f    8b0b                    mov ecx, dword ptr [ebx]
'00718291    53                      push ebx
'00718292    ff9110030000            call dword ptr [ecx+00000310]
'00718298    50                      push eax
'00718299    8d55b8                  lea edx, dword ptr [ebp-48]
'0071829c    52                      push edx

' *** Reference to "__vbaObjSet"
'0071829d    ff15b4104000            call dword ptr [004010b4]
    Set var_61 = Nothing
'007182a3    50                      push eax
'007182a4    8d459c                  lea eax, dword ptr [ebp-64]
'007182a7    50                      push eax

' *** Reference to "__vbaLateIdCallLd"
'007182a8    ff157c114000            call dword ptr [0040117c]
    var_51 = var_61.UNK_-256 - 12_158
'007182ae    8b10                    mov edx, dword ptr [eax]
'007182b0    83c430                  add esp, 30
'007182b3    8bcc                    mov ecx, esp
'007182b5    8911                    mov dword ptr [ecx], edx
'007182b7    8b5004                  mov edx, dword ptr [eax+04]
'007182ba    895104                  mov dword ptr [ecx+04], edx
'007182bd    8b5008                  mov edx, dword ptr [eax+08]
'007182c0    8b400c                  mov eax, dword ptr [eax+0c]
'007182c3    895108                  mov dword ptr [ecx+08], edx
'007182c6    56                      push esi
'007182c7    89410c                  mov dword ptr [ecx+0c], eax
'007182ca    ff5748                  call dword ptr [edi+48]
'007182cd    dbe2                    fnclex
'007182cf    85c0                    test ax, ax
'007182d1    7d0f                    jge 7182e2
'007182d3    6a48                    push 48
'007182d5    6880304300              push 00433080
'007182da    56                      push esi
'007182db    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'007182dc    ff1580104000            call dword ptr [00401080]
'007182e2    8d4db0                  lea ecx, dword ptr [ebp-50]
'007182e5    51                      push ecx
'007182e6    8d55b4                  lea edx, dword ptr [ebp-4c]
'007182e9    52                      push edx
'007182ea    8d45b8                  lea eax, dword ptr [ebp-48]
'007182ed    50                      push eax
'007182ee    6a03                    push 03

' *** Reference to "__vbaFreeObjList"
'007182f0    ff154c104000            call dword ptr [0040104c]
    '#Cleanup( var_61, var_62, var_6)
'007182f6    83c410                  add esp, 10
'007182f9    8d4d9c                  lea ecx, dword ptr [ebp-64]

' *** Reference to "__vbaFreeVar"
'007182fc    ff1528104000            call dword ptr [00401028]
    '#Cleanup(var_51)

' *** Reference to "__vbaStrMove"
'00718302    8b3dd0124000            mov edi, dword ptr [004012d0]
'00718308    e985010000              jmp 718492
    
End If
'0071830d    8d55b0                  lea edx, dword ptr [ebp-50]
'00718310    52                      push edx
'00718311    50                      push eax
'00718312    ff91b4000000            call dword ptr [ecx+000000b4]
'00718318    dbe2                    fnclex
'0071831a    85c0                    test ax, ax
'0071831c    7d15                    jge 718333
'0071831e    8b4ddc                  mov ecx, dword ptr [ebp-24]
'00718321    68b4000000              push 000000b4
'00718326    6830314300              push 00433130
'0071832b    51                      push ecx
'0071832c    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0071832d    ff1580104000            call dword ptr [00401080]
'00718333    8b45b0                  mov eax, dword ptr [ebp-50]
'00718336    8b10                    mov edx, dword ptr [eax]
'00718338    6689b544ffffff          mov word ptr [ebp+ffffff44], si
'0071833f    8d75ac                  lea esi, dword ptr [ebp-54]
'00718342    56                      push esi
'00718343    83ec10                  sub esp, 10
'00718346    8bf4                    mov esi, esp
'00718348    b902000000              mov ecx, 00000002
'0071834d    890e                    mov dword ptr [esi], ecx
'0071834f    8b8d40ffffff            mov ecx, dword ptr [ebp+ffffff40]
'00718355    894e04                  mov dword ptr [esi+04], ecx
'00718358    8b8d44ffffff            mov ecx, dword ptr [ebp+ffffff44]
'0071835e    894e08                  mov dword ptr [esi+08], ecx
'00718361    8b8d48ffffff            mov ecx, dword ptr [ebp+ffffff48]
'00718367    50                      push eax
'00718368    898524feffff            mov dword ptr [ebp+fffffe24], eax
'0071836e    894e0c                  mov dword ptr [esi+0c], ecx
'00718371    ff5230                  call dword ptr [edx+30]
'00718374    dbe2                    fnclex
'00718376    85c0                    test ax, ax
'00718378    7d15                    jge 71838f
'0071837a    8b9524feffff            mov edx, dword ptr [ebp+fffffe24]
'00718380    6a30                    push 30
'00718382    68d8304300              push 004330d8
'00718387    52                      push edx
'00718388    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00718389    ff1580104000            call dword ptr [00401080]
'0071838f    8b45ac                  mov eax, dword ptr [ebp-54]
'00718392    8d55b8                  lea edx, dword ptr [ebp-48]
'00718395    89851cfeffff            mov dword ptr [ebp+fffffe1c], eax
'0071839b    8b45d8                  mov eax, dword ptr [ebp-28]
'0071839e    8b08                    mov ecx, dword ptr [eax]
'007183a0    52                      push edx
'007183a1    50                      push eax
'007183a2    ff91b4000000            call dword ptr [ecx+000000b4]
'007183a8    dbe2                    fnclex
'007183aa    85c0                    test ax, ax
'007183ac    7d15                    jge 7183c3
'007183ae    8b4dd8                  mov ecx, dword ptr [ebp-28]
'007183b1    68b4000000              push 000000b4
'007183b6    6830314300              push 00433130
'007183bb    51                      push ecx
'007183bc    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'007183bd    ff1580104000            call dword ptr [00401080]
'007183c3    668b55e0                mov dx, word ptr [ebp-20]
'007183c7    8b45b8                  mov eax, dword ptr [ebp-48]
'007183ca    8d75b4                  lea esi, dword ptr [ebp-4c]
'007183cd    56                      push esi
'007183ce    83ec10                  sub esp, 10
'007183d1    8bf4                    mov esi, esp
'007183d3    b902000000              mov ecx, 00000002
'007183d8    890e                    mov dword ptr [esi], ecx
'007183da    8b8d50ffffff            mov ecx, dword ptr [ebp+ffffff50]
'007183e0    894e04                  mov dword ptr [esi+04], ecx
'007183e3    66899554ffffff          mov word ptr [ebp+ffffff54], dx
'007183ea    8b8d54ffffff            mov ecx, dword ptr [ebp+ffffff54]
'007183f0    8b10                    mov edx, dword ptr [eax]
'007183f2    894e08                  mov dword ptr [esi+08], ecx
'007183f5    8b8d58ffffff            mov ecx, dword ptr [ebp+ffffff58]
'007183fb    50                      push eax
'007183fc    898530feffff            mov dword ptr [ebp+fffffe30], eax
'00718402    894e0c                  mov dword ptr [esi+0c], ecx
'00718405    ff5230                  call dword ptr [edx+30]
'00718408    dbe2                    fnclex
'0071840a    85c0                    test ax, ax
'0071840c    7d15                    jge 718423
'0071840e    8b9530feffff            mov edx, dword ptr [ebp+fffffe30]
'00718414    6a30                    push 30
'00718416    68d8304300              push 004330d8
'0071841b    52                      push edx
'0071841c    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0071841d    ff1580104000            call dword ptr [00401080]
'00718423    8b45b4                  mov eax, dword ptr [ebp-4c]
'00718426    8bb51cfeffff            mov esi, dword ptr [ebp+fffffe1c]
'0071842c    8945a4                  mov dword ptr [ebp-5c], eax
'0071842f    83ec10                  sub esp, 10
'00718432    b809000000              mov eax, 00000009
'00718437    89459c                  mov dword ptr [ebp-64], eax
'0071843a    8bd4                    mov edx, esp
'0071843c    8902                    mov dword ptr [edx], eax
'0071843e    8b45a0                  mov eax, dword ptr [ebp-60]
'00718441    894204                  mov dword ptr [edx+04], eax
'00718444    8b45a4                  mov eax, dword ptr [ebp-5c]
'00718447    894208                  mov dword ptr [edx+08], eax
'0071844a    8b45a8                  mov eax, dword ptr [ebp-58]
'0071844d    c745b400000000          mov dword ptr [ebp-4c], 00000000
'00718454    8b0e                    mov ecx, dword ptr [esi]
'00718456    56                      push esi
'00718457    89420c                  mov dword ptr [edx+0c], eax
'0071845a    ff5148                  call dword ptr [ecx+48]
'0071845d    dbe2                    fnclex
'0071845f    85c0                    test ax, ax
'00718461    7d0f                    jge 718472
'00718463    6a48                    push 48
'00718465    6880304300              push 00433080
'0071846a    56                      push esi
'0071846b    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0071846c    ff1580104000            call dword ptr [00401080]
'00718472    8d4dac                  lea ecx, dword ptr [ebp-54]
'00718475    51                      push ecx
'00718476    8d55b0                  lea edx, dword ptr [ebp-50]
'00718479    52                      push edx
'0071847a    8d45b8                  lea eax, dword ptr [ebp-48]
'0071847d    50                      push eax
'0071847e    6a03                    push 03

' *** Reference to "__vbaFreeObjList"
'00718480    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_61, var_6, var_50)
'00718486    83c410                  add esp, 10
'00718489    8d4d9c                  lea ecx, dword ptr [ebp-64]

' *** Reference to "__vbaFreeVar"
'0071848c    ff1528104000            call dword ptr [00401028]
'#Cleanup(var_51)
'ERROR: Two many next close:
End If
'00718492    b801000000              mov eax, 00000001
'00718497    660345e0                add ax, word ptr [ebp-20]
var_num1 = 1 + __vbaHresultCheckObj
'0071849b    0f80d2230000            jo 71a873
'007184a1    8945e0                  mov dword ptr [ebp-20], eax
'007184a4    8bf0                    mov esi, eax
'007184a6    e9edfbffff              jmp 718098

'ERROR: Two many next close:
Loop
'007184ab    8b45dc                  mov eax, dword ptr [ebp-24]
'007184ae    8b08                    mov ecx, dword ptr [eax]
'007184b0    6a00                    push 00
'007184b2    6a01                    push 01
'007184b4    50                      push eax
'007184b5    ff9164010000            call dword ptr [ecx+00000164]
'007184bb    dbe2                    fnclex
'007184bd    85c0                    test ax, ax
'007184bf    7d15                    jge 7184d6
'007184c1    8b55dc                  mov edx, dword ptr [ebp-24]
'007184c4    6864010000              push 00000164
'007184c9    6830314300              push 00433130
'007184ce    52                      push edx
'007184cf    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'007184d0    ff1580104000            call dword ptr [00401080]
'007184d6    8b45d8                  mov eax, dword ptr [ebp-28]
'007184d9    8b08                    mov ecx, dword ptr [eax]
'007184db    50                      push eax
'007184dc    ff91ec000000            call dword ptr [ecx+000000ec]
'007184e2    dbe2                    fnclex
'007184e4    85c0                    test ax, ax
'007184e6    0f8d44fbffff            jge 718030
'007184ec    8b55d8                  mov edx, dword ptr [ebp-28]
'007184ef    68ec000000              push 000000ec
'007184f4    6830314300              push 00433130
'007184f9    52                      push edx
'007184fa    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'007184fb    ff1580104000            call dword ptr [00401080]
'00718501    e92afbffff              jmp 718030

'ERROR: Two many next close:
Loop
'00718506    8b08                    mov ecx, dword ptr [eax]
'00718508    50                      push eax
'00718509    ff91c4000000            call dword ptr [ecx+000000c4]
'0071850f    dbe2                    fnclex
'00718511    85c0                    test ax, ax
'00718513    7d11                    jge 718526
'00718515    8b55dc                  mov edx, dword ptr [ebp-24]
'00718518    68c4000000              push 000000c4
'0071851d    6830314300              push 00433130
'00718522    52                      push edx
'00718523    50                      push eax
'00718524    ffd6                    call esi
'00718526    8b45d8                  mov eax, dword ptr [ebp-28]
'00718529    8b08                    mov ecx, dword ptr [eax]
'0071852b    50                      push eax
'0071852c    ff91c4000000            call dword ptr [ecx+000000c4]
'00718532    dbe2                    fnclex
'00718534    85c0                    test ax, ax
'00718536    7d19                    jge 718551
'00718538    8b55d8                  mov edx, dword ptr [ebp-28]
'0071853b    68c4000000              push 000000c4
'00718540    6830314300              push 00433130
'00718545    52                      push edx
'00718546    50                      push eax
'00718547    ffd6                    call esi
'00718549    eb06                    jmp 718551

' *** Reference to "__vbaStrMove"
'0071854b    8b3dd0124000            mov edi, dword ptr [004012d0]
'00718551    668b55e8                mov dx, word ptr [ebp-18]
'00718555    83ec10                  sub esp, 10
'00718558    8bf4                    mov esi, esp
'0071855a    33c0                    xor eax, eax
var_num1 = Empty
'0071855c    b903000000              mov ecx, 00000003
'00718561    890e                    mov dword ptr [esi], ecx
'00718563    8b8d50ffffff            mov ecx, dword ptr [ebp+ffffff50]
'00718569    894e04                  mov dword ptr [esi+04], ecx
'0071856c    894608                  mov dword ptr [esi+08], eax
'0071856f    8b8558ffffff            mov eax, dword ptr [ebp+ffffff58]
'00718575    89460c                  mov dword ptr [esi+0c], eax
'00718578    83ec10                  sub esp, 10
'0071857b    66899534ffffff          mov word ptr [ebp+ffffff34], dx
'00718582    8b8534ffffff            mov eax, dword ptr [ebp+ffffff34]
'00718588    8bcc                    mov ecx, esp
'0071858a    ba02000000              mov edx, 00000002
'0071858f    8911                    mov dword ptr [ecx], edx
'00718591    8b9530ffffff            mov edx, dword ptr [ebp+ffffff30]
'00718597    895104                  mov dword ptr [ecx+04], edx
'0071859a    8b9538ffffff            mov edx, dword ptr [ebp+ffffff38]
'007185a0    894108                  mov dword ptr [ecx+08], eax
'007185a3    89510c                  mov dword ptr [ecx+0c], edx
'007185a6    8b9510ffffff            mov edx, dword ptr [ebp+ffffff10]
'007185ac    83ec10                  sub esp, 10
'007185af    8bcc                    mov ecx, esp
'007185b1    b802000000              mov eax, 00000002
'007185b6    8901                    mov dword ptr [ecx], eax
'007185b8    895104                  mov dword ptr [ecx+04], edx
'007185bb    33c0                    xor eax, eax
var_num1 = Empty
'007185bd    894108                  mov dword ptr [ecx+08], eax
'007185c0    8b8518ffffff            mov eax, dword ptr [ebp+ffffff18]
'007185c6    6a03                    push 03
'007185c8    689e000000              push 0000009e
'007185cd    89410c                  mov dword ptr [ecx+0c], eax
'007185d0    8b0b                    mov ecx, dword ptr [ebx]
'007185d2    53                      push ebx
'007185d3    ff9110030000            call dword ptr [ecx+00000310]
'007185d9    50                      push eax
'007185da    8d55b8                  lea edx, dword ptr [ebp-48]
'007185dd    52                      push edx

' *** Reference to "__vbaObjSet"
'007185de    ff15b4104000            call dword ptr [004010b4]
Set var_61 = Nothing
'007185e4    50                      push eax
'007185e5    8d459c                  lea eax, dword ptr [ebp-64]
'007185e8    50                      push eax

' *** Reference to "__vbaLateIdCallLd"
'007185e9    ff157c114000            call dword ptr [0040117c]
var_51 = var_61.UNK_-256 - 12_158
'007185ef    83c440                  add esp, 40
'007185f2    50                      push eax

' *** Reference to "__vbaStrVarMove"
'007185f3    ff153c104000            call dword ptr [0040103c]
'007185f9    8bd0                    mov edx, eax
'007185fb    8d4dcc                  lea ecx, dword ptr [ebp-34]
'007185fe    ffd7                    call edi
'00718600    8d4dcc                  lea ecx, dword ptr [ebp-34]
'00718603    51                      push ecx
'00718604    e8e7b5ddff              call 4f3bf0
Call sub_4F3BF0()
'00718609    8bd0                    mov edx, eax
'0071860b    8d4dbc                  lea ecx, dword ptr [ebp-44]
'0071860e    ffd7                    call edi
'00718610    8b55bc                  mov edx, dword ptr [ebp-44]
'00718613    899530fdffff            mov dword ptr [ebp+fffffd30], edx
'00718619    8b5338                  mov edx, dword ptr [ebx+38]
'0071861c    c785f4feffff02000000    mov dword ptr [ebp+fffffef4], 00000002
'00718626    c785ecfeffff03000000    mov dword ptr [ebp+fffffeec], 00000003
'00718630    c745bc00000000          mov dword ptr [ebp-44], 00000000
'00718637    8b32                    mov esi, dword ptr [edx]
'00718639    8d55b4                  lea edx, dword ptr [ebp-4c]
'0071863c    52                      push edx
'0071863d    83ec10                  sub esp, 10
'00718640    b90a000000              mov ecx, 0000000a
'00718645    8bd4                    mov edx, esp
'00718647    890a                    mov dword ptr [edx], ecx
'00718649    898ddcfeffff            mov dword ptr [ebp+fffffedc], ecx
'0071864f    8b8dd0feffff            mov ecx, dword ptr [ebp+fffffed0]
'00718655    894a04                  mov dword ptr [edx+04], ecx
'00718658    b804000280              mov eax, 80020004
'0071865d    894208                  mov dword ptr [edx+08], eax
'00718660    8985e4feffff            mov dword ptr [ebp+fffffee4], eax
'00718666    8b85d8feffff            mov eax, dword ptr [ebp+fffffed8]
'0071866c    83ec10                  sub esp, 10
'0071866f    89420c                  mov dword ptr [edx+0c], eax
'00718672    8b95dcfeffff            mov edx, dword ptr [ebp+fffffedc]
'00718678    8bcc                    mov ecx, esp
'0071867a    8b85e0feffff            mov eax, dword ptr [ebp+fffffee0]
'00718680    8911                    mov dword ptr [ecx], edx
'00718682    8b95e4feffff            mov edx, dword ptr [ebp+fffffee4]
'00718688    894104                  mov dword ptr [ecx+04], eax
'0071868b    8b85e8feffff            mov eax, dword ptr [ebp+fffffee8]
'00718691    895108                  mov dword ptr [ecx+08], edx
'00718694    8b95ecfeffff            mov edx, dword ptr [ebp+fffffeec]
'0071869a    89410c                  mov dword ptr [ecx+0c], eax
'0071869d    8b85f0feffff            mov eax, dword ptr [ebp+fffffef0]
'007186a3    83ec10                  sub esp, 10
'007186a6    8bcc                    mov ecx, esp
'007186a8    8911                    mov dword ptr [ecx], edx
'007186aa    8b95f4feffff            mov edx, dword ptr [ebp+fffffef4]
'007186b0    894104                  mov dword ptr [ecx+04], eax
'007186b3    8b85f8feffff            mov eax, dword ptr [ebp+fffffef8]
'007186b9    895108                  mov dword ptr [ecx+08], edx
'007186bc    8b9530fdffff            mov edx, dword ptr [ebp+fffffd30]
'007186c2    89410c                  mov dword ptr [ecx+0c], eax
'007186c5    68046c4500              push 00456c04
'007186ca    8d4dc8                  lea ecx, dword ptr [ebp-38]
'007186cd    ffd7                    call edi
'007186cf    50                      push eax

' *** Reference to "__vbaStrCat"
'007186d0    ff1570104000            call dword ptr [00401070]
var_885 = ("select * from PersonnageDons where nom='") & (var_51)
'007186d6    8bd0                    mov edx, eax
'007186d8    8d4dc4                  lea ecx, dword ptr [ebp-3c]
'007186db    ffd7                    call edi
'007186dd    50                      push eax
'007186de    6854a44300              push 0043a454

' *** Reference to "__vbaStrCat"
'007186e3    ff1570104000            call dword ptr [00401070]
var_1470 = (var_885) & ("'")
'007186e9    8bd0                    mov edx, eax
'007186eb    8d4dc0                  lea ecx, dword ptr [ebp-40]
'007186ee    ffd7                    call edi
'007186f0    8b4b38                  mov ecx, dword ptr [ebx+38]
'007186f3    50                      push eax
'007186f4    51                      push ecx
'007186f5    ff96bc000000            call dword ptr [esi+000000bc]
'007186fb    dbe2                    fnclex
'007186fd    85c0                    test ax, ax
'007186ff    7d15                    jge 718716
'00718701    8b5338                  mov edx, dword ptr [ebx+38]
'00718704    68bc000000              push 000000bc
'00718709    68301f4300              push 00431f30
'0071870e    52                      push edx
'0071870f    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00718710    ff1580104000            call dword ptr [00401080]
'00718716    8b45b4                  mov eax, dword ptr [ebp-4c]
'00718719    50                      push eax
'0071871a    8d45d8                  lea eax, dword ptr [ebp-28]
'0071871d    50                      push eax
'0071871e    c745b400000000          mov dword ptr [ebp-4c], 00000000

' *** Reference to "__vbaObjSet"
'00718725    ff15b4104000            call dword ptr [004010b4]
Set var_45 = Nothing
'0071872b    8d4dbc                  lea ecx, dword ptr [ebp-44]
'0071872e    51                      push ecx
'0071872f    8d55c0                  lea edx, dword ptr [ebp-40]
'00718732    52                      push edx
'00718733    8d45c4                  lea eax, dword ptr [ebp-3c]
'00718736    50                      push eax
'00718737    8d4dc8                  lea ecx, dword ptr [ebp-38]
'0071873a    51                      push ecx
'0071873b    8d55cc                  lea edx, dword ptr [ebp-34]
'0071873e    52                      push edx
'0071873f    6a05                    push 05

' *** Reference to "__vbaFreeStrList"
'00718741    ff155c124000            call dword ptr [0040125c]
'#Cleanup( -4732, -4732, -4736, -4740, var_58)
'00718747    83c418                  add esp, 18
'0071874a    8d4db8                  lea ecx, dword ptr [ebp-48]

' *** Reference to "__vbaFreeObj"
'0071874d    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_61)
'00718753    8d4d9c                  lea ecx, dword ptr [ebp-64]

' *** Reference to "__vbaFreeVar"
'00718756    ff1528104000            call dword ptr [00401028]
'#Cleanup(var_51)
'0071875c    8d75b8                  lea esi, dword ptr [ebp-48]
'0071875f    56                      push esi
'00718760    83ec10                  sub esp, 10
'00718763    8bf4                    mov esi, esp
'00718765    b90a000000              mov ecx, 0000000a
'0071876a    890e                    mov dword ptr [esi], ecx
'0071876c    898d3cffffff            mov dword ptr [ebp+ffffff3c], ecx
'00718772    8b8d30ffffff            mov ecx, dword ptr [ebp+ffffff30]
'00718778    894e04                  mov dword ptr [esi+04], ecx
'0071877b    b804000280              mov eax, 80020004
'00718780    894608                  mov dword ptr [esi+08], eax
'00718783    898534ffffff            mov dword ptr [ebp+ffffff34], eax
'00718789    898544ffffff            mov dword ptr [ebp+ffffff44], eax
'0071878f    8b8538ffffff            mov eax, dword ptr [ebp+ffffff38]
'00718795    89460c                  mov dword ptr [esi+0c], eax
'00718798    8b853cffffff            mov eax, dword ptr [ebp+ffffff3c]
'0071879e    8b1548b07200            mov edx, dword ptr [0072b048]
'007187a4    8b12                    mov edx, dword ptr [edx]
'007187a6    83ec10                  sub esp, 10
'007187a9    8bcc                    mov ecx, esp
'007187ab    8901                    mov dword ptr [ecx], eax
'007187ad    8b8540ffffff            mov eax, dword ptr [ebp+ffffff40]
'007187b3    894104                  mov dword ptr [ecx+04], eax
'007187b6    8b8544ffffff            mov eax, dword ptr [ebp+ffffff44]
'007187bc    894108                  mov dword ptr [ecx+08], eax
'007187bf    8b8548ffffff            mov eax, dword ptr [ebp+ffffff48]
'007187c5    89410c                  mov dword ptr [ecx+0c], eax
'007187c8    83ec10                  sub esp, 10
'007187cb    8bcc                    mov ecx, esp
'007187cd    b803000000              mov eax, 00000003
'007187d2    8901                    mov dword ptr [ecx], eax
'007187d4    8b8550ffffff            mov eax, dword ptr [ebp+ffffff50]
'007187da    894104                  mov dword ptr [ecx+04], eax
'007187dd    b801000000              mov eax, 00000001
'007187e2    894108                  mov dword ptr [ecx+08], eax
'007187e5    8b8558ffffff            mov eax, dword ptr [ebp+ffffff58]
'007187eb    89410c                  mov dword ptr [ecx+0c], eax
'007187ee    8b0d48b07200            mov ecx, dword ptr [0072b048]
'007187f4    6828b54300              push 0043b528
'007187f9    51                      push ecx
'007187fa    ff92bc000000            call dword ptr [edx+000000bc]
'00718800    dbe2                    fnclex
'00718802    85c0                    test ax, ax
'00718804    7d18                    jge 71881e
'00718806    8b1548b07200            mov edx, dword ptr [0072b048]
'0071880c    68bc000000              push 000000bc
'00718811    68301f4300              push 00431f30
'00718816    52                      push edx
'00718817    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00718818    ff1580104000            call dword ptr [00401080]
'0071881e    8b45b8                  mov eax, dword ptr [ebp-48]
'00718821    50                      push eax
'00718822    8d45dc                  lea eax, dword ptr [ebp-24]
'00718825    50                      push eax
'00718826    c745b800000000          mov dword ptr [ebp-48], 00000000

' *** Reference to "__vbaObjSet"
'0071882d    ff15b4104000            call dword ptr [004010b4]
Set var_39 = Nothing
'00718833    eb03                    jmp 718838
'00718835    8b5d08                  mov ebx, dword ptr [ebp+08]

' *** Reference to "__vbaHresultCheckObj"
'00718838    8b3580104000            mov esi, dword ptr [00401080]
'0071883e    8b45d8                  mov eax, dword ptr [ebp-28]
'00718841    8b08                    mov ecx, dword ptr [eax]
'00718843    8d9538feffff            lea edx, dword ptr [ebp+fffffe38]
'00718849    52                      push edx
'0071884a    50                      push eax
'0071884b    ff5134                  call dword ptr [ecx+34]
'0071884e    dbe2                    fnclex
'00718850    85c0                    test ax, ax
'00718852    7d0e                    jge 718862
'00718854    8b4dd8                  mov ecx, dword ptr [ebp-28]
'00718857    6a34                    push 34
'00718859    6830314300              push 00433130
'0071885e    51                      push ecx
'0071885f    50                      push eax
'00718860    ffd6                    call esi
'00718862    6683bd38feffff00        cmp word ptr [ebp+fffffe38], 00
'0071886a    8b45dc                  mov eax, dword ptr [ebp-24]
'0071886d    0f8581040000            jne 718cf4

Do While (0 = 0)
'00718873    8b10                    mov edx, dword ptr [eax]
'00718875    50                      push eax
'00718876    ff92c0000000            call dword ptr [edx+000000c0]
'0071887c    dbe2                    fnclex
'0071887e    85c0                    test ax, ax
'00718880    7d11                    jge 718893
'00718882    8b4ddc                  mov ecx, dword ptr [ebp-24]
'00718885    68c0000000              push 000000c0
'0071888a    6830314300              push 00433130
'0071888f    51                      push ecx
'00718890    50                      push eax
'00718891    ffd6                    call esi
'00718893    8b45dc                  mov eax, dword ptr [ebp-24]
'00718896    8b10                    mov edx, dword ptr [eax]
'00718898    8d4db4                  lea ecx, dword ptr [ebp-4c]
'0071889b    51                      push ecx
'0071889c    50                      push eax
'0071889d    ff92b4000000            call dword ptr [edx+000000b4]
'007188a3    dbe2                    fnclex
'007188a5    85c0                    test ax, ax
'007188a7    7d11                    jge 7188ba
'007188a9    8b55dc                  mov edx, dword ptr [ebp-24]
'007188ac    68b4000000              push 000000b4
'007188b1    6830314300              push 00433130
'007188b6    52                      push edx
'007188b7    50                      push eax
'007188b8    ffd6                    call esi
'007188ba    8b45b4                  mov eax, dword ptr [ebp-4c]
'007188bd    8d5db0                  lea ebx, dword ptr [ebp-50]
'007188c0    53                      push ebx
'007188c1    83ec10                  sub esp, 10
'007188c4    b902000000              mov ecx, 00000002
'007188c9    898decfeffff            mov dword ptr [ebp+fffffeec], ecx
'007188cf    8bdc                    mov ebx, esp
'007188d1    890b                    mov dword ptr [ebx], ecx
'007188d3    8b8df0feffff            mov ecx, dword ptr [ebp+fffffef0]
'007188d9    894b04                  mov dword ptr [ebx+04], ecx
'007188dc    c785f4feffff00000000    mov dword ptr [ebp+fffffef4], 00000000
'007188e6    8b8df4feffff            mov ecx, dword ptr [ebp+fffffef4]
'007188ec    8b10                    mov edx, dword ptr [eax]
'007188ee    894b08                  mov dword ptr [ebx+08], ecx
'007188f1    8b8df8feffff            mov ecx, dword ptr [ebp+fffffef8]
'007188f7    50                      push eax
'007188f8    898530feffff            mov dword ptr [ebp+fffffe30], eax
'007188fe    894b0c                  mov dword ptr [ebx+0c], ecx
'00718901    ff5230                  call dword ptr [edx+30]
'00718904    dbe2                    fnclex
'00718906    85c0                    test ax, ax
'00718908    7d11                    jge 71891b
'0071890a    8b9530feffff            mov edx, dword ptr [ebp+fffffe30]
'00718910    6a30                    push 30
'00718912    68d8304300              push 004330d8
'00718917    52                      push edx
'00718918    50                      push eax
'00718919    ffd6                    call esi
'0071891b    668b55e8                mov dx, word ptr [ebp-18]
'0071891f    66899534ffffff          mov word ptr [ebp+ffffff34], dx
'00718926    83ec10                  sub esp, 10
'00718929    8bd4                    mov edx, esp
'0071892b    33c0                    xor eax, eax
    var_num1 = Empty
'0071892d    8b75b0                  mov esi, dword ptr [ebp-50]
'00718930    898554ffffff            mov dword ptr [ebp+ffffff54], eax
'00718936    b903000000              mov ecx, 00000003
'0071893b    890a                    mov dword ptr [edx], ecx
'0071893d    8b8d50ffffff            mov ecx, dword ptr [ebp+ffffff50]
'00718943    894a04                  mov dword ptr [edx+04], ecx
'00718946    894208                  mov dword ptr [edx+08], eax
'00718949    8b8558ffffff            mov eax, dword ptr [ebp+ffffff58]
'0071894f    89420c                  mov dword ptr [edx+0c], eax
'00718952    8b9530ffffff            mov edx, dword ptr [ebp+ffffff30]
'00718958    8b1e                    mov ebx, dword ptr [esi]
'0071895a    83ec10                  sub esp, 10
'0071895d    8bcc                    mov ecx, esp
'0071895f    b802000000              mov eax, 00000002
'00718964    8901                    mov dword ptr [ecx], eax
'00718966    8b8534ffffff            mov eax, dword ptr [ebp+ffffff34]
'0071896c    895104                  mov dword ptr [ecx+04], edx
'0071896f    8b9538ffffff            mov edx, dword ptr [ebp+ffffff38]
'00718975    894108                  mov dword ptr [ecx+08], eax
'00718978    89510c                  mov dword ptr [ecx+0c], edx
'0071897b    8b9510ffffff            mov edx, dword ptr [ebp+ffffff10]
'00718981    83ec10                  sub esp, 10
'00718984    8bc4                    mov eax, esp
'00718986    c7850cffffff02000000    mov dword ptr [ebp+ffffff0c], 00000002
'00718990    8b8d0cffffff            mov ecx, dword ptr [ebp+ffffff0c]
'00718996    8908                    mov dword ptr [eax], ecx
'00718998    895004                  mov dword ptr [eax+04], edx
'0071899b    8b9518ffffff            mov edx, dword ptr [ebp+ffffff18]
'007189a1    c78514ffffff01000000    mov dword ptr [ebp+ffffff14], 00000001
'007189ab    8b8d14ffffff            mov ecx, dword ptr [ebp+ffffff14]
'007189b1    894808                  mov dword ptr [eax+08], ecx
'007189b4    6a03                    push 03
'007189b6    89500c                  mov dword ptr [eax+0c], edx
'007189b9    8b4508                  mov eax, dword ptr [ebp+08]
'007189bc    8b08                    mov ecx, dword ptr [eax]
'007189be    689e000000              push 0000009e
'007189c3    50                      push eax
'007189c4    ff9110030000            call dword ptr [ecx+00000310]
'007189ca    50                      push eax
'007189cb    8d55b8                  lea edx, dword ptr [ebp-48]
'007189ce    52                      push edx

' *** Reference to "__vbaObjSet"
'007189cf    ff15b4104000            call dword ptr [004010b4]
    Set var_61 = Me
'007189d5    50                      push eax
'007189d6    8d459c                  lea eax, dword ptr [ebp-64]
'007189d9    50                      push eax

' *** Reference to "__vbaLateIdCallLd"
'007189da    ff157c114000            call dword ptr [0040117c]
    var_51 = var_61.UNK_frmImport_158
'007189e0    8b10                    mov edx, dword ptr [eax]
'007189e2    83c430                  add esp, 30
'007189e5    8bcc                    mov ecx, esp
'007189e7    8911                    mov dword ptr [ecx], edx
'007189e9    8b5004                  mov edx, dword ptr [eax+04]
'007189ec    895104                  mov dword ptr [ecx+04], edx
'007189ef    8b5008                  mov edx, dword ptr [eax+08]
'007189f2    8b400c                  mov eax, dword ptr [eax+0c]
'007189f5    895108                  mov dword ptr [ecx+08], edx
'007189f8    56                      push esi
'007189f9    89410c                  mov dword ptr [ecx+0c], eax
'007189fc    ff5348                  call dword ptr [ebx+48]
'007189ff    dbe2                    fnclex
'00718a01    85c0                    test ax, ax
'00718a03    7d0f                    jge 718a14
'00718a05    6a48                    push 48
'00718a07    6880304300              push 00433080
'00718a0c    56                      push esi
'00718a0d    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00718a0e    ff1580104000            call dword ptr [00401080]
'00718a14    8d4db0                  lea ecx, dword ptr [ebp-50]
'00718a17    51                      push ecx
'00718a18    8d55b4                  lea edx, dword ptr [ebp-4c]
'00718a1b    52                      push edx
'00718a1c    8d45b8                  lea eax, dword ptr [ebp-48]
'00718a1f    50                      push eax
'00718a20    6a03                    push 03

' *** Reference to "__vbaFreeObjList"
'00718a22    ff154c104000            call dword ptr [0040104c]
    '#Cleanup( var_61, var_62, var_6)
'00718a28    83c410                  add esp, 10
'00718a2b    8d4d9c                  lea ecx, dword ptr [ebp-64]

' *** Reference to "__vbaFreeVar"
'00718a2e    ff1528104000            call dword ptr [00401028]
    '#Cleanup(var_51)
'00718a34    bb01000000              mov ebx, 00000001
'00718a39    b805000000              mov eax, 00000005
'00718a3e    663bd8                  cmp bx, ax
'00718a41    895de0                  mov dword ptr [ebp-20], ebx
'00718a44    0f8f4a020000            jg 718c94
    
    Do While (    1 <= 5)
'00718a4a    8b45d8                  mov eax, dword ptr [ebp-28]
'00718a4d    8b08                    mov ecx, dword ptr [eax]
'00718a4f    8d55b8                  lea edx, dword ptr [ebp-48]
'00718a52    52                      push edx
'00718a53    50                      push eax
'00718a54    ff91b4000000            call dword ptr [ecx+000000b4]
'00718a5a    dbe2                    fnclex
'00718a5c    85c0                    test ax, ax
'00718a5e    7d15                    jge 718a75
'00718a60    8b4dd8                  mov ecx, dword ptr [ebp-28]
'00718a63    68b4000000              push 000000b4
'00718a68    6830314300              push 00433130
'00718a6d    51                      push ecx
'00718a6e    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00718a6f    ff1580104000            call dword ptr [00401080]
'00718a75    8b45b8                  mov eax, dword ptr [ebp-48]
'00718a78    8b10                    mov edx, dword ptr [eax]
'00718a7a    66899d54ffffff          mov word ptr [ebp+ffffff54], bx
'00718a81    8d5db4                  lea ebx, dword ptr [ebp-4c]
'00718a84    53                      push ebx
'00718a85    83ec10                  sub esp, 10
'00718a88    8bdc                    mov ebx, esp
'00718a8a    b902000000              mov ecx, 00000002
'00718a8f    890b                    mov dword ptr [ebx], ecx
'00718a91    8b8d50ffffff            mov ecx, dword ptr [ebp+ffffff50]
'00718a97    894b04                  mov dword ptr [ebx+04], ecx
'00718a9a    8b8d54ffffff            mov ecx, dword ptr [ebp+ffffff54]
'00718aa0    894b08                  mov dword ptr [ebx+08], ecx
'00718aa3    8b8d58ffffff            mov ecx, dword ptr [ebp+ffffff58]
'00718aa9    50                      push eax
'00718aaa    8bf0                    mov esi, eax
'00718aac    894b0c                  mov dword ptr [ebx+0c], ecx
'00718aaf    ff5230                  call dword ptr [edx+30]
'00718ab2    dbe2                    fnclex
'00718ab4    85c0                    test ax, ax
'00718ab6    7d0f                    jge 718ac7
'00718ab8    6a30                    push 30
'00718aba    68d8304300              push 004330d8
'00718abf    56                      push esi
'00718ac0    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00718ac1    ff1580104000            call dword ptr [00401080]
'00718ac7    8b45b4                  mov eax, dword ptr [ebp-4c]
'00718aca    8d559c                  lea edx, dword ptr [ebp-64]
'00718acd    52                      push edx
'00718ace    c745b400000000          mov dword ptr [ebp-4c], 00000000
'00718ad5    8945a4                  mov dword ptr [ebp-5c], eax
'00718ad8    c7459c09000000          mov dword ptr [ebp-64], 00000009

' *** Reference to "rtcIsNull"
'00718adf    ff1540114000            call dword ptr [00401140]
'00718ae5    33f6                    xor esi, esi
    var_num8 = Empty
'00718ae7    668bf0                  mov si, ax
'00718aea    8d4db8                  lea ecx, dword ptr [ebp-48]
'00718aed    f7d6                    not esi

' *** Reference to "__vbaFreeObj"
'00718aef    ff1508134000            call dword ptr [00401308]
    '#Cleanup(var_61)
'00718af5    8d4d9c                  lea ecx, dword ptr [ebp-64]

' *** Reference to "__vbaFreeVar"
'00718af8    ff1528104000            call dword ptr [00401028]
    '#Cleanup(var_51)
'00718afe    6685f6                  test esi, esi
'00718b01    0f8477010000            je 718c7e
    
    If (    Not (IsNull(var_62))) Then
'00718b07    8b45dc                  mov eax, dword ptr [ebp-24]
'00718b0a    8b08                    mov ecx, dword ptr [eax]
'00718b0c    8d55b0                  lea edx, dword ptr [ebp-50]
'00718b0f    52                      push edx
'00718b10    50                      push eax
'00718b11    ff91b4000000            call dword ptr [ecx+000000b4]
'00718b17    dbe2                    fnclex
'00718b19    85c0                    test ax, ax
'00718b1b    7d15                    jge 718b32
'00718b1d    8b4ddc                  mov ecx, dword ptr [ebp-24]
'00718b20    68b4000000              push 000000b4
'00718b25    6830314300              push 00433130
'00718b2a    51                      push ecx
'00718b2b    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00718b2c    ff1580104000            call dword ptr [00401080]
'00718b32    668b55e0                mov dx, word ptr [ebp-20]
'00718b36    8b45b0                  mov eax, dword ptr [ebp-50]
'00718b39    8d5dac                  lea ebx, dword ptr [ebp-54]
'00718b3c    53                      push ebx
'00718b3d    83ec10                  sub esp, 10
'00718b40    8bdc                    mov ebx, esp
'00718b42    b902000000              mov ecx, 00000002
'00718b47    890b                    mov dword ptr [ebx], ecx
'00718b49    8b8d40ffffff            mov ecx, dword ptr [ebp+ffffff40]
'00718b4f    894b04                  mov dword ptr [ebx+04], ecx
'00718b52    66899544ffffff          mov word ptr [ebp+ffffff44], dx
'00718b59    8b8d44ffffff            mov ecx, dword ptr [ebp+ffffff44]
'00718b5f    8b10                    mov edx, dword ptr [eax]
'00718b61    894b08                  mov dword ptr [ebx+08], ecx
'00718b64    8b8d48ffffff            mov ecx, dword ptr [ebp+ffffff48]
'00718b6a    50                      push eax
'00718b6b    8bf0                    mov esi, eax
'00718b6d    894b0c                  mov dword ptr [ebx+0c], ecx
'00718b70    ff5230                  call dword ptr [edx+30]
'00718b73    dbe2                    fnclex
'00718b75    85c0                    test ax, ax
'00718b77    7d0f                    jge 718b88
'00718b79    6a30                    push 30
'00718b7b    68d8304300              push 004330d8
'00718b80    56                      push esi
'00718b81    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00718b82    ff1580104000            call dword ptr [00401080]
'00718b88    8b55ac                  mov edx, dword ptr [ebp-54]
'00718b8b    8b45d8                  mov eax, dword ptr [ebp-28]
'00718b8e    8b08                    mov ecx, dword ptr [eax]
'00718b90    89951cfeffff            mov dword ptr [ebp+fffffe1c], edx
'00718b96    8d55b8                  lea edx, dword ptr [ebp-48]
'00718b99    52                      push edx
'00718b9a    50                      push eax
'00718b9b    ff91b4000000            call dword ptr [ecx+000000b4]
'00718ba1    dbe2                    fnclex
'00718ba3    85c0                    test ax, ax
'00718ba5    7d15                    jge 718bbc
'00718ba7    8b4dd8                  mov ecx, dword ptr [ebp-28]
'00718baa    68b4000000              push 000000b4
'00718baf    6830314300              push 00433130
'00718bb4    51                      push ecx
'00718bb5    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00718bb6    ff1580104000            call dword ptr [00401080]
'00718bbc    668b55e0                mov dx, word ptr [ebp-20]
'00718bc0    8b45b8                  mov eax, dword ptr [ebp-48]
'00718bc3    8d5db4                  lea ebx, dword ptr [ebp-4c]
'00718bc6    53                      push ebx
'00718bc7    83ec10                  sub esp, 10
'00718bca    8bdc                    mov ebx, esp
'00718bcc    b902000000              mov ecx, 00000002
'00718bd1    890b                    mov dword ptr [ebx], ecx
'00718bd3    8b8d50ffffff            mov ecx, dword ptr [ebp+ffffff50]
'00718bd9    894b04                  mov dword ptr [ebx+04], ecx
'00718bdc    66899554ffffff          mov word ptr [ebp+ffffff54], dx
'00718be3    8b8d54ffffff            mov ecx, dword ptr [ebp+ffffff54]
'00718be9    8b10                    mov edx, dword ptr [eax]
'00718beb    894b08                  mov dword ptr [ebx+08], ecx
'00718bee    8b8d58ffffff            mov ecx, dword ptr [ebp+ffffff58]
'00718bf4    50                      push eax
'00718bf5    8bf0                    mov esi, eax
'00718bf7    894b0c                  mov dword ptr [ebx+0c], ecx
'00718bfa    ff5230                  call dword ptr [edx+30]
'00718bfd    dbe2                    fnclex
'00718bff    85c0                    test ax, ax
'00718c01    7d0f                    jge 718c12
'00718c03    6a30                    push 30
'00718c05    68d8304300              push 004330d8
'00718c0a    56                      push esi
'00718c0b    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00718c0c    ff1580104000            call dword ptr [00401080]
'00718c12    8b45b4                  mov eax, dword ptr [ebp-4c]
'00718c15    8bb51cfeffff            mov esi, dword ptr [ebp+fffffe1c]
'00718c1b    83ec10                  sub esp, 10
'00718c1e    b909000000              mov ecx, 00000009
'00718c23    8bdc                    mov ebx, esp
'00718c25    890b                    mov dword ptr [ebx], ecx
'00718c27    894d9c                  mov dword ptr [ebp-64], ecx
'00718c2a    8b4da0                  mov ecx, dword ptr [ebp-60]
'00718c2d    894b04                  mov dword ptr [ebx+04], ecx
'00718c30    8945a4                  mov dword ptr [ebp-5c], eax
'00718c33    894308                  mov dword ptr [ebx+08], eax
'00718c36    8b45a8                  mov eax, dword ptr [ebp-58]
'00718c39    c745b400000000          mov dword ptr [ebp-4c], 00000000
'00718c40    8b16                    mov edx, dword ptr [esi]
'00718c42    56                      push esi
'00718c43    89430c                  mov dword ptr [ebx+0c], eax
'00718c46    ff5248                  call dword ptr [edx+48]
'00718c49    dbe2                    fnclex
'00718c4b    85c0                    test ax, ax
'00718c4d    7d0f                    jge 718c5e
'00718c4f    6a48                    push 48
'00718c51    6880304300              push 00433080
'00718c56    56                      push esi
'00718c57    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00718c58    ff1580104000            call dword ptr [00401080]
'00718c5e    8d4dac                  lea ecx, dword ptr [ebp-54]
'00718c61    51                      push ecx
'00718c62    8d55b0                  lea edx, dword ptr [ebp-50]
'00718c65    52                      push edx
'00718c66    8d45b8                  lea eax, dword ptr [ebp-48]
'00718c69    50                      push eax
'00718c6a    6a03                    push 03

' *** Reference to "__vbaFreeObjList"
'00718c6c    ff154c104000            call dword ptr [0040104c]
    '#Cleanup( var_61, var_6, var_50)
'00718c72    83c410                  add esp, 10
'00718c75    8d4d9c                  lea ecx, dword ptr [ebp-64]

' *** Reference to "__vbaFreeVar"
'00718c78    ff1528104000            call dword ptr [00401028]
    '#Cleanup(var_51)
End If
'00718c7e    b801000000              mov eax, 00000001
'00718c83    660345e0                add ax, word ptr [ebp-20]
var_num1 = 1 + 1
'00718c87    0f80e61b0000            jo 71a873
'00718c8d    8bd8                    mov ebx, eax
'00718c8f    e9a5fdffff              jmp 718a39

'ERROR: Two many next close:
Loop
'00718c94    8b45dc                  mov eax, dword ptr [ebp-24]
'00718c97    8b08                    mov ecx, dword ptr [eax]
'00718c99    6a00                    push 00
'00718c9b    6a01                    push 01
'00718c9d    50                      push eax
'00718c9e    ff9164010000            call dword ptr [ecx+00000164]
'00718ca4    dbe2                    fnclex
'00718ca6    85c0                    test ax, ax
'00718ca8    7d15                    jge 718cbf
'00718caa    8b55dc                  mov edx, dword ptr [ebp-24]
'00718cad    6864010000              push 00000164
'00718cb2    6830314300              push 00433130
'00718cb7    52                      push edx
'00718cb8    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00718cb9    ff1580104000            call dword ptr [00401080]
'00718cbf    8b45d8                  mov eax, dword ptr [ebp-28]
'00718cc2    8b08                    mov ecx, dword ptr [eax]
'00718cc4    50                      push eax
'00718cc5    ff91ec000000            call dword ptr [ecx+000000ec]
'00718ccb    dbe2                    fnclex
'00718ccd    85c0                    test ax, ax
'00718ccf    0f8d60fbffff            jge 718835
'00718cd5    8b55d8                  mov edx, dword ptr [ebp-28]

' *** Reference to "__vbaHresultCheckObj"
'00718cd8    8b3580104000            mov esi, dword ptr [00401080]
'00718cde    68ec000000              push 000000ec
'00718ce3    6830314300              push 00433130
'00718ce8    52                      push edx
'00718ce9    50                      push eax
'00718cea    ffd6                    call esi
'00718cec    8b5d08                  mov ebx, dword ptr [ebp+08]
'00718cef    e94afbffff              jmp 71883e

'ERROR: Two many next close:
Loop
'00718cf4    8b08                    mov ecx, dword ptr [eax]
'00718cf6    50                      push eax
'00718cf7    ff91c4000000            call dword ptr [ecx+000000c4]
'00718cfd    dbe2                    fnclex
'00718cff    85c0                    test ax, ax
'00718d01    7d11                    jge 718d14
'00718d03    8b55dc                  mov edx, dword ptr [ebp-24]
'00718d06    68c4000000              push 000000c4
'00718d0b    6830314300              push 00433130
'00718d10    52                      push edx
'00718d11    50                      push eax
'00718d12    ffd6                    call esi
'00718d14    8b45d8                  mov eax, dword ptr [ebp-28]
'00718d17    8b08                    mov ecx, dword ptr [eax]
'00718d19    50                      push eax
'00718d1a    ff91c4000000            call dword ptr [ecx+000000c4]
'00718d20    dbe2                    fnclex
'00718d22    85c0                    test ax, ax
'00718d24    7d11                    jge 718d37
'00718d26    8b55d8                  mov edx, dword ptr [ebp-28]
'00718d29    68c4000000              push 000000c4
'00718d2e    6830314300              push 00433130
'00718d33    52                      push edx
'00718d34    50                      push eax
'00718d35    ffd6                    call esi
'00718d37    668b55e8                mov dx, word ptr [ebp-18]
'00718d3b    83ec10                  sub esp, 10
'00718d3e    8bf4                    mov esi, esp
'00718d40    33c0                    xor eax, eax
var_num1 = Empty
'00718d42    b903000000              mov ecx, 00000003
'00718d47    890e                    mov dword ptr [esi], ecx
'00718d49    8b8d50ffffff            mov ecx, dword ptr [ebp+ffffff50]
'00718d4f    894e04                  mov dword ptr [esi+04], ecx
'00718d52    894608                  mov dword ptr [esi+08], eax
'00718d55    8b8558ffffff            mov eax, dword ptr [ebp+ffffff58]
'00718d5b    89460c                  mov dword ptr [esi+0c], eax
'00718d5e    83ec10                  sub esp, 10
'00718d61    66899534ffffff          mov word ptr [ebp+ffffff34], dx
'00718d68    8b8534ffffff            mov eax, dword ptr [ebp+ffffff34]
'00718d6e    8bcc                    mov ecx, esp
'00718d70    ba02000000              mov edx, 00000002
'00718d75    8911                    mov dword ptr [ecx], edx
'00718d77    8b9530ffffff            mov edx, dword ptr [ebp+ffffff30]
'00718d7d    895104                  mov dword ptr [ecx+04], edx
'00718d80    8b9538ffffff            mov edx, dword ptr [ebp+ffffff38]
'00718d86    894108                  mov dword ptr [ecx+08], eax
'00718d89    89510c                  mov dword ptr [ecx+0c], edx
'00718d8c    8b9510ffffff            mov edx, dword ptr [ebp+ffffff10]
'00718d92    83ec10                  sub esp, 10
'00718d95    8bcc                    mov ecx, esp
'00718d97    b802000000              mov eax, 00000002
'00718d9c    8901                    mov dword ptr [ecx], eax
'00718d9e    895104                  mov dword ptr [ecx+04], edx
'00718da1    33c0                    xor eax, eax
var_num1 = Empty
'00718da3    894108                  mov dword ptr [ecx+08], eax
'00718da6    8b8518ffffff            mov eax, dword ptr [ebp+ffffff18]
'00718dac    6a03                    push 03
'00718dae    689e000000              push 0000009e
'00718db3    89410c                  mov dword ptr [ecx+0c], eax
'00718db6    8b0b                    mov ecx, dword ptr [ebx]
'00718db8    53                      push ebx
'00718db9    ff9110030000            call dword ptr [ecx+00000310]
'00718dbf    50                      push eax
'00718dc0    8d55b8                  lea edx, dword ptr [ebp-48]
'00718dc3    52                      push edx

' *** Reference to "__vbaObjSet"
'00718dc4    ff15b4104000            call dword ptr [004010b4]
Set var_61 = Nothing
'00718dca    50                      push eax
'00718dcb    8d459c                  lea eax, dword ptr [ebp-64]
'00718dce    50                      push eax

' *** Reference to "__vbaLateIdCallLd"
'00718dcf    ff157c114000            call dword ptr [0040117c]
var_51 = var_61.UNK_-256 - 12_158
'00718dd5    83c440                  add esp, 40
'00718dd8    50                      push eax

' *** Reference to "__vbaStrVarMove"
'00718dd9    ff153c104000            call dword ptr [0040103c]
'00718ddf    8bd0                    mov edx, eax
'00718de1    8d4dcc                  lea ecx, dword ptr [ebp-34]
'00718de4    ffd7                    call edi
'00718de6    8d4dcc                  lea ecx, dword ptr [ebp-34]
'00718de9    51                      push ecx
'00718dea    e801aeddff              call 4f3bf0
Call sub_4F3BF0()
'00718def    8bd0                    mov edx, eax
'00718df1    8d4dbc                  lea ecx, dword ptr [ebp-44]
'00718df4    ffd7                    call edi
'00718df6    8b55bc                  mov edx, dword ptr [ebp-44]
'00718df9    89951cfdffff            mov dword ptr [ebp+fffffd1c], edx
'00718dff    8b5338                  mov edx, dword ptr [ebx+38]
'00718e02    c785f4feffff02000000    mov dword ptr [ebp+fffffef4], 00000002
'00718e0c    c785ecfeffff03000000    mov dword ptr [ebp+fffffeec], 00000003
'00718e16    c745bc00000000          mov dword ptr [ebp-44], 00000000
'00718e1d    8b32                    mov esi, dword ptr [edx]
'00718e1f    8d55b4                  lea edx, dword ptr [ebp-4c]
'00718e22    52                      push edx
'00718e23    83ec10                  sub esp, 10
'00718e26    b90a000000              mov ecx, 0000000a
'00718e2b    8bd4                    mov edx, esp
'00718e2d    890a                    mov dword ptr [edx], ecx
'00718e2f    898ddcfeffff            mov dword ptr [ebp+fffffedc], ecx
'00718e35    8b8dd0feffff            mov ecx, dword ptr [ebp+fffffed0]
'00718e3b    894a04                  mov dword ptr [edx+04], ecx
'00718e3e    b804000280              mov eax, 80020004
'00718e43    894208                  mov dword ptr [edx+08], eax
'00718e46    8985e4feffff            mov dword ptr [ebp+fffffee4], eax
'00718e4c    8b85d8feffff            mov eax, dword ptr [ebp+fffffed8]
'00718e52    83ec10                  sub esp, 10
'00718e55    89420c                  mov dword ptr [edx+0c], eax
'00718e58    8b95dcfeffff            mov edx, dword ptr [ebp+fffffedc]
'00718e5e    8bcc                    mov ecx, esp
'00718e60    8b85e0feffff            mov eax, dword ptr [ebp+fffffee0]
'00718e66    8911                    mov dword ptr [ecx], edx
'00718e68    8b95e4feffff            mov edx, dword ptr [ebp+fffffee4]
'00718e6e    894104                  mov dword ptr [ecx+04], eax
'00718e71    8b85e8feffff            mov eax, dword ptr [ebp+fffffee8]
'00718e77    895108                  mov dword ptr [ecx+08], edx
'00718e7a    8b95ecfeffff            mov edx, dword ptr [ebp+fffffeec]
'00718e80    89410c                  mov dword ptr [ecx+0c], eax
'00718e83    8b85f0feffff            mov eax, dword ptr [ebp+fffffef0]
'00718e89    83ec10                  sub esp, 10
'00718e8c    8bcc                    mov ecx, esp
'00718e8e    8911                    mov dword ptr [ecx], edx
'00718e90    8b95f4feffff            mov edx, dword ptr [ebp+fffffef4]
'00718e96    894104                  mov dword ptr [ecx+04], eax
'00718e99    8b85f8feffff            mov eax, dword ptr [ebp+fffffef8]
'00718e9f    895108                  mov dword ptr [ecx+08], edx
'00718ea2    8b951cfdffff            mov edx, dword ptr [ebp+fffffd1c]
'00718ea8    89410c                  mov dword ptr [ecx+0c], eax
'00718eab    68e01e4500              push 00451ee0
'00718eb0    8d4dc8                  lea ecx, dword ptr [ebp-38]
'00718eb3    ffd7                    call edi
'00718eb5    50                      push eax

' *** Reference to "__vbaStrCat"
'00718eb6    ff1570104000            call dword ptr [00401070]
var_2436 = ("select * from PersonnageProgression where nom='") & (var_51)
'00718ebc    8bd0                    mov edx, eax
'00718ebe    8d4dc4                  lea ecx, dword ptr [ebp-3c]
'00718ec1    ffd7                    call edi
'00718ec3    50                      push eax
'00718ec4    6854a44300              push 0043a454

' *** Reference to "__vbaStrCat"
'00718ec9    ff1570104000            call dword ptr [00401070]
var_505 = (var_2436) & ("'")
'00718ecf    8bd0                    mov edx, eax
'00718ed1    8d4dc0                  lea ecx, dword ptr [ebp-40]
'00718ed4    ffd7                    call edi
'00718ed6    8b4b38                  mov ecx, dword ptr [ebx+38]
'00718ed9    50                      push eax
'00718eda    51                      push ecx
'00718edb    ff96bc000000            call dword ptr [esi+000000bc]
'00718ee1    dbe2                    fnclex
'00718ee3    85c0                    test ax, ax
'00718ee5    7d15                    jge 718efc
'00718ee7    8b5338                  mov edx, dword ptr [ebx+38]
'00718eea    68bc000000              push 000000bc
'00718eef    68301f4300              push 00431f30
'00718ef4    52                      push edx
'00718ef5    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00718ef6    ff1580104000            call dword ptr [00401080]
'00718efc    8b45b4                  mov eax, dword ptr [ebp-4c]
'00718eff    50                      push eax
'00718f00    8d45d8                  lea eax, dword ptr [ebp-28]
'00718f03    50                      push eax
'00718f04    c745b400000000          mov dword ptr [ebp-4c], 00000000

' *** Reference to "__vbaObjSet"
'00718f0b    ff15b4104000            call dword ptr [004010b4]
Set var_45 = Nothing
'00718f11    8d4dbc                  lea ecx, dword ptr [ebp-44]
'00718f14    51                      push ecx
'00718f15    8d55c0                  lea edx, dword ptr [ebp-40]
'00718f18    52                      push edx
'00718f19    8d45c4                  lea eax, dword ptr [ebp-3c]
'00718f1c    50                      push eax
'00718f1d    8d4dc8                  lea ecx, dword ptr [ebp-38]
'00718f20    51                      push ecx
'00718f21    8d55cc                  lea edx, dword ptr [ebp-34]
'00718f24    52                      push edx
'00718f25    6a05                    push 05

' *** Reference to "__vbaFreeStrList"
'00718f27    ff155c124000            call dword ptr [0040125c]
'#Cleanup( -4796, -4796, -4800, -4804, var_58)
'00718f2d    83c418                  add esp, 18
'00718f30    8d4db8                  lea ecx, dword ptr [ebp-48]

' *** Reference to "__vbaFreeObj"
'00718f33    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_61)
'00718f39    8d4d9c                  lea ecx, dword ptr [ebp-64]

' *** Reference to "__vbaFreeVar"
'00718f3c    ff1528104000            call dword ptr [00401028]
'#Cleanup(var_51)
'00718f42    8d75b8                  lea esi, dword ptr [ebp-48]
'00718f45    56                      push esi
'00718f46    83ec10                  sub esp, 10
'00718f49    8bf4                    mov esi, esp
'00718f4b    b90a000000              mov ecx, 0000000a
'00718f50    890e                    mov dword ptr [esi], ecx
'00718f52    898d3cffffff            mov dword ptr [ebp+ffffff3c], ecx
'00718f58    8b8d30ffffff            mov ecx, dword ptr [ebp+ffffff30]
'00718f5e    894e04                  mov dword ptr [esi+04], ecx
'00718f61    b804000280              mov eax, 80020004
'00718f66    894608                  mov dword ptr [esi+08], eax
'00718f69    898534ffffff            mov dword ptr [ebp+ffffff34], eax
'00718f6f    898544ffffff            mov dword ptr [ebp+ffffff44], eax
'00718f75    8b8538ffffff            mov eax, dword ptr [ebp+ffffff38]
'00718f7b    89460c                  mov dword ptr [esi+0c], eax
'00718f7e    8b853cffffff            mov eax, dword ptr [ebp+ffffff3c]
'00718f84    8b1548b07200            mov edx, dword ptr [0072b048]
'00718f8a    8b12                    mov edx, dword ptr [edx]
'00718f8c    83ec10                  sub esp, 10
'00718f8f    8bcc                    mov ecx, esp
'00718f91    8901                    mov dword ptr [ecx], eax
'00718f93    8b8540ffffff            mov eax, dword ptr [ebp+ffffff40]
'00718f99    894104                  mov dword ptr [ecx+04], eax
'00718f9c    8b8544ffffff            mov eax, dword ptr [ebp+ffffff44]
'00718fa2    894108                  mov dword ptr [ecx+08], eax
'00718fa5    8b8548ffffff            mov eax, dword ptr [ebp+ffffff48]
'00718fab    89410c                  mov dword ptr [ecx+0c], eax
'00718fae    83ec10                  sub esp, 10
'00718fb1    8bcc                    mov ecx, esp
'00718fb3    b803000000              mov eax, 00000003
'00718fb8    8901                    mov dword ptr [ecx], eax
'00718fba    8b8550ffffff            mov eax, dword ptr [ebp+ffffff50]
'00718fc0    894104                  mov dword ptr [ecx+04], eax
'00718fc3    b801000000              mov eax, 00000001
'00718fc8    894108                  mov dword ptr [ecx+08], eax
'00718fcb    8b8558ffffff            mov eax, dword ptr [ebp+ffffff58]
'00718fd1    89410c                  mov dword ptr [ecx+0c], eax
'00718fd4    8b0d48b07200            mov ecx, dword ptr [0072b048]
'00718fda    68e4b44300              push 0043b4e4
'00718fdf    51                      push ecx
'00718fe0    ff92bc000000            call dword ptr [edx+000000bc]
'00718fe6    dbe2                    fnclex
'00718fe8    85c0                    test ax, ax
'00718fea    7d18                    jge 719004
'00718fec    8b1548b07200            mov edx, dword ptr [0072b048]
'00718ff2    68bc000000              push 000000bc
'00718ff7    68301f4300              push 00431f30
'00718ffc    52                      push edx
'00718ffd    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00718ffe    ff1580104000            call dword ptr [00401080]
'00719004    8b45b8                  mov eax, dword ptr [ebp-48]
'00719007    50                      push eax
'00719008    8d45dc                  lea eax, dword ptr [ebp-24]
'0071900b    50                      push eax
'0071900c    c745b800000000          mov dword ptr [ebp-48], 00000000

' *** Reference to "__vbaObjSet"
'00719013    ff15b4104000            call dword ptr [004010b4]
Set var_39 = Nothing
'00719019    eb06                    jmp 719021

' *** Reference to "__vbaStrMove"
'0071901b    8b3dd0124000            mov edi, dword ptr [004012d0]

' *** Reference to "__vbaHresultCheckObj"
'00719021    8b3580104000            mov esi, dword ptr [00401080]
'00719027    8b45d8                  mov eax, dword ptr [ebp-28]
'0071902a    8b08                    mov ecx, dword ptr [eax]
'0071902c    8d9538feffff            lea edx, dword ptr [ebp+fffffe38]
'00719032    52                      push edx
'00719033    50                      push eax
'00719034    ff5134                  call dword ptr [ecx+34]
'00719037    dbe2                    fnclex
'00719039    85c0                    test ax, ax
'0071903b    7d0e                    jge 71904b
'0071903d    8b4dd8                  mov ecx, dword ptr [ebp-28]
'00719040    6a34                    push 34
'00719042    6830314300              push 00433130
'00719047    51                      push ecx
'00719048    50                      push eax
'00719049    ffd6                    call esi
'0071904b    6683bd38feffff00        cmp word ptr [ebp+fffffe38], 00
'00719053    8b45dc                  mov eax, dword ptr [ebp-24]
'00719056    0f8566040000            jne 7194c2

Do While (0 = 0)
'0071905c    8b10                    mov edx, dword ptr [eax]
'0071905e    50                      push eax
'0071905f    ff92c0000000            call dword ptr [edx+000000c0]
'00719065    dbe2                    fnclex
'00719067    85c0                    test ax, ax
'00719069    7d11                    jge 71907c
'0071906b    8b4ddc                  mov ecx, dword ptr [ebp-24]
'0071906e    68c0000000              push 000000c0
'00719073    6830314300              push 00433130
'00719078    51                      push ecx
'00719079    50                      push eax
'0071907a    ffd6                    call esi
'0071907c    8b45dc                  mov eax, dword ptr [ebp-24]
'0071907f    8b10                    mov edx, dword ptr [eax]
'00719081    8d4db4                  lea ecx, dword ptr [ebp-4c]
'00719084    51                      push ecx
'00719085    50                      push eax
'00719086    ff92b4000000            call dword ptr [edx+000000b4]
'0071908c    dbe2                    fnclex
'0071908e    85c0                    test ax, ax
'00719090    7d11                    jge 7190a3
'00719092    8b55dc                  mov edx, dword ptr [ebp-24]
'00719095    68b4000000              push 000000b4
'0071909a    6830314300              push 00433130
'0071909f    52                      push edx
'007190a0    50                      push eax
'007190a1    ffd6                    call esi
'007190a3    8b45b4                  mov eax, dword ptr [ebp-4c]
'007190a6    8d5db0                  lea ebx, dword ptr [ebp-50]
'007190a9    53                      push ebx
'007190aa    83ec10                  sub esp, 10
'007190ad    ba02000000              mov edx, 00000002
'007190b2    8bdc                    mov ebx, esp
'007190b4    8913                    mov dword ptr [ebx], edx
'007190b6    8995ecfeffff            mov dword ptr [ebp+fffffeec], edx
'007190bc    8b95f0feffff            mov edx, dword ptr [ebp+fffffef0]
'007190c2    33c9                    xor ecx, ecx
'007190c4    895304                  mov dword ptr [ebx+04], edx
'007190c7    898df4feffff            mov dword ptr [ebp+fffffef4], ecx
'007190cd    8b38                    mov edi, dword ptr [eax]
'007190cf    894b08                  mov dword ptr [ebx+08], ecx
'007190d2    8b8df8feffff            mov ecx, dword ptr [ebp+fffffef8]
'007190d8    50                      push eax
'007190d9    898530feffff            mov dword ptr [ebp+fffffe30], eax
'007190df    894b0c                  mov dword ptr [ebx+0c], ecx
'007190e2    ff5730                  call dword ptr [edi+30]
'007190e5    dbe2                    fnclex
'007190e7    85c0                    test ax, ax
'007190e9    7d11                    jge 7190fc
'007190eb    8b9530feffff            mov edx, dword ptr [ebp+fffffe30]
'007190f1    6a30                    push 30
'007190f3    68d8304300              push 004330d8
'007190f8    52                      push edx
'007190f9    50                      push eax
'007190fa    ffd6                    call esi
'007190fc    668b55e8                mov dx, word ptr [ebp-18]
'00719100    83ec10                  sub esp, 10
'00719103    8bdc                    mov ebx, esp
'00719105    8b75b0                  mov esi, dword ptr [ebp-50]
'00719108    33c0                    xor eax, eax
    var_num1 = Empty
'0071910a    8b3e                    mov edi, dword ptr [esi]
'0071910c    b903000000              mov ecx, 00000003
'00719111    890b                    mov dword ptr [ebx], ecx
'00719113    8b8d50ffffff            mov ecx, dword ptr [ebp+ffffff50]
'00719119    894b04                  mov dword ptr [ebx+04], ecx
'0071911c    894308                  mov dword ptr [ebx+08], eax
'0071911f    898554ffffff            mov dword ptr [ebp+ffffff54], eax
'00719125    8b8558ffffff            mov eax, dword ptr [ebp+ffffff58]
'0071912b    89430c                  mov dword ptr [ebx+0c], eax
'0071912e    8b5d08                  mov ebx, dword ptr [ebp+08]
'00719131    83ec10                  sub esp, 10
'00719134    66899534ffffff          mov word ptr [ebp+ffffff34], dx
'0071913b    8b8534ffffff            mov eax, dword ptr [ebp+ffffff34]
'00719141    8bcc                    mov ecx, esp
'00719143    ba02000000              mov edx, 00000002
'00719148    8911                    mov dword ptr [ecx], edx
'0071914a    8b9530ffffff            mov edx, dword ptr [ebp+ffffff30]
'00719150    895104                  mov dword ptr [ecx+04], edx
'00719153    8b9538ffffff            mov edx, dword ptr [ebp+ffffff38]
'00719159    894108                  mov dword ptr [ecx+08], eax
'0071915c    89510c                  mov dword ptr [ecx+0c], edx
'0071915f    8b9510ffffff            mov edx, dword ptr [ebp+ffffff10]
'00719165    83ec10                  sub esp, 10
'00719168    8bcc                    mov ecx, esp
'0071916a    b802000000              mov eax, 00000002
'0071916f    8901                    mov dword ptr [ecx], eax
'00719171    895104                  mov dword ptr [ecx+04], edx
'00719174    b801000000              mov eax, 00000001
'00719179    894108                  mov dword ptr [ecx+08], eax
'0071917c    8b8518ffffff            mov eax, dword ptr [ebp+ffffff18]
'00719182    6a03                    push 03
'00719184    689e000000              push 0000009e
'00719189    89410c                  mov dword ptr [ecx+0c], eax
'0071918c    8b0b                    mov ecx, dword ptr [ebx]
'0071918e    53                      push ebx
'0071918f    ff9110030000            call dword ptr [ecx+00000310]
'00719195    50                      push eax
'00719196    8d55b8                  lea edx, dword ptr [ebp-48]
'00719199    52                      push edx

' *** Reference to "__vbaObjSet"
'0071919a    ff15b4104000            call dword ptr [004010b4]
    Set var_61 = Nothing
'007191a0    50                      push eax
'007191a1    8d459c                  lea eax, dword ptr [ebp-64]
'007191a4    50                      push eax

' *** Reference to "__vbaLateIdCallLd"
'007191a5    ff157c114000            call dword ptr [0040117c]
    var_51 = var_61.UNK_-256 - 12_158
'007191ab    8b10                    mov edx, dword ptr [eax]
'007191ad    83c430                  add esp, 30
'007191b0    8bcc                    mov ecx, esp
'007191b2    8911                    mov dword ptr [ecx], edx
'007191b4    8b5004                  mov edx, dword ptr [eax+04]
'007191b7    895104                  mov dword ptr [ecx+04], edx
'007191ba    8b5008                  mov edx, dword ptr [eax+08]
'007191bd    8b400c                  mov eax, dword ptr [eax+0c]
'007191c0    895108                  mov dword ptr [ecx+08], edx
'007191c3    56                      push esi
'007191c4    89410c                  mov dword ptr [ecx+0c], eax
'007191c7    ff5748                  call dword ptr [edi+48]
'007191ca    dbe2                    fnclex
'007191cc    85c0                    test ax, ax
'007191ce    7d0f                    jge 7191df
'007191d0    6a48                    push 48
'007191d2    6880304300              push 00433080
'007191d7    56                      push esi
'007191d8    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'007191d9    ff1580104000            call dword ptr [00401080]
'007191df    8d4db0                  lea ecx, dword ptr [ebp-50]
'007191e2    51                      push ecx
'007191e3    8d55b4                  lea edx, dword ptr [ebp-4c]
'007191e6    52                      push edx
'007191e7    8d45b8                  lea eax, dword ptr [ebp-48]
'007191ea    50                      push eax
'007191eb    6a03                    push 03

' *** Reference to "__vbaFreeObjList"
'007191ed    ff154c104000            call dword ptr [0040104c]
    '#Cleanup( var_61, var_62, var_6)
'007191f3    83c410                  add esp, 10
'007191f6    8d4d9c                  lea ecx, dword ptr [ebp-64]

' *** Reference to "__vbaFreeVar"
'007191f9    ff1528104000            call dword ptr [00401028]
    '#Cleanup(var_51)
'007191ff    bf01000000              mov edi, 00000001
'00719204    b80c000000              mov eax, 0000000c
'00719209    663bf8                  cmp di, ax
'0071920c    897de0                  mov dword ptr [ebp-20], edi
'0071920f    0f8f4a020000            jg 71945f
    
    Do While (    1 <= 12)
'00719215    8b45d8                  mov eax, dword ptr [ebp-28]
'00719218    8b08                    mov ecx, dword ptr [eax]
'0071921a    8d55b8                  lea edx, dword ptr [ebp-48]
'0071921d    52                      push edx
'0071921e    50                      push eax
'0071921f    ff91b4000000            call dword ptr [ecx+000000b4]
'00719225    dbe2                    fnclex
'00719227    85c0                    test ax, ax
'00719229    7d15                    jge 719240
'0071922b    8b4dd8                  mov ecx, dword ptr [ebp-28]
'0071922e    68b4000000              push 000000b4
'00719233    6830314300              push 00433130
'00719238    51                      push ecx
'00719239    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0071923a    ff1580104000            call dword ptr [00401080]
'00719240    8b45b8                  mov eax, dword ptr [ebp-48]
'00719243    8b10                    mov edx, dword ptr [eax]
'00719245    6689bd54ffffff          mov word ptr [ebp+ffffff54], di
'0071924c    8d7db4                  lea edi, dword ptr [ebp-4c]
'0071924f    57                      push edi
'00719250    83ec10                  sub esp, 10
'00719253    8bfc                    mov edi, esp
'00719255    b902000000              mov ecx, 00000002
'0071925a    890f                    mov dword ptr [edi], ecx
'0071925c    8b8d50ffffff            mov ecx, dword ptr [ebp+ffffff50]
'00719262    894f04                  mov dword ptr [edi+04], ecx
'00719265    8b8d54ffffff            mov ecx, dword ptr [ebp+ffffff54]
'0071926b    894f08                  mov dword ptr [edi+08], ecx
'0071926e    8b8d58ffffff            mov ecx, dword ptr [ebp+ffffff58]
'00719274    50                      push eax
'00719275    8bf0                    mov esi, eax
'00719277    894f0c                  mov dword ptr [edi+0c], ecx
'0071927a    ff5230                  call dword ptr [edx+30]
'0071927d    dbe2                    fnclex
'0071927f    85c0                    test ax, ax
'00719281    7d0f                    jge 719292
'00719283    6a30                    push 30
'00719285    68d8304300              push 004330d8
'0071928a    56                      push esi
'0071928b    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0071928c    ff1580104000            call dword ptr [00401080]
'00719292    8b45b4                  mov eax, dword ptr [ebp-4c]
'00719295    8d559c                  lea edx, dword ptr [ebp-64]
'00719298    52                      push edx
'00719299    c745b400000000          mov dword ptr [ebp-4c], 00000000
'007192a0    8945a4                  mov dword ptr [ebp-5c], eax
'007192a3    c7459c09000000          mov dword ptr [ebp-64], 00000009

' *** Reference to "rtcIsNull"
'007192aa    ff1540114000            call dword ptr [00401140]
'007192b0    33f6                    xor esi, esi
    var_num8 = Empty
'007192b2    668bf0                  mov si, ax
'007192b5    8d4db8                  lea ecx, dword ptr [ebp-48]
'007192b8    f7d6                    not esi

' *** Reference to "__vbaFreeObj"
'007192ba    ff1508134000            call dword ptr [00401308]
    '#Cleanup(var_61)
'007192c0    8d4d9c                  lea ecx, dword ptr [ebp-64]

' *** Reference to "__vbaFreeVar"
'007192c3    ff1528104000            call dword ptr [00401028]
    '#Cleanup(var_51)
'007192c9    6685f6                  test esi, esi
'007192cc    0f8477010000            je 719449
    
    If (    Not (IsNull(var_62))) Then
'007192d2    8b45dc                  mov eax, dword ptr [ebp-24]
'007192d5    8b08                    mov ecx, dword ptr [eax]
'007192d7    8d55b0                  lea edx, dword ptr [ebp-50]
'007192da    52                      push edx
'007192db    50                      push eax
'007192dc    ff91b4000000            call dword ptr [ecx+000000b4]
'007192e2    dbe2                    fnclex
'007192e4    85c0                    test ax, ax
'007192e6    7d15                    jge 7192fd
'007192e8    8b4ddc                  mov ecx, dword ptr [ebp-24]
'007192eb    68b4000000              push 000000b4
'007192f0    6830314300              push 00433130
'007192f5    51                      push ecx
'007192f6    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'007192f7    ff1580104000            call dword ptr [00401080]
'007192fd    668b55e0                mov dx, word ptr [ebp-20]
'00719301    8b45b0                  mov eax, dword ptr [ebp-50]
'00719304    8d7dac                  lea edi, dword ptr [ebp-54]
'00719307    57                      push edi
'00719308    83ec10                  sub esp, 10
'0071930b    8bfc                    mov edi, esp
'0071930d    b902000000              mov ecx, 00000002
'00719312    890f                    mov dword ptr [edi], ecx
'00719314    8b8d40ffffff            mov ecx, dword ptr [ebp+ffffff40]
'0071931a    894f04                  mov dword ptr [edi+04], ecx
'0071931d    66899544ffffff          mov word ptr [ebp+ffffff44], dx
'00719324    8b8d44ffffff            mov ecx, dword ptr [ebp+ffffff44]
'0071932a    8b10                    mov edx, dword ptr [eax]
'0071932c    894f08                  mov dword ptr [edi+08], ecx
'0071932f    8b8d48ffffff            mov ecx, dword ptr [ebp+ffffff48]
'00719335    50                      push eax
'00719336    8bf0                    mov esi, eax
'00719338    894f0c                  mov dword ptr [edi+0c], ecx
'0071933b    ff5230                  call dword ptr [edx+30]
'0071933e    dbe2                    fnclex
'00719340    85c0                    test ax, ax
'00719342    7d0f                    jge 719353
'00719344    6a30                    push 30
'00719346    68d8304300              push 004330d8
'0071934b    56                      push esi
'0071934c    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0071934d    ff1580104000            call dword ptr [00401080]
'00719353    8b55ac                  mov edx, dword ptr [ebp-54]
'00719356    8b45d8                  mov eax, dword ptr [ebp-28]
'00719359    8b08                    mov ecx, dword ptr [eax]
'0071935b    89951cfeffff            mov dword ptr [ebp+fffffe1c], edx
'00719361    8d55b8                  lea edx, dword ptr [ebp-48]
'00719364    52                      push edx
'00719365    50                      push eax
'00719366    ff91b4000000            call dword ptr [ecx+000000b4]
'0071936c    dbe2                    fnclex
'0071936e    85c0                    test ax, ax
'00719370    7d15                    jge 719387
'00719372    8b4dd8                  mov ecx, dword ptr [ebp-28]
'00719375    68b4000000              push 000000b4
'0071937a    6830314300              push 00433130
'0071937f    51                      push ecx
'00719380    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00719381    ff1580104000            call dword ptr [00401080]
'00719387    668b55e0                mov dx, word ptr [ebp-20]
'0071938b    8b45b8                  mov eax, dword ptr [ebp-48]
'0071938e    8d7db4                  lea edi, dword ptr [ebp-4c]
'00719391    57                      push edi
'00719392    83ec10                  sub esp, 10
'00719395    8bfc                    mov edi, esp
'00719397    b902000000              mov ecx, 00000002
'0071939c    890f                    mov dword ptr [edi], ecx
'0071939e    8b8d50ffffff            mov ecx, dword ptr [ebp+ffffff50]
'007193a4    894f04                  mov dword ptr [edi+04], ecx
'007193a7    66899554ffffff          mov word ptr [ebp+ffffff54], dx
'007193ae    8b8d54ffffff            mov ecx, dword ptr [ebp+ffffff54]
'007193b4    8b10                    mov edx, dword ptr [eax]
'007193b6    894f08                  mov dword ptr [edi+08], ecx
'007193b9    8b8d58ffffff            mov ecx, dword ptr [ebp+ffffff58]
'007193bf    50                      push eax
'007193c0    8bf0                    mov esi, eax
'007193c2    894f0c                  mov dword ptr [edi+0c], ecx
'007193c5    ff5230                  call dword ptr [edx+30]
'007193c8    dbe2                    fnclex
'007193ca    85c0                    test ax, ax
'007193cc    7d0f                    jge 7193dd
'007193ce    6a30                    push 30
'007193d0    68d8304300              push 004330d8
'007193d5    56                      push esi
'007193d6    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'007193d7    ff1580104000            call dword ptr [00401080]
'007193dd    8b45b4                  mov eax, dword ptr [ebp-4c]
'007193e0    8bb51cfeffff            mov esi, dword ptr [ebp+fffffe1c]
'007193e6    83ec10                  sub esp, 10
'007193e9    b909000000              mov ecx, 00000009
'007193ee    8bfc                    mov edi, esp
'007193f0    890f                    mov dword ptr [edi], ecx
'007193f2    894d9c                  mov dword ptr [ebp-64], ecx
'007193f5    8b4da0                  mov ecx, dword ptr [ebp-60]
'007193f8    894f04                  mov dword ptr [edi+04], ecx
'007193fb    8945a4                  mov dword ptr [ebp-5c], eax
'007193fe    894708                  mov dword ptr [edi+08], eax
'00719401    8b45a8                  mov eax, dword ptr [ebp-58]
'00719404    c745b400000000          mov dword ptr [ebp-4c], 00000000
'0071940b    8b16                    mov edx, dword ptr [esi]
'0071940d    56                      push esi
'0071940e    89470c                  mov dword ptr [edi+0c], eax
'00719411    ff5248                  call dword ptr [edx+48]
'00719414    dbe2                    fnclex
'00719416    85c0                    test ax, ax
'00719418    7d0f                    jge 719429
'0071941a    6a48                    push 48
'0071941c    6880304300              push 00433080
'00719421    56                      push esi
'00719422    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00719423    ff1580104000            call dword ptr [00401080]
'00719429    8d4dac                  lea ecx, dword ptr [ebp-54]
'0071942c    51                      push ecx
'0071942d    8d55b0                  lea edx, dword ptr [ebp-50]
'00719430    52                      push edx
'00719431    8d45b8                  lea eax, dword ptr [ebp-48]
'00719434    50                      push eax
'00719435    6a03                    push 03

' *** Reference to "__vbaFreeObjList"
'00719437    ff154c104000            call dword ptr [0040104c]
    '#Cleanup( var_61, var_6, var_50)
'0071943d    83c410                  add esp, 10
'00719440    8d4d9c                  lea ecx, dword ptr [ebp-64]

' *** Reference to "__vbaFreeVar"
'00719443    ff1528104000            call dword ptr [00401028]
    '#Cleanup(var_51)
End If
'00719449    b801000000              mov eax, 00000001
'0071944e    660345e0                add ax, word ptr [ebp-20]
var_num1 = 1 + 1
'00719452    0f801b140000            jo 71a873
'00719458    8bf8                    mov edi, eax
'0071945a    e9a5fdffff              jmp 719204

'ERROR: Two many next close:
Loop
'0071945f    8b45dc                  mov eax, dword ptr [ebp-24]
'00719462    8b08                    mov ecx, dword ptr [eax]
'00719464    6a00                    push 00
'00719466    6a01                    push 01
'00719468    50                      push eax
'00719469    ff9164010000            call dword ptr [ecx+00000164]
'0071946f    dbe2                    fnclex
'00719471    85c0                    test ax, ax
'00719473    7d15                    jge 71948a
'00719475    8b55dc                  mov edx, dword ptr [ebp-24]
'00719478    6864010000              push 00000164
'0071947d    6830314300              push 00433130
'00719482    52                      push edx
'00719483    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00719484    ff1580104000            call dword ptr [00401080]
'0071948a    8b45d8                  mov eax, dword ptr [ebp-28]
'0071948d    8b08                    mov ecx, dword ptr [eax]
'0071948f    50                      push eax
'00719490    ff91ec000000            call dword ptr [ecx+000000ec]
'00719496    dbe2                    fnclex
'00719498    85c0                    test ax, ax
'0071949a    0f8d7bfbffff            jge 71901b
'007194a0    8b55d8                  mov edx, dword ptr [ebp-28]

' *** Reference to "__vbaHresultCheckObj"
'007194a3    8b3580104000            mov esi, dword ptr [00401080]
'007194a9    68ec000000              push 000000ec
'007194ae    6830314300              push 00433130
'007194b3    52                      push edx
'007194b4    50                      push eax
'007194b5    ffd6                    call esi

' *** Reference to "__vbaStrMove"
'007194b7    8b3dd0124000            mov edi, dword ptr [004012d0]
'007194bd    e965fbffff              jmp 719027

'ERROR: Two many next close:
Loop
'007194c2    8b08                    mov ecx, dword ptr [eax]
'007194c4    50                      push eax
'007194c5    ff91c4000000            call dword ptr [ecx+000000c4]
'007194cb    dbe2                    fnclex
'007194cd    85c0                    test ax, ax
'007194cf    7d11                    jge 7194e2
'007194d1    8b55dc                  mov edx, dword ptr [ebp-24]
'007194d4    68c4000000              push 000000c4
'007194d9    6830314300              push 00433130
'007194de    52                      push edx
'007194df    50                      push eax
'007194e0    ffd6                    call esi
'007194e2    8b45d8                  mov eax, dword ptr [ebp-28]
'007194e5    8b08                    mov ecx, dword ptr [eax]
'007194e7    50                      push eax
'007194e8    ff91c4000000            call dword ptr [ecx+000000c4]
'007194ee    dbe2                    fnclex
'007194f0    85c0                    test ax, ax
'007194f2    7d11                    jge 719505
'007194f4    8b55d8                  mov edx, dword ptr [ebp-28]
'007194f7    68c4000000              push 000000c4
'007194fc    6830314300              push 00433130
'00719501    52                      push edx
'00719502    50                      push eax
'00719503    ffd6                    call esi
'00719505    668b55e8                mov dx, word ptr [ebp-18]
'00719509    83ec10                  sub esp, 10
'0071950c    8bf4                    mov esi, esp
'0071950e    33c0                    xor eax, eax
var_num1 = Empty
'00719510    b903000000              mov ecx, 00000003
'00719515    890e                    mov dword ptr [esi], ecx
'00719517    8b8d50ffffff            mov ecx, dword ptr [ebp+ffffff50]
'0071951d    894e04                  mov dword ptr [esi+04], ecx
'00719520    894608                  mov dword ptr [esi+08], eax
'00719523    8b8558ffffff            mov eax, dword ptr [ebp+ffffff58]
'00719529    89460c                  mov dword ptr [esi+0c], eax
'0071952c    83ec10                  sub esp, 10
'0071952f    66899534ffffff          mov word ptr [ebp+ffffff34], dx
'00719536    8b8534ffffff            mov eax, dword ptr [ebp+ffffff34]
'0071953c    8bcc                    mov ecx, esp
'0071953e    ba02000000              mov edx, 00000002
'00719543    8911                    mov dword ptr [ecx], edx
'00719545    8b9530ffffff            mov edx, dword ptr [ebp+ffffff30]
'0071954b    895104                  mov dword ptr [ecx+04], edx
'0071954e    8b9538ffffff            mov edx, dword ptr [ebp+ffffff38]
'00719554    894108                  mov dword ptr [ecx+08], eax
'00719557    89510c                  mov dword ptr [ecx+0c], edx
'0071955a    8b9510ffffff            mov edx, dword ptr [ebp+ffffff10]
'00719560    83ec10                  sub esp, 10
'00719563    8bcc                    mov ecx, esp
'00719565    b802000000              mov eax, 00000002
'0071956a    8901                    mov dword ptr [ecx], eax
'0071956c    895104                  mov dword ptr [ecx+04], edx
'0071956f    33c0                    xor eax, eax
var_num1 = Empty
'00719571    894108                  mov dword ptr [ecx+08], eax
'00719574    8b8518ffffff            mov eax, dword ptr [ebp+ffffff18]
'0071957a    6a03                    push 03
'0071957c    689e000000              push 0000009e
'00719581    89410c                  mov dword ptr [ecx+0c], eax
'00719584    8b0b                    mov ecx, dword ptr [ebx]
'00719586    53                      push ebx
'00719587    ff9110030000            call dword ptr [ecx+00000310]
'0071958d    50                      push eax
'0071958e    8d55b8                  lea edx, dword ptr [ebp-48]
'00719591    52                      push edx

' *** Reference to "__vbaObjSet"
'00719592    ff15b4104000            call dword ptr [004010b4]
Set var_61 = Nothing
'00719598    50                      push eax
'00719599    8d459c                  lea eax, dword ptr [ebp-64]
'0071959c    50                      push eax

' *** Reference to "__vbaLateIdCallLd"
'0071959d    ff157c114000            call dword ptr [0040117c]
var_51 = var_61.UNK_-256 - 12_158
'007195a3    83c440                  add esp, 40
'007195a6    50                      push eax

' *** Reference to "__vbaStrVarMove"
'007195a7    ff153c104000            call dword ptr [0040103c]
'007195ad    8bd0                    mov edx, eax
'007195af    8d4dcc                  lea ecx, dword ptr [ebp-34]
'007195b2    ffd7                    call edi
'007195b4    8d4dcc                  lea ecx, dword ptr [ebp-34]
'007195b7    51                      push ecx
'007195b8    e833a6ddff              call 4f3bf0
Call sub_4F3BF0()
'007195bd    8bd0                    mov edx, eax
'007195bf    8d4dbc                  lea ecx, dword ptr [ebp-44]
'007195c2    ffd7                    call edi
'007195c4    8b55bc                  mov edx, dword ptr [ebp-44]
'007195c7    899508fdffff            mov dword ptr [ebp+fffffd08], edx
'007195cd    8b5338                  mov edx, dword ptr [ebx+38]
'007195d0    c785f4feffff02000000    mov dword ptr [ebp+fffffef4], 00000002
'007195da    c785ecfeffff03000000    mov dword ptr [ebp+fffffeec], 00000003
'007195e4    c745bc00000000          mov dword ptr [ebp-44], 00000000
'007195eb    8b32                    mov esi, dword ptr [edx]
'007195ed    8d55b4                  lea edx, dword ptr [ebp-4c]
'007195f0    52                      push edx
'007195f1    83ec10                  sub esp, 10
'007195f4    b90a000000              mov ecx, 0000000a
'007195f9    8bd4                    mov edx, esp
'007195fb    890a                    mov dword ptr [edx], ecx
'007195fd    898ddcfeffff            mov dword ptr [ebp+fffffedc], ecx
'00719603    8b8dd0feffff            mov ecx, dword ptr [ebp+fffffed0]
'00719609    894a04                  mov dword ptr [edx+04], ecx
'0071960c    b804000280              mov eax, 80020004
'00719611    894208                  mov dword ptr [edx+08], eax
'00719614    8985e4feffff            mov dword ptr [ebp+fffffee4], eax
'0071961a    8b85d8feffff            mov eax, dword ptr [ebp+fffffed8]
'00719620    83ec10                  sub esp, 10
'00719623    89420c                  mov dword ptr [edx+0c], eax
'00719626    8b95dcfeffff            mov edx, dword ptr [ebp+fffffedc]
'0071962c    8bcc                    mov ecx, esp
'0071962e    8b85e0feffff            mov eax, dword ptr [ebp+fffffee0]
'00719634    8911                    mov dword ptr [ecx], edx
'00719636    8b95e4feffff            mov edx, dword ptr [ebp+fffffee4]
'0071963c    894104                  mov dword ptr [ecx+04], eax
'0071963f    8b85e8feffff            mov eax, dword ptr [ebp+fffffee8]
'00719645    895108                  mov dword ptr [ecx+08], edx
'00719648    8b95ecfeffff            mov edx, dword ptr [ebp+fffffeec]
'0071964e    89410c                  mov dword ptr [ecx+0c], eax
'00719651    8b85f0feffff            mov eax, dword ptr [ebp+fffffef0]
'00719657    83ec10                  sub esp, 10
'0071965a    8bcc                    mov ecx, esp
'0071965c    8911                    mov dword ptr [ecx], edx
'0071965e    8b95f4feffff            mov edx, dword ptr [ebp+fffffef4]
'00719664    894104                  mov dword ptr [ecx+04], eax
'00719667    8b85f8feffff            mov eax, dword ptr [ebp+fffffef8]
'0071966d    895108                  mov dword ptr [ecx+08], edx
'00719670    8b9508fdffff            mov edx, dword ptr [ebp+fffffd08]
'00719676    89410c                  mov dword ptr [ecx+0c], eax
'00719679    685c6c4500              push 00456c5c
'0071967e    8d4dc8                  lea ecx, dword ptr [ebp-38]
'00719681    ffd7                    call edi
'00719683    50                      push eax

' *** Reference to "__vbaStrCat"
'00719684    ff1570104000            call dword ptr [00401070]
var_1632 = ("select * from SortPersonnage where NomPerso='") & (var_51)
'0071968a    8bd0                    mov edx, eax
'0071968c    8d4dc4                  lea ecx, dword ptr [ebp-3c]
'0071968f    ffd7                    call edi
'00719691    50                      push eax
'00719692    6854a44300              push 0043a454

' *** Reference to "__vbaStrCat"
'00719697    ff1570104000            call dword ptr [00401070]
var_2120 = (var_1632) & ("'")
'0071969d    8bd0                    mov edx, eax
'0071969f    8d4dc0                  lea ecx, dword ptr [ebp-40]
'007196a2    ffd7                    call edi
'007196a4    8b4b38                  mov ecx, dword ptr [ebx+38]
'007196a7    50                      push eax
'007196a8    51                      push ecx
'007196a9    ff96bc000000            call dword ptr [esi+000000bc]
'007196af    dbe2                    fnclex
'007196b1    85c0                    test ax, ax
'007196b3    7d15                    jge 7196ca
'007196b5    8b5338                  mov edx, dword ptr [ebx+38]
'007196b8    68bc000000              push 000000bc
'007196bd    68301f4300              push 00431f30
'007196c2    52                      push edx
'007196c3    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'007196c4    ff1580104000            call dword ptr [00401080]
'007196ca    8b45b4                  mov eax, dword ptr [ebp-4c]
'007196cd    50                      push eax
'007196ce    8d45d8                  lea eax, dword ptr [ebp-28]
'007196d1    50                      push eax
'007196d2    c745b400000000          mov dword ptr [ebp-4c], 00000000

' *** Reference to "__vbaObjSet"
'007196d9    ff15b4104000            call dword ptr [004010b4]
Set var_45 = Nothing
'007196df    8d4dbc                  lea ecx, dword ptr [ebp-44]
'007196e2    51                      push ecx
'007196e3    8d55c0                  lea edx, dword ptr [ebp-40]
'007196e6    52                      push edx
'007196e7    8d45c4                  lea eax, dword ptr [ebp-3c]
'007196ea    50                      push eax
'007196eb    8d4dc8                  lea ecx, dword ptr [ebp-38]
'007196ee    51                      push ecx
'007196ef    8d55cc                  lea edx, dword ptr [ebp-34]
'007196f2    52                      push edx
'007196f3    6a05                    push 05

' *** Reference to "__vbaFreeStrList"
'007196f5    ff155c124000            call dword ptr [0040125c]
'#Cleanup( -4860, -4860, -4864, -4868, var_58)
'007196fb    83c418                  add esp, 18
'007196fe    8d4db8                  lea ecx, dword ptr [ebp-48]

' *** Reference to "__vbaFreeObj"
'00719701    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_61)
'00719707    8d4d9c                  lea ecx, dword ptr [ebp-64]

' *** Reference to "__vbaFreeVar"
'0071970a    ff1528104000            call dword ptr [00401028]
'#Cleanup(var_51)
'00719710    8d5db8                  lea ebx, dword ptr [ebp-48]
'00719713    53                      push ebx
'00719714    83ec10                  sub esp, 10
'00719717    8bdc                    mov ebx, esp
'00719719    b90a000000              mov ecx, 0000000a
'0071971e    890b                    mov dword ptr [ebx], ecx
'00719720    898d3cffffff            mov dword ptr [ebp+ffffff3c], ecx
'00719726    8b8d30ffffff            mov ecx, dword ptr [ebp+ffffff30]
'0071972c    894b04                  mov dword ptr [ebx+04], ecx
'0071972f    8b3548b07200            mov esi, dword ptr [0072b048]
'00719735    b804000280              mov eax, 80020004
'0071973a    894308                  mov dword ptr [ebx+08], eax
'0071973d    8bd0                    mov edx, eax
'0071973f    898534ffffff            mov dword ptr [ebp+ffffff34], eax
'00719745    8b8538ffffff            mov eax, dword ptr [ebp+ffffff38]
'0071974b    89430c                  mov dword ptr [ebx+0c], eax
'0071974e    8b853cffffff            mov eax, dword ptr [ebp+ffffff3c]
'00719754    8b36                    mov esi, dword ptr [esi]
'00719756    83ec10                  sub esp, 10
'00719759    8bcc                    mov ecx, esp
'0071975b    8901                    mov dword ptr [ecx], eax
'0071975d    8b8540ffffff            mov eax, dword ptr [ebp+ffffff40]
'00719763    894104                  mov dword ptr [ecx+04], eax
'00719766    895108                  mov dword ptr [ecx+08], edx
'00719769    899544ffffff            mov dword ptr [ebp+ffffff44], edx
'0071976f    8b9548ffffff            mov edx, dword ptr [ebp+ffffff48]
'00719775    89510c                  mov dword ptr [ecx+0c], edx
'00719778    8b9550ffffff            mov edx, dword ptr [ebp+ffffff50]
'0071977e    83ec10                  sub esp, 10
'00719781    8bcc                    mov ecx, esp
'00719783    b803000000              mov eax, 00000003
'00719788    8901                    mov dword ptr [ecx], eax
'0071978a    895104                  mov dword ptr [ecx+04], edx
'0071978d    b801000000              mov eax, 00000001
'00719792    894108                  mov dword ptr [ecx+08], eax
'00719795    8b8558ffffff            mov eax, dword ptr [ebp+ffffff58]
'0071979b    89410c                  mov dword ptr [ecx+0c], eax
'0071979e    8b0d48b07200            mov ecx, dword ptr [0072b048]
'007197a4    685caa4300              push 0043aa5c
'007197a9    51                      push ecx
'007197aa    ff96bc000000            call dword ptr [esi+000000bc]
'007197b0    dbe2                    fnclex
'007197b2    85c0                    test ax, ax
'007197b4    7d18                    jge 7197ce
'007197b6    8b1548b07200            mov edx, dword ptr [0072b048]
'007197bc    68bc000000              push 000000bc
'007197c1    68301f4300              push 00431f30
'007197c6    52                      push edx
'007197c7    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'007197c8    ff1580104000            call dword ptr [00401080]
'007197ce    8b45b8                  mov eax, dword ptr [ebp-48]
'007197d1    50                      push eax
'007197d2    8d45dc                  lea eax, dword ptr [ebp-24]
'007197d5    50                      push eax
'007197d6    c745b800000000          mov dword ptr [ebp-48], 00000000

' *** Reference to "__vbaObjSet"
'007197dd    ff15b4104000            call dword ptr [004010b4]
Set var_39 = Nothing
'007197e3    eb06                    jmp 7197eb

' *** Reference to "__vbaStrMove"
'007197e5    8b3dd0124000            mov edi, dword ptr [004012d0]

' *** Reference to "__vbaHresultCheckObj"
'007197eb    8b3580104000            mov esi, dword ptr [00401080]
'007197f1    8b45d8                  mov eax, dword ptr [ebp-28]
'007197f4    8b08                    mov ecx, dword ptr [eax]
'007197f6    8d9538feffff            lea edx, dword ptr [ebp+fffffe38]
'007197fc    52                      push edx
'007197fd    50                      push eax
'007197fe    ff5134                  call dword ptr [ecx+34]
'00719801    dbe2                    fnclex
'00719803    85c0                    test ax, ax
'00719805    7d0e                    jge 719815
'00719807    8b4dd8                  mov ecx, dword ptr [ebp-28]
'0071980a    6a34                    push 34
'0071980c    6830314300              push 00433130
'00719811    51                      push ecx
'00719812    50                      push eax
'00719813    ffd6                    call esi
'00719815    6683bd38feffff00        cmp word ptr [ebp+fffffe38], 00
'0071981d    8b45dc                  mov eax, dword ptr [ebp-24]
'00719820    0f854e040000            jne 719c74

Do While (0 = 0)
'00719826    8b10                    mov edx, dword ptr [eax]
'00719828    50                      push eax
'00719829    ff92c0000000            call dword ptr [edx+000000c0]
'0071982f    dbe2                    fnclex
'00719831    85c0                    test ax, ax
'00719833    7d11                    jge 719846
'00719835    8b4ddc                  mov ecx, dword ptr [ebp-24]
'00719838    68c0000000              push 000000c0
'0071983d    6830314300              push 00433130
'00719842    51                      push ecx
'00719843    50                      push eax
'00719844    ffd6                    call esi
'00719846    8b45dc                  mov eax, dword ptr [ebp-24]
'00719849    8b10                    mov edx, dword ptr [eax]
'0071984b    8d4db4                  lea ecx, dword ptr [ebp-4c]
'0071984e    51                      push ecx
'0071984f    50                      push eax
'00719850    ff92b4000000            call dword ptr [edx+000000b4]
'00719856    dbe2                    fnclex
'00719858    85c0                    test ax, ax
'0071985a    7d11                    jge 71986d
'0071985c    8b55dc                  mov edx, dword ptr [ebp-24]
'0071985f    68b4000000              push 000000b4
'00719864    6830314300              push 00433130
'00719869    52                      push edx
'0071986a    50                      push eax
'0071986b    ffd6                    call esi
'0071986d    8b45b4                  mov eax, dword ptr [ebp-4c]
'00719870    8d5db0                  lea ebx, dword ptr [ebp-50]
'00719873    53                      push ebx
'00719874    83ec10                  sub esp, 10
'00719877    ba02000000              mov edx, 00000002
'0071987c    8bdc                    mov ebx, esp
'0071987e    8913                    mov dword ptr [ebx], edx
'00719880    8995ecfeffff            mov dword ptr [ebp+fffffeec], edx
'00719886    8b95f0feffff            mov edx, dword ptr [ebp+fffffef0]
'0071988c    33c9                    xor ecx, ecx
'0071988e    895304                  mov dword ptr [ebx+04], edx
'00719891    898df4feffff            mov dword ptr [ebp+fffffef4], ecx
'00719897    8b38                    mov edi, dword ptr [eax]
'00719899    894b08                  mov dword ptr [ebx+08], ecx
'0071989c    8b8df8feffff            mov ecx, dword ptr [ebp+fffffef8]
'007198a2    50                      push eax
'007198a3    8bf0                    mov esi, eax
'007198a5    894b0c                  mov dword ptr [ebx+0c], ecx
'007198a8    ff5730                  call dword ptr [edi+30]
'007198ab    dbe2                    fnclex
'007198ad    85c0                    test ax, ax
'007198af    7d0f                    jge 7198c0
'007198b1    6a30                    push 30
'007198b3    68d8304300              push 004330d8
'007198b8    56                      push esi
'007198b9    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'007198ba    ff1580104000            call dword ptr [00401080]
'007198c0    668b55e8                mov dx, word ptr [ebp-18]
'007198c4    83ec10                  sub esp, 10
'007198c7    8bdc                    mov ebx, esp
'007198c9    8b75b0                  mov esi, dword ptr [ebp-50]
'007198cc    33c0                    xor eax, eax
    var_num1 = Empty
'007198ce    b903000000              mov ecx, 00000003
'007198d3    890b                    mov dword ptr [ebx], ecx
'007198d5    8b8d50ffffff            mov ecx, dword ptr [ebp+ffffff50]
'007198db    894b04                  mov dword ptr [ebx+04], ecx
'007198de    894308                  mov dword ptr [ebx+08], eax
'007198e1    8b3e                    mov edi, dword ptr [esi]
'007198e3    898554ffffff            mov dword ptr [ebp+ffffff54], eax
'007198e9    8b8558ffffff            mov eax, dword ptr [ebp+ffffff58]
'007198ef    89430c                  mov dword ptr [ebx+0c], eax
'007198f2    83ec10                  sub esp, 10
'007198f5    66899534ffffff          mov word ptr [ebp+ffffff34], dx
'007198fc    8b8534ffffff            mov eax, dword ptr [ebp+ffffff34]
'00719902    8bcc                    mov ecx, esp
'00719904    ba02000000              mov edx, 00000002
'00719909    8911                    mov dword ptr [ecx], edx
'0071990b    8b9530ffffff            mov edx, dword ptr [ebp+ffffff30]
'00719911    895104                  mov dword ptr [ecx+04], edx
'00719914    8b9538ffffff            mov edx, dword ptr [ebp+ffffff38]
'0071991a    894108                  mov dword ptr [ecx+08], eax
'0071991d    89510c                  mov dword ptr [ecx+0c], edx
'00719920    8b9510ffffff            mov edx, dword ptr [ebp+ffffff10]
'00719926    83ec10                  sub esp, 10
'00719929    8bcc                    mov ecx, esp
'0071992b    b802000000              mov eax, 00000002
'00719930    8901                    mov dword ptr [ecx], eax
'00719932    895104                  mov dword ptr [ecx+04], edx
'00719935    b801000000              mov eax, 00000001
'0071993a    894108                  mov dword ptr [ecx+08], eax
'0071993d    8b8518ffffff            mov eax, dword ptr [ebp+ffffff18]
'00719943    6a03                    push 03
'00719945    89410c                  mov dword ptr [ecx+0c], eax
'00719948    8b4508                  mov eax, dword ptr [ebp+08]
'0071994b    8b08                    mov ecx, dword ptr [eax]
'0071994d    689e000000              push 0000009e
'00719952    50                      push eax
'00719953    ff9110030000            call dword ptr [ecx+00000310]
'00719959    50                      push eax
'0071995a    8d55b8                  lea edx, dword ptr [ebp-48]
'0071995d    52                      push edx

' *** Reference to "__vbaObjSet"
'0071995e    ff15b4104000            call dword ptr [004010b4]
    Set var_61 = Me
'00719964    50                      push eax
'00719965    8d459c                  lea eax, dword ptr [ebp-64]
'00719968    50                      push eax

' *** Reference to "__vbaLateIdCallLd"
'00719969    ff157c114000            call dword ptr [0040117c]
    var_51 = var_61.UNK_frmImport_158
'0071996f    8b10                    mov edx, dword ptr [eax]
'00719971    83c430                  add esp, 30
'00719974    8bcc                    mov ecx, esp
'00719976    8911                    mov dword ptr [ecx], edx
'00719978    8b5004                  mov edx, dword ptr [eax+04]
'0071997b    895104                  mov dword ptr [ecx+04], edx
'0071997e    8b5008                  mov edx, dword ptr [eax+08]
'00719981    8b400c                  mov eax, dword ptr [eax+0c]
'00719984    895108                  mov dword ptr [ecx+08], edx
'00719987    56                      push esi
'00719988    89410c                  mov dword ptr [ecx+0c], eax
'0071998b    ff5748                  call dword ptr [edi+48]
'0071998e    dbe2                    fnclex
'00719990    85c0                    test ax, ax
'00719992    7d0f                    jge 7199a3
'00719994    6a48                    push 48
'00719996    6880304300              push 00433080
'0071999b    56                      push esi
'0071999c    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0071999d    ff1580104000            call dword ptr [00401080]
'007199a3    8d4db0                  lea ecx, dword ptr [ebp-50]
'007199a6    51                      push ecx
'007199a7    8d55b4                  lea edx, dword ptr [ebp-4c]
'007199aa    52                      push edx
'007199ab    8d45b8                  lea eax, dword ptr [ebp-48]
'007199ae    50                      push eax
'007199af    6a03                    push 03

' *** Reference to "__vbaFreeObjList"
'007199b1    ff154c104000            call dword ptr [0040104c]
    '#Cleanup( var_61, var_62, var_6)
'007199b7    83c410                  add esp, 10
'007199ba    8d4d9c                  lea ecx, dword ptr [ebp-64]

' *** Reference to "__vbaFreeVar"
'007199bd    ff1528104000            call dword ptr [00401028]
    '#Cleanup(var_51)
'007199c3    bb01000000              mov ebx, 00000001
'007199c8    b803000000              mov eax, 00000003
'007199cd    663bd8                  cmp bx, ax
'007199d0    895de0                  mov dword ptr [ebp-20], ebx
'007199d3    0f8f38020000            jg 719c11
    
    Do While (    1 <= 3)
'007199d9    8b45d8                  mov eax, dword ptr [ebp-28]
'007199dc    8b08                    mov ecx, dword ptr [eax]
'007199de    8d55b8                  lea edx, dword ptr [ebp-48]
'007199e1    52                      push edx
'007199e2    50                      push eax
'007199e3    ff91b4000000            call dword ptr [ecx+000000b4]
'007199e9    dbe2                    fnclex
'007199eb    85c0                    test ax, ax
'007199ed    7d15                    jge 719a04
'007199ef    8b4dd8                  mov ecx, dword ptr [ebp-28]
'007199f2    68b4000000              push 000000b4
'007199f7    6830314300              push 00433130
'007199fc    51                      push ecx
'007199fd    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'007199fe    ff1580104000            call dword ptr [00401080]
'00719a04    8b45b8                  mov eax, dword ptr [ebp-48]
'00719a07    8b10                    mov edx, dword ptr [eax]
'00719a09    8d7db4                  lea edi, dword ptr [ebp-4c]
'00719a0c    57                      push edi
'00719a0d    83ec10                  sub esp, 10
'00719a10    8bfc                    mov edi, esp
'00719a12    b902000000              mov ecx, 00000002
'00719a17    890f                    mov dword ptr [edi], ecx
'00719a19    8b8d50ffffff            mov ecx, dword ptr [ebp+ffffff50]
'00719a1f    894f04                  mov dword ptr [edi+04], ecx
'00719a22    66899d54ffffff          mov word ptr [ebp+ffffff54], bx
'00719a29    8b8d54ffffff            mov ecx, dword ptr [ebp+ffffff54]
'00719a2f    894f08                  mov dword ptr [edi+08], ecx
'00719a32    8b8d58ffffff            mov ecx, dword ptr [ebp+ffffff58]
'00719a38    50                      push eax
'00719a39    8bf0                    mov esi, eax
'00719a3b    894f0c                  mov dword ptr [edi+0c], ecx
'00719a3e    ff5230                  call dword ptr [edx+30]
'00719a41    dbe2                    fnclex
'00719a43    85c0                    test ax, ax
'00719a45    7d0f                    jge 719a56
'00719a47    6a30                    push 30
'00719a49    68d8304300              push 004330d8
'00719a4e    56                      push esi
'00719a4f    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00719a50    ff1580104000            call dword ptr [00401080]
'00719a56    8b45b4                  mov eax, dword ptr [ebp-4c]
'00719a59    8d559c                  lea edx, dword ptr [ebp-64]
'00719a5c    52                      push edx
'00719a5d    c745b400000000          mov dword ptr [ebp-4c], 00000000
'00719a64    8945a4                  mov dword ptr [ebp-5c], eax
'00719a67    c7459c09000000          mov dword ptr [ebp-64], 00000009

' *** Reference to "rtcIsNull"
'00719a6e    ff1540114000            call dword ptr [00401140]
'00719a74    33f6                    xor esi, esi
    var_num8 = Empty
'00719a76    668bf0                  mov si, ax
'00719a79    8d4db8                  lea ecx, dword ptr [ebp-48]
'00719a7c    f7d6                    not esi

' *** Reference to "__vbaFreeObj"
'00719a7e    ff1508134000            call dword ptr [00401308]
    '#Cleanup(var_61)
'00719a84    8d4d9c                  lea ecx, dword ptr [ebp-64]

' *** Reference to "__vbaFreeVar"
'00719a87    ff1528104000            call dword ptr [00401028]
    '#Cleanup(var_51)
'00719a8d    6685f6                  test esi, esi
'00719a90    0f8466010000            je 719bfc
    
    If (    Not (IsNull(var_62))) Then
'00719a96    8b45dc                  mov eax, dword ptr [ebp-24]
'00719a99    8b08                    mov ecx, dword ptr [eax]
'00719a9b    8d55b0                  lea edx, dword ptr [ebp-50]
'00719a9e    52                      push edx
'00719a9f    50                      push eax
'00719aa0    ff91b4000000            call dword ptr [ecx+000000b4]
'00719aa6    dbe2                    fnclex
'00719aa8    85c0                    test ax, ax
'00719aaa    7d15                    jge 719ac1
'00719aac    8b4ddc                  mov ecx, dword ptr [ebp-24]
'00719aaf    68b4000000              push 000000b4
'00719ab4    6830314300              push 00433130
'00719ab9    51                      push ecx
'00719aba    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00719abb    ff1580104000            call dword ptr [00401080]
'00719ac1    8b45b0                  mov eax, dword ptr [ebp-50]
'00719ac4    8b10                    mov edx, dword ptr [eax]
'00719ac6    8d7dac                  lea edi, dword ptr [ebp-54]
'00719ac9    57                      push edi
'00719aca    83ec10                  sub esp, 10
'00719acd    8bfc                    mov edi, esp
'00719acf    b902000000              mov ecx, 00000002
'00719ad4    890f                    mov dword ptr [edi], ecx
'00719ad6    8b8d40ffffff            mov ecx, dword ptr [ebp+ffffff40]
'00719adc    894f04                  mov dword ptr [edi+04], ecx
'00719adf    66899d44ffffff          mov word ptr [ebp+ffffff44], bx
'00719ae6    8b8d44ffffff            mov ecx, dword ptr [ebp+ffffff44]
'00719aec    894f08                  mov dword ptr [edi+08], ecx
'00719aef    8b8d48ffffff            mov ecx, dword ptr [ebp+ffffff48]
'00719af5    50                      push eax
'00719af6    8bf0                    mov esi, eax
'00719af8    894f0c                  mov dword ptr [edi+0c], ecx
'00719afb    ff5230                  call dword ptr [edx+30]
'00719afe    dbe2                    fnclex
'00719b00    85c0                    test ax, ax
'00719b02    7d0f                    jge 719b13
'00719b04    6a30                    push 30
'00719b06    68d8304300              push 004330d8
'00719b0b    56                      push esi
'00719b0c    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00719b0d    ff1580104000            call dword ptr [00401080]
'00719b13    8b45d8                  mov eax, dword ptr [ebp-28]
'00719b16    8b10                    mov edx, dword ptr [eax]
'00719b18    8b7dac                  mov edi, dword ptr [ebp-54]
'00719b1b    8d4db8                  lea ecx, dword ptr [ebp-48]
'00719b1e    51                      push ecx
'00719b1f    50                      push eax
'00719b20    ff92b4000000            call dword ptr [edx+000000b4]
'00719b26    dbe2                    fnclex
'00719b28    85c0                    test ax, ax
'00719b2a    7d15                    jge 719b41
'00719b2c    8b55d8                  mov edx, dword ptr [ebp-28]
'00719b2f    68b4000000              push 000000b4
'00719b34    6830314300              push 00433130
'00719b39    52                      push edx
'00719b3a    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00719b3b    ff1580104000            call dword ptr [00401080]
'00719b41    8b45b8                  mov eax, dword ptr [ebp-48]
'00719b44    8b10                    mov edx, dword ptr [eax]
'00719b46    66899d54ffffff          mov word ptr [ebp+ffffff54], bx
'00719b4d    8d5db4                  lea ebx, dword ptr [ebp-4c]
'00719b50    53                      push ebx
'00719b51    83ec10                  sub esp, 10
'00719b54    8bdc                    mov ebx, esp
'00719b56    b902000000              mov ecx, 00000002
'00719b5b    890b                    mov dword ptr [ebx], ecx
'00719b5d    8b8d50ffffff            mov ecx, dword ptr [ebp+ffffff50]
'00719b63    894b04                  mov dword ptr [ebx+04], ecx
'00719b66    8b8d54ffffff            mov ecx, dword ptr [ebp+ffffff54]
'00719b6c    894b08                  mov dword ptr [ebx+08], ecx
'00719b6f    8b8d58ffffff            mov ecx, dword ptr [ebp+ffffff58]
'00719b75    50                      push eax
'00719b76    8bf0                    mov esi, eax
'00719b78    894b0c                  mov dword ptr [ebx+0c], ecx
'00719b7b    ff5230                  call dword ptr [edx+30]
'00719b7e    dbe2                    fnclex
'00719b80    85c0                    test ax, ax
'00719b82    7d0f                    jge 719b93
'00719b84    6a30                    push 30
'00719b86    68d8304300              push 004330d8
'00719b8b    56                      push esi
'00719b8c    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00719b8d    ff1580104000            call dword ptr [00401080]
'00719b93    8b45b4                  mov eax, dword ptr [ebp-4c]
'00719b96    83ec10                  sub esp, 10
'00719b99    b909000000              mov ecx, 00000009
'00719b9e    8bf4                    mov esi, esp
'00719ba0    890e                    mov dword ptr [esi], ecx
'00719ba2    894d9c                  mov dword ptr [ebp-64], ecx
'00719ba5    8b4da0                  mov ecx, dword ptr [ebp-60]
'00719ba8    894e04                  mov dword ptr [esi+04], ecx
'00719bab    8945a4                  mov dword ptr [ebp-5c], eax
'00719bae    894608                  mov dword ptr [esi+08], eax
'00719bb1    8b45a8                  mov eax, dword ptr [ebp-58]
'00719bb4    c745b400000000          mov dword ptr [ebp-4c], 00000000
'00719bbb    8b17                    mov edx, dword ptr [edi]
'00719bbd    57                      push edi
'00719bbe    89460c                  mov dword ptr [esi+0c], eax
'00719bc1    ff5248                  call dword ptr [edx+48]
'00719bc4    dbe2                    fnclex
'00719bc6    85c0                    test ax, ax
'00719bc8    7d0f                    jge 719bd9
'00719bca    6a48                    push 48
'00719bcc    6880304300              push 00433080
'00719bd1    57                      push edi
'00719bd2    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00719bd3    ff1580104000            call dword ptr [00401080]
'00719bd9    8d4dac                  lea ecx, dword ptr [ebp-54]
'00719bdc    51                      push ecx
'00719bdd    8d55b0                  lea edx, dword ptr [ebp-50]
'00719be0    52                      push edx
'00719be1    8d45b8                  lea eax, dword ptr [ebp-48]
'00719be4    50                      push eax
'00719be5    6a03                    push 03

' *** Reference to "__vbaFreeObjList"
'00719be7    ff154c104000            call dword ptr [0040104c]
    '#Cleanup( var_61, var_6, var_50)
'00719bed    83c410                  add esp, 10
'00719bf0    8d4d9c                  lea ecx, dword ptr [ebp-64]

' *** Reference to "__vbaFreeVar"
'00719bf3    ff1528104000            call dword ptr [00401028]
    '#Cleanup(var_51)
'00719bf9    8b5de0                  mov ebx, dword ptr [ebp-20]
    
End If
'00719bfc    b801000000              mov eax, 00000001
'00719c01    6603c3                  add ax, bx
var_num1 = 1 + 1
'00719c04    0f80690c0000            jo 71a873
'00719c0a    8bd8                    mov ebx, eax
'00719c0c    e9b7fdffff              jmp 7199c8

'ERROR: Two many next close:
Loop
'00719c11    8b45dc                  mov eax, dword ptr [ebp-24]
'00719c14    8b08                    mov ecx, dword ptr [eax]
'00719c16    6a00                    push 00
'00719c18    6a01                    push 01
'00719c1a    50                      push eax
'00719c1b    ff9164010000            call dword ptr [ecx+00000164]
'00719c21    dbe2                    fnclex
'00719c23    85c0                    test ax, ax
'00719c25    7d15                    jge 719c3c
'00719c27    8b55dc                  mov edx, dword ptr [ebp-24]
'00719c2a    6864010000              push 00000164
'00719c2f    6830314300              push 00433130
'00719c34    52                      push edx
'00719c35    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00719c36    ff1580104000            call dword ptr [00401080]
'00719c3c    8b45d8                  mov eax, dword ptr [ebp-28]
'00719c3f    8b08                    mov ecx, dword ptr [eax]
'00719c41    50                      push eax
'00719c42    ff91ec000000            call dword ptr [ecx+000000ec]
'00719c48    dbe2                    fnclex
'00719c4a    85c0                    test ax, ax
'00719c4c    0f8d93fbffff            jge 7197e5
'00719c52    8b55d8                  mov edx, dword ptr [ebp-28]

' *** Reference to "__vbaHresultCheckObj"
'00719c55    8b3580104000            mov esi, dword ptr [00401080]
'00719c5b    68ec000000              push 000000ec
'00719c60    6830314300              push 00433130
'00719c65    52                      push edx
'00719c66    50                      push eax
'00719c67    ffd6                    call esi

' *** Reference to "__vbaStrMove"
'00719c69    8b3dd0124000            mov edi, dword ptr [004012d0]
'00719c6f    e97dfbffff              jmp 7197f1

'ERROR: Two many next close:
Loop
'00719c74    8b08                    mov ecx, dword ptr [eax]
'00719c76    50                      push eax
'00719c77    ff91c4000000            call dword ptr [ecx+000000c4]
'00719c7d    dbe2                    fnclex
'00719c7f    85c0                    test ax, ax
'00719c81    7d11                    jge 719c94
'00719c83    8b55dc                  mov edx, dword ptr [ebp-24]
'00719c86    68c4000000              push 000000c4
'00719c8b    6830314300              push 00433130
'00719c90    52                      push edx
'00719c91    50                      push eax
'00719c92    ffd6                    call esi
'00719c94    8b45d8                  mov eax, dword ptr [ebp-28]
'00719c97    8b08                    mov ecx, dword ptr [eax]
'00719c99    50                      push eax
'00719c9a    ff91c4000000            call dword ptr [ecx+000000c4]
'00719ca0    dbe2                    fnclex
'00719ca2    85c0                    test ax, ax
'00719ca4    7d11                    jge 719cb7
'00719ca6    8b55d8                  mov edx, dword ptr [ebp-28]
'00719ca9    68c4000000              push 000000c4
'00719cae    6830314300              push 00433130
'00719cb3    52                      push edx
'00719cb4    50                      push eax
'00719cb5    ffd6                    call esi
'00719cb7    668b55e8                mov dx, word ptr [ebp-18]
'00719cbb    83ec10                  sub esp, 10
'00719cbe    8bdc                    mov ebx, esp
'00719cc0    b903000000              mov ecx, 00000003
'00719cc5    890b                    mov dword ptr [ebx], ecx
'00719cc7    8b8d50ffffff            mov ecx, dword ptr [ebp+ffffff50]
'00719ccd    894b04                  mov dword ptr [ebx+04], ecx
'00719cd0    33c0                    xor eax, eax
var_num1 = Empty
'00719cd2    894308                  mov dword ptr [ebx+08], eax
'00719cd5    8b8558ffffff            mov eax, dword ptr [ebp+ffffff58]
'00719cdb    89430c                  mov dword ptr [ebx+0c], eax
'00719cde    8b5d08                  mov ebx, dword ptr [ebp+08]
'00719ce1    83ec10                  sub esp, 10
'00719ce4    66899534ffffff          mov word ptr [ebp+ffffff34], dx
'00719ceb    8b8534ffffff            mov eax, dword ptr [ebp+ffffff34]
'00719cf1    8bcc                    mov ecx, esp
'00719cf3    ba02000000              mov edx, 00000002
'00719cf8    8911                    mov dword ptr [ecx], edx
'00719cfa    8b9530ffffff            mov edx, dword ptr [ebp+ffffff30]
'00719d00    895104                  mov dword ptr [ecx+04], edx
'00719d03    8b9538ffffff            mov edx, dword ptr [ebp+ffffff38]
'00719d09    894108                  mov dword ptr [ecx+08], eax
'00719d0c    89510c                  mov dword ptr [ecx+0c], edx
'00719d0f    8b9510ffffff            mov edx, dword ptr [ebp+ffffff10]
'00719d15    83ec10                  sub esp, 10
'00719d18    8bcc                    mov ecx, esp
'00719d1a    b802000000              mov eax, 00000002
'00719d1f    8901                    mov dword ptr [ecx], eax
'00719d21    8b8518ffffff            mov eax, dword ptr [ebp+ffffff18]
'00719d27    895104                  mov dword ptr [ecx+04], edx
'00719d2a    6a03                    push 03
'00719d2c    33f6                    xor esi, esi
var_num8 = Empty
'00719d2e    897108                  mov dword ptr [ecx+08], esi
'00719d31    689e000000              push 0000009e
'00719d36    89410c                  mov dword ptr [ecx+0c], eax
'00719d39    8b0b                    mov ecx, dword ptr [ebx]
'00719d3b    53                      push ebx
'00719d3c    ff9110030000            call dword ptr [ecx+00000310]
'00719d42    50                      push eax
'00719d43    8d55b8                  lea edx, dword ptr [ebp-48]
'00719d46    52                      push edx

' *** Reference to "__vbaObjSet"
'00719d47    ff15b4104000            call dword ptr [004010b4]
Set var_61 = Nothing
'00719d4d    50                      push eax
'00719d4e    8d459c                  lea eax, dword ptr [ebp-64]
'00719d51    50                      push eax

' *** Reference to "__vbaLateIdCallLd"
'00719d52    ff157c114000            call dword ptr [0040117c]
var_51 = var_61.UNK_-256 - 12_158
'00719d58    83c440                  add esp, 40
'00719d5b    50                      push eax

' *** Reference to "__vbaStrVarMove"
'00719d5c    ff153c104000            call dword ptr [0040103c]
'00719d62    8bd0                    mov edx, eax
'00719d64    8d4dcc                  lea ecx, dword ptr [ebp-34]
'00719d67    ffd7                    call edi
'00719d69    8d4dcc                  lea ecx, dword ptr [ebp-34]
'00719d6c    51                      push ecx
'00719d6d    e87e9eddff              call 4f3bf0
Call sub_4F3BF0()
'00719d72    8bd0                    mov edx, eax
'00719d74    8d4dbc                  lea ecx, dword ptr [ebp-44]
'00719d77    ffd7                    call edi
'00719d79    8b55bc                  mov edx, dword ptr [ebp-44]
'00719d7c    8995f4fcffff            mov dword ptr [ebp+fffffcf4], edx
'00719d82    8b5338                  mov edx, dword ptr [ebx+38]
'00719d85    c785f4feffff02000000    mov dword ptr [ebp+fffffef4], 00000002
'00719d8f    c785ecfeffff03000000    mov dword ptr [ebp+fffffeec], 00000003
'00719d99    c745bc00000000          mov dword ptr [ebp-44], 00000000
'00719da0    8b1a                    mov ebx, dword ptr [edx]
'00719da2    8d55b4                  lea edx, dword ptr [ebp-4c]
'00719da5    52                      push edx
'00719da6    83ec10                  sub esp, 10
'00719da9    b90a000000              mov ecx, 0000000a
'00719dae    8bd4                    mov edx, esp
'00719db0    890a                    mov dword ptr [edx], ecx
'00719db2    8b8dd0feffff            mov ecx, dword ptr [ebp+fffffed0]
'00719db8    b804000280              mov eax, 80020004
'00719dbd    894a04                  mov dword ptr [edx+04], ecx
'00719dc0    894208                  mov dword ptr [edx+08], eax
'00719dc3    8bf0                    mov esi, eax
'00719dc5    8b85d8feffff            mov eax, dword ptr [ebp+fffffed8]
'00719dcb    83ec10                  sub esp, 10
'00719dce    89420c                  mov dword ptr [edx+0c], eax
'00719dd1    8bcc                    mov ecx, esp
'00719dd3    b80a000000              mov eax, 0000000a
'00719dd8    8b95e0feffff            mov edx, dword ptr [ebp+fffffee0]
'00719dde    8901                    mov dword ptr [ecx], eax
'00719de0    8b85e8feffff            mov eax, dword ptr [ebp+fffffee8]
'00719de6    895104                  mov dword ptr [ecx+04], edx
'00719de9    8b95ecfeffff            mov edx, dword ptr [ebp+fffffeec]
'00719def    897108                  mov dword ptr [ecx+08], esi
'00719df2    89410c                  mov dword ptr [ecx+0c], eax
'00719df5    8b85f0feffff            mov eax, dword ptr [ebp+fffffef0]
'00719dfb    83ec10                  sub esp, 10
'00719dfe    8bcc                    mov ecx, esp
'00719e00    8911                    mov dword ptr [ecx], edx
'00719e02    8b95f4feffff            mov edx, dword ptr [ebp+fffffef4]
'00719e08    894104                  mov dword ptr [ecx+04], eax
'00719e0b    8b85f8feffff            mov eax, dword ptr [ebp+fffffef8]
'00719e11    895108                  mov dword ptr [ecx+08], edx
'00719e14    8b95f4fcffff            mov edx, dword ptr [ebp+fffffcf4]
'00719e1a    89410c                  mov dword ptr [ecx+0c], eax
'00719e1d    68bc6c4500              push 00456cbc
'00719e22    8d4dc8                  lea ecx, dword ptr [ebp-38]
'00719e25    ffd7                    call edi

' *** Reference to "__vbaStrCat"
'00719e27    8b3570104000            mov esi, dword ptr [00401070]
'00719e2d    50                      push eax
'00719e2e    ffd6                    call esi
var_1666 = ("select * from SortListe where NomPerso='") & (var_51)
'00719e30    8bd0                    mov edx, eax
'00719e32    8d4dc4                  lea ecx, dword ptr [ebp-3c]
'00719e35    ffd7                    call edi
'00719e37    50                      push eax
'00719e38    6854a44300              push 0043a454
'00719e3d    ffd6                    call esi
var_852 = (var_1666) & ("'")
'00719e3f    8bd0                    mov edx, eax
'00719e41    8d4dc0                  lea ecx, dword ptr [ebp-40]
'00719e44    ffd7                    call edi
'00719e46    8b4d08                  mov ecx, dword ptr [ebp+08]
'00719e49    8b5138                  mov edx, dword ptr [ecx+38]
'00719e4c    50                      push eax
'00719e4d    52                      push edx
'00719e4e    ff93bc000000            call dword ptr [ebx+000000bc]
'00719e54    dbe2                    fnclex
'00719e56    85c0                    test ax, ax
'00719e58    7d18                    jge 719e72
'00719e5a    8b4d08                  mov ecx, dword ptr [ebp+08]
'00719e5d    8b5138                  mov edx, dword ptr [ecx+38]
'00719e60    68bc000000              push 000000bc
'00719e65    68301f4300              push 00431f30
'00719e6a    52                      push edx
'00719e6b    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00719e6c    ff1580104000            call dword ptr [00401080]
'00719e72    8b45b4                  mov eax, dword ptr [ebp-4c]
'00719e75    50                      push eax
'00719e76    8d45d8                  lea eax, dword ptr [ebp-28]
'00719e79    50                      push eax
'00719e7a    c745b400000000          mov dword ptr [ebp-4c], 00000000

' *** Reference to "__vbaObjSet"
'00719e81    ff15b4104000            call dword ptr [004010b4]
Set var_45 = Nothing
'00719e87    8d4dbc                  lea ecx, dword ptr [ebp-44]
'00719e8a    51                      push ecx
'00719e8b    8d55c0                  lea edx, dword ptr [ebp-40]
'00719e8e    52                      push edx
'00719e8f    8d45c4                  lea eax, dword ptr [ebp-3c]
'00719e92    50                      push eax
'00719e93    8d4dc8                  lea ecx, dword ptr [ebp-38]
'00719e96    51                      push ecx
'00719e97    8d55cc                  lea edx, dword ptr [ebp-34]
'00719e9a    52                      push edx
'00719e9b    6a05                    push 05

' *** Reference to "__vbaFreeStrList"
'00719e9d    ff155c124000            call dword ptr [0040125c]
'#Cleanup( -4924, -4924, -4928, -4932, var_58)
'00719ea3    83c418                  add esp, 18
'00719ea6    8d4db8                  lea ecx, dword ptr [ebp-48]

' *** Reference to "__vbaFreeObj"
'00719ea9    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_61)
'00719eaf    8d4d9c                  lea ecx, dword ptr [ebp-64]

' *** Reference to "__vbaFreeVar"
'00719eb2    ff1528104000            call dword ptr [00401028]
'#Cleanup(var_51)
'00719eb8    8d5db8                  lea ebx, dword ptr [ebp-48]
'00719ebb    53                      push ebx
'00719ebc    83ec10                  sub esp, 10
'00719ebf    8bdc                    mov ebx, esp
'00719ec1    8b3d48b07200            mov edi, dword ptr [0072b048]
'00719ec7    b90a000000              mov ecx, 0000000a
'00719ecc    890b                    mov dword ptr [ebx], ecx
'00719ece    8bf1                    mov esi, ecx
'00719ed0    8b8d30ffffff            mov ecx, dword ptr [ebp+ffffff30]
'00719ed6    894b04                  mov dword ptr [ebx+04], ecx
'00719ed9    8b3f                    mov edi, dword ptr [edi]
'00719edb    b804000280              mov eax, 80020004
'00719ee0    894308                  mov dword ptr [ebx+08], eax
'00719ee3    8bd0                    mov edx, eax
'00719ee5    898534ffffff            mov dword ptr [ebp+ffffff34], eax
'00719eeb    8b8538ffffff            mov eax, dword ptr [ebp+ffffff38]
'00719ef1    89430c                  mov dword ptr [ebx+0c], eax
'00719ef4    8b8540ffffff            mov eax, dword ptr [ebp+ffffff40]
'00719efa    83ec10                  sub esp, 10
'00719efd    8bcc                    mov ecx, esp
'00719eff    8931                    mov dword ptr [ecx], esi
'00719f01    894104                  mov dword ptr [ecx+04], eax
'00719f04    895108                  mov dword ptr [ecx+08], edx
'00719f07    899544ffffff            mov dword ptr [ebp+ffffff44], edx
'00719f0d    8b9548ffffff            mov edx, dword ptr [ebp+ffffff48]
'00719f13    89510c                  mov dword ptr [ecx+0c], edx
'00719f16    8b9550ffffff            mov edx, dword ptr [ebp+ffffff50]
'00719f1c    83ec10                  sub esp, 10
'00719f1f    8bcc                    mov ecx, esp
'00719f21    b803000000              mov eax, 00000003
'00719f26    8901                    mov dword ptr [ecx], eax
'00719f28    895104                  mov dword ptr [ecx+04], edx
'00719f2b    b801000000              mov eax, 00000001
'00719f30    894108                  mov dword ptr [ecx+08], eax
'00719f33    8b8558ffffff            mov eax, dword ptr [ebp+ffffff58]
'00719f39    89410c                  mov dword ptr [ecx+0c], eax
'00719f3c    8b0d48b07200            mov ecx, dword ptr [0072b048]
'00719f42    6818b64300              push 0043b618
'00719f47    51                      push ecx
'00719f48    ff97bc000000            call dword ptr [edi+000000bc]
'00719f4e    dbe2                    fnclex
'00719f50    85c0                    test ax, ax
'00719f52    7d18                    jge 719f6c
'00719f54    8b1548b07200            mov edx, dword ptr [0072b048]
'00719f5a    68bc000000              push 000000bc
'00719f5f    68301f4300              push 00431f30
'00719f64    52                      push edx
'00719f65    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00719f66    ff1580104000            call dword ptr [00401080]
'00719f6c    8b45b8                  mov eax, dword ptr [ebp-48]
'00719f6f    50                      push eax
'00719f70    8d45dc                  lea eax, dword ptr [ebp-24]
'00719f73    50                      push eax
'00719f74    c745b800000000          mov dword ptr [ebp-48], 00000000

' *** Reference to "__vbaObjSet"
'00719f7b    ff15b4104000            call dword ptr [004010b4]
Set var_39 = Nothing
'00719f81    8b45d8                  mov eax, dword ptr [ebp-28]
'00719f84    8b08                    mov ecx, dword ptr [eax]
'00719f86    8b7de8                  mov edi, dword ptr [ebp-18]
'00719f89    8d9538feffff            lea edx, dword ptr [ebp+fffffe38]
'00719f8f    52                      push edx
'00719f90    50                      push eax
'00719f91    ff5134                  call dword ptr [ecx+34]
'00719f94    dbe2                    fnclex
'00719f96    85c0                    test ax, ax
'00719f98    7d16                    jge 719fb0
'00719f9a    8b4dd8                  mov ecx, dword ptr [ebp-28]

' *** Reference to "__vbaHresultCheckObj"
'00719f9d    8b3580104000            mov esi, dword ptr [00401080]
'00719fa3    6a34                    push 34
'00719fa5    6830314300              push 00433130
'00719faa    51                      push ecx
'00719fab    50                      push eax
'00719fac    ffd6                    call esi
'00719fae    eb06                    jmp 719fb6

' *** Reference to "__vbaHresultCheckObj"
'00719fb0    8b3580104000            mov esi, dword ptr [00401080]
'00719fb6    6683bd38feffff00        cmp word ptr [ebp+fffffe38], 00
'00719fbe    8b45dc                  mov eax, dword ptr [ebp-24]
'00719fc1    0f8556040000            jne 71a41d

Do While (0 = 0)
'00719fc7    8b10                    mov edx, dword ptr [eax]
'00719fc9    50                      push eax
'00719fca    ff92c0000000            call dword ptr [edx+000000c0]
'00719fd0    dbe2                    fnclex
'00719fd2    85c0                    test ax, ax
'00719fd4    7d11                    jge 719fe7
'00719fd6    8b4ddc                  mov ecx, dword ptr [ebp-24]
'00719fd9    68c0000000              push 000000c0
'00719fde    6830314300              push 00433130
'00719fe3    51                      push ecx
'00719fe4    50                      push eax
'00719fe5    ffd6                    call esi
'00719fe7    8b45dc                  mov eax, dword ptr [ebp-24]
'00719fea    8b10                    mov edx, dword ptr [eax]
'00719fec    8d4db4                  lea ecx, dword ptr [ebp-4c]
'00719fef    51                      push ecx
'00719ff0    50                      push eax
'00719ff1    ff92b4000000            call dword ptr [edx+000000b4]
'00719ff7    dbe2                    fnclex
'00719ff9    85c0                    test ax, ax
'00719ffb    7d11                    jge 71a00e
'00719ffd    8b55dc                  mov edx, dword ptr [ebp-24]
'0071a000    68b4000000              push 000000b4
'0071a005    6830314300              push 00433130
'0071a00a    52                      push edx
'0071a00b    50                      push eax
'0071a00c    ffd6                    call esi
'0071a00e    8b45b4                  mov eax, dword ptr [ebp-4c]
'0071a011    8d5db0                  lea ebx, dword ptr [ebp-50]
'0071a014    53                      push ebx
'0071a015    83ec10                  sub esp, 10
'0071a018    ba02000000              mov edx, 00000002
'0071a01d    8bdc                    mov ebx, esp
'0071a01f    8913                    mov dword ptr [ebx], edx
'0071a021    8995ecfeffff            mov dword ptr [ebp+fffffeec], edx
'0071a027    8b95f0feffff            mov edx, dword ptr [ebp+fffffef0]
'0071a02d    33c9                    xor ecx, ecx
'0071a02f    895304                  mov dword ptr [ebx+04], edx
'0071a032    898df4feffff            mov dword ptr [ebp+fffffef4], ecx
'0071a038    8b30                    mov esi, dword ptr [eax]
'0071a03a    894b08                  mov dword ptr [ebx+08], ecx
'0071a03d    8b8df8feffff            mov ecx, dword ptr [ebp+fffffef8]
'0071a043    50                      push eax
'0071a044    898530feffff            mov dword ptr [ebp+fffffe30], eax
'0071a04a    894b0c                  mov dword ptr [ebx+0c], ecx
'0071a04d    ff5630                  call dword ptr [esi+30]
'0071a050    dbe2                    fnclex
'0071a052    85c0                    test ax, ax
'0071a054    7d15                    jge 71a06b
'0071a056    8b9530feffff            mov edx, dword ptr [ebp+fffffe30]
'0071a05c    6a30                    push 30
'0071a05e    68d8304300              push 004330d8
'0071a063    52                      push edx
'0071a064    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0071a065    ff1580104000            call dword ptr [00401080]
'0071a06b    83ec10                  sub esp, 10
'0071a06e    8bdc                    mov ebx, esp
'0071a070    8b75b0                  mov esi, dword ptr [ebp-50]
'0071a073    33c0                    xor eax, eax
    var_num1 = Empty
'0071a075    b903000000              mov ecx, 00000003
'0071a07a    890b                    mov dword ptr [ebx], ecx
'0071a07c    8b8d50ffffff            mov ecx, dword ptr [ebp+ffffff50]
'0071a082    894b04                  mov dword ptr [ebx+04], ecx
'0071a085    894308                  mov dword ptr [ebx+08], eax
'0071a088    898554ffffff            mov dword ptr [ebp+ffffff54], eax
'0071a08e    8b8558ffffff            mov eax, dword ptr [ebp+ffffff58]
'0071a094    89430c                  mov dword ptr [ebx+0c], eax
'0071a097    83ec10                  sub esp, 10
'0071a09a    8bcc                    mov ecx, esp
'0071a09c    ba02000000              mov edx, 00000002
'0071a0a1    8911                    mov dword ptr [ecx], edx
'0071a0a3    8b9530ffffff            mov edx, dword ptr [ebp+ffffff30]
'0071a0a9    895104                  mov dword ptr [ecx+04], edx
'0071a0ac    8b9538ffffff            mov edx, dword ptr [ebp+ffffff38]
'0071a0b2    6689bd34ffffff          mov word ptr [ebp+ffffff34], di
'0071a0b9    8b8534ffffff            mov eax, dword ptr [ebp+ffffff34]
'0071a0bf    8b3e                    mov edi, dword ptr [esi]
'0071a0c1    894108                  mov dword ptr [ecx+08], eax
'0071a0c4    89510c                  mov dword ptr [ecx+0c], edx
'0071a0c7    8b9510ffffff            mov edx, dword ptr [ebp+ffffff10]
'0071a0cd    83ec10                  sub esp, 10
'0071a0d0    8bcc                    mov ecx, esp
'0071a0d2    b802000000              mov eax, 00000002
'0071a0d7    8901                    mov dword ptr [ecx], eax
'0071a0d9    895104                  mov dword ptr [ecx+04], edx
'0071a0dc    b801000000              mov eax, 00000001
'0071a0e1    894108                  mov dword ptr [ecx+08], eax
'0071a0e4    8b8518ffffff            mov eax, dword ptr [ebp+ffffff18]
'0071a0ea    6a03                    push 03
'0071a0ec    89410c                  mov dword ptr [ecx+0c], eax
'0071a0ef    8b4508                  mov eax, dword ptr [ebp+08]
'0071a0f2    8b08                    mov ecx, dword ptr [eax]
'0071a0f4    689e000000              push 0000009e
'0071a0f9    50                      push eax
'0071a0fa    ff9110030000            call dword ptr [ecx+00000310]
'0071a100    50                      push eax
'0071a101    8d55b8                  lea edx, dword ptr [ebp-48]
'0071a104    52                      push edx

' *** Reference to "__vbaObjSet"
'0071a105    ff15b4104000            call dword ptr [004010b4]
    Set var_61 = Me
'0071a10b    50                      push eax
'0071a10c    8d459c                  lea eax, dword ptr [ebp-64]
'0071a10f    50                      push eax

' *** Reference to "__vbaLateIdCallLd"
'0071a110    ff157c114000            call dword ptr [0040117c]
    var_51 = var_61.UNK_frmImport_158
'0071a116    8b10                    mov edx, dword ptr [eax]
'0071a118    83c430                  add esp, 30
'0071a11b    8bcc                    mov ecx, esp
'0071a11d    8911                    mov dword ptr [ecx], edx
'0071a11f    8b5004                  mov edx, dword ptr [eax+04]
'0071a122    895104                  mov dword ptr [ecx+04], edx
'0071a125    8b5008                  mov edx, dword ptr [eax+08]
'0071a128    8b400c                  mov eax, dword ptr [eax+0c]
'0071a12b    895108                  mov dword ptr [ecx+08], edx
'0071a12e    56                      push esi
'0071a12f    89410c                  mov dword ptr [ecx+0c], eax
'0071a132    ff5748                  call dword ptr [edi+48]
'0071a135    dbe2                    fnclex
'0071a137    85c0                    test ax, ax
'0071a139    7d0f                    jge 71a14a
'0071a13b    6a48                    push 48
'0071a13d    6880304300              push 00433080
'0071a142    56                      push esi
'0071a143    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0071a144    ff1580104000            call dword ptr [00401080]
'0071a14a    8d4db0                  lea ecx, dword ptr [ebp-50]
'0071a14d    51                      push ecx
'0071a14e    8d55b4                  lea edx, dword ptr [ebp-4c]
'0071a151    52                      push edx
'0071a152    8d45b8                  lea eax, dword ptr [ebp-48]
'0071a155    50                      push eax
'0071a156    6a03                    push 03

' *** Reference to "__vbaFreeObjList"
'0071a158    ff154c104000            call dword ptr [0040104c]
    '#Cleanup( var_61, var_62, var_6)

' *** Reference to "__vbaFreeVar"
'0071a15e    8b3d28104000            mov edi, dword ptr [00401028]
'0071a164    83c410                  add esp, 10
'0071a167    8d4d9c                  lea ecx, dword ptr [ebp-64]
'0071a16a    ffd7                    call edi
    '#Cleanup(var_51)
'0071a16c    bb01000000              mov ebx, 00000001
'0071a171    b805000000              mov eax, 00000005
'0071a176    663bd8                  cmp bx, ax
'0071a179    895de0                  mov dword ptr [ebp-20], ebx
'0071a17c    0f8f40020000            jg 71a3c2
    
    Do While (    1 <= 5)
'0071a182    8b45d8                  mov eax, dword ptr [ebp-28]
'0071a185    8b08                    mov ecx, dword ptr [eax]
'0071a187    8d55b8                  lea edx, dword ptr [ebp-48]
'0071a18a    52                      push edx
'0071a18b    50                      push eax
'0071a18c    ff91b4000000            call dword ptr [ecx+000000b4]
'0071a192    dbe2                    fnclex
'0071a194    85c0                    test ax, ax
'0071a196    7d15                    jge 71a1ad
'0071a198    8b4dd8                  mov ecx, dword ptr [ebp-28]
'0071a19b    68b4000000              push 000000b4
'0071a1a0    6830314300              push 00433130
'0071a1a5    51                      push ecx
'0071a1a6    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0071a1a7    ff1580104000            call dword ptr [00401080]
'0071a1ad    8b45b8                  mov eax, dword ptr [ebp-48]
'0071a1b0    8b10                    mov edx, dword ptr [eax]
'0071a1b2    8d75b4                  lea esi, dword ptr [ebp-4c]
'0071a1b5    56                      push esi
'0071a1b6    83ec10                  sub esp, 10
'0071a1b9    8bf4                    mov esi, esp
'0071a1bb    b902000000              mov ecx, 00000002
'0071a1c0    890e                    mov dword ptr [esi], ecx
'0071a1c2    8b8d50ffffff            mov ecx, dword ptr [ebp+ffffff50]
'0071a1c8    894e04                  mov dword ptr [esi+04], ecx
'0071a1cb    66899d54ffffff          mov word ptr [ebp+ffffff54], bx
'0071a1d2    8b8d54ffffff            mov ecx, dword ptr [ebp+ffffff54]
'0071a1d8    894e08                  mov dword ptr [esi+08], ecx
'0071a1db    8b8d58ffffff            mov ecx, dword ptr [ebp+ffffff58]
'0071a1e1    50                      push eax
'0071a1e2    898530feffff            mov dword ptr [ebp+fffffe30], eax
'0071a1e8    894e0c                  mov dword ptr [esi+0c], ecx
'0071a1eb    ff5230                  call dword ptr [edx+30]
'0071a1ee    dbe2                    fnclex
'0071a1f0    85c0                    test ax, ax
'0071a1f2    7d15                    jge 71a209
'0071a1f4    8b9530feffff            mov edx, dword ptr [ebp+fffffe30]
'0071a1fa    6a30                    push 30
'0071a1fc    68d8304300              push 004330d8
'0071a201    52                      push edx
'0071a202    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0071a203    ff1580104000            call dword ptr [00401080]
'0071a209    8b45b4                  mov eax, dword ptr [ebp-4c]
'0071a20c    8945a4                  mov dword ptr [ebp-5c], eax
'0071a20f    8d459c                  lea eax, dword ptr [ebp-64]
'0071a212    50                      push eax
'0071a213    c745b400000000          mov dword ptr [ebp-4c], 00000000
'0071a21a    c7459c09000000          mov dword ptr [ebp-64], 00000009

' *** Reference to "rtcIsNull"
'0071a221    ff1540114000            call dword ptr [00401140]
'0071a227    33f6                    xor esi, esi
    var_num8 = Empty
'0071a229    668bf0                  mov si, ax
'0071a22c    8d4db8                  lea ecx, dword ptr [ebp-48]
'0071a22f    f7d6                    not esi

' *** Reference to "__vbaFreeObj"
'0071a231    ff1508134000            call dword ptr [00401308]
    '#Cleanup(var_61)
'0071a237    8d4d9c                  lea ecx, dword ptr [ebp-64]
'0071a23a    ffd7                    call edi
    '#Cleanup(var_51)
'0071a23c    6685f6                  test esi, esi
'0071a23f    0f8468010000            je 71a3ad
    
    If (    Not (IsNull(var_62))) Then
'0071a245    8b45dc                  mov eax, dword ptr [ebp-24]
'0071a248    8b08                    mov ecx, dword ptr [eax]
'0071a24a    8d55b0                  lea edx, dword ptr [ebp-50]
'0071a24d    52                      push edx
'0071a24e    50                      push eax
'0071a24f    ff91b4000000            call dword ptr [ecx+000000b4]
'0071a255    dbe2                    fnclex
'0071a257    85c0                    test ax, ax
'0071a259    7d15                    jge 71a270
'0071a25b    8b4ddc                  mov ecx, dword ptr [ebp-24]
'0071a25e    68b4000000              push 000000b4
'0071a263    6830314300              push 00433130
'0071a268    51                      push ecx
'0071a269    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0071a26a    ff1580104000            call dword ptr [00401080]
'0071a270    8b45b0                  mov eax, dword ptr [ebp-50]
'0071a273    8b10                    mov edx, dword ptr [eax]
'0071a275    8d75ac                  lea esi, dword ptr [ebp-54]
'0071a278    56                      push esi
'0071a279    83ec10                  sub esp, 10
'0071a27c    8bf4                    mov esi, esp
'0071a27e    b902000000              mov ecx, 00000002
'0071a283    890e                    mov dword ptr [esi], ecx
'0071a285    8b8d40ffffff            mov ecx, dword ptr [ebp+ffffff40]
'0071a28b    894e04                  mov dword ptr [esi+04], ecx
'0071a28e    66899d44ffffff          mov word ptr [ebp+ffffff44], bx
'0071a295    8b8d44ffffff            mov ecx, dword ptr [ebp+ffffff44]
'0071a29b    894e08                  mov dword ptr [esi+08], ecx
'0071a29e    8b8d48ffffff            mov ecx, dword ptr [ebp+ffffff48]
'0071a2a4    50                      push eax
'0071a2a5    8bf8                    mov edi, eax
'0071a2a7    894e0c                  mov dword ptr [esi+0c], ecx
'0071a2aa    ff5230                  call dword ptr [edx+30]
'0071a2ad    dbe2                    fnclex
'0071a2af    85c0                    test ax, ax
'0071a2b1    7d0f                    jge 71a2c2
'0071a2b3    6a30                    push 30
'0071a2b5    68d8304300              push 004330d8
'0071a2ba    57                      push edi
'0071a2bb    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0071a2bc    ff1580104000            call dword ptr [00401080]
'0071a2c2    8b45d8                  mov eax, dword ptr [ebp-28]
'0071a2c5    8b10                    mov edx, dword ptr [eax]
'0071a2c7    8b7dac                  mov edi, dword ptr [ebp-54]
'0071a2ca    8d4db8                  lea ecx, dword ptr [ebp-48]
'0071a2cd    51                      push ecx
'0071a2ce    50                      push eax
'0071a2cf    ff92b4000000            call dword ptr [edx+000000b4]
'0071a2d5    dbe2                    fnclex
'0071a2d7    85c0                    test ax, ax
'0071a2d9    7d15                    jge 71a2f0
'0071a2db    8b55d8                  mov edx, dword ptr [ebp-28]
'0071a2de    68b4000000              push 000000b4
'0071a2e3    6830314300              push 00433130
'0071a2e8    52                      push edx
'0071a2e9    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0071a2ea    ff1580104000            call dword ptr [00401080]
'0071a2f0    8b45b8                  mov eax, dword ptr [ebp-48]
'0071a2f3    8b10                    mov edx, dword ptr [eax]
'0071a2f5    66899d54ffffff          mov word ptr [ebp+ffffff54], bx
'0071a2fc    8d5db4                  lea ebx, dword ptr [ebp-4c]
'0071a2ff    53                      push ebx
'0071a300    83ec10                  sub esp, 10
'0071a303    8bdc                    mov ebx, esp
'0071a305    b902000000              mov ecx, 00000002
'0071a30a    890b                    mov dword ptr [ebx], ecx
'0071a30c    8b8d50ffffff            mov ecx, dword ptr [ebp+ffffff50]
'0071a312    894b04                  mov dword ptr [ebx+04], ecx
'0071a315    8b8d54ffffff            mov ecx, dword ptr [ebp+ffffff54]
'0071a31b    894b08                  mov dword ptr [ebx+08], ecx
'0071a31e    8b8d58ffffff            mov ecx, dword ptr [ebp+ffffff58]
'0071a324    50                      push eax
'0071a325    8bf0                    mov esi, eax
'0071a327    894b0c                  mov dword ptr [ebx+0c], ecx
'0071a32a    ff5230                  call dword ptr [edx+30]
'0071a32d    dbe2                    fnclex
'0071a32f    85c0                    test ax, ax
'0071a331    7d0f                    jge 71a342
'0071a333    6a30                    push 30
'0071a335    68d8304300              push 004330d8
'0071a33a    56                      push esi
'0071a33b    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0071a33c    ff1580104000            call dword ptr [00401080]
'0071a342    8b45b4                  mov eax, dword ptr [ebp-4c]
'0071a345    83ec10                  sub esp, 10
'0071a348    b909000000              mov ecx, 00000009
'0071a34d    8bf4                    mov esi, esp
'0071a34f    890e                    mov dword ptr [esi], ecx
'0071a351    894d9c                  mov dword ptr [ebp-64], ecx
'0071a354    8b4da0                  mov ecx, dword ptr [ebp-60]
'0071a357    894e04                  mov dword ptr [esi+04], ecx
'0071a35a    8945a4                  mov dword ptr [ebp-5c], eax
'0071a35d    894608                  mov dword ptr [esi+08], eax
'0071a360    8b45a8                  mov eax, dword ptr [ebp-58]
'0071a363    c745b400000000          mov dword ptr [ebp-4c], 00000000
'0071a36a    8b17                    mov edx, dword ptr [edi]
'0071a36c    57                      push edi
'0071a36d    89460c                  mov dword ptr [esi+0c], eax
'0071a370    ff5248                  call dword ptr [edx+48]
'0071a373    dbe2                    fnclex
'0071a375    85c0                    test ax, ax
'0071a377    7d0f                    jge 71a388
'0071a379    6a48                    push 48
'0071a37b    6880304300              push 00433080
'0071a380    57                      push edi
'0071a381    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0071a382    ff1580104000            call dword ptr [00401080]
'0071a388    8d4dac                  lea ecx, dword ptr [ebp-54]
'0071a38b    51                      push ecx
'0071a38c    8d55b0                  lea edx, dword ptr [ebp-50]
'0071a38f    52                      push edx
'0071a390    8d45b8                  lea eax, dword ptr [ebp-48]
'0071a393    50                      push eax
'0071a394    6a03                    push 03

' *** Reference to "__vbaFreeObjList"
'0071a396    ff154c104000            call dword ptr [0040104c]
    '#Cleanup( var_61, var_6, var_50)

' *** Reference to "__vbaFreeVar"
'0071a39c    8b3d28104000            mov edi, dword ptr [00401028]
'0071a3a2    83c410                  add esp, 10
'0071a3a5    8d4d9c                  lea ecx, dword ptr [ebp-64]
'0071a3a8    ffd7                    call edi
    '#Cleanup(var_51)
'0071a3aa    8b5de0                  mov ebx, dword ptr [ebp-20]
    
End If
'0071a3ad    b801000000              mov eax, 00000001
'0071a3b2    6603c3                  add ax, bx
var_num1 = 1 + 1
'0071a3b5    0f80b8040000            jo 71a873
'0071a3bb    8bd8                    mov ebx, eax
'0071a3bd    e9affdffff              jmp 71a171

'ERROR: Two many next close:
Loop
'0071a3c2    8b45dc                  mov eax, dword ptr [ebp-24]
'0071a3c5    8b08                    mov ecx, dword ptr [eax]
'0071a3c7    6a00                    push 00
'0071a3c9    6a01                    push 01
'0071a3cb    50                      push eax
'0071a3cc    ff9164010000            call dword ptr [ecx+00000164]
'0071a3d2    dbe2                    fnclex
'0071a3d4    85c0                    test ax, ax
'0071a3d6    7d15                    jge 71a3ed
'0071a3d8    8b55dc                  mov edx, dword ptr [ebp-24]
'0071a3db    6864010000              push 00000164
'0071a3e0    6830314300              push 00433130
'0071a3e5    52                      push edx
'0071a3e6    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0071a3e7    ff1580104000            call dword ptr [00401080]
'0071a3ed    8b45d8                  mov eax, dword ptr [ebp-28]
'0071a3f0    8b08                    mov ecx, dword ptr [eax]
'0071a3f2    50                      push eax
'0071a3f3    ff91ec000000            call dword ptr [ecx+000000ec]
'0071a3f9    dbe2                    fnclex
'0071a3fb    85c0                    test ax, ax
'0071a3fd    0f8d7efbffff            jge 719f81
'0071a403    8b55d8                  mov edx, dword ptr [ebp-28]
'0071a406    68ec000000              push 000000ec
'0071a40b    6830314300              push 00433130
'0071a410    52                      push edx
'0071a411    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0071a412    ff1580104000            call dword ptr [00401080]
'0071a418    e964fbffff              jmp 719f81

'ERROR: Two many next close:
Loop
'0071a41d    8b08                    mov ecx, dword ptr [eax]
'0071a41f    50                      push eax
'0071a420    ff91c4000000            call dword ptr [ecx+000000c4]
'0071a426    dbe2                    fnclex
'0071a428    85c0                    test ax, ax
'0071a42a    7d11                    jge 71a43d
'0071a42c    8b55dc                  mov edx, dword ptr [ebp-24]
'0071a42f    68c4000000              push 000000c4
'0071a434    6830314300              push 00433130
'0071a439    52                      push edx
'0071a43a    50                      push eax
'0071a43b    ffd6                    call esi
'0071a43d    8b45d8                  mov eax, dword ptr [ebp-28]
'0071a440    8b08                    mov ecx, dword ptr [eax]
'0071a442    50                      push eax
'0071a443    ff91c4000000            call dword ptr [ecx+000000c4]
'0071a449    dbe2                    fnclex
'0071a44b    85c0                    test ax, ax
'0071a44d    0f8d5f010000            jge 71a5b2
'0071a453    8b55d8                  mov edx, dword ptr [ebp-28]
'0071a456    68c4000000              push 000000c4
'0071a45b    6830314300              push 00433130
'0071a460    52                      push edx
'0071a461    50                      push eax
'0071a462    ffd6                    call esi
'0071a464    e949010000              jmp 71a5b2

'ERROR: Two many next close:
End If
'0071a469    8b7de8                  mov edi, dword ptr [ebp-18]
'0071a46c    83ec10                  sub esp, 10
'0071a46f    8bdc                    mov ebx, esp
'0071a471    33c0                    xor eax, eax
var_num1 = Empty
'0071a473    b903000000              mov ecx, 00000003
'0071a478    890b                    mov dword ptr [ebx], ecx
'0071a47a    8b8d50ffffff            mov ecx, dword ptr [ebp+ffffff50]
'0071a480    894b04                  mov dword ptr [ebx+04], ecx
'0071a483    894308                  mov dword ptr [ebx+08], eax
'0071a486    8b8558ffffff            mov eax, dword ptr [ebp+ffffff58]
'0071a48c    89430c                  mov dword ptr [ebx+0c], eax
'0071a48f    83ec10                  sub esp, 10
'0071a492    8bcc                    mov ecx, esp
'0071a494    ba02000000              mov edx, 00000002
'0071a499    8911                    mov dword ptr [ecx], edx
'0071a49b    8b9530ffffff            mov edx, dword ptr [ebp+ffffff30]
'0071a4a1    895104                  mov dword ptr [ecx+04], edx
'0071a4a4    8b9538ffffff            mov edx, dword ptr [ebp+ffffff38]
'0071a4aa    6689bd34ffffff          mov word ptr [ebp+ffffff34], di
'0071a4b1    8b8534ffffff            mov eax, dword ptr [ebp+ffffff34]
'0071a4b7    894108                  mov dword ptr [ecx+08], eax
'0071a4ba    89510c                  mov dword ptr [ecx+0c], edx
'0071a4bd    8b9510ffffff            mov edx, dword ptr [ebp+ffffff10]
'0071a4c3    83ec10                  sub esp, 10
'0071a4c6    8bcc                    mov ecx, esp
'0071a4c8    b802000000              mov eax, 00000002
'0071a4cd    8901                    mov dword ptr [ecx], eax
'0071a4cf    8b8518ffffff            mov eax, dword ptr [ebp+ffffff18]
'0071a4d5    895104                  mov dword ptr [ecx+04], edx
'0071a4d8    be01000000              mov esi, 00000001
'0071a4dd    897108                  mov dword ptr [ecx+08], esi
'0071a4e0    6a03                    push 03
'0071a4e2    89410c                  mov dword ptr [ecx+0c], eax
'0071a4e5    8b4508                  mov eax, dword ptr [ebp+08]
'0071a4e8    8b08                    mov ecx, dword ptr [eax]
'0071a4ea    689e000000              push 0000009e
'0071a4ef    50                      push eax
'0071a4f0    ff9110030000            call dword ptr [ecx+00000310]
'0071a4f6    50                      push eax
'0071a4f7    8d55b8                  lea edx, dword ptr [ebp-48]
'0071a4fa    52                      push edx

' *** Reference to "__vbaObjSet"
'0071a4fb    ff15b4104000            call dword ptr [004010b4]
Set var_61 = Me
'0071a501    50                      push eax
'0071a502    8d459c                  lea eax, dword ptr [ebp-64]
'0071a505    50                      push eax

' *** Reference to "__vbaLateIdCallLd"
'0071a506    ff157c114000            call dword ptr [0040117c]
var_51 = var_61.UNK_frmImport_158
'0071a50c    83c440                  add esp, 40
'0071a50f    b90a000000              mov ecx, 0000000a
'0071a514    898d5cffffff            mov dword ptr [ebp+ffffff5c], ecx
'0071a51a    898d6cffffff            mov dword ptr [ebp+ffffff6c], ecx
'0071a520    898d7cffffff            mov dword ptr [ebp+ffffff7c], ecx
'0071a526    b804000280              mov eax, 80020004
'0071a52b    8d8d5cffffff            lea ecx, dword ptr [ebp+ffffff5c]
'0071a531    51                      push ecx
'0071a532    898564ffffff            mov dword ptr [ebp+ffffff64], eax
'0071a538    898574ffffff            mov dword ptr [ebp+ffffff74], eax
'0071a53e    894584                  mov dword ptr [ebp-7c], eax
'0071a541    8d956cffffff            lea edx, dword ptr [ebp+ffffff6c]
'0071a547    52                      push edx
'0071a548    8d857cffffff            lea eax, dword ptr [ebp+ffffff7c]
'0071a54e    50                      push eax
'0071a54f    6a00                    push 00
'0071a551    8d4d9c                  lea ecx, dword ptr [ebp-64]
'0071a554    51                      push ecx
'0071a555    8d95ecfeffff            lea edx, dword ptr [ebp+fffffeec]
'0071a55b    52                      push edx
'0071a55c    8d458c                  lea eax, dword ptr [ebp-74]
'0071a55f    50                      push eax
'0071a560    c785f4feffff146d4500    mov dword ptr [ebp+fffffef4], 00456d14
'0071a56a    c785ecfeffff08000000    mov dword ptr [ebp+fffffeec], 00000008

' *** Reference to "__vbaVarCat"
'0071a574    ff1508124000            call dword ptr [00401208]
'0071a57a    50                      push eax

' *** Reference to "rtcMsgBox"
'0071a57b    ff15b8104000            call dword ptr [004010b8]
var_1690 = MsgBox(#NOT SUPPORTED#, 0)
'0071a581    8d4db8                  lea ecx, dword ptr [ebp-48]

' *** Reference to "__vbaFreeObj"
'0071a584    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_61)
'0071a58a    8d8d5cffffff            lea ecx, dword ptr [ebp+ffffff5c]
'0071a590    51                      push ecx
'0071a591    8d956cffffff            lea edx, dword ptr [ebp+ffffff6c]
'0071a597    52                      push edx
'0071a598    8d857cffffff            lea eax, dword ptr [ebp+ffffff7c]
'0071a59e    50                      push eax
'0071a59f    8d4d8c                  lea ecx, dword ptr [ebp-74]
'0071a5a2    51                      push ecx
'0071a5a3    8d559c                  lea edx, dword ptr [ebp-64]
'0071a5a6    52                      push edx
'0071a5a7    6a05                    push 05

' *** Reference to "__vbaFreeVarList"
'0071a5a9    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_51, var_52, var_59, var_94, var_88)
'0071a5af    83c418                  add esp, 18

'ERROR: Two many next close:
End If
'0071a5b2    8b7508                  mov esi, dword ptr [ebp+08]
'0071a5b5    b801000000              mov eax, 00000001
'0071a5ba    6603c7                  add ax, di
var_num1 = 1 + 1
'0071a5bd    0f80b0020000            jo 71a873
'0071a5c3    8bf8                    mov edi, eax
'0071a5c5    897de8                  mov dword ptr [ebp-18], edi
'0071a5c8    e920ccffff              jmp 7171ed

'ERROR: Two many next close:
Loop
'0071a5cd    8b45d4                  mov eax, dword ptr [ebp-2c]
'0071a5d0    8b08                    mov ecx, dword ptr [eax]
'0071a5d2    50                      push eax
'0071a5d3    ff91c4000000            call dword ptr [ecx+000000c4]
'0071a5d9    dbe2                    fnclex
'0071a5db    85c0                    test ax, ax
'0071a5dd    7d15                    jge 71a5f4
'0071a5df    8b55d4                  mov edx, dword ptr [ebp-2c]
'0071a5e2    68c4000000              push 000000c4
'0071a5e7    6830314300              push 00433130
'0071a5ec    52                      push edx
'0071a5ed    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0071a5ee    ff1580104000            call dword ptr [00401080]
'0071a5f4    8b45d0                  mov eax, dword ptr [ebp-30]
'0071a5f7    8b08                    mov ecx, dword ptr [eax]
'0071a5f9    50                      push eax
'0071a5fa    ff91c4000000            call dword ptr [ecx+000000c4]
'0071a600    dbe2                    fnclex
'0071a602    85c0                    test ax, ax
'0071a604    7d15                    jge 71a61b
'0071a606    8b55d0                  mov edx, dword ptr [ebp-30]
'0071a609    68c4000000              push 000000c4
'0071a60e    6830314300              push 00433130
'0071a613    52                      push edx
'0071a614    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0071a615    ff1580104000            call dword ptr [00401080]
'0071a61b    a124be7200              mov ax, word ptr [0072be24]
'0071a620    85c0                    test ax, ax
'0071a622    7514                    jne 71a638

' *** Reference to "__vbaNew2"
'0071a624    8b1d2c124000            mov ebx, dword ptr [0040122c]
'0071a62a    6824be7200              push 0072be24
'0071a62f    6870004300              push 00430070
'0071a634    ffd3                    call ebx
Dim var_2 As New Global
'0071a636    eb06                    jmp 71a63e

' *** Reference to "__vbaNew2"
'0071a638    8b1d2c124000            mov ebx, dword ptr [0040122c]
'0071a63e    8b3d24be7200            mov edi, dword ptr [0072be24]
'0071a644    8b17                    mov edx, dword ptr [edi]
'0071a646    56                      push esi
'0071a647    8d45b8                  lea eax, dword ptr [ebp-48]
'0071a64a    50                      push eax
'0071a64b    8995e0fcffff            mov dword ptr [ebp+fffffce0], edx

' *** Reference to "__vbaObjSetAddref"
'0071a651    ff15c8104000            call dword ptr [004010c8]
Set var_61 = Me
'0071a657    8b8de0fcffff            mov ecx, dword ptr [ebp+fffffce0]
'0071a65d    50                      push eax
'0071a65e    57                      push edi
'0071a65f    ff5110                  call dword ptr [ecx+10]
Call var_2.Unload(var_61)
'0071a662    dbe2                    fnclex
'0071a664    85c0                    test ax, ax
'0071a666    7d0f                    jge 71a677
'0071a668    6a10                    push 10
'0071a66a    6860004300              push 00430060
'0071a66f    57                      push edi
'0071a670    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0071a671    ff1580104000            call dword ptr [00401080]
'0071a677    8d4db8                  lea ecx, dword ptr [ebp-48]

' *** Reference to "__vbaFreeObj"
'0071a67a    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_61)
'0071a680    a110b07200              mov ax, word ptr [0072b010]
'0071a685    85c0                    test ax, ax
'0071a687    750c                    jne 71a695
'0071a689    6810b07200              push 0072b010
'0071a68e    68b0e54000              push 0040e5b0
'0071a693    ffd3                    call ebx
Dim var_60 As New frmMain
'0071a695    8b3510b07200            mov esi, dword ptr [0072b010]
'0071a69b    baf8064300              mov edx, 004306f8
'0071a6a0    8d4dcc                  lea ecx, dword ptr [ebp-34]

' *** Reference to "__vbaStrCopy"
'0071a6a3    ff1548124000            call dword ptr [00401248]
var_43 = ("frmAcceder")
'0071a6a9    8b16                    mov edx, dword ptr [esi]
'0071a6ab    8d8538feffff            lea eax, dword ptr [ebp+fffffe38]
'0071a6b1    50                      push eax
'0071a6b2    8d4dcc                  lea ecx, dword ptr [ebp-34]
'0071a6b5    51                      push ecx
'0071a6b6    56                      push esi
'0071a6b7    ff92f8060000            call dword ptr [edx+000006f8]
'0071a6bd    dbe2                    fnclex
'0071a6bf    85c0                    test ax, ax
'0071a6c1    7d12                    jge 71a6d5
'0071a6c3    68f8060000              push 000006f8
'0071a6c8    687cfd4200              push 0042fd7c
'0071a6cd    56                      push esi
'0071a6ce    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0071a6cf    ff1580104000            call dword ptr [00401080]
'0071a6d5    33d2                    xor edx, edx
var_num4 = Empty
'0071a6d7    6683bd38feffffff        cmp word ptr [ebp+fffffe38], ffffffff
'0071a6df    8d4dcc                  lea ecx, dword ptr [ebp-34]
'0071a6e2    0f94c2                  sete dl
'0071a6e5    f7da                    neg edx
'0071a6e7    8bf2                    mov esi, edx

' *** Reference to "__vbaFreeStr"
'0071a6e9    ff150c134000            call dword ptr [0040130c]
'#Cleanup(var_43)
'0071a6ef    6685f6                  test esi, esi
'0071a6f2    a15cb07200              mov ax, word ptr [0072b05c]
'0071a6f7    0f8481000000            je 71a77e

If (0 = -1) Then
'0071a6fd    85c0                    test ax, ax
'0071a6ff    750c                    jne 71a70d
'0071a701    685cb07200              push 0072b05c
'0071a706    68347b4000              push 00407b34
'0071a70b    ffd3                    call ebx
    Dim var_24 As New frmAcceder
'0071a70d    8b355cb07200            mov esi, dword ptr [0072b05c]
'0071a713    8b3e                    mov edi, dword ptr [esi]
'0071a715    83ec10                  sub esp, 10
'0071a718    8bdc                    mov ebx, esp
'0071a71a    b90a000000              mov ecx, 0000000a
'0071a71f    890b                    mov dword ptr [ebx], ecx
'0071a721    898d4cffffff            mov dword ptr [ebp+ffffff4c], ecx
'0071a727    8b8d40ffffff            mov ecx, dword ptr [ebp+ffffff40]
'0071a72d    894b04                  mov dword ptr [ebx+04], ecx
'0071a730    b804000280              mov eax, 80020004
'0071a735    894308                  mov dword ptr [ebx+08], eax
'0071a738    8bd0                    mov edx, eax
'0071a73a    8b8548ffffff            mov eax, dword ptr [ebp+ffffff48]
'0071a740    89430c                  mov dword ptr [ebx+0c], eax
'0071a743    8b854cffffff            mov eax, dword ptr [ebp+ffffff4c]
'0071a749    83ec10                  sub esp, 10
'0071a74c    8bcc                    mov ecx, esp
'0071a74e    8901                    mov dword ptr [ecx], eax
'0071a750    8b8550ffffff            mov eax, dword ptr [ebp+ffffff50]
'0071a756    894104                  mov dword ptr [ecx+04], eax
'0071a759    895108                  mov dword ptr [ecx+08], edx
'0071a75c    8b9558ffffff            mov edx, dword ptr [ebp+ffffff58]
'0071a762    56                      push esi
'0071a763    89510c                  mov dword ptr [ecx+0c], edx
'0071a766    ff97b0020000            call dword ptr [edi+000002b0]
'0071a76c    dbe2                    fnclex
'0071a76e    85c0                    test ax, ax
'0071a770    7d43                    jge 71a7b5
'0071a772    68b0020000              push 000002b0
'0071a777    6810074300              push 00430710
'0071a77c    eb2f                    jmp 71a7ad
    
End If
'0071a77e    85c0                    test ax, ax
'0071a780    750c                    jne 71a78e
'0071a782    685cb07200              push 0072b05c
'0071a787    68347b4000              push 00407b34
'0071a78c    ffd3                    call ebx
Set var_24 = New frmAcceder
'0071a78e    8b355cb07200            mov esi, dword ptr [0072b05c]
'0071a794    8b06                    mov eax, dword ptr [esi]
'0071a796    56                      push esi
'0071a797    ff90fc060000            call dword ptr [eax+000006fc]
'0071a79d    dbe2                    fnclex
'0071a79f    85c0                    test ax, ax
'0071a7a1    7d12                    jge 71a7b5
'0071a7a3    68fc060000              push 000006fc
'0071a7a8    6840074300              push 00430740
'0071a7ad    56                      push esi
'0071a7ae    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0071a7af    ff1580104000            call dword ptr [00401080]
'0071a7b5    c745fc00000000          mov dword ptr [ebp-04], 00000000
'0071a7bc    6854a87100              push 0071a854
'0071a7c1    eb5d                    jmp 71a820
'0071a7c3    8d4dbc                  lea ecx, dword ptr [ebp-44]
'0071a7c6    51                      push ecx
'0071a7c7    8d55c0                  lea edx, dword ptr [ebp-40]
'0071a7ca    52                      push edx
'0071a7cb    8d45c4                  lea eax, dword ptr [ebp-3c]
'0071a7ce    50                      push eax
'0071a7cf    8d4dc8                  lea ecx, dword ptr [ebp-38]
'0071a7d2    51                      push ecx
'0071a7d3    8d55cc                  lea edx, dword ptr [ebp-34]
'0071a7d6    52                      push edx
'0071a7d7    6a05                    push 05

' *** Reference to "__vbaFreeStrList"
'0071a7d9    ff155c124000            call dword ptr [0040125c]
'#Cleanup( , -4924, -4928, -4932, var_58)
'0071a7df    8d45ac                  lea eax, dword ptr [ebp-54]
'0071a7e2    50                      push eax
'0071a7e3    8d4db0                  lea ecx, dword ptr [ebp-50]
'0071a7e6    51                      push ecx
'0071a7e7    8d55b4                  lea edx, dword ptr [ebp-4c]
'0071a7ea    52                      push edx
'0071a7eb    8d45b8                  lea eax, dword ptr [ebp-48]
'0071a7ee    50                      push eax
'0071a7ef    6a04                    push 04

' *** Reference to "__vbaFreeObjList"
'0071a7f1    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_61, var_62, var_6, var_50)
'0071a7f7    8d8d5cffffff            lea ecx, dword ptr [ebp+ffffff5c]
'0071a7fd    51                      push ecx
'0071a7fe    8d956cffffff            lea edx, dword ptr [ebp+ffffff6c]
'0071a804    52                      push edx
'0071a805    8d857cffffff            lea eax, dword ptr [ebp+ffffff7c]
'0071a80b    50                      push eax
'0071a80c    8d4d8c                  lea ecx, dword ptr [ebp-74]
'0071a80f    51                      push ecx
'0071a810    8d559c                  lea edx, dword ptr [ebp-64]
'0071a813    52                      push edx
'0071a814    6a05                    push 05

' *** Reference to "__vbaFreeVarList"
'0071a816    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_51, var_52, var_59, var_94, var_88)
'0071a81c    83c444                  add esp, 44
'0071a81f    c3                      ret
'0071a820    8d8510feffff            lea eax, dword ptr [ebp+fffffe10]
'0071a826    50                      push eax
'0071a827    8d8d14feffff            lea ecx, dword ptr [ebp+fffffe14]
'0071a82d    51                      push ecx
'0071a82e    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'0071a830    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_695, var_915)

' *** Reference to "__vbaFreeObj"
'0071a836    8b3508134000            mov esi, dword ptr [00401308]
'0071a83c    83c40c                  add esp, 0c
'0071a83f    8d4ddc                  lea ecx, dword ptr [ebp-24]
'0071a842    ffd6                    call esi
'#Cleanup(var_39)
'0071a844    8d4dd8                  lea ecx, dword ptr [ebp-28]
'0071a847    ffd6                    call esi
'#Cleanup(var_45)
'0071a849    8d4dd4                  lea ecx, dword ptr [ebp-2c]
'0071a84c    ffd6                    call esi
'#Cleanup(var_44)
'0071a84e    8d4dd0                  lea ecx, dword ptr [ebp-30]
'0071a851    ffd6                    call esi
'#Cleanup(var_4)
'0071a853    c3                      ret
'0071a854    8b4508                  mov eax, dword ptr [ebp+08]
'0071a857    8b10                    mov edx, dword ptr [eax]
'0071a859    50                      push eax
'0071a85a    ff5208                  call dword ptr [edx+08]
'0071a85d    8b45fc                  mov eax, dword ptr [ebp-04]
'0071a860    8b4dec                  mov ecx, dword ptr [ebp-14]
'0071a863    5f                      pop edi
'0071a864    5e                      pop esi
    'Reference to '__except_list'
'0071a865    64890d00000000          mov dword ptr fs:[00000000], ecx
'0071a86c    5b                      pop ebx
'0071a86d    8be5                    mov esp, ebp
'0071a86f    5d                      pop ebp
'0071a870    c20400                  ret 0004


End Sub


'Event for btnAucun
Private Sub btnAucun_Click()
'00714a30    55                      push ebp
'00714a31    8bec                    mov ebp, esp
'00714a33    83ec0c                  sub esp, 0c

' *** Reference to "__vbaExceptHandler"
'00714a36    6866724000              push 00407266
'00714a3b    64a100000000            mov ax, word ptr fs:[00000000]
'00714a41    50                      push eax
    'Reference to '__except_list'
'00714a42    64892500000000          mov dword ptr fs:[00000000], esp
'00714a49    83ec10                  sub esp, 10
'00714a4c    53                      push ebx
'00714a4d    56                      push esi
'00714a4e    57                      push edi
'00714a4f    8965f4                  mov dword ptr [ebp-0c], esp
'00714a52    c745f8b06d4000          mov dword ptr [ebp-08], 00406db0
'00714a59    8b7508                  mov esi, dword ptr [ebp+08]
'00714a5c    8bc6                    mov eax, esi
'00714a5e    83e001                  and eax, 01
'00714a61    8945fc                  mov dword ptr [ebp-04], eax
'00714a64    83e6fe                  and esi, fffffffe
'00714a67    8b0e                    mov ecx, dword ptr [esi]
'00714a69    56                      push esi
'00714a6a    897508                  mov dword ptr [ebp+08], esi
'00714a6d    ff5104                  call dword ptr [ecx+04]
'00714a70    33ff                    xor edi, edi
'00714a72    ba70164500              mov edx, 00451670
'00714a77    8d4de8                  lea ecx, dword ptr [ebp-18]
'00714a7a    897de8                  mov dword ptr [ebp-18], edi

' *** Reference to "__vbaStrCopy"
'00714a7d    ff1548124000            call dword ptr [00401248]
var_41 = ("Non")
'00714a83    8b16                    mov edx, dword ptr [esi]
'00714a85    8d45e8                  lea eax, dword ptr [ebp-18]
'00714a88    50                      push eax
'00714a89    56                      push esi
'00714a8a    ff920c070000            call dword ptr [edx+0000070c]
'00714a90    3bc7                    cmp eax, edi
'00714a92    7d12                    jge 714aa6
'00714a94    680c070000              push 0000070c
'00714a99    68601e4300              push 00431e60
'00714a9e    56                      push esi
'00714a9f    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00714aa0    ff1580104000            call dword ptr [00401080]
'00714aa6    8d4de8                  lea ecx, dword ptr [ebp-18]

' *** Reference to "__vbaFreeStr"
'00714aa9    ff150c134000            call dword ptr [0040130c]
'#Cleanup(var_41)
'00714aaf    897dfc                  mov dword ptr [ebp-04], edi
'00714ab2    68c44a7100              push 00714ac4
'00714ab7    eb0a                    jmp 714ac3
'00714ab9    8d4de8                  lea ecx, dword ptr [ebp-18]

' *** Reference to "__vbaFreeStr"
'00714abc    ff150c134000            call dword ptr [0040130c]
'#Cleanup(var_41)
'00714ac2    c3                      ret
'00714ac3    c3                      ret
'00714ac4    8b4508                  mov eax, dword ptr [ebp+08]
'00714ac7    8b08                    mov ecx, dword ptr [eax]
'00714ac9    50                      push eax
'00714aca    ff5108                  call dword ptr [ecx+08]
'00714acd    8b45fc                  mov eax, dword ptr [ebp-04]
'00714ad0    8b4dec                  mov ecx, dword ptr [ebp-14]
'00714ad3    5f                      pop edi
'00714ad4    5e                      pop esi
    'Reference to '__except_list'
'00714ad5    64890d00000000          mov dword ptr fs:[00000000], ecx
'00714adc    5b                      pop ebx
'00714add    8be5                    mov esp, ebp
'00714adf    5d                      pop ebp
'00714ae0    c20400                  ret 0004


End Sub


'Event for BtnFermer
Private Sub BtnFermer_Click()
'00714af0    55                      push ebp
'00714af1    8bec                    mov ebp, esp
'00714af3    83ec0c                  sub esp, 0c

' *** Reference to "__vbaExceptHandler"
'00714af6    6866724000              push 00407266
'00714afb    64a100000000            mov ax, word ptr fs:[00000000]
'00714b01    50                      push eax
    'Reference to '__except_list'
'00714b02    64892500000000          mov dword ptr fs:[00000000], esp
'00714b09    83ec18                  sub esp, 18
'00714b0c    53                      push ebx
'00714b0d    56                      push esi
'00714b0e    57                      push edi
'00714b0f    8965f4                  mov dword ptr [ebp-0c], esp
'00714b12    c745f8c06d4000          mov dword ptr [ebp-08], 00406dc0
'00714b19    8b7d08                  mov edi, dword ptr [ebp+08]
'00714b1c    8bc7                    mov eax, edi
'00714b1e    83e001                  and eax, 01
'00714b21    8945fc                  mov dword ptr [ebp-04], eax
'00714b24    83e7fe                  and edi, fffffffe
'00714b27    8b0f                    mov ecx, dword ptr [edi]
'00714b29    57                      push edi
'00714b2a    897d08                  mov dword ptr [ebp+08], edi
'00714b2d    ff5104                  call dword ptr [ecx+04]
'00714b30    a124be7200              mov ax, word ptr [0072be24]
'00714b35    33db                    xor ebx, ebx
'00714b37    3bc3                    cmp eax, ebx
'00714b39    895de8                  mov dword ptr [ebp-18], ebx
'00714b3c    7510                    jne 714b4e

If (0 = 0) Then
'00714b3e    6824be7200              push 0072be24
'00714b43    6870004300              push 00430070

' *** Reference to "__vbaNew2"
'00714b48    ff152c124000            call dword ptr [0040122c]
    Dim var_2 As New Global
End If
'00714b4e    8b3524be7200            mov esi, dword ptr [0072be24]
'00714b54    8b16                    mov edx, dword ptr [esi]
'00714b56    57                      push edi
'00714b57    8d45e8                  lea eax, dword ptr [ebp-18]
'00714b5a    50                      push eax
'00714b5b    8955d4                  mov dword ptr [ebp-2c], edx

' *** Reference to "__vbaObjSetAddref"
'00714b5e    ff15c8104000            call dword ptr [004010c8]
Set var_41 = Me
'00714b64    8b4dd4                  mov ecx, dword ptr [ebp-2c]
'00714b67    50                      push eax
'00714b68    56                      push esi
'00714b69    ff5110                  call dword ptr [ecx+10]
Call var_2.Unload(var_41)
'00714b6c    dbe2                    fnclex
'00714b6e    3bc3                    cmp eax, ebx
'00714b70    7d0f                    jge 714b81
'00714b72    6a10                    push 10
'00714b74    6860004300              push 00430060
'00714b79    56                      push esi
'00714b7a    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00714b7b    ff1580104000            call dword ptr [00401080]
'00714b81    8d4de8                  lea ecx, dword ptr [ebp-18]

' *** Reference to "__vbaFreeObj"
'00714b84    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_41)
'00714b8a    895dfc                  mov dword ptr [ebp-04], ebx
'00714b8d    689f4b7100              push 00714b9f
'00714b92    eb0a                    jmp 714b9e
'00714b94    8d4de8                  lea ecx, dword ptr [ebp-18]

' *** Reference to "__vbaFreeObj"
'00714b97    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_41)
'00714b9d    c3                      ret
'00714b9e    c3                      ret
'00714b9f    8b4508                  mov eax, dword ptr [ebp+08]
'00714ba2    8b10                    mov edx, dword ptr [eax]
'00714ba4    50                      push eax
'00714ba5    ff5208                  call dword ptr [edx+08]
'00714ba8    8b45fc                  mov eax, dword ptr [ebp-04]
'00714bab    8b4dec                  mov ecx, dword ptr [ebp-14]
'00714bae    5f                      pop edi
'00714baf    5e                      pop esi
    'Reference to '__except_list'
'00714bb0    64890d00000000          mov dword ptr fs:[00000000], ecx
'00714bb7    5b                      pop ebx
'00714bb8    8be5                    mov esp, ebp
'00714bba    5d                      pop ebp
'00714bbb    c20400                  ret 0004


End Sub


'Event for vsPersoImport
Private Sub vsPersoImport_Event37()
'0071e110    55                      push ebp
'0071e111    8bec                    mov ebp, esp
'0071e113    83ec14                  sub esp, 14

' *** Reference to "__vbaExceptHandler"
'0071e116    6866724000              push 00407266
'0071e11b    64a100000000            mov ax, word ptr fs:[00000000]
'0071e121    50                      push eax
    'Reference to '__except_list'
'0071e122    64892500000000          mov dword ptr fs:[00000000], esp
'0071e129    81ec2c010000            sub esp, 0000012c
'0071e12f    53                      push ebx
'0071e130    56                      push esi
'0071e131    57                      push edi
'0071e132    8965ec                  mov dword ptr [ebp-14], esp
'0071e135    c745f0206f4000          mov dword ptr [ebp-10], 00406f20
'0071e13c    8b7d08                  mov edi, dword ptr [ebp+08]
'0071e13f    8bc7                    mov eax, edi
'0071e141    83e001                  and eax, 01
'0071e144    8945f4                  mov dword ptr [ebp-0c], eax
'0071e147    83e7fe                  and edi, fffffffe
'0071e14a    897d08                  mov dword ptr [ebp+08], edi
'0071e14d    33f6                    xor esi, esi
'0071e14f    8975f8                  mov dword ptr [ebp-08], esi
'0071e152    8b0f                    mov ecx, dword ptr [edi]
'0071e154    57                      push edi
'0071e155    ff5104                  call dword ptr [ecx+04]
'0071e158    8975e0                  mov dword ptr [ebp-20], esi
'0071e15b    8975dc                  mov dword ptr [ebp-24], esi
'0071e15e    8975d8                  mov dword ptr [ebp-28], esi
'0071e161    8975d4                  mov dword ptr [ebp-2c], esi
'0071e164    8975d0                  mov dword ptr [ebp-30], esi
'0071e167    8975cc                  mov dword ptr [ebp-34], esi
'0071e16a    8975c8                  mov dword ptr [ebp-38], esi
'0071e16d    8975c4                  mov dword ptr [ebp-3c], esi
'0071e170    8975c0                  mov dword ptr [ebp-40], esi
'0071e173    8975b0                  mov dword ptr [ebp-50], esi
'0071e176    8975a0                  mov dword ptr [ebp-60], esi
'0071e179    897590                  mov dword ptr [ebp-70], esi
'0071e17c    897580                  mov dword ptr [ebp-80], esi
'0071e17f    89b570ffffff            mov dword ptr [ebp+ffffff70], esi
'0071e185    89b560ffffff            mov dword ptr [ebp+ffffff60], esi
'0071e18b    89b550ffffff            mov dword ptr [ebp+ffffff50], esi
'0071e191    89b540ffffff            mov dword ptr [ebp+ffffff40], esi
'0071e197    89b530ffffff            mov dword ptr [ebp+ffffff30], esi
'0071e19d    89b520ffffff            mov dword ptr [ebp+ffffff20], esi
'0071e1a3    89b510ffffff            mov dword ptr [ebp+ffffff10], esi
'0071e1a9    89b500ffffff            mov dword ptr [ebp+ffffff00], esi
'0071e1af    89b5f0feffff            mov dword ptr [ebp+fffffef0], esi
'0071e1b5    89b5e0feffff            mov dword ptr [ebp+fffffee0], esi
'0071e1bb    89b5dcfeffff            mov dword ptr [ebp+fffffedc], esi
'0071e1c1    66393528b07200          cmp word ptr [0072b028], si
'0071e1c8    7508                    jne 71e1d2
'0071e1ca    6a01                    push 01

' *** Reference to "__vbaOnError"
'0071e1cc    ff15b0104000            call dword ptr [004010b0]
On Error Goto handler_0
'0071e1d2    39750c                  cmp dword ptr [ebp+0c], esi
'0071e1d5    0f8e15040000            jle 71e5f0

If (arg_0 > 0) Then
'0071e1db    8b4510                  mov eax, dword ptr [ebp+10]
'0071e1de    83f801                  cmp eax, 01
'0071e1e1    0f84b1030000            je 71e598
    
    If (    arg_1 <> 1) Then
'0071e1e7    83f805                  cmp eax, 05
'0071e1ea    0f84a8030000            je 71e598
    
    If (    arg_1 <> 5) Then
'0071e1f0    b904000280              mov ecx, 80020004
'0071e1f5    894d88                  mov dword ptr [ebp-78], ecx
'0071e1f8    b80a000000              mov eax, 0000000a
'0071e1fd    894580                  mov dword ptr [ebp-80], eax
'0071e200    894d98                  mov dword ptr [ebp-68], ecx
'0071e203    894590                  mov dword ptr [ebp-70], eax
'0071e206    894da8                  mov dword ptr [ebp-58], ecx
'0071e209    8945a0                  mov dword ptr [ebp-60], eax
'0071e20c    c78528ffffffa0664500    mov dword ptr [ebp+ffffff28], 004566a0
'0071e216    c78520ffffff08000000    mov dword ptr [ebp+ffffff20], 00000008
'0071e220    8d9520ffffff            lea edx, dword ptr [ebp+ffffff20]
'0071e226    8d4db0                  lea ecx, dword ptr [ebp-50]

' *** Reference to "__vbaVarDup"
'0071e229    ff1598124000            call dword ptr [00401298]
    var_6 = ("Vous ne pouvez pas changer cette céllule")
'0071e22f    8d4580                  lea eax, dword ptr [ebp-80]
'0071e232    50                      push eax
'0071e233    8d4d90                  lea ecx, dword ptr [ebp-70]
'0071e236    51                      push ecx
'0071e237    8d55a0                  lea edx, dword ptr [ebp-60]
'0071e23a    52                      push edx
'0071e23b    56                      push esi
'0071e23c    8d45b0                  lea eax, dword ptr [ebp-50]
'0071e23f    50                      push eax

' *** Reference to "rtcMsgBox"
'0071e240    ff15b8104000            call dword ptr [004010b8]
    var_63 = MsgBox(var_6, 0)
'0071e246    8d4d80                  lea ecx, dword ptr [ebp-80]
'0071e249    51                      push ecx
'0071e24a    8d5590                  lea edx, dword ptr [ebp-70]
'0071e24d    52                      push edx
'0071e24e    8d45a0                  lea eax, dword ptr [ebp-60]
'0071e251    50                      push eax
'0071e252    8d4db0                  lea ecx, dword ptr [ebp-50]
'0071e255    51                      push ecx
'0071e256    6a04                    push 04

' *** Reference to "__vbaFreeVarList"
'0071e258    ff1540104000            call dword ptr [00401040]
    '#Cleanup( var_6, var_7, var_8, var_18)
'0071e25e    83c414                  add esp, 14
'0071e261    e98a030000              jmp 71e5f0

' *** Reference to "rtcErrObj"
'0071e266    8b1d6c124000            mov ebx, dword ptr [0040126c]
'0071e26c    ffd3                    call ebx
'0071e26e    50                      push eax
'0071e26f    8d55c4                  lea edx, dword ptr [ebp-3c]
'0071e272    52                      push edx

' *** Reference to "__vbaObjSet"
'0071e273    8b3db4104000            mov edi, dword ptr [004010b4]
'0071e279    ffd7                    call edi
    Set var_9 = Err
'0071e27b    8bf0                    mov esi, eax
'0071e27d    8b06                    mov eax, dword ptr [esi]
'0071e27f    8d4de0                  lea ecx, dword ptr [ebp-20]
'0071e282    51                      push ecx
'0071e283    56                      push esi
'0071e284    ff502c                  call dword ptr [eax+2c]
    var_3 = var_9.Description
'0071e287    dbe2                    fnclex
'0071e289    85c0                    test ax, ax
'0071e28b    7d0f                    jge 71e29c
'0071e28d    6a2c                    push 2c
'0071e28f    685c084300              push 0043085c
'0071e294    56                      push esi
'0071e295    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0071e296    ff1580104000            call dword ptr [00401080]
'0071e29c    ffd3                    call ebx
'0071e29e    50                      push eax
'0071e29f    8d55c0                  lea edx, dword ptr [ebp-40]
'0071e2a2    52                      push edx
'0071e2a3    ffd7                    call edi
    Set var_5 = Err
'0071e2a5    8bf0                    mov esi, eax
'0071e2a7    8b06                    mov eax, dword ptr [esi]
'0071e2a9    8d8ddcfeffff            lea ecx, dword ptr [ebp+fffffedc]
'0071e2af    51                      push ecx
'0071e2b0    56                      push esi
'0071e2b1    ff501c                  call dword ptr [eax+1c]
    var_10 = var_5.Number
'0071e2b4    dbe2                    fnclex
'0071e2b6    85c0                    test ax, ax
'0071e2b8    7d0f                    jge 71e2c9
'0071e2ba    6a1c                    push 1c
'0071e2bc    685c084300              push 0043085c
'0071e2c1    56                      push esi
'0071e2c2    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0071e2c3    ff1580104000            call dword ptr [00401080]
'0071e2c9    b904000280              mov ecx, 80020004
'0071e2ce    894d98                  mov dword ptr [ebp-68], ecx
'0071e2d1    b80a000000              mov eax, 0000000a
'0071e2d6    894590                  mov dword ptr [ebp-70], eax
'0071e2d9    894da8                  mov dword ptr [ebp-58], ecx
'0071e2dc    8945a0                  mov dword ptr [ebp-60], eax
'0071e2df    c78528ffffff24b07200    mov dword ptr [ebp+ffffff28], 0072b024
'0071e2e9    c78520ffffff08400000    mov dword ptr [ebp+ffffff20], 00004008
'0071e2f3    6814084300              push 00430814
'0071e2f8    8b55e0                  mov edx, dword ptr [ebp-20]
'0071e2fb    52                      push edx

' *** Reference to "__vbaStrCat"
'0071e2fc    8b3570104000            mov esi, dword ptr [00401070]
'0071e302    ffd6                    call esi
    var_11 = ("L'erreur suivante s'est produite : ") & (var_3)
'0071e304    8bd0                    mov edx, eax
'0071e306    8d4ddc                  lea ecx, dword ptr [ebp-24]

' *** Reference to "__vbaStrMove"
'0071e309    8b3dd0124000            mov edi, dword ptr [004012d0]
'0071e30f    ffd7                    call edi
'0071e311    50                      push eax
'0071e312    6870084300              push 00430870
'0071e317    ffd6                    call esi
    var_14 = (var_11) & (vbCrLf)
'0071e319    8bd0                    mov edx, eax
'0071e31b    8d4dd8                  lea ecx, dword ptr [ebp-28]
'0071e31e    ffd7                    call edi
'0071e320    50                      push eax
'0071e321    6870084300              push 00430870
'0071e326    ffd6                    call esi
    var_127 = (var_14) & (vbCrLf)
'0071e328    8bd0                    mov edx, eax
'0071e32a    8d4dd4                  lea ecx, dword ptr [ebp-2c]
'0071e32d    ffd7                    call edi
'0071e32f    50                      push eax
'0071e330    6880084300              push 00430880
'0071e335    ffd6                    call esi
    var_15 = (var_127) & (" numéro d'erreur (")
'0071e337    8bd0                    mov edx, eax
'0071e339    8d4dd0                  lea ecx, dword ptr [ebp-30]
'0071e33c    ffd7                    call edi
'0071e33e    50                      push eax
'0071e33f    8b85dcfeffff            mov eax, dword ptr [ebp+fffffedc]
'0071e345    50                      push eax

' *** Reference to "__vbaStrI4"
'0071e346    ff1520104000            call dword ptr [00401020]
'0071e34c    8bd0                    mov edx, eax
'0071e34e    8d4dcc                  lea ecx, dword ptr [ebp-34]
'0071e351    ffd7                    call edi
'0071e353    50                      push eax
'0071e354    ffd6                    call esi
    var_128 = (var_15) & (CStr(var_10))
'0071e356    8bd0                    mov edx, eax
'0071e358    8d4dc8                  lea ecx, dword ptr [ebp-38]
'0071e35b    ffd7                    call edi
'0071e35d    50                      push eax
'0071e35e    68ac084300              push 004308ac
'0071e363    ffd6                    call esi
    var_17 = (var_128) & (" )")
'0071e365    8945b8                  mov dword ptr [ebp-48], eax
'0071e368    bf08000000              mov edi, 00000008
'0071e36d    897db0                  mov dword ptr [ebp-50], edi
'0071e370    8d4d90                  lea ecx, dword ptr [ebp-70]
'0071e373    51                      push ecx
'0071e374    8d55a0                  lea edx, dword ptr [ebp-60]
'0071e377    52                      push edx
'0071e378    8d8520ffffff            lea eax, dword ptr [ebp+ffffff20]
'0071e37e    50                      push eax
'0071e37f    6a10                    push 10
'0071e381    8d4db0                  lea ecx, dword ptr [ebp-50]
'0071e384    51                      push ecx

' *** Reference to "rtcMsgBox"
'0071e385    ff15b8104000            call dword ptr [004010b8]
    var_53 = MsgBox(var_17, 16, vbNullString)
'0071e38b    8d55c8                  lea edx, dword ptr [ebp-38]
'0071e38e    52                      push edx
'0071e38f    8d45cc                  lea eax, dword ptr [ebp-34]
'0071e392    50                      push eax
'0071e393    8d4dd0                  lea ecx, dword ptr [ebp-30]
'0071e396    51                      push ecx
'0071e397    8d55d4                  lea edx, dword ptr [ebp-2c]
'0071e39a    52                      push edx
'0071e39b    8d45d8                  lea eax, dword ptr [ebp-28]
'0071e39e    50                      push eax
'0071e39f    8d4ddc                  lea ecx, dword ptr [ebp-24]
'0071e3a2    51                      push ecx
'0071e3a3    8d55e0                  lea edx, dword ptr [ebp-20]
'0071e3a6    52                      push edx
'0071e3a7    6a07                    push 07

' *** Reference to "__vbaFreeStrList"
'0071e3a9    ff155c124000            call dword ptr [0040125c]
    '#Cleanup( 0, -4520, -4524, -4528, -4532, -4536, -4540)
'0071e3af    8d45c0                  lea eax, dword ptr [ebp-40]
'0071e3b2    50                      push eax
'0071e3b3    8d4dc4                  lea ecx, dword ptr [ebp-3c]
'0071e3b6    51                      push ecx
'0071e3b7    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'0071e3b9    ff154c104000            call dword ptr [0040104c]
    '#Cleanup( var_9, var_5)
'0071e3bf    8d5590                  lea edx, dword ptr [ebp-70]
'0071e3c2    52                      push edx
'0071e3c3    8d45a0                  lea eax, dword ptr [ebp-60]
'0071e3c6    50                      push eax
'0071e3c7    8d4db0                  lea ecx, dword ptr [ebp-50]
'0071e3ca    51                      push ecx
'0071e3cb    6a03                    push 03

' *** Reference to "__vbaFreeVarList"
'0071e3cd    ff1540104000            call dword ptr [00401040]
    '#Cleanup( var_6, var_7, var_8)
'0071e3d3    83c43c                  add esp, 3c
'0071e3d6    8d55b0                  lea edx, dword ptr [ebp-50]
'0071e3d9    52                      push edx

' *** Reference to "rtcGetPresentDate"
'0071e3da    ff15f4124000            call dword ptr [004012f4]
    var_17 = Now()
'0071e3e0    c78528ffffffb8084300    mov dword ptr [ebp+ffffff28], 004308b8
'0071e3ea    89bd20ffffff            mov dword ptr [ebp+ffffff20], edi
'0071e3f0    8d9520ffffff            lea edx, dword ptr [ebp+ffffff20]
'0071e3f6    8d4da0                  lea ecx, dword ptr [ebp-60]

' *** Reference to "__vbaVarDup"
'0071e3f9    ff1598124000            call dword ptr [00401298]
    var_7 = ("dd/MM/yyyy hh:mm:ss")
'0071e3ff    6a01                    push 01
'0071e401    6a01                    push 01
'0071e403    8d45a0                  lea eax, dword ptr [ebp-60]
'0071e406    50                      push eax
'0071e407    8d4db0                  lea ecx, dword ptr [ebp-50]
'0071e40a    51                      push ecx
'0071e40b    8d5590                  lea edx, dword ptr [ebp-70]
'0071e40e    52                      push edx

' *** Reference to "rtcVarFromFormatVar"
'0071e40f    ff1574104000            call dword ptr [00401074]
'0071e415    ffd3                    call ebx
'0071e417    50                      push eax
'0071e418    8d45c4                  lea eax, dword ptr [ebp-3c]
'0071e41b    50                      push eax

' *** Reference to "__vbaObjSet"
'0071e41c    ff15b4104000            call dword ptr [004010b4]
    Set var_9 = Err
'0071e422    8bf0                    mov esi, eax
'0071e424    8b0e                    mov ecx, dword ptr [esi]
'0071e426    8d95dcfeffff            lea edx, dword ptr [ebp+fffffedc]
'0071e42c    52                      push edx
'0071e42d    56                      push esi
'0071e42e    ff511c                  call dword ptr [ecx+1c]
    var_10 = var_9.Number
'0071e431    dbe2                    fnclex
'0071e433    85c0                    test ax, ax
'0071e435    7d0f                    jge 71e446
'0071e437    6a1c                    push 1c
'0071e439    685c084300              push 0043085c
'0071e43e    56                      push esi
'0071e43f    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0071e440    ff1580104000            call dword ptr [00401080]
'0071e446    ffd3                    call ebx
'0071e448    50                      push eax
'0071e449    8d45c0                  lea eax, dword ptr [ebp-40]
'0071e44c    50                      push eax

' *** Reference to "__vbaObjSet"
'0071e44d    ff15b4104000            call dword ptr [004010b4]
    Set var_5 = Err
'0071e453    8bf0                    mov esi, eax
'0071e455    8b0e                    mov ecx, dword ptr [esi]
'0071e457    8d55e0                  lea edx, dword ptr [ebp-20]
'0071e45a    52                      push edx
'0071e45b    56                      push esi
'0071e45c    ff512c                  call dword ptr [ecx+2c]
    var_3 = var_5.Description
'0071e45f    dbe2                    fnclex
'0071e461    85c0                    test ax, ax
'0071e463    7d0f                    jge 71e474
'0071e465    6a2c                    push 2c
'0071e467    685c084300              push 0043085c
'0071e46c    56                      push esi
'0071e46d    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0071e46e    ff1580104000            call dword ptr [00401080]
'0071e474    c78518ffffffe4084300    mov dword ptr [ebp+ffffff18], 004308e4
'0071e47e    89bd10ffffff            mov dword ptr [ebp+ffffff10], edi
'0071e484    8b85dcfeffff            mov eax, dword ptr [ebp+fffffedc]
'0071e48a    898508ffffff            mov dword ptr [ebp+ffffff08], eax
'0071e490    c78500ffffff03000000    mov dword ptr [ebp+ffffff00], 00000003
'0071e49a    c785f8feffff08094300    mov dword ptr [ebp+fffffef8], 00430908
'0071e4a4    89bdf0feffff            mov dword ptr [ebp+fffffef0], edi
'0071e4aa    8b45e0                  mov eax, dword ptr [ebp-20]
'0071e4ad    c745e000000000          mov dword ptr [ebp-20], 00000000
'0071e4b4    898558ffffff            mov dword ptr [ebp+ffffff58], eax
'0071e4ba    89bd50ffffff            mov dword ptr [ebp+ffffff50], edi
'0071e4c0    c785e8fefffff06d4500    mov dword ptr [ebp+fffffee8], 00456df0
'0071e4ca    89bde0feffff            mov dword ptr [ebp+fffffee0], edi
'0071e4d0    8d4d90                  lea ecx, dword ptr [ebp-70]
'0071e4d3    51                      push ecx
'0071e4d4    8d9510ffffff            lea edx, dword ptr [ebp+ffffff10]
'0071e4da    52                      push edx
'0071e4db    8d4580                  lea eax, dword ptr [ebp-80]
'0071e4de    50                      push eax

' *** Reference to "__vbaVarCat"
'0071e4df    8b3508124000            mov esi, dword ptr [00401208]
'0071e4e5    ffd6                    call esi
'0071e4e7    50                      push eax
'0071e4e8    8d8d00ffffff            lea ecx, dword ptr [ebp+ffffff00]
'0071e4ee    51                      push ecx
'0071e4ef    8d9570ffffff            lea edx, dword ptr [ebp+ffffff70]
'0071e4f5    52                      push edx
'0071e4f6    ffd6                    call esi
'0071e4f8    50                      push eax
'0071e4f9    8d85f0feffff            lea eax, dword ptr [ebp+fffffef0]
'0071e4ff    50                      push eax
'0071e500    8d8d60ffffff            lea ecx, dword ptr [ebp+ffffff60]
'0071e506    51                      push ecx
'0071e507    ffd6                    call esi
'0071e509    50                      push eax
'0071e50a    8d9550ffffff            lea edx, dword ptr [ebp+ffffff50]
'0071e510    52                      push edx
'0071e511    8d8540ffffff            lea eax, dword ptr [ebp+ffffff40]
'0071e517    50                      push eax
'0071e518    ffd6                    call esi
'0071e51a    50                      push eax
'0071e51b    8d8de0feffff            lea ecx, dword ptr [ebp+fffffee0]
'0071e521    51                      push ecx
'0071e522    8d9530ffffff            lea edx, dword ptr [ebp+ffffff30]
'0071e528    52                      push edx
'0071e529    ffd6                    call esi
'0071e52b    50                      push eax
'0071e52c    33c0                    xor eax, eax
'0071e52e    66a12ab07200            mov eax, dword ptr [0072b02a]
'0071e534    50                      push eax
'0071e535    6884094300              push 00430984

' *** Reference to "__vbaPrintFile"
'0071e53a    ff15b8114000            call dword ptr [004011b8]
    Print #0, 
'0071e540    8d4dc0                  lea ecx, dword ptr [ebp-40]
'0071e543    51                      push ecx
'0071e544    8d55c4                  lea edx, dword ptr [ebp-3c]
'0071e547    52                      push edx
'0071e548    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'0071e54a    ff154c104000            call dword ptr [0040104c]
    '#Cleanup( var_9, var_5)
'0071e550    8d8530ffffff            lea eax, dword ptr [ebp+ffffff30]
'0071e556    50                      push eax
'0071e557    8d8d40ffffff            lea ecx, dword ptr [ebp+ffffff40]
'0071e55d    51                      push ecx
'0071e55e    8d9550ffffff            lea edx, dword ptr [ebp+ffffff50]
'0071e564    52                      push edx
'0071e565    8d8560ffffff            lea eax, dword ptr [ebp+ffffff60]
'0071e56b    50                      push eax
'0071e56c    8d8d70ffffff            lea ecx, dword ptr [ebp+ffffff70]
'0071e572    51                      push ecx
'0071e573    8d5580                  lea edx, dword ptr [ebp-80]
'0071e576    52                      push edx
'0071e577    8d4590                  lea eax, dword ptr [ebp-70]
'0071e57a    50                      push eax
'0071e57b    8d4da0                  lea ecx, dword ptr [ebp-60]
'0071e57e    51                      push ecx
'0071e57f    8d55b0                  lea edx, dword ptr [ebp-50]
'0071e582    52                      push edx
'0071e583    6a09                    push 09

' *** Reference to "__vbaFreeVarList"
'0071e585    ff1540104000            call dword ptr [00401040]
    '#Cleanup( var_6, var_7, var_8, var_18, var_19, var_20, var_21, var_22, var_23)
'0071e58b    83c440                  add esp, 40
'0071e58e    6a00                    push 00

' *** Reference to "__vbaResume"
'0071e590    ff1568104000            call dword ptr [00401068]
'0071e596    eb58                    jmp 71e5f0
    Resume handler_71E5F0
End If
'0071e598    b8cc134300              mov eax, 004313cc
'0071e59d    898528ffffff            mov dword ptr [ebp+ffffff28], eax
'0071e5a3    b908000000              mov ecx, 00000008
'0071e5a8    898d20ffffff            mov dword ptr [ebp+ffffff20], ecx
'0071e5ae    83ec10                  sub esp, 10
'0071e5b1    8bd4                    mov edx, esp
'0071e5b3    890a                    mov dword ptr [edx], ecx
'0071e5b5    8b8d24ffffff            mov ecx, dword ptr [ebp+ffffff24]
'0071e5bb    894a04                  mov dword ptr [edx+04], ecx
'0071e5be    894208                  mov dword ptr [edx+08], eax
'0071e5c1    8b852cffffff            mov eax, dword ptr [ebp+ffffff2c]
'0071e5c7    89420c                  mov dword ptr [edx+0c], eax
'0071e5ca    6a48                    push 48
'0071e5cc    8b0f                    mov ecx, dword ptr [edi]
'0071e5ce    57                      push edi
'0071e5cf    ff9110030000            call dword ptr [ecx+00000310]
'0071e5d5    50                      push eax
'0071e5d6    8d55c4                  lea edx, dword ptr [ebp-3c]
'0071e5d9    52                      push edx

' *** Reference to "__vbaObjSet"
'0071e5da    ff15b4104000            call dword ptr [004010b4]
Set var_9 = Nothing
'0071e5e0    50                      push eax

' *** Reference to "__vbaLateIdSt"
'0071e5e1    ff15ec124000            call dword ptr [004012ec]
var_9.[UNMANAGED] = vbNullChar
'0071e5e7    8d4dc4                  lea ecx, dword ptr [ebp-3c]

' *** Reference to "__vbaFreeObj"
'0071e5ea    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_9)
'ERROR: Two many next close:
End If

' *** Reference to "__vbaExitProc"
'0071e5f0    ff15a0104000            call dword ptr [004010a0]
'0071e5f6    6871e67100              push 0071e671
'0071e5fb    eb73                    jmp 71e670
'0071e5fd    8d45c8                  lea eax, dword ptr [ebp-38]
'0071e600    50                      push eax
'0071e601    8d4dcc                  lea ecx, dword ptr [ebp-34]
'0071e604    51                      push ecx
'0071e605    8d55d0                  lea edx, dword ptr [ebp-30]
'0071e608    52                      push edx
'0071e609    8d45d4                  lea eax, dword ptr [ebp-2c]
'0071e60c    50                      push eax
'0071e60d    8d4dd8                  lea ecx, dword ptr [ebp-28]
'0071e610    51                      push ecx
'0071e611    8d55dc                  lea edx, dword ptr [ebp-24]
'0071e614    52                      push edx
'0071e615    8d45e0                  lea eax, dword ptr [ebp-20]
'0071e618    50                      push eax
'0071e619    6a07                    push 07

' *** Reference to "__vbaFreeStrList"
'0071e61b    ff155c124000            call dword ptr [0040125c]
'#Cleanup( var_3, -4520, -4524, -4528, -4532, -4536, -4540)
'0071e621    8d4dc0                  lea ecx, dword ptr [ebp-40]
'0071e624    51                      push ecx
'0071e625    8d55c4                  lea edx, dword ptr [ebp-3c]
'0071e628    52                      push edx
'0071e629    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'0071e62b    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_9, var_5)
'0071e631    8d8530ffffff            lea eax, dword ptr [ebp+ffffff30]
'0071e637    50                      push eax
'0071e638    8d8d40ffffff            lea ecx, dword ptr [ebp+ffffff40]
'0071e63e    51                      push ecx
'0071e63f    8d9550ffffff            lea edx, dword ptr [ebp+ffffff50]
'0071e645    52                      push edx
'0071e646    8d8560ffffff            lea eax, dword ptr [ebp+ffffff60]
'0071e64c    50                      push eax
'0071e64d    8d8d70ffffff            lea ecx, dword ptr [ebp+ffffff70]
'0071e653    51                      push ecx
'0071e654    8d5580                  lea edx, dword ptr [ebp-80]
'0071e657    52                      push edx
'0071e658    8d4590                  lea eax, dword ptr [ebp-70]
'0071e65b    50                      push eax
'0071e65c    8d4da0                  lea ecx, dword ptr [ebp-60]
'0071e65f    51                      push ecx
'0071e660    8d55b0                  lea edx, dword ptr [ebp-50]
'0071e663    52                      push edx
'0071e664    6a09                    push 09

' *** Reference to "__vbaFreeVarList"
'0071e666    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_6, var_7, var_8, var_18, var_19, var_20, var_21, var_22, var_23)
'0071e66c    83c454                  add esp, 54
'0071e66f    c3                      ret
'0071e670    c3                      ret
'0071e671    8b4508                  mov eax, dword ptr [ebp+08]
'0071e674    8b08                    mov ecx, dword ptr [eax]
'0071e676    50                      push eax
'0071e677    ff5108                  call dword ptr [ecx+08]
'0071e67a    8b45f4                  mov eax, dword ptr [ebp-0c]
'0071e67d    8b4de4                  mov ecx, dword ptr [ebp-1c]
    'Reference to '__except_list'
'0071e680    64890d00000000          mov dword ptr fs:[00000000], ecx
'0071e687    5f                      pop edi
'0071e688    5e                      pop esi
'0071e689    5b                      pop ebx
'0071e68a    8be5                    mov esp, ebp
'0071e68c    5d                      pop ebp
'0071e68d    c21000                  ret 0010


End Sub


