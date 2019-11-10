VERSION 5.00
Object = "{D76D7130-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "C:\WINDOWS\SysWow64\Vsflex7d.ocx"

Begin VB.Form frmAcceder
    Caption = "Accéder à un personnage"
    ScaleMode = 1
    AutoRedraw = 0              'False
    FontTransparent = -1              'True
    LinkTopic = "Form1"
    MDIChild = -1              'True
    ClientLeft   = 6420
    ClientTop    = 4725
    ClientWidth  = 6660
    ClientHeight = 3840
    BeginProperty Font
        Name          = "Times New Roman"
        Size          = 9,75
        Charset       = 0
        Weight        = 400
        Underline     = 0              'False
        Italic        = 0              'False
        Strikethrough = 0              'False
    EndProperty
    Begin VSFlex7DAOCtl.VSFlexGrid vsnom
        Left   = 0
        Top    = 0
        Width  = 6360
        Height = 3825
        TabIndex = 0
        OleObjectBlob = frmAcceder.frx:0000
    End
End
Public Function RempPerso(arg_0 As Unknow, arg_1 As Unknow, arg_2 As Unknow, arg_3 As Unknow, arg_4 As Unknow, arg_5 As Unknow, arg_6 As Unknow, arg_7 As Unknow, arg_8 As Unknow, arg_9 As Unknow, arg_A As Unknow, arg_B As Unknow, arg_C As Unknow, arg_D As Unknow, arg_E As Unknow, arg_F As Unknow, arg_10 As Unknow, arg_11 As Unknow, arg_12 As Unknow, arg_13 As Unknow, arg_14 As Unknow, arg_15 As Unknow, arg_16 As Unknow, arg_17 As Unknow, arg_18 As Unknow, arg_19 As Unknow, arg_1A As Unknow, arg_1B As Unknow, arg_1C As Unknow, arg_1D As Unknow, arg_1E As Unknow, arg_1F As Unknow, arg_20 As Unknow, arg_21 As Unknow, arg_22 As Unknow, arg_23 As Unknow, arg_24 As Unknow, arg_25 As Unknow, arg_26 As Unknow, arg_27 As Unknow, arg_28 As Unknow, arg_29 As Unknow, arg_2A As Unknow, arg_2B As Unknow, arg_2C As Unknow, arg_2D As Unknow, arg_2E As Unknow, arg_2F As Unknow, arg_30 As Unknow, arg_31 As Unknow, arg_32 As Unknow, arg_33 As Unknow, arg_34 As Unknow, arg_35 As Unknow, arg_36 As Unknow, arg_37 As Unknow, arg_38 As Unknow, arg_39 As Unknow, arg_3A As Unknow, arg_3B As Unknow, arg_3C As Unknow)
'00633cd0    55                      push ebp
'00633cd1    8bec                    mov ebp, esp
'00633cd3    83ec14                  sub esp, 14

' *** Reference to "__vbaExceptHandler"
'00633cd6    6866724000              push 00407266
'00633cdb    64a100000000            mov ax, word ptr fs:[00000000]
'00633ce1    50                      push eax
    'Reference to '__except_list'
'00633ce2    64892500000000          mov dword ptr fs:[00000000], esp
'00633ce9    81ecf4010000            sub esp, 000001f4
'00633cef    53                      push ebx
'00633cf0    56                      push esi
'00633cf1    57                      push edi
'00633cf2    8965ec                  mov dword ptr [ebp-14], esp
'00633cf5    c745f0e8544000          mov dword ptr [ebp-10], 004054e8
'00633cfc    33ff                    xor edi, edi
'00633cfe    897df4                  mov dword ptr [ebp-0c], edi
'00633d01    897df8                  mov dword ptr [ebp-08], edi
'00633d04    8b7508                  mov esi, dword ptr [ebp+08]
'00633d07    8b06                    mov eax, dword ptr [esi]
'00633d09    56                      push esi
'00633d0a    ff5004                  call dword ptr [eax+04]
'00633d0d    897ddc                  mov dword ptr [ebp-24], edi
'00633d10    897dc0                  mov dword ptr [ebp-40], edi
'00633d13    897db4                  mov dword ptr [ebp-4c], edi
'00633d16    897db0                  mov dword ptr [ebp-50], edi
'00633d19    897dac                  mov dword ptr [ebp-54], edi
'00633d1c    897da8                  mov dword ptr [ebp-58], edi
'00633d1f    897da4                  mov dword ptr [ebp-5c], edi
'00633d22    897da0                  mov dword ptr [ebp-60], edi
'00633d25    897d9c                  mov dword ptr [ebp-64], edi
'00633d28    897d98                  mov dword ptr [ebp-68], edi
'00633d2b    897d94                  mov dword ptr [ebp-6c], edi
'00633d2e    897d90                  mov dword ptr [ebp-70], edi
'00633d31    897d8c                  mov dword ptr [ebp-74], edi
'00633d34    897d88                  mov dword ptr [ebp-78], edi
'00633d37    89bd78ffffff            mov dword ptr [ebp+ffffff78], edi
'00633d3d    89bd68ffffff            mov dword ptr [ebp+ffffff68], edi
'00633d43    89bd58ffffff            mov dword ptr [ebp+ffffff58], edi
'00633d49    89bd48ffffff            mov dword ptr [ebp+ffffff48], edi
'00633d4f    89bd38ffffff            mov dword ptr [ebp+ffffff38], edi
'00633d55    89bd28ffffff            mov dword ptr [ebp+ffffff28], edi
'00633d5b    89bd18ffffff            mov dword ptr [ebp+ffffff18], edi
'00633d61    89bd08ffffff            mov dword ptr [ebp+ffffff08], edi
'00633d67    89bdf8feffff            mov dword ptr [ebp+fffffef8], edi
'00633d6d    89bde8feffff            mov dword ptr [ebp+fffffee8], edi
'00633d73    89bdd8feffff            mov dword ptr [ebp+fffffed8], edi
'00633d79    89bdc8feffff            mov dword ptr [ebp+fffffec8], edi
'00633d7f    89bdb8feffff            mov dword ptr [ebp+fffffeb8], edi
'00633d85    89bda8feffff            mov dword ptr [ebp+fffffea8], edi
'00633d8b    89bd98feffff            mov dword ptr [ebp+fffffe98], edi
'00633d91    89bd88feffff            mov dword ptr [ebp+fffffe88], edi
'00633d97    89bd74feffff            mov dword ptr [ebp+fffffe74], edi
'00633d9d    89bd70feffff            mov dword ptr [ebp+fffffe70], edi
'00633da3    66393d28b07200          cmp word ptr [0072b028], di
'00633daa    7508                    jne 633db4
'00633dac    6a01                    push 01

' *** Reference to "__vbaOnError"
'00633dae    ff15b0104000            call dword ptr [004010b0]
On Error Goto handler_0
'00633db4    b801000000              mov eax, 00000001
'00633db9    8985f0feffff            mov dword ptr [ebp+fffffef0], eax
'00633dbf    b903000000              mov ecx, 00000003
'00633dc4    898de8feffff            mov dword ptr [ebp+fffffee8], ecx
'00633dca    83ec10                  sub esp, 10
'00633dcd    8bd4                    mov edx, esp
'00633dcf    890a                    mov dword ptr [edx], ecx
'00633dd1    8b8decfeffff            mov ecx, dword ptr [ebp+fffffeec]
'00633dd7    894a04                  mov dword ptr [edx+04], ecx
'00633dda    894208                  mov dword ptr [edx+08], eax
'00633ddd    8b85f4feffff            mov eax, dword ptr [ebp+fffffef4]
'00633de3    89420c                  mov dword ptr [edx+0c], eax
'00633de6    6a07                    push 07
'00633de8    8b0e                    mov ecx, dword ptr [esi]
'00633dea    56                      push esi
'00633deb    ff91fc020000            call dword ptr [ecx+000002fc]
'00633df1    50                      push eax
'00633df2    8d5598                  lea edx, dword ptr [ebp-68]
'00633df5    52                      push edx

' *** Reference to "__vbaObjSet"
'00633df6    ff15b4104000            call dword ptr [004010b4]
Set var_130 = Nothing
'00633dfc    50                      push eax

' *** Reference to "__vbaLateIdSt"
'00633dfd    ff15ec124000            call dword ptr [004012ec]
var_130.[UNMANAGED] = 1
'00633e03    8d4d98                  lea ecx, dword ptr [ebp-68]

' *** Reference to "__vbaFreeObj"
'00633e06    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_130)
'00633e0c    ba04000280              mov edx, 80020004
'00633e11    8bc2                    mov eax, edx
'00633e13    8985d0feffff            mov dword ptr [ebp+fffffed0], eax
'00633e19    be0a000000              mov esi, 0000000a
'00633e1e    8bce                    mov ecx, esi
'00633e20    898dc8feffff            mov dword ptr [ebp+fffffec8], ecx
'00633e26    8995e0feffff            mov dword ptr [ebp+fffffee0], edx
'00633e2c    89b5d8feffff            mov dword ptr [ebp+fffffed8], esi
'00633e32    c785f0feffff02000000    mov dword ptr [ebp+fffffef0], 00000002
'00633e3c    c785e8feffff03000000    mov dword ptr [ebp+fffffee8], 00000003
'00633e46    8b3d48b07200            mov edi, dword ptr [0072b048]
'00633e4c    8b3f                    mov edi, dword ptr [edi]
'00633e4e    8d5d98                  lea ebx, dword ptr [ebp-68]
'00633e51    53                      push ebx
'00633e52    83ec10                  sub esp, 10
'00633e55    8bdc                    mov ebx, esp
'00633e57    890b                    mov dword ptr [ebx], ecx
'00633e59    8b8dccfeffff            mov ecx, dword ptr [ebp+fffffecc]
'00633e5f    894b04                  mov dword ptr [ebx+04], ecx
'00633e62    894308                  mov dword ptr [ebx+08], eax
'00633e65    8b85d4feffff            mov eax, dword ptr [ebp+fffffed4]
'00633e6b    89430c                  mov dword ptr [ebx+0c], eax
'00633e6e    83ec10                  sub esp, 10
'00633e71    8bcc                    mov ecx, esp
'00633e73    8931                    mov dword ptr [ecx], esi
'00633e75    8b85dcfeffff            mov eax, dword ptr [ebp+fffffedc]
'00633e7b    894104                  mov dword ptr [ecx+04], eax
'00633e7e    895108                  mov dword ptr [ecx+08], edx
'00633e81    8b95e4feffff            mov edx, dword ptr [ebp+fffffee4]
'00633e87    89510c                  mov dword ptr [ecx+0c], edx
'00633e8a    83ec10                  sub esp, 10
'00633e8d    8bc4                    mov eax, esp
'00633e8f    8b8de8feffff            mov ecx, dword ptr [ebp+fffffee8]
'00633e95    8908                    mov dword ptr [eax], ecx
'00633e97    8b95ecfeffff            mov edx, dword ptr [ebp+fffffeec]
'00633e9d    895004                  mov dword ptr [eax+04], edx
'00633ea0    8b8df0feffff            mov ecx, dword ptr [ebp+fffffef0]
'00633ea6    894808                  mov dword ptr [eax+08], ecx
'00633ea9    8b95f4feffff            mov edx, dword ptr [ebp+fffffef4]
'00633eaf    89500c                  mov dword ptr [eax+0c], edx
'00633eb2    68d8db4400              push 0044dbd8
'00633eb7    a148b07200              mov ax, word ptr [0072b048]
'00633ebc    50                      push eax
'00633ebd    ff97bc000000            call dword ptr [edi+000000bc]
'00633ec3    dbe2                    fnclex
'00633ec5    85c0                    test ax, ax
'00633ec7    0f8d97030000            jge 634264
'00633ecd    68bc000000              push 000000bc
'00633ed2    68301f4300              push 00431f30
'00633ed7    8b0d48b07200            mov ecx, dword ptr [0072b048]
'00633edd    51                      push ecx
'00633ede    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00633edf    8b1d80104000            mov ebx, dword ptr [00401080]
'00633ee5    ffd3                    call ebx
'00633ee7    e97e030000              jmp 63426a

' *** Reference to "rtcErrObj"
'00633eec    8b1d6c124000            mov ebx, dword ptr [0040126c]
'00633ef2    ffd3                    call ebx
'00633ef4    50                      push eax
'00633ef5    8d4d98                  lea ecx, dword ptr [ebp-68]
'00633ef8    51                      push ecx

' *** Reference to "__vbaObjSet"
'00633ef9    8b3db4104000            mov edi, dword ptr [004010b4]
'00633eff    ffd7                    call edi
Set var_130 = Err
'00633f01    8bf0                    mov esi, eax
'00633f03    8b16                    mov edx, dword ptr [esi]
'00633f05    8d45b4                  lea eax, dword ptr [ebp-4c]
'00633f08    50                      push eax
'00633f09    56                      push esi
'00633f0a    ff522c                  call dword ptr [edx+2c]
var_62 = var_130.Description
'00633f0d    dbe2                    fnclex
'00633f0f    85c0                    test ax, ax
'00633f11    7d0f                    jge 633f22
'00633f13    6a2c                    push 2c
'00633f15    685c084300              push 0043085c
'00633f1a    56                      push esi
'00633f1b    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00633f1c    ff1580104000            call dword ptr [00401080]
'00633f22    ffd3                    call ebx
'00633f24    50                      push eax
'00633f25    8d4d94                  lea ecx, dword ptr [ebp-6c]
'00633f28    51                      push ecx
'00633f29    ffd7                    call edi
Set var_121 = Err
'00633f2b    8bf0                    mov esi, eax
'00633f2d    8b16                    mov edx, dword ptr [esi]
'00633f2f    8d8570feffff            lea eax, dword ptr [ebp+fffffe70]
'00633f35    50                      push eax
'00633f36    56                      push esi
'00633f37    ff521c                  call dword ptr [edx+1c]
var_74 = var_121.Number
'00633f3a    dbe2                    fnclex
'00633f3c    85c0                    test ax, ax
'00633f3e    7d0f                    jge 633f4f
'00633f40    6a1c                    push 1c
'00633f42    685c084300              push 0043085c
'00633f47    56                      push esi
'00633f48    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00633f49    ff1580104000            call dword ptr [00401080]
'00633f4f    ba04000280              mov edx, 80020004
'00633f54    899560ffffff            mov dword ptr [ebp+ffffff60], edx
'00633f5a    be0a000000              mov esi, 0000000a
'00633f5f    89b558ffffff            mov dword ptr [ebp+ffffff58], esi
'00633f65    899570ffffff            mov dword ptr [ebp+ffffff70], edx
'00633f6b    89b568ffffff            mov dword ptr [ebp+ffffff68], esi
'00633f71    c785f0feffff24b07200    mov dword ptr [ebp+fffffef0], 0072b024
'00633f7b    c785e8feffff08400000    mov dword ptr [ebp+fffffee8], 00004008
'00633f85    6814084300              push 00430814
'00633f8a    8b4db4                  mov ecx, dword ptr [ebp-4c]
'00633f8d    51                      push ecx

' *** Reference to "__vbaStrCat"
'00633f8e    8b3570104000            mov esi, dword ptr [00401070]
'00633f94    ffd6                    call esi
var_49 = ("L'erreur suivante s'est produite : ") & (var_62)
'00633f96    8bd0                    mov edx, eax
'00633f98    8d4db0                  lea ecx, dword ptr [ebp-50]

' *** Reference to "__vbaStrMove"
'00633f9b    8b3dd0124000            mov edi, dword ptr [004012d0]
'00633fa1    ffd7                    call edi
'00633fa3    50                      push eax
'00633fa4    6870084300              push 00430870
'00633fa9    ffd6                    call esi
var_127 = (var_49) & (vbCrLf)
'00633fab    8bd0                    mov edx, eax
'00633fad    8d4dac                  lea ecx, dword ptr [ebp-54]
'00633fb0    ffd7                    call edi
'00633fb2    50                      push eax
'00633fb3    6870084300              push 00430870
'00633fb8    ffd6                    call esi
var_15 = (var_127) & (vbCrLf)
'00633fba    8bd0                    mov edx, eax
'00633fbc    8d4da8                  lea ecx, dword ptr [ebp-58]
'00633fbf    ffd7                    call edi
'00633fc1    50                      push eax
'00633fc2    6880084300              push 00430880
'00633fc7    ffd6                    call esi
var_16 = (var_15) & (" numéro d'erreur (")
'00633fc9    8bd0                    mov edx, eax
'00633fcb    8d4da4                  lea ecx, dword ptr [ebp-5c]
'00633fce    ffd7                    call edi
'00633fd0    50                      push eax
'00633fd1    8b9570feffff            mov edx, dword ptr [ebp+fffffe70]
'00633fd7    52                      push edx

' *** Reference to "__vbaStrI4"
'00633fd8    ff1520104000            call dword ptr [00401020]
'00633fde    8bd0                    mov edx, eax
'00633fe0    8d4da0                  lea ecx, dword ptr [ebp-60]
'00633fe3    ffd7                    call edi
'00633fe5    50                      push eax
'00633fe6    ffd6                    call esi
var_17 = (var_16) & (CStr(var_74))
'00633fe8    8bd0                    mov edx, eax
'00633fea    8d4d9c                  lea ecx, dword ptr [ebp-64]
'00633fed    ffd7                    call edi
'00633fef    50                      push eax
'00633ff0    68ac084300              push 004308ac
'00633ff5    ffd6                    call esi
var_129 = (var_17) & (" )")
'00633ff7    894580                  mov dword ptr [ebp-80], eax
'00633ffa    bf08000000              mov edi, 00000008
'00633fff    89bd78ffffff            mov dword ptr [ebp+ffffff78], edi
'00634005    8d8558ffffff            lea eax, dword ptr [ebp+ffffff58]
'0063400b    50                      push eax
'0063400c    8d8d68ffffff            lea ecx, dword ptr [ebp+ffffff68]
'00634012    51                      push ecx
'00634013    8d95e8feffff            lea edx, dword ptr [ebp+fffffee8]
'00634019    52                      push edx
'0063401a    6a10                    push 10
'0063401c    8d8578ffffff            lea eax, dword ptr [ebp+ffffff78]
'00634022    50                      push eax

' *** Reference to "rtcMsgBox"
'00634023    ff15b8104000            call dword ptr [004010b8]
var_285 = MsgBox(var_129, 16, vbNullString)
'00634029    8d4d9c                  lea ecx, dword ptr [ebp-64]
'0063402c    51                      push ecx
'0063402d    8d55a0                  lea edx, dword ptr [ebp-60]
'00634030    52                      push edx
'00634031    8d45a4                  lea eax, dword ptr [ebp-5c]
'00634034    50                      push eax
'00634035    8d4da8                  lea ecx, dword ptr [ebp-58]
'00634038    51                      push ecx
'00634039    8d55ac                  lea edx, dword ptr [ebp-54]
'0063403c    52                      push edx
'0063403d    8d45b0                  lea eax, dword ptr [ebp-50]
'00634040    50                      push eax
'00634041    8d4db4                  lea ecx, dword ptr [ebp-4c]
'00634044    51                      push ecx
'00634045    6a07                    push 07

' *** Reference to "__vbaFreeStrList"
'00634047    ff155c124000            call dword ptr [0040125c]
'#Cleanup( 0, -4524, -4528, -4532, -4536, -4540, -4544)
'0063404d    8d5594                  lea edx, dword ptr [ebp-6c]
'00634050    52                      push edx
'00634051    8d4598                  lea eax, dword ptr [ebp-68]
'00634054    50                      push eax
'00634055    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'00634057    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_130, var_121)
'0063405d    8d8d58ffffff            lea ecx, dword ptr [ebp+ffffff58]
'00634063    51                      push ecx
'00634064    8d9568ffffff            lea edx, dword ptr [ebp+ffffff68]
'0063406a    52                      push edx
'0063406b    8d8578ffffff            lea eax, dword ptr [ebp+ffffff78]
'00634071    50                      push eax
'00634072    6a03                    push 03

' *** Reference to "__vbaFreeVarList"
'00634074    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_87, var_132, var_133)
'0063407a    83c43c                  add esp, 3c
'0063407d    8d8d78ffffff            lea ecx, dword ptr [ebp+ffffff78]
'00634083    51                      push ecx

' *** Reference to "rtcGetPresentDate"
'00634084    ff15f4124000            call dword ptr [004012f4]
var_129 = Now()
'0063408a    c785f0feffffb8084300    mov dword ptr [ebp+fffffef0], 004308b8
'00634094    89bde8feffff            mov dword ptr [ebp+fffffee8], edi
'0063409a    8d95e8feffff            lea edx, dword ptr [ebp+fffffee8]
'006340a0    8d8d68ffffff            lea ecx, dword ptr [ebp+ffffff68]

' *** Reference to "__vbaVarDup"
'006340a6    ff1598124000            call dword ptr [00401298]
var_132 = ("dd/MM/yyyy hh:mm:ss")
'006340ac    6a01                    push 01
'006340ae    6a01                    push 01
'006340b0    8d9568ffffff            lea edx, dword ptr [ebp+ffffff68]
'006340b6    52                      push edx
'006340b7    8d8578ffffff            lea eax, dword ptr [ebp+ffffff78]
'006340bd    50                      push eax
'006340be    8d8d58ffffff            lea ecx, dword ptr [ebp+ffffff58]
'006340c4    51                      push ecx

' *** Reference to "rtcVarFromFormatVar"
'006340c5    ff1574104000            call dword ptr [00401074]
'006340cb    ffd3                    call ebx
'006340cd    50                      push eax
'006340ce    8d5598                  lea edx, dword ptr [ebp-68]
'006340d1    52                      push edx

' *** Reference to "__vbaObjSet"
'006340d2    ff15b4104000            call dword ptr [004010b4]
Set var_130 = Err
'006340d8    8bf0                    mov esi, eax
'006340da    8b06                    mov eax, dword ptr [esi]
'006340dc    8d8d70feffff            lea ecx, dword ptr [ebp+fffffe70]
'006340e2    51                      push ecx
'006340e3    56                      push esi
'006340e4    ff501c                  call dword ptr [eax+1c]
var_74 = var_130.Number
'006340e7    dbe2                    fnclex
'006340e9    85c0                    test ax, ax
'006340eb    7d0f                    jge 6340fc
'006340ed    6a1c                    push 1c
'006340ef    685c084300              push 0043085c
'006340f4    56                      push esi
'006340f5    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006340f6    ff1580104000            call dword ptr [00401080]
'006340fc    ffd3                    call ebx
'006340fe    50                      push eax
'006340ff    8d5594                  lea edx, dword ptr [ebp-6c]
'00634102    52                      push edx

' *** Reference to "__vbaObjSet"
'00634103    ff15b4104000            call dword ptr [004010b4]
Set var_121 = Err
'00634109    8bf0                    mov esi, eax
'0063410b    8b06                    mov eax, dword ptr [esi]
'0063410d    8d4db4                  lea ecx, dword ptr [ebp-4c]
'00634110    51                      push ecx
'00634111    56                      push esi
'00634112    ff502c                  call dword ptr [eax+2c]
var_62 = var_121.Description
'00634115    dbe2                    fnclex
'00634117    85c0                    test ax, ax
'00634119    7d0f                    jge 63412a
'0063411b    6a2c                    push 2c
'0063411d    685c084300              push 0043085c
'00634122    56                      push esi
'00634123    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00634124    ff1580104000            call dword ptr [00401080]
'0063412a    c785e0feffffe4084300    mov dword ptr [ebp+fffffee0], 004308e4
'00634134    89bdd8feffff            mov dword ptr [ebp+fffffed8], edi
'0063413a    8b9570feffff            mov edx, dword ptr [ebp+fffffe70]
'00634140    8995d0feffff            mov dword ptr [ebp+fffffed0], edx
'00634146    c785c8feffff03000000    mov dword ptr [ebp+fffffec8], 00000003
'00634150    c785c0feffff08094300    mov dword ptr [ebp+fffffec0], 00430908
'0063415a    89bdb8feffff            mov dword ptr [ebp+fffffeb8], edi
'00634160    8b45b4                  mov eax, dword ptr [ebp-4c]
'00634163    c745b400000000          mov dword ptr [ebp-4c], 00000000
'0063416a    898520ffffff            mov dword ptr [ebp+ffffff20], eax
'00634170    89bd18ffffff            mov dword ptr [ebp+ffffff18], edi
'00634176    c785b0feffff50dd4400    mov dword ptr [ebp+fffffeb0], 0044dd50
'00634180    89bda8feffff            mov dword ptr [ebp+fffffea8], edi
'00634186    8d8558ffffff            lea eax, dword ptr [ebp+ffffff58]
'0063418c    50                      push eax
'0063418d    8d8dd8feffff            lea ecx, dword ptr [ebp+fffffed8]
'00634193    51                      push ecx
'00634194    8d9548ffffff            lea edx, dword ptr [ebp+ffffff48]
'0063419a    52                      push edx

' *** Reference to "__vbaVarCat"
'0063419b    8b3508124000            mov esi, dword ptr [00401208]
'006341a1    ffd6                    call esi
'006341a3    50                      push eax
'006341a4    8d85c8feffff            lea eax, dword ptr [ebp+fffffec8]
'006341aa    50                      push eax
'006341ab    8d8d38ffffff            lea ecx, dword ptr [ebp+ffffff38]
'006341b1    51                      push ecx
'006341b2    ffd6                    call esi
'006341b4    50                      push eax
'006341b5    8d95b8feffff            lea edx, dword ptr [ebp+fffffeb8]
'006341bb    52                      push edx
'006341bc    8d8528ffffff            lea eax, dword ptr [ebp+ffffff28]
'006341c2    50                      push eax
'006341c3    ffd6                    call esi
'006341c5    50                      push eax
'006341c6    8d8d18ffffff            lea ecx, dword ptr [ebp+ffffff18]
'006341cc    51                      push ecx
'006341cd    8d9508ffffff            lea edx, dword ptr [ebp+ffffff08]
'006341d3    52                      push edx
'006341d4    ffd6                    call esi
'006341d6    50                      push eax
'006341d7    8d85a8feffff            lea eax, dword ptr [ebp+fffffea8]
'006341dd    50                      push eax
'006341de    8d8df8feffff            lea ecx, dword ptr [ebp+fffffef8]
'006341e4    51                      push ecx
'006341e5    ffd6                    call esi
'006341e7    50                      push eax
'006341e8    33d2                    xor edx, edx
'006341ea    668b152ab07200          mov dx, word ptr [0072b02a]
'006341f1    52                      push edx
'006341f2    6884094300              push 00430984

' *** Reference to "__vbaPrintFile"
'006341f7    ff15b8114000            call dword ptr [004011b8]
Print #0, 
'006341fd    8d4594                  lea eax, dword ptr [ebp-6c]
'00634200    50                      push eax
'00634201    8d4d98                  lea ecx, dword ptr [ebp-68]
'00634204    51                      push ecx
'00634205    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'00634207    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_130, var_121)
'0063420d    8d95f8feffff            lea edx, dword ptr [ebp+fffffef8]
'00634213    52                      push edx
'00634214    8d8508ffffff            lea eax, dword ptr [ebp+ffffff08]
'0063421a    50                      push eax
'0063421b    8d8d18ffffff            lea ecx, dword ptr [ebp+ffffff18]
'00634221    51                      push ecx
'00634222    8d9528ffffff            lea edx, dword ptr [ebp+ffffff28]
'00634228    52                      push edx
'00634229    8d8538ffffff            lea eax, dword ptr [ebp+ffffff38]
'0063422f    50                      push eax
'00634230    8d8d48ffffff            lea ecx, dword ptr [ebp+ffffff48]
'00634236    51                      push ecx
'00634237    8d9558ffffff            lea edx, dword ptr [ebp+ffffff58]
'0063423d    52                      push edx
'0063423e    8d8568ffffff            lea eax, dword ptr [ebp+ffffff68]
'00634244    50                      push eax
'00634245    8d8d78ffffff            lea ecx, dword ptr [ebp+ffffff78]
'0063424b    51                      push ecx
'0063424c    6a09                    push 09

' *** Reference to "__vbaFreeVarList"
'0063424e    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_87, var_132, var_133, var_134, var_135, var_136, var_849, var_310, var_298)
'00634254    83c440                  add esp, 40
'00634257    6a00                    push 00

' *** Reference to "__vbaResume"
'00634259    ff1568104000            call dword ptr [00401068]
'0063425f    e9311b0000              jmp 635d95
Resume handler_635D95

' *** Reference to "__vbaHresultCheckObj"
'00634264    8b1d80104000            mov ebx, dword ptr [00401080]
'0063426a    8b4598                  mov eax, dword ptr [ebp-68]
'0063426d    c7459800000000          mov dword ptr [ebp-68], 00000000
'00634274    50                      push eax
'00634275    8d55c0                  lea edx, dword ptr [ebp-40]
'00634278    52                      push edx

' *** Reference to "__vbaObjSet"
'00634279    ff15b4104000            call dword ptr [004010b4]
Set var_5 = var_130
'0063427f    8b45c0                  mov eax, dword ptr [ebp-40]
'00634282    8b08                    mov ecx, dword ptr [eax]
'00634284    8d9574feffff            lea edx, dword ptr [ebp+fffffe74]
'0063428a    52                      push edx
'0063428b    50                      push eax
'0063428c    ff5134                  call dword ptr [ecx+34]
var_304 = var_5.HelpFile
'0063428f    dbe2                    fnclex
'00634291    85c0                    test ax, ax
'00634293    7d0e                    jge 6342a3
'00634295    6a34                    push 34
'00634297    6830314300              push 00433130
'0063429c    8b4dc0                  mov ecx, dword ptr [ebp-40]
'0063429f    51                      push ecx
'006342a0    50                      push eax
'006342a1    ffd3                    call ebx
'006342a3    6683bd74feffff00        cmp word ptr [ebp+fffffe74], 00
'006342ab    8b45c0                  mov eax, dword ptr [ebp-40]
'006342ae    8b10                    mov edx, dword ptr [eax]
'006342b0    0f85c11a0000            jne 635d77

Do While (0 = 0)
'006342b6    8d4d98                  lea ecx, dword ptr [ebp-68]
'006342b9    51                      push ecx
'006342ba    50                      push eax
'006342bb    ff92b4000000            call dword ptr [edx+000000b4]
'006342c1    dbe2                    fnclex
'006342c3    85c0                    test ax, ax
'006342c5    7d11                    jge 6342d8
'006342c7    68b4000000              push 000000b4
'006342cc    6830314300              push 00433130
'006342d1    8b55c0                  mov edx, dword ptr [ebp-40]
'006342d4    52                      push edx
'006342d5    50                      push eax
'006342d6    ffd3                    call ebx
'006342d8    8b4598                  mov eax, dword ptr [ebp-68]
'006342db    898568feffff            mov dword ptr [ebp+fffffe68], eax
'006342e1    b934d44300              mov ecx, 0043d434
'006342e6    898df0feffff            mov dword ptr [ebp+fffffef0], ecx
'006342ec    ba08000000              mov edx, 00000008
'006342f1    8995e8feffff            mov dword ptr [ebp+fffffee8], edx
'006342f7    8b30                    mov esi, dword ptr [eax]
'006342f9    8d7d94                  lea edi, dword ptr [ebp-6c]
'006342fc    57                      push edi
'006342fd    83ec10                  sub esp, 10
'00634300    8bfc                    mov edi, esp
'00634302    8917                    mov dword ptr [edi], edx
'00634304    8b95ecfeffff            mov edx, dword ptr [ebp+fffffeec]
'0063430a    895704                  mov dword ptr [edi+04], edx
'0063430d    894f08                  mov dword ptr [edi+08], ecx
'00634310    8b8df4feffff            mov ecx, dword ptr [ebp+fffffef4]
'00634316    894f0c                  mov dword ptr [edi+0c], ecx
'00634319    50                      push eax
'0063431a    ff5630                  call dword ptr [esi+30]
'0063431d    dbe2                    fnclex
'0063431f    85c0                    test ax, ax
'00634321    7d11                    jge 634334
'00634323    6a30                    push 30
'00634325    68d8304300              push 004330d8
'0063432a    8b9568feffff            mov edx, dword ptr [ebp+fffffe68]
'00634330    52                      push edx
'00634331    50                      push eax
'00634332    ffd3                    call ebx
'00634334    8b4594                  mov eax, dword ptr [ebp-6c]
'00634337    c7459400000000          mov dword ptr [ebp-6c], 00000000
'0063433e    894580                  mov dword ptr [ebp-80], eax
'00634341    c78578ffffff09000000    mov dword ptr [ebp+ffffff78], 00000009
'0063434b    8d8578ffffff            lea eax, dword ptr [ebp+ffffff78]
'00634351    50                      push eax

' *** Reference to "rtcIsNull"
'00634352    ff1540114000            call dword ptr [00401140]
'00634358    898574feffff            mov dword ptr [ebp+fffffe74], eax
'0063435e    8b45c0                  mov eax, dword ptr [ebp-40]
'00634361    8b08                    mov ecx, dword ptr [eax]
'00634363    8d5590                  lea edx, dword ptr [ebp-70]
'00634366    52                      push edx
'00634367    50                      push eax
'00634368    ff91b4000000            call dword ptr [ecx+000000b4]
'0063436e    dbe2                    fnclex
'00634370    85c0                    test ax, ax
'00634372    7d11                    jge 634385
'00634374    68b4000000              push 000000b4
'00634379    6830314300              push 00433130
'0063437e    8b4dc0                  mov ecx, dword ptr [ebp-40]
'00634381    51                      push ecx
'00634382    50                      push eax
'00634383    ffd3                    call ebx
'00634385    8b4590                  mov eax, dword ptr [ebp-70]
'00634388    89855cfeffff            mov dword ptr [ebp+fffffe5c], eax
'0063438e    b934d44300              mov ecx, 0043d434
'00634393    898de0feffff            mov dword ptr [ebp+fffffee0], ecx
'00634399    ba08000000              mov edx, 00000008
'0063439e    8995d8feffff            mov dword ptr [ebp+fffffed8], edx
'006343a4    8b30                    mov esi, dword ptr [eax]
'006343a6    8d7d8c                  lea edi, dword ptr [ebp-74]
'006343a9    57                      push edi
'006343aa    83ec10                  sub esp, 10
'006343ad    8bfc                    mov edi, esp
'006343af    8917                    mov dword ptr [edi], edx
'006343b1    8b95dcfeffff            mov edx, dword ptr [ebp+fffffedc]
'006343b7    895704                  mov dword ptr [edi+04], edx
'006343ba    894f08                  mov dword ptr [edi+08], ecx
'006343bd    8b8de4feffff            mov ecx, dword ptr [ebp+fffffee4]
'006343c3    894f0c                  mov dword ptr [edi+0c], ecx
'006343c6    50                      push eax
'006343c7    ff5630                  call dword ptr [esi+30]
'006343ca    dbe2                    fnclex
'006343cc    33f6                    xor esi, esi
    var_num8 = Empty
'006343ce    3bc6                    cmp eax, esi
'006343d0    7d11                    jge 6343e3
    
    If (    0 < -256 - 20) Then
'006343d2    6a30                    push 30
'006343d4    68d8304300              push 004330d8
'006343d9    8b955cfeffff            mov edx, dword ptr [ebp+fffffe5c]
'006343df    52                      push edx
'006343e0    50                      push eax
'006343e1    ffd3                    call ebx
    
End If
'006343e3    8b458c                  mov eax, dword ptr [ebp-74]
'006343e6    89758c                  mov dword ptr [ebp-74], esi
'006343e9    898560ffffff            mov dword ptr [ebp+ffffff60], eax
'006343ef    c78558ffffff09000000    mov dword ptr [ebp+ffffff58], 00000009
'006343f9    89b570ffffff            mov dword ptr [ebp+ffffff70], esi
'006343ff    c78568ffffff02000000    mov dword ptr [ebp+ffffff68], 00000002
'00634409    668b8574feffff          mov ax, word ptr [ebp+fffffe74]
'00634410    668985d0feffff          mov word ptr [ebp+fffffed0], ax
'00634417    c785c8feffff0b000000    mov dword ptr [ebp+fffffec8], 0000000b
'00634421    8d8d58ffffff            lea ecx, dword ptr [ebp+ffffff58]
'00634427    51                      push ecx
'00634428    8d9568ffffff            lea edx, dword ptr [ebp+ffffff68]
'0063442e    52                      push edx
'0063442f    8d85c8feffff            lea eax, dword ptr [ebp+fffffec8]
'00634435    50                      push eax
'00634436    8d8d48ffffff            lea ecx, dword ptr [ebp+ffffff48]
'0063443c    51                      push ecx

' *** Reference to "rtcImmediateIf"
'0063443d    ff154c124000            call dword ptr [0040124c]
'00634443    8d9548ffffff            lea edx, dword ptr [ebp+ffffff48]
'00634449    52                      push edx

' *** Reference to "__vbaI2Var"
'0063444a    ff150c124000            call dword ptr [0040120c]
'00634450    8945bc                  mov dword ptr [ebp-44], eax
'00634453    8d4590                  lea eax, dword ptr [ebp-70]
'00634456    50                      push eax
'00634457    8d4d98                  lea ecx, dword ptr [ebp-68]
'0063445a    51                      push ecx
'0063445b    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'0063445d    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_130, var_8)
'00634463    8d9548ffffff            lea edx, dword ptr [ebp+ffffff48]
'00634469    52                      push edx
'0063446a    8d8558ffffff            lea eax, dword ptr [ebp+ffffff58]
'00634470    50                      push eax
'00634471    8d8d68ffffff            lea ecx, dword ptr [ebp+ffffff68]
'00634477    51                      push ecx
'00634478    8d95c8feffff            lea edx, dword ptr [ebp+fffffec8]
'0063447e    52                      push edx
'0063447f    8d8578ffffff            lea eax, dword ptr [ebp+ffffff78]
'00634485    50                      push eax
'00634486    6a05                    push 05

' *** Reference to "__vbaFreeVarList"
'00634488    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_87, var_65, var_132, var_133, var_134)
'0063448e    83c424                  add esp, 24
'00634491    8b45c0                  mov eax, dword ptr [ebp-40]
'00634494    8b08                    mov ecx, dword ptr [eax]
'00634496    8d5598                  lea edx, dword ptr [ebp-68]
'00634499    52                      push edx
'0063449a    50                      push eax
'0063449b    ff91b4000000            call dword ptr [ecx+000000b4]
'006344a1    dbe2                    fnclex
'006344a3    3bc6                    cmp eax, esi
'006344a5    7d11                    jge 6344b8
'006344a7    68b4000000              push 000000b4
'006344ac    6830314300              push 00433130
'006344b1    8b4dc0                  mov ecx, dword ptr [ebp-40]
'006344b4    51                      push ecx
'006344b5    50                      push eax
'006344b6    ffd3                    call ebx
'006344b8    8b4598                  mov eax, dword ptr [ebp-68]
'006344bb    898568feffff            mov dword ptr [ebp+fffffe68], eax
'006344c1    b95cd44300              mov ecx, 0043d45c
'006344c6    898df0feffff            mov dword ptr [ebp+fffffef0], ecx
'006344cc    ba08000000              mov edx, 00000008
'006344d1    8995e8feffff            mov dword ptr [ebp+fffffee8], edx
'006344d7    8b30                    mov esi, dword ptr [eax]
'006344d9    8d7d94                  lea edi, dword ptr [ebp-6c]
'006344dc    57                      push edi
'006344dd    83ec10                  sub esp, 10
'006344e0    8bfc                    mov edi, esp
'006344e2    8917                    mov dword ptr [edi], edx
'006344e4    8b95ecfeffff            mov edx, dword ptr [ebp+fffffeec]
'006344ea    895704                  mov dword ptr [edi+04], edx
'006344ed    894f08                  mov dword ptr [edi+08], ecx
'006344f0    8b8df4feffff            mov ecx, dword ptr [ebp+fffffef4]
'006344f6    894f0c                  mov dword ptr [edi+0c], ecx
'006344f9    50                      push eax
'006344fa    ff5630                  call dword ptr [esi+30]
'006344fd    dbe2                    fnclex
'006344ff    85c0                    test ax, ax
'00634501    7d11                    jge 634514
'00634503    6a30                    push 30
'00634505    68d8304300              push 004330d8
'0063450a    8b9568feffff            mov edx, dword ptr [ebp+fffffe68]
'00634510    52                      push edx
'00634511    50                      push eax
'00634512    ffd3                    call ebx
'00634514    8b4594                  mov eax, dword ptr [ebp-6c]
'00634517    c7459400000000          mov dword ptr [ebp-6c], 00000000
'0063451e    894580                  mov dword ptr [ebp-80], eax
'00634521    c78578ffffff09000000    mov dword ptr [ebp+ffffff78], 00000009
'0063452b    8d8578ffffff            lea eax, dword ptr [ebp+ffffff78]
'00634531    50                      push eax

' *** Reference to "rtcIsNull"
'00634532    ff1540114000            call dword ptr [00401140]
'00634538    898574feffff            mov dword ptr [ebp+fffffe74], eax
'0063453e    8b45c0                  mov eax, dword ptr [ebp-40]
'00634541    8b08                    mov ecx, dword ptr [eax]
'00634543    8d5590                  lea edx, dword ptr [ebp-70]
'00634546    52                      push edx
'00634547    50                      push eax
'00634548    ff91b4000000            call dword ptr [ecx+000000b4]
'0063454e    dbe2                    fnclex
'00634550    85c0                    test ax, ax
'00634552    7d11                    jge 634565
'00634554    68b4000000              push 000000b4
'00634559    6830314300              push 00433130
'0063455e    8b4dc0                  mov ecx, dword ptr [ebp-40]
'00634561    51                      push ecx
'00634562    50                      push eax
'00634563    ffd3                    call ebx
'00634565    8b4590                  mov eax, dword ptr [ebp-70]
'00634568    89855cfeffff            mov dword ptr [ebp+fffffe5c], eax
'0063456e    b95cd44300              mov ecx, 0043d45c
'00634573    898de0feffff            mov dword ptr [ebp+fffffee0], ecx
'00634579    ba08000000              mov edx, 00000008
'0063457e    8995d8feffff            mov dword ptr [ebp+fffffed8], edx
'00634584    8b30                    mov esi, dword ptr [eax]
'00634586    8d7d8c                  lea edi, dword ptr [ebp-74]
'00634589    57                      push edi
'0063458a    83ec10                  sub esp, 10
'0063458d    8bfc                    mov edi, esp
'0063458f    8917                    mov dword ptr [edi], edx
'00634591    8b95dcfeffff            mov edx, dword ptr [ebp+fffffedc]
'00634597    895704                  mov dword ptr [edi+04], edx
'0063459a    894f08                  mov dword ptr [edi+08], ecx
'0063459d    8b8de4feffff            mov ecx, dword ptr [ebp+fffffee4]
'006345a3    894f0c                  mov dword ptr [edi+0c], ecx
'006345a6    50                      push eax
'006345a7    ff5630                  call dword ptr [esi+30]
'006345aa    dbe2                    fnclex
'006345ac    33ff                    xor edi, edi
var_num7 = Empty
'006345ae    3bc7                    cmp eax, edi
'006345b0    7d11                    jge 6345c3

If (0 < -744 - 16) Then
'006345b2    6a30                    push 30
'006345b4    68d8304300              push 004330d8
'006345b9    8b955cfeffff            mov edx, dword ptr [ebp+fffffe5c]
'006345bf    52                      push edx
'006345c0    50                      push eax
'006345c1    ffd3                    call ebx
    
End If
'006345c3    8b458c                  mov eax, dword ptr [ebp-74]
'006345c6    897d8c                  mov dword ptr [ebp-74], edi
'006345c9    898560ffffff            mov dword ptr [ebp+ffffff60], eax
'006345cf    c78558ffffff09000000    mov dword ptr [ebp+ffffff58], 00000009
'006345d9    89bd70ffffff            mov dword ptr [ebp+ffffff70], edi
'006345df    c78568ffffff02000000    mov dword ptr [ebp+ffffff68], 00000002
'006345e9    668b8574feffff          mov ax, word ptr [ebp+fffffe74]
'006345f0    668985d0feffff          mov word ptr [ebp+fffffed0], ax
'006345f7    c785c8feffff0b000000    mov dword ptr [ebp+fffffec8], 0000000b
'00634601    8d8d58ffffff            lea ecx, dword ptr [ebp+ffffff58]
'00634607    51                      push ecx
'00634608    8d9568ffffff            lea edx, dword ptr [ebp+ffffff68]
'0063460e    52                      push edx
'0063460f    8d85c8feffff            lea eax, dword ptr [ebp+fffffec8]
'00634615    50                      push eax
'00634616    8d8d48ffffff            lea ecx, dword ptr [ebp+ffffff48]
'0063461c    51                      push ecx

' *** Reference to "rtcImmediateIf"
'0063461d    ff154c124000            call dword ptr [0040124c]
'00634623    8d9548ffffff            lea edx, dword ptr [ebp+ffffff48]
'00634629    52                      push edx

' *** Reference to "__vbaI2Var"
'0063462a    ff150c124000            call dword ptr [0040120c]
'00634630    8bf0                    mov esi, eax
'00634632    8d4590                  lea eax, dword ptr [ebp-70]
'00634635    50                      push eax
'00634636    8d4d98                  lea ecx, dword ptr [ebp-68]
'00634639    51                      push ecx
'0063463a    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'0063463c    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_130, var_8)
'00634642    8d9548ffffff            lea edx, dword ptr [ebp+ffffff48]
'00634648    52                      push edx
'00634649    8d8558ffffff            lea eax, dword ptr [ebp+ffffff58]
'0063464f    50                      push eax
'00634650    8d8d68ffffff            lea ecx, dword ptr [ebp+ffffff68]
'00634656    51                      push ecx
'00634657    8d95c8feffff            lea edx, dword ptr [ebp+fffffec8]
'0063465d    52                      push edx
'0063465e    8d8578ffffff            lea eax, dword ptr [ebp+ffffff78]
'00634664    50                      push eax
'00634665    6a05                    push 05

' *** Reference to "__vbaFreeVarList"
'00634667    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_87, var_65, var_132, var_133, var_134)
'0063466d    83c424                  add esp, 24
'00634670    8b45c0                  mov eax, dword ptr [ebp-40]
'00634673    8b08                    mov ecx, dword ptr [eax]
'00634675    8d5598                  lea edx, dword ptr [ebp-68]
'00634678    52                      push edx
'00634679    50                      push eax
'0063467a    ff91b4000000            call dword ptr [ecx+000000b4]
'00634680    dbe2                    fnclex
'00634682    3bc7                    cmp eax, edi
'00634684    7d11                    jge 634697
'00634686    68b4000000              push 000000b4
'0063468b    6830314300              push 00433130
'00634690    8b4dc0                  mov ecx, dword ptr [ebp-40]
'00634693    51                      push ecx
'00634694    50                      push eax
'00634695    ffd3                    call ebx
'00634697    8b4598                  mov eax, dword ptr [ebp-68]
'0063469a    898568feffff            mov dword ptr [ebp+fffffe68], eax
'006346a0    c785f0feffff84d44300    mov dword ptr [ebp+fffffef0], 0043d484
'006346aa    b908000000              mov ecx, 00000008
'006346af    898de8feffff            mov dword ptr [ebp+fffffee8], ecx
'006346b5    8b10                    mov edx, dword ptr [eax]
'006346b7    8d7d94                  lea edi, dword ptr [ebp-6c]
'006346ba    57                      push edi
'006346bb    83ec10                  sub esp, 10
'006346be    8bfc                    mov edi, esp
'006346c0    890f                    mov dword ptr [edi], ecx
'006346c2    8b8decfeffff            mov ecx, dword ptr [ebp+fffffeec]
'006346c8    894f04                  mov dword ptr [edi+04], ecx
'006346cb    8b8df0feffff            mov ecx, dword ptr [ebp+fffffef0]
'006346d1    894f08                  mov dword ptr [edi+08], ecx
'006346d4    8b8df4feffff            mov ecx, dword ptr [ebp+fffffef4]
'006346da    894f0c                  mov dword ptr [edi+0c], ecx
'006346dd    50                      push eax
'006346de    ff5230                  call dword ptr [edx+30]
'006346e1    dbe2                    fnclex
'006346e3    85c0                    test ax, ax
'006346e5    7d11                    jge 6346f8
'006346e7    6a30                    push 30
'006346e9    68d8304300              push 004330d8
'006346ee    8b9568feffff            mov edx, dword ptr [ebp+fffffe68]
'006346f4    52                      push edx
'006346f5    50                      push eax
'006346f6    ffd3                    call ebx
'006346f8    8b4594                  mov eax, dword ptr [ebp-6c]
'006346fb    c7459400000000          mov dword ptr [ebp-6c], 00000000
'00634702    894580                  mov dword ptr [ebp-80], eax
'00634705    c78578ffffff09000000    mov dword ptr [ebp+ffffff78], 00000009
'0063470f    8d8578ffffff            lea eax, dword ptr [ebp+ffffff78]
'00634715    50                      push eax

' *** Reference to "rtcIsNull"
'00634716    ff1540114000            call dword ptr [00401140]
'0063471c    898574feffff            mov dword ptr [ebp+fffffe74], eax
'00634722    8b45c0                  mov eax, dword ptr [ebp-40]
'00634725    8b08                    mov ecx, dword ptr [eax]
'00634727    8d5590                  lea edx, dword ptr [ebp-70]
'0063472a    52                      push edx
'0063472b    50                      push eax
'0063472c    ff91b4000000            call dword ptr [ecx+000000b4]
'00634732    dbe2                    fnclex
'00634734    85c0                    test ax, ax
'00634736    7d11                    jge 634749
'00634738    68b4000000              push 000000b4
'0063473d    6830314300              push 00433130
'00634742    8b4dc0                  mov ecx, dword ptr [ebp-40]
'00634745    51                      push ecx
'00634746    50                      push eax
'00634747    ffd3                    call ebx
'00634749    8b4590                  mov eax, dword ptr [ebp-70]
'0063474c    89855cfeffff            mov dword ptr [ebp+fffffe5c], eax
'00634752    c785e0feffff84d44300    mov dword ptr [ebp+fffffee0], 0043d484
'0063475c    b908000000              mov ecx, 00000008
'00634761    898dd8feffff            mov dword ptr [ebp+fffffed8], ecx
'00634767    8b10                    mov edx, dword ptr [eax]
'00634769    8d7d8c                  lea edi, dword ptr [ebp-74]
'0063476c    57                      push edi
'0063476d    83ec10                  sub esp, 10
'00634770    8bfc                    mov edi, esp
'00634772    890f                    mov dword ptr [edi], ecx
'00634774    8b8ddcfeffff            mov ecx, dword ptr [ebp+fffffedc]
'0063477a    894f04                  mov dword ptr [edi+04], ecx
'0063477d    8b8de0feffff            mov ecx, dword ptr [ebp+fffffee0]
'00634783    894f08                  mov dword ptr [edi+08], ecx
'00634786    8b8de4feffff            mov ecx, dword ptr [ebp+fffffee4]
'0063478c    894f0c                  mov dword ptr [edi+0c], ecx
'0063478f    50                      push eax
'00634790    ff5230                  call dword ptr [edx+30]
'00634793    dbe2                    fnclex
'00634795    33ff                    xor edi, edi
var_num7 = Empty
'00634797    3bc7                    cmp eax, edi
'00634799    7d11                    jge 6347ac

If (0 < -756 - 16) Then
'0063479b    6a30                    push 30
'0063479d    68d8304300              push 004330d8
'006347a2    8b955cfeffff            mov edx, dword ptr [ebp+fffffe5c]
'006347a8    52                      push edx
'006347a9    50                      push eax
'006347aa    ffd3                    call ebx
    
End If
'006347ac    8b458c                  mov eax, dword ptr [ebp-74]
'006347af    897d8c                  mov dword ptr [ebp-74], edi
'006347b2    898560ffffff            mov dword ptr [ebp+ffffff60], eax
'006347b8    c78558ffffff09000000    mov dword ptr [ebp+ffffff58], 00000009
'006347c2    89bd70ffffff            mov dword ptr [ebp+ffffff70], edi
'006347c8    c78568ffffff02000000    mov dword ptr [ebp+ffffff68], 00000002
'006347d2    668b8574feffff          mov ax, word ptr [ebp+fffffe74]
'006347d9    668985d0feffff          mov word ptr [ebp+fffffed0], ax
'006347e0    c785c8feffff0b000000    mov dword ptr [ebp+fffffec8], 0000000b
'006347ea    8d8d58ffffff            lea ecx, dword ptr [ebp+ffffff58]
'006347f0    51                      push ecx
'006347f1    8d9568ffffff            lea edx, dword ptr [ebp+ffffff68]
'006347f7    52                      push edx
'006347f8    8d85c8feffff            lea eax, dword ptr [ebp+fffffec8]
'006347fe    50                      push eax
'006347ff    8d8d48ffffff            lea ecx, dword ptr [ebp+ffffff48]
'00634805    51                      push ecx

' *** Reference to "rtcImmediateIf"
'00634806    ff154c124000            call dword ptr [0040124c]
'0063480c    8d9548ffffff            lea edx, dword ptr [ebp+ffffff48]
'00634812    52                      push edx

' *** Reference to "__vbaI2Var"
'00634813    ff150c124000            call dword ptr [0040120c]
'00634819    8945e0                  mov dword ptr [ebp-20], eax
'0063481c    8d4590                  lea eax, dword ptr [ebp-70]
'0063481f    50                      push eax
'00634820    8d4d98                  lea ecx, dword ptr [ebp-68]
'00634823    51                      push ecx
'00634824    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'00634826    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_130, var_8)
'0063482c    8d9548ffffff            lea edx, dword ptr [ebp+ffffff48]
'00634832    52                      push edx
'00634833    8d8558ffffff            lea eax, dword ptr [ebp+ffffff58]
'00634839    50                      push eax
'0063483a    8d8d68ffffff            lea ecx, dword ptr [ebp+ffffff68]
'00634840    51                      push ecx
'00634841    8d95c8feffff            lea edx, dword ptr [ebp+fffffec8]
'00634847    52                      push edx
'00634848    8d8578ffffff            lea eax, dword ptr [ebp+ffffff78]
'0063484e    50                      push eax
'0063484f    6a05                    push 05

' *** Reference to "__vbaFreeVarList"
'00634851    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_87, var_65, var_132, var_133, var_134)
'00634857    83c424                  add esp, 24
'0063485a    8b45c0                  mov eax, dword ptr [ebp-40]
'0063485d    8b08                    mov ecx, dword ptr [eax]
'0063485f    8d5598                  lea edx, dword ptr [ebp-68]
'00634862    52                      push edx
'00634863    50                      push eax
'00634864    ff91b4000000            call dword ptr [ecx+000000b4]
'0063486a    dbe2                    fnclex
'0063486c    3bc7                    cmp eax, edi
'0063486e    7d11                    jge 634881
'00634870    68b4000000              push 000000b4
'00634875    6830314300              push 00433130
'0063487a    8b4dc0                  mov ecx, dword ptr [ebp-40]
'0063487d    51                      push ecx
'0063487e    50                      push eax
'0063487f    ffd3                    call ebx
'00634881    8b4598                  mov eax, dword ptr [ebp-68]
'00634884    898568feffff            mov dword ptr [ebp+fffffe68], eax
'0063488a    c785f0feffff08ab4300    mov dword ptr [ebp+fffffef0], 0043ab08
'00634894    b908000000              mov ecx, 00000008
'00634899    898de8feffff            mov dword ptr [ebp+fffffee8], ecx
'0063489f    8b10                    mov edx, dword ptr [eax]
'006348a1    8d7d94                  lea edi, dword ptr [ebp-6c]
'006348a4    57                      push edi
'006348a5    83ec10                  sub esp, 10
'006348a8    8bfc                    mov edi, esp
'006348aa    890f                    mov dword ptr [edi], ecx
'006348ac    8b8decfeffff            mov ecx, dword ptr [ebp+fffffeec]
'006348b2    894f04                  mov dword ptr [edi+04], ecx
'006348b5    8b8df0feffff            mov ecx, dword ptr [ebp+fffffef0]
'006348bb    894f08                  mov dword ptr [edi+08], ecx
'006348be    8b8df4feffff            mov ecx, dword ptr [ebp+fffffef4]
'006348c4    894f0c                  mov dword ptr [edi+0c], ecx
'006348c7    50                      push eax
'006348c8    ff5230                  call dword ptr [edx+30]
'006348cb    dbe2                    fnclex
'006348cd    85c0                    test ax, ax
'006348cf    7d11                    jge 6348e2
'006348d1    6a30                    push 30
'006348d3    68d8304300              push 004330d8
'006348d8    8b9568feffff            mov edx, dword ptr [ebp+fffffe68]
'006348de    52                      push edx
'006348df    50                      push eax
'006348e0    ffd3                    call ebx
'006348e2    8b4594                  mov eax, dword ptr [ebp-6c]
'006348e5    c7459400000000          mov dword ptr [ebp-6c], 00000000
'006348ec    894580                  mov dword ptr [ebp-80], eax
'006348ef    c78578ffffff09000000    mov dword ptr [ebp+ffffff78], 00000009
'006348f9    8d8578ffffff            lea eax, dword ptr [ebp+ffffff78]
'006348ff    50                      push eax

' *** Reference to "rtcIsNull"
'00634900    ff1540114000            call dword ptr [00401140]
'00634906    898574feffff            mov dword ptr [ebp+fffffe74], eax
'0063490c    8b45c0                  mov eax, dword ptr [ebp-40]
'0063490f    8b08                    mov ecx, dword ptr [eax]
'00634911    8d5590                  lea edx, dword ptr [ebp-70]
'00634914    52                      push edx
'00634915    50                      push eax
'00634916    ff91b4000000            call dword ptr [ecx+000000b4]
'0063491c    dbe2                    fnclex
'0063491e    85c0                    test ax, ax
'00634920    7d11                    jge 634933
'00634922    68b4000000              push 000000b4
'00634927    6830314300              push 00433130
'0063492c    8b4dc0                  mov ecx, dword ptr [ebp-40]
'0063492f    51                      push ecx
'00634930    50                      push eax
'00634931    ffd3                    call ebx
'00634933    8b4590                  mov eax, dword ptr [ebp-70]
'00634936    89855cfeffff            mov dword ptr [ebp+fffffe5c], eax
'0063493c    c785e0feffff08ab4300    mov dword ptr [ebp+fffffee0], 0043ab08
'00634946    b908000000              mov ecx, 00000008
'0063494b    898dd8feffff            mov dword ptr [ebp+fffffed8], ecx
'00634951    8b10                    mov edx, dword ptr [eax]
'00634953    8d7d8c                  lea edi, dword ptr [ebp-74]
'00634956    57                      push edi
'00634957    83ec10                  sub esp, 10
'0063495a    8bfc                    mov edi, esp
'0063495c    890f                    mov dword ptr [edi], ecx
'0063495e    8b8ddcfeffff            mov ecx, dword ptr [ebp+fffffedc]
'00634964    894f04                  mov dword ptr [edi+04], ecx
'00634967    8b8de0feffff            mov ecx, dword ptr [ebp+fffffee0]
'0063496d    894f08                  mov dword ptr [edi+08], ecx
'00634970    8b8de4feffff            mov ecx, dword ptr [ebp+fffffee4]
'00634976    894f0c                  mov dword ptr [edi+0c], ecx
'00634979    50                      push eax
'0063497a    ff5230                  call dword ptr [edx+30]
'0063497d    dbe2                    fnclex
'0063497f    33ff                    xor edi, edi
var_num7 = Empty
'00634981    3bc7                    cmp eax, edi
'00634983    7d11                    jge 634996

If (0 < -768 - 16) Then
'00634985    6a30                    push 30
'00634987    68d8304300              push 004330d8
'0063498c    8b955cfeffff            mov edx, dword ptr [ebp+fffffe5c]
'00634992    52                      push edx
'00634993    50                      push eax
'00634994    ffd3                    call ebx
    
End If
'00634996    8b458c                  mov eax, dword ptr [ebp-74]
'00634999    897d8c                  mov dword ptr [ebp-74], edi
'0063499c    898560ffffff            mov dword ptr [ebp+ffffff60], eax
'006349a2    c78558ffffff09000000    mov dword ptr [ebp+ffffff58], 00000009
'006349ac    89bd70ffffff            mov dword ptr [ebp+ffffff70], edi
'006349b2    c78568ffffff02000000    mov dword ptr [ebp+ffffff68], 00000002
'006349bc    668b8574feffff          mov ax, word ptr [ebp+fffffe74]
'006349c3    668985d0feffff          mov word ptr [ebp+fffffed0], ax
'006349ca    c785c8feffff0b000000    mov dword ptr [ebp+fffffec8], 0000000b
'006349d4    8d8d58ffffff            lea ecx, dword ptr [ebp+ffffff58]
'006349da    51                      push ecx
'006349db    8d9568ffffff            lea edx, dword ptr [ebp+ffffff68]
'006349e1    52                      push edx
'006349e2    8d85c8feffff            lea eax, dword ptr [ebp+fffffec8]
'006349e8    50                      push eax
'006349e9    8d8d48ffffff            lea ecx, dword ptr [ebp+ffffff48]
'006349ef    51                      push ecx

' *** Reference to "rtcImmediateIf"
'006349f0    ff154c124000            call dword ptr [0040124c]
'006349f6    8d9548ffffff            lea edx, dword ptr [ebp+ffffff48]
'006349fc    52                      push edx

' *** Reference to "__vbaI2Var"
'006349fd    ff150c124000            call dword ptr [0040120c]
'00634a03    8945d8                  mov dword ptr [ebp-28], eax
'00634a06    8d4590                  lea eax, dword ptr [ebp-70]
'00634a09    50                      push eax
'00634a0a    8d4d98                  lea ecx, dword ptr [ebp-68]
'00634a0d    51                      push ecx
'00634a0e    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'00634a10    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_130, var_8)
'00634a16    8d9548ffffff            lea edx, dword ptr [ebp+ffffff48]
'00634a1c    52                      push edx
'00634a1d    8d8558ffffff            lea eax, dword ptr [ebp+ffffff58]
'00634a23    50                      push eax
'00634a24    8d8d68ffffff            lea ecx, dword ptr [ebp+ffffff68]
'00634a2a    51                      push ecx
'00634a2b    8d95c8feffff            lea edx, dword ptr [ebp+fffffec8]
'00634a31    52                      push edx
'00634a32    8d8578ffffff            lea eax, dword ptr [ebp+ffffff78]
'00634a38    50                      push eax
'00634a39    6a05                    push 05

' *** Reference to "__vbaFreeVarList"
'00634a3b    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_87, var_65, var_132, var_133, var_134)
'00634a41    83c424                  add esp, 24
'00634a44    8b45c0                  mov eax, dword ptr [ebp-40]
'00634a47    8b08                    mov ecx, dword ptr [eax]
'00634a49    8d5598                  lea edx, dword ptr [ebp-68]
'00634a4c    52                      push edx
'00634a4d    50                      push eax
'00634a4e    ff91b4000000            call dword ptr [ecx+000000b4]
'00634a54    dbe2                    fnclex
'00634a56    3bc7                    cmp eax, edi
'00634a58    7d11                    jge 634a6b
'00634a5a    68b4000000              push 000000b4
'00634a5f    6830314300              push 00433130
'00634a64    8b4dc0                  mov ecx, dword ptr [ebp-40]
'00634a67    51                      push ecx
'00634a68    50                      push eax
'00634a69    ffd3                    call ebx
'00634a6b    8b4598                  mov eax, dword ptr [ebp-68]
'00634a6e    898568feffff            mov dword ptr [ebp+fffffe68], eax
'00634a74    c785f0feffff18ab4300    mov dword ptr [ebp+fffffef0], 0043ab18
'00634a7e    b908000000              mov ecx, 00000008
'00634a83    898de8feffff            mov dword ptr [ebp+fffffee8], ecx
'00634a89    8b10                    mov edx, dword ptr [eax]
'00634a8b    8d7d94                  lea edi, dword ptr [ebp-6c]
'00634a8e    57                      push edi
'00634a8f    83ec10                  sub esp, 10
'00634a92    8bfc                    mov edi, esp
'00634a94    890f                    mov dword ptr [edi], ecx
'00634a96    8b8decfeffff            mov ecx, dword ptr [ebp+fffffeec]
'00634a9c    894f04                  mov dword ptr [edi+04], ecx
'00634a9f    8b8df0feffff            mov ecx, dword ptr [ebp+fffffef0]
'00634aa5    894f08                  mov dword ptr [edi+08], ecx
'00634aa8    8b8df4feffff            mov ecx, dword ptr [ebp+fffffef4]
'00634aae    894f0c                  mov dword ptr [edi+0c], ecx
'00634ab1    50                      push eax
'00634ab2    ff5230                  call dword ptr [edx+30]
'00634ab5    dbe2                    fnclex
'00634ab7    85c0                    test ax, ax
'00634ab9    7d11                    jge 634acc
'00634abb    6a30                    push 30
'00634abd    68d8304300              push 004330d8
'00634ac2    8b9568feffff            mov edx, dword ptr [ebp+fffffe68]
'00634ac8    52                      push edx
'00634ac9    50                      push eax
'00634aca    ffd3                    call ebx
'00634acc    8b4594                  mov eax, dword ptr [ebp-6c]
'00634acf    c7459400000000          mov dword ptr [ebp-6c], 00000000
'00634ad6    894580                  mov dword ptr [ebp-80], eax
'00634ad9    c78578ffffff09000000    mov dword ptr [ebp+ffffff78], 00000009
'00634ae3    8d8578ffffff            lea eax, dword ptr [ebp+ffffff78]
'00634ae9    50                      push eax

' *** Reference to "rtcIsNull"
'00634aea    ff1540114000            call dword ptr [00401140]
'00634af0    898574feffff            mov dword ptr [ebp+fffffe74], eax
'00634af6    8b45c0                  mov eax, dword ptr [ebp-40]
'00634af9    8b08                    mov ecx, dword ptr [eax]
'00634afb    8d5590                  lea edx, dword ptr [ebp-70]
'00634afe    52                      push edx
'00634aff    50                      push eax
'00634b00    ff91b4000000            call dword ptr [ecx+000000b4]
'00634b06    dbe2                    fnclex
'00634b08    85c0                    test ax, ax
'00634b0a    7d11                    jge 634b1d
'00634b0c    68b4000000              push 000000b4
'00634b11    6830314300              push 00433130
'00634b16    8b4dc0                  mov ecx, dword ptr [ebp-40]
'00634b19    51                      push ecx
'00634b1a    50                      push eax
'00634b1b    ffd3                    call ebx
'00634b1d    8b4590                  mov eax, dword ptr [ebp-70]
'00634b20    89855cfeffff            mov dword ptr [ebp+fffffe5c], eax
'00634b26    c785e0feffff18ab4300    mov dword ptr [ebp+fffffee0], 0043ab18
'00634b30    b908000000              mov ecx, 00000008
'00634b35    898dd8feffff            mov dword ptr [ebp+fffffed8], ecx
'00634b3b    8b10                    mov edx, dword ptr [eax]
'00634b3d    8d7d8c                  lea edi, dword ptr [ebp-74]
'00634b40    57                      push edi
'00634b41    83ec10                  sub esp, 10
'00634b44    8bfc                    mov edi, esp
'00634b46    890f                    mov dword ptr [edi], ecx
'00634b48    8b8ddcfeffff            mov ecx, dword ptr [ebp+fffffedc]
'00634b4e    894f04                  mov dword ptr [edi+04], ecx
'00634b51    8b8de0feffff            mov ecx, dword ptr [ebp+fffffee0]
'00634b57    894f08                  mov dword ptr [edi+08], ecx
'00634b5a    8b8de4feffff            mov ecx, dword ptr [ebp+fffffee4]
'00634b60    894f0c                  mov dword ptr [edi+0c], ecx
'00634b63    50                      push eax
'00634b64    ff5230                  call dword ptr [edx+30]
'00634b67    dbe2                    fnclex
'00634b69    33ff                    xor edi, edi
var_num7 = Empty
'00634b6b    3bc7                    cmp eax, edi
'00634b6d    7d11                    jge 634b80

If (0 < -780 - 16) Then
'00634b6f    6a30                    push 30
'00634b71    68d8304300              push 004330d8
'00634b76    8b955cfeffff            mov edx, dword ptr [ebp+fffffe5c]
'00634b7c    52                      push edx
'00634b7d    50                      push eax
'00634b7e    ffd3                    call ebx
    
End If
'00634b80    8b458c                  mov eax, dword ptr [ebp-74]
'00634b83    897d8c                  mov dword ptr [ebp-74], edi
'00634b86    898560ffffff            mov dword ptr [ebp+ffffff60], eax
'00634b8c    c78558ffffff09000000    mov dword ptr [ebp+ffffff58], 00000009
'00634b96    89bd70ffffff            mov dword ptr [ebp+ffffff70], edi
'00634b9c    c78568ffffff02000000    mov dword ptr [ebp+ffffff68], 00000002
'00634ba6    668b8574feffff          mov ax, word ptr [ebp+fffffe74]
'00634bad    668985d0feffff          mov word ptr [ebp+fffffed0], ax
'00634bb4    c785c8feffff0b000000    mov dword ptr [ebp+fffffec8], 0000000b
'00634bbe    8d8d58ffffff            lea ecx, dword ptr [ebp+ffffff58]
'00634bc4    51                      push ecx
'00634bc5    8d9568ffffff            lea edx, dword ptr [ebp+ffffff68]
'00634bcb    52                      push edx
'00634bcc    8d85c8feffff            lea eax, dword ptr [ebp+fffffec8]
'00634bd2    50                      push eax
'00634bd3    8d8d48ffffff            lea ecx, dword ptr [ebp+ffffff48]
'00634bd9    51                      push ecx

' *** Reference to "rtcImmediateIf"
'00634bda    ff154c124000            call dword ptr [0040124c]
'00634be0    8d9548ffffff            lea edx, dword ptr [ebp+ffffff48]
'00634be6    52                      push edx

' *** Reference to "__vbaI2Var"
'00634be7    ff150c124000            call dword ptr [0040120c]
'00634bed    8945d0                  mov dword ptr [ebp-30], eax
'00634bf0    8d4590                  lea eax, dword ptr [ebp-70]
'00634bf3    50                      push eax
'00634bf4    8d4d98                  lea ecx, dword ptr [ebp-68]
'00634bf7    51                      push ecx
'00634bf8    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'00634bfa    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_130, var_8)
'00634c00    8d9548ffffff            lea edx, dword ptr [ebp+ffffff48]
'00634c06    52                      push edx
'00634c07    8d8558ffffff            lea eax, dword ptr [ebp+ffffff58]
'00634c0d    50                      push eax
'00634c0e    8d8d68ffffff            lea ecx, dword ptr [ebp+ffffff68]
'00634c14    51                      push ecx
'00634c15    8d95c8feffff            lea edx, dword ptr [ebp+fffffec8]
'00634c1b    52                      push edx
'00634c1c    8d8578ffffff            lea eax, dword ptr [ebp+ffffff78]
'00634c22    50                      push eax
'00634c23    6a05                    push 05

' *** Reference to "__vbaFreeVarList"
'00634c25    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_87, var_65, var_132, var_133, var_134)
'00634c2b    83c424                  add esp, 24
'00634c2e    8b45c0                  mov eax, dword ptr [ebp-40]
'00634c31    8b08                    mov ecx, dword ptr [eax]
'00634c33    8d5598                  lea edx, dword ptr [ebp-68]
'00634c36    52                      push edx
'00634c37    50                      push eax
'00634c38    ff91b4000000            call dword ptr [ecx+000000b4]
'00634c3e    dbe2                    fnclex
'00634c40    3bc7                    cmp eax, edi
'00634c42    7d11                    jge 634c55
'00634c44    68b4000000              push 000000b4
'00634c49    6830314300              push 00433130
'00634c4e    8b4dc0                  mov ecx, dword ptr [ebp-40]
'00634c51    51                      push ecx
'00634c52    50                      push eax
'00634c53    ffd3                    call ebx
'00634c55    8b4598                  mov eax, dword ptr [ebp-68]
'00634c58    898568feffff            mov dword ptr [ebp+fffffe68], eax
'00634c5e    c785f0feffff28ab4300    mov dword ptr [ebp+fffffef0], 0043ab28
'00634c68    b908000000              mov ecx, 00000008
'00634c6d    898de8feffff            mov dword ptr [ebp+fffffee8], ecx
'00634c73    8b10                    mov edx, dword ptr [eax]
'00634c75    8d7d94                  lea edi, dword ptr [ebp-6c]
'00634c78    57                      push edi
'00634c79    83ec10                  sub esp, 10
'00634c7c    8bfc                    mov edi, esp
'00634c7e    890f                    mov dword ptr [edi], ecx
'00634c80    8b8decfeffff            mov ecx, dword ptr [ebp+fffffeec]
'00634c86    894f04                  mov dword ptr [edi+04], ecx
'00634c89    8b8df0feffff            mov ecx, dword ptr [ebp+fffffef0]
'00634c8f    894f08                  mov dword ptr [edi+08], ecx
'00634c92    8b8df4feffff            mov ecx, dword ptr [ebp+fffffef4]
'00634c98    894f0c                  mov dword ptr [edi+0c], ecx
'00634c9b    50                      push eax
'00634c9c    ff5230                  call dword ptr [edx+30]
'00634c9f    dbe2                    fnclex
'00634ca1    85c0                    test ax, ax
'00634ca3    7d11                    jge 634cb6
'00634ca5    6a30                    push 30
'00634ca7    68d8304300              push 004330d8
'00634cac    8b9568feffff            mov edx, dword ptr [ebp+fffffe68]
'00634cb2    52                      push edx
'00634cb3    50                      push eax
'00634cb4    ffd3                    call ebx
'00634cb6    8b4594                  mov eax, dword ptr [ebp-6c]
'00634cb9    c7459400000000          mov dword ptr [ebp-6c], 00000000
'00634cc0    894580                  mov dword ptr [ebp-80], eax
'00634cc3    c78578ffffff09000000    mov dword ptr [ebp+ffffff78], 00000009
'00634ccd    8d8578ffffff            lea eax, dword ptr [ebp+ffffff78]
'00634cd3    50                      push eax

' *** Reference to "rtcIsNull"
'00634cd4    ff1540114000            call dword ptr [00401140]
'00634cda    898574feffff            mov dword ptr [ebp+fffffe74], eax
'00634ce0    8b45c0                  mov eax, dword ptr [ebp-40]
'00634ce3    8b08                    mov ecx, dword ptr [eax]
'00634ce5    8d5590                  lea edx, dword ptr [ebp-70]
'00634ce8    52                      push edx
'00634ce9    50                      push eax
'00634cea    ff91b4000000            call dword ptr [ecx+000000b4]
'00634cf0    dbe2                    fnclex
'00634cf2    85c0                    test ax, ax
'00634cf4    7d11                    jge 634d07
'00634cf6    68b4000000              push 000000b4
'00634cfb    6830314300              push 00433130
'00634d00    8b4dc0                  mov ecx, dword ptr [ebp-40]
'00634d03    51                      push ecx
'00634d04    50                      push eax
'00634d05    ffd3                    call ebx
'00634d07    8b4590                  mov eax, dword ptr [ebp-70]
'00634d0a    89855cfeffff            mov dword ptr [ebp+fffffe5c], eax
'00634d10    c785e0feffff28ab4300    mov dword ptr [ebp+fffffee0], 0043ab28
'00634d1a    b908000000              mov ecx, 00000008
'00634d1f    898dd8feffff            mov dword ptr [ebp+fffffed8], ecx
'00634d25    8b10                    mov edx, dword ptr [eax]
'00634d27    8d7d8c                  lea edi, dword ptr [ebp-74]
'00634d2a    57                      push edi
'00634d2b    83ec10                  sub esp, 10
'00634d2e    8bfc                    mov edi, esp
'00634d30    890f                    mov dword ptr [edi], ecx
'00634d32    8b8ddcfeffff            mov ecx, dword ptr [ebp+fffffedc]
'00634d38    894f04                  mov dword ptr [edi+04], ecx
'00634d3b    8b8de0feffff            mov ecx, dword ptr [ebp+fffffee0]
'00634d41    894f08                  mov dword ptr [edi+08], ecx
'00634d44    8b8de4feffff            mov ecx, dword ptr [ebp+fffffee4]
'00634d4a    894f0c                  mov dword ptr [edi+0c], ecx
'00634d4d    50                      push eax
'00634d4e    ff5230                  call dword ptr [edx+30]
'00634d51    dbe2                    fnclex
'00634d53    33ff                    xor edi, edi
var_num7 = Empty
'00634d55    3bc7                    cmp eax, edi
'00634d57    7d11                    jge 634d6a

If (0 < -792 - 16) Then
'00634d59    6a30                    push 30
'00634d5b    68d8304300              push 004330d8
'00634d60    8b955cfeffff            mov edx, dword ptr [ebp+fffffe5c]
'00634d66    52                      push edx
'00634d67    50                      push eax
'00634d68    ffd3                    call ebx
    
End If
'00634d6a    8b458c                  mov eax, dword ptr [ebp-74]
'00634d6d    897d8c                  mov dword ptr [ebp-74], edi
'00634d70    898560ffffff            mov dword ptr [ebp+ffffff60], eax
'00634d76    c78558ffffff09000000    mov dword ptr [ebp+ffffff58], 00000009
'00634d80    89bd70ffffff            mov dword ptr [ebp+ffffff70], edi
'00634d86    c78568ffffff02000000    mov dword ptr [ebp+ffffff68], 00000002
'00634d90    668b8574feffff          mov ax, word ptr [ebp+fffffe74]
'00634d97    668985d0feffff          mov word ptr [ebp+fffffed0], ax
'00634d9e    c785c8feffff0b000000    mov dword ptr [ebp+fffffec8], 0000000b
'00634da8    8d8d58ffffff            lea ecx, dword ptr [ebp+ffffff58]
'00634dae    51                      push ecx
'00634daf    8d9568ffffff            lea edx, dword ptr [ebp+ffffff68]
'00634db5    52                      push edx
'00634db6    8d85c8feffff            lea eax, dword ptr [ebp+fffffec8]
'00634dbc    50                      push eax
'00634dbd    8d8d48ffffff            lea ecx, dword ptr [ebp+ffffff48]
'00634dc3    51                      push ecx

' *** Reference to "rtcImmediateIf"
'00634dc4    ff154c124000            call dword ptr [0040124c]
'00634dca    8d9548ffffff            lea edx, dword ptr [ebp+ffffff48]
'00634dd0    52                      push edx

' *** Reference to "__vbaI2Var"
'00634dd1    ff150c124000            call dword ptr [0040120c]
'00634dd7    8945cc                  mov dword ptr [ebp-34], eax
'00634dda    8d4590                  lea eax, dword ptr [ebp-70]
'00634ddd    50                      push eax
'00634dde    8d4d98                  lea ecx, dword ptr [ebp-68]
'00634de1    51                      push ecx
'00634de2    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'00634de4    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_130, var_8)
'00634dea    8d9548ffffff            lea edx, dword ptr [ebp+ffffff48]
'00634df0    52                      push edx
'00634df1    8d8558ffffff            lea eax, dword ptr [ebp+ffffff58]
'00634df7    50                      push eax
'00634df8    8d8d68ffffff            lea ecx, dword ptr [ebp+ffffff68]
'00634dfe    51                      push ecx
'00634dff    8d95c8feffff            lea edx, dword ptr [ebp+fffffec8]
'00634e05    52                      push edx
'00634e06    8d8578ffffff            lea eax, dword ptr [ebp+ffffff78]
'00634e0c    50                      push eax
'00634e0d    6a05                    push 05

' *** Reference to "__vbaFreeVarList"
'00634e0f    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_87, var_65, var_132, var_133, var_134)
'00634e15    83c424                  add esp, 24
'00634e18    8b45c0                  mov eax, dword ptr [ebp-40]
'00634e1b    8b08                    mov ecx, dword ptr [eax]
'00634e1d    8d5598                  lea edx, dword ptr [ebp-68]
'00634e20    52                      push edx
'00634e21    50                      push eax
'00634e22    ff91b4000000            call dword ptr [ecx+000000b4]
'00634e28    dbe2                    fnclex
'00634e2a    3bc7                    cmp eax, edi
'00634e2c    7d11                    jge 634e3f
'00634e2e    68b4000000              push 000000b4
'00634e33    6830314300              push 00433130
'00634e38    8b4dc0                  mov ecx, dword ptr [ebp-40]
'00634e3b    51                      push ecx
'00634e3c    50                      push eax
'00634e3d    ffd3                    call ebx
'00634e3f    8b4598                  mov eax, dword ptr [ebp-68]
'00634e42    898568feffff            mov dword ptr [ebp+fffffe68], eax
'00634e48    c785f0feffffac874300    mov dword ptr [ebp+fffffef0], 004387ac
'00634e52    b908000000              mov ecx, 00000008
'00634e57    898de8feffff            mov dword ptr [ebp+fffffee8], ecx
'00634e5d    8b10                    mov edx, dword ptr [eax]
'00634e5f    8d7d94                  lea edi, dword ptr [ebp-6c]
'00634e62    57                      push edi
'00634e63    83ec10                  sub esp, 10
'00634e66    8bfc                    mov edi, esp
'00634e68    890f                    mov dword ptr [edi], ecx
'00634e6a    8b8decfeffff            mov ecx, dword ptr [ebp+fffffeec]
'00634e70    894f04                  mov dword ptr [edi+04], ecx
'00634e73    8b8df0feffff            mov ecx, dword ptr [ebp+fffffef0]
'00634e79    894f08                  mov dword ptr [edi+08], ecx
'00634e7c    8b8df4feffff            mov ecx, dword ptr [ebp+fffffef4]
'00634e82    894f0c                  mov dword ptr [edi+0c], ecx
'00634e85    50                      push eax
'00634e86    ff5230                  call dword ptr [edx+30]
'00634e89    dbe2                    fnclex
'00634e8b    85c0                    test ax, ax
'00634e8d    7d11                    jge 634ea0
'00634e8f    6a30                    push 30
'00634e91    68d8304300              push 004330d8
'00634e96    8b9568feffff            mov edx, dword ptr [ebp+fffffe68]
'00634e9c    52                      push edx
'00634e9d    50                      push eax
'00634e9e    ffd3                    call ebx
'00634ea0    8b4594                  mov eax, dword ptr [ebp-6c]
'00634ea3    c7459400000000          mov dword ptr [ebp-6c], 00000000
'00634eaa    894580                  mov dword ptr [ebp-80], eax
'00634ead    c78578ffffff09000000    mov dword ptr [ebp+ffffff78], 00000009
'00634eb7    8d8578ffffff            lea eax, dword ptr [ebp+ffffff78]
'00634ebd    50                      push eax

' *** Reference to "rtcIsNull"
'00634ebe    ff1540114000            call dword ptr [00401140]
'00634ec4    898574feffff            mov dword ptr [ebp+fffffe74], eax
'00634eca    8b45c0                  mov eax, dword ptr [ebp-40]
'00634ecd    8b08                    mov ecx, dword ptr [eax]
'00634ecf    8d5590                  lea edx, dword ptr [ebp-70]
'00634ed2    52                      push edx
'00634ed3    50                      push eax
'00634ed4    ff91b4000000            call dword ptr [ecx+000000b4]
'00634eda    dbe2                    fnclex
'00634edc    85c0                    test ax, ax
'00634ede    7d11                    jge 634ef1
'00634ee0    68b4000000              push 000000b4
'00634ee5    6830314300              push 00433130
'00634eea    8b4dc0                  mov ecx, dword ptr [ebp-40]
'00634eed    51                      push ecx
'00634eee    50                      push eax
'00634eef    ffd3                    call ebx
'00634ef1    8b4590                  mov eax, dword ptr [ebp-70]
'00634ef4    89855cfeffff            mov dword ptr [ebp+fffffe5c], eax
'00634efa    c785e0feffffac874300    mov dword ptr [ebp+fffffee0], 004387ac
'00634f04    b908000000              mov ecx, 00000008
'00634f09    898dd8feffff            mov dword ptr [ebp+fffffed8], ecx
'00634f0f    8b10                    mov edx, dword ptr [eax]
'00634f11    8d7d8c                  lea edi, dword ptr [ebp-74]
'00634f14    57                      push edi
'00634f15    83ec10                  sub esp, 10
'00634f18    8bfc                    mov edi, esp
'00634f1a    890f                    mov dword ptr [edi], ecx
'00634f1c    8b8ddcfeffff            mov ecx, dword ptr [ebp+fffffedc]
'00634f22    894f04                  mov dword ptr [edi+04], ecx
'00634f25    8b8de0feffff            mov ecx, dword ptr [ebp+fffffee0]
'00634f2b    894f08                  mov dword ptr [edi+08], ecx
'00634f2e    8b8de4feffff            mov ecx, dword ptr [ebp+fffffee4]
'00634f34    894f0c                  mov dword ptr [edi+0c], ecx
'00634f37    50                      push eax
'00634f38    ff5230                  call dword ptr [edx+30]
'00634f3b    dbe2                    fnclex
'00634f3d    33ff                    xor edi, edi
var_num7 = Empty
'00634f3f    3bc7                    cmp eax, edi
'00634f41    7d11                    jge 634f54

If (0 < -804 - 16) Then
'00634f43    6a30                    push 30
'00634f45    68d8304300              push 004330d8
'00634f4a    8b955cfeffff            mov edx, dword ptr [ebp+fffffe5c]
'00634f50    52                      push edx
'00634f51    50                      push eax
'00634f52    ffd3                    call ebx
    
End If
'00634f54    8b458c                  mov eax, dword ptr [ebp-74]
'00634f57    897d8c                  mov dword ptr [ebp-74], edi
'00634f5a    898560ffffff            mov dword ptr [ebp+ffffff60], eax
'00634f60    c78558ffffff09000000    mov dword ptr [ebp+ffffff58], 00000009
'00634f6a    89bd70ffffff            mov dword ptr [ebp+ffffff70], edi
'00634f70    c78568ffffff02000000    mov dword ptr [ebp+ffffff68], 00000002
'00634f7a    668b8574feffff          mov ax, word ptr [ebp+fffffe74]
'00634f81    668985d0feffff          mov word ptr [ebp+fffffed0], ax
'00634f88    c785c8feffff0b000000    mov dword ptr [ebp+fffffec8], 0000000b
'00634f92    8d8d58ffffff            lea ecx, dword ptr [ebp+ffffff58]
'00634f98    51                      push ecx
'00634f99    8d9568ffffff            lea edx, dword ptr [ebp+ffffff68]
'00634f9f    52                      push edx
'00634fa0    8d85c8feffff            lea eax, dword ptr [ebp+fffffec8]
'00634fa6    50                      push eax
'00634fa7    8d8d48ffffff            lea ecx, dword ptr [ebp+ffffff48]
'00634fad    51                      push ecx

' *** Reference to "rtcImmediateIf"
'00634fae    ff154c124000            call dword ptr [0040124c]
'00634fb4    8d9548ffffff            lea edx, dword ptr [ebp+ffffff48]
'00634fba    52                      push edx

' *** Reference to "__vbaI2Var"
'00634fbb    ff150c124000            call dword ptr [0040120c]
'00634fc1    8945c8                  mov dword ptr [ebp-38], eax
'00634fc4    8d4590                  lea eax, dword ptr [ebp-70]
'00634fc7    50                      push eax
'00634fc8    8d4d98                  lea ecx, dword ptr [ebp-68]
'00634fcb    51                      push ecx
'00634fcc    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'00634fce    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_130, var_8)
'00634fd4    8d9548ffffff            lea edx, dword ptr [ebp+ffffff48]
'00634fda    52                      push edx
'00634fdb    8d8558ffffff            lea eax, dword ptr [ebp+ffffff58]
'00634fe1    50                      push eax
'00634fe2    8d8d68ffffff            lea ecx, dword ptr [ebp+ffffff68]
'00634fe8    51                      push ecx
'00634fe9    8d95c8feffff            lea edx, dword ptr [ebp+fffffec8]
'00634fef    52                      push edx
'00634ff0    8d8578ffffff            lea eax, dword ptr [ebp+ffffff78]
'00634ff6    50                      push eax
'00634ff7    6a05                    push 05

' *** Reference to "__vbaFreeVarList"
'00634ff9    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_87, var_65, var_132, var_133, var_134)
'00634fff    83c424                  add esp, 24
'00635002    8b45c0                  mov eax, dword ptr [ebp-40]
'00635005    8b08                    mov ecx, dword ptr [eax]
'00635007    8d5598                  lea edx, dword ptr [ebp-68]
'0063500a    52                      push edx
'0063500b    50                      push eax
'0063500c    ff91b4000000            call dword ptr [ecx+000000b4]
'00635012    dbe2                    fnclex
'00635014    3bc7                    cmp eax, edi
'00635016    7d11                    jge 635029
'00635018    68b4000000              push 000000b4
'0063501d    6830314300              push 00433130
'00635022    8b4dc0                  mov ecx, dword ptr [ebp-40]
'00635025    51                      push ecx
'00635026    50                      push eax
'00635027    ffd3                    call ebx
'00635029    8b4598                  mov eax, dword ptr [ebp-68]
'0063502c    898568feffff            mov dword ptr [ebp+fffffe68], eax
'00635032    c785f0feffff38784300    mov dword ptr [ebp+fffffef0], 00437838
'0063503c    b908000000              mov ecx, 00000008
'00635041    898de8feffff            mov dword ptr [ebp+fffffee8], ecx
'00635047    8b10                    mov edx, dword ptr [eax]
'00635049    8d7d94                  lea edi, dword ptr [ebp-6c]
'0063504c    57                      push edi
'0063504d    83ec10                  sub esp, 10
'00635050    8bfc                    mov edi, esp
'00635052    890f                    mov dword ptr [edi], ecx
'00635054    8b8decfeffff            mov ecx, dword ptr [ebp+fffffeec]
'0063505a    894f04                  mov dword ptr [edi+04], ecx
'0063505d    8b8df0feffff            mov ecx, dword ptr [ebp+fffffef0]
'00635063    894f08                  mov dword ptr [edi+08], ecx
'00635066    8b8df4feffff            mov ecx, dword ptr [ebp+fffffef4]
'0063506c    894f0c                  mov dword ptr [edi+0c], ecx
'0063506f    50                      push eax
'00635070    ff5230                  call dword ptr [edx+30]
'00635073    dbe2                    fnclex
'00635075    85c0                    test ax, ax
'00635077    7d11                    jge 63508a
'00635079    6a30                    push 30
'0063507b    68d8304300              push 004330d8
'00635080    8b9568feffff            mov edx, dword ptr [ebp+fffffe68]
'00635086    52                      push edx
'00635087    50                      push eax
'00635088    ffd3                    call ebx
'0063508a    8b4594                  mov eax, dword ptr [ebp-6c]
'0063508d    c7459400000000          mov dword ptr [ebp-6c], 00000000
'00635094    894580                  mov dword ptr [ebp-80], eax
'00635097    c78578ffffff09000000    mov dword ptr [ebp+ffffff78], 00000009
'006350a1    8d8578ffffff            lea eax, dword ptr [ebp+ffffff78]
'006350a7    50                      push eax

' *** Reference to "rtcIsNull"
'006350a8    ff1540114000            call dword ptr [00401140]
'006350ae    898574feffff            mov dword ptr [ebp+fffffe74], eax
'006350b4    8b45c0                  mov eax, dword ptr [ebp-40]
'006350b7    8b08                    mov ecx, dword ptr [eax]
'006350b9    8d5590                  lea edx, dword ptr [ebp-70]
'006350bc    52                      push edx
'006350bd    50                      push eax
'006350be    ff91b4000000            call dword ptr [ecx+000000b4]
'006350c4    dbe2                    fnclex
'006350c6    85c0                    test ax, ax
'006350c8    7d11                    jge 6350db
'006350ca    68b4000000              push 000000b4
'006350cf    6830314300              push 00433130
'006350d4    8b4dc0                  mov ecx, dword ptr [ebp-40]
'006350d7    51                      push ecx
'006350d8    50                      push eax
'006350d9    ffd3                    call ebx
'006350db    8b4590                  mov eax, dword ptr [ebp-70]
'006350de    89855cfeffff            mov dword ptr [ebp+fffffe5c], eax
'006350e4    c785e0feffff38784300    mov dword ptr [ebp+fffffee0], 00437838
'006350ee    b908000000              mov ecx, 00000008
'006350f3    898dd8feffff            mov dword ptr [ebp+fffffed8], ecx
'006350f9    8b10                    mov edx, dword ptr [eax]
'006350fb    8d7d8c                  lea edi, dword ptr [ebp-74]
'006350fe    57                      push edi
'006350ff    83ec10                  sub esp, 10
'00635102    8bfc                    mov edi, esp
'00635104    890f                    mov dword ptr [edi], ecx
'00635106    8b8ddcfeffff            mov ecx, dword ptr [ebp+fffffedc]
'0063510c    894f04                  mov dword ptr [edi+04], ecx
'0063510f    8b8de0feffff            mov ecx, dword ptr [ebp+fffffee0]
'00635115    894f08                  mov dword ptr [edi+08], ecx
'00635118    8b8de4feffff            mov ecx, dword ptr [ebp+fffffee4]
'0063511e    894f0c                  mov dword ptr [edi+0c], ecx
'00635121    50                      push eax
'00635122    ff5230                  call dword ptr [edx+30]
'00635125    dbe2                    fnclex
'00635127    33ff                    xor edi, edi
var_num7 = Empty
'00635129    3bc7                    cmp eax, edi
'0063512b    7d11                    jge 63513e

If (0 < -816 - 16) Then
'0063512d    6a30                    push 30
'0063512f    68d8304300              push 004330d8
'00635134    8b955cfeffff            mov edx, dword ptr [ebp+fffffe5c]
'0063513a    52                      push edx
'0063513b    50                      push eax
'0063513c    ffd3                    call ebx
    
End If
'0063513e    8b458c                  mov eax, dword ptr [ebp-74]
'00635141    897d8c                  mov dword ptr [ebp-74], edi
'00635144    898560ffffff            mov dword ptr [ebp+ffffff60], eax
'0063514a    c78558ffffff09000000    mov dword ptr [ebp+ffffff58], 00000009
'00635154    89bd70ffffff            mov dword ptr [ebp+ffffff70], edi
'0063515a    c78568ffffff02000000    mov dword ptr [ebp+ffffff68], 00000002
'00635164    668b8574feffff          mov ax, word ptr [ebp+fffffe74]
'0063516b    668985d0feffff          mov word ptr [ebp+fffffed0], ax
'00635172    c785c8feffff0b000000    mov dword ptr [ebp+fffffec8], 0000000b
'0063517c    8d8d58ffffff            lea ecx, dword ptr [ebp+ffffff58]
'00635182    51                      push ecx
'00635183    8d9568ffffff            lea edx, dword ptr [ebp+ffffff68]
'00635189    52                      push edx
'0063518a    8d85c8feffff            lea eax, dword ptr [ebp+fffffec8]
'00635190    50                      push eax
'00635191    8d8d48ffffff            lea ecx, dword ptr [ebp+ffffff48]
'00635197    51                      push ecx

' *** Reference to "rtcImmediateIf"
'00635198    ff154c124000            call dword ptr [0040124c]
'0063519e    8d9548ffffff            lea edx, dword ptr [ebp+ffffff48]
'006351a4    52                      push edx

' *** Reference to "__vbaI2Var"
'006351a5    ff150c124000            call dword ptr [0040120c]
'006351ab    8945c4                  mov dword ptr [ebp-3c], eax
'006351ae    8d4590                  lea eax, dword ptr [ebp-70]
'006351b1    50                      push eax
'006351b2    8d4d98                  lea ecx, dword ptr [ebp-68]
'006351b5    51                      push ecx
'006351b6    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'006351b8    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_130, var_8)
'006351be    8d9548ffffff            lea edx, dword ptr [ebp+ffffff48]
'006351c4    52                      push edx
'006351c5    8d8558ffffff            lea eax, dword ptr [ebp+ffffff58]
'006351cb    50                      push eax
'006351cc    8d8d68ffffff            lea ecx, dword ptr [ebp+ffffff68]
'006351d2    51                      push ecx
'006351d3    8d95c8feffff            lea edx, dword ptr [ebp+fffffec8]
'006351d9    52                      push edx
'006351da    8d8578ffffff            lea eax, dword ptr [ebp+ffffff78]
'006351e0    50                      push eax
'006351e1    6a05                    push 05

' *** Reference to "__vbaFreeVarList"
'006351e3    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_87, var_65, var_132, var_133, var_134)
'006351e9    83c424                  add esp, 24
'006351ec    8b45c0                  mov eax, dword ptr [ebp-40]
'006351ef    8b08                    mov ecx, dword ptr [eax]
'006351f1    8d5598                  lea edx, dword ptr [ebp-68]
'006351f4    52                      push edx
'006351f5    50                      push eax
'006351f6    ff91b4000000            call dword ptr [ecx+000000b4]
'006351fc    dbe2                    fnclex
'006351fe    3bc7                    cmp eax, edi
'00635200    7d11                    jge 635213
'00635202    68b4000000              push 000000b4
'00635207    6830314300              push 00433130
'0063520c    8b4dc0                  mov ecx, dword ptr [ebp-40]
'0063520f    51                      push ecx
'00635210    50                      push eax
'00635211    ffd3                    call ebx
'00635213    8b4598                  mov eax, dword ptr [ebp-68]
'00635216    898568feffff            mov dword ptr [ebp+fffffe68], eax
'0063521c    c785f0feffff1cd44300    mov dword ptr [ebp+fffffef0], 0043d41c
'00635226    b908000000              mov ecx, 00000008
'0063522b    898de8feffff            mov dword ptr [ebp+fffffee8], ecx
'00635231    8b10                    mov edx, dword ptr [eax]
'00635233    8d7d94                  lea edi, dword ptr [ebp-6c]
'00635236    57                      push edi
'00635237    83ec10                  sub esp, 10
'0063523a    8bfc                    mov edi, esp
'0063523c    890f                    mov dword ptr [edi], ecx
'0063523e    8b8decfeffff            mov ecx, dword ptr [ebp+fffffeec]
'00635244    894f04                  mov dword ptr [edi+04], ecx
'00635247    8b8df0feffff            mov ecx, dword ptr [ebp+fffffef0]
'0063524d    894f08                  mov dword ptr [edi+08], ecx
'00635250    8b8df4feffff            mov ecx, dword ptr [ebp+fffffef4]
'00635256    894f0c                  mov dword ptr [edi+0c], ecx
'00635259    50                      push eax
'0063525a    ff5230                  call dword ptr [edx+30]
'0063525d    dbe2                    fnclex
'0063525f    85c0                    test ax, ax
'00635261    7d11                    jge 635274
'00635263    6a30                    push 30
'00635265    68d8304300              push 004330d8
'0063526a    8b9568feffff            mov edx, dword ptr [ebp+fffffe68]
'00635270    52                      push edx
'00635271    50                      push eax
'00635272    ffd3                    call ebx
'00635274    8b4594                  mov eax, dword ptr [ebp-6c]
'00635277    8bf8                    mov edi, eax
'00635279    8b08                    mov ecx, dword ptr [eax]
'0063527b    8d9578ffffff            lea edx, dword ptr [ebp+ffffff78]
'00635281    52                      push edx
'00635282    50                      push eax
'00635283    ff5144                  call dword ptr [ecx+44]
'00635286    dbe2                    fnclex
'00635288    85c0                    test ax, ax
'0063528a    7d0b                    jge 635297
'0063528c    6a44                    push 44
'0063528e    6880304300              push 00433080
'00635293    57                      push edi
'00635294    50                      push eax
'00635295    ffd3                    call ebx
'00635297    8d8578ffffff            lea eax, dword ptr [ebp+ffffff78]
'0063529d    50                      push eax

' *** Reference to "__vbaStrVarMove"
'0063529e    ff153c104000            call dword ptr [0040103c]
'006352a4    8bd0                    mov edx, eax
'006352a6    8d4ddc                  lea ecx, dword ptr [ebp-24]

' *** Reference to "__vbaStrMove"
'006352a9    ff15d0124000            call dword ptr [004012d0]
'006352af    8d4d94                  lea ecx, dword ptr [ebp-6c]
'006352b2    51                      push ecx
'006352b3    8d5598                  lea edx, dword ptr [ebp-68]
'006352b6    52                      push edx
'006352b7    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'006352b9    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_130, var_121)
'006352bf    83c40c                  add esp, 0c
'006352c2    8d8d78ffffff            lea ecx, dword ptr [ebp+ffffff78]

' *** Reference to "__vbaFreeVar"
'006352c8    ff1528104000            call dword ptr [00401028]
'#Cleanup(var_87)
'006352ce    8b45c0                  mov eax, dword ptr [ebp-40]
'006352d1    8b08                    mov ecx, dword ptr [eax]
'006352d3    8d5598                  lea edx, dword ptr [ebp-68]
'006352d6    52                      push edx
'006352d7    50                      push eax
'006352d8    ff91b4000000            call dword ptr [ecx+000000b4]
'006352de    dbe2                    fnclex
'006352e0    85c0                    test ax, ax
'006352e2    7d11                    jge 6352f5
'006352e4    68b4000000              push 000000b4
'006352e9    6830314300              push 00433130
'006352ee    8b4dc0                  mov ecx, dword ptr [ebp-40]
'006352f1    51                      push ecx
'006352f2    50                      push eax
'006352f3    ffd3                    call ebx
'006352f5    8b4598                  mov eax, dword ptr [ebp-68]
'006352f8    898568feffff            mov dword ptr [ebp+fffffe68], eax
'006352fe    c785f0feffff34d44300    mov dword ptr [ebp+fffffef0], 0043d434
'00635308    b908000000              mov ecx, 00000008
'0063530d    898de8feffff            mov dword ptr [ebp+fffffee8], ecx
'00635313    8b10                    mov edx, dword ptr [eax]
'00635315    8d7d94                  lea edi, dword ptr [ebp-6c]
'00635318    57                      push edi
'00635319    83ec10                  sub esp, 10
'0063531c    8bfc                    mov edi, esp
'0063531e    890f                    mov dword ptr [edi], ecx
'00635320    8b8decfeffff            mov ecx, dword ptr [ebp+fffffeec]
'00635326    894f04                  mov dword ptr [edi+04], ecx
'00635329    8b8df0feffff            mov ecx, dword ptr [ebp+fffffef0]
'0063532f    894f08                  mov dword ptr [edi+08], ecx
'00635332    8b8df4feffff            mov ecx, dword ptr [ebp+fffffef4]
'00635338    894f0c                  mov dword ptr [edi+0c], ecx
'0063533b    50                      push eax
'0063533c    ff5230                  call dword ptr [edx+30]
'0063533f    dbe2                    fnclex
'00635341    85c0                    test ax, ax
'00635343    7d11                    jge 635356
'00635345    6a30                    push 30
'00635347    68d8304300              push 004330d8
'0063534c    8b9568feffff            mov edx, dword ptr [ebp+fffffe68]
'00635352    52                      push edx
'00635353    50                      push eax
'00635354    ffd3                    call ebx
'00635356    8b4594                  mov eax, dword ptr [ebp-6c]
'00635359    8bf8                    mov edi, eax
'0063535b    8b08                    mov ecx, dword ptr [eax]
'0063535d    8d9578ffffff            lea edx, dword ptr [ebp+ffffff78]
'00635363    52                      push edx
'00635364    50                      push eax
'00635365    ff5144                  call dword ptr [ecx+44]
'00635368    dbe2                    fnclex
'0063536a    85c0                    test ax, ax
'0063536c    7d0b                    jge 635379
'0063536e    6a44                    push 44
'00635370    6880304300              push 00433080
'00635375    57                      push edi
'00635376    50                      push eax
'00635377    ffd3                    call ebx
'00635379    8d8578ffffff            lea eax, dword ptr [ebp+ffffff78]
'0063537f    50                      push eax

' *** Reference to "__vbaI2Var"
'00635380    ff150c124000            call dword ptr [0040120c]
'00635386    8bf8                    mov edi, eax
'00635388    8d4d94                  lea ecx, dword ptr [ebp-6c]
'0063538b    51                      push ecx
'0063538c    8d5598                  lea edx, dword ptr [ebp-68]
'0063538f    52                      push edx
'00635390    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'00635392    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_130, var_121)
'00635398    83c40c                  add esp, 0c
'0063539b    8d8d78ffffff            lea ecx, dword ptr [ebp+ffffff78]

' *** Reference to "__vbaFreeVar"
'006353a1    ff1528104000            call dword ptr [00401028]
'#Cleanup(var_87)
'006353a7    663bf7                  cmp si, di
'006353aa    0f8ee4000000            jle 635494

If (CInt(IIf(IsNull(var_121), -744 - 16, -256 - 20)) > CInt(var_121)) Then
'006353b0    8b45c0                  mov eax, dword ptr [ebp-40]
'006353b3    8b08                    mov ecx, dword ptr [eax]
'006353b5    8d5598                  lea edx, dword ptr [ebp-68]
'006353b8    52                      push edx
'006353b9    50                      push eax
'006353ba    ff91b4000000            call dword ptr [ecx+000000b4]
'006353c0    dbe2                    fnclex
'006353c2    85c0                    test ax, ax
'006353c4    7d11                    jge 6353d7
'006353c6    68b4000000              push 000000b4
'006353cb    6830314300              push 00433130
'006353d0    8b4dc0                  mov ecx, dword ptr [ebp-40]
'006353d3    51                      push ecx
'006353d4    50                      push eax
'006353d5    ffd3                    call ebx
'006353d7    8b4598                  mov eax, dword ptr [ebp-68]
'006353da    898568feffff            mov dword ptr [ebp+fffffe68], eax
'006353e0    c785f0feffff44d44300    mov dword ptr [ebp+fffffef0], 0043d444
'006353ea    b908000000              mov ecx, 00000008
'006353ef    898de8feffff            mov dword ptr [ebp+fffffee8], ecx
'006353f5    8b10                    mov edx, dword ptr [eax]
'006353f7    8d7d94                  lea edi, dword ptr [ebp-6c]
'006353fa    57                      push edi
'006353fb    83ec10                  sub esp, 10
'006353fe    8bfc                    mov edi, esp
'00635400    890f                    mov dword ptr [edi], ecx
'00635402    8b8decfeffff            mov ecx, dword ptr [ebp+fffffeec]
'00635408    894f04                  mov dword ptr [edi+04], ecx
'0063540b    8b8df0feffff            mov ecx, dword ptr [ebp+fffffef0]
'00635411    894f08                  mov dword ptr [edi+08], ecx
'00635414    8b8df4feffff            mov ecx, dword ptr [ebp+fffffef4]
'0063541a    894f0c                  mov dword ptr [edi+0c], ecx
'0063541d    50                      push eax
'0063541e    ff5230                  call dword ptr [edx+30]
'00635421    dbe2                    fnclex
'00635423    85c0                    test ax, ax
'00635425    7d11                    jge 635438
'00635427    6a30                    push 30
'00635429    68d8304300              push 004330d8
'0063542e    8b9568feffff            mov edx, dword ptr [ebp+fffffe68]
'00635434    52                      push edx
'00635435    50                      push eax
'00635436    ffd3                    call ebx
'00635438    8b4594                  mov eax, dword ptr [ebp-6c]
'0063543b    8bf8                    mov edi, eax
'0063543d    8b08                    mov ecx, dword ptr [eax]
'0063543f    8d9578ffffff            lea edx, dword ptr [ebp+ffffff78]
'00635445    52                      push edx
'00635446    50                      push eax
'00635447    ff5144                  call dword ptr [ecx+44]
'0063544a    dbe2                    fnclex
'0063544c    85c0                    test ax, ax
'0063544e    7d0b                    jge 63545b
'00635450    6a44                    push 44
'00635452    6880304300              push 00433080
'00635457    57                      push edi
'00635458    50                      push eax
'00635459    ffd3                    call ebx
'0063545b    8d8578ffffff            lea eax, dword ptr [ebp+ffffff78]
'00635461    50                      push eax

' *** Reference to "__vbaStrVarMove"
'00635462    ff153c104000            call dword ptr [0040103c]
'00635468    8bd0                    mov edx, eax
'0063546a    8d4ddc                  lea ecx, dword ptr [ebp-24]

' *** Reference to "__vbaStrMove"
'0063546d    ff15d0124000            call dword ptr [004012d0]
'00635473    8d4d94                  lea ecx, dword ptr [ebp-6c]
'00635476    51                      push ecx
'00635477    8d5598                  lea edx, dword ptr [ebp-68]
'0063547a    52                      push edx
'0063547b    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'0063547d    ff154c104000            call dword ptr [0040104c]
    '#Cleanup( var_130, var_121)
'00635483    83c40c                  add esp, 0c
'00635486    8d8d78ffffff            lea ecx, dword ptr [ebp+ffffff78]

' *** Reference to "__vbaFreeVar"
'0063548c    ff1528104000            call dword ptr [00401028]
    '#Cleanup(var_87)
'00635492    8bfe                    mov edi, esi
    
End If
'00635494    66397de0                cmp word ptr [ebp-20], di
'00635498    0f8ee5000000            jle 635583

If (CInt(IIf(IsNull(var_121), -756 - 16, -744 - 16)) > CInt(var_121)) Then
'0063549e    8b45c0                  mov eax, dword ptr [ebp-40]
'006354a1    8b08                    mov ecx, dword ptr [eax]
'006354a3    8d5598                  lea edx, dword ptr [ebp-68]
'006354a6    52                      push edx
'006354a7    50                      push eax
'006354a8    ff91b4000000            call dword ptr [ecx+000000b4]
'006354ae    dbe2                    fnclex
'006354b0    85c0                    test ax, ax
'006354b2    7d11                    jge 6354c5
'006354b4    68b4000000              push 000000b4
'006354b9    6830314300              push 00433130
'006354be    8b4dc0                  mov ecx, dword ptr [ebp-40]
'006354c1    51                      push ecx
'006354c2    50                      push eax
'006354c3    ffd3                    call ebx
'006354c5    8b4598                  mov eax, dword ptr [ebp-68]
'006354c8    898568feffff            mov dword ptr [ebp+fffffe68], eax
'006354ce    c785f0feffff6cd44300    mov dword ptr [ebp+fffffef0], 0043d46c
'006354d8    b908000000              mov ecx, 00000008
'006354dd    898de8feffff            mov dword ptr [ebp+fffffee8], ecx
'006354e3    8b10                    mov edx, dword ptr [eax]
'006354e5    8d7d94                  lea edi, dword ptr [ebp-6c]
'006354e8    57                      push edi
'006354e9    83ec10                  sub esp, 10
'006354ec    8bfc                    mov edi, esp
'006354ee    890f                    mov dword ptr [edi], ecx
'006354f0    8b8decfeffff            mov ecx, dword ptr [ebp+fffffeec]
'006354f6    894f04                  mov dword ptr [edi+04], ecx
'006354f9    8b8df0feffff            mov ecx, dword ptr [ebp+fffffef0]
'006354ff    894f08                  mov dword ptr [edi+08], ecx
'00635502    8b8df4feffff            mov ecx, dword ptr [ebp+fffffef4]
'00635508    894f0c                  mov dword ptr [edi+0c], ecx
'0063550b    50                      push eax
'0063550c    ff5230                  call dword ptr [edx+30]
'0063550f    dbe2                    fnclex
'00635511    85c0                    test ax, ax
'00635513    7d11                    jge 635526
'00635515    6a30                    push 30
'00635517    68d8304300              push 004330d8
'0063551c    8b9568feffff            mov edx, dword ptr [ebp+fffffe68]
'00635522    52                      push edx
'00635523    50                      push eax
'00635524    ffd3                    call ebx
'00635526    8b4594                  mov eax, dword ptr [ebp-6c]
'00635529    8bf8                    mov edi, eax
'0063552b    8b08                    mov ecx, dword ptr [eax]
'0063552d    8d9578ffffff            lea edx, dword ptr [ebp+ffffff78]
'00635533    52                      push edx
'00635534    50                      push eax
'00635535    ff5144                  call dword ptr [ecx+44]
'00635538    dbe2                    fnclex
'0063553a    85c0                    test ax, ax
'0063553c    7d0b                    jge 635549
'0063553e    6a44                    push 44
'00635540    6880304300              push 00433080
'00635545    57                      push edi
'00635546    50                      push eax
'00635547    ffd3                    call ebx
'00635549    8d8578ffffff            lea eax, dword ptr [ebp+ffffff78]
'0063554f    50                      push eax

' *** Reference to "__vbaStrVarMove"
'00635550    ff153c104000            call dword ptr [0040103c]
'00635556    8bd0                    mov edx, eax
'00635558    8d4ddc                  lea ecx, dword ptr [ebp-24]

' *** Reference to "__vbaStrMove"
'0063555b    ff15d0124000            call dword ptr [004012d0]
'00635561    8d4d94                  lea ecx, dword ptr [ebp-6c]
'00635564    51                      push ecx
'00635565    8d5598                  lea edx, dword ptr [ebp-68]
'00635568    52                      push edx
'00635569    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'0063556b    ff154c104000            call dword ptr [0040104c]
    '#Cleanup( var_130, var_121)
'00635571    83c40c                  add esp, 0c
'00635574    8d8d78ffffff            lea ecx, dword ptr [ebp+ffffff78]

' *** Reference to "__vbaFreeVar"
'0063557a    ff1528104000            call dword ptr [00401028]
    '#Cleanup(var_87)
'00635580    8b7de0                  mov edi, dword ptr [ebp-20]
    
End If
'00635583    66397dd8                cmp word ptr [ebp-28], di
'00635587    0f8ee5000000            jle 635672

If (CInt(IIf(IsNull(var_121), -768 - 16, -756 - 16)) > CInt(var_121)) Then
'0063558d    8b45c0                  mov eax, dword ptr [ebp-40]
'00635590    8b08                    mov ecx, dword ptr [eax]
'00635592    8d5598                  lea edx, dword ptr [ebp-68]
'00635595    52                      push edx
'00635596    50                      push eax
'00635597    ff91b4000000            call dword ptr [ecx+000000b4]
'0063559d    dbe2                    fnclex
'0063559f    85c0                    test ax, ax
'006355a1    7d11                    jge 6355b4
'006355a3    68b4000000              push 000000b4
'006355a8    6830314300              push 00433130
'006355ad    8b4dc0                  mov ecx, dword ptr [ebp-40]
'006355b0    51                      push ecx
'006355b1    50                      push eax
'006355b2    ffd3                    call ebx
'006355b4    8b4598                  mov eax, dword ptr [ebp-68]
'006355b7    898568feffff            mov dword ptr [ebp+fffffe68], eax
'006355bd    c785f0feffff0cad4300    mov dword ptr [ebp+fffffef0], 0043ad0c
'006355c7    b908000000              mov ecx, 00000008
'006355cc    898de8feffff            mov dword ptr [ebp+fffffee8], ecx
'006355d2    8b10                    mov edx, dword ptr [eax]
'006355d4    8d7d94                  lea edi, dword ptr [ebp-6c]
'006355d7    57                      push edi
'006355d8    83ec10                  sub esp, 10
'006355db    8bfc                    mov edi, esp
'006355dd    890f                    mov dword ptr [edi], ecx
'006355df    8b8decfeffff            mov ecx, dword ptr [ebp+fffffeec]
'006355e5    894f04                  mov dword ptr [edi+04], ecx
'006355e8    8b8df0feffff            mov ecx, dword ptr [ebp+fffffef0]
'006355ee    894f08                  mov dword ptr [edi+08], ecx
'006355f1    8b8df4feffff            mov ecx, dword ptr [ebp+fffffef4]
'006355f7    894f0c                  mov dword ptr [edi+0c], ecx
'006355fa    50                      push eax
'006355fb    ff5230                  call dword ptr [edx+30]
'006355fe    dbe2                    fnclex
'00635600    85c0                    test ax, ax
'00635602    7d11                    jge 635615
'00635604    6a30                    push 30
'00635606    68d8304300              push 004330d8
'0063560b    8b9568feffff            mov edx, dword ptr [ebp+fffffe68]
'00635611    52                      push edx
'00635612    50                      push eax
'00635613    ffd3                    call ebx
'00635615    8b4594                  mov eax, dword ptr [ebp-6c]
'00635618    8bf8                    mov edi, eax
'0063561a    8b08                    mov ecx, dword ptr [eax]
'0063561c    8d9578ffffff            lea edx, dword ptr [ebp+ffffff78]
'00635622    52                      push edx
'00635623    50                      push eax
'00635624    ff5144                  call dword ptr [ecx+44]
'00635627    dbe2                    fnclex
'00635629    85c0                    test ax, ax
'0063562b    7d0b                    jge 635638
'0063562d    6a44                    push 44
'0063562f    6880304300              push 00433080
'00635634    57                      push edi
'00635635    50                      push eax
'00635636    ffd3                    call ebx
'00635638    8d8578ffffff            lea eax, dword ptr [ebp+ffffff78]
'0063563e    50                      push eax

' *** Reference to "__vbaStrVarMove"
'0063563f    ff153c104000            call dword ptr [0040103c]
'00635645    8bd0                    mov edx, eax
'00635647    8d4ddc                  lea ecx, dword ptr [ebp-24]

' *** Reference to "__vbaStrMove"
'0063564a    ff15d0124000            call dword ptr [004012d0]
'00635650    8d4d94                  lea ecx, dword ptr [ebp-6c]
'00635653    51                      push ecx
'00635654    8d5598                  lea edx, dword ptr [ebp-68]
'00635657    52                      push edx
'00635658    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'0063565a    ff154c104000            call dword ptr [0040104c]
    '#Cleanup( var_130, var_121)
'00635660    83c40c                  add esp, 0c
'00635663    8d8d78ffffff            lea ecx, dword ptr [ebp+ffffff78]

' *** Reference to "__vbaFreeVar"
'00635669    ff1528104000            call dword ptr [00401028]
    '#Cleanup(var_87)
'0063566f    8b7dd8                  mov edi, dword ptr [ebp-28]
    
End If
'00635672    66397dd0                cmp word ptr [ebp-30], di
'00635676    0f8ee5000000            jle 635761

If (CInt(IIf(IsNull(var_121), -780 - 16, -768 - 16)) > CInt(var_121)) Then
'0063567c    8b45c0                  mov eax, dword ptr [ebp-40]
'0063567f    8b08                    mov ecx, dword ptr [eax]
'00635681    8d5598                  lea edx, dword ptr [ebp-68]
'00635684    52                      push edx
'00635685    50                      push eax
'00635686    ff91b4000000            call dword ptr [ecx+000000b4]
'0063568c    dbe2                    fnclex
'0063568e    85c0                    test ax, ax
'00635690    7d11                    jge 6356a3
'00635692    68b4000000              push 000000b4
'00635697    6830314300              push 00433130
'0063569c    8b4dc0                  mov ecx, dword ptr [ebp-40]
'0063569f    51                      push ecx
'006356a0    50                      push eax
'006356a1    ffd3                    call ebx
'006356a3    8b4598                  mov eax, dword ptr [ebp-68]
'006356a6    898568feffff            mov dword ptr [ebp+fffffe68], eax
'006356ac    c785f0feffff24ad4300    mov dword ptr [ebp+fffffef0], 0043ad24
'006356b6    b908000000              mov ecx, 00000008
'006356bb    898de8feffff            mov dword ptr [ebp+fffffee8], ecx
'006356c1    8b10                    mov edx, dword ptr [eax]
'006356c3    8d7d94                  lea edi, dword ptr [ebp-6c]
'006356c6    57                      push edi
'006356c7    83ec10                  sub esp, 10
'006356ca    8bfc                    mov edi, esp
'006356cc    890f                    mov dword ptr [edi], ecx
'006356ce    8b8decfeffff            mov ecx, dword ptr [ebp+fffffeec]
'006356d4    894f04                  mov dword ptr [edi+04], ecx
'006356d7    8b8df0feffff            mov ecx, dword ptr [ebp+fffffef0]
'006356dd    894f08                  mov dword ptr [edi+08], ecx
'006356e0    8b8df4feffff            mov ecx, dword ptr [ebp+fffffef4]
'006356e6    894f0c                  mov dword ptr [edi+0c], ecx
'006356e9    50                      push eax
'006356ea    ff5230                  call dword ptr [edx+30]
'006356ed    dbe2                    fnclex
'006356ef    85c0                    test ax, ax
'006356f1    7d11                    jge 635704
'006356f3    6a30                    push 30
'006356f5    68d8304300              push 004330d8
'006356fa    8b9568feffff            mov edx, dword ptr [ebp+fffffe68]
'00635700    52                      push edx
'00635701    50                      push eax
'00635702    ffd3                    call ebx
'00635704    8b4594                  mov eax, dword ptr [ebp-6c]
'00635707    8bf8                    mov edi, eax
'00635709    8b08                    mov ecx, dword ptr [eax]
'0063570b    8d9578ffffff            lea edx, dword ptr [ebp+ffffff78]
'00635711    52                      push edx
'00635712    50                      push eax
'00635713    ff5144                  call dword ptr [ecx+44]
'00635716    dbe2                    fnclex
'00635718    85c0                    test ax, ax
'0063571a    7d0b                    jge 635727
'0063571c    6a44                    push 44
'0063571e    6880304300              push 00433080
'00635723    57                      push edi
'00635724    50                      push eax
'00635725    ffd3                    call ebx
'00635727    8d8578ffffff            lea eax, dword ptr [ebp+ffffff78]
'0063572d    50                      push eax

' *** Reference to "__vbaStrVarMove"
'0063572e    ff153c104000            call dword ptr [0040103c]
'00635734    8bd0                    mov edx, eax
'00635736    8d4ddc                  lea ecx, dword ptr [ebp-24]

' *** Reference to "__vbaStrMove"
'00635739    ff15d0124000            call dword ptr [004012d0]
'0063573f    8d4d94                  lea ecx, dword ptr [ebp-6c]
'00635742    51                      push ecx
'00635743    8d5598                  lea edx, dword ptr [ebp-68]
'00635746    52                      push edx
'00635747    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'00635749    ff154c104000            call dword ptr [0040104c]
    '#Cleanup( var_130, var_121)
'0063574f    83c40c                  add esp, 0c
'00635752    8d8d78ffffff            lea ecx, dword ptr [ebp+ffffff78]

' *** Reference to "__vbaFreeVar"
'00635758    ff1528104000            call dword ptr [00401028]
    '#Cleanup(var_87)
'0063575e    8b7dd0                  mov edi, dword ptr [ebp-30]
    
End If
'00635761    66397dcc                cmp word ptr [ebp-34], di
'00635765    0f8ee5000000            jle 635850

If (CInt(IIf(IsNull(var_121), -792 - 16, -780 - 16)) > CInt(var_121)) Then
'0063576b    8b45c0                  mov eax, dword ptr [ebp-40]
'0063576e    8b08                    mov ecx, dword ptr [eax]
'00635770    8d5598                  lea edx, dword ptr [ebp-68]
'00635773    52                      push edx
'00635774    50                      push eax
'00635775    ff91b4000000            call dword ptr [ecx+000000b4]
'0063577b    dbe2                    fnclex
'0063577d    85c0                    test ax, ax
'0063577f    7d11                    jge 635792
'00635781    68b4000000              push 000000b4
'00635786    6830314300              push 00433130
'0063578b    8b4dc0                  mov ecx, dword ptr [ebp-40]
'0063578e    51                      push ecx
'0063578f    50                      push eax
'00635790    ffd3                    call ebx
'00635792    8b4598                  mov eax, dword ptr [ebp-68]
'00635795    898568feffff            mov dword ptr [ebp+fffffe68], eax
'0063579b    c785f0feffff3cad4300    mov dword ptr [ebp+fffffef0], 0043ad3c
'006357a5    b908000000              mov ecx, 00000008
'006357aa    898de8feffff            mov dword ptr [ebp+fffffee8], ecx
'006357b0    8b10                    mov edx, dword ptr [eax]
'006357b2    8d7d94                  lea edi, dword ptr [ebp-6c]
'006357b5    57                      push edi
'006357b6    83ec10                  sub esp, 10
'006357b9    8bfc                    mov edi, esp
'006357bb    890f                    mov dword ptr [edi], ecx
'006357bd    8b8decfeffff            mov ecx, dword ptr [ebp+fffffeec]
'006357c3    894f04                  mov dword ptr [edi+04], ecx
'006357c6    8b8df0feffff            mov ecx, dword ptr [ebp+fffffef0]
'006357cc    894f08                  mov dword ptr [edi+08], ecx
'006357cf    8b8df4feffff            mov ecx, dword ptr [ebp+fffffef4]
'006357d5    894f0c                  mov dword ptr [edi+0c], ecx
'006357d8    50                      push eax
'006357d9    ff5230                  call dword ptr [edx+30]
'006357dc    dbe2                    fnclex
'006357de    85c0                    test ax, ax
'006357e0    7d11                    jge 6357f3
'006357e2    6a30                    push 30
'006357e4    68d8304300              push 004330d8
'006357e9    8b9568feffff            mov edx, dword ptr [ebp+fffffe68]
'006357ef    52                      push edx
'006357f0    50                      push eax
'006357f1    ffd3                    call ebx
'006357f3    8b4594                  mov eax, dword ptr [ebp-6c]
'006357f6    8bf8                    mov edi, eax
'006357f8    8b08                    mov ecx, dword ptr [eax]
'006357fa    8d9578ffffff            lea edx, dword ptr [ebp+ffffff78]
'00635800    52                      push edx
'00635801    50                      push eax
'00635802    ff5144                  call dword ptr [ecx+44]
'00635805    dbe2                    fnclex
'00635807    85c0                    test ax, ax
'00635809    7d0b                    jge 635816
'0063580b    6a44                    push 44
'0063580d    6880304300              push 00433080
'00635812    57                      push edi
'00635813    50                      push eax
'00635814    ffd3                    call ebx
'00635816    8d8578ffffff            lea eax, dword ptr [ebp+ffffff78]
'0063581c    50                      push eax

' *** Reference to "__vbaStrVarMove"
'0063581d    ff153c104000            call dword ptr [0040103c]
'00635823    8bd0                    mov edx, eax
'00635825    8d4ddc                  lea ecx, dword ptr [ebp-24]

' *** Reference to "__vbaStrMove"
'00635828    ff15d0124000            call dword ptr [004012d0]
'0063582e    8d4d94                  lea ecx, dword ptr [ebp-6c]
'00635831    51                      push ecx
'00635832    8d5598                  lea edx, dword ptr [ebp-68]
'00635835    52                      push edx
'00635836    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'00635838    ff154c104000            call dword ptr [0040104c]
    '#Cleanup( var_130, var_121)
'0063583e    83c40c                  add esp, 0c
'00635841    8d8d78ffffff            lea ecx, dword ptr [ebp+ffffff78]

' *** Reference to "__vbaFreeVar"
'00635847    ff1528104000            call dword ptr [00401028]
    '#Cleanup(var_87)
'0063584d    8b7dcc                  mov edi, dword ptr [ebp-34]
    
End If
'00635850    66397dc8                cmp word ptr [ebp-38], di
'00635854    0f8ee5000000            jle 63593f

If (CInt(IIf(IsNull(var_121), -804 - 16, -792 - 16)) > CInt(var_121)) Then
'0063585a    8b45c0                  mov eax, dword ptr [ebp-40]
'0063585d    8b08                    mov ecx, dword ptr [eax]
'0063585f    8d5598                  lea edx, dword ptr [ebp-68]
'00635862    52                      push edx
'00635863    50                      push eax
'00635864    ff91b4000000            call dword ptr [ecx+000000b4]
'0063586a    dbe2                    fnclex
'0063586c    85c0                    test ax, ax
'0063586e    7d11                    jge 635881
'00635870    68b4000000              push 000000b4
'00635875    6830314300              push 00433130
'0063587a    8b4dc0                  mov ecx, dword ptr [ebp-40]
'0063587d    51                      push ecx
'0063587e    50                      push eax
'0063587f    ffd3                    call ebx
'00635881    8b4598                  mov eax, dword ptr [ebp-68]
'00635884    898568feffff            mov dword ptr [ebp+fffffe68], eax
'0063588a    c785f0feffffb4b54300    mov dword ptr [ebp+fffffef0], 0043b5b4
'00635894    b908000000              mov ecx, 00000008
'00635899    898de8feffff            mov dword ptr [ebp+fffffee8], ecx
'0063589f    8b10                    mov edx, dword ptr [eax]
'006358a1    8d7d94                  lea edi, dword ptr [ebp-6c]
'006358a4    57                      push edi
'006358a5    83ec10                  sub esp, 10
'006358a8    8bfc                    mov edi, esp
'006358aa    890f                    mov dword ptr [edi], ecx
'006358ac    8b8decfeffff            mov ecx, dword ptr [ebp+fffffeec]
'006358b2    894f04                  mov dword ptr [edi+04], ecx
'006358b5    8b8df0feffff            mov ecx, dword ptr [ebp+fffffef0]
'006358bb    894f08                  mov dword ptr [edi+08], ecx
'006358be    8b8df4feffff            mov ecx, dword ptr [ebp+fffffef4]
'006358c4    894f0c                  mov dword ptr [edi+0c], ecx
'006358c7    50                      push eax
'006358c8    ff5230                  call dword ptr [edx+30]
'006358cb    dbe2                    fnclex
'006358cd    85c0                    test ax, ax
'006358cf    7d11                    jge 6358e2
'006358d1    6a30                    push 30
'006358d3    68d8304300              push 004330d8
'006358d8    8b9568feffff            mov edx, dword ptr [ebp+fffffe68]
'006358de    52                      push edx
'006358df    50                      push eax
'006358e0    ffd3                    call ebx
'006358e2    8b4594                  mov eax, dword ptr [ebp-6c]
'006358e5    8bf8                    mov edi, eax
'006358e7    8b08                    mov ecx, dword ptr [eax]
'006358e9    8d9578ffffff            lea edx, dword ptr [ebp+ffffff78]
'006358ef    52                      push edx
'006358f0    50                      push eax
'006358f1    ff5144                  call dword ptr [ecx+44]
'006358f4    dbe2                    fnclex
'006358f6    85c0                    test ax, ax
'006358f8    7d0b                    jge 635905
'006358fa    6a44                    push 44
'006358fc    6880304300              push 00433080
'00635901    57                      push edi
'00635902    50                      push eax
'00635903    ffd3                    call ebx
'00635905    8d8578ffffff            lea eax, dword ptr [ebp+ffffff78]
'0063590b    50                      push eax

' *** Reference to "__vbaStrVarMove"
'0063590c    ff153c104000            call dword ptr [0040103c]
'00635912    8bd0                    mov edx, eax
'00635914    8d4ddc                  lea ecx, dword ptr [ebp-24]

' *** Reference to "__vbaStrMove"
'00635917    ff15d0124000            call dword ptr [004012d0]
'0063591d    8d4d94                  lea ecx, dword ptr [ebp-6c]
'00635920    51                      push ecx
'00635921    8d5598                  lea edx, dword ptr [ebp-68]
'00635924    52                      push edx
'00635925    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'00635927    ff154c104000            call dword ptr [0040104c]
    '#Cleanup( var_130, var_121)
'0063592d    83c40c                  add esp, 0c
'00635930    8d8d78ffffff            lea ecx, dword ptr [ebp+ffffff78]

' *** Reference to "__vbaFreeVar"
'00635936    ff1528104000            call dword ptr [00401028]
    '#Cleanup(var_87)
'0063593c    8b7dc8                  mov edi, dword ptr [ebp-38]
    
End If
'0063593f    66397dc4                cmp word ptr [ebp-3c], di
'00635943    0f8ee2000000            jle 635a2b

If (CInt(IIf(IsNull(var_121), -816 - 16, -804 - 16)) > CInt(var_121)) Then
'00635949    8b45c0                  mov eax, dword ptr [ebp-40]
'0063594c    8b08                    mov ecx, dword ptr [eax]
'0063594e    8d5598                  lea edx, dword ptr [ebp-68]
'00635951    52                      push edx
'00635952    50                      push eax
'00635953    ff91b4000000            call dword ptr [ecx+000000b4]
'00635959    dbe2                    fnclex
'0063595b    85c0                    test ax, ax
'0063595d    7d11                    jge 635970
'0063595f    68b4000000              push 000000b4
'00635964    6830314300              push 00433130
'00635969    8b4dc0                  mov ecx, dword ptr [ebp-40]
'0063596c    51                      push ecx
'0063596d    50                      push eax
'0063596e    ffd3                    call ebx
'00635970    8b4598                  mov eax, dword ptr [ebp-68]
'00635973    898568feffff            mov dword ptr [ebp+fffffe68], eax
'00635979    c785f0feffffccb54300    mov dword ptr [ebp+fffffef0], 0043b5cc
'00635983    b908000000              mov ecx, 00000008
'00635988    898de8feffff            mov dword ptr [ebp+fffffee8], ecx
'0063598e    8b10                    mov edx, dword ptr [eax]
'00635990    8d7d94                  lea edi, dword ptr [ebp-6c]
'00635993    57                      push edi
'00635994    83ec10                  sub esp, 10
'00635997    8bfc                    mov edi, esp
'00635999    890f                    mov dword ptr [edi], ecx
'0063599b    8b8decfeffff            mov ecx, dword ptr [ebp+fffffeec]
'006359a1    894f04                  mov dword ptr [edi+04], ecx
'006359a4    8b8df0feffff            mov ecx, dword ptr [ebp+fffffef0]
'006359aa    894f08                  mov dword ptr [edi+08], ecx
'006359ad    8b8df4feffff            mov ecx, dword ptr [ebp+fffffef4]
'006359b3    894f0c                  mov dword ptr [edi+0c], ecx
'006359b6    50                      push eax
'006359b7    ff5230                  call dword ptr [edx+30]
'006359ba    dbe2                    fnclex
'006359bc    85c0                    test ax, ax
'006359be    7d11                    jge 6359d1
'006359c0    6a30                    push 30
'006359c2    68d8304300              push 004330d8
'006359c7    8b9568feffff            mov edx, dword ptr [ebp+fffffe68]
'006359cd    52                      push edx
'006359ce    50                      push eax
'006359cf    ffd3                    call ebx
'006359d1    8b4594                  mov eax, dword ptr [ebp-6c]
'006359d4    8bf8                    mov edi, eax
'006359d6    8b08                    mov ecx, dword ptr [eax]
'006359d8    8d9578ffffff            lea edx, dword ptr [ebp+ffffff78]
'006359de    52                      push edx
'006359df    50                      push eax
'006359e0    ff5144                  call dword ptr [ecx+44]
'006359e3    dbe2                    fnclex
'006359e5    85c0                    test ax, ax
'006359e7    7d0b                    jge 6359f4
'006359e9    6a44                    push 44
'006359eb    6880304300              push 00433080
'006359f0    57                      push edi
'006359f1    50                      push eax
'006359f2    ffd3                    call ebx
'006359f4    8d8578ffffff            lea eax, dword ptr [ebp+ffffff78]
'006359fa    50                      push eax

' *** Reference to "__vbaStrVarMove"
'006359fb    ff153c104000            call dword ptr [0040103c]
'00635a01    8bd0                    mov edx, eax
'00635a03    8d4ddc                  lea ecx, dword ptr [ebp-24]

' *** Reference to "__vbaStrMove"
'00635a06    ff15d0124000            call dword ptr [004012d0]
'00635a0c    8d4d94                  lea ecx, dword ptr [ebp-6c]
'00635a0f    51                      push ecx
'00635a10    8d5598                  lea edx, dword ptr [ebp-68]
'00635a13    52                      push edx
'00635a14    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'00635a16    ff154c104000            call dword ptr [0040104c]
    '#Cleanup( var_130, var_121)
'00635a1c    83c40c                  add esp, 0c
'00635a1f    8d8d78ffffff            lea ecx, dword ptr [ebp+ffffff78]

' *** Reference to "__vbaFreeVar"
'00635a25    ff1528104000            call dword ptr [00401028]
    '#Cleanup(var_87)
End If
'00635a2b    8b45c0                  mov eax, dword ptr [ebp-40]
'00635a2e    8b08                    mov ecx, dword ptr [eax]
'00635a30    8d5598                  lea edx, dword ptr [ebp-68]
'00635a33    52                      push edx
'00635a34    50                      push eax
'00635a35    ff91b4000000            call dword ptr [ecx+000000b4]
'00635a3b    dbe2                    fnclex
'00635a3d    85c0                    test ax, ax
'00635a3f    7d11                    jge 635a52
'00635a41    68b4000000              push 000000b4
'00635a46    6830314300              push 00433130
'00635a4b    8b4dc0                  mov ecx, dword ptr [ebp-40]
'00635a4e    51                      push ecx
'00635a4f    50                      push eax
'00635a50    ffd3                    call ebx
'00635a52    8b4598                  mov eax, dword ptr [ebp-68]
'00635a55    898568feffff            mov dword ptr [ebp+fffffe68], eax
'00635a5b    c785f0feffff64b14300    mov dword ptr [ebp+fffffef0], 0043b164
'00635a65    b908000000              mov ecx, 00000008
'00635a6a    898de8feffff            mov dword ptr [ebp+fffffee8], ecx
'00635a70    8b10                    mov edx, dword ptr [eax]
'00635a72    8d7d94                  lea edi, dword ptr [ebp-6c]
'00635a75    57                      push edi
'00635a76    83ec10                  sub esp, 10
'00635a79    8bfc                    mov edi, esp
'00635a7b    890f                    mov dword ptr [edi], ecx
'00635a7d    8b8decfeffff            mov ecx, dword ptr [ebp+fffffeec]
'00635a83    894f04                  mov dword ptr [edi+04], ecx
'00635a86    8b8df0feffff            mov ecx, dword ptr [ebp+fffffef0]
'00635a8c    894f08                  mov dword ptr [edi+08], ecx
'00635a8f    8b8df4feffff            mov ecx, dword ptr [ebp+fffffef4]
'00635a95    894f0c                  mov dword ptr [edi+0c], ecx
'00635a98    50                      push eax
'00635a99    ff5230                  call dword ptr [edx+30]
'00635a9c    dbe2                    fnclex
'00635a9e    85c0                    test ax, ax
'00635aa0    7d11                    jge 635ab3
'00635aa2    6a30                    push 30
'00635aa4    68d8304300              push 004330d8
'00635aa9    8b9568feffff            mov edx, dword ptr [ebp+fffffe68]
'00635aaf    52                      push edx
'00635ab0    50                      push eax
'00635ab1    ffd3                    call ebx
'00635ab3    8b4594                  mov eax, dword ptr [ebp-6c]
'00635ab6    8bf8                    mov edi, eax
'00635ab8    8b08                    mov ecx, dword ptr [eax]
'00635aba    8d9578ffffff            lea edx, dword ptr [ebp+ffffff78]
'00635ac0    52                      push edx
'00635ac1    50                      push eax
'00635ac2    ff5144                  call dword ptr [ecx+44]
'00635ac5    dbe2                    fnclex
'00635ac7    85c0                    test ax, ax
'00635ac9    7d0b                    jge 635ad6
'00635acb    6a44                    push 44
'00635acd    6880304300              push 00433080
'00635ad2    57                      push edi
'00635ad3    50                      push eax
'00635ad4    ffd3                    call ebx
'00635ad6    c785e0feffff44ed4300    mov dword ptr [ebp+fffffee0], 0043ed44
'00635ae0    c785d8feffff08000000    mov dword ptr [ebp+fffffed8], 00000008
'00635aea    8b45c0                  mov eax, dword ptr [ebp-40]
'00635aed    8b08                    mov ecx, dword ptr [eax]
'00635aef    8d5590                  lea edx, dword ptr [ebp-70]
'00635af2    52                      push edx
'00635af3    50                      push eax
'00635af4    ff91b4000000            call dword ptr [ecx+000000b4]
'00635afa    dbe2                    fnclex
'00635afc    85c0                    test ax, ax
'00635afe    7d11                    jge 635b11
'00635b00    68b4000000              push 000000b4
'00635b05    6830314300              push 00433130
'00635b0a    8b4dc0                  mov ecx, dword ptr [ebp-40]
'00635b0d    51                      push ecx
'00635b0e    50                      push eax
'00635b0f    ffd3                    call ebx
'00635b11    8b4590                  mov eax, dword ptr [ebp-70]
'00635b14    898554feffff            mov dword ptr [ebp+fffffe54], eax
'00635b1a    c785d0feffffe4b14300    mov dword ptr [ebp+fffffed0], 0043b1e4
'00635b24    b908000000              mov ecx, 00000008
'00635b29    898dc8feffff            mov dword ptr [ebp+fffffec8], ecx
'00635b2f    8b10                    mov edx, dword ptr [eax]
'00635b31    8d7d8c                  lea edi, dword ptr [ebp-74]
'00635b34    57                      push edi
'00635b35    83ec10                  sub esp, 10
'00635b38    8bfc                    mov edi, esp
'00635b3a    890f                    mov dword ptr [edi], ecx
'00635b3c    8b8dccfeffff            mov ecx, dword ptr [ebp+fffffecc]
'00635b42    894f04                  mov dword ptr [edi+04], ecx
'00635b45    8b8dd0feffff            mov ecx, dword ptr [ebp+fffffed0]
'00635b4b    894f08                  mov dword ptr [edi+08], ecx
'00635b4e    8b8dd4feffff            mov ecx, dword ptr [ebp+fffffed4]
'00635b54    894f0c                  mov dword ptr [edi+0c], ecx
'00635b57    50                      push eax
'00635b58    ff5230                  call dword ptr [edx+30]
'00635b5b    dbe2                    fnclex
'00635b5d    85c0                    test ax, ax
'00635b5f    7d11                    jge 635b72
'00635b61    6a30                    push 30
'00635b63    68d8304300              push 004330d8
'00635b68    8b9554feffff            mov edx, dword ptr [ebp+fffffe54]
'00635b6e    52                      push edx
'00635b6f    50                      push eax
'00635b70    ffd3                    call ebx
'00635b72    8b458c                  mov eax, dword ptr [ebp-74]
'00635b75    8bf8                    mov edi, eax
'00635b77    8b08                    mov ecx, dword ptr [eax]
'00635b79    8d9558ffffff            lea edx, dword ptr [ebp+ffffff58]
'00635b7f    52                      push edx
'00635b80    50                      push eax
'00635b81    ff5144                  call dword ptr [ecx+44]
'00635b84    dbe2                    fnclex
'00635b86    85c0                    test ax, ax
'00635b88    7d0b                    jge 635b95
'00635b8a    6a44                    push 44
'00635b8c    6880304300              push 00433080
'00635b91    57                      push edi
'00635b92    50                      push eax
'00635b93    ffd3                    call ebx
'00635b95    b844ed4300              mov eax, 0043ed44
'00635b9a    8985c0feffff            mov dword ptr [ebp+fffffec0], eax
'00635ba0    bf08000000              mov edi, 00000008
'00635ba5    89bdb8feffff            mov dword ptr [ebp+fffffeb8], edi
'00635bab    660375bc                add si, word ptr [ebp-44]
var_num8 = CInt(IIf(IsNull(var_121), -744 - 16, -256 - 20)) + CInt(IIf(IsNull(var_121), -256 - 20, 0))
'00635baf    0f80ad020000            jo 635e62
'00635bb5    660375e0                add si, word ptr [ebp-20]
var_num8 = var_num8 + CInt(IIf(IsNull(var_121), -756 - 16, -744 - 16))
'00635bb9    0f80a3020000            jo 635e62
'00635bbf    660375d8                add si, word ptr [ebp-28]
var_num8 = var_num8 + CInt(IIf(IsNull(var_121), -768 - 16, -756 - 16))
'00635bc3    0f8099020000            jo 635e62
'00635bc9    660375d0                add si, word ptr [ebp-30]
var_num8 = var_num8 + CInt(IIf(IsNull(var_121), -780 - 16, -768 - 16))
'00635bcd    0f808f020000            jo 635e62
'00635bd3    660375cc                add si, word ptr [ebp-34]
var_num8 = var_num8 + CInt(IIf(IsNull(var_121), -792 - 16, -780 - 16))
'00635bd7    0f8085020000            jo 635e62
'00635bdd    660375c8                add si, word ptr [ebp-38]
var_num8 = var_num8 + CInt(IIf(IsNull(var_121), -804 - 16, -792 - 16))
'00635be1    0f807b020000            jo 635e62
'00635be7    660375c4                add si, word ptr [ebp-3c]
var_num8 = var_num8 + CInt(IIf(IsNull(var_121), -816 - 16, -804 - 16))
'00635beb    0f8071020000            jo 635e62
'00635bf1    6689b5b0feffff          mov word ptr [ebp+fffffeb0], si
'00635bf8    c785a8feffff02000000    mov dword ptr [ebp+fffffea8], 00000002
'00635c02    8985a0feffff            mov dword ptr [ebp+fffffea0], eax
'00635c08    89bd98feffff            mov dword ptr [ebp+fffffe98], edi
'00635c0e    8b45dc                  mov eax, dword ptr [ebp-24]
'00635c11    898590feffff            mov dword ptr [ebp+fffffe90], eax
'00635c17    89bd88feffff            mov dword ptr [ebp+fffffe88], edi
'00635c1d    8d8d78ffffff            lea ecx, dword ptr [ebp+ffffff78]
'00635c23    51                      push ecx
'00635c24    8d95d8feffff            lea edx, dword ptr [ebp+fffffed8]
'00635c2a    52                      push edx
'00635c2b    8d8568ffffff            lea eax, dword ptr [ebp+ffffff68]
'00635c31    50                      push eax

' *** Reference to "__vbaVarCat"
'00635c32    8b3508124000            mov esi, dword ptr [00401208]
'00635c38    ffd6                    call esi
'00635c3a    50                      push eax
'00635c3b    8d8d58ffffff            lea ecx, dword ptr [ebp+ffffff58]
'00635c41    51                      push ecx
'00635c42    8d9548ffffff            lea edx, dword ptr [ebp+ffffff48]
'00635c48    52                      push edx
'00635c49    ffd6                    call esi
'00635c4b    50                      push eax
'00635c4c    8d85b8feffff            lea eax, dword ptr [ebp+fffffeb8]
'00635c52    50                      push eax
'00635c53    8d8d38ffffff            lea ecx, dword ptr [ebp+ffffff38]
'00635c59    51                      push ecx
'00635c5a    ffd6                    call esi
'00635c5c    50                      push eax
'00635c5d    8d95a8feffff            lea edx, dword ptr [ebp+fffffea8]
'00635c63    52                      push edx
'00635c64    8d8528ffffff            lea eax, dword ptr [ebp+ffffff28]
'00635c6a    50                      push eax
'00635c6b    ffd6                    call esi
'00635c6d    50                      push eax
'00635c6e    8d8d98feffff            lea ecx, dword ptr [ebp+fffffe98]
'00635c74    51                      push ecx
'00635c75    8d9518ffffff            lea edx, dword ptr [ebp+ffffff18]
'00635c7b    52                      push edx
'00635c7c    ffd6                    call esi
'00635c7e    50                      push eax
'00635c7f    8d8588feffff            lea eax, dword ptr [ebp+fffffe88]
'00635c85    50                      push eax
'00635c86    8d8d08ffffff            lea ecx, dword ptr [ebp+ffffff08]
'00635c8c    51                      push ecx
'00635c8d    ffd6                    call esi
'00635c8f    50                      push eax

' *** Reference to "__vbaStrVarMove"
'00635c90    ff153c104000            call dword ptr [0040103c]
'00635c96    898500ffffff            mov dword ptr [ebp+ffffff00], eax
'00635c9c    8bcf                    mov ecx, edi
'00635c9e    898df8feffff            mov dword ptr [ebp+fffffef8], ecx
'00635ca4    83ec10                  sub esp, 10
'00635ca7    8bd4                    mov edx, esp
'00635ca9    890a                    mov dword ptr [edx], ecx
'00635cab    8b8dfcfeffff            mov ecx, dword ptr [ebp+fffffefc]
'00635cb1    894a04                  mov dword ptr [edx+04], ecx
'00635cb4    894208                  mov dword ptr [edx+08], eax
'00635cb7    8b8504ffffff            mov eax, dword ptr [ebp+ffffff04]
'00635cbd    89420c                  mov dword ptr [edx+0c], eax
'00635cc0    6a01                    push 01
'00635cc2    6880000000              push 00000080
'00635cc7    8b4508                  mov eax, dword ptr [ebp+08]
'00635cca    8b08                    mov ecx, dword ptr [eax]
'00635ccc    50                      push eax
'00635ccd    ff91fc020000            call dword ptr [ecx+000002fc]
'00635cd3    50                      push eax
'00635cd4    8d5588                  lea edx, dword ptr [ebp-78]
'00635cd7    52                      push edx

' *** Reference to "__vbaObjSet"
'00635cd8    ff15b4104000            call dword ptr [004010b4]
Set var_131 = Me
'00635cde    50                      push eax

' *** Reference to "__vbaLateIdCall"
'00635cdf    ff1538104000            call dword ptr [00401038]
Call frmAcceder.Member_0x80(#NOT SUPPORTED#)
'00635ce5    8d4588                  lea eax, dword ptr [ebp-78]
'00635ce8    50                      push eax
'00635ce9    8d4d8c                  lea ecx, dword ptr [ebp-74]
'00635cec    51                      push ecx
'00635ced    8d5590                  lea edx, dword ptr [ebp-70]
'00635cf0    52                      push edx
'00635cf1    8d4594                  lea eax, dword ptr [ebp-6c]
'00635cf4    50                      push eax
'00635cf5    8d4d98                  lea ecx, dword ptr [ebp-68]
'00635cf8    51                      push ecx
'00635cf9    6a05                    push 05

' *** Reference to "__vbaFreeObjList"
'00635cfb    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_130, var_121, var_8, var_52, var_131)
'00635d01    8d95f8feffff            lea edx, dword ptr [ebp+fffffef8]
'00635d07    52                      push edx
'00635d08    8d8508ffffff            lea eax, dword ptr [ebp+ffffff08]
'00635d0e    50                      push eax
'00635d0f    8d8d18ffffff            lea ecx, dword ptr [ebp+ffffff18]
'00635d15    51                      push ecx
'00635d16    8d9528ffffff            lea edx, dword ptr [ebp+ffffff28]
'00635d1c    52                      push edx
'00635d1d    8d8538ffffff            lea eax, dword ptr [ebp+ffffff38]
'00635d23    50                      push eax
'00635d24    8d8d48ffffff            lea ecx, dword ptr [ebp+ffffff48]
'00635d2a    51                      push ecx
'00635d2b    8d9558ffffff            lea edx, dword ptr [ebp+ffffff58]
'00635d31    52                      push edx
'00635d32    8d8568ffffff            lea eax, dword ptr [ebp+ffffff68]
'00635d38    50                      push eax
'00635d39    8d8d78ffffff            lea ecx, dword ptr [ebp+ffffff78]
'00635d3f    51                      push ecx
'00635d40    6a09                    push 09

' *** Reference to "__vbaFreeVarList"
'00635d42    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_87, var_132, var_133, var_134, var_135, var_136, var_849, var_310, var_298)
'00635d48    83c45c                  add esp, 5c
'00635d4b    8b45c0                  mov eax, dword ptr [ebp-40]
'00635d4e    8b10                    mov edx, dword ptr [eax]
'00635d50    50                      push eax
'00635d51    ff92ec000000            call dword ptr [edx+000000ec]
'00635d57    dbe2                    fnclex
'00635d59    85c0                    test ax, ax
'00635d5b    0f8d1ee5ffff            jge 63427f
'00635d61    68ec000000              push 000000ec
'00635d66    6830314300              push 00433130
'00635d6b    8b4dc0                  mov ecx, dword ptr [ebp-40]
'00635d6e    51                      push ecx
'00635d6f    50                      push eax
'00635d70    ffd3                    call ebx
'00635d72    e908e5ffff              jmp 63427f

'ERROR: Two many next close:
Loop
'00635d77    50                      push eax
'00635d78    ff92c4000000            call dword ptr [edx+000000c4]
'00635d7e    dbe2                    fnclex
'00635d80    85c0                    test ax, ax
'00635d82    7d11                    jge 635d95
'00635d84    68c4000000              push 000000c4
'00635d89    6830314300              push 00433130
'00635d8e    8b4dc0                  mov ecx, dword ptr [ebp-40]
'00635d91    51                      push ecx
'00635d92    50                      push eax
'00635d93    ffd3                    call ebx

' *** Reference to "__vbaExitProc"
'00635d95    ff15a0104000            call dword ptr [004010a0]
'00635d9b    68435e6300              push 00635e43
'00635da0    e98b000000              jmp 635e30
'00635da5    8d559c                  lea edx, dword ptr [ebp-64]
'00635da8    52                      push edx
'00635da9    8d45a0                  lea eax, dword ptr [ebp-60]
'00635dac    50                      push eax
'00635dad    8d4da4                  lea ecx, dword ptr [ebp-5c]
'00635db0    51                      push ecx
'00635db1    8d55a8                  lea edx, dword ptr [ebp-58]
'00635db4    52                      push edx
'00635db5    8d45ac                  lea eax, dword ptr [ebp-54]
'00635db8    50                      push eax
'00635db9    8d4db0                  lea ecx, dword ptr [ebp-50]
'00635dbc    51                      push ecx
'00635dbd    8d55b4                  lea edx, dword ptr [ebp-4c]
'00635dc0    52                      push edx
'00635dc1    6a07                    push 07

' *** Reference to "__vbaFreeStrList"
'00635dc3    ff155c124000            call dword ptr [0040125c]
'#Cleanup( var_62, -4524, -4528, -4532, -4536, -4540, -4544)
'00635dc9    8d4588                  lea eax, dword ptr [ebp-78]
'00635dcc    50                      push eax
'00635dcd    8d4d8c                  lea ecx, dword ptr [ebp-74]
'00635dd0    51                      push ecx
'00635dd1    8d5590                  lea edx, dword ptr [ebp-70]
'00635dd4    52                      push edx
'00635dd5    8d4594                  lea eax, dword ptr [ebp-6c]
'00635dd8    50                      push eax
'00635dd9    8d4d98                  lea ecx, dword ptr [ebp-68]
'00635ddc    51                      push ecx
'00635ddd    6a05                    push 05

' *** Reference to "__vbaFreeObjList"
'00635ddf    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_130, var_121, var_8, var_52, var_131)
'00635de5    8d95f8feffff            lea edx, dword ptr [ebp+fffffef8]
'00635deb    52                      push edx
'00635dec    8d8508ffffff            lea eax, dword ptr [ebp+ffffff08]
'00635df2    50                      push eax
'00635df3    8d8d18ffffff            lea ecx, dword ptr [ebp+ffffff18]
'00635df9    51                      push ecx
'00635dfa    8d9528ffffff            lea edx, dword ptr [ebp+ffffff28]
'00635e00    52                      push edx
'00635e01    8d8538ffffff            lea eax, dword ptr [ebp+ffffff38]
'00635e07    50                      push eax
'00635e08    8d8d48ffffff            lea ecx, dword ptr [ebp+ffffff48]
'00635e0e    51                      push ecx
'00635e0f    8d9558ffffff            lea edx, dword ptr [ebp+ffffff58]
'00635e15    52                      push edx
'00635e16    8d8568ffffff            lea eax, dword ptr [ebp+ffffff68]
'00635e1c    50                      push eax
'00635e1d    8d8d78ffffff            lea ecx, dword ptr [ebp+ffffff78]
'00635e23    51                      push ecx
'00635e24    6a09                    push 09

' *** Reference to "__vbaFreeVarList"
'00635e26    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_87, var_132, var_133, var_134, var_135, var_136, var_849, var_310, var_298)
'00635e2c    83c460                  add esp, 60
'00635e2f    c3                      ret
'00635e30    8d4ddc                  lea ecx, dword ptr [ebp-24]

' *** Reference to "__vbaFreeStr"
'00635e33    ff150c134000            call dword ptr [0040130c]
'#Cleanup(var_39)
'00635e39    8d4dc0                  lea ecx, dword ptr [ebp-40]

' *** Reference to "__vbaFreeObj"
'00635e3c    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_5)
'00635e42    c3                      ret
'00635e43    8b4508                  mov eax, dword ptr [ebp+08]
'00635e46    8b10                    mov edx, dword ptr [eax]
'00635e48    50                      push eax
'00635e49    ff5208                  call dword ptr [edx+08]
'00635e4c    8b45f4                  mov eax, dword ptr [ebp-0c]
'00635e4f    8b4de4                  mov ecx, dword ptr [ebp-1c]
    'Reference to '__except_list'
'00635e52    64890d00000000          mov dword ptr fs:[00000000], ecx
'00635e59    5f                      pop edi
'00635e5a    5e                      pop esi
'00635e5b    5b                      pop ebx
'00635e5c    8be5                    mov esp, ebp
'00635e5e    5d                      pop ebp
'00635e5f    c20400                  ret 0004


End Function


Public Function AccedePerso(arg_0 As Unknow, arg_1 As Unknow, arg_2 As Unknow, arg_3 As Unknow, arg_4 As Unknow, arg_5 As Unknow, arg_6 As Unknow, arg_7 As Unknow, arg_8 As Unknow, arg_9 As Unknow, arg_A As Unknow, arg_B As Unknow, arg_C As Unknow, arg_D As Unknow, arg_E As Unknow, arg_F As Unknow, arg_10 As Unknow, arg_11 As Unknow, arg_12 As Unknow, arg_13 As Unknow, arg_14 As Unknow, arg_15 As Unknow, arg_16 As Unknow, arg_17 As Unknow, arg_18 As Unknow, arg_19 As Unknow, arg_1A As Unknow, arg_1B As Unknow, arg_1C As Unknow, arg_1D As Unknow, arg_1E As Unknow, arg_1F As Unknow, arg_20 As Unknow, arg_21 As Unknow, arg_22 As Unknow, arg_23 As Unknow, arg_24 As Unknow, arg_25 As Unknow, arg_26 As Unknow, arg_27 As Unknow, arg_28 As Unknow, arg_29 As Unknow, arg_2A As Unknow, arg_2B As Unknow, arg_2C As Unknow, arg_2D As Unknow, arg_2E As Unknow, arg_2F As Unknow, arg_30 As Unknow, arg_31 As Unknow, arg_32 As Unknow, arg_33 As Unknow, arg_34 As Unknow, arg_35 As Unknow, arg_36 As Unknow, arg_37 As Unknow, arg_38 As Unknow, arg_39 As Unknow, arg_3A As Unknow, arg_3B As Unknow, arg_3C As Unknow)
'00636870    55                      push ebp
'00636871    8bec                    mov ebp, esp
'00636873    83ec14                  sub esp, 14

' *** Reference to "__vbaExceptHandler"
'00636876    6866724000              push 00407266
'0063687b    64a100000000            mov ax, word ptr fs:[00000000]
'00636881    50                      push eax
    'Reference to '__except_list'
'00636882    64892500000000          mov dword ptr fs:[00000000], esp
'00636889    81ec64010000            sub esp, 00000164
'0063688f    53                      push ebx
'00636890    56                      push esi
'00636891    57                      push edi
'00636892    8965ec                  mov dword ptr [ebp-14], esp
'00636895    c745f060554000          mov dword ptr [ebp-10], 00405560
'0063689c    33ff                    xor edi, edi
'0063689e    897df4                  mov dword ptr [ebp-0c], edi
'006368a1    897df8                  mov dword ptr [ebp-08], edi
'006368a4    8b7508                  mov esi, dword ptr [ebp+08]
'006368a7    8b06                    mov eax, dword ptr [esi]
'006368a9    56                      push esi
'006368aa    ff5004                  call dword ptr [eax+04]
'006368ad    897de0                  mov dword ptr [ebp-20], edi
'006368b0    897ddc                  mov dword ptr [ebp-24], edi
'006368b3    897dd8                  mov dword ptr [ebp-28], edi
'006368b6    897dd4                  mov dword ptr [ebp-2c], edi
'006368b9    897dd0                  mov dword ptr [ebp-30], edi
'006368bc    897dcc                  mov dword ptr [ebp-34], edi
'006368bf    897dc8                  mov dword ptr [ebp-38], edi
'006368c2    897dc4                  mov dword ptr [ebp-3c], edi
'006368c5    897dc0                  mov dword ptr [ebp-40], edi
'006368c8    897db0                  mov dword ptr [ebp-50], edi
'006368cb    897da0                  mov dword ptr [ebp-60], edi
'006368ce    897d90                  mov dword ptr [ebp-70], edi
'006368d1    897d80                  mov dword ptr [ebp-80], edi
'006368d4    89bd70ffffff            mov dword ptr [ebp+ffffff70], edi
'006368da    89bd60ffffff            mov dword ptr [ebp+ffffff60], edi
'006368e0    89bd50ffffff            mov dword ptr [ebp+ffffff50], edi
'006368e6    89bd40ffffff            mov dword ptr [ebp+ffffff40], edi
'006368ec    89bd30ffffff            mov dword ptr [ebp+ffffff30], edi
'006368f2    89bd20ffffff            mov dword ptr [ebp+ffffff20], edi
'006368f8    89bd10ffffff            mov dword ptr [ebp+ffffff10], edi
'006368fe    89bd00ffffff            mov dword ptr [ebp+ffffff00], edi
'00636904    89bdf0feffff            mov dword ptr [ebp+fffffef0], edi
'0063690a    89bde0feffff            mov dword ptr [ebp+fffffee0], edi
'00636910    89bdc0feffff            mov dword ptr [ebp+fffffec0], edi
'00636916    89bdacfeffff            mov dword ptr [ebp+fffffeac], edi
'0063691c    89bda8feffff            mov dword ptr [ebp+fffffea8], edi
'00636922    66393d28b07200          cmp word ptr [0072b028], di
'00636929    7508                    jne 636933
'0063692b    6a01                    push 01

' *** Reference to "__vbaOnError"
'0063692d    ff15b0104000            call dword ptr [004010b0]
On Error Goto handler_0
'00636933    57                      push edi
'00636934    68a7000000              push 000000a7
'00636939    8b16                    mov edx, dword ptr [esi]
'0063693b    56                      push esi
'0063693c    ff92fc020000            call dword ptr [edx+000002fc]
'00636942    50                      push eax
'00636943    8d45c4                  lea eax, dword ptr [ebp-3c]
'00636946    50                      push eax

' *** Reference to "__vbaObjSet"
'00636947    8b1db4104000            mov ebx, dword ptr [004010b4]
'0063694d    ffd3                    call ebx
Set var_9 = Nothing
'0063694f    50                      push eax
'00636950    8d4db0                  lea ecx, dword ptr [ebp-50]
'00636953    51                      push ecx

' *** Reference to "__vbaLateIdCallLd"
'00636954    ff157c114000            call dword ptr [0040117c]
var_6 = var_9.UNK_-256 - 20_167
'0063695a    83c410                  add esp, 10
'0063695d    50                      push eax

' *** Reference to "__vbaI4Var"
'0063695e    ff157c124000            call dword ptr [0040127c]
'00636964    33d2                    xor edx, edx
var_num4 = Empty
'00636966    85c0                    test ax, ax
'00636968    0f9fc2                  setg dl
'0063696b    f7da                    neg edx
'0063696d    668995a4feffff          mov word ptr [ebp+fffffea4], dx
'00636974    8d4dc4                  lea ecx, dword ptr [ebp-3c]

' *** Reference to "__vbaFreeObj"
'00636977    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_9)
'0063697d    8d4db0                  lea ecx, dword ptr [ebp-50]

' *** Reference to "__vbaFreeVar"
'00636980    ff1528104000            call dword ptr [00401028]
'#Cleanup(var_6)
'00636986    6639bda4feffff          cmp word ptr [ebp+fffffea4], di
'0063698d    0f8475060000            je 637008

If (-(frmAcceder) <> 0) Then
'00636993    393d74b17200            cmp dword ptr [0072b174], edi
'00636999    7510                    jne 6369ab
'0063699b    6874b17200              push 0072b174
'006369a0    6890c04100              push 0041c090

' *** Reference to "__vbaNew2"
'006369a5    ff152c124000            call dword ptr [0040122c]
    Dim var_77 As New frmFichePerso
'006369ab    8b3d74b17200            mov edi, dword ptr [0072b174]
'006369b1    33c0                    xor eax, eax
    var_num1 = Empty
'006369b3    898508ffffff            mov dword ptr [ebp+ffffff08], eax
'006369b9    b903000000              mov ecx, 00000003
'006369be    898d00ffffff            mov dword ptr [ebp+ffffff00], ecx
'006369c4    898528ffffff            mov dword ptr [ebp+ffffff28], eax
'006369ca    898d20ffffff            mov dword ptr [ebp+ffffff20], ecx
'006369d0    83ec10                  sub esp, 10
'006369d3    8bd4                    mov edx, esp
'006369d5    890a                    mov dword ptr [edx], ecx
'006369d7    8b8d24ffffff            mov ecx, dword ptr [ebp+ffffff24]
'006369dd    894a04                  mov dword ptr [edx+04], ecx
'006369e0    894208                  mov dword ptr [edx+08], eax
'006369e3    8b852cffffff            mov eax, dword ptr [ebp+ffffff2c]
'006369e9    89420c                  mov dword ptr [edx+0c], eax
'006369ec    6a01                    push 01
'006369ee    68a8000000              push 000000a8
'006369f3    8b0e                    mov ecx, dword ptr [esi]
'006369f5    56                      push esi
'006369f6    ff91fc020000            call dword ptr [ecx+000002fc]
'006369fc    50                      push eax
'006369fd    8d55c4                  lea edx, dword ptr [ebp-3c]
'00636a00    52                      push edx
'00636a01    ffd3                    call ebx
    Set var_9 = Nothing
'00636a03    50                      push eax
'00636a04    8d45b0                  lea eax, dword ptr [ebp-50]
'00636a07    50                      push eax

' *** Reference to "__vbaLateIdCallLd"
'00636a08    ff157c114000            call dword ptr [0040117c]
    var_6 = var_9.UNK_-256 - 20_168
'00636a0e    83c420                  add esp, 20
'00636a11    50                      push eax

' *** Reference to "__vbaI4Var"
'00636a12    ff157c124000            call dword ptr [0040127c]
'00636a18    8985e8feffff            mov dword ptr [ebp+fffffee8], eax
'00636a1e    c785e0feffff03000000    mov dword ptr [ebp+fffffee0], 00000003
'00636a28    8b17                    mov edx, dword ptr [edi]
'00636a2a    83ec10                  sub esp, 10
'00636a2d    8bcc                    mov ecx, esp
'00636a2f    8b8500ffffff            mov eax, dword ptr [ebp+ffffff00]
'00636a35    8901                    mov dword ptr [ecx], eax
'00636a37    8b8504ffffff            mov eax, dword ptr [ebp+ffffff04]
'00636a3d    894104                  mov dword ptr [ecx+04], eax
'00636a40    8b8508ffffff            mov eax, dword ptr [ebp+ffffff08]
'00636a46    894108                  mov dword ptr [ecx+08], eax
'00636a49    8b850cffffff            mov eax, dword ptr [ebp+ffffff0c]
'00636a4f    89410c                  mov dword ptr [ecx+0c], eax
'00636a52    83ec10                  sub esp, 10
'00636a55    8bcc                    mov ecx, esp
'00636a57    8b85e0feffff            mov eax, dword ptr [ebp+fffffee0]
'00636a5d    8901                    mov dword ptr [ecx], eax
'00636a5f    8b85e4feffff            mov eax, dword ptr [ebp+fffffee4]
'00636a65    894104                  mov dword ptr [ecx+04], eax
'00636a68    8b85e8feffff            mov eax, dword ptr [ebp+fffffee8]
'00636a6e    894108                  mov dword ptr [ecx+08], eax
'00636a71    8b85ecfeffff            mov eax, dword ptr [ebp+fffffeec]
'00636a77    89410c                  mov dword ptr [ecx+0c], eax
'00636a7a    83ec10                  sub esp, 10
'00636a7d    8bcc                    mov ecx, esp
'00636a7f    b802000000              mov eax, 00000002
'00636a84    8901                    mov dword ptr [ecx], eax
'00636a86    8b85c4feffff            mov eax, dword ptr [ebp+fffffec4]
'00636a8c    894104                  mov dword ptr [ecx+04], eax
'00636a8f    33c0                    xor eax, eax
    var_num1 = Empty
'00636a91    894108                  mov dword ptr [ecx+08], eax
'00636a94    8b85ccfeffff            mov eax, dword ptr [ebp+fffffecc]
'00636a9a    89410c                  mov dword ptr [ecx+0c], eax
'00636a9d    6a03                    push 03
'00636a9f    689e000000              push 0000009e
'00636aa4    8b0e                    mov ecx, dword ptr [esi]
'00636aa6    56                      push esi
'00636aa7    899580feffff            mov dword ptr [ebp+fffffe80], edx
'00636aad    ff91fc020000            call dword ptr [ecx+000002fc]
'00636ab3    50                      push eax
'00636ab4    8d55c0                  lea edx, dword ptr [ebp-40]
'00636ab7    52                      push edx
'00636ab8    ffd3                    call ebx
    Set var_5 = Nothing
'00636aba    50                      push eax
'00636abb    8d45a0                  lea eax, dword ptr [ebp-60]
'00636abe    50                      push eax

' *** Reference to "__vbaLateIdCallLd"
'00636abf    ff157c114000            call dword ptr [0040117c]
    var_7 = var_5.UNK_-256 - 20_158
'00636ac5    83c440                  add esp, 40
'00636ac8    50                      push eax
'00636ac9    8d4de0                  lea ecx, dword ptr [ebp-20]
'00636acc    51                      push ecx

' *** Reference to "__vbaStrVarVal"
'00636acd    ff15fc114000            call dword ptr [004011fc]
'00636ad3    50                      push eax
'00636ad4    57                      push edi
'00636ad5    8b9580feffff            mov edx, dword ptr [ebp+fffffe80]
'00636adb    ff92fc060000            call dword ptr [edx+000006fc]
'00636ae1    dbe2                    fnclex
'00636ae3    85c0                    test ax, ax
'00636ae5    7d12                    jge 636af9
'00636ae7    68fc060000              push 000006fc
'00636aec    68e0234300              push 004323e0
'00636af1    57                      push edi
'00636af2    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00636af3    ff1580104000            call dword ptr [00401080]
'00636af9    8d4de0                  lea ecx, dword ptr [ebp-20]

' *** Reference to "__vbaFreeStr"
'00636afc    ff150c134000            call dword ptr [0040130c]
    '#Cleanup(var_3)
'00636b02    8d45c0                  lea eax, dword ptr [ebp-40]
'00636b05    50                      push eax
'00636b06    8d4dc4                  lea ecx, dword ptr [ebp-3c]
'00636b09    51                      push ecx
'00636b0a    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'00636b0c    ff154c104000            call dword ptr [0040104c]
    '#Cleanup( var_9, var_5)
'00636b12    8d55a0                  lea edx, dword ptr [ebp-60]
'00636b15    52                      push edx
'00636b16    8d45b0                  lea eax, dword ptr [ebp-50]
'00636b19    50                      push eax
'00636b1a    6a02                    push 02

' *** Reference to "__vbaFreeVarList"
'00636b1c    ff1540104000            call dword ptr [00401040]
    '#Cleanup( var_6, var_7)
'00636b22    83c418                  add esp, 18
'00636b25    a174b17200              mov ax, word ptr [0072b174]
'00636b2a    85c0                    test ax, ax
'00636b2c    7510                    jne 636b3e
'00636b2e    6874b17200              push 0072b174
'00636b33    6890c04100              push 0041c090

' *** Reference to "__vbaNew2"
'00636b38    ff152c124000            call dword ptr [0040122c]
    Set var_77 = New frmFichePerso
'00636b3e    8b3574b17200            mov esi, dword ptr [0072b174]
'00636b44    ba04000280              mov edx, 80020004
'00636b49    8bc2                    mov eax, edx
'00636b4b    898518ffffff            mov dword ptr [ebp+ffffff18], eax
'00636b51    b90a000000              mov ecx, 0000000a
'00636b56    898d10ffffff            mov dword ptr [ebp+ffffff10], ecx
'00636b5c    899528ffffff            mov dword ptr [ebp+ffffff28], edx
'00636b62    898d20ffffff            mov dword ptr [ebp+ffffff20], ecx
'00636b68    8b3e                    mov edi, dword ptr [esi]
'00636b6a    83ec10                  sub esp, 10
'00636b6d    8bdc                    mov ebx, esp
'00636b6f    890b                    mov dword ptr [ebx], ecx
'00636b71    8b8d14ffffff            mov ecx, dword ptr [ebp+ffffff14]
'00636b77    894b04                  mov dword ptr [ebx+04], ecx
'00636b7a    894308                  mov dword ptr [ebx+08], eax
'00636b7d    8b851cffffff            mov eax, dword ptr [ebp+ffffff1c]
'00636b83    89430c                  mov dword ptr [ebx+0c], eax
'00636b86    83ec10                  sub esp, 10
'00636b89    8bcc                    mov ecx, esp
'00636b8b    8b8520ffffff            mov eax, dword ptr [ebp+ffffff20]
'00636b91    8901                    mov dword ptr [ecx], eax
'00636b93    8b8524ffffff            mov eax, dword ptr [ebp+ffffff24]
'00636b99    894104                  mov dword ptr [ecx+04], eax
'00636b9c    895108                  mov dword ptr [ecx+08], edx
'00636b9f    8b952cffffff            mov edx, dword ptr [ebp+ffffff2c]
'00636ba5    89510c                  mov dword ptr [ecx+0c], edx
'00636ba8    56                      push esi
'00636ba9    ff97b0020000            call dword ptr [edi+000002b0]
'00636baf    dbe2                    fnclex
'00636bb1    85c0                    test ax, ax
'00636bb3    7d12                    jge 636bc7
'00636bb5    68b0020000              push 000002b0
'00636bba    68b0234300              push 004323b0
'00636bbf    56                      push esi
'00636bc0    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00636bc1    ff1580104000            call dword ptr [00401080]
'00636bc7    a174b17200              mov ax, word ptr [0072b174]
'00636bcc    85c0                    test ax, ax
'00636bce    0f854d030000            jne 636f21
'00636bd4    6874b17200              push 0072b174
'00636bd9    6890c04100              push 0041c090

' *** Reference to "__vbaNew2"
'00636bde    8b3d2c124000            mov edi, dword ptr [0040122c]
'00636be4    ffd7                    call edi
    Set var_77 = New frmFichePerso
'00636be6    e93c030000              jmp 636f27

' *** Reference to "rtcErrObj"
'00636beb    8b1d6c124000            mov ebx, dword ptr [0040126c]
'00636bf1    ffd3                    call ebx
'00636bf3    50                      push eax
'00636bf4    8d4dc4                  lea ecx, dword ptr [ebp-3c]
'00636bf7    51                      push ecx

' *** Reference to "__vbaObjSet"
'00636bf8    8b3db4104000            mov edi, dword ptr [004010b4]
'00636bfe    ffd7                    call edi
    Set var_9 = Err
'00636c00    8bf0                    mov esi, eax
'00636c02    8b16                    mov edx, dword ptr [esi]
'00636c04    8d45e0                  lea eax, dword ptr [ebp-20]
'00636c07    50                      push eax
'00636c08    56                      push esi
'00636c09    ff522c                  call dword ptr [edx+2c]
    var_3 = var_9.Description
'00636c0c    dbe2                    fnclex
'00636c0e    85c0                    test ax, ax
'00636c10    7d0f                    jge 636c21
'00636c12    6a2c                    push 2c
'00636c14    685c084300              push 0043085c
'00636c19    56                      push esi
'00636c1a    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00636c1b    ff1580104000            call dword ptr [00401080]
'00636c21    ffd3                    call ebx
'00636c23    50                      push eax
'00636c24    8d4dc0                  lea ecx, dword ptr [ebp-40]
'00636c27    51                      push ecx
'00636c28    ffd7                    call edi
    Set var_5 = Err
'00636c2a    8bf0                    mov esi, eax
'00636c2c    8b16                    mov edx, dword ptr [esi]
'00636c2e    8d85a8feffff            lea eax, dword ptr [ebp+fffffea8]
'00636c34    50                      push eax
'00636c35    56                      push esi
'00636c36    ff521c                  call dword ptr [edx+1c]
    var_519 = var_5.Number
'00636c39    dbe2                    fnclex
'00636c3b    85c0                    test ax, ax
'00636c3d    7d0f                    jge 636c4e
'00636c3f    6a1c                    push 1c
'00636c41    685c084300              push 0043085c
'00636c46    56                      push esi
'00636c47    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00636c48    ff1580104000            call dword ptr [00401080]
'00636c4e    ba04000280              mov edx, 80020004
'00636c53    895598                  mov dword ptr [ebp-68], edx
'00636c56    b90a000000              mov ecx, 0000000a
'00636c5b    894d90                  mov dword ptr [ebp-70], ecx
'00636c5e    8955a8                  mov dword ptr [ebp-58], edx
'00636c61    894da0                  mov dword ptr [ebp-60], ecx
'00636c64    c78528ffffff24b07200    mov dword ptr [ebp+ffffff28], 0072b024
'00636c6e    c78520ffffff08400000    mov dword ptr [ebp+ffffff20], 00004008
'00636c78    6814084300              push 00430814
'00636c7d    8b4de0                  mov ecx, dword ptr [ebp-20]
'00636c80    51                      push ecx

' *** Reference to "__vbaStrCat"
'00636c81    8b3570104000            mov esi, dword ptr [00401070]
'00636c87    ffd6                    call esi
    var_49 = ("L'erreur suivante s'est produite : ") & (var_3)
'00636c89    8bd0                    mov edx, eax
'00636c8b    8d4ddc                  lea ecx, dword ptr [ebp-24]

' *** Reference to "__vbaStrMove"
'00636c8e    8b3dd0124000            mov edi, dword ptr [004012d0]
'00636c94    ffd7                    call edi
'00636c96    50                      push eax
'00636c97    6870084300              push 00430870
'00636c9c    ffd6                    call esi
    var_2461 = (var_49) & (vbCrLf)
'00636c9e    8bd0                    mov edx, eax
'00636ca0    8d4dd8                  lea ecx, dword ptr [ebp-28]
'00636ca3    ffd7                    call edi
'00636ca5    50                      push eax
'00636ca6    6870084300              push 00430870
'00636cab    ffd6                    call esi
    var_870 = (var_2461) & (vbCrLf)
'00636cad    8bd0                    mov edx, eax
'00636caf    8d4dd4                  lea ecx, dword ptr [ebp-2c]
'00636cb2    ffd7                    call edi
'00636cb4    50                      push eax
'00636cb5    6880084300              push 00430880
'00636cba    ffd6                    call esi
    var_64 = (var_870) & (" numéro d'erreur (")
'00636cbc    8bd0                    mov edx, eax
'00636cbe    8d4dd0                  lea ecx, dword ptr [ebp-30]
'00636cc1    ffd7                    call edi
'00636cc3    50                      push eax
'00636cc4    8b95a8feffff            mov edx, dword ptr [ebp+fffffea8]
'00636cca    52                      push edx

' *** Reference to "__vbaStrI4"
'00636ccb    ff1520104000            call dword ptr [00401020]
'00636cd1    8bd0                    mov edx, eax
'00636cd3    8d4dcc                  lea ecx, dword ptr [ebp-34]
'00636cd6    ffd7                    call edi
'00636cd8    50                      push eax
'00636cd9    ffd6                    call esi
    var_871 = (var_64) & (CStr(var_519))
'00636cdb    8bd0                    mov edx, eax
'00636cdd    8d4dc8                  lea ecx, dword ptr [ebp-38]
'00636ce0    ffd7                    call edi
'00636ce2    50                      push eax
'00636ce3    68ac084300              push 004308ac
'00636ce8    ffd6                    call esi
    var_79 = (var_871) & (" )")
'00636cea    8945b8                  mov dword ptr [ebp-48], eax
'00636ced    bf08000000              mov edi, 00000008
'00636cf2    897db0                  mov dword ptr [ebp-50], edi
'00636cf5    8d4590                  lea eax, dword ptr [ebp-70]
'00636cf8    50                      push eax
'00636cf9    8d4da0                  lea ecx, dword ptr [ebp-60]
'00636cfc    51                      push ecx
'00636cfd    8d9520ffffff            lea edx, dword ptr [ebp+ffffff20]
'00636d03    52                      push edx
'00636d04    6a10                    push 10
'00636d06    8d45b0                  lea eax, dword ptr [ebp-50]
'00636d09    50                      push eax

' *** Reference to "rtcMsgBox"
'00636d0a    ff15b8104000            call dword ptr [004010b8]
    var_66 = MsgBox(var_79, 16, vbNullString)
'00636d10    8d4dc8                  lea ecx, dword ptr [ebp-38]
'00636d13    51                      push ecx
'00636d14    8d55cc                  lea edx, dword ptr [ebp-34]
'00636d17    52                      push edx
'00636d18    8d45d0                  lea eax, dword ptr [ebp-30]
'00636d1b    50                      push eax
'00636d1c    8d4dd4                  lea ecx, dword ptr [ebp-2c]
'00636d1f    51                      push ecx
'00636d20    8d55d8                  lea edx, dword ptr [ebp-28]
'00636d23    52                      push edx
'00636d24    8d45dc                  lea eax, dword ptr [ebp-24]
'00636d27    50                      push eax
'00636d28    8d4de0                  lea ecx, dword ptr [ebp-20]
'00636d2b    51                      push ecx
'00636d2c    6a07                    push 07

' *** Reference to "__vbaFreeStrList"
'00636d2e    ff155c124000            call dword ptr [0040125c]
    '#Cleanup( , -4604, -4608, -4612, -4616, -4620, -4624)
'00636d34    8d55c0                  lea edx, dword ptr [ebp-40]
'00636d37    52                      push edx
'00636d38    8d45c4                  lea eax, dword ptr [ebp-3c]
'00636d3b    50                      push eax
'00636d3c    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'00636d3e    ff154c104000            call dword ptr [0040104c]
    '#Cleanup( var_9, var_5)
'00636d44    8d4d90                  lea ecx, dword ptr [ebp-70]
'00636d47    51                      push ecx
'00636d48    8d55a0                  lea edx, dword ptr [ebp-60]
'00636d4b    52                      push edx
'00636d4c    8d45b0                  lea eax, dword ptr [ebp-50]
'00636d4f    50                      push eax
'00636d50    6a03                    push 03

' *** Reference to "__vbaFreeVarList"
'00636d52    ff1540104000            call dword ptr [00401040]
    '#Cleanup( var_6, var_7, var_8)
'00636d58    83c43c                  add esp, 3c
'00636d5b    8d4db0                  lea ecx, dword ptr [ebp-50]
'00636d5e    51                      push ecx

' *** Reference to "rtcGetPresentDate"
'00636d5f    ff15f4124000            call dword ptr [004012f4]
    var_79 = Now()
'00636d65    c78528ffffffb8084300    mov dword ptr [ebp+ffffff28], 004308b8
'00636d6f    89bd20ffffff            mov dword ptr [ebp+ffffff20], edi
'00636d75    8d9520ffffff            lea edx, dword ptr [ebp+ffffff20]
'00636d7b    8d4da0                  lea ecx, dword ptr [ebp-60]

' *** Reference to "__vbaVarDup"
'00636d7e    ff1598124000            call dword ptr [00401298]
    var_7 = ("dd/MM/yyyy hh:mm:ss")
'00636d84    6a01                    push 01
'00636d86    6a01                    push 01
'00636d88    8d55a0                  lea edx, dword ptr [ebp-60]
'00636d8b    52                      push edx
'00636d8c    8d45b0                  lea eax, dword ptr [ebp-50]
'00636d8f    50                      push eax
'00636d90    8d4d90                  lea ecx, dword ptr [ebp-70]
'00636d93    51                      push ecx

' *** Reference to "rtcVarFromFormatVar"
'00636d94    ff1574104000            call dword ptr [00401074]
'00636d9a    ffd3                    call ebx
'00636d9c    50                      push eax
'00636d9d    8d55c4                  lea edx, dword ptr [ebp-3c]
'00636da0    52                      push edx

' *** Reference to "__vbaObjSet"
'00636da1    ff15b4104000            call dword ptr [004010b4]
    Set var_9 = Err
'00636da7    8bf0                    mov esi, eax
'00636da9    8b06                    mov eax, dword ptr [esi]
'00636dab    8d8da8feffff            lea ecx, dword ptr [ebp+fffffea8]
'00636db1    51                      push ecx
'00636db2    56                      push esi
'00636db3    ff501c                  call dword ptr [eax+1c]
    var_519 = var_9.Number
'00636db6    dbe2                    fnclex
'00636db8    85c0                    test ax, ax
'00636dba    7d0f                    jge 636dcb
'00636dbc    6a1c                    push 1c
'00636dbe    685c084300              push 0043085c
'00636dc3    56                      push esi
'00636dc4    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00636dc5    ff1580104000            call dword ptr [00401080]
'00636dcb    ffd3                    call ebx
'00636dcd    50                      push eax
'00636dce    8d55c0                  lea edx, dword ptr [ebp-40]
'00636dd1    52                      push edx

' *** Reference to "__vbaObjSet"
'00636dd2    ff15b4104000            call dword ptr [004010b4]
    Set var_5 = Err
'00636dd8    8bf0                    mov esi, eax
'00636dda    8b06                    mov eax, dword ptr [esi]
'00636ddc    8d4de0                  lea ecx, dword ptr [ebp-20]
'00636ddf    51                      push ecx
'00636de0    56                      push esi
'00636de1    ff502c                  call dword ptr [eax+2c]
    var_3 = var_5.Description
'00636de4    dbe2                    fnclex
'00636de6    85c0                    test ax, ax
'00636de8    7d0f                    jge 636df9
'00636dea    6a2c                    push 2c
'00636dec    685c084300              push 0043085c
'00636df1    56                      push esi
'00636df2    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00636df3    ff1580104000            call dword ptr [00401080]
'00636df9    c78518ffffffe4084300    mov dword ptr [ebp+ffffff18], 004308e4
'00636e03    89bd10ffffff            mov dword ptr [ebp+ffffff10], edi
'00636e09    8b95a8feffff            mov edx, dword ptr [ebp+fffffea8]
'00636e0f    899508ffffff            mov dword ptr [ebp+ffffff08], edx
'00636e15    c78500ffffff03000000    mov dword ptr [ebp+ffffff00], 00000003
'00636e1f    c785f8feffff08094300    mov dword ptr [ebp+fffffef8], 00430908
'00636e29    89bdf0feffff            mov dword ptr [ebp+fffffef0], edi
'00636e2f    8b45e0                  mov eax, dword ptr [ebp-20]
'00636e32    c745e000000000          mov dword ptr [ebp-20], 00000000
'00636e39    898558ffffff            mov dword ptr [ebp+ffffff58], eax
'00636e3f    89bd50ffffff            mov dword ptr [ebp+ffffff50], edi
'00636e45    c785e8feffff18de4400    mov dword ptr [ebp+fffffee8], 0044de18
'00636e4f    89bde0feffff            mov dword ptr [ebp+fffffee0], edi
'00636e55    8d4590                  lea eax, dword ptr [ebp-70]
'00636e58    50                      push eax
'00636e59    8d8d10ffffff            lea ecx, dword ptr [ebp+ffffff10]
'00636e5f    51                      push ecx
'00636e60    8d5580                  lea edx, dword ptr [ebp-80]
'00636e63    52                      push edx

' *** Reference to "__vbaVarCat"
'00636e64    8b3508124000            mov esi, dword ptr [00401208]
'00636e6a    ffd6                    call esi
'00636e6c    50                      push eax
'00636e6d    8d8500ffffff            lea eax, dword ptr [ebp+ffffff00]
'00636e73    50                      push eax
'00636e74    8d8d70ffffff            lea ecx, dword ptr [ebp+ffffff70]
'00636e7a    51                      push ecx
'00636e7b    ffd6                    call esi
'00636e7d    50                      push eax
'00636e7e    8d95f0feffff            lea edx, dword ptr [ebp+fffffef0]
'00636e84    52                      push edx
'00636e85    8d8560ffffff            lea eax, dword ptr [ebp+ffffff60]
'00636e8b    50                      push eax
'00636e8c    ffd6                    call esi
'00636e8e    50                      push eax
'00636e8f    8d8d50ffffff            lea ecx, dword ptr [ebp+ffffff50]
'00636e95    51                      push ecx
'00636e96    8d9540ffffff            lea edx, dword ptr [ebp+ffffff40]
'00636e9c    52                      push edx
'00636e9d    ffd6                    call esi
'00636e9f    50                      push eax
'00636ea0    8d85e0feffff            lea eax, dword ptr [ebp+fffffee0]
'00636ea6    50                      push eax
'00636ea7    8d8d30ffffff            lea ecx, dword ptr [ebp+ffffff30]
'00636ead    51                      push ecx
'00636eae    ffd6                    call esi
'00636eb0    50                      push eax
'00636eb1    33d2                    xor edx, edx
'00636eb3    668b152ab07200          mov dx, word ptr [0072b02a]
'00636eba    52                      push edx
'00636ebb    6884094300              push 00430984

' *** Reference to "__vbaPrintFile"
'00636ec0    ff15b8114000            call dword ptr [004011b8]
    Print #0, 
'00636ec6    8d45c0                  lea eax, dword ptr [ebp-40]
'00636ec9    50                      push eax
'00636eca    8d4dc4                  lea ecx, dword ptr [ebp-3c]
'00636ecd    51                      push ecx
'00636ece    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'00636ed0    ff154c104000            call dword ptr [0040104c]
    '#Cleanup( var_9, var_5)
'00636ed6    8d9530ffffff            lea edx, dword ptr [ebp+ffffff30]
'00636edc    52                      push edx
'00636edd    8d8540ffffff            lea eax, dword ptr [ebp+ffffff40]
'00636ee3    50                      push eax
'00636ee4    8d8d50ffffff            lea ecx, dword ptr [ebp+ffffff50]
'00636eea    51                      push ecx
'00636eeb    8d9560ffffff            lea edx, dword ptr [ebp+ffffff60]
'00636ef1    52                      push edx
'00636ef2    8d8570ffffff            lea eax, dword ptr [ebp+ffffff70]
'00636ef8    50                      push eax
'00636ef9    8d4d80                  lea ecx, dword ptr [ebp-80]
'00636efc    51                      push ecx
'00636efd    8d5590                  lea edx, dword ptr [ebp-70]
'00636f00    52                      push edx
'00636f01    8d45a0                  lea eax, dword ptr [ebp-60]
'00636f04    50                      push eax
'00636f05    8d4db0                  lea ecx, dword ptr [ebp-50]
'00636f08    51                      push ecx
'00636f09    6a09                    push 09

' *** Reference to "__vbaFreeVarList"
'00636f0b    ff1540104000            call dword ptr [00401040]
    '#Cleanup( var_6, var_7, var_8, var_18, var_19, var_20, var_21, var_22, var_23)
'00636f11    83c440                  add esp, 40
'00636f14    6a00                    push 00

' *** Reference to "__vbaResume"
'00636f16    ff1568104000            call dword ptr [00401068]
'00636f1c    e9e7000000              jmp 637008
    Resume handler_637008

' *** Reference to "__vbaNew2"
'00636f21    8b3d2c124000            mov edi, dword ptr [0040122c]
'00636f27    8b3574b17200            mov esi, dword ptr [0072b174]
'00636f2d    33c0                    xor eax, eax
'00636f2f    898528ffffff            mov dword ptr [ebp+ffffff28], eax
'00636f35    b902000000              mov ecx, 00000002
'00636f3a    898d20ffffff            mov dword ptr [ebp+ffffff20], ecx
'00636f40    8b16                    mov edx, dword ptr [esi]
'00636f42    83ec10                  sub esp, 10
'00636f45    8bdc                    mov ebx, esp
'00636f47    890b                    mov dword ptr [ebx], ecx
'00636f49    8b8d24ffffff            mov ecx, dword ptr [ebp+ffffff24]
'00636f4f    894b04                  mov dword ptr [ebx+04], ecx
'00636f52    894308                  mov dword ptr [ebx+08], eax
'00636f55    8b852cffffff            mov eax, dword ptr [ebp+ffffff2c]
'00636f5b    89430c                  mov dword ptr [ebx+0c], eax
'00636f5e    56                      push esi
'00636f5f    ff92ac020000            call dword ptr [edx+000002ac]
'00636f65    dbe2                    fnclex
'00636f67    85c0                    test ax, ax
'00636f69    7d16                    jge 636f81
'00636f6b    68ac020000              push 000002ac
'00636f70    68b0234300              push 004323b0
'00636f75    56                      push esi
'00636f76    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00636f77    8b1d80104000            mov ebx, dword ptr [00401080]
'00636f7d    ffd3                    call ebx
'00636f7f    eb06                    jmp 636f87

' *** Reference to "__vbaHresultCheckObj"
'00636f81    8b1d80104000            mov ebx, dword ptr [00401080]
'00636f87    a174b17200              mov ax, word ptr [0072b174]
'00636f8c    85c0                    test ax, ax
'00636f8e    750c                    jne 636f9c
'00636f90    6874b17200              push 0072b174
'00636f95    6890c04100              push 0041c090
'00636f9a    ffd7                    call edi
    Set var_77 = New frmFichePerso
'00636f9c    8b3574b17200            mov esi, dword ptr [0072b174]
'00636fa2    8b0e                    mov ecx, dword ptr [esi]
'00636fa4    56                      push esi
'00636fa5    ff9168070000            call dword ptr [ecx+00000768]
'00636fab    dbe2                    fnclex
'00636fad    85c0                    test ax, ax
'00636faf    7d0e                    jge 636fbf
'00636fb1    6868070000              push 00000768
'00636fb6    68e0234300              push 004323e0
'00636fbb    56                      push esi
'00636fbc    50                      push eax
'00636fbd    ffd3                    call ebx
'00636fbf    a174b17200              mov ax, word ptr [0072b174]
'00636fc4    85c0                    test ax, ax
'00636fc6    750c                    jne 636fd4
'00636fc8    6874b17200              push 0072b174
'00636fcd    6890c04100              push 0041c090
'00636fd2    ffd7                    call edi
    Set var_77 = New frmFichePerso
'00636fd4    8b3574b17200            mov esi, dword ptr [0072b174]
'00636fda    c785acfeffff01000000    mov dword ptr [ebp+fffffeac], 00000001
'00636fe4    8b16                    mov edx, dword ptr [esi]
'00636fe6    8d85acfeffff            lea eax, dword ptr [ebp+fffffeac]
'00636fec    50                      push eax
'00636fed    56                      push esi
'00636fee    ff9264070000            call dword ptr [edx+00000764]
'00636ff4    dbe2                    fnclex
'00636ff6    85c0                    test ax, ax
'00636ff8    7d0e                    jge 637008
'00636ffa    6864070000              push 00000764
'00636fff    68e0234300              push 004323e0
'00637004    56                      push esi
'00637005    50                      push eax
'00637006    ffd3                    call ebx
    
End If

' *** Reference to "__vbaExitProc"
'00637008    ff15a0104000            call dword ptr [004010a0]
'0063700e    6889706300              push 00637089
'00637013    eb73                    jmp 637088
'00637015    8d4dc8                  lea ecx, dword ptr [ebp-38]
'00637018    51                      push ecx
'00637019    8d55cc                  lea edx, dword ptr [ebp-34]
'0063701c    52                      push edx
'0063701d    8d45d0                  lea eax, dword ptr [ebp-30]
'00637020    50                      push eax
'00637021    8d4dd4                  lea ecx, dword ptr [ebp-2c]
'00637024    51                      push ecx
'00637025    8d55d8                  lea edx, dword ptr [ebp-28]
'00637028    52                      push edx
'00637029    8d45dc                  lea eax, dword ptr [ebp-24]
'0063702c    50                      push eax
'0063702d    8d4de0                  lea ecx, dword ptr [ebp-20]
'00637030    51                      push ecx
'00637031    6a07                    push 07

' *** Reference to "__vbaFreeStrList"
'00637033    ff155c124000            call dword ptr [0040125c]
'#Cleanup( var_3, -4604, -4608, -4612, -4616, -4620, -4624)
'00637039    8d55c0                  lea edx, dword ptr [ebp-40]
'0063703c    52                      push edx
'0063703d    8d45c4                  lea eax, dword ptr [ebp-3c]
'00637040    50                      push eax
'00637041    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'00637043    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_9, var_5)
'00637049    8d8d30ffffff            lea ecx, dword ptr [ebp+ffffff30]
'0063704f    51                      push ecx
'00637050    8d9540ffffff            lea edx, dword ptr [ebp+ffffff40]
'00637056    52                      push edx
'00637057    8d8550ffffff            lea eax, dword ptr [ebp+ffffff50]
'0063705d    50                      push eax
'0063705e    8d8d60ffffff            lea ecx, dword ptr [ebp+ffffff60]
'00637064    51                      push ecx
'00637065    8d9570ffffff            lea edx, dword ptr [ebp+ffffff70]
'0063706b    52                      push edx
'0063706c    8d4580                  lea eax, dword ptr [ebp-80]
'0063706f    50                      push eax
'00637070    8d4d90                  lea ecx, dword ptr [ebp-70]
'00637073    51                      push ecx
'00637074    8d55a0                  lea edx, dword ptr [ebp-60]
'00637077    52                      push edx
'00637078    8d45b0                  lea eax, dword ptr [ebp-50]
'0063707b    50                      push eax
'0063707c    6a09                    push 09

' *** Reference to "__vbaFreeVarList"
'0063707e    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_6, var_7, var_8, var_18, var_19, var_20, var_21, var_22, var_23)
'00637084    83c454                  add esp, 54
'00637087    c3                      ret
'00637088    c3                      ret
'00637089    8b4508                  mov eax, dword ptr [ebp+08]
'0063708c    8b08                    mov ecx, dword ptr [eax]
'0063708e    50                      push eax
'0063708f    ff5108                  call dword ptr [ecx+08]
'00637092    8b45f4                  mov eax, dword ptr [ebp-0c]
'00637095    8b4de4                  mov ecx, dword ptr [ebp-1c]
    'Reference to '__except_list'
'00637098    64890d00000000          mov dword ptr fs:[00000000], ecx
'0063709f    5f                      pop edi
'006370a0    5e                      pop esi
'006370a1    5b                      pop ebx
'006370a2    8be5                    mov esp, ebp
'006370a4    5d                      pop ebp
'006370a5    c20400                  ret 0004


End Function


'Event for vsnom
Private Sub vsnom_Event9()
'006367e0    55                      push ebp
'006367e1    8bec                    mov ebp, esp
'006367e3    83ec0c                  sub esp, 0c

' *** Reference to "__vbaExceptHandler"
'006367e6    6866724000              push 00407266
'006367eb    64a100000000            mov ax, word ptr fs:[00000000]
'006367f1    50                      push eax
    'Reference to '__except_list'
'006367f2    64892500000000          mov dword ptr fs:[00000000], esp
'006367f9    83ec0c                  sub esp, 0c
'006367fc    53                      push ebx
'006367fd    56                      push esi
'006367fe    57                      push edi
'006367ff    8965f4                  mov dword ptr [ebp-0c], esp
'00636802    c745f858554000          mov dword ptr [ebp-08], 00405558
'00636809    8b7508                  mov esi, dword ptr [ebp+08]
'0063680c    8bc6                    mov eax, esi
'0063680e    83e001                  and eax, 01
'00636811    8945fc                  mov dword ptr [ebp-04], eax
'00636814    83e6fe                  and esi, fffffffe
'00636817    8b0e                    mov ecx, dword ptr [esi]
'00636819    56                      push esi
'0063681a    897508                  mov dword ptr [ebp+08], esi
'0063681d    ff5104                  call dword ptr [ecx+04]
'00636820    8b16                    mov edx, dword ptr [esi]
'00636822    56                      push esi
'00636823    ff92fc060000            call dword ptr [edx+000006fc]
'00636829    85c0                    test ax, ax
'0063682b    7d12                    jge 63683f
'0063682d    68fc060000              push 000006fc
'00636832    6840074300              push 00430740
'00636837    56                      push esi
'00636838    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00636839    ff1580104000            call dword ptr [00401080]
'0063683f    c745fc00000000          mov dword ptr [ebp-04], 00000000
'00636846    8b4508                  mov eax, dword ptr [ebp+08]
'00636849    8b08                    mov ecx, dword ptr [eax]
'0063684b    50                      push eax
'0063684c    ff5108                  call dword ptr [ecx+08]
'0063684f    8b45fc                  mov eax, dword ptr [ebp-04]
'00636852    8b4dec                  mov ecx, dword ptr [ebp-14]
'00636855    5f                      pop edi
'00636856    5e                      pop esi
    'Reference to '__except_list'
'00636857    64890d00000000          mov dword ptr fs:[00000000], ecx
'0063685e    5b                      pop ebx
'0063685f    8be5                    mov esp, ebp
'00636861    5d                      pop ebp
'00636862    c20400                  ret 0004


End Sub


'Event for Form
Private Sub Form_Load()
'00635e70    55                      push ebp
'00635e71    8bec                    mov ebp, esp
'00635e73    83ec14                  sub esp, 14

' *** Reference to "__vbaExceptHandler"
'00635e76    6866724000              push 00407266
'00635e7b    64a100000000            mov ax, word ptr fs:[00000000]
'00635e81    50                      push eax
    'Reference to '__except_list'
'00635e82    64892500000000          mov dword ptr fs:[00000000], esp
'00635e89    81ec2c010000            sub esp, 0000012c
'00635e8f    53                      push ebx
'00635e90    56                      push esi
'00635e91    57                      push edi
'00635e92    8965ec                  mov dword ptr [ebp-14], esp
'00635e95    c745f010554000          mov dword ptr [ebp-10], 00405510
'00635e9c    8b7508                  mov esi, dword ptr [ebp+08]
'00635e9f    8bc6                    mov eax, esi
'00635ea1    83e001                  and eax, 01
'00635ea4    8945f4                  mov dword ptr [ebp-0c], eax
'00635ea7    83e6fe                  and esi, fffffffe
'00635eaa    897508                  mov dword ptr [ebp+08], esi
'00635ead    33db                    xor ebx, ebx
'00635eaf    895df8                  mov dword ptr [ebp-08], ebx
'00635eb2    8b0e                    mov ecx, dword ptr [esi]
'00635eb4    56                      push esi
'00635eb5    ff5104                  call dword ptr [ecx+04]
'00635eb8    895de0                  mov dword ptr [ebp-20], ebx
'00635ebb    895ddc                  mov dword ptr [ebp-24], ebx
'00635ebe    895dd8                  mov dword ptr [ebp-28], ebx
'00635ec1    895dd4                  mov dword ptr [ebp-2c], ebx
'00635ec4    895dd0                  mov dword ptr [ebp-30], ebx
'00635ec7    895dcc                  mov dword ptr [ebp-34], ebx
'00635eca    895dc8                  mov dword ptr [ebp-38], ebx
'00635ecd    895dc4                  mov dword ptr [ebp-3c], ebx
'00635ed0    895dc0                  mov dword ptr [ebp-40], ebx
'00635ed3    895db0                  mov dword ptr [ebp-50], ebx
'00635ed6    895da0                  mov dword ptr [ebp-60], ebx
'00635ed9    895d90                  mov dword ptr [ebp-70], ebx
'00635edc    895d80                  mov dword ptr [ebp-80], ebx
'00635edf    899d70ffffff            mov dword ptr [ebp+ffffff70], ebx
'00635ee5    899d60ffffff            mov dword ptr [ebp+ffffff60], ebx
'00635eeb    899d50ffffff            mov dword ptr [ebp+ffffff50], ebx
'00635ef1    899d40ffffff            mov dword ptr [ebp+ffffff40], ebx
'00635ef7    899d30ffffff            mov dword ptr [ebp+ffffff30], ebx
'00635efd    899d20ffffff            mov dword ptr [ebp+ffffff20], ebx
'00635f03    899d10ffffff            mov dword ptr [ebp+ffffff10], ebx
'00635f09    899d00ffffff            mov dword ptr [ebp+ffffff00], ebx
'00635f0f    899df0feffff            mov dword ptr [ebp+fffffef0], ebx
'00635f15    899de0feffff            mov dword ptr [ebp+fffffee0], ebx
'00635f1b    899ddcfeffff            mov dword ptr [ebp+fffffedc], ebx
'00635f21    66391d28b07200          cmp word ptr [0072b028], bx
'00635f28    7508                    jne 635f32
'00635f2a    6a01                    push 01

' *** Reference to "__vbaOnError"
'00635f2c    ff15b0104000            call dword ptr [004010b0]
On Error Goto handler_0
'00635f32    391d5cb07200            cmp dword ptr [0072b05c], ebx
'00635f38    7510                    jne 635f4a
'00635f3a    685cb07200              push 0072b05c
'00635f3f    68347b4000              push 00407b34

' *** Reference to "__vbaNew2"
'00635f44    ff152c124000            call dword ptr [0040122c]
Dim var_24 As New frmAcceder
'00635f4a    a15cb07200              mov ax, word ptr [0072b05c]
'00635f4f    8985d0feffff            mov dword ptr [ebp+fffffed0], eax
'00635f55    391d10b07200            cmp dword ptr [0072b010], ebx
'00635f5b    7510                    jne 635f6d
'00635f5d    6810b07200              push 0072b010
'00635f62    68b0e54000              push 0040e5b0

' *** Reference to "__vbaNew2"
'00635f67    ff152c124000            call dword ptr [0040122c]
Dim var_60 As New frmMain
'00635f6d    8b3d10b07200            mov edi, dword ptr [0072b010]
'00635f73    8b0f                    mov ecx, dword ptr [edi]
'00635f75    8d95dcfeffff            lea edx, dword ptr [ebp+fffffedc]
'00635f7b    52                      push edx
'00635f7c    57                      push edi
'00635f7d    ff9180000000            call dword ptr [ecx+00000080]
'00635f83    dbe2                    fnclex
'00635f85    3bc3                    cmp eax, ebx
'00635f87    7d12                    jge 635f9b

If (var_24 < 0) Then
'00635f89    6880000000              push 00000080
'00635f8e    684cfd4200              push 0042fd4c
'00635f93    57                      push edi
'00635f94    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00635f95    ff1580104000            call dword ptr [00401080]
    
End If
'00635f9b    8bbdd0feffff            mov edi, dword ptr [ebp+fffffed0]
'00635fa1    8b0f                    mov ecx, dword ptr [edi]
'00635fa3    d985dcfeffff            fld dword ptr [ebp+fffffedc]
'00635fa9    d82538554000            fsub dword ptr [00405538]
'00635faf    dfe0                    fnstsw ax
'00635fb1    a80d                    test al, 0d
'00635fb3    0f85eb050000            jne 6365a4
'00635fb9    51                      push ecx
'00635fba    d91c24                  fstp dword ptr [esp]
'var_534 = (00)
'00635fbd    57                      push edi
'00635fbe    ff5174                  call dword ptr [ecx+74]
'00635fc1    dbe2                    fnclex
'00635fc3    3bc3                    cmp eax, ebx
'00635fc5    7d0f                    jge 635fd6
'00635fc7    6a74                    push 74
'00635fc9    6810074300              push 00430710
'00635fce    57                      push edi
'00635fcf    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00635fd0    ff1580104000            call dword ptr [00401080]
'00635fd6    391d5cb07200            cmp dword ptr [0072b05c], ebx
'00635fdc    7510                    jne 635fee

If (var_24 = 0) Then
'00635fde    685cb07200              push 0072b05c
'00635fe3    68347b4000              push 00407b34

' *** Reference to "__vbaNew2"
'00635fe8    ff152c124000            call dword ptr [0040122c]
    Set var_24 = New frmAcceder
End If
'00635fee    8b1d5cb07200            mov ebx, dword ptr [0072b05c]
'00635ff4    a110b07200              mov ax, word ptr [0072b010]
'00635ff9    85c0                    test ax, ax
'00635ffb    7510                    jne 63600d
'00635ffd    6810b07200              push 0072b010
'00636002    68b0e54000              push 0040e5b0

' *** Reference to "__vbaNew2"
'00636007    ff152c124000            call dword ptr [0040122c]
Set var_60 = New frmMain
'0063600d    8b3d10b07200            mov edi, dword ptr [0072b010]
'00636013    8b17                    mov edx, dword ptr [edi]
'00636015    8d85dcfeffff            lea eax, dword ptr [ebp+fffffedc]
'0063601b    50                      push eax
'0063601c    57                      push edi
'0063601d    ff9288000000            call dword ptr [edx+00000088]
'00636023    dbe2                    fnclex
'00636025    85c0                    test ax, ax
'00636027    7d12                    jge 63603b
'00636029    6888000000              push 00000088
'0063602e    684cfd4200              push 0042fd4c
'00636033    57                      push edi
'00636034    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00636035    ff1580104000            call dword ptr [00401080]
'0063603b    8b0b                    mov ecx, dword ptr [ebx]
'0063603d    d985dcfeffff            fld dword ptr [ebp+fffffedc]
'00636043    d82534554000            fsub dword ptr [00405534]
'00636049    dfe0                    fnstsw ax
'0063604b    a80d                    test al, 0d
'0063604d    0f8551050000            jne 6365a4
'00636053    51                      push ecx
'00636054    d91c24                  fstp dword ptr [esp]
'var_822 = (00)
'00636057    53                      push ebx
'00636058    ff517c                  call dword ptr [ecx+7c]
'0063605b    dbe2                    fnclex
'0063605d    85c0                    test ax, ax
'0063605f    7d0f                    jge 636070
'00636061    6a7c                    push 7c
'00636063    6810074300              push 00430710
'00636068    53                      push ebx
'00636069    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0063606a    ff1580104000            call dword ptr [00401080]
'00636070    a15cb07200              mov ax, word ptr [0072b05c]
'00636075    85c0                    test ax, ax
'00636077    7510                    jne 636089
'00636079    685cb07200              push 0072b05c
'0063607e    68347b4000              push 00407b34

' *** Reference to "__vbaNew2"
'00636083    ff152c124000            call dword ptr [0040122c]
Set var_24 = New frmAcceder
'00636089    8b3d5cb07200            mov edi, dword ptr [0072b05c]
'0063608f    8b17                    mov edx, dword ptr [edi]
'00636091    6800008745              push 45870000
'00636096    57                      push edi
'00636097    ff928c000000            call dword ptr [edx+0000008c]
'0063609d    dbe2                    fnclex
'0063609f    85c0                    test ax, ax
'006360a1    7d12                    jge 6360b5
'006360a3    688c000000              push 0000008c
'006360a8    6810074300              push 00430710
'006360ad    57                      push edi
'006360ae    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006360af    ff1580104000            call dword ptr [00401080]
'006360b5    a15cb07200              mov ax, word ptr [0072b05c]
'006360ba    85c0                    test ax, ax
'006360bc    7510                    jne 6360ce
'006360be    685cb07200              push 0072b05c
'006360c3    68347b4000              push 00407b34

' *** Reference to "__vbaNew2"
'006360c8    ff152c124000            call dword ptr [0040122c]
Set var_24 = New frmAcceder
'006360ce    8b3d5cb07200            mov edi, dword ptr [0072b05c]
'006360d4    8b07                    mov eax, dword ptr [edi]
'006360d6    680040ce45              push 45ce4000
'006360db    57                      push edi
'006360dc    ff9084000000            call dword ptr [eax+00000084]
'006360e2    dbe2                    fnclex
'006360e4    85c0                    test ax, ax
'006360e6    7d12                    jge 6360fa
'006360e8    6884000000              push 00000084
'006360ed    6810074300              push 00430710
'006360f2    57                      push edi
'006360f3    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006360f4    ff1580104000            call dword ptr [00401080]
'006360fa    b800106f45              mov eax, 456f1000
'006360ff    898528ffffff            mov dword ptr [ebp+ffffff28], eax
'00636105    b904000000              mov ecx, 00000004
'0063610a    898d20ffffff            mov dword ptr [ebp+ffffff20], ecx
'00636110    83ec10                  sub esp, 10
'00636113    8bd4                    mov edx, esp
'00636115    890a                    mov dword ptr [edx], ecx
'00636117    8b8d24ffffff            mov ecx, dword ptr [ebp+ffffff24]
'0063611d    894a04                  mov dword ptr [edx+04], ecx
'00636120    894208                  mov dword ptr [edx+08], eax
'00636123    8b852cffffff            mov eax, dword ptr [ebp+ffffff2c]
'00636129    89420c                  mov dword ptr [edx+0c], eax
'0063612c    6806000180              push 80010006
'00636131    8b0e                    mov ecx, dword ptr [esi]
'00636133    56                      push esi
'00636134    ff91fc020000            call dword ptr [ecx+000002fc]
'0063613a    50                      push eax
'0063613b    8d55c4                  lea edx, dword ptr [ebp-3c]
'0063613e    52                      push edx

' *** Reference to "__vbaObjSet"
'0063613f    8b3db4104000            mov edi, dword ptr [004010b4]
'00636145    ffd7                    call edi
Set var_9 = Nothing
'00636147    50                      push eax

' *** Reference to "__vbaLateIdSt"
'00636148    8b1dec124000            mov ebx, dword ptr [004012ec]
'0063614e    ffd3                    call ebx
var_9.[UNMANAGED] = 1164906496
'00636150    8d4dc4                  lea ecx, dword ptr [ebp-3c]

' *** Reference to "__vbaFreeObj"
'00636153    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_9)
'00636159    b800c0c645              mov eax, 45c6c000
'0063615e    898528ffffff            mov dword ptr [ebp+ffffff28], eax
'00636164    b904000000              mov ecx, 00000004
'00636169    898d20ffffff            mov dword ptr [ebp+ffffff20], ecx
'0063616f    83ec10                  sub esp, 10
'00636172    8bd4                    mov edx, esp
'00636174    890a                    mov dword ptr [edx], ecx
'00636176    8b8d24ffffff            mov ecx, dword ptr [ebp+ffffff24]
'0063617c    894a04                  mov dword ptr [edx+04], ecx
'0063617f    894208                  mov dword ptr [edx+08], eax
'00636182    8b852cffffff            mov eax, dword ptr [ebp+ffffff2c]
'00636188    89420c                  mov dword ptr [edx+0c], eax
'0063618b    6805000180              push 80010005
'00636190    8b0e                    mov ecx, dword ptr [esi]
'00636192    56                      push esi
'00636193    ff91fc020000            call dword ptr [ecx+000002fc]
'00636199    50                      push eax
'0063619a    8d55c4                  lea edx, dword ptr [ebp-3c]
'0063619d    52                      push edx
'0063619e    ffd7                    call edi
Set var_9 = Nothing
'006361a0    50                      push eax
'006361a1    ffd3                    call ebx
var_9.[UNMANAGED] = 1170653184
'006361a3    8d4dc4                  lea ecx, dword ptr [ebp-3c]

' *** Reference to "__vbaFreeObj"
'006361a6    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_9)
'006361ac    8b06                    mov eax, dword ptr [esi]
'006361ae    56                      push esi
'006361af    ff90f8060000            call dword ptr [eax+000006f8]
'006361b5    85c0                    test ax, ax
'006361b7    7d12                    jge 6361cb
'006361b9    68f8060000              push 000006f8
'006361be    6840074300              push 00430740
'006361c3    56                      push esi
'006361c4    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006361c5    ff1580104000            call dword ptr [00401080]

' *** Reference to "__vbaExitProc"
'006361cb    ff15a0104000            call dword ptr [004010a0]
'006361d1    9b                      fwait
'006361d2    6885656300              push 00636585
'006361d7    e9a8030000              jmp 636584

' *** Reference to "rtcErrObj"
'006361dc    8b1d6c124000            mov ebx, dword ptr [0040126c]
'006361e2    ffd3                    call ebx
'006361e4    50                      push eax
'006361e5    8d55c4                  lea edx, dword ptr [ebp-3c]
'006361e8    52                      push edx

' *** Reference to "__vbaObjSet"
'006361e9    8b3db4104000            mov edi, dword ptr [004010b4]
'006361ef    ffd7                    call edi
Set var_9 = Err
'006361f1    8bf0                    mov esi, eax
'006361f3    8b06                    mov eax, dword ptr [esi]
'006361f5    8d4de0                  lea ecx, dword ptr [ebp-20]
'006361f8    51                      push ecx
'006361f9    56                      push esi
'006361fa    ff502c                  call dword ptr [eax+2c]
var_3 = var_9.Description
'006361fd    dbe2                    fnclex
'006361ff    85c0                    test ax, ax
'00636201    7d0f                    jge 636212
'00636203    6a2c                    push 2c
'00636205    685c084300              push 0043085c
'0063620a    56                      push esi
'0063620b    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0063620c    ff1580104000            call dword ptr [00401080]
'00636212    ffd3                    call ebx
'00636214    50                      push eax
'00636215    8d55c0                  lea edx, dword ptr [ebp-40]
'00636218    52                      push edx
'00636219    ffd7                    call edi
Set var_5 = Err
'0063621b    8bf0                    mov esi, eax
'0063621d    8b06                    mov eax, dword ptr [esi]
'0063621f    8d8ddcfeffff            lea ecx, dword ptr [ebp+fffffedc]
'00636225    51                      push ecx
'00636226    56                      push esi
'00636227    ff501c                  call dword ptr [eax+1c]
var_10 = var_5.Number
'0063622a    dbe2                    fnclex
'0063622c    85c0                    test ax, ax
'0063622e    7d0f                    jge 63623f
'00636230    6a1c                    push 1c
'00636232    685c084300              push 0043085c
'00636237    56                      push esi
'00636238    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00636239    ff1580104000            call dword ptr [00401080]
'0063623f    b904000280              mov ecx, 80020004
'00636244    894d98                  mov dword ptr [ebp-68], ecx
'00636247    b80a000000              mov eax, 0000000a
'0063624c    894590                  mov dword ptr [ebp-70], eax
'0063624f    894da8                  mov dword ptr [ebp-58], ecx
'00636252    8945a0                  mov dword ptr [ebp-60], eax
'00636255    c78528ffffff24b07200    mov dword ptr [ebp+ffffff28], 0072b024
'0063625f    c78520ffffff08400000    mov dword ptr [ebp+ffffff20], 00004008
'00636269    6814084300              push 00430814
'0063626e    8b55e0                  mov edx, dword ptr [ebp-20]
'00636271    52                      push edx

' *** Reference to "__vbaStrCat"
'00636272    8b3570104000            mov esi, dword ptr [00401070]
'00636278    ffd6                    call esi
var_11 = ("L'erreur suivante s'est produite : ") & (var_3)
'0063627a    8bd0                    mov edx, eax
'0063627c    8d4ddc                  lea ecx, dword ptr [ebp-24]

' *** Reference to "__vbaStrMove"
'0063627f    8b3dd0124000            mov edi, dword ptr [004012d0]
'00636285    ffd7                    call edi
'00636287    50                      push eax
'00636288    6870084300              push 00430870
'0063628d    ffd6                    call esi
var_139 = (var_11) & (vbCrLf)
'0063628f    8bd0                    mov edx, eax
'00636291    8d4dd8                  lea ecx, dword ptr [ebp-28]
'00636294    ffd7                    call edi
'00636296    50                      push eax
'00636297    6870084300              push 00430870
'0063629c    ffd6                    call esi
var_56 = (var_139) & (vbCrLf)
'0063629e    8bd0                    mov edx, eax
'006362a0    8d4dd4                  lea ecx, dword ptr [ebp-2c]
'006362a3    ffd7                    call edi
'006362a5    50                      push eax
'006362a6    6880084300              push 00430880
'006362ab    ffd6                    call esi
var_2462 = (var_56) & (" numéro d'erreur (")
'006362ad    8bd0                    mov edx, eax
'006362af    8d4dd0                  lea ecx, dword ptr [ebp-30]
'006362b2    ffd7                    call edi
'006362b4    50                      push eax
'006362b5    8b85dcfeffff            mov eax, dword ptr [ebp+fffffedc]
'006362bb    50                      push eax

' *** Reference to "__vbaStrI4"
'006362bc    ff1520104000            call dword ptr [00401020]
'006362c2    8bd0                    mov edx, eax
'006362c4    8d4dcc                  lea ecx, dword ptr [ebp-34]
'006362c7    ffd7                    call edi
'006362c9    50                      push eax
'006362ca    ffd6                    call esi
var_1231 = (var_2462) & (CStr(var_10))
'006362cc    8bd0                    mov edx, eax
'006362ce    8d4dc8                  lea ecx, dword ptr [ebp-38]
'006362d1    ffd7                    call edi
'006362d3    50                      push eax
'006362d4    68ac084300              push 004308ac
'006362d9    ffd6                    call esi
var_2068 = (var_1231) & (" )")
'006362db    8945b8                  mov dword ptr [ebp-48], eax
'006362de    bf08000000              mov edi, 00000008
'006362e3    897db0                  mov dword ptr [ebp-50], edi
'006362e6    8d4d90                  lea ecx, dword ptr [ebp-70]
'006362e9    51                      push ecx
'006362ea    8d55a0                  lea edx, dword ptr [ebp-60]
'006362ed    52                      push edx
'006362ee    8d8520ffffff            lea eax, dword ptr [ebp+ffffff20]
'006362f4    50                      push eax
'006362f5    6a10                    push 10
'006362f7    8d4db0                  lea ecx, dword ptr [ebp-50]
'006362fa    51                      push ecx

' *** Reference to "rtcMsgBox"
'006362fb    ff15b8104000            call dword ptr [004010b8]
var_2069 = MsgBox(var_2068, 16, vbNullString)
'00636301    8d55c8                  lea edx, dword ptr [ebp-38]
'00636304    52                      push edx
'00636305    8d45cc                  lea eax, dword ptr [ebp-34]
'00636308    50                      push eax
'00636309    8d4dd0                  lea ecx, dword ptr [ebp-30]
'0063630c    51                      push ecx
'0063630d    8d55d4                  lea edx, dword ptr [ebp-2c]
'00636310    52                      push edx
'00636311    8d45d8                  lea eax, dword ptr [ebp-28]
'00636314    50                      push eax
'00636315    8d4ddc                  lea ecx, dword ptr [ebp-24]
'00636318    51                      push ecx
'00636319    8d55e0                  lea edx, dword ptr [ebp-20]
'0063631c    52                      push edx
'0063631d    6a07                    push 07

' *** Reference to "__vbaFreeStrList"
'0063631f    ff155c124000            call dword ptr [0040125c]
'#Cleanup( 0, -4564, -4568, -4572, -4576, -4580, -4584)
'00636325    8d45c0                  lea eax, dword ptr [ebp-40]
'00636328    50                      push eax
'00636329    8d4dc4                  lea ecx, dword ptr [ebp-3c]
'0063632c    51                      push ecx
'0063632d    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'0063632f    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_9, var_5)
'00636335    8d5590                  lea edx, dword ptr [ebp-70]
'00636338    52                      push edx
'00636339    8d45a0                  lea eax, dword ptr [ebp-60]
'0063633c    50                      push eax
'0063633d    8d4db0                  lea ecx, dword ptr [ebp-50]
'00636340    51                      push ecx
'00636341    6a03                    push 03

' *** Reference to "__vbaFreeVarList"
'00636343    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_6, var_7, var_8)
'00636349    83c43c                  add esp, 3c
'0063634c    8d55b0                  lea edx, dword ptr [ebp-50]
'0063634f    52                      push edx

' *** Reference to "rtcGetPresentDate"
'00636350    ff15f4124000            call dword ptr [004012f4]
var_2068 = Now()
'00636356    c78528ffffffb8084300    mov dword ptr [ebp+ffffff28], 004308b8
'00636360    89bd20ffffff            mov dword ptr [ebp+ffffff20], edi
'00636366    8d9520ffffff            lea edx, dword ptr [ebp+ffffff20]
'0063636c    8d4da0                  lea ecx, dword ptr [ebp-60]

' *** Reference to "__vbaVarDup"
'0063636f    ff1598124000            call dword ptr [00401298]
var_7 = ("dd/MM/yyyy hh:mm:ss")
'00636375    6a01                    push 01
'00636377    6a01                    push 01
'00636379    8d45a0                  lea eax, dword ptr [ebp-60]
'0063637c    50                      push eax
'0063637d    8d4db0                  lea ecx, dword ptr [ebp-50]
'00636380    51                      push ecx
'00636381    8d5590                  lea edx, dword ptr [ebp-70]
'00636384    52                      push edx

' *** Reference to "rtcVarFromFormatVar"
'00636385    ff1574104000            call dword ptr [00401074]
'0063638b    ffd3                    call ebx
'0063638d    50                      push eax
'0063638e    8d45c4                  lea eax, dword ptr [ebp-3c]
'00636391    50                      push eax

' *** Reference to "__vbaObjSet"
'00636392    ff15b4104000            call dword ptr [004010b4]
Set var_9 = Err
'00636398    8bf0                    mov esi, eax
'0063639a    8b0e                    mov ecx, dword ptr [esi]
'0063639c    8d95dcfeffff            lea edx, dword ptr [ebp+fffffedc]
'006363a2    52                      push edx
'006363a3    56                      push esi
'006363a4    ff511c                  call dword ptr [ecx+1c]
var_10 = var_9.Number
'006363a7    dbe2                    fnclex
'006363a9    85c0                    test ax, ax
'006363ab    7d0f                    jge 6363bc
'006363ad    6a1c                    push 1c
'006363af    685c084300              push 0043085c
'006363b4    56                      push esi
'006363b5    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006363b6    ff1580104000            call dword ptr [00401080]
'006363bc    ffd3                    call ebx
'006363be    50                      push eax
'006363bf    8d45c0                  lea eax, dword ptr [ebp-40]
'006363c2    50                      push eax

' *** Reference to "__vbaObjSet"
'006363c3    ff15b4104000            call dword ptr [004010b4]
Set var_5 = Err
'006363c9    8bf0                    mov esi, eax
'006363cb    8b0e                    mov ecx, dword ptr [esi]
'006363cd    8d55e0                  lea edx, dword ptr [ebp-20]
'006363d0    52                      push edx
'006363d1    56                      push esi
'006363d2    ff512c                  call dword ptr [ecx+2c]
var_3 = var_5.Description
'006363d5    dbe2                    fnclex
'006363d7    85c0                    test ax, ax
'006363d9    7d0f                    jge 6363ea
'006363db    6a2c                    push 2c
'006363dd    685c084300              push 0043085c
'006363e2    56                      push esi
'006363e3    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006363e4    ff1580104000            call dword ptr [00401080]
'006363ea    c78518ffffffe4084300    mov dword ptr [ebp+ffffff18], 004308e4
'006363f4    89bd10ffffff            mov dword ptr [ebp+ffffff10], edi
'006363fa    8b85dcfeffff            mov eax, dword ptr [ebp+fffffedc]
'00636400    898508ffffff            mov dword ptr [ebp+ffffff08], eax
'00636406    c78500ffffff03000000    mov dword ptr [ebp+ffffff00], 00000003
'00636410    c785f8feffff08094300    mov dword ptr [ebp+fffffef8], 00430908
'0063641a    89bdf0feffff            mov dword ptr [ebp+fffffef0], edi
'00636420    8b45e0                  mov eax, dword ptr [ebp-20]
'00636423    c745e000000000          mov dword ptr [ebp-20], 00000000
'0063642a    898558ffffff            mov dword ptr [ebp+ffffff58], eax
'00636430    89bd50ffffff            mov dword ptr [ebp+ffffff50], edi
'00636436    c785e8feffffb4dd4400    mov dword ptr [ebp+fffffee8], 0044ddb4
'00636440    89bde0feffff            mov dword ptr [ebp+fffffee0], edi
'00636446    8d4d90                  lea ecx, dword ptr [ebp-70]
'00636449    51                      push ecx
'0063644a    8d9510ffffff            lea edx, dword ptr [ebp+ffffff10]
'00636450    52                      push edx
'00636451    8d4580                  lea eax, dword ptr [ebp-80]
'00636454    50                      push eax

' *** Reference to "__vbaVarCat"
'00636455    8b3508124000            mov esi, dword ptr [00401208]
'0063645b    ffd6                    call esi
'0063645d    50                      push eax
'0063645e    8d8d00ffffff            lea ecx, dword ptr [ebp+ffffff00]
'00636464    51                      push ecx
'00636465    8d9570ffffff            lea edx, dword ptr [ebp+ffffff70]
'0063646b    52                      push edx
'0063646c    ffd6                    call esi
'0063646e    50                      push eax
'0063646f    8d85f0feffff            lea eax, dword ptr [ebp+fffffef0]
'00636475    50                      push eax
'00636476    8d8d60ffffff            lea ecx, dword ptr [ebp+ffffff60]
'0063647c    51                      push ecx
'0063647d    ffd6                    call esi
'0063647f    50                      push eax
'00636480    8d9550ffffff            lea edx, dword ptr [ebp+ffffff50]
'00636486    52                      push edx
'00636487    8d8540ffffff            lea eax, dword ptr [ebp+ffffff40]
'0063648d    50                      push eax
'0063648e    ffd6                    call esi
'00636490    50                      push eax
'00636491    8d8de0feffff            lea ecx, dword ptr [ebp+fffffee0]
'00636497    51                      push ecx
'00636498    8d9530ffffff            lea edx, dword ptr [ebp+ffffff30]
'0063649e    52                      push edx
'0063649f    ffd6                    call esi
'006364a1    50                      push eax
'006364a2    33c0                    xor eax, eax
'006364a4    66a12ab07200            mov eax, dword ptr [0072b02a]
'006364aa    50                      push eax
'006364ab    6884094300              push 00430984

' *** Reference to "__vbaPrintFile"
'006364b0    ff15b8114000            call dword ptr [004011b8]
Print #0, 
'006364b6    8d4dc0                  lea ecx, dword ptr [ebp-40]
'006364b9    51                      push ecx
'006364ba    8d55c4                  lea edx, dword ptr [ebp-3c]
'006364bd    52                      push edx
'006364be    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'006364c0    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_9, var_5)
'006364c6    8d8530ffffff            lea eax, dword ptr [ebp+ffffff30]
'006364cc    50                      push eax
'006364cd    8d8d40ffffff            lea ecx, dword ptr [ebp+ffffff40]
'006364d3    51                      push ecx
'006364d4    8d9550ffffff            lea edx, dword ptr [ebp+ffffff50]
'006364da    52                      push edx
'006364db    8d8560ffffff            lea eax, dword ptr [ebp+ffffff60]
'006364e1    50                      push eax
'006364e2    8d8d70ffffff            lea ecx, dword ptr [ebp+ffffff70]
'006364e8    51                      push ecx
'006364e9    8d5580                  lea edx, dword ptr [ebp-80]
'006364ec    52                      push edx
'006364ed    8d4590                  lea eax, dword ptr [ebp-70]
'006364f0    50                      push eax
'006364f1    8d4da0                  lea ecx, dword ptr [ebp-60]
'006364f4    51                      push ecx
'006364f5    8d55b0                  lea edx, dword ptr [ebp-50]
'006364f8    52                      push edx
'006364f9    6a09                    push 09

' *** Reference to "__vbaFreeVarList"
'006364fb    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_6, var_7, var_8, var_18, var_19, var_20, var_21, var_22, var_23)
'00636501    83c440                  add esp, 40
'00636504    6a00                    push 00

' *** Reference to "__vbaResume"
'00636506    ff1568104000            call dword ptr [00401068]
'0063650c    e9bafcffff              jmp 6361cb
Resume handler_6361CB
'00636511    8d4dc8                  lea ecx, dword ptr [ebp-38]
'00636514    51                      push ecx
'00636515    8d55cc                  lea edx, dword ptr [ebp-34]
'00636518    52                      push edx
'00636519    8d45d0                  lea eax, dword ptr [ebp-30]
'0063651c    50                      push eax
'0063651d    8d4dd4                  lea ecx, dword ptr [ebp-2c]
'00636520    51                      push ecx
'00636521    8d55d8                  lea edx, dword ptr [ebp-28]
'00636524    52                      push edx
'00636525    8d45dc                  lea eax, dword ptr [ebp-24]
'00636528    50                      push eax
'00636529    8d4de0                  lea ecx, dword ptr [ebp-20]
'0063652c    51                      push ecx
'0063652d    6a07                    push 07

' *** Reference to "__vbaFreeStrList"
'0063652f    ff155c124000            call dword ptr [0040125c]
'#Cleanup( var_3, -4564, -4568, -4572, -4576, -4580, -4584)
'00636535    8d55c0                  lea edx, dword ptr [ebp-40]
'00636538    52                      push edx
'00636539    8d45c4                  lea eax, dword ptr [ebp-3c]
'0063653c    50                      push eax
'0063653d    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'0063653f    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_9, var_5)
'00636545    8d8d30ffffff            lea ecx, dword ptr [ebp+ffffff30]
'0063654b    51                      push ecx
'0063654c    8d9540ffffff            lea edx, dword ptr [ebp+ffffff40]
'00636552    52                      push edx
'00636553    8d8550ffffff            lea eax, dword ptr [ebp+ffffff50]
'00636559    50                      push eax
'0063655a    8d8d60ffffff            lea ecx, dword ptr [ebp+ffffff60]
'00636560    51                      push ecx
'00636561    8d9570ffffff            lea edx, dword ptr [ebp+ffffff70]
'00636567    52                      push edx
'00636568    8d4580                  lea eax, dword ptr [ebp-80]
'0063656b    50                      push eax
'0063656c    8d4d90                  lea ecx, dword ptr [ebp-70]
'0063656f    51                      push ecx
'00636570    8d55a0                  lea edx, dword ptr [ebp-60]
'00636573    52                      push edx
'00636574    8d45b0                  lea eax, dword ptr [ebp-50]
'00636577    50                      push eax
'00636578    6a09                    push 09

' *** Reference to "__vbaFreeVarList"
'0063657a    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_6, var_7, var_8, var_18, var_19, var_20, var_21, var_22, var_23)
'00636580    83c454                  add esp, 54
'00636583    c3                      ret
'00636584    c3                      ret
'00636585    8b4508                  mov eax, dword ptr [ebp+08]
'00636588    8b08                    mov ecx, dword ptr [eax]
'0063658a    50                      push eax
'0063658b    ff5108                  call dword ptr [ecx+08]
'0063658e    8b45f4                  mov eax, dword ptr [ebp-0c]
'00636591    8b4de4                  mov ecx, dword ptr [ebp-1c]
    'Reference to '__except_list'
'00636594    64890d00000000          mov dword ptr fs:[00000000], ecx
'0063659b    5f                      pop edi
'0063659c    5e                      pop esi
'0063659d    5b                      pop ebx
'0063659e    8be5                    mov esp, ebp
'006365a0    5d                      pop ebp
'006365a1    c20400                  ret 0004


End Sub


Private Sub Form_Resize()
'006365b0    55                      push ebp
'006365b1    8bec                    mov ebp, esp
'006365b3    83ec0c                  sub esp, 0c

' *** Reference to "__vbaExceptHandler"
'006365b6    6866724000              push 00407266
'006365bb    64a100000000            mov ax, word ptr fs:[00000000]
'006365c1    50                      push eax
    'Reference to '__except_list'
'006365c2    64892500000000          mov dword ptr fs:[00000000], esp
'006365c9    83ec7c                  sub esp, 7c
'006365cc    53                      push ebx
'006365cd    56                      push esi
'006365ce    57                      push edi
'006365cf    8965f4                  mov dword ptr [ebp-0c], esp
'006365d2    c745f848554000          mov dword ptr [ebp-08], 00405548
'006365d9    8b7508                  mov esi, dword ptr [ebp+08]
'006365dc    8bc6                    mov eax, esi
'006365de    83e001                  and eax, 01
'006365e1    8945fc                  mov dword ptr [ebp-04], eax
'006365e4    83e6fe                  and esi, fffffffe
'006365e7    8b0e                    mov ecx, dword ptr [esi]
'006365e9    56                      push esi
'006365ea    897508                  mov dword ptr [ebp+08], esi
'006365ed    ff5104                  call dword ptr [ecx+04]
'006365f0    a15cb07200              mov ax, word ptr [0072b05c]
'006365f5    33db                    xor ebx, ebx
'006365f7    3bc3                    cmp eax, ebx
'006365f9    895de8                  mov dword ptr [ebp-18], ebx
'006365fc    895dd8                  mov dword ptr [ebp-28], ebx
'006365ff    895dc8                  mov dword ptr [ebp-38], ebx
'00636602    895d84                  mov dword ptr [ebp-7c], ebx
'00636605    7510                    jne 636617

If (0 = 0) Then
'00636607    685cb07200              push 0072b05c
'0063660c    68347b4000              push 00407b34

' *** Reference to "__vbaNew2"
'00636611    ff152c124000            call dword ptr [0040122c]
    Dim var_24 As New frmAcceder
End If
'00636617    8b3d5cb07200            mov edi, dword ptr [0072b05c]
'0063661d    8b17                    mov edx, dword ptr [edi]
'0063661f    8d4584                  lea eax, dword ptr [ebp-7c]
'00636622    50                      push eax
'00636623    57                      push edi
'00636624    ff9288000000            call dword ptr [edx+00000088]
'0063662a    dbe2                    fnclex
'0063662c    3bc3                    cmp eax, ebx
'0063662e    7d12                    jge 636642
'00636630    6888000000              push 00000088
'00636635    6810074300              push 00430710
'0063663a    57                      push edi
'0063663b    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0063663c    ff1580104000            call dword ptr [00401080]
'00636642    d94584                  fld dword ptr [ebp-7c]
'00636645    8d4dc8                  lea ecx, dword ptr [ebp-38]
'00636648    d82540554000            fsub dword ptr [00405540]
'0063664e    51                      push ecx
'0063664f    8d55d8                  lea edx, dword ptr [ebp-28]
'00636652    52                      push edx
'00636653    d95dd0                  fstp dword ptr [ebp-30]
'var_4 = (00)
'00636656    dfe0                    fnstsw ax
'00636658    a80d                    test al, 0d
'0063665a    0f8571010000            jne 6367d1
'00636660    c745c804000000          mov dword ptr [ebp-38], 00000004
'00636667    c745e001000000          mov dword ptr [ebp-20], 00000001
'0063666e    c745d802000000          mov dword ptr [ebp-28], 00000002
'00636675    e876c0ebff              call 4f26f0
Call sub_4F26F0()
'0063667a    8b559c                  mov edx, dword ptr [ebp-64]
'0063667d    0fbfc0                  movsx eax, ax
'00636680    898570ffffff            mov dword ptr [ebp+ffffff70], eax
'00636686    83ec10                  sub esp, 10
'00636689    8bcc                    mov ecx, esp
'0063668b    db8570ffffff            fild dword ptr [ebp+ffffff70]
'00636691    b804000000              mov eax, 00000004
'00636696    8901                    mov dword ptr [ecx], eax
'00636698    895104                  mov dword ptr [ecx+04], edx
'0063669b    8b55a4                  mov edx, dword ptr [ebp-5c]
'0063669e    d95da0                  fstp dword ptr [ebp-60]
'var_7 = (00)
'006366a1    8b45a0                  mov eax, dword ptr [ebp-60]
'006366a4    894108                  mov dword ptr [ecx+08], eax
'006366a7    8b06                    mov eax, dword ptr [esi]
'006366a9    6806000180              push 80010006
'006366ae    56                      push esi
'006366af    89510c                  mov dword ptr [ecx+0c], edx
'006366b2    ff90fc020000            call dword ptr [eax+000002fc]

' *** Reference to "__vbaObjSet"
'006366b8    8b1db4104000            mov ebx, dword ptr [004010b4]
'006366be    50                      push eax
'006366bf    8d4de8                  lea ecx, dword ptr [ebp-18]
'006366c2    51                      push ecx
'006366c3    ffd3                    call ebx
Set var_41 = Nothing
'006366c5    50                      push eax

' *** Reference to "__vbaLateIdSt"
'006366c6    ff15ec124000            call dword ptr [004012ec]
var_41.[UNMANAGED] = (-380)
'006366cc    8d4de8                  lea ecx, dword ptr [ebp-18]

' *** Reference to "__vbaFreeObj"
'006366cf    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_41)
'006366d5    8d55c8                  lea edx, dword ptr [ebp-38]
'006366d8    52                      push edx
'006366d9    8d45d8                  lea eax, dword ptr [ebp-28]
'006366dc    50                      push eax
'006366dd    6a02                    push 02

' *** Reference to "__vbaFreeVarList"
'006366df    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_45, var_46)
'006366e5    a15cb07200              mov ax, word ptr [0072b05c]
'006366ea    83c40c                  add esp, 0c
'006366ed    85c0                    test ax, ax
'006366ef    7510                    jne 636701
'006366f1    685cb07200              push 0072b05c
'006366f6    68347b4000              push 00407b34

' *** Reference to "__vbaNew2"
'006366fb    ff152c124000            call dword ptr [0040122c]
Set var_24 = New frmAcceder
'00636701    8b3d5cb07200            mov edi, dword ptr [0072b05c]
'00636707    8b0f                    mov ecx, dword ptr [edi]
'00636709    8d5584                  lea edx, dword ptr [ebp-7c]
'0063670c    52                      push edx
'0063670d    57                      push edi
'0063670e    ff9180000000            call dword ptr [ecx+00000080]
'00636714    dbe2                    fnclex
'00636716    85c0                    test ax, ax
'00636718    7d12                    jge 63672c
'0063671a    6880000000              push 00000080
'0063671f    6810074300              push 00430710
'00636724    57                      push edi
'00636725    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00636726    ff1580104000            call dword ptr [00401080]
'0063672c    d94584                  fld dword ptr [ebp-7c]
'0063672f    8b55bc                  mov edx, dword ptr [ebp-44]
'00636732    d8253c554000            fsub dword ptr [0040553c]
'00636738    d95dc0                  fstp dword ptr [ebp-40]
'var_5 = (00)
'0063673b    dfe0                    fnstsw ax
'0063673d    a80d                    test al, 0d
'0063673f    0f858c000000            jne 6367d1
'00636745    83ec10                  sub esp, 10
'00636748    8bcc                    mov ecx, esp
'0063674a    b804000000              mov eax, 00000004
'0063674f    8901                    mov dword ptr [ecx], eax
'00636751    8b45c0                  mov eax, dword ptr [ebp-40]
'00636754    895104                  mov dword ptr [ecx+04], edx
'00636757    8b55c4                  mov edx, dword ptr [ebp-3c]
'0063675a    894108                  mov dword ptr [ecx+08], eax
'0063675d    8b06                    mov eax, dword ptr [esi]
'0063675f    6805000180              push 80010005
'00636764    56                      push esi
'00636765    89510c                  mov dword ptr [ecx+0c], edx
'00636768    ff90fc020000            call dword ptr [eax+000002fc]
'0063676e    50                      push eax
'0063676f    8d4de8                  lea ecx, dword ptr [ebp-18]
'00636772    51                      push ecx
'00636773    ffd3                    call ebx
Set var_41 = Nothing
'00636775    50                      push eax

' *** Reference to "__vbaLateIdSt"
'00636776    ff15ec124000            call dword ptr [004012ec]
var_41.[UNMANAGED] = ((0) - 0!)
'0063677c    8d4de8                  lea ecx, dword ptr [ebp-18]

' *** Reference to "__vbaFreeObj"
'0063677f    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_41)
'00636785    c745fc00000000          mov dword ptr [ebp-04], 00000000
'0063678c    9b                      fwait
'0063678d    68b2676300              push 006367b2
'00636792    eb1d                    jmp 6367b1
'00636794    8d4de8                  lea ecx, dword ptr [ebp-18]

' *** Reference to "__vbaFreeObj"
'00636797    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_41)
'0063679d    8d55c8                  lea edx, dword ptr [ebp-38]
'006367a0    52                      push edx
'006367a1    8d45d8                  lea eax, dword ptr [ebp-28]
'006367a4    50                      push eax
'006367a5    6a02                    push 02

' *** Reference to "__vbaFreeVarList"
'006367a7    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_45, var_46)
'006367ad    83c40c                  add esp, 0c
'006367b0    c3                      ret
'006367b1    c3                      ret
'006367b2    8b4508                  mov eax, dword ptr [ebp+08]
'006367b5    8b08                    mov ecx, dword ptr [eax]
'006367b7    50                      push eax
'006367b8    ff5108                  call dword ptr [ecx+08]
'006367bb    8b45fc                  mov eax, dword ptr [ebp-04]
'006367be    8b4dec                  mov ecx, dword ptr [ebp-14]
'006367c1    5f                      pop edi
'006367c2    5e                      pop esi
    'Reference to '__except_list'
'006367c3    64890d00000000          mov dword ptr fs:[00000000], ecx
'006367ca    5b                      pop ebx
'006367cb    8be5                    mov esp, ebp
'006367cd    5d                      pop ebp
'006367ce    c20400                  ret 0004


End Sub


