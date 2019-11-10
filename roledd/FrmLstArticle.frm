VERSION 5.00
Object = "{D76D7130-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "C:\WINDOWS\SysWow64\Vsflex7d.ocx"

Begin VB.Form FrmLstArticle
    Caption = "Liste des objets"
    ScaleMode = 1
    AutoRedraw = 0              'False
    FontTransparent = -1              'True
    LinkTopic = "Form1"
    ControlBox = 0              'False
    KeyPreview = -1              'True
    ClientLeft   = 60
    ClientTop    = 420
    ClientWidth  = 8760
    ClientHeight = 5655
    BeginProperty Font
        Name          = "Times New Roman"
        Size          = 9
        Charset       = 0
        Weight        = 400
        Underline     = 0              'False
        Italic        = 0              'False
        Strikethrough = 0              'False
    EndProperty
    StartupPosition = 2
    Begin VB.TextBox fldstrdescription
        Left   = 45
        Top    = 3645
        Width  = 8700
        Height = 1500
        TabIndex = 7
        MultiLine = -1              'True
        ScrollBars = 2
    End
    Begin VB.TextBox fldnPrixUnitaire
        Left   = 1800
        Top    = 5205
        Width  = 975
        Height = 375
        TabIndex = 4
    End
    Begin VB.TextBox fldnNombre
        Left   = 5160
        Top    = 5205
        Width  = 495
        Height = 375
        Text = "1"
        TabIndex = 3
    End
    Begin VB.CommandButton BtnAnnuler
        Caption = "&Annuler (ESC)"
        Left   = 6210
        Top    = 5220
        Width  = 1335
        Height = 350
        TabIndex = 2
        Cancel = -1              'True
    End
    Begin VB.CommandButton BtnOK
        Caption = "&OK"
        Left   = 7695
        Top    = 5220
        Width  = 855
        Height = 350
        TabIndex = 1
        Default = -1              'True
    End
    Begin VSFlex7DAOCtl.VSFlexGrid FGarticle
        Left   = 0
        Top    = 0
        Width  = 8760
        Height = 3570
        TabIndex = 0
        OleObjectBlob = FrmLstArticle.frx:0000
        Align = 1
        Bindings = [UNPREDICTABLE]
    End
    Begin VB.Label LblObjet
        Caption = "Nombre d'objet à acheter"
        Index = 1
        Left   = 3120
        Top    = 5280
        Width  = 1785
        Height = 225
        TabIndex = 6
        AutoSize = -1              'True
    End
    Begin VB.Label LblObjet
        Caption = "Prix unitaire de l'objet"
        Index = 0
        Left   = 120
        Top    = 5280
        Width  = 1560
        Height = 225
        TabIndex = 5
        AutoSize = -1              'True
    End
End
Public Function remp(arg_0 As Unknow, arg_1 As Unknow, arg_2 As Unknow, arg_3 As Unknow, arg_4 As Unknow, arg_5 As Unknow, arg_6 As Unknow, arg_7 As Unknow, arg_8 As Unknow, arg_9 As Unknow, arg_A As Unknow, arg_B As Unknow, arg_C As Unknow, arg_D As Unknow, arg_E As Unknow, arg_F As Unknow, arg_10 As Unknow, arg_11 As Unknow, arg_12 As Unknow, arg_13 As Unknow, arg_14 As Unknow, arg_15 As Unknow, arg_16 As Unknow, arg_17 As Unknow, arg_18 As Unknow, arg_19 As Unknow, arg_1A As Unknow, arg_1B As Unknow, arg_1C As Unknow, arg_1D As Unknow, arg_1E As Unknow, arg_1F As Unknow, arg_20 As Unknow, arg_21 As Unknow, arg_22 As Unknow, arg_23 As Unknow, arg_24 As Unknow, arg_25 As Unknow, arg_26 As Unknow, arg_27 As Unknow, arg_28 As Unknow, arg_29 As Unknow, arg_2A As Unknow, arg_2B As Unknow, arg_2C As Unknow, arg_2D As Unknow, arg_2E As Unknow, arg_2F As Unknow, arg_30 As Unknow, arg_31 As Unknow, arg_32 As Unknow, arg_33 As Unknow, arg_34 As Unknow, arg_35 As Unknow, arg_36 As Unknow, arg_37 As Unknow, arg_38 As Unknow, arg_39 As Unknow, arg_3A As Unknow, arg_3B As Unknow, arg_3C As Unknow)
'006a1ab0    55                      push ebp
'006a1ab1    8bec                    mov ebp, esp
'006a1ab3    83ec14                  sub esp, 14

' *** Reference to "__vbaExceptHandler"
'006a1ab6    6866724000              push 00407266
'006a1abb    64a100000000            mov ax, word ptr fs:[00000000]
'006a1ac1    50                      push eax
    'Reference to '__except_list'
'006a1ac2    64892500000000          mov dword ptr fs:[00000000], esp
'006a1ac9    81ec54020000            sub esp, 00000254
'006a1acf    53                      push ebx
'006a1ad0    56                      push esi
'006a1ad1    57                      push edi
'006a1ad2    8965ec                  mov dword ptr [ebp-14], esp
'006a1ad5    c745f078624000          mov dword ptr [ebp-10], 00406278
'006a1adc    33f6                    xor esi, esi
'006a1ade    8975f4                  mov dword ptr [ebp-0c], esi
'006a1ae1    8975f8                  mov dword ptr [ebp-08], esi
'006a1ae4    8b4508                  mov eax, dword ptr [ebp+08]
'006a1ae7    8b08                    mov ecx, dword ptr [eax]
'006a1ae9    50                      push eax
'006a1aea    ff5104                  call dword ptr [ecx+04]
'006a1aed    8975e0                  mov dword ptr [ebp-20], esi
'006a1af0    8975dc                  mov dword ptr [ebp-24], esi
'006a1af3    8975d8                  mov dword ptr [ebp-28], esi
'006a1af6    8975d4                  mov dword ptr [ebp-2c], esi
'006a1af9    8975d0                  mov dword ptr [ebp-30], esi
'006a1afc    8975cc                  mov dword ptr [ebp-34], esi
'006a1aff    8975c8                  mov dword ptr [ebp-38], esi
'006a1b02    8975c4                  mov dword ptr [ebp-3c], esi
'006a1b05    8975c0                  mov dword ptr [ebp-40], esi
'006a1b08    8975bc                  mov dword ptr [ebp-44], esi
'006a1b0b    8975b8                  mov dword ptr [ebp-48], esi
'006a1b0e    8975b4                  mov dword ptr [ebp-4c], esi
'006a1b11    8975b0                  mov dword ptr [ebp-50], esi
'006a1b14    8975ac                  mov dword ptr [ebp-54], esi
'006a1b17    8975a8                  mov dword ptr [ebp-58], esi
'006a1b1a    8975a4                  mov dword ptr [ebp-5c], esi
'006a1b1d    8975a0                  mov dword ptr [ebp-60], esi
'006a1b20    89759c                  mov dword ptr [ebp-64], esi
'006a1b23    897598                  mov dword ptr [ebp-68], esi
'006a1b26    897588                  mov dword ptr [ebp-78], esi
'006a1b29    89b578ffffff            mov dword ptr [ebp+ffffff78], esi
'006a1b2f    89b568ffffff            mov dword ptr [ebp+ffffff68], esi
'006a1b35    89b558ffffff            mov dword ptr [ebp+ffffff58], esi
'006a1b3b    89b548ffffff            mov dword ptr [ebp+ffffff48], esi
'006a1b41    89b538ffffff            mov dword ptr [ebp+ffffff38], esi
'006a1b47    89b528ffffff            mov dword ptr [ebp+ffffff28], esi
'006a1b4d    89b518ffffff            mov dword ptr [ebp+ffffff18], esi
'006a1b53    89b508ffffff            mov dword ptr [ebp+ffffff08], esi
'006a1b59    89b5f8feffff            mov dword ptr [ebp+fffffef8], esi
'006a1b5f    89b5e8feffff            mov dword ptr [ebp+fffffee8], esi
'006a1b65    89b5d8feffff            mov dword ptr [ebp+fffffed8], esi
'006a1b6b    89b5c8feffff            mov dword ptr [ebp+fffffec8], esi
'006a1b71    89b5b8feffff            mov dword ptr [ebp+fffffeb8], esi
'006a1b77    89b5a8feffff            mov dword ptr [ebp+fffffea8], esi
'006a1b7d    89b598feffff            mov dword ptr [ebp+fffffe98], esi
'006a1b83    89b588feffff            mov dword ptr [ebp+fffffe88], esi
'006a1b89    89b578feffff            mov dword ptr [ebp+fffffe78], esi
'006a1b8f    89b568feffff            mov dword ptr [ebp+fffffe68], esi
'006a1b95    89b558feffff            mov dword ptr [ebp+fffffe58], esi
'006a1b9b    89b548feffff            mov dword ptr [ebp+fffffe48], esi
'006a1ba1    89b538feffff            mov dword ptr [ebp+fffffe38], esi
'006a1ba7    89b528feffff            mov dword ptr [ebp+fffffe28], esi
'006a1bad    89b514feffff            mov dword ptr [ebp+fffffe14], esi
'006a1bb3    89b510feffff            mov dword ptr [ebp+fffffe10], esi
'006a1bb9    66393528b07200          cmp word ptr [0072b028], si
'006a1bc0    7508                    jne 6a1bca
'006a1bc2    6a01                    push 01

' *** Reference to "__vbaOnError"
'006a1bc4    ff15b0104000            call dword ptr [004010b0]
On Error Goto handler_0
'006a1bca    393524be7200            cmp dword ptr [0072be24], esi
'006a1bd0    7510                    jne 6a1be2
'006a1bd2    6824be7200              push 0072be24
'006a1bd7    6870004300              push 00430070

' *** Reference to "__vbaNew2"
'006a1bdc    ff152c124000            call dword ptr [0040122c]
Dim var_2 As New Global
'006a1be2    8b3d24be7200            mov edi, dword ptr [0072be24]
'006a1be8    8b07                    mov eax, dword ptr [edi]
'006a1bea    8d4dc0                  lea ecx, dword ptr [ebp-40]
'006a1bed    51                      push ecx
'006a1bee    57                      push edi
'006a1bef    ff5018                  call dword ptr [eax+18]
Set var_5 = var_2.Screen
'006a1bf2    dbe2                    fnclex
'006a1bf4    3bc6                    cmp eax, esi
'006a1bf6    0f8d78030000            jge 6a1f74

If (var_2.Screen < 0) Then
'006a1bfc    6a18                    push 18
'006a1bfe    6860004300              push 00430060
'006a1c03    57                      push edi
'006a1c04    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006a1c05    8b3d80104000            mov edi, dword ptr [00401080]
'006a1c0b    ffd7                    call edi
'006a1c0d    e968030000              jmp 6a1f7a

' *** Reference to "rtcErrObj"
'006a1c12    8b1d6c124000            mov ebx, dword ptr [0040126c]
'006a1c18    ffd3                    call ebx
'006a1c1a    50                      push eax
'006a1c1b    8d55c0                  lea edx, dword ptr [ebp-40]
'006a1c1e    52                      push edx

' *** Reference to "__vbaObjSet"
'006a1c1f    8b3db4104000            mov edi, dword ptr [004010b4]
'006a1c25    ffd7                    call edi
    Set var_5 = Err
'006a1c27    8bf0                    mov esi, eax
'006a1c29    8b06                    mov eax, dword ptr [esi]
'006a1c2b    8d4ddc                  lea ecx, dword ptr [ebp-24]
'006a1c2e    51                      push ecx
'006a1c2f    56                      push esi
'006a1c30    ff502c                  call dword ptr [eax+2c]
    var_39 = var_5.Description
'006a1c33    dbe2                    fnclex
'006a1c35    85c0                    test ax, ax
'006a1c37    7d0f                    jge 6a1c48
'006a1c39    6a2c                    push 2c
'006a1c3b    685c084300              push 0043085c
'006a1c40    56                      push esi
'006a1c41    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006a1c42    ff1580104000            call dword ptr [00401080]
'006a1c48    ffd3                    call ebx
'006a1c4a    50                      push eax
'006a1c4b    8d55bc                  lea edx, dword ptr [ebp-44]
'006a1c4e    52                      push edx
'006a1c4f    ffd7                    call edi
    Set var_58 = Err
'006a1c51    8bf0                    mov esi, eax
'006a1c53    8b06                    mov eax, dword ptr [esi]
'006a1c55    8d8d10feffff            lea ecx, dword ptr [ebp+fffffe10]
'006a1c5b    51                      push ecx
'006a1c5c    56                      push esi
'006a1c5d    ff501c                  call dword ptr [eax+1c]
    var_915 = var_58.Number
'006a1c60    dbe2                    fnclex
'006a1c62    85c0                    test ax, ax
'006a1c64    7d0f                    jge 6a1c75
'006a1c66    6a1c                    push 1c
'006a1c68    685c084300              push 0043085c
'006a1c6d    56                      push esi
'006a1c6e    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006a1c6f    ff1580104000            call dword ptr [00401080]
'006a1c75    ba04000280              mov edx, 80020004
'006a1c7a    899570ffffff            mov dword ptr [ebp+ffffff70], edx
'006a1c80    b90a000000              mov ecx, 0000000a
'006a1c85    898d68ffffff            mov dword ptr [ebp+ffffff68], ecx
'006a1c8b    895580                  mov dword ptr [ebp-80], edx
'006a1c8e    898d78ffffff            mov dword ptr [ebp+ffffff78], ecx
'006a1c94    c785b0feffff24b07200    mov dword ptr [ebp+fffffeb0], 0072b024
'006a1c9e    c785a8feffff08400000    mov dword ptr [ebp+fffffea8], 00004008
'006a1ca8    6814084300              push 00430814
'006a1cad    8b55dc                  mov edx, dword ptr [ebp-24]
'006a1cb0    52                      push edx

' *** Reference to "__vbaStrCat"
'006a1cb1    8b3570104000            mov esi, dword ptr [00401070]
'006a1cb7    ffd6                    call esi
    var_49 = ("L'erreur suivante s'est produite : ") & (var_39)
'006a1cb9    8bd0                    mov edx, eax
'006a1cbb    8d4dd8                  lea ecx, dword ptr [ebp-28]

' *** Reference to "__vbaStrMove"
'006a1cbe    8b3dd0124000            mov edi, dword ptr [004012d0]
'006a1cc4    ffd7                    call edi
'006a1cc6    50                      push eax
'006a1cc7    6870084300              push 00430870
'006a1ccc    ffd6                    call esi
    var_13 = (var_49) & (vbCrLf)
'006a1cce    8bd0                    mov edx, eax
'006a1cd0    8d4dd4                  lea ecx, dword ptr [ebp-2c]
'006a1cd3    ffd7                    call edi
'006a1cd5    50                      push eax
'006a1cd6    6870084300              push 00430870
'006a1cdb    ffd6                    call esi
    var_14 = (var_13) & (vbCrLf)
'006a1cdd    8bd0                    mov edx, eax
'006a1cdf    8d4dd0                  lea ecx, dword ptr [ebp-30]
'006a1ce2    ffd7                    call edi
'006a1ce4    50                      push eax
'006a1ce5    6880084300              push 00430880
'006a1cea    ffd6                    call esi
    var_127 = (var_14) & (" numéro d'erreur (")
'006a1cec    8bd0                    mov edx, eax
'006a1cee    8d4dcc                  lea ecx, dword ptr [ebp-34]
'006a1cf1    ffd7                    call edi
'006a1cf3    50                      push eax
'006a1cf4    8b8510feffff            mov eax, dword ptr [ebp+fffffe10]
'006a1cfa    50                      push eax

' *** Reference to "__vbaStrI4"
'006a1cfb    ff1520104000            call dword ptr [00401020]
'006a1d01    8bd0                    mov edx, eax
'006a1d03    8d4dc8                  lea ecx, dword ptr [ebp-38]
'006a1d06    ffd7                    call edi
'006a1d08    50                      push eax
'006a1d09    ffd6                    call esi
    var_16 = (var_127) & (CStr(var_915))
'006a1d0b    8bd0                    mov edx, eax
'006a1d0d    8d4dc4                  lea ecx, dword ptr [ebp-3c]
'006a1d10    ffd7                    call edi
'006a1d12    50                      push eax
'006a1d13    68ac084300              push 004308ac
'006a1d18    ffd6                    call esi
    var_128 = (var_16) & (" )")
'006a1d1a    894590                  mov dword ptr [ebp-70], eax
'006a1d1d    bf08000000              mov edi, 00000008
'006a1d22    897d88                  mov dword ptr [ebp-78], edi
'006a1d25    8d8d68ffffff            lea ecx, dword ptr [ebp+ffffff68]
'006a1d2b    51                      push ecx
'006a1d2c    8d9578ffffff            lea edx, dword ptr [ebp+ffffff78]
'006a1d32    52                      push edx
'006a1d33    8d85a8feffff            lea eax, dword ptr [ebp+fffffea8]
'006a1d39    50                      push eax
'006a1d3a    6a10                    push 10
'006a1d3c    8d4d88                  lea ecx, dword ptr [ebp-78]
'006a1d3f    51                      push ecx

' *** Reference to "rtcMsgBox"
'006a1d40    ff15b8104000            call dword ptr [004010b8]
    var_129 = MsgBox(var_128, 16, vbNullString)
'006a1d46    8d55c4                  lea edx, dword ptr [ebp-3c]
'006a1d49    52                      push edx
'006a1d4a    8d45c8                  lea eax, dword ptr [ebp-38]
'006a1d4d    50                      push eax
'006a1d4e    8d4dcc                  lea ecx, dword ptr [ebp-34]
'006a1d51    51                      push ecx
'006a1d52    8d55d0                  lea edx, dword ptr [ebp-30]
'006a1d55    52                      push edx
'006a1d56    8d45d4                  lea eax, dword ptr [ebp-2c]
'006a1d59    50                      push eax
'006a1d5a    8d4dd8                  lea ecx, dword ptr [ebp-28]
'006a1d5d    51                      push ecx
'006a1d5e    8d55dc                  lea edx, dword ptr [ebp-24]
'006a1d61    52                      push edx
'006a1d62    6a07                    push 07

' *** Reference to "__vbaFreeStrList"
'006a1d64    ff155c124000            call dword ptr [0040125c]
    '#Cleanup( 0, -4516, -4520, -4524, -4528, -4532, -4536)
'006a1d6a    8d45bc                  lea eax, dword ptr [ebp-44]
'006a1d6d    50                      push eax
'006a1d6e    8d4dc0                  lea ecx, dword ptr [ebp-40]
'006a1d71    51                      push ecx
'006a1d72    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'006a1d74    ff154c104000            call dword ptr [0040104c]
    '#Cleanup( var_5, var_58)
'006a1d7a    8d9568ffffff            lea edx, dword ptr [ebp+ffffff68]
'006a1d80    52                      push edx
'006a1d81    8d8578ffffff            lea eax, dword ptr [ebp+ffffff78]
'006a1d87    50                      push eax
'006a1d88    8d4d88                  lea ecx, dword ptr [ebp-78]
'006a1d8b    51                      push ecx
'006a1d8c    6a03                    push 03

' *** Reference to "__vbaFreeVarList"
'006a1d8e    ff1540104000            call dword ptr [00401040]
    '#Cleanup( var_131, var_87, var_132)
'006a1d94    83c43c                  add esp, 3c
'006a1d97    8d5588                  lea edx, dword ptr [ebp-78]
'006a1d9a    52                      push edx

' *** Reference to "rtcGetPresentDate"
'006a1d9b    ff15f4124000            call dword ptr [004012f4]
    var_128 = Now()
'006a1da1    c785b0feffffb8084300    mov dword ptr [ebp+fffffeb0], 004308b8
'006a1dab    89bda8feffff            mov dword ptr [ebp+fffffea8], edi
'006a1db1    8d95a8feffff            lea edx, dword ptr [ebp+fffffea8]
'006a1db7    8d8d78ffffff            lea ecx, dword ptr [ebp+ffffff78]

' *** Reference to "__vbaVarDup"
'006a1dbd    ff1598124000            call dword ptr [00401298]
    var_87 = ("dd/MM/yyyy hh:mm:ss")
'006a1dc3    6a01                    push 01
'006a1dc5    6a01                    push 01
'006a1dc7    8d8578ffffff            lea eax, dword ptr [ebp+ffffff78]
'006a1dcd    50                      push eax
'006a1dce    8d4d88                  lea ecx, dword ptr [ebp-78]
'006a1dd1    51                      push ecx
'006a1dd2    8d9568ffffff            lea edx, dword ptr [ebp+ffffff68]
'006a1dd8    52                      push edx

' *** Reference to "rtcVarFromFormatVar"
'006a1dd9    ff1574104000            call dword ptr [00401074]
'006a1ddf    ffd3                    call ebx
'006a1de1    50                      push eax
'006a1de2    8d45c0                  lea eax, dword ptr [ebp-40]
'006a1de5    50                      push eax

' *** Reference to "__vbaObjSet"
'006a1de6    ff15b4104000            call dword ptr [004010b4]
    Set var_5 = Err
'006a1dec    8bf0                    mov esi, eax
'006a1dee    8b0e                    mov ecx, dword ptr [esi]
'006a1df0    8d9510feffff            lea edx, dword ptr [ebp+fffffe10]
'006a1df6    52                      push edx
'006a1df7    56                      push esi
'006a1df8    ff511c                  call dword ptr [ecx+1c]
    var_915 = var_5.Number
'006a1dfb    dbe2                    fnclex
'006a1dfd    85c0                    test ax, ax
'006a1dff    7d0f                    jge 6a1e10
'006a1e01    6a1c                    push 1c
'006a1e03    685c084300              push 0043085c
'006a1e08    56                      push esi
'006a1e09    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006a1e0a    ff1580104000            call dword ptr [00401080]
'006a1e10    ffd3                    call ebx
'006a1e12    50                      push eax
'006a1e13    8d45bc                  lea eax, dword ptr [ebp-44]
'006a1e16    50                      push eax

' *** Reference to "__vbaObjSet"
'006a1e17    ff15b4104000            call dword ptr [004010b4]
    Set var_58 = Err
'006a1e1d    8bf0                    mov esi, eax
'006a1e1f    8b0e                    mov ecx, dword ptr [esi]
'006a1e21    8d55dc                  lea edx, dword ptr [ebp-24]
'006a1e24    52                      push edx
'006a1e25    56                      push esi
'006a1e26    ff512c                  call dword ptr [ecx+2c]
    var_39 = var_58.Description
'006a1e29    dbe2                    fnclex
'006a1e2b    85c0                    test ax, ax
'006a1e2d    7d0f                    jge 6a1e3e
'006a1e2f    6a2c                    push 2c
'006a1e31    685c084300              push 0043085c
'006a1e36    56                      push esi
'006a1e37    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006a1e38    ff1580104000            call dword ptr [00401080]
'006a1e3e    c785a0feffffe4084300    mov dword ptr [ebp+fffffea0], 004308e4
'006a1e48    89bd98feffff            mov dword ptr [ebp+fffffe98], edi
'006a1e4e    8b8510feffff            mov eax, dword ptr [ebp+fffffe10]
'006a1e54    898590feffff            mov dword ptr [ebp+fffffe90], eax
'006a1e5a    c78588feffff03000000    mov dword ptr [ebp+fffffe88], 00000003
'006a1e64    c78580feffff08094300    mov dword ptr [ebp+fffffe80], 00430908
'006a1e6e    89bd78feffff            mov dword ptr [ebp+fffffe78], edi
'006a1e74    8b45dc                  mov eax, dword ptr [ebp-24]
'006a1e77    c745dc00000000          mov dword ptr [ebp-24], 00000000
'006a1e7e    898530ffffff            mov dword ptr [ebp+ffffff30], eax
'006a1e84    89bd28ffffff            mov dword ptr [ebp+ffffff28], edi
'006a1e8a    c78570feffff08264500    mov dword ptr [ebp+fffffe70], 00452608
'006a1e94    89bd68feffff            mov dword ptr [ebp+fffffe68], edi
'006a1e9a    8d8d68ffffff            lea ecx, dword ptr [ebp+ffffff68]
'006a1ea0    51                      push ecx
'006a1ea1    8d9598feffff            lea edx, dword ptr [ebp+fffffe98]
'006a1ea7    52                      push edx
'006a1ea8    8d8558ffffff            lea eax, dword ptr [ebp+ffffff58]
'006a1eae    50                      push eax

' *** Reference to "__vbaVarCat"
'006a1eaf    8b3508124000            mov esi, dword ptr [00401208]
'006a1eb5    ffd6                    call esi
'006a1eb7    50                      push eax
'006a1eb8    8d8d88feffff            lea ecx, dword ptr [ebp+fffffe88]
'006a1ebe    51                      push ecx
'006a1ebf    8d9548ffffff            lea edx, dword ptr [ebp+ffffff48]
'006a1ec5    52                      push edx
'006a1ec6    ffd6                    call esi
'006a1ec8    50                      push eax
'006a1ec9    8d8578feffff            lea eax, dword ptr [ebp+fffffe78]
'006a1ecf    50                      push eax
'006a1ed0    8d8d38ffffff            lea ecx, dword ptr [ebp+ffffff38]
'006a1ed6    51                      push ecx
'006a1ed7    ffd6                    call esi
'006a1ed9    50                      push eax
'006a1eda    8d9528ffffff            lea edx, dword ptr [ebp+ffffff28]
'006a1ee0    52                      push edx
'006a1ee1    8d8518ffffff            lea eax, dword ptr [ebp+ffffff18]
'006a1ee7    50                      push eax
'006a1ee8    ffd6                    call esi
'006a1eea    50                      push eax
'006a1eeb    8d8d68feffff            lea ecx, dword ptr [ebp+fffffe68]
'006a1ef1    51                      push ecx
'006a1ef2    8d9508ffffff            lea edx, dword ptr [ebp+ffffff08]
'006a1ef8    52                      push edx
'006a1ef9    ffd6                    call esi
'006a1efb    50                      push eax
'006a1efc    33c0                    xor eax, eax
'006a1efe    66a12ab07200            mov eax, dword ptr [0072b02a]
'006a1f04    50                      push eax
'006a1f05    6884094300              push 00430984

' *** Reference to "__vbaPrintFile"
'006a1f0a    ff15b8114000            call dword ptr [004011b8]
    Print #0, 
'006a1f10    8d4dbc                  lea ecx, dword ptr [ebp-44]
'006a1f13    51                      push ecx
'006a1f14    8d55c0                  lea edx, dword ptr [ebp-40]
'006a1f17    52                      push edx
'006a1f18    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'006a1f1a    ff154c104000            call dword ptr [0040104c]
    '#Cleanup( var_5, var_58)
'006a1f20    8d8508ffffff            lea eax, dword ptr [ebp+ffffff08]
'006a1f26    50                      push eax
'006a1f27    8d8d18ffffff            lea ecx, dword ptr [ebp+ffffff18]
'006a1f2d    51                      push ecx
'006a1f2e    8d9528ffffff            lea edx, dword ptr [ebp+ffffff28]
'006a1f34    52                      push edx
'006a1f35    8d8538ffffff            lea eax, dword ptr [ebp+ffffff38]
'006a1f3b    50                      push eax
'006a1f3c    8d8d48ffffff            lea ecx, dword ptr [ebp+ffffff48]
'006a1f42    51                      push ecx
'006a1f43    8d9558ffffff            lea edx, dword ptr [ebp+ffffff58]
'006a1f49    52                      push edx
'006a1f4a    8d8568ffffff            lea eax, dword ptr [ebp+ffffff68]
'006a1f50    50                      push eax
'006a1f51    8d8d78ffffff            lea ecx, dword ptr [ebp+ffffff78]
'006a1f57    51                      push ecx
'006a1f58    8d5588                  lea edx, dword ptr [ebp-78]
'006a1f5b    52                      push edx
'006a1f5c    6a09                    push 09

' *** Reference to "__vbaFreeVarList"
'006a1f5e    ff1540104000            call dword ptr [00401040]
    '#Cleanup( var_131, var_87, var_132, var_133, var_134, var_135, var_136, var_849, var_310)
'006a1f64    83c440                  add esp, 40
'006a1f67    6a00                    push 00

' *** Reference to "__vbaResume"
'006a1f69    ff1568104000            call dword ptr [00401068]
'006a1f6f    e985070000              jmp 6a26f9
    Resume handler_6A26F9
End If

' *** Reference to "__vbaHresultCheckObj"
'006a1f74    8b3d80104000            mov edi, dword ptr [00401080]
'006a1f7a    8b45c0                  mov eax, dword ptr [ebp-40]
'006a1f7d    898504feffff            mov dword ptr [ebp+fffffe04], eax
'006a1f83    8b18                    mov ebx, dword ptr [eax]
'006a1f85    b90b000000              mov ecx, 0000000b

' *** Reference to "__vbaI2I4"
'006a1f8a    ff1550114000            call dword ptr [00401150]
'006a1f90    50                      push eax
'006a1f91    8bd3                    mov edx, ebx
'006a1f93    8b9d04feffff            mov ebx, dword ptr [ebp+fffffe04]
'006a1f99    53                      push ebx
'006a1f9a    ff527c                  call dword ptr [edx+7c]
'006a1f9d    dbe2                    fnclex
'006a1f9f    3bc6                    cmp eax, esi
'006a1fa1    7d0b                    jge 6a1fae
'006a1fa3    6a7c                    push 7c
'006a1fa5    6810be4300              push 0043be10
'006a1faa    53                      push ebx
'006a1fab    50                      push eax
'006a1fac    ffd7                    call edi
'006a1fae    8d4dc0                  lea ecx, dword ptr [ebp-40]

' *** Reference to "__vbaFreeObj"
'006a1fb1    8b3508134000            mov esi, dword ptr [00401308]
'006a1fb7    ffd6                    call esi
'#Cleanup(var_5)
'006a1fb9    b801000000              mov eax, 00000001
'006a1fbe    8985b0feffff            mov dword ptr [ebp+fffffeb0], eax
'006a1fc4    b903000000              mov ecx, 00000003
'006a1fc9    898da8feffff            mov dword ptr [ebp+fffffea8], ecx
'006a1fcf    83ec10                  sub esp, 10
'006a1fd2    8bd4                    mov edx, esp
'006a1fd4    890a                    mov dword ptr [edx], ecx
'006a1fd6    8b8dacfeffff            mov ecx, dword ptr [ebp+fffffeac]
'006a1fdc    894a04                  mov dword ptr [edx+04], ecx
'006a1fdf    894208                  mov dword ptr [edx+08], eax
'006a1fe2    8b85b4feffff            mov eax, dword ptr [ebp+fffffeb4]
'006a1fe8    89420c                  mov dword ptr [edx+0c], eax
'006a1feb    6a07                    push 07
'006a1fed    8b4508                  mov eax, dword ptr [ebp+08]
'006a1ff0    8b08                    mov ecx, dword ptr [eax]
'006a1ff2    50                      push eax
'006a1ff3    ff9114030000            call dword ptr [ecx+00000314]
'006a1ff9    50                      push eax
'006a1ffa    8d55c0                  lea edx, dword ptr [ebp-40]
'006a1ffd    52                      push edx

' *** Reference to "__vbaObjSet"
'006a1ffe    ff15b4104000            call dword ptr [004010b4]
Set var_5 = Me
'006a2004    50                      push eax

' *** Reference to "__vbaLateIdSt"
'006a2005    ff15ec124000            call dword ptr [004012ec]
var_5.[UNMANAGED] = 1
'006a200b    8d4dc0                  lea ecx, dword ptr [ebp-40]
'006a200e    ffd6                    call esi
'#Cleanup(var_5)
'006a2010    ba04000280              mov edx, 80020004
'006a2015    8bc2                    mov eax, edx
'006a2017    898590feffff            mov dword ptr [ebp+fffffe90], eax
'006a201d    b90a000000              mov ecx, 0000000a
'006a2022    898d88feffff            mov dword ptr [ebp+fffffe88], ecx
'006a2028    8995a0feffff            mov dword ptr [ebp+fffffea0], edx
'006a202e    898d98feffff            mov dword ptr [ebp+fffffe98], ecx
'006a2034    c785b0feffff02000000    mov dword ptr [ebp+fffffeb0], 00000002
'006a203e    c785a8feffff03000000    mov dword ptr [ebp+fffffea8], 00000003
'006a2048    8b3548b07200            mov esi, dword ptr [0072b048]
'006a204e    8b36                    mov esi, dword ptr [esi]
'006a2050    8d5dc0                  lea ebx, dword ptr [ebp-40]
'006a2053    53                      push ebx
'006a2054    83ec10                  sub esp, 10
'006a2057    8bdc                    mov ebx, esp
'006a2059    890b                    mov dword ptr [ebx], ecx
'006a205b    8b8d8cfeffff            mov ecx, dword ptr [ebp+fffffe8c]
'006a2061    894b04                  mov dword ptr [ebx+04], ecx
'006a2064    894308                  mov dword ptr [ebx+08], eax
'006a2067    8b8594feffff            mov eax, dword ptr [ebp+fffffe94]
'006a206d    89430c                  mov dword ptr [ebx+0c], eax
'006a2070    83ec10                  sub esp, 10
'006a2073    8bcc                    mov ecx, esp
'006a2075    8b8598feffff            mov eax, dword ptr [ebp+fffffe98]
'006a207b    8901                    mov dword ptr [ecx], eax
'006a207d    8b859cfeffff            mov eax, dword ptr [ebp+fffffe9c]
'006a2083    894104                  mov dword ptr [ecx+04], eax
'006a2086    895108                  mov dword ptr [ecx+08], edx
'006a2089    8b95a4feffff            mov edx, dword ptr [ebp+fffffea4]
'006a208f    89510c                  mov dword ptr [ecx+0c], edx
'006a2092    83ec10                  sub esp, 10
'006a2095    8bc4                    mov eax, esp
'006a2097    8b8da8feffff            mov ecx, dword ptr [ebp+fffffea8]
'006a209d    8908                    mov dword ptr [eax], ecx
'006a209f    8b95acfeffff            mov edx, dword ptr [ebp+fffffeac]
'006a20a5    895004                  mov dword ptr [eax+04], edx
'006a20a8    8b8db0feffff            mov ecx, dword ptr [ebp+fffffeb0]
'006a20ae    894808                  mov dword ptr [eax+08], ecx
'006a20b1    8b95b4feffff            mov edx, dword ptr [ebp+fffffeb4]
'006a20b7    89500c                  mov dword ptr [eax+0c], edx
'006a20ba    68e0274500              push 004527e0
'006a20bf    a148b07200              mov ax, word ptr [0072b048]
'006a20c4    50                      push eax
'006a20c5    ff96bc000000            call dword ptr [esi+000000bc]
'006a20cb    dbe2                    fnclex
'006a20cd    85c0                    test ax, ax
'006a20cf    7d14                    jge 6a20e5
'006a20d1    68bc000000              push 000000bc
'006a20d6    68301f4300              push 00431f30
'006a20db    8b0d48b07200            mov ecx, dword ptr [0072b048]
'006a20e1    51                      push ecx
'006a20e2    50                      push eax
'006a20e3    ffd7                    call edi
'006a20e5    8b45c0                  mov eax, dword ptr [ebp-40]
'006a20e8    c745c000000000          mov dword ptr [ebp-40], 00000000
'006a20ef    50                      push eax
'006a20f0    8d55e0                  lea edx, dword ptr [ebp-20]
'006a20f3    52                      push edx

' *** Reference to "__vbaObjSet"
'006a20f4    ff15b4104000            call dword ptr [004010b4]
Set var_3 = var_5

' *** Reference to "__vbaVarCat"
'006a20fa    8b3508124000            mov esi, dword ptr [00401208]
'006a2100    8b45e0                  mov eax, dword ptr [ebp-20]
'006a2103    8b08                    mov ecx, dword ptr [eax]
'006a2105    8d9514feffff            lea edx, dword ptr [ebp+fffffe14]
'006a210b    52                      push edx
'006a210c    50                      push eax
'006a210d    ff5134                  call dword ptr [ecx+34]
'006a2110    dbe2                    fnclex
'006a2112    85c0                    test ax, ax
'006a2114    7d0e                    jge 6a2124
'006a2116    6a34                    push 34
'006a2118    6830314300              push 00433130
'006a211d    8b4de0                  mov ecx, dword ptr [ebp-20]
'006a2120    51                      push ecx
'006a2121    50                      push eax
'006a2122    ffd7                    call edi
'006a2124    6683bd14feffff00        cmp word ptr [ebp+fffffe14], 00
'006a212c    8b45e0                  mov eax, dword ptr [ebp-20]
'006a212f    0f853e050000            jne 6a2673

Do While (0 = 0)
'006a2135    8b10                    mov edx, dword ptr [eax]
'006a2137    8d4dc0                  lea ecx, dword ptr [ebp-40]
'006a213a    51                      push ecx
'006a213b    50                      push eax
'006a213c    ff92b4000000            call dword ptr [edx+000000b4]
'006a2142    dbe2                    fnclex
'006a2144    85c0                    test ax, ax
'006a2146    7d11                    jge 6a2159
'006a2148    68b4000000              push 000000b4
'006a214d    6830314300              push 00433130
'006a2152    8b55e0                  mov edx, dword ptr [ebp-20]
'006a2155    52                      push edx
'006a2156    50                      push eax
'006a2157    ffd7                    call edi
'006a2159    8b45c0                  mov eax, dword ptr [ebp-40]
'006a215c    898508feffff            mov dword ptr [ebp+fffffe08], eax
'006a2162    c785b0feffffc4a54300    mov dword ptr [ebp+fffffeb0], 0043a5c4
'006a216c    b908000000              mov ecx, 00000008
'006a2171    898da8feffff            mov dword ptr [ebp+fffffea8], ecx
'006a2177    8b10                    mov edx, dword ptr [eax]
'006a2179    8d5dbc                  lea ebx, dword ptr [ebp-44]
'006a217c    53                      push ebx
'006a217d    83ec10                  sub esp, 10
'006a2180    8bdc                    mov ebx, esp
'006a2182    890b                    mov dword ptr [ebx], ecx
'006a2184    8b8dacfeffff            mov ecx, dword ptr [ebp+fffffeac]
'006a218a    894b04                  mov dword ptr [ebx+04], ecx
'006a218d    8b8db0feffff            mov ecx, dword ptr [ebp+fffffeb0]
'006a2193    894b08                  mov dword ptr [ebx+08], ecx
'006a2196    8b8db4feffff            mov ecx, dword ptr [ebp+fffffeb4]
'006a219c    894b0c                  mov dword ptr [ebx+0c], ecx
'006a219f    50                      push eax
'006a21a0    ff5230                  call dword ptr [edx+30]
'006a21a3    dbe2                    fnclex
'006a21a5    85c0                    test ax, ax
'006a21a7    7d11                    jge 6a21ba
'006a21a9    6a30                    push 30
'006a21ab    68d8304300              push 004330d8
'006a21b0    8b9508feffff            mov edx, dword ptr [ebp+fffffe08]
'006a21b6    52                      push edx
'006a21b7    50                      push eax
'006a21b8    ffd7                    call edi
'006a21ba    8b45bc                  mov eax, dword ptr [ebp-44]
'006a21bd    8bd8                    mov ebx, eax
'006a21bf    8b08                    mov ecx, dword ptr [eax]
'006a21c1    8d5588                  lea edx, dword ptr [ebp-78]
'006a21c4    52                      push edx
'006a21c5    50                      push eax
'006a21c6    ff5144                  call dword ptr [ecx+44]
    Call var_58.Raise(-376, var_827, -256 - 20, , -256 - 20)
'006a21c9    dbe2                    fnclex
'006a21cb    85c0                    test ax, ax
'006a21cd    7d0b                    jge 6a21da
'006a21cf    6a44                    push 44
'006a21d1    6880304300              push 00433080
'006a21d6    53                      push ebx
'006a21d7    50                      push eax
'006a21d8    ffd7                    call edi
'006a21da    c785a0feffff44ed4300    mov dword ptr [ebp+fffffea0], 0043ed44
'006a21e4    c78598feffff08000000    mov dword ptr [ebp+fffffe98], 00000008
'006a21ee    8b45e0                  mov eax, dword ptr [ebp-20]
'006a21f1    8b08                    mov ecx, dword ptr [eax]
'006a21f3    8d55b8                  lea edx, dword ptr [ebp-48]
'006a21f6    52                      push edx
'006a21f7    50                      push eax
'006a21f8    ff91b4000000            call dword ptr [ecx+000000b4]
'006a21fe    dbe2                    fnclex
'006a2200    85c0                    test ax, ax
'006a2202    7d11                    jge 6a2215
'006a2204    68b4000000              push 000000b4
'006a2209    6830314300              push 00433130
'006a220e    8b4de0                  mov ecx, dword ptr [ebp-20]
'006a2211    51                      push ecx
'006a2212    50                      push eax
'006a2213    ffd7                    call edi
'006a2215    8b45b8                  mov eax, dword ptr [ebp-48]
'006a2218    8985f4fdffff            mov dword ptr [ebp+fffffdf4], eax
'006a221e    c78590feffff20a74300    mov dword ptr [ebp+fffffe90], 0043a720
'006a2228    b908000000              mov ecx, 00000008
'006a222d    898d88feffff            mov dword ptr [ebp+fffffe88], ecx
'006a2233    8b10                    mov edx, dword ptr [eax]
'006a2235    8d5db4                  lea ebx, dword ptr [ebp-4c]
'006a2238    53                      push ebx
'006a2239    83ec10                  sub esp, 10
'006a223c    8bdc                    mov ebx, esp
'006a223e    890b                    mov dword ptr [ebx], ecx
'006a2240    8b8d8cfeffff            mov ecx, dword ptr [ebp+fffffe8c]
'006a2246    894b04                  mov dword ptr [ebx+04], ecx
'006a2249    8b8d90feffff            mov ecx, dword ptr [ebp+fffffe90]
'006a224f    894b08                  mov dword ptr [ebx+08], ecx
'006a2252    8b8d94feffff            mov ecx, dword ptr [ebp+fffffe94]
'006a2258    894b0c                  mov dword ptr [ebx+0c], ecx
'006a225b    50                      push eax
'006a225c    ff5230                  call dword ptr [edx+30]
'006a225f    dbe2                    fnclex
'006a2261    85c0                    test ax, ax
'006a2263    7d11                    jge 6a2276
'006a2265    6a30                    push 30
'006a2267    68d8304300              push 004330d8
'006a226c    8b95f4fdffff            mov edx, dword ptr [ebp+fffffdf4]
'006a2272    52                      push edx
'006a2273    50                      push eax
'006a2274    ffd7                    call edi
'006a2276    8b45b4                  mov eax, dword ptr [ebp-4c]
'006a2279    8bd8                    mov ebx, eax
'006a227b    8b08                    mov ecx, dword ptr [eax]
'006a227d    8d9568ffffff            lea edx, dword ptr [ebp+ffffff68]
'006a2283    52                      push edx
'006a2284    50                      push eax
'006a2285    ff5144                  call dword ptr [ecx+44]
'006a2288    dbe2                    fnclex
'006a228a    85c0                    test ax, ax
'006a228c    7d0b                    jge 6a2299
'006a228e    6a44                    push 44
'006a2290    6880304300              push 00433080
'006a2295    53                      push ebx
'006a2296    50                      push eax
'006a2297    ffd7                    call edi
'006a2299    c78580feffff44ed4300    mov dword ptr [ebp+fffffe80], 0043ed44
'006a22a3    c78578feffff08000000    mov dword ptr [ebp+fffffe78], 00000008
'006a22ad    8b45e0                  mov eax, dword ptr [ebp-20]
'006a22b0    8b08                    mov ecx, dword ptr [eax]
'006a22b2    8d55b0                  lea edx, dword ptr [ebp-50]
'006a22b5    52                      push edx
'006a22b6    50                      push eax
'006a22b7    ff91b4000000            call dword ptr [ecx+000000b4]
'006a22bd    dbe2                    fnclex
'006a22bf    85c0                    test ax, ax
'006a22c1    7d11                    jge 6a22d4
'006a22c3    68b4000000              push 000000b4
'006a22c8    6830314300              push 00433130
'006a22cd    8b4de0                  mov ecx, dword ptr [ebp-20]
'006a22d0    51                      push ecx
'006a22d1    50                      push eax
'006a22d2    ffd7                    call edi
'006a22d4    8b45b0                  mov eax, dword ptr [ebp-50]
'006a22d7    8985e0fdffff            mov dword ptr [ebp+fffffde0], eax
'006a22dd    c78570feffffdca54300    mov dword ptr [ebp+fffffe70], 0043a5dc
'006a22e7    b908000000              mov ecx, 00000008
'006a22ec    898d68feffff            mov dword ptr [ebp+fffffe68], ecx
'006a22f2    8b10                    mov edx, dword ptr [eax]
'006a22f4    8d5dac                  lea ebx, dword ptr [ebp-54]
'006a22f7    53                      push ebx
'006a22f8    83ec10                  sub esp, 10
'006a22fb    8bdc                    mov ebx, esp
'006a22fd    890b                    mov dword ptr [ebx], ecx
'006a22ff    8b8d6cfeffff            mov ecx, dword ptr [ebp+fffffe6c]
'006a2305    894b04                  mov dword ptr [ebx+04], ecx
'006a2308    8b8d70feffff            mov ecx, dword ptr [ebp+fffffe70]
'006a230e    894b08                  mov dword ptr [ebx+08], ecx
'006a2311    8b8d74feffff            mov ecx, dword ptr [ebp+fffffe74]
'006a2317    894b0c                  mov dword ptr [ebx+0c], ecx
'006a231a    50                      push eax
'006a231b    ff5230                  call dword ptr [edx+30]
'006a231e    dbe2                    fnclex
'006a2320    85c0                    test ax, ax
'006a2322    7d11                    jge 6a2335
'006a2324    6a30                    push 30
'006a2326    68d8304300              push 004330d8
'006a232b    8b95e0fdffff            mov edx, dword ptr [ebp+fffffde0]
'006a2331    52                      push edx
'006a2332    50                      push eax
'006a2333    ffd7                    call edi
'006a2335    8b45ac                  mov eax, dword ptr [ebp-54]
'006a2338    8bd8                    mov ebx, eax
'006a233a    8b08                    mov ecx, dword ptr [eax]
'006a233c    8d9538ffffff            lea edx, dword ptr [ebp+ffffff38]
'006a2342    52                      push edx
'006a2343    50                      push eax
'006a2344    ff5144                  call dword ptr [ecx+44]
'006a2347    dbe2                    fnclex
'006a2349    85c0                    test ax, ax
'006a234b    7d0b                    jge 6a2358
'006a234d    6a44                    push 44
'006a234f    6880304300              push 00433080
'006a2354    53                      push ebx
'006a2355    50                      push eax
'006a2356    ffd7                    call edi
'006a2358    c78560feffff44ed4300    mov dword ptr [ebp+fffffe60], 0043ed44
'006a2362    c78558feffff08000000    mov dword ptr [ebp+fffffe58], 00000008
'006a236c    8b45e0                  mov eax, dword ptr [ebp-20]
'006a236f    8b08                    mov ecx, dword ptr [eax]
'006a2371    8d55a8                  lea edx, dword ptr [ebp-58]
'006a2374    52                      push edx
'006a2375    50                      push eax
'006a2376    ff91b4000000            call dword ptr [ecx+000000b4]
'006a237c    dbe2                    fnclex
'006a237e    85c0                    test ax, ax
'006a2380    7d11                    jge 6a2393
'006a2382    68b4000000              push 000000b4
'006a2387    6830314300              push 00433130
'006a238c    8b4de0                  mov ecx, dword ptr [ebp-20]
'006a238f    51                      push ecx
'006a2390    50                      push eax
'006a2391    ffd7                    call edi
'006a2393    8b45a8                  mov eax, dword ptr [ebp-58]
'006a2396    8985ccfdffff            mov dword ptr [ebp+fffffdcc], eax
'006a239c    b908000000              mov ecx, 00000008
'006a23a1    8b10                    mov edx, dword ptr [eax]
'006a23a3    8d5da4                  lea ebx, dword ptr [ebp-5c]
'006a23a6    53                      push ebx
'006a23a7    83ec10                  sub esp, 10
'006a23aa    8bdc                    mov ebx, esp
'006a23ac    890b                    mov dword ptr [ebx], ecx
'006a23ae    8b8d4cfeffff            mov ecx, dword ptr [ebp+fffffe4c]
'006a23b4    894b04                  mov dword ptr [ebx+04], ecx
'006a23b7    b94ca84300              mov ecx, 0043a84c
'006a23bc    894b08                  mov dword ptr [ebx+08], ecx
'006a23bf    8b8d54feffff            mov ecx, dword ptr [ebp+fffffe54]
'006a23c5    894b0c                  mov dword ptr [ebx+0c], ecx
'006a23c8    50                      push eax
'006a23c9    ff5230                  call dword ptr [edx+30]
'006a23cc    dbe2                    fnclex
'006a23ce    85c0                    test ax, ax
'006a23d0    7d11                    jge 6a23e3
'006a23d2    6a30                    push 30
'006a23d4    68d8304300              push 004330d8
'006a23d9    8b95ccfdffff            mov edx, dword ptr [ebp+fffffdcc]
'006a23df    52                      push edx
'006a23e0    50                      push eax
'006a23e1    ffd7                    call edi
'006a23e3    8b45a4                  mov eax, dword ptr [ebp-5c]
'006a23e6    8bd8                    mov ebx, eax
'006a23e8    8b08                    mov ecx, dword ptr [eax]
'006a23ea    8d9508ffffff            lea edx, dword ptr [ebp+ffffff08]
'006a23f0    52                      push edx
'006a23f1    50                      push eax
'006a23f2    ff5144                  call dword ptr [ecx+44]
'006a23f5    dbe2                    fnclex
'006a23f7    85c0                    test ax, ax
'006a23f9    7d0b                    jge 6a2406
'006a23fb    6a44                    push 44
'006a23fd    6880304300              push 00433080
'006a2402    53                      push ebx
'006a2403    50                      push eax
'006a2404    ffd7                    call edi
'006a2406    c78540feffff44ed4300    mov dword ptr [ebp+fffffe40], 0043ed44
'006a2410    c78538feffff08000000    mov dword ptr [ebp+fffffe38], 00000008
'006a241a    8b45e0                  mov eax, dword ptr [ebp-20]
'006a241d    8b08                    mov ecx, dword ptr [eax]
'006a241f    8d55a0                  lea edx, dword ptr [ebp-60]
'006a2422    52                      push edx
'006a2423    50                      push eax
'006a2424    ff91b4000000            call dword ptr [ecx+000000b4]
'006a242a    dbe2                    fnclex
'006a242c    85c0                    test ax, ax
'006a242e    7d11                    jge 6a2441
'006a2430    68b4000000              push 000000b4
'006a2435    6830314300              push 00433130
'006a243a    8b4de0                  mov ecx, dword ptr [ebp-20]
'006a243d    51                      push ecx
'006a243e    50                      push eax
'006a243f    ffd7                    call edi
'006a2441    8b45a0                  mov eax, dword ptr [ebp-60]
'006a2444    8985b8fdffff            mov dword ptr [ebp+fffffdb8], eax
'006a244a    b9cca64300              mov ecx, 0043a6cc
'006a244f    ba08000000              mov edx, 00000008
'006a2454    8b38                    mov edi, dword ptr [eax]
'006a2456    8d5d9c                  lea ebx, dword ptr [ebp-64]
'006a2459    53                      push ebx
'006a245a    83ec10                  sub esp, 10
'006a245d    8bdc                    mov ebx, esp
'006a245f    8913                    mov dword ptr [ebx], edx
'006a2461    8b952cfeffff            mov edx, dword ptr [ebp+fffffe2c]
'006a2467    895304                  mov dword ptr [ebx+04], edx
'006a246a    894b08                  mov dword ptr [ebx+08], ecx
'006a246d    8b8d34feffff            mov ecx, dword ptr [ebp+fffffe34]
'006a2473    894b0c                  mov dword ptr [ebx+0c], ecx
'006a2476    50                      push eax
'006a2477    ff5730                  call dword ptr [edi+30]
'006a247a    dbe2                    fnclex
'006a247c    85c0                    test ax, ax
'006a247e    7d19                    jge 6a2499
'006a2480    6a30                    push 30
'006a2482    68d8304300              push 004330d8
'006a2487    8b95b8fdffff            mov edx, dword ptr [ebp+fffffdb8]
'006a248d    52                      push edx
'006a248e    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006a248f    8b3d80104000            mov edi, dword ptr [00401080]
'006a2495    ffd7                    call edi
'006a2497    eb06                    jmp 6a249f

' *** Reference to "__vbaHresultCheckObj"
'006a2499    8b3d80104000            mov edi, dword ptr [00401080]
'006a249f    8b459c                  mov eax, dword ptr [ebp-64]
'006a24a2    8bd8                    mov ebx, eax
'006a24a4    8b08                    mov ecx, dword ptr [eax]
'006a24a6    8d95d8feffff            lea edx, dword ptr [ebp+fffffed8]
'006a24ac    52                      push edx
'006a24ad    50                      push eax
'006a24ae    ff5144                  call dword ptr [ecx+44]
'006a24b1    dbe2                    fnclex
'006a24b3    85c0                    test ax, ax
'006a24b5    7d0b                    jge 6a24c2
'006a24b7    6a44                    push 44
'006a24b9    6880304300              push 00433080
'006a24be    53                      push ebx
'006a24bf    50                      push eax
'006a24c0    ffd7                    call edi
'006a24c2    8d4588                  lea eax, dword ptr [ebp-78]
'006a24c5    50                      push eax
'006a24c6    8d8d98feffff            lea ecx, dword ptr [ebp+fffffe98]
'006a24cc    51                      push ecx
'006a24cd    8d9578ffffff            lea edx, dword ptr [ebp+ffffff78]
'006a24d3    52                      push edx
'006a24d4    ffd6                    call esi
'006a24d6    50                      push eax
'006a24d7    8d8568ffffff            lea eax, dword ptr [ebp+ffffff68]
'006a24dd    50                      push eax
'006a24de    8d8d58ffffff            lea ecx, dword ptr [ebp+ffffff58]
'006a24e4    51                      push ecx
'006a24e5    ffd6                    call esi
'006a24e7    50                      push eax
'006a24e8    8d9578feffff            lea edx, dword ptr [ebp+fffffe78]
'006a24ee    52                      push edx
'006a24ef    8d8548ffffff            lea eax, dword ptr [ebp+ffffff48]
'006a24f5    50                      push eax
'006a24f6    ffd6                    call esi
'006a24f8    50                      push eax
'006a24f9    8d8d38ffffff            lea ecx, dword ptr [ebp+ffffff38]
'006a24ff    51                      push ecx
'006a2500    8d9528ffffff            lea edx, dword ptr [ebp+ffffff28]
'006a2506    52                      push edx
'006a2507    ffd6                    call esi
'006a2509    50                      push eax
'006a250a    8d8558feffff            lea eax, dword ptr [ebp+fffffe58]
'006a2510    50                      push eax
'006a2511    8d8d18ffffff            lea ecx, dword ptr [ebp+ffffff18]
'006a2517    51                      push ecx
'006a2518    ffd6                    call esi
'006a251a    50                      push eax
'006a251b    8d9508ffffff            lea edx, dword ptr [ebp+ffffff08]
'006a2521    52                      push edx
'006a2522    8d85f8feffff            lea eax, dword ptr [ebp+fffffef8]
'006a2528    50                      push eax
'006a2529    ffd6                    call esi
'006a252b    50                      push eax
'006a252c    8d8d38feffff            lea ecx, dword ptr [ebp+fffffe38]
'006a2532    51                      push ecx
'006a2533    8d95e8feffff            lea edx, dword ptr [ebp+fffffee8]
'006a2539    52                      push edx
'006a253a    ffd6                    call esi
'006a253c    50                      push eax
'006a253d    8d85d8feffff            lea eax, dword ptr [ebp+fffffed8]
'006a2543    50                      push eax
'006a2544    8d8dc8feffff            lea ecx, dword ptr [ebp+fffffec8]
'006a254a    51                      push ecx
'006a254b    ffd6                    call esi
'006a254d    50                      push eax

' *** Reference to "__vbaStrVarMove"
'006a254e    ff153c104000            call dword ptr [0040103c]
'006a2554    8985c0feffff            mov dword ptr [ebp+fffffec0], eax
'006a255a    b908000000              mov ecx, 00000008
'006a255f    898db8feffff            mov dword ptr [ebp+fffffeb8], ecx
'006a2565    83ec10                  sub esp, 10
'006a2568    8bd4                    mov edx, esp
'006a256a    890a                    mov dword ptr [edx], ecx
'006a256c    8b8dbcfeffff            mov ecx, dword ptr [ebp+fffffebc]
'006a2572    894a04                  mov dword ptr [edx+04], ecx
'006a2575    894208                  mov dword ptr [edx+08], eax
'006a2578    8b85c4feffff            mov eax, dword ptr [ebp+fffffec4]
'006a257e    89420c                  mov dword ptr [edx+0c], eax
'006a2581    6a01                    push 01
'006a2583    6880000000              push 00000080
'006a2588    8b4508                  mov eax, dword ptr [ebp+08]
'006a258b    8b08                    mov ecx, dword ptr [eax]
'006a258d    50                      push eax
'006a258e    ff9114030000            call dword ptr [ecx+00000314]
'006a2594    50                      push eax
'006a2595    8d5598                  lea edx, dword ptr [ebp-68]
'006a2598    52                      push edx

' *** Reference to "__vbaObjSet"
'006a2599    ff15b4104000            call dword ptr [004010b4]
    Set var_130 = Me
'006a259f    50                      push eax

' *** Reference to "__vbaLateIdCall"
'006a25a0    ff1538104000            call dword ptr [00401038]
    Call FrmLstArticle.Member_0x80(#NOT SUPPORTED#)
'006a25a6    8d4598                  lea eax, dword ptr [ebp-68]
'006a25a9    50                      push eax
'006a25aa    8d4d9c                  lea ecx, dword ptr [ebp-64]
'006a25ad    51                      push ecx
'006a25ae    8d55a0                  lea edx, dword ptr [ebp-60]
'006a25b1    52                      push edx
'006a25b2    8d45a4                  lea eax, dword ptr [ebp-5c]
'006a25b5    50                      push eax
'006a25b6    8d4da8                  lea ecx, dword ptr [ebp-58]
'006a25b9    51                      push ecx
'006a25ba    8d55ac                  lea edx, dword ptr [ebp-54]
'006a25bd    52                      push edx
'006a25be    8d45b0                  lea eax, dword ptr [ebp-50]
'006a25c1    50                      push eax
'006a25c2    8d4db4                  lea ecx, dword ptr [ebp-4c]
'006a25c5    51                      push ecx
'006a25c6    8d55b8                  lea edx, dword ptr [ebp-48]
'006a25c9    52                      push edx
'006a25ca    8d45bc                  lea eax, dword ptr [ebp-44]
'006a25cd    50                      push eax
'006a25ce    8d4dc0                  lea ecx, dword ptr [ebp-40]
'006a25d1    51                      push ecx
'006a25d2    6a0b                    push 0b

' *** Reference to "__vbaFreeObjList"
'006a25d4    ff154c104000            call dword ptr [0040104c]
    '#Cleanup( var_5, var_58, var_61, var_62, var_6, var_50, var_86, var_83, var_7, var_51, var_130)
'006a25da    83c44c                  add esp, 4c
'006a25dd    8d95b8feffff            lea edx, dword ptr [ebp+fffffeb8]
'006a25e3    52                      push edx
'006a25e4    8d85c8feffff            lea eax, dword ptr [ebp+fffffec8]
'006a25ea    50                      push eax
'006a25eb    8d8dd8feffff            lea ecx, dword ptr [ebp+fffffed8]
'006a25f1    51                      push ecx
'006a25f2    8d95e8feffff            lea edx, dword ptr [ebp+fffffee8]
'006a25f8    52                      push edx
'006a25f9    8d85f8feffff            lea eax, dword ptr [ebp+fffffef8]
'006a25ff    50                      push eax
'006a2600    8d8d08ffffff            lea ecx, dword ptr [ebp+ffffff08]
'006a2606    51                      push ecx
'006a2607    8d9518ffffff            lea edx, dword ptr [ebp+ffffff18]
'006a260d    52                      push edx
'006a260e    8d8528ffffff            lea eax, dword ptr [ebp+ffffff28]
'006a2614    50                      push eax
'006a2615    8d8d38ffffff            lea ecx, dword ptr [ebp+ffffff38]
'006a261b    51                      push ecx
'006a261c    8d9548ffffff            lea edx, dword ptr [ebp+ffffff48]
'006a2622    52                      push edx
'006a2623    8d8558ffffff            lea eax, dword ptr [ebp+ffffff58]
'006a2629    50                      push eax
'006a262a    8d8d68ffffff            lea ecx, dword ptr [ebp+ffffff68]
'006a2630    51                      push ecx
'006a2631    8d9578ffffff            lea edx, dword ptr [ebp+ffffff78]
'006a2637    52                      push edx
'006a2638    8d4588                  lea eax, dword ptr [ebp-78]
'006a263b    50                      push eax
'006a263c    6a0e                    push 0e

' *** Reference to "__vbaFreeVarList"
'006a263e    ff1540104000            call dword ptr [00401040]
    '#Cleanup( var_131, var_87, var_132, var_133, var_134, var_135, var_136, var_849, var_310, var_298, var_539, var_25, var_65, var_767)
'006a2644    83c43c                  add esp, 3c
'006a2647    8b45e0                  mov eax, dword ptr [ebp-20]
'006a264a    8b08                    mov ecx, dword ptr [eax]
'006a264c    50                      push eax
'006a264d    ff91ec000000            call dword ptr [ecx+000000ec]
'006a2653    dbe2                    fnclex
'006a2655    85c0                    test ax, ax
'006a2657    0f8da3faffff            jge 6a2100
'006a265d    68ec000000              push 000000ec
'006a2662    6830314300              push 00433130
'006a2667    8b55e0                  mov edx, dword ptr [ebp-20]
'006a266a    52                      push edx
'006a266b    50                      push eax
'006a266c    ffd7                    call edi
'006a266e    e98dfaffff              jmp 6a2100
    
Loop
'006a2673    8b08                    mov ecx, dword ptr [eax]
'006a2675    50                      push eax
'006a2676    ff91c4000000            call dword ptr [ecx+000000c4]
'006a267c    dbe2                    fnclex
'006a267e    85c0                    test ax, ax
'006a2680    7d11                    jge 6a2693
'006a2682    68c4000000              push 000000c4
'006a2687    6830314300              push 00433130
'006a268c    8b55e0                  mov edx, dword ptr [ebp-20]
'006a268f    52                      push edx
'006a2690    50                      push eax
'006a2691    ffd7                    call edi
'006a2693    a124be7200              mov ax, word ptr [0072be24]
'006a2698    85c0                    test ax, ax
'006a269a    7510                    jne 6a26ac
'006a269c    6824be7200              push 0072be24
'006a26a1    6870004300              push 00430070

' *** Reference to "__vbaNew2"
'006a26a6    ff152c124000            call dword ptr [0040122c]
Set var_2 = New Global
'006a26ac    8b3524be7200            mov esi, dword ptr [0072be24]
'006a26b2    8b06                    mov eax, dword ptr [esi]
'006a26b4    8d4dc0                  lea ecx, dword ptr [ebp-40]
'006a26b7    51                      push ecx
'006a26b8    56                      push esi
'006a26b9    ff5018                  call dword ptr [eax+18]
Set var_5 = var_2.Screen
'006a26bc    dbe2                    fnclex
'006a26be    85c0                    test ax, ax
'006a26c0    7d0b                    jge 6a26cd
'006a26c2    6a18                    push 18
'006a26c4    6860004300              push 00430060
'006a26c9    56                      push esi
'006a26ca    50                      push eax
'006a26cb    ffd7                    call edi
'006a26cd    8b75c0                  mov esi, dword ptr [ebp-40]
'006a26d0    8b1e                    mov ebx, dword ptr [esi]
'006a26d2    33c9                    xor ecx, ecx

' *** Reference to "__vbaI2I4"
'006a26d4    ff1550114000            call dword ptr [00401150]
'006a26da    50                      push eax
'006a26db    56                      push esi
'006a26dc    ff537c                  call dword ptr [ebx+7c]
var_5.MousePointer = CInt(0)
'006a26df    dbe2                    fnclex
'006a26e1    85c0                    test ax, ax
'006a26e3    7d0b                    jge 6a26f0
'006a26e5    6a7c                    push 7c
'006a26e7    6810be4300              push 0043be10
'006a26ec    56                      push esi
'006a26ed    50                      push eax
'006a26ee    ffd7                    call edi
'006a26f0    8d4dc0                  lea ecx, dword ptr [ebp-40]

' *** Reference to "__vbaFreeObj"
'006a26f3    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_5)

' *** Reference to "__vbaExitProc"
'006a26f9    ff15a0104000            call dword ptr [004010a0]
'006a26ff    68d9276a00              push 006a27d9
'006a2704    e9c6000000              jmp 6a27cf
'006a2709    8d55c4                  lea edx, dword ptr [ebp-3c]
'006a270c    52                      push edx
'006a270d    8d45c8                  lea eax, dword ptr [ebp-38]
'006a2710    50                      push eax
'006a2711    8d4dcc                  lea ecx, dword ptr [ebp-34]
'006a2714    51                      push ecx
'006a2715    8d55d0                  lea edx, dword ptr [ebp-30]
'006a2718    52                      push edx
'006a2719    8d45d4                  lea eax, dword ptr [ebp-2c]
'006a271c    50                      push eax
'006a271d    8d4dd8                  lea ecx, dword ptr [ebp-28]
'006a2720    51                      push ecx
'006a2721    8d55dc                  lea edx, dword ptr [ebp-24]
'006a2724    52                      push edx
'006a2725    6a07                    push 07

' *** Reference to "__vbaFreeStrList"
'006a2727    ff155c124000            call dword ptr [0040125c]
'#Cleanup( var_39, -4516, -4520, -4524, -4528, -4532, -4536)
'006a272d    8d4598                  lea eax, dword ptr [ebp-68]
'006a2730    50                      push eax
'006a2731    8d4d9c                  lea ecx, dword ptr [ebp-64]
'006a2734    51                      push ecx
'006a2735    8d55a0                  lea edx, dword ptr [ebp-60]
'006a2738    52                      push edx
'006a2739    8d45a4                  lea eax, dword ptr [ebp-5c]
'006a273c    50                      push eax
'006a273d    8d4da8                  lea ecx, dword ptr [ebp-58]
'006a2740    51                      push ecx
'006a2741    8d55ac                  lea edx, dword ptr [ebp-54]
'006a2744    52                      push edx
'006a2745    8d45b0                  lea eax, dword ptr [ebp-50]
'006a2748    50                      push eax
'006a2749    8d4db4                  lea ecx, dword ptr [ebp-4c]
'006a274c    51                      push ecx
'006a274d    8d55b8                  lea edx, dword ptr [ebp-48]
'006a2750    52                      push edx
'006a2751    8d45bc                  lea eax, dword ptr [ebp-44]
'006a2754    50                      push eax
'006a2755    8d4dc0                  lea ecx, dword ptr [ebp-40]
'006a2758    51                      push ecx
'006a2759    6a0b                    push 0b

' *** Reference to "__vbaFreeObjList"
'006a275b    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_5, var_58, var_61, var_62, var_6, var_50, var_86, var_83, var_7, var_51, var_130)
'006a2761    83c450                  add esp, 50
'006a2764    8d95b8feffff            lea edx, dword ptr [ebp+fffffeb8]
'006a276a    52                      push edx
'006a276b    8d85c8feffff            lea eax, dword ptr [ebp+fffffec8]
'006a2771    50                      push eax
'006a2772    8d8dd8feffff            lea ecx, dword ptr [ebp+fffffed8]
'006a2778    51                      push ecx
'006a2779    8d95e8feffff            lea edx, dword ptr [ebp+fffffee8]
'006a277f    52                      push edx
'006a2780    8d85f8feffff            lea eax, dword ptr [ebp+fffffef8]
'006a2786    50                      push eax
'006a2787    8d8d08ffffff            lea ecx, dword ptr [ebp+ffffff08]
'006a278d    51                      push ecx
'006a278e    8d9518ffffff            lea edx, dword ptr [ebp+ffffff18]
'006a2794    52                      push edx
'006a2795    8d8528ffffff            lea eax, dword ptr [ebp+ffffff28]
'006a279b    50                      push eax
'006a279c    8d8d38ffffff            lea ecx, dword ptr [ebp+ffffff38]
'006a27a2    51                      push ecx
'006a27a3    8d9548ffffff            lea edx, dword ptr [ebp+ffffff48]
'006a27a9    52                      push edx
'006a27aa    8d8558ffffff            lea eax, dword ptr [ebp+ffffff58]
'006a27b0    50                      push eax
'006a27b1    8d8d68ffffff            lea ecx, dword ptr [ebp+ffffff68]
'006a27b7    51                      push ecx
'006a27b8    8d9578ffffff            lea edx, dword ptr [ebp+ffffff78]
'006a27be    52                      push edx
'006a27bf    8d4588                  lea eax, dword ptr [ebp-78]
'006a27c2    50                      push eax
'006a27c3    6a0e                    push 0e

' *** Reference to "__vbaFreeVarList"
'006a27c5    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_131, var_87, var_132, var_133, var_134, var_135, var_136, var_849, var_310, var_298, var_539, var_25, var_65, var_767)
'006a27cb    83c43c                  add esp, 3c
'006a27ce    c3                      ret
'006a27cf    8d4de0                  lea ecx, dword ptr [ebp-20]

' *** Reference to "__vbaFreeObj"
'006a27d2    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_3)
'006a27d8    c3                      ret
'006a27d9    8b4508                  mov eax, dword ptr [ebp+08]
'006a27dc    8b08                    mov ecx, dword ptr [eax]
'006a27de    50                      push eax
'006a27df    ff5108                  call dword ptr [ecx+08]
'006a27e2    8b45f4                  mov eax, dword ptr [ebp-0c]
'006a27e5    8b4de4                  mov ecx, dword ptr [ebp-1c]
    'Reference to '__except_list'
'006a27e8    64890d00000000          mov dword ptr fs:[00000000], ecx
'006a27ef    5f                      pop edi
'006a27f0    5e                      pop esi
'006a27f1    5b                      pop ebx
'006a27f2    8be5                    mov esp, ebp
'006a27f4    5d                      pop ebp
'006a27f5    c20400                  ret 0004


End Function


'Event for FGarticle
Private Sub FGarticle_Event10()
'006a0cb0    55                      push ebp
'006a0cb1    8bec                    mov ebp, esp
'006a0cb3    83ec14                  sub esp, 14

' *** Reference to "__vbaExceptHandler"
'006a0cb6    6866724000              push 00407266
'006a0cbb    64a100000000            mov ax, word ptr fs:[00000000]
'006a0cc1    50                      push eax
    'Reference to '__except_list'
'006a0cc2    64892500000000          mov dword ptr fs:[00000000], esp
'006a0cc9    81ec2c010000            sub esp, 0000012c
'006a0ccf    53                      push ebx
'006a0cd0    56                      push esi
'006a0cd1    57                      push edi
'006a0cd2    8965ec                  mov dword ptr [ebp-14], esp
'006a0cd5    c745f018624000          mov dword ptr [ebp-10], 00406218
'006a0cdc    8b7d08                  mov edi, dword ptr [ebp+08]
'006a0cdf    8bc7                    mov eax, edi
'006a0ce1    83e001                  and eax, 01
'006a0ce4    8945f4                  mov dword ptr [ebp-0c], eax
'006a0ce7    83e7fe                  and edi, fffffffe
'006a0cea    897d08                  mov dword ptr [ebp+08], edi
'006a0ced    33f6                    xor esi, esi
'006a0cef    8975f8                  mov dword ptr [ebp-08], esi
'006a0cf2    8b0f                    mov ecx, dword ptr [edi]
'006a0cf4    57                      push edi
'006a0cf5    ff5104                  call dword ptr [ecx+04]
'006a0cf8    8975e0                  mov dword ptr [ebp-20], esi
'006a0cfb    8975dc                  mov dword ptr [ebp-24], esi
'006a0cfe    8975d8                  mov dword ptr [ebp-28], esi
'006a0d01    8975d4                  mov dword ptr [ebp-2c], esi
'006a0d04    8975d0                  mov dword ptr [ebp-30], esi
'006a0d07    8975cc                  mov dword ptr [ebp-34], esi
'006a0d0a    8975c8                  mov dword ptr [ebp-38], esi
'006a0d0d    8975c4                  mov dword ptr [ebp-3c], esi
'006a0d10    8975c0                  mov dword ptr [ebp-40], esi
'006a0d13    8975b0                  mov dword ptr [ebp-50], esi
'006a0d16    8975a0                  mov dword ptr [ebp-60], esi
'006a0d19    897590                  mov dword ptr [ebp-70], esi
'006a0d1c    897580                  mov dword ptr [ebp-80], esi
'006a0d1f    89b570ffffff            mov dword ptr [ebp+ffffff70], esi
'006a0d25    89b560ffffff            mov dword ptr [ebp+ffffff60], esi
'006a0d2b    89b550ffffff            mov dword ptr [ebp+ffffff50], esi
'006a0d31    89b540ffffff            mov dword ptr [ebp+ffffff40], esi
'006a0d37    89b530ffffff            mov dword ptr [ebp+ffffff30], esi
'006a0d3d    89b520ffffff            mov dword ptr [ebp+ffffff20], esi
'006a0d43    89b510ffffff            mov dword ptr [ebp+ffffff10], esi
'006a0d49    89b500ffffff            mov dword ptr [ebp+ffffff00], esi
'006a0d4f    89b5f0feffff            mov dword ptr [ebp+fffffef0], esi
'006a0d55    89b5e0feffff            mov dword ptr [ebp+fffffee0], esi
'006a0d5b    89b5dcfeffff            mov dword ptr [ebp+fffffedc], esi
'006a0d61    66393528b07200          cmp word ptr [0072b028], si
'006a0d68    7508                    jne 6a0d72
'006a0d6a    6a01                    push 01

' *** Reference to "__vbaOnError"
'006a0d6c    ff15b0104000            call dword ptr [004010b0]
On Error Goto handler_0
'006a0d72    8b07                    mov eax, dword ptr [edi]
'006a0d74    57                      push edi
'006a0d75    ff9008070000            call dword ptr [eax+00000708]
'006a0d7b    3bc6                    cmp eax, esi
'006a0d7d    7d12                    jge 6a0d91

If (FrmLstArticle < 0) Then
'006a0d7f    6808070000              push 00000708
'006a0d84    6870f44400              push 0044f470
'006a0d89    57                      push edi
'006a0d8a    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006a0d8b    ff1580104000            call dword ptr [00401080]
    
End If

' *** Reference to "__vbaExitProc"
'006a0d91    ff15a0104000            call dword ptr [004010a0]
'006a0d97    684a116a00              push 006a114a
'006a0d9c    e9a8030000              jmp 6a1149

' *** Reference to "rtcErrObj"
'006a0da1    8b1d6c124000            mov ebx, dword ptr [0040126c]
'006a0da7    ffd3                    call ebx
'006a0da9    50                      push eax
'006a0daa    8d55c4                  lea edx, dword ptr [ebp-3c]
'006a0dad    52                      push edx

' *** Reference to "__vbaObjSet"
'006a0dae    8b3db4104000            mov edi, dword ptr [004010b4]
'006a0db4    ffd7                    call edi
Set var_9 = Err
'006a0db6    8bf0                    mov esi, eax
'006a0db8    8b06                    mov eax, dword ptr [esi]
'006a0dba    8d4de0                  lea ecx, dword ptr [ebp-20]
'006a0dbd    51                      push ecx
'006a0dbe    56                      push esi
'006a0dbf    ff502c                  call dword ptr [eax+2c]
var_3 = var_9.Description
'006a0dc2    dbe2                    fnclex
'006a0dc4    85c0                    test ax, ax
'006a0dc6    7d0f                    jge 6a0dd7
'006a0dc8    6a2c                    push 2c
'006a0dca    685c084300              push 0043085c
'006a0dcf    56                      push esi
'006a0dd0    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006a0dd1    ff1580104000            call dword ptr [00401080]
'006a0dd7    ffd3                    call ebx
'006a0dd9    50                      push eax
'006a0dda    8d55c0                  lea edx, dword ptr [ebp-40]
'006a0ddd    52                      push edx
'006a0dde    ffd7                    call edi
Set var_5 = Err
'006a0de0    8bf0                    mov esi, eax
'006a0de2    8b06                    mov eax, dword ptr [esi]
'006a0de4    8d8ddcfeffff            lea ecx, dword ptr [ebp+fffffedc]
'006a0dea    51                      push ecx
'006a0deb    56                      push esi
'006a0dec    ff501c                  call dword ptr [eax+1c]
var_10 = var_5.Number
'006a0def    dbe2                    fnclex
'006a0df1    85c0                    test ax, ax
'006a0df3    7d0f                    jge 6a0e04
'006a0df5    6a1c                    push 1c
'006a0df7    685c084300              push 0043085c
'006a0dfc    56                      push esi
'006a0dfd    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006a0dfe    ff1580104000            call dword ptr [00401080]
'006a0e04    b904000280              mov ecx, 80020004
'006a0e09    894d98                  mov dword ptr [ebp-68], ecx
'006a0e0c    b80a000000              mov eax, 0000000a
'006a0e11    894590                  mov dword ptr [ebp-70], eax
'006a0e14    894da8                  mov dword ptr [ebp-58], ecx
'006a0e17    8945a0                  mov dword ptr [ebp-60], eax
'006a0e1a    c78528ffffff24b07200    mov dword ptr [ebp+ffffff28], 0072b024
'006a0e24    c78520ffffff08400000    mov dword ptr [ebp+ffffff20], 00004008
'006a0e2e    6814084300              push 00430814
'006a0e33    8b55e0                  mov edx, dword ptr [ebp-20]
'006a0e36    52                      push edx

' *** Reference to "__vbaStrCat"
'006a0e37    8b3570104000            mov esi, dword ptr [00401070]
'006a0e3d    ffd6                    call esi
var_11 = ("L'erreur suivante s'est produite : ") & (var_3)
'006a0e3f    8bd0                    mov edx, eax
'006a0e41    8d4ddc                  lea ecx, dword ptr [ebp-24]

' *** Reference to "__vbaStrMove"
'006a0e44    8b3dd0124000            mov edi, dword ptr [004012d0]
'006a0e4a    ffd7                    call edi
'006a0e4c    50                      push eax
'006a0e4d    6870084300              push 00430870
'006a0e52    ffd6                    call esi
var_76 = (var_11) & (vbCrLf)
'006a0e54    8bd0                    mov edx, eax
'006a0e56    8d4dd8                  lea ecx, dword ptr [ebp-28]
'006a0e59    ffd7                    call edi
'006a0e5b    50                      push eax
'006a0e5c    6870084300              push 00430870
'006a0e61    ffd6                    call esi
var_12 = (var_76) & (vbCrLf)
'006a0e63    8bd0                    mov edx, eax
'006a0e65    8d4dd4                  lea ecx, dword ptr [ebp-2c]
'006a0e68    ffd7                    call edi
'006a0e6a    50                      push eax
'006a0e6b    6880084300              push 00430880
'006a0e70    ffd6                    call esi
var_13 = (var_12) & (" numéro d'erreur (")
'006a0e72    8bd0                    mov edx, eax
'006a0e74    8d4dd0                  lea ecx, dword ptr [ebp-30]
'006a0e77    ffd7                    call edi
'006a0e79    50                      push eax
'006a0e7a    8b85dcfeffff            mov eax, dword ptr [ebp+fffffedc]
'006a0e80    50                      push eax

' *** Reference to "__vbaStrI4"
'006a0e81    ff1520104000            call dword ptr [00401020]
'006a0e87    8bd0                    mov edx, eax
'006a0e89    8d4dcc                  lea ecx, dword ptr [ebp-34]
'006a0e8c    ffd7                    call edi
'006a0e8e    50                      push eax
'006a0e8f    ffd6                    call esi
var_127 = (var_13) & (CStr(var_10))
'006a0e91    8bd0                    mov edx, eax
'006a0e93    8d4dc8                  lea ecx, dword ptr [ebp-38]
'006a0e96    ffd7                    call edi
'006a0e98    50                      push eax
'006a0e99    68ac084300              push 004308ac
'006a0e9e    ffd6                    call esi
var_15 = (var_127) & (" )")
'006a0ea0    8945b8                  mov dword ptr [ebp-48], eax
'006a0ea3    bf08000000              mov edi, 00000008
'006a0ea8    897db0                  mov dword ptr [ebp-50], edi
'006a0eab    8d4d90                  lea ecx, dword ptr [ebp-70]
'006a0eae    51                      push ecx
'006a0eaf    8d55a0                  lea edx, dword ptr [ebp-60]
'006a0eb2    52                      push edx
'006a0eb3    8d8520ffffff            lea eax, dword ptr [ebp+ffffff20]
'006a0eb9    50                      push eax
'006a0eba    6a10                    push 10
'006a0ebc    8d4db0                  lea ecx, dword ptr [ebp-50]
'006a0ebf    51                      push ecx

' *** Reference to "rtcMsgBox"
'006a0ec0    ff15b8104000            call dword ptr [004010b8]
var_128 = MsgBox(var_15, 16, vbNullString)
'006a0ec6    8d55c8                  lea edx, dword ptr [ebp-38]
'006a0ec9    52                      push edx
'006a0eca    8d45cc                  lea eax, dword ptr [ebp-34]
'006a0ecd    50                      push eax
'006a0ece    8d4dd0                  lea ecx, dword ptr [ebp-30]
'006a0ed1    51                      push ecx
'006a0ed2    8d55d4                  lea edx, dword ptr [ebp-2c]
'006a0ed5    52                      push edx
'006a0ed6    8d45d8                  lea eax, dword ptr [ebp-28]
'006a0ed9    50                      push eax
'006a0eda    8d4ddc                  lea ecx, dword ptr [ebp-24]
'006a0edd    51                      push ecx
'006a0ede    8d55e0                  lea edx, dword ptr [ebp-20]
'006a0ee1    52                      push edx
'006a0ee2    6a07                    push 07

' *** Reference to "__vbaFreeStrList"
'006a0ee4    ff155c124000            call dword ptr [0040125c]
'#Cleanup( 0, -4508, -4512, -4516, -4520, -4524, -4528)
'006a0eea    8d45c0                  lea eax, dword ptr [ebp-40]
'006a0eed    50                      push eax
'006a0eee    8d4dc4                  lea ecx, dword ptr [ebp-3c]
'006a0ef1    51                      push ecx
'006a0ef2    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'006a0ef4    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_9, var_5)
'006a0efa    8d5590                  lea edx, dword ptr [ebp-70]
'006a0efd    52                      push edx
'006a0efe    8d45a0                  lea eax, dword ptr [ebp-60]
'006a0f01    50                      push eax
'006a0f02    8d4db0                  lea ecx, dword ptr [ebp-50]
'006a0f05    51                      push ecx
'006a0f06    6a03                    push 03

' *** Reference to "__vbaFreeVarList"
'006a0f08    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_6, var_7, var_8)
'006a0f0e    83c43c                  add esp, 3c
'006a0f11    8d55b0                  lea edx, dword ptr [ebp-50]
'006a0f14    52                      push edx

' *** Reference to "rtcGetPresentDate"
'006a0f15    ff15f4124000            call dword ptr [004012f4]
var_15 = Now()
'006a0f1b    c78528ffffffb8084300    mov dword ptr [ebp+ffffff28], 004308b8
'006a0f25    89bd20ffffff            mov dword ptr [ebp+ffffff20], edi
'006a0f2b    8d9520ffffff            lea edx, dword ptr [ebp+ffffff20]
'006a0f31    8d4da0                  lea ecx, dword ptr [ebp-60]

' *** Reference to "__vbaVarDup"
'006a0f34    ff1598124000            call dword ptr [00401298]
var_7 = ("dd/MM/yyyy hh:mm:ss")
'006a0f3a    6a01                    push 01
'006a0f3c    6a01                    push 01
'006a0f3e    8d45a0                  lea eax, dword ptr [ebp-60]
'006a0f41    50                      push eax
'006a0f42    8d4db0                  lea ecx, dword ptr [ebp-50]
'006a0f45    51                      push ecx
'006a0f46    8d5590                  lea edx, dword ptr [ebp-70]
'006a0f49    52                      push edx

' *** Reference to "rtcVarFromFormatVar"
'006a0f4a    ff1574104000            call dword ptr [00401074]
'006a0f50    ffd3                    call ebx
'006a0f52    50                      push eax
'006a0f53    8d45c4                  lea eax, dword ptr [ebp-3c]
'006a0f56    50                      push eax

' *** Reference to "__vbaObjSet"
'006a0f57    ff15b4104000            call dword ptr [004010b4]
Set var_9 = Err
'006a0f5d    8bf0                    mov esi, eax
'006a0f5f    8b0e                    mov ecx, dword ptr [esi]
'006a0f61    8d95dcfeffff            lea edx, dword ptr [ebp+fffffedc]
'006a0f67    52                      push edx
'006a0f68    56                      push esi
'006a0f69    ff511c                  call dword ptr [ecx+1c]
var_10 = var_9.Number
'006a0f6c    dbe2                    fnclex
'006a0f6e    85c0                    test ax, ax
'006a0f70    7d0f                    jge 6a0f81
'006a0f72    6a1c                    push 1c
'006a0f74    685c084300              push 0043085c
'006a0f79    56                      push esi
'006a0f7a    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006a0f7b    ff1580104000            call dword ptr [00401080]
'006a0f81    ffd3                    call ebx
'006a0f83    50                      push eax
'006a0f84    8d45c0                  lea eax, dword ptr [ebp-40]
'006a0f87    50                      push eax

' *** Reference to "__vbaObjSet"
'006a0f88    ff15b4104000            call dword ptr [004010b4]
Set var_5 = Err
'006a0f8e    8bf0                    mov esi, eax
'006a0f90    8b0e                    mov ecx, dword ptr [esi]
'006a0f92    8d55e0                  lea edx, dword ptr [ebp-20]
'006a0f95    52                      push edx
'006a0f96    56                      push esi
'006a0f97    ff512c                  call dword ptr [ecx+2c]
var_3 = var_5.Description
'006a0f9a    dbe2                    fnclex
'006a0f9c    85c0                    test ax, ax
'006a0f9e    7d0f                    jge 6a0faf
'006a0fa0    6a2c                    push 2c
'006a0fa2    685c084300              push 0043085c
'006a0fa7    56                      push esi
'006a0fa8    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006a0fa9    ff1580104000            call dword ptr [00401080]
'006a0faf    c78518ffffffe4084300    mov dword ptr [ebp+ffffff18], 004308e4
'006a0fb9    89bd10ffffff            mov dword ptr [ebp+ffffff10], edi
'006a0fbf    8b85dcfeffff            mov eax, dword ptr [ebp+fffffedc]
'006a0fc5    898508ffffff            mov dword ptr [ebp+ffffff08], eax
'006a0fcb    c78500ffffff03000000    mov dword ptr [ebp+ffffff00], 00000003
'006a0fd5    c785f8feffff08094300    mov dword ptr [ebp+fffffef8], 00430908
'006a0fdf    89bdf0feffff            mov dword ptr [ebp+fffffef0], edi
'006a0fe5    8b45e0                  mov eax, dword ptr [ebp-20]
'006a0fe8    c745e000000000          mov dword ptr [ebp-20], 00000000
'006a0fef    898558ffffff            mov dword ptr [ebp+ffffff58], eax
'006a0ff5    89bd50ffffff            mov dword ptr [ebp+ffffff50], edi
'006a0ffb    c785e8fefffff8264500    mov dword ptr [ebp+fffffee8], 004526f8
'006a1005    89bde0feffff            mov dword ptr [ebp+fffffee0], edi
'006a100b    8d4d90                  lea ecx, dword ptr [ebp-70]
'006a100e    51                      push ecx
'006a100f    8d9510ffffff            lea edx, dword ptr [ebp+ffffff10]
'006a1015    52                      push edx
'006a1016    8d4580                  lea eax, dword ptr [ebp-80]
'006a1019    50                      push eax

' *** Reference to "__vbaVarCat"
'006a101a    8b3508124000            mov esi, dword ptr [00401208]
'006a1020    ffd6                    call esi
'006a1022    50                      push eax
'006a1023    8d8d00ffffff            lea ecx, dword ptr [ebp+ffffff00]
'006a1029    51                      push ecx
'006a102a    8d9570ffffff            lea edx, dword ptr [ebp+ffffff70]
'006a1030    52                      push edx
'006a1031    ffd6                    call esi
'006a1033    50                      push eax
'006a1034    8d85f0feffff            lea eax, dword ptr [ebp+fffffef0]
'006a103a    50                      push eax
'006a103b    8d8d60ffffff            lea ecx, dword ptr [ebp+ffffff60]
'006a1041    51                      push ecx
'006a1042    ffd6                    call esi
'006a1044    50                      push eax
'006a1045    8d9550ffffff            lea edx, dword ptr [ebp+ffffff50]
'006a104b    52                      push edx
'006a104c    8d8540ffffff            lea eax, dword ptr [ebp+ffffff40]
'006a1052    50                      push eax
'006a1053    ffd6                    call esi
'006a1055    50                      push eax
'006a1056    8d8de0feffff            lea ecx, dword ptr [ebp+fffffee0]
'006a105c    51                      push ecx
'006a105d    8d9530ffffff            lea edx, dword ptr [ebp+ffffff30]
'006a1063    52                      push edx
'006a1064    ffd6                    call esi
'006a1066    50                      push eax
'006a1067    33c0                    xor eax, eax
'006a1069    66a12ab07200            mov eax, dword ptr [0072b02a]
'006a106f    50                      push eax
'006a1070    6884094300              push 00430984

' *** Reference to "__vbaPrintFile"
'006a1075    ff15b8114000            call dword ptr [004011b8]
Print #0, 
'006a107b    8d4dc0                  lea ecx, dword ptr [ebp-40]
'006a107e    51                      push ecx
'006a107f    8d55c4                  lea edx, dword ptr [ebp-3c]
'006a1082    52                      push edx
'006a1083    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'006a1085    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_9, var_5)
'006a108b    8d8530ffffff            lea eax, dword ptr [ebp+ffffff30]
'006a1091    50                      push eax
'006a1092    8d8d40ffffff            lea ecx, dword ptr [ebp+ffffff40]
'006a1098    51                      push ecx
'006a1099    8d9550ffffff            lea edx, dword ptr [ebp+ffffff50]
'006a109f    52                      push edx
'006a10a0    8d8560ffffff            lea eax, dword ptr [ebp+ffffff60]
'006a10a6    50                      push eax
'006a10a7    8d8d70ffffff            lea ecx, dword ptr [ebp+ffffff70]
'006a10ad    51                      push ecx
'006a10ae    8d5580                  lea edx, dword ptr [ebp-80]
'006a10b1    52                      push edx
'006a10b2    8d4590                  lea eax, dword ptr [ebp-70]
'006a10b5    50                      push eax
'006a10b6    8d4da0                  lea ecx, dword ptr [ebp-60]
'006a10b9    51                      push ecx
'006a10ba    8d55b0                  lea edx, dword ptr [ebp-50]
'006a10bd    52                      push edx
'006a10be    6a09                    push 09

' *** Reference to "__vbaFreeVarList"
'006a10c0    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_6, var_7, var_8, var_18, var_19, var_20, var_21, var_22, var_23)
'006a10c6    83c440                  add esp, 40
'006a10c9    6a00                    push 00

' *** Reference to "__vbaResume"
'006a10cb    ff1568104000            call dword ptr [00401068]
'006a10d1    e9bbfcffff              jmp 6a0d91
Resume handler_6A0D91
'006a10d6    8d4dc8                  lea ecx, dword ptr [ebp-38]
'006a10d9    51                      push ecx
'006a10da    8d55cc                  lea edx, dword ptr [ebp-34]
'006a10dd    52                      push edx
'006a10de    8d45d0                  lea eax, dword ptr [ebp-30]
'006a10e1    50                      push eax
'006a10e2    8d4dd4                  lea ecx, dword ptr [ebp-2c]
'006a10e5    51                      push ecx
'006a10e6    8d55d8                  lea edx, dword ptr [ebp-28]
'006a10e9    52                      push edx
'006a10ea    8d45dc                  lea eax, dword ptr [ebp-24]
'006a10ed    50                      push eax
'006a10ee    8d4de0                  lea ecx, dword ptr [ebp-20]
'006a10f1    51                      push ecx
'006a10f2    6a07                    push 07

' *** Reference to "__vbaFreeStrList"
'006a10f4    ff155c124000            call dword ptr [0040125c]
'#Cleanup( var_3, -4508, -4512, -4516, -4520, -4524, -4528)
'006a10fa    8d55c0                  lea edx, dword ptr [ebp-40]
'006a10fd    52                      push edx
'006a10fe    8d45c4                  lea eax, dword ptr [ebp-3c]
'006a1101    50                      push eax
'006a1102    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'006a1104    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_9, var_5)
'006a110a    8d8d30ffffff            lea ecx, dword ptr [ebp+ffffff30]
'006a1110    51                      push ecx
'006a1111    8d9540ffffff            lea edx, dword ptr [ebp+ffffff40]
'006a1117    52                      push edx
'006a1118    8d8550ffffff            lea eax, dword ptr [ebp+ffffff50]
'006a111e    50                      push eax
'006a111f    8d8d60ffffff            lea ecx, dword ptr [ebp+ffffff60]
'006a1125    51                      push ecx
'006a1126    8d9570ffffff            lea edx, dword ptr [ebp+ffffff70]
'006a112c    52                      push edx
'006a112d    8d4580                  lea eax, dword ptr [ebp-80]
'006a1130    50                      push eax
'006a1131    8d4d90                  lea ecx, dword ptr [ebp-70]
'006a1134    51                      push ecx
'006a1135    8d55a0                  lea edx, dword ptr [ebp-60]
'006a1138    52                      push edx
'006a1139    8d45b0                  lea eax, dword ptr [ebp-50]
'006a113c    50                      push eax
'006a113d    6a09                    push 09

' *** Reference to "__vbaFreeVarList"
'006a113f    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_6, var_7, var_8, var_18, var_19, var_20, var_21, var_22, var_23)
'006a1145    83c454                  add esp, 54
'006a1148    c3                      ret
'006a1149    c3                      ret
'006a114a    8b4508                  mov eax, dword ptr [ebp+08]
'006a114d    8b08                    mov ecx, dword ptr [eax]
'006a114f    50                      push eax
'006a1150    ff5108                  call dword ptr [ecx+08]
'006a1153    8b45f4                  mov eax, dword ptr [ebp-0c]
'006a1156    8b4de4                  mov ecx, dword ptr [ebp-1c]
    'Reference to '__except_list'
'006a1159    64890d00000000          mov dword ptr fs:[00000000], ecx
'006a1160    5f                      pop edi
'006a1161    5e                      pop esi
'006a1162    5b                      pop ebx
'006a1163    8be5                    mov esp, ebp
'006a1165    5d                      pop ebp
'006a1166    c20400                  ret 0004


End Sub


Private Sub FGarticle_Event17()
'006a1170    55                      push ebp
'006a1171    8bec                    mov ebp, esp
'006a1173    83ec0c                  sub esp, 0c

' *** Reference to "__vbaExceptHandler"
'006a1176    6866724000              push 00407266
'006a117b    64a100000000            mov ax, word ptr fs:[00000000]
'006a1181    50                      push eax
    'Reference to '__except_list'
'006a1182    64892500000000          mov dword ptr fs:[00000000], esp
'006a1189    81eca8000000            sub esp, 000000a8
'006a118f    53                      push ebx
'006a1190    56                      push esi
'006a1191    57                      push edi
'006a1192    8965f4                  mov dword ptr [ebp-0c], esp
'006a1195    c745f840624000          mov dword ptr [ebp-08], 00406240
'006a119c    8b7508                  mov esi, dword ptr [ebp+08]
'006a119f    8bc6                    mov eax, esi
'006a11a1    83e001                  and eax, 01
'006a11a4    8945fc                  mov dword ptr [ebp-04], eax
'006a11a7    83e6fe                  and esi, fffffffe
'006a11aa    8b0e                    mov ecx, dword ptr [esi]
'006a11ac    56                      push esi
'006a11ad    897508                  mov dword ptr [ebp+08], esi
'006a11b0    ff5104                  call dword ptr [ecx+04]
'006a11b3    8b16                    mov edx, dword ptr [esi]
'006a11b5    33db                    xor ebx, ebx
'006a11b7    56                      push esi
'006a11b8    895de8                  mov dword ptr [ebp-18], ebx
'006a11bb    895de4                  mov dword ptr [ebp-1c], ebx
'006a11be    895de0                  mov dword ptr [ebp-20], ebx
'006a11c1    895ddc                  mov dword ptr [ebp-24], ebx
'006a11c4    895dcc                  mov dword ptr [ebp-34], ebx
'006a11c7    895dbc                  mov dword ptr [ebp-44], ebx
'006a11ca    ff9200030000            call dword ptr [edx+00000300]

' *** Reference to "__vbaObjSet"
'006a11d0    8b3db4104000            mov edi, dword ptr [004010b4]
'006a11d6    50                      push eax
'006a11d7    8d45dc                  lea eax, dword ptr [ebp-24]
'006a11da    50                      push eax
'006a11db    ffd7                    call edi
Set var_39 = Me
'006a11dd    8b0e                    mov ecx, dword ptr [esi]
'006a11df    53                      push ebx
'006a11e0    6a11                    push 11
'006a11e2    56                      push esi
'006a11e3    898558ffffff            mov dword ptr [ebp+ffffff58], eax
'006a11e9    ff9114030000            call dword ptr [ecx+00000314]
'006a11ef    50                      push eax
'006a11f0    8d55e4                  lea edx, dword ptr [ebp-1c]
'006a11f3    52                      push edx
'006a11f4    ffd7                    call edi
Set var_40 = var_39
'006a11f6    50                      push eax
'006a11f7    8d45cc                  lea eax, dword ptr [ebp-34]
'006a11fa    50                      push eax

' *** Reference to "__vbaLateIdCallLd"
'006a11fb    ff157c114000            call dword ptr [0040117c]
var_43 = var_40.UNK_FrmLstArticle_17
'006a1201    83c410                  add esp, 10
'006a1204    50                      push eax

' *** Reference to "__vbaI4Var"
'006a1205    ff157c124000            call dword ptr [0040127c]
'006a120b    8b8d58ffffff            mov ecx, dword ptr [ebp+ffffff58]
'006a1211    8b11                    mov edx, dword ptr [ecx]
'006a1213    894594                  mov dword ptr [ebp-6c], eax
'006a1216    83ec10                  sub esp, 10
'006a1219    8bcc                    mov ecx, esp
'006a121b    b803000000              mov eax, 00000003
'006a1220    8901                    mov dword ptr [ecx], eax
'006a1222    8b45b0                  mov eax, dword ptr [ebp-50]
'006a1225    894104                  mov dword ptr [ecx+04], eax
'006a1228    8b45b8                  mov eax, dword ptr [ebp-48]
'006a122b    895908                  mov dword ptr [ecx+08], ebx
'006a122e    89410c                  mov dword ptr [ecx+0c], eax
'006a1231    83ec10                  sub esp, 10
'006a1234    8bcc                    mov ecx, esp
'006a1236    b803000000              mov eax, 00000003
'006a123b    8901                    mov dword ptr [ecx], eax
'006a123d    8b4590                  mov eax, dword ptr [ebp-70]
'006a1240    894104                  mov dword ptr [ecx+04], eax
'006a1243    8b4594                  mov eax, dword ptr [ebp-6c]
'006a1246    894108                  mov dword ptr [ecx+08], eax
'006a1249    8b4598                  mov eax, dword ptr [ebp-68]
'006a124c    89410c                  mov dword ptr [ecx+0c], eax
'006a124f    83ec10                  sub esp, 10
'006a1252    8bcc                    mov ecx, esp
'006a1254    c7856cffffff02000000    mov dword ptr [ebp+ffffff6c], 00000002
'006a125e    8b856cffffff            mov eax, dword ptr [ebp+ffffff6c]
'006a1264    8901                    mov dword ptr [ecx], eax
'006a1266    8b8570ffffff            mov eax, dword ptr [ebp+ffffff70]
'006a126c    894104                  mov dword ptr [ecx+04], eax
'006a126f    b802000000              mov eax, 00000002
'006a1274    894108                  mov dword ptr [ecx+08], eax
'006a1277    8b8578ffffff            mov eax, dword ptr [ebp+ffffff78]
'006a127d    6a03                    push 03
'006a127f    89410c                  mov dword ptr [ecx+0c], eax
'006a1282    689e000000              push 0000009e
'006a1287    8b0e                    mov ecx, dword ptr [esi]
'006a1289    56                      push esi
'006a128a    899548ffffff            mov dword ptr [ebp+ffffff48], edx
'006a1290    ff9114030000            call dword ptr [ecx+00000314]
'006a1296    50                      push eax
'006a1297    8d55e0                  lea edx, dword ptr [ebp-20]
'006a129a    52                      push edx
'006a129b    ffd7                    call edi
Set var_3 = Nothing
'006a129d    50                      push eax
'006a129e    8d45bc                  lea eax, dword ptr [ebp-44]
'006a12a1    50                      push eax

' *** Reference to "__vbaLateIdCallLd"
'006a12a2    ff157c114000            call dword ptr [0040117c]
var_58 = var_3.UNK_-256 - 12_158
'006a12a8    83c440                  add esp, 40
'006a12ab    50                      push eax
'006a12ac    8d4de8                  lea ecx, dword ptr [ebp-18]
'006a12af    51                      push ecx

' *** Reference to "__vbaStrVarVal"
'006a12b0    ff15fc114000            call dword ptr [004011fc]
'006a12b6    8b9d58ffffff            mov ebx, dword ptr [ebp+ffffff58]
'006a12bc    8b9548ffffff            mov edx, dword ptr [ebp+ffffff48]
'006a12c2    50                      push eax
'006a12c3    53                      push ebx
'006a12c4    ff92a4000000            call dword ptr [edx+000000a4]
'006a12ca    dbe2                    fnclex
'006a12cc    85c0                    test ax, ax
'006a12ce    7d12                    jge 6a12e2
'006a12d0    68a4000000              push 000000a4
'006a12d5    68d00d4300              push 00430dd0
'006a12da    53                      push ebx
'006a12db    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006a12dc    ff1580104000            call dword ptr [00401080]
'006a12e2    8d4de8                  lea ecx, dword ptr [ebp-18]

' *** Reference to "__vbaFreeStr"
'006a12e5    ff150c134000            call dword ptr [0040130c]
'#Cleanup(var_41)
'006a12eb    8d45dc                  lea eax, dword ptr [ebp-24]
'006a12ee    50                      push eax
'006a12ef    8d4de0                  lea ecx, dword ptr [ebp-20]
'006a12f2    51                      push ecx
'006a12f3    8d55e4                  lea edx, dword ptr [ebp-1c]
'006a12f6    52                      push edx
'006a12f7    6a03                    push 03

' *** Reference to "__vbaFreeObjList"
'006a12f9    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_40, var_3, var_39)
'006a12ff    8d45bc                  lea eax, dword ptr [ebp-44]
'006a1302    50                      push eax
'006a1303    8d4dcc                  lea ecx, dword ptr [ebp-34]
'006a1306    51                      push ecx
'006a1307    6a02                    push 02

' *** Reference to "__vbaFreeVarList"
'006a1309    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_43, var_58)
'006a130f    8b16                    mov edx, dword ptr [esi]
'006a1311    83c41c                  add esp, 1c
'006a1314    56                      push esi
'006a1315    ff92fc020000            call dword ptr [edx+000002fc]
'006a131b    50                      push eax
'006a131c    8d45dc                  lea eax, dword ptr [ebp-24]
'006a131f    50                      push eax
'006a1320    ffd7                    call edi
Set var_39 = 
'006a1322    8b0e                    mov ecx, dword ptr [esi]
'006a1324    6a00                    push 00
'006a1326    6a11                    push 11
'006a1328    56                      push esi
'006a1329    8bd8                    mov ebx, eax
'006a132b    ff9114030000            call dword ptr [ecx+00000314]
'006a1331    50                      push eax
'006a1332    8d55e4                  lea edx, dword ptr [ebp-1c]
'006a1335    52                      push edx
'006a1336    ffd7                    call edi
Set var_40 = var_39
'006a1338    50                      push eax
'006a1339    8d45cc                  lea eax, dword ptr [ebp-34]
'006a133c    50                      push eax

' *** Reference to "__vbaLateIdCallLd"
'006a133d    ff157c114000            call dword ptr [0040117c]
var_43 = var_40.UNK_var_58_17
'006a1343    83c410                  add esp, 10
'006a1346    50                      push eax

' *** Reference to "__vbaI4Var"
'006a1347    ff157c124000            call dword ptr [0040127c]
'006a134d    894594                  mov dword ptr [ebp-6c], eax
'006a1350    83ec10                  sub esp, 10
'006a1353    8bcc                    mov ecx, esp
'006a1355    b803000000              mov eax, 00000003
'006a135a    8901                    mov dword ptr [ecx], eax
'006a135c    8b45b0                  mov eax, dword ptr [ebp-50]
'006a135f    894104                  mov dword ptr [ecx+04], eax
'006a1362    33c0                    xor eax, eax
var_num1 = Empty
'006a1364    894108                  mov dword ptr [ecx+08], eax
'006a1367    8b45b8                  mov eax, dword ptr [ebp-48]
'006a136a    89410c                  mov dword ptr [ecx+0c], eax
'006a136d    8b13                    mov edx, dword ptr [ebx]
'006a136f    83ec10                  sub esp, 10
'006a1372    8bcc                    mov ecx, esp
'006a1374    b803000000              mov eax, 00000003
'006a1379    8901                    mov dword ptr [ecx], eax
'006a137b    8b4590                  mov eax, dword ptr [ebp-70]
'006a137e    894104                  mov dword ptr [ecx+04], eax
'006a1381    8b4594                  mov eax, dword ptr [ebp-6c]
'006a1384    894108                  mov dword ptr [ecx+08], eax
'006a1387    8b4598                  mov eax, dword ptr [ebp-68]
'006a138a    89410c                  mov dword ptr [ecx+0c], eax
'006a138d    83ec10                  sub esp, 10
'006a1390    8bcc                    mov ecx, esp
'006a1392    b802000000              mov eax, 00000002
'006a1397    8901                    mov dword ptr [ecx], eax
'006a1399    8b8570ffffff            mov eax, dword ptr [ebp+ffffff70]
'006a139f    894104                  mov dword ptr [ecx+04], eax
'006a13a2    b804000000              mov eax, 00000004
'006a13a7    894108                  mov dword ptr [ecx+08], eax
'006a13aa    8b8578ffffff            mov eax, dword ptr [ebp+ffffff78]
'006a13b0    6a03                    push 03
'006a13b2    689e000000              push 0000009e
'006a13b7    89410c                  mov dword ptr [ecx+0c], eax
'006a13ba    8b0e                    mov ecx, dword ptr [esi]
'006a13bc    56                      push esi
'006a13bd    899544ffffff            mov dword ptr [ebp+ffffff44], edx
'006a13c3    ff9114030000            call dword ptr [ecx+00000314]
'006a13c9    50                      push eax
'006a13ca    8d55e0                  lea edx, dword ptr [ebp-20]
'006a13cd    52                      push edx
'006a13ce    ffd7                    call edi
Set var_3 = Nothing
'006a13d0    50                      push eax
'006a13d1    8d45bc                  lea eax, dword ptr [ebp-44]
'006a13d4    50                      push eax

' *** Reference to "__vbaLateIdCallLd"
'006a13d5    ff157c114000            call dword ptr [0040117c]
var_58 = var_3.UNK_-256 - 12_158
'006a13db    83c440                  add esp, 40
'006a13de    50                      push eax
'006a13df    8d4de8                  lea ecx, dword ptr [ebp-18]
'006a13e2    51                      push ecx

' *** Reference to "__vbaStrVarVal"
'006a13e3    ff15fc114000            call dword ptr [004011fc]
'006a13e9    8b9544ffffff            mov edx, dword ptr [ebp+ffffff44]
'006a13ef    50                      push eax
'006a13f0    53                      push ebx
'006a13f1    ff92a4000000            call dword ptr [edx+000000a4]
'006a13f7    dbe2                    fnclex
'006a13f9    85c0                    test ax, ax
'006a13fb    7d12                    jge 6a140f
'006a13fd    68a4000000              push 000000a4
'006a1402    68d00d4300              push 00430dd0
'006a1407    53                      push ebx
'006a1408    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006a1409    ff1580104000            call dword ptr [00401080]
'006a140f    8d4de8                  lea ecx, dword ptr [ebp-18]

' *** Reference to "__vbaFreeStr"
'006a1412    ff150c134000            call dword ptr [0040130c]
'#Cleanup(var_41)
'006a1418    8d45dc                  lea eax, dword ptr [ebp-24]
'006a141b    50                      push eax
'006a141c    8d4de0                  lea ecx, dword ptr [ebp-20]
'006a141f    51                      push ecx
'006a1420    8d55e4                  lea edx, dword ptr [ebp-1c]
'006a1423    52                      push edx
'006a1424    6a03                    push 03

' *** Reference to "__vbaFreeObjList"
'006a1426    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_40, var_3, var_39)
'006a142c    8d45bc                  lea eax, dword ptr [ebp-44]
'006a142f    50                      push eax
'006a1430    8d4dcc                  lea ecx, dword ptr [ebp-34]
'006a1433    51                      push ecx
'006a1434    6a02                    push 02

' *** Reference to "__vbaFreeVarList"
'006a1436    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_43, var_58)
'006a143c    83c41c                  add esp, 1c
'006a143f    c745fc00000000          mov dword ptr [ebp-04], 00000000
'006a1446    687f146a00              push 006a147f
'006a144b    eb31                    jmp 6a147e
'006a144d    8d4de8                  lea ecx, dword ptr [ebp-18]

' *** Reference to "__vbaFreeStr"
'006a1450    ff150c134000            call dword ptr [0040130c]
'#Cleanup(var_41)
'006a1456    8d55dc                  lea edx, dword ptr [ebp-24]
'006a1459    52                      push edx
'006a145a    8d45e0                  lea eax, dword ptr [ebp-20]
'006a145d    50                      push eax
'006a145e    8d4de4                  lea ecx, dword ptr [ebp-1c]
'006a1461    51                      push ecx
'006a1462    6a03                    push 03

' *** Reference to "__vbaFreeObjList"
'006a1464    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_40, var_3, var_39)
'006a146a    8d55bc                  lea edx, dword ptr [ebp-44]
'006a146d    52                      push edx
'006a146e    8d45cc                  lea eax, dword ptr [ebp-34]
'006a1471    50                      push eax
'006a1472    6a02                    push 02

' *** Reference to "__vbaFreeVarList"
'006a1474    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_43, var_58)
'006a147a    83c41c                  add esp, 1c
'006a147d    c3                      ret
'006a147e    c3                      ret
'006a147f    8b4508                  mov eax, dword ptr [ebp+08]
'006a1482    8b08                    mov ecx, dword ptr [eax]
'006a1484    50                      push eax
'006a1485    ff5108                  call dword ptr [ecx+08]
'006a1488    8b45fc                  mov eax, dword ptr [ebp-04]
'006a148b    8b4dec                  mov ecx, dword ptr [ebp-14]
'006a148e    5f                      pop edi
'006a148f    5e                      pop esi
    'Reference to '__except_list'
'006a1490    64890d00000000          mov dword ptr fs:[00000000], ecx
'006a1497    5b                      pop ebx
'006a1498    8be5                    mov esp, ebp
'006a149a    5d                      pop ebp
'006a149b    c20400                  ret 0004


End Sub


Private Sub FGarticle_Event69()
'006a06c0    55                      push ebp
'006a06c1    8bec                    mov ebp, esp
'006a06c3    83ec14                  sub esp, 14

' *** Reference to "__vbaExceptHandler"
'006a06c6    6866724000              push 00407266
'006a06cb    64a100000000            mov ax, word ptr fs:[00000000]
'006a06d1    50                      push eax
    'Reference to '__except_list'
'006a06d2    64892500000000          mov dword ptr fs:[00000000], esp
'006a06d9    81ec5c010000            sub esp, 0000015c
'006a06df    53                      push ebx
'006a06e0    56                      push esi
'006a06e1    57                      push edi
'006a06e2    8965ec                  mov dword ptr [ebp-14], esp
'006a06e5    c745f0f0614000          mov dword ptr [ebp-10], 004061f0
'006a06ec    8b7508                  mov esi, dword ptr [ebp+08]
'006a06ef    8bc6                    mov eax, esi
'006a06f1    83e001                  and eax, 01
'006a06f4    8945f4                  mov dword ptr [ebp-0c], eax
'006a06f7    83e6fe                  and esi, fffffffe
'006a06fa    897508                  mov dword ptr [ebp+08], esi
'006a06fd    33db                    xor ebx, ebx
'006a06ff    895df8                  mov dword ptr [ebp-08], ebx
'006a0702    8b0e                    mov ecx, dword ptr [esi]
'006a0704    56                      push esi
'006a0705    ff5104                  call dword ptr [ecx+04]
'006a0708    895de0                  mov dword ptr [ebp-20], ebx
'006a070b    895ddc                  mov dword ptr [ebp-24], ebx
'006a070e    895dd8                  mov dword ptr [ebp-28], ebx
'006a0711    895dd4                  mov dword ptr [ebp-2c], ebx
'006a0714    895dd0                  mov dword ptr [ebp-30], ebx
'006a0717    895dcc                  mov dword ptr [ebp-34], ebx
'006a071a    895dc8                  mov dword ptr [ebp-38], ebx
'006a071d    895dc4                  mov dword ptr [ebp-3c], ebx
'006a0720    895dc0                  mov dword ptr [ebp-40], ebx
'006a0723    895db0                  mov dword ptr [ebp-50], ebx
'006a0726    895da0                  mov dword ptr [ebp-60], ebx
'006a0729    895d90                  mov dword ptr [ebp-70], ebx
'006a072c    895d80                  mov dword ptr [ebp-80], ebx
'006a072f    899d70ffffff            mov dword ptr [ebp+ffffff70], ebx
'006a0735    899d60ffffff            mov dword ptr [ebp+ffffff60], ebx
'006a073b    899d50ffffff            mov dword ptr [ebp+ffffff50], ebx
'006a0741    899d40ffffff            mov dword ptr [ebp+ffffff40], ebx
'006a0747    899d30ffffff            mov dword ptr [ebp+ffffff30], ebx
'006a074d    899d20ffffff            mov dword ptr [ebp+ffffff20], ebx
'006a0753    899d10ffffff            mov dword ptr [ebp+ffffff10], ebx
'006a0759    899d00ffffff            mov dword ptr [ebp+ffffff00], ebx
'006a075f    899df0feffff            mov dword ptr [ebp+fffffef0], ebx
'006a0765    899de0feffff            mov dword ptr [ebp+fffffee0], ebx
'006a076b    899dc0feffff            mov dword ptr [ebp+fffffec0], ebx
'006a0771    899dacfeffff            mov dword ptr [ebp+fffffeac], ebx
'006a0777    66391d28b07200          cmp word ptr [0072b028], bx
'006a077e    7508                    jne 6a0788
'006a0780    6a01                    push 01

' *** Reference to "__vbaOnError"
'006a0782    ff15b0104000            call dword ptr [004010b0]
On Error Goto handler_0
'006a0788    899d28ffffff            mov dword ptr [ebp+ffffff28], ebx
'006a078e    c78520ffffff03000000    mov dword ptr [ebp+ffffff20], 00000003
'006a0798    53                      push ebx
'006a0799    6a08                    push 08
'006a079b    8b06                    mov eax, dword ptr [esi]
'006a079d    56                      push esi
'006a079e    ff9014030000            call dword ptr [eax+00000314]
'006a07a4    50                      push eax
'006a07a5    8d4dc4                  lea ecx, dword ptr [ebp-3c]
'006a07a8    51                      push ecx

' *** Reference to "__vbaObjSet"
'006a07a9    8b3db4104000            mov edi, dword ptr [004010b4]
'006a07af    ffd7                    call edi
Set var_9 = Nothing
'006a07b1    50                      push eax
'006a07b2    8d55b0                  lea edx, dword ptr [ebp-50]
'006a07b5    52                      push edx

' *** Reference to "__vbaLateIdCallLd"
'006a07b6    ff157c114000            call dword ptr [0040117c]
var_6 = var_9.UNK_-256 - 20_8
'006a07bc    83c410                  add esp, 10
'006a07bf    50                      push eax

' *** Reference to "__vbaI4Var"
'006a07c0    ff157c124000            call dword ptr [0040127c]
'006a07c6    83e801                  sub eax, 01
var_num1 = CLng(var_6) - 1
'006a07c9    0f80d0040000            jo 6a0c9f
'006a07cf    898508ffffff            mov dword ptr [ebp+ffffff08], eax
'006a07d5    b903000000              mov ecx, 00000003
'006a07da    898d00ffffff            mov dword ptr [ebp+ffffff00], ecx
'006a07e0    c785e8feffff04000280    mov dword ptr [ebp+fffffee8], 80020004
'006a07ea    c785e0feffff0a000000    mov dword ptr [ebp+fffffee0], 0000000a
'006a07f4    83ec10                  sub esp, 10
'006a07f7    8bd4                    mov edx, esp
'006a07f9    8b9d20ffffff            mov ebx, dword ptr [ebp+ffffff20]
'006a07ff    891a                    mov dword ptr [edx], ebx
'006a0801    8b9d24ffffff            mov ebx, dword ptr [ebp+ffffff24]
'006a0807    895a04                  mov dword ptr [edx+04], ebx
'006a080a    8b9d28ffffff            mov ebx, dword ptr [ebp+ffffff28]
'006a0810    895a08                  mov dword ptr [edx+08], ebx
'006a0813    8b9d2cffffff            mov ebx, dword ptr [ebp+ffffff2c]
'006a0819    895a0c                  mov dword ptr [edx+0c], ebx
'006a081c    83ec10                  sub esp, 10
'006a081f    8bd4                    mov edx, esp
'006a0821    890a                    mov dword ptr [edx], ecx
'006a0823    8b8d04ffffff            mov ecx, dword ptr [ebp+ffffff04]
'006a0829    894a04                  mov dword ptr [edx+04], ecx
'006a082c    894208                  mov dword ptr [edx+08], eax
'006a082f    8b850cffffff            mov eax, dword ptr [ebp+ffffff0c]
'006a0835    89420c                  mov dword ptr [edx+0c], eax
'006a0838    83ec10                  sub esp, 10
'006a083b    8bcc                    mov ecx, esp
'006a083d    8b95e0feffff            mov edx, dword ptr [ebp+fffffee0]
'006a0843    8911                    mov dword ptr [ecx], edx
'006a0845    8b85e4feffff            mov eax, dword ptr [ebp+fffffee4]
'006a084b    894104                  mov dword ptr [ecx+04], eax
'006a084e    8b95e8feffff            mov edx, dword ptr [ebp+fffffee8]
'006a0854    895108                  mov dword ptr [ecx+08], edx
'006a0857    8b85ecfeffff            mov eax, dword ptr [ebp+fffffeec]
'006a085d    89410c                  mov dword ptr [ecx+0c], eax
'006a0860    83ec10                  sub esp, 10
'006a0863    8bcc                    mov ecx, esp
'006a0865    b802000000              mov eax, 00000002
'006a086a    8901                    mov dword ptr [ecx], eax
'006a086c    8b95c4feffff            mov edx, dword ptr [ebp+fffffec4]
'006a0872    895104                  mov dword ptr [ecx+04], edx
'006a0875    b864000000              mov eax, 00000064
'006a087a    894108                  mov dword ptr [ecx+08], eax
'006a087d    8b85ccfeffff            mov eax, dword ptr [ebp+fffffecc]
'006a0883    89410c                  mov dword ptr [ecx+0c], eax
'006a0886    6a04                    push 04
'006a0888    6893000000              push 00000093
'006a088d    8b0e                    mov ecx, dword ptr [esi]
'006a088f    56                      push esi
'006a0890    ff9114030000            call dword ptr [ecx+00000314]
'006a0896    50                      push eax
'006a0897    8d55c0                  lea edx, dword ptr [ebp-40]
'006a089a    52                      push edx
'006a089b    ffd7                    call edi
Set var_5 = Nothing
'006a089d    50                      push eax

' *** Reference to "__vbaLateIdCall"
'006a089e    ff1538104000            call dword ptr [00401038]
Call -256 - 20.Member_0x93(0, var_num1, , 100)
'006a08a4    8d45c0                  lea eax, dword ptr [ebp-40]
'006a08a7    50                      push eax
'006a08a8    8d4dc4                  lea ecx, dword ptr [ebp-3c]
'006a08ab    51                      push ecx
'006a08ac    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'006a08ae    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_9, var_5)
'006a08b4    83c458                  add esp, 58
'006a08b7    8d4db0                  lea ecx, dword ptr [ebp-50]

' *** Reference to "__vbaFreeVar"
'006a08ba    ff1528104000            call dword ptr [00401028]
'#Cleanup(var_6)

' *** Reference to "__vbaExitProc"
'006a08c0    ff15a0104000            call dword ptr [004010a0]
'006a08c6    68800c6a00              push 006a0c80
'006a08cb    e9af030000              jmp 6a0c7f

' *** Reference to "rtcErrObj"
'006a08d0    8b3d6c124000            mov edi, dword ptr [0040126c]
'006a08d6    ffd7                    call edi
'006a08d8    50                      push eax
'006a08d9    8d55c4                  lea edx, dword ptr [ebp-3c]
'006a08dc    52                      push edx

' *** Reference to "__vbaObjSet"
'006a08dd    ff15b4104000            call dword ptr [004010b4]
Set var_9 = Err
'006a08e3    8bf0                    mov esi, eax
'006a08e5    8b06                    mov eax, dword ptr [esi]
'006a08e7    8d4de0                  lea ecx, dword ptr [ebp-20]
'006a08ea    51                      push ecx
'006a08eb    56                      push esi
'006a08ec    ff502c                  call dword ptr [eax+2c]
var_3 = var_9.Description
'006a08ef    dbe2                    fnclex
'006a08f1    33db                    xor ebx, ebx
var_num2 = Empty
'006a08f3    3bc3                    cmp eax, ebx
'006a08f5    7d0f                    jge 6a0906
'006a08f7    6a2c                    push 2c
'006a08f9    685c084300              push 0043085c
'006a08fe    56                      push esi
'006a08ff    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006a0900    ff1580104000            call dword ptr [00401080]
'006a0906    ffd7                    call edi
'006a0908    50                      push eax
'006a0909    8d55c0                  lea edx, dword ptr [ebp-40]
'006a090c    52                      push edx

' *** Reference to "__vbaObjSet"
'006a090d    ff15b4104000            call dword ptr [004010b4]
Set var_5 = Err
'006a0913    8bf0                    mov esi, eax
'006a0915    8b06                    mov eax, dword ptr [esi]
'006a0917    8d8dacfeffff            lea ecx, dword ptr [ebp+fffffeac]
'006a091d    51                      push ecx
'006a091e    56                      push esi
'006a091f    ff501c                  call dword ptr [eax+1c]
var_520 = var_5.Number
'006a0922    dbe2                    fnclex
'006a0924    3bc3                    cmp eax, ebx
'006a0926    7d0f                    jge 6a0937
'006a0928    6a1c                    push 1c
'006a092a    685c084300              push 0043085c
'006a092f    56                      push esi
'006a0930    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006a0931    ff1580104000            call dword ptr [00401080]
'006a0937    b904000280              mov ecx, 80020004
'006a093c    894d98                  mov dword ptr [ebp-68], ecx
'006a093f    b80a000000              mov eax, 0000000a
'006a0944    894590                  mov dword ptr [ebp-70], eax
'006a0947    894da8                  mov dword ptr [ebp-58], ecx
'006a094a    8945a0                  mov dword ptr [ebp-60], eax
'006a094d    c78528ffffff24b07200    mov dword ptr [ebp+ffffff28], 0072b024
'006a0957    c78520ffffff08400000    mov dword ptr [ebp+ffffff20], 00004008
'006a0961    6814084300              push 00430814
'006a0966    8b55e0                  mov edx, dword ptr [ebp-20]
'006a0969    52                      push edx

' *** Reference to "__vbaStrCat"
'006a096a    8b3570104000            mov esi, dword ptr [00401070]
'006a0970    ffd6                    call esi
var_11 = ("L'erreur suivante s'est produite : ") & (var_3)
'006a0972    8bd0                    mov edx, eax
'006a0974    8d4ddc                  lea ecx, dword ptr [ebp-24]

' *** Reference to "__vbaStrMove"
'006a0977    8b3dd0124000            mov edi, dword ptr [004012d0]
'006a097d    ffd7                    call edi
'006a097f    50                      push eax
'006a0980    6870084300              push 00430870
'006a0985    ffd6                    call esi
var_868 = (var_11) & (vbCrLf)
'006a0987    8bd0                    mov edx, eax
'006a0989    8d4dd8                  lea ecx, dword ptr [ebp-28]
'006a098c    ffd7                    call edi
'006a098e    50                      push eax
'006a098f    6870084300              push 00430870
'006a0994    ffd6                    call esi
var_2461 = (var_868) & (vbCrLf)
'006a0996    8bd0                    mov edx, eax
'006a0998    8d4dd4                  lea ecx, dword ptr [ebp-2c]
'006a099b    ffd7                    call edi
'006a099d    50                      push eax
'006a099e    6880084300              push 00430880
'006a09a3    ffd6                    call esi
var_870 = (var_2461) & (" numéro d'erreur (")
'006a09a5    8bd0                    mov edx, eax
'006a09a7    8d4dd0                  lea ecx, dword ptr [ebp-30]
'006a09aa    ffd7                    call edi
'006a09ac    50                      push eax
'006a09ad    8b85acfeffff            mov eax, dword ptr [ebp+fffffeac]
'006a09b3    50                      push eax

' *** Reference to "__vbaStrI4"
'006a09b4    ff1520104000            call dword ptr [00401020]
'006a09ba    8bd0                    mov edx, eax
'006a09bc    8d4dcc                  lea ecx, dword ptr [ebp-34]
'006a09bf    ffd7                    call edi
'006a09c1    50                      push eax
'006a09c2    ffd6                    call esi
var_2070 = (var_870) & (CStr(var_520))
'006a09c4    8bd0                    mov edx, eax
'006a09c6    8d4dc8                  lea ecx, dword ptr [ebp-38]
'006a09c9    ffd7                    call edi
'006a09cb    50                      push eax
'006a09cc    68ac084300              push 004308ac
'006a09d1    ffd6                    call esi
var_871 = (var_2070) & (" )")
'006a09d3    8945b8                  mov dword ptr [ebp-48], eax
'006a09d6    bf08000000              mov edi, 00000008
'006a09db    897db0                  mov dword ptr [ebp-50], edi
'006a09de    8d4d90                  lea ecx, dword ptr [ebp-70]
'006a09e1    51                      push ecx
'006a09e2    8d55a0                  lea edx, dword ptr [ebp-60]
'006a09e5    52                      push edx
'006a09e6    8d8520ffffff            lea eax, dword ptr [ebp+ffffff20]
'006a09ec    50                      push eax
'006a09ed    6a10                    push 10
'006a09ef    8d4db0                  lea ecx, dword ptr [ebp-50]
'006a09f2    51                      push ecx

' *** Reference to "rtcMsgBox"
'006a09f3    ff15b8104000            call dword ptr [004010b8]
var_2071 = MsgBox(var_871, 16, vbNullString)
'006a09f9    8d55c8                  lea edx, dword ptr [ebp-38]
'006a09fc    52                      push edx
'006a09fd    8d45cc                  lea eax, dword ptr [ebp-34]
'006a0a00    50                      push eax
'006a0a01    8d4dd0                  lea ecx, dword ptr [ebp-30]
'006a0a04    51                      push ecx
'006a0a05    8d55d4                  lea edx, dword ptr [ebp-2c]
'006a0a08    52                      push edx
'006a0a09    8d45d8                  lea eax, dword ptr [ebp-28]
'006a0a0c    50                      push eax
'006a0a0d    8d4ddc                  lea ecx, dword ptr [ebp-24]
'006a0a10    51                      push ecx
'006a0a11    8d55e0                  lea edx, dword ptr [ebp-20]
'006a0a14    52                      push edx
'006a0a15    6a07                    push 07

' *** Reference to "__vbaFreeStrList"
'006a0a17    ff155c124000            call dword ptr [0040125c]
'#Cleanup( 0, -4600, -4604, -4608, -4612, -4616, -4620)
'006a0a1d    8d45c0                  lea eax, dword ptr [ebp-40]
'006a0a20    50                      push eax
'006a0a21    8d4dc4                  lea ecx, dword ptr [ebp-3c]
'006a0a24    51                      push ecx
'006a0a25    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'006a0a27    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_9, var_5)
'006a0a2d    8d5590                  lea edx, dword ptr [ebp-70]
'006a0a30    52                      push edx
'006a0a31    8d45a0                  lea eax, dword ptr [ebp-60]
'006a0a34    50                      push eax
'006a0a35    8d4db0                  lea ecx, dword ptr [ebp-50]
'006a0a38    51                      push ecx
'006a0a39    6a03                    push 03

' *** Reference to "__vbaFreeVarList"
'006a0a3b    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_6, var_7, var_8)
'006a0a41    83c43c                  add esp, 3c
'006a0a44    8d55b0                  lea edx, dword ptr [ebp-50]
'006a0a47    52                      push edx

' *** Reference to "rtcGetPresentDate"
'006a0a48    ff15f4124000            call dword ptr [004012f4]
var_871 = Now()
'006a0a4e    c78528ffffffb8084300    mov dword ptr [ebp+ffffff28], 004308b8
'006a0a58    89bd20ffffff            mov dword ptr [ebp+ffffff20], edi
'006a0a5e    8d9520ffffff            lea edx, dword ptr [ebp+ffffff20]
'006a0a64    8d4da0                  lea ecx, dword ptr [ebp-60]

' *** Reference to "__vbaVarDup"
'006a0a67    ff1598124000            call dword ptr [00401298]
var_7 = ("dd/MM/yyyy hh:mm:ss")
'006a0a6d    6a01                    push 01
'006a0a6f    6a01                    push 01
'006a0a71    8d45a0                  lea eax, dword ptr [ebp-60]
'006a0a74    50                      push eax
'006a0a75    8d4db0                  lea ecx, dword ptr [ebp-50]
'006a0a78    51                      push ecx
'006a0a79    8d5590                  lea edx, dword ptr [ebp-70]
'006a0a7c    52                      push edx

' *** Reference to "rtcVarFromFormatVar"
'006a0a7d    ff1574104000            call dword ptr [00401074]

' *** Reference to "rtcErrObj"
'006a0a83    ff156c124000            call dword ptr [0040126c]
'006a0a89    50                      push eax
'006a0a8a    8d45c4                  lea eax, dword ptr [ebp-3c]
'006a0a8d    50                      push eax

' *** Reference to "__vbaObjSet"
'006a0a8e    ff15b4104000            call dword ptr [004010b4]
Set var_9 = Err
'006a0a94    8bf0                    mov esi, eax
'006a0a96    8b0e                    mov ecx, dword ptr [esi]
'006a0a98    8d95acfeffff            lea edx, dword ptr [ebp+fffffeac]
'006a0a9e    52                      push edx
'006a0a9f    56                      push esi
'006a0aa0    ff511c                  call dword ptr [ecx+1c]
var_520 = var_9.Number
'006a0aa3    dbe2                    fnclex
'006a0aa5    3bc3                    cmp eax, ebx
'006a0aa7    7d0f                    jge 6a0ab8
'006a0aa9    6a1c                    push 1c
'006a0aab    685c084300              push 0043085c
'006a0ab0    56                      push esi
'006a0ab1    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006a0ab2    ff1580104000            call dword ptr [00401080]

' *** Reference to "rtcErrObj"
'006a0ab8    ff156c124000            call dword ptr [0040126c]
'006a0abe    50                      push eax
'006a0abf    8d45c0                  lea eax, dword ptr [ebp-40]
'006a0ac2    50                      push eax

' *** Reference to "__vbaObjSet"
'006a0ac3    ff15b4104000            call dword ptr [004010b4]
Set var_5 = Err
'006a0ac9    8bf0                    mov esi, eax
'006a0acb    8b0e                    mov ecx, dword ptr [esi]
'006a0acd    8d55e0                  lea edx, dword ptr [ebp-20]
'006a0ad0    52                      push edx
'006a0ad1    56                      push esi
'006a0ad2    ff512c                  call dword ptr [ecx+2c]
var_3 = var_5.Description
'006a0ad5    dbe2                    fnclex
'006a0ad7    3bc3                    cmp eax, ebx
'006a0ad9    7d0f                    jge 6a0aea
'006a0adb    6a2c                    push 2c
'006a0add    685c084300              push 0043085c
'006a0ae2    56                      push esi
'006a0ae3    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006a0ae4    ff1580104000            call dword ptr [00401080]
'006a0aea    c78518ffffffe4084300    mov dword ptr [ebp+ffffff18], 004308e4
'006a0af4    89bd10ffffff            mov dword ptr [ebp+ffffff10], edi
'006a0afa    8b85acfeffff            mov eax, dword ptr [ebp+fffffeac]
'006a0b00    898508ffffff            mov dword ptr [ebp+ffffff08], eax
'006a0b06    c78500ffffff03000000    mov dword ptr [ebp+ffffff00], 00000003
'006a0b10    c785f8feffff08094300    mov dword ptr [ebp+fffffef8], 00430908
'006a0b1a    89bdf0feffff            mov dword ptr [ebp+fffffef0], edi
'006a0b20    8b45e0                  mov eax, dword ptr [ebp-20]
'006a0b23    895de0                  mov dword ptr [ebp-20], ebx
'006a0b26    898558ffffff            mov dword ptr [ebp+ffffff58], eax
'006a0b2c    89bd50ffffff            mov dword ptr [ebp+ffffff50], edi
'006a0b32    c785e8feffff6c264500    mov dword ptr [ebp+fffffee8], 0045266c
'006a0b3c    89bde0feffff            mov dword ptr [ebp+fffffee0], edi
'006a0b42    8d4d90                  lea ecx, dword ptr [ebp-70]
'006a0b45    51                      push ecx
'006a0b46    8d9510ffffff            lea edx, dword ptr [ebp+ffffff10]
'006a0b4c    52                      push edx
'006a0b4d    8d4580                  lea eax, dword ptr [ebp-80]
'006a0b50    50                      push eax

' *** Reference to "__vbaVarCat"
'006a0b51    8b3508124000            mov esi, dword ptr [00401208]
'006a0b57    ffd6                    call esi
'006a0b59    50                      push eax
'006a0b5a    8d8d00ffffff            lea ecx, dword ptr [ebp+ffffff00]
'006a0b60    51                      push ecx
'006a0b61    8d9570ffffff            lea edx, dword ptr [ebp+ffffff70]
'006a0b67    52                      push edx
'006a0b68    ffd6                    call esi
'006a0b6a    50                      push eax
'006a0b6b    8d85f0feffff            lea eax, dword ptr [ebp+fffffef0]
'006a0b71    50                      push eax
'006a0b72    8d8d60ffffff            lea ecx, dword ptr [ebp+ffffff60]
'006a0b78    51                      push ecx
'006a0b79    ffd6                    call esi
'006a0b7b    50                      push eax
'006a0b7c    8d9550ffffff            lea edx, dword ptr [ebp+ffffff50]
'006a0b82    52                      push edx
'006a0b83    8d8540ffffff            lea eax, dword ptr [ebp+ffffff40]
'006a0b89    50                      push eax
'006a0b8a    ffd6                    call esi
'006a0b8c    50                      push eax
'006a0b8d    8d8de0feffff            lea ecx, dword ptr [ebp+fffffee0]
'006a0b93    51                      push ecx
'006a0b94    8d9530ffffff            lea edx, dword ptr [ebp+ffffff30]
'006a0b9a    52                      push edx
'006a0b9b    ffd6                    call esi
'006a0b9d    50                      push eax
'006a0b9e    33c0                    xor eax, eax
'006a0ba0    66a12ab07200            mov eax, dword ptr [0072b02a]
'006a0ba6    50                      push eax
'006a0ba7    6884094300              push 00430984

' *** Reference to "__vbaPrintFile"
'006a0bac    ff15b8114000            call dword ptr [004011b8]
Print #0, 
'006a0bb2    8d4dc0                  lea ecx, dword ptr [ebp-40]
'006a0bb5    51                      push ecx
'006a0bb6    8d55c4                  lea edx, dword ptr [ebp-3c]
'006a0bb9    52                      push edx
'006a0bba    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'006a0bbc    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_9, var_5)
'006a0bc2    8d8530ffffff            lea eax, dword ptr [ebp+ffffff30]
'006a0bc8    50                      push eax
'006a0bc9    8d8d40ffffff            lea ecx, dword ptr [ebp+ffffff40]
'006a0bcf    51                      push ecx
'006a0bd0    8d9550ffffff            lea edx, dword ptr [ebp+ffffff50]
'006a0bd6    52                      push edx
'006a0bd7    8d8560ffffff            lea eax, dword ptr [ebp+ffffff60]
'006a0bdd    50                      push eax
'006a0bde    8d8d70ffffff            lea ecx, dword ptr [ebp+ffffff70]
'006a0be4    51                      push ecx
'006a0be5    8d5580                  lea edx, dword ptr [ebp-80]
'006a0be8    52                      push edx
'006a0be9    8d4590                  lea eax, dword ptr [ebp-70]
'006a0bec    50                      push eax
'006a0bed    8d4da0                  lea ecx, dword ptr [ebp-60]
'006a0bf0    51                      push ecx
'006a0bf1    8d55b0                  lea edx, dword ptr [ebp-50]
'006a0bf4    52                      push edx
'006a0bf5    6a09                    push 09

' *** Reference to "__vbaFreeVarList"
'006a0bf7    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_6, var_7, var_8, var_18, var_19, var_20, var_21, var_22, var_23)
'006a0bfd    83c440                  add esp, 40
'006a0c00    53                      push ebx

' *** Reference to "__vbaResume"
'006a0c01    ff1568104000            call dword ptr [00401068]
'006a0c07    e9b4fcffff              jmp 6a08c0
Resume handler_6A08C0
'006a0c0c    8d55c8                  lea edx, dword ptr [ebp-38]
'006a0c0f    52                      push edx
'006a0c10    8d45cc                  lea eax, dword ptr [ebp-34]
'006a0c13    50                      push eax
'006a0c14    8d4dd0                  lea ecx, dword ptr [ebp-30]
'006a0c17    51                      push ecx
'006a0c18    8d55d4                  lea edx, dword ptr [ebp-2c]
'006a0c1b    52                      push edx
'006a0c1c    8d45d8                  lea eax, dword ptr [ebp-28]
'006a0c1f    50                      push eax
'006a0c20    8d4ddc                  lea ecx, dword ptr [ebp-24]
'006a0c23    51                      push ecx
'006a0c24    8d55e0                  lea edx, dword ptr [ebp-20]
'006a0c27    52                      push edx
'006a0c28    6a07                    push 07

' *** Reference to "__vbaFreeStrList"
'006a0c2a    ff155c124000            call dword ptr [0040125c]
'#Cleanup( 0, -4600, -4604, -4608, -4612, -4616, -4620)
'006a0c30    8d45c0                  lea eax, dword ptr [ebp-40]
'006a0c33    50                      push eax
'006a0c34    8d4dc4                  lea ecx, dword ptr [ebp-3c]
'006a0c37    51                      push ecx
'006a0c38    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'006a0c3a    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_9, var_5)
'006a0c40    8d9530ffffff            lea edx, dword ptr [ebp+ffffff30]
'006a0c46    52                      push edx
'006a0c47    8d8540ffffff            lea eax, dword ptr [ebp+ffffff40]
'006a0c4d    50                      push eax
'006a0c4e    8d8d50ffffff            lea ecx, dword ptr [ebp+ffffff50]
'006a0c54    51                      push ecx
'006a0c55    8d9560ffffff            lea edx, dword ptr [ebp+ffffff60]
'006a0c5b    52                      push edx
'006a0c5c    8d8570ffffff            lea eax, dword ptr [ebp+ffffff70]
'006a0c62    50                      push eax
'006a0c63    8d4d80                  lea ecx, dword ptr [ebp-80]
'006a0c66    51                      push ecx
'006a0c67    8d5590                  lea edx, dword ptr [ebp-70]
'006a0c6a    52                      push edx
'006a0c6b    8d45a0                  lea eax, dword ptr [ebp-60]
'006a0c6e    50                      push eax
'006a0c6f    8d4db0                  lea ecx, dword ptr [ebp-50]
'006a0c72    51                      push ecx
'006a0c73    6a09                    push 09

' *** Reference to "__vbaFreeVarList"
'006a0c75    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_6, var_7, var_8, var_18, var_19, var_20, var_21, var_22, var_23)
'006a0c7b    83c454                  add esp, 54
'006a0c7e    c3                      ret
'006a0c7f    c3                      ret
'006a0c80    8b4508                  mov eax, dword ptr [ebp+08]
'006a0c83    8b10                    mov edx, dword ptr [eax]
'006a0c85    50                      push eax
'006a0c86    ff5208                  call dword ptr [edx+08]
'006a0c89    8b45f4                  mov eax, dword ptr [ebp-0c]
'006a0c8c    8b4de4                  mov ecx, dword ptr [ebp-1c]
    'Reference to '__except_list'
'006a0c8f    64890d00000000          mov dword ptr fs:[00000000], ecx
'006a0c96    5f                      pop edi
'006a0c97    5e                      pop esi
'006a0c98    5b                      pop ebx
'006a0c99    8be5                    mov esp, ebp
'006a0c9b    5d                      pop ebp
'006a0c9c    c20400                  ret 0004


End Sub


'Event for btnOk
Private Sub btnOk_Click()
'0069f120    55                      push ebp
'0069f121    8bec                    mov ebp, esp
'0069f123    83ec14                  sub esp, 14

' *** Reference to "__vbaExceptHandler"
'0069f126    6866724000              push 00407266
'0069f12b    64a100000000            mov ax, word ptr fs:[00000000]
'0069f131    50                      push eax
    'Reference to '__except_list'
'0069f132    64892500000000          mov dword ptr fs:[00000000], esp
'0069f139    81ec04020000            sub esp, 00000204
'0069f13f    53                      push ebx
'0069f140    56                      push esi
'0069f141    57                      push edi
'0069f142    8965ec                  mov dword ptr [ebp-14], esp
'0069f145    c745f0c8614000          mov dword ptr [ebp-10], 004061c8
'0069f14c    8b7508                  mov esi, dword ptr [ebp+08]
'0069f14f    8bc6                    mov eax, esi
'0069f151    83e001                  and eax, 01
'0069f154    8945f4                  mov dword ptr [ebp-0c], eax
'0069f157    83e6fe                  and esi, fffffffe
'0069f15a    897508                  mov dword ptr [ebp+08], esi
'0069f15d    33db                    xor ebx, ebx
'0069f15f    895df8                  mov dword ptr [ebp-08], ebx
'0069f162    8b0e                    mov ecx, dword ptr [esi]
'0069f164    56                      push esi
'0069f165    ff5104                  call dword ptr [ecx+04]
'0069f168    895de0                  mov dword ptr [ebp-20], ebx
'0069f16b    895ddc                  mov dword ptr [ebp-24], ebx
'0069f16e    895dd8                  mov dword ptr [ebp-28], ebx
'0069f171    895dd4                  mov dword ptr [ebp-2c], ebx
'0069f174    895dd0                  mov dword ptr [ebp-30], ebx
'0069f177    895dcc                  mov dword ptr [ebp-34], ebx
'0069f17a    895dc8                  mov dword ptr [ebp-38], ebx
'0069f17d    895dc4                  mov dword ptr [ebp-3c], ebx
'0069f180    895dc0                  mov dword ptr [ebp-40], ebx
'0069f183    895dbc                  mov dword ptr [ebp-44], ebx
'0069f186    895db8                  mov dword ptr [ebp-48], ebx
'0069f189    895da8                  mov dword ptr [ebp-58], ebx
'0069f18c    895d98                  mov dword ptr [ebp-68], ebx
'0069f18f    895d88                  mov dword ptr [ebp-78], ebx
'0069f192    899d78ffffff            mov dword ptr [ebp+ffffff78], ebx
'0069f198    899d68ffffff            mov dword ptr [ebp+ffffff68], ebx
'0069f19e    899d58ffffff            mov dword ptr [ebp+ffffff58], ebx
'0069f1a4    899d48ffffff            mov dword ptr [ebp+ffffff48], ebx
'0069f1aa    899d38ffffff            mov dword ptr [ebp+ffffff38], ebx
'0069f1b0    899d28ffffff            mov dword ptr [ebp+ffffff28], ebx
'0069f1b6    899d18ffffff            mov dword ptr [ebp+ffffff18], ebx
'0069f1bc    899d08ffffff            mov dword ptr [ebp+ffffff08], ebx
'0069f1c2    899df8feffff            mov dword ptr [ebp+fffffef8], ebx
'0069f1c8    899de8feffff            mov dword ptr [ebp+fffffee8], ebx
'0069f1ce    899dd8feffff            mov dword ptr [ebp+fffffed8], ebx
'0069f1d4    899db8feffff            mov dword ptr [ebp+fffffeb8], ebx
'0069f1da    899d98feffff            mov dword ptr [ebp+fffffe98], ebx
'0069f1e0    899d78feffff            mov dword ptr [ebp+fffffe78], ebx
'0069f1e6    899d54feffff            mov dword ptr [ebp+fffffe54], ebx
'0069f1ec    66391d28b07200          cmp word ptr [0072b028], bx
'0069f1f3    7508                    jne 69f1fd
'0069f1f5    6a01                    push 01

' *** Reference to "__vbaOnError"
'0069f1f7    ff15b0104000            call dword ptr [004010b0]
On Error Goto handler_0
'0069f1fd    8b06                    mov eax, dword ptr [esi]
'0069f1ff    56                      push esi
'0069f200    ff9004030000            call dword ptr [eax+00000304]
'0069f206    50                      push eax
'0069f207    8d4dc4                  lea ecx, dword ptr [ebp-3c]
'0069f20a    51                      push ecx

' *** Reference to "__vbaObjSet"
'0069f20b    8b3db4104000            mov edi, dword ptr [004010b4]
'0069f211    ffd7                    call edi
Set var_9 = Nothing
'0069f213    898540feffff            mov dword ptr [ebp+fffffe40], eax
'0069f219    8b10                    mov edx, dword ptr [eax]
'0069f21b    8d4de0                  lea ecx, dword ptr [ebp-20]
'0069f21e    51                      push ecx
'0069f21f    50                      push eax
'0069f220    ff92a0000000            call dword ptr [edx+000000a0]
'0069f226    dbe2                    fnclex
'0069f228    3bc3                    cmp eax, ebx
'0069f22a    7d18                    jge 69f244

If (var_9 < 0) Then
'0069f22c    68a0000000              push 000000a0
'0069f231    68d00d4300              push 00430dd0
'0069f236    8b9540feffff            mov edx, dword ptr [ebp+fffffe40]
'0069f23c    52                      push edx
'0069f23d    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0069f23e    ff1580104000            call dword ptr [00401080]
    
End If
'0069f244    8b45e0                  mov eax, dword ptr [ebp-20]
'0069f247    50                      push eax
'0069f248    68cc134300              push 004313cc

' *** Reference to "__vbaStrCmp"
'0069f24d    ff1530114000            call dword ptr [00401130]
'0069f253    8bd8                    mov ebx, eax
'0069f255    f7db                    neg ebx
'0069f257    1bdb                    sbb ebx, ebx
'0069f259    f7db                    neg ebx
'0069f25b    f7db                    neg ebx
'0069f25d    8d4de0                  lea ecx, dword ptr [ebp-20]

' *** Reference to "__vbaFreeStr"
'0069f260    ff150c134000            call dword ptr [0040130c]
'#Cleanup(var_3)
'0069f266    8d4dc4                  lea ecx, dword ptr [ebp-3c]

' *** Reference to "__vbaFreeObj"
'0069f269    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_9)
'0069f26f    6685db                  test ebx, ebx
'0069f272    0f8430120000            je 6a04a8

If (((vbNullString) <> (vbNullChar))) Then
'0069f278    8b0e                    mov ecx, dword ptr [esi]
'0069f27a    56                      push esi
'0069f27b    ff9104030000            call dword ptr [ecx+00000304]
'0069f281    8945b0                  mov dword ptr [ebp-50], eax
'0069f284    c745a809000000          mov dword ptr [ebp-58], 00000009
'0069f28b    8d55a8                  lea edx, dword ptr [ebp-58]
'0069f28e    52                      push edx

' *** Reference to "rtcIsNumeric"
'0069f28f    ff154c114000            call dword ptr [0040114c]
'0069f295    33db                    xor ebx, ebx
    var_num2 = Empty
'0069f297    668bd8                  mov bx, ax
'0069f29a    f7d3                    not ebx
'0069f29c    8d4da8                  lea ecx, dword ptr [ebp-58]

' *** Reference to "__vbaFreeVar"
'0069f29f    ff1528104000            call dword ptr [00401028]
    '#Cleanup(var_86)
'0069f2a5    6685db                  test ebx, ebx
'0069f2a8    0f848c040000            je 69f73a
    
    If (    IsNumeric(((vbNullString) [#!#] (vbNullChar)))) Then
'0069f2ae    a124be7200              mov ax, word ptr [0072be24]
'0069f2b3    85c0                    test ax, ax
'0069f2b5    7510                    jne 69f2c7
'0069f2b7    6824be7200              push 0072be24
'0069f2bc    6870004300              push 00430070

' *** Reference to "__vbaNew2"
'0069f2c1    ff152c124000            call dword ptr [0040122c]
    Dim var_2 As New Global
'0069f2c7    8b1d24be7200            mov ebx, dword ptr [0072be24]
'0069f2cd    8b03                    mov eax, dword ptr [ebx]
'0069f2cf    8d4dc4                  lea ecx, dword ptr [ebp-3c]
'0069f2d2    51                      push ecx
'0069f2d3    53                      push ebx
'0069f2d4    ff5018                  call dword ptr [eax+18]
    Set var_9 = var_2.Screen
'0069f2d7    dbe2                    fnclex
'0069f2d9    85c0                    test ax, ax
'0069f2db    7d0f                    jge 69f2ec
'0069f2dd    6a18                    push 18
'0069f2df    6860004300              push 00430060
'0069f2e4    53                      push ebx
'0069f2e5    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0069f2e6    ff1580104000            call dword ptr [00401080]
'0069f2ec    8b45c4                  mov eax, dword ptr [ebp-3c]
'0069f2ef    898538feffff            mov dword ptr [ebp+fffffe38], eax
'0069f2f5    8b18                    mov ebx, dword ptr [eax]
'0069f2f7    33c9                    xor ecx, ecx

' *** Reference to "__vbaI2I4"
'0069f2f9    ff1550114000            call dword ptr [00401150]
'0069f2ff    50                      push eax
'0069f300    8bd3                    mov edx, ebx
'0069f302    8b9d38feffff            mov ebx, dword ptr [ebp+fffffe38]
'0069f308    53                      push ebx
'0069f309    ff527c                  call dword ptr [edx+7c]
    var_9.MousePointer = CInt(0)
'0069f30c    dbe2                    fnclex
'0069f30e    85c0                    test ax, ax
'0069f310    7d0f                    jge 69f321
'0069f312    6a7c                    push 7c
'0069f314    6810be4300              push 0043be10
'0069f319    53                      push ebx
'0069f31a    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0069f31b    ff1580104000            call dword ptr [00401080]
'0069f321    8d4dc4                  lea ecx, dword ptr [ebp-3c]

' *** Reference to "__vbaFreeObj"
'0069f324    ff1508134000            call dword ptr [00401308]
    '#Cleanup(var_9)
'0069f32a    8b06                    mov eax, dword ptr [esi]
'0069f32c    8d4de0                  lea ecx, dword ptr [ebp-20]
'0069f32f    51                      push ecx
'0069f330    56                      push esi
'0069f331    ff5048                  call dword ptr [eax+48]
'0069f334    dbe2                    fnclex
'0069f336    85c0                    test ax, ax
'0069f338    7d0f                    jge 69f349
'0069f33a    6a48                    push 48
'0069f33c    6838f64400              push 0044f638
'0069f341    56                      push esi
'0069f342    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0069f343    ff1580104000            call dword ptr [00401080]
'0069f349    b904000280              mov ecx, 80020004
'0069f34e    894d80                  mov dword ptr [ebp-80], ecx
'0069f351    b80a000000              mov eax, 0000000a
'0069f356    898578ffffff            mov dword ptr [ebp+ffffff78], eax
'0069f35c    894d90                  mov dword ptr [ebp-70], ecx
'0069f35f    894588                  mov dword ptr [ebp-78], eax
'0069f362    8b45e0                  mov eax, dword ptr [ebp-20]
'0069f365    c745e000000000          mov dword ptr [ebp-20], 00000000
'0069f36c    8945a0                  mov dword ptr [ebp-60], eax
'0069f36f    bb08000000              mov ebx, 00000008
'0069f374    895d98                  mov dword ptr [ebp-68], ebx
'0069f377    c78520ffffff7c244500    mov dword ptr [ebp+ffffff20], 0045247c
'0069f381    899d18ffffff            mov dword ptr [ebp+ffffff18], ebx
'0069f387    8d9518ffffff            lea edx, dword ptr [ebp+ffffff18]
'0069f38d    8d4da8                  lea ecx, dword ptr [ebp-58]

' *** Reference to "__vbaVarDup"
'0069f390    ff1598124000            call dword ptr [00401298]
    var_86 = ("Le nombre d'objet doit être entrée en tant que valeur numérique")
'0069f396    8d9578ffffff            lea edx, dword ptr [ebp+ffffff78]
'0069f39c    52                      push edx
'0069f39d    8d4588                  lea eax, dword ptr [ebp-78]
'0069f3a0    50                      push eax
'0069f3a1    8d4d98                  lea ecx, dword ptr [ebp-68]
'0069f3a4    51                      push ecx
'0069f3a5    6a30                    push 30
'0069f3a7    8d55a8                  lea edx, dword ptr [ebp-58]
'0069f3aa    52                      push edx

' *** Reference to "rtcMsgBox"
'0069f3ab    ff15b8104000            call dword ptr [004010b8]
    var_14 = MsgBox(var_86, 48, vbNullString)
'0069f3b1    8d8578ffffff            lea eax, dword ptr [ebp+ffffff78]
'0069f3b7    50                      push eax
'0069f3b8    8d4d88                  lea ecx, dword ptr [ebp-78]
'0069f3bb    51                      push ecx
'0069f3bc    8d5598                  lea edx, dword ptr [ebp-68]
'0069f3bf    52                      push edx
'0069f3c0    8d45a8                  lea eax, dword ptr [ebp-58]
'0069f3c3    50                      push eax
'0069f3c4    6a04                    push 04

' *** Reference to "__vbaFreeVarList"
'0069f3c6    ff1540104000            call dword ptr [00401040]
    '#Cleanup( var_86, var_130, var_131, var_87)
'0069f3cc    83c414                  add esp, 14
'0069f3cf    8b0e                    mov ecx, dword ptr [esi]
'0069f3d1    56                      push esi
'0069f3d2    ff9104030000            call dword ptr [ecx+00000304]
'0069f3d8    50                      push eax
'0069f3d9    8d55c4                  lea edx, dword ptr [ebp-3c]
'0069f3dc    52                      push edx
'0069f3dd    ffd7                    call edi
    Set var_9 = 
'0069f3df    8bf0                    mov esi, eax
'0069f3e1    8b06                    mov eax, dword ptr [esi]
'0069f3e3    56                      push esi
'0069f3e4    ff9004020000            call dword ptr [eax+00000204]
'0069f3ea    dbe2                    fnclex
'0069f3ec    85c0                    test ax, ax
'0069f3ee    0f8d08120000            jge 6a05fc
'0069f3f4    e9f1110000              jmp 6a05ea

' *** Reference to "rtcErrObj"
'0069f3f9    8b3d6c124000            mov edi, dword ptr [0040126c]
'0069f3ff    ffd7                    call edi
'0069f401    50                      push eax
'0069f402    8d55c4                  lea edx, dword ptr [ebp-3c]
'0069f405    52                      push edx

' *** Reference to "__vbaObjSet"
'0069f406    8b1db4104000            mov ebx, dword ptr [004010b4]
'0069f40c    ffd3                    call ebx
    Set var_9 = Err
'0069f40e    8bf0                    mov esi, eax
'0069f410    8b06                    mov eax, dword ptr [esi]
'0069f412    8d4de0                  lea ecx, dword ptr [ebp-20]
'0069f415    51                      push ecx
'0069f416    56                      push esi
'0069f417    ff502c                  call dword ptr [eax+2c]
    var_3 = var_9.Description
'0069f41a    dbe2                    fnclex
'0069f41c    85c0                    test ax, ax
'0069f41e    7d0f                    jge 69f42f
'0069f420    6a2c                    push 2c
'0069f422    685c084300              push 0043085c
'0069f427    56                      push esi
'0069f428    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0069f429    ff1580104000            call dword ptr [00401080]
'0069f42f    ffd7                    call edi
'0069f431    50                      push eax
'0069f432    8d55c0                  lea edx, dword ptr [ebp-40]
'0069f435    52                      push edx
'0069f436    ffd3                    call ebx
    Set var_5 = Err
'0069f438    8bf0                    mov esi, eax
'0069f43a    8b06                    mov eax, dword ptr [esi]
'0069f43c    8d8d54feffff            lea ecx, dword ptr [ebp+fffffe54]
'0069f442    51                      push ecx
'0069f443    56                      push esi
'0069f444    ff501c                  call dword ptr [eax+1c]
    var_376 = var_5.Number
'0069f447    dbe2                    fnclex
'0069f449    85c0                    test ax, ax
'0069f44b    7d0f                    jge 69f45c
'0069f44d    6a1c                    push 1c
'0069f44f    685c084300              push 0043085c
'0069f454    56                      push esi
'0069f455    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0069f456    ff1580104000            call dword ptr [00401080]
'0069f45c    b804000280              mov eax, 80020004
'0069f461    894590                  mov dword ptr [ebp-70], eax
'0069f464    b90a000000              mov ecx, 0000000a
'0069f469    894d88                  mov dword ptr [ebp-78], ecx
'0069f46c    8945a0                  mov dword ptr [ebp-60], eax
'0069f46f    894d98                  mov dword ptr [ebp-68], ecx
'0069f472    c78520ffffff24b07200    mov dword ptr [ebp+ffffff20], 0072b024
'0069f47c    c78518ffffff08400000    mov dword ptr [ebp+ffffff18], 00004008
'0069f486    6814084300              push 00430814
'0069f48b    8b55e0                  mov edx, dword ptr [ebp-20]
'0069f48e    52                      push edx

' *** Reference to "__vbaStrCat"
'0069f48f    8b3570104000            mov esi, dword ptr [00401070]
'0069f495    ffd6                    call esi
    var_26 = ("L'erreur suivante s'est produite : ") & (var_3)
'0069f497    8bd0                    mov edx, eax
'0069f499    8d4ddc                  lea ecx, dword ptr [ebp-24]

' *** Reference to "__vbaStrMove"
'0069f49c    8b3dd0124000            mov edi, dword ptr [004012d0]
'0069f4a2    ffd7                    call edi
'0069f4a4    50                      push eax
'0069f4a5    6870084300              push 00430870
'0069f4aa    ffd6                    call esi
    var_17 = (var_26) & (vbCrLf)
'0069f4ac    8bd0                    mov edx, eax
'0069f4ae    8d4dd8                  lea ecx, dword ptr [ebp-28]
'0069f4b1    ffd7                    call edi
'0069f4b3    50                      push eax
'0069f4b4    6870084300              push 00430870
'0069f4b9    ffd6                    call esi
    var_129 = (var_17) & (vbCrLf)
'0069f4bb    8bd0                    mov edx, eax
'0069f4bd    8d4dd4                  lea ecx, dword ptr [ebp-2c]
'0069f4c0    ffd7                    call edi
'0069f4c2    50                      push eax
'0069f4c3    6880084300              push 00430880
'0069f4c8    ffd6                    call esi
    var_53 = (var_129) & (" numéro d'erreur (")
'0069f4ca    8bd0                    mov edx, eax
'0069f4cc    8d4dd0                  lea ecx, dword ptr [ebp-30]
'0069f4cf    ffd7                    call edi
'0069f4d1    50                      push eax
'0069f4d2    8b8554feffff            mov eax, dword ptr [ebp+fffffe54]
'0069f4d8    50                      push eax

' *** Reference to "__vbaStrI4"
'0069f4d9    ff1520104000            call dword ptr [00401020]
'0069f4df    8bd0                    mov edx, eax
'0069f4e1    8d4dcc                  lea ecx, dword ptr [ebp-34]
'0069f4e4    ffd7                    call edi
'0069f4e6    50                      push eax
'0069f4e7    ffd6                    call esi
    var_54 = (var_53) & (CStr(var_376))
'0069f4e9    8bd0                    mov edx, eax
'0069f4eb    8d4dc8                  lea ecx, dword ptr [ebp-38]
'0069f4ee    ffd7                    call edi
'0069f4f0    50                      push eax
'0069f4f1    68ac084300              push 004308ac
'0069f4f6    ffd6                    call esi
    var_55 = (var_54) & (" )")
'0069f4f8    8945b0                  mov dword ptr [ebp-50], eax
'0069f4fb    bb08000000              mov ebx, 00000008
'0069f500    895da8                  mov dword ptr [ebp-58], ebx
'0069f503    8d4d88                  lea ecx, dword ptr [ebp-78]
'0069f506    51                      push ecx
'0069f507    8d5598                  lea edx, dword ptr [ebp-68]
'0069f50a    52                      push edx
'0069f50b    8d8518ffffff            lea eax, dword ptr [ebp+ffffff18]
'0069f511    50                      push eax
'0069f512    6a10                    push 10
'0069f514    8d4da8                  lea ecx, dword ptr [ebp-58]
'0069f517    51                      push ecx

' *** Reference to "rtcMsgBox"
'0069f518    ff15b8104000            call dword ptr [004010b8]
    var_56 = MsgBox(var_55, 16, vbNullString)
'0069f51e    8d55c8                  lea edx, dword ptr [ebp-38]
'0069f521    52                      push edx
'0069f522    8d45cc                  lea eax, dword ptr [ebp-34]
'0069f525    50                      push eax
'0069f526    8d4dd0                  lea ecx, dword ptr [ebp-30]
'0069f529    51                      push ecx
'0069f52a    8d55d4                  lea edx, dword ptr [ebp-2c]
'0069f52d    52                      push edx
'0069f52e    8d45d8                  lea eax, dword ptr [ebp-28]
'0069f531    50                      push eax
'0069f532    8d4ddc                  lea ecx, dword ptr [ebp-24]
'0069f535    51                      push ecx
'0069f536    8d55e0                  lea edx, dword ptr [ebp-20]
'0069f539    52                      push edx
'0069f53a    6a07                    push 07

' *** Reference to "__vbaFreeStrList"
'0069f53c    ff155c124000            call dword ptr [0040125c]
    '#Cleanup( var_3, -4540, -4544, -4548, -4552, -4556, -4560)
'0069f542    8d45c0                  lea eax, dword ptr [ebp-40]
'0069f545    50                      push eax
'0069f546    8d4dc4                  lea ecx, dword ptr [ebp-3c]
'0069f549    51                      push ecx
'0069f54a    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'0069f54c    8b3d4c104000            mov edi, dword ptr [0040104c]
'0069f552    ffd7                    call edi
    '#Cleanup( var_9, var_5)
'0069f554    8d5588                  lea edx, dword ptr [ebp-78]
'0069f557    52                      push edx
'0069f558    8d4598                  lea eax, dword ptr [ebp-68]
'0069f55b    50                      push eax
'0069f55c    8d4da8                  lea ecx, dword ptr [ebp-58]
'0069f55f    51                      push ecx
'0069f560    6a03                    push 03

' *** Reference to "__vbaFreeVarList"
'0069f562    ff1540104000            call dword ptr [00401040]
    '#Cleanup( var_86, var_130, var_131)
'0069f568    83c43c                  add esp, 3c
'0069f56b    8d55a8                  lea edx, dword ptr [ebp-58]
'0069f56e    52                      push edx

' *** Reference to "rtcGetPresentDate"
'0069f56f    ff15f4124000            call dword ptr [004012f4]
    var_55 = Now()
'0069f575    c78520ffffffb8084300    mov dword ptr [ebp+ffffff20], 004308b8
'0069f57f    899d18ffffff            mov dword ptr [ebp+ffffff18], ebx
'0069f585    8d9518ffffff            lea edx, dword ptr [ebp+ffffff18]
'0069f58b    8d4d98                  lea ecx, dword ptr [ebp-68]

' *** Reference to "__vbaVarDup"
'0069f58e    ff1598124000            call dword ptr [00401298]
    var_130 = ("dd/MM/yyyy hh:mm:ss")
'0069f594    6a01                    push 01
'0069f596    6a01                    push 01
'0069f598    8d4598                  lea eax, dword ptr [ebp-68]
'0069f59b    50                      push eax
'0069f59c    8d4da8                  lea ecx, dword ptr [ebp-58]
'0069f59f    51                      push ecx
'0069f5a0    8d5588                  lea edx, dword ptr [ebp-78]
'0069f5a3    52                      push edx

' *** Reference to "rtcVarFromFormatVar"
'0069f5a4    ff1574104000            call dword ptr [00401074]

' *** Reference to "rtcErrObj"
'0069f5aa    ff156c124000            call dword ptr [0040126c]
'0069f5b0    50                      push eax
'0069f5b1    8d45c4                  lea eax, dword ptr [ebp-3c]
'0069f5b4    50                      push eax

' *** Reference to "__vbaObjSet"
'0069f5b5    ff15b4104000            call dword ptr [004010b4]
    Set var_9 = Err
'0069f5bb    8bf0                    mov esi, eax
'0069f5bd    8b0e                    mov ecx, dword ptr [esi]
'0069f5bf    8d9554feffff            lea edx, dword ptr [ebp+fffffe54]
'0069f5c5    52                      push edx
'0069f5c6    56                      push esi
'0069f5c7    ff511c                  call dword ptr [ecx+1c]
    var_376 = var_9.Number
'0069f5ca    dbe2                    fnclex
'0069f5cc    85c0                    test ax, ax
'0069f5ce    7d0f                    jge 69f5df
'0069f5d0    6a1c                    push 1c
'0069f5d2    685c084300              push 0043085c
'0069f5d7    56                      push esi
'0069f5d8    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0069f5d9    ff1580104000            call dword ptr [00401080]

' *** Reference to "rtcErrObj"
'0069f5df    ff156c124000            call dword ptr [0040126c]
'0069f5e5    50                      push eax
'0069f5e6    8d45c0                  lea eax, dword ptr [ebp-40]
'0069f5e9    50                      push eax

' *** Reference to "__vbaObjSet"
'0069f5ea    ff15b4104000            call dword ptr [004010b4]
    Set var_5 = Err
'0069f5f0    8bf0                    mov esi, eax
'0069f5f2    8b0e                    mov ecx, dword ptr [esi]
'0069f5f4    8d55e0                  lea edx, dword ptr [ebp-20]
'0069f5f7    52                      push edx
'0069f5f8    56                      push esi
'0069f5f9    ff512c                  call dword ptr [ecx+2c]
    var_3 = var_5.Description
'0069f5fc    dbe2                    fnclex
'0069f5fe    85c0                    test ax, ax
'0069f600    7d0f                    jge 69f611
'0069f602    6a2c                    push 2c
'0069f604    685c084300              push 0043085c
'0069f609    56                      push esi
'0069f60a    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0069f60b    ff1580104000            call dword ptr [00401080]
'0069f611    c78510ffffffe4084300    mov dword ptr [ebp+ffffff10], 004308e4
'0069f61b    899d08ffffff            mov dword ptr [ebp+ffffff08], ebx
'0069f621    8b8554feffff            mov eax, dword ptr [ebp+fffffe54]
'0069f627    898500ffffff            mov dword ptr [ebp+ffffff00], eax
'0069f62d    c785f8feffff03000000    mov dword ptr [ebp+fffffef8], 00000003
'0069f637    c785f0feffff08094300    mov dword ptr [ebp+fffffef0], 00430908
'0069f641    899de8feffff            mov dword ptr [ebp+fffffee8], ebx
'0069f647    8b45e0                  mov eax, dword ptr [ebp-20]
'0069f64a    c745e000000000          mov dword ptr [ebp-20], 00000000
'0069f651    898550ffffff            mov dword ptr [ebp+ffffff50], eax
'0069f657    899d48ffffff            mov dword ptr [ebp+ffffff48], ebx
'0069f65d    c785e0feffff98254500    mov dword ptr [ebp+fffffee0], 00452598
'0069f667    899dd8feffff            mov dword ptr [ebp+fffffed8], ebx
'0069f66d    8d4d88                  lea ecx, dword ptr [ebp-78]
'0069f670    51                      push ecx
'0069f671    8d9508ffffff            lea edx, dword ptr [ebp+ffffff08]
'0069f677    52                      push edx
'0069f678    8d8578ffffff            lea eax, dword ptr [ebp+ffffff78]
'0069f67e    50                      push eax

' *** Reference to "__vbaVarCat"
'0069f67f    8b3508124000            mov esi, dword ptr [00401208]
'0069f685    ffd6                    call esi
'0069f687    50                      push eax
'0069f688    8d8df8feffff            lea ecx, dword ptr [ebp+fffffef8]
'0069f68e    51                      push ecx
'0069f68f    8d9568ffffff            lea edx, dword ptr [ebp+ffffff68]
'0069f695    52                      push edx
'0069f696    ffd6                    call esi
'0069f698    50                      push eax
'0069f699    8d85e8feffff            lea eax, dword ptr [ebp+fffffee8]
'0069f69f    50                      push eax
'0069f6a0    8d8d58ffffff            lea ecx, dword ptr [ebp+ffffff58]
'0069f6a6    51                      push ecx
'0069f6a7    ffd6                    call esi
'0069f6a9    50                      push eax
'0069f6aa    8d9548ffffff            lea edx, dword ptr [ebp+ffffff48]
'0069f6b0    52                      push edx
'0069f6b1    8d8538ffffff            lea eax, dword ptr [ebp+ffffff38]
'0069f6b7    50                      push eax
'0069f6b8    ffd6                    call esi
'0069f6ba    50                      push eax
'0069f6bb    8d8dd8feffff            lea ecx, dword ptr [ebp+fffffed8]
'0069f6c1    51                      push ecx
'0069f6c2    8d9528ffffff            lea edx, dword ptr [ebp+ffffff28]
'0069f6c8    52                      push edx
'0069f6c9    ffd6                    call esi
'0069f6cb    50                      push eax
'0069f6cc    33c0                    xor eax, eax
'0069f6ce    66a12ab07200            mov eax, dword ptr [0072b02a]
'0069f6d4    50                      push eax
'0069f6d5    6884094300              push 00430984

' *** Reference to "__vbaPrintFile"
'0069f6da    ff15b8114000            call dword ptr [004011b8]
    Print #0, 
'0069f6e0    8d4dc0                  lea ecx, dword ptr [ebp-40]
'0069f6e3    51                      push ecx
'0069f6e4    8d55c4                  lea edx, dword ptr [ebp-3c]
'0069f6e7    52                      push edx
'0069f6e8    6a02                    push 02
'0069f6ea    ffd7                    call edi
    '#Cleanup( var_9, var_5)
'0069f6ec    8d8528ffffff            lea eax, dword ptr [ebp+ffffff28]
'0069f6f2    50                      push eax
'0069f6f3    8d8d38ffffff            lea ecx, dword ptr [ebp+ffffff38]
'0069f6f9    51                      push ecx
'0069f6fa    8d9548ffffff            lea edx, dword ptr [ebp+ffffff48]
'0069f700    52                      push edx
'0069f701    8d8558ffffff            lea eax, dword ptr [ebp+ffffff58]
'0069f707    50                      push eax
'0069f708    8d8d68ffffff            lea ecx, dword ptr [ebp+ffffff68]
'0069f70e    51                      push ecx
'0069f70f    8d9578ffffff            lea edx, dword ptr [ebp+ffffff78]
'0069f715    52                      push edx
'0069f716    8d4588                  lea eax, dword ptr [ebp-78]
'0069f719    50                      push eax
'0069f71a    8d4d98                  lea ecx, dword ptr [ebp-68]
'0069f71d    51                      push ecx
'0069f71e    8d55a8                  lea edx, dword ptr [ebp-58]
'0069f721    52                      push edx
'0069f722    6a09                    push 09

' *** Reference to "__vbaFreeVarList"
'0069f724    ff1540104000            call dword ptr [00401040]
    '#Cleanup( var_86, var_130, var_131, var_87, var_132, var_133, var_134, var_135, var_136)
'0069f72a    83c440                  add esp, 40
'0069f72d    6a00                    push 00

' *** Reference to "__vbaResume"
'0069f72f    ff1568104000            call dword ptr [00401068]
'0069f735    e9cb0e0000              jmp 6a0605
    Resume handler_6A0605
End If
'0069f73a    8b0e                    mov ecx, dword ptr [esi]
'0069f73c    56                      push esi
'0069f73d    ff9100030000            call dword ptr [ecx+00000300]
'0069f743    50                      push eax
'0069f744    8d55c4                  lea edx, dword ptr [ebp-3c]
'0069f747    52                      push edx
'0069f748    ffd7                    call edi
Set var_9 = IsNumeric(((vbNullString) [##] (vbNullChar)))
'0069f74a    8bd8                    mov ebx, eax
'0069f74c    8b03                    mov eax, dword ptr [ebx]
'0069f74e    8d4de0                  lea ecx, dword ptr [ebp-20]
'0069f751    51                      push ecx
'0069f752    53                      push ebx
'0069f753    ff90a0000000            call dword ptr [eax+000000a0]
'0069f759    dbe2                    fnclex
'0069f75b    85c0                    test ax, ax
'0069f75d    7d12                    jge 69f771
'0069f75f    68a0000000              push 000000a0
'0069f764    68d00d4300              push 00430dd0
'0069f769    53                      push ebx
'0069f76a    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0069f76b    ff1580104000            call dword ptr [00401080]
'0069f771    8b55e0                  mov edx, dword ptr [ebp-20]
'0069f774    52                      push edx
'0069f775    68cc134300              push 004313cc

' *** Reference to "__vbaStrCmp"
'0069f77a    ff1530114000            call dword ptr [00401130]
'0069f780    8bd8                    mov ebx, eax
'0069f782    f7db                    neg ebx
'0069f784    1bdb                    sbb ebx, ebx
'0069f786    f7db                    neg ebx
'0069f788    f7db                    neg ebx
'0069f78a    8d4de0                  lea ecx, dword ptr [ebp-20]

' *** Reference to "__vbaFreeStr"
'0069f78d    ff150c134000            call dword ptr [0040130c]
'#Cleanup(var_3)
'0069f793    8d4dc4                  lea ecx, dword ptr [ebp-3c]

' *** Reference to "__vbaFreeObj"
'0069f796    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_9)
'0069f79c    6685db                  test ebx, ebx
'0069f79f    8b06                    mov eax, dword ptr [esi]
'0069f7a1    56                      push esi
'0069f7a2    0f847e010000            je 69f926

If (((vbNullString) <> (vbNullChar))) Then
'0069f7a8    ff9000030000            call dword ptr [eax+00000300]
'0069f7ae    8945b0                  mov dword ptr [ebp-50], eax
'0069f7b1    c745a809000000          mov dword ptr [ebp-58], 00000009
'0069f7b8    8d4da8                  lea ecx, dword ptr [ebp-58]
'0069f7bb    51                      push ecx

' *** Reference to "rtcIsNumeric"
'0069f7bc    ff154c114000            call dword ptr [0040114c]
'0069f7c2    33db                    xor ebx, ebx
    var_num2 = Empty
'0069f7c4    668bd8                  mov bx, ax
'0069f7c7    f7d3                    not ebx
'0069f7c9    8d4da8                  lea ecx, dword ptr [ebp-58]

' *** Reference to "__vbaFreeVar"
'0069f7cc    ff1528104000            call dword ptr [00401028]
    '#Cleanup(var_86)
'0069f7d2    6685db                  test ebx, ebx
'0069f7d5    0f84ad010000            je 69f988
    
    If (    Not (IsNumeric(FrmLstArticle))) Then
'0069f7db    a124be7200              mov ax, word ptr [0072be24]
'0069f7e0    85c0                    test ax, ax
'0069f7e2    7510                    jne 69f7f4
'0069f7e4    6824be7200              push 0072be24
'0069f7e9    6870004300              push 00430070

' *** Reference to "__vbaNew2"
'0069f7ee    ff152c124000            call dword ptr [0040122c]
    Set var_2 = New Global
'0069f7f4    8b1d24be7200            mov ebx, dword ptr [0072be24]
'0069f7fa    8b13                    mov edx, dword ptr [ebx]
'0069f7fc    8d45c4                  lea eax, dword ptr [ebp-3c]
'0069f7ff    50                      push eax
'0069f800    53                      push ebx
'0069f801    ff5218                  call dword ptr [edx+18]
    Set var_9 = var_2.Screen
'0069f804    dbe2                    fnclex
'0069f806    85c0                    test ax, ax
'0069f808    7d0f                    jge 69f819
'0069f80a    6a18                    push 18
'0069f80c    6860004300              push 00430060
'0069f811    53                      push ebx
'0069f812    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0069f813    ff1580104000            call dword ptr [00401080]
'0069f819    8b45c4                  mov eax, dword ptr [ebp-3c]
'0069f81c    898538feffff            mov dword ptr [ebp+fffffe38], eax
'0069f822    8b18                    mov ebx, dword ptr [eax]
'0069f824    33c9                    xor ecx, ecx

' *** Reference to "__vbaI2I4"
'0069f826    ff1550114000            call dword ptr [00401150]
'0069f82c    50                      push eax
'0069f82d    8bcb                    mov ecx, ebx
'0069f82f    8b9d38feffff            mov ebx, dword ptr [ebp+fffffe38]
'0069f835    53                      push ebx
'0069f836    ff517c                  call dword ptr [ecx+7c]
    var_9.MousePointer = CInt(0)
'0069f839    dbe2                    fnclex
'0069f83b    85c0                    test ax, ax
'0069f83d    7d0f                    jge 69f84e
'0069f83f    6a7c                    push 7c
'0069f841    6810be4300              push 0043be10
'0069f846    53                      push ebx
'0069f847    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0069f848    ff1580104000            call dword ptr [00401080]
'0069f84e    8d4dc4                  lea ecx, dword ptr [ebp-3c]

' *** Reference to "__vbaFreeObj"
'0069f851    ff1508134000            call dword ptr [00401308]
    '#Cleanup(var_9)
'0069f857    8b16                    mov edx, dword ptr [esi]
'0069f859    8d45e0                  lea eax, dword ptr [ebp-20]
'0069f85c    50                      push eax
'0069f85d    56                      push esi
'0069f85e    ff5248                  call dword ptr [edx+48]
'0069f861    dbe2                    fnclex
'0069f863    85c0                    test ax, ax
'0069f865    7d0f                    jge 69f876
'0069f867    6a48                    push 48
'0069f869    6838f64400              push 0044f638
'0069f86e    56                      push esi
'0069f86f    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0069f870    ff1580104000            call dword ptr [00401080]
'0069f876    b904000280              mov ecx, 80020004
'0069f87b    894d80                  mov dword ptr [ebp-80], ecx
'0069f87e    b80a000000              mov eax, 0000000a
'0069f883    898578ffffff            mov dword ptr [ebp+ffffff78], eax
'0069f889    894d90                  mov dword ptr [ebp-70], ecx
'0069f88c    894588                  mov dword ptr [ebp-78], eax
'0069f88f    8b45e0                  mov eax, dword ptr [ebp-20]
'0069f892    c745e000000000          mov dword ptr [ebp-20], 00000000
'0069f899    8945a0                  mov dword ptr [ebp-60], eax
'0069f89c    bb08000000              mov ebx, 00000008
'0069f8a1    895d98                  mov dword ptr [ebp-68], ebx
'0069f8a4    c78520ffffff00254500    mov dword ptr [ebp+ffffff20], 00452500
'0069f8ae    899d18ffffff            mov dword ptr [ebp+ffffff18], ebx
'0069f8b4    8d9518ffffff            lea edx, dword ptr [ebp+ffffff18]
'0069f8ba    8d4da8                  lea ecx, dword ptr [ebp-58]

' *** Reference to "__vbaVarDup"
'0069f8bd    ff1598124000            call dword ptr [00401298]
    var_86 = ("Le prix unitaire de l'objet doit être entrée en tant que valeur numérique")
'0069f8c3    8d8d78ffffff            lea ecx, dword ptr [ebp+ffffff78]
'0069f8c9    51                      push ecx
'0069f8ca    8d5588                  lea edx, dword ptr [ebp-78]
'0069f8cd    52                      push edx
'0069f8ce    8d4598                  lea eax, dword ptr [ebp-68]
'0069f8d1    50                      push eax
'0069f8d2    6a30                    push 30
'0069f8d4    8d4da8                  lea ecx, dword ptr [ebp-58]
'0069f8d7    51                      push ecx

' *** Reference to "rtcMsgBox"
'0069f8d8    ff15b8104000            call dword ptr [004010b8]
    var_1386 = MsgBox(var_86, 48, vbNullString)
'0069f8de    8d9578ffffff            lea edx, dword ptr [ebp+ffffff78]
'0069f8e4    52                      push edx
'0069f8e5    8d4588                  lea eax, dword ptr [ebp-78]
'0069f8e8    50                      push eax
'0069f8e9    8d4d98                  lea ecx, dword ptr [ebp-68]
'0069f8ec    51                      push ecx
'0069f8ed    8d55a8                  lea edx, dword ptr [ebp-58]
'0069f8f0    52                      push edx
'0069f8f1    6a04                    push 04

' *** Reference to "__vbaFreeVarList"
'0069f8f3    ff1540104000            call dword ptr [00401040]
    '#Cleanup( var_86, var_130, var_131, var_87)
'0069f8f9    83c414                  add esp, 14
'0069f8fc    8b06                    mov eax, dword ptr [esi]
'0069f8fe    56                      push esi
'0069f8ff    ff9000030000            call dword ptr [eax+00000300]
'0069f905    50                      push eax
'0069f906    8d4dc4                  lea ecx, dword ptr [ebp-3c]
'0069f909    51                      push ecx
'0069f90a    ffd7                    call edi
    Set var_9 = Nothing
'0069f90c    8bf0                    mov esi, eax
'0069f90e    8b16                    mov edx, dword ptr [esi]
'0069f910    56                      push esi
'0069f911    ff9204020000            call dword ptr [edx+00000204]
'0069f917    dbe2                    fnclex
'0069f919    85c0                    test ax, ax
'0069f91b    0f8ddb0c0000            jge 6a05fc
'0069f921    e9c40c0000              jmp 6a05ea
    
End If
'0069f926    ff9000030000            call dword ptr [eax+00000300]
'0069f92c    50                      push eax
'0069f92d    8d4dc4                  lea ecx, dword ptr [ebp-3c]
'0069f930    51                      push ecx
'0069f931    ffd7                    call edi
Set var_9 = Nothing
'0069f933    898540feffff            mov dword ptr [ebp+fffffe40], eax
'0069f939    8b18                    mov ebx, dword ptr [eax]
'0069f93b    6a00                    push 00

' *** Reference to "__vbaStrI2"
'0069f93d    ff1510104000            call dword ptr [00401010]
'0069f943    8bd0                    mov edx, eax
'0069f945    8d4de0                  lea ecx, dword ptr [ebp-20]

' *** Reference to "__vbaStrMove"
'0069f948    ff15d0124000            call dword ptr [004012d0]
'0069f94e    50                      push eax
'0069f94f    8bd3                    mov edx, ebx
'0069f951    8b9d40feffff            mov ebx, dword ptr [ebp+fffffe40]
'0069f957    53                      push ebx
'0069f958    ff92a4000000            call dword ptr [edx+000000a4]
'0069f95e    dbe2                    fnclex
'0069f960    85c0                    test ax, ax
'0069f962    7d12                    jge 69f976
'0069f964    68a4000000              push 000000a4
'0069f969    68d00d4300              push 00430dd0
'0069f96e    53                      push ebx
'0069f96f    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0069f970    ff1580104000            call dword ptr [00401080]
'0069f976    8d4de0                  lea ecx, dword ptr [ebp-20]

' *** Reference to "__vbaFreeStr"
'0069f979    ff150c134000            call dword ptr [0040130c]
'#Cleanup(var_3)
'0069f97f    8d4dc4                  lea ecx, dword ptr [ebp-3c]

' *** Reference to "__vbaFreeObj"
'0069f982    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_9)
'ERROR: Two many next close:
End If
'0069f988    6a00                    push 00
'0069f98a    6a11                    push 11
'0069f98c    8b06                    mov eax, dword ptr [esi]
'0069f98e    56                      push esi
'0069f98f    ff9014030000            call dword ptr [eax+00000314]
'0069f995    50                      push eax
'0069f996    8d4dc4                  lea ecx, dword ptr [ebp-3c]
'0069f999    51                      push ecx
'0069f99a    ffd7                    call edi
Set var_9 = Nothing
'0069f99c    50                      push eax
'0069f99d    8d55a8                  lea edx, dword ptr [ebp-58]
'0069f9a0    52                      push edx

' *** Reference to "__vbaLateIdCallLd"
'0069f9a1    8b1d7c114000            mov ebx, dword ptr [0040117c]
'0069f9a7    ffd3                    call ebx
var_86 = var_9.UNK_-256 - 20_17
'0069f9a9    83c410                  add esp, 10
'0069f9ac    50                      push eax

' *** Reference to "__vbaI4Var"
'0069f9ad    ff157c124000            call dword ptr [0040127c]
'0069f9b3    33c9                    xor ecx, ecx
'0069f9b5    85c0                    test ax, ax
'0069f9b7    0f9fc1                  setg cl
'0069f9ba    f7d9                    neg ecx
'0069f9bc    66898d40feffff          mov word ptr [ebp+fffffe40], cx
'0069f9c3    8d4dc4                  lea ecx, dword ptr [ebp-3c]

' *** Reference to "__vbaFreeObj"
'0069f9c6    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_9)
'0069f9cc    8d4da8                  lea ecx, dword ptr [ebp-58]

' *** Reference to "__vbaFreeVar"
'0069f9cf    ff1528104000            call dword ptr [00401028]
'#Cleanup(var_86)
'0069f9d5    6683bd40feffff00        cmp word ptr [ebp+fffffe40], 00
'0069f9dd    0f84220c0000            je 6a0605

If (-() <> 0) Then

' *** Reference to "rtcErrObj"
'0069f9e3    ff156c124000            call dword ptr [0040126c]
'0069f9e9    50                      push eax
'0069f9ea    8d55c4                  lea edx, dword ptr [ebp-3c]
'0069f9ed    52                      push edx
'0069f9ee    ffd7                    call edi
    Set var_9 = Err
'0069f9f0    8b08                    mov ecx, dword ptr [eax]
'0069f9f2    50                      push eax
'0069f9f3    ff5148                  call dword ptr [ecx+48]
    Call var_9.Clear()
'0069f9f6    8d4dc4                  lea ecx, dword ptr [ebp-3c]

' *** Reference to "__vbaFreeObj"
'0069f9f9    ff1508134000            call dword ptr [00401308]
    '#Cleanup(var_9)
'0069f9ff    c78520ffffff00000000    mov dword ptr [ebp+ffffff20], 00000000
'0069fa09    c78518ffffff03000000    mov dword ptr [ebp+ffffff18], 00000003
'0069fa13    6a00                    push 00
'0069fa15    6a11                    push 11
'0069fa17    8b16                    mov edx, dword ptr [esi]
'0069fa19    56                      push esi
'0069fa1a    ff9214030000            call dword ptr [edx+00000314]
'0069fa20    50                      push eax
'0069fa21    8d45c4                  lea eax, dword ptr [ebp-3c]
'0069fa24    50                      push eax
'0069fa25    ffd7                    call edi
    Set var_9 = Nothing
'0069fa27    50                      push eax
'0069fa28    8d4da8                  lea ecx, dword ptr [ebp-58]
'0069fa2b    51                      push ecx
'0069fa2c    ffd3                    call ebx
    var_86 = var_9.UNK_-256 - 20_17
'0069fa2e    83c410                  add esp, 10
'0069fa31    50                      push eax

' *** Reference to "__vbaI4Var"
'0069fa32    ff157c124000            call dword ptr [0040127c]
'0069fa38    898500ffffff            mov dword ptr [ebp+ffffff00], eax
'0069fa3e    c785f8feffff03000000    mov dword ptr [ebp+fffffef8], 00000003
'0069fa48    c785e0feffff00000000    mov dword ptr [ebp+fffffee0], 00000000
'0069fa52    c785d8feffff02000000    mov dword ptr [ebp+fffffed8], 00000002
'0069fa5c    a184b07200              mov ax, word ptr [0072b084]
'0069fa61    85c0                    test ax, ax
'0069fa63    7510                    jne 69fa75
'0069fa65    6884b07200              push 0072b084
'0069fa6a    68f8894100              push 004189f8

' *** Reference to "__vbaNew2"
'0069fa6f    ff152c124000            call dword ptr [0040122c]
    Dim var_28 As New frmGestPerso
'0069fa75    83ec10                  sub esp, 10
'0069fa78    8bd4                    mov edx, esp
'0069fa7a    b803000000              mov eax, 00000003
'0069fa7f    8902                    mov dword ptr [edx], eax
'0069fa81    8b85bcfeffff            mov eax, dword ptr [ebp+fffffebc]
'0069fa87    894204                  mov dword ptr [edx+04], eax
'0069fa8a    33c0                    xor eax, eax
    var_num1 = Empty
'0069fa8c    894208                  mov dword ptr [edx+08], eax
'0069fa8f    8b8dc4feffff            mov ecx, dword ptr [ebp+fffffec4]
'0069fa95    894a0c                  mov dword ptr [edx+0c], ecx
'0069fa98    83ec10                  sub esp, 10
'0069fa9b    8bd4                    mov edx, esp
'0069fa9d    b80a000000              mov eax, 0000000a
'0069faa2    8902                    mov dword ptr [edx], eax
'0069faa4    8b859cfeffff            mov eax, dword ptr [ebp+fffffe9c]
'0069faaa    894204                  mov dword ptr [edx+04], eax
'0069faad    b804000280              mov eax, 80020004
'0069fab2    894208                  mov dword ptr [edx+08], eax
'0069fab5    8b8da4feffff            mov ecx, dword ptr [ebp+fffffea4]
'0069fabb    894a0c                  mov dword ptr [edx+0c], ecx
'0069fabe    83ec10                  sub esp, 10
'0069fac1    8bd4                    mov edx, esp
'0069fac3    b802000000              mov eax, 00000002
'0069fac8    8902                    mov dword ptr [edx], eax
'0069faca    8b857cfeffff            mov eax, dword ptr [ebp+fffffe7c]
'0069fad0    894204                  mov dword ptr [edx+04], eax
'0069fad3    33c0                    xor eax, eax
    var_num1 = Empty
'0069fad5    894208                  mov dword ptr [edx+08], eax
'0069fad8    8b8d84feffff            mov ecx, dword ptr [ebp+fffffe84]
'0069fade    894a0c                  mov dword ptr [edx+0c], ecx
'0069fae1    83ec10                  sub esp, 10
'0069fae4    8bd4                    mov edx, esp
'0069fae6    8b8518ffffff            mov eax, dword ptr [ebp+ffffff18]
'0069faec    8902                    mov dword ptr [edx], eax
'0069faee    8b8d1cffffff            mov ecx, dword ptr [ebp+ffffff1c]
'0069faf4    894a04                  mov dword ptr [edx+04], ecx
'0069faf7    8b8520ffffff            mov eax, dword ptr [ebp+ffffff20]
'0069fafd    894208                  mov dword ptr [edx+08], eax
'0069fb00    8b8d24ffffff            mov ecx, dword ptr [ebp+ffffff24]
'0069fb06    894a0c                  mov dword ptr [edx+0c], ecx
'0069fb09    83ec10                  sub esp, 10
'0069fb0c    8bd4                    mov edx, esp
'0069fb0e    8b85f8feffff            mov eax, dword ptr [ebp+fffffef8]
'0069fb14    8902                    mov dword ptr [edx], eax
'0069fb16    8b8dfcfeffff            mov ecx, dword ptr [ebp+fffffefc]
'0069fb1c    894a04                  mov dword ptr [edx+04], ecx
'0069fb1f    8b8500ffffff            mov eax, dword ptr [ebp+ffffff00]
'0069fb25    894208                  mov dword ptr [edx+08], eax
'0069fb28    8b8d04ffffff            mov ecx, dword ptr [ebp+ffffff04]
'0069fb2e    894a0c                  mov dword ptr [edx+0c], ecx
'0069fb31    83ec10                  sub esp, 10
'0069fb34    8bd4                    mov edx, esp
'0069fb36    8b85d8feffff            mov eax, dword ptr [ebp+fffffed8]
'0069fb3c    8902                    mov dword ptr [edx], eax
'0069fb3e    8b8ddcfeffff            mov ecx, dword ptr [ebp+fffffedc]
'0069fb44    894a04                  mov dword ptr [edx+04], ecx
'0069fb47    8b85e0feffff            mov eax, dword ptr [ebp+fffffee0]
'0069fb4d    894208                  mov dword ptr [edx+08], eax
'0069fb50    8b8de4feffff            mov ecx, dword ptr [ebp+fffffee4]
'0069fb56    894a0c                  mov dword ptr [edx+0c], ecx
'0069fb59    6a03                    push 03
'0069fb5b    689e000000              push 0000009e
'0069fb60    8b16                    mov edx, dword ptr [esi]
'0069fb62    56                      push esi
'0069fb63    ff9214030000            call dword ptr [edx+00000314]
'0069fb69    50                      push eax
'0069fb6a    8d45c0                  lea eax, dword ptr [ebp-40]
'0069fb6d    50                      push eax
'0069fb6e    ffd7                    call edi
    Set var_5 = Nothing
'0069fb70    50                      push eax
'0069fb71    8d4d98                  lea ecx, dword ptr [ebp-68]
'0069fb74    51                      push ecx
'0069fb75    ffd3                    call ebx
    var_130 = var_5.UNK_-256 - 20_158
'0069fb77    83c430                  add esp, 30
'0069fb7a    8bd4                    mov edx, esp
'0069fb7c    8b08                    mov ecx, dword ptr [eax]
'0069fb7e    890a                    mov dword ptr [edx], ecx
'0069fb80    8b4804                  mov ecx, dword ptr [eax+04]
'0069fb83    894a04                  mov dword ptr [edx+04], ecx
'0069fb86    8b4808                  mov ecx, dword ptr [eax+08]
'0069fb89    894a08                  mov dword ptr [edx+08], ecx
'0069fb8c    8b400c                  mov eax, dword ptr [eax+0c]
'0069fb8f    89420c                  mov dword ptr [edx+0c], eax
'0069fb92    6a03                    push 03
'0069fb94    689e000000              push 0000009e
'0069fb99    a184b07200              mov ax, word ptr [0072b084]
'0069fb9e    8b08                    mov ecx, dword ptr [eax]
'0069fba0    50                      push eax
'0069fba1    ff9140040000            call dword ptr [ecx+00000440]
'0069fba7    50                      push eax
'0069fba8    8d55bc                  lea edx, dword ptr [ebp-44]
'0069fbab    52                      push edx
'0069fbac    ffd7                    call edi
    Set var_58 = var_28
'0069fbae    50                      push eax

' *** Reference to "__vbaLateIdCallSt"
'0069fbaf    ff159c114000            call dword ptr [0040119c]
'0069fbb5    8d45bc                  lea eax, dword ptr [ebp-44]
'0069fbb8    50                      push eax
'0069fbb9    8d4dc0                  lea ecx, dword ptr [ebp-40]
'0069fbbc    51                      push ecx
'0069fbbd    8d55c4                  lea edx, dword ptr [ebp-3c]
'0069fbc0    52                      push edx
'0069fbc1    6a03                    push 03

' *** Reference to "__vbaFreeObjList"
'0069fbc3    ff154c104000            call dword ptr [0040104c]
    '#Cleanup( var_9, var_5, var_58)
'0069fbc9    83c45c                  add esp, 5c
'0069fbcc    8d4598                  lea eax, dword ptr [ebp-68]
'0069fbcf    50                      push eax
'0069fbd0    8d4da8                  lea ecx, dword ptr [ebp-58]
'0069fbd3    51                      push ecx
'0069fbd4    6a02                    push 02

' *** Reference to "__vbaFreeVarList"
'0069fbd6    ff1540104000            call dword ptr [00401040]
    '#Cleanup( var_86, var_130)
'0069fbdc    83c40c                  add esp, 0c
'0069fbdf    c78520ffffff00000000    mov dword ptr [ebp+ffffff20], 00000000
'0069fbe9    c78518ffffff03000000    mov dword ptr [ebp+ffffff18], 00000003
'0069fbf3    6a00                    push 00
'0069fbf5    6a11                    push 11
'0069fbf7    8b16                    mov edx, dword ptr [esi]
'0069fbf9    56                      push esi
'0069fbfa    ff9214030000            call dword ptr [edx+00000314]
'0069fc00    50                      push eax
'0069fc01    8d45c4                  lea eax, dword ptr [ebp-3c]
'0069fc04    50                      push eax
'0069fc05    ffd7                    call edi
    Set var_9 = 
'0069fc07    50                      push eax
'0069fc08    8d4da8                  lea ecx, dword ptr [ebp-58]
'0069fc0b    51                      push ecx
'0069fc0c    ffd3                    call ebx
    var_86 = var_9.UNK_var_130_17
'0069fc0e    83c410                  add esp, 10
'0069fc11    50                      push eax

' *** Reference to "__vbaI4Var"
'0069fc12    ff157c124000            call dword ptr [0040127c]
'0069fc18    898500ffffff            mov dword ptr [ebp+ffffff00], eax
'0069fc1e    c785f8feffff03000000    mov dword ptr [ebp+fffffef8], 00000003
'0069fc28    c785e0feffff01000000    mov dword ptr [ebp+fffffee0], 00000001
'0069fc32    c785d8feffff02000000    mov dword ptr [ebp+fffffed8], 00000002
'0069fc3c    a184b07200              mov ax, word ptr [0072b084]
'0069fc41    85c0                    test ax, ax
'0069fc43    7510                    jne 69fc55
'0069fc45    6884b07200              push 0072b084
'0069fc4a    68f8894100              push 004189f8

' *** Reference to "__vbaNew2"
'0069fc4f    ff152c124000            call dword ptr [0040122c]
    Set var_28 = New frmGestPerso
'0069fc55    83ec10                  sub esp, 10
'0069fc58    8bd4                    mov edx, esp
'0069fc5a    b803000000              mov eax, 00000003
'0069fc5f    8902                    mov dword ptr [edx], eax
'0069fc61    8b85bcfeffff            mov eax, dword ptr [ebp+fffffebc]
'0069fc67    894204                  mov dword ptr [edx+04], eax
'0069fc6a    33c0                    xor eax, eax
    var_num1 = Empty
'0069fc6c    894208                  mov dword ptr [edx+08], eax
'0069fc6f    8b8dc4feffff            mov ecx, dword ptr [ebp+fffffec4]
'0069fc75    894a0c                  mov dword ptr [edx+0c], ecx
'0069fc78    83ec10                  sub esp, 10
'0069fc7b    8bd4                    mov edx, esp
'0069fc7d    b80a000000              mov eax, 0000000a
'0069fc82    8902                    mov dword ptr [edx], eax
'0069fc84    8b859cfeffff            mov eax, dword ptr [ebp+fffffe9c]
'0069fc8a    894204                  mov dword ptr [edx+04], eax
'0069fc8d    b804000280              mov eax, 80020004
'0069fc92    894208                  mov dword ptr [edx+08], eax
'0069fc95    8b8da4feffff            mov ecx, dword ptr [ebp+fffffea4]
'0069fc9b    894a0c                  mov dword ptr [edx+0c], ecx
'0069fc9e    83ec10                  sub esp, 10
'0069fca1    8bd4                    mov edx, esp
'0069fca3    b802000000              mov eax, 00000002
'0069fca8    8902                    mov dword ptr [edx], eax
'0069fcaa    8b857cfeffff            mov eax, dword ptr [ebp+fffffe7c]
'0069fcb0    894204                  mov dword ptr [edx+04], eax
'0069fcb3    b801000000              mov eax, 00000001
'0069fcb8    894208                  mov dword ptr [edx+08], eax
'0069fcbb    8b8d84feffff            mov ecx, dword ptr [ebp+fffffe84]
'0069fcc1    894a0c                  mov dword ptr [edx+0c], ecx
'0069fcc4    83ec10                  sub esp, 10
'0069fcc7    8bd4                    mov edx, esp
'0069fcc9    8b8518ffffff            mov eax, dword ptr [ebp+ffffff18]
'0069fccf    8902                    mov dword ptr [edx], eax
'0069fcd1    8b8d1cffffff            mov ecx, dword ptr [ebp+ffffff1c]
'0069fcd7    894a04                  mov dword ptr [edx+04], ecx
'0069fcda    8b8520ffffff            mov eax, dword ptr [ebp+ffffff20]
'0069fce0    894208                  mov dword ptr [edx+08], eax
'0069fce3    8b8d24ffffff            mov ecx, dword ptr [ebp+ffffff24]
'0069fce9    894a0c                  mov dword ptr [edx+0c], ecx
'0069fcec    83ec10                  sub esp, 10
'0069fcef    8bd4                    mov edx, esp
'0069fcf1    8b85f8feffff            mov eax, dword ptr [ebp+fffffef8]
'0069fcf7    8902                    mov dword ptr [edx], eax
'0069fcf9    8b8dfcfeffff            mov ecx, dword ptr [ebp+fffffefc]
'0069fcff    894a04                  mov dword ptr [edx+04], ecx
'0069fd02    8b8500ffffff            mov eax, dword ptr [ebp+ffffff00]
'0069fd08    894208                  mov dword ptr [edx+08], eax
'0069fd0b    8b8d04ffffff            mov ecx, dword ptr [ebp+ffffff04]
'0069fd11    894a0c                  mov dword ptr [edx+0c], ecx
'0069fd14    83ec10                  sub esp, 10
'0069fd17    8bd4                    mov edx, esp
'0069fd19    8b85d8feffff            mov eax, dword ptr [ebp+fffffed8]
'0069fd1f    8902                    mov dword ptr [edx], eax
'0069fd21    8b8ddcfeffff            mov ecx, dword ptr [ebp+fffffedc]
'0069fd27    894a04                  mov dword ptr [edx+04], ecx
'0069fd2a    8b85e0feffff            mov eax, dword ptr [ebp+fffffee0]
'0069fd30    894208                  mov dword ptr [edx+08], eax
'0069fd33    8b8de4feffff            mov ecx, dword ptr [ebp+fffffee4]
'0069fd39    894a0c                  mov dword ptr [edx+0c], ecx
'0069fd3c    6a03                    push 03
'0069fd3e    689e000000              push 0000009e
'0069fd43    8b16                    mov edx, dword ptr [esi]
'0069fd45    56                      push esi
'0069fd46    ff9214030000            call dword ptr [edx+00000314]
'0069fd4c    50                      push eax
'0069fd4d    8d45c0                  lea eax, dword ptr [ebp-40]
'0069fd50    50                      push eax
'0069fd51    ffd7                    call edi
    Set var_5 = var_427
'0069fd53    50                      push eax
'0069fd54    8d4d98                  lea ecx, dword ptr [ebp-68]
'0069fd57    51                      push ecx
'0069fd58    ffd3                    call ebx
    var_130 = var_5.UNK__158
'0069fd5a    83c430                  add esp, 30
'0069fd5d    8bd4                    mov edx, esp
'0069fd5f    8b08                    mov ecx, dword ptr [eax]
'0069fd61    890a                    mov dword ptr [edx], ecx
'0069fd63    8b4804                  mov ecx, dword ptr [eax+04]
'0069fd66    894a04                  mov dword ptr [edx+04], ecx
'0069fd69    8b4808                  mov ecx, dword ptr [eax+08]
'0069fd6c    894a08                  mov dword ptr [edx+08], ecx
'0069fd6f    8b400c                  mov eax, dword ptr [eax+0c]
'0069fd72    89420c                  mov dword ptr [edx+0c], eax
'0069fd75    6a03                    push 03
'0069fd77    689e000000              push 0000009e
'0069fd7c    a184b07200              mov ax, word ptr [0072b084]
'0069fd81    8b08                    mov ecx, dword ptr [eax]
'0069fd83    50                      push eax
'0069fd84    ff9140040000            call dword ptr [ecx+00000440]
'0069fd8a    50                      push eax
'0069fd8b    8d55bc                  lea edx, dword ptr [ebp-44]
'0069fd8e    52                      push edx
'0069fd8f    ffd7                    call edi
    Set var_58 = var_28
'0069fd91    50                      push eax

' *** Reference to "__vbaLateIdCallSt"
'0069fd92    ff159c114000            call dword ptr [0040119c]
'0069fd98    8d45bc                  lea eax, dword ptr [ebp-44]
'0069fd9b    50                      push eax
'0069fd9c    8d4dc0                  lea ecx, dword ptr [ebp-40]
'0069fd9f    51                      push ecx
'0069fda0    8d55c4                  lea edx, dword ptr [ebp-3c]
'0069fda3    52                      push edx
'0069fda4    6a03                    push 03

' *** Reference to "__vbaFreeObjList"
'0069fda6    ff154c104000            call dword ptr [0040104c]
    '#Cleanup( var_9, var_5, var_58)
'0069fdac    83c45c                  add esp, 5c
'0069fdaf    8d4598                  lea eax, dword ptr [ebp-68]
'0069fdb2    50                      push eax
'0069fdb3    8d4da8                  lea ecx, dword ptr [ebp-58]
'0069fdb6    51                      push ecx
'0069fdb7    6a02                    push 02

' *** Reference to "__vbaFreeVarList"
'0069fdb9    ff1540104000            call dword ptr [00401040]
    '#Cleanup( var_86, var_130)
'0069fdbf    83c40c                  add esp, 0c
'0069fdc2    c78520ffffff00000000    mov dword ptr [ebp+ffffff20], 00000000
'0069fdcc    c78518ffffff03000000    mov dword ptr [ebp+ffffff18], 00000003
'0069fdd6    6a00                    push 00
'0069fdd8    6a11                    push 11
'0069fdda    8b16                    mov edx, dword ptr [esi]
'0069fddc    56                      push esi
'0069fddd    ff9214030000            call dword ptr [edx+00000314]
'0069fde3    50                      push eax
'0069fde4    8d45c4                  lea eax, dword ptr [ebp-3c]
'0069fde7    50                      push eax
'0069fde8    ffd7                    call edi
    Set var_9 = 
'0069fdea    50                      push eax
'0069fdeb    8d4da8                  lea ecx, dword ptr [ebp-58]
'0069fdee    51                      push ecx
'0069fdef    ffd3                    call ebx
    var_86 = var_9.UNK_var_130_17
'0069fdf1    83c410                  add esp, 10
'0069fdf4    50                      push eax

' *** Reference to "__vbaI4Var"
'0069fdf5    ff157c124000            call dword ptr [0040127c]
'0069fdfb    898500ffffff            mov dword ptr [ebp+ffffff00], eax
'0069fe01    b803000000              mov eax, 00000003
'0069fe06    8985f8feffff            mov dword ptr [ebp+fffffef8], eax
'0069fe0c    8985e0feffff            mov dword ptr [ebp+fffffee0], eax
'0069fe12    c785d8feffff02000000    mov dword ptr [ebp+fffffed8], 00000002
'0069fe1c    a184b07200              mov ax, word ptr [0072b084]
'0069fe21    85c0                    test ax, ax
'0069fe23    7510                    jne 69fe35
'0069fe25    6884b07200              push 0072b084
'0069fe2a    68f8894100              push 004189f8

' *** Reference to "__vbaNew2"
'0069fe2f    ff152c124000            call dword ptr [0040122c]
    Set var_28 = New frmGestPerso
'0069fe35    83ec10                  sub esp, 10
'0069fe38    8bd4                    mov edx, esp
'0069fe3a    b803000000              mov eax, 00000003
'0069fe3f    8902                    mov dword ptr [edx], eax
'0069fe41    8b85bcfeffff            mov eax, dword ptr [ebp+fffffebc]
'0069fe47    894204                  mov dword ptr [edx+04], eax
'0069fe4a    33c0                    xor eax, eax
    var_num1 = Empty
'0069fe4c    894208                  mov dword ptr [edx+08], eax
'0069fe4f    8b8dc4feffff            mov ecx, dword ptr [ebp+fffffec4]
'0069fe55    894a0c                  mov dword ptr [edx+0c], ecx
'0069fe58    83ec10                  sub esp, 10
'0069fe5b    8bd4                    mov edx, esp
'0069fe5d    b80a000000              mov eax, 0000000a
'0069fe62    8902                    mov dword ptr [edx], eax
'0069fe64    8b859cfeffff            mov eax, dword ptr [ebp+fffffe9c]
'0069fe6a    894204                  mov dword ptr [edx+04], eax
'0069fe6d    b804000280              mov eax, 80020004
'0069fe72    894208                  mov dword ptr [edx+08], eax
'0069fe75    8b8da4feffff            mov ecx, dword ptr [ebp+fffffea4]
'0069fe7b    894a0c                  mov dword ptr [edx+0c], ecx
'0069fe7e    83ec10                  sub esp, 10
'0069fe81    8bd4                    mov edx, esp
'0069fe83    b802000000              mov eax, 00000002
'0069fe88    8902                    mov dword ptr [edx], eax
'0069fe8a    8b857cfeffff            mov eax, dword ptr [ebp+fffffe7c]
'0069fe90    894204                  mov dword ptr [edx+04], eax
'0069fe93    b803000000              mov eax, 00000003
'0069fe98    894208                  mov dword ptr [edx+08], eax
'0069fe9b    8b8d84feffff            mov ecx, dword ptr [ebp+fffffe84]
'0069fea1    894a0c                  mov dword ptr [edx+0c], ecx
'0069fea4    83ec10                  sub esp, 10
'0069fea7    8bd4                    mov edx, esp
'0069fea9    8b8518ffffff            mov eax, dword ptr [ebp+ffffff18]
'0069feaf    8902                    mov dword ptr [edx], eax
'0069feb1    8b8d1cffffff            mov ecx, dword ptr [ebp+ffffff1c]
'0069feb7    894a04                  mov dword ptr [edx+04], ecx
'0069feba    8b8520ffffff            mov eax, dword ptr [ebp+ffffff20]
'0069fec0    894208                  mov dword ptr [edx+08], eax
'0069fec3    8b8d24ffffff            mov ecx, dword ptr [ebp+ffffff24]
'0069fec9    894a0c                  mov dword ptr [edx+0c], ecx
'0069fecc    83ec10                  sub esp, 10
'0069fecf    8bd4                    mov edx, esp
'0069fed1    8b85f8feffff            mov eax, dword ptr [ebp+fffffef8]
'0069fed7    8902                    mov dword ptr [edx], eax
'0069fed9    8b8dfcfeffff            mov ecx, dword ptr [ebp+fffffefc]
'0069fedf    894a04                  mov dword ptr [edx+04], ecx
'0069fee2    8b8500ffffff            mov eax, dword ptr [ebp+ffffff00]
'0069fee8    894208                  mov dword ptr [edx+08], eax
'0069feeb    8b8d04ffffff            mov ecx, dword ptr [ebp+ffffff04]
'0069fef1    894a0c                  mov dword ptr [edx+0c], ecx
'0069fef4    83ec10                  sub esp, 10
'0069fef7    8bd4                    mov edx, esp
'0069fef9    8b85d8feffff            mov eax, dword ptr [ebp+fffffed8]
'0069feff    8902                    mov dword ptr [edx], eax
'0069ff01    8b8ddcfeffff            mov ecx, dword ptr [ebp+fffffedc]
'0069ff07    894a04                  mov dword ptr [edx+04], ecx
'0069ff0a    8b85e0feffff            mov eax, dword ptr [ebp+fffffee0]
'0069ff10    894208                  mov dword ptr [edx+08], eax
'0069ff13    8b8de4feffff            mov ecx, dword ptr [ebp+fffffee4]
'0069ff19    894a0c                  mov dword ptr [edx+0c], ecx
'0069ff1c    6a03                    push 03
'0069ff1e    689e000000              push 0000009e
'0069ff23    8b16                    mov edx, dword ptr [esi]
'0069ff25    56                      push esi
'0069ff26    ff9214030000            call dword ptr [edx+00000314]
'0069ff2c    50                      push eax
'0069ff2d    8d45c0                  lea eax, dword ptr [ebp-40]
'0069ff30    50                      push eax
'0069ff31    ffd7                    call edi
    Set var_5 = 3
'0069ff33    50                      push eax
'0069ff34    8d4d98                  lea ecx, dword ptr [ebp-68]
'0069ff37    51                      push ecx
'0069ff38    ffd3                    call ebx
    var_130 = var_5.UNK__158
'0069ff3a    83c430                  add esp, 30
'0069ff3d    8bd4                    mov edx, esp
'0069ff3f    8b08                    mov ecx, dword ptr [eax]
'0069ff41    890a                    mov dword ptr [edx], ecx
'0069ff43    8b4804                  mov ecx, dword ptr [eax+04]
'0069ff46    894a04                  mov dword ptr [edx+04], ecx
'0069ff49    8b4808                  mov ecx, dword ptr [eax+08]
'0069ff4c    894a08                  mov dword ptr [edx+08], ecx
'0069ff4f    8b400c                  mov eax, dword ptr [eax+0c]
'0069ff52    89420c                  mov dword ptr [edx+0c], eax
'0069ff55    6a03                    push 03
'0069ff57    689e000000              push 0000009e
'0069ff5c    a184b07200              mov ax, word ptr [0072b084]
'0069ff61    8b08                    mov ecx, dword ptr [eax]
'0069ff63    50                      push eax
'0069ff64    ff9140040000            call dword ptr [ecx+00000440]
'0069ff6a    50                      push eax
'0069ff6b    8d55bc                  lea edx, dword ptr [ebp-44]
'0069ff6e    52                      push edx
'0069ff6f    ffd7                    call edi
    Set var_58 = var_28
'0069ff71    50                      push eax

' *** Reference to "__vbaLateIdCallSt"
'0069ff72    ff159c114000            call dword ptr [0040119c]
'0069ff78    8d45bc                  lea eax, dword ptr [ebp-44]
'0069ff7b    50                      push eax
'0069ff7c    8d4dc0                  lea ecx, dword ptr [ebp-40]
'0069ff7f    51                      push ecx
'0069ff80    8d55c4                  lea edx, dword ptr [ebp-3c]
'0069ff83    52                      push edx
'0069ff84    6a03                    push 03

' *** Reference to "__vbaFreeObjList"
'0069ff86    ff154c104000            call dword ptr [0040104c]
    '#Cleanup( var_9, var_5, var_58)
'0069ff8c    83c45c                  add esp, 5c
'0069ff8f    8d4598                  lea eax, dword ptr [ebp-68]
'0069ff92    50                      push eax
'0069ff93    8d4da8                  lea ecx, dword ptr [ebp-58]
'0069ff96    51                      push ecx
'0069ff97    6a02                    push 02

' *** Reference to "__vbaFreeVarList"
'0069ff99    ff1540104000            call dword ptr [00401040]
    '#Cleanup( var_86, var_130)
'0069ff9f    83c40c                  add esp, 0c
'0069ffa2    c78520ffffff00000000    mov dword ptr [ebp+ffffff20], 00000000
'0069ffac    c78518ffffff03000000    mov dword ptr [ebp+ffffff18], 00000003
'0069ffb6    c78500ffffff04000280    mov dword ptr [ebp+ffffff00], 80020004
'0069ffc0    c785f8feffff0a000000    mov dword ptr [ebp+fffffef8], 0000000a
'0069ffca    c785e0feffff05000000    mov dword ptr [ebp+fffffee0], 00000005
'0069ffd4    c785d8feffff02000000    mov dword ptr [ebp+fffffed8], 00000002
'0069ffde    8b16                    mov edx, dword ptr [esi]
'0069ffe0    56                      push esi
'0069ffe1    ff9204030000            call dword ptr [edx+00000304]
'0069ffe7    50                      push eax
'0069ffe8    8d45c4                  lea eax, dword ptr [ebp-3c]
'0069ffeb    50                      push eax
'0069ffec    ffd7                    call edi
    Set var_9 = 
'0069ffee    8bd8                    mov ebx, eax
'0069fff0    8b0b                    mov ecx, dword ptr [ebx]
'0069fff2    8d55e0                  lea edx, dword ptr [ebp-20]
'0069fff5    52                      push edx
'0069fff6    53                      push ebx
'0069fff7    ff91a0000000            call dword ptr [ecx+000000a0]
'0069fffd    dbe2                    fnclex
'0069ffff    85c0                    test ax, ax
'006a0001    7d12                    jge 6a0015
'006a0003    68a0000000              push 000000a0
'006a0008    68d00d4300              push 00430dd0
'006a000d    53                      push ebx
'006a000e    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006a000f    ff1580104000            call dword ptr [00401080]
'006a0015    8b45e0                  mov eax, dword ptr [ebp-20]
'006a0018    50                      push eax

' *** Reference to "rtcR8ValFromBstr"
'006a0019    ff1510134000            call dword ptr [00401310]
'006a001f    dd9dc0feffff            fstp qword ptr [ebp+fffffec0]
    'var_314 = (00)
'006a0025    bb05000000              mov ebx, 00000005
'006a002a    a184b07200              mov ax, word ptr [0072b084]
'006a002f    85c0                    test ax, ax
'006a0031    7515                    jne 6a0048
'006a0033    6884b07200              push 0072b084
'006a0038    68f8894100              push 004189f8

' *** Reference to "__vbaNew2"
'006a003d    ff152c124000            call dword ptr [0040122c]
    Set var_28 = New frmGestPerso
'006a0043    a184b07200              mov ax, word ptr [0072b084]
'006a0048    83ec10                  sub esp, 10
'006a004b    8bcc                    mov ecx, esp
'006a004d    8b9518ffffff            mov edx, dword ptr [ebp+ffffff18]
'006a0053    8911                    mov dword ptr [ecx], edx
'006a0055    8b951cffffff            mov edx, dword ptr [ebp+ffffff1c]
'006a005b    895104                  mov dword ptr [ecx+04], edx
'006a005e    8b9520ffffff            mov edx, dword ptr [ebp+ffffff20]
'006a0064    895108                  mov dword ptr [ecx+08], edx
'006a0067    8b9524ffffff            mov edx, dword ptr [ebp+ffffff24]
'006a006d    89510c                  mov dword ptr [ecx+0c], edx
'006a0070    83ec10                  sub esp, 10
'006a0073    8bcc                    mov ecx, esp
'006a0075    8b95f8feffff            mov edx, dword ptr [ebp+fffffef8]
'006a007b    8911                    mov dword ptr [ecx], edx
'006a007d    8b95fcfeffff            mov edx, dword ptr [ebp+fffffefc]
'006a0083    895104                  mov dword ptr [ecx+04], edx
'006a0086    8b9500ffffff            mov edx, dword ptr [ebp+ffffff00]
'006a008c    895108                  mov dword ptr [ecx+08], edx
'006a008f    8b9504ffffff            mov edx, dword ptr [ebp+ffffff04]
'006a0095    89510c                  mov dword ptr [ecx+0c], edx
'006a0098    83ec10                  sub esp, 10
'006a009b    8bcc                    mov ecx, esp
'006a009d    8b95d8feffff            mov edx, dword ptr [ebp+fffffed8]
'006a00a3    8911                    mov dword ptr [ecx], edx
'006a00a5    8b95dcfeffff            mov edx, dword ptr [ebp+fffffedc]
'006a00ab    895104                  mov dword ptr [ecx+04], edx
'006a00ae    8b95e0feffff            mov edx, dword ptr [ebp+fffffee0]
'006a00b4    895108                  mov dword ptr [ecx+08], edx
'006a00b7    8b95e4feffff            mov edx, dword ptr [ebp+fffffee4]
'006a00bd    89510c                  mov dword ptr [ecx+0c], edx
'006a00c0    83ec10                  sub esp, 10
'006a00c3    8bcc                    mov ecx, esp
'006a00c5    8919                    mov dword ptr [ecx], ebx
'006a00c7    8b95bcfeffff            mov edx, dword ptr [ebp+fffffebc]
'006a00cd    895104                  mov dword ptr [ecx+04], edx
'006a00d0    8b95c0feffff            mov edx, dword ptr [ebp+fffffec0]
'006a00d6    895108                  mov dword ptr [ecx+08], edx
'006a00d9    8b9dc4feffff            mov ebx, dword ptr [ebp+fffffec4]
'006a00df    89590c                  mov dword ptr [ecx+0c], ebx
'006a00e2    6a03                    push 03
'006a00e4    689e000000              push 0000009e
'006a00e9    8b08                    mov ecx, dword ptr [eax]
'006a00eb    50                      push eax
'006a00ec    ff9140040000            call dword ptr [ecx+00000440]
'006a00f2    50                      push eax
'006a00f3    8d55c0                  lea edx, dword ptr [ebp-40]
'006a00f6    52                      push edx
'006a00f7    ffd7                    call edi
    Set var_5 = var_28
'006a00f9    50                      push eax

' *** Reference to "__vbaLateIdCallSt"
'006a00fa    ff159c114000            call dword ptr [0040119c]
'006a0100    83c44c                  add esp, 4c
'006a0103    8d4de0                  lea ecx, dword ptr [ebp-20]

' *** Reference to "__vbaFreeStr"
'006a0106    ff150c134000            call dword ptr [0040130c]
    '#Cleanup(var_3)
'006a010c    8d45c0                  lea eax, dword ptr [ebp-40]
'006a010f    50                      push eax
'006a0110    8d4dc4                  lea ecx, dword ptr [ebp-3c]
'006a0113    51                      push ecx
'006a0114    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'006a0116    ff154c104000            call dword ptr [0040104c]
    '#Cleanup( var_9, var_5)
'006a011c    83c40c                  add esp, 0c
'006a011f    c78520ffffff00000000    mov dword ptr [ebp+ffffff20], 00000000
'006a0129    c78518ffffff03000000    mov dword ptr [ebp+ffffff18], 00000003
'006a0133    c78500ffffff04000280    mov dword ptr [ebp+ffffff00], 80020004
'006a013d    c785f8feffff0a000000    mov dword ptr [ebp+fffffef8], 0000000a
'006a0147    c785e0feffff04000000    mov dword ptr [ebp+fffffee0], 00000004
'006a0151    c785d8feffff02000000    mov dword ptr [ebp+fffffed8], 00000002
'006a015b    a184b07200              mov ax, word ptr [0072b084]
'006a0160    85c0                    test ax, ax
'006a0162    7515                    jne 6a0179
'006a0164    6884b07200              push 0072b084
'006a0169    68f8894100              push 004189f8

' *** Reference to "__vbaNew2"
'006a016e    ff152c124000            call dword ptr [0040122c]
    Set var_28 = New frmGestPerso
'006a0174    a184b07200              mov ax, word ptr [0072b084]
'006a0179    83ec10                  sub esp, 10
'006a017c    8bd4                    mov edx, esp
'006a017e    8b8d18ffffff            mov ecx, dword ptr [ebp+ffffff18]
'006a0184    890a                    mov dword ptr [edx], ecx
'006a0186    8b8d1cffffff            mov ecx, dword ptr [ebp+ffffff1c]
'006a018c    894a04                  mov dword ptr [edx+04], ecx
'006a018f    8b8d20ffffff            mov ecx, dword ptr [ebp+ffffff20]
'006a0195    894a08                  mov dword ptr [edx+08], ecx
'006a0198    8b8d24ffffff            mov ecx, dword ptr [ebp+ffffff24]
'006a019e    894a0c                  mov dword ptr [edx+0c], ecx
'006a01a1    83ec10                  sub esp, 10
'006a01a4    8bd4                    mov edx, esp
'006a01a6    8b8df8feffff            mov ecx, dword ptr [ebp+fffffef8]
'006a01ac    890a                    mov dword ptr [edx], ecx
'006a01ae    8b8dfcfeffff            mov ecx, dword ptr [ebp+fffffefc]
'006a01b4    894a04                  mov dword ptr [edx+04], ecx
'006a01b7    8b8d00ffffff            mov ecx, dword ptr [ebp+ffffff00]
'006a01bd    894a08                  mov dword ptr [edx+08], ecx
'006a01c0    8b8d04ffffff            mov ecx, dword ptr [ebp+ffffff04]
'006a01c6    894a0c                  mov dword ptr [edx+0c], ecx
'006a01c9    83ec10                  sub esp, 10
'006a01cc    8bd4                    mov edx, esp
'006a01ce    8b8dd8feffff            mov ecx, dword ptr [ebp+fffffed8]
'006a01d4    890a                    mov dword ptr [edx], ecx
'006a01d6    8b8ddcfeffff            mov ecx, dword ptr [ebp+fffffedc]
'006a01dc    894a04                  mov dword ptr [edx+04], ecx
'006a01df    8b8de0feffff            mov ecx, dword ptr [ebp+fffffee0]
'006a01e5    894a08                  mov dword ptr [edx+08], ecx
'006a01e8    8b8de4feffff            mov ecx, dword ptr [ebp+fffffee4]
'006a01ee    894a0c                  mov dword ptr [edx+0c], ecx
'006a01f1    83ec10                  sub esp, 10
'006a01f4    8bd4                    mov edx, esp
'006a01f6    b908000000              mov ecx, 00000008
'006a01fb    890a                    mov dword ptr [edx], ecx
'006a01fd    8b8dbcfeffff            mov ecx, dword ptr [ebp+fffffebc]
'006a0203    894a04                  mov dword ptr [edx+04], ecx
'006a0206    b9fce94400              mov ecx, 0044e9fc
'006a020b    894a08                  mov dword ptr [edx+08], ecx
'006a020e    895a0c                  mov dword ptr [edx+0c], ebx
'006a0211    6a03                    push 03
'006a0213    689e000000              push 0000009e
'006a0218    8b10                    mov edx, dword ptr [eax]
'006a021a    50                      push eax
'006a021b    ff9240040000            call dword ptr [edx+00000440]
'006a0221    50                      push eax
'006a0222    8d45c4                  lea eax, dword ptr [ebp-3c]
'006a0225    50                      push eax
'006a0226    ffd7                    call edi
    Set var_9 = var_28
'006a0228    50                      push eax

' *** Reference to "__vbaLateIdCallSt"
'006a0229    ff159c114000            call dword ptr [0040119c]
'006a022f    83c44c                  add esp, 4c
'006a0232    8d4dc4                  lea ecx, dword ptr [ebp-3c]

' *** Reference to "__vbaFreeObj"
'006a0235    ff1508134000            call dword ptr [00401308]
    '#Cleanup(var_9)
'006a023b    8b0e                    mov ecx, dword ptr [esi]
'006a023d    56                      push esi
'006a023e    ff9104030000            call dword ptr [ecx+00000304]
'006a0244    50                      push eax
'006a0245    8d55c0                  lea edx, dword ptr [ebp-40]
'006a0248    52                      push edx
'006a0249    ffd7                    call edi
    Set var_5 = var_9
'006a024b    8bd8                    mov ebx, eax
'006a024d    8b03                    mov eax, dword ptr [ebx]
'006a024f    8d4ddc                  lea ecx, dword ptr [ebp-24]
'006a0252    51                      push ecx
'006a0253    53                      push ebx
'006a0254    ff90a0000000            call dword ptr [eax+000000a0]
'006a025a    dbe2                    fnclex
'006a025c    85c0                    test ax, ax
'006a025e    7d12                    jge 6a0272
'006a0260    68a0000000              push 000000a0
'006a0265    68d00d4300              push 00430dd0
'006a026a    53                      push ebx
'006a026b    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006a026c    ff1580104000            call dword ptr [00401080]
'006a0272    8b55dc                  mov edx, dword ptr [ebp-24]
'006a0275    52                      push edx

' *** Reference to "rtcR8ValFromBstr"
'006a0276    ff1510134000            call dword ptr [00401310]
'006a027c    dd9d4cfeffff            fstp qword ptr [ebp+fffffe4c]
    'var_287 = (00)
'006a0282    8b06                    mov eax, dword ptr [esi]
'006a0284    56                      push esi
'006a0285    ff9000030000            call dword ptr [eax+00000300]
'006a028b    50                      push eax
'006a028c    8d4dbc                  lea ecx, dword ptr [ebp-44]
'006a028f    51                      push ecx
'006a0290    ffd7                    call edi
    Set var_58 = Nothing
'006a0292    8bd8                    mov ebx, eax
'006a0294    8b13                    mov edx, dword ptr [ebx]
'006a0296    8d45d8                  lea eax, dword ptr [ebp-28]
'006a0299    50                      push eax
'006a029a    53                      push ebx
'006a029b    ff92a0000000            call dword ptr [edx+000000a0]
'006a02a1    dbe2                    fnclex
'006a02a3    85c0                    test ax, ax
'006a02a5    7d12                    jge 6a02b9
'006a02a7    68a0000000              push 000000a0
'006a02ac    68d00d4300              push 00430dd0
'006a02b1    53                      push ebx
'006a02b2    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006a02b3    ff1580104000            call dword ptr [00401080]
'006a02b9    8b4dd8                  mov ecx, dword ptr [ebp-28]
'006a02bc    51                      push ecx

' *** Reference to "rtcR8ValFromBstr"
'006a02bd    ff1510134000            call dword ptr [00401310]
'006a02c3    dd9d44feffff            fstp qword ptr [ebp+fffffe44]
    'var_325 = (00)
'006a02c9    a184b07200              mov ax, word ptr [0072b084]
'006a02ce    85c0                    test ax, ax
'006a02d0    7515                    jne 6a02e7
'006a02d2    6884b07200              push 0072b084
'006a02d7    68f8894100              push 004189f8

' *** Reference to "__vbaNew2"
'006a02dc    ff152c124000            call dword ptr [0040122c]
    Set var_28 = New frmGestPerso
'006a02e2    a184b07200              mov ax, word ptr [0072b084]
'006a02e7    8b10                    mov edx, dword ptr [eax]
'006a02e9    50                      push eax
'006a02ea    ff9228030000            call dword ptr [edx+00000328]
'006a02f0    50                      push eax
'006a02f1    8d45b8                  lea eax, dword ptr [ebp-48]
'006a02f4    50                      push eax
'006a02f5    ffd7                    call edi
    Set var_61 = var_28
'006a02f7    898528feffff            mov dword ptr [ebp+fffffe28], eax
'006a02fd    a184b07200              mov ax, word ptr [0072b084]
'006a0302    85c0                    test ax, ax
'006a0304    7515                    jne 6a031b
'006a0306    6884b07200              push 0072b084
'006a030b    68f8894100              push 004189f8

' *** Reference to "__vbaNew2"
'006a0310    ff152c124000            call dword ptr [0040122c]
    Set var_28 = New frmGestPerso
'006a0316    a184b07200              mov ax, word ptr [0072b084]
'006a031b    8b08                    mov ecx, dword ptr [eax]
'006a031d    50                      push eax
'006a031e    ff9128030000            call dword ptr [ecx+00000328]
'006a0324    50                      push eax
'006a0325    8d55c4                  lea edx, dword ptr [ebp-3c]
'006a0328    52                      push edx
'006a0329    ffd7                    call edi
    Set var_9 = var_28
'006a032b    8bd8                    mov ebx, eax
'006a032d    8b03                    mov eax, dword ptr [ebx]
'006a032f    8d4de0                  lea ecx, dword ptr [ebp-20]
'006a0332    51                      push ecx
'006a0333    53                      push ebx
'006a0334    ff90a0000000            call dword ptr [eax+000000a0]
'006a033a    dbe2                    fnclex
'006a033c    85c0                    test ax, ax
'006a033e    7d12                    jge 6a0352
'006a0340    68a0000000              push 000000a0
'006a0345    68d00d4300              push 00430dd0
'006a034a    53                      push ebx
'006a034b    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006a034c    ff1580104000            call dword ptr [00401080]
'006a0352    8b9528feffff            mov edx, dword ptr [ebp+fffffe28]
'006a0358    8b1a                    mov ebx, dword ptr [edx]
'006a035a    8b45e0                  mov eax, dword ptr [ebp-20]
'006a035d    50                      push eax

' *** Reference to "rtcR8ValFromBstr"
'006a035e    ff1510134000            call dword ptr [00401310]
'006a0364    0fbf4e34                movsx ecx, word ptr [esi+34]
'006a0368    898df4fdffff            mov dword ptr [ebp+fffffdf4], ecx
'006a036e    db85f4fdffff            fild dword ptr [ebp+fffffdf4]
'006a0374    dd9decfdffff            fstp qword ptr [ebp+fffffdec]
    'var_293 = (00)
'006a037a    dd85ecfdffff            fld qword ptr [ebp+fffffdec]
'006a0380    dc8d4cfeffff            fmul qword ptr [ebp+fffffe4c]
'006a0386    dc8d44feffff            fmul qword ptr [ebp+fffffe44]
'006a038c    dee9                    fsubp
'006a038e    dfe0                    fnstsw ax
'006a0390    a80d                    test al, 0d
'006a0392    0f8519030000            jne 6a06b1
'006a0398    83ec08                  sub esp, 08
'006a039b    dd1c24                  fstp qword ptr [esp]
    'var_282 = (00)

' *** Reference to "__vbaStrR8"
'006a039e    ff1580114000            call dword ptr [00401180]
'006a03a4    8bd0                    mov edx, eax
'006a03a6    8d4dd4                  lea ecx, dword ptr [ebp-2c]

' *** Reference to "__vbaStrMove"
'006a03a9    ff15d0124000            call dword ptr [004012d0]
'006a03af    50                      push eax
'006a03b0    8bd3                    mov edx, ebx
'006a03b2    8b9d28feffff            mov ebx, dword ptr [ebp+fffffe28]
'006a03b8    53                      push ebx
'006a03b9    ff92a4000000            call dword ptr [edx+000000a4]
'006a03bf    dbe2                    fnclex
'006a03c1    85c0                    test ax, ax
'006a03c3    7d12                    jge 6a03d7
    
    If (    -4928) Then
'006a03c5    68a4000000              push 000000a4
'006a03ca    68d00d4300              push 00430dd0
'006a03cf    53                      push ebx
'006a03d0    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006a03d1    ff1580104000            call dword ptr [00401080]
    
End If
'006a03d7    8d45d4                  lea eax, dword ptr [ebp-2c]
'006a03da    50                      push eax
'006a03db    8d4dd8                  lea ecx, dword ptr [ebp-28]
'006a03de    51                      push ecx
'006a03df    8d55dc                  lea edx, dword ptr [ebp-24]
'006a03e2    52                      push edx
'006a03e3    8d45e0                  lea eax, dword ptr [ebp-20]
'006a03e6    50                      push eax
'006a03e7    6a04                    push 04

' *** Reference to "__vbaFreeStrList"
'006a03e9    ff155c124000            call dword ptr [0040125c]
'#Cleanup( , -4540, -4544, -4928)
'006a03ef    8d4db8                  lea ecx, dword ptr [ebp-48]
'006a03f2    51                      push ecx
'006a03f3    8d55bc                  lea edx, dword ptr [ebp-44]
'006a03f6    52                      push edx
'006a03f7    8d45c0                  lea eax, dword ptr [ebp-40]
'006a03fa    50                      push eax
'006a03fb    8d4dc4                  lea ecx, dword ptr [ebp-3c]
'006a03fe    51                      push ecx
'006a03ff    6a04                    push 04

' *** Reference to "__vbaFreeObjList"
'006a0401    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_9, var_5, var_58, var_61)
'006a0407    83c428                  add esp, 28
'006a040a    6a00                    push 00
'006a040c    6a11                    push 11
'006a040e    8b16                    mov edx, dword ptr [esi]
'006a0410    56                      push esi
'006a0411    ff9214030000            call dword ptr [edx+00000314]
'006a0417    50                      push eax
'006a0418    8d45c4                  lea eax, dword ptr [ebp-3c]
'006a041b    50                      push eax
'006a041c    ffd7                    call edi
Set var_9 = 
'006a041e    50                      push eax
'006a041f    8d4da8                  lea ecx, dword ptr [ebp-58]
'006a0422    51                      push ecx

' *** Reference to "__vbaLateIdCallLd"
'006a0423    ff157c114000            call dword ptr [0040117c]
var_86 = var_9.UNK_var_5_17
'006a0429    83c410                  add esp, 10
'006a042c    50                      push eax

' *** Reference to "__vbaI4Var"
'006a042d    ff157c124000            call dword ptr [0040127c]
'006a0433    a340b07200              mov word ptr [0072b040], ax
'006a0438    8d4dc4                  lea ecx, dword ptr [ebp-3c]

' *** Reference to "__vbaFreeObj"
'006a043b    8b1d08134000            mov ebx, dword ptr [00401308]
'006a0441    ffd3                    call ebx
'#Cleanup(var_9)
'006a0443    8d4da8                  lea ecx, dword ptr [ebp-58]

' *** Reference to "__vbaFreeVar"
'006a0446    ff1528104000            call dword ptr [00401028]
'#Cleanup(var_86)
'006a044c    a124be7200              mov ax, word ptr [0072be24]
'006a0451    85c0                    test ax, ax
'006a0453    7510                    jne 6a0465
'006a0455    6824be7200              push 0072be24
'006a045a    6870004300              push 00430070

' *** Reference to "__vbaNew2"
'006a045f    ff152c124000            call dword ptr [0040122c]
Set var_2 = New Global
'006a0465    8b3d24be7200            mov edi, dword ptr [0072be24]
'006a046b    8b17                    mov edx, dword ptr [edi]
'006a046d    56                      push esi
'006a046e    8d45c4                  lea eax, dword ptr [ebp-3c]
'006a0471    50                      push eax
'006a0472    8995e4fdffff            mov dword ptr [ebp+fffffde4], edx

' *** Reference to "__vbaObjSetAddref"
'006a0478    ff15c8104000            call dword ptr [004010c8]
Set var_9 = Me
'006a047e    50                      push eax
'006a047f    57                      push edi
'006a0480    8b8de4fdffff            mov ecx, dword ptr [ebp+fffffde4]
'006a0486    ff5110                  call dword ptr [ecx+10]
Call var_2.Unload(var_9)
'006a0489    dbe2                    fnclex
'006a048b    85c0                    test ax, ax
'006a048d    7d0f                    jge 6a049e
'006a048f    6a10                    push 10
'006a0491    6860004300              push 00430060
'006a0496    57                      push edi
'006a0497    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006a0498    ff1580104000            call dword ptr [00401080]
'006a049e    8d4dc4                  lea ecx, dword ptr [ebp-3c]
'006a04a1    ffd3                    call ebx
'#Cleanup(var_9)
'006a04a3    e95d010000              jmp 6a0605

'ERROR: Two many next close:
End If
'006a04a8    a124be7200              mov ax, word ptr [0072be24]
'006a04ad    85c0                    test ax, ax
'006a04af    7510                    jne 6a04c1
'006a04b1    6824be7200              push 0072be24
'006a04b6    6870004300              push 00430070

' *** Reference to "__vbaNew2"
'006a04bb    ff152c124000            call dword ptr [0040122c]
Set var_2 = New Global
'006a04c1    8b1d24be7200            mov ebx, dword ptr [0072be24]
'006a04c7    8b13                    mov edx, dword ptr [ebx]
'006a04c9    8d45c4                  lea eax, dword ptr [ebp-3c]
'006a04cc    50                      push eax
'006a04cd    53                      push ebx
'006a04ce    ff5218                  call dword ptr [edx+18]
Set var_9 = var_2.Screen
'006a04d1    dbe2                    fnclex
'006a04d3    85c0                    test ax, ax
'006a04d5    7d0f                    jge 6a04e6
'006a04d7    6a18                    push 18
'006a04d9    6860004300              push 00430060
'006a04de    53                      push ebx
'006a04df    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006a04e0    ff1580104000            call dword ptr [00401080]
'006a04e6    8b45c4                  mov eax, dword ptr [ebp-3c]
'006a04e9    898538feffff            mov dword ptr [ebp+fffffe38], eax
'006a04ef    8b18                    mov ebx, dword ptr [eax]
'006a04f1    33c9                    xor ecx, ecx

' *** Reference to "__vbaI2I4"
'006a04f3    ff1550114000            call dword ptr [00401150]
'006a04f9    50                      push eax
'006a04fa    8bcb                    mov ecx, ebx
'006a04fc    8b9d38feffff            mov ebx, dword ptr [ebp+fffffe38]
'006a0502    53                      push ebx
'006a0503    ff517c                  call dword ptr [ecx+7c]
var_9.MousePointer = CInt(0)
'006a0506    dbe2                    fnclex
'006a0508    85c0                    test ax, ax
'006a050a    7d0f                    jge 6a051b
'006a050c    6a7c                    push 7c
'006a050e    6810be4300              push 0043be10
'006a0513    53                      push ebx
'006a0514    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006a0515    ff1580104000            call dword ptr [00401080]
'006a051b    8d4dc4                  lea ecx, dword ptr [ebp-3c]

' *** Reference to "__vbaFreeObj"
'006a051e    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_9)
'006a0524    8b16                    mov edx, dword ptr [esi]
'006a0526    8d45e0                  lea eax, dword ptr [ebp-20]
'006a0529    50                      push eax
'006a052a    56                      push esi
'006a052b    ff5248                  call dword ptr [edx+48]
'006a052e    dbe2                    fnclex
'006a0530    85c0                    test ax, ax
'006a0532    7d0f                    jge 6a0543
'006a0534    6a48                    push 48
'006a0536    6838f64400              push 0044f638
'006a053b    56                      push esi
'006a053c    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006a053d    ff1580104000            call dword ptr [00401080]
'006a0543    b904000280              mov ecx, 80020004
'006a0548    894d80                  mov dword ptr [ebp-80], ecx
'006a054b    b80a000000              mov eax, 0000000a
'006a0550    898578ffffff            mov dword ptr [ebp+ffffff78], eax
'006a0556    894d90                  mov dword ptr [ebp-70], ecx
'006a0559    894588                  mov dword ptr [ebp-78], eax
'006a055c    8b45e0                  mov eax, dword ptr [ebp-20]
'006a055f    c745e000000000          mov dword ptr [ebp-20], 00000000
'006a0566    8945a0                  mov dword ptr [ebp-60], eax
'006a0569    bb08000000              mov ebx, 00000008
'006a056e    895d98                  mov dword ptr [ebp-68], ebx
'006a0571    c78520ffffff7c244500    mov dword ptr [ebp+ffffff20], 0045247c
'006a057b    899d18ffffff            mov dword ptr [ebp+ffffff18], ebx
'006a0581    8d9518ffffff            lea edx, dword ptr [ebp+ffffff18]
'006a0587    8d4da8                  lea ecx, dword ptr [ebp-58]

' *** Reference to "__vbaVarDup"
'006a058a    ff1598124000            call dword ptr [00401298]
var_86 = ("Le nombre d'objet doit être entrée en tant que valeur numérique")
'006a0590    8d8d78ffffff            lea ecx, dword ptr [ebp+ffffff78]
'006a0596    51                      push ecx
'006a0597    8d5588                  lea edx, dword ptr [ebp-78]
'006a059a    52                      push edx
'006a059b    8d4598                  lea eax, dword ptr [ebp-68]
'006a059e    50                      push eax
'006a059f    6a30                    push 30
'006a05a1    8d4da8                  lea ecx, dword ptr [ebp-58]
'006a05a4    51                      push ecx

' *** Reference to "rtcMsgBox"
'006a05a5    ff15b8104000            call dword ptr [004010b8]
var_1684 = MsgBox(var_86, 48, vbNullString)
'006a05ab    8d9578ffffff            lea edx, dword ptr [ebp+ffffff78]
'006a05b1    52                      push edx
'006a05b2    8d4588                  lea eax, dword ptr [ebp-78]
'006a05b5    50                      push eax
'006a05b6    8d4d98                  lea ecx, dword ptr [ebp-68]
'006a05b9    51                      push ecx
'006a05ba    8d55a8                  lea edx, dword ptr [ebp-58]
'006a05bd    52                      push edx
'006a05be    6a04                    push 04

' *** Reference to "__vbaFreeVarList"
'006a05c0    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_86, var_130, var_131, var_87)
'006a05c6    83c414                  add esp, 14
'006a05c9    8b06                    mov eax, dword ptr [esi]
'006a05cb    56                      push esi
'006a05cc    ff9004030000            call dword ptr [eax+00000304]
'006a05d2    50                      push eax
'006a05d3    8d4dc4                  lea ecx, dword ptr [ebp-3c]
'006a05d6    51                      push ecx
'006a05d7    ffd7                    call edi
Set var_9 = Nothing
'006a05d9    8bf0                    mov esi, eax
'006a05db    8b16                    mov edx, dword ptr [esi]
'006a05dd    56                      push esi
'006a05de    ff9204020000            call dword ptr [edx+00000204]
'006a05e4    dbe2                    fnclex
'006a05e6    85c0                    test ax, ax
'006a05e8    7d12                    jge 6a05fc
'006a05ea    6804020000              push 00000204
'006a05ef    68d00d4300              push 00430dd0
'006a05f4    56                      push esi
'006a05f5    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006a05f6    ff1580104000            call dword ptr [00401080]
'006a05fc    8d4dc4                  lea ecx, dword ptr [ebp-3c]

' *** Reference to "__vbaFreeObj"
'006a05ff    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_9)
'ERROR: Two many next close:
End If

' *** Reference to "__vbaExitProc"
'006a0605    ff15a0104000            call dword ptr [004010a0]
'006a060b    9b                      fwait
'006a060c    6892066a00              push 006a0692
'006a0611    eb7e                    jmp 6a0691
'006a0613    8d45c8                  lea eax, dword ptr [ebp-38]
'006a0616    50                      push eax
'006a0617    8d4dcc                  lea ecx, dword ptr [ebp-34]
'006a061a    51                      push ecx
'006a061b    8d55d0                  lea edx, dword ptr [ebp-30]
'006a061e    52                      push edx
'006a061f    8d45d4                  lea eax, dword ptr [ebp-2c]
'006a0622    50                      push eax
'006a0623    8d4dd8                  lea ecx, dword ptr [ebp-28]
'006a0626    51                      push ecx
'006a0627    8d55dc                  lea edx, dword ptr [ebp-24]
'006a062a    52                      push edx
'006a062b    8d45e0                  lea eax, dword ptr [ebp-20]
'006a062e    50                      push eax
'006a062f    6a07                    push 07

' *** Reference to "__vbaFreeStrList"
'006a0631    ff155c124000            call dword ptr [0040125c]
'#Cleanup( var_3, -4540, -4544, -4928, -4552, -4556, -4560)
'006a0637    8d4db8                  lea ecx, dword ptr [ebp-48]
'006a063a    51                      push ecx
'006a063b    8d55bc                  lea edx, dword ptr [ebp-44]
'006a063e    52                      push edx
'006a063f    8d45c0                  lea eax, dword ptr [ebp-40]
'006a0642    50                      push eax
'006a0643    8d4dc4                  lea ecx, dword ptr [ebp-3c]
'006a0646    51                      push ecx
'006a0647    6a04                    push 04

' *** Reference to "__vbaFreeObjList"
'006a0649    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_9, var_5, var_58, var_61)
'006a064f    8d9528ffffff            lea edx, dword ptr [ebp+ffffff28]
'006a0655    52                      push edx
'006a0656    8d8538ffffff            lea eax, dword ptr [ebp+ffffff38]
'006a065c    50                      push eax
'006a065d    8d8d48ffffff            lea ecx, dword ptr [ebp+ffffff48]
'006a0663    51                      push ecx
'006a0664    8d9558ffffff            lea edx, dword ptr [ebp+ffffff58]
'006a066a    52                      push edx
'006a066b    8d8568ffffff            lea eax, dword ptr [ebp+ffffff68]
'006a0671    50                      push eax
'006a0672    8d8d78ffffff            lea ecx, dword ptr [ebp+ffffff78]
'006a0678    51                      push ecx
'006a0679    8d5588                  lea edx, dword ptr [ebp-78]
'006a067c    52                      push edx
'006a067d    8d4598                  lea eax, dword ptr [ebp-68]
'006a0680    50                      push eax
'006a0681    8d4da8                  lea ecx, dword ptr [ebp-58]
'006a0684    51                      push ecx
'006a0685    6a09                    push 09

' *** Reference to "__vbaFreeVarList"
'006a0687    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_86, var_130, var_131, var_87, var_132, var_133, var_134, var_135, var_136)
'006a068d    83c45c                  add esp, 5c
'006a0690    c3                      ret
'006a0691    c3                      ret
'006a0692    8b4508                  mov eax, dword ptr [ebp+08]
'006a0695    8b10                    mov edx, dword ptr [eax]
'006a0697    50                      push eax
'006a0698    ff5208                  call dword ptr [edx+08]
'006a069b    8b45f4                  mov eax, dword ptr [ebp-0c]
'006a069e    8b4de4                  mov ecx, dword ptr [ebp-1c]
    'Reference to '__except_list'
'006a06a1    64890d00000000          mov dword ptr fs:[00000000], ecx
'006a06a8    5f                      pop edi
'006a06a9    5e                      pop esi
'006a06aa    5b                      pop ebx
'006a06ab    8be5                    mov esp, ebp
'006a06ad    5d                      pop ebp
'006a06ae    c20400                  ret 0004


End Sub


'Event for btnAnnuler
Private Sub btnAnnuler_Click()
'0069ec20    55                      push ebp
'0069ec21    8bec                    mov ebp, esp
'0069ec23    83ec14                  sub esp, 14

' *** Reference to "__vbaExceptHandler"
'0069ec26    6866724000              push 00407266
'0069ec2b    64a100000000            mov ax, word ptr fs:[00000000]
'0069ec31    50                      push eax
    'Reference to '__except_list'
'0069ec32    64892500000000          mov dword ptr fs:[00000000], esp
'0069ec39    81ec30010000            sub esp, 00000130
'0069ec3f    53                      push ebx
'0069ec40    56                      push esi
'0069ec41    57                      push edi
'0069ec42    8965ec                  mov dword ptr [ebp-14], esp
'0069ec45    c745f0a0614000          mov dword ptr [ebp-10], 004061a0
'0069ec4c    8b5d08                  mov ebx, dword ptr [ebp+08]
'0069ec4f    8bc3                    mov eax, ebx
'0069ec51    83e001                  and eax, 01
'0069ec54    8945f4                  mov dword ptr [ebp-0c], eax
'0069ec57    83e3fe                  and ebx, fffffffe
'0069ec5a    895d08                  mov dword ptr [ebp+08], ebx
'0069ec5d    33f6                    xor esi, esi
'0069ec5f    8975f8                  mov dword ptr [ebp-08], esi
'0069ec62    8b0b                    mov ecx, dword ptr [ebx]
'0069ec64    53                      push ebx
'0069ec65    ff5104                  call dword ptr [ecx+04]
'0069ec68    8975e0                  mov dword ptr [ebp-20], esi
'0069ec6b    8975dc                  mov dword ptr [ebp-24], esi
'0069ec6e    8975d8                  mov dword ptr [ebp-28], esi
'0069ec71    8975d4                  mov dword ptr [ebp-2c], esi
'0069ec74    8975d0                  mov dword ptr [ebp-30], esi
'0069ec77    8975cc                  mov dword ptr [ebp-34], esi
'0069ec7a    8975c8                  mov dword ptr [ebp-38], esi
'0069ec7d    8975c4                  mov dword ptr [ebp-3c], esi
'0069ec80    8975c0                  mov dword ptr [ebp-40], esi
'0069ec83    8975b0                  mov dword ptr [ebp-50], esi
'0069ec86    8975a0                  mov dword ptr [ebp-60], esi
'0069ec89    897590                  mov dword ptr [ebp-70], esi
'0069ec8c    897580                  mov dword ptr [ebp-80], esi
'0069ec8f    89b570ffffff            mov dword ptr [ebp+ffffff70], esi
'0069ec95    89b560ffffff            mov dword ptr [ebp+ffffff60], esi
'0069ec9b    89b550ffffff            mov dword ptr [ebp+ffffff50], esi
'0069eca1    89b540ffffff            mov dword ptr [ebp+ffffff40], esi
'0069eca7    89b530ffffff            mov dword ptr [ebp+ffffff30], esi
'0069ecad    89b520ffffff            mov dword ptr [ebp+ffffff20], esi
'0069ecb3    89b510ffffff            mov dword ptr [ebp+ffffff10], esi
'0069ecb9    89b500ffffff            mov dword ptr [ebp+ffffff00], esi
'0069ecbf    89b5f0feffff            mov dword ptr [ebp+fffffef0], esi
'0069ecc5    89b5e0feffff            mov dword ptr [ebp+fffffee0], esi
'0069eccb    89b5dcfeffff            mov dword ptr [ebp+fffffedc], esi
'0069ecd1    66393528b07200          cmp word ptr [0072b028], si
'0069ecd8    7508                    jne 69ece2
'0069ecda    6a01                    push 01

' *** Reference to "__vbaOnError"
'0069ecdc    ff15b0104000            call dword ptr [004010b0]
On Error Goto handler_0
'0069ece2    393524be7200            cmp dword ptr [0072be24], esi
'0069ece8    7510                    jne 69ecfa
'0069ecea    6824be7200              push 0072be24
'0069ecef    6870004300              push 00430070

' *** Reference to "__vbaNew2"
'0069ecf4    ff152c124000            call dword ptr [0040122c]
Dim var_2 As New Global
'0069ecfa    8b3d24be7200            mov edi, dword ptr [0072be24]
'0069ed00    8b17                    mov edx, dword ptr [edi]
'0069ed02    53                      push ebx
'0069ed03    8d45c4                  lea eax, dword ptr [ebp-3c]
'0069ed06    50                      push eax
'0069ed07    8995b4feffff            mov dword ptr [ebp+fffffeb4], edx

' *** Reference to "__vbaObjSetAddref"
'0069ed0d    ff15c8104000            call dword ptr [004010c8]
Set var_9 = Me
'0069ed13    50                      push eax
'0069ed14    57                      push edi
'0069ed15    8b8db4feffff            mov ecx, dword ptr [ebp+fffffeb4]
'0069ed1b    ff5110                  call dword ptr [ecx+10]
Call var_2.Unload(var_9)
'0069ed1e    dbe2                    fnclex
'0069ed20    3bc6                    cmp eax, esi
'0069ed22    7d0f                    jge 69ed33
'0069ed24    6a10                    push 10
'0069ed26    6860004300              push 00430060
'0069ed2b    57                      push edi
'0069ed2c    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0069ed2d    ff1580104000            call dword ptr [00401080]
'0069ed33    8d4dc4                  lea ecx, dword ptr [ebp-3c]

' *** Reference to "__vbaFreeObj"
'0069ed36    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_9)

' *** Reference to "__vbaExitProc"
'0069ed3c    ff15a0104000            call dword ptr [004010a0]
'0069ed42    68f5f06900              push 0069f0f5
'0069ed47    e9a8030000              jmp 69f0f4

' *** Reference to "rtcErrObj"
'0069ed4c    8b1d6c124000            mov ebx, dword ptr [0040126c]
'0069ed52    ffd3                    call ebx
'0069ed54    50                      push eax
'0069ed55    8d55c4                  lea edx, dword ptr [ebp-3c]
'0069ed58    52                      push edx

' *** Reference to "__vbaObjSet"
'0069ed59    8b3db4104000            mov edi, dword ptr [004010b4]
'0069ed5f    ffd7                    call edi
Set var_9 = Err
'0069ed61    8bf0                    mov esi, eax
'0069ed63    8b06                    mov eax, dword ptr [esi]
'0069ed65    8d4de0                  lea ecx, dword ptr [ebp-20]
'0069ed68    51                      push ecx
'0069ed69    56                      push esi
'0069ed6a    ff502c                  call dword ptr [eax+2c]
var_3 = var_9.Description
'0069ed6d    dbe2                    fnclex
'0069ed6f    85c0                    test ax, ax
'0069ed71    7d0f                    jge 69ed82
'0069ed73    6a2c                    push 2c
'0069ed75    685c084300              push 0043085c
'0069ed7a    56                      push esi
'0069ed7b    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0069ed7c    ff1580104000            call dword ptr [00401080]
'0069ed82    ffd3                    call ebx
'0069ed84    50                      push eax
'0069ed85    8d55c0                  lea edx, dword ptr [ebp-40]
'0069ed88    52                      push edx
'0069ed89    ffd7                    call edi
Set var_5 = Err
'0069ed8b    8bf0                    mov esi, eax
'0069ed8d    8b06                    mov eax, dword ptr [esi]
'0069ed8f    8d8ddcfeffff            lea ecx, dword ptr [ebp+fffffedc]
'0069ed95    51                      push ecx
'0069ed96    56                      push esi
'0069ed97    ff501c                  call dword ptr [eax+1c]
var_10 = var_5.Number
'0069ed9a    dbe2                    fnclex
'0069ed9c    85c0                    test ax, ax
'0069ed9e    7d0f                    jge 69edaf
'0069eda0    6a1c                    push 1c
'0069eda2    685c084300              push 0043085c
'0069eda7    56                      push esi
'0069eda8    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0069eda9    ff1580104000            call dword ptr [00401080]
'0069edaf    b904000280              mov ecx, 80020004
'0069edb4    894d98                  mov dword ptr [ebp-68], ecx
'0069edb7    b80a000000              mov eax, 0000000a
'0069edbc    894590                  mov dword ptr [ebp-70], eax
'0069edbf    894da8                  mov dword ptr [ebp-58], ecx
'0069edc2    8945a0                  mov dword ptr [ebp-60], eax
'0069edc5    c78528ffffff24b07200    mov dword ptr [ebp+ffffff28], 0072b024
'0069edcf    c78520ffffff08400000    mov dword ptr [ebp+ffffff20], 00004008
'0069edd9    6814084300              push 00430814
'0069edde    8b55e0                  mov edx, dword ptr [ebp-20]
'0069ede1    52                      push edx

' *** Reference to "__vbaStrCat"
'0069ede2    8b3570104000            mov esi, dword ptr [00401070]
'0069ede8    ffd6                    call esi
var_11 = ("L'erreur suivante s'est produite : ") & (var_3)
'0069edea    8bd0                    mov edx, eax
'0069edec    8d4ddc                  lea ecx, dword ptr [ebp-24]

' *** Reference to "__vbaStrMove"
'0069edef    8b3dd0124000            mov edi, dword ptr [004012d0]
'0069edf5    ffd7                    call edi
'0069edf7    50                      push eax
'0069edf8    6870084300              push 00430870
'0069edfd    ffd6                    call esi
var_12 = (var_11) & (vbCrLf)
'0069edff    8bd0                    mov edx, eax
'0069ee01    8d4dd8                  lea ecx, dword ptr [ebp-28]
'0069ee04    ffd7                    call edi
'0069ee06    50                      push eax
'0069ee07    6870084300              push 00430870
'0069ee0c    ffd6                    call esi
var_13 = (var_12) & (vbCrLf)
'0069ee0e    8bd0                    mov edx, eax
'0069ee10    8d4dd4                  lea ecx, dword ptr [ebp-2c]
'0069ee13    ffd7                    call edi
'0069ee15    50                      push eax
'0069ee16    6880084300              push 00430880
'0069ee1b    ffd6                    call esi
var_14 = (var_13) & (" numéro d'erreur (")
'0069ee1d    8bd0                    mov edx, eax
'0069ee1f    8d4dd0                  lea ecx, dword ptr [ebp-30]
'0069ee22    ffd7                    call edi
'0069ee24    50                      push eax
'0069ee25    8b85dcfeffff            mov eax, dword ptr [ebp+fffffedc]
'0069ee2b    50                      push eax

' *** Reference to "__vbaStrI4"
'0069ee2c    ff1520104000            call dword ptr [00401020]
'0069ee32    8bd0                    mov edx, eax
'0069ee34    8d4dcc                  lea ecx, dword ptr [ebp-34]
'0069ee37    ffd7                    call edi
'0069ee39    50                      push eax
'0069ee3a    ffd6                    call esi
var_15 = (var_14) & (CStr(var_10))
'0069ee3c    8bd0                    mov edx, eax
'0069ee3e    8d4dc8                  lea ecx, dword ptr [ebp-38]
'0069ee41    ffd7                    call edi
'0069ee43    50                      push eax
'0069ee44    68ac084300              push 004308ac
'0069ee49    ffd6                    call esi
var_16 = (var_15) & (" )")
'0069ee4b    8945b8                  mov dword ptr [ebp-48], eax
'0069ee4e    bf08000000              mov edi, 00000008
'0069ee53    897db0                  mov dword ptr [ebp-50], edi
'0069ee56    8d4d90                  lea ecx, dword ptr [ebp-70]
'0069ee59    51                      push ecx
'0069ee5a    8d55a0                  lea edx, dword ptr [ebp-60]
'0069ee5d    52                      push edx
'0069ee5e    8d8520ffffff            lea eax, dword ptr [ebp+ffffff20]
'0069ee64    50                      push eax
'0069ee65    6a10                    push 10
'0069ee67    8d4db0                  lea ecx, dword ptr [ebp-50]
'0069ee6a    51                      push ecx

' *** Reference to "rtcMsgBox"
'0069ee6b    ff15b8104000            call dword ptr [004010b8]
var_17 = MsgBox(var_16, 16, vbNullString)
'0069ee71    8d55c8                  lea edx, dword ptr [ebp-38]
'0069ee74    52                      push edx
'0069ee75    8d45cc                  lea eax, dword ptr [ebp-34]
'0069ee78    50                      push eax
'0069ee79    8d4dd0                  lea ecx, dword ptr [ebp-30]
'0069ee7c    51                      push ecx
'0069ee7d    8d55d4                  lea edx, dword ptr [ebp-2c]
'0069ee80    52                      push edx
'0069ee81    8d45d8                  lea eax, dword ptr [ebp-28]
'0069ee84    50                      push eax
'0069ee85    8d4ddc                  lea ecx, dword ptr [ebp-24]
'0069ee88    51                      push ecx
'0069ee89    8d55e0                  lea edx, dword ptr [ebp-20]
'0069ee8c    52                      push edx
'0069ee8d    6a07                    push 07

' *** Reference to "__vbaFreeStrList"
'0069ee8f    ff155c124000            call dword ptr [0040125c]
'#Cleanup( 0, -4512, -4516, -4520, -4524, -4528, -4532)
'0069ee95    8d45c0                  lea eax, dword ptr [ebp-40]
'0069ee98    50                      push eax
'0069ee99    8d4dc4                  lea ecx, dword ptr [ebp-3c]
'0069ee9c    51                      push ecx
'0069ee9d    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'0069ee9f    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_9, var_5)
'0069eea5    8d5590                  lea edx, dword ptr [ebp-70]
'0069eea8    52                      push edx
'0069eea9    8d45a0                  lea eax, dword ptr [ebp-60]
'0069eeac    50                      push eax
'0069eead    8d4db0                  lea ecx, dword ptr [ebp-50]
'0069eeb0    51                      push ecx
'0069eeb1    6a03                    push 03

' *** Reference to "__vbaFreeVarList"
'0069eeb3    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_6, var_7, var_8)
'0069eeb9    83c43c                  add esp, 3c
'0069eebc    8d55b0                  lea edx, dword ptr [ebp-50]
'0069eebf    52                      push edx

' *** Reference to "rtcGetPresentDate"
'0069eec0    ff15f4124000            call dword ptr [004012f4]
var_16 = Now()
'0069eec6    c78528ffffffb8084300    mov dword ptr [ebp+ffffff28], 004308b8
'0069eed0    89bd20ffffff            mov dword ptr [ebp+ffffff20], edi
'0069eed6    8d9520ffffff            lea edx, dword ptr [ebp+ffffff20]
'0069eedc    8d4da0                  lea ecx, dword ptr [ebp-60]

' *** Reference to "__vbaVarDup"
'0069eedf    ff1598124000            call dword ptr [00401298]
var_7 = ("dd/MM/yyyy hh:mm:ss")
'0069eee5    6a01                    push 01
'0069eee7    6a01                    push 01
'0069eee9    8d45a0                  lea eax, dword ptr [ebp-60]
'0069eeec    50                      push eax
'0069eeed    8d4db0                  lea ecx, dword ptr [ebp-50]
'0069eef0    51                      push ecx
'0069eef1    8d5590                  lea edx, dword ptr [ebp-70]
'0069eef4    52                      push edx

' *** Reference to "rtcVarFromFormatVar"
'0069eef5    ff1574104000            call dword ptr [00401074]
'0069eefb    ffd3                    call ebx
'0069eefd    50                      push eax
'0069eefe    8d45c4                  lea eax, dword ptr [ebp-3c]
'0069ef01    50                      push eax

' *** Reference to "__vbaObjSet"
'0069ef02    ff15b4104000            call dword ptr [004010b4]
Set var_9 = Err
'0069ef08    8bf0                    mov esi, eax
'0069ef0a    8b0e                    mov ecx, dword ptr [esi]
'0069ef0c    8d95dcfeffff            lea edx, dword ptr [ebp+fffffedc]
'0069ef12    52                      push edx
'0069ef13    56                      push esi
'0069ef14    ff511c                  call dword ptr [ecx+1c]
var_10 = var_9.Number
'0069ef17    dbe2                    fnclex
'0069ef19    85c0                    test ax, ax
'0069ef1b    7d0f                    jge 69ef2c
'0069ef1d    6a1c                    push 1c
'0069ef1f    685c084300              push 0043085c
'0069ef24    56                      push esi
'0069ef25    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0069ef26    ff1580104000            call dword ptr [00401080]
'0069ef2c    ffd3                    call ebx
'0069ef2e    50                      push eax
'0069ef2f    8d45c0                  lea eax, dword ptr [ebp-40]
'0069ef32    50                      push eax

' *** Reference to "__vbaObjSet"
'0069ef33    ff15b4104000            call dword ptr [004010b4]
Set var_5 = Err
'0069ef39    8bf0                    mov esi, eax
'0069ef3b    8b0e                    mov ecx, dword ptr [esi]
'0069ef3d    8d55e0                  lea edx, dword ptr [ebp-20]
'0069ef40    52                      push edx
'0069ef41    56                      push esi
'0069ef42    ff512c                  call dword ptr [ecx+2c]
var_3 = var_5.Description
'0069ef45    dbe2                    fnclex
'0069ef47    85c0                    test ax, ax
'0069ef49    7d0f                    jge 69ef5a
'0069ef4b    6a2c                    push 2c
'0069ef4d    685c084300              push 0043085c
'0069ef52    56                      push esi
'0069ef53    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0069ef54    ff1580104000            call dword ptr [00401080]
'0069ef5a    c78518ffffffe4084300    mov dword ptr [ebp+ffffff18], 004308e4
'0069ef64    89bd10ffffff            mov dword ptr [ebp+ffffff10], edi
'0069ef6a    8b85dcfeffff            mov eax, dword ptr [ebp+fffffedc]
'0069ef70    898508ffffff            mov dword ptr [ebp+ffffff08], eax
'0069ef76    c78500ffffff03000000    mov dword ptr [ebp+ffffff00], 00000003
'0069ef80    c785f8feffff08094300    mov dword ptr [ebp+fffffef8], 00430908
'0069ef8a    89bdf0feffff            mov dword ptr [ebp+fffffef0], edi
'0069ef90    8b45e0                  mov eax, dword ptr [ebp-20]
'0069ef93    c745e000000000          mov dword ptr [ebp-20], 00000000
'0069ef9a    898558ffffff            mov dword ptr [ebp+ffffff58], eax
'0069efa0    89bd50ffffff            mov dword ptr [ebp+ffffff50], edi
'0069efa6    c785e8feffffd0234500    mov dword ptr [ebp+fffffee8], 004523d0
'0069efb0    89bde0feffff            mov dword ptr [ebp+fffffee0], edi
'0069efb6    8d4d90                  lea ecx, dword ptr [ebp-70]
'0069efb9    51                      push ecx
'0069efba    8d9510ffffff            lea edx, dword ptr [ebp+ffffff10]
'0069efc0    52                      push edx
'0069efc1    8d4580                  lea eax, dword ptr [ebp-80]
'0069efc4    50                      push eax

' *** Reference to "__vbaVarCat"
'0069efc5    8b3508124000            mov esi, dword ptr [00401208]
'0069efcb    ffd6                    call esi
'0069efcd    50                      push eax
'0069efce    8d8d00ffffff            lea ecx, dword ptr [ebp+ffffff00]
'0069efd4    51                      push ecx
'0069efd5    8d9570ffffff            lea edx, dword ptr [ebp+ffffff70]
'0069efdb    52                      push edx
'0069efdc    ffd6                    call esi
'0069efde    50                      push eax
'0069efdf    8d85f0feffff            lea eax, dword ptr [ebp+fffffef0]
'0069efe5    50                      push eax
'0069efe6    8d8d60ffffff            lea ecx, dword ptr [ebp+ffffff60]
'0069efec    51                      push ecx
'0069efed    ffd6                    call esi
'0069efef    50                      push eax
'0069eff0    8d9550ffffff            lea edx, dword ptr [ebp+ffffff50]
'0069eff6    52                      push edx
'0069eff7    8d8540ffffff            lea eax, dword ptr [ebp+ffffff40]
'0069effd    50                      push eax
'0069effe    ffd6                    call esi
'0069f000    50                      push eax
'0069f001    8d8de0feffff            lea ecx, dword ptr [ebp+fffffee0]
'0069f007    51                      push ecx
'0069f008    8d9530ffffff            lea edx, dword ptr [ebp+ffffff30]
'0069f00e    52                      push edx
'0069f00f    ffd6                    call esi
'0069f011    50                      push eax
'0069f012    33c0                    xor eax, eax
'0069f014    66a12ab07200            mov eax, dword ptr [0072b02a]
'0069f01a    50                      push eax
'0069f01b    6884094300              push 00430984

' *** Reference to "__vbaPrintFile"
'0069f020    ff15b8114000            call dword ptr [004011b8]
Print #0, 
'0069f026    8d4dc0                  lea ecx, dword ptr [ebp-40]
'0069f029    51                      push ecx
'0069f02a    8d55c4                  lea edx, dword ptr [ebp-3c]
'0069f02d    52                      push edx
'0069f02e    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'0069f030    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_9, var_5)
'0069f036    8d8530ffffff            lea eax, dword ptr [ebp+ffffff30]
'0069f03c    50                      push eax
'0069f03d    8d8d40ffffff            lea ecx, dword ptr [ebp+ffffff40]
'0069f043    51                      push ecx
'0069f044    8d9550ffffff            lea edx, dword ptr [ebp+ffffff50]
'0069f04a    52                      push edx
'0069f04b    8d8560ffffff            lea eax, dword ptr [ebp+ffffff60]
'0069f051    50                      push eax
'0069f052    8d8d70ffffff            lea ecx, dword ptr [ebp+ffffff70]
'0069f058    51                      push ecx
'0069f059    8d5580                  lea edx, dword ptr [ebp-80]
'0069f05c    52                      push edx
'0069f05d    8d4590                  lea eax, dword ptr [ebp-70]
'0069f060    50                      push eax
'0069f061    8d4da0                  lea ecx, dword ptr [ebp-60]
'0069f064    51                      push ecx
'0069f065    8d55b0                  lea edx, dword ptr [ebp-50]
'0069f068    52                      push edx
'0069f069    6a09                    push 09

' *** Reference to "__vbaFreeVarList"
'0069f06b    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_6, var_7, var_8, var_18, var_19, var_20, var_21, var_22, var_23)
'0069f071    83c440                  add esp, 40
'0069f074    6a00                    push 00

' *** Reference to "__vbaResume"
'0069f076    ff1568104000            call dword ptr [00401068]
'0069f07c    e9bbfcffff              jmp 69ed3c
Resume handler_69ED3C
'0069f081    8d55c8                  lea edx, dword ptr [ebp-38]
'0069f084    52                      push edx
'0069f085    8d45cc                  lea eax, dword ptr [ebp-34]
'0069f088    50                      push eax
'0069f089    8d4dd0                  lea ecx, dword ptr [ebp-30]
'0069f08c    51                      push ecx
'0069f08d    8d55d4                  lea edx, dword ptr [ebp-2c]
'0069f090    52                      push edx
'0069f091    8d45d8                  lea eax, dword ptr [ebp-28]
'0069f094    50                      push eax
'0069f095    8d4ddc                  lea ecx, dword ptr [ebp-24]
'0069f098    51                      push ecx
'0069f099    8d55e0                  lea edx, dword ptr [ebp-20]
'0069f09c    52                      push edx
'0069f09d    6a07                    push 07

' *** Reference to "__vbaFreeStrList"
'0069f09f    ff155c124000            call dword ptr [0040125c]
'#Cleanup( var_3, -4512, -4516, -4520, -4524, -4528, -4532)
'0069f0a5    8d45c0                  lea eax, dword ptr [ebp-40]
'0069f0a8    50                      push eax
'0069f0a9    8d4dc4                  lea ecx, dword ptr [ebp-3c]
'0069f0ac    51                      push ecx
'0069f0ad    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'0069f0af    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_9, var_5)
'0069f0b5    8d9530ffffff            lea edx, dword ptr [ebp+ffffff30]
'0069f0bb    52                      push edx
'0069f0bc    8d8540ffffff            lea eax, dword ptr [ebp+ffffff40]
'0069f0c2    50                      push eax
'0069f0c3    8d8d50ffffff            lea ecx, dword ptr [ebp+ffffff50]
'0069f0c9    51                      push ecx
'0069f0ca    8d9560ffffff            lea edx, dword ptr [ebp+ffffff60]
'0069f0d0    52                      push edx
'0069f0d1    8d8570ffffff            lea eax, dword ptr [ebp+ffffff70]
'0069f0d7    50                      push eax
'0069f0d8    8d4d80                  lea ecx, dword ptr [ebp-80]
'0069f0db    51                      push ecx
'0069f0dc    8d5590                  lea edx, dword ptr [ebp-70]
'0069f0df    52                      push edx
'0069f0e0    8d45a0                  lea eax, dword ptr [ebp-60]
'0069f0e3    50                      push eax
'0069f0e4    8d4db0                  lea ecx, dword ptr [ebp-50]
'0069f0e7    51                      push ecx
'0069f0e8    6a09                    push 09

' *** Reference to "__vbaFreeVarList"
'0069f0ea    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_6, var_7, var_8, var_18, var_19, var_20, var_21, var_22, var_23)
'0069f0f0    83c454                  add esp, 54
'0069f0f3    c3                      ret
'0069f0f4    c3                      ret
'0069f0f5    8b4508                  mov eax, dword ptr [ebp+08]
'0069f0f8    8b10                    mov edx, dword ptr [eax]
'0069f0fa    50                      push eax
'0069f0fb    ff5208                  call dword ptr [edx+08]
'0069f0fe    8b45f4                  mov eax, dword ptr [ebp-0c]
'0069f101    8b4de4                  mov ecx, dword ptr [ebp-1c]
    'Reference to '__except_list'
'0069f104    64890d00000000          mov dword ptr fs:[00000000], ecx
'0069f10b    5f                      pop edi
'0069f10c    5e                      pop esi
'0069f10d    5b                      pop ebx
'0069f10e    8be5                    mov esp, ebp
'0069f110    5d                      pop ebp
'0069f111    c20400                  ret 0004


End Sub


'Event for Form
Private Sub Form_Load()
'006a14a0    55                      push ebp
'006a14a1    8bec                    mov ebp, esp
'006a14a3    83ec14                  sub esp, 14

' *** Reference to "__vbaExceptHandler"
'006a14a6    6866724000              push 00407266
'006a14ab    64a100000000            mov ax, word ptr fs:[00000000]
'006a14b1    50                      push eax
    'Reference to '__except_list'
'006a14b2    64892500000000          mov dword ptr fs:[00000000], esp
'006a14b9    81ec2c010000            sub esp, 0000012c
'006a14bf    53                      push ebx
'006a14c0    56                      push esi
'006a14c1    57                      push edi
'006a14c2    8965ec                  mov dword ptr [ebp-14], esp
'006a14c5    c745f050624000          mov dword ptr [ebp-10], 00406250
'006a14cc    8b7508                  mov esi, dword ptr [ebp+08]
'006a14cf    8bc6                    mov eax, esi
'006a14d1    83e001                  and eax, 01
'006a14d4    8945f4                  mov dword ptr [ebp-0c], eax
'006a14d7    83e6fe                  and esi, fffffffe
'006a14da    897508                  mov dword ptr [ebp+08], esi
'006a14dd    33ff                    xor edi, edi
'006a14df    897df8                  mov dword ptr [ebp-08], edi
'006a14e2    8b0e                    mov ecx, dword ptr [esi]
'006a14e4    56                      push esi
'006a14e5    ff5104                  call dword ptr [ecx+04]
'006a14e8    897de0                  mov dword ptr [ebp-20], edi
'006a14eb    897ddc                  mov dword ptr [ebp-24], edi
'006a14ee    897dd8                  mov dword ptr [ebp-28], edi
'006a14f1    897dd4                  mov dword ptr [ebp-2c], edi
'006a14f4    897dd0                  mov dword ptr [ebp-30], edi
'006a14f7    897dcc                  mov dword ptr [ebp-34], edi
'006a14fa    897dc8                  mov dword ptr [ebp-38], edi
'006a14fd    897dc4                  mov dword ptr [ebp-3c], edi
'006a1500    897dc0                  mov dword ptr [ebp-40], edi
'006a1503    897db0                  mov dword ptr [ebp-50], edi
'006a1506    897da0                  mov dword ptr [ebp-60], edi
'006a1509    897d90                  mov dword ptr [ebp-70], edi
'006a150c    897d80                  mov dword ptr [ebp-80], edi
'006a150f    89bd70ffffff            mov dword ptr [ebp+ffffff70], edi
'006a1515    89bd60ffffff            mov dword ptr [ebp+ffffff60], edi
'006a151b    89bd50ffffff            mov dword ptr [ebp+ffffff50], edi
'006a1521    89bd40ffffff            mov dword ptr [ebp+ffffff40], edi
'006a1527    89bd30ffffff            mov dword ptr [ebp+ffffff30], edi
'006a152d    89bd20ffffff            mov dword ptr [ebp+ffffff20], edi
'006a1533    89bd10ffffff            mov dword ptr [ebp+ffffff10], edi
'006a1539    89bd00ffffff            mov dword ptr [ebp+ffffff00], edi
'006a153f    89bdf0feffff            mov dword ptr [ebp+fffffef0], edi
'006a1545    89bde0feffff            mov dword ptr [ebp+fffffee0], edi
'006a154b    89bddcfeffff            mov dword ptr [ebp+fffffedc], edi
'006a1551    66393d28b07200          cmp word ptr [0072b028], di
'006a1558    7508                    jne 6a1562
'006a155a    6a01                    push 01

' *** Reference to "__vbaOnError"
'006a155c    ff15b0104000            call dword ptr [004010b0]
On Error Goto handler_0
'006a1562    893d40b07200            mov dword ptr [0072b040], edi
'006a1568    66397e34                cmp word ptr [esi+34], di
'006a156c    8b06                    mov eax, dword ptr [esi]
'006a156e    56                      push esi
'006a156f    0f85a6030000            jne 6a191b

If (arg_6 = 0) Then
'006a1575    ff9000030000            call dword ptr [eax+00000300]
'006a157b    50                      push eax
'006a157c    8d4dc4                  lea ecx, dword ptr [ebp-3c]
'006a157f    51                      push ecx

' *** Reference to "__vbaObjSet"
'006a1580    ff15b4104000            call dword ptr [004010b4]
    Set var_9 = Nothing
'006a1586    8bd8                    mov ebx, eax
'006a1588    8b13                    mov edx, dword ptr [ebx]
'006a158a    682cbf4300              push 0043bf2c
'006a158f    53                      push ebx
'006a1590    ff92a4000000            call dword ptr [edx+000000a4]
'006a1596    dbe2                    fnclex
'006a1598    3bc7                    cmp eax, edi
'006a159a    7d12                    jge 6a15ae
    
    If (    var_9 < 0) Then
'006a159c    68a4000000              push 000000a4
'006a15a1    68d00d4300              push 00430dd0
'006a15a6    53                      push ebx
'006a15a7    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006a15a8    ff1580104000            call dword ptr [00401080]
    
End If
'006a15ae    8d4dc4                  lea ecx, dword ptr [ebp-3c]

' *** Reference to "__vbaFreeObj"
'006a15b1    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_9)
'006a15b7    8b06                    mov eax, dword ptr [esi]
'006a15b9    56                      push esi
'006a15ba    ff9000030000            call dword ptr [eax+00000300]
'006a15c0    50                      push eax
'006a15c1    8d4dc4                  lea ecx, dword ptr [ebp-3c]
'006a15c4    51                      push ecx

' *** Reference to "__vbaObjSet"
'006a15c5    ff15b4104000            call dword ptr [004010b4]
Set var_9 = Nothing
'006a15cb    8bd8                    mov ebx, eax
'006a15cd    8b13                    mov edx, dword ptr [ebx]
'006a15cf    57                      push edi
'006a15d0    53                      push ebx
'006a15d1    ff928c000000            call dword ptr [edx+0000008c]
'006a15d7    dbe2                    fnclex
'006a15d9    3bc7                    cmp eax, edi
'006a15db    0f8d70030000            jge 6a1951

If (var_9 < 0) Then
'006a15e1    e959030000              jmp 6a193f

' *** Reference to "rtcErrObj"
'006a15e6    8b1d6c124000            mov ebx, dword ptr [0040126c]
'006a15ec    ffd3                    call ebx
'006a15ee    50                      push eax
'006a15ef    8d55c4                  lea edx, dword ptr [ebp-3c]
'006a15f2    52                      push edx

' *** Reference to "__vbaObjSet"
'006a15f3    8b3db4104000            mov edi, dword ptr [004010b4]
'006a15f9    ffd7                    call edi
    Set var_9 = Err
'006a15fb    8bf0                    mov esi, eax
'006a15fd    8b06                    mov eax, dword ptr [esi]
'006a15ff    8d4de0                  lea ecx, dword ptr [ebp-20]
'006a1602    51                      push ecx
'006a1603    56                      push esi
'006a1604    ff502c                  call dword ptr [eax+2c]
    var_3 = var_9.Description
'006a1607    dbe2                    fnclex
'006a1609    85c0                    test ax, ax
'006a160b    7d0f                    jge 6a161c
'006a160d    6a2c                    push 2c
'006a160f    685c084300              push 0043085c
'006a1614    56                      push esi
'006a1615    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006a1616    ff1580104000            call dword ptr [00401080]
'006a161c    ffd3                    call ebx
'006a161e    50                      push eax
'006a161f    8d55c0                  lea edx, dword ptr [ebp-40]
'006a1622    52                      push edx
'006a1623    ffd7                    call edi
    Set var_5 = Err
'006a1625    8bf0                    mov esi, eax
'006a1627    8b06                    mov eax, dword ptr [esi]
'006a1629    8d8ddcfeffff            lea ecx, dword ptr [ebp+fffffedc]
'006a162f    51                      push ecx
'006a1630    56                      push esi
'006a1631    ff501c                  call dword ptr [eax+1c]
    var_10 = var_5.Number
'006a1634    dbe2                    fnclex
'006a1636    85c0                    test ax, ax
'006a1638    7d0f                    jge 6a1649
'006a163a    6a1c                    push 1c
'006a163c    685c084300              push 0043085c
'006a1641    56                      push esi
'006a1642    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006a1643    ff1580104000            call dword ptr [00401080]
'006a1649    b904000280              mov ecx, 80020004
'006a164e    894d98                  mov dword ptr [ebp-68], ecx
'006a1651    b80a000000              mov eax, 0000000a
'006a1656    894590                  mov dword ptr [ebp-70], eax
'006a1659    894da8                  mov dword ptr [ebp-58], ecx
'006a165c    8945a0                  mov dword ptr [ebp-60], eax
'006a165f    c78528ffffff24b07200    mov dword ptr [ebp+ffffff28], 0072b024
'006a1669    c78520ffffff08400000    mov dword ptr [ebp+ffffff20], 00004008
'006a1673    6814084300              push 00430814
'006a1678    8b55e0                  mov edx, dword ptr [ebp-20]
'006a167b    52                      push edx

' *** Reference to "__vbaStrCat"
'006a167c    8b3570104000            mov esi, dword ptr [00401070]
'006a1682    ffd6                    call esi
    var_11 = ("L'erreur suivante s'est produite : ") & (var_3)
'006a1684    8bd0                    mov edx, eax
'006a1686    8d4ddc                  lea ecx, dword ptr [ebp-24]

' *** Reference to "__vbaStrMove"
'006a1689    8b3dd0124000            mov edi, dword ptr [004012d0]
'006a168f    ffd7                    call edi
'006a1691    50                      push eax
'006a1692    6870084300              push 00430870
'006a1697    ffd6                    call esi
    var_76 = (var_11) & (vbCrLf)
'006a1699    8bd0                    mov edx, eax
'006a169b    8d4dd8                  lea ecx, dword ptr [ebp-28]
'006a169e    ffd7                    call edi
'006a16a0    50                      push eax
'006a16a1    6870084300              push 00430870
'006a16a6    ffd6                    call esi
    var_12 = (var_76) & (vbCrLf)
'006a16a8    8bd0                    mov edx, eax
'006a16aa    8d4dd4                  lea ecx, dword ptr [ebp-2c]
'006a16ad    ffd7                    call edi
'006a16af    50                      push eax
'006a16b0    6880084300              push 00430880
'006a16b5    ffd6                    call esi
    var_13 = (var_12) & (" numéro d'erreur (")
'006a16b7    8bd0                    mov edx, eax
'006a16b9    8d4dd0                  lea ecx, dword ptr [ebp-30]
'006a16bc    ffd7                    call edi
'006a16be    50                      push eax
'006a16bf    8b85dcfeffff            mov eax, dword ptr [ebp+fffffedc]
'006a16c5    50                      push eax

' *** Reference to "__vbaStrI4"
'006a16c6    ff1520104000            call dword ptr [00401020]
'006a16cc    8bd0                    mov edx, eax
'006a16ce    8d4dcc                  lea ecx, dword ptr [ebp-34]
'006a16d1    ffd7                    call edi
'006a16d3    50                      push eax
'006a16d4    ffd6                    call esi
    var_127 = (var_13) & (CStr(var_10))
'006a16d6    8bd0                    mov edx, eax
'006a16d8    8d4dc8                  lea ecx, dword ptr [ebp-38]
'006a16db    ffd7                    call edi
'006a16dd    50                      push eax
'006a16de    68ac084300              push 004308ac
'006a16e3    ffd6                    call esi
    var_15 = (var_127) & (" )")
'006a16e5    8945b8                  mov dword ptr [ebp-48], eax
'006a16e8    bf08000000              mov edi, 00000008
'006a16ed    897db0                  mov dword ptr [ebp-50], edi
'006a16f0    8d4d90                  lea ecx, dword ptr [ebp-70]
'006a16f3    51                      push ecx
'006a16f4    8d55a0                  lea edx, dword ptr [ebp-60]
'006a16f7    52                      push edx
'006a16f8    8d8520ffffff            lea eax, dword ptr [ebp+ffffff20]
'006a16fe    50                      push eax
'006a16ff    6a10                    push 10
'006a1701    8d4db0                  lea ecx, dword ptr [ebp-50]
'006a1704    51                      push ecx

' *** Reference to "rtcMsgBox"
'006a1705    ff15b8104000            call dword ptr [004010b8]
    var_128 = MsgBox(var_15, 16, vbNullString)
'006a170b    8d55c8                  lea edx, dword ptr [ebp-38]
'006a170e    52                      push edx
'006a170f    8d45cc                  lea eax, dword ptr [ebp-34]
'006a1712    50                      push eax
'006a1713    8d4dd0                  lea ecx, dword ptr [ebp-30]
'006a1716    51                      push ecx
'006a1717    8d55d4                  lea edx, dword ptr [ebp-2c]
'006a171a    52                      push edx
'006a171b    8d45d8                  lea eax, dword ptr [ebp-28]
'006a171e    50                      push eax
'006a171f    8d4ddc                  lea ecx, dword ptr [ebp-24]
'006a1722    51                      push ecx
'006a1723    8d55e0                  lea edx, dword ptr [ebp-20]
'006a1726    52                      push edx
'006a1727    6a07                    push 07

' *** Reference to "__vbaFreeStrList"
'006a1729    ff155c124000            call dword ptr [0040125c]
    '#Cleanup( 0, -4508, -4512, -4516, -4520, -4524, -4528)
'006a172f    8d45c0                  lea eax, dword ptr [ebp-40]
'006a1732    50                      push eax
'006a1733    8d4dc4                  lea ecx, dword ptr [ebp-3c]
'006a1736    51                      push ecx
'006a1737    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'006a1739    ff154c104000            call dword ptr [0040104c]
    '#Cleanup( var_9, var_5)
'006a173f    8d5590                  lea edx, dword ptr [ebp-70]
'006a1742    52                      push edx
'006a1743    8d45a0                  lea eax, dword ptr [ebp-60]
'006a1746    50                      push eax
'006a1747    8d4db0                  lea ecx, dword ptr [ebp-50]
'006a174a    51                      push ecx
'006a174b    6a03                    push 03

' *** Reference to "__vbaFreeVarList"
'006a174d    ff1540104000            call dword ptr [00401040]
    '#Cleanup( var_6, var_7, var_8)
'006a1753    83c43c                  add esp, 3c
'006a1756    8d55b0                  lea edx, dword ptr [ebp-50]
'006a1759    52                      push edx

' *** Reference to "rtcGetPresentDate"
'006a175a    ff15f4124000            call dword ptr [004012f4]
    var_15 = Now()
'006a1760    c78528ffffffb8084300    mov dword ptr [ebp+ffffff28], 004308b8
'006a176a    89bd20ffffff            mov dword ptr [ebp+ffffff20], edi
'006a1770    8d9520ffffff            lea edx, dword ptr [ebp+ffffff20]
'006a1776    8d4da0                  lea ecx, dword ptr [ebp-60]

' *** Reference to "__vbaVarDup"
'006a1779    ff1598124000            call dword ptr [00401298]
    var_7 = ("dd/MM/yyyy hh:mm:ss")
'006a177f    6a01                    push 01
'006a1781    6a01                    push 01
'006a1783    8d45a0                  lea eax, dword ptr [ebp-60]
'006a1786    50                      push eax
'006a1787    8d4db0                  lea ecx, dword ptr [ebp-50]
'006a178a    51                      push ecx
'006a178b    8d5590                  lea edx, dword ptr [ebp-70]
'006a178e    52                      push edx

' *** Reference to "rtcVarFromFormatVar"
'006a178f    ff1574104000            call dword ptr [00401074]
'006a1795    ffd3                    call ebx
'006a1797    50                      push eax
'006a1798    8d45c4                  lea eax, dword ptr [ebp-3c]
'006a179b    50                      push eax

' *** Reference to "__vbaObjSet"
'006a179c    ff15b4104000            call dword ptr [004010b4]
    Set var_9 = Err
'006a17a2    8bf0                    mov esi, eax
'006a17a4    8b0e                    mov ecx, dword ptr [esi]
'006a17a6    8d95dcfeffff            lea edx, dword ptr [ebp+fffffedc]
'006a17ac    52                      push edx
'006a17ad    56                      push esi
'006a17ae    ff511c                  call dword ptr [ecx+1c]
    var_10 = var_9.Number
'006a17b1    dbe2                    fnclex
'006a17b3    85c0                    test ax, ax
'006a17b5    7d0f                    jge 6a17c6
'006a17b7    6a1c                    push 1c
'006a17b9    685c084300              push 0043085c
'006a17be    56                      push esi
'006a17bf    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006a17c0    ff1580104000            call dword ptr [00401080]
'006a17c6    ffd3                    call ebx
'006a17c8    50                      push eax
'006a17c9    8d45c0                  lea eax, dword ptr [ebp-40]
'006a17cc    50                      push eax

' *** Reference to "__vbaObjSet"
'006a17cd    ff15b4104000            call dword ptr [004010b4]
    Set var_5 = Err
'006a17d3    8bf0                    mov esi, eax
'006a17d5    8b0e                    mov ecx, dword ptr [esi]
'006a17d7    8d55e0                  lea edx, dword ptr [ebp-20]
'006a17da    52                      push edx
'006a17db    56                      push esi
'006a17dc    ff512c                  call dword ptr [ecx+2c]
    var_3 = var_5.Description
'006a17df    dbe2                    fnclex
'006a17e1    85c0                    test ax, ax
'006a17e3    7d0f                    jge 6a17f4
'006a17e5    6a2c                    push 2c
'006a17e7    685c084300              push 0043085c
'006a17ec    56                      push esi
'006a17ed    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006a17ee    ff1580104000            call dword ptr [00401080]
'006a17f4    c78518ffffffe4084300    mov dword ptr [ebp+ffffff18], 004308e4
'006a17fe    89bd10ffffff            mov dword ptr [ebp+ffffff10], edi
'006a1804    8b85dcfeffff            mov eax, dword ptr [ebp+fffffedc]
'006a180a    898508ffffff            mov dword ptr [ebp+ffffff08], eax
'006a1810    c78500ffffff03000000    mov dword ptr [ebp+ffffff00], 00000003
'006a181a    c785f8feffff08094300    mov dword ptr [ebp+fffffef8], 00430908
'006a1824    89bdf0feffff            mov dword ptr [ebp+fffffef0], edi
'006a182a    8b45e0                  mov eax, dword ptr [ebp-20]
'006a182d    c745e000000000          mov dword ptr [ebp-20], 00000000
'006a1834    898558ffffff            mov dword ptr [ebp+ffffff58], eax
'006a183a    89bd50ffffff            mov dword ptr [ebp+ffffff50], edi
'006a1840    c785e8feffff74274500    mov dword ptr [ebp+fffffee8], 00452774
'006a184a    89bde0feffff            mov dword ptr [ebp+fffffee0], edi
'006a1850    8d4d90                  lea ecx, dword ptr [ebp-70]
'006a1853    51                      push ecx
'006a1854    8d9510ffffff            lea edx, dword ptr [ebp+ffffff10]
'006a185a    52                      push edx
'006a185b    8d4580                  lea eax, dword ptr [ebp-80]
'006a185e    50                      push eax

' *** Reference to "__vbaVarCat"
'006a185f    8b3508124000            mov esi, dword ptr [00401208]
'006a1865    ffd6                    call esi
'006a1867    50                      push eax
'006a1868    8d8d00ffffff            lea ecx, dword ptr [ebp+ffffff00]
'006a186e    51                      push ecx
'006a186f    8d9570ffffff            lea edx, dword ptr [ebp+ffffff70]
'006a1875    52                      push edx
'006a1876    ffd6                    call esi
'006a1878    50                      push eax
'006a1879    8d85f0feffff            lea eax, dword ptr [ebp+fffffef0]
'006a187f    50                      push eax
'006a1880    8d8d60ffffff            lea ecx, dword ptr [ebp+ffffff60]
'006a1886    51                      push ecx
'006a1887    ffd6                    call esi
'006a1889    50                      push eax
'006a188a    8d9550ffffff            lea edx, dword ptr [ebp+ffffff50]
'006a1890    52                      push edx
'006a1891    8d8540ffffff            lea eax, dword ptr [ebp+ffffff40]
'006a1897    50                      push eax
'006a1898    ffd6                    call esi
'006a189a    50                      push eax
'006a189b    8d8de0feffff            lea ecx, dword ptr [ebp+fffffee0]
'006a18a1    51                      push ecx
'006a18a2    8d9530ffffff            lea edx, dword ptr [ebp+ffffff30]
'006a18a8    52                      push edx
'006a18a9    ffd6                    call esi
'006a18ab    50                      push eax
'006a18ac    33c0                    xor eax, eax
'006a18ae    66a12ab07200            mov eax, dword ptr [0072b02a]
'006a18b4    50                      push eax
'006a18b5    6884094300              push 00430984

' *** Reference to "__vbaPrintFile"
'006a18ba    ff15b8114000            call dword ptr [004011b8]
    Print #0, 
'006a18c0    8d4dc0                  lea ecx, dword ptr [ebp-40]
'006a18c3    51                      push ecx
'006a18c4    8d55c4                  lea edx, dword ptr [ebp-3c]
'006a18c7    52                      push edx
'006a18c8    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'006a18ca    ff154c104000            call dword ptr [0040104c]
    '#Cleanup( var_9, var_5)
'006a18d0    8d8530ffffff            lea eax, dword ptr [ebp+ffffff30]
'006a18d6    50                      push eax
'006a18d7    8d8d40ffffff            lea ecx, dword ptr [ebp+ffffff40]
'006a18dd    51                      push ecx
'006a18de    8d9550ffffff            lea edx, dword ptr [ebp+ffffff50]
'006a18e4    52                      push edx
'006a18e5    8d8560ffffff            lea eax, dword ptr [ebp+ffffff60]
'006a18eb    50                      push eax
'006a18ec    8d8d70ffffff            lea ecx, dword ptr [ebp+ffffff70]
'006a18f2    51                      push ecx
'006a18f3    8d5580                  lea edx, dword ptr [ebp-80]
'006a18f6    52                      push edx
'006a18f7    8d4590                  lea eax, dword ptr [ebp-70]
'006a18fa    50                      push eax
'006a18fb    8d4da0                  lea ecx, dword ptr [ebp-60]
'006a18fe    51                      push ecx
'006a18ff    8d55b0                  lea edx, dword ptr [ebp-50]
'006a1902    52                      push edx
'006a1903    6a09                    push 09

' *** Reference to "__vbaFreeVarList"
'006a1905    ff1540104000            call dword ptr [00401040]
    '#Cleanup( var_6, var_7, var_8, var_18, var_19, var_20, var_21, var_22, var_23)
'006a190b    83c440                  add esp, 40
'006a190e    6a00                    push 00

' *** Reference to "__vbaResume"
'006a1910    ff1568104000            call dword ptr [00401068]
'006a1916    e9ee000000              jmp 6a1a09
    Resume handler_6A1A09
End If
'006a191b    ff9000030000            call dword ptr [eax+00000300]
'006a1921    50                      push eax
'006a1922    8d4dc4                  lea ecx, dword ptr [ebp-3c]
'006a1925    51                      push ecx

' *** Reference to "__vbaObjSet"
'006a1926    ff15b4104000            call dword ptr [004010b4]
Set var_9 = Nothing
'006a192c    8bd8                    mov ebx, eax
'006a192e    8b13                    mov edx, dword ptr [ebx]
'006a1930    6aff                    push ffffffff
'006a1932    53                      push ebx
'006a1933    ff928c000000            call dword ptr [edx+0000008c]
'006a1939    dbe2                    fnclex
'006a193b    3bc7                    cmp eax, edi
'006a193d    7d12                    jge 6a1951

If (var_9 < 0) Then
'006a193f    688c000000              push 0000008c
'006a1944    68d00d4300              push 00430dd0
'006a1949    53                      push ebx
'006a194a    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006a194b    ff1580104000            call dword ptr [00401080]
    
End If
'006a1951    8d4dc4                  lea ecx, dword ptr [ebp-3c]

' *** Reference to "__vbaFreeObj"
'006a1954    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_9)
'006a195a    b804000000              mov eax, 00000004
'006a195f    898528ffffff            mov dword ptr [ebp+ffffff28], eax
'006a1965    b903000000              mov ecx, 00000003
'006a196a    898d20ffffff            mov dword ptr [ebp+ffffff20], ecx
'006a1970    83caff                  or edx, ffffffff
var_num4 = -256 - 20 Or -1
'006a1973    899508ffffff            mov dword ptr [ebp+ffffff08], edx
'006a1979    bf0b000000              mov edi, 0000000b
'006a197e    89bd00ffffff            mov dword ptr [ebp+ffffff00], edi
'006a1984    83ec10                  sub esp, 10
'006a1987    8bdc                    mov ebx, esp
'006a1989    890b                    mov dword ptr [ebx], ecx
'006a198b    8b8d24ffffff            mov ecx, dword ptr [ebp+ffffff24]
'006a1991    894b04                  mov dword ptr [ebx+04], ecx
'006a1994    894308                  mov dword ptr [ebx+08], eax
'006a1997    8b852cffffff            mov eax, dword ptr [ebp+ffffff2c]
'006a199d    89430c                  mov dword ptr [ebx+0c], eax
'006a19a0    83ec10                  sub esp, 10
'006a19a3    8bcc                    mov ecx, esp
'006a19a5    8939                    mov dword ptr [ecx], edi
'006a19a7    8b8504ffffff            mov eax, dword ptr [ebp+ffffff04]
'006a19ad    894104                  mov dword ptr [ecx+04], eax
'006a19b0    895108                  mov dword ptr [ecx+08], edx
'006a19b3    8b950cffffff            mov edx, dword ptr [ebp+ffffff0c]
'006a19b9    89510c                  mov dword ptr [ecx+0c], edx
'006a19bc    6a01                    push 01
'006a19be    68ac000000              push 000000ac
'006a19c3    8b06                    mov eax, dword ptr [esi]
'006a19c5    56                      push esi
'006a19c6    ff9014030000            call dword ptr [eax+00000314]
'006a19cc    50                      push eax
'006a19cd    8d4dc4                  lea ecx, dword ptr [ebp-3c]
'006a19d0    51                      push ecx

' *** Reference to "__vbaObjSet"
'006a19d1    ff15b4104000            call dword ptr [004010b4]
Set var_9 = Nothing
'006a19d7    50                      push eax

' *** Reference to "__vbaLateIdCallSt"
'006a19d8    ff159c114000            call dword ptr [0040119c]
'006a19de    83c42c                  add esp, 2c
'006a19e1    8d4dc4                  lea ecx, dword ptr [ebp-3c]

' *** Reference to "__vbaFreeObj"
'006a19e4    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_9)
'006a19ea    8b16                    mov edx, dword ptr [esi]
'006a19ec    56                      push esi
'006a19ed    ff9200070000            call dword ptr [edx+00000700]
'006a19f3    85c0                    test ax, ax
'006a19f5    7d12                    jge 6a1a09
'006a19f7    6800070000              push 00000700
'006a19fc    6870f44400              push 0044f470
'006a1a01    56                      push esi
'006a1a02    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006a1a03    ff1580104000            call dword ptr [00401080]

' *** Reference to "__vbaExitProc"
'006a1a09    ff15a0104000            call dword ptr [004010a0]
'006a1a0f    688a1a6a00              push 006a1a8a
'006a1a14    eb73                    jmp 6a1a89
'006a1a16    8d45c8                  lea eax, dword ptr [ebp-38]
'006a1a19    50                      push eax
'006a1a1a    8d4dcc                  lea ecx, dword ptr [ebp-34]
'006a1a1d    51                      push ecx
'006a1a1e    8d55d0                  lea edx, dword ptr [ebp-30]
'006a1a21    52                      push edx
'006a1a22    8d45d4                  lea eax, dword ptr [ebp-2c]
'006a1a25    50                      push eax
'006a1a26    8d4dd8                  lea ecx, dword ptr [ebp-28]
'006a1a29    51                      push ecx
'006a1a2a    8d55dc                  lea edx, dword ptr [ebp-24]
'006a1a2d    52                      push edx
'006a1a2e    8d45e0                  lea eax, dword ptr [ebp-20]
'006a1a31    50                      push eax
'006a1a32    6a07                    push 07

' *** Reference to "__vbaFreeStrList"
'006a1a34    ff155c124000            call dword ptr [0040125c]
'#Cleanup( var_3, -4508, -4512, -4516, -4520, -4524, -4528)
'006a1a3a    8d4dc0                  lea ecx, dword ptr [ebp-40]
'006a1a3d    51                      push ecx
'006a1a3e    8d55c4                  lea edx, dword ptr [ebp-3c]
'006a1a41    52                      push edx
'006a1a42    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'006a1a44    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_9, var_5)
'006a1a4a    8d8530ffffff            lea eax, dword ptr [ebp+ffffff30]
'006a1a50    50                      push eax
'006a1a51    8d8d40ffffff            lea ecx, dword ptr [ebp+ffffff40]
'006a1a57    51                      push ecx
'006a1a58    8d9550ffffff            lea edx, dword ptr [ebp+ffffff50]
'006a1a5e    52                      push edx
'006a1a5f    8d8560ffffff            lea eax, dword ptr [ebp+ffffff60]
'006a1a65    50                      push eax
'006a1a66    8d8d70ffffff            lea ecx, dword ptr [ebp+ffffff70]
'006a1a6c    51                      push ecx
'006a1a6d    8d5580                  lea edx, dword ptr [ebp-80]
'006a1a70    52                      push edx
'006a1a71    8d4590                  lea eax, dword ptr [ebp-70]
'006a1a74    50                      push eax
'006a1a75    8d4da0                  lea ecx, dword ptr [ebp-60]
'006a1a78    51                      push ecx
'006a1a79    8d55b0                  lea edx, dword ptr [ebp-50]
'006a1a7c    52                      push edx
'006a1a7d    6a09                    push 09

' *** Reference to "__vbaFreeVarList"
'006a1a7f    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_6, var_7, var_8, var_18, var_19, var_20, var_21, var_22, var_23)
'006a1a85    83c454                  add esp, 54
'006a1a88    c3                      ret
'006a1a89    c3                      ret
'006a1a8a    8b4508                  mov eax, dword ptr [ebp+08]
'006a1a8d    8b08                    mov ecx, dword ptr [eax]
'006a1a8f    50                      push eax
'006a1a90    ff5108                  call dword ptr [ecx+08]
'006a1a93    8b45f4                  mov eax, dword ptr [ebp-0c]
'006a1a96    8b4de4                  mov ecx, dword ptr [ebp-1c]
    'Reference to '__except_list'
'006a1a99    64890d00000000          mov dword ptr fs:[00000000], ecx
'006a1aa0    5f                      pop edi
'006a1aa1    5e                      pop esi
'006a1aa2    5b                      pop ebx
'006a1aa3    8be5                    mov esp, ebp
'006a1aa5    5d                      pop ebp
'006a1aa6    c20400                  ret 0004


End Sub


Private Sub Form_Resize()
'006a2800    55                      push ebp
'006a2801    8bec                    mov ebp, esp
'006a2803    83ec0c                  sub esp, 0c

' *** Reference to "__vbaExceptHandler"
'006a2806    6866724000              push 00407266
'006a280b    64a100000000            mov ax, word ptr fs:[00000000]
'006a2811    50                      push eax
    'Reference to '__except_list'
'006a2812    64892500000000          mov dword ptr fs:[00000000], esp
'006a2819    81ecac000000            sub esp, 000000ac
'006a281f    53                      push ebx
'006a2820    56                      push esi
'006a2821    57                      push edi
'006a2822    8965f4                  mov dword ptr [ebp-0c], esp
'006a2825    c745f8a8624000          mov dword ptr [ebp-08], 004062a8
'006a282c    8b7508                  mov esi, dword ptr [ebp+08]
'006a282f    8bc6                    mov eax, esi
'006a2831    83e001                  and eax, 01
'006a2834    8945fc                  mov dword ptr [ebp-04], eax
'006a2837    83e6fe                  and esi, fffffffe
'006a283a    8b0e                    mov ecx, dword ptr [esi]
'006a283c    56                      push esi
'006a283d    897508                  mov dword ptr [ebp+08], esi
'006a2840    ff5104                  call dword ptr [ecx+04]
'006a2843    8b16                    mov edx, dword ptr [esi]
'006a2845    33ff                    xor edi, edi
'006a2847    56                      push esi
'006a2848    897de8                  mov dword ptr [ebp-18], edi
'006a284b    897de4                  mov dword ptr [ebp-1c], edi
'006a284e    897dd4                  mov dword ptr [ebp-2c], edi
'006a2851    897dc4                  mov dword ptr [ebp-3c], edi
'006a2854    897da0                  mov dword ptr [ebp-60], edi
'006a2857    ff92fc020000            call dword ptr [edx+000002fc]
'006a285d    50                      push eax
'006a285e    8d45e8                  lea eax, dword ptr [ebp-18]
'006a2861    50                      push eax

' *** Reference to "__vbaObjSet"
'006a2862    ff15b4104000            call dword ptr [004010b4]
Set var_41 = Me
'006a2868    8bd8                    mov ebx, eax
'006a286a    393d10b27200            cmp dword ptr [0072b210], edi
'006a2870    7510                    jne 6a2882
'006a2872    6810b27200              push 0072b210
'006a2877    68c8b44000              push 0040b4c8

' *** Reference to "__vbaNew2"
'006a287c    ff152c124000            call dword ptr [0040122c]
Dim var_2479 As New FrmLstArticle
'006a2882    8b3d10b27200            mov edi, dword ptr [0072b210]
'006a2888    8b0f                    mov ecx, dword ptr [edi]
'006a288a    8d55a0                  lea edx, dword ptr [ebp-60]
'006a288d    52                      push edx
'006a288e    57                      push edi
'006a288f    ff9188000000            call dword ptr [ecx+00000088]
'006a2895    dbe2                    fnclex
'006a2897    85c0                    test ax, ax
'006a2899    7d12                    jge 6a28ad
'006a289b    6888000000              push 00000088
'006a28a0    6838f64400              push 0044f638
'006a28a5    57                      push edi
'006a28a6    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006a28a7    ff1580104000            call dword ptr [00401080]
'006a28ad    d945a0                  fld dword ptr [ebp-60]
'006a28b0    8d4dd4                  lea ecx, dword ptr [ebp-2c]
'006a28b3    d825a4624000            fsub dword ptr [004062a4]
'006a28b9    c745c404000000          mov dword ptr [ebp-3c], 00000004
'006a28c0    c745dc01000000          mov dword ptr [ebp-24], 00000001
'006a28c7    c745d402000000          mov dword ptr [ebp-2c], 00000002
'006a28ce    d95dcc                  fstp dword ptr [ebp-34]
'var_43 = (00)
'006a28d1    dfe0                    fnstsw ax
'006a28d3    a80d                    test al, 0d
'006a28d5    0f854d070000            jne 6a3028
'006a28db    8b3b                    mov edi, dword ptr [ebx]
'006a28dd    8d45c4                  lea eax, dword ptr [ebp-3c]
'006a28e0    50                      push eax
'006a28e1    51                      push ecx
'006a28e2    e809fee4ff              call 4f26f0
Call sub_4F26F0()
'006a28e7    0fbfd0                  movsx edx, ax
'006a28ea    89957cffffff            mov dword ptr [ebp+ffffff7c], edx
'006a28f0    db857cffffff            fild dword ptr [ebp+ffffff7c]
'006a28f6    d99d78ffffff            fstp dword ptr [ebp+ffffff78]
'var_87 = (00)
'006a28fc    8b8578ffffff            mov eax, dword ptr [ebp+ffffff78]
'006a2902    50                      push eax
'006a2903    53                      push ebx
'006a2904    ff9784000000            call dword ptr [edi+00000084]
'006a290a    dbe2                    fnclex
'006a290c    85c0                    test ax, ax
'006a290e    7d12                    jge 6a2922
'006a2910    6884000000              push 00000084
'006a2915    68d00d4300              push 00430dd0
'006a291a    53                      push ebx
'006a291b    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006a291c    ff1580104000            call dword ptr [00401080]
'006a2922    8d4de8                  lea ecx, dword ptr [ebp-18]

' *** Reference to "__vbaFreeObj"
'006a2925    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_41)
'006a292b    8d4dc4                  lea ecx, dword ptr [ebp-3c]
'006a292e    51                      push ecx
'006a292f    8d55d4                  lea edx, dword ptr [ebp-2c]
'006a2932    52                      push edx
'006a2933    6a02                    push 02

' *** Reference to "__vbaFreeVarList"
'006a2935    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_44, var_9)
'006a293b    8b06                    mov eax, dword ptr [esi]
'006a293d    83c40c                  add esp, 0c
'006a2940    56                      push esi
'006a2941    ff90fc020000            call dword ptr [eax+000002fc]
'006a2947    50                      push eax
'006a2948    8d4de8                  lea ecx, dword ptr [ebp-18]
'006a294b    51                      push ecx

' *** Reference to "__vbaObjSet"
'006a294c    ff15b4104000            call dword ptr [004010b4]
Set var_41 = Nothing
'006a2952    8bf8                    mov edi, eax
'006a2954    a110b27200              mov ax, word ptr [0072b210]
'006a2959    85c0                    test ax, ax
'006a295b    7510                    jne 6a296d
'006a295d    6810b27200              push 0072b210
'006a2962    68c8b44000              push 0040b4c8

' *** Reference to "__vbaNew2"
'006a2967    ff152c124000            call dword ptr [0040122c]
Set var_2479 = New FrmLstArticle
'006a296d    8b1d10b27200            mov ebx, dword ptr [0072b210]
'006a2973    8b13                    mov edx, dword ptr [ebx]
'006a2975    8d45a0                  lea eax, dword ptr [ebp-60]
'006a2978    50                      push eax
'006a2979    53                      push ebx
'006a297a    ff9280000000            call dword ptr [edx+00000080]
'006a2980    dbe2                    fnclex
'006a2982    85c0                    test ax, ax
'006a2984    7d12                    jge 6a2998
'006a2986    6880000000              push 00000080
'006a298b    6838f64400              push 0044f638
'006a2990    53                      push ebx
'006a2991    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006a2992    ff1580104000            call dword ptr [00401080]
'006a2998    d945a0                  fld dword ptr [ebp-60]
'006a299b    8d4dc4                  lea ecx, dword ptr [ebp-3c]
'006a299e    d825a0624000            fsub dword ptr [004062a0]
'006a29a4    51                      push ecx
'006a29a5    8d55d4                  lea edx, dword ptr [ebp-2c]
'006a29a8    c745c404000000          mov dword ptr [ebp-3c], 00000004
'006a29af    d95dcc                  fstp dword ptr [ebp-34]
'var_43 = (00)
'006a29b2    dfe0                    fnstsw ax
'006a29b4    a80d                    test al, 0d
'006a29b6    0f856c060000            jne 6a3028
'006a29bc    c745dc01000000          mov dword ptr [ebp-24], 00000001
'006a29c3    c745d402000000          mov dword ptr [ebp-2c], 00000002
'006a29ca    8b1f                    mov ebx, dword ptr [edi]
'006a29cc    52                      push edx
'006a29cd    e81efde4ff              call 4f26f0
Call sub_4F26F0()
'006a29d2    0fbfc0                  movsx eax, ax
'006a29d5    898574ffffff            mov dword ptr [ebp+ffffff74], eax
'006a29db    db8574ffffff            fild dword ptr [ebp+ffffff74]
'006a29e1    d99d70ffffff            fstp dword ptr [ebp+ffffff70]
'var_19 = (00)
'006a29e7    8b8d70ffffff            mov ecx, dword ptr [ebp+ffffff70]
'006a29ed    51                      push ecx
'006a29ee    57                      push edi
'006a29ef    ff537c                  call dword ptr [ebx+7c]
'006a29f2    dbe2                    fnclex
'006a29f4    85c0                    test ax, ax
'006a29f6    7d0f                    jge 6a2a07
'006a29f8    6a7c                    push 7c
'006a29fa    68d00d4300              push 00430dd0
'006a29ff    57                      push edi
'006a2a00    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006a2a01    ff1580104000            call dword ptr [00401080]
'006a2a07    8d4de8                  lea ecx, dword ptr [ebp-18]

' *** Reference to "__vbaFreeObj"
'006a2a0a    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_41)
'006a2a10    8d55c4                  lea edx, dword ptr [ebp-3c]
'006a2a13    52                      push edx
'006a2a14    8d45d4                  lea eax, dword ptr [ebp-2c]
'006a2a17    50                      push eax
'006a2a18    6a02                    push 02

' *** Reference to "__vbaFreeVarList"
'006a2a1a    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_44, var_9)
'006a2a20    8b0e                    mov ecx, dword ptr [esi]
'006a2a22    83c40c                  add esp, 0c
'006a2a25    56                      push esi
'006a2a26    ff910c030000            call dword ptr [ecx+0000030c]
'006a2a2c    50                      push eax
'006a2a2d    8d55e8                  lea edx, dword ptr [ebp-18]
'006a2a30    52                      push edx

' *** Reference to "__vbaObjSet"
'006a2a31    ff15b4104000            call dword ptr [004010b4]
Set var_41 = 
'006a2a37    8bf8                    mov edi, eax
'006a2a39    a110b27200              mov ax, word ptr [0072b210]
'006a2a3e    85c0                    test ax, ax
'006a2a40    7510                    jne 6a2a52
'006a2a42    6810b27200              push 0072b210
'006a2a47    68c8b44000              push 0040b4c8

' *** Reference to "__vbaNew2"
'006a2a4c    ff152c124000            call dword ptr [0040122c]
Set var_2479 = New FrmLstArticle
'006a2a52    8b1d10b27200            mov ebx, dword ptr [0072b210]
'006a2a58    8b03                    mov eax, dword ptr [ebx]
'006a2a5a    8d4da0                  lea ecx, dword ptr [ebp-60]
'006a2a5d    51                      push ecx
'006a2a5e    53                      push ebx
'006a2a5f    ff9088000000            call dword ptr [eax+00000088]
'006a2a65    dbe2                    fnclex
'006a2a67    85c0                    test ax, ax
'006a2a69    7d12                    jge 6a2a7d
'006a2a6b    6888000000              push 00000088
'006a2a70    6838f64400              push 0044f638
'006a2a75    53                      push ebx
'006a2a76    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006a2a77    ff1580104000            call dword ptr [00401080]
'006a2a7d    d945a0                  fld dword ptr [ebp-60]
'006a2a80    8d55c4                  lea edx, dword ptr [ebp-3c]
'006a2a83    d8259c624000            fsub dword ptr [0040629c]
'006a2a89    52                      push edx
'006a2a8a    c745c404000000          mov dword ptr [ebp-3c], 00000004
'006a2a91    c745dc64140000          mov dword ptr [ebp-24], 00001464
'006a2a98    d95dcc                  fstp dword ptr [ebp-34]
'var_43 = (00)
'006a2a9b    dfe0                    fnstsw ax
'006a2a9d    a80d                    test al, 0d
'006a2a9f    0f8583050000            jne 6a3028
'006a2aa5    8d45d4                  lea eax, dword ptr [ebp-2c]
'006a2aa8    c745d402000000          mov dword ptr [ebp-2c], 00000002
'006a2aaf    8b1f                    mov ebx, dword ptr [edi]
'006a2ab1    50                      push eax
'006a2ab2    e839fce4ff              call 4f26f0
Call sub_4F26F0()
'006a2ab7    0fbfc8                  movsx ecx, ax
'006a2aba    898d6cffffff            mov dword ptr [ebp+ffffff6c], ecx
'006a2ac0    db856cffffff            fild dword ptr [ebp+ffffff6c]
'006a2ac6    d99d68ffffff            fstp dword ptr [ebp+ffffff68]
'var_132 = (00)
'006a2acc    8b9568ffffff            mov edx, dword ptr [ebp+ffffff68]
'006a2ad2    52                      push edx
'006a2ad3    57                      push edi
'006a2ad4    ff5374                  call dword ptr [ebx+74]
'006a2ad7    dbe2                    fnclex
'006a2ad9    85c0                    test ax, ax
'006a2adb    7d0f                    jge 6a2aec
'006a2add    6a74                    push 74
'006a2adf    6838eb4300              push 0043eb38
'006a2ae4    57                      push edi
'006a2ae5    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006a2ae6    ff1580104000            call dword ptr [00401080]
'006a2aec    8d4de8                  lea ecx, dword ptr [ebp-18]

' *** Reference to "__vbaFreeObj"
'006a2aef    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_41)
'006a2af5    8d45c4                  lea eax, dword ptr [ebp-3c]
'006a2af8    50                      push eax
'006a2af9    8d4dd4                  lea ecx, dword ptr [ebp-2c]
'006a2afc    51                      push ecx
'006a2afd    6a02                    push 02

' *** Reference to "__vbaFreeVarList"
'006a2aff    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_44, var_9)
'006a2b05    8b16                    mov edx, dword ptr [esi]
'006a2b07    83c40c                  add esp, 0c
'006a2b0a    56                      push esi
'006a2b0b    ff9208030000            call dword ptr [edx+00000308]
'006a2b11    50                      push eax
'006a2b12    8d45e8                  lea eax, dword ptr [ebp-18]
'006a2b15    50                      push eax

' *** Reference to "__vbaObjSet"
'006a2b16    ff15b4104000            call dword ptr [004010b4]
Set var_41 = 
'006a2b1c    8bf8                    mov edi, eax
'006a2b1e    a110b27200              mov ax, word ptr [0072b210]
'006a2b23    85c0                    test ax, ax
'006a2b25    7510                    jne 6a2b37
'006a2b27    6810b27200              push 0072b210
'006a2b2c    68c8b44000              push 0040b4c8

' *** Reference to "__vbaNew2"
'006a2b31    ff152c124000            call dword ptr [0040122c]
Set var_2479 = New FrmLstArticle
'006a2b37    8b1d10b27200            mov ebx, dword ptr [0072b210]
'006a2b3d    8b0b                    mov ecx, dword ptr [ebx]
'006a2b3f    8d55a0                  lea edx, dword ptr [ebp-60]
'006a2b42    52                      push edx
'006a2b43    53                      push ebx
'006a2b44    ff9188000000            call dword ptr [ecx+00000088]
'006a2b4a    dbe2                    fnclex
'006a2b4c    85c0                    test ax, ax
'006a2b4e    7d12                    jge 6a2b62
'006a2b50    6888000000              push 00000088
'006a2b55    6838f64400              push 0044f638
'006a2b5a    53                      push ebx
'006a2b5b    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006a2b5c    ff1580104000            call dword ptr [00401080]
'006a2b62    d945a0                  fld dword ptr [ebp-60]
'006a2b65    8d4dd4                  lea ecx, dword ptr [ebp-2c]
'006a2b68    d8259c624000            fsub dword ptr [0040629c]
'006a2b6e    c745c404000000          mov dword ptr [ebp-3c], 00000004
'006a2b75    c745dc64140000          mov dword ptr [ebp-24], 00001464
'006a2b7c    c745d402000000          mov dword ptr [ebp-2c], 00000002
'006a2b83    d95dcc                  fstp dword ptr [ebp-34]
'var_43 = (00)
'006a2b86    dfe0                    fnstsw ax
'006a2b88    a80d                    test al, 0d
'006a2b8a    0f8598040000            jne 6a3028
'006a2b90    8b1f                    mov ebx, dword ptr [edi]
'006a2b92    8d45c4                  lea eax, dword ptr [ebp-3c]
'006a2b95    50                      push eax
'006a2b96    51                      push ecx
'006a2b97    e854fbe4ff              call 4f26f0
Call sub_4F26F0()
'006a2b9c    0fbfd0                  movsx edx, ax
'006a2b9f    899564ffffff            mov dword ptr [ebp+ffffff64], edx
'006a2ba5    db8564ffffff            fild dword ptr [ebp+ffffff64]
'006a2bab    d99d60ffffff            fstp dword ptr [ebp+ffffff60]
'var_20 = (00)
'006a2bb1    8b8560ffffff            mov eax, dword ptr [ebp+ffffff60]
'006a2bb7    50                      push eax
'006a2bb8    57                      push edi
'006a2bb9    ff5374                  call dword ptr [ebx+74]
'006a2bbc    dbe2                    fnclex
'006a2bbe    85c0                    test ax, ax
'006a2bc0    7d0f                    jge 6a2bd1
'006a2bc2    6a74                    push 74
'006a2bc4    6838eb4300              push 0043eb38
'006a2bc9    57                      push edi
'006a2bca    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006a2bcb    ff1580104000            call dword ptr [00401080]
'006a2bd1    8d4de8                  lea ecx, dword ptr [ebp-18]

' *** Reference to "__vbaFreeObj"
'006a2bd4    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_41)
'006a2bda    8d4dc4                  lea ecx, dword ptr [ebp-3c]
'006a2bdd    51                      push ecx
'006a2bde    8d55d4                  lea edx, dword ptr [ebp-2c]
'006a2be1    52                      push edx
'006a2be2    6a02                    push 02

' *** Reference to "__vbaFreeVarList"
'006a2be4    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_44, var_9)
'006a2bea    8b06                    mov eax, dword ptr [esi]
'006a2bec    83c40c                  add esp, 0c
'006a2bef    56                      push esi
'006a2bf0    ff9004030000            call dword ptr [eax+00000304]
'006a2bf6    50                      push eax
'006a2bf7    8d4de8                  lea ecx, dword ptr [ebp-18]
'006a2bfa    51                      push ecx

' *** Reference to "__vbaObjSet"
'006a2bfb    ff15b4104000            call dword ptr [004010b4]
Set var_41 = Nothing
'006a2c01    8bf8                    mov edi, eax
'006a2c03    a110b27200              mov ax, word ptr [0072b210]
'006a2c08    85c0                    test ax, ax
'006a2c0a    7510                    jne 6a2c1c
'006a2c0c    6810b27200              push 0072b210
'006a2c11    68c8b44000              push 0040b4c8

' *** Reference to "__vbaNew2"
'006a2c16    ff152c124000            call dword ptr [0040122c]
Set var_2479 = New FrmLstArticle
'006a2c1c    8b1d10b27200            mov ebx, dword ptr [0072b210]
'006a2c22    8b13                    mov edx, dword ptr [ebx]
'006a2c24    8d45a0                  lea eax, dword ptr [ebp-60]
'006a2c27    50                      push eax
'006a2c28    53                      push ebx
'006a2c29    ff9288000000            call dword ptr [edx+00000088]
'006a2c2f    dbe2                    fnclex
'006a2c31    85c0                    test ax, ax
'006a2c33    7d12                    jge 6a2c47
'006a2c35    6888000000              push 00000088
'006a2c3a    6838f64400              push 0044f638
'006a2c3f    53                      push ebx
'006a2c40    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006a2c41    ff1580104000            call dword ptr [00401080]
'006a2c47    d945a0                  fld dword ptr [ebp-60]
'006a2c4a    8d4dc4                  lea ecx, dword ptr [ebp-3c]
'006a2c4d    d8259c624000            fsub dword ptr [0040629c]
'006a2c53    51                      push ecx
'006a2c54    8d55d4                  lea edx, dword ptr [ebp-2c]
'006a2c57    c745c404000000          mov dword ptr [ebp-3c], 00000004
'006a2c5e    d95dcc                  fstp dword ptr [ebp-34]
'var_43 = (00)
'006a2c61    dfe0                    fnstsw ax
'006a2c63    a80d                    test al, 0d
'006a2c65    0f85bd030000            jne 6a3028
'006a2c6b    c745dc64140000          mov dword ptr [ebp-24], 00001464
'006a2c72    c745d402000000          mov dword ptr [ebp-2c], 00000002
'006a2c79    8b1f                    mov ebx, dword ptr [edi]
'006a2c7b    52                      push edx
'006a2c7c    e86ffae4ff              call 4f26f0
Call sub_4F26F0()
'006a2c81    0fbfc0                  movsx eax, ax
'006a2c84    89855cffffff            mov dword ptr [ebp+ffffff5c], eax
'006a2c8a    db855cffffff            fild dword ptr [ebp+ffffff5c]
'006a2c90    d99d58ffffff            fstp dword ptr [ebp+ffffff58]
'var_133 = (00)
'006a2c96    8b8d58ffffff            mov ecx, dword ptr [ebp+ffffff58]
'006a2c9c    51                      push ecx
'006a2c9d    57                      push edi
'006a2c9e    ff5374                  call dword ptr [ebx+74]
'006a2ca1    dbe2                    fnclex
'006a2ca3    85c0                    test ax, ax
'006a2ca5    7d0f                    jge 6a2cb6
'006a2ca7    6a74                    push 74
'006a2ca9    68d00d4300              push 00430dd0
'006a2cae    57                      push edi
'006a2caf    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006a2cb0    ff1580104000            call dword ptr [00401080]
'006a2cb6    8d4de8                  lea ecx, dword ptr [ebp-18]

' *** Reference to "__vbaFreeObj"
'006a2cb9    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_41)
'006a2cbf    8d55c4                  lea edx, dword ptr [ebp-3c]
'006a2cc2    52                      push edx
'006a2cc3    8d45d4                  lea eax, dword ptr [ebp-2c]
'006a2cc6    50                      push eax
'006a2cc7    6a02                    push 02

' *** Reference to "__vbaFreeVarList"
'006a2cc9    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_44, var_9)
'006a2ccf    8b0e                    mov ecx, dword ptr [esi]
'006a2cd1    83c40c                  add esp, 0c
'006a2cd4    56                      push esi
'006a2cd5    ff9100030000            call dword ptr [ecx+00000300]
'006a2cdb    50                      push eax
'006a2cdc    8d55e8                  lea edx, dword ptr [ebp-18]
'006a2cdf    52                      push edx

' *** Reference to "__vbaObjSet"
'006a2ce0    ff15b4104000            call dword ptr [004010b4]
Set var_41 = 
'006a2ce6    8bf8                    mov edi, eax
'006a2ce8    a110b27200              mov ax, word ptr [0072b210]
'006a2ced    85c0                    test ax, ax
'006a2cef    7510                    jne 6a2d01
'006a2cf1    6810b27200              push 0072b210
'006a2cf6    68c8b44000              push 0040b4c8

' *** Reference to "__vbaNew2"
'006a2cfb    ff152c124000            call dword ptr [0040122c]
Set var_2479 = New FrmLstArticle
'006a2d01    8b1d10b27200            mov ebx, dword ptr [0072b210]
'006a2d07    8b03                    mov eax, dword ptr [ebx]
'006a2d09    8d4da0                  lea ecx, dword ptr [ebp-60]
'006a2d0c    51                      push ecx
'006a2d0d    53                      push ebx
'006a2d0e    ff9088000000            call dword ptr [eax+00000088]
'006a2d14    dbe2                    fnclex
'006a2d16    85c0                    test ax, ax
'006a2d18    7d12                    jge 6a2d2c
'006a2d1a    6888000000              push 00000088
'006a2d1f    6838f64400              push 0044f638
'006a2d24    53                      push ebx
'006a2d25    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006a2d26    ff1580104000            call dword ptr [00401080]
'006a2d2c    d945a0                  fld dword ptr [ebp-60]
'006a2d2f    8d55c4                  lea edx, dword ptr [ebp-3c]
'006a2d32    d8259c624000            fsub dword ptr [0040629c]
'006a2d38    52                      push edx
'006a2d39    c745c404000000          mov dword ptr [ebp-3c], 00000004
'006a2d40    c745dc64140000          mov dword ptr [ebp-24], 00001464
'006a2d47    d95dcc                  fstp dword ptr [ebp-34]
'var_43 = (00)
'006a2d4a    dfe0                    fnstsw ax
'006a2d4c    a80d                    test al, 0d
'006a2d4e    0f85d4020000            jne 6a3028
'006a2d54    8d45d4                  lea eax, dword ptr [ebp-2c]
'006a2d57    c745d402000000          mov dword ptr [ebp-2c], 00000002
'006a2d5e    8b1f                    mov ebx, dword ptr [edi]
'006a2d60    50                      push eax
'006a2d61    e88af9e4ff              call 4f26f0
Call sub_4F26F0()
'006a2d66    0fbfc8                  movsx ecx, ax
'006a2d69    898d54ffffff            mov dword ptr [ebp+ffffff54], ecx
'006a2d6f    db8554ffffff            fild dword ptr [ebp+ffffff54]
'006a2d75    d99d50ffffff            fstp dword ptr [ebp+ffffff50]
'var_21 = (00)
'006a2d7b    8b9550ffffff            mov edx, dword ptr [ebp+ffffff50]
'006a2d81    52                      push edx
'006a2d82    57                      push edi
'006a2d83    ff5374                  call dword ptr [ebx+74]
'006a2d86    dbe2                    fnclex
'006a2d88    85c0                    test ax, ax
'006a2d8a    7d0f                    jge 6a2d9b
'006a2d8c    6a74                    push 74
'006a2d8e    68d00d4300              push 00430dd0
'006a2d93    57                      push edi
'006a2d94    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006a2d95    ff1580104000            call dword ptr [00401080]
'006a2d9b    8d4de8                  lea ecx, dword ptr [ebp-18]

' *** Reference to "__vbaFreeObj"
'006a2d9e    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_41)
'006a2da4    8d45c4                  lea eax, dword ptr [ebp-3c]
'006a2da7    50                      push eax
'006a2da8    8d4dd4                  lea ecx, dword ptr [ebp-2c]
'006a2dab    51                      push ecx
'006a2dac    6a02                    push 02

' *** Reference to "__vbaFreeVarList"
'006a2dae    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_44, var_9)
'006a2db4    8b16                    mov edx, dword ptr [esi]
'006a2db6    83c40c                  add esp, 0c
'006a2db9    56                      push esi
'006a2dba    ff9210030000            call dword ptr [edx+00000310]
'006a2dc0    50                      push eax
'006a2dc1    8d45e8                  lea eax, dword ptr [ebp-18]
'006a2dc4    50                      push eax

' *** Reference to "__vbaObjSet"
'006a2dc5    ff15b4104000            call dword ptr [004010b4]
Set var_41 = 
'006a2dcb    8d55e4                  lea edx, dword ptr [ebp-1c]
'006a2dce    52                      push edx
'006a2dcf    8bf8                    mov edi, eax
'006a2dd1    8b0f                    mov ecx, dword ptr [edi]
'006a2dd3    6a00                    push 00
'006a2dd5    57                      push edi
'006a2dd6    ff5140                  call dword ptr [ecx+40]
'006a2dd9    dbe2                    fnclex
'006a2ddb    85c0                    test ax, ax
'006a2ddd    7d0f                    jge 6a2dee
'006a2ddf    6a40                    push 40
'006a2de1    686c384300              push 0043386c
'006a2de6    57                      push edi
'006a2de7    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006a2de8    ff1580104000            call dword ptr [00401080]
'006a2dee    a110b27200              mov ax, word ptr [0072b210]
'006a2df3    85c0                    test ax, ax
'006a2df5    8b5de4                  mov ebx, dword ptr [ebp-1c]
'006a2df8    7510                    jne 6a2e0a
'006a2dfa    6810b27200              push 0072b210
'006a2dff    68c8b44000              push 0040b4c8

' *** Reference to "__vbaNew2"
'006a2e04    ff152c124000            call dword ptr [0040122c]
Set var_2479 = New FrmLstArticle
'006a2e0a    8b3d10b27200            mov edi, dword ptr [0072b210]
'006a2e10    8b07                    mov eax, dword ptr [edi]
'006a2e12    8d4da0                  lea ecx, dword ptr [ebp-60]
'006a2e15    51                      push ecx
'006a2e16    57                      push edi
'006a2e17    ff9088000000            call dword ptr [eax+00000088]
'006a2e1d    dbe2                    fnclex
'006a2e1f    85c0                    test ax, ax
'006a2e21    7d12                    jge 6a2e35
'006a2e23    6888000000              push 00000088
'006a2e28    6838f64400              push 0044f638
'006a2e2d    57                      push edi
'006a2e2e    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006a2e2f    ff1580104000            call dword ptr [00401080]
'006a2e35    d945a0                  fld dword ptr [ebp-60]
'006a2e38    8d55c4                  lea edx, dword ptr [ebp-3c]
'006a2e3b    d8259c624000            fsub dword ptr [0040629c]
'006a2e41    52                      push edx
'006a2e42    c745c404000000          mov dword ptr [ebp-3c], 00000004
'006a2e49    c745dc64140000          mov dword ptr [ebp-24], 00001464
'006a2e50    d95dcc                  fstp dword ptr [ebp-34]
'var_43 = (00)
'006a2e53    dfe0                    fnstsw ax
'006a2e55    a80d                    test al, 0d
'006a2e57    0f85cb010000            jne 6a3028
'006a2e5d    8d45d4                  lea eax, dword ptr [ebp-2c]
'006a2e60    c745d402000000          mov dword ptr [ebp-2c], 00000002
'006a2e67    8b3b                    mov edi, dword ptr [ebx]
'006a2e69    50                      push eax
'006a2e6a    e881f8e4ff              call 4f26f0
Call sub_4F26F0()
'006a2e6f    0fbfc8                  movsx ecx, ax
'006a2e72    898d4cffffff            mov dword ptr [ebp+ffffff4c], ecx
'006a2e78    db854cffffff            fild dword ptr [ebp+ffffff4c]
'006a2e7e    d99d48ffffff            fstp dword ptr [ebp+ffffff48]
'var_134 = (00)
'006a2e84    8b9548ffffff            mov edx, dword ptr [ebp+ffffff48]
'006a2e8a    52                      push edx
'006a2e8b    53                      push ebx
'006a2e8c    ff577c                  call dword ptr [edi+7c]
'006a2e8f    dbe2                    fnclex
'006a2e91    85c0                    test ax, ax
'006a2e93    7d13                    jge 6a2ea8
'006a2e95    6a7c                    push 7c
'006a2e97    683ce04300              push 0043e03c
'006a2e9c    53                      push ebx

' *** Reference to "__vbaHresultCheckObj"
'006a2e9d    8b1d80104000            mov ebx, dword ptr [00401080]
'006a2ea3    50                      push eax
'006a2ea4    ffd3                    call ebx
'006a2ea6    eb06                    jmp 6a2eae

' *** Reference to "__vbaHresultCheckObj"
'006a2ea8    8b1d80104000            mov ebx, dword ptr [00401080]
'006a2eae    8d45e4                  lea eax, dword ptr [ebp-1c]
'006a2eb1    50                      push eax
'006a2eb2    8d4de8                  lea ecx, dword ptr [ebp-18]
'006a2eb5    51                      push ecx
'006a2eb6    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'006a2eb8    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_41, var_40)
'006a2ebe    8d55c4                  lea edx, dword ptr [ebp-3c]
'006a2ec1    52                      push edx
'006a2ec2    8d45d4                  lea eax, dword ptr [ebp-2c]
'006a2ec5    50                      push eax
'006a2ec6    6a02                    push 02

' *** Reference to "__vbaFreeVarList"
'006a2ec8    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_44, var_9)
'006a2ece    8b0e                    mov ecx, dword ptr [esi]
'006a2ed0    83c418                  add esp, 18
'006a2ed3    56                      push esi
'006a2ed4    ff9110030000            call dword ptr [ecx+00000310]
'006a2eda    50                      push eax
'006a2edb    8d55e8                  lea edx, dword ptr [ebp-18]
'006a2ede    52                      push edx

' *** Reference to "__vbaObjSet"
'006a2edf    ff15b4104000            call dword ptr [004010b4]
Set var_41 = 
'006a2ee5    8d4de4                  lea ecx, dword ptr [ebp-1c]
'006a2ee8    51                      push ecx
'006a2ee9    8bf0                    mov esi, eax
'006a2eeb    8b06                    mov eax, dword ptr [esi]
'006a2eed    6a01                    push 01
'006a2eef    56                      push esi
'006a2ef0    ff5040                  call dword ptr [eax+40]
'006a2ef3    dbe2                    fnclex
'006a2ef5    85c0                    test ax, ax
'006a2ef7    7d0b                    jge 6a2f04
'006a2ef9    6a40                    push 40
'006a2efb    686c384300              push 0043386c
'006a2f00    56                      push esi
'006a2f01    50                      push eax
'006a2f02    ffd3                    call ebx
'006a2f04    a110b27200              mov ax, word ptr [0072b210]
'006a2f09    85c0                    test ax, ax
'006a2f0b    8b7de4                  mov edi, dword ptr [ebp-1c]
'006a2f0e    7510                    jne 6a2f20
'006a2f10    6810b27200              push 0072b210
'006a2f15    68c8b44000              push 0040b4c8

' *** Reference to "__vbaNew2"
'006a2f1a    ff152c124000            call dword ptr [0040122c]
Set var_2479 = New FrmLstArticle
'006a2f20    8b3510b27200            mov esi, dword ptr [0072b210]
'006a2f26    8b16                    mov edx, dword ptr [esi]
'006a2f28    8d45a0                  lea eax, dword ptr [ebp-60]
'006a2f2b    50                      push eax
'006a2f2c    56                      push esi
'006a2f2d    ff9288000000            call dword ptr [edx+00000088]
'006a2f33    dbe2                    fnclex
'006a2f35    85c0                    test ax, ax
'006a2f37    7d0e                    jge 6a2f47
'006a2f39    6888000000              push 00000088
'006a2f3e    6838f64400              push 0044f638
'006a2f43    56                      push esi
'006a2f44    50                      push eax
'006a2f45    ffd3                    call ebx
'006a2f47    d945a0                  fld dword ptr [ebp-60]
'006a2f4a    8d4dc4                  lea ecx, dword ptr [ebp-3c]
'006a2f4d    d8259c624000            fsub dword ptr [0040629c]
'006a2f53    51                      push ecx
'006a2f54    8d55d4                  lea edx, dword ptr [ebp-2c]
'006a2f57    c745c404000000          mov dword ptr [ebp-3c], 00000004
'006a2f5e    d95dcc                  fstp dword ptr [ebp-34]
'var_43 = (00)
'006a2f61    dfe0                    fnstsw ax
'006a2f63    a80d                    test al, 0d
'006a2f65    0f85bd000000            jne 6a3028
'006a2f6b    c745dc64140000          mov dword ptr [ebp-24], 00001464
'006a2f72    c745d402000000          mov dword ptr [ebp-2c], 00000002
'006a2f79    8b37                    mov esi, dword ptr [edi]
'006a2f7b    52                      push edx
'006a2f7c    e86ff7e4ff              call 4f26f0
Call sub_4F26F0()
'006a2f81    0fbfc0                  movsx eax, ax
'006a2f84    898544ffffff            mov dword ptr [ebp+ffffff44], eax
'006a2f8a    db8544ffffff            fild dword ptr [ebp+ffffff44]
'006a2f90    d99d40ffffff            fstp dword ptr [ebp+ffffff40]
'var_22 = (00)
'006a2f96    8b8d40ffffff            mov ecx, dword ptr [ebp+ffffff40]
'006a2f9c    51                      push ecx
'006a2f9d    57                      push edi
'006a2f9e    ff567c                  call dword ptr [esi+7c]
'006a2fa1    dbe2                    fnclex
'006a2fa3    85c0                    test ax, ax
'006a2fa5    7d0b                    jge 6a2fb2
'006a2fa7    6a7c                    push 7c
'006a2fa9    683ce04300              push 0043e03c
'006a2fae    57                      push edi
'006a2faf    50                      push eax
'006a2fb0    ffd3                    call ebx
'006a2fb2    8d55e4                  lea edx, dword ptr [ebp-1c]
'006a2fb5    52                      push edx
'006a2fb6    8d45e8                  lea eax, dword ptr [ebp-18]
'006a2fb9    50                      push eax
'006a2fba    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'006a2fbc    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_41, var_40)
'006a2fc2    8d4dc4                  lea ecx, dword ptr [ebp-3c]
'006a2fc5    51                      push ecx
'006a2fc6    8d55d4                  lea edx, dword ptr [ebp-2c]
'006a2fc9    52                      push edx
'006a2fca    6a02                    push 02

' *** Reference to "__vbaFreeVarList"
'006a2fcc    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_44, var_9)
'006a2fd2    83c418                  add esp, 18
'006a2fd5    c745fc00000000          mov dword ptr [ebp-04], 00000000
'006a2fdc    9b                      fwait
'006a2fdd    6809306a00              push 006a3009
'006a2fe2    eb24                    jmp 6a3008
'006a2fe4    8d45e4                  lea eax, dword ptr [ebp-1c]
'006a2fe7    50                      push eax
'006a2fe8    8d4de8                  lea ecx, dword ptr [ebp-18]
'006a2feb    51                      push ecx
'006a2fec    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'006a2fee    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_41, var_40)
'006a2ff4    8d55c4                  lea edx, dword ptr [ebp-3c]
'006a2ff7    52                      push edx
'006a2ff8    8d45d4                  lea eax, dword ptr [ebp-2c]
'006a2ffb    50                      push eax
'006a2ffc    6a02                    push 02

' *** Reference to "__vbaFreeVarList"
'006a2ffe    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_44, var_9)
'006a3004    83c418                  add esp, 18
'006a3007    c3                      ret
'006a3008    c3                      ret
'006a3009    8b4508                  mov eax, dword ptr [ebp+08]
'006a300c    8b08                    mov ecx, dword ptr [eax]
'006a300e    50                      push eax
'006a300f    ff5108                  call dword ptr [ecx+08]
'006a3012    8b45fc                  mov eax, dword ptr [ebp-04]
'006a3015    8b4dec                  mov ecx, dword ptr [ebp-14]
'006a3018    5f                      pop edi
'006a3019    5e                      pop esi
    'Reference to '__except_list'
'006a301a    64890d00000000          mov dword ptr fs:[00000000], ecx
'006a3021    5b                      pop ebx
'006a3022    8be5                    mov esp, ebp
'006a3024    5d                      pop ebp
'006a3025    c20400                  ret 0004


End Sub


