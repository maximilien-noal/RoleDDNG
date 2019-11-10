VERSION 5.00

Begin VB.Form frmLanceurDes
    Caption = "Lanceur de dès"
    ScaleMode = 1
    AutoRedraw = 0              'False
    FontTransparent = -1              'True
    BorderStyle = 3
    LinkTopic = "Form1"
    MaxButton = 0              'False
    MinButton = 0              'False
    ControlBox = 0              'False
    MDIChild = -1              'True
    KeyPreview = -1              'True
    ClientLeft   = 45
    ClientTop    = 330
    ClientWidth  = 4605
    ClientHeight = 2100
    BeginProperty Font
        Name          = "Times New Roman"
        Size          = 9
        Charset       = 0
        Weight        = 400
        Underline     = 0              'False
        Italic        = 0              'False
        Strikethrough = 0              'False
    EndProperty
    ShowInTaskbar = 0              'False
    Begin VB.TextBox fldnResultat
        Left   = 1215
        Top    = 1650
        Width  = 690
        Height = 345
        TabIndex = 7
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
    Begin VB.TextBox fldstrDetailResultat
        Left   = 135
        Top    = 1170
        Width  = 4470
        Height = 375
        TabIndex = 5
    End
    Begin VB.CommandButton btnGenerer
        Left   = 2295
        Top    = 225
        Width  = 2265
        Height = 870
        TabIndex = 4
        BeginProperty Font
            Name          = "MS Sans Serif"
            Size          = 8,25
            Charset       = 0
            Weight        = 400
            Underline     = 0              'False
            Italic        = 0              'False
            Strikethrough = 0              'False
        EndProperty
        Picture = frmLanceurDes.frx:0000
        Style = 1
    End
    Begin VB.ComboBox CmbTypeDes
        Left   = 1485
        Top    = 668
        Width  = 690
        Height = 345
        Text = "6"
        TabIndex = 3
        BeginProperty Font
            Name          = "Times New Roman"
            Size          = 9,75
            Charset       = 0
            Weight        = 400
            Underline     = 0              'False
            Italic        = 0              'False
            Strikethrough = 0              'False
        EndProperty
    End
    Begin VB.TextBox fldnNombreDes
        Left   = 1485
        Top    = 270
        Width  = 690
        Height = 285
        Text = "1"
        TabIndex = 1
    End
    Begin VB.Label Label3
        Caption = "Résultat"
        Left   = 225
        Top    = 1650
        Width  = 960
        Height = 345
        TabIndex = 6
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
    Begin VB.Label Label2
        Caption = "Type de dès"
        Left   = 180
        Top    = 720
        Width  = 1140
        Height = 240
        TabIndex = 2
        BeginProperty Font
            Name          = "Times New Roman"
            Size          = 9,75
            Charset       = 0
            Weight        = 400
            Underline     = 0              'False
            Italic        = 0              'False
            Strikethrough = 0              'False
        EndProperty
    End
    Begin VB.Label Label1
        Caption = "Nombre de dès"
        Left   = 180
        Top    = 292
        Width  = 1230
        Height = 240
        TabIndex = 0
        BeginProperty Font
            Name          = "Times New Roman"
            Size          = 9,75
            Charset       = 0
            Weight        = 400
            Underline     = 0              'False
            Italic        = 0              'False
            Strikethrough = 0              'False
        EndProperty
    End
End
'Event for btnGenerer
Private Sub btnGenerer_Click()
'006cfcb0    55                      push ebp
'006cfcb1    8bec                    mov ebp, esp
'006cfcb3    83ec0c                  sub esp, 0c

' *** Reference to "__vbaExceptHandler"
'006cfcb6    6866724000              push 00407266
'006cfcbb    64a100000000            mov ax, word ptr fs:[00000000]
'006cfcc1    50                      push eax
    'Reference to '__except_list'
'006cfcc2    64892500000000          mov dword ptr fs:[00000000], esp
'006cfcc9    83ec70                  sub esp, 70
'006cfccc    53                      push ebx
'006cfccd    56                      push esi
'006cfcce    57                      push edi
'006cfccf    8965f4                  mov dword ptr [ebp-0c], esp
'006cfcd2    c745f8a8684000          mov dword ptr [ebp-08], 004068a8
'006cfcd9    8b7508                  mov esi, dword ptr [ebp+08]
'006cfcdc    8bc6                    mov eax, esi
'006cfcde    83e001                  and eax, 01
'006cfce1    8945fc                  mov dword ptr [ebp-04], eax
'006cfce4    83e6fe                  and esi, fffffffe
'006cfce7    8b0e                    mov ecx, dword ptr [esi]
'006cfce9    56                      push esi
'006cfcea    897508                  mov dword ptr [ebp+08], esi
'006cfced    ff5104                  call dword ptr [ecx+04]
'006cfcf0    8b16                    mov edx, dword ptr [esi]
'006cfcf2    33c0                    xor eax, eax
var_num1 = Empty
'006cfcf4    56                      push esi
'006cfcf5    8945dc                  mov dword ptr [ebp-24], eax
'006cfcf8    8945d8                  mov dword ptr [ebp-28], eax
'006cfcfb    8945d4                  mov dword ptr [ebp-2c], eax
'006cfcfe    8945d0                  mov dword ptr [ebp-30], eax
'006cfd01    8945cc                  mov dword ptr [ebp-34], eax
'006cfd04    8945c8                  mov dword ptr [ebp-38], eax
'006cfd07    8945b8                  mov dword ptr [ebp-48], eax
'006cfd0a    8945e4                  mov dword ptr [ebp-1c], eax
'006cfd0d    ff9200030000            call dword ptr [edx+00000300]

' *** Reference to "__vbaObjSet"
'006cfd13    8b3db4104000            mov edi, dword ptr [004010b4]
'006cfd19    50                      push eax
'006cfd1a    8d45cc                  lea eax, dword ptr [ebp-34]
'006cfd1d    50                      push eax
'006cfd1e    ffd7                    call edi
Set var_43 = Me
'006cfd20    8bd8                    mov ebx, eax
'006cfd22    8b0b                    mov ecx, dword ptr [ebx]
'006cfd24    68cc134300              push 004313cc
'006cfd29    53                      push ebx
'006cfd2a    ff91a4000000            call dword ptr [ecx+000000a4]
'006cfd30    dbe2                    fnclex
'006cfd32    85c0                    test ax, ax
'006cfd34    7d12                    jge 6cfd48
'006cfd36    68a4000000              push 000000a4
'006cfd3b    68d00d4300              push 00430dd0
'006cfd40    53                      push ebx
'006cfd41    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006cfd42    ff1580104000            call dword ptr [00401080]
'006cfd48    8d4dcc                  lea ecx, dword ptr [ebp-34]

' *** Reference to "__vbaFreeObj"
'006cfd4b    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_43)
'006cfd51    8b16                    mov edx, dword ptr [esi]
'006cfd53    56                      push esi
'006cfd54    ff920c030000            call dword ptr [edx+0000030c]
'006cfd5a    50                      push eax
'006cfd5b    8d45cc                  lea eax, dword ptr [ebp-34]
'006cfd5e    50                      push eax
'006cfd5f    ffd7                    call edi
Set var_43 = var_43
'006cfd61    8d55dc                  lea edx, dword ptr [ebp-24]
'006cfd64    8bd8                    mov ebx, eax
'006cfd66    8b0b                    mov ecx, dword ptr [ebx]
'006cfd68    52                      push edx
'006cfd69    53                      push ebx
'006cfd6a    ff91a0000000            call dword ptr [ecx+000000a0]
'006cfd70    dbe2                    fnclex
'006cfd72    85c0                    test ax, ax
'006cfd74    7d12                    jge 6cfd88
'006cfd76    68a0000000              push 000000a0
'006cfd7b    68d00d4300              push 00430dd0
'006cfd80    53                      push ebx
'006cfd81    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006cfd82    ff1580104000            call dword ptr [00401080]
'006cfd88    8b45dc                  mov eax, dword ptr [ebp-24]
'006cfd8b    50                      push eax

' *** Reference to "rtcR8ValFromBstr"
'006cfd8c    ff1510134000            call dword ptr [00401310]

' *** Reference to "__vbaFpI2"
'006cfd92    ff15a0124000            call dword ptr [004012a0]
var_num1 = CLng(Val(frmLanceurDes))
'006cfd98    8d4ddc                  lea ecx, dword ptr [ebp-24]
'006cfd9b    c745e801000000          mov dword ptr [ebp-18], 00000001
'006cfda2    894588                  mov dword ptr [ebp-78], eax

' *** Reference to "__vbaFreeStr"
'006cfda5    ff150c134000            call dword ptr [0040130c]
'#Cleanup(var_39)
'006cfdab    8d4dcc                  lea ecx, dword ptr [ebp-34]

' *** Reference to "__vbaFreeObj"
'006cfdae    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_43)
'006cfdb4    668b4d88                mov cx, word ptr [ebp-78]
'006cfdb8    66394de8                cmp word ptr [ebp-18], cx
'006cfdbc    0f8fa8010000            jg 6cff6a

Do While (var_41 <= var_num1)
'006cfdc2    8b16                    mov edx, dword ptr [esi]
'006cfdc4    56                      push esi
'006cfdc5    ff9208030000            call dword ptr [edx+00000308]
'006cfdcb    50                      push eax
'006cfdcc    8d45cc                  lea eax, dword ptr [ebp-34]
'006cfdcf    50                      push eax
'006cfdd0    ffd7                    call edi
    Set var_43 = var_num1
'006cfdd2    8d55dc                  lea edx, dword ptr [ebp-24]
'006cfdd5    8bd8                    mov ebx, eax
'006cfdd7    8b0b                    mov ecx, dword ptr [ebx]
'006cfdd9    52                      push edx
'006cfdda    53                      push ebx
'006cfddb    ff91a8000000            call dword ptr [ecx+000000a8]
'006cfde1    dbe2                    fnclex
'006cfde3    85c0                    test ax, ax
'006cfde5    7d12                    jge 6cfdf9
'006cfde7    68a8000000              push 000000a8
'006cfdec    681cb94300              push 0043b91c
'006cfdf1    53                      push ebx
'006cfdf2    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006cfdf3    ff1580104000            call dword ptr [00401080]
'006cfdf9    8b45dc                  mov eax, dword ptr [ebp-24]
'006cfdfc    50                      push eax

' *** Reference to "rtcR8ValFromBstr"
'006cfdfd    ff1510134000            call dword ptr [00401310]
'006cfe03    dd5da0                  fstp qword ptr [ebp-60]
    'var_7 = (00)
'006cfe06    8d4db8                  lea ecx, dword ptr [ebp-48]
'006cfe09    51                      push ecx
'006cfe0a    c745c0e8030000          mov dword ptr [ebp-40], 000003e8
'006cfe11    c745b802000000          mov dword ptr [ebp-48], 00000002

' *** Reference to "rtcRandomNext"
'006cfe18    ff15a4104000            call dword ptr [004010a4]
'006cfe1e    dc4da0                  fmul qword ptr [ebp-60]
'006cfe21    dfe0                    fnstsw ax
'006cfe23    a80d                    test al, 0d
'006cfe25    0f85fe010000            jne 6d0029

' *** Reference to "__vbaFPInt"
'006cfe2b    ff15fc124000            call dword ptr [004012fc]
'006cfe31    dc0558164000            fadd qword ptr [00401658]
'006cfe37    dfe0                    fnstsw ax
'006cfe39    a80d                    test al, 0d
'006cfe3b    0f85e8010000            jne 6d0029

' *** Reference to "__vbaFpI2"
'006cfe41    ff15a0124000            call dword ptr [004012a0]
    var_num1 = CLng((Int((Rnd(var_5) * Val(vbNullString))) + 1#))
'006cfe47    8d4ddc                  lea ecx, dword ptr [ebp-24]
'006cfe4a    8bd8                    mov ebx, eax

' *** Reference to "__vbaFreeStr"
'006cfe4c    ff150c134000            call dword ptr [0040130c]
    '#Cleanup(var_39)
'006cfe52    8d4dcc                  lea ecx, dword ptr [ebp-34]

' *** Reference to "__vbaFreeObj"
'006cfe55    ff1508134000            call dword ptr [00401308]
    '#Cleanup(var_43)
'006cfe5b    8d4db8                  lea ecx, dword ptr [ebp-48]

' *** Reference to "__vbaFreeVar"
'006cfe5e    ff1528104000            call dword ptr [00401028]
    '#Cleanup(var_61)
'006cfe64    8b45e4                  mov eax, dword ptr [ebp-1c]
'006cfe67    0fbfd3                  movsx edx, bx
'006cfe6a    03d0                    add edx, eax
    var_num4 = var_num1 + Me
'006cfe6c    8b06                    mov eax, dword ptr [esi]
'006cfe6e    0f80ba010000            jo 6d002e
'006cfe74    56                      push esi
'006cfe75    8955e4                  mov dword ptr [ebp-1c], edx
'006cfe78    ff9000030000            call dword ptr [eax+00000300]
'006cfe7e    50                      push eax
'006cfe7f    8d4dc8                  lea ecx, dword ptr [ebp-38]
'006cfe82    51                      push ecx
'006cfe83    ffd7                    call edi
    Set var_46 = Nothing
'006cfe85    8b16                    mov edx, dword ptr [esi]
'006cfe87    56                      push esi
'006cfe88    894594                  mov dword ptr [ebp-6c], eax
'006cfe8b    ff9200030000            call dword ptr [edx+00000300]
'006cfe91    50                      push eax
'006cfe92    8d45cc                  lea eax, dword ptr [ebp-34]
'006cfe95    50                      push eax
'006cfe96    ffd7                    call edi
    Set var_43 = Nothing
'006cfe98    8d55dc                  lea edx, dword ptr [ebp-24]
'006cfe9b    8bf8                    mov edi, eax
'006cfe9d    8b0f                    mov ecx, dword ptr [edi]
'006cfe9f    52                      push edx
'006cfea0    57                      push edi
'006cfea1    ff91a0000000            call dword ptr [ecx+000000a0]
'006cfea7    dbe2                    fnclex
'006cfea9    85c0                    test ax, ax
'006cfeab    7d12                    jge 6cfebf
'006cfead    68a0000000              push 000000a0
'006cfeb2    68d00d4300              push 00430dd0
'006cfeb7    57                      push edi
'006cfeb8    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006cfeb9    ff1580104000            call dword ptr [00401080]
'006cfebf    8b4ddc                  mov ecx, dword ptr [ebp-24]
'006cfec2    8b4594                  mov eax, dword ptr [ebp-6c]
'006cfec5    8b38                    mov edi, dword ptr [eax]
'006cfec7    51                      push ecx
'006cfec8    53                      push ebx

' *** Reference to "__vbaStrI2"
'006cfec9    ff1510104000            call dword ptr [00401010]

' *** Reference to "__vbaStrMove"
'006cfecf    8b1dd0124000            mov ebx, dword ptr [004012d0]
'006cfed5    8bd0                    mov edx, eax
'006cfed7    8d4dd8                  lea ecx, dword ptr [ebp-28]
'006cfeda    ffd3                    call ebx
'006cfedc    50                      push eax

' *** Reference to "__vbaStrCat"
'006cfedd    ff1570104000            call dword ptr [00401070]
    var_36 = (vbNullString) & (CStr(var_num1))
'006cfee3    8bd0                    mov edx, eax
'006cfee5    8d4dd4                  lea ecx, dword ptr [ebp-2c]
'006cfee8    ffd3                    call ebx
'006cfeea    50                      push eax
'006cfeeb    68c48d4300              push 00438dc4

' *** Reference to "__vbaStrCat"
'006cfef0    ff1570104000            call dword ptr [00401070]
    var_76 = (var_36) & (" ")
'006cfef6    8bd0                    mov edx, eax
'006cfef8    8d4dd0                  lea ecx, dword ptr [ebp-30]
'006cfefb    ffd3                    call ebx
'006cfefd    8bd7                    mov edx, edi
'006cfeff    8b7d94                  mov edi, dword ptr [ebp-6c]
'006cff02    50                      push eax
'006cff03    57                      push edi
'006cff04    ff92a4000000            call dword ptr [edx+000000a4]
'006cff0a    dbe2                    fnclex
'006cff0c    85c0                    test ax, ax
'006cff0e    7d12                    jge 6cff22
'006cff10    68a4000000              push 000000a4
'006cff15    68d00d4300              push 00430dd0
'006cff1a    57                      push edi
'006cff1b    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006cff1c    ff1580104000            call dword ptr [00401080]
'006cff22    8d45d0                  lea eax, dword ptr [ebp-30]
'006cff25    50                      push eax
'006cff26    8d4dd4                  lea ecx, dword ptr [ebp-2c]
'006cff29    51                      push ecx
'006cff2a    8d55d8                  lea edx, dword ptr [ebp-28]
'006cff2d    52                      push edx
'006cff2e    8d45dc                  lea eax, dword ptr [ebp-24]
'006cff31    50                      push eax
'006cff32    6a04                    push 04

' *** Reference to "__vbaFreeStrList"
'006cff34    ff155c124000            call dword ptr [0040125c]
    '#Cleanup( , -4504, -4508, -4512)
'006cff3a    8d4dc8                  lea ecx, dword ptr [ebp-38]
'006cff3d    51                      push ecx
'006cff3e    8d55cc                  lea edx, dword ptr [ebp-34]
'006cff41    52                      push edx
'006cff42    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'006cff44    ff154c104000            call dword ptr [0040104c]
    '#Cleanup( var_43, var_46)

' *** Reference to "__vbaObjSet"
'006cff4a    8b3db4104000            mov edi, dword ptr [004010b4]
'006cff50    b801000000              mov eax, 00000001
'006cff55    83c420                  add esp, 20
'006cff58    660345e8                add ax, word ptr [ebp-18]
    var_num1 = 1 + var_41
'006cff5c    0f80cc000000            jo 6d002e
'006cff62    8945e8                  mov dword ptr [ebp-18], eax
'006cff65    e94afeffff              jmp 6cfdb4
    
Loop
'006cff6a    8b06                    mov eax, dword ptr [esi]
'006cff6c    56                      push esi
'006cff6d    ff90fc020000            call dword ptr [eax+000002fc]
'006cff73    50                      push eax
'006cff74    8d4dcc                  lea ecx, dword ptr [ebp-34]
'006cff77    51                      push ecx
'006cff78    ffd7                    call edi
Set var_43 = Nothing
'006cff7a    8b55e4                  mov edx, dword ptr [ebp-1c]
'006cff7d    8bf0                    mov esi, eax
'006cff7f    8b3e                    mov edi, dword ptr [esi]
'006cff81    52                      push edx

' *** Reference to "__vbaStrI4"
'006cff82    ff1520104000            call dword ptr [00401020]
'006cff88    8bd0                    mov edx, eax
'006cff8a    8d4ddc                  lea ecx, dword ptr [ebp-24]

' *** Reference to "__vbaStrMove"
'006cff8d    ff15d0124000            call dword ptr [004012d0]
'006cff93    50                      push eax
'006cff94    56                      push esi
'006cff95    ff97a4000000            call dword ptr [edi+000000a4]
'006cff9b    dbe2                    fnclex
'006cff9d    85c0                    test ax, ax
'006cff9f    7d12                    jge 6cffb3
'006cffa1    68a4000000              push 000000a4
'006cffa6    68d00d4300              push 00430dd0
'006cffab    56                      push esi
'006cffac    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006cffad    ff1580104000            call dword ptr [00401080]
'006cffb3    8d4ddc                  lea ecx, dword ptr [ebp-24]

' *** Reference to "__vbaFreeStr"
'006cffb6    ff150c134000            call dword ptr [0040130c]
'#Cleanup(var_39)
'006cffbc    8d4dcc                  lea ecx, dword ptr [ebp-34]

' *** Reference to "__vbaFreeObj"
'006cffbf    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_43)
'006cffc5    c745fc00000000          mov dword ptr [ebp-04], 00000000
'006cffcc    9b                      fwait
'006cffcd    680a006d00              push 006d000a
'006cffd2    eb35                    jmp 6d0009
'006cffd4    8d45d0                  lea eax, dword ptr [ebp-30]
'006cffd7    50                      push eax
'006cffd8    8d4dd4                  lea ecx, dword ptr [ebp-2c]
'006cffdb    51                      push ecx
'006cffdc    8d55d8                  lea edx, dword ptr [ebp-28]
'006cffdf    52                      push edx
'006cffe0    8d45dc                  lea eax, dword ptr [ebp-24]
'006cffe3    50                      push eax
'006cffe4    6a04                    push 04

' *** Reference to "__vbaFreeStrList"
'006cffe6    ff155c124000            call dword ptr [0040125c]
'#Cleanup( , -4504, -4508, -4512)
'006cffec    8d4dc8                  lea ecx, dword ptr [ebp-38]
'006cffef    51                      push ecx
'006cfff0    8d55cc                  lea edx, dword ptr [ebp-34]
'006cfff3    52                      push edx
'006cfff4    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'006cfff6    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_43, var_46)
'006cfffc    83c420                  add esp, 20
'006cffff    8d4db8                  lea ecx, dword ptr [ebp-48]

' *** Reference to "__vbaFreeVar"
'006d0002    ff1528104000            call dword ptr [00401028]
'#Cleanup(var_61)
'006d0008    c3                      ret
'006d0009    c3                      ret
'006d000a    8b4508                  mov eax, dword ptr [ebp+08]
'006d000d    8b08                    mov ecx, dword ptr [eax]
'006d000f    50                      push eax
'006d0010    ff5108                  call dword ptr [ecx+08]
'006d0013    8b45fc                  mov eax, dword ptr [ebp-04]
'006d0016    8b4dec                  mov ecx, dword ptr [ebp-14]
'006d0019    5f                      pop edi
'006d001a    5e                      pop esi
    'Reference to '__except_list'
'006d001b    64890d00000000          mov dword ptr fs:[00000000], ecx
'006d0022    5b                      pop ebx
'006d0023    8be5                    mov esp, ebp
'006d0025    5d                      pop ebp
'006d0026    c20400                  ret 0004


End Sub


'Event for Form
Private Sub Form_KeyDown(KeyCode as Integer, Shift as Integer)
'006d0040    55                      push ebp
'006d0041    8bec                    mov ebp, esp
'006d0043    83ec0c                  sub esp, 0c

' *** Reference to "__vbaExceptHandler"
'006d0046    6866724000              push 00407266
'006d004b    64a100000000            mov ax, word ptr fs:[00000000]
'006d0051    50                      push eax
    'Reference to '__except_list'
'006d0052    64892500000000          mov dword ptr fs:[00000000], esp
'006d0059    83ec18                  sub esp, 18
'006d005c    53                      push ebx
'006d005d    56                      push esi
'006d005e    57                      push edi
'006d005f    8965f4                  mov dword ptr [ebp-0c], esp
'006d0062    c745f8b8684000          mov dword ptr [ebp-08], 004068b8
'006d0069    8b7508                  mov esi, dword ptr [ebp+08]
'006d006c    8bc6                    mov eax, esi
'006d006e    83e001                  and eax, 01
'006d0071    8945fc                  mov dword ptr [ebp-04], eax
'006d0074    83e6fe                  and esi, fffffffe
'006d0077    8b0e                    mov ecx, dword ptr [esi]
'006d0079    56                      push esi
'006d007a    897508                  mov dword ptr [ebp+08], esi
'006d007d    ff5104                  call dword ptr [ecx+04]
'006d0080    8b550c                  mov edx, dword ptr [ebp+0c]
'006d0083    668b02                  mov ax, word ptr [edx]
'006d0086    33db                    xor ebx, ebx
'006d0088    663d0d00                cmp ax, 000d
'006d008c    895de8                  mov dword ptr [ebp-18], ebx
'006d008f    7521                    jne 6d00b2

If (arg_0 = 13) Then
'006d0091    8b06                    mov eax, dword ptr [esi]
'006d0093    56                      push esi
'006d0094    ff90f8060000            call dword ptr [eax+000006f8]
'006d009a    3bc3                    cmp eax, ebx
'006d009c    7d6e                    jge 6d010c
    
    If (    frmLanceurDes < 0) Then
'006d009e    68f8060000              push 000006f8
'006d00a3    68fc1a4300              push 00431afc
'006d00a8    56                      push esi
'006d00a9    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006d00aa    ff1580104000            call dword ptr [00401080]
'006d00b0    eb5a                    jmp 6d010c
    
End If
'006d00b2    663d1b00                cmp ax, 001b
'006d00b6    7554                    jne 6d010c

If (arg_0 = 27) Then
'006d00b8    391d24be7200            cmp dword ptr [0072be24], ebx
'006d00be    7510                    jne 6d00d0
'006d00c0    6824be7200              push 0072be24
'006d00c5    6870004300              push 00430070

' *** Reference to "__vbaNew2"
'006d00ca    ff152c124000            call dword ptr [0040122c]
    Dim var_2 As New Global
'006d00d0    8b3d24be7200            mov edi, dword ptr [0072be24]
'006d00d6    8b17                    mov edx, dword ptr [edi]
'006d00d8    56                      push esi
'006d00d9    8d4de8                  lea ecx, dword ptr [ebp-18]
'006d00dc    51                      push ecx
'006d00dd    8955d4                  mov dword ptr [ebp-2c], edx

' *** Reference to "__vbaObjSetAddref"
'006d00e0    ff15c8104000            call dword ptr [004010c8]
    Set var_41 = Me
'006d00e6    8b55d4                  mov edx, dword ptr [ebp-2c]
'006d00e9    50                      push eax
'006d00ea    57                      push edi
'006d00eb    ff5210                  call dword ptr [edx+10]
    Call var_2.Unload(var_41)
'006d00ee    dbe2                    fnclex
'006d00f0    3bc3                    cmp eax, ebx
'006d00f2    7d0f                    jge 6d0103
'006d00f4    6a10                    push 10
'006d00f6    6860004300              push 00430060
'006d00fb    57                      push edi
'006d00fc    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006d00fd    ff1580104000            call dword ptr [00401080]
'006d0103    8d4de8                  lea ecx, dword ptr [ebp-18]

' *** Reference to "__vbaFreeObj"
'006d0106    ff1508134000            call dword ptr [00401308]
    '#Cleanup(var_41)
End If
'006d010c    895dfc                  mov dword ptr [ebp-04], ebx
'006d010f    6821016d00              push 006d0121
'006d0114    eb0a                    jmp 6d0120
'006d0116    8d4de8                  lea ecx, dword ptr [ebp-18]

' *** Reference to "__vbaFreeObj"
'006d0119    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_41)
'006d011f    c3                      ret
'006d0120    c3                      ret
'006d0121    8b4508                  mov eax, dword ptr [ebp+08]
'006d0124    8b08                    mov ecx, dword ptr [eax]
'006d0126    50                      push eax
'006d0127    ff5108                  call dword ptr [ecx+08]
'006d012a    8b45fc                  mov eax, dword ptr [ebp-04]
'006d012d    8b4dec                  mov ecx, dword ptr [ebp-14]
'006d0130    5f                      pop edi
'006d0131    5e                      pop esi
    'Reference to '__except_list'
'006d0132    64890d00000000          mov dword ptr fs:[00000000], ecx
'006d0139    5b                      pop ebx
'006d013a    8be5                    mov esp, ebp
'006d013c    5d                      pop ebp
'006d013d    c20c00                  ret 000c


End Sub


