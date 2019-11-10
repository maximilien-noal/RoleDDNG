VERSION 5.00

Begin VB.Form frmDescriptionRaces
    Caption = "Description des races"
    ScaleMode = 1
    AutoRedraw = 0              'False
    FontTransparent = -1              'True
    LinkTopic = "Form1"
    MDIChild = -1              'True
    ClientLeft   = 60
    ClientTop    = 345
    ClientWidth  = 10695
    ClientHeight = 4230
    BeginProperty Font
        Name          = "Times New Roman"
        Size          = 9
        Charset       = 0
        Weight        = 400
        Underline     = 0              'False
        Italic        = 0              'False
        Strikethrough = 0              'False
    EndProperty
    Begin VB.CommandButton btnEnregistrer
        Caption = "&Enregistrer"
        Left   = 5880
        Top    = 120
        Width  = 1275
        Height = 285
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
    End
    Begin VB.ComboBox CmbRace
        Left   = 90
        Top    = 90
        Width  = 3615
        Height = 405
        Text = "Combo1&"
        TabIndex = 1
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
    Begin VB.TextBox fldstrDescriptionRaces
        Left   = 45
        Top    = 585
        Width  = 10650
        Height = 3660
        TabIndex = 0
        MultiLine = -1              'True
        ScrollBars = 2
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
'Event for btnEnregistrer
Private Sub btnEnregistrer_Click()
'006d0140    55                      push ebp
'006d0141    8bec                    mov ebp, esp
'006d0143    83ec0c                  sub esp, 0c

' *** Reference to "__vbaExceptHandler"
'006d0146    6866724000              push 00407266
'006d014b    64a100000000            mov ax, word ptr fs:[00000000]
'006d0151    50                      push eax
    'Reference to '__except_list'
'006d0152    64892500000000          mov dword ptr fs:[00000000], esp
'006d0159    81ecc4000000            sub esp, 000000c4
'006d015f    53                      push ebx
'006d0160    56                      push esi
'006d0161    57                      push edi
'006d0162    8965f4                  mov dword ptr [ebp-0c], esp
'006d0165    c745f8c8684000          mov dword ptr [ebp-08], 004068c8
'006d016c    8b5d08                  mov ebx, dword ptr [ebp+08]
'006d016f    8bc3                    mov eax, ebx
'006d0171    83e001                  and eax, 01
'006d0174    8945fc                  mov dword ptr [ebp-04], eax
'006d0177    83e3fe                  and ebx, fffffffe
'006d017a    8b0b                    mov ecx, dword ptr [ebx]
'006d017c    53                      push ebx
'006d017d    895d08                  mov dword ptr [ebp+08], ebx
'006d0180    ff5104                  call dword ptr [ecx+04]
'006d0183    8b13                    mov edx, dword ptr [ebx]
'006d0185    33ff                    xor edi, edi
'006d0187    53                      push ebx
'006d0188    897de8                  mov dword ptr [ebp-18], edi
'006d018b    897de4                  mov dword ptr [ebp-1c], edi
'006d018e    897de0                  mov dword ptr [ebp-20], edi
'006d0191    897ddc                  mov dword ptr [ebp-24], edi
'006d0194    897dd8                  mov dword ptr [ebp-28], edi
'006d0197    897dd4                  mov dword ptr [ebp-2c], edi
'006d019a    897dd0                  mov dword ptr [ebp-30], edi
'006d019d    897dcc                  mov dword ptr [ebp-34], edi
'006d01a0    897dc8                  mov dword ptr [ebp-38], edi
'006d01a3    897dc4                  mov dword ptr [ebp-3c], edi
'006d01a6    897dc0                  mov dword ptr [ebp-40], edi
'006d01a9    897dbc                  mov dword ptr [ebp-44], edi
'006d01ac    897db8                  mov dword ptr [ebp-48], edi
'006d01af    897db4                  mov dword ptr [ebp-4c], edi
'006d01b2    897da4                  mov dword ptr [ebp-5c], edi
'006d01b5    897d94                  mov dword ptr [ebp-6c], edi
'006d01b8    897d84                  mov dword ptr [ebp-7c], edi
'006d01bb    89bd74ffffff            mov dword ptr [ebp+ffffff74], edi
'006d01c1    89bd64ffffff            mov dword ptr [ebp+ffffff64], edi
'006d01c7    89bd54ffffff            mov dword ptr [ebp+ffffff54], edi
'006d01cd    ff9200030000            call dword ptr [edx+00000300]
'006d01d3    50                      push eax
'006d01d4    8d45b8                  lea eax, dword ptr [ebp-48]
'006d01d7    50                      push eax

' *** Reference to "__vbaObjSet"
'006d01d8    ff15b4104000            call dword ptr [004010b4]
Set var_61 = Me
'006d01de    8d55e8                  lea edx, dword ptr [ebp-18]
'006d01e1    8bf0                    mov esi, eax
'006d01e3    8b0e                    mov ecx, dword ptr [esi]
'006d01e5    52                      push edx
'006d01e6    56                      push esi
'006d01e7    ff91a8000000            call dword ptr [ecx+000000a8]
'006d01ed    dbe2                    fnclex
'006d01ef    3bc7                    cmp eax, edi
'006d01f1    7d12                    jge 6d0205

If (var_61 < 0) Then
'006d01f3    68a8000000              push 000000a8
'006d01f8    681cb94300              push 0043b91c
'006d01fd    56                      push esi
'006d01fe    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006d01ff    ff1580104000            call dword ptr [00401080]
    
End If
'006d0205    8b45e8                  mov eax, dword ptr [ebp-18]
'006d0208    50                      push eax
'006d0209    68cc134300              push 004313cc

' *** Reference to "__vbaStrCmp"
'006d020e    ff1530114000            call dword ptr [00401130]
'006d0214    8b0b                    mov ecx, dword ptr [ebx]
'006d0216    f7d8                    neg eax
'006d0218    1bc0                    sbb eax, eax
'006d021a    f7d8                    neg eax
'006d021c    f7d8                    neg eax
'006d021e    53                      push ebx
'006d021f    6689855cffffff          mov word ptr [ebp+ffffff5c], ax
'006d0226    c78554ffffff0b000000    mov dword ptr [ebp+ffffff54], 0000000b
'006d0230    ff9104030000            call dword ptr [ecx+00000304]
'006d0236    8d55a4                  lea edx, dword ptr [ebp-5c]
'006d0239    8945ac                  mov dword ptr [ebp-54], eax
'006d023c    52                      push edx
'006d023d    8d4594                  lea eax, dword ptr [ebp-6c]
'006d0240    50                      push eax
'006d0241    c745a409000000          mov dword ptr [ebp-5c], 00000009

' *** Reference to "rtcTrimVar"
'006d0248    ff15e4104000            call dword ptr [004010e4]
'006d024e    8d8d54ffffff            lea ecx, dword ptr [ebp+ffffff54]
'006d0254    51                      push ecx
'006d0255    8d5594                  lea edx, dword ptr [ebp-6c]
'006d0258    52                      push edx
'006d0259    8d8564ffffff            lea eax, dword ptr [ebp+ffffff64]
'006d025f    50                      push eax
'006d0260    8d4d84                  lea ecx, dword ptr [ebp-7c]
'006d0263    51                      push ecx
'006d0264    c7856cffffffcc134300    mov dword ptr [ebp+ffffff6c], 004313cc
'006d026e    c78564ffffff08800000    mov dword ptr [ebp+ffffff64], 00008008

' *** Reference to "__vbaVarCmpNe"
'006d0278    ff156c104000            call dword ptr [0040106c]
'006d027e    50                      push eax
'006d027f    8d9574ffffff            lea edx, dword ptr [ebp+ffffff74]
'006d0285    52                      push edx

' *** Reference to "__vbaVarAnd"
'006d0286    ff1598114000            call dword ptr [00401198]
'006d028c    50                      push eax

' *** Reference to "__vbaBoolVarNull"
'006d028d    ff15fc104000            call dword ptr [004010fc]
'006d0293    8d4de8                  lea ecx, dword ptr [ebp-18]
'006d0296    668bf0                  mov si, ax

' *** Reference to "__vbaFreeStr"
'006d0299    ff150c134000            call dword ptr [0040130c]
'#Cleanup(var_41)
'006d029f    8d4db8                  lea ecx, dword ptr [ebp-48]

' *** Reference to "__vbaFreeObj"
'006d02a2    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_61)
'006d02a8    8d8554ffffff            lea eax, dword ptr [ebp+ffffff54]
'006d02ae    50                      push eax
'006d02af    8d4d94                  lea ecx, dword ptr [ebp-6c]
'006d02b2    51                      push ecx
'006d02b3    8d55a4                  lea edx, dword ptr [ebp-5c]
'006d02b6    52                      push edx
'006d02b7    6a03                    push 03

' *** Reference to "__vbaFreeVarList"
'006d02b9    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_83, var_121, var_124)
'006d02bf    83c410                  add esp, 10
'006d02c2    663bf7                  cmp si, di
'006d02c5    0f84cb010000            je 6d0496

If (CBool(#NOT SUPPORTED#) <> 0) Then
'006d02cb    8b03                    mov eax, dword ptr [ebx]
'006d02cd    53                      push ebx
'006d02ce    ff9004030000            call dword ptr [eax+00000304]
'006d02d4    50                      push eax
'006d02d5    8d4db8                  lea ecx, dword ptr [ebp-48]
'006d02d8    51                      push ecx

' *** Reference to "__vbaObjSet"
'006d02d9    ff15b4104000            call dword ptr [004010b4]
    Set var_61 = Nothing
'006d02df    8bf0                    mov esi, eax
'006d02e1    8b16                    mov edx, dword ptr [esi]
'006d02e3    8d45e8                  lea eax, dword ptr [ebp-18]
'006d02e6    50                      push eax
'006d02e7    56                      push esi
'006d02e8    ff92a0000000            call dword ptr [edx+000000a0]
'006d02ee    dbe2                    fnclex
'006d02f0    3bc7                    cmp eax, edi
'006d02f2    7d12                    jge 6d0306
'006d02f4    68a0000000              push 000000a0
'006d02f9    68d00d4300              push 00430dd0
'006d02fe    56                      push esi
'006d02ff    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006d0300    ff1580104000            call dword ptr [00401080]
'006d0306    8b55e8                  mov edx, dword ptr [ebp-18]

' *** Reference to "__vbaStrMove"
'006d0309    8b35d0124000            mov esi, dword ptr [004012d0]
'006d030f    8d4de4                  lea ecx, dword ptr [ebp-1c]
'006d0312    897de8                  mov dword ptr [ebp-18], edi
'006d0315    ffd6                    call esi
'006d0317    8d4de4                  lea ecx, dword ptr [ebp-1c]
'006d031a    51                      push ecx
'006d031b    e8d038e2ff              call 4f3bf0
    Call sub_4F3BF0()
'006d0320    8bd0                    mov edx, eax
'006d0322    8d4dc0                  lea ecx, dword ptr [ebp-40]
'006d0325    ffd6                    call esi
'006d0327    8b13                    mov edx, dword ptr [ebx]
'006d0329    53                      push ebx
'006d032a    ff9200030000            call dword ptr [edx+00000300]
'006d0330    50                      push eax
'006d0331    8d45b4                  lea eax, dword ptr [ebp-4c]
'006d0334    50                      push eax

' *** Reference to "__vbaObjSet"
'006d0335    ff15b4104000            call dword ptr [004010b4]
    Set var_62 = Nothing
'006d033b    8d55d8                  lea edx, dword ptr [ebp-28]
'006d033e    8bd8                    mov ebx, eax
'006d0340    8b0b                    mov ecx, dword ptr [ebx]
'006d0342    52                      push edx
'006d0343    53                      push ebx
'006d0344    ff91a8000000            call dword ptr [ecx+000000a8]
'006d034a    dbe2                    fnclex
'006d034c    3bc7                    cmp eax, edi
'006d034e    7d12                    jge 6d0362
    
    If (    var_62 < 0) Then
'006d0350    68a8000000              push 000000a8
'006d0355    681cb94300              push 0043b91c
'006d035a    53                      push ebx
'006d035b    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006d035c    ff1580104000            call dword ptr [00401080]
    
End If
'006d0362    8b55d8                  mov edx, dword ptr [ebp-28]
'006d0365    8d4dd4                  lea ecx, dword ptr [ebp-2c]
'006d0368    897dd8                  mov dword ptr [ebp-28], edi
'006d036b    ffd6                    call esi
'006d036d    8d45d4                  lea eax, dword ptr [ebp-2c]
'006d0370    50                      push eax
'006d0371    e87a38e2ff              call 4f3bf0
Call sub_4F3BF0()
'006d0376    8bd0                    mov edx, eax
'006d0378    8d4dbc                  lea ecx, dword ptr [ebp-44]
'006d037b    ffd6                    call esi
'006d037d    8b55c0                  mov edx, dword ptr [ebp-40]
'006d0380    89952cffffff            mov dword ptr [ebp+ffffff2c], edx
'006d0386    8b55bc                  mov edx, dword ptr [ebp-44]
'006d0389    899528ffffff            mov dword ptr [ebp+ffffff28], edx
'006d038f    8b154cb07200            mov edx, dword ptr [0072b04c]
'006d0395    b880000000              mov eax, 00000080
'006d039a    b903000000              mov ecx, 00000003
'006d039f    898d64ffffff            mov dword ptr [ebp+ffffff64], ecx
'006d03a5    89856cffffff            mov dword ptr [ebp+ffffff6c], eax
'006d03ab    83ec10                  sub esp, 10
'006d03ae    897dc0                  mov dword ptr [ebp-40], edi
'006d03b1    897dbc                  mov dword ptr [ebp-44], edi
'006d03b4    8b1a                    mov ebx, dword ptr [edx]
'006d03b6    8bd4                    mov edx, esp
'006d03b8    890a                    mov dword ptr [edx], ecx
'006d03ba    8b8d68ffffff            mov ecx, dword ptr [ebp+ffffff68]
'006d03c0    894a04                  mov dword ptr [edx+04], ecx
'006d03c3    894208                  mov dword ptr [edx+08], eax
'006d03c6    8b8570ffffff            mov eax, dword ptr [ebp+ffffff70]
'006d03cc    89420c                  mov dword ptr [edx+0c], eax
'006d03cf    8b952cffffff            mov edx, dword ptr [ebp+ffffff2c]
'006d03d5    68083c4500              push 00453c08
'006d03da    8d4de0                  lea ecx, dword ptr [ebp-20]
'006d03dd    ffd6                    call esi
'006d03df    50                      push eax

' *** Reference to "__vbaStrCat"
'006d03e0    ff1570104000            call dword ptr [00401070]
var_49 = ("update race set Description='") & (vbNullString)
'006d03e6    8bd0                    mov edx, eax
'006d03e8    8d4ddc                  lea ecx, dword ptr [ebp-24]
'006d03eb    ffd6                    call esi
'006d03ed    50                      push eax
'006d03ee    6840354500              push 00453540

' *** Reference to "__vbaStrCat"
'006d03f3    ff1570104000            call dword ptr [00401070]
var_17 = (var_49) & ("' where race='")
'006d03f9    8bd0                    mov edx, eax
'006d03fb    8d4dd0                  lea ecx, dword ptr [ebp-30]
'006d03fe    ffd6                    call esi
'006d0400    8b9528ffffff            mov edx, dword ptr [ebp+ffffff28]
'006d0406    50                      push eax
'006d0407    8d4dcc                  lea ecx, dword ptr [ebp-34]
'006d040a    ffd6                    call esi
'006d040c    50                      push eax

' *** Reference to "__vbaStrCat"
'006d040d    ff1570104000            call dword ptr [00401070]
var_44 = (var_17) & (0)
'006d0413    8bd0                    mov edx, eax
'006d0415    8d4dc8                  lea ecx, dword ptr [ebp-38]
'006d0418    ffd6                    call esi
'006d041a    50                      push eax
'006d041b    6854a44300              push 0043a454

' *** Reference to "__vbaStrCat"
'006d0420    ff1570104000            call dword ptr [00401070]
var_53 = (var_44) & ("'")
'006d0426    8bd0                    mov edx, eax
'006d0428    8d4dc4                  lea ecx, dword ptr [ebp-3c]
'006d042b    ffd6                    call esi
'006d042d    8b0d4cb07200            mov ecx, dword ptr [0072b04c]
'006d0433    50                      push eax
'006d0434    51                      push ecx
'006d0435    ff535c                  call dword ptr [ebx+5c]
'006d0438    dbe2                    fnclex
'006d043a    3bc7                    cmp eax, edi
'006d043c    7d15                    jge 6d0453

If (-4552 < 0) Then
'006d043e    8b154cb07200            mov edx, dword ptr [0072b04c]
'006d0444    6a5c                    push 5c
'006d0446    68301f4300              push 00431f30
'006d044b    52                      push edx
'006d044c    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006d044d    ff1580104000            call dword ptr [00401080]
    
End If
'006d0453    8d45bc                  lea eax, dword ptr [ebp-44]
'006d0456    50                      push eax
'006d0457    8d4dc0                  lea ecx, dword ptr [ebp-40]
'006d045a    51                      push ecx
'006d045b    8d55c4                  lea edx, dword ptr [ebp-3c]
'006d045e    52                      push edx
'006d045f    8d45c8                  lea eax, dword ptr [ebp-38]
'006d0462    50                      push eax
'006d0463    8d4dcc                  lea ecx, dword ptr [ebp-34]
'006d0466    51                      push ecx
'006d0467    8d55d0                  lea edx, dword ptr [ebp-30]
'006d046a    52                      push edx
'006d046b    8d45d4                  lea eax, dword ptr [ebp-2c]
'006d046e    50                      push eax
'006d046f    8d4ddc                  lea ecx, dword ptr [ebp-24]
'006d0472    51                      push ecx
'006d0473    8d55e0                  lea edx, dword ptr [ebp-20]
'006d0476    52                      push edx
'006d0477    8d45e4                  lea eax, dword ptr [ebp-1c]
'006d047a    50                      push eax
'006d047b    6a0a                    push 0a

' *** Reference to "__vbaFreeStrList"
'006d047d    ff155c124000            call dword ptr [0040125c]
'#Cleanup( 0, 0, -4540, 0, -4544, -300, -4548, -4552, 0, 0)
'006d0483    8d4db4                  lea ecx, dword ptr [ebp-4c]
'006d0486    51                      push ecx
'006d0487    8d55b8                  lea edx, dword ptr [ebp-48]
'006d048a    52                      push edx
'006d048b    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'006d048d    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_61, var_62)
'006d0493    83c438                  add esp, 38

'ERROR: Two many next close:
End If
'006d0496    897dfc                  mov dword ptr [ebp-04], edi
'006d0499    680b056d00              push 006d050b
'006d049e    eb6a                    jmp 6d050a
'006d04a0    8d45bc                  lea eax, dword ptr [ebp-44]
'006d04a3    50                      push eax
'006d04a4    8d4dc0                  lea ecx, dword ptr [ebp-40]
'006d04a7    51                      push ecx
'006d04a8    8d55c4                  lea edx, dword ptr [ebp-3c]
'006d04ab    52                      push edx
'006d04ac    8d45c8                  lea eax, dword ptr [ebp-38]
'006d04af    50                      push eax
'006d04b0    8d4dcc                  lea ecx, dword ptr [ebp-34]
'006d04b3    51                      push ecx
'006d04b4    8d55d0                  lea edx, dword ptr [ebp-30]
'006d04b7    52                      push edx
'006d04b8    8d45d4                  lea eax, dword ptr [ebp-2c]
'006d04bb    50                      push eax
'006d04bc    8d4dd8                  lea ecx, dword ptr [ebp-28]
'006d04bf    51                      push ecx
'006d04c0    8d55dc                  lea edx, dword ptr [ebp-24]
'006d04c3    52                      push edx
'006d04c4    8d45e0                  lea eax, dword ptr [ebp-20]
'006d04c7    50                      push eax
'006d04c8    8d4de4                  lea ecx, dword ptr [ebp-1c]
'006d04cb    51                      push ecx
'006d04cc    8d55e8                  lea edx, dword ptr [ebp-18]
'006d04cf    52                      push edx
'006d04d0    6a0c                    push 0c

' *** Reference to "__vbaFreeStrList"
'006d04d2    ff155c124000            call dword ptr [0040125c]
'#Cleanup( 0, 0, 0, -4540, 0, 0, -4544, -300, -4548, -4552, 0, 0)
'006d04d8    8d45b4                  lea eax, dword ptr [ebp-4c]
'006d04db    50                      push eax
'006d04dc    8d4db8                  lea ecx, dword ptr [ebp-48]
'006d04df    51                      push ecx
'006d04e0    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'006d04e2    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_61, var_62)
'006d04e8    83c440                  add esp, 40
'006d04eb    8d9574ffffff            lea edx, dword ptr [ebp+ffffff74]
'006d04f1    52                      push edx
'006d04f2    8d4584                  lea eax, dword ptr [ebp-7c]
'006d04f5    50                      push eax
'006d04f6    8d4d94                  lea ecx, dword ptr [ebp-6c]
'006d04f9    51                      push ecx
'006d04fa    8d55a4                  lea edx, dword ptr [ebp-5c]
'006d04fd    52                      push edx
'006d04fe    6a04                    push 04

' *** Reference to "__vbaFreeVarList"
'006d0500    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_83, var_121, var_119, var_91)
'006d0506    83c414                  add esp, 14
'006d0509    c3                      ret
'006d050a    c3                      ret
'006d050b    8b4508                  mov eax, dword ptr [ebp+08]
'006d050e    8b08                    mov ecx, dword ptr [eax]
'006d0510    50                      push eax
'006d0511    ff5108                  call dword ptr [ecx+08]
'006d0514    8b45fc                  mov eax, dword ptr [ebp-04]
'006d0517    8b4dec                  mov ecx, dword ptr [ebp-14]
'006d051a    5f                      pop edi
'006d051b    5e                      pop esi
    'Reference to '__except_list'
'006d051c    64890d00000000          mov dword ptr fs:[00000000], ecx
'006d0523    5b                      pop ebx
'006d0524    8be5                    mov esp, ebp
'006d0526    5d                      pop ebp
'006d0527    c20400                  ret 0004


End Sub


'Event for CmbRace
Private Sub CmbRace_Click()
'006d0530    55                      push ebp
'006d0531    8bec                    mov ebp, esp
'006d0533    83ec0c                  sub esp, 0c

' *** Reference to "__vbaExceptHandler"
'006d0536    6866724000              push 00407266
'006d053b    64a100000000            mov ax, word ptr fs:[00000000]
'006d0541    50                      push eax
    'Reference to '__except_list'
'006d0542    64892500000000          mov dword ptr fs:[00000000], esp
'006d0549    81ecf0000000            sub esp, 000000f0
'006d054f    53                      push ebx
'006d0550    56                      push esi
'006d0551    57                      push edi
'006d0552    8965f4                  mov dword ptr [ebp-0c], esp
'006d0555    c745f8d8684000          mov dword ptr [ebp-08], 004068d8
'006d055c    8b7508                  mov esi, dword ptr [ebp+08]
'006d055f    8bc6                    mov eax, esi
'006d0561    83e001                  and eax, 01
'006d0564    8945fc                  mov dword ptr [ebp-04], eax
'006d0567    83e6fe                  and esi, fffffffe
'006d056a    8b0e                    mov ecx, dword ptr [esi]
'006d056c    56                      push esi
'006d056d    897508                  mov dword ptr [ebp+08], esi
'006d0570    ff5104                  call dword ptr [ecx+04]
'006d0573    8b16                    mov edx, dword ptr [esi]
'006d0575    33ff                    xor edi, edi
'006d0577    56                      push esi
'006d0578    897de8                  mov dword ptr [ebp-18], edi
'006d057b    897de4                  mov dword ptr [ebp-1c], edi
'006d057e    897de0                  mov dword ptr [ebp-20], edi
'006d0581    897ddc                  mov dword ptr [ebp-24], edi
'006d0584    897dd8                  mov dword ptr [ebp-28], edi
'006d0587    897dd4                  mov dword ptr [ebp-2c], edi
'006d058a    897dd0                  mov dword ptr [ebp-30], edi
'006d058d    897dcc                  mov dword ptr [ebp-34], edi
'006d0590    897dc8                  mov dword ptr [ebp-38], edi
'006d0593    897dc4                  mov dword ptr [ebp-3c], edi
'006d0596    897dc0                  mov dword ptr [ebp-40], edi
'006d0599    897dbc                  mov dword ptr [ebp-44], edi
'006d059c    897dac                  mov dword ptr [ebp-54], edi
'006d059f    897d9c                  mov dword ptr [ebp-64], edi
'006d05a2    897d8c                  mov dword ptr [ebp-74], edi
'006d05a5    89bd7cffffff            mov dword ptr [ebp+ffffff7c], edi
'006d05ab    89bd4cffffff            mov dword ptr [ebp+ffffff4c], edi
'006d05b1    89bd3cffffff            mov dword ptr [ebp+ffffff3c], edi
'006d05b7    ff9200030000            call dword ptr [edx+00000300]
'006d05bd    50                      push eax
'006d05be    8d45cc                  lea eax, dword ptr [ebp-34]
'006d05c1    50                      push eax

' *** Reference to "__vbaObjSet"
'006d05c2    ff15b4104000            call dword ptr [004010b4]
Set var_43 = Me
'006d05c8    8d55e4                  lea edx, dword ptr [ebp-1c]
'006d05cb    8bf0                    mov esi, eax
'006d05cd    8b0e                    mov ecx, dword ptr [esi]
'006d05cf    52                      push edx
'006d05d0    56                      push esi
'006d05d1    ff91a8000000            call dword ptr [ecx+000000a8]
'006d05d7    dbe2                    fnclex
'006d05d9    3bc7                    cmp eax, edi
'006d05db    7d12                    jge 6d05ef

If (var_43 < 0) Then
'006d05dd    68a8000000              push 000000a8
'006d05e2    681cb94300              push 0043b91c
'006d05e7    56                      push esi
'006d05e8    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006d05e9    ff1580104000            call dword ptr [00401080]
    
End If
'006d05ef    8b55e4                  mov edx, dword ptr [ebp-1c]

' *** Reference to "__vbaStrMove"
'006d05f2    8b35d0124000            mov esi, dword ptr [004012d0]
'006d05f8    8d4de0                  lea ecx, dword ptr [ebp-20]
'006d05fb    897de4                  mov dword ptr [ebp-1c], edi
'006d05fe    ffd6                    call esi
'006d0600    8d45e0                  lea eax, dword ptr [ebp-20]
'006d0603    50                      push eax
'006d0604    e8e735e2ff              call 4f3bf0
Call sub_4F3BF0()
'006d0609    8bd0                    mov edx, eax
'006d060b    8d4dd0                  lea ecx, dword ptr [ebp-30]
'006d060e    ffd6                    call esi
'006d0610    8b55d0                  mov edx, dword ptr [ebp-30]
'006d0613    899508ffffff            mov dword ptr [ebp+ffffff08], edx
'006d0619    8b154cb07200            mov edx, dword ptr [0072b04c]
'006d061f    b804000280              mov eax, 80020004
'006d0624    898554ffffff            mov dword ptr [ebp+ffffff54], eax
'006d062a    b90a000000              mov ecx, 0000000a
'006d062f    898d4cffffff            mov dword ptr [ebp+ffffff4c], ecx
'006d0635    898d5cffffff            mov dword ptr [ebp+ffffff5c], ecx
'006d063b    897dd0                  mov dword ptr [ebp-30], edi
'006d063e    8b1a                    mov ebx, dword ptr [edx]
'006d0640    8d55c8                  lea edx, dword ptr [ebp-38]
'006d0643    52                      push edx
'006d0644    83ec10                  sub esp, 10
'006d0647    8bd4                    mov edx, esp
'006d0649    890a                    mov dword ptr [edx], ecx
'006d064b    8b8d50ffffff            mov ecx, dword ptr [ebp+ffffff50]
'006d0651    894a04                  mov dword ptr [edx+04], ecx
'006d0654    894208                  mov dword ptr [edx+08], eax
'006d0657    898564ffffff            mov dword ptr [ebp+ffffff64], eax
'006d065d    8b8558ffffff            mov eax, dword ptr [ebp+ffffff58]
'006d0663    89420c                  mov dword ptr [edx+0c], eax
'006d0666    8b955cffffff            mov edx, dword ptr [ebp+ffffff5c]
'006d066c    8b8560ffffff            mov eax, dword ptr [ebp+ffffff60]
'006d0672    83ec10                  sub esp, 10
'006d0675    8bcc                    mov ecx, esp
'006d0677    8911                    mov dword ptr [ecx], edx
'006d0679    8b9564ffffff            mov edx, dword ptr [ebp+ffffff64]
'006d067f    894104                  mov dword ptr [ecx+04], eax
'006d0682    8b8568ffffff            mov eax, dword ptr [ebp+ffffff68]
'006d0688    895108                  mov dword ptr [ecx+08], edx
'006d068b    8b9570ffffff            mov edx, dword ptr [ebp+ffffff70]
'006d0691    89410c                  mov dword ptr [ecx+0c], eax
'006d0694    83ec10                  sub esp, 10
'006d0697    8bcc                    mov ecx, esp
'006d0699    b803000000              mov eax, 00000003
'006d069e    8901                    mov dword ptr [ecx], eax
'006d06a0    895104                  mov dword ptr [ecx+04], edx
'006d06a3    8b9508ffffff            mov edx, dword ptr [ebp+ffffff08]
'006d06a9    b802000000              mov eax, 00000002
'006d06ae    894108                  mov dword ptr [ecx+08], eax
'006d06b1    8b8578ffffff            mov eax, dword ptr [ebp+ffffff78]
'006d06b7    89410c                  mov dword ptr [ecx+0c], eax
'006d06ba    68583c4500              push 00453c58
'006d06bf    8d4ddc                  lea ecx, dword ptr [ebp-24]
'006d06c2    ffd6                    call esi
'006d06c4    50                      push eax

' *** Reference to "__vbaStrCat"
'006d06c5    ff1570104000            call dword ptr [00401070]
var_3 = ("select * from race where race='") & (0)
'006d06cb    8bd0                    mov edx, eax
'006d06cd    8d4dd8                  lea ecx, dword ptr [ebp-28]
'006d06d0    ffd6                    call esi
'006d06d2    50                      push eax
'006d06d3    6854a44300              push 0043a454

' *** Reference to "__vbaStrCat"
'006d06d8    ff1570104000            call dword ptr [00401070]
var_84 = (var_3) & ("'")
'006d06de    8bd0                    mov edx, eax
'006d06e0    8d4dd4                  lea ecx, dword ptr [ebp-2c]
'006d06e3    ffd6                    call esi
'006d06e5    8b0d4cb07200            mov ecx, dword ptr [0072b04c]
'006d06eb    50                      push eax
'006d06ec    51                      push ecx
'006d06ed    ff93bc000000            call dword ptr [ebx+000000bc]
'006d06f3    dbe2                    fnclex
'006d06f5    3bc7                    cmp eax, edi
'006d06f7    7d18                    jge 6d0711

If (-4500 < 0) Then
'006d06f9    8b154cb07200            mov edx, dword ptr [0072b04c]
'006d06ff    68bc000000              push 000000bc
'006d0704    68301f4300              push 00431f30
'006d0709    52                      push edx
'006d070a    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006d070b    ff1580104000            call dword ptr [00401080]
    
End If
'006d0711    8b45c8                  mov eax, dword ptr [ebp-38]
'006d0714    50                      push eax
'006d0715    8d45e8                  lea eax, dword ptr [ebp-18]
'006d0718    50                      push eax
'006d0719    897dc8                  mov dword ptr [ebp-38], edi

' *** Reference to "__vbaObjSet"
'006d071c    ff15b4104000            call dword ptr [004010b4]
Set var_41 = Nothing
'006d0722    8d4dd0                  lea ecx, dword ptr [ebp-30]
'006d0725    51                      push ecx
'006d0726    8d55d4                  lea edx, dword ptr [ebp-2c]
'006d0729    52                      push edx
'006d072a    8d45d8                  lea eax, dword ptr [ebp-28]
'006d072d    50                      push eax
'006d072e    8d4ddc                  lea ecx, dword ptr [ebp-24]
'006d0731    51                      push ecx
'006d0732    8d55e0                  lea edx, dword ptr [ebp-20]
'006d0735    52                      push edx
'006d0736    6a05                    push 05

' *** Reference to "__vbaFreeStrList"
'006d0738    ff155c124000            call dword ptr [0040125c]
'#Cleanup( 0, -288, -4496, -4500, 0)
'006d073e    83c418                  add esp, 18
'006d0741    8d4dcc                  lea ecx, dword ptr [ebp-34]

' *** Reference to "__vbaFreeObj"
'006d0744    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_43)
'006d074a    8b45e8                  mov eax, dword ptr [ebp-18]
'006d074d    8b08                    mov ecx, dword ptr [eax]
'006d074f    8d55cc                  lea edx, dword ptr [ebp-34]
'006d0752    52                      push edx
'006d0753    50                      push eax
'006d0754    ff91b4000000            call dword ptr [ecx+000000b4]
'006d075a    dbe2                    fnclex
'006d075c    3bc7                    cmp eax, edi
'006d075e    7d15                    jge 6d0775

If (var_41 < 0) Then
'006d0760    8b4de8                  mov ecx, dword ptr [ebp-18]
'006d0763    68b4000000              push 000000b4
'006d0768    6830314300              push 00433130
'006d076d    51                      push ecx
'006d076e    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006d076f    ff1580104000            call dword ptr [00401080]
    
End If
'006d0775    8b45cc                  mov eax, dword ptr [ebp-34]
'006d0778    8b30                    mov esi, dword ptr [eax]
'006d077a    8d5dc8                  lea ebx, dword ptr [ebp-38]
'006d077d    53                      push ebx
'006d077e    83ec10                  sub esp, 10
'006d0781    8bdc                    mov ebx, esp
'006d0783    ba08000000              mov edx, 00000008
'006d0788    8913                    mov dword ptr [ebx], edx
'006d078a    8b9570ffffff            mov edx, dword ptr [ebp+ffffff70]
'006d0790    895304                  mov dword ptr [ebx+04], edx
'006d0793    b9cca64300              mov ecx, 0043a6cc
'006d0798    894b08                  mov dword ptr [ebx+08], ecx
'006d079b    8b8d78ffffff            mov ecx, dword ptr [ebp+ffffff78]
'006d07a1    50                      push eax
'006d07a2    898530ffffff            mov dword ptr [ebp+ffffff30], eax
'006d07a8    894b0c                  mov dword ptr [ebx+0c], ecx
'006d07ab    ff5630                  call dword ptr [esi+30]
'006d07ae    dbe2                    fnclex
'006d07b0    3bc7                    cmp eax, edi
'006d07b2    7d15                    jge 6d07c9

If (var_43 < 0) Then
'006d07b4    8b9530ffffff            mov edx, dword ptr [ebp+ffffff30]
'006d07ba    6a30                    push 30
'006d07bc    68d8304300              push 004330d8
'006d07c1    52                      push edx
'006d07c2    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006d07c3    ff1580104000            call dword ptr [00401080]
    
End If
'006d07c9    8b45c8                  mov eax, dword ptr [ebp-38]
'006d07cc    8945b4                  mov dword ptr [ebp-4c], eax
'006d07cf    8d45ac                  lea eax, dword ptr [ebp-54]
'006d07d2    50                      push eax
'006d07d3    897dc8                  mov dword ptr [ebp-38], edi
'006d07d6    c745ac09000000          mov dword ptr [ebp-54], 00000009

' *** Reference to "rtcIsNull"
'006d07dd    ff1540114000            call dword ptr [00401140]
'006d07e3    898538ffffff            mov dword ptr [ebp+ffffff38], eax
'006d07e9    8b4508                  mov eax, dword ptr [ebp+08]
'006d07ec    8b08                    mov ecx, dword ptr [eax]
'006d07ee    50                      push eax
'006d07ef    ff9104030000            call dword ptr [ecx+00000304]
'006d07f5    50                      push eax
'006d07f6    8d55bc                  lea edx, dword ptr [ebp-44]
'006d07f9    52                      push edx

' *** Reference to "__vbaObjSet"
'006d07fa    ff15b4104000            call dword ptr [004010b4]
Set var_58 = Me
'006d0800    8d55c4                  lea edx, dword ptr [ebp-3c]
'006d0803    89851cffffff            mov dword ptr [ebp+ffffff1c], eax
'006d0809    8b45e8                  mov eax, dword ptr [ebp-18]
'006d080c    8b08                    mov ecx, dword ptr [eax]
'006d080e    52                      push edx
'006d080f    50                      push eax
'006d0810    ff91b4000000            call dword ptr [ecx+000000b4]
'006d0816    dbe2                    fnclex
'006d0818    3bc7                    cmp eax, edi
'006d081a    7d15                    jge 6d0831

If (var_41 < 0) Then
'006d081c    8b4de8                  mov ecx, dword ptr [ebp-18]
'006d081f    68b4000000              push 000000b4
'006d0824    6830314300              push 00433130
'006d0829    51                      push ecx
'006d082a    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006d082b    ff1580104000            call dword ptr [00401080]
    
End If
'006d0831    8b45c4                  mov eax, dword ptr [ebp-3c]
'006d0834    8b30                    mov esi, dword ptr [eax]
'006d0836    8d5dc0                  lea ebx, dword ptr [ebp-40]
'006d0839    53                      push ebx
'006d083a    83ec10                  sub esp, 10
'006d083d    8bdc                    mov ebx, esp
'006d083f    ba08000000              mov edx, 00000008
'006d0844    8913                    mov dword ptr [ebx], edx
'006d0846    8b9560ffffff            mov edx, dword ptr [ebp+ffffff60]
'006d084c    895304                  mov dword ptr [ebx+04], edx
'006d084f    b9cca64300              mov ecx, 0043a6cc
'006d0854    894b08                  mov dword ptr [ebx+08], ecx
'006d0857    8b8d68ffffff            mov ecx, dword ptr [ebp+ffffff68]
'006d085d    50                      push eax
'006d085e    898524ffffff            mov dword ptr [ebp+ffffff24], eax
'006d0864    894b0c                  mov dword ptr [ebx+0c], ecx
'006d0867    ff5630                  call dword ptr [esi+30]
'006d086a    dbe2                    fnclex
'006d086c    3bc7                    cmp eax, edi
'006d086e    7d15                    jge 6d0885

If (0 < 0) Then
'006d0870    8b9524ffffff            mov edx, dword ptr [ebp+ffffff24]
'006d0876    6a30                    push 30
'006d0878    68d8304300              push 004330d8
'006d087d    52                      push edx
'006d087e    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006d087f    ff1580104000            call dword ptr [00401080]
    
End If
'006d0885    8b45c0                  mov eax, dword ptr [ebp-40]
'006d0888    8d953cffffff            lea edx, dword ptr [ebp+ffffff3c]
'006d088e    8d4d9c                  lea ecx, dword ptr [ebp-64]
'006d0891    897dc0                  mov dword ptr [ebp-40], edi
'006d0894    894594                  mov dword ptr [ebp-6c], eax
'006d0897    c7458c09000000          mov dword ptr [ebp-74], 00000009
'006d089e    c78544ffffffcc134300    mov dword ptr [ebp+ffffff44], 004313cc
'006d08a8    c7853cffffff08000000    mov dword ptr [ebp+ffffff3c], 00000008

' *** Reference to "__vbaVarDup"
'006d08b2    ff1598124000            call dword ptr [00401298]
var_51 = (vbNullChar)
'006d08b8    668b8538ffffff          mov ax, word ptr [ebp+ffffff38]
'006d08bf    8d4d8c                  lea ecx, dword ptr [ebp-74]
'006d08c2    51                      push ecx
'006d08c3    8d559c                  lea edx, dword ptr [ebp-64]
'006d08c6    66898554ffffff          mov word ptr [ebp+ffffff54], ax
'006d08cd    52                      push edx
'006d08ce    8d854cffffff            lea eax, dword ptr [ebp+ffffff4c]
'006d08d4    50                      push eax
'006d08d5    8d8d7cffffff            lea ecx, dword ptr [ebp+ffffff7c]
'006d08db    51                      push ecx
'006d08dc    c7854cffffff0b000000    mov dword ptr [ebp+ffffff4c], 0000000b

' *** Reference to "rtcImmediateIf"
'006d08e6    ff154c124000            call dword ptr [0040124c]
'006d08ec    8bb51cffffff            mov esi, dword ptr [ebp+ffffff1c]
'006d08f2    8b1e                    mov ebx, dword ptr [esi]
'006d08f4    8d957cffffff            lea edx, dword ptr [ebp+ffffff7c]
'006d08fa    52                      push edx
'006d08fb    8d45e4                  lea eax, dword ptr [ebp-1c]
'006d08fe    50                      push eax

' *** Reference to "__vbaStrVarVal"
'006d08ff    ff15fc114000            call dword ptr [004011fc]
'006d0905    50                      push eax
'006d0906    56                      push esi
'006d0907    ff93a4000000            call dword ptr [ebx+000000a4]
'006d090d    dbe2                    fnclex
'006d090f    3bc7                    cmp eax, edi
'006d0911    7d12                    jge 6d0925
'006d0913    68a4000000              push 000000a4
'006d0918    68d00d4300              push 00430dd0
'006d091d    56                      push esi
'006d091e    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006d091f    ff1580104000            call dword ptr [00401080]
'006d0925    8d4de4                  lea ecx, dword ptr [ebp-1c]

' *** Reference to "__vbaFreeStr"
'006d0928    ff150c134000            call dword ptr [0040130c]
'#Cleanup(var_40)
'006d092e    8d4dbc                  lea ecx, dword ptr [ebp-44]
'006d0931    51                      push ecx
'006d0932    8d55c4                  lea edx, dword ptr [ebp-3c]
'006d0935    52                      push edx
'006d0936    8d45cc                  lea eax, dword ptr [ebp-34]
'006d0939    50                      push eax
'006d093a    6a03                    push 03

' *** Reference to "__vbaFreeObjList"
'006d093c    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_43, var_9, var_58)
'006d0942    8d8d7cffffff            lea ecx, dword ptr [ebp+ffffff7c]
'006d0948    51                      push ecx
'006d0949    8d558c                  lea edx, dword ptr [ebp-74]
'006d094c    52                      push edx
'006d094d    8d459c                  lea eax, dword ptr [ebp-64]
'006d0950    50                      push eax
'006d0951    8d8d4cffffff            lea ecx, dword ptr [ebp+ffffff4c]
'006d0957    51                      push ecx
'006d0958    8d55ac                  lea edx, dword ptr [ebp-54]
'006d095b    52                      push edx
'006d095c    6a05                    push 05

' *** Reference to "__vbaFreeVarList"
'006d095e    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_50, var_89, var_51, var_52, var_59)
'006d0964    8b45e8                  mov eax, dword ptr [ebp-18]
'006d0967    8b08                    mov ecx, dword ptr [eax]
'006d0969    83c428                  add esp, 28
'006d096c    50                      push eax
'006d096d    ff91c4000000            call dword ptr [ecx+000000c4]
'006d0973    dbe2                    fnclex
'006d0975    3bc7                    cmp eax, edi
'006d0977    7d15                    jge 6d098e

If (var_41 < 0) Then
'006d0979    8b55e8                  mov edx, dword ptr [ebp-18]
'006d097c    68c4000000              push 000000c4
'006d0981    6830314300              push 00433130
'006d0986    52                      push edx
'006d0987    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006d0988    ff1580104000            call dword ptr [00401080]
    
End If
'006d098e    897dfc                  mov dword ptr [ebp-04], edi
'006d0991    68fd096d00              push 006d09fd
'006d0996    eb5b                    jmp 6d09f3
'006d0998    8d45d0                  lea eax, dword ptr [ebp-30]
'006d099b    50                      push eax
'006d099c    8d4dd4                  lea ecx, dword ptr [ebp-2c]
'006d099f    51                      push ecx
'006d09a0    8d55d8                  lea edx, dword ptr [ebp-28]
'006d09a3    52                      push edx
'006d09a4    8d45dc                  lea eax, dword ptr [ebp-24]
'006d09a7    50                      push eax
'006d09a8    8d4de0                  lea ecx, dword ptr [ebp-20]
'006d09ab    51                      push ecx
'006d09ac    8d55e4                  lea edx, dword ptr [ebp-1c]
'006d09af    52                      push edx
'006d09b0    6a06                    push 06

' *** Reference to "__vbaFreeStrList"
'006d09b2    ff155c124000            call dword ptr [0040125c]
'#Cleanup( , 0, -288, -4496, -4500, 0)
'006d09b8    8d45bc                  lea eax, dword ptr [ebp-44]
'006d09bb    50                      push eax
'006d09bc    8d4dc0                  lea ecx, dword ptr [ebp-40]
'006d09bf    51                      push ecx
'006d09c0    8d55c4                  lea edx, dword ptr [ebp-3c]
'006d09c3    52                      push edx
'006d09c4    8d45c8                  lea eax, dword ptr [ebp-38]
'006d09c7    50                      push eax
'006d09c8    8d4dcc                  lea ecx, dword ptr [ebp-34]
'006d09cb    51                      push ecx
'006d09cc    6a05                    push 05

' *** Reference to "__vbaFreeObjList"
'006d09ce    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_43, var_46, var_9, var_5, var_58)
'006d09d4    8d957cffffff            lea edx, dword ptr [ebp+ffffff7c]
'006d09da    52                      push edx
'006d09db    8d458c                  lea eax, dword ptr [ebp-74]
'006d09de    50                      push eax
'006d09df    8d4d9c                  lea ecx, dword ptr [ebp-64]
'006d09e2    51                      push ecx
'006d09e3    8d55ac                  lea edx, dword ptr [ebp-54]
'006d09e6    52                      push edx
'006d09e7    6a04                    push 04

' *** Reference to "__vbaFreeVarList"
'006d09e9    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_50, var_51, var_52, var_59)
'006d09ef    83c448                  add esp, 48
'006d09f2    c3                      ret
'006d09f3    8d4de8                  lea ecx, dword ptr [ebp-18]

' *** Reference to "__vbaFreeObj"
'006d09f6    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_41)
'006d09fc    c3                      ret
'006d09fd    8b4508                  mov eax, dword ptr [ebp+08]
'006d0a00    8b08                    mov ecx, dword ptr [eax]
'006d0a02    50                      push eax
'006d0a03    ff5108                  call dword ptr [ecx+08]
'006d0a06    8b45fc                  mov eax, dword ptr [ebp-04]
'006d0a09    8b4dec                  mov ecx, dword ptr [ebp-14]
'006d0a0c    5f                      pop edi
'006d0a0d    5e                      pop esi
    'Reference to '__except_list'
'006d0a0e    64890d00000000          mov dword ptr fs:[00000000], ecx
'006d0a15    5b                      pop ebx
'006d0a16    8be5                    mov esp, ebp
'006d0a18    5d                      pop ebp
'006d0a19    c20400                  ret 0004


End Sub


'Event for Form
Private Sub Form_Load()
'006d0a20    55                      push ebp
'006d0a21    8bec                    mov ebp, esp
'006d0a23    83ec0c                  sub esp, 0c

' *** Reference to "__vbaExceptHandler"
'006d0a26    6866724000              push 00407266
'006d0a2b    64a100000000            mov ax, word ptr fs:[00000000]
'006d0a31    50                      push eax
    'Reference to '__except_list'
'006d0a32    64892500000000          mov dword ptr fs:[00000000], esp
'006d0a39    81ec00010000            sub esp, 00000100
'006d0a3f    53                      push ebx
'006d0a40    56                      push esi
'006d0a41    57                      push edi
'006d0a42    8965f4                  mov dword ptr [ebp-0c], esp
'006d0a45    c745f8f0684000          mov dword ptr [ebp-08], 004068f0
'006d0a4c    8b7508                  mov esi, dword ptr [ebp+08]
'006d0a4f    8bc6                    mov eax, esi
'006d0a51    83e001                  and eax, 01
'006d0a54    8945fc                  mov dword ptr [ebp-04], eax
'006d0a57    83e6fe                  and esi, fffffffe
'006d0a5a    8b0e                    mov ecx, dword ptr [esi]
'006d0a5c    56                      push esi
'006d0a5d    897508                  mov dword ptr [ebp+08], esi
'006d0a60    ff5104                  call dword ptr [ecx+04]
'006d0a63    8b16                    mov edx, dword ptr [esi]
'006d0a65    33ff                    xor edi, edi
'006d0a67    56                      push esi
'006d0a68    897de8                  mov dword ptr [ebp-18], edi
'006d0a6b    897de4                  mov dword ptr [ebp-1c], edi
'006d0a6e    897de0                  mov dword ptr [ebp-20], edi
'006d0a71    897ddc                  mov dword ptr [ebp-24], edi
'006d0a74    897dd8                  mov dword ptr [ebp-28], edi
'006d0a77    897dd4                  mov dword ptr [ebp-2c], edi
'006d0a7a    897dd0                  mov dword ptr [ebp-30], edi
'006d0a7d    897dcc                  mov dword ptr [ebp-34], edi
'006d0a80    897dc8                  mov dword ptr [ebp-38], edi
'006d0a83    897dc4                  mov dword ptr [ebp-3c], edi
'006d0a86    897dc0                  mov dword ptr [ebp-40], edi
'006d0a89    897dbc                  mov dword ptr [ebp-44], edi
'006d0a8c    897dac                  mov dword ptr [ebp-54], edi
'006d0a8f    897d9c                  mov dword ptr [ebp-64], edi
'006d0a92    897d8c                  mov dword ptr [ebp-74], edi
'006d0a95    89bd7cffffff            mov dword ptr [ebp+ffffff7c], edi
'006d0a9b    89bd4cffffff            mov dword ptr [ebp+ffffff4c], edi
'006d0aa1    89bd3cffffff            mov dword ptr [ebp+ffffff3c], edi
'006d0aa7    89bd38ffffff            mov dword ptr [ebp+ffffff38], edi
'006d0aad    89bd34ffffff            mov dword ptr [ebp+ffffff34], edi
'006d0ab3    ff9204030000            call dword ptr [edx+00000304]
'006d0ab9    50                      push eax
'006d0aba    8d45cc                  lea eax, dword ptr [ebp-34]
'006d0abd    50                      push eax

' *** Reference to "__vbaObjSet"
'006d0abe    ff15b4104000            call dword ptr [004010b4]
Set var_43 = Me
'006d0ac4    898528ffffff            mov dword ptr [ebp+ffffff28], eax
'006d0aca    393dd4b07200            cmp dword ptr [0072b0d4], edi
'006d0ad0    7510                    jne 6d0ae2
'006d0ad2    68d4b07200              push 0072b0d4
'006d0ad7    68207e4000              push 00407e20

' *** Reference to "__vbaNew2"
'006d0adc    ff152c124000            call dword ptr [0040122c]
Dim var_32 As New frmDescriptionRaces
'006d0ae2    8b1dd4b07200            mov ebx, dword ptr [0072b0d4]
'006d0ae8    8b0b                    mov ecx, dword ptr [ebx]
'006d0aea    8d9534ffffff            lea edx, dword ptr [ebp+ffffff34]
'006d0af0    52                      push edx
'006d0af1    53                      push ebx
'006d0af2    ff9188000000            call dword ptr [ecx+00000088]
'006d0af8    dbe2                    fnclex
'006d0afa    3bc7                    cmp eax, edi
'006d0afc    7d12                    jge 6d0b10

If (var_43 < 0) Then
'006d0afe    6888000000              push 00000088
'006d0b03    6854134300              push 00431354
'006d0b08    53                      push ebx
'006d0b09    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006d0b0a    ff1580104000            call dword ptr [00401080]
    
End If
'006d0b10    d98534ffffff            fld dword ptr [ebp+ffffff34]
'006d0b16    8b9d28ffffff            mov ebx, dword ptr [ebp+ffffff28]
'006d0b1c    d825ec684000            fsub dword ptr [004068ec]
'006d0b22    8b0b                    mov ecx, dword ptr [ebx]
'006d0b24    51                      push ecx
'006d0b25    dfe0                    fnstsw ax
'006d0b27    a80d                    test al, 0d
'006d0b29    0f85240a0000            jne 6d1553
'006d0b2f    d91c24                  fstp dword ptr [esp]
'var_25 = (00)
'006d0b32    53                      push ebx
'006d0b33    ff9184000000            call dword ptr [ecx+00000084]
'006d0b39    dbe2                    fnclex
'006d0b3b    3bc7                    cmp eax, edi
'006d0b3d    7d12                    jge 6d0b51
'006d0b3f    6884000000              push 00000084
'006d0b44    68d00d4300              push 00430dd0
'006d0b49    53                      push ebx
'006d0b4a    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006d0b4b    ff1580104000            call dword ptr [00401080]
'006d0b51    8d4dcc                  lea ecx, dword ptr [ebp-34]

' *** Reference to "__vbaFreeObj"
'006d0b54    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_43)
'006d0b5a    8b16                    mov edx, dword ptr [esi]
'006d0b5c    56                      push esi
'006d0b5d    ff9204030000            call dword ptr [edx+00000304]
'006d0b63    50                      push eax
'006d0b64    8d45cc                  lea eax, dword ptr [ebp-34]
'006d0b67    50                      push eax

' *** Reference to "__vbaObjSet"
'006d0b68    ff15b4104000            call dword ptr [004010b4]
Set var_43 = 
'006d0b6e    898528ffffff            mov dword ptr [ebp+ffffff28], eax
'006d0b74    393dd4b07200            cmp dword ptr [0072b0d4], edi
'006d0b7a    7510                    jne 6d0b8c

If (var_32 = 0) Then
'006d0b7c    68d4b07200              push 0072b0d4
'006d0b81    68207e4000              push 00407e20

' *** Reference to "__vbaNew2"
'006d0b86    ff152c124000            call dword ptr [0040122c]
    Set var_32 = New frmDescriptionRaces
End If
'006d0b8c    8b1dd4b07200            mov ebx, dword ptr [0072b0d4]
'006d0b92    8b0b                    mov ecx, dword ptr [ebx]
'006d0b94    8d9534ffffff            lea edx, dword ptr [ebp+ffffff34]
'006d0b9a    52                      push edx
'006d0b9b    53                      push ebx
'006d0b9c    ff9180000000            call dword ptr [ecx+00000080]
'006d0ba2    dbe2                    fnclex
'006d0ba4    3bc7                    cmp eax, edi
'006d0ba6    7d12                    jge 6d0bba

If (var_43 < 0) Then
'006d0ba8    6880000000              push 00000080
'006d0bad    6854134300              push 00431354
'006d0bb2    53                      push ebx
'006d0bb3    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006d0bb4    ff1580104000            call dword ptr [00401080]
    
End If
'006d0bba    d98534ffffff            fld dword ptr [ebp+ffffff34]
'006d0bc0    8b9d28ffffff            mov ebx, dword ptr [ebp+ffffff28]
'006d0bc6    d825e8684000            fsub dword ptr [004068e8]
'006d0bcc    8b0b                    mov ecx, dword ptr [ebx]
'006d0bce    51                      push ecx
'006d0bcf    dfe0                    fnstsw ax
'006d0bd1    a80d                    test al, 0d
'006d0bd3    0f857a090000            jne 6d1553
'006d0bd9    d91c24                  fstp dword ptr [esp]
'var_2085 = (00)
'006d0bdc    53                      push ebx
'006d0bdd    ff517c                  call dword ptr [ecx+7c]
'006d0be0    dbe2                    fnclex
'006d0be2    3bc7                    cmp eax, edi
'006d0be4    7d0f                    jge 6d0bf5
'006d0be6    6a7c                    push 7c
'006d0be8    68d00d4300              push 00430dd0
'006d0bed    53                      push ebx
'006d0bee    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006d0bef    ff1580104000            call dword ptr [00401080]
'006d0bf5    8d4dcc                  lea ecx, dword ptr [ebp-34]

' *** Reference to "__vbaFreeObj"
'006d0bf8    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_43)
'006d0bfe    8b16                    mov edx, dword ptr [esi]
'006d0c00    56                      push esi
'006d0c01    ff9200030000            call dword ptr [edx+00000300]
'006d0c07    50                      push eax
'006d0c08    8d45cc                  lea eax, dword ptr [ebp-34]
'006d0c0b    50                      push eax

' *** Reference to "__vbaObjSet"
'006d0c0c    ff15b4104000            call dword ptr [004010b4]
Set var_43 = 
'006d0c12    8bd8                    mov ebx, eax
'006d0c14    8b0b                    mov ecx, dword ptr [ebx]
'006d0c16    53                      push ebx
'006d0c17    ff91e8010000            call dword ptr [ecx+000001e8]
'006d0c1d    dbe2                    fnclex
'006d0c1f    3bc7                    cmp eax, edi
'006d0c21    7d12                    jge 6d0c35

If (var_43 < 0) Then
'006d0c23    68e8010000              push 000001e8
'006d0c28    681cb94300              push 0043b91c
'006d0c2d    53                      push ebx
'006d0c2e    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006d0c2f    ff1580104000            call dword ptr [00401080]
    
End If
'006d0c35    8d4dcc                  lea ecx, dword ptr [ebp-34]

' *** Reference to "__vbaFreeObj"
'006d0c38    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_43)
'006d0c3e    8d5dcc                  lea ebx, dword ptr [ebp-34]
'006d0c41    53                      push ebx
'006d0c42    83ec10                  sub esp, 10
'006d0c45    8bdc                    mov ebx, esp
'006d0c47    b90a000000              mov ecx, 0000000a
'006d0c4c    890b                    mov dword ptr [ebx], ecx
'006d0c4e    898d4cffffff            mov dword ptr [ebp+ffffff4c], ecx
'006d0c54    8b8d50ffffff            mov ecx, dword ptr [ebp+ffffff50]
'006d0c5a    894b04                  mov dword ptr [ebx+04], ecx
'006d0c5d    8b3d4cb07200            mov edi, dword ptr [0072b04c]
'006d0c63    b804000280              mov eax, 80020004
'006d0c68    894308                  mov dword ptr [ebx+08], eax
'006d0c6b    8bd0                    mov edx, eax
'006d0c6d    898554ffffff            mov dword ptr [ebp+ffffff54], eax
'006d0c73    8b8558ffffff            mov eax, dword ptr [ebp+ffffff58]
'006d0c79    89430c                  mov dword ptr [ebx+0c], eax
'006d0c7c    8b3f                    mov edi, dword ptr [edi]
'006d0c7e    83ec10                  sub esp, 10
'006d0c81    8bcc                    mov ecx, esp
'006d0c83    b80a000000              mov eax, 0000000a
'006d0c88    8901                    mov dword ptr [ecx], eax
'006d0c8a    8b8560ffffff            mov eax, dword ptr [ebp+ffffff60]
'006d0c90    894104                  mov dword ptr [ecx+04], eax
'006d0c93    895108                  mov dword ptr [ecx+08], edx
'006d0c96    8b9568ffffff            mov edx, dword ptr [ebp+ffffff68]
'006d0c9c    89510c                  mov dword ptr [ecx+0c], edx
'006d0c9f    8b9570ffffff            mov edx, dword ptr [ebp+ffffff70]
'006d0ca5    83ec10                  sub esp, 10
'006d0ca8    8bcc                    mov ecx, esp
'006d0caa    b803000000              mov eax, 00000003
'006d0caf    8901                    mov dword ptr [ecx], eax
'006d0cb1    895104                  mov dword ptr [ecx+04], edx
'006d0cb4    b802000000              mov eax, 00000002
'006d0cb9    894108                  mov dword ptr [ecx+08], eax
'006d0cbc    8b8578ffffff            mov eax, dword ptr [ebp+ffffff78]
'006d0cc2    89410c                  mov dword ptr [ecx+0c], eax
'006d0cc5    8b0d4cb07200            mov ecx, dword ptr [0072b04c]
'006d0ccb    68c8044500              push 004504c8
'006d0cd0    51                      push ecx
'006d0cd1    ff97bc000000            call dword ptr [edi+000000bc]
'006d0cd7    dbe2                    fnclex
'006d0cd9    85c0                    test ax, ax
'006d0cdb    7d1c                    jge 6d0cf9
'006d0cdd    8b154cb07200            mov edx, dword ptr [0072b04c]

' *** Reference to "__vbaHresultCheckObj"
'006d0ce3    8b1d80104000            mov ebx, dword ptr [00401080]
'006d0ce9    68bc000000              push 000000bc
'006d0cee    68301f4300              push 00431f30
'006d0cf3    52                      push edx
'006d0cf4    50                      push eax
'006d0cf5    ffd3                    call ebx
'006d0cf7    eb06                    jmp 6d0cff

' *** Reference to "__vbaHresultCheckObj"
'006d0cf9    8b1d80104000            mov ebx, dword ptr [00401080]
'006d0cff    8b45cc                  mov eax, dword ptr [ebp-34]
'006d0d02    50                      push eax
'006d0d03    8d45e8                  lea eax, dword ptr [ebp-18]
'006d0d06    50                      push eax
'006d0d07    c745cc00000000          mov dword ptr [ebp-34], 00000000

' *** Reference to "__vbaObjSet"
'006d0d0e    ff15b4104000            call dword ptr [004010b4]
Set var_41 = var_43
'006d0d14    8b45e8                  mov eax, dword ptr [ebp-18]
'006d0d17    8b08                    mov ecx, dword ptr [eax]
'006d0d19    8d9538ffffff            lea edx, dword ptr [ebp+ffffff38]
'006d0d1f    52                      push edx
'006d0d20    50                      push eax
'006d0d21    ff5134                  call dword ptr [ecx+34]
'006d0d24    dbe2                    fnclex
'006d0d26    85c0                    test ax, ax
'006d0d28    7d0e                    jge 6d0d38
'006d0d2a    8b4de8                  mov ecx, dword ptr [ebp-18]
'006d0d2d    6a34                    push 34
'006d0d2f    6830314300              push 00433130
'006d0d34    51                      push ecx
'006d0d35    50                      push eax
'006d0d36    ffd3                    call ebx
'006d0d38    6683bd38ffffff00        cmp word ptr [ebp+ffffff38], 00
'006d0d40    0f85f0020000            jne 6d1036

Do While (0 = 0)
'006d0d46    8b16                    mov edx, dword ptr [esi]
'006d0d48    56                      push esi
'006d0d49    ff9200030000            call dword ptr [edx+00000300]
'006d0d4f    50                      push eax
'006d0d50    8d45c4                  lea eax, dword ptr [ebp-3c]
'006d0d53    50                      push eax

' *** Reference to "__vbaObjSet"
'006d0d54    ff15b4104000            call dword ptr [004010b4]
    Set var_9 = var_41
'006d0d5a    8d55cc                  lea edx, dword ptr [ebp-34]
'006d0d5d    89851cffffff            mov dword ptr [ebp+ffffff1c], eax
'006d0d63    8b45e8                  mov eax, dword ptr [ebp-18]
'006d0d66    8b08                    mov ecx, dword ptr [eax]
'006d0d68    52                      push edx
'006d0d69    50                      push eax
'006d0d6a    c78564ffffff04000280    mov dword ptr [ebp+ffffff64], 80020004
'006d0d74    c7855cffffff0a000000    mov dword ptr [ebp+ffffff5c], 0000000a
'006d0d7e    ff91b4000000            call dword ptr [ecx+000000b4]
'006d0d84    dbe2                    fnclex
'006d0d86    85c0                    test ax, ax
'006d0d88    7d11                    jge 6d0d9b
'006d0d8a    8b4de8                  mov ecx, dword ptr [ebp-18]
'006d0d8d    68b4000000              push 000000b4
'006d0d92    6830314300              push 00433130
'006d0d97    51                      push ecx
'006d0d98    50                      push eax
'006d0d99    ffd3                    call ebx
'006d0d9b    8b45cc                  mov eax, dword ptr [ebp-34]
'006d0d9e    8b38                    mov edi, dword ptr [eax]
'006d0da0    8d5dc8                  lea ebx, dword ptr [ebp-38]
'006d0da3    53                      push ebx
'006d0da4    83ec10                  sub esp, 10
'006d0da7    8bdc                    mov ebx, esp
'006d0da9    ba08000000              mov edx, 00000008
'006d0dae    8913                    mov dword ptr [ebx], edx
'006d0db0    8b9570ffffff            mov edx, dword ptr [ebp+ffffff70]
'006d0db6    895304                  mov dword ptr [ebx+04], edx
'006d0db9    b9e4b14300              mov ecx, 0043b1e4
'006d0dbe    894b08                  mov dword ptr [ebx+08], ecx
'006d0dc1    8b8d78ffffff            mov ecx, dword ptr [ebp+ffffff78]
'006d0dc7    50                      push eax
'006d0dc8    89852cffffff            mov dword ptr [ebp+ffffff2c], eax
'006d0dce    894b0c                  mov dword ptr [ebx+0c], ecx
'006d0dd1    ff5730                  call dword ptr [edi+30]
'006d0dd4    dbe2                    fnclex
'006d0dd6    85c0                    test ax, ax
'006d0dd8    7d15                    jge 6d0def
'006d0dda    8b952cffffff            mov edx, dword ptr [ebp+ffffff2c]
'006d0de0    6a30                    push 30
'006d0de2    68d8304300              push 004330d8
'006d0de7    52                      push edx
'006d0de8    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006d0de9    ff1580104000            call dword ptr [00401080]
'006d0def    8b45c8                  mov eax, dword ptr [ebp-38]
'006d0df2    8b08                    mov ecx, dword ptr [eax]
'006d0df4    8d55ac                  lea edx, dword ptr [ebp-54]
'006d0df7    52                      push edx
'006d0df8    50                      push eax
'006d0df9    8bf8                    mov edi, eax
'006d0dfb    ff5144                  call dword ptr [ecx+44]
'006d0dfe    dbe2                    fnclex
'006d0e00    85c0                    test ax, ax
'006d0e02    7d0f                    jge 6d0e13
'006d0e04    6a44                    push 44
'006d0e06    6880304300              push 00433080
'006d0e0b    57                      push edi
'006d0e0c    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006d0e0d    ff1580104000            call dword ptr [00401080]
'006d0e13    8b8d5cffffff            mov ecx, dword ptr [ebp+ffffff5c]
'006d0e19    8b9560ffffff            mov edx, dword ptr [ebp+ffffff60]
'006d0e1f    8bbd1cffffff            mov edi, dword ptr [ebp+ffffff1c]
'006d0e25    8b1f                    mov ebx, dword ptr [edi]
'006d0e27    83ec10                  sub esp, 10
'006d0e2a    8bc4                    mov eax, esp
'006d0e2c    8908                    mov dword ptr [eax], ecx
'006d0e2e    8b8d64ffffff            mov ecx, dword ptr [ebp+ffffff64]
'006d0e34    895004                  mov dword ptr [eax+04], edx
'006d0e37    8b9568ffffff            mov edx, dword ptr [ebp+ffffff68]
'006d0e3d    894808                  mov dword ptr [eax+08], ecx
'006d0e40    89500c                  mov dword ptr [eax+0c], edx
'006d0e43    8d45ac                  lea eax, dword ptr [ebp-54]
'006d0e46    50                      push eax

' *** Reference to "__vbaStrVarMove"
'006d0e47    ff153c104000            call dword ptr [0040103c]
'006d0e4d    8bd0                    mov edx, eax
'006d0e4f    8d4de4                  lea ecx, dword ptr [ebp-1c]

' *** Reference to "__vbaStrMove"
'006d0e52    ff15d0124000            call dword ptr [004012d0]
'006d0e58    50                      push eax
'006d0e59    57                      push edi
'006d0e5a    ff93ec010000            call dword ptr [ebx+000001ec]
'006d0e60    dbe2                    fnclex
'006d0e62    85c0                    test ax, ax
'006d0e64    7d12                    jge 6d0e78
'006d0e66    68ec010000              push 000001ec
'006d0e6b    681cb94300              push 0043b91c
'006d0e70    57                      push edi
'006d0e71    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006d0e72    ff1580104000            call dword ptr [00401080]
'006d0e78    8d4de4                  lea ecx, dword ptr [ebp-1c]

' *** Reference to "__vbaFreeStr"
'006d0e7b    ff150c134000            call dword ptr [0040130c]
    '#Cleanup(var_40)
'006d0e81    8d4dc4                  lea ecx, dword ptr [ebp-3c]
'006d0e84    51                      push ecx
'006d0e85    8d55c8                  lea edx, dword ptr [ebp-38]
'006d0e88    52                      push edx
'006d0e89    8d45cc                  lea eax, dword ptr [ebp-34]
'006d0e8c    50                      push eax
'006d0e8d    6a03                    push 03

' *** Reference to "__vbaFreeObjList"
'006d0e8f    ff154c104000            call dword ptr [0040104c]
    '#Cleanup( var_43, var_46, var_9)
'006d0e95    83c410                  add esp, 10
'006d0e98    8d4dac                  lea ecx, dword ptr [ebp-54]

' *** Reference to "__vbaFreeVar"
'006d0e9b    ff1528104000            call dword ptr [00401028]
    '#Cleanup(var_50)
'006d0ea1    8b0e                    mov ecx, dword ptr [esi]
'006d0ea3    56                      push esi
'006d0ea4    ff9100030000            call dword ptr [ecx+00000300]
'006d0eaa    50                      push eax
'006d0eab    8d55c0                  lea edx, dword ptr [ebp-40]
'006d0eae    52                      push edx

' *** Reference to "__vbaObjSet"
'006d0eaf    ff15b4104000            call dword ptr [004010b4]
    Set var_5 = 
'006d0eb5    8d55cc                  lea edx, dword ptr [ebp-34]
'006d0eb8    898514ffffff            mov dword ptr [ebp+ffffff14], eax
'006d0ebe    8b45e8                  mov eax, dword ptr [ebp-18]
'006d0ec1    8b08                    mov ecx, dword ptr [eax]
'006d0ec3    52                      push edx
'006d0ec4    50                      push eax
'006d0ec5    ff91b4000000            call dword ptr [ecx+000000b4]
'006d0ecb    dbe2                    fnclex
'006d0ecd    85c0                    test ax, ax
'006d0ecf    7d15                    jge 6d0ee6
'006d0ed1    8b4de8                  mov ecx, dword ptr [ebp-18]
'006d0ed4    68b4000000              push 000000b4
'006d0ed9    6830314300              push 00433130
'006d0ede    51                      push ecx
'006d0edf    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006d0ee0    ff1580104000            call dword ptr [00401080]
'006d0ee6    8b45cc                  mov eax, dword ptr [ebp-34]
'006d0ee9    8b38                    mov edi, dword ptr [eax]
'006d0eeb    8d5dc8                  lea ebx, dword ptr [ebp-38]
'006d0eee    53                      push ebx
'006d0eef    83ec10                  sub esp, 10
'006d0ef2    8bdc                    mov ebx, esp
'006d0ef4    ba08000000              mov edx, 00000008
'006d0ef9    8913                    mov dword ptr [ebx], edx
'006d0efb    8b9570ffffff            mov edx, dword ptr [ebp+ffffff70]
'006d0f01    895304                  mov dword ptr [ebx+04], edx
'006d0f04    b9cc284400              mov ecx, 004428cc
'006d0f09    894b08                  mov dword ptr [ebx+08], ecx
'006d0f0c    8b8d78ffffff            mov ecx, dword ptr [ebp+ffffff78]
'006d0f12    50                      push eax
'006d0f13    898524ffffff            mov dword ptr [ebp+ffffff24], eax
'006d0f19    894b0c                  mov dword ptr [ebx+0c], ecx
'006d0f1c    ff5730                  call dword ptr [edi+30]
'006d0f1f    dbe2                    fnclex
'006d0f21    85c0                    test ax, ax
'006d0f23    7d19                    jge 6d0f3e
'006d0f25    8b9524ffffff            mov edx, dword ptr [ebp+ffffff24]

' *** Reference to "__vbaHresultCheckObj"
'006d0f2b    8b1d80104000            mov ebx, dword ptr [00401080]
'006d0f31    6a30                    push 30
'006d0f33    68d8304300              push 004330d8
'006d0f38    52                      push edx
'006d0f39    50                      push eax
'006d0f3a    ffd3                    call ebx
'006d0f3c    eb06                    jmp 6d0f44

' *** Reference to "__vbaHresultCheckObj"
'006d0f3e    8b1d80104000            mov ebx, dword ptr [00401080]
'006d0f44    8b45c8                  mov eax, dword ptr [ebp-38]
'006d0f47    8b08                    mov ecx, dword ptr [eax]
'006d0f49    8d55ac                  lea edx, dword ptr [ebp-54]
'006d0f4c    52                      push edx
'006d0f4d    50                      push eax
'006d0f4e    8bf8                    mov edi, eax
'006d0f50    ff5144                  call dword ptr [ecx+44]
'006d0f53    dbe2                    fnclex
'006d0f55    85c0                    test ax, ax
'006d0f57    7d0b                    jge 6d0f64
'006d0f59    6a44                    push 44
'006d0f5b    6880304300              push 00433080
'006d0f60    57                      push edi
'006d0f61    50                      push eax
'006d0f62    ffd3                    call ebx
'006d0f64    8b06                    mov eax, dword ptr [esi]
'006d0f66    56                      push esi
'006d0f67    ff9000030000            call dword ptr [eax+00000300]
'006d0f6d    50                      push eax
'006d0f6e    8d4dc4                  lea ecx, dword ptr [ebp-3c]
'006d0f71    51                      push ecx

' *** Reference to "__vbaObjSet"
'006d0f72    ff15b4104000            call dword ptr [004010b4]
    Set var_9 = Nothing
'006d0f78    8bf8                    mov edi, eax
'006d0f7a    8b17                    mov edx, dword ptr [edi]
'006d0f7c    8d8538ffffff            lea eax, dword ptr [ebp+ffffff38]
'006d0f82    50                      push eax
'006d0f83    57                      push edi
'006d0f84    ff92e8000000            call dword ptr [edx+000000e8]
'006d0f8a    dbe2                    fnclex
'006d0f8c    85c0                    test ax, ax
'006d0f8e    7d0e                    jge 6d0f9e
'006d0f90    68e8000000              push 000000e8
'006d0f95    681cb94300              push 0043b91c
'006d0f9a    57                      push edi
'006d0f9b    50                      push eax
'006d0f9c    ffd3                    call ebx
'006d0f9e    8b8d14ffffff            mov ecx, dword ptr [ebp+ffffff14]
'006d0fa4    8b39                    mov edi, dword ptr [ecx]
'006d0fa6    8d55ac                  lea edx, dword ptr [ebp-54]
'006d0fa9    52                      push edx

' *** Reference to "__vbaI4Var"
'006d0faa    ff157c124000            call dword ptr [0040127c]
'006d0fb0    50                      push eax
'006d0fb1    668b8538ffffff          mov ax, word ptr [ebp+ffffff38]
'006d0fb8    662d0100                sub ax, 0001
    var_num1 = 0 - 1
'006d0fbc    0f8096050000            jo 6d1558
'006d0fc2    8bcf                    mov ecx, edi
'006d0fc4    8bbd14ffffff            mov edi, dword ptr [ebp+ffffff14]
'006d0fca    50                      push eax
'006d0fcb    57                      push edi
'006d0fcc    ff9154010000            call dword ptr [ecx+00000154]
'006d0fd2    dbe2                    fnclex
'006d0fd4    85c0                    test ax, ax
'006d0fd6    7d0e                    jge 6d0fe6
'006d0fd8    6854010000              push 00000154
'006d0fdd    681cb94300              push 0043b91c
'006d0fe2    57                      push edi
'006d0fe3    50                      push eax
'006d0fe4    ffd3                    call ebx
'006d0fe6    8d55c0                  lea edx, dword ptr [ebp-40]
'006d0fe9    52                      push edx
'006d0fea    8d45c8                  lea eax, dword ptr [ebp-38]
'006d0fed    50                      push eax
'006d0fee    8d4dc4                  lea ecx, dword ptr [ebp-3c]
'006d0ff1    51                      push ecx
'006d0ff2    8d55cc                  lea edx, dword ptr [ebp-34]
'006d0ff5    52                      push edx
'006d0ff6    6a04                    push 04

' *** Reference to "__vbaFreeObjList"
'006d0ff8    ff154c104000            call dword ptr [0040104c]
    '#Cleanup( var_43, var_9, var_46, var_5)
'006d0ffe    83c414                  add esp, 14
'006d1001    8d4dac                  lea ecx, dword ptr [ebp-54]

' *** Reference to "__vbaFreeVar"
'006d1004    ff1528104000            call dword ptr [00401028]
    '#Cleanup(var_50)
'006d100a    8b45e8                  mov eax, dword ptr [ebp-18]
'006d100d    8b08                    mov ecx, dword ptr [eax]
'006d100f    50                      push eax
'006d1010    ff91ec000000            call dword ptr [ecx+000000ec]
'006d1016    dbe2                    fnclex
'006d1018    85c0                    test ax, ax
'006d101a    0f8df4fcffff            jge 6d0d14
'006d1020    8b55e8                  mov edx, dword ptr [ebp-18]
'006d1023    68ec000000              push 000000ec
'006d1028    6830314300              push 00433130
'006d102d    52                      push edx
'006d102e    50                      push eax
'006d102f    ffd3                    call ebx
'006d1031    e9defcffff              jmp 6d0d14
    
Loop
'006d1036    8b45e8                  mov eax, dword ptr [ebp-18]
'006d1039    8b08                    mov ecx, dword ptr [eax]
'006d103b    50                      push eax
'006d103c    ff91c4000000            call dword ptr [ecx+000000c4]
'006d1042    dbe2                    fnclex
'006d1044    85c0                    test ax, ax
'006d1046    7d11                    jge 6d1059
'006d1048    8b55e8                  mov edx, dword ptr [ebp-18]
'006d104b    68c4000000              push 000000c4
'006d1050    6830314300              push 00433130
'006d1055    52                      push edx
'006d1056    50                      push eax
'006d1057    ffd3                    call ebx
'006d1059    8b06                    mov eax, dword ptr [esi]
'006d105b    56                      push esi
'006d105c    ff9000030000            call dword ptr [eax+00000300]
'006d1062    50                      push eax
'006d1063    8d4dcc                  lea ecx, dword ptr [ebp-34]
'006d1066    51                      push ecx

' *** Reference to "__vbaObjSet"
'006d1067    ff15b4104000            call dword ptr [004010b4]
Set var_43 = Nothing
'006d106d    8bf8                    mov edi, eax
'006d106f    8b17                    mov edx, dword ptr [edi]
'006d1071    6a00                    push 00
'006d1073    57                      push edi
'006d1074    ff92f4000000            call dword ptr [edx+000000f4]
'006d107a    dbe2                    fnclex
'006d107c    85c0                    test ax, ax
'006d107e    7d0e                    jge 6d108e
'006d1080    68f4000000              push 000000f4
'006d1085    681cb94300              push 0043b91c
'006d108a    57                      push edi
'006d108b    50                      push eax
'006d108c    ffd3                    call ebx
'006d108e    8d4dcc                  lea ecx, dword ptr [ebp-34]

' *** Reference to "__vbaFreeObj"
'006d1091    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_43)
'006d1097    8b4634                  mov eax, dword ptr [esi+34]
'006d109a    50                      push eax
'006d109b    68cc134300              push 004313cc

' *** Reference to "__vbaStrCmp"
'006d10a0    ff1530114000            call dword ptr [00401130]
'006d10a6    85c0                    test ax, ax
'006d10a8    0f8412040000            je 6d14c0

If (((vbNullString) <> (vbNullChar))) Then
'006d10ae    8b0e                    mov ecx, dword ptr [esi]
'006d10b0    56                      push esi
'006d10b1    ff9100030000            call dword ptr [ecx+00000300]
'006d10b7    50                      push eax
'006d10b8    8d55cc                  lea edx, dword ptr [ebp-34]
'006d10bb    52                      push edx

' *** Reference to "__vbaObjSet"
'006d10bc    ff15b4104000            call dword ptr [004010b4]
    Set var_43 = ((vbNullString) [##] (vbNullChar))
'006d10c2    8b4e34                  mov ecx, dword ptr [esi+34]
'006d10c5    8bf8                    mov edi, eax
'006d10c7    8b07                    mov eax, dword ptr [edi]
'006d10c9    51                      push ecx
'006d10ca    57                      push edi
'006d10cb    ff90ac000000            call dword ptr [eax+000000ac]
'006d10d1    dbe2                    fnclex
'006d10d3    85c0                    test ax, ax
'006d10d5    7d0e                    jge 6d10e5
'006d10d7    68ac000000              push 000000ac
'006d10dc    681cb94300              push 0043b91c
'006d10e1    57                      push edi
'006d10e2    50                      push eax
'006d10e3    ffd3                    call ebx
'006d10e5    8d4dcc                  lea ecx, dword ptr [ebp-34]

' *** Reference to "__vbaFreeObj"
'006d10e8    ff1508134000            call dword ptr [00401308]
    '#Cleanup(var_43)
'006d10ee    8b16                    mov edx, dword ptr [esi]
'006d10f0    56                      push esi
'006d10f1    ff9200030000            call dword ptr [edx+00000300]
'006d10f7    50                      push eax
'006d10f8    8d45cc                  lea eax, dword ptr [ebp-34]
'006d10fb    50                      push eax

' *** Reference to "__vbaObjSet"
'006d10fc    ff15b4104000            call dword ptr [004010b4]
    Set var_43 = Nothing
'006d1102    8d55e4                  lea edx, dword ptr [ebp-1c]
'006d1105    8bf0                    mov esi, eax
'006d1107    8b0e                    mov ecx, dword ptr [esi]
'006d1109    52                      push edx
'006d110a    56                      push esi
'006d110b    ff91a8000000            call dword ptr [ecx+000000a8]
'006d1111    dbe2                    fnclex
'006d1113    85c0                    test ax, ax
'006d1115    7d0e                    jge 6d1125
'006d1117    68a8000000              push 000000a8
'006d111c    681cb94300              push 0043b91c
'006d1121    56                      push esi
'006d1122    50                      push eax
'006d1123    ffd3                    call ebx
'006d1125    8b55e4                  mov edx, dword ptr [ebp-1c]

' *** Reference to "__vbaStrMove"
'006d1128    8b35d0124000            mov esi, dword ptr [004012d0]
'006d112e    8d4de0                  lea ecx, dword ptr [ebp-20]
'006d1131    c745e400000000          mov dword ptr [ebp-1c], 00000000
'006d1138    ffd6                    call esi
'006d113a    8d45e0                  lea eax, dword ptr [ebp-20]
'006d113d    50                      push eax
'006d113e    e8ad2ae2ff              call 4f3bf0
    Call sub_4F3BF0()
'006d1143    8bd0                    mov edx, eax
'006d1145    8d4dd0                  lea ecx, dword ptr [ebp-30]
'006d1148    ffd6                    call esi
'006d114a    8b55d0                  mov edx, dword ptr [ebp-30]
'006d114d    8995fcfeffff            mov dword ptr [ebp+fffffefc], edx
'006d1153    8b154cb07200            mov edx, dword ptr [0072b04c]
'006d1159    b804000280              mov eax, 80020004
'006d115e    898554ffffff            mov dword ptr [ebp+ffffff54], eax
'006d1164    b90a000000              mov ecx, 0000000a
'006d1169    898d4cffffff            mov dword ptr [ebp+ffffff4c], ecx
'006d116f    c745d000000000          mov dword ptr [ebp-30], 00000000
'006d1176    8b1a                    mov ebx, dword ptr [edx]
'006d1178    8bf9                    mov edi, ecx
'006d117a    8d55c8                  lea edx, dword ptr [ebp-38]
'006d117d    52                      push edx
'006d117e    83ec10                  sub esp, 10
'006d1181    8bd4                    mov edx, esp
'006d1183    890a                    mov dword ptr [edx], ecx
'006d1185    8b8d50ffffff            mov ecx, dword ptr [ebp+ffffff50]
'006d118b    894a04                  mov dword ptr [edx+04], ecx
'006d118e    894208                  mov dword ptr [edx+08], eax
'006d1191    8bf0                    mov esi, eax
'006d1193    8b8558ffffff            mov eax, dword ptr [ebp+ffffff58]
'006d1199    89420c                  mov dword ptr [edx+0c], eax
'006d119c    8b9560ffffff            mov edx, dword ptr [ebp+ffffff60]
'006d11a2    8b8568ffffff            mov eax, dword ptr [ebp+ffffff68]
'006d11a8    83ec10                  sub esp, 10
'006d11ab    8bcc                    mov ecx, esp
'006d11ad    8939                    mov dword ptr [ecx], edi
'006d11af    895104                  mov dword ptr [ecx+04], edx
'006d11b2    8b9570ffffff            mov edx, dword ptr [ebp+ffffff70]
'006d11b8    897108                  mov dword ptr [ecx+08], esi

' *** Reference to "__vbaStrMove"
'006d11bb    8b35d0124000            mov esi, dword ptr [004012d0]
'006d11c1    89410c                  mov dword ptr [ecx+0c], eax
'006d11c4    83ec10                  sub esp, 10
'006d11c7    8bcc                    mov ecx, esp
'006d11c9    b803000000              mov eax, 00000003
'006d11ce    8901                    mov dword ptr [ecx], eax
'006d11d0    895104                  mov dword ptr [ecx+04], edx
'006d11d3    8b9578ffffff            mov edx, dword ptr [ebp+ffffff78]
'006d11d9    c78574ffffff02000000    mov dword ptr [ebp+ffffff74], 00000002
'006d11e3    8b8574ffffff            mov eax, dword ptr [ebp+ffffff74]
'006d11e9    894108                  mov dword ptr [ecx+08], eax
'006d11ec    89510c                  mov dword ptr [ecx+0c], edx
'006d11ef    8b95fcfeffff            mov edx, dword ptr [ebp+fffffefc]
'006d11f5    68583c4500              push 00453c58
'006d11fa    8d4ddc                  lea ecx, dword ptr [ebp-24]
'006d11fd    ffd6                    call esi

' *** Reference to "__vbaStrCat"
'006d11ff    8b3d70104000            mov edi, dword ptr [00401070]
'006d1205    50                      push eax
'006d1206    ffd7                    call edi
    var_3 = ("select * from race where race='") & (0)
'006d1208    8bd0                    mov edx, eax
'006d120a    8d4dd8                  lea ecx, dword ptr [ebp-28]
'006d120d    ffd6                    call esi
'006d120f    50                      push eax
'006d1210    6854a44300              push 0043a454
'006d1215    ffd7                    call edi
    var_13 = (var_3) & ("'")
'006d1217    8bd0                    mov edx, eax
'006d1219    8d4dd4                  lea ecx, dword ptr [ebp-2c]
'006d121c    ffd6                    call esi
'006d121e    50                      push eax
'006d121f    a14cb07200              mov ax, word ptr [0072b04c]
'006d1224    50                      push eax
'006d1225    ff93bc000000            call dword ptr [ebx+000000bc]
'006d122b    dbe2                    fnclex
'006d122d    85c0                    test ax, ax
'006d122f    7d18                    jge 6d1249
'006d1231    8b0d4cb07200            mov ecx, dword ptr [0072b04c]
'006d1237    68bc000000              push 000000bc
'006d123c    68301f4300              push 00431f30
'006d1241    51                      push ecx
'006d1242    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006d1243    ff1580104000            call dword ptr [00401080]
'006d1249    8b45c8                  mov eax, dword ptr [ebp-38]
'006d124c    50                      push eax
'006d124d    8d55e8                  lea edx, dword ptr [ebp-18]
'006d1250    52                      push edx
'006d1251    c745c800000000          mov dword ptr [ebp-38], 00000000

' *** Reference to "__vbaObjSet"
'006d1258    ff15b4104000            call dword ptr [004010b4]
    Set var_41 = Nothing
'006d125e    8d45d0                  lea eax, dword ptr [ebp-30]
'006d1261    50                      push eax
'006d1262    8d4dd4                  lea ecx, dword ptr [ebp-2c]
'006d1265    51                      push ecx
'006d1266    8d55d8                  lea edx, dword ptr [ebp-28]
'006d1269    52                      push edx
'006d126a    8d45dc                  lea eax, dword ptr [ebp-24]
'006d126d    50                      push eax
'006d126e    8d4de0                  lea ecx, dword ptr [ebp-20]
'006d1271    51                      push ecx
'006d1272    6a05                    push 05

' *** Reference to "__vbaFreeStrList"
'006d1274    ff155c124000            call dword ptr [0040125c]
    '#Cleanup( 0, -288, -4516, -4520, var_4)
'006d127a    83c418                  add esp, 18
'006d127d    8d4dcc                  lea ecx, dword ptr [ebp-34]

' *** Reference to "__vbaFreeObj"
'006d1280    ff1508134000            call dword ptr [00401308]
    '#Cleanup(var_43)
'006d1286    8b45e8                  mov eax, dword ptr [ebp-18]
'006d1289    8b10                    mov edx, dword ptr [eax]
'006d128b    8d4dcc                  lea ecx, dword ptr [ebp-34]
'006d128e    51                      push ecx
'006d128f    50                      push eax
'006d1290    ff92b4000000            call dword ptr [edx+000000b4]
'006d1296    dbe2                    fnclex
'006d1298    85c0                    test ax, ax
'006d129a    7d15                    jge 6d12b1
'006d129c    8b55e8                  mov edx, dword ptr [ebp-18]
'006d129f    68b4000000              push 000000b4
'006d12a4    6830314300              push 00433130
'006d12a9    52                      push edx
'006d12aa    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006d12ab    ff1580104000            call dword ptr [00401080]
'006d12b1    8b45cc                  mov eax, dword ptr [ebp-34]
'006d12b4    8b38                    mov edi, dword ptr [eax]
'006d12b6    8d5dc8                  lea ebx, dword ptr [ebp-38]
'006d12b9    53                      push ebx
'006d12ba    83ec10                  sub esp, 10
'006d12bd    8bdc                    mov ebx, esp
'006d12bf    ba08000000              mov edx, 00000008
'006d12c4    8913                    mov dword ptr [ebx], edx
'006d12c6    8b9570ffffff            mov edx, dword ptr [ebp+ffffff70]
'006d12cc    895304                  mov dword ptr [ebx+04], edx
'006d12cf    b9cca64300              mov ecx, 0043a6cc
'006d12d4    894b08                  mov dword ptr [ebx+08], ecx
'006d12d7    8b8d78ffffff            mov ecx, dword ptr [ebp+ffffff78]
'006d12dd    50                      push eax
'006d12de    8bf0                    mov esi, eax
'006d12e0    894b0c                  mov dword ptr [ebx+0c], ecx
'006d12e3    ff5730                  call dword ptr [edi+30]
'006d12e6    dbe2                    fnclex
'006d12e8    85c0                    test ax, ax
'006d12ea    7d0f                    jge 6d12fb
'006d12ec    6a30                    push 30
'006d12ee    68d8304300              push 004330d8
'006d12f3    56                      push esi
'006d12f4    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006d12f5    ff1580104000            call dword ptr [00401080]
'006d12fb    8b45c8                  mov eax, dword ptr [ebp-38]
'006d12fe    8d55ac                  lea edx, dword ptr [ebp-54]
'006d1301    52                      push edx
'006d1302    c745c800000000          mov dword ptr [ebp-38], 00000000
'006d1309    8945b4                  mov dword ptr [ebp-4c], eax
'006d130c    c745ac09000000          mov dword ptr [ebp-54], 00000009

' *** Reference to "rtcIsNull"
'006d1313    ff1540114000            call dword ptr [00401140]
'006d1319    898538ffffff            mov dword ptr [ebp+ffffff38], eax
'006d131f    8b4508                  mov eax, dword ptr [ebp+08]
'006d1322    8b08                    mov ecx, dword ptr [eax]
'006d1324    50                      push eax
'006d1325    ff9104030000            call dword ptr [ecx+00000304]
'006d132b    50                      push eax
'006d132c    8d55bc                  lea edx, dword ptr [ebp-44]
'006d132f    52                      push edx

' *** Reference to "__vbaObjSet"
'006d1330    ff15b4104000            call dword ptr [004010b4]
    Set var_58 = Me
'006d1336    8d55c4                  lea edx, dword ptr [ebp-3c]
'006d1339    8bf0                    mov esi, eax
'006d133b    8b45e8                  mov eax, dword ptr [ebp-18]
'006d133e    8b08                    mov ecx, dword ptr [eax]
'006d1340    52                      push edx
'006d1341    50                      push eax
'006d1342    ff91b4000000            call dword ptr [ecx+000000b4]
'006d1348    dbe2                    fnclex
'006d134a    85c0                    test ax, ax
'006d134c    7d15                    jge 6d1363
'006d134e    8b4de8                  mov ecx, dword ptr [ebp-18]
'006d1351    68b4000000              push 000000b4
'006d1356    6830314300              push 00433130
'006d135b    51                      push ecx
'006d135c    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006d135d    ff1580104000            call dword ptr [00401080]
'006d1363    8b45c4                  mov eax, dword ptr [ebp-3c]
'006d1366    8b38                    mov edi, dword ptr [eax]
'006d1368    8d5dc0                  lea ebx, dword ptr [ebp-40]
'006d136b    53                      push ebx
'006d136c    83ec10                  sub esp, 10
'006d136f    8bdc                    mov ebx, esp
'006d1371    ba08000000              mov edx, 00000008
'006d1376    8913                    mov dword ptr [ebx], edx
'006d1378    8b9560ffffff            mov edx, dword ptr [ebp+ffffff60]
'006d137e    895304                  mov dword ptr [ebx+04], edx
'006d1381    b9cca64300              mov ecx, 0043a6cc
'006d1386    894b08                  mov dword ptr [ebx+08], ecx
'006d1389    8b8d68ffffff            mov ecx, dword ptr [ebp+ffffff68]
'006d138f    50                      push eax
'006d1390    898520ffffff            mov dword ptr [ebp+ffffff20], eax
'006d1396    894b0c                  mov dword ptr [ebx+0c], ecx
'006d1399    ff5730                  call dword ptr [edi+30]
'006d139c    dbe2                    fnclex
'006d139e    85c0                    test ax, ax
'006d13a0    7d19                    jge 6d13bb
'006d13a2    8b9520ffffff            mov edx, dword ptr [ebp+ffffff20]

' *** Reference to "__vbaHresultCheckObj"
'006d13a8    8b3d80104000            mov edi, dword ptr [00401080]
'006d13ae    6a30                    push 30
'006d13b0    68d8304300              push 004330d8
'006d13b5    52                      push edx
'006d13b6    50                      push eax
'006d13b7    ffd7                    call edi
'006d13b9    eb06                    jmp 6d13c1

' *** Reference to "__vbaHresultCheckObj"
'006d13bb    8b3d80104000            mov edi, dword ptr [00401080]
'006d13c1    8b45c0                  mov eax, dword ptr [ebp-40]
'006d13c4    8d953cffffff            lea edx, dword ptr [ebp+ffffff3c]
'006d13ca    8d4d9c                  lea ecx, dword ptr [ebp-64]
'006d13cd    c745c000000000          mov dword ptr [ebp-40], 00000000
'006d13d4    894594                  mov dword ptr [ebp-6c], eax
'006d13d7    c7458c09000000          mov dword ptr [ebp-74], 00000009
'006d13de    c78544ffffffcc134300    mov dword ptr [ebp+ffffff44], 004313cc
'006d13e8    c7853cffffff08000000    mov dword ptr [ebp+ffffff3c], 00000008

' *** Reference to "__vbaVarDup"
'006d13f2    ff1598124000            call dword ptr [00401298]
    var_51 = (vbNullChar)
'006d13f8    668b8538ffffff          mov ax, word ptr [ebp+ffffff38]
'006d13ff    8d4d8c                  lea ecx, dword ptr [ebp-74]
'006d1402    51                      push ecx
'006d1403    8d559c                  lea edx, dword ptr [ebp-64]
'006d1406    66898554ffffff          mov word ptr [ebp+ffffff54], ax
'006d140d    52                      push edx
'006d140e    8d854cffffff            lea eax, dword ptr [ebp+ffffff4c]
'006d1414    50                      push eax
'006d1415    8d8d7cffffff            lea ecx, dword ptr [ebp+ffffff7c]
'006d141b    51                      push ecx
'006d141c    c7854cffffff0b000000    mov dword ptr [ebp+ffffff4c], 0000000b

' *** Reference to "rtcImmediateIf"
'006d1426    ff154c124000            call dword ptr [0040124c]
'006d142c    8b1e                    mov ebx, dword ptr [esi]
'006d142e    8d957cffffff            lea edx, dword ptr [ebp+ffffff7c]
'006d1434    52                      push edx
'006d1435    8d45e4                  lea eax, dword ptr [ebp-1c]
'006d1438    50                      push eax

' *** Reference to "__vbaStrVarVal"
'006d1439    ff15fc114000            call dword ptr [004011fc]
'006d143f    50                      push eax
'006d1440    56                      push esi
'006d1441    ff93a4000000            call dword ptr [ebx+000000a4]
'006d1447    dbe2                    fnclex
'006d1449    85c0                    test ax, ax
'006d144b    7d0e                    jge 6d145b
'006d144d    68a4000000              push 000000a4
'006d1452    68d00d4300              push 00430dd0
'006d1457    56                      push esi
'006d1458    50                      push eax
'006d1459    ffd7                    call edi
'006d145b    8d4de4                  lea ecx, dword ptr [ebp-1c]

' *** Reference to "__vbaFreeStr"
'006d145e    ff150c134000            call dword ptr [0040130c]
    '#Cleanup(var_40)
'006d1464    8d4dbc                  lea ecx, dword ptr [ebp-44]
'006d1467    51                      push ecx
'006d1468    8d55c4                  lea edx, dword ptr [ebp-3c]
'006d146b    52                      push edx
'006d146c    8d45cc                  lea eax, dword ptr [ebp-34]
'006d146f    50                      push eax
'006d1470    6a03                    push 03

' *** Reference to "__vbaFreeObjList"
'006d1472    ff154c104000            call dword ptr [0040104c]
    '#Cleanup( var_43, var_9, var_58)
'006d1478    8d8d7cffffff            lea ecx, dword ptr [ebp+ffffff7c]
'006d147e    51                      push ecx
'006d147f    8d558c                  lea edx, dword ptr [ebp-74]
'006d1482    52                      push edx
'006d1483    8d459c                  lea eax, dword ptr [ebp-64]
'006d1486    50                      push eax
'006d1487    8d8d4cffffff            lea ecx, dword ptr [ebp+ffffff4c]
'006d148d    51                      push ecx
'006d148e    8d55ac                  lea edx, dword ptr [ebp-54]
'006d1491    52                      push edx
'006d1492    6a05                    push 05

' *** Reference to "__vbaFreeVarList"
'006d1494    ff1540104000            call dword ptr [00401040]
    '#Cleanup( var_50, var_89, var_51, var_52, var_59)
'006d149a    8b45e8                  mov eax, dword ptr [ebp-18]
'006d149d    8b08                    mov ecx, dword ptr [eax]
'006d149f    83c428                  add esp, 28
'006d14a2    50                      push eax
'006d14a3    ff91c4000000            call dword ptr [ecx+000000c4]
'006d14a9    dbe2                    fnclex
'006d14ab    85c0                    test ax, ax
'006d14ad    7d11                    jge 6d14c0
'006d14af    8b55e8                  mov edx, dword ptr [ebp-18]
'006d14b2    68c4000000              push 000000c4
'006d14b7    6830314300              push 00433130
'006d14bc    52                      push edx
'006d14bd    50                      push eax
'006d14be    ffd7                    call edi
    
End If
'006d14c0    c745fc00000000          mov dword ptr [ebp-04], 00000000
'006d14c7    9b                      fwait
'006d14c8    6834156d00              push 006d1534
'006d14cd    eb5b                    jmp 6d152a
'006d14cf    8d45d0                  lea eax, dword ptr [ebp-30]
'006d14d2    50                      push eax
'006d14d3    8d4dd4                  lea ecx, dword ptr [ebp-2c]
'006d14d6    51                      push ecx
'006d14d7    8d55d8                  lea edx, dword ptr [ebp-28]
'006d14da    52                      push edx
'006d14db    8d45dc                  lea eax, dword ptr [ebp-24]
'006d14de    50                      push eax
'006d14df    8d4de0                  lea ecx, dword ptr [ebp-20]
'006d14e2    51                      push ecx
'006d14e3    8d55e4                  lea edx, dword ptr [ebp-1c]
'006d14e6    52                      push edx
'006d14e7    6a06                    push 06

' *** Reference to "__vbaFreeStrList"
'006d14e9    ff155c124000            call dword ptr [0040125c]
'#Cleanup( , 0, -288, -4516, -4520, var_4)
'006d14ef    8d45bc                  lea eax, dword ptr [ebp-44]
'006d14f2    50                      push eax
'006d14f3    8d4dc0                  lea ecx, dword ptr [ebp-40]
'006d14f6    51                      push ecx
'006d14f7    8d55c4                  lea edx, dword ptr [ebp-3c]
'006d14fa    52                      push edx
'006d14fb    8d45c8                  lea eax, dword ptr [ebp-38]
'006d14fe    50                      push eax
'006d14ff    8d4dcc                  lea ecx, dword ptr [ebp-34]
'006d1502    51                      push ecx
'006d1503    6a05                    push 05

' *** Reference to "__vbaFreeObjList"
'006d1505    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_43, var_46, var_9, var_5, var_58)
'006d150b    8d957cffffff            lea edx, dword ptr [ebp+ffffff7c]
'006d1511    52                      push edx
'006d1512    8d458c                  lea eax, dword ptr [ebp-74]
'006d1515    50                      push eax
'006d1516    8d4d9c                  lea ecx, dword ptr [ebp-64]
'006d1519    51                      push ecx
'006d151a    8d55ac                  lea edx, dword ptr [ebp-54]
'006d151d    52                      push edx
'006d151e    6a04                    push 04

' *** Reference to "__vbaFreeVarList"
'006d1520    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_50, var_51, var_52, var_59)
'006d1526    83c448                  add esp, 48
'006d1529    c3                      ret
'006d152a    8d4de8                  lea ecx, dword ptr [ebp-18]

' *** Reference to "__vbaFreeObj"
'006d152d    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_41)
'006d1533    c3                      ret
'006d1534    8b4508                  mov eax, dword ptr [ebp+08]
'006d1537    8b08                    mov ecx, dword ptr [eax]
'006d1539    50                      push eax
'006d153a    ff5108                  call dword ptr [ecx+08]
'006d153d    8b45fc                  mov eax, dword ptr [ebp-04]
'006d1540    8b4dec                  mov ecx, dword ptr [ebp-14]
'006d1543    5f                      pop edi
'006d1544    5e                      pop esi
    'Reference to '__except_list'
'006d1545    64890d00000000          mov dword ptr fs:[00000000], ecx
'006d154c    5b                      pop ebx
'006d154d    8be5                    mov esp, ebp
'006d154f    5d                      pop ebp
'006d1550    c20400                  ret 0004


End Sub


Private Sub Form_Resize()
'006d1560    55                      push ebp
'006d1561    8bec                    mov ebp, esp
'006d1563    83ec0c                  sub esp, 0c

' *** Reference to "__vbaExceptHandler"
'006d1566    6866724000              push 00407266
'006d156b    64a100000000            mov ax, word ptr fs:[00000000]
'006d1571    50                      push eax
    'Reference to '__except_list'
'006d1572    64892500000000          mov dword ptr fs:[00000000], esp
'006d1579    83ec70                  sub esp, 70
'006d157c    53                      push ebx
'006d157d    56                      push esi
'006d157e    57                      push edi
'006d157f    8965f4                  mov dword ptr [ebp-0c], esp
'006d1582    c745f800694000          mov dword ptr [ebp-08], 00406900
'006d1589    8b7508                  mov esi, dword ptr [ebp+08]
'006d158c    8bc6                    mov eax, esi
'006d158e    83e001                  and eax, 01
'006d1591    8945fc                  mov dword ptr [ebp-04], eax
'006d1594    83e6fe                  and esi, fffffffe
'006d1597    8b0e                    mov ecx, dword ptr [esi]
'006d1599    56                      push esi
'006d159a    897508                  mov dword ptr [ebp+08], esi
'006d159d    ff5104                  call dword ptr [ecx+04]
'006d15a0    8b16                    mov edx, dword ptr [esi]
'006d15a2    33ff                    xor edi, edi
'006d15a4    56                      push esi
'006d15a5    897de8                  mov dword ptr [ebp-18], edi
'006d15a8    897dd8                  mov dword ptr [ebp-28], edi
'006d15ab    897dc8                  mov dword ptr [ebp-38], edi
'006d15ae    897da4                  mov dword ptr [ebp-5c], edi
'006d15b1    ff9204030000            call dword ptr [edx+00000304]
'006d15b7    50                      push eax
'006d15b8    8d45e8                  lea eax, dword ptr [ebp-18]
'006d15bb    50                      push eax

' *** Reference to "__vbaObjSet"
'006d15bc    ff15b4104000            call dword ptr [004010b4]
Set var_41 = Me
'006d15c2    8bd8                    mov ebx, eax
'006d15c4    393dd4b07200            cmp dword ptr [0072b0d4], edi
'006d15ca    7510                    jne 6d15dc
'006d15cc    68d4b07200              push 0072b0d4
'006d15d1    68207e4000              push 00407e20

' *** Reference to "__vbaNew2"
'006d15d6    ff152c124000            call dword ptr [0040122c]
Dim var_32 As New frmDescriptionRaces
'006d15dc    8b3dd4b07200            mov edi, dword ptr [0072b0d4]
'006d15e2    8b0f                    mov ecx, dword ptr [edi]
'006d15e4    8d55a4                  lea edx, dword ptr [ebp-5c]
'006d15e7    52                      push edx
'006d15e8    57                      push edi
'006d15e9    ff9188000000            call dword ptr [ecx+00000088]
'006d15ef    dbe2                    fnclex
'006d15f1    85c0                    test ax, ax
'006d15f3    7d12                    jge 6d1607
'006d15f5    6888000000              push 00000088
'006d15fa    6854134300              push 00431354
'006d15ff    57                      push edi
'006d1600    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006d1601    ff1580104000            call dword ptr [00401080]
'006d1607    d945a4                  fld dword ptr [ebp-5c]
'006d160a    8d4dd8                  lea ecx, dword ptr [ebp-28]
'006d160d    d825ec684000            fsub dword ptr [004068ec]
'006d1613    c745c804000000          mov dword ptr [ebp-38], 00000004
'006d161a    c745e001000000          mov dword ptr [ebp-20], 00000001
'006d1621    c745d802000000          mov dword ptr [ebp-28], 00000002
'006d1628    d95dd0                  fstp dword ptr [ebp-30]
'var_4 = (00)
'006d162b    dfe0                    fnstsw ax
'006d162d    a80d                    test al, 0d
'006d162f    0f8580010000            jne 6d17b5
'006d1635    8b3b                    mov edi, dword ptr [ebx]
'006d1637    8d45c8                  lea eax, dword ptr [ebp-38]
'006d163a    50                      push eax
'006d163b    51                      push ecx
'006d163c    e8af10e2ff              call 4f26f0
Call sub_4F26F0()
'006d1641    0fbfd0                  movsx edx, ax
'006d1644    895588                  mov dword ptr [ebp-78], edx
'006d1647    db4588                  fild dword ptr [ebp-78]
'006d164a    d95d84                  fstp dword ptr [ebp-7c]
'var_119 = (00)
'006d164d    8b4584                  mov eax, dword ptr [ebp-7c]
'006d1650    50                      push eax
'006d1651    53                      push ebx
'006d1652    ff9784000000            call dword ptr [edi+00000084]
'006d1658    dbe2                    fnclex
'006d165a    85c0                    test ax, ax
'006d165c    7d12                    jge 6d1670
'006d165e    6884000000              push 00000084
'006d1663    68d00d4300              push 00430dd0
'006d1668    53                      push ebx
'006d1669    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006d166a    ff1580104000            call dword ptr [00401080]

' *** Reference to "__vbaFreeObj"
'006d1670    8b1d08134000            mov ebx, dword ptr [00401308]
'006d1676    8d4de8                  lea ecx, dword ptr [ebp-18]
'006d1679    ffd3                    call ebx
'#Cleanup(var_41)
'006d167b    8d4dc8                  lea ecx, dword ptr [ebp-38]
'006d167e    51                      push ecx
'006d167f    8d55d8                  lea edx, dword ptr [ebp-28]
'006d1682    52                      push edx
'006d1683    6a02                    push 02

' *** Reference to "__vbaFreeVarList"
'006d1685    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_45, var_46)
'006d168b    8b06                    mov eax, dword ptr [esi]
'006d168d    83c40c                  add esp, 0c
'006d1690    56                      push esi
'006d1691    ff9004030000            call dword ptr [eax+00000304]
'006d1697    50                      push eax
'006d1698    8d4de8                  lea ecx, dword ptr [ebp-18]
'006d169b    51                      push ecx

' *** Reference to "__vbaObjSet"
'006d169c    ff15b4104000            call dword ptr [004010b4]
Set var_41 = Nothing
'006d16a2    8bf0                    mov esi, eax
'006d16a4    a1d4b07200              mov ax, word ptr [0072b0d4]
'006d16a9    85c0                    test ax, ax
'006d16ab    7510                    jne 6d16bd
'006d16ad    68d4b07200              push 0072b0d4
'006d16b2    68207e4000              push 00407e20

' *** Reference to "__vbaNew2"
'006d16b7    ff152c124000            call dword ptr [0040122c]
Set var_32 = New frmDescriptionRaces
'006d16bd    8b3dd4b07200            mov edi, dword ptr [0072b0d4]
'006d16c3    8b17                    mov edx, dword ptr [edi]
'006d16c5    8d45a4                  lea eax, dword ptr [ebp-5c]
'006d16c8    50                      push eax
'006d16c9    57                      push edi
'006d16ca    ff9280000000            call dword ptr [edx+00000080]
'006d16d0    dbe2                    fnclex
'006d16d2    85c0                    test ax, ax
'006d16d4    7d12                    jge 6d16e8
'006d16d6    6880000000              push 00000080
'006d16db    6854134300              push 00431354
'006d16e0    57                      push edi
'006d16e1    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006d16e2    ff1580104000            call dword ptr [00401080]
'006d16e8    d945a4                  fld dword ptr [ebp-5c]
'006d16eb    8d4dc8                  lea ecx, dword ptr [ebp-38]
'006d16ee    d825e8684000            fsub dword ptr [004068e8]
'006d16f4    51                      push ecx
'006d16f5    8d55d8                  lea edx, dword ptr [ebp-28]
'006d16f8    c745c804000000          mov dword ptr [ebp-38], 00000004
'006d16ff    d95dd0                  fstp dword ptr [ebp-30]
'var_4 = (00)
'006d1702    dfe0                    fnstsw ax
'006d1704    a80d                    test al, 0d
'006d1706    0f85a9000000            jne 6d17b5
'006d170c    c745e001000000          mov dword ptr [ebp-20], 00000001
'006d1713    c745d802000000          mov dword ptr [ebp-28], 00000002
'006d171a    8b3e                    mov edi, dword ptr [esi]
'006d171c    52                      push edx
'006d171d    e8ce0fe2ff              call 4f26f0
Call sub_4F26F0()
'006d1722    0fbfc0                  movsx eax, ax
'006d1725    894580                  mov dword ptr [ebp-80], eax
'006d1728    db4580                  fild dword ptr [ebp-80]
'006d172b    d99d7cffffff            fstp dword ptr [ebp+ffffff7c]
'var_59 = (00)
'006d1731    8b8d7cffffff            mov ecx, dword ptr [ebp+ffffff7c]
'006d1737    51                      push ecx
'006d1738    56                      push esi
'006d1739    ff577c                  call dword ptr [edi+7c]
'006d173c    dbe2                    fnclex
'006d173e    85c0                    test ax, ax
'006d1740    7d0f                    jge 6d1751
'006d1742    6a7c                    push 7c
'006d1744    68d00d4300              push 00430dd0
'006d1749    56                      push esi
'006d174a    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006d174b    ff1580104000            call dword ptr [00401080]
'006d1751    8d4de8                  lea ecx, dword ptr [ebp-18]
'006d1754    ffd3                    call ebx
'#Cleanup(var_41)
'006d1756    8d55c8                  lea edx, dword ptr [ebp-38]
'006d1759    52                      push edx
'006d175a    8d45d8                  lea eax, dword ptr [ebp-28]
'006d175d    50                      push eax
'006d175e    6a02                    push 02

' *** Reference to "__vbaFreeVarList"
'006d1760    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_45, var_46)
'006d1766    83c40c                  add esp, 0c
'006d1769    c745fc00000000          mov dword ptr [ebp-04], 00000000
'006d1770    9b                      fwait
'006d1771    6896176d00              push 006d1796
'006d1776    eb1d                    jmp 6d1795
'006d1778    8d4de8                  lea ecx, dword ptr [ebp-18]

' *** Reference to "__vbaFreeObj"
'006d177b    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_41)
'006d1781    8d4dc8                  lea ecx, dword ptr [ebp-38]
'006d1784    51                      push ecx
'006d1785    8d55d8                  lea edx, dword ptr [ebp-28]
'006d1788    52                      push edx
'006d1789    6a02                    push 02

' *** Reference to "__vbaFreeVarList"
'006d178b    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_45, var_46)
'006d1791    83c40c                  add esp, 0c
'006d1794    c3                      ret
'006d1795    c3                      ret
'006d1796    8b4508                  mov eax, dword ptr [ebp+08]
'006d1799    8b08                    mov ecx, dword ptr [eax]
'006d179b    50                      push eax
'006d179c    ff5108                  call dword ptr [ecx+08]
'006d179f    8b45fc                  mov eax, dword ptr [ebp-04]
'006d17a2    8b4dec                  mov ecx, dword ptr [ebp-14]
'006d17a5    5f                      pop edi
'006d17a6    5e                      pop esi
    'Reference to '__except_list'
'006d17a7    64890d00000000          mov dword ptr fs:[00000000], ecx
'006d17ae    5b                      pop ebx
'006d17af    8be5                    mov esp, ebp
'006d17b1    5d                      pop ebp
'006d17b2    c20400                  ret 0004


End Sub


