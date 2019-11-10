VERSION 5.00

Begin VB.Form frmRenommer
    Caption = "Renommer"
    ScaleMode = 1
    AutoRedraw = 0              'False
    FontTransparent = -1              'True
    LinkTopic = "Form1"
    ControlBox = 0              'False
    ClipControls = 0              'False
    ClientLeft   = 60
    ClientTop    = 450
    ClientWidth  = 5445
    ClientHeight = 945
    StartupPosition = 1
    Begin VB.CommandButton btnAnnuler
        Caption = "&Annuler"
        Left   = 4230
        Top    = 540
        Width  = 1140
        Height = 330
        TabIndex = 2
        Cancel = -1              'True
    End
    Begin VB.CommandButton btnOk
        Caption = "&Ok"
        Left   = 3015
        Top    = 540
        Width  = 1140
        Height = 330
        Enabled = 0              'False
        TabIndex = 1
    End
    Begin VB.TextBox fldstrNouveauNom
        Left   = 1395
        Top    = 135
        Width  = 3930
        Height = 285
        TabIndex = 0
        MaxLength = 50
    End
    Begin VB.Label Label1
        Caption = "Nouveau nom:"
        Left   = 135
        Top    = 180
        Width  = 1185
        Height = 240
        TabIndex = 3
    End
End
'Event for btnOk
Private Sub btnOk_Click()
'006cd990    55                      push ebp
'006cd991    8bec                    mov ebp, esp
'006cd993    83ec0c                  sub esp, 0c

' *** Reference to "__vbaExceptHandler"
'006cd996    6866724000              push 00407266
'006cd99b    64a100000000            mov ax, word ptr fs:[00000000]
'006cd9a1    50                      push eax
    'Reference to '__except_list'
'006cd9a2    64892500000000          mov dword ptr fs:[00000000], esp
'006cd9a9    83ec5c                  sub esp, 5c
'006cd9ac    53                      push ebx
'006cd9ad    56                      push esi
'006cd9ae    57                      push edi
'006cd9af    8965f4                  mov dword ptr [ebp-0c], esp
'006cd9b2    c745f818684000          mov dword ptr [ebp-08], 00406818
'006cd9b9    8b7508                  mov esi, dword ptr [ebp+08]
'006cd9bc    8bc6                    mov eax, esi
'006cd9be    83e001                  and eax, 01
'006cd9c1    8945fc                  mov dword ptr [ebp-04], eax
'006cd9c4    83e6fe                  and esi, fffffffe
'006cd9c7    8b0e                    mov ecx, dword ptr [esi]
'006cd9c9    56                      push esi
'006cd9ca    897508                  mov dword ptr [ebp+08], esi
'006cd9cd    ff5104                  call dword ptr [ecx+04]
'006cd9d0    8b16                    mov edx, dword ptr [esi]
'006cd9d2    33db                    xor ebx, ebx
'006cd9d4    56                      push esi
'006cd9d5    895de8                  mov dword ptr [ebp-18], ebx
'006cd9d8    895de4                  mov dword ptr [ebp-1c], ebx
'006cd9db    895dd4                  mov dword ptr [ebp-2c], ebx
'006cd9de    895dc4                  mov dword ptr [ebp-3c], ebx
'006cd9e1    895db4                  mov dword ptr [ebp-4c], ebx
'006cd9e4    895da4                  mov dword ptr [ebp-5c], ebx
'006cd9e7    ff9204030000            call dword ptr [edx+00000304]
'006cd9ed    8945dc                  mov dword ptr [ebp-24], eax
'006cd9f0    8d45d4                  lea eax, dword ptr [ebp-2c]
'006cd9f3    50                      push eax
'006cd9f4    8d4dc4                  lea ecx, dword ptr [ebp-3c]
'006cd9f7    51                      push ecx
'006cd9f8    c745d409000000          mov dword ptr [ebp-2c], 00000009

' *** Reference to "rtcTrimVar"
'006cd9ff    ff15e4104000            call dword ptr [004010e4]
'006cda05    8d55c4                  lea edx, dword ptr [ebp-3c]
'006cda08    52                      push edx
'006cda09    8d45a4                  lea eax, dword ptr [ebp-5c]
'006cda0c    50                      push eax
'006cda0d    c745accc134300          mov dword ptr [ebp-54], 004313cc
'006cda14    c745a408800000          mov dword ptr [ebp-5c], 00008008

' *** Reference to "__vbaVarTstEq"
'006cda1b    ff153c114000            call dword ptr [0040113c]
'006cda21    8d4dc4                  lea ecx, dword ptr [ebp-3c]
'006cda24    51                      push ecx
'006cda25    8d55d4                  lea edx, dword ptr [ebp-2c]
'006cda28    52                      push edx
'006cda29    6a02                    push 02
'006cda2b    668bf8                  mov di, ax

' *** Reference to "__vbaFreeVarList"
'006cda2e    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_44, var_9)
'006cda34    8b06                    mov eax, dword ptr [esi]
'006cda36    83c40c                  add esp, 0c
'006cda39    663bfb                  cmp di, bx
'006cda3c    56                      push esi
'006cda3d    7436                    je 6cda75

If (((Trim(Me)) = (vbNullChar))) Then
'006cda3f    ff9004030000            call dword ptr [eax+00000304]
'006cda45    50                      push eax
'006cda46    8d4de4                  lea ecx, dword ptr [ebp-1c]
'006cda49    51                      push ecx

' *** Reference to "__vbaObjSet"
'006cda4a    ff15b4104000            call dword ptr [004010b4]
    Set var_40 = Nothing
'006cda50    8bf0                    mov esi, eax
'006cda52    8b16                    mov edx, dword ptr [esi]
'006cda54    56                      push esi
'006cda55    ff9204020000            call dword ptr [edx+00000204]
'006cda5b    dbe2                    fnclex
'006cda5d    3bc3                    cmp eax, ebx
'006cda5f    0f8db1000000            jge 6cdb16
    
    If (    var_40 < 0) Then
'006cda65    6804020000              push 00000204
'006cda6a    68d00d4300              push 00430dd0
'006cda6f    56                      push esi
'006cda70    e99a000000              jmp 6cdb0f
    
End If
'006cda75    ff9004030000            call dword ptr [eax+00000304]
'006cda7b    50                      push eax
'006cda7c    8d4de4                  lea ecx, dword ptr [ebp-1c]
'006cda7f    51                      push ecx

' *** Reference to "__vbaObjSet"
'006cda80    ff15b4104000            call dword ptr [004010b4]
Set var_40 = Nothing
'006cda86    8bf8                    mov edi, eax
'006cda88    8b17                    mov edx, dword ptr [edi]
'006cda8a    8d45e8                  lea eax, dword ptr [ebp-18]
'006cda8d    50                      push eax
'006cda8e    57                      push edi
'006cda8f    ff92a0000000            call dword ptr [edx+000000a0]
'006cda95    dbe2                    fnclex
'006cda97    3bc3                    cmp eax, ebx
'006cda99    7d12                    jge 6cdaad
'006cda9b    68a0000000              push 000000a0
'006cdaa0    68d00d4300              push 00430dd0
'006cdaa5    57                      push edi
'006cdaa6    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006cdaa7    ff1580104000            call dword ptr [00401080]
'006cdaad    8b55e8                  mov edx, dword ptr [ebp-18]
'006cdab0    8d4e34                  lea ecx, dword ptr [esi+34]

' *** Reference to "__vbaStrCopy"
'006cdab3    ff1548124000            call dword ptr [00401248]
var_138 = (vbNullString)
'006cdab9    8d4de8                  lea ecx, dword ptr [ebp-18]

' *** Reference to "__vbaFreeStr"
'006cdabc    ff150c134000            call dword ptr [0040130c]
'#Cleanup(var_41)
'006cdac2    8d4de4                  lea ecx, dword ptr [ebp-1c]

' *** Reference to "__vbaFreeObj"
'006cdac5    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_40)
'006cdacb    391d24be7200            cmp dword ptr [0072be24], ebx
'006cdad1    7510                    jne 6cdae3
'006cdad3    6824be7200              push 0072be24
'006cdad8    6870004300              push 00430070

' *** Reference to "__vbaNew2"
'006cdadd    ff152c124000            call dword ptr [0040122c]
Dim var_2 As New Global
'006cdae3    8b3d24be7200            mov edi, dword ptr [0072be24]
'006cdae9    8b17                    mov edx, dword ptr [edi]
'006cdaeb    56                      push esi
'006cdaec    8d4de4                  lea ecx, dword ptr [ebp-1c]
'006cdaef    51                      push ecx
'006cdaf0    895590                  mov dword ptr [ebp-70], edx

' *** Reference to "__vbaObjSetAddref"
'006cdaf3    ff15c8104000            call dword ptr [004010c8]
Set var_40 = Me
'006cdaf9    8b5590                  mov edx, dword ptr [ebp-70]
'006cdafc    50                      push eax
'006cdafd    57                      push edi
'006cdafe    ff5210                  call dword ptr [edx+10]
Call var_2.Unload(var_40)
'006cdb01    dbe2                    fnclex
'006cdb03    3bc3                    cmp eax, ebx
'006cdb05    7d0f                    jge 6cdb16
'006cdb07    6a10                    push 10
'006cdb09    6860004300              push 00430060
'006cdb0e    57                      push edi
'006cdb0f    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006cdb10    ff1580104000            call dword ptr [00401080]

'ERROR: Two many next close:
End If
'006cdb16    8d4de4                  lea ecx, dword ptr [ebp-1c]

' *** Reference to "__vbaFreeObj"
'006cdb19    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_40)
'006cdb1f    895dfc                  mov dword ptr [ebp-04], ebx
'006cdb22    6854db6c00              push 006cdb54
'006cdb27    eb2a                    jmp 6cdb53
'006cdb29    8d4de8                  lea ecx, dword ptr [ebp-18]

' *** Reference to "__vbaFreeStr"
'006cdb2c    ff150c134000            call dword ptr [0040130c]
'#Cleanup(var_41)
'006cdb32    8d4de4                  lea ecx, dword ptr [ebp-1c]

' *** Reference to "__vbaFreeObj"
'006cdb35    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_40)
'006cdb3b    8d45b4                  lea eax, dword ptr [ebp-4c]
'006cdb3e    50                      push eax
'006cdb3f    8d4dc4                  lea ecx, dword ptr [ebp-3c]
'006cdb42    51                      push ecx
'006cdb43    8d55d4                  lea edx, dword ptr [ebp-2c]
'006cdb46    52                      push edx
'006cdb47    6a03                    push 03

' *** Reference to "__vbaFreeVarList"
'006cdb49    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_44, var_9, var_62)
'006cdb4f    83c410                  add esp, 10
'006cdb52    c3                      ret
'006cdb53    c3                      ret
'006cdb54    8b4508                  mov eax, dword ptr [ebp+08]
'006cdb57    8b08                    mov ecx, dword ptr [eax]
'006cdb59    50                      push eax
'006cdb5a    ff5108                  call dword ptr [ecx+08]
'006cdb5d    8b45fc                  mov eax, dword ptr [ebp-04]
'006cdb60    8b4dec                  mov ecx, dword ptr [ebp-14]
'006cdb63    5f                      pop edi
'006cdb64    5e                      pop esi
    'Reference to '__except_list'
'006cdb65    64890d00000000          mov dword ptr fs:[00000000], ecx
'006cdb6c    5b                      pop ebx
'006cdb6d    8be5                    mov esp, ebp
'006cdb6f    5d                      pop ebp
'006cdb70    c20400                  ret 0004


End Sub


'Event for fldstrNouveauNom
Private Sub fldstrNouveauNom_Change()
'006cdb80    55                      push ebp
'006cdb81    8bec                    mov ebp, esp
'006cdb83    83ec0c                  sub esp, 0c

' *** Reference to "__vbaExceptHandler"
'006cdb86    6866724000              push 00407266
'006cdb8b    64a100000000            mov ax, word ptr fs:[00000000]
'006cdb91    50                      push eax
    'Reference to '__except_list'
'006cdb92    64892500000000          mov dword ptr fs:[00000000], esp
'006cdb99    83ec64                  sub esp, 64
'006cdb9c    53                      push ebx
'006cdb9d    56                      push esi
'006cdb9e    57                      push edi
'006cdb9f    8965f4                  mov dword ptr [ebp-0c], esp
'006cdba2    c745f828684000          mov dword ptr [ebp-08], 00406828
'006cdba9    8b7508                  mov esi, dword ptr [ebp+08]
'006cdbac    8bc6                    mov eax, esi
'006cdbae    83e001                  and eax, 01
'006cdbb1    8945fc                  mov dword ptr [ebp-04], eax
'006cdbb4    83e6fe                  and esi, fffffffe
'006cdbb7    8b0e                    mov ecx, dword ptr [esi]
'006cdbb9    56                      push esi
'006cdbba    897508                  mov dword ptr [ebp+08], esi
'006cdbbd    ff5104                  call dword ptr [ecx+04]
'006cdbc0    8b16                    mov edx, dword ptr [esi]
'006cdbc2    33db                    xor ebx, ebx
'006cdbc4    56                      push esi
'006cdbc5    895de8                  mov dword ptr [ebp-18], ebx
'006cdbc8    895dd8                  mov dword ptr [ebp-28], ebx
'006cdbcb    895dc8                  mov dword ptr [ebp-38], ebx
'006cdbce    895db8                  mov dword ptr [ebp-48], ebx
'006cdbd1    895da8                  mov dword ptr [ebp-58], ebx
'006cdbd4    895d98                  mov dword ptr [ebp-68], ebx
'006cdbd7    ff9200030000            call dword ptr [edx+00000300]
'006cdbdd    50                      push eax
'006cdbde    8d45e8                  lea eax, dword ptr [ebp-18]
'006cdbe1    50                      push eax

' *** Reference to "__vbaObjSet"
'006cdbe2    ff15b4104000            call dword ptr [004010b4]
Set var_41 = Me
'006cdbe8    8b0e                    mov ecx, dword ptr [esi]
'006cdbea    56                      push esi
'006cdbeb    8bf8                    mov edi, eax
'006cdbed    ff9104030000            call dword ptr [ecx+00000304]
'006cdbf3    8d55d8                  lea edx, dword ptr [ebp-28]
'006cdbf6    8945e0                  mov dword ptr [ebp-20], eax
'006cdbf9    52                      push edx
'006cdbfa    8d45c8                  lea eax, dword ptr [ebp-38]
'006cdbfd    50                      push eax
'006cdbfe    c745d809000000          mov dword ptr [ebp-28], 00000009

' *** Reference to "rtcTrimVar"
'006cdc05    ff15e4104000            call dword ptr [004010e4]
'006cdc0b    8d4dc8                  lea ecx, dword ptr [ebp-38]
'006cdc0e    51                      push ecx
'006cdc0f    8d55b8                  lea edx, dword ptr [ebp-48]
'006cdc12    895da0                  mov dword ptr [ebp-60], ebx
'006cdc15    c7459802800000          mov dword ptr [ebp-68], 00008002
'006cdc1c    8b37                    mov esi, dword ptr [edi]
'006cdc1e    52                      push edx

' *** Reference to "__vbaLenVar"
'006cdc1f    ff1584104000            call dword ptr [00401084]
var_61 = Len(Trim(var_41))
'006cdc25    50                      push eax
'006cdc26    8d4598                  lea eax, dword ptr [ebp-68]
'006cdc29    50                      push eax
'006cdc2a    8d4da8                  lea ecx, dword ptr [ebp-58]
'006cdc2d    51                      push ecx

' *** Reference to "__vbaVarCmpGt"
'006cdc2e    ff1510114000            call dword ptr [00401110]
'006cdc34    50                      push eax

' *** Reference to "__vbaBoolVar"
'006cdc35    ff15e0104000            call dword ptr [004010e0]
'006cdc3b    50                      push eax
'006cdc3c    57                      push edi
'006cdc3d    ff968c000000            call dword ptr [esi+0000008c]
'006cdc43    dbe2                    fnclex
'006cdc45    3bc3                    cmp eax, ebx
'006cdc47    7d12                    jge 6cdc5b

If (CBool(#NOT SUPPORTED#) < 0) Then
'006cdc49    688c000000              push 0000008c
'006cdc4e    6838eb4300              push 0043eb38
'006cdc53    57                      push edi
'006cdc54    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006cdc55    ff1580104000            call dword ptr [00401080]
    
End If
'006cdc5b    8d4de8                  lea ecx, dword ptr [ebp-18]

' *** Reference to "__vbaFreeObj"
'006cdc5e    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_41)
'006cdc64    8d55c8                  lea edx, dword ptr [ebp-38]
'006cdc67    52                      push edx
'006cdc68    8d45d8                  lea eax, dword ptr [ebp-28]
'006cdc6b    50                      push eax
'006cdc6c    6a02                    push 02

' *** Reference to "__vbaFreeVarList"
'006cdc6e    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_45, var_46)
'006cdc74    83c40c                  add esp, 0c
'006cdc77    895dfc                  mov dword ptr [ebp-04], ebx
'006cdc7a    68a7dc6c00              push 006cdca7
'006cdc7f    eb25                    jmp 6cdca6
'006cdc81    8d4de8                  lea ecx, dword ptr [ebp-18]

' *** Reference to "__vbaFreeObj"
'006cdc84    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_41)
'006cdc8a    8d4da8                  lea ecx, dword ptr [ebp-58]
'006cdc8d    51                      push ecx
'006cdc8e    8d55b8                  lea edx, dword ptr [ebp-48]
'006cdc91    52                      push edx
'006cdc92    8d45c8                  lea eax, dword ptr [ebp-38]
'006cdc95    50                      push eax
'006cdc96    8d4dd8                  lea ecx, dword ptr [ebp-28]
'006cdc99    51                      push ecx
'006cdc9a    6a04                    push 04

' *** Reference to "__vbaFreeVarList"
'006cdc9c    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_45, var_46, var_61, var_86)
'006cdca2    83c414                  add esp, 14
'006cdca5    c3                      ret
'006cdca6    c3                      ret
'006cdca7    8b4508                  mov eax, dword ptr [ebp+08]
'006cdcaa    8b10                    mov edx, dword ptr [eax]
'006cdcac    50                      push eax
'006cdcad    ff5208                  call dword ptr [edx+08]
'006cdcb0    8b45fc                  mov eax, dword ptr [ebp-04]
'006cdcb3    8b4dec                  mov ecx, dword ptr [ebp-14]
'006cdcb6    5f                      pop edi
'006cdcb7    5e                      pop esi
    'Reference to '__except_list'
'006cdcb8    64890d00000000          mov dword ptr fs:[00000000], ecx
'006cdcbf    5b                      pop ebx
'006cdcc0    8be5                    mov esp, ebp
'006cdcc2    5d                      pop ebp
'006cdcc3    c20400                  ret 0004


End Sub


Private Sub fldstrNouveauNom_GotFocus()
'006cdcd0    55                      push ebp
'006cdcd1    8bec                    mov ebp, esp
'006cdcd3    83ec0c                  sub esp, 0c

' *** Reference to "__vbaExceptHandler"
'006cdcd6    6866724000              push 00407266
'006cdcdb    64a100000000            mov ax, word ptr fs:[00000000]
'006cdce1    50                      push eax
    'Reference to '__except_list'
'006cdce2    64892500000000          mov dword ptr fs:[00000000], esp
'006cdce9    83ec08                  sub esp, 08
'006cdcec    53                      push ebx
'006cdced    56                      push esi
'006cdcee    57                      push edi
'006cdcef    8965f4                  mov dword ptr [ebp-0c], esp
'006cdcf2    c745f838684000          mov dword ptr [ebp-08], 00406838
'006cdcf9    8b4508                  mov eax, dword ptr [ebp+08]
'006cdcfc    8bc8                    mov ecx, eax
'006cdcfe    83e101                  and ecx, 01
'006cdd01    894dfc                  mov dword ptr [ebp-04], ecx
'006cdd04    83e0fe                  and eax, fffffffe
'006cdd07    8b10                    mov edx, dword ptr [eax]
'006cdd09    50                      push eax
'006cdd0a    894508                  mov dword ptr [ebp+08], eax
'006cdd0d    ff5204                  call dword ptr [edx+04]
'006cdd10    e8fb4fe2ff              call 4f2d10
Call sub_4F2D10()
'006cdd15    c745fc00000000          mov dword ptr [ebp-04], 00000000
'006cdd1c    8b4508                  mov eax, dword ptr [ebp+08]
'006cdd1f    8b08                    mov ecx, dword ptr [eax]
'006cdd21    50                      push eax
'006cdd22    ff5108                  call dword ptr [ecx+08]
'006cdd25    8b45fc                  mov eax, dword ptr [ebp-04]
'006cdd28    8b4dec                  mov ecx, dword ptr [ebp-14]
'006cdd2b    5f                      pop edi
'006cdd2c    5e                      pop esi
    'Reference to '__except_list'
'006cdd2d    64890d00000000          mov dword ptr fs:[00000000], ecx
'006cdd34    5b                      pop ebx
'006cdd35    8be5                    mov esp, ebp
'006cdd37    5d                      pop ebp
'006cdd38    c20400                  ret 0004


End Sub


'Event for Form
Private Sub Form_Load()
'006cdd40    55                      push ebp
'006cdd41    8bec                    mov ebp, esp
'006cdd43    83ec0c                  sub esp, 0c

' *** Reference to "__vbaExceptHandler"
'006cdd46    6866724000              push 00407266
'006cdd4b    64a100000000            mov ax, word ptr fs:[00000000]
'006cdd51    50                      push eax
    'Reference to '__except_list'
'006cdd52    64892500000000          mov dword ptr fs:[00000000], esp
'006cdd59    83ec08                  sub esp, 08
'006cdd5c    53                      push ebx
'006cdd5d    56                      push esi
'006cdd5e    57                      push edi
'006cdd5f    8965f4                  mov dword ptr [ebp-0c], esp
'006cdd62    c745f840684000          mov dword ptr [ebp-08], 00406840
'006cdd69    8b7508                  mov esi, dword ptr [ebp+08]
'006cdd6c    8bc6                    mov eax, esi
'006cdd6e    83e001                  and eax, 01
'006cdd71    8945fc                  mov dword ptr [ebp-04], eax
'006cdd74    83e6fe                  and esi, fffffffe
'006cdd77    8b0e                    mov ecx, dword ptr [esi]
'006cdd79    56                      push esi
'006cdd7a    897508                  mov dword ptr [ebp+08], esi
'006cdd7d    ff5104                  call dword ptr [ecx+04]
'006cdd80    bacc134300              mov edx, 004313cc
'006cdd85    8d4e34                  lea ecx, dword ptr [esi+34]

' *** Reference to "__vbaStrCopy"
'006cdd88    ff1548124000            call dword ptr [00401248]
var_138 = (vbNullChar)
'006cdd8e    c745fc00000000          mov dword ptr [ebp-04], 00000000
'006cdd95    8b4508                  mov eax, dword ptr [ebp+08]
'006cdd98    8b10                    mov edx, dword ptr [eax]
'006cdd9a    50                      push eax
'006cdd9b    ff5208                  call dword ptr [edx+08]
'006cdd9e    8b45fc                  mov eax, dword ptr [ebp-04]
'006cdda1    8b4dec                  mov ecx, dword ptr [ebp-14]
'006cdda4    5f                      pop edi
'006cdda5    5e                      pop esi
    'Reference to '__except_list'
'006cdda6    64890d00000000          mov dword ptr fs:[00000000], ecx
'006cddad    5b                      pop ebx
'006cddae    8be5                    mov esp, ebp
'006cddb0    5d                      pop ebp
'006cddb1    c20400                  ret 0004


End Sub


'Event for btnAnnuler
Private Sub btnAnnuler_Click()
'006cd8c0    55                      push ebp
'006cd8c1    8bec                    mov ebp, esp
'006cd8c3    83ec0c                  sub esp, 0c

' *** Reference to "__vbaExceptHandler"
'006cd8c6    6866724000              push 00407266
'006cd8cb    64a100000000            mov ax, word ptr fs:[00000000]
'006cd8d1    50                      push eax
    'Reference to '__except_list'
'006cd8d2    64892500000000          mov dword ptr fs:[00000000], esp
'006cd8d9    83ec18                  sub esp, 18
'006cd8dc    53                      push ebx
'006cd8dd    56                      push esi
'006cd8de    57                      push edi
'006cd8df    8965f4                  mov dword ptr [ebp-0c], esp
'006cd8e2    c745f808684000          mov dword ptr [ebp-08], 00406808
'006cd8e9    8b7d08                  mov edi, dword ptr [ebp+08]
'006cd8ec    8bc7                    mov eax, edi
'006cd8ee    83e001                  and eax, 01
'006cd8f1    8945fc                  mov dword ptr [ebp-04], eax
'006cd8f4    83e7fe                  and edi, fffffffe
'006cd8f7    8b0f                    mov ecx, dword ptr [edi]
'006cd8f9    57                      push edi
'006cd8fa    897d08                  mov dword ptr [ebp+08], edi
'006cd8fd    ff5104                  call dword ptr [ecx+04]
'006cd900    a124be7200              mov ax, word ptr [0072be24]
'006cd905    33db                    xor ebx, ebx
'006cd907    3bc3                    cmp eax, ebx
'006cd909    895de8                  mov dword ptr [ebp-18], ebx
'006cd90c    7510                    jne 6cd91e

If (0 = 0) Then
'006cd90e    6824be7200              push 0072be24
'006cd913    6870004300              push 00430070

' *** Reference to "__vbaNew2"
'006cd918    ff152c124000            call dword ptr [0040122c]
    Dim var_2 As New Global
End If
'006cd91e    8b3524be7200            mov esi, dword ptr [0072be24]
'006cd924    8b16                    mov edx, dword ptr [esi]
'006cd926    57                      push edi
'006cd927    8d45e8                  lea eax, dword ptr [ebp-18]
'006cd92a    50                      push eax
'006cd92b    8955d4                  mov dword ptr [ebp-2c], edx

' *** Reference to "__vbaObjSetAddref"
'006cd92e    ff15c8104000            call dword ptr [004010c8]
Set var_41 = Me
'006cd934    8b4dd4                  mov ecx, dword ptr [ebp-2c]
'006cd937    50                      push eax
'006cd938    56                      push esi
'006cd939    ff5110                  call dword ptr [ecx+10]
Call var_2.Unload(var_41)
'006cd93c    dbe2                    fnclex
'006cd93e    3bc3                    cmp eax, ebx
'006cd940    7d0f                    jge 6cd951
'006cd942    6a10                    push 10
'006cd944    6860004300              push 00430060
'006cd949    56                      push esi
'006cd94a    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006cd94b    ff1580104000            call dword ptr [00401080]
'006cd951    8d4de8                  lea ecx, dword ptr [ebp-18]

' *** Reference to "__vbaFreeObj"
'006cd954    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_41)
'006cd95a    895dfc                  mov dword ptr [ebp-04], ebx
'006cd95d    686fd96c00              push 006cd96f
'006cd962    eb0a                    jmp 6cd96e
'006cd964    8d4de8                  lea ecx, dword ptr [ebp-18]

' *** Reference to "__vbaFreeObj"
'006cd967    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_41)
'006cd96d    c3                      ret
'006cd96e    c3                      ret
'006cd96f    8b4508                  mov eax, dword ptr [ebp+08]
'006cd972    8b10                    mov edx, dword ptr [eax]
'006cd974    50                      push eax
'006cd975    ff5208                  call dword ptr [edx+08]
'006cd978    8b45fc                  mov eax, dword ptr [ebp-04]
'006cd97b    8b4dec                  mov ecx, dword ptr [ebp-14]
'006cd97e    5f                      pop edi
'006cd97f    5e                      pop esi
    'Reference to '__except_list'
'006cd980    64890d00000000          mov dword ptr fs:[00000000], ecx
'006cd987    5b                      pop ebx
'006cd988    8be5                    mov esp, ebp
'006cd98a    5d                      pop ebp
'006cd98b    c20400                  ret 0004


End Sub


